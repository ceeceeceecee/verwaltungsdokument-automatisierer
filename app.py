"""Verwaltungsdokument-Automatisierer - Streamlit App"""
import streamlit as st, os, sys, json, re, plotly.express as px, pandas as pd
from datetime import datetime
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from database.db import DatabaseManager
from processor.analyzer import DokumentAnalyzer

DB_PATH = os.getenv("DB_PATH", str(Path(__file__).parent / "data" / "dokumente.db"))
db = DatabaseManager(DB_PATH)

st.set_page_config(page_title="Verwaltungsdokument-Automatisierer", page_icon="📄", layout="wide", initial_sidebar_state="expanded")
st.markdown("[data-testid='stSidebar']{background-color:#151e2e;} [data-testid='stSidebar'] *{color:#c8d6e5 !important;}", unsafe_allow_html=True)

if "analyzer" not in st.session_state:
    st.session_state.analyzer = DokumentAnalyzer(db.get_setting("ollama_url","http://localhost:11434"), db.get_setting("ollama_model","llama3.1:8b"))

with st.sidebar:
    st.markdown("📄 **Verwaltungsdokument**")
    st.caption("Automatisierte Dokumenterstellung")
    st.divider()
    s=db.get_stats()
    st.markdown(f"Vorlagen: **{s['vorlagen']}**")
    st.markdown(f"Dokumente: **{s['dokumente']}**")
    st.divider()
    ok=st.session_state.analyzer.is_available()
    st.markdown("✅ Ollama verbunden" if ok else "⚠️ Demo-Modus")

page = st.sidebar.radio("Navigation", ["📊 Dashboard","📝 Vorlagen","📄 Dokument erstellen","⚙️ Einstellungen"], label_visibility="collapsed")

if page == "📊 Dashboard":
    st.title("📊 Dashboard")
    s=db.get_stats()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("📝 Vorlagen",s["vorlagen"])
    c2.metric("📄 Dokumente",s["dokumente"])
    c3.metric("📋 Entwuerfe",s["entwurf"])
    c4.metric("✅ Fertig",s["fertig"])
    st.divider()
    docs=db.get_all_dokumente()
    if docs:
        for d in docs: d["status_emoji"]="📋" if d["status"]=="entwurf" else "✅"
        df=pd.DataFrame(docs)
        st.dataframe(df[["titel","vorlage_name","status_emoji","erstellt_am"]],use_container_width=True,hide_index=True,height=400)
        cc=db.get_category_counts()
        if cc:
            cdf=pd.DataFrame(list(cc.items()),columns=["Kategorie","Anzahl"])
            fig=px.pie(cdf,values="Anzahl",names="Kategorie",title="Vorlagen nach Kategorie")
            st.plotly_chart(fig,use_container_width=True)

elif page == "📝 Vorlagen":
    st.title("📝 Vorlagen verwalten")
    tab1,tab2=st.tabs(["Vorlagen","Neue Vorlage"])
    with tab1:
        vorlagen=db.get_all_vorlagen()
        for v in vorlagen:
            with st.expander(f"📋 {v['name']} ({v['kategorie']})"):
                st.code(v["inhalt"][:500]+"..." if len(v["inhalt"])>500 else v["inhalt"])
                vars=json.loads(v["variablen_json"]) if v.get("variablen_json") else []
                if vars: st.caption(f"Variablen: {', '.join(vars)}")
    with tab2:
        with st.form("neu_vorlage"):
            name=st.text_input("Name*")
            kat=st.selectbox("Kategorie",["Behoerde","Korrespondenz","Finanzen"])
            inhalt=st.text_area("Inhalt mit {{Variable}}",height=200,placeholder="Sehr geehrte/r {{Anrede}} {{Name}},...")
            if st.form_submit_button("💾 Speichern",type="primary"):
                if not name or not inhalt: st.error("Pflichtfelder!")
                else:
                    vars=list(set(re.findall(r'\{\{(\w+)\}\}',inhalt)))
                    db.create_vorlage(name,kat,inhalt,vars)
                    st.success(f"Vorlage gespeichert! {len(vars)} Variablen gefunden.")
                    st.rerun()

elif page == "📄 Dokument erstellen":
    st.title("📄 Dokument erstellen")
    vorlagen=db.get_all_vorlagen()
    if not vorlagen: st.info("Keine Vorlagen vorhanden."); st.stop()
    selected=st.selectbox("Vorlage waehlen", [v["name"] for v in vorlagen])
    vid=[v["id"] for v in vorlagen if v["name"]==selected][0]
    v=db.get_vorlage(vid)
    if v:
        vars=json.loads(v["variablen_json"]) if v.get("variablen_json") else []
        st.caption(f"Kategorie: {v['kategorie']} | {len(vars)} Variablen")
        st.divider()
        inputs={}
        c1,c2=st.columns(2)
        cols=[c1,c2]
        for i,var in enumerate(vars):
            with cols[i%2]:
                # Default values
                defaults={"Anrede":"Herr","Name":"Mustermann","Datum":datetime.now().strftime("%d.%m.%Y"),
                          "Behoerde":"Stadtverwaltung Musterhausen","Adresse":"Musterstr. 1, 99867 Musterhausen"}
                inputs[var]=st.text_input(var,defaults.get(var,""))
        st.divider()
        # Generate preview
        inhalt=v["inhalt"]
        for var,val in inputs.items():
            inhalt=inhalt.replace(f"{{{{{var}}}}}",val)
        st.subheader("Vorschau")
        st.text_area("Dokument",inhalt,height=300,label_visibility="collapsed")
        c1,c2=st.columns(2)
        with c1:
            if st.button("📄 Als Entwurf speichern",use_container_width=True):
                db.create_dokument(vid,f"{selected} - {inputs.get('Name','')}",inputs,inhalt)
                st.success("Dokument gespeichert!")
        with c2:
            if st.button("🤖 KI-Verbesserung",use_container_width=True):
                with st.spinner("KI verbessert Text..."):
                    result=st.session_state.analyzer.improve_text(inhalt,selected)
                    if result.get("demo_mode"): st.caption("⚠️ Demo-Modus")
                    st.text_area("Verbessert",result["text"],height=300,label_visibility="collapsed")

elif page == "⚙️ Einstellungen":
    st.title("⚙️ Einstellungen")
    c1,c2=st.columns(2)
    with c1:
        url=st.text_input("Ollama Server",db.get_setting("ollama_url","http://localhost:11434"))
        model=st.text_input("Modell",db.get_setting("ollama_model","llama3.1:8b"))
    with c2:
        behoerde=st.text_input("Behoerde",db.get_setting("behoerde","Stadtverwaltung Musterhausen"))
        abt=st.text_input("Standard-Abteilung","Buergeramt")
    if st.button("💾 Speichern"):
        for k,v in [("ollama_url",url),("ollama_model",model),("behoerde",behoerde)]: db.set_setting(k,v)
        st.success("Gespeichert!")
    st.divider()
    st.markdown("🔒 **100% DSGVO-konform | Self-Hosted | Ollama**")
