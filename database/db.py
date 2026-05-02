"""SQLite DB fuer Verwaltungsdokument-Automatisierer."""
import sqlite3, json
from datetime import datetime, timedelta
from pathlib import Path

class DatabaseManager:
    def __init__(self, db_path=None):
        if db_path is None:
            base = Path(__file__).parent.parent / "data"
            base.mkdir(parents=True, exist_ok=True)
            db_path = str(base / "dokumente.db")
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create()
        self._seed()

    def _create(self):
        self.conn.execute("""CREATE TABLE IF NOT EXISTS vorlagen (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, kategorie TEXT NOT NULL,
            inhalt TEXT NOT NULL, variablen_json TEXT DEFAULT '{}', erstellt_am TEXT NOT NULL)""")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS dokumente (
            id INTEGER PRIMARY KEY AUTOINCREMENT, vorlage_id INTEGER, titel TEXT NOT NULL,
            variablen_json TEXT DEFAULT '{}', erstellt_am TEXT NOT NULL, status TEXT DEFAULT 'entwurf',
            inhalt TEXT, FOREIGN KEY(vorlage_id) REFERENCES vorlagen(id))""")
        self.conn.execute("""CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)""")
        self.conn.commit()

    def _seed(self):
        if self.conn.execute("SELECT COUNT(*) FROM vorlagen").fetchone()[0] > 0: return
        vorlagen = [
            ("Bescheid - Antrag genehmigt","Behoerde","Sehr geehrte/r {{Anrede}} {{Name}},\n\nhiermit teile ich Ihnen mit, dass Ihr Antrag vom {{Datum}} (Ref: {{Aktenzeichen}}) genehmigt wurde.\n\n{{Behoerde}} stellt Ihnen hiermit den Bescheid zu.\n\nMit freundlichen Gruessen\n{{Behoerde}}\n\n{{Zusatz}}",json.dumps(["Anrede","Name","Datum","Aktenzeichen","Behoerde","Zusatz"])),
            ("Widerspruchsbescheid","Behoerde","Widerspruchsbescheid\n\nBetreff: Ihr Widerspruch vom {{Datum}} (Ref: {{Aktenzeichen}})\n\nSehr geehrte/r {{Anrede}} {{Name}},\n\nIhr Widerspruch gegen den Bescheid vom {{Bescheid_Datum}} wurde wie folgt beschieden:\n\n{{Entscheidung}}\n\nBegruendung: {{Begruendung}}\n\nMit freundlichen Gruessen\n{{Behoerde}}",json.dumps(["Anrede","Name","Datum","Aktenzeichen","Bescheid_Datum","Entscheidung","Begruendung","Behoerde"])),
            ("Amtschreiben","Korrespondenz","{{Behoerde}}\n\n{{Abteilung}}\n\nAn: {{Anrede}} {{Name}}\n{{Adresse}}\n\nDatum: {{Datum}}\n\nBetreff: {{Betreff}}\n\nSehr geehrte/r {{Anrede}} {{Name}},\n\n{{Text}}\n\nMit freundlichen Gruessen\n{{Unterschrift}}",json.dumps(["Behoerde","Abteilung","Anrede","Name","Adresse","Datum","Betreff","Text","Unterschrift"])),
            ("Ladung","Behoerde","Ladung\n\nAn: {{Anrede}} {{Name}}\n{{Adresse}}\n\nSehr geehrte/r {{Anrede}} {{Name}},\n\nSie werden hiermit zu folgender Sitzung geladen:\n\nTermin: {{Datum}} um {{Uhrzeit}} Uhr\nOrt: {{Ort}}\nGremium: {{Gremium}}\n\nTagesordnungspunkt: {{TOP}}\n\nBitte erscheinen Sie puenktlich oder entschuldigen Sie sich rechtzeitig.\n\n{{Behoerde}}",json.dumps(["Anrede","Name","Adresse","Datum","Uhrzeit","Ort","Gremium","TOP","Behoerde"])),
            ("Bescheinigung","Korrespondenz","Bescheinigung\n\nHiermit wird bescheinigt, dass\n\n{{Anrede}} {{Name}}\nGeboren am: {{Geburtsdatum}} in {{Geburtsort}}\n\n{{Beschreibung}}\n\nMusterhausen, den {{Datum}}\n\n{{Unterschrift}}\n{{Behoerde}}",json.dumps(["Anrede","Name","Geburtsdatum","Geburtsort","Beschreibung","Datum","Unterschrift","Behoerde"])),
            ("Auskunft","Korrespondenz","Auskunft\n\nIhre Anfrage vom {{Anfrage_Datum}}:\n\n{{Frage}}\n\nAntwort:\n{{Antwort}}\n\nMusterhausen, den {{Datum}}\n\n{{Behoerde}}",json.dumps(["Anfrage_Datum","Frage","Antwort","Datum","Behoerde"])),
            ("Mahnung","Finanzen","Mahnung\n\nRef: {{Aktenzeichen}}\n\nSehr geehrte/r {{Anrede}} {{Name}},\n\nSie haben die folgende Forderung noch nicht beglichen:\n\nBetrag: {{Betrag}} EUR\nFaellig seit: {{Faelligkeit}}\n\nWir fordern Sie auf, den Betrag bis zum {{Neue_Faelligkeit}} zu ueberweisen.\n\nBei Nichtbeachtung behalten wir uns weitere Massnahmen vor.\n\n{{Behoerde}}",json.dumps(["Aktenzeichen","Anrede","Name","Betrag","Faelligkeit","Neue_Faelligkeit","Behoerde"])),
            ("Bestaetigung","Korrespondenz","Bestaetigung\n\nSehr geehrte/r {{Anrede}} {{Name}},\n\nhiermit bestaetigen wir Ihnen:\n\n{{Bestaetigungstext}}\n\nDatum: {{Datum}}\nRef: {{Aktenzeichen}}\n\nMit freundlichen Gruessen\n{{Behoerde}}",json.dumps(["Anrede","Name","Bestaetigungstext","Datum","Aktenzeichen","Behoerde"])),
            ("Anzeige","Behoerde","Anzeige\n\nAn {{Datum}} wurde durch {{Beamter}} folgende Ordnungswidrigkeit festgestellt:\n\nAdresse: {{Adresse}}\nTatbestand: {{Tatbestand}}\n\nSie haben die Moeglichkeit, zu dem Vorwurf schriftlich Stellung zu nehmen bis zum {{Frist}}.\n\n{{Behoerde}}",json.dumps(["Datum","Beamter","Adresse","Tatbestand","Frist","Behoerde"])),
            ("Mitteilung","Korrespondenz","Mitteilung\n\nAn: {{Anrede}} {{Name}}\n{{Adresse}}\n\n{{Mitteilungstext}}\n\nMusterhausen, den {{Datum}}\n\n{{Behoerde}}",json.dumps(["Anrede","Name","Adresse","Mitteilungstext","Datum","Behoerde"])),
        ]
        for name,kat,inhalt,vars in vorlagen:
            d=(datetime.now()-timedelta(days=random.randint(30,180))).strftime("%Y-%m-%d")
            self.conn.execute("INSERT INTO vorlagen (name,kategorie,inhalt,variablen_json,erstellt_am) VALUES (?,?,?,?,?)",(name,kat,inhalt,vars,d))
        # 5 generierte Dokumente
        docs = [
            (1,"Bescheid Mueller","{\"Anrede\":\"Herr\",\"Name\":\"Thomas Mueller\",\"Datum\":\"01.05.2026\",\"Aktenzeichen\":\"BA-2026-0147\",\"Behoerde\":\"Stadtverwaltung Musterhausen\",\"Zusatz\":\"Der Bescheid ergeht vollumfaenglich.\"}","fertig","Sehr geehrte/r Herr Thomas Mueller,..."),
            (3,"Amtschreiben Weber","{\"Behoerde\":\"Stadtverwaltung\",\"Abteilung\":\"Buergeramt\",\"Anrede\":\"Frau\",\"Name\":\"Anna Weber\",\"Adresse\":\"Hauptstr. 22\",\"Datum\":\"28.04.2026\",\"Betreff\":\"Meldebescheinigung\",\"Text\":\"Ihre Meldebescheinigung liegt bei...\",\"Unterschrift\":\"Buergeramt\"}","fertig","Stadtverwaltung..."),
            (4,"Ladung Gemeinderat","{\"Anrede\":\"Herr\",\"Name\":\"Buergermeister Schmidt\",\"Adresse\":\"Rathaus\",\"Datum\":\"15.05.2026\",\"Uhrzeit\":\"18:00\",\"Ort\":\"Ratssaal\",\"Gremium\":\"Gemeinderat\",\"TOP\":\"Haushaltsplan 2027\",\"Behoerde\":\"Stadtverwaltung\"}","entwurf","Ladung..."),
            (5,"Bescheinigung Fischer","{\"Anrede\":\"Frau\",\"Name\":\"Petra Fischer\",\"Geburtsdatum\":\"15.03.1985\",\"Geburtsort\":\"Musterhausen\",\"Beschreibung\":\"...ist wohnhaft in Musterhausen\",\"Datum\":\"20.04.2026\",\"Unterschrift\":\"Buergeramt\",\"Behoerde\":\"Stadtverwaltung\"}","fertig","Bescheinigung..."),
            (7,"Mahnung Koch","{\"Aktenzeichen\":\"ST-2026-0089\",\"Anrede\":\"Herr\",\"Name\":\"Hans Koch\",\"Betrag\":\"245,00\",\"Faelligkeit\":\"01.03.2026\",\"Neue_Faelligkeit\":\"15.05.2026\",\"Behoerde\":\"Stadtverwaltung Musterhausen\"}","entwurf","Mahnung..."),
        ]
        for vid,titel,vars,status,inhalt in docs:
            d=(datetime.now()-timedelta(days=random.randint(1,14))).strftime("%Y-%m-%d")
            self.conn.execute("INSERT INTO dokumente (vorlage_id,titel,variablen_json,erstellt_am,status,inhalt) VALUES (?,?,?,?,?,?)",(vid,titel,vars,d,status,inhalt))
        defaults={"ollama_url":"http://localhost:11434","ollama_model":"llama3.1:8b","behoerde":"Stadtverwaltung Musterhausen","datenspeicherung":"Lokal","verschluesselung":"AES-256"}
        for k,v in defaults.items(): self.conn.execute("INSERT OR IGNORE INTO settings VALUES (?,?)",(k,v))
        self.conn.commit()

    def get_all_vorlagen(self): return [dict(r) for r in self.conn.execute("SELECT * FROM vorlagen ORDER BY name").fetchall()]
    def get_vorlage(self, vid): r=self.conn.execute("SELECT * FROM vorlagen WHERE id=?",(vid,)).fetchone(); return dict(r) if r else None
    def get_all_dokumente(self): return [dict(r) for r in self.conn.execute("SELECT d.*, v.name as vorlage_name FROM dokumente d LEFT JOIN vorlagen v ON d.vorlage_id=v.id ORDER BY d.erstellt_am DESC").fetchall()]
    def create_dokument(self, vid, titel, vars, inhalt):
        d=datetime.now().strftime("%Y-%m-%d %H:%M")
        self.conn.execute("INSERT INTO dokumente (vorlage_id,titel,variablen_json,erstellt_am,status,inhalt) VALUES (?,?,?,?,?,?)",(vid,titel,json.dumps(vars),d,"entwurf",inhalt))
        self.conn.commit()
    def create_vorlage(self, name, kat, inhalt, vars):
        d=datetime.now().strftime("%Y-%m-%d")
        self.conn.execute("INSERT INTO vorlagen (name,kategorie,inhalt,variablen_json,erstellt_am) VALUES (?,?,?,?,?)",(name,kat,inhalt,json.dumps(vars),d))
        self.conn.commit()
    def get_setting(self, k, d=""): r=self.conn.execute("SELECT value FROM settings WHERE key=?", (k,)).fetchone(); return r[0] if r else d
    def set_setting(self, k, v): self.conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)",(k,v)); self.conn.commit()
    def get_stats(self):
        v=self.conn.execute("SELECT COUNT(*) FROM vorlagen").fetchone()[0]
        d=self.conn.execute("SELECT COUNT(*) FROM dokumente").fetchone()[0]
        return {"vorlagen":v,"dokumente":d,"entwurf":self.conn.execute("SELECT COUNT(*) FROM dokumente WHERE status='entwurf'").fetchone()[0],"fertig":self.conn.execute("SELECT COUNT(*) FROM dokumente WHERE status='fertig'").fetchone()[0]}
    def get_category_counts(self): return dict(self.conn.execute("SELECT kategorie,COUNT(*) FROM vorlagen GROUP BY kategorie").fetchall())
