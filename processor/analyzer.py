"""Ollama KI fuer Verwaltungsdokumente."""
import json, requests

class DokumentAnalyzer:
    def __init__(self, url="http://localhost:11434", model="llama3.1:8b", temp=0.3, tokens=4096):
        self.url=url.rstrip("/"); self.model=model; self.temp=temp; self.tokens=tokens
    def is_available(self):
        try: return requests.get(f"{self.url}/api/tags",timeout=3).status_code==200
        except: return False
    def improve_text(self, inhalt, vorlage_name=""):
        prompt=f"""Verbessere den folgenden Verwaltungstext. Er sollte professionell, klar und behoerdengerecht formuliert sein.
Vorlage: {vorlage_name}
Text: {inhalt}
Gib nur den verbesserten Text zurueck, kein JSON."""
        if self.is_available():
            try:
                r=requests.post(f"{self.url}/api/generate",json={"model":self.model,"prompt":prompt,"stream":False,"options":{"temperature":self.temp,"num_predict":self.tokens}},timeout=120)
                return {"text":r.json().get("response","")}
            except: pass
        return {"text":inhalt,"demo_mode":True}
