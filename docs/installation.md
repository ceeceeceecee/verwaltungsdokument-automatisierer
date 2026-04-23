# Installation – Verwaltungsdokument-Automatisierer

## Voraussetzungen

- Microsoft Office 2016 oder neuer (Word + Excel)
- Windows 10 oder 11
- Administrator-Rechte (für VBA-Makros)
- Optional: Python 3.8+ und Ollama (für KI-Optimierung)

## 🔒 Sicherheitscheckliste

- [ ] Office-Makros sind aktiviert (Vertrauenswürdige Speicherorte)
- [ ] Makro-Sicherheit auf "Warnung vor allen Makros" gestellt
- [ ] Ollama ist lokal installiert (kein Cloud-Zugriff)
- [ ] Python-Proxy läuft nur auf localhost:5000
- [ ] Keine API-Keys im Code gespeichert
- [ ] Ausgabe-Ordner ist zugriffsgeschützt
- [ ] Log-Dateien werden regelmäßig bereinigt

## Schritt 1: VBA-Makros installieren

### Makro-Sicherheit einstellen
1. Excel → Datei → Optionen → Trust Center → Trust Center-Einstellungen
2. Makroeinstellungen → "Alle Makros deaktivieren mit Benachrichtigung"
3. Vertrauenswürdige Speicherorte → Hinzufügen → Projektordner

### Makro-Code importieren
1. Excel öffnen → Alt+F11 (VBA-Editor)
2. Im Projekt-Explorer: DieseArbeitsmappe → Rechtsklick → Einfügen → Modul
3. `vba/DokumentGenerator.bas` öffnen und Code kopieren
4. Neues Modul: `vba/KIConnector.bas` Code kopieren
5. Speichern als `.xlsm`

## Schritt 2: Word-Vorlagen erstellen

1. Word öffnen
2. Briefkopf mit Behördenname erstellen
3. Platzhalter einfügen: `{{VORNAME}}`, `{{NACHNAME}}`, etc.
4. Als `.dotx` speichern im Ordner `templates/`

👉 [Vorlagen-Anleitung](vorlagen-erstellen.md)

## Schritt 3: KI-Proxy einrichten (optional)

### Ollama installieren
1. https://ollama.ai/download
2. Modell herunterladen: `ollama pull llama3`
3. Ollama starten

### Proxy starten
```bash
cd python
pip install flask
python local_ai_proxy.py
```

### Testen
```bash
curl http://localhost:5000/health
curl http://localhost:5000/api/status
```

## Schritt 4: Erste Nutzung

1. Excel-Datei öffnen
2. Daten in Zeile 2 eintragen
3. Zelle Q2 = "JA" (für KI) oder "NEIN"
4. Alt+F8 → `GenerateDokument` → Ausführen
5. Dokument wird im Ordner `output/` gespeichert

## ⚠️ Fehlerbehebung

| Problem | Lösung |
|---------|--------|
| Makros nicht sichtbar | Vertrauenswürdiger Speicherort prüfen |
| Word öffnet nicht | Word-Version kompatibel? (2016+) |
| Vorlage nicht gefunden | Pfad in Zelle Q1 prüfen |
| KI nicht erreichbar | Ollama läuft? Proxy gestartet? |
| XMLHTTP Fehler | Internet Explorer aktiviert? |
