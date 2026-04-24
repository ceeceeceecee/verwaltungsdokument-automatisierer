# Verwaltungsdokument-Automatisierer – VBA + KI für Behörden

![DSGVO-Konform](https://img.shields.io/badge/DSGVO-Konform-brightgreen)
![Self-Hosted](https://img.shields.io/badge/Self--Hosted-100%25-blue)
![Office 2016-365](https://img.shields.io/badge/Office-2016--365-D04423)
![VBA](https://img.shields.io/badge/VBA-Makros-7FBB3F)
![License: MIT](https://img.shields.io/badge/Lizenz-MIT-yellow)

> 🏛️ **VBA-Makros mit KI-Integration für deutsche Behörden** — Generiert Bescheide, Anschreiben und Berichte aus Excel-Daten per Knopfdruck. Läuft in der bestehenden Office-Infrastruktur.

---

## ✨ Features

| Feature | Beschreibung |
|---------|-------------|
| 🔘 **Knopfdruck** | Dokumente direkt aus Excel mit einem Klick generieren |
| 🤖 **KI-Optimierung** | Lokale KI (Ollama) verbessert Texte automatisch |
| 📄 **Word-Vorlagen** | Beliebige Word-Vorlagen mit Platzhaltern nutzen |
| ✉️ **Serienbrief** | Massendokumente aus Datenlisten erstellen |
| 🔒 **100% Lokal** | Keine Daten verlassen den Rechner |
| 🏢 **Behördenstil** | Formatiert für amtliche Korrespondenz |
| 📋 **Protokollierung** | Alle Generierungen werden geloggt |
| 🔄 **Office 2016-365** | Kompatibel mit allen aktuellen Office-Versionen |

---

## 🏗️ Architektur

```
┌─────────────────┐    ┌──────────────┐    ┌─────────────┐
│  Excel          │───▶│  VBA Makros  │───▶│  Flask-Proxy│
│  Dateneingabe   │    │  (Generator) │    │  localhost  │
└─────────────────┘    └──────┬───────┘    └──────┬──────┘
                             │                   │
                       ┌─────▼─────┐       ┌─────▼─────┐
                       │  Word      │       │  Ollama   │
                       │  Dokument  │       │  (lokal)  │
                       └───────────┘       └───────────┘
```

---

## 🚀 Schnellstart

### Voraussetzungen

| Komponente | Version | Zweck |
|---|---|---|
| Microsoft Office | 2016+ | Excel & Word |
| Windows | 10/11 | VBA-Makro-Ausführung |
| Ollama | neueste | KI-Textoptimierung |
| Python | 3.8+ | KI-Proxy (Flask) |

### Installation

```bash
git clone https://github.com/ceeceeceecee/verwaltungsdokument-automatisierer.git
cd verwaltungsdokument-automatisierer

# KI-Proxy installieren (optional, für Ollama-Integration)
pip install flask requests
```

### Erste Schritte

1. **VBA-Makros installieren** — `vba/DokumentGenerator.bas` und `vba/KIConnector.bas` in Excel (Alt+F11 → Einfügen → Modul) importieren
2. **Excel-Datei als `.xlsm` speichern** — Makros aktivieren
3. **Word-Vorlagen anpassen** — Platzhalter: `{{VORNAME}}`, `{{NACHNAME}}`, etc.
4. **Dokument generieren** — Button in Excel klicken, Word-Dokument wird automatisch erstellt

👉 [Detaillierte Installationsanleitung](docs/installation.md)

---

## 📸 Screenshots

![Excel-Vorlage](screenshots/excel-vorlage.png)
*Excel-Tabelle mit Beispieldaten und Formular

![Word-Output](screenshots/word-output.png)
*Generiertes Bescheid-Dokument

![KI-Proxy](screenshots/ki-proxy.png)
*Flask-Proxy Terminal-Output

![Workflow](screenshots/workflow.png)
*Workflow-Diagramm

---

## 🔒 Datenschutz & DSGVO

- **100% lokal:** Keine Daten werden an externe Dienste gesendet
- **KI-Proxy:** Lokaler Flask-Server als Zwischenschicht
- **Keine Cloud:** Ollama läuft auf dem eigenen Rechner
- **Protokollierung:** Jede Dokumentgenerierung wird erfasst

---

## 🏛️ Use Cases für Behörden

| Szenario | Beschreibung |
|----------|-------------|
| **Bescheide** | Baugenehmigungen, Gebührenbescheide, Sozialbescheide |
| **Anschreiben** | Bürgeranfragen, Terminbestätigungen, Einladungen |
| **Berichte** | Monatsberichte, Jahresberichte, Statistiken |
| **Serienbriefe** | Massenbenachrichtigungen, Einladungen, Info-Schreiben |

---

## 🗺️ Roadmap

- [x] Basis-VBA-Generator
- [x] KI-Proxy (Flask + Ollama)
- [x] Word-Vorlagen-System
- [x] Serienbrief-Funktion
- [ ] PDF-Export
- [ ] Outlook-Integration
- [ ] Datensatz-Validierung
- [ ] Mehrsprachige Vorlagen

---

## 🤝 Contributing

1. Fork erstellen
2. Feature-Branch anlegen
3. Änderungen committen
4. Pull Request erstellen

---

## 📄 Lizenz

[MIT-Lizenz](LICENSE)

---

## 👤 Autor

**Cela** — Freelancer für digitale Verwaltungslösungen

[GitHub](https://github.com/ceeceeceecee)
