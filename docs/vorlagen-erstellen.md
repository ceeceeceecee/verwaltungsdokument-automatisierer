# Word-Vorlagen erstellen

## Vorlagen-System

Das System verwendet Platzhalter im Format `{{PLATZHALTER}}`, die automatisch durch die Excel-Daten ersetzt werden.

## Erstellen einer neuen Vorlage

### Schritt 1: Word-Dokument erstellen

1. Microsoft Word öffnen
2. Neues Dokument erstellen
3. Seitenränder: 2,5 cm alle Seiten

### Schritt 2: Briefkopf

```
Stadt Musterhausen
Bauamt
Rathausplatz 1
12345 Musterhausen
```

Optional: Logo/Wappen einfügen

### Schritt 3: Platzhalter einfügen

Verwenden Sie exakt folgende Platzhalter:

**Empfängerdaten:**
- `{{ANREDE}}` – Sehr geehrte Frau/Herr
- `{{VORNAME}}` – Vorname
- `{{NACHNAME}}` – Nachname
- `{{STRASSE}}` – Straße und Hausnummer
- `{{PLZ}}` – Postleitzahl
- `{{ORT}}` – Ort

**Metadaten:**
- `{{DATUM}}` – Datum
- `{{AKTENZEICHEN}}` – Aktenzeichen
- `{{BEHOERDE}}` – Behördenname
- `{{ABTEILUNG}}` – Abteilung
- `{{BETREFF}}` – Betreffzeile

**Inhalt:**
- `{{INHALT}}` – Haupttext (wird ggf. von KI optimiert)
- `{{FRIST}}` – Fristangabe
- `{{BETRAG}}` – Geldbetrag
- `{{KONTAKT}}` – Kontaktinformationen
- `{{GEBURTSDATUM}}` – Geburtsdatum

### Schritt 4: Formatierung

- Schriftart: Arial, 11pt
- Absatzformat: Blocksatz
- Zeilenabstand: 1,5
- Ausrichtung: Linksbündig (Text), Rechtsbündig (Datum/Aktenzeichen)

### Schritt 5: Speichern

1. Datei → Speichern unter
2. Dateityp: Word-Vorlage (*.dotx)
3. Speicherort: `templates/` im Projektordner
4. Dateiname: z.B. `Bescheid-Vorlage.dotx`

## Tipps

- Platzhalter genau so schreiben (Großbuchstaben, doppelte geschweifte Klammern)
- Testen Sie die Vorlage mit Beispieldaten
- Erstellen Sie separate Vorlagen für verschiedene Dokumenttypen
- Verwenden Sie Word-Formatvorlagen für einheitliches Erscheinungsbild
