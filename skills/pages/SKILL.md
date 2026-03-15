---
name: pages
description: >-
  Steuert Apple Pages Dokumente vollständig aus Claude heraus.
  Trigger: "/pages", "pages dokument", "pages erstellen", "seiten dokument",
  "brief schreiben in pages", "report in pages", "pages exportieren",
  "pages öffnen", "dokument formatieren in pages", "Pages document",
  "create pages document", "open pages file", "export pages".
  Nutze diesen Skill wenn der User ein Pages-Dokument erstellen, bearbeiten,
  formatieren, reviewen oder exportieren möchte. Unterstützt Vorlagen,
  Text-Styling, Tabellen, Bilder, Export als PDF/Word/EPUB/Text.
---

# Apple Pages CLI Skill

Steuere Apple Pages vollständig aus Claude heraus — Dokumente erstellen, bearbeiten, formatieren, reviewen und exportieren.

## Voraussetzungen

Die CLI `cli-anything-pages` muss installiert sein. Falls nicht vorhanden, installiere sie automatisch:

```bash
which cli-anything-pages || (cd ${CLAUDE_PLUGIN_ROOT}/agent-harness && python3 -m venv .venv && source .venv/bin/activate && pip install -e . && echo "cli-anything-pages installed")
```

Falls eine venv existiert, aktiviere sie:
```bash
test -f ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate && source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate
```

**Systemanforderungen:**
- macOS (Apple Pages ist macOS-only)
- Apple Pages muss installiert sein (vorinstalliert oder App Store)
- Python 3.10+

## Aufruf-Syntax

| Befehl | Aktion |
|--------|--------|
| `/pages new` | Neues leeres Dokument erstellen |
| `/pages new <Vorlage>` | Dokument aus Vorlage erstellen |
| `/pages open <Pfad>` | Bestehendes Dokument öffnen |
| `/pages info` | Dokumentinfo anzeigen |
| `/pages close` | Dokument schließen |
| `/pages export pdf <Pfad>` | Als PDF exportieren |
| `/pages export word <Pfad>` | Als Word exportieren |
| `/pages templates` | Verfügbare Vorlagen anzeigen |
| `/pages status` | Session-Status anzeigen |

## CLI-Befehle Referenz

**Dokument-Management:**
```bash
cli-anything-pages --json document new
cli-anything-pages --json document new --template "Professional Report"
cli-anything-pages --json document open "/pfad/zum/dokument.pages"
cli-anything-pages --json document info
cli-anything-pages --json document list
cli-anything-pages document save
cli-anything-pages document save --path "/pfad/speichern.pages"
cli-anything-pages document close
cli-anything-pages document close --no-save
```

**Text:**
```bash
cli-anything-pages text add "Neuer Text"
cli-anything-pages text set "Gesamten Text ersetzen"
cli-anything-pages --json text get
cli-anything-pages text set-font --name "Helvetica Neue" --size 14
cli-anything-pages text set-font --name "Helvetica Neue" --size 24 --paragraph 1
cli-anything-pages text set-color --r 0 --g 0 --b 65535 --paragraph 1
cli-anything-pages --json text word-count
```

**Tabellen:**
```bash
cli-anything-pages table add --rows 5 --cols 3 --name "Daten"
cli-anything-pages table set-cell "Daten" 1 1 "Spalte A"
cli-anything-pages --json table get-cell "Daten" 1 1
cli-anything-pages --json table list
cli-anything-pages table merge "Daten" "A1:B1"
cli-anything-pages table sort "Daten" --column 1
```

**Medien:**
```bash
cli-anything-pages media add-image "/pfad/bild.png" --x 100 --y 200
cli-anything-pages media add-shape --type rectangle --w 200 --h 100 --text "Box"
cli-anything-pages --json media list
```

**Export:**
```bash
cli-anything-pages export pdf ~/Desktop/dokument.pdf
cli-anything-pages export word ~/Desktop/dokument.docx
cli-anything-pages export epub ~/Desktop/buch.epub --title "Titel" --author "Autor"
cli-anything-pages export text ~/Desktop/dokument.txt
cli-anything-pages export rtf ~/Desktop/dokument.rtf
cli-anything-pages --json export formats
```

**Vorlagen & Session:**
```bash
cli-anything-pages --json template list
cli-anything-pages --json session status
```

## Natürliche Sprache → CLI

Übersetze User-Anweisungen in CLI-Befehle:

| User sagt | CLI-Befehl |
|-----------|------------|
| "Schreibe einen Titel" | `text add "Titel"` + `text set-font --size 24 --paragraph N` |
| "Mach den Text größer" | `text set-font --size 16` |
| "Füge eine Tabelle ein" | `table add --rows N --cols N` |
| "Füge das Bild ein" | `media add-image "Pfad"` |
| "Exportiere als PDF" | `export pdf ~/Desktop/output.pdf` |
| "Wie viele Wörter?" | `--json text word-count` |
| "Zeig mir den Text" | `--json text get` |
| "Prüfe das Dokument" | `--json text get` → analysieren + Vorschläge |

## Farben (RGB 0-65535)

Pages verwendet RGB-Werte von 0-65535:
- Schwarz: `--r 0 --g 0 --b 0`
- Rot: `--r 65535 --g 0 --b 0`
- Blau: `--r 0 --g 0 --b 65535`
- Grün: `--r 0 --g 65535 --b 0`

Umrechnung: `Pages-Wert = round((RGB-0-255-Wert / 255) * 65535)`

## Verfügbare Vorlagen (Auswahl)

- **Leer**: Blank, Blank Landscape, Blank Black
- **Berichte**: Simple Report, Modern Report, Professional Report, Research Paper
- **Briefe**: Classic Letter, Professional Letter, Modern Letter, Business Letter
- **CVs**: Contemporary CV, Classic CV, Professional CV, Modern CV
- **Newsletter**: Classic Newsletter, Simple Newsletter
- **Poster/Flyer**: Photo Poster, Event Poster, Type Poster
- **Karten**: Birthday Card, Photo Card, Party Invitation
- **Eigene**: Alle in Pages gespeicherten eigenen Vorlagen

Vollständige Liste: `cli-anything-pages --json template list`

## Wichtig

- Verwende IMMER `--json` bei Abfrage-Befehlen (info, list, get, status, word-count)
- Pages wird automatisch gestartet falls nicht bereits offen
- Falls CLI nicht im PATH: `source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate`
