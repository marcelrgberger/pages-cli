# pages-cli

Claude Code Plugin zur vollständigen Steuerung von **Apple Pages** direkt aus Claude heraus. Erstelle, bearbeite, formatiere und exportiere Pages-Dokumente per Slash-Command oder natürlicher Sprache.

## Features

- Dokumente erstellen (aus 100+ Vorlagen inkl. eigener)
- Text hinzufügen, formatieren, Schriftart/Farbe setzen
- Tabellen einfügen, Zellen bearbeiten, sortieren, mergen
- Bilder und Shapes einfügen
- Export als PDF, Word, EPUB, Text, RTF
- Interaktiver REPL-Modus
- JSON-Output für Agent-Integration
- Funktioniert mit natürlicher Sprache nach dem Öffnen

## Systemanforderungen

- **macOS** (Apple Pages ist macOS-only)
- **Apple Pages** installiert (vorinstalliert oder [App Store](https://apps.apple.com/app/pages/id409201541))
- **Python 3.10+**
- **Claude Code** CLI

## Installation

### In Claude Code

```bash
claude plugins add marcelrgberger/pages-cli
```

Beim ersten Aufruf von `/pages` wird die CLI automatisch installiert.

### Manuell

```bash
git clone https://github.com/marcelrgberger/pages-cli.git
cd pages-cli/agent-harness
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Verwendung

### Slash-Commands

```
/pages                              Neues leeres Dokument
/pages new                          Neues leeres Dokument
/pages new Professional Report      Dokument aus Vorlage
/pages open ~/Documents/brief.pages Dokument öffnen
/pages info                         Dokumentinfo anzeigen
/pages export pdf ~/Desktop/out.pdf Als PDF exportieren
/pages export word ~/Desktop/out.docx Als Word exportieren
/pages templates                    Vorlagen anzeigen
/pages close                        Schließen
/pages status                       Session-Status
```

### Natürliche Sprache

Nach `/pages new` oder `/pages open` einfach sagen was passieren soll:

```
User: /pages new Professional Report
Claude: Dokument "Professional Report" erstellt. Was soll ich damit machen?

User: Schreibe einen Titel "Quartalsreport Q1 2026" und darunter eine Zusammenfassung
Claude: [fügt Text ein, formatiert den Titel größer]

User: Füge eine Tabelle mit Umsatzzahlen ein, 4 Zeilen, 3 Spalten
Claude: [erstellt Tabelle, füllt Header]

User: Exportiere das als PDF auf den Desktop
Claude: PDF exportiert: ~/Desktop/Quartalsreport.pdf (12,340 bytes)
```

### CLI direkt

```bash
cli-anything-pages --help
cli-anything-pages --json document new --template "Blank"
cli-anything-pages text add "Hello World"
cli-anything-pages export pdf ~/Desktop/test.pdf
cli-anything-pages  # Startet interaktiven REPL-Modus
```

## Verfügbare Vorlagen (Auswahl)

| Kategorie | Vorlagen |
|-----------|----------|
| Leer | Blank, Blank Landscape, Blank Black |
| Berichte | Simple Report, Modern Report, Professional Report, Research Paper |
| Briefe | Classic Letter, Professional Letter, Modern Letter, Business Letter |
| CVs | Contemporary CV, Classic CV, Professional CV, Modern CV |
| Newsletter | Classic Newsletter, Simple Newsletter |
| Poster | Photo Poster, Event Poster, Type Poster |
| Karten | Birthday Card, Photo Card, Party Invitation |

Plus alle eigenen in Pages gespeicherten Vorlagen. Vollständige Liste: `/pages templates`

## Export-Formate

| Format | Befehl | Erweiterung |
|--------|--------|-------------|
| PDF | `/pages export pdf <pfad>` | .pdf |
| Microsoft Word | `/pages export word <pfad>` | .docx |
| EPUB | `/pages export epub <pfad>` | .epub |
| Klartext | `/pages export text <pfad>` | .txt |
| Rich Text | `/pages export rtf <pfad>` | .rtf |

## Architektur

```
pages-cli/
├── .claude-plugin/plugin.json    Plugin-Metadaten
├── commands/pages.md             /pages Slash-Command
├── skills/pages/SKILL.md         Skill mit NLP-Mapping
├── agent-harness/                CLI-Backend
│   ├── setup.py                  PyPI-Package
│   └── cli_anything/pages/       Python-Module
│       ├── pages_cli.py          Click CLI + REPL
│       ├── core/                 document, text, tables, media, export, templates, session
│       ├── utils/                AppleScript-Backend, REPL-Skin
│       └── tests/                42 Tests (Unit + E2E)
└── README.md
```

Die CLI steuert Pages über dessen native **AppleScript-API** (`osascript`). Alle Operationen werden vom echten Apple Pages ausgeführt — die CLI ist ein Interface zu Pages, kein Ersatz.

## Tests

```bash
cd agent-harness
source .venv/bin/activate
python3 -m pytest cli_anything/pages/tests/ -v -s
```

42 Tests (27 Unit + 15 E2E), 100% Pass Rate. E2E-Tests erstellen echte Dokumente und exportieren echte PDFs/Word-Dateien.

## Lizenz

MIT
