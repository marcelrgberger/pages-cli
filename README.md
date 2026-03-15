# pages-cli

Claude Code plugin for full control over **Apple Pages**. Create, edit, format, and export Pages documents using slash commands or natural language.

## Features

- Create documents from 100+ templates (including custom ones)
- Add and format text — font, size, color per paragraph
- Insert and edit tables — cells, sorting, merging
- Add images and shapes
- Export to PDF, Word, EPUB, plain text, RTF
- Interactive REPL mode
- JSON output for agent integration
- Natural language interaction after opening a document

## Requirements

- **macOS** (Apple Pages is macOS-only)
- **Apple Pages** installed (pre-installed or from the [App Store](https://apps.apple.com/app/pages/id409201541))
- **Python 3.10+**
- **Claude Code** CLI

## Installation

### Via Claude Code

```bash
claude plugins marketplace add marcelrgberger/pages-cli
claude plugins install pages-cli
```

The CLI backend is installed automatically on first use of `/pages`.

### Manual

```bash
git clone https://github.com/marcelrgberger/pages-cli.git
cd pages-cli/agent-harness
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Usage

### Slash Commands

```
/pages                              Create a new blank document
/pages new                          Create a new blank document
/pages new Professional Report      Create from template
/pages open ~/Documents/report.pages Open existing document
/pages info                         Show document info
/pages export pdf ~/Desktop/out.pdf Export as PDF
/pages export word ~/Desktop/out.docx Export as Word
/pages templates                    List available templates
/pages close                        Close document
/pages status                       Show session status
```

### Natural Language

After `/pages new` or `/pages open`, just tell Claude what to do:

```
User: /pages new Professional Report
Claude: Created document from "Professional Report" template. What would you like me to do?

User: Write a title "Q1 2026 Quarterly Report" and a summary paragraph below it
Claude: [adds text, formats the title in 24pt]

User: Add a table with revenue figures, 4 rows and 3 columns
Claude: [creates table, fills headers]

User: Export as PDF to the Desktop
Claude: Exported: ~/Desktop/Quarterly_Report.pdf (12,340 bytes)
```

### Direct CLI

```bash
pages-cli --help
pages-cli --json document new --template "Blank"
pages-cli text add "Hello World"
pages-cli export pdf ~/Desktop/test.pdf
pages-cli  # Launches interactive REPL
```

## Available Templates (Selection)

| Category | Templates |
|----------|-----------|
| Blank | Blank, Blank Landscape, Blank Black |
| Reports | Simple Report, Modern Report, Professional Report, Research Paper |
| Letters | Classic Letter, Professional Letter, Modern Letter, Business Letter |
| CVs | Contemporary CV, Classic CV, Professional CV, Modern CV |
| Newsletters | Classic Newsletter, Simple Newsletter |
| Posters | Photo Poster, Event Poster, Type Poster |
| Cards | Birthday Card, Photo Card, Party Invitation |

Plus all custom templates saved in Pages. Full list: `/pages templates`

## Export Formats

| Format | Command | Extension |
|--------|---------|-----------|
| PDF | `/pages export pdf <path>` | .pdf |
| Microsoft Word | `/pages export word <path>` | .docx |
| EPUB | `/pages export epub <path>` | .epub |
| Plain Text | `/pages export text <path>` | .txt |
| Rich Text | `/pages export rtf <path>` | .rtf |

## Architecture

```
pages-cli/
├── .claude-plugin/plugin.json    Plugin metadata
├── commands/pages.md             /pages slash command
├── skills/pages/SKILL.md         Skill with NLP mapping
├── agent-harness/                CLI backend
│   ├── setup.py                  PyPI package config
│   └── pages_cli/                Python modules
│       ├── pages_cli.py          Click CLI + REPL
│       ├── core/                 document, text, tables, media, export, templates, session
│       ├── utils/                AppleScript backend, REPL skin
│       └── tests/                42 tests (unit + E2E)
└── README.md
```

The CLI controls Pages through its native **AppleScript API** (`osascript`). All operations are performed by the real Apple Pages application — the CLI is an interface to Pages, not a replacement.

## How It Works

1. The plugin provides a `/pages` slash command and a skill that triggers on natural language
2. Commands are translated into AppleScript calls via `osascript`
3. Pages is launched automatically if not already running
4. All document manipulation happens in the real Pages application
5. Export produces real files (PDF with `%PDF-` magic bytes, Word as valid OOXML ZIP, etc.)

## Tests

```bash
cd agent-harness
source .venv/bin/activate
python3 -m pytest pages_cli/tests/ -v -s
```

42 tests (27 unit + 15 E2E), 100% pass rate. E2E tests create real documents and export real PDF/Word files with format verification.

## License

MIT
