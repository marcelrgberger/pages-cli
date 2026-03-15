# pages-cli

Claude Code plugin for full programmatic control over **Apple Pages**. Create, edit, format, and export Pages documents using slash commands, CLI, or natural language.

Covers 100% of the Pages AppleScript API — every command, class, property, and enum from the Pages scripting definition (sdef).

## Features

- **Documents** — create (100+ templates incl. custom), open, close, save, password protect
- **Text** — add, replace, style per paragraph/word/character (font, size, color, bold, italic)
- **Tables** — create, fill, formulas, cell formatting, range styling (alignment, colors, wrap), merge/unmerge/sort/clear, row height, column width, header/footer rows & columns
- **Media** — images (with alt text), shapes (with text), text boxes, audio clips, movies, lines, groups — rotation, opacity, reflection, lock, position, size
- **Export** — PDF, Word, EPUB, plain text, RTF, Pages 09 — with password, image quality, comments, annotations, full EPUB metadata
- **Templates** — list all (incl. custom), create from template, placeholder filling
- **Sections & Pages** — multi-section/page body text access
- **Session** — state tracking, history, REPL with prompt_toolkit
- **JSON output** — `--json` flag on all commands for agent integration

## Requirements

- **macOS** with **Apple Pages** (pre-installed or [App Store](https://apps.apple.com/app/pages/id409201541))
- **Python 3.10+**
- **Claude Code** CLI (for plugin usage)

## Installation

### Via Claude Code

```bash
claude plugins marketplace add marcelrgberger/pages-cli
claude plugins install pages-cli
```

The CLI backend installs automatically on first `/pages` use.

### Manual

```bash
git clone https://github.com/marcelrgberger/pages-cli.git
cd pages-cli/agent-harness
python3 -m venv .venv && source .venv/bin/activate
pip install -e .
```

## Quick Start

```
/pages new Professional Report     Create from template
/pages open ~/doc.pages            Open existing document
/pages export pdf ~/Desktop/out.pdf Export as PDF
/pages templates                   Browse 100+ templates
```

After opening a document, use natural language:

```
User: Write a title "Q1 Report" in 24pt bold, then add 3 paragraphs about revenue
User: Add a 5x3 table with headers Name, Amount, Date
User: Set the header row background to blue
User: Export as PDF and Word to the Desktop
```

## Command Reference

### Document
| Command | Description |
|---------|-------------|
| `document new [--template T] [--name N]` | Create document |
| `document open <path>` | Open .pages file |
| `document close [--no-save]` | Close document |
| `document save [--path P]` | Save document |
| `document info` | Name, pages, words, characters |
| `document list` | List open documents |
| `document set-password <pw> [--hint H]` | Password protect |
| `document remove-password <pw>` | Remove password |
| `document placeholders` | List template placeholders |
| `document set-placeholder <idx> <tag>` | Set placeholder tag |
| `document get-placeholder <idx>` | Get placeholder tag |
| `document sections` | Section count |
| `document section-text <idx>` | Get section body text |
| `document page-text <idx>` | Get page body text |
| `document delete <type> <idx>` | Delete object on page |

### Text
| Command | Description |
|---------|-------------|
| `text add <text>` | Append text |
| `text add-paragraph <text> [styling opts]` | Add styled paragraph |
| `text set <text>` | Replace all body text |
| `text get` | Get body text |
| `text set-font [--name F] [--size S] [--paragraph/--word/--char N]` | Set font |
| `text get-font [--paragraph/--word/--char N]` | Get font name |
| `text get-font-size [--paragraph/--word/--char N]` | Get font size |
| `text set-color --r R --g G --b B [--paragraph/--word/--char N]` | Set color (0-65535) |
| `text get-color [--paragraph/--word/--char N]` | Get color |
| `text bold [--paragraph/--word N]` | Set bold |
| `text italic [--paragraph/--word N]` | Set italic |
| `text bold-italic [--paragraph/--word N]` | Set bold+italic |
| `text style-paragraph <idx> [all styling opts]` | Style paragraph |
| `text style-word <idx> [all styling opts]` | Style word |
| `text style-character <idx> [all styling opts]` | Style character |
| `text word-count` | Words, characters, paragraphs |

### Table
| Command | Description |
|---------|-------------|
| `table add [--rows N] [--cols N] [--name N] [--headers N]` | Create table |
| `table delete <name>` | Delete table |
| `table set-cell <table> <row> <col> <value>` | Set cell value |
| `table get-cell <table> <row> <col>` | Get cell value |
| `table get-formula <table> <row> <col>` | Get cell formula |
| `table get-formatted-value <table> <row> <col>` | Get formatted value |
| `table list` | List all tables |
| `table info <table>` | Table details |
| `table get-data <table>` | Dump all cells as JSON |
| `table fill <table> --data '[[...]]'` | Fill from JSON |
| `table merge <table> <range>` | Merge cells |
| `table unmerge <table> <range>` | Unmerge cells |
| `table clear <table> <range>` | Clear range |
| `table sort <table> --column N [--descending]` | Sort by column |
| `table set-cell-format <table> <row> <col> --format F` | Cell format |
| `table set-header-rows <table> <count>` | Header row count |
| `table set-header-cols <table> <count>` | Header column count |
| `table set-footer-rows <table> <count>` | Footer row count |
| `table set-row-height/get-row-height` | Row height |
| `table set-col-width/get-col-width` | Column width |
| `table range-set-font/alignment/text-color/bg-color/text-wrap/format` | Range styling |

### Media
| Command | Description |
|---------|-------------|
| `media add-image <path> [--x X --y Y --w W --h H]` | Add image |
| `media add-shape [--type T --x X --y Y --w W --h H --text T]` | Add shape |
| `media add-text-item <text> [position/size]` | Add text box |
| `media add-audio <path> [--volume V]` | Add audio clip |
| `media add-movie <path> [position/volume]` | Add movie |
| `media add-line [--start-x/y --end-x/y]` | Add line |
| `media delete <type> <index>` | Delete item |
| `media set-rotation/set-opacity/set-reflection/set-locked` | Item properties |
| `media set-position/set-size` | Position & size |
| `media set-image-description` | VoiceOver alt text |
| `media set-audio-volume/set-audio-repeat` | Audio properties |
| `media set-movie-volume/set-movie-repeat` | Movie properties |
| `media set-line-start/set-line-end` | Line endpoints |
| `media properties <type> <index>` | Get all properties |
| `media list [--type T]` | List items |

### Export
| Command | Description |
|---------|-------------|
| `export pdf <path> [--password P --hint H --image-quality Q --include-comments --include-annotations]` | Export PDF |
| `export word <path> [shared options]` | Export Word |
| `export epub <path> [--title --author --genre --language --publisher --cover --fixed-layout] [shared options]` | Export EPUB |
| `export text <path> [shared options]` | Export plain text |
| `export rtf <path> [shared options]` | Export RTF |
| `export pages09 <path> [shared options]` | Export Pages 09 |
| `export formats` | List all formats |

### Template & Session
| Command | Description |
|---------|-------------|
| `template list` | List all templates |
| `template info <name>` | Template details |
| `session status` | Session state |
| `session save <path>` | Save session |
| `session load <path>` | Load session |

## Architecture

```
pages-cli/
├── .claude-plugin/           Plugin metadata + marketplace
├── commands/pages.md         /pages slash command
├── skills/pages/SKILL.md     NLP skill
├── agent-harness/
│   ├── setup.py
│   └── pages_cli/
│       ├── pages_cli.py      2200+ lines Click CLI + REPL
│       ├── core/             document, text, tables, media, export, templates, session
│       ├── utils/            AppleScript backend, shared helpers, REPL skin
│       └── tests/            Unit + E2E tests
└── README.md
```

Controls Pages through its native **AppleScript API** (`osascript`). All operations are performed by the real Apple Pages — the CLI is an interface, not a replacement.

## License

MIT
