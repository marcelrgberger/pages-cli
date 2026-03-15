---
name: pages
description: >-
  Control Apple Pages documents from Claude. Create, edit, format, review,
  and export Pages documents programmatically.
  Trigger: "/pages", "pages document", "create pages document", "write letter in pages",
  "report in pages", "export pages", "open pages file", "format document in pages",
  "pages erstellen", "pages dokument", "brief schreiben in pages".
  Use this skill when the user wants to create, edit, format, review,
  or export a Pages document. Supports 100+ templates, text styling,
  tables, images, export to PDF/Word/EPUB/text.
---

# Apple Pages CLI Skill

Control Apple Pages from Claude — create, edit, format, review, and export documents.

## Prerequisites

The `pages-cli` command must be installed. If not available, install it automatically:

```bash
which pages-cli || (cd ${CLAUDE_PLUGIN_ROOT}/agent-harness && python3 -m venv .venv && source .venv/bin/activate && pip install -e . && echo "pages-cli installed")
```

If a venv already exists, activate it:
```bash
test -f ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate && source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate
```

**System requirements:**
- macOS (Apple Pages is macOS-only)
- Apple Pages installed (pre-installed or from the App Store)
- Python 3.10+

## Command Syntax

| Command | Action |
|---------|--------|
| `/pages new` | Create a new blank document |
| `/pages new <template>` | Create document from template |
| `/pages open <path>` | Open an existing document |
| `/pages info` | Show document info |
| `/pages close` | Close document |
| `/pages export pdf <path>` | Export as PDF |
| `/pages export word <path>` | Export as Word |
| `/pages templates` | List available templates |
| `/pages status` | Show session status |

## CLI Command Reference

**Document management:**
```bash
pages-cli --json document new
pages-cli --json document new --template "Professional Report"
pages-cli --json document open "/path/to/document.pages"
pages-cli --json document info
pages-cli --json document list
pages-cli document save
pages-cli document save --path "/path/to/save.pages"
pages-cli document close
pages-cli document close --no-save
```

**Text:**
```bash
pages-cli text add "New text"
pages-cli text set "Replace entire body text"
pages-cli --json text get
pages-cli text set-font --name "Helvetica Neue" --size 14
pages-cli text set-font --name "Helvetica Neue" --size 24 --paragraph 1
pages-cli text set-color --r 0 --g 0 --b 65535 --paragraph 1
pages-cli --json text word-count
```

**Tables:**
```bash
pages-cli table add --rows 5 --cols 3 --name "Data"
pages-cli table set-cell "Data" 1 1 "Column A"
pages-cli --json table get-cell "Data" 1 1
pages-cli --json table list
pages-cli table merge "Data" "A1:B1"
pages-cli table sort "Data" --column 1
```

**Media:**
```bash
pages-cli media add-image "/path/to/image.png" --x 100 --y 200
pages-cli media add-shape --type rectangle --w 200 --h 100 --text "Box"
pages-cli --json media list
```

**Export:**
```bash
pages-cli export pdf ~/Desktop/document.pdf
pages-cli export word ~/Desktop/document.docx
pages-cli export epub ~/Desktop/book.epub --title "Title" --author "Author"
pages-cli export text ~/Desktop/document.txt
pages-cli export rtf ~/Desktop/document.rtf
pages-cli --json export formats
```

**Templates & session:**
```bash
pages-cli --json template list
pages-cli --json session status
```

## Natural Language → CLI Mapping

Translate user instructions into CLI commands:

| User says | CLI command |
|-----------|------------|
| "Write a title" | `text add "Title"` + `text set-font --size 24 --paragraph N` |
| "Make the text bigger" | `text set-font --size 16` |
| "Add a table" | `table add --rows N --cols N` |
| "Insert this image" | `media add-image "path"` |
| "Export as PDF" | `export pdf ~/Desktop/output.pdf` |
| "How many words?" | `--json text word-count` |
| "Show me the text" | `--json text get` |
| "Review the document" | `--json text get` → analyze + suggest improvements |

## Colors (RGB 0-65535)

Pages uses RGB values from 0-65535 (not 0-255):
- Black: `--r 0 --g 0 --b 0`
- Red: `--r 65535 --g 0 --b 0`
- Blue: `--r 0 --g 0 --b 65535`
- Green: `--r 0 --g 65535 --b 0`

Conversion: `Pages value = round((RGB_0_255 / 255) * 65535)`

## Available Templates (Selection)

- **Blank**: Blank, Blank Landscape, Blank Black
- **Reports**: Simple Report, Modern Report, Professional Report, Research Paper
- **Letters**: Classic Letter, Professional Letter, Modern Letter, Business Letter
- **CVs**: Contemporary CV, Classic CV, Professional CV, Modern CV
- **Newsletters**: Classic Newsletter, Simple Newsletter
- **Posters/Flyers**: Photo Poster, Event Poster, Type Poster
- **Cards**: Birthday Card, Photo Card, Party Invitation
- **Custom**: All user-saved custom templates in Pages

Full list: `pages-cli --json template list`

## Important

- ALWAYS use `--json` for query commands (info, list, get, status, word-count)
- Pages launches automatically if not already running
- If CLI is not in PATH: `source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate`
