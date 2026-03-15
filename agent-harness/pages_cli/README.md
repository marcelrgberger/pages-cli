# pages-cli

CLI harness for **Apple Pages** — create, edit, and export Pages documents from the command line or via AI agents.

## Prerequisites

- **macOS** (Pages is macOS-only)
- **Apple Pages** installed (comes pre-installed or from the App Store)
- **Python 3.10+**

## Installation

```bash
cd agent-harness
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

Verify:

```bash
which pages-cli
pages-cli --help
```

## Usage

### Interactive REPL (default)

```bash
pages-cli
```

Launches the interactive REPL with command history and styled prompts.

### One-shot commands

```bash
# Create a new document from template
pages-cli document new --template "Blank"

# Open an existing document
pages-cli document open ~/Documents/report.pages

# Add text
pages-cli text add "Hello, World!"

# Add a table
pages-cli table add --rows 5 --cols 3

# Export as PDF
pages-cli export pdf ~/Desktop/report.pdf

# Export as Word
pages-cli export word ~/Desktop/report.docx

# List templates
pages-cli template list
```

### JSON output (for agents)

```bash
pages-cli --json document info
pages-cli --json template list
pages-cli --json document list
```

## Command Reference

| Group | Command | Description |
|-------|---------|-------------|
| `document` | `new` | Create new document from template |
| `document` | `open` | Open existing .pages file |
| `document` | `close` | Close document |
| `document` | `save` | Save document |
| `document` | `info` | Show document info |
| `document` | `list` | List open documents |
| `text` | `add` | Append text |
| `text` | `set` | Set entire body text |
| `text` | `get` | Get body text |
| `text` | `set-font` | Set font name and size |
| `text` | `set-color` | Set text color (RGB) |
| `text` | `word-count` | Get word count |
| `table` | `add` | Add a table |
| `table` | `set-cell` | Set cell value |
| `table` | `get-cell` | Get cell value |
| `table` | `list` | List tables |
| `table` | `merge` | Merge cells |
| `table` | `sort` | Sort table by column |
| `media` | `add-image` | Add image from file |
| `media` | `add-shape` | Add a shape |
| `media` | `list` | List media items |
| `export` | `pdf` | Export as PDF |
| `export` | `word` | Export as Microsoft Word |
| `export` | `epub` | Export as EPUB |
| `export` | `text` | Export as plain text |
| `export` | `rtf` | Export as RTF |
| `export` | `formats` | List export formats |
| `template` | `list` | List available templates |
| `session` | `status` | Show session status |

## How it works

This CLI uses Apple Pages' built-in AppleScript support to control the application programmatically. Pages is launched automatically if needed. All operations are performed by the real Pages application — the CLI is an interface to Pages, not a replacement.

## Running tests

```bash
cd agent-harness
source .venv/bin/activate
python3 -m pytest pages_cli/pages/tests/ -v -s

# Force installed CLI command (for CI):
CLI_ANYTHING_FORCE_INSTALLED=1 python3 -m pytest pages_cli/pages/tests/ -v -s
```
