---
name: pages
description: >-
  Operate on Apple Pages documents (.pages files) from Claude with the pages-cli
  command-line tool: create documents from templates, edit text without losing
  formatting, read a file's text without opening Pages, style text, add tables and
  images, export to PDF/Word/EPUB/text, and audit formatting consistency in large
  documents (manuscripts, reports, books). Use whenever the user mentions Pages,
  Apple Pages, pages-cli or a .pages file, or a document they are working on in Pages,
  even without naming the tool. Also use for checking formatting consistency, auditing
  styles or fonts, or looking for formatting mistakes in a large Pages document.
  Trigger: "/pages", "pages document", "create pages document", "write letter in pages",
  "report in pages", "export pages", "open pages file", "format document in pages",
  "pages erstellen", "pages dokument", "brief schreiben in pages".
---

# Apple Pages via pages-cli

`pages-cli` drives the real Apple Pages app on this Mac through AppleScript: documents,
text, tables, media, templates and export, as a command-line tool. It only works in a
session with Bash access to this Mac (macOS, Pages installed, Python 3.10+). It is not a
claude.ai chat file format and does nothing on other devices.

## Find or install the CLI

The repository (the folder holding `agent-harness/`) is two levels above this skill's
folder, so it can be found wherever it was cloned. This skill's folder is the "Base
directory" shown when the skill loads (the plugin root, `${CLAUDE_PLUGIN_ROOT}`, for a
plugin install). Use the first of these that works:

```bash
which pages-cli                                              # already on PATH
SKILL_DIR="<this skill's base directory>"
REPO="$(cd -P "$SKILL_DIR/../.." && pwd)"                    # the repository
source "$REPO/agent-harness/.venv/bin/activate"              # its virtual environment
```

If there is no virtual environment yet, create it in the repository:

```bash
cd "$REPO/agent-harness" && python3 -m venv .venv && source .venv/bin/activate && pip install -e .
```

The plugin-marketplace route needs the `claude` CLI, which a desktop session lacks, so
install manually there. With a manual clone, link this skill into your skills folder
once, from the repository root: `ln -s "$PWD/skills/pages" ~/.claude/skills/pages`.

Global flags (`--json`, `--document NAME`) go *before* the subcommand. Use `--help` on any
group (`document`, `text`, `table`, `media`, `export`, `template`, `session`) rather than
guessing flags. Commands marked **(new)** below (`text find/replace/replace-word/
replace-char/delete-range`, `document read-file`, `document save --reopen`) are missing
from older installs. Check that `pages-cli text --help` lists `find` before relying on
them, and update the install if it doesn't.

## Working rules (every Pages job)

- **Work only when asked, and only once the person has stopped editing.** Scripted changes
  go straight into the open document and can collide with unsaved work.
- **Say so before opening Pages,** so a window appearing is no surprise.
- **Prefer reading from disk.** `document read-file` shows a file's text without Pages
  running and without touching any open document.
- **Never open a backup in Pages.** Back up with `cp -R` (a .pages file can be a folder)
  and leave the copy alone.
- **Close only the document you opened; never close every document.** That can discard
  someone's unsaved work. Use `pages-cli --document "<name>" document close --no-save` on
  the one document, never a bare close. `document open` reports `was_open`: if the file was
  already open, leave it open.
- **Pages autosaves a saved document in place, and closing does not undo it.** About
  10-25 seconds after an edit Pages writes it into the file; closing with saving no does
  not revert it (measured on macOS 27.2 / Pages Creator Studio 15.4; re-check on other
  versions). The `cp` backup is the only reliable undo, so work on a copy unless the
  person wants the original changed.
- **Keep each command short.** Every CLI call has a 30-second limit. Send long output to a
  file in a scratch folder and delete what you leave there.
- **Test on a copy first** when the job is new or large.

## Command reference

**Documents:**
```bash
pages-cli --json document new [--template "Professional Report"]
pages-cli --json document open "/path/to/document.pages"      # reports was_open
pages-cli --json document info
pages-cli --json document list
pages-cli document read-file "/path/to/document.pages" --numbered   # (new) no Pages needed
pages-cli document save
pages-cli document save --path "/path/to/save.pages" [--reopen]     # --reopen (new)
pages-cli document close [--no-save]
```

**Text:**
```bash
pages-cli --json text get
pages-cli text add "New text"
pages-cli text add-paragraph "Heading" --size 24 --bold
pages-cli text set "Replace entire body"      # resets paragraph styles: not for edits
pages-cli --json text find "phrase" --ignore-case                  # (new)
pages-cli text replace "old" "new" --paragraph 3 --occurrence 1    # (new) keeps styles
pages-cli text replace-word 3 2 "new"                              # (new)
pages-cli text replace-char 3 14 ","                               # (new)
pages-cli text delete-range 3 10 20                                # (new) inclusive
pages-cli text set-font --name "Helvetica Neue" --size 14 [--paragraph N | --word N | --char N]
pages-cli text set-color --r 0 --g 0 --b 65535 --paragraph 1
pages-cli --json text word-count
```

**Tables and media:**
```bash
pages-cli table add --rows 5 --cols 3 --name "Data"
pages-cli table set-cell "Data" 1 1 "Column A"
pages-cli --json table get-cell "Data" 1 1
pages-cli table merge "Data" "A1:B1"
pages-cli table sort "Data" --column 1
pages-cli media add-image "/path/to/image.png" --x 100 --y 200
pages-cli media add-shape --type rectangle --w 200 --h 100 --text "Box"
pages-cli --json media list
```

**Export, templates, session:**
```bash
pages-cli export pdf ~/Desktop/document.pdf
pages-cli export word ~/Desktop/document.docx
pages-cli export epub ~/Desktop/book.epub --title "Title" --author "Author"
pages-cli export text ~/Desktop/document.txt
pages-cli --json export formats
pages-cli --json template list
pages-cli --json session status
```

## User request to command

| User says | Do |
|-----------|----|
| "Write a title" | `text add "Title"` + `text set-font --size 24 --paragraph N` |
| "Make the text bigger" | `text set-font --size 16` |
| "Fix this word / change X to Y" | `text find`, then `text replace` (see editing reference) |
| "Show me the text" | `document read-file` for a file, or `--json text get` for the open document |
| "Review the document" | read the text, then analyze and suggest improvements |
| "Add a table" / "Insert this image" | `table add --rows N --cols N` / `media add-image "path"` |
| "Export as PDF" | `export pdf ~/Desktop/output.pdf` |
| "How many words?" | `--json text word-count` |
| "Check the formatting" | formatting audit reference (never loop per paragraph) |

## Colors (RGB 0-65535)

Pages uses 0-65535, not 0-255: black `0 0 0`, red `65535 0 0`, green `0 65535 0`, blue
`0 0 65535`. Conversion: `round(value_0_255 / 255 * 65535)`.

## Templates (selection)

Blank (Blank, Blank Landscape, Blank Black); Reports (Simple, Modern, Professional,
Research Paper); Letters (Classic, Professional, Modern, Business); CVs (Contemporary,
Classic, Professional, Modern); Newsletters (Classic, Simple); Posters/Flyers (Photo,
Event, Type); Cards (Birthday, Photo, Party Invitation); Custom: any template the user saved
in Pages. Full list: `pages-cli --json template list`.

## Read these when the job needs them

| Job | Read |
|-----|------|
| Changing text in a document that already has content: reading it from disk, targeted edits, saving to a new path, the backup routine, what was tested | [references/editing-existing-documents.md](references/editing-existing-documents.md) |
| Exporting PDF and text and verifying them | [references/exporting-and-verifying.md](references/exporting-and-verifying.md) |
| Building a document from a template with the placeholders' styles | [references/building-from-template.md](references/building-from-template.md) |
| Checking formatting consistency in a large document (uses `scripts/analyze_formatting.py`) | [references/formatting-audit.md](references/formatting-audit.md) |
| Raw AppleScript, odd behaviour, stray files, shell quirks | [references/pitfalls.md](references/pitfalls.md) |

## Important

- Use `--json` for query commands (info, list, get, status, word-count, find).
- Pages launches automatically if it isn't running.
- If `pages-cli` is not found, activate the venv as above.
