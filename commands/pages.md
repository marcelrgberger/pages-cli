---
name: pages
description: "Control Apple Pages — /pages new, /pages open <path>, /pages export pdf <path>, /pages templates, /pages info, /pages close, /pages status"
arguments:
  - name: action
    description: "Action: new [template], open <path>, info, close, export <format> <path>, templates, status"
    required: false
---

# Pages CLI — Quick Access

You have access to the Apple Pages CLI (`pages-cli`).

## Step 0: Ensure CLI is available

Check if `pages-cli` is installed. If not, install it automatically:

```bash
which pages-cli || (cd ${CLAUDE_PLUGIN_ROOT}/agent-harness && python3 -m venv .venv && source .venv/bin/activate && pip install -e . && echo "pages-cli installed")
```

If a venv already exists, activate it:

```bash
test -f ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate && source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate
```

## Your task

Parse the argument `$ARGUMENTS` and execute the matching action:

### If no argument or "new":
1. Create a new blank document:
   ```bash
   pages-cli --json document new
   ```
2. Inform the user and wait for instructions on what to do with it.

### If "new <template name>":
1. Create a document from the named template:
   ```bash
   pages-cli --json document new --template "<template name>"
   ```
2. Inform the user and wait for instructions.

### If "open <path>":
1. Open the document:
   ```bash
   pages-cli --json document open "<path>"
   ```
2. Get document info:
   ```bash
   pages-cli --json document info
   ```
3. Show the user a summary (name, pages, word count) and wait for instructions.

### If "info":
```bash
pages-cli --json document info
pages-cli --json text word-count
```
Display the information in a clear format.

### If "close":
```bash
pages-cli document close
```

### If "export pdf <path>" or "export word <path>" etc.:
```bash
pages-cli export <format> "<path>"
```
Confirm the export with file size.

### If "templates":
```bash
pages-cli --json template list
```
Display the templates grouped by category (Reports, Letters, CVs, etc.).

### If "status":
```bash
pages-cli --json session status
```

## After opening/creating

Once a document is open, translate the user's natural language into CLI commands:

- Write text → `pages-cli text add "..."`
- Format text → `pages-cli text set-font --name "..." --size N`
- Add table → `pages-cli table add --rows N --cols N`
- Add image → `pages-cli media add-image "path"`
- Export → `pages-cli export pdf/word/epub "path"`
- Review → `pages-cli --json text get` → Analyze the text and suggest improvements
- Check formatting → `pages-cli --json document info` → Check consistency

ALWAYS use `--json` for query commands to get structured output.
