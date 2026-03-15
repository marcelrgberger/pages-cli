# PAGES.md — Apple Pages CLI Harness SOP

> Software-specific analysis, architecture decisions, command mapping, and
> implementation notes for **pages-cli**.

---

## Phase 1: Codebase Analysis Results

### Application Identity

| Property             | Value                                             |
|----------------------|---------------------------------------------------|
| Application          | Apple Pages                                       |
| Bundle ID            | `com.apple.Pages`                                 |
| Installed Version    | 14.5 (ships with macOS; also on the App Store)    |
| Platform             | macOS only                                        |
| Scripting Enabled    | Yes (`NSAppleScriptEnabled = true`)               |
| Scripting Definition | `/Applications/Pages.app/Contents/Resources/Pages.sdef` |
| Scripting Bridge     | AppleScript via `osascript -e '...'`              |
| Data Format          | `.pages` (proprietary ZIP-based package)          |

### Backend Engine

Apple Pages has **no headless CLI** and **no subprocess rendering mode** (unlike
LibreOffice or Blender). The only programmatic interface is **AppleScript**,
invoked through `osascript`. This is fundamentally different from the typical
HARNESS.md pattern where we generate intermediate files and hand them to a
backend CLI.

**Key architectural difference:** Pages is the _only_ rendering engine. We
cannot generate an intermediate file format and convert it with a separate tool.
Every operation -- document creation, text manipulation, table editing, media
insertion, and export -- must be executed as AppleScript commands sent to the
running Pages application.

**Consequence:** Pages must be running (or will be launched automatically by
`osascript`). There is no true "headless" mode, but the application can run
without any visible windows if no document windows are opened to the foreground.

### Scripting Definition: Pages.sdef

The scripting dictionary is organized into three suites:

#### 1. iWork Text Suite (`iWtx`)

Provides the text model shared across all iWork applications.

| Class              | Code   | Description                        | Key Properties                        |
|--------------------|--------|------------------------------------|---------------------------------------|
| `rich text`        | `rtxt` | Base rich text class               | `color`, `font`, `size`               |
| `character`        | `cha ` | Single character (inherits rich text) | (inherited)                        |
| `paragraph`        | `cpar` | A paragraph (inherits rich text)   | (inherited); contains words, chars    |
| `word`             | `cwor` | A single word (inherits rich text) | (inherited); contains characters      |

**Text hierarchy:** `rich text` > `paragraph` > `word` > `character`

#### 2. iWork Suite (`iwrk`)

Core iWork classes and commands shared across Pages, Numbers, and Keynote.

**Commands:**

| Command           | Code       | Description                                    |
|-------------------|------------|------------------------------------------------|
| `set`             | `coresetd` | Sets the value of a specified object           |
| `delete`          | `coredelo` | Delete an object                               |
| `make`            | `corecrel` | Create a new object                            |
| `clear`           | `NmTbCLR ` | Clear cell contents including formatting       |
| `merge`           | `NMTbMRGE` | Merge a range of cells                         |
| `sort`            | `NmTbSORT` | Sort table rows by column (ascending/descending) |
| `unmerge`         | `NmTbSpUm` | Unmerge all merged cells in a range            |
| `set password`    | `NmTbPset` | Set password on an unencrypted document        |
| `remove password` | `NmTbPdel` | Remove password from a document                |

**Document extensions (iWork Suite):**

| Property             | Code   | Type      | Access | Description                    |
|----------------------|--------|-----------|--------|--------------------------------|
| `selection`          | `sele` | list      | rw     | Currently selected items       |
| `password protected` | `pwpt` | boolean   | r      | Whether document has password  |

**Classes:**

| Class           | Code   | Inherits        | Key Properties                                                       |
|-----------------|--------|-----------------|----------------------------------------------------------------------|
| `iWork container` | `iwkc` | --            | Elements: audio clip, chart, image, iWork item, group, line, movie, shape, table, text item |
| `iWork item`    | `fmti` | --              | `height`, `locked`, `parent`, `position`, `width`                    |
| `audio clip`    | `shau` | iWork item      | `file name`, `clip volume` (0-100), `repetition method`              |
| `shape`         | `sshp` | iWork item      | `background fill type`, `object text`, `reflection showing/value`, `rotation`, `opacity` |
| `chart`         | `shct` | iWork item      | (base properties only)                                               |
| `image`         | `imag` | iWork item      | `description`, `file`, `file name`, `opacity`, `reflection showing/value`, `rotation` |
| `group`         | `igrp` | iWork container | `height`, `parent`, `position`, `width`, `rotation`                  |
| `line`          | `iWln` | iWork item      | `end point`, `start point`, `reflection showing/value`, `rotation`   |
| `movie`         | `shmv` | iWork item      | `file name`, `movie volume` (0-100), `opacity`, `reflection showing/value`, `repetition method`, `rotation` |
| `table`         | `NmTb` | iWork item      | `name`, `cell range`, `selection range`, `row count`, `column count`, `header row count`, `header column count`, `footer row count` |
| `text item`     | `shtx` | iWork item      | `background fill type`, `object text`, `opacity`, `reflection showing/value`, `rotation` |
| `range`         | `NmCR` | --              | `font name`, `font size`, `format`, `alignment`, `name`, `text color`, `text wrap`, `background color`, `vertical alignment` |
| `cell`          | `NmCl` | range           | `column`, `row`, `value`, `formatted value`, `formula`               |
| `row`           | `NMRw` | range           | `address`, `height`                                                  |
| `column`        | `NMCo` | range           | `address`, `width`                                                   |

**Enumerations (iWork Suite):**

| Enumeration             | Values                                                                                              |
|-------------------------|-----------------------------------------------------------------------------------------------------|
| Vertical alignment      | `top`, `center`, `bottom`                                                                           |
| Horizontal alignment    | `auto align`, `center`, `justify`, `left`, `right`                                                  |
| Sort direction           | `ascending`, `descending`                                                                          |
| Cell format              | `automatic`, `checkbox`, `currency`, `date and time`, `fraction`, `number`, `percent`, `pop up menu`, `scientific`, `slider`, `stepper`, `text`, `duration`, `rating`, `numeral system` |
| Item fill options        | `no fill`, `color fill`, `gradient fill`, `advanced gradient fill`, `image fill`, `advanced image fill` |
| Playback repetition      | `none`, `loop`, `loop back and forth`                                                              |

#### 3. Pages Suite (`Pgst`)

Pages-specific classes, commands, and export functionality.

**Document extensions (Pages Suite):**

| Property            | Code   | Type      | Access | Description                              |
|---------------------|--------|-----------|--------|------------------------------------------|
| `id`                | `ID  ` | text      | r      | Document identifier                      |
| `document template` | `Tmpl` | template  | r      | Template assigned to the document        |
| `body text`         | `pTxt` | rich text | rw     | The document body text                   |
| `document body`     | `pDbo` | boolean   | r      | Whether document has body text           |
| `facing pages`      | `pFPa` | boolean   | rw     | Whether document has facing pages        |
| `current page`      | `pCpa` | page      | r      | Current page of the document             |

**Document elements:** audio clip, chart, group, image, iWork item, line, movie,
page, section, shape, table, text item, placeholder text.

**Pages-specific classes:**

| Class              | Code   | Inherits         | Key Properties                                                    |
|--------------------|--------|------------------|-------------------------------------------------------------------|
| `template`         | `tmpl` | --               | `id`, `name`                                                      |
| `section`          | `cSec` | --               | `body text`; elements: audio clip, chart, group, image, iWork item, line, movie, page, shape, table, text item |
| `page`             | `cPag` | iWork container  | `body text`                                                       |
| `placeholder text` | `cpla` | rich text        | `tag`                                                             |

**Window extensions:**

| Property  | Code   | Type    | Access | Description                        |
|-----------|--------|---------|--------|------------------------------------|
| `two up`  | `p2Up` | boolean | rw     | Whether showing pages side-by-side |

**Export command:**

```
export <document> to <file> as <export format> [with properties <export options>]
```

**Export formats:**

| Format             | Code   | Extension  | UTI                                          |
|--------------------|--------|------------|----------------------------------------------|
| PDF                | `Ppdf` | `.pdf`     | `com.apple.iWork.Pages.exportPDF`            |
| Microsoft Word     | `Pwrd` | `.docx`    | `com.apple.iWork.Pages.exportWord`           |
| EPUB               | `Pepu` | `.epub`    | `com.apple.iWork.Pages.exportEPUB`           |
| Formatted text     | `Prtf` | `.rtf`     | `com.apple.iWork.Pages.exportRTF`            |
| Unformatted text   | `Ptxf` | `.txt`     | `com.apple.iWork.Pages.exportText`           |
| Pages 09           | `PPag` | `.pages`   | `com.apple.iWork.Pages.exportPages`          |

**Export options (record type `Pxop`):**

| Option               | Code   | Type          | Description                              |
|----------------------|--------|---------------|------------------------------------------|
| `title`              | `Pxet` | text          | EPUB title                               |
| `author`             | `Pxea` | text          | EPUB author                              |
| `genre`              | `Pxeg` | text          | EPUB genre                               |
| `language`           | `Pxel` | text          | EPUB language (name or ISO code)         |
| `publisher`          | `Pxep` | text          | EPUB publisher                           |
| `cover`              | `Pxec` | boolean       | EPUB first page is cover                 |
| `fixed layout`       | `Pxef` | boolean       | EPUB fixed layout                        |
| `password`           | `PxPW` | text          | Export password                          |
| `password hint`      | `PxPH` | text          | Password hint                            |
| `image quality`      | `PxPI` | image quality | Good (0) / Better (1) / Best (2)        |
| `include comments`   | `PxRC` | boolean       | Include comments in export               |
| `include annotations`| `PxRA` | boolean       | Include smart annotations in export      |

**Saveable file format:**

| Format        | Code   | UTI                                        |
|---------------|--------|--------------------------------------------|
| Pages Format  | `Pgff` | `com.apple.iwork.pages.pages-tef`          |

### Templates

Pages ships with 100+ templates. A partial listing from the live application:

```
Blank, Blank Layout, Blank Landscape, Blank Black, Note Taking,
Simple Report, Essay, Minimalist Report, Contemporary Report,
Photo Report, End Of Term Essay, School Report, Visual Report,
Academic Report, Research Paper, Modern Report, Professional Report,
Project Proposal, Technical Certificate, Classic Certificate,
Kids Certificate, Classic Newsletter, Journal Newsletter,
Simple Newsletter, Serif Newsletter, School Newsletter,
Elegant Brochure, Museum Brochure, Blank Book Portrait, Basic Photo, ...
```

Templates are enumerated via `tell application "Pages" to get name of every template`.

### Data Model

- `.pages` files are ZIP-based packages (directories on disk).
- Internal structure is proprietary (protobuf-based `.iwa` files).
- Unlike ODF or OOXML, the `.pages` format is NOT designed for third-party
  manipulation. Direct editing of `.pages` files is fragile and unsupported.
- All data manipulation MUST go through the AppleScript API.

---

## Phase 2: Architecture Decisions

### Backend Strategy: AppleScript-Only

Unlike most pages-cli harnesses (which generate intermediate files and call a
backend CLI for rendering), the Pages harness operates entirely through
AppleScript. There is no intermediate file format to manipulate.

**Pipeline:**

```
CLI command --> Python --> osascript -e 'tell application "Pages" ...' --> Pages app
```

**Backend module:** `utils/pages_backend.py` wraps all `osascript` calls:

```python
# utils/pages_backend.py
import subprocess
import shutil

def run_applescript(script: str) -> str:
    """Execute AppleScript via osascript and return stdout."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=30,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
    return result.stdout.strip()

def ensure_pages_running():
    """Launch Pages if not already running."""
    run_applescript('tell application "Pages" to activate')
```

**Why not manipulate .pages files directly?**
- The `.pages` format uses protobuf-encoded `.iwa` files inside a ZIP.
- Apple does not document the protobuf schemas.
- Reverse-engineered parsers exist but are fragile across versions.
- AppleScript provides a stable, version-resilient API.
- Pages MUST be running, but `osascript` auto-launches it if needed.

### Interaction Model

Both REPL and subcommand CLI, with REPL as default (per HARNESS.md convention).

### State Model

State is maintained by the running Pages application itself. The CLI tracks:
- **Active document name** (the `--document` flag or REPL context)
- **Session metadata** (JSON in `~/.pages-cli/session.json`)
- Pages' own document state is the source of truth (no shadow copies)

### Output Modes

- **Human-readable:** Tables, colored output via `ReplSkin`
- **Machine-readable:** `--json` flag on every command

---

## Phase 3: CLI Command Mapping

### Command Groups

#### `document` -- Document Management

| CLI Command                              | AppleScript                                                                                     |
|------------------------------------------|-------------------------------------------------------------------------------------------------|
| `document new`                           | `tell application "Pages" to make new document with properties {document template:template "Blank"}` |
| `document new --template "Essay"`        | `tell application "Pages" to make new document with properties {document template:template "Essay"}` |
| `document open <path>`                   | `tell application "Pages" to open POSIX file "<path>"`                                          |
| `document close`                         | `tell application "Pages" to close document "<name>"`                                           |
| `document close --save`                  | `tell application "Pages" to close document "<name>" saving yes`                                |
| `document save`                          | `tell application "Pages" to save document "<name>"`                                            |
| `document save --to <path>`              | `tell application "Pages" to save document "<name>" in POSIX file "<path>"`                     |
| `document info`                          | Queries: `name`, `id`, `document body`, `facing pages`, `password protected`, `current page`, count of pages/sections/tables/images |
| `document list`                          | `tell application "Pages" to get name of every document`                                        |

**`document new` AppleScript details:**
```applescript
-- With built-in template (by name)
tell application "Pages"
    make new document with properties {document template:template "Essay"}
end tell

-- With custom template file
tell application "Pages"
    make new document with data POSIX file "/path/to/template.template"
end tell

-- Blank document (default)
tell application "Pages"
    make new document with properties {document template:template "Blank"}
end tell
```

**`document info` -- compound query:**
```applescript
tell application "Pages"
    tell document "<name>"
        set docName to name
        set docId to id
        set hasBody to document body
        set hasFacing to facing pages
        set isProtected to password protected
        set pageCount to count of pages
        set sectionCount to count of sections
        set tableCount to count of tables
        set imageCount to count of images
        set shapeCount to count of shapes
    end tell
end tell
```

#### `text` -- Text Operations

| CLI Command                                        | AppleScript                                                                                                 |
|----------------------------------------------------|-------------------------------------------------------------------------------------------------------------|
| `text get`                                         | `tell application "Pages" to get body text of document "<name>"`                                            |
| `text set "<content>"`                             | `tell application "Pages" to set body text of document "<name>" to "<content>"`                              |
| `text add "<content>"`                             | `tell application "Pages" to tell document "<name>" to set body text to (body text as text) & "<content>"`  |
| `text set-font --name "Helvetica" --size 14`       | See below                                                                                                   |
| `text set-color --r 255 --g 0 --b 0`              | See below                                                                                                   |
| `text word-count`                                  | `tell application "Pages" to count words of body text of document "<name>"`                                 |

**`text set-font` -- setting font on a paragraph range:**
```applescript
tell application "Pages"
    tell document "<name>"
        set font of paragraph <n> of body text to "Helvetica"
        set size of paragraph <n> of body text to 14
    end tell
end tell
```

**`text set-color` -- RGB color (0-65535 scale):**
```applescript
tell application "Pages"
    tell document "<name>"
        -- Pages uses {R, G, B} with 0-65535 range
        -- CLI accepts 0-255 and multiplies by 257
        set color of paragraph <n> of body text to {65535, 0, 0}
    end tell
end tell
```

**Note on color values:** Pages uses a 16-bit color scale (0-65535 per channel).
The CLI accepts standard 0-255 RGB values and converts: `pages_value = cli_value * 257`.

#### `table` -- Table Operations

| CLI Command                                        | AppleScript                                                                                                 |
|----------------------------------------------------|-------------------------------------------------------------------------------------------------------------|
| `table add --rows 5 --cols 3`                      | `tell application "Pages" to tell document "<name>" to make new table with properties {row count:5, column count:3}` |
| `table add --rows 5 --cols 3 --name "Data"`        | `... with properties {name:"Data", row count:5, column count:3}`                                            |
| `table list`                                       | `tell application "Pages" to get name of every table of document "<name>"`                                  |
| `table set-cell --table "Table 1" --row 1 --col A --value "Hello"` | See below                                                                                 |
| `table get-cell --table "Table 1" --row 1 --col A` | See below                                                                                                  |
| `table merge --table "Table 1" --range "A1:B2"`   | See below                                                                                                   |
| `table unmerge --table "Table 1" --range "A1:B2"` | See below                                                                                                   |
| `table sort --table "Table 1" --by A --direction ascending` | See below                                                                                        |

**`table set-cell`:**
```applescript
tell application "Pages"
    tell table "Table 1" of document "<name>"
        set value of cell "A1" to "Hello"
    end tell
end tell
```

**`table get-cell`:**
```applescript
tell application "Pages"
    tell table "Table 1" of document "<name>"
        get value of cell "A1"
    end tell
end tell
```

**`table merge`:**
```applescript
tell application "Pages"
    tell table "Table 1" of document "<name>"
        merge range "A1:B2"
    end tell
end tell
```

**`table sort`:**
```applescript
tell application "Pages"
    tell document "<name>"
        sort table "Table 1" by column "A" direction ascending
    end tell
end tell
```

**Table cell addressing:** Cells are addressed as `"A1"`, `"B3"`, etc. (column
letter + row number). The CLI can also accept `--row 1 --col 1` (both numeric)
and convert to the letter-number format internally.

**Cell range formatting properties (via `range`):**

```applescript
tell application "Pages"
    tell table "Table 1" of document "<name>"
        tell range "A1:C3"
            set font name to "Helvetica Neue"
            set font size to 12
            set text color to {0, 0, 0}
            set background color to {65535, 65535, 0}
            set alignment to center
            set vertical alignment to center
            set text wrap to true
            set format to number
        end tell
    end tell
end tell
```

#### `media` -- Media Operations

| CLI Command                                                    | AppleScript                                                                                         |
|----------------------------------------------------------------|-----------------------------------------------------------------------------------------------------|
| `media add-image --file <path>`                                | `tell application "Pages" to tell document "<name>" to make new image with properties {file name:POSIX file "<path>"}` |
| `media add-image --file <path> --width 400 --height 300`      | `... with properties {file name:POSIX file "<path>", width:400, height:300}`                        |
| `media add-shape`                                              | `tell application "Pages" to tell document "<name>" to make new shape`                              |
| `media add-shape --text "Label" --x 100 --y 200`              | `... with properties {position:{100, 200}, object text:"Label"}`                                   |
| `media list`                                                   | Queries count/properties of images, shapes, text items, lines, groups, audio clips, movies          |
| `media set-opacity --type image --index 1 --value 50`         | `set opacity of image 1 of document "<name>" to 50`                                                |
| `media set-rotation --type shape --index 1 --value 90`        | `set rotation of shape 1 of document "<name>" to 90`                                               |

**`media list` -- compound query:**
```applescript
tell application "Pages"
    tell document "<name>"
        set imgCount to count of images
        set shpCount to count of shapes
        set tblCount to count of tables
        set txtCount to count of text items
        set lnCount to count of lines
        set grpCount to count of groups
        set audCount to count of audio clips
        set movCount to count of movies
    end tell
end tell
```

#### `export` -- Export Operations

| CLI Command                                                          | AppleScript                                                                                    |
|----------------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| `export pdf <output_path>`                                           | `tell application "Pages" to export document "<name>" to POSIX file "<path>" as PDF`           |
| `export word <output_path>`                                          | `... as Microsoft Word`                                                                        |
| `export epub <output_path>`                                          | `... as EPUB`                                                                                  |
| `export epub <path> --title "T" --author "A" --genre "G"`           | `... as EPUB with properties {title:"T", author:"A", genre:"G"}`                               |
| `export text <output_path>`                                          | `... as unformatted text`                                                                      |
| `export rtf <output_path>`                                           | `... as formatted text`                                                                        |
| `export pdf <path> --password "secret" --hint "hint"`                | `... as PDF with properties {password:"secret", password hint:"hint"}`                         |
| `export pdf <path> --image-quality best --include-comments`          | `... as PDF with properties {image quality:Best, include comments:true}`                       |
| `export formats`                                                     | Lists all supported export formats (static data)                                               |

**Full export AppleScript with all options:**
```applescript
tell application "Pages"
    export document "<name>" to POSIX file "<path>" as PDF with properties {
        image quality:Best,
        include comments:true,
        include annotations:true,
        password:"secret",
        password hint:"remember"
    }
end tell
```

**EPUB export with metadata:**
```applescript
tell application "Pages"
    export document "<name>" to POSIX file "<path>" as EPUB with properties {
        title:"My Book",
        author:"Marcel R. G. Berger",
        genre:"Non-Fiction",
        language:"en",
        publisher:"Self",
        cover:true,
        fixed layout:false
    }
end tell
```

#### `template` -- Template Operations

| CLI Command         | AppleScript                                                         |
|---------------------|---------------------------------------------------------------------|
| `template list`     | `tell application "Pages" to get name of every template`            |
| `template info <n>` | `tell application "Pages" to get {name, id} of template "<name>"`   |

#### `password` -- Password Operations

| CLI Command                                       | AppleScript                                                                                        |
|---------------------------------------------------|----------------------------------------------------------------------------------------------------|
| `password set --password "secret"`                | `tell application "Pages" to tell document "<name>" to set password "secret"`                      |
| `password set --password "secret" --hint "memo"`  | `tell application "Pages" to tell document "<name>" to set password "secret" hint "memo"`          |
| `password set --password "secret" --keychain`     | `... saving in keychain true`                                                                      |
| `password remove --password "secret"`             | `tell application "Pages" to tell document "<name>" to remove password "secret"`                   |
| `password status`                                 | `tell application "Pages" to get password protected of document "<name>"`                          |

#### `page` -- Page Inspection

| CLI Command                    | AppleScript                                                                    |
|--------------------------------|--------------------------------------------------------------------------------|
| `page list`                    | `tell application "Pages" to get count of pages of document "<name>"`          |
| `page info --page 1`          | Queries body text and element counts of page 1                                 |
| `page current`                | `tell application "Pages" to get current page of document "<name>"`            |

#### `section` -- Section Operations

| CLI Command                      | AppleScript                                                                    |
|----------------------------------|--------------------------------------------------------------------------------|
| `section list`                   | `tell application "Pages" to get count of sections of document "<name>"`       |
| `section info --section 1`      | Queries body text and element counts of section 1                              |

#### `selection` -- Selection Operations

| CLI Command                       | AppleScript                                                                    |
|-----------------------------------|--------------------------------------------------------------------------------|
| `selection get`                   | `tell application "Pages" to get selection of document "<name>"`               |
| `selection set --items <spec>`    | `tell application "Pages" to set selection of document "<name>" to {<spec>}`   |

#### `session` -- Session Management

| CLI Command     | Description                                  |
|-----------------|----------------------------------------------|
| `session status`| Show active document, Pages running state    |

---

## Phase 3 Addendum: AppleScript Patterns & Pitfalls

### String Escaping

AppleScript strings use `\"` for literal quotes. When building AppleScript from
Python, all user-supplied text must be sanitized:

```python
def escape_applescript_string(s: str) -> str:
    """Escape a string for embedding in AppleScript."""
    return s.replace("\\", "\\\\").replace('"', '\\"')
```

### Multi-line AppleScript

For complex scripts, use `osascript` with multiple `-e` flags or pipe a script
file:

```python
def run_applescript_lines(lines: list[str]) -> str:
    cmd = ["osascript"]
    for line in lines:
        cmd.extend(["-e", line])
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    ...
```

### Error Handling

Common AppleScript errors from Pages:

| Error                                         | Cause                                    | CLI Response                          |
|-----------------------------------------------|------------------------------------------|---------------------------------------|
| `Can't get document "X"`                      | Document not open                        | "Document 'X' is not open."          |
| `The document "X" is not open`                | Document was closed                      | "Document 'X' was closed."           |
| `Incorrect password`                          | Wrong password for remove/open           | "Incorrect password."                |
| `Can't make table`                            | Invalid context (e.g., page layout doc)  | "Cannot add table in this context."  |
| `Can't get template "X"`                      | Template does not exist                  | "Template 'X' not found."           |
| `Application isn't running`                   | Pages not installed or crashed           | "Pages is not available."            |
| `execution error: The document has changes`   | Close without save                       | Prompt to save or discard            |

### Document Identification

Documents can be referenced by:
- **Name:** `document "Untitled"` (the window title, without `.pages` extension)
- **Index:** `document 1` (front-most document)
- **All:** `every document`

The CLI uses `--document <name>` as the primary targeting mechanism. If omitted,
the front-most document (index 1) is used.

### Timeout Considerations

- Most AppleScript operations complete in < 1 second.
- Large document exports (many pages, high-res images) may take 5-15 seconds.
- Template listing enumerates all installed templates and takes 1-3 seconds.
- The CLI sets a 30-second default timeout, configurable per-command.

### Pages Launch Behavior

- `osascript -e 'tell application "Pages" to ...'` auto-launches Pages if not
  running. The first command after launch takes 2-5 seconds longer.
- The CLI's `session status` command checks if Pages is running via:
  ```applescript
  tell application "System Events" to (name of processes) contains "Pages"
  ```
- To prevent visible window flashing, use `activate` only when needed; otherwise
  operations run against the application without bringing it to the foreground.

---

## Implementation Notes

### Directory Structure

```
pages/
└── agent-harness/
    ├── PAGES.md               # This file
    ├── setup.py               # PyPI package configuration
    └── pages_cli/          # Namespace package (NO __init__.py)
        └── pages/             # Sub-package
            ├── __init__.py
            ├── __main__.py    # python3 -m pages_cli.pages
            ├── README.md      # How to run
            ├── pages_cli.py   # Main CLI entry point (Click + REPL)
            ├── core/
            │   ├── __init__.py
            │   ├── document.py    # Document CRUD: new, open, close, save, info, list
            │   ├── text.py        # Text operations: get, set, add, font, color, word-count
            │   ├── table.py       # Table operations: add, set-cell, get-cell, merge, sort
            │   ├── media.py       # Media operations: add-image, add-shape, list
            │   ├── export.py      # Export pipeline: pdf, word, epub, text, rtf
            │   ├── template.py    # Template listing and info
            │   ├── password.py    # Password set/remove/status
            │   └── session.py     # Session state, Pages status
            ├── utils/
            │   ├── __init__.py
            │   ├── pages_backend.py   # osascript wrapper, AppleScript builder
            │   └── repl_skin.py       # Unified REPL skin (copy from plugin)
            └── tests/
                ├── TEST.md            # Test plan and results
                ├── test_core.py       # Unit tests
                └── test_full_e2e.py   # E2E tests (real Pages interaction)
```

### Core Module Design

#### `utils/pages_backend.py`

The single most important module. All AppleScript execution flows through it.

**Key functions:**

```python
def find_pages() -> str:
    """Verify Pages is installed. Raises RuntimeError if not."""
    app_path = "/Applications/Pages.app"
    if not os.path.exists(app_path):
        raise RuntimeError(
            "Apple Pages is not installed.\n"
            "Install from the Mac App Store: https://apps.apple.com/app/pages/id409201541"
        )
    return app_path

def run_applescript(script: str, timeout: int = 30) -> str:
    """Execute an AppleScript string via osascript."""

def is_pages_running() -> bool:
    """Check if Pages process is active."""

def ensure_pages_running() -> None:
    """Launch Pages if not already running."""

def escape_string(s: str) -> str:
    """Escape a Python string for AppleScript embedding."""

def build_tell_document(doc_name: str, *statements: str) -> str:
    """Build a 'tell document' block."""

def build_tell_table(doc_name: str, table_name: str, *statements: str) -> str:
    """Build a nested tell for table operations."""
```

#### `core/document.py`

```python
def new_document(template: str = "Blank") -> dict:
    """Create a new Pages document from a template."""

def open_document(path: str) -> dict:
    """Open an existing .pages file."""

def close_document(name: str, save: bool = False) -> dict:
    """Close a document, optionally saving."""

def save_document(name: str, path: str | None = None) -> dict:
    """Save a document, optionally to a new location."""

def get_document_info(name: str) -> dict:
    """Return comprehensive document metadata."""

def list_documents() -> list[dict]:
    """List all open documents."""
```

#### `core/export.py`

```python
EXPORT_FORMATS = {
    "pdf": "PDF",
    "word": "Microsoft Word",
    "epub": "EPUB",
    "rtf": "formatted text",
    "text": "unformatted text",
    "pages09": "Pages 09",
}

def export_document(
    name: str,
    output_path: str,
    format: str,
    *,
    title: str | None = None,
    author: str | None = None,
    genre: str | None = None,
    language: str | None = None,
    publisher: str | None = None,
    cover: bool | None = None,
    fixed_layout: bool | None = None,
    password: str | None = None,
    password_hint: str | None = None,
    image_quality: str | None = None,  # "good", "better", "best"
    include_comments: bool | None = None,
    include_annotations: bool | None = None,
) -> dict:
    """Export a document to the specified format with optional properties."""
```

#### `core/table.py`

```python
def column_number_to_letter(n: int) -> str:
    """Convert 1-based column number to letter (1='A', 27='AA')."""

def cell_ref(row: int, col: int) -> str:
    """Build a cell reference like 'A1' from numeric coordinates."""
```

### REPL Behavior

The REPL shows the active document name in the prompt:

```
pages [Untitled] > text add "Hello, World!"
pages [Untitled] > table add --rows 3 --cols 4
pages [Untitled] > export pdf ~/Desktop/output.pdf
pages [Untitled*] > document save
pages [Untitled] >
```

The asterisk (`*`) indicates unsaved changes (detected by querying Pages).

### Differences from Generic HARNESS.md Pattern

| Aspect                  | Generic HARNESS.md                                  | Pages Harness                                    |
|-------------------------|-----------------------------------------------------|--------------------------------------------------|
| Backend                 | CLI subprocess (`libreoffice`, `blender`, `melt`)   | AppleScript via `osascript`                      |
| Intermediate format     | ODF, MLT XML, SVG, .blend-cli.json                  | None -- all operations are live                  |
| Headless mode           | Fully headless (`--headless`, `--background`)        | Pages must be running (GUI app)                  |
| Rendering               | Backend CLI renders from intermediate file           | Pages renders directly (export command)           |
| Platform                | Cross-platform (Linux, macOS, Windows)               | macOS only                                       |
| Installation            | System package manager (`apt`, `brew`)               | Mac App Store (pre-installed on most Macs)       |
| State management        | CLI manipulates project files, state in JSON         | State lives in running Pages app                 |
| File format manipulation| Direct XML/ZIP editing of project files              | Not possible -- use AppleScript API only         |

### E2E Test Strategy

Because Pages must be running for tests, the E2E test suite:

1. **Launches Pages** if not already running (via `osascript`)
2. Creates real documents with templates
3. Adds text, tables, images, shapes
4. Exports to all supported formats (PDF, DOCX, EPUB, RTF, TXT)
5. **Verifies output files:**
   - PDF: magic bytes `%PDF-`, size > 0
   - DOCX: valid ZIP with `[Content_Types].xml` (OOXML)
   - EPUB: valid ZIP with `META-INF/container.xml`
   - RTF: starts with `{\rtf`
   - TXT: contains expected text content
6. Closes test documents without saving
7. **Prints artifact paths** for manual inspection

**No graceful degradation:** If Pages is not installed, tests fail with a clear
error message. Pages is a hard dependency.

**Subprocess tests** use `_resolve_cli("pages-cli")` per HARNESS.md
convention. Full workflow tests create documents, add content, export, and
verify output -- all via the installed CLI command.

### Security Considerations

- **Password handling:** The `password set` and `export --password` commands
  pass passwords through AppleScript strings. The CLI must NOT log passwords.
  Use `--password` flag (not positional arguments) so passwords are less likely
  to appear in shell history.
- **File access:** Pages can read/write any file the user has access to.
  The CLI does not add or remove restrictions.
- **AppleScript injection:** All user-supplied strings MUST be escaped via
  `escape_applescript_string()` before embedding in AppleScript code to prevent
  injection attacks.

### Performance Notes

- **First command latency:** If Pages is not running, the first `osascript` call
  launches it (2-5 seconds). Subsequent calls are fast (< 500ms).
- **Template enumeration:** Listing all templates takes 1-3 seconds.
- **Large document export:** Exporting a 100+ page document with images can take
  10-30 seconds. The CLI should provide progress feedback.
- **Batch operations:** Multiple cell writes to a large table are faster when
  batched into a single `osascript` call with multiple `set value` statements
  inside one `tell` block, rather than one `osascript` call per cell.

---

## Quick Reference: Full AppleScript Examples

### Create document, add content, export as PDF
```applescript
tell application "Pages"
    -- Create from template
    set newDoc to make new document with properties {document template:template "Blank"}
    set docName to name of newDoc

    -- Add body text
    set body text of document docName to "Quarterly Report - Q4 2025"

    -- Add a table
    tell document docName
        set newTable to make new table with properties {row count:4, column count:3, name:"Sales Data"}
        tell newTable
            set value of cell "A1" to "Region"
            set value of cell "B1" to "Revenue"
            set value of cell "C1" to "Growth"
            set value of cell "A2" to "EMEA"
            set value of cell "B2" to 1250000
            set value of cell "C2" to 0.12
        end tell
    end tell

    -- Export as PDF
    export document docName to POSIX file "/Users/marcelrgberger/Desktop/report.pdf" as PDF with properties {image quality:Best, include comments:true}
end tell
```

### Add image and shape
```applescript
tell application "Pages"
    tell document "My Document"
        make new image with properties {file name:POSIX file "/path/to/photo.jpg", width:400, height:300}
        make new shape with properties {position:{100, 500}, width:200, height:100, object text:"Caption"}
    end tell
end tell
```

### Batch cell formatting
```applescript
tell application "Pages"
    tell table "Sales Data" of document "Report"
        tell range "A1:C1"
            set font name to "Helvetica Neue Bold"
            set font size to 14
            set background color to {0, 0, 40000}
            set text color to {65535, 65535, 65535}
            set alignment to center
        end tell
    end tell
end tell
```

### EPUB export with full metadata
```applescript
tell application "Pages"
    export document "Novel" to POSIX file "/Users/marcelrgberger/Desktop/novel.epub" as EPUB with properties {title:"The Great Adventure", author:"Marcel R. G. Berger", genre:"Fiction", language:"de", publisher:"Self-Published", cover:true, fixed layout:false}
end tell
```
