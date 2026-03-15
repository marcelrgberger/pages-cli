"""cli-anything-pages — Click-based CLI for Apple Pages automation.

Main entry point for the CLI harness. Supports both one-shot commands
and an interactive REPL mode (default when no subcommand is given).
"""

import json
import sys
from pathlib import Path

import click

from cli_anything.pages import __version__
from cli_anything.pages.utils.repl_skin import ReplSkin
from cli_anything.pages.utils.pages_backend import (
    _run_applescript,
    _run_jxa,
    ensure_pages_running,
    is_pages_running,
    find_pages,
    quit_pages,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class SessionState:
    """Holds runtime state shared across CLI invocations."""

    def __init__(self):
        self.current_document: str | None = None
        self.json_mode: bool = False

    @property
    def document_name(self) -> str:
        """Short display name of the current document."""
        if self.current_document:
            return Path(self.current_document).stem
        return ""


pass_state = click.make_pass_decorator(SessionState, ensure=True)


def _output(ctx: click.Context, data: dict):
    """Emit output: JSON when --json is active, otherwise human-readable."""
    state: SessionState = ctx.ensure_object(SessionState)
    if state.json_mode:
        click.echo(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        if "error" in data:
            click.echo(f"Error: {data['error']}", err=True)
        elif "message" in data:
            click.echo(data["message"])
        elif "rows" in data and "headers" in data:
            skin = ReplSkin("pages", version=__version__)
            skin.table(data["headers"], data["rows"])
        else:
            for key, value in data.items():
                click.echo(f"{key}: {value}")


def _safe_run(func, *args, **kwargs):
    """Run a function and return (result, None) or (None, error_string)."""
    try:
        result = func(*args, **kwargs)
        return result, None
    except Exception as exc:
        return None, str(exc)


def _get_front_document_name() -> str:
    """Get the name of the front document in Pages."""
    try:
        return _run_applescript(
            'tell application "Pages" to get name of front document'
        )
    except RuntimeError:
        return ""


def _get_front_document_path() -> str:
    """Get the file path of the front document in Pages."""
    try:
        result = _run_applescript(
            'tell application "Pages" to get file of front document as text'
        )
        # Convert HFS path to POSIX
        posix = _run_applescript(f'POSIX path of "{result}"')
        return posix
    except RuntimeError:
        return ""


# ---------------------------------------------------------------------------
# Main group
# ---------------------------------------------------------------------------

@click.group(invoke_without_command=True)
@click.option("--json", "json_mode", is_flag=True, default=False,
              help="Output results as JSON.")
@click.option("--document", "document_name", default=None,
              help="Target document name.")
@click.version_option(version=__version__, prog_name="cli-anything-pages")
@click.pass_context
def cli(ctx, json_mode, document_name):
    """cli-anything-pages — Agent-operable CLI for Apple Pages."""
    state = ctx.ensure_object(SessionState)
    state.json_mode = json_mode
    if document_name:
        state.current_document = document_name

    if ctx.invoked_subcommand is None:
        _run_repl(ctx)


# ---------------------------------------------------------------------------
# document group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def document(ctx):
    """Document management commands."""
    pass


@document.command("new")
@click.option("--template", "template_name", default=None,
              help="Template name to use.")
@click.option("--name", "doc_name", default=None,
              help="Name for the new document.")
@click.pass_context
def document_new(ctx, template_name, doc_name):
    """Create a new document."""
    ensure_pages_running()
    if template_name:
        script = (
            f'tell application "Pages"\n'
            f'  set newDoc to make new document with properties '
            f'{{document template:template "{template_name}"}}\n'
            f'end tell'
        )
    else:
        script = (
            'tell application "Pages"\n'
            '  set newDoc to make new document\n'
            'end tell'
        )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to create document: {err}"})
        return

    state = ctx.ensure_object(SessionState)
    name = _get_front_document_name()
    state.current_document = name

    if doc_name:
        # Save with the given name if a name was provided
        save_path = str(Path.home() / "Documents" / f"{doc_name}.pages")
        save_script = (
            f'tell application "Pages"\n'
            f'  save front document in POSIX file "{save_path}"\n'
            f'end tell'
        )
        _, save_err = _safe_run(_run_applescript, save_script)
        if save_err:
            _output(ctx, {"error": f"Document created but save failed: {save_err}"})
            return
        state.current_document = doc_name
        name = doc_name

    _output(ctx, {"message": f"Created new document: {name}", "document": name})


@document.command("open")
@click.argument("file_path")
@click.pass_context
def document_open(ctx, file_path):
    """Open a document from FILE_PATH."""
    ensure_pages_running()
    abs_path = str(Path(file_path).expanduser().resolve())
    if not Path(abs_path).exists():
        _output(ctx, {"error": f"File not found: {abs_path}"})
        return

    script = (
        f'tell application "Pages"\n'
        f'  open POSIX file "{abs_path}"\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to open document: {err}"})
        return

    state = ctx.ensure_object(SessionState)
    name = _get_front_document_name()
    state.current_document = name or Path(abs_path).stem
    _output(ctx, {"message": f"Opened: {state.current_document}",
                  "document": state.current_document, "path": abs_path})


@document.command("close")
@click.option("--no-save", is_flag=True, default=False,
              help="Close without saving.")
@click.pass_context
def document_close(ctx, no_save):
    """Close the current document."""
    saving = "saving no" if no_save else "saving yes"
    script = f'tell application "Pages" to close front document {saving}'
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to close document: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    closed_name = state.current_document or "document"
    state.current_document = None
    _output(ctx, {"message": f"Closed: {closed_name}"})


@document.command("save")
@click.option("--path", "save_path", default=None,
              help="File path to save to.")
@click.pass_context
def document_save(ctx, save_path):
    """Save the current document."""
    if save_path:
        abs_path = str(Path(save_path).expanduser().resolve())
        script = (
            f'tell application "Pages"\n'
            f'  save front document in POSIX file "{abs_path}"\n'
            f'end tell'
        )
    else:
        script = 'tell application "Pages" to save front document'

    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to save document: {err}"})
        return
    name = _get_front_document_name()
    _output(ctx, {"message": f"Saved: {name}", "document": name})


@document.command("info")
@click.pass_context
def document_info(ctx):
    """Show information about the current document."""
    script = (
        'tell application "Pages"\n'
        '  set d to front document\n'
        '  set docName to name of d\n'
        '  set docModified to modified of d\n'
        '  set pCount to count of pages of d\n'
        '  set wCount to count of words of body text of d\n'
        '  set cCount to count of characters of body text of d\n'
        '  return docName & "||" & docModified & "||" & pCount '
        '& "||" & wCount & "||" & cCount\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to get document info: {err}"})
        return

    parts = result.split("||")
    if len(parts) >= 5:
        info = {
            "name": parts[0],
            "modified": parts[1],
            "pages": parts[2],
            "words": parts[3],
            "characters": parts[4],
        }
    else:
        info = {"raw": result}
    _output(ctx, info)


@document.command("list")
@click.pass_context
def document_list(ctx):
    """List all open documents."""
    script = (
        'tell application "Pages"\n'
        '  set docNames to {}\n'
        '  repeat with d in documents\n'
        '    set end of docNames to name of d\n'
        '  end repeat\n'
        '  set AppleScript\'s text item delimiters to "||"\n'
        '  return docNames as text\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to list documents: {err}"})
        return

    docs = [d.strip() for d in result.split("||") if d.strip()] if result else []
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"documents": docs, "count": len(docs)})
    else:
        if not docs:
            click.echo("No documents open.")
        else:
            for i, doc in enumerate(docs, 1):
                click.echo(f"  {i}. {doc}")


# ---------------------------------------------------------------------------
# text group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def text(ctx):
    """Text operations on the current document."""
    pass


@text.command("add")
@click.argument("text_content")
@click.pass_context
def text_add(ctx, text_content):
    """Add text to the end of the document body."""
    escaped = text_content.replace('\\', '\\\\').replace('"', '\\"')
    script = (
        f'tell application "Pages"\n'
        f'  tell front document\n'
        f'    set body text to (body text as text) & "{escaped}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to add text: {err}"})
        return
    _output(ctx, {"message": "Text added.", "length": len(text_content)})


@text.command("set")
@click.argument("text_content")
@click.pass_context
def text_set(ctx, text_content):
    """Set the entire body text of the document."""
    escaped = text_content.replace('\\', '\\\\').replace('"', '\\"')
    script = (
        f'tell application "Pages"\n'
        f'  set body text of front document to "{escaped}"\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to set text: {err}"})
        return
    _output(ctx, {"message": "Body text replaced.", "length": len(text_content)})


@text.command("get")
@click.pass_context
def text_get(ctx):
    """Get the body text of the document."""
    script = (
        'tell application "Pages"\n'
        '  get body text of front document as text\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to get text: {err}"})
        return
    _output(ctx, {"text": result, "length": len(result) if result else 0})


@text.command("set-font")
@click.option("--name", "font_name", default=None, help="Font family name.")
@click.option("--size", "font_size", type=float, default=None,
              help="Font size in points.")
@click.option("--paragraph", "para_index", type=int, default=None,
              help="Paragraph index (1-based). Applies to all if omitted.")
@click.pass_context
def text_set_font(ctx, font_name, font_size, para_index):
    """Set font properties on the document text."""
    if not font_name and font_size is None:
        _output(ctx, {"error": "Provide at least --name or --size."})
        return

    target = "body text of front document"
    if para_index is not None:
        target = f"paragraph {para_index} of body text of front document"

    lines = [f'tell application "Pages"']
    if font_name:
        lines.append(f'  set font of {target} to "{font_name}"')
    if font_size is not None:
        lines.append(f'  set size of {target} to {font_size}')
    lines.append('end tell')
    script = "\n".join(lines)

    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to set font: {err}"})
        return
    parts = []
    if font_name:
        parts.append(f"font={font_name}")
    if font_size is not None:
        parts.append(f"size={font_size}")
    scope = f"paragraph {para_index}" if para_index else "all text"
    _output(ctx, {"message": f"Font updated ({', '.join(parts)}) on {scope}."})


@text.command("set-color")
@click.option("--r", "red", type=int, required=True, help="Red (0-65535).")
@click.option("--g", "green", type=int, required=True, help="Green (0-65535).")
@click.option("--b", "blue", type=int, required=True, help="Blue (0-65535).")
@click.option("--paragraph", "para_index", type=int, default=None,
              help="Paragraph index (1-based).")
@click.pass_context
def text_set_color(ctx, red, green, blue, para_index):
    """Set text color using RGB values (Apple's 0-65535 range)."""
    target = "body text of front document"
    if para_index is not None:
        target = f"paragraph {para_index} of body text of front document"

    script = (
        f'tell application "Pages"\n'
        f'  set color of {target} to {{{red}, {green}, {blue}}}\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to set color: {err}"})
        return
    _output(ctx, {"message": f"Color set to ({red}, {green}, {blue})."})


@text.command("word-count")
@click.pass_context
def text_word_count(ctx):
    """Get the word count of the document."""
    script = (
        'tell application "Pages"\n'
        '  set wc to count of words of body text of front document\n'
        '  set cc to count of characters of body text of front document\n'
        '  set pc to count of paragraphs of body text of front document\n'
        '  return wc & "||" & cc & "||" & pc\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to get word count: {err}"})
        return
    parts = result.split("||")
    if len(parts) >= 3:
        _output(ctx, {"words": int(parts[0]), "characters": int(parts[1]),
                       "paragraphs": int(parts[2])})
    else:
        _output(ctx, {"words": result})


# ---------------------------------------------------------------------------
# table group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def table(ctx):
    """Table operations on the current document."""
    pass


@table.command("add")
@click.option("--rows", "row_count", type=int, default=3,
              help="Number of rows (default 3).")
@click.option("--cols", "col_count", type=int, default=3,
              help="Number of columns (default 3).")
@click.option("--name", "table_name", default=None,
              help="Name for the table.")
@click.pass_context
def table_add(ctx, row_count, col_count, table_name):
    """Add a table to the document."""
    name_prop = f', name:"{table_name}"' if table_name else ""
    script = (
        f'tell application "Pages"\n'
        f'  tell front document\n'
        f'    set newTable to make new table with properties '
        f'{{row count:{row_count}, column count:{col_count}{name_prop}}}\n'
        f'    return name of newTable\n'
        f'  end tell\n'
        f'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to add table: {err}"})
        return
    _output(ctx, {"message": f"Table created: {result}",
                  "table": result, "rows": row_count, "columns": col_count})


@table.command("set-cell")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.argument("value")
@click.pass_context
def table_set_cell(ctx, table_ref, row, col, value):
    """Set a cell value: TABLE ROW COL VALUE."""
    escaped_val = value.replace('\\', '\\\\').replace('"', '\\"')
    script = (
        f'tell application "Pages"\n'
        f'  tell table "{table_ref}" of front document\n'
        f'    set value of cell {col} of row {row} to "{escaped_val}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to set cell: {err}"})
        return
    _output(ctx, {"message": f"Cell ({row},{col}) set to '{value}'.",
                  "table": table_ref, "row": row, "column": col, "value": value})


@table.command("get-cell")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.pass_context
def table_get_cell(ctx, table_ref, row, col):
    """Get a cell value: TABLE ROW COL."""
    script = (
        f'tell application "Pages"\n'
        f'  tell table "{table_ref}" of front document\n'
        f'    get value of cell {col} of row {row}\n'
        f'  end tell\n'
        f'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to get cell: {err}"})
        return
    _output(ctx, {"value": result, "table": table_ref, "row": row, "column": col})


@table.command("list")
@click.pass_context
def table_list(ctx):
    """List all tables in the document."""
    script = (
        'tell application "Pages"\n'
        '  set tableInfo to {}\n'
        '  repeat with t in tables of front document\n'
        '    set end of tableInfo to (name of t) & ":" & '
        '(row count of t) & "x" & (column count of t)\n'
        '  end repeat\n'
        '  set AppleScript\'s text item delimiters to "||"\n'
        '  return tableInfo as text\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to list tables: {err}"})
        return

    tables = []
    if result:
        for entry in result.split("||"):
            entry = entry.strip()
            if ":" in entry:
                name, dims = entry.rsplit(":", 1)
                tables.append({"name": name, "dimensions": dims})
            elif entry:
                tables.append({"name": entry, "dimensions": "?"})

    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"tables": tables, "count": len(tables)})
    else:
        if not tables:
            click.echo("No tables in document.")
        else:
            for t in tables:
                click.echo(f"  {t['name']}  ({t['dimensions']})")


@table.command("merge")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_merge(ctx, table_ref, cell_range):
    """Merge cells in a table: TABLE RANGE (e.g. 'A1:B2')."""
    script = (
        f'tell application "Pages"\n'
        f'  tell table "{table_ref}" of front document\n'
        f'    merge range "{cell_range}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to merge cells: {err}"})
        return
    _output(ctx, {"message": f"Merged cells {cell_range} in {table_ref}."})


@table.command("sort")
@click.argument("table_ref")
@click.option("--column", "sort_col", type=int, required=True,
              help="Column to sort by (1-based).")
@click.option("--descending", is_flag=True, default=False,
              help="Sort in descending order.")
@click.pass_context
def table_sort(ctx, table_ref, sort_col, descending):
    """Sort a table by a column."""
    order = "descending" if descending else "ascending"
    script = (
        f'tell application "Pages"\n'
        f'  tell table "{table_ref}" of front document\n'
        f'    sort by column {sort_col} direction {order}\n'
        f'  end tell\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to sort table: {err}"})
        return
    _output(ctx, {"message": f"Sorted {table_ref} by column {sort_col} ({order})."})


# ---------------------------------------------------------------------------
# media group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def media(ctx):
    """Media operations (images, shapes)."""
    pass


@media.command("add-image")
@click.argument("file_path")
@click.option("--x", "x_pos", type=int, default=100, help="X position.")
@click.option("--y", "y_pos", type=int, default=100, help="Y position.")
@click.option("--width", "width", type=int, default=None,
              help="Width in points.")
@click.option("--height", "height", type=int, default=None,
              help="Height in points.")
@click.pass_context
def media_add_image(ctx, file_path, x_pos, y_pos, width, height):
    """Add an image to the document."""
    abs_path = str(Path(file_path).expanduser().resolve())
    if not Path(abs_path).exists():
        _output(ctx, {"error": f"Image file not found: {abs_path}"})
        return

    props = f'position:{{{x_pos}, {y_pos}}}'
    if width is not None:
        props += f', width:{width}'
    if height is not None:
        props += f', height:{height}'

    script = (
        f'tell application "Pages"\n'
        f'  tell front document\n'
        f'    set img to make new image with properties '
        f'{{file:POSIX file "{abs_path}", {props}}}\n'
        f'    return name of img\n'
        f'  end tell\n'
        f'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to add image: {err}"})
        return
    _output(ctx, {"message": f"Image added: {result or abs_path}",
                  "path": abs_path, "position": {"x": x_pos, "y": y_pos}})


@media.command("add-shape")
@click.option("--type", "shape_type", default="rectangle",
              help="Shape type (rectangle, oval, star, etc).")
@click.option("--x", "x_pos", type=int, default=100, help="X position.")
@click.option("--y", "y_pos", type=int, default=100, help="Y position.")
@click.option("--w", "width", type=int, default=200, help="Width.")
@click.option("--h", "height", type=int, default=100, help="Height.")
@click.option("--text", "shape_text", default=None,
              help="Text inside the shape.")
@click.pass_context
def media_add_shape(ctx, shape_type, x_pos, y_pos, width, height, shape_text):
    """Add a shape to the document."""
    props = (
        f'position:{{{x_pos}, {y_pos}}}, width:{width}, height:{height}'
    )

    lines = [
        'tell application "Pages"',
        '  tell front document',
        f'    set s to make new shape with properties {{{props}}}',
    ]
    if shape_text:
        escaped = shape_text.replace('\\', '\\\\').replace('"', '\\"')
        lines.append(f'    set object text of s to "{escaped}"')
    lines.append('    return name of s')
    lines.append('  end tell')
    lines.append('end tell')
    script = "\n".join(lines)

    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to add shape: {err}"})
        return
    _output(ctx, {"message": f"Shape added: {result or shape_type}",
                  "type": shape_type,
                  "position": {"x": x_pos, "y": y_pos},
                  "size": {"width": width, "height": height}})


@media.command("list")
@click.pass_context
def media_list(ctx):
    """List media items in the document."""
    script = (
        'tell application "Pages"\n'
        '  set mediaInfo to {}\n'
        '  tell front document\n'
        '    repeat with img in images\n'
        '      set end of mediaInfo to "image:" & (name of img)\n'
        '    end repeat\n'
        '    repeat with s in shapes\n'
        '      set end of mediaInfo to "shape:" & (name of s)\n'
        '    end repeat\n'
        '  end tell\n'
        '  set AppleScript\'s text item delimiters to "||"\n'
        '  return mediaInfo as text\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to list media: {err}"})
        return

    items = []
    if result:
        for entry in result.split("||"):
            entry = entry.strip()
            if ":" in entry:
                kind, name = entry.split(":", 1)
                items.append({"type": kind, "name": name})
            elif entry:
                items.append({"type": "unknown", "name": entry})

    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"media": items, "count": len(items)})
    else:
        if not items:
            click.echo("No media items in document.")
        else:
            for it in items:
                click.echo(f"  [{it['type']}] {it['name']}")


# ---------------------------------------------------------------------------
# export group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def export(ctx):
    """Export the current document to various formats."""
    pass


def _export_document(ctx: click.Context, output_path: str, fmt: str,
                     extra_props: str = ""):
    """Generic export helper."""
    abs_path = str(Path(output_path).expanduser().resolve())
    parent = Path(abs_path).parent
    if not parent.exists():
        parent.mkdir(parents=True, exist_ok=True)

    with_clause = ""
    if extra_props:
        with_clause = f' with properties {{{extra_props}}}'

    script = (
        f'tell application "Pages"\n'
        f'  export front document to POSIX file "{abs_path}" '
        f'as {fmt}{with_clause}\n'
        f'end tell'
    )
    _, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}",
                  "path": abs_path, "format": fmt})


@export.command("pdf")
@click.argument("output_path")
@click.option("--password", "password", default=None,
              help="Password-protect the PDF.")
@click.pass_context
def export_pdf(ctx, output_path, password):
    """Export as PDF."""
    extra = ""
    if password:
        extra = f'password:"{password}"'
    _export_document(ctx, output_path, "PDF", extra)


@export.command("word")
@click.argument("output_path")
@click.pass_context
def export_word(ctx, output_path):
    """Export as Microsoft Word (.docx)."""
    _export_document(ctx, output_path, "Microsoft Word")


@export.command("epub")
@click.argument("output_path")
@click.option("--title", "epub_title", default=None,
              help="EPUB title metadata.")
@click.option("--author", "epub_author", default=None,
              help="EPUB author metadata.")
@click.pass_context
def export_epub(ctx, output_path, epub_title, epub_author):
    """Export as EPUB."""
    extra_parts = []
    if epub_title:
        extra_parts.append(f'title:"{epub_title}"')
    if epub_author:
        extra_parts.append(f'author:"{epub_author}"')
    extra = ", ".join(extra_parts)
    _export_document(ctx, output_path, "EPUB", extra)


@export.command("text")
@click.argument("output_path")
@click.pass_context
def export_text(ctx, output_path):
    """Export as plain text."""
    _export_document(ctx, output_path, "unformatted text")


@export.command("rtf")
@click.argument("output_path")
@click.pass_context
def export_rtf(ctx, output_path):
    """Export as RTF (formatted text)."""
    _export_document(ctx, output_path, "formatted text")


@export.command("formats")
@click.pass_context
def export_formats(ctx):
    """List available export formats."""
    formats = [
        {"name": "pdf", "description": "PDF document", "extension": ".pdf"},
        {"name": "word", "description": "Microsoft Word", "extension": ".docx"},
        {"name": "epub", "description": "EPUB e-book", "extension": ".epub"},
        {"name": "text", "description": "Plain text", "extension": ".txt"},
        {"name": "rtf", "description": "Rich Text Format", "extension": ".rtf"},
    ]
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"formats": formats})
    else:
        skin = ReplSkin("pages", version=__version__)
        headers = ["Format", "Description", "Extension"]
        rows = [[f["name"], f["description"], f["extension"]] for f in formats]
        skin.table(headers, rows)


# ---------------------------------------------------------------------------
# template group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def template(ctx):
    """Template management."""
    pass


@template.command("list")
@click.pass_context
def template_list(ctx):
    """List available templates."""
    script = (
        'tell application "Pages"\n'
        '  set tNames to {}\n'
        '  repeat with t in templates\n'
        '    set end of tNames to name of t\n'
        '  end repeat\n'
        '  set AppleScript\'s text item delimiters to "||"\n'
        '  return tNames as text\n'
        'end tell'
    )
    result, err = _safe_run(_run_applescript, script)
    if err:
        _output(ctx, {"error": f"Failed to list templates: {err}"})
        return

    templates = [t.strip() for t in result.split("||") if t.strip()] if result else []
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"templates": templates, "count": len(templates)})
    else:
        if not templates:
            click.echo("No templates found.")
        else:
            for i, t in enumerate(templates, 1):
                click.echo(f"  {i}. {t}")


# ---------------------------------------------------------------------------
# session group
# ---------------------------------------------------------------------------

@cli.group()
@click.pass_context
def session(ctx):
    """Session management."""
    pass


@session.command("status")
@click.pass_context
def session_status(ctx):
    """Show session status."""
    running = is_pages_running()
    state = ctx.ensure_object(SessionState)

    doc_count = 0
    doc_names = []
    if running:
        script = (
            'tell application "Pages"\n'
            '  set docNames to {}\n'
            '  repeat with d in documents\n'
            '    set end of docNames to name of d\n'
            '  end repeat\n'
            '  set AppleScript\'s text item delimiters to "||"\n'
            '  return docNames as text\n'
            'end tell'
        )
        result, _ = _safe_run(_run_applescript, script)
        if result:
            doc_names = [d.strip() for d in result.split("||") if d.strip()]
            doc_count = len(doc_names)

    info = {
        "pages_running": running,
        "documents_open": doc_count,
        "document_names": doc_names,
        "current_document": state.current_document or "(none)",
        "json_mode": state.json_mode,
        "version": __version__,
    }

    if state.json_mode:
        _output(ctx, info)
    else:
        skin = ReplSkin("pages", version=__version__)
        skin.status_block({
            "Pages running": str(running),
            "Documents open": str(doc_count),
            "Current document": state.current_document or "(none)",
            "CLI version": __version__,
        }, title="Session Status")
        if doc_names:
            click.echo()
            for name in doc_names:
                click.echo(f"    - {name}")


@session.command("save")
@click.argument("path")
@click.pass_context
def session_save(ctx, path):
    """Save session state to a JSON file."""
    state = ctx.ensure_object(SessionState)
    abs_path = str(Path(path).expanduser().resolve())

    session_data = {
        "current_document": state.current_document,
        "json_mode": state.json_mode,
        "version": __version__,
    }
    try:
        Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
        with open(abs_path, "w") as f:
            json.dump(session_data, f, indent=2)
        _output(ctx, {"message": f"Session saved to: {abs_path}", "path": abs_path})
    except OSError as e:
        _output(ctx, {"error": f"Failed to save session: {e}"})


@session.command("load")
@click.argument("path")
@click.pass_context
def session_load(ctx, path):
    """Load session state from a JSON file."""
    abs_path = str(Path(path).expanduser().resolve())
    if not Path(abs_path).exists():
        _output(ctx, {"error": f"Session file not found: {abs_path}"})
        return

    try:
        with open(abs_path) as f:
            session_data = json.load(f)
        state = ctx.ensure_object(SessionState)
        if "current_document" in session_data:
            state.current_document = session_data["current_document"]
        _output(ctx, {"message": f"Session loaded from: {abs_path}",
                      "session": session_data})
    except (OSError, json.JSONDecodeError) as e:
        _output(ctx, {"error": f"Failed to load session: {e}"})


# ---------------------------------------------------------------------------
# repl command (explicit entry)
# ---------------------------------------------------------------------------

@cli.command("repl")
@click.pass_context
def repl_command(ctx):
    """Enter interactive REPL mode."""
    _run_repl(ctx)


# ---------------------------------------------------------------------------
# REPL implementation
# ---------------------------------------------------------------------------

_REPL_HELP = {
    "document new":         "Create a new document [--template T] [--name N]",
    "document open":        "Open a document <file_path>",
    "document close":       "Close document [--no-save]",
    "document save":        "Save document [--path PATH]",
    "document info":        "Show document info",
    "document list":        "List open documents",
    "text add":             "Add text <text>",
    "text set":             "Set body text <text>",
    "text get":             "Get body text",
    "text set-font":        "Set font [--name F] [--size S] [--paragraph N]",
    "text set-color":       "Set text color --r R --g G --b B [--paragraph N]",
    "text word-count":      "Get word count",
    "table add":            "Add table [--rows N] [--cols N] [--name NAME]",
    "table set-cell":       "Set cell <table> <row> <col> <value>",
    "table get-cell":       "Get cell <table> <row> <col>",
    "table list":           "List tables",
    "table merge":          "Merge cells <table> <range>",
    "table sort":           "Sort table <table> --column COL [--descending]",
    "media add-image":      "Add image <file> [--x X] [--y Y] [--width W] [--height H]",
    "media add-shape":      "Add shape [--type T] [--x X] [--y Y] [--w W] [--h H] [--text T]",
    "media list":           "List media items",
    "export pdf":           "Export as PDF <output> [--password P]",
    "export word":          "Export as Word <output>",
    "export epub":          "Export as EPUB <output> [--title T] [--author A]",
    "export text":          "Export as plain text <output>",
    "export rtf":           "Export as RTF <output>",
    "export formats":       "List available export formats",
    "template list":        "List available templates",
    "session status":       "Show session status",
    "session save":         "Save session <path>",
    "session load":         "Load session <path>",
    "help":                 "Show this help message",
    "quit / exit":          "Exit the REPL",
}


def _tokenize_input(raw: str) -> list[str]:
    """Split user input respecting quoted strings."""
    import shlex
    try:
        return shlex.split(raw)
    except ValueError:
        return raw.split()


def _run_repl(ctx: click.Context):
    """Run the interactive REPL loop."""
    state = ctx.ensure_object(SessionState)
    skin = ReplSkin("pages", version=__version__)

    skin.print_banner()

    pt_session = skin.create_prompt_session()

    while True:
        # Determine current document name for prompt
        doc_name = state.document_name
        try:
            if not doc_name:
                live_name = _get_front_document_name()
                if live_name:
                    state.current_document = live_name
                    doc_name = live_name
        except Exception:
            pass

        try:
            user_input = skin.get_input(
                pt_session,
                project_name=doc_name,
            )
        except (EOFError, KeyboardInterrupt):
            skin.print_goodbye()
            break

        if not user_input:
            continue

        lowered = user_input.strip().lower()
        if lowered in ("quit", "exit", "q"):
            skin.print_goodbye()
            break

        if lowered in ("help", "?"):
            skin.help(_REPL_HELP)
            continue

        # Parse and dispatch to Click commands
        tokens = _tokenize_input(user_input)
        if not tokens:
            continue

        # Build args list — inject --json if in json mode
        args = list(tokens)
        if state.json_mode and "--json" not in args:
            args = ["--json"] + args

        try:
            # Create a fresh context for each invocation so Click
            # does not think the command was already invoked
            cli.main(
                args=args,
                prog_name="cli-anything-pages",
                standalone_mode=False,
                **{"obj": state},
            )
        except click.exceptions.UsageError as exc:
            skin.error(str(exc))
            skin.hint("Type 'help' to see available commands.")
        except click.exceptions.Abort:
            skin.warning("Command aborted.")
        except SystemExit:
            # Click sometimes raises SystemExit; catch it in REPL mode
            pass
        except Exception as exc:
            skin.error(f"Unexpected error: {exc}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Entry point for setup.py console_scripts."""
    cli(obj=SessionState())


if __name__ == "__main__":
    main()
