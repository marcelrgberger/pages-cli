"""pages-cli -- Click-based CLI for Apple Pages automation.

Main entry point for the CLI harness. Supports both one-shot commands
and an interactive REPL mode (default when no subcommand is given).

This file provides 100% coverage of the Apple Pages AppleScript/sdef API.
"""

import json
import sys
from pathlib import Path

import click

from pages_cli import __version__
from pages_cli.utils.repl_skin import ReplSkin
from pages_cli.utils.pages_backend import (
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


def _get_doc(ctx: click.Context) -> str | None:
    """Return the --document value from session state (may be None)."""
    state = ctx.ensure_object(SessionState)
    return state.current_document


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
              help="Target document name (used by all subcommands).")
@click.version_option(version=__version__, prog_name="pages-cli")
@click.pass_context
def cli(ctx, json_mode, document_name):
    """pages-cli -- Agent-operable CLI for Apple Pages."""
    state = ctx.ensure_object(SessionState)
    state.json_mode = json_mode
    if document_name:
        state.current_document = document_name

    if ctx.invoked_subcommand is None:
        _run_repl(ctx)


# ===========================================================================
# document group
# ===========================================================================

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
    from pages_cli.core.document import create_document
    result, err = _safe_run(create_document, template=template_name or "Blank", name=doc_name)
    if err:
        _output(ctx, {"error": f"Failed to create document: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    state.current_document = result.get("name", "")
    _output(ctx, {"message": f"Created: {result.get('name', '')}", **result})


@document.command("open")
@click.argument("file_path")
@click.pass_context
def document_open(ctx, file_path):
    """Open a document from FILE_PATH."""
    abs_path = str(Path(file_path).expanduser().resolve())
    if not Path(abs_path).exists():
        _output(ctx, {"error": f"File not found: {abs_path}"})
        return
    from pages_cli.core.document import open_document
    result, err = _safe_run(open_document, abs_path)
    if err:
        _output(ctx, {"error": f"Failed to open document: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    state.current_document = result.get("name", Path(abs_path).stem)
    _output(ctx, {"message": f"Opened: {result.get('name', '')}", **result})


@document.command("close")
@click.option("--no-save", is_flag=True, default=False,
              help="Close without saving.")
@click.pass_context
def document_close(ctx, no_save):
    """Close the current document."""
    from pages_cli.core.document import close_document
    doc = _get_doc(ctx)
    _, err = _safe_run(close_document, name=doc, saving=not no_save)
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
    from pages_cli.core.document import save_document
    doc = _get_doc(ctx)
    abs_path = str(Path(save_path).expanduser().resolve()) if save_path else None
    _, err = _safe_run(save_document, name=doc, path=abs_path)
    if err:
        _output(ctx, {"error": f"Failed to save document: {err}"})
        return
    name = _get_front_document_name()
    _output(ctx, {"message": f"Saved: {name}", "document": name})


@document.command("info")
@click.pass_context
def document_info(ctx):
    """Show information about the current document."""
    from pages_cli.core.document import get_document_info
    doc = _get_doc(ctx)
    result, err = _safe_run(get_document_info, name=doc)
    if err:
        _output(ctx, {"error": f"Failed to get document info: {err}"})
        return
    _output(ctx, result)


@document.command("list")
@click.pass_context
def document_list(ctx):
    """List all open documents."""
    from pages_cli.core.document import list_documents
    result, err = _safe_run(list_documents)
    if err:
        _output(ctx, {"error": f"Failed to list documents: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"documents": result, "count": len(result)})
    else:
        if not result:
            click.echo("No documents open.")
        else:
            for i, doc in enumerate(result, 1):
                click.echo(f"  {i}. {doc['name']}")


@document.command("set-password")
@click.argument("password")
@click.option("--hint", default=None, help="Password hint.")
@click.pass_context
def document_set_password(ctx, password, hint):
    """Set a password on the document."""
    from pages_cli.core.document import set_password
    doc = _get_doc(ctx)
    _, err = _safe_run(set_password, password=password, hint=hint, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set password: {err}"})
        return
    _output(ctx, {"message": "Password set."})


@document.command("remove-password")
@click.argument("password")
@click.pass_context
def document_remove_password(ctx, password):
    """Remove the password from the document."""
    from pages_cli.core.document import remove_password
    doc = _get_doc(ctx)
    _, err = _safe_run(remove_password, password=password, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to remove password: {err}"})
        return
    _output(ctx, {"message": "Password removed."})


@document.command("placeholders")
@click.pass_context
def document_placeholders(ctx):
    """List all placeholder texts in the document."""
    from pages_cli.core.document import list_placeholders
    doc = _get_doc(ctx)
    result, err = _safe_run(list_placeholders, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to list placeholders: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"placeholders": result, "count": len(result)})
    else:
        if not result:
            click.echo("No placeholders.")
        else:
            for p in result:
                click.echo(f"  {p['index']}. tag=\"{p['tag']}\"")


@document.command("set-placeholder")
@click.argument("index", type=int)
@click.argument("tag")
@click.pass_context
def document_set_placeholder(ctx, index, tag):
    """Set a placeholder tag by INDEX to TAG value."""
    from pages_cli.core.document import set_placeholder_tag
    doc = _get_doc(ctx)
    _, err = _safe_run(set_placeholder_tag, index=index, tag=tag, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set placeholder: {err}"})
        return
    _output(ctx, {"message": f"Placeholder {index} set to '{tag}'."})


@document.command("get-placeholder")
@click.argument("index", type=int)
@click.pass_context
def document_get_placeholder(ctx, index):
    """Get a placeholder tag by INDEX."""
    from pages_cli.core.document import get_placeholder_tag
    doc = _get_doc(ctx)
    result, err = _safe_run(get_placeholder_tag, index=index, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get placeholder: {err}"})
        return
    _output(ctx, {"index": index, "tag": result})


@document.command("sections")
@click.pass_context
def document_sections(ctx):
    """Show section count of the document."""
    from pages_cli.core.document import get_section_count
    doc = _get_doc(ctx)
    result, err = _safe_run(get_section_count, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get sections: {err}"})
        return
    _output(ctx, {"section_count": result})


@document.command("section-text")
@click.argument("section_index", type=int)
@click.pass_context
def document_section_text(ctx, section_index):
    """Get the body text of a section by SECTION_INDEX."""
    from pages_cli.core.document import get_section_body_text
    doc = _get_doc(ctx)
    result, err = _safe_run(get_section_body_text, section_index=section_index, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get section text: {err}"})
        return
    _output(ctx, {"text": result, "section": section_index})


@document.command("page-text")
@click.argument("page_index", type=int, default=1)
@click.pass_context
def document_page_text(ctx, page_index):
    """Get the body text of a page by PAGE_INDEX."""
    from pages_cli.core.document import get_page_body_text
    doc = _get_doc(ctx)
    result, err = _safe_run(get_page_body_text, page_index=page_index, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get page text: {err}"})
        return
    _output(ctx, {"text": result, "page": page_index})


@document.command("delete")
@click.argument("object_specifier")
@click.pass_context
def document_delete(ctx, object_specifier):
    """Delete an object by AppleScript specifier (e.g. 'table 1 of page 1')."""
    from pages_cli.core.document import delete_object
    doc = _get_doc(ctx)
    _, err = _safe_run(delete_object, object_specifier=object_specifier, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to delete: {err}"})
        return
    _output(ctx, {"message": f"Deleted: {object_specifier}"})


# ===========================================================================
# text group
# ===========================================================================

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
    from pages_cli.core.text import add_text
    doc = _get_doc(ctx)
    _, err = _safe_run(add_text, text=text_content, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add text: {err}"})
        return
    _output(ctx, {"message": "Text added.", "length": len(text_content)})


@text.command("add-paragraph")
@click.argument("text_content")
@click.option("--font", default=None, help="Font name.")
@click.option("--size", type=float, default=None, help="Font size.")
@click.option("--color-r", type=int, default=None, help="Red (0-65535).")
@click.option("--color-g", type=int, default=None, help="Green (0-65535).")
@click.option("--color-b", type=int, default=None, help="Blue (0-65535).")
@click.option("--bold", is_flag=True, default=False)
@click.option("--italic", is_flag=True, default=False)
@click.pass_context
def text_add_paragraph(ctx, text_content, font, size, color_r, color_g, color_b, bold, italic):
    """Add a styled paragraph to the document."""
    from pages_cli.core.text import add_paragraph
    doc = _get_doc(ctx)
    color = None
    if color_r is not None and color_g is not None and color_b is not None:
        color = (color_r, color_g, color_b)
    _, err = _safe_run(
        add_paragraph, text=text_content, font=font, size=size,
        color=color, bold=bold, italic=italic, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to add paragraph: {err}"})
        return
    _output(ctx, {"message": "Paragraph added."})


@text.command("set")
@click.argument("text_content")
@click.pass_context
def text_set(ctx, text_content):
    """Set the entire body text of the document."""
    from pages_cli.core.text import set_body_text
    doc = _get_doc(ctx)
    _, err = _safe_run(set_body_text, text=text_content, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set text: {err}"})
        return
    _output(ctx, {"message": "Body text replaced.", "length": len(text_content)})


@text.command("get")
@click.pass_context
def text_get(ctx):
    """Get the body text of the document."""
    from pages_cli.core.text import get_body_text
    doc = _get_doc(ctx)
    result, err = _safe_run(get_body_text, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get text: {err}"})
        return
    _output(ctx, {"text": result, "length": len(result) if result else 0})


@text.command("set-font")
@click.option("--name", "font_name", default=None, help="Font family name.")
@click.option("--size", "font_size", type=float, default=None, help="Font size.")
@click.option("--paragraph", "para_index", type=int, default=None, help="Paragraph (1-based).")
@click.option("--word", "word_index", type=int, default=None, help="Word (1-based).")
@click.option("--character", "char_index", type=int, default=None, help="Character (1-based).")
@click.pass_context
def text_set_font(ctx, font_name, font_size, para_index, word_index, char_index):
    """Set font properties on the document text at any level."""
    from pages_cli.core.text import set_font, set_font_size
    doc = _get_doc(ctx)
    if not font_name and font_size is None:
        _output(ctx, {"error": "Provide at least --name or --size."})
        return
    if font_name:
        _, err = _safe_run(
            set_font, font_name=font_name,
            paragraph_index=para_index, word_index=word_index,
            character_index=char_index, document=doc,
        )
        if err:
            _output(ctx, {"error": f"Failed to set font: {err}"})
            return
    if font_size is not None:
        _, err = _safe_run(
            set_font_size, size=font_size,
            paragraph_index=para_index, word_index=word_index,
            character_index=char_index, document=doc,
        )
        if err:
            _output(ctx, {"error": f"Failed to set font size: {err}"})
            return
    parts = []
    if font_name:
        parts.append(f"font={font_name}")
    if font_size is not None:
        parts.append(f"size={font_size}")
    scope = "all text"
    if char_index:
        scope = f"character {char_index}"
    elif word_index:
        scope = f"word {word_index}"
    elif para_index:
        scope = f"paragraph {para_index}"
    _output(ctx, {"message": f"Font updated ({', '.join(parts)}) on {scope}."})


@text.command("get-font")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_get_font(ctx, para_index, word_index, char_index):
    """Get font name at any text level."""
    from pages_cli.core.text import get_font
    doc = _get_doc(ctx)
    result, err = _safe_run(
        get_font, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to get font: {err}"})
        return
    _output(ctx, {"font": result})


@text.command("get-font-size")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_get_font_size(ctx, para_index, word_index, char_index):
    """Get font size at any text level."""
    from pages_cli.core.text import get_font_size
    doc = _get_doc(ctx)
    result, err = _safe_run(
        get_font_size, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to get font size: {err}"})
        return
    _output(ctx, {"font_size": result})


@text.command("set-color")
@click.option("--r", "red", type=int, required=True, help="Red (0-65535).")
@click.option("--g", "green", type=int, required=True, help="Green (0-65535).")
@click.option("--b", "blue", type=int, required=True, help="Blue (0-65535).")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_set_color(ctx, red, green, blue, para_index, word_index, char_index):
    """Set text color using RGB values (Apple's 0-65535 range) at any level."""
    from pages_cli.core.text import set_text_color
    doc = _get_doc(ctx)
    _, err = _safe_run(
        set_text_color, r=red, g=green, b=blue,
        paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to set color: {err}"})
        return
    _output(ctx, {"message": f"Color set to ({red}, {green}, {blue})."})


@text.command("get-color")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_get_color(ctx, para_index, word_index, char_index):
    """Get text color at any level."""
    from pages_cli.core.text import get_text_color
    doc = _get_doc(ctx)
    result, err = _safe_run(
        get_text_color, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to get color: {err}"})
        return
    _output(ctx, {"color": result})


@text.command("bold")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_bold(ctx, para_index, word_index, char_index):
    """Set text to bold via font variant."""
    from pages_cli.core.text import set_bold
    doc = _get_doc(ctx)
    _, err = _safe_run(
        set_bold, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to set bold: {err}"})
        return
    _output(ctx, {"message": "Bold applied."})


@text.command("italic")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_italic(ctx, para_index, word_index, char_index):
    """Set text to italic via font variant."""
    from pages_cli.core.text import set_italic
    doc = _get_doc(ctx)
    _, err = _safe_run(
        set_italic, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to set italic: {err}"})
        return
    _output(ctx, {"message": "Italic applied."})


@text.command("bold-italic")
@click.option("--paragraph", "para_index", type=int, default=None)
@click.option("--word", "word_index", type=int, default=None)
@click.option("--character", "char_index", type=int, default=None)
@click.pass_context
def text_bold_italic(ctx, para_index, word_index, char_index):
    """Set text to bold-italic via font variant."""
    from pages_cli.core.text import set_bold_italic
    doc = _get_doc(ctx)
    _, err = _safe_run(
        set_bold_italic, paragraph_index=para_index, word_index=word_index,
        character_index=char_index, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to set bold-italic: {err}"})
        return
    _output(ctx, {"message": "Bold-italic applied."})


@text.command("style-paragraph")
@click.argument("para_index", type=int)
@click.option("--font", default=None)
@click.option("--size", type=float, default=None)
@click.option("--color-r", type=int, default=None)
@click.option("--color-g", type=int, default=None)
@click.option("--color-b", type=int, default=None)
@click.option("--bold", is_flag=True, default=False)
@click.option("--italic", is_flag=True, default=False)
@click.pass_context
def text_style_paragraph(ctx, para_index, font, size, color_r, color_g, color_b, bold, italic):
    """Style paragraph PARA_INDEX with font, size, color, bold, italic."""
    from pages_cli.core.text import set_paragraph_style
    doc = _get_doc(ctx)
    color = None
    if color_r is not None and color_g is not None and color_b is not None:
        color = (color_r, color_g, color_b)
    _, err = _safe_run(
        set_paragraph_style, paragraph_index=para_index, font=font,
        size=size, color=color, bold=bold, italic=italic, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to style paragraph: {err}"})
        return
    _output(ctx, {"message": f"Paragraph {para_index} styled."})


@text.command("style-word")
@click.argument("word_index", type=int)
@click.option("--font", default=None)
@click.option("--size", type=float, default=None)
@click.option("--color-r", type=int, default=None)
@click.option("--color-g", type=int, default=None)
@click.option("--color-b", type=int, default=None)
@click.option("--bold", is_flag=True, default=False)
@click.option("--italic", is_flag=True, default=False)
@click.pass_context
def text_style_word(ctx, word_index, font, size, color_r, color_g, color_b, bold, italic):
    """Style word WORD_INDEX with font, size, color, bold, italic."""
    from pages_cli.core.text import set_word_style
    doc = _get_doc(ctx)
    color = None
    if color_r is not None and color_g is not None and color_b is not None:
        color = (color_r, color_g, color_b)
    _, err = _safe_run(
        set_word_style, word_index=word_index, font=font,
        size=size, color=color, bold=bold, italic=italic, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to style word: {err}"})
        return
    _output(ctx, {"message": f"Word {word_index} styled."})


@text.command("style-character")
@click.argument("char_index", type=int)
@click.option("--font", default=None)
@click.option("--size", type=float, default=None)
@click.option("--color-r", type=int, default=None)
@click.option("--color-g", type=int, default=None)
@click.option("--color-b", type=int, default=None)
@click.option("--bold", is_flag=True, default=False)
@click.option("--italic", is_flag=True, default=False)
@click.pass_context
def text_style_character(ctx, char_index, font, size, color_r, color_g, color_b, bold, italic):
    """Style character CHAR_INDEX with font, size, color, bold, italic."""
    from pages_cli.core.text import set_character_style
    doc = _get_doc(ctx)
    color = None
    if color_r is not None and color_g is not None and color_b is not None:
        color = (color_r, color_g, color_b)
    _, err = _safe_run(
        set_character_style, character_index=char_index, font=font,
        size=size, color=color, bold=bold, italic=italic, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to style character: {err}"})
        return
    _output(ctx, {"message": f"Character {char_index} styled."})


@text.command("word-count")
@click.pass_context
def text_word_count(ctx):
    """Get word, paragraph, and character counts."""
    from pages_cli.core.text import get_counts
    doc = _get_doc(ctx)
    result, err = _safe_run(get_counts, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get counts: {err}"})
        return
    _output(ctx, result)


# ===========================================================================
# table group
# ===========================================================================

@cli.group()
@click.pass_context
def table(ctx):
    """Table operations on the current document."""
    pass


@table.command("add")
@click.option("--rows", "row_count", type=int, default=3, help="Number of rows.")
@click.option("--cols", "col_count", type=int, default=3, help="Number of columns.")
@click.option("--name", "table_name", default=None, help="Name for the table.")
@click.option("--header-rows", type=int, default=1, help="Header rows.")
@click.option("--header-cols", type=int, default=0, help="Header columns.")
@click.option("--footer-rows", type=int, default=0, help="Footer rows.")
@click.pass_context
def table_add(ctx, row_count, col_count, table_name, header_rows, header_cols, footer_rows):
    """Add a table to the document."""
    from pages_cli.core.tables import add_table
    doc = _get_doc(ctx)
    result, err = _safe_run(
        add_table, rows=row_count, cols=col_count, name=table_name,
        header_rows=header_rows, header_columns=header_cols,
        footer_rows=footer_rows, document=doc,
    )
    if err:
        _output(ctx, {"error": f"Failed to add table: {err}"})
        return
    _output(ctx, {"message": f"Table created: {result.get('name', '')}", **result})


@table.command("delete")
@click.option("--name", "table_name", default=None, help="Table name.")
@click.option("--index", "table_index", type=int, default=None, help="Table index.")
@click.pass_context
def table_delete(ctx, table_name, table_index):
    """Delete a table by name or index."""
    from pages_cli.core.tables import delete_table
    doc = _get_doc(ctx)
    _, err = _safe_run(delete_table, table_name=table_name, index=table_index, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to delete table: {err}"})
        return
    _output(ctx, {"message": f"Table deleted."})


@table.command("set-cell")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.argument("value")
@click.pass_context
def table_set_cell(ctx, table_ref, row, col, value):
    """Set a cell value: TABLE ROW COL VALUE (use '=SUM(...)' for formulas)."""
    from pages_cli.core.tables import set_cell
    doc = _get_doc(ctx)
    _, err = _safe_run(set_cell, table_name=table_ref, row=row, col=col, value=value, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set cell: {err}"})
        return
    _output(ctx, {"message": f"Cell ({row},{col}) set.", "value": value})


@table.command("get-cell")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.pass_context
def table_get_cell(ctx, table_ref, row, col):
    """Get a cell value: TABLE ROW COL."""
    from pages_cli.core.tables import get_cell
    doc = _get_doc(ctx)
    result, err = _safe_run(get_cell, table_name=table_ref, row=row, col=col, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get cell: {err}"})
        return
    _output(ctx, {"value": result, "row": row, "column": col})


@table.command("get-formula")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.pass_context
def table_get_formula(ctx, table_ref, row, col):
    """Get the formula of a cell."""
    from pages_cli.core.tables import get_cell_formula
    doc = _get_doc(ctx)
    result, err = _safe_run(get_cell_formula, table_name=table_ref, row=row, col=col, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get formula: {err}"})
        return
    _output(ctx, {"formula": result, "row": row, "column": col})


@table.command("get-formatted-value")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.pass_context
def table_get_formatted_value(ctx, table_ref, row, col):
    """Get the formatted display value of a cell."""
    from pages_cli.core.tables import get_cell_formatted_value
    doc = _get_doc(ctx)
    result, err = _safe_run(get_cell_formatted_value, table_name=table_ref, row=row, col=col, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get formatted value: {err}"})
        return
    _output(ctx, {"formatted_value": result, "row": row, "column": col})


@table.command("list")
@click.pass_context
def table_list(ctx):
    """List all tables in the document."""
    from pages_cli.core.tables import list_tables
    doc = _get_doc(ctx)
    result, err = _safe_run(list_tables, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to list tables: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"tables": result, "count": len(result)})
    else:
        if not result:
            click.echo("No tables in document.")
        else:
            for t in result:
                click.echo(f"  {t['name']}  ({t['rows']}x{t['columns']})")


@table.command("info")
@click.argument("table_ref")
@click.pass_context
def table_info(ctx, table_ref):
    """Get detailed info for table TABLE_REF."""
    from pages_cli.core.tables import get_table_info
    doc = _get_doc(ctx)
    result, err = _safe_run(get_table_info, table_name=table_ref, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get table info: {err}"})
        return
    _output(ctx, result)


@table.command("get-data")
@click.argument("table_ref")
@click.pass_context
def table_get_data(ctx, table_ref):
    """Get all cell values of a table as a 2D array."""
    from pages_cli.core.tables import get_table_data
    doc = _get_doc(ctx)
    result, err = _safe_run(get_table_data, table_name=table_ref, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get table data: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"data": result, "rows": len(result)})
    else:
        for row in result:
            click.echo("\t".join(str(v) for v in row))


@table.command("fill")
@click.argument("table_ref")
@click.argument("json_data")
@click.pass_context
def table_fill(ctx, table_ref, json_data):
    """Fill a table from JSON 2D array: TABLE '[[\"a\",\"b\"],[\"c\",\"d\"]]'."""
    from pages_cli.core.tables import fill_table
    doc = _get_doc(ctx)
    try:
        data = json.loads(json_data)
    except json.JSONDecodeError as e:
        _output(ctx, {"error": f"Invalid JSON: {e}"})
        return
    _, err = _safe_run(fill_table, table_name=table_ref, data=data, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to fill table: {err}"})
        return
    _output(ctx, {"message": f"Table {table_ref} filled."})


@table.command("merge")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_merge(ctx, table_ref, cell_range):
    """Merge cells in a table: TABLE RANGE (e.g. 'A1:B2')."""
    from pages_cli.core.tables import merge_cells
    doc = _get_doc(ctx)
    _, err = _safe_run(merge_cells, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to merge cells: {err}"})
        return
    _output(ctx, {"message": f"Merged {cell_range} in {table_ref}."})


@table.command("unmerge")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_unmerge(ctx, table_ref, cell_range):
    """Unmerge cells in a table: TABLE RANGE."""
    from pages_cli.core.tables import unmerge_cells
    doc = _get_doc(ctx)
    _, err = _safe_run(unmerge_cells, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to unmerge cells: {err}"})
        return
    _output(ctx, {"message": f"Unmerged {cell_range} in {table_ref}."})


@table.command("clear")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_clear(ctx, table_ref, cell_range):
    """Clear contents+formatting of a cell range."""
    from pages_cli.core.tables import clear_range
    doc = _get_doc(ctx)
    _, err = _safe_run(clear_range, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to clear range: {err}"})
        return
    _output(ctx, {"message": f"Cleared {cell_range} in {table_ref}."})


@table.command("sort")
@click.argument("table_ref")
@click.option("--column", "sort_col", type=int, required=True, help="Column (1-based).")
@click.option("--descending", is_flag=True, default=False)
@click.pass_context
def table_sort(ctx, table_ref, sort_col, descending):
    """Sort a table by a column."""
    from pages_cli.core.tables import sort_table
    doc = _get_doc(ctx)
    direction = "descending" if descending else "ascending"
    _, err = _safe_run(sort_table, table_name=table_ref, column=sort_col, direction=direction, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to sort table: {err}"})
        return
    _output(ctx, {"message": f"Sorted {table_ref} by column {sort_col} ({direction})."})


@table.command("set-cell-format")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("col", type=int)
@click.argument("format_type")
@click.pass_context
def table_set_cell_format(ctx, table_ref, row, col, format_type):
    """Set cell format: TABLE ROW COL FORMAT (e.g. currency, percent)."""
    from pages_cli.core.tables import set_cell_format
    doc = _get_doc(ctx)
    _, err = _safe_run(set_cell_format, table_name=table_ref, row=row, col=col, format_type=format_type, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set cell format: {err}"})
        return
    _output(ctx, {"message": f"Cell ({row},{col}) format set to {format_type}."})


@table.command("set-header-rows")
@click.argument("table_ref")
@click.argument("count", type=int)
@click.pass_context
def table_set_header_rows(ctx, table_ref, count):
    """Set header row count for a table."""
    from pages_cli.core.tables import set_header_rows
    doc = _get_doc(ctx)
    _, err = _safe_run(set_header_rows, table_name=table_ref, count=count, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set header rows: {err}"})
        return
    _output(ctx, {"message": f"Header rows set to {count}."})


@table.command("set-header-cols")
@click.argument("table_ref")
@click.argument("count", type=int)
@click.pass_context
def table_set_header_cols(ctx, table_ref, count):
    """Set header column count for a table."""
    from pages_cli.core.tables import set_header_columns
    doc = _get_doc(ctx)
    _, err = _safe_run(set_header_columns, table_name=table_ref, count=count, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set header columns: {err}"})
        return
    _output(ctx, {"message": f"Header columns set to {count}."})


@table.command("set-footer-rows")
@click.argument("table_ref")
@click.argument("count", type=int)
@click.pass_context
def table_set_footer_rows(ctx, table_ref, count):
    """Set footer row count for a table."""
    from pages_cli.core.tables import set_footer_rows
    doc = _get_doc(ctx)
    _, err = _safe_run(set_footer_rows, table_name=table_ref, count=count, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set footer rows: {err}"})
        return
    _output(ctx, {"message": f"Footer rows set to {count}."})


@table.command("set-row-height")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.argument("height", type=float)
@click.pass_context
def table_set_row_height(ctx, table_ref, row, height):
    """Set the height of a row."""
    from pages_cli.core.tables import set_row_height
    doc = _get_doc(ctx)
    _, err = _safe_run(set_row_height, table_name=table_ref, row=row, height=height, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set row height: {err}"})
        return
    _output(ctx, {"message": f"Row {row} height set to {height}pt."})


@table.command("get-row-height")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.pass_context
def table_get_row_height(ctx, table_ref, row):
    """Get the height of a row."""
    from pages_cli.core.tables import get_row_height
    doc = _get_doc(ctx)
    result, err = _safe_run(get_row_height, table_name=table_ref, row=row, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get row height: {err}"})
        return
    _output(ctx, {"row": row, "height": result})


@table.command("set-col-width")
@click.argument("table_ref")
@click.argument("col", type=int)
@click.argument("width", type=float)
@click.pass_context
def table_set_col_width(ctx, table_ref, col, width):
    """Set the width of a column."""
    from pages_cli.core.tables import set_column_width
    doc = _get_doc(ctx)
    _, err = _safe_run(set_column_width, table_name=table_ref, col=col, width=width, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set column width: {err}"})
        return
    _output(ctx, {"message": f"Column {col} width set to {width}pt."})


@table.command("get-col-width")
@click.argument("table_ref")
@click.argument("col", type=int)
@click.pass_context
def table_get_col_width(ctx, table_ref, col):
    """Get the width of a column."""
    from pages_cli.core.tables import get_column_width
    doc = _get_doc(ctx)
    result, err = _safe_run(get_column_width, table_name=table_ref, col=col, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get column width: {err}"})
        return
    _output(ctx, {"column": col, "width": result})


# --- Range property commands ---

@table.command("range-set-font")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("font_name")
@click.pass_context
def table_range_set_font(ctx, table_ref, cell_range, font_name):
    """Set font name on a cell range."""
    from pages_cli.core.tables import set_range_font
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_font, table_name=table_ref, range_str=cell_range, font_name=font_name, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set range font: {err}"})
        return
    _output(ctx, {"message": f"Font set to {font_name} on {cell_range}."})


@table.command("range-get-font")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_font(ctx, table_ref, cell_range):
    """Get font name of a cell range."""
    from pages_cli.core.tables import get_range_font
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_font, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get range font: {err}"})
        return
    _output(ctx, {"font": result, "range": cell_range})


@table.command("range-set-font-size")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("size", type=float)
@click.pass_context
def table_range_set_font_size(ctx, table_ref, cell_range, size):
    """Set font size on a cell range."""
    from pages_cli.core.tables import set_range_font_size
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_font_size, table_name=table_ref, range_str=cell_range, size=size, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set range font size: {err}"})
        return
    _output(ctx, {"message": f"Font size set to {size} on {cell_range}."})


@table.command("range-set-format")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("format_type")
@click.pass_context
def table_range_set_format(ctx, table_ref, cell_range, format_type):
    """Set cell format on a range (e.g. currency, percent, number)."""
    from pages_cli.core.tables import set_range_format
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_format, table_name=table_ref, range_str=cell_range, format_type=format_type, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set range format: {err}"})
        return
    _output(ctx, {"message": f"Format set to {format_type} on {cell_range}."})


@table.command("range-get-format")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_format(ctx, table_ref, cell_range):
    """Get cell format of a range."""
    from pages_cli.core.tables import get_range_format
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_format, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get range format: {err}"})
        return
    _output(ctx, {"format": result, "range": cell_range})


@table.command("range-set-alignment")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("alignment")
@click.pass_context
def table_range_set_alignment(ctx, table_ref, cell_range, alignment):
    """Set horizontal alignment: auto align, center, justify, left, right."""
    from pages_cli.core.tables import set_range_alignment
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_alignment, table_name=table_ref, range_str=cell_range, alignment=alignment, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set alignment: {err}"})
        return
    _output(ctx, {"message": f"Alignment set to {alignment} on {cell_range}."})


@table.command("range-get-alignment")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_alignment(ctx, table_ref, cell_range):
    """Get horizontal alignment of a range."""
    from pages_cli.core.tables import get_range_alignment
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_alignment, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get alignment: {err}"})
        return
    _output(ctx, {"alignment": result, "range": cell_range})


@table.command("range-set-vertical-alignment")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("alignment")
@click.pass_context
def table_range_set_valign(ctx, table_ref, cell_range, alignment):
    """Set vertical alignment: top, center, bottom."""
    from pages_cli.core.tables import set_range_vertical_alignment
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_vertical_alignment, table_name=table_ref, range_str=cell_range, alignment=alignment, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set vertical alignment: {err}"})
        return
    _output(ctx, {"message": f"Vertical alignment set to {alignment} on {cell_range}."})


@table.command("range-set-text-color")
@click.argument("table_ref")
@click.argument("cell_range")
@click.option("--r", "red", type=int, required=True)
@click.option("--g", "green", type=int, required=True)
@click.option("--b", "blue", type=int, required=True)
@click.pass_context
def table_range_set_text_color(ctx, table_ref, cell_range, red, green, blue):
    """Set text color on a range (RGB 0-65535)."""
    from pages_cli.core.tables import set_range_text_color
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_text_color, table_name=table_ref, range_str=cell_range, r=red, g=green, b=blue, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set text color: {err}"})
        return
    _output(ctx, {"message": f"Text color set on {cell_range}."})


@table.command("range-set-bg-color")
@click.argument("table_ref")
@click.argument("cell_range")
@click.option("--r", "red", type=int, required=True)
@click.option("--g", "green", type=int, required=True)
@click.option("--b", "blue", type=int, required=True)
@click.pass_context
def table_range_set_bg_color(ctx, table_ref, cell_range, red, green, blue):
    """Set background color on a range (RGB 0-65535)."""
    from pages_cli.core.tables import set_range_background_color
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_background_color, table_name=table_ref, range_str=cell_range, r=red, g=green, b=blue, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set background color: {err}"})
        return
    _output(ctx, {"message": f"Background color set on {cell_range}."})


@table.command("range-set-text-wrap")
@click.argument("table_ref")
@click.argument("cell_range")
@click.argument("wrap", type=bool)
@click.pass_context
def table_range_set_text_wrap(ctx, table_ref, cell_range, wrap):
    """Set text wrap on a range (true/false)."""
    from pages_cli.core.tables import set_range_text_wrap
    doc = _get_doc(ctx)
    _, err = _safe_run(set_range_text_wrap, table_name=table_ref, range_str=cell_range, wrap=wrap, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set text wrap: {err}"})
        return
    _output(ctx, {"message": f"Text wrap set to {wrap} on {cell_range}."})


@table.command("range-get-font-size")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_font_size(ctx, table_ref, cell_range):
    """Get font size of a cell range."""
    from pages_cli.core.tables import get_range_font_size
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_font_size, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get range font size: {err}"})
        return
    _output(ctx, {"font_size": result, "range": cell_range})


@table.command("range-get-vertical-alignment")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_valign(ctx, table_ref, cell_range):
    """Get vertical alignment of a cell range."""
    from pages_cli.core.tables import get_range_vertical_alignment
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_vertical_alignment, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get vertical alignment: {err}"})
        return
    _output(ctx, {"vertical_alignment": result, "range": cell_range})


@table.command("range-get-text-color")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_text_color(ctx, table_ref, cell_range):
    """Get text color of a cell range."""
    from pages_cli.core.tables import get_range_text_color
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_text_color, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get text color: {err}"})
        return
    _output(ctx, {"text_color": result, "range": cell_range})


@table.command("range-get-bg-color")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_bg_color(ctx, table_ref, cell_range):
    """Get background color of a cell range."""
    from pages_cli.core.tables import get_range_background_color
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_background_color, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get background color: {err}"})
        return
    _output(ctx, {"background_color": result, "range": cell_range})


@table.command("range-get-text-wrap")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_text_wrap(ctx, table_ref, cell_range):
    """Get text wrap setting of a cell range."""
    from pages_cli.core.tables import get_range_text_wrap
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_text_wrap, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get text wrap: {err}"})
        return
    _output(ctx, {"text_wrap": result, "range": cell_range})


@table.command("range-get-name")
@click.argument("table_ref")
@click.argument("cell_range")
@click.pass_context
def table_range_get_name(ctx, table_ref, cell_range):
    """Get the name (coordinates) of a cell range."""
    from pages_cli.core.tables import get_range_name
    doc = _get_doc(ctx)
    result, err = _safe_run(get_range_name, table_name=table_ref, range_str=cell_range, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get range name: {err}"})
        return
    _output(ctx, {"name": result, "range": cell_range})


@table.command("get-row-address")
@click.argument("table_ref")
@click.argument("row", type=int)
@click.pass_context
def table_get_row_address(ctx, table_ref, row):
    """Get the address of a row."""
    from pages_cli.core.tables import get_row_address
    doc = _get_doc(ctx)
    result, err = _safe_run(get_row_address, table_name=table_ref, row=row, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get row address: {err}"})
        return
    _output(ctx, {"row": row, "address": result})


@table.command("get-col-address")
@click.argument("table_ref")
@click.argument("col", type=int)
@click.pass_context
def table_get_col_address(ctx, table_ref, col):
    """Get the address of a column."""
    from pages_cli.core.tables import get_column_address
    doc = _get_doc(ctx)
    result, err = _safe_run(get_column_address, table_name=table_ref, col=col, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get column address: {err}"})
        return
    _output(ctx, {"column": col, "address": result})


# ===========================================================================
# media group
# ===========================================================================

@cli.group()
@click.pass_context
def media(ctx):
    """Media operations (images, shapes, text items, audio, movies, lines, groups)."""
    pass


@media.command("add-image")
@click.argument("file_path")
@click.option("--x", "x_pos", type=int, default=100, help="X position.")
@click.option("--y", "y_pos", type=int, default=100, help="Y position.")
@click.option("--width", type=int, default=None, help="Width in points.")
@click.option("--height", type=int, default=None, help="Height in points.")
@click.option("--description", "alt_text", default=None, help="VoiceOver description.")
@click.pass_context
def media_add_image(ctx, file_path, x_pos, y_pos, width, height, alt_text):
    """Add an image to the document."""
    abs_path = str(Path(file_path).expanduser().resolve())
    if not Path(abs_path).exists():
        _output(ctx, {"error": f"Image file not found: {abs_path}"})
        return
    from pages_cli.core.media import add_image
    doc = _get_doc(ctx)
    result, err = _safe_run(add_image, file_path=abs_path, x=x_pos, y=y_pos, width=width, height=height, description=alt_text, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add image: {err}"})
        return
    _output(ctx, {"message": f"Image added.", **result})


@media.command("add-shape")
@click.option("--type", "shape_type", default="rectangle", help="Shape type.")
@click.option("--x", "x_pos", type=int, default=100, help="X position.")
@click.option("--y", "y_pos", type=int, default=100, help="Y position.")
@click.option("--w", "width", type=int, default=200, help="Width.")
@click.option("--h", "height", type=int, default=100, help="Height.")
@click.option("--text", "shape_text", default=None, help="Text inside shape.")
@click.pass_context
def media_add_shape(ctx, shape_type, x_pos, y_pos, width, height, shape_text):
    """Add a shape to the document."""
    from pages_cli.core.media import add_shape
    doc = _get_doc(ctx)
    result, err = _safe_run(add_shape, shape_type=shape_type, x=x_pos, y=y_pos, width=width, height=height, text=shape_text, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add shape: {err}"})
        return
    _output(ctx, {"message": f"Shape added.", **result})


@media.command("add-text-item")
@click.argument("text_content")
@click.option("--x", "x_pos", type=int, default=100, help="X position.")
@click.option("--y", "y_pos", type=int, default=100, help="Y position.")
@click.option("--w", "width", type=int, default=300, help="Width.")
@click.option("--h", "height", type=int, default=60, help="Height.")
@click.pass_context
def media_add_text_item(ctx, text_content, x_pos, y_pos, width, height):
    """Add a text item (text box) to the document."""
    from pages_cli.core.media import add_text_item
    doc = _get_doc(ctx)
    result, err = _safe_run(add_text_item, text=text_content, x=x_pos, y=y_pos, width=width, height=height, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add text item: {err}"})
        return
    _output(ctx, {"message": "Text item added.", **result})


@media.command("add-audio")
@click.argument("file_path")
@click.option("--x", "x_pos", type=int, default=100)
@click.option("--y", "y_pos", type=int, default=100)
@click.pass_context
def media_add_audio(ctx, file_path, x_pos, y_pos):
    """Add an audio clip to the document."""
    abs_path = str(Path(file_path).expanduser().resolve())
    from pages_cli.core.media import add_audio_clip
    doc = _get_doc(ctx)
    result, err = _safe_run(add_audio_clip, file_path=abs_path, x=x_pos, y=y_pos, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add audio clip: {err}"})
        return
    _output(ctx, {"message": "Audio clip added.", **result})


@media.command("add-movie")
@click.argument("file_path")
@click.option("--x", "x_pos", type=int, default=100)
@click.option("--y", "y_pos", type=int, default=100)
@click.option("--width", type=int, default=None)
@click.option("--height", type=int, default=None)
@click.pass_context
def media_add_movie(ctx, file_path, x_pos, y_pos, width, height):
    """Add a movie to the document."""
    abs_path = str(Path(file_path).expanduser().resolve())
    from pages_cli.core.media import add_movie
    doc = _get_doc(ctx)
    result, err = _safe_run(add_movie, file_path=abs_path, x=x_pos, y=y_pos, width=width, height=height, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add movie: {err}"})
        return
    _output(ctx, {"message": "Movie added.", **result})


@media.command("add-line")
@click.option("--start-x", type=int, required=True)
@click.option("--start-y", type=int, required=True)
@click.option("--end-x", type=int, required=True)
@click.option("--end-y", type=int, required=True)
@click.pass_context
def media_add_line(ctx, start_x, start_y, end_x, end_y):
    """Add a line to the document (may fail in some Pages versions)."""
    from pages_cli.core.media import add_line
    doc = _get_doc(ctx)
    result, err = _safe_run(add_line, start_x=start_x, start_y=start_y, end_x=end_x, end_y=end_y, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to add line: {err}"})
        return
    _output(ctx, {"message": "Line added.", **result})


@media.command("delete")
@click.argument("item_type")
@click.option("--index", type=int, default=None, help="Item index (1-based).")
@click.option("--name", "item_name", default=None, help="Item name.")
@click.pass_context
def media_delete(ctx, item_type, index, item_name):
    """Delete an item: ITEM_TYPE (image, shape, text item, ...) --index N or --name N."""
    from pages_cli.core.media import delete_item
    doc = _get_doc(ctx)
    _, err = _safe_run(delete_item, item_type=item_type, index=index, name=item_name, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to delete item: {err}"})
        return
    _output(ctx, {"message": f"Deleted {item_type}."})


@media.command("set-property")
@click.argument("item_type")
@click.argument("index", type=int)
@click.argument("prop_name")
@click.argument("prop_value")
@click.pass_context
def media_set_property(ctx, item_type, index, prop_name, prop_value):
    """Set a property on any media item: TYPE INDEX PROP VALUE."""
    from pages_cli.core.media import set_item_property
    doc = _get_doc(ctx)
    _, err = _safe_run(set_item_property, item_type=item_type, index=index, prop_name=prop_name, prop_value=prop_value, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set property: {err}"})
        return
    _output(ctx, {"message": f"Set {prop_name}={prop_value} on {item_type} {index}."})


@media.command("get-property")
@click.argument("item_type")
@click.argument("index", type=int)
@click.argument("prop_name")
@click.pass_context
def media_get_property(ctx, item_type, index, prop_name):
    """Get a property of any media item: TYPE INDEX PROP."""
    from pages_cli.core.media import get_item_property
    doc = _get_doc(ctx)
    result, err = _safe_run(get_item_property, item_type=item_type, index=index, prop_name=prop_name, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get property: {err}"})
        return
    _output(ctx, {"property": prop_name, "value": result})


@media.command("properties")
@click.argument("item_type")
@click.argument("index", type=int)
@click.pass_context
def media_properties(ctx, item_type, index):
    """Get all properties of a media item: TYPE INDEX."""
    from pages_cli.core.media import get_item_properties
    doc = _get_doc(ctx)
    result, err = _safe_run(get_item_properties, item_type=item_type, index=index, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to get properties: {err}"})
        return
    _output(ctx, result)


@media.command("set-rotation")
@click.argument("item_type")
@click.argument("index", type=int)
@click.argument("degrees", type=float)
@click.pass_context
def media_set_rotation(ctx, item_type, index, degrees):
    """Set rotation of an item (0-359)."""
    from pages_cli.core.media import set_rotation
    doc = _get_doc(ctx)
    _, err = _safe_run(set_rotation, item_type=item_type, index=index, degrees=degrees, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set rotation: {err}"})
        return
    _output(ctx, {"message": f"Rotation set to {degrees} on {item_type} {index}."})


@media.command("set-opacity")
@click.argument("item_type")
@click.argument("index", type=int)
@click.argument("opacity", type=int)
@click.pass_context
def media_set_opacity(ctx, item_type, index, opacity):
    """Set opacity of an item (0-100)."""
    from pages_cli.core.media import set_opacity
    doc = _get_doc(ctx)
    _, err = _safe_run(set_opacity, item_type=item_type, index=index, opacity=opacity, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set opacity: {err}"})
        return
    _output(ctx, {"message": f"Opacity set to {opacity} on {item_type} {index}."})


@media.command("set-reflection")
@click.argument("item_type")
@click.argument("index", type=int)
@click.option("--showing/--no-showing", default=True)
@click.option("--value", type=int, default=None, help="Reflection value 0-100.")
@click.pass_context
def media_set_reflection(ctx, item_type, index, showing, value):
    """Set reflection on an item."""
    from pages_cli.core.media import set_reflection_showing, set_reflection_value
    doc = _get_doc(ctx)
    _, err = _safe_run(set_reflection_showing, item_type=item_type, index=index, showing=showing, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set reflection: {err}"})
        return
    if value is not None:
        _, err = _safe_run(set_reflection_value, item_type=item_type, index=index, value=value, document=doc)
        if err:
            _output(ctx, {"error": f"Failed to set reflection value: {err}"})
            return
    _output(ctx, {"message": f"Reflection set on {item_type} {index}."})


@media.command("set-locked")
@click.argument("item_type")
@click.argument("index", type=int)
@click.argument("locked", type=bool)
@click.pass_context
def media_set_locked(ctx, item_type, index, locked):
    """Set locked state of an item."""
    from pages_cli.core.media import set_locked
    doc = _get_doc(ctx)
    _, err = _safe_run(set_locked, item_type=item_type, index=index, locked=locked, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set locked: {err}"})
        return
    _output(ctx, {"message": f"Locked set to {locked} on {item_type} {index}."})


@media.command("set-position")
@click.argument("item_type")
@click.argument("index", type=int)
@click.option("--x", type=int, required=True)
@click.option("--y", type=int, required=True)
@click.pass_context
def media_set_position(ctx, item_type, index, x, y):
    """Set position of an item."""
    from pages_cli.core.media import set_position
    doc = _get_doc(ctx)
    _, err = _safe_run(set_position, item_type=item_type, index=index, x=x, y=y, document=doc)
    if err:
        _output(ctx, {"error": f"Failed to set position: {err}"})
        return
    _output(ctx, {"message": f"Position set to ({x}, {y}) on {item_type} {index}."})


@media.command("set-size")
@click.argument("item_type")
@click.argument("index", type=int)
@click.option("--width", type=float, default=None)
@click.option("--height", type=float, default=None)
@click.pass_context
def media_set_size(ctx, item_type, index, width, height):
    """Set width and/or height of an item."""
    from pages_cli.core.media import set_width, set_height
    doc = _get_doc(ctx)
    if width is not None:
        _, err = _safe_run(set_width, item_type=item_type, index=index, width=width, document=doc)
        if err:
            _output(ctx, {"error": f"Failed to set width: {err}"})
            return
    if height is not None:
        _, err = _safe_run(set_height, item_type=item_type, index=index, height=height, document=doc)
        if err:
            _output(ctx, {"error": f"Failed to set height: {err}"})
            return
    _output(ctx, {"message": f"Size updated on {item_type} {index}."})


@media.command("list")
@click.option("--type", "item_type", default=None,
              help="Filter by type: image, shape, 'text item', 'audio clip', movie, line, group.")
@click.pass_context
def media_list(ctx, item_type):
    """List media items in the document."""
    doc = _get_doc(ctx)
    from pages_cli.core import media as media_mod
    state = ctx.ensure_object(SessionState)

    if item_type:
        list_fn = {
            "image": media_mod.list_images,
            "shape": media_mod.list_shapes,
            "text item": media_mod.list_text_items,
            "audio clip": media_mod.list_audio_clips,
            "movie": media_mod.list_movies,
            "line": media_mod.list_lines,
            "group": media_mod.list_groups,
        }.get(item_type)
        if not list_fn:
            _output(ctx, {"error": f"Unknown item type: {item_type}"})
            return
        result, err = _safe_run(list_fn, document=doc)
        if err:
            _output(ctx, {"error": f"Failed to list {item_type}s: {err}"})
            return
        if state.json_mode:
            _output(ctx, {"items": result, "count": len(result)})
        else:
            if not result:
                click.echo(f"No {item_type}s.")
            else:
                for it in result:
                    click.echo(f"  [{it.get('type', item_type)}] index={it.get('index', '?')}")
    else:
        result, err = _safe_run(media_mod.list_all_items, document=doc)
        if err:
            _output(ctx, {"error": f"Failed to list items: {err}"})
            return
        if state.json_mode:
            _output(ctx, result)
        else:
            if result.get("total", 0) == 0:
                click.echo("No media items.")
            else:
                for kind in ["images", "shapes", "text_items", "audio_clips", "movies", "lines", "groups", "tables"]:
                    items = result.get(kind, [])
                    if items:
                        click.echo(f"  {kind}: {len(items)} items")


# --- Audio-specific commands ---

@media.command("set-audio-volume")
@click.argument("index", type=int)
@click.argument("volume", type=int)
@click.pass_context
def media_set_audio_volume(ctx, index, volume):
    """Set audio clip volume (0-100)."""
    from pages_cli.core.media import set_audio_clip_volume
    doc = _get_doc(ctx)
    _, err = _safe_run(set_audio_clip_volume, index=index, volume=volume, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Audio clip {index} volume set to {volume}."})


@media.command("set-audio-repeat")
@click.argument("index", type=int)
@click.argument("method")
@click.pass_context
def media_set_audio_repeat(ctx, index, method):
    """Set audio clip repetition: none, loop, 'loop back and forth'."""
    from pages_cli.core.media import set_audio_repetition_method
    doc = _get_doc(ctx)
    _, err = _safe_run(set_audio_repetition_method, index=index, method=method, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Audio clip {index} repeat set to {method}."})


# --- Movie-specific commands ---

@media.command("set-movie-volume")
@click.argument("index", type=int)
@click.argument("volume", type=int)
@click.pass_context
def media_set_movie_volume(ctx, index, volume):
    """Set movie volume (0-100)."""
    from pages_cli.core.media import set_movie_volume
    doc = _get_doc(ctx)
    _, err = _safe_run(set_movie_volume, index=index, volume=volume, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Movie {index} volume set to {volume}."})


@media.command("set-movie-repeat")
@click.argument("index", type=int)
@click.argument("method")
@click.pass_context
def media_set_movie_repeat(ctx, index, method):
    """Set movie repetition: none, loop, 'loop back and forth'."""
    from pages_cli.core.media import set_movie_repetition_method
    doc = _get_doc(ctx)
    _, err = _safe_run(set_movie_repetition_method, index=index, method=method, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Movie {index} repeat set to {method}."})


# --- Line-specific commands ---

@media.command("set-line-start")
@click.argument("index", type=int)
@click.option("--x", type=int, required=True)
@click.option("--y", type=int, required=True)
@click.pass_context
def media_set_line_start(ctx, index, x, y):
    """Set line start point."""
    from pages_cli.core.media import set_line_start_point
    doc = _get_doc(ctx)
    _, err = _safe_run(set_line_start_point, index=index, x=x, y=y, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Line {index} start set to ({x}, {y})."})


@media.command("set-line-end")
@click.argument("index", type=int)
@click.option("--x", type=int, required=True)
@click.option("--y", type=int, required=True)
@click.pass_context
def media_set_line_end(ctx, index, x, y):
    """Set line end point."""
    from pages_cli.core.media import set_line_end_point
    doc = _get_doc(ctx)
    _, err = _safe_run(set_line_end_point, index=index, x=x, y=y, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Line {index} end set to ({x}, {y})."})


# --- Image description ---

@media.command("set-image-description")
@click.argument("index", type=int)
@click.argument("description")
@click.pass_context
def media_set_image_description(ctx, index, description):
    """Set VoiceOver description on image INDEX."""
    from pages_cli.core.media import set_image_description
    doc = _get_doc(ctx)
    _, err = _safe_run(set_image_description, index=index, description=description, document=doc)
    if err:
        _output(ctx, {"error": f"Failed: {err}"})
        return
    _output(ctx, {"message": f"Image {index} description set."})


# ===========================================================================
# export group
# ===========================================================================

@cli.group()
@click.pass_context
def export(ctx):
    """Export the current document to various formats."""
    pass


@export.command("pdf")
@click.argument("output_path")
@click.option("--password", default=None, help="Password-protect.")
@click.option("--password-hint", default=None, help="Password hint.")
@click.option("--image-quality", default=None, type=click.Choice(["Good", "Better", "Best"]))
@click.option("--include-comments/--no-comments", default=None)
@click.option("--include-annotations/--no-annotations", default=None)
@click.pass_context
def export_pdf(ctx, output_path, password, password_hint, image_quality, include_comments, include_annotations):
    """Export as PDF."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="PDF", document=doc,
        password=password, password_hint=password_hint,
        image_quality=image_quality,
        include_comments=include_comments,
        include_annotations=include_annotations,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("word")
@click.argument("output_path")
@click.option("--password", default=None)
@click.option("--image-quality", default=None, type=click.Choice(["Good", "Better", "Best"]))
@click.option("--include-comments/--no-comments", default=None)
@click.option("--include-annotations/--no-annotations", default=None)
@click.pass_context
def export_word(ctx, output_path, password, image_quality, include_comments, include_annotations):
    """Export as Microsoft Word (.docx)."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="Microsoft Word",
        document=doc, password=password, image_quality=image_quality,
        include_comments=include_comments,
        include_annotations=include_annotations,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("epub")
@click.argument("output_path")
@click.option("--title", "epub_title", default=None, help="EPUB title.")
@click.option("--author", "epub_author", default=None, help="EPUB author.")
@click.option("--genre", "epub_genre", default=None, help="EPUB genre.")
@click.option("--language", "epub_language", default=None, help="EPUB language.")
@click.option("--publisher", "epub_publisher", default=None, help="EPUB publisher.")
@click.option("--cover", "epub_cover", default=None, help="Cover image path.")
@click.option("--fixed-layout/--no-fixed-layout", default=None)
@click.pass_context
def export_epub(ctx, output_path, epub_title, epub_author, epub_genre,
                epub_language, epub_publisher, epub_cover, fixed_layout):
    """Export as EPUB."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="EPUB", document=doc,
        epub_title=epub_title, epub_author=epub_author, epub_genre=epub_genre,
        epub_language=epub_language, epub_publisher=epub_publisher,
        epub_cover=epub_cover, epub_fixed_layout=fixed_layout,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("text")
@click.argument("output_path")
@click.option("--image-quality", default=None, type=click.Choice(["Good", "Better", "Best"]))
@click.option("--include-comments/--no-comments", default=None)
@click.option("--include-annotations/--no-annotations", default=None)
@click.pass_context
def export_text(ctx, output_path, image_quality, include_comments, include_annotations):
    """Export as plain text."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="unformatted text",
        document=doc, image_quality=image_quality,
        include_comments=include_comments,
        include_annotations=include_annotations,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("rtf")
@click.argument("output_path")
@click.option("--image-quality", default=None, type=click.Choice(["Good", "Better", "Best"]))
@click.option("--include-comments/--no-comments", default=None)
@click.option("--include-annotations/--no-annotations", default=None)
@click.pass_context
def export_rtf(ctx, output_path, image_quality, include_comments, include_annotations):
    """Export as RTF (formatted text)."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="formatted text",
        document=doc, image_quality=image_quality,
        include_comments=include_comments,
        include_annotations=include_annotations,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("pages09")
@click.argument("output_path")
@click.option("--image-quality", default=None, type=click.Choice(["Good", "Better", "Best"]))
@click.option("--include-comments/--no-comments", default=None)
@click.option("--include-annotations/--no-annotations", default=None)
@click.pass_context
def export_pages09(ctx, output_path, image_quality, include_comments, include_annotations):
    """Export as Pages 09 format."""
    from pages_cli.core.export import export_document
    doc = _get_doc(ctx)
    abs_path = str(Path(output_path).expanduser().resolve())
    Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
    result, err = _safe_run(
        export_document, output_path=abs_path, format="Pages 09",
        document=doc, image_quality=image_quality,
        include_comments=include_comments,
        include_annotations=include_annotations,
    )
    if err:
        _output(ctx, {"error": f"Export failed: {err}"})
        return
    _output(ctx, {"message": f"Exported to: {abs_path}", **result})


@export.command("formats")
@click.pass_context
def export_formats(ctx):
    """List available export formats."""
    from pages_cli.core.export import get_export_formats
    format_names = get_export_formats()
    _EXT_MAP = {
        "PDF": ".pdf",
        "Microsoft Word": ".docx",
        "EPUB": ".epub",
        "unformatted text": ".txt",
        "formatted text": ".rtf",
        "Pages 09": ".pages",
    }
    _CLI_MAP = {
        "PDF": "pdf",
        "Microsoft Word": "word",
        "EPUB": "epub",
        "unformatted text": "text",
        "formatted text": "rtf",
        "Pages 09": "pages09",
    }
    formats = [
        {
            "name": _CLI_MAP.get(f, f),
            "description": f,
            "extension": _EXT_MAP.get(f, ""),
        }
        for f in format_names
    ]
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"formats": formats})
    else:
        skin = ReplSkin("pages", version=__version__)
        headers = ["Format", "Description", "Extension"]
        rows = [[f["name"], f["description"], f["extension"]] for f in formats]
        skin.table(headers, rows)


# ===========================================================================
# template group
# ===========================================================================

@cli.group()
@click.pass_context
def template(ctx):
    """Template management."""
    pass


@template.command("list")
@click.pass_context
def template_list(ctx):
    """List available templates."""
    from pages_cli.core.templates import list_templates
    result, err = _safe_run(list_templates)
    if err:
        _output(ctx, {"error": f"Failed to list templates: {err}"})
        return
    state = ctx.ensure_object(SessionState)
    if state.json_mode:
        _output(ctx, {"templates": result, "count": len(result)})
    else:
        if not result:
            click.echo("No templates found.")
        else:
            for i, t in enumerate(result, 1):
                click.echo(f"  {i}. {t}")


@template.command("info")
@click.argument("name")
@click.pass_context
def template_info(ctx, name):
    """Get details about a specific template."""
    from pages_cli.core.templates import get_template_info
    result, err = _safe_run(get_template_info, name=name)
    if err:
        _output(ctx, {"error": f"Failed to get template info: {err}"})
        return
    _output(ctx, result)


# ===========================================================================
# session group
# ===========================================================================

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


# ===========================================================================
# repl command
# ===========================================================================

@cli.command("repl")
@click.pass_context
def repl_command(ctx):
    """Enter interactive REPL mode."""
    _run_repl(ctx)


# ===========================================================================
# REPL implementation
# ===========================================================================

_REPL_HELP = {
    # document
    "document new":             "Create document [--template T] [--name N]",
    "document open":            "Open document <file_path>",
    "document close":           "Close document [--no-save]",
    "document save":            "Save document [--path PATH]",
    "document info":            "Show document info",
    "document list":            "List open documents",
    "document set-password":    "Set password <pw> [--hint H]",
    "document remove-password": "Remove password <pw>",
    "document placeholders":    "List placeholder texts",
    "document set-placeholder": "Set placeholder <index> <tag>",
    "document get-placeholder": "Get placeholder tag <index>",
    "document sections":        "Show section count",
    "document section-text":    "Get section body text <index>",
    "document page-text":       "Get page body text [index]",
    "document delete":          "Delete object <specifier>",
    # text
    "text add":                 "Add text <text>",
    "text add-paragraph":       "Add styled paragraph <text> [--font F] [--size S] ...",
    "text set":                 "Set body text <text>",
    "text get":                 "Get body text",
    "text set-font":            "Set font [--name F] [--size S] [--paragraph N] [--word N] [--character N]",
    "text get-font":            "Get font [--paragraph N] [--word N] [--character N]",
    "text get-font-size":       "Get font size [--paragraph N] [--word N] [--character N]",
    "text set-color":           "Set text color --r R --g G --b B [--paragraph] [--word] [--character]",
    "text get-color":           "Get text color [--paragraph] [--word] [--character]",
    "text bold":                "Set bold [--paragraph N] [--word N] [--character N]",
    "text italic":              "Set italic [--paragraph N] [--word N] [--character N]",
    "text bold-italic":         "Set bold-italic [--paragraph N] [--word N] [--character N]",
    "text style-paragraph":     "Style paragraph N [--font] [--size] [--bold] [--italic] ...",
    "text style-word":          "Style word N ...",
    "text style-character":     "Style character N ...",
    "text word-count":          "Get word/paragraph/character counts",
    # table
    "table add":                "Add table [--rows N] [--cols N] [--name N] ...",
    "table delete":             "Delete table [--name N] [--index N]",
    "table set-cell":           "Set cell <table> <row> <col> <value>",
    "table get-cell":           "Get cell <table> <row> <col>",
    "table get-formula":        "Get formula <table> <row> <col>",
    "table get-formatted-value":"Get formatted value <table> <row> <col>",
    "table list":               "List tables",
    "table info":               "Table info <table>",
    "table get-data":           "Get all data <table>",
    "table fill":               "Fill table <table> <json_data>",
    "table merge":              "Merge cells <table> <range>",
    "table unmerge":            "Unmerge cells <table> <range>",
    "table clear":              "Clear range <table> <range>",
    "table sort":               "Sort <table> --column COL [--descending]",
    "table set-cell-format":    "Set cell format <table> <row> <col> <format>",
    "table set-header-rows":    "Set header rows <table> <count>",
    "table set-header-cols":    "Set header cols <table> <count>",
    "table set-footer-rows":    "Set footer rows <table> <count>",
    "table set-row-height":     "Set row height <table> <row> <height>",
    "table get-row-height":     "Get row height <table> <row>",
    "table set-col-width":      "Set col width <table> <col> <width>",
    "table get-col-width":      "Get col width <table> <col>",
    "table range-set-font":     "Set range font <table> <range> <font>",
    "table range-set-font-size":"Set range font size <table> <range> <size>",
    "table range-set-format":   "Set range format <table> <range> <format>",
    "table range-set-alignment":"Set alignment <table> <range> <alignment>",
    "table range-set-vertical-alignment": "Set valign <table> <range> <alignment>",
    "table range-set-text-color":"Set text color on range --r --g --b",
    "table range-set-bg-color": "Set background color on range --r --g --b",
    "table range-set-text-wrap":"Set text wrap <table> <range> <true/false>",
    "table range-get-font-size":"Get range font size <table> <range>",
    "table range-get-vertical-alignment":"Get vertical alignment <table> <range>",
    "table range-get-text-color":"Get text color <table> <range>",
    "table range-get-bg-color": "Get background color <table> <range>",
    "table range-get-text-wrap":"Get text wrap <table> <range>",
    "table range-get-name":     "Get range name <table> <range>",
    "table get-row-address":    "Get row address <table> <row>",
    "table get-col-address":    "Get col address <table> <col>",
    # media
    "media add-image":          "Add image <file> [--x X] [--y Y] [--width W] [--height H]",
    "media add-shape":          "Add shape [--type T] [--x X] [--y Y] [--w W] [--h H] [--text T]",
    "media add-text-item":      "Add text box <text> [--x X] [--y Y]",
    "media add-audio":          "Add audio <file>",
    "media add-movie":          "Add movie <file>",
    "media add-line":           "Add line --start-x --start-y --end-x --end-y",
    "media delete":             "Delete item <type> [--index N] [--name N]",
    "media properties":         "Get properties <type> <index>",
    "media set-property":       "Set property <type> <index> <prop> <value>",
    "media get-property":       "Get property <type> <index> <prop>",
    "media set-rotation":       "Set rotation <type> <index> <degrees>",
    "media set-opacity":        "Set opacity <type> <index> <value>",
    "media set-reflection":     "Set reflection <type> <index> [--showing] [--value V]",
    "media set-locked":         "Set locked <type> <index> <true/false>",
    "media set-position":       "Set position <type> <index> --x X --y Y",
    "media set-size":           "Set size <type> <index> [--width W] [--height H]",
    "media set-audio-volume":   "Set audio volume <index> <0-100>",
    "media set-audio-repeat":   "Set audio repeat <index> <method>",
    "media set-movie-volume":   "Set movie volume <index> <0-100>",
    "media set-movie-repeat":   "Set movie repeat <index> <method>",
    "media set-line-start":     "Set line start <index> --x X --y Y",
    "media set-line-end":       "Set line end <index> --x X --y Y",
    "media set-image-description":"Set image alt text <index> <description>",
    "media list":               "List media items [--type T]",
    # export
    "export pdf":               "Export as PDF <output> [--password P] [--image-quality Q] ...",
    "export word":              "Export as Word <output> [--password P]",
    "export epub":              "Export as EPUB <output> [--title] [--author] [--genre] ...",
    "export text":              "Export as plain text <output>",
    "export rtf":               "Export as RTF <output>",
    "export pages09":           "Export as Pages 09 <output>",
    "export formats":           "List export formats",
    # template
    "template list":            "List templates",
    "template info":            "Get template details <name>",
    # session
    "session status":           "Show session status",
    "session save":             "Save session <path>",
    "session load":             "Load session <path>",
    # misc
    "help":                     "Show this help message",
    "quit / exit":              "Exit the REPL",
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

        tokens = _tokenize_input(user_input)
        if not tokens:
            continue

        args = list(tokens)
        if state.json_mode and "--json" not in args:
            args = ["--json"] + args

        try:
            cli.main(
                args=args,
                prog_name="pages-cli",
                standalone_mode=False,
                **{"obj": state},
            )
        except click.exceptions.UsageError as exc:
            skin.error(str(exc))
            skin.hint("Type 'help' to see available commands.")
        except click.exceptions.Abort:
            skin.warning("Command aborted.")
        except SystemExit:
            pass
        except Exception as exc:
            skin.error(f"Unexpected error: {exc}")


# ===========================================================================
# Entry point
# ===========================================================================

def main():
    """Entry point for setup.py console_scripts."""
    cli(obj=SessionState())


if __name__ == "__main__":
    main()
