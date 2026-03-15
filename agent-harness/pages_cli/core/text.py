"""Text manipulation module for Apple Pages."""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running


def _doc_target(document: str | None) -> str:
    """Return the AppleScript target reference for a document.

    Args:
        document: Document name, or None for front document.

    Returns:
        str: AppleScript document reference.
    """
    if document:
        escaped = document.replace('"', '\\"')
        return f'document "{escaped}"'
    return "front document"


def add_text(text: str, document: str | None = None) -> None:
    """Append text to the body of the current document.

    Args:
        text: The text to append.
        document: Optional document name. Uses front document if None.

    Raises:
        RuntimeError: If the text cannot be added.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_text = text.replace("\\", "\\\\").replace('"', '\\"')

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to (body text as text) & "{escaped_text}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def add_paragraph(
    text: str,
    font: str | None = None,
    size: int | None = None,
    color: tuple[int, int, int] | None = None,
    bold: bool = False,
    italic: bool = False,
    document: str | None = None,
) -> None:
    """Add a styled paragraph to the document.

    Args:
        text: The paragraph text.
        font: Optional font name (e.g. "Helvetica").
        size: Optional font size in points.
        color: Optional RGB color tuple with values 0-65535.
        bold: Whether to make text bold.
        italic: Whether to make text italic.
        document: Optional document name.

    Raises:
        RuntimeError: If the paragraph cannot be added.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_text = text.replace("\\", "\\\\").replace('"', '\\"')

    # First append the text as a new paragraph
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to (body text as text) & "\\n{escaped_text}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)

    # Now style the last paragraph
    style_props = []
    if font:
        escaped_font = font.replace('"', '\\"')
        style_props.append(f'font:"{escaped_font}"')
    if size is not None:
        style_props.append(f"size:{size}")
    if color is not None:
        r, g, b = color
        style_props.append(f"color:{{{r}, {g}, {b}}}")
    if bold:
        style_props.append("bold:true")
    if italic:
        style_props.append("italic:true")

    if style_props:
        props_str = ", ".join(style_props)
        style_script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            f'    set paraCount to count of paragraphs of body text\n'
            f'    set properties of paragraph paraCount of body text to {{{props_str}}}\n'
            f'  end tell\n'
            f'end tell'
        )
        _run_applescript(style_script)


def set_body_text(text: str, document: str | None = None) -> None:
    """Set the entire body text of the document.

    Args:
        text: The text to set as the body.
        document: Optional document name.

    Raises:
        RuntimeError: If the body text cannot be set.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_text = text.replace("\\", "\\\\").replace('"', '\\"')

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to "{escaped_text}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_body_text(document: str | None = None) -> str:
    """Get the body text of the document.

    Args:
        document: Optional document name.

    Returns:
        str: The body text of the document.

    Raises:
        RuntimeError: If the body text cannot be retrieved.
    """
    ensure_pages_running()
    target = _doc_target(document)

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get body text\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script)


def get_word_count(document: str | None = None) -> int:
    """Get the word count of the document.

    Args:
        document: Optional document name.

    Returns:
        int: The number of words in the document.

    Raises:
        RuntimeError: If the word count cannot be retrieved.
    """
    ensure_pages_running()
    target = _doc_target(document)

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of words of body text\n'
        f'  end tell\n'
        f'end tell'
    )
    result = _run_applescript(script)
    return int(result.strip())


def set_font(
    font_name: str,
    paragraph_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the font for the document or a specific paragraph.

    Args:
        font_name: The font name (e.g. "Helvetica", "Times New Roman").
        paragraph_index: Optional 1-based paragraph index. If None, applies to all text.
        document: Optional document name.

    Raises:
        RuntimeError: If the font cannot be set.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_font = font_name.replace('"', '\\"')

    if paragraph_index is not None:
        text_target = f"paragraph {paragraph_index} of body text"
    else:
        text_target = "body text"

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set font of {text_target} to "{escaped_font}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def set_font_size(
    size: int,
    paragraph_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the font size for the document or a specific paragraph.

    Args:
        size: The font size in points.
        paragraph_index: Optional 1-based paragraph index. If None, applies to all text.
        document: Optional document name.

    Raises:
        RuntimeError: If the font size cannot be set.
    """
    ensure_pages_running()
    target = _doc_target(document)

    if paragraph_index is not None:
        text_target = f"paragraph {paragraph_index} of body text"
    else:
        text_target = "body text"

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set size of {text_target} to {size}\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def set_text_color(
    r: int,
    g: int,
    b: int,
    paragraph_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the text color using RGB values (0-65535 range).

    Args:
        r: Red component (0-65535).
        g: Green component (0-65535).
        b: Blue component (0-65535).
        paragraph_index: Optional 1-based paragraph index. If None, applies to all text.
        document: Optional document name.

    Raises:
        RuntimeError: If the color cannot be set.
    """
    ensure_pages_running()
    target = _doc_target(document)

    if paragraph_index is not None:
        text_target = f"paragraph {paragraph_index} of body text"
    else:
        text_target = "body text"

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set color of {text_target} to {{{r}, {g}, {b}}}\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)
