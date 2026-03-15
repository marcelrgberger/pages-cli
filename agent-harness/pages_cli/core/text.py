"""Text manipulation module for Apple Pages.

Supports paragraph-level, word-level, and character-level formatting
using the verified AppleScript API. Bold/italic are achieved by switching
to the appropriate font variant (e.g. Helvetica -> Helvetica-Bold).

Full coverage: font, size, color on body text / paragraph N / word N /
character N, bold/italic via font variant names, paragraph count, word
count, character count.
"""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running
from pages_cli.utils.helpers import _doc_target, _esc


# ---------------------------------------------------------------------------
# Font variant mappings (known system font families)
# ---------------------------------------------------------------------------

_BOLD_MAP = {
    "Helvetica": "Helvetica-Bold",
    "Helvetica-Oblique": "Helvetica-BoldOblique",
    "HelveticaNeue": "HelveticaNeue-Bold",
    "HelveticaNeue-Italic": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Light": "HelveticaNeue-Bold",
    "HelveticaNeue-Medium": "HelveticaNeue-Bold",
    "HelveticaNeue-Thin": "HelveticaNeue-Bold",
    "Georgia": "Georgia-Bold",
    "Georgia-Italic": "Georgia-BoldItalic",
    "TimesNewRomanPSMT": "TimesNewRomanPS-BoldMT",
    "TimesNewRomanPS-ItalicMT": "TimesNewRomanPS-BoldItalicMT",
    "Avenir-Book": "Avenir-Heavy",
    "Avenir-Light": "Avenir-Heavy",
    "Avenir-Medium": "Avenir-Heavy",
}

_ITALIC_MAP = {
    "Helvetica": "Helvetica-Oblique",
    "Helvetica-Bold": "Helvetica-BoldOblique",
    "HelveticaNeue": "HelveticaNeue-Italic",
    "HelveticaNeue-Bold": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Light": "HelveticaNeue-Italic",
    "HelveticaNeue-Medium": "HelveticaNeue-Italic",
    "HelveticaNeue-Thin": "HelveticaNeue-Italic",
    "Georgia": "Georgia-Italic",
    "Georgia-Bold": "Georgia-BoldItalic",
    "TimesNewRomanPSMT": "TimesNewRomanPS-ItalicMT",
    "TimesNewRomanPS-BoldMT": "TimesNewRomanPS-BoldItalicMT",
    "Avenir-Book": "Avenir-BookOblique",
    "Avenir-Light": "Avenir-LightOblique",
    "Avenir-Medium": "Avenir-MediumOblique",
    "Avenir-Heavy": "Avenir-HeavyOblique",
}

_BOLD_ITALIC_MAP = {
    "Helvetica": "Helvetica-BoldOblique",
    "Helvetica-Bold": "Helvetica-BoldOblique",
    "Helvetica-Oblique": "Helvetica-BoldOblique",
    "HelveticaNeue": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Bold": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Italic": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Light": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Medium": "HelveticaNeue-BoldItalic",
    "HelveticaNeue-Thin": "HelveticaNeue-BoldItalic",
    "Georgia": "Georgia-BoldItalic",
    "Georgia-Bold": "Georgia-BoldItalic",
    "Georgia-Italic": "Georgia-BoldItalic",
    "TimesNewRomanPSMT": "TimesNewRomanPS-BoldItalicMT",
    "TimesNewRomanPS-BoldMT": "TimesNewRomanPS-BoldItalicMT",
    "TimesNewRomanPS-ItalicMT": "TimesNewRomanPS-BoldItalicMT",
    "Avenir-Book": "Avenir-HeavyOblique",
    "Avenir-Light": "Avenir-HeavyOblique",
    "Avenir-Medium": "Avenir-HeavyOblique",
    "Avenir-Heavy": "Avenir-HeavyOblique",
}


def _text_target(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
) -> str:
    """Build the AppleScript text target specifier.

    Levels are mutually exclusive -- character is most specific, then word,
    then paragraph. If none are given, targets all body text.
    """
    if character_index is not None:
        return f"character {character_index} of body text"
    if word_index is not None:
        return f"word {word_index} of body text"
    if paragraph_index is not None:
        return f"paragraph {paragraph_index} of body text"
    return "body text"


def _resolve_font(current_font: str, variant_map: dict[str, str]) -> str:
    """Resolve a font to its variant. Returns as-is if not found."""
    return variant_map.get(current_font, current_font)


# ---------------------------------------------------------------------------
# Core text operations
# ---------------------------------------------------------------------------

def add_text(text: str, document: str | None = None) -> None:
    """Append text to the body of the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to (body text as text) & "{_esc(text)}"\n'
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
    """Add a styled paragraph to the document."""
    ensure_pages_running()
    target = _doc_target(document)

    # Append text as new paragraph
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to (body text as text) & "\\n{_esc(text)}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)

    # Determine effective font
    effective_font = font
    if bold and italic:
        effective_font = _resolve_font(
            effective_font or "Helvetica", _BOLD_ITALIC_MAP
        )
    elif bold:
        effective_font = _resolve_font(
            effective_font or "Helvetica", _BOLD_MAP
        )
    elif italic:
        effective_font = _resolve_font(
            effective_font or "Helvetica", _ITALIC_MAP
        )

    # Style the last paragraph
    style_lines = []
    if effective_font:
        style_lines.append(
            f'    set font of paragraph paraCount of body text to "{_esc(effective_font)}"'
        )
    if size is not None:
        style_lines.append(
            f'    set size of paragraph paraCount of body text to {size}'
        )
    if color is not None:
        r, g, b = color
        style_lines.append(
            f'    set color of paragraph paraCount of body text to {{{r}, {g}, {b}}}'
        )

    if style_lines:
        style_script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            f'    set paraCount to count of paragraphs of body text\n'
            + "\n".join(style_lines) + "\n"
            f'  end tell\n'
            f'end tell'
        )
        _run_applescript(style_script)


def set_body_text(text: str, document: str | None = None) -> None:
    """Set the entire body text of the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set body text to "{_esc(text)}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_body_text(document: str | None = None) -> str:
    """Get the body text of the document."""
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


# ---------------------------------------------------------------------------
# Counting
# ---------------------------------------------------------------------------

def get_word_count(document: str | None = None) -> int:
    """Get the word count of the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of words of body text\n'
        f'  end tell\n'
        f'end tell'
    )
    return int(_run_applescript(script).strip())


def get_paragraph_count(document: str | None = None) -> int:
    """Get the paragraph count of the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of paragraphs of body text\n'
        f'  end tell\n'
        f'end tell'
    )
    return int(_run_applescript(script).strip())


def get_character_count(document: str | None = None) -> int:
    """Get the character count of the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of characters of body text\n'
        f'  end tell\n'
        f'end tell'
    )
    return int(_run_applescript(script).strip())


def get_counts(document: str | None = None) -> dict:
    """Get word, paragraph, and character counts in one call."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set wc to count of words of body text\n'
        f'    set pc to count of paragraphs of body text\n'
        f'    set cc to count of characters of body text\n'
        f'    return (wc as text) & "|" & (pc as text) & "|" & (cc as text)\n'
        f'  end tell\n'
        f'end tell'
    )
    result = _run_applescript(script).strip()
    parts = result.split("|")
    return {
        "words": int(parts[0]) if len(parts) > 0 else 0,
        "paragraphs": int(parts[1]) if len(parts) > 1 else 0,
        "characters": int(parts[2]) if len(parts) > 2 else 0,
    }


# ---------------------------------------------------------------------------
# Font / size / color  --  ALL levels (body, paragraph, word, character)
# ---------------------------------------------------------------------------

def set_font(
    font_name: str,
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the font at any level (all text, paragraph, word, character)."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set font of {text_ref} to "{_esc(font_name)}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_font(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> str:
    """Get the font at any level."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get font of {text_ref}\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script).strip()


def set_font_size(
    size: float,
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the font size at any level."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set size of {text_ref} to {size}\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_font_size(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> float:
    """Get the font size at any level."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get size of {text_ref}\n'
        f'  end tell\n'
        f'end tell'
    )
    return float(_run_applescript(script).strip())


def set_text_color(
    r: int, g: int, b: int,
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set the text color (RGB 0-65535) at any level."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set color of {text_ref} to {{{r}, {g}, {b}}}\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_text_color(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> str:
    """Get the text color at any level (returns comma-separated RGB)."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get color of {text_ref}\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Bold / Italic convenience  --  ALL levels
# ---------------------------------------------------------------------------

def _apply_variant(
    variant_map: dict[str, str],
    fallback_font: str,
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Read the current font and switch to the variant."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = _text_target(paragraph_index, word_index, character_index)

    get_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get font of {text_ref}\n'
        f'  end tell\n'
        f'end tell'
    )
    current_font = _run_applescript(get_script).strip()
    new_font = _resolve_font(current_font, variant_map)
    if new_font == current_font:
        # Not in map -- use fallback
        new_font = fallback_font

    set_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set font of {text_ref} to "{_esc(new_font)}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(set_script)


def set_bold(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set text to bold by switching to the bold font variant."""
    _apply_variant(
        _BOLD_MAP, "Helvetica-Bold",
        paragraph_index, word_index, character_index, document,
    )


def set_italic(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set text to italic by switching to the italic font variant."""
    _apply_variant(
        _ITALIC_MAP, "Helvetica-Oblique",
        paragraph_index, word_index, character_index, document,
    )


def set_bold_italic(
    paragraph_index: int | None = None,
    word_index: int | None = None,
    character_index: int | None = None,
    document: str | None = None,
) -> None:
    """Set text to bold-italic by switching to the bold-italic font variant."""
    _apply_variant(
        _BOLD_ITALIC_MAP, "Helvetica-BoldOblique",
        paragraph_index, word_index, character_index, document,
    )


# ---------------------------------------------------------------------------
# Convenience: style a paragraph in one call
# ---------------------------------------------------------------------------

def set_paragraph_style(
    paragraph_index: int,
    font: str | None = None,
    size: float | None = None,
    color: tuple[int, int, int] | None = None,
    bold: bool = False,
    italic: bool = False,
    document: str | None = None,
) -> None:
    """Style a paragraph with font, size, color, bold and italic."""
    ensure_pages_running()
    target = _doc_target(document)
    text_ref = f"paragraph {paragraph_index} of body text"

    # Determine effective font
    effective_font = font
    if bold and italic:
        if effective_font:
            effective_font = _resolve_font(effective_font, _BOLD_ITALIC_MAP)
        else:
            get_script = (
                f'tell application "Pages"\n'
                f'  tell {target}\n'
                f'    get font of {text_ref}\n'
                f'  end tell\n'
                f'end tell'
            )
            current = _run_applescript(get_script).strip()
            effective_font = _resolve_font(current, _BOLD_ITALIC_MAP)
    elif bold:
        if effective_font:
            effective_font = _resolve_font(effective_font, _BOLD_MAP)
        else:
            get_script = (
                f'tell application "Pages"\n'
                f'  tell {target}\n'
                f'    get font of {text_ref}\n'
                f'  end tell\n'
                f'end tell'
            )
            current = _run_applescript(get_script).strip()
            effective_font = _resolve_font(current, _BOLD_MAP)
    elif italic:
        if effective_font:
            effective_font = _resolve_font(effective_font, _ITALIC_MAP)
        else:
            get_script = (
                f'tell application "Pages"\n'
                f'  tell {target}\n'
                f'    get font of {text_ref}\n'
                f'  end tell\n'
                f'end tell'
            )
            current = _run_applescript(get_script).strip()
            effective_font = _resolve_font(current, _ITALIC_MAP)

    style_lines = []
    if effective_font:
        style_lines.append(f'    set font of {text_ref} to "{_esc(effective_font)}"')
    if size is not None:
        style_lines.append(f'    set size of {text_ref} to {size}')
    if color is not None:
        r, g, b = color
        style_lines.append(f'    set color of {text_ref} to {{{r}, {g}, {b}}}')

    if style_lines:
        script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            + "\n".join(style_lines) + "\n"
            f'  end tell\n'
            f'end tell'
        )
        _run_applescript(script)


def set_word_style(
    word_index: int,
    font: str | None = None,
    size: float | None = None,
    color: tuple[int, int, int] | None = None,
    bold: bool = False,
    italic: bool = False,
    document: str | None = None,
) -> None:
    """Style a word with font, size, color, bold and italic."""
    if font:
        effective_font = font
        if bold and italic:
            effective_font = _resolve_font(font, _BOLD_ITALIC_MAP)
        elif bold:
            effective_font = _resolve_font(font, _BOLD_MAP)
        elif italic:
            effective_font = _resolve_font(font, _ITALIC_MAP)
        set_font(effective_font, word_index=word_index, document=document)
    elif bold and italic:
        set_bold_italic(word_index=word_index, document=document)
    elif bold:
        set_bold(word_index=word_index, document=document)
    elif italic:
        set_italic(word_index=word_index, document=document)

    if size is not None:
        set_font_size(size, word_index=word_index, document=document)
    if color is not None:
        r, g, b = color
        set_text_color(r, g, b, word_index=word_index, document=document)


def set_character_style(
    character_index: int,
    font: str | None = None,
    size: float | None = None,
    color: tuple[int, int, int] | None = None,
    bold: bool = False,
    italic: bool = False,
    document: str | None = None,
) -> None:
    """Style a character with font, size, color, bold and italic."""
    if font:
        effective_font = font
        if bold and italic:
            effective_font = _resolve_font(font, _BOLD_ITALIC_MAP)
        elif bold:
            effective_font = _resolve_font(font, _BOLD_MAP)
        elif italic:
            effective_font = _resolve_font(font, _ITALIC_MAP)
        set_font(effective_font, character_index=character_index, document=document)
    elif bold and italic:
        set_bold_italic(character_index=character_index, document=document)
    elif bold:
        set_bold(character_index=character_index, document=document)
    elif italic:
        set_italic(character_index=character_index, document=document)

    if size is not None:
        set_font_size(size, character_index=character_index, document=document)
    if color is not None:
        r, g, b = color
        set_text_color(r, g, b, character_index=character_index, document=document)
