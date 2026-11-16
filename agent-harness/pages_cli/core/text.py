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


# ---------------------------------------------------------------------------
# Find / replace that keeps paragraph styles
#
# Pages scripting can replace a single word or character in place and the
# paragraph keeps its font, size and style. Replacing the whole body text, or
# assigning to a character range, does not (a range assignment corrupts the
# text). So every span edit here is built from single-character assignments:
# delete the span's trailing characters from last to first, then set its first
# character to the replacement. Edits are applied last-to-first so earlier
# positions stay valid, then the paragraphs are re-read and compared with what
# the edit should have produced.
# ---------------------------------------------------------------------------

# Not whitespace, so osascript output stripping cannot eat an empty or
# space-padded paragraph at either end.
_PARA_MARK = "\u241e"
_PARA_MARK_ID = ord(_PARA_MARK)


def get_paragraphs(document: str | None = None) -> list[str]:
    """Return the body text paragraphs, without their line terminators."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  set d to {target}\n'
        f'  set mark to character id {_PARA_MARK_ID}\n'
        f'  set out to ""\n'
        f'  repeat with i from 1 to (count of paragraphs of body text of d)\n'
        f'    set out to out & mark & ((paragraph i of body text of d) as text)\n'
        f'  end repeat\n'
        f'  return out & mark\n'
        f'end tell'
    )
    raw = _run_applescript(script)
    if not raw.startswith(_PARA_MARK) or not raw.endswith(_PARA_MARK):
        raise RuntimeError("Unexpected paragraph listing from Pages.")
    if raw == _PARA_MARK:
        return []
    return raw[1:-1].split(_PARA_MARK)


def find_in_paragraphs(
    paragraphs: list[str],
    phrase: str,
    paragraph: int | None = None,
    ignore_case: bool = False,
) -> list[dict]:
    """Find every non-overlapping occurrence of ``phrase``.

    Positions are 1-based character offsets within the paragraph; ``end`` is
    the position of the last matched character.
    """
    if not phrase:
        raise ValueError("Search text must not be empty.")
    if "\n" in phrase or "\r" in phrase:
        raise ValueError("Search text must sit within one paragraph (no line breaks).")
    if paragraph is not None and not 1 <= paragraph <= len(paragraphs):
        raise ValueError(
            f"Paragraph {paragraph} out of range (document has {len(paragraphs)})."
        )
    matches: list[dict] = []
    for number, text in enumerate(paragraphs, start=1):
        if paragraph is not None and number != paragraph:
            continue
        haystack, needle = text, phrase
        if ignore_case:
            lowered = text.lower()
            # str.lower() can change length (e.g. U+0130); positions would drift.
            if len(lowered) == len(text):
                haystack, needle = lowered, phrase.lower()
        pos = haystack.find(needle)
        while pos != -1:
            end = pos + len(phrase)
            matches.append({
                "paragraph": number,
                "start": pos + 1,
                "end": end,
                "match": text[pos:end],
                "context": text[max(0, pos - 20):end + 20],
            })
            pos = haystack.find(needle, end)
    return matches


def find_text(
    phrase: str,
    paragraph: int | None = None,
    ignore_case: bool = False,
    document: str | None = None,
) -> list[dict]:
    """Find ``phrase`` in the document body (see find_in_paragraphs)."""
    return find_in_paragraphs(
        get_paragraphs(document), phrase, paragraph, ignore_case
    )


def _apply_ops(paragraphs: list[str], ops: list[tuple[int, int, int, str]]) -> list[str]:
    """Predict the paragraphs after span edits (paragraph, start, end, new).

    A line feed in ``new`` starts a new paragraph, as it does in Pages.
    """
    result = list(paragraphs)
    for number, start, end, new in sorted(ops, reverse=True):
        text = result[number - 1]
        edited = text[:start - 1] + new + text[end:]
        result[number - 1:number] = edited.split("\n")
    return result


# Each character assignment is one Apple event of roughly 20 ms, and every
# osascript call is killed after 30 s. A bulk edit is therefore sent in batches
# of this many assignments (about 12 s each) rather than as one long script.
_MAX_STEPS_PER_SCRIPT = 600

_AUTOSAVE_WARNING = (
    "Pages autosaves open documents into their file within about 30 seconds, and "
    "closing without saving does not undo that, so restore from your backup copy "
    "if you need to undo."
)


def _span_steps(number: int, start: int, end: int, new: str) -> list[tuple[int, int, str]]:
    """Character assignments for one span: trailing characters last to first, then the first."""
    steps = [(number, k, "") for k in range(end, start, -1)]
    steps.append((number, start, new))
    return steps


def _steps_script(steps: list[tuple[int, int, str]], target: str) -> str:
    lines = "\n".join(
        f'    set character {k} of paragraph {number} of body text of d to "{_esc(text)}"'
        for number, k, text in steps
    )
    return (
        f'tell application "Pages"\n'
        f'  set d to {target}\n'
        f'  tell d\n'
        f'{lines}\n'
        f'  end tell\n'
        f'end tell'
    )


def _edit_spans(
    paragraphs: list[str],
    ops: list[tuple[int, int, int, str]],
    document: str | None,
) -> list[str]:
    """Apply span edits in one script, then verify and return the paragraphs."""
    for number, start, end, new in ops:
        if not 1 <= number <= len(paragraphs):
            raise ValueError(
                f"Paragraph {number} out of range (document has {len(paragraphs)})."
            )
        length = len(paragraphs[number - 1])
        if not 1 <= start <= end <= length:
            raise ValueError(
                f"Characters {start}-{end} out of range "
                f"(paragraph {number} has {length})."
            )
        if "\r" in new:
            raise ValueError("Use a line feed, not a carriage return, for new paragraphs.")

    expected = _apply_ops(paragraphs, ops)
    steps = [
        step
        for number, start, end, new in sorted(ops, reverse=True)
        for step in _span_steps(number, start, end, new)
    ]
    target = _doc_target(document)
    for first in range(0, len(steps), _MAX_STEPS_PER_SCRIPT):
        batch = steps[first:first + _MAX_STEPS_PER_SCRIPT]
        try:
            _run_applescript(_steps_script(batch, target))
        except RuntimeError as exc:
            raise RuntimeError(
                f"Edit stopped part-way: {first} of {len(steps)} character "
                f"assignments were applied, so the document is partly edited. "
                f"{_AUTOSAVE_WARNING} Cause: {exc}"
            ) from None

    actual = get_paragraphs(document)
    if actual != expected:
        for i in range(max(len(actual), len(expected))):
            got = actual[i] if i < len(actual) else None
            want = expected[i] if i < len(expected) else None
            if got != want:
                raise RuntimeError(
                    f"Edit did not produce the expected text at paragraph {i + 1}: "
                    f"expected {want!r}, found {got!r}. The document is modified. "
                    f"{_AUTOSAVE_WARNING}"
                )
    return actual


def replace_text(
    old: str,
    new: str,
    paragraph: int | None = None,
    occurrence: int | None = None,
    ignore_case: bool = False,
    document: str | None = None,
) -> dict:
    """Replace occurrences of ``old`` with ``new``, keeping paragraph styles.

    Replaces every occurrence unless ``occurrence`` (1-based, counted among
    the matches within scope) selects one. Nothing is saved.
    """
    paragraphs = get_paragraphs(document)
    matches = find_in_paragraphs(paragraphs, old, paragraph, ignore_case)
    if occurrence is not None:
        if not 1 <= occurrence <= len(matches):
            raise ValueError(
                f"Occurrence {occurrence} out of range ({len(matches)} found)."
            )
        matches = [matches[occurrence - 1]]
    if not matches:
        return {"replaced": 0, "paragraphs": []}
    ops = [(m["paragraph"], m["start"], m["end"], new) for m in matches]
    _edit_spans(paragraphs, ops, document)
    return {
        "replaced": len(ops),
        "paragraphs": sorted({m["paragraph"] for m in matches}),
    }


def replace_range(
    paragraph: int,
    start: int,
    end: int,
    new: str,
    document: str | None = None,
) -> dict:
    """Replace characters ``start``-``end`` of a paragraph with ``new``.

    ``new`` may be empty (delete). Nothing is saved.
    """
    paragraphs = get_paragraphs(document)
    actual = _edit_spans(paragraphs, [(paragraph, start, end, new)], document)
    return {"paragraph": paragraph, "text": actual[paragraph - 1]}


def replace_word(
    paragraph: int,
    word_index: int,
    new: str,
    document: str | None = None,
) -> dict:
    """Replace word ``word_index`` of a paragraph, keeping its style.

    Word boundaries are Pages' own, so the result is checked only for the
    other paragraphs being untouched. Nothing is saved.
    """
    if "\r" in new:
        raise ValueError("Use a line feed, not a carriage return, for new paragraphs.")
    paragraphs = get_paragraphs(document)
    if not 1 <= paragraph <= len(paragraphs):
        raise ValueError(
            f"Paragraph {paragraph} out of range (document has {len(paragraphs)})."
        )
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  set d to {target}\n'
        f'  set word {word_index} of paragraph {paragraph} of body text of d to "{_esc(new)}"\n'
        f'end tell'
    )
    _run_applescript(script)
    actual = get_paragraphs(document)
    shift = len(actual) - len(paragraphs)
    untouched = (
        actual[:paragraph - 1] == paragraphs[:paragraph - 1]
        and (actual[paragraph + shift:] == paragraphs[paragraph:])
    )
    if shift < 0 or not untouched:
        raise RuntimeError(
            "Word replacement changed text outside the target paragraph. "
            f"The document is modified. {_AUTOSAVE_WARNING}"
        )
    return {"paragraph": paragraph, "text": "\n".join(actual[paragraph - 1:paragraph + shift])}


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
