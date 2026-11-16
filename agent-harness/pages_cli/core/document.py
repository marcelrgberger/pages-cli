"""Document management module for Apple Pages.

Full sdef coverage: create, open, close, save, info, list,
set password, remove password, placeholder text, sections, page body text.
"""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running
from pages_cli.utils.helpers import _doc_target, _esc


# ---------------------------------------------------------------------------
# Create / Open / Close / Save
# ---------------------------------------------------------------------------

def create_document(template: str = "Blank", name: str | None = None) -> dict:
    """Create a new Pages document from a template.

    Returns:
        dict: Document info.
    """
    ensure_pages_running()

    if template == "Blank":
        script = 'tell application "Pages" to make new document'
    else:
        script = (
            'tell application "Pages"\n'
            f'  make new document with properties {{document template:template "{_esc(template)}"}}\n'
            'end tell'
        )
    _run_applescript(script)

    if name:
        _run_applescript(
            f'tell application "Pages" to set name of front document to "{_esc(name)}"'
        )

    return get_document_info()


def open_document(path: str) -> dict:
    """Open an existing .pages file.

    Raises if Pages hands back no document. The returned info describes the
    opened document itself (not whichever window is frontmost) and carries
    ``was_open``: True if the file was already open before this call, so the
    caller can avoid closing a document someone else had open.
    """
    ensure_pages_running()
    script = (
        f'set wanted to (POSIX file "{_esc(path)}") as text\n'
        f'tell application "Pages"\n'
        f'  set wasOpen to false\n'
        f'  repeat with x in documents\n'
        f'    try\n'
        f'      if (file of x as text) is wanted then set wasOpen to true\n'
        f'    end try\n'
        f'  end repeat\n'
        f'  set d to open POSIX file "{_esc(path)}"\n'
        f'  if d is missing value then error "Pages returned no document for {_esc(path)}"\n'
        f'  return (name of d) & linefeed & (wasOpen as text)\n'
        f'end tell'
    )
    name, _, was_open = _run_applescript(script).partition("\n")
    info = get_document_info(name)
    info["was_open"] = was_open.strip() == "true"
    return info


def close_document(name: str | None = None, saving: bool = True) -> None:
    """Close a Pages document."""
    ensure_pages_running()
    save_flag = "saving yes" if saving else "saving no"
    target = _doc_target(name)
    script = f'tell application "Pages" to close {target} {save_flag}'
    _run_applescript(script)


def save_document(
    name: str | None = None,
    path: str | None = None,
    reopen: bool = False,
) -> str:
    """Save the current or named document and return the document's name.

    ``save ... in POSIX file`` writes the file but leaves a new document's
    window Untitled and disconnected from it. With ``reopen=True`` (requires
    ``path``) that window is closed without saving and the written file is
    opened, so the open document is the real file.
    """
    if reopen and not path:
        raise ValueError("reopen requires a path to save to")
    ensure_pages_running()
    target = _doc_target(name)
    if path and reopen:
        script = (
            f'tell application "Pages"\n'
            f'  set d to {target}\n'
            f'  save d in POSIX file "{_esc(path)}"\n'
            f'  delay 1\n'
            f'  close d saving no\n'
            f'  delay 1\n'
            f'  set f to open POSIX file "{_esc(path)}"\n'
            f'  if f is missing value then error "Saved {_esc(path)} but Pages would not reopen it"\n'
            f'  return name of f\n'
            f'end tell'
        )
    elif path:
        script = (
            f'tell application "Pages"\n'
            f'  set d to {target}\n'
            f'  save d in POSIX file "{_esc(path)}"\n'
            f'  return name of d\n'
            f'end tell'
        )
    else:
        script = (
            f'tell application "Pages"\n'
            f'  set d to {target}\n'
            f'  save d\n'
            f'  return name of d\n'
            f'end tell'
        )
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Info / List
# ---------------------------------------------------------------------------

def get_document_info(name: str | None = None) -> dict:
    """Get information about a Pages document."""
    ensure_pages_running()
    target = _doc_target(name)

    doc_name = _run_applescript(
        f'tell application "Pages" to get name of {target}'
    ).strip()

    try:
        doc_path = _run_applescript(
            f'tell application "Pages" to get file of {target} as text'
        ).strip()
    except RuntimeError:
        doc_path = ""

    try:
        modified = _run_applescript(
            f'tell application "Pages" to get modified of {target}'
        ).strip().lower() == "true"
    except RuntimeError:
        modified = False

    try:
        page_count = int(_run_applescript(
            f'tell application "Pages" to get count of pages of {target}'
        ).strip())
    except (RuntimeError, ValueError):
        page_count = 0

    try:
        word_count = int(_run_applescript(
            f'tell application "Pages" to get count of words of body text of {target}'
        ).strip())
    except (RuntimeError, ValueError):
        word_count = 0

    try:
        char_count = int(_run_applescript(
            f'tell application "Pages" to get count of characters of body text of {target}'
        ).strip())
    except (RuntimeError, ValueError):
        char_count = 0

    return {
        "name": doc_name,
        "path": doc_path,
        "modified": modified,
        "page_count": page_count,
        "word_count": word_count,
        "character_count": char_count,
    }


def list_documents() -> list[dict]:
    """List all open Pages documents."""
    ensure_pages_running()
    count = int(_run_applescript(
        'tell application "Pages" to get count of documents'
    ).strip())

    if count == 0:
        return []

    documents = []
    for i in range(1, count + 1):
        try:
            doc_name = _run_applescript(
                f'tell application "Pages" to get name of document {i}'
            ).strip()
            try:
                doc_path = _run_applescript(
                    f'tell application "Pages" to get file of document {i} as text'
                ).strip()
            except RuntimeError:
                doc_path = ""
            try:
                modified = _run_applescript(
                    f'tell application "Pages" to get modified of document {i}'
                ).strip().lower() == "true"
            except RuntimeError:
                modified = False
            documents.append({
                "name": doc_name,
                "path": doc_path,
                "modified": modified,
            })
        except RuntimeError:
            continue
    return documents


# ---------------------------------------------------------------------------
# Password management
# ---------------------------------------------------------------------------

def set_password(
    password: str,
    hint: str | None = None,
    document: str | None = None,
) -> None:
    """Set a password on a document.

    AppleScript: set password "pw" to <document> [hint "hint"]
    """
    ensure_pages_running()
    target = _doc_target(document)
    hint_clause = ""
    if hint:
        hint_clause = f' hint "{_esc(hint)}"'
    script = (
        f'tell application "Pages"\n'
        f'  set password "{_esc(password)}" to {target}{hint_clause}\n'
        f'end tell'
    )
    _run_applescript(script)


def remove_password(
    password: str,
    document: str | None = None,
) -> None:
    """Remove the password from a document.

    AppleScript: remove password "pw" from <document>
    """
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  remove password "{_esc(password)}" from {target}\n'
        f'end tell'
    )
    _run_applescript(script)


# ---------------------------------------------------------------------------
# Placeholder text
# ---------------------------------------------------------------------------

def list_placeholders(document: str | None = None) -> list[dict]:
    """List all placeholder texts in the document.

    Returns:
        list[dict]: Each with keys: index, tag.
    """
    ensure_pages_running()
    target = _doc_target(document)

    count_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of placeholder texts\n'
        f'  end tell\n'
        f'end tell'
    )
    count = int(_run_applescript(count_script).strip())

    if count == 0:
        return []

    placeholders = []
    for i in range(1, count + 1):
        try:
            tag_script = (
                f'tell application "Pages"\n'
                f'  tell {target}\n'
                f'    get tag of placeholder text {i}\n'
                f'  end tell\n'
                f'end tell'
            )
            tag = _run_applescript(tag_script).strip()
            placeholders.append({"index": i, "tag": tag})
        except RuntimeError:
            placeholders.append({"index": i, "tag": ""})
    return placeholders


def set_placeholder_tag(
    index: int,
    tag: str,
    document: str | None = None,
) -> None:
    """Set the tag value of a placeholder text by index."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set tag of placeholder text {index} to "{_esc(tag)}"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_placeholder_tag(
    index: int,
    document: str | None = None,
) -> str:
    """Get the tag value of a placeholder text by index."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get tag of placeholder text {index}\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------

def get_section_count(document: str | None = None) -> int:
    """Get the number of sections in the document."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of sections\n'
        f'  end tell\n'
        f'end tell'
    )
    return int(_run_applescript(script).strip())


def get_section_body_text(
    section_index: int,
    document: str | None = None,
) -> str:
    """Get the body text of a specific section."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get body text of section {section_index}\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script)


# ---------------------------------------------------------------------------
# Page body text
# ---------------------------------------------------------------------------

def get_page_body_text(
    page_index: int = 1,
    document: str | None = None,
) -> str:
    """Get the body text of a specific page."""
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    get body text of page {page_index}\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script)


# ---------------------------------------------------------------------------
# Generic delete command
# ---------------------------------------------------------------------------

def delete_object(
    object_specifier: str,
    document: str | None = None,
) -> None:
    """Delete any object in the document by AppleScript specifier.

    Example: delete_object("table 1 of page 1")
    """
    ensure_pages_running()
    target = _doc_target(document)
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    delete {object_specifier}\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)
