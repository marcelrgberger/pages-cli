"""Document management module for Apple Pages."""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running


def create_document(template: str = "Blank", name: str | None = None) -> dict:
    """Create a new Pages document from a template.

    Args:
        template: Template name to use. Defaults to "Blank".
        name: Optional name for the document.

    Returns:
        dict: Document info with keys: name, modified, page_count.

    Raises:
        RuntimeError: If document creation fails.
    """
    ensure_pages_running()

    if template == "Blank":
        script = 'tell application "Pages" to make new document'
    else:
        escaped_template = template.replace('"', '\\"')
        script = (
            'tell application "Pages"\n'
            f'  make new document with properties {{document template:template "{escaped_template}"}}\n'
            'end tell'
        )

    _run_applescript(script)

    if name:
        escaped_name = name.replace('"', '\\"')
        _run_applescript(
            f'tell application "Pages" to set name of front document to "{escaped_name}"'
        )

    return get_document_info()


def open_document(path: str) -> dict:
    """Open an existing .pages file.

    Args:
        path: The file path to the .pages document.

    Returns:
        dict: Document info dict.

    Raises:
        RuntimeError: If the file cannot be opened.
    """
    ensure_pages_running()

    escaped_path = path.replace('"', '\\"')
    script = (
        f'tell application "Pages"\n'
        f'  open POSIX file "{escaped_path}"\n'
        f'end tell'
    )
    _run_applescript(script)
    return get_document_info()


def close_document(name: str | None = None, saving: bool = True) -> None:
    """Close a Pages document.

    Args:
        name: The name of the document to close. If None, closes the front document.
        saving: Whether to save before closing. Defaults to True.

    Raises:
        RuntimeError: If the document cannot be closed.
    """
    ensure_pages_running()

    save_flag = "saving yes" if saving else "saving no"

    if name:
        escaped_name = name.replace('"', '\\"')
        target = f'document "{escaped_name}"'
    else:
        target = "front document"

    script = f'tell application "Pages" to close {target} {save_flag}'
    _run_applescript(script)


def save_document(name: str | None = None, path: str | None = None) -> None:
    """Save the current or named document.

    Args:
        name: The name of the document to save. If None, saves the front document.
        path: Optional path to save to (Save As). If None, saves in place.

    Raises:
        RuntimeError: If the document cannot be saved.
    """
    ensure_pages_running()

    if name:
        escaped_name = name.replace('"', '\\"')
        target = f'document "{escaped_name}"'
    else:
        target = "front document"

    if path:
        escaped_path = path.replace('"', '\\"')
        script = (
            f'tell application "Pages"\n'
            f'  save {target} in POSIX file "{escaped_path}"\n'
            f'end tell'
        )
    else:
        script = f'tell application "Pages" to save {target}'

    _run_applescript(script)


def get_document_info(name: str | None = None) -> dict:
    """Get information about a Pages document.

    Args:
        name: The name of the document. If None, uses the front document.

    Returns:
        dict: Document info with keys: name, path, modified, page_count,
              word_count, character_count.

    Raises:
        RuntimeError: If the document info cannot be retrieved.
    """
    ensure_pages_running()

    if name:
        escaped_name = name.replace('"', '\\"')
        target = f'document "{escaped_name}"'
    else:
        target = "front document"

    # Get name
    doc_name = _run_applescript(
        f'tell application "Pages" to get name of {target}'
    )

    # Get file path (may not exist for unsaved documents)
    try:
        doc_path = _run_applescript(
            f'tell application "Pages" to get file of {target} as text'
        )
    except RuntimeError:
        doc_path = ""

    # Get modified state
    try:
        modified = _run_applescript(
            f'tell application "Pages" to get modified of {target}'
        )
        modified = modified.strip().lower() == "true"
    except RuntimeError:
        modified = False

    # Get page count
    try:
        page_count = _run_applescript(
            f'tell application "Pages" to get count of pages of {target}'
        )
        page_count = int(page_count.strip())
    except (RuntimeError, ValueError):
        page_count = 0

    # Get word count via body text
    try:
        word_count = _run_applescript(
            f'tell application "Pages" to get count of words of body text of {target}'
        )
        word_count = int(word_count.strip())
    except (RuntimeError, ValueError):
        word_count = 0

    # Get character count via body text
    try:
        char_count = _run_applescript(
            f'tell application "Pages" to get count of characters of body text of {target}'
        )
        char_count = int(char_count.strip())
    except (RuntimeError, ValueError):
        char_count = 0

    return {
        "name": doc_name.strip(),
        "path": doc_path.strip(),
        "modified": modified,
        "page_count": page_count,
        "word_count": word_count,
        "character_count": char_count,
    }


def list_documents() -> list[dict]:
    """List all open Pages documents.

    Returns:
        list[dict]: List of document info dicts with keys: name, path, modified.

    Raises:
        RuntimeError: If document listing fails.
    """
    ensure_pages_running()

    # Get the count of open documents
    count_str = _run_applescript(
        'tell application "Pages" to get count of documents'
    )
    count = int(count_str.strip())

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
