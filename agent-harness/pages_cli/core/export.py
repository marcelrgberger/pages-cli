"""Export module for Apple Pages.

Full coverage of all export formats and options including:
- PDF with password, password hint, image quality, comments, annotations
- EPUB with title, author, genre, language, publisher, cover, fixed layout
- Microsoft Word, plain text, RTF, Pages 09
"""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running
from pages_cli.utils.helpers import _doc_target, _esc


# Maps user-friendly format names to Pages AppleScript export format constants
_FORMAT_MAP = {
    "PDF": "PDF",
    "pdf": "PDF",
    "Microsoft Word": "Microsoft Word",
    "word": "Microsoft Word",
    "docx": "Microsoft Word",
    "EPUB": "EPUB",
    "epub": "EPUB",
    "unformatted text": "unformatted text",
    "text": "unformatted text",
    "txt": "unformatted text",
    "formatted text": "formatted text",
    "rtf": "formatted text",
    "Pages 09": "Pages 09",
    "pages09": "Pages 09",
}

# Valid image quality values
IMAGE_QUALITIES = ["Good", "Better", "Best"]


def export_document(
    output_path: str,
    format: str = "PDF",
    document: str | None = None,
    password: str | None = None,
    password_hint: str | None = None,
    image_quality: str | None = None,
    include_comments: bool | None = None,
    include_annotations: bool | None = None,
    # EPUB-specific
    epub_title: str | None = None,
    epub_author: str | None = None,
    epub_genre: str | None = None,
    epub_language: str | None = None,
    epub_publisher: str | None = None,
    epub_cover: str | None = None,
    epub_fixed_layout: bool | None = None,
) -> dict:
    """Export the document to the specified format with full option support.

    Args:
        output_path: The output file path.
        format: Export format name (see _FORMAT_MAP keys).
        document: Optional target document name.
        password: Password-protect the exported file.
        password_hint: Hint for the password.
        image_quality: "Good", "Better", or "Best".
        include_comments: Include comments in export.
        include_annotations: Include annotations in export.
        epub_title: EPUB title metadata.
        epub_author: EPUB author metadata.
        epub_genre: EPUB genre metadata.
        epub_language: EPUB language metadata.
        epub_publisher: EPUB publisher metadata.
        epub_cover: Path to cover image for EPUB.
        epub_fixed_layout: Use fixed layout for EPUB.

    Returns:
        dict with keys: path, format.
    """
    ensure_pages_running()
    target = _doc_target(document)

    resolved_format = _FORMAT_MAP.get(format)
    if resolved_format is None:
        raise ValueError(
            f"Unsupported export format '{format}'. "
            f"Supported: {', '.join(get_export_formats())}"
        )

    escaped_path = _esc(output_path)

    # Build export properties
    export_props = []

    if password:
        export_props.append(f'password:"{_esc(password)}"')
    if password_hint:
        export_props.append(f'password hint:"{_esc(password_hint)}"')
    if image_quality:
        export_props.append(f"image quality:{image_quality}")
    if include_comments is not None:
        val = "true" if include_comments else "false"
        export_props.append(f"include comments:{val}")
    if include_annotations is not None:
        val = "true" if include_annotations else "false"
        export_props.append(f"include annotations:{val}")

    # EPUB-specific
    if resolved_format == "EPUB":
        if epub_title:
            export_props.append(f'epub title:"{_esc(epub_title)}"')
        if epub_author:
            export_props.append(f'epub author:"{_esc(epub_author)}"')
        if epub_genre:
            export_props.append(f'epub genre:"{_esc(epub_genre)}"')
        if epub_language:
            export_props.append(f'epub language:"{_esc(epub_language)}"')
        if epub_publisher:
            export_props.append(f'epub publisher:"{_esc(epub_publisher)}"')
        if epub_cover:
            export_props.append(f'epub cover:POSIX file "{_esc(epub_cover)}"')
        if epub_fixed_layout is not None:
            val = "true" if epub_fixed_layout else "false"
            export_props.append(f"fixed layout:{val}")

    # Build the export command
    props_str = ""
    if export_props:
        props_str = " with properties {" + ", ".join(export_props) + "}"

    script = (
        f'tell application "Pages"\n'
        f'  export {target} to POSIX file "{escaped_path}" '
        f'as {resolved_format}{props_str}\n'
        f'end tell'
    )
    _run_applescript(script)

    return {
        "path": output_path,
        "format": resolved_format,
    }


def get_export_formats() -> list[str]:
    """Return a list of supported export format names."""
    return [
        "PDF",
        "Microsoft Word",
        "EPUB",
        "unformatted text",
        "formatted text",
        "Pages 09",
    ]
