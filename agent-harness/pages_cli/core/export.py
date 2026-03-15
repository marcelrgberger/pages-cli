"""Export module for Apple Pages."""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running


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


def _doc_target(document: str | None) -> str:
    """Return the AppleScript target reference for a document."""
    if document:
        escaped = document.replace('"', '\\"')
        return f'document "{escaped}"'
    return "front document"


def export_document(
    output_path: str,
    format: str = "PDF",
    document: str | None = None,
    **options,
) -> dict:
    """Export the document to the specified format.

    Args:
        output_path: The output file path.
        format: Export format. Supported values:
            "PDF", "Microsoft Word" / "word" / "docx",
            "EPUB" / "epub",
            "unformatted text" / "text" / "txt",
            "formatted text" / "rtf",
            "Pages 09" / "pages09".
        document: Optional document name.
        **options: Additional export options:
            - password (str): Password protect the export.
            - password_hint (str): Password hint.
            - image_quality (str): Image quality ("Good", "Better", "Best").
            - include_comments (bool): Include comments in export.
            - include_annotations (bool): Include annotations.
            For EPUB exports:
            - title (str): EPUB title.
            - author (str): EPUB author.
            - genre (str): EPUB genre.
            - language (str): EPUB language.
            - publisher (str): EPUB publisher.
            - cover (str): Path to cover image.
            - fixed_layout (bool): Use fixed layout for EPUB.

    Returns:
        dict: Export result with keys: path, format.

    Raises:
        RuntimeError: If the export fails.
        ValueError: If an unsupported format is provided.
    """
    ensure_pages_running()
    target = _doc_target(document)

    resolved_format = _FORMAT_MAP.get(format)
    if resolved_format is None:
        raise ValueError(
            f"Unsupported export format '{format}'. "
            f"Supported formats: {', '.join(get_export_formats())}"
        )

    escaped_path = output_path.replace('"', '\\"')

    # Build the export properties
    export_props = []

    if options.get("password"):
        pw = options["password"].replace('"', '\\"')
        export_props.append(f'password:"{pw}"')

    if options.get("password_hint"):
        hint = options["password_hint"].replace('"', '\\"')
        export_props.append(f'password hint:"{hint}"')

    if options.get("image_quality"):
        quality = options["image_quality"].replace('"', '\\"')
        export_props.append(f'image quality:{quality}')

    if options.get("include_comments") is not None:
        val = "true" if options["include_comments"] else "false"
        export_props.append(f"include comments:{val}")

    if options.get("include_annotations") is not None:
        val = "true" if options["include_annotations"] else "false"
        export_props.append(f"include annotations:{val}")

    # EPUB-specific properties
    if resolved_format == "EPUB":
        if options.get("title"):
            title = options["title"].replace('"', '\\"')
            export_props.append(f'epub title:"{title}"')
        if options.get("author"):
            author = options["author"].replace('"', '\\"')
            export_props.append(f'epub author:"{author}"')
        if options.get("genre"):
            genre = options["genre"].replace('"', '\\"')
            export_props.append(f'epub genre:"{genre}"')
        if options.get("language"):
            lang = options["language"].replace('"', '\\"')
            export_props.append(f'epub language:"{lang}"')
        if options.get("publisher"):
            pub = options["publisher"].replace('"', '\\"')
            export_props.append(f'epub publisher:"{pub}"')
        if options.get("cover"):
            cover = options["cover"].replace('"', '\\"')
            export_props.append(f'epub cover:POSIX file "{cover}"')
        if options.get("fixed_layout") is not None:
            val = "true" if options["fixed_layout"] else "false"
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
    """Return a list of supported export format names.

    Returns:
        list[str]: Supported export formats.
    """
    return [
        "PDF",
        "Microsoft Word",
        "EPUB",
        "unformatted text",
        "formatted text",
        "Pages 09",
    ]
