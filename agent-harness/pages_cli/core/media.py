"""Media (images, shapes) operations module for Apple Pages."""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running


def _doc_target(document: str | None) -> str:
    """Return the AppleScript target reference for a document."""
    if document:
        escaped = document.replace('"', '\\"')
        return f'document "{escaped}"'
    return "front document"


def add_image(
    file_path: str,
    x: int = 0,
    y: int = 0,
    width: int | None = None,
    height: int | None = None,
    document: str | None = None,
) -> dict:
    """Add an image to the document.

    Args:
        file_path: Path to the image file.
        x: Horizontal position in points. Defaults to 0.
        y: Vertical position in points. Defaults to 0.
        width: Optional width in points. Uses original size if None.
        height: Optional height in points. Uses original size if None.
        document: Optional document name.

    Returns:
        dict: Image info with keys: position, file.

    Raises:
        RuntimeError: If the image cannot be added.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_path = file_path.replace('"', '\\"')

    props = f"position:{{{x}, {y}}}"
    if width is not None:
        props += f", width:{width}"
    if height is not None:
        props += f", height:{height}"

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set imgFile to POSIX file "{escaped_path}"\n'
        f'    set newImage to make new image with properties {{file:imgFile, {props}}}\n'
        f'    return "ok"\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)

    return {
        "file": file_path,
        "position": {"x": x, "y": y},
        "width": width,
        "height": height,
    }


def add_shape(
    shape_type: str = "rectangle",
    x: int = 0,
    y: int = 0,
    width: int = 100,
    height: int = 100,
    text: str | None = None,
    document: str | None = None,
) -> dict:
    """Add a shape to the document.

    Supported shape types include: rectangle, rounded rectangle, oval, circle,
    diamond, triangle, right triangle, arrow, double arrow, star, plus, minus.

    Args:
        shape_type: Type of shape. Defaults to "rectangle".
        x: Horizontal position in points. Defaults to 0.
        y: Vertical position in points. Defaults to 0.
        width: Width in points. Defaults to 100.
        height: Height in points. Defaults to 100.
        text: Optional text to place inside the shape.
        document: Optional document name.

    Returns:
        dict: Shape info with keys: type, position, size.

    Raises:
        RuntimeError: If the shape cannot be added.
    """
    ensure_pages_running()
    target = _doc_target(document)

    # Map common names to Pages iWork item types
    shape_map = {
        "rectangle": "iWork item type rectangle",
        "rounded rectangle": "iWork item type rounded rectangle",
        "oval": "iWork item type oval",
        "circle": "iWork item type oval",
        "diamond": "iWork item type diamond",
        "triangle": "iWork item type triangle",
        "right triangle": "iWork item type right triangle",
        "arrow": "iWork item type arrow",
        "double arrow": "iWork item type double arrow",
        "star": "iWork item type star",
        "plus": "iWork item type plus",
        "minus": "iWork item type minus",
    }

    iwork_type = shape_map.get(shape_type.lower(), f"iWork item type {shape_type}")

    lines = [
        f'tell application "Pages"',
        f'  tell {target}',
        f'    set newShape to make new shape with properties '
        f'{{position:{{{x}, {y}}}, width:{width}, height:{height}}}',
    ]

    if text:
        escaped_text = text.replace("\\", "\\\\").replace('"', '\\"')
        lines.append(
            f'    set object text of newShape to "{escaped_text}"'
        )

    lines.extend([
        f'    return "ok"',
        f'  end tell',
        f'end tell',
    ])

    script = "\n".join(lines)
    _run_applescript(script)

    return {
        "type": shape_type,
        "position": {"x": x, "y": y},
        "size": {"width": width, "height": height},
        "text": text,
    }


def list_images(document: str | None = None) -> list[dict]:
    """List all images in the document.

    Args:
        document: Optional document name.

    Returns:
        list[dict]: List of image info dicts with keys: index, position, width, height.

    Raises:
        RuntimeError: If images cannot be listed.
    """
    ensure_pages_running()
    target = _doc_target(document)

    count_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of images\n'
        f'  end tell\n'
        f'end tell'
    )
    count = int(_run_applescript(count_script).strip())

    if count == 0:
        return []

    images = []
    for i in range(1, count + 1):
        info_script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            f'    set img to image {i}\n'
            f'    set imgPos to position of img\n'
            f'    set imgW to width of img\n'
            f'    set imgH to height of img\n'
            f'    return (item 1 of imgPos as text) & "|" & '
            f'(item 2 of imgPos as text) & "|" & '
            f'(imgW as text) & "|" & (imgH as text)\n'
            f'  end tell\n'
            f'end tell'
        )
        try:
            result = _run_applescript(info_script).strip()
            parts = result.split("|")
            images.append({
                "index": i,
                "position": {
                    "x": float(parts[0]) if len(parts) > 0 else 0,
                    "y": float(parts[1]) if len(parts) > 1 else 0,
                },
                "width": float(parts[2]) if len(parts) > 2 else 0,
                "height": float(parts[3]) if len(parts) > 3 else 0,
            })
        except (RuntimeError, ValueError, IndexError):
            images.append({"index": i, "position": {"x": 0, "y": 0}, "width": 0, "height": 0})

    return images


def list_shapes(document: str | None = None) -> list[dict]:
    """List all shapes in the document.

    Args:
        document: Optional document name.

    Returns:
        list[dict]: List of shape info dicts with keys: index, position, width, height, text.

    Raises:
        RuntimeError: If shapes cannot be listed.
    """
    ensure_pages_running()
    target = _doc_target(document)

    count_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of shapes\n'
        f'  end tell\n'
        f'end tell'
    )
    count = int(_run_applescript(count_script).strip())

    if count == 0:
        return []

    shapes = []
    for i in range(1, count + 1):
        info_script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            f'    set s to shape {i}\n'
            f'    set sPos to position of s\n'
            f'    set sW to width of s\n'
            f'    set sH to height of s\n'
            f'    try\n'
            f'      set sText to object text of s\n'
            f'    on error\n'
            f'      set sText to ""\n'
            f'    end try\n'
            f'    return (item 1 of sPos as text) & "|" & '
            f'(item 2 of sPos as text) & "|" & '
            f'(sW as text) & "|" & (sH as text) & "|" & sText\n'
            f'  end tell\n'
            f'end tell'
        )
        try:
            result = _run_applescript(info_script).strip()
            parts = result.split("|")
            shapes.append({
                "index": i,
                "position": {
                    "x": float(parts[0]) if len(parts) > 0 else 0,
                    "y": float(parts[1]) if len(parts) > 1 else 0,
                },
                "width": float(parts[2]) if len(parts) > 2 else 0,
                "height": float(parts[3]) if len(parts) > 3 else 0,
                "text": parts[4] if len(parts) > 4 else "",
            })
        except (RuntimeError, ValueError, IndexError):
            shapes.append({
                "index": i,
                "position": {"x": 0, "y": 0},
                "width": 0,
                "height": 0,
                "text": "",
            })

    return shapes
