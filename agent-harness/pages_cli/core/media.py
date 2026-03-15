"""Media operations module for Apple Pages.

CRITICAL: All objects (images, shapes, text items, audio clips, movies,
lines, groups) must be created and accessed via ``tell page 1`` inside
the document tell block.

Covers the full sdef API for: shape, image, text item, audio clip, movie,
line, group, and the base iWork item properties (height, width, locked,
parent, position).
"""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running
from pages_cli.utils.helpers import _doc_target, _esc


def _page_tell(document: str | None) -> tuple[str, str]:
    """Return (open_block, close_block) for page 1 tell chain."""
    target = _doc_target(document)
    open_block = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell page 1\n'
    )
    close_block = (
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    return open_block, close_block


# ---------------------------------------------------------------------------
# Generic property getters / setters (work for any item type)
# ---------------------------------------------------------------------------

def set_item_property(
    item_type: str,
    index: int,
    prop_name: str,
    prop_value: str,
    document: str | None = None,
) -> None:
    """Set an arbitrary property on an item.

    prop_value is inserted verbatim into the AppleScript, so callers
    must format it correctly (e.g. quoted strings, numeric values, etc.).
    """
    ensure_pages_running()
    o, c = _page_tell(document)
    script = (
        f'{o}'
        f'      set {prop_name} of {item_type} {index} to {prop_value}\n'
        f'{c}'
    )
    _run_applescript(script)


def get_item_property(
    item_type: str,
    index: int,
    prop_name: str,
    document: str | None = None,
) -> str:
    """Get an arbitrary property of an item. Returns the raw string."""
    ensure_pages_running()
    o, c = _page_tell(document)
    script = (
        f'{o}'
        f'      get {prop_name} of {item_type} {index}\n'
        f'{c}'
    )
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------

def add_image(
    file_path: str,
    x: int = 0,
    y: int = 0,
    width: int | None = None,
    height: int | None = None,
    description: str | None = None,
    document: str | None = None,
) -> dict:
    """Add an image to page 1.

    Args:
        description: VoiceOver alt text (accessibility).
    """
    ensure_pages_running()
    o, c = _page_tell(document)
    escaped_path = _esc(file_path)

    props = f"position:{{{x}, {y}}}"
    if width is not None:
        props += f", width:{width}"
    if height is not None:
        props += f", height:{height}"

    lines = [
        o,
        f'      set imgFile to POSIX file "{escaped_path}"',
        f'      set newImage to make new image with properties {{file:imgFile, {props}}}',
    ]
    if description is not None:
        lines.append(f'      set description of newImage to "{_esc(description)}"')
    lines.append('      return "ok"')
    lines.append(c)

    _run_applescript("\n".join(lines))
    return {
        "file": file_path,
        "position": {"x": x, "y": y},
        "width": width,
        "height": height,
        "description": description,
    }


def set_image_description(
    index: int, description: str, document: str | None = None,
) -> None:
    """Set the VoiceOver description (alt text) of an image."""
    set_item_property("image", index, "description", f'"{_esc(description)}"', document)


def get_image_description(
    index: int, document: str | None = None,
) -> str:
    """Get the VoiceOver description of an image."""
    return get_item_property("image", index, "description", document)


def get_image_file(
    index: int, document: str | None = None,
) -> str:
    """Get the file reference of an image (read-only)."""
    try:
        return get_item_property("image", index, "file", document)
    except RuntimeError:
        return ""


def set_image_file_name(
    index: int, file_name: str, document: str | None = None,
) -> None:
    """Set the file name of an image."""
    set_item_property("image", index, "file name", f'"{_esc(file_name)}"', document)


def get_image_file_name(
    index: int, document: str | None = None,
) -> str:
    """Get the file name of an image."""
    try:
        return get_item_property("image", index, "file name", document)
    except RuntimeError:
        return ""


# ---------------------------------------------------------------------------
# Shapes
# ---------------------------------------------------------------------------

def add_shape(
    shape_type: str = "rectangle",
    x: int = 0,
    y: int = 0,
    width: int = 100,
    height: int = 100,
    text: str | None = None,
    document: str | None = None,
) -> dict:
    """Add a shape to page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)

    lines = [
        o,
        f'      set newShape to make new shape with properties '
        f'{{position:{{{x}, {y}}}, width:{width}, height:{height}}}',
    ]
    if text:
        lines.append(f'      set object text of newShape to "{_esc(text)}"')
    lines.append('      return "ok"')
    lines.append(c)
    _run_applescript("\n".join(lines))

    return {
        "type": shape_type,
        "position": {"x": x, "y": y},
        "size": {"width": width, "height": height},
        "text": text,
    }


def get_shape_background_fill_type(
    index: int, document: str | None = None,
) -> str:
    """Get the background fill type of a shape (read-only)."""
    try:
        return get_item_property("shape", index, "background fill type", document)
    except RuntimeError:
        return ""


def set_shape_object_text(
    index: int, text: str, document: str | None = None,
) -> None:
    """Set the object text of a shape."""
    set_item_property("shape", index, "object text", f'"{_esc(text)}"', document)


def get_shape_object_text(
    index: int, document: str | None = None,
) -> str:
    """Get the object text of a shape."""
    try:
        return get_item_property("shape", index, "object text", document)
    except RuntimeError:
        return ""


# ---------------------------------------------------------------------------
# Text Items (text boxes)
# ---------------------------------------------------------------------------

def add_text_item(
    text: str,
    x: int = 0,
    y: int = 0,
    width: int = 300,
    height: int = 60,
    document: str | None = None,
) -> dict:
    """Add a text item (text box) to page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)
    script = (
        f'{o}'
        f'      set newTI to make new text item with properties '
        f'{{position:{{{x}, {y}}}, width:{width}, height:{height}}}\n'
        f'      set object text of newTI to "{_esc(text)}"\n'
        f'      return "ok"\n'
        f'{c}'
    )
    _run_applescript(script)
    return {
        "position": {"x": x, "y": y},
        "size": {"width": width, "height": height},
        "text": text,
    }


def get_text_item_background_fill_type(
    index: int, document: str | None = None,
) -> str:
    """Get the background fill type of a text item (read-only)."""
    try:
        return get_item_property("text item", index, "background fill type", document)
    except RuntimeError:
        return ""


def set_text_item_object_text(
    index: int, text: str, document: str | None = None,
) -> None:
    """Set the object text of a text item."""
    set_item_property("text item", index, "object text", f'"{_esc(text)}"', document)


def get_text_item_object_text(
    index: int, document: str | None = None,
) -> str:
    """Get the object text of a text item."""
    try:
        return get_item_property("text item", index, "object text", document)
    except RuntimeError:
        return ""


# ---------------------------------------------------------------------------
# Audio Clips
# ---------------------------------------------------------------------------

def add_audio_clip(
    file_path: str,
    x: int = 0,
    y: int = 0,
    document: str | None = None,
) -> dict:
    """Add an audio clip to page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)
    script = (
        f'{o}'
        f'      set audioFile to POSIX file "{_esc(file_path)}"\n'
        f'      set newAudio to make new audio clip with properties '
        f'{{file:audioFile, position:{{{x}, {y}}}}}\n'
        f'      return "ok"\n'
        f'{c}'
    )
    _run_applescript(script)
    return {"file": file_path, "position": {"x": x, "y": y}}


def set_audio_clip_volume(
    index: int, volume: int, document: str | None = None,
) -> None:
    """Set audio clip volume (0-100)."""
    set_item_property("audio clip", index, "clip volume", str(volume), document)


def get_audio_clip_volume(
    index: int, document: str | None = None,
) -> int:
    """Get audio clip volume."""
    return int(get_item_property("audio clip", index, "clip volume", document))


def set_audio_file_name(
    index: int, file_name: str, document: str | None = None,
) -> None:
    """Set file name of an audio clip."""
    set_item_property("audio clip", index, "file name", f'"{_esc(file_name)}"', document)


def get_audio_file_name(
    index: int, document: str | None = None,
) -> str:
    """Get file name of an audio clip."""
    try:
        return get_item_property("audio clip", index, "file name", document)
    except RuntimeError:
        return ""


def set_audio_repetition_method(
    index: int, method: str, document: str | None = None,
) -> None:
    """Set repetition method: none, loop, loop back and forth."""
    set_item_property("audio clip", index, "repetition method", method, document)


def get_audio_repetition_method(
    index: int, document: str | None = None,
) -> str:
    """Get repetition method of an audio clip."""
    return get_item_property("audio clip", index, "repetition method", document)


# ---------------------------------------------------------------------------
# Movies
# ---------------------------------------------------------------------------

def add_movie(
    file_path: str,
    x: int = 0,
    y: int = 0,
    width: int | None = None,
    height: int | None = None,
    document: str | None = None,
) -> dict:
    """Add a movie to page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)
    escaped_path = _esc(file_path)
    props = f"position:{{{x}, {y}}}"
    if width is not None:
        props += f", width:{width}"
    if height is not None:
        props += f", height:{height}"
    script = (
        f'{o}'
        f'      set movieFile to POSIX file "{escaped_path}"\n'
        f'      set newMovie to make new movie with properties '
        f'{{file:movieFile, {props}}}\n'
        f'      return "ok"\n'
        f'{c}'
    )
    _run_applescript(script)
    return {"file": file_path, "position": {"x": x, "y": y}}


def set_movie_volume(
    index: int, volume: int, document: str | None = None,
) -> None:
    """Set movie volume (0-100)."""
    set_item_property("movie", index, "movie volume", str(volume), document)


def get_movie_volume(
    index: int, document: str | None = None,
) -> int:
    """Get movie volume."""
    return int(get_item_property("movie", index, "movie volume", document))


def set_movie_file_name(
    index: int, file_name: str, document: str | None = None,
) -> None:
    """Set file name of a movie."""
    set_item_property("movie", index, "file name", f'"{_esc(file_name)}"', document)


def get_movie_file_name(
    index: int, document: str | None = None,
) -> str:
    """Get file name of a movie."""
    try:
        return get_item_property("movie", index, "file name", document)
    except RuntimeError:
        return ""


def set_movie_repetition_method(
    index: int, method: str, document: str | None = None,
) -> None:
    """Set repetition method: none, loop, loop back and forth."""
    set_item_property("movie", index, "repetition method", method, document)


def get_movie_repetition_method(
    index: int, document: str | None = None,
) -> str:
    """Get repetition method of a movie."""
    return get_item_property("movie", index, "repetition method", document)


# ---------------------------------------------------------------------------
# Lines
# ---------------------------------------------------------------------------

def add_line(
    start_x: int, start_y: int,
    end_x: int, end_y: int,
    document: str | None = None,
) -> dict:
    """Add a line to page 1.

    NOTE: Line creation via AppleScript may fail in some Pages versions.
    """
    ensure_pages_running()
    o, c = _page_tell(document)
    script = (
        f'{o}'
        f'      set newLine to make new line with properties '
        f'{{start point:{{{start_x}, {start_y}}}, end point:{{{end_x}, {end_y}}}}}\n'
        f'      return "ok"\n'
        f'{c}'
    )
    _run_applescript(script)
    return {
        "start_point": {"x": start_x, "y": start_y},
        "end_point": {"x": end_x, "y": end_y},
    }


def set_line_start_point(
    index: int, x: int, y: int, document: str | None = None,
) -> None:
    """Set start point of a line."""
    set_item_property("line", index, "start point", f"{{{x}, {y}}}", document)


def get_line_start_point(
    index: int, document: str | None = None,
) -> str:
    """Get start point of a line."""
    return get_item_property("line", index, "start point", document)


def set_line_end_point(
    index: int, x: int, y: int, document: str | None = None,
) -> None:
    """Set end point of a line."""
    set_item_property("line", index, "end point", f"{{{x}, {y}}}", document)


def get_line_end_point(
    index: int, document: str | None = None,
) -> str:
    """Get end point of a line."""
    return get_item_property("line", index, "end point", document)


# ---------------------------------------------------------------------------
# Common iWork item properties: rotation, opacity, reflection, locked, position
# These work on any item_type: image, shape, text item, movie, line, group
# ---------------------------------------------------------------------------

def set_rotation(
    item_type: str, index: int, degrees: float,
    document: str | None = None,
) -> None:
    """Set rotation of an item (0-359)."""
    set_item_property(item_type, index, "rotation", str(degrees), document)


def get_rotation(
    item_type: str, index: int, document: str | None = None,
) -> float:
    """Get rotation of an item."""
    return float(get_item_property(item_type, index, "rotation", document))


def set_opacity(
    item_type: str, index: int, opacity: int,
    document: str | None = None,
) -> None:
    """Set opacity of an item (0-100)."""
    set_item_property(item_type, index, "opacity", str(opacity), document)


def get_opacity(
    item_type: str, index: int, document: str | None = None,
) -> int:
    """Get opacity of an item."""
    return int(get_item_property(item_type, index, "opacity", document))


def set_reflection_showing(
    item_type: str, index: int, showing: bool,
    document: str | None = None,
) -> None:
    """Set whether reflection is showing."""
    val = "true" if showing else "false"
    set_item_property(item_type, index, "reflection showing", val, document)


def get_reflection_showing(
    item_type: str, index: int, document: str | None = None,
) -> bool:
    """Get whether reflection is showing."""
    result = get_item_property(item_type, index, "reflection showing", document)
    return result.lower() == "true"


def set_reflection_value(
    item_type: str, index: int, value: int,
    document: str | None = None,
) -> None:
    """Set reflection value (0-100)."""
    set_item_property(item_type, index, "reflection value", str(value), document)


def get_reflection_value(
    item_type: str, index: int, document: str | None = None,
) -> int:
    """Get reflection value."""
    return int(get_item_property(item_type, index, "reflection value", document))


def set_locked(
    item_type: str, index: int, locked: bool,
    document: str | None = None,
) -> None:
    """Set locked state of an item."""
    val = "true" if locked else "false"
    set_item_property(item_type, index, "locked", val, document)


def get_locked(
    item_type: str, index: int, document: str | None = None,
) -> bool:
    """Get locked state of an item."""
    result = get_item_property(item_type, index, "locked", document)
    return result.lower() == "true"


def set_position(
    item_type: str, index: int, x: int, y: int,
    document: str | None = None,
) -> None:
    """Set position of an item."""
    set_item_property(item_type, index, "position", f"{{{x}, {y}}}", document)


def get_position(
    item_type: str, index: int, document: str | None = None,
) -> str:
    """Get position of an item."""
    return get_item_property(item_type, index, "position", document)


def set_width(
    item_type: str, index: int, width: float,
    document: str | None = None,
) -> None:
    """Set width of an item."""
    set_item_property(item_type, index, "width", str(width), document)


def get_width(
    item_type: str, index: int, document: str | None = None,
) -> float:
    """Get width of an item."""
    return float(get_item_property(item_type, index, "width", document))


def set_height(
    item_type: str, index: int, height: float,
    document: str | None = None,
) -> None:
    """Set height of an item."""
    set_item_property(item_type, index, "height", str(height), document)


def get_height(
    item_type: str, index: int, document: str | None = None,
) -> float:
    """Get height of an item."""
    return float(get_item_property(item_type, index, "height", document))


def get_parent(
    item_type: str, index: int, document: str | None = None,
) -> str:
    """Get the parent of an item (read-only)."""
    try:
        return get_item_property(item_type, index, "parent", document)
    except RuntimeError:
        return ""


# ---------------------------------------------------------------------------
# Delete
# ---------------------------------------------------------------------------

def delete_item(
    item_type: str,
    index: int | None = None,
    name: str | None = None,
    document: str | None = None,
) -> None:
    """Delete an item from page 1 by index or name."""
    if index is None and name is None:
        raise ValueError("Provide either index or name.")
    ensure_pages_running()
    o, c = _page_tell(document)
    if name:
        item_ref = f'{item_type} "{_esc(name)}"'
    else:
        item_ref = f"{item_type} {index}"
    script = f'{o}      delete {item_ref}\n{c}'
    _run_applescript(script)


# ---------------------------------------------------------------------------
# Full property inspection
# ---------------------------------------------------------------------------

def get_item_properties(
    item_type: str,
    index: int,
    document: str | None = None,
) -> dict:
    """Get common properties of an item on page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)

    script = (
        f'{o}'
        f'      set itm to {item_type} {index}\n'
        f'      set itmPos to position of itm\n'
        f'      set itmW to width of itm\n'
        f'      set itmH to height of itm\n'
        f'      set itmRot to rotation of itm\n'
        f'      set itmOp to opacity of itm\n'
        f'      try\n'
        f'        set itmRef to reflection showing of itm\n'
        f'      on error\n'
        f'        set itmRef to false\n'
        f'      end try\n'
        f'      try\n'
        f'        set itmRefVal to reflection value of itm\n'
        f'      on error\n'
        f'        set itmRefVal to 0\n'
        f'      end try\n'
        f'      try\n'
        f'        set itmLock to locked of itm\n'
        f'      on error\n'
        f'        set itmLock to false\n'
        f'      end try\n'
        f'      return (item 1 of itmPos as text) & "|" & '
        f'(item 2 of itmPos as text) & "|" & '
        f'(itmW as text) & "|" & (itmH as text) & "|" & '
        f'(itmRot as text) & "|" & (itmOp as text) & "|" & '
        f'(itmRef as text) & "|" & (itmRefVal as text) & "|" & '
        f'(itmLock as text)\n'
        f'{c}'
    )
    result = _run_applescript(script).strip()
    parts = result.split("|")

    props = {
        "type": item_type,
        "index": index,
        "position": {
            "x": float(parts[0]) if len(parts) > 0 else 0,
            "y": float(parts[1]) if len(parts) > 1 else 0,
        },
        "width": float(parts[2]) if len(parts) > 2 else 0,
        "height": float(parts[3]) if len(parts) > 3 else 0,
        "rotation": float(parts[4]) if len(parts) > 4 else 0,
        "opacity": float(parts[5]) if len(parts) > 5 else 100,
        "reflection_showing": parts[6].lower() == "true" if len(parts) > 6 else False,
        "reflection_value": float(parts[7]) if len(parts) > 7 else 0,
        "locked": parts[8].lower() == "true" if len(parts) > 8 else False,
    }

    # Type-specific additional properties
    if item_type in ("shape", "text item"):
        try:
            text_script = (
                f'{o}      get object text of {item_type} {index}\n{c}'
            )
            props["object_text"] = _run_applescript(text_script).strip()
        except RuntimeError:
            props["object_text"] = ""
        try:
            bf_script = (
                f'{o}      get background fill type of {item_type} {index}\n{c}'
            )
            props["background_fill_type"] = _run_applescript(bf_script).strip()
        except RuntimeError:
            props["background_fill_type"] = ""

    if item_type == "image":
        try:
            fn_script = f'{o}      get file name of image {index}\n{c}'
            props["file_name"] = _run_applescript(fn_script).strip()
        except RuntimeError:
            props["file_name"] = ""
        try:
            desc_script = f'{o}      get description of image {index}\n{c}'
            props["description"] = _run_applescript(desc_script).strip()
        except RuntimeError:
            props["description"] = ""

    if item_type == "audio clip":
        try:
            vol_script = f'{o}      get clip volume of audio clip {index}\n{c}'
            props["clip_volume"] = int(_run_applescript(vol_script).strip())
        except (RuntimeError, ValueError):
            props["clip_volume"] = 0
        try:
            rep_script = f'{o}      get repetition method of audio clip {index}\n{c}'
            props["repetition_method"] = _run_applescript(rep_script).strip()
        except RuntimeError:
            props["repetition_method"] = ""

    if item_type == "movie":
        try:
            vol_script = f'{o}      get movie volume of movie {index}\n{c}'
            props["movie_volume"] = int(_run_applescript(vol_script).strip())
        except (RuntimeError, ValueError):
            props["movie_volume"] = 0
        try:
            rep_script = f'{o}      get repetition method of movie {index}\n{c}'
            props["repetition_method"] = _run_applescript(rep_script).strip()
        except RuntimeError:
            props["repetition_method"] = ""

    if item_type == "line":
        try:
            sp_script = f'{o}      get start point of line {index}\n{c}'
            props["start_point"] = _run_applescript(sp_script).strip()
        except RuntimeError:
            props["start_point"] = ""
        try:
            ep_script = f'{o}      get end point of line {index}\n{c}'
            props["end_point"] = _run_applescript(ep_script).strip()
        except RuntimeError:
            props["end_point"] = ""

    return props


# ---------------------------------------------------------------------------
# Listing helpers
# ---------------------------------------------------------------------------

def _count_items(item_type: str, document: str | None = None) -> int:
    """Count items of a given type on page 1."""
    ensure_pages_running()
    o, c = _page_tell(document)
    plural = item_type + "s" if not item_type.endswith("s") else item_type
    # Handle special plurals
    if item_type == "audio clip":
        plural = "audio clips"
    elif item_type == "movie":
        plural = "movies"
    elif item_type == "text item":
        plural = "text items"

    script = f'{o}      count of {plural}\n{c}'
    return int(_run_applescript(script).strip())


def list_images(document: str | None = None) -> list[dict]:
    """List all images on page 1."""
    count = _count_items("image", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("image", i, document)
        for i in range(1, count + 1)
    ]


def list_shapes(document: str | None = None) -> list[dict]:
    """List all shapes on page 1."""
    count = _count_items("shape", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("shape", i, document)
        for i in range(1, count + 1)
    ]


def list_text_items(document: str | None = None) -> list[dict]:
    """List all text items on page 1."""
    count = _count_items("text item", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("text item", i, document)
        for i in range(1, count + 1)
    ]


def list_audio_clips(document: str | None = None) -> list[dict]:
    """List all audio clips on page 1."""
    count = _count_items("audio clip", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("audio clip", i, document)
        for i in range(1, count + 1)
    ]


def list_movies(document: str | None = None) -> list[dict]:
    """List all movies on page 1."""
    count = _count_items("movie", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("movie", i, document)
        for i in range(1, count + 1)
    ]


def list_lines(document: str | None = None) -> list[dict]:
    """List all lines on page 1."""
    count = _count_items("line", document)
    if count == 0:
        return []
    return [
        _safe_get_item_props("line", i, document)
        for i in range(1, count + 1)
    ]


def list_groups(document: str | None = None) -> list[dict]:
    """List all groups on page 1."""
    count = _count_items("group", document)
    if count == 0:
        return []
    results = []
    for i in range(1, count + 1):
        try:
            props = get_item_properties("group", i, document)
            results.append(props)
        except RuntimeError:
            results.append({"type": "group", "index": i})
    return results


def list_all_items(document: str | None = None) -> dict:
    """List ALL items on page 1."""
    images = list_images(document)
    shapes = list_shapes(document)
    text_items = list_text_items(document)
    audio_clips = list_audio_clips(document)
    movies = list_movies(document)
    lines = list_lines(document)
    groups = list_groups(document)

    from pages_cli.core.tables import list_tables
    tables = list_tables(document)

    total = (
        len(images) + len(shapes) + len(text_items) + len(audio_clips)
        + len(movies) + len(lines) + len(groups) + len(tables)
    )

    return {
        "images": images,
        "shapes": shapes,
        "text_items": text_items,
        "audio_clips": audio_clips,
        "movies": movies,
        "lines": lines,
        "groups": groups,
        "tables": tables,
        "total": total,
    }


def _safe_get_item_props(
    item_type: str, index: int, document: str | None = None,
) -> dict:
    """Get item properties, returning a stub on error."""
    try:
        return get_item_properties(item_type, index, document)
    except RuntimeError:
        return {
            "type": item_type,
            "index": index,
            "position": {"x": 0, "y": 0},
            "width": 0,
            "height": 0,
        }
