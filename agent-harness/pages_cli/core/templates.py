"""Template management module for Apple Pages."""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running


def list_templates() -> list[str]:
    """List all available Pages templates.

    Returns:
        list[str]: List of template names.

    Raises:
        RuntimeError: If templates cannot be listed.
    """
    ensure_pages_running()

    script = (
        'tell application "Pages"\n'
        '  set templateNames to name of every template\n'
        '  set AppleScript\'s text item delimiters to "||"\n'
        '  set nameStr to templateNames as text\n'
        '  set AppleScript\'s text item delimiters to ""\n'
        '  return nameStr\n'
        'end tell'
    )
    result = _run_applescript(script)

    if not result.strip():
        return []

    return [t.strip() for t in result.split("||") if t.strip()]


def get_template_info(name: str) -> dict:
    """Get details about a specific template.

    Args:
        name: The template name to look up.

    Returns:
        dict: Template info with keys: name, id.

    Raises:
        RuntimeError: If the template cannot be found or info cannot be retrieved.
    """
    ensure_pages_running()

    escaped_name = name.replace('"', '\\"')

    # Get template id
    id_script = (
        f'tell application "Pages"\n'
        f'  set t to first template whose name is "{escaped_name}"\n'
        f'  return id of t\n'
        f'end tell'
    )
    try:
        template_id = _run_applescript(id_script).strip()
    except RuntimeError:
        template_id = ""

    return {
        "name": name,
        "id": template_id,
    }
