"""Shared helper functions for Pages CLI core modules."""


def _doc_target(document: str | None) -> str:
    """Return the AppleScript target reference for a document."""
    if document:
        escaped = document.replace('\\', '\\\\').replace('"', '\\"')
        return f'document "{escaped}"'
    return "front document"


def _esc(value: str) -> str:
    """Escape a string for safe embedding in AppleScript."""
    return value.replace('\\', '\\\\').replace('"', '\\"')
