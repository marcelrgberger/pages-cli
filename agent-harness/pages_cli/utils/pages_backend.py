"""Backend module that wraps osascript execution for Apple Pages."""

import shutil
import subprocess
from pathlib import Path


def find_pages() -> str:
    """Check if Apple Pages is installed and return the application path.

    Returns:
        str: The path to the Pages application.

    Raises:
        RuntimeError: If Pages is not installed or osascript is not available.
    """
    osascript = shutil.which("osascript")
    if not osascript:
        raise RuntimeError(
            "osascript not found. This tool requires macOS with osascript available."
        )

    # Check if Pages.app exists in standard locations
    pages_paths = [
        Path("/Applications/Pages.app"),
        Path("/Applications/Pages Creator Studio.app"),
        Path("/System/Applications/Pages.app"),
    ]
    for p in pages_paths:
        if p.exists():
            return str(p)

    # Try via AppleScript as a fallback
    try:
        result = _run_applescript(
            'tell application "Finder" to get application file id "com.apple.iWork.Pages" as text'
        )
        if result.strip():
            return result.strip()
    except RuntimeError:
        pass

    raise RuntimeError(
        "Apple Pages is not installed. Install it from the Mac App Store: "
        "https://apps.apple.com/app/pages/id409201541"
    )


def _run_applescript(script: str) -> str:
    """Run an AppleScript string via osascript.

    Args:
        script: The AppleScript code to execute.

    Returns:
        str: The stdout output from osascript.

    Raises:
        RuntimeError: If the script fails to execute.
    """
    osascript = shutil.which("osascript")
    if not osascript:
        raise RuntimeError(
            "osascript not found. This tool requires macOS."
        )

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"AppleScript timed out after 30 seconds. Script: {script[:200]}..."
        )

    if result.returncode != 0:
        error_msg = result.stderr.strip() if result.stderr else "Unknown error"
        raise RuntimeError(
            f"AppleScript execution failed (exit code {result.returncode}): {error_msg}"
        )

    return result.stdout.strip()


def _run_jxa(script: str) -> str:
    """Run JavaScript for Automation via osascript.

    Args:
        script: The JXA code to execute.

    Returns:
        str: The stdout output from osascript.

    Raises:
        RuntimeError: If the script fails to execute.
    """
    osascript = shutil.which("osascript")
    if not osascript:
        raise RuntimeError(
            "osascript not found. This tool requires macOS."
        )

    try:
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", script],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"JXA script timed out after 30 seconds. Script: {script[:200]}..."
        )

    if result.returncode != 0:
        error_msg = result.stderr.strip() if result.stderr else "Unknown error"
        raise RuntimeError(
            f"JXA execution failed (exit code {result.returncode}): {error_msg}"
        )

    return result.stdout.strip()


def ensure_pages_running() -> None:
    """Ensure Apple Pages is running, launching it if necessary.

    Raises:
        RuntimeError: If Pages cannot be launched.
    """
    find_pages()  # Verify Pages is installed

    if not is_pages_running():
        try:
            _run_applescript('tell application "Pages" to activate')
        except RuntimeError as e:
            raise RuntimeError(f"Failed to launch Apple Pages: {e}")


def is_pages_running() -> bool:
    """Check if Apple Pages is currently running.

    Returns:
        bool: True if Pages is running, False otherwise.
    """
    try:
        result = _run_applescript(
            'tell application "System Events" to (name of processes) contains "Pages"'
        )
        return result.strip().lower() == "true"
    except RuntimeError:
        return False


def quit_pages(saving: bool = True) -> None:
    """Quit Apple Pages.

    Args:
        saving: If True, save all documents before quitting. Defaults to True.

    Raises:
        RuntimeError: If Pages cannot be quit.
    """
    save_flag = "saving yes" if saving else "saving no"
    _run_applescript(f'tell application "Pages" to quit {save_flag}')
