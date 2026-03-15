"""Session state management for Apple Pages CLI harness."""

import json
from datetime import datetime, timezone
from pathlib import Path


class Session:
    """Tracks the current session state for the Pages CLI harness.

    Maintains the current document reference, command history, and provides
    serialisation for session persistence.
    """

    def __init__(self) -> None:
        """Initialize a new empty session."""
        self._document_name: str | None = None
        self._document_path: str | None = None
        self._history: list[dict] = []
        self._created_at: str = datetime.now(timezone.utc).isoformat()

    def set_document(self, name: str, path: str | None = None) -> None:
        """Set the current active document.

        Args:
            name: The document name.
            path: Optional file path to the document.
        """
        self._document_name = name
        self._document_path = path

    def get_document(self) -> str | None:
        """Get the name of the current active document.

        Returns:
            str or None: The current document name, or None if no document is set.
        """
        return self._document_name

    def status(self) -> dict:
        """Get the current session status.

        Returns:
            dict: Session state with keys: document_name, document_path,
                  created_at, command_count.
        """
        return {
            "document_name": self._document_name,
            "document_path": self._document_path,
            "created_at": self._created_at,
            "command_count": len(self._history),
        }

    def save_session(self, path: str) -> None:
        """Save the session state to a JSON file.

        Args:
            path: File path to save the session to.
        """
        data = {
            "document_name": self._document_name,
            "document_path": self._document_path,
            "created_at": self._created_at,
            "history": self._history,
        }
        file_path = Path(path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        file_path.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")

    def load_session(self, path: str) -> None:
        """Load a session state from a JSON file.

        Args:
            path: File path to load the session from.

        Raises:
            FileNotFoundError: If the session file does not exist.
            json.JSONDecodeError: If the file is not valid JSON.
        """
        file_path = Path(path)
        data = json.loads(file_path.read_text(encoding="utf-8"))

        self._document_name = data.get("document_name")
        self._document_path = data.get("document_path")
        self._created_at = data.get("created_at", self._created_at)
        self._history = data.get("history", [])

    def add_to_history(self, command: str) -> None:
        """Add a command to the session history.

        Args:
            command: The command string that was executed.
        """
        self._history.append({
            "command": command,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        })

    def get_history(self) -> list[dict]:
        """Get the command history.

        Returns:
            list[dict]: List of history entries, each with keys: command, timestamp.
        """
        return list(self._history)
