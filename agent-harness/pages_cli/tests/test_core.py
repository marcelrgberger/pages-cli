"""Unit tests for pages-cli core modules.

These tests verify pure Python logic without requiring Apple Pages.
External calls to osascript are mocked.
"""
import json
import os
import sys
import pytest
from unittest.mock import patch, MagicMock

# Test session module (pure Python, no mocking needed)
from pages_cli.core.session import Session

# Test export format mapping
from pages_cli.core.export import _FORMAT_MAP, get_export_formats


class TestSession:
    """Test the Session state management class."""

    def test_create_empty_session(self):
        s = Session()
        assert s.get_document() is None
        assert s.get_history() == []
        status = s.status()
        assert status["document_name"] is None
        assert status["command_count"] == 0

    def test_set_document(self):
        s = Session()
        s.set_document("MyReport", "/path/to/report.pages")
        assert s.get_document() == "MyReport"
        status = s.status()
        assert status["document_name"] == "MyReport"
        assert status["document_path"] == "/path/to/report.pages"

    def test_add_history(self):
        s = Session()
        s.add_to_history("document new")
        s.add_to_history("text add 'Hello'")
        history = s.get_history()
        assert len(history) == 2
        assert history[0]["command"] == "document new"
        assert history[1]["command"] == "text add 'Hello'"
        assert "timestamp" in history[0]

    def test_save_load_session(self, tmp_path):
        s = Session()
        s.set_document("TestDoc", "/test.pages")
        s.add_to_history("cmd1")
        s.add_to_history("cmd2")

        save_path = str(tmp_path / "session.json")
        s.save_session(save_path)

        assert os.path.exists(save_path)

        s2 = Session()
        s2.load_session(save_path)
        assert s2.get_document() == "TestDoc"
        assert len(s2.get_history()) == 2

    def test_load_missing_session(self, tmp_path):
        s = Session()
        with pytest.raises(FileNotFoundError):
            s.load_session(str(tmp_path / "nonexistent.json"))

    def test_load_invalid_json(self, tmp_path):
        bad_file = tmp_path / "bad.json"
        bad_file.write_text("not json")
        s = Session()
        with pytest.raises(json.JSONDecodeError):
            s.load_session(str(bad_file))

    def test_session_status_fields(self):
        s = Session()
        status = s.status()
        assert "document_name" in status
        assert "document_path" in status
        assert "created_at" in status
        assert "command_count" in status

    def test_history_immutability(self):
        s = Session()
        s.add_to_history("cmd1")
        history = s.get_history()
        history.append({"command": "fake"})
        assert len(s.get_history()) == 1  # Original unchanged


class TestExportFormatMapping:
    """Test export format resolution."""

    def test_pdf_aliases(self):
        assert _FORMAT_MAP["PDF"] == "PDF"
        assert _FORMAT_MAP["pdf"] == "PDF"

    def test_word_aliases(self):
        assert _FORMAT_MAP["Microsoft Word"] == "Microsoft Word"
        assert _FORMAT_MAP["word"] == "Microsoft Word"
        assert _FORMAT_MAP["docx"] == "Microsoft Word"

    def test_epub_aliases(self):
        assert _FORMAT_MAP["EPUB"] == "EPUB"
        assert _FORMAT_MAP["epub"] == "EPUB"

    def test_text_aliases(self):
        assert _FORMAT_MAP["text"] == "unformatted text"
        assert _FORMAT_MAP["txt"] == "unformatted text"

    def test_rtf_alias(self):
        assert _FORMAT_MAP["rtf"] == "formatted text"

    def test_pages09_alias(self):
        assert _FORMAT_MAP["pages09"] == "Pages 09"

    def test_invalid_format(self):
        assert _FORMAT_MAP.get("invalid_format") is None

    def test_get_export_formats(self):
        formats = get_export_formats()
        assert "PDF" in formats
        assert "Microsoft Word" in formats
        assert "EPUB" in formats
        assert len(formats) >= 5


class TestBackend:
    """Test the pages_backend module with mocked subprocess calls."""

    @patch("pages_cli.utils.pages_backend.subprocess.run")
    @patch("pages_cli.utils.pages_backend.shutil.which", return_value="/usr/bin/osascript")
    def test_run_applescript_success(self, mock_which, mock_run):
        from pages_cli.utils.pages_backend import _run_applescript
        mock_run.return_value = MagicMock(returncode=0, stdout="result\n", stderr="")
        result = _run_applescript('tell application "Pages" to get name of front document')
        assert result == "result"
        mock_run.assert_called_once()

    @patch("pages_cli.utils.pages_backend.subprocess.run")
    @patch("pages_cli.utils.pages_backend.shutil.which", return_value="/usr/bin/osascript")
    def test_run_applescript_failure(self, mock_which, mock_run):
        from pages_cli.utils.pages_backend import _run_applescript
        mock_run.return_value = MagicMock(returncode=1, stdout="", stderr="error message")
        with pytest.raises(RuntimeError, match="error message"):
            _run_applescript("bad script")

    @patch("pages_cli.utils.pages_backend.shutil.which", return_value=None)
    def test_run_applescript_no_osascript(self, mock_which):
        from pages_cli.utils.pages_backend import _run_applescript
        with pytest.raises(RuntimeError, match="osascript not found"):
            _run_applescript("any script")

    @patch("pages_cli.utils.pages_backend.subprocess.run")
    @patch("pages_cli.utils.pages_backend.shutil.which", return_value="/usr/bin/osascript")
    def test_run_applescript_timeout(self, mock_which, mock_run):
        import subprocess
        from pages_cli.utils.pages_backend import _run_applescript
        mock_run.side_effect = subprocess.TimeoutExpired(cmd="osascript", timeout=30)
        with pytest.raises(RuntimeError, match="timed out"):
            _run_applescript("slow script")

    def test_find_pages_app_exists(self):
        from pages_cli.utils.pages_backend import find_pages
        # This test runs on macOS where Pages should be installed
        result = find_pages()
        assert "Pages" in result

    @patch("pages_cli.utils.pages_backend.shutil.which", return_value="/usr/bin/osascript")
    def test_is_pages_running_returns_bool(self, mock_which):
        from pages_cli.utils.pages_backend import is_pages_running
        result = is_pages_running()
        assert isinstance(result, bool)


class TestCLI:
    """Test CLI structure and help output."""

    def test_cli_help(self):
        from click.testing import CliRunner
        from pages_cli.pages_cli import cli
        runner = CliRunner()
        result = runner.invoke(cli, ["--help"])
        assert result.exit_code == 0
        assert "pages-cli" in result.output

    def test_cli_version(self):
        from click.testing import CliRunner
        from pages_cli.pages_cli import cli
        runner = CliRunner()
        result = runner.invoke(cli, ["--version"])
        assert result.exit_code == 0
        assert "1.0.0" in result.output

    def test_document_group_help(self):
        from click.testing import CliRunner
        from pages_cli.pages_cli import cli
        runner = CliRunner()
        result = runner.invoke(cli, ["document", "--help"])
        assert result.exit_code == 0
        assert "new" in result.output
        assert "open" in result.output

    def test_export_formats_command(self):
        from click.testing import CliRunner
        from pages_cli.pages_cli import cli
        runner = CliRunner()
        result = runner.invoke(cli, ["--json", "export", "formats"])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert "formats" in data
        assert len(data["formats"]) >= 5

    def test_all_command_groups_registered(self):
        from pages_cli.pages_cli import cli
        group_names = [cmd for cmd in cli.commands]
        assert "document" in group_names
        assert "text" in group_names
        assert "table" in group_names
        assert "media" in group_names
        assert "export" in group_names
        assert "template" in group_names
        assert "session" in group_names
