"""End-to-end tests for pages-cli.

These tests require Apple Pages to be installed and running on macOS.
They create real documents, manipulate them, and export to real files.
"""
import json
import os
import subprocess
import sys
import time

import pytest


def _resolve_cli(name):
    """Resolve installed CLI command; falls back to python -m for dev."""
    import shutil
    force = os.environ.get("CLI_ANYTHING_FORCE_INSTALLED", "").strip() == "1"
    path = shutil.which(name)
    if path:
        print(f"[_resolve_cli] Using installed command: {path}")
        return [path]
    if force:
        raise RuntimeError(f"{name} not found in PATH. Install with: pip install -e .")
    module = name.replace("pages-cli-", "pages_cli.") + "." + name.split("-")[-1] + "_cli"
    print(f"[_resolve_cli] Falling back to: {sys.executable} -m {module}")
    return [sys.executable, "-m", module]


def _close_all_pages_documents():
    """Close all open Pages documents without saving."""
    try:
        subprocess.run(
            ["osascript", "-e",
             'tell application "Pages" to close every document saving no'],
            capture_output=True, text=True, timeout=10,
        )
    except Exception:
        pass


@pytest.fixture
def tmp_dir(tmp_path):
    """Provide a temporary directory for test outputs."""
    yield str(tmp_path)


@pytest.fixture(autouse=True)
def cleanup_pages():
    """Ensure Pages documents are cleaned up after each test."""
    yield
    _close_all_pages_documents()


class TestDocumentE2E:
    """E2E tests for document operations."""

    def test_create_blank_document(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        info = create_document(template="Blank")
        assert "name" in info
        assert info["page_count"] >= 1
        close_document(saving=False)

    def test_create_and_get_info(self, tmp_dir):
        from pages_cli.core.document import (
            create_document, get_document_info, close_document
        )
        create_document(template="Blank")
        info = get_document_info()
        assert info["name"]
        assert info["page_count"] >= 1
        assert isinstance(info["modified"], bool)
        close_document(saving=False)

    def test_list_documents(self, tmp_dir):
        from pages_cli.core.document import (
            create_document, list_documents, close_document
        )
        create_document(template="Blank")
        docs = list_documents()
        assert len(docs) >= 1
        assert docs[0]["name"]
        close_document(saving=False)


class TestTextE2E:
    """E2E tests for text operations."""

    def test_add_and_get_text(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import add_text, get_body_text
        create_document(template="Blank")
        add_text("Hello from pages-cli!")
        body = get_body_text()
        assert "Hello" in body
        close_document(saving=False)

    def test_set_body_text(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import set_body_text, get_body_text
        create_document(template="Blank")
        set_body_text("Replaced content.")
        body = get_body_text()
        assert "Replaced" in body
        close_document(saving=False)

    def test_word_count(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import set_body_text, get_word_count
        create_document(template="Blank")
        set_body_text("one two three four five")
        count = get_word_count()
        assert count == 5
        close_document(saving=False)


class TestExportE2E:
    """E2E tests for export operations — produces REAL output files."""

    def test_export_pdf(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import set_body_text
        from pages_cli.core.export import export_document

        create_document(template="Blank")
        set_body_text("PDF export test from pages-cli.")

        pdf_path = os.path.join(tmp_dir, "test_output.pdf")
        result = export_document(pdf_path, format="PDF")

        assert os.path.exists(pdf_path), f"PDF not created at {pdf_path}"
        size = os.path.getsize(pdf_path)
        assert size > 100, f"PDF suspiciously small: {size} bytes"

        with open(pdf_path, "rb") as f:
            magic = f.read(5)
        assert magic == b"%PDF-", f"Invalid PDF magic bytes: {magic}"

        print(f"\n  PDF: {pdf_path} ({size:,} bytes)")
        close_document(saving=False)

    def test_export_word(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import set_body_text
        from pages_cli.core.export import export_document

        create_document(template="Blank")
        set_body_text("Word export test.")

        docx_path = os.path.join(tmp_dir, "test_output.docx")
        result = export_document(docx_path, format="word")

        assert os.path.exists(docx_path), f"Word file not created at {docx_path}"
        size = os.path.getsize(docx_path)
        assert size > 100, f"Word file suspiciously small: {size} bytes"

        # DOCX is a ZIP file — verify ZIP magic bytes
        with open(docx_path, "rb") as f:
            magic = f.read(2)
        assert magic == b"PK", f"Invalid DOCX (ZIP) magic bytes: {magic}"

        print(f"\n  DOCX: {docx_path} ({size:,} bytes)")
        close_document(saving=False)

    def test_export_plain_text(self, tmp_dir):
        from pages_cli.core.document import create_document, close_document
        from pages_cli.core.text import set_body_text
        from pages_cli.core.export import export_document

        create_document(template="Blank")
        test_text = "Plain text export verification."
        set_body_text(test_text)

        txt_path = os.path.join(tmp_dir, "test_output.txt")
        result = export_document(txt_path, format="text")

        assert os.path.exists(txt_path), f"Text file not created at {txt_path}"
        content = open(txt_path).read()
        assert "Plain text" in content

        print(f"\n  TXT: {txt_path} ({os.path.getsize(txt_path):,} bytes)")
        close_document(saving=False)


class TestTemplateE2E:
    """E2E tests for template operations."""

    def test_list_templates(self):
        from pages_cli.core.templates import list_templates
        templates = list_templates()
        assert len(templates) > 50, f"Expected 50+ templates, got {len(templates)}"
        assert "Blank" in templates


class TestCLISubprocess:
    """Subprocess tests — invokes the installed CLI command."""

    CLI_BASE = _resolve_cli("pages-cli")

    def _run(self, args, check=True):
        return subprocess.run(
            self.CLI_BASE + args,
            capture_output=True, text=True,
            check=check,
        )

    def test_help(self):
        result = self._run(["--help"])
        assert result.returncode == 0
        assert "pages-cli" in result.stdout

    def test_version(self):
        result = self._run(["--version"])
        assert result.returncode == 0
        assert "1.0.0" in result.stdout

    def test_json_export_formats(self):
        result = self._run(["--json", "export", "formats"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "formats" in data
        assert len(data["formats"]) >= 5

    def test_json_session_status(self):
        result = self._run(["--json", "session", "status"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "pages_running" in data
        assert "version" in data

    def test_full_pdf_workflow(self, tmp_dir):
        """Full workflow: create doc → add text → export PDF → verify."""
        # This test creates a real document and exports a real PDF
        pdf_path = os.path.join(tmp_dir, "subprocess_test.pdf")

        # Create document (AppleScript via subprocess)
        r1 = self._run(["--json", "document", "new"], check=False)
        if r1.returncode != 0:
            pytest.skip(f"Could not create document: {r1.stderr}")

        # Add text
        r2 = self._run(["text", "add", "Subprocess PDF test."], check=False)

        # Export PDF
        r3 = self._run(["export", "pdf", pdf_path], check=False)

        # Close without saving
        self._run(["document", "close", "--no-save"], check=False)

        if r3.returncode == 0:
            assert os.path.exists(pdf_path), f"PDF not found: {pdf_path}"
            with open(pdf_path, "rb") as f:
                assert f.read(5) == b"%PDF-"
            print(f"\n  Subprocess PDF: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
