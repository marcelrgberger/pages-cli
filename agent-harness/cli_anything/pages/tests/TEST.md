# TEST.md — cli-anything-pages Test Plan & Results

## Test Inventory

- `test_core.py`: ~25 unit tests (synthetic data, no Pages required)
- `test_full_e2e.py`: ~15 E2E tests (requires Apple Pages running)

## Unit Test Plan (test_core.py)

### Session module
- Session creation, document tracking, history, serialization/deserialization
- Edge cases: empty session, missing fields in JSON

### Export module
- Format mapping resolution (all aliases)
- get_export_formats() returns complete list
- Invalid format raises ValueError

### CLI module
- --help flag works
- --version flag works
- JSON output mode flag
- All command groups registered

### Backend module
- find_pages() locates Pages app
- osascript availability check
- Error handling for failed scripts

## E2E Test Plan (test_full_e2e.py)

### Document workflows
- Create blank document → verify info → close
- Create from template → verify → export PDF → verify PDF magic bytes → close
- Open existing → get info → close

### Text workflows
- Create doc → add text → get text → verify content → close
- Create doc → add paragraph with styling → word count → close

### Table workflows
- Create doc → add table → set cells → get cells → verify → close

### Export workflows
- Export as PDF → verify %PDF- magic bytes and file size
- Export as Word → verify output exists
- Export as plain text → verify content

### CLI subprocess tests
- TestCLISubprocess with _resolve_cli("cli-anything-pages")
- test_help, test_version, test_json_session_status, test_json_export_formats
- Full workflow: create → add text → export PDF → verify

## Realistic Workflow Scenarios

### 1. Report Generation
Simulates creating a business report:
- New document from "Blank" template
- Add heading text, body paragraphs
- Add a data table with values
- Export as PDF
- Verify PDF exists and has correct magic bytes

### 2. Template Browsing
Simulates browsing templates:
- List all templates via CLI
- Verify template count > 50
- Create document from specific template

### 3. Multi-format Export
Simulates exporting to all formats:
- Create document with content
- Export as PDF, Word, plain text
- Verify each output exists

---

## Test Results

Run: `CLI_ANYTHING_FORCE_INSTALLED=1 python3 -m pytest cli_anything/pages/tests/ -v -s --tb=short`

```
[_resolve_cli] Using installed command: .venv/bin/cli-anything-pages

cli_anything/pages/tests/test_core.py::TestSession::test_create_empty_session PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_set_document PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_add_history PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_save_load_session PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_load_missing_session PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_load_invalid_json PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_session_status_fields PASSED
cli_anything/pages/tests/test_core.py::TestSession::test_history_immutability PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_pdf_aliases PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_word_aliases PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_epub_aliases PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_text_aliases PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_rtf_alias PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_pages09_alias PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_invalid_format PASSED
cli_anything/pages/tests/test_core.py::TestExportFormatMapping::test_get_export_formats PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_run_applescript_success PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_run_applescript_failure PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_run_applescript_no_osascript PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_run_applescript_timeout PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_find_pages_app_exists PASSED
cli_anything/pages/tests/test_core.py::TestBackend::test_is_pages_running_returns_bool PASSED
cli_anything/pages/tests/test_core.py::TestCLI::test_cli_help PASSED
cli_anything/pages/tests/test_core.py::TestCLI::test_cli_version PASSED
cli_anything/pages/tests/test_core.py::TestCLI::test_document_group_help PASSED
cli_anything/pages/tests/test_core.py::TestCLI::test_export_formats_command PASSED
cli_anything/pages/tests/test_core.py::TestCLI::test_all_command_groups_registered PASSED
cli_anything/pages/tests/test_full_e2e.py::TestDocumentE2E::test_create_blank_document PASSED
cli_anything/pages/tests/test_full_e2e.py::TestDocumentE2E::test_create_and_get_info PASSED
cli_anything/pages/tests/test_full_e2e.py::TestDocumentE2E::test_list_documents PASSED
cli_anything/pages/tests/test_full_e2e.py::TestTextE2E::test_add_and_get_text PASSED
cli_anything/pages/tests/test_full_e2e.py::TestTextE2E::test_set_body_text PASSED
cli_anything/pages/tests/test_full_e2e.py::TestTextE2E::test_word_count PASSED
cli_anything/pages/tests/test_full_e2e.py::TestExportE2E::test_export_pdf PASSED (9,049 bytes)
cli_anything/pages/tests/test_full_e2e.py::TestExportE2E::test_export_word PASSED (7,388 bytes)
cli_anything/pages/tests/test_full_e2e.py::TestExportE2E::test_export_plain_text PASSED (31 bytes)
cli_anything/pages/tests/test_full_e2e.py::TestTemplateE2E::test_list_templates PASSED
cli_anything/pages/tests/test_full_e2e.py::TestCLISubprocess::test_help PASSED
cli_anything/pages/tests/test_full_e2e.py::TestCLISubprocess::test_version PASSED
cli_anything/pages/tests/test_full_e2e.py::TestCLISubprocess::test_json_export_formats PASSED
cli_anything/pages/tests/test_full_e2e.py::TestCLISubprocess::test_json_session_status PASSED
cli_anything/pages/tests/test_full_e2e.py::TestCLISubprocess::test_full_pdf_workflow PASSED (7,987 bytes)

42 passed in 36.84s
```

## Summary

| Metric | Value |
|--------|-------|
| Total tests | 42 |
| Passed | 42 |
| Failed | 0 |
| Pass rate | 100% |
| Execution time | 36.84s |
| Unit tests (test_core.py) | 27 |
| E2E tests (test_full_e2e.py) | 15 |
| Subprocess tests | 5 (using installed CLI) |

## Output Verification

- **PDF**: %PDF- magic bytes verified, 9,049 bytes
- **Word/DOCX**: PK (ZIP) magic bytes verified, 7,388 bytes
- **Plain text**: Content verified, 31 bytes
- **Subprocess PDF**: Full workflow (create → text → export → verify), 7,987 bytes

## Coverage Notes

- All core modules tested: session, export, backend, CLI structure
- Real Apple Pages used for all E2E tests (no mocking)
- Subprocess tests use `_resolve_cli()` and run against installed CLI command
- Template listing verified (100+ templates available)
- Missing: table E2E tests, media E2E tests, EPUB export test (could be added)
