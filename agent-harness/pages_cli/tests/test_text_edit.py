"""Unit tests for style-preserving find/replace and the safer open/save/backend.

Pure Python: osascript is mocked, so Apple Pages is not needed.
"""
from unittest.mock import patch

import pytest

from pages_cli.core import document as doc_mod
from pages_cli.core import text as T
from pages_cli.utils import pages_backend

MARK = T._PARA_MARK


def listing(*paragraphs):
    return MARK + MARK.join(paragraphs) + MARK


@pytest.fixture(autouse=True)
def pages_running():
    with patch("pages_cli.core.text.ensure_pages_running"), \
         patch("pages_cli.core.document.ensure_pages_running"):
        yield


class FakePages:
    """Stands in for osascript: serves paragraph listings, records edits."""

    def __init__(self, before, after=None):
        self.before = list(before)
        self.after = after
        self.edit_scripts = []
        self.edited = False

    def __call__(self, script):
        if "character id" in script and "paragraph i of body text" in script:
            current = self.after if (self.edited and self.after is not None) else self.before
            return listing(*current)
        self.edited = True
        self.edit_scripts.append(script)
        return ""


class TestGetParagraphs:
    def test_parses_listing_keeping_padding_and_empty_paragraphs(self):
        with patch("pages_cli.core.text._run_applescript",
                   return_value=listing("  padded  ", "", "last")):
            assert T.get_paragraphs() == ["  padded  ", "", "last"]

    def test_empty_document(self):
        with patch("pages_cli.core.text._run_applescript", return_value=MARK):
            assert T.get_paragraphs() == []

    def test_rejects_unexpected_output(self):
        with patch("pages_cli.core.text._run_applescript", return_value="oops"):
            with pytest.raises(RuntimeError):
                T.get_paragraphs()

    def test_marker_is_not_whitespace(self):
        assert not MARK.isspace()


class TestFind:
    paras = ["The cat sat. The cat left.", "Cat nap", "none here"]

    def test_all_occurrences_with_positions(self):
        found = T.find_in_paragraphs(self.paras, "cat")
        assert [(m["paragraph"], m["start"], m["end"]) for m in found] == [(1, 5, 7), (1, 18, 20)]

    def test_ignore_case(self):
        found = T.find_in_paragraphs(self.paras, "cat", ignore_case=True)
        assert [(m["paragraph"], m["start"]) for m in found] == [(1, 5), (1, 18), (2, 1)]
        assert found[2]["match"] == "Cat"

    def test_paragraph_scope(self):
        found = T.find_in_paragraphs(self.paras, "a", paragraph=3)
        assert found == []
        found = T.find_in_paragraphs(self.paras, "Cat", paragraph=2)
        assert len(found) == 1

    def test_matches_do_not_overlap(self):
        found = T.find_in_paragraphs(["aaaa"], "aa")
        assert [m["start"] for m in found] == [1, 3]

    def test_errors(self):
        with pytest.raises(ValueError):
            T.find_in_paragraphs(self.paras, "")
        with pytest.raises(ValueError):
            T.find_in_paragraphs(self.paras, "a\nb")
        with pytest.raises(ValueError):
            T.find_in_paragraphs(self.paras, "x", paragraph=9)

    def test_case_folding_that_changes_length_does_not_shift_positions(self):
        # "İ".lower() is two code points; positions must stay those of the original.
        paras = ["İx cat", "Cat"]
        found = T.find_in_paragraphs(paras, "cat", ignore_case=True)
        assert (2, 1) in [(m["paragraph"], m["start"]) for m in found]
        for m in found:
            assert paras[m["paragraph"] - 1][m["start"] - 1:m["end"]] == m["match"]


class TestApplyOps:
    def test_replace_and_delete(self):
        assert T._apply_ops(["hello world"], [(1, 1, 5, "HI")]) == ["HI world"]
        assert T._apply_ops(["hello world"], [(1, 6, 11, "")]) == ["hello"]

    def test_several_ops_in_one_paragraph_use_original_positions(self):
        ops = [(1, 1, 3, "X"), (1, 9, 11, "YY")]
        assert T._apply_ops(["abc def ghi"], ops) == ["X def YY"]

    def test_line_feed_splits_paragraph_without_shifting_earlier_edits(self):
        ops = [(1, 2, 2, "-\n-"), (2, 1, 1, "Z")]
        assert T._apply_ops(["abc", "def"], ops) == ["a-", "-c", "Zef"]


class TestSpanSteps:
    def test_deletes_trailing_characters_last_to_first_then_sets_first(self):
        steps = T._span_steps(2, 5, 8, 'say "hi"')
        assert steps == [(2, 8, ""), (2, 7, ""), (2, 6, ""), (2, 5, 'say "hi"')]

    def test_single_character_has_no_deletions(self):
        assert T._span_steps(1, 3, 3, "x") == [(1, 3, "x")]

    def test_script_escapes_text_and_targets_one_document(self):
        script = T._steps_script([(2, 5, 'say "hi"')], 'document "Draft"')
        assert 'set d to document "Draft"' in script
        assert 'set character 5 of paragraph 2 of body text of d to "say \\"hi\\""' in script


class TestEditSpans:
    def test_applies_edits_last_to_first_and_returns_verified_text(self):
        fake = FakePages(["abc abc", "abc"], after=["X abc", "X"])
        with patch("pages_cli.core.text._run_applescript", fake):
            result = T._edit_spans(
                ["abc abc", "abc"], [(1, 1, 3, "X"), (2, 1, 3, "X")], None
            )
        assert result == ["X abc", "X"]
        (script,) = fake.edit_scripts
        assert script.index("paragraph 2") < script.index("paragraph 1")

    def test_unexpected_result_is_reported_not_swallowed(self):
        fake = FakePages(["abc"], after=["garbled"])
        with patch("pages_cli.core.text._run_applescript", fake):
            with pytest.raises(RuntimeError, match="paragraph 1.*backup copy"):
                T._edit_spans(["abc"], [(1, 1, 3, "X")], None)

    @pytest.mark.parametrize("op", [(2, 1, 1, "x"), (1, 0, 2, "x"), (1, 3, 2, "x"), (1, 1, 9, "x")])
    def test_bounds_checked_before_touching_pages(self, op):
        fake = FakePages(["abc"])
        with patch("pages_cli.core.text._run_applescript", fake):
            with pytest.raises(ValueError):
                T._edit_spans(["abc"], [op], None)
        assert fake.edit_scripts == []

    def test_large_edit_is_sent_in_batches_last_to_first(self):
        fake = FakePages(["aaaa", "bbbb"], after=["Z", "Y"])
        with patch("pages_cli.core.text._run_applescript", fake), \
             patch.object(T, "_MAX_STEPS_PER_SCRIPT", 3):
            T._edit_spans(["aaaa", "bbbb"], [(1, 1, 4, "Z"), (2, 1, 4, "Y")], None)
        assert len(fake.edit_scripts) == 3
        order = [int(m) for script in fake.edit_scripts
                 for m in __import__("re").findall(r"of paragraph (\d) of body", script)]
        assert order == [2, 2, 2, 2, 1, 1, 1, 1]
        assert all(script.count("set character") <= 3 for script in fake.edit_scripts)

    def test_batch_failure_reports_partial_edit_and_the_backup(self):
        calls = []

        def flaky(script):
            if "character id" in script:
                return listing("aaaa", "bbbb")
            calls.append(script)
            if len(calls) == 2:
                raise RuntimeError("AppleScript timed out after 30 seconds.")
            return ""

        with patch("pages_cli.core.text._run_applescript", flaky), \
             patch.object(T, "_MAX_STEPS_PER_SCRIPT", 3):
            with pytest.raises(RuntimeError) as err:
                T._edit_spans(["aaaa", "bbbb"], [(1, 1, 4, "Z"), (2, 1, 4, "Y")], None)
        message = str(err.value)
        assert "part-way: 3 of 8" in message
        assert "backup copy" in message and "timed out" in message

    def test_carriage_return_rejected(self):
        with pytest.raises(ValueError):
            T._edit_spans(["abc"], [(1, 1, 1, "a\rb")], None)


class TestReplaceText:
    def test_replaces_all_and_reports_paragraphs(self):
        fake = FakePages(["cat dog cat", "cat"], after=["bird dog bird", "bird"])
        with patch("pages_cli.core.text._run_applescript", fake):
            result = T.replace_text("cat", "bird")
        assert result == {"replaced": 3, "paragraphs": [1, 2]}

    def test_single_occurrence(self):
        fake = FakePages(["cat dog cat"], after=["cat dog bird"])
        with patch("pages_cli.core.text._run_applescript", fake):
            assert T.replace_text("cat", "bird", occurrence=2)["replaced"] == 1

    def test_occurrence_out_of_range(self):
        fake = FakePages(["cat"])
        with patch("pages_cli.core.text._run_applescript", fake):
            with pytest.raises(ValueError):
                T.replace_text("cat", "x", occurrence=2)

    def test_no_match_changes_nothing(self):
        fake = FakePages(["cat"])
        with patch("pages_cli.core.text._run_applescript", fake):
            assert T.replace_text("zzz", "x") == {"replaced": 0, "paragraphs": []}
        assert fake.edit_scripts == []


class TestReplaceWord:
    def test_ok_when_other_paragraphs_untouched(self):
        fake = FakePages(["a b", "c d", "e f"], after=["a b", "c NEW", "e f"])
        with patch("pages_cli.core.text._run_applescript", fake):
            result = T.replace_word(2, 2, "NEW")
        assert result == {"paragraph": 2, "text": "c NEW"}
        assert 'set word 2 of paragraph 2' in fake.edit_scripts[0]

    def test_new_paragraph_from_line_feed_is_allowed(self):
        fake = FakePages(["a b", "c d", "e f"], after=["a b", "c x", "y", "e f"])
        with patch("pages_cli.core.text._run_applescript", fake):
            assert T.replace_word(2, 2, "x\ny")["text"] == "c x\ny"

    def test_collateral_change_detected(self):
        fake = FakePages(["a b", "c d"], after=["a CHANGED", "c d"])
        with patch("pages_cli.core.text._run_applescript", fake):
            with pytest.raises(RuntimeError):
                T.replace_word(2, 1, "x")

    def test_paragraph_out_of_range(self):
        fake = FakePages(["a"])
        with patch("pages_cli.core.text._run_applescript", fake):
            with pytest.raises(ValueError):
                T.replace_word(5, 1, "x")


class TestBackend:
    @pytest.mark.parametrize("code,expected", [(0, True), (1, False)])
    def test_running_check_uses_pgrep_exact_name(self, code, expected):
        with patch("pages_cli.utils.pages_backend.subprocess.run") as run:
            run.return_value.returncode = code
            assert pages_backend.is_pages_running() is expected
        assert run.call_args.args[0] == ["pgrep", "-x", "Pages"]

    def test_running_check_survives_missing_pgrep(self):
        with patch("pages_cli.utils.pages_backend.subprocess.run", side_effect=OSError):
            assert pages_backend.is_pages_running() is False


class TestOpenAndSave:
    def test_open_reports_was_open_and_describes_that_document(self):
        with patch("pages_cli.core.document._run_applescript",
                   return_value="b.pages\ntrue") as osa, \
             patch("pages_cli.core.document.get_document_info",
                   return_value={"name": "b.pages"}) as info:
            result = doc_mod.open_document("/x/b.pages")
        assert result == {"name": "b.pages", "was_open": True}
        info.assert_called_once_with("b.pages")
        script = osa.call_args.args[0]
        assert "missing value" in script and "/x/b.pages" in script

    def test_open_after_closed_is_not_was_open(self):
        with patch("pages_cli.core.document._run_applescript", return_value="b.pages\nfalse"), \
             patch("pages_cli.core.document.get_document_info", return_value={}):
            assert doc_mod.open_document("/x/b.pages")["was_open"] is False

    def test_save_returns_name_and_targets_one_document(self):
        with patch("pages_cli.core.document._run_applescript", return_value=" a.pages \n") as osa:
            assert doc_mod.save_document("Draft", "/x/a.pages") == "a.pages"
        script = osa.call_args.args[0]
        assert 'set d to document "Draft"' in script
        assert "close" not in script

    def test_reopen_closes_without_saving_then_opens_written_file(self):
        with patch("pages_cli.core.document._run_applescript", return_value="a.pages") as osa:
            doc_mod.save_document(None, "/x/a.pages", reopen=True)
        script = osa.call_args.args[0]
        assert script.index("save d in POSIX file") < script.index("close d saving no") \
            < script.index("open POSIX file")
        assert "missing value" in script

    def test_reopen_needs_a_path(self):
        with pytest.raises(ValueError):
            doc_mod.save_document(None, None, reopen=True)


class TestSaveCommandWarning:
    def invoke(self, saved_name, *args):
        from click.testing import CliRunner
        from pages_cli.pages_cli import cli
        with patch("pages_cli.core.document.save_document", return_value=saved_name):
            return CliRunner().invoke(cli, ["document", "save", *args])

    def test_warns_when_window_stays_untitled(self):
        result = self.invoke("Untitled", "--path", "/x/report.pages")
        assert "still 'Untitled'" in result.output and "--reopen" in result.output

    def test_no_warning_when_linked(self):
        result = self.invoke("report.pages", "--path", "/x/report.pages")
        assert "Warning" not in result.output

    def test_no_warning_after_reopen(self):
        result = self.invoke("report.pages", "--path", "/x/report.pages", "--reopen")
        assert "Warning" not in result.output

    def test_reopen_without_path_is_an_error(self):
        assert "--reopen requires --path" in self.invoke("x", "--reopen").output
