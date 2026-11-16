# Checking formatting consistency in a large document

This is the part that isn't obvious, and it matters a lot once a document has more
than a few dozen paragraphs.

**Don't** loop `pages-cli text get-font --paragraph N` / `get-font-size` / `get-color`
once per paragraph to audit a whole document. Each call is a separate AppleScript
round-trip through `osascript`, and a document with a few hundred paragraphs needs on
the order of a thousand-plus individual calls that way -- slow, and liable to time out
or flake partway through.

**Do** export the document once and analyze the export with a script:

```bash
pages-cli --json export word "/tmp/check.docx"
```

Then set up a throwaway venv with `python-docx` (the system Python here is externally
managed -- PEP 668 -- so don't `pip install --break-system-packages` into it):

```bash
python3 -m venv /tmp/pages-analysis-venv
source /tmp/pages-analysis-venv/bin/activate
pip install python-docx
python3 <this skill's folder>/scripts/analyze_formatting.py "/tmp/check.docx"
```

(`<this skill's folder>` is the folder holding `SKILL.md`; with the usual link that is
`~/.claude/skills/pages`.) That script walks every paragraph and run in the exported
docx and reports:
- font/color usage overall
- **font size broken down per paragraph style** -- this is the real consistency
  check. A style having its own fixed size (headings bigger than body, etc.) is
  normal; a style whose size *varies* across its own instances is the actual signal
  of a mistake.
- any paragraph that mixes more than one font or size within itself

On a book-length document this analyzed over a thousand text runs in one pass, versus
over a thousand individual CLI calls the naive per-paragraph approach would have
needed.

Clean up the export and the throwaway venv when done (they're scratch artifacts, not
something to leave lying around):

```bash
rm -rf /tmp/check.docx /tmp/pages-analysis-venv
```

Only fall back to live per-paragraph `pages-cli` queries when the document is small,
or when the check specifically needs the live/unsaved state of the open document
(the exported copy is a snapshot as of export time).

## Reading the results: investigate anomalies, don't just flag counts

When a style's usage count looks like it doesn't add up -- e.g. a "no first-line
indent" style shows up more times than there are section openings it's supposedly
for -- resist the urge to immediately call that an inconsistency. Pull up the
surrounding paragraphs and read them in context first.

The reason this matters: in one real case, a no-first-line-indent body style (used for
the paragraph right after each section heading) showed up on one extra paragraph that
wasn't a heading. Reading the surrounding text showed it was the paragraph resuming a
sentence that had been interrupted by an inset quotation -- the same typographic reason
for dropping the indent after a heading applies after a block quote too. It was a
deliberate, coherent choice, not a mistake. A pure count-based check would have wrongly
flagged it.

So: use the script's output to find *candidates* worth a look, then actually read
the paragraph text (and what precedes/follows it) before reporting something as a
formatting problem to the user.
