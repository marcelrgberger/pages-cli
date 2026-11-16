# Editing a document that already has content

Read this before changing text in an existing document. The working rules in SKILL.md
(back up, never close every document, Pages autosaves in place) apply throughout.

## Reading a file without opening Pages

```bash
pages-cli document read-file "<file>.pages" --numbered   # add --other for text boxes, headers, footnotes; --json for structure
```

Needs no running Pages. Table cell text is not included. Paragraphs are separated by
line feeds, but a line feed is not the only paragraph break: Pages also breaks
paragraphs at U+0004 and U+0005 (stored in real documents) and, by test, at U+000B,
U+000C, U+000D and U+2029. `--numbered` and the JSON `paragraphs` list apply that rule,
and the numbering matched live Pages exactly on a 632-paragraph document containing
51 U+0005 and 8 U+0004 characters. (A plain split on line feeds gave 573 paragraphs
there, so never number paragraphs yourself from `body`.) U+2028 soft breaks and U+FFFC
object anchors stay inside their paragraph. Other stored control characters such as
U+000E are left inline, unchecked. It reads the saved file, so unsaved edits in an
open window do not appear, and Pages may not yet have autosaved a recent edit.

## Small text fixes

Pages' scripting can replace a word or character in place and the paragraph keeps its
style. Replacing the whole body (`text set`) resets styles, so don't use it for edits.
Scripting cannot set paragraph styles at all, and assigning to a character *range*
corrupts the text -- larger rewrites and style changes go through the person.

```bash
pages-cli --json text find "phrase" --ignore-case       # paragraph + character positions
pages-cli text replace "old" "new" --paragraph 3 --occurrence 1
pages-cli text replace-word 3 2 "new"                    # paragraph, word (Pages' own word boundaries)
pages-cli text replace-char 3 14 ","
pages-cli text delete-range 3 10 20                      # inclusive
```

Each edit re-reads every paragraph and reports a mismatch instead of hiding it. The
commands don't call save, but Pages autosaves (see Working rules in SKILL.md). A line
feed in the new text starts a new paragraph in the same style. Text is matched exactly,
so a straight quote will not match a curly one.

Tested on a copy of a bilingual document (632 paragraphs, thousands of combining marks,
mixed fonts and styles inside paragraphs). These results were measured on one setup
(macOS 27.2, Pages Creator Studio 15.4); timings and Pages' behaviour may differ on
other versions, so treat them as observations to re-check, not guarantees:
- Positions are Pages' own characters: they equalled Python's code-point counts in all
  632 paragraphs, including those with combining marks.
- Style survives: an italic run, a bold-italic run shortened by a letter, a right-to-left
  word in its own font replaced by a shorter one, and a lone combining mark deleted all
  kept their fonts and sizes, and later runs shifted correctly.
- A replacement that straddles a style change takes the style of the first character
  of the span; the untouched remainder keeps its own. Replace each styled run
  separately if that matters.
- Speed: about 20 ms per character edited. Edits are sent in batches of 600 so no
  script hits the 30-second limit; 247 matches took about 27 s in all, 1,929 about
  46 s. For big replaces run the command in the background or expect it to be slow.
- If a batch fails midway the error says how many assignments were applied: the
  document is then partly edited, so restore the backup.

Routine: (1) copy the `.pages` file with `cp -R` as the backup; (2) make the edits;
(3) if a check fails, close the document and copy the backup back over the file, then
report; (4) otherwise confirm on disk with `document read-file` (allow up to ~30 s for
the autosave, or `document save`); (5) delete the backup.

## Saving to a new path

`document save --path P` writes the file, but a document that was edited by script
before its first save-as can stay an **Untitled** window that is not linked to the
file (observed on macOS 27.2 / Pages Creator Studio 15.4); opening the file again then
gives a second copy. The command warns when this happens. Use
`document save --path P --reopen` to close that window without saving and open the
written file. Pages may leave an extra `Untitled.pages` autosave behind (see
pitfalls.md).
