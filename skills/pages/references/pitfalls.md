# Pitfalls

- `set bt to body text of d` gives plain text, so `font` and similar fail on it. Use
  `paragraph i of body text of d` directly, or a reference to it.
- Pages won't report paragraph style names. To check formatting, compare `font` and
  `size` of a few paragraphs against a known good file.
- `row` is a reserved word inside `tell application "Pages"`.
- Always guard `open` returning `missing value` (the CLI's `document open` does).
- Check whether Pages is running with `pgrep -x Pages`, not System Events, which may
  not be authorised.
- In zsh `timeout` doesn't exist and `log` is a built-in.
- Text copied back out of Pages may lose blank lines between paragraphs; that's an
  artefact, not an edit.
- **Pages autosaves unsaved documents into the iCloud Pages folder**
  (`~/Library/Mobile Documents/com~apple~Pages/Documents`) as `Untitled N.pages`, even
  if you close with saving no. List that folder before and after any job that creates
  documents, and move only the files you created to the Trash.
- **Pages renames unsaved documents from their content** (a new document containing
  "The cat sat on the mat" became "Cat on Mat"), so targeting an unsaved document by
  name breaks mid-job. Use `front document`, or save it first and use its file name.
