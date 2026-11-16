# Building a document from a template

The CLI has no command for this. Fill each placeholder with `set word 1` of its
paragraph so the new text takes that placeholder's style; setting `body text` directly
puts everything in the first style. Prepare the body text in a scratch folder with
exactly one line feed between paragraphs and curly quotes, and report any quote that
can't be settled (a leading apostrophe such as 'cause or 'em would become an opening
quote). For a template with a title in paragraph 1 and body in paragraph 2:

```applescript
set body to read (POSIX file "<scratch>/body.txt") as «class utf8»
tell application "Pages"
	set d to make new document with properties {document template:template "<template>"}
	delay 2
	set word 1 of paragraph 2 of body text of d to body
	delay 1
	set word 1 of paragraph 1 of body text of d to "<title>"
	delay 1
	save d in POSIX file "<target>.pages"
	delay 1
	close d saving no
end tell
```

Then check the paragraph count and the saved text on disk (`document read-file`), and
never overwrite an existing file until the new one has passed.
