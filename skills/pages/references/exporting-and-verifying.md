# Exporting text and PDF, and verifying them

Export both together from the finished file, to temporary names, and check before
putting them in place:

```bash
pages-cli export pdf "<folder>/<name> export.pdf" --image-quality Best --password "<pw>"   # omit --password if none
pages-cli export text "<folder>/<name> export.txt"
grep -a -c '/Encrypt' "<folder>/<name> export.pdf"       # must be above 0 when a password was set
```

The text export has single line breaks between paragraphs. Compare it with
`document read-file` output, then move the files into place; don't overwrite an
existing export until the new one has passed.
