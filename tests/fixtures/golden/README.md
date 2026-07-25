# Golden Fixtures

Golden fixtures live under one directory per conversion path. Each scenario keeps
the source input and expected artifacts together so future conversion work can
add deterministic assertions without reorganizing the test tree.

Current paths:

- `publish/basic-note/input.md` - a small Obsidian-flavored source note for the
  first Markdown-to-publish fixture.
- `publish/basic-note/expected.document.xml` - normalized `word/document.xml`
  expected from the DOCX output. Normalization removes volatile Word revision
  IDs, bookmark numeric IDs, relationship numeric IDs, and whitespace-only
  layout differences while preserving document structure, styles, text, and
  links.
- `../reference/reference.docx` - the required Pandoc reference document used by
  integration tests for deterministic DOCX generation.
