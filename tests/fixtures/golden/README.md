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
- `roundtrip/basic-note/expected.md` - the Markdown produced by feeding
  `publish/basic-note/input.md` through `convertMarkdownToDocx()` and then
  `convertDocxToMarkdown()`. The two fixtures deliberately share one source
  note so the round-trip golden stays comparable to the publish golden.

  This fixture is a record of what the round trip actually loses, not an
  aspiration. Compared with the source note it has no YAML frontmatter, its
  `[[wikilinks]]` survive only as backslash-escaped literal text, its callout
  marker is escaped and the callout's internal line break is folded into one
  blockquote paragraph, and paragraphs are re-wrapped at Pandoc's default
  column width. A consumer restores that layer with its own sidecar through a
  `MarkdownPostprocessor`; the library does not invent it.

  Regenerate it only from real Pandoc output, and only alongside a check that
  the losses it records are still the losses Pandoc produces. The neighbouring
  test asserts the frontmatter and wikilink losses directly so this file cannot
  be quietly regenerated into something that hides one.

- `../reference/reference.docx` - the required Pandoc reference document used by
  integration tests for deterministic DOCX generation.
