# polydoc-core

`polydoc-core` is the reusable TypeScript library layer for Polydoc: deterministic
Markdown-to-DOCX conversion plus hook contracts for consumers that need their own
Markdown transforms. It is intended to be consumed by the standalone `polydoc`
tool and by TeamWiki workflows that need the same conversion behavior without
taking on a CLI.

The original Polydoc design work scoped a local-first Markdown-to-Word and
Markdown-to-Google Docs workflow. This package keeps only the reusable library
boundary from that work. CLI commands, project manifests, watch mode, OAuth user
experience, transports, and sidecar storage stay outside this repository.

## API

The package is ESM-only and exports the core Pandoc contract from
`@agentic-tooling/polydoc-core`.

```ts
import {
  SUPPORTED_PANDOC_MAJOR,
  convertMarkdownToDocx,
  doctor,
} from "@agentic-tooling/polydoc-core";

const probe = await doctor();

if (!probe.ok) {
  throw new Error(probe.message);
}

const docxBytes = await convertMarkdownToDocx({
  markdown: "# Publish me\n\nTeamWiki can pass no hooks here.",
  referenceDocxPath: "./reference.docx",
  sourceDateEpoch: 1_704_067_200,
  preprocessors: [
    async (markdown) => markdown.replaceAll("[[TeamWiki]]", "TeamWiki"),
  ],
});

console.log(SUPPORTED_PANDOC_MAJOR); // 3
```

`convertMarkdownToDocx()` returns DOCX bytes as a `Uint8Array`. Callers can then
write those bytes to disk, upload them to a transport, or pass them to another
library.

## Pandoc Contract

This package shells out to the system `pandoc` binary through `execa` with an
argument array. It does not bundle Pandoc and does not invoke a shell.

- Supported Pandoc policy is exported as `SUPPORTED_PANDOC_MAJOR`; the current
  supported major is `3`.
- `doctor()` runs `pandoc --version`, parses the installed version and feature
  line, and returns a typed success or failure result.
- Every conversion probes Pandoc first and fails closed before creating
  conversion files when Pandoc is missing, unparseable, or outside the supported
  major.
- Failures throw `PandocError` with a stable `code` and actionable `guidance`.
- A readable `referenceDocxPath` is required. Pandoc's `--reference-doc` option
  is the styling contract for generated Word documents.

Markdown-to-DOCX conversion uses:

```txt
pandoc --from gfm --to docx --reference-doc <reference.docx> --output <output.docx> <input.md>
```

The forward DOCX writer intentionally does not pass `--wrap=none` or
`--markdown-headings=atx`; those Pandoc options affect textual Markdown output,
not DOCX generation.

## Hooks

Forward conversion accepts Markdown preprocessors:

```ts
const docxBytes = await convertMarkdownToDocx({
  markdown,
  referenceDocxPath,
  preprocessors: [
    (source) => source.replaceAll("[[", "").replaceAll("]]", ""),
  ],
});
```

Preprocessors run sequentially before Pandoc receives the Markdown. They receive
a context object with `phase: "preprocess"` and `targetFormat: "docx"`.

The package also exports `MarkdownPostprocessor` and
`applyMarkdownPostprocessors()` for future reverse or textual Markdown pipelines.
They do not run during Markdown-to-DOCX conversion because that pipeline returns
DOCX bytes, not Markdown text.

## Determinism

`convertMarkdownToDocx()` sets `SOURCE_DATE_EPOCH` for Pandoc. By default it uses
`"0"`; callers can pass `sourceDateEpoch` as a non-negative Unix timestamp string
or number. Identical Markdown, options, reference DOCX, source date epoch, and
Pandoc binary/version are expected to produce identical bytes.

## Requirements

- Node.js 20 or newer
- pnpm 11.9.0
- Pandoc 3.x for conversion

## Development

Install dependencies:

```sh
pnpm install
```

Run the main gates:

```sh
pnpm build
pnpm test
```

Additional checks:

```sh
pnpm lint
pnpm typecheck
pnpm format:check
```

Pandoc integration tests are included in `pnpm test`. They skip cleanly when a
supported Pandoc binary is unavailable.

Format files:

```sh
pnpm format
```

Build output is emitted to `dist/` with JavaScript, source maps, TypeScript
declarations, and declaration maps. Published package contents are constrained by
the `files` allowlist in `package.json`.

## License

MIT (c) 2026 Logan Lindquist Land
