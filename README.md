# polydoc-core

`polydoc-core` is the reusable TypeScript library layer for Polydoc: deterministic
Markdown-to-DOCX conversion plus hook contracts for consumers that need their own
Markdown transforms. It is intended to be consumed by the standalone `polydoc`
tool and by TeamWiki workflows that need the same conversion behavior without
taking on a CLI.

The original Polydoc design work scoped a local-first Markdown-to-Word and
Markdown-to-Google Docs workflow. This package keeps only the reusable library
boundary from that work. CLI commands, project manifests, watch mode, OAuth user
experience, remote service integrations, and sidecar storage stay outside this
repository. The library does include the shared transport contract and a local
file transport for consumers that need deterministic DOCX writes.

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

## Transports

Transports are side-effecting adapters that publish DOCX bytes for a canonical
document ID and return a stable destination handle for later consumers.

```ts
import {
  LocalFileTransport,
  convertMarkdownToDocx,
} from "@agentic-tooling/polydoc-core";

const docx = await convertMarkdownToDocx({
  markdown,
  referenceDocxPath: "./reference.docx",
});

const transport = new LocalFileTransport({ rootDir: "./generated-docx" });
const destination = await transport.upload("teamwiki/basic-note", docx);

console.log(destination.destinationId); // absolute path to generated-docx/teamwiki/basic-note.docx
console.log(destination.path); // same path, for local-file consumers
```

`LocalFileTransport` writes or overwrites one deterministic `.docx` destination
per canonical ID, creating parent directories as needed. By default,
`teamwiki/basic-note` maps to `<rootDir>/teamwiki/basic-note.docx`. A custom
`mapCanonicalId` option can map opaque IDs, such as `urn:teamwiki:note:123`, to
a relative destination under the same root.

The local file transport validates paths lexically before writing:

- Canonical IDs must be non-empty trimmed identifiers and must not contain NUL
  bytes.
- Mapped destinations must be non-empty relative paths.
- Parent-directory segments, absolute paths, Windows drive syntax, backslashes,
  colons, Windows-invalid filename characters, control characters, empty path
  segments, Windows-reserved device names, path segments ending in dot or space,
  and NUL bytes are rejected in mapped destinations.
- The resolved destination must stay under the configured `rootDir`.

This boundary check prevents accidental lexical path escape. It does not resolve
symlinks or claim protection against a writable root that already contains
hostile symlinks.

Transport failures throw `TransportError` with a stable `code`, actionable
`guidance`, and the original `cause` when filesystem or mapper operations fail.

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
- pnpm 10.34.5
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
supported Pandoc binary is unavailable. CI has a dedicated Pandoc-backed job
that installs Pandoc 3.10.1 and requires the integration path to be available,
while the regular Node matrix can still run without Pandoc.

GitHub Actions runs blocking checks for pushes to `main` and pull requests:
format check, lint, typecheck, tests, and build across supported Node majors
20.x, 22.x, and 24.x. Dependency installation uses the checked-in lockfile with
`pnpm install --frozen-lockfile`.

Format files:

```sh
pnpm format
```

Build output is emitted to `dist/` with JavaScript, source maps, TypeScript
declarations, and declaration maps. Published package contents are constrained by
the `files` allowlist in `package.json`.

## License

MIT (c) 2026 Logan Lindquist Land
