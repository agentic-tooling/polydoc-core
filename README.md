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

### SharePoint Transport

`SharePointTransport` publishes DOCX bytes to a SharePoint document library via
Microsoft Graph app-only auth. It implements the same create-or-update
`Transport.upload(canonicalId, docx)` contract as `LocalFileTransport`.

```ts
import {
  SharePointTransport,
  convertMarkdownToDocx,
} from "@agentic-tooling/polydoc-core";

const docx = await convertMarkdownToDocx({
  markdown,
  referenceDocxPath: "./reference.docx",
});

const transport = new SharePointTransport({
  siteId: "contoso.sharepoint.com,site-collection-id,site-id",
  driveId: "document-library-drive-id",
  baseFolder: "TeamWiki",
  tenantId,
  clientId,
  clientSecret,
  mapCanonicalId: (canonicalId) => `published/${canonicalId}`,
});

const destination = await transport.upload("handbook/intro", docx);

console.log(destination.destinationId); // Microsoft Graph driveItem id
console.log(destination.path); // TeamWiki/published/handbook/intro.docx
console.log(destination.webUrl); // optional Graph driveItem webUrl
```

Auth uses `@azure/msal-node` with the Microsoft Graph scope
`https://graph.microsoft.com/.default`. `Sites.Selected` is the required Entra
application permission to admin-consent on the app registration, and a separate
site-specific `write` grant must be provisioned out of band. The library does
not request `Sites.Selected` as an OAuth scope, does not require
tenant-wide `Sites.ReadWrite.All`, does not provision site grants, and does not
verify token roles.

Typical setup is:

- Create or reuse an Entra app registration and client credential.
- Add the Microsoft Graph application permission `Sites.Selected` and grant
  tenant admin consent.
- Grant that app `write` permission to the target SharePoint site using the
  Microsoft Graph site permissions API or an equivalent admin process.
- Pass `tenantId`, `clientId`, and `clientSecret` from the consuming
  application. This package never reads environment variables directly and does
  not log or expose secrets.

Consumers that use managed identity, certificates, or their own token cache can
pass `accessTokenProvider` instead of client-secret fields. The constructor
rejects configurations that provide both auth shapes.

Uploads use Microsoft Graph simple upload:

```txt
PUT /v1.0/sites/{site-id}/drives/{drive-id}/root:/{relative-path}:/content
```

The body is the DOCX bytes with DOCX content type. Every 2xx response is treated
as success, but the response body must include a non-empty Graph driveItem `id`.
The returned `destinationId` and `driveItemId` are that stable Graph ID, which
callers should persist if they need a durable handle. Re-uploading the same
canonical ID maps to the same relative path and lets SharePoint/Graph update the
existing item/version history by path; the library does not keep local upload
state.

Path mapping is deterministic. By default, a path-safe canonical ID becomes the
relative SharePoint destination. `mapCanonicalId` can map opaque IDs such as
`urn:teamwiki:note:123` to a SharePoint-safe relative path. `.docx` is appended
once, `baseFolder` is prepended when provided, and the final decoded destination
is validated before auth or network work. Path segments are encoded
individually, so spaces, `#`, and `%` are percent-encoded instead of corrupting
the URL.

Destination validation rejects traversal, empty segments, control characters,
segments with leading or trailing whitespace, decoded segments over 255
characters, decoded relative paths over 400 characters, double quote, `*`, `:`,
`<`, `>`, `?`, `\`, `|`, structural slash misuse, names beginning with `~` or
`~$`, segments ending with `.`, Windows device names including `COM0`-`COM9` and
`LPT0`-`LPT9`, `.lock`, `desktop.ini`, names containing `_vti_`, and root-level
`Forms`. The 400-character check applies to the library-relative destination;
SharePoint's actual decoded full-path limit also includes the site and document
library prefix, so Microsoft Graph can still reject a shorter relative path in a
deep site or library.

Microsoft Graph simple upload supports files up to 250 MB. `SharePointTransport`
checks this limit before token acquisition or upload. This package only supports
forward Markdown-to-DOCX publishing; it does not import, diff, or reverse-convert
SharePoint documents.

Microsoft references:

- [Client credentials and `.default`](https://learn.microsoft.com/en-us/graph/auth-v2-service)
- [MSAL Node client credential requests](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests)
- [Sites.Selected permission](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesselected)
- [Site-specific permission grants](https://learn.microsoft.com/en-us/graph/api/site-post-permissions?view=graph-rest-1.0)
- [DriveItem simple upload](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
- [OneDrive/SharePoint path addressing](https://learn.microsoft.com/en-us/graph/onedrive-addressing-driveitems)

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
