# Changelog

All notable changes to this project are documented in this file.

Entries are generated from [Conventional Commits](https://www.conventionalcommits.org/)
by [git-cliff](https://git-cliff.org/). Changes that a commit subject cannot
convey — dependency moves, notes about the shape of the public API — are added
to the release entry by hand. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.1] - 2026-07-28

### Bug Fixes

- Source maps now resolve to original TypeScript sources

### Miscellaneous

- Add npm release workflow triggered by version tags

## [0.1.0] - 2026-07-28

### Features

- Add deterministic Pandoc conversion
- Add local file transport and CI
- Add SharePoint transport
- Add GoogleDriveTransport for uploading docx as Google Docs
- Add docx-to-markdown reverse conversion (Tier-2) ([#16](https://github.com/agentic-tooling/polydoc-core/pull/16))
- Split the public surface into subpath exports ([#13](https://github.com/agentic-tooling/polydoc-core/issues/13))

### Miscellaneous

- Scaffold TypeScript library package

### Dependencies

- `fflate` is a runtime dependency (previously a dev dependency). DOCX archive
  inspection runs at conversion time, so consumers now install it.
- `google-auth-library` is an **optional** peer dependency (`^10`). It is only
  needed when passing an `auth` client to `GoogleDriveTransport`; the package
  imports it for types alone and never at runtime.
- `@googleapis/drive` is a runtime dependency, loaded lazily on the first Drive
  upload so consumers who never publish to Drive do not pay for the SDK.
- `@azure/msal-node` is a runtime dependency, used for SharePoint client-secret
  auth.

### Notes

- The transports are subpath-only. `SharePointTransport` is imported from
  `@llbbl/polydoc-core/sharepoint` and `GoogleDriveTransport` from
  `@llbbl/polydoc-core/google`; neither is re-exported from the root
  entry point, types included. The root entry point loads no third-party SDK —
  Node builtins plus `execa` and `fflate` — so a Markdown-to-DOCX consumer never
  pays for a cloud SDK it will not call. Both transport entry points re-export
  the shared contract (`Transport`, `TransportUploadResult`, `TransportError`,
  `TransportErrorCode`, `DOCX_MIME_TYPE`) as the same bindings the root entry
  point exposes.
- `instanceof TransportError` holds across entry points under Node's resolver,
  `require()` included, because every entry point resolves to one
  `transport.js` module instance. That is a property of module identity, not of
  the package layout: a bundler that emits the root and a transport entry point
  into separate bundles with no shared chunk produces two distinct classes and
  `instanceof` then evaluates false. `TransportError.code` is the
  resolution-independent discriminant. If CJS output is ever added, `transport.js`
  must remain a single shared artifact rather than shipping alongside a
  `transport.cjs`, or this becomes the dual-package hazard.
- The package ships only the entry points listed above plus `./package.json`.
  Subpath types are declared through both `exports` and `typesVersions`, so they
  resolve under `moduleResolution: "node"` as well as `node16`/`nodenext` and
  bundler resolution. `attw` and `publint` run in CI to keep that true.
- A Pandoc 3.x binary must be on `PATH`. It is not bundled. `doctor()` probes
  for it and reports actionable failures.
- `PandocErrorCode` and `TransportErrorCode` are open unions that grow as
  features land. Narrowing on them with an exhaustive `switch` and a `never`
  assertion may need updating between releases.
- Reverse conversion is lossy and Word-only by design; Google Docs is
  publish-only. An empty `report.unmappable` is not proof of a lossless
  round trip — see the reverse-conversion section of the README.

