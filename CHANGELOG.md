# Changelog

All notable changes to this project are documented in this file.

Entries are generated from [Conventional Commits](https://www.conventionalcommits.org/)
by [git-cliff](https://git-cliff.org/). Changes that a commit subject cannot
convey — dependency moves, notes about the shape of the public API — are added
to the release entry by hand. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-07-26

### Features

- Add deterministic Pandoc conversion
- Add local file transport and CI
- Add SharePoint transport
- Add GoogleDriveTransport for uploading docx as Google Docs
- Add docx-to-markdown reverse conversion (Tier-2) ([#16](https://github.com/agentic-tooling/polydoc-core/pull/16))

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

- A Pandoc 3.x binary must be on `PATH`. It is not bundled. `doctor()` probes
  for it and reports actionable failures.
- `PandocErrorCode` and `TransportErrorCode` are open unions that grow as
  features land. Narrowing on them with an exhaustive `switch` and a `never`
  assertion may need updating between releases.
- Reverse conversion is lossy and Word-only by design; Google Docs is
  publish-only. An empty `report.unmappable` is not proof of a lossless
  round trip — see the reverse-conversion section of the README.

