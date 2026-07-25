# polydoc-core

`polydoc-core` is the reusable TypeScript library layer for Polydoc: the
conversion core plus pluggable transports for publishing Markdown-derived
documents. It is intended to be consumed by the standalone `polydoc` tool and by
TeamWiki workflows that need the same conversion behavior without taking on a
CLI.

The original Polydoc design work scoped a local-first Markdown-to-Word and
Markdown-to-Google Docs workflow. This package keeps only the reusable library
boundary from that work. CLI commands, project manifests, watch mode, OAuth user
experience, and sidecar storage stay outside this repository.

## Current Status

This repository is at its initial package scaffold. The public API is deliberately
small until the dedicated conversion and transport issues define stable contracts.
Today it exports package identity and a descriptor of the current library
boundary.

## Requirements

- Node.js 20 or newer
- pnpm 11.9.0

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

Format files:

```sh
pnpm format
```

Build output is emitted to `dist/` with JavaScript, source maps, TypeScript
declarations, and declaration maps. Published package contents are constrained by
the `files` allowlist in `package.json`.

## License

MIT (c) 2026 Logan Lindquist Land
