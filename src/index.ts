export const POLYDOC_CORE_PACKAGE_NAME = "@agentic-tooling/polydoc-core";

export const POLYDOC_CORE_DESCRIPTOR = {
  packageName: POLYDOC_CORE_PACKAGE_NAME,
  purpose: "conversion-core-and-transports",
  includes: ["conversion-core", "pluggable-transports"],
  excludes: ["cli", "manifest", "watch-mode", "oauth-ux", "sidecar-management"],
} as const;

export type PolydocCoreDescriptor = typeof POLYDOC_CORE_DESCRIPTOR;
