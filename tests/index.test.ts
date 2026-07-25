import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";

import { POLYDOC_CORE_DESCRIPTOR, POLYDOC_CORE_PACKAGE_NAME } from "../src/index.js";

describe("polydoc-core scaffold", () => {
  it("exports the package identity and current library boundary", () => {
    expect(POLYDOC_CORE_PACKAGE_NAME).toBe("@agentic-tooling/polydoc-core");
    expect(POLYDOC_CORE_DESCRIPTOR).toEqual({
      packageName: "@agentic-tooling/polydoc-core",
      purpose: "conversion-core-and-transports",
      includes: ["conversion-core", "pluggable-transports"],
      excludes: ["cli", "manifest", "watch-mode", "oauth-ux", "sidecar-management"],
    });
  });

  it("tracks golden fixture inputs for future conversion behavior", async () => {
    const fixture = await readFile("tests/fixtures/golden/publish/basic-note/input.md", "utf8");

    expect(fixture).toContain("[[TeamWiki]]");
    expect(fixture).toContain("> [!note] Publish target");
  });
});
