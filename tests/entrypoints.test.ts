import { execFile } from "node:child_process";
import { access, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { beforeAll, describe, expect, it } from "vitest";

/**
 * Entry-point guard.
 *
 * The package publishes three entry points and the root one is contractually
 * SDK-free: node builtins plus `execa` and `fflate`, nothing else. A single
 * re-export added back to `src/index.ts` would silently undo that, and neither
 * `tsc` nor the behavioral suite would notice. These tests assert against the
 * built output and the real module graph, so source-level shuffling cannot
 * satisfy them.
 */

const execFileAsync = promisify(execFile);

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const distDir = join(repoRoot, "dist");
const supportDir = join(repoRoot, "tests", "support");

const PACKAGE_NAME = "@agentic-tooling/polydoc-core";

/**
 * Every third-party package the root entry point may load. This list is the
 * contract; widening it is a deliberate act, not an accident.
 */
const CORE_ALLOWED_PACKAGES = ["execa", "fflate"];

/** Cloud SDKs that must stay behind their own entry points. */
const SDK_PACKAGE_PATTERN = /^(?:@azure\/|@googleapis\/|google-auth-library$|googleapis)/;

/**
 * Ceiling on how many modules the root entry point may load. The exact count is
 * a moving target — it tracks `execa`'s own dependency closure — so this guards
 * the order of magnitude the README describes rather than a specific number.
 * Importing either cloud SDK blows straight through it.
 */
const CORE_MODULE_CEILING = 80;

interface EntryPointProbe {
  readonly specifier: string;
  readonly elapsedMs: number;
  readonly totalModules: number;
  readonly builtinModules: number;
  readonly thirdPartyModules: number;
  readonly packages: readonly string[];
  readonly exports: readonly string[];
}

type ExportConditions = Readonly<Record<string, string | undefined>>;

interface PackageManifest {
  readonly name: string;
  readonly files: readonly string[];
  readonly exports: Readonly<Record<string, string | ExportConditions>>;
  readonly typesVersions?: Readonly<Record<string, Readonly<Record<string, readonly string[]>>>>;
}

let manifest: PackageManifest;

beforeAll(async () => {
  manifest = JSON.parse(await readFile(join(repoRoot, "package.json"), "utf8")) as PackageManifest;

  // The assertions below only mean something against real build output, so
  // build it here rather than trusting whatever dist/ happens to hold. tsc never
  // removes stale output, so a renamed source file would otherwise leave a
  // resolvable target behind locally and publish a broken one from CI's fresh
  // checkout. The compiler entry script is invoked directly instead of the .bin
  // shim so this does not depend on shell semantics.
  await rm(distDir, { recursive: true, force: true });
  await execFileAsync(
    process.execPath,
    [join(repoRoot, "node_modules", "typescript", "bin", "tsc"), "-p", "tsconfig.build.json"],
    { cwd: repoRoot },
  );
}, 60_000);

describe("package entry points", () => {
  it("publishes the cloud transports behind their own subpaths", () => {
    expect(Object.keys(manifest.exports)).toEqual(
      expect.arrayContaining([".", "./google", "./sharepoint"]),
    );
  });

  it("gives every entry point a types target and a resolvable JS target", () => {
    for (const [subpath, conditions] of libraryEntryPoints()) {
      const conditionNames = Object.keys(conditions);

      // Resolution order, not cosmetics: `types` has to win before any runtime
      // condition, and `default` has to be the last fallback so a resolver that
      // matches none of the earlier conditions still lands on a file.
      expect(conditionNames.at(0), `exports["${subpath}"] first condition`).toBe("types");
      expect(conditionNames.at(-1), `exports["${subpath}"] last condition`).toBe("default");

      expect(conditions.types, `exports["${subpath}"].types`).toMatch(/^\.\/dist\/.+\.d\.ts$/);

      const runtimeTargets = conditionNames
        .filter((name) => name !== "types")
        .map((name) => conditions[name]);

      expect(runtimeTargets.length, `exports["${subpath}"] runtime targets`).toBeGreaterThan(0);

      for (const target of runtimeTargets) {
        expect(target, `exports["${subpath}"] runtime target`).toMatch(/^\.\/dist\/.+\.js$/);
      }
    }
  });

  it("keeps subpath types resolvable under node10 module resolution", () => {
    // `exports` is invisible to `moduleResolution: "node"`, which is still the
    // default for plain tsc and for older toolchains. Without typesVersions a
    // consumer there gets TS2307 on a subpath that their bundler resolves fine,
    // and the diagnostic points at their config rather than at this package.
    const typesVersions = manifest.typesVersions?.["*"] ?? {};

    for (const [subpath, conditions] of libraryEntryPoints()) {
      if (subpath === ".") {
        continue;
      }

      expect(typesVersions[subpath.replace(/^\.\//, "")], `typesVersions for ${subpath}`).toEqual([
        conditions.types,
      ]);
    }
  });

  it("ships every file the exports map resolves to", async () => {
    expect(manifest.files).toContain("dist");

    for (const [subpath, target] of Object.entries(manifest.exports)) {
      const targets = typeof target === "string" ? [target] : Object.values(target);

      for (const resolved of targets) {
        expect(resolved, `exports["${subpath}"] target`).toBeDefined();
        await expect(
          access(join(repoRoot, String(resolved))),
          `${String(resolved)} is missing from the packed output`,
        ).resolves.toBeUndefined();
      }
    }
  });
});

describe("root entry point isolation", () => {
  it("reaches no third-party package other than execa and fflate in the built graph", async () => {
    const specifiers = await reachableBareSpecifiers(join(distDir, "index.js"));

    expect([...specifiers].sort()).toEqual([...CORE_ALLOWED_PACKAGES].sort());
  });

  it("keeps SDK types out of the root declaration graph as well", async () => {
    // The runtime walk above cannot see `export type { … } from "./google.js"`,
    // because a type-only re-export is erased from the emitted JS. It survives
    // in index.d.ts, and the contract covers types too: a consumer of the root
    // entry point should not need the SDK's types installed to typecheck.
    // Declarations reference their siblings by `.js` specifier under NodeNext,
    // so this walk lands in the same built graph and has the same closure.
    const specifiers = await reachableBareSpecifiers(join(distDir, "index.d.ts"));

    expect([...specifiers].sort()).toEqual([...CORE_ALLOWED_PACKAGES].sort());
  });

  it("loads no cloud SDK when imported through the exports map", async () => {
    const probe = await probeEntryPoint(PACKAGE_NAME);

    expect(probe.packages.filter((name) => SDK_PACKAGE_PATTERN.test(name))).toEqual([]);
    // The allowed dependencies really are loaded, so the assertion above is
    // measuring a populated module graph rather than an empty one.
    expect(probe.packages).toEqual(expect.arrayContaining(CORE_ALLOWED_PACKAGES));
    expect(probe.totalModules).toBeLessThanOrEqual(CORE_MODULE_CEILING);
  }, 30_000);

  it("exports the conversion core and the transport contract, and no cloud transport", async () => {
    const probe = await probeEntryPoint(PACKAGE_NAME);

    expect(probe.exports).toEqual(
      expect.arrayContaining([
        "convertMarkdownToDocx",
        "convertDocxToMarkdown",
        "doctor",
        "DOCX_MIME_TYPE",
        "LocalFileTransport",
        "TransportError",
      ]),
    );
    expect(probe.exports).not.toContain("GoogleDriveTransport");
    expect(probe.exports).not.toContain("SharePointTransport");
    expect(probe.exports).not.toContain("GoogleDriveTransportError");
    expect(probe.exports).not.toContain("SharePointTransportError");
  }, 30_000);
});

describe("transport entry points", () => {
  it("resolves the google entry point and keeps the Drive SDK lazy", async () => {
    const probe = await probeEntryPoint(`${PACKAGE_NAME}/google`);

    expect(probe.exports).toEqual(
      expect.arrayContaining([
        "GoogleDriveTransport",
        "GoogleDriveTransportError",
        "DOCX_MIME_TYPE",
        "TransportError",
      ]),
    );
    // @googleapis/drive is imported inside upload(), so importing the entry
    // point alone must still load no third-party module at all — while the
    // built graph does reach the SDK, which is why it lives behind this entry
    // point and not in the barrel. The module count keeps the empty-package
    // assertion from passing on a probe that recorded nothing at all.
    expect(probe.totalModules).toBeGreaterThan(0);
    expect(probe.packages).toEqual([]);
    expect([...(await reachableBareSpecifiers(join(distDir, "google.js")))]).toContain(
      "@googleapis/drive",
    );
  }, 30_000);

  it("resolves the sharepoint entry point and owns the MSAL dependency", async () => {
    const probe = await probeEntryPoint(`${PACKAGE_NAME}/sharepoint`);

    expect(probe.exports).toEqual(
      expect.arrayContaining([
        "SharePointTransport",
        "SharePointTransportError",
        "DOCX_MIME_TYPE",
        "TransportError",
      ]),
    );
    expect(probe.packages).toContain("@azure/msal-node");
  }, 30_000);

  it("re-exports the shared transport contract as one binding, not a copy", async () => {
    const { stdout } = await execFileAsync(
      process.execPath,
      [join(supportDir, "shared-bindings-probe.mjs")],
      { cwd: repoRoot },
    );

    expect(JSON.parse(stdout)).toEqual({
      googleReExportsSameTransportError: true,
      sharePointReExportsSameTransportError: true,
      googleReExportsSameDocxMimeType: true,
      sharePointReExportsSameDocxMimeType: true,
      googleErrorIsCoreTransportError: true,
      sharePointErrorIsCoreTransportError: true,
    });
  }, 30_000);
});

/**
 * Entry points declared with a condition object, which is every entry point
 * except the bare `"./package.json": "./package.json"` passthrough.
 */
function libraryEntryPoints(): [string, ExportConditions][] {
  return Object.entries(manifest.exports).flatMap(([subpath, target]) =>
    typeof target === "string" ? [] : [[subpath, target] satisfies [string, ExportConditions]],
  );
}

const probeCache = new Map<string, Promise<EntryPointProbe>>();

function probeEntryPoint(specifier: string): Promise<EntryPointProbe> {
  const cached = probeCache.get(specifier);

  if (cached !== undefined) {
    return cached;
  }

  const probe = runEntryPointProbe(specifier);
  probeCache.set(specifier, probe);

  return probe;
}

async function runEntryPointProbe(specifier: string): Promise<EntryPointProbe> {
  const recordPath = join(
    tmpdir(),
    `polydoc-core-module-loads-${process.pid}-${Math.random().toString(36).slice(2)}.txt`,
  );
  const { stdout } = await execFileAsync(
    process.execPath,
    [join(supportDir, "entrypoint-probe.mjs"), specifier, recordPath],
    { cwd: repoRoot },
  );

  return JSON.parse(stdout) as EntryPointProbe;
}

/**
 * Walks the built module graph from one entry file and returns every bare
 * specifier it can reach, static or dynamic. A lazily imported SDK counts:
 * deferring the cost of a dependency does not remove the dependency.
 */
async function reachableBareSpecifiers(entryFile: string): Promise<Set<string>> {
  const visited = new Set<string>();
  const bareSpecifiers = new Set<string>();
  const pending = [entryFile];

  while (pending.length > 0) {
    const file = pending.pop();

    if (file === undefined || visited.has(file)) {
      continue;
    }

    visited.add(file);

    for (const specifier of moduleSpecifiers(await readFile(file, "utf8"))) {
      if (specifier.startsWith("node:")) {
        continue;
      }

      if (specifier.startsWith(".")) {
        pending.push(resolve(dirname(file), specifier));
        continue;
      }

      bareSpecifiers.add(specifier);
    }
  }

  return bareSpecifiers;
}

/**
 * Recovers module specifiers from emitted output by pattern, which assumes what
 * this build actually produces: tsc-shaped output, top-level import and export
 * statements at column zero, and biome's double-quote style.
 *
 * It does not see template-literal dynamic imports, single-quoted specifiers, or
 * `createRequire(import.meta.url)("…")`. That last one is the plausible escape
 * hatch — CJS-only SDK interop — and it is invisible to the runtime probe too,
 * since nothing is loaded at import time either way. The tradeoff is accepted
 * because the caller asserts exact equality against the allowlist: a walker that
 * stops matching anything fails loudly rather than passing silently, so only a
 * deliberately evasive import shape slips through, and that is a code review
 * problem rather than a test problem.
 */
function moduleSpecifiers(source: string): string[] {
  const patterns = [
    // `import ... from "x";` / `export ... from "x";`, including the multi-line
    // specifier lists tsc emits for a barrel.
    /^(?:import|export)\b[\s\S]*?from\s*"([^"]+)";/gm,
    // Bare side-effect import.
    /^import\s*"([^"]+)";/gm,
    // Dynamic import, which is how the Drive SDK is loaded.
    /\bimport\(\s*"([^"]+)"\s*\)/g,
  ];

  return patterns.flatMap((pattern) =>
    [...source.matchAll(pattern)].flatMap((match) => (match[1] === undefined ? [] : [match[1]])),
  );
}
