/**
 * Imports one published entry point in a fresh process and reports what Node
 * actually loaded.
 *
 *     node tests/support/entrypoint-probe.mjs <specifier> <recordPath>
 *
 * The specifier is the public one — `@llbbl/polydoc-core/google`, not
 * a guessed `dist/` path — so the probe exercises the package.json `exports`
 * map through Node's own resolver by self-reference. A typo in an export key or
 * target fails here the same way it would fail for a consumer after publish.
 */
import { readFileSync, rmSync, writeFileSync } from "node:fs";
import { register } from "node:module";

const [specifier, recordPath] = process.argv.slice(2);

if (specifier === undefined || recordPath === undefined) {
  throw new Error("Usage: entrypoint-probe.mjs <specifier> <recordPath>");
}

writeFileSync(recordPath, "");
register("./record-module-loads.mjs", import.meta.url, { data: { recordPath } });

let namespace;
let elapsedMs;
let urls;

try {
  const startedAt = performance.now();
  namespace = await import(specifier);
  elapsedMs = performance.now() - startedAt;
  urls = readFileSync(recordPath, "utf8").split("\n").filter(Boolean);
} finally {
  // A failed import still has to clean up after itself, or a resolution error
  // silently litters the temp directory on every run.
  rmSync(recordPath, { force: true });
}

const thirdPartyUrls = urls.filter((url) => url.includes("/node_modules/"));
const packages = new Set(thirdPartyUrls.map(packageNameFromUrl).filter((name) => name !== ""));

process.stdout.write(
  JSON.stringify({
    specifier,
    elapsedMs: Number(elapsedMs.toFixed(1)),
    totalModules: urls.length,
    builtinModules: urls.filter((url) => url.startsWith("node:")).length,
    thirdPartyModules: thirdPartyUrls.length,
    packages: [...packages].sort(),
    exports: Object.keys(namespace).sort(),
  }),
);

function packageNameFromUrl(url) {
  const afterLastNodeModules = url.split("/node_modules/").pop() ?? "";
  const [first = "", second = ""] = afterLastNodeModules.split("/");

  return first.startsWith("@") ? `${first}/${second}` : first;
}
