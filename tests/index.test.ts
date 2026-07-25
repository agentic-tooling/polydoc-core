import { chmod, mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  applyMarkdownPostprocessors,
  applyMarkdownPreprocessors,
  type ConvertMarkdownToDocxOptions,
  convertMarkdownToDocx,
  DEFAULT_SOURCE_DATE_EPOCH,
  doctor,
  LocalFileTransport,
  PandocError,
  type PandocRunner,
  SUPPORTED_PANDOC_MAJOR,
} from "../src/index.js";

const fixtureInputPath = "tests/fixtures/golden/publish/basic-note/input.md";
const fixtureExpectedDocumentXmlPath =
  "tests/fixtures/golden/publish/basic-note/expected.document.xml";
const fixtureReferenceDocxPath = "tests/fixtures/reference/reference.docx";

const realPandocProbe = await doctor();
const canRepresentUnreadableFiles = process.platform !== "win32" && process.getuid?.() !== 0;
const requirePandocIntegration = process.env.POLYDOC_REQUIRE_PANDOC === "1";

const tempDirectories: string[] = [];
const unsupportedMajorOptionRegression: ConvertMarkdownToDocxOptions = {
  markdown: "# Type regression",
  referenceDocxPath: "reference.docx",
  // @ts-expect-error supportedMajor is intentionally not a public policy override.
  supportedMajor: 2,
};
void unsupportedMajorOptionRegression;

afterEach(async () => {
  while (tempDirectories.length > 0) {
    const tempDirectory = tempDirectories.pop();

    if (tempDirectory !== undefined) {
      await rm(tempDirectory, { force: true, recursive: true });
    }
  }
});

describe("pandoc doctor", () => {
  it("parses a supported Pandoc 3 version and features", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: ["pandoc 3.10", "Features: +server +lua", "Scripting engine: Lua 5.4"].join("\n"),
      stderr: "",
      exitCode: 0,
    }));

    const result = await doctor({ runner });

    expect(result).toMatchObject({
      ok: true,
      binary: "pandoc",
      supportedMajor: SUPPORTED_PANDOC_MAJOR,
      version: {
        major: 3,
        minor: 10,
        patch: undefined,
        raw: "3.10",
      },
      features: ["+server", "+lua"],
    });
    expect(runner).toHaveBeenCalledWith("pandoc", ["--version"], { reject: false });
  });

  it("reports a missing Pandoc binary without throwing", async () => {
    const error = Object.assign(new Error("spawn pandoc ENOENT"), { code: "ENOENT" });
    const runner: PandocRunner = vi.fn(async () => {
      throw error;
    });

    const result = await doctor({ runner });

    expect(result).toMatchObject({
      ok: false,
      code: "PANDOC_NOT_FOUND",
      message: "Pandoc was not found on PATH.",
      supportedMajor: SUPPORTED_PANDOC_MAJOR,
    });
    expect(result.ok ? undefined : result.guidance).toContain("Install Pandoc 3.x");
  });

  it("rejects incompatible Pandoc major versions", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: "pandoc 2.19.2\nFeatures: +lua",
      stderr: "",
      exitCode: 0,
    }));

    const result = await doctor({ runner });

    expect(result).toMatchObject({
      ok: false,
      code: "PANDOC_UNSUPPORTED_MAJOR",
      detectedVersion: {
        major: 2,
        minor: 19,
        patch: 2,
        raw: "2.19.2",
      },
    });
  });

  it("rejects unparseable Pandoc version output", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: "not pandoc",
      stderr: "",
      exitCode: 0,
    }));

    const result = await doctor({ runner });

    expect(result).toMatchObject({
      ok: false,
      code: "PANDOC_VERSION_UNPARSEABLE",
    });
  });

  it("does not allow callers to override the supported Pandoc major through conversion", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: "pandoc 2.19",
      stderr: "",
      exitCode: 0,
    }));

    await expect(
      convertMarkdownToDocx({
        markdown: "# Nope",
        referenceDocxPath: "ignored-before-reference-validation.docx",
        runner,
        supportedMajor: 2,
      } as unknown as ConvertMarkdownToDocxOptions),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_UNSUPPORTED_MAJOR",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });
});

describe("markdown hooks", () => {
  it("runs preprocessors in order with forward DOCX context", async () => {
    const seenTargets: string[] = [];
    const result = await applyMarkdownPreprocessors(
      "Hello",
      [
        (markdown, context) => {
          seenTargets.push(`${context.phase}:${context.targetFormat}`);
          return `${markdown}, TeamWiki`;
        },
        async (markdown, context) => {
          seenTargets.push(`${context.phase}:${String(context.metadata.source)}`);
          return `${markdown}!`;
        },
      ],
      { source: "fixture" },
    );

    expect(result).toBe("Hello, TeamWiki!");
    expect(seenTargets).toEqual(["preprocess:docx", "preprocess:fixture"]);
  });

  it("exports postprocessors as Markdown-to-Markdown contracts", async () => {
    const result = await applyMarkdownPostprocessors("Hello", [
      (markdown, context) => `${markdown} ${context.phase} ${context.targetFormat}`,
    ]);

    expect(result).toBe("Hello postprocess markdown");
  });

  it("wraps hook failures in typed errors", async () => {
    await expect(
      applyMarkdownPreprocessors("Hello", [
        () => {
          throw new Error("boom");
        },
      ]),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "MARKDOWN_HOOK_FAILED",
    });
  });
});

describe("markdown to DOCX conversion", () => {
  it("fails closed when Pandoc is unsupported before touching conversion files", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: "pandoc 2.19",
      stderr: "",
      exitCode: 0,
    }));
    const referenceDocxPath = await createTempDocx("reference.docx");

    await expect(
      convertMarkdownToDocx({
        markdown: "# Nope",
        referenceDocxPath,
        runner,
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_UNSUPPORTED_MAJOR",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it("requires a readable reference DOCX after a successful probe", async () => {
    const runner = createSuccessfulDoctorRunner();

    await expect(
      convertMarkdownToDocx({
        markdown: "# Missing reference",
        referenceDocxPath: "tests/fixtures/reference/missing.docx",
        runner,
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "REFERENCE_DOC_INVALID",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it.skipIf(!canRepresentUnreadableFiles)(
    "wraps unreadable reference DOCX access in a typed validation error",
    async () => {
      const runner = createSuccessfulDoctorRunner();
      const referenceDocxPath = await createTempDocx("reference.docx");

      try {
        await chmod(referenceDocxPath, 0o000);

        await expect(
          convertMarkdownToDocx({
            markdown: "# Unreadable reference",
            referenceDocxPath,
            runner,
          }),
        ).rejects.toMatchObject({
          name: "PandocError",
          code: "REFERENCE_DOC_INVALID",
          guidance: expect.stringContaining("referenceDocxPath"),
        });
        expect(runner).toHaveBeenCalledTimes(1);
      } finally {
        await chmod(referenceDocxPath, 0o600);
      }
    },
  );

  it("uses argument arrays, source date, reference DOCX, and no reverse-only writer flags", async () => {
    const calls: Array<{
      binary: string;
      args: readonly string[];
      env: Readonly<Record<string, string>> | undefined;
    }> = [];
    const referenceDocxPath = await createTempDocx("reference.docx");
    const runner: PandocRunner = vi.fn(async (binary, args, options) => {
      calls.push({ binary, args, env: options?.env });

      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10\nFeatures: +lua",
          stderr: "",
          exitCode: 0,
        };
      }

      const outputPath = args[args.indexOf("--output") + 1];
      const inputPath = args.at(-1);

      if (outputPath === undefined || inputPath === undefined) {
        throw new Error("test runner received an incomplete invocation");
      }

      expect(await readFile(inputPath, "utf8")).toBe("# Title\n\nProcessed twice\n");
      await writeFile(outputPath, "docx bytes");

      return {
        stdout: "",
        stderr: "",
        exitCode: 0,
      };
    });

    const bytes = await convertMarkdownToDocx({
      markdown: "# Title\n",
      referenceDocxPath,
      sourceDateEpoch: 123,
      preprocessors: [(markdown) => `${markdown}\nProcessed`, (markdown) => `${markdown} twice\n`],
      runner,
    });

    expect(Buffer.from(bytes).toString("utf8")).toBe("docx bytes");
    expect(calls).toHaveLength(2);
    expect(calls[1]).toEqual({
      binary: "pandoc",
      args: [
        "--from",
        "gfm",
        "--to",
        "docx",
        "--reference-doc",
        referenceDocxPath,
        "--output",
        expect.stringMatching(/output\.docx$/),
        expect.stringMatching(/input\.md$/),
      ],
      env: { SOURCE_DATE_EPOCH: "123" },
    });
    expect(calls[1]?.args).not.toContain("--wrap=none");
    expect(calls[1]?.args).not.toContain("--markdown-headings=atx");
  });

  it("wraps unexpected runner failures after a successful probe and cleans up", async () => {
    let tempOutputPath: string | undefined;
    const referenceDocxPath = await createTempDocx("reference.docx");
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10",
          stderr: "",
          exitCode: 0,
        };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];
      throw new Error("process runner exploded");
    });

    await expect(
      convertMarkdownToDocx({
        markdown: "# Broken",
        referenceDocxPath,
        runner,
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_CONVERSION_FAILED",
      guidance: expect.stringContaining("pandocPath"),
    });

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("preserves typed conversion errors thrown by the runner and cleans up", async () => {
    let tempOutputPath: string | undefined;
    const referenceDocxPath = await createTempDocx("reference.docx");
    const typedError = new PandocError(
      "PANDOC_CONVERSION_FAILED",
      "Injected typed failure.",
      "Caller guidance.",
    );
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10",
          stderr: "",
          exitCode: 0,
        };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];
      throw typedError;
    });

    await expect(
      convertMarkdownToDocx({
        markdown: "# Broken",
        referenceDocxPath,
        runner,
      }),
    ).rejects.toBe(typedError);

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("uses a fixed SOURCE_DATE_EPOCH by default", async () => {
    const referenceDocxPath = await createTempDocx("reference.docx");
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10",
          stderr: "",
          exitCode: 0,
        };
      }

      const outputPath = args[args.indexOf("--output") + 1];

      if (outputPath === undefined) {
        throw new Error("test runner received no output path");
      }

      await writeFile(outputPath, "docx bytes");

      return {
        stdout: "",
        stderr: "",
        exitCode: 0,
      };
    });

    await convertMarkdownToDocx({
      markdown: "# Title",
      referenceDocxPath,
      runner,
    });

    expect(runner).toHaveBeenLastCalledWith(
      "pandoc",
      expect.any(Array),
      expect.objectContaining({
        env: { SOURCE_DATE_EPOCH: DEFAULT_SOURCE_DATE_EPOCH },
        reject: false,
      }),
    );
  });

  it("cleans up temporary files when Pandoc conversion fails", async () => {
    let tempOutputPath: string | undefined;
    const referenceDocxPath = await createTempDocx("reference.docx");
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10",
          stderr: "",
          exitCode: 0,
        };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];

      return {
        stdout: "",
        stderr: "invalid reference doc",
        exitCode: 43,
      };
    });

    await expect(
      convertMarkdownToDocx({
        markdown: "# Broken",
        referenceDocxPath,
        runner,
      }),
    ).rejects.toBeInstanceOf(PandocError);

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("wraps missing output after exit 0 and cleans up", async () => {
    let tempOutputPath: string | undefined;
    const referenceDocxPath = await createTempDocx("reference.docx");
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return {
          stdout: "pandoc 3.10",
          stderr: "",
          exitCode: 0,
        };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];

      return {
        stdout: "",
        stderr: "",
        exitCode: 0,
      };
    });

    await expect(
      convertMarkdownToDocx({
        markdown: "# Missing output",
        referenceDocxPath,
        runner,
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_CONVERSION_FAILED",
      message: expect.stringContaining("DOCX output could not be read"),
    });

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("rejects invalid SOURCE_DATE_EPOCH values", async () => {
    await expect(
      convertMarkdownToDocx({
        markdown: "# Bad time",
        referenceDocxPath: await createTempDocx("reference.docx"),
        sourceDateEpoch: "now",
        runner: createSuccessfulDoctorRunner(),
      }),
    ).rejects.toMatchObject({
      code: "SOURCE_DATE_EPOCH_INVALID",
    });
  });
});

describe.skipIf(!realPandocProbe.ok)("Pandoc integration", () => {
  it("converts the golden Markdown fixture to deterministic DOCX bytes", async () => {
    const markdown = await readFile(fixtureInputPath, "utf8");
    const expectedDocumentXml = await readFile(fixtureExpectedDocumentXmlPath, "utf8");
    const first = await convertMarkdownToDocx({
      markdown,
      referenceDocxPath: fixtureReferenceDocxPath,
      sourceDateEpoch: 1_704_067_200,
      preprocessors: [stripYamlFrontmatter],
    });
    const second = await convertMarkdownToDocx({
      markdown,
      referenceDocxPath: fixtureReferenceDocxPath,
      sourceDateEpoch: 1_704_067_200,
      preprocessors: [stripYamlFrontmatter],
    });

    expect(Buffer.compare(Buffer.from(first), Buffer.from(second))).toBe(0);

    const documentXml = normalizeDocumentXml(first);
    expect(`${documentXml}\n`).toBe(expectedDocumentXml);
  });
});

describe.skipIf(realPandocProbe.ok)("Pandoc integration skip behavior", () => {
  it("skips golden conversion tests when Pandoc is absent or unsupported", () => {
    expect(realPandocProbe.ok).toBe(false);
  });
});

describe.skipIf(!requirePandocIntegration)("Pandoc integration requirement", () => {
  it("has a supported Pandoc binary when CI requires integration coverage", () => {
    expect(realPandocProbe).toMatchObject({ ok: true });
  });
});

describe("local file transport", () => {
  it("creates a deterministic DOCX destination for a canonical ID", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({ rootDir });
    const result = await transport.upload("teamwiki/basic-note", new TextEncoder().encode("docx"));

    expect(result).toEqual({
      kind: "local-file",
      destinationId: join(rootDir, "teamwiki", "basic-note.docx"),
      path: join(rootDir, "teamwiki", "basic-note.docx"),
    });
    expect(await readFile(result.path, "utf8")).toBe("docx");
  });

  it("overwrites the same path instead of creating duplicate destinations", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({ rootDir });

    const first = await transport.upload("notes/update", new TextEncoder().encode("first"));
    const second = await transport.upload("notes/update", new TextEncoder().encode("second"));

    expect(second).toEqual(first);
    expect(await readFile(first.path, "utf8")).toBe("second");
  });

  it("resolves safe nested mapped destinations under the configured root", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({
      rootDir,
      mapCanonicalId: (canonicalId) => `published/${canonicalId}/index`,
    });

    const result = transport.resolveDestination("handbook/intro");

    expect(result).toEqual({
      kind: "local-file",
      destinationId: join(rootDir, "published", "handbook", "intro", "index.docx"),
      path: join(rootDir, "published", "handbook", "intro", "index.docx"),
    });
  });

  it("accepts opaque canonical IDs when a custom mapper returns a safe destination", async () => {
    const rootDir = await createTempDirectory();
    const seenIds: string[] = [];
    const transport = new LocalFileTransport({
      rootDir,
      mapCanonicalId: (canonicalId) => {
        seenIds.push(canonicalId);
        return "opaque/teamwiki-note-123";
      },
    });

    const result = await transport.upload(
      "urn:teamwiki:note:123",
      new TextEncoder().encode("docx"),
    );

    expect(seenIds).toEqual(["urn:teamwiki:note:123"]);
    expect(result.path).toBe(join(rootDir, "opaque", "teamwiki-note-123.docx"));
    expect(await readFile(result.path, "utf8")).toBe("docx");
  });

  it("does not treat relative paths starting with '..' characters as parent traversal", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({ rootDir });

    const result = await transport.upload("..safe/note", new TextEncoder().encode("docx"));

    expect(result.path).toBe(join(rootDir, "..safe", "note.docx"));
    expect(await readFile(result.path, "utf8")).toBe("docx");
  });

  it("preserves an explicit .docx extension from the mapped destination", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({
      rootDir,
      mapCanonicalId: () => "exports/final.DOCX",
    });

    const result = await transport.upload("final", new TextEncoder().encode("docx"));

    expect(result.path).toBe(join(rootDir, "exports", "final.DOCX"));
  });

  it.each(["", "   ", " leading", "trailing ", "bad\0id"])(
    "rejects invalid canonical ID %j",
    async (canonicalId) => {
      const transport = new LocalFileTransport({ rootDir: await createTempDirectory() });

      await expect(transport.upload(canonicalId, new Uint8Array())).rejects.toMatchObject({
        name: "TransportError",
        code: "TRANSPORT_ID_INVALID",
      });
    },
  );

  it.each([
    "../secret",
    "team/../secret",
    "team//secret",
    "./secret",
    "/absolute",
    "C:/absolute",
    "C:relative",
    String.raw`team\secret`,
  ])("rejects default-mapped canonical ID %j as an unsafe destination", async (canonicalId) => {
    const transport = new LocalFileTransport({ rootDir: await createTempDirectory() });

    await expect(transport.upload(canonicalId, new Uint8Array())).rejects.toMatchObject({
      name: "TransportError",
      code: "TRANSPORT_DESTINATION_INVALID",
    });
  });

  it.each([
    "../outside",
    "safe/../../outside",
    "/absolute",
    "C:/absolute",
    "C:relative",
    String.raw`safe\outside`,
    "safe//outside",
    "bad\0destination",
    "CON.docx",
    "COM1",
    "name?",
    "name*",
    "trailing.",
    "nested/trailing ",
    "AUX",
    "LPT9.txt",
    "control\u001Fchar",
  ])("rejects unsafe mapped destination %j", (mappedDestination) => {
    const transport = new LocalFileTransport({
      rootDir: resolve("/tmp/polydoc-root"),
      mapCanonicalId: () => mappedDestination,
    });

    expect(() => transport.resolveDestination("safe")).toThrow(
      expect.objectContaining({
        name: "TransportError",
        code: "TRANSPORT_DESTINATION_INVALID",
      }),
    );
  });

  it("wraps mapper failures with an actionable typed error and original cause", async () => {
    const cause = new Error("mapper exploded");
    const transport = new LocalFileTransport({
      rootDir: await createTempDirectory(),
      mapCanonicalId: () => {
        throw cause;
      },
    });

    await expect(transport.upload("opaque:id", new Uint8Array())).rejects.toMatchObject({
      name: "TransportError",
      code: "TRANSPORT_DESTINATION_INVALID",
      guidance: expect.stringContaining("mapCanonicalId"),
      cause,
    });
  });

  it("snapshots DOCX bytes before asynchronous writes can observe caller mutation", async () => {
    const rootDir = await createTempDirectory();
    const transport = new LocalFileTransport({ rootDir });
    const bytes = new TextEncoder().encode("before");

    const upload = transport.upload("snapshot", bytes);
    bytes.fill("x".charCodeAt(0));
    const result = await upload;

    expect(await readFile(result.path, "utf8")).toBe("before");
  });

  it("wraps write failures with an actionable typed error and original cause", async () => {
    const rootDir = join(await createTempDirectory(), "file-root");
    await writeFile(rootDir, "not a directory");
    const transport = new LocalFileTransport({ rootDir });

    await expect(transport.upload("doc", new Uint8Array([1, 2, 3]))).rejects.toMatchObject({
      name: "TransportError",
      code: "TRANSPORT_WRITE_FAILED",
      guidance: expect.stringContaining("rootDir"),
      cause: expect.objectContaining({ code: expect.any(String) }),
    });
  });

  it("rejects an empty local root", () => {
    expect(() => new LocalFileTransport({ rootDir: " " })).toThrow(
      expect.objectContaining({
        name: "TransportError",
        code: "TRANSPORT_ROOT_INVALID",
      }),
    );
  });
});

async function createTempDocx(name: string): Promise<string> {
  const tempDirectory = await createTempDirectory();
  const docxPath = join(tempDirectory, name);
  await mkdir(join(docxPath, ".."), { recursive: true });
  await writeFile(docxPath, "fake docx");

  return docxPath;
}

async function createTempDirectory(): Promise<string> {
  const tempDirectory = await mkdtemp(join(tmpdir(), "polydoc-core-test-"));
  tempDirectories.push(tempDirectory);

  return tempDirectory;
}

function createSuccessfulDoctorRunner(): PandocRunner {
  return vi.fn(async () => ({
    stdout: "pandoc 3.10",
    stderr: "",
    exitCode: 0,
  }));
}

function stripYamlFrontmatter(markdown: string): string {
  return markdown.replace(/^---\n[\s\S]*?\n---\n\n?/, "");
}

function normalizeDocumentXml(docxBytes: Uint8Array): string {
  const files = unzipSync(docxBytes);
  const documentXml = files["word/document.xml"];

  if (documentXml === undefined) {
    throw new Error("DOCX did not include word/document.xml");
  }

  return strFromU8(documentXml)
    .replace(/ w:rsid\w+="[^"]*"/g, "")
    .replace(/w:id="\d+"/g, 'w:id="ID"')
    .replace(/r:id="rId\d+"/g, 'r:id="rId"')
    .replace(/<w:bookmarkStart[^>]*\/><w:bookmarkEnd[^>]*\/>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
