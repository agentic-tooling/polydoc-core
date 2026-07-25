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
  createSharePointClientSecretAccessTokenProvider,
  DEFAULT_SOURCE_DATE_EPOCH,
  DOCX_MIME_TYPE,
  doctor,
  encodeSharePointRelativePath,
  LocalFileTransport,
  MICROSOFT_GRAPH_DEFAULT_SCOPE,
  PandocError,
  type PandocRunner,
  SHAREPOINT_REQUIRED_APPLICATION_PERMISSION,
  SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES,
  type SharePointAccessTokenProvider,
  type SharePointConfidentialClient,
  SharePointTransport,
  SharePointTransportError,
  type SharePointTransportOptions,
  SUPPORTED_PANDOC_MAJOR,
  validateSharePointDocxSize,
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

describe("SharePoint transport", () => {
  it("uses MSAL client-credential token requests with exactly the Microsoft Graph .default scope", async () => {
    const clientApplication: SharePointConfidentialClient = {
      acquireTokenByClientCredential: vi.fn(async () => ({ accessToken: "token" })),
    };
    const provider = createSharePointClientSecretAccessTokenProvider({
      tenantId: "tenant-id",
      clientId: "client-id",
      clientSecret: "client-secret",
      clientApplication,
    });

    await expect(provider({ scopes: [MICROSOFT_GRAPH_DEFAULT_SCOPE] })).resolves.toBe("token");

    expect(clientApplication.acquireTokenByClientCredential).toHaveBeenCalledWith({
      scopes: [MICROSOFT_GRAPH_DEFAULT_SCOPE],
    });
    expect(MICROSOFT_GRAPH_DEFAULT_SCOPE).toBe("https://graph.microsoft.com/.default");
    expect(SHAREPOINT_REQUIRED_APPLICATION_PERMISSION).toBe("Sites.Selected");
  });

  it("treats client secrets as opaque non-empty strings with normal punctuation", () => {
    const clientApplication: SharePointConfidentialClient = {
      acquireTokenByClientCredential: vi.fn(async () => ({ accessToken: "token" })),
    };

    expect(() =>
      createSharePointClientSecretAccessTokenProvider({
        tenantId: "tenant-id",
        clientId: "client-id",
        clientSecret: "opaque/secret:with?punctuation|#% and spaces",
        clientApplication,
      }),
    ).not.toThrow();
    expect(
      () =>
        new SharePointTransport({
          siteId: "site-id",
          driveId: "drive-id",
          tenantId: "tenant-id",
          clientId: "client-id",
          clientSecret: "opaque/secret:with?punctuation|#% and spaces",
        }),
    ).not.toThrow();
  });

  it("does not attach secret-bearing MSAL failures as auth error causes", async () => {
    const secretBearingCause = new Error("client-secret-value");
    const clientApplication: SharePointConfidentialClient = {
      acquireTokenByClientCredential: vi.fn(async () => {
        throw secretBearingCause;
      }),
    };
    const provider = createSharePointClientSecretAccessTokenProvider({
      tenantId: "tenant-id",
      clientId: "client-id",
      clientSecret: "client-secret-value",
      clientApplication,
    });

    await expect(provider({ scopes: [MICROSOFT_GRAPH_DEFAULT_SCOPE] })).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_AUTH_FAILED" &&
        error.cause === undefined &&
        !String(error).includes("client-secret-value"),
    );
  });

  it("PUTs a DOCX body to the encoded Graph path with bearer auth and content type", async () => {
    const calls: FetchCall[] = [];
    const transport = createSharePointTransport({
      fetchImpl: createSuccessfulSharePointFetch(calls, { id: "drive-item-1" }),
    });
    const docx = new Uint8Array([1, 2, 3]);

    const result = await transport.upload("Team Docs/#plan 100%", docx);

    expect(result).toEqual({
      kind: "sharepoint",
      destinationId: "drive-item-1",
      driveItemId: "drive-item-1",
      path: "Team Docs/#plan 100%.docx",
    });
    expect(calls).toHaveLength(1);
    expect(calls[0]).toEqual({
      url: "https://graph.microsoft.com/v1.0/sites/site%20id%23100%25/drives/drive%20id%23100%25/root:/Team%20Docs/%23plan%20100%25.docx:/content",
      method: "PUT",
      authorization: "Bearer test-token",
      contentType: DOCX_MIME_TYPE,
      body: [1, 2, 3],
    });
  });

  it("snapshots DOCX bytes before asynchronous token and network work", async () => {
    const calls: FetchCall[] = [];
    let releaseToken: ((token: string) => void) | undefined;
    const accessTokenProvider: SharePointAccessTokenProvider = vi.fn(
      () =>
        new Promise<string>((resolve) => {
          releaseToken = resolve;
        }),
    );
    const transport = createSharePointTransport({
      accessTokenProvider,
      fetchImpl: createSuccessfulSharePointFetch(calls, { id: "snapshot-id" }),
    });
    const docx = new Uint8Array([1, 2, 3]);

    const upload = transport.upload("snapshot", docx);
    docx.fill(9);
    releaseToken?.("test-token");
    await upload;

    expect(calls[0]?.body).toEqual([1, 2, 3]);
  });

  it("repeated uploads to the same canonical ID use the same URL and stable driveItem ID", async () => {
    const calls: FetchCall[] = [];
    const transport = createSharePointTransport({
      fetchImpl: createSuccessfulSharePointFetch(calls, { id: "stable-drive-item-id" }),
    });

    const first = await transport.upload("stable/path", new Uint8Array([1]));
    const second = await transport.upload("stable/path", new Uint8Array([2]));

    expect(first.driveItemId).toBe("stable-drive-item-id");
    expect(second.driveItemId).toBe("stable-drive-item-id");
    expect(calls).toHaveLength(2);
    expect(calls[1]?.url).toBe(calls[0]?.url);
    expect(calls[1]?.method).toBe("PUT");
  });

  it("supports custom opaque ID mapping, base folders, and .docx extension handling", async () => {
    const calls: FetchCall[] = [];
    const seenIds: string[] = [];
    const transport = createSharePointTransport({
      baseFolder: "Published Docs",
      fetchImpl: createSuccessfulSharePointFetch(calls, {
        id: "opaque-id",
        webUrl: "https://tenant.sharepoint.com/sites/wiki/Shared%20Documents/opaque.docx",
        eTag: '"etag"',
      }),
      mapCanonicalId: (canonicalId) => {
        seenIds.push(canonicalId);
        return "Opaque IDs/note-123.DOCX";
      },
    });

    const result = await transport.upload("urn:teamwiki:note:123", new Uint8Array([1]));

    expect(seenIds).toEqual(["urn:teamwiki:note:123"]);
    expect(result).toEqual({
      kind: "sharepoint",
      destinationId: "opaque-id",
      driveItemId: "opaque-id",
      path: "Published Docs/Opaque IDs/note-123.DOCX",
      webUrl: "https://tenant.sharepoint.com/sites/wiki/Shared%20Documents/opaque.docx",
      eTag: '"etag"',
    });
    expect(calls[0]?.url).toContain("/root:/Published%20Docs/Opaque%20IDs/note-123.DOCX:/content");
  });

  it("encodes supported spaces, hash, and percent characters per path segment", () => {
    expect(encodeSharePointRelativePath("Team Docs/#plan 100%.docx")).toBe(
      "Team%20Docs/%23plan%20100%25.docx",
    );
  });

  it.each([
    [
      "empty site ID",
      () =>
        ({
          siteId: "",
          driveId: "drive-id",
          accessTokenProvider: async () => "token",
        }) as SharePointTransportOptions,
    ],
    [
      "empty drive ID",
      () =>
        ({
          siteId: "site-id",
          driveId: " ",
          accessTokenProvider: async () => "token",
        }) as SharePointTransportOptions,
    ],
    [
      "mixed provider and credentials",
      () =>
        ({
          siteId: "site-id",
          driveId: "drive-id",
          accessTokenProvider: async () => "token",
          tenantId: "tenant",
          clientId: "client",
          clientSecret: "secret",
        }) as unknown as SharePointTransportOptions,
    ],
    [
      "missing auth",
      () =>
        ({
          siteId: "site-id",
          driveId: "drive-id",
        }) as unknown as SharePointTransportOptions,
    ],
  ])("rejects invalid config: %s", (_label, buildOptions) => {
    expect(() => new SharePointTransport(buildOptions())).toThrow(
      expect.objectContaining({
        name: "SharePointTransportError",
        code: "SHAREPOINT_CONFIG_INVALID",
      }),
    );
  });

  it.each(["", "  ", " leading", "trailing ", "bad\0id"])(
    "rejects invalid canonical ID %j",
    async (canonicalId) => {
      const transport = createSharePointTransport();

      await expect(transport.upload(canonicalId, new Uint8Array([1]))).rejects.toMatchObject({
        name: "SharePointTransportError",
        code: "SHAREPOINT_PATH_INVALID",
      });
    },
  );

  it.each([
    "../secret",
    "team/../secret",
    "team//secret",
    "./secret",
    "/absolute",
    String.raw`team\secret`,
    "name?",
    "name*",
    "name<",
    "name>",
    'name"',
    "name:",
    "name|",
    "~temporary",
    "~$temporary",
    "folder/trailing.",
    "CON",
    "COM0",
    "LPT1.docx",
    "LPT0.docx",
    ".lock",
    "desktop.ini",
    "Team/_vti_/note",
  ])("rejects unsafe SharePoint destination %j", async (mappedDestination) => {
    const transport = createSharePointTransport({
      mapCanonicalId: () => mappedDestination,
    });

    await expect(transport.upload("safe", new Uint8Array([1]))).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_PATH_INVALID",
    });
  });

  it.each([
    ["base folder trailing whitespace segment", { baseFolder: "TeamWiki /Published" }],
    ["mapped nested leading whitespace segment", { mapCanonicalId: (): string => "Team/ note" }],
    ["mapped nested trailing whitespace segment", { mapCanonicalId: (): string => "Team/note " }],
    [
      "root-level Forms folder after final combination",
      { mapCanonicalId: (): string => "Forms/note" },
    ],
    [
      "root-level Forms base folder after final combination",
      { baseFolder: "Forms", mapCanonicalId: (): string => "note" },
    ],
    [
      "appended extension blocked .lock final destination",
      { mapCanonicalId: (): string => ".lock" },
    ],
    [
      "appended extension blocked desktop.ini final destination",
      { mapCanonicalId: (): string => "desktop.ini" },
    ],
  ])("rejects unsafe final SharePoint destination: %s", async (_label, options) => {
    const transport = createSharePointTransport(options);

    await expect(transport.upload("safe", new Uint8Array([1]))).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_PATH_INVALID",
    });
  });

  it("rejects decoded SharePoint path segments longer than 255 characters", async () => {
    const transport = createSharePointTransport({
      mapCanonicalId: () => "a".repeat(256),
    });

    await expect(transport.upload("safe", new Uint8Array([1]))).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_PATH_INVALID",
    });
  });

  it("rejects decoded SharePoint relative paths longer than 400 characters", async () => {
    const transport = createSharePointTransport({
      baseFolder: "a".repeat(200),
      mapCanonicalId: () => `${"b".repeat(196)}.docx`,
    });

    await expect(transport.upload("safe", new Uint8Array([1]))).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_PATH_INVALID",
    });
  });

  it("enforces the 250 MB simple-upload cap without allocating a 250 MB DOCX", () => {
    expect(() => validateSharePointDocxSize(SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES)).not.toThrow();
    expect(() => validateSharePointDocxSize(SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES + 1)).toThrow(
      expect.objectContaining({
        name: "SharePointTransportError",
        code: "SHAREPOINT_DOCX_TOO_LARGE",
      }),
    );
  });

  it("guards upload size before auth or network work", async () => {
    const accessTokenProvider = vi.fn(async () => "token");
    const fetchImpl = vi.fn<typeof fetch>();
    const transport = createSharePointTransport({ accessTokenProvider, fetchImpl });
    const tooLarge = { byteLength: SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES + 1 } as Uint8Array;

    await expect(transport.upload("too-large", tooLarge)).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_DOCX_TOO_LARGE",
    });
    expect(accessTokenProvider).not.toHaveBeenCalled();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("wraps auth failures without exposing tokens or secrets in the error string", async () => {
    const secretBearingCause = new Error("secret-value");
    const transport = createSharePointTransport({
      accessTokenProvider: async () => {
        throw secretBearingCause;
      },
    });

    await expect(transport.upload("auth", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_AUTH_FAILED" &&
        error.cause === undefined &&
        !String(error).includes("secret-value"),
    );
  });

  it("wraps network failures without exposing bearer tokens in the error string", async () => {
    const tokenBearingCause = new Error("test-token");
    const transport = createSharePointTransport({
      accessTokenProvider: async () => "test-token",
      fetchImpl: async () => {
        throw tokenBearingCause;
      },
    });

    await expect(transport.upload("network", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_NETWORK_FAILED" &&
        error.cause === undefined &&
        !String(error).includes("test-token"),
    );
  });

  it("returns bounded Graph error context for non-2xx responses without leaking body text", async () => {
    const transport = createSharePointTransport({
      fetchImpl: async () =>
        new Response(
          JSON.stringify({
            error: {
              code: "accessDenied",
              message: "test-token secret-value",
            },
          }),
          {
            status: 403,
            headers: {
              "request-id": "request-123",
            },
          },
        ),
    });

    await expect(transport.upload("denied", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_HTTP_FAILED" &&
        error.context?.status === 403 &&
        error.context.requestId === "request-123" &&
        error.context.graphErrorCode === "accessDenied" &&
        !String(error).includes("test-token") &&
        !String(error).includes("secret-value"),
    );
  });

  it("sanitizes Graph request IDs and error codes before adding them to errors", async () => {
    const transport = createSharePointTransport({
      fetchImpl: async () =>
        ({
          ok: false,
          status: 429,
          headers: {
            get: (name: string) =>
              name.toLowerCase() === "request-id" ? "request-\n123\u007F" : null,
          },
          text: async () =>
            JSON.stringify({
              error: {
                code: "tooMany\nRequests\u007F",
              },
            }),
        }) as Response,
    });

    await expect(transport.upload("throttled", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_HTTP_FAILED" &&
        error.context?.requestId === "request-123" &&
        error.context.graphErrorCode === "tooManyRequests" &&
        !String(error).includes("\n") &&
        !String(error).includes("\u007F"),
    );
  });

  it("does not treat x-ms-ags-diagnostic as a Graph request ID", async () => {
    const transport = createSharePointTransport({
      fetchImpl: async () =>
        new Response(JSON.stringify({ error: { code: "accessDenied" } }), {
          status: 403,
          headers: {
            "x-ms-ags-diagnostic": "diagnostic-should-not-be-request-id",
          },
        }),
    });

    await expect(transport.upload("denied", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_HTTP_FAILED" &&
        error.context?.requestId === undefined,
    );
  });

  it("rejects invalid JSON and missing driveItem IDs", async () => {
    const invalidJsonTransport = createSharePointTransport({
      fetchImpl: async () => new Response("not-json secret-value", { status: 200 }),
    });
    const missingIdTransport = createSharePointTransport({
      fetchImpl: async () => Response.json({ name: "missing id" }, { status: 201 }),
    });

    await expect(
      invalidJsonTransport.upload("invalid-json", new Uint8Array([1])),
    ).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof SharePointTransportError &&
        error.code === "SHAREPOINT_RESPONSE_INVALID" &&
        error.cause === undefined &&
        !String(error).includes("secret-value"),
    );
    await expect(
      missingIdTransport.upload("missing-id", new Uint8Array([1])),
    ).rejects.toMatchObject({
      name: "SharePointTransportError",
      code: "SHAREPOINT_RESPONSE_INVALID",
    });
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

interface FetchCall {
  readonly url: string;
  readonly method: string | undefined;
  readonly authorization: string | undefined;
  readonly contentType: string | undefined;
  readonly body: readonly number[];
}

function createSharePointTransport(
  options: {
    readonly accessTokenProvider?: SharePointAccessTokenProvider;
    readonly baseFolder?: string;
    readonly fetchImpl?: typeof fetch;
    readonly mapCanonicalId?: (canonicalId: string) => string;
  } = {},
): SharePointTransport {
  let transportOptions: SharePointTransportOptions = {
    siteId: "site id#100%",
    driveId: "drive id#100%",
    accessTokenProvider: options.accessTokenProvider ?? (async () => "test-token"),
    fetch: options.fetchImpl ?? createSuccessfulSharePointFetch([], { id: "drive-item" }),
  };

  if (options.baseFolder !== undefined) {
    transportOptions = { ...transportOptions, baseFolder: options.baseFolder };
  }

  if (options.mapCanonicalId !== undefined) {
    transportOptions = { ...transportOptions, mapCanonicalId: options.mapCanonicalId };
  }

  return new SharePointTransport(transportOptions);
}

function createSuccessfulSharePointFetch(
  calls: FetchCall[],
  driveItem: {
    readonly id: string;
    readonly webUrl?: string;
    readonly eTag?: string;
  },
): typeof fetch {
  return async (input, init) => {
    const headers = new Headers(init?.headers);
    const body = init?.body;

    calls.push({
      url: String(input),
      method: init?.method,
      authorization: headers.get("Authorization") ?? undefined,
      contentType: headers.get("Content-Type") ?? undefined,
      body: body instanceof Uint8Array ? [...body] : [],
    });

    const responseBody: {
      id: string;
      webUrl?: string;
      eTag?: string;
    } = {
      id: driveItem.id,
    };

    if (driveItem.webUrl !== undefined) {
      responseBody.webUrl = driveItem.webUrl;
    }

    if (driveItem.eTag !== undefined) {
      responseBody.eTag = driveItem.eTag;
    }

    return Response.json(responseBody, { status: 201 });
  };
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
