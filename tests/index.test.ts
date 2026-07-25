import { chmod, mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, unzipSync } from "fflate";
import { OAuth2Client } from "google-auth-library";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  applyMarkdownPostprocessors,
  applyMarkdownPreprocessors,
  type ConvertMarkdownToDocxOptions,
  convertMarkdownToDocx,
  createSharePointClientSecretAccessTokenProvider,
  DEFAULT_SOURCE_DATE_EPOCH,
  DOCX_MIME_TYPE,
  defaultGoogleDocName,
  doctor,
  encodeSharePointRelativePath,
  GOOGLE_DOC_MIME_TYPE,
  GOOGLE_DRIVE_DOCX_IMPORT_MAX_BYTES,
  GOOGLE_DRIVE_FILE_SCOPE,
  type GoogleDriveCreateFileParams,
  type GoogleDriveFileMetadata,
  type GoogleDriveFileResponse,
  type GoogleDriveFilesClient,
  GoogleDriveTransport,
  GoogleDriveTransportError,
  type GoogleDriveTransportOptions,
  type GoogleDriveUpdateFileParams,
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

// The only module mock in this suite. GoogleDriveTransport imports
// @googleapis/drive lazily and memoizes the client it builds, so that path is
// unreachable through dependency injection; every other seam stays DI.
const googleDriveModuleMock = vi.hoisted(() => ({ drive: vi.fn() }));

vi.mock("@googleapis/drive", () => ({ drive: googleDriveModuleMock.drive }));

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

describe("Google Drive transport", () => {
  it("creates a converted Google Doc from DOCX bytes when no file ID is known", async () => {
    const harness = createGoogleDriveHarness({
      file: {
        id: "file-1",
        name: "Quarterly Plan",
        mimeType: GOOGLE_DOC_MIME_TYPE,
        webViewLink: "https://docs.google.com/document/d/file-1/edit",
      },
    });
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    const result = await transport.upload("Quarterly Plan.docx", new Uint8Array([1, 2, 3]));

    expect(result).toEqual({
      kind: "google-drive",
      destinationId: "file-1",
      fileId: "file-1",
      name: "Quarterly Plan",
      mimeType: GOOGLE_DOC_MIME_TYPE,
      webViewLink: "https://docs.google.com/document/d/file-1/edit",
    });
    expect(harness.update).not.toHaveBeenCalled();
    expect(harness.calls).toEqual([
      {
        operation: "create",
        fileId: undefined,
        name: "Quarterly Plan",
        requestMimeType: GOOGLE_DOC_MIME_TYPE,
        mediaMimeType: DOCX_MIME_TYPE,
        parents: undefined,
        parameterKeys: ["fields", "media", "requestBody", "supportsAllDrives"],
        requestBodyKeys: ["mimeType", "name"],
        body: [1, 2, 3],
        fields: "id,name,mimeType,webViewLink",
        supportsAllDrives: true,
      },
    ]);
  });

  it("omits optional destination fields Drive did not return", async () => {
    const harness = createGoogleDriveHarness({ file: { id: "file-bare" } });
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await expect(transport.upload("Bare", new Uint8Array([1]))).resolves.toEqual({
      kind: "google-drive",
      destinationId: "file-bare",
      fileId: "file-bare",
      name: "Bare",
    });
  });

  it("updates the existing Doc when resolveExistingFileId yields a stored file ID", async () => {
    const harness = createGoogleDriveHarness({ file: { id: "file-existing", name: "Handbook" } });
    const seenIds: string[] = [];
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: async (canonicalId) => {
        seenIds.push(canonicalId);
        return "file-existing";
      },
    });

    const result = await transport.upload("Handbook", new Uint8Array([4, 5]));

    expect(seenIds).toEqual(["Handbook"]);
    expect(harness.create).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      kind: "google-drive",
      destinationId: "file-existing",
      fileId: "file-existing",
    });
    expect(harness.calls).toEqual([
      {
        operation: "update",
        fileId: "file-existing",
        name: "Handbook",
        requestMimeType: GOOGLE_DOC_MIME_TYPE,
        mediaMimeType: DOCX_MIME_TYPE,
        parents: undefined,
        parameterKeys: ["fields", "fileId", "media", "requestBody", "supportsAllDrives"],
        requestBodyKeys: ["mimeType", "name"],
        body: [4, 5],
        fields: "id,name,mimeType,webViewLink",
        supportsAllDrives: true,
      },
    ]);
  });

  it("creates new Docs inside folderId but never re-parents an existing Doc", async () => {
    const harness = createGoogleDriveHarness();
    const created = new GoogleDriveTransport({
      driveClient: harness.client,
      folderId: "folder-1",
    });
    const updated = new GoogleDriveTransport({
      driveClient: harness.client,
      folderId: "folder-1",
      resolveExistingFileId: () => "file-existing",
    });

    await created.upload("New Doc", new Uint8Array([1]));
    await updated.upload("Existing Doc", new Uint8Array([2]));

    expect(harness.calls[0]).toMatchObject({
      operation: "create",
      parents: ["folder-1"],
      requestBodyKeys: ["mimeType", "name", "parents"],
    });
    expect(harness.calls[1]).toMatchObject({
      operation: "update",
      fileId: "file-existing",
      parents: undefined,
      parameterKeys: ["fields", "fileId", "media", "requestBody", "supportsAllDrives"],
      requestBodyKeys: ["mimeType", "name"],
    });
    expect(harness.calls[1]?.parameterKeys).not.toContain("addParents");
  });

  it("names no Drive scope other than drive.file anywhere in the module source", async () => {
    const googleSourcePath = fileURLToPath(new URL("../src/google.ts", import.meta.url));
    const googleSource = await readFile(googleSourcePath, "utf8");
    const namedScopes = googleSource.match(/https:\/\/www\.googleapis\.com\/auth\/[\w.]+/g) ?? [];

    expect(GOOGLE_DRIVE_FILE_SCOPE).toBe("https://www.googleapis.com/auth/drive.file");
    expect(namedScopes.length).toBeGreaterThan(0);
    expect([...new Set(namedScopes)]).toEqual([GOOGLE_DRIVE_FILE_SCOPE]);
  });

  it("exposes no credentials and sends no auth material in Drive requests", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      folderId: "folder-1",
    });

    await transport.upload("No Secrets", new Uint8Array([1]));

    // The transport does retain the consumer's auth client, but every internal
    // field is a `#private`, so the only own property it can expose is the
    // folder ID the consumer already gave it.
    expect(Object.keys(transport)).toEqual(["folderId"]);
    expect(JSON.stringify(transport)).toBe('{"folderId":"folder-1"}');

    const params = harness.create.mock.calls[0]?.[0];

    expect(Object.keys(params ?? {}).sort()).toEqual([
      "fields",
      "media",
      "requestBody",
      "supportsAllDrives",
    ]);
    expect(params).not.toHaveProperty("auth");
    expect(params).not.toHaveProperty("headers");
    expect(params).not.toHaveProperty("oauth_token");
  });

  it("strips a trailing .docx extension from the default Google Doc name", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await transport.upload("Release Notes.DOCX", new Uint8Array([1]));

    expect(defaultGoogleDocName("Release Notes.docx")).toBe("Release Notes");
    expect(defaultGoogleDocName("Release Notes.DOCX")).toBe("Release Notes");
    expect(defaultGoogleDocName("docx-primer")).toBe("docx-primer");
    expect(harness.calls[0]?.name).toBe("Release Notes");
  });

  it("maps opaque canonical IDs to Drive names through mapCanonicalId", async () => {
    const harness = createGoogleDriveHarness();
    const seenIds: string[] = [];
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      mapCanonicalId: (canonicalId) => {
        seenIds.push(canonicalId);
        return "TeamWiki Note 123";
      },
    });

    await transport.upload("urn:teamwiki:note:123", new Uint8Array([1]));

    expect(seenIds).toEqual(["urn:teamwiki:note:123"]);
    expect(harness.calls[0]?.name).toBe("TeamWiki Note 123");
  });

  it("snapshots DOCX bytes before asynchronous Drive work", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: async () => undefined,
    });
    const docx = new Uint8Array([1, 2, 3]);

    const upload = transport.upload("Snapshot", docx);
    docx.fill(9);
    await upload;

    expect(harness.calls[0]?.body).toEqual([1, 2, 3]);
  });

  it.each([
    [
      "both auth and driveClient",
      (client: GoogleDriveFilesClient) =>
        ({
          auth: new OAuth2Client({ clientId: "client-id" }),
          driveClient: client,
        }) as unknown as GoogleDriveTransportOptions,
    ],
    ["neither auth nor driveClient", () => ({}) as unknown as GoogleDriveTransportOptions],
    [
      "blank folder ID",
      (client: GoogleDriveFilesClient) =>
        ({ driveClient: client, folderId: " " }) as GoogleDriveTransportOptions,
    ],
    [
      "folder ID with a path separator",
      (client: GoogleDriveFilesClient) =>
        ({ driveClient: client, folderId: "folder/child" }) as GoogleDriveTransportOptions,
    ],
  ])("rejects invalid config: %s", (_label, buildOptions) => {
    expect(() => new GoogleDriveTransport(buildOptions(createGoogleDriveHarness().client))).toThrow(
      expect.objectContaining({
        name: "GoogleDriveTransportError",
        code: "GOOGLE_DRIVE_CONFIG_INVALID",
      }),
    );
  });

  it.each(["", "  ", " leading", "trailing ", "bad\0id"])(
    "rejects invalid canonical ID %j",
    async (canonicalId) => {
      const harness = createGoogleDriveHarness();
      const transport = new GoogleDriveTransport({ driveClient: harness.client });

      await expect(transport.upload(canonicalId, new Uint8Array([1]))).rejects.toMatchObject({
        name: "GoogleDriveTransportError",
        code: "GOOGLE_DRIVE_NAME_INVALID",
      });
      expect(harness.create).not.toHaveBeenCalled();
    },
  );

  it.each([
    "teamwiki/basic-note",
    " leading",
    "trailing ",
    "control\u001Fname",
    "\u007Fdelete",
    "",
    "a".repeat(256),
  ])("rejects unsafe mapped Google Doc name %j", async (mappedName) => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      mapCanonicalId: () => mappedName,
    });

    await expect(transport.upload("safe", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_NAME_INVALID",
      guidance: expect.stringContaining("mapCanonicalId"),
    });
    expect(harness.create).not.toHaveBeenCalled();
  });

  it("wraps mapper failures with an actionable typed error and original cause", async () => {
    const cause = new Error("mapper exploded");
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      mapCanonicalId: () => {
        throw cause;
      },
    });

    await expect(transport.upload("opaque:id", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_NAME_INVALID",
      guidance: expect.stringContaining("mapCanonicalId"),
      cause,
    });
    expect(harness.create).not.toHaveBeenCalled();
  });

  it("wraps resolveExistingFileId failures and rejects unusable resolved IDs", async () => {
    const cause = new Error("manifest unavailable");
    const throwingHarness = createGoogleDriveHarness();
    const throwingTransport = new GoogleDriveTransport({
      driveClient: throwingHarness.client,
      resolveExistingFileId: () => {
        throw cause;
      },
    });
    const invalidHarness = createGoogleDriveHarness();
    const invalidTransport = new GoogleDriveTransport({
      driveClient: invalidHarness.client,
      resolveExistingFileId: () => " ",
    });

    await expect(throwingTransport.upload("doc", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_CONFIG_INVALID",
      guidance: expect.stringContaining("resolveExistingFileId"),
      cause,
    });
    await expect(invalidTransport.upload("doc", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_CONFIG_INVALID",
    });
    expect(throwingHarness.create).not.toHaveBeenCalled();
    expect(invalidHarness.create).not.toHaveBeenCalled();
    expect(invalidHarness.update).not.toHaveBeenCalled();
  });

  it("enforces the 50 MB Google Docs import cap before any Drive call", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({ driveClient: harness.client });
    const tooLarge = { byteLength: GOOGLE_DRIVE_DOCX_IMPORT_MAX_BYTES + 1 } as Uint8Array;

    await expect(transport.upload("too-large", tooLarge)).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_DOCX_TOO_LARGE",
    });
    expect(harness.create).not.toHaveBeenCalled();
    expect(harness.update).not.toHaveBeenCalled();
  });

  it("maps Drive API rejections to bounded context without leaking tokens", async () => {
    const harness = createGoogleDriveHarness({
      createError: Object.assign(new Error("ya29.secret-access-token was rejected"), {
        status: 403,
        response: {
          status: 403,
          data: {
            error: {
              code: 403,
              message: "ya29.secret-access-token was rejected",
              errors: [{ domain: "global", reason: "storageQuota\nExceeded\u007F" }],
            },
          },
        },
      }),
    });
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await expect(transport.upload("denied", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof GoogleDriveTransportError &&
        error.code === "GOOGLE_DRIVE_API_FAILED" &&
        error.context?.status === 403 &&
        error.context.reason === "storageQuotaExceeded" &&
        error.cause === undefined &&
        !String(error).includes("ya29.secret-access-token") &&
        !String(error).includes("\n") &&
        !String(error).includes("\u007F"),
    );
  });

  it("maps a 403 with an insufficient-permission reason to an auth failure", async () => {
    const harness = createGoogleDriveHarness({
      createError: Object.assign(new Error("Insufficient permissions"), {
        status: 403,
        response: {
          status: 403,
          data: {
            error: {
              code: 403,
              errors: [{ domain: "global", reason: "insufficientFilePermissions" }],
            },
          },
        },
      }),
    });
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await expect(transport.upload("no-scope", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_AUTH_FAILED",
      guidance: expect.stringContaining(GOOGLE_DRIVE_FILE_SCOPE),
      context: { status: 403, reason: "insufficientFilePermissions" },
    });
  });

  it("distinguishes an unreachable Drive API from a Drive API decision", async () => {
    const networkHarness = createGoogleDriveHarness({
      createError: Object.assign(new Error("getaddrinfo ENOTFOUND www.googleapis.com"), {
        code: "ENOTFOUND",
      }),
    });
    // A non-Error rejection carries no status either, so it must not be
    // reported as a terminal Drive decision.
    const opaqueHarness = createGoogleDriveHarness({ createError: "boom" });
    const networkTransport = new GoogleDriveTransport({ driveClient: networkHarness.client });
    const opaqueTransport = new GoogleDriveTransport({ driveClient: opaqueHarness.client });

    await expect(networkTransport.upload("offline", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_NETWORK_FAILED",
      context: {},
    });
    await expect(opaqueTransport.upload("opaque", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof GoogleDriveTransportError &&
        error.code === "GOOGLE_DRIVE_NETWORK_FAILED" &&
        error.cause === undefined &&
        JSON.stringify(error.context) === "{}",
    );
  });

  it("leaves context undefined on validation errors that never reached Drive", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await expect(transport.upload("bad/name", new Uint8Array([1]))).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof GoogleDriveTransportError &&
        error.code === "GOOGLE_DRIVE_NAME_INVALID" &&
        error.context === undefined,
    );
  });

  it("rejects a Drive response whose stored file is not a Google Doc", async () => {
    const harness = createGoogleDriveHarness({
      file: { id: "file-blob", name: "Handbook", mimeType: DOCX_MIME_TYPE },
    });
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: () => "file-blob",
    });

    await expect(transport.upload("Handbook", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_RESPONSE_INVALID",
      message: expect.stringContaining("file-blob"),
      guidance: expect.stringContaining("file-blob"),
      context: { fileId: "file-blob" },
    });
  });

  it("propagates a 404 from files.update without silently creating a duplicate", async () => {
    const harness = createGoogleDriveHarness({
      updateError: Object.assign(new Error("File not found: file-gone."), {
        status: 404,
        response: {
          status: 404,
          data: { error: { code: 404, errors: [{ domain: "global", reason: "notFound" }] } },
        },
      }),
    });
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: () => "file-gone",
    });

    await expect(transport.upload("Vanished", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_API_FAILED",
      context: { status: 404, reason: "notFound" },
    });
    expect(harness.create).not.toHaveBeenCalled();
  });

  it("rejects a mapCanonicalId that does not return a string", async () => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      mapCanonicalId: () => 42 as unknown as string,
    });

    await expect(transport.upload("doc", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_NAME_INVALID",
      cause: expect.objectContaining({ name: "TypeError" }),
    });
    expect(harness.create).not.toHaveBeenCalled();
  });

  it.each([
    ["null, the natural miss for a Map-style manifest", null],
    ["a number", 42],
    ["an object", {}],
    ["a Drive ID with a slash", "folder/file"],
  ])("rejects resolveExistingFileId returning %s", async (_label, resolved) => {
    const harness = createGoogleDriveHarness();
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: () => resolved as unknown as string | undefined,
    });

    await expect(transport.upload("doc", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_CONFIG_INVALID",
      guidance: expect.stringContaining("resolveExistingFileId"),
    });
    expect(harness.create).not.toHaveBeenCalled();
    expect(harness.update).not.toHaveBeenCalled();
  });

  it("blames the default mapper only when the default mapper produced the bad name", async () => {
    const harness = createGoogleDriveHarness();
    const defaultMapped = new GoogleDriveTransport({ driveClient: harness.client });
    const customMapped = new GoogleDriveTransport({
      driveClient: harness.client,
      mapCanonicalId: (canonicalId) => canonicalId,
    });

    await expect(defaultMapped.upload("teamwiki/note", new Uint8Array([1]))).rejects.toMatchObject({
      code: "GOOGLE_DRIVE_NAME_INVALID",
      message: expect.stringContaining("default-mapped"),
      guidance: expect.stringContaining("pass mapCanonicalId"),
    });
    await expect(customMapped.upload("teamwiki/note", new Uint8Array([1]))).rejects.toMatchObject({
      code: "GOOGLE_DRIVE_NAME_INVALID",
      guidance: expect.stringContaining("Return a single trimmed Drive document name"),
    });
  });

  it("builds the lazily imported Drive client once for concurrent uploads", async () => {
    const files = createStubDriveFiles();
    googleDriveModuleMock.drive.mockReset();
    googleDriveModuleMock.drive.mockReturnValue({ files });
    const auth = new OAuth2Client({ clientId: "client-id" });
    const transport = new GoogleDriveTransport({ auth });

    const [first, second] = await Promise.all([
      transport.upload("A", new Uint8Array([1])),
      transport.upload("B", new Uint8Array([2])),
    ]);

    expect(googleDriveModuleMock.drive).toHaveBeenCalledTimes(1);
    expect(googleDriveModuleMock.drive).toHaveBeenCalledWith({ version: "v3", auth });
    expect(files.create).toHaveBeenCalledTimes(2);
    expect(first?.fileId).toBe("file-1");
    expect(second?.fileId).toBe("file-1");
  });

  it("does not memoize a failed Drive client construction", async () => {
    const failure = new Error("Drive client construction failed");
    const files = createStubDriveFiles();
    googleDriveModuleMock.drive.mockReset();
    googleDriveModuleMock.drive
      .mockImplementationOnce(() => {
        throw failure;
      })
      .mockImplementation(() => ({ files }));
    const transport = new GoogleDriveTransport({
      auth: new OAuth2Client({ clientId: "client-id" }),
    });

    await expect(transport.upload("Retry", new Uint8Array([1]))).rejects.toBe(failure);
    expect(files.create).not.toHaveBeenCalled();

    // The first attempt must not be cached, so this retries the import.
    await expect(transport.upload("Retry", new Uint8Array([1]))).resolves.toMatchObject({
      kind: "google-drive",
      fileId: "file-1",
    });
    expect(googleDriveModuleMock.drive).toHaveBeenCalledTimes(2);
  });

  it("maps Drive 401 responses to an auth failure", async () => {
    const harness = createGoogleDriveHarness({
      file: { id: "file-1" },
      updateError: Object.assign(new Error("Invalid Credentials"), {
        status: 401,
        response: {
          status: 401,
          data: { error: { code: 401, status: "UNAUTHENTICATED" } },
        },
      }),
    });
    const transport = new GoogleDriveTransport({
      driveClient: harness.client,
      resolveExistingFileId: () => "file-1",
    });

    await expect(transport.upload("stale-token", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_AUTH_FAILED",
      guidance: expect.stringContaining(GOOGLE_DRIVE_FILE_SCOPE),
      context: { status: 401, reason: "UNAUTHENTICATED" },
    });
  });

  it.each([
    ["missing data", null],
    ["missing id", { name: "no id here" }],
    ["blank id", { id: " " }],
  ])("fails closed on an invalid Drive response: %s", async (_label, file) => {
    const harness = createGoogleDriveHarness({ file });
    const transport = new GoogleDriveTransport({ driveClient: harness.client });

    await expect(transport.upload("bad-response", new Uint8Array([1]))).rejects.toMatchObject({
      name: "GoogleDriveTransportError",
      code: "GOOGLE_DRIVE_RESPONSE_INVALID",
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

interface GoogleDriveCall {
  readonly operation: "create" | "update";
  readonly fileId: string | undefined;
  readonly name: string;
  readonly requestMimeType: string;
  readonly mediaMimeType: string;
  readonly parents: readonly string[] | undefined;
  readonly parameterKeys: readonly string[];
  readonly requestBodyKeys: readonly string[];
  readonly body: readonly number[];
  readonly fields: string;
  readonly supportsAllDrives: boolean;
}

function createGoogleDriveHarness(
  options: {
    readonly file?: GoogleDriveFileMetadata | null;
    readonly createError?: unknown;
    readonly updateError?: unknown;
  } = {},
) {
  const calls: GoogleDriveCall[] = [];
  const file = "file" in options ? options.file : { id: "file-1" };
  const respond = async (
    operation: "create" | "update",
    params: GoogleDriveCreateFileParams | GoogleDriveUpdateFileParams,
    error: unknown,
  ): Promise<GoogleDriveFileResponse> => {
    calls.push(await recordGoogleDriveCall(operation, params));

    if (error !== undefined) {
      throw error;
    }

    return { data: file };
  };
  const create = vi.fn(async (params: GoogleDriveCreateFileParams) =>
    respond("create", params, options.createError),
  );
  const update = vi.fn(async (params: GoogleDriveUpdateFileParams) =>
    respond("update", params, options.updateError),
  );
  const client: GoogleDriveFilesClient = { files: { create, update } };

  return { calls, client, create, update };
}

function createStubDriveFiles() {
  const response: GoogleDriveFileResponse = {
    data: { id: "file-1", name: "Stub", mimeType: GOOGLE_DOC_MIME_TYPE },
  };

  return {
    create: vi.fn(async (_params: GoogleDriveCreateFileParams) => response),
    update: vi.fn(async (_params: GoogleDriveUpdateFileParams) => response),
  };
}

async function recordGoogleDriveCall(
  operation: "create" | "update",
  params: GoogleDriveCreateFileParams | GoogleDriveUpdateFileParams,
): Promise<GoogleDriveCall> {
  const chunks: Buffer[] = [];

  for await (const chunk of params.media.body) {
    chunks.push(Buffer.from(chunk as Uint8Array));
  }

  return {
    operation,
    fileId: "fileId" in params ? params.fileId : undefined,
    name: params.requestBody.name,
    requestMimeType: params.requestBody.mimeType,
    mediaMimeType: params.media.mimeType,
    parents: "parents" in params.requestBody ? params.requestBody.parents : undefined,
    parameterKeys: Object.keys(params).sort(),
    requestBodyKeys: Object.keys(params.requestBody).sort(),
    body: [...Buffer.concat(chunks)],
    fields: params.fields,
    supportsAllDrives: params.supportsAllDrives,
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
