import { chmod, mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { OAuth2Client } from "google-auth-library";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  applyMarkdownPostprocessors,
  applyMarkdownPreprocessors,
  buildDocxImportReport,
  type ConvertMarkdownToDocxOptions,
  convertDocxToMarkdown,
  convertMarkdownToDocx,
  createSharePointClientSecretAccessTokenProvider,
  DEFAULT_DOCX_TRACK_CHANGES,
  DEFAULT_SOURCE_DATE_EPOCH,
  DOCX_MAX_ARCHIVE_ENTRIES,
  DOCX_MAX_HEADER_FOOTER_PARTS,
  DOCX_MAX_PART_BYTES,
  DOCX_MAX_TOTAL_BYTES,
  DOCX_MIME_TYPE,
  DOCX_TRACK_CHANGES_MODES,
  type DocxImportReport,
  type DocxUnmappableItem,
  type DocxUnmappableKind,
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
  type PandocRunnerOptions,
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
const fixtureRoundTripExpectedPath = "tests/fixtures/golden/roundtrip/basic-note/expected.md";
/** Smallest valid 1x1 PNG, used as embedded DOCX media. */
const TINY_PNG = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

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

describe("DOCX import report", () => {
  it("reports nothing for a clean document, including Pandoc's always-present empty comments part", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml("<w:p><w:r><w:t>Plain text.</w:t></w:r></w:p>"),
        // Every Pandoc-generated DOCX carries this part with no comment bodies,
        // so its presence alone must never be reported as a loss.
        "word/comments.xml": '<w:comments xmlns:w="urn:w" />',
      }),
    );

    expect(report).toEqual({
      unmappable: [],
      hasUnmappableContent: false,
      trackChanges: "accept",
      trackedChanges: { insertions: 0, deletions: 0, formattingChanges: 0 },
      documentPart: "word/document.xml",
      scannedParts: ["word/document.xml"],
    });
  });

  it("counts comment bodies rather than the comments part or its root element", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/comments.xml": [
          '<w:comments xmlns:w="urn:w">',
          '<w:comment w:id="1" w:author="Reviewer One"><w:p><w:r><w:t>First.</w:t></w:r></w:p></w:comment>',
          '<w:comment w:id="2" w:author="Reviewer Two"><w:p><w:r><w:t>Second.</w:t></w:r></w:p></w:comment>',
          "</w:comments>",
        ].join(""),
      }),
    );

    expect(findUnmappable(report, "comment")).toMatchObject({
      kind: "comment",
      count: 2,
      summary: expect.stringContaining("--track-changes=all"),
    });
  });

  it("counts tracked insertions and deletions without matching lookalike elements", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml(
          [
            "<w:tbl><w:tblPr><w:tblBorders>",
            // `w:insideH`/`w:insideV` are table borders, `w:delText` is the body
            // of a deletion, and `w:moveFromRangeStart` is a range marker, so
            // none of them is a revision mark itself.
            '<w:insideH w:val="single" /><w:insideV w:val="single" />',
            "</w:tblBorders></w:tblPr></w:tbl>",
            "<w:p>",
            '<w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins>',
            '<w:ins w:id="2" w:author="B"><w:r><w:t>also added</w:t></w:r></w:ins>',
            '<w:moveFromRangeStart w:id="8" w:name="move1" />',
            '<w:del w:id="3" w:author="A"><w:r><w:delText>removed</w:delText></w:r></w:del>',
            "</w:p>",
          ].join(""),
        ),
      }),
    );

    expect(report.trackedChanges).toEqual({
      insertions: 2,
      deletions: 1,
      formattingChanges: 0,
    });
    expect(findUnmappable(report, "tracked-change")).toMatchObject({
      kind: "tracked-change",
      count: 3,
      summary: expect.stringContaining("--track-changes=accept"),
    });
  });

  it("counts tracked moves, which Word writes as neither w:ins nor w:del", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml(
          [
            "<w:p>",
            '<w:moveFrom w:id="1" w:author="A"><w:r><w:t>moved away</w:t></w:r></w:moveFrom>',
            '<w:moveTo w:id="2" w:author="A"><w:r><w:t>moved here</w:t></w:r></w:moveTo>',
            "</w:p>",
          ].join(""),
        ),
      }),
    );

    // Verified against Pandoc 3.10: --track-changes=all renders w:moveFrom as
    // class="deletion" and w:moveTo as class="insertion", so a reorganized
    // document loses its moveFrom text under accept exactly like a deletion.
    expect(report.trackedChanges).toMatchObject({ insertions: 1, deletions: 1 });
    expect(findUnmappable(report, "tracked-change")).toBeDefined();
  });

  it("counts formatting-only revisions, which no track-changes mode preserves", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml(
          '<w:p><w:r><w:rPr><w:b /><w:rPrChange w:id="1" w:author="A"><w:rPr /></w:rPrChange></w:rPr><w:t>styled</w:t></w:r></w:p>',
        ),
      }),
      "all",
    );

    expect(report.trackedChanges).toMatchObject({ formattingChanges: 1 });
    expect(findUnmappable(report, "tracked-change")).toMatchObject({
      count: 1,
      summary: expect.stringContaining("no representation for a formatting-only revision"),
    });
  });

  it("classifies text boxes by whether a fallback exists, not by namespace", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml(
          [
            // Word 2010+ writes DrawingML plus a VML mc:Fallback. Verified
            // against Pandoc 3.10: the fallback is read, so the text survives.
            "<w:p><w:r><mc:AlternateContent>",
            '<mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>Modern</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice>',
            "<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Modern</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>",
            "</mc:AlternateContent></w:r></w:p>",
            // Bare legacy VML is also inlined.
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Legacy</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>",
            // DrawingML with no fallback is the only form actually dropped.
            "<w:p><w:r><mc:AlternateContent>",
            '<mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>Orphan</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice>',
            "</mc:AlternateContent></w:r></w:p>",
            // Bare DrawingML, no wrapper at all, is dropped too.
            "<w:p><w:r><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>Bare</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></w:r></w:p>",
          ].join(""),
        ),
      }),
    );

    expect(findUnmappable(report, "text-box")).toMatchObject({
      kind: "text-box",
      count: 4,
      summary: expect.stringContaining("2 inlined as ordinary body paragraphs"),
    });
    expect(findUnmappable(report, "text-box")?.summary).toContain("2 dropped outright");
  });

  it("counts a text box once when an AlternateContent is nested inside a fallback", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml(
          [
            "<w:p><w:r><mc:AlternateContent>",
            '<mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>Outer</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice>',
            // A non-greedy paired-tag pattern would stop at the inner closing
            // tag here, orphan the outer one, and leak un-stripped content.
            "<mc:Fallback><mc:AlternateContent><mc:Choice><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Outer</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Choice></mc:AlternateContent></mc:Fallback>",
            "</mc:AlternateContent></w:r></w:p>",
          ].join(""),
        ),
      }),
    );

    expect(findUnmappable(report, "text-box")).toMatchObject({ count: 1 });
  });

  it("reports embedded media and OLE objects with their sorted archive part names", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/media/image2.png": new Uint8Array([2]),
        "word/media/image1.png": new Uint8Array([1]),
        "word/embeddings/oleObject1.bin": new Uint8Array([3]),
      }),
    );

    expect(findUnmappable(report, "embedded-media")).toMatchObject({
      kind: "embedded-media",
      count: 2,
      entries: ["word/media/image1.png", "word/media/image2.png"],
      summary: expect.stringContaining("--extract-media"),
    });
    expect(findUnmappable(report, "embedded-object")).toMatchObject({
      kind: "embedded-object",
      count: 1,
      entries: ["word/embeddings/oleObject1.bin"],
    });
  });

  it("reports headers and footers that hold text, and ignores empty boilerplate ones", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/header1.xml":
          '<w:hdr xmlns:w="urn:w"><w:p><w:r><w:t>Running title.</w:t></w:r></w:p></w:hdr>',
        "word/footer1.xml":
          '<w:ftr xmlns:w="urn:w"><w:p><w:r><w:t>Page footer.</w:t></w:r></w:p></w:ftr>',
        // No w:t anywhere, so this one is Word furniture rather than content.
        "word/header2.xml": '<w:hdr xmlns:w="urn:w"><w:p /></w:hdr>',
      }),
    );

    expect(findUnmappable(report, "header-footer")).toMatchObject({
      kind: "header-footer",
      count: 2,
      entries: ["word/footer1.xml", "word/header1.xml"],
      summary: expect.stringContaining("ignores headers and footers entirely"),
    });
  });

  it("counts revisions inside footnotes and endnotes, whose text does reach the Markdown", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/footnotes.xml": [
          '<w:footnotes xmlns:w="urn:w"><w:footnote w:id="2"><w:p>',
          '<w:del w:id="1" w:author="A"><w:r><w:delText>cut from a footnote</w:delText></w:r></w:del>',
          "</w:p></w:footnote></w:footnotes>",
        ].join(""),
        "word/endnotes.xml": [
          '<w:endnotes xmlns:w="urn:w"><w:endnote w:id="2"><w:p>',
          '<w:ins w:id="2" w:author="A"><w:r><w:t>added to an endnote</w:t></w:r></w:ins>',
          "</w:p></w:endnote></w:endnotes>",
        ].join(""),
      }),
    );

    // Verified against Pandoc 3.10: a tracked deletion inside a footnote is
    // resolved away exactly like one in the body, so scanning only the document
    // part would let it vanish with a clean report.
    expect(report.scannedParts).toEqual([
      "word/document.xml",
      "word/footnotes.xml",
      "word/endnotes.xml",
    ]);
    expect(report.trackedChanges).toMatchObject({ insertions: 1, deletions: 1 });
  });

  it("reports SmartArt and charts, which live outside the media and embeddings prefixes", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/diagrams/data1.xml": '<dgm:dataModel xmlns:dgm="urn:d" />',
        "word/diagrams/layout1.xml": '<dgm:layoutDef xmlns:dgm="urn:d" />',
        "word/charts/chart1.xml": '<c:chartSpace xmlns:c="urn:c" />',
        "word/charts/style1.xml": '<cs:chartStyle xmlns:cs="urn:cs" />',
      }),
    );

    // One SmartArt graphic spans several parts, so the data part is what makes
    // it countable; the same holds for a chart.
    expect(findUnmappable(report, "smart-art")).toMatchObject({
      kind: "smart-art",
      count: 1,
      entries: ["word/diagrams/data1.xml", "word/diagrams/layout1.xml"],
    });
    expect(findUnmappable(report, "chart")).toMatchObject({
      kind: "chart",
      count: 1,
      entries: ["word/charts/chart1.xml", "word/charts/style1.xml"],
    });
  });

  it("reports altChunk imports, whose embedded document Pandoc never reads", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "word/document.xml": documentXml('<w:altChunk r:id="rId9" /><w:altChunk r:id="rId10" />'),
      }),
    );

    expect(findUnmappable(report, "external-content")).toMatchObject({
      kind: "external-content",
      count: 2,
      summary: expect.stringContaining("altChunk"),
    });
  });

  it("resolves the main document part from the package relationships", () => {
    // Verified against Pandoc 3.10: it converts this document fine, so assuming
    // word/document.xml would reject a valid Word file as "not a Word document".
    const report = buildDocxImportReport(
      createDocxArchive({
        "_rels/.rels": [
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document2.xml" />',
          "</Relationships>",
        ].join(""),
        "word/document2.xml": documentXml(
          '<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins></w:p>',
        ),
        "word/document.xml": undefined,
      }),
    );

    expect(report.documentPart).toBe("word/document2.xml");
    expect(report.scannedParts).toEqual(["word/document2.xml"]);
    expect(report.trackedChanges).toMatchObject({ insertions: 1 });
  });

  it("refuses to inflate an oversized part instead of exhausting memory", () => {
    // A small archive whose main part inflates past the per-part cap: a
    // compression ratio near 300:1. Without the cap an earlier build of this
    // shape allocated 242 MB of resident memory in 545 ms, scaling linearly
    // with the input. Sized just over the cap so the fixture stays cheap.
    const bomb = zipSync(
      { "word/document.xml": createRepetitiveXmlPart(DOCX_MAX_PART_BYTES + 1024) },
      { level: 1 },
    );

    expect(bomb.byteLength).toBeLessThan(2 * 1024 * 1024);
    expect(() => buildDocxImportReport(bomb)).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("over the"),
      }),
    );
  }, 30_000);

  it("refuses many individually plausible parts that exceed the total budget", () => {
    // The shape neither other limit catches: every part is honestly declared
    // and comfortably under the per-part cap, and the entry count is modest,
    // but the aggregate blows past the budget. Enumerating header parts is
    // what removed the fixed-part-set invariant that used to make the
    // per-part cap an aggregate cap, so this is the test that holds the
    // budget in place.
    //
    // The fixture is the least data that can prove that: just over the
    // budget, spread across the fewest parts, at the cheapest compression
    // level. It still costs ~1.3s because ~128 MB genuinely has to be
    // deflated, hence the explicit timeout below rather than silently
    // sitting near the default.
    const partCount = 10;
    const part = createRepetitiveXmlPart(Math.ceil(DOCX_MAX_TOTAL_BYTES / partCount) + 1024);
    const entries: Record<string, Uint8Array> = {
      "word/document.xml": strToU8(documentXml("<w:p />")),
    };

    for (let index = 1; index <= partCount; index += 1) {
      entries[`word/header${index}.xml`] = part;
    }

    // Every other guard must be satisfied, or this would prove nothing about
    // the budget specifically.
    expect(part.byteLength).toBeLessThan(DOCX_MAX_PART_BYTES);
    expect(partCount).toBeLessThanOrEqual(DOCX_MAX_HEADER_FOOTER_PARTS);
    expect(Object.keys(entries).length).toBeLessThan(DOCX_MAX_ARCHIVE_ENTRIES);
    expect(part.byteLength * partCount).toBeGreaterThan(DOCX_MAX_TOTAL_BYTES);

    expect(() => buildDocxImportReport(zipSync(entries, { level: 1 }))).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("in total"),
      }),
    );
  }, 30_000);

  it("refuses an implausible number of header and footer parts", () => {
    const entries: Record<string, Uint8Array> = {
      "word/document.xml": strToU8(documentXml("<w:p />")),
    };

    for (let index = 0; index <= DOCX_MAX_HEADER_FOOTER_PARTS; index += 1) {
      entries[`word/header${index}.xml`] = strToU8('<w:hdr xmlns:w="urn:w" />');
    }

    expect(() => buildDocxImportReport(zipSync(entries))).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("header and footer parts"),
      }),
    );
  });

  // Both levels matter. A stored entry is sliced using its COMPRESSED size, so
  // an understated uncompressed size is visible as a length mismatch. A deflate
  // entry — every real .docx, and anything an attacker would build — is
  // inflated into a buffer pre-sized from that same understated number, so the
  // inflated length always equals it and a length comparison is tautological.
  // Only the CRC catches the deflate case.
  it.each([
    ["stored", 0],
    ["deflate", 9],
  ])(
    "refuses a %s part whose entry understates its real size instead of analyzing a fragment",
    (_label, level) => {
      // The revision sits past byte 5000, so a 512-byte fragment hides it and
      // the report would come back clean for a document Pandoc reads in full.
      const body = documentXml(
        `<w:p><w:r><w:t>${"padding ".repeat(
          700,
        )}</w:t></w:r></w:p><w:ins w:id="1" w:author="A"><w:r><w:t>SECRET-INSERTION</w:t></w:r></w:ins>`,
      );
      const archive = zipSync({ "word/document.xml": strToU8(body) }, { level: level as 0 | 9 });
      const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
      let patched = false;

      // Rewrite the uncompressed-size field in both the local header and the
      // central directory record so the entry claims far less than it holds.
      for (let offset = 0; offset + 4 <= archive.byteLength; offset += 1) {
        const signature = view.getUint32(offset, true);

        if (signature === 0x0403_4b50) {
          view.setUint32(offset + 22, 512, true);
          patched = true;
        } else if (signature === 0x0201_4b50) {
          view.setUint32(offset + 24, 512, true);
          patched = true;
        }
      }

      expect(patched).toBe(true);
      expect(() => buildDocxImportReport(archive)).toThrow(
        expect.objectContaining({
          name: "PandocError",
          code: "DOCX_ARCHIVE_INVALID",
        }),
      );
    },
  );

  it("catches the deflate truncation by checksum, the only check that can", () => {
    // Pinning the mechanism, not just the rejection: for a deflate entry the
    // inflated length always equals the declared size, so if this ever starts
    // failing on the size check instead, the CRC has stopped doing the work.
    const body = documentXml(
      `<w:p><w:r><w:t>${"padding ".repeat(
        700,
      )}</w:t></w:r></w:p><w:ins w:id="1" w:author="A"><w:r><w:t>SECRET-INSERTION</w:t></w:r></w:ins>`,
    );
    const archive = zipSync({ "word/document.xml": strToU8(body) }, { level: 9 });
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);

    for (let offset = 0; offset + 4 <= archive.byteLength; offset += 1) {
      const signature = view.getUint32(offset, true);

      if (signature === 0x0403_4b50) {
        view.setUint32(offset + 22, 512, true);
      } else if (signature === 0x0201_4b50) {
        view.setUint32(offset + 24, 512, true);
      }
    }

    expect(() => buildDocxImportReport(archive)).toThrow(
      expect.objectContaining({
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("CRC-32"),
      }),
    );
  });

  it("accepts an honest archive at every compression level the CRC check sees", () => {
    // A checksum check that false-positives on real Word output would be worse
    // than the hole it closes, so both storage modes are asserted to pass.
    for (const level of [0, 6, 9] as const) {
      const archive = zipSync(
        {
          "word/document.xml": strToU8(documentXml("<w:p><w:r><w:t>Plain text.</w:t></w:r></w:p>")),
          "word/comments.xml": strToU8('<w:comments xmlns:w="urn:w" />'),
          "word/footnotes.xml": strToU8('<w:footnotes xmlns:w="urn:w" />'),
        },
        { level },
      );

      expect(buildDocxImportReport(archive).unmappable).toEqual([]);
    }
  });

  it("resolves a main document part named with single-quoted attributes", () => {
    const report = buildDocxImportReport(
      createDocxArchive({
        "_rels/.rels": [
          "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>",
          "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document2.xml' />",
          "</Relationships>",
        ].join(""),
        "word/document2.xml": documentXml("<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"),
        "word/document.xml": undefined,
      }),
    );

    expect(report.documentPart).toBe("word/document2.xml");
  });

  it("explains that the relationships named no usable target, not that a part is missing", () => {
    expect(() =>
      buildDocxImportReport(
        createDocxArchive({
          "_rels/.rels": [
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/absent.xml" />',
            "</Relationships>",
          ].join(""),
          "word/document.xml": undefined,
        }),
      ),
    ).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("no officeDocument target present in the archive"),
      }),
    );
  });

  it("refuses an archive with an implausible number of entries", () => {
    const entries: Record<string, Uint8Array> = {};

    for (let index = 0; index <= DOCX_MAX_ARCHIVE_ENTRIES; index += 1) {
      entries[`word/media/image${index}.png`] = new Uint8Array([0]);
    }

    expect(() => buildDocxImportReport(zipSync(entries))).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("entries"),
      }),
    );
  });

  it("surfaces every unmappable kind at once for a document that has all of them", () => {
    const report = buildDocxImportReport(createKitchenSinkDocx());

    expect(report.hasUnmappableContent).toBe(true);
    expect(report.unmappable.map((item) => item.kind)).toEqual([
      "tracked-change",
      "comment",
      "header-footer",
      "text-box",
      "external-content",
      "smart-art",
      "chart",
      "embedded-media",
      "embedded-object",
    ]);
  });

  it("stops reporting comments and revisions as losses when Pandoc keeps them as spans", () => {
    const report = buildDocxImportReport(createKitchenSinkDocx(), "all");

    // Verified against Pandoc 3.10: --track-changes=all emits insertion,
    // deletion, comment-start, and comment-end spans instead of dropping them,
    // so neither is a loss in that mode. The counts stay on the report, and
    // every loss that does not depend on the mode still reports.
    expect(report.unmappable.map((item) => item.kind)).toEqual([
      "header-footer",
      "text-box",
      "external-content",
      "smart-art",
      "chart",
      "embedded-media",
      "embedded-object",
    ]);
    expect(report.trackChanges).toBe("all");
    expect(report.trackedChanges).toMatchObject({ insertions: 1, deletions: 1 });
  });

  it.each([
    [
      "a manual page break, which is layout rather than content",
      documentXml('<w:p><w:r><w:br w:type="page" /></w:r></w:p>'),
    ],
    [
      "direct character formatting with no Markdown equivalent",
      documentXml(
        '<w:p><w:r><w:rPr><w:highlight w:val="yellow" /><w:spacing w:val="40" /></w:rPr><w:t>styled</w:t></w:r></w:p>',
      ),
    ],
  ])("does not claim to detect %s", (_label, body) => {
    const report = buildDocxImportReport(createDocxArchive({ "word/document.xml": body }));

    // These are real fidelity losses that this report deliberately does not
    // cover. The test exists so the gap stays a decision rather than becoming
    // an accidental promise that an empty `unmappable` means a lossless import.
    expect(report.unmappable).toEqual([]);
  });

  it("reports insertions as discarded when revisions are rejected", () => {
    const report = buildDocxImportReport(createKitchenSinkDocx(), "reject");

    expect(findUnmappable(report, "tracked-change")?.summary).toContain(
      "deletions were restored as ordinary text and insertions were discarded",
    );
  });

  it.each([
    ["empty bytes", new Uint8Array()],
    ["plain text that is not a ZIP", strToU8("# This is Markdown, not a DOCX")],
    ["a legacy .doc OLE compound file header", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1])],
  ])("rejects %s as invalid DOCX input", (_label, bytes) => {
    expect(() => buildDocxImportReport(bytes)).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_INPUT_INVALID",
      }),
    );
  });

  it("rejects a ZIP archive that is not a Word document", () => {
    expect(() => buildDocxImportReport(zipSync({ mimetype: strToU8("text/plain") }))).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
        message: expect.stringContaining("word/document.xml"),
      }),
    );
  });

  it("rejects a truncated ZIP archive", () => {
    const truncated = createDocxArchive().slice(0, 40);

    expect(() => buildDocxImportReport(truncated)).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_ARCHIVE_INVALID",
      }),
    );
  });

  it.each(["", "ACCEPT", "yes", 1, null])("rejects invalid trackChanges value %j", (mode) => {
    expect(() => buildDocxImportReport(createDocxArchive(), mode as unknown as "accept")).toThrow(
      expect.objectContaining({
        name: "PandocError",
        code: "DOCX_TRACK_CHANGES_INVALID",
      }),
    );
  });

  it("exposes the supported track-changes modes and the clean-Markdown default", () => {
    expect(DOCX_TRACK_CHANGES_MODES).toEqual(["accept", "reject", "all"]);
    expect(DEFAULT_DOCX_TRACK_CHANGES).toBe("accept");
  });
});

describe("DOCX to Markdown conversion", () => {
  it("fails closed when Pandoc is unsupported before touching the DOCX input", async () => {
    const runner: PandocRunner = vi.fn(async () => ({
      stdout: "pandoc 2.19",
      stderr: "",
      exitCode: 0,
    }));

    await expect(
      convertDocxToMarkdown({ docx: createDocxArchive(), runner }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_UNSUPPORTED_MAJOR",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it("uses argument arrays with an explicit track-changes value and no writer env", async () => {
    const calls: Array<{
      binary: string;
      args: readonly string[];
      options: PandocRunnerOptions | undefined;
    }> = [];
    const docx = createDocxArchive();
    const runner: PandocRunner = vi.fn(async (binary, args, options) => {
      calls.push({ binary, args, options });

      if (args[0] === "--version") {
        return { stdout: "pandoc 3.10\nFeatures: +lua", stderr: "", exitCode: 0 };
      }

      const outputPath = args[args.indexOf("--output") + 1];
      const inputPath = args.at(-1);

      if (outputPath === undefined || inputPath === undefined) {
        throw new Error("test runner received an incomplete invocation");
      }

      // The bytes handed to Pandoc are the bytes the caller passed in.
      expect(new Uint8Array(await readFile(inputPath))).toEqual(docx);
      await writeFile(outputPath, "# Converted\n");

      return { stdout: "", stderr: "", exitCode: 0 };
    });

    const result = await convertDocxToMarkdown({ docx, runner });

    expect(result.markdown).toBe("# Converted\n");
    expect(calls).toHaveLength(2);
    expect(calls[1]).toEqual({
      binary: "pandoc",
      args: [
        "--from",
        "docx",
        "--to",
        "gfm",
        "--track-changes",
        "accept",
        "--output",
        expect.stringMatching(/output\.md$/),
        expect.stringMatching(/input\.docx$/),
      ],
      options: { reject: false },
    });
    // SOURCE_DATE_EPOCH only matters to the DOCX writer; Markdown output has no
    // embedded timestamps, so the reverse path passes no environment at all.
    expect(calls[1]?.options?.env).toBeUndefined();
  });

  it("passes an explicitly requested track-changes value through to Pandoc", async () => {
    const runner = createReverseRunner("# Converted\n");

    const result = await convertDocxToMarkdown({
      docx: createKitchenSinkDocx(),
      trackChanges: "all",
      runner,
    });

    expect(runner).toHaveBeenLastCalledWith(
      "pandoc",
      expect.arrayContaining(["--track-changes", "all"]),
      { reject: false },
    );
    expect(result.report.trackChanges).toBe("all");
  });

  it("rejects an invalid track-changes value before writing any conversion files", async () => {
    const runner = createSuccessfulDoctorRunner();

    await expect(
      convertDocxToMarkdown({
        docx: createDocxArchive(),
        trackChanges: "maybe" as unknown as "accept",
        runner,
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "DOCX_TRACK_CHANGES_INVALID",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it("runs postprocessors in order with reverse Markdown context and returns the report", async () => {
    const seen: string[] = [];
    const result = await convertDocxToMarkdown({
      docx: createKitchenSinkDocx(),
      metadata: { source: "fixture" },
      postprocessors: [
        (markdown, context) => {
          seen.push(`${context.phase}:${context.sourceFormat}:${context.targetFormat}`);
          return `${markdown}first\n`;
        },
        async (markdown, context) => {
          seen.push(`${context.phase}:${String(context.metadata.source)}`);
          return `${markdown}second\n`;
        },
      ],
      runner: createReverseRunner("# Converted\n"),
    });

    expect(result.markdown).toBe("# Converted\nfirst\nsecond\n");
    expect(seen).toEqual(["postprocess:markdown:markdown", "postprocess:fixture"]);
    // The report describes the DOCX, so postprocessors cannot mask a loss.
    expect(result.report.hasUnmappableContent).toBe(true);
  });

  it("wraps postprocessor failures in typed hook errors", async () => {
    await expect(
      convertDocxToMarkdown({
        docx: createDocxArchive(),
        postprocessors: [
          () => {
            throw new Error("sidecar lookup exploded");
          },
        ],
        runner: createReverseRunner("# Converted\n"),
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "MARKDOWN_HOOK_FAILED",
    });
  });

  it("reads a DOCX from a filesystem path", async () => {
    const docxPath = join(await createTempDirectory(), "note.docx");
    await writeFile(docxPath, createDocxArchive());

    const result = await convertDocxToMarkdown({
      docx: docxPath,
      runner: createReverseRunner("# From path\n"),
    });

    expect(result.markdown).toBe("# From path\n");
  });

  it.each([
    ["a missing path", "tests/fixtures/golden/missing-note.docx"],
    ["a blank path", "   "],
  ])("rejects %s with a typed input error", async (_label, docx) => {
    const runner = createSuccessfulDoctorRunner();

    await expect(convertDocxToMarkdown({ docx, runner })).rejects.toMatchObject({
      name: "PandocError",
      code: "DOCX_INPUT_INVALID",
    });
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["garbage that is not a ZIP", strToU8("not a docx"), "DOCX_INPUT_INVALID"],
    [
      "a ZIP that is not a Word document",
      zipSync({ "a.txt": strToU8("hi") }),
      "DOCX_ARCHIVE_INVALID",
    ],
  ])("rejects %s before invoking Pandoc's converter", async (_label, docx, code) => {
    const runner = createSuccessfulDoctorRunner();

    await expect(convertDocxToMarkdown({ docx, runner })).rejects.toMatchObject({
      name: "PandocError",
      code,
    });
    // Only the version probe ran, so garbage never reached the converter.
    expect(runner).toHaveBeenCalledTimes(1);
  });

  it("cleans up temporary files when Pandoc conversion fails", async () => {
    let tempOutputPath: string | undefined;
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return { stdout: "pandoc 3.10", stderr: "", exitCode: 0 };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];

      return { stdout: "", stderr: "unsupported docx feature", exitCode: 43 };
    });

    await expect(
      convertDocxToMarkdown({ docx: createDocxArchive(), runner }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_CONVERSION_FAILED",
      guidance: expect.stringContaining("Pandoc stderr: unsupported docx feature"),
    });

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("wraps missing Markdown output after exit 0 and cleans up", async () => {
    let tempOutputPath: string | undefined;
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return { stdout: "pandoc 3.10", stderr: "", exitCode: 0 };
      }

      tempOutputPath = args[args.indexOf("--output") + 1];

      return { stdout: "", stderr: "", exitCode: 0 };
    });

    await expect(
      convertDocxToMarkdown({ docx: createDocxArchive(), runner }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_CONVERSION_FAILED",
      message: expect.stringContaining("Markdown output could not be read"),
    });

    expect(tempOutputPath).toBeDefined();
    await expect(stat(join(tempOutputPath ?? "", ".."))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("snapshots DOCX bytes so caller mutation cannot desync the report from the output", async () => {
    let convertedBytes: Uint8Array | undefined;
    const docx = createKitchenSinkDocx();
    const runner: PandocRunner = vi.fn(async (_binary, args) => {
      if (args[0] === "--version") {
        return { stdout: "pandoc 3.10", stderr: "", exitCode: 0 };
      }

      const outputPath = args[args.indexOf("--output") + 1];
      const inputPath = args.at(-1);

      if (outputPath === undefined || inputPath === undefined) {
        throw new Error("test runner received an incomplete invocation");
      }

      convertedBytes = new Uint8Array(await readFile(inputPath));
      await writeFile(outputPath, "# Converted\n");

      return { stdout: "", stderr: "", exitCode: 0 };
    });
    const original = Uint8Array.from(docx);

    const conversion = convertDocxToMarkdown({ docx, runner });
    docx.fill(0);
    const result = await conversion;

    expect(convertedBytes).toEqual(original);
    expect(result.report.hasUnmappableContent).toBe(true);
  });

  it("wraps unexpected runner failures and preserves typed ones", async () => {
    const typedError = new PandocError(
      "PANDOC_CONVERSION_FAILED",
      "Injected typed failure.",
      "Caller guidance.",
    );
    const throwingRunner = (cause: unknown): PandocRunner =>
      vi.fn(async (_binary, args) => {
        if (args[0] === "--version") {
          return { stdout: "pandoc 3.10", stderr: "", exitCode: 0 };
        }

        throw cause;
      });

    await expect(
      convertDocxToMarkdown({
        docx: createDocxArchive(),
        runner: throwingRunner(new Error("process runner exploded")),
      }),
    ).rejects.toMatchObject({
      name: "PandocError",
      code: "PANDOC_CONVERSION_FAILED",
      guidance: expect.stringContaining("pandocPath"),
    });
    await expect(
      convertDocxToMarkdown({ docx: createDocxArchive(), runner: throwingRunner(typedError) }),
    ).rejects.toBe(typedError);
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

  it("round-trips the golden note back to the golden Markdown fixture", async () => {
    const markdown = await readFile(fixtureInputPath, "utf8");
    const expectedMarkdown = await readFile(fixtureRoundTripExpectedPath, "utf8");
    const docx = await convertMarkdownToDocx({
      markdown,
      referenceDocxPath: fixtureReferenceDocxPath,
      sourceDateEpoch: 1_704_067_200,
      preprocessors: [stripYamlFrontmatter],
    });

    const result = await convertDocxToMarkdown({ docx });

    expect(result.markdown).toBe(expectedMarkdown);
    // A Pandoc-written DOCX carries an empty comments part and no media, so a
    // clean publish round trip must report no losses at all.
    expect(result.report).toEqual({
      unmappable: [],
      hasUnmappableContent: false,
      trackChanges: "accept",
      trackedChanges: { insertions: 0, deletions: 0, formattingChanges: 0 },
      documentPart: "word/document.xml",
      scannedParts: ["word/document.xml", "word/footnotes.xml"],
    });
  });

  it("documents the round-trip losses the golden fixture records", async () => {
    const source = await readFile(fixtureInputPath, "utf8");
    const roundTripped = await readFile(fixtureRoundTripExpectedPath, "utf8");

    // These assertions exist so the golden file cannot be quietly regenerated
    // into something that hides a loss: the round trip really does drop
    // frontmatter and really does destroy wikilink syntax.
    expect(source).toContain("title: Basic TeamWiki Note");
    expect(roundTripped).not.toContain("title: Basic TeamWiki Note");
    expect(source).toContain("[[TeamWiki]]");
    expect(roundTripped).not.toContain("[[TeamWiki]]");
    expect(roundTripped).toContain(String.raw`\[\[TeamWiki\]\]`);
  });

  it("restores frontmatter and wikilinks from a consumer-provided sidecar", async () => {
    const sourceNote = [
      "---",
      "title: Sidecar Note",
      "tags:",
      "  - teamwiki",
      "---",
      "",
      "# Sidecar Note",
      "",
      "This note links to [[TeamWiki]] and [[Handbook]].",
      "",
      "A second paragraph with **bold** text.",
      "",
      "- First item",
      "- Second item",
      "",
    ].join("\n");
    // This note is deliberately built to survive Pandoc's re-wrap: short lines,
    // no callout, no paragraph long enough to be reflowed. That is what makes a
    // byte-exact assertion possible. The test proves the postprocessor hook
    // runs and can restore the Obsidian layer, NOT that byte-exact restoration
    // is achievable for an arbitrary note — the golden round-trip fixture next
    // to it shows how much a realistic note actually loses.
    //
    // The sidecar is the consumer's, not the library's: publishing records what
    // the DOCX cannot carry, and the reverse import puts it back.
    const sidecar: { frontmatter: string; wikilinks: string[] } = {
      frontmatter: "",
      wikilinks: [],
    };

    const docx = await convertMarkdownToDocx({
      markdown: sourceNote,
      referenceDocxPath: fixtureReferenceDocxPath,
      sourceDateEpoch: 1_704_067_200,
      preprocessors: [
        (markdown) => {
          const match = /^---\n[\s\S]*?\n---\n\n?/.exec(markdown);
          sidecar.frontmatter = match?.[0] ?? "";

          return markdown.slice(sidecar.frontmatter.length);
        },
        (markdown) =>
          markdown.replace(/\[\[([^\]]+)\]\]/g, (_match, target: string) => {
            sidecar.wikilinks.push(target);

            return target;
          }),
      ],
    });

    const result = await convertDocxToMarkdown({
      docx,
      postprocessors: [
        (markdown) => {
          let restored = markdown;

          // replaceAll, not replace: a string first argument to replace()
          // rewrites only the first match, so a target that also appears
          // earlier as ordinary prose would capture the wrapping and leave the
          // real wikilink bare. Deduplicating targets keeps a repeated link
          // from being wrapped twice.
          //
          // A production consumer needs more than this: a bare target string is
          // indistinguishable from prose that happens to use the same words, so
          // the sidecar should record source offsets rather than just names.
          for (const target of new Set(sidecar.wikilinks)) {
            restored = restored.replaceAll(target, `[[${target}]]`);
          }

          return restored;
        },
        (markdown) => `${sidecar.frontmatter}${markdown}`,
      ],
    });

    expect(sidecar.wikilinks).toEqual(["TeamWiki", "Handbook"]);
    expect(result.markdown).toBe(sourceNote);
    expect(result.report.hasUnmappableContent).toBe(false);
  });

  it("surfaces every unmappable construct from a real DOCX", async () => {
    const result = await convertDocxToMarkdown({ docx: createKitchenSinkDocx() });

    expect(result.report.unmappable.map((item) => item.kind)).toEqual([
      "tracked-change",
      "comment",
      "header-footer",
      "text-box",
      "external-content",
      "smart-art",
      "chart",
      "embedded-media",
      "embedded-object",
    ]);
    // These assertions pin the behavior each summary claims, so a Pandoc
    // upgrade that changes any of it fails here rather than leaving the report
    // quietly wrong.
    //
    // Accepted: the insertion became ordinary text, the deletion is gone.
    expect(result.markdown).toContain("INSERTED-TEXT");
    expect(result.markdown).not.toContain("DELETED-TEXT");
    // Dropped outright, each with no placeholder of any kind.
    expect(result.markdown).not.toContain("First reviewer comment.");
    expect(result.markdown).not.toContain("EMBEDDED-OLE-OBJECT");
    expect(result.markdown).not.toContain("HEADER-TEXT-CONTENT");
    expect(result.markdown).not.toContain("SMARTART-NODE-TEXT");
    expect(result.markdown).not.toContain("CHART-TITLE-TEXT");
    expect(result.markdown).not.toContain("ALTCHUNK-IMPORTED-TEXT");
    // Inlined, losing its text-box framing rather than its text.
    expect(result.markdown).toContain("TEXTBOX-CONTENT");
    // Referenced but never written: the image link dangles.
    expect(result.markdown).toContain("media/image1.png");
  });

  it.each([
    [
      "a DrawingML text box with the mc:Fallback Word 2010+ writes",
      "<w:p><w:r><mc:AlternateContent>" +
        '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="2000000" cy="500000" /><wp:docPr id="9" name="Text Box 9" /><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1" /><wps:spPr><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>TEXTBOX-PROBE</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" /></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
        '<mc:Fallback><w:pict><v:shape id="_x0000_s2051" type="#_x0000_t202" style="width:150pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>TEXTBOX-PROBE</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>' +
        "</mc:AlternateContent></w:r></w:p>",
      1,
      0,
      true,
    ],
    [
      "a DrawingML text box with no fallback",
      "<w:p><w:r><mc:AlternateContent>" +
        '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="2000000" cy="500000" /><wp:docPr id="9" name="Text Box 9" /><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1" /><wps:spPr><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>TEXTBOX-PROBE</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" /></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
        "</mc:AlternateContent></w:r></w:p>",
      0,
      1,
      false,
    ],
    [
      "a bare legacy VML text box",
      '<w:p><w:r><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="width:150pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>TEXTBOX-PROBE</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>',
      1,
      0,
      true,
    ],
  ])(
    "classifies %s the way Pandoc actually treats it",
    async (_label, body, inlined, dropped, textSurvives) => {
      const result = await convertDocxToMarkdown({
        docx: createDocxArchive({ "word/document.xml": documentXml(body) }),
      });

      // The report's classification and Pandoc's real behavior are asserted
      // against each other. This case is what caught the classification being
      // inverted: the fallback, not the namespace, decides.
      expect(findUnmappable(result.report, "text-box")?.summary).toContain(
        `${inlined} inlined as ordinary body paragraphs`,
      );
      expect(findUnmappable(result.report, "text-box")?.summary).toContain(
        `${dropped} dropped outright`,
      );
      expect(result.markdown.includes("TEXTBOX-PROBE")).toBe(textSurvives);
    },
  );

  it("reports a header that Pandoc drops, so an empty report never hides one", async () => {
    const result = await convertDocxToMarkdown({
      docx: createDocxArchive({
        "word/header1.xml": `<w:hdr ${DOCX_NAMESPACES}><w:p><w:r><w:t>HEADER-TEXT-CONTENT</w:t></w:r></w:p></w:hdr>`,
      }),
    });

    expect(result.markdown).not.toContain("HEADER-TEXT-CONTENT");
    expect(findUnmappable(result.report, "header-footer")).toMatchObject({
      count: 1,
      entries: ["word/header1.xml"],
    });
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

const DOCX_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><w:document ${DOCX_NAMESPACES}><w:body>${body}</w:body></w:document>`;
}

/**
 * Builds a real DOCX archive in memory. The default parts are the minimum a
 * Word document needs, so tests can add exactly the one construct they are
 * about without hand-rolling a package each time.
 */
function createDocxArchive(
  parts: Readonly<Record<string, string | Uint8Array | undefined>> = {},
): Uint8Array {
  const defaults: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />',
      '<Default Extension="xml" ContentType="application/xml" />',
      '<Default Extension="png" ContentType="image/png" />',
      '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject" />',
      '<Default Extension="html" ContentType="text/html" />',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" />',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />',
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" />',
      "</Types>",
    ].join(""),
    "_rels/.rels": [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />',
      "</Relationships>",
    ].join(""),
    "word/_rels/document.xml.rels": [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" />',
      '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png" />',
      '<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin" />',
      '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml" />',
      '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml" />',
      '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml" />',
      '<Relationship Id="rId62" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.html" />',
      "</Relationships>",
    ].join(""),
    "word/document.xml": documentXml("<w:p><w:r><w:t>Plain text.</w:t></w:r></w:p>"),
  };
  const files: Record<string, Uint8Array> = {};

  for (const [name, content] of Object.entries({ ...defaults, ...parts })) {
    // An explicit undefined removes a default part, so a test can build an
    // archive that deliberately lacks one.
    if (content !== undefined) {
      files[name] = typeof content === "string" ? strToU8(content) : content;
    }
  }

  return zipSync(files);
}

/**
 * A well-formed, highly repetitive XML part of at least `minimumBytes`.
 *
 * The archive-limit tests have to hand the inspector real, honestly declared
 * data to be worth anything, so their fixtures are sized from the limit they
 * probe rather than from a hardcoded number. That keeps them minimal — building
 * one is dominated by deflating the bytes, so 4x more data is 4x the runtime —
 * and keeps them correct if a limit is ever retuned.
 */
function createRepetitiveXmlPart(minimumBytes: number): Uint8Array {
  const unit = "<w:p><w:r><w:t>Header padding text.</w:t></w:r></w:p>";
  const open = '<w:hdr xmlns:w="urn:w">';
  const close = "</w:hdr>";
  const repeats = Math.ceil((minimumBytes - open.length - close.length) / unit.length);

  return strToU8(`${open}${unit.repeat(Math.max(repeats, 1))}${close}`);
}

/** A DOCX containing one of every construct the import report knows about. */
function createKitchenSinkDocx(): Uint8Array {
  return createDocxArchive({
    "word/document.xml": documentXml(
      [
        "<w:p>",
        '<w:r><w:t xml:space="preserve">Base paragraph. </w:t></w:r>',
        '<w:ins w:id="501" w:author="Reviewer One" w:date="2024-01-01T00:00:00Z"><w:r><w:t>INSERTED-TEXT</w:t></w:r></w:ins>',
        '<w:del w:id="502" w:author="Reviewer One" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>DELETED-TEXT</w:delText></w:r></w:del>',
        "</w:p>",
        "<w:p>",
        '<w:commentRangeStart w:id="1" /><w:r><w:t>Commented sentence.</w:t></w:r><w:commentRangeEnd w:id="1" />',
        '<w:r><w:commentReference w:id="1" /></w:r>',
        "</w:p>",
        "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>TEXTBOX-CONTENT</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>",
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500" /><wp:docPr id="1" name="Picture 1" /><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png" /><pic:cNvPicPr /></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId100" /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="952500" cy="952500" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
        '<w:p><w:r><w:object><v:shape id="_x0000_i1025" type="#_x0000_t75" /><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rId101" /></w:object></w:r></w:p>',
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="2000000" /><wp:docPr id="20" name="Diagram" /><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId60" /></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="2000000" /><wp:docPr id="21" name="Chart" /><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId61" /></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
        '<w:altChunk r:id="rId62" />',
        "<w:p><w:r><w:t>Trailing paragraph.</w:t></w:r></w:p>",
      ].join(""),
    ),
    "word/comments.xml": [
      `<?xml version="1.0" encoding="UTF-8"?><w:comments ${DOCX_NAMESPACES}>`,
      '<w:comment w:id="1" w:author="Reviewer One" w:date="2024-01-01T00:00:00Z" w:initials="R1">',
      "<w:p><w:r><w:t>First reviewer comment.</w:t></w:r></w:p></w:comment>",
      "</w:comments>",
    ].join(""),
    "word/header1.xml": `<w:hdr ${DOCX_NAMESPACES}><w:p><w:r><w:t>HEADER-TEXT-CONTENT</w:t></w:r></w:p></w:hdr>`,
    "word/diagrams/data1.xml":
      '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>SMARTART-NODE-TEXT</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>',
    "word/charts/chart1.xml":
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>CHART-TITLE-TEXT</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>',
    "word/afchunk.html": "<html><body><p>ALTCHUNK-IMPORTED-TEXT</p></body></html>",
    "word/media/image1.png": TINY_PNG,
    "word/embeddings/oleObject1.bin": strToU8("EMBEDDED-OLE-OBJECT"),
  });
}

function findUnmappable(
  report: DocxImportReport,
  kind: DocxUnmappableKind,
): DocxUnmappableItem | undefined {
  return report.unmappable.find((item) => item.kind === kind);
}

/** A runner that answers the version probe, then writes fixed Markdown output. */
function createReverseRunner(markdown: string): PandocRunner {
  return vi.fn(async (_binary, args) => {
    if (args[0] === "--version") {
      return { stdout: "pandoc 3.10", stderr: "", exitCode: 0 };
    }

    const outputPath = args[args.indexOf("--output") + 1];

    if (outputPath === undefined) {
      throw new Error("test runner received no output path");
    }

    await writeFile(outputPath, markdown);

    return { stdout: "", stderr: "", exitCode: 0 };
  });
}

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
