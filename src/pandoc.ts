import { constants } from "node:fs";
import { access, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, extname, join } from "node:path";
import { execa } from "execa";

export const SUPPORTED_PANDOC_MAJOR = 3;
export const DEFAULT_SOURCE_DATE_EPOCH = "0";

const DEFAULT_PANDOC_BINARY = "pandoc";

export type PandocDoctorFailureCode =
  | "PANDOC_NOT_FOUND"
  | "PANDOC_PROBE_FAILED"
  | "PANDOC_VERSION_UNPARSEABLE"
  | "PANDOC_UNSUPPORTED_MAJOR";

export type PandocErrorCode =
  | PandocDoctorFailureCode
  | "MARKDOWN_HOOK_FAILED"
  | "PANDOC_CONVERSION_FAILED"
  | "REFERENCE_DOC_REQUIRED"
  | "REFERENCE_DOC_INVALID"
  | "SOURCE_DATE_EPOCH_INVALID";

export interface PandocVersion {
  readonly major: number;
  readonly minor: number | undefined;
  readonly patch: number | undefined;
  readonly raw: string;
}

export interface PandocDoctorSuccess {
  readonly ok: true;
  readonly binary: string;
  readonly supportedMajor: number;
  readonly version: PandocVersion;
  readonly features: readonly string[];
  readonly rawOutput: string;
}

export interface PandocDoctorFailure {
  readonly ok: false;
  readonly binary: string;
  readonly code: PandocDoctorFailureCode;
  readonly message: string;
  readonly guidance: string;
  readonly supportedMajor: number;
  readonly detectedVersion?: PandocVersion;
  readonly rawOutput?: string;
  readonly cause?: unknown;
}

export type PandocDoctorResult = PandocDoctorSuccess | PandocDoctorFailure;

export interface PandocRunnerOptions {
  readonly env?: Readonly<Record<string, string>>;
  readonly reject?: boolean;
}

export interface PandocRunnerResult {
  readonly stdout: string;
  readonly stderr: string;
  readonly exitCode: number;
}

export type PandocRunner = (
  binary: string,
  args: readonly string[],
  options?: PandocRunnerOptions,
) => Promise<PandocRunnerResult>;

export interface PandocDoctorOptions {
  readonly pandocPath?: string;
  readonly runner?: PandocRunner;
}

export interface MarkdownProcessorContext {
  readonly phase: "preprocess" | "postprocess";
  readonly sourceFormat: "markdown";
  readonly targetFormat: "docx" | "markdown";
  readonly metadata: Readonly<Record<string, unknown>>;
}

export type MarkdownProcessor = (
  markdown: string,
  context: MarkdownProcessorContext,
) => string | Promise<string>;

export type MarkdownPreprocessor = MarkdownProcessor;

/**
 * Typed contract for future reverse/textual Markdown pipelines.
 *
 * Markdown-to-DOCX conversion does not run postprocessors because the forward
 * pipeline returns DOCX bytes, not Markdown text.
 */
export type MarkdownPostprocessor = MarkdownProcessor;

export interface ConvertMarkdownToDocxOptions extends PandocDoctorOptions {
  readonly markdown: string | Uint8Array;
  readonly referenceDocxPath: string;
  readonly sourceDateEpoch?: number | string;
  readonly preprocessors?: readonly MarkdownPreprocessor[];
  readonly metadata?: Readonly<Record<string, unknown>>;
}

export class PandocError extends Error {
  readonly code: PandocErrorCode;
  readonly guidance: string;

  constructor(
    code: PandocErrorCode,
    message: string,
    guidance: string,
    options: { readonly cause?: unknown } = {},
  ) {
    const fullMessage = `${message} ${guidance}`;

    if ("cause" in options) {
      super(fullMessage, { cause: options.cause });
    } else {
      super(fullMessage);
    }

    this.name = "PandocError";
    this.code = code;
    this.guidance = guidance;
  }
}

export async function doctor(options: PandocDoctorOptions = {}): Promise<PandocDoctorResult> {
  const binary = options.pandocPath ?? DEFAULT_PANDOC_BINARY;
  const runner = options.runner ?? defaultPandocRunner;

  try {
    const result = await runner(binary, ["--version"], { reject: false });

    if (result.exitCode !== 0) {
      return {
        ok: false,
        binary,
        code: "PANDOC_PROBE_FAILED",
        message: "Pandoc was found but `pandoc --version` failed.",
        guidance:
          "Run `pandoc --version` locally and fix the installation before converting Markdown to DOCX.",
        supportedMajor: SUPPORTED_PANDOC_MAJOR,
        rawOutput: joinOutput(result.stdout, result.stderr),
      };
    }

    const rawOutput = joinOutput(result.stdout, result.stderr);
    const version = parsePandocVersion(rawOutput);

    if (version === undefined) {
      return {
        ok: false,
        binary,
        code: "PANDOC_VERSION_UNPARSEABLE",
        message: "Pandoc responded, but its version could not be parsed.",
        guidance: `Install a supported Pandoc ${SUPPORTED_PANDOC_MAJOR}.x release and retry.`,
        supportedMajor: SUPPORTED_PANDOC_MAJOR,
        rawOutput,
      };
    }

    if (version.major !== SUPPORTED_PANDOC_MAJOR) {
      return {
        ok: false,
        binary,
        code: "PANDOC_UNSUPPORTED_MAJOR",
        message: `Pandoc ${version.raw} is installed, but this package supports Pandoc ${SUPPORTED_PANDOC_MAJOR}.x.`,
        guidance: `Install or select Pandoc ${SUPPORTED_PANDOC_MAJOR}.x with the pandocPath option before converting.`,
        supportedMajor: SUPPORTED_PANDOC_MAJOR,
        detectedVersion: version,
        rawOutput,
      };
    }

    return {
      ok: true,
      binary,
      supportedMajor: SUPPORTED_PANDOC_MAJOR,
      version,
      features: parsePandocFeatures(rawOutput),
      rawOutput,
    };
  } catch (cause) {
    return {
      ok: false,
      binary,
      code: isNotFoundError(cause) ? "PANDOC_NOT_FOUND" : "PANDOC_PROBE_FAILED",
      message: isNotFoundError(cause)
        ? "Pandoc was not found on PATH."
        : "Pandoc could not be probed.",
      guidance:
        "Install Pandoc 3.x from https://pandoc.org/installing.html or pass pandocPath to the installed binary.",
      supportedMajor: SUPPORTED_PANDOC_MAJOR,
      cause,
    };
  }
}

export async function applyMarkdownPreprocessors(
  markdown: string,
  preprocessors: readonly MarkdownPreprocessor[] = [],
  metadata: Readonly<Record<string, unknown>> = {},
): Promise<string> {
  return applyMarkdownProcessors(markdown, preprocessors, {
    phase: "preprocess",
    sourceFormat: "markdown",
    targetFormat: "docx",
    metadata,
  });
}

export async function applyMarkdownPostprocessors(
  markdown: string,
  postprocessors: readonly MarkdownPostprocessor[] = [],
  metadata: Readonly<Record<string, unknown>> = {},
): Promise<string> {
  return applyMarkdownProcessors(markdown, postprocessors, {
    phase: "postprocess",
    sourceFormat: "markdown",
    targetFormat: "markdown",
    metadata,
  });
}

export async function convertMarkdownToDocx(
  options: ConvertMarkdownToDocxOptions,
): Promise<Uint8Array> {
  const pandoc = await doctor(options);

  if (!pandoc.ok) {
    throw pandocFailureToError(pandoc);
  }

  const referenceDocxPath = await validateReferenceDocxPath(options.referenceDocxPath);
  const sourceDateEpoch = normalizeSourceDateEpoch(options.sourceDateEpoch);
  const markdown = await applyMarkdownPreprocessors(
    decodeMarkdown(options.markdown),
    options.preprocessors,
    options.metadata,
  );
  const runner = options.runner ?? defaultPandocRunner;
  const tempDirectory = await mkdtemp(join(tmpdir(), "polydoc-core-"));

  try {
    const inputPath = join(tempDirectory, "input.md");
    const outputPath = join(tempDirectory, "output.docx");
    await writeFile(inputPath, markdown, "utf8");

    let result: PandocRunnerResult;

    try {
      result = await runner(
        pandoc.binary,
        [
          "--from",
          "gfm",
          "--to",
          "docx",
          "--reference-doc",
          referenceDocxPath,
          "--output",
          outputPath,
          inputPath,
        ],
        {
          env: { SOURCE_DATE_EPOCH: sourceDateEpoch },
          reject: false,
        },
      );
    } catch (cause) {
      if (cause instanceof PandocError) {
        throw cause;
      }

      throw new PandocError(
        "PANDOC_CONVERSION_FAILED",
        "Pandoc execution failed while converting Markdown to DOCX.",
        "Confirm pandocPath points at a working Pandoc 3.x binary and retry the conversion.",
        { cause },
      );
    }

    if (result.exitCode !== 0) {
      throw new PandocError(
        "PANDOC_CONVERSION_FAILED",
        "Pandoc failed while converting Markdown to DOCX.",
        conversionGuidance(result.stderr),
      );
    }

    try {
      return await readFile(outputPath);
    } catch (cause) {
      throw new PandocError(
        "PANDOC_CONVERSION_FAILED",
        "Pandoc completed but the DOCX output could not be read.",
        "Check that the conversion process can write to the system temp directory and that disk space is available.",
        { cause },
      );
    }
  } finally {
    await rm(tempDirectory, { force: true, recursive: true });
  }
}

async function defaultPandocRunner(
  binary: string,
  args: readonly string[],
  options: PandocRunnerOptions = {},
): Promise<PandocRunnerResult> {
  const execaOptions =
    options.env === undefined
      ? { reject: options.reject ?? true }
      : { env: { ...options.env }, reject: options.reject ?? true };
  const result = await execa(binary, [...args], execaOptions);

  return {
    stdout: result.stdout,
    stderr: result.stderr,
    exitCode: result.exitCode ?? 0,
  };
}

async function applyMarkdownProcessors(
  markdown: string,
  processors: readonly MarkdownProcessor[],
  context: MarkdownProcessorContext,
): Promise<string> {
  let current = markdown;

  for (const processor of processors) {
    try {
      const next = await processor(current, context);

      if (typeof next !== "string") {
        throw new TypeError("Markdown processors must return a string.");
      }

      current = next;
    } catch (cause) {
      if (cause instanceof PandocError) {
        throw cause;
      }

      throw new PandocError(
        "MARKDOWN_HOOK_FAILED",
        `A Markdown ${context.phase}or failed.`,
        "Fix the configured Markdown hook or remove it before retrying the conversion.",
        { cause },
      );
    }
  }

  return current;
}

async function validateReferenceDocxPath(referenceDocxPath: string): Promise<string> {
  if (referenceDocxPath.trim() === "") {
    throw new PandocError(
      "REFERENCE_DOC_REQUIRED",
      "A reference DOCX path is required for Markdown-to-DOCX conversion.",
      "Pass referenceDocxPath pointing at a readable .docx file that defines the Word styling contract.",
    );
  }

  if (extname(referenceDocxPath).toLowerCase() !== ".docx") {
    throw new PandocError(
      "REFERENCE_DOC_INVALID",
      `Reference document ${basename(referenceDocxPath)} is not a .docx file.`,
      "Pass a readable .docx file through referenceDocxPath.",
    );
  }

  try {
    const referenceStat = await stat(referenceDocxPath);

    if (!referenceStat.isFile()) {
      throw new PandocError(
        "REFERENCE_DOC_INVALID",
        `Reference document ${basename(referenceDocxPath)} is not a file.`,
        "Pass a readable .docx file through referenceDocxPath.",
      );
    }

    await access(referenceDocxPath, constants.R_OK);
  } catch (cause) {
    if (cause instanceof PandocError) {
      throw cause;
    }

    throw new PandocError(
      "REFERENCE_DOC_INVALID",
      `Reference document ${basename(referenceDocxPath)} could not be read.`,
      "Create the reference DOCX first or pass the correct referenceDocxPath.",
      { cause },
    );
  }

  return referenceDocxPath;
}

function normalizeSourceDateEpoch(sourceDateEpoch: number | string | undefined): string {
  const normalized = sourceDateEpoch ?? DEFAULT_SOURCE_DATE_EPOCH;

  if (typeof normalized === "number") {
    if (!Number.isSafeInteger(normalized) || normalized < 0) {
      throw invalidSourceDateEpochError();
    }

    return String(normalized);
  }

  if (!/^(0|[1-9]\d*)$/.test(normalized)) {
    throw invalidSourceDateEpochError();
  }

  return normalized;
}

function invalidSourceDateEpochError(): PandocError {
  return new PandocError(
    "SOURCE_DATE_EPOCH_INVALID",
    "SOURCE_DATE_EPOCH must be a non-negative Unix timestamp.",
    "Pass sourceDateEpoch as a non-negative integer string or number.",
  );
}

function parsePandocVersion(output: string): PandocVersion | undefined {
  const match = /^pandoc\s+(\d+)(?:\.(\d+))?(?:\.(\d+))?([^\s]*)?/m.exec(output);

  if (match === null) {
    return undefined;
  }

  const major = Number.parseInt(match[1] ?? "", 10);
  const minor = match[2] === undefined ? undefined : Number.parseInt(match[2], 10);
  const patch = match[3] === undefined ? undefined : Number.parseInt(match[3], 10);

  if (!Number.isInteger(major)) {
    return undefined;
  }

  return {
    major,
    minor: Number.isNaN(minor) ? undefined : minor,
    patch: Number.isNaN(patch) ? undefined : patch,
    raw: [match[1], match[2], match[3]].filter((part) => part !== undefined).join("."),
  };
}

function parsePandocFeatures(output: string): readonly string[] {
  const featuresLine = output
    .split(/\r?\n/)
    .find((line) => line.toLowerCase().startsWith("features:"));

  if (featuresLine === undefined) {
    return [];
  }

  return featuresLine
    .replace(/^features:\s*/i, "")
    .split(/\s+/)
    .filter((feature) => feature.length > 0);
}

function pandocFailureToError(failure: PandocDoctorFailure): PandocError {
  return new PandocError(failure.code, failure.message, failure.guidance, { cause: failure.cause });
}

function decodeMarkdown(markdown: string | Uint8Array): string {
  if (typeof markdown === "string") {
    return markdown;
  }

  return new TextDecoder().decode(markdown);
}

function joinOutput(stdout: string, stderr: string): string {
  return [stdout, stderr].filter((text) => text.length > 0).join("\n");
}

function conversionGuidance(stderr: string): string {
  const trimmed = stderr.trim();
  const stderrSuffix = trimmed.length > 0 ? ` Pandoc stderr: ${trimmed.slice(0, 800)}` : "";

  return `Confirm the Markdown input and reference DOCX are valid, then retry.${stderrSuffix}`;
}

function isNotFoundError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    (error as { readonly code?: unknown }).code === "ENOENT"
  );
}
