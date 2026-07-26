/**
 * Shared typed error for every Pandoc-backed conversion path.
 *
 * This lives in its own module so the pure DOCX archive inspector and the
 * Pandoc process layer can both raise the same error type without importing
 * each other.
 */

export type PandocDoctorFailureCode =
  | "PANDOC_NOT_FOUND"
  | "PANDOC_PROBE_FAILED"
  | "PANDOC_VERSION_UNPARSEABLE"
  | "PANDOC_UNSUPPORTED_MAJOR";

export type PandocErrorCode =
  | PandocDoctorFailureCode
  | "DOCX_ARCHIVE_INVALID"
  | "DOCX_INPUT_INVALID"
  | "DOCX_TRACK_CHANGES_INVALID"
  | "MARKDOWN_HOOK_FAILED"
  | "PANDOC_CONVERSION_FAILED"
  | "REFERENCE_DOC_REQUIRED"
  | "REFERENCE_DOC_INVALID"
  | "SOURCE_DATE_EPOCH_INVALID";

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
