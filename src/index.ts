/**
 * Root entry point: conversion core, the transport contract, and the local file
 * transport.
 *
 * This barrel deliberately pulls in no third-party SDK. Its whole dependency
 * closure is Node builtins plus `execa` and `fflate`, so a consumer that only
 * converts Markdown to DOCX never loads a cloud SDK. The cloud transports live
 * behind their own entry points and are opt-in:
 *
 * - `@agentic-tooling/polydoc-core/sharepoint` — `SharePointTransport`
 * - `@agentic-tooling/polydoc-core/google` — `GoogleDriveTransport`
 *
 * Both re-export the shared transport contract, so a consumer of one transport
 * never has to reach back into this barrel for the types it needs. Adding a
 * cloud transport export here would undo the split; `tests/entrypoints.test.ts`
 * fails if the built barrel regains an SDK dependency.
 */
export type {
  DocxImportReport,
  DocxTrackChangesMode,
  DocxTrackedChangeCounts,
  DocxUnmappableItem,
  DocxUnmappableKind,
} from "./docx.js";
export {
  buildDocxImportReport,
  DEFAULT_DOCX_DOCUMENT_PART,
  DEFAULT_DOCX_TRACK_CHANGES,
  DOCX_MAX_ARCHIVE_ENTRIES,
  DOCX_MAX_HEADER_FOOTER_PARTS,
  DOCX_MAX_PART_BYTES,
  DOCX_MAX_TOTAL_BYTES,
  DOCX_TRACK_CHANGES_MODES,
} from "./docx.js";
export type {
  ConvertDocxToMarkdownOptions,
  ConvertMarkdownToDocxOptions,
  DocxToMarkdownResult,
  MarkdownPostprocessor,
  MarkdownPreprocessor,
  MarkdownProcessor,
  MarkdownProcessorContext,
  PandocDoctorFailure,
  PandocDoctorFailureCode,
  PandocDoctorOptions,
  PandocDoctorResult,
  PandocDoctorSuccess,
  PandocErrorCode,
  PandocRunner,
  PandocRunnerOptions,
  PandocRunnerResult,
  PandocVersion,
} from "./pandoc.js";
export {
  applyMarkdownPostprocessors,
  applyMarkdownPreprocessors,
  convertDocxToMarkdown,
  convertMarkdownToDocx,
  DEFAULT_SOURCE_DATE_EPOCH,
  doctor,
  PandocError,
  SUPPORTED_PANDOC_MAJOR,
} from "./pandoc.js";
export type {
  LocalFileDestinationMapper,
  LocalFileTransportDestination,
  LocalFileTransportOptions,
  Transport,
  TransportErrorCode,
  TransportUploadResult,
} from "./transport.js";
export { DOCX_MIME_TYPE, LocalFileTransport, TransportError } from "./transport.js";
