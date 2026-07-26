import { Readable } from "node:stream";

import type { BaseExternalAccountClient, GoogleAuth, OAuth2Client } from "google-auth-library";

import type { Transport, TransportUploadResult } from "./transport.js";
import { DOCX_MIME_TYPE, TransportError, type TransportErrorCode } from "./transport.js";

/**
 * The shared transport contract is re-exported here so a consumer that only
 * imports this entry point can type a `Transport` and catch a `TransportError`
 * without also importing the root barrel. These are re-exports of the same
 * bindings the barrel exposes, not copies: both entry points resolve to one
 * `dist/transport.js` module instance, so `instanceof TransportError` holds
 * across entry points.
 */
export type { Transport, TransportErrorCode, TransportUploadResult } from "./transport.js";
export { DOCX_MIME_TYPE, TransportError } from "./transport.js";

/**
 * Least-privilege Drive scope: the app may only see and manage files it created
 * itself. This library never requests a broader Drive scope.
 */
export const GOOGLE_DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file";

/** Drive converts an uploaded DOCX on import when the file metadata uses this MIME type. */
export const GOOGLE_DOC_MIME_TYPE = "application/vnd.google-apps.document";

/**
 * Google documents a 50 MB ceiling for a text document converted to Google Docs
 * format, so a DOCX above this size cannot be imported as a Doc.
 * https://support.google.com/drive/answer/37603
 */
export const GOOGLE_DRIVE_DOCX_IMPORT_MAX_BYTES = 50 * 1024 * 1024;

/** Fields requested back from Drive after a create or update. */
const GOOGLE_DRIVE_RESPONSE_FIELDS = "id,name,mimeType,webViewLink";

const MAX_GOOGLE_DRIVE_NAME_LENGTH = 255;
/** Drive resource IDs are short opaque tokens, unrelated to document name limits. */
const MAX_GOOGLE_DRIVE_ID_LENGTH = 128;
const MAX_GOOGLE_DRIVE_CONTEXT_VALUE_LENGTH = 128;
const TRAILING_DOCX_EXTENSION = /\.docx$/i;
const GOOGLE_DRIVE_ID_PATTERN = /^[A-Za-z0-9_-]+$/;
const INSUFFICIENT_PERMISSION_REASON = /^insufficient/i;

export type GoogleDriveTransportErrorCode = Extract<
  TransportErrorCode,
  | "GOOGLE_DRIVE_CONFIG_INVALID"
  | "GOOGLE_DRIVE_AUTH_FAILED"
  | "GOOGLE_DRIVE_NAME_INVALID"
  | "GOOGLE_DRIVE_DOCX_TOO_LARGE"
  | "GOOGLE_DRIVE_NETWORK_FAILED"
  | "GOOGLE_DRIVE_API_FAILED"
  | "GOOGLE_DRIVE_RESPONSE_INVALID"
>;

export interface GoogleDriveTransportErrorContext {
  readonly status?: number;
  readonly reason?: string;
  readonly fileId?: string;
}

export class GoogleDriveTransportError extends TransportError {
  readonly context: GoogleDriveTransportErrorContext | undefined;

  constructor(
    code: GoogleDriveTransportErrorCode,
    message: string,
    guidance: string,
    options: {
      readonly cause?: unknown;
      readonly context?: GoogleDriveTransportErrorContext;
    } = {},
  ) {
    super(code, message, guidance, "cause" in options ? { cause: options.cause } : {});
    this.name = "GoogleDriveTransportError";
    this.context = options.context;
  }
}

/** Auth clients accepted by the Drive v3 client. `OAuth2Client` is the usual choice. */
export type GoogleDriveAuthClient = BaseExternalAccountClient | GoogleAuth | OAuth2Client;

export interface GoogleDriveFileMetadata {
  readonly id?: string | null;
  readonly name?: string | null;
  readonly mimeType?: string | null;
  readonly webViewLink?: string | null;
}

export interface GoogleDriveFileResponse {
  readonly data?: GoogleDriveFileMetadata | null;
}

export interface GoogleDriveMediaBody {
  readonly mimeType: string;
  readonly body: Readable;
}

export interface GoogleDriveCreateFileParams {
  readonly requestBody: {
    readonly name: string;
    readonly mimeType: string;
    readonly parents?: string[];
  };
  readonly media: GoogleDriveMediaBody;
  readonly fields: string;
  readonly supportsAllDrives: boolean;
}

export interface GoogleDriveUpdateFileParams {
  readonly fileId: string;
  readonly requestBody: {
    readonly name: string;
    readonly mimeType: string;
  };
  readonly media: GoogleDriveMediaBody;
  readonly fields: string;
  readonly supportsAllDrives: boolean;
}

/**
 * Minimal structural view of the Drive v3 files resource used by this transport.
 * Consumers can inject a stand-in for tests without constructing a real client.
 */
export interface GoogleDriveFilesClient {
  readonly files: {
    create(params: GoogleDriveCreateFileParams): Promise<GoogleDriveFileResponse>;
    update(params: GoogleDriveUpdateFileParams): Promise<GoogleDriveFileResponse>;
  };
}

export type GoogleDriveNameMapper = (canonicalId: string) => string;

export type GoogleDriveExistingFileIdResolver = (
  canonicalId: string,
) => string | undefined | Promise<string | undefined>;

export interface GoogleDriveTransportDestination extends TransportUploadResult {
  readonly kind: "google-drive";
  readonly fileId: string;
  readonly name: string;
  readonly webViewLink?: string;
  readonly mimeType?: string;
}

export interface GoogleDriveTransportOptionsBase {
  readonly folderId?: string;
  readonly mapCanonicalId?: GoogleDriveNameMapper;
  readonly resolveExistingFileId?: GoogleDriveExistingFileIdResolver;
}

export type GoogleDriveTransportOptions = GoogleDriveTransportOptionsBase &
  (
    | {
        readonly auth: GoogleDriveAuthClient;
        readonly driveClient?: never;
      }
    | {
        readonly driveClient: GoogleDriveFilesClient;
        readonly auth?: never;
      }
  );

export class GoogleDriveTransport implements Transport {
  readonly folderId: string | undefined;

  readonly #auth: GoogleDriveAuthClient | undefined;
  readonly #mapCanonicalId: GoogleDriveNameMapper;
  readonly #usesDefaultNameMapper: boolean;
  readonly #resolveExistingFileId: GoogleDriveExistingFileIdResolver | undefined;
  #driveClient: Promise<GoogleDriveFilesClient> | undefined;

  constructor(options: GoogleDriveTransportOptions) {
    validateAuthShape(options);

    if (options.folderId !== undefined) {
      validateFolderId(options.folderId);
    }

    this.folderId = options.folderId;
    this.#usesDefaultNameMapper = options.mapCanonicalId === undefined;
    this.#mapCanonicalId = options.mapCanonicalId ?? defaultGoogleDocName;
    this.#resolveExistingFileId = options.resolveExistingFileId;
    this.#auth = options.auth;
    this.#driveClient =
      options.driveClient === undefined ? undefined : Promise.resolve(options.driveClient);
  }

  async upload(canonicalId: string, docx: Uint8Array): Promise<GoogleDriveTransportDestination> {
    validateGoogleDriveDocxSize(docx.byteLength);
    validateCanonicalId(canonicalId);

    const name = validateGoogleDocName(
      mapCanonicalId(canonicalId, this.#mapCanonicalId),
      this.#usesDefaultNameMapper,
    );
    // Load-bearing copy: Readable.from() special-cases Buffer into a single
    // chunk, but iterates a raw Uint8Array byte by byte, which would emit one
    // chunk per byte and corrupt every upload. It also snapshots the bytes
    // before any await, so caller mutation cannot race the upload.
    const bytes = Buffer.from(docx);
    const existingFileId = await this.#resolveFileId(canonicalId);
    const driveClient = await this.#getDriveClient();
    const response =
      existingFileId === undefined
        ? await createFile(driveClient, canonicalId, name, bytes, this.folderId)
        : await updateFile(driveClient, canonicalId, existingFileId, name, bytes);

    return parseFileResponse(response, canonicalId, name);
  }

  /**
   * The Drive SDK is a heavy transitive dependency, so it is imported on first
   * upload rather than at module load. Consumers that only use another
   * transport never pay for it.
   */
  #getDriveClient(): Promise<GoogleDriveFilesClient> {
    this.#driveClient ??= importDriveFilesClient(this.#auth).catch((cause: unknown) => {
      // Never memoize a failed import, or one bad load would poison the
      // transport for the rest of the process.
      this.#driveClient = undefined;
      throw cause;
    });

    return this.#driveClient;
  }

  async #resolveFileId(canonicalId: string): Promise<string | undefined> {
    const resolver = this.#resolveExistingFileId;

    if (resolver === undefined) {
      return undefined;
    }

    let resolved: string | undefined;

    try {
      resolved = await resolver(canonicalId);
    } catch (cause) {
      if (cause instanceof GoogleDriveTransportError) {
        throw cause;
      }

      throw new GoogleDriveTransportError(
        "GOOGLE_DRIVE_CONFIG_INVALID",
        `Resolving the existing Google Drive file ID failed for canonical ID ${JSON.stringify(
          canonicalId,
        )}.`,
        "Fix resolveExistingFileId so it returns a stored Drive file ID or undefined.",
        { cause },
      );
    }

    if (resolved === undefined) {
      return undefined;
    }

    if (typeof resolved !== "string" || resolved.trim() === "" || !isSafeDriveId(resolved)) {
      throw new GoogleDriveTransportError(
        "GOOGLE_DRIVE_CONFIG_INVALID",
        `resolveExistingFileId returned an invalid Google Drive file ID for canonical ID ${JSON.stringify(
          canonicalId,
        )}.`,
        "Fix resolveExistingFileId so it returns the trimmed Drive file ID recorded by the consumer manifest, or undefined for a new document.",
      );
    }

    return resolved;
  }
}

async function createFile(
  driveClient: GoogleDriveFilesClient,
  canonicalId: string,
  name: string,
  bytes: Buffer,
  folderId: string | undefined,
): Promise<GoogleDriveFileResponse> {
  const requestBody: { name: string; mimeType: string; parents?: string[] } = {
    name,
    mimeType: GOOGLE_DOC_MIME_TYPE,
  };

  if (folderId !== undefined) {
    requestBody.parents = [folderId];
  }

  try {
    return await driveClient.files.create({
      requestBody,
      media: { mimeType: DOCX_MIME_TYPE, body: Readable.from(bytes) },
      fields: GOOGLE_DRIVE_RESPONSE_FIELDS,
      supportsAllDrives: true,
    });
  } catch (cause) {
    throw createApiError(cause, canonicalId, "create");
  }
}

async function updateFile(
  driveClient: GoogleDriveFilesClient,
  canonicalId: string,
  fileId: string,
  name: string,
  bytes: Buffer,
): Promise<GoogleDriveFileResponse> {
  try {
    // Drive rejects `parents` on update, and this transport does not silently
    // move an existing document, so folderId is intentionally create-only.
    return await driveClient.files.update({
      fileId,
      requestBody: { name, mimeType: GOOGLE_DOC_MIME_TYPE },
      media: { mimeType: DOCX_MIME_TYPE, body: Readable.from(bytes) },
      fields: GOOGLE_DRIVE_RESPONSE_FIELDS,
      supportsAllDrives: true,
    });
  } catch (cause) {
    throw createApiError(cause, canonicalId, "update");
  }
}

export function validateGoogleDriveDocxSize(byteLength: number): void {
  if (!Number.isSafeInteger(byteLength) || byteLength < 0) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_DOCX_TOO_LARGE",
      "DOCX byte length must be a non-negative safe integer.",
      "Pass a real Uint8Array produced by convertMarkdownToDocx().",
    );
  }

  if (byteLength > GOOGLE_DRIVE_DOCX_IMPORT_MAX_BYTES) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_DOCX_TOO_LARGE",
      "Google Drive only converts text documents up to 50 MB into Google Docs format.",
      "Split the source document or publish a smaller DOCX.",
    );
  }
}

/**
 * Default Google Doc name mapper. Drive has no path namespace, so the canonical
 * ID is used verbatim except for a trailing `.docx` extension: a converted
 * Google Doc should not be named `foo.docx`.
 */
export function defaultGoogleDocName(canonicalId: string): string {
  return canonicalId.replace(TRAILING_DOCX_EXTENSION, "");
}

async function importDriveFilesClient(
  auth: GoogleDriveAuthClient | undefined,
): Promise<GoogleDriveFilesClient> {
  if (auth === undefined) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_CONFIG_INVALID",
      "Google Drive transport has no auth client to build a Drive client with.",
      `Pass a google-auth-library client as auth, authorized for the ${GOOGLE_DRIVE_FILE_SCOPE} scope, or pass driveClient.`,
    );
  }

  const { drive } = await import("@googleapis/drive");
  const driveClient = drive({ version: "v3", auth });

  return {
    files: {
      create: async (params) => driveClient.files.create(params),
      update: async (params) => driveClient.files.update(params),
    },
  };
}

function validateAuthShape(options: GoogleDriveTransportOptions): void {
  const candidate = options as {
    readonly auth?: unknown;
    readonly driveClient?: unknown;
  };
  const hasAuth = candidate.auth !== undefined;
  const hasDriveClient = candidate.driveClient !== undefined;

  if (hasAuth && hasDriveClient) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_CONFIG_INVALID",
      "Google Drive transport auth configuration is ambiguous.",
      `Pass either auth or driveClient, not both. The auth client must already hold the ${GOOGLE_DRIVE_FILE_SCOPE} scope.`,
    );
  }

  if (!hasAuth && !hasDriveClient) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_CONFIG_INVALID",
      "Google Drive transport auth configuration is required.",
      `Pass a google-auth-library client as auth, authorized for the ${GOOGLE_DRIVE_FILE_SCOPE} scope, or pass driveClient.`,
    );
  }
}

function validateFolderId(folderId: string): void {
  if (typeof folderId !== "string" || folderId.trim() === "") {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_CONFIG_INVALID",
      "Google Drive folderId must be a non-empty string when provided.",
      "Pass the Drive folder ID that should own newly created documents, or omit folderId.",
    );
  }

  if (!isSafeDriveId(folderId)) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_CONFIG_INVALID",
      "Google Drive folderId is not a valid Drive resource ID.",
      "Pass the opaque Drive folder ID as it appears in the Drive URL: letters, digits, hyphens, and underscores only.",
    );
  }
}

function isSafeDriveId(value: string): boolean {
  return value.length <= MAX_GOOGLE_DRIVE_ID_LENGTH && GOOGLE_DRIVE_ID_PATTERN.test(value);
}

function validateCanonicalId(canonicalId: string): void {
  if (typeof canonicalId !== "string" || canonicalId.trim() === "") {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_NAME_INVALID",
      "A non-empty canonical ID is required for Google Drive upload.",
      "Pass the stable document ID used by the source system.",
    );
  }

  if (canonicalId !== canonicalId.trim() || canonicalId.includes("\0")) {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_NAME_INVALID",
      "Google Drive canonical IDs must be trimmed and must not contain NUL bytes.",
      "Normalize the source document ID before uploading.",
    );
  }
}

function mapCanonicalId(canonicalId: string, mapper: GoogleDriveNameMapper): string {
  try {
    const mappedName = mapper(canonicalId);

    if (typeof mappedName !== "string") {
      throw new TypeError("Google Drive name mapper must return a string.");
    }

    return mappedName;
  } catch (cause) {
    if (cause instanceof GoogleDriveTransportError) {
      throw cause;
    }

    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_NAME_INVALID",
      `Google Drive name mapping failed for canonical ID ${JSON.stringify(canonicalId)}.`,
      "Fix mapCanonicalId so it returns a valid Google Doc name.",
      { cause },
    );
  }
}

function validateGoogleDocName(name: string, usesDefaultMapper: boolean): string {
  if (name.trim() === "") {
    throw invalidGoogleDocNameError("must be non-empty", usesDefaultMapper);
  }

  if (name !== name.trim()) {
    throw invalidGoogleDocNameError(
      "must not have leading or trailing whitespace",
      usesDefaultMapper,
    );
  }

  if (name.includes("\0") || containsControlCharacter(name)) {
    throw invalidGoogleDocNameError(
      "must not contain NUL bytes or control characters",
      usesDefaultMapper,
    );
  }

  if (name.includes("/")) {
    throw invalidGoogleDocNameError(
      "must not contain slashes, because Drive names are not paths",
      usesDefaultMapper,
    );
  }

  if (name.length > MAX_GOOGLE_DRIVE_NAME_LENGTH) {
    throw invalidGoogleDocNameError(
      `must not exceed ${MAX_GOOGLE_DRIVE_NAME_LENGTH} characters`,
      usesDefaultMapper,
    );
  }

  return name;
}

function invalidGoogleDocNameError(
  reason: string,
  usesDefaultMapper: boolean,
): GoogleDriveTransportError {
  return new GoogleDriveTransportError(
    "GOOGLE_DRIVE_NAME_INVALID",
    `The ${usesDefaultMapper ? "default-mapped" : "mapped"} Google Doc name ${reason}.`,
    usesDefaultMapper
      ? "This canonical ID is not usable as a Drive document name on its own; pass mapCanonicalId to map it to a valid name."
      : "Return a single trimmed Drive document name from mapCanonicalId, not a path.",
  );
}

function containsControlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0);

    if (codePoint !== undefined && (codePoint <= 31 || codePoint === 127)) {
      return true;
    }
  }

  return false;
}

function createApiError(
  cause: unknown,
  canonicalId: string,
  operation: "create" | "update",
): GoogleDriveTransportError {
  if (cause instanceof GoogleDriveTransportError) {
    return cause;
  }

  const context = readDriveErrorContext(cause);
  const details = [
    context.status === undefined ? undefined : `status ${context.status}`,
    context.reason === undefined ? undefined : `Drive reason ${context.reason}`,
  ].filter((detail): detail is string => detail !== undefined);
  const suffix = details.length === 0 ? "" : ` (${details.join(", ")})`;

  if (context.status === undefined) {
    // No HTTP status reached us, so the request never completed: transport-level
    // failure (DNS, TLS, socket) rather than a decision by the Drive API.
    return new GoogleDriveTransportError(
      "GOOGLE_DRIVE_NETWORK_FAILED",
      `Failed to reach the Google Drive API for canonical ID ${JSON.stringify(canonicalId)}.`,
      "The request did not complete, so Drive never decided on it. This is usually connectivity, but a client misconfiguration can also prevent the request from being sent — do not retry indefinitely.",
      { context },
    );
  }

  if (isAuthFailure(context)) {
    return new GoogleDriveTransportError(
      "GOOGLE_DRIVE_AUTH_FAILED",
      `Google Drive rejected the credentials for canonical ID ${JSON.stringify(
        canonicalId,
      )}${suffix}.`,
      `Refresh the auth client the consumer owns and confirm it is authorized for the ${GOOGLE_DRIVE_FILE_SCOPE} scope.`,
      { context },
    );
  }

  return new GoogleDriveTransportError(
    "GOOGLE_DRIVE_API_FAILED",
    `Google Drive files.${operation} failed for canonical ID ${JSON.stringify(
      canonicalId,
    )}${suffix}.`,
    "Check the Drive folder ID, the stored file ID, and the drive.file authorization of the supplied auth client.",
    { context },
  );
}

/**
 * A 401 is always an auth failure. A 403 is overloaded (quota, rate limit,
 * sharing policy), so only the `insufficient*` reasons — the shape Drive
 * returns for a token missing `drive.file` — are treated as auth failures.
 */
function isAuthFailure(context: GoogleDriveTransportErrorContext): boolean {
  if (context.status === 401) {
    return true;
  }

  return (
    context.status === 403 &&
    context.reason !== undefined &&
    INSUFFICIENT_PERMISSION_REASON.test(context.reason)
  );
}

function readDriveErrorContext(cause: unknown): GoogleDriveTransportErrorContext {
  const status = readStatus(cause);
  const reason = readReason(cause);
  const context: { status?: number; reason?: string } = {};

  if (status !== undefined) {
    context.status = status;
  }

  if (reason !== undefined) {
    context.reason = reason;
  }

  return context;
}

function readStatus(cause: unknown): number | undefined {
  const direct = getNumberProperty(cause, "status") ?? getNumberProperty(cause, "code");

  if (direct !== undefined) {
    return direct;
  }

  const response = getObjectProperty(cause, "response");
  const responseStatus = getNumberProperty(response, "status");

  if (responseStatus !== undefined) {
    return responseStatus;
  }

  return getNumberProperty(getObjectProperty(getObjectProperty(response, "data"), "error"), "code");
}

function readReason(cause: unknown): string | undefined {
  const error = getObjectProperty(
    getObjectProperty(getObjectProperty(cause, "response"), "data"),
    "error",
  );
  const errors = error === undefined ? undefined : error.errors;
  const firstError = Array.isArray(errors) ? errors[0] : undefined;

  return (
    sanitizeContextValue(getStringProperty(firstError, "reason")) ??
    sanitizeContextValue(getStringProperty(error, "status"))
  );
}

function sanitizeContextValue(value: string | undefined): string | undefined {
  if (value === undefined) {
    return undefined;
  }

  let sanitized = "";

  for (const character of value) {
    const codePoint = character.codePointAt(0);

    if (codePoint !== undefined && codePoint > 31 && codePoint !== 127) {
      sanitized += character;
    }
  }

  const trimmed = sanitized.trim();

  return trimmed === "" ? undefined : trimmed.slice(0, MAX_GOOGLE_DRIVE_CONTEXT_VALUE_LENGTH);
}

function parseFileResponse(
  response: GoogleDriveFileResponse,
  canonicalId: string,
  requestedName: string,
): GoogleDriveTransportDestination {
  const file = response.data;
  const fileId = getStringProperty(file, "id");

  if (fileId === undefined || fileId.trim() === "") {
    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_RESPONSE_INVALID",
      `Google Drive did not return a file ID for canonical ID ${JSON.stringify(canonicalId)}.`,
      "Retry the upload or inspect the Drive API response outside this library.",
    );
  }

  const returnedMimeType = getStringProperty(file, "mimeType");

  // Drive reports what the stored file actually is. Anything other than a
  // Google Doc means the DOCX was stored as a blob instead of being converted:
  // typically resolveExistingFileId pointing at a plain DOCX upload, where every
  // future publish would silently write a file nobody can open as a Doc.
  if (
    returnedMimeType !== undefined &&
    returnedMimeType !== "" &&
    returnedMimeType !== GOOGLE_DOC_MIME_TYPE
  ) {
    // fileId is unvalidated Drive response data, so it is sanitized for the
    // human-readable strings. context.fileId stays raw: the consumer needs the
    // real ID to delete or adopt the orphaned file.
    const safeFileId = sanitizeContextValue(fileId) ?? "unknown";

    throw new GoogleDriveTransportError(
      "GOOGLE_DRIVE_RESPONSE_INVALID",
      `Google Drive stored canonical ID ${JSON.stringify(canonicalId)} as ${
        sanitizeContextValue(returnedMimeType) ?? "an unknown type"
      } instead of a Google Doc (file ID ${safeFileId}).`,
      `Drive did not convert the upload. Remove or re-point the stored file ID for this canonical ID, and delete or adopt Drive file ${safeFileId} by hand.`,
      { context: { fileId } },
    );
  }

  const returnedName = getStringProperty(file, "name");
  const destination: {
    kind: "google-drive";
    destinationId: string;
    fileId: string;
    name: string;
    webViewLink?: string;
    mimeType?: string;
  } = {
    kind: "google-drive",
    destinationId: fileId,
    fileId,
    name: returnedName === undefined || returnedName === "" ? requestedName : returnedName,
  };
  const webViewLink = getStringProperty(file, "webViewLink");

  if (webViewLink !== undefined && webViewLink.trim() !== "") {
    destination.webViewLink = webViewLink;
  }

  if (returnedMimeType !== undefined && returnedMimeType.trim() !== "") {
    destination.mimeType = returnedMimeType;
  }

  return destination;
}

function getObjectProperty(value: unknown, key: string): Record<string, unknown> | undefined {
  if (typeof value !== "object" || value === null || !Object.hasOwn(value, key)) {
    return undefined;
  }

  const property = (value as Record<string, unknown>)[key];

  return typeof property === "object" && property !== null
    ? (property as Record<string, unknown>)
    : undefined;
}

function getStringProperty(value: unknown, key: string): string | undefined {
  if (typeof value !== "object" || value === null || !Object.hasOwn(value, key)) {
    return undefined;
  }

  const property = (value as Record<string, unknown>)[key];

  return typeof property === "string" ? property : undefined;
}

function getNumberProperty(value: unknown, key: string): number | undefined {
  if (typeof value !== "object" || value === null || !Object.hasOwn(value, key)) {
    return undefined;
  }

  const property = (value as Record<string, unknown>)[key];

  return typeof property === "number" && Number.isFinite(property) ? property : undefined;
}
