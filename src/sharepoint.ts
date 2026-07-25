import { ConfidentialClientApplication } from "@azure/msal-node";

import type { Transport, TransportUploadResult } from "./transport.js";
import { TransportError, type TransportErrorCode } from "./transport.js";

export const MICROSOFT_GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default";
export const SHAREPOINT_REQUIRED_APPLICATION_PERMISSION = "Sites.Selected";
export const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
export const SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES = 250 * 1024 * 1024;
export const MICROSOFT_GRAPH_V1_BASE_URL = "https://graph.microsoft.com/v1.0";

export type SharePointTransportErrorCode = Extract<
  TransportErrorCode,
  | "SHAREPOINT_CONFIG_INVALID"
  | "SHAREPOINT_AUTH_FAILED"
  | "SHAREPOINT_PATH_INVALID"
  | "SHAREPOINT_DOCX_TOO_LARGE"
  | "SHAREPOINT_NETWORK_FAILED"
  | "SHAREPOINT_HTTP_FAILED"
  | "SHAREPOINT_RESPONSE_INVALID"
>;

export interface SharePointTransportErrorContext {
  readonly status?: number;
  readonly requestId?: string;
  readonly graphErrorCode?: string;
}

export class SharePointTransportError extends TransportError {
  readonly context: SharePointTransportErrorContext | undefined;

  constructor(
    code: SharePointTransportErrorCode,
    message: string,
    guidance: string,
    options: {
      readonly cause?: unknown;
      readonly context?: SharePointTransportErrorContext;
    } = {},
  ) {
    super(code, message, guidance, "cause" in options ? { cause: options.cause } : {});
    this.name = "SharePointTransportError";
    this.context = options.context;
  }
}

export interface SharePointAccessTokenRequest {
  readonly scopes: readonly [typeof MICROSOFT_GRAPH_DEFAULT_SCOPE];
}

export type SharePointAccessTokenProvider = (
  request: SharePointAccessTokenRequest,
) => Promise<string>;

export interface SharePointClientSecretCredentials {
  readonly tenantId: string;
  readonly clientId: string;
  readonly clientSecret: string;
}

export interface SharePointConfidentialClient {
  acquireTokenByClientCredential(request: {
    readonly scopes: readonly string[];
  }): Promise<{ readonly accessToken?: string | null } | null>;
}

export interface SharePointClientSecretAccessTokenProviderOptions
  extends SharePointClientSecretCredentials {
  readonly clientApplication?: SharePointConfidentialClient;
}

export type SharePointDestinationMapper = (canonicalId: string) => string;

export interface SharePointTransportDestination extends TransportUploadResult {
  readonly kind: "sharepoint";
  readonly driveItemId: string;
  readonly path: string;
  readonly webUrl?: string;
  readonly eTag?: string;
}

export interface SharePointTransportOptionsBase {
  readonly siteId: string;
  readonly driveId: string;
  readonly baseFolder?: string;
  readonly mapCanonicalId?: SharePointDestinationMapper;
  readonly fetch?: typeof fetch;
}

export type SharePointTransportOptions = SharePointTransportOptionsBase &
  (
    | {
        readonly tenantId: string;
        readonly clientId: string;
        readonly clientSecret: string;
        readonly accessTokenProvider?: never;
      }
    | {
        readonly accessTokenProvider: SharePointAccessTokenProvider;
        readonly tenantId?: never;
        readonly clientId?: never;
        readonly clientSecret?: never;
      }
  );

interface ResolvedSharePointDestination {
  readonly path: string;
  readonly url: string;
}

const INVALID_SHAREPOINT_NAME_CHARS = /["*<>?:|\\]/;
const RESERVED_WINDOWS_DEVICE_BASENAME = /^(?:con|prn|aux|nul|com[0-9]|lpt[0-9])(?:\..*)?$/i;
const MAX_DECODED_SHAREPOINT_PATH_LENGTH = 400;
const MAX_DECODED_SHAREPOINT_SEGMENT_LENGTH = 255;
const MAX_GRAPH_CONTEXT_VALUE_LENGTH = 128;

export function createSharePointClientSecretAccessTokenProvider(
  options: SharePointClientSecretAccessTokenProviderOptions,
): SharePointAccessTokenProvider {
  validateClientSecretCredentials(options);

  const clientApplication =
    options.clientApplication ??
    new ConfidentialClientApplication({
      auth: {
        authority: `https://login.microsoftonline.com/${encodeGraphPathSegment(options.tenantId)}`,
        clientId: options.clientId,
        clientSecret: options.clientSecret,
      },
    });

  return async (request) => {
    let response: { readonly accessToken?: string | null } | null;

    try {
      response = await clientApplication.acquireTokenByClientCredential({
        scopes: [...request.scopes],
      });
    } catch {
      throw new SharePointTransportError(
        "SHAREPOINT_AUTH_FAILED",
        "Failed to acquire a Microsoft Graph app-only access token.",
        "Check the Entra app registration, client credential, admin consent, and tenant ID.",
      );
    }

    const accessToken = response?.accessToken;

    if (typeof accessToken !== "string" || accessToken.trim() === "") {
      throw new SharePointTransportError(
        "SHAREPOINT_AUTH_FAILED",
        "Microsoft Graph token acquisition did not return an access token.",
        "Check the Entra app registration, client credential, admin consent, and tenant ID.",
      );
    }

    return accessToken;
  };
}

export class SharePointTransport implements Transport {
  readonly siteId: string;
  readonly driveId: string;
  readonly baseFolder: string | undefined;

  readonly #accessTokenProvider: SharePointAccessTokenProvider;
  readonly #fetch: typeof fetch;
  readonly #graphBaseUrl: string;
  readonly #mapCanonicalId: SharePointDestinationMapper;

  constructor(options: SharePointTransportOptions) {
    validateAuthShape(options);
    validateGraphPathParameter(options.siteId, "siteId");
    validateGraphPathParameter(options.driveId, "driveId");

    const baseFolder = options.baseFolder;

    this.siteId = options.siteId;
    this.driveId = options.driveId;
    this.baseFolder = baseFolder;
    this.#mapCanonicalId = options.mapCanonicalId ?? ((canonicalId) => canonicalId);
    this.#graphBaseUrl = MICROSOFT_GRAPH_V1_BASE_URL;
    this.#fetch = options.fetch ?? globalThis.fetch;

    if (typeof this.#fetch !== "function") {
      throw new SharePointTransportError(
        "SHAREPOINT_CONFIG_INVALID",
        "A fetch implementation is required for SharePoint transport.",
        "Run on Node.js 20 or newer, or pass a compatible fetch implementation.",
      );
    }

    this.#accessTokenProvider =
      "accessTokenProvider" in options && options.accessTokenProvider !== undefined
        ? options.accessTokenProvider
        : createSharePointClientSecretAccessTokenProvider({
            tenantId: options.tenantId,
            clientId: options.clientId,
            clientSecret: options.clientSecret,
          });
  }

  async upload(canonicalId: string, docx: Uint8Array): Promise<SharePointTransportDestination> {
    validateSharePointDocxSize(docx.byteLength);

    const destination = this.#resolveUploadDestination(canonicalId);
    const bytes = Buffer.from(docx);
    const accessToken = await this.#getAccessToken();
    const response = await this.#putDocx(destination.url, accessToken, bytes, canonicalId);

    return parseDriveItemResponse(response, destination.path, canonicalId);
  }

  #resolveUploadDestination(canonicalId: string): ResolvedSharePointDestination {
    validateCanonicalId(canonicalId);

    const mappedDestination = mapCanonicalId(canonicalId, this.#mapCanonicalId);
    const relativeDestination = ensureDocxExtension(
      normalizeSharePointRelativePath(mappedDestination, "mapped SharePoint destination"),
    );
    const destinationPath =
      this.baseFolder === undefined
        ? relativeDestination
        : `${this.baseFolder}/${relativeDestination}`;
    const normalizedDestinationPath = normalizeSharePointRelativePath(
      destinationPath,
      "SharePoint destination",
    );
    const encodedSiteId = encodeGraphPathSegment(this.siteId);
    const encodedDriveId = encodeGraphPathSegment(this.driveId);
    const encodedPath = encodeSharePointRelativePath(normalizedDestinationPath);
    const url = `${this.#graphBaseUrl}/sites/${encodedSiteId}/drives/${encodedDriveId}/root:/${encodedPath}:/content`;

    return {
      path: normalizedDestinationPath,
      url,
    };
  }

  async #getAccessToken(): Promise<string> {
    try {
      const accessToken = await this.#accessTokenProvider({
        scopes: [MICROSOFT_GRAPH_DEFAULT_SCOPE],
      });

      if (typeof accessToken !== "string" || accessToken.trim() === "") {
        throw new TypeError("SharePoint access token provider returned an empty token.");
      }

      return accessToken;
    } catch (error) {
      if (error instanceof SharePointTransportError) {
        throw error;
      }

      throw new SharePointTransportError(
        "SHAREPOINT_AUTH_FAILED",
        "Failed to acquire a Microsoft Graph app-only access token.",
        "Check the Entra app registration, client credential, admin consent, and tenant ID.",
      );
    }
  }

  async #putDocx(
    url: string,
    accessToken: string,
    bytes: Uint8Array,
    canonicalId: string,
  ): Promise<Response> {
    let response: Response;

    try {
      response = await this.#fetch(url, {
        body: bytes,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": DOCX_MIME_TYPE,
        },
        method: "PUT",
      });
    } catch {
      throw new SharePointTransportError(
        "SHAREPOINT_NETWORK_FAILED",
        `Failed to upload DOCX for canonical ID ${JSON.stringify(canonicalId)} to SharePoint.`,
        "Check network connectivity to Microsoft Graph and retry the upload.",
      );
    }

    if (!response.ok) {
      throw await createHttpError(response, canonicalId);
    }

    return response;
  }
}

export function validateSharePointDocxSize(byteLength: number): void {
  if (!Number.isSafeInteger(byteLength) || byteLength < 0) {
    throw new SharePointTransportError(
      "SHAREPOINT_DOCX_TOO_LARGE",
      "DOCX byte length must be a non-negative safe integer.",
      "Pass a real Uint8Array produced by convertMarkdownToDocx().",
    );
  }

  if (byteLength > SHAREPOINT_SIMPLE_UPLOAD_MAX_BYTES) {
    throw new SharePointTransportError(
      "SHAREPOINT_DOCX_TOO_LARGE",
      "SharePoint simple upload only supports DOCX payloads up to 250 MB.",
      "Use a smaller DOCX or implement an upload-session transport for larger files.",
    );
  }
}

export function encodeSharePointRelativePath(path: string): string {
  return normalizeSharePointRelativePath(path, "SharePoint relative path")
    .split("/")
    .map(encodeGraphPathSegment)
    .join("/");
}

function validateAuthShape(options: SharePointTransportOptions): void {
  const candidate = options as {
    readonly accessTokenProvider?: unknown;
    readonly tenantId?: unknown;
    readonly clientId?: unknown;
    readonly clientSecret?: unknown;
  };
  const hasAccessTokenProvider = candidate.accessTokenProvider !== undefined;
  const hasAnyCredential =
    candidate.tenantId !== undefined ||
    candidate.clientId !== undefined ||
    candidate.clientSecret !== undefined;

  if (hasAccessTokenProvider && hasAnyCredential) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      "SharePoint transport auth configuration is ambiguous.",
      "Pass either accessTokenProvider or tenantId/clientId/clientSecret, not both.",
    );
  }

  if (!hasAccessTokenProvider && !hasAnyCredential) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      "SharePoint transport auth configuration is required.",
      "Pass accessTokenProvider, or tenantId/clientId/clientSecret for MSAL client-secret auth.",
    );
  }
}

function validateClientSecretCredentials(options: SharePointClientSecretCredentials): void {
  validateTenantOrClientId(options.tenantId, "tenantId");
  validateTenantOrClientId(options.clientId, "clientId");
  validateClientSecret(options.clientSecret);
}

function validateTenantOrClientId(value: string, label: "tenantId" | "clientId"): void {
  if (typeof value !== "string" || value.trim() === "") {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      `SharePoint ${label} is required for MSAL client-secret auth.`,
      "Pass non-empty tenantId, clientId, and clientSecret values from the consuming application.",
    );
  }

  if (value !== value.trim() || containsControlCharacter(value) || value.includes("/")) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      `SharePoint ${label} is not a valid credential field.`,
      "Pass trimmed credential fields without control characters or URL path separators.",
    );
  }
}

function validateClientSecret(value: string): void {
  if (typeof value !== "string" || value.length === 0) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      "SharePoint clientSecret is required for MSAL client-secret auth.",
      "Pass the opaque client secret value from the consuming application.",
    );
  }

  if (containsControlCharacter(value)) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      "SharePoint clientSecret must not contain control characters.",
      "Pass the opaque client secret value from the consuming application.",
    );
  }
}

function validateGraphPathParameter(value: string, label: "siteId" | "driveId"): void {
  if (typeof value !== "string" || value.trim() === "") {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      `SharePoint ${label} is required.`,
      "Pass the stable Microsoft Graph site ID and document-library drive ID.",
    );
  }

  if (value !== value.trim() || containsControlCharacter(value) || /[/\\]/.test(value)) {
    throw new SharePointTransportError(
      "SHAREPOINT_CONFIG_INVALID",
      `SharePoint ${label} is not valid for a Microsoft Graph path parameter.`,
      "Pass a trimmed Graph ID without path separators or control characters.",
    );
  }
}

function validateCanonicalId(canonicalId: string): void {
  if (typeof canonicalId !== "string" || canonicalId.trim() === "") {
    throw new SharePointTransportError(
      "SHAREPOINT_PATH_INVALID",
      "A non-empty canonical ID is required for SharePoint upload.",
      "Pass the stable document ID used by the source system.",
    );
  }

  if (canonicalId !== canonicalId.trim() || canonicalId.includes("\0")) {
    throw new SharePointTransportError(
      "SHAREPOINT_PATH_INVALID",
      "SharePoint canonical IDs must be trimmed and must not contain NUL bytes.",
      "Normalize the source document ID before uploading.",
    );
  }
}

function mapCanonicalId(canonicalId: string, mapper: SharePointDestinationMapper): string {
  try {
    const mappedDestination = mapper(canonicalId);

    if (typeof mappedDestination !== "string") {
      throw new TypeError("SharePoint destination mapper must return a string.");
    }

    return mappedDestination;
  } catch (cause) {
    if (cause instanceof SharePointTransportError) {
      throw cause;
    }

    throw new SharePointTransportError(
      "SHAREPOINT_PATH_INVALID",
      `SharePoint destination mapping failed for canonical ID ${JSON.stringify(canonicalId)}.`,
      "Fix mapCanonicalId so it returns a safe relative DOCX destination.",
      { cause },
    );
  }
}

function normalizeSharePointRelativePath(path: string, label: string): string {
  if (typeof path !== "string" || path.trim() === "") {
    throw invalidSharePointPathError(label, "must be non-empty");
  }

  if (path.includes("\0") || containsControlCharacter(path)) {
    throw invalidSharePointPathError(label, "must not contain NUL bytes or control characters");
  }

  if (path.startsWith("/") || path.startsWith("//")) {
    throw invalidSharePointPathError(label, "must be relative to the document library root");
  }

  if (path.length > MAX_DECODED_SHAREPOINT_PATH_LENGTH) {
    throw invalidSharePointPathError(
      label,
      `must not exceed ${MAX_DECODED_SHAREPOINT_PATH_LENGTH} decoded characters`,
    );
  }

  const segments = path.split("/");

  for (const [index, segment] of segments.entries()) {
    validateSharePointPathSegment(segment, label, index);
  }

  return path;
}

function validateSharePointPathSegment(segment: string, label: string, index: number): void {
  if (segment === "" || segment === "." || segment === "..") {
    throw invalidSharePointPathError(
      label,
      "must not contain empty, current-directory, or parent-directory segments",
    );
  }

  if (segment !== segment.trim()) {
    throw invalidSharePointPathError(
      label,
      "must not contain path segments with leading or trailing whitespace",
    );
  }

  if (segment.length > MAX_DECODED_SHAREPOINT_SEGMENT_LENGTH) {
    throw invalidSharePointPathError(
      label,
      `must not contain path segments longer than ${MAX_DECODED_SHAREPOINT_SEGMENT_LENGTH} decoded characters`,
    );
  }

  if (segment.startsWith("~$")) {
    throw invalidSharePointPathError(label, "must not contain names beginning with ~$");
  }

  if (segment.startsWith("~")) {
    throw invalidSharePointPathError(label, "must not contain segments beginning with tilde");
  }

  if (segment.endsWith(".")) {
    throw invalidSharePointPathError(label, "must not contain segments ending with a period");
  }

  if (INVALID_SHAREPOINT_NAME_CHARS.test(segment)) {
    throw invalidSharePointPathError(label, 'must not contain " * < > ? : | or \\ characters');
  }

  if (RESERVED_WINDOWS_DEVICE_BASENAME.test(segment)) {
    throw invalidSharePointPathError(label, "must not use Windows-reserved device names");
  }

  validateBlockedSharePointName(segment, label, index);
}

function validateBlockedSharePointName(segment: string, label: string, index: number): void {
  const lowerSegment = segment.toLowerCase();

  if (lowerSegment === ".lock" || lowerSegment === ".lock.docx") {
    throw invalidSharePointPathError(label, "must not use the blocked name .lock");
  }

  if (lowerSegment === "desktop.ini" || lowerSegment === "desktop.ini.docx") {
    throw invalidSharePointPathError(label, "must not use the blocked name desktop.ini");
  }

  if (lowerSegment.includes("_vti_")) {
    throw invalidSharePointPathError(label, "must not contain _vti_ in path segment names");
  }

  if (index === 0 && lowerSegment === "forms") {
    throw invalidSharePointPathError(
      label,
      "must not use Forms as a root-level document-library item name",
    );
  }
}

function invalidSharePointPathError(label: string, reason: string): SharePointTransportError {
  return new SharePointTransportError(
    "SHAREPOINT_PATH_INVALID",
    `The ${label} ${reason}.`,
    "Return a relative SharePoint path without traversal, backslashes, or reserved SharePoint name characters.",
  );
}

function ensureDocxExtension(path: string): string {
  return path.toLowerCase().endsWith(".docx") ? path : `${path}.docx`;
}

function encodeGraphPathSegment(segment: string): string {
  return encodeURIComponent(segment).replace(
    /[!'()*]/g,
    (character) => `%${character.codePointAt(0)?.toString(16).toUpperCase()}`,
  );
}

function containsControlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0);

    if (codePoint !== undefined && codePoint <= 31) {
      return true;
    }
  }

  return false;
}

async function createHttpError(
  response: Response,
  canonicalId: string,
): Promise<SharePointTransportError> {
  const context = await readGraphErrorContext(response);
  const details = [
    `status ${response.status}`,
    context.requestId === undefined ? undefined : `request-id ${context.requestId}`,
    context.graphErrorCode === undefined ? undefined : `Graph error ${context.graphErrorCode}`,
  ].filter((detail): detail is string => detail !== undefined);

  return new SharePointTransportError(
    "SHAREPOINT_HTTP_FAILED",
    `Microsoft Graph rejected the SharePoint upload for canonical ID ${JSON.stringify(
      canonicalId,
    )} (${details.join(", ")}).`,
    "Check the site-specific write grant, document-library drive ID, mapped path, and Graph service response.",
    { context },
  );
}

async function readGraphErrorContext(response: Response): Promise<SharePointTransportErrorContext> {
  const requestId =
    sanitizeGraphContextValue(response.headers.get("request-id")) ??
    sanitizeGraphContextValue(response.headers.get("client-request-id"));
  const bodyText = await readBoundedResponseText(response);
  const graphErrorCode = extractGraphErrorCode(bodyText);
  const context: {
    status?: number;
    requestId?: string;
    graphErrorCode?: string;
  } = {
    status: response.status,
  };

  if (requestId !== undefined) {
    context.requestId = requestId;
  }

  if (graphErrorCode !== undefined) {
    context.graphErrorCode = graphErrorCode;
  }

  return context;
}

async function readBoundedResponseText(response: Response): Promise<string> {
  try {
    return (await response.text()).slice(0, 4096);
  } catch {
    return "";
  }
}

function extractGraphErrorCode(bodyText: string): string | undefined {
  if (bodyText.trim() === "") {
    return undefined;
  }

  try {
    const body = JSON.parse(bodyText) as unknown;
    const code = getStringProperty(getObjectProperty(body, "error"), "code");

    return sanitizeGraphContextValue(code);
  } catch {
    return undefined;
  }
}

function sanitizeGraphContextValue(value: string | null | undefined): string | undefined {
  if (value === null || value === undefined) {
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

  return trimmed === "" ? undefined : trimmed.slice(0, MAX_GRAPH_CONTEXT_VALUE_LENGTH);
}

async function parseDriveItemResponse(
  response: Response,
  path: string,
  canonicalId: string,
): Promise<SharePointTransportDestination> {
  let body: unknown;

  try {
    body = (await response.json()) as unknown;
  } catch {
    throw new SharePointTransportError(
      "SHAREPOINT_RESPONSE_INVALID",
      `Microsoft Graph returned invalid JSON for SharePoint upload of canonical ID ${JSON.stringify(
        canonicalId,
      )}.`,
      "Retry the upload or inspect the Graph response outside this library.",
    );
  }

  const driveItemId = getStringProperty(body, "id");

  if (driveItemId === undefined || driveItemId.trim() === "") {
    throw new SharePointTransportError(
      "SHAREPOINT_RESPONSE_INVALID",
      `Microsoft Graph did not return a driveItem id for SharePoint upload of canonical ID ${JSON.stringify(
        canonicalId,
      )}.`,
      "Check that the upload endpoint returns a driveItem response body.",
    );
  }

  const destination: {
    kind: "sharepoint";
    destinationId: string;
    driveItemId: string;
    path: string;
    webUrl?: string;
    eTag?: string;
  } = {
    kind: "sharepoint",
    destinationId: driveItemId,
    driveItemId,
    path,
  };
  const webUrl = getStringProperty(body, "webUrl");
  const eTag = getStringProperty(body, "eTag");

  if (webUrl !== undefined && webUrl.trim() !== "") {
    destination.webUrl = webUrl;
  }

  if (eTag !== undefined && eTag.trim() !== "") {
    destination.eTag = eTag;
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
