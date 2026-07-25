import { mkdir, writeFile } from "node:fs/promises";
import { dirname, isAbsolute, relative, resolve, sep, win32 } from "node:path";

export type TransportErrorCode =
  | "TRANSPORT_ROOT_INVALID"
  | "TRANSPORT_ID_INVALID"
  | "TRANSPORT_DESTINATION_INVALID"
  | "TRANSPORT_DESTINATION_OUTSIDE_ROOT"
  | "TRANSPORT_WRITE_FAILED"
  | "SHAREPOINT_CONFIG_INVALID"
  | "SHAREPOINT_AUTH_FAILED"
  | "SHAREPOINT_PATH_INVALID"
  | "SHAREPOINT_DOCX_TOO_LARGE"
  | "SHAREPOINT_NETWORK_FAILED"
  | "SHAREPOINT_HTTP_FAILED"
  | "SHAREPOINT_RESPONSE_INVALID";

export interface TransportUploadResult {
  readonly destinationId: string;
}

export interface Transport {
  upload(canonicalId: string, docx: Uint8Array): Promise<TransportUploadResult>;
}

export interface LocalFileTransportDestination extends TransportUploadResult {
  readonly kind: "local-file";
  readonly path: string;
}

export type LocalFileDestinationMapper = (canonicalId: string) => string;

export interface LocalFileTransportOptions {
  readonly rootDir: string;
  readonly mapCanonicalId?: LocalFileDestinationMapper;
}

const WINDOWS_RESERVED_DEVICE_BASENAME = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?$/i;
const WINDOWS_INVALID_FILENAME_CHARS = /[<>"|?*]/;

export class TransportError extends Error {
  readonly code: TransportErrorCode;
  readonly guidance: string;

  constructor(
    code: TransportErrorCode,
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

    this.name = "TransportError";
    this.code = code;
    this.guidance = guidance;
  }
}

export class LocalFileTransport implements Transport {
  readonly rootDir: string;
  readonly #mapCanonicalId: LocalFileDestinationMapper;

  constructor(options: LocalFileTransportOptions) {
    if (options.rootDir.trim() === "") {
      throw new TransportError(
        "TRANSPORT_ROOT_INVALID",
        "A local file transport rootDir is required.",
        "Pass rootDir pointing at the directory where generated DOCX files may be written.",
      );
    }

    this.rootDir = resolve(options.rootDir);
    this.#mapCanonicalId = options.mapCanonicalId ?? ((canonicalId) => canonicalId);
  }

  resolveDestination(canonicalId: string): LocalFileTransportDestination {
    validateCanonicalId(canonicalId);

    const mappedDestination = mapCanonicalId(canonicalId, this.#mapCanonicalId);
    const relativeDestination = normalizeRelativeDestination(mappedDestination);
    const destinationPath = resolve(this.rootDir, ensureDocxExtension(relativeDestination));
    const relativeToRoot = relative(this.rootDir, destinationPath);

    if (
      relativeToRoot === "" ||
      relativeToRoot === ".." ||
      relativeToRoot.startsWith(`..${sep}`) ||
      isAbsolute(relativeToRoot) ||
      win32.isAbsolute(relativeToRoot)
    ) {
      throw new TransportError(
        "TRANSPORT_DESTINATION_OUTSIDE_ROOT",
        `Canonical ID ${JSON.stringify(canonicalId)} maps outside the local file transport root.`,
        "Return a relative DOCX destination under rootDir from mapCanonicalId.",
      );
    }

    return {
      kind: "local-file",
      destinationId: destinationPath,
      path: destinationPath,
    };
  }

  async upload(canonicalId: string, docx: Uint8Array): Promise<LocalFileTransportDestination> {
    const destination = this.resolveDestination(canonicalId);
    const bytes = Buffer.from(docx);

    try {
      await mkdir(dirname(destination.path), { recursive: true });
      await writeFile(destination.path, bytes);
    } catch (cause) {
      throw new TransportError(
        "TRANSPORT_WRITE_FAILED",
        `Failed to write DOCX for canonical ID ${JSON.stringify(canonicalId)}.`,
        "Check that rootDir is writable and that the destination parent can be created.",
        { cause },
      );
    }

    return destination;
  }
}

function mapCanonicalId(canonicalId: string, mapper: LocalFileDestinationMapper): string {
  try {
    const mappedDestination = mapper(canonicalId);

    if (typeof mappedDestination !== "string") {
      throw new TypeError("Local file destination mapper must return a string.");
    }

    return mappedDestination;
  } catch (cause) {
    if (cause instanceof TransportError) {
      throw cause;
    }

    throw new TransportError(
      "TRANSPORT_DESTINATION_INVALID",
      `Local file transport mapping failed for canonical ID ${JSON.stringify(canonicalId)}.`,
      "Fix mapCanonicalId so it returns a safe relative DOCX destination.",
      { cause },
    );
  }
}

function validateCanonicalId(canonicalId: string): void {
  if (canonicalId.trim() === "") {
    throw new TransportError(
      "TRANSPORT_ID_INVALID",
      "A non-empty canonical ID is required for transport upload.",
      "Pass the stable document ID used by the source system.",
    );
  }

  if (canonicalId !== canonicalId.trim()) {
    throw new TransportError(
      "TRANSPORT_ID_INVALID",
      "Canonical IDs must not include leading or trailing whitespace.",
      "Normalize the document ID before uploading.",
    );
  }

  if (canonicalId.includes("\0")) {
    throw new TransportError(
      "TRANSPORT_ID_INVALID",
      "Canonical IDs must not contain NUL bytes.",
      "Pass the stable document ID used by the source system.",
    );
  }
}

function normalizeRelativeDestination(destination: string): string {
  if (destination.trim() === "") {
    throw new TransportError(
      "TRANSPORT_DESTINATION_INVALID",
      "Local file transport mapping returned an empty destination.",
      "Return a relative DOCX destination under rootDir from mapCanonicalId.",
    );
  }

  if (destination !== destination.trim()) {
    throw new TransportError(
      "TRANSPORT_DESTINATION_INVALID",
      "Local file transport destinations must not include leading or trailing whitespace.",
      "Normalize mapped destinations before returning them from mapCanonicalId.",
    );
  }

  validateRelativePathText(
    destination,
    "TRANSPORT_DESTINATION_INVALID",
    "local file transport destination",
  );

  return destination;
}

function validateRelativePathText(
  value: string,
  code: "TRANSPORT_ID_INVALID" | "TRANSPORT_DESTINATION_INVALID",
  label: string,
): void {
  if (value.includes("\0")) {
    throw invalidRelativePathError(code, label, "must not contain NUL bytes");
  }

  if (containsWindowsInvalidControlCharacter(value)) {
    throw invalidRelativePathError(code, label, "must not contain control characters");
  }

  if (value.includes("\\")) {
    throw invalidRelativePathError(code, label, "must use forward slashes, not backslashes");
  }

  if (value.includes(":")) {
    throw invalidRelativePathError(code, label, "must not contain colon characters");
  }

  if (WINDOWS_INVALID_FILENAME_CHARS.test(value)) {
    throw invalidRelativePathError(code, label, 'must not contain < > " | ? or * characters');
  }

  if (value.startsWith("/") || value.startsWith("//") || win32.isAbsolute(value)) {
    throw invalidRelativePathError(code, label, "must be a relative path");
  }

  for (const segment of value.split("/")) {
    if (segment === "" || segment === "." || segment === "..") {
      throw invalidRelativePathError(
        code,
        label,
        "must not contain empty, current-directory, or parent-directory path segments",
      );
    }

    if (segment.endsWith(".") || segment.endsWith(" ")) {
      throw invalidRelativePathError(
        code,
        label,
        "must not contain path segments ending in a dot or space",
      );
    }

    if (WINDOWS_RESERVED_DEVICE_BASENAME.test(segment)) {
      throw invalidRelativePathError(
        code,
        label,
        "must not use Windows-reserved device names as path segments",
      );
    }
  }
}

function invalidRelativePathError(
  code: "TRANSPORT_ID_INVALID" | "TRANSPORT_DESTINATION_INVALID",
  label: string,
  reason: string,
): TransportError {
  return new TransportError(
    code,
    `The ${label} ${reason}.`,
    "Use a stable relative ID such as docs/handbook without traversal, platform separators, or drive syntax.",
  );
}

function ensureDocxExtension(destination: string): string {
  return destination.toLowerCase().endsWith(".docx") ? destination : `${destination}.docx`;
}

function containsWindowsInvalidControlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0);

    if (codePoint !== undefined && codePoint <= 31) {
      return true;
    }
  }

  return false;
}
