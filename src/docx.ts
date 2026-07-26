/**
 * Pure DOCX archive inspection: no Pandoc, no filesystem access.
 *
 * Pandoc's DOCX reader silently discards several kinds of Word content. The
 * fidelity contract for reverse conversion is that nothing is dropped without
 * being surfaced, so the archive is inspected directly and every known loss is
 * reported back to the caller alongside the Markdown.
 *
 * Every behavior asserted in a summary string below was verified by feeding
 * Pandoc 3.10 a DOCX containing that construct and reading the Markdown it
 * produced. Notably, the DrawingML/VML text box split is not the intuitive one:
 * see `analyzeTextBoxes`.
 *
 * ## Why this scans XML with regular expressions instead of parsing it
 *
 * This is deliberate, and it should stay this way:
 *
 * - These are structural presence and count checks ("is there a `w:ins`
 *   anywhere"), not queries that need a tree.
 * - WordprocessingML entity-escapes `<` in text content, so an open-tag pattern
 *   can never collide with user-authored text. A `<w:ins` match is always
 *   markup.
 * - A parser would fix none of the known failure modes below, while adding an
 *   XML entity-expansion denial-of-service surface to a path that already takes
 *   untrusted bytes.
 *
 * Known, accepted failure modes:
 *
 * - Element counting is namespace-prefix-agnostic, so an element with the same
 *   local name in a foreign namespace would be counted. None of the local names
 *   used here collide in practice inside a WordprocessingML part.
 * - Tag scanning assumes `>` does not appear unescaped inside an attribute
 *   value. That is legal XML but Word does not emit it.
 * - Namespaces declared inline on `mc:AlternateContent`, rather than on the
 *   document root, change what Pandoc does without changing what this reports:
 *   Pandoc drops such a text box while the report still calls it inlined. The
 *   `text-box` item still fires, so the caller is still told to review it and
 *   only the prose misdescribes the outcome. Word declares on the root.
 * - On unbalanced XML an unclosed region never terminates, so its content falls
 *   through into the `analyzeTextBoxes` remainder and can be counted twice.
 *   Malformed input only; a size-mismatched archive is rejected before this.
 */

import { unzipSync } from "fflate";

import { PandocError } from "./errors.js";

/** Package relationship part naming the main document part. */
export const DOCX_PACKAGE_RELS_PART = "_rels/.rels";
/** Used when the package relationships do not name a main document part. */
export const DEFAULT_DOCX_DOCUMENT_PART = "word/document.xml";
/** Archive prefix holding embedded images and other media. */
export const DOCX_MEDIA_PREFIX = "word/media/";
/** Archive prefix holding embedded OLE objects. */
export const DOCX_EMBEDDINGS_PREFIX = "word/embeddings/";
/** Archive prefix holding SmartArt graphics. */
export const DOCX_DIAGRAMS_PREFIX = "word/diagrams/";
/** Archive prefix holding charts. */
export const DOCX_CHARTS_PREFIX = "word/charts/";

/**
 * Largest single part this inspector will decompress.
 *
 * Word's own main document part is low single-digit megabytes even for a large
 * manuscript. The cap exists because a hostile archive can reach a ratio near
 * 300:1, so a few hundred kilobytes of input would otherwise inflate into
 * hundreds of megabytes of resident memory.
 */
export const DOCX_MAX_PART_BYTES = 64 * 1024 * 1024;

/**
 * Largest total this inspector will decompress across every part of one
 * archive.
 *
 * The per-part cap alone is not a memory bound. The set of parts read is not
 * fixed — headers and footers are enumerated from the archive — so an archive
 * can stay under the per-part cap on every entry and still force gigabytes of
 * decompression in aggregate. This budget is what actually holds, and it holds
 * regardless of which part kind is enumerated next.
 *
 * Note that this bounds bytes decompressed, not peak resident memory, which
 * runs roughly 2.4x higher: the inflated `Uint8Array`s and the UTF-16 strings
 * decoded from them coexist, on top of the caller's copy of the input. An
 * archive sitting just under this budget peaks around 300 MB.
 */
export const DOCX_MAX_TOTAL_BYTES = 128 * 1024 * 1024;

/** Largest number of archive entries this inspector will walk. */
export const DOCX_MAX_ARCHIVE_ENTRIES = 4096;

/**
 * Largest number of header and footer parts this inspector will read. Word
 * writes at most six per section, so anything near this is already pathological.
 */
export const DOCX_MAX_HEADER_FOOTER_PARTS = 64;

/** ZIP local file header, the first four bytes of every DOCX. */
const ZIP_LOCAL_FILE_HEADER = [0x50, 0x4b, 0x03, 0x04] as const;

const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/**
 * How Pandoc resolves Word revision marks.
 *
 * `accept` produces the clean Markdown a reviewed import wants. `all` keeps
 * insertions, deletions, and comments as HTML spans instead of dropping them.
 */
export type DocxTrackChangesMode = "accept" | "reject" | "all";

export const DOCX_TRACK_CHANGES_MODES: readonly DocxTrackChangesMode[] = [
  "accept",
  "reject",
  "all",
];

export const DEFAULT_DOCX_TRACK_CHANGES: DocxTrackChangesMode = "accept";

export type DocxUnmappableKind =
  | "chart"
  | "comment"
  | "embedded-media"
  | "embedded-object"
  | "external-content"
  | "header-footer"
  | "smart-art"
  | "text-box"
  | "tracked-change";

export interface DocxUnmappableItem {
  /** Stable machine-readable discriminator for this class of loss. */
  readonly kind: DocxUnmappableKind;
  /** Human-readable explanation of what was found and what Pandoc did with it. */
  readonly summary: string;
  /** How many of this item the archive contains. */
  readonly count: number;
  /** Archive part names, for the kinds that map to concrete files. */
  readonly entries?: readonly string[];
}

export interface DocxTrackedChangeCounts {
  /** `w:ins` and `w:moveTo` elements. */
  readonly insertions: number;
  /** `w:del` and `w:moveFrom` elements. */
  readonly deletions: number;
  /** `w:rPrChange` and friends: revisions that restyle rather than edit text. */
  readonly formattingChanges: number;
}

export interface DocxImportReport {
  /**
   * Everything the conversion could not carry into Markdown faithfully.
   * Empty for a document Pandoc can represent completely.
   */
  readonly unmappable: readonly DocxUnmappableItem[];
  /** Convenience flag: `unmappable.length > 0`. */
  readonly hasUnmappableContent: boolean;
  /** The `--track-changes` value the conversion ran with. */
  readonly trackChanges: DocxTrackChangesMode;
  /**
   * Revision marks counted across `scannedParts`.
   *
   * `insertions` and `deletions` count revision-mark elements, which includes
   * paragraph-mark revisions that Word writes on every Enter pressed with
   * tracking on and that cost nothing in Markdown. Formatting-only revisions
   * are not revision marks and are counted separately as `formattingChanges`.
   */
  readonly trackedChanges: DocxTrackedChangeCounts;
  /** The main document part, resolved from the package relationships. */
  readonly documentPart: string;
  /**
   * Every part the structural counters ran over. Content in a part outside this
   * list is covered by a whole-part item, such as `header-footer`, or is not
   * inspected at all.
   */
  readonly scannedParts: readonly string[];
}

interface DocxArchive {
  readonly documentPart: string;
  readonly scannedParts: readonly string[];
  /** Concatenated XML of every scanned content part, exactly as stored. */
  readonly contentXml: string;
  readonly commentsXml: string | undefined;
  readonly headerFooterParts: readonly string[];
  readonly mediaEntries: readonly string[];
  readonly embeddingEntries: readonly string[];
  readonly diagramEntries: readonly string[];
  readonly chartEntries: readonly string[];
}

interface TextBoxCounts {
  readonly inlined: number;
  readonly dropped: number;
}

interface ElementRegion {
  readonly start: number;
  readonly end: number;
  readonly inner: string;
}

interface ArchiveScan {
  /** Running total of declared bytes decompressed from one archive. */
  bytes: number;
  /** CRC-32 per entry, read from the archive's own central directory. */
  readonly checksums: ReadonlyMap<string, number>;
}

function validateDocxTrackChanges(value: unknown): DocxTrackChangesMode {
  if (typeof value === "string" && isTrackChangesMode(value)) {
    return value;
  }

  throw new PandocError(
    "DOCX_TRACK_CHANGES_INVALID",
    `trackChanges must be one of ${DOCX_TRACK_CHANGES_MODES.join(", ")}.`,
    "Pass trackChanges as accept to resolve Word revision marks into clean Markdown, reject to restore the pre-edit text, or all to keep insertions, deletions, and comments as inline spans.",
  );
}

/**
 * Inspects a DOCX archive and reports everything reverse conversion cannot
 * represent in Markdown. Throws when the bytes are not a readable Word
 * document, so garbage never reaches Pandoc.
 */
export function buildDocxImportReport(
  docx: Uint8Array,
  trackChanges: DocxTrackChangesMode = DEFAULT_DOCX_TRACK_CHANGES,
): DocxImportReport {
  const mode = validateDocxTrackChanges(trackChanges);
  const archive = readDocxArchive(docx);
  // `mc:Fallback` restates the preceding `mc:Choice` in a legacy representation
  // rather than adding content, so counting both would double-report one edit.
  // Text boxes are the exception and are analyzed on the unstripped XML,
  // because for them the fallback is the branch Pandoc actually reads.
  const contentXml = stripFallbackRegions(archive.contentXml);
  const trackedChanges: DocxTrackedChangeCounts = {
    insertions: countElements(contentXml, "ins") + countElements(contentXml, "moveTo"),
    deletions: countElements(contentXml, "del") + countElements(contentXml, "moveFrom"),
    formattingChanges: countFormattingRevisions(contentXml),
  };
  const scope = describeParts(archive.scannedParts);
  const unmappable = [
    buildTrackedChangeItem(trackedChanges, mode, scope),
    buildCommentItem(archive.commentsXml, mode),
    buildHeaderFooterItem(archive.headerFooterParts),
    buildTextBoxItem(analyzeTextBoxes(archive.contentXml), scope),
    buildExternalContentItem(countElements(contentXml, "altChunk"), scope),
    buildSmartArtItem(archive.diagramEntries),
    buildChartItem(archive.chartEntries),
    buildMediaItem(archive.mediaEntries),
    buildEmbeddedObjectItem(archive.embeddingEntries),
  ].filter((item): item is DocxUnmappableItem => item !== undefined);

  return {
    unmappable,
    hasUnmappableContent: unmappable.length > 0,
    trackChanges: mode,
    trackedChanges,
    documentPart: archive.documentPart,
    scannedParts: archive.scannedParts,
  };
}

function readDocxArchive(docx: Uint8Array): DocxArchive {
  validateDocxBytes(docx);

  const names: string[] = [];
  // One scan spans both passes, so the total decompressed for an archive is
  // bounded no matter how the work is split across them.
  const scan: ArchiveScan = { bytes: 0, checksums: readCentralDirectoryChecksums(docx) };
  // The first pass parses only the central directory and inflates the tiny
  // package relationships part, so the main document part can be resolved by
  // name before anything large is decompressed.
  const rels = inflateParts(docx, names, (name) => name === DOCX_PACKAGE_RELS_PART, scan);
  const relsXml = rels[DOCX_PACKAGE_RELS_PART];
  const resolved = resolveDocumentPart(relsXml, names);

  if (!names.includes(resolved.part)) {
    throw new PandocError(
      "DOCX_ARCHIVE_INVALID",
      `The input is a ZIP archive with no readable main document part: ${
        relsXml === undefined
          ? `it has no ${DOCX_PACKAGE_RELS_PART} and no ${DEFAULT_DOCX_DOCUMENT_PART}`
          : `its package relationships name no officeDocument target present in the archive, and there is no ${DEFAULT_DOCX_DOCUMENT_PART}`
      }.`,
      "Pass a Word .docx file. Other Office formats, OpenDocument files, and plain ZIP archives are not supported.",
    );
  }

  const documentPart = resolved.part;
  const directory = documentPart.slice(0, documentPart.lastIndexOf("/") + 1);
  const commentsPart = `${directory}comments.xml`;
  const headerFooterPattern = new RegExp(`^${escapeRegExp(directory)}(?:header|footer)\\d*\\.xml$`);
  // Footnote and endnote text does flow into the Markdown, so revisions inside
  // them belong in the same counters as the body.
  const notesParts = [`${directory}footnotes.xml`, `${directory}endnotes.xml`].filter((part) =>
    names.includes(part),
  );
  const headerFooterCandidates = names.filter((name) => headerFooterPattern.test(name)).sort();

  if (headerFooterCandidates.length > DOCX_MAX_HEADER_FOOTER_PARTS) {
    throw new PandocError(
      "DOCX_ARCHIVE_INVALID",
      `The DOCX archive declares ${headerFooterCandidates.length} header and footer parts, over the ${DOCX_MAX_HEADER_FOOTER_PARTS} part limit.`,
      "Pass a Word .docx file. Word writes at most six header and footer parts per section, so a count this high indicates a corrupt or hostile archive.",
    );
  }

  const scannedParts = [documentPart, ...notesParts];
  const wanted = new Set([...scannedParts, commentsPart, ...headerFooterCandidates]);
  const parts = inflateParts(docx, undefined, (name) => wanted.has(name), scan);

  return {
    documentPart,
    scannedParts,
    contentXml: scannedParts.map((part) => parts[part] ?? "").join("\n"),
    commentsXml: parts[commentsPart],
    // A header or footer with no text is Word boilerplate, not lost content.
    headerFooterParts: headerFooterCandidates.filter((part) =>
      containsElement(parts[part] ?? "", "t"),
    ),
    mediaEntries: filterEntries(names, DOCX_MEDIA_PREFIX),
    embeddingEntries: filterEntries(names, DOCX_EMBEDDINGS_PREFIX),
    diagramEntries: filterEntries(names, DOCX_DIAGRAMS_PREFIX),
    chartEntries: filterEntries(names, DOCX_CHARTS_PREFIX),
  };
}

/**
 * Decompresses the parts `wanted` selects, optionally recording every entry
 * name into `names`. The filter callback runs once per archive entry, so it
 * doubles as the inventory and as the place to enforce the archive limits
 * before any decompression happens.
 */
function inflateParts(
  docx: Uint8Array,
  names: string[] | undefined,
  wanted: (name: string) => boolean,
  scan: ArchiveScan,
): Record<string, string> {
  const declaredSizes = new Map<string, number>();
  let entryCount = 0;
  let files: Record<string, Uint8Array>;

  try {
    files = unzipSync(docx, {
      filter: (file) => {
        entryCount += 1;

        if (entryCount > DOCX_MAX_ARCHIVE_ENTRIES) {
          throw new PandocError(
            "DOCX_ARCHIVE_INVALID",
            `The DOCX archive declares more than ${DOCX_MAX_ARCHIVE_ENTRIES} entries.`,
            "Pass a Word .docx file. An archive this large is not a document Word produced.",
          );
        }

        if (names !== undefined && !file.name.endsWith("/")) {
          names.push(file.name);
        }

        if (!wanted(file.name)) {
          return false;
        }

        // `originalSize` comes from the central directory, and fflate allocates
        // exactly that many bytes to inflate into, so it bounds the allocation
        // even when an archive lies about it. A lie in the other direction is
        // caught after inflation, below.
        if (file.originalSize > DOCX_MAX_PART_BYTES) {
          throw new PandocError(
            "DOCX_ARCHIVE_INVALID",
            `The DOCX part ${file.name} declares ${file.originalSize} bytes, over the ${DOCX_MAX_PART_BYTES} byte limit.`,
            "Pass a Word .docx file. A part this large indicates a corrupt or hostile archive rather than a document.",
          );
        }

        scan.bytes += file.originalSize;

        if (scan.bytes > DOCX_MAX_TOTAL_BYTES) {
          throw new PandocError(
            "DOCX_ARCHIVE_INVALID",
            `The DOCX parts this inspector must read declare more than ${DOCX_MAX_TOTAL_BYTES} bytes in total.`,
            "Pass a Word .docx file. Many individually plausible parts adding up to this much data indicate a corrupt or hostile archive.",
          );
        }

        declaredSizes.set(file.name, file.originalSize);

        return true;
      },
    });
  } catch (cause) {
    // The limit checks above throw through fflate, so they must not be remapped
    // into the generic unreadable-archive error.
    if (cause instanceof PandocError) {
      throw cause;
    }

    throw new PandocError(
      "DOCX_ARCHIVE_INVALID",
      "The DOCX input could not be read as a ZIP archive.",
      "Pass an intact .docx file. A truncated download or a file rewritten by a text editor produces this error.",
      { cause },
    );
  }

  const decoded: Record<string, string> = {};
  const decoder = new TextDecoder();

  for (const [name, bytes] of Object.entries(files)) {
    // fflate inflates a deflate entry into a buffer pre-sized from the central
    // directory, so a part declaring less than it really holds is silently
    // truncated and analyzing the fragment would produce a clean report for a
    // document Pandoc reads in full. Comparing lengths cannot catch that: the
    // declared size is what sized the buffer, so it always matches. The CRC is
    // the archive's own integrity mechanism and catches both directions.
    // A stored entry is sliced by its compressed size, so an understated
    // uncompressed size leaves the content intact but the index inconsistent.
    // That is caught here; it is a no-op for a deflate entry, where this
    // comparison is against the very number that sized the buffer.
    if (bytes.length !== declaredSizes.get(name)) {
      throw new PandocError(
        "DOCX_ARCHIVE_INVALID",
        `The DOCX part ${name} holds ${bytes.length} bytes but its archive entry declares ${declaredSizes.get(name)}.`,
        "Pass an intact .docx file. A size mismatch means the archive index disagrees with its contents, so no report over it could be trusted.",
      );
    }

    const expected = scan.checksums.get(name);

    if (expected === undefined || expected !== crc32(bytes)) {
      throw new PandocError(
        "DOCX_ARCHIVE_INVALID",
        `The DOCX part ${name} does not match the CRC-32 recorded in the archive index.`,
        "Pass an intact .docx file. A checksum mismatch means the archive index disagrees with its contents, so the part may be truncated and no report over it could be trusted.",
      );
    }

    decoded[name] = decoder.decode(bytes);
  }

  return decoded;
}

/**
 * Reads the CRC-32 of every entry from the archive's central directory.
 *
 * fflate keeps its own `crc` helper internal and exposes no checksum on the
 * filter callback, so the directory is walked here. The traversal mirrors
 * fflate's own end-of-central-directory discovery, including its ZIP64 handling,
 * so the two agree on which records exist.
 */
function readCentralDirectoryChecksums(docx: Uint8Array): ReadonlyMap<string, number> {
  const checksums = new Map<string, number>();

  if (docx.byteLength < 22) {
    return checksums;
  }

  const view = new DataView(docx.buffer, docx.byteOffset, docx.byteLength);
  let end = docx.byteLength - 22;

  while (view.getUint32(end, true) !== 0x0605_4b50) {
    if (end === 0 || docx.byteLength - end > 65_558) {
      return checksums;
    }

    end -= 1;
  }

  let entries = view.getUint16(end + 8, true);
  let offset = view.getUint32(end + 16, true);

  if (end >= 20 && view.getUint32(end - 20, true) === 0x0706_4b50) {
    const zip64End = view.getUint32(end - 12, true);

    if (zip64End + 52 <= docx.byteLength && view.getUint32(zip64End, true) === 0x0606_4b50) {
      entries = view.getUint32(zip64End + 32, true);
      offset = view.getUint32(zip64End + 48, true);
    }
  }

  const decoder = new TextDecoder();

  for (let index = 0; index < entries; index += 1) {
    if (offset + 46 > docx.byteLength || view.getUint32(offset, true) !== 0x0201_4b50) {
      break;
    }

    const nameLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 46;

    if (nameStart + nameLength > docx.byteLength) {
      break;
    }

    checksums.set(
      decoder.decode(docx.subarray(nameStart, nameStart + nameLength)),
      view.getUint32(offset + 16, true),
    );
    offset =
      nameStart +
      nameLength +
      view.getUint16(offset + 30, true) +
      view.getUint16(offset + 32, true);
  }

  return checksums;
}

const CRC32_TABLE = /* @__PURE__ */ (() => {
  const table = new Int32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;

    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? (value >>> 1) ^ 0xedb8_8320 : value >>> 1;
    }

    table[index] = value;
  }

  return table;
})();

/** Standard ZIP CRC-32. */
function crc32(bytes: Uint8Array): number {
  let checksum = -1;

  for (let index = 0; index < bytes.length; index += 1) {
    checksum = (checksum >>> 8) ^ (CRC32_TABLE[(checksum ^ (bytes[index] ?? 0)) & 0xff] ?? 0);
  }

  return (checksum ^ -1) >>> 0;
}

/**
 * Resolves the main document part from the package relationships.
 *
 * The name is not fixed by the specification. `word/document2.xml` occurs in
 * the wild and Pandoc converts it happily, so assuming `word/document.xml`
 * would reject valid Word documents.
 */
function resolveDocumentPart(
  relsXml: string | undefined,
  names: readonly string[],
): { part: string; fromRels: boolean } {
  // The element may carry a namespace prefix, so match on the local name.
  for (const tag of relsXml?.match(/<(?:[A-Za-z_][\w.-]*:)?Relationship(?=[\s/>])[^>]*>/g) ?? []) {
    if (readAttribute(tag, "Type") !== OFFICE_DOCUMENT_RELATIONSHIP) {
      continue;
    }

    // Package-relative targets may be written absolute or explicitly relative.
    const target = (readAttribute(tag, "Target") ?? "").replace(/^\/+/, "").replace(/^\.\//, "");

    if (target !== "" && names.includes(target)) {
      return { part: target, fromRels: true };
    }
  }

  return { part: DEFAULT_DOCX_DOCUMENT_PART, fromRels: false };
}

function validateDocxBytes(docx: Uint8Array): void {
  if (!(docx instanceof Uint8Array)) {
    throw new PandocError(
      "DOCX_INPUT_INVALID",
      "The DOCX input must be a Uint8Array of file bytes.",
      "Pass DOCX bytes, or a path to a .docx file for the library to read.",
    );
  }

  if (docx.byteLength === 0) {
    throw new PandocError(
      "DOCX_INPUT_INVALID",
      "The DOCX input is empty.",
      "Pass a non-empty .docx file. An empty file usually means an interrupted download or an unwritten export.",
    );
  }

  if (!ZIP_LOCAL_FILE_HEADER.every((byte, index) => docx[index] === byte)) {
    throw new PandocError(
      "DOCX_INPUT_INVALID",
      "The DOCX input does not start with a ZIP local file header, so it is not a .docx file.",
      "Pass a real Word .docx file. Legacy .doc, RTF, and PDF files are different formats and are not supported.",
    );
  }
}

function buildTrackedChangeItem(
  counts: DocxTrackedChangeCounts,
  mode: DocxTrackChangesMode,
  scope: string,
): DocxUnmappableItem | undefined {
  const revisions = counts.insertions + counts.deletions;

  // `all` keeps insertions and deletions as inline spans, so only the
  // formatting revisions it cannot represent remain a loss in that mode.
  if (mode === "all") {
    if (counts.formattingChanges === 0) {
      return undefined;
    }

    return {
      kind: "tracked-change",
      count: counts.formattingChanges,
      summary: `${plural(
        counts.formattingChanges,
        "formatting revision",
      )} in ${scope}. Pandoc ran --track-changes=all, which preserves insertions and deletions as inline spans but has no representation for a formatting-only revision, so the superseded formatting is gone.`,
    };
  }

  if (revisions + counts.formattingChanges === 0) {
    return undefined;
  }

  const resolution =
    mode === "accept"
      ? "insertions were kept as ordinary text and deletions were discarded"
      : "deletions were restored as ordinary text and insertions were discarded";
  const formattingNote =
    counts.formattingChanges === 0
      ? ""
      : ` A further ${plural(
          counts.formattingChanges,
          "formatting revision",
        )} discarded the superseded formatting.`;

  return {
    kind: "tracked-change",
    count: revisions + counts.formattingChanges,
    summary: `${plural(revisions, "tracked change")} in ${scope} (${plural(
      counts.insertions,
      "insertion",
    )}, ${plural(
      counts.deletions,
      "deletion",
    )}, counting moved text as both). Pandoc ran --track-changes=${mode}, so ${resolution}; the revision marks and their authors are not represented in the Markdown.${formattingNote}`,
  };
}

function buildCommentItem(
  commentsXml: string | undefined,
  mode: DocxTrackChangesMode,
): DocxUnmappableItem | undefined {
  if (commentsXml === undefined) {
    return undefined;
  }

  // Pandoc writes an empty `<w:comments/>` part into every DOCX it generates,
  // so the part existing proves nothing. Only comment bodies count.
  const count = countElements(commentsXml, "comment");

  if (count === 0 || mode === "all") {
    return undefined;
  }

  return {
    kind: "comment",
    count,
    summary: `${plural(count, "Word comment")}. Pandoc only emits comment text with --track-changes=all; this conversion ran --track-changes=${mode}, so the comments were dropped and appear nowhere in the Markdown.`,
  };
}

function buildHeaderFooterItem(parts: readonly string[]): DocxUnmappableItem | undefined {
  if (parts.length === 0) {
    return undefined;
  }

  return {
    kind: "header-footer",
    count: parts.length,
    entries: parts,
    summary: `${plural(
      parts.length,
      "header or footer part",
      "header and footer parts",
    )} containing text. Pandoc's DOCX reader ignores headers and footers entirely, so running titles, page furniture, and any content parked there are absent from the Markdown.`,
  };
}

function buildTextBoxItem(counts: TextBoxCounts, scope: string): DocxUnmappableItem | undefined {
  const total = counts.inlined + counts.dropped;

  if (total === 0) {
    return undefined;
  }

  return {
    kind: "text-box",
    count: total,
    summary: `${plural(total, "text box", "text boxes")} in ${scope} (${
      counts.inlined
    } inlined as ordinary body paragraphs, losing their text-box framing and position; ${
      counts.dropped
    } dropped outright). Pandoc 3.x reaches text box content only through the legacy VML representation: a DrawingML text box carrying an mc:Fallback, which is what Word 2010 and newer writes, is read from that fallback, while one without a fallback is discarded along with its text.`,
  };
}

function buildExternalContentItem(count: number, scope: string): DocxUnmappableItem | undefined {
  if (count === 0) {
    return undefined;
  }

  return {
    kind: "external-content",
    count,
    summary: `${plural(
      count,
      "altChunk import",
    )} in ${scope}. An altChunk embeds another document, usually HTML or a nested DOCX, for Word to merge in. Pandoc's DOCX reader ignores it, so all of that text is missing from the Markdown.`,
  };
}

function buildSmartArtItem(entries: readonly string[]): DocxUnmappableItem | undefined {
  if (entries.length === 0) {
    return undefined;
  }

  // One SmartArt graphic spans several parts; the data part is the one holding
  // the node text, so it is what makes the graphic countable.
  const count = countMatchingEntries(entries, /\/data\d*\.xml$/);

  return {
    kind: "smart-art",
    count,
    entries,
    summary: `${plural(
      count,
      "SmartArt graphic",
    )} under ${DOCX_DIAGRAMS_PREFIX}. Pandoc's DOCX reader drops SmartArt entirely, and its node labels are ordinary prose, so every word in the diagram is missing from the Markdown.`,
  };
}

function buildChartItem(entries: readonly string[]): DocxUnmappableItem | undefined {
  if (entries.length === 0) {
    return undefined;
  }

  const count = countMatchingEntries(entries, /\/chart\d*\.xml$/);

  return {
    kind: "chart",
    count,
    entries,
    summary: `${plural(
      count,
      "chart",
    )} under ${DOCX_CHARTS_PREFIX}. Pandoc's DOCX reader drops charts entirely, including their titles and axis labels. A chart also stores its data as a separate embedded workbook, which is reported on its own.`,
  };
}

function buildMediaItem(entries: readonly string[]): DocxUnmappableItem | undefined {
  if (entries.length === 0) {
    return undefined;
  }

  return {
    kind: "embedded-media",
    count: entries.length,
    entries,
    summary: `${plural(
      entries.length,
      "embedded media file",
    )} under ${DOCX_MEDIA_PREFIX}. Pandoc references media by path, such as media/image1.png, but does not write the bytes unless --extract-media is used, so those Markdown references point at files that do not exist.`,
  };
}

function buildEmbeddedObjectItem(entries: readonly string[]): DocxUnmappableItem | undefined {
  if (entries.length === 0) {
    return undefined;
  }

  return {
    kind: "embedded-object",
    count: entries.length,
    entries,
    summary: `${plural(
      entries.length,
      "embedded OLE object",
    )} under ${DOCX_EMBEDDINGS_PREFIX}. Pandoc's DOCX reader drops embedded objects entirely, so neither the object nor a placeholder appears in the Markdown.`,
  };
}

/**
 * Classifies text boxes by whether Pandoc can reach their content.
 *
 * The intuitive split, DrawingML versus legacy VML, is the wrong one. Verified
 * against Pandoc 3.10:
 *
 * - bare legacy VML `w:pict`/`v:textbox` is inlined
 * - DrawingML inside `mc:AlternateContent` *with* an `mc:Fallback` is inlined,
 *   read out of the fallback rather than out of the `mc:Choice`
 * - DrawingML with no fallback, wrapped or bare, is dropped
 *
 * Word 2010 and newer always writes the fallback, so the common real-world case
 * is preserved. Classifying on the post-strip namespace prefix would report
 * exactly that case as a total loss.
 */
function analyzeTextBoxes(xml: string): TextBoxCounts {
  let inlined = 0;
  let dropped = 0;
  let remainder = "";
  let cursor = 0;

  for (const region of findElementRegions(xml, "AlternateContent")) {
    remainder += xml.slice(cursor, region.start);
    cursor = region.end;

    const hasFallbackTextBox = findElementRegions(region.inner, "Fallback").some((fallback) =>
      containsElement(fallback.inner, "txbxContent"),
    );

    if (hasFallbackTextBox) {
      inlined += 1;
    } else if (containsElement(region.inner, "txbxContent")) {
      dropped += 1;
    }
  }

  remainder += xml.slice(cursor);

  // Outside any AlternateContent, a text box is dropped only when its content
  // sits directly in a DrawingML `wps:txbx` with no legacy representation.
  const remainderTotal = countElements(remainder, "txbxContent");
  const remainderDropped = Math.min(countElements(remainder, "txbx"), remainderTotal);

  return {
    inlined: inlined + (remainderTotal - remainderDropped),
    dropped: dropped + remainderDropped,
  };
}

/**
 * Removes `mc:Fallback` regions, which restate the preceding `mc:Choice`
 * content in a legacy representation rather than adding new content.
 */
function stripFallbackRegions(xml: string): string {
  let stripped = "";
  let cursor = 0;

  for (const region of findElementRegions(xml, "Fallback")) {
    stripped += xml.slice(cursor, region.start);
    cursor = region.end;
  }

  return stripped + xml.slice(cursor);
}

/**
 * Finds the outermost balanced regions of an element by local name.
 *
 * Depth tracking is what makes this correct for a nested `mc:AlternateContent`
 * inside an `mc:Fallback`, which is legal and which a non-greedy pattern would
 * mis-terminate at the inner closing tag.
 */
function findElementRegions(xml: string, localName: string): readonly ElementRegion[] {
  const pattern = new RegExp(`<(/)?(?:[A-Za-z_][\\w.-]*:)?${localName}(?=[\\s/>])([^>]*)>`, "g");
  const regions: ElementRegion[] = [];
  let depth = 0;
  let start = 0;
  let innerStart = 0;
  let match = pattern.exec(xml);

  while (match !== null) {
    const isClosing = match[1] === "/";
    const isSelfClosing = (match[2] ?? "").endsWith("/");

    if (!isClosing && !isSelfClosing) {
      if (depth === 0) {
        start = match.index;
        innerStart = match.index + match[0].length;
      }

      depth += 1;
    } else if (isClosing && depth > 0) {
      depth -= 1;

      if (depth === 0) {
        regions.push({
          start,
          end: match.index + match[0].length,
          inner: xml.slice(innerStart, match.index),
        });
      }
    }

    match = pattern.exec(xml);
  }

  return regions;
}

/**
 * Counts elements by local name, ignoring the namespace prefix.
 *
 * The lookahead is load-bearing: without it `ins` would match the table border
 * elements `w:insideH` and `w:insideV`, `del` would match `w:delText`,
 * `moveFrom` would match `w:moveFromRangeStart`, `txbx` would match
 * `w:txbxContent`, and `comment` would match the `w:comments` root element.
 */
function countElements(xml: string, localName: string): number {
  return (xml.match(elementPattern(localName)) ?? []).length;
}

function containsElement(xml: string, localName: string): boolean {
  return elementPattern(localName).test(xml);
}

function elementPattern(localName: string): RegExp {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}(?=[\\s/>])`, "g");
}

/** Counts `w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, and the rest. */
function countFormattingRevisions(xml: string): number {
  return (xml.match(/<(?:[A-Za-z_][\w.-]*:)?\w*PrChange(?=[\s/>])/g) ?? []).length;
}

function countMatchingEntries(entries: readonly string[], pattern: RegExp): number {
  const matched = entries.filter((entry) => pattern.test(entry)).length;

  return matched === 0 ? entries.length : matched;
}

function filterEntries(names: readonly string[], prefix: string): readonly string[] {
  return names.filter((name) => name.startsWith(prefix)).sort();
}

function describeParts(parts: readonly string[]): string {
  if (parts.length === 1) {
    return parts[0] ?? "";
  }

  return `${parts.slice(0, -1).join(", ")} and ${parts.at(-1)}`;
}

/** Reads an XML attribute value, accepting either legal quote style. */
function readAttribute(tag: string, name: string): string | undefined {
  const match = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);

  return match?.[1] ?? match?.[2];
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, String.raw`\$&`);
}

function isTrackChangesMode(value: string): value is DocxTrackChangesMode {
  return (DOCX_TRACK_CHANGES_MODES as readonly string[]).includes(value);
}

function plural(count: number, singular: string, pluralForm = `${singular}s`): string {
  return `${count} ${count === 1 ? singular : pluralForm}`;
}
