/**
 * Parses document filename to extract IFS metadata.
 * TypeScript/Node.js port of the React useDocumentInfo parseDocumentFilename logic.
 *
 * Handles IFS document naming format:
 * - "Title (01-TEST - 1028340 - 1 - A1) - 1.docx"
 * - "Title (extra text)(01-TEST - 1028340 - 1 - A1) - 1.docx"
 * - "150 mb file docx (sri lanaka)(01-TEST - 1028340 - 1 - A1) - 1.docx"
 *
 * Format: CLASS - NUMBER - SHEET - REV
 * - CLASS: alphanumeric with hyphens/underscores (e.g. "01-TEST", "ISU_YES_N")
 * - NUMBER: digits (e.g. "1028340")
 * - SHEET: digits (e.g. "1")
 * - REV: alphanumeric (e.g. "A1")
 */

export interface ParsedDocumentInfo {
  title: string;
  docClass: string;
  docNo: string;
  docSheet: string;
  docRev: string;
  fileExtension: string;
  isIFSFormat: boolean;
  isCopy: boolean;
}

/**
 * Check if a string contains IFS metadata format (CLASS - NUMBER - SHEET - REV)
 */
function isIFSMetadataPattern(str: string): boolean {
  const tokens = str.split(/[-_\s]+/).map((t) => t.trim()).filter(Boolean);

  if (tokens.length < 4) return false;

  const docNo = tokens[tokens.length - 3];
  const docSheet = tokens[tokens.length - 2];
  const docRev = tokens[tokens.length - 1];

  return (
    /^\d+$/.test(docNo) &&
    /^\d+$/.test(docSheet) &&
    /^[A-Z0-9]+$/i.test(docRev)
  );
}

/**
 * Parses document filename to extract IFS metadata.
 * @param filename - Document filename (e.g. "Title (01-TEST - 1028340 - 1 - A1) - 1.docx")
 * @returns ParsedDocumentInfo with title, docClass, docNo, docSheet, docRev, or null if invalid
 */
export function parseDocumentFilename(filename: string | null | undefined): ParsedDocumentInfo | null {
  if (!filename || typeof filename !== "string") return null;

  // Clean prefixes like "Copy of " / "Word add-in "
  let cleanFileName = filename.replace(/^Copy of\s+/i, "").trim();

  // Extract extension first
  const dot = cleanFileName.lastIndexOf(".");
  const extension = dot >= 0 ? cleanFileName.slice(dot) : "";
  const nameWithoutExt = dot >= 0 ? cleanFileName.slice(0, dot) : cleanFileName;

  // Remove version suffix like "(1).docx" or " - 1" before extension
  const withoutVersion = nameWithoutExt.replace(/\s*-\s*\d+\s*$/, "");

  // Find bracket groups (...) or [...]
  const parenMatches = [...withoutVersion.matchAll(/\(([^)]+)\)/g)];
  const bracketMatches = [...withoutVersion.matchAll(/\[([^\]]+)\]/g)];
  const allBracketMatches = [...parenMatches, ...bracketMatches].sort(
    (a, b) => (a.index ?? 0) - (b.index ?? 0)
  );

  let metadataString: string | null = null;
  let metadataMatch: RegExpMatchArray | null = null;
  let titlePart = withoutVersion;

  if (allBracketMatches.length > 0) {
    for (let i = allBracketMatches.length - 1; i >= 0; i--) {
      const match = allBracketMatches[i];
      const content = match[1];
      if (isIFSMetadataPattern(content)) {
        metadataString = content;
        metadataMatch = match;
        break;
      }
    }

    if (metadataMatch) {
      titlePart = withoutVersion.slice(0, metadataMatch.index ?? 0).trim();
      titlePart = titlePart.replace(/[\s()[\]\]]+$/, "").trim();
      const remainingBrackets = titlePart.match(/\(.+?\)|\[.+?\]/g);
      if (remainingBrackets) {
        titlePart = titlePart.replace(/\(.+?\)|\[.+?\]/g, "").trim();
      }
    } else {
      metadataString = null;
    }
  }

  // If no brackets found, try alternative pattern matching
  if (!metadataString) {
    const altPattern = /([A-Z0-9-]+)\s*[-_]\s*(\d+)\s*[-_]\s*(\d+)\s*[-_]\s*([A-Z0-9]+)/i;
    const altMatch = withoutVersion.match(altPattern);

    if (altMatch) {
      const [, docClass, docNo, docSheet, docRev] = altMatch;
      const matchIndex = altMatch.index ?? 0;
      titlePart = withoutVersion.slice(0, matchIndex).trim();

      return {
        title: titlePart || cleanFileName.replace(/\.[^.]*$/, ""),
        docClass: (docClass ?? "").trim(),
        docNo: (docNo ?? "").trim(),
        docSheet: (docSheet ?? "").trim(),
        docRev: (docRev ?? "").trim().toUpperCase(),
        fileExtension: extension ? extension.toUpperCase() : "",
        isIFSFormat: true,
        isCopy: /^copy of\s+/i.test(filename),
      };
    }

    return {
      title: titlePart || nameWithoutExt,
      docClass: "",
      docNo: "",
      docSheet: "",
      docRev: "",
      fileExtension: extension ? extension.toUpperCase() : "",
      isIFSFormat: false,
      isCopy: /^copy of\s+/i.test(filename),
    };
  }

  return parseMetadataTokens(metadataString, titlePart, nameWithoutExt, extension, filename);
}

/**
 * Parse metadata string "CLASS - NUMBER - SHEET - REV" into structured parts
 */
function parseMetadataTokens(
  metadataString: string,
  titlePart: string,
  nameWithoutExt: string,
  extension: string,
  filename: string
): ParsedDocumentInfo {
  const tokens = metadataString
    .split(/\s+-\s+/)
    .map((t) => t.trim())
    .filter(Boolean);

  if (tokens.length < 4) {
    return {
      title: titlePart || nameWithoutExt,
      docClass: "",
      docNo: "",
      docSheet: "",
      docRev: "",
      fileExtension: extension ? extension.toUpperCase() : "",
      isIFSFormat: false,
      isCopy: /^copy of\s+/i.test(filename),
    };
  }

  const docRev = tokens.pop() ?? "";
  const docSheet = tokens.pop() ?? "";
  const docNo = tokens.pop() ?? "";
  const docClass = tokens.join(" - ").trim();

  return {
    title: titlePart || "Untitled",
    docClass: (docClass ?? "").trim(),
    docNo: (docNo ?? "").trim(),
    docSheet: (docSheet ?? "").trim(),
    docRev: (docRev ?? "").trim().toUpperCase(),
    fileExtension: extension ? extension.toUpperCase() : "",
    isIFSFormat: true,
    isCopy: /^copy of\s+/i.test(filename),
  };
}

/**
 * Formats parsed document info back to IFS filename format.
 * @param info - Parsed document info from parseDocumentFilename
 * @returns Formatted filename (e.g. "Title (01-TEST - 1028340 - 1 - A1) - 1.docx")
 */
export function formatDocumentName(info: ParsedDocumentInfo): string {
  if (!info) return "";
  const { title, docClass, docNo, docSheet, docRev, fileExtension } = info;
  const base = `${title} (${docClass} - ${docNo} - ${docSheet} - ${docRev}) - 1`;
  return `${base}${fileExtension ? fileExtension : ""}`;
}
