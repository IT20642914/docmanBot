/**
 * Identifies document type/source from document title (filename).
 * Used when user sends a file - we analyze the name to provide specific responses.
 *
 * Uses parseDocumentFilename for IFS format detection (CLASS - NUMBER - SHEET - REV).
 */

import { parseDocumentFilename, type ParsedDocumentInfo } from "./parseDocumentFilename";

export interface DocumentTypeResult {
  /** Identified document type, e.g. "IFS", "SAP", "Oracle" */
  type: string;
  /** Human-readable description for the bot response */
  message: string;
  /** Whether we identified a known document type */
  isIdentified: boolean;
  /** Parsed IFS metadata when document is IFS format */
  ifsMetadata?: ParsedDocumentInfo;
}

/**
 * Identifies document type from the document title (filename).
 * Uses parseDocumentFilename for IFS format (CLASS - NUMBER - SHEET - REV).
 * @param documentTitle - Filename or document title (e.g. "Title (01-TEST - 1028340 - 1 - A1) - 1.docx")
 * @returns DocumentTypeResult with type, message, isIdentified, and optional ifsMetadata
 */
export function identifyDocumentType(documentTitle: string): DocumentTypeResult {
  if (!documentTitle || typeof documentTitle !== "string") {
    return {
      type: "Unknown",
      message: "I couldn't determine the document type.",
      isIdentified: false,
    };
  }

  // First: try IFS metadata parsing (CLASS - NUMBER - SHEET - REV format)
  const parsed = parseDocumentFilename(documentTitle);
  if (parsed?.isIFSFormat) {
    const { title, docClass, docNo, docSheet, docRev } = parsed;
    const metadataStr = `${docClass} - ${docNo} - ${docSheet} - ${docRev}`;
    return {
      type: "IFS",
      message:
        "This document is from **IFS** (IFS Applications).\n\n" +
        "**Title:** " +
        title +
        "\n\n" +
        "**Metadata:** " +
        metadataStr +
        "\n\n" +
        "*(Class: " +
        docClass +
        ", Doc No: " +
        docNo +
        ", Sheet: " +
        docSheet +
        ", Rev: " +
        docRev +
        ")*",
      isIdentified: true,
      ifsMetadata: parsed,
    };
  }

  return {
    type: "NotIFS",
    message: "This document is not from IFS.",
    isIdentified: false,
  };
}
