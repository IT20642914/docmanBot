import { readFile } from "node:fs/promises";
import path from "node:path";

export interface PendingApprovalDocument {
  id: string;
  title: string;
  fileName: string;
  /** Relative path (recommended) or absolute path to local text content */
  localPath?: string;
  docClass?: string;
  docNo?: string;
  docSheet?: string;
  docRev?: string;
  source?: string;
  submittedBy?: string;
  submittedAt?: string;
}

interface ApprovalDocumentsJson {
  pending: PendingApprovalDocument[];
}

function getCandidatePaths(): string[] {
  // 1) project root (dev + typical prod)
  const fromCwd = path.join(process.cwd(), "data", "approvalDocuments.json");

  // 2) relative to this module (works for ts-node and lib/services output)
  // - ts-node: services/../../data -> data
  // - compiled: lib/services/../../data -> data (repo root)
  const fromRepoRootRelative = path.join(__dirname, "..", "..", "data", "approvalDocuments.json");

  // 3) if assets are copied into lib/data at build time
  const fromLib = path.join(__dirname, "..", "data", "approvalDocuments.json");

  return [fromCwd, fromRepoRootRelative, fromLib];
}

export async function getPendingApprovalDocuments(): Promise<PendingApprovalDocument[]> {
  const candidates = getCandidatePaths();
  let lastErr: unknown;

  for (const p of candidates) {
    try {
      const raw = await readFile(p, "utf8");
      const parsed = JSON.parse(raw) as ApprovalDocumentsJson;
      if (!parsed || !Array.isArray(parsed.pending)) return [];
      return parsed.pending;
    } catch (err) {
      lastErr = err;
    }
  }

  // If file missing / unreadable, return empty list (keep bot responsive)
  void lastErr;
  return [];
}

export async function getPendingApprovalDocumentById(
  id: string
): Promise<PendingApprovalDocument | null> {
  if (!id) return null;
  const docs = await getPendingApprovalDocuments();
  return docs.find((d) => d.id === id) ?? null;
}

function resolveLocalPath(p: string): string {
  if (path.isAbsolute(p)) return p;
  return path.join(process.cwd(), p);
}

export async function getDocumentText(doc: PendingApprovalDocument): Promise<string> {
  const localPath = doc.localPath?.trim();
  if (!localPath) return "";

  const candidates: string[] = [];
  if (path.isAbsolute(localPath)) {
    candidates.push(localPath);
  } else {
    // 1) relative to working directory
    candidates.push(resolveLocalPath(localPath));
    // 2) relative to repo root (ts-node + compiled lib/services)
    candidates.push(path.join(__dirname, "..", "..", localPath));
    // 3) relative to lib (if assets copied under lib/)
    candidates.push(path.join(__dirname, "..", localPath));
  }

  for (const p of candidates) {
    try {
      return await readFile(p, "utf8");
    } catch {
      // try next candidate
    }
  }
  return "";
}

