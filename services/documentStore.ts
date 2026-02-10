import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { extractTextFromFile } from "./documentTextExtractor";

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
  /** Workflow state */
  state?: "pendingApproval" | "approved" | "rejected" | "pendigApprovel";
  submittedBy?: string;
  submittedAt?: string;
}

interface ApprovalDocumentsJson {
  // new format
  documents?: PendingApprovalDocument[];
  // legacy format
  pending?: PendingApprovalDocument[];
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

async function tryReadJsonFile(p: string): Promise<ApprovalDocumentsJson | null> {
  try {
    const raw = await readFile(p, "utf8");
    return JSON.parse(raw) as ApprovalDocumentsJson;
  } catch {
    return null;
  }
}

async function resolveDocumentsFilePath(): Promise<string> {
  const candidates = getCandidatePaths();
  for (const p of candidates) {
    const parsed = await tryReadJsonFile(p);
    if (parsed) return p;
  }
  // default to cwd location if none exist
  return candidates[0];
}

export async function getAllDocuments(): Promise<PendingApprovalDocument[]> {
  const candidates = getCandidatePaths();
  let lastErr: unknown;

  for (const p of candidates) {
    try {
      const raw = await readFile(p, "utf8");
      const parsed = JSON.parse(raw) as ApprovalDocumentsJson;
      if (!parsed) return [];

      if (Array.isArray(parsed.documents)) return parsed.documents;

      // Legacy: "pending" array -> convert to documents with state pendingApproval
      if (Array.isArray(parsed.pending)) {
        return parsed.pending.map((d) => ({ ...d, state: d.state ?? "pendingApproval" }));
      }
      return [];
    } catch (err) {
      lastErr = err;
    }
  }

  // If file missing / unreadable, return empty list (keep bot responsive)
  void lastErr;
  return [];
}

export async function getPendingApprovalDocuments(): Promise<PendingApprovalDocument[]> {
  const docs = await getAllDocuments();
  return docs.filter((d) => {
    const s = String(d.state ?? "pendingApproval");
    return s === "pendingApproval" || s === "pendigApprovel";
  });
}

export async function getPendingApprovalDocumentById(
  id: string
): Promise<PendingApprovalDocument | null> {
  if (!id) return null;
  const docs = await getAllDocuments();
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
      return await extractTextFromFile(p);
    } catch {
      // try next candidate
    }
  }
  return "";
}

export async function setDocumentState(
  id: string,
  state: "pendingApproval" | "approved" | "rejected"
): Promise<boolean> {
  if (!id) return false;
  try {
    const filePath = await resolveDocumentsFilePath();
    const docs = await getAllDocuments();
    const idx = docs.findIndex((d) => d.id === id);
    if (idx < 0) return false;

    docs[idx] = { ...docs[idx], state };

    // Ensure directory exists
    await mkdir(path.dirname(filePath), { recursive: true }).catch(() => {});
    const out: ApprovalDocumentsJson = { documents: docs };
    await writeFile(filePath, JSON.stringify(out, null, 2) + "\n", "utf8");
    return true;
  } catch {
    return false;
  }
}

