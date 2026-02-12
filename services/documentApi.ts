import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import type { ApprovalDocument } from "./documentStore";

type ApprovalDocumentsJson = { documents: ApprovalDocument[] };

function inferDocTypeFromName(name: string): string | undefined {
  const n = String(name || "").toLowerCase();
  if (n.endsWith(".pdf")) return "PDF";
  if (n.endsWith(".docx")) return "DOCX";
  if (n.endsWith(".xlsx")) return "XLSX";
  if (n.endsWith(".xlsb")) return "XLSB";
  if (n.endsWith(".xls")) return "XLS";
  if (n.endsWith(".pptx")) return "PPTX";
  if (n.endsWith(".ppt")) return "PPT";
  if (n.endsWith(".txt")) return "TXT";
  if (n.endsWith(".md")) return "MD";
  return undefined;
}

function getDocumentsFilePath(): string {
  return path.join(process.cwd(), "data", "approvalDocuments.json");
}

async function readDocumentsFile(): Promise<ApprovalDocumentsJson> {
  const p = getDocumentsFilePath();
  try {
    const raw = await readFile(p, "utf8");
    const parsed = JSON.parse(raw) as any;
    const docs = Array.isArray(parsed?.documents) ? (parsed.documents as ApprovalDocument[]) : [];
    return { documents: docs };
  } catch {
    return { documents: [] };
  }
}

async function writeDocumentsFile(data: ApprovalDocumentsJson): Promise<void> {
  const p = getDocumentsFilePath();
  await mkdir(path.dirname(p), { recursive: true }).catch(() => {});
  await writeFile(p, JSON.stringify(data, null, 2) + "\n", "utf8");
}

function nextId(existing: ApprovalDocument[]): string {
  const nums = existing
    .map((d) => String(d.id || "").match(/^DOC-(\d+)$/i)?.[1])
    .filter(Boolean)
    .map((s) => Number(s))
    .filter((n) => Number.isFinite(n));
  const max = nums.length ? Math.max(...nums) : 0;
  return `DOC-${String(max + 1).padStart(3, "0")}`;
}

export interface AddDocumentInput {
  localPath: string;
  docType?: string;
  state?: "pendingApproval";
  /** Optional: who to notify */
  notifyEmail?: string;
  /** Optional: Teams AAD object id of approver */
  notifyAadObjectId?: string;
  Title: string;
  DocumentNo?: string;
  DocumentClass?: string;
  DocumentRevision?: string;
  DocumentSheet?: string;
  OriginalFileType?: string;
  DocumentStatus?: string;
  FileStatus?: string;
  Language?: string;
  ResponsiblePerson?: string;
  ModifiedBy?: string;
  CreatedBy?: string;
  OriginalCreator?: string;
  DateCreated?: string;
  Modified?: string;
  CheckedOutBy?: string;
  DocumentType?: string;
  OriginalFileName?: string;
  Format?: string;
}

export async function addDocumentToApprovalList(input: AddDocumentInput): Promise<ApprovalDocument> {
  const store = await readDocumentsFile();
  const id = nextId(store.documents);

  const fileName = input.OriginalFileName || path.basename(input.localPath);
  const docType = input.docType || inferDocTypeFromName(fileName) || inferDocTypeFromName(input.localPath);

  const nowIso = new Date().toISOString();

  const doc: ApprovalDocument = {
    id,
    state: "pendingApproval",
    localPath: input.localPath,
    docType,
    Title: input.Title,
    DocumentNo: input.DocumentNo ?? "",
    DocumentClass: input.DocumentClass ?? "",
    Format: input.Format ?? "*",
    DocumentSheet: input.DocumentSheet ?? "1",
    DocumentRevision: input.DocumentRevision ?? "",
    OriginalFileType: input.OriginalFileType ?? docType ?? "",
    DocumentStatus: input.DocumentStatus ?? "Preliminary",
    FileStatus: input.FileStatus ?? "Checked In",
    Language: input.Language ?? "en",
    ResponsiblePerson: input.ResponsiblePerson ?? "",
    ModifiedBy: input.ModifiedBy ?? "",
    CreatedBy: input.CreatedBy ?? "",
    OriginalCreator: input.OriginalCreator ?? "",
    DateCreated: input.DateCreated ?? nowIso,
    Modified: input.Modified ?? nowIso,
    CheckedOutBy: input.CheckedOutBy ?? "",
    DocumentType: input.DocumentType ?? "ORIGINAL",
    OriginalFileName: fileName,
  };

  store.documents.push(doc);
  await writeDocumentsFile(store);
  return doc;
}

