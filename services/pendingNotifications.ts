import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import type { ApprovalDocument } from "./documentStore";

export interface NotificationTarget {
  /** Best identifier in Teams: activity.from.aadObjectId */
  aadObjectId?: string;
  /** Email/UPN if available */
  email?: string;
}

export interface PendingNotification {
  id: string;
  createdAt: string;
  target?: NotificationTarget;
  doc: Pick<
    ApprovalDocument,
    "id" | "Title" | "DocumentNo" | "DocumentClass" | "DocumentRevision" | "OriginalFileName" | "docType"
  >;
}

type NotificationsJson = { notifications: PendingNotification[] };

function filePath(): string {
  return path.join(process.cwd(), "data", "pendingNotifications.json");
}

function tryExtractEmail(input: unknown): string | null {
  const s = String(input ?? "");
  // Support placeholders like "samudra@skyforce" (no TLD)
  const m = s.match(/[^\s()<>]+@[^\s()<>]+/);
  return m?.[0]?.trim().toLowerCase() || null;
}

export async function enqueueNewDocumentNotification(
  doc: ApprovalDocument,
  target?: NotificationTarget
): Promise<void> {
  const p = filePath();
  let current: NotificationsJson = { notifications: [] };
  try {
    const raw = await readFile(p, "utf8");
    const parsed = JSON.parse(raw) as any;
    if (Array.isArray(parsed?.notifications)) current.notifications = parsed.notifications;
  } catch {
    // ignore
  }

  const derivedEmail =
    target?.email ||
    tryExtractEmail(doc.ResponsiblePerson) ||
    tryExtractEmail(doc.ModifiedBy) ||
    tryExtractEmail(doc.CreatedBy) ||
    tryExtractEmail(doc.OriginalCreator);

  const n: PendingNotification = {
    id: `N-${Date.now()}`,
    createdAt: new Date().toISOString(),
    target: {
      aadObjectId: target?.aadObjectId?.trim() || undefined,
      email: derivedEmail || undefined,
    },
    doc: {
      id: doc.id,
      Title: doc.Title,
      DocumentNo: doc.DocumentNo,
      DocumentClass: doc.DocumentClass,
      DocumentRevision: doc.DocumentRevision,
      OriginalFileName: doc.OriginalFileName,
      docType: doc.docType,
    },
  };
  current.notifications.unshift(n);
  current.notifications = current.notifications.slice(0, 25);

  await mkdir(path.dirname(p), { recursive: true }).catch(() => {});
  await writeFile(p, JSON.stringify(current, null, 2) + "\n", "utf8");
}

function matchesTarget(n: PendingNotification, user: NotificationTarget): boolean {
  const t = n.target;
  if (!t) return true; // broadcast if not targeted
  if (t.aadObjectId && user.aadObjectId) {
    return t.aadObjectId.toLowerCase() === user.aadObjectId.toLowerCase();
  }
  if (t.email && user.email) {
    return t.email.toLowerCase() === user.email.toLowerCase();
  }
  // if notification is targeted but we don't have match keys, don't deliver
  return false;
}

export async function drainNotificationsForUser(user: NotificationTarget): Promise<PendingNotification[]> {
  const p = filePath();
  try {
    const raw = await readFile(p, "utf8");
    const parsed = JSON.parse(raw) as any;
    const list: PendingNotification[] = Array.isArray(parsed?.notifications) ? parsed.notifications : [];
    const deliver = list.filter((n) => matchesTarget(n, user));
    const keep = list.filter((n) => !matchesTarget(n, user));
    await writeFile(p, JSON.stringify({ notifications: keep }, null, 2) + "\n", "utf8");
    return deliver;
  } catch {
    return [];
  }
}

