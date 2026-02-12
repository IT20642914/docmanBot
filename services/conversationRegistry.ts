import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

export interface ConversationRef {
  conversationId: string;
  serviceUrl?: string;
  updatedAt: string;
  userLabel?: string;
  /** Email/UPN if we were able to resolve it (e.g., via Graph) */
  email?: string;
}

type RegistryJson = {
  byAadObjectId: Record<string, ConversationRef>;
  byEmail: Record<string, ConversationRef>;
};

function filePath(): string {
  // Prefer a stable location relative to this module rather than process.cwd(),
  // because dev/prod startup paths may differ (e.g., running from a parent folder).
  return path.join(__dirname, "..", "..", "data", "conversationRegistry.json");
}

async function readRegistry(): Promise<RegistryJson> {
  const p = filePath();
  try {
    const raw = await readFile(p, "utf8");
    const parsed = JSON.parse(raw) as any;
    return {
      byAadObjectId: parsed?.byAadObjectId ?? {},
      byEmail: parsed?.byEmail ?? {},
    };
  } catch {
    return { byAadObjectId: {}, byEmail: {} };
  }
}

async function writeRegistry(r: RegistryJson): Promise<void> {
  const p = filePath();
  await mkdir(path.dirname(p), { recursive: true }).catch(() => {});
  await writeFile(p, JSON.stringify(r, null, 2) + "\n", "utf8");
}

function isDebug(): boolean {
  return String(process.env.DOCUMATE_DEBUG || "").trim() === "1";
}

export async function upsertConversationRef(input: {
  aadObjectId?: string;
  email?: string;
  conversationId: string;
  serviceUrl?: string;
  userLabel?: string;
}): Promise<void> {
  const p = filePath();
  try {
    const r = await readRegistry();
    const ref: ConversationRef = {
      conversationId: input.conversationId,
      serviceUrl: input.serviceUrl,
      updatedAt: new Date().toISOString(),
      userLabel: input.userLabel,
      email: input.email ? input.email.toLowerCase() : undefined,
    };

    if (input.aadObjectId) r.byAadObjectId[input.aadObjectId.toLowerCase()] = ref;
    if (ref.email) r.byEmail[ref.email.toLowerCase()] = ref;

    await writeRegistry(r);
    if (isDebug()) {
      console.log("[conversationRegistry] upsert ok", {
        path: p,
        aadObjectId: input.aadObjectId,
        email: ref.email,
      });
    }
  } catch (e: any) {
    console.warn("[conversationRegistry] upsert failed", { path: p, error: String(e?.message ?? e) });
  }
}

export async function getConversationRefByAadObjectId(aadObjectId: string): Promise<ConversationRef | null> {
  const id = String(aadObjectId || "").trim().toLowerCase();
  if (!id) return null;
  const r = await readRegistry();
  return r.byAadObjectId[id] ?? null;
}

export async function findConversationIdForTarget(target: {
  aadObjectId?: string;
  email?: string;
}): Promise<string | null> {
  const r = await readRegistry();
  if (target.aadObjectId) {
    const v = r.byAadObjectId[target.aadObjectId.toLowerCase()];
    if (v?.conversationId) return v.conversationId;
  }
  if (target.email) {
    const v = r.byEmail[target.email.toLowerCase()];
    if (v?.conversationId) return v.conversationId;
  }
  return null;
}

