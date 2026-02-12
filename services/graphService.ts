import { ClientSecretCredential, ManagedIdentityCredential } from "@azure/identity";

type GraphUser = {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
};

function getCredential(): ClientSecretCredential | ManagedIdentityCredential | null {
  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID;
  const clientSecret = process.env.CLIENT_SECRET || process.env.CLIENT_PASSWORD;

  // Local dev / client secret auth
  if (clientId && tenantId && clientSecret) {
    return new ClientSecretCredential(tenantId, clientId, clientSecret);
  }

  // Managed identity (UMI/System). This will only work when running in Azure.
  if (clientId) {
    return new ManagedIdentityCredential({ clientId });
  }

  return null;
}

function isGraphDebug(): boolean {
  return String(process.env.DOCUMATE_GRAPH_DEBUG || process.env.DOCUMATE_DEBUG || "").trim() === "1";
}

async function getGraphToken(): Promise<string | null> {
  try {
    const cred = getCredential();
    if (!cred) {
      if (isGraphDebug()) console.warn("[graph] no credential available (missing env vars?)");
      return null;
    }
    const token = await cred.getToken("https://graph.microsoft.com/.default");
    if (!token?.token && isGraphDebug()) console.warn("[graph] token empty");
    return token?.token || null;
  } catch (e: any) {
    if (isGraphDebug()) console.warn("[graph] token fetch failed", String(e?.message ?? e));
    return null;
  }
}

function pickEmail(u: GraphUser | null): string | null {
  const mail = typeof u?.mail === "string" ? u.mail.trim() : "";
  if (mail) return mail.toLowerCase();
  const upn = typeof u?.userPrincipalName === "string" ? u.userPrincipalName.trim() : "";
  if (upn) return upn.toLowerCase();
  return null;
}

export async function getEmailForAadObjectId(
  aadObjectId: string
): Promise<{ email: string } | null> {
  const id = String(aadObjectId || "").trim();
  if (!id) return null;

  const token = await getGraphToken();
  if (!token) return null;

  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
    id
  )}?$select=id,displayName,mail,userPrincipalName`;

  // Node 18+ has global fetch; if not present, skip gracefully.
  const f = (globalThis as any)?.fetch;
  if (typeof f !== "function") {
    if (isGraphDebug()) console.warn("[graph] fetch not available in this Node runtime");
    return null;
  }

  const res = await f(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!res.ok) {
    if (isGraphDebug()) {
      let body = "";
      try {
        body = await res.text();
      } catch {
        body = "";
      }
      console.warn("[graph] /users/{id} failed", { status: res.status, body: body?.slice(0, 500) });
    }
    return null;
  }

  const data = (await res.json()) as GraphUser;
  const email = pickEmail(data);
  if (!email) return null;
  return { email };
}

