export const FILE_DOWNLOAD_INFO_TYPE =
  "application/vnd.microsoft.teams.file.download.info";

export function decodeHtmlEntities(input: string): string {
  return input
    .replaceAll("&nbsp;", " ")
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replaceAll("&#x2F;", "/");
}

function tryExtractFileNameFromHtml(html: string): string | null {
  if (!html || typeof html !== "string") return null;

  // 1) Prefer anchor text: <a ...>filename</a>
  const anchorMatches = [...html.matchAll(/<a\b[^>]*>([^<]{1,300})<\/a>/gi)]
    .map((m) => decodeHtmlEntities(String(m[1] ?? "")).trim())
    .filter(Boolean);
  if (anchorMatches.length > 0) return anchorMatches[0];

  // 2) Try title="filename"
  const titleMatch = html.match(/\btitle\s*=\s*"([^"]{1,300})"/i);
  if (titleMatch?.[1]) return decodeHtmlEntities(titleMatch[1]).trim();

  // 3) Fallback: find something that looks like a filename with extension in visible text
  const textOnly = decodeHtmlEntities(html.replace(/<[^>]+>/g, " "))
    .replace(/\s+/g, " ")
    .trim();
  const fileLike = textOnly.match(/([^\s]{1,200}\.[a-z0-9]{1,8})/i);
  if (fileLike?.[1]) return fileLike[1].trim();

  return null;
}

export function isFileLikeAttachment(attachment: any): boolean {
  const ct = String(attachment?.contentType ?? "").toLowerCase();
  if (!ct) return false;
  if (ct === FILE_DOWNLOAD_INFO_TYPE) return true;
  if (ct.startsWith("application/vnd.microsoft.teams.file.")) return true;
  if (ct.includes("application/vnd.microsoft.teams.card.file")) return true;
  // Teams sometimes represents a file message as HTML content
  if (ct === "text/html") return true;
  return false;
}

export function getAttachmentFileName(attachment: any): string | null {
  if (!attachment) return null;
  const direct =
    (typeof attachment.name === "string" && attachment.name.trim()) ||
    (typeof attachment.filename === "string" && attachment.filename.trim());
  if (direct) return direct;

  const content = attachment.content;
  if (
    typeof content === "string" &&
    String(attachment?.contentType ?? "").toLowerCase() === "text/html"
  ) {
    const fromHtml = tryExtractFileNameFromHtml(content);
    if (fromHtml) return fromHtml;
  }
  if (content && typeof content === "object") {
    const fromContent =
      (typeof content.name === "string" && content.name.trim()) ||
      (typeof content.fileName === "string" && content.fileName.trim()) ||
      (typeof content.filename === "string" && content.filename.trim()) ||
      (typeof content.originalName === "string" && content.originalName.trim()) ||
      (typeof content?.item?.name === "string" && content.item.name.trim());
    if (fromContent) return fromContent;
  }

  return null;
}

function deepFindFirstStringByKeys(
  input: any,
  keys: string[],
  depth = 0,
  maxDepth = 4
): string | null {
  if (input == null) return null;
  if (depth > maxDepth) return null;

  if (typeof input === "object") {
    for (const k of keys) {
      const v = (input as any)[k];
      if (typeof v === "string" && v.trim()) return v.trim();
    }

    if (Array.isArray(input)) {
      for (const item of input) {
        const found = deepFindFirstStringByKeys(item, keys, depth + 1, maxDepth);
        if (found) return found;
      }
      return null;
    }

    for (const v of Object.values(input)) {
      const found = deepFindFirstStringByKeys(v, keys, depth + 1, maxDepth);
      if (found) return found;
    }
  }

  return null;
}

export function getFileNameFromActivity(activity: any): string | null {
  const attachments = activity?.attachments ?? [];
  for (const a of attachments) {
    const name = getAttachmentFileName(a);
    if (name) return name;
  }

  // Some Teams clients put file info in channelData/entities instead of attachments
  const keys = ["fileName", "filename", "name", "originalName", "title"];
  const fromChannelData = deepFindFirstStringByKeys(activity?.channelData, keys);
  if (fromChannelData && /\.[a-z0-9]{1,8}$/i.test(fromChannelData)) return fromChannelData;

  const fromEntities = deepFindFirstStringByKeys(activity?.entities, keys);
  if (fromEntities && /\.[a-z0-9]{1,8}$/i.test(fromEntities)) return fromEntities;

  return null;
}

function toBase64Url(input: string): string {
  return Buffer.from(input, "utf8")
    .toString("base64")
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replaceAll("=", "");
}

export function toGraphShareId(shareUrl: string): string {
  // Graph shareId format: "u!" + base64url(shareUrl)
  return `u!${toBase64Url(shareUrl)}`;
}

function extractUrlsFromText(text: string): string[] {
  if (!text || typeof text !== "string") return [];
  const decoded = decodeHtmlEntities(text);
  const matches = decoded.match(/https?:\/\/[^\s"'<>()]+/gi) ?? [];
  return matches
    .map((u) => u.replace(/[)\].,;]+$/g, ""))
    .filter(
      (u) =>
        // Ignore schema / mention URLs which are not file links
        !/^https?:\/\/schema\.skype\.com\//i.test(u) &&
        !/^https?:\/\/schema\.org\//i.test(u)
    );
}

function extractUrlsFromHtml(html: string): string[] {
  if (!html || typeof html !== "string") return [];
  const decoded = decodeHtmlEntities(html);
  const hrefs = [...decoded.matchAll(/href\s*=\s*"([^"]+)"/gi)]
    .map((m) => String(m[1] ?? "").trim())
    .filter((u) => /^https?:\/\//i.test(u));
  const plain = extractUrlsFromText(decoded);
  return Array.from(new Set([...hrefs, ...plain]));
}

function pickBestFileUrl(urls: string[]): string | null {
  if (!urls || urls.length === 0) return null;
  // Only treat these as "downloadable via Graph shareId" candidates
  const preferred = urls.find(
    (u) =>
      /^https:\/\//i.test(u) &&
      /(sharepoint\.com|1drv\.ms|onedrive\.live\.com|my\.sharepoint\.com)/i.test(u)
  );
  return preferred ?? null;
}

export function getShareUrlFromActivity(activity: any): string | null {
  const urls: string[] = [];

  // From plain text (Teams sometimes includes the link in text)
  if (typeof activity?.text === "string") {
    urls.push(...extractUrlsFromText(activity.text));
  }

  // From HTML attachments (common for Word add-in / share cards)
  const attachments = activity?.attachments ?? [];
  for (const a of attachments) {
    if (typeof a?.content === "string") {
      urls.push(...extractUrlsFromHtml(a.content));
    }
  }

  return pickBestFileUrl(Array.from(new Set(urls)));
}

export function buildIncomingActivityLog(activity: any, text: unknown) {
  const attachments = activity?.attachments ?? [];
  const attachmentSummaries = attachments.map((a: any, i: number) => ({
    i,
    contentType: a?.contentType,
    name: a?.name,
    contentTypeOfContent: typeof a?.content,
    contentLength:
      typeof a?.content === "string"
        ? a.content.length
        : a?.content && typeof a.content === "object"
          ? JSON.stringify(a.content).length
          : undefined,
    contentKeys: a?.content && typeof a.content === "object" ? Object.keys(a.content) : undefined,
    contentSnippet: typeof a?.content === "string" ? a.content.slice(0, 140) : undefined,
  }));

  const channelDataKeys =
    activity?.channelData && typeof activity.channelData === "object"
      ? Object.keys(activity.channelData)
      : undefined;

  const entitiesSummary = Array.isArray(activity?.entities)
    ? activity.entities.map((e: any) => ({
        type: e?.type,
        keys: e && typeof e === "object" ? Object.keys(e).slice(0, 12) : undefined,
      }))
    : undefined;

  return {
    type: activity?.type,
    name: activity?.name,
    channelId: activity?.channelId,
    conversationId: activity?.conversation?.id,
    fromId: activity?.from?.id,
    text,
    attachmentsCount: attachments.length,
    attachmentSummaries,
    channelDataKeys,
    entitiesSummary,
  };
}

