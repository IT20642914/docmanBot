import { stripMentionsText } from "@microsoft/teams.api";
import type { App } from "@microsoft/teams.apps";
import type { IStorage } from "@microsoft/teams.common";
import { identifyDocumentType } from "./documentIdentifier";
import {
  buildIncomingActivityLog,
  getAttachmentFileName,
  getFileNameFromActivity,
  getShareUrlFromActivity,
  isFileLikeAttachment,
  toGraphShareId,
} from "./teamsActivityUtils";

interface ConversationState {
  count: number;
}

async function getConversationState(
  storage: IStorage<string, any>,
  conversationId: string
): Promise<ConversationState> {
  let state = await storage.get(conversationId);
  if (!state) {
    state = { count: 0 };
    await storage.set(conversationId, state);
  }
  return state as ConversationState;
}

async function handleShareLinkViaGraph(context: any): Promise<boolean> {
  const rawActivity = context.activity as any;
  const shareUrl = getShareUrlFromActivity(rawActivity);
  if (!shareUrl) return false;

  try {
    const shareId = toGraphShareId(shareUrl);

    // Fetch metadata (name/size/webUrl)
    const driveItem = await context.appGraph.call(() => ({
      method: "get",
      path: `/shares/${shareId}/driveItem?$select=name,size,webUrl`,
    }));

    const fileName: string = (driveItem as any)?.name ?? "shared-file";

    // Download bytes
    const arrayBuffer = await context.appGraph.call(
      () => ({
        method: "get",
        path: `/shares/${shareId}/driveItem/content`,
      }),
      { requestConfig: { responseType: "arraybuffer" } }
    );

    const buffer = Buffer.from(arrayBuffer as any);

    // Optional: forward to backend
    const backendUrl = process.env.BACKEND_INGEST_URL;
    if (backendUrl) {
      const form = new FormData();
      form.append("file", new Blob([buffer]), fileName);
      form.append("source", "teams-graph-share");
      form.append("shareUrl", shareUrl);

      const res = await fetch(backendUrl, { method: "POST", body: form as any });
      if (!res.ok) {
        throw new Error(`Backend upload failed: ${res.status} ${res.statusText}`);
      }
    }

    await context.send(
      `📄 **${fileName}**\n` +
        `Downloaded from Graph (${buffer.length} bytes).` +
        (process.env.BACKEND_INGEST_URL ? ` Sent to backend.` : "")
    );
    return true;
  } catch (err: any) {
    console.error("[docmanBot] Graph share download failed", {
      message: err?.message,
      shareUrl,
    });
    await context.send(
      `I found a shared file link, but couldn't download it via Graph.\n` +
        `Most common reasons:\n` +
        `- Missing Graph permissions (Files/Sites)\n` +
        `- Admin consent not granted\n` +
        `- Link not accessible to the app\n`
    );
    return true; // handled (we replied)
  }
}

async function handleFileAttachments(context: any): Promise<boolean> {
  const activity = context.activity as any;
  const attachments = activity.attachments ?? [];

  const fileLikeAttachments = attachments.filter((a: any) => isFileLikeAttachment(a));
  const namedAttachments = attachments.filter((a: any) => !!getAttachmentFileName(a));
  const candidateFileAttachments =
    fileLikeAttachments.length > 0 ? fileLikeAttachments : namedAttachments;

  if (candidateFileAttachments.length === 0) return false;

  // Prefer filenames extracted from the attachments themselves (avoids duplicates like:
  // file.download.info + extra text/html attachment in the same message).
  const attachmentFileNames = candidateFileAttachments
    .map((a: any) => getAttachmentFileName(a))
    .filter((n: string | null): n is string => !!n && n.trim().length > 0);

  const uniqueFileNames: string[] =
    attachmentFileNames.length > 0
      ? Array.from(new Set(attachmentFileNames))
      : (() => {
          const fallback = getFileNameFromActivity(activity);
          return fallback ? [fallback] : [];
        })();

  if (uniqueFileNames.length === 0) {
    await context.send(
      `📄 I received an attachment, but couldn't extract the filename (contentTypes: ${candidateFileAttachments
        .map((a: any) => String(a?.contentType ?? "unknown"))
        .join(", ")}).`
    );
    return true;
  }

  const messages: string[] = [];
  for (const fileName of uniqueFileNames) {
    const result = identifyDocumentType(fileName);
    messages.push(`📄 **${fileName}**\n\n${result.message}`);
  }

  const text = messages.join("\n\n---\n\n");
  await context.send({
    type: "message",
    text,
    textFormat: "markdown",
  });
  return true;
}

async function handleCommands(context: any, storage: IStorage<string, any>): Promise<boolean> {
  const activity = context.activity as any;
  const text = stripMentionsText(activity);

  if (text === "/reset") {
    await storage.delete(activity.conversation.id);
    await context.send("Ok I've deleted the current conversation state.");
    return true;
  }

  if (text === "/count") {
    const state = await getConversationState(storage, activity.conversation.id);
    await context.send(`The count is ${state.count}`);
    return true;
  }

  if (text === "/diag") {
    await context.send(JSON.stringify(activity));
    return true;
  }

  if (text === "/state") {
    const state = await getConversationState(storage, activity.conversation.id);
    await context.send(JSON.stringify(state));
    return true;
  }

  if (text === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0", // Microsoft Teams SDK
    };
    await context.send(JSON.stringify(runtime));
    return true;
  }

  return false;
}

export function registerMessageRoutes(app: App, storage: IStorage<string, any>) {
  // 1) Logging middleware (runs first)
  app.use(async (ctx: any) => {
    if (ctx.activity?.type !== "message") return ctx.next();

    const text = stripMentionsText(ctx.activity);
    console.log("[docmanBot] incoming", buildIncomingActivityLog(ctx.activity, text));
    if (process.env.DEBUG_ACTIVITY === "true") {
      console.log("[docmanBot] activity payload", JSON.stringify(ctx.activity, null, 2));
    }

    return ctx.next();
  });

  // 2) Share-link route (Graph download)
  app.use(async (ctx: any) => {
    if (ctx.activity?.type !== "message") return ctx.next();
    const handled = await handleShareLinkViaGraph(ctx);
    if (handled) return;
    return ctx.next();
  });

  // 3) File-attachment route (filename based classification)
  app.use(async (ctx: any) => {
    if (ctx.activity?.type !== "message") return ctx.next();
    const handled = await handleFileAttachments(ctx);
    if (handled) return;
    return ctx.next();
  });

  // 4) Command route (/reset, /count, etc.)
  app.use(async (ctx: any) => {
    if (ctx.activity?.type !== "message") return ctx.next();
    const handled = await handleCommands(ctx, storage);
    if (handled) return;
    return ctx.next();
  });

  // 5) Default message route (echo)
  app.on("message", async (ctx: any) => {
    const activity = ctx.activity as any;
    const text = stripMentionsText(activity);
    const state = await getConversationState(storage, activity.conversation.id);
    state.count++;
    await ctx.send(`[${state.count}] you said: ${text}`);
  });
}

