import { stripMentionsText } from "@microsoft/teams.api";
import type { App } from "@microsoft/teams.apps";
import type { IStorage } from "@microsoft/teams.common";
import { identifyDocumentType } from "./documentIdentifier";
import { getFileNameFromActivity } from "./teamsActivityUtils";
import {
  buildDocAnswerCard,
  buildDocSummaryAndQuestionCard,
  buildDocumentInfoCard,
  buildPendingApprovalDetailsCard,
  buildPendingApprovalsCard,
  buildPendingApprovalsListCard,
  buildUserChangeCard,
} from "./adaptiveCards";
import {
  getDocumentText,
  getPendingApprovalDocumentById,
  getPendingApprovalDocuments,
} from "./documentStore";
import { answerQuestionFromDocument, summarizeDocumentText } from "./azureOpenAi";

export function registerMessageRoutes(app: App, storage: IStorage<string, any>) {
  app.on("message", async (ctx: any) => {
    const activity = ctx?.activity;

    // Handle Adaptive Card Action.Submit callbacks
    const submittedAction = activity?.value?.action;
    if (submittedAction === "show_pending_approvals") {
      const docs = await getPendingApprovalDocuments();
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildPendingApprovalsListCard({ docs }),
          },
        ],
      });
      return;
    }
    if (submittedAction === "select_pending_doc") {
      const docId = String(activity?.value?.docId ?? "").trim();
      const doc = await getPendingApprovalDocumentById(docId);
      if (!doc) {
        await ctx.send("I couldn’t find that document. Please refresh the list.");
        return;
      }

      const docText = await getDocumentText(doc);
      if (!docText) {
        await ctx.send("I couldn't read the local document text for this item.");
        return;
      }
      const summary = await summarizeDocumentText(docText).catch((e: any) => {
        return `Failed to summarize: ${String(e?.message ?? e)}`;
      });
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildDocSummaryAndQuestionCard({ doc, summary }),
          },
        ],
      });
      return;
    }
    if (submittedAction === "ask_doc_question") {
      const docId = String(activity?.value?.docId ?? "").trim();
      const question = String(activity?.value?.question ?? "").trim();
      if (!question) {
        await ctx.send("Please type a question first.");
        return;
      }

      const doc = await getPendingApprovalDocumentById(docId);
      if (!doc) {
        await ctx.send("I couldn’t find that document. Please refresh the list.");
        return;
      }
      const docText = await getDocumentText(doc);
      if (!docText) {
        await ctx.send("I couldn't read the local document text for this item.");
        return;
      }

      const answer = await answerQuestionFromDocument(docText, question).catch((e: any) => {
        return `Failed to answer: ${String(e?.message ?? e)}`;
      });

      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildDocAnswerCard({ doc, question, answer }),
          },
        ],
      });
      return;
    }
    if (submittedAction === "approve_doc") {
      const docId = String(activity?.value?.docId ?? "").trim();
      await ctx.send(`Approved: ${docId}`);
      return;
    }
    if (submittedAction === "reject_doc") {
      const docId = String(activity?.value?.docId ?? "").trim();
      await ctx.send(`Rejected: ${docId}`);
      return;
    }
    if (submittedAction === "dismiss_pending_approvals") {
      await ctx.send("Ok.");
      return;
    }
    if (submittedAction === "ack_change") {
      await ctx.send("Acknowledged. I’ll track this change.");
      return;
    }
    if (submittedAction === "request_change_details") {
      await ctx.send("Tell me what you want to change (metadata, rename, revision, etc.).");
      return;
    }
    if (submittedAction === "confirm_document") {
      await ctx.send("Confirmed. I’ll proceed with this document.");
      return;
    }
    if (submittedAction === "reject_document") {
      await ctx.send("Ok — please share the correct document name or upload the correct file.");
      return;
    }

    const text = String(stripMentionsText(activity) ?? "").trim();
    const lower = text.toLowerCase();

    // Default / greeting: show pending-approvals Adaptive Card
    const conversationId = activity?.conversation?.id;
    const welcomeKey = conversationId ? `welcomeShown:${conversationId}` : null;
    const alreadyWelcomed = welcomeKey ? Boolean(await storage.get(welcomeKey)) : false;
    const isGreeting =
      lower === "hi" ||
      lower === "hello" ||
      lower === "hey" ||
      lower === "hi!" ||
      lower === "hello!" ||
      lower === "hey!";

    if (!alreadyWelcomed || isGreeting) {
      if (welcomeKey && !alreadyWelcomed) {
        await storage.set(welcomeKey, true);
      }
      const docs = await getPendingApprovalDocuments();
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildPendingApprovalsCard({ docs }),
          },
        ],
      });
      return;
    }

    // If user sent a file, show document info as an Adaptive Card
    const fileName = getFileNameFromActivity(activity);
    if (fileName) {
      const docType = identifyDocumentType(fileName);
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildDocumentInfoCard({ fileName, documentType: docType }),
          },
        ],
      });
      return;
    }

    // Manual trigger to show a "change for user" Adaptive Card
    if (lower === "change" || lower.startsWith("change ")) {
      const userDisplayName = activity?.from?.name || activity?.from?.aadObjectId || undefined;
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildUserChangeCard({
              summary: text.length > "change".length ? text : "Please review the requested change.",
              userDisplayName,
            }),
          },
        ],
      });
      return;
    }

    await ctx.send(`you said: ${text}`);
  });
}

