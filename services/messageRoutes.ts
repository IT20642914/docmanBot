import { stripMentionsText } from "@microsoft/teams.api";
import type { App } from "@microsoft/teams.apps";
import type { IStorage } from "@microsoft/teams.common";
import { identifyDocumentType } from "./documentIdentifier";
import { getFileNameFromActivity } from "./teamsActivityUtils";
import {
  buildInfoCard,
  buildLoadingCard,
  buildNewDocumentNotificationCard,
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
  getDocumentImages,
  getPendingApprovalDocumentById,
  getPendingApprovalDocuments,
  setDocumentState,
} from "./documentStore";
import { answerQuestionFromDocument, summarizeDocumentText } from "./azureOpenAi";
import { drainNotificationsForUser } from "./pendingNotifications";
import { getConversationRefByAadObjectId, upsertConversationRef } from "./conversationRegistry";
import { getEmailForAadObjectId } from "./graphService";

export function registerMessageRoutes(app: App, storage: IStorage<string, any>) {
  app.on("message", async (ctx: any) => {
    const activity = ctx?.activity;

    function getUserLabel(): string | null {
      const name =
        (typeof activity?.from?.name === "string" && activity.from.name.trim()) ||
        (typeof ctx?.userName === "string" && ctx.userName.trim()) ||
        null;
      // Email/UPN is not reliably available in normal message activities without Graph.
      const upn =
        (typeof activity?.from?.userPrincipalName === "string" && activity.from.userPrincipalName.trim()) ||
        (typeof activity?.channelData?.tenant?.userPrincipalName === "string" &&
          activity.channelData.tenant.userPrincipalName.trim()) ||
        null;
      if (name && upn && !name.includes(upn)) return `${name} (${upn})`;
      return name ?? upn;
    }

    function getUserIdentity(): { aadObjectId?: string; email?: string } {
      const aadObjectId =
        (typeof activity?.from?.aadObjectId === "string" && activity.from.aadObjectId.trim()) || undefined;
      const email =
        (typeof activity?.from?.userPrincipalName === "string" && activity.from.userPrincipalName.trim()) ||
        (typeof ctx?.userName === "string" && ctx.userName.includes("@") ? ctx.userName.trim() : "") ||
        undefined;
      return { aadObjectId, email: email ? email.toLowerCase() : undefined };
    }

    function isDebug(): boolean {
      return String(process.env.DOCUMATE_DEBUG || "").trim() === "1";
    }

    // Always store conversation reference (works for normal messages + card submits)
    const convId = activity?.conversation?.id;
    if (typeof convId === "string" && convId.trim()) {
      const ident = getUserIdentity();
      await upsertConversationRef({
        aadObjectId: ident.aadObjectId,
        email: ident.email,
        conversationId: convId.trim(),
        serviceUrl: activity?.serviceUrl,
        userLabel: getUserLabel() ?? undefined,
      });

      // If Teams didn't provide email, try Graph lookup (aadObjectId -> email) and store it.
      if (!ident.email && ident.aadObjectId) {
        void (async () => {
          try {
            const existing = await getConversationRefByAadObjectId(ident.aadObjectId!);
            if (existing?.email) return;
            const r = await getEmailForAadObjectId(ident.aadObjectId!);
            if (!r?.email) {
              if (isDebug()) console.warn("[graph] email lookup returned empty");
              return;
            }
            await upsertConversationRef({
              aadObjectId: ident.aadObjectId,
              email: r.email,
              conversationId: convId.trim(),
              serviceUrl: activity?.serviceUrl,
              userLabel: getUserLabel() ?? undefined,
            });
          } catch (e: any) {
            if (isDebug()) console.warn("[graph] email lookup failed", String(e?.message ?? e));
          }
        })();
      }
    } else if (isDebug()) {
      console.warn("[conversationRegistry] missing conversationId on activity");
    }

    async function sendTyping(): Promise<void> {
      // Teams shows a "typing" indicator; safe no-op if not supported.
      await ctx.send({ type: "typing" }).catch(() => {});
    }

    function startTypingLoop(): () => void {
      void sendTyping();
      const t = setInterval(() => {
        void sendTyping();
      }, 3000);
      return () => clearInterval(t);
    }

    function getSentActivityId(sent: any): string | null {
      return (
        (typeof sent?.id === "string" && sent.id) ||
        (typeof sent?.activityId === "string" && sent.activityId) ||
        null
      );
    }

    function getConversationId(): string | null {
      const id = activity?.conversation?.id;
      return typeof id === "string" && id.trim() ? id.trim() : null;
    }

    async function tryUpdateActivity(activityId: string, next: any): Promise<boolean> {
      const conversationId = getConversationId();
      if (!conversationId) return false;
      const updateFn = ctx?.api?.conversations?.activities?.(conversationId)?.update;
      if (typeof updateFn !== "function") return false;
      try {
        await updateFn(activityId, next);
        return true;
      } catch (e) {
        console.warn("activities.update failed", { activityId, e });
        return false;
      }
    }

    async function tryDeleteActivity(activityId: string): Promise<void> {
      const conversationId = getConversationId();
      if (!conversationId) return;
      const delFn = ctx?.api?.conversations?.activities?.(conversationId)?.delete;
      if (typeof delFn !== "function") return;
      try {
        await delFn(activityId);
      } catch (e) {
        console.warn("activities.delete failed", { activityId, e });
      }
    }

    async function replaceOrSendFinal(placeholderId: string | null, finalActivity: any): Promise<void> {
      // Prefer updating the placeholder message if available; otherwise just send.
      if (placeholderId) {
        const updated = await tryUpdateActivity(placeholderId, finalActivity);
        if (updated) return;
      }

      await ctx.send(finalActivity);

      // If we couldn't update, try deleting the placeholder to reduce clutter / disable buttons.
      if (placeholderId) await tryDeleteActivity(placeholderId);
    }

    async function runWithLoading(params: {
      targetId: string | null;
      loadingTitle: string;
      loadingMessage: string;
      work: () => Promise<any>;
    }): Promise<void> {
      const stopTyping = startTypingLoop();

      const loadingActivity = {
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildLoadingCard({ title: params.loadingTitle, message: params.loadingMessage }),
          },
        ],
      };

      let placeholderId: string | null = params.targetId;
      let didUpdate = false;
      if (params.targetId) {
        didUpdate = await tryUpdateActivity(params.targetId, loadingActivity);
      }

      if (!didUpdate) {
        // If we couldn't update the original card, try deleting it to avoid stale buttons.
        if (params.targetId) await tryDeleteActivity(params.targetId);
        const sent = await ctx.send(loadingActivity).catch(() => null);
        placeholderId = getSentActivityId(sent);
      }

      let finalActivity: any;
      try {
        finalActivity = await params.work();
      } catch (e: any) {
        finalActivity = {
          type: "message",
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: buildInfoCard({
                title: "Something went wrong",
                message: String(e?.message ?? e ?? "Unknown error"),
              }),
            },
          ],
        };
      } finally {
        stopTyping();
      }

      await replaceOrSendFinal(placeholderId, finalActivity);
    }

    // Handle Adaptive Card Action.Submit callbacks
    const submittedAction = activity?.value?.action;
    if (submittedAction === "show_pending_approvals") {
      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      await runWithLoading({
        targetId,
        loadingTitle: "Loading documents…",
        loadingMessage: "Fetching pending approvals list.",
        work: async () => {
          const docs = await getPendingApprovalDocuments();
          return {
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: buildPendingApprovalsListCard({ docs }),
              },
            ],
          };
        },
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

      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      await runWithLoading({
        targetId,
        loadingTitle: "Generating summary…",
        loadingMessage: `${doc.Title || doc.title || doc.OriginalFileName || doc.fileName || doc.id}`,
        work: async () => {
          const docText = await getDocumentText(doc);
          if (!docText) {
            return {
              type: "message",
              attachments: [
                {
                  contentType: "application/vnd.microsoft.card.adaptive",
                  content: buildInfoCard({
                    title: "Can’t read document",
                    message: "I couldn't read the local document content for this item.",
                  }),
                },
              ],
            };
          }
          const images = await getDocumentImages(doc);
          const summary = await summarizeDocumentText(docText, images).catch((e: any) => {
            return `Failed to summarize: ${String(e?.message ?? e)}`;
          });
          return {
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: buildDocSummaryAndQuestionCard({ doc, summary }),
              },
            ],
          };
        },
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
      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      await runWithLoading({
        targetId,
        loadingTitle: "Answering…",
        loadingMessage: "Reading the document and generating an answer.",
        work: async () => {
          const docText = await getDocumentText(doc);
          if (!docText) {
            return {
              type: "message",
              attachments: [
                {
                  contentType: "application/vnd.microsoft.card.adaptive",
                  content: buildInfoCard({
                    title: "Can’t read document",
                    message: "I couldn't read the local document content for this item.",
                  }),
                },
              ],
            };
          }
          const images = await getDocumentImages(doc);
          const answer = await answerQuestionFromDocument(docText, question, images).catch((e: any) => {
            return `Failed to answer: ${String(e?.message ?? e)}`;
          });
          return {
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: buildDocAnswerCard({ doc, question, answer }),
              },
            ],
          };
        },
      });
      return;
    }
    if (submittedAction === "approve_doc") {
      const docId = String(activity?.value?.docId ?? "").trim();
      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      const doc = await getPendingApprovalDocumentById(docId);
      await runWithLoading({
        targetId,
        loadingTitle: "Approving…",
        loadingMessage: `Updating ${docId}`,
        work: async () => {
          const ok = await setDocumentState(docId, "approved");
          if (!ok) {
            return {
              type: "message",
              attachments: [
                {
                  contentType: "application/vnd.microsoft.card.adaptive",
                  content: buildInfoCard({ title: "Couldn’t approve", message: "Failed to update document state." }),
                },
              ],
            };
          }
          return {
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: buildInfoCard({
                  title: "Approved",
                  message: `${doc?.Title || doc?.title || doc?.OriginalFileName || doc?.fileName || docId} was approved. I’ve updated the status.`,
                }),
              },
            ],
          };
        },
      });

      // Then show the refreshed pending list again
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
    if (submittedAction === "reject_doc") {
      const docId = String(activity?.value?.docId ?? "").trim();
      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      const doc = await getPendingApprovalDocumentById(docId);
      await runWithLoading({
        targetId,
        loadingTitle: "Rejecting…",
        loadingMessage: `Updating ${docId}`,
        work: async () => {
          const ok = await setDocumentState(docId, "rejected");
          if (!ok) {
            return {
              type: "message",
              attachments: [
                {
                  contentType: "application/vnd.microsoft.card.adaptive",
                  content: buildInfoCard({ title: "Couldn’t reject", message: "Failed to update document state." }),
                },
              ],
            };
          }
          return {
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: buildInfoCard({
                  title: "Rejected",
                  message: `${doc?.Title || doc?.title || doc?.OriginalFileName || doc?.fileName || docId} was rejected. I’ve updated the status.`,
                }),
              },
            ],
          };
        },
      });

      // Then show the refreshed pending list again
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
    if (submittedAction === "dismiss_pending_approvals") {
      const targetId = typeof activity?.replyToId === "string" ? activity.replyToId : null;
      const user = getUserLabel();
      await runWithLoading({
        targetId,
        loadingTitle: "Closing…",
        loadingMessage: "Got it.",
        work: async () => ({
          type: "message",
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: buildInfoCard({
                title: user ? `Thanks, ${user}` : "Thanks",
                message: 'If you need anything else, just say "hi".',
              }),
            },
          ],
        }),
      });
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
      const userLabel = getUserLabel() ?? undefined;
      await ctx.send({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildPendingApprovalsCard({ docs, userLabel }),
          },
        ],
      });

      // If any documents were added via API, notify the user.
      const notes = await drainNotificationsForUser(getUserIdentity());
      if (notes.length > 0) {
        const n = notes[0];
        const d = n.doc;
        await ctx.send({
          type: "message",
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: buildNewDocumentNotificationCard({
                title: d.Title || d.OriginalFileName || d.id,
                documentNo: d.DocumentNo,
                documentClass: d.DocumentClass,
                documentRevision: d.DocumentRevision,
                fileName: d.OriginalFileName,
                docType: d.docType,
              }),
            },
          ],
        });
      }
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

