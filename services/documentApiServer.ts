import http from "node:http";
import { addDocumentToApprovalList, type AddDocumentInput } from "./documentApi";
import { enqueueNewDocumentNotification } from "./pendingNotifications";
import { findConversationIdForTarget } from "./conversationRegistry";
import type { App } from "@microsoft/teams.apps";
import { buildNewDocumentNotificationCard } from "./adaptiveCards";

function tryExtractEmail(input: unknown): string | null {
  const s = String(input ?? "");
  // Support placeholders like "samudra@skyforce" (no TLD)
  const m = s.match(/[^\s()<>]+@[^\s()<>]+/);
  return m?.[0]?.trim().toLowerCase() || null;
}

function readJsonBody(req: http.IncomingMessage): Promise<any> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (c) => chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c)));
    req.on("end", () => {
      try {
        const raw = Buffer.concat(chunks).toString("utf8");
        resolve(raw ? JSON.parse(raw) : {});
      } catch (e) {
        reject(e);
      }
    });
    req.on("error", reject);
  });
}

export function startDocumentApiServer(app?: App) {
  const port = Number(process.env.DOC_API_PORT ?? 3981);

  const server = http.createServer(async (req, res) => {
    try {
      const url = req.url || "/";
      if (req.method === "POST" && url === "/api/documents") {
        const body = (await readJsonBody(req)) as Partial<AddDocumentInput>;
        if (!body?.localPath || !body?.Title) {
          res.writeHead(400, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ ok: false, error: "localPath and Title are required" }));
          return;
        }

        const doc = await addDocumentToApprovalList(body as AddDocumentInput);
        const targetEmail =
          (body.notifyEmail && body.notifyEmail.trim().toLowerCase()) ||
          tryExtractEmail(body.ResponsiblePerson) ||
          tryExtractEmail(doc.ResponsiblePerson) ||
          tryExtractEmail(body.ModifiedBy) ||
          tryExtractEmail(doc.ModifiedBy) ||
          tryExtractEmail(body.CreatedBy) ||
          tryExtractEmail(doc.CreatedBy) ||
          tryExtractEmail(body.OriginalCreator) ||
          tryExtractEmail(doc.OriginalCreator) ||
          undefined;

        await enqueueNewDocumentNotification(doc, {
          email: targetEmail,
          aadObjectId: body.notifyAadObjectId,
        });

        // Try to proactively notify immediately (no need for user to say hi)
        let notified = false;
        if (app) {
          const conversationId = await findConversationIdForTarget({
            aadObjectId: body.notifyAadObjectId,
            email: targetEmail,
          });
          if (conversationId) {
            try {
              await app.send(conversationId, {
                type: "message",
                attachments: [
                  {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: buildNewDocumentNotificationCard({
                      title: doc.Title || doc.OriginalFileName || doc.id,
                      documentNo: doc.DocumentNo,
                      documentClass: doc.DocumentClass,
                      documentRevision: doc.DocumentRevision,
                      fileName: doc.OriginalFileName,
                      docType: doc.docType,
                    }),
                  },
                ],
              });
              notified = true;
            } catch (e) {
              console.warn("proactive notify failed", e);
            }
          }
        }

        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true, doc, notified }));
        return;
      }

      if (req.method === "GET" && (url === "/" || url === "/healthz")) {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true, service: "document-api" }));
        return;
      }

      res.writeHead(404, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "not_found" }));
    } catch (e: any) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: String(e?.message ?? e) }));
    }
  });

  server.listen(port, () => {
    console.log(`[document-api] listening on ${port}`);
  });

  return server;
}

