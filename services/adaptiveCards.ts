import type { DocumentTypeResult } from "./documentIdentifier";
import type { PendingApprovalDocument } from "./documentStore";

function kv(label: string, value: string): any {
  return {
    type: "ColumnSet",
    columns: [
      { type: "Column", width: "auto", items: [{ type: "TextBlock", text: label, weight: "Bolder" }] },
      { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: value, wrap: true }] },
    ],
  };
}

export function buildLoadingCard(input: { title?: string; message?: string }): any {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: input.title ?? "Working…", weight: "Bolder", size: "Large" },
      {
        type: "TextBlock",
        text: input.message ?? "Please wait while I process your request.",
        wrap: true,
        spacing: "Small",
      },
    ],
  };
}

export function buildUserChangeCard(input: {
  title?: string;
  summary: string;
  userDisplayName?: string;
}): any {
  const title = input.title ?? "Change for user";
  const user = input.userDisplayName?.trim() || "Unknown user";

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: input.summary, wrap: true },
      { type: "FactSet", facts: [{ title: "User", value: user }] },
      {
        type: "TextBlock",
        text: "Choose an action:",
        weight: "Bolder",
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Acknowledge",
        data: { action: "ack_change" },
      },
      {
        type: "Action.Submit",
        title: "Request details",
        data: { action: "request_change_details" },
      },
    ],
  };
}

export function buildDocumentInfoCard(input: {
  fileName: string;
  documentType: DocumentTypeResult;
}): any {
  const { documentType } = input;
  const ifs = documentType.ifsMetadata;

  const body: any[] = [
    { type: "TextBlock", text: "Document detected", weight: "Bolder", size: "Large" },
    kv("File", input.fileName),
    kv("Type", documentType.type),
  ];

  if (ifs?.isIFSFormat) {
    body.push({ type: "TextBlock", text: "IFS metadata", weight: "Bolder", spacing: "Medium" });
    body.push(kv("Title", ifs.title || "Untitled"));
    body.push(kv("Class", ifs.docClass || "-"));
    body.push(kv("Doc No", ifs.docNo || "-"));
    body.push(kv("Sheet", ifs.docSheet || "-"));
    body.push(kv("Rev", ifs.docRev || "-"));
  } else {
    body.push({
      type: "TextBlock",
      text: "This filename doesn’t match the IFS naming format.",
      wrap: true,
      spacing: "Medium",
    });
  }

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body,
    actions: [
      {
        type: "Action.Submit",
        title: "Confirm",
        data: { action: "confirm_document", fileName: input.fileName, docType: documentType.type },
      },
      {
        type: "Action.Submit",
        title: "Not correct",
        data: { action: "reject_document", fileName: input.fileName, docType: documentType.type },
      },
    ],
  };
}

export function buildPendingApprovalsCard(input: {
  docs: PendingApprovalDocument[];
}): any {
  const docs = input.docs ?? [];
  const count = docs.length;

  const listItems =
    count === 0
      ? [{ type: "TextBlock", text: "No documents pending approval.", wrap: true }]
      : docs.slice(0, 10).map((d) => ({
          type: "TextBlock",
          text: `• ${d.title || d.fileName || d.id}`,
          wrap: true,
          spacing: "Small",
        }));

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Documents pending approval", weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: `Pending: ${count}`, wrap: true, spacing: "Small" },
      { type: "TextBlock", text: "Do you want to see the list?", wrap: true, spacing: "Medium" },
      { type: "Container", items: listItems, spacing: "Medium" },
    ],
    actions:
      count === 0
        ? [
            {
              type: "Action.Submit",
              title: "OK",
              data: { action: "dismiss_pending_approvals" },
            },
          ]
        : [
            {
              type: "Action.Submit",
              title: "Yes",
              data: { action: "show_pending_approvals" },
            },
            {
              type: "Action.Submit",
              title: "No",
              data: { action: "dismiss_pending_approvals" },
            },
          ],
  };
}

export function buildPendingApprovalsListCard(input: { docs: PendingApprovalDocument[] }): any {
  const docs = input.docs ?? [];

  const items =
    docs.length === 0
      ? [{ type: "TextBlock", text: "No documents pending approval.", wrap: true }]
      : docs.slice(0, 10).map((d) => {
          const meta =
            d.docClass && d.docNo && d.docSheet && d.docRev
              ? `${d.docClass} - ${d.docNo} - ${d.docSheet} - ${d.docRev}`
              : undefined;
          const typeLine = d.docType ? `Type: ${d.docType}` : undefined;

          return {
            type: "Container",
            style: "emphasis",
            spacing: "Small",
            items: [
              { type: "TextBlock", text: d.title || d.fileName || d.id, wrap: true, weight: "Bolder" },
              ...(typeLine
                ? [{ type: "TextBlock", text: typeLine, wrap: true, isSubtle: true, spacing: "None" }]
                : []),
              ...(meta ? [{ type: "TextBlock", text: meta, wrap: true, isSubtle: true, spacing: "None" }] : []),
              ...(d.submittedBy
                ? [{ type: "TextBlock", text: `Submitted by: ${d.submittedBy}`, wrap: true, isSubtle: true, spacing: "None" }]
                : []),
            ],
            selectAction: {
              type: "Action.Submit",
              title: "Select",
              data: { action: "select_pending_doc", docId: d.id },
            },
          };
        });

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Pending approvals", weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: "Select a document to view details.", wrap: true, spacing: "Small" },
      { type: "Container", items, spacing: "Medium" },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Refresh",
        data: { action: "show_pending_approvals" },
      },
      {
        type: "Action.Submit",
        title: "Close",
        data: { action: "dismiss_pending_approvals" },
      },
    ],
  };
}

export function buildPendingApprovalDetailsCard(input: { doc: PendingApprovalDocument }): any {
  const d = input.doc;
  const meta =
    d.docClass && d.docNo && d.docSheet && d.docRev ? `${d.docClass} - ${d.docNo} - ${d.docSheet} - ${d.docRev}` : null;

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Document details", weight: "Bolder", size: "Large" },
      kv("ID", d.id),
      kv("Title", d.title || "-"),
      kv("File", d.fileName || "-"),
      ...(meta ? [kv("Metadata", meta)] : []),
      ...(d.source ? [kv("Source", d.source)] : []),
      ...(d.submittedBy ? [kv("Submitted by", d.submittedBy)] : []),
      ...(d.submittedAt ? [kv("Submitted at", d.submittedAt)] : []),
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Approve",
        data: { action: "approve_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Reject",
        data: { action: "reject_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Back to list",
        data: { action: "show_pending_approvals" },
      },
    ],
  };
}

export function buildDocSummaryAndQuestionCard(input: {
  doc: PendingApprovalDocument;
  summary: string;
}): any {
  const d = input.doc;
  const meta =
    d.docClass && d.docNo && d.docSheet && d.docRev ? `${d.docClass} - ${d.docNo} - ${d.docSheet} - ${d.docRev}` : null;

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Document summary", weight: "Bolder", size: "Large" },
      kv("ID", d.id),
      kv("Title", d.title || "-"),
      ...(d.docType ? [kv("Type", d.docType)] : []),
      ...(meta ? [kv("Metadata", meta)] : []),
      { type: "TextBlock", text: input.summary || "(No summary)", wrap: true, spacing: "Medium" },
      { type: "TextBlock", text: "Ask a question about this document:", weight: "Bolder", spacing: "Medium" },
      {
        type: "Input.Text",
        id: "question",
        isMultiline: true,
        placeholder: "e.g. What are the approval requirements?",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Ask",
        data: { action: "ask_doc_question", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Approve",
        data: { action: "approve_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Reject",
        data: { action: "reject_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Back to list",
        data: { action: "show_pending_approvals" },
      },
    ],
  };
}

export function buildDocAnswerCard(input: {
  doc: PendingApprovalDocument;
  question: string;
  answer: string;
}): any {
  const d = input.doc;
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Answer from document", weight: "Bolder", size: "Large" },
      kv("Document", d.title || d.id),
      ...(d.docType ? [kv("Type", d.docType)] : []),
      { type: "TextBlock", text: `Q: ${input.question}`, wrap: true, weight: "Bolder", spacing: "Medium" },
      { type: "TextBlock", text: input.answer || "(No answer)", wrap: true, spacing: "Small" },
      { type: "TextBlock", text: "Ask another question:", weight: "Bolder", spacing: "Medium" },
      {
        type: "Input.Text",
        id: "question",
        isMultiline: true,
        placeholder: "Type your next question…",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Ask",
        data: { action: "ask_doc_question", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Approve",
        data: { action: "approve_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Reject",
        data: { action: "reject_doc", docId: d.id },
      },
      {
        type: "Action.Submit",
        title: "Back to list",
        data: { action: "show_pending_approvals" },
      },
    ],
  };
}

