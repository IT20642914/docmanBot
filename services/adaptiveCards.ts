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

export function buildInfoCard(input: { title: string; message?: string }): any {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: input.title, weight: "Bolder", size: "Large" },
      ...(input.message
        ? [
            {
              type: "TextBlock",
              text: input.message,
              wrap: true,
              spacing: "Small",
            },
          ]
        : []),
    ],
  };
}

export function buildNewDocumentNotificationCard(input: {
  title: string;
  documentNo?: string;
  documentClass?: string;
  documentRevision?: string;
  fileName?: string;
  docType?: string;
}): any {
  const metaParts = [
    input.documentClass?.trim(),
    input.documentNo?.trim(),
    input.documentRevision?.trim(),
  ].filter(Boolean);
  const meta = metaParts.length ? metaParts.join(" - ") : undefined;

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "New document to approve", weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: input.title, wrap: true, weight: "Bolder", spacing: "Small" },
      ...(meta ? [{ type: "TextBlock", text: meta, wrap: true, spacing: "None", isSubtle: true }] : []),
      ...(input.docType ? [{ type: "TextBlock", text: `Type: ${input.docType}`, wrap: true, spacing: "None", isSubtle: true }] : []),
      ...(input.fileName ? [{ type: "TextBlock", text: `File: ${input.fileName}`, wrap: true, spacing: "Small", isSubtle: true }] : []),
    ],
    actions: [
      { type: "Action.Submit", title: "View pending list", data: { action: "show_pending_approvals" } },
      { type: "Action.Submit", title: "Close", data: { action: "dismiss_pending_approvals" } },
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
  userLabel?: string;
}): any {
  const docs = input.docs ?? [];
  const count = docs.length;
  const user = input.userLabel?.trim();

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      ...(user
        ? [
            {
              type: "TextBlock",
              text: `Hi ${user}`,
              wrap: true,
              spacing: "None",
            },
          ]
        : []),
      { type: "TextBlock", text: "Documents pending approval", weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: `Pending: ${count}`, wrap: true, spacing: "Small" },
      ...(count === 0
        ? [{ type: "TextBlock", text: "No documents pending approval.", wrap: true, spacing: "Medium" }]
        : [{ type: "TextBlock", text: "Do you want to see the list?", wrap: true, spacing: "Medium" }]),
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
            d.DocumentClass && d.DocumentNo && d.DocumentSheet && d.DocumentRevision
              ? `${d.DocumentClass} - ${d.DocumentNo} - ${d.DocumentSheet} - ${d.DocumentRevision}`
              : undefined;
          const typeLine = d.docType ? `Type: ${d.docType}` : undefined;

          return {
            type: "Container",
            style: "emphasis",
            spacing: "Small",
            items: [
              { type: "TextBlock", text: d.Title || d.title || d.OriginalFileName || d.fileName || d.id, wrap: true, weight: "Bolder" },
              ...(typeLine
                ? [{ type: "TextBlock", text: typeLine, wrap: true, isSubtle: true, spacing: "None" }]
                : []),
              ...(meta ? [{ type: "TextBlock", text: meta, wrap: true, isSubtle: true, spacing: "None" }] : []),
              ...(d.ResponsiblePerson
                ? [{ type: "TextBlock", text: `Responsible: ${d.ResponsiblePerson}`, wrap: true, isSubtle: true, spacing: "None" }]
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
    d.DocumentClass && d.DocumentNo && d.DocumentSheet && d.DocumentRevision
      ? `${d.DocumentClass} - ${d.DocumentNo} - ${d.DocumentSheet} - ${d.DocumentRevision}`
      : null;

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Document details", weight: "Bolder", size: "Large" },
      kv("ID", d.id),
      kv("Title", d.Title || d.title || "-"),
      kv("Original file", d.OriginalFileName || d.fileName || "-"),
      ...(meta ? [kv("Metadata", meta)] : []),
      ...(d.DocumentStatus ? [kv("Doc status", d.DocumentStatus)] : []),
      ...(d.FileStatus ? [kv("File status", d.FileStatus)] : []),
      ...(d.Language ? [kv("Language", d.Language)] : []),
      ...(d.ResponsiblePerson ? [kv("Responsible", d.ResponsiblePerson)] : []),
      ...(d.Modified ? [kv("Modified", d.Modified)] : []),
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
    d.DocumentClass && d.DocumentNo && d.DocumentSheet && d.DocumentRevision
      ? `${d.DocumentClass} - ${d.DocumentNo} - ${d.DocumentSheet} - ${d.DocumentRevision}`
      : null;

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Document summary", weight: "Bolder", size: "Large" },
      kv("ID", d.id),
      kv("Title", d.Title || d.title || "-"),
      ...(d.docType ? [kv("Type", d.docType)] : []),
      ...(d.DocumentNo ? [kv("Document No", d.DocumentNo)] : []),
      ...(d.DocumentClass ? [kv("Document Class", d.DocumentClass)] : []),
      ...(d.DocumentRevision ? [kv("Document Rev", d.DocumentRevision)] : []),
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
      kv("Document", d.Title || d.title || d.id),
      ...(d.docType ? [kv("Type", d.docType)] : []),
      ...(d.DocumentNo ? [kv("Document No", d.DocumentNo)] : []),
      ...(d.DocumentClass ? [kv("Document Class", d.DocumentClass)] : []),
      ...(d.DocumentRevision ? [kv("Document Rev", d.DocumentRevision)] : []),
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

