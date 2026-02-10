import { loadEnvFromLocalFilesIfMissing } from "./dotenvLite";

export interface AzureOpenAiConfig {
  /** Azure OpenAI resource endpoint, e.g. https://<resource>.cognitiveservices.azure.com/ */
  endpoint: string;
  apiKey: string;
  /** Azure OpenAI deployment name (the "Name" column in Deployments) */
  deployment: string;
  /** API version for Azure OpenAI, e.g. 2024-04-01-preview */
  apiVersion: string;
  /** Optional model name passed to SDK call (some examples require it) */
  model: string;
}

export function getAzureOpenAiConfigFromEnv(): AzureOpenAiConfig | null {
  const endpoint = String(process.env.AZURE_OPENAI_ENDPOINT ?? "https://ai-anis5304ai744397091529.cognitiveservices.azure.com/").trim();
  const apiKey = String(process.env.AZURE_OPENAI_API_KEY ?? "BuEaamvBcxruTHQCwjM3yaA4waFvvZdKiyVrfmVsh5kIWM9wdtGHJQQJ99AKACfhMk5XJ3w3AAAAACOG3dhV").trim();
  const deployment = String(process.env.AZURE_OPENAI_DEPLOYMENT ?? "gpt-5.2-chat-2").trim();
  const apiVersion = String(process.env.AZURE_OPENAI_API_VERSION ?? "2024-04-01-preview").trim();
  const model = String(process.env.AZURE_OPENAI_MODEL ?? "gpt-5.2-chat").trim() || deployment;
  if (!endpoint || !apiKey || !deployment) return null;
  return { endpoint: endpoint.replace(/\/+$/, ""), apiKey, deployment, apiVersion, model };
}

function clampText(input: string, maxChars: number): string {
  const s = String(input ?? "");
  if (s.length <= maxChars) return s;
  return `${s.slice(0, maxChars)}\n\n[TRUNCATED]`;
}

let _clientPromise: Promise<any> | null = null;
async function getSdkClient(cfg: AzureOpenAiConfig) {
  if (_clientPromise) return _clientPromise;

  _clientPromise = (async () => {
    // `openai` is ESM-first; use dynamic import to work in this TS/Node setup.
    const mod: any = await import("openai");
    const AzureOpenAI = mod.AzureOpenAI;
    if (!AzureOpenAI) throw new Error("AzureOpenAI SDK not found (openai package).");

    return new AzureOpenAI({
      endpoint: cfg.endpoint,
      apiKey: cfg.apiKey,
      deployment: cfg.deployment,
      apiVersion: cfg.apiVersion,
    });
  })();

  return _clientPromise;
}

async function chatComplete(cfg: AzureOpenAiConfig, messages: Array<{ role: string; content: string }>, maxTokens: number) {
  const client = await getSdkClient(cfg);
  const resp = await client.chat.completions.create({
    messages,
    model: cfg.model,
    max_completion_tokens: maxTokens,
    // Some Azure OpenAI deployments only support the default temperature (1).
    // Omitting it keeps compatibility across models.
  });
  const content = resp?.choices?.[0]?.message?.content;
  if (typeof content !== "string" || !content.trim()) return "";
  return content.trim();
}

export async function summarizeDocumentText(docText: string): Promise<string> {
  await loadEnvFromLocalFilesIfMissing([
    "AZURE_OPENAI_ENDPOINT",
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_DEPLOYMENT",
    "AZURE_OPENAI_API_VERSION",
    "AZURE_OPENAI_MODEL",
  ]);
  const cfg = getAzureOpenAiConfigFromEnv();
  if (!cfg) {
    return "Azure OpenAI is not configured. Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY (AZURE_OPENAI_DEPLOYMENT is optional).";
  }

  const safeText = clampText(docText, 12000);
  return await chatComplete(
    cfg,
    [
      {
        role: "system",
        content:
          "You summarize engineering documents for approval. Output concise bullet points, then a short 'Approval checklist' section.",
      },
      { role: "user", content: `Summarize this document:\n\n${safeText}` },
    ],
    600
  );
}

export async function answerQuestionFromDocument(docText: string, question: string): Promise<string> {
  await loadEnvFromLocalFilesIfMissing([
    "AZURE_OPENAI_ENDPOINT",
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_DEPLOYMENT",
    "AZURE_OPENAI_API_VERSION",
    "AZURE_OPENAI_MODEL",
  ]);
  const cfg = getAzureOpenAiConfigFromEnv();
  if (!cfg) {
    return "Azure OpenAI is not configured. Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY (AZURE_OPENAI_DEPLOYMENT is optional).";
  }

  const safeText = clampText(docText, 12000);
  const safeQ = clampText(question, 800);
  return await chatComplete(
    cfg,
    [
      {
        role: "system",
        content:
          "Answer ONLY from the provided document text. If the answer is not in the document, say: 'Not found in the document.' Keep it short and precise.",
      },
      { role: "user", content: `Document:\n\n${safeText}\n\nQuestion: ${safeQ}` },
    ],
    700
  );
}

