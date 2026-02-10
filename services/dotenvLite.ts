import { readFile } from "node:fs/promises";
import path from "node:path";

function parseDotenv(contents: string): Record<string, string> {
  const out: Record<string, string> = {};
  const lines = contents.split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq <= 0) continue;
    const key = line.slice(0, eq).trim();
    let value = line.slice(eq + 1).trim();
    // Remove wrapping quotes if present
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (key) out[key] = value;
  }
  return out;
}

async function tryReadFileUtf8(filePath: string): Promise<string | null> {
  try {
    return await readFile(filePath, "utf8");
  } catch {
    return null;
  }
}

function resolveCandidates(rel: string): string[] {
  // Try from CWD and from repo-root-ish relative to compiled location.
  return [
    path.join(process.cwd(), rel),
    path.join(__dirname, "..", "..", rel), // services -> repo root
    path.join(__dirname, "..", rel), // lib/services -> lib
  ];
}

/**
 * Load specific env vars from local dotenv-style files into process.env
 * ONLY when they're missing.
 */
export async function loadEnvFromLocalFilesIfMissing(keys: string[]): Promise<void> {
  const missing = keys.filter((k) => !String(process.env[k] ?? "").trim());
  if (missing.length === 0) return;

  const candidateFiles = [
    ...resolveCandidates(".localConfigs"),
    ...resolveCandidates(".localConfigs.playground"),
    ...resolveCandidates("env/.env.local"),
    ...resolveCandidates("env/.env"),
    ...resolveCandidates("env/.env.dev"),
    ...resolveCandidates("env/.env.playground"),
  ];

  for (const p of candidateFiles) {
    const raw = await tryReadFileUtf8(p);
    if (!raw) continue;
    const parsed = parseDotenv(raw);
    for (const k of missing) {
      const v = parsed[k];
      if (typeof v === "string" && v.trim()) {
        process.env[k] = v.trim();
      }
    }
    // recompute missing
    const stillMissing = keys.filter((k) => !String(process.env[k] ?? "").trim());
    if (stillMissing.length === 0) return;
  }
}

