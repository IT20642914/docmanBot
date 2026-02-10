import { readFile } from "node:fs/promises";
import path from "node:path";

function clampText(input: string, maxChars: number): string {
  const s = String(input ?? "");
  if (s.length <= maxChars) return s;
  return `${s.slice(0, maxChars)}\n\n[TRUNCATED]`;
}

export async function extractTextFromFile(filePath: string): Promise<string> {
  const ext = path.extname(filePath).toLowerCase();

  // Plain text
  if (ext === ".txt" || ext === ".md") {
    try {
      return await readFile(filePath, "utf8");
    } catch {
      return "";
    }
  }

  // PDF
  if (ext === ".pdf") {
    try {
      const pdfParseMod: any = await import("pdf-parse");
      const pdfParse = pdfParseMod?.default ?? pdfParseMod;
      const buf = await readFile(filePath);
      const parsed = await pdfParse(buf);
      return clampText(String(parsed?.text ?? ""), 12000);
    } catch {
      return "";
    }
  }

  // DOCX
  if (ext === ".docx") {
    try {
      const mammoth: any = await import("mammoth");
      const buf = await readFile(filePath);
      const result = await mammoth.extractRawText({ buffer: buf });
      return clampText(String(result?.value ?? ""), 12000);
    } catch {
      return "";
    }
  }

  // Excel: xlsx / xlsb
  if (ext === ".xlsx" || ext === ".xlsb" || ext === ".xls") {
    try {
      const xlsx: any = await import("xlsx");
      const buf = await readFile(filePath);
      const wb = xlsx.read(buf, { type: "buffer" });
      const out: string[] = [];
      const sheetNames: string[] = Array.isArray(wb?.SheetNames) ? wb.SheetNames : [];

      for (const name of sheetNames.slice(0, 5)) {
        const ws = wb.Sheets?.[name];
        if (!ws) continue;
        const csv = xlsx.utils.sheet_to_csv(ws, { FS: "\t" });
        out.push(`Sheet: ${name}\n${csv}`);
      }

      return clampText(out.join("\n\n"), 12000);
    } catch {
      return "";
    }
  }

  // Unknown/binary
  return "";
}

