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
      const buf = await readFile(filePath);
      const pdfMod: any = await import("pdf-parse");
      const PDFParse = pdfMod?.PDFParse;
      if (!PDFParse) return "";
      const parser = new PDFParse({ data: buf });
      const result = await parser.getText();
      await parser.destroy().catch(() => {});
      return clampText(String(result?.text ?? ""), 12000);
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

  // PowerPoint: pptx (zip) - extract text from slide XML
  if (ext === ".pptx") {
    try {
      const JSZipMod: any = await import("jszip");
      const JSZip = JSZipMod?.default ?? JSZipMod;
      const buf = await readFile(filePath);
      const zip = await JSZip.loadAsync(buf);

      const slidePaths = Object.keys(zip.files)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
        .sort((a, b) => {
          const na = Number(a.match(/slide(\d+)\.xml/i)?.[1] ?? 0);
          const nb = Number(b.match(/slide(\d+)\.xml/i)?.[1] ?? 0);
          return na - nb;
        });

      const out: string[] = [];
      for (const sp of slidePaths.slice(0, 30)) {
        const xml = await zip.file(sp)?.async("string");
        if (!xml) continue;
        // Extract <a:t>Text</a:t> nodes
        const texts = [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/gi)]
          .map((m) => String(m[1] ?? "").replace(/\s+/g, " ").trim())
          .filter(Boolean);
        if (texts.length > 0) out.push(`Slide ${out.length + 1}: ${texts.join(" ")}`);
      }

      return clampText(out.join("\n"), 12000);
    } catch {
      return "";
    }
  }

  // Legacy .ppt is binary; not supported without specialized parser/OCR.
  if (ext === ".ppt") return "";

  // Unknown/binary
  return "";
}

function mimeFromExt(ext: string): string | null {
  const e = ext.toLowerCase();
  if (e === ".png") return "image/png";
  if (e === ".jpg" || e === ".jpeg") return "image/jpeg";
  if (e === ".webp") return "image/webp";
  return null;
}

export async function extractImagesFromPptx(filePath: string): Promise<string[]> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".pptx") return [];

  try {
    const JSZipMod: any = await import("jszip");
    const JSZip = JSZipMod?.default ?? JSZipMod;
    const buf = await readFile(filePath);
    const zip = await JSZip.loadAsync(buf);

    const allFiles = Object.keys(zip.files);

    // Prefer the thumbnail if present (often a rendered slide preview)
    const thumb = allFiles.find((p) => /^docProps\/thumbnail\.(jpe?g|png|webp)$/i.test(p));

    const mediaPaths = allFiles
      .filter((p) => /^ppt\/media\//i.test(p))
      .filter((p) => Boolean(mimeFromExt(path.extname(p))));

    const pick = [
      ...(thumb ? [thumb] : []),
      ...mediaPaths,
    ]
      .filter(Boolean)
      .slice(0, 4); // limit images to avoid huge prompts

    const out: string[] = [];
    for (const mp of pick) {
      const mime = mimeFromExt(path.extname(mp));
      if (!mime) continue;
      const file = zip.file(mp);
      if (!file) continue;
      const b64 = await file.async("base64");
      out.push(`data:${mime};base64,${b64}`);
    }
    return out;
  } catch {
    return [];
  }
}

