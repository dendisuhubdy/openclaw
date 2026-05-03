import { Buffer } from "node:buffer";
import { spawn } from "node:child_process";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
import { Type } from "typebox";
import { saveMediaBuffer } from "../../media/store.js";
import {
  asToolParamsRecord,
  jsonResult,
  readStringParam,
  ToolInputError,
  type AnyAgentTool,
} from "./common.js";

const DEFAULT_MAX_BYTES = 25 * 1024 * 1024; // 25 MB
const MEDIA_SUBDIR = "tool-office-generation";

const SUPPORTED_FORMATS = ["docx", "pptx", "xlsx", "pdf"] as const;
type OfficeFormat = (typeof SUPPORTED_FORMATS)[number];

const FORMAT_MIME: Record<OfficeFormat, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pdf: "application/pdf",
};

export const OfficeGenerateSchema = Type.Object({
  format: Type.String({
    description: 'Output format. One of "docx", "pptx", "xlsx", "pdf".',
  }),
  filename: Type.Optional(
    Type.String({
      description:
        "Filename without extension. Will be sanitized; defaults to 'document' if omitted.",
    }),
  ),
  title: Type.Optional(
    Type.String({
      description:
        "Optional document title. Prepended to content as a top-level heading if provided.",
    }),
  ),
  content: Type.String({
    description:
      "Markdown source. Headings (#, ##), **bold**, *italic*, lists, tables, and code " +
      "blocks all render natively into docx/pptx via pandoc. For pptx, use a `---` line " +
      "or any heading to start a new slide. For xlsx, supply one or more markdown tables " +
      "(each table = sheet); if no tables, content is parsed as CSV (one row per line).",
  }),
});

function parseFormat(raw: unknown): OfficeFormat {
  if (typeof raw !== "string") {
    throw new ToolInputError("format is required");
  }
  const normalized = raw.trim().toLowerCase();
  if (!SUPPORTED_FORMATS.includes(normalized as OfficeFormat)) {
    throw new ToolInputError(
      `format must be one of: ${SUPPORTED_FORMATS.join(", ")} (got "${raw}")`,
    );
  }
  return normalized as OfficeFormat;
}

function sanitizeFilenameStem(input: string | undefined): string {
  const fallback = "document";
  if (!input) return fallback;
  const stem = input
    .replace(/\.[a-z0-9]{1,8}$/i, "")
    .replace(/[^A-Za-z0-9_\-. ]+/g, "_")
    .trim();
  return stem || fallback;
}

function prependTitle(title: string | undefined, content: string): string {
  if (!title) return content;
  return `# ${title}\n\n${content}`;
}

// Run pandoc with stdin → stdout. Collects the binary output buffer.
function runPandoc(args: string[], input: string): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const child = spawn("pandoc", args, { stdio: ["pipe", "pipe", "pipe"] });
    const out: Buffer[] = [];
    const err: Buffer[] = [];
    child.stdout.on("data", (c: Buffer) => out.push(c));
    child.stderr.on("data", (c: Buffer) => err.push(c));
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) {
        resolve(Buffer.concat(out));
      } else {
        reject(
          new Error(
            `pandoc ${args.join(" ")} exited with code ${code}: ${Buffer.concat(err).toString("utf8").slice(0, 500)}`,
          ),
        );
      }
    });
    child.stdin.end(input, "utf8");
  });
}

async function buildDocxFromMarkdown(title: string | undefined, content: string): Promise<Buffer> {
  return runPandoc(["-f", "markdown", "-t", "docx"], prependTitle(title, content));
}

async function buildPptxFromMarkdown(title: string | undefined, content: string): Promise<Buffer> {
  // Pandoc's slide-level convention: headings start new slides. Honor `---` markers
  // by translating them to top-level headings so the user can author either way.
  const normalized = content.replace(/\r\n/g, "\n").replace(/^---\s*$/gm, "## ---");
  return runPandoc(["-f", "markdown", "-t", "pptx"], prependTitle(title, normalized));
}

// ───── XLSX ─────────────────────────────────────────────────────────────────
// Strategy: scan content for GitHub-flavored markdown tables. Each table → sheet.
// If no tables found, fall back to CSV-per-line (backward compat for the v0.6.79 callers).

type MarkdownTable = { name: string; rows: string[][] };

function extractMarkdownTables(content: string): MarkdownTable[] {
  const lines = content.replace(/\r\n/g, "\n").split("\n");
  const tables: MarkdownTable[] = [];
  let i = 0;
  let lastHeading = "";
  while (i < lines.length) {
    const line = lines[i] ?? "";
    const headingMatch = line.match(/^#{1,6}\s+(.+?)\s*$/);
    if (headingMatch) {
      lastHeading = headingMatch[1] ?? "";
      i++;
      continue;
    }
    // Detect a table: a `| … |` line followed by a `| --- | --- |` separator.
    if (line.startsWith("|") && i + 1 < lines.length) {
      const sep = lines[i + 1] ?? "";
      if (/^\|[\s\-:|]+\|$/.test(sep.trim())) {
        const rows: string[][] = [splitMdRow(line)];
        i += 2;
        while (i < lines.length && (lines[i] ?? "").trim().startsWith("|")) {
          rows.push(splitMdRow(lines[i] ?? ""));
          i++;
        }
        tables.push({ name: lastHeading || `Sheet${tables.length + 1}`, rows });
        continue;
      }
    }
    i++;
  }
  return tables;
}

function splitMdRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  return trimmed.split("|").map((cell) => cell.trim());
}

function parseCsvLine(line: string): string[] {
  const cells: string[] = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cur += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ",") {
        cells.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
  }
  cells.push(cur);
  return cells;
}

function safeSheetName(name: string, fallback: string): string {
  const cleaned = name
    .slice(0, 31)
    .replace(/[\\/?*\[\]:]/g, "_")
    .trim();
  return cleaned || fallback;
}

async function buildXlsx(title: string | undefined, content: string): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  if (title) {
    workbook.creator = title;
    workbook.title = title;
  }
  const tables = extractMarkdownTables(content);
  if (tables.length > 0) {
    tables.forEach((t, idx) => {
      const sheet = workbook.addWorksheet(safeSheetName(t.name, `Sheet${idx + 1}`));
      for (const row of t.rows) sheet.addRow(row);
    });
  } else {
    // CSV fallback: blank-line-separated blocks become sheets; first non-csv line of a block names it.
    const blocks = content
      .replace(/\r\n/g, "\n")
      .split(/\n\s*\n/)
      .map((b) => b.trim())
      .filter(Boolean);
    const sheetBlocks = blocks.length > 0 ? blocks : [""];
    sheetBlocks.forEach((block, idx) => {
      const lines = block.split("\n");
      let name = `Sheet${idx + 1}`;
      let dataLines = lines;
      if (sheetBlocks.length > 1 && lines.length > 0 && !lines[0].includes(",")) {
        const candidate = lines[0].trim();
        if (candidate) {
          name = safeSheetName(candidate, name);
          dataLines = lines.slice(1);
        }
      }
      const sheet = workbook.addWorksheet(name);
      for (const line of dataLines) sheet.addRow(parseCsvLine(line));
    });
  }
  const ab = await workbook.xlsx.writeBuffer();
  return Buffer.from(ab as ArrayBuffer);
}

// ───── PDF ──────────────────────────────────────────────────────────────────
// No PDF engine (wkhtmltopdf/latex) is installed in the agent container, so
// pandoc -t pdf can't be used. Convert markdown → plain via pandoc, then render
// with pdfkit. This loses some formatting but produces a valid PDF.

async function buildPdfFromMarkdown(title: string | undefined, content: string): Promise<Buffer> {
  let plain: string;
  try {
    const buf = await runPandoc(
      ["-f", "markdown", "-t", "plain", "--wrap=preserve"],
      prependTitle(title, content),
    );
    plain = buf.toString("utf8");
  } catch {
    // Fallback: render the markdown source as-is if pandoc isn't reachable.
    plain = prependTitle(title, content);
  }
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: "LETTER", margin: 60 });
      const chunks: Buffer[] = [];
      doc.on("data", (c: Buffer) => chunks.push(c));
      doc.on("end", () => resolve(Buffer.concat(chunks)));
      doc.on("error", reject);
      doc.fontSize(12);
      const paragraphs = plain
        .replace(/\r\n/g, "\n")
        .split(/\n\s*\n/)
        .map((p) => p.trim())
        .filter(Boolean);
      for (const para of paragraphs) {
        doc.text(para, { align: "left" });
        doc.moveDown(0.5);
      }
      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}

async function buildDocument(
  format: OfficeFormat,
  title: string | undefined,
  content: string,
): Promise<Buffer> {
  switch (format) {
    case "docx":
      return buildDocxFromMarkdown(title, content);
    case "pptx":
      return buildPptxFromMarkdown(title, content);
    case "xlsx":
      return buildXlsx(title, content);
    case "pdf":
      return buildPdfFromMarkdown(title, content);
  }
}

export function createOfficeGenerateTool(): AnyAgentTool {
  return {
    label: "Office",
    name: "office_generate",
    description:
      "Convert markdown into a real Microsoft Word (.docx), PowerPoint (.pptx), Excel " +
      "(.xlsx), or PDF binary file and return a download link. " +
      "MANDATORY: Use this tool — NOT the `write` tool — for any user request to produce " +
      "a .docx / .pptx / .xlsx / .pdf file. The `write` tool only writes plain text and " +
      "will produce a corrupt non-Word file if used for these formats. This tool runs " +
      "the markdown through pandoc, so headings, bold, italic, lists, and tables all " +
      "render as native Word/PowerPoint/Excel content.",
    parameters: OfficeGenerateSchema,
    execute: async (_toolCallId, args) => {
      const params = asToolParamsRecord(args);
      const format = parseFormat(params.format);
      const title = readStringParam(params, "title");
      const content = readStringParam(params, "content", { required: true });
      const stem = sanitizeFilenameStem(readStringParam(params, "filename"));
      const filename = `${stem}.${format}`;

      const buffer = await buildDocument(format, title, content);
      const saved = await saveMediaBuffer(
        buffer,
        FORMAT_MIME[format],
        MEDIA_SUBDIR,
        DEFAULT_MAX_BYTES,
        filename,
      );

      return jsonResult({
        format,
        filename,
        size_bytes: saved.size,
        path: saved.path,
        media_paths: [saved.path],
      });
    },
  };
}
