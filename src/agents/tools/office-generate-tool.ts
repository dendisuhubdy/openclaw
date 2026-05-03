import { Buffer } from "node:buffer";
import {
  AlignmentType,
  Document as DocxDocument,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
// pptxgenjs has both `export default` and `export as namespace`; under NodeNext + tsgo
// the default-import binding is treated as the namespace type. Re-import via `* as`
// and pull off `.default` to land on the actual class.
import * as PptxGenJSModule from "pptxgenjs";
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
      description: "Document title. Used as docx/pdf heading and pptx title slide.",
    }),
  ),
  content: Type.String({
    description:
      "Body content. Markdown-ish for docx/pdf (one paragraph per blank line). " +
      "For pptx, slides are separated by a line containing only '---'; each slide's first " +
      "line becomes the title and remaining lines become bullets. For xlsx, content is " +
      "interpreted as CSV (one row per line, comma-separated cells); use a blank line to " +
      "split into multiple sheets, where the first line of each sheet is the sheet name.",
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

function splitParagraphs(content: string): string[] {
  return content
    .replace(/\r\n/g, "\n")
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);
}

async function buildDocx(title: string | undefined, content: string): Promise<Buffer> {
  const children: Paragraph[] = [];
  if (title) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text: title, bold: true })],
      }),
    );
  }
  for (const para of splitParagraphs(content)) {
    children.push(new Paragraph({ children: [new TextRun(para)] }));
  }
  if (children.length === 0) {
    children.push(new Paragraph({ children: [new TextRun("")] }));
  }
  const doc = new DocxDocument({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}

type PptxSlide = {
  addText: (
    text: string | { text: string; options?: Record<string, unknown> }[],
    options?: Record<string, unknown>,
  ) => unknown;
};

type PptxGenJSInstance = {
  layout: string;
  addSlide: () => PptxSlide;
  write: (opts: { outputType: "nodebuffer" }) => Promise<Buffer | ArrayBuffer | string>;
};

async function buildPptx(title: string | undefined, content: string): Promise<Buffer> {
  const PptxGenJSCtor = (PptxGenJSModule as unknown as { default: new () => PptxGenJSInstance })
    .default;
  const pptx: PptxGenJSInstance = new PptxGenJSCtor();
  pptx.layout = "LAYOUT_WIDE";

  if (title) {
    const titleSlide = pptx.addSlide();
    titleSlide.addText(title, {
      x: 0.5,
      y: 2.5,
      w: 12,
      h: 1.5,
      fontSize: 44,
      bold: true,
      align: "center",
    });
  }

  const slides = content
    .replace(/\r\n/g, "\n")
    .split(/\n---\n/)
    .map((s) => s.trim())
    .filter(Boolean);

  if (slides.length === 0) {
    slides.push(content.trim() || "");
  }

  for (const slide of slides) {
    const lines = slide
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean);
    const slideTitle = lines[0] ?? "";
    const bullets = lines.slice(1);
    const s = pptx.addSlide();
    if (slideTitle) {
      s.addText(slideTitle, {
        x: 0.5,
        y: 0.4,
        w: 12,
        h: 0.9,
        fontSize: 32,
        bold: true,
      });
    }
    if (bullets.length > 0) {
      s.addText(
        bullets.map((b) => ({ text: b.replace(/^[-*]\s*/, ""), options: { bullet: true } })),
        { x: 0.5, y: 1.4, w: 12, h: 5.5, fontSize: 20 },
      );
    }
  }

  // pptxgenjs returns Promise<string | ArrayBuffer | Buffer | Blob> depending on outputType
  const out = await pptx.write({ outputType: "nodebuffer" });
  if (Buffer.isBuffer(out)) return out;
  if (out instanceof ArrayBuffer) return Buffer.from(out);
  if (typeof out === "string") return Buffer.from(out, "binary");
  throw new Error("pptxgenjs returned unexpected output type");
}

function parseCsvLine(line: string): string[] {
  // Minimal CSV parsing: supports quoted fields with embedded commas / escaped quotes.
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

async function buildXlsx(title: string | undefined, content: string): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  if (title) {
    workbook.creator = title;
    workbook.title = title;
  }

  const blocks = content
    .replace(/\r\n/g, "\n")
    .split(/\n\s*\n/)
    .map((b) => b.trim())
    .filter(Boolean);

  const sheetBlocks = blocks.length > 0 ? blocks : [""];

  let sheetIndex = 0;
  for (const block of sheetBlocks) {
    sheetIndex++;
    const lines = block.split("\n");
    let sheetName = `Sheet${sheetIndex}`;
    let dataLines = lines;
    if (sheetBlocks.length > 1 && lines.length > 0 && !lines[0].includes(",")) {
      const candidate = lines[0].trim();
      if (candidate) {
        sheetName = candidate.slice(0, 31).replace(/[\\/?*\[\]:]/g, "_");
        dataLines = lines.slice(1);
      }
    }
    const sheet = workbook.addWorksheet(sheetName);
    for (const line of dataLines) {
      sheet.addRow(parseCsvLine(line));
    }
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer as ArrayBuffer);
}

async function buildPdf(title: string | undefined, content: string): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: "LETTER", margin: 60 });
      const chunks: Buffer[] = [];
      doc.on("data", (c: Buffer) => chunks.push(c));
      doc.on("end", () => resolve(Buffer.concat(chunks)));
      doc.on("error", reject);

      if (title) {
        doc.fontSize(20).text(title, { align: "left" });
        doc.moveDown();
      }
      doc.fontSize(12);
      for (const para of splitParagraphs(content)) {
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
      return buildDocx(title, content);
    case "pptx":
      return buildPptx(title, content);
    case "xlsx":
      return buildXlsx(title, content);
    case "pdf":
      return buildPdf(title, content);
  }
}

export function createOfficeGenerateTool(): AnyAgentTool {
  return {
    label: "Office Document Generator",
    name: "office_generate",
    description:
      "Generate an Office document (DOCX, PPTX, XLSX) or PDF from structured text and " +
      "save it to the user's files. Returns the saved media path which the host bridges to " +
      "a downloadable URL. Use for: drafting reports (DOCX/PDF), slide decks (PPTX), " +
      "or tabular data (XLSX). Content syntax: see the 'content' parameter description.",
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
