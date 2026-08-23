"use client";

import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  TextRun,
  type ISectionOptions,
} from "docx";
import { cmykToRgbHex, estimateTextColor, rgbToHex } from "./color";
import { outputName } from "./file-utils";
import type { ConversionResult, ConversionSettings, ProgressUpdate } from "./types";

type ProgressCallback = (update: ProgressUpdate) => void;
type PdfTextItem = {
  str: string;
  transform: number[];
  width: number;
  height: number;
  fontName: string;
  hasEOL?: boolean;
};

type StyledTextItem = PdfTextItem & { color: string };
type PdfLine = { items: StyledTextItem[]; y: number; fontSize: number };
type BBox = { x0: number; y0: number; x1: number; y1: number };

type OcrWord = { text: string; bbox: BBox; confidence?: number };
type OcrLine = { text: string; bbox: BBox; words: OcrWord[]; confidence?: number };

type TesseractWorker = Awaited<ReturnType<typeof import("tesseract.js")["createWorker"]>>;

const PAGE_WIDTH_TWIPS = 12240; // 8.5"
const PAGE_MARGIN_TWIPS = 720; // 0.5"
const CONTENT_WIDTH_TWIPS = PAGE_WIDTH_TWIPS - PAGE_MARGIN_TWIPS * 2;
const PDFJS_VERSION = "6.2.108";
const PDFJS_CDN = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}`;
const MAX_RENDER_SIDE = 3600;
const MAX_RENDER_PIXELS = 12_000_000;

function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

function cleanText(text: string) {
  return text.replace(/\u0000/g, "").replace(/[ \t]+/g, " ").trim();
}

function numberArrayFromView(view: ArrayBufferView): number[] {
  if (view instanceof DataView) return [];
  return Array.from(view as unknown as ArrayLike<number>);
}

function normalizeColorArgs(args: unknown): number[] {
  if (Array.isArray(args)) {
    return args.flatMap((value) => {
      if (typeof value === "number") return [value];
      if (ArrayBuffer.isView(value)) return numberArrayFromView(value);
      return [];
    });
  }
  if (ArrayBuffer.isView(args)) return numberArrayFromView(args);
  return [];
}

async function extractPdfColors(page: any, pdfjs: typeof import("pdfjs-dist")) {
  const list = await page.getOperatorList();
  const colors: string[] = [];
  let current = "111827";

  for (let i = 0; i < list.fnArray.length; i += 1) {
    const fn = list.fnArray[i];
    const args = normalizeColorArgs(list.argsArray[i]);
    if (fn === pdfjs.OPS.setFillRGBColor && args.length >= 3) {
      current = rgbToHex(args[0], args[1], args[2]);
    } else if (fn === pdfjs.OPS.setFillGray && args.length >= 1) {
      current = rgbToHex(args[0], args[0], args[0]);
    } else if (fn === pdfjs.OPS.setFillCMYKColor && args.length >= 4) {
      current = cmykToRgbHex(args[0], args[1], args[2], args[3]);
    } else if (
      fn === pdfjs.OPS.showText ||
      fn === pdfjs.OPS.showSpacedText ||
      fn === pdfjs.OPS.nextLineShowText ||
      fn === pdfjs.OPS.nextLineSetSpacingShowText
    ) {
      colors.push(current);
    }
  }
  return colors;
}

function groupPdfLines(items: StyledTextItem[]) {
  if (!items.length) return [] as PdfLine[];
  const sorted = [...items].sort((a, b) => b.transform[5] - a.transform[5] || a.transform[4] - b.transform[4]);
  const sizes = sorted.map((item) => Math.max(5, Math.hypot(item.transform[0], item.transform[1]))).sort((a, b) => a - b);
  const median = sizes[Math.floor(sizes.length / 2)] || 11;
  const tolerance = Math.max(2.5, median * 0.45);
  const lines: PdfLine[] = [];

  for (const item of sorted) {
    if (!cleanText(item.str)) continue;
    const y = item.transform[5];
    let line = lines.find((candidate) => Math.abs(candidate.y - y) <= tolerance);
    const fontSize = Math.max(6, Math.hypot(item.transform[0], item.transform[1]));
    if (!line) {
      line = { items: [], y, fontSize };
      lines.push(line);
    }
    line.items.push(item);
    line.fontSize = Math.max(line.fontSize, fontSize);
  }

  return lines
    .map((line) => ({ ...line, items: line.items.sort((a, b) => a.transform[4] - b.transform[4]) }))
    .sort((a, b) => b.y - a.y);
}

function alignmentForBounds(x0: number, x1: number, width: number) {
  const center = (x0 + x1) / 2;
  const pageCenter = width / 2;
  const lineWidth = x1 - x0;
  if (lineWidth < width * 0.78 && Math.abs(center - pageCenter) < width * 0.07) return AlignmentType.CENTER;
  if (x0 > width * 0.52 && x1 > width * 0.88) return AlignmentType.RIGHT;
  return AlignmentType.LEFT;
}

function pdfLineToParagraph(line: PdfLine, pageWidth: number, previousY?: number) {
  const first = line.items[0];
  const last = line.items[line.items.length - 1];
  const x0 = first.transform[4];
  const x1 = last.transform[4] + (last.width || 0);
  const alignment = alignmentForBounds(x0, x1, pageWidth);
  const fontSize = clamp(line.fontSize, 7, 44);
  const isHeading = fontSize >= 16 && line.items.map((i) => i.str).join("").length < 100;
  const leftIndent = alignment === AlignmentType.LEFT ? Math.max(0, Math.round((x0 / pageWidth) * CONTENT_WIDTH_TWIPS) - 80) : 0;
  const gap = previousY == null ? 0 : Math.max(0, previousY - line.y - fontSize);
  const runs: TextRun[] = [];

  line.items.forEach((item, index) => {
    const currentX = item.transform[4];
    const prev = line.items[index - 1];
    if (prev) {
      const prevEnd = prev.transform[4] + (prev.width || 0);
      const gapX = currentX - prevEnd;
      if (gapX > Math.max(1.5, fontSize * 0.2) && !item.str.startsWith(" ")) runs.push(new TextRun({ text: " " }));
    }
    const itemSize = clamp(Math.hypot(item.transform[0], item.transform[1]), 7, 44);
    const fontHint = item.fontName.toLowerCase();
    runs.push(
      new TextRun({
        text: item.str,
        size: Math.round(itemSize * 2),
        color: item.color,
        bold: isHeading || /bold|black|heavy|semibold|demi/.test(fontHint),
        italics: /italic|oblique/.test(fontHint),
        font: "Aptos",
      }),
    );
  });

  return new Paragraph({
    children: runs,
    alignment,
    heading: isHeading ? HeadingLevel.HEADING_2 : undefined,
    indent: leftIndent ? { left: leftIndent } : undefined,
    spacing: {
      before: Math.round(clamp(gap * 10, 0, 240)),
      after: Math.round(clamp(fontSize * 2.5, 15, 90)),
      line: Math.round(clamp(fontSize * 20 * 1.08, 180, 660)),
    },
  });
}

async function canvasFromPdfPage(page: any, scale = 2) {
  const base = page.getViewport({ scale });
  const sideRatio = Math.min(1, MAX_RENDER_SIDE / Math.max(base.width, base.height));
  const pixelRatio = Math.min(1, Math.sqrt(MAX_RENDER_PIXELS / Math.max(1, base.width * base.height)));
  const safeScale = scale * Math.min(sideRatio, pixelRatio);
  const viewport = page.getViewport({ scale: safeScale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) throw new Error("Your browser could not create a canvas context.");
  await page.render({ canvasContext: ctx, viewport, canvas }).promise;
  return { canvas, ctx, viewport, scale };
}

function canvasToPngBytes(canvas: HTMLCanvasElement) {
  return new Promise<Uint8Array>((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) return reject(new Error("Could not render page image."));
      resolve(new Uint8Array(await blob.arrayBuffer()));
    }, "image/png", 0.96);
  });
}

async function imageFileToCanvas(file: File) {
  let source: CanvasImageSource;
  let width = 0;
  let height = 0;
  let cleanup = () => {};

  try {
    const bitmap = await createImageBitmap(file);
    source = bitmap;
    width = bitmap.width;
    height = bitmap.height;
    cleanup = () => bitmap.close();
  } catch {
    const url = URL.createObjectURL(file);
    const image = new Image();
    image.decoding = "async";
    image.src = url;
    try {
      await image.decode();
    } catch {
      URL.revokeObjectURL(url);
      throw new Error("This image format could not be decoded by your browser. If it is HEIC/HEIF, export it as JPG or PNG and try again.");
    }
    source = image;
    width = image.naturalWidth;
    height = image.naturalHeight;
    cleanup = () => URL.revokeObjectURL(url);
  }

  const sideRatio = Math.min(1, MAX_RENDER_SIDE / Math.max(width, height));
  const pixelRatio = Math.min(1, Math.sqrt(MAX_RENDER_PIXELS / Math.max(1, width * height)));
  const ratio = Math.min(sideRatio, pixelRatio);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width * ratio));
  canvas.height = Math.max(1, Math.round(height * ratio));
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) {
    cleanup();
    throw new Error("Your browser could not create a canvas context.");
  }
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  cleanup();
  return { canvas, ctx, scale: ratio };
}

function flattenOcrLines(blocks: any[] | null | undefined): OcrLine[] {
  if (!blocks?.length) return [];
  const lines: OcrLine[] = [];
  for (const block of blocks) {
    for (const paragraph of block.paragraphs ?? []) {
      for (const line of paragraph.lines ?? []) {
        const words: OcrWord[] = (line.words ?? [])
          .map((word: any) => ({ text: cleanText(word.text ?? ""), bbox: word.bbox, confidence: word.confidence }))
          .filter((word: OcrWord) => word.text);
        const text = cleanText(line.text ?? words.map((w) => w.text).join(" "));
        if (text && line.bbox) lines.push({ text, bbox: line.bbox, words, confidence: line.confidence });
      }
    }
  }
  return lines;
}

async function ocrCanvas(
  canvas: HTMLCanvasElement,
  ctx: CanvasRenderingContext2D,
  worker: TesseractWorker,
  settings: ConversionSettings,
  pageNumber: number,
  pageCount: number,
  onProgress: ProgressCallback,
) {
  const result = await worker.recognize(canvas, {}, { blocks: true });
  const lines = flattenOcrLines(result.data.blocks as any[] | null);
  const medianHeights = lines.map((line) => line.bbox.y1 - line.bbox.y0).sort((a, b) => a - b);
  const median = medianHeights[Math.floor(medianHeights.length / 2)] || 26;
  const children: Paragraph[] = [];

  lines.sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
  lines.forEach((line, lineIndex) => {
    const h = Math.max(8, line.bbox.y1 - line.bbox.y0);
    const fontPt = clamp((h / canvas.height) * 792 * 0.85, 7, 38);
    const isHeading = h > median * 1.35 && line.text.length < 100;
    const alignment = alignmentForBounds(line.bbox.x0, line.bbox.x1, canvas.width);
    const leftIndent = alignment === AlignmentType.LEFT ? Math.max(0, Math.round((line.bbox.x0 / canvas.width) * CONTENT_WIDTH_TWIPS) - 100) : 0;
    const runs: TextRun[] = [];

    if (settings.mode === "clean" || !line.words.length) {
      const color = settings.preserveColor ? estimateTextColor(ctx, line.bbox) : "111827";
      runs.push(new TextRun({ text: line.text, size: Math.round(fontPt * 2), color, bold: isHeading, font: "Aptos" }));
    } else {
      line.words.forEach((word, wordIndex) => {
        const color = settings.preserveColor ? estimateTextColor(ctx, word.bbox) : "111827";
        runs.push(
          new TextRun({
            text: `${wordIndex ? " " : ""}${word.text}`,
            size: Math.round(fontPt * 2),
            color,
            bold: isHeading,
            font: "Aptos",
          }),
        );
      });
    }

    const next = lines[lineIndex + 1];
    const verticalGap = next ? Math.max(0, next.bbox.y0 - line.bbox.y1) : 0;
    children.push(
      new Paragraph({
        children: runs,
        alignment,
        heading: isHeading ? HeadingLevel.HEADING_2 : undefined,
        indent: settings.preserveLayout && leftIndent ? { left: leftIndent } : undefined,
        spacing: {
          after: Math.round(clamp((verticalGap / canvas.height) * 12000, 20, 220)),
          line: Math.round(clamp(fontPt * 21, 180, 620)),
        },
      }),
    );
  });

  onProgress({
    value: 22 + Math.round((pageNumber / pageCount) * 62),
    phase: "OCR complete",
    detail: `Read ${lines.length} lines on page ${pageNumber}.`,
  });
  return children;
}

function pixelCloneParagraph(png: Uint8Array, canvas: HTMLCanvasElement) {
  const targetWidth = 720;
  const targetHeight = Math.round((canvas.height / canvas.width) * targetWidth);
  return new Paragraph({
    children: [
      new ImageRun({
        data: png,
        transformation: { width: targetWidth, height: targetHeight },
        type: "png",
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
  });
}

async function makeWorker(settings: ConversionSettings, onProgress: ProgressCallback) {
  const { createWorker, OEM } = await import("tesseract.js");
  const worker = await createWorker(settings.language.split("+"), OEM.LSTM_ONLY, {
    logger: (message) => {
      if (message.status === "recognizing text") {
        onProgress({
          value: 25 + Math.round(message.progress * 50),
          phase: "Reading the page",
          detail: `${Math.round(message.progress * 100)}% OCR`,
        });
      }
    },
  });
  await worker.setParameters({ preserve_interword_spaces: "1", user_defined_dpi: "300" });
  return worker;
}

function baseSections(children: Paragraph[]): ISectionOptions[] {
  return [
    {
      properties: {
        page: {
          margin: { top: PAGE_MARGIN_TWIPS, right: PAGE_MARGIN_TWIPS, bottom: PAGE_MARGIN_TWIPS, left: PAGE_MARGIN_TWIPS },
        },
      },
      children,
    },
  ];
}

export async function convertToDocx(
  file: File,
  settings: ConversionSettings,
  onProgress: ProgressCallback,
): Promise<ConversionResult> {
  onProgress({ value: 2, phase: "Opening file", detail: "Preparing your document locally…" });
  const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
  const docChildren: Paragraph[] = [];
  let pageCount = 1;
  let usedOcr = false;
  let worker: TesseractWorker | null = null;

  try {
    if (isPdf) {
      const pdfjs = await import("pdfjs-dist");
      pdfjs.GlobalWorkerOptions.workerSrc = `${PDFJS_CDN}/build/pdf.worker.min.mjs`;
      const bytes = new Uint8Array(await file.arrayBuffer());
      const loadingTask = pdfjs.getDocument({
        data: bytes,
        cMapUrl: `${PDFJS_CDN}/cmaps/`,
        cMapPacked: true,
        standardFontDataUrl: `${PDFJS_CDN}/standard_fonts/`,
        wasmUrl: `${PDFJS_CDN}/wasm/`,
      });
      const pdf = await loadingTask.promise;
      pageCount = pdf.numPages;

      try {
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
          const page = await pdf.getPage(pageNumber);
          onProgress({
            value: 8 + Math.round(((pageNumber - 1) / pageCount) * 70),
            phase: `Page ${pageNumber} of ${pageCount}`,
            detail: "Analyzing layout, text and color…",
          });

          const textContent = await page.getTextContent({ includeMarkedContent: false });
          const rawItems = (textContent.items as any[]).filter((item) => "str" in item) as PdfTextItem[];
          const characterCount = rawItems.reduce((sum, item) => sum + cleanText(item.str).length, 0);
          const digitalPage = characterCount >= 24 && rawItems.length >= 3;
          const viewport = page.getViewport({ scale: 1 });

          if (settings.mode === "pixel") {
            const rendered = await canvasFromPdfPage(page, 2);
            const png = await canvasToPngBytes(rendered.canvas);
            docChildren.push(pixelCloneParagraph(png, rendered.canvas));
            if (!worker) worker = await makeWorker(settings, onProgress);
            usedOcr = true;
            docChildren.push(
              new Paragraph({
                children: [new TextRun({ text: "Editable text", bold: true, color: "64748B", size: 18 })],
                spacing: { before: 120, after: 70 },
              }),
            );
            docChildren.push(...(await ocrCanvas(rendered.canvas, rendered.ctx, worker, { ...settings, preserveLayout: false }, pageNumber, pageCount, onProgress)));
          } else if (digitalPage) {
            const colorSequence = settings.preserveColor ? await extractPdfColors(page, pdfjs) : [];
            const styled: StyledTextItem[] = rawItems.map((item, index) => ({
              ...item,
              color: settings.preserveColor ? colorSequence[Math.min(index, Math.max(0, colorSequence.length - 1))] ?? "111827" : "111827",
            }));
            const lines = groupPdfLines(styled);
            lines.forEach((line, index) => {
              docChildren.push(pdfLineToParagraph(line, viewport.width, index ? lines[index - 1].y : undefined));
            });
          } else {
            const rendered = await canvasFromPdfPage(page, 2.2);
            if (!worker) worker = await makeWorker(settings, onProgress);
            usedOcr = true;
            docChildren.push(...(await ocrCanvas(rendered.canvas, rendered.ctx, worker, settings, pageNumber, pageCount, onProgress)));
          }

          if (pageNumber < pdf.numPages) docChildren.push(new Paragraph({ children: [new PageBreak()] }));
          page.cleanup();
        }
      } finally {
        await loadingTask.destroy();
      }
    } else {
      const rendered = await imageFileToCanvas(file);
      worker = await makeWorker(settings, onProgress);
      usedOcr = true;
      if (settings.mode === "pixel") {
        const png = await canvasToPngBytes(rendered.canvas);
        docChildren.push(pixelCloneParagraph(png, rendered.canvas));
        docChildren.push(
          new Paragraph({
            children: [new TextRun({ text: "Editable text", bold: true, color: "64748B", size: 18 })],
            spacing: { before: 120, after: 70 },
          }),
        );
      }
      docChildren.push(...(await ocrCanvas(rendered.canvas, rendered.ctx, worker, settings, 1, 1, onProgress)));
    }

    onProgress({ value: 90, phase: "Building Word file", detail: "Packing editable DOCX…" });
    if (!docChildren.length) {
      docChildren.push(new Paragraph({ children: [new TextRun("No readable text was detected in this file.")] }));
    }

    const doc = new Document({
      creator: "PX Paper — by pxcheco",
      title: file.name,
      description: "Editable document generated locally by PX Paper.",
      sections: baseSections(docChildren),
      styles: {
        default: {
          document: { run: { font: "Aptos", size: 22, color: "111827" }, paragraph: { spacing: { after: 70, line: 276 } } },
          heading2: { run: { font: "Aptos Display", bold: true, color: "111827" } },
        },
      },
    });
    const blob = await Packer.toBlob(doc);
    onProgress({ value: 100, phase: "Done", detail: "Your editable Word file is ready." });
    return { blob, fileName: outputName(file.name), pages: pageCount, sourceKind: isPdf ? "pdf" : "image", usedOcr };
  } finally {
    if (worker) await worker.terminate();
  }
}
