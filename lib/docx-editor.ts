"use client";

import JSZip from "jszip";
import {
  AlignmentType,
  BorderStyle,
  Document as DocxDocument,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const WORD_NS = "w:";

type RunState = {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  size?: number;
  font?: string;
  superScript?: boolean;
  subScript?: boolean;
};

type EditorChild = Paragraph | Table;
type RelationshipMap = Map<string, string>;

function childrenByTag(parent: Element, tagName: string) {
  return Array.from(parent.children).filter((child) => child.tagName === tagName);
}

function firstDirect(parent: Element | null | undefined, tagName: string) {
  if (!parent) return null;
  return Array.from(parent.children).find((child) => child.tagName === tagName) ?? null;
}

function firstDescendant(parent: Element | null | undefined, tagName: string) {
  if (!parent) return null;
  return parent.getElementsByTagName(tagName)[0] ?? null;
}

function attr(element: Element | null | undefined, name: string) {
  if (!element) return null;
  return element.getAttribute(name) ?? element.getAttribute(name.includes(":") ? name.split(":")[1] : name);
}

function escapeHtml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function safeHex(value: string | null, fallback = "111827") {
  if (!value || value.toLowerCase() === "auto") return fallback;
  const normalized = value.replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : fallback;
}

function emuToPx(value: string | null, fallback: number) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.max(1, Math.round(parsed / 9525));
}

function uint8ToBase64(bytes: Uint8Array) {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(binary);
}

function mimeForPath(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".bmp")) return "image/bmp";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  return "image/png";
}

function parseRelationships(xml: XMLDocument): RelationshipMap {
  const map = new Map<string, string>();
  for (const rel of Array.from(xml.getElementsByTagName("Relationship"))) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target && !target.includes("://")) map.set(id, target.replace(/^\.\//, ""));
  }
  return map;
}

async function imageHtmlFromRun(run: Element, zip: JSZip, relationships: RelationshipMap) {
  const blip = firstDescendant(run, "a:blip");
  const relId = attr(blip, "r:embed");
  if (!relId) return "";
  const target = relationships.get(relId);
  if (!target) return "";
  const normalized = target.startsWith("word/") ? target : `word/${target.replace(/^\.\.\//, "")}`;
  const imageFile = zip.file(normalized);
  if (!imageFile) return "";
  const bytes = await imageFile.async("uint8array");
  const extent = firstDescendant(run, "wp:extent");
  const width = Math.min(720, emuToPx(attr(extent, "cx"), 520));
  const height = Math.min(980, emuToPx(attr(extent, "cy"), 320));
  const data = `data:${mimeForPath(normalized)};base64,${uint8ToBase64(bytes)}`;
  return `<img src="${data}" alt="Document image" style="max-width:100%;width:${width}px;height:auto" data-px-height="${height}" />`;
}

async function runToHtml(run: Element, zip: JSZip, relationships: RelationshipMap) {
  const image = await imageHtmlFromRun(run, zip, relationships);
  if (image) return image;
  const rPr = firstDirect(run, `${WORD_NS}rPr`);
  const styles: string[] = [];
  const color = safeHex(attr(firstDirect(rPr, `${WORD_NS}color`), `${WORD_NS}val`));
  const sizeHalfPoints = Number(attr(firstDirect(rPr, `${WORD_NS}sz`), `${WORD_NS}val`));
  const fontElement = firstDirect(rPr, `${WORD_NS}rFonts`);
  const font = attr(fontElement, `${WORD_NS}ascii`) ?? attr(fontElement, `${WORD_NS}hAnsi`);
  const highlight = attr(firstDirect(rPr, `${WORD_NS}highlight`), `${WORD_NS}val`);
  const vertAlign = attr(firstDirect(rPr, `${WORD_NS}vertAlign`), `${WORD_NS}val`);
  styles.push(`color:#${color}`);
  if (Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0) styles.push(`font-size:${sizeHalfPoints / 2}pt`);
  if (font) styles.push(`font-family:${JSON.stringify(font)}`);
  if (firstDirect(rPr, `${WORD_NS}b`)) styles.push("font-weight:700");
  if (firstDirect(rPr, `${WORD_NS}i`)) styles.push("font-style:italic");
  if (firstDirect(rPr, `${WORD_NS}u`)) styles.push("text-decoration:underline");
  if (firstDirect(rPr, `${WORD_NS}strike`)) styles.push("text-decoration:line-through");
  if (highlight && highlight !== "none") styles.push(`background-color:${highlight}`);
  if (vertAlign === "superscript") styles.push("vertical-align:super;font-size:.78em");
  if (vertAlign === "subscript") styles.push("vertical-align:sub;font-size:.78em");
  let content = "";
  for (const child of Array.from(run.children)) {
    if (child.tagName === `${WORD_NS}t`) content += escapeHtml(child.textContent ?? "");
    else if (child.tagName === `${WORD_NS}tab`) content += "&emsp;";
    else if (child.tagName === `${WORD_NS}br`) content += attr(child, `${WORD_NS}type`) === "page" ? '<span class="px-page-break-inline"></span>' : "<br>";
  }
  if (!content) return "";
  return `<span style="${styles.join(";")}">${content}</span>`;
}

async function paragraphToHtml(paragraph: Element, zip: JSZip, relationships: RelationshipMap) {
  const pPr = firstDirect(paragraph, `${WORD_NS}pPr`);
  const align = attr(firstDirect(pPr, `${WORD_NS}jc`), `${WORD_NS}val`);
  const indent = Number(attr(firstDirect(pPr, `${WORD_NS}ind`), `${WORD_NS}left`));
  const pStyle = attr(firstDirect(pPr, `${WORD_NS}pStyle`), `${WORD_NS}val`) ?? "";
  const styles: string[] = [];
  if (align === "center") styles.push("text-align:center");
  else if (align === "right") styles.push("text-align:right");
  else if (align === "both") styles.push("text-align:justify");
  if (Number.isFinite(indent) && indent > 0) styles.push(`margin-left:${Math.round(indent / 20)}pt`);
  let content = "";
  for (const child of Array.from(paragraph.children)) {
    if (child.tagName === `${WORD_NS}r`) content += await runToHtml(child, zip, relationships);
    else if (child.tagName === `${WORD_NS}hyperlink`) for (const run of childrenByTag(child, `${WORD_NS}r`)) content += await runToHtml(run, zip, relationships);
  }
  const headingMatch = pStyle.match(/Heading\s*([1-6])/i) ?? pStyle.match(/Heading([1-6])/i);
  const tag = headingMatch ? `h${headingMatch[1]}` : "p";
  return `<${tag} style="${styles.join(";")}">${content || "<br>"}</${tag}>`;
}

async function tableToHtml(table: Element, zip: JSZip, relationships: RelationshipMap) {
  const rows: string[] = [];
  for (const row of childrenByTag(table, `${WORD_NS}tr`)) {
    const cells: string[] = [];
    for (const cell of childrenByTag(row, `${WORD_NS}tc`)) {
      const tcPr = firstDirect(cell, `${WORD_NS}tcPr`);
      const fill = attr(firstDirect(tcPr, `${WORD_NS}shd`), `${WORD_NS}fill`);
      const style = fill && fill !== "auto" && fill !== "FFFFFF" ? `background:#${safeHex(fill, "FFFFFF")};` : "";
      let cellHtml = "";
      for (const child of Array.from(cell.children)) {
        if (child.tagName === `${WORD_NS}p`) cellHtml += await paragraphToHtml(child, zip, relationships);
        else if (child.tagName === `${WORD_NS}tbl`) cellHtml += await tableToHtml(child, zip, relationships);
      }
      cells.push(`<td style="${style}">${cellHtml || "<p><br></p>"}</td>`);
    }
    rows.push(`<tr>${cells.join("")}</tr>`);
  }
  return `<table><tbody>${rows.join("")}</tbody></table>`;
}

export async function docxBlobToEditorHtml(blob: Blob) {
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentText = await zip.file("word/document.xml")?.async("text");
  if (!documentText) throw new Error("PX Editor could not read this DOCX.");
  const relText = await zip.file("word/_rels/document.xml.rels")?.async("text");
  const parser = new DOMParser();
  const documentXml = parser.parseFromString(documentText, "application/xml");
  const relXml = relText ? parser.parseFromString(relText, "application/xml") : parser.parseFromString("<Relationships />", "application/xml");
  const relationships = parseRelationships(relXml);
  const body = firstDescendant(documentXml.documentElement, `${WORD_NS}body`);
  if (!body) throw new Error("PX Editor could not find the Word document body.");
  const output: string[] = [];
  for (const child of Array.from(body.children)) {
    if (child.tagName === `${WORD_NS}p`) output.push(await paragraphToHtml(child, zip, relationships));
    else if (child.tagName === `${WORD_NS}tbl`) output.push(await tableToHtml(child, zip, relationships));
  }
  return output.join("\n") || "<p><br></p>";
}

function cssColorToHex(value: string | null | undefined) {
  if (!value) return undefined;
  const hex = value.trim().match(/^#([0-9a-f]{6})$/i);
  if (hex) return hex[1].toUpperCase();
  const rgb = value.trim().match(/^rgba?\((\d+)[, ]+\s*(\d+)[, ]+\s*(\d+)/i);
  if (!rgb) return undefined;
  return [rgb[1], rgb[2], rgb[3]].map((part) => Math.max(0, Math.min(255, Number(part))).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function cssSizeToHalfPoints(value: string | null | undefined) {
  if (!value) return undefined;
  const number = Number.parseFloat(value);
  if (!Number.isFinite(number)) return undefined;
  if (value.includes("pt")) return Math.round(number * 2);
  if (value.includes("px")) return Math.round(number * 0.75 * 2);
  return undefined;
}

function fontTagSizeToHalfPoints(value: string | null) {
  const index = Number(value);
  const points = [8, 10, 12, 14, 18, 24, 36];
  if (!Number.isInteger(index) || index < 1 || index > 7) return undefined;
  return points[index - 1] * 2;
}

function mergeRunState(base: RunState, element: Element) {
  const next: RunState = { ...base };
  const tag = element.tagName.toLowerCase();
  const htmlElement = element as HTMLElement;
  if (tag === "b" || tag === "strong") next.bold = true;
  if (tag === "i" || tag === "em") next.italics = true;
  if (tag === "u") next.underline = true;
  if (tag === "s" || tag === "strike") next.strike = true;
  if (tag === "sup") next.superScript = true;
  if (tag === "sub") next.subScript = true;
  const weight = htmlElement.style.fontWeight;
  if (weight === "bold" || Number(weight) >= 600) next.bold = true;
  if (htmlElement.style.fontStyle === "italic") next.italics = true;
  const decoration = htmlElement.style.textDecoration;
  if (decoration.includes("underline")) next.underline = true;
  if (decoration.includes("line-through")) next.strike = true;
  const color = cssColorToHex(htmlElement.style.color || element.getAttribute("color"));
  if (color) next.color = color;
  const size = cssSizeToHalfPoints(htmlElement.style.fontSize) ?? (tag === "font" ? fontTagSizeToHalfPoints(element.getAttribute("size")) : undefined);
  if (size) next.size = size;
  const font = htmlElement.style.fontFamily || element.getAttribute("face");
  if (font) next.font = font.replace(/["']/g, "").split(",")[0].trim();
  return next;
}

function imageRunFromElement(element: Element) {
  const src = element.getAttribute("src") ?? "";
  const match = src.match(/^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i);
  if (!match) return null;
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  const htmlImage = element as HTMLImageElement;
  const width = Math.max(40, Math.min(720, Number.parseFloat(htmlImage.style.width) || Number(htmlImage.getAttribute("width")) || 520));
  const heightHint = Number(element.getAttribute("data-px-height"));
  const height = Math.max(30, Math.min(980, Number.parseFloat(htmlImage.style.height) || Number(htmlImage.getAttribute("height")) || heightHint || Math.round(width * 0.65)));
  const ext = match[1].toLowerCase() === "jpeg" ? "jpg" : match[1].toLowerCase();
  return new ImageRun({ data: bytes, transformation: { width, height }, type: ext as "png" | "jpg" | "gif" | "bmp" });
}

function inlineRuns(node: Node, inherited: RunState = {}): Array<TextRun | ImageRun> {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? "";
    if (!text) return [];
    return [new TextRun({ text, bold: inherited.bold, italics: inherited.italics, underline: inherited.underline ? {} : undefined, strike: inherited.strike, color: inherited.color, size: inherited.size, font: inherited.font, superScript: inherited.superScript, subScript: inherited.subScript })];
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return [];
  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  if (tag === "br") return [new TextRun({ break: 1 })];
  if (tag === "img") {
    const image = imageRunFromElement(element);
    return image ? [image] : [];
  }
  const state = mergeRunState(inherited, element);
  return Array.from(element.childNodes).flatMap((child) => inlineRuns(child, state));
}

function alignmentFromElement(element: Element) {
  const htmlElement = element as HTMLElement;
  const align = htmlElement.style.textAlign || element.getAttribute("align") || "left";
  if (align === "center") return AlignmentType.CENTER;
  if (align === "right") return AlignmentType.RIGHT;
  if (align === "justify") return AlignmentType.JUSTIFIED;
  return AlignmentType.LEFT;
}

function paragraphFromElement(element: Element, prefix = "") {
  const tag = element.tagName.toLowerCase();
  const runs = inlineRuns(element);
  if (prefix) runs.unshift(new TextRun({ text: prefix }));
  const marginLeft = Number.parseFloat((element as HTMLElement).style.marginLeft || "0");
  const heading = tag === "h1" ? HeadingLevel.HEADING_1 : tag === "h2" ? HeadingLevel.HEADING_2 : tag === "h3" ? HeadingLevel.HEADING_3 : undefined;
  return new Paragraph({ children: runs.length ? runs : [new TextRun("")], heading, alignment: alignmentFromElement(element), indent: marginLeft > 0 ? { left: Math.round(marginLeft * 20) } : undefined, spacing: { after: 90, line: 276 } });
}

function tableFromElement(element: Element) {
  const rows: TableRow[] = [];
  const rowElements = Array.from(element.querySelectorAll(":scope > tbody > tr, :scope > thead > tr, :scope > tr"));
  for (const row of rowElements) {
    const cells = Array.from(row.children).filter((cell) => ["td", "th"].includes(cell.tagName.toLowerCase())).map((cell) => {
      const paragraphs = Array.from(cell.children).filter((child) => ["p", "div", "h1", "h2", "h3", "h4", "h5", "h6"].includes(child.tagName.toLowerCase())).map((child) => paragraphFromElement(child));
      if (!paragraphs.length) paragraphs.push(new Paragraph({ children: inlineRuns(cell) }));
      return new TableCell({ children: paragraphs });
    });
    if (cells.length) rows.push(new TableRow({ children: cells }));
  }
  return new Table({
    rows: rows.length ? rows : [new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun("")] })] })] })],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "B8C0D0" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "B8C0D0" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "B8C0D0" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "B8C0D0" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: "D6DBE5" },
      insideVertical: { style: BorderStyle.SINGLE, size: 3, color: "D6DBE5" },
    },
  });
}

function htmlToDocChildren(html: string) {
  const parsed = new DOMParser().parseFromString(html, "text/html");
  const children: EditorChild[] = [];
  let orderedIndex = 1;
  for (const node of Array.from(parsed.body.children)) {
    const tag = node.tagName.toLowerCase();
    if (node.classList.contains("px-page-break") || node.querySelector(".px-page-break-inline")) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      if (!node.classList.contains("px-page-break")) children.push(paragraphFromElement(node));
      continue;
    }
    if (tag === "table") {
      children.push(tableFromElement(node));
      continue;
    }
    if (tag === "ul" || tag === "ol") {
      orderedIndex = 1;
      for (const li of Array.from(node.children).filter((child) => child.tagName.toLowerCase() === "li")) children.push(paragraphFromElement(li, tag === "ol" ? `${orderedIndex++}. ` : "• "));
      continue;
    }
    if (["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote"].includes(tag)) {
      children.push(paragraphFromElement(node));
      continue;
    }
    if (tag === "img") children.push(new Paragraph({ children: inlineRuns(node), alignment: AlignmentType.CENTER }));
  }
  return children;
}

export async function editorHtmlToDocx(html: string, originalFileName: string) {
  const children = htmlToDocChildren(html);
  if (!children.length) children.push(new Paragraph({ children: [new TextRun("")] }));
  const doc = new DocxDocument({
    creator: "PX Editor — by pxcheco",
    title: originalFileName,
    description: "Edited locally in PX Editor.",
    styles: { default: { document: { run: { font: "Aptos", size: 22, color: "111827" }, paragraph: { spacing: { after: 90, line: 276 } } } } },
    sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }],
  });
  return Packer.toBlob(doc);
}

export function editedOutputName(fileName: string) {
  const base = fileName.replace(/\.docx$/i, "");
  return `${base}_PX_Edited.docx`;
}
