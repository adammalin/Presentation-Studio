import JSZip from "jszip";
import type {
  ProjectResource,
  ResourceKind,
  ResourceRole,
  ResourceSupportState,
} from "../types";
import { sha256 } from "./hash";

export const MAX_SINGLE_RESOURCE_BYTES = 1_000_000_000;
export const MAX_PROJECT_RESOURCE_BYTES = 1_250_000_000;
const MAX_TEXT_INPUT_BYTES = 4 * 1024 * 1024;
const MAX_TEXT_DERIVATIVE_BYTES = 2 * 1024 * 1024;
const MAX_OFFICE_PACKAGE_ENTRIES = 20_000;
const MAX_OFFICE_EXPANDED_BYTES = 250 * 1024 * 1024;

export interface ResourceInput {
  name: string;
  filePath?: string;
  mediaType?: string;
  bytes: Uint8Array;
}
interface ResourceDescriptor {
  kind: ResourceKind;
  mediaType: string;
  roles: ResourceRole[];
  support: ResourceSupportState[];
  extraction?: "text" | "docx" | "xlsx";
  status: "indexed" | "stored-only" | "needs-review";
  summary: string;
  warnings: string[];
}

const blockedExtensions = new Set([
  "app", "bat", "cmd", "com", "cpl", "dll", "dmg", "exe", "jar", "js", "mjs", "cjs",
  "msi", "pkg", "ps1", "py", "scr", "sh", "vbs",
]);

const knownMediaTypes: Record<string, string> = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  md: "text/markdown",
  markdown: "text/markdown",
  txt: "text/plain",
  csv: "text/csv",
  tsv: "text/tab-separated-values",
  json: "application/json",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  wav: "audio/wav",
  mp3: "audio/mpeg",
  m4a: "audio/mp4",
  mp4: "video/mp4",
  mov: "video/quicktime",
  doc: "application/msword",
  xls: "application/vnd.ms-excel",
};

function extensionFor(name: string): string {
  const leaf = name.replaceAll("\\", "/").split("/").pop() ?? name;
  const dot = leaf.lastIndexOf(".");
  return dot > 0 ? leaf.slice(dot + 1).toLowerCase() : "";
}

function descriptorFor(name: string, suppliedMediaType?: string): ResourceDescriptor {
  const extension = extensionFor(name);
  const mediaType = knownMediaTypes[extension] ?? (suppliedMediaType?.trim() || "application/octet-stream");
  if (extension === "pptx") return { kind: "presentation", mediaType, roles: ["import-origin"], support: ["pptx-preserved", "source-readable"], status: "indexed", summary: "Embedded and queued for read-only PowerPoint audit.", warnings: [] };
  if (extension === "potx") return { kind: "presentation", mediaType, roles: ["template-source"], support: ["pptx-preserved"], status: "stored-only", summary: "Embedded as a portable PowerPoint template source.", warnings: ["Template compilation is not available in this build."] };
  if (["md", "markdown", "txt", "json"].includes(extension)) return { kind: "document", mediaType, roles: ["grounding-source"], support: ["source-readable"], extraction: "text", status: "indexed", summary: "Text extracted locally and stored with the project.", warnings: [] };
  if (["csv", "tsv"].includes(extension)) return { kind: "data", mediaType, roles: ["chart-data", "grounding-source"], support: ["source-readable"], extraction: "text", status: "indexed", summary: "Table text extracted locally and stored with the project.", warnings: [] };
  if (extension === "docx") return { kind: "document", mediaType, roles: ["grounding-source"], support: ["source-readable"], extraction: "docx", status: "indexed", summary: "Document text extracted locally and stored with the project.", warnings: [] };
  if (extension === "xlsx") return { kind: "data", mediaType, roles: ["chart-data", "grounding-source"], support: ["source-readable"], extraction: "xlsx", status: "indexed", summary: "Workbook values extracted locally and stored with the project.", warnings: [] };
  if (extension === "pdf") return { kind: "document", mediaType, roles: ["grounding-source"], support: ["previewable"], status: "stored-only", summary: "PDF embedded safely; text extraction is not available in this build.", warnings: ["Use the original PDF as a local reference until a bounded PDF extractor is enabled."] };
  if (["png", "jpg", "jpeg", "webp", "tif", "tiff", "svg"].includes(extension)) return { kind: "image", mediaType, roles: ["slide-media"], support: ["previewable", "placeable"], status: "indexed", summary: "Image indexed as embedded slide media.", warnings: [] };
  if (["wav", "mp3", "m4a"].includes(extension)) return { kind: "audio", mediaType, roles: ["slide-media"], support: ["previewable", "placeable"], status: "stored-only", summary: "Audio embedded as slide media; transcript extraction is not available.", warnings: [] };
  if (["mp4", "mov"].includes(extension)) return { kind: "video", mediaType, roles: ["slide-media"], support: ["previewable", "placeable"], status: "stored-only", summary: "Video embedded as slide media; poster-frame and transcript extraction are not available.", warnings: [] };
  if (extension === "doc") return { kind: "document", mediaType, roles: ["grounding-source"], support: ["unsupported"], status: "needs-review", summary: "Legacy Word file embedded without extraction.", warnings: ["Convert this file to DOCX for local text extraction."] };
  if (extension === "xls") return { kind: "data", mediaType, roles: ["chart-data"], support: ["unsupported"], status: "needs-review", summary: "Legacy Excel file embedded without extraction.", warnings: ["Convert this file to XLSX or CSV for local value extraction."] };
  return { kind: "other", mediaType, roles: ["reference-only"], support: ["unsupported"], status: "needs-review", summary: "File embedded as reference-only; no safe processor is registered for this type.", warnings: ["The file is stored but cannot yet be previewed, read, or placed by Presentation Studio."] };
}

function beginsWith(bytes: Uint8Array, signature: number[]): boolean {
  return signature.every((value, index) => bytes[index] === value);
}

function isZip(bytes: Uint8Array): boolean {
  return beginsWith(bytes, [0x50, 0x4b, 0x03, 0x04]) || beginsWith(bytes, [0x50, 0x4b, 0x05, 0x06]) || beginsWith(bytes, [0x50, 0x4b, 0x07, 0x08]);
}

function validateKnownSignature(extension: string, bytes: Uint8Array): string[] {
  const warnings: string[] = [];
  if (["pptx", "potx", "docx", "xlsx"].includes(extension) && !isZip(bytes)) throw new Error(`.${extension} file signature does not match its filename.`);
  if (extension === "pdf" && new TextDecoder("ascii").decode(bytes.slice(0, 5)) !== "%PDF-") throw new Error("PDF file signature does not match its filename.");
  if (extension === "png" && !beginsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) throw new Error("PNG file signature does not match its filename.");
  if (["jpg", "jpeg"].includes(extension) && !beginsWith(bytes, [0xff, 0xd8, 0xff])) throw new Error("JPEG file signature does not match its filename.");
  if (extension === "webp") {
    const riff = new TextDecoder("ascii").decode(bytes.slice(0, 4));
    const webp = new TextDecoder("ascii").decode(bytes.slice(8, 12));
    if (riff !== "RIFF" || webp !== "WEBP") throw new Error("WebP file signature does not match its filename.");
  }
  if (extension === "wav") {
    const riff = new TextDecoder("ascii").decode(bytes.slice(0, 4));
    const wave = new TextDecoder("ascii").decode(bytes.slice(8, 12));
    if (riff !== "RIFF" || wave !== "WAVE") throw new Error("WAV file signature does not match its filename.");
  }
  if (["mp4", "m4a", "mov"].includes(extension) && new TextDecoder("ascii").decode(bytes.slice(4, 8)) !== "ftyp") throw new Error(`.${extension} file signature does not match its filename.`);
  if (extension === "svg") {
    const source = new TextDecoder().decode(bytes.slice(0, Math.min(bytes.byteLength, 256 * 1024)));
    if (!/<svg(?:\s|>)/i.test(source)) throw new Error("SVG markup does not contain an SVG root element.");
    if (/<script(?:\s|>)/i.test(source) || /\son\w+\s*=/i.test(source) || /(?:href|src)\s*=\s*["']\s*(?:https?:|javascript:|data:text\/html)/i.test(source)) {
      warnings.push("Active or externally linked SVG content was detected. The SVG is stored but will not be previewed or placed until sanitized.");
    }
  }
  return warnings;
}

function decodeXmlEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&")
    .replace(/&#(\d+);/g, (_match, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_match, code) => String.fromCodePoint(Number.parseInt(code, 16)));
}

function normalizedText(value: string): string {
  return value
    .replaceAll("\u0000", "")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function xmlText(xml: string): string {
  return normalizedText(decodeXmlEntities(xml
    .replace(/<w:tab\b[^>]*\/?\s*>/gi, "\t")
    .replace(/<w:br\b[^>]*\/?\s*>/gi, "\n")
    .replace(/<\/w:p\s*>/gi, "\n")
    .replace(/<\/w:tr\s*>/gi, "\n")
    .replace(/<\/w:tc\s*>/gi, "\t")
    .replace(/<[^>]+>/g, "")));
}

function assertBoundedOfficePackage(zip: JSZip): void {
  const entries = Object.values(zip.files);
  if (entries.length > MAX_OFFICE_PACKAGE_ENTRIES) throw new Error("The Office file contains too many package entries.");
  let expanded = 0;
  for (const entry of entries) {
    const size = (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
    if (typeof size === "number" && Number.isFinite(size)) expanded += size;
  }
  if (expanded > MAX_OFFICE_EXPANDED_BYTES) throw new Error("The Office file expands beyond the 250 MB processing limit.");
}

async function extractDocxText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  assertBoundedOfficePackage(zip);
  const paths = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((path) => /^word\/(?:header|footer|footnotes|endnotes)\d*\.xml$/i.test(path)).sort(),
  ];
  const sections: string[] = [];
  for (const path of paths) {
    const entry = zip.file(path);
    if (!entry) continue;
    const text = xmlText(await entry.async("string"));
    if (text) sections.push(text);
  }
  return normalizedText(sections.join("\n\n"));
}

function worksheetText(xml: string, sharedStrings: string[]): string {
  const rows: string[] = [];
  for (const rowMatch of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/gi)) {
    const values: string[] = [];
    for (const cellMatch of rowMatch[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/gi)) {
      const attributes = cellMatch[1];
      const body = cellMatch[2];
      const type = /\bt=["']([^"']+)["']/i.exec(attributes)?.[1];
      const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/i.exec(body)?.[1] ?? /<t\b[^>]*>([\s\S]*?)<\/t>/i.exec(body)?.[1] ?? "";
      const decoded = decodeXmlEntities(raw);
      values.push(type === "s" ? (sharedStrings[Number(decoded)] ?? decoded) : decoded);
    }
    if (values.some(Boolean)) rows.push(values.join("\t"));
  }
  return rows.join("\n");
}

async function extractXlsxText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  assertBoundedOfficePackage(zip);
  const sharedEntry = zip.file("xl/sharedStrings.xml");
  const sharedXml = sharedEntry ? await sharedEntry.async("string") : "";
  const sharedStrings = [...sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)].map((match) => xmlText(match[1]));
  const sheetPaths = Object.keys(zip.files).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const sheets: string[] = [];
  for (let index = 0; index < sheetPaths.length; index += 1) {
    const entry = zip.file(sheetPaths[index]);
    if (!entry) continue;
    const text = worksheetText(await entry.async("string"), sharedStrings);
    if (text) sheets.push(`[Sheet ${index + 1}]\n${text}`);
  }
  return normalizedText(sheets.join("\n\n"));
}

function extractPlainText(bytes: Uint8Array): { text: string; truncated: boolean } {
  const sample = bytes.slice(0, Math.min(bytes.byteLength, MAX_TEXT_INPUT_BYTES));
  const nullCount = sample.slice(0, Math.min(sample.byteLength, 4096)).reduce((count, value) => count + (value === 0 ? 1 : 0), 0);
  if (nullCount > 2) throw new Error("The file appears to contain binary data rather than readable text.");
  return { text: normalizedText(new TextDecoder().decode(sample)), truncated: bytes.byteLength > sample.byteLength };
}

async function createTextDerivative(text: string, initiallyTruncated: boolean, processor: string) {
  let bytes = new TextEncoder().encode(text);
  let truncated = initiallyTruncated;
  if (bytes.byteLength > MAX_TEXT_DERIVATIVE_BYTES) {
    bytes = bytes.slice(0, MAX_TEXT_DERIVATIVE_BYTES);
    const safeText = new TextDecoder().decode(bytes).replace(/\uFFFD$/, "");
    bytes = new TextEncoder().encode(`${safeText}\n\n[Extraction truncated by Presentation Studio]`);
    truncated = true;
  }
  return {
    id: crypto.randomUUID(),
    kind: "extracted-text" as const,
    mediaType: "text/plain" as const,
    byteLength: bytes.byteLength,
    sha256: await sha256(bytes),
    createdAt: new Date().toISOString(),
    processor,
    truncated,
    bytes,
  };
}

export function isPowerPointResource(name: string): boolean {
  return extensionFor(name) === "pptx";
}

export async function processResourceInput(input: ResourceInput): Promise<ProjectResource> {
  const extension = extensionFor(input.name);
  if (blockedExtensions.has(extension)) throw new Error(`${input.name} is an active program or script type and cannot be added as a project Resource.`);
  if (input.bytes.byteLength > MAX_SINGLE_RESOURCE_BYTES) throw new Error(`${input.name} exceeds the 1 GB per-Resource limit.`);
  const descriptor = descriptorFor(input.name, input.mediaType);
  const signatureWarnings = validateKnownSignature(extension, input.bytes);
  descriptor.warnings.push(...signatureWarnings);
  if (signatureWarnings.length > 0) {
    descriptor.status = "needs-review";
    descriptor.support = ["unsupported"];
    descriptor.summary = "File embedded but held for review because active or external content was detected.";
  }

  const derivatives = [];
  if (descriptor.extraction) {
    try {
      let text = "";
      let truncated = false;
      let processor = "presentation-studio/text-v1";
      if (descriptor.extraction === "text") ({ text, truncated } = extractPlainText(input.bytes));
      if (descriptor.extraction === "docx") {
        text = await extractDocxText(input.bytes);
        processor = "presentation-studio/docx-v1";
      }
      if (descriptor.extraction === "xlsx") {
        text = await extractXlsxText(input.bytes);
        processor = "presentation-studio/xlsx-v1";
      }
      if (text) derivatives.push(await createTextDerivative(text, truncated, processor));
      else descriptor.warnings.push("The local extractor found no readable text or values.");
      if (truncated) descriptor.warnings.push("The extracted text derivative was truncated to the local processing limit.");
    } catch (error) {
      descriptor.status = "needs-review";
      descriptor.support = ["unsupported"];
      descriptor.summary = "Original file embedded, but local extraction did not complete.";
      descriptor.warnings.push(error instanceof Error ? error.message : "The local extractor failed.");
    }
  }

  return {
    id: crypto.randomUUID(),
    name: input.name.replaceAll("\\", "/").split("/").pop() || "Untitled Resource",
    mediaType: descriptor.mediaType,
    byteLength: input.bytes.byteLength,
    sha256: await sha256(input.bytes),
    roles: descriptor.roles,
    kind: descriptor.kind,
    support: descriptor.support,
    processing: {
      status: descriptor.status,
      summary: descriptor.summary,
      processedAt: new Date().toISOString(),
      warnings: descriptor.warnings,
    },
    derivatives,
    createdAt: new Date().toISOString(),
    sourcePath: input.filePath,
    embedded: true,
    bytes: input.bytes,
    mcpAccess: "none",
  };
}
