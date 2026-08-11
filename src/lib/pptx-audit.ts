import JSZip from "jszip";
import type {
  AuditFinding,
  FontInventoryItem,
  PictureInventoryItem,
  PptxAudit,
  SlideInventoryItem,
  TableInventoryItem,
  TemplateClassification,
} from "../types";
import { sha256Text } from "./hash";

const MAX_PACKAGE_FILES = 25_000;
const MAX_EXPANDED_BYTES = 750 * 1024 * 1024;
const TEXT_RUN_RE = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
const TYPEFACE_RE = /\btypeface=(?:"([^"]*)"|'([^']*)')/g;
const XML_ENTITY_RE = /&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi;
const SYMBOL_FONT_RE = /^(symbol|wingdings(?:\s*[23])?|webdings|cambria math|stix|mt extra)$/i;

function decodeXml(value: string): string {
  return value.replace(XML_ENTITY_RE, (entity, code: string) => {
    if (code === "amp") return "&";
    if (code === "lt") return "<";
    if (code === "gt") return ">";
    if (code === "quot") return '"';
    if (code === "apos") return "'";
    const radix = code.toLowerCase().startsWith("#x") ? 16 : 10;
    const numeric = Number.parseInt(code.replace(/^#x?/i, ""), radix);
    return Number.isFinite(numeric) ? String.fromCodePoint(numeric) : entity;
  });
}

function normalizeFont(value: string): string {
  return decodeXml(value).replace(/\s+/g, " ").trim().toLowerCase();
}

function partKind(path: string): string {
  if (/ppt\/slides\/slide\d+\.xml$/i.test(path)) return "slide";
  if (path.includes("/slideMasters/")) return "master";
  if (path.includes("/slideLayouts/")) return "layout";
  if (path.includes("/theme/")) return "theme";
  if (path.includes("/notesSlides/")) return "notes";
  return "other";
}

function slideNumberForPart(path: string): number | undefined {
  const match = path.match(/ppt\/slides\/slide(\d+)\.xml$/i);
  return match ? Number(match[1]) : undefined;
}

function extractTextRuns(xml: string): string[] {
  return [...xml.matchAll(TEXT_RUN_RE)].map((match) => decodeXml(match[1] ?? ""));
}

function extractFonts(xml: string): string[] {
  return [...xml.matchAll(TYPEFACE_RE)]
    .map((match) => decodeXml(match[1] ?? match[2] ?? "").trim())
    .filter((font) => font && !font.startsWith("+") && !/^none$/i.test(font));
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))].sort((left, right) => left.localeCompare(right));
}

function attributeValue(attributes: string, name: string): string | undefined {
  const match = attributes.match(new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "i"));
  return match ? decodeXml(match[1] ?? match[2] ?? "") : undefined;
}

async function extractTableInventory(slideNumber: number, xml: string): Promise<TableInventoryItem[]> {
  const blocks = [...xml.matchAll(/<a:tbl\b[\s\S]*?<\/a:tbl>/g)].map((match) => match[0]);
  const tables: TableInventoryItem[] = [];
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const properties = block.match(/<a:tblPr\b([^>]*)>/)?.[1] ?? "";
    const styleId = block.match(/<a:tableStyleId>([\s\S]*?)<\/a:tableStyleId>/)?.[1];
    const styleFlags = uniqueSorted(["firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"].filter((name) => attributeValue(properties, name) === "1"));
    const cellFonts = uniqueSorted(extractFonts(block));
    const colorTokens = uniqueSorted([...block.matchAll(/<a:(?:srgbClr|schemeClr|sysClr)\b[^>]*\bval=(?:"([^"]+)"|'([^']+)')/g)].map((match) => (match[1] ?? match[2] ?? "").toLowerCase()));
    const marginSignatures = uniqueSorted([...block.matchAll(/<a:tcPr\b([^>]*)>/g)].map((match) => {
      const attrs = match[1] ?? "";
      return ["marL", "marR", "marT", "marB", "anchor"].map((name) => `${name}:${attributeValue(attrs, name) ?? "default"}`).join("|");
    }));
    const styleFingerprint = await sha256Text(JSON.stringify({ styleId: styleId ? decodeXml(styleId).trim() : null, styleFlags, cellFonts: cellFonts.map(normalizeFont), colorTokens, marginSignatures }));
    tables.push({
      id: `slide-${slideNumber}-table-${index + 1}`,
      slideNumber,
      ordinal: index + 1,
      rowCount: xmlCount(block, /<a:tr\b/g),
      columnCount: xmlCount(block.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/)?.[0] ?? "", /<a:gridCol\b/g),
      mergedCellCount: xmlCount(block, /\b(?:gridSpan|rowSpan|hMerge|vMerge)=/g),
      styleId: styleId ? decodeXml(styleId).trim() : undefined,
      styleFlags,
      cellFonts,
      colorTokens,
      marginSignatures,
      styleFingerprint,
    });
  }
  return tables;
}

function extractPictureInventory(slideNumber: number, xml: string): PictureInventoryItem[] {
  const blocks = [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map((match) => match[0]);
  return blocks.map((block, index) => {
    const nonVisual = block.match(/<p:cNvPr\b([^>]*)>/)?.[1] ?? "";
    const extent = block.match(/<a:ext\b([^>]*)\/>/)?.[1] ?? "";
    const relationshipMatch = block.match(/<a:blip\b[^>]*\br:embed=(?:"([^"]+)"|'([^']+)')/);
    const description = attributeValue(nonVisual, "descr")?.trim();
    return {
      id: `slide-${slideNumber}-picture-${index + 1}`,
      slideNumber,
      ordinal: index + 1,
      name: attributeValue(nonVisual, "name")?.trim() || `Picture ${index + 1}`,
      description: description || undefined,
      relationshipId: relationshipMatch ? relationshipMatch[1] ?? relationshipMatch[2] : undefined,
      widthEmu: Number(attributeValue(extent, "cx")) || undefined,
      heightEmu: Number(attributeValue(extent, "cy")) || undefined,
      cropped: /<a:srcRect\b[^>]*(?:\bl=|\br=|\bt=|\bb=)/.test(block),
      hasOutline: /<a:ln\b/.test(block) && !/<a:ln\b[\s\S]*?<a:noFill\s*\/>[\s\S]*?<\/a:ln>/.test(block),
      hasEffect: /<a:(?:effectLst|effectDag)\b/.test(block),
    };
  });
}

function xmlCount(xml: string, expression: RegExp): number {
  return [...xml.matchAll(expression)].length;
}

function classifyTemplate(allText: string, fonts: FontInventoryItem[]): { classification: TemplateClassification; evidence: string[] } {
  const normalized = allText.toLowerCase();
  const fontNames = new Set(fonts.map((font) => font.normalizedFamily));
  const evidence: string[] = [];
  const hasOrnlLanguage = normalized.includes("oak ridge national laboratory") || /\bornl\b/.test(normalized);
  const hasAptos = fontNames.has("aptos") || fontNames.has("aptos display");
  const hasLegacyOrnlFont = fontNames.has("century gothic");

  if (hasOrnlLanguage) evidence.push("The package contains ORNL or Oak Ridge National Laboratory text.");
  if (hasAptos) evidence.push("The package contains Aptos typography.");
  if (hasLegacyOrnlFont) evidence.push("The package contains Century Gothic legacy typography.");

  if (hasOrnlLanguage) {
    evidence.push("A template hash is not installed, so the current official revision cannot be proven automatically.");
    return { classification: "older-or-modified-ornl", evidence };
  }
  evidence.push("No authoritative template identity was found; a person must classify the deck before cleanup.");
  return { classification: "unknown", evidence };
}

function finding(input: Omit<AuditFinding, "id">): AuditFinding {
  const scope = `${input.ruleId}:${input.slideNumber ?? "deck"}:${input.evidence}`;
  let hash = 2166136261;
  for (let index = 0; index < scope.length; index += 1) {
    hash ^= scope.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return { ...input, id: `finding-${(hash >>> 0).toString(16).padStart(8, "0")}` };
}

export async function auditPptx(bytes: Uint8Array): Promise<PptxAudit> {
  if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("This file is not a ZIP-based PowerPoint presentation.");
  }
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: false, createFolders: false });
  const paths = Object.keys(zip.files).filter((path) => !zip.files[path].dir);
  if (paths.length > MAX_PACKAGE_FILES) throw new Error(`The PowerPoint contains ${paths.length} package parts, above the ${MAX_PACKAGE_FILES.toLocaleString()}-part safety limit.`);
  if (paths.some((path) => path.startsWith("/") || path.includes("\\") || path.split("/").includes(".."))) {
    throw new Error("The PowerPoint contains an unsafe package path.");
  }

  const xmlByPath = new Map<string, string>();
  let expandedByteLength = 0;
  for (const path of paths) {
    const entry = zip.file(path) as ({ _data?: { uncompressedSize?: number; compressedSize?: number } }) | null;
    const uncompressed = entry?._data?.uncompressedSize;
    const compressed = entry?._data?.compressedSize;
    if (Number.isFinite(uncompressed)) {
      expandedByteLength += uncompressed ?? 0;
      if ((uncompressed ?? 0) > 50 * 1024 * 1024 && (compressed ?? 0) > 0 && (uncompressed ?? 0) / (compressed ?? 1) > 1_000) {
        throw new Error(`The PowerPoint package part ${path} has an unsafe compression ratio.`);
      }
    }
  }
  if (expandedByteLength > MAX_EXPANDED_BYTES) throw new Error("The expanded PowerPoint package exceeds the 750 MB safety limit.");
  const hasDeclaredSizes = expandedByteLength > 0;
  if (!hasDeclaredSizes) expandedByteLength = 0;
  for (const path of paths) {
    const entry = zip.file(path);
    if (!entry) continue;
    if (/\.(xml|rels)$/i.test(path) || path === "[Content_Types].xml") {
      const text = await entry.async("text");
      if (!hasDeclaredSizes) expandedByteLength += new TextEncoder().encode(text).byteLength;
      xmlByPath.set(path, text);
    } else if (!hasDeclaredSizes) {
      const data = await entry.async("uint8array");
      expandedByteLength += data.byteLength;
    }
    if (expandedByteLength > MAX_EXPANDED_BYTES) {
      throw new Error("The expanded PowerPoint package exceeds the 750 MB safety limit.");
    }
  }

  const slidePaths = paths
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort((left, right) => (slideNumberForPart(left) ?? 0) - (slideNumberForPart(right) ?? 0));
  const fontMap = new Map<string, FontInventoryItem>();
  let searchableText = "";

  for (const [path, xml] of xmlByPath.entries()) {
    searchableText += ` ${extractTextRuns(xml).join(" ")}`;
    for (const family of extractFonts(xml)) {
      const normalizedFamily = normalizeFont(family);
      const existing = fontMap.get(normalizedFamily) ?? {
        family,
        normalizedFamily,
        count: 0,
        directSlideCount: 0,
        slideNumbers: [],
        partKinds: [],
        isThemeFont: false,
        isLikelySymbolFont: SYMBOL_FONT_RE.test(family),
      };
      existing.count += 1;
      const kind = partKind(path);
      if (kind === "slide") existing.directSlideCount += 1;
      if (!existing.partKinds.includes(kind)) existing.partKinds.push(kind);
      if (kind === "theme") existing.isThemeFont = true;
      const slideNumber = slideNumberForPart(path);
      if (slideNumber && !existing.slideNumbers.includes(slideNumber)) existing.slideNumbers.push(slideNumber);
      fontMap.set(normalizedFamily, existing);
    }
  }
  const fonts = [...fontMap.values()].sort((left, right) => right.count - left.count || left.family.localeCompare(right.family));
  for (const font of fonts) font.slideNumbers.sort((left, right) => left - right);
  const classified = classifyTemplate(searchableText, fonts);

  const slides: SlideInventoryItem[] = [];
  const tables: TableInventoryItem[] = [];
  const pictures: PictureInventoryItem[] = [];
  for (const path of slidePaths) {
    const xml = xmlByPath.get(path) ?? "";
    const number = slideNumberForPart(path) ?? slides.length + 1;
    const runs = extractTextRuns(xml);
    const text = runs.join(" ").replace(/\s+/g, " ").trim();
    const title = runs.find((run) => run.trim().length > 0)?.trim().slice(0, 160) || `Slide ${number}`;
    const slideFonts = [...new Set(extractFonts(xml))].sort();
    const fontSizes = uniqueSorted([...xml.matchAll(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)].map((match) => String(Number(match[1] ?? match[2]) / 100))).map(Number);
    const warnings: string[] = [];
    const tableCount = xmlCount(xml, /<a:tbl\b/g);
    tables.push(...await extractTableInventory(number, xml));
    const pictureCount = xmlCount(xml, /<p:pic\b/g);
    pictures.push(...extractPictureInventory(number, xml));
    const connectorCount = xmlCount(xml, /<p:cxnSp\b/g);
    const relationXml = xmlByPath.get(`ppt/slides/_rels/slide${number}.xml.rels`) ?? "";
    const chartCount = xmlCount(relationXml, /relationships\/chart(?:"|\/)/gi);
    const commentPart = xmlByPath.get(`ppt/comments/comment${number}.xml`) ?? "";
    const commentCount = xmlCount(commentPart, /<p:cm\b/g);
    if (!text) warnings.push("No visible text was detected in the slide XML.");
    if (text.length > 1_400) warnings.push("Dense visible text should receive visual fit review.");
    slides.push({
      id: `slide-${number}`,
      number,
      title,
      text,
      textHash: await sha256Text(text),
      textRunCount: runs.length,
      tableCount,
      pictureCount,
      chartCount,
      connectorCount,
      commentCount,
      fonts: slideFonts,
      fontSizes,
      warnings,
    });
  }

  const contentTypes = xmlByPath.get("[Content_Types].xml") ?? "";
  const relationshipXml = [...xmlByPath.entries()]
    .filter(([path]) => path.endsWith(".rels"))
    .map(([, xml]) => xml)
    .join("\n");
  const containsMacros = /macroEnabled|vbaProject/i.test(contentTypes + relationshipXml) || paths.some((path) => /vbaProject\.bin$/i.test(path));
  const containsOleObjects = paths.some((path) => /^ppt\/embeddings\//i.test(path)) || /relationships\/oleObject/i.test(relationshipXml);
  const containsExternalRelationships = /TargetMode=(?:"|')External(?:"|')/i.test(relationshipXml);
  const legacyCommentCount = paths.filter((path) => /^ppt\/comments\/comment\d+\.xml$/i.test(path)).length;
  const modernCommentCount = paths.filter((path) => /ppt\/(?:comments|people|authors).*\.xml$/i.test(path) && !/^ppt\/comments\/comment\d+\.xml$/i.test(path)).length;
  const findings: AuditFinding[] = [];

  if (classified.classification !== "current-ornl") {
    findings.push(finding({
      ruleId: "template.confirm-target",
      category: "template",
      severity: "warning",
      confidence: "high",
      message: "Confirm the target template before staging cleanup.",
      evidence: classified.evidence.join(" "),
      autoFixable: false,
    }));
  }
  for (const font of fonts) {
    if (font.isLikelySymbolFont || font.directSlideCount === 0 || !["century gothic", "arial"].includes(font.normalizedFamily)) continue;
    findings.push(finding({
      ruleId: `font.legacy.${font.normalizedFamily.replaceAll(" ", "-")}`,
      category: "font",
      severity: "warning",
      confidence: font.normalizedFamily === "century gothic" ? "high" : "medium",
      message: `${font.family} appears ${font.directSlideCount} time${font.directSlideCount === 1 ? "" : "s"} in editable slide markup.`,
      evidence: font.slideNumbers.length ? `Detected on slide${font.slideNumbers.length === 1 ? "" : "s"} ${font.slideNumbers.join(", ")}.` : `Detected in ${font.partKinds.join(", ")} parts.`,
      autoFixable: true,
    }));
  }
  const tableStyleIds = uniqueSorted(tables.map((table) => table.styleId ?? "No table style ID"));
  if (tableStyleIds.length > 1) {
    findings.push(finding({
      ruleId: "table.style-clusters",
      category: "table",
      severity: "warning",
      confidence: "high",
      message: `${tables.length} native tables use ${tableStyleIds.length} table style IDs.`,
      evidence: tableStyleIds.join(", "),
      autoFixable: false,
    }));
  }
  const tableFingerprints = uniqueSorted(tables.map((table) => table.styleFingerprint));
  if (tables.length > 1 && tableFingerprints.length > tableStyleIds.length) {
    findings.push(finding({
      ruleId: "table.local-format-clusters",
      category: "table",
      severity: "info",
      confidence: "medium",
      message: `Native tables contain ${tableFingerprints.length} local formatting fingerprints.`,
      evidence: "Differences may reflect cell margins, direct fonts, semantic color, or other table-local treatment and require exemplar-aware review.",
      autoFixable: false,
    }));
  }
  const missingPictureDescriptions = pictures.filter((picture) => !picture.description);
  if (missingPictureDescriptions.length > 0) {
    const missingDescriptionSlides = [...new Set(missingPictureDescriptions.map((picture) => picture.slideNumber))].sort((left, right) => left - right);
    findings.push(finding({
      ruleId: "figure.missing-description",
      category: "figure",
      severity: "info",
      confidence: "high",
      message: `${missingPictureDescriptions.length} of ${pictures.length} native pictures have no stored description.`,
      evidence: `Review slide${missingDescriptionSlides.length === 1 ? "" : "s"} ${missingDescriptionSlides.join(", ")} for meaningful alt text or decorative status.`,
      autoFixable: false,
    }));
  }
  for (const slide of slides) {
    const minimum = slide.fontSizes.length ? Math.min(...slide.fontSizes) : undefined;
    if (minimum !== undefined && minimum < 10) {
      findings.push(finding({
        ruleId: "production.small-direct-type",
        category: "production",
        severity: "info",
        confidence: "high",
        slideNumber: slide.number,
        message: `Direct text as small as ${minimum} pt appears on this slide.`,
        evidence: "Small type may be intentional for citations, equations, or dense technical tables; confirm legibility in the final PowerPoint render.",
        autoFixable: false,
      }));
    }
  }
  if (containsMacros || containsOleObjects || containsExternalRelationships) {
    findings.push(finding({
      ruleId: "production.advanced-content",
      category: "production",
      severity: "error",
      confidence: "high",
      message: "Advanced or externally linked content requires manual review before cleanup export.",
      evidence: `Macros: ${containsMacros ? "yes" : "no"}; embedded OLE: ${containsOleObjects ? "yes" : "no"}; external relationships: ${containsExternalRelationships ? "yes" : "no"}.`,
      autoFixable: false,
    }));
  }
  for (const slide of slides.filter((item) => item.warnings.length > 0)) {
    findings.push(finding({
      ruleId: "layout.visual-review",
      category: "layout",
      severity: "info",
      confidence: "medium",
      slideNumber: slide.number,
      message: slide.warnings[0],
      evidence: `${slide.textRunCount} text runs; ${slide.text.length} visible characters.`,
      autoFixable: false,
    }));
  }

  return {
    scannedAt: new Date().toISOString(),
    supportLevel: containsMacros || containsOleObjects ? "partial" : "native-ooxml",
    slideCount: slides.length,
    masterCount: paths.filter((path) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(path)).length,
    layoutCount: paths.filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(path)).length,
    themeCount: paths.filter((path) => /^ppt\/theme\/theme\d+\.xml$/i.test(path)).length,
    notesCount: paths.filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path)).length,
    legacyCommentCount,
    modernCommentCount,
    mediaCount: paths.filter((path) => /^ppt\/media\//i.test(path)).length,
    tableCount: slides.reduce((sum, slide) => sum + slide.tableCount, 0),
    chartCount: slides.reduce((sum, slide) => sum + slide.chartCount, 0),
    pictureCount: slides.reduce((sum, slide) => sum + slide.pictureCount, 0),
    containsMacros,
    containsOleObjects,
    containsExternalRelationships,
    packageFileCount: paths.length,
    expandedByteLength,
    classification: classified.classification,
    classificationEvidence: classified.evidence,
    fonts,
    slides,
    tables,
    pictures,
    findings,
    warnings: modernCommentCount > 0 ? ["Modern PowerPoint comments were retained as unsupported package parts and do not block the audit."] : [],
  };
}

export function extractSlideText(xml: string): string {
  return extractTextRuns(xml).join(" ").replace(/\s+/g, " ").trim();
}
