import JSZip from "jszip";
import type {
  AlignmentRepairCandidate,
  AuditFinding,
  FontInventoryItem,
  LayoutReviewItem,
  PictureInventoryItem,
  PptxAudit,
  SlideEditableObject,
  SlideEditableObjectElement,
  SlideInventoryItem,
  TableInventoryItem,
  TemplateClassification,
  TextBoxInventoryItem,
  TextParagraphInventoryItem,
} from "../types";
import { normalizeCellFillToken, semanticColorRoleForToken } from "./semantic-visuals";

export const PPTX_AUDIT_SEMANTIC_VISUAL_VERSION = 6;
import { sha256Text } from "./hash";

const MAX_PACKAGE_FILES = 25_000;
const MAX_EXPANDED_BYTES = 750 * 1024 * 1024;
const TEXT_RUN_RE = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
const TYPEFACE_RE = /\btypeface=(?:"([^"]*)"|'([^']*)')/g;
const XML_ENTITY_RE = /&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi;
const SYMBOL_FONT_RE = /^(symbol|wingdings(?:\s*[23])?|webdings|cambria math|stix|mt extra)$/i;
const EMU_PER_POINT = 12_700;
const EMU_PER_INCH = 914_400;
const DEFAULT_SLIDE_WIDTH_EMU = 12_192_000;
const DEFAULT_SLIDE_HEIGHT_EMU = 6_858_000;
const SAFE_MARGIN_EMU = Math.round(0.25 * EMU_PER_INCH);
const OFF_SLIDE_TOLERANCE_EMU = 2 * EMU_PER_POINT;

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

function extractTableCellRuns(xml: string): { textRuns: string[]; paragraphRunCounts: number[]; runBreaksBefore: Array<"none" | "line" | "paragraph"> } {
  const textRuns: string[] = [];
  const paragraphRunCounts: number[] = [];
  const runBreaksBefore: Array<"none" | "line" | "paragraph"> = [];
  const paragraphs = [...xml.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g)];
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    const paragraphXml = paragraphs[paragraphIndex][1] ?? "";
    let count = 0;
    let pendingBreak: "none" | "line" | "paragraph" = paragraphIndex > 0 ? "paragraph" : "none";
    const items = [...paragraphXml.matchAll(/<a:br\b[^>]*(?:\/>|>[\s\S]*?<\/a:br>)|<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)];
    for (const item of items) {
      if (item[0].startsWith("<a:br")) {
        pendingBreak = "line";
        continue;
      }
      textRuns.push(decodeXml(item[1] ?? ""));
      runBreaksBefore.push(pendingBreak);
      pendingBreak = "none";
      count += 1;
    }
    paragraphRunCounts.push(count);
  }
  return { textRuns, paragraphRunCounts, runBreaksBefore };
}

function visibleTextFromRuns(content: ReturnType<typeof extractTableCellRuns>): string {
  let text = "";
  for (let index = 0; index < content.textRuns.length; index += 1) {
    const value = content.textRuns[index] ?? "";
    const breakBefore = content.runBreaksBefore[index] ?? "none";
    if (breakBefore !== "none" && text && value && !/\s$/.test(text) && !/^\s/.test(value)) text += " ";
    text += value;
  }
  return text.replace(/\s+/g, " ").trim();
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

interface ExternalRelationshipInventory {
  totalCount: number;
  hyperlinkCount: number;
  blockingCount: number;
  blockingTypes: string[];
}

/**
 * Hyperlinks are normal presentation content. They are stored as external
 * relationships, but surgical slide edits do not dereference or rewrite them.
 * Linked media, OLE/package links, external data, and every unknown external
 * relationship remain blocking because their payload can change outside the
 * project or require an adapter that Studio does not yet provide.
 */
function inventoryExternalRelationships(entries: Array<[string, string]>): ExternalRelationshipInventory {
  let totalCount = 0;
  let hyperlinkCount = 0;
  const blockingTypes: string[] = [];
  for (const [, xml] of entries) {
    for (const match of xml.matchAll(/<Relationship\b([^>]*?)\/?\s*>/g)) {
      const attributes = match[1] ?? "";
      if (!/^external$/i.test(attributeValue(attributes, "TargetMode") ?? "")) continue;
      totalCount += 1;
      const type = attributeValue(attributes, "Type") ?? "unknown";
      if (/\/hyperlink$/i.test(type)) hyperlinkCount += 1;
      else blockingTypes.push(type);
    }
  }
  return {
    totalCount,
    hyperlinkCount,
    blockingCount: blockingTypes.length,
    blockingTypes: uniqueSorted(blockingTypes),
  };
}

/**
 * PowerPoint stores modern content such as Office Math in mc:AlternateContent.
 * A capable Office client renders the first compatible mc:Choice and ignores
 * mc:Fallback, but a raw OOXML scan otherwise sees both copies. Selecting the
 * active branch before inventory prevents duplicate shape IDs, duplicate text
 * boxes, and false native-measurement bindings during export acceptance.
 */
function selectActiveMarkupCompatibilityContent(xml: string): string {
  let result = xml;
  for (let pass = 0; pass < 8; pass += 1) {
    let changed = false;
    result = result.replace(/<mc:AlternateContent\b[^>]*>([\s\S]*?)<\/mc:AlternateContent>/g, (_alternate, body: string) => {
      const choice = body.match(/<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/)?.[1];
      const fallback = body.match(/<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/)?.[1];
      changed = true;
      return choice ?? fallback ?? "";
    });
    if (!changed) break;
  }
  return result;
}

function directCellFillToken(cellXml: string): string | undefined {
  const properties = cellXml.match(/<a:tcPr\b[^>]*>([\s\S]*?)<\/a:tcPr>|<a:tcPr\b[^>]*\/>/)?.[0];
  // The first solid fill inside tcPr may belong to a border line. Semantic
  // table meaning is carried by the direct cell fill, so ignore line colors.
  const cellProperties = properties
    ?.replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>/g, "")
    .replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[^>]*\/>/g, "");
  const fill = cellProperties?.match(/<a:solidFill\b[^>]*>([\s\S]*?)<\/a:solidFill>/)?.[1];
  const color = fill?.match(/<a:(srgbClr|schemeClr|sysClr)\b[^>]*\bval=(?:"([^"]+)"|'([^']+)')/i);
  return color ? normalizeCellFillToken(color[1], color[2] ?? color[3] ?? "") : undefined;
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
    const cellBlocks = [...block.matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((match) => match[0]);
    const semanticColorTokens = uniqueSorted(cellBlocks.map((cell) => semanticColorRoleForToken(directCellFillToken(cell)) ?? ""));
    const cellTexts = cellBlocks.map((cell) => extractTextRuns(cell).join(""));
    const tableId = `slide-${slideNumber}-table-${index + 1}`;
    const columns = [...(block.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/)?.[0] ?? "").matchAll(/<a:gridCol\b([^>]*)\/?\s*>/g)].map((match, columnIndex) => ({
      id: `${tableId}-column-${columnIndex + 1}`,
      index: columnIndex + 1,
      widthEmu: Number(attributeValue(match[1] ?? "", "w")) || 0,
    }));
    const rowBlocks = [...block.matchAll(/<a:tr\b([^>]*)>([\s\S]*?)<\/a:tr>/g)];
    const rows = rowBlocks.map((match, rowIndex) => ({
      id: `${tableId}-row-${rowIndex + 1}`,
      index: rowIndex + 1,
      heightEmu: Number(attributeValue(match[1] ?? "", "h")) || 0,
    }));
    const cells = [];
    for (let rowIndex = 0; rowIndex < rowBlocks.length; rowIndex += 1) {
      const rowXml = rowBlocks[rowIndex][2] ?? "";
      const rowCells = [...rowXml.matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)];
      for (let columnIndex = 0; columnIndex < rowCells.length; columnIndex += 1) {
        const cellAttributes = rowCells[columnIndex][1] ?? "";
        const cellXml = rowCells[columnIndex][0];
        const cellProperties = cellXml.match(/<a:tcPr\b([^>]*)/)?.[1] ?? "";
        const numberAttribute = (name: string, fallback: number) => {
          const value = Number(attributeValue(cellProperties, name));
          return Number.isFinite(value) ? value : fallback;
        };
        const spanAttribute = (name: string) => Math.max(1, Number(attributeValue(cellAttributes, name)) || 1);
        const { textRuns, paragraphRunCounts, runBreaksBefore } = extractTableCellRuns(cellXml);
        const text = visibleTextFromRuns({ textRuns, paragraphRunCounts, runBreaksBefore });
        const fillToken = directCellFillToken(cellXml);
        const semanticColorRole = semanticColorRoleForToken(fillToken);
        const fontSizes = uniqueSorted([...cellXml.matchAll(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)].map((match) => String(Number(match[1] ?? match[2]) / 100))).map(Number);
        const anchor = attributeValue(cellProperties, "anchor")?.toLowerCase();
        cells.push({
          id: `${tableId}-cell-r${rowIndex + 1}-c${columnIndex + 1}`,
          row: rowIndex + 1,
          column: columnIndex + 1,
          rowSpan: spanAttribute("rowSpan"),
          columnSpan: spanAttribute("gridSpan"),
          horizontalMergeContinuation: attributeValue(cellAttributes, "hMerge") === "1",
          verticalMergeContinuation: attributeValue(cellAttributes, "vMerge") === "1",
          text,
          textRuns,
          paragraphRunCounts,
          runBreaksBefore,
          textHash: await sha256Text(text),
          characterCount: text.length,
          paragraphCount: Math.max(1, xmlCount(cellXml, /<a:p\b/g)),
          fontFamilies: uniqueSorted(extractFonts(cellXml)),
          fontSizes,
          fillToken,
          semanticColorRole,
          marginsEmu: {
            left: numberAttribute("marL", 91_440),
            right: numberAttribute("marR", 91_440),
            top: numberAttribute("marT", 45_720),
            bottom: numberAttribute("marB", 45_720),
          },
          horizontalAlignment: paragraphAlignment(cellXml),
          verticalAlignment: (anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top") as "top" | "middle" | "bottom",
        });
      }
    }
    const contentHash = await sha256Text(JSON.stringify(cellBlocks.map((cell) => {
      const content = extractTableCellRuns(cell);
      return { textRuns: content.textRuns, breaksBefore: content.runBreaksBefore.map((value) => value === "none" ? "none" : "break") };
    })));
    const structureHash = await sha256Text(JSON.stringify({
      rows: [...block.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)].map((row) => [...row[1].matchAll(/<a:tc\b([^>]*)>[\s\S]*?<\/a:tc>/g)].map((cell) => ({
        gridSpan: attributeValue(cell[1] ?? "", "gridSpan") ?? null,
        rowSpan: attributeValue(cell[1] ?? "", "rowSpan") ?? null,
        hMerge: attributeValue(cell[1] ?? "", "hMerge") ?? null,
        vMerge: attributeValue(cell[1] ?? "", "vMerge") ?? null,
      }))),
      columns: xmlCount(block.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/)?.[0] ?? "", /<a:gridCol\b/g),
    }));
    tables.push({
      id: tableId,
      slideNumber,
      ordinal: index + 1,
      rowCount: xmlCount(block, /<a:tr\b/g),
      columnCount: xmlCount(block.match(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/)?.[0] ?? "", /<a:gridCol\b/g),
      mergedCellCount: xmlCount(block, /\b(?:gridSpan|rowSpan|hMerge|vMerge)=/g),
      totalCellCharacterCount: cellTexts.reduce((sum, text) => sum + text.length, 0),
      maximumCellCharacterCount: Math.max(0, ...cellTexts.map((text) => text.length)),
      styleId: styleId ? decodeXml(styleId).trim() : undefined,
      styleFlags,
      cellFonts,
      colorTokens,
      semanticColorTokens,
      marginSignatures,
      styleFingerprint,
      contentHash,
      structureHash,
      columns,
      rows,
      cells,
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

interface ParsedTextShape {
  inventory: TextBoxInventoryItem;
  text: string;
}

function extractDirectShapeBlocks(xml: string): string[] {
  const tokens = [...xml.matchAll(/<p:grpSp\b|<\/p:grpSp>|<p:sp\b|<\/p:sp>/g)];
  const blocks: string[] = [];
  let groupDepth = 0;
  let shapeStart: number | undefined;
  for (const token of tokens) {
    if (token[0] === "<p:grpSp") groupDepth += 1;
    else if (token[0] === "</p:grpSp>") groupDepth = Math.max(0, groupDepth - 1);
    else if (token[0] === "<p:sp" && groupDepth === 0) shapeStart = token.index;
    else if (token[0] === "</p:sp>" && groupDepth === 0 && shapeStart !== undefined) {
      blocks.push(xml.slice(shapeStart, (token.index ?? shapeStart) + token[0].length));
      shapeStart = undefined;
    }
  }
  return blocks;
}

interface DirectObjectBlock {
  sourceElement: SlideEditableObjectElement;
  block: string;
}

function extractDirectObjectBlocks(xml: string): DirectObjectBlock[] {
  const tokens = [...xml.matchAll(/<\/?p:(grpSp|sp|pic|graphicFrame|cxnSp)\b[^>]*>/g)];
  const blocks: DirectObjectBlock[] = [];
  let groupDepth = 0;
  let active: { sourceElement: SlideEditableObjectElement; start: number } | undefined;
  for (const token of tokens) {
    const whole = token[0];
    const tag = token[1] as "grpSp" | "sp" | "pic" | "graphicFrame" | "cxnSp";
    const closing = whole.startsWith("</");
    const sourceElement = `p:${tag}` as SlideEditableObjectElement;
    if (!closing) {
      if (tag === "grpSp") {
        if (groupDepth === 0) active = { sourceElement, start: token.index ?? 0 };
        groupDepth += 1;
      } else if (groupDepth === 0) active = { sourceElement, start: token.index ?? 0 };
      continue;
    }
    if (tag === "grpSp") {
      groupDepth = Math.max(0, groupDepth - 1);
      if (groupDepth === 0 && active?.sourceElement === sourceElement) {
        blocks.push({ sourceElement, block: xml.slice(active.start, (token.index ?? active.start) + whole.length) });
        active = undefined;
      }
    } else if (groupDepth === 0 && active?.sourceElement === sourceElement) {
      blocks.push({ sourceElement, block: xml.slice(active.start, (token.index ?? active.start) + whole.length) });
      active = undefined;
    }
  }
  return blocks;
}

async function extractEditableObjects(slideNumber: number, xml: string): Promise<SlideEditableObject[]> {
  const objects: SlideEditableObject[] = [];
  let tableOrdinal = 0;
  let pictureOrdinal = 0;
  for (const { sourceElement, block } of extractDirectObjectBlocks(xml)) {
    const nonVisual = block.match(/<p:cNvPr\b([^>]*)>/)?.[1] ?? "";
    const shapeId = attributeValue(nonVisual, "id")?.trim();
    if (!shapeId) continue;
    const transformTag = sourceElement === "p:graphicFrame" ? "p:xfrm" : "a:xfrm";
    const transform = block.match(new RegExp(`<${transformTag}\\b([^>]*)>[\\s\\S]*?<a:off\\b([^>]*)\\/>[\\s\\S]*?<a:ext\\b([^>]*)\\/>[\\s\\S]*?<\\/${transformTag}>`));
    if (!transform) continue;
    const x = Number(attributeValue(transform[2] ?? "", "x"));
    const y = Number(attributeValue(transform[2] ?? "", "y"));
    const width = Number(attributeValue(transform[3] ?? "", "cx"));
    const height = Number(attributeValue(transform[3] ?? "", "cy"));
    if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) continue;
    const rotationUnits = Number(attributeValue(transform[1] ?? "", "rot"));
    const rotation = Number.isFinite(rotationUnits) ? rotationUnits / 60_000 : 0;
    const visibleText = visibleTextFromRuns(extractTableCellRuns(block));
    let kind: SlideEditableObject["kind"];
    if (sourceElement === "p:pic") { kind = "picture"; pictureOrdinal += 1; }
    else if (sourceElement === "p:graphicFrame" && /<a:tbl\b/.test(block)) { kind = "table"; tableOrdinal += 1; }
    else if (sourceElement === "p:graphicFrame" && /<c:chart\b/.test(block)) kind = "chart";
    else if (sourceElement === "p:graphicFrame") kind = "graphic-frame";
    else if (sourceElement === "p:cxnSp") kind = "connector";
    else if (sourceElement === "p:grpSp") kind = "group";
    else kind = visibleText ? "text" : "shape";
    objects.push({
      id: `slide-${slideNumber}-object-${shapeId}`,
      slideNumber,
      shapeId,
      name: attributeValue(nonVisual, "name")?.trim() || `${kind.replace("-", " ")} ${shapeId}`,
      kind,
      sourceElement,
      geometry: { x, y, width, height, rotation },
      canMove: kind !== "group",
      canResize: !["connector", "group"].includes(kind),
      textHash: visibleText ? await sha256Text(visibleText) : undefined,
      tableId: kind === "table" ? `slide-${slideNumber}-table-${tableOrdinal}` : undefined,
      pictureId: kind === "picture" ? `slide-${slideNumber}-picture-${pictureOrdinal}` : undefined,
    });
  }
  return objects;
}

function textRole(block: string, y: number, height: number, maximumFontSize: number | undefined, characterCount: number): TextBoxInventoryItem["role"] {
  const placeholder = block.match(/<p:ph\b([^>]*)\/?\s*>/)?.[1] ?? "";
  const placeholderType = attributeValue(placeholder, "type")?.toLowerCase();
  if (placeholderType === "title" || placeholderType === "ctrtitle" || placeholderType === "subtitle") return "title";
  if (y < 1.5 * EMU_PER_INCH && (maximumFontSize ?? 0) >= 24) return "title";
  if ((maximumFontSize ?? 100) <= 14 && (y > 5.25 * EMU_PER_INCH || characterCount < 160)) return "caption";
  if (height <= 0.65 * EMU_PER_INCH && characterCount < 120) return "label";
  return characterCount > 0 ? "body" : "other";
}

function paragraphAlignment(block: string): TextBoxInventoryItem["paragraphAlignment"] {
  const values = [...block.matchAll(/<a:pPr\b([^>]*)/g)]
    .map((match) => attributeValue(match[1] ?? "", "algn")?.toLowerCase())
    .filter((value): value is string => Boolean(value))
    .map((value) => value === "ctr" ? "center" : value === "r" ? "right" : ["just", "justlow", "dist", "thaidist"].includes(value) ? "justified" : "left");
  if (values.length === 0 && /<p:ph\b[^>]*\btype=(?:"ctrTitle"|'ctrTitle')/i.test(block)) return "center";
  const distinct = uniqueSorted(values);
  return distinct.length > 1 ? "mixed" : (distinct[0] as TextBoxInventoryItem["paragraphAlignment"] | undefined) ?? "left";
}

function autoFitMode(block: string): TextBoxInventoryItem["autoFit"] {
  if (/<a:noAutofit\b/i.test(block)) return "none";
  if (/<a:normAutofit\b/i.test(block)) return "shrink-text";
  if (/<a:spAutoFit\b/i.test(block)) return "resize-shape";
  return "unspecified";
}

function verticalAlignment(block: string): TextBoxInventoryItem["verticalAlignment"] {
  const bodyProperties = block.match(/<a:bodyPr\b([^>]*)/)?.[1] ?? "";
  const anchor = attributeValue(bodyProperties, "anchor")?.toLowerCase();
  return anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
}

function textMargins(block: string): { left: number; right: number; top: number; bottom: number } {
  const bodyProperties = block.match(/<a:bodyPr\b([^>]*)/)?.[1] ?? "";
  const value = (name: string, fallback: number) => {
    const parsed = Number(attributeValue(bodyProperties, name));
    return Number.isFinite(parsed) ? parsed : fallback;
  };
  return {
    left: value("lIns", 91_440),
    right: value("rIns", 91_440),
    top: value("tIns", 45_720),
    bottom: value("bIns", 45_720),
  };
}

function paragraphOpticalMetrics(block: string, margins: ReturnType<typeof textMargins>) {
  const paragraphs = [...block.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)].map((match) => match[0]);
  const paragraphLeftMargins: number[] = [];
  const paragraphIndents: number[] = [];
  const textStartOffsets: number[] = [];
  let bulletParagraphCount = 0;
  let directMarginCount = 0;

  for (const paragraph of paragraphs) {
    const properties = paragraph.match(/<a:pPr\b([^>]*)/)?.[1] ?? "";
    const marginValue = attributeValue(properties, "marL");
    const indentValue = attributeValue(properties, "indent");
    const leftMargin = marginValue === undefined ? 0 : Number(marginValue);
    const indent = indentValue === undefined ? 0 : Number(indentValue);
    const hasDirectMargin = marginValue !== undefined && Number.isFinite(leftMargin);
    const hasDirectIndent = indentValue !== undefined && Number.isFinite(indent);
    const hasBullet = !/<a:buNone\b/i.test(paragraph) && /<a:(?:buChar|buAutoNum|buBlip)\b/i.test(paragraph);

    if (hasDirectMargin) {
      paragraphLeftMargins.push(leftMargin);
      directMarginCount += 1;
    }
    if (hasDirectIndent) paragraphIndents.push(indent);
    if (hasBullet) bulletParagraphCount += 1;

    // PowerPoint positions bullet text at marL; a non-bulleted first line also
    // applies its indent. Inherited list styles remain intentionally marked as
    // partial confidence instead of being presented as exact geometry.
    textStartOffsets.push(margins.left + (hasDirectMargin ? leftMargin : 0) + (hasBullet ? 0 : hasDirectIndent ? indent : 0));
  }

  const opticalLeftOffsetEmu = textStartOffsets.length > 0 ? Math.min(...textStartOffsets) : margins.left;
  return {
    paragraphLeftMarginsEmu: uniqueSorted(paragraphLeftMargins.map(String)).map(Number),
    paragraphIndentsEmu: uniqueSorted(paragraphIndents.map(String)).map(Number),
    bulletParagraphCount,
    opticalLeftOffsetEmu,
    opticalAlignmentConfidence: paragraphs.length > 0 && directMarginCount === paragraphs.length ? "direct" as const : "partial-inheritance" as const,
  };
}

function estimatedLineCount(block: string, availableWidthEmu: number, fontSizePt: number): number {
  const widthPt = Math.max(1, availableWidthEmu / EMU_PER_POINT);
  const charactersPerLine = Math.max(4, Math.floor(widthPt / Math.max(1, fontSizePt * 0.54)));
  const paragraphs = [...block.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)].map((match) => match[0]);
  return Math.max(1, (paragraphs.length ? paragraphs : [block]).reduce((sum, paragraph) => {
    const characters = extractTextRuns(paragraph).join("").length;
    const explicitBreaks = xmlCount(paragraph, /<a:br\b/g);
    return sum + Math.max(1, Math.ceil(Math.max(1, characters) / charactersPerLine) + explicitBreaks);
  }, 0));
}

async function extractTextParagraphs(block: string): Promise<TextParagraphInventoryItem[]> {
  const paragraphs = [...block.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)].map((match) => match[0]);
  const inventory: TextParagraphInventoryItem[] = [];
  for (const paragraph of paragraphs) {
    const text = visibleTextFromRuns(extractTableCellRuns(paragraph));
    if (!text) continue;
    const properties = paragraph.match(/<a:pPr\b([^>]*)/)?.[1] ?? "";
    const directBullet = /<a:(?:buChar|buAutoNum|buBlip)\b/i.test(paragraph);
    const directNoBullet = /<a:buNone\b/i.test(paragraph);
    const rawLevel = Number(attributeValue(properties, "lvl"));
    const fontSizes = uniqueSorted([...paragraph.matchAll(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)]
      .map((match) => String(Number(match[1] ?? match[2]) / 100))).map(Number);
    inventory.push({
      index: inventory.length + 1,
      text,
      textHash: await sha256Text(text),
      characterCount: text.length,
      bullet: directBullet,
      bulletConfidence: directBullet || directNoBullet ? "direct" : "inherited-possible",
      level: Number.isFinite(rawLevel) && rawLevel >= 0 ? rawLevel : 0,
      fontFamilies: uniqueSorted(extractFonts(paragraph)),
      fontSizes,
    });
  }
  return inventory;
}

async function extractTextBoxes(slideNumber: number, xml: string, slideWidth: number, slideHeight: number): Promise<{ shapes: ParsedTextShape[]; reviews: LayoutReviewItem[] }> {
  const shapes: ParsedTextShape[] = [];
  const reviews: LayoutReviewItem[] = [];
  const blocks = extractDirectShapeBlocks(xml);
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (!/<p:txBody\b/.test(block)) continue;
    const text = visibleTextFromRuns(extractTableCellRuns(block));
    const shapeId = block.match(/<p:cNvPr\b[^>]*\bid=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean);
    const transform = block.match(/<a:xfrm\b[^>]*>[\s\S]*?<a:off\b([^>]*)\/>[\s\S]*?<a:ext\b([^>]*)\/>[\s\S]*?<\/a:xfrm>/);
    if (!text || !shapeId || !transform) continue;
    const x = Number(attributeValue(transform[1] ?? "", "x"));
    const y = Number(attributeValue(transform[1] ?? "", "y"));
    const width = Number(attributeValue(transform[2] ?? "", "cx"));
    const height = Number(attributeValue(transform[2] ?? "", "cy"));
    if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) continue;
    const runFontSizes = [...block.matchAll(/<a:rPr\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)].map((match) => String(Number(match[1] ?? match[2]) / 100));
    const fallbackFontSizes = runFontSizes.length === 0 ? [...block.matchAll(/<a:endParaRPr\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)].map((match) => String(Number(match[1] ?? match[2]) / 100)) : [];
    const fontSizes = uniqueSorted([...runFontSizes, ...fallbackFontSizes]).map(Number);
    const fontSizePt = fontSizes.length > 0 ? Math.max(...fontSizes) : 16;
    const margins = textMargins(block);
    const opticalMetrics = paragraphOpticalMetrics(block, margins);
    const lineCount = estimatedLineCount(block, Math.max(1, width - margins.left - margins.right), fontSizePt);
    const requiredHeight = Math.round((lineCount * fontSizePt * 1.18) * EMU_PER_POINT + margins.top + margins.bottom);
    const fitRatio = Number((requiredHeight / height).toFixed(3));
    const offSlide = x < -OFF_SLIDE_TOLERANCE_EMU || y < -OFF_SLIDE_TOLERANCE_EMU || x + width > slideWidth + OFF_SLIDE_TOLERANCE_EMU || y + height > slideHeight + OFF_SLIDE_TOLERANCE_EMU;
    const nearEdge = !offSlide && (x < SAFE_MARGIN_EMU || y < SAFE_MARGIN_EMU || x + width > slideWidth - SAFE_MARGIN_EMU || y + height > slideHeight - SAFE_MARGIN_EMU);
    const autoFit = autoFitMode(block);
    const paragraphs = await extractTextParagraphs(block);
    const warnings: string[] = [];
    if (offSlide) warnings.push("The editable text box extends beyond the physical slide boundary.");
    else if (nearEdge) warnings.push("The editable text box enters the 0.25-inch safe-margin review zone.");
    const overflowThreshold = autoFit === "resize-shape" ? 1.35 : 1.18;
    const overflowRisk = fontSizes.length > 0 && fitRatio > overflowThreshold;
    if (overflowRisk) warnings.push(autoFit === "shrink-text" ? "The current text metrics may trigger automatic font shrinking." : "The current text metrics may exceed the available text-box height.");
    const geometry = { x, y, width, height };
    const inventory: TextBoxInventoryItem = {
      id: `slide-${slideNumber}-text-box-${index + 1}`,
      slideNumber,
      ordinal: index + 1,
      shapeId,
      text,
      textHash: await sha256Text(text),
      characterCount: text.length,
      paragraphCount: Math.max(1, xmlCount(block, /<a:p\b/g)),
      paragraphs,
      geometry,
      textInsets: margins,
      ...opticalMetrics,
      estimatedOpticalLeftEmu: x + opticalMetrics.opticalLeftOffsetEmu,
      fontFamilies: uniqueSorted(extractFonts(block)),
      fontSizes,
      directFontSizeKnown: fontSizes.length > 0,
      paragraphAlignment: paragraphAlignment(block),
      verticalAlignment: verticalAlignment(block),
      role: textRole(block, y, height, fontSizes.length ? Math.max(...fontSizes) : undefined, text.length),
      autoFit,
      estimatedLineCount: lineCount,
      estimatedRequiredHeightEmu: requiredHeight,
      fitRatio,
      safeAreaStatus: offSlide ? "off-slide" : nearEdge ? "near-edge" : "inside",
      warnings,
    };
    shapes.push({ inventory, text });
    if (offSlide) reviews.push({
      id: `slide-${slideNumber}-shape-${shapeId}-off-slide`,
      slideNumber,
      shapeId,
      rule: "off-slide",
      severity: "error",
      confidence: "high",
      reason: "Editable text extends outside the physical slide. Reposition or resize it without deleting, rewriting, or hiding content.",
      geometry,
    });
    else if (nearEdge) reviews.push({
      id: `slide-${slideNumber}-shape-${shapeId}-safe-area`,
      slideNumber,
      shapeId,
      rule: "safe-area",
      severity: "info",
      confidence: "medium",
      reason: "Text enters the 0.25-inch safe-margin review zone; confirm that this is intentional template furniture or a readable edge treatment.",
      geometry,
    });
    if (overflowRisk) reviews.push({
      id: `slide-${slideNumber}-shape-${shapeId}-overflow-risk`,
      slideNumber,
      shapeId,
      rule: "overflow-risk",
      severity: "warning",
      confidence: autoFit === "none" ? "high" : "medium",
      reason: autoFit === "shrink-text" ? `Estimated text demand is ${fitRatio.toFixed(2)}× the available height and may force PowerPoint to shrink the type.` : `Estimated text demand is ${fitRatio.toFixed(2)}× the available height; inspect the native PowerPoint render for clipping or unexpected wrapping.`,
      geometry,
      fitRatio,
    });
  }
  return { shapes, reviews };
}

function overlapRatio(target: TextBoxInventoryItem["geometry"], other: TextBoxInventoryItem["geometry"]): number {
  const overlapWidth = Math.max(0, Math.min(target.x + target.width, other.x + other.width) - Math.max(target.x, other.x));
  const overlapHeight = Math.max(0, Math.min(target.y + target.height, other.y + other.height) - Math.max(target.y, other.y));
  return target.width * target.height > 0 ? (overlapWidth * overlapHeight) / (target.width * target.height) : 0;
}

function roleBucket(role: TextBoxInventoryItem["role"]): "title" | "content" | "other" {
  if (role === "title") return "title";
  if (["body", "caption", "label"].includes(role)) return "content";
  return "other";
}

async function extractAlignmentRepairs(slideNumber: number, shapes: ParsedTextShape[], slideWidth: number): Promise<{ repairs: AlignmentRepairCandidate[]; reviews: LayoutReviewItem[] }> {
  const repairs: AlignmentRepairCandidate[] = [];
  const reviews: LayoutReviewItem[] = [];
  const leftAligned = shapes.filter((shape) => shape.inventory.paragraphAlignment === "left" && shape.inventory.safeAreaStatus !== "off-slide");
  if (leftAligned.length < 3) return { repairs, reviews };

  const clusterByLeftEdge = (candidates: ParsedTextShape[]) => {
    const clusters: Array<{ center: number; members: ParsedTextShape[] }> = [];
    for (const shape of candidates) {
      const opticalLeft = shape.inventory.estimatedOpticalLeftEmu;
      const cluster = clusters.find((item) => Math.abs(item.center - opticalLeft) <= 38_100);
      if (cluster) {
        cluster.members.push(shape);
        cluster.center = Math.round(cluster.members.reduce((sum, member) => sum + member.inventory.estimatedOpticalLeftEmu, 0) / cluster.members.length);
      } else clusters.push({ center: opticalLeft, members: [shape] });
    }
    return clusters.sort((left, right) => right.members.length - left.members.length);
  };

  if (slideNumber === 1) {
    const dominant = clusterByLeftEdge(leftAligned)[0];
    if (dominant && dominant.members.length >= 2) {
      for (const candidate of leftAligned.filter((shape) => {
        const box = shape.inventory.geometry;
        const delta = Math.abs(shape.inventory.estimatedOpticalLeftEmu - dominant.center);
        return delta >= 152_400 && delta <= 914_400 && box.y >= 3_500_000 && box.height <= 508_000 && shape.text.length >= 10;
      })) {
        const box = candidate.inventory.geometry;
        repairs.push({
          id: `slide-${slideNumber}-shape-${candidate.inventory.shapeId}-align-left`,
          slideNumber,
          shapeId: candidate.inventory.shapeId,
          textHash: candidate.inventory.textHash,
          source: box,
          target: { ...box, x: dominant.center - candidate.inventory.opticalLeftOffsetEmu },
          ruleId: "cover.dominant-left-edge",
          confidence: "high",
          rationale: `Align the visible text start of a lower cover block to the dominant optical edge used by ${dominant.members.length} peer text blocks; account for PowerPoint text insets and paragraph indents.`,
        });
      }
    }
  }

  const repairedShapeIds = new Set(repairs.map((repair) => repair.shapeId));
  const contentShapes = leftAligned.filter((shape) => roleBucket(shape.inventory.role) === "content");
  const clusters = clusterByLeftEdge(contentShapes);
  const dominant = clusters[0];
  const runnerUp = clusters[1];
  if (!dominant || dominant.members.length < 3 || (runnerUp && dominant.members.length - runnerUp.members.length < 2)) return { repairs, reviews };
  const peerWidths = dominant.members.map((shape) => shape.inventory.geometry.width).sort((left, right) => left - right);
  const medianWidth = peerWidths[Math.floor(peerWidths.length / 2)] ?? 1;
  const minY = Math.min(...dominant.members.map((shape) => shape.inventory.geometry.y));
  const maxBottom = Math.max(...dominant.members.map((shape) => shape.inventory.geometry.y + shape.inventory.geometry.height));
  for (const candidate of contentShapes.filter((shape) => !dominant.members.includes(shape) && !repairedShapeIds.has(shape.inventory.shapeId))) {
    const box = candidate.inventory.geometry;
    const delta = Math.abs(candidate.inventory.estimatedOpticalLeftEmu - dominant.center);
    if (delta < 127_000 || delta > 762_000) continue;
    const target = { ...box, x: dominant.center - candidate.inventory.opticalLeftOffsetEmu };
    const widthRatio = box.width / medianWidth;
    const verticalPeer = box.y >= minY - 0.75 * EMU_PER_INCH && box.y + box.height <= maxBottom + 0.75 * EMU_PER_INCH;
    const collision = shapes.some((shape) => shape !== candidate && overlapRatio(target, shape.inventory.geometry) > 0.12);
    const safelyInside = target.x >= 0 && target.x + target.width <= slideWidth;
    const highConfidence = delta <= 457_200 && widthRatio >= 0.65 && widthRatio <= 1.55 && verticalPeer && !collision && safelyInside;
    if (highConfidence) repairs.push({
      id: `slide-${slideNumber}-shape-${candidate.inventory.shapeId}-align-left`,
      slideNumber,
      shapeId: candidate.inventory.shapeId,
      textHash: candidate.inventory.textHash,
      source: box,
      target,
      ruleId: "peer.dominant-left-edge",
      confidence: "high",
      rationale: `Align a text-box outlier to the visible text edge shared by ${dominant.members.length} nearby content peers; PowerPoint text insets and paragraph indents are included while text, vertical position, size, and order remain unchanged.`,
    });
    else reviews.push({
      id: `slide-${slideNumber}-shape-${candidate.inventory.shapeId}-alignment-ambiguous`,
      slideNumber,
      shapeId: candidate.inventory.shapeId,
      rule: "alignment-ambiguous",
      severity: "info",
      confidence: "medium",
      reason: collision ? "A likely left-edge outlier cannot be moved safely because the target geometry would collide with another editable text box." : "A likely left-edge outlier needs visual review because peer role, distance, width, or vertical grouping is not decisive enough for an automatic move.",
      geometry: box,
    });
  }
  return { repairs, reviews };
}

function xmlCount(xml: string, expression: RegExp): number {
  return [...xml.matchAll(expression)].length;
}

function classifyTemplate(input: { allText: string; structuralText: string; themeIdentity: string; fonts: FontInventoryItem[] }): { classification: TemplateClassification; evidence: string[] } {
  const normalized = input.allText.toLowerCase();
  const structural = input.structuralText.toLowerCase();
  const themeIdentity = input.themeIdentity.toLowerCase();
  const fonts = input.fonts;
  const fontNames = new Set(fonts.map((font) => font.normalizedFamily));
  const evidence: string[] = [];
  const hasOrnlLanguage = normalized.includes("oak ridge national laboratory") || /\bornl\b/.test(normalized);
  const hasStructuralOrnlIdentity = structural.includes("oak ridge national laboratory") || /\bornl\b/.test(structural) || /\b(?:ornl|oak[\s_-]*ridge)\b/.test(themeIdentity);
  const sponsorThemeMarkers = [
    /\beere\b/,
    /\bdoe(?:\b|[-_ ])/,
    /department[\s_-]+of[\s_-]+energy/,
    /building[\s_-]+technologies[\s_-]+office/,
    /\bbto(?:\b|[-_ ])/,
  ];
  const sponsorThemeIdentity = sponsorThemeMarkers.some((marker) => marker.test(themeIdentity));
  const hasAptos = fontNames.has("aptos") || fontNames.has("aptos display");
  const hasLegacyOrnlFont = fontNames.has("century gothic");

  if (sponsorThemeIdentity) evidence.push("The PowerPoint theme identity contains a recognized sponsor marker; theme and master identity take precedence over organization names in ordinary slide copy.");
  if (hasStructuralOrnlIdentity) evidence.push("ORNL identity was found in theme, master, or layout structure rather than only in ordinary slide copy.");
  else if (hasOrnlLanguage) evidence.push("ORNL or Oak Ridge National Laboratory appears only in package copy and is not sufficient evidence of an ORNL template.");
  if (hasAptos) evidence.push("The package contains Aptos typography.");
  if (hasLegacyOrnlFont) evidence.push("The package contains Century Gothic legacy typography.");

  if (sponsorThemeIdentity) {
    evidence.push("Retain the detected sponsor theme as read-only source evidence. Presentation Studio still targets the current ORNL Template Pack by default; preserve the sponsor template only when a person explicitly selects that override.");
    return { classification: "sponsor", evidence };
  }
  if (hasStructuralOrnlIdentity) {
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
  let structuralText = "";
  let themeIdentity = "";

  for (const [path, xml] of xmlByPath.entries()) {
    const activeXml = selectActiveMarkupCompatibilityContent(xml);
    searchableText += ` ${extractTextRuns(activeXml).join(" ")}`;
    const kind = partKind(path);
    if (["master", "layout", "theme"].includes(kind)) structuralText += ` ${extractTextRuns(activeXml).join(" ")}`;
    if (kind === "theme") {
      themeIdentity += ` ${[...activeXml.matchAll(/\b(?:name|id|vid)=(?:"([^"]*)"|'([^']*)')/gi)].map((match) => decodeXml(match[1] ?? match[2] ?? "")).join(" ")}`;
    }
    for (const family of extractFonts(activeXml)) {
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
  const classified = classifyTemplate({ allText: searchableText, structuralText, themeIdentity, fonts });
  const presentationXml = xmlByPath.get("ppt/presentation.xml") ?? "";
  const slideSizeAttributes = presentationXml.match(/<p:sldSz\b([^>]*)/)?.[1] ?? "";
  const declaredSlideWidth = Number(attributeValue(slideSizeAttributes, "cx"));
  const declaredSlideHeight = Number(attributeValue(slideSizeAttributes, "cy"));
  const slideWidth = Number.isFinite(declaredSlideWidth) && declaredSlideWidth > 0 ? declaredSlideWidth : DEFAULT_SLIDE_WIDTH_EMU;
  const slideHeight = Number.isFinite(declaredSlideHeight) && declaredSlideHeight > 0 ? declaredSlideHeight : DEFAULT_SLIDE_HEIGHT_EMU;

  const slides: SlideInventoryItem[] = [];
  const tables: TableInventoryItem[] = [];
  const pictures: PictureInventoryItem[] = [];
  const textBoxes: TextBoxInventoryItem[] = [];
  const editableObjects: SlideEditableObject[] = [];
  const layoutReviews: LayoutReviewItem[] = [];
  const alignmentRepairs: AlignmentRepairCandidate[] = [];
  for (const path of slidePaths) {
    const xml = xmlByPath.get(path) ?? "";
    const activeXml = selectActiveMarkupCompatibilityContent(xml);
    const number = slideNumberForPart(path) ?? slides.length + 1;
    const relationshipPart = `ppt/slides/_rels/slide${number}.xml.rels`;
    const relationXml = xmlByPath.get(relationshipPart) ?? "";
    const runs = extractTextRuns(activeXml);
    const text = visibleTextFromRuns(extractTableCellRuns(activeXml));
    const title = runs.find((run) => run.trim().length > 0)?.trim().slice(0, 160) || `Slide ${number}`;
    const slideFonts = [...new Set(extractFonts(activeXml))].sort();
    const fontSizes = uniqueSorted([...activeXml.matchAll(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz=(?:"(\d+)"|'(\d+)')/g)].map((match) => String(Number(match[1] ?? match[2]) / 100))).map(Number);
    const warnings: string[] = [];
    const tableCount = xmlCount(activeXml, /<a:tbl\b/g);
    tables.push(...await extractTableInventory(number, activeXml));
    const pictureCount = xmlCount(activeXml, /<p:pic\b/g);
    pictures.push(...extractPictureInventory(number, activeXml));
    const extractedTextBoxes = await extractTextBoxes(number, activeXml, slideWidth, slideHeight);
    textBoxes.push(...extractedTextBoxes.shapes.map((shape) => shape.inventory));
    editableObjects.push(...await extractEditableObjects(number, activeXml));
    layoutReviews.push(...extractedTextBoxes.reviews);
    const alignment = await extractAlignmentRepairs(number, extractedTextBoxes.shapes, slideWidth);
    alignmentRepairs.push(...alignment.repairs);
    layoutReviews.push(...alignment.reviews);
    const connectorCount = xmlCount(activeXml, /<p:cxnSp\b/g);
    const chartCount = xmlCount(relationXml, /relationships\/chart(?:"|\/)/gi);
    const commentPart = xmlByPath.get(`ppt/comments/comment${number}.xml`) ?? "";
    const commentCount = xmlCount(commentPart, /<p:cm\b/g);
    if (!text) warnings.push("No visible text was detected in the slide XML.");
    if (text.length > 1_400) warnings.push("Dense visible text should receive visual fit review.");
    slides.push({
      id: `slide-${number}`,
      number,
      sourcePart: path,
      sourcePartSha256: await sha256Text(xml),
      relationshipPart: relationXml ? relationshipPart : undefined,
      relationshipPartSha256: relationXml ? await sha256Text(relationXml) : undefined,
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
  const relationshipEntries = [...xmlByPath.entries()].filter(([path]) => path.endsWith(".rels"));
  const relationshipXml = relationshipEntries.map(([, xml]) => xml).join("\n");
  const containsMacros = /macroEnabled|vbaProject/i.test(contentTypes + relationshipXml) || paths.some((path) => /vbaProject\.bin$/i.test(path));
  const containsOleObjects = paths.some((path) => /^ppt\/embeddings\//i.test(path)) || /relationships\/oleObject/i.test(relationshipXml);
  const externalRelationships = inventoryExternalRelationships(relationshipEntries);
  const containsExternalRelationships = externalRelationships.totalCount > 0;
  const containsBlockingExternalRelationships = externalRelationships.blockingCount > 0;
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
  for (const repair of alignmentRepairs) findings.push(finding({
    ruleId: repair.ruleId,
    category: "layout",
    severity: "warning",
    confidence: repair.confidence,
    slideNumber: repair.slideNumber,
    message: repair.ruleId === "cover.dominant-left-edge" ? "A cover text block is offset from the dominant content alignment." : "A text box is offset from the dominant edge used by nearby content peers.",
    evidence: `Text box left edge ${repair.source.x} EMU; dominant peer edge ${repair.target.x} EMU.`,
    autoFixable: true,
  }));
  for (const review of layoutReviews) findings.push(finding({
    ruleId: `layout.${review.rule}`,
    category: "layout",
    severity: review.severity,
    confidence: review.confidence,
    slideNumber: review.slideNumber,
    message: review.rule === "overflow-risk" ? "Text-box fit needs native-render review." : review.rule === "off-slide" ? "Editable text extends beyond the slide boundary." : review.rule === "safe-area" ? "Editable text enters the safe-margin review zone." : "A possible alignment outlier is not safe to move automatically.",
    evidence: review.reason,
    autoFixable: false,
  }));
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
  if (externalRelationships.hyperlinkCount > 0) {
    findings.push(finding({
      ruleId: "production.external-hyperlinks",
      category: "production",
      severity: "info",
      confidence: "high",
      message: `${externalRelationships.hyperlinkCount} ordinary hyperlink relationship${externalRelationships.hyperlinkCount === 1 ? " is" : "s are"} preserved during editing.`,
      evidence: "Hyperlink relationship parts remain byte-preserved unless the linked object itself is intentionally replaced.",
      autoFixable: false,
    }));
  }
  if (containsMacros || containsOleObjects || containsBlockingExternalRelationships) {
    findings.push(finding({
      ruleId: "production.advanced-content",
      category: "production",
      severity: "error",
      confidence: "high",
      message: "Advanced or externally linked content requires manual review before cleanup export.",
      evidence: `Macros: ${containsMacros ? "yes" : "no"}; embedded OLE: ${containsOleObjects ? "yes" : "no"}; blocking external relationships: ${externalRelationships.blockingCount}${externalRelationships.blockingTypes.length > 0 ? ` (${externalRelationships.blockingTypes.join(", ")})` : ""}.`,
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
    semanticVisualVersion: PPTX_AUDIT_SEMANTIC_VISUAL_VERSION,
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
    externalHyperlinkCount: externalRelationships.hyperlinkCount,
    blockingExternalRelationshipCount: externalRelationships.blockingCount,
    containsBlockingExternalRelationships,
    packageFileCount: paths.length,
    expandedByteLength,
    slideSize: { width: slideWidth, height: slideHeight },
    classification: classified.classification,
    classificationEvidence: classified.evidence,
    fonts,
    slides,
    tables,
    pictures,
    textBoxes,
    editableObjects,
    layoutReviews,
    alignmentRepairs,
    findings,
    warnings: modernCommentCount > 0 ? ["Modern PowerPoint comments were retained as unsupported package parts and do not block the audit."] : [],
  };
}

export function extractSlideText(xml: string): string {
  return extractTextRuns(xml).join(" ").replace(/\s+/g, " ").trim();
}
