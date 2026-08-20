import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { sha256, sha256Text } from "./hash";
import { deriveLayoutSemantics, type TemplateLayoutSemantics } from "./layout-semantics";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@",
  removeNSPrefix: true,
  parseAttributeValue: false,
  trimValues: false,
});

type XmlRecord = Record<string, unknown>;

export interface TemplatePreviewElement {
  id: string;
  kind: "shape" | "text" | "image";
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  geometry: "rect" | "ellipse" | "line";
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  opacity?: number;
  text?: string;
  textColor?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: number;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
  placeholderType?: string;
  placeholderIndex?: string;
  mediaId?: string;
  sourceCropped?: boolean;
  sourcePart?: string;
  sourceShapeId?: string;
  origin?: "master" | "layout" | "slide";
  textHash?: string;
  sourceParagraphs?: Array<{
    index: number;
    text: string;
    textHash: string;
    characterCount: number;
    bullet: boolean;
    bulletConfidence: "direct" | "inherited-possible";
    level: number;
    fontFamilies: string[];
    fontSizes: number[];
  }>;
}

export interface TemplateLayoutPreview {
  id: string;
  name: string;
  category: "title" | "content" | "image" | "conclusion" | "other";
  background: string;
  elements: TemplatePreviewElement[];
  placeholderTypes: string[];
  sourcePart: string;
  semantic?: TemplateLayoutSemantics;
}

export interface TemplateCatalog {
  id: string;
  name: string;
  sha256: string;
  slideWidth: number;
  slideHeight: number;
  masterCount: number;
  layouts: TemplateLayoutPreview[];
  media: Record<string, string>;
  generatedAt: string;
}

export interface SlideRenderPreview extends TemplateLayoutPreview {
  number: number;
  title: string;
  hidden: boolean;
  renderWarnings: string[];
}

export interface SlideRenderCatalog {
  id: string;
  name: string;
  sha256: string;
  slideWidth: number;
  slideHeight: number;
  slides: SlideRenderPreview[];
  media: Record<string, string>;
  generatedAt: string;
  renderer: "local-ooxml-preview";
}

interface Relationship {
  id: string;
  type: string;
  target: string;
}

interface Theme {
  colors: Record<string, string>;
  majorFont: string;
  minorFont: string;
}

interface GroupContext {
  x: number;
  y: number;
  width: number;
  height: number;
  childX: number;
  childY: number;
  childWidth: number;
  childHeight: number;
}

function record(value: unknown): XmlRecord {
  return value && typeof value === "object" && !Array.isArray(value) ? value as XmlRecord : {};
}

function array<T = unknown>(value: unknown): T[] {
  if (value === undefined || value === null) return [];
  return (Array.isArray(value) ? value : [value]) as T[];
}

function textValue(value: unknown): string {
  if (typeof value === "string" || typeof value === "number") return String(value);
  const item = record(value);
  return typeof item["#text"] === "string" ? item["#text"] as string : "";
}

function numeric(value: unknown, fallback = 0): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizePart(basePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segments = `${basePart.slice(0, basePart.lastIndexOf("/") + 1)}${target}`.split("/");
  const normalized: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") normalized.pop();
    else normalized.push(segment);
  }
  return normalized.join("/");
}

function relationshipsPart(part: string): string {
  const slash = part.lastIndexOf("/");
  return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
}

async function readXml(zip: JSZip, part: string): Promise<XmlRecord> {
  const file = zip.file(part);
  if (!file) throw new Error(`The template is missing ${part}.`);
  return record(parser.parse(await file.async("text")));
}

async function readRelationships(zip: JSZip, part: string): Promise<Relationship[]> {
  const file = zip.file(relationshipsPart(part));
  if (!file) return [];
  const parsed = record(parser.parse(await file.async("text")));
  return array(record(parsed.Relationships).Relationship).map((value) => {
    const item = record(value);
    return {
      id: String(item["@Id"] ?? ""),
      type: String(item["@Type"] ?? ""),
      target: normalizePart(part, String(item["@Target"] ?? "")),
    };
  }).filter((item) => item.id && item.target);
}

function findRelationship(relationships: Relationship[], id: string): Relationship | undefined {
  return relationships.find((item) => item.id === id);
}

function colorNode(value: unknown): string | undefined {
  const item = record(value);
  const direct = record(item.srgbClr)["@val"];
  if (direct) return `#${String(direct).replace(/^#/, "")}`;
  const system = record(item.sysClr)["@lastClr"];
  return system ? `#${String(system).replace(/^#/, "")}` : undefined;
}

function parseTheme(themeXml: XmlRecord): Theme {
  const theme = record(themeXml.theme);
  const elements = record(theme.themeElements);
  const scheme = record(elements.clrScheme);
  const colors: Record<string, string> = {};
  for (const [key, value] of Object.entries(scheme)) {
    if (key.startsWith("@")) continue;
    const color = colorNode(value);
    if (color) colors[key] = color;
  }
  colors.tx1 = colors.dk1 ?? "#373A36";
  colors.tx2 = colors.dk2 ?? "#00454D";
  colors.bg1 = colors.lt1 ?? "#FFFFFF";
  colors.bg2 = colors.lt2 ?? "#DBDCDB";
  const fontScheme = record(elements.fontScheme);
  const majorFont = String(record(record(fontScheme.majorFont).latin)["@typeface"] ?? "Aptos Display");
  const minorFont = String(record(record(fontScheme.minorFont).latin)["@typeface"] ?? "Aptos");
  return { colors, majorFont, minorFont };
}

function schemeColor(value: unknown, theme: Theme): { color?: string; opacity?: number } {
  const fill = record(value);
  if (fill.noFill !== undefined) return {};
  const srgb = record(fill.srgbClr);
  const scheme = record(fill.schemeClr);
  const raw = srgb["@val"] ? `#${String(srgb["@val"]).replace(/^#/, "")}` : theme.colors[String(scheme["@val"] ?? "")];
  const alpha = numeric(record(srgb.alpha)["@val"] ?? record(scheme.alpha)["@val"], 100000) / 100000;
  return { color: raw, opacity: Math.min(1, Math.max(0, alpha)) };
}

function fillFor(node: XmlRecord, theme: Theme): { color?: string; opacity?: number } {
  if (node.noFill !== undefined) return {};
  if (node.solidFill !== undefined) return schemeColor(record(node.solidFill), theme);
  const gradient = record(node.gradFill);
  const stops = array(record(gradient.gsLst).gs);
  if (stops.length > 0) return schemeColor(record(stops[0]), theme);
  return {};
}

function collectText(value: unknown): string {
  if (typeof value === "string" || typeof value === "number") return String(value);
  if (Array.isArray(value)) return value.map(collectText).join("");
  const item = record(value);
  let output = "";
  for (const [key, child] of Object.entries(item)) {
    if (key === "t") output += textValue(child);
    else if (key === "br") output += "\n";
    else if (!key.startsWith("@")) output += collectText(child);
  }
  return output;
}

function paragraphText(txBody: XmlRecord): string {
  return array(txBody.p).map((paragraph) => collectText(paragraph).trim()).filter(Boolean).join("\n");
}

function firstRunProperties(txBody: XmlRecord): XmlRecord {
  const paragraph = record(array(txBody.p)[0]);
  const run = record(array(paragraph.r)[0]);
  const paragraphProperties = record(paragraph.pPr);
  const levelProperties = record(record(txBody.lstStyle).lvl1pPr);
  return {
    ...record(levelProperties.defRPr),
    ...record(paragraphProperties.defRPr),
    ...record(run.rPr),
    ...record(paragraph.endParaRPr),
  };
}

function transformFor(value: unknown): { x: number; y: number; width: number; height: number; rotation: number } | null {
  const xfrm = record(value);
  const off = record(xfrm.off);
  const ext = record(xfrm.ext);
  if (off["@x"] === undefined || off["@y"] === undefined || ext["@cx"] === undefined || ext["@cy"] === undefined) return null;
  return {
    x: numeric(off["@x"]),
    y: numeric(off["@y"]),
    width: Math.max(0, numeric(ext["@cx"])),
    height: Math.max(0, numeric(ext["@cy"])),
    rotation: numeric(xfrm["@rot"]) / 60000,
  };
}

function mapTransform(transform: NonNullable<ReturnType<typeof transformFor>>, context: GroupContext) {
  const scaleX = context.childWidth ? context.width / context.childWidth : 1;
  const scaleY = context.childHeight ? context.height / context.childHeight : 1;
  return {
    x: context.x + (transform.x - context.childX) * scaleX,
    y: context.y + (transform.y - context.childY) * scaleY,
    width: transform.width * scaleX,
    height: transform.height * scaleY,
    rotation: transform.rotation,
  };
}

function placeholderInfo(shape: XmlRecord): { type?: string; index?: string } {
  const placeholder = record(record(record(shape.nvSpPr).nvPr).ph);
  const type = String(placeholder["@type"] ?? "");
  const index = placeholder["@idx"] === undefined ? undefined : String(placeholder["@idx"]);
  return { type: type || (index !== undefined ? "body" : undefined), index };
}

function placeholderFor(shape: XmlRecord): string | undefined {
  return placeholderInfo(shape).type;
}

function shapeName(shape: XmlRecord, picture = false): string {
  const nonVisual = picture ? record(record(shape.nvPicPr).cNvPr) : record(record(shape.nvSpPr).cNvPr);
  return String(nonVisual["@name"] ?? "Template element");
}

function shapeElements(shape: XmlRecord, context: GroupContext, theme: Theme, index: number, fallback?: TemplatePreviewElement): TemplatePreviewElement[] {
  const spPr = record(shape.spPr);
  const transform = transformFor(spPr.xfrm) ?? (fallback ? { x: fallback.x, y: fallback.y, width: fallback.width, height: fallback.height, rotation: fallback.rotation } : null);
  if (!transform) return [];
  const mapped = mapTransform(transform, context);
  const txBody = record(shape.txBody);
  const text = paragraphText(txBody);
  const fill = fillFor(spPr, theme);
  const line = record(spPr.ln);
  const stroke = fillFor(line, theme);
  const geometryName = String(record(spPr.prstGeom)["@prst"] ?? "rect");
  const geometry = geometryName === "ellipse" ? "ellipse" : geometryName === "line" ? "line" : "rect";
  const name = shapeName(shape);
  const placeholder = placeholderInfo(shape);
  const sourceShapeId = String(record(record(shape.nvSpPr).cNvPr)["@id"] ?? index);
  const base: TemplatePreviewElement = {
    id: `shape-${index}-${sourceShapeId}`,
    kind: "shape",
    name,
    sourceShapeId,
    ...mapped,
    geometry,
    fill: fill.color,
    stroke: stroke.color,
    strokeWidth: numeric(line["@w"], 12700),
    opacity: fill.opacity,
    placeholderType: placeholder.type,
    placeholderIndex: placeholder.index,
  };
  if (!text) return fill.color || stroke.color || base.placeholderType ? [base] : [];
  const run = firstRunProperties(txBody);
  const fontRef = record(record(shape.style).fontRef);
  const runFill = run.solidFill !== undefined ? schemeColor(record(run.solidFill), theme) : schemeColor(fontRef, theme);
  const typeface = String(record(run.latin)["@typeface"] ?? "");
  const resolvedTypeface = typeface === "+mj-lt" ? theme.majorFont : typeface === "+mn-lt" ? theme.minorFont : typeface || fallback?.fontFamily || theme.minorFont;
  const paragraph = record(array(txBody.p)[0]);
  const alignment = String(record(paragraph.pPr)["@algn"] ?? record(record(txBody.lstStyle).lvl1pPr)["@algn"] ?? (fallback?.textAlign === "center" ? "ctr" : fallback?.textAlign === "right" ? "r" : "l"));
  const anchor = String(record(txBody.bodyPr)["@anchor"] ?? (fallback?.verticalAlign === "center" ? "ctr" : fallback?.verticalAlign === "bottom" ? "b" : "t"));
  const rawSize = run["@sz"];
  const rawBold = run["@b"];
  return [
    ...(fill.color || stroke.color ? [base] : []),
    {
      ...base,
      id: `${base.id}-text`,
      kind: "text",
      fill: undefined,
      stroke: undefined,
      text,
      textColor: runFill.color ?? fallback?.textColor ?? theme.colors.tx1 ?? "#373A36",
      fontFamily: resolvedTypeface,
      fontSize: rawSize === undefined ? fallback?.fontSize ?? 18 : Math.max(8, numeric(rawSize, 1800) / 100),
      fontWeight: rawBold === undefined ? fallback?.fontWeight ?? 400 : String(rawBold) === "1" ? 700 : 400,
      textAlign: alignment === "ctr" ? "center" : alignment === "r" ? "right" : "left",
      verticalAlign: anchor === "ctr" ? "center" : anchor === "b" ? "bottom" : "top",
      opacity: 1,
    },
  ];
}

function graphicFrameName(frame: XmlRecord): string {
  return String(record(record(frame.nvGraphicFramePr).cNvPr)["@name"] ?? "Graphic frame");
}

function graphicFrameElements(frame: XmlRecord, context: GroupContext, theme: Theme, index: number): TemplatePreviewElement[] {
  const transform = transformFor(frame.xfrm);
  if (!transform) return [];
  const mapped = mapTransform(transform, context);
  const data = record(record(frame.graphic).graphicData);
  const table = record(data.tbl);
  const frameId = String(record(record(frame.nvGraphicFramePr).cNvPr)["@id"] ?? index);
  if (Object.keys(table).length > 0) {
    const widths = array<XmlRecord>(record(table.tblGrid).gridCol).map((column) => Math.max(1, numeric(record(column)["@w"], 1)));
    const rows = array<XmlRecord>(table.tr);
    if (widths.length === 0 || rows.length === 0) return [];
    const totalWidth = widths.reduce((sum, value) => sum + value, 0);
    const rowHeights = rows.map((row) => Math.max(1, numeric(record(row)["@h"], 1)));
    const totalHeight = rowHeights.reduce((sum, value) => sum + value, 0);
    const columnOffsets = widths.reduce<number[]>((offsets, value) => [...offsets, offsets[offsets.length - 1] + value], [0]);
    const rowOffsets = rowHeights.reduce<number[]>((offsets, value) => [...offsets, offsets[offsets.length - 1] + value], [0]);
    const elements: TemplatePreviewElement[] = [];
    rows.forEach((rowValue, rowIndex) => {
      let columnIndex = 0;
      array<XmlRecord>(record(rowValue).tc).forEach((cellValue, cellIndex) => {
        const cell = record(cellValue);
        const properties = record(cell.tcPr);
        const span = Math.max(1, numeric(properties["@gridSpan"], 1));
        const x = mapped.x + mapped.width * (columnOffsets[columnIndex] / totalWidth);
        const y = mapped.y + mapped.height * (rowOffsets[rowIndex] / totalHeight);
        const width = mapped.width * ((columnOffsets[Math.min(widths.length, columnIndex + span)] - columnOffsets[columnIndex]) / totalWidth);
        const height = mapped.height * (rowHeights[rowIndex] / totalHeight);
        const cellFill = fillFor(properties, theme).color ?? "#FFFFFF";
        const cellId = `table-${index}-${frameId}-${rowIndex}-${cellIndex}`;
        elements.push({ id: `${cellId}-cell`, kind: "shape", name: `${graphicFrameName(frame)} cell`, sourceShapeId: frameId, x, y, width, height, rotation: 0, geometry: "rect", fill: cellFill, stroke: "#DBDCDB", strokeWidth: 9000, opacity: 1 });
        const txBody = record(cell.txBody);
        const text = paragraphText(txBody);
        if (text) {
          const run = firstRunProperties(txBody);
          const runFill = run.solidFill !== undefined ? schemeColor(record(run.solidFill), theme) : {};
          const typeface = String(record(run.latin)["@typeface"] ?? "");
          const resolvedTypeface = typeface === "+mj-lt" ? theme.majorFont : typeface === "+mn-lt" || !typeface ? theme.minorFont : typeface;
          const marginLeft = numeric(properties["@marL"], 72000);
          const marginRight = numeric(properties["@marR"], 72000);
          const marginTop = numeric(properties["@marT"], 36000);
          const marginBottom = numeric(properties["@marB"], 36000);
          elements.push({ id: `${cellId}-text`, kind: "text", name: `${graphicFrameName(frame)} cell text`, sourceShapeId: frameId, x: x + marginLeft, y: y + marginTop, width: Math.max(1, width - marginLeft - marginRight), height: Math.max(1, height - marginTop - marginBottom), rotation: 0, geometry: "rect", text, textColor: runFill.color ?? theme.colors.tx1 ?? "#373A36", fontFamily: resolvedTypeface, fontSize: Math.max(7, numeric(run["@sz"], 1400) / 100), fontWeight: String(run["@b"] ?? "0") === "1" || rowIndex === 0 ? 700 : 400, textAlign: "left", verticalAlign: "center", opacity: 1 });
        }
        columnIndex += span;
      });
    });
    return elements;
  }
  const kind = data.chart !== undefined ? "Chart" : data.diagram !== undefined || data.relIds !== undefined ? "Diagram" : "Embedded object";
  return [
    { id: `graphic-${index}-${frameId}`, kind: "shape", name: graphicFrameName(frame), sourceShapeId: frameId, ...mapped, geometry: "rect", fill: "#F4F6F5", stroke: "#A8B5AE", strokeWidth: 12700, opacity: 1 },
    { id: `graphic-${index}-${frameId}-text`, kind: "text", name: `${kind} label`, sourceShapeId: frameId, ...mapped, geometry: "rect", text: kind, textColor: "#68736E", fontFamily: theme.minorFont, fontSize: 14, fontWeight: 700, textAlign: "center", verticalAlign: "center", opacity: 1 },
  ];
}

function connectorElements(connector: XmlRecord, context: GroupContext, theme: Theme, index: number): TemplatePreviewElement[] {
  const spPr = record(connector.spPr);
  const transform = transformFor(spPr.xfrm);
  if (!transform) return [];
  const line = record(spPr.ln);
  const stroke = fillFor(line, theme).color ?? theme.colors.tx1 ?? "#373A36";
  const sourceShapeId = String(record(record(connector.nvCxnSpPr).cNvPr)["@id"] ?? index);
  return [{ id: `connector-${index}`, kind: "shape", name: "Connector", sourceShapeId, ...mapTransform(transform, context), geometry: "line", stroke, strokeWidth: numeric(line["@w"], 12700), opacity: 1 }];
}

function pictureElements(picture: XmlRecord, context: GroupContext, relationships: Relationship[], index: number): TemplatePreviewElement[] {
  const transform = transformFor(record(picture.spPr).xfrm);
  if (!transform) return [];
  const blipFill = record(picture.blipFill);
  const blip = record(blipFill.blip);
  const sourceCrop = record(blipFill.srcRect);
  const sourceCropped = ["@l", "@r", "@t", "@b"].some((attribute) => numeric(sourceCrop[attribute], 0) !== 0);
  const svgRelationshipId = array<XmlRecord>(record(blip.extLst).ext)
    .map((extension) => String(record(extension.svgBlip)["@embed"] ?? ""))
    .find(Boolean);
  const relationshipId = String(blip["@embed"] ?? svgRelationshipId ?? "");
  const relationship = findRelationship(relationships, relationshipId);
  if (!relationship) return [];
  const sourceShapeId = String(record(record(picture.nvPicPr).cNvPr)["@id"] ?? index);
  return [{
    id: `picture-${index}-${sourceShapeId}`,
    kind: "image",
    name: shapeName(picture, true),
    sourceShapeId,
    ...mapTransform(transform, context),
    geometry: "rect",
    mediaId: relationship.target,
    sourceCropped,
  }];
}

async function enrichTextMetadata(elements: TemplatePreviewElement[]): Promise<TemplatePreviewElement[]> {
  return Promise.all(elements.map(async (element) => {
    if (element.kind !== "text" || !element.text?.trim()) return element;
    const paragraphs = element.text.split(/\r?\n/).map((text) => text.replace(/\s+/g, " ").trim()).filter(Boolean);
    return {
      ...element,
      textHash: await sha256Text(element.text.replace(/\s+/g, " ").trim()),
      sourceParagraphs: await Promise.all(paragraphs.map(async (text, index) => ({
        index: index + 1,
        text,
        textHash: await sha256Text(text),
        characterCount: text.length,
        bullet: false,
        bulletConfidence: "inherited-possible" as const,
        level: 0,
        fontFamilies: element.fontFamily ? [element.fontFamily] : [],
        fontSizes: element.fontSize ? [element.fontSize] : [],
      }))),
    };
  }));
}

function flattenTree(spTree: XmlRecord, context: GroupContext, theme: Theme, relationships: Relationship[], prefix = "root"): TemplatePreviewElement[] {
  const result: TemplatePreviewElement[] = [];
  array<XmlRecord>(spTree.pic).forEach((picture, index) => result.push(...pictureElements(record(picture), context, relationships, index)));
  array<XmlRecord>(spTree.sp).forEach((shape, index) => result.push(...shapeElements(record(shape), context, theme, index)));
  array<XmlRecord>(spTree.graphicFrame).forEach((frame, index) => result.push(...graphicFrameElements(record(frame), context, theme, index)));
  array<XmlRecord>(spTree.cxnSp).forEach((connector, index) => result.push(...connectorElements(record(connector), context, theme, index)));
  array<XmlRecord>(spTree.grpSp).forEach((groupValue, index) => {
    const group = record(groupValue);
    const xfrm = record(record(group.grpSpPr).xfrm);
    const transform = transformFor(xfrm);
    if (!transform) return;
    const mapped = mapTransform(transform, context);
    const childOff = record(xfrm.chOff);
    const childExt = record(xfrm.chExt);
    const childContext: GroupContext = {
      x: mapped.x,
      y: mapped.y,
      width: mapped.width,
      height: mapped.height,
      childX: numeric(childOff["@x"]),
      childY: numeric(childOff["@y"]),
      childWidth: Math.max(1, numeric(childExt["@cx"], mapped.width)),
      childHeight: Math.max(1, numeric(childExt["@cy"], mapped.height)),
    };
    result.push(...flattenTree(group, childContext, theme, relationships, `${prefix}-group-${index}`));
  });
  return result.map((element, index) => ({ ...element, id: `${prefix}-${index}-${element.id}` }));
}

function backgroundFor(commonSlideData: XmlRecord, theme: Theme): string {
  const background = record(commonSlideData.bg);
  const backgroundProperties = record(background.bgPr);
  const direct = fillFor(backgroundProperties, theme).color;
  if (direct) return direct;
  const reference = record(background.bgRef);
  return schemeColor(reference, theme).color ?? theme.colors.bg1 ?? "#FFFFFF";
}

function categoryFor(name: string): TemplateLayoutPreview["category"] {
  const normalized = name.toLowerCase();
  if (normalized.includes("conclusion") || normalized.includes("closing") || normalized.includes("thank")) return "conclusion";
  if (normalized.includes("title")) return "title";
  if (normalized.includes("image") || normalized.includes("photo") || normalized.includes("portrait")) return "image";
  if (normalized.includes("column") || normalized.includes("content") || normalized.includes("text") || normalized.includes("green bar")) return "content";
  return "other";
}

function mediaTypeFor(part: string): string | undefined {
  const extension = part.split(".").pop()?.toLowerCase();
  return ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", svg: "image/svg+xml", webp: "image/webp" } as Record<string, string>)[extension ?? ""];
}

function sortPartNames(left: string, right: string): number {
  const leftNumber = numeric(left.match(/(\d+)(?=\.xml$)/)?.[1]);
  const rightNumber = numeric(right.match(/(\d+)(?=\.xml$)/)?.[1]);
  return leftNumber - rightNumber || left.localeCompare(right);
}

export async function buildTemplateCatalog(bytes: Uint8Array, sourceName: string): Promise<TemplateCatalog> {
  if (bytes.byteLength < 100 || bytes.byteLength > 250 * 1024 * 1024) throw new Error("Choose a valid PowerPoint template smaller than 250 MB.");
  if (!/\.(potx|pptx)$/i.test(sourceName)) throw new Error("Template designs require a POTX or PPTX file.");
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const layoutParts = Object.keys(zip.files).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)).sort(sortPartNames);
  const masterParts = Object.keys(zip.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)).sort(sortPartNames);
  if (layoutParts.length === 0 || layoutParts.length > 200 || masterParts.length === 0) throw new Error("The selected file does not contain a supported PowerPoint master and layout catalog.");

  const presentation = record((await readXml(zip, "ppt/presentation.xml")).presentation);
  const slideSize = record(presentation.sldSz);
  const slideWidth = numeric(slideSize["@cx"]);
  const slideHeight = numeric(slideSize["@cy"]);
  if (slideWidth <= 0 || slideHeight <= 0) throw new Error("The template slide dimensions are invalid.");

  const defaultThemePart = Object.keys(zip.files).find((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name));
  const defaultTheme = defaultThemePart ? parseTheme(await readXml(zip, defaultThemePart)) : { colors: { tx1: "#373A36", bg1: "#FFFFFF", accent1: "#00662C" }, majorFont: "Aptos Display", minorFont: "Aptos" };
  const masterCache = new Map<string, { elements: TemplatePreviewElement[]; background: string; theme: Theme; relationships: Relationship[] }>();

  async function loadMaster(part: string) {
    const cached = masterCache.get(part);
    if (cached) return cached;
    const relationships = await readRelationships(zip, part);
    const themeRelationship = relationships.find((item) => item.type.endsWith("/theme"));
    const theme = themeRelationship && zip.file(themeRelationship.target) ? parseTheme(await readXml(zip, themeRelationship.target)) : defaultTheme;
    const masterXml = record((await readXml(zip, part)).sldMaster);
    const commonSlideData = record(masterXml.cSld);
    const context: GroupContext = { x: 0, y: 0, width: slideWidth, height: slideHeight, childX: 0, childY: 0, childWidth: slideWidth, childHeight: slideHeight };
    const loaded = {
      elements: (await enrichTextMetadata(flattenTree(record(commonSlideData.spTree), context, theme, relationships, part))).filter((element) => !element.placeholderType).map((element) => ({ ...element, sourcePart: part, origin: "master" as const })),
      background: backgroundFor(commonSlideData, theme),
      theme,
      relationships,
    };
    masterCache.set(part, loaded);
    return loaded;
  }

  const layouts: TemplateLayoutPreview[] = [];
  const mediaParts = new Set<string>();
  for (const [index, layoutPart] of layoutParts.entries()) {
    const relationships = await readRelationships(zip, layoutPart);
    const masterRelationship = relationships.find((item) => item.type.endsWith("/slideMaster"));
    const masterPart = masterRelationship?.target ?? masterParts[0];
    const master = await loadMaster(masterPart);
    const layoutXml = record((await readXml(zip, layoutPart)).sldLayout);
    const commonSlideData = record(layoutXml.cSld);
    const context: GroupContext = { x: 0, y: 0, width: slideWidth, height: slideHeight, childX: 0, childY: 0, childWidth: slideWidth, childHeight: slideHeight };
    const layoutElements = (await enrichTextMetadata(flattenTree(record(commonSlideData.spTree), context, master.theme, relationships, layoutPart))).map((element) => ({ ...element, sourcePart: layoutPart, origin: "layout" as const }));
    const showMasterShapes = String(layoutXml["@showMasterSp"] ?? "1") !== "0";
    const elements = [...(showMasterShapes ? master.elements : []), ...layoutElements];
    for (const element of elements) if (element.mediaId) mediaParts.add(element.mediaId);
    const name = String(commonSlideData["@name"] ?? `Layout ${index + 1}`).trim() || `Layout ${index + 1}`;
    const preview = {
      id: `layout-${index + 1}`,
      name,
      category: categoryFor(name),
      background: backgroundFor(commonSlideData, master.theme) || master.background,
      elements,
      placeholderTypes: [...new Set(elements.map((element) => element.placeholderType).filter((value): value is string => Boolean(value)))],
      sourcePart: layoutPart,
    };
    layouts.push({ ...preview, semantic: deriveLayoutSemantics(preview, slideWidth, slideHeight) });
  }

  const media: Record<string, string> = {};
  let packagedMediaBytes = 0;
  for (const part of mediaParts) {
    const type = mediaTypeFor(part);
    const file = zip.file(part);
    if (!type || !file) continue;
    const data = await file.async("uint8array");
    packagedMediaBytes += data.byteLength;
    if (packagedMediaBytes > 150 * 1024 * 1024) throw new Error("The template media required for previews exceeds the 150 MB safety limit.");
    media[part] = `data:${type};base64,${await file.async("base64")}`;
  }

  const digest = await sha256(bytes);
  return {
    id: `template-${digest.slice(0, 16)}`,
    name: sourceName,
    sha256: digest,
    slideWidth,
    slideHeight,
    masterCount: masterParts.length,
    layouts,
    media,
    generatedAt: new Date().toISOString(),
  };
}

interface LoadedSlideLayout {
  elements: TemplatePreviewElement[];
  placeholders: TemplatePreviewElement[];
  background: string;
  theme: Theme;
}

function placeholderMatches(element: TemplatePreviewElement, info: { type?: string; index?: string }): boolean {
  if (!element.placeholderType || !info.type) return false;
  if (info.index !== undefined && element.placeholderIndex !== undefined) return info.index === element.placeholderIndex;
  return element.placeholderType === info.type;
}

function flattenSlideTree(spTree: XmlRecord, context: GroupContext, theme: Theme, relationships: Relationship[], placeholders: TemplatePreviewElement[], prefix: string): TemplatePreviewElement[] {
  const result: TemplatePreviewElement[] = [];
  const used = new Set<string>();
  array<XmlRecord>(spTree.pic).forEach((picture, index) => result.push(...pictureElements(record(picture), context, relationships, index)));
  array<XmlRecord>(spTree.graphicFrame).forEach((frame, index) => result.push(...graphicFrameElements(record(frame), context, theme, index)));
  array<XmlRecord>(spTree.cxnSp).forEach((connector, index) => result.push(...connectorElements(record(connector), context, theme, index)));
  array<XmlRecord>(spTree.sp).forEach((shapeValue, index) => {
    const shape = record(shapeValue);
    const info = placeholderInfo(shape);
    const fallback = placeholders.find((element) => element.kind === "text" && !used.has(element.id) && placeholderMatches(element, info));
    if (fallback) used.add(fallback.id);
    result.push(...shapeElements(shape, context, theme, index, fallback));
  });
  array<XmlRecord>(spTree.grpSp).forEach((groupValue, index) => {
    const group = record(groupValue);
    const xfrm = record(record(group.grpSpPr).xfrm);
    const transform = transformFor(xfrm);
    if (!transform) return;
    const mapped = mapTransform(transform, context);
    const childOff = record(xfrm.chOff);
    const childExt = record(xfrm.chExt);
    result.push(...flattenTree(group, { x: mapped.x, y: mapped.y, width: mapped.width, height: mapped.height, childX: numeric(childOff["@x"]), childY: numeric(childOff["@y"]), childWidth: Math.max(1, numeric(childExt["@cx"], mapped.width)), childHeight: Math.max(1, numeric(childExt["@cy"], mapped.height)) }, theme, relationships, `${prefix}-group-${index}`));
  });
  return result.map((element, index) => ({ ...element, id: `${prefix}-${index}-${element.id}` }));
}

export async function buildSlideRenderCatalog(bytes: Uint8Array, sourceName: string): Promise<SlideRenderCatalog> {
  if (bytes.byteLength < 100 || bytes.byteLength > 1_000_000_000) throw new Error("Choose a valid PowerPoint presentation smaller than 1 GB.");
  if (!/\.pptx$/i.test(sourceName)) throw new Error("Slide previews require a PPTX file.");
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  const slideParts = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).sort(sortPartNames);
  const layoutParts = Object.keys(zip.files).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)).sort(sortPartNames);
  const masterParts = Object.keys(zip.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)).sort(sortPartNames);
  if (slideParts.length === 0 || layoutParts.length === 0 || masterParts.length === 0) throw new Error("The PowerPoint file does not contain supported slides, masters, and layouts.");

  const presentation = record((await readXml(zip, "ppt/presentation.xml")).presentation);
  const slideSize = record(presentation.sldSz);
  const slideWidth = numeric(slideSize["@cx"]);
  const slideHeight = numeric(slideSize["@cy"]);
  if (slideWidth <= 0 || slideHeight <= 0) throw new Error("The presentation slide dimensions are invalid.");
  const rootContext: GroupContext = { x: 0, y: 0, width: slideWidth, height: slideHeight, childX: 0, childY: 0, childWidth: slideWidth, childHeight: slideHeight };
  const defaultThemePart = Object.keys(zip.files).find((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name));
  const defaultTheme = defaultThemePart ? parseTheme(await readXml(zip, defaultThemePart)) : { colors: { tx1: "#373A36", bg1: "#FFFFFF", accent1: "#00662C" }, majorFont: "Aptos Display", minorFont: "Aptos" };
  const masterCache = new Map<string, { elements: TemplatePreviewElement[]; background: string; theme: Theme }>();
  const layoutCache = new Map<string, LoadedSlideLayout>();

  async function loadMaster(part: string) {
    const cached = masterCache.get(part);
    if (cached) return cached;
    const relationships = await readRelationships(zip, part);
    const themeRelationship = relationships.find((item) => item.type.endsWith("/theme"));
    const theme = themeRelationship && zip.file(themeRelationship.target) ? parseTheme(await readXml(zip, themeRelationship.target)) : defaultTheme;
    const masterXml = record((await readXml(zip, part)).sldMaster);
    const common = record(masterXml.cSld);
    const loaded = { elements: (await enrichTextMetadata(flattenTree(record(common.spTree), rootContext, theme, relationships, part))).filter((element) => !element.placeholderType).map((element) => ({ ...element, sourcePart: part, origin: "master" as const })), background: backgroundFor(common, theme), theme };
    masterCache.set(part, loaded);
    return loaded;
  }

  async function loadLayout(part: string): Promise<LoadedSlideLayout> {
    const cached = layoutCache.get(part);
    if (cached) return cached;
    const relationships = await readRelationships(zip, part);
    const masterPart = relationships.find((item) => item.type.endsWith("/slideMaster"))?.target ?? masterParts[0];
    const master = await loadMaster(masterPart);
    const layoutXml = record((await readXml(zip, part)).sldLayout);
    const common = record(layoutXml.cSld);
    const layoutElements = (await enrichTextMetadata(flattenTree(record(common.spTree), rootContext, master.theme, relationships, part))).map((element) => ({ ...element, sourcePart: part, origin: "layout" as const }));
    const placeholders = layoutElements.filter((element) => Boolean(element.placeholderType));
    const showMasterShapes = String(layoutXml["@showMasterSp"] ?? "1") !== "0";
    const loaded = { elements: [...(showMasterShapes ? master.elements : []), ...layoutElements.filter((element) => !element.placeholderType)], placeholders, background: Object.keys(record(common.bg)).length ? backgroundFor(common, master.theme) : master.background, theme: master.theme };
    layoutCache.set(part, loaded);
    return loaded;
  }

  const slides: SlideRenderPreview[] = [];
  const mediaParts = new Set<string>();
  for (const [index, slidePart] of slideParts.entries()) {
    const relationships = await readRelationships(zip, slidePart);
    const layoutPart = relationships.find((item) => item.type.endsWith("/slideLayout"))?.target ?? layoutParts[0];
    const layout = await loadLayout(layoutPart);
    const slideXml = record((await readXml(zip, slidePart)).sld);
    const common = record(slideXml.cSld);
    const slideElements = (await enrichTextMetadata(flattenSlideTree(record(common.spTree), rootContext, layout.theme, relationships, layout.placeholders, slidePart))).map((element) => ({ ...element, sourcePart: slidePart, origin: "slide" as const }));
    const elements = [...layout.elements, ...slideElements];
    for (const element of elements) if (element.mediaId) mediaParts.add(element.mediaId);
    const titleElement = slideElements.find((element) => element.kind === "text" && ["title", "ctrTitle"].includes(element.placeholderType ?? "") && element.text?.trim()) ?? slideElements.find((element) => element.kind === "text" && element.text?.trim());
    const title = titleElement?.text?.split(/\r?\n/)[0].trim().slice(0, 160) || `Slide ${index + 1}`;
    const warnings: string[] = [];
    if (slideElements.length === 0) warnings.push("No supported editable slide elements were available to the local preview renderer.");
    if (elements.some((element) => element.kind === "text" && !element.fontFamily)) warnings.push("Some text uses inherited formatting that may differ in native PowerPoint.");
    const preview = { id: `slide-render-${index + 1}`, name: title, category: "content" as const, background: Object.keys(record(common.bg)).length ? backgroundFor(common, layout.theme) : layout.background, elements, placeholderTypes: [], sourcePart: slidePart };
    slides.push({ ...preview, number: index + 1, title, hidden: String(slideXml["@show"] ?? "1") === "0", renderWarnings: warnings, semantic: deriveLayoutSemantics(preview, slideWidth, slideHeight) });
  }

  const media: Record<string, string> = {};
  let packagedMediaBytes = 0;
  for (const part of mediaParts) {
    const type = mediaTypeFor(part);
    const file = zip.file(part);
    if (!type || !file) continue;
    const data = await file.async("uint8array");
    packagedMediaBytes += data.byteLength;
    if (packagedMediaBytes > 200 * 1024 * 1024) throw new Error("The slide media required for local previews exceeds the 200 MB safety limit.");
    media[part] = `data:${type};base64,${await file.async("base64")}`;
  }

  const digest = await sha256(bytes);
  return { id: `slide-catalog-${digest.slice(0, 16)}`, name: sourceName, sha256: digest, slideWidth, slideHeight, slides, media, generatedAt: new Date().toISOString(), renderer: "local-ooxml-preview" };
}
