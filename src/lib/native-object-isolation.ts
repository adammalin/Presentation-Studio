import JSZip from "jszip";
import type { StudioFigureTreatment, StudioWebFrame, StudioWebSlide } from "../types";

const SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SHAPE_TAGS = new Set(["p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp"]);
const EMU_PER_INCH = 914_400;

function frameContains(container: StudioWebFrame, candidate: StudioWebFrame, padding = .12 * EMU_PER_INCH): boolean {
  const centerX = candidate.x + candidate.width / 2;
  const centerY = candidate.y + candidate.height / 2;
  return centerX >= container.x - padding
    && centerX <= container.x + container.width + padding
    && centerY >= container.y - padding
    && centerY <= container.y + container.height + padding;
}

/**
 * Returns top-level source shape IDs only when every catalog-derived member of
 * the requested evidence unit is contained by a selected native PowerPoint
 * group. Otherwise the caller must crop the authoritative full-slide render;
 * isolating only the known top-level shapes would silently drop part of a
 * composite technical figure.
 */
export function nativeIsolationShapeIds(slide: StudioWebSlide, treatment: StudioFigureTreatment): string[] {
  const nodes = treatment.nodeIds.map((id) => slide.nodes.find((node) => node.id === id)).filter((node): node is StudioWebSlide["nodes"][number] => Boolean(node?.visible));
  const editableNodes = nodes.filter((node) => node.sourceBinding === "editable-object" && Boolean(node.sourceShapeId));
  if (!editableNodes.length) return [];
  const nativeContainers = editableNodes.filter((node) => node.kind === "native-object");
  const unresolved = nodes.filter((node) => node.sourceBinding !== "editable-object");
  if (unresolved.some((node) => !nativeContainers.some((container) => frameContains(container.sourceFrame, node.sourceFrame)))) return [];
  return [...new Set(editableNodes.map((node) => node.sourceShapeId))];
}

function decodeXml(value: string): string {
  return value.replace(/&(amp|lt|gt|quot|apos);/g, (_entity, name: string) => name === "amp" ? "&" : name === "lt" ? "<" : name === "gt" ? ">" : name === "quot" ? '"' : "'");
}

function attribute(attributes: string, name: string): string | undefined {
  return decodeXml(attributes.match(new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "i"))?.slice(1).find((value) => value !== undefined) ?? "");
}

function setHidden(attributes: string, hidden: boolean): string {
  const pattern = /(\bhidden=)(?:"[^"]*"|'[^']*')/i;
  if (!hidden) return attributes.replace(pattern, "").replace(/\s{2,}/g, " ");
  return pattern.test(attributes) ? attributes.replace(pattern, '$1"1"') : `${attributes.trimEnd()} hidden="1"`;
}

function isolateTopLevelShapes(slideXml: string, allowedShapeIds: Set<string>): { xml: string; hiddenShapeIds: string[]; preservedShapeIds: string[] } {
  const shapeTreeStart = slideXml.search(/<p:spTree\b/i);
  if (shapeTreeStart < 0) throw new Error("The source slide does not contain a PowerPoint shape tree.");
  const tokenPattern = /<(\/)?([A-Za-z_][\w.-]*(?::[\w.-]+)?)(\s[^<>]*?)?(\/?)>/g;
  tokenPattern.lastIndex = shapeTreeStart;
  const stack: Array<{ name: string; start: number; topLevelShape: boolean }> = [];
  const ranges: Array<{ start: number; end: number }> = [];
  for (let match = tokenPattern.exec(slideXml); match; match = tokenPattern.exec(slideXml)) {
    const closing = match[1] === "/";
    const name = match[2];
    const selfClosing = match[4] === "/";
    if (!closing) {
      const topLevelShape = stack.at(-1)?.name === "p:spTree" && SHAPE_TAGS.has(name);
      if (!selfClosing) stack.push({ name, start: match.index, topLevelShape });
      if (name === "p:spTree" && stack.length > 1) throw new Error("The source slide contains an unexpected nested shape tree.");
      continue;
    }
    const open = stack.pop();
    if (!open || open.name !== name) throw new Error("The source slide contains malformed shape XML.");
    if (open.topLevelShape) ranges.push({ start: open.start, end: tokenPattern.lastIndex });
    if (name === "p:spTree") break;
  }
  if (!ranges.length) throw new Error("The source slide does not contain isolatable top-level shapes.");
  const hiddenShapeIds: string[] = [];
  const preservedShapeIds: string[] = [];
  let xml = slideXml;
  for (const range of [...ranges].sort((left, right) => right.start - left.start)) {
    const shape = xml.slice(range.start, range.end);
    const next = shape.replace(/<p:cNvPr\b([^>]*)>/i, (tag, attributes: string) => {
      const id = attribute(attributes, "id");
      if (!id) throw new Error("A top-level PowerPoint shape is missing its nonvisual ID.");
      const keep = allowedShapeIds.has(id);
      (keep ? preservedShapeIds : hiddenShapeIds).push(id);
      return `<p:cNvPr${setHidden(attributes, !keep)}>`;
    });
    xml = `${xml.slice(0, range.start)}${next}${xml.slice(range.end)}`;
  }
  const missing = [...allowedShapeIds].filter((id) => !preservedShapeIds.includes(id));
  if (missing.length) throw new Error(`The requested source shape IDs are not top-level objects on this slide: ${missing.join(", ")}.`);
  return { xml, hiddenShapeIds: hiddenShapeIds.reverse(), preservedShapeIds: preservedShapeIds.reverse() };
}

function slideRelationshipId(relationships: string, slideNumber: number): string {
  const target = `slides/slide${slideNumber}.xml`;
  for (const match of relationships.matchAll(/<Relationship\b([^>]*?)\/?\s*>/g)) {
    const attributes = match[1] ?? "";
    if (attribute(attributes, "Type") === SLIDE_RELATIONSHIP && attribute(attributes, "Target")?.replace(/^\.\//, "") === target) {
      const id = attribute(attributes, "Id");
      if (id) return id;
    }
  }
  throw new Error(`The PowerPoint package does not list slide ${slideNumber} in its presentation relationships.`);
}

function keepOnlySlide(presentationXml: string, relationshipId: string): string {
  return presentationXml.replace(/<p:sldIdLst\b([^>]*)>([\s\S]*?)<\/p:sldIdLst>/i, (_whole, attributes: string, contents: string) => {
    const entries = [...contents.matchAll(/<p:sldId\b[^>]*\br:id=(?:"([^"]*)"|'([^']*)')[^>]*\/?\s*>/gi)];
    const selected = entries.find((match) => (match[1] ?? match[2]) === relationshipId)?.[0];
    if (!selected) throw new Error("The requested source slide is not present in the presentation slide list.");
    return `<p:sldIdLst${attributes}>${selected}</p:sldIdLst>`;
  });
}

export interface NativeObjectIsolationReceipt {
  slideNumber: number;
  preservedShapeIds: string[];
  hiddenShapeIds: string[];
}

/**
 * Produces a private, one-visible-slide PowerPoint render source in which only
 * the requested top-level source shapes remain visible. The original package
 * and every meaning-bearing child inside a preserved group remain untouched.
 */
export async function isolateNativePowerPointObjects(input: { sourceBytes: Uint8Array; slideNumber: number; shapeIds: string[] }): Promise<{ bytes: Uint8Array; receipt: NativeObjectIsolationReceipt }> {
  const shapeIds = [...new Set(input.shapeIds.map(String).filter(Boolean))];
  if (!Number.isInteger(input.slideNumber) || input.slideNumber < 1) throw new Error("Native object isolation requires a positive source slide number.");
  if (!shapeIds.length) throw new Error("Native object isolation requires at least one top-level source shape ID.");
  const zip = await JSZip.loadAsync(input.sourceBytes, { checkCRC32: false });
  const slidePart = `ppt/slides/slide${input.slideNumber}.xml`;
  const slideEntry = zip.file(slidePart);
  const presentationEntry = zip.file("ppt/presentation.xml");
  const relationshipsEntry = zip.file("ppt/_rels/presentation.xml.rels");
  if (!slideEntry || !presentationEntry || !relationshipsEntry) throw new Error("The PowerPoint package is missing the requested slide or presentation metadata.");
  const isolated = isolateTopLevelShapes(await slideEntry.async("text"), new Set(shapeIds));
  const relationshipId = slideRelationshipId(await relationshipsEntry.async("text"), input.slideNumber);
  zip.file(slidePart, isolated.xml);
  zip.file("ppt/presentation.xml", keepOnlySlide(await presentationEntry.async("text"), relationshipId));
  const application = zip.file("docProps/app.xml");
  if (application) zip.file("docProps/app.xml", (await application.async("text")).replace(/<Slides>\d+<\/Slides>/i, "<Slides>1</Slides>"));
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { bytes, receipt: { slideNumber: input.slideNumber, preservedShapeIds: isolated.preservedShapeIds, hiddenShapeIds: isolated.hiddenShapeIds } };
}
