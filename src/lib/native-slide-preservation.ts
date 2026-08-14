import JSZip from "jszip";
import { cloneTemplateLayoutForSlide } from "./native-layout-remap";
import { sha256 } from "./hash";

const SLIDE_LAYOUT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

function decodeXml(value: string): string {
  return value.replace(/&(amp|lt|gt|quot|apos);/g, (_entity, name: string) => name === "amp" ? "&" : name === "lt" ? "<" : name === "gt" ? ">" : name === "quot" ? '"' : "'");
}

function escapeXml(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function attributeValue(attributes: string, name: string): string | undefined {
  return attributes.match(new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "i"))?.slice(1).find((value) => value !== undefined);
}

function setAttribute(attributes: string, name: string, value: string): string {
  const escaped = escapeXml(value);
  const pattern = new RegExp(`(\\b${name}=)(?:"[^"]*"|'[^']*')`, "i");
  return pattern.test(attributes) ? attributes.replace(pattern, `$1"${escaped}"`) : `${attributes.trimEnd()} ${name}="${escaped}"`;
}

function normalizePart(value: string): string {
  const result: string[] = [];
  for (const segment of value.replaceAll("\\", "/").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") result.pop();
    else result.push(segment);
  }
  return result.join("/");
}

function dirname(part: string): string { return part.slice(0, Math.max(0, part.lastIndexOf("/"))); }
function basename(part: string): string { return part.slice(part.lastIndexOf("/") + 1); }
function relationshipPart(part: string): string { return `${dirname(part)}/_rels/${basename(part)}.rels`; }
function resolveTarget(sourcePart: string, target: string): string { return normalizePart(`${dirname(sourcePart)}/${decodeXml(target)}`); }

function relativeTarget(sourcePart: string, targetPart: string): string {
  const source = dirname(sourcePart).split("/").filter(Boolean);
  const target = targetPart.split("/").filter(Boolean);
  let shared = 0;
  while (shared < source.length && shared < target.length && source[shared] === target[shared]) shared += 1;
  return [...Array(source.length - shared).fill(".."), ...target.slice(shared)].join("/") || basename(targetPart);
}

async function text(zip: JSZip, part: string): Promise<string> {
  const entry = zip.file(part);
  if (!entry) throw new Error(`Required PowerPoint package part ${part} is missing.`);
  return entry.async("text");
}

function layoutTarget(slidePart: string, relationships: string): string {
  for (const match of relationships.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributes = match[1] ?? "";
    if (decodeXml(attributeValue(attributes, "Type") ?? "") === SLIDE_LAYOUT_TYPE) {
      return resolveTarget(slidePart, attributeValue(attributes, "Target") ?? "");
    }
  }
  throw new Error("The source title slide does not have one native layout relationship.");
}

async function masterTarget(zip: JSZip, layoutPart: string): Promise<string | undefined> {
  const relationshipsEntry = zip.file(relationshipPart(layoutPart));
  if (!relationshipsEntry) return undefined;
  const relationships = await relationshipsEntry.async("text");
  for (const match of relationships.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributes = match[1] ?? "";
    if (decodeXml(attributeValue(attributes, "Type") ?? "") === SLIDE_MASTER_TYPE) {
      return resolveTarget(layoutPart, attributeValue(attributes, "Target") ?? "");
    }
  }
  return undefined;
}

function addMediaContentType(destinationXml: string, sourceXml: string, sourcePart: string): string {
  const extension = basename(sourcePart).match(/\.([^.]+)$/)?.[1];
  if (!extension || new RegExp(`<Default\\b[^>]*\\bExtension=(?:"${extension}"|'${extension}')`, "i").test(destinationXml)) return destinationXml;
  const sourceDefault = sourceXml.match(new RegExp(`<Default\\b([^>]*\\bExtension=(?:"${extension}"|'${extension}')[^>]*)/>`, "i"));
  return sourceDefault ? destinationXml.replace(/<\/Types>\s*$/i, `<Default${sourceDefault[1]}/></Types>`) : destinationXml;
}

export interface NativeSlidePreservationReceipt {
  slideNumber: number;
  sourceSlideSha256: string;
  clonedLayoutPart: string;
  clonedMasterPart?: string;
  copiedMediaCount: number;
}

/** Transplants one source slide and its native title-layout dependency graph. */
export async function preserveNativeSlide(input: { destinationBytes: Uint8Array; sourceBytes: Uint8Array; slideNumber: number }): Promise<{ bytes: Uint8Array; receipt: NativeSlidePreservationReceipt }> {
  const { destinationBytes, sourceBytes, slideNumber } = input;
  const source = await JSZip.loadAsync(sourceBytes, { checkCRC32: false });
  const slidePart = `ppt/slides/slide${slideNumber}.xml`;
  const sourceSlide = source.file(slidePart);
  const sourceRelationshipsPart = relationshipPart(slidePart);
  if (!sourceSlide || !source.file(sourceRelationshipsPart)) throw new Error(`The source PowerPoint is missing native slide ${slideNumber}.`);
  const sourceRelationships = await text(source, sourceRelationshipsPart);
  const sourceLayoutPart = layoutTarget(slidePart, sourceRelationships);
  const sourceLayoutBytes = await source.file(sourceLayoutPart)?.async("uint8array");
  if (!sourceLayoutBytes) throw new Error("The source title slide's native layout part is missing.");

  const cloned = await cloneTemplateLayoutForSlide({
    sourceBytes: destinationBytes,
    templateBytes: sourceBytes,
    command: {
      id: `preserve-native-slide-${slideNumber}`,
      slideNumber,
      templateSha256: await sha256(sourceBytes),
      templateLayoutPart: sourceLayoutPart,
      templateLayoutSha256: await sha256(sourceLayoutBytes),
      templateLayoutName: `Source-preserved slide ${slideNumber} layout`,
      rationale: "Preserve the populated ORNL title slide and its native template dependency graph exactly.",
      author: "ai",
    },
  });
  const destination = await JSZip.loadAsync(cloned.bytes, { checkCRC32: false });
  const preservedMasterPart = cloned.receipt.clonedMasterPart ?? await masterTarget(destination, cloned.receipt.clonedLayoutPart);
  if (!preservedMasterPart || !destination.file(preservedMasterPart)) throw new Error("The sacred title slide's preserved layout is not connected to a native slide master.");
  destination.file(slidePart, await sourceSlide.async("uint8array"));

  const sourceContentTypes = await text(source, "[Content_Types].xml");
  let destinationContentTypes = await text(destination, "[Content_Types].xml");
  let copiedMediaCount = 0;
  const rewrittenRelationships = await Promise.all([...sourceRelationships.matchAll(/<Relationship\b([^>]*?)\/>/g)].map(async (match) => {
    let attributes = match[1] ?? "";
    const type = decodeXml(attributeValue(attributes, "Type") ?? "");
    const target = attributeValue(attributes, "Target");
    const external = /^external$/i.test(attributeValue(attributes, "TargetMode") ?? "");
    if (!target || external) return `<Relationship${attributes}/>`;
    if (type === SLIDE_LAYOUT_TYPE) {
      attributes = setAttribute(attributes, "Target", relativeTarget(slidePart, cloned.receipt.clonedLayoutPart));
      return `<Relationship${attributes}/>`;
    }
    const sourceTargetPart = resolveTarget(slidePart, target);
    if (!/^ppt\/media\//i.test(sourceTargetPart)) throw new Error(`Sacred title-slide preservation does not yet support relationship ${sourceTargetPart}; export is held rather than flattening it.`);
    const media = await source.file(sourceTargetPart)?.async("uint8array");
    if (!media) throw new Error(`The sacred title slide is missing related media ${sourceTargetPart}.`);
    const extension = basename(sourceTargetPart).match(/\.([^.]+)$/)?.[1]?.toLowerCase() ?? "bin";
    const destinationPart = `ppt/media/source-title-${(await sha256(media)).slice(0, 12)}-${copiedMediaCount + 1}.${extension}`;
    destination.file(destinationPart, media);
    destinationContentTypes = addMediaContentType(destinationContentTypes, sourceContentTypes, sourceTargetPart);
    copiedMediaCount += 1;
    attributes = setAttribute(attributes, "Target", relativeTarget(slidePart, destinationPart));
    return `<Relationship${attributes}/>`;
  }));
  const relationshipRoot = sourceRelationships.match(/^\s*<\?xml[^>]*\?>\s*<Relationships\b([^>]*)>/i)?.[1] ?? ' xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  destination.file(sourceRelationshipsPart, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships${relationshipRoot}>${rewrittenRelationships.join("")}</Relationships>`);
  destination.file("[Content_Types].xml", destinationContentTypes);
  const bytes = await destination.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { bytes, receipt: { slideNumber, sourceSlideSha256: await sha256(await sourceSlide.async("uint8array")), clonedLayoutPart: cloned.receipt.clonedLayoutPart, clonedMasterPart: preservedMasterPart, copiedMediaCount } };
}
