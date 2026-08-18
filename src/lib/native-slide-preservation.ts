import JSZip from "jszip";
import { cloneTemplateLayoutForSlide } from "./native-layout-remap";
import { sha256 } from "./hash";

const SLIDE_LAYOUT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const NOTES_SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const SLIDE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

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
  throw new Error("The source slide does not have one native layout relationship.");
}

function relationshipOfType(sourcePart: string, relationships: string, expectedType: string): { attributes: string; targetPart: string } | undefined {
  for (const match of relationships.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributes = match[1] ?? "";
    if (decodeXml(attributeValue(attributes, "Type") ?? "") !== expectedType) continue;
    const target = attributeValue(attributes, "Target");
    if (!target || /^external$/i.test(attributeValue(attributes, "TargetMode") ?? "")) continue;
    return { attributes, targetPart: resolveTarget(sourcePart, target) };
  }
  return undefined;
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

function nextRelationshipId(used: Set<string>): string {
  let number = 1;
  while (used.has(`rId${number}`)) number += 1;
  const result = `rId${number}`;
  used.add(result);
  return result;
}

export interface NativeSlidePreservationReceipt {
  slideNumber: number;
  sourceSlideNumber: number;
  destinationSlideNumber: number;
  sourceSlideSha256: string;
  clonedLayoutPart: string;
  clonedMasterPart?: string;
  copiedMediaCount: number;
  preservedNotes: boolean;
}

/** Transplants one protected source slide and its native layout dependency graph. */
export async function preserveNativeSlide(input: { destinationBytes: Uint8Array; sourceBytes: Uint8Array; slideNumber?: number; sourceSlideNumber?: number; destinationSlideNumber?: number }): Promise<{ bytes: Uint8Array; receipt: NativeSlidePreservationReceipt }> {
  const { destinationBytes, sourceBytes } = input;
  const sourceSlideNumber = input.sourceSlideNumber ?? input.slideNumber;
  const destinationSlideNumber = input.destinationSlideNumber ?? input.slideNumber;
  if (!sourceSlideNumber || !destinationSlideNumber) throw new Error("Native slide preservation requires both a source and destination slide number.");
  const source = await JSZip.loadAsync(sourceBytes, { checkCRC32: false });
  const sourceSlidePart = `ppt/slides/slide${sourceSlideNumber}.xml`;
  const destinationSlidePart = `ppt/slides/slide${destinationSlideNumber}.xml`;
  const sourceSlide = source.file(sourceSlidePart);
  const sourceRelationshipsPart = relationshipPart(sourceSlidePart);
  const destinationRelationshipsPart = relationshipPart(destinationSlidePart);
  if (!sourceSlide || !source.file(sourceRelationshipsPart)) throw new Error(`The source PowerPoint is missing native slide ${sourceSlideNumber}.`);
  const sourceRelationships = await text(source, sourceRelationshipsPart);
  const sourceLayoutPart = layoutTarget(sourceSlidePart, sourceRelationships);
  const sourceLayoutBytes = await source.file(sourceLayoutPart)?.async("uint8array");
  if (!sourceLayoutBytes) throw new Error(`Source slide ${sourceSlideNumber}'s native layout part is missing.`);

  const cloned = await cloneTemplateLayoutForSlide({
    sourceBytes: destinationBytes,
    templateBytes: sourceBytes,
    command: {
      id: `preserve-native-slide-${sourceSlideNumber}-into-${destinationSlideNumber}`,
      slideNumber: destinationSlideNumber,
      templateSha256: await sha256(sourceBytes),
      templateLayoutPart: sourceLayoutPart,
      templateLayoutSha256: await sha256(sourceLayoutBytes),
      templateLayoutName: `Source-preserved slide ${sourceSlideNumber} layout`,
      rationale: `Preserve the approved ORNL template composition from source slide ${sourceSlideNumber} on output slide ${destinationSlideNumber} and retain its native dependency graph exactly.`,
      author: "ai",
    },
  });
  const destination = await JSZip.loadAsync(cloned.bytes, { checkCRC32: false });
  const preservedMasterPart = cloned.receipt.clonedMasterPart ?? await masterTarget(destination, cloned.receipt.clonedLayoutPart);
  if (!preservedMasterPart || !destination.file(preservedMasterPart)) throw new Error(`The protected slide ${sourceSlideNumber} layout is not connected to a native slide master.`);
  const destinationRelationships = await text(destination, destinationRelationshipsPart);
  destination.file(destinationSlidePart, await sourceSlide.async("uint8array"));

  const sourceContentTypes = await text(source, "[Content_Types].xml");
  let destinationContentTypes = await text(destination, "[Content_Types].xml");
  let copiedMediaCount = 0;
  let preservedNotes = false;
  const sourceNotesRelationship = relationshipOfType(sourceSlidePart, sourceRelationships, NOTES_SLIDE_TYPE);
  const destinationNotesRelationship = relationshipOfType(destinationSlidePart, destinationRelationships, NOTES_SLIDE_TYPE);
  if (sourceNotesRelationship) {
    if (!destinationNotesRelationship) throw new Error(`Protected slide ${sourceSlideNumber} has speaker notes, but the editable destination has no notes container. Export is held rather than dropping notes.`);
    const sourceNotesXml = await text(source, sourceNotesRelationship.targetPart);
    destination.file(destinationNotesRelationship.targetPart, sourceNotesXml.replace(/(<a:fld\b[^>]*\btype=(?:"slidenum"|'slidenum')[^>]*>[\s\S]*?<a:t>)[^<]*(<\/a:t>)/gi, `$1${destinationSlideNumber}$2`));
    const sourceNotesRelationshipsPart = relationshipPart(sourceNotesRelationship.targetPart);
    const destinationNotesRelationshipsPart = relationshipPart(destinationNotesRelationship.targetPart);
    const sourceNotesRelationshipsEntry = source.file(sourceNotesRelationshipsPart);
    const destinationNotesRelationshipsEntry = destination.file(destinationNotesRelationshipsPart);
    if (sourceNotesRelationshipsEntry) {
      if (!destinationNotesRelationshipsEntry) throw new Error(`Protected slide ${sourceSlideNumber}'s speaker-note relationships cannot be represented in the editable destination.`);
      const sourceNotesRelationships = await sourceNotesRelationshipsEntry.async("text");
      const destinationNotesRelationships = await destinationNotesRelationshipsEntry.async("text");
      const destinationNotesMaster = relationshipOfType(destinationNotesRelationship.targetPart, destinationNotesRelationships, NOTES_MASTER_TYPE);
      const destinationSlideBacklink = relationshipOfType(destinationNotesRelationship.targetPart, destinationNotesRelationships, SLIDE_TYPE);
      const rewrittenNotesRelationships = await Promise.all([...sourceNotesRelationships.matchAll(/<Relationship\b([^>]*?)\/>/g)].map(async (match) => {
        let attributes = match[1] ?? "";
        const type = decodeXml(attributeValue(attributes, "Type") ?? "");
        const target = attributeValue(attributes, "Target");
        const external = /^external$/i.test(attributeValue(attributes, "TargetMode") ?? "");
        if (!target || external) return `<Relationship${attributes}/>`;
        if (type === NOTES_MASTER_TYPE) {
          if (!destinationNotesMaster) throw new Error(`Protected slide ${sourceSlideNumber}'s speaker notes have no destination notes master.`);
          attributes = setAttribute(attributes, "Target", relativeTarget(destinationNotesRelationship.targetPart, destinationNotesMaster.targetPart));
          return `<Relationship${attributes}/>`;
        }
        if (type === SLIDE_TYPE) {
          if (!destinationSlideBacklink) throw new Error(`Protected slide ${sourceSlideNumber}'s speaker notes have no destination slide backlink.`);
          attributes = setAttribute(attributes, "Target", relativeTarget(destinationNotesRelationship.targetPart, destinationSlidePart));
          return `<Relationship${attributes}/>`;
        }
        const sourceTargetPart = resolveTarget(sourceNotesRelationship.targetPart, target);
        if (!/^ppt\/media\//i.test(sourceTargetPart)) throw new Error(`Protected slide ${sourceSlideNumber}'s speaker notes contain unsupported relationship ${sourceTargetPart}; export is held rather than dropping it.`);
        const media = await source.file(sourceTargetPart)?.async("uint8array");
        if (!media) throw new Error(`Protected slide ${sourceSlideNumber}'s speaker notes are missing related media ${sourceTargetPart}.`);
        const extension = basename(sourceTargetPart).match(/\.([^.]+)$/)?.[1]?.toLowerCase() ?? "bin";
        const destinationPart = `ppt/media/source-notes-slide-${sourceSlideNumber}-${(await sha256(media)).slice(0, 12)}-${copiedMediaCount + 1}.${extension}`;
        destination.file(destinationPart, media);
        destinationContentTypes = addMediaContentType(destinationContentTypes, sourceContentTypes, sourceTargetPart);
        copiedMediaCount += 1;
        attributes = setAttribute(attributes, "Target", relativeTarget(destinationNotesRelationship.targetPart, destinationPart));
        return `<Relationship${attributes}/>`;
      }));
      const notesRelationshipRoot = sourceNotesRelationships.match(/^\s*<\?xml[^>]*\?>\s*<Relationships\b([^>]*)>/i)?.[1] ?? ' xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
      destination.file(destinationNotesRelationshipsPart, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships${notesRelationshipRoot}>${rewrittenNotesRelationships.join("")}</Relationships>`);
    }
    preservedNotes = true;
  }
  const rewrittenRelationships = (await Promise.all([...sourceRelationships.matchAll(/<Relationship\b([^>]*?)\/>/g)].map(async (match) => {
    let attributes = match[1] ?? "";
    const type = decodeXml(attributeValue(attributes, "Type") ?? "");
    const target = attributeValue(attributes, "Target");
    const external = /^external$/i.test(attributeValue(attributes, "TargetMode") ?? "");
    if (!target || external) return `<Relationship${attributes}/>`;
    if (type === SLIDE_LAYOUT_TYPE) {
      attributes = setAttribute(attributes, "Target", relativeTarget(destinationSlidePart, cloned.receipt.clonedLayoutPart));
      return `<Relationship${attributes}/>`;
    }
    if (type === NOTES_SLIDE_TYPE) return undefined;
    const sourceTargetPart = resolveTarget(sourceSlidePart, target);
    if (!/^ppt\/media\//i.test(sourceTargetPart)) throw new Error(`Protected slide ${sourceSlideNumber} preservation does not yet support relationship ${sourceTargetPart}; export is held rather than flattening it.`);
    const media = await source.file(sourceTargetPart)?.async("uint8array");
    if (!media) throw new Error(`Protected slide ${sourceSlideNumber} is missing related media ${sourceTargetPart}.`);
    const extension = basename(sourceTargetPart).match(/\.([^.]+)$/)?.[1]?.toLowerCase() ?? "bin";
    const destinationPart = `ppt/media/source-slide-${sourceSlideNumber}-${(await sha256(media)).slice(0, 12)}-${copiedMediaCount + 1}.${extension}`;
    destination.file(destinationPart, media);
    destinationContentTypes = addMediaContentType(destinationContentTypes, sourceContentTypes, sourceTargetPart);
    copiedMediaCount += 1;
    attributes = setAttribute(attributes, "Target", relativeTarget(destinationSlidePart, destinationPart));
    return `<Relationship${attributes}/>`;
  }))).filter((relationship): relationship is string => Boolean(relationship));
  const usedRelationshipIds = new Set(rewrittenRelationships.map((relationship) => attributeValue(relationship, "Id") ?? "").filter(Boolean));
  for (const match of destinationRelationships.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    let attributes = match[1] ?? "";
    if (decodeXml(attributeValue(attributes, "Type") ?? "") !== NOTES_SLIDE_TYPE) continue;
    attributes = setAttribute(attributes, "Id", nextRelationshipId(usedRelationshipIds));
    rewrittenRelationships.push(`<Relationship${attributes}/>`);
  }
  const relationshipRoot = sourceRelationships.match(/^\s*<\?xml[^>]*\?>\s*<Relationships\b([^>]*)>/i)?.[1] ?? ' xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
  destination.file(destinationRelationshipsPart, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships${relationshipRoot}>${rewrittenRelationships.join("")}</Relationships>`);
  destination.file("[Content_Types].xml", destinationContentTypes);
  const bytes = await destination.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { bytes, receipt: { slideNumber: destinationSlideNumber, sourceSlideNumber, destinationSlideNumber, sourceSlideSha256: await sha256(await sourceSlide.async("uint8array")), clonedLayoutPart: cloned.receipt.clonedLayoutPart, clonedMasterPart: preservedMasterPart, copiedMediaCount, preservedNotes } };
}
