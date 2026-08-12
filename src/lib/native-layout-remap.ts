import JSZip from "jszip";
import type { NativeLayoutRemapCommand } from "../types";
import { sha256 } from "./hash";

export interface NativeLayoutCloneReceipt {
  strategy: "reused-source-layout" | "cloned-template-dependency-graph";
  slideNumber: number;
  priorLayoutPart: string;
  templateLayoutPart: string;
  clonedLayoutPart: string;
  clonedMasterPart?: string;
  clonedPartCount: number;
  clonedLayoutCount: number;
  clonedMasterCount: number;
  clonedThemeCount: number;
  clonedMediaCount: number;
  remappedPlaceholderCount: number;
}

const ALLOWED_TEMPLATE_PART = /^ppt\/(?:slideLayouts|slideMasters|theme|media)\//;
const RELATIONSHIP_TYPE_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const RELATIONSHIP_TYPE_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

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

function removeAttribute(attributes: string, name: string): string {
  return attributes.replace(new RegExp(`\\s+\\b${name}=(?:"[^"]*"|'[^']*')`, "i"), "");
}

function normalizePart(value: string): string {
  const result: string[] = [];
  for (const segment of value.replaceAll("\\", "/").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      if (result.length === 0) throw new Error("A template relationship escapes the package root.");
      result.pop();
    } else result.push(segment);
  }
  return result.join("/");
}

function dirname(part: string): string {
  const index = part.lastIndexOf("/");
  return index < 0 ? "" : part.slice(0, index);
}

function basename(part: string): string {
  return part.slice(part.lastIndexOf("/") + 1);
}

function relationshipPart(part: string): string {
  return `${dirname(part)}/_rels/${basename(part)}.rels`;
}

function resolveRelationshipTarget(sourcePart: string, target: string): string {
  return normalizePart(`${dirname(sourcePart)}/${decodeXml(target)}`);
}

function relativeTarget(sourcePart: string, targetPart: string): string {
  const source = dirname(sourcePart).split("/").filter(Boolean);
  const target = targetPart.split("/").filter(Boolean);
  let shared = 0;
  while (shared < source.length && shared < target.length && source[shared] === target[shared]) shared += 1;
  return [...Array(source.length - shared).fill(".."), ...target.slice(shared)].join("/") || basename(targetPart);
}

function relationshipRecords(xml: string): Array<{ attributes: string; id: string; type: string; target: string; external: boolean }> {
  return [...xml.matchAll(/<Relationship\b([^>]*?)\/>/g)].map((match) => {
    const attributes = match[1] ?? "";
    return {
      attributes,
      id: decodeXml(attributeValue(attributes, "Id") ?? ""),
      type: decodeXml(attributeValue(attributes, "Type") ?? ""),
      target: decodeXml(attributeValue(attributes, "Target") ?? ""),
      external: /^external$/i.test(attributeValue(attributes, "TargetMode") ?? ""),
    };
  });
}

interface PlaceholderIdentity { type?: string; idx?: string; semantic: string }

function placeholderSemantic(type?: string): string {
  const normalized = type?.toLowerCase() || "obj";
  if (["title", "ctrtitle"].includes(normalized)) return "title";
  if (["obj", "body"].includes(normalized)) return "content";
  return normalized;
}

function placeholderIdentities(xml: string): PlaceholderIdentity[] {
  return [...xml.matchAll(/<p:ph\b([^>]*?)\/>/g)].map((match) => {
    const attributes = match[1] ?? "";
    const type = attributeValue(attributes, "type");
    return { type, idx: attributeValue(attributes, "idx"), semantic: placeholderSemantic(type) };
  });
}

function remapSlidePlaceholderIdentities(slideXml: string, targetLayoutXml: string): { xml: string; remappedCount: number } {
  const targets = placeholderIdentities(targetLayoutXml);
  const usedTargets = new Set<number>();
  let remappedCount = 0;
  const xml = slideXml.replace(/<p:ph\b([^>]*?)\/>/g, (_whole, rawAttributes: string) => {
    const sourceType = attributeValue(rawAttributes, "type");
    const sourceIdx = attributeValue(rawAttributes, "idx");
    const semantic = placeholderSemantic(sourceType);
    let targetIndex = targets.findIndex((target, index) => !usedTargets.has(index) && target.semantic === semantic && target.type === sourceType && target.idx === sourceIdx);
    if (targetIndex < 0) {
      const candidates = targets.map((target, index) => ({ target, index })).filter(({ target, index }) => !usedTargets.has(index) && target.semantic === semantic);
      if (candidates.length === 1) targetIndex = candidates[0].index;
    }
    if (targetIndex < 0) throw new Error(`The selected native layout cannot unambiguously map the slide's ${semantic} placeholder.`);
    usedTargets.add(targetIndex);
    const target = targets[targetIndex];
    let attributes = rawAttributes;
    attributes = target.type === undefined ? removeAttribute(attributes, "type") : setAttribute(attributes, "type", target.type);
    attributes = target.idx === undefined ? removeAttribute(attributes, "idx") : setAttribute(attributes, "idx", target.idx);
    if (attributes !== rawAttributes) remappedCount += 1;
    return `<p:ph${attributes}/>`;
  });
  return { xml, remappedCount };
}

async function matchingDestinationLayout(destination: JSZip, preferredPart: string, expectedSha256: string): Promise<string | undefined> {
  const parts = Object.keys(destination.files).filter((part) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(part));
  const ordered = parts.includes(preferredPart) ? [preferredPart, ...parts.filter((part) => part !== preferredPart).sort()] : parts.sort();
  for (const part of ordered) {
    const bytes = await destination.file(part)!.async("uint8array");
    if (await sha256(bytes) === expectedSha256) return part;
  }
  return undefined;
}

async function applyLayoutToSlide(destination: JSZip, slideNumber: number, layoutPart: string): Promise<{ priorLayoutPart: string; remappedPlaceholderCount: number }> {
  const slidePart = `ppt/slides/slide${slideNumber}.xml`;
  if (!destination.file(slidePart)) throw new Error(`Slide ${slideNumber} is not present in the source package.`);
  const layoutXml = await text(destination, layoutPart);
  const remapped = remapSlidePlaceholderIdentities(await text(destination, slidePart), layoutXml);
  destination.file(slidePart, remapped.xml);

  const slideRelationshipsPart = relationshipPart(slidePart);
  let slideRelationships = await text(destination, slideRelationshipsPart);
  const layoutRelationships = relationshipRecords(slideRelationships).filter((relationship) => relationship.type === RELATIONSHIP_TYPE_SLIDE_LAYOUT);
  if (layoutRelationships.length !== 1) throw new Error(`Slide ${slideNumber} must have exactly one native layout relationship.`);
  const priorLayoutPart = resolveRelationshipTarget(slidePart, layoutRelationships[0].target);
  slideRelationships = slideRelationships.replace(/<Relationship\b([^>]*?)\/>/g, (whole, attributes: string) => {
    if (decodeXml(attributeValue(attributes, "Type") ?? "") !== RELATIONSHIP_TYPE_SLIDE_LAYOUT) return whole;
    return `<Relationship${setAttribute(attributes, "Target", relativeTarget(slidePart, layoutPart))}/>`;
  });
  destination.file(slideRelationshipsPart, slideRelationships);
  return { priorLayoutPart, remappedPlaceholderCount: remapped.remappedCount };
}

async function text(zip: JSZip, part: string): Promise<string> {
  const entry = zip.file(part);
  if (!entry) throw new Error(`Required PowerPoint package part ${part} is missing.`);
  return entry.async("text");
}

async function templateClosure(template: JSZip, initialLayoutPart: string): Promise<Set<string>> {
  const initial = normalizePart(initialLayoutPart);
  if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(initial)) throw new Error("Choose a concrete slideLayout XML part from the active Template Pack.");
  const pending = [initial];
  const visited = new Set<string>();
  while (pending.length > 0) {
    const part = pending.shift()!;
    if (visited.has(part)) continue;
    if (!ALLOWED_TEMPLATE_PART.test(part)) throw new Error(`Template relationship cloning does not allow ${part}.`);
    if (!template.file(part)) throw new Error(`The Template Pack is missing ${part}.`);
    visited.add(part);
    const relsEntry = template.file(relationshipPart(part));
    if (!relsEntry) continue;
    const rels = await relsEntry.async("text");
    for (const relationship of relationshipRecords(rels)) {
      if (relationship.external) throw new Error(`The approved layout dependency ${part} contains an external relationship.`);
      const target = resolveRelationshipTarget(part, relationship.target);
      if (!ALLOWED_TEMPLATE_PART.test(target)) throw new Error(`The approved layout dependency graph includes unsupported part ${target}.`);
      if (!visited.has(target)) pending.push(target);
    }
  }
  const masters = [...visited].filter((part) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(part));
  if (masters.length !== 1) throw new Error(`A native layout clone requires exactly one reachable slide master; found ${masters.length}.`);
  return visited;
}

function nextNumber(parts: string[], pattern: RegExp): number {
  return Math.max(0, ...parts.map((part) => Number(part.match(pattern)?.[1] ?? 0)).filter(Number.isFinite)) + 1;
}

async function destinationMap(destination: JSZip, template: JSZip, parts: Set<string>): Promise<Map<string, string>> {
  const destinationParts = Object.keys(destination.files);
  let layoutNumber = nextNumber(destinationParts, /^ppt\/slideLayouts\/slideLayout(\d+)\.xml$/i);
  let masterNumber = nextNumber(destinationParts, /^ppt\/slideMasters\/slideMaster(\d+)\.xml$/i);
  let themeNumber = nextNumber(destinationParts, /^ppt\/theme\/theme(\d+)\.xml$/i);
  let mediaNumber = 1;
  const used = new Set(destinationParts);
  const result = new Map<string, string>();
  for (const part of [...parts].sort()) {
    let candidate: string;
    if (/^ppt\/slideLayouts\//i.test(part)) candidate = `ppt/slideLayouts/slideLayout${layoutNumber++}.xml`;
    else if (/^ppt\/slideMasters\//i.test(part)) candidate = `ppt/slideMasters/slideMaster${masterNumber++}.xml`;
    else if (/^ppt\/theme\//i.test(part)) candidate = `ppt/theme/theme${themeNumber++}.xml`;
    else {
      const extension = basename(part).match(/\.([^.]+)$/)?.[1]?.toLowerCase() ?? "bin";
      const digest = (await sha256(await template.file(part)!.async("uint8array"))).slice(0, 10);
      candidate = `ppt/media/pstudio-${digest}-${mediaNumber++}.${extension}`;
      while (used.has(candidate)) candidate = `ppt/media/pstudio-${digest}-${mediaNumber++}.${extension}`;
    }
    used.add(candidate);
    result.set(part, candidate);
  }
  return result;
}

function rewriteRelationships(xml: string, sourcePart: string, destinationPart: string, mapping: Map<string, string>): string {
  return xml.replace(/<Relationship\b([^>]*?)\/>/g, (whole, rawAttributes: string) => {
    const external = /^external$/i.test(attributeValue(rawAttributes, "TargetMode") ?? "");
    if (external) return whole;
    const target = attributeValue(rawAttributes, "Target");
    if (!target) throw new Error(`A relationship in ${sourcePart} has no target.`);
    const resolved = resolveRelationshipTarget(sourcePart, target);
    const mapped = mapping.get(resolved);
    if (!mapped) throw new Error(`The cloned relationship graph omitted ${resolved}.`);
    return `<Relationship${setAttribute(rawAttributes, "Target", relativeTarget(destinationPart, mapped))}/>`;
  });
}

function nextRelationshipId(xml: string): string {
  const used = new Set([...xml.matchAll(/\bId=(?:"([^"]+)"|'([^']+)')/g)].map((match) => match[1] ?? match[2] ?? ""));
  let number = 1;
  while (used.has(`rId${number}`)) number += 1;
  return `rId${number}`;
}

function addContentTypes(destinationXml: string, templateXml: string, mapping: Map<string, string>): string {
  let result = destinationXml;
  const additions: string[] = [];
  for (const [sourcePart, destinationPart] of mapping) {
    const escapedSource = sourcePart.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const override = templateXml.match(new RegExp(`<Override\\b([^>]*\\bPartName=(?:"/${escapedSource}"|'/${escapedSource}')[^>]*)/>`, "i"));
    if (override && !new RegExp(`\\bPartName=(?:"/${destinationPart}"|'/${destinationPart}')`, "i").test(result)) {
      additions.push(`<Override${setAttribute(override[1] ?? "", "PartName", `/${destinationPart}`)}/>`);
      continue;
    }
    const extension = basename(sourcePart).match(/\.([^.]+)$/)?.[1];
    if (!extension || new RegExp(`<Default\\b[^>]*\\bExtension=(?:"${extension}"|'${extension}')`, "i").test(result)) continue;
    const templateDefault = templateXml.match(new RegExp(`<Default\\b([^>]*\\bExtension=(?:"${extension}"|'${extension}')[^>]*)/>`, "i"));
    if (templateDefault) additions.push(`<Default${templateDefault[1]}/>`);
  }
  if (additions.length > 0) result = result.replace(/<\/Types>\s*$/i, `${additions.join("")}</Types>`);
  return result;
}

async function maximumLayoutId(destination: JSZip): Promise<number> {
  let maximum = 0;
  for (const part of Object.keys(destination.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name))) {
    const xml = await text(destination, part);
    for (const match of xml.matchAll(/<p:sldLayoutId\b[^>]*\bid=(?:"(\d+)"|'(\d+)')/g)) maximum = Math.max(maximum, Number(match[1] ?? match[2] ?? 0));
  }
  return maximum;
}

function replaceLayoutIds(masterXml: string, firstId: number): { xml: string; nextId: number } {
  let nextId = firstId;
  return {
    xml: masterXml.replace(/<p:sldLayoutId\b([^>]*?)\/>/g, (_whole, attributes: string) => `<p:sldLayoutId${setAttribute(attributes, "id", String(nextId++))}/>`),
    nextId,
  };
}

export async function templateLayoutPartSha256(templateBytes: Uint8Array, layoutPart: string): Promise<string> {
  const template = await JSZip.loadAsync(templateBytes, { checkCRC32: false });
  const normalized = normalizePart(layoutPart);
  if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(normalized)) throw new Error("Choose a concrete slideLayout XML part from the active Template Pack.");
  const bytes = await template.file(normalized)?.async("uint8array");
  if (!bytes) throw new Error(`The active Template Pack does not contain ${normalized}.`);
  return sha256(bytes);
}

export async function cloneTemplateLayoutForSlide(input: {
  sourceBytes: Uint8Array;
  templateBytes: Uint8Array;
  command: NativeLayoutRemapCommand;
}): Promise<{ bytes: Uint8Array; receipt: NativeLayoutCloneReceipt }> {
  const { sourceBytes, templateBytes, command } = input;
  const [sourceDigest, templateDigest] = await Promise.all([sha256(sourceBytes), sha256(templateBytes)]);
  if (templateDigest !== command.templateSha256) throw new Error("The active Template Pack changed after the native layout proposal was staged.");
  if (sourceDigest.length !== 64) throw new Error("The source package could not be revision-bound.");
  const [destination, template] = await Promise.all([JSZip.loadAsync(sourceBytes, { checkCRC32: false }), JSZip.loadAsync(templateBytes, { checkCRC32: false })]);
  const targetLayoutPart = normalizePart(command.templateLayoutPart);
  const targetLayoutBytes = await template.file(targetLayoutPart)?.async("uint8array");
  if (!targetLayoutBytes) throw new Error(`The active Template Pack no longer contains ${targetLayoutPart}.`);
  if (await sha256(targetLayoutBytes) !== command.templateLayoutSha256) throw new Error("The selected Template Pack layout changed after staging.");
  const reusableLayout = await matchingDestinationLayout(destination, targetLayoutPart, command.templateLayoutSha256);
  if (reusableLayout) {
    const applied = await applyLayoutToSlide(destination, command.slideNumber, reusableLayout);
    const bytes = await destination.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    return {
      bytes,
      receipt: {
        strategy: "reused-source-layout",
        slideNumber: command.slideNumber,
        priorLayoutPart: applied.priorLayoutPart,
        templateLayoutPart: targetLayoutPart,
        clonedLayoutPart: reusableLayout,
        clonedPartCount: 0,
        clonedLayoutCount: 0,
        clonedMasterCount: 0,
        clonedThemeCount: 0,
        clonedMediaCount: 0,
        remappedPlaceholderCount: applied.remappedPlaceholderCount,
      },
    };
  }
  const closure = await templateClosure(template, targetLayoutPart);
  if (closure.size > 160) throw new Error(`The approved layout dependency graph is unexpectedly large (${closure.size} parts).`);
  const mapping = await destinationMap(destination, template, closure);
  let nextLayoutId = (await maximumLayoutId(destination)) + 1;
  for (const sourcePart of [...closure].sort()) {
    const destinationPart = mapping.get(sourcePart)!;
    const bytes = await template.file(sourcePart)!.async("uint8array");
    let output: Uint8Array | string = bytes;
    if (/\.xml$/i.test(sourcePart)) {
      let xml = new TextDecoder().decode(bytes);
      if (/^ppt\/slideMasters\//i.test(sourcePart)) {
        const replaced = replaceLayoutIds(xml, nextLayoutId);
        xml = replaced.xml;
        nextLayoutId = replaced.nextId;
      }
      output = xml;
    }
    destination.file(destinationPart, output);
    const sourceRelationships = template.file(relationshipPart(sourcePart));
    if (sourceRelationships) {
      const rewritten = rewriteRelationships(await sourceRelationships.async("text"), sourcePart, destinationPart, mapping);
      destination.file(relationshipPart(destinationPart), rewritten);
    }
  }

  const contentTypes = await text(destination, "[Content_Types].xml");
  const templateContentTypes = await text(template, "[Content_Types].xml");
  destination.file("[Content_Types].xml", addContentTypes(contentTypes, templateContentTypes, mapping));

  const clonedMaster = [...mapping.entries()].find(([part]) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(part))?.[1];
  const clonedLayout = mapping.get(targetLayoutPart);
  if (!clonedMaster || !clonedLayout) throw new Error("The native layout clone did not produce a master and target layout.");

  let presentationRelationships = await text(destination, "ppt/_rels/presentation.xml.rels");
  const presentationRelationshipId = nextRelationshipId(presentationRelationships);
  presentationRelationships = presentationRelationships.replace(/<\/Relationships>\s*$/i, `<Relationship Id="${presentationRelationshipId}" Type="${RELATIONSHIP_TYPE_SLIDE_MASTER}" Target="${escapeXml(relativeTarget("ppt/presentation.xml", clonedMaster))}"/></Relationships>`);
  destination.file("ppt/_rels/presentation.xml.rels", presentationRelationships);

  let presentation = await text(destination, "ppt/presentation.xml");
  const existingMasterIds = [...presentation.matchAll(/<p:sldMasterId\b[^>]*\bid=(?:"(\d+)"|'(\d+)')/g)].map((match) => Number(match[1] ?? match[2] ?? 0));
  const masterId = Math.max(2_147_483_647, ...existingMasterIds) + 1;
  if (masterId > 4_294_967_295) throw new Error("The PowerPoint master ID range is exhausted.");
  if (/<p:sldMasterIdLst\b[^>]*>/i.test(presentation)) presentation = presentation.replace(/<\/p:sldMasterIdLst>/i, `<p:sldMasterId id="${masterId}" r:id="${presentationRelationshipId}"/></p:sldMasterIdLst>`);
  else presentation = presentation.replace(/<p:notesMasterIdLst\b/i, `<p:sldMasterIdLst><p:sldMasterId id="${masterId}" r:id="${presentationRelationshipId}"/></p:sldMasterIdLst><p:notesMasterIdLst`);
  destination.file("ppt/presentation.xml", presentation);

  const applied = await applyLayoutToSlide(destination, command.slideNumber, clonedLayout);

  const bytes = await destination.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return {
    bytes,
    receipt: {
      strategy: "cloned-template-dependency-graph",
      slideNumber: command.slideNumber,
      priorLayoutPart: applied.priorLayoutPart,
      templateLayoutPart: targetLayoutPart,
      clonedLayoutPart: clonedLayout,
      clonedMasterPart: clonedMaster,
      clonedPartCount: closure.size,
      clonedLayoutCount: [...closure].filter((part) => /^ppt\/slideLayouts\//i.test(part)).length,
      clonedMasterCount: [...closure].filter((part) => /^ppt\/slideMasters\//i.test(part)).length,
      clonedThemeCount: [...closure].filter((part) => /^ppt\/theme\//i.test(part)).length,
      clonedMediaCount: [...closure].filter((part) => /^ppt\/media\//i.test(part)).length,
      remappedPlaceholderCount: applied.remappedPlaceholderCount,
    },
  };
}
