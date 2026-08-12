import JSZip from "jszip";

const PRESENTATION_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const PRESENTATION_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

function sortPartNames(left: string, right: string): number {
  const numberFor = (value: string) => Number(value.match(/(\d+)(?=\.xml$)/)?.[1] ?? 0);
  return numberFor(left) - numberFor(right) || left.localeCompare(right);
}

function insertBeforeClosing(xml: string, tag: string, addition: string): string {
  const closing = `</${tag}>`;
  const index = xml.lastIndexOf(closing);
  if (index < 0) throw new Error(`The template is missing ${closing}.`);
  return `${xml.slice(0, index)}${addition}${xml.slice(index)}`;
}

function replaceSlideIdList(xml: string, ids: string): string {
  if (/<p:sldIdLst\b[^>]*\/>/.test(xml)) return xml.replace(/<p:sldIdLst\b[^>]*\/>/, `<p:sldIdLst>${ids}</p:sldIdLst>`);
  if (/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/.test(xml)) return xml.replace(/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids}</p:sldIdLst>`);
  const insertionPoint = xml.search(/<p:(sldSz|notesSz|defaultTextStyle)\b/);
  if (insertionPoint < 0) return insertBeforeClosing(xml, "p:presentation", `<p:sldIdLst>${ids}</p:sldIdLst>`);
  return `${xml.slice(0, insertionPoint)}<p:sldIdLst>${ids}</p:sldIdLst>${xml.slice(insertionPoint)}`;
}

function minimalSlideXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

export async function buildTemplatePreviewDeck(templateBytes: Uint8Array): Promise<{ bytes: Uint8Array; layoutParts: string[] }> {
  if (templateBytes.byteLength < 100 || templateBytes.byteLength > 250 * 1024 * 1024) throw new Error("The installed template cannot be materialized safely.");
  const zip = await JSZip.loadAsync(templateBytes, { checkCRC32: true });
  const layoutParts = Object.keys(zip.files).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)).sort(sortPartNames);
  if (layoutParts.length === 0 || layoutParts.length > 200) throw new Error("The installed template has no supported layout catalog.");
  const presentationFile = zip.file("ppt/presentation.xml");
  const relationshipsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const contentTypesFile = zip.file("[Content_Types].xml");
  if (!presentationFile || !relationshipsFile || !contentTypesFile) throw new Error("The installed template is missing required PowerPoint package parts.");

  for (const part of Object.keys(zip.files)) {
    if (/^ppt\/slides\/slide\d+\.xml$/i.test(part) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(part)) zip.remove(part);
  }

  let relationships = await relationshipsFile.async("string");
  relationships = relationships.replace(/<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*\/>/g, "");
  const usedIds = [...relationships.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]));
  let nextRelationshipId = Math.max(0, ...usedIds) + 1;
  const slideIds: string[] = [];
  const slideRelationships: string[] = [];
  for (const [index, layoutPart] of layoutParts.entries()) {
    const slideNumber = index + 1;
    const relationshipId = `rId${nextRelationshipId++}`;
    slideIds.push(`<p:sldId id="${256 + index}" r:id="${relationshipId}"/>`);
    slideRelationships.push(`<Relationship Id="${relationshipId}" Type="${PRESENTATION_RELATIONSHIP}" Target="slides/slide${slideNumber}.xml"/>`);
    zip.file(`ppt/slides/slide${slideNumber}.xml`, minimalSlideXml());
    zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${SLIDE_LAYOUT_RELATIONSHIP}" Target="../slideLayouts/${layoutPart.split("/").pop()}"/></Relationships>`);
  }
  relationships = insertBeforeClosing(relationships, "Relationships", slideRelationships.join(""));
  zip.file("ppt/_rels/presentation.xml.rels", relationships);

  let presentation = await presentationFile.async("string");
  presentation = replaceSlideIdList(presentation, slideIds.join(""));
  zip.file("ppt/presentation.xml", presentation);

  let contentTypes = await contentTypesFile.async("string");
  contentTypes = contentTypes.replace(/<Override\b[^>]*\/>/gi, (tag) => /\bPartName="\/ppt\/slides\/slide\d+\.xml"/i.test(tag) ? "" : tag);
  contentTypes = contentTypes.replace(/<Override\b[^>]*\/>/gi, (tag) => /\bPartName="\/ppt\/presentation\.xml"/i.test(tag) ? tag.replace(/\bContentType="[^"]+"/i, `ContentType="${PRESENTATION_CONTENT_TYPE}"`) : tag);
  const slideOverrides = layoutParts.map((_part, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>`).join("");
  contentTypes = insertBeforeClosing(contentTypes, "Types", slideOverrides);
  zip.file("[Content_Types].xml", contentTypes);

  return { bytes: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }), layoutParts };
}
