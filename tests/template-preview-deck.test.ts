import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { buildTemplatePreviewDeck } from "../src/lib/template-preview-deck";

async function syntheticTemplate(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.template.main+xml" PartName="/ppt/presentation.xml"/></Types>`);
  zip.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst/><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`);
  const layout = (label: string) => `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody>${label} preview number</p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr/></p:nvSpPr><p:txBody><a:p><a:fld type="slidenum"><a:t>${label} field number</a:t></a:fld></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody>${label} body</p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`;
  zip.file("ppt/slideLayouts/slideLayout2.xml", layout("second"));
  zip.file("ppt/slideLayouts/slideLayout1.xml", layout("first"));
  return zip.generateAsync({ type: "uint8array" });
}

test("materializes one ordered native-preview slide per template layout", async () => {
  const result = await buildTemplatePreviewDeck(await syntheticTemplate());
  assert.deepEqual(result.layoutParts, ["ppt/slideLayouts/slideLayout1.xml", "ppt/slideLayouts/slideLayout2.xml"]);
  const zip = await JSZip.loadAsync(result.bytes);
  assert.ok(zip.file("ppt/slides/slide1.xml"));
  assert.ok(zip.file("ppt/slides/slide2.xml"));
  assert.match(await zip.file("ppt/slides/_rels/slide2.xml.rels")!.async("string"), /slideLayout2\.xml/);
  const firstLayout = await zip.file("ppt/slideLayouts/slideLayout1.xml")!.async("string");
  assert.doesNotMatch(firstLayout, /type="sldNum"/);
  assert.doesNotMatch(firstLayout, /first preview number/);
  assert.doesNotMatch(firstLayout, /type="slidenum"/);
  assert.doesNotMatch(firstLayout, /first field number/);
  assert.match(firstLayout, /first body/);
  const presentation = await zip.file("ppt/presentation.xml")!.async("string");
  assert.match(presentation, /<p:sldId id="256"/);
  assert.match(presentation, /<p:sldId id="257"/);
  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
  assert.match(contentTypes, /presentationml\.presentation\.main\+xml/);
  assert.doesNotMatch(contentTypes, /presentationml\.template\.main\+xml/);
});
