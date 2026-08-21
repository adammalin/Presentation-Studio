import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { buildSlideRenderCatalog, buildTemplateCatalog } from "../src/lib/template-catalog";
import { previewElementFor } from "../src/lib/studio-web-scene";

async function syntheticTemplate(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("ppt/presentation.xml", `<?xml version="1.0"?><p:presentation xmlns:p="p"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  zip.file("ppt/theme/theme1.xml", `<?xml version="1.0"?><a:theme xmlns:a="a"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="373A36"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="00662C"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  zip.file("ppt/slideMasters/slideMaster1.xml", `<?xml version="1.0"?><p:sldMaster xmlns:p="p" xmlns:a="a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Master bar"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="350000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp></p:spTree></p:cSld></p:sldMaster>`);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`);
  zip.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld name="Title | Synthetic"><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="6" name="Image frame"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill><p:spPr><a:xfrm><a:off x="7000000" y="0"/><a:ext cx="5192000" cy="6858000"/></a:xfrm></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="3" name="Title 1"/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="900000" y="1900000"/><a:ext cx="5200000" cy="1200000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Presentation title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`);
  zip.file("ppt/media/image1.png", Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]));
  zip.file("ppt/slides/slide1.xml", `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Current slide title</a:t></a:r></a:p></p:txBody></p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="800000" y="3500000"/><a:ext cx="5200000" cy="1500000"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="2600000"/><a:gridCol w="2600000"/></a:tblGrid><a:tr h="750000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Header A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Header B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="750000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame><p:pic><p:nvPicPr><p:cNvPr id="6" name="Partner SVG"/></p:nvPicPr><p:blipFill><a:blip><a:extLst><a:ext><asvg:svgBlip r:embed="rId2"/></a:ext></a:extLst></a:blip><a:srcRect l="5000"/></p:blipFill><p:spPr><a:xfrm><a:off x="7000000" y="3500000"/><a:ext cx="1800000" cy="700000"/></a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>`);
  zip.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/></Relationships>`);
  zip.file("ppt/media/image2.svg", `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="40"><rect width="100" height="40" fill="#007833"/></svg>`);
  return zip.generateAsync({ type: "uint8array" });
}

test("builds a visual catalog from native PowerPoint master and layout parts", async () => {
  const catalog = await buildTemplateCatalog(await syntheticTemplate(), "synthetic-template.potx");
  assert.equal(catalog.masterCount, 1);
  assert.equal(catalog.layouts.length, 1);
  assert.equal(catalog.layouts[0].name, "Title | Synthetic");
  assert.equal(catalog.layouts[0].category, "title");
  assert.deepEqual(catalog.layouts[0].placeholderTypes, ["ctrTitle"]);
  assert.ok(catalog.layouts[0].elements.some((element) => element.kind === "shape" && element.fill === "#00662C"));
  assert.ok(catalog.layouts[0].elements.some((element) => element.kind === "text" && element.text === "Presentation title"));
  assert.ok(catalog.layouts[0].elements.some((element) => element.kind === "image"));
  assert.match(catalog.media["ppt/media/image1.png"], /^data:image\/png;base64,/);
  assert.match(catalog.sha256, /^[0-9a-f]{64}$/);
});

test("builds current slide previews with inherited title geometry and native table cells", async () => {
  const catalog = await buildSlideRenderCatalog(await syntheticTemplate(), "synthetic-deck.pptx");
  assert.equal(catalog.renderer, "local-ooxml-preview");
  assert.equal(catalog.slides.length, 1);
  assert.equal(catalog.slides[0].title, "Current slide title");
  const title = catalog.slides[0].elements.find((element) => element.kind === "text" && element.text === "Current slide title");
  assert.ok(title);
  assert.equal(title.x, 900000);
  assert.equal(title.fontFamily, "Aptos Display");
  assert.equal(title.origin, "slide");
  assert.equal(title.sourcePart, "ppt/slides/slide1.xml");
  assert.equal(title.sourceShapeId, "2");
  assert.match(title.textHash ?? "", /^[0-9a-f]{64}$/);
  assert.deepEqual(title.sourceParagraphs?.map((paragraph) => paragraph.text), ["Current slide title"]);
  assert.ok(catalog.slides[0].elements.some((element) => element.kind === "text" && element.text === "Header A"));
  assert.ok(catalog.slides[0].elements.some((element) => element.kind === "shape" && element.name.includes("cell")));
  const svg = catalog.slides[0].elements.find((element) => element.sourceShapeId === "6");
  assert.equal(svg?.mediaId, "ppt/media/image1.png");
  const exactSlideSvg = previewElementFor(catalog, 1, "6", "image");
  assert.equal(exactSlideSvg?.origin, "slide");
  assert.equal(exactSlideSvg?.mediaId, "ppt/media/image2.svg");
  assert.equal(exactSlideSvg?.sourceCropped, true);
  assert.deepEqual(exactSlideSvg?.sourceCrop, { left: .05, top: 0, right: 0, bottom: 0 });
  assert.match(catalog.media["ppt/media/image2.svg"], /^data:image\/svg\+xml;base64,/);
});

test("slide catalog preserves ordered inline PowerPoint breaks as semantic text boundaries", async () => {
  const zip = await JSZip.loadAsync(await syntheticTemplate());
  const slide = await zip.file("ppt/slides/slide1.xml")!.async("text");
  zip.file("ppt/slides/slide1.xml", slide.replace(
    "<a:t>Current slide title</a:t></a:r>",
    "<a:t>Current slide title</a:t></a:r><a:br/><a:r><a:rPr b=\"0\"/><a:t>Supporting attribution</a:t></a:r>",
  ));
  const catalog = await buildSlideRenderCatalog(await zip.generateAsync({ type: "uint8array" }), "inline-break.pptx");
  const title = catalog.slides[0].elements.find((element) => element.origin === "slide" && element.kind === "text" && element.sourceShapeId === "2");
  assert.equal(title?.text, "Current slide title\nSupporting attribution");
  assert.deepEqual(title?.sourceParagraphs?.map((paragraph) => paragraph.text), ["Current slide title", "Supporting attribution"]);
});

test("rejects files without a PowerPoint layout catalog", async () => {
  const zip = new JSZip();
  zip.file("readme.txt", "not a template");
  const bytes = await zip.generateAsync({ type: "uint8array" });
  await assert.rejects(() => buildTemplateCatalog(bytes, "bad.potx"), /master and layout catalog/);
});
