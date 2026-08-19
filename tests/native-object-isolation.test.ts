import assert from "node:assert/strict";
import fs from "node:fs/promises";
import JSZip from "jszip";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { isolateNativePowerPointObjects, isolateNativePowerPointSlide } from "../src/lib/native-object-isolation";
import { auditPptx } from "../src/lib/pptx-audit";

test("native object isolation keeps one requested top-level shape without flattening its package", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-isolation-"));
  const sourcePath = path.join(directory, "source.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const audit = await auditPptx(sourceBytes);
  const sourceObject = audit.editableObjects.find((object) => object.slideNumber === 2);
  assert.ok(sourceObject);
  const isolated = await isolateNativePowerPointObjects({ sourceBytes, slideNumber: 2, shapeIds: [sourceObject.shapeId] });
  assert.deepEqual(isolated.receipt.preservedShapeIds, [sourceObject.shapeId]);
  assert.ok(isolated.receipt.hiddenShapeIds.length > 0);
  const zip = await JSZip.loadAsync(isolated.bytes);
  const presentation = await zip.file("ppt/presentation.xml")!.async("text");
  const slide = await zip.file("ppt/slides/slide2.xml")!.async("text");
  assert.equal([...presentation.matchAll(/<p:sldId\b/g)].length, 1);
  assert.match(slide, new RegExp(`<p:cNvPr\\b(?=[^>]*\\bid=(?:"${sourceObject.shapeId}"|'${sourceObject.shapeId}'))(?![^>]*\\bhidden=)[^>]*>`, "i"));
  for (const hiddenId of isolated.receipt.hiddenShapeIds) assert.match(slide, new RegExp(`<p:cNvPr\\b(?=[^>]*\\bid=(?:"${hiddenId}"|'${hiddenId}'))(?=[^>]*\\bhidden=(?:"1"|'1'))[^>]*>`, "i"));
  assert.ok(zip.file("ppt/slides/slide1.xml"));
  assert.ok(zip.file("ppt/slides/slide2.xml"));
});

test("native slide isolation keeps the selected slide relationship graph byte-identical", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-slide-isolation-"));
  const sourcePath = path.join(directory, "source.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const sourceSlideRelationships = await sourceZip.file("ppt/slides/_rels/slide1.xml.rels")!.async("text");
  const sourceMaster = await sourceZip.file("ppt/slideMasters/slideMaster1.xml")!.async("text");
  const sourceLayout = await sourceZip.file("ppt/slideLayouts/slideLayout1.xml")!.async("text");

  const isolated = await isolateNativePowerPointSlide({ sourceBytes, slideNumber: 1 });
  assert.deepEqual(isolated.receipt, { slideNumber: 1 });
  const zip = await JSZip.loadAsync(isolated.bytes);
  const presentation = await zip.file("ppt/presentation.xml")!.async("text");
  assert.equal([...presentation.matchAll(/<p:sldId\b/g)].length, 1);
  assert.equal(await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("text"), sourceSlideRelationships);
  assert.equal(await zip.file("ppt/slideMasters/slideMaster1.xml")!.async("text"), sourceMaster);
  assert.equal(await zip.file("ppt/slideLayouts/slideLayout1.xml")!.async("text"), sourceLayout);
  assert.match(sourceSlideRelationships, /relationships\/hyperlink/i);
  assert.match(sourceSlideRelationships, /TargetMode=(?:"External"|'External')/i);
  assert.ok(zip.file("ppt/slides/slide2.xml"), "Unreferenced source parts remain available to preserve the package graph without reconstruction.");
});
