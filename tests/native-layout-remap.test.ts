import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { sha256 } from "../src/lib/hash";
import { cloneTemplateLayoutForSlide } from "../src/lib/native-layout-remap";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildCleanupProposalPptx, createNativeLayoutProposal, createNativeLayoutRecompositionProposal } from "../src/lib/cleanup";
import type { DeckJob, NativeLayoutRemapCommand } from "../src/types";

async function fixture(options: { distinctTemplateLayout?: boolean } = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-layout-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const zip = await JSZip.loadAsync(bytes);
  const layoutPart = Object.keys(zip.files).find((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name));
  assert.ok(layoutPart);
  if (options.distinctTemplateLayout) {
    const layoutXml = await zip.file(layoutPart)?.async("text");
    assert.ok(layoutXml);
    zip.file(layoutPart, layoutXml.replace(/<p:cSld\b/, "<!-- Presentation Studio approved layout variant --><p:cSld"));
  }
  const templateBytes = options.distinctTemplateLayout ? await zip.generateAsync({ type: "uint8array" }) : bytes;
  const templateZip = await JSZip.loadAsync(templateBytes);
  const layoutBytes = await templateZip.file(layoutPart)?.async("uint8array");
  assert.ok(layoutBytes);
  const command: NativeLayoutRemapCommand = {
    id: "native-layout-slide-1",
    slideNumber: 1,
    templateSha256: await sha256(templateBytes),
    templateLayoutPart: layoutPart,
    templateLayoutSha256: await sha256(layoutBytes),
    templateLayoutName: "Synthetic layout",
    rationale: "Exercise a source-bound native layout clone.",
    author: "ai",
  };
  return { bytes, templateBytes, command };
}

test("clones a complete native layout/master dependency graph and remaps only the target slide", async () => {
  const { bytes, templateBytes, command } = await fixture({ distinctTemplateLayout: true });
  const sourceAudit = await auditPptx(bytes);
  const result = await cloneTemplateLayoutForSlide({ sourceBytes: bytes, templateBytes, command });
  const outputAudit = await auditPptx(result.bytes);
  assert.equal(outputAudit.slideCount, sourceAudit.slideCount);
  assert.deepEqual(outputAudit.slides.map((slide) => slide.textHash), sourceAudit.slides.map((slide) => slide.textHash));
  assert.equal(outputAudit.masterCount, sourceAudit.masterCount + 1);
  assert.equal(result.receipt.strategy, "cloned-template-dependency-graph");
  assert.ok(result.receipt.clonedPartCount >= 3);
  assert.equal(result.receipt.clonedMasterCount, 1);
  assert.notEqual(result.receipt.priorLayoutPart, result.receipt.clonedLayoutPart);

  const output = await JSZip.loadAsync(result.bytes);
  const slideRelationships = await output.file("ppt/slides/_rels/slide1.xml.rels")?.async("text");
  assert.match(slideRelationships ?? "", new RegExp(result.receipt.clonedLayoutPart.replace("ppt/", "../").replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  const presentationRelationships = await output.file("ppt/_rels/presentation.xml.rels")?.async("text");
  assert.match(presentationRelationships ?? "", /relationships\/slideMaster/);
  assert.ok(result.receipt.clonedMasterPart);
  assert.match(presentationRelationships ?? "", new RegExp(result.receipt.clonedMasterPart.replace("ppt/", "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  const contentTypes = await output.file("[Content_Types].xml")?.async("text");
  assert.match(contentTypes ?? "", new RegExp(result.receipt.clonedLayoutPart.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  const presentation = await output.file("ppt/presentation.xml")!.async("text");
  const masterAndLayoutIds = [...presentation.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g)].map((match) => match[1]);
  for (const masterPart of Object.keys(output.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name))) {
    const masterXml = await output.file(masterPart)!.async("text");
    masterAndLayoutIds.push(...[...masterXml.matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)].map((match) => match[1]));
  }
  assert.equal(new Set(masterAndLayoutIds).size, masterAndLayoutIds.length);
});

test("isolates one selected layout instead of cloning every sibling on its master", async () => {
  const { bytes, templateBytes, command } = await fixture({ distinctTemplateLayout: true });
  const template = await JSZip.loadAsync(templateBytes);
  const masterPart = Object.keys(template.files).find((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name));
  assert.ok(masterPart);
  const masterRelationshipsPart = masterPart.replace("ppt/slideMasters/", "ppt/slideMasters/_rels/") + ".rels";
  const layoutParts = Object.keys(template.files).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name));
  const siblingNumber = Math.max(...layoutParts.map((name) => Number(name.match(/slideLayout(\d+)\.xml$/i)?.[1] ?? 0))) + 1;
  const siblingPart = `ppt/slideLayouts/slideLayout${siblingNumber}.xml`;
  const siblingRelationshipsPart = `ppt/slideLayouts/_rels/slideLayout${siblingNumber}.xml.rels`;
  template.file(siblingPart, await template.file(command.templateLayoutPart)!.async("uint8array"));
  const sourceLayoutRelationshipsPart = command.templateLayoutPart.replace("ppt/slideLayouts/", "ppt/slideLayouts/_rels/") + ".rels";
  template.file(siblingRelationshipsPart, await template.file(sourceLayoutRelationshipsPart)!.async("uint8array"));
  const siblingRelationshipId = "rId999";
  const masterRelationships = await template.file(masterRelationshipsPart)!.async("text");
  template.file(masterRelationshipsPart, masterRelationships.replace(/<\/Relationships>\s*$/i, `<Relationship Id="${siblingRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${siblingNumber}.xml"/></Relationships>`));
  const masterXml = await template.file(masterPart)!.async("text");
  template.file(masterPart, masterXml.replace(/<\/p:sldLayoutIdLst>/i, `<p:sldLayoutId id="4294960000" r:id="${siblingRelationshipId}"/></p:sldLayoutIdLst>`));
  const expandedTemplateBytes = await template.generateAsync({ type: "uint8array" });
  const result = await cloneTemplateLayoutForSlide({
    sourceBytes: bytes,
    templateBytes: expandedTemplateBytes,
    command: { ...command, templateSha256: await sha256(expandedTemplateBytes) },
  });
  assert.equal(result.receipt.clonedLayoutCount, 1);
  assert.ok(result.receipt.clonedMasterPart);
  const output = await JSZip.loadAsync(result.bytes);
  const clonedMasterRelationshipsPart = result.receipt.clonedMasterPart.replace("ppt/slideMasters/", "ppt/slideMasters/_rels/") + ".rels";
  const clonedMasterRelationships = await output.file(clonedMasterRelationshipsPart)!.async("text");
  assert.equal((clonedMasterRelationships.match(/relationships\/slideLayout/g) ?? []).length, 1);
  const clonedMasterXml = await output.file(result.receipt.clonedMasterPart)!.async("text");
  assert.equal((clonedMasterXml.match(/<p:sldLayoutId\b/g) ?? []).length, 1);
});

test("reuses an exact approved layout already in the source package without duplicating masters", async () => {
  const { bytes, templateBytes, command } = await fixture();
  const sourceAudit = await auditPptx(bytes);
  const result = await cloneTemplateLayoutForSlide({ sourceBytes: bytes, templateBytes, command });
  const outputAudit = await auditPptx(result.bytes);
  assert.equal(result.receipt.strategy, "reused-source-layout");
  assert.equal(result.receipt.clonedPartCount, 0);
  assert.equal(outputAudit.masterCount, sourceAudit.masterCount);
  assert.deepEqual(outputAudit.slides.map((slide) => slide.textHash), sourceAudit.slides.map((slide) => slide.textHash));
});

test("materializes a staged native layout proposal through the guarded cleanup pipeline", async () => {
  const { bytes, templateBytes, command } = await fixture();
  const audit = await auditPptx(bytes);
  const deck: DeckJob = { id: "deck-native-layout", name: "synthetic.pptx", sourceResourceId: "resource-native-layout", sourceSha256: await sha256(bytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const proposal = createNativeLayoutProposal(deck, "2026-08-12T20:00:00.000Z", command);
  const result = await buildCleanupProposalPptx(bytes, proposal, { templateBytes });
  assert.equal(result.layoutCount, 1);
  assert.equal(result.layoutReceipts[0]?.strategy, "reused-source-layout");
  assert.deepEqual((await auditPptx(result.bytes)).slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
});

test("combines native layout remapping and object placement in one exact-content transaction", async () => {
  const { bytes, templateBytes, command } = await fixture();
  const audit = await auditPptx(bytes);
  const object = audit.editableObjects.find((item) => item.slideNumber === 1 && item.kind === "text" && item.canMove);
  assert.ok(object);
  const deck: DeckJob = { id: "deck-native-recompose", name: "synthetic.pptx", sourceResourceId: "resource-native-recompose", sourceSha256: await sha256(bytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const target = { ...object.geometry, x: object.geometry.x + 45_720 };
  const proposal = createNativeLayoutRecompositionProposal(deck, "2026-08-12T20:00:00.000Z", command, [{ objectId: object.id, target, rationale: "Place the source-bound title in the approved native grid.", author: "ai" }]);
  const result = await buildCleanupProposalPptx(bytes, proposal, { templateBytes });
  const outputAudit = await auditPptx(result.bytes);
  assert.equal(proposal.mode, "slide-reflow");
  assert.deepEqual(proposal.changes.map((change) => change.kind), ["layout-remap", "geometry"]);
  assert.equal(result.layoutCount, 1);
  assert.equal(result.geometryCount, 1);
  assert.equal(outputAudit.editableObjects.find((item) => item.id === object.id)?.geometry.x, target.x);
  assert.deepEqual(outputAudit.slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
});

test("accumulates native layout recompositions for different slides in one pending proposal", async () => {
  const { bytes, templateBytes, command } = await fixture();
  const audit = await auditPptx(bytes);
  const firstObject = audit.editableObjects.find((item) => item.slideNumber === 1 && item.kind === "text" && item.canMove);
  const secondObject = audit.editableObjects.find((item) => item.slideNumber === 2 && item.kind === "text" && item.canMove);
  assert.ok(firstObject && secondObject);
  const timestamp = "2026-08-12T20:00:00.000Z";
  const deck: DeckJob = { id: "deck-native-multi", name: "synthetic.pptx", sourceResourceId: "resource-native-multi", sourceSha256: await sha256(bytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: timestamp, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const first = createNativeLayoutRecompositionProposal(deck, timestamp, command, [{ objectId: firstObject.id, target: { ...firstObject.geometry, x: firstObject.geometry.x + 45_720 }, rationale: "Place slide one on the approved grid.", author: "ai" }]);
  const nextTimestamp = "2026-08-12T20:00:01.000Z";
  deck.proposal = { ...first, baseUpdatedAt: nextTimestamp };
  const secondCommand = { ...command, id: "native-layout-slide-2", slideNumber: 2 };
  const combined = createNativeLayoutRecompositionProposal(deck, nextTimestamp, secondCommand, [{ objectId: secondObject.id, target: { ...secondObject.geometry, x: secondObject.geometry.x + 45_720 }, rationale: "Place slide two on the approved grid.", author: "ai" }]);
  assert.equal(combined.changes.filter((change) => change.kind === "layout-remap").length, 2);
  assert.deepEqual(combined.changes.filter((change) => change.kind === "layout-remap").map((change) => change.affectedSlideNumbers[0]), [1, 2]);
  const result = await buildCleanupProposalPptx(bytes, combined, { templateBytes });
  assert.equal(result.layoutCount, 2);
  assert.equal(result.geometryCount, 2);
  assert.deepEqual((await auditPptx(result.bytes)).slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
});

test("native layout remap rejects a stale Template Pack", async () => {
  const { bytes, templateBytes, command } = await fixture();
  await assert.rejects(() => cloneTemplateLayoutForSlide({ sourceBytes: bytes, templateBytes, command: { ...command, templateSha256: "0".repeat(64) } }), /Template Pack changed/i);
});
