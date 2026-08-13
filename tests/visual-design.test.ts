import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildCleanupProposalPptx, createVisualDesignProposal } from "../src/lib/cleanup";
import { sha256 } from "../src/lib/hash";
import type { DeckJob } from "../src/types";

test("native visual design adds editable brand geometry and styles text without changing content", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-visual-design-"));
  const sourcePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceZip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  const sourceSlideXml = await sourceZip.file("ppt/slides/slide1.xml")?.async("text");
  assert.ok(sourceSlideXml);
  sourceZip.file("ppt/slides/slide1.xml", sourceSlideXml.replace(/<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/, '<a:pPr lvl="0"/>'));
  const sourceBytes = await sourceZip.generateAsync({ type: "uint8array" });
  const audit = await auditPptx(sourceBytes);
  const title = audit.editableObjects.find((object) => object.slideNumber === 1 && object.kind === "text" && object.name.includes("Text"));
  assert.ok(title);
  const deck: DeckJob = { id: "deck-visual", name: "synthetic.pptx", sourceResourceId: "resource-visual", sourceSha256: await sha256(sourceBytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const proposal = createVisualDesignProposal(deck, "2026-08-12T20:00:00.000Z", {
    slideNumber: 1,
    textStyles: [{ objectId: title.id, fontSizePt: 32, bold: true, color: "#373A36", alignment: "left", verticalAlignment: "middle", insetsInches: { top: .04, right: .06, bottom: .04, left: .06 }, paragraphStyle: { lineSpacingMultiple: 1.08, spaceAfterPt: 7, bulletLeftMarginInches: .36, bulletHangingInches: .22 }, rationale: "Establish a clear editable title hierarchy.", author: "ai" }],
    decorations: [{ id: "slide-1-title-rule", name: "ORNL title rule", geometry: { x: 594_360, y: 1_097_280, width: 822_960, height: 54_864 }, fillColor: "#00662C", lineWidthPt: 0, behindContent: true, rationale: "Add a restrained ORNL title accent.", author: "ai" }],
  });
  const result = await buildCleanupProposalPptx(sourceBytes, proposal);
  assert.equal(result.textStyleCount, 1);
  assert.equal(result.decorationCount, 1);
  const outputAudit = await auditPptx(result.bytes);
  assert.deepEqual(outputAudit.slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
  const zip = await JSZip.loadAsync(result.bytes);
  const slideXml = await zip.file("ppt/slides/slide1.xml")?.async("text");
  assert.match(slideXml ?? "", /name="ORNL title rule"/);
  assert.match(slideXml ?? "", /typeface="Aptos"/);
  assert.match(slideXml ?? "", /sz="3200"/);
  assert.doesNotMatch(slideXml ?? "", /\/\s+(?:sz|b|i|algn)=/);
  assert.match(slideXml ?? "", /<a:pPr lvl="0" algn="l"><a:lnSpc><a:spcPct val="108000"\/><\/a:lnSpc><a:spcAft><a:spcPts val="700"\/><\/a:spcAft><\/a:pPr>/);
});

test("an accepted design remains the editable baseline and reopens only the revised slide", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-iterative-review-"));
  const sourcePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const audit = await auditPptx(sourceBytes);
  const slideOneTitle = audit.editableObjects.find((object) => object.slideNumber === 1 && object.kind === "text");
  const slideTwoTitle = audit.editableObjects.find((object) => object.slideNumber === 2 && object.kind === "text");
  assert.ok(slideOneTitle && slideTwoTitle);
  const revision = "2026-08-12T20:00:00.000Z";
  const deck: DeckJob = { id: "deck-iterative", name: "synthetic.pptx", sourceResourceId: "resource-iterative", sourceSha256: await sha256(sourceBytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const accepted = createVisualDesignProposal(deck, revision, { slideNumber: 1, textStyles: [{ objectId: slideOneTitle.id, fontSizePt: 28, bold: true, rationale: "Establish the first approved title hierarchy.", author: "ai" }], decorations: [] });
  accepted.status = "applied";
  accepted.slideReviews = [{ slideNumber: 1, decision: "approved", reviewedAt: revision }, { slideNumber: 2, decision: "approved", reviewedAt: revision }];
  const revised = createVisualDesignProposal({ ...deck, proposal: accepted, status: "approved" }, revision, { slideNumber: 2, textStyles: [{ objectId: slideTwoTitle.id, fontSizePt: 26, bold: true, rationale: "Revise only the second slide title.", author: "ai" }], decorations: [] });
  assert.equal(revised.status, "pending");
  assert.equal(revised.changes.filter((change) => change.kind === "text-style").length, 2);
  assert.deepEqual(revised.slideReviews, [{ slideNumber: 1, decision: "approved", reviewedAt: revision }]);
});

test("a visual revision can remove obsolete one-off decorations while retaining shared components", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-decoration-removal-"));
  const sourcePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const audit = await auditPptx(sourceBytes);
  const revision = "2026-08-12T20:00:00.000Z";
  const deck: DeckJob = { id: "deck-decoration-removal", name: "synthetic.pptx", sourceResourceId: "resource-decoration-removal", sourceSha256: await sha256(sourceBytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const accepted = createVisualDesignProposal(deck, revision, { slideNumber: 1, textStyles: [], decorations: [
    { id: "shared-title-rule", name: "Shared title rule", geometry: { x: 400_000, y: 800_000, width: 700_000, height: 50_000 }, fillColor: "#00662C", lineWidthPt: 0, behindContent: true, rationale: "Keep the shared title component.", author: "ai" },
    { id: "one-off-box", name: "One-off box", geometry: { x: 400_000, y: 1_000_000, width: 100_000, height: 500_000 }, fillColor: "#00454D", lineWidthPt: 0, behindContent: true, rationale: "Temporary grouping cue.", author: "ai" },
  ] });
  accepted.status = "applied";
  const revised = createVisualDesignProposal({ ...deck, proposal: accepted, status: "approved" }, revision, { slideNumber: 1, removeDecorationIds: ["one-off-box"], textStyles: [], decorations: [] });
  const decorations = revised.changes.flatMap((change) => change.decorationCommands ?? []);
  assert.deepEqual(decorations.map((decoration) => decoration.id), ["shared-title-rule"]);
  const result = await buildCleanupProposalPptx(sourceBytes, revised);
  const zip = await JSZip.loadAsync(result.bytes);
  const slideXml = await zip.file("ppt/slides/slide1.xml")?.async("text");
  assert.match(slideXml ?? "", /name="Shared title rule"/);
  assert.doesNotMatch(slideXml ?? "", /name="One-off box"/);
});

test("repeated visual revisions consolidate shared slide components instead of redrawing the slide", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-component-consolidation-"));
  const sourcePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const audit = await auditPptx(sourceBytes);
  const textObjects = audit.editableObjects.filter((object) => object.slideNumber === 1 && object.kind === "text");
  assert.ok(textObjects.length >= 2);
  const revision = "2026-08-12T20:00:00.000Z";
  const deck: DeckJob = { id: "deck-component-consolidation", name: "synthetic.pptx", sourceResourceId: "resource-component-consolidation", sourceSha256: await sha256(sourceBytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const first = createVisualDesignProposal(deck, revision, {
    slideNumber: 1,
    textStyles: [{ objectId: textObjects[0].id, fontSizePt: 30, bold: true, rationale: "Use the shared title component.", author: "ai" }],
    decorations: [{ id: "shared-title-rule", name: "Shared title rule", geometry: { x: 400_000, y: 800_000, width: 700_000, height: 50_000 }, fillColor: "#00662C", lineWidthPt: 0, behindContent: true, rationale: "Use the shared title rule component.", author: "ai" }],
  });
  const revised = createVisualDesignProposal({ ...deck, proposal: first }, revision, {
    slideNumber: 1,
    textStyles: [
      { objectId: textObjects[0].id, fontSizePt: 32, bold: true, rationale: "Update the shared title component.", author: "ai" },
      { objectId: textObjects[1].id, fontSizePt: 18, rationale: "Use the shared body component.", author: "ai" },
    ],
    decorations: [{ id: "shared-title-rule", name: "Shared title rule", geometry: { x: 400_000, y: 800_000, width: 800_000, height: 50_000 }, fillColor: "#00662C", lineWidthPt: 0, behindContent: true, rationale: "Update the shared title rule component.", author: "ai" }],
  });
  assert.equal(new Set(revised.changes.map((change) => change.id)).size, revised.changes.length);
  const textChange = revised.changes.find((change) => change.id === "text-style-slide-1");
  const decorationChange = revised.changes.find((change) => change.id === "decoration-slide-1");
  assert.equal(textChange?.textStyleCommands?.length, 2);
  assert.equal(textChange?.textStyleCommands?.find((command) => command.objectId === textObjects[0].id)?.fontSizePt, 32);
  assert.equal(decorationChange?.decorationCommands?.length, 1);
  assert.equal(decorationChange?.decorationCommands?.[0].geometry.width, 800_000);
});
