import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { applyCleanupToPptx, buildCleanupProposalPptx, createDesignerCleanupProposal, createFontCleanupProposal, createGeometryBatchProposal, createGeometryEditProposal, createTableStyleProposal } from "../src/lib/cleanup";
import { PRESENTATION_DESIGN_STANDARD } from "../src/lib/design-standard";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildAuditReport } from "../src/lib/report";
import { createProject, projectSchema } from "../src/lib/project";
import type { DeckJob } from "../src/types";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";

async function fixtureBytes(): Promise<Uint8Array> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-pptx-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  return new Uint8Array(await fs.readFile(filePath));
}

test("OOXML audit inventories synthetic slides, fonts, and tables", async () => {
  const audit = await auditPptx(await fixtureBytes());
  assert.equal(audit.slideCount, 2);
  assert.equal(audit.tableCount, 1);
  assert.equal(audit.tables.length, 1);
  assert.equal(audit.tables[0].rowCount, 3);
  assert.equal(audit.tables[0].columnCount, 3);
  assert.match(audit.tables[0].styleFingerprint, /^[0-9a-f]{64}$/);
  assert.match(audit.tables[0].contentHash, /^[0-9a-f]{64}$/);
  assert.match(audit.tables[0].structureHash, /^[0-9a-f]{64}$/);
  assert.equal(audit.alignmentRepairs.length, 2);
  assert.ok(audit.alignmentRepairs.some((repair) => repair.ruleId === "cover.dominant-left-edge"));
  assert.ok(audit.alignmentRepairs.some((repair) => repair.ruleId === "peer.dominant-left-edge"));
  assert.ok(audit.textBoxes.length >= 10);
  const bulletTextBox = audit.textBoxes.find((textBox) => textBox.text.startsWith("This deliberately dense"));
  assert.ok(bulletTextBox);
  assert.equal(bulletTextBox.bulletParagraphCount, 1);
  assert.equal(bulletTextBox.opticalAlignmentConfidence, "direct");
  assert.ok(bulletTextBox.paragraphLeftMarginsEmu.some((margin) => margin > 0));
  assert.equal(bulletTextBox.estimatedOpticalLeftEmu, bulletTextBox.geometry.x + bulletTextBox.textInsets.left + Math.min(...bulletTextBox.paragraphLeftMarginsEmu));
  assert.ok(audit.editableObjects.length >= audit.textBoxes.length);
  assert.ok(audit.editableObjects.some((object) => object.kind === "table" && object.tableId === "slide-2-table-1"));
  assert.deepEqual(audit.slideSize, { width: 12_192_000, height: 6_858_000 });
  assert.ok(audit.layoutReviews.some((review) => review.rule === "overflow-risk"));
  assert.ok(audit.layoutReviews.some((review) => review.rule === "off-slide"));
  assert.ok(audit.fonts.some((font) => font.normalizedFamily === "century gothic"));
  assert.ok(audit.fonts.some((font) => font.normalizedFamily === "arial"));
  assert.ok(audit.findings.some((finding) => finding.ruleId === "font.legacy.century-gothic"));
  assert.equal(audit.containsMacros, false);
  assert.equal(audit.containsOleObjects, false);
});

test("object geometry proposals move native objects while preserving exact slide and table content", async () => {
  const bytes = await fixtureBytes();
  const audit = await auditPptx(bytes);
  const table = audit.editableObjects.find((object) => object.kind === "table");
  assert.ok(table);
  const deck: DeckJob = { id: "deck-geometry", name: "synthetic.pptx", sourceResourceId: "resource-geometry", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: "older-or-modified-ornl", targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const source = table.geometry;
  const proposal = createGeometryEditProposal(deck, "2026-08-11T12:00:00.000Z", { objectId: table.id, target: { x: source.x + 91_440, y: source.y, width: source.width - 91_440, height: source.height }, rationale: "Align the native table to the content grid.", author: "ai" });
  assert.equal(proposal.mode, "slide-geometry");
  assert.deepEqual(proposal.changes.map((change) => change.kind), ["geometry"]);
  assert.ok(proposal.changes.some((change) => change.kind === "geometry"));
  const preview = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(preview.bytes);
  const moved = after.editableObjects.find((object) => object.id === table.id);
  assert.equal(preview.geometryCount, 1);
  assert.equal(moved?.geometry.x, source.x + 91_440);
  assert.equal(moved?.geometry.width, source.width - 91_440);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
  assert.deepEqual(after.tables.map((item) => item.contentHash), audit.tables.map((item) => item.contentHash));
  assert.deepEqual(after.tables.map((item) => item.structureHash), audit.tables.map((item) => item.structureHash));
  assert.throws(() => createGeometryEditProposal(deck, "2026-08-11T12:00:00.000Z", { objectId: table.id, target: { x: audit.slideSize.width, y: 0, width: source.width, height: source.height }, rationale: "Invalid", author: "ai" }), /physical slide boundary/);
  const appliedDeck: DeckJob = { ...deck, proposal: { ...createDesignerCleanupProposal(deck, "2026-08-11T11:00:00.000Z"), status: "applied" } };
  const afterApplied = createGeometryEditProposal(appliedDeck, "2026-08-11T12:00:00.000Z", { objectId: table.id, target: { x: source.x + 91_440, y: source.y, width: source.width - 91_440, height: source.height }, rationale: "Do not revive the applied cleanup.", author: "ai" });
  assert.deepEqual(afterApplied.changes.map((change) => change.kind), ["geometry"]);
});

test("atomic geometry validation rejects new collisions and worsened text fit before materialization", async () => {
  const audit = await auditPptx(await fixtureBytes());
  const deck: DeckJob = { id: "deck-geometry-constraints", name: "synthetic.pptx", sourceResourceId: "resource-geometry-constraints", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: "older-or-modified-ornl", targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const textObjects = audit.editableObjects.filter((object) => object.kind === "text" && object.canResize);
  const target = textObjects.find((object) => textObjects.some((other) => other.slideNumber === object.slideNumber && other.id !== object.id));
  const other = target && textObjects.find((object) => object.slideNumber === target.slideNumber && object.id !== target.id);
  assert.ok(target && other);
  assert.throws(() => createGeometryBatchProposal(deck, "2026-08-11T12:00:00.000Z", [{ objectId: target.id, target: { x: other.geometry.x, y: other.geometry.y, width: target.geometry.width, height: target.geometry.height }, rationale: "Invalid collision", author: "ai" }]), /overlap/i);
  const fitObject = textObjects.find((object) => audit.textBoxes.some((textBox) => textBox.shapeId === object.shapeId && textBox.slideNumber === object.slideNumber));
  assert.ok(fitObject);
  assert.throws(() => createGeometryEditProposal(deck, "2026-08-11T12:00:00.000Z", { objectId: fitObject.id, target: { x: fitObject.geometry.x, y: fitObject.geometry.y, width: fitObject.geometry.width, height: 91_440 }, rationale: "Invalid fit", author: "ai" }), /text fit/i);
});

test("designer cleanup reviews every slide and normalizes compatible native tables without changing their meaning", async () => {
  const bytes = await fixtureBytes();
  const audit = await auditPptx(bytes);
  const deck: DeckJob = {
    id: "deck-designer",
    name: "synthetic.pptx",
    sourceResourceId: "resource-designer",
    sourceSha256: "0".repeat(64),
    operationScope: "reflow",
    templateClassification: "older-or-modified-ornl",
    targetTemplateId: "ornl-16x9-v1",
    targetTemplateConfirmedAt: new Date().toISOString(),
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
  const proposal = createDesignerCleanupProposal(deck, "2026-08-11T12:00:00.000Z");
  assert.equal(proposal.mode, "designer-cleanup");
  assert.equal(proposal.slideDispositions.length, audit.slideCount);
  assert.equal(proposal.slideDispositions.find((item) => item.slideNumber === 1)?.status, "change-proposed");
  assert.equal(proposal.slideDispositions.find((item) => item.slideNumber === 2)?.status, "needs-review");
  assert.ok(proposal.changes.some((change) => change.kind === "table-style"));
  assert.ok(proposal.changes.some((change) => change.kind === "alignment"));
  const preview = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(preview.bytes);
  assert.equal(preview.tableCount, 1);
  assert.equal(preview.alignmentCount, 2);
  assert.equal(after.tables[0].cellFonts.includes("Aptos"), true);
  assert.equal(after.tables[0].colorTokens.includes("00454d"), true);
  assert.equal(after.tables[0].marginSignatures.includes("marL:76200|marR:76200|marT:50800|marB:50800|anchor:ctr"), true);
  assert.equal(after.tables[0].contentHash, audit.tables[0].contentHash);
  assert.equal(after.tables[0].structureHash, audit.tables[0].structureHash);
  assert.notEqual(after.tables[0].styleFingerprint, audit.tables[0].styleFingerprint);
  assert.equal(after.alignmentRepairs.length, 0);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
  assert.deepEqual(after.textBoxes.map((textBox) => textBox.textHash), audit.textBoxes.map((textBox) => textBox.textHash));
});

test("dense technical tables use one compact ORNL component without changing table content or structure", async () => {
  const bytes = await fixtureBytes();
  const audit = await auditPptx(bytes);
  const revision = "2026-08-12T20:00:00.000Z";
  const deck: DeckJob = { id: "deck-dense-table", name: "synthetic.pptx", sourceResourceId: "resource-dense-table", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const sourceTable = audit.tables[0];
  assert.ok(sourceTable);
  const proposal = createTableStyleProposal(deck, revision, { tableIds: [sourceTable.id], variant: "dense-technical" });
  const preview = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(preview.bytes);
  assert.equal(preview.tableCount, 1);
  assert.equal(after.tables[0].contentHash, sourceTable.contentHash);
  assert.equal(after.tables[0].structureHash, sourceTable.structureHash);
  const denseTokens = PRESENTATION_DESIGN_STANDARD.tableVariants.denseTechnical;
  const horizontalMargin = Math.round(denseTokens.horizontalPaddingPt * 12_700);
  const verticalMargin = Math.round(denseTokens.verticalPaddingPt * 12_700);
  assert.equal(after.tables[0].marginSignatures.includes(`marL:${horizontalMargin}|marR:${horizontalMargin}|marT:${verticalMargin}|marB:${verticalMargin}|anchor:ctr`), true);
  const zip = await JSZip.loadAsync(preview.bytes);
  const xml = await zip.file("ppt/slides/slide2.xml")?.async("text");
  assert.match(xml ?? "", /sz="1000"/);
  assert.match(xml ?? "", /sz="1050"/);
  assert.doesNotMatch(xml ?? "", /<a:tableStyleId\b/);
});

test("designer cleanup preserves semantic table colors as explicit review exceptions", async () => {
  const audit = await auditPptx(await fixtureBytes());
  const deck: DeckJob = {
    id: "deck-semantic-table",
    name: "synthetic.pptx",
    sourceResourceId: "resource-semantic-table",
    sourceSha256: "0".repeat(64),
    operationScope: "reflow",
    templateClassification: "older-or-modified-ornl",
    targetTemplateId: "ornl-16x9-v1",
    targetTemplateConfirmedAt: new Date().toISOString(),
    status: "ready-for-cleanup",
    audit: { ...audit, tables: audit.tables.map((table) => ({ ...table, colorTokens: [...table.colorTokens, "accent6"] })) },
    protectedSlideNumbers: [],
  };
  const proposal = createDesignerCleanupProposal(deck, "2026-08-11T12:00:00.000Z");
  assert.equal(proposal.tableExceptions[0]?.rule, "semantic-color");
  assert.equal(proposal.slideDispositions.find((item) => item.slideNumber === 2)?.status, "needs-review");
  assert.equal(proposal.changes.some((change) => change.kind === "table-style"), false);
});

test("designer cleanup does not force the 16-point table profile onto text-dense cells", async () => {
  const audit = await auditPptx(await fixtureBytes());
  const deck: DeckJob = {
    id: "deck-dense-table",
    name: "synthetic.pptx",
    sourceResourceId: "resource-dense-table",
    sourceSha256: "0".repeat(64),
    operationScope: "reflow",
    templateClassification: "older-or-modified-ornl",
    targetTemplateId: "ornl-16x9-v1",
    targetTemplateConfirmedAt: new Date().toISOString(),
    status: "ready-for-cleanup",
    audit: { ...audit, tables: audit.tables.map((table) => ({ ...table, totalCellCharacterCount: 900, maximumCellCharacterCount: 190 })) },
    protectedSlideNumbers: [],
  };
  const proposal = createDesignerCleanupProposal(deck, "2026-08-11T12:00:00.000Z");
  assert.equal(proposal.tableExceptions[0]?.rule, "dense-table");
  assert.equal(proposal.changes.some((change) => change.kind === "table-style"), false);
});

test("font cleanup creates a new PPTX while preserving all visible text", async () => {
  const bytes = await fixtureBytes();
  const audit = await auditPptx(bytes);
  const deck: DeckJob = {
    id: "deck-synthetic",
    name: "synthetic.pptx",
    sourceResourceId: "resource-synthetic",
    sourceSha256: "0".repeat(64),
    operationScope: "cleanup-only",
    templateClassification: "older-or-modified-ornl",
    targetTemplateId: "ornl-16x9-v1",
    targetTemplateConfirmedAt: new Date().toISOString(),
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
  const proposal = createFontCleanupProposal(deck, "2026-08-11T12:00:00.000Z");
  await assert.rejects(() => applyCleanupToPptx(bytes, proposal), /Accept the cleanup plan/);
  const output = await applyCleanupToPptx(bytes, { ...proposal, status: "applied" });
  const after = await auditPptx(output.bytes);
  assert.ok(output.replacementCount >= 2);
  assert.equal(after.slideCount, audit.slideCount);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), audit.slides.map((slide) => slide.textHash));
  assert.notDeepEqual(output.bytes, bytes);
});

test("audit report excludes slide and picture content", async () => {
  const audit = await auditPptx(await fixtureBytes());
  const project = createProject("Synthetic report test");
  const deck: DeckJob = { id: "deck-report", name: "synthetic.pptx", sourceResourceId: "resource-report", sourceSha256: "0".repeat(64), operationScope: "audit-only", templateClassification: audit.classification, status: "audited", audit, protectedSlideNumbers: [] };
  const report = new TextDecoder().decode(buildAuditReport(project, deck));
  assert.match(report, /presentation-studio\/audit-report/);
  assert.doesNotMatch(report, /Every sentence in this file/);
  assert.doesNotMatch(report, /Visible text must remain/);
});

test("older saved audits and proposals adopt empty geometry collections", async () => {
  const audit = await auditPptx(await fixtureBytes());
  const deck: DeckJob = { id: "deck-legacy-geometry", name: "synthetic.pptx", sourceResourceId: "resource-legacy-geometry", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const proposal = createDesignerCleanupProposal(deck, new Date().toISOString());
  const legacyProject = createProject("Legacy geometry test") as unknown as Record<string, unknown> & { decks: Array<Record<string, unknown>> };
  legacyProject.decks = [{ ...deck, audit: { ...audit }, proposal: { ...proposal } }];
  delete (legacyProject.decks[0].audit as Record<string, unknown>).textBoxes;
  delete (legacyProject.decks[0].audit as Record<string, unknown>).layoutReviews;
  delete (legacyProject.decks[0].audit as Record<string, unknown>).editableObjects;
  delete (legacyProject.decks[0].audit as Record<string, unknown>).slideSize;
  delete (legacyProject.decks[0].proposal as Record<string, unknown>).layoutExceptions;
  delete (legacyProject.decks[0].proposal as Record<string, unknown>).slideReviews;
  const parsed = projectSchema.parse(legacyProject);
  assert.deepEqual(parsed.decks[0].audit?.textBoxes, []);
  assert.deepEqual(parsed.decks[0].audit?.layoutReviews, []);
  assert.deepEqual(parsed.decks[0].audit?.editableObjects, []);
  assert.deepEqual(parsed.decks[0].audit?.slideSize, { width: 12_192_000, height: 6_858_000 });
  assert.deepEqual(parsed.decks[0].proposal?.layoutExceptions, []);
  assert.deepEqual(parsed.decks[0].proposal?.slideReviews, []);
});
