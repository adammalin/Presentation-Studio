import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { applyCleanupToPptx, buildCleanupProposalPptx, createDesignerCleanupProposal, createFontCleanupProposal } from "../src/lib/cleanup";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildAuditReport } from "../src/lib/report";
import { createProject } from "../src/lib/project";
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
  assert.equal(audit.alignmentRepairs.length, 1);
  assert.equal(audit.alignmentRepairs[0].ruleId, "cover.dominant-left-edge");
  assert.ok(audit.fonts.some((font) => font.normalizedFamily === "century gothic"));
  assert.ok(audit.fonts.some((font) => font.normalizedFamily === "arial"));
  assert.ok(audit.findings.some((finding) => finding.ruleId === "font.legacy.century-gothic"));
  assert.equal(audit.containsMacros, false);
  assert.equal(audit.containsOleObjects, false);
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
  assert.ok(proposal.slideDispositions.every((item) => item.status === "change-proposed"));
  assert.ok(proposal.changes.some((change) => change.kind === "table-style"));
  assert.ok(proposal.changes.some((change) => change.kind === "alignment"));
  const preview = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(preview.bytes);
  assert.equal(preview.tableCount, 1);
  assert.equal(preview.alignmentCount, 1);
  assert.equal(after.tables[0].cellFonts.includes("Aptos"), true);
  assert.equal(after.tables[0].colorTokens.includes("00454d"), true);
  assert.equal(after.tables[0].marginSignatures.includes("marL:76200|marR:76200|marT:50800|marB:50800|anchor:ctr"), true);
  assert.equal(after.tables[0].contentHash, audit.tables[0].contentHash);
  assert.equal(after.tables[0].structureHash, audit.tables[0].structureHash);
  assert.notEqual(after.tables[0].styleFingerprint, audit.tables[0].styleFingerprint);
  assert.equal(after.alignmentRepairs.length, 0);
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
