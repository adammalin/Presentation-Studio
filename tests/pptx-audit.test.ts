import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { applyCleanupToPptx, buildCleanupProposalPptx, createDesignerCleanupProposal, createFontCleanupProposal, createGeometryBatchProposal, createGeometryEditProposal, createTableStyleProposal, withDesignerCleanupScope } from "../src/lib/cleanup";
import { createOrnlDesignProfile, PRESENTATION_DESIGN_STANDARD } from "../src/lib/design-standard";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildAuditReport } from "../src/lib/report";
import { createProject, projectSchema } from "../src/lib/project";
import type { DeckJob } from "../src/types";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { deckTemplateWorkflow, deckWithAutomaticTemplateRouting } from "../src/lib/template-routing";

async function fixtureBytes(): Promise<Uint8Array> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-pptx-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  return new Uint8Array(await fs.readFile(filePath));
}

async function semanticTableFixtureBytes(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixtureBytes());
  const entry = zip.file("ppt/slides/slide2.xml");
  assert.ok(entry);
  const xml = await entry.async("text");
  let rowIndex = 0;
  const next = xml.replace(/<a:tbl\b[\s\S]*?<\/a:tbl>/, (table) => table.replace(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g, (row) => {
    const token = rowIndex === 1 ? "accent6" : rowIndex === 2 ? "accent3" : undefined;
    rowIndex += 1;
    if (!token) return row;
    return row.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => {
      const fill = `<a:solidFill><a:schemeClr val="${token}"/></a:solidFill>`;
      if (/<a:tcPr\b[^>]*\/>/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)\/>/, `<a:tcPr$1>${fill}</a:tcPr>`);
      if (/<a:tcPr\b/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)>([\s\S]*?)<\/a:tcPr>/, (_tag, attributes, children) => `<a:tcPr${attributes}>${String(children).replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>/g, "")}${fill}</a:tcPr>`);
      return cell.replace(/<\/a:tc>/, `<a:tcPr>${fill}</a:tcPr></a:tc>`);
    });
  }));
  zip.file("ppt/slides/slide2.xml", next);
  return zip.generateAsync({ type: "uint8array" });
}

async function alternateContentFixtureBytes(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixtureBytes());
  const entry = zip.file("ppt/slides/slide1.xml");
  assert.ok(entry);
  const xml = await entry.async("text");
  const shape = [...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].map((match) => match[0]).find((candidate) => /<p:txBody\b/.test(candidate) && /<a:t>/.test(candidate));
  assert.ok(shape);
  const fallback = shape.replace(/<a:t>([\s\S]*?)<\/a:t>/, "<a:t>FALLBACK ONLY</a:t>");
  const alternate = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><mc:Choice Requires="a14">${shape}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  zip.file("ppt/slides/slide1.xml", xml.replace(shape, alternate));
  return zip.generateAsync({ type: "uint8array" });
}

async function fixtureWithOrnlSlideCopyAndSponsorTheme(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixtureBytes());
  const slide = zip.file("ppt/slides/slide1.xml");
  const theme = zip.file("ppt/theme/theme1.xml");
  assert.ok(slide && theme);
  zip.file("ppt/slides/slide1.xml", (await slide.async("text")).replace("Legacy typography should be auditable", "ORNL legacy typography should be auditable"));
  const themeXml = await theme.async("text");
  zip.file("ppt/theme/theme1.xml", themeXml
    .replace(/<a:clrScheme\b([^>]*)\bname=(?:"[^"]*"|'[^']*')/, '<a:clrScheme$1 name="EERE-DOE-2025"')
    .replace(/<\/a:theme>/, '<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="EERE-Presentation-Toolkit.pptx"/></a:theme>'));
  return zip.generateAsync({ type: "uint8array" });
}

async function fixtureWithOrnlSlideCopyOnly(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixtureBytes());
  const slide = zip.file("ppt/slides/slide1.xml");
  assert.ok(slide);
  zip.file("ppt/slides/slide1.xml", (await slide.async("text")).replace("Legacy typography should be auditable", "ORNL legacy typography should be auditable"));
  return zip.generateAsync({ type: "uint8array" });
}

async function fixtureWithBlockingExternalImage(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await fixtureBytes());
  const relationships = zip.file("ppt/slides/_rels/slide1.xml.rels");
  assert.ok(relationships);
  const xml = await relationships.async("text");
  zip.file("ppt/slides/_rels/slide1.xml.rels", xml.replace("</Relationships>", '<Relationship Id="rIdLinkedImageTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/linked-image.png" TargetMode="External"/></Relationships>'));
  return zip.generateAsync({ type: "uint8array" });
}

test("OOXML audit selects the active AlternateContent branch instead of duplicating PowerPoint shapes", async () => {
  const audit = await auditPptx(await alternateContentFixtureBytes());
  const slideObjects = audit.editableObjects.filter((object) => object.slideNumber === 1);
  const objectIds = slideObjects.map((object) => object.id);
  assert.equal(new Set(objectIds).size, objectIds.length);
  const textBoxShapeIds = audit.textBoxes.filter((textBox) => textBox.slideNumber === 1).map((textBox) => textBox.shapeId);
  assert.ok(textBoxShapeIds.length > 0);
  assert.equal(new Set(textBoxShapeIds).size, textBoxShapeIds.length);
  assert.doesNotMatch(audit.slides[0].text, /FALLBACK ONLY/);
});

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
  assert.equal(audit.tables[0].rows?.length, 3);
  assert.equal(audit.tables[0].columns?.length, 3);
  assert.equal(audit.tables[0].cells?.length, 9);
  assert.equal(audit.tables[0].cells?.[0].id, "slide-2-table-1-cell-r1-c1");
  assert.equal(audit.tables[0].cells?.[0].text, "Category");
  assert.match(audit.tables[0].cells?.[0].textHash ?? "", /^[0-9a-f]{64}$/);
  assert.equal(audit.alignmentRepairs.length, 2);
  assert.ok(audit.alignmentRepairs.some((repair) => repair.ruleId === "cover.dominant-left-edge"));
  assert.ok(audit.alignmentRepairs.some((repair) => repair.ruleId === "peer.dominant-left-edge"));
  assert.ok(audit.textBoxes.length >= 10);
  const bulletTextBox = audit.textBoxes.find((textBox) => textBox.text.startsWith("This deliberately dense"));
  assert.ok(bulletTextBox);
  assert.equal(bulletTextBox.bulletParagraphCount, 1);
  assert.ok(bulletTextBox.paragraphs.length >= 1);
  assert.equal(bulletTextBox.paragraphs.map((paragraph) => paragraph.text).join(" "), bulletTextBox.text);
  assert.equal(bulletTextBox.paragraphs.some((paragraph) => paragraph.bullet), true);
  assert.equal(bulletTextBox.paragraphs.every((paragraph) => /^[0-9a-f]{64}$/.test(paragraph.textHash)), true);
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
  assert.equal(audit.containsExternalRelationships, true);
  assert.equal(audit.externalHyperlinkCount, 1);
  assert.equal(audit.blockingExternalRelationshipCount, 0);
  assert.equal(audit.containsBlockingExternalRelationships, false);
  assert.ok(audit.findings.some((finding) => finding.ruleId === "production.external-hyperlinks"));
  assert.equal(audit.findings.some((finding) => finding.ruleId === "production.advanced-content"), false);
});

test("ordinary hyperlinks remain editable while linked media stays blocked", async () => {
  const hyperlinkBytes = await fixtureBytes();
  const hyperlinkAudit = await auditPptx(hyperlinkBytes);
  const revision = "2026-08-18T19:00:00.000Z";
  const hyperlinkDeck: DeckJob = { id: "deck-hyperlinks", name: "hyperlinks.pptx", sourceResourceId: "resource-hyperlinks", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: hyperlinkAudit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit: hyperlinkAudit, protectedSlideNumbers: [] };
  const proposal = createDesignerCleanupProposal(hyperlinkDeck, revision);
  const preview = await buildCleanupProposalPptx(hyperlinkBytes, proposal);
  const beforeZip = await JSZip.loadAsync(hyperlinkBytes);
  const afterZip = await JSZip.loadAsync(preview.bytes);
  const beforeRelationships = await beforeZip.file("ppt/slides/_rels/slide1.xml.rels")?.async("text");
  const afterRelationships = await afterZip.file("ppt/slides/_rels/slide1.xml.rels")?.async("text");
  assert.equal(afterRelationships, beforeRelationships);
  assert.match(afterRelationships ?? "", /relationships\/hyperlink/);
  assert.match(afterRelationships ?? "", /TargetMode="External"/);

  const linkedAudit = await auditPptx(await fixtureWithBlockingExternalImage());
  assert.equal(linkedAudit.externalHyperlinkCount, 1);
  assert.equal(linkedAudit.blockingExternalRelationshipCount, 1);
  assert.equal(linkedAudit.containsBlockingExternalRelationships, true);
  const linkedDeck: DeckJob = { ...hyperlinkDeck, id: "deck-linked-media", audit: linkedAudit };
  assert.throws(() => createDesignerCleanupProposal(linkedDeck, revision), /externally linked content/i);
});

test("template classification preserves sponsor source identity while ORNL remains the automatic design target", async () => {
  const audit = await auditPptx(await fixtureWithOrnlSlideCopyAndSponsorTheme());
  assert.equal(audit.classification, "sponsor");
  assert.ok(audit.classificationEvidence.some((item) => /theme identity.*sponsor/i.test(item)));
  const base: DeckJob = { id: "sponsor-routing", name: "sponsor.pptx", sourceResourceId: "resource", sourceSha256: "0".repeat(64), operationScope: "cleanup-only", templateClassification: "older-or-modified-ornl", targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: "2026-08-18T12:00:00.000Z", designProfile: createOrnlDesignProfile("automatic-default", "2026-08-18T12:00:00.000Z"), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const routed = deckWithAutomaticTemplateRouting({ deck: base, audit, adoptedAt: "2026-08-18T13:00:00.000Z" });
  assert.equal(routed.targetTemplateId, "ornl-16x9-v1");
  assert.equal(routed.targetTemplateDecisionSource, "automatic-default");
  assert.ok(routed.designProfile);
  assert.equal(deckTemplateWorkflow(routed), "ornl-studio");
  assert.equal(routed.status, "ready-for-cleanup");
});

test("Designer Cleanup enables source-bound sponsor reflow without converting its template", async () => {
  const audit = await auditPptx(await fixtureWithOrnlSlideCopyAndSponsorTheme());
  const revision = "2026-08-18T12:00:00.000Z";
  const deck: DeckJob = { id: "sponsor-designer", name: "sponsor.pptx", sourceResourceId: "resource", sourceSha256: "0".repeat(64), operationScope: "cleanup-only", templateClassification: "sponsor", targetTemplateId: "sponsor-source", targetTemplateDecisionSource: "user-selected", targetTemplateConfirmedAt: revision, status: "audited", audit, protectedSlideNumbers: [] };
  const designerDeck = withDesignerCleanupScope(deck);
  assert.equal(designerDeck.operationScope, "reflow");
  assert.equal(designerDeck.targetTemplateId, "sponsor-source");
  assert.equal(designerDeck.templateClassification, "sponsor");
  const proposal = createDesignerCleanupProposal(designerDeck, revision);
  assert.equal(proposal.changes.some((change) => change.kind === "font-family" || change.kind === "table-style"), false);
});

test("ordinary ORNL slide copy alone does not alter source classification while ORNL remains the product default", async () => {
  const audit = await auditPptx(await fixtureWithOrnlSlideCopyOnly());
  assert.equal(audit.classification, "unknown");
  assert.ok(audit.classificationEvidence.some((item) => /ordinary slide copy|package copy/i.test(item)));
  const deck: DeckJob = { id: "unknown-source", name: "unknown.pptx", sourceResourceId: "resource", sourceSha256: "0".repeat(64), operationScope: "cleanup-only", templateClassification: "unknown", status: "audited", audit, protectedSlideNumbers: [] };
  const routed = deckWithAutomaticTemplateRouting({ deck, audit, adoptedAt: "2026-08-18T13:00:00.000Z" });
  assert.equal(routed.targetTemplateId, "ornl-16x9-v1");
  assert.equal(deckTemplateWorkflow(routed), "ornl-studio");
});

test("an explicit user-selected ORNL conversion survives later sponsor detection", async () => {
  const audit = await auditPptx(await fixtureWithOrnlSlideCopyAndSponsorTheme());
  const deck: DeckJob = { id: "explicit-conversion", name: "sponsor.pptx", sourceResourceId: "resource", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: "sponsor", targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "user-selected", targetTemplateConfirmedAt: "2026-08-18T12:00:00.000Z", designProfile: createOrnlDesignProfile("user-selected", "2026-08-18T12:00:00.000Z"), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const routed = deckWithAutomaticTemplateRouting({ deck, audit, adoptedAt: "2026-08-18T13:00:00.000Z" });
  assert.equal(routed.targetTemplateId, "ornl-16x9-v1");
  assert.equal(routed.targetTemplateDecisionSource, "user-selected");
  assert.equal(deckTemplateWorkflow(routed), "ornl-studio");
});

test("an explicit user-selected source-template override survives automatic ORNL routing", async () => {
  const audit = await auditPptx(await fixtureWithOrnlSlideCopyAndSponsorTheme());
  const deck: DeckJob = { id: "explicit-source", name: "sponsor.pptx", sourceResourceId: "resource", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: "sponsor", targetTemplateId: "sponsor-source", targetTemplateDecisionSource: "user-selected", targetTemplateConfirmedAt: "2026-08-18T12:00:00.000Z", status: "audited", audit, protectedSlideNumbers: [] };
  const routed = deckWithAutomaticTemplateRouting({ deck, audit, adoptedAt: "2026-08-18T13:00:00.000Z" });
  assert.equal(routed.targetTemplateId, "sponsor-source");
  assert.equal(routed.targetTemplateDecisionSource, "user-selected");
  assert.equal(deckTemplateWorkflow(routed), "source-template-cleanup");
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
  const proposal = createTableStyleProposal(deck, revision, { tableIds: [sourceTable.id], variant: "dense-technical", semanticColorPolicy: "preserve-source" });
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

test("shared table components preserve semantic roles with approved ORNL tints", async () => {
  const bytes = await semanticTableFixtureBytes();
  const audit = await auditPptx(bytes);
  const sourceTable = audit.tables[0];
  assert.deepEqual(sourceTable.semanticColorTokens, ["accent3", "accent6"]);
  assert.equal(sourceTable.cells?.filter((cell) => cell.semanticColorRole === "accent6").length, 3);
  assert.equal(sourceTable.cells?.filter((cell) => cell.semanticColorRole === "accent3").length, 3);
  const revision = "2026-08-13T12:00:00.000Z";
  const deck: DeckJob = { id: "deck-semantic-tints", name: "semantic.pptx", sourceResourceId: "resource-semantic-tints", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const proposal = createTableStyleProposal(deck, revision, { tableIds: [sourceTable.id], variant: "standard", semanticColorPolicy: "preserve-source" });
  const preview = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(preview.bytes);
  const outputTable = after.tables[0];
  assert.ok(outputTable.semanticColorTokens.includes("accent3"));
  assert.ok(outputTable.semanticColorTokens.includes("accent6"));
  assert.ok(outputTable.cells?.filter((cell) => cell.semanticColorRole === "accent6").every((cell) => cell.fillToken === "srgb:fbc9df"));
  assert.ok(outputTable.cells?.filter((cell) => cell.semanticColorRole === "accent3").every((cell) => cell.fillToken === "srgb:c8fbdd"));
  assert.equal(outputTable.contentHash, sourceTable.contentHash);
  assert.equal(outputTable.structureHash, sourceTable.structureHash);
  const unsafe = { ...proposal, changes: proposal.changes.map((change) => change.kind === "table-style" ? { ...change, semanticColorPolicy: undefined } : change) };
  await assert.rejects(() => buildCleanupProposalPptx(bytes, unsafe), /semantic-color preservation policy/i);
});

test("restaging a table upgrades retained legacy table changes to semantic preservation", async () => {
  const bytes = await fixtureBytes();
  const audit = await auditPptx(bytes);
  const revision = "2026-08-13T12:30:00.000Z";
  const deck: DeckJob = { id: "deck-table-upgrade", name: "legacy-table.pptx", sourceResourceId: "resource-table-upgrade", sourceSha256: "0".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateConfirmedAt: revision, status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  const first = createTableStyleProposal(deck, revision, { tableIds: [audit.tables[0].id], variant: "standard", semanticColorPolicy: "preserve-source" });
  const legacy = { ...first, status: "applied" as const, changes: first.changes.map((change) => change.kind === "table-style" ? { ...change, tableIds: ["legacy-retained-table", audit.tables[0].id], semanticColorPolicy: undefined } : change) };
  const restaged = createTableStyleProposal({ ...deck, proposal: legacy }, revision, { tableIds: [audit.tables[0].id], variant: "dense-technical", semanticColorPolicy: "preserve-source" });
  assert.ok(restaged.changes.filter((change) => change.kind === "table-style").every((change) => change.semanticColorPolicy === "preserve-source"));
  assert.ok(restaged.changes.some((change) => change.kind === "table-style" && change.tableIds?.includes("legacy-retained-table")));
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
