import assert from "node:assert/strict";
import { XMLValidator } from "fast-xml-parser";
import fs from "node:fs/promises";
import JSZip from "jszip";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createDesignCanaryDeck } from "../scripts/create-design-canary";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createTableLayoutProposal } from "../src/lib/cleanup";
import { solveAlignment, solveDistribution, solveGroupLayout, solveSafeRegion, solveSceneToLayout } from "../src/lib/layout-solver";
import { bindNativeMeasurement } from "../src/lib/native-measurement";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { recommendedTableGrowthPlan, solveTableLayout } from "../src/lib/table-layout-solver";
import type { DeckJob } from "../src/types";

async function fixture() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-solver-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const audit = await auditPptx(bytes);
  const deck: DeckJob = { id: "solver-deck", name: "synthetic.pptx", sourceResourceId: "solver-source", sourceSha256: "a".repeat(64), operationScope: "reflow", templateClassification: "custom", targetTemplateId: "synthetic-template", targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  return { bytes, audit, deck, measurement: bindNativeMeasurement(deck) };
}

async function canaryFixture() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-design-canary-"));
  const filePath = path.join(directory, "canary.pptx");
  await createDesignCanaryDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const audit = await auditPptx(bytes);
  const deck: DeckJob = { id: "canary-solver", name: "canary.pptx", sourceResourceId: "canary-source", sourceSha256: "c".repeat(64), operationScope: "reflow", templateClassification: "custom", targetTemplateId: "synthetic-canary", targetTemplateConfirmedAt: "2026-08-12T20:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  return { deck, measurement: bindNativeMeasurement(deck) };
}

test("optical alignment solver anchors visible text starts without AI coordinate arithmetic", async () => {
  const { deck, audit, measurement } = await fixture();
  const anchorText = audit.textBoxes.find((item) => item.text === "Aligned peer label two");
  const outlierText = audit.textBoxes.find((item) => item.text === "Offset peer label four");
  const anchor = deck.scene?.objects.find((object) => object.shapeId === anchorText?.shapeId && object.slideNumber === anchorText?.slideNumber);
  const outlier = deck.scene?.objects.find((object) => object.shapeId === outlierText?.shapeId && object.slideNumber === outlierText?.slideNumber);
  assert.ok(anchor && outlier);
  const solved = solveAlignment({ deck, measurement, slideNumber: 2, objectIds: [anchor.id, outlier.id], anchorObjectId: anchor.id, mode: "optical-left", rationale: "Align the visible starts." });
  assert.equal(solved.status, "solved");
  assert.equal(solved.commands.length, 1);
  assert.equal(solved.commands[0].objectId, outlier.id);
  assert.equal(solved.commands[0].target.x, anchor.geometry.x);
  assert.equal(solved.commands[0].target.y, outlier.geometry.y);
});

test("cell-level table solver materializes row, column, and margin changes without content loss", async () => {
  const { bytes, audit, deck, measurement } = await fixture();
  const table = audit.tables[0];
  assert.ok(table.cells?.length && table.rows?.length && table.columns?.length);
  const standard = solveTableLayout({ deck, measurement, tableId: table.id, rationale: "Check the standard ORNL table floor.", variant: "standard" });
  assert.equal(standard.status, "infeasible");
  assert.match(standard.diagnostics.reasons.join(" "), /below the 16 pt standard floor/i);
  const solved = solveTableLayout({ deck, measurement, tableId: table.id, rationale: "Fit the synthetic native table with readable padding.", variant: "dense-technical" });
  assert.equal(solved.status, "solved");
  assert.ok(solved.command);
  assert.equal(solved.command.constraints.minimumFontPt, 10);
  assert.equal(solved.command.constraints.minimumHorizontalPaddingPt, 6);
  assert.equal(solved.command.constraints.minimumVerticalPaddingPt, 4);
  const proposal = createTableLayoutProposal(deck, "2026-08-12T20:00:00.000Z", solved.command);
  const output = await buildCleanupProposalPptx(bytes, proposal);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const slideXml = await outputZip.file("ppt/slides/slide1.xml")!.async("text");
  assert.equal(XMLValidator.validate(slideXml), true);
  assert.doesNotMatch(slideXml, /<a:gridCol\b[^>]*\/\/>/);
  const after = await auditPptx(output.bytes);
  assert.equal(output.tableLayoutCount, 1);
  assert.deepEqual(after.tables[0].columns?.map((column) => column.widthEmu), solved.command.columnWidthsEmu);
  assert.deepEqual(after.tables[0].rows?.map((row) => row.heightEmu), solved.command.rowHeightsEmu);
  assert.equal(after.tables[0].contentHash, table.contentHash);
  assert.equal(after.tables[0].structureHash, table.structureHash);
});

test("native table refinement retains the preceding geometry-growth transaction", async () => {
  const { deck, audit, measurement } = await fixture();
  const table = audit.tables[0];
  const object = deck.scene!.objects.find((item) => item.sourceLocator.tableId === table.id)!;
  const solved = solveTableLayout({ deck, measurement, tableId: table.id, rationale: "Fit the synthetic native table.", variant: "dense-technical" });
  assert.equal(solved.status, "solved");
  assert.ok(solved.command);
  const updatedAt = "2026-08-12T20:00:00.000Z";
  const growth = createGeometryBatchProposal(deck, updatedAt, [{ objectId: object.id, target: { ...object.geometry, height: object.geometry.height + 12_700 }, rationale: "Add one measured point of table height.", author: "ai", constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false } }]);
  const first = createTableLayoutProposal({ ...deck, proposal: growth }, updatedAt, solved.command);
  const refined = { ...solved.command, id: `${solved.command.id}-refined`, rowHeightsEmu: solved.command.rowHeightsEmu.map((value, index) => index === 0 ? value + 1 : value - 1) };
  const second = createTableLayoutProposal({ ...deck, proposal: first }, updatedAt, refined);
  assert.equal(second.changes.filter((change) => change.kind === "geometry").length, 1);
  assert.equal(second.changes.filter((change) => change.kind === "table-layout").length, 1);
  assert.equal(second.changes.find((change) => change.kind === "table-layout")?.tableLayoutCommands?.[0].id, refined.id);
});

test("table solver preserves a PowerPoint-native table that already meets the shared constraints", async () => {
  const { deck, measurement } = await canaryFixture();
  const table = deck.audit!.tables.find((item) => item.slideNumber === 9)!;
  const object = measurement.objects.find((item) => item.tableId === table.id)!;
  measurement.authority = "powerpoint-native";
  measurement.status = "ready";
  object.provenance.authority = "powerpoint-native";
  for (const cell of object.table!.cells) {
    cell.marginsPt = { left: 6, right: 6, top: 4, bottom: 4 };
    cell.clearancesPt = { left: 6, right: 6, top: 4, bottom: 4 };
    cell.lineCount = 1;
  }
  const result = solveTableLayout({ deck, measurement, tableId: table.id, rationale: "Preserve a compliant native table.", variant: "standard" });
  assert.equal(result.status, "already-fit");
  assert.equal(result.command, undefined);
  assert.match(result.diagnostics.recommendations.join(" "), /preserve it as-is/i);
});

test("table solver returns a minimum safe-region growth plan for a clipped short table", async () => {
  const { deck, measurement } = await canaryFixture();
  const table = deck.audit!.tables.find((item) => item.slideNumber === 10)!;
  const object = deck.scene!.objects.find((item) => item.sourceLocator.tableId === table.id)!;
  const result = solveTableLayout({ deck, measurement, tableId: table.id, rationale: "Create enough native row height without shrinking type.", variant: "standard" });
  assert.equal(result.status, "infeasible");
  const plan = recommendedTableGrowthPlan(result, object.id, "Grow the table by the measured minimum.");
  assert.ok(plan);
  assert.ok(plan.target.height > object.geometry.height);
  assert.ok(plan.target.y >= 18 * 12_700);
  assert.ok(plan.target.y + plan.target.height <= deck.audit!.slideSize.height - 18 * 12_700);
  const atomic = solveTableLayout({
    deck,
    measurement,
    tableId: table.id,
    rationale: "Grow the table frame and row grid in one transaction.",
    variant: "standard",
    targetBoundsPt: { width: plan.target.width / 12_700, height: plan.target.height / 12_700 },
  });
  assert.equal(atomic.status, "solved");
  assert.ok(atomic.command);
  assert.ok(Math.abs(atomic.command.rowHeightsEmu.reduce((sum, value) => sum + value, 0) - plan.target.height) <= 1);
});

test("grouped distribution preserves panel-content relationships while equalizing outer gaps", async () => {
  const { deck } = await canaryFixture();
  const groups = [1, 2, 3, 4].map((index) => [
    deck.scene!.objects.find((object) => object.name === `canary-5-panel-${index}`)!.id,
    deck.scene!.objects.find((object) => object.name === `canary-5-body-${index}`)!.id,
  ]);
  const solved = solveDistribution({ deck, slideNumber: 5, objectIds: groups.flat(), groups, mode: "vertical-equal-gap", rationale: "Normalize the repeated evidence-block rhythm as four intact groups." });
  assert.equal(solved.status, "solved");
  assert.ok(solved.commands.length >= 2);
  const commandByObject = new Map(solved.commands.map((command) => [command.objectId, command]));
  for (const group of groups) {
    const panel = deck.scene!.objects.find((object) => object.id === group[0])!;
    const body = deck.scene!.objects.find((object) => object.id === group[1])!;
    const panelDelta = (commandByObject.get(panel.id)?.target.y ?? panel.geometry.y) - panel.geometry.y;
    const bodyDelta = (commandByObject.get(body.id)?.target.y ?? body.geometry.y) - body.geometry.y;
    assert.equal(panelDelta, bodyDelta);
  }
  const finalPanels = groups.map(([panelId]) => {
    const panel = deck.scene!.objects.find((object) => object.id === panelId)!;
    const command = commandByObject.get(panelId);
    return { top: (command?.target.y ?? panel.geometry.y) / 12_700, height: panel.geometry.height / 12_700 };
  });
  const gaps = finalPanels.slice(1).map((panel, index) => panel.top - finalPanels[index].top - finalPanels[index].height);
  assert.ok(Math.max(...gaps) - Math.min(...gaps) <= .02);
});

test("safe-region solver moves an overlapping panel-content group by the minimum required distance", async () => {
  const { deck } = await canaryFixture();
  const panel = deck.scene!.objects.find((object) => object.name === "canary-6-safe-margin-minus-4pt")!;
  const copy = deck.scene!.objects.find((object) => object.name === "canary-6-safe-margin-copy")!;
  const solved = solveSafeRegion({ deck, slideNumber: 6, objectIds: [panel.id, copy.id], rationale: "Move the intact callout group to the 18-point safe margin." });
  assert.equal(solved.status, "solved");
  assert.equal(solved.commands.length, 2);
  const panelCommand = solved.commands.find((command) => command.objectId === panel.id)!;
  const copyCommand = solved.commands.find((command) => command.objectId === copy.id)!;
  assert.equal(panelCommand.target.x - panel.geometry.x, 4 * 12_700);
  assert.equal(copyCommand.target.x - copy.geometry.x, 4 * 12_700);
});

test("approved-region group layout composes intact components with design-rhythm spacing", async () => {
  const { deck } = await canaryFixture();
  const groups = [1, 2, 3, 4].map((index) => [
    deck.scene!.objects.find((object) => object.name === `canary-5-panel-${index}`)!.id,
    deck.scene!.objects.find((object) => object.name === `canary-5-body-${index}`)!.id,
  ]);
  const solved = solveGroupLayout({ deck, slideNumber: 5, groups, regionPt: { left: 51.84, top: 100, width: 446.4, height: 200 }, mode: "vertical-stack", alignment: "start", preferredGapPt: 18, rationale: "Compose repeated evidence components inside one approved semantic region." });
  assert.equal(solved.status, "solved");
  assert.ok(solved.commands.length > 0);
  const commandByObject = new Map(solved.commands.map((command) => [command.objectId, command]));
  const finalPanels = groups.map(([panelId]) => {
    const panel = deck.scene!.objects.find((object) => object.id === panelId)!;
    const command = commandByObject.get(panelId);
    return { top: (command?.target.y ?? panel.geometry.y) / 12_700, height: panel.geometry.height / 12_700 };
  });
  const gaps = finalPanels.slice(1).map((panel, index) => panel.top - finalPanels[index].top - finalPanels[index].height);
  assert.ok(gaps.every((gap) => Math.abs(gap - 18) <= .02));
  for (const [panelId, bodyId] of groups) {
    const panel = deck.scene!.objects.find((object) => object.id === panelId)!;
    const body = deck.scene!.objects.find((object) => object.id === bodyId)!;
    assert.equal((commandByObject.get(panelId)?.target.y ?? panel.geometry.y) - panel.geometry.y, (commandByObject.get(bodyId)?.target.y ?? body.geometry.y) - body.geometry.y);
  }
});

test("approved-region group layout preserves semantic hierarchy with primary and caption spacing", async () => {
  const { deck } = await canaryFixture();
  const groups = [1, 2, 3, 4].map((index) => [
    deck.scene!.objects.find((object) => object.name === `canary-5-panel-${index}`)!.id,
    deck.scene!.objects.find((object) => object.name === `canary-5-body-${index}`)!.id,
  ]);
  const solved = solveGroupLayout({
    deck,
    slideNumber: 5,
    groups,
    groupRoles: ["primary", "supporting", "supporting", "caption"],
    regionPt: { left: 51.84, top: 90, width: 446.4, height: 220 },
    mode: "vertical-stack",
    alignment: "start",
    preferredGapPt: 18,
    rationale: "Preserve one primary assertion, supporting evidence rhythm, and a tighter caption relationship.",
  });
  assert.equal(solved.status, "solved");
  const commandByObject = new Map(solved.commands.map((command) => [command.objectId, command]));
  const finalPanels = groups.map(([panelId]) => {
    const panel = deck.scene!.objects.find((object) => object.id === panelId)!;
    const command = commandByObject.get(panelId);
    return { top: (command?.target.y ?? panel.geometry.y) / 12_700, height: panel.geometry.height / 12_700 };
  });
  const gaps = finalPanels.slice(1).map((panel, index) => panel.top - finalPanels[index].top - finalPanels[index].height);
  assert.ok(Math.abs(gaps[0] - 24) <= .02, `Expected primary separation of 24 pt, received ${gaps[0]}.`);
  assert.ok(Math.abs(gaps[1] - 18) <= .02, `Expected supporting rhythm of 18 pt, received ${gaps[1]}.`);
  assert.ok(Math.abs(gaps[2] - 8) <= .02, `Expected compact caption relationship of 8 pt, received ${gaps[2]}.`);
});

test("approved-region group layout rejects ambiguous multiple-primary hierarchy", async () => {
  const { deck } = await canaryFixture();
  const groups = [1, 2].map((index) => [deck.scene!.objects.find((object) => object.name === `canary-5-panel-${index}`)!.id]);
  assert.throws(() => solveGroupLayout({ deck, slideNumber: 5, groups, groupRoles: ["primary", "primary"], regionPt: { left: 51.84, top: 90, width: 446.4, height: 220 }, mode: "vertical-stack", alignment: "start", rationale: "Invalid ambiguous hierarchy." }), /at most one primary/i);
});

test("approved-region group layout proportionally scales a resize-capable visual system above its readability floor", async () => {
  const { deck } = await canaryFixture();
  const panel = deck.scene!.objects.find((object) => object.name === "canary-14-large-right-panel")!;
  const copy = deck.scene!.objects.find((object) => object.name === "canary-14-right-copy")!;
  const solved = solveGroupLayout({
    deck,
    slideNumber: 14,
    groups: [[panel.id, copy.id]],
    groupRoles: ["primary"],
    regionPt: { left: 350, top: 100, width: 520, height: 240 },
    mode: "vertical-stack",
    alignment: "center",
    allowResponsiveScale: true,
    minimumScale: .75,
    rationale: "Fit the complete visual system proportionally inside the approved region.",
  });
  assert.equal(solved.status, "solved");
  const panelCommand = solved.commands.find((command) => command.objectId === panel.id)!;
  const copyCommand = solved.commands.find((command) => command.objectId === copy.id)!;
  const panelScale = panelCommand.target.width / panel.geometry.width;
  const copyScale = copyCommand.target.width / copy.geometry.width;
  assert.ok(panelScale < 1 && panelScale >= .75);
  assert.ok(Math.abs(panelScale - copyScale) < .001);
  assert.equal(panelCommand.operation, "move-and-resize");
  assert.equal(copyCommand.operation, "move-and-resize");
});

test("approved-region group layout refuses scaling below its declared floor", async () => {
  const { deck } = await canaryFixture();
  const panel = deck.scene!.objects.find((object) => object.name === "canary-14-large-right-panel")!;
  const copy = deck.scene!.objects.find((object) => object.name === "canary-14-right-copy")!;
  const solved = solveGroupLayout({ deck, slideNumber: 14, groups: [[panel.id, copy.id]], regionPt: { left: 350, top: 100, width: 300, height: 180 }, mode: "vertical-stack", alignment: "center", allowResponsiveScale: true, minimumScale: .75, rationale: "Do not force this system below its scale floor." });
  assert.equal(solved.status, "infeasible");
  assert.match(solved.diagnostics.join(" "), /below the allowed 75\.0% floor/i);
});

test("scene-to-layout solver validates two responsive regions as one atomic composition", async () => {
  const { deck } = await canaryFixture();
  const leftPanel = deck.scene!.objects.find((object) => object.name === "canary-14-small-left-panel")!;
  const leftLabel = deck.scene!.objects.find((object) => object.name === "canary-14-left-label")!;
  const rightPanel = deck.scene!.objects.find((object) => object.name === "canary-14-large-right-panel")!;
  const rightCopy = deck.scene!.objects.find((object) => object.name === "canary-14-right-copy")!;
  const solved = solveSceneToLayout({
    deck,
    slideNumber: 14,
    rationale: "Fit the complete comparison scene into one shared two-region layout recipe.",
    regions: [
      { id: "left-content", groups: [[leftPanel.id, leftLabel.id]], groupRoles: ["supporting"], regionPt: { left: 36, top: 100, width: 210, height: 300 }, mode: "vertical-stack", alignment: "center" },
      { id: "right-content", groups: [[rightPanel.id, rightCopy.id]], groupRoles: ["primary"], regionPt: { left: 270, top: 100, width: 650, height: 300 }, mode: "vertical-stack", alignment: "center" },
    ],
  });
  assert.equal(solved.status, "solved");
  assert.equal(solved.operation, "fit-scene-to-layout");
  assert.ok(solved.commands.length >= 2);
  const commandById = new Map(solved.commands.map((command) => [command.objectId, command]));
  const targetBox = (object: typeof leftPanel) => {
    const command = commandById.get(object.id);
    return {
      left: (command?.target.x ?? object.geometry.x) / 12_700,
      top: (command?.target.y ?? object.geometry.y) / 12_700,
      width: object.geometry.width / 12_700,
      height: object.geometry.height / 12_700,
    };
  };
  const left = targetBox(leftPanel);
  const right = targetBox(rightPanel);
  assert.ok(left.left >= 36 && left.left + left.width <= 246);
  assert.ok(right.left >= 270 && right.left + right.width <= 920);
  assert.ok(left.left + left.width < right.left);
  const leftPanelDelta = (commandById.get(leftPanel.id)?.target.x ?? leftPanel.geometry.x) - leftPanel.geometry.x;
  const leftLabelDelta = (commandById.get(leftLabel.id)?.target.x ?? leftLabel.geometry.x) - leftLabel.geometry.x;
  assert.equal(leftPanelDelta, leftLabelDelta);
});
