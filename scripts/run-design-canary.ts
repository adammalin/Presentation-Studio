import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { createDesignCanaryDeck } from "./create-design-canary";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { bindNativeMeasurement, compareNativeMeasurementPackets, type NativeMeasurementPacket } from "../src/lib/native-measurement";
import { calculateDesignMetrics } from "../src/lib/design-metrics";
import { solveAlignment, solveDistribution, solveSafeRegion } from "../src/lib/layout-solver";
import { recommendedTableGrowthPlan, solveTableLayout, type TableSolverResult } from "../src/lib/table-layout-solver";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createTableLayoutProposal } from "../src/lib/cleanup";
import { PRESENTATION_DESIGN_STANDARD } from "../src/lib/design-standard";
import { sha256 } from "../src/lib/hash";
import type { DeckJob, GeometryEditCommand, PresentationSceneObject } from "../src/types";
import type { NativeMeasurementResult, NativeRenderResult } from "../src/lib/desktop";

const require = createRequire(import.meta.url);
const { measurePowerPointNative } = require("../electron/native-measurement.cjs") as { measurePowerPointNative(input: { bytes: Uint8Array; name: string }): Promise<NativeMeasurementResult> };
const { renderPowerPointNative } = require("../electron/native-render.cjs") as { renderPowerPointNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }): Promise<NativeRenderResult> };
const SAFE_MARGIN_PT = PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt;

function objectByName(deck: DeckJob, name: string) {
  const object = deck.scene?.objects.find((candidate) => candidate.name === name);
  if (!object) throw new Error(`Canary object ${name} was not found by stable PowerPoint name.`);
  return object;
}

function opticalLeft(packet: NativeMeasurementPacket, objectId: string) {
  const value = packet.objects.find((object) => object.objectId === objectId)?.text?.renderedBoundsPt?.left;
  if (value === undefined) throw new Error(`Native rendered-text bounds are missing for ${objectId}.`);
  return value;
}

function measuredBox(packet: NativeMeasurementPacket, object: PresentationSceneObject) {
  const box = packet.objects.find((candidate) => candidate.objectId === object.id)?.measuredGeometryPt;
  if (!box) throw new Error(`Native object geometry is missing for ${object.id}.`);
  return box;
}

function unionMeasured(packet: NativeMeasurementPacket, objects: PresentationSceneObject[]) {
  const boxes = objects.map((object) => measuredBox(packet, object));
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

function verticalGroupGaps(packet: NativeMeasurementPacket, groups: PresentationSceneObject[][]) {
  const boxes = groups.map((group) => unionMeasured(packet, group)).sort((left, right) => left.top - right.top);
  return boxes.slice(1).map((box, index) => box.top - boxes[index].top - boxes[index].height);
}

function tableClearance(packet: NativeMeasurementPacket, tableId: string) {
  const cells = packet.objects.find((object) => object.tableId === tableId)?.table?.cells ?? [];
  const horizontal = cells.flatMap((cell) => cell.clearancesPt ? [cell.clearancesPt.left, cell.clearancesPt.right] : []);
  const vertical = cells.flatMap((cell) => cell.clearancesPt ? [cell.clearancesPt.top, cell.clearancesPt.bottom] : []);
  return {
    measuredCellCount: cells.length,
    minimumHorizontalPt: horizontal.length ? Math.min(...horizontal) : undefined,
    minimumVerticalPt: vertical.length ? Math.min(...vertical) : undefined,
  };
}

function tableWrap(packet: NativeMeasurementPacket, tableId: string) {
  const cells = packet.objects.find((object) => object.tableId === tableId)?.table?.cells ?? [];
  const lineCounts = cells.filter((cell) => cell.textLength > 0).map((cell) => cell.lineCount);
  return { maximumLineCount: lineCounts.length ? Math.max(...lineCounts) : 0, cellsOverTwoLines: lineCounts.filter((count) => count > 2).length };
}

function tableBounds(packet: NativeMeasurementPacket, tableId: string) {
  const box = packet.objects.find((object) => object.tableId === tableId)?.measuredGeometryPt;
  if (!box) throw new Error(`Native table bounds are missing for ${tableId}.`);
  return box;
}

function tableConstraints(variant: "standard" | "dense-technical") {
  const settings = variant === "dense-technical" ? PRESENTATION_DESIGN_STANDARD.tableVariants.denseTechnical : PRESENTATION_DESIGN_STANDARD.tableVariants.standard;
  return { minimumFontPt: settings.bodyFontSizePt, minimumHorizontalPaddingPt: settings.horizontalPaddingPt, minimumVerticalPaddingPt: settings.verticalPaddingPt };
}

function tableClearancePasses(packet: NativeMeasurementPacket, tableId: string, constraints: ReturnType<typeof tableConstraints>) {
  const clearance = tableClearance(packet, tableId);
  return clearance.measuredCellCount > 0
    && (clearance.minimumHorizontalPt ?? -Infinity) >= constraints.minimumHorizontalPaddingPt - .5
    && (clearance.minimumVerticalPt ?? -Infinity) >= constraints.minimumVerticalPaddingPt - .5;
}

function bindingCounts(packet: NativeMeasurementPacket) {
  return packet.objects.reduce<Record<string, number>>((counts, object) => {
    counts[object.binding.method] = (counts[object.binding.method] ?? 0) + 1;
    return counts;
  }, {});
}

function bindRequiredNativeMeasurement(deck: DeckJob, result: NativeMeasurementResult, label: string) {
  if (result.status !== "ready" || result.authority !== "powerpoint-native") {
    throw new Error(`${label} requires PowerPoint-native measurement${result.reason ? ` (${result.reason})` : ""}: ${result.warnings.join(" ") || "PowerPoint returned no native measurements."}`);
  }
  return bindNativeMeasurement(deck, result);
}

async function writeRenderSlides(render: NativeRenderResult, slideNumbers: number[], outputRoot: string, label: string) {
  if (render.status !== "ready") throw new Error(`${label} native render is not ready: ${render.reason ?? render.warnings.join(" ")}`);
  for (const slideNumber of slideNumbers) {
    const slide = render.slides.find((item) => item.number === slideNumber);
    if (!slide) throw new Error(`Native render did not return slide ${slideNumber}.`);
    await fs.writeFile(path.join(outputRoot, `slide-${String(slideNumber).padStart(2, "0")}-${label}.png`), Buffer.from(slide.bytes));
  }
}

function exactContentPreserved(source: Awaited<ReturnType<typeof auditPptx>>, candidate: Awaited<ReturnType<typeof auditPptx>>) {
  return source.slideCount === candidate.slideCount
    && source.slides.every((slide, index) => slide.textHash === candidate.slides[index]?.textHash)
    && source.tables.every((table) => candidate.tables.find((item) => item.id === table.id)?.contentHash === table.contentHash && candidate.tables.find((item) => item.id === table.id)?.structureHash === table.structureHash);
}

export async function runDesignCanary(root: string) {
  const generated = path.join(root, "fixtures", "generated", "precision-layout-canary.pptx");
  const groundTruthPath = path.join(root, "fixtures", "design-canary-ground-truth.json");
  const outputRoot = path.join(root, "tmp", "design-canary");
  await fs.mkdir(outputRoot, { recursive: true });
  const groundTruthBytes = new Uint8Array(await fs.readFile(groundTruthPath));
  const groundTruth = JSON.parse(Buffer.from(groundTruthBytes).toString("utf8")) as { schema?: string; version?: number; slides?: Array<{ slideNumber?: number; defect?: string; objects?: string[] }> };
  if (groundTruth.schema !== "presentation-studio/design-canary-ground-truth" || groundTruth.version !== 2 || groundTruth.slides?.length !== 14 || new Set(groundTruth.slides.map((slide) => slide.slideNumber)).size !== 14) throw new Error("The version-2 design-canary ground truth is missing or incomplete.");
  await createDesignCanaryDeck(generated);
  const sourceBytes = new Uint8Array(await fs.readFile(generated));
  const sourceSha256 = await sha256(sourceBytes);
  const audit = await auditPptx(sourceBytes);
  const now = new Date().toISOString();
  const deck: DeckJob = {
    id: "precision-layout-canary",
    name: "precision-layout-canary.pptx",
    sourceResourceId: "precision-layout-canary-source",
    sourceSha256,
    operationScope: "reflow",
    templateClassification: "custom",
    targetTemplateId: "synthetic-canary",
    targetTemplateConfirmedAt: now,
    targetTemplateDecisionSource: "user-selected",
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
  deck.scene = compilePresentationScene({ ...deck, audit });

  const currentNative = await measurePowerPointNative({ bytes: sourceBytes, name: deck.name });
  const currentMeasurement = bindRequiredNativeMeasurement(deck, currentNative, "The source canary baseline");
  const currentMetrics = calculateDesignMetrics(deck, currentMeasurement);

  const alignmentSpecs = [
    { slideNumber: 1, outlierName: "canary-1-title-off-2pt", anchorName: "canary-1-body-reference", expectedPt: 2 },
    { slideNumber: 2, outlierName: "canary-2-title-off-4pt", anchorName: "canary-2-body-reference", expectedPt: 4 },
    { slideNumber: 3, outlierName: "canary-3-text-margin-5_4pt", anchorName: "canary-3-text-zero-margin", expectedPt: 5.4 },
    { slideNumber: 4, outlierName: "canary-4-bullet-indent", anchorName: "canary-4-plain", expectedPt: 18 },
  ];
  const alignmentResults = alignmentSpecs.map((spec) => {
    const outlier = objectByName(deck, spec.outlierName);
    const anchor = objectByName(deck, spec.anchorName);
    const detectedMagnitudePt = Math.abs(opticalLeft(currentMeasurement, outlier.id) - opticalLeft(currentMeasurement, anchor.id));
    const solved = solveAlignment({ deck, measurement: currentMeasurement, slideNumber: spec.slideNumber, objectIds: [outlier.id, anchor.id], anchorObjectId: anchor.id, mode: "optical-left", rationale: `Repair the slide ${spec.slideNumber} optical-left canary while preserving the intended reference grid.` });
    if (solved.status !== "solved" || !solved.commands.length) throw new Error(`The slide ${spec.slideNumber} optical solver failed: ${solved.diagnostics.join(" ")}`);
    return { ...spec, outlier, anchor, detectedMagnitudePt, solved };
  });

  const distributionGroups = [1, 2, 3, 4].map((index) => [objectByName(deck, `canary-5-panel-${index}`), objectByName(deck, `canary-5-body-${index}`)]);
  const currentVerticalGapsPt = verticalGroupGaps(currentMeasurement, distributionGroups);
  const distribution = solveDistribution({ deck, slideNumber: 5, objectIds: distributionGroups.flat().map((object) => object.id), groups: distributionGroups.map((group) => group.map((object) => object.id)), mode: "vertical-equal-gap", rationale: "Normalize the repeated evidence blocks as four intact visual groups." });
  if (distribution.status !== "solved" || !distribution.commands.length) throw new Error(`The grouped distribution solver failed: ${distribution.diagnostics.join(" ")}`);

  const safeGroup = [objectByName(deck, "canary-6-safe-margin-minus-4pt"), objectByName(deck, "canary-6-safe-margin-copy")];
  const currentSafeLeftPt = unionMeasured(currentMeasurement, safeGroup).left;
  const safeRegion = solveSafeRegion({ deck, slideNumber: 6, objectIds: safeGroup.map((object) => object.id), rationale: "Move the complete callout to the 18-point safe region with minimum movement." });
  if (safeRegion.status !== "solved" || !safeRegion.commands.length) throw new Error(`The safe-region solver failed: ${safeRegion.diagnostics.join(" ")}`);

  const geometryCommands: GeometryEditCommand[] = [...alignmentResults.flatMap((result) => result.solved.commands), ...distribution.commands, ...safeRegion.commands];
  let proposal = createGeometryBatchProposal(deck, now, geometryCommands.map((command) => ({ objectId: command.objectId, target: command.target, rationale: command.rationale, author: "ai", constraints: command.constraints })));

  const tableSpecs: Array<{ slideNumber: number; variant: "standard" | "dense-technical"; expectation: string }> = [
    { slideNumber: 7, variant: "standard", expectation: "repair-two-point-padding" },
    { slideNumber: 8, variant: "standard", expectation: "repair-four-point-horizontal-padding" },
    { slideNumber: 9, variant: "standard", expectation: "preserve-compliant-six-point-control" },
    { slideNumber: 10, variant: "standard", expectation: "grow-clipped-rows-and-fit" },
    { slideNumber: 11, variant: "dense-technical", expectation: "redistribute-wrap-pressured-columns" },
    { slideNumber: 12, variant: "standard", expectation: "refit-native-auto-grown-rows" },
    { slideNumber: 13, variant: "dense-technical", expectation: "fit-with-merged-topology-preserved" },
  ];
  const tableCanaries: Array<{
    slideNumber: number;
    tableId: string;
    variant: "standard" | "dense-technical";
    expectation: string;
    initialStatus: TableSolverResult["status"];
    finalStatus: TableSolverResult["status"];
    mutation: "preserved-as-is" | "table-layout" | "grow-and-table-layout" | "grow-only";
    growthPlan?: ReturnType<typeof recommendedTableGrowthPlan>;
    iterations: Array<{ iteration: number; clearance: ReturnType<typeof tableClearance>; wrap: ReturnType<typeof tableWrap>; passed: boolean }>;
    currentClearance: ReturnType<typeof tableClearance>;
    currentWrap: ReturnType<typeof tableWrap>;
    currentBounds: ReturnType<typeof tableBounds>;
    proposalClearance?: ReturnType<typeof tableClearance>;
    proposalWrap?: ReturnType<typeof tableWrap>;
    proposalBounds?: ReturnType<typeof tableBounds>;
    constraints: ReturnType<typeof tableConstraints>;
  }> = [];
  let tableLayoutCommandCount = 0;
  for (const spec of tableSpecs) {
    const inventory = audit.tables.find((table) => table.slideNumber === spec.slideNumber);
    if (!inventory) throw new Error(`The slide ${spec.slideNumber} native table canary is missing.`);
    const object = deck.scene.objects.find((item) => item.sourceLocator.tableId === inventory.id);
    if (!object) throw new Error(`The slide ${spec.slideNumber} table object is not source-bound.`);
    const constraints = tableConstraints(spec.variant);
    let result = solveTableLayout({ deck, measurement: currentMeasurement, tableId: inventory.id, rationale: `Qualify ${spec.expectation} while preserving exact native table content and structure.`, variant: spec.variant });
    const initialStatus = result.status;
    const growthPlan = recommendedTableGrowthPlan(result, object.id, `Grow slide ${spec.slideNumber}'s native table by the minimum measured amount before fitting its cell grid.`);
    let mutation: "preserved-as-is" | "table-layout" | "grow-and-table-layout" | "grow-only" = result.status === "already-fit" ? "preserved-as-is" : "table-layout";
    let workingMeasurement = currentMeasurement;
    if (growthPlan) {
      proposal = createGeometryBatchProposal({ ...deck, proposal }, now, [{ objectId: growthPlan.objectId, target: growthPlan.target, rationale: growthPlan.rationale, author: "ai", constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false } }]);
      result = solveTableLayout({
        deck,
        measurement: currentMeasurement,
        tableId: inventory.id,
        rationale: `Grow and fit slide ${spec.slideNumber}'s native table atomically with the resolved ${spec.variant} constraints.`,
        variant: spec.variant,
        targetBoundsPt: { width: growthPlan.target.width / 12_700, height: growthPlan.target.height / 12_700 },
      });
      mutation = "grow-and-table-layout";
    }
    if (result.status === "infeasible") throw new Error(`The slide ${spec.slideNumber} table canary remained infeasible: ${result.diagnostics.reasons.join(" ")} ${result.diagnostics.recommendations.join(" ")}`);
    if (result.status === "solved" && result.command) {
      proposal = createTableLayoutProposal({ ...deck, proposal }, now, result.command);
      tableLayoutCommandCount += 1;
    }
    const iterations: Array<{ iteration: number; clearance: ReturnType<typeof tableClearance>; wrap: ReturnType<typeof tableWrap>; passed: boolean }> = [];
    let passed = result.status === "already-fit" && tableClearancePasses(workingMeasurement, inventory.id, constraints);
    for (let iteration = 1; !passed && iteration <= 3; iteration += 1) {
      const iterationCandidate = await buildCleanupProposalPptx(sourceBytes, proposal);
      await fs.writeFile(path.join(outputRoot, `slide-${String(spec.slideNumber).padStart(2, "0")}-iteration-${iteration}.pptx`), Buffer.from(iterationCandidate.bytes));
      const iterationNative = await measurePowerPointNative({ bytes: iterationCandidate.bytes, name: `precision-layout-canary-slide-${spec.slideNumber}-iteration-${iteration}.pptx` });
      workingMeasurement = bindRequiredNativeMeasurement(deck, iterationNative, `The slide ${spec.slideNumber} iteration ${iteration} candidate`);
      const clearance = tableClearance(workingMeasurement, inventory.id);
      const wrap = tableWrap(workingMeasurement, inventory.id);
      passed = tableClearancePasses(workingMeasurement, inventory.id, constraints);
      iterations.push({ iteration, clearance, wrap, passed });
      if (passed) break;
      result = solveTableLayout({ deck, measurement: workingMeasurement, tableId: inventory.id, rationale: `Refine slide ${spec.slideNumber}'s table from PowerPoint-native iteration ${iteration}.`, variant: spec.variant });
      if (result.status === "infeasible" || !result.command) throw new Error(`The slide ${spec.slideNumber} table canary could not satisfy native clearances after iteration ${iteration}: ${result.diagnostics.reasons.join(" ")} ${result.diagnostics.recommendations.join(" ")}`);
      proposal = createTableLayoutProposal({ ...deck, proposal }, now, result.command);
    }
    if (!passed) throw new Error(`The slide ${spec.slideNumber} table canary exceeded the three-round native-clearance limit.`);
    tableCanaries.push({ slideNumber: spec.slideNumber, tableId: inventory.id, variant: spec.variant, expectation: spec.expectation, initialStatus, finalStatus: result.status, mutation, growthPlan, iterations, currentClearance: tableClearance(currentMeasurement, inventory.id), currentWrap: tableWrap(currentMeasurement, inventory.id), currentBounds: tableBounds(currentMeasurement, inventory.id), constraints });
  }

  const materialized = await buildCleanupProposalPptx(sourceBytes, proposal);
  const candidatePath = path.join(outputRoot, "precision-layout-canary-proposal.pptx");
  await fs.writeFile(candidatePath, Buffer.from(materialized.bytes));
  const proposalAudit = await auditPptx(materialized.bytes);
  const proposalNative = await measurePowerPointNative({ bytes: materialized.bytes, name: path.basename(candidatePath) });
  const proposalMeasurement = bindRequiredNativeMeasurement(deck, proposalNative, "The final proposal candidate");
  const proposalMetrics = calculateDesignMetrics(deck, proposalMeasurement, currentMeasurement);

  const alignmentEvidence = alignmentResults.map((result) => ({
    slideNumber: result.slideNumber,
    expectedMagnitudePt: result.expectedPt,
    detectedMagnitudePt: result.detectedMagnitudePt,
    repairedMagnitudePt: Math.abs(opticalLeft(proposalMeasurement, result.outlier.id) - opticalLeft(proposalMeasurement, result.anchor.id)),
    commandCount: result.solved.commands.length,
    movementCostPt: result.solved.objective.movementCostPt,
  }));
  const proposalVerticalGapsPt = verticalGroupGaps(proposalMeasurement, distributionGroups);
  const proposalSafeLeftPt = unionMeasured(proposalMeasurement, safeGroup).left;
  for (const canary of tableCanaries) {
    canary.proposalClearance = tableClearance(proposalMeasurement, canary.tableId);
    canary.proposalWrap = tableWrap(proposalMeasurement, canary.tableId);
    canary.proposalBounds = tableBounds(proposalMeasurement, canary.tableId);
  }

  const exportedBytes = new Uint8Array(await fs.readFile(candidatePath));
  const exportedAudit = await auditPptx(exportedBytes);
  const exportedNative = await measurePowerPointNative({ bytes: exportedBytes, name: "precision-layout-canary-exported.pptx" });
  const exportedMeasurement = bindRequiredNativeMeasurement(deck, exportedNative, "The reopened export candidate");
  const measurementComparison = compareNativeMeasurementPackets(proposalMeasurement, exportedMeasurement, .2);

  const renderSlides = Array.from({ length: 14 }, (_value, index) => index + 1);
  const currentRender = await renderPowerPointNative({ bytes: sourceBytes, name: deck.name, width: 2200, format: "png" });
  const proposalRender = await renderPowerPointNative({ bytes: materialized.bytes, name: path.basename(candidatePath), width: 2200, format: "png" });
  const exportedRender = await renderPowerPointNative({ bytes: exportedBytes, name: "precision-layout-canary-exported.pptx", width: 2200, format: "png" });
  await writeRenderSlides(currentRender, renderSlides, outputRoot, "current");
  await writeRenderSlides(proposalRender, renderSlides, outputRoot, "proposal");
  await writeRenderSlides(exportedRender, renderSlides, outputRoot, "export");

  const proposalRenderHashes = proposalRender.status === "ready" ? new Map(proposalRender.slides.map((slide) => [slide.number, slide.sha256])) : new Map<number, string>();
  const exportedRenderHashes = exportedRender.status === "ready" ? new Map(exportedRender.slides.map((slide) => [slide.number, slide.sha256])) : new Map<number, string>();
  const binding = { current: bindingCounts(currentMeasurement), proposal: bindingCounts(proposalMeasurement), export: bindingCounts(exportedMeasurement) };
  const stableBindingCount = (counts: Record<string, number>) => (counts["shape-id"] ?? 0) + (counts.name ?? 0);
  const sourceExact = exactContentPreserved(audit, proposalAudit);
  const exportExact = exactContentPreserved(audit, exportedAudit);
  const slide13Source = audit.tables.find((table) => table.slideNumber === 13);
  const slide13Export = exportedAudit.tables.find((table) => table.slideNumber === 13);
  const checks = {
    groundTruthCoversEverySlide: groundTruth.slides.every((item, index) => item.slideNumber === index + 1 && Boolean(item.defect) && (item.objects?.length ?? 0) > 0),
    nativeMeasurementReady: [currentMeasurement, proposalMeasurement, exportedMeasurement].every((packet) => packet.authority === "powerpoint-native"),
    stablePowerPointIdentityBinding: [currentMeasurement, proposalMeasurement, exportedMeasurement].every((packet) => stableBindingCount(bindingCounts(packet)) === packet.objects.length),
    allOpticalDefectsDetected: alignmentEvidence.every((item) => Math.abs(item.detectedMagnitudePt - item.expectedMagnitudePt) <= 1),
    allOpticalDefectsRepaired: alignmentEvidence.every((item) => item.repairedMagnitudePt <= .5),
    groupedVerticalRhythmDetected: Math.max(...currentVerticalGapsPt) - Math.min(...currentVerticalGapsPt) >= 7,
    groupedVerticalRhythmRepaired: Math.max(...proposalVerticalGapsPt) - Math.min(...proposalVerticalGapsPt) <= .5,
    safeRegionDefectDetected: Math.abs(SAFE_MARGIN_PT - currentSafeLeftPt - 4) <= .5,
    safeRegionRepaired: proposalSafeLeftPt >= SAFE_MARGIN_PT - .2,
    nativeTableCellMeasurementReady: tableCanaries.every((canary) => canary.currentClearance.measuredCellCount > 0 && (canary.proposalClearance?.measuredCellCount ?? 0) > 0),
    nativeTableSolverRepairedClearance: tableCanaries.every((canary) => tableClearancePasses(proposalMeasurement, canary.tableId, canary.constraints)),
    compliantTablePreservedAsIs: tableCanaries.find((canary) => canary.slideNumber === 9)?.mutation === "preserved-as-is",
    clippedTableGrewAndNativeAutoRowsRefit: (() => {
      const clipped = tableCanaries.find((item) => item.slideNumber === 10);
      const autoGrown = tableCanaries.find((item) => item.slideNumber === 12);
      return Boolean(clipped?.growthPlan
        && (clipped.proposalBounds?.height ?? 0) > clipped.currentBounds.height + .2
        && autoGrown?.mutation === "table-layout"
        && !autoGrown.growthPlan
        && tableClearancePasses(proposalMeasurement, autoGrown.tableId, autoGrown.constraints));
    })(),
    wrapPressureDidNotRegress: (tableCanaries.find((canary) => canary.slideNumber === 11)?.proposalWrap?.maximumLineCount ?? Infinity) <= (tableCanaries.find((canary) => canary.slideNumber === 11)?.currentWrap.maximumLineCount ?? -Infinity),
    exactContentAndTableStructurePreserved: sourceExact && exportExact,
    mergedTableTopologyPreserved: Boolean(slide13Source && slide13Export && slide13Source.structureHash === slide13Export.structureHash && slide13Source.mergedCellCount === slide13Export.mergedCellCount),
    nativeRenderBaselineReady: [currentRender, proposalRender, exportedRender].every((render) => render.status === "ready"),
    exportedMeasurementMatchesProposal: measurementComparison.equivalent,
    exportedRenderMatchesProposal: renderSlides.every((slideNumber) => proposalRenderHashes.get(slideNumber) === exportedRenderHashes.get(slideNumber)),
    aiOnlyCanaryReservedForVisualJudgment: objectByName(deck, "canary-14-small-left-panel").slideNumber === 14 && objectByName(deck, "canary-14-large-right-panel").slideNumber === 14,
  };
  const report = {
    schema: "presentation-studio/design-canary-report",
    version: 2,
    generatedAt: new Date().toISOString(),
    source: { file: generated, sha256: sourceSha256, slideCount: audit.slideCount },
    groundTruth: { file: groundTruthPath, sha256: await sha256(groundTruthBytes), version: groundTruth.version, slideCount: groundTruth.slides.length },
    candidate: { file: candidatePath, sha256: await sha256(exportedBytes), geometryCommandCount: geometryCommands.length + tableCanaries.filter((canary) => canary.growthPlan).length, tableLayoutCommandCount },
    native: { adapter: currentMeasurement.adapter, powerPointVersion: currentMeasurement.powerPointVersion, binding },
    alignmentCanaries: alignmentEvidence,
    distributionCanary: { slideNumber: 5, currentGapsPt: currentVerticalGapsPt, proposalGapsPt: proposalVerticalGapsPt, commandCount: distribution.commands.length },
    safeRegionCanary: { slideNumber: 6, requiredMarginPt: SAFE_MARGIN_PT, currentLeftPt: currentSafeLeftPt, proposalLeftPt: proposalSafeLeftPt, commandCount: safeRegion.commands.length },
    tableCanaries,
    visualJudgmentCanary: { slideNumber: 14, deterministicMutationCount: 0, status: "requires-ai-judgment", instruction: "Use the native pixels to judge balance and choose semantic recomposition; no precision solver should invent this design decision." },
    metrics: { current: currentMetrics.slides.filter((slide) => slide.slideNumber <= 13), proposal: proposalMetrics.slides.filter((slide) => slide.slideNumber <= 13) },
    exportAcceptance: { exactContent: exportExact, measurementComparison, renderHashesMatch: checks.exportedRenderMatchesProposal },
    checks,
    passed: Object.values(checks).every(Boolean),
  };
  await fs.writeFile(path.join(outputRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`);
  return report;
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
export async function runDesignCanaryCli(
  root: string,
  options: {
    run?: typeof runDesignCanary;
    stdout?: (message: string) => void;
    stderr?: (message: string) => void;
  } = {},
) {
  const run = options.run ?? runDesignCanary;
  const stdout = options.stdout ?? console.log;
  const stderr = options.stderr ?? console.error;
  try {
    const report = await run(root);
    stdout(JSON.stringify(report, null, 2));
    return report.passed ? 0 : 1;
  } catch (error) {
    const failure = { schema: "presentation-studio/design-canary-report", version: 2, generatedAt: new Date().toISOString(), passed: false, checks: { canaryCompleted: false }, error: error instanceof Error ? error.message : String(error) };
    const outputRoot = path.join(root, "tmp", "design-canary");
    await fs.mkdir(outputRoot, { recursive: true });
    await fs.writeFile(path.join(outputRoot, "report.json"), `${JSON.stringify(failure, null, 2)}\n`);
    stderr(JSON.stringify(failure, null, 2));
    return 1;
  }
}

if (invokedPath === fileURLToPath(import.meta.url)) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  process.exitCode = await runDesignCanaryCli(root);
}
