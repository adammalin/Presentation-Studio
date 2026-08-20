import type { DeckJob } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { calculateNativeTextOverflowEdges, type BoundObjectMeasurement, type NativeMeasurementPacket } from "./native-measurement";

const EMU_PER_POINT = 12_700;
const SAFE_MARGIN_PT = PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt;
const GEOMETRY_TOLERANCE_PT = 0.75;

export const DESIGN_METRICS_SCHEMA = "presentation-studio/design-metrics" as const;

export interface SlideDesignMetrics {
  slideNumber: number;
  authority: NativeMeasurementPacket["authority"];
  opticalLeftMadPt?: number;
  opticalLeftRangePt?: number;
  verticalGapVariancePt2?: number;
  verticalGapMadPt?: number;
  safeRegionViolationCount: number;
  offSlideObjectCount: number;
  offSlideObjectIds: string[];
  textOverflowCount: number;
  textOverflowObjectIds: string[];
  minimumTableCellClearancePt?: number;
  minimumTableHorizontalClearancePt?: number;
  minimumTableVerticalClearancePt?: number;
  tableCellClearanceViolationCount: number;
  tableColumnImbalanceRatio?: number;
  tableRowHeightVariancePt2?: number;
  movementCostPt?: number;
  changedObjectCount?: number;
  tableCellFindings: Array<{
    tableId: string;
    cellId: string;
    row: number;
    column: number;
    rule: "insufficient-clearance" | "wrap-pressure" | "native-overflow";
    severity: "warning" | "error";
    evidence: string;
  }>;
  warnings: string[];
}

export interface DesignMetricsReport {
  schema: typeof DESIGN_METRICS_SCHEMA;
  version: 2;
  measurementRevision: string;
  sourceSha256: string;
  slides: SlideDesignMetrics[];
  totals: {
    safeRegionViolationCount: number;
    offSlideObjectCount: number;
    textOverflowCount: number;
    tableCellClearanceViolationCount: number;
    movementCostPt: number;
    changedObjectCount: number;
  };
}

function rounded(value: number, places = 3) {
  const factor = 10 ** places;
  return Math.round(value * factor) / factor;
}

function mean(values: number[]) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : undefined;
}

function median(values: number[]) {
  if (!values.length) return undefined;
  const sorted = [...values].sort((left, right) => left - right);
  const midpoint = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[midpoint] : (sorted[midpoint - 1] + sorted[midpoint]) / 2;
}

function variance(values: number[]) {
  const average = mean(values);
  return average === undefined ? undefined : values.reduce((sum, value) => sum + (value - average) ** 2, 0) / values.length;
}

function mad(values: number[]) {
  const center = median(values);
  return center === undefined ? undefined : median(values.map((value) => Math.abs(value - center)));
}

function validTextObjects(objects: BoundObjectMeasurement[]) {
  return objects.filter((object) => object.text?.renderedBoundsPt && object.text.textLength > 0);
}

function verticalGaps(objects: BoundObjectMeasurement[]) {
  const sorted = objects.filter((object) => object.measuredGeometryPt).sort((left, right) => left.measuredGeometryPt!.top - right.measuredGeometryPt!.top);
  const gaps: number[] = [];
  for (let index = 1; index < sorted.length; index += 1) {
    const previous = sorted[index - 1].measuredGeometryPt!;
    const current = sorted[index].measuredGeometryPt!;
    const gap = current.top - (previous.top + previous.height);
    if (gap >= -GEOMETRY_TOLERANCE_PT) gaps.push(Math.max(0, gap));
  }
  return gaps;
}

function tableMetrics(deck: DeckJob, objects: BoundObjectMeasurement[]) {
  const cellClearances = objects.flatMap((object) => object.table?.cells.flatMap((cell) => cell.clearancesPt ? [{ horizontal: Math.min(cell.clearancesPt.left, cell.clearancesPt.right), vertical: Math.min(cell.clearancesPt.top, cell.clearancesPt.bottom) }] : []) ?? []);
  const clearances = cellClearances.flatMap((clearance) => [clearance.horizontal, clearance.vertical]);
  const tables = objects.filter((object) => object.table);
  const columnRatios = tables.map((object) => {
    const widths = object.table!.columnWidthsPt.filter((value) => value > 0);
    return widths.length ? Math.max(...widths) / Math.max(0.001, Math.min(...widths)) : undefined;
  }).filter((value): value is number => value !== undefined);
  const rowVariances = tables.map((object) => variance(object.table!.rowHeightsPt.filter((value) => value > 0))).filter((value): value is number => value !== undefined);
  const findings = tables.flatMap((object) => {
    const auditedTable = deck.audit?.tables.find((table) => table.id === object.tableId);
    const extremeDense = Boolean(auditedTable && (auditedTable.totalCellCharacterCount >= 2_200 || auditedTable.maximumCellCharacterCount >= 700));
    const horizontalFloorPt = extremeDense ? 4 : PRESENTATION_DESIGN_STANDARD.tableVariants.standard.horizontalPaddingPt;
    const verticalFloorPt = extremeDense ? 2 : PRESENTATION_DESIGN_STANDARD.tableVariants.standard.verticalPaddingPt;
    return object.table!.cells.flatMap((cell) => {
    const result: SlideDesignMetrics["tableCellFindings"] = [];
    if (cell.clearancesPt) {
      const horizontal = Math.min(cell.clearancesPt.left, cell.clearancesPt.right);
      const vertical = Math.min(cell.clearancesPt.top, cell.clearancesPt.bottom);
      if (horizontal < horizontalFloorPt - .5 || vertical < verticalFloorPt - .5) result.push({ tableId: object.tableId ?? object.objectId, cellId: cell.cellId, row: cell.row, column: cell.column, rule: "insufficient-clearance", severity: horizontal < 1 || vertical < 1 ? "error" : "warning", evidence: `PowerPoint-native clearance is ${horizontal.toFixed(2)} pt horizontal and ${vertical.toFixed(2)} pt vertical; the resolved floors are ${horizontalFloorPt} pt and ${verticalFloorPt} pt.` });
    }
    if (cell.lineCount > 2) result.push({ tableId: object.tableId ?? object.objectId, cellId: cell.cellId, row: cell.row, column: cell.column, rule: "wrap-pressure", severity: !extremeDense && cell.lineCount > 4 ? "error" : "warning", evidence: `PowerPoint rendered ${cell.lineCount} lines in this cell${extremeDense ? "; this is an explicitly compact exact-content table" : ""}.` });
    if (cell.boundsPt && cell.renderedTextBoundsPt && cell.marginsPt && (cell.renderedTextBoundsPt.width > cell.boundsPt.width - cell.marginsPt.left - cell.marginsPt.right + .5 || cell.renderedTextBoundsPt.height > cell.boundsPt.height - cell.marginsPt.top - cell.marginsPt.bottom + .5)) result.push({ tableId: object.tableId ?? object.objectId, cellId: cell.cellId, row: cell.row, column: cell.column, rule: "native-overflow", severity: "error", evidence: "PowerPoint-native rendered text exceeds the measured inner cell frame." });
    return result;
    });
  });
  const clearanceViolations = findings.filter((finding) => finding.rule === "insufficient-clearance").length;
  return {
    minimumClearance: clearances.length ? Math.min(...clearances) : undefined,
    minimumHorizontalClearance: cellClearances.length ? Math.min(...cellClearances.map((item) => item.horizontal)) : undefined,
    minimumVerticalClearance: cellClearances.length ? Math.min(...cellClearances.map((item) => item.vertical)) : undefined,
    clearanceViolations,
    columnImbalance: mean(columnRatios),
    rowVariance: mean(rowVariances),
    findings,
  };
}

function textOverflowTolerance(deck: DeckJob, object: BoundObjectMeasurement) {
  const sourceObject = deck.audit?.editableObjects.find((candidate) => candidate.id === object.objectId);
  const textBox = sourceObject ? deck.audit?.textBoxes.find((candidate) => candidate.slideNumber === object.slideNumber && candidate.shapeId === sourceObject.shapeId) : undefined;
  // PowerPoint's native TextRange2 bounds include a small right-side glyph
  // overhang for bulleted paragraphs. A 3 pt allowance suppresses that known
  // measurement artifact while still reporting material wrap or clipping.
  return (textBox?.bulletParagraphCount ?? 0) > 0 ? 3 : GEOMETRY_TOLERANCE_PT;
}

function textOverflows(deck: DeckJob, object: BoundObjectMeasurement) {
  return calculateNativeTextOverflowEdges(object, textOverflowTolerance(deck, object)).length > 0;
}

function isIntentionalEdgeDecoration(deck: DeckJob, object: BoundObjectMeasurement, slideWidthPt: number, slideHeightPt: number) {
  const sceneObject = deck.scene?.objects.find((candidate) => candidate.id === object.objectId);
  const box = object.measuredGeometryPt;
  if (sceneObject?.semanticRole !== "decoration" || !box) return false;
  const touchesEdge = box.left <= GEOMETRY_TOLERANCE_PT
    || box.top <= GEOMETRY_TOLERANCE_PT
    || box.left + box.width >= slideWidthPt - GEOMETRY_TOLERANCE_PT
    || box.top + box.height >= slideHeightPt - GEOMETRY_TOLERANCE_PT;
  const formsEdgeBand = box.width >= slideWidthPt * .25 || box.height >= slideHeightPt * .25;
  return touchesEdge && formsEdgeBand;
}

function movement(current: BoundObjectMeasurement[], baseline?: NativeMeasurementPacket) {
  if (!baseline) return { cost: 0, count: 0 };
  let cost = 0;
  let count = 0;
  for (const object of current) {
    const before = baseline.objects.find((candidate) => candidate.objectId === object.objectId)?.measuredGeometryPt;
    const after = object.measuredGeometryPt;
    if (!before || !after) continue;
    const delta = Math.abs(after.left - before.left) + Math.abs(after.top - before.top) + Math.abs(after.width - before.width) + Math.abs(after.height - before.height);
    if (delta > GEOMETRY_TOLERANCE_PT) count += 1;
    cost += delta;
  }
  return { cost, count };
}

export function calculateSlideDesignMetrics(deck: DeckJob, packet: NativeMeasurementPacket, slideNumber: number, baseline?: NativeMeasurementPacket): SlideDesignMetrics {
  if (!deck.audit) throw new Error("A current audit is required before calculating design metrics.");
  const slideWidthPt = deck.audit.slideSize.width / EMU_PER_POINT;
  const slideHeightPt = deck.audit.slideSize.height / EMU_PER_POINT;
  const objects = packet.objects.filter((object) => object.slideNumber === slideNumber);
  const textObjects = validTextObjects(objects);
  const lefts = textObjects.map((object) => object.text!.renderedBoundsPt!.left);
  const gaps = verticalGaps(objects);
  const safeViolations = objects.filter((object) => {
    const box = object.measuredGeometryPt;
    return box && !isIntentionalEdgeDecoration(deck, object, slideWidthPt, slideHeightPt) && (box.left < SAFE_MARGIN_PT || box.top < SAFE_MARGIN_PT || box.left + box.width > slideWidthPt - SAFE_MARGIN_PT || box.top + box.height > slideHeightPt - SAFE_MARGIN_PT);
  });
  const offSlide = objects.filter((object) => {
    const box = object.measuredGeometryPt;
    return box && !isIntentionalEdgeDecoration(deck, object, slideWidthPt, slideHeightPt) && (box.left < -GEOMETRY_TOLERANCE_PT || box.top < -GEOMETRY_TOLERANCE_PT || box.left + box.width > slideWidthPt + GEOMETRY_TOLERANCE_PT || box.top + box.height > slideHeightPt + GEOMETRY_TOLERANCE_PT);
  });
  const tables = tableMetrics(deck, objects);
  const overflowingText = textObjects.filter((object) => textOverflows(deck, object));
  const move = movement(objects, baseline);
  const warnings: string[] = [];
  if (packet.authority !== "powerpoint-native") warnings.push("Rendered-text and cell-clearance metrics use OOXML fallback estimates; native PowerPoint measurement is required for final acceptance.");
  return {
    slideNumber,
    authority: packet.authority,
    opticalLeftMadPt: lefts.length > 1 ? rounded(mad(lefts)!) : undefined,
    opticalLeftRangePt: lefts.length > 1 ? rounded(Math.max(...lefts) - Math.min(...lefts)) : undefined,
    verticalGapVariancePt2: gaps.length > 1 ? rounded(variance(gaps)!) : undefined,
    verticalGapMadPt: gaps.length > 1 ? rounded(mad(gaps)!) : undefined,
    safeRegionViolationCount: safeViolations.length,
    offSlideObjectCount: offSlide.length,
    offSlideObjectIds: offSlide.map((object) => object.objectId),
    textOverflowCount: overflowingText.length,
    textOverflowObjectIds: overflowingText.map((object) => object.objectId),
    minimumTableCellClearancePt: tables.minimumClearance === undefined ? undefined : rounded(tables.minimumClearance),
    minimumTableHorizontalClearancePt: tables.minimumHorizontalClearance === undefined ? undefined : rounded(tables.minimumHorizontalClearance),
    minimumTableVerticalClearancePt: tables.minimumVerticalClearance === undefined ? undefined : rounded(tables.minimumVerticalClearance),
    tableCellClearanceViolationCount: tables.clearanceViolations,
    tableColumnImbalanceRatio: tables.columnImbalance === undefined ? undefined : rounded(tables.columnImbalance),
    tableRowHeightVariancePt2: tables.rowVariance === undefined ? undefined : rounded(tables.rowVariance),
    movementCostPt: rounded(move.cost),
    changedObjectCount: move.count,
    tableCellFindings: tables.findings,
    warnings,
  };
}

export function calculateDesignMetrics(deck: DeckJob, packet: NativeMeasurementPacket, baseline?: NativeMeasurementPacket): DesignMetricsReport {
  const slides = (deck.audit?.slides ?? []).map((slide) => calculateSlideDesignMetrics(deck, packet, slide.number, baseline));
  return {
    schema: DESIGN_METRICS_SCHEMA,
    version: 2,
    measurementRevision: packet.revision,
    sourceSha256: packet.sourceSha256,
    slides,
    totals: {
      safeRegionViolationCount: slides.reduce((sum, slide) => sum + slide.safeRegionViolationCount, 0),
      offSlideObjectCount: slides.reduce((sum, slide) => sum + slide.offSlideObjectCount, 0),
      textOverflowCount: slides.reduce((sum, slide) => sum + slide.textOverflowCount, 0),
      tableCellClearanceViolationCount: slides.reduce((sum, slide) => sum + slide.tableCellClearanceViolationCount, 0),
      movementCostPt: rounded(slides.reduce((sum, slide) => sum + (slide.movementCostPt ?? 0), 0)),
      changedObjectCount: slides.reduce((sum, slide) => sum + (slide.changedObjectCount ?? 0), 0),
    },
  };
}

export function metricsImproved(current: SlideDesignMetrics, proposal: SlideDesignMetrics) {
  const regressions: string[] = [];
  const improvements: string[] = [];
  const compareLower = (label: string, before?: number, after?: number) => {
    if (before === undefined || after === undefined) return;
    if (after < before - 0.01) improvements.push(label);
    else if (after > before + 0.01) regressions.push(label);
  };
  compareLower("text overflow", current.textOverflowCount, proposal.textOverflowCount);
  compareLower("off-slide geometry", current.offSlideObjectCount, proposal.offSlideObjectCount);
  compareLower("safe-region violations", current.safeRegionViolationCount, proposal.safeRegionViolationCount);
  compareLower("optical left-edge deviation", current.opticalLeftMadPt, proposal.opticalLeftMadPt);
  compareLower("vertical gap variance", current.verticalGapVariancePt2, proposal.verticalGapVariancePt2);
  compareLower("table clearance violations", current.tableCellClearanceViolationCount, proposal.tableCellClearanceViolationCount);
  return { acceptedByMetrics: regressions.length === 0 && improvements.length > 0, improvements, regressions };
}
