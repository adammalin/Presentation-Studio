import type { DeckJob, TableLayoutCommand } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import type { NativeMeasurementPacket } from "./native-measurement";

const EMU_PER_POINT = 12_700;

export interface TableSolverResult {
  status: "solved" | "already-fit" | "infeasible";
  command?: TableLayoutCommand;
  diagnostics: {
    currentWidthPt: number;
    currentHeightPt: number;
    minimumRequiredWidthPt: number;
    minimumRequiredHeightPt: number;
    recommendedBoundsPt?: { left: number; top: number; width: number; height: number };
    limitingCellIds: string[];
    reasons: string[];
    recommendations: string[];
  };
}

export interface TableGrowthPlan {
  objectId: string;
  target: { x: number; y: number; width: number; height: number };
  rationale: string;
}

function ptToEmu(value: number) { return Math.round(value * EMU_PER_POINT); }

function allocate(total: number, minimums: number[]) {
  const minimumTotal = minimums.reduce((sum, value) => sum + value, 0);
  if (minimumTotal > total + 0.01) return undefined;
  const surplus = total - minimumTotal;
  const weightTotal = minimums.reduce((sum, value) => sum + value, 0) || minimums.length;
  const allocated = minimums.map((value) => value + surplus * ((value || 1) / weightTotal));
  const correction = total - allocated.reduce((sum, value) => sum + value, 0);
  if (allocated.length) allocated[allocated.length - 1] += correction;
  return allocated;
}

function ensureSpanMinimum(minimums: number[], startIndex: number, span: number, required: number) {
  const safeSpan = Math.max(1, Math.min(span, minimums.length - startIndex));
  const current = minimums.slice(startIndex, startIndex + safeSpan).reduce((sum, value) => sum + value, 0);
  const deficit = required - current;
  if (deficit <= 0) return;
  for (let index = startIndex; index < startIndex + safeSpan; index += 1) minimums[index] += deficit / safeSpan;
}

export function solveTableLayout(input: {
  deck: DeckJob;
  measurement: NativeMeasurementPacket;
  tableId: string;
  rationale: string;
  variant?: "standard" | "dense-technical";
  targetBoundsPt?: { width: number; height: number };
}): TableSolverResult {
  const { deck, measurement, tableId } = input;
  if (!deck.audit || !deck.scene) throw new Error("A current audit and scene are required before solving a table layout.");
  const inventory = deck.audit.tables.find((table) => table.id === tableId);
  const tableScene = deck.scene.tables?.find((table) => table.id === tableId);
  const object = deck.scene.objects.find((candidate) => candidate.sourceLocator.tableId === tableId);
  const measuredObject = measurement.objects.find((candidate) => candidate.tableId === tableId);
  if (!inventory || !tableScene || !object || !measuredObject?.table || !measuredObject.measuredGeometryPt) throw new Error("The requested table does not have a complete cell-level scene and measurement binding.");
  if (!object.operations.resize || !object.operations.editTableStyle || object.protected) throw new Error("The requested table is protected or is not safely editable in the current scene.");
  const dense = input.variant === "dense-technical";
  const variant = dense ? PRESENTATION_DESIGN_STANDARD.tableVariants.denseTechnical : PRESENTATION_DESIGN_STANDARD.tableVariants.standard;
  const minimumFontPt = variant.bodyFontSizePt;
  const horizontalPaddingPt = variant.horizontalPaddingPt;
  const verticalPaddingPt = variant.verticalPaddingPt;
  const measuredWidthPt = measuredObject.measuredGeometryPt.width;
  const measuredHeightPt = measuredObject.measuredGeometryPt.height;
  const widthPt = input.targetBoundsPt?.width ?? measuredWidthPt;
  const heightPt = input.targetBoundsPt?.height ?? measuredHeightPt;
  const columnMinimums: number[] = Array.from({ length: inventory.columnCount }, () => dense ? 42 : 54);
  const rowMinimums: number[] = Array.from({ length: inventory.rowCount }, (_value, index) => {
    const fontPt = index === 0 ? variant.headerFontSizePt : variant.bodyFontSizePt;
    return Math.max(dense ? 20 : 24, fontPt * 1.2 + verticalPaddingPt * 2);
  });
  const limitingCellIds: string[] = [];
  const measuredCells = new Map(measuredObject.table.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  for (const cell of inventory.cells ?? []) {
    if (cell.horizontalMergeContinuation || cell.verticalMergeContinuation) continue;
    const measured = measuredCells.get(`${cell.row}:${cell.column}`);
    const textWidth = measured?.renderedTextBoundsPt?.width ?? 0;
    const textHeight = measured?.renderedTextBoundsPt?.height ?? 0;
    const wrapAdjustedWidth = measured && measured.lineCount > 2 && measured.boundsPt
      ? Math.max(textWidth, Math.min(widthPt * .6, measured.boundsPt.width * Math.min(1.8, 1 + (measured.lineCount - 2) * .2)))
      : textWidth;
    ensureSpanMinimum(columnMinimums, cell.column - 1, cell.columnSpan, wrapAdjustedWidth + horizontalPaddingPt * 2);
    ensureSpanMinimum(rowMinimums, cell.row - 1, cell.rowSpan, textHeight + verticalPaddingPt * 2);
    if (measured?.clearancesPt && Math.min(...Object.values(measured.clearancesPt)) < Math.min(horizontalPaddingPt, verticalPaddingPt)) limitingCellIds.push(cell.id);
  }
  const minimumRequiredWidthPt = columnMinimums.reduce((sum, value) => sum + value, 0);
  const minimumRequiredHeightPt = rowMinimums.reduce((sum, value) => sum + value, 0);
  const reasons: string[] = [];
  const recommendations: string[] = [];
  const minimumObservedFont = Math.min(...(inventory.cells ?? []).flatMap((cell) => cell.fontSizes).filter((value) => value > 0), Number.POSITIVE_INFINITY);
  if (minimumObservedFont < minimumFontPt) {
    reasons.push(`The table already contains ${minimumObservedFont} pt text, below the ${minimumFontPt} pt ${dense ? "dense technical" : "standard"} floor.`);
    recommendations.push("Increase the table region, reduce surrounding content, or use a continuation slide instead of shrinking type further.");
  }
  if (minimumRequiredWidthPt > widthPt + 0.5) {
    reasons.push(`The fixed table width is ${widthPt.toFixed(1)} pt but measured content and minimum padding require at least ${minimumRequiredWidthPt.toFixed(1)} pt.`);
    recommendations.push(`Widen the table by at least ${(minimumRequiredWidthPt - widthPt).toFixed(1)} pt or move it into a wider approved layout region.`);
  }
  if (minimumRequiredHeightPt > heightPt + 0.5) {
    reasons.push(`The fixed table height is ${heightPt.toFixed(1)} pt but measured content and minimum padding require at least ${minimumRequiredHeightPt.toFixed(1)} pt.`);
    recommendations.push(`Increase table height by at least ${(minimumRequiredHeightPt - heightPt).toFixed(1)} pt, reduce surrounding content, or continue the table on another slide.`);
  }
  const slideWidthPt = deck.audit.slideSize.width / EMU_PER_POINT;
  const slideHeightPt = deck.audit.slideSize.height / EMU_PER_POINT;
  const safeMarginPt = PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt;
  const targetWidthPt = Math.max(widthPt, minimumRequiredWidthPt);
  const targetHeightPt = Math.max(heightPt, minimumRequiredHeightPt);
  const fontFloorFailure = minimumObservedFont < minimumFontPt;
  const recommendedBoundsPt = !fontFloorFailure && targetWidthPt <= slideWidthPt - safeMarginPt * 2 && targetHeightPt <= slideHeightPt - safeMarginPt * 2
    ? {
      left: Math.max(safeMarginPt, Math.min(measuredObject.measuredGeometryPt.left, slideWidthPt - safeMarginPt - targetWidthPt)),
      top: Math.max(safeMarginPt, Math.min(measuredObject.measuredGeometryPt.top, slideHeightPt - safeMarginPt - targetHeightPt)),
      width: targetWidthPt,
      height: targetHeightPt,
    }
    : undefined;
  const measurableCells = measuredObject.table.cells.filter((cell) => cell.textLength > 0);
  const alreadyFits = !input.targetBoundsPt
    && reasons.length === 0
    && measurement.authority === "powerpoint-native"
    && measurableCells.length > 0
    && measurableCells.every((cell) => cell.marginsPt && cell.clearancesPt
      && cell.marginsPt.left >= horizontalPaddingPt - .2
      && cell.marginsPt.right >= horizontalPaddingPt - .2
      && cell.marginsPt.top >= verticalPaddingPt - .2
      && cell.marginsPt.bottom >= verticalPaddingPt - .2
      && cell.clearancesPt.left >= horizontalPaddingPt - .5
      && cell.clearancesPt.right >= horizontalPaddingPt - .5
      && cell.clearancesPt.top >= verticalPaddingPt - .5
      && cell.clearancesPt.bottom >= verticalPaddingPt - .5
      && cell.lineCount <= 2);
  if (alreadyFits) {
    return { status: "already-fit", diagnostics: { currentWidthPt: measuredWidthPt, currentHeightPt: measuredHeightPt, minimumRequiredWidthPt, minimumRequiredHeightPt, recommendedBoundsPt, limitingCellIds: [], reasons: [], recommendations: ["PowerPoint confirms that this table already satisfies the resolved type, padding, clearance, and wrap constraints; preserve it as-is."] } };
  }
  const columnWidths = allocate(widthPt, columnMinimums);
  const rowHeights = allocate(heightPt, rowMinimums);
  if (reasons.length || !columnWidths || !rowHeights) {
    return { status: "infeasible", diagnostics: { currentWidthPt: measuredWidthPt, currentHeightPt: measuredHeightPt, minimumRequiredWidthPt, minimumRequiredHeightPt, recommendedBoundsPt, limitingCellIds: [...new Set(limitingCellIds)], reasons, recommendations } };
  }
  const command: TableLayoutCommand = {
    id: `table-layout-${tableId}`,
    slideNumber: inventory.slideNumber,
    tableId,
    objectId: object.id,
    columnWidthsEmu: columnWidths.map(ptToEmu),
    rowHeightsEmu: rowHeights.map(ptToEmu),
    cellMarginsEmu: { left: ptToEmu(horizontalPaddingPt), right: ptToEmu(horizontalPaddingPt), top: ptToEmu(verticalPaddingPt), bottom: ptToEmu(verticalPaddingPt) },
    rationale: input.rationale.trim().slice(0, 700) || "Fit the native table with a deterministic cell-level solver while preserving exact cell text and merged-cell structure.",
    author: "ai",
    constraints: { minimumFontPt, minimumHorizontalPaddingPt: horizontalPaddingPt, minimumVerticalPaddingPt: verticalPaddingPt, preserveTableBounds: true },
    validation: { feasible: true, predictedOverflowCellIds: [], warnings: measurement.authority === "powerpoint-native" ? [] : ["This solution uses OOXML fallback measurements and must be remeasured in PowerPoint before acceptance."] },
  };
  return { status: "solved", command, diagnostics: { currentWidthPt: measuredWidthPt, currentHeightPt: measuredHeightPt, minimumRequiredWidthPt, minimumRequiredHeightPt, recommendedBoundsPt, limitingCellIds: [...new Set(limitingCellIds)], reasons: [], recommendations: [] } };
}

export function recommendedTableGrowthPlan(result: TableSolverResult, objectId: string, rationale: string): TableGrowthPlan | undefined {
  const bounds = result.diagnostics.recommendedBoundsPt;
  if (result.status !== "infeasible" || !bounds) return undefined;
  if (bounds.width <= result.diagnostics.currentWidthPt + .1 && bounds.height <= result.diagnostics.currentHeightPt + .1) return undefined;
  return {
    objectId,
    target: { x: ptToEmu(bounds.left), y: ptToEmu(bounds.top), width: ptToEmu(bounds.width), height: ptToEmu(bounds.height) },
    rationale: rationale.trim().slice(0, 700) || "Grow the native table by the minimum measured amount that satisfies readable cell geometry while preserving exact content and structure.",
  };
}
