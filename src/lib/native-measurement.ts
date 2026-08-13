import type { DeckJob, PresentationSceneObject, PresentationSceneTable } from "../types";
import type {
  MeasurementAuthority,
  NativeBoundsPt,
  NativeMeasurementResult,
  NativeShapeMeasurement,
  NativeTableCellMeasurement,
} from "./desktop";

const EMU_PER_POINT = 12_700;

export const NATIVE_MEASUREMENT_PACKET_SCHEMA = "presentation-studio/native-measurement-packet" as const;
export const NATIVE_MEASUREMENT_PACKET_VERSION = 1 as const;

export interface MeasurementProvenance {
  authority: MeasurementAuthority;
  adapter: string;
  confidence: "high" | "medium" | "low";
  note: string;
}

export interface BoundCellMeasurement extends NativeTableCellMeasurement {
  cellId: string;
  clearancesPt?: { left: number; right: number; top: number; bottom: number };
  clearanceProvenance?: MeasurementProvenance;
}

export interface BoundObjectMeasurement {
  objectId: string;
  shapeId: string;
  tableId?: string;
  slideNumber: number;
  sourceGeometryPt: NativeBoundsPt;
  measuredGeometryPt?: NativeBoundsPt;
  text?: NativeShapeMeasurement["text"];
  table?: Omit<NonNullable<NativeShapeMeasurement["table"]>, "cells"> & { cells: BoundCellMeasurement[] };
  binding: { method: "shape-id" | "name" | "z-order" | "geometry" | "ooxml-fallback"; confidence: "high" | "medium" | "low"; nativeShapeIndex?: number };
  provenance: MeasurementProvenance;
}

export interface NativeMeasurementPacket {
  schema: typeof NATIVE_MEASUREMENT_PACKET_SCHEMA;
  version: typeof NATIVE_MEASUREMENT_PACKET_VERSION;
  status: NativeMeasurementResult["status"] | "fallback";
  revision: string;
  sourceSha256: string;
  adapter: string;
  authority: MeasurementAuthority;
  powerPointVersion?: string;
  generatedAt: string;
  objects: BoundObjectMeasurement[];
  warnings: string[];
}

export function calculateNativeTextOverflowEdges(object: BoundObjectMeasurement, tolerancePt = .5): Array<"left" | "top" | "right" | "bottom"> {
  const box = object.measuredGeometryPt;
  const text = object.text?.renderedBoundsPt;
  const margins = object.text?.marginsPt;
  if (!box || !text || !margins || object.text?.coordinateSpace !== "slide") return [];
  const inner = { left: box.left + margins.left, top: box.top + margins.top, right: box.left + box.width - margins.right, bottom: box.top + box.height - margins.bottom };
  return [
    text.left < inner.left - tolerancePt ? "left" as const : undefined,
    text.top < inner.top - tolerancePt ? "top" as const : undefined,
    text.left + text.width > inner.right + tolerancePt ? "right" as const : undefined,
    text.top + text.height > inner.bottom + tolerancePt ? "bottom" as const : undefined,
  ].filter((edge): edge is "left" | "top" | "right" | "bottom" => Boolean(edge));
}

function emuToPt(value: number) {
  return value / EMU_PER_POINT;
}

function sourceGeometryPt(object: PresentationSceneObject): NativeBoundsPt {
  return { left: emuToPt(object.geometry.x), top: emuToPt(object.geometry.y), width: emuToPt(object.geometry.width), height: emuToPt(object.geometry.height) };
}

function geometryDistance(left: NativeBoundsPt, right: NativeBoundsPt) {
  return Math.abs(left.left - right.left) + Math.abs(left.top - right.top) + Math.abs(left.width - right.width) + Math.abs(left.height - right.height);
}

export function calculateNativeCellClearances(cell: NativeTableCellMeasurement) {
  const bounds = cell.boundsPt;
  const text = cell.renderedTextBoundsPt;
  const margins = cell.marginsPt;
  if (!bounds || !text || !margins) return undefined;
  const left = text.left;
  const right = bounds.width - text.left - text.width;
  const verticalRoom = bounds.height - margins.top - margins.bottom - text.height;
  const anchor = cell.verticalAnchor.toLowerCase();
  const top = anchor.includes("middle") ? margins.top + verticalRoom / 2 : anchor.includes("bottom") ? margins.top + verticalRoom : margins.top;
  const bottom = bounds.height - top - text.height;
  return { left, right, top, bottom };
}

function bindCells(tableScene: PresentationSceneTable | undefined, cells: NativeTableCellMeasurement[]): BoundCellMeasurement[] {
  return cells.map((cell) => {
    const sceneCell = tableScene?.cells.find((item) => item.row === cell.row && item.column === cell.column);
    return {
      ...cell,
      cellId: sceneCell?.id ?? `${tableScene?.id ?? "unbound-table"}-cell-r${cell.row}-c${cell.column}`,
      clearancesPt: calculateNativeCellClearances(cell),
      clearanceProvenance: {
        authority: "powerpoint-native",
        adapter: "macos-powerpoint-applescript",
        confidence: "high",
        note: "Horizontal bounds are native. Vertical edge clearance combines native cell size, margins, rendered text height, and PowerPoint vertical anchoring.",
      },
    };
  });
}

function fallbackObject(deck: DeckJob, object: PresentationSceneObject): BoundObjectMeasurement {
  const tableScene = deck.scene?.tables?.find((table) => table.objectId === object.id);
  const tableInventory = object.sourceLocator.tableId ? deck.audit?.tables.find((table) => table.id === object.sourceLocator.tableId) : undefined;
  const textBox = deck.audit?.textBoxes.find((text) => text.slideNumber === object.slideNumber && text.shapeId === object.shapeId);
  return {
    objectId: object.id,
    shapeId: object.shapeId,
    tableId: object.sourceLocator.tableId,
    slideNumber: object.slideNumber,
    sourceGeometryPt: sourceGeometryPt(object),
    measuredGeometryPt: sourceGeometryPt(object),
    text: textBox ? {
      marginsPt: {
        left: emuToPt(textBox.textInsets.left),
        right: emuToPt(textBox.textInsets.right),
        top: emuToPt(textBox.textInsets.top),
        bottom: emuToPt(textBox.textInsets.bottom),
      },
      renderedBoundsPt: {
        left: emuToPt(textBox.estimatedOpticalLeftEmu),
        top: emuToPt(textBox.geometry.y + textBox.textInsets.top),
        width: emuToPt(Math.max(0, textBox.geometry.width - textBox.textInsets.left - textBox.textInsets.right)),
        height: emuToPt(textBox.estimatedRequiredHeightEmu),
      },
      coordinateSpace: "slide",
      textLength: textBox.characterCount,
      lineCount: textBox.estimatedLineCount,
      verticalAnchor: textBox.verticalAlignment,
    } : undefined,
    table: tableInventory && tableScene ? {
      rowCount: tableInventory.rowCount,
      columnCount: tableInventory.columnCount,
      rowHeightsPt: (tableInventory.rows ?? []).map((row) => emuToPt(row.heightEmu)),
      columnWidthsPt: (tableInventory.columns ?? []).map((column) => emuToPt(column.widthEmu)),
      cells: (tableInventory.cells ?? []).map((cell) => ({
        row: cell.row,
        column: cell.column,
        boundsPt: tableScene.cells.find((item) => item.id === cell.id) ? (() => {
          const geometry = tableScene.cells.find((item) => item.id === cell.id)!.geometry;
          return { left: emuToPt(geometry.x), top: emuToPt(geometry.y), width: emuToPt(geometry.width), height: emuToPt(geometry.height) };
        })() : undefined,
        marginsPt: { left: emuToPt(cell.marginsEmu.left), right: emuToPt(cell.marginsEmu.right), top: emuToPt(cell.marginsEmu.top), bottom: emuToPt(cell.marginsEmu.bottom) },
        renderedTextBoundsPt: undefined,
        textCoordinateSpace: "cell-relative",
        textLength: cell.characterCount,
        lineCount: cell.paragraphCount,
        verticalAnchor: cell.verticalAlignment,
        cellId: cell.id,
      })),
    } : undefined,
    binding: { method: "ooxml-fallback", confidence: "medium" },
    provenance: { authority: "derived-ooxml", adapter: "ooxml-fallback", confidence: "medium", note: "Shape and table grid geometry come from OOXML; rendered text bounds remain heuristic until native PowerPoint measurement succeeds." },
  };
}

export function bindNativeMeasurement(deck: DeckJob, native?: NativeMeasurementResult): NativeMeasurementPacket {
  if (!deck.audit || !deck.scene) throw new Error("A current audit and scene are required before binding measurements.");
  const generatedAt = native?.generatedAt ?? new Date().toISOString();
  const nativeReady = native?.status === "ready" && native.authority === "powerpoint-native";
  const warnings = [...(native?.warnings ?? [])];
  const used = new Set<string>();
  const objects: BoundObjectMeasurement[] = deck.scene.objects.map((object): BoundObjectMeasurement => {
    if (!nativeReady) return fallbackObject(deck, object);
    const slide = native.slides.find((item) => item.number === object.slideNumber);
    const expected = sourceGeometryPt(object);
    const available = slide?.shapes.filter((shape) => !used.has(`${object.slideNumber}:${shape.shapeIndex}`)) ?? [];
    const idMatches = available.filter((shape) => shape.nativeShapeId === object.shapeId);
    const nameMatches = object.name ? available.filter((shape) => shape.name === object.name) : [];
    const zCandidate = available.find((shape) => shape.zOrder === object.zIndex + 1);
    let candidate = idMatches.length === 1 ? idMatches[0] : nameMatches.length === 1 ? nameMatches[0] : zCandidate;
    let method: BoundObjectMeasurement["binding"]["method"] = idMatches.length === 1 ? "shape-id" : nameMatches.length === 1 ? "name" : "z-order";
    let confidence: BoundObjectMeasurement["binding"]["confidence"] = idMatches.length === 1 || nameMatches.length === 1 ? "high" : zCandidate?.boundsPt && geometryDistance(expected, zCandidate.boundsPt) <= 2 ? "high" : "medium";
    const hasStableIdentity = method === "shape-id" || method === "name";
    if (!candidate?.boundsPt || (!hasStableIdentity && geometryDistance(expected, candidate.boundsPt) > 16)) {
      candidate = available.filter((shape) => shape.boundsPt)
        .sort((left, right) => geometryDistance(expected, left.boundsPt!) - geometryDistance(expected, right.boundsPt!))[0];
      method = "geometry";
      confidence = candidate?.boundsPt && geometryDistance(expected, candidate.boundsPt) <= 4 ? "high" : candidate?.boundsPt && geometryDistance(expected, candidate.boundsPt) <= 16 ? "medium" : "low";
    }
    if (!candidate) {
      warnings.push(`No native PowerPoint shape measurement could be bound to ${object.id}; OOXML fallback is used for that object.`);
      return fallbackObject(deck, object);
    }
    used.add(`${object.slideNumber}:${candidate.shapeIndex}`);
    const tableScene = deck.scene?.tables?.find((table) => table.objectId === object.id);
    return {
      objectId: object.id,
      shapeId: object.shapeId,
      tableId: object.sourceLocator.tableId,
      slideNumber: object.slideNumber,
      sourceGeometryPt: expected,
      measuredGeometryPt: candidate.boundsPt,
      text: candidate.text,
      table: candidate.table ? { ...candidate.table, cells: bindCells(tableScene, candidate.table.cells) } : undefined,
      binding: { method, confidence, nativeShapeIndex: candidate.shapeIndex },
      provenance: { authority: "powerpoint-native", adapter: native.adapter, confidence, note: "Measurements were returned by Microsoft PowerPoint and bound to the source scene without reading or rewriting visible copy." },
    };
  });
  return {
    schema: NATIVE_MEASUREMENT_PACKET_SCHEMA,
    version: NATIVE_MEASUREMENT_PACKET_VERSION,
    status: nativeReady ? "ready" : "fallback",
    revision: `${deck.scene.revision}:measurement-${nativeReady ? native.sourceSha256 ?? "native" : "ooxml"}`,
    sourceSha256: deck.sourceSha256,
    adapter: nativeReady ? native.adapter : "ooxml-fallback",
    authority: nativeReady ? "powerpoint-native" : "derived-ooxml",
    powerPointVersion: native?.powerPointVersion,
    generatedAt,
    objects,
    warnings,
  };
}

export function compareNativeMeasurementPackets(expected: NativeMeasurementPacket, actual: NativeMeasurementPacket, tolerancePt = 0.2) {
  const mismatches: string[] = [];
  if (expected.authority !== "powerpoint-native" || actual.authority !== "powerpoint-native") mismatches.push("Both measurement packets must be PowerPoint-native.");
  const compareNumber = (label: string, left: number | undefined, right: number | undefined) => {
    if (left === undefined && right === undefined) return;
    if (left === undefined || right === undefined || Math.abs(left - right) > tolerancePt) mismatches.push(label);
  };
  for (const object of expected.objects) {
    const candidate = actual.objects.find((item) => item.objectId === object.objectId);
    if (!candidate) { mismatches.push(`${object.objectId}:missing`); continue; }
    for (const key of ["left", "top", "width", "height"] as const) compareNumber(`${object.objectId}:geometry:${key}`, object.measuredGeometryPt?.[key], candidate.measuredGeometryPt?.[key]);
    for (const key of ["left", "top", "width", "height"] as const) compareNumber(`${object.objectId}:text:${key}`, object.text?.renderedBoundsPt?.[key], candidate.text?.renderedBoundsPt?.[key]);
    if ((object.text?.lineCount ?? 0) !== (candidate.text?.lineCount ?? 0)) mismatches.push(`${object.objectId}:text:line-count`);
    if (object.table || candidate.table) {
      if (!object.table || !candidate.table) { mismatches.push(`${object.objectId}:table:missing`); continue; }
      object.table.columnWidthsPt.forEach((value, index) => compareNumber(`${object.objectId}:column-${index + 1}`, value, candidate.table?.columnWidthsPt[index]));
      object.table.rowHeightsPt.forEach((value, index) => compareNumber(`${object.objectId}:row-${index + 1}`, value, candidate.table?.rowHeightsPt[index]));
      for (const cell of object.table.cells) {
        const actualCell = candidate.table.cells.find((item) => item.cellId === cell.cellId);
        if (!actualCell) { mismatches.push(`${cell.cellId}:missing`); continue; }
        for (const key of ["left", "top", "width", "height"] as const) compareNumber(`${cell.cellId}:text:${key}`, cell.renderedTextBoundsPt?.[key], actualCell.renderedTextBoundsPt?.[key]);
        if (cell.lineCount !== actualCell.lineCount) mismatches.push(`${cell.cellId}:line-count`);
      }
    }
  }
  return { equivalent: mismatches.length === 0, tolerancePt, mismatches: mismatches.slice(0, 100) };
}
