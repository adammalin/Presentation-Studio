import type { DeckJob } from "../types";
import type { DesignMetricsReport, SlideDesignMetrics } from "./design-metrics";
import type { NativeMeasurementPacket } from "./native-measurement";
import type { NativeRenderResult } from "./desktop";

export const INSPECTION_PACKET_SCHEMA = "presentation-studio/inspection-packet" as const;
export const INSPECTION_PACKET_VERSION = 1 as const;

export interface InspectionCropRegion {
  id: string;
  kind: "title" | "table" | "text" | "issue";
  objectIds: string[];
  normalized: { x: number; y: number; width: number; height: number };
  reason: string;
}

function clamp(value: number, minimum: number, maximum: number) { return Math.min(maximum, Math.max(minimum, value)); }

function normalizedRegion(bounds: { left: number; top: number; width: number; height: number }, slideWidth: number, slideHeight: number, paddingPt = 12) {
  const left = clamp(bounds.left - paddingPt, 0, slideWidth);
  const top = clamp(bounds.top - paddingPt, 0, slideHeight);
  const right = clamp(bounds.left + bounds.width + paddingPt, 0, slideWidth);
  const bottom = clamp(bounds.top + bounds.height + paddingPt, 0, slideHeight);
  return { x: left / slideWidth, y: top / slideHeight, width: Math.max(0.01, (right - left) / slideWidth), height: Math.max(0.01, (bottom - top) / slideHeight) };
}

export function inspectionCropRegions(deck: DeckJob, measurement: NativeMeasurementPacket, slideNumber: number): InspectionCropRegion[] {
  if (!deck.audit || !deck.scene) return [];
  const slideWidthPt = deck.audit.slideSize.width / 12_700;
  const slideHeightPt = deck.audit.slideSize.height / 12_700;
  const sceneObjects = deck.scene.objects.filter((object) => object.slideNumber === slideNumber);
  const measured = measurement.objects.filter((object) => object.slideNumber === slideNumber && object.measuredGeometryPt);
  const regions: InspectionCropRegion[] = [];
  const title = sceneObjects.find((object) => object.semanticRole === "title");
  const titleMeasurement = measured.find((object) => object.objectId === title?.id);
  if (titleMeasurement?.measuredGeometryPt) regions.push({ id: `slide-${slideNumber}-title`, kind: "title", objectIds: [titleMeasurement.objectId], normalized: normalizedRegion(titleMeasurement.measuredGeometryPt, slideWidthPt, slideHeightPt, 16), reason: "Inspect native title wrapping, optical left edge, and hierarchy at readable scale." });
  for (const object of measured.filter((item) => item.table)) {
    regions.push({ id: `slide-${slideNumber}-${object.tableId}-table`, kind: "table", objectIds: [object.objectId], normalized: normalizedRegion(object.measuredGeometryPt!, slideWidthPt, slideHeightPt, 18), reason: "Inspect cell padding, wrapping, row rhythm, column balance, header hierarchy, and rule consistency." });
  }
  for (const object of measured.filter((item) => item.text?.renderedBoundsPt && item.text.textLength > 0).slice(0, 5)) {
    if (object.objectId === title?.id) continue;
    regions.push({ id: `slide-${slideNumber}-${object.objectId}-text`, kind: "text", objectIds: [object.objectId], normalized: normalizedRegion(object.measuredGeometryPt!, slideWidthPt, slideHeightPt, 10), reason: "Inspect native text fit, leading, alignment, and relationship to neighboring content." });
  }
  return regions.slice(0, 8);
}

export function buildInspectionPacket(input: {
  deck: DeckJob;
  slideNumber: number;
  projectUpdatedAt: string;
  workOrder: unknown;
  render: NativeRenderResult;
  measurement: NativeMeasurementPacket;
  metrics: DesignMetricsReport | SlideDesignMetrics;
}) {
  const slide = input.deck.audit?.slides.find((item) => item.number === input.slideNumber);
  const raster = input.render.status === "ready" ? input.render.slides.find((item) => item.number === input.slideNumber) : undefined;
  if (!slide) throw new Error("The requested slide is not present in the current audit.");
  return {
    schema: INSPECTION_PACKET_SCHEMA,
    version: INSPECTION_PACKET_VERSION,
    revision: `${input.projectUpdatedAt}:${input.deck.scene?.revision}:slide-${input.slideNumber}:raster-${raster?.sha256 ?? "unavailable"}:measurement-${input.measurement.revision}`,
    projectUpdatedAt: input.projectUpdatedAt,
    deck: { id: input.deck.id, name: input.deck.name, sourceSha256: input.deck.sourceSha256 },
    slide: { id: slide.id, number: slide.number, textHash: slide.textHash },
    workOrder: input.workOrder,
    measurement: input.measurement,
    metrics: input.metrics,
    visualEvidence: {
      authority: raster && input.render.authoritative ? "powerpoint-native" : "unavailable",
      renderer: input.render.renderer,
      pipeline: input.render.pipeline,
      powerPointVersion: input.render.powerPointVersion,
      rasterSha256: raster?.sha256,
      width: raster?.width,
      height: raster?.height,
      crops: inspectionCropRegions(input.deck, input.measurement, input.slideNumber),
    },
    instruction: "Use pixels for gestalt and design judgment; use PowerPoint-native measurements and deterministic solvers for exact geometry. Never estimate point adjustments from the image alone.",
  };
}
