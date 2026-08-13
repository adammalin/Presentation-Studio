import type { DeckJob } from "../types";
import type { GeometryEditRequest, VisualDesignRequest } from "./cleanup";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { calculateNativeTextOverflowEdges, type NativeMeasurementPacket } from "./native-measurement";

const EMU_PER_POINT = 12_700;

export interface TextFitSolverResult {
  status: "solved" | "already-fit" | "infeasible";
  geometry?: GeometryEditRequest;
  textStyle?: VisualDesignRequest["textStyles"][number];
  diagnostics: {
    currentFontPt?: number;
    targetFontPt?: number;
    currentHeightPt: number;
    requiredHeightPt: number;
    nativeOverflow: boolean;
    overflowEdges: Array<"left" | "top" | "right" | "bottom">;
    reasons: string[];
    recommendations: string[];
  };
}

export function nativeTextFrameOverflowEdges(measured: NativeMeasurementPacket["objects"][number], tolerancePt = .5): Array<"left" | "top" | "right" | "bottom"> {
  return calculateNativeTextOverflowEdges(measured, tolerancePt);
}

export function nativeTextFrameOverflows(measured: NativeMeasurementPacket["objects"][number], tolerancePt = .5) {
  return nativeTextFrameOverflowEdges(measured, tolerancePt).length > 0;
}

export function solveTextFit(input: { deck: DeckJob; measurement: NativeMeasurementPacket; objectId: string; rationale: string }): TextFitSolverResult {
  const { deck, measurement, objectId } = input;
  if (!deck.audit || !deck.scene) throw new Error("A current audit and hybrid scene are required before fitting text.");
  const object = deck.scene.objects.find((item) => item.id === objectId);
  const auditObject = deck.audit.editableObjects.find((item) => item.id === objectId);
  const textBox = object ? deck.audit.textBoxes.find((item) => item.slideNumber === object.slideNumber && item.shapeId === object.shapeId) : undefined;
  const measured = measurement.objects.find((item) => item.objectId === objectId);
  if (!object || !auditObject || !textBox || !measured?.measuredGeometryPt || !measured.text?.renderedBoundsPt || !measured.text.marginsPt) throw new Error("The requested object does not have complete source-bound text geometry and native measurement.");
  if (!object.operations.move || !object.operations.resize || !object.operations.restyle || object.protected) throw new Error("The requested text object is protected or cannot be safely fitted by the current scene contract.");
  const fontSizes = [...new Set(textBox.fontSizes.filter((value) => value > 0))];
  const currentFontPt = fontSizes.length === 1 ? fontSizes[0] : undefined;
  const minimumFontPt = object.semanticRole === "caption" || object.semanticRole === "label" ? PRESENTATION_DESIGN_STANDARD.defaults.typography.captionMinimumPt : PRESENTATION_DESIGN_STANDARD.defaults.typography.bodyMinimumPt;
  const targetFontPt = currentFontPt === undefined ? undefined : Math.max(currentFontPt, minimumFontPt);
  const reasons: string[] = [];
  const recommendations: string[] = [];
  if (measurement.authority !== "powerpoint-native" || measured.provenance.authority !== "powerpoint-native") {
    reasons.push("PowerPoint-native rendered-text measurement is unavailable for this object; OOXML estimates cannot prove a non-clipping fit.");
    recommendations.push("Unlock/open Microsoft PowerPoint and rerun native measurement before fitting text.");
  }
  if (fontSizes.length > 1 && Math.min(...fontSizes) < minimumFontPt) {
    reasons.push(`The text frame has mixed type sizes including values below the ${minimumFontPt} pt ${object.semanticRole} floor; a uniform font-size mutation could destroy intentional hierarchy.`);
    recommendations.push("Use a wider/taller approved region or a targeted run-level typography operation before fitting the frame.");
  }
  const box = measured.measuredGeometryPt;
  const text = measured.text.renderedBoundsPt;
  const margins = measured.text.marginsPt;
  const scale = currentFontPt && targetFontPt ? targetFontPt / currentFontPt : 1;
  const requiredHeightPt = Math.max(box.height, text.height * scale + margins.top + margins.bottom);
  const overflowEdges = nativeTextFrameOverflowEdges(measured);
  const nativeOverflow = overflowEdges.length > 0;
  if (overflowEdges.includes("left") || overflowEdges.includes("right")) {
    reasons.push(`PowerPoint reports horizontal text clipping at the ${overflowEdges.filter((edge) => edge === "left" || edge === "right").join(" and ")} edge; increasing frame height cannot resolve it safely.`);
    recommendations.push("Widen the text region or recompose the surrounding content into an approved layout.");
  }
  if (reasons.length) return { status: "infeasible", diagnostics: { currentFontPt, targetFontPt, currentHeightPt: box.height, requiredHeightPt, nativeOverflow, overflowEdges, reasons, recommendations } };
  const needsFont = Boolean(currentFontPt && targetFontPt && targetFontPt > currentFontPt + .01);
  const needsHeight = requiredHeightPt > box.height + .25 || nativeOverflow;
  if (!needsFont && !needsHeight) return { status: "already-fit", diagnostics: { currentFontPt, targetFontPt, currentHeightPt: box.height, requiredHeightPt, nativeOverflow, overflowEdges, reasons: [], recommendations: [] } };
  const anchor = measured.text.verticalAnchor.toLowerCase();
  const targetHeightPt = Math.max(requiredHeightPt, box.height);
  const heightDeltaPt = targetHeightPt - box.height;
  const targetTopPt = anchor.includes("bottom") ? box.top - heightDeltaPt : anchor.includes("middle") ? box.top - heightDeltaPt / 2 : box.top;
  const slideHeightPt = deck.audit.slideSize.height / EMU_PER_POINT;
  const safeMarginPt = PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt;
  if (targetTopPt < safeMarginPt - .1 || targetTopPt + targetHeightPt > slideHeightPt - safeMarginPt + .1) {
    reasons.push(`The minimum measured fit would cross the ${safeMarginPt} pt slide safe region.`);
    recommendations.push("Move related content as a group, use a taller approved region, or continue the content onto another slide.");
    return { status: "infeasible", diagnostics: { currentFontPt, targetFontPt, currentHeightPt: box.height, requiredHeightPt, nativeOverflow, overflowEdges, reasons, recommendations } };
  }
  const geometry: GeometryEditRequest | undefined = heightDeltaPt > .1 ? {
    objectId,
    target: { x: auditObject.geometry.x, y: Math.round(targetTopPt * EMU_PER_POINT), width: auditObject.geometry.width, height: Math.round(targetHeightPt * EMU_PER_POINT) },
    rationale: input.rationale.trim().slice(0, 700) || "Grow the native text frame by the minimum measured amount while preserving its vertical anchor and exact copy.",
    author: "ai",
    constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false },
  } : undefined;
  const textStyle: VisualDesignRequest["textStyles"][number] | undefined = needsFont && targetFontPt ? { objectId, fontSizePt: targetFontPt, rationale: `Raise uniform ${object.semanticRole} type to the resolved ${minimumFontPt} pt readability floor, then verify the PowerPoint-native fit.`, author: "ai" } : undefined;
  return { status: "solved", geometry, textStyle, diagnostics: { currentFontPt, targetFontPt, currentHeightPt: box.height, requiredHeightPt, nativeOverflow, overflowEdges, reasons: [], recommendations: [] } };
}
