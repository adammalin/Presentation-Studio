import assert from "node:assert/strict";
import test from "node:test";
import { calculateSlideDesignMetrics } from "../src/lib/design-metrics";
import type { NativeMeasurementPacket } from "../src/lib/native-measurement";
import type { DeckJob, PresentationSceneObject } from "../src/types";

function sceneObject(id: string, semanticRole: PresentationSceneObject["semanticRole"]): PresentationSceneObject {
  return {
    id,
    slideId: "slide-1",
    slideNumber: 1,
    shapeId: id,
    name: id,
    kind: semanticRole === "decoration" ? "shape" : "text",
    sourceElement: "p:sp",
    semanticRole,
    fidelityState: "editable-native",
    fidelityReason: "test",
    geometry: { x: 0, y: 0, width: 1, height: 1, rotation: 0 },
    zIndex: 0,
    sourceLocator: { slidePart: "ppt/slides/slide1.xml", shapeId: id },
    representation: { geometry: "native", text: semanticRole === "body" ? "native" : "none", style: "partial", internalStructure: "partial" },
    operations: { move: true, resize: true, restyle: true, editText: false, editTableStyle: false, replaceMedia: false, editChartData: false, editInternalStructure: false },
    protected: false,
  };
}

test("design metrics exclude intentional full-bleed brand bands but retain real safe-region defects", () => {
  const objects = [sceneObject("brand-band", "decoration"), sceneObject("body", "body"), sceneObject("small-callout", "decoration")];
  const deck = {
    id: "metrics",
    name: "metrics.pptx",
    sourceResourceId: "source",
    sourceSha256: "a".repeat(64),
    operationScope: "reflow",
    templateClassification: "custom",
    status: "ready-for-cleanup",
    protectedSlideNumbers: [],
    audit: { slideSize: { width: 720 * 12_700, height: 405 * 12_700 } },
    scene: { objects },
  } as unknown as DeckJob;
  const packet = {
    schema: "presentation-studio/native-measurement-packet",
    version: 1,
    status: "ready",
    revision: "r1",
    sourceSha256: deck.sourceSha256,
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    generatedAt: "2026-08-12T00:00:00.000Z",
    warnings: [],
    objects: [
      { objectId: "brand-band", shapeId: "brand-band", slideNumber: 1, sourceGeometryPt: { left: 0, top: -4, width: 720, height: 44 }, measuredGeometryPt: { left: 0, top: -4, width: 720, height: 44 }, binding: { method: "shape-id", confidence: "high" }, provenance: { authority: "powerpoint-native", adapter: "macos-powerpoint-applescript", confidence: "high", note: "test" } },
      { objectId: "body", shapeId: "body", slideNumber: 1, sourceGeometryPt: { left: 10, top: 60, width: 200, height: 40 }, measuredGeometryPt: { left: 10, top: 60, width: 200, height: 40 }, binding: { method: "shape-id", confidence: "high" }, provenance: { authority: "powerpoint-native", adapter: "macos-powerpoint-applescript", confidence: "high", note: "test" } },
      { objectId: "small-callout", shapeId: "small-callout", slideNumber: 1, sourceGeometryPt: { left: 14, top: 180, width: 90, height: 50 }, measuredGeometryPt: { left: 14, top: 180, width: 90, height: 50 }, binding: { method: "shape-id", confidence: "high" }, provenance: { authority: "powerpoint-native", adapter: "macos-powerpoint-applescript", confidence: "high", note: "test" } },
    ],
  } as NativeMeasurementPacket;

  const metrics = calculateSlideDesignMetrics(deck, packet, 1);
  assert.equal(metrics.safeRegionViolationCount, 2);
  assert.equal(metrics.offSlideObjectCount, 0);
});
