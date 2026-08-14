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

test("design metrics tolerate PowerPoint bullet glyph overhang without hiding ordinary text overflow", () => {
  const objects = [sceneObject("bullet", "body"), sceneObject("plain", "body")];
  const deck = {
    id: "bullet-metrics",
    name: "bullet-metrics.pptx",
    sourceResourceId: "source",
    sourceSha256: "b".repeat(64),
    operationScope: "reflow",
    templateClassification: "custom",
    status: "ready-for-cleanup",
    protectedSlideNumbers: [],
    audit: {
      slideSize: { width: 720 * 12_700, height: 405 * 12_700 },
      editableObjects: objects.map((object) => ({ id: object.id, slideNumber: 1, shapeId: object.shapeId, name: object.name })),
      textBoxes: objects.map((object) => ({ slideNumber: 1, shapeId: object.shapeId, bulletParagraphCount: object.id === "bullet" ? 1 : 0 })),
    },
    scene: { objects },
  } as unknown as DeckJob;
  const measured = (objectId: string) => ({
    objectId,
    shapeId: objectId,
    slideNumber: 1,
    sourceGeometryPt: { left: 100, top: objectId === "bullet" ? 100 : 180, width: 200, height: 60 },
    measuredGeometryPt: { left: 100, top: objectId === "bullet" ? 100 : 180, width: 200, height: 60 },
    text: { marginsPt: { left: 0, right: 0, top: 0, bottom: 0 }, renderedBoundsPt: { left: 118, top: objectId === "bullet" ? 100 : 180, width: 184.1, height: 30 }, coordinateSpace: "slide" as const, textLength: 80, lineCount: 2, verticalAnchor: "anchor top" },
    binding: { method: "shape-id" as const, confidence: "high" as const },
    provenance: { authority: "powerpoint-native" as const, adapter: "macos-powerpoint-applescript", confidence: "high" as const, note: "test" },
  });
  const packet = {
    schema: "presentation-studio/native-measurement-packet",
    version: 1,
    status: "ready",
    revision: "r2",
    sourceSha256: deck.sourceSha256,
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    generatedAt: "2026-08-14T00:00:00.000Z",
    warnings: [],
    objects: [measured("bullet"), measured("plain")],
  } as NativeMeasurementPacket;

  assert.equal(calculateSlideDesignMetrics(deck, packet, 1).textOverflowCount, 1);
});
