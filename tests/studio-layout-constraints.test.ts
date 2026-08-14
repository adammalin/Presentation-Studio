import assert from "node:assert/strict";
import test from "node:test";
import type { NativeMeasurementResult } from "../src/lib/desktop";
import { applyStudioLayoutConstraints } from "../src/lib/studio-layout-constraints";
import type { StudioWebNode, StudioWebScene } from "../src/types";

const PT = 12_700;

function node(id: string, xPt: number, yPt: number, widthPt = 24, heightPt = 12): StudioWebNode {
  const frame = { x: xPt * PT, y: yPt * PT, width: widthPt * PT, height: heightPt * PT, rotation: 0 };
  return {
    id,
    sourceObjectId: `source-${id}`,
    sourceShapeId: id,
    sourceBinding: "editable-object",
    name: id,
    kind: "text",
    role: "body",
    sourceFrame: frame,
    frame,
    zIndex: 1,
    sourceTextOrder: 1,
    visible: true,
    locked: false,
    exactContent: true,
    text: id,
    opticalInsets: { left: 4 * PT, top: 2 * PT, right: 0, bottom: 0, authority: "source-estimate", basis: "rendered-text" },
    style: { fontFamily: "Aptos", fontSizePt: 18, fontWeight: 400, lineHeight: 1.1, color: "#373A36", borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } },
  };
}

function scene(nodes: StudioWebNode[]): StudioWebScene {
  return {
    schema: "presentation-studio/web-scene",
    version: 5,
    revision: "source:web-v5:before",
    deckId: "deck",
    sourceSha256: "source",
    slideSize: { width: 960 * PT, height: 540 * PT },
    sourceSlideSize: { width: 960 * PT, height: 540 * PT },
    rhythm: { safeMarginPt: 18, gridPt: 6, compactGapPt: 8, normalGapPt: 12, primaryGapPt: 18, captionGapPt: 8, titleContentGapPt: 18 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{
      id: "slide-1",
      slideNumber: 1,
      sourceSlideId: "source-slide-1",
      sourceTextHash: "hash",
      contentCoverage: { exactTextMapped: true, sourceCharacterCount: 3, mappedCharacterCount: 3, sourceTextBoxCount: nodes.length, mappedTextNodeCount: nodes.length, groupedOrUnsupportedTextPresent: false },
      sourceRevision: "source",
      recipe: "ornl-title-content",
      background: "#FFFFFF",
      status: "designed",
      designRationale: "test",
      figureTreatments: [],
      constraints: [],
      nodes,
      updatedAt: "2026-08-13T12:00:00.000Z",
    }],
  };
}

test("Studio optical alignment carries PowerPoint-rendered text insets without AI coordinate arithmetic", () => {
  const first = node("first", 10, 80);
  const second = node("second", 50, 180);
  const measurement: NativeMeasurementResult = {
    status: "ready",
    adapter: "macos-powerpoint-applescript",
    authority: "powerpoint-native",
    slides: [{ number: 1, shapeCount: 2, shapes: [
      { slideNumber: 1, shapeIndex: 1, name: "first", zOrder: 1, boundsPt: { left: 10, top: 80, width: 24, height: 12 }, rotation: 0, hasTextFrame: true, hasTable: false, text: { coordinateSpace: "slide", textLength: 5, lineCount: 1, verticalAnchor: "top", renderedBoundsPt: { left: 14, top: 82, width: 18, height: 8 } } },
      { slideNumber: 1, shapeIndex: 2, name: "second", zOrder: 2, boundsPt: { left: 50, top: 180, width: 24, height: 12 }, rotation: 0, hasTextFrame: true, hasTable: false, text: { coordinateSpace: "slide", textLength: 6, lineCount: 1, verticalAnchor: "top", renderedBoundsPt: { left: 58, top: 182, width: 14, height: 8 } } },
    ] }],
    warnings: [],
  };
  const result = applyStudioLayoutConstraints(scene([first, second]), 1, [{ kind: "align", mode: "optical-left", nodeIds: [first.id, second.id], anchorNodeId: first.id, rationale: "Align the visible glyph starts.", author: "ai" }], measurement);
  const moved = result.scene.slides[0].nodes.find((item) => item.id === second.id)!;
  assert.equal(moved.frame.x, 6 * PT);
  assert.equal(result.evidenceAuthority, "powerpoint-native");
  assert.equal(result.constraints[0].evidenceAuthority, "powerpoint-native");
});

test("Studio safe-region fitting moves a declared relationship group as one intact figure", () => {
  const image = { ...node("figure", 4, 100, 60, 80), kind: "image" as const, text: undefined };
  const caption = node("caption", 8, 184, 52, 20);
  const beforeDelta = caption.frame.y - image.frame.y;
  const result = applyStudioLayoutConstraints(scene([image, caption]), 1, [{ kind: "fit-safe-region", mode: "both", nodeIds: [image.id, caption.id], groups: [[image.id, caption.id]], rationale: "Keep the figure and caption together inside the deck safe region.", author: "ai" }]);
  const afterImage = result.scene.slides[0].nodes.find((item) => item.id === image.id)!;
  const afterCaption = result.scene.slides[0].nodes.find((item) => item.id === caption.id)!;
  assert.equal(afterImage.frame.x, 18 * PT);
  assert.equal(afterCaption.frame.x, 22 * PT);
  assert.equal(afterCaption.frame.y - afterImage.frame.y, beforeDelta);
  assert.equal(result.evidenceAuthority, "scene-estimate");
  assert.match(result.diagnostics.join(" "), /PowerPoint-native build/i);
});

test("Studio constraints move a source-locked first-class figure through its group frame", () => {
  const image = { ...node("locked-figure", 2, 90, 120, 90), kind: "native-object" as const, text: undefined, locked: true };
  const label = node("figure-label", 4, 184, 110, 16);
  const input = scene([image, label]);
  input.slides[0].figureTreatments = [{ id: "technical-figure", nodeIds: [image.id, label.id], mode: "preserve-and-frame", verificationStatus: "source-locked", intentSummary: "Source-locked technical figure", informationInventory: ["figure", "label"], invariants: ["Preserve relationships"], rationale: "Move as one evidence unit", groupFrame: { x: 2 * PT, y: 90 * PT, width: 120 * PT, height: 110 * PT, rotation: 0 }, relationshipPolicy: "preserve-internal", lockAspectRatio: true }];
  const result = applyStudioLayoutConstraints(input, 1, [{ kind: "fit-safe-region", mode: "both", nodeIds: [image.id, label.id], groups: [[image.id, label.id]], rationale: "Move the complete locked figure inside the safe region.", author: "ai" }]);
  const treatment = result.scene.slides[0].figureTreatments[0];
  assert.equal(treatment.groupFrame?.x, 18 * PT);
  assert.equal(result.scene.slides[0].nodes.find((item) => item.id === image.id)?.frame.x, 18 * PT);
  assert.deepEqual(new Set(result.changedNodeIds), new Set([image.id, label.id]));
});

test("Studio constraints reject a newly introduced collision", () => {
  const first = node("first", 20, 100, 60, 20);
  const second = node("second", 220, 100, 60, 20);
  const obstacle = node("obstacle", 110, 100, 60, 20);
  assert.throws(() => applyStudioLayoutConstraints(scene([first, second, obstacle]), 1, [{ kind: "align", mode: "left", nodeIds: [first.id, second.id], anchorNodeId: obstacle.id, rationale: "Invalid collision test.", author: "ai" }]), /anchor must belong/i);
  assert.throws(() => applyStudioLayoutConstraints(scene([first, second, obstacle]), 1, [{ kind: "align", mode: "left", nodeIds: [first.id, second.id, obstacle.id], anchorNodeId: obstacle.id, rationale: "Invalid collision test.", author: "ai" }]), /new overlap/i);
});
