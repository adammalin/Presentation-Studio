import assert from "node:assert/strict";
import test from "node:test";
import type { NativeMeasurementResult } from "../src/lib/desktop";
import { critiqueStudioSlide } from "../src/lib/studio-visual-critic";
import { applyStudioDeterministicRepairPass } from "../src/lib/studio-repair-pass";
import type { StudioWebNode, StudioWebScene } from "../src/types";

const PT = 12_700;

function textNode(id: string, role: StudioWebNode["role"], x: number, y: number, width: number, height: number, fontSizePt: number): StudioWebNode {
  const frame = { x: x * PT, y: y * PT, width: width * PT, height: height * PT, rotation: 0 };
  return { id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "editable-object", name: id, kind: "text", role, sourceFrame: frame, frame, zIndex: 1, sourceTextOrder: 1, visible: true, locked: false, exactContent: true, text: id, style: { fontFamily: "Aptos", fontSizePt, fontWeight: role === "title" ? 700 : 400, lineHeight: 1.1, color: "#373A36", borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } } };
}

function scene(): StudioWebScene {
  const title = textNode("title", "title", 24, 30, 300, 40, 24);
  const body = textNode("body", "body", 32, 120, 160, 60, 26);
  return {
    schema: "presentation-studio/web-scene", version: 5, revision: "sha:web-v5:test", deckId: "deck", sourceSha256: "a".repeat(64), slideSize: { width: 960 * PT, height: 540 * PT }, sourceSlideSize: { width: 960 * PT, height: 540 * PT }, rhythm: { safeMarginPt: 18, gridPt: 6, compactGapPt: 8, normalGapPt: 12, primaryGapPt: 18, captionGapPt: 8, titleContentGapPt: 18 }, designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{ id: "slide", slideNumber: 1, sourceSlideId: "source", sourceTextHash: "b".repeat(64), contentCoverage: { exactTextMapped: true, sourceCharacterCount: 9, mappedCharacterCount: 9, sourceTextBoxCount: 2, mappedTextNodeCount: 2, groupedOrUnsupportedTextPresent: false }, sourceRevision: "source", recipe: "ornl-title-content", background: "#FFFFFF", status: "designed", designRationale: "test", figureTreatments: [], constraints: [{ id: "align", kind: "align", mode: "optical-left", nodeIds: [title.id, body.id], rationale: "Align visible text starts", author: "ai", evidenceAuthority: "scene-estimate", appliedAt: "2026-08-13T12:00:00.000Z" }], nodes: [title, body], updatedAt: "2026-08-13T12:00:00.000Z" }],
  };
}

function measurement(): NativeMeasurementResult {
  return { status: "ready", adapter: "macos-powerpoint-applescript", authority: "powerpoint-native", slides: [{ number: 1, shapeCount: 2, shapes: [
    { slideNumber: 1, shapeIndex: 1, name: "Title · title", zOrder: 1, boundsPt: { left: 24, top: 30, width: 300, height: 40 }, rotation: 0, hasTextFrame: true, hasTable: false, text: { coordinateSpace: "slide", textLength: 5, lineCount: 1, verticalAnchor: "top", marginsPt: { left: 0, right: 0, top: 0, bottom: 0 }, renderedBoundsPt: { left: 28, top: 30, width: 300, height: 40 } } },
    { slideNumber: 1, shapeIndex: 2, name: "Body · body", zOrder: 2, boundsPt: { left: 32, top: 120, width: 160, height: 60 }, rotation: 0, hasTextFrame: true, hasTable: false, text: { coordinateSpace: "slide", textLength: 4, lineCount: 2, verticalAnchor: "top", marginsPt: { left: 0, right: 0, top: 0, bottom: 0 }, renderedBoundsPt: { left: 40, top: 120, width: 150, height: 55 } } },
  ] }], warnings: [] };
}

test("Studio critic combines PowerPoint overflow, optical alignment, and hierarchy evidence", () => {
  const result = critiqueStudioSlide(scene(), 1, measurement());
  assert.equal(result.evidenceAuthority, "powerpoint-native");
  assert.equal(result.verdict, "revise");
  assert.equal(result.issues.some((issue) => issue.category === "overflow" && issue.severity === "blocker"), true);
  assert.equal(result.issues.some((issue) => issue.category === "alignment" && issue.source === "powerpoint-native"), true);
  assert.equal(result.issues.some((issue) => issue.category === "hierarchy"), true);
  assert.equal(result.iteration.maxPasses, 3);
});

test("Studio critic refuses non-native measurement evidence", () => {
  assert.throws(() => critiqueStudioSlide(scene(), 1, { ...measurement(), authority: "direct-ooxml" }), /Microsoft PowerPoint/i);
});

test("deterministic Studio repair replays native optical constraints and restores title hierarchy without changing copy", () => {
  const source = scene();
  const sourceText = source.slides[0].nodes.map((node) => node.text);
  const result = applyStudioDeterministicRepairPass(source, 1, measurement());
  assert.equal(result.requiresNativeRerender, true);
  assert.equal(result.actions.some((action) => action.operation === "reapply-constraint" && action.status === "fixed"), true);
  assert.equal(result.actions.some((action) => action.operation === "restore-title-hierarchy" && action.status === "fixed"), true);
  assert.equal(result.actions.some((action) => action.issueId.startsWith("overflow") && action.status === "deferred"), true);
  assert.deepEqual(result.scene.slides[0].nodes.map((node) => node.text), sourceText);
  assert.equal(result.scene.slides[0].nodes.find((node) => node.id === "title")?.style.fontSizePt, 30);
});

test("deterministic Studio repair grows a vertically overflowing text frame only by native measured need", () => {
  const source = scene();
  source.slides[0] = { ...source.slides[0], constraints: [], nodes: source.slides[0].nodes.map((node) => node.id === "title" ? { ...node, style: { ...node.style, fontSizePt: 32 } } : node.id === "body" ? { ...node, style: { ...node.style, fontSizePt: 16 } } : node) };
  const native = measurement();
  native.slides[0].shapes[0].text!.renderedBoundsPt = { left: 24, top: 30, width: 290, height: 35 };
  native.slides[0].shapes[1].text!.renderedBoundsPt = { left: 32, top: 120, width: 150, height: 70 };
  const before = source.slides[0].nodes.find((node) => node.id === "body")!.frame;
  const result = applyStudioDeterministicRepairPass(source, 1, native);
  const after = result.scene.slides[0].nodes.find((node) => node.id === "body")!.frame;
  assert.equal(result.actions.some((action) => action.operation === "grow-text-frame" && action.status === "fixed"), true);
  assert.ok(after.height > before.height);
  assert.equal(after.x, before.x);
  assert.equal(after.y, before.y);
});
