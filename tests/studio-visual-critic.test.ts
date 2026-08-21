import assert from "node:assert/strict";
import test from "node:test";
import type { NativeMeasurementResult } from "../src/lib/desktop";
import { critiqueStudioSlide, nativeStudioProductionIssues, preflightStudioScene } from "../src/lib/studio-visual-critic";
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

test("Studio critic blocks visible source media and source-locked treatments missing from the compiled candidate", () => {
  const source = scene();
  const seed = source.slides[0].nodes.find((node) => node.id === "body")!;
  const image: StudioWebNode = { ...seed, id: "partner-logo", sourceObjectId: "partner-logo", sourceShapeId: "partner-logo", name: "Partner logo", kind: "image", role: "image", text: undefined, textHash: undefined, sourceParagraphs: undefined, mediaPart: "logo.png", exactContent: false, style: { ...seed.style, objectFit: "contain" } };
  source.slides[0].nodes.push(image);
  let result = critiqueStudioSlide(source, 1, measurement(), { renderedNodeIds: ["title", "body"], renderedFigureTreatmentIds: [] });
  assert.equal(result.issues.some((issue) => issue.category === "figure" && issue.severity === "blocker" && issue.nodeIds.includes(image.id)), true);

  source.slides[0].figureTreatments = [{
    id: "locked-logo",
    nodeIds: [image.id],
    mode: "preserve-as-unit",
    verificationStatus: "source-locked",
    intentSummary: "Complete source logo",
    informationInventory: ["Every source pixel"],
    invariants: ["No omission"],
    rationale: "Preserve exact identity.",
  }];
  result = critiqueStudioSlide(source, 1, measurement(), { renderedNodeIds: ["title", "body"], renderedFigureTreatmentIds: [] });
  assert.equal(result.issues.some((issue) => issue.category === "figure" && issue.severity === "blocker" && /source-locked figure treatment/.test(issue.message)), true);
});

test("Studio production preflight blocks undersized ordinary type and crop-prone protected marks", () => {
  const source = scene();
  const body = source.slides[0].nodes.find((node) => node.id === "body")!;
  const logoFrame = { x: 760 * PT, y: 470 * PT, width: 120 * PT, height: 35 * PT, rotation: 0 };
  const logo: StudioWebNode = { ...body, id: "doe-logo", sourceObjectId: "doe-logo", sourceShapeId: "doe-logo", name: "DOE logo", kind: "image", role: "image", text: undefined, exactContent: false, sourceFrame: logoFrame, frame: logoFrame, style: { ...body.style, fontSizePt: 16, objectFit: "cover" } };
  source.slides[0] = { ...source.slides[0], nodes: source.slides[0].nodes.map((node) => node.id === "body" ? { ...node, style: { ...node.style, fontSizePt: 12 } } : node).concat(logo) };
  const result = preflightStudioScene(source);
  assert.equal(result.ready, false);
  assert.equal(result.issues.some((issue) => issue.category === "legibility" && /12 pt/.test(issue.message)), true);
  assert.equal(result.issues.some((issue) => issue.category === "brand" && issue.severity === "blocker"), true);
});

test("Studio production preflight preserves approved sacred-template typography", () => {
  const source = scene();
  source.slides[0] = { ...source.slides[0], recipe: "template-layout", nodes: source.slides[0].nodes.map((node) => node.id === "body" ? { ...node, style: { ...node.style, fontSizePt: 14 } } : node) };
  const result = preflightStudioScene(source, { protectedSlideNumbers: [1] });
  assert.equal(result.issues.some((issue) => issue.category === "legibility"), false);
});

test("Studio production preflight honors the qualified 10.5 pt editorial-record grid exceptions only for those components", () => {
  const source = scene();
  const editorial = source.slides[0].nodes.find((node) => node.id === "body")!;
  source.slides[0] = {
    ...source.slides[0],
    recipe: "ornl-title-card-grid",
    nodes: source.slides[0].nodes.map((node) => node.id === editorial.id ? {
      ...node,
      component: { groupId: "studio-editorial-record-grid-1", role: "technical-annotation", ordinal: 0 },
      style: { ...node.style, fontSizePt: 10.5 },
    } : node),
  };
  const qualified = preflightStudioScene(source);
  assert.equal(qualified.issues.some((issue) => issue.nodeIds.includes(editorial.id) && issue.category === "legibility"), false);

  const sourceGeometryQualified = structuredClone(source);
  const sourceGeometryNode = sourceGeometryQualified.slides[0].nodes.find((node) => node.id === editorial.id)!;
  sourceGeometryNode.component = { groupId: "studio-dense-source-grid-1", role: "technical-annotation", ordinal: 0 };
  sourceGeometryNode.style.fontSizePt = 8.5;
  assert.equal(preflightStudioScene(sourceGeometryQualified).issues.some((issue) => issue.nodeIds.includes(editorial.id) && issue.category === "legibility"), false);

  const tooSmall = structuredClone(source);
  tooSmall.slides[0].nodes.find((node) => node.id === editorial.id)!.style.fontSizePt = 10;
  assert.equal(preflightStudioScene(tooSmall).issues.some((issue) => issue.nodeIds.includes(editorial.id) && /10 pt/.test(issue.message)), true);

  const ordinary = structuredClone(source);
  const ordinaryNode = ordinary.slides[0].nodes.find((node) => node.id === editorial.id)!;
  ordinaryNode.component = { groupId: "studio-card-1", role: "card-body", ordinal: 0 };
  assert.equal(preflightStudioScene(ordinary).issues.some((issue) => issue.nodeIds.includes(editorial.id) && /10.5 pt/.test(issue.message)), true);
});

test("Studio critic uses source-relative legibility for dense peer-logo walls", () => {
  const source = scene();
  const seed = source.slides[0].nodes.find((node) => node.id === "body")!;
  const sourceFrame = { x: 250 * PT, y: 180 * PT, width: 20 * PT, height: 40 * PT, rotation: 0 };
  const logo: StudioWebNode = {
    ...seed,
    id: "narrow-authentic-logo",
    sourceObjectId: "narrow-authentic-logo",
    sourceShapeId: "narrow-authentic-logo",
    name: "Narrow authentic partner mark",
    kind: "image",
    role: "image",
    text: undefined,
    exactContent: false,
    sourceFrame,
    frame: { ...sourceFrame, width: 18 * PT, height: 38 * PT },
    component: { groupId: "studio-logo-grid-1", role: "logo-grid-item", ordinal: 0 },
    style: { ...seed.style, objectFit: "contain" },
  };
  source.slides[0].nodes.push(logo);
  const acceptable = critiqueStudioSlide(source, 1, measurement(), { renderedNodeIds: ["title", "body", logo.id], renderedFigureTreatmentIds: [] });
  assert.equal(acceptable.issues.some((issue) => issue.nodeIds.includes(logo.id) && issue.category === "legibility"), false);

  logo.frame = { ...logo.frame, width: 12 * PT };
  const reduced = critiqueStudioSlide(source, 1, measurement(), { renderedNodeIds: ["title", "body", logo.id], renderedFigureTreatmentIds: [] });
  assert.equal(reduced.issues.some((issue) => issue.nodeIds.includes(logo.id) && issue.category === "legibility"), true);
});

test("Studio critic allows a native-clean dense editorial record grid without flattening its hierarchy", () => {
  const source = scene();
  source.slides[0].constraints = [];
  source.slides[0].nodes = [source.slides[0].nodes[0], ...Array.from({ length: 7 }, (_, index) => {
    const record = textNode(`record-${index + 1}`, "body", 36 + (index % 2) * 430, 100 + Math.floor(index / 2) * 92, 390, 78, 10.5);
    record.text = `${2020 + index}\nA substantive laboratory award title that must remain visibly emphasized\nNamed researcher and complete source attribution ${"evidence ".repeat(11)}`;
    record.component = { groupId: `studio-editorial-record-grid-${index + 1}`, role: "technical-annotation", ordinal: index };
    return record;
  })];
  const result = critiqueStudioSlide(source, 1, measurement());
  assert.equal(result.issues.some((item) => item.category === "legibility" && item.severity === "major" && /composition carries/.test(item.message)), false);
  assert.equal(result.issues.some((item) => item.category === "legibility" && item.severity === "minor" && /editorial-record grid/.test(item.message)), true);
});

test("Studio critic accepts complete supporting metric-icon crops while retaining the strict technical-figure crop gate", () => {
  const supporting = scene();
  supporting.slides[0].figureTreatments = [{
    id: "studio-auto-metric-icon-1",
    nodeIds: ["body"],
    mode: "preserve-and-frame",
    verificationStatus: "source-locked",
    intentSummary: "Source metric icon",
    informationInventory: ["Complete source icon"],
    invariants: ["Keep paired with its metric"],
    rationale: "Use the small source icon as supporting evidence.",
    groupFrame: { x: 40 * PT, y: 220 * PT, width: 32 * PT, height: 32 * PT, rotation: 0 },
    crop: { left: 0, top: 0, right: .8, bottom: 0 },
    lockAspectRatio: true,
  }];
  const supportingResult = critiqueStudioSlide(supporting, 1, measurement());
  assert.equal(supportingResult.issues.some((issue) => issue.category === "figure"), false);

  const technical = structuredClone(supporting);
  technical.slides[0].figureTreatments[0] = { ...technical.slides[0].figureTreatments[0], id: "technical-figure", intentSummary: "Technical evidence figure" };
  const technicalResult = critiqueStudioSlide(technical, 1, measurement());
  assert.equal(technicalResult.issues.some((issue) => issue.category === "figure" && /less than 25%/.test(issue.message)), true);
});

test("Studio critic judges a source-locked treatment as one unit instead of flagging its hidden legacy inventory", () => {
  const source = scene();
  const legacy = source.slides[0].nodes.find((node) => node.id === "body")!;
  legacy.kind = "native-object";
  legacy.role = "group";
  legacy.text = undefined;
  legacy.frame = { x: 4 * PT, y: 4 * PT, width: 48 * PT, height: 32 * PT, rotation: 0 };
  source.slides[0].figureTreatments = [{
    id: "source-locked-legacy-inventory",
    nodeIds: [legacy.id],
    mode: "preserve-as-unit",
    verificationStatus: "source-locked",
    intentSummary: "Complete source relationship unit",
    informationInventory: ["Every source member"],
    invariants: ["Preserve internal relationships"],
    rationale: "Judge the group frame, not obsolete member geometry.",
    groupFrame: { x: 90 * PT, y: 120 * PT, width: 420 * PT, height: 220 * PT, rotation: 0 },
    lockAspectRatio: true,
  }];
  const result = critiqueStudioSlide(source, 1, measurement());
  assert.equal(result.issues.some((item) => item.nodeIds.includes(legacy.id) && ["safe-region", "legibility"].includes(item.category)), false);
});

test("native production QA treats editable table-cell overflow as a hard issue", () => {
  const native = measurement();
  native.slides[0].shapes.push({ slideNumber: 1, shapeIndex: 3, name: "Results table", zOrder: 3, boundsPt: { left: 30, top: 210, width: 300, height: 100 }, rotation: 0, hasTextFrame: false, hasTable: true, table: { rowCount: 1, columnCount: 1, rowHeightsPt: [100], columnWidthsPt: [300], cells: [{ row: 1, column: 1, boundsPt: { left: 0, top: 0, width: 300, height: 100 }, marginsPt: { left: 6, right: 6, top: 4, bottom: 4 }, renderedTextBoundsPt: { left: 6, top: 4, width: 290, height: 96 }, textCoordinateSpace: "cell-relative", textLength: 20, lineCount: 5, verticalAnchor: "top" }] } });
  assert.equal(nativeStudioProductionIssues(native).some((issue) => /cell r1c1 overflows/.test(issue.message)), true);
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
