import assert from "node:assert/strict";
import test from "node:test";
import { reconstructStudioConcept } from "../src/lib/studio-concept-reconstruction";
import { attachStudioConceptReference } from "../src/lib/studio-concept-reference";
import { analyzeStudioDesignImpact } from "../src/lib/studio-design-impact";
import { createStudioVisualNeed } from "../src/lib/studio-visual-needs";
import type { ProjectResource, StudioWebNode, StudioWebScene } from "../src/types";

const PT = 12_700;

function node(id: string, kind: StudioWebNode["kind"], role: StudioWebNode["role"], x: number, y: number, width: number, height: number): StudioWebNode {
  const frame = { x: x * PT, y: y * PT, width: width * PT, height: height * PT, rotation: 0 };
  return {
    id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "editable-object", name: id, kind, role,
    sourceFrame: frame, frame, zIndex: 1, sourceTextOrder: 1, visible: true, locked: false, exactContent: true,
    text: kind === "text" ? `${id} exact source text` : undefined,
    style: { fontFamily: "Aptos", fontSizePt: role === "title" ? 30 : 18, fontWeight: role === "title" ? 700 : 400, lineHeight: 1.1, color: "#373A36", borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 }, objectFit: kind === "image" ? "contain" : undefined },
  };
}

function scene(): StudioWebScene {
  return {
    schema: "presentation-studio/web-scene", version: 5, revision: "source:web-v5:before", deckId: "deck", sourceSha256: "a".repeat(64),
    slideSize: { width: 960 * PT, height: 540 * PT }, sourceSlideSize: { width: 960 * PT, height: 540 * PT },
    rhythm: { safeMarginPt: 18, gridPt: 6, compactGapPt: 8, normalGapPt: 12, primaryGapPt: 18, captionGapPt: 8, titleContentGapPt: 18 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{
      id: "slide-2", slideNumber: 2, sourceSlideId: "source-2", sourceTextHash: "b".repeat(64), sourceRevision: "source", recipe: "source", background: "#FFFFFF", status: "imported", designRationale: "source", figureTreatments: [], conceptReferences: [], visualNeeds: [],
      contentCoverage: { exactTextMapped: true, sourceCharacterCount: 55, mappedCharacterCount: 55, sourceTextBoxCount: 2, mappedTextNodeCount: 2, groupedOrUnsupportedTextPresent: false },
      nodes: [node("title", "text", "title", 40, 30, 880, 60), node("body", "text", "body", 60, 140, 360, 220), node("image", "image", "image", 500, 150, 350, 220)],
      updatedAt: "2026-08-14T12:00:00.000Z",
    }],
  };
}

const conceptResource: ProjectResource = {
  id: "concept", name: "concept.png", mediaType: "image/png", byteLength: 3, sha256: "d".repeat(64), roles: ["concept-reference"], kind: "image", support: ["previewable", "placeable"], createdAt: "2026-08-14T12:00:00.000Z", embedded: true, bytes: new Uint8Array([1, 2, 3]), mcpAccess: "preview",
};

test("design impact does not mistake a type-only edit for redesign", () => {
  const slide = { ...scene().slides[0], status: "designed" as const, nodes: scene().slides[0].nodes.map((item) => item.id === "body" ? { ...item, style: { ...item.style, fontSizePt: 16 } } : item) };
  const impact = analyzeStudioDesignImpact(slide);
  assert.equal(impact.level, "typography-only");
  assert.equal(impact.meaningful, false);
});

test("design impact recognizes a material shared-recipe composition", () => {
  const source = scene().slides[0];
  const slide = { ...source, status: "designed" as const, recipe: "ornl-title-two-column" as const, nodes: source.nodes.map((item) => item.id === "body" ? { ...item, frame: { ...item.frame, x: item.frame.x + 100 * PT }, component: { groupId: "content", role: "card-body" as const } } : item) };
  const impact = analyzeStudioDesignImpact(slide);
  assert.equal(impact.level, "layout-redesign");
  assert.equal(impact.meaningful, true);
  assert.deepEqual(impact.geometryChangedNodeIds, ["body"]);
});

test("concept reconstruction maps approved zones to editable content and advances only after material change", () => {
  const created = createStudioVisualNeed(scene(), 2, { id: "need-layout", type: "layout-concept", reason: "The source layout is weak.", communicationJob: "Create one clear assertion and evidence relationship." });
  const attached = attachStudioConceptReference(created, 2, conceptResource, {
    id: "concept-layout", visualNeedId: "need-layout", origin: "imagegen", approvedInfluences: ["composition", "visual-hierarchy", "negative-space"],
    blueprint: {
      summary: "A concise title above a left explanation and right evidence field.",
      zones: [
        { id: "title-zone", role: "title", x: .04, y: .04, width: .92, height: .14 },
        { id: "copy-zone", role: "supporting-evidence", x: .05, y: .24, width: .38, height: .62 },
        { id: "visual-zone", role: "primary-visual", x: .49, y: .22, width: .46, height: .64 },
      ],
      styleNotes: ["Balanced ORNL expression."], reconstructionNotes: ["Use exact source content."],
    },
  });
  const result = reconstructStudioConcept(attached, 2, "concept-layout", "ornl-title-two-column");
  const slide = result.scene.slides[0];
  assert.equal(slide.visualNeeds?.[0].status, "reconstruction-ready");
  assert.equal(result.mappedNodeIds.length, 3);
  assert.equal(analyzeStudioDesignImpact(slide).requirements.every((item) => item.passed), true);
  assert.notDeepEqual(slide.nodes.find((item) => item.id === "body")?.frame, slide.nodes.find((item) => item.id === "body")?.sourceFrame);
});

test("concept reconstruction refuses a concept with no semantic zones", () => {
  const created = createStudioVisualNeed(scene(), 2, { id: "need-layout", type: "layout-concept", reason: "weak", communicationJob: "clarify" });
  const attached = attachStudioConceptReference(created, 2, conceptResource, { id: "concept-empty", visualNeedId: "need-layout", origin: "imagegen", approvedInfluences: ["composition"], blueprint: { summary: "No zone analysis yet.", zones: [], styleNotes: [], reconstructionNotes: [] } });
  assert.throws(() => reconstructStudioConcept(attached, 2, "concept-empty"), /semantic zone/i);
});

