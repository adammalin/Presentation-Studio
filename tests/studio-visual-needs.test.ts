import assert from "node:assert/strict";
import test from "node:test";
import { attachStudioConceptReference, removeStudioConceptReference } from "../src/lib/studio-concept-reference";
import { createStudioVisualNeed, holdStudioVisualNeed, markStudioVisualNeedsReconstructionReady, resolveStudioVisualNeeds } from "../src/lib/studio-visual-needs";
import type { ProjectResource, StudioWebNode, StudioWebScene } from "../src/types";

function node(id: string, kind: StudioWebNode["kind"], role: StudioWebNode["role"], name = id): StudioWebNode {
  return {
    id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "editable-object", name, kind, role,
    sourceFrame: { x: 0, y: 0, width: 1_000_000, height: 500_000, rotation: 0 },
    frame: { x: 0, y: 0, width: 1_000_000, height: 500_000, rotation: 0 },
    zIndex: 1, sourceTextOrder: 1, visible: true, locked: false, exactContent: true,
    text: kind === "text" ? "classified source wording" : undefined,
    style: { fontFamily: "Aptos", fontSizePt: 18, fontWeight: 400, lineHeight: 1.1, color: "#111111", borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } },
  };
}

function scene(): StudioWebScene {
  return {
    schema: "presentation-studio/web-scene", version: 5, revision: "source:web-v5:before", deckId: "deck", sourceSha256: "a".repeat(64),
    slideSize: { width: 12_192_000, height: 6_858_000 }, sourceSlideSize: { width: 12_192_000, height: 6_858_000 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{
      id: "slide-2", slideNumber: 2, sourceSlideId: "source-2", sourceTextHash: "b".repeat(64), sourceRevision: "source", recipe: "ornl-title-content", background: "#FFFFFF", status: "designed", designRationale: "test", figureTreatments: [], conceptReferences: [], visualNeeds: [],
      contentCoverage: { exactTextMapped: true, sourceCharacterCount: 26, mappedCharacterCount: 26, sourceTextBoxCount: 2, mappedTextNodeCount: 2, groupedOrUnsupportedTextPresent: false },
      nodes: [node("title", "text", "title"), node("body", "text", "body"), node("image", "image", "image"), node("arrow", "connector", "connector", "callout arrow")],
      updatedAt: "2026-08-14T12:00:00.000Z",
    }],
  };
}

const conceptResource: ProjectResource = {
  id: "concept-image", name: "concept.png", mediaType: "image/png", byteLength: 3, sha256: "d".repeat(64), roles: ["concept-reference"], kind: "image", support: ["previewable", "placeable"], createdAt: "2026-08-14T12:00:00.000Z", embedded: true, bytes: new Uint8Array([1, 2, 3]), mcpAccess: "preview",
};

test("visual-needs briefs expose structure without leaking source wording by default", () => {
  const created = createStudioVisualNeed(scene(), 2, {
    id: "need-layout",
    type: "layout-concept",
    reason: "The classified source wording has no visual hierarchy.",
    communicationJob: "Explain classified source wording clearly.",
    expression: "balanced",
  });
  const need = created.slides[0].visualNeeds?.[0];
  assert.equal(need?.status, "brief-ready");
  assert.equal(need?.disclosurePolicy, "abstract-structure-only");
  assert.equal(need?.structureInventory.titleCount, 1);
  assert.equal(need?.structureInventory.imageCount, 1);
  assert.equal(need?.structureInventory.calloutCount, 1);
  assert.doesNotMatch(need?.promptPackage.prompt ?? "", /classified source wording/i);
  assert.match(need?.promptPackage.prompt ?? "", /concept-only art-direction/i);
  assert.match(need?.promptPackage.prompt ?? "", /PRIMARY ARTIFACT COUNT: 1.*ONLY ALLOWED TEXT: NONE.*LOGO: OMIT/i);
  assert.equal(need?.brandExpression.typographyStrategy, "no-generated-type-reserve-editable-aptos-zones");
  assert.match(need?.promptPackage.prompt ?? "", /square 90-degree corners/i);
  assert.match(need?.promptPackage.negativePrompt ?? "", /No readable text.*logos.*data/i);
  assert.match(need?.promptPackage.negativePrompt ?? "", /generic green-and-navy panel flooding/i);
});

test("exact-content disclosure requires an explicitly bounded approved summary", () => {
  assert.throws(() => createStudioVisualNeed(scene(), 2, { type: "figure-concept", reason: "weak", communicationJob: "show a process", disclosurePolicy: "exact-content-approved" }), /approved content summary/i);
  assert.throws(() => createStudioVisualNeed(scene(), 2, { type: "figure-concept", reason: "weak", communicationJob: "show a process", approvedContentSummary: "approved" }), /requires exact-content-approved/i);
  const created = createStudioVisualNeed(scene(), 2, { id: "need-figure", type: "figure-concept", reason: "weak", communicationJob: "show a process", disclosurePolicy: "exact-content-approved", approvedContentSummary: "Three generic stages connected left to right." });
  assert.match(created.slides[0].visualNeeds?.[0].promptPackage.prompt ?? "", /Three generic stages connected left to right/);
});

test("concept linking, editable reconstruction, native-ready resolution, and detachment use an explicit lifecycle", () => {
  const created = createStudioVisualNeed(scene(), 2, { id: "need-figure", type: "figure-concept", reason: "The source figure is hard to scan.", communicationJob: "Clarify the relationship without changing its meaning." });
  const attached = attachStudioConceptReference(created, 2, conceptResource, {
    id: "concept-figure",
    visualNeedId: "need-figure",
    origin: "imagegen",
    approvedInfluences: ["figure-concept", "negative-space"],
    blueprint: { summary: "One evidence unit beside one explanation zone.", zones: [], styleNotes: [], reconstructionNotes: [] },
  });
  assert.equal(attached.slides[0].visualNeeds?.[0].status, "concept-attached");
  assert.equal(attached.slides[0].visualNeeds?.[0].linkedConceptReferenceId, "concept-figure");
  assert.equal(attached.slides[0].conceptReferences?.[0].visualNeedId, "need-figure");

  const designed = {
    ...attached,
    slides: attached.slides.map((slide) => ({
      ...slide,
      nodes: slide.nodes.map((item) => item.id === "image" ? { ...item, frame: { ...item.frame, x: item.frame.x + 100_000 } } : item),
      figureTreatments: [{ id: "figure-treatment", nodeIds: ["image"], mode: "preserve-and-frame" as const, verificationStatus: "source-locked" as const, intentSummary: "Preserve the source evidence.", informationInventory: ["source image"], invariants: ["all source pixels"], rationale: "Give the evidence a deliberate frame.", relationshipPolicy: "preserve-internal" as const }],
    })),
  };
  const reconstructed = markStudioVisualNeedsReconstructionReady(designed, 2);
  assert.equal(reconstructed.slides[0].visualNeeds?.[0].status, "reconstruction-ready");
  const resolved = resolveStudioVisualNeeds(reconstructed, 2, "Editable reconstruction passed native PowerPoint review.");
  assert.equal(resolved.slides[0].visualNeeds?.[0].status, "resolved");

  const detached = removeStudioConceptReference(attached, 2, "concept-figure");
  assert.equal(detached.slides[0].visualNeeds?.[0].status, "brief-ready");
  assert.equal(detached.slides[0].visualNeeds?.[0].linkedConceptReferenceId, undefined);
});

test("a held visual need remains visible without deleting its brief", () => {
  const created = createStudioVisualNeed(scene(), 2, { id: "need-hold", type: "supporting-visual", reason: "Optional supporting art may help.", communicationJob: "Support the main evidence." });
  const held = holdStudioVisualNeed(created, 2, "need-hold", "No concept is needed for this version.");
  assert.equal(held.slides[0].visualNeeds?.[0].status, "held");
  assert.match(held.slides[0].visualNeeds?.[0].resolutionNote ?? "", /No concept is needed/i);
});
