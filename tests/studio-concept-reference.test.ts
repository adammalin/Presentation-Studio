import assert from "node:assert/strict";
import test from "node:test";
import { attachStudioConceptReference, removeStudioConceptReference } from "../src/lib/studio-concept-reference";
import type { ProjectResource, StudioWebScene } from "../src/types";

function scene(): StudioWebScene {
  return {
    schema: "presentation-studio/web-scene",
    version: 5,
    revision: "source:web-v5:before",
    deckId: "deck",
    sourceSha256: "a".repeat(64),
    slideSize: { width: 12_192_000, height: 6_858_000 },
    sourceSlideSize: { width: 12_192_000, height: 6_858_000 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{
      id: "slide-2", slideNumber: 2, sourceSlideId: "source-2", sourceTextHash: "b".repeat(64), sourceRevision: "source", recipe: "ornl-title-content", background: "#FFFFFF", status: "designed", designRationale: "test", figureTreatments: [],
      contentCoverage: { exactTextMapped: true, sourceCharacterCount: 4, mappedCharacterCount: 4, sourceTextBoxCount: 1, mappedTextNodeCount: 1, groupedOrUnsupportedTextPresent: false },
      qualityReview: { sceneRevision: "old", slideUpdatedAt: "2026-08-14T12:00:00.000Z", rasterSha256: "c".repeat(64), pass: 1, maxPasses: 3, requestedVerdict: "ready", recordedVerdict: "ready", rationale: "old", objectiveIssues: [], visualIssues: [], recordedAt: "2026-08-14T12:00:00.000Z" },
      nodes: [], updatedAt: "2026-08-14T12:00:00.000Z",
    }],
  };
}

const conceptResource: ProjectResource = {
  id: "concept-image", name: "concept.png", mediaType: "image/png", byteLength: 3, sha256: "d".repeat(64), roles: ["reference-only"], kind: "image", support: ["previewable", "placeable"], createdAt: "2026-08-14T12:00:00.000Z", embedded: true, bytes: new Uint8Array([1, 2, 3]), mcpAccess: "preview",
};

test("concept references preserve source authority and record only approved visual influence", () => {
  const attached = attachStudioConceptReference(scene(), 2, conceptResource, {
    origin: "imagegen",
    approvedInfluences: ["composition", "negative-space", "figure-concept"],
    blueprint: {
      summary: "Use one large technical visual and one concise evidence column.",
      zones: [{ id: "visual", role: "primary-visual", x: .05, y: .2, width: .58, height: .65 }],
      styleNotes: ["Use green as an anchor, not a full-slide fill."],
      reconstructionNotes: ["Replace every generated label with exact source content."],
    },
    provenance: { model: "example-image-model", promptSummary: "ORNL technical slide concept", generatedAt: "2026-08-14T12:00:00.000Z" },
  });
  const slide = attached.slides[0];
  const reference = slide.conceptReferences?.[0];
  assert.equal(reference?.status, "concept-only");
  assert.equal(reference?.sourceTextHash, "b".repeat(64));
  assert.deepEqual(reference?.approvedInfluences, ["composition", "negative-space", "figure-concept"]);
  assert.deepEqual(reference?.untrustedElements, ["generated-text", "generated-logos", "generated-data", "generated-technical-details", "generated-claims"]);
  assert.equal(slide.qualityReview, undefined);
  assert.notEqual(attached.revision, "source:web-v5:before");

  const removed = removeStudioConceptReference(attached, 2, reference!.id);
  assert.deepEqual(removed.slides[0].conceptReferences, []);
  assert.equal(conceptResource.bytes?.byteLength, 3);
});

test("concept references reject non-image Resources and invalid normalized zones", () => {
  assert.throws(() => attachStudioConceptReference(scene(), 2, { ...conceptResource, kind: "document", mediaType: "application/pdf" }, { origin: "other", approvedInfluences: ["composition"], blueprint: { summary: "test", zones: [], styleNotes: [], reconstructionNotes: [] } }), /image Resource/i);
  assert.throws(() => attachStudioConceptReference(scene(), 2, conceptResource, { origin: "imagegen", approvedInfluences: ["composition"], blueprint: { summary: "test", zones: [{ id: "bad", role: "other", x: .8, y: 0, width: .4, height: .4 }], styleNotes: [], reconstructionNotes: [] } }), /normalized 0-1 geometry/i);
});
