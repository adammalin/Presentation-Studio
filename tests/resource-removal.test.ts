import assert from "node:assert/strict";
import test from "node:test";
import type { DeckJob, ProjectResource, StudioWebScene } from "../src/types";
import { createProject } from "../src/lib/project";
import { removeResourceFromProject, resourceRemovalImpact } from "../src/lib/resource-removal";

function resource(id: string, name: string): ProjectResource {
  return { id, name, mediaType: "application/octet-stream", byteLength: 1, sha256: id.padEnd(64, "0").slice(0, 64), roles: ["grounding-source"], createdAt: "2026-08-13T12:00:00.000Z", embedded: true, bytes: new Uint8Array([1]), mcpAccess: "none" };
}

test("removing an ordinary Resource changes only the self-contained project", () => {
  const project = createProject("Removal test");
  project.resources = [resource("a", "notes.txt"), resource("b", "image.png")];
  const result = removeResourceFromProject(project, "a");
  assert.deepEqual(result.project.resources.map((item) => item.id), ["b"]);
  assert.equal(result.project.decks.length, 0);
  assert.equal(project.resources.length, 2);
  assert.match(result.project.activity.at(-1)?.detail ?? "", /No external source file was changed or deleted/);
});

test("removing a deck source also removes dependent project state without touching unrelated work", () => {
  const project = createProject("Linked removal test");
  project.resources = [resource("deck-source", "source.pptx"), resource("notes", "notes.txt")];
  const linkedDeck: DeckJob = { id: "deck-one", name: "source.pptx", sourceResourceId: "deck-source", sourceSha256: "d".repeat(64), operationScope: "reflow", templateClassification: "custom", status: "audited", protectedSlideNumbers: [] };
  const otherDeck: DeckJob = { ...linkedDeck, id: "deck-two", name: "other.pptx", sourceResourceId: "notes" };
  project.decks = [linkedDeck, otherDeck];
  project.styleExemplars = [{ id: "example", name: "Table", kind: "table", resourceId: "deck-source", deckId: "deck-one", slideNumber: 1, objectOrdinal: 1, scope: "deck", createdAt: "2026-08-13T12:00:00.000Z" }];
  project.designThreads = [{ id: "thread", deckId: "deck-one", slideId: "slide-1", slideNumber: 1, baseRevision: "2026-08-13T12:00:00.000Z", anchor: { kind: "region", x: 0, y: 0, width: .1, height: .1 }, comment: "Adjust table", status: "submitted", createdAt: "2026-08-13T12:00:00.000Z", updatedAt: "2026-08-13T12:00:00.000Z" }];
  const impact = resourceRemovalImpact(project, "deck-source");
  assert.deepEqual(impact.linkedDeckIds, ["deck-one"]);
  assert.equal(impact.removedExemplarCount, 1);
  assert.equal(impact.removedThreadCount, 1);
  const result = removeResourceFromProject(project, "deck-source");
  assert.deepEqual(result.project.resources.map((item) => item.id), ["notes"]);
  assert.deepEqual(result.project.decks.map((deck) => deck.id), ["deck-two"]);
  assert.equal(result.project.styleExemplars.length, 0);
  assert.equal(result.project.designThreads.length, 0);
});

test("removing a concept Resource reopens its visual need without deleting the slide", () => {
  const project = createProject("Concept removal test");
  const source = resource("deck-source", "source.pptx");
  const concept = { ...resource("concept", "concept.png"), kind: "image" as const, mediaType: "image/png", mcpAccess: "preview" as const, roles: ["concept-reference" as const] };
  project.resources = [source, concept];
  const studioScene: StudioWebScene = {
    schema: "presentation-studio/web-scene", version: 5, revision: "scene", deckId: "deck", sourceSha256: "d".repeat(64), slideSize: { width: 12_192_000, height: 6_858_000 }, sourceSlideSize: { width: 12_192_000, height: 6_858_000 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{
      id: "slide-2", slideNumber: 2, sourceSlideId: "source-2", sourceTextHash: "b".repeat(64), contentCoverage: { exactTextMapped: true, sourceCharacterCount: 0, mappedCharacterCount: 0, sourceTextBoxCount: 0, mappedTextNodeCount: 0, groupedOrUnsupportedTextPresent: false }, sourceRevision: "source", recipe: "ornl-title-content", background: "#fff", status: "designed", designRationale: "test", figureTreatments: [], nodes: [], updatedAt: "2026-08-14T12:00:00.000Z",
      conceptReferences: [{ id: "reference", resourceId: concept.id, resourceSha256: concept.sha256, sourceTextHash: "b".repeat(64), status: "concept-only", origin: "imagegen", approvedInfluences: ["composition"], untrustedElements: ["generated-text", "generated-logos", "generated-data", "generated-technical-details", "generated-claims"], blueprint: { summary: "test", zones: [], styleNotes: [], reconstructionNotes: [] }, visualNeedId: "need", attachedAt: "2026-08-14T12:00:00.000Z" }],
      visualNeeds: [{ id: "need", type: "layout-concept", status: "concept-attached", sourceTextHash: "b".repeat(64), reason: "test", communicationJob: "test", expression: "balanced", approvedInfluences: ["composition"], disclosurePolicy: "abstract-structure-only", brandExpression: { motif: "modular-square-grid", accent: "Aqua", accentRole: "focused", typographyStrategy: "no-generated-type-reserve-editable-aptos-zones", rationale: "test" }, structureInventory: { titleCount: 0, textGroupCount: 0, imageCount: 0, tableCount: 0, figureCount: 0, calloutCount: 0 }, targetSlot: { role: "whole-slide", aspectRatio: "16:9", placementNotes: "test" }, promptPackage: { prompt: "test", negativePrompt: "test", contentSafety: "test" }, linkedConceptReferenceId: "reference", createdAt: "2026-08-14T12:00:00.000Z", updatedAt: "2026-08-14T12:00:00.000Z" }],
    }],
  };
  project.decks = [{ id: "deck", name: "source.pptx", sourceResourceId: source.id, sourceSha256: source.sha256, operationScope: "reflow", templateClassification: "custom", status: "ready-for-cleanup", protectedSlideNumbers: [], studioScene }];
  const result = removeResourceFromProject(project, concept.id);
  const slide = result.project.decks[0].studioScene?.slides[0];
  assert.deepEqual(slide?.conceptReferences, []);
  assert.equal(slide?.visualNeeds?.[0].status, "brief-ready");
  assert.equal(slide?.visualNeeds?.[0].linkedConceptReferenceId, undefined);
  assert.equal(result.project.decks.length, 1);
});
