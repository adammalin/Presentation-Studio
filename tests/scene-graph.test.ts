import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { auditPptx } from "../src/lib/pptx-audit";
import { createProject, projectSchema } from "../src/lib/project";
import { compilePresentationScene, sceneNeedsRebuild } from "../src/lib/scene-graph";
import type { DeckJob, PptxAudit, SlideEditableObject } from "../src/types";

async function fixtureAudit(): Promise<PptxAudit> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-scene-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  return auditPptx(new Uint8Array(await fs.readFile(filePath)));
}

function deckForAudit(audit: PptxAudit): DeckJob {
  return {
    id: "deck-scene",
    name: "synthetic.pptx",
    sourceResourceId: "resource-scene",
    sourceSha256: "a".repeat(64),
    operationScope: "reflow",
    templateClassification: audit.classification,
    targetTemplateId: "ornl-16x9-v1",
    targetTemplateDecisionSource: "automatic-default",
    status: "ready-for-cleanup",
    audit,
    protectedSlideNumbers: [],
  };
}

test("hybrid scene binds every audited object to hashed native slide parts", async () => {
  const audit = await fixtureAudit();
  const deck = deckForAudit(audit);
  const scene = compilePresentationScene({ ...deck, audit });
  assert.equal(scene.slides.length, audit.slideCount);
  assert.equal(scene.objects.length, audit.editableObjects.length);
  assert.equal(scene.preservationEnvelope.sourceSha256, deck.sourceSha256);
  assert.equal(scene.preservationEnvelope.sourceBytesAuthoritative, true);
  assert.equal(scene.preservationEnvelope.nativeRenderAuthoritativeForAppearance, true);
  assert.equal(scene.preservationEnvelope.exportStrategy, "surgical-ooxml-overlay");
  assert.equal(scene.slides.every((slide) => /^[0-9a-f]{64}$/.test(slide.sourcePartSha256 ?? "")), true);
  assert.equal(scene.slides.every((slide) => slide.objectIds.every((id) => scene.objects.some((object) => object.id === id))), true);
  assert.equal(scene.objects.every((object) => object.sourceLocator.slidePart === `ppt/slides/slide${object.slideNumber}.xml`), true);
  const table = scene.objects.find((object) => object.kind === "table");
  assert.ok(table);
  assert.equal(table.fidelityState, "editable-native");
  assert.equal(table.operations.editTableStyle, true);
  assert.equal(table.operations.editText, false);
  assert.equal(table.representation.internalStructure, "native");
});

test("scene fidelity distinguishes preserved and conversion-only PowerPoint objects", async () => {
  const audit = await fixtureAudit();
  const base = audit.editableObjects[0];
  assert.ok(base);
  const picture: SlideEditableObject = { ...base, id: "slide-1-object-preserved", shapeId: "preserved", name: "Preserved picture", kind: "picture", sourceElement: "p:pic", textHash: undefined, pictureId: "slide-1-picture-99" };
  const graphicFrame: SlideEditableObject = { ...base, id: "slide-1-object-conversion", shapeId: "conversion", name: "Unsupported diagram", kind: "graphic-frame", sourceElement: "p:graphicFrame", textHash: undefined };
  const advancedAudit: PptxAudit = { ...audit, containsOleObjects: true, editableObjects: [...audit.editableObjects, picture, graphicFrame] };
  const deck = deckForAudit(advancedAudit);
  const scene = compilePresentationScene({ ...deck, audit: advancedAudit });
  assert.equal(scene.objects.find((object) => object.id === picture.id)?.fidelityState, "preserved-native");
  assert.equal(scene.objects.find((object) => object.id === graphicFrame.id)?.fidelityState, "conversion-required");
  assert.equal(scene.fidelityCounts["preserved-native"], 1);
  assert.equal(scene.fidelityCounts["conversion-required"], 1);
  assert.deepEqual(scene.preservationEnvelope.blockingFeatures, ["ole-objects"]);
});

test("project schema persists scenes while legacy projects remain rebuildable", async () => {
  const audit = await fixtureAudit();
  const deck = deckForAudit(audit);
  deck.scene = compilePresentationScene({ ...deck, audit });
  const project = createProject("Scene persistence");
  project.decks = [deck];
  const parsed = projectSchema.parse(project);
  assert.equal(parsed.decks[0].scene?.revision, deck.scene.revision);
  assert.equal(sceneNeedsRebuild(parsed.decks[0]), false);

  const legacy = structuredClone(project) as typeof project;
  delete legacy.decks[0].scene;
  const parsedLegacy = projectSchema.parse(legacy);
  assert.equal(parsedLegacy.decks[0].scene, undefined);
  assert.equal(sceneNeedsRebuild(parsedLegacy.decks[0]), true);
});
