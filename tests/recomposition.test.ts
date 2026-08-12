import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { deriveLayoutSemantics } from "../src/lib/layout-semantics";
import { auditPptx } from "../src/lib/pptx-audit";
import { semanticRecompositionRequests } from "../src/lib/recomposition";
import { compilePresentationScene } from "../src/lib/scene-graph";
import type { TemplateLayoutPreview } from "../src/lib/template-catalog";
import type { DeckJob } from "../src/types";

async function fixtureDeck(): Promise<DeckJob> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-recompose-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const audit = await auditPptx(new Uint8Array(await fs.readFile(filePath)));
  const deck: DeckJob = { id: "deck-recompose", name: "synthetic.pptx", sourceResourceId: "resource-recompose", sourceSha256: "e".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  return deck;
}

function layout(): TemplateLayoutPreview {
  const base: TemplateLayoutPreview = {
    id: "layout-semantic",
    name: "Title and Content",
    category: "content",
    background: "#FFFFFF",
    sourcePart: "ppt/slideLayouts/slideLayout2.xml",
    placeholderTypes: ["title", "body"],
    elements: [
      { id: "title", kind: "text", name: "Title", x: 914_400, y: 457_200, width: 10_363_200, height: 914_400, rotation: 0, geometry: "rect", placeholderType: "title" },
      { id: "body", kind: "text", name: "Body", x: 914_400, y: 1_600_200, width: 10_363_200, height: 4_572_000, rotation: 0, geometry: "rect", placeholderType: "body" },
    ],
  };
  return { ...base, semantic: deriveLayoutSemantics(base, 12_192_000, 6_858_000) };
}

test("semantic recomposition maps source objects into approved unique slots", async () => {
  const deck = await fixtureDeck();
  const targetLayout = layout();
  const titleObject = deck.scene?.objects.find((object) => object.slideNumber === 1 && object.semanticRole === "title");
  const bodyObject = deck.scene?.objects.find((object) => object.slideNumber === 1 && object.kind === "text" && object.id !== titleObject?.id);
  const titleSlot = targetLayout.semantic?.slots.find((slot) => slot.role === "title");
  const bodySlot = targetLayout.semantic?.slots.find((slot) => slot.role === "body");
  assert.ok(titleObject && bodyObject && titleSlot && bodySlot);
  const result = semanticRecompositionRequests({ deck, slideNumber: 1, layout: targetLayout, rationale: "Establish approved title and body hierarchy.", bindings: [{ objectId: titleObject.id, slotId: titleSlot.id, fit: "align-horizontal" }, { objectId: bodyObject.id, slotId: bodySlot.id, fit: "align-horizontal", insetInches: .05 }] });
  assert.equal(result.requests.length, 2);
  assert.equal(result.requests[0].target.x, titleSlot.x);
  assert.equal(result.requests[1].target.x, bodySlot.x + Math.round(.05 * 914_400));
  assert.equal(result.requests[0].target.y, titleObject.geometry.y);
  assert.equal(result.requests[1].target.height, bodyObject.geometry.height);
  assert.ok(result.unboundObjectIds.length > 0);
  assert.throws(() => semanticRecompositionRequests({ deck, slideNumber: 1, layout: targetLayout, rationale: "Invalid duplicate slot.", bindings: [{ objectId: titleObject.id, slotId: titleSlot.id, fit: "fill" }, { objectId: bodyObject.id, slotId: titleSlot.id, fit: "fill" }] }), /only one source object/i);
});

test("semantic recomposition rejects substantial unmeasured text-frame replacement", async () => {
  const deck = await fixtureDeck();
  const targetLayout = layout();
  const bodyObject = deck.scene?.objects.find((object) => object.slideNumber === 1 && object.kind === "text" && object.semanticRole !== "title");
  const bodySlot = targetLayout.semantic?.slots.find((slot) => slot.role === "body");
  assert.ok(bodyObject && bodySlot);
  const compressedLayout = {
    ...targetLayout,
    semantic: {
      ...targetLayout.semantic!,
      slots: targetLayout.semantic!.slots.map((slot) => slot.id === bodySlot.id ? { ...slot, y: slot.y + 914_400, height: Math.round(slot.height * .5) } : slot),
    },
  };
  assert.throws(() => semanticRecompositionRequests({ deck, slideNumber: 1, layout: compressedLayout, rationale: "Attempt unsafe replacement.", bindings: [{ objectId: bodyObject.id, slotId: bodySlot.id, fit: "fill" }] }), /substantially replace a proven text frame/i);
});
