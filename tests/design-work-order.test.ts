import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { buildDeckDesignWorkOrder, buildSlideDesignWorkOrder, selectRepresentativeSlides } from "../src/lib/design-work-order";
import { deriveLayoutSemantics } from "../src/lib/layout-semantics";
import { auditPptx } from "../src/lib/pptx-audit";
import { compilePresentationScene } from "../src/lib/scene-graph";
import type { TemplateCatalog, TemplateLayoutPreview } from "../src/lib/template-catalog";
import type { DeckJob } from "../src/types";

async function fixtureDeck(): Promise<DeckJob> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-work-order-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const audit = await auditPptx(new Uint8Array(await fs.readFile(filePath)));
  const deck: DeckJob = { id: "deck-work-order", name: "synthetic.pptx", sourceResourceId: "resource-work-order", sourceSha256: "a".repeat(64), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  return deck;
}

function catalog(): TemplateCatalog {
  const make = (id: string, name: string, category: TemplateLayoutPreview["category"], elements: TemplateLayoutPreview["elements"]): TemplateLayoutPreview => {
    const base = { id, name, category, background: "#FFFFFF", elements, placeholderTypes: elements.flatMap((element) => element.placeholderType ? [element.placeholderType] : []), sourcePart: `ppt/slideLayouts/${id}.xml` };
    return { ...base, semantic: deriveLayoutSemantics(base, 12_192_000, 6_858_000) };
  };
  const title = { id: "title", kind: "text" as const, name: "Title", x: 609_600, y: 457_200, width: 10_972_800, height: 914_400, rotation: 0, geometry: "rect" as const, placeholderType: "title" };
  return {
    id: "template-test",
    name: "Synthetic ORNL-compatible template",
    sha256: "d".repeat(64),
    slideWidth: 12_192_000,
    slideHeight: 6_858_000,
    masterCount: 1,
    layouts: [
      make("layout-cover", "Title Slide", "title", [title, { ...title, id: "subtitle", name: "Subtitle", y: 1_828_800, height: 1_371_600, placeholderType: "subTitle" }]),
      make("layout-content", "Title and Content", "content", [title, { ...title, id: "body", name: "Body", y: 1_600_200, height: 4_572_000, placeholderType: "body" }]),
      make("layout-table", "Title and Table", "content", [title, { ...title, id: "table", name: "Table", y: 1_600_200, height: 4_572_000, placeholderType: "tbl" }]),
    ],
    media: {},
    generatedAt: "2026-08-12T00:00:00.000Z",
  };
}

test("slide design work order closes content and binds evidence to scene, template, and render revisions", async () => {
  const deck = await fixtureDeck();
  const updatedAt = "2026-08-12T20:00:00.000Z";
  const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber: 1, projectUpdatedAt: updatedAt, templateCatalog: catalog() });
  assert.equal(workOrder.schema, "presentation-studio/design-work-order");
  assert.equal(workOrder.version, 1);
  assert.match(workOrder.revision, /render-unavailable$/);
  assert.equal(workOrder.slide.exactVisibleText, deck.audit?.slides[0].text);
  assert.equal(workOrder.closedContentInventory.lockedTextHash, deck.audit?.slides[0].textHash);
  assert.equal(workOrder.objects.length, deck.scene?.objects.filter((object) => object.slideNumber === 1).length);
  assert.equal(workOrder.layoutCandidates[0].layout.semantic?.intent, "cover");
  assert.match(workOrder.requiredSequence.join(" "), /PowerPoint/i);
  assert.match(workOrder.definitionOfDone, /visibly stronger/i);
});

test("slide design work order exposes optical text starts instead of only shape coordinates", async () => {
  const deck = await fixtureDeck();
  const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber: 2, projectUpdatedAt: "2026-08-12T20:00:00.000Z", templateCatalog: catalog() });
  const bulletObject = workOrder.objects.find((object) => object.exactText?.startsWith("This deliberately dense"));
  assert.ok(bulletObject?.opticalText);
  assert.equal(bulletObject.opticalText.bulletParagraphCount, 1);
  assert.ok(bulletObject.opticalText.estimatedOpticalLeftEmu > bulletObject.geometry.x);
  assert.match(bulletObject.opticalText.instruction, /visible text starts/i);
});

test("deck work order selects a bounded representative qualification set", async () => {
  const deck = await fixtureDeck();
  const representatives = selectRepresentativeSlides(deck);
  assert.equal(representatives[0].role, "cover");
  assert.ok(representatives.some((item) => item.role === "table"));
  assert.equal(new Set(representatives.map((item) => item.slideNumber)).size, representatives.length);
  assert.ok(representatives.length <= 5);
  const workOrder = buildDeckDesignWorkOrder({ deck, projectUpdatedAt: "2026-08-12T20:00:00.000Z", templateCatalog: catalog() });
  assert.equal(workOrder.workOrders.length, representatives.length);
  assert.match(workOrder.executionPolicy, /representative set/i);
});
