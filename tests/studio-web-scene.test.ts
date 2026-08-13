import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { auditPptx } from "../src/lib/pptx-audit";
import { createProject, projectSchema } from "../src/lib/project";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { buildSlideRenderCatalog } from "../src/lib/template-catalog";
import { compileStudioWebScene, recomposeStudioWebSlide, studioGeometryRequests, studioSceneNeedsRebuild, studioVisualDesignRequest, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "../src/lib/studio-web-scene";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createVisualDesignProposal } from "../src/lib/cleanup";
import { sha256 } from "../src/lib/hash";
import type { DeckJob } from "../src/types";

async function fixture() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-web-scene-"));
  const filePath = path.join(directory, "synthetic.pptx");
  await createSyntheticLegacyDeck(filePath);
  const bytes = new Uint8Array(await fs.readFile(filePath));
  const audit = await auditPptx(bytes);
  const deck: DeckJob = { id: "studio-web-deck", name: "synthetic.pptx", sourceResourceId: "studio-web-resource", sourceSha256: await sha256(bytes), operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: "2026-08-13T16:00:00.000Z", status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
  deck.scene = compilePresentationScene({ ...deck, audit });
  const catalog = await buildSlideRenderCatalog(bytes, deck.name);
  return { bytes, deck, catalog };
}

test("Studio Web Scene separates exact PowerPoint content from redesignable web geometry", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  assert.equal(scene.designSystem.renderer, "html-css");
  assert.equal(scene.designSystem.exportTarget, "editable-powerpoint");
  assert.equal(scene.slides.length, deck.audit?.slideCount);
  assert.equal(scene.slides.every((slide) => slide.sourceTextHash === deck.audit?.slides.find((item) => item.number === slide.slideNumber)?.textHash), true);
  assert.equal(scene.slides.flatMap((slide) => slide.nodes).every((node) => node.sourceObjectId && node.sourceShapeId && node.style.fontFamily === "Aptos"), true);
  const sourceText = deck.audit?.textBoxes.map((item) => item.text).filter(Boolean).sort();
  const sceneText = scene.slides.flatMap((slide) => slide.nodes.map((node) => node.text).filter((value): value is string => Boolean(value))).sort();
  assert.deepEqual(sceneText, sourceText);
  assert.equal(scene.slides.some((slide) => slide.nodes.some((node) => node.kind === "table" && node.table?.cells.some((cell) => cell.text))), true);
});

test("shared ORNL web recipes recompose complete slides and compile back to source-bound PowerPoint commands", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const slideNumber = source.slides.find((slide) => slide.nodes.some((node) => node.kind === "table"))?.slideNumber ?? 1;
  const designed = recomposeStudioWebSlide(source, slideNumber, "ornl-title-table", undefined, "Use the shared ORNL table composition.");
  const slide = designed.slides.find((item) => item.slideNumber === slideNumber);
  assert.equal(slide?.status, "designed");
  assert.equal(slide?.recipe, "ornl-title-table");
  assert.equal(slide?.nodes.map((node) => node.text).filter(Boolean).sort().join("\n"), source.slides.find((item) => item.slideNumber === slideNumber)?.nodes.map((node) => node.text).filter(Boolean).sort().join("\n"));
  const geometry = studioGeometryRequests(deck, designed, slideNumber);
  assert.ok(geometry.length > 0);
  assert.equal(geometry.every((request) => deck.audit?.editableObjects.some((object) => object.id === request.objectId)), true);
  const visual = studioVisualDesignRequest(designed, slideNumber);
  assert.equal(visual.slideNumber, slideNumber);
  assert.equal(visual.decorations[0]?.name, "ORNL title rule");
  assert.equal(visual.textStyles.every((style) => style.fontSizePt && style.fontSizePt >= 14), true);
});

test("a Studio web composition round-trips to editable PowerPoint without changing exact content", async () => {
  const { bytes, deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const slideNumber = source.slides.find((slide) => slide.nodes.some((node) => node.kind === "table"))?.slideNumber ?? 1;
  const designed = recomposeStudioWebSlide(source, slideNumber, "ornl-title-table", undefined, "Use the shared ORNL table page and preserve the technical content exactly.");
  const revision = "2026-08-13T16:00:00.000Z";
  const geometryProposal = createGeometryBatchProposal(deck, revision, studioGeometryRequests(deck, designed, slideNumber));
  const proposal = createVisualDesignProposal({ ...deck, proposal: geometryProposal }, revision, studioVisualDesignRequest(designed, slideNumber));
  const compiled = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(compiled.bytes);
  assert.ok(compiled.geometryCount > 0);
  assert.ok(compiled.textStyleCount > 0);
  assert.equal(compiled.decorationCount, 1);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), deck.audit?.slides.map((slide) => slide.textHash));
  assert.deepEqual(after.tables.map((table) => table.contentHash), deck.audit?.tables.map((table) => table.contentHash));
  assert.deepEqual(after.tables.map((table) => table.structureHash), deck.audit?.tables.map((table) => table.structureHash));
  const sourceTable = deck.audit?.editableObjects.find((object) => object.slideNumber === slideNumber && object.kind === "table");
  const outputTable = after.editableObjects.find((object) => object.id === sourceTable?.id);
  assert.ok(sourceTable && outputTable);
  assert.notDeepEqual(outputTable.geometry, sourceTable.geometry);
});

test("human canvas edits remain bounded and the self-contained project persists the web scene", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const node = scene.slides[0].nodes.find((item) => !item.locked);
  assert.ok(node);
  const updated = updateStudioWebNodeFrame(scene, 1, node.id, { x: -10_000, y: -20_000, width: scene.slideSize.width * 2, height: node.frame.height, rotation: 0 });
  const updatedNode = updated.slides[0].nodes.find((item) => item.id === node.id);
  assert.equal(updatedNode?.frame.x, 0);
  assert.equal(updatedNode?.frame.y, 0);
  assert.equal(updatedNode?.frame.width, scene.slideSize.width);
  const styled = updateStudioWebNodeStyle(updated, 1, node.id, { fontSizePt: 27, fontWeight: 600, color: "#00662C", textAlign: "left" });
  const styledNode = styled.slides[0].nodes.find((item) => item.id === node.id);
  assert.equal(styledNode?.style.fontSizePt, 27);
  assert.equal(styledNode?.style.fontFamily, "Aptos");
  deck.studioScene = styled;
  const project = createProject("Studio web scene persistence");
  project.decks = [deck];
  const parsed = projectSchema.parse(project);
  assert.equal(parsed.decks[0].studioScene?.revision, styled.revision);
  assert.equal(studioSceneNeedsRebuild(parsed.decks[0]), false);
});
