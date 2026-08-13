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
import { atomizeStudioWebSlide, compileStudioWebScene, recommendedStudioRecipe, recomposeStudioWebSlide, studioGeneratedComponents, studioGeometryRequests, studioSceneNeedsRebuild, studioVisualDesignRequest, updateStudioFigureTreatment, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "../src/lib/studio-web-scene";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createVisualDesignProposal } from "../src/lib/cleanup";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { sha256, sha256Text } from "../src/lib/hash";
import type { DeckJob, StudioWebNode, StudioWebScene } from "../src/types";

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
  assert.deepEqual(scene.designSystem.compilerModes, ["source-bound-overlay", "fresh-composition"]);
  assert.equal(scene.slides.length, deck.audit?.slideCount);
  assert.equal(scene.slides.every((slide) => slide.sourceTextHash === deck.audit?.slides.find((item) => item.number === slide.slideNumber)?.textHash), true);
  assert.equal(scene.slides.flatMap((slide) => slide.nodes).every((node) => node.sourceObjectId && node.sourceShapeId && node.style.fontFamily === "Aptos"), true);
  const sourceText = deck.audit?.textBoxes.map((item) => item.text).filter(Boolean).sort();
  const sceneText = scene.slides.flatMap((slide) => slide.nodes.map((node) => node.text).filter((value): value is string => Boolean(value))).sort();
  assert.deepEqual(sceneText, sourceText);
  assert.equal(scene.slides.every((slide) => slide.contentCoverage.exactTextMapped), true);
  assert.equal(scene.slides.flatMap((slide) => slide.nodes).filter((node) => node.kind === "text").every((node) => (node.sourceParagraphs?.length ?? 0) > 0), true);
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

test("source-locked figure treatments preserve the technical object and add one shared ORNL frame", async () => {
  const { bytes, deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const figureNode = source.slides.flatMap((slide) => slide.nodes.map((node) => ({ slide, node }))).find(({ node }) => ["image", "native-object", "shape", "connector"].includes(node.kind));
  assert.ok(figureNode);
  const treated = updateStudioFigureTreatment(source, figureNode.slide.slideNumber, {
    id: `technical-figure-${figureNode.slide.slideNumber}`,
    nodeIds: [figureNode.node.id],
    mode: "preserve-and-frame",
    verificationStatus: "source-locked",
    intentSummary: "Preserve the original technical evidence while improving its relationship to the page grid.",
    informationInventory: ["The complete source visual"],
    invariants: ["Do not change labels, values, arrows, code, screenshots, or technical relationships."],
    rationale: "The source object carries more meaning than can safely be inferred and redrawn automatically.",
  });
  const slide = treated.slides.find((item) => item.slideNumber === figureNode.slide.slideNumber)!;
  assert.equal(slide.figureTreatments.length, 1);
  assert.equal(slide.nodes.find((node) => node.id === figureNode.node.id)?.visible, true);
  assert.equal(slide.nodes.find((node) => node.id === figureNode.node.id)?.mediaPart, figureNode.node.mediaPart);
  assert.deepEqual(studioGeneratedComponents(slide).map((component) => component.id), [`${slide.figureTreatments[0].id}-surface`, `${slide.figureTreatments[0].id}-accent`]);
  const proposal = createVisualDesignProposal(deck, "2026-08-13T17:00:00.000Z", studioVisualDesignRequest(treated, slide.slideNumber));
  const compiled = await buildCleanupProposalPptx(bytes, proposal);
  const after = await auditPptx(compiled.bytes);
  assert.equal(compiled.decorationCount, 2);
  assert.deepEqual(after.slides.map((item) => item.textHash), deck.audit?.slides.map((item) => item.textHash));
  assert.throws(() => updateStudioFigureTreatment(source, figureNode.slide.slideNumber, { ...slide.figureTreatments[0], mode: "redraw-candidate", verificationStatus: "source-locked" }), /needs-content-review/i);
});

test("comparison-card recipe ignores footer furniture and composes repeated semantic groups", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const base = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const makeText = (id: string, text: string, role: StudioWebNode["role"], zIndex: number, y: number): StudioWebNode => ({
    ...base,
    id,
    sourceObjectId: id,
    sourceShapeId: id,
    name: id,
    kind: "text",
    role,
    zIndex,
    text,
    sourceFrame: { x: 500_000, y, width: 3_000_000, height: 400_000, rotation: 0 },
    frame: { x: 500_000, y, width: 3_000_000, height: 400_000, rotation: 0 },
    locked: false,
    visible: true,
  });
  const nodes: StudioWebNode[] = [makeText("eyebrow", "SECTION", "label", 1, 200_000), makeText("title", "Four comparable systems", "title", 2, 500_000)];
  for (let index = 0; index < 4; index += 1) {
    nodes.push(makeText(`kicker-${index}`, `S${index + 1}`, "label", 3 + index * 3, 1_400_000 + index * 500_000));
    nodes.push(makeText(`heading-${index}`, `System ${index + 1}`, "label", 4 + index * 3, 1_400_000 + index * 500_000));
    nodes.push(makeText(`body-${index}`, `Exact technical explanation ${index + 1}.`, "body", 5 + index * 3, 1_800_000 + index * 500_000));
  }
  nodes.push({ ...makeText("footer", "Presentation · 20", "caption", 40, 6_500_000), sourceFrame: { x: 8_000_000, y: 6_300_000, width: 3_000_000, height: 200_000, rotation: 0 }, frame: { x: 8_000_000, y: 6_300_000, width: 3_000_000, height: 200_000, rotation: 0 } });
  const cardSource = { ...scene, slides: [{ ...sourceSlide, nodes }] };
  assert.equal(recommendedStudioRecipe(cardSource.slides[0]), "ornl-title-card-grid");
  const designed = recomposeStudioWebSlide(cardSource, sourceSlide.slideNumber);
  const slide = designed.slides[0];
  assert.equal(slide.recipe, "ornl-title-card-grid");
  assert.equal(slide.nodes.filter((node) => node.component?.role === "card-body").length, 4);
  assert.equal(new Set(slide.nodes.filter((node) => node.component?.role === "card-body").map((node) => node.component?.groupId)).size, 4);
  assert.equal(studioGeneratedComponents(slide).length, 10);
  assert.deepEqual(slide.nodes.map((node) => node.text).filter(Boolean).sort(), nodes.map((node) => node.text).filter(Boolean).sort());
});

test("semantic atomization turns one exact multi-paragraph source box into reusable objective columns", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const title = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const paragraphText = [
    "Understand the EMT simulation workflow.",
    "Build and run representative power-system models.",
    "Interpret the resulting technical evidence.",
  ];
  const paragraphs = await Promise.all(paragraphText.map(async (text, index) => ({ index: index + 1, text, textHash: await sha256Text(text), characterCount: text.length, bullet: false, bulletConfidence: "direct" as const, level: 0, fontFamilies: ["Aptos"], fontSizes: [18] })));
  const bodyText = paragraphText.join(" ");
  const body: StudioWebNode = {
    ...title,
    id: "objective-source",
    sourceObjectId: "slide-1-object-objective",
    sourceShapeId: "objective",
    sourceBinding: "editable-object",
    name: "Objectives",
    role: "body",
    zIndex: title.zIndex + 1,
    text: bodyText,
    textHash: await sha256Text(bodyText),
    sourceParagraphs: paragraphs,
    sourceFrame: { x: 500_000, y: 1_500_000, width: 11_000_000, height: 4_500_000, rotation: 0 },
    frame: { x: 500_000, y: 1_500_000, width: 11_000_000, height: 4_500_000, rotation: 0 },
  };
  const objectiveTitle: StudioWebNode = { ...title, role: "title", text: "Training Objectives", textHash: await sha256Text("Training Objectives"), sourceParagraphs: [{ ...paragraphs[0], text: "Training Objectives", textHash: await sha256Text("Training Objectives"), characterCount: 19 }] };
  const objectiveSource: StudioWebScene = { ...scene, slides: [{ ...sourceSlide, nodes: [objectiveTitle, body], contentCoverage: { exactTextMapped: true, sourceCharacterCount: 19 + 1 + bodyText.length, mappedCharacterCount: 19 + 1 + bodyText.length, sourceTextBoxCount: 2, mappedTextNodeCount: 2, groupedOrUnsupportedTextPresent: false } }] };
  assert.equal(recommendedStudioRecipe(objectiveSource.slides[0]), "ornl-title-objective-columns");
  const atomized = atomizeStudioWebSlide(objectiveSource, 1);
  assert.equal(atomized.slides[0].nodes.find((node) => node.id === body.id)?.visible, false);
  assert.deepEqual(atomized.slides[0].nodes.filter((node) => node.sourceBinding === "semantic-atom").map((node) => node.text), paragraphText);
  const designed = recomposeStudioWebSlide(objectiveSource, 1, "ornl-title-objective-columns");
  const atoms = designed.slides[0].nodes.filter((node) => node.sourceBinding === "semantic-atom");
  assert.equal(atoms.length, 3);
  assert.equal(atoms.every((node) => node.component?.role === "objective-body"), true);
  assert.equal(studioGeneratedComponents(designed.slides[0]).filter((component) => component.id.includes("objective")).length, 3);
  assert.equal(studioGeometryRequests(deck, designed, 1).some((request) => atoms.some((node) => node.sourceObjectId === request.objectId)), false);
  assert.equal(studioVisualDesignRequest(designed, 1).textStyles.some((request) => atoms.some((node) => node.sourceObjectId === request.objectId)), false);
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

test("fresh-composition mode builds a new editable native deck from the web scene", async () => {
  const { bytes, deck, catalog } = await fixture();
  const before = await auditPptx(bytes);
  let scene = compileStudioWebScene(deck, catalog);
  for (const slide of scene.slides) scene = recomposeStudioWebSlide(scene, slide.slideNumber);
  const rebuilt = await buildStudioCompositionPptx(scene, { catalog, title: "Synthetic Studio rebuild" });
  const after = await auditPptx(rebuilt.bytes);
  assert.equal(rebuilt.slideCount, before.slideCount);
  assert.ok(rebuilt.textNodeCount > 0);
  assert.equal(rebuilt.tableCount, before.tableCount);
  assert.equal(rebuilt.warnings.length, 0);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), before.slides.map((slide) => slide.textHash));
  assert.deepEqual(after.tables.map((table) => table.contentHash), before.tables.map((table) => table.contentHash));
  assert.deepEqual(after.tables.map((table) => table.structureHash), before.tables.map((table) => table.structureHash));
  assert.equal(after.fonts.some((font) => font.family === "Aptos"), true);
  assert.equal(after.textBoxes.every((textBox) => textBox.fontFamilies.every((family) => family === "Aptos")), true);
  assert.equal(after.tables.every((table) => table.cellFonts.every((family) => family === "Aptos")), true);
});

test("fresh table composition preserves an explicit visible cell break in the editable PowerPoint grid", async () => {
  const { deck, catalog } = await fixture();
  let scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides.find((slide) => slide.nodes.some((node) => node.kind === "table" && node.table));
  const sourceTable = sourceSlide?.nodes.find((node) => node.kind === "table" && node.table);
  const sourceCell = sourceTable?.table?.cells.find((cell) => cell.text.length > 0);
  assert.ok(sourceSlide && sourceTable?.table && sourceCell);
  const updatedText = "First Second";
  const updatedCell: typeof sourceCell = { ...sourceCell, text: updatedText, textRuns: ["First", "Second"], paragraphRunCounts: [2], runBreaksBefore: ["none", "line"] };
  scene = {
    ...scene,
    slides: scene.slides.map((slide) => slide.slideNumber !== sourceSlide.slideNumber ? slide : {
      ...slide,
      nodes: slide.nodes.map((node) => node.id !== sourceTable.id || !node.table ? node : { ...node, table: { ...node.table, cells: node.table.cells.map((cell) => cell.id === sourceCell.id ? updatedCell : cell) } }),
    }),
  };
  for (const slide of scene.slides) scene = recomposeStudioWebSlide(scene, slide.slideNumber);
  const rebuilt = await buildStudioCompositionPptx(scene, { catalog, title: "Table line-break round trip" });
  const after = await auditPptx(rebuilt.bytes);
  const rebuiltCell = after.tables.find((table) => table.slideNumber === sourceSlide.slideNumber)?.cells?.find((cell) => cell.row === sourceCell.row && cell.column === sourceCell.column);
  assert.equal(rebuiltCell?.text, updatedText);
  assert.equal(rebuiltCell?.runBreaksBefore?.some((value) => value !== "none"), true);
  assert.deepEqual(after.tables.map((table) => table.structureHash), deck.audit?.tables.map((table) => table.structureHash));
});

test("fresh composition stops before writing a slide whose grouped or unsupported text is not completely mapped", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const incomplete = {
    ...scene,
    slides: scene.slides.map((slide, index) => index === 0 ? {
      ...slide,
      contentCoverage: { ...slide.contentCoverage, exactTextMapped: false, mappedCharacterCount: Math.max(0, slide.contentCoverage.sourceCharacterCount - 12), groupedOrUnsupportedTextPresent: true },
    } : slide),
  };
  await assert.rejects(() => buildStudioCompositionPptx(incomplete, { catalog }), /must be atomized before fresh composition/);
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
