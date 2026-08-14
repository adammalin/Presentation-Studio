import assert from "node:assert/strict";
import fs from "node:fs/promises";
import JSZip from "jszip";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { auditPptx } from "../src/lib/pptx-audit";
import { createProject, projectSchema } from "../src/lib/project";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { buildSlideRenderCatalog } from "../src/lib/template-catalog";
import { atomizeStudioWebSlide, compileStudioWebScene, planStudioExportBuild, recommendedStudioRecipe, recomposeStudioWebSlide, studioGeneratedComponents, studioGeometryRequests, studioSceneNeedsRebuild, studioVisualDesignRequest, updateStudioFigureTreatment, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "../src/lib/studio-web-scene";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createVisualDesignProposal } from "../src/lib/cleanup";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { preserveNativeSlide } from "../src/lib/native-slide-preservation";
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

test("process-flow recipe keeps low technical inputs out of the footer and pairs the true output by source order", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const base = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const textNode = async (id: string, text: string, role: StudioWebNode["role"], zIndex: number, x: number, y: number, paragraphs = [text]): Promise<StudioWebNode> => ({
    ...base,
    id,
    sourceObjectId: id,
    sourceShapeId: id,
    sourceBinding: "editable-object",
    name: id,
    kind: "text",
    role,
    zIndex,
    sourceTextOrder: zIndex * 100,
    text,
    textHash: await sha256Text(text),
    sourceParagraphs: await Promise.all(paragraphs.map(async (paragraph, index) => ({ index: index + 1, text: paragraph, textHash: await sha256Text(paragraph), characterCount: paragraph.length, bullet: role === "body", bulletConfidence: "direct" as const, level: 0, fontFamilies: ["Aptos"], fontSizes: [16] }))),
    sourceFrame: { x: x * 914_400, y: y * 914_400, width: 2.4 * 914_400, height: .5 * 914_400, rotation: 0 },
    frame: { x: x * 914_400, y: y * 914_400, width: 2.4 * 914_400, height: .5 * 914_400, rotation: 0 },
    visible: true,
    locked: false,
  });
  const imageNode = (id: string, zIndex: number, x: number, y: number): StudioWebNode => ({
    ...base,
    id,
    sourceObjectId: id,
    sourceShapeId: id,
    sourceBinding: "editable-object",
    name: id,
    kind: "image",
    role: "image",
    zIndex,
    sourceTextOrder: zIndex * 100,
    text: undefined,
    textHash: undefined,
    sourceParagraphs: undefined,
    sourceFrame: { x: x * 914_400, y: y * 914_400, width: .6 * 914_400, height: .6 * 914_400, rotation: 0 },
    frame: { x: x * 914_400, y: y * 914_400, width: .6 * 914_400, height: .6 * 914_400, rotation: 0 },
    visible: true,
    locked: false,
  });
  const nodes: StudioWebNode[] = [
    await textNode("process-title", "Exact conversion process", "title", 0, .5, .3),
    await textNode("stage-one", "Conversion Tool", "body", 1, 4.5, 4.6),
    imageNode("steady-icon", 2, 3.0, 4.7), await textNode("steady-label", "Steady-state network data files", "caption", 3, .8, 4.7),
    imageNode("dynamic-icon", 4, 3.0, 5.4), await textNode("dynamic-label", "Dynamic model data files", "caption", 5, .8, 5.4),
    imageNode("sequence-icon", 6, 3.0, 6.1), await textNode("sequence-label", "Sequence network data files", "caption", 7, .8, 6.1),
    imageNode("location-icon", 8, 3.2, 6.9), await textNode("location-label", "Location and diagram data files", "caption", 9, .8, 6.8),
    await textNode("stage-two", "Fix conversion errors", "body", 10, 7.4, 4.6),
    imageNode("output-icon", 11, 11.0, 5.6), await textNode("output-label", "EMT input data files", "caption", 12, 9.4, 5.3),
    await textNode("support", "The process crosses parties. Exact source detail remains available.", "body", 13, .5, 1.4, ["The process crosses parties.", "Exact source detail remains available."]),
  ];
  const processSource = { ...scene, slides: [{ ...sourceSlide, nodes }] };
  assert.equal(recommendedStudioRecipe(processSource.slides[0]), "ornl-title-process-flow");
  const designed = recomposeStudioWebSlide(processSource, sourceSlide.slideNumber);
  const designedSlide = designed.slides[0];
  const inputs = designedSlide.nodes.filter((node) => node.component?.role === "process-input");
  assert.equal(inputs.length, 4);
  assert.deepEqual(inputs.map((node) => node.text), ["Steady-state network data files", "Dynamic model data files", "Sequence network data files", "Location and diagram data files"]);
  assert.equal(designedSlide.nodes.find((node) => node.text === "EMT input data files")?.component?.role, "process-output");
  assert.equal(designedSlide.nodes.find((node) => node.id === "location-icon")?.component?.role, "process-icon");
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
  assert.equal(after.textBoxes.reduce((sum, item) => sum + item.bulletParagraphCount, 0), before.textBoxes.reduce((sum, item) => sum + item.bulletParagraphCount, 0));
  assert.equal(after.fonts.some((font) => font.family === "Aptos"), true);
  assert.equal(after.textBoxes.every((textBox) => textBox.fontFamilies.every((family) => family === "Aptos")), true);
  assert.equal(after.tables.every((table) => table.cellFonts.every((family) => family === "Aptos")), true);
});

test("fresh composition builds converted template artwork into the same editable slide result", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const templateScene: StudioWebScene = {
    ...scene,
    slides: [{ ...sourceSlide, status: "designed", recipe: "template-layout", targetLayoutId: "installed-layout" }],
  };
  const templateCatalog = {
    id: "template",
    name: "Converted ORNL Template",
    sha256: "template-sha",
    slideWidth: templateScene.slideSize.width,
    slideHeight: templateScene.slideSize.height,
    masterCount: 1,
    layouts: [{ id: "installed-layout", name: "ORNL content", category: "content" as const, background: "#FFFFFF", placeholderTypes: ["title", "body"], sourcePart: "ppt/slideLayouts/slideLayout1.xml", elements: [
      { id: "brand-rule", kind: "shape" as const, name: "ORNL brand rule", x: 420_000, y: 900_000, width: 1_000_000, height: 32_000, rotation: 0, geometry: "rect" as const, fill: "#007833", origin: "master" as const },
      { id: "brand-image", kind: "image" as const, name: "ORNL brand image", x: 10_900_000, y: 6_300_000, width: 700_000, height: 300_000, rotation: 0, geometry: "rect" as const, mediaId: "template-logo", origin: "master" as const },
    ] }],
    media: { "template-logo": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=" },
    generatedAt: "2026-08-13T16:00:00.000Z",
  };
  const rebuilt = await buildStudioCompositionPptx(templateScene, { catalog, templateCatalog, strict: false });
  const audited = await auditPptx(rebuilt.bytes);
  assert.equal(rebuilt.slideCount, 1);
  assert.equal(audited.pictures.some((picture) => picture.name === "Template · ORNL brand image"), true);
  assert.ok(rebuilt.bytes.length > 0);
});

test("build-all planning preserves source slides and routes each designed slide to its truthful PowerPoint compiler", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const fresh = { ...source.slides[1], status: "designed" as const, recipe: "ornl-title-content" as const };
  const template = { ...source.slides[2], status: "designed" as const, recipe: "template-layout" as const, targetLayoutId: "installed-layout" };
  const plan = planStudioExportBuild({ ...source, slides: [source.slides[0], fresh, template] });
  assert.deepEqual(plan, {
    preservedSourceSlideNumbers: [source.slides[0].slideNumber],
    freshCompositionSlideNumbers: [fresh.slideNumber, template.slideNumber],
    nativeTemplateSlideNumbers: [],
  });
});

test("source-preserved title composition remains visually locked while exact text stays auditable", async () => {
  const { bytes, deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const sourceSlide = source.slides[0];
  const rasterRebuilt = await buildStudioCompositionPptx({ ...source, slides: [sourceSlide] }, {
    catalog,
    sourceSlideText: { [sourceSlide.slideNumber]: deck.audit!.slides[0].text },
    sourceSlideRasters: { [sourceSlide.slideNumber]: { data: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=", width: 1, height: 1 } },
  });
  await assert.rejects(() => preserveNativeSlide({ destinationBytes: rasterRebuilt.bytes, sourceBytes: bytes, slideNumber: 1 }), /notesSlide1\.xml.*held rather than flattening/i);
  const minimalSource = await JSZip.loadAsync(bytes);
  const slideRelationshipsPart = "ppt/slides/_rels/slide1.xml.rels";
  const relationships = await minimalSource.file(slideRelationshipsPart)!.async("text");
  minimalSource.file(slideRelationshipsPart, relationships.replace(/<Relationship\b[^>]*Type="[^"]*\/notesSlide"[^>]*\/>/g, ""));
  const minimalSourceBytes = await minimalSource.generateAsync({ type: "uint8array" });
  const minimalAudit = await auditPptx(minimalSourceBytes);
  const rebuilt = await preserveNativeSlide({ destinationBytes: rasterRebuilt.bytes, sourceBytes: minimalSourceBytes, slideNumber: 1 });
  const after = await auditPptx(rebuilt.bytes);
  assert.equal(after.slides[0].textHash, minimalAudit.slides[0].textHash);
  assert.equal(rebuilt.receipt.slideNumber, 1);
  assert.match(rebuilt.receipt.clonedLayoutPart, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/);
  assert.match(rebuilt.receipt.clonedMasterPart ?? "", /^ppt\/slideMasters\/slideMaster\d+\.xml$/);
  assert.equal(rebuilt.receipt.sourceSlideSha256, minimalAudit.slides[0].sourcePartSha256);
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

test("fresh composition carries a disclosed complex source figure as one PowerPoint-rendered evidence unit", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const sourceSlide = source.slides[0];
  const seed = sourceSlide.nodes[0];
  assert.ok(seed);
  const nativeObject: StudioWebNode = {
    ...seed,
    id: "source-locked-native-object",
    sourceObjectId: "source-locked-native-object",
    sourceShapeId: "source-locked-native-object",
    sourceBinding: "catalog-derived",
    name: "Complex technical schematic",
    kind: "native-object",
    role: "group",
    text: undefined,
    textHash: undefined,
    sourceParagraphs: undefined,
    sourceFrame: { x: 914_400, y: 1_828_800, width: 3_657_600, height: 1_828_800, rotation: 0 },
    frame: { x: 4_571_999, y: 1_371_600, width: 6_400_800, height: 3_200_400, rotation: 0 },
    visible: true,
    locked: true,
    exactContent: false,
  };
  const scene: StudioWebScene = {
    ...source,
    slides: [{
      ...sourceSlide,
      status: "designed",
      recipe: "ornl-title-two-column",
      nodes: [...sourceSlide.nodes, nativeObject],
      figureTreatments: [{
        id: "source-locked-schematic",
        nodeIds: [nativeObject.id],
        mode: "preserve-and-frame",
        verificationStatus: "source-locked",
        intentSummary: "Technical schematic with meaning-bearing labels and relationships",
        informationInventory: ["Schematic labels", "Connections"],
        invariants: ["Preserve all labels and relationships"],
        rationale: "Retain exact source pixels while Studio improves the surrounding composition.",
      }],
    }],
  };
  const rebuilt = await buildStudioCompositionPptx(scene, {
    catalog,
    sourceSlideRasters: { 1: { data: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=", width: 1, height: 1 } },
  });
  const after = await auditPptx(rebuilt.bytes);
  assert.equal(after.pictures.some((picture) => picture.name.startsWith("Source-locked · Technical schematic")), true);
  assert.equal(rebuilt.warnings.some((warning) => warning.includes("source-locked PowerPoint-rendered evidence unit")), true);
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
