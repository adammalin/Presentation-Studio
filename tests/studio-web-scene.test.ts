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
import { atomizeStudioWebSlide, compileStudioWebScene, inferRepeatedImageSeries, planStudioExportBuild, recommendedStudioRecipe, recomposeStudioWebSlide, resizeStudioTableColumn, resizeStudioTableRow, resolvedStudioTableDesign, studioGeneratedComponents, studioGeometryRequests, studioSceneNeedsRebuild, studioVisualDesignRequest, updateStudioConnectorDesign, updateStudioFigureTreatment, updateStudioTableCellDesign, updateStudioTableDesign, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "../src/lib/studio-web-scene";
import { preflightStudioScene } from "../src/lib/studio-visual-critic";
import { buildCleanupProposalPptx, createGeometryBatchProposal, createVisualDesignProposal } from "../src/lib/cleanup";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { preserveNativeSlide } from "../src/lib/native-slide-preservation";
import { nativeIsolationShapeIds } from "../src/lib/native-object-isolation";
import { adoptStudioComponentStyle, compatibleStudioComponentInstances } from "../src/lib/studio-component-library";
import { planStudioTableContinuation, publishStudioTableExemplar } from "../src/lib/studio-table-workflow";
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

test("Studio content coverage excludes inherited slide-number furniture", async () => {
  const { deck, catalog } = await fixture();
  const slide = catalog.slides[0];
  slide.elements.push({ id: "synthetic-slide-number", kind: "text", name: "Slide Number Placeholder", x: 11_000_000, y: 6_200_000, width: 500_000, height: 200_000, rotation: 0, geometry: "rect", text: "1", placeholderType: "sldNum", sourcePart: slide.sourcePart, sourceShapeId: "999", origin: "slide" });
  const scene = compileStudioWebScene(deck, catalog);
  assert.equal(scene.slides[0].contentCoverage.exactTextMapped, true);
  assert.equal(scene.slides[0].nodes.some((node) => node.sourceShapeId === "999"), false);
});

test("Studio import preserves intentional source table column and row proportions", async () => {
  const { deck, catalog } = await fixture();
  const table = deck.audit?.tables.find((item) => item.columns?.length === 3 && item.rows?.length === 3);
  assert.ok(table?.columns && table.rows);
  table.columns = table.columns.map((column, index) => ({ ...column, widthEmu: [1, 2, 1][index] * 914_400 }));
  table.rows = table.rows.map((row, index) => ({ ...row, heightEmu: [2, 1, 1][index] * 457_200 }));
  const scene = compileStudioWebScene(deck, catalog);
  const tableNode = scene.slides.flatMap((slide) => slide.nodes).find((node) => node.kind === "table" && node.table);
  assert.ok(tableNode);
  const design = resolvedStudioTableDesign(tableNode);
  assert.deepEqual(design.columnWidths, [.25, .5, .25]);
  assert.deepEqual(design.rowHeights, [.5, .25, .25]);
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
  const titleNode = slide?.nodes.find((node) => node.visible && node.role === "title");
  const titleRule = slide ? studioGeneratedComponents(slide).find((component) => component.id.includes("title-rule")) : undefined;
  assert.equal(titleRule?.fillColor, "#00B38F");
  assert.equal(titleRule?.frame.width, Math.round(.934 * 914_400));
  assert.equal(titleRule?.frame.height, Math.round(.079 * 914_400));
  assert.equal(titleNode && titleRule ? titleRule.frame.y - titleNode.frame.y - titleNode.frame.height : undefined, Math.round(.04 * 914_400));
  assert.equal(visual.textStyles.every((style) => style.fontSizePt && style.fontSizePt >= 14), true);
});

test("recomposition keeps only the topmost source heading as the deck title and demotes false duplicate titles", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const sourceSlide = source.slides.find((slide) => slide.nodes.filter((node) => node.kind === "text").length >= 2)!;
  const textNodes = sourceSlide.nodes.filter((node) => node.kind === "text").slice(0, 2);
  const emu = (value: number) => value * 914_400;
  const falseTitle = { ...textNodes[0], id: "false-title", role: "title" as const, text: "A long evidence assertion that is not the page title", sourceFrame: { x: emu(7.5), y: emu(1.1), width: emu(5), height: emu(3), rotation: 0 } };
  const trueTitle = { ...textNodes[1], id: "true-title", role: "title" as const, text: "3.2 - S&T Management", sourceFrame: { x: emu(.47), y: emu(.12), width: emu(8), height: emu(.6), rotation: 0 } };
  const scene = { ...source, slides: source.slides.map((slide) => slide.slideNumber === sourceSlide.slideNumber ? { ...slide, nodes: [falseTitle, trueTitle] } : slide) };
  const designed = recomposeStudioWebSlide(scene, sourceSlide.slideNumber, "ornl-title-content").slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!;
  assert.equal(designed.nodes.find((node) => node.id === "true-title")?.role, "title");
  assert.equal(designed.nodes.find((node) => node.id === "false-title")?.role, "body");
  assert.equal(designed.nodes.filter((node) => node.role === "title").length, 1);
  assert.equal(designed.nodes.find((node) => node.id === "true-title")?.frame.y, emu(.29));
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

test("two labeled figures reserve a real narrative region instead of squeezing exact copy into a footer strip", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const node = (id: string, kind: StudioWebNode["kind"], role: StudioWebNode["role"], x: number, y: number, width: number, height: number, text?: string): StudioWebNode => ({
    ...seed,
    id,
    sourceObjectId: id,
    sourceShapeId: id,
    name: id,
    kind,
    role,
    text,
    sourceParagraphs: text ? [{ index: 1, text, textHash: seed.textHash ?? "hash", characterCount: text.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [18] }] : undefined,
    sourceFrame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 },
    frame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 },
    zIndex: Number(id.match(/\d+/)?.[0] ?? 0),
    visible: true,
    locked: false,
  });
  const longNarrative = "Step 1: Use the first exact source view. Step 2: Compare the second exact source view. Once the project is loaded, retain every approved instruction and technical qualifier without rewriting or hiding it in a footer-sized frame.";
  const nodes = [
    node("title-1", "text", "title", .34, .4, 12.57, .48, "Inverter Basics Demo: PSCAD Model Inverter Control Implementation"),
    node("body-2", "text", "body", .34, 1.9, 12.57, 4.4, longNarrative),
    node("caption-3", "text", "caption", 2.2, 2.2, 2, .3, "First source view"),
    node("caption-4", "text", "caption", 8.4, 2.2, 2, .3, "Second source view"),
    node("image-5", "image", "image", 1.5, 2.7, 4.1, 2.1),
    node("image-6", "image", "image", 8.4, 2.4, 2.4, 3.4),
  ];
  const designed = recomposeStudioWebSlide({ ...scene, slides: [{ ...sourceSlide, nodes }] }, sourceSlide.slideNumber, "ornl-title-labeled-figure-grid");
  const slide = designed.slides[0];
  const body = slide.nodes.find((candidate) => candidate.id === "body-2")!;
  const images = slide.nodes.filter((candidate) => candidate.kind === "image");
  const visualFields = images.map((image) => image.component?.frame);
  assert.ok(body.frame.height >= emu(1.09));
  assert.ok(images.every((image) => image.frame.y + image.frame.height <= body.frame.y));
  assert.ok(body.frame.y + body.frame.height <= emu(6.65));
  assert.equal(images.every((image) => image.component?.role === "figure-media"), true);
  assert.equal(visualFields.every(Boolean), true);
  assert.equal(visualFields[0]?.width, visualFields[1]?.width);
  assert.equal(visualFields[0]?.height, visualFields[1]?.height);
  assert.ok(Math.abs(images[0].frame.width / images[0].frame.height - images[0].sourceFrame.width / images[0].sourceFrame.height) < .001);
  assert.ok(Math.abs(images[1].frame.width / images[1].frame.height - images[1].sourceFrame.width / images[1].sourceFrame.height) < .001);
  assert.equal(studioGeneratedComponents(slide).filter((component) => component.id.endsWith("-visual-field")).length, 2);
  const designedTitle = slide.nodes.find((candidate) => candidate.id === "title-1")!;
  const titleRule = studioGeneratedComponents(slide).find((component) => component.id.includes("title-rule"))!;
  assert.ok(designedTitle.frame.height >= emu(.93));
  assert.ok(titleRule.frame.y >= designedTitle.frame.y + designedTitle.frame.height);
  assert.ok(titleRule.frame.y + titleRule.frame.height <= Math.min(...images.map((image) => image.component!.frame!.y)));

  const plain = recomposeStudioWebSlide({ ...scene, slides: [{ ...sourceSlide, nodes: nodes.filter((candidate) => candidate.role !== "caption") }] }, sourceSlide.slideNumber, "ornl-title-figure-grid").slides[0];
  const plainBody = plain.nodes.find((candidate) => candidate.id === "body-2")!;
  const plainImages = plain.nodes.filter((candidate) => candidate.kind === "image");
  assert.ok(plainBody.frame.height >= emu(1.09));
  assert.ok(plainImages.every((image) => image.frame.y + image.frame.height <= plainBody.frame.y));
  assert.ok(plainBody.frame.y + plainBody.frame.height <= emu(6.65));
});

test("cross-image callouts remain one source-locked evidence field instead of becoming detached cards", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const make = (id: string, kind: StudioWebNode["kind"], role: StudioWebNode["role"], x: number, y: number, width: number, height: number, text?: string): StudioWebNode => ({ ...seed, id, sourceObjectId: id, sourceShapeId: id, name: id, kind, role, text, sourceParagraphs: text ? [{ index: 1, text, textHash: seed.textHash ?? "hash", characterCount: text.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [14] }] : undefined, sourceFrame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, frame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, zIndex: Number(id.match(/\d+/)?.[0] ?? 0), visible: true, locked: false });
  const nodes = [
    make("title-1", "text", "title", .3, .4, 12.5, .5, "Technical sequence"),
    make("body-2", "text", "body", .3, 1.2, 12.5, .6, "Use the exact source relationship field."),
    make("image-3", "image", "image", 1.1, 2.2, 3, 2),
    make("image-4", "image", "image", 7.8, 2.2, 3, 2),
    make("caption-5", "text", "caption", .6, 2.5, 1.3, .3, "Input"),
    make("caption-6", "text", "caption", 10.9, 2.5, 1.3, .3, "Output"),
    make("connector-7", "connector", "connector", 3.9, 2.8, 4.1, .4),
  ];
  const designed = recomposeStudioWebSlide({ ...scene, slides: [{ ...sourceSlide, nodes }] }, sourceSlide.slideNumber, "ornl-title-labeled-figure-grid");
  const treatment = designed.slides[0].figureTreatments.find((candidate) => candidate.id.startsWith("studio-auto-figure-field"));
  assert.ok(treatment);
  assert.deepEqual(new Set(treatment.nodeIds), new Set(["image-3", "image-4", "caption-5", "caption-6", "connector-7"]));
  assert.equal(treatment.relationshipPolicy, "preserve-internal");
  assert.equal(treatment.mode, "preserve-as-unit");
  assert.equal(treatment.verificationStatus, "source-locked");
  const plainDesigned = recomposeStudioWebSlide({ ...scene, slides: [{ ...sourceSlide, nodes }] }, sourceSlide.slideNumber, "ornl-title-figure-grid");
  const plainTreatment = plainDesigned.slides[0].figureTreatments.find((candidate) => candidate.id.startsWith("studio-auto-figure-field"));
  assert.ok(plainTreatment);
  assert.deepEqual(new Set(plainTreatment.nodeIds), new Set(["image-3", "image-4", "caption-5", "caption-6", "connector-7"]));
  assert.equal(plainTreatment.relationshipPolicy, "preserve-internal");
});

test("a grouped native diagram is automatically preserved while surrounding narrative remains editable", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const makeText = (id: string, text: string, binding: StudioWebNode["sourceBinding"], x: number, y: number, width: number, height: number): StudioWebNode => ({ ...seed, id, sourceObjectId: id, sourceShapeId: id, sourceBinding: binding, name: id, kind: "text", role: "body", text, sourceParagraphs: [{ index: 1, text, textHash: seed.textHash ?? "hash", characterCount: text.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [14] }], sourceFrame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, frame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, visible: true, locked: false });
  const title = { ...makeText("title", "Theory revisit", "editable-object", .3, .4, 12.5, .5), role: "title" as const };
  const narrative = makeText("narrative", "Where is the point of interconnection? Where is the main transformer? Preserve every exact question.", "editable-object", .3, 1.2, 12.5, 1.1);
  const native: StudioWebNode = { ...seed, id: "native-group", sourceObjectId: "native-group", sourceShapeId: "native-group", sourceBinding: "editable-object", name: "Grouped native diagram", kind: "native-object", role: "group", text: undefined, sourceParagraphs: undefined, sourceFrame: { x: emu(2.8), y: emu(2.5), width: emu(10), height: emu(4) , rotation: 0 }, frame: { x: emu(2.8), y: emu(2.5), width: emu(10), height: emu(4), rotation: 0 }, visible: true, locked: true };
  const secondNative: StudioWebNode = { ...native, id: "native-group-2", sourceObjectId: "native-group-2", sourceShapeId: "native-group-2", name: "Second grouped native diagram", sourceFrame: { x: emu(9.2), y: emu(3), width: emu(3), height: emu(2), rotation: 0 }, frame: { x: emu(9.2), y: emu(3), width: emu(3), height: emu(2), rotation: 0 } };
  const embedded = makeText("embedded-label", "POI", "catalog-derived", 8.3, 3.2, .8, .3);
  const secondEmbedded = makeText("embedded-label-2", "Second unit", "catalog-derived", 9.6, 3.4, 1.2, .3);
  const designed = recomposeStudioWebSlide({ ...scene, slides: [{ ...sourceSlide, nodes: [title, narrative, native, secondNative, embedded, secondEmbedded] }] }, sourceSlide.slideNumber);
  const slide = designed.slides[0];
  const treatment = slide.figureTreatments.find((candidate) => candidate.id.startsWith("studio-auto-technical-figure"));
  assert.equal(slide.recipe, "ornl-title-two-column");
  assert.ok(treatment?.nodeIds.includes(native.id));
  assert.ok(treatment?.nodeIds.includes(secondNative.id));
  assert.ok(treatment?.nodeIds.includes(embedded.id));
  assert.ok(treatment?.nodeIds.includes(secondEmbedded.id));
  assert.ok(slide.nodes.find((candidate) => candidate.id === narrative.id)!.frame.height >= emu(.8));
  assert.ok(!treatment?.nodeIds.includes(narrative.id));
});

test("a mixed legacy control overview stays one complete source-locked region when object isolation would omit catalog members", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const title: StudioWebNode = { ...seed, id: "overview-title", sourceObjectId: "overview-title", sourceShapeId: "1", sourceBinding: "editable-object", name: "Overview title", kind: "text", role: "title", text: "Complete control overview", sourceFrame: { x: emu(.4), y: emu(.3), width: emu(12.5), height: emu(.7), rotation: 0 }, frame: { x: emu(.4), y: emu(.3), width: emu(12.5), height: emu(.7), rotation: 0 }, visible: true, locked: false };
  const native: StudioWebNode = { ...seed, id: "overview-native", sourceObjectId: "overview-native", sourceShapeId: "9", sourceBinding: "editable-object", name: "Legacy controller group", kind: "native-object", role: "group", text: undefined, textHash: undefined, sourceParagraphs: undefined, sourceFrame: { x: emu(.5), y: emu(1.3), width: emu(5.7), height: emu(1.5), rotation: 0 }, frame: { x: emu(.5), y: emu(1.3), width: emu(5.7), height: emu(1.5), rotation: 0 }, visible: true, locked: false };
  const image = (id: string, shapeId: string, x: number, y: number, width: number, height: number, binding: StudioWebNode["sourceBinding"] = "editable-object"): StudioWebNode => ({ ...seed, id, sourceObjectId: id, sourceShapeId: shapeId, sourceBinding: binding, name: id, kind: "image", role: "image", text: undefined, textHash: undefined, sourceParagraphs: undefined, sourceFrame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, frame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, visible: true, locked: false });
  const leftPanel = image("overview-left", "10", .5, 3.1, 5.8, 2.7);
  const rightPanel = image("overview-right", "13", 6.5, 2.2, 6.2, 4.5);
  const catalogPanel = image("overview-catalog-panel", "31", 2.6, 5.1, 1.2, .7, "catalog-derived");
  const source = { ...scene, slides: [{ ...sourceSlide, nodes: [title, native, leftPanel, rightPanel, catalogPanel] }] };
  assert.equal(recommendedStudioRecipe(source.slides[0]), "ornl-title-two-column");
  const designed = recomposeStudioWebSlide(source, sourceSlide.slideNumber);
  const slide = designed.slides[0];
  const treatment = slide.figureTreatments.find((candidate) => candidate.id.startsWith("studio-auto-technical-overview"));
  assert.ok(treatment);
  assert.deepEqual(new Set(treatment.nodeIds), new Set([native.id, leftPanel.id, rightPanel.id, catalogPanel.id]));
  assert.equal(treatment.mode, "preserve-as-unit");
  assert.equal(treatment.verificationStatus, "source-locked");
  assert.deepEqual(nativeIsolationShapeIds(slide, treatment), []);
  assert.equal(studioGeneratedComponents(slide).some((component) => component.id.includes("technical-overview") && component.lineWidthPt > 0), false);
});

test("question-and-diagram recipe atomizes exact questions and preserves the complete native evidence unit", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const paragraphTexts = [
    "Revisit the exact source model before comparing the evidence.",
    "Where is the point of interconnection?",
    "Which equipment defines the first boundary?",
    "How do the grouped systems relate?",
    "Which glossary terms explain the diagram?",
  ];
  const bodyText = paragraphTexts.join("\n");
  const body: StudioWebNode = {
    ...seed,
    id: "question-body",
    sourceObjectId: "question-body",
    sourceShapeId: "question-body",
    sourceBinding: "editable-object",
    name: "Question body",
    kind: "text",
    role: "body",
    text: bodyText,
    textHash: await sha256Text(bodyText),
    sourceParagraphs: await Promise.all(paragraphTexts.map(async (text, index) => ({ index: index + 1, text, textHash: await sha256Text(text), characterCount: text.length, bullet: false, bulletConfidence: "direct" as const, level: 0, fontFamilies: ["Aptos"], fontSizes: [18] }))),
    sourceFrame: { x: emu(.34), y: emu(1.26), width: emu(12.57), height: emu(4.4), rotation: 0 },
    frame: { x: emu(.34), y: emu(1.26), width: emu(12.57), height: emu(4.4), rotation: 0 },
    visible: true,
    locked: false,
  };
  const title: StudioWebNode = { ...body, id: "question-title", sourceObjectId: "question-title", sourceShapeId: "question-title", name: "Question title", role: "title", text: "Theory revisit", textHash: await sha256Text("Theory revisit"), sourceParagraphs: [{ index: 1, text: "Theory revisit", textHash: await sha256Text("Theory revisit"), characterCount: 14, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [18] }], sourceFrame: { x: emu(.34), y: emu(.4), width: emu(12.57), height: emu(.48), rotation: 0 }, frame: { x: emu(.34), y: emu(.4), width: emu(12.57), height: emu(.48), rotation: 0 } };
  const native: StudioWebNode = { ...seed, id: "question-native", sourceObjectId: "question-native", sourceShapeId: "question-native", sourceBinding: "editable-object", name: "Grouped technical diagram", kind: "native-object", role: "group", text: undefined, sourceParagraphs: undefined, sourceFrame: { x: emu(2.83), y: emu(2.53), width: emu(10.08), height: emu(4), rotation: 0 }, frame: { x: emu(2.83), y: emu(2.53), width: emu(10.08), height: emu(4), rotation: 0 }, visible: true, locked: true };
  const glossaryText = "POI: exact source definition";
  const embedded: StudioWebNode = { ...body, id: "question-glossary", sourceObjectId: "question-glossary", sourceShapeId: "question-glossary", sourceBinding: "catalog-derived", name: "Diagram glossary", text: glossaryText, textHash: await sha256Text(glossaryText), sourceParagraphs: [{ index: 1, text: glossaryText, textHash: await sha256Text(glossaryText), characterCount: glossaryText.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [12] }], sourceFrame: { x: emu(8.1), y: emu(5.7), width: emu(3.2), height: emu(.7), rotation: 0 }, frame: { x: emu(8.1), y: emu(5.7), width: emu(3.2), height: emu(.7), rotation: 0 } };
  const source = { ...scene, slides: [{ ...sourceSlide, nodes: [title, body, native, embedded] }] };
  assert.equal(recommendedStudioRecipe(source.slides[0]), "ornl-title-question-diagram");
  const designed = recomposeStudioWebSlide(source, sourceSlide.slideNumber);
  const slide = designed.slides[0];
  const questions = slide.nodes.filter((node) => node.component?.role === "question-item");
  const intro = slide.nodes.find((node) => node.component?.role === "question-intro");
  const treatment = slide.figureTreatments.find((candidate) => candidate.id.startsWith("studio-auto-question-diagram"));
  assert.equal(slide.recipe, "ornl-title-question-diagram");
  assert.equal(intro?.text, paragraphTexts[0]);
  assert.deepEqual(questions.map((node) => node.text), paragraphTexts.slice(1));
  assert.equal(questions.every((node) => node.sourceBinding === "semantic-atom"), true);
  assert.ok(treatment?.nodeIds.includes(native.id));
  assert.ok(treatment?.nodeIds.includes(embedded.id));
  assert.equal(treatment?.verificationStatus, "source-locked");
  assert.equal(studioGeneratedComponents(slide).filter((component) => component.id.includes("question-rail")).length, 1);
  assert.equal(studioGeneratedComponents(slide).filter((component) => component.id.includes("question-separator")).length, 3);
  const visibleText = slide.nodes.filter((node) => node.visible && node.kind === "text").map((node) => node.text).filter((text): text is string => Boolean(text));
  assert.deepEqual(new Set(visibleText), new Set(["Theory revisit", ...paragraphTexts, glossaryText]));
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

test("repeated sponsor metric groups become an editable ORNL metric grid instead of a preserved legacy figure", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => value * 914_400;
  const title: StudioWebNode = {
    ...seed,
    id: "metric-title",
    sourceObjectId: "metric-title",
    sourceShapeId: "metric-title",
    sourceBinding: "editable-object",
    name: "Metric title",
    kind: "text",
    role: "title",
    text: "Program outcomes",
    sourceFrame: { x: emu(.47), y: emu(.28), width: emu(12), height: emu(.6), rotation: 0 },
    frame: { x: emu(.47), y: emu(.28), width: emu(12), height: emu(.6), rotation: 0 },
    visible: true,
    locked: false,
  };
  const metricTexts = ["27 projects", "$18M invested", "14 partners", "6 states", "42 publications", "9 patents", "31 demonstrations", "12 awards", "85% complete", "4 technologies"];
  const metrics = metricTexts.map((text, index): StudioWebNode => ({
    ...seed,
    id: `metric-${index + 1}`,
    sourceObjectId: `metric-${index + 1}`,
    sourceShapeId: `metric-${index + 1}`,
    sourceBinding: "editable-object",
    name: `Metric ${index + 1}`,
    kind: "text",
    role: "body",
    text,
    sourceTextOrder: index + 2,
    zIndex: index + 2,
    sourceFrame: { x: emu(index % 2 ? 6.8 : .7), y: emu(1.4 + Math.floor(index / 2) * .95), width: emu(5.6), height: emu(.62), rotation: 0 },
    frame: { x: emu(index % 2 ? 6.8 : .7), y: emu(1.4 + Math.floor(index / 2) * .95), width: emu(5.6), height: emu(.62), rotation: 0 },
    visible: true,
    locked: false,
  }));
  const groups = Array.from({ length: 5 }, (_, index): StudioWebNode => ({
    ...seed,
    id: `legacy-group-${index + 1}`,
    sourceObjectId: `legacy-group-${index + 1}`,
    sourceShapeId: `legacy-group-${index + 1}`,
    sourceBinding: "editable-object",
    name: `Legacy metric row ${index + 1}`,
    kind: "native-object",
    role: "group",
    text: undefined,
    textHash: undefined,
    sourceParagraphs: undefined,
    sourceFrame: { x: emu(.5), y: emu(1.25 + index * 1.02), width: emu(12.3), height: emu(.86), rotation: 0 },
    frame: { x: emu(.5), y: emu(1.25 + index * 1.02), width: emu(12.3), height: emu(.86), rotation: 0 },
    zIndex: 30 + index,
    visible: true,
    locked: true,
  }));
  const source = { ...scene, slides: [{ ...sourceSlide, nodes: [title, ...metrics, ...groups] }] };
  assert.equal(recommendedStudioRecipe(source.slides[0]), "ornl-title-metric-grid");
  const slide = recomposeStudioWebSlide(source, sourceSlide.slideNumber).slides[0];
  const metricCards = slide.nodes.filter((node) => node.component?.role === "metric-card");
  assert.equal(slide.recipe, "ornl-title-metric-grid");
  assert.equal(metricCards.length, metricTexts.length);
  assert.deepEqual(metricCards.map((node) => node.text), metricTexts);
  assert.equal(groups.every((group) => slide.nodes.find((node) => node.id === group.id)?.visible === true), true);
  assert.equal(slide.figureTreatments.filter((treatment) => treatment.id.startsWith("studio-auto-metric-icon-")).length, metricTexts.length);
  assert.equal(slide.figureTreatments.every((treatment) => treatment.crop && treatment.lockAspectRatio), true);
  assert.equal(metricCards.every((node) => node.style.fontFamily === "Aptos" && node.style.fontSizePt >= 15.5), true);
  assert.equal(studioGeneratedComponents(slide).filter((component) => component.id.includes("studio-metric-") && component.id.endsWith("-surface")).length, metricTexts.length);
  assert.equal(studioGeneratedComponents(slide).some((component) => component.id.includes("studio-auto") || component.lineWidthPt > 0), false);
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

test("golden five-column image series preserves every source image-heading-evidence relationship", async () => {
  const { deck, catalog } = await fixture();
  const scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.kind === "text")!;
  const emu = (value: number) => Math.round(value * 914_400);
  const textNode = (id: string, text: string, role: StudioWebNode["role"], x: number, y: number, width: number, height: number, zIndex: number): StudioWebNode => ({
    ...seed, id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "editable-object", name: id, kind: "text", role, text, zIndex, sourceTextOrder: zIndex * 100,
    sourceParagraphs: [{ index: 1, text, textHash: seed.textHash ?? "a".repeat(64), characterCount: text.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [16] }],
    sourceFrame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, frame: { x: emu(x), y: emu(y), width: emu(width), height: emu(height), rotation: 0 }, visible: true, locked: false,
  });
  const imageNode = (id: string, x: number, zIndex: number): StudioWebNode => ({
    ...seed, id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "editable-object", name: id, kind: "image", role: "image", text: undefined, textHash: undefined, sourceParagraphs: undefined, zIndex, sourceTextOrder: zIndex * 100,
    sourceFrame: { x: emu(x), y: emu(1.18), width: emu(2.20), height: emu(1.10), rotation: 0 }, frame: { x: emu(x), y: emu(1.18), width: emu(2.20), height: emu(1.10), rotation: 0 }, visible: true, locked: false, style: { ...seed.style, objectFit: "contain" },
  });
  const headings = ["Grid Flexible Solutions", "Building Envelopes & Industrialized Construction", "Energy Storage / Multi-functional Products", "Systems Integration", "High-performance Equipment"];
  const bodies = [
    "Advanced wireless sensor technologies, building energy modeling, communications and controls, and urban-scale energy-optimized solutions.",
    "New and emerging materials, components and systems; productivity, affordability, quality and safety in building construction; durable walls, attics and foundations.",
    "Integrating advanced energy storage in equipment and envelope systems, flexible building loads, dynamic facades and thermal energy storage materials.",
    "Testing new components, equipment and systems in realistic environments before market introduction, including research-house and light-commercial platforms.",
    "Helping industry launch some of the most advanced building equipment technologies on the market for a wide range of applications.",
  ];
  const nodes: StudioWebNode[] = [textNode("series-title", "ORNL is accelerating affordable building solutions", "title", .47, .29, 12.39, .56, 1)];
  for (let ordinal = 0; ordinal < 5; ordinal += 1) {
    const x = .47 + ordinal * 2.45;
    nodes.push(imageNode(`series-image-${ordinal + 1}`, x, 10 + ordinal * 3));
    nodes.push(textNode(`series-heading-${ordinal + 1}`, headings[ordinal], ordinal % 2 ? "body" : "label", x, 2.35, 2.20, .48, 11 + ordinal * 3));
    nodes.push(textNode(`series-body-${ordinal + 1}`, bodies[ordinal], ordinal % 2 ? "caption" : "body", x, 2.95, 2.20, 2.60, 12 + ordinal * 3));
  }
  const source: StudioWebScene = { ...scene, slides: [{ ...sourceSlide, status: "imported", recipe: "source", nodes }] };
  const inferred = inferRepeatedImageSeries(source.slides[0]);
  assert.equal(inferred?.groups.length, 5);
  assert.equal(recommendedStudioRecipe(source.slides[0]), "ornl-title-image-series");

  const designed = recomposeStudioWebSlide(source, sourceSlide.slideNumber);
  const slide = designed.slides[0];
  assert.equal(slide.recipe, "ornl-title-image-series");
  assert.equal(preflightStudioScene(designed).ready, true);
  assert.equal(slide.nodes.filter((node) => node.component?.role === "image-series-media").length, 5);
  assert.equal(slide.nodes.filter((node) => node.component?.role === "image-series-heading").length, 5);
  assert.equal(slide.nodes.filter((node) => node.component?.role === "image-series-body").length, 5);
  for (let ordinal = 0; ordinal < 5; ordinal += 1) {
    const groupId = `studio-image-series-${slide.slideNumber}-${ordinal + 1}`;
    const group = slide.nodes.filter((node) => node.component?.groupId === groupId);
    assert.deepEqual(new Set(group.map((node) => node.component?.role)), new Set(["image-series-media", "image-series-heading", "image-series-body"]));
    assert.equal(group.every((node) => node.component?.ordinal === ordinal), true);
    assert.equal(group.find((node) => node.component?.role === "image-series-body")?.style.fontSizePt, 16);
  }
  assert.equal(studioGeneratedComponents(slide).filter((component) => component.id.endsWith("-heading-band")).length, 5);
  assert.equal(studioGeneratedComponents(slide).some((component) => component.lineWidthPt > 0), false);

  const damaged = recomposeStudioWebSlide(source, sourceSlide.slideNumber, "ornl-title-content");
  const failed = preflightStudioScene(damaged);
  assert.equal(failed.ready, false);
  assert.equal(failed.issues.some((issue) => issue.severity === "blocker" && /image-heading-evidence groups/i.test(issue.message)), true);
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
  const objectiveComponents = studioGeneratedComponents(designed.slides[0]).filter((component) => component.id.includes("objective"));
  assert.equal(objectiveComponents.length, 6);
  assert.equal(objectiveComponents.every((component) => component.lineWidthPt === 0), true);
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
      { id: "unused-panel", kind: "shape" as const, name: "Unused semantic panel", x: 8_000_000, y: 1_500_000, width: 3_000_000, height: 400_000, rotation: 0, geometry: "rect" as const, fill: "#00662C", placeholderType: "body", origin: "layout" as const },
    ] }],
    media: { "template-logo": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=" },
    generatedAt: "2026-08-13T16:00:00.000Z",
  };
  const rebuilt = await buildStudioCompositionPptx(templateScene, { catalog, templateCatalog, strict: false });
  const audited = await auditPptx(rebuilt.bytes);
  assert.equal(rebuilt.slideCount, 1);
  assert.equal(audited.pictures.some((picture) => picture.name === "Template · ORNL brand image"), true);
  assert.equal(audited.editableObjects.some((object) => object.name === "Template panel · Unused semantic panel"), false);
  const occupiedFrame = { x: 8_000_000, y: 1_500_000, width: 3_000_000, height: 400_000, rotation: 0 };
  const firstVisibleNode = sourceSlide.nodes.find((node) => node.visible)!;
  const occupiedScene: StudioWebScene = {
    ...templateScene,
    slides: [{ ...templateScene.slides[0], nodes: templateScene.slides[0].nodes.map((node) => node.id === firstVisibleNode.id ? { ...node, frame: occupiedFrame } : node) }],
  };
  const occupied = await buildStudioCompositionPptx(occupiedScene, { catalog, templateCatalog, strict: false });
  const occupiedAudit = await auditPptx(occupied.bytes);
  assert.equal(occupiedAudit.editableObjects.some((object) => object.name === "Template panel · Unused semantic panel"), true);
  assert.ok(rebuilt.bytes.length > 0);
});

test("fresh composition contains extracted images without distorting their source aspect ratio", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const sourceSlide = source.slides[0];
  const seed = sourceSlide.nodes.find((node) => node.visible)!;
  const imageNode: StudioWebNode = {
    ...seed,
    id: "contained-image",
    sourceObjectId: "contained-image",
    sourceShapeId: "contained-image",
    name: "Contained evidence image",
    kind: "image",
    role: "image",
    text: undefined,
    textHash: undefined,
    sourceParagraphs: undefined,
    sourceFrame: { x: 0, y: 0, width: 4 * 914_400, height: 2 * 914_400, rotation: 0 },
    frame: { x: 5 * 914_400, y: 2 * 914_400, width: 3 * 914_400, height: 3 * 914_400, rotation: 0 },
    mediaPart: "contained-image.png",
    style: { ...seed.style, objectFit: "contain" },
    visible: true,
    locked: false,
  };
  const scene: StudioWebScene = { ...source, slides: [{ ...sourceSlide, recipe: "ornl-title-content", status: "designed", nodes: [...sourceSlide.nodes, imageNode] }] };
  const rebuilt = await buildStudioCompositionPptx(scene, {
    catalog: { ...catalog, media: { ...catalog.media, "contained-image.png": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=" } },
    strict: false,
  });
  const audited = await auditPptx(rebuilt.bytes);
  const image = audited.pictures.find((picture) => picture.name.includes("Contained evidence image"));
  assert.ok(image?.widthEmu && image.heightEmu);
  assert.ok(Math.abs(image.widthEmu / image.heightEmu - 2) < .01);
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
  const sourceWithNotes = await JSZip.loadAsync(bytes);
  const notesPart = "ppt/notesSlides/notesSlide1.xml";
  const notesXml = await sourceWithNotes.file(notesPart)!.async("text");
  sourceWithNotes.file(notesPart, notesXml.replace("<a:t></a:t>", "<a:t>Protected speaker note</a:t>"));
  const sourceWithNotesBytes = await sourceWithNotes.generateAsync({ type: "uint8array" });
  const sourceAudit = await auditPptx(sourceWithNotesBytes);
  const rebuilt = await preserveNativeSlide({ destinationBytes: rasterRebuilt.bytes, sourceBytes: sourceWithNotesBytes, slideNumber: 1 });
  const after = await auditPptx(rebuilt.bytes);
  assert.equal(after.slides[0].textHash, sourceAudit.slides[0].textHash);
  assert.equal(rebuilt.receipt.slideNumber, 1);
  assert.match(rebuilt.receipt.clonedLayoutPart, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/);
  assert.match(rebuilt.receipt.clonedMasterPart ?? "", /^ppt\/slideMasters\/slideMaster\d+\.xml$/);
  assert.equal(rebuilt.receipt.sourceSlideSha256, sourceAudit.slides[0].sourcePartSha256);
  assert.equal(rebuilt.receipt.preservedNotes, true);
  const rebuiltZip = await JSZip.loadAsync(rebuilt.bytes);
  const preservedRelationships = await rebuiltZip.file("ppt/slides/_rels/slide1.xml.rels")!.async("text");
  assert.match(preservedRelationships, /relationships\/notesSlide/);
  const preservedNotes = await rebuiltZip.file("ppt/notesSlides/notesSlide1.xml")!.async("text");
  assert.match(preservedNotes, /Protected speaker note/);
  const relationshipIds = [...preservedRelationships.matchAll(/\bId="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(new Set(relationshipIds).size, relationshipIds.length);
  const twoSlideDestination = await buildStudioCompositionPptx({
    ...source,
    slides: source.slides.slice(0, 2).map((slide) => ({ ...slide, status: "designed" as const, recipe: "ornl-title-content" as const })),
  }, { catalog, strict: false });
  const shifted = await preserveNativeSlide({ destinationBytes: twoSlideDestination.bytes, sourceBytes: sourceWithNotesBytes, sourceSlideNumber: 1, destinationSlideNumber: 2 });
  const shiftedAudit = await auditPptx(shifted.bytes);
  assert.equal(shiftedAudit.slideCount, 2);
  assert.equal(shiftedAudit.slides[1].textHash, sourceAudit.slides[0].textHash);
  assert.equal(shifted.receipt.sourceSlideNumber, 1);
  assert.equal(shifted.receipt.destinationSlideNumber, 2);
  const shiftedZip = await JSZip.loadAsync(shifted.bytes);
  assert.match(await shiftedZip.file("ppt/notesSlides/notesSlide2.xml")!.async("text"), /Protected speaker note/);
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

test("first-class Studio table design survives project persistence and editable PowerPoint export without changing source cells or merge topology", async () => {
  const { deck, catalog } = await fixture();
  let scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides.find((slide) => slide.nodes.some((node) => node.kind === "table" && node.table));
  const sourceTable = sourceSlide?.nodes.find((node) => node.kind === "table" && node.table);
  assert.ok(sourceSlide && sourceTable?.table);
  const sourceCells = sourceTable.table.cells.map((cell) => ({ id: cell.id, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, text: cell.text, semanticColorRole: cell.semanticColorRole }));
  scene = recomposeStudioWebSlide(scene, sourceSlide.slideNumber, "ornl-title-table", undefined, "Use one consistent editable ORNL table component.");
  const designedTable = scene.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!.nodes.find((node) => node.id === sourceTable.id)!;
  const firstCell = designedTable.table!.cells[0];
  scene = updateStudioTableDesign(scene, sourceSlide.slideNumber, sourceTable.id, { borderMode: "none", borderWidthPt: 0, defaultPaddingPt: { top: 5, right: 8, bottom: 5, left: 8 } });
  scene = resizeStudioTableColumn(scene, sourceSlide.slideNumber, sourceTable.id, 1, designedTable.frame.width * .4);
  scene = resizeStudioTableRow(scene, sourceSlide.slideNumber, sourceTable.id, 1, designedTable.frame.height * .24);
  scene = updateStudioTableCellDesign(scene, sourceSlide.slideNumber, sourceTable.id, firstCell.id, { fill: "#00662C", color: "#FFFFFF", fontSizePt: 15, fontWeight: 700, textAlign: "center", verticalAlign: "middle", borders: { top: { type: "none", color: "#DBDCDB", widthPt: 0 }, bottom: { type: "solid", color: "#00662C", widthPt: 2 } } });
  scene = updateStudioTableCellDesign(scene, sourceSlide.slideNumber, sourceTable.id, firstCell.id, { fontSizePt: 16 });
  const resultTable = scene.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!.nodes.find((node) => node.id === sourceTable.id)!;
  const resultCells = resultTable.table!.cells.map((cell) => ({ id: cell.id, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, text: cell.text, semanticColorRole: cell.semanticColorRole }));
  const design = resolvedStudioTableDesign(resultTable);
  assert.deepEqual(resultCells, sourceCells);
  assert.equal(design.borderMode, "none");
  assert.ok(design.columnWidths[0] > .35);
  assert.ok(design.rowHeights[0] > .2);
  assert.deepEqual(design.cellStyles.find((item) => item.cellId === firstCell.id), { cellId: firstCell.id, fill: "#00662C", color: "#FFFFFF", fontSizePt: 16, fontWeight: 700, textAlign: "center", verticalAlign: "middle", borders: { top: { type: "none", color: "#DBDCDB", widthPt: 0 }, bottom: { type: "solid", color: "#00662C", widthPt: 2 } } });
  deck.studioScene = scene;
  const project = createProject("First-class table persistence");
  project.decks = [deck];
  const parsed = projectSchema.parse(project);
  const parsedTable = parsed.decks[0].studioScene?.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)?.nodes.find((node) => node.id === sourceTable.id);
  assert.equal(parsedTable?.table?.design?.borderMode, "none");
  assert.equal(parsedTable?.table?.design?.cellStyles[0]?.borders?.bottom?.color, "#00662C");
  const rebuilt = await buildStudioCompositionPptx({ ...scene, slides: [scene.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!] }, { catalog, strict: false, title: "Editable Studio table" });
  const after = await auditPptx(rebuilt.bytes);
  const sourceAuditTable = deck.audit!.tables.find((table) => table.id === sourceTable.tableId)!;
  const afterTable = after.tables[0];
  assert.ok(afterTable);
  assert.equal(afterTable.structureHash, sourceAuditTable.structureHash);
  assert.deepEqual(afterTable.cells?.filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation).map((cell) => cell.text), sourceAuditTable.cells?.filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation).map((cell) => cell.text));
});

test("fresh table composition normalizes legacy body fills and preserves srgb semantic colors", async () => {
  const { deck, catalog } = await fixture();
  let scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides.find((slide) => slide.nodes.some((node) => node.kind === "table" && node.table));
  const sourceTable = sourceSlide?.nodes.find((node) => node.kind === "table" && node.table);
  assert.ok(sourceSlide && sourceTable?.table);
  const semanticCell = sourceTable.table.cells.find((cell) => cell.row > 1);
  assert.ok(semanticCell);
  scene = recomposeStudioWebSlide(scene, sourceSlide.slideNumber, "ornl-title-table", undefined, "Normalize accidental table styling while retaining meaning-bearing color.");
  scene = {
    ...scene,
    slides: scene.slides.map((slide) => slide.slideNumber !== sourceSlide.slideNumber ? slide : {
      ...slide,
      nodes: slide.nodes.map((node) => node.id !== sourceTable.id || !node.table ? node : {
        ...node,
        table: {
          ...node.table,
          cells: node.table.cells.map((cell) => ({
            ...cell,
            fill: cell.id === semanticCell.id ? "srgb:B50094" : "srgb:005C82",
            semanticColorRole: cell.id === semanticCell.id ? "source-category-plasma" : undefined,
          })),
        },
      }),
    }),
  };
  const rebuilt = await buildStudioCompositionPptx({ ...scene, slides: [scene.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!] }, { catalog, strict: false, title: "ORNL table fill normalization" });
  const after = await auditPptx(rebuilt.bytes);
  const cells = after.tables[0]?.cells?.filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation) ?? [];
  assert.ok(cells.length > 0);
  assert.ok(cells.filter((cell) => cell.row === 1).every((cell) => cell.fillToken === "srgb:00454d"));
  assert.ok(cells.filter((cell) => cell.row > 1 && (cell.row !== semanticCell.row || cell.column !== semanticCell.column)).every((cell) => cell.fillToken === (cell.row % 2 === 0 ? "srgb:f0f2f1" : "srgb:ffffff")));
  assert.equal(cells.find((cell) => cell.row === semanticCell.row && cell.column === semanticCell.column)?.fillToken, "srgb:b50094");
  assert.equal(cells.some((cell) => cell.fillToken === "srgb:000000"), false);
});

test("approved table exemplars propagate only to compatible structures and preserve semantic fills", async () => {
  const { deck, catalog } = await fixture();
  let scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides.find((slide) => slide.nodes.some((node) => node.kind === "table" && node.table));
  const sourceTable = sourceSlide?.nodes.find((node) => node.kind === "table" && node.table);
  const targetSlide = scene.slides.find((slide) => slide.slideNumber !== sourceSlide?.slideNumber);
  assert.ok(sourceSlide && sourceTable?.table && targetSlide);
  scene = updateStudioTableDesign(scene, sourceSlide.slideNumber, sourceTable.id, { borderMode: "none", borderWidthPt: 0, defaultPaddingPt: { top: 5, right: 8, bottom: 5, left: 8 } });
  const styledSource = scene.slides.find((slide) => slide.slideNumber === sourceSlide.slideNumber)!.nodes.find((node) => node.id === sourceTable.id)!;
  const semanticSourceCell = styledSource.table!.cells.find((cell) => cell.row > 1) ?? styledSource.table!.cells.at(-1)!;
  const targetTable: StudioWebNode = {
    ...styledSource,
    id: "studio-compatible-target-table",
    sourceObjectId: "compatible-target-table",
    sourceShapeId: "compatible-target-table",
    name: "Compatible semantic table",
    frame: { ...styledSource.frame },
    sourceFrame: { ...styledSource.sourceFrame },
    table: {
      ...styledSource.table!,
      cells: styledSource.table!.cells.map((cell) => ({
        ...cell,
        id: `target-${cell.id}`,
        ...(cell.id === semanticSourceCell.id ? { fill: "#B50094", semanticColorRole: "source-category-plasma" } : {}),
      })),
      design: {
        ...resolvedStudioTableDesign(styledSource),
        borderMode: "full",
        borderWidthPt: 2,
        cellStyles: [],
      },
    },
  };
  scene = { ...scene, slides: scene.slides.map((slide) => slide.slideNumber !== targetSlide.slideNumber ? slide : { ...slide, nodes: [...slide.nodes, targetTable] }) };
  const beforeTarget = scene.slides.find((slide) => slide.slideNumber === targetSlide.slideNumber)!.nodes.find((node) => node.id === targetTable.id)!;
  const beforeContent = beforeTarget.table!.cells.map((cell) => ({ id: cell.id, text: cell.text, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, fill: cell.fill, semanticColorRole: cell.semanticColorRole }));
  const result = publishStudioTableExemplar(scene, { slideNumber: sourceSlide.slideNumber, tableNodeId: styledSource.id, name: "Approved minimal technical table" });
  const afterTarget = result.scene.slides.find((slide) => slide.slideNumber === targetSlide.slideNumber)!.nodes.find((node) => node.id === targetTable.id)!;
  const afterContent = afterTarget.table!.cells.map((cell) => ({ id: cell.id, text: cell.text, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, fill: cell.fill, semanticColorRole: cell.semanticColorRole }));
  const afterDesign = resolvedStudioTableDesign(afterTarget);
  const targetSemanticCellId = `target-${semanticSourceCell.id}`;
  assert.deepEqual(afterContent, beforeContent);
  assert.equal(afterDesign.borderMode, "none");
  assert.equal(afterDesign.defaultPaddingPt.left, 8);
  assert.equal(afterTarget.table!.cells.find((cell) => cell.id === targetSemanticCellId)?.fill, "#B50094");
  assert.notEqual(afterDesign.cellStyles.find((style) => style.cellId === targetSemanticCellId)?.fill, "#FFFFFF");
  assert.equal(result.affectedTableNodeIds.includes(targetTable.id), true);
  assert.equal(result.scene.tableLibrary?.[0]?.name, "Approved minimal technical table");
  deck.studioScene = result.scene;
  const project = createProject("Table exemplar persistence");
  project.decks = [deck];
  const parsed = projectSchema.parse(project);
  assert.equal(parsed.decks[0].studioScene?.tableLibrary?.[0]?.compatibility.columns, sourceTable.table.columns);
});

test("table continuation planning repeats headers and never splits a merged body unit", async () => {
  const { deck, catalog } = await fixture();
  let scene = compileStudioWebScene(deck, catalog);
  const sourceSlide = scene.slides.find((slide) => slide.nodes.some((node) => node.kind === "table" && node.table));
  const sourceTable = sourceSlide?.nodes.find((node) => node.kind === "table" && node.table);
  assert.ok(sourceSlide && sourceTable?.table);
  const rows = 10;
  const columns = 2;
  const cells: NonNullable<StudioWebNode["table"]>["cells"] = [];
  for (let row = 1; row <= rows; row += 1) {
    if (row !== 5) cells.push({ id: `continuation-r${row}c1`, row, column: 1, rowSpan: row === 4 ? 2 : 1, columnSpan: 1, text: `R${row}C1` });
    cells.push({ id: `continuation-r${row}c2`, row, column: 2, rowSpan: 1, columnSpan: 1, text: `R${row}C2` });
  }
  const denseTable: StudioWebNode = {
    ...sourceTable,
    id: "continuation-table",
    sourceObjectId: "continuation-table",
    sourceShapeId: "continuation-table",
    table: {
      rows,
      columns,
      cells,
      design: {
        ...resolvedStudioTableDesign(sourceTable),
        headerRows: 1,
        columnWidths: [.5, .5],
        rowHeights: Array.from({ length: rows }, () => 1 / rows),
        cellStyles: [],
      },
    },
  };
  scene = {
    ...scene,
    slides: scene.slides.map((slide) => slide.slideNumber !== sourceSlide.slideNumber ? slide : {
      ...slide,
      recipe: "ornl-title-table",
      status: "designed",
      nodes: slide.nodes.map((node) => node.id === sourceTable.id ? denseTable : node),
    }),
  };
  const result = planStudioTableContinuation(scene, { slideNumber: sourceSlide.slideNumber, tableNodeId: denseTable.id, maximumBodyRowsPerSlide: 3 });
  assert.equal(result.plan.status, "ready");
  assert.equal(result.plan.segments.length, 4);
  assert.equal(result.plan.segments.some((segment) => segment.bodyRowStart <= 4 && segment.bodyRowEnd >= 5), true);
  assert.equal(result.plan.segments.some((segment) => segment.bodyRowEnd === 4 || segment.bodyRowStart === 5), false);
  const headerIds = new Set(cells.filter((cell) => cell.row === 1).map((cell) => cell.id));
  assert.equal(result.plan.segments.every((segment) => [...headerIds].every((id) => segment.sourceCellIds.includes(id))), true);
  const bodyIds = cells.filter((cell) => cell.row > 1).map((cell) => cell.id);
  assert.deepEqual(result.plan.segments.flatMap((segment) => segment.sourceCellIds.filter((id) => !headerIds.has(id))).sort(), [...bodyIds].sort());
  const blocked = planStudioTableContinuation(scene, { slideNumber: sourceSlide.slideNumber, tableNodeId: denseTable.id, maximumBodyRowsPerSlide: 1 });
  assert.equal(blocked.plan.status, "blocked");
  assert.match(blocked.plan.blockers.join(" "), /merged unit/i);
  deck.studioScene = result.scene;
  const project = createProject("Continuation persistence");
  project.decks = [deck];
  assert.equal(projectSchema.parse(project).decks[0].studioScene?.tableContinuationPlans?.[0]?.segments.length, 4);
  const targetScene = { ...result.scene, slides: result.scene.slides.filter((slide) => slide.slideNumber === sourceSlide.slideNumber) };
  const rebuilt = await buildStudioCompositionPptx(targetScene, { catalog, strict: false, title: "Native table continuation" });
  assert.equal(rebuilt.slideCount, 4);
  assert.deepEqual(rebuilt.outputSlides.map((slide) => slide.sourceSlideNumber), [sourceSlide.slideNumber, sourceSlide.slideNumber, sourceSlide.slideNumber, sourceSlide.slideNumber]);
  assert.deepEqual(rebuilt.outputSlides.map((slide) => slide.continuation?.segmentOrdinal), [1, 2, 3, 4]);
  const audit = await auditPptx(rebuilt.bytes);
  assert.equal(audit.slideCount, 4);
  assert.equal(audit.tables.length, 4);
  assert.equal(audit.tables.every((table) => (table.cells ?? []).some((cell) => cell.text === "R1C1") && (table.cells ?? []).some((cell) => cell.text === "R1C2")), true);
});

test("verified Studio connector authoring binds stable endpoints and exports one editable PowerPoint connector", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const sourceSlide = source.slides[0];
  const seed = sourceSlide.nodes[0];
  assert.ok(seed);
  const box = (id: string, x: number): StudioWebNode => ({ ...seed, id, sourceObjectId: id, sourceShapeId: id, sourceBinding: "semantic-atom", name: id, kind: "shape", role: "other", sourceFrame: { x, y: 2 * 914_400, width: 2 * 914_400, height: 1 * 914_400, rotation: 0 }, frame: { x, y: 2 * 914_400, width: 2 * 914_400, height: 1 * 914_400, rotation: 0 }, visible: true, locked: false, exactContent: false, text: undefined, textHash: undefined, sourceParagraphs: undefined, table: undefined, tableId: undefined, connector: undefined, mediaPart: undefined });
  const from = box("verified-from", 2 * 914_400);
  const to = box("verified-to", 8 * 914_400);
  const connector: StudioWebNode = { ...box("verified-connector", 4 * 914_400), kind: "connector", role: "connector", frame: { x: 4 * 914_400, y: 2.5 * 914_400, width: 4 * 914_400, height: 1, rotation: 0 } };
  let scene: StudioWebScene = { ...source, slides: [{ ...sourceSlide, recipe: "ornl-title-content", status: "designed", contentCoverage: { exactTextMapped: true, sourceCharacterCount: 0, mappedCharacterCount: 0, sourceTextBoxCount: 0, mappedTextNodeCount: 0, groupedOrUnsupportedTextPresent: false }, nodes: [from, to, connector], figureTreatments: [] }] };
  scene = updateStudioFigureTreatment(scene, sourceSlide.slideNumber, { id: "verified-diagram", nodeIds: [from.id, to.id, connector.id], mode: "redraw-candidate", verificationStatus: "verified", intentSummary: "One verified causal relationship", informationInventory: ["Source", "Destination", "Directed connector"], invariants: ["Direction and endpoint identity remain unchanged."], rationale: "The relationship has been verified from authoritative source content.", relationshipPolicy: "editable-diagram" });
  scene = updateStudioConnectorDesign(scene, sourceSlide.slideNumber, connector.id, { fromNodeId: from.id, toNodeId: to.id, fromSide: "right", toSide: "left", stroke: "#00662C", widthPt: 1.5, dash: "solid", beginArrow: "none", endArrow: "triangle", verificationStatus: "verified" });
  const authored = scene.slides[0].nodes.find((node) => node.id === connector.id)!;
  assert.equal(authored.connector?.fromNodeId, from.id);
  assert.deepEqual(scene.slides[0].figureTreatments[0].relationships, [{ fromNodeId: connector.id, toNodeId: from.id, kind: "connects-from" }, { fromNodeId: connector.id, toNodeId: to.id, kind: "connects-to" }]);
  const priorX = authored.frame.x;
  scene = updateStudioWebNodeFrame(scene, sourceSlide.slideNumber, from.id, { ...from.frame, x: from.frame.x + .5 * 914_400 });
  assert.notEqual(scene.slides[0].nodes.find((node) => node.id === connector.id)?.frame.x, priorX);
  deck.studioScene = scene;
  const project = createProject("Verified connector persistence");
  project.decks = [deck];
  assert.equal(projectSchema.parse(project).decks[0].studioScene?.slides[0].nodes.find((node) => node.id === connector.id)?.connector?.verificationStatus, "verified");
  const rebuilt = await buildStudioCompositionPptx(scene, { catalog, title: "Verified connector" });
  const after = await auditPptx(rebuilt.bytes);
  const exportedConnector = after.editableObjects.find((object) => /verified-connector/i.test(object.name));
  assert.ok(exportedConnector);
  assert.equal(exportedConnector.kind, "shape");
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
    sourceFigureRasters: { "source-locked-schematic": { data: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=", width: 1, height: 1 } },
  });
  const after = await auditPptx(rebuilt.bytes);
  assert.equal(after.pictures.some((picture) => picture.name.startsWith("Source-locked · Technical schematic")), true);
  assert.equal(rebuilt.warnings.some((warning) => warning.includes("source-locked PowerPoint-rendered evidence unit from an object-isolated native render")), true);
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

test("a reusable component definition propagates style only to compatible instances and survives project persistence", async () => {
  const { deck, catalog } = await fixture();
  const source = compileStudioWebScene(deck, catalog);
  const seed = source.slides.flatMap((slide) => slide.nodes).find((node) => node.kind === "text" && !node.locked);
  assert.ok(seed);
  const sourceFrame = { x: 914_400, y: 1_828_800, width: 3_657_600, height: 914_400, rotation: 0 };
  const first: StudioWebNode = { ...seed, id: "component-source", sourceObjectId: "component-source", sourceShapeId: "component-source", text: "Exact source component copy", frame: sourceFrame, sourceFrame, component: { groupId: "component-group-1", role: "supporting-copy" }, style: { ...seed.style, fontSizePt: 19, fontWeight: 600, color: "#00454D", paddingPt: { top: 4, right: 6, bottom: 4, left: 6 } } };
  const second: StudioWebNode = { ...seed, id: "component-target", sourceObjectId: "component-target", sourceShapeId: "component-target", text: "Different exact target copy", frame: { ...sourceFrame, y: 2_971_800 }, sourceFrame: { ...sourceFrame, y: 2_971_800 }, component: { groupId: "component-group-2", role: "supporting-copy" }, style: { ...seed.style, fontSizePt: 13, fontWeight: 400, color: "#373A36", paddingPt: { ...seed.style.paddingPt } } };
  const dark: StudioWebNode = { ...second, id: "component-dark", sourceObjectId: "component-dark", sourceShapeId: "component-dark", text: "Dark surface variant", component: { groupId: "component-group-3", role: "supporting-copy" }, style: { ...second.style, background: "#00454D", color: "#FFFFFF" } };
  const scene: StudioWebScene = {
    ...source,
    slides: [
      { ...source.slides[0], status: "designed", recipe: "ornl-title-content", background: "#FFFFFF", nodes: [first] },
      { ...source.slides[1], status: "designed", recipe: "ornl-title-content", background: "#FFFFFF", nodes: [second, dark] },
    ],
  };
  assert.deepEqual(compatibleStudioComponentInstances(scene, scene.slides[0].slideNumber, first.id), [{ slideNumber: scene.slides[0].slideNumber, nodeId: first.id }, { slideNumber: scene.slides[1].slideNumber, nodeId: second.id }]);
  const before = scene.slides.flatMap((slide) => slide.nodes.map((node) => ({ id: node.id, text: node.text, frame: node.frame })));
  const result = adoptStudioComponentStyle(scene, { slideNumber: scene.slides[0].slideNumber, nodeId: first.id });
  const after = result.scene.slides.flatMap((slide) => slide.nodes.map((node) => ({ id: node.id, text: node.text, frame: node.frame })));
  const updatedSecond = result.scene.slides[1].nodes.find((node) => node.id === second.id)!;
  const unchangedDark = result.scene.slides[1].nodes.find((node) => node.id === dark.id)!;
  assert.deepEqual(after, before);
  assert.equal(result.affectedNodeIds.length, 2);
  assert.equal(updatedSecond.style.fontSizePt, 19);
  assert.equal(updatedSecond.style.color, "#00454D");
  assert.match(updatedSecond.component?.definitionId ?? "", /supporting-copy-light/);
  assert.equal(unchangedDark.style.fontSizePt, 13);
  assert.equal(result.scene.componentLibrary?.length, 1);
  deck.studioScene = result.scene;
  const project = createProject("Reusable component persistence");
  project.decks = [deck];
  const parsed = projectSchema.parse(project);
  assert.equal(parsed.decks[0].studioScene?.componentLibrary?.[0].role, "supporting-copy");
  assert.equal(parsed.decks[0].studioScene?.slides[1].nodes.find((node) => node.id === second.id)?.component?.definitionId, result.definition.id);
});
