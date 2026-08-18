import assert from "node:assert/strict";
import test from "node:test";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { bindNewStudioSceneToGeneratedPowerPoint, createNewStudioPresentationScene } from "../src/lib/new-studio-presentation";
import { processResourceInput } from "../src/lib/resource-ingestion";
import { assertExactResourceExcerpt, resourceTextPage } from "../src/lib/resource-text-access";
import { deriveLayoutSemantics } from "../src/lib/layout-semantics";
import { createProject, projectSchema } from "../src/lib/project";
import type { TemplateCatalog, TemplateLayoutPreview } from "../src/lib/template-catalog";

const ONE_PIXEL_PNG = Uint8Array.from(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3ZQAAAAASUVORK5CYII=", "base64"));

function layout(input: Omit<TemplateLayoutPreview, "semantic">): TemplateLayoutPreview {
  return { ...input, semantic: deriveLayoutSemantics(input, 12_192_000, 6_858_000) };
}

function templateCatalog(): TemplateCatalog {
  const title = layout({
    id: "approved-title",
    name: "ORNL Title",
    category: "title",
    background: "#FFFFFF",
    sourcePart: "ppt/slideLayouts/slideLayout1.xml",
    placeholderTypes: ["ctrTitle", "subTitle"],
    elements: [
      { id: "green-anchor", kind: "shape", name: "ORNL green anchor", x: 0, y: 0, width: 1_100_000, height: 6_858_000, rotation: 0, geometry: "rect", fill: "#00662C", origin: "master" },
      { id: "title", kind: "text", name: "Title", x: 1_650_000, y: 1_500_000, width: 9_750_000, height: 1_300_000, rotation: 0, geometry: "rect", placeholderType: "ctrTitle", origin: "layout" },
      { id: "subtitle", kind: "text", name: "Subtitle", x: 1_650_000, y: 3_000_000, width: 8_400_000, height: 900_000, rotation: 0, geometry: "rect", placeholderType: "subTitle", origin: "layout" },
    ],
  });
  const content = layout({
    id: "approved-content",
    name: "ORNL Title and Content",
    category: "content",
    background: "#FFFFFF",
    sourcePart: "ppt/slideLayouts/slideLayout2.xml",
    placeholderTypes: ["title", "body"],
    elements: [
      { id: "rule", kind: "shape", name: "ORNL title rule", x: 430_000, y: 1_080_000, width: 950_000, height: 35_000, rotation: 0, geometry: "rect", fill: "#00662C", origin: "master" },
      { id: "content-title", kind: "text", name: "Title", x: 430_000, y: 220_000, width: 11_300_000, height: 800_000, rotation: 0, geometry: "rect", placeholderType: "title", origin: "layout" },
      { id: "content-body", kind: "text", name: "Body", x: 520_000, y: 1_350_000, width: 10_900_000, height: 4_900_000, rotation: 0, geometry: "rect", placeholderType: "body", origin: "layout" },
    ],
  });
  return { id: "ornl-test", name: "Synthetic ORNL Template", sha256: "f".repeat(64), slideWidth: 12_192_000, slideHeight: 6_858_000, masterCount: 1, layouts: [title, content], media: {}, generatedAt: "2026-08-17T16:00:00.000Z" };
}

test("Resource text requires the active global AI session and remains bounded", async () => {
  const resource = await processResourceInput({ name: "source.txt", bytes: new TextEncoder().encode("Alpha evidence with 42 units.\nBeta evidence preserves attribution.") });
  assert.throws(() => resourceTextPage(resource), /no extracted text available/i);
  resource.mcpAccess = "text";
  const page = resourceTextPage(resource, 0, 1_000);
  assert.match(page.text, /42 units/);
  assert.equal(page.derivativeSha256, resource.derivatives?.[0].sha256);
  assert.doesNotThrow(() => assertExactResourceExcerpt(resource, "Alpha evidence with 42 units."));
  assert.throws(() => assertExactResourceExcerpt(resource, "Invented evidence"), /not found/);
});

test("document Resources instantiate approved ORNL designs in the native Studio JSON scene and editable PowerPoint", async () => {
  const source = await processResourceInput({ name: "brief.txt", bytes: new TextEncoder().encode("The program reduces processing time by 42 percent. The verified pilot retained all required technical controls. Implementation proceeds in three stages.") });
  source.mcpAccess = "text";
  const image = await processResourceInput({ name: "evidence.png", mediaType: "image/png", bytes: ONE_PIXEL_PNG });
  image.mcpAccess = "preview";
  const catalog = templateCatalog();
  const scene = await createNewStudioPresentationScene({
    deckId: "new-deck",
    name: "Resource-grounded presentation",
    communicationJob: "By the end, technical leaders should understand the verified program result and implementation path.",
    expression: "balanced",
    slides: [
      { title: "Verified controls support a faster processing path", subtitle: "Source-grounded program overview", body: [], recipe: "template-layout", layoutId: "approved-title", sourceReferences: [{ resourceId: source.id, exactExcerpt: "The program reduces processing time by 42 percent." }], rationale: "Open with the source-supported result using the protected approved title design." },
      { title: "The pilot retained every required technical control", body: ["Processing time decreased by 42 percent.", "All required technical controls remained in place.", "Implementation proceeds in three stages."], recipe: "ornl-title-two-column", imageResourceIds: [image.id], sourceReferences: [{ resourceId: source.id, exactExcerpt: "The verified pilot retained all required technical controls. Implementation proceeds in three stages." }], rationale: "Pair the assertion and implementation evidence with the authorized source image." },
    ],
  }, [source, image], catalog);

  assert.equal(scene.slides.length, 2);
  assert.equal(scene.slides[0].recipe, "template-layout");
  assert.equal(scene.slides.every((slide) => (slide.resourceBindings?.length ?? 0) > 0), true);
  assert.equal(scene.slides.flatMap((slide) => slide.nodes).every((node) => node.style.fontFamily === "Aptos" && node.style.borderWidthPt === 0), true);
  assert.equal(scene.slides[1].nodes.some((node) => node.mediaPart === `resource:${image.id}`), true);

  const resourceImage = `data:image/png;base64,${Buffer.from(ONE_PIXEL_PNG).toString("base64")}`;
  const built = await buildStudioCompositionPptx(scene, {
    catalog: { id: "resources", name: "Resources", sha256: scene.sourceSha256, slideWidth: scene.slideSize.width, slideHeight: scene.slideSize.height, slides: [], media: { [`resource:${image.id}`]: resourceImage }, generatedAt: new Date().toISOString(), renderer: "local-ooxml-preview" },
    templateCatalog: catalog,
    strict: true,
    title: "Resource-grounded presentation",
  });
  const audit = await auditPptx(built.bytes);
  assert.equal(audit.slideCount, 2);
  assert.equal(audit.slides[0].text.includes("Verified controls support a faster processing path"), true);
  assert.equal(audit.slides[1].text.includes("42 percent"), true);
  assert.equal(audit.textBoxes.every((textBox) => textBox.fontFamilies.every((font) => font === "Aptos")), true);
  const rebound = bindNewStudioSceneToGeneratedPowerPoint(scene, "a".repeat(64), audit);
  assert.equal(rebound.sourceSha256, "a".repeat(64));
  assert.equal(rebound.slides[0].recipe, "source");
  assert.equal(rebound.slides[1].recipe, "ornl-title-two-column");
  assert.equal(rebound.slides[0].resourceBindings?.[0].resourceSha256, source.sha256);

  const generated = await processResourceInput({ name: "Resource-grounded-presentation.pptx", mediaType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", bytes: built.bytes });
  const project = createProject("Resource-grounded presentation");
  project.project.type = "new-presentation";
  project.settings.contentPolicy = "source-grounded-generative";
  project.settings.defaultOperationScope = "compose";
  project.resources = [source, image, generated];
  project.decks = [{
    id: scene.deckId,
    name: generated.name,
    sourceResourceId: generated.id,
    sourceSha256: generated.sha256,
    operationScope: "compose",
    templateClassification: "current-ornl",
    status: "ready-for-cleanup",
    audit,
    studioScene: rebound,
    protectedSlideNumbers: [1],
  }];
  const persisted = projectSchema.parse(project);
  assert.equal(persisted.decks[0].studioScene?.slides[1].resourceBindings?.[0].exactExcerptHash?.length, 64);
  assert.equal(persisted.decks[0].studioScene?.slides[1].nodes.find((node) => node.kind === "image")?.resourceBindings?.[0].resourceSha256, image.sha256);
});
