import assert from "node:assert/strict";
import test from "node:test";
import { deriveLayoutSemantics, type LayoutContentProfile } from "../src/lib/layout-semantics";
import { inferStudioDesignArchetype, planStudioComposition } from "../src/lib/studio-archetypes";
import type { TemplateLayoutPreview, TemplatePreviewElement } from "../src/lib/template-catalog";

const WIDTH = 12_192_000;
const HEIGHT = 6_858_000;

function placeholder(type: string, index: string | undefined, x: number, y: number, width: number, height: number): TemplatePreviewElement {
  return { id: `${type}-${index ?? "default"}-${x}`, kind: "shape", name: type, x, y, width, height, rotation: 0, geometry: "rect", placeholderType: type, placeholderIndex: index };
}

function layout(id: string, name: string, category: TemplateLayoutPreview["category"], elements: TemplatePreviewElement[]): TemplateLayoutPreview {
  const base = { id, name, category, background: "#FFFFFF", elements, placeholderTypes: [...new Set(elements.map((element) => element.placeholderType).filter((value): value is string => Boolean(value)))], sourcePart: `ppt/slideLayouts/${id}.xml` };
  return { ...base, semantic: deriveLayoutSemantics(base, WIDTH, HEIGHT) };
}

function profile(patch: Partial<LayoutContentProfile> = {}): LayoutContentProfile {
  return { titleCharacterCount: 32, bodyBlockCount: 1, bodyBlockCharacterCounts: [180], captionBlockCount: 0, bodyCharacterCount: 180, imageCount: 0, tableCount: 0, chartCount: 0, mediaCount: 0, ...patch };
}

const title = placeholder("title", undefined, 300_000, 300_000, 11_000_000, 900_000);
const layouts = [
  layout("cover", "Title Slide", "title", [title, placeholder("subTitle", "1", 600_000, 4_900_000, 5_500_000, 800_000)]),
  layout("reading", "1-Column", "content", [title, placeholder("body", "1", 500_000, 1_500_000, 11_000_000, 4_700_000)]),
  layout("hero", "1-Column Key Image", "image", [title, placeholder("pic", "1", 5_600_000, 1_500_000, 5_800_000, 4_700_000), placeholder("body", "1", 500_000, 1_500_000, 4_700_000, 4_700_000)]),
  layout("four-images", "4-Image Series", "image", [
    title,
    ...[0, 1, 2, 3].flatMap((index) => [
      placeholder("pic", String(index + 1), 500_000 + index * 2_900_000, 1_500_000, 2_500_000, 2_600_000),
      placeholder("body", String(index + 1), 500_000 + index * 2_900_000, 4_300_000, 2_500_000, 600_000),
    ]),
  ]),
  layout("table", "Table | Full Width", "content", [title, placeholder("tbl", "1", 500_000, 1_500_000, 11_000_000, 4_700_000)]),
  layout("conclusion", "Conclusion", "conclusion", [title, placeholder("body", "1", 500_000, 1_800_000, 11_000_000, 3_600_000)]),
];

test("infers communication archetypes from source structure before choosing geometry", () => {
  assert.equal(inferStudioDesignArchetype(profile(), { slideNumber: 1 }).archetype, "cover");
  assert.equal(inferStudioDesignArchetype(profile({ tableCount: 1 })).archetype, "table");
  assert.equal(inferStudioDesignArchetype(profile({ imageCount: 1 })).archetype, "hero-figure");
  assert.equal(inferStudioDesignArchetype(profile(), { connectorCount: 3 }).archetype, "technical-diagram");
  assert.equal(inferStudioDesignArchetype(profile({ bodyBlockCount: 10 }), { connectorCount: 5, nativeObjectCount: 5, recommendedRecipe: "ornl-title-metric-grid" }).archetype, "comparison");
  assert.equal(inferStudioDesignArchetype(profile({ imageCount: 5, bodyBlockCount: 5, captionBlockCount: 5 }), { repeatedImageSeries: true }).archetype, "image-series");
  assert.equal(inferStudioDesignArchetype(profile({ imageCount: 24, bodyBlockCount: 0, bodyBlockCharacterCounts: [], bodyCharacterCount: 0 }), { recommendedRecipe: "ornl-title-two-column" }).archetype, "technical-diagram");
  assert.equal(inferStudioDesignArchetype(profile({ bodyCharacterCount: 949, bodyBlockCharacterCounts: [949], desiredIntent: "conclusion" }), { title: "Self-Assessment Summary" }).archetype, "text-led");
  assert.equal(inferStudioDesignArchetype(profile(), { protectedSourceComposition: true }).archetype, "source-preserve");
});

test("dense peer-logo fields stay on the relationship-preserving contained-image recipe", () => {
  const plan = planStudioComposition(profile({ imageCount: 24, bodyBlockCount: 0, bodyBlockCharacterCounts: [], bodyCharacterCount: 0, desiredIntent: "visual" }), layouts, { recommendedRecipe: "ornl-title-two-column", connectorCount: 0 });
  assert.equal(plan.archetype, "technical-diagram");
  assert.equal(plan.strategy, "shared-archetype-on-native-base");
  assert.equal(plan.recipe, "ornl-title-two-column");
  assert.match(plan.reasons.join(" "), /peer-logo field/i);
});

test("uses an exact native layout only when its complete relationship contract fits", () => {
  const coverPlan = planStudioComposition(profile({ bodyBlockCount: 0, bodyCharacterCount: 0, desiredIntent: "cover" }), layouts, { slideNumber: 1 });
  assert.equal(coverPlan.strategy, "converted-template-layout");
  assert.equal(coverPlan.recipe, "template-layout");
  assert.equal(coverPlan.layoutId, "cover");

  const fourImagePlan = planStudioComposition(profile({ bodyBlockCount: 0, bodyBlockCharacterCounts: [], bodyCharacterCount: 0, imageCount: 4, captionBlockCount: 4, desiredIntent: "visual" }), layouts, { repeatedImageSeries: true });
  assert.equal(fourImagePlan.archetype, "image-series");
  assert.equal(fourImagePlan.strategy, "converted-template-layout");
  assert.equal(fourImagePlan.layoutId, "four-images");

  const fiveImagePlan = planStudioComposition(profile({ bodyBlockCount: 0, bodyBlockCharacterCounts: [], bodyCharacterCount: 0, imageCount: 5, captionBlockCount: 5, desiredIntent: "visual" }), layouts, { repeatedImageSeries: true });
  assert.equal(fiveImagePlan.strategy, "shared-archetype-on-native-base");
  assert.equal(fiveImagePlan.recipe, "ornl-title-image-series");
  assert.equal(fiveImagePlan.layoutId, undefined);
});

test("does not force the strict image-series recipe when source relationships are only inferred from counts", () => {
  const ambiguousSeriesPlan = planStudioComposition(profile({
    bodyBlockCount: 5,
    bodyBlockCharacterCounts: [80, 80, 80, 80, 80],
    bodyCharacterCount: 400,
    imageCount: 5,
    captionBlockCount: 5,
    desiredIntent: "visual",
  }), layouts, { repeatedImageSeries: false });
  assert.equal(ambiguousSeriesPlan.archetype, "image-series");
  assert.equal(ambiguousSeriesPlan.strategy, "shared-archetype-on-native-base");
  assert.equal(ambiguousSeriesPlan.recipe, "ornl-title-labeled-figure-grid");
  assert.match(ambiguousSeriesPlan.reasons.join(" "), /does not prove complete image-heading-evidence groups/i);

  const connectedSeriesPlan = planStudioComposition(profile({
    bodyBlockCount: 5,
    bodyBlockCharacterCounts: [80, 80, 80, 80, 80],
    bodyCharacterCount: 400,
    imageCount: 5,
    captionBlockCount: 5,
    desiredIntent: "visual",
  }), layouts, { repeatedImageSeries: false, connectorCount: 1 });
  assert.equal(connectedSeriesPlan.recipe, "ornl-title-figure-grid");
  assert.match(connectedSeriesPlan.reasons.join(" "), /connector carries visual relationships/i);
});

test("routes tables and technical diagrams to controlled shared archetypes instead of forcing a named layout", () => {
  const tablePlan = planStudioComposition(profile({ bodyBlockCount: 0, bodyBlockCharacterCounts: [], bodyCharacterCount: 0, tableCount: 1, desiredIntent: "data" }), layouts);
  assert.equal(tablePlan.archetype, "table");
  assert.equal(tablePlan.strategy, "shared-archetype-on-native-base");
  assert.equal(tablePlan.recipe, "ornl-title-table");

  const diagramPlan = planStudioComposition(profile(), layouts, { connectorCount: 4 });
  assert.equal(diagramPlan.archetype, "technical-diagram");
  assert.equal(diagramPlan.strategy, "shared-archetype-on-native-base");
  assert.equal(diagramPlan.recipe, "ornl-title-two-column");
  assert.match(diagramPlan.requiredChecks.join(" "), /connector/i);

  const connectedFigurePlan = planStudioComposition(profile({ imageCount: 2, imageAspectRatios: [1.2, 1.4] }), layouts, { connectorCount: 2 });
  assert.equal(connectedFigurePlan.archetype, "technical-diagram");
  assert.equal(connectedFigurePlan.recipe, "ornl-title-figure-grid");
  assert.match(connectedFigurePlan.reasons.join(" "), /relationship-bearing technical field/i);
});

test("uses detailed source structure to choose a compatible recipe inside the broader archetype", () => {
  const figureComparison = planStudioComposition(profile({
    bodyBlockCount: 1,
    bodyBlockCharacterCounts: [610],
    bodyCharacterCount: 610,
    captionBlockCount: 1,
    imageCount: 2,
    imageAspectRatios: [1.78, 1.59],
    desiredIntent: "visual",
  }), layouts, { recommendedRecipe: "ornl-title-figure-grid" });
  assert.equal(figureComparison.archetype, "comparison");
  assert.equal(figureComparison.strategy, "shared-archetype-on-native-base");
  assert.equal(figureComparison.recipe, "ornl-title-figure-grid");
  assert.match(figureComparison.reasons.join(" "), /source structure supports/i);
});

test("routes a dense editorial record grid around legacy native carrier furniture", () => {
  const denseRecords = profile({
    bodyBlockCount: 8,
    bodyBlockCharacterCounts: Array.from({ length: 8 }, () => 180),
    bodyCharacterCount: 1_440,
    imageCount: 0,
    desiredIntent: "comparison",
  });
  const plan = planStudioComposition(denseRecords, layouts, { nativeObjectCount: 1, connectorCount: 0, recommendedRecipe: "ornl-title-card-grid" });
  assert.equal(plan.archetype, "comparison");
  assert.equal(plan.strategy, "shared-archetype-on-native-base");
  assert.equal(plan.recipe, "ornl-title-card-grid");
  assert.match(plan.reasons.join(" "), /editorial records/i);
  assert.match(plan.requiredChecks.join(" "), /10\.5 pt record-grid exception/i);
});
