import assert from "node:assert/strict";
import test from "node:test";
import { deriveLayoutSemantics, rankLayoutCompatibility, uniquePlaceholderElements, type LayoutContentProfile } from "../src/lib/layout-semantics";
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

test("deduplicates inherited placeholders and derives stable semantic slots", () => {
  const image = placeholder("pic", "12", 6_000_000, 1_500_000, 4_000_000, 3_000_000);
  const elements = [placeholder("title", undefined, 300_000, 300_000, 11_000_000, 900_000), image, { ...image, id: "duplicate-image" }];
  assert.equal(uniquePlaceholderElements(elements).length, 2);
  const preview = layout("layout-visual", "1-Column Key Image", "image", elements);
  assert.equal(preview.semantic?.intent, "visual");
  assert.equal(preview.semantic?.capabilities.imageSlots, 1);
  assert.equal(preview.semantic?.constraints.requiresVisual, true);
  assert.equal(preview.semantic?.contract.family, "hero-visual");
  assert.equal(preview.semantic?.contract.selectionPolicy, "special-purpose");
  assert.equal(preview.semantic?.contract.surface, "light");
  assert.deepEqual(preview.semantic?.contract.nativeAuthority, { masterRequired: true, layoutRequired: true, preserveInheritedArtwork: true, inheritedArtworkCount: 0, footerArtworkExpected: false });
  assert.ok(preview.semantic?.contract.compatibleArchetypes.includes("hero-figure"));
  const titleSlot = preview.semantic?.slots.find((slot) => slot.role === "title");
  const imageSlot = preview.semantic?.slots.find((slot) => slot.role === "image");
  assert.deepEqual(titleSlot?.preferredBounds, { x: 300_000, y: 300_000, width: 11_000_000, height: 900_000 });
  assert.ok((titleSlot?.minimumBounds.width ?? 0) < (titleSlot?.preferredBounds.width ?? 0));
  assert.deepEqual(titleSlot?.maximumBounds, titleSlot?.preferredBounds);
  assert.equal(titleSlot?.alignmentIntent, "optical-left");
  assert.equal(titleSlot?.priority, 100);
  assert.deepEqual(imageSlot?.allowedObjectKinds, ["image"]);
  assert.equal(imageSlot?.alignmentIntent, "contain");
  assert.equal(imageSlot?.priority, 80);
});

test("ranks content, visual, and data layouts from exact-content needs", () => {
  const title = placeholder("title", undefined, 300_000, 300_000, 11_000_000, 900_000);
  const candidates = [
    layout("layout-content", "1-Column", "content", [title, placeholder("body", "1", 300_000, 1_500_000, 11_000_000, 4_600_000)]),
    layout("layout-images", "3-Image Series", "image", [title, placeholder("pic", "1", 400_000, 1_500_000, 3_400_000, 3_000_000), placeholder("pic", "2", 4_400_000, 1_500_000, 3_400_000, 3_000_000), placeholder("pic", "3", 8_400_000, 1_500_000, 3_400_000, 3_000_000)]),
    layout("layout-table", "Table | Full Width", "content", [title, placeholder("tbl", "1", 500_000, 1_500_000, 11_000_000, 4_700_000)]),
  ];
  const visualProfile: LayoutContentProfile = { titleCharacterCount: 30, bodyBlockCount: 0, captionBlockCount: 0, bodyCharacterCount: 0, imageCount: 3, tableCount: 0, chartCount: 0, mediaCount: 0, desiredIntent: "visual" };
  assert.equal(rankLayoutCompatibility(candidates, visualProfile)[0].layoutId, "layout-images");
  const oneImage = layout("layout-one-image", "1-Column Key Image", "image", [title, placeholder("pic", "1", 4_400_000, 1_500_000, 7_000_000, 4_500_000)]);
  const oneImageProfile: LayoutContentProfile = { ...visualProfile, imageCount: 1 };
  assert.equal(rankLayoutCompatibility([...candidates, oneImage], oneImageProfile)[0].layoutId, "layout-one-image");
  const dataProfile: LayoutContentProfile = { titleCharacterCount: 30, bodyBlockCount: 0, captionBlockCount: 0, bodyCharacterCount: 0, imageCount: 0, tableCount: 1, chartCount: 0, mediaCount: 0, desiredIntent: "data" };
  assert.equal(rankLayoutCompatibility(candidates, dataProfile)[0].layoutId, "layout-table");
  const denseProfile: LayoutContentProfile = { titleCharacterCount: 30, bodyBlockCount: 1, captionBlockCount: 0, bodyCharacterCount: 1_100, imageCount: 0, tableCount: 0, chartCount: 0, mediaCount: 0, desiredIntent: "content" };
  assert.equal(rankLayoutCompatibility(candidates, denseProfile)[0].layoutId, "layout-content");
  const shortStack = layout("layout-short-stack", "2-Image Short Stack", "image", [title, placeholder("body", "1", 300_000, 1_500_000, 1_700_000, 1_400_000), placeholder("pic", "1", 2_200_000, 1_500_000, 4_000_000, 1_400_000), placeholder("pic", "2", 6_500_000, 3_200_000, 4_000_000, 1_400_000)]);
  const longVisual = layout("layout-long-visual", "1-Column Stacked Image Series", "image", [title, placeholder("body", "1", 300_000, 1_500_000, 5_500_000, 4_400_000), placeholder("pic", "1", 6_200_000, 1_500_000, 2_700_000, 2_000_000), placeholder("pic", "2", 9_100_000, 1_500_000, 2_700_000, 2_000_000)]);
  const unsplitLongBlock: LayoutContentProfile = { titleCharacterCount: 30, bodyBlockCount: 1, bodyBlockCharacterCounts: [360], captionBlockCount: 0, bodyCharacterCount: 360, imageCount: 2, tableCount: 0, chartCount: 0, mediaCount: 0, desiredIntent: "visual" };
  assert.equal(rankLayoutCompatibility([shortStack, longVisual], unsplitLongBlock)[0].layoutId, "layout-long-visual");
  const oneBodyOneImage = layout("layout-shared-slot", "1-Column Key Image", "image", [title, placeholder("body", "1", 300_000, 1_500_000, 5_500_000, 4_400_000), placeholder("pic", "1", 6_200_000, 1_500_000, 5_500_000, 4_400_000)]);
  assert.equal(rankLayoutCompatibility([oneBodyOneImage, longVisual], unsplitLongBlock)[0].layoutId, "layout-long-visual", "a body slot cannot be counted once for exact text and again as fallback image capacity");
});

test("builds repeated-slot relationships and keeps sacred layouts out of ordinary content routing", () => {
  const title = placeholder("title", undefined, 300_000, 300_000, 11_000_000, 900_000);
  const imageSeries = layout("layout-images", "3-Image Series", "image", [
    title,
    placeholder("pic", "1", 400_000, 1_500_000, 3_400_000, 2_700_000),
    placeholder("pic", "2", 4_400_000, 1_500_000, 3_400_000, 2_700_000),
    placeholder("pic", "3", 8_400_000, 1_500_000, 3_400_000, 2_700_000),
    placeholder("body", "1", 400_000, 4_350_000, 3_400_000, 650_000),
    placeholder("body", "2", 4_400_000, 4_350_000, 3_400_000, 650_000),
    placeholder("body", "3", 8_400_000, 4_350_000, 3_400_000, 650_000),
  ]);
  assert.equal(imageSeries.semantic?.contract.family, "image-series");
  assert.equal(imageSeries.semantic?.contract.slotGroups.length, 3);
  assert.ok(imageSeries.semantic?.contract.slotGroups.every((group) => group.kind === "image-evidence" && group.relationship === "image-heading-evidence" && group.slotIds.length === 2));
  const stackedSeries = layout("layout-stacked", "1-Column Stacked Image Series", "image", [
    title,
    placeholder("body", "1", 400_000, 1_500_000, 4_000_000, 4_400_000),
    placeholder("pic", "1", 5_000_000, 1_500_000, 2_000_000, 1_200_000),
    placeholder("pic", "2", 7_300_000, 2_900_000, 2_000_000, 1_200_000),
    placeholder("pic", "3", 9_600_000, 4_300_000, 2_000_000, 1_200_000),
  ]);
  assert.equal(stackedSeries.semantic?.contract.slotGroups.length, 1);
  assert.equal(stackedSeries.semantic?.contract.slotGroups[0]?.kind, "image-series");
  assert.equal(stackedSeries.semantic?.contract.slotGroups[0]?.relationship, "image-collection-evidence");
  assert.equal(stackedSeries.semantic?.contract.slotGroups[0]?.slotIds.length, 4);

  const cover = layout("layout-cover", "Title Slide", "title", [title, placeholder("subTitle", "1", 700_000, 4_900_000, 5_000_000, 800_000)]);
  assert.equal(cover.semantic?.contract.selectionPolicy, "sacred");
  const tableProfile: LayoutContentProfile = { titleCharacterCount: 30, bodyBlockCount: 0, captionBlockCount: 0, bodyCharacterCount: 0, imageCount: 0, tableCount: 1, chartCount: 0, mediaCount: 0, desiredIntent: "data", designArchetype: "table" };
  const [tableResult, coverResult] = rankLayoutCompatibility([
    layout("layout-table", "Table | Full Width", "content", [title, placeholder("tbl", "1", 500_000, 1_500_000, 11_000_000, 4_700_000)]),
    cover,
  ], tableProfile);
  assert.equal(tableResult.layoutId, "layout-table");
  assert.equal(coverResult.status, "incompatible");
  assert.match(coverResult.unmetNeeds.join(" "), /Sacred cover layout/i);
});
