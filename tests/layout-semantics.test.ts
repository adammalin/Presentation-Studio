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
});
