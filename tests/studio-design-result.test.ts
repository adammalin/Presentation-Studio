import assert from "node:assert/strict";
import test from "node:test";
import type { NativeRenderResult } from "../src/lib/desktop";
import { composeLatestStudioNativeRender, type StudioSlideBuildResult } from "../src/lib/studio-design-result";
import { STUDIO_WEB_SCENE_SCHEMA, STUDIO_WEB_SCENE_VERSION, type StudioWebScene, type StudioWebSlide } from "../src/types";

function native(numbers: number[]): NativeRenderResult {
  return {
    status: "ready",
    renderer: "powerpoint-native",
    authoritative: true,
    slides: numbers.map((number) => ({ number, mimeType: "image/png", width: 1600, height: 900, sha256: `slide-${number}`, bytes: new Uint8Array([number]) })),
    warnings: [],
  };
}

function scene(): StudioWebScene {
  const slide = (slideNumber: number, recipe: StudioWebSlide["recipe"], updatedAt: string): StudioWebSlide => ({
    id: `studio-slide-${slideNumber}`,
    slideNumber,
    sourceSlideId: `slide-${slideNumber}`,
    sourceTextHash: `text-${slideNumber}`,
    contentCoverage: { exactTextMapped: true, sourceCharacterCount: 1, mappedCharacterCount: 1, sourceTextBoxCount: 1, mappedTextNodeCount: 1, groupedOrUnsupportedTextPresent: false },
    sourceRevision: "source",
    recipe,
    background: "#FFFFFF",
    status: recipe === "source" ? "imported" : "designed",
    designRationale: "test",
    figureTreatments: [],
    nodes: [],
    updatedAt,
  });
  return {
    schema: STUDIO_WEB_SCENE_SCHEMA,
    version: STUDIO_WEB_SCENE_VERSION,
    revision: "scene-revision",
    deckId: "deck-1",
    sourceSha256: "source",
    slideSize: { width: 12_192_000, height: 6_858_000 },
    sourceSlideSize: { width: 12_192_000, height: 6_858_000 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [slide(1, "source", "one"), slide(2, "ornl-title-content", "two")],
  };
}

test("latest Studio render preserves source slides and replaces designed slides with their exact build", () => {
  const builds: Record<string, StudioSlideBuildResult> = {
    "deck-1:2": { deckId: "deck-1", sourceSlideNumber: 2, slideUpdatedAt: "two", nativeRender: native([1]) },
  };
  const result = composeLatestStudioNativeRender(scene(), native([1, 2]), builds);
  assert.deepEqual(result?.slides.map((slide) => [slide.number, slide.sha256]), [[1, "slide-1"], [2, "slide-1"]]);
  assert.equal(result?.reason, undefined);
});

test("latest Studio render holds a redesigned slide when only a stale build exists", () => {
  const builds: Record<string, StudioSlideBuildResult> = {
    "deck-1:2": { deckId: "deck-1", sourceSlideNumber: 2, slideUpdatedAt: "old", nativeRender: native([1]) },
  };
  const result = composeLatestStudioNativeRender(scene(), native([1, 2]), builds);
  assert.deepEqual(result?.slides.map((slide) => slide.number), [1]);
  assert.match(result?.reason ?? "", /not been built/i);
  assert.match(result?.warnings.at(-1) ?? "", /slide 2/i);
});
