import assert from "node:assert/strict";
import test from "node:test";
import { assertSacredOrnlTitleSlideIntegrity, isSacredOrnlTitleSlide, unsupportedSourceSlideNumbers } from "../src/lib/template-guardrails";
import type { StudioWebScene } from "../src/types";

const ornlDeck = { targetTemplateId: "ornl-16x9-v1", templateClassification: "current-ornl" as const };

function scene(recipe: StudioWebScene["slides"][number]["recipe"]): StudioWebScene {
  return {
    schema: "presentation-studio/web-scene",
    version: 5,
    revision: "guardrail-test",
    deckId: "deck",
    sourceSha256: "a".repeat(64),
    slideSize: { width: 12_192_000, height: 6_858_000 },
    sourceSlideSize: { width: 12_192_000, height: 6_858_000 },
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: "test", unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides: [{ id: "studio-slide-1", slideNumber: 1, sourceSlideId: "slide-1", sourceTextHash: "b".repeat(64), sourceRevision: "source", contentCoverage: { exactTextMapped: true, sourceCharacterCount: 0, mappedCharacterCount: 0, sourceTextBoxCount: 0, mappedTextNodeCount: 0, groupedOrUnsupportedTextPresent: false }, recipe, background: "#FFFFFF", status: recipe === "source" ? "imported" : "designed", designRationale: "test", figureTreatments: [], nodes: [], updatedAt: "2026-08-13T20:00:00.000Z" }],
  };
}

test("existing ORNL title slides are sacred source-preserved compositions", () => {
  assert.equal(isSacredOrnlTitleSlide(ornlDeck, 1), true);
  assert.equal(isSacredOrnlTitleSlide({ ...ornlDeck, templateClassification: "mixed" }, 1), true);
  assert.equal(isSacredOrnlTitleSlide(ornlDeck, 2), false);
  assert.equal(isSacredOrnlTitleSlide({ ...ornlDeck, templateClassification: "sponsor" }, 1), false);
  assert.doesNotThrow(() => assertSacredOrnlTitleSlideIntegrity(ornlDeck, scene("source")));
  assert.throws(() => assertSacredOrnlTitleSlideIntegrity(ornlDeck, scene("template-layout")), /title slide is sacred/i);
  assert.deepEqual(unsupportedSourceSlideNumbers(ornlDeck, scene("source")), []);
});
