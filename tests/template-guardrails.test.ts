import assert from "node:assert/strict";
import test from "node:test";
import { assertSacredOrnlTitleSlideIntegrity, isProtectedOrnlTemplateSlide, isSacredOrnlClosingSlide, isSacredOrnlTitleSlide, markNativeQualifiedConvertedOrnlTitle, unsupportedSourceSlideNumbers } from "../src/lib/template-guardrails";
import type { DeckJob, StudioWebScene } from "../src/types";

const ornlDeck = { targetTemplateId: "ornl-16x9-v1", templateClassification: "current-ornl" as const, protectedSlideNumbers: [] as number[] };

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
  assert.throws(() => assertSacredOrnlTitleSlideIntegrity(ornlDeck, scene("template-layout")), /template composition.*sacred/i);
  assert.deepEqual(unsupportedSourceSlideNumbers(ornlDeck, scene("source")), []);
});

test("a final text-only ORNL Thank you slide remains approved source template composition", () => {
  const audit = { slides: [
    { number: 1, text: "Presentation title" },
    { number: 13, text: "Thank you" },
  ] } as DeckJob["audit"];
  const deck = { ...ornlDeck, audit };
  assert.equal(isSacredOrnlClosingSlide(deck, 13), true);
  assert.equal(isProtectedOrnlTemplateSlide(deck, 13), true);
  assert.equal(isSacredOrnlClosingSlide(deck, 12), false);
  assert.equal(isSacredOrnlClosingSlide({ ...deck, audit: { ...audit!, slides: [...audit!.slides.slice(0, -1), { ...audit!.slides.at(-1)!, text: "Conclusions and next steps" }] } }, 13), false);
});

test("a converted non-ORNL title becomes sacred only after its exact native-qualified revision is recorded", () => {
  const deck = { targetTemplateId: "ornl-16x9-v1", templateClassification: "sponsor" as const, protectedSlideNumbers: [1] };
  const converted = scene("template-layout");
  converted.slides[0].targetLayoutId = "layout-1";
  converted.slides[0].targetLayoutName = "Title | Standard";
  assert.equal(isProtectedOrnlTemplateSlide({ ...deck, studioScene: converted }, 1), false);
  assert.doesNotThrow(() => assertSacredOrnlTitleSlideIntegrity({ ...deck, studioScene: converted }, converted));
  const qualified = markNativeQualifiedConvertedOrnlTitle({ ...deck, studioScene: converted } as DeckJob, converted);
  assert.equal(isProtectedOrnlTemplateSlide(qualified, 1), true);
  assert.doesNotThrow(() => assertSacredOrnlTitleSlideIntegrity(qualified, converted));
  const changed = scene("ornl-title-content");
  assert.equal(isProtectedOrnlTemplateSlide({ ...qualified, studioScene: changed }, 1), false);
});
