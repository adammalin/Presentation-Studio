import type { DeckJob, StudioWebScene } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";

const ORNL_TEMPLATE_CLASSIFICATIONS = new Set(["current-ornl", "older-or-modified-ornl", "mixed"]);

type OrnlTemplateDeck = Pick<DeckJob, "targetTemplateId" | "templateClassification" | "audit" | "protectedSlideNumbers">;

function usesOrnlTemplate(deck: Pick<DeckJob, "targetTemplateId" | "templateClassification">): boolean {
  return deck.targetTemplateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id
    && ORNL_TEMPLATE_CLASSIFICATIONS.has(deck.templateClassification);
}

/**
 * The populated opening title slide in an ORNL deck is approved brand
 * composition, not raw content for the redesign engine. Existing artwork,
 * marks, photography, legal copy, and geometry therefore remain source locked.
 */
export function isSacredOrnlTitleSlide(deck: Pick<DeckJob, "targetTemplateId" | "templateClassification">, slideNumber: number): boolean {
  return slideNumber === 1 && usesOrnlTemplate(deck);
}

/**
 * A populated ORNL closing slide is also approved template composition. Keep
 * this deliberately narrow: only the final, text-only Thank you slide is
 * protected automatically. Content-heavy conclusions remain designable.
 */
export function isSacredOrnlClosingSlide(deck: OrnlTemplateDeck, slideNumber: number): boolean {
  if (!usesOrnlTemplate(deck) || !deck.audit?.slides.length) return false;
  const finalSlide = deck.audit.slides.at(-1);
  if (!finalSlide || slideNumber !== finalSlide.number) return false;
  const normalizedText = finalSlide.text.trim().toLowerCase().replace(/[.!]+$/g, "").replace(/\s+/g, " ");
  return normalizedText === "thank you";
}

export function isProtectedOrnlTemplateSlide(deck: OrnlTemplateDeck, slideNumber: number): boolean {
  return deck.targetTemplateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id
    && (deck.protectedSlideNumbers?.includes(slideNumber) || isSacredOrnlTitleSlide(deck, slideNumber) || isSacredOrnlClosingSlide(deck, slideNumber));
}

export function assertSacredOrnlTitleSlideIntegrity(deck: OrnlTemplateDeck, scene: StudioWebScene): void {
  const protectedSlides = scene.slides.filter((slide) => isProtectedOrnlTemplateSlide(deck, slide.slideNumber));
  const changed = protectedSlides.find((slide) => isSacredOrnlTitleSlide(deck, slide.slideNumber) || isSacredOrnlClosingSlide(deck, slide.slideNumber)
    ? slide.recipe !== "source"
    : slide.recipe !== "template-layout" || !slide.targetLayoutId);
  if (changed) {
    throw new Error(`The approved ORNL template composition on slide ${changed.slideNumber} is sacred. Restore its approved source or converted-template layout; do not recompose, restyle, move, resize, or replace its template artwork.`);
  }
}

export function unsupportedSourceSlideNumbers(deck: OrnlTemplateDeck, scene: StudioWebScene): number[] {
  return scene.slides
    .filter((slide) => (slide.status !== "designed" || slide.recipe === "source") && !isProtectedOrnlTemplateSlide(deck, slide.slideNumber))
    .map((slide) => slide.slideNumber);
}
