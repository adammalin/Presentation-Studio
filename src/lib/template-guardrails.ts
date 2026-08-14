import type { DeckJob, StudioWebScene } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";

const ORNL_TEMPLATE_CLASSIFICATIONS = new Set(["current-ornl", "older-or-modified-ornl", "mixed"]);

/**
 * The populated opening title slide in an ORNL deck is approved brand
 * composition, not raw content for the redesign engine. Existing artwork,
 * marks, photography, legal copy, and geometry therefore remain source locked.
 */
export function isSacredOrnlTitleSlide(deck: Pick<DeckJob, "targetTemplateId" | "templateClassification">, slideNumber: number): boolean {
  return slideNumber === 1
    && deck.targetTemplateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id
    && ORNL_TEMPLATE_CLASSIFICATIONS.has(deck.templateClassification);
}

export function assertSacredOrnlTitleSlideIntegrity(deck: Pick<DeckJob, "targetTemplateId" | "templateClassification">, scene: StudioWebScene): void {
  const titleSlide = scene.slides.find((slide) => isSacredOrnlTitleSlide(deck, slide.slideNumber));
  if (titleSlide && titleSlide.recipe !== "source") {
    throw new Error("The existing ORNL title slide is sacred and must remain source-preserved. Restore slide 1 to Source geometry; do not recompose, restyle, move, resize, or replace its template artwork.");
  }
}

export function unsupportedSourceSlideNumbers(deck: Pick<DeckJob, "targetTemplateId" | "templateClassification">, scene: StudioWebScene): number[] {
  return scene.slides
    .filter((slide) => (slide.status !== "designed" || slide.recipe === "source") && !isSacredOrnlTitleSlide(deck, slide.slideNumber))
    .map((slide) => slide.slideNumber);
}
