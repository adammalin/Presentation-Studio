import type { DeckJob, StudioDesignArchetype, StudioInterventionDecision, StudioInterventionLevel, StudioWebSlide } from "../types";
import { isProtectedOrnlTemplateSlide } from "./template-guardrails";
import { deckTemplateWorkflow } from "./template-routing";

export const STUDIO_INTERVENTION_LEVELS: readonly StudioInterventionLevel[] = ["preserve", "polish", "recompose", "rebuild-figure"];

function decision(level: StudioInterventionLevel, reason: string, selectedBy: StudioInterventionDecision["selectedBy"], selectedAt: string): StudioInterventionDecision {
  const acceptanceRule = level === "preserve"
    ? "The source is the approved result. The candidate must remain visually identical unless a person explicitly changes the intervention level."
    : level === "polish"
      ? "Accept only bounded improvements to alignment, fit, typography, spacing, table consistency, or brand fidelity. If the candidate weakens the source composition, keep the source."
      : level === "rebuild-figure"
        ? "Accept only when the complete figure information inventory and every verified relationship are preserved and the candidate is materially clearer than the source. Otherwise keep the source figure as one evidence unit."
        : "Accept only when the whole-slide composition is materially stronger than the source while preserving exact content, technical meaning, and relationships. Otherwise keep the source.";
  return { level, selectedBy, reason, sourceWins: true, acceptanceRule, selectedAt };
}

export function automaticStudioIntervention(deck: DeckJob, slideNumber: number, studioSlide?: StudioWebSlide, archetype?: StudioDesignArchetype, now = new Date().toISOString()): StudioInterventionDecision {
  if (isProtectedOrnlTemplateSlide(deck, slideNumber)) return decision("preserve", "This is an approved protected ORNL template composition.", "automatic", now);
  const verifiedFigureRebuild = studioSlide?.figureTreatments.some((treatment) => ["hybrid-rebuild", "redraw-candidate"].includes(treatment.mode) && treatment.verificationStatus === "verified");
  if (verifiedFigureRebuild) return decision("rebuild-figure", "A verified relationship inventory authorizes a bounded editable figure reconstruction.", "automatic", now);
  if (deckTemplateWorkflow(deck) === "source-template-cleanup" || ["audit-only", "cleanup-only"].includes(deck.operationScope)) {
    return decision("polish", "The deck is scoped to restrained source-template cleanup rather than whole-slide recomposition.", "automatic", now);
  }
  if (archetype === "source-preserve") return decision("preserve", "The communication archetype is explicitly source-preserve.", "automatic", now);
  if (studioSlide?.recipe === "source" && studioSlide.status === "designed") return decision("polish", "The current approved composition remains the baseline; only bounded refinements are warranted.", "automatic", now);
  return decision("recompose", "The slide is in the ORNL Studio workflow and requires a deliberate shared-layout composition decision.", "automatic", now);
}

export function resolveStudioIntervention(deck: DeckJob, slideNumber: number, studioSlide?: StudioWebSlide, archetype?: StudioDesignArchetype): StudioInterventionDecision {
  return studioSlide?.intervention ?? automaticStudioIntervention(deck, slideNumber, studioSlide, archetype);
}

export function withStudioIntervention(
  slide: StudioWebSlide,
  level: StudioInterventionLevel,
  reason: string,
  selectedBy: StudioInterventionDecision["selectedBy"],
  selectedAt = new Date().toISOString(),
): StudioWebSlide {
  return { ...slide, intervention: decision(level, reason, selectedBy, selectedAt) };
}
