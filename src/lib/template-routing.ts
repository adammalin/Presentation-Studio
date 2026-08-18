import type { DeckJob, PptxAudit, TemplateClassification } from "../types";
import { createOrnlDesignProfile, PRESENTATION_DESIGN_STANDARD } from "./design-standard";

export type DeckTemplateWorkflow = "ornl-studio" | "source-template-cleanup" | "template-decision-required";

export function deckTemplateWorkflow(deck: Pick<DeckJob, "templateClassification" | "targetTemplateId" | "targetTemplateDecisionSource">): DeckTemplateWorkflow {
  if (deck.targetTemplateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id) return "ornl-studio";
  if (deck.targetTemplateDecisionSource === "automatic-source-preservation" || ["sponsor-source", "custom-source"].includes(deck.targetTemplateId ?? "")) return "source-template-cleanup";
  if (["sponsor", "custom"].includes(deck.templateClassification)) return "source-template-cleanup";
  if (["current-ornl", "older-or-modified-ornl"].includes(deck.templateClassification)) return "ornl-studio";
  return "template-decision-required";
}

function automaticTarget(classification: TemplateClassification, adoptedAt: string): Partial<DeckJob> {
  if (["current-ornl", "older-or-modified-ornl"].includes(classification)) return {
    targetTemplateId: PRESENTATION_DESIGN_STANDARD.defaults.template.id,
    targetTemplateConfirmedAt: adoptedAt,
    targetTemplateDecisionSource: "automatic-default",
    designProfile: createOrnlDesignProfile("automatic-default", adoptedAt),
    status: "ready-for-cleanup",
  };
  if (classification === "sponsor" || classification === "custom") return {
    targetTemplateId: classification === "sponsor" ? "sponsor-source" : "custom-source",
    targetTemplateConfirmedAt: adoptedAt,
    targetTemplateDecisionSource: "automatic-source-preservation",
    designProfile: undefined,
    status: "audited",
  };
  return {
    targetTemplateId: undefined,
    targetTemplateConfirmedAt: undefined,
    targetTemplateDecisionSource: undefined,
    designProfile: undefined,
    status: "needs-template-decision",
  };
}

export function deckWithAutomaticTemplateRouting(input: {
  deck: DeckJob;
  audit: PptxAudit;
  adoptedAt?: string;
}): DeckJob {
  const adoptedAt = input.adoptedAt ?? new Date().toISOString();
  const previousWorkflow = deckTemplateWorkflow(input.deck);
  if (input.deck.targetTemplateDecisionSource === "user-selected") return {
    ...input.deck,
    audit: input.audit,
    templateClassification: input.audit.classification,
  };
  const routed = { ...input.deck, ...automaticTarget(input.audit.classification, adoptedAt), audit: input.audit, templateClassification: input.audit.classification };
  const workflowChanged = previousWorkflow !== deckTemplateWorkflow(routed);
  return workflowChanged ? { ...routed, studioScene: undefined, proposal: undefined } : routed;
}
