import type { DeckJob, DesignThread, StudioDesignArchetype } from "../types";
import type { DeckQualificationReport } from "./deck-qualification";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { contentProfileForSlide, selectRepresentativeSlides } from "./design-work-order";
import { analyzeStudioDeckConsistency } from "./studio-deck-consistency";
import { resolveStudioIntervention } from "./studio-intervention";
import { deckTemplateWorkflow } from "./template-routing";
import { isProtectedOrnlTemplateSlide } from "./template-guardrails";

export const AGENT_RUNBOOK_SCHEMA = "presentation-studio/agent-runbook" as const;

export interface BuildAgentRunbookInput {
  deck: DeckJob;
  projectUpdatedAt: string;
  templateInstalled: boolean;
  buildSceneRevision?: string;
  qualification?: DeckQualificationReport;
  threads?: DesignThread[];
}

export function buildAgentRunbook(input: BuildAgentRunbookInput) {
  const { deck } = input;
  const scene = deck.studioScene;
  const representatives = selectRepresentativeSlides(deck);
  const interventions = (deck.audit?.slides ?? []).map((audited) => {
    const studioSlide = scene?.slides.find((slide) => slide.slideNumber === audited.number);
    const archetype = (studioSlide?.designArchetype ?? contentProfileForSlide(deck, audited.number).designArchetype ?? "assertion-evidence") as StudioDesignArchetype;
    const intervention = resolveStudioIntervention(deck, audited.number, studioSlide, archetype);
    return { slideNumber: audited.number, archetype, status: studioSlide?.status ?? "not-created", recipe: studioSlide?.recipe, protected: isProtectedOrnlTemplateSlide(deck, audited.number), ...intervention };
  });
  const consistency = scene ? analyzeStudioDeckConsistency(scene) : undefined;
  const activeThreads = (input.threads ?? []).filter((thread) => thread.deckId === deck.id && ["submitted", "needs-reanchor"].includes(thread.status));
  const buildCurrent = Boolean(scene && input.buildSceneRevision === scene.revision);
  const qualificationCurrent = Boolean(scene && input.qualification?.sceneRevision === scene.revision);
  const undesigned = interventions.filter((item) => !item.protected && (item.status !== "designed" || item.recipe === "source" && item.level !== "preserve"));

  let nextAction: { tool: string; reason: string; input: Record<string, unknown> };
  if (!input.templateInstalled && deckTemplateWorkflow(deck) === "ornl-studio") nextAction = { tool: "get_app_status", reason: "Install or restore the authorized ORNL Template Pack before selecting layouts or building ORNL results.", input: {} };
  else if (deckTemplateWorkflow(deck) === "template-decision-required") nextAction = { tool: "set_deck_template_target", reason: "Resolve the design target before any composition work. ORNL is the default unless the user explicitly requests source-template preservation.", input: { deckId: deck.id, expectedUpdatedAt: input.projectUpdatedAt, target: "ornl-default" } };
  else if (deckTemplateWorkflow(deck) === "source-template-cleanup") {
    const slideNumber = representatives[0]?.slideNumber ?? 1;
    nextAction = { tool: "get_slide_design_work_order", reason: "This user-selected source-template workflow uses bounded source-native cleanup and Current/Proposal review rather than ORNL Studio recomposition.", input: { deckId: deck.id, slideNumber } };
  } else if (!scene || undesigned.length) {
    const slideNumber = undesigned[0]?.slideNumber ?? representatives[0]?.slideNumber ?? 1;
    nextAction = { tool: "get_studio_composition_plan", reason: `Slide ${slideNumber} is the next unqualified ${undesigned[0]?.archetype ?? "representative"} communication archetype.`, input: { deckId: deck.id, slideNumber } };
  } else if (consistency?.issues.some((issue) => issue.severity === "major")) nextAction = { tool: "get_studio_deck_consistency", reason: "Resolve major shared-grid, component, table, or archetype-pattern inconsistency before building the final deck.", input: { deckId: deck.id } };
  else if (!buildCurrent) nextAction = { tool: "build_studio_presentation", reason: "The canonical Studio scene has not been compiled into an exact current editable PowerPoint candidate.", input: { deckId: deck.id, expectedUpdatedAt: input.projectUpdatedAt } };
  else if (!qualificationCurrent) nextAction = { tool: "run_deck_qualification", reason: "The exact current candidate needs PowerPoint-native whole-deck evidence and source comparison.", input: { deckId: deck.id, expectedUpdatedAt: input.projectUpdatedAt } };
  else if (input.qualification?.status === "objective-failure" || input.qualification?.status === "revision-required" || input.qualification?.status === "held") nextAction = { tool: "get_deck_qualification", reason: "Read the routed issue ledger, repair the central scene, and create a new candidate rather than accepting a weaker result.", input: { deckId: deck.id } };
  else if (input.qualification?.status !== "review-complete") nextAction = { tool: "get_qualification_contact_sheet", reason: "Inspect every candidate overview page, then every full-size source/candidate pair and record source-wins comparisons.", input: { deckId: deck.id, representation: "candidate", page: 1 } };
  else nextAction = { tool: "get_deck_qualification", reason: "The exact candidate completed draft qualification. Leave it visible for human review; save/export remains a separate user action.", input: { deckId: deck.id } };

  return {
    schema: AGENT_RUNBOOK_SCHEMA,
    version: 1,
    designStandardVersion: PRESENTATION_DESIGN_STANDARD.version,
    projectUpdatedAt: input.projectUpdatedAt,
    deck: { id: deck.id, name: deck.name, slideCount: deck.audit?.slideCount ?? 0, sceneRevision: scene?.revision, templateWorkflow: deckTemplateWorkflow(deck) },
    operatingRules: [
      "ORNL is the default target unless the user explicitly requests another installed brand or source-template preservation.",
      "Read the source, exact content inventory, communication archetype, intervention level, and native source pixels before design.",
      "Use one canonical Studio scene and shared archetype patterns. Do not redraw every slide as an unrelated canvas.",
      "Preserve exact wording, meaning-bearing table colors, technical relationships, and source relational geometry.",
      "Source wins: reject or hold any candidate that is weaker than the source, even when objective checks pass.",
      "Build, render, remeasure, inspect, fix, and recheck original intent. Never call browser preview or successful compilation finished output.",
    ],
    representativeQualification: {
      slides: representatives,
      required: "Qualify at least one representative of every communication archetype present before propagating that pattern deck-wide.",
    },
    interventions,
    consistency,
    activeDesignThreads: activeThreads.map((thread) => ({ id: thread.id, slideNumber: thread.slideNumber, status: thread.status })),
    readiness: { buildCurrent, qualificationCurrent, qualificationStatus: input.qualification?.status ?? "not-run" },
    nextAction,
    stopConditions: PRESENTATION_DESIGN_STANDARD.askOnlyWhen,
    completionRule: "No blocker or major issue remains; every exact source/candidate pair has a raster-bound comparison; the candidate is preferred or equivalent under its intervention rule; same-archetype slides share a coherent deck system; and the result remains editable PowerPoint. This is draft qualification, not formal ORNL approval.",
  };
}
