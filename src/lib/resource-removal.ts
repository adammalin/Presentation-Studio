import type { PresentationStudioProject, ProjectResource } from "../types";
import { touchProject } from "./project";

export interface ResourceRemovalImpact {
  resource: ProjectResource;
  linkedDeckIds: string[];
  linkedDeckNames: string[];
  removedExemplarCount: number;
  removedThreadCount: number;
  removedConceptReferenceCount: number;
}

export function resourceRemovalImpact(project: PresentationStudioProject, resourceId: string): ResourceRemovalImpact {
  const resource = project.resources.find((item) => item.id === resourceId);
  if (!resource) throw new Error("The selected Resource is no longer present in this project.");
  const linkedDecks = project.decks.filter((deck) => deck.sourceResourceId === resourceId);
  const linkedDeckIds = new Set(linkedDecks.map((deck) => deck.id));
  return {
    resource,
    linkedDeckIds: [...linkedDeckIds],
    linkedDeckNames: linkedDecks.map((deck) => deck.name),
    removedExemplarCount: project.styleExemplars.filter((item) => item.resourceId === resourceId || linkedDeckIds.has(item.deckId)).length,
    removedThreadCount: project.designThreads.filter((thread) => linkedDeckIds.has(thread.deckId)).length,
    removedConceptReferenceCount: project.decks.filter((deck) => !linkedDeckIds.has(deck.id)).reduce((count, deck) => count + (deck.studioScene?.slides.reduce((slideCount, slide) => slideCount + (slide.conceptReferences?.filter((reference) => reference.resourceId === resourceId).length ?? 0), 0) ?? 0), 0),
  };
}

export function removeResourceFromProject(project: PresentationStudioProject, resourceId: string): { project: PresentationStudioProject; impact: ResourceRemovalImpact } {
  const impact = resourceRemovalImpact(project, resourceId);
  const linkedDeckIds = new Set(impact.linkedDeckIds);
  const removedAt = new Date().toISOString();
  const next = touchProject({
    ...project,
    resources: project.resources.filter((resource) => resource.id !== resourceId),
    decks: project.decks.filter((deck) => !linkedDeckIds.has(deck.id)).map((deck) => {
      if (!deck.studioScene?.slides.some((slide) => slide.conceptReferences?.some((reference) => reference.resourceId === resourceId))) return deck;
      return {
        ...deck,
        studioScene: {
          ...deck.studioScene,
          revision: `${deck.studioScene.sourceSha256}:web-v${deck.studioScene.version}:${removedAt}`,
          slides: deck.studioScene.slides.map((slide) => {
            const removedReferenceIds = new Set(slide.conceptReferences?.filter((reference) => reference.resourceId === resourceId).map((reference) => reference.id) ?? []);
            if (!removedReferenceIds.size) return slide;
            return {
              ...slide,
              conceptReferences: (slide.conceptReferences ?? []).filter((reference) => reference.resourceId !== resourceId),
              visualNeeds: (slide.visualNeeds ?? []).map((need) => need.linkedConceptReferenceId && removedReferenceIds.has(need.linkedConceptReferenceId) ? {
                ...need,
                status: "brief-ready" as const,
                linkedConceptReferenceId: undefined,
                resolutionNote: "The linked concept Resource was removed from the project; the brief is ready for a new concept.",
                updatedAt: removedAt,
              } : need),
              qualityReview: undefined,
              updatedAt: removedAt,
            };
          }),
        },
      };
    }),
    styleExemplars: project.styleExemplars.filter((item) => item.resourceId !== resourceId && !linkedDeckIds.has(item.deckId)),
    designThreads: project.designThreads.filter((thread) => !linkedDeckIds.has(thread.deckId)),
  }, "resource-removed", `Removed the embedded Resource ${impact.resource.name}${impact.linkedDeckIds.length > 0 ? ` and ${impact.linkedDeckIds.length} linked deck${impact.linkedDeckIds.length === 1 ? "" : "s"}` : ""}${impact.removedConceptReferenceCount > 0 ? ` plus ${impact.removedConceptReferenceCount} concept binding${impact.removedConceptReferenceCount === 1 ? "" : "s"}` : ""} from this project only. No external source file was changed or deleted.`);
  return { project: next, impact };
}
