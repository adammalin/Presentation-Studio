import type { PresentationStudioProject, ProjectResource } from "../types";
import { touchProject } from "./project";

export interface ResourceRemovalImpact {
  resource: ProjectResource;
  linkedDeckIds: string[];
  linkedDeckNames: string[];
  removedExemplarCount: number;
  removedThreadCount: number;
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
  };
}

export function removeResourceFromProject(project: PresentationStudioProject, resourceId: string): { project: PresentationStudioProject; impact: ResourceRemovalImpact } {
  const impact = resourceRemovalImpact(project, resourceId);
  const linkedDeckIds = new Set(impact.linkedDeckIds);
  const next = touchProject({
    ...project,
    resources: project.resources.filter((resource) => resource.id !== resourceId),
    decks: project.decks.filter((deck) => !linkedDeckIds.has(deck.id)),
    styleExemplars: project.styleExemplars.filter((item) => item.resourceId !== resourceId && !linkedDeckIds.has(item.deckId)),
    designThreads: project.designThreads.filter((thread) => !linkedDeckIds.has(thread.deckId)),
  }, "resource-removed", `Removed the embedded Resource ${impact.resource.name}${impact.linkedDeckIds.length > 0 ? ` and ${impact.linkedDeckIds.length} linked deck${impact.linkedDeckIds.length === 1 ? "" : "s"}` : ""} from this project only. No external source file was changed or deleted.`);
  return { project: next, impact };
}
