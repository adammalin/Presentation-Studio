import { STUDIO_WEB_SCENE_VERSION, type ProjectResource, type StudioConceptInfluence, type StudioConceptReference, type StudioConceptUntrustedElement, type StudioWebScene } from "../types";
import { linkStudioVisualNeed, reopenStudioVisualNeedForDetachedConcept } from "./studio-visual-needs";

const REQUIRED_UNTRUSTED: StudioConceptUntrustedElement[] = ["generated-text", "generated-logos", "generated-data", "generated-technical-details", "generated-claims"];

export interface StudioConceptReferenceRequest {
  id?: string;
  origin: StudioConceptReference["origin"];
  approvedInfluences: StudioConceptInfluence[];
  blueprint: StudioConceptReference["blueprint"];
  provenance?: StudioConceptReference["provenance"];
  visualNeedId?: string;
}

function boundedText(value: string, maximum: number, field: string) {
  const result = value.trim();
  if (!result) throw new Error(`${field} is required.`);
  if (result.length > maximum) throw new Error(`${field} must be ${maximum} characters or fewer.`);
  return result;
}

function validateBlueprint(blueprint: StudioConceptReference["blueprint"]): StudioConceptReference["blueprint"] {
  if (blueprint.zones.length > 20 || blueprint.styleNotes.length > 20 || blueprint.reconstructionNotes.length > 20) throw new Error("A concept blueprint may contain at most 20 zones, style notes, and reconstruction notes.");
  const zones = blueprint.zones.map((zone, index) => {
    const values = [zone.x, zone.y, zone.width, zone.height];
    if (values.some((value) => !Number.isFinite(value)) || zone.x < 0 || zone.y < 0 || zone.width <= 0 || zone.height <= 0 || zone.x + zone.width > 1 || zone.y + zone.height > 1) throw new Error(`Concept zone ${index + 1} must use normalized 0-1 geometry inside the slide.`);
    return { ...zone, id: boundedText(zone.id, 120, `Concept zone ${index + 1} ID`) };
  });
  return {
    summary: boundedText(blueprint.summary, 1_000, "Concept summary"),
    zones,
    styleNotes: blueprint.styleNotes.map((note, index) => boundedText(note, 500, `Style note ${index + 1}`)),
    reconstructionNotes: blueprint.reconstructionNotes.map((note, index) => boundedText(note, 500, `Reconstruction note ${index + 1}`)),
  };
}

export function attachStudioConceptReference(scene: StudioWebScene, slideNumber: number, resource: ProjectResource, request: StudioConceptReferenceRequest): StudioWebScene {
  if (resource.kind !== "image" || !resource.mediaType.startsWith("image/")) throw new Error("A Studio concept reference must be an embedded image Resource.");
  if (!resource.bytes?.byteLength) throw new Error("The concept image is not embedded in this project.");
  if (!request.approvedInfluences.length) throw new Error("Approve at least one visual characteristic to follow from the concept.");
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the Studio scene.`);
  const now = new Date().toISOString();
  const reference: StudioConceptReference = {
    id: request.id?.trim().slice(0, 180) || `concept-${slideNumber}-${resource.sha256.slice(0, 12)}`,
    resourceId: resource.id,
    resourceSha256: resource.sha256,
    sourceTextHash: slide.sourceTextHash,
    status: "concept-only",
    origin: request.origin,
    approvedInfluences: [...new Set(request.approvedInfluences)],
    untrustedElements: REQUIRED_UNTRUSTED,
    blueprint: validateBlueprint(request.blueprint),
    provenance: request.provenance ? {
      model: request.provenance.model?.trim().slice(0, 180),
      promptSummary: request.provenance.promptSummary?.trim().slice(0, 1_000),
      generatedAt: request.provenance.generatedAt,
    } : undefined,
    visualNeedId: request.visualNeedId,
    attachedAt: now,
  };
  const existing = slide.conceptReferences ?? [];
  const conceptReferences = [...existing.filter((item) => item.id !== reference.id && item.resourceId !== resource.id), reference];
  const attached = {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber === slideNumber ? { ...item, conceptReferences, qualityReview: undefined, updatedAt: now } : item),
  };
  return request.visualNeedId ? linkStudioVisualNeed(attached, slideNumber, request.visualNeedId, reference.id) : attached;
}

export function removeStudioConceptReference(scene: StudioWebScene, slideNumber: number, referenceId: string): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the Studio scene.`);
  const existing = slide.conceptReferences ?? [];
  if (!existing.some((item) => item.id === referenceId)) throw new Error("The concept reference is not attached to this slide.");
  const now = new Date().toISOString();
  const detached = {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber === slideNumber ? { ...item, conceptReferences: existing.filter((reference) => reference.id !== referenceId), qualityReview: undefined, updatedAt: now } : item),
  };
  return reopenStudioVisualNeedForDetachedConcept(detached, slideNumber, referenceId);
}
