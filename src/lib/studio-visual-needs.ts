import {
  STUDIO_WEB_SCENE_VERSION,
  type StudioConceptInfluence,
  type StudioVisualNeed,
  type StudioVisualNeedType,
  type StudioWebScene,
  type StudioWebSlide,
} from "../types";
import { analyzeStudioDesignImpact } from "./studio-design-impact";

export interface StudioVisualNeedRequest {
  id?: string;
  type: StudioVisualNeedType;
  reason: string;
  communicationJob: string;
  expression?: StudioVisualNeed["expression"];
  approvedInfluences?: StudioConceptInfluence[];
  disclosurePolicy?: StudioVisualNeed["disclosurePolicy"];
  approvedContentSummary?: string;
  targetSlot?: Partial<StudioVisualNeed["targetSlot"]>;
}

const DEFAULT_INFLUENCES: Record<StudioVisualNeedType, StudioConceptInfluence[]> = {
  "layout-concept": ["composition", "visual-hierarchy", "negative-space", "visual-rhythm"],
  "figure-concept": ["figure-concept", "visual-hierarchy", "negative-space"],
  "image-treatment": ["image-treatment", "composition", "color-balance"],
  "supporting-visual": ["composition", "negative-space", "image-treatment"],
  "diagram-rebuild": ["figure-concept", "visual-hierarchy", "visual-rhythm"],
};

const DEFAULT_TARGET_ROLE: Record<StudioVisualNeedType, StudioVisualNeed["targetSlot"]["role"]> = {
  "layout-concept": "whole-slide",
  "figure-concept": "figure",
  "image-treatment": "primary-visual",
  "supporting-visual": "supporting-evidence",
  "diagram-rebuild": "figure",
};

function bounded(value: string, maximum: number, label: string) {
  const normalized = value.trim();
  if (!normalized) throw new Error(`${label} is required.`);
  if (normalized.length > maximum) throw new Error(`${label} must be ${maximum} characters or fewer.`);
  return normalized;
}

function structureInventory(slide: StudioWebSlide): StudioVisualNeed["structureInventory"] {
  const visible = slide.nodes.filter((node) => node.visible);
  return {
    titleCount: visible.filter((node) => node.role === "title").length,
    textGroupCount: visible.filter((node) => node.kind === "text" && node.role !== "title").length,
    imageCount: visible.filter((node) => node.kind === "image").length,
    tableCount: visible.filter((node) => node.kind === "table").length,
    figureCount: slide.figureTreatments.length || visible.filter((node) => ["chart", "group"].includes(node.role) || node.kind === "native-object").length,
    calloutCount: visible.filter((node) => node.kind === "connector" || /callout/i.test(node.name)).length,
  };
}

function inventorySentence(inventory: StudioVisualNeed["structureInventory"]) {
  const parts = [
    `${inventory.titleCount} title zone${inventory.titleCount === 1 ? "" : "s"}`,
    `${inventory.textGroupCount} editable text group${inventory.textGroupCount === 1 ? "" : "s"}`,
    `${inventory.imageCount} source image${inventory.imageCount === 1 ? "" : "s"}`,
    `${inventory.tableCount} table${inventory.tableCount === 1 ? "" : "s"}`,
    `${inventory.figureCount} technical figure group${inventory.figureCount === 1 ? "" : "s"}`,
    `${inventory.calloutCount} callout or connector${inventory.calloutCount === 1 ? "" : "s"}`,
  ];
  return parts.join(", ");
}

function brandExpression(scene: StudioWebScene, type: StudioVisualNeedType, expression: StudioVisualNeed["expression"]): StudioVisualNeed["brandExpression"] {
  const used = new Set(scene.slides.flatMap((slide) => (slide.visualNeeds ?? []).map((need) => need.brandExpression?.motif).filter(Boolean)));
  const candidates: Record<StudioVisualNeedType, StudioVisualNeed["brandExpression"]["motif"][]> = {
    "layout-concept": ["modular-square-grid", "green-motion-gradient", "pattern-free"],
    "figure-concept": ["directional-rule", "modular-square-grid", "subordinate-hex-system"],
    "image-treatment": ["editorial-layering", "green-motion-gradient", "pattern-free"],
    "supporting-visual": ["green-motion-gradient", "modular-square-grid", "pattern-free"],
    "diagram-rebuild": ["directional-rule", "subordinate-hex-system", "modular-square-grid"],
  };
  const motif = expression === "restrained" ? "pattern-free" : candidates[type].find((candidate) => !used.has(candidate)) ?? candidates[type][0];
  const accent: StudioVisualNeed["brandExpression"]["accent"] = expression === "restrained" ? "none" : type === "figure-concept" || type === "diagram-rebuild" ? "Infinity" : type === "image-treatment" ? "Forge" : type === "supporting-visual" ? "Biome" : "Aqua";
  return {
    motif,
    accent,
    accentRole: accent === "none" ? "No accent; use green, navy, and neutrals only." : `Use ${accent} for one focused evidence or directional role, below 25% of the concept.`,
    typographyStrategy: "no-generated-type-reserve-editable-aptos-zones",
    rationale: `${motif.replaceAll("-", " ")} supports this ${type.replaceAll("-", " ")} without repeating every ORNL device or flooding the slide with green and navy panels.`,
  };
}

function promptPackage(need: Pick<StudioVisualNeed, "type" | "expression" | "approvedInfluences" | "disclosurePolicy" | "approvedContentSummary" | "brandExpression" | "structureInventory" | "targetSlot">): StudioVisualNeed["promptPackage"] {
  const expression = need.expression === "restrained"
    ? "restrained editorial clarity with generous Polar whitespace and one quiet ORNL Green anchor"
    : need.expression === "expressive"
      ? "expressive but disciplined visual storytelling with ORNL Green as an anchor, not a full-slide color flood"
      : "balanced technical storytelling with clear hierarchy, Polar whitespace, restrained ORNL Green, Hale Navy, and Graphite accents";
  const approvedSummary = need.disclosurePolicy === "exact-content-approved" && need.approvedContentSummary
    ? ` The person explicitly approved this bounded content summary for concept generation: ${need.approvedContentSummary}`
    : " Use abstract placeholder zones only; no source wording or technical content is disclosed for this concept.";
  return {
    prompt: `MODE: Image Generation. PRIMARY ARTIFACT COUNT: 1. ONLY ALLOWED TEXT: NONE. LOGO: OMIT. PLACEHOLDERS: NONE. Create one concept-only art-direction raster for a 16:9 ORNL technical presentation ${need.type.replaceAll("-", " ")}. Use ${expression}. BRAND EXPRESSION: ${need.brandExpression.motif.replaceAll("-", " ")}; ${need.brandExpression.accentRole} ${need.brandExpression.rationale} Reserve editable-content zones for ${inventorySentence(need.structureInventory)}. Target the ${need.targetSlot.role.replaceAll("-", " ")} area with a ${need.targetSlot.aspectRatio} treatment; ${need.targetSlot.placementNotes}. ALLOWED VISUALS: abstract structural fields, flat geometric zones, directional rhythm, and non-readable evidence placeholders needed to express ${need.approvedInfluences.join(", ")}.${approvedSummary} MUST PRESERVE: clear safe margins, one clean title zone, production-feasible flat geometry, square 90-degree corners, and empty zones that can be rebuilt with editable Aptos text and source-bound evidence. FORBIDDEN: all readable text, logos, seals, slogans, URLs, CTAs, numbers, data, claims, labels, methods, event details, footers, reviewer notes, working notes, fake screenshots, fake software UI, and invented scientific content. Do not create a finished slide; create visual direction only.`,
    negativePrompt: "No readable text, letters, numbers, labels, logos, seals, watermarks, data, charts with invented values, claims, scientific details, fake software UI, fake screenshots, decorative rounded cards, arbitrary borders around text boxes, generic green-and-navy panel flooding, uncontrolled multi-color gradients, shadows, glass effects, or rasterized final-slide typography.",
    contentSafety: need.disclosurePolicy === "abstract-structure-only"
      ? "Structure only. Do not send source text, source pixels, names, technical claims, data, or screenshots to an image generator."
      : "Only the explicitly approved bounded summary may be sent. The original slide, exact wording, data, logos, technical details, and source pixels remain excluded unless separately authorized.",
  };
}

function updateSlide(scene: StudioWebScene, slideNumber: number, updater: (slide: StudioWebSlide, now: string) => StudioWebSlide): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the Studio scene.`);
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber === slideNumber ? updater(item, now) : item),
  };
}

function updateQueueMetadata(scene: StudioWebScene, slideNumber: number, updater: (slide: StudioWebSlide, now: string) => StudioWebSlide): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the Studio scene.`);
  const now = new Date().toISOString();
  return { ...scene, slides: scene.slides.map((item) => item.slideNumber === slideNumber ? updater(item, now) : item) };
}

export function createStudioVisualNeed(scene: StudioWebScene, slideNumber: number, request: StudioVisualNeedRequest): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the Studio scene.`);
  const disclosurePolicy = request.disclosurePolicy ?? "abstract-structure-only";
  if (request.approvedContentSummary && disclosurePolicy !== "exact-content-approved") throw new Error("An approved content summary requires exact-content-approved disclosure.");
  const approvedContentSummary = request.approvedContentSummary ? bounded(request.approvedContentSummary, 800, "Approved content summary") : undefined;
  if (disclosurePolicy === "exact-content-approved" && !approvedContentSummary) throw new Error("Exact-content-approved disclosure requires a bounded approved content summary.");
  const inventory = structureInventory(slide);
  const approvedInfluences = [...new Set(request.approvedInfluences?.length ? request.approvedInfluences : DEFAULT_INFLUENCES[request.type])];
  const expression = request.expression ?? "balanced";
  const expressionRecipe = brandExpression(scene, request.type, expression);
  const targetSlot: StudioVisualNeed["targetSlot"] = {
    role: request.targetSlot?.role ?? DEFAULT_TARGET_ROLE[request.type],
    aspectRatio: request.targetSlot?.aspectRatio ?? (request.type === "layout-concept" ? "16:9" : "free"),
    placementNotes: bounded(request.targetSlot?.placementNotes ?? "preserve a clean title zone, safe margins, and room for exact editable source content", 500, "Target placement notes"),
  };
  const now = new Date().toISOString();
  const need: StudioVisualNeed = {
    id: request.id?.trim().slice(0, 180) || `visual-need-${slideNumber}-${now.replace(/\D/g, "").slice(0, 17)}`,
    type: request.type,
    status: "brief-ready",
    sourceTextHash: slide.sourceTextHash,
    reason: bounded(request.reason, 1_000, "Visual-need reason"),
    communicationJob: bounded(request.communicationJob, 1_000, "Communication job"),
    expression,
    approvedInfluences,
    disclosurePolicy,
    approvedContentSummary,
    brandExpression: expressionRecipe,
    structureInventory: inventory,
    targetSlot,
    promptPackage: promptPackage({ type: request.type, expression, approvedInfluences, disclosurePolicy, approvedContentSummary, brandExpression: expressionRecipe, structureInventory: inventory, targetSlot }),
    createdAt: now,
    updatedAt: now,
  };
  return updateQueueMetadata(scene, slideNumber, (current, updatedAt) => ({
    ...current,
    visualNeeds: [...(current.visualNeeds ?? []).filter((item) => item.id !== need.id), { ...need, updatedAt }],
  }));
}

export function linkStudioVisualNeed(scene: StudioWebScene, slideNumber: number, visualNeedId: string, referenceId: string): StudioWebScene {
  return updateSlide(scene, slideNumber, (slide, now) => {
    const needs = slide.visualNeeds ?? [];
    const need = needs.find((item) => item.id === visualNeedId);
    if (!need) throw new Error("The requested visual need is not attached to this slide.");
    if (need.sourceTextHash !== slide.sourceTextHash) throw new Error("The visual need is stale because the source content binding changed.");
    if (need.status === "held" || need.status === "resolved") throw new Error("Reopen this visual need before attaching another concept.");
    return {
      ...slide,
      visualNeeds: needs.map((item) => item.id === visualNeedId ? { ...item, status: "concept-attached", linkedConceptReferenceId: referenceId, resolutionNote: undefined, updatedAt: now } : item),
      qualityReview: undefined,
      updatedAt: now,
    };
  });
}

export function reopenStudioVisualNeedForDetachedConcept(scene: StudioWebScene, slideNumber: number, referenceId: string): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide?.visualNeeds?.some((item) => item.linkedConceptReferenceId === referenceId)) return scene;
  return updateSlide(scene, slideNumber, (current, now) => ({
    ...current,
    visualNeeds: (current.visualNeeds ?? []).map((item) => item.linkedConceptReferenceId === referenceId ? { ...item, status: "brief-ready", linkedConceptReferenceId: undefined, resolutionNote: "The linked concept was detached; the brief is ready for a new concept.", updatedAt: now } : item),
    qualityReview: undefined,
    updatedAt: now,
  }));
}

export function holdStudioVisualNeed(scene: StudioWebScene, slideNumber: number, visualNeedId: string, note: string): StudioWebScene {
  return updateQueueMetadata(scene, slideNumber, (slide, now) => {
    const needs = slide.visualNeeds ?? [];
    if (!needs.some((item) => item.id === visualNeedId)) throw new Error("The requested visual need is not attached to this slide.");
    return { ...slide, visualNeeds: needs.map((item) => item.id === visualNeedId ? { ...item, status: "held", resolutionNote: bounded(note, 1_000, "Hold note"), updatedAt: now } : item), updatedAt: now };
  });
}

export function markStudioVisualNeedsReconstructionReady(scene: StudioWebScene, slideNumber: number, visualNeedIds?: string[]): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const selected = visualNeedIds?.length ? new Set(visualNeedIds) : undefined;
  if (selected && [...selected].some((id) => !slide?.visualNeeds?.some((item) => item.id === id && item.status === "concept-attached"))) throw new Error("Every reconstructed visual need must be concept-attached on this slide.");
  if (!slide?.visualNeeds?.some((item) => item.status === "concept-attached" && (!selected || selected.has(item.id)))) return scene;
  const impact = analyzeStudioDesignImpact(slide);
  const insufficient = impact.requirements.filter((item) => (!selected || selected.has(item.visualNeedId)) && !item.passed);
  if (insufficient.length) throw new Error(`Editable concept reconstruction is not yet material: ${insufficient.map((item) => item.reason).join(" ")}`);
  return updateSlide(scene, slideNumber, (current, now) => ({
    ...current,
    visualNeeds: (current.visualNeeds ?? []).map((item) => item.status === "concept-attached" && (!selected || selected.has(item.id)) ? { ...item, status: "reconstruction-ready", resolutionNote: "An editable Studio reconstruction was staged from this concept direction; native PowerPoint review is still required.", updatedAt: now } : item),
    updatedAt: now,
  }));
}

export function resolveStudioVisualNeeds(scene: StudioWebScene, slideNumber: number, note: string): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide?.visualNeeds?.some((item) => item.status === "reconstruction-ready")) return scene;
  const impact = analyzeStudioDesignImpact(slide);
  const insufficient = impact.requirements.filter((item) => slide.visualNeeds?.some((need) => need.id === item.visualNeedId && need.status === "reconstruction-ready") && !item.passed);
  if (insufficient.length) throw new Error(`Visual needs cannot resolve before material reconstruction: ${insufficient.map((item) => item.reason).join(" ")}`);
  return updateQueueMetadata(scene, slideNumber, (current, now) => ({
    ...current,
    visualNeeds: (current.visualNeeds ?? []).map((item) => item.status === "reconstruction-ready" ? { ...item, status: "resolved", resolutionNote: bounded(note, 1_000, "Resolution note"), updatedAt: now } : item),
  }));
}
