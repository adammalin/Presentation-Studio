import type { StudioDesignArchetype, StudioLayoutRecipe } from "../types";
import type { LayoutCompatibilityResult, LayoutContentProfile, TemplateLayoutIntent } from "./layout-semantics";
import { rankLayoutCompatibility } from "./layout-semantics";
import type { TemplateLayoutPreview } from "./template-catalog";

export const STUDIO_DESIGN_ARCHETYPES: readonly StudioDesignArchetype[] = ["cover", "section", "assertion-evidence", "text-led", "hero-figure", "comparison", "image-series", "portrait-series", "table", "data-visualization", "process-flow", "technical-diagram", "conclusion", "source-preserve"];

export interface StudioArchetypeHints {
  slideNumber?: number;
  title?: string;
  connectorCount?: number;
  nativeObjectCount?: number;
  repeatedImageSeries?: boolean;
  recommendedRecipe?: StudioLayoutRecipe;
  protectedSourceComposition?: boolean;
  requestedArchetype?: StudioDesignArchetype;
}

export interface StudioArchetypeInference {
  archetype: StudioDesignArchetype;
  confidence: "high" | "medium" | "low";
  reasons: string[];
  preservationPolicy: "editable-composition" | "relationship-preserving" | "source-preserve";
}

export interface StudioCompositionPlan {
  archetype: StudioDesignArchetype;
  confidence: StudioArchetypeInference["confidence"];
  strategy: "source-preserve" | "converted-template-layout" | "shared-archetype-on-native-base";
  recipe: StudioLayoutRecipe;
  layoutId?: string;
  layoutName?: string;
  layoutScore?: number;
  reasons: string[];
  preservationPolicy: StudioArchetypeInference["preservationPolicy"];
  alternativeLayouts: Array<LayoutCompatibilityResult & { family?: string; selectionPolicy?: string }>;
  requiredChecks: string[];
}

const CONCLUSION_PATTERN = /\b(conclusion|summary|questions?|thank you|next steps?)\b/i;

export function inferStudioDesignArchetype(profile: LayoutContentProfile, hints: StudioArchetypeHints = {}): StudioArchetypeInference {
  if (hints.requestedArchetype) return { archetype: hints.requestedArchetype, confidence: "high", reasons: ["The requested communication archetype is explicit."], preservationPolicy: hints.requestedArchetype === "source-preserve" ? "source-preserve" : ["technical-diagram", "process-flow", "image-series"].includes(hints.requestedArchetype) ? "relationship-preserving" : "editable-composition" };
  if (hints.protectedSourceComposition) return { archetype: "source-preserve", confidence: "high", reasons: ["The slide is a protected approved source composition."], preservationPolicy: "source-preserve" };
  if (hints.slideNumber === 1 || profile.desiredIntent === "cover") return { archetype: "cover", confidence: "high", reasons: ["The opening slide requires an approved cover composition."], preservationPolicy: "editable-composition" };
  if ((profile.desiredIntent === "conclusion" || CONCLUSION_PATTERN.test(hints.title ?? "")) && profile.bodyCharacterCount <= 600) return { archetype: "conclusion", confidence: "high", reasons: ["The title and bounded closing message indicate a closing composition."], preservationPolicy: "editable-composition" };
  if (profile.tableCount > 0) return { archetype: "table", confidence: "high", reasons: ["An editable table is the primary structured evidence."], preservationPolicy: "relationship-preserving" };
  if (profile.chartCount > 0) return { archetype: "data-visualization", confidence: "high", reasons: ["A chart or plotted data object is primary evidence."], preservationPolicy: "relationship-preserving" };
  if (hints.recommendedRecipe === "ornl-title-metric-grid") return { archetype: "comparison", confidence: "high", reasons: ["Repeated source metric rows form a comparison system; native grouping and connectors are row furniture rather than technical-diagram evidence."], preservationPolicy: "relationship-preserving" };
  if (hints.recommendedRecipe === "ornl-title-card-grid" && profile.imageCount === 0 && (hints.connectorCount ?? 0) === 0 && profile.bodyBlockCount >= 7) return { archetype: "comparison", confidence: "high", reasons: ["Repeated exact editorial records form one dense peer grid; a legacy native group is record furniture rather than technical-diagram evidence."], preservationPolicy: "editable-composition" };
  if (hints.recommendedRecipe === "ornl-title-two-column" && profile.imageCount >= 12 && (hints.connectorCount ?? 0) === 0) return { archetype: "technical-diagram", confidence: "high", reasons: ["A dense peer-logo field requires one relationship-preserving contained-image system rather than a generic figure stack or aggregate raster."], preservationPolicy: "relationship-preserving" };
  if (hints.repeatedImageSeries || profile.imageCount >= 3 && profile.imageCount <= 6 && profile.bodyBlockCount + profile.captionBlockCount >= profile.imageCount) return { archetype: "image-series", confidence: hints.repeatedImageSeries ? "high" : "medium", reasons: [hints.repeatedImageSeries ? "Immutable source geometry exposes repeated image-heading-evidence groups." : "The slide contains a peer image set with enough supporting text to form repeated evidence groups."], preservationPolicy: "relationship-preserving" };
  if (profile.imageCount >= 8 && profile.captionBlockCount >= Math.floor(profile.imageCount * .7)) return { archetype: "portrait-series", confidence: "medium", reasons: ["A large image-label inventory is best treated as a portrait or people series."], preservationPolicy: "relationship-preserving" };
  if ((hints.connectorCount ?? 0) >= 2) return { archetype: "technical-diagram", confidence: "high", reasons: ["Multiple connectors encode a relationship-bearing technical diagram."], preservationPolicy: "relationship-preserving" };
  if ((hints.nativeObjectCount ?? 0) > 0) return { archetype: "technical-diagram", confidence: "medium", reasons: ["A grouped or unsupported native object should remain one technical evidence unit unless verified for reconstruction."], preservationPolicy: "relationship-preserving" };
  if (profile.imageCount === 1) return { archetype: "hero-figure", confidence: "high", reasons: ["One visual is available to serve as the primary evidence field."], preservationPolicy: "relationship-preserving" };
  if (profile.imageCount > 1) return { archetype: "comparison", confidence: "medium", reasons: ["Multiple peer visuals require a balanced comparison system."], preservationPolicy: "relationship-preserving" };
  if (profile.bodyBlockCount >= 2 && profile.bodyBlockCount <= 4) return { archetype: "comparison", confidence: "medium", reasons: ["Two to four peer text groups form a comparison or objective system."], preservationPolicy: "editable-composition" };
  if (profile.bodyCharacterCount > 700 || profile.bodyBlockCount === 1 && (profile.bodyBlockCharacterCounts?.[0] ?? 0) > 420) return { archetype: "text-led", confidence: "high", reasons: ["One dense exact-content reading flow requires the roomiest approved content treatment."], preservationPolicy: "editable-composition" };
  return { archetype: "assertion-evidence", confidence: "medium", reasons: ["The slide has one primary assertion and a bounded supporting-content inventory."], preservationPolicy: "editable-composition" };
}

export function defaultRecipeForArchetype(archetype: StudioDesignArchetype): StudioLayoutRecipe {
  if (archetype === "source-preserve") return "source";
  if (["cover", "section", "portrait-series", "conclusion"].includes(archetype)) return "template-layout";
  if (["assertion-evidence", "text-led"].includes(archetype)) return "ornl-title-content";
  if (["hero-figure", "data-visualization", "technical-diagram"].includes(archetype)) return "ornl-title-two-column";
  if (archetype === "comparison") return "ornl-title-card-grid";
  if (archetype === "image-series") return "ornl-title-image-series";
  if (archetype === "table") return "ornl-title-table";
  if (archetype === "process-flow") return "ornl-title-process-flow";
  return "ornl-title-content";
}

function compatibleDetailedRecipe(archetype: StudioDesignArchetype, recipe: StudioLayoutRecipe | undefined): StudioLayoutRecipe | undefined {
  if (!recipe || recipe === "source" || recipe === "template-layout") return undefined;
  const allowed: Partial<Record<StudioDesignArchetype, StudioLayoutRecipe[]>> = {
    "assertion-evidence": ["ornl-title-content", "ornl-title-objective-columns", "ornl-title-card-grid", "ornl-title-steps-evidence"],
    "text-led": ["ornl-title-content", "ornl-title-objective-columns", "ornl-title-card-grid"],
    "hero-figure": ["ornl-title-two-column", "ornl-title-steps-evidence", "ornl-title-figure-grid", "ornl-title-labeled-figure-grid"],
    comparison: ["ornl-title-card-grid", "ornl-title-objective-columns", "ornl-title-figure-grid", "ornl-title-labeled-figure-grid", "ornl-title-image-series", "ornl-title-metric-grid"],
    "image-series": ["ornl-title-image-series", "ornl-title-figure-grid", "ornl-title-labeled-figure-grid"],
    table: ["ornl-title-table"],
    "data-visualization": ["ornl-title-two-column", "ornl-title-figure-grid", "ornl-title-labeled-figure-grid", "ornl-title-table"],
    "process-flow": ["ornl-title-process-flow", "ornl-title-steps-evidence", "ornl-title-question-diagram"],
    "technical-diagram": ["ornl-title-two-column", "ornl-title-figure-grid", "ornl-title-labeled-figure-grid", "ornl-title-question-diagram", "ornl-title-challenges-evidence", "ornl-title-process-flow"],
  };
  return allowed[archetype]?.includes(recipe) ? recipe : undefined;
}

function desiredIntentFor(archetype: StudioDesignArchetype): TemplateLayoutIntent | undefined {
  if (archetype === "cover") return "cover";
  if (archetype === "section") return "section";
  if (archetype === "conclusion") return "conclusion";
  if (["hero-figure", "image-series", "portrait-series", "technical-diagram"].includes(archetype)) return "visual";
  if (["table", "data-visualization"].includes(archetype)) return "data";
  if (archetype === "comparison") return "comparison";
  if (["assertion-evidence", "text-led", "process-flow"].includes(archetype)) return "content";
  return undefined;
}

function exactTemplateFit(layout: TemplateLayoutPreview, profile: LayoutContentProfile, archetype: StudioDesignArchetype): boolean {
  const semantics = layout.semantic;
  if (!semantics || !semantics.contract.compatibleArchetypes.includes(archetype)) return false;
  if (["cover", "section", "conclusion"].includes(archetype)) return semantics.capabilities.title;
  if (archetype === "portrait-series") return semantics.contract.family === "portrait-series" && semantics.capabilities.imageSlots === profile.imageCount;
  if (archetype === "image-series") return semantics.contract.family === "image-series" && semantics.capabilities.imageSlots === profile.imageCount && semantics.capabilities.bodySlots + semantics.capabilities.captionSlots >= Math.max(profile.bodyBlockCount, profile.captionBlockCount);
  if (archetype === "comparison") return semantics.contract.family === "comparison"
    && semantics.capabilities.columns === profile.bodyBlockCount
    && profile.titleCharacterCount <= 70
    && profile.bodyCharacterCount <= 520;
  if (archetype === "hero-figure") return semantics.contract.family === "hero-visual" && semantics.capabilities.imageSlots === 1 && profile.titleCharacterCount <= 54 && profile.bodyCharacterCount <= 220;
  if (archetype === "assertion-evidence") return semantics.contract.family === "reading" && profile.bodyBlockCount <= 1 && profile.titleCharacterCount <= 82 && profile.bodyCharacterCount <= 520;
  return false;
}

export function planStudioComposition(profile: LayoutContentProfile, layouts: TemplateLayoutPreview[], hints: StudioArchetypeHints = {}): StudioCompositionPlan {
  const inference = inferStudioDesignArchetype(profile, hints);
  const enrichedProfile: LayoutContentProfile = { ...profile, desiredIntent: desiredIntentFor(inference.archetype) ?? profile.desiredIntent, designArchetype: inference.archetype };
  const ranked = rankLayoutCompatibility(layouts, enrichedProfile);
  const alternatives = ranked.slice(0, 6).map((result) => {
    const layout = layouts.find((candidate) => candidate.id === result.layoutId);
    return { ...result, family: layout?.semantic?.contract.family, selectionPolicy: layout?.semantic?.contract.selectionPolicy };
  });
  if (inference.archetype === "source-preserve") return { archetype: inference.archetype, confidence: inference.confidence, strategy: "source-preserve", recipe: "source", reasons: inference.reasons, preservationPolicy: inference.preservationPolicy, alternativeLayouts: alternatives, requiredChecks: ["Source PowerPoint pixels remain unchanged.", "No source content, artwork, or relationships are replaced."] };

  const exact = ranked.map((result) => ({ result, layout: layouts.find((candidate) => candidate.id === result.layoutId) })).find(({ result, layout }) => Boolean(layout) && result.score >= 60 && result.unmetNeeds.length === 0 && exactTemplateFit(layout!, enrichedProfile, inference.archetype));
  const forceTemplate = ["cover", "section", "portrait-series", "conclusion"].includes(inference.archetype);
  const canUseExactTemplate = Boolean(exact) && !["text-led", "table", "data-visualization", "process-flow", "technical-diagram"].includes(inference.archetype);
  const strategy = forceTemplate || canUseExactTemplate ? "converted-template-layout" as const : "shared-archetype-on-native-base" as const;
  const detailedRecipe = compatibleDetailedRecipe(inference.archetype, hints.recommendedRecipe);
  const recipe = strategy === "converted-template-layout"
    ? "template-layout"
    : inference.archetype === "technical-diagram" && profile.imageCount >= 2 && (hints.connectorCount ?? 0) > 0
      ? "ornl-title-figure-grid"
    : inference.archetype === "image-series" && !hints.repeatedImageSeries
      ? (hints.connectorCount ?? 0) > 0 ? "ornl-title-figure-grid" : "ornl-title-labeled-figure-grid"
      : detailedRecipe ?? defaultRecipeForArchetype(inference.archetype);
  const denseEditorialRecordGrid = recipe === "ornl-title-card-grid" && profile.bodyBlockCount >= 7 && profile.bodyCharacterCount >= 900;
  const denseSourceEditorialField = recipe === "ornl-title-two-column" && profile.imageCount === 1 && profile.bodyBlockCount >= 8 && profile.bodyCharacterCount >= 700;
  const standardCover = inference.archetype === "cover"
    ? layouts.find((candidate) => candidate.category === "title" && candidate.semantic?.contract.family === "cover" && candidate.semantic.contract.selectionPolicy === "sacred" && !candidate.placeholderTypes.includes("pic") && !/text[- ]only/i.test(candidate.name))
    : undefined;
  const layout = standardCover ?? exact?.layout ?? (forceTemplate ? layouts.find((candidate) => ranked[0]?.layoutId === candidate.id) : undefined);
  const layoutResult = layout ? ranked.find((result) => result.layoutId === layout.id) : undefined;
  const reasons = [...inference.reasons];
  if (strategy === "converted-template-layout" && layout) reasons.push(`${layout.name} supplies an exact compatible native ORNL slot system.`);
  else reasons.push(
    inference.archetype === "technical-diagram" && profile.imageCount >= 2 && (hints.connectorCount ?? 0) > 0
      ? "Multiple source visuals and connectors form one relationship-bearing technical field. Preserve and scale that field as a unit in the shared figure-grid composition."
      : inference.archetype === "image-series" && !hints.repeatedImageSeries
        ? (hints.connectorCount ?? 0) > 0
          ? "The slide resembles an image series, but immutable source geometry does not prove complete image-heading-evidence groups and at least one connector carries visual relationships. Preserve and scale the complete figure field in the relationship-safe figure grid instead of guessing column assignments."
          : "The slide resembles an image series, but immutable source geometry does not prove complete image-heading-evidence groups. Use the relationship-safe labeled figure grid instead of guessing column assignments."
        : detailedRecipe
          ? `Source structure supports the compatible ${detailedRecipe} recipe inside the ${inference.archetype} communication archetype.`
          : "No exact approved layout contract safely holds the complete content, so Studio should use a responsive shared archetype on the neutral native ORNL base."
  );
  return {
    archetype: inference.archetype,
    confidence: inference.confidence,
    strategy,
    recipe,
    layoutId: strategy === "converted-template-layout" ? layout?.id : undefined,
    layoutName: strategy === "converted-template-layout" ? layout?.name : undefined,
    layoutScore: layoutResult?.score,
    reasons,
    preservationPolicy: inference.preservationPolicy,
    alternativeLayouts: alternatives,
    requiredChecks: [
      "Preserve exact source wording, numbers, units, qualifications, attribution, and semantic colors.",
      "Attach the real native ORNL master and layout; do not recreate protected marks or footer artwork.",
      denseEditorialRecordGrid
        ? "This exact-content editorial record grid may use the bounded 10.5 pt record-grid exception only when PowerPoint-native measurement proves every record is readable and collision-free; do not misclassify a peer record as a footer or caption."
        : denseSourceEditorialField
          ? "This exact-content source-geometry editorial field may use the bounded 8.5 pt compact-grid exception only when it stays at least source-equivalent and PowerPoint-native measurement proves every record readable and collision-free; otherwise hold or use a verified continuation."
        : "Keep body type at or above 16 pt and captions or labels at or above 14 pt.",
      "Preserve every image-caption, heading-evidence, table-cell, connector, and technical-figure relationship.",
      "Build, render, and measure the exact editable candidate in Microsoft PowerPoint before visual acceptance.",
    ],
  };
}
