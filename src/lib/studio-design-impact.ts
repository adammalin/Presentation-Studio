import type { StudioVisualNeed, StudioWebFrame, StudioWebSlide } from "../types";

const EMU_PER_POINT = 12_700;
const MEANINGFUL_DELTA = 6 * EMU_PER_POINT;

export type StudioDesignImpactLevel = "unchanged" | "typography-only" | "cleanup" | "layout-redesign" | "figure-redesign" | "full-redesign";

export interface StudioDesignImpactRequirement {
  visualNeedId: string;
  type: StudioVisualNeed["type"];
  passed: boolean;
  reason: string;
}

export interface StudioDesignImpact {
  level: StudioDesignImpactLevel;
  meaningful: boolean;
  geometryChangedNodeIds: string[];
  visualGeometryChangedNodeIds: string[];
  componentNodeIds: string[];
  figureTreatmentIds: string[];
  activeVisualNeedCount: number;
  requirements: StudioDesignImpactRequirement[];
  summary: string;
}

function frameChanged(source: StudioWebFrame, current: StudioWebFrame) {
  return ["x", "y", "width", "height"].some((field) => Math.abs(source[field as keyof StudioWebFrame] - current[field as keyof StudioWebFrame]) >= MEANINGFUL_DELTA)
    || Math.abs(source.rotation - current.rotation) >= 1;
}

function nonzeroCrop(crop: { left: number; top: number; right: number; bottom: number } | undefined) {
  return Boolean(crop && [crop.left, crop.top, crop.right, crop.bottom].some((value) => value >= .01));
}

function requirementFor(need: StudioVisualNeed, slide: StudioWebSlide, impact: Omit<StudioDesignImpact, "requirements" | "activeVisualNeedCount" | "summary">): StudioDesignImpactRequirement {
  const treatments = slide.figureTreatments;
  const visualChange = impact.visualGeometryChangedNodeIds.length > 0;
  const framedFigure = treatments.some((item) => item.mode === "preserve-and-frame" || item.mode === "hybrid-rebuild" || item.mode === "redraw-candidate" || nonzeroCrop(item.crop));
  const verifiedDiagram = treatments.some((item) =>
    ["hybrid-rebuild", "redraw-candidate"].includes(item.mode)
    && item.verificationStatus === "verified"
    && item.relationshipPolicy === "editable-diagram"
    && (item.relationships?.length ?? 0) > 0,
  );
  let passed = false;
  let reason = "";
  if (need.type === "layout-concept" || need.type === "supporting-visual") {
    passed = impact.level === "layout-redesign" || impact.level === "full-redesign";
    reason = passed ? "The slide uses a materially new semantic composition." : "Choose and apply a materially different shared recipe or approved Template Pack layout; typography-only changes do not reconstruct this concept.";
  } else if (need.type === "figure-concept") {
    passed = framedFigure && (visualChange || impact.level === "figure-redesign" || impact.level === "full-redesign");
    reason = passed ? "The figure is treated as a first-class evidence unit in a changed composition." : "Record a relationship-preserving figure treatment and materially reframe or recompose the source visual evidence.";
  } else if (need.type === "image-treatment") {
    passed = visualChange || treatments.some((item) => nonzeroCrop(item.crop));
    reason = passed ? "The source image received a material crop, scale, or placement treatment." : "Apply a material crop, scale, or placement treatment to the source image; changing only surrounding type is insufficient.";
  } else {
    passed = verifiedDiagram && visualChange;
    reason = passed ? "The verified relationship graph was reconstructed with a material geometry change." : "A diagram rebuild requires verified relationships, an editable-diagram policy, and a material geometry reconstruction—not visual guesswork.";
  }
  return { visualNeedId: need.id, type: need.type, passed, reason };
}

export function analyzeStudioDesignImpact(slide: StudioWebSlide): StudioDesignImpact {
  const visible = slide.nodes.filter((node) => node.visible);
  const geometryChangedNodeIds = visible.filter((node) => frameChanged(node.sourceFrame, node.frame)).map((node) => node.id);
  const visualGeometryChangedNodeIds = visible.filter((node) => ["image", "native-object", "shape", "connector"].includes(node.kind) && frameChanged(node.sourceFrame, node.frame)).map((node) => node.id);
  const componentNodeIds = visible.filter((node) => Boolean(node.component)).map((node) => node.id);
  const figureTreatmentIds = slide.figureTreatments.filter((item) => item.mode !== "preserve-as-unit" || Boolean(item.groupFrame) || nonzeroCrop(item.crop)).map((item) => item.id);
  const contentNodeCount = Math.max(1, visible.filter((node) => !["footer", "slide-number", "date", "logo"].includes(node.role)).length);
  const minimumLayoutChange = Math.max(1, Math.ceil(contentNodeCount * .25));
  const layoutChanged = slide.recipe !== "source" && (geometryChangedNodeIds.length >= minimumLayoutChange || componentNodeIds.length >= minimumLayoutChange || slide.recipe === "template-layout");
  const figureChanged = figureTreatmentIds.length > 0 && (visualGeometryChangedNodeIds.length > 0 || slide.figureTreatments.some((item) => item.mode === "preserve-and-frame" || Boolean(item.groupFrame) || nonzeroCrop(item.crop)));
  let level: StudioDesignImpactLevel;
  if (layoutChanged && figureChanged) level = "full-redesign";
  else if (figureChanged) level = "figure-redesign";
  else if (layoutChanged) level = "layout-redesign";
  else if (geometryChangedNodeIds.length) level = "cleanup";
  else if (slide.status === "designed") level = "typography-only";
  else level = "unchanged";
  const base = {
    level,
    meaningful: ["layout-redesign", "figure-redesign", "full-redesign"].includes(level),
    geometryChangedNodeIds,
    visualGeometryChangedNodeIds,
    componentNodeIds,
    figureTreatmentIds,
  };
  const activeNeeds = (slide.visualNeeds ?? []).filter((need) => !["held", "resolved"].includes(need.status));
  const requirements = activeNeeds.map((need) => requirementFor(need, slide, base));
  const failed = requirements.filter((item) => !item.passed);
  const summary = failed.length
    ? `${level.replaceAll("-", " ")} · ${failed.length} active visual need${failed.length === 1 ? "" : "s"} still require material reconstruction`
    : `${level.replaceAll("-", " ")} · ${geometryChangedNodeIds.length} materially changed frame${geometryChangedNodeIds.length === 1 ? "" : "s"}`;
  return { ...base, activeVisualNeedCount: activeNeeds.length, requirements, summary };
}
