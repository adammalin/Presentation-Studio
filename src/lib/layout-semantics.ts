import type { TemplateLayoutPreview, TemplatePreviewElement } from "./template-catalog";

export type TemplateSemanticRole = "title" | "subtitle" | "body" | "caption" | "image" | "table" | "chart" | "media" | "content" | "footer" | "date" | "slide-number" | "other";
export type TemplateContentKind = "text" | "image" | "table" | "chart" | "media";
export type TemplateLayoutIntent = "cover" | "section" | "assertion" | "content" | "comparison" | "visual" | "data" | "conclusion" | "blank";
export type TemplateSlotAlignmentIntent = "optical-left" | "contain" | "fill" | "structural";

export interface TemplateSlotBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TemplateSemanticSlot {
  id: string;
  role: TemplateSemanticRole;
  placeholderType: string;
  placeholderIndex?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  acceptedContent: TemplateContentKind[];
  allowedObjectKinds: TemplateContentKind[];
  required: boolean;
  capacity: "micro" | "short" | "medium" | "long" | "visual";
  preferredBounds: TemplateSlotBounds;
  minimumBounds: TemplateSlotBounds;
  maximumBounds: TemplateSlotBounds;
  alignmentIntent: TemplateSlotAlignmentIntent;
  paddingIntentPt: { top: number; right: number; bottom: number; left: number };
  priority: number;
}

export interface TemplateLayoutSemantics {
  intent: TemplateLayoutIntent;
  slots: TemplateSemanticSlot[];
  capabilities: {
    title: boolean;
    subtitle: boolean;
    bodySlots: number;
    captionSlots: number;
    imageSlots: number;
    tableSlots: number;
    chartSlots: number;
    mediaSlots: number;
    flexibleContentSlots: number;
    columns: number;
  };
  constraints: {
    maxTitleLines: number;
    preferredBodyDensity: "low" | "medium" | "high";
    minimumBodyFontPt: number;
    minimumCaptionFontPt: number;
    requiresVisual: boolean;
    supportsDenseText: boolean;
  };
  summary: string;
}

export interface LayoutContentProfile {
  titleCharacterCount: number;
  bodyBlockCount: number;
  captionBlockCount: number;
  bodyCharacterCount: number;
  imageCount: number;
  imageAspectRatios?: number[];
  tableCount: number;
  chartCount: number;
  mediaCount: number;
  desiredIntent?: TemplateLayoutIntent;
}

export interface LayoutCompatibilityResult {
  layoutId: string;
  layoutName: string;
  score: number;
  status: "recommended" | "compatible" | "poor" | "incompatible";
  reasons: string[];
  unmetNeeds: string[];
}

const SYSTEM_PLACEHOLDERS = new Set(["ftr", "dt", "sldNum", "hdr"]);

function placeholderKey(element: TemplatePreviewElement): string {
  const round = (value: number) => Math.round(value / 1000);
  return [element.placeholderType, element.placeholderIndex ?? "", round(element.x), round(element.y), round(element.width), round(element.height)].join("|");
}

export function uniquePlaceholderElements(elements: TemplatePreviewElement[]): TemplatePreviewElement[] {
  const seen = new Set<string>();
  return elements.filter((element) => {
    if (!element.placeholderType) return false;
    const key = placeholderKey(element);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function roleFor(element: TemplatePreviewElement, slideHeight: number): TemplateSemanticRole {
  const type = element.placeholderType ?? "other";
  if (["title", "ctrTitle"].includes(type)) return "title";
  if (type === "subTitle") return "subtitle";
  if (type === "pic") return "image";
  if (type === "tbl") return "table";
  if (type === "chart") return "chart";
  if (type === "media") return "media";
  if (type === "obj") return "content";
  if (type === "ftr" || type === "hdr") return "footer";
  if (type === "dt") return "date";
  if (type === "sldNum") return "slide-number";
  if (type === "body") return element.height / slideHeight < 0.105 ? "caption" : "body";
  return "other";
}

function acceptedContentFor(role: TemplateSemanticRole): TemplateContentKind[] {
  if (["title", "subtitle", "caption", "footer", "date", "slide-number"].includes(role)) return ["text"];
  if (role === "body") return ["text", "image", "table", "chart"];
  if (role === "image") return ["image"];
  if (role === "table") return ["table"];
  if (role === "chart") return ["chart"];
  if (role === "media") return ["media"];
  if (role === "content") return ["text", "image", "table", "chart", "media"];
  return ["text"];
}

function capacityFor(element: TemplatePreviewElement, role: TemplateSemanticRole, slideWidth: number, slideHeight: number): TemplateSemanticSlot["capacity"] {
  if (["image", "table", "chart", "media", "content"].includes(role)) return "visual";
  const area = (element.width * element.height) / (slideWidth * slideHeight);
  const height = element.height / slideHeight;
  if (area < 0.03 || height < 0.075) return "micro";
  if (area < 0.09 || height < 0.16) return "short";
  if (area < 0.24 || height < 0.34) return "medium";
  return "long";
}

function responsiveSlotContract(element: TemplatePreviewElement, role: TemplateSemanticRole) {
  const preferredBounds = { x: element.x, y: element.y, width: element.width, height: element.height };
  const scale = ["title", "subtitle"].includes(role) ? { width: .55, height: .5 }
    : ["body", "caption"].includes(role) ? { width: .45, height: .35 }
      : ["image", "table", "chart", "media", "content"].includes(role) ? { width: .4, height: .4 }
        : { width: .6, height: .5 };
  const minimumBounds = { ...preferredBounds, width: Math.round(element.width * scale.width), height: Math.round(element.height * scale.height) };
  const alignmentIntent: TemplateSlotAlignmentIntent = ["title", "subtitle", "body", "caption"].includes(role) ? "optical-left"
    : role === "image" || role === "media" ? "contain"
      : ["table", "chart", "content"].includes(role) ? "fill"
        : "structural";
  const padding = ["body", "caption", "content"].includes(role) ? 8 : 0;
  const priority = role === "title" ? 100
    : role === "subtitle" ? 90
      : ["body", "table", "chart", "image", "media", "content"].includes(role) ? 80
        : role === "caption" ? 60
          : 10;
  return {
    preferredBounds,
    minimumBounds,
    maximumBounds: { ...preferredBounds },
    alignmentIntent,
    paddingIntentPt: { top: padding, right: padding, bottom: padding, left: padding },
    priority,
  };
}

function inferIntent(name: string, slots: TemplateSemanticSlot[], category: TemplateLayoutPreview["category"]): TemplateLayoutIntent {
  const normalized = name.toLowerCase();
  const imageSlots = slots.filter((slot) => slot.role === "image").length;
  const dataSlots = slots.filter((slot) => ["table", "chart"].includes(slot.role)).length;
  const bodySlots = slots.filter((slot) => ["body", "content"].includes(slot.role)).length;
  if (normalized.includes("conclusion") || category === "conclusion") return "conclusion";
  if (normalized.includes("section")) return "section";
  if (normalized.startsWith("title") || category === "title") return "cover";
  if (dataSlots > 0 || normalized.includes("table") || normalized.includes("chart") || normalized.includes("data")) return "data";
  if (imageSlots > 0 || category === "image" || normalized.includes("portrait") || normalized.includes("image")) return "visual";
  if (bodySlots > 1 || /[2-9]-column/.test(normalized)) return "comparison";
  if (slots.filter((slot) => !SYSTEM_PLACEHOLDERS.has(slot.placeholderType)).length === 0) return "blank";
  if (normalized.includes("assert")) return "assertion";
  return "content";
}

function summaryFor(intent: TemplateLayoutIntent, capabilities: TemplateLayoutSemantics["capabilities"]): string {
  const parts: string[] = [];
  if (capabilities.bodySlots) parts.push(`${capabilities.bodySlots} text region${capabilities.bodySlots === 1 ? "" : "s"}`);
  if (capabilities.imageSlots) parts.push(`${capabilities.imageSlots} image${capabilities.imageSlots === 1 ? "" : "s"}`);
  if (capabilities.tableSlots) parts.push(`${capabilities.tableSlots} table${capabilities.tableSlots === 1 ? "" : "s"}`);
  if (capabilities.chartSlots) parts.push(`${capabilities.chartSlots} chart${capabilities.chartSlots === 1 ? "" : "s"}`);
  if (capabilities.captionSlots) parts.push(`${capabilities.captionSlots} short label${capabilities.captionSlots === 1 ? "" : "s"}`);
  return `${intent[0].toUpperCase()}${intent.slice(1)} layout${parts.length ? ` · ${parts.join(" · ")}` : ""}`;
}

export function deriveLayoutSemantics(layout: Pick<TemplateLayoutPreview, "id" | "name" | "category" | "elements">, slideWidth: number, slideHeight: number): TemplateLayoutSemantics {
  const slots = uniquePlaceholderElements(layout.elements).map((element, index): TemplateSemanticSlot => {
    const role = roleFor(element, slideHeight);
    const acceptedContent = acceptedContentFor(role);
    return {
      id: `${layout.id}-slot-${index + 1}`,
      role,
      placeholderType: element.placeholderType ?? "other",
      placeholderIndex: element.placeholderIndex,
      x: element.x,
      y: element.y,
      width: element.width,
      height: element.height,
      acceptedContent,
      allowedObjectKinds: acceptedContent,
      required: role === "title",
      capacity: capacityFor(element, role, slideWidth, slideHeight),
      ...responsiveSlotContract(element, role),
    };
  });
  const count = (role: TemplateSemanticRole) => slots.filter((slot) => slot.role === role).length;
  const bodySlots = count("body") + count("content");
  const capabilities = {
    title: count("title") > 0,
    subtitle: count("subtitle") > 0,
    bodySlots,
    captionSlots: count("caption"),
    imageSlots: count("image"),
    tableSlots: count("table"),
    chartSlots: count("chart"),
    mediaSlots: count("media"),
    flexibleContentSlots: count("content"),
    columns: bodySlots > 0 ? bodySlots : 0,
  };
  const intent = inferIntent(layout.name, slots, layout.category);
  const requiresVisual = intent === "visual" && capabilities.imageSlots > 0;
  const supportsDenseText = bodySlots === 1 && slots.some((slot) => ["body", "content"].includes(slot.role) && slot.capacity === "long") && !requiresVisual;
  return {
    intent,
    slots,
    capabilities,
    constraints: {
      maxTitleLines: intent === "cover" ? 3 : 2,
      preferredBodyDensity: supportsDenseText ? "high" : ["cover", "visual", "conclusion"].includes(intent) ? "low" : "medium",
      minimumBodyFontPt: 16,
      minimumCaptionFontPt: 14,
      requiresVisual,
      supportsDenseText,
    },
    summary: summaryFor(intent, capabilities),
  };
}

function compatibleVisualCapacity(layout: TemplateLayoutPreview, kind: "image" | "table" | "chart" | "media"): { exact: number; fallback: number } {
  const semantics = layout.semantic;
  if (!semantics) return { exact: 0, fallback: 0 };
  const exact = kind === "image" ? semantics.capabilities.imageSlots : kind === "table" ? semantics.capabilities.tableSlots : kind === "chart" ? semantics.capabilities.chartSlots : semantics.capabilities.mediaSlots;
  return { exact: exact + semantics.capabilities.flexibleContentSlots, fallback: kind === "media" ? 0 : Math.max(0, semantics.capabilities.bodySlots - semantics.capabilities.flexibleContentSlots) };
}

export function scoreLayoutCompatibility(layout: TemplateLayoutPreview, profile: LayoutContentProfile): LayoutCompatibilityResult {
  const semantics = layout.semantic;
  if (!semantics) return { layoutId: layout.id, layoutName: layout.name, score: 0, status: "incompatible", reasons: [], unmetNeeds: ["Layout semantic metadata is unavailable."] };
  let score = 70;
  const reasons: string[] = [];
  const unmetNeeds: string[] = [];
  const needsTitle = profile.titleCharacterCount > 0;
  if (needsTitle && semantics.capabilities.title) { score += 7; reasons.push("Includes an approved title region."); }
  else if (needsTitle) { score -= 42; unmetNeeds.push("No title region for the supplied headline."); }

  const availableBody = semantics.capabilities.bodySlots;
  if (profile.bodyBlockCount > 0 && availableBody === 0) { score -= 38; unmetNeeds.push("No primary text region for the supplied body content."); }
  else if (profile.bodyBlockCount > availableBody) { score -= Math.min(34, (profile.bodyBlockCount - availableBody) * 14); unmetNeeds.push(`Needs ${profile.bodyBlockCount} text regions; layout provides ${availableBody}.`); }
  else if (profile.bodyBlockCount > 0) { score += Math.min(14, profile.bodyBlockCount * 5); reasons.push(`Accommodates ${profile.bodyBlockCount} text region${profile.bodyBlockCount === 1 ? "" : "s"}.`); }

  const availableCaptions = semantics.capabilities.captionSlots + Math.max(0, availableBody - profile.bodyBlockCount);
  if (profile.captionBlockCount > availableCaptions) { score -= Math.min(24, (profile.captionBlockCount - availableCaptions) * 8); unmetNeeds.push(`Needs ${profile.captionBlockCount} short label regions; layout supports ${availableCaptions}.`); }
  else if (profile.captionBlockCount > 0) { score += Math.min(10, profile.captionBlockCount * 3); reasons.push(`Accommodates ${profile.captionBlockCount} short label${profile.captionBlockCount === 1 ? "" : "s"}.`); }

  for (const [kind, count] of [["image", profile.imageCount], ["table", profile.tableCount], ["chart", profile.chartCount], ["media", profile.mediaCount]] as const) {
    if (count <= 0) continue;
    const capacity = compatibleVisualCapacity(layout, kind);
    const totalCapacity = capacity.exact + capacity.fallback;
    if (totalCapacity < count) { score -= Math.min(42, (count - totalCapacity) * 24); unmetNeeds.push(`Needs ${count} ${kind} region${count === 1 ? "" : "s"}; layout supports ${totalCapacity}.`); }
    else if (capacity.exact >= count) { score += Math.min(16, count * 6); reasons.push(`Provides purpose-built ${kind} capacity.`); }
    else { score += Math.min(6, count * 2); reasons.push(`Can compose ${kind} content inside a broad editable content region; native fit QA is still required.`); }
  }

  if (profile.imageCount > 0 && semantics.capabilities.imageSlots > profile.imageCount) score -= Math.min(48, (semantics.capabilities.imageSlots - profile.imageCount) * 12);
  const sourceImageAspects = (profile.imageAspectRatios ?? []).filter((value) => Number.isFinite(value) && value > 0);
  const targetImageAspects = semantics.slots.filter((slot) => slot.role === "image" && slot.height > 0).map((slot) => slot.width / slot.height);
  if (sourceImageAspects.length > 0 && targetImageAspects.length > 0) {
    const mismatch = sourceImageAspects.slice(0, targetImageAspects.length).reduce((sum, sourceAspect) => {
      const closest = Math.min(...targetImageAspects.map((targetAspect) => Math.abs(Math.log(sourceAspect / targetAspect))));
      return sum + closest;
    }, 0) / Math.min(sourceImageAspects.length, targetImageAspects.length);
    if (mismatch < 0.16) { score += 7; reasons.push("Image-slot aspect ratios closely match the current visual geometry."); }
    else if (mismatch < 0.38) score += 3;
    else if (mismatch > 0.8) score -= 7;
  }

  const hasVisual = profile.imageCount + profile.tableCount + profile.chartCount + profile.mediaCount > 0;
  if (semantics.constraints.requiresVisual && !hasVisual) { score -= 28; unmetNeeds.push("This layout depends on a purposeful visual resource."); }
  if (profile.bodyCharacterCount > 1200 && !semantics.constraints.supportsDenseText) { score -= 24; unmetNeeds.push("The supplied body is too dense for this layout at the 16 pt minimum."); }
  else if (profile.bodyCharacterCount > 700 && semantics.constraints.preferredBodyDensity === "low") { score -= 14; unmetNeeds.push("Body density exceeds this layout's preferred range."); }
  else if (profile.bodyCharacterCount > 700 && semantics.constraints.supportsDenseText) { score += 8; reasons.push("Supports a dense single-column reading flow."); }

  if (profile.desiredIntent) {
    if (profile.desiredIntent === semantics.intent) { score += 12; reasons.push(`Matches the requested ${profile.desiredIntent} intent.`); }
    else if (["cover", "conclusion", "section"].includes(profile.desiredIntent)) { score -= 28; unmetNeeds.push(`Requested ${profile.desiredIntent} intent does not match this ${semantics.intent} layout.`); }
    else score -= 8;
  }

  score = Math.max(0, Math.min(100, Math.round(score)));
  const status = score >= 82 ? "recommended" : score >= 60 ? "compatible" : score >= 35 ? "poor" : "incompatible";
  return { layoutId: layout.id, layoutName: layout.name, score, status, reasons: reasons.slice(0, 4), unmetNeeds: unmetNeeds.slice(0, 4) };
}

export function rankLayoutCompatibility(layouts: TemplateLayoutPreview[], profile: LayoutContentProfile): LayoutCompatibilityResult[] {
  return layouts.map((layout) => scoreLayoutCompatibility(layout, profile)).sort((left, right) => right.score - left.score || left.layoutName.localeCompare(right.layoutName));
}
