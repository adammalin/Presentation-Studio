import designStandard from "../../shared/presentation-design-standard.json";
import type { ResolvedDesignProfile, TemplateDecisionSource } from "../types";

export interface PresentationDesignStandard {
  schema: "presentation-studio/design-standard";
  version: string;
  name: string;
  status: "product-default-draft";
  mission: string;
  defaultMode: { name: string; behavior: string; definitionOfDone: string };
  defaults: {
    profileId: string;
    profileLabel: string;
    slide: { aspectRatio: "16:9"; widthInches: number; heightInches: number };
    typography: { family: "Aptos"; headlineMaximumPt: number; subheadlineMinimumPt: number; subheadlineMaximumPt: number; bodyMinimumPt: number; captionMinimumPt: number };
    template: { id: string; policy: string };
    contentPolicy: "preserve-exact";
    sourceProtection: "read-only-source-new-output";
    geometry: { brandContainerRadius: 0 };
    palette: Record<string, string>;
  };
  autonomy: { routineDecisions: string[]; approvalPolicy: string };
  componentSystem: {
    version: number;
    principle: string;
    page: Record<string, number>;
    title: Record<string, number>;
    patterns: Record<string, string>;
    paragraph: Record<string, number>;
    components: Record<string, string>;
    layoutRecipes: Record<string, { useWhen: string; regions: string }>;
  };
  tableProfile: {
    id: string;
    version: number;
    authority: string;
    nativeEditable: boolean;
    cornerRadius: number;
    fontFamily: string;
    header: { fontSizePt: number; fontWeight: number; fill: string; textColor: string; verticalAlignment: string };
    body: { fontSizePt: number; minimumFontSizePt: number; textColor: string; fill: string; alternateFill: string };
    caption: { minimumFontSizePt: number };
    cellPaddingPt: { left: number; right: number; top: number; bottom: number };
    strokes: { outer: string; horizontal: { color: string; widthPt: number }; vertical: string };
    alignment: Record<string, string>;
    overflowPolicy: string;
    preserve: string[];
  };
  tableVariants: {
    standard: { headerFontSizePt: number; bodyFontSizePt: number; horizontalPaddingPt: number; verticalPaddingPt: number };
    denseTechnical: { headerFontSizePt: number; bodyFontSizePt: number; horizontalPaddingPt: number; verticalPaddingPt: number; useWhen: string };
  };
  preservationRules: string[];
  everySlideChecklist: string[];
  collaborationRules: string[];
  tableRules: string[];
  ornlRules: string[];
  visualQaLoop: string[];
  askOnlyWhen: string[];
  reviewOutput: string[];
}

export const PRESENTATION_DESIGN_STANDARD = designStandard as PresentationDesignStandard;

export function defaultProjectDesignSettings() {
  return {
    designStandardVersion: PRESENTATION_DESIGN_STANDARD.version,
    defaultProfileId: PRESENTATION_DESIGN_STANDARD.defaults.profileId,
    defaultSlideSize: PRESENTATION_DESIGN_STANDARD.defaults.slide.aspectRatio,
    defaultFontFamily: PRESENTATION_DESIGN_STANDARD.defaults.typography.family,
  } as const;
}

export function designStandardSummary() {
  const defaults = PRESENTATION_DESIGN_STANDARD.defaults;
  return `${defaults.profileLabel} · ${defaults.slide.aspectRatio} · ${defaults.typography.family} · exact content`;
}

export function createOrnlDesignProfile(source: TemplateDecisionSource, adoptedAt = new Date().toISOString()): ResolvedDesignProfile {
  const defaults = PRESENTATION_DESIGN_STANDARD.defaults;
  return {
    id: defaults.profileId,
    standardVersion: PRESENTATION_DESIGN_STANDARD.version,
    templateId: defaults.template.id,
    slideSize: defaults.slide.aspectRatio,
    fontFamily: defaults.typography.family,
    contentPolicy: defaults.contentPolicy,
    adoptedAt,
    source,
    customized: false,
  };
}
