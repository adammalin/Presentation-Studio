import type { DeckJob, DesignThread } from "../types";
import type { SlideDesignMetrics } from "./design-metrics";

export const DESIGN_REPAIR_LEDGER_SCHEMA = "presentation-studio/design-repair-ledger" as const;

export type DesignRepairIssueCategory = "alignment" | "fit" | "safe-area" | "table" | "figure" | "layout" | "human-feedback" | "intent";

export interface DesignRepairIssue {
  id: string;
  slideNumber: number;
  category: DesignRepairIssueCategory;
  severity: "info" | "warning" | "error";
  status: "found" | "fixing" | "fixed" | "needs-review";
  summary: string;
  evidence: string;
  suggestedTreatment: string;
  autoFixable: boolean;
}

function metricIssues(slideNumber: number, metrics?: SlideDesignMetrics): DesignRepairIssue[] {
  if (!metrics) return [];
  const issues: DesignRepairIssue[] = [];
  if (metrics.textOverflowCount > 0) issues.push({ id: `slide-${slideNumber}-native-overflow`, slideNumber, category: "fit", severity: "error", status: "found", summary: `${metrics.textOverflowCount} PowerPoint-native text frame${metrics.textOverflowCount === 1 ? "" : "s"} overflow.`, evidence: "Rendered text escapes its measured outer PowerPoint frame.", suggestedTreatment: "Choose a better shared layout or grow/reflow the measured frame; never silently shrink or rewrite approved text.", autoFixable: true });
  if (metrics.offSlideObjectCount > 0 || metrics.safeRegionViolationCount > 0) issues.push({ id: `slide-${slideNumber}-safe-region`, slideNumber, category: "safe-area", severity: metrics.offSlideObjectCount > 0 ? "error" : "warning", status: "found", summary: `${metrics.offSlideObjectCount} off-slide and ${metrics.safeRegionViolationCount} safe-region placement issue${metrics.offSlideObjectCount + metrics.safeRegionViolationCount === 1 ? "" : "s"}.`, evidence: "PowerPoint-native object geometry crosses the slide or resolved safe region.", suggestedTreatment: "Fit the complete semantic group into the approved region with minimum movement.", autoFixable: true });
  if ((metrics.opticalLeftRangePt ?? 0) > 4) issues.push({ id: `slide-${slideNumber}-optical-alignment`, slideNumber, category: "alignment", severity: "warning", status: "found", summary: "Visible text starts do not share a stable optical edge.", evidence: `PowerPoint-native optical left-edge range is ${metrics.opticalLeftRangePt?.toFixed(2)} pt.`, suggestedTreatment: "Align visible text starts using native insets and bullet margins, not only shape coordinates.", autoFixable: true });
  if (metrics.tableCellFindings.length > 0) issues.push({ id: `slide-${slideNumber}-table-cells`, slideNumber, category: "table", severity: metrics.tableCellFindings.some((finding) => finding.severity === "error") ? "error" : "warning", status: "found", summary: `${metrics.tableCellFindings.length} table-cell legibility issue${metrics.tableCellFindings.length === 1 ? "" : "s"}.`, evidence: metrics.tableCellFindings.slice(0, 3).map((finding) => finding.evidence).join(" "), suggestedTreatment: "Run the native table solver while preserving exact cells, spans, semantic colors, and visible breaks.", autoFixable: true });
  return issues;
}

export function buildDesignRepairLedger(input: { deck: DeckJob; slideNumber: number; representation: "current" | "proposal"; metrics?: SlideDesignMetrics; threads?: DesignThread[] }) {
  const slide = input.deck.audit?.slides.find((item) => item.number === input.slideNumber);
  if (!slide) throw new Error("The requested slide is not present in the current audit.");
  const issues: DesignRepairIssue[] = [
    ...metricIssues(input.slideNumber, input.metrics),
    ...(input.deck.audit?.findings ?? []).filter((finding) => finding.slideNumber === input.slideNumber).map((finding): DesignRepairIssue => ({ id: finding.id, slideNumber: input.slideNumber, category: finding.category === "table" ? "table" : finding.category === "figure" ? "figure" : finding.category === "layout" ? "layout" : "fit", severity: finding.severity, status: finding.autoFixable ? "found" : "needs-review", summary: finding.message, evidence: finding.evidence, suggestedTreatment: finding.autoFixable ? "Use the bounded deterministic repair declared by the design work order." : "Preserve the source and include this item in human review.", autoFixable: finding.autoFixable })),
    ...(input.threads ?? []).filter((thread) => thread.deckId === input.deck.id && thread.slideNumber === input.slideNumber && ["submitted", "needs-reanchor"].includes(thread.status)).map((thread): DesignRepairIssue => ({ id: `thread-${thread.id}`, slideNumber: input.slideNumber, category: "human-feedback", severity: "warning", status: thread.status === "needs-reanchor" ? "needs-review" : "found", summary: thread.comment, evidence: `Human comment is bound to normalized slide region ${thread.anchor.x.toFixed(3)}, ${thread.anchor.y.toFixed(3)}, ${thread.anchor.width.toFixed(3)}, ${thread.anchor.height.toFixed(3)}.`, suggestedTreatment: thread.status === "needs-reanchor" ? "Ask the user to re-anchor this exact comment." : "Stage a bounded fix and clear only this exact thread when fully addressed.", autoFixable: thread.status === "submitted" })),
  ];
  const sourceVisualObjectIds = input.deck.scene?.objects.filter((object) => object.slideNumber === input.slideNumber && ["picture", "group", "graphic-frame", "connector"].includes(object.kind)).map((object) => object.id) ?? [];
  const relationshipObjectIds = input.deck.scene?.objects.filter((object) => object.slideNumber === input.slideNumber && ["group", "connector"].includes(object.kind)).map((object) => object.id) ?? [];
  if (slide.pictureCount > 0 || slide.connectorCount > 0) issues.push({ id: `slide-${input.slideNumber}-intent-translation`, slideNumber: input.slideNumber, category: "intent", severity: "warning", status: "needs-review", summary: "The proposal must be checked against the original figure's message and technical relationships.", evidence: `The source contains ${slide.pictureCount} picture${slide.pictureCount === 1 ? "" : "s"} and ${slide.connectorCount} connector${slide.connectorCount === 1 ? "" : "s"}; exact text hash ${slide.textHash}.`, suggestedTreatment: "Compare Current and Proposal pixels, confirm every meaning-bearing label/value/relationship, and record an intent review before calling the design better.", autoFixable: false });
  const unique = [...new Map(issues.map((issue) => [issue.id, issue])).values()];
  return {
    schema: DESIGN_REPAIR_LEDGER_SCHEMA,
    version: 1 as const,
    slideNumber: input.slideNumber,
    representation: input.representation,
    phase: input.representation === "current" ? "found-issues" as const : "rechecking-original-intent" as const,
    issueCount: unique.length,
    autoFixableCount: unique.filter((issue) => issue.autoFixable).length,
    needsReviewCount: unique.filter((issue) => issue.status === "needs-review").length,
    issues: unique,
    originalIntentReference: {
      sourceTextHash: slide.textHash,
      exactVisibleText: slide.text,
      sourceVisualObjectIds,
      relationshipObjectIds,
      requiredChecks: ["exact approved wording remains", "source visual identity remains or a verified replacement is disclosed", "meaning-bearing labels and values remain", "arrows, sequence, causality, and grouping retain the original message"],
    },
    instruction: input.representation === "current" ? "Report the found issues, fix the bounded items as one coherent composition, then inspect Proposal." : "Recheck the Proposal against this original-intent reference. A prettier slide is not better if technical meaning, labels, values, or relationships drifted.",
  };
}
