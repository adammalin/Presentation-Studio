import type { NativeMeasurementResult } from "./desktop";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { applyStudioLayoutConstraints, type StudioConstraintRequest } from "./studio-layout-constraints";
import { critiqueStudioSlide, type StudioVisualCritique } from "./studio-visual-critic";
import { updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "./studio-web-scene";
import type { StudioQualityIssue, StudioWebFrame, StudioWebNode, StudioWebScene } from "../types";

const PT = 12_700;

export interface StudioRepairAction {
  issueId: string;
  status: "fixed" | "deferred" | "already-satisfied";
  operation: "fit-safe-region" | "reapply-constraint" | "grow-text-frame" | "restore-title-hierarchy" | "human-or-ai-design";
  nodeIds: string[];
  message: string;
}

export interface StudioRepairPassResult {
  scene: StudioWebScene;
  critique: StudioVisualCritique;
  actions: StudioRepairAction[];
  changedNodeIds: string[];
  fixedIssueIds: string[];
  deferredIssueIds: string[];
  requiresNativeRerender: boolean;
}

function sameIds(left: string[], right: string[]) {
  return left.length === right.length && left.every((id) => right.includes(id));
}

function intersects(left: StudioWebFrame, right: StudioWebFrame) {
  return left.x < right.x + right.width && left.x + left.width > right.x && left.y < right.y + right.height && left.y + left.height > right.y;
}

function nativeShape(node: StudioWebNode, measurement: NativeMeasurementResult) {
  return measurement.slides.flatMap((slide) => slide.shapes).find((shape) => shape.name === node.id || shape.name?.endsWith(` · ${node.id}`));
}

function completeRelationshipGroup(scene: StudioWebScene, slideNumber: number, nodeIds: string[]) {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const treatment = slide?.figureTreatments.find((item) => nodeIds.some((id) => item.nodeIds.includes(id)));
  return treatment ? { nodeIds: [...treatment.nodeIds], groups: [[...treatment.nodeIds]] } : { nodeIds: [...nodeIds], groups: nodeIds.map((id) => [id]) };
}

function minimumNativeTextGrowth(scene: StudioWebScene, slideNumber: number, node: StudioWebNode, measurement: NativeMeasurementResult): StudioWebFrame | undefined {
  if (node.kind !== "text" || node.locked) return undefined;
  const shape = nativeShape(node, measurement);
  const bounds = shape?.boundsPt;
  const text = shape?.text?.renderedBoundsPt;
  const margins = shape?.text?.marginsPt;
  if (!bounds || !text || !margins || shape?.text?.coordinateSpace !== "slide") return undefined;
  const innerLeft = bounds.left + margins.left;
  const innerRight = bounds.left + bounds.width - margins.right;
  if (text.left < innerLeft - 2.25 || text.left + text.width > innerRight + 2.25) return undefined;
  const requiredHeightPt = Math.max(bounds.height, text.height + margins.top + margins.bottom + 2.25);
  if (requiredHeightPt <= bounds.height + .25) return undefined;
  const delta = Math.round((requiredHeightPt - bounds.height) * PT);
  const anchor = shape.text?.verticalAnchor.toLowerCase() ?? "top";
  const frame = { ...node.frame, height: node.frame.height + delta, y: node.frame.y - (anchor.includes("bottom") ? delta : anchor.includes("middle") ? Math.round(delta / 2) : 0) };
  const safe = (scene.rhythm?.safeMarginPt ?? PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt) * PT;
  if (frame.y < safe || frame.y + frame.height > scene.slideSize.height - safe) return undefined;
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber)!;
  const obstacles = slide.nodes.filter((candidate) => candidate.visible && candidate.id !== node.id);
  if (obstacles.some((candidate) => !intersects(node.frame, candidate.frame) && intersects(frame, candidate.frame))) return undefined;
  return frame;
}

function selectedIssues(critique: StudioVisualCritique, requestedIssueIds?: string[]) {
  if (!requestedIssueIds?.length) return critique.issues.filter((issue) => issue.autoFixable);
  const unique = [...new Set(requestedIssueIds)];
  const missing = unique.filter((id) => !critique.issues.some((issue) => issue.id === id));
  if (missing.length) throw new Error(`The requested Studio issue IDs are stale or unavailable: ${missing.join(", ")}.`);
  return critique.issues.filter((issue) => unique.includes(issue.id));
}

export function applyStudioDeterministicRepairPass(scene: StudioWebScene, slideNumber: number, measurement: NativeMeasurementResult, requestedIssueIds?: string[]): StudioRepairPassResult {
  const critique = critiqueStudioSlide(scene, slideNumber, measurement);
  const issues = selectedIssues(critique, requestedIssueIds);
  const originalSlide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!originalSlide) throw new Error("The requested Studio slide is unavailable.");
  let next = scene;
  const actions: StudioRepairAction[] = [];
  const changed = new Set<string>();

  const defer = (issue: StudioQualityIssue, message: string) => actions.push({ issueId: issue.id, status: "deferred", operation: "human-or-ai-design", nodeIds: issue.nodeIds, message });
  for (const issue of issues) {
    try {
      if (issue.category === "safe-region" && issue.nodeIds.length) {
        const group = completeRelationshipGroup(next, slideNumber, issue.nodeIds);
        const request: StudioConstraintRequest = { kind: "fit-safe-region", mode: "both", nodeIds: group.nodeIds, groups: group.groups, rationale: `Repair ${issue.id} from the matching PowerPoint-native issue ledger without breaking a figure or component relationship.`, author: "ai" };
        const result = applyStudioLayoutConstraints(next, slideNumber, [request], measurement);
        next = result.scene;
        result.changedNodeIds.forEach((id) => changed.add(id));
        actions.push({ issueId: issue.id, status: result.changedNodeIds.length ? "fixed" : "already-satisfied", operation: "fit-safe-region", nodeIds: group.nodeIds, message: result.changedNodeIds.length ? "Moved the complete relationship group minimally into the resolved safe region." : "The matching scene already satisfies the safe region." });
        continue;
      }
      if (["alignment", "spacing"].includes(issue.category) && issue.nodeIds.length) {
        const constraint = (originalSlide.constraints ?? []).find((item) => sameIds(item.nodeIds, issue.nodeIds) && (issue.category === "alignment" ? item.kind === "align" : item.kind === "distribute"));
        if (!constraint) { defer(issue, "No revision-bound high-level constraint remains to replay; use a deliberate Studio alignment or distribution decision."); continue; }
        const request: StudioConstraintRequest = { kind: constraint.kind, mode: constraint.mode, nodeIds: [...constraint.nodeIds], groups: constraint.groups?.map((group) => [...group]), anchorNodeId: constraint.anchorNodeId, gridPt: constraint.gridPt, rationale: `Reapply ${constraint.kind} ${constraint.mode} from the native issue ledger.`, author: "ai" };
        const result = applyStudioLayoutConstraints(next, slideNumber, [request], measurement);
        next = result.scene;
        result.changedNodeIds.forEach((id) => changed.add(id));
        actions.push({ issueId: issue.id, status: result.changedNodeIds.length ? "fixed" : "already-satisfied", operation: "reapply-constraint", nodeIds: constraint.nodeIds, message: result.changedNodeIds.length ? "Reapplied the recorded semantic constraint using PowerPoint-native optical evidence." : "The matching scene already satisfies the recorded constraint." });
        continue;
      }
      if (issue.category === "overflow" && issue.nodeIds.length === 1) {
        const slide = next.slides.find((item) => item.slideNumber === slideNumber)!;
        const node = slide.nodes.find((item) => item.id === issue.nodeIds[0]);
        const insideFigure = slide.figureTreatments.some((item) => item.nodeIds.includes(issue.nodeIds[0]));
        const frame = node && !insideFigure ? minimumNativeTextGrowth(next, slideNumber, node, measurement) : undefined;
        if (!node || !frame) { defer(issue, insideFigure ? "The overflowing text belongs to a relationship-preserving figure; recompose the complete figure instead of growing one member." : "The minimum measured text-frame growth would cross a safe region, create a new overlap, or cannot solve horizontal clipping; choose a larger semantic region."); continue; }
        next = updateStudioWebNodeFrame(next, slideNumber, node.id, frame);
        changed.add(node.id);
        actions.push({ issueId: issue.id, status: "fixed", operation: "grow-text-frame", nodeIds: [node.id], message: "Grew the editable text frame by the minimum PowerPoint-measured amount without crossing the safe region or creating a new collision." });
        continue;
      }
      if (issue.category === "hierarchy" && issue.nodeIds.length === 1) {
        const slide = next.slides.find((item) => item.slideNumber === slideNumber)!;
        const title = slide.nodes.find((item) => item.id === issue.nodeIds[0]);
        if (!title || title.kind !== "text" || title.locked) { defer(issue, "The title hierarchy cannot be edited safely in this scene revision."); continue; }
        const largestBody = Math.max(0, ...slide.nodes.filter((node) => node.visible && node.kind === "text" && ["body", "label", "caption"].includes(node.role)).map((node) => node.style.fontSizePt));
        const target = Math.min(PRESENTATION_DESIGN_STANDARD.defaults.typography.headlineMaximumPt, Math.max(title.style.fontSizePt, largestBody + 4));
        if (target <= title.style.fontSizePt + .01) { defer(issue, "The ORNL headline maximum cannot restore a clear hierarchy without recomposing competing content."); continue; }
        next = updateStudioWebNodeStyle(next, slideNumber, title.id, { fontSizePt: target, fontWeight: 700 });
        changed.add(title.id);
        actions.push({ issueId: issue.id, status: "fixed", operation: "restore-title-hierarchy", nodeIds: [title.id], message: `Raised the editable title to ${target} pt within the ORNL headline maximum; the fresh render must confirm fit.` });
        continue;
      }
      defer(issue, "This issue requires a material composition, table, figure, concept, or human decision rather than a deterministic geometry patch.");
    } catch (error) {
      defer(issue, error instanceof Error ? error.message : "The deterministic repair could not be applied safely.");
    }
  }

  const fixedIssueIds = actions.filter((item) => item.status === "fixed" || item.status === "already-satisfied").map((item) => item.issueId);
  const deferredIssueIds = actions.filter((item) => item.status === "deferred").map((item) => item.issueId);
  return { scene: next, critique, actions, changedNodeIds: [...changed], fixedIssueIds, deferredIssueIds, requiresNativeRerender: changed.size > 0 };
}
