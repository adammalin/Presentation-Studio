import type { NativeMeasurementResult } from "./desktop";
import { nativeTextOverflows } from "./fresh-composition-qa";
import { analyzeStudioDesignImpact, type StudioDesignImpact } from "./studio-design-impact";
import { studioNodeOpticalBox } from "./studio-layout-constraints";
import type { StudioLayoutConstraint, StudioQualityIssue, StudioWebFrame, StudioWebNode, StudioWebScene } from "../types";

const PT = 12_700;
const ALIGNMENT_TOLERANCE = PT;

export interface StudioVisualCritique {
  slideNumber: number;
  evidenceAuthority: "powerpoint-native";
  designImpact: StudioDesignImpact;
  issues: StudioQualityIssue[];
  blockerCount: number;
  majorCount: number;
  minorCount: number;
  autoFixableCount: number;
  verdict: "ready" | "revise";
  iteration: { currentPass: number; maxPasses: 3; remainingPasses: number };
  checks: string[];
}

function union(frames: StudioWebFrame[]) {
  const x = Math.min(...frames.map((frame) => frame.x));
  const y = Math.min(...frames.map((frame) => frame.y));
  const right = Math.max(...frames.map((frame) => frame.x + frame.width));
  const bottom = Math.max(...frames.map((frame) => frame.y + frame.height));
  return { x, y, width: right - x, height: bottom - y };
}

function issue(index: number, input: Omit<StudioQualityIssue, "id">): StudioQualityIssue {
  return { id: `${input.category}-${index + 1}`, ...input };
}

function groupsFor(constraint: StudioLayoutConstraint, nodes: Map<string, StudioWebNode>) {
  const ids = constraint.groups?.length ? constraint.groups : constraint.nodeIds.map((id) => [id]);
  return ids.map((group) => group.map((id) => nodes.get(id)).filter((node): node is StudioWebNode => Boolean(node)));
}

function constraintDelta(constraint: StudioLayoutConstraint, nodes: Map<string, StudioWebNode>, measurement: NativeMeasurementResult) {
  const groups = groupsFor(constraint, nodes).filter((group) => group.length);
  if (groups.length < 2) return 0;
  if (constraint.kind === "align") {
    const values = groups.map((group) => {
      const optical = union(group.map((node) => ({ ...studioNodeOpticalBox(node, measurement).box, rotation: 0 })));
      const structural = union(group.map((node) => node.frame));
      const box = constraint.mode.startsWith("optical-") ? optical : structural;
      if (["left", "optical-left"].includes(constraint.mode)) return box.x;
      if (constraint.mode === "center") return box.x + box.width / 2;
      if (constraint.mode === "right") return box.x + box.width;
      if (["top", "optical-top"].includes(constraint.mode)) return box.y;
      if (constraint.mode === "middle") return box.y + box.height / 2;
      return box.y + box.height;
    });
    return Math.max(...values) - Math.min(...values);
  }
  if (constraint.kind === "distribute" && groups.length >= 3) {
    const horizontal = constraint.mode === "horizontal-equal-gap";
    const frames = groups.map((group) => union(group.map((node) => node.frame))).sort((left, right) => horizontal ? left.x - right.x : left.y - right.y);
    const gaps = frames.slice(1).map((frame, index) => horizontal ? frame.x - frames[index].x - frames[index].width : frame.y - frames[index].y - frames[index].height);
    return Math.max(...gaps) - Math.min(...gaps);
  }
  return 0;
}

export function critiqueStudioSlide(scene: StudioWebScene, slideNumber: number, measurement: NativeMeasurementResult): StudioVisualCritique {
  if (measurement.status !== "ready" || measurement.authority !== "powerpoint-native") throw new Error("Studio visual critique requires the exact slide revision measured by Microsoft PowerPoint.");
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is unavailable in the current Studio scene.`);
  const issues: StudioQualityIssue[] = [];
  const add = (input: Omit<StudioQualityIssue, "id">) => issues.push(issue(issues.length, input));
  const visible = slide.nodes.filter((node) => node.visible);
  const nodes = new Map(visible.map((node) => [node.id, node]));
  const safe = (scene.rhythm?.safeMarginPt ?? 18) * PT;

  for (const overflow of nativeTextOverflows(measurement)) {
    const node = visible.find((candidate) => overflow.name === candidate.id || overflow.name?.endsWith(` · ${candidate.id}`));
    add({ category: "overflow", severity: "blocker", source: "powerpoint-native", nodeIds: node ? [node.id] : [], message: `${overflow.name || "Text"} renders outside its PowerPoint frame (${overflow.edges.join(", ")}).`, recommendation: "Recompose into a larger semantic region or use measured text fitting; do not silently shrink below the readability floor.", autoFixable: Boolean(node) });
  }

  for (const node of visible) {
    const frame = node.frame;
    if (frame.x < 0 || frame.y < 0 || frame.x + frame.width > scene.slideSize.width || frame.y + frame.height > scene.slideSize.height) add({ category: "safe-region", severity: "blocker", source: "scene", nodeIds: [node.id], message: `${node.name} leaves the slide canvas.`, recommendation: "Fit the complete component into the safe region while preserving its internal relationships.", autoFixable: true });
    else if (!["footer", "slide-number", "date", "logo"].includes(node.role) && (frame.x < safe || frame.y < safe || frame.x + frame.width > scene.slideSize.width - safe || frame.y + frame.height > scene.slideSize.height - safe)) add({ category: "safe-region", severity: "minor", source: "scene", nodeIds: [node.id], message: `${node.name} enters the ${scene.rhythm?.safeMarginPt ?? 18}-point working safe region.`, recommendation: "Confirm intentional full-bleed/template placement or fit the relationship group into the safe region.", autoFixable: true });
    if ((node.kind === "image" || node.kind === "native-object") && (frame.width < 72 * PT || frame.height < 54 * PT)) add({ category: "legibility", severity: "major", source: "scene", nodeIds: [node.id], message: `${node.name} is too small to function as readable technical evidence.`, recommendation: "Give the figure a larger primary or supporting visual region, or crop to the meaning-bearing content with a verified focal point.", autoFixable: false });
  }

  const title = visible.find((node) => node.kind === "text" && node.role === "title");
  const bodySizes = visible.filter((node) => node.kind === "text" && ["body", "label", "caption"].includes(node.role)).map((node) => node.style.fontSizePt);
  if (title && bodySizes.length && title.style.fontSizePt <= Math.max(...bodySizes)) add({ category: "hierarchy", severity: "major", source: "scene", nodeIds: [title.id], message: "The slide title is not typographically dominant over supporting copy.", recommendation: "Use the ORNL title scale or reduce competing headings while preserving readable body type.", autoFixable: true });
  const characterCount = visible.reduce((sum, node) => sum + (node.text?.length ?? node.table?.cells.reduce((cellSum, cell) => cellSum + cell.text.length, 0) ?? 0), 0);
  if (characterCount > 1_100 || visible.length > 24) add({ category: "legibility", severity: "major", source: "scene", nodeIds: [], message: `The composition carries ${characterCount} characters across ${visible.length} visible semantic nodes.`, recommendation: "Use hierarchy, a denser approved table treatment, or a continuation slide; do not miniaturize the entire composition.", autoFixable: false });
  else if (characterCount > 750 || visible.length > 16) add({ category: "legibility", severity: "minor", source: "scene", nodeIds: [], message: "The slide is visually dense and needs a deliberate reading path.", recommendation: "Strengthen grouping and progressive hierarchy while keeping every source statement exact.", autoFixable: false });

  for (const constraint of slide.constraints ?? []) {
    const delta = constraintDelta(constraint, nodes, measurement);
    const source = constraint.kind === "align" && constraint.mode.startsWith("optical-") ? "powerpoint-native" : "scene";
    if (delta > ALIGNMENT_TOLERANCE) add({ category: constraint.kind === "distribute" ? "spacing" : "alignment", severity: "major", source, nodeIds: constraint.nodeIds, message: `${constraint.kind} ${constraint.mode} misses its recorded constraint by ${(delta / PT).toFixed(1)} points in the current built geometry.`, recommendation: "Run refine_studio_layout again against this matching native measurement instead of guessing correction coordinates.", autoFixable: true });
  }

  for (const treatment of slide.figureTreatments) {
    const frame = treatment.groupFrame;
    if (treatment.mode === "redraw-candidate" && treatment.verificationStatus !== "verified") add({ category: "figure", severity: "blocker", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} is marked for redraw but its information and relationships are not verified.`, recommendation: "Keep the original technical object source-locked until a content owner verifies the complete inventory and relationship map.", autoFixable: false });
    if (frame && (frame.width < 108 * PT || frame.height < 72 * PT)) add({ category: "figure", severity: "major", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} is framed too small for reliable review.`, recommendation: "Assign the figure a larger evidence region or crop only after its meaning-bearing area is verified.", autoFixable: false });
    if (treatment.crop && (1 - treatment.crop.left - treatment.crop.right) * (1 - treatment.crop.top - treatment.crop.bottom) < .25) add({ category: "figure", severity: "major", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} retains less than 25% of the source figure area.`, recommendation: "Verify that no labels, values, arrows, or context were removed before accepting the crop.", autoFixable: false });
  }

  const designImpact = analyzeStudioDesignImpact(slide);
  for (const need of (slide.visualNeeds ?? []).filter((item) => !["held", "resolved"].includes(item.status))) {
    const requirement = designImpact.requirements.find((item) => item.visualNeedId === need.id);
    if (need.status === "brief-ready") add({ category: "other", severity: "major", source: "scene", nodeIds: [], message: `The ${need.type.replaceAll("-", " ")} brief is still unfulfilled.`, recommendation: "Attach an explicitly approved concept or hold the brief before declaring this slide ready.", autoFixable: false });
    else if (need.status === "concept-attached") add({ category: "other", severity: "major", source: "scene", nodeIds: [], message: `The ${need.type.replaceAll("-", " ")} concept is attached but has not been reconstructed as editable Studio content.`, recommendation: requirement?.reason ?? "Reconstruct the approved visual characteristics with exact source-bound content before native review.", autoFixable: false });
    else if (requirement && !requirement.passed) add({ category: need.type.includes("figure") || need.type === "diagram-rebuild" ? "figure" : "other", severity: "major", source: "scene", nodeIds: designImpact.geometryChangedNodeIds, message: `The editable result does not materially fulfill the ${need.type.replaceAll("-", " ")} brief.`, recommendation: requirement.reason, autoFixable: false });
  }

  const blockerCount = issues.filter((item) => item.severity === "blocker").length;
  const majorCount = issues.filter((item) => item.severity === "major").length;
  const minorCount = issues.filter((item) => item.severity === "minor").length;
  const currentPass = Math.min(3, (slide.qualityReview?.pass ?? 0) + 1);
  return {
    slideNumber,
    evidenceAuthority: "powerpoint-native",
    designImpact,
    issues,
    blockerCount,
    majorCount,
    minorCount,
    autoFixableCount: issues.filter((item) => item.autoFixable).length,
    verdict: blockerCount || majorCount ? "revise" : "ready",
    iteration: { currentPass, maxPasses: 3, remainingPasses: 3 - currentPass },
    checks: [
      "Compare the export pixels to the original source slide for exact message, labels, values, arrows, grouping, and causal relationships.",
      "Judge hierarchy, alignment, spacing rhythm, figure legibility, table consistency, and ORNL restraint at presentation scale.",
      "Record concrete visual issues only; after pass 3, hold unresolved ambiguity for human review instead of looping.",
    ],
  };
}
