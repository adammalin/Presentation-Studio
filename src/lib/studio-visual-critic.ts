import type { NativeMeasurementResult } from "./desktop";
import { nativeTextOverflows } from "./fresh-composition-qa";
import { calculateNativeCellClearances } from "./native-measurement";
import { analyzeStudioDesignImpact, type StudioDesignImpact } from "./studio-design-impact";
import { studioNodeOpticalBox } from "./studio-layout-constraints";
import { inferRepeatedImageSeries, studioGeneratedComponents } from "./studio-web-scene";
import type { StudioLayoutConstraint, StudioQualityIssue, StudioWebFrame, StudioWebNode, StudioWebScene } from "../types";

const PT = 12_700;
const ALIGNMENT_TOLERANCE = PT;

function inchesForPreflight(value: number): number {
  return Math.round(value * 72 * PT);
}

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

export interface StudioScenePreflight {
  issues: StudioQualityIssue[];
  blockerCount: number;
  majorCount: number;
  bySlide: Array<{ slideNumber: number; issues: StudioQualityIssue[] }>;
  ready: boolean;
}

function protectedBrandMark(node: StudioWebNode): boolean {
  return node.component?.role === "footer-logo" || /(?:^|\b)(?:ornl|doe|department of energy|oak ridge|wordmark|logo)(?:\b|$)/i.test(node.name);
}

function sourceLockedNodeIds(scene: StudioWebScene, slideNumber: number): Set<string> {
  const slide = scene.slides.find((candidate) => candidate.slideNumber === slideNumber);
  return new Set((slide?.figureTreatments ?? [])
    .filter((treatment) => ["preserve-as-unit", "preserve-and-frame"].includes(treatment.mode) && ["source-locked", "verified"].includes(treatment.verificationStatus))
    .flatMap((treatment) => treatment.nodeIds));
}

function overlapArea(left: StudioWebFrame, right: StudioWebFrame): number {
  return Math.max(0, Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x))
    * Math.max(0, Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y));
}

/**
 * Fast, deterministic fail-closed checks that run before Microsoft PowerPoint.
 * They prevent known-bad typography, header collisions, unsafe logo treatment,
 * and off-canvas geometry from entering an expensive native deck build.
 */
export function preflightStudioScene(scene: StudioWebScene, options: { protectedSlideNumbers?: readonly number[] } = {}): StudioScenePreflight {
  const protectedSlideNumbers = new Set(options.protectedSlideNumbers ?? []);
  const bySlide = scene.slides.map((slide) => {
    const issues: StudioQualityIssue[] = [];
    const add = (input: Omit<StudioQualityIssue, "id">) => issues.push({ id: `slide-${slide.slideNumber}-${input.category}-${issues.length + 1}`, ...input });
    const visible = slide.nodes.filter((node) => node.visible);
    const sourceLocked = sourceLockedNodeIds(scene, slide.slideNumber);
    const editable = visible.filter((node) => !sourceLocked.has(node.id));
    const footerRoles = new Set(["footer", "slide-number", "date", "logo"]);
    const protectedTemplateSlide = protectedSlideNumbers.has(slide.slideNumber);

    const repeatedImageSeries = inferRepeatedImageSeries(slide);
    if (!protectedTemplateSlide && repeatedImageSeries && slide.status === "designed" && slide.recipe !== "source") {
      if (slide.recipe !== "ornl-title-image-series") {
        add({ category: "figure", severity: "blocker", source: "scene", nodeIds: repeatedImageSeries.sourceNodeIds, message: `The source contains ${repeatedImageSeries.groups.length} image-heading-evidence groups, but ${slide.recipe} does not preserve that repeated relationship system.`, recommendation: "Use the ORNL image-series recipe or keep the faithful source slide; do not flatten peer groups into unrelated text and thumbnails.", autoFixable: true });
      } else {
        repeatedImageSeries.groups.forEach((group) => {
          const expected = [group.visual, group.heading, ...group.body];
          const groupIds = new Set(expected.map((node) => node.component?.groupId).filter((value): value is string => Boolean(value)));
          const roles = new Map(expected.map((node) => [node.id, node.component?.role]));
          const visibleExpected = expected.every((node) => node.visible);
          const ordinalMatches = expected.every((node) => node.component?.ordinal === group.ordinal);
          const rolesMatch = roles.get(group.visual.id) === "image-series-media"
            && roles.get(group.heading.id) === "image-series-heading"
            && group.body.every((node) => roles.get(node.id) === "image-series-body");
          if (!visibleExpected || groupIds.size !== 1 || !rolesMatch || !ordinalMatches) add({ category: "figure", severity: "blocker", source: "scene", nodeIds: expected.map((node) => node.id), message: `Image-series group ${group.ordinal + 1} lost its source image, heading, or evidence binding.`, recommendation: "Recompose the complete group as one image-series column with a shared relationship ID and source ordinal.", autoFixable: true });
          const visualCenter = group.visual.frame.x + group.visual.frame.width / 2;
          const aligned = [group.heading, ...group.body].every((node) => Math.abs(node.frame.x + node.frame.width / 2 - visualCenter) <= inchesForPreflight(.22));
          if (!aligned) add({ category: "alignment", severity: "blocker", source: "scene", nodeIds: expected.map((node) => node.id), message: `Image-series group ${group.ordinal + 1} no longer shares one visual column.`, recommendation: "Restore the common column center and equal series grid before native rendering.", autoFixable: true });
          if (group.visual.frame.width < inchesForPreflight(.82) || group.visual.frame.height < inchesForPreflight(.55)) add({ category: "legibility", severity: "blocker", source: "scene", nodeIds: [group.visual.id], message: `Image-series visual ${group.ordinal + 1} is too small to communicate its source evidence.`, recommendation: "Use the shared equal-column image field; never demote source visuals to decorative thumbnails.", autoFixable: true });
          const headingBottom = group.heading.frame.y + group.heading.frame.height;
          if (group.body.some((node) => node.frame.y < headingBottom)) add({ category: "spacing", severity: "blocker", source: "scene", nodeIds: [group.heading.id, ...group.body.map((node) => node.id)], message: `Image-series group ${group.ordinal + 1} places evidence copy inside its heading band.`, recommendation: "Keep the green heading band and body evidence in separate, aligned regions.", autoFixable: true });
        });
      }
    }

    for (const node of editable) {
      if (node.frame.x < 0 || node.frame.y < 0 || node.frame.x + node.frame.width > scene.slideSize.width || node.frame.y + node.frame.height > scene.slideSize.height) add({ category: "safe-region", severity: "blocker", source: "scene", nodeIds: [node.id], message: `${node.name} leaves the slide canvas.`, recommendation: "Recompose it inside the 16:9 canvas before building.", autoFixable: true });
      if (!protectedTemplateSlide && node.kind === "text" && node.text?.trim() && node.component?.role !== "footer-meta" && !footerRoles.has(node.role)) {
        const floor = node.role === "title" ? 24 : node.role === "caption" || node.role === "label" || node.component?.role === "eyebrow" || node.component?.role === "image-series-heading" ? 14 : 16;
        if (node.style.fontFamily.trim().toLowerCase() !== "aptos") add({ category: "brand", severity: "major", source: "scene", nodeIds: [node.id], message: `${node.name} uses ${node.style.fontFamily || "an unspecified font"} instead of Aptos.`, recommendation: "Use Aptos throughout ordinary ORNL presentation content.", autoFixable: true });
        if (node.style.fontSizePt < floor) add({ category: "legibility", severity: "major", source: "scene", nodeIds: [node.id], message: `${node.name} is ${node.style.fontSizePt} pt; the production floor for this role is ${floor} pt.`, recommendation: "Choose a roomier recipe, enlarge the region, or use an explicit continuation slide instead of miniaturizing type.", autoFixable: true });
      }
      if (protectedBrandMark(node) && node.kind === "image" && node.style.objectFit !== "contain") add({ category: "brand", severity: "blocker", source: "scene", nodeIds: [node.id], message: `${node.name} is protected brand artwork but is not using aspect-preserving contain behavior.`, recommendation: "Restore the approved asset with locked aspect ratio; never crop or stretch ORNL or DOE marks.", autoFixable: true });
    }

    for (const treatment of slide.figureTreatments) {
      const protectedIds = treatment.nodeIds.filter((id) => {
        const node = visible.find((candidate) => candidate.id === id);
        return node ? protectedBrandMark(node) : false;
      });
      if (protectedIds.length && treatment.lockAspectRatio === false) add({ category: "brand", severity: "blocker", source: "scene", nodeIds: protectedIds, message: "A figure treatment can distort protected ORNL or DOE artwork.", recommendation: "Lock the complete mark group aspect ratio and use contain behavior.", autoFixable: true });
    }

    const textNodes = editable.filter((node) => node.kind === "text" && node.text?.trim() && !footerRoles.has(node.role));
    if (!protectedTemplateSlide) for (let leftIndex = 0; leftIndex < textNodes.length; leftIndex += 1) for (let rightIndex = leftIndex + 1; rightIndex < textNodes.length; rightIndex += 1) {
      const left = textNodes[leftIndex];
      const right = textNodes[rightIndex];
      const overlap = overlapArea(left.frame, right.frame);
      const smaller = Math.max(1, Math.min(left.frame.width * left.frame.height, right.frame.width * right.frame.height));
      if (overlap / smaller > .08) add({ category: "spacing", severity: "blocker", source: "scene", nodeIds: [left.id, right.id], message: `${left.name} and ${right.name} occupy overlapping text regions.`, recommendation: "Recompose both regions on the shared grid before native rendering.", autoFixable: true });
    }

    if (!["source", "template-layout"].includes(slide.recipe)) {
      const rule = studioGeneratedComponents(slide).find((component) => component.id.includes("title-rule"));
      if (rule) for (const node of textNodes.filter((candidate) => candidate.role !== "title")) if (overlapArea(rule.frame, node.frame) > 0) add({ category: "spacing", severity: "blocker", source: "scene", nodeIds: [node.id], message: `The ORNL title accent overlaps ${node.name}.`, recommendation: "Place body content below the complete title and accent system.", autoFixable: true });
    }

    return { slideNumber: slide.slideNumber, issues };
  });
  const issues = bySlide.flatMap((slide) => slide.issues);
  const blockerCount = issues.filter((item) => item.severity === "blocker").length;
  const majorCount = issues.filter((item) => item.severity === "major").length;
  return { issues, blockerCount, majorCount, bySlide: bySlide.filter((slide) => slide.issues.length > 0), ready: blockerCount === 0 && majorCount === 0 };
}

export function nativeStudioProductionIssues(measurement: NativeMeasurementResult): Array<{ slideNumber: number; message: string }> {
  if (measurement.status !== "ready" || measurement.authority !== "powerpoint-native") return [{ slideNumber: 0, message: measurement.reason ?? "PowerPoint-native measurement is unavailable." }];
  const issues = nativeTextOverflows(measurement).map((overflow) => ({ slideNumber: overflow.slideNumber, message: `${overflow.name || "Text"} renders outside its frame (${overflow.edges.join(", ")}).` }));
  for (const slide of measurement.slides) for (const shape of slide.shapes) for (const cell of shape.table?.cells ?? []) {
    const bounds = cell.boundsPt;
    const text = cell.renderedTextBoundsPt;
    const margins = cell.marginsPt;
    if (bounds && text && margins && (text.width > bounds.width - margins.left - margins.right + .5 || text.height > bounds.height - margins.top - margins.bottom + .5)) issues.push({ slideNumber: slide.number, message: `${shape.name ?? "Table"} cell r${cell.row}c${cell.column} overflows its editable PowerPoint cell.` });
    const clearance = calculateNativeCellClearances(cell);
    if (clearance && (Math.min(clearance.left, clearance.right) < 1 || Math.min(clearance.top, clearance.bottom) < 1)) issues.push({ slideNumber: slide.number, message: `${shape.name ?? "Table"} cell r${cell.row}c${cell.column} has less than 1 pt native text clearance.` });
  }
  return issues;
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
    const supportingIcon = treatment.id.startsWith("studio-auto-metric-icon-") || /(?:source )?metric icon/i.test(treatment.intentSummary);
    const minimumWidth = supportingIcon ? 28 * PT : 108 * PT;
    const minimumHeight = supportingIcon ? 28 * PT : 72 * PT;
    if (treatment.mode === "redraw-candidate" && treatment.verificationStatus !== "verified") add({ category: "figure", severity: "blocker", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} is marked for redraw but its information and relationships are not verified.`, recommendation: "Keep the original technical object source-locked until a content owner verifies the complete inventory and relationship map.", autoFixable: false });
    if (frame && (frame.width < minimumWidth || frame.height < minimumHeight)) add({ category: "figure", severity: "major", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} is framed too small for reliable review.`, recommendation: supportingIcon ? "Keep the complete source icon at least 28 points in both dimensions and paired with its metric." : "Assign the figure a larger evidence region or crop only after its meaning-bearing area is verified.", autoFixable: false });
    if (!supportingIcon && treatment.crop && (1 - treatment.crop.left - treatment.crop.right) * (1 - treatment.crop.top - treatment.crop.bottom) < .25) add({ category: "figure", severity: "major", source: "scene", nodeIds: treatment.nodeIds, message: `${treatment.intentSummary} retains less than 25% of the source figure area.`, recommendation: "Verify that no labels, values, arrows, or context were removed before accepting the crop.", autoFixable: false });
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
