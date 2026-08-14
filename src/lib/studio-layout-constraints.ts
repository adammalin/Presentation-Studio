import type { NativeMeasurementResult } from "./desktop";
import type { StudioConstraintKind, StudioConstraintMode, StudioFigureTreatment, StudioLayoutConstraint, StudioWebFrame, StudioWebNode, StudioWebScene, StudioWebSlide } from "../types";
import { STUDIO_WEB_SCENE_VERSION } from "../types";
import { defaultStudioDeckRhythm } from "./studio-web-scene";

const EMU_PER_POINT = 12_700;
const TOLERANCE_EMU = EMU_PER_POINT * .5;

export interface StudioConstraintRequest {
  kind: StudioConstraintKind;
  mode: StudioConstraintMode;
  nodeIds: string[];
  groups?: string[][];
  anchorNodeId?: string;
  gridPt?: number;
  rationale: string;
  author?: "human" | "ai";
}

export interface StudioConstraintResult {
  scene: StudioWebScene;
  slideNumber: number;
  changedNodeIds: string[];
  constraints: StudioLayoutConstraint[];
  evidenceAuthority: "scene-estimate" | "powerpoint-native";
  diagnostics: string[];
}

type Box = { x: number; y: number; width: number; height: number };
type Unit = { ids: string[]; nodes: StudioWebNode[]; frame: Box; optical: Box; figureTreatmentId?: string };

function median(values: number[]) {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function union(boxes: Box[]): Box {
  const x = Math.min(...boxes.map((box) => box.x));
  const y = Math.min(...boxes.map((box) => box.y));
  const right = Math.max(...boxes.map((box) => box.x + box.width));
  const bottom = Math.max(...boxes.map((box) => box.y + box.height));
  return { x, y, width: right - x, height: bottom - y };
}

function intersects(left: Box, right: Box) {
  return left.x < right.x + right.width - TOLERANCE_EMU
    && left.x + left.width > right.x + TOLERANCE_EMU
    && left.y < right.y + right.height - TOLERANCE_EMU
    && left.y + left.height > right.y + TOLERANCE_EMU;
}

function nativeShapeFor(node: StudioWebNode, measurement?: NativeMeasurementResult) {
  if (measurement?.status !== "ready" || measurement.authority !== "powerpoint-native") return undefined;
  const shapes = measurement.slides.flatMap((slide) => slide.shapes);
  return shapes.find((shape) => shape.name === node.id || shape.name?.endsWith(` · ${node.id}`));
}

export function studioNodeOpticalBox(node: StudioWebNode, measurement?: NativeMeasurementResult): { box: Box; authority: "scene-estimate" | "powerpoint-native" } {
  const shape = nativeShapeFor(node, measurement);
  const native = node.kind === "text" ? shape?.text?.renderedBoundsPt : shape?.boundsPt;
  if (native && shape?.boundsPt) {
    const measuredShape = { x: shape.boundsPt.left * EMU_PER_POINT, y: shape.boundsPt.top * EMU_PER_POINT };
    // A bounded constraint transaction only translates nodes. Carry the
    // measured PowerPoint optical inset with the node after each sequential
    // translation rather than continuing to use the old absolute pixels.
    const dx = node.frame.x - measuredShape.x;
    const dy = node.frame.y - measuredShape.y;
    return {
      box: { x: Math.round(native.left * EMU_PER_POINT + dx), y: Math.round(native.top * EMU_PER_POINT + dy), width: Math.round(native.width * EMU_PER_POINT), height: Math.round(native.height * EMU_PER_POINT) },
      authority: "powerpoint-native",
    };
  }
  const inset = node.opticalInsets ?? { left: 0, top: 0, right: 0, bottom: 0 };
  return {
    box: {
      x: node.frame.x + inset.left,
      y: node.frame.y + inset.top,
      width: Math.max(1, node.frame.width - inset.left - inset.right),
      height: Math.max(1, node.frame.height - inset.top - inset.bottom),
    },
    authority: "scene-estimate",
  };
}

function unitsFor(slide: StudioWebSlide, request: StudioConstraintRequest, measurement?: NativeMeasurementResult): Unit[] {
  const groups = request.groups?.length ? request.groups : request.nodeIds.map((id) => [id]);
  const flattened = groups.flat();
  if (new Set(flattened).size !== flattened.length) throw new Error("Each Studio node may belong to only one constraint group.");
  if (flattened.length !== request.nodeIds.length || request.nodeIds.some((id) => !flattened.includes(id))) throw new Error("Constraint nodeIds must exactly match the supplied groups.");
  const byId = new Map(slide.nodes.map((node) => [node.id, node]));
  return groups.map((ids) => {
    const members = ids.map((id) => byId.get(id));
    if (members.some((node) => !node)) throw new Error("A constrained Studio node is not present in the current slide revision.");
    const present = members as StudioWebNode[];
    const treatment = slide.figureTreatments.find((candidate) => candidate.nodeIds.length === ids.length && candidate.nodeIds.every((id) => ids.includes(id)));
    if (present.some((node) => !node.visible)) throw new Error("Hidden Studio nodes cannot enter an automatic layout constraint.");
    if (!treatment && present.some((node) => node.locked)) throw new Error("Locked Studio nodes may move only as one complete first-class figure treatment.");
    const figureFrame = treatment?.groupFrame;
    return figureFrame
      ? { ids, nodes: present, frame: figureFrame, optical: figureFrame, figureTreatmentId: treatment.id }
      : { ids, nodes: present, frame: union(present.map((node) => node.frame)), optical: union(present.map((node) => studioNodeOpticalBox(node, measurement).box)), figureTreatmentId: treatment?.id };
  });
}

function unitValue(unit: Unit, mode: StudioConstraintMode) {
  const box = mode === "optical-left" || mode === "optical-top" ? unit.optical : unit.frame;
  if (mode === "left" || mode === "optical-left") return box.x;
  if (mode === "center") return box.x + box.width / 2;
  if (mode === "right") return box.x + box.width;
  if (mode === "top" || mode === "optical-top") return box.y;
  if (mode === "middle") return box.y + box.height / 2;
  return box.y + box.height;
}

function translateUnit(unit: Unit, dx: number, dy: number, frames: Map<string, StudioWebFrame>, figureFrames: Map<string, StudioWebFrame>) {
  if (unit.figureTreatmentId) {
    figureFrames.set(unit.figureTreatmentId, { ...unit.frame, x: Math.round(unit.frame.x + dx), y: Math.round(unit.frame.y + dy), rotation: 0 });
    for (const node of unit.nodes) frames.set(node.id, { ...node.frame, x: Math.round(node.frame.x + dx), y: Math.round(node.frame.y + dy) });
    return;
  }
  for (const node of unit.nodes) frames.set(node.id, { ...node.frame, x: Math.round(node.frame.x + dx), y: Math.round(node.frame.y + dy) });
}

function applyOne(scene: StudioWebScene, slideNumber: number, request: StudioConstraintRequest, measurement?: NativeMeasurementResult) {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error("The constrained Studio slide is not present in the current scene revision.");
  if (!request.nodeIds.length || request.nodeIds.length > 60 || new Set(request.nodeIds).size !== request.nodeIds.length) throw new Error("A Studio constraint requires 1–60 unique node IDs.");
  if (!request.rationale.trim()) throw new Error("A Studio constraint requires a design rationale.");
  const units = unitsFor(slide, request, measurement);
  const frames = new Map<string, StudioWebFrame>();
  const figureFrames = new Map<string, StudioWebFrame>();
  const rhythm = scene.rhythm ?? defaultStudioDeckRhythm();

  if (request.kind === "align") {
    if (units.length < 2) throw new Error("Alignment requires at least two nodes or groups.");
    if (!["left", "optical-left", "center", "right", "top", "optical-top", "middle", "bottom"].includes(request.mode)) throw new Error("Choose a supported Studio alignment mode.");
    const anchorUnit = request.anchorNodeId ? units.find((unit) => unit.ids.includes(request.anchorNodeId!)) : undefined;
    if (request.anchorNodeId && !anchorUnit) throw new Error("The alignment anchor must belong to the constrained nodes.");
    const target = anchorUnit ? unitValue(anchorUnit, request.mode) : median(units.map((unit) => unitValue(unit, request.mode)));
    for (const unit of units) {
      const delta = target - unitValue(unit, request.mode);
      translateUnit(unit, ["left", "optical-left", "center", "right"].includes(request.mode) ? delta : 0, ["top", "optical-top", "middle", "bottom"].includes(request.mode) ? delta : 0, frames, figureFrames);
    }
  } else if (request.kind === "distribute") {
    if (units.length < 3) throw new Error("Equal-gap distribution requires at least three nodes or groups.");
    if (!["horizontal-equal-gap", "vertical-equal-gap"].includes(request.mode)) throw new Error("Choose horizontal or vertical equal-gap distribution.");
    const horizontal = request.mode === "horizontal-equal-gap";
    const ordered = [...units].sort((left, right) => horizontal ? left.frame.x - right.frame.x : left.frame.y - right.frame.y);
    const first = ordered[0].frame;
    const last = ordered.at(-1)!.frame;
    const span = horizontal ? last.x + last.width - first.x : last.y + last.height - first.y;
    const occupied = ordered.reduce((sum, unit) => sum + (horizontal ? unit.frame.width : unit.frame.height), 0);
    const gap = (span - occupied) / (ordered.length - 1);
    if (gap < 0) throw new Error("The selected Studio groups overlap inside their current span and cannot be distributed without resizing or recomposition.");
    let cursor = horizontal ? first.x : first.y;
    for (const unit of ordered) {
      const origin = horizontal ? unit.frame.x : unit.frame.y;
      translateUnit(unit, horizontal ? cursor - origin : 0, horizontal ? 0 : cursor - origin, frames, figureFrames);
      cursor += (horizontal ? unit.frame.width : unit.frame.height) + gap;
    }
  } else if (request.kind === "snap-to-grid") {
    const grid = Math.round(Math.max(1, Math.min(72, request.gridPt ?? rhythm.gridPt)) * EMU_PER_POINT);
    if (!['left', 'top', 'both'].includes(request.mode)) throw new Error("Grid snapping supports left, top, or both axes.");
    for (const unit of units) {
      const dx = request.mode === "top" ? 0 : Math.round(unit.frame.x / grid) * grid - unit.frame.x;
      const dy = request.mode === "left" ? 0 : Math.round(unit.frame.y / grid) * grid - unit.frame.y;
      translateUnit(unit, dx, dy, frames, figureFrames);
    }
  } else if (request.kind === "fit-safe-region") {
    const safe = Math.round(rhythm.safeMarginPt * EMU_PER_POINT);
    const group = union(units.map((unit) => unit.frame));
    if (group.width > scene.slideSize.width - safe * 2 || group.height > scene.slideSize.height - safe * 2) throw new Error("The selected Studio group is larger than the safe region; choose another recipe or resize the whole component intentionally.");
    const dx = group.x < safe ? safe - group.x : group.x + group.width > scene.slideSize.width - safe ? scene.slideSize.width - safe - group.x - group.width : 0;
    const dy = group.y < safe ? safe - group.y : group.y + group.height > scene.slideSize.height - safe ? scene.slideSize.height - safe - group.y - group.height : 0;
    for (const unit of units) translateUnit(unit, dx, dy, frames, figureFrames);
  }

  const beforeById = new Map(slide.nodes.map((node) => [node.id, node.frame]));
  const afterNodes = slide.nodes.map((node) => frames.has(node.id) ? { ...node, frame: frames.get(node.id)! } : node);
  const afterTreatments = slide.figureTreatments.map((treatment): StudioFigureTreatment => figureFrames.has(treatment.id) ? { ...treatment, groupFrame: figureFrames.get(treatment.id)! } : treatment);
  const selected = new Set(request.nodeIds);
  const afterById = new Map(afterNodes.map((node) => [node.id, node]));
  const afterUnitFrame = (unit: Unit) => unit.figureTreatmentId ? figureFrames.get(unit.figureTreatmentId) ?? unit.frame : union(unit.ids.map((id) => afterById.get(id)!.frame));
  const obstacles = afterNodes.filter((node) => node.visible && !selected.has(node.id));
  for (const unit of units) {
    const after = afterUnitFrame(unit);
    if (after.x < 0 || after.y < 0 || after.x + after.width > scene.slideSize.width || after.y + after.height > scene.slideSize.height) throw new Error(`${unit.figureTreatmentId ?? unit.ids.join(", ")} would leave the Studio slide canvas.`);
    for (const obstacle of obstacles) if (!intersects(unit.frame, obstacle.frame) && intersects(after, obstacle.frame)) throw new Error(`${unit.figureTreatmentId ?? unit.ids.join(", ")} would create a new overlap with ${obstacle.id}; recompose the group or choose a different constraint anchor.`);
  }
  for (let left = 0; left < units.length; left += 1) for (let right = left + 1; right < units.length; right += 1) {
    if (!intersects(units[left].frame, units[right].frame) && intersects(afterUnitFrame(units[left]), afterUnitFrame(units[right]))) throw new Error("The requested constraint would create a new overlap between selected Studio groups.");
  }
  const changedNodeIds = [...selected].filter((id) => {
    const unit = units.find((candidate) => candidate.ids.includes(id));
    return Boolean(unit?.figureTreatmentId ? figureFrames.has(unit.figureTreatmentId) && JSON.stringify(unit.frame) !== JSON.stringify(figureFrames.get(unit.figureTreatmentId)) : JSON.stringify(beforeById.get(id)) !== JSON.stringify(afterById.get(id)?.frame));
  });
  const evidenceAuthority = units.every((unit) => !unit.figureTreatmentId && unit.nodes.every((node) => studioNodeOpticalBox(afterById.get(node.id)!, measurement).authority === "powerpoint-native")) ? "powerpoint-native" : "scene-estimate";
  const now = new Date().toISOString();
  const record: StudioLayoutConstraint = {
    id: crypto.randomUUID(), kind: request.kind, mode: request.mode, nodeIds: [...request.nodeIds], groups: request.groups?.map((group) => [...group]), anchorNodeId: request.anchorNodeId, gridPt: request.gridPt,
    rationale: request.rationale.trim().slice(0, 1_000), author: request.author ?? "ai", evidenceAuthority, appliedAt: now,
  };
  const nextSlide = { ...slide, status: "designed" as const, nodes: afterNodes, figureTreatments: afterTreatments, constraints: [...(slide.constraints ?? []), record].slice(-80), updatedAt: now, designRationale: `${slide.designRationale} ${record.rationale}`.trim().slice(0, 1_000) };
  const nextScene = { ...scene, rhythm, revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`, slides: scene.slides.map((item) => item.slideNumber === slideNumber ? nextSlide : item) };
  return { scene: nextScene, changedNodeIds, record, evidenceAuthority };
}

export function applyStudioLayoutConstraints(scene: StudioWebScene, slideNumber: number, requests: StudioConstraintRequest[], measurement?: NativeMeasurementResult): StudioConstraintResult {
  if (!requests.length || requests.length > 20) throw new Error("Apply between 1 and 20 Studio layout constraints in one bounded transaction.");
  let next = scene;
  const changed = new Set<string>();
  const records: StudioLayoutConstraint[] = [];
  const diagnostics: string[] = [];
  let authority: "scene-estimate" | "powerpoint-native" = "powerpoint-native";
  for (const request of requests) {
    const result = applyOne(next, slideNumber, request, measurement);
    next = result.scene;
    result.changedNodeIds.forEach((id) => changed.add(id));
    records.push(result.record);
    if (result.evidenceAuthority !== "powerpoint-native") authority = "scene-estimate";
    if (!result.changedNodeIds.length) diagnostics.push(`${request.kind} ${request.mode} was already satisfied.`);
  }
  if (authority !== "powerpoint-native") diagnostics.push("Final optical acceptance still requires a matching PowerPoint-native build and measurement pass.");
  return { scene: next, slideNumber, changedNodeIds: [...changed], constraints: records, evidenceAuthority: authority, diagnostics };
}
