import type { DeckJob, GeometryEditCommand, PresentationSceneObject } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import type { NativeMeasurementPacket } from "./native-measurement";

const EMU_PER_POINT = 12_700;
const SAFE_MARGIN_PT = PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt;

export type AlignmentMode = "left" | "optical-left" | "center" | "right" | "top" | "middle" | "bottom";
export type DistributionMode = "horizontal-equal-gap" | "vertical-equal-gap";
export type GroupLayoutMode = "horizontal-stack" | "vertical-stack";
export type GroupLayoutAlignment = "start" | "center" | "end";
export type GroupHierarchyRole = "primary" | "supporting" | "caption";

export interface LayoutSolverResult {
  status: "solved" | "infeasible";
  operation: "align" | "distribute" | "fit-safe-region" | "layout-group" | "fit-scene-to-layout";
  commands: GeometryEditCommand[];
  objective: { movementCostPt: number; maximumMovementPt: number };
  diagnostics: string[];
}

function emuToPt(value: number) { return value / EMU_PER_POINT; }
function ptToEmu(value: number) { return Math.round(value * EMU_PER_POINT); }
function median(values: number[]) {
  const sorted = [...values].sort((left, right) => left - right);
  const midpoint = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[midpoint] : (sorted[midpoint - 1] + sorted[midpoint]) / 2;
}

function sceneObjects(deck: DeckJob, slideNumber: number, objectIds: string[]) {
  if (!deck.scene || !deck.audit) throw new Error("A current scene and audit are required before solving layout geometry.");
  const objects = objectIds.map((id) => deck.scene!.objects.find((object) => object.id === id)).filter((object): object is PresentationSceneObject => Boolean(object));
  if (objects.length !== objectIds.length) throw new Error("At least one requested object is not present in the current scene revision.");
  if (objects.some((object) => object.slideNumber !== slideNumber)) throw new Error("All requested objects must belong to the same slide.");
  if (objects.some((object) => !object.operations.move || object.protected)) throw new Error("At least one requested object is protected or cannot be moved safely.");
  return objects;
}

function objectBoxPt(object: PresentationSceneObject) {
  return { left: emuToPt(object.geometry.x), top: emuToPt(object.geometry.y), width: emuToPt(object.geometry.width), height: emuToPt(object.geometry.height) };
}

function overlap(left: { left: number; top: number; width: number; height: number }, right: { left: number; top: number; width: number; height: number }) {
  return left.left < right.left + right.width - 0.5 && left.left + left.width > right.left + 0.5 && left.top < right.top + right.height - 0.5 && left.top + left.height > right.top + 0.5;
}

function unionBox(objects: PresentationSceneObject[]) {
  const boxes = objects.map(objectBoxPt);
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

type LayoutTarget = { left: number; top: number; width?: number; height?: number };

function buildCommands(deck: DeckJob, slideNumber: number, objects: PresentationSceneObject[], targets: Map<string, LayoutTarget>, rationale: string, groupByObjectId = new Map<string, number>(), ignoreCollisionObjectIds = new Set<string>()): LayoutSolverResult {
  const slideWidthPt = deck.audit!.slideSize.width / EMU_PER_POINT;
  const slideHeightPt = deck.audit!.slideSize.height / EMU_PER_POINT;
  const diagnostics: string[] = [];
  const proposedBoxes = new Map<string, { left: number; top: number; width: number; height: number }>();
  let movementCostPt = 0;
  let maximumMovementPt = 0;
  for (const object of objects) {
    const source = objectBoxPt(object);
    const target = targets.get(object.id) ?? { left: source.left, top: source.top };
    const box = { left: target.left, top: target.top, width: target.width ?? source.width, height: target.height ?? source.height };
    if ((Math.abs(box.width - source.width) >= .1 || Math.abs(box.height - source.height) >= .1) && !object.operations.resize) diagnostics.push(`${object.id} cannot be resized safely.`);
    if (box.left < SAFE_MARGIN_PT || box.top < SAFE_MARGIN_PT || box.left + box.width > slideWidthPt - SAFE_MARGIN_PT || box.top + box.height > slideHeightPt - SAFE_MARGIN_PT) {
      diagnostics.push(`${object.id} would enter the 0.25-inch safe region.`);
    }
    proposedBoxes.set(object.id, box);
    const movement = Math.abs(target.left - source.left) + Math.abs(target.top - source.top) + Math.abs(box.width - source.width) + Math.abs(box.height - source.height);
    movementCostPt += movement;
    maximumMovementPt = Math.max(maximumMovementPt, movement);
  }
  const selectedIds = new Set(objects.map((object) => object.id));
  const allSlideObjects = deck.scene!.objects.filter((object) => object.slideNumber === slideNumber && !selectedIds.has(object.id) && !ignoreCollisionObjectIds.has(object.id));
  for (const object of objects) {
    const target = proposedBoxes.get(object.id)!;
    const source = objectBoxPt(object);
    const collisions = allSlideObjects.filter((candidate) => overlap(target, objectBoxPt(candidate)) && !overlap(source, objectBoxPt(candidate))).map((candidate) => candidate.id);
    if (collisions.length) diagnostics.push(`${object.id} would overlap ${collisions.join(", ")}.`);
  }
  for (let leftIndex = 0; leftIndex < objects.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < objects.length; rightIndex += 1) {
      const leftObject = objects[leftIndex];
      const rightObject = objects[rightIndex];
      if (groupByObjectId.get(leftObject.id) === groupByObjectId.get(rightObject.id) && groupByObjectId.has(leftObject.id)) continue;
      const sourceOverlaps = overlap(objectBoxPt(leftObject), objectBoxPt(rightObject));
      const targetOverlaps = overlap(proposedBoxes.get(leftObject.id)!, proposedBoxes.get(rightObject.id)!);
      if (targetOverlaps && !sourceOverlaps) diagnostics.push(`${leftObject.id} would overlap ${rightObject.id}.`);
    }
  }
  if (diagnostics.length) return { status: "infeasible", operation: "align", commands: [], objective: { movementCostPt, maximumMovementPt }, diagnostics };
  const commands = objects.flatMap((object) => {
    const source = objectBoxPt(object);
    const target = proposedBoxes.get(object.id)!;
    const moved = Math.abs(target.left - source.left) >= .1 || Math.abs(target.top - source.top) >= .1;
    const resized = Math.abs(target.width - source.width) >= .1 || Math.abs(target.height - source.height) >= .1;
    if (!moved && !resized) return [];
    return [{
      id: crypto.randomUUID(),
      slideNumber,
      objectId: object.id,
      shapeId: object.shapeId,
      sourceElement: object.sourceElement,
      objectKind: object.kind,
      operation: moved && resized ? "move-and-resize" as const : moved ? "move" as const : "resize" as const,
      source: { x: object.geometry.x, y: object.geometry.y, width: object.geometry.width, height: object.geometry.height },
      target: { x: ptToEmu(target.left), y: ptToEmu(target.top), width: ptToEmu(target.width), height: ptToEmu(target.height) },
      rationale,
      author: "ai" as const,
      constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false },
      validation: { safeAreaStatus: "inside" as const, overlapObjectIds: [], warnings: [] },
    }];
  });
  return { status: "solved", operation: "align", commands, objective: { movementCostPt, maximumMovementPt }, diagnostics: [] };
}

export function solveAlignment(input: { deck: DeckJob; measurement: NativeMeasurementPacket; slideNumber: number; objectIds: string[]; mode: AlignmentMode; rationale: string; anchorObjectId?: string }): LayoutSolverResult {
  const { deck, measurement, slideNumber, objectIds, mode } = input;
  if (objectIds.length < 2) throw new Error("Alignment requires at least two source-bound objects.");
  if (!["left", "optical-left", "center", "right", "top", "middle", "bottom"].includes(mode)) throw new Error("Choose a supported semantic alignment mode.");
  const objects = sceneObjects(deck, slideNumber, objectIds);
  const boxes = new Map(objects.map((object) => [object.id, objectBoxPt(object)]));
  const measured = new Map(measurement.objects.filter((object) => object.slideNumber === slideNumber).map((object) => [object.objectId, object]));
  const valueFor = (object: PresentationSceneObject) => {
    const box = boxes.get(object.id)!;
    if (mode === "optical-left") return measured.get(object.id)?.text?.renderedBoundsPt?.left ?? box.left;
    if (mode === "left") return box.left;
    if (mode === "center") return box.left + box.width / 2;
    if (mode === "right") return box.left + box.width;
    if (mode === "top") return box.top;
    if (mode === "middle") return box.top + box.height / 2;
    return box.top + box.height;
  };
  const anchor = input.anchorObjectId ? objects.find((object) => object.id === input.anchorObjectId) : undefined;
  if (input.anchorObjectId && !anchor) throw new Error("The alignment anchor must be one of the selected objects.");
  const targetValue = anchor ? valueFor(anchor) : median(objects.map(valueFor));
  const targets = new Map<string, { left: number; top: number }>();
  for (const object of objects) {
    const box = boxes.get(object.id)!;
    const delta = targetValue - valueFor(object);
    targets.set(object.id, ["left", "optical-left", "center", "right"].includes(mode) ? { left: box.left + delta, top: box.top } : { left: box.left, top: box.top + delta });
  }
  const result = buildCommands(deck, slideNumber, objects, targets, input.rationale);
  result.operation = "align";
  if (measurement.authority !== "powerpoint-native" && mode === "optical-left") result.diagnostics.push("Optical alignment used OOXML fallback; acquire native PowerPoint measurements before final acceptance.");
  return result;
}

export function solveDistribution(input: { deck: DeckJob; slideNumber: number; objectIds: string[]; groups?: string[][]; mode: DistributionMode; rationale: string }): LayoutSolverResult {
  const { deck, slideNumber, objectIds, mode } = input;
  const groups = input.groups?.length ? input.groups : objectIds.map((id) => [id]);
  if (groups.length < 3) throw new Error("Equal-gap distribution requires at least three source-bound objects or groups.");
  if (!["horizontal-equal-gap", "vertical-equal-gap"].includes(mode)) throw new Error("Choose a supported equal-gap distribution mode.");
  if (groups.some((group) => group.length === 0)) throw new Error("Distribution groups cannot be empty.");
  const flattenedIds = groups.flat();
  if (new Set(flattenedIds).size !== flattenedIds.length) throw new Error("Each object may belong to only one distribution group.");
  if (objectIds.length && (objectIds.length !== flattenedIds.length || objectIds.some((id) => !flattenedIds.includes(id)))) throw new Error("objectIds must match the objects supplied in distribution groups.");
  const objects = sceneObjects(deck, slideNumber, flattenedIds);
  const objectById = new Map(objects.map((object) => [object.id, object]));
  const groupedObjects = groups.map((group) => group.map((id) => objectById.get(id)!));
  const groupByObjectId = new Map(groups.flatMap((group, index) => group.map((id) => [id, index] as const)));
  const horizontal = mode === "horizontal-equal-gap";
  const sorted = [...groupedObjects].sort((left, right) => horizontal ? unionBox(left).left - unionBox(right).left : unionBox(left).top - unionBox(right).top);
  const first = unionBox(sorted[0]);
  const last = unionBox(sorted[sorted.length - 1]);
  const totalSize = sorted.reduce((sum, group) => sum + (horizontal ? unionBox(group).width : unionBox(group).height), 0);
  const span = horizontal ? last.left + last.width - first.left : last.top + last.height - first.top;
  const gap = (span - totalSize) / (sorted.length - 1);
  if (gap < 0) return { status: "infeasible", operation: "distribute", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: ["The selected objects do not fit within their current outer span without overlap. Increase the region or resize content before distributing."] };
  const targets = new Map<string, { left: number; top: number }>();
  let cursor = horizontal ? first.left : first.top;
  for (const group of sorted) {
    const groupBox = unionBox(group);
    const delta = cursor - (horizontal ? groupBox.left : groupBox.top);
    for (const object of group) {
      const box = objectBoxPt(object);
      targets.set(object.id, horizontal ? { left: box.left + delta, top: box.top } : { left: box.left, top: box.top + delta });
    }
    cursor += (horizontal ? groupBox.width : groupBox.height) + gap;
  }
  const result = buildCommands(deck, slideNumber, objects, targets, input.rationale, groupByObjectId);
  result.operation = "distribute";
  return result;
}

export function solveSafeRegion(input: { deck: DeckJob; slideNumber: number; objectIds: string[]; rationale: string }): LayoutSolverResult {
  const { deck, slideNumber, objectIds } = input;
  if (!objectIds.length) throw new Error("Safe-region fitting requires at least one source-bound object.");
  const objects = sceneObjects(deck, slideNumber, objectIds);
  const slideWidthPt = deck.audit!.slideSize.width / EMU_PER_POINT;
  const slideHeightPt = deck.audit!.slideSize.height / EMU_PER_POINT;
  const bounds = unionBox(objects);
  if (bounds.width > slideWidthPt - SAFE_MARGIN_PT * 2 || bounds.height > slideHeightPt - SAFE_MARGIN_PT * 2) {
    return { status: "infeasible", operation: "fit-safe-region", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: ["The selected group is larger than the safe region. Resize or recompose it before fitting."] };
  }
  let deltaX = bounds.left < SAFE_MARGIN_PT ? SAFE_MARGIN_PT - bounds.left : 0;
  let deltaY = bounds.top < SAFE_MARGIN_PT ? SAFE_MARGIN_PT - bounds.top : 0;
  if (bounds.left + bounds.width + deltaX > slideWidthPt - SAFE_MARGIN_PT) deltaX = slideWidthPt - SAFE_MARGIN_PT - bounds.left - bounds.width;
  if (bounds.top + bounds.height + deltaY > slideHeightPt - SAFE_MARGIN_PT) deltaY = slideHeightPt - SAFE_MARGIN_PT - bounds.top - bounds.height;
  if (Math.abs(deltaX) < .1 && Math.abs(deltaY) < .1) return { status: "solved", operation: "fit-safe-region", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: ["The selected group already fits inside the safe region."] };
  const targets = new Map(objects.map((object) => {
    const box = objectBoxPt(object);
    return [object.id, { left: box.left + deltaX, top: box.top + deltaY }] as const;
  }));
  const groupByObjectId = new Map(objects.map((object) => [object.id, 0] as const));
  const result = buildCommands(deck, slideNumber, objects, targets, input.rationale, groupByObjectId);
  result.operation = "fit-safe-region";
  return result;
}

export function solveGroupLayout(input: {
  deck: DeckJob;
  slideNumber: number;
  groups: string[][];
  regionPt: { left: number; top: number; width: number; height: number };
  mode: GroupLayoutMode;
  alignment: GroupLayoutAlignment;
  groupRoles?: GroupHierarchyRole[];
  preferredGapPt?: number;
  allowResponsiveScale?: boolean;
  minimumScale?: number;
  rationale: string;
  coordinatedObjectIds?: string[];
}): LayoutSolverResult {
  const { deck, slideNumber, regionPt, mode, alignment } = input;
  if (!input.groups.length || input.groups.length > 10 || input.groups.some((group) => !group.length)) throw new Error("Group layout requires between 1 and 10 non-empty visual groups.");
  const objectIds = input.groups.flat();
  if (objectIds.length > 20) throw new Error("Group layout supports at most 20 source-bound objects in one transaction.");
  if (new Set(objectIds).size !== objectIds.length) throw new Error("Each object may belong to only one visual group.");
  if (![regionPt.left, regionPt.top, regionPt.width, regionPt.height].every(Number.isFinite) || regionPt.width <= 0 || regionPt.height <= 0) throw new Error("The approved layout region must have finite positive bounds.");
  const objects = sceneObjects(deck, slideNumber, objectIds);
  const objectById = new Map(objects.map((object) => [object.id, object]));
  const groups = input.groups.map((group) => group.map((id) => objectById.get(id)!));
  const groupRoles = input.groupRoles ?? groups.map(() => "supporting" as const);
  if (groupRoles.length !== groups.length) throw new Error("groupRoles must provide one semantic role for every visual group.");
  if (groupRoles.some((role) => !["primary", "supporting", "caption"].includes(role))) throw new Error("Each visual group role must be primary, supporting, or caption.");
  if (groupRoles.filter((role) => role === "primary").length > 1) throw new Error("One approved region may contain at most one primary visual group.");
  const sourceBoxes = groups.map(unionBox);
  const horizontal = mode === "horizontal-stack";
  const regionPrimary = horizontal ? regionPt.width : regionPt.height;
  const regionCross = horizontal ? regionPt.height : regionPt.width;
  const spacing = PRESENTATION_DESIGN_STANDARD.componentSystem.spacing;
  const preferredGapPt = Math.max(0, Math.min(72, input.preferredGapPt ?? spacing.normalPt));
  const gapCount = Math.max(0, groups.length - 1);
  const preferredGaps = Array.from({ length: gapCount }, (_value, index) => {
    const pair = [groupRoles[index], groupRoles[index + 1]];
    if (pair.includes("primary")) return Math.max(preferredGapPt, spacing.primarySeparationPt);
    if (pair.includes("caption")) return Math.min(preferredGapPt, spacing.compactPt);
    return preferredGapPt;
  });
  const minimumGaps = Array.from({ length: gapCount }, (_value, index) => {
    const pair = [groupRoles[index], groupRoles[index + 1]];
    if (pair.includes("primary")) return Math.max(spacing.normalPt, spacing.compactPt);
    if (pair.includes("caption")) return Math.max(4, spacing.compactPt * .75);
    return spacing.compactPt;
  });
  const minimumGapTotal = minimumGaps.reduce((sum, gap) => sum + gap, 0);
  const totalSourcePrimary = sourceBoxes.reduce((sum, box) => sum + (horizontal ? box.width : box.height), 0);
  const maximumSourceCross = Math.max(...sourceBoxes.map((box) => horizontal ? box.height : box.width));
  const availablePrimaryForObjects = regionPrimary - minimumGapTotal;
  const requiredScale = Math.max(0, Math.min(1, availablePrimaryForObjects / Math.max(1, totalSourcePrimary), regionCross / Math.max(1, maximumSourceCross)));
  const minimumScale = Math.max(.5, Math.min(1, input.minimumScale ?? .75));
  if (requiredScale < 1 - .001 && !input.allowResponsiveScale) {
    return { status: "infeasible", operation: "layout-group", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: [`The selected groups exceed the approved region at their current size and require proportional scaling to ${(requiredScale * 100).toFixed(1)}%. Enable bounded responsive scaling, choose a larger approved region, or split the composition.`] };
  }
  if (requiredScale < minimumScale - .001) {
    return { status: "infeasible", operation: "layout-group", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: [`The selected groups require ${(requiredScale * 100).toFixed(1)}% scaling, below the allowed ${(minimumScale * 100).toFixed(1)}% floor. Choose a larger layout region or continuation slide instead of forcing unreadable content.`] };
  }
  const scale = requiredScale < 1 ? requiredScale : 1;
  if (scale < 1 && objects.some((object) => !object.operations.resize)) {
    return { status: "infeasible", operation: "layout-group", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: ["At least one object in the responsive group cannot be resized safely. Choose a larger region or preserve that object outside the scaled group."] };
  }
  const boxes = sourceBoxes.map((box) => ({ ...box, width: box.width * scale, height: box.height * scale }));
  const totalPrimary = boxes.reduce((sum, box) => sum + (horizontal ? box.width : box.height), 0);
  const availableForGaps = regionPrimary - totalPrimary;
  if (availableForGaps < minimumGapTotal - .01) {
    return { status: "infeasible", operation: "layout-group", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: [`The selected groups require ${totalPrimary.toFixed(1)} pt before spacing, but the approved region has only ${regionPrimary.toFixed(1)} pt. Enlarge the region, choose another approved layout, resize content safely, or split the composition.`] };
  }
  const preferredGapTotal = preferredGaps.reduce((sum, gap) => sum + gap, 0);
  const compression = preferredGapTotal <= minimumGapTotal || availableForGaps >= preferredGapTotal ? 1 : (availableForGaps - minimumGapTotal) / (preferredGapTotal - minimumGapTotal);
  const gaps = preferredGaps.map((preferred, index) => minimumGaps[index] + (preferred - minimumGaps[index]) * compression);
  const stackSize = totalPrimary + gaps.reduce((sum, gap) => sum + gap, 0);
  let cursor = (horizontal ? regionPt.left : regionPt.top) + Math.max(0, (regionPrimary - stackSize) / 2);
  const targets = new Map<string, LayoutTarget>();
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    const box = boxes[index];
    const crossSize = horizontal ? box.height : box.width;
    if (crossSize > regionCross + .01) {
      return { status: "infeasible", operation: "layout-group", commands: [], objective: { movementCostPt: 0, maximumMovementPt: 0 }, diagnostics: [`Visual group ${index + 1} is ${crossSize.toFixed(1)} pt across, exceeding the approved region's ${regionCross.toFixed(1)} pt cross-axis capacity.`] };
    }
    const crossStart = horizontal ? regionPt.top : regionPt.left;
    const alignedCross = alignment === "start" ? crossStart : alignment === "end" ? crossStart + regionCross - crossSize : crossStart + (regionCross - crossSize) / 2;
    for (const object of group) {
      const source = objectBoxPt(object);
      const sourceGroupBox = sourceBoxes[index];
      const scaledRelativeLeft = (source.left - sourceGroupBox.left) * scale;
      const scaledRelativeTop = (source.top - sourceGroupBox.top) * scale;
      const targetGroupLeft = horizontal ? cursor : alignedCross;
      const targetGroupTop = horizontal ? alignedCross : cursor;
      targets.set(object.id, { left: targetGroupLeft + scaledRelativeLeft, top: targetGroupTop + scaledRelativeTop, width: source.width * scale, height: source.height * scale });
    }
    cursor += (horizontal ? box.width : box.height) + (gaps[index] ?? 0);
  }
  const groupByObjectId = new Map(input.groups.flatMap((group, index) => group.map((id) => [id, index] as const)));
  const result = buildCommands(deck, slideNumber, objects, targets, input.rationale, groupByObjectId, new Set(input.coordinatedObjectIds ?? []));
  result.operation = "layout-group";
  return result;
}

export interface SceneLayoutRegionRequest {
  id: string;
  groups: string[][];
  groupRoles?: GroupHierarchyRole[];
  regionPt: { left: number; top: number; width: number; height: number };
  mode: GroupLayoutMode;
  alignment: GroupLayoutAlignment;
  preferredGapPt?: number;
  allowResponsiveScale?: boolean;
  minimumScale?: number;
}

export function solveSceneToLayout(input: { deck: DeckJob; slideNumber: number; regions: SceneLayoutRegionRequest[]; rationale: string }): LayoutSolverResult {
  if (!input.regions.length || input.regions.length > 8) throw new Error("Scene fitting requires between 1 and 8 approved semantic regions.");
  if (new Set(input.regions.map((region) => region.id)).size !== input.regions.length) throw new Error("Each approved semantic region may be used only once in a scene-fit transaction.");
  const objectIds = input.regions.flatMap((region) => region.groups.flat());
  if (!objectIds.length || objectIds.length > 30) throw new Error("Scene fitting requires between 1 and 30 source-bound objects.");
  if (new Set(objectIds).size !== objectIds.length) throw new Error("Each source-bound object may belong to only one approved semantic region.");

  const targets = new Map<string, LayoutTarget>();
  const groupByObjectId = new Map<string, number>();
  let groupIndex = 0;
  for (const region of input.regions) {
    const solved = solveGroupLayout({ ...region, deck: input.deck, slideNumber: input.slideNumber, rationale: input.rationale, coordinatedObjectIds: objectIds });
    if (solved.status === "infeasible") return { ...solved, operation: "fit-scene-to-layout", diagnostics: solved.diagnostics.map((diagnostic) => `${region.id}: ${diagnostic}`) };
    const commandById = new Map(solved.commands.map((command) => [command.objectId, command]));
    for (const group of region.groups) {
      for (const objectId of group) {
        const object = input.deck.scene?.objects.find((candidate) => candidate.id === objectId);
        if (!object) throw new Error(`Scene fitting could not resolve ${objectId}.`);
        const command = commandById.get(objectId);
        targets.set(objectId, command ? { left: emuToPt(command.target.x), top: emuToPt(command.target.y), width: emuToPt(command.target.width), height: emuToPt(command.target.height) } : { left: emuToPt(object.geometry.x), top: emuToPt(object.geometry.y), width: emuToPt(object.geometry.width), height: emuToPt(object.geometry.height) });
        groupByObjectId.set(objectId, groupIndex);
      }
      groupIndex += 1;
    }
  }
  const objects = sceneObjects(input.deck, input.slideNumber, objectIds);
  const result = buildCommands(input.deck, input.slideNumber, objects, targets, input.rationale, groupByObjectId);
  result.operation = "fit-scene-to-layout";
  return result;
}
