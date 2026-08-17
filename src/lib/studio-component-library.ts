import type {
  StudioComponentDefinition,
  StudioComponentRole,
  StudioComponentSurface,
  StudioWebNode,
  StudioWebScene,
  StudioWebSlide,
} from "../types";
import { STUDIO_WEB_SCENE_VERSION } from "../types";

const HEX = /^#([0-9a-f]{6})$/i;

export interface StudioComponentAdoptionResult {
  scene: StudioWebScene;
  definition: StudioComponentDefinition;
  affectedSlideNumbers: number[];
  affectedNodeIds: string[];
  skippedNodeIds: string[];
}

function channel(value: string, offset: number): number {
  return Number.parseInt(value.slice(offset, offset + 2), 16) / 255;
}

function linear(value: number): number {
  return value <= .03928 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4;
}

function luminance(value: string | undefined): number {
  const match = value?.match(HEX);
  if (!match) return 1;
  const hex = match[1];
  return .2126 * linear(channel(hex, 0)) + .7152 * linear(channel(hex, 2)) + .0722 * linear(channel(hex, 4));
}

function cloneStyle(style: StudioWebNode["style"]): StudioWebNode["style"] {
  return { ...style, paddingPt: { ...style.paddingPt } };
}

function surfaceFor(slide: StudioWebSlide, node: StudioWebNode): StudioComponentSurface {
  const fill = node.style.background && HEX.test(node.style.background) ? node.style.background : slide.background;
  return luminance(fill) < .36 ? "dark" : "light";
}

function definitionId(role: StudioComponentRole, surface: StudioComponentSurface): string {
  return `ornl-component-${role}-${surface}`;
}

function sameStyle(left: StudioWebNode["style"], right: StudioWebNode["style"]): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function lockedBySourceFigure(slide: StudioWebSlide, nodeId: string): boolean {
  return slide.figureTreatments.some((treatment) => treatment.nodeIds.includes(nodeId) && treatment.verificationStatus === "source-locked");
}

function nextRevision(scene: StudioWebScene, now: string): string {
  return `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`;
}

export function compatibleStudioComponentInstances(scene: StudioWebScene, slideNumber: number, nodeId: string): Array<{ slideNumber: number; nodeId: string }> {
  const sourceSlide = scene.slides.find((slide) => slide.slideNumber === slideNumber);
  const sourceNode = sourceSlide?.nodes.find((node) => node.id === nodeId);
  if (!sourceSlide || !sourceNode?.component) return [];
  const surface = surfaceFor(sourceSlide, sourceNode);
  return scene.slides.flatMap((slide) => slide.nodes.flatMap((node) => node.visible && node.component?.role === sourceNode.component?.role && surfaceFor(slide, node) === surface
    ? [{ slideNumber: slide.slideNumber, nodeId: node.id }]
    : []));
}

export function adoptStudioComponentStyle(scene: StudioWebScene, input: {
  slideNumber: number;
  nodeId: string;
  name?: string;
  targetSlideNumbers?: number[];
}): StudioComponentAdoptionResult {
  const sourceSlide = scene.slides.find((slide) => slide.slideNumber === input.slideNumber);
  const sourceNode = sourceSlide?.nodes.find((node) => node.id === input.nodeId);
  if (!sourceSlide || !sourceNode) throw new Error("The source component node is not present in the current Studio scene revision.");
  if (!sourceNode.component) throw new Error("Choose a node created by a shared Studio recipe before publishing a reusable component style.");
  if (sourceNode.locked || lockedBySourceFigure(sourceSlide, sourceNode.id)) throw new Error("A locked or source-preserved figure node cannot define a reusable component style.");
  const targetSet = input.targetSlideNumbers?.length ? new Set(input.targetSlideNumbers) : undefined;
  if (targetSet?.has(input.slideNumber) === false) targetSet.add(input.slideNumber);
  const surface = surfaceFor(sourceSlide, sourceNode);
  const id = definitionId(sourceNode.component.role, surface);
  const now = new Date().toISOString();
  const definition: StudioComponentDefinition = {
    id,
    name: input.name?.trim().slice(0, 120) || `${sourceNode.component.role.replaceAll("-", " ")} · ${surface} surface`,
    role: sourceNode.component.role,
    surface,
    sourceNodeId: sourceNode.id,
    adoptedFromSlideNumber: sourceSlide.slideNumber,
    style: cloneStyle(sourceNode.style),
    updatedAt: now,
  };
  const affectedSlideNumbers = new Set<number>();
  const affectedNodeIds: string[] = [];
  const skippedNodeIds: string[] = [];
  const slides = scene.slides.map((slide) => {
    if (targetSet && !targetSet.has(slide.slideNumber)) return slide;
    let changed = false;
    const nodes = slide.nodes.map((node) => {
      if (!node.visible || node.component?.role !== definition.role || surfaceFor(slide, node) !== definition.surface) return node;
      if (node.locked || lockedBySourceFigure(slide, node.id)) {
        skippedNodeIds.push(node.id);
        return node;
      }
      changed = true;
      affectedNodeIds.push(node.id);
      return { ...node, style: cloneStyle(definition.style), component: { ...node.component, definitionId: definition.id } };
    });
    if (!changed) return slide;
    affectedSlideNumbers.add(slide.slideNumber);
    return { ...slide, nodes, qualityReview: undefined, updatedAt: now };
  });
  if (!affectedNodeIds.length) throw new Error("No compatible unlocked component instances were found for this role and surface.");
  const componentLibrary = [...(scene.componentLibrary ?? []).filter((item) => item.id !== definition.id), definition];
  return {
    scene: { ...scene, revision: nextRevision(scene, now), componentLibrary, slides },
    definition,
    affectedSlideNumbers: [...affectedSlideNumbers].sort((left, right) => left - right),
    affectedNodeIds,
    skippedNodeIds,
  };
}

export function applyStudioComponentDefinition(scene: StudioWebScene, input: { definitionId: string; targetSlideNumbers?: number[] }): StudioComponentAdoptionResult {
  const definition = scene.componentLibrary?.find((item) => item.id === input.definitionId);
  if (!definition) throw new Error("The requested reusable component definition is not present in this Studio scene.");
  const sourceSlide = scene.slides.find((slide) => slide.slideNumber === definition.adoptedFromSlideNumber);
  const sourceNode = sourceSlide?.nodes.find((node) => node.id === definition.sourceNodeId);
  if (!sourceSlide || !sourceNode?.component) throw new Error("The reusable component definition no longer has a valid source instance.");
  if (!sameStyle(sourceNode.style, definition.style)) {
    const refreshedScene = {
      ...scene,
      slides: scene.slides.map((slide) => slide.slideNumber === sourceSlide.slideNumber ? { ...slide, nodes: slide.nodes.map((node) => node.id === sourceNode.id ? { ...node, style: cloneStyle(definition.style) } : node) } : slide),
    };
    return adoptStudioComponentStyle(refreshedScene, { slideNumber: sourceSlide.slideNumber, nodeId: sourceNode.id, name: definition.name, targetSlideNumbers: input.targetSlideNumbers });
  }
  return adoptStudioComponentStyle(scene, { slideNumber: sourceSlide.slideNumber, nodeId: sourceNode.id, name: definition.name, targetSlideNumbers: input.targetSlideNumbers });
}
