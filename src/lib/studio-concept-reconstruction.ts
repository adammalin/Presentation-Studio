import { STUDIO_WEB_SCENE_VERSION, type StudioLayoutRecipe, type StudioWebFrame, type StudioWebNode, type StudioWebScene } from "../types";
import { markStudioVisualNeedsReconstructionReady } from "./studio-visual-needs";
import { recommendedStudioRecipe, recomposeStudioWebSlide } from "./studio-web-scene";

const EMU_PER_POINT = 12_700;

export interface StudioConceptReconstructionResult {
  scene: StudioWebScene;
  slideNumber: number;
  referenceId: string;
  recipe: StudioLayoutRecipe;
  mappedNodeIds: string[];
  unmappedNodeIds: string[];
  diagnostics: string[];
}

function zoneFrame(scene: StudioWebScene, zone: { x: number; y: number; width: number; height: number }): StudioWebFrame {
  const padding = 8 * EMU_PER_POINT;
  const raw = {
    x: Math.round(zone.x * scene.slideSize.width),
    y: Math.round(zone.y * scene.slideSize.height),
    width: Math.round(zone.width * scene.slideSize.width),
    height: Math.round(zone.height * scene.slideSize.height),
    rotation: 0,
  };
  return {
    x: raw.x + padding,
    y: raw.y + padding,
    width: Math.max(EMU_PER_POINT, raw.width - padding * 2),
    height: Math.max(EMU_PER_POINT, raw.height - padding * 2),
    rotation: 0,
  };
}

function contained(node: StudioWebNode, target: StudioWebFrame): StudioWebFrame {
  if (!["image", "native-object", "shape", "connector"].includes(node.kind) || node.sourceFrame.width <= 0 || node.sourceFrame.height <= 0) return target;
  const scale = Math.min(target.width / node.sourceFrame.width, target.height / node.sourceFrame.height);
  const width = Math.max(EMU_PER_POINT, Math.round(node.sourceFrame.width * scale));
  const height = Math.max(EMU_PER_POINT, Math.round(node.sourceFrame.height * scale));
  return { x: Math.round(target.x + (target.width - width) / 2), y: Math.round(target.y + (target.height - height) / 2), width, height, rotation: 0 };
}

function layoutNodes(nodes: StudioWebNode[], target: StudioWebFrame): Map<string, StudioWebFrame> {
  const placements = new Map<string, StudioWebFrame>();
  if (!nodes.length) return placements;
  const visual = nodes.every((node) => ["image", "native-object", "shape", "connector"].includes(node.kind));
  const gap = 12 * EMU_PER_POINT;
  if (visual && nodes.length > 1) {
    const columns = nodes.length <= 2 ? nodes.length : Math.ceil(Math.sqrt(nodes.length));
    const rows = Math.ceil(nodes.length / columns);
    const cellWidth = (target.width - gap * (columns - 1)) / columns;
    const cellHeight = (target.height - gap * (rows - 1)) / rows;
    nodes.forEach((node, index) => {
      const column = index % columns;
      const row = Math.floor(index / columns);
      placements.set(node.id, contained(node, { x: Math.round(target.x + column * (cellWidth + gap)), y: Math.round(target.y + row * (cellHeight + gap)), width: Math.round(cellWidth), height: Math.round(cellHeight), rotation: 0 }));
    });
    return placements;
  }
  const height = Math.max(EMU_PER_POINT, Math.round((target.height - gap * Math.max(0, nodes.length - 1)) / nodes.length));
  nodes.forEach((node, index) => placements.set(node.id, contained(node, { x: target.x, y: target.y + index * (height + gap), width: target.width, height, rotation: 0 })));
  return placements;
}

function candidateNodes(nodes: StudioWebNode[], role: string, assigned: Set<string>) {
  const available = nodes.filter((node) => node.visible && !assigned.has(node.id) && !["footer", "slide-number", "date", "logo"].includes(node.role));
  if (role === "title") return available.filter((node) => node.role === "title");
  if (role === "caption") return available.filter((node) => node.role === "caption" || node.role === "label");
  if (role === "primary-visual") return available.filter((node) => node.kind === "image" || node.kind === "table" || ["chart", "media"].includes(node.role) || node.kind === "native-object");
  if (role === "supporting-evidence") return available.filter((node) => node.kind === "text" && node.role !== "title");
  if (role === "other") return available;
  return [];
}

export function reconstructStudioConcept(scene: StudioWebScene, slideNumber: number, referenceId: string, requestedRecipe?: StudioLayoutRecipe): StudioConceptReconstructionResult {
  const sourceSlide = scene.slides.find((slide) => slide.slideNumber === slideNumber);
  const reference = sourceSlide?.conceptReferences?.find((item) => item.id === referenceId);
  if (!sourceSlide || !reference) throw new Error("The requested concept reference is not attached to this Studio slide.");
  if (reference.sourceTextHash !== sourceSlide.sourceTextHash) throw new Error("The concept reference is stale because the source-content binding changed.");
  if (!reference.visualNeedId) throw new Error("Concept reconstruction requires a linked visual need so the result can be reviewed against its communication job.");
  if (!reference.blueprint.zones.length) throw new Error("The concept needs at least one normalized semantic zone before editable reconstruction.");
  const recipe = requestedRecipe ?? recommendedStudioRecipe(sourceSlide);
  if (recipe === "source" || recipe === "template-layout") throw new Error("Concept reconstruction requires a shared Studio composition recipe; source and converted-template layout modes use separate workflows.");
  let recomposed = recomposeStudioWebSlide(scene, slideNumber, recipe, undefined, `Reconstruct approved ${reference.approvedInfluences.join(", ")} from concept ${reference.id} with exact source-bound editable content.`);
  const slide = recomposed.slides.find((item) => item.slideNumber === slideNumber)!;
  const assigned = new Set<string>();
  const placements = new Map<string, StudioWebFrame>();
  for (const zone of reference.blueprint.zones.filter((item) => item.role !== "footer-safe")) {
    const candidates = candidateNodes(slide.nodes, zone.role, assigned);
    if (!candidates.length) continue;
    const local = layoutNodes(candidates, zoneFrame(recomposed, zone));
    for (const [nodeId, frame] of local) {
      assigned.add(nodeId);
      placements.set(nodeId, frame);
    }
  }
  const now = new Date().toISOString();
  recomposed = {
    ...recomposed,
    revision: `${recomposed.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: recomposed.slides.map((item) => item.slideNumber === slideNumber ? {
      ...item,
      status: "designed",
      designRationale: `Editable reconstruction of concept ${reference.id}: ${reference.blueprint.summary}`.slice(0, 1_000),
      nodes: item.nodes.map((node) => placements.has(node.id) ? { ...node, frame: placements.get(node.id)! } : node),
      qualityReview: undefined,
      updatedAt: now,
    } : item),
  };
  recomposed = markStudioVisualNeedsReconstructionReady(recomposed, slideNumber, [reference.visualNeedId]);
  const finalSlide = recomposed.slides.find((item) => item.slideNumber === slideNumber)!;
  const mappedNodeIds = [...placements.keys()];
  const unmappedNodeIds = finalSlide.nodes.filter((node) => node.visible && !assigned.has(node.id) && !["footer", "slide-number", "date", "logo"].includes(node.role)).map((node) => node.id);
  return {
    scene: recomposed,
    slideNumber,
    referenceId,
    recipe,
    mappedNodeIds,
    unmappedNodeIds,
    diagnostics: unmappedNodeIds.length ? [`${unmappedNodeIds.length} semantic node${unmappedNodeIds.length === 1 ? "" : "s"} retained the shared recipe placement because the concept supplied no matching zone.`] : [],
  };
}

