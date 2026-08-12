import type { DeckJob } from "../types";
import type { GeometryEditRequest } from "./cleanup";
import type { TemplateContentKind, TemplateSemanticSlot } from "./layout-semantics";
import type { TemplateLayoutPreview } from "./template-catalog";

export interface SemanticSlotBinding {
  objectId: string;
  slotId: string;
  fit: "fill" | "contain" | "align-horizontal";
  insetInches?: number;
}

function contentKindFor(kind: string): TemplateContentKind | undefined {
  if (["text", "shape"].includes(kind)) return "text";
  if (kind === "picture") return "image";
  if (kind === "table") return "table";
  if (kind === "chart") return "chart";
  return undefined;
}

function targetForSlot(
  slot: TemplateSemanticSlot,
  source: { width: number; height: number },
  binding: SemanticSlotBinding,
  sourcePosition?: { x: number; y: number },
): { x: number; y: number; width: number; height: number } {
  const inset = Math.round(Math.max(0, Math.min(.25, binding.insetInches ?? 0)) * 914_400);
  const frame = {
    x: slot.x + inset,
    y: slot.y + inset,
    width: Math.max(91_440, slot.width - inset * 2),
    height: Math.max(91_440, slot.height - inset * 2),
  };
  if (binding.fit === "align-horizontal") {
    if (!sourcePosition) throw new Error("Horizontal semantic alignment requires the source position.");
    return { x: frame.x, y: sourcePosition.y, width: frame.width, height: source.height };
  }
  if (binding.fit !== "contain") return frame;
  const scale = Math.min(frame.width / source.width, frame.height / source.height);
  const width = Math.max(91_440, Math.round(source.width * scale));
  const height = Math.max(91_440, Math.round(source.height * scale));
  return { x: Math.round(frame.x + (frame.width - width) / 2), y: Math.round(frame.y + (frame.height - height) / 2), width, height };
}

export function semanticRecompositionRequests(input: {
  deck: DeckJob;
  slideNumber: number;
  layout: TemplateLayoutPreview;
  bindings: SemanticSlotBinding[];
  rationale: string;
}): { requests: GeometryEditRequest[]; unboundObjectIds: string[] } {
  const { deck, slideNumber, layout, bindings } = input;
  if (!deck.audit || !deck.scene) throw new Error("A current audit and hybrid scene are required before semantic recomposition.");
  const scene = deck.scene;
  if (!layout.semantic) throw new Error("The requested layout has no semantic slot contract.");
  if (bindings.length === 0 || bindings.length > 20) throw new Error("Bind between 1 and 20 source objects to approved layout slots.");
  if (new Set(bindings.map((binding) => binding.objectId)).size !== bindings.length) throw new Error("Each source object may be bound only once.");
  if (new Set(bindings.map((binding) => binding.slotId)).size !== bindings.length) throw new Error("Each approved layout slot may receive only one source object in this recomposition slice.");

  const requests = bindings.map((binding): GeometryEditRequest => {
    const object = scene.objects.find((item) => item.id === binding.objectId && item.slideNumber === slideNumber);
    const auditObject = deck.audit?.editableObjects.find((item) => item.id === binding.objectId && item.slideNumber === slideNumber);
    if (!object || !auditObject) throw new Error(`Object ${binding.objectId} is not source-bound to slide ${slideNumber}.`);
    const slot = layout.semantic?.slots.find((item) => item.id === binding.slotId);
    if (!slot) throw new Error(`Slot ${binding.slotId} is not present in layout ${layout.name}.`);
    const contentKind = contentKindFor(object.kind);
    if (!contentKind || !slot.acceptedContent.includes(contentKind)) throw new Error(`${object.name} (${object.kind}) is not compatible with the ${slot.role} slot in ${layout.name}.`);
    if (!object.operations.move || !object.operations.resize) throw new Error(`${object.name} cannot be safely moved and resized by the current scene contract.`);
    if (object.kind === "picture" && binding.fit === "fill") throw new Error(`${object.name} must use contain because crop-aware picture filling is not implemented yet.`);
    if (object.kind === "picture" && binding.fit === "align-horizontal") throw new Error(`${object.name} must use contain because horizontal-only picture resizing can distort its frame.`);
    const target = targetForSlot(slot, auditObject.geometry, binding, auditObject.geometry);
    if (["text", "shape"].includes(object.kind) && binding.fit === "fill") {
      const source = auditObject.geometry;
      const sourceArea = source.width * source.height;
      const targetArea = target.width * target.height;
      const verticalDisplacement = Math.abs(target.y - source.y);
      const substantialFrameChange = target.height < source.height * .9 || targetArea < sourceArea * .9 || verticalDisplacement > .25 * 914_400;
      if (substantialFrameChange) throw new Error(`${object.name} would substantially replace a proven text frame. Use align-horizontal to adopt the approved layout edges while preserving vertical fit, or stage a measured text-fitting edit after native review.`);
    }
    return {
      objectId: object.id,
      target,
      rationale: `${input.rationale.trim()} Map ${object.name} to the approved ${slot.role} region in ${layout.name}.`.trim().slice(0, 700),
      author: "ai",
      constraints: {
        allowIntentionalOverlap: false,
        allowFitRisk: false,
        allowSafeArea: slot.x < 228_600 || slot.y < 228_600 || slot.x + slot.width > scene.slideSize.width - 228_600 || slot.y + slot.height > scene.slideSize.height - 228_600,
        allowAspectRatioChange: false,
      },
    };
  });
  const bound = new Set(bindings.map((binding) => binding.objectId));
  const unboundObjectIds = scene.objects.filter((object) => object.slideNumber === slideNumber && !bound.has(object.id)).map((object) => object.id);
  return { requests, unboundObjectIds };
}
