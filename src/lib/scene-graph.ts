import type {
  DeckJob,
  PptxAudit,
  PresentationScene,
  PresentationSceneObject,
  SceneFidelityCounts,
  SceneFidelityState,
  SceneSemanticRole,
  SlideEditableObject,
  TextBoxInventoryItem,
} from "../types";
import {
  PRESENTATION_SCENE_SCHEMA,
  PRESENTATION_SCENE_VERSION,
  PRESERVATION_ENVELOPE_SCHEMA,
  PRESERVATION_ENVELOPE_VERSION,
} from "../types";

export type SceneCompileInput = Pick<DeckJob,
  "sourceResourceId" | "sourceSha256" | "templateClassification" | "targetTemplateId" | "targetTemplateDecisionSource" | "protectedSlideNumbers"
> & { audit: PptxAudit };

export function emptySceneFidelityCounts(): SceneFidelityCounts {
  return { "editable-native": 0, "preserved-native": 0, "conversion-required": 0, "unsupported-blocking": 0 };
}

function countFidelity(states: SceneFidelityState[]): SceneFidelityCounts {
  const counts = emptySceneFidelityCounts();
  for (const state of states) counts[state] += 1;
  return counts;
}

function roleForObject(object: SlideEditableObject, textBox?: TextBoxInventoryItem): SceneSemanticRole {
  if (object.kind === "picture") return "image";
  if (object.kind === "table") return "table";
  if (object.kind === "chart") return "chart";
  if (object.kind === "connector") return "connector";
  if (object.kind === "group") return "group";
  if (textBox && ["title", "body", "caption", "label"].includes(textBox.role)) return textBox.role as SceneSemanticRole;
  if (object.kind === "shape" && !object.textHash) return "decoration";
  return "other";
}

function fidelityForObject(object: SlideEditableObject): { state: SceneFidelityState; reason: string } {
  if (["text", "shape", "table"].includes(object.kind)) {
    return { state: "editable-native", reason: "Studio represents this object's native geometry and its supported text, shape, or table properties." };
  }
  if (object.kind === "graphic-frame") {
    return { state: "conversion-required", reason: "The PowerPoint graphic frame is preserved, but its internal content must be converted or gain a dedicated adapter before Studio can edit it." };
  }
  return { state: "preserved-native", reason: "Studio keeps this native PowerPoint object intact and exposes only the supported whole-object operations." };
}

function sceneObject(
  input: SceneCompileInput,
  object: SlideEditableObject,
  zIndex: number,
  slidePart: string,
): PresentationSceneObject {
  const textBox = input.audit.textBoxes.find((item) => item.slideNumber === object.slideNumber && item.shapeId === object.shapeId);
  const table = object.tableId ? input.audit.tables.find((item) => item.id === object.tableId) : undefined;
  const fidelity = fidelityForObject(object);
  const hasText = Boolean(object.textHash);
  const isEditableText = ["text", "shape"].includes(object.kind) && hasText;
  const isTable = object.kind === "table";
  const protectedSlide = input.protectedSlideNumbers.includes(object.slideNumber);
  return {
    id: object.id,
    slideId: `slide-${object.slideNumber}`,
    slideNumber: object.slideNumber,
    shapeId: object.shapeId,
    name: object.name,
    kind: object.kind,
    sourceElement: object.sourceElement,
    semanticRole: roleForObject(object, textBox),
    fidelityState: fidelity.state,
    fidelityReason: fidelity.reason,
    geometry: { ...object.geometry },
    zIndex,
    sourceLocator: { slidePart, shapeId: object.shapeId, tableId: object.tableId, pictureId: object.pictureId },
    representation: {
      geometry: "native",
      text: hasText ? "native" : "none",
      style: isEditableText || isTable ? "partial" : fidelity.state === "editable-native" ? "partial" : "preserved",
      internalStructure: isTable ? "native" : fidelity.state === "editable-native" ? "partial" : "preserved",
    },
    operations: {
      move: !protectedSlide && object.canMove,
      resize: !protectedSlide && object.canResize,
      restyle: !protectedSlide && (isEditableText || isTable),
      editText: false,
      editTableStyle: !protectedSlide && isTable,
      replaceMedia: false,
      editChartData: false,
      editInternalStructure: false,
    },
    contentHash: object.textHash ?? table?.contentHash,
    protected: protectedSlide,
  };
}

export function compilePresentationScene(input: SceneCompileInput): PresentationScene {
  const objects: PresentationSceneObject[] = [];
  for (const slide of input.audit.slides) {
    const slidePart = slide.sourcePart ?? `ppt/slides/slide${slide.number}.xml`;
    const slideObjects = input.audit.editableObjects.filter((object) => object.slideNumber === slide.number);
    slideObjects.forEach((object, index) => objects.push(sceneObject(input, object, index, slidePart)));
  }

  const slides = input.audit.slides.map((slide) => {
    const slideObjects = objects.filter((object) => object.slideNumber === slide.number);
    const fidelityCounts = countFidelity(slideObjects.map((object) => object.fidelityState));
    const protectedSlide = input.protectedSlideNumbers.includes(slide.number);
    return {
      id: slide.id,
      number: slide.number,
      sourcePart: slide.sourcePart ?? `ppt/slides/slide${slide.number}.xml`,
      sourcePartSha256: slide.sourcePartSha256,
      relationshipPart: slide.relationshipPart,
      relationshipPartSha256: slide.relationshipPartSha256,
      sourceTextHash: slide.textHash,
      objectIds: slideObjects.map((object) => object.id),
      fidelityCounts,
      preservationRequired: protectedSlide || slideObjects.some((object) => object.fidelityState !== "editable-native"),
      protected: protectedSlide,
    };
  });
  const blockingFeatures: Array<"macros" | "ole-objects" | "external-relationships"> = [];
  if (input.audit.containsMacros) blockingFeatures.push("macros");
  if (input.audit.containsOleObjects) blockingFeatures.push("ole-objects");
  if (input.audit.containsExternalRelationships) blockingFeatures.push("external-relationships");
  const templateBinding = {
    sourceClassification: input.templateClassification,
    targetTemplateId: input.targetTemplateId,
    targetDecisionSource: input.targetTemplateDecisionSource,
  };
  return {
    schema: PRESENTATION_SCENE_SCHEMA,
    version: PRESENTATION_SCENE_VERSION,
    revision: `${input.sourceSha256}:scene-v${PRESENTATION_SCENE_VERSION}:${input.targetTemplateId ?? "source"}`,
    sourceSha256: input.sourceSha256,
    slideSize: { ...input.audit.slideSize },
    templateBinding,
    slides,
    objects,
    fidelityCounts: countFidelity(objects.map((object) => object.fidelityState)),
    preservationEnvelope: {
      schema: PRESERVATION_ENVELOPE_SCHEMA,
      version: PRESERVATION_ENVELOPE_VERSION,
      sourceResourceId: input.sourceResourceId,
      sourceSha256: input.sourceSha256,
      sourceBytesAuthoritative: true,
      nativeRenderAuthoritativeForAppearance: true,
      exportStrategy: "surgical-ooxml-overlay",
      packageFileCount: input.audit.packageFileCount,
      expandedByteLength: input.audit.expandedByteLength,
      protectedFeatures: {
        macros: input.audit.containsMacros,
        oleObjects: input.audit.containsOleObjects,
        externalRelationships: input.audit.containsExternalRelationships,
      },
      blockingFeatures,
      slides: slides.map((slide) => ({
        slideId: slide.id,
        slideNumber: slide.number,
        sourcePart: slide.sourcePart,
        sourcePartSha256: slide.sourcePartSha256,
        relationshipPart: slide.relationshipPart,
        relationshipPartSha256: slide.relationshipPartSha256,
        sourceTextHash: slide.sourceTextHash,
        objectIds: slide.objectIds,
      })),
    },
  };
}

export function sceneNeedsRebuild(deck: DeckJob): boolean {
  return Boolean(deck.audit) && (
    !deck.scene
    || deck.scene.schema !== PRESENTATION_SCENE_SCHEMA
    || deck.scene.version !== PRESENTATION_SCENE_VERSION
    || deck.scene.sourceSha256 !== deck.sourceSha256
    || deck.scene.templateBinding.targetTemplateId !== deck.targetTemplateId
    || deck.scene.templateBinding.targetDecisionSource !== deck.targetTemplateDecisionSource
  );
}
