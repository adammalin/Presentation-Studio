import type { DeckJob, DesignThread } from "../types";
import type { NativeRenderResult } from "./desktop";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { rankLayoutCompatibility, type LayoutContentProfile, type TemplateLayoutIntent } from "./layout-semantics";
import type { TemplateCatalog } from "./template-catalog";

export const DESIGN_WORK_ORDER_SCHEMA = "presentation-studio/design-work-order" as const;
export const DESIGN_WORK_ORDER_VERSION = 1 as const;

export interface RepresentativeSlide {
  slideNumber: number;
  role: "cover" | "content" | "diagram" | "image-heavy" | "table";
  reason: string;
}

export function contentProfileForSlide(deck: DeckJob, slideNumber: number): LayoutContentProfile {
  const slide = deck.audit?.slides.find((item) => item.number === slideNumber);
  if (!slide) throw new Error("The requested slide is not present in the current audit.");
  const textBoxes = (deck.audit?.textBoxes ?? []).filter((textBox) => textBox.slideNumber === slideNumber);
  const bodyTextBoxes = textBoxes.filter((textBox) => textBox.role !== "title" && textBox.characterCount > 0);
  const slideHeight = deck.audit?.slideSize.height ?? 6_858_000;
  const captionTextBoxes = bodyTextBoxes.filter((textBox) => ["caption", "label"].includes(textBox.role) || textBox.geometry.height / slideHeight < 0.105);
  const primaryBodyTextBoxes = bodyTextBoxes.filter((textBox) => !captionTextBoxes.includes(textBox));
  const fallbackBodyCharacters = Math.max(0, slide.text.length - slide.title.length);
  const bodyCharacterCount = primaryBodyTextBoxes.length ? primaryBodyTextBoxes.reduce((sum, textBox) => sum + textBox.characterCount, 0) : bodyTextBoxes.length ? 0 : fallbackBodyCharacters;
  const normalizedTitle = slide.title.toLowerCase();
  const desiredIntent: TemplateLayoutIntent = slideNumber === 1 ? "cover"
    : /\b(conclusion|summary|questions?|thank you)\b/.test(normalizedTitle) ? "conclusion"
      : slide.tableCount + slide.chartCount > 0 ? "data"
        : slide.pictureCount > 0 ? "visual"
          : bodyTextBoxes.filter((textBox) => textBox.role === "body").length > 1 ? "comparison"
            : "content";
  return {
    titleCharacterCount: slide.title.trim().length,
    bodyBlockCount: primaryBodyTextBoxes.length || !textBoxes.length && fallbackBodyCharacters > 0 ? Math.max(1, primaryBodyTextBoxes.length) : 0,
    captionBlockCount: captionTextBoxes.length,
    bodyCharacterCount,
    imageCount: slide.pictureCount,
    imageAspectRatios: (deck.audit?.pictures ?? []).filter((picture) => picture.slideNumber === slideNumber && picture.widthEmu && picture.heightEmu).map((picture) => Number(picture.widthEmu) / Number(picture.heightEmu)),
    tableCount: slide.tableCount,
    chartCount: slide.chartCount,
    mediaCount: 0,
    desiredIntent,
  };
}

function highestSlide(deck: DeckJob, score: (slideNumber: number) => number, predicate: (slideNumber: number) => boolean): number | undefined {
  return deck.audit?.slides.map((slide) => slide.number).filter(predicate).sort((left, right) => score(right) - score(left) || left - right)[0];
}

export function selectRepresentativeSlides(deck: DeckJob, limit = 5): RepresentativeSlide[] {
  if (!deck.audit?.slideCount) return [];
  const result: RepresentativeSlide[] = [];
  const used = new Set<number>();
  const add = (slideNumber: number | undefined, role: RepresentativeSlide["role"], reason: string) => {
    if (!slideNumber || used.has(slideNumber) || result.length >= limit) return;
    used.add(slideNumber);
    result.push({ slideNumber, role, reason });
  };
  add(1, "cover", "Opening hierarchy, identity, and title alignment establish the deck's visual standard.");
  const plainContent = highestSlide(deck,
    (number) => (deck.audit?.textBoxes ?? []).filter((box) => box.slideNumber === number && box.role === "body").length * 100 + (deck.audit?.slides.find((slide) => slide.number === number)?.text.length ?? 0),
    (number) => number > 1 && Boolean(deck.audit?.slides.some((slide) => slide.number === number && slide.tableCount === 0 && slide.pictureCount === 0 && slide.chartCount === 0 && slide.connectorCount === 0)),
  );
  add(plainContent, "content", "Representative text-led content tests hierarchy, spacing, and readable exact-content fit.");
  const diagram = highestSlide(deck,
    (number) => { const slide = deck.audit?.slides.find((item) => item.number === number); return (slide?.connectorCount ?? 0) * 10 + (slide?.pictureCount ?? 0); },
    (number) => Boolean(deck.audit?.slides.some((slide) => slide.number === number && slide.connectorCount > 0)),
  );
  add(diagram, "diagram", "Connector-rich content tests relationship preservation and whole-composition alignment.");
  const imageHeavy = highestSlide(deck,
    (number) => deck.audit?.slides.find((slide) => slide.number === number)?.pictureCount ?? 0,
    (number) => Boolean(deck.audit?.slides.some((slide) => slide.number === number && slide.pictureCount > 0)),
  );
  add(imageHeavy, "image-heavy", "The largest image set tests crop-safe placement, visual hierarchy, and caption relationships.");
  const tableSlide = highestSlide(deck,
    (number) => (deck.audit?.tables ?? []).filter((table) => table.slideNumber === number).reduce((sum, table) => sum + table.totalCellCharacterCount + table.rowCount * table.columnCount * 10, 0),
    (number) => Boolean(deck.audit?.slides.some((slide) => slide.number === number && slide.tableCount > 0)),
  );
  add(tableSlide, "table", "The densest native table tests structure preservation, legibility, and surrounding layout.");
  for (const slide of deck.audit.slides) add(slide.number, "content", "Additional representative slide selected to complete the bounded qualification set.");
  return result;
}

function visualEvidence(render: NativeRenderResult | undefined, slideNumber: number) {
  const slide = render?.status === "ready" ? render.slides.find((item) => item.number === slideNumber) : undefined;
  return slide && render ? {
    authority: render.authoritative ? "powerpoint-native" as const : "studio-approximate" as const,
    renderer: render.renderer,
    pipeline: render.pipeline,
    powerPointVersion: render.powerPointVersion,
    rasterSha256: slide.sha256,
    width: slide.width,
    height: slide.height,
  } : { authority: "unavailable" as const, renderer: render?.renderer, warnings: render?.warnings ?? ["A native Current render has not been acquired for this work order."] };
}

export function buildSlideDesignWorkOrder(input: {
  deck: DeckJob;
  slideNumber: number;
  projectUpdatedAt: string;
  templateCatalog: TemplateCatalog;
  currentRender?: NativeRenderResult;
  threads?: DesignThread[];
}) {
  const { deck, slideNumber, projectUpdatedAt, templateCatalog } = input;
  if (!deck.audit || !deck.scene) throw new Error("A current audit and hybrid scene are required before building a design work order.");
  const slide = deck.audit.slides.find((item) => item.number === slideNumber);
  const sceneSlide = deck.scene.slides.find((item) => item.number === slideNumber);
  if (!slide || !sceneSlide) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
  const profile = contentProfileForSlide(deck, slideNumber);
  const ranked = rankLayoutCompatibility(templateCatalog.layouts, profile).slice(0, 6).map((result) => {
    const layout = templateCatalog.layouts.find((item) => item.id === result.layoutId)!;
    return { ...result, layout: { id: layout.id, name: layout.name, sourcePart: layout.sourcePart, category: layout.category, semantic: layout.semantic } };
  });
  const objects = deck.scene.objects.filter((object) => object.slideNumber === slideNumber).map((object) => {
    const textBox = deck.audit?.textBoxes.find((box) => box.slideNumber === slideNumber && box.shapeId === object.shapeId);
    return {
      ...object,
      geometryInches: { x: object.geometry.x / 914_400, y: object.geometry.y / 914_400, width: object.geometry.width / 914_400, height: object.geometry.height / 914_400 },
      exactText: textBox?.text,
      opticalText: textBox ? {
        textInsetsEmu: textBox.textInsets,
        paragraphLeftMarginsEmu: textBox.paragraphLeftMarginsEmu,
        paragraphIndentsEmu: textBox.paragraphIndentsEmu,
        bulletParagraphCount: textBox.bulletParagraphCount,
        opticalLeftOffsetEmu: textBox.opticalLeftOffsetEmu,
        estimatedOpticalLeftEmu: textBox.estimatedOpticalLeftEmu,
        estimatedOpticalLeftInches: textBox.estimatedOpticalLeftEmu / 914_400,
        confidence: textBox.opticalAlignmentConfidence,
        instruction: "Align visible text starts, not only shape x coordinates. Treat partial-inheritance values as visual-review evidence, not exact inherited PowerPoint layout values.",
      } : undefined,
    };
  });
  const submittedThreads = (input.threads ?? []).filter((thread) => thread.deckId === deck.id && thread.slideNumber === slideNumber && thread.status !== "note");
  const currentVisualEvidence = visualEvidence(input.currentRender, slideNumber);
  const revision = `${projectUpdatedAt}:${deck.scene.revision}:slide-${slideNumber}:standard-${PRESENTATION_DESIGN_STANDARD.version}:template-${templateCatalog.sha256}:render-${"rasterSha256" in currentVisualEvidence ? currentVisualEvidence.rasterSha256 : "unavailable"}`;
  return {
    schema: DESIGN_WORK_ORDER_SCHEMA,
    version: DESIGN_WORK_ORDER_VERSION,
    revision,
    projectUpdatedAt,
    deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId, operationScope: deck.operationScope },
    slide: {
      id: slide.id,
      number: slide.number,
      exactTitle: slide.title,
      exactVisibleText: slide.text,
      sourceTextHash: slide.textHash,
      sourcePartSha256: slide.sourcePartSha256,
      sceneRevision: deck.scene.revision,
      fidelityCounts: sceneSlide.fidelityCounts,
      preservationRequired: sceneSlide.preservationRequired,
    },
    closedContentInventory: {
      onlyAllowedVisibleText: slide.text,
      lockedTextHash: slide.textHash,
      allowedVisuals: "Existing source visuals plus locally authorized Resources explicitly shared for this session. No invented technical content or pseudo-official marks.",
      mustPreserve: ["exact visible wording", "technical meaning", "native tables/charts/equations", "source media identity", "notes/comments/relationships", "slide count and order"],
      forbidden: ["rewriting", "omitting approved content", "silent type shrink", "flattening supported objects", "fabricated icons/images/data/claims", "browser-only effects without PowerPoint representation"],
    },
    communicationJob: `Improve slide ${slideNumber} as a restrained ORNL technical presentation slide while preserving every approved word and technical relationship.`,
    contentProfile: profile,
    objects,
    findings: deck.audit.findings.filter((finding) => finding.slideNumber === slideNumber),
    geometryChecks: deck.audit.layoutReviews.filter((finding) => finding.slideNumber === slideNumber),
    submittedThreads: submittedThreads.map((thread) => ({ id: thread.id, baseRevision: thread.baseRevision, anchor: thread.anchor, comment: thread.comment, status: thread.status })),
    layoutCandidates: ranked,
    currentVisualEvidence,
    designRules: {
      expression: "restrained",
      typography: PRESENTATION_DESIGN_STANDARD.defaults.typography,
      palette: PRESENTATION_DESIGN_STANDARD.defaults.palette,
      geometry: PRESENTATION_DESIGN_STANDARD.defaults.geometry,
      componentSystem: PRESENTATION_DESIGN_STANDARD.componentSystem,
      tableProfile: PRESENTATION_DESIGN_STANDARD.tableProfile,
      tableVariants: PRESENTATION_DESIGN_STANDARD.tableVariants,
      everySlideChecklist: PRESENTATION_DESIGN_STANDARD.everySlideChecklist,
      visualQaLoop: PRESENTATION_DESIGN_STANDARD.visualQaLoop,
    },
    requiredSequence: [
      "Inspect the authoritative Current render and every structured object.",
      "Diagnose hierarchy, optical text alignment, fit, visual balance, table/figure treatment, and layout compatibility; align visible text starts using insets and paragraph indents rather than shape x coordinates alone.",
      "Classify the slide into one shared layout recipe and compose it from the named components; do not invent a new spacing system for the slide.",
      "Choose a compatible approved PowerPoint layout or explain why the current native arrangement is stronger.",
      "Stage one bounded semantic recomposition or atomic geometry transaction without rewriting content.",
      "When the staged transaction fully addresses a submitted comment, pass that exact comment ID in addressedThreadIds so it leaves the clean canvas; never clear unrelated or partially addressed feedback.",
      "Materialize and render the Proposal through Microsoft PowerPoint.",
      "Compare Current and Proposal, reject regressions, and revise until the proposal is materially better or record approved-as-is evidence.",
    ],
    pauseOnlyWhen: PRESENTATION_DESIGN_STANDARD.askOnlyWhen,
    definitionOfDone: "The native Proposal render is visibly stronger than Current, every approved word and native technical object remains present, no fit/collision/safe-area regression is known, and the result stays editable in PowerPoint.",
  };
}

export function buildDeckDesignWorkOrder(input: {
  deck: DeckJob;
  projectUpdatedAt: string;
  templateCatalog: TemplateCatalog;
  currentRender?: NativeRenderResult;
  threads?: DesignThread[];
}) {
  const representatives = selectRepresentativeSlides(input.deck);
  return {
    schema: "presentation-studio/deck-design-work-order" as const,
    version: 1,
    projectUpdatedAt: input.projectUpdatedAt,
    deck: { id: input.deck.id, name: input.deck.name, slideCount: input.deck.audit?.slideCount ?? 0, sceneRevision: input.deck.scene?.revision, targetTemplateId: input.deck.targetTemplateId },
    representativeSlides: representatives,
    workOrders: representatives.map((item) => ({
      ...item,
      workOrder: buildSlideDesignWorkOrder({ ...input, slideNumber: item.slideNumber }),
    })),
    executionPolicy: "Qualify this representative set before expanding to the full deck. Routine reversible design decisions do not require per-slide questions; every staged change still enters Current/Proposal review.",
  };
}
