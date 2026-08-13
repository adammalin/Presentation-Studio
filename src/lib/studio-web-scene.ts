import type {
  DeckJob,
  SceneSemanticRole,
  StudioLayoutRecipe,
  StudioWebFrame,
  StudioWebNode,
  StudioWebScene,
  StudioWebSlide,
} from "../types";
import { STUDIO_WEB_SCENE_SCHEMA, STUDIO_WEB_SCENE_VERSION } from "../types";
import type { GeometryEditRequest, VisualDesignRequest } from "./cleanup";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import type { TemplateLayoutPreview, SlideRenderCatalog, TemplatePreviewElement } from "./template-catalog";

const EMU_PER_INCH = 914_400;
const EMU_PER_POINT = 12_700;
const STUDIO_WIDTH_INCHES = PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches;
const STUDIO_HEIGHT_INCHES = PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches;

function inches(value: number): number {
  return Math.round(value * EMU_PER_INCH);
}

function emuInches(value: number): number {
  return value / EMU_PER_INCH;
}

function points(value: number): number {
  return Math.round(value * EMU_PER_POINT);
}

function frame(x: number, y: number, width: number, height: number, rotation = 0): StudioWebFrame {
  return { x: inches(x), y: inches(y), width: inches(width), height: inches(height), rotation };
}

function sourceFrame(value: { x: number; y: number; width: number; height: number; rotation: number }): StudioWebFrame {
  return { x: value.x, y: value.y, width: value.width, height: value.height, rotation: value.rotation };
}

function scaleFrame(value: StudioWebFrame, from: { width: number; height: number }, to: { width: number; height: number }): StudioWebFrame {
  return {
    x: Math.round(value.x * to.width / from.width),
    y: Math.round(value.y * to.height / from.height),
    width: Math.round(value.width * to.width / from.width),
    height: Math.round(value.height * to.height / from.height),
    rotation: value.rotation,
  };
}

function previewElementFor(catalog: SlideRenderCatalog | undefined, slideNumber: number, shapeId: string, kind?: TemplatePreviewElement["kind"]): TemplatePreviewElement | undefined {
  const slide = catalog?.slides.find((item) => item.number === slideNumber);
  if (!slide) return undefined;
  const suffixes = [`-${shapeId}`, `-${shapeId}-text`];
  return slide.elements.find((element) => (!kind || element.kind === kind) && suffixes.some((suffix) => element.id.endsWith(suffix)))
    ?? slide.elements.find((element) => (!kind || element.kind === kind) && element.id.includes(`-${shapeId}-`));
}

function nodeKind(kind: string, hasText: boolean): StudioWebNode["kind"] {
  if (kind === "picture") return "image";
  if (kind === "table") return "table";
  if (kind === "connector") return "connector";
  if (hasText) return "text";
  if (kind === "shape") return "shape";
  return "native-object";
}

function color(value: string | undefined, fallback: string): string {
  return /^#[0-9a-f]{6}$/i.test(value ?? "") ? value!.toUpperCase() : fallback;
}

function roleStyle(role: SceneSemanticRole, preview: TemplatePreviewElement | undefined, textBox: DeckJob["audit"] extends infer _ ? NonNullable<DeckJob["audit"]>["textBoxes"][number] | undefined : never): StudioWebNode["style"] {
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  const observedSize = preview?.fontSize ?? textBox?.fontSizes.find((value) => value >= 8);
  const defaultSize = role === "title" ? 29.25 : role === "caption" || role === "label" ? 14 : 16;
  const insets = textBox?.textInsets;
  return {
    fontFamily: "Aptos",
    fontSizePt: Math.max(role === "caption" || role === "label" ? 10 : 12, Math.min(44, observedSize ?? defaultSize)),
    fontWeight: preview?.fontWeight === 700 || role === "title" ? 700 : 400,
    lineHeight: role === "title" ? 1.02 : 1.08,
    color: color(preview?.textColor, palette.darkMatter),
    background: preview?.kind === "shape" ? preview.fill : undefined,
    borderColor: preview?.stroke,
    borderWidthPt: Math.max(0, Math.min(6, (preview?.strokeWidth ?? 0) / EMU_PER_POINT)),
    textAlign: preview?.textAlign ?? (textBox?.paragraphAlignment === "center" || textBox?.paragraphAlignment === "right" ? textBox.paragraphAlignment : "left"),
    verticalAlign: preview?.verticalAlign === "center" ? "middle" : preview?.verticalAlign ?? (textBox?.verticalAlignment ?? "top"),
    paddingPt: {
      top: Math.max(0, Math.min(18, (insets?.top ?? 45_720) / EMU_PER_POINT)),
      right: Math.max(0, Math.min(18, (insets?.right ?? 91_440) / EMU_PER_POINT)),
      bottom: Math.max(0, Math.min(18, (insets?.bottom ?? 45_720) / EMU_PER_POINT)),
      left: Math.max(0, Math.min(18, (insets?.left ?? 91_440) / EMU_PER_POINT)),
    },
    objectFit: role === "image" ? "contain" : undefined,
  };
}

function compileNode(deck: DeckJob, objectId: string, studioSlideSize: { width: number; height: number }, catalog?: SlideRenderCatalog): StudioWebNode | undefined {
  const audit = deck.audit;
  const sceneObject = deck.scene?.objects.find((item) => item.id === objectId);
  const object = audit?.editableObjects.find((item) => item.id === objectId);
  if (!audit || !sceneObject || !object) return undefined;
  const textBox = audit.textBoxes.find((item) => item.slideNumber === object.slideNumber && item.shapeId === object.shapeId);
  const table = object.tableId ? audit.tables.find((item) => item.id === object.tableId) : undefined;
  const previewKind = object.kind === "picture" ? "image" : textBox ? "text" : undefined;
  const preview = previewElementFor(catalog, object.slideNumber, object.shapeId, previewKind);
  const kind = nodeKind(object.kind, Boolean(textBox?.text));
  const tableCells = table?.cells?.filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation).map((cell) => ({
    id: cell.id,
    row: cell.row,
    column: cell.column,
    rowSpan: cell.rowSpan,
    columnSpan: cell.columnSpan,
    text: cell.text,
    fill: cell.fillToken,
    semanticColorRole: cell.semanticColorRole,
  }));
  const normalizedSourceFrame = scaleFrame(sourceFrame(object.geometry), audit.slideSize, studioSlideSize);
  return {
    id: `studio-${object.id}`,
    sourceObjectId: object.id,
    sourceShapeId: object.shapeId,
    name: object.name,
    kind,
    role: sceneObject.semanticRole,
    sourceFrame: normalizedSourceFrame,
    frame: normalizedSourceFrame,
    zIndex: sceneObject.zIndex,
    visible: true,
    locked: sceneObject.protected || sceneObject.fidelityState === "unsupported-blocking" || sceneObject.fidelityState === "conversion-required",
    exactContent: Boolean(textBox?.text || table),
    text: textBox?.text,
    textHash: textBox?.textHash,
    sourceParagraphs: textBox?.paragraphs,
    tableId: table?.id,
    table: table ? { rows: table.rowCount, columns: table.columnCount, cells: tableCells ?? [] } : undefined,
    mediaPart: preview?.mediaId,
    style: roleStyle(sceneObject.semanticRole, preview, textBox),
  };
}

function normalizedVisibleText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function nodeVisibleText(node: StudioWebNode): string {
  if (node.kind === "text") return node.text ?? "";
  if (node.kind === "table" && node.table) {
    return [...node.table.cells]
      .sort((left, right) => left.row - right.row || left.column - right.column)
      .map((cell) => cell.text)
      .join(" ");
  }
  return "";
}

export function compileStudioWebScene(deck: DeckJob, catalog?: SlideRenderCatalog): StudioWebScene {
  if (!deck.audit || !deck.scene) throw new Error("Audit and compile the PowerPoint preservation scene before creating a Studio Web Scene.");
  const now = new Date().toISOString();
  const studioSlideSize = { width: inches(STUDIO_WIDTH_INCHES), height: inches(STUDIO_HEIGHT_INCHES) };
  const slides: StudioWebSlide[] = deck.scene.slides.map((slide) => {
    const sourceSlide = deck.audit!.slides.find((item) => item.number === slide.number);
    const preview = catalog?.slides.find((item) => item.number === slide.number);
    let nodes = slide.objectIds.map((objectId) => compileNode(deck, objectId, studioSlideSize, catalog)).filter((node): node is StudioWebNode => Boolean(node));
    if (!nodes.some((node) => node.role === "title")) {
      const inferredTitle = nodes
        .filter((node) => node.kind === "text" && Boolean(node.text?.trim()) && node.sourceFrame.y < inches(1.3))
        .sort((left, right) => left.sourceFrame.y - right.sourceFrame.y || right.style.fontSizePt - left.style.fontSizePt || (left.text?.length ?? 0) - (right.text?.length ?? 0))[0];
      if (inferredTitle) {
        nodes = nodes.map((node) => node.id !== inferredTitle.id ? node : {
          ...node,
          role: "title",
          style: { ...node.style, fontFamily: "Aptos", fontSizePt: Math.max(24, Math.min(32, node.style.fontSizePt)), fontWeight: 700, lineHeight: 1.02 },
        });
      }
    }
    const mappedTextNodes = nodes.filter((node) => Boolean(nodeVisibleText(node)));
    const sourceText = normalizedVisibleText(sourceSlide?.text ?? "");
    const mappedText = normalizedVisibleText(mappedTextNodes.map(nodeVisibleText).join(" "));
    const exactTextMapped = sourceText === mappedText;
    const sourceObjects = deck.scene!.objects.filter((object) => object.slideNumber === slide.number);
    return {
      id: `studio-${slide.id}`,
      slideNumber: slide.number,
      sourceSlideId: slide.id,
      sourceTextHash: slide.sourceTextHash,
      contentCoverage: {
        exactTextMapped,
        sourceCharacterCount: sourceText.length,
        mappedCharacterCount: mappedText.length,
        sourceTextBoxCount: deck.audit!.textBoxes.filter((item) => item.slideNumber === slide.number).length,
        mappedTextNodeCount: mappedTextNodes.length,
        groupedOrUnsupportedTextPresent: !exactTextMapped && (sourceText.length !== mappedText.length || sourceObjects.some((object) => ["group", "graphic-frame", "chart"].includes(object.kind))),
      },
      sourceRevision: deck.scene!.revision,
      recipe: "source",
      background: color(preview?.background, PRESENTATION_DESIGN_STANDARD.defaults.palette.polar),
      status: "imported",
      designRationale: "Faithful semantic web representation of the imported PowerPoint slide before Studio recomposition.",
      nodes,
      updatedAt: now,
    };
  });
  return {
    schema: STUDIO_WEB_SCENE_SCHEMA,
    version: STUDIO_WEB_SCENE_VERSION,
    revision: `${deck.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    deckId: deck.id,
    sourceSha256: deck.sourceSha256,
    slideSize: studioSlideSize,
    sourceSlideSize: { ...deck.audit.slideSize },
    designSystem: {
      id: "ornl-presentation-web-v1",
      standardVersion: PRESENTATION_DESIGN_STANDARD.version,
      unit: "emu",
      renderer: "html-css",
      exportTarget: "editable-powerpoint",
      compilerModes: ["source-bound-overlay", "fresh-composition"],
    },
    slides,
  };
}

function activeNodes(slide: StudioWebSlide): StudioWebNode[] {
  return slide.nodes.filter((node) => node.visible && !node.locked && !["shape", "connector"].includes(node.kind));
}

function footerNode(node: StudioWebNode): boolean {
  return node.sourceFrame.y >= inches(6.78) || node.component?.role === "footer-logo" || node.component?.role === "footer-meta";
}

function meaningfulImage(node: StudioWebNode): boolean {
  return node.kind === "image" && !footerNode(node) && node.sourceFrame.width * node.sourceFrame.height >= inches(.45) * inches(.35);
}

function styleForDesignedNode(node: StudioWebNode): StudioWebNode["style"] {
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  if (node.role === "title") return { ...node.style, fontFamily: "Aptos", fontSizePt: 29.25, fontWeight: 700, lineHeight: 1.02, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } };
  if (node.role === "caption" || node.role === "label") return { ...node.style, fontFamily: "Aptos", fontSizePt: 14, fontWeight: 400, lineHeight: 1.08, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } };
  if (node.kind === "table") return { ...node.style, fontFamily: "Aptos", fontSizePt: 16, fontWeight: 400, lineHeight: 1.05, color: palette.darkMatter, background: palette.polar, borderColor: palette.graphite, borderWidthPt: .75, textAlign: "left", verticalAlign: "middle", paddingPt: { top: 4, right: 6, bottom: 4, left: 6 } };
  return { ...node.style, fontFamily: "Aptos", fontSizePt: Math.max(16, Math.min(22, node.style.fontSizePt)), fontWeight: node.style.fontWeight === 700 ? 600 : node.style.fontWeight, lineHeight: 1.08, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 }, objectFit: node.kind === "image" ? "contain" : node.style.objectFit };
}

function styleForComponent(node: StudioWebNode): StudioWebNode["style"] {
  const base = styleForDesignedNode(node);
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  if (node.component?.role === "eyebrow") return { ...base, fontSizePt: 10.5, fontWeight: 700, lineHeight: 1, color: palette.ornlGreen, textAlign: "left" };
  if (node.component?.role === "card-kicker") return { ...base, fontSizePt: 18, fontWeight: 700, lineHeight: 1, color: [palette.ornlGreen, palette.infinity, palette.hydro, palette.darkMatter][node.component.ordinal ?? 0] ?? palette.ornlGreen };
  if (node.component?.role === "card-heading") return { ...base, fontSizePt: 13.5, fontWeight: 400, lineHeight: 1.05, color: "#666B68" };
  if (node.component?.role === "card-body") return { ...base, fontSizePt: 16, fontWeight: 400, lineHeight: 1.13, color: palette.darkMatter };
  if (node.component?.role === "footer-logo") return { ...base, paddingPt: { top: 0, right: 0, bottom: 0, left: 0 }, objectFit: "contain" };
  if (node.component?.role === "footer-meta") return { ...base, fontSizePt: 9, fontWeight: 400, lineHeight: 1, color: "#6B716E", textAlign: "right", verticalAlign: "middle" };
  return base;
}

function contained(node: StudioWebNode, target: StudioWebFrame): StudioWebFrame {
  if (node.kind !== "image" || node.sourceFrame.width <= 0 || node.sourceFrame.height <= 0) return { ...target, rotation: 0 };
  const scale = Math.min(target.width / node.sourceFrame.width, target.height / node.sourceFrame.height);
  const width = Math.max(inches(.1), Math.round(node.sourceFrame.width * scale));
  const height = Math.max(inches(.1), Math.round(node.sourceFrame.height * scale));
  return { x: Math.round(target.x + (target.width - width) / 2), y: Math.round(target.y + (target.height - height) / 2), width, height, rotation: 0 };
}

function stack(nodes: StudioWebNode[], target: StudioWebFrame, gapPt = 18): Map<string, StudioWebFrame> {
  const placements = new Map<string, StudioWebFrame>();
  if (nodes.length === 0) return placements;
  const minimumHeight = inches(.1);
  const gap = nodes.length <= 1 ? 0 : Math.max(0, Math.min(points(gapPt), (target.height - minimumHeight * nodes.length) / (nodes.length - 1)));
  const available = Math.max(minimumHeight * nodes.length, target.height - gap * Math.max(0, nodes.length - 1));
  const weights = nodes.map((node) => Math.max(inches(.35), node.sourceFrame.height));
  const total = weights.reduce((sum, value) => sum + value, 0);
  const rawHeights = weights.map((weight) => available * weight / total);
  const fixed = rawHeights.map((height) => height < minimumHeight);
  const fixedTotal = fixed.filter(Boolean).length * minimumHeight;
  const flexibleWeight = weights.reduce((sum, weight, index) => sum + (fixed[index] ? 0 : weight), 0);
  const heights = rawHeights.map((height, index) => fixed[index] ? minimumHeight : Math.max(minimumHeight, (available - fixedTotal) * weights[index] / Math.max(1, flexibleWeight)));
  let y = target.y;
  nodes.forEach((node, index) => {
    const remainingNodes = nodes.length - index;
    const remainingBottom = target.y + target.height;
    const height = index === nodes.length - 1 ? Math.max(minimumHeight, remainingBottom - y) : Math.max(minimumHeight, Math.round(heights[index]));
    const next = { x: target.x, y, width: target.width, height, rotation: 0 };
    placements.set(node.id, node.kind === "image" ? contained(node, next) : next);
    y += height + (remainingNodes > 1 ? gap : 0);
  });
  return placements;
}

function grid(nodes: StudioWebNode[], target: StudioWebFrame, columns = 2, gapPt = 18): Map<string, StudioWebFrame> {
  const placements = new Map<string, StudioWebFrame>();
  if (nodes.length === 0) return placements;
  const gap = points(gapPt);
  const rows = Math.ceil(nodes.length / columns);
  const width = Math.max(inches(.25), (target.width - gap * (columns - 1)) / columns);
  const height = Math.max(inches(.25), (target.height - gap * (rows - 1)) / rows);
  nodes.forEach((node, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const next = { x: Math.round(target.x + column * (width + gap)), y: Math.round(target.y + row * (height + gap)), width: Math.round(width), height: Math.round(height), rotation: 0 };
    placements.set(node.id, node.kind === "image" ? contained(node, next) : next);
  });
  return placements;
}

export function recommendedStudioRecipe(slide: StudioWebSlide): StudioLayoutRecipe {
  const nodes = activeNodes(slide);
  if (nodes.some((node) => node.kind === "table")) return "ornl-title-table";
  const bodyCount = nodes.filter((node) => node.kind === "text" && node.role === "body" && !footerNode(node)).length;
  const labelCount = nodes.filter((node) => node.kind === "text" && node.role === "label" && !footerNode(node)).length;
  if (bodyCount >= 3 && bodyCount <= 6 && labelCount >= bodyCount) return "ornl-title-card-grid";
  const images = nodes.filter(meaningfulImage).length;
  if (images >= 2) return "ornl-title-figure-grid";
  if (images === 1) return "ornl-title-two-column";
  return "ornl-title-content";
}

function cardFrame(ordinal: number, count: number): StudioWebFrame {
  const columns = count <= 2 ? Math.max(1, count) : 2;
  const rows = Math.ceil(count / columns);
  const gapX = .24;
  const gapY = .24;
  const region = { x: .47, y: 1.36, width: 12.39, height: rows === 1 ? 2.42 : 5.05 };
  const width = (region.width - gapX * (columns - 1)) / columns;
  const height = (region.height - gapY * (rows - 1)) / rows;
  const column = ordinal % columns;
  const row = Math.floor(ordinal / columns);
  return frame(region.x + column * (width + gapX), region.y + row * (height + gapY), width, height);
}

export interface StudioGeneratedComponent {
  id: string;
  kind: "rect" | "line";
  frame: StudioWebFrame;
  fillColor?: string;
  lineColor?: string;
  lineWidthPt: number;
  behindContent: boolean;
}

export function studioGeneratedComponents(slide: StudioWebSlide): StudioGeneratedComponent[] {
  if (slide.recipe === "source" || slide.recipe === "template-layout") return [];
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  const hasEyebrow = slide.nodes.some((node) => node.component?.role === "eyebrow");
  const components: StudioGeneratedComponent[] = [{ id: `studio-title-rule-${slide.slideNumber}`, kind: "rect", frame: frame(.47, hasEyebrow ? 1.10 : .93, hasEyebrow ? .62 : .96, .035), fillColor: palette.ornlGreen, lineWidthPt: 0, behindContent: true }];
  if (slide.recipe !== "ornl-title-card-grid") return components;
  const groups = [...new Set(slide.nodes.filter((node) => node.component?.role === "card-body").map((node) => node.component!.groupId))];
  const accents = [palette.ornlGreen, palette.aqua, palette.infinity, palette.forge, palette.plasma, palette.pulsar];
  groups.forEach((groupId, ordinal) => {
    const card = cardFrame(ordinal, groups.length);
    components.push({ id: `${groupId}-surface`, kind: "rect", frame: card, fillColor: palette.polar, lineColor: palette.graphite, lineWidthPt: .75, behindContent: true });
    components.push({ id: `${groupId}-accent`, kind: "rect", frame: { ...card, height: points(2) }, fillColor: accents[ordinal] ?? palette.ornlGreen, lineWidthPt: 0, behindContent: true });
  });
  components.push({ id: `studio-footer-rule-${slide.slideNumber}`, kind: "rect", frame: frame(.47, 6.92, 12.39, .012), fillColor: palette.graphite, lineWidthPt: 0, behindContent: true });
  return components;
}

function templatePlacements(nodes: StudioWebNode[], layout: TemplateLayoutPreview): Map<string, StudioWebFrame> {
  const placements = new Map<string, StudioWebFrame>();
  const slots = layout.semantic?.slots.filter((slot) => !["footer", "date", "slide-number"].includes(slot.role)) ?? [];
  const used = new Set<string>();
  const contentKind = (node: StudioWebNode) => node.kind === "image" ? "image" : node.kind === "table" ? "table" : node.role === "chart" ? "chart" : "text";
  const ordered = [...nodes].sort((left, right) => Number(right.role === "title") - Number(left.role === "title") || left.zIndex - right.zIndex);
  for (const node of ordered) {
    const kind = contentKind(node);
    const slot = slots.find((candidate) => !used.has(candidate.id) && (candidate.role === node.role || candidate.acceptedContent.includes(kind)))
      ?? slots.find((candidate) => !used.has(candidate.id) && candidate.acceptedContent.includes(kind));
    if (!slot) continue;
    used.add(slot.id);
    const next = { x: slot.x, y: slot.y, width: slot.width, height: slot.height, rotation: 0 };
    placements.set(node.id, node.kind === "image" ? contained(node, next) : next);
  }
  return placements;
}

export function recomposeStudioWebSlide(scene: StudioWebScene, slideNumber: number, requestedRecipe?: StudioLayoutRecipe, layout?: TemplateLayoutPreview, rationale?: string): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  const recipe = requestedRecipe ?? recommendedStudioRecipe(slide);
  if (recipe === "template-layout" && !layout?.semantic) throw new Error("Choose an installed template layout with semantic regions before applying template-layout.");
  const nodes = activeNodes(slide);
  const title = nodes.find((node) => node.role === "title");
  const footer = nodes.filter(footerNode);
  const eyebrow = nodes.find((node) => node.kind === "text" && node.role === "label" && !footerNode(node) && (!title || node.sourceFrame.y < title.sourceFrame.y));
  const content = nodes.filter((node) => node.id !== title?.id && node.id !== eyebrow?.id && !footerNode(node) && node.role !== "caption" && node.role !== "label");
  const captions = nodes.filter((node) => node.id !== eyebrow?.id && !footerNode(node) && (node.role === "caption" || node.role === "label"));
  const placements = new Map<string, StudioWebFrame>();
  const components = new Map<string, StudioWebNode["component"]>();
  const titleFrame = frame(.47, eyebrow ? .58 : .29, 12.39, eyebrow ? .50 : .68);
  if (title) placements.set(title.id, titleFrame);
  if (eyebrow) {
    placements.set(eyebrow.id, frame(.47, .32, 12.39, .18));
    components.set(eyebrow.id, { groupId: `studio-header-${slideNumber}`, role: "eyebrow" });
  }
  if (recipe === "source") for (const node of nodes) placements.set(node.id, { ...node.sourceFrame });
  else if (recipe === "template-layout" && layout) for (const [id, value] of templatePlacements(nodes, layout)) placements.set(id, value);
  else if (recipe === "ornl-title-table") {
    const tables = content.filter((node) => node.kind === "table");
    const remaining = content.filter((node) => node.kind !== "table");
    const support = [...remaining, ...captions];
    if (tables.length === 1 && support.length >= 2) {
      placements.set(tables[0].id, frame(.47, 1.15, 8.05, 5.30));
      for (const [id, value] of stack(support, frame(8.90, 1.15, 3.96, 5.30), 12)) placements.set(id, value);
    } else {
      for (const [id, value] of stack(tables, frame(.47, 1.15, 12.39, support.length ? 4.68 : 5.30), 12)) placements.set(id, value);
      for (const [id, value] of stack(support, frame(.47, 5.98, 12.39, .64), 8)) placements.set(id, value);
    }
  } else if (recipe === "ornl-title-two-column") {
    const visual = content.find((node) => meaningfulImage(node) || node.kind === "table");
    const left = content.filter((node) => node.id !== visual?.id);
    for (const [id, value] of stack(left, frame(.47, 1.15, 5.78, 5.40), 18)) placements.set(id, value);
    if (visual) placements.set(visual.id, contained(visual, frame(6.63, 1.15, 6.23, 5.40)));
    for (const [id, value] of stack(captions, frame(6.63, 6.58, 6.23, .18), 4)) placements.set(id, value);
  } else if (recipe === "ornl-title-figure-grid") {
    const visuals = content.filter(meaningfulImage);
    const remaining = content.filter((node) => node.kind !== "image");
    for (const [id, value] of grid(visuals, frame(.47, 1.15, 12.39, 4.70), visuals.length <= 2 ? 2 : 3, 18)) placements.set(id, value);
    for (const [id, value] of stack([...remaining, ...captions], frame(.47, 5.98, 12.39, .64), 8)) placements.set(id, value);
  } else if (recipe === "ornl-title-card-grid") {
    const bodies = content.filter((node) => node.kind === "text" && node.role === "body").sort((left, right) => left.zIndex - right.zIndex);
    const groupNodeIds = new Set<string>();
    let priorBodyZ = title?.zIndex ?? -Infinity;
    bodies.forEach((body, ordinal) => {
      const groupId = `studio-card-${slideNumber}-${ordinal + 1}`;
      const card = cardFrame(ordinal, bodies.length);
      const labels = captions.filter((node) => node.kind === "text" && node.zIndex > priorBodyZ && node.zIndex < body.zIndex).sort((left, right) => left.zIndex - right.zIndex);
      const kicker = labels[0];
      const heading = labels.at(-1);
      if (kicker) {
        placements.set(kicker.id, frame(emuInches(card.x) + .27, emuInches(card.y) + .27, heading && heading.id !== kicker.id ? .75 : emuInches(card.width) - .54, .30));
        components.set(kicker.id, { groupId, role: "card-kicker", ordinal });
        groupNodeIds.add(kicker.id);
      }
      if (heading && heading.id !== kicker?.id) {
        placements.set(heading.id, frame(emuInches(card.x) + 1.04, emuInches(card.y) + .29, emuInches(card.width) - 1.31, .28));
        components.set(heading.id, { groupId, role: "card-heading", ordinal });
        groupNodeIds.add(heading.id);
      }
      placements.set(body.id, frame(emuInches(card.x) + .27, emuInches(card.y) + .72, emuInches(card.width) - .49, emuInches(card.height) - .88));
      components.set(body.id, { groupId, role: "card-body", ordinal });
      groupNodeIds.add(body.id);
      priorBodyZ = body.zIndex;
    });
    const extras = [...content, ...captions].filter((node) => !groupNodeIds.has(node.id));
    for (const [id, value] of stack(extras, frame(.47, 5.96, 12.39, .52), 6)) placements.set(id, value);
  } else {
    for (const [id, value] of stack([...content, ...captions], frame(.47, 1.15, 12.39, 5.57), 18)) placements.set(id, value);
  }
  const footerLogo = footer.find((node) => node.kind === "image");
  const footerMeta = footer.find((node) => node.kind === "text");
  if (recipe !== "source" && recipe !== "template-layout") {
    if (footerLogo) {
      placements.set(footerLogo.id, contained(footerLogo, frame(.47, 7.06, .76, .20)));
      components.set(footerLogo.id, { groupId: `studio-footer-${slideNumber}`, role: "footer-logo" });
    }
    if (footerMeta) {
      placements.set(footerMeta.id, frame(9.01, 7.04, 3.89, .20));
      components.set(footerMeta.id, { groupId: `studio-footer-${slideNumber}`, role: "footer-meta" });
    }
  }
  const now = new Date().toISOString();
  const nextSlide: StudioWebSlide = {
    ...slide,
    recipe,
    targetLayoutId: recipe === "template-layout" ? layout?.id : undefined,
    targetLayoutName: recipe === "template-layout" ? layout?.name : undefined,
    status: recipe === "source" ? "imported" : "designed",
    designRationale: (rationale ?? `Recompose exact source content with the shared ${recipe} ORNL web component recipe.`).trim().slice(0, 1_000),
    nodes: slide.nodes.map((node) => {
      const component = recipe === "source" || recipe === "template-layout" ? undefined : components.get(node.id);
      if (!placements.has(node.id)) return { ...node, component };
      const nextNode = { ...node, component };
      return { ...nextNode, frame: placements.get(node.id)!, style: component ? styleForComponent(nextNode) : styleForDesignedNode(nextNode) };
    }),
    updatedAt: now,
  };
  return { ...scene, revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`, slides: scene.slides.map((item) => item.slideNumber === slideNumber ? nextSlide : item) };
}

export function updateStudioWebNodeFrame(scene: StudioWebScene, slideNumber: number, nodeId: string, nextFrame: StudioWebFrame): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const node = slide?.nodes.find((item) => item.id === nodeId);
  if (!slide || !node) throw new Error("The requested Studio node is not present in the current slide revision.");
  if (node.locked) throw new Error(`${node.name} is locked because its PowerPoint representation is not safely editable.`);
  const width = Math.max(inches(.1), Math.min(scene.slideSize.width, Math.round(nextFrame.width)));
  const height = Math.max(inches(.1), Math.min(scene.slideSize.height, Math.round(nextFrame.height)));
  const bounded = { ...nextFrame, width, height, x: Math.max(0, Math.min(scene.slideSize.width - width, Math.round(nextFrame.x))), y: Math.max(0, Math.min(scene.slideSize.height - height, Math.round(nextFrame.y))), rotation: Math.max(-360, Math.min(360, nextFrame.rotation)) };
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : { ...item, status: "designed", updatedAt: now, designRationale: "Human-adjusted on the shared Studio web canvas.", nodes: item.nodes.map((candidate) => candidate.id === nodeId ? { ...candidate, frame: bounded } : candidate) }),
  };
}

export function updateStudioWebNodeStyle(scene: StudioWebScene, slideNumber: number, nodeId: string, patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">>): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const node = slide?.nodes.find((item) => item.id === nodeId);
  if (!slide || !node) throw new Error("The requested Studio node is not present in the current slide revision.");
  if (node.locked) throw new Error(`${node.name} is locked because its PowerPoint representation is not safely editable.`);
  if (patch.fontSizePt !== undefined && (!Number.isFinite(patch.fontSizePt) || patch.fontSizePt < 10 || patch.fontSizePt > 60)) throw new Error("Studio text size must be between 10 and 60 pt.");
  if (patch.fontWeight !== undefined && ![400, 600, 700].includes(patch.fontWeight)) throw new Error("Studio font weight must be 400, 600, or 700.");
  if (patch.color !== undefined && !/^#[0-9a-f]{6}$/i.test(patch.color)) throw new Error("Studio text color must be a six-digit hex color.");
  if (patch.textAlign !== undefined && !["left", "center", "right"].includes(patch.textAlign)) throw new Error("Studio text alignment must be left, center, or right.");
  if (patch.verticalAlign !== undefined && !["top", "middle", "bottom"].includes(patch.verticalAlign)) throw new Error("Studio vertical alignment must be top, middle, or bottom.");
  if (patch.objectFit !== undefined && !["contain", "cover"].includes(patch.objectFit)) throw new Error("Studio media fit must be contain or cover.");
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      status: "designed",
      updatedAt: now,
      designRationale: "Human- or AI-refined with the shared Studio web design controls.",
      nodes: item.nodes.map((candidate) => candidate.id === nodeId ? { ...candidate, style: { ...candidate.style, ...patch, fontFamily: "Aptos" } } : candidate),
    }),
  };
}

function frameChanged(left: StudioWebFrame, right: StudioWebFrame): boolean {
  return Math.max(Math.abs(left.x - right.x), Math.abs(left.y - right.y), Math.abs(left.width - right.width), Math.abs(left.height - right.height)) > 1_000;
}

export function studioGeometryRequests(deck: DeckJob, scene: StudioWebScene, slideNumber: number, author: "human" | "ai" = "ai"): GeometryEditRequest[] {
  if (scene.deckId !== deck.id || scene.sourceSha256 !== deck.sourceSha256) throw new Error("The Studio Web Scene is stale for this deck.");
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  return slide.nodes.filter((node) => node.visible && !node.locked && frameChanged(node.sourceFrame, node.frame)).map((node) => ({
    objectId: node.sourceObjectId,
    target: scaleFrame(node.frame, scene.slideSize, scene.sourceSlideSize),
    rationale: `${slide.designRationale} Compile the web-computed frame for ${node.name} back to its editable PowerPoint object.`.slice(0, 700),
    author,
    constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false },
  }));
}

export function studioVisualDesignRequest(scene: StudioWebScene, slideNumber: number, author: "human" | "ai" = "ai"): VisualDesignRequest {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  const textStyles = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "text").map((node) => ({
    objectId: node.sourceObjectId,
    fontSizePt: node.style.fontSizePt,
    bold: node.style.fontWeight >= 600,
    color: color(node.style.color, PRESENTATION_DESIGN_STANDARD.defaults.palette.darkMatter),
    alignment: node.style.textAlign,
    verticalAlignment: node.style.verticalAlign,
    insetsInches: {
      top: Math.min(.25, node.style.paddingPt.top / 72),
      right: Math.min(.25, node.style.paddingPt.right / 72),
      bottom: Math.min(.25, node.style.paddingPt.bottom / 72),
      left: Math.min(.25, node.style.paddingPt.left / 72),
    },
    paragraphStyle: { lineSpacingMultiple: node.style.lineHeight },
    rationale: `${slide.designRationale} Apply the Studio web typography for ${node.name}.`.slice(0, 700),
    author,
  }));
  const decorations: VisualDesignRequest["decorations"] = studioGeneratedComponents(slide).map((component) => ({
    id: component.id,
    name: component.id.includes("title-rule") ? "ORNL title rule" : component.id.includes("footer-rule") ? "ORNL footer rule" : component.id.endsWith("-accent") ? "ORNL card accent" : "ORNL content card",
    geometry: scaleFrame(component.frame, scene.slideSize, scene.sourceSlideSize),
    fillColor: component.fillColor,
    lineColor: component.lineColor,
    lineWidthPt: component.lineWidthPt,
    behindContent: component.behindContent,
    rationale: "Use the shared square-cornered ORNL web component rather than a slide-specific decoration.",
    author,
  }));
  return { slideNumber, textStyles, decorations };
}

export function studioSceneNeedsRebuild(deck: DeckJob): boolean {
  return Boolean(deck.audit && deck.scene) && (!deck.studioScene || deck.studioScene.schema !== STUDIO_WEB_SCENE_SCHEMA || deck.studioScene.version !== STUDIO_WEB_SCENE_VERSION || deck.studioScene.deckId !== deck.id || deck.studioScene.sourceSha256 !== deck.sourceSha256);
}
