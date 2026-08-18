import type {
  DeckJob,
  SceneSemanticRole,
  StudioConnectorDesign,
  StudioFigureTreatment,
  StudioLayoutRecipe,
  StudioDeckRhythm,
  StudioTableCellDesign,
  StudioTableDesign,
  StudioWebFrame,
  StudioWebNode,
  StudioWebScene,
  StudioWebSlide,
} from "../types";
import { STUDIO_WEB_SCENE_SCHEMA, STUDIO_WEB_SCENE_VERSION } from "../types";
import type { GeometryEditRequest, VisualDesignRequest } from "./cleanup";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import type { TemplateLayoutPreview, SlideRenderCatalog, TemplatePreviewElement } from "./template-catalog";
import { contentCharacterSignature } from "./content-integrity";

const EMU_PER_INCH = 914_400;
const EMU_PER_POINT = 12_700;
const STUDIO_WIDTH_INCHES = PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches;
const STUDIO_HEIGHT_INCHES = PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches;

export function defaultStudioDeckRhythm(): StudioDeckRhythm {
  const spacing = PRESENTATION_DESIGN_STANDARD.componentSystem.spacing;
  return {
    safeMarginPt: PRESENTATION_DESIGN_STANDARD.defaults.geometry.safeMarginPt,
    gridPt: 6,
    compactGapPt: spacing.compactPt,
    normalGapPt: spacing.normalPt,
    primaryGapPt: spacing.primarySeparationPt,
    captionGapPt: 8,
    titleContentGapPt: spacing.primarySeparationPt,
  };
}

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
  return slide.elements.find((element) => (!kind || element.kind === kind) && element.sourceShapeId === shapeId)
    ?? slide.elements.find((element) => (!kind || element.kind === kind) && suffixes.some((suffix) => element.id.endsWith(suffix)))
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

function templateTextColor(layout: TemplateLayoutPreview | undefined, node: StudioWebNode, nextFrame: StudioWebFrame): string | undefined {
  if (!layout || node.kind !== "text") return undefined;
  const candidates = layout.elements.filter((element) => element.placeholderType && element.textColor);
  const direct = candidates.find((element) => element.x === nextFrame.x && element.y === nextFrame.y && element.width === nextFrame.width && element.height === nextFrame.height);
  if (direct?.textColor) return color(direct.textColor, node.style.color);
  const roleTypes = node.role === "title" ? new Set(["title", "ctrTitle"]) : node.role === "caption" || node.role === "label" ? new Set(["body", "subTitle"]) : new Set(["body", "obj"]);
  return color(candidates.find((element) => roleTypes.has(element.placeholderType ?? ""))?.textColor, node.style.color);
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
  if (preview && inheritedPlaceholder(preview)) return undefined;
  const kind = nodeKind(object.kind, Boolean(textBox?.text));
  const tableCells = table?.cells?.filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation).map((cell) => ({
    id: cell.id,
    row: cell.row,
    column: cell.column,
    rowSpan: cell.rowSpan,
    columnSpan: cell.columnSpan,
    text: cell.text,
    textRuns: cell.textRuns,
    paragraphRunCounts: cell.paragraphRunCounts,
    runBreaksBefore: cell.runBreaksBefore,
    fill: cell.fillToken,
    semanticColorRole: cell.semanticColorRole,
  }));
  const normalizedSourceFrame = scaleFrame(sourceFrame(object.geometry), audit.slideSize, studioSlideSize);
  const style = roleStyle(sceneObject.semanticRole, preview, textBox);
  if (table) {
    const dense = table.columnCount >= 6 || table.rowCount >= 9 || table.totalCellCharacterCount >= 900 || table.maximumCellCharacterCount >= 180;
    style.fontSizePt = dense ? PRESENTATION_DESIGN_STANDARD.tableVariants.denseTechnical.bodyFontSizePt : PRESENTATION_DESIGN_STANDARD.tableVariants.standard.bodyFontSizePt;
  }
  return {
    id: `studio-${object.id}`,
    sourceObjectId: object.id,
    sourceShapeId: object.shapeId,
    sourceBinding: "editable-object",
    name: object.name,
    kind,
    role: sceneObject.semanticRole,
    sourceFrame: normalizedSourceFrame,
    frame: normalizedSourceFrame,
    zIndex: sceneObject.zIndex,
    sourceTextOrder: 0,
    visible: true,
    locked: sceneObject.protected || sceneObject.fidelityState === "unsupported-blocking" || sceneObject.fidelityState === "conversion-required",
    exactContent: Boolean(textBox?.text || table),
    text: textBox?.text,
    textHash: textBox?.textHash,
    sourceParagraphs: textBox?.paragraphs,
    tableId: table?.id,
    table: table ? {
      rows: table.rowCount,
      columns: table.columnCount,
      cells: tableCells ?? [],
      design: {
        headerRows: table.rowCount > 0 ? 1 : 0,
        columnWidths: Array.from({ length: Math.max(1, table.columnCount) }, () => 1 / Math.max(1, table.columnCount)),
        rowHeights: Array.from({ length: Math.max(1, table.rowCount) }, () => 1 / Math.max(1, table.rowCount)),
        borderMode: "subtle",
        borderColor: "#DBDCDB",
        borderWidthPt: .75,
        defaultPaddingPt: { top: 4, right: 7, bottom: 4, left: 7 },
        cellStyles: [],
      },
    } : undefined,
    mediaPart: preview?.mediaId,
    opticalInsets: textBox ? {
      left: Math.max(0, Math.round(textBox.opticalLeftOffsetEmu * studioSlideSize.width / audit.slideSize.width)),
      top: Math.max(0, Math.round(textBox.textInsets.top * studioSlideSize.height / audit.slideSize.height)),
      right: Math.max(0, Math.round(textBox.textInsets.right * studioSlideSize.width / audit.slideSize.width)),
      bottom: Math.max(0, Math.round(textBox.textInsets.bottom * studioSlideSize.height / audit.slideSize.height)),
      authority: "source-estimate",
      basis: "rendered-text",
    } : { left: 0, top: 0, right: 0, bottom: 0, authority: "scene-frame", basis: kind === "image" ? "active-image-content" : "shape" },
    style,
  };
}

function roleForCatalogElement(element: TemplatePreviewElement, slideSize: { width: number; height: number }): SceneSemanticRole {
  if (element.kind === "image") return "image";
  if (["title", "ctrTitle"].includes(element.placeholderType ?? "")) return "title";
  const normalizedName = element.name.toLowerCase();
  if (element.y < slideSize.height * .2 && (element.fontSize ?? 0) >= 20) return "title";
  if ((element.text?.length ?? 0) <= 42 && (normalizedName.includes("label") || normalizedName.includes("placeholder"))) return "label";
  return "body";
}

function compileCatalogDerivedNodes(deck: DeckJob, slideNumber: number, studioSlideSize: { width: number; height: number }, nodes: StudioWebNode[], catalog?: SlideRenderCatalog): StudioWebNode[] {
  const previewSlide = catalog?.slides.find((item) => item.number === slideNumber);
  if (!previewSlide || !catalog) return [];
  const mappedShapeIds = new Set(nodes.map((node) => node.sourceShapeId));
  return previewSlide.elements.flatMap((element, index): StudioWebNode[] => {
    if (element.origin !== "slide" || inheritedPlaceholder(element) || !element.sourceShapeId || mappedShapeIds.has(element.sourceShapeId) || !["text", "image"].includes(element.kind)) return [];
    if (element.kind === "text" && !element.text?.trim()) return [];
    if (element.kind === "image" && !element.mediaId) return [];
    const role = roleForCatalogElement(element, { width: catalog.slideWidth, height: catalog.slideHeight });
    const normalizedFrame = scaleFrame(sourceFrame(element), { width: catalog.slideWidth, height: catalog.slideHeight }, studioSlideSize);
    const text = element.kind === "text" ? normalizedVisibleText(element.text ?? "") : undefined;
    return [{
      id: `studio-catalog-${slideNumber}-${element.sourceShapeId}-${index}`,
      sourceObjectId: `catalog:${previewSlide.sourcePart}:${element.sourceShapeId}:${index}`,
      sourceShapeId: element.sourceShapeId,
      sourceBinding: "catalog-derived",
      name: element.name,
      kind: element.kind,
      role,
      sourceFrame: normalizedFrame,
      frame: normalizedFrame,
      zIndex: index,
      sourceTextOrder: 0,
      visible: true,
      locked: false,
      exactContent: true,
      text,
      textHash: element.textHash,
      sourceParagraphs: element.sourceParagraphs,
      mediaPart: element.mediaId,
      style: roleStyle(role, element, undefined),
    }];
  });
}

function normalizedVisibleText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function inheritedPlaceholder(element: TemplatePreviewElement): boolean {
  return ["sldNum", "dt", "ftr", "hdr"].includes(element.placeholderType ?? "");
}

function catalogSlideContentValues(preview: SlideRenderCatalog["slides"][number] | undefined): string[] | undefined {
  if (!preview) return undefined;
  return preview.elements
    .filter((element) => element.origin === "slide" && element.kind === "text" && !inheritedPlaceholder(element) && Boolean(element.text?.trim()))
    .map((element) => element.text ?? "");
}

function assignSourceTextOrder(nodes: StudioWebNode[], sourceText: string): StudioWebNode[] {
  const normalizedSource = normalizedVisibleText(sourceText);
  const occurrences = new Map<string, number>();
  return nodes.map((node) => {
    const text = normalizedVisibleText(nodeVisibleText(node));
    if (!text) return { ...node, sourceTextOrder: normalizedSource.length + 10_000 + node.zIndex };
    const start = occurrences.get(text) ?? 0;
    const found = normalizedSource.indexOf(text, start);
    const fallback = normalizedSource.indexOf(text);
    const sourceTextOrder = found >= 0 ? found : fallback >= 0 ? fallback : normalizedSource.length + node.zIndex;
    occurrences.set(text, sourceTextOrder + text.length);
    return { ...node, sourceTextOrder };
  });
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

function nodeVisibleTextValues(node: StudioWebNode): string[] {
  if (node.kind === "text") return node.text ? [node.text] : [];
  if (node.kind === "table" && node.table) return [...node.table.cells]
    .sort((left, right) => left.row - right.row || left.column - right.column)
    .map((cell) => cell.text);
  return [];
}

export function studioSlideContentSignature(slide: StudioWebSlide): string {
  const visible = slide.nodes.filter((node) => node.visible);
  const count = (kind: StudioWebNode["kind"]) => visible.filter((node) => node.kind === kind).length;
  const textCharacters = visible.reduce((sum, node) => sum + nodeVisibleText(node).length, 0);
  const density = textCharacters > 1_100 ? "very-dense" : textCharacters > 700 ? "dense" : textCharacters > 300 ? "moderate" : "light";
  const paragraphAtoms = visible.filter((node) => node.sourceBinding === "semantic-atom").length;
  return `text:${count("text")};atoms:${paragraphAtoms};images:${count("image")};tables:${count("table")};native:${count("native-object")};connectors:${count("connector")};density:${density}`;
}

export function compileStudioWebScene(deck: DeckJob, catalog?: SlideRenderCatalog): StudioWebScene {
  if (!deck.audit || !deck.scene) throw new Error("Audit and compile the PowerPoint preservation scene before creating a Studio Web Scene.");
  const now = new Date().toISOString();
  const studioSlideSize = { width: inches(STUDIO_WIDTH_INCHES), height: inches(STUDIO_HEIGHT_INCHES) };
  const slides: StudioWebSlide[] = deck.scene.slides.map((slide) => {
    const sourceSlide = deck.audit!.slides.find((item) => item.number === slide.number);
    const preview = catalog?.slides.find((item) => item.number === slide.number);
    let nodes = slide.objectIds.map((objectId) => compileNode(deck, objectId, studioSlideSize, catalog)).filter((node): node is StudioWebNode => Boolean(node));
    nodes = [...nodes, ...compileCatalogDerivedNodes(deck, slide.number, studioSlideSize, nodes, catalog)];
    nodes = assignSourceTextOrder(nodes, sourceSlide?.text ?? "");
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
    // The render catalog distinguishes authored slide content from inherited
    // master/layout furniture such as automatic slide numbers. Raw slide XML
    // does not, and its run boundaries can split a visible token (for example,
    // `20` + `24`). Use the catalog inventory as the source-content authority
    // whenever it is available, while retaining the audit text as a fallback.
    const sourceTextValues = catalogSlideContentValues(preview) ?? [sourceSlide?.text ?? ""];
    const mappedTextValues = mappedTextNodes.flatMap(nodeVisibleTextValues);
    const sourceText = normalizedVisibleText(sourceTextValues.join(" "));
    const mappedText = normalizedVisibleText(mappedTextValues.join(" "));
    const sourceContentSignature = contentCharacterSignature(sourceTextValues);
    const exactTextMapped = sourceContentSignature === contentCharacterSignature(mappedTextValues);
    const sourceObjects = deck.scene!.objects.filter((object) => object.slideNumber === slide.number);
    return {
      id: `studio-${slide.id}`,
      slideNumber: slide.number,
      sourceSlideId: slide.id,
      sourceTextHash: slide.sourceTextHash,
      contentCoverage: {
        exactTextMapped,
        sourceContentSignature,
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
      figureTreatments: [],
      conceptReferences: [],
      visualNeeds: [],
      constraints: [],
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
    rhythm: defaultStudioDeckRhythm(),
    designMemory: [],
    componentLibrary: [],
    tableLibrary: [],
    tableContinuationPlans: [],
    designSystem: {
      id: deck.targetTemplateDecisionSource === "automatic-source-preservation" ? "source-template-preservation-web-v1" : "ornl-presentation-web-v1",
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
  if (node.component?.role === "footer-logo" || node.component?.role === "footer-meta") return true;
  if (node.sourceFrame.y < inches(6.78)) return false;
  const normalizedName = node.name.toLowerCase();
  if (normalizedName.includes("footer") || normalizedName.includes("slide number")) return true;
  if (node.kind === "text") return node.style.fontSizePt <= 10.5;
  if (node.kind === "image") return node.sourceFrame.x <= inches(1.5) && node.sourceFrame.width >= inches(.55);
  return false;
}

function meaningfulImage(node: StudioWebNode): boolean {
  return node.kind === "image" && !footerNode(node) && node.sourceFrame.width * node.sourceFrame.height >= inches(.45) * inches(.35);
}

function styleForDesignedNode(node: StudioWebNode): StudioWebNode["style"] {
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  if (node.role === "title") return { ...node.style, fontFamily: "Aptos", fontSizePt: 29.25, fontWeight: 700, lineHeight: 1.02, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } };
  if (node.role === "caption" || node.role === "label") return { ...node.style, fontFamily: "Aptos", fontSizePt: 14, fontWeight: 400, lineHeight: 1.08, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } };
  if (node.kind === "table") return { ...node.style, fontFamily: "Aptos", fontSizePt: node.style.fontSizePt, fontWeight: 400, lineHeight: 1.05, color: palette.darkMatter, background: palette.polar, borderColor: palette.graphite, borderWidthPt: .75, textAlign: "left", verticalAlign: "middle", paddingPt: { top: 4, right: 6, bottom: 4, left: 6 } };
  const presentationBodySize = node.kind === "text" && node.role === "body" && (node.text?.length ?? 0) > 360 ? PRESENTATION_DESIGN_STANDARD.defaults.typography.bodyMinimumPt : Math.max(16, Math.min(22, node.style.fontSizePt));
  return { ...node.style, fontFamily: "Aptos", fontSizePt: presentationBodySize, fontWeight: node.style.fontWeight === 700 ? 600 : node.style.fontWeight, lineHeight: 1.08, color: palette.darkMatter, background: undefined, borderColor: undefined, borderWidthPt: 0, textAlign: "left", verticalAlign: "top", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 }, objectFit: node.kind === "image" ? "contain" : node.style.objectFit };
}

function styleForComponent(node: StudioWebNode): StudioWebNode["style"] {
  const base = styleForDesignedNode(node);
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  if (node.component?.role === "eyebrow") return { ...base, fontSizePt: 10.5, fontWeight: 700, lineHeight: 1, color: palette.ornlGreen, textAlign: "left" };
  if (node.component?.role === "card-kicker") return { ...base, fontSizePt: 18, fontWeight: 700, lineHeight: 1, color: [palette.ornlGreen, palette.infinity, palette.hydro, palette.darkMatter][node.component.ordinal ?? 0] ?? palette.ornlGreen };
  if (node.component?.role === "card-heading") return { ...base, fontSizePt: 13.5, fontWeight: 400, lineHeight: 1.05, color: "#666B68" };
  if (node.component?.role === "card-body") return { ...base, fontSizePt: 15, fontWeight: 400, lineHeight: 1.13, color: palette.darkMatter };
  if (node.component?.role === "objective-body") return { ...base, fontSizePt: 18, fontWeight: 400, lineHeight: 1.16, color: palette.darkMatter, verticalAlign: "middle" };
  if (node.component?.role === "step-heading") return { ...base, fontSizePt: 18, fontWeight: 700, lineHeight: 1.05, color: palette.ornlGreen };
  if (node.component?.role === "step-body") return { ...base, fontSizePt: 16, fontWeight: 400, lineHeight: 1.14, color: palette.darkMatter };
  if (node.component?.role === "figure-label") return { ...base, fontSizePt: 14, fontWeight: 700, lineHeight: 1.05, color: palette.ornlGreen, verticalAlign: "middle" };
  if (node.component?.role === "figure-caption") return { ...base, fontSizePt: 14, fontWeight: 400, lineHeight: 1.12, color: palette.darkMatter, verticalAlign: "middle" };
  if (node.component?.role === "technical-annotation") return { ...base, fontSizePt: Math.max(10, Math.min(12, node.style.fontSizePt)), fontWeight: node.style.fontWeight, lineHeight: 1.04, color: node.style.color, verticalAlign: "middle" };
  if (node.component?.role === "question-intro") return { ...base, fontSizePt: 15, fontWeight: 400, lineHeight: 1.12, color: "#666B68" };
  if (node.component?.role === "question-item") return { ...base, fontSizePt: 17, fontWeight: 600, lineHeight: 1.12, color: palette.darkMatter, verticalAlign: "middle" };
  if (node.component?.role === "challenge-assertion") return { ...base, fontSizePt: 18, fontWeight: 700, lineHeight: 1.08, color: palette.ornlGreen };
  if (node.component?.role === "challenge-intro") return { ...base, fontSizePt: 12.5, fontWeight: 700, lineHeight: 1, color: "#666B68" };
  if (node.component?.role === "challenge-body") return { ...base, fontSizePt: 14, fontWeight: 400, lineHeight: 1.12, color: palette.darkMatter };
  if (node.component?.role === "process-input") return { ...base, fontSizePt: 13.5, fontWeight: 600, lineHeight: 1.08, color: palette.darkMatter, verticalAlign: "middle" };
  if (node.component?.role === "process-stage" || node.component?.role === "process-output") return { ...base, fontSizePt: 15, fontWeight: 700, lineHeight: 1.04, color: "#FFFFFF", textAlign: "center", verticalAlign: "middle" };
  if (node.component?.role === "supporting-copy") return { ...base, fontSizePt: 14, fontWeight: 400, lineHeight: 1.14, color: palette.darkMatter };
  if (node.component?.role === "footer-logo") return { ...base, paddingPt: { top: 0, right: 0, bottom: 0, left: 0 }, objectFit: "contain" };
  if (node.component?.role === "footer-meta") return { ...base, fontSizePt: 9, fontWeight: 400, lineHeight: 1, color: "#6B716E", textAlign: "right", verticalAlign: "middle" };
  return base;
}

function fittedStyle(node: StudioWebNode, style: StudioWebNode["style"], target: StudioWebFrame): StudioWebNode["style"] {
  if (node.kind !== "text" || !node.text?.trim()) return style;
  const widthPt = target.width / EMU_PER_POINT - style.paddingPt.left - style.paddingPt.right;
  const heightPt = target.height / EMU_PER_POINT - style.paddingPt.top - style.paddingPt.bottom;
  if (widthPt <= 0 || heightPt <= 0) return style;
  const minimum = node.role === "title" ? 22 : node.component?.role === "challenge-intro" ? 10.5 : node.component?.role === "footer-meta" ? 8.5 : node.component?.role === "process-input" ? 12.5 : node.role === "caption" || node.role === "label" ? 10 : 14;
  const paragraphs = node.sourceParagraphs?.filter((paragraph) => paragraph.text.trim()) ?? [{ text: node.text }];
  const fits = (fontSizePt: number) => {
    const averageGlyphWidth = fontSizePt * .60;
    const lineCapacity = Math.max(8, Math.floor(widthPt / averageGlyphWidth));
    const lines = paragraphs.reduce((sum, paragraph) => sum + Math.max(1, Math.ceil(paragraph.text.length / lineCapacity)), 0);
    return lines * fontSizePt * style.lineHeight <= heightPt - 1;
  };
  let fontSizePt = style.fontSizePt;
  while (fontSizePt > minimum && !fits(fontSizePt)) fontSizePt = Math.max(minimum, Math.round((fontSizePt - .5) * 2) / 2);
  return { ...style, fontSizePt };
}

function splitVertical(value: StudioWebFrame, weights: number[], ordinal: number): StudioWebFrame {
  const total = Math.max(1, weights.reduce((sum, weight) => sum + weight, 0));
  const before = weights.slice(0, ordinal).reduce((sum, weight) => sum + weight, 0);
  const current = weights[ordinal] ?? 1;
  const y = value.y + Math.round(value.height * before / total);
  const bottom = value.y + Math.round(value.height * (before + current) / total);
  return { x: value.x, y, width: value.width, height: Math.max(inches(.1), bottom - y), rotation: value.rotation };
}

function restoreSemanticAtoms(scene: StudioWebScene, slideNumber: number): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide?.nodes.some((node) => node.sourceBinding === "semantic-atom")) return scene;
  const sourceNodeIds = new Set(slide.nodes.flatMap((node) => node.sourceAtom ? [node.sourceAtom.sourceNodeId] : []));
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      contentCoverage: { ...item.contentCoverage, mappedTextNodeCount: item.nodes.filter((node) => node.sourceBinding !== "semantic-atom" && (sourceNodeIds.has(node.id) || (node.visible && Boolean(nodeVisibleText(node))))).length },
      nodes: item.nodes.filter((node) => node.sourceBinding !== "semantic-atom").map((node) => sourceNodeIds.has(node.id) ? { ...node, visible: true } : node),
      updatedAt: now,
    }),
  };
}

export function atomizeStudioWebSlide(scene: StudioWebScene, slideNumber: number, sourceNodeIds?: string[]): StudioWebScene {
  const restored = restoreSemanticAtoms(scene, slideNumber);
  const slide = restored.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  const requested = sourceNodeIds ? new Set(sourceNodeIds) : undefined;
  const atoms: StudioWebNode[] = [];
  const atomized = new Set<string>();
  for (const node of slide.nodes) {
    const paragraphs = node.sourceParagraphs?.filter((paragraph) => paragraph.text.trim()) ?? [];
    if (node.kind !== "text" || node.role !== "body" || paragraphs.length < 2 || (requested && !requested.has(node.id))) continue;
    atomized.add(node.id);
    const weights = paragraphs.map((paragraph) => Math.max(24, paragraph.characterCount));
    paragraphs.forEach((paragraph, ordinal) => {
      atoms.push({
        ...node,
        id: `${node.id}-atom-${paragraph.index}`,
        sourceShapeId: `${node.sourceShapeId}#p${paragraph.index}`,
        sourceBinding: "semantic-atom",
        name: `${node.name} · paragraph ${paragraph.index}`,
        sourceFrame: splitVertical(node.sourceFrame, weights, ordinal),
        frame: splitVertical(node.frame, weights, ordinal),
        zIndex: node.zIndex * 100 + ordinal,
        sourceTextOrder: node.sourceTextOrder + paragraphs.slice(0, ordinal).reduce((sum, candidate) => sum + candidate.characterCount + 1, 0),
        visible: true,
        text: paragraph.text,
        textHash: paragraph.textHash,
        sourceParagraphs: [paragraph],
        sourceAtom: {
          sourceNodeId: node.id,
          sourceObjectId: node.sourceObjectId,
          paragraphStart: paragraph.index,
          paragraphEnd: paragraph.index,
          ordinal,
          count: paragraphs.length,
          aggregateSourceTextHash: node.textHash,
        },
      });
    });
  }
  if (atoms.length === 0) return restored;
  const now = new Date().toISOString();
  return {
    ...restored,
    revision: `${restored.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: restored.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      contentCoverage: { ...item.contentCoverage, mappedTextNodeCount: item.contentCoverage.mappedTextNodeCount - atomized.size + atoms.length },
      nodes: [...item.nodes.map((node) => atomized.has(node.id) ? { ...node, visible: false } : node), ...atoms],
      updatedAt: now,
    }),
  };
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

function scaleSourceGroup(nodes: StudioWebNode[], target: StudioWebFrame): Map<string, StudioWebFrame> {
  const placements = new Map<string, StudioWebFrame>();
  if (nodes.length === 0) return placements;
  const source = unionStudioFrames(nodes.map((node) => node.sourceFrame));
  if (source.width <= 0 || source.height <= 0) return placements;
  const scale = Math.min(target.width / source.width, target.height / source.height);
  const width = source.width * scale;
  const height = source.height * scale;
  const originX = target.x + (target.width - width) / 2;
  const originY = target.y + (target.height - height) / 2;
  for (const node of nodes) {
    placements.set(node.id, {
      x: Math.round(originX + (node.sourceFrame.x - source.x) * scale),
      y: Math.round(originY + (node.sourceFrame.y - source.y) * scale),
      width: Math.max(inches(.02), Math.round(node.sourceFrame.width * scale)),
      height: Math.max(inches(.02), Math.round(node.sourceFrame.height * scale)),
      rotation: node.sourceFrame.rotation,
    });
  }
  return placements;
}

function sourceCenterDistance(left: StudioWebNode, right: StudioWebNode): number {
  const leftX = left.sourceFrame.x + left.sourceFrame.width / 2;
  const leftY = left.sourceFrame.y + left.sourceFrame.height / 2;
  const rightX = right.sourceFrame.x + right.sourceFrame.width / 2;
  const rightY = right.sourceFrame.y + right.sourceFrame.height / 2;
  return Math.hypot(leftX - rightX, leftY - rightY);
}

function shortFigureAnnotation(node: StudioWebNode): boolean {
  return node.kind === "text"
    && ["caption", "label"].includes(node.role)
    && (node.text?.length ?? 0) > 0
    && (node.text?.length ?? 0) <= 80
    && node.sourceFrame.height <= inches(1.05);
}

function assignFigureRelationships(visuals: StudioWebNode[], candidates: StudioWebNode[]): Map<string, StudioWebNode[]> {
  const assignments = new Map(visuals.map((visual) => [visual.id, [] as StudioWebNode[]]));
  for (const candidate of candidates) {
    const visual = [...visuals].sort((left, right) => sourceCenterDistance(left, candidate) - sourceCenterDistance(right, candidate))[0];
    if (visual) assignments.get(visual.id)?.push(candidate);
  }
  return assignments;
}

function sourceFrameContains(container: StudioWebFrame, candidate: StudioWebFrame, padding = inches(.12)): boolean {
  const centerX = candidate.x + candidate.width / 2;
  const centerY = candidate.y + candidate.height / 2;
  return centerX >= container.x - padding
    && centerX <= container.x + container.width + padding
    && centerY >= container.y - padding
    && centerY <= container.y + container.height + padding;
}

function narrativeHeightInches(nodes: StudioWebNode[], widthInches: number, minimum = 1.05, maximum = 1.75): number {
  const characters = nodes.reduce((sum, node) => sum + (node.text?.length ?? 0), 0);
  const paragraphs = nodes.reduce((sum, node) => sum + Math.max(1, node.sourceParagraphs?.length ?? 1), 0);
  const estimatedLines = characters / Math.max(36, widthInches * 8.4) + Math.max(0, paragraphs - 1) * .35;
  return Math.max(minimum, Math.min(maximum, .48 + estimatedLines * .27));
}

export function recommendedStudioRecipe(slide: StudioWebSlide): StudioLayoutRecipe {
  const nodes = activeNodes(slide);
  if (nodes.some((node) => node.kind === "table")) return "ornl-title-table";
  const nativeObject = slide.nodes.some((node) => node.visible && node.kind === "native-object" && !footerNode(node));
  const directBodyParagraphs = nodes
    .filter((node) => node.kind === "text" && node.role === "body" && node.sourceBinding === "editable-object" && !footerNode(node))
    .flatMap((node) => node.sourceParagraphs?.filter((paragraph) => paragraph.text.trim()) ?? []);
  const questionCount = directBodyParagraphs.filter((paragraph) => paragraph.text.trim().endsWith("?")).length;
  if (nativeObject && questionCount >= 3 && questionCount <= 6) return "ornl-title-question-diagram";
  if (nativeObject) return "ornl-title-two-column";
  const eligibleParagraphs = nodes.filter((node) => node.kind === "text" && node.role === "body" && !footerNode(node)).reduce((sum, node) => sum + Math.max(1, node.sourceParagraphs?.length ?? 1), 0);
  const bodyCharacters = nodes.filter((node) => node.kind === "text" && node.role === "body" && !footerNode(node)).reduce((sum, node) => sum + (node.text?.length ?? 0), 0);
  const bodyCount = nodes.filter((node) => node.kind === "text" && node.role === "body" && !footerNode(node)).length;
  const labelCount = nodes.filter((node) => node.kind === "text" && node.role === "label" && !footerNode(node)).length;
  const structuredParagraphs = nodes
    .filter((node) => node.kind === "text" && node.role === "body" && !footerNode(node))
    .flatMap((node) => node.sourceParagraphs?.filter((paragraph) => paragraph.text.trim()) ?? []);
  const structuredSectionCount = structuredParagraphs.filter((paragraph) => !paragraph.bullet && paragraph.level === 0).length;
  if (!nodes.some(meaningfulImage) && structuredParagraphs.length >= 4 && structuredParagraphs.length <= 10 && structuredSectionCount >= 2 && structuredSectionCount <= 4 && structuredParagraphs.some((paragraph) => paragraph.level > 0)) return "ornl-title-card-grid";
  if (bodyCount >= 3 && bodyCount <= 6 && labelCount >= bodyCount) return "ornl-title-card-grid";
  const images = nodes.filter(meaningfulImage).length;
  const explanatoryLabels = nodes.filter((node) => node.kind === "text" && ["label", "caption"].includes(node.role) && !footerNode(node)).length;
  const connectorCount = slide.nodes.filter((node) => node.visible && node.kind === "connector" && !footerNode(node)).length;
  const shapeCount = slide.nodes.filter((node) => node.visible && node.kind === "shape" && !footerNode(node)).length;
  if (images >= 4 && explanatoryLabels >= 4 && bodyCount >= 3) return "ornl-title-process-flow";
  if (images >= 1 && eligibleParagraphs >= 5 && connectorCount >= 3 && shapeCount >= 3) return "ornl-title-challenges-evidence";
  if (images >= 2 && explanatoryLabels >= images) return "ornl-title-labeled-figure-grid";
  if (images >= 2) return "ornl-title-figure-grid";
  if (images === 1 && eligibleParagraphs >= 2 && eligibleParagraphs <= 5) return "ornl-title-steps-evidence";
  if (images === 1) return "ornl-title-two-column";
  if (images === 0 && eligibleParagraphs >= 3 && eligibleParagraphs <= 6 && bodyCharacters / eligibleParagraphs >= 180) return "ornl-title-card-grid";
  if (images === 0 && eligibleParagraphs >= 2 && eligibleParagraphs <= 4) return "ornl-title-objective-columns";
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
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  const hasEyebrow = slide.nodes.some((node) => node.component?.role === "eyebrow");
  const title = slide.nodes.find((node) => node.visible && node.role === "title");
  const titleBottom = title ? emuInches(title.frame.y + title.frame.height) : hasEyebrow ? 1.25 : 1.12;
  const components: StudioGeneratedComponent[] = slide.recipe === "source" || slide.recipe === "template-layout" ? [] : [{ id: `studio-title-rule-${slide.slideNumber}`, kind: "rect", frame: frame(.47, titleBottom + .03, hasEyebrow ? .62 : .96, .035), fillColor: palette.ornlGreen, lineWidthPt: 0, behindContent: true }];
  for (const treatment of (slide.figureTreatments ?? []).filter((item) => item.mode === "preserve-and-frame")) {
    const nodes = treatment.nodeIds.map((id) => slide.nodes.find((node) => node.id === id)).filter((node): node is StudioWebNode => Boolean(node?.visible));
    if (!nodes.length) continue;
    const padding = points(6);
    const slideWidth = PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches * EMU_PER_INCH;
    const slideHeight = PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches * EMU_PER_INCH;
    const left = Math.max(0, Math.min(...nodes.map((node) => node.frame.x)) - padding);
    const top = Math.max(0, Math.min(...nodes.map((node) => node.frame.y)) - padding);
    const right = Math.min(slideWidth, Math.max(...nodes.map((node) => node.frame.x + node.frame.width)) + padding);
    const bottom = Math.min(slideHeight, Math.max(...nodes.map((node) => node.frame.y + node.frame.height)) + padding);
    const figureFrame = { x: left, y: top, width: Math.max(points(12), right - left), height: Math.max(points(12), bottom - top), rotation: 0 };
    components.push({ id: `${treatment.id}-surface`, kind: "rect", frame: figureFrame, fillColor: palette.polar, lineColor: palette.graphite, lineWidthPt: .75, behindContent: true });
    components.push({ id: `${treatment.id}-accent`, kind: "rect", frame: { ...figureFrame, height: points(2) }, fillColor: palette.ornlGreen, lineWidthPt: 0, behindContent: true });
  }
  if (slide.recipe === "source" || slide.recipe === "template-layout") return components;
  if (slide.recipe === "ornl-title-objective-columns") {
    const groups = slide.nodes.filter((node) => node.component?.role === "objective-body");
    groups.forEach((node, ordinal) => {
      const padding = points(12);
      const surface = { x: node.frame.x - padding, y: node.frame.y - padding, width: node.frame.width + padding * 2, height: node.frame.height + padding * 2, rotation: 0 };
      components.push({ id: `${node.component!.groupId}-surface`, kind: "rect", frame: surface, fillColor: "#F2F5F3", lineWidthPt: 0, behindContent: true });
      components.push({ id: `${node.component!.groupId}-accent`, kind: "rect", frame: { ...surface, height: points(3) }, fillColor: [palette.ornlGreen, palette.aqua, palette.forge, palette.infinity][ordinal] ?? palette.ornlGreen, lineWidthPt: 0, behindContent: true });
    });
    return components;
  }
  if (slide.recipe === "ornl-title-steps-evidence") {
    slide.nodes.filter((node) => node.component?.role === "step-heading").forEach((node) => components.push({ id: `${node.component!.groupId}-rail`, kind: "rect", frame: { x: node.frame.x - points(12), y: node.frame.y, width: points(3), height: node.frame.height, rotation: 0 }, fillColor: palette.ornlGreen, lineWidthPt: 0, behindContent: true }));
    return components;
  }
  if (slide.recipe === "ornl-title-labeled-figure-grid") {
    slide.nodes.filter((node) => node.component?.role === "figure-media" && node.component.frame).forEach((node) => {
      components.push({ id: `${node.component!.groupId}-visual-field`, kind: "rect", frame: node.component!.frame!, fillColor: palette.graphite, lineWidthPt: 0, behindContent: true });
    });
    const groups = [...new Set(slide.nodes.filter((node) => node.component?.role === "figure-label").map((node) => node.component!.groupId))];
    groups.slice(1).forEach((groupId) => {
      const node = slide.nodes.find((candidate) => candidate.component?.groupId === groupId);
      if (node) components.push({ id: `${groupId}-separator`, kind: "rect", frame: { x: .47 * EMU_PER_INCH, y: node.frame.y - points(10), width: 12.39 * EMU_PER_INCH, height: points(.75), rotation: 0 }, fillColor: palette.graphite, lineWidthPt: 0, behindContent: true });
    });
    return components;
  }
  if (slide.recipe === "ornl-title-question-diagram") {
    const questions = slide.nodes.filter((node) => node.component?.role === "question-item").sort((left, right) => (left.component?.ordinal ?? 0) - (right.component?.ordinal ?? 0));
    if (questions.length) {
      const top = Math.min(...questions.map((node) => node.frame.y));
      const bottom = Math.max(...questions.map((node) => node.frame.y + node.frame.height));
      components.push({ id: `studio-question-rail-${slide.slideNumber}`, kind: "rect", frame: { x: questions[0].frame.x - points(13), y: top, width: points(3), height: bottom - top, rotation: 0 }, fillColor: palette.ornlGreen, lineWidthPt: 0, behindContent: true });
      questions.slice(1).forEach((node, ordinal) => components.push({ id: `studio-question-separator-${slide.slideNumber}-${ordinal + 1}`, kind: "rect", frame: { x: node.frame.x, y: node.frame.y - points(7), width: node.frame.width, height: points(.75), rotation: 0 }, fillColor: palette.graphite, lineWidthPt: 0, behindContent: true }));
    }
    return components;
  }
  if (slide.recipe === "ornl-title-challenges-evidence") {
    const challenges = slide.nodes.filter((node) => node.component?.role === "challenge-body");
    const accents = [palette.ornlGreen, palette.aqua, palette.forge];
    challenges.forEach((node, ordinal) => {
      const padding = points(10);
      components.push({ id: `${node.component!.groupId}-surface`, kind: "rect", frame: { x: node.frame.x - padding, y: node.frame.y - padding, width: node.frame.width + padding * 2, height: node.frame.height + padding * 2, rotation: 0 }, fillColor: palette.polar, lineColor: palette.graphite, lineWidthPt: .75, behindContent: true });
      components.push({ id: `${node.component!.groupId}-accent`, kind: "rect", frame: { x: node.frame.x - padding, y: node.frame.y - padding, width: node.frame.width + padding * 2, height: points(2), rotation: 0 }, fillColor: accents[ordinal] ?? palette.ornlGreen, lineWidthPt: 0, behindContent: true });
    });
    return components;
  }
  if (slide.recipe === "ornl-title-process-flow") {
    const inputGroups = [...new Set(slide.nodes.filter((node) => node.component?.role === "process-input").map((node) => node.component!.groupId))];
    inputGroups.forEach((groupId, ordinal) => {
      const nodes = slide.nodes.filter((node) => node.component?.groupId === groupId);
      if (!nodes.length) return;
      const bounds = unionStudioFrames(nodes.map((node) => node.frame));
      const padding = points(8);
      components.push({ id: `${groupId}-surface`, kind: "rect", frame: { x: bounds.x - padding, y: bounds.y - padding, width: bounds.width + padding * 2, height: bounds.height + padding * 2, rotation: 0 }, fillColor: palette.polar, lineColor: palette.graphite, lineWidthPt: .75, behindContent: true });
      components.push({ id: `${groupId}-accent`, kind: "rect", frame: { x: bounds.x - padding, y: bounds.y - padding, width: points(2), height: bounds.height + padding * 2, rotation: 0 }, fillColor: [palette.ornlGreen, palette.aqua, palette.infinity, palette.forge][ordinal] ?? palette.ornlGreen, lineWidthPt: 0, behindContent: true });
    });
    slide.nodes.filter((node) => node.component?.role === "process-stage" || node.component?.role === "process-output").forEach((node) => components.push({ id: `${node.component!.groupId}-surface`, kind: "rect", frame: node.frame, fillColor: node.component?.role === "process-output" ? palette.haleNavy : palette.ornlGreen, lineWidthPt: 0, behindContent: true }));
    return components;
  }
  if (slide.recipe !== "ornl-title-card-grid") return components;
  const groups = [...new Set(slide.nodes.filter((node) => node.component?.role === "card-body").map((node) => node.component!.groupId))];
  const accents = [palette.ornlGreen, palette.aqua, palette.infinity, palette.forge, palette.plasma, palette.pulsar];
  groups.forEach((groupId, ordinal) => {
    const groupNodes = slide.nodes.filter((node) => node.component?.groupId === groupId);
    const card = groupNodes.find((node) => node.component?.frame)?.component?.frame ?? cardFrame(ordinal, groups.length);
    components.push({ id: `${groupId}-surface`, kind: "rect", frame: card, fillColor: "#F2F5F3", lineWidthPt: 0, behindContent: true });
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
  const originalSlide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!originalSlide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  const recipe = requestedRecipe ?? recommendedStudioRecipe(originalSlide);
  const atomRecipes: StudioLayoutRecipe[] = ["ornl-title-objective-columns", "ornl-title-steps-evidence", "ornl-title-card-grid", "ornl-title-challenges-evidence", "ornl-title-process-flow"];
  const questionSourceIds = recipe === "ornl-title-question-diagram"
    ? originalSlide.nodes.filter((node) => node.kind === "text" && node.role === "body" && node.sourceBinding === "editable-object" && (node.sourceParagraphs?.filter((paragraph) => paragraph.text.trim()).filter((paragraph) => paragraph.text.trim().endsWith("?")).length ?? 0) >= 3).map((node) => node.id)
    : [];
  const workingScene = recipe === "ornl-title-question-diagram"
    ? atomizeStudioWebSlide(scene, slideNumber, questionSourceIds)
    : atomRecipes.includes(recipe) ? atomizeStudioWebSlide(scene, slideNumber) : restoreSemanticAtoms(scene, slideNumber);
  const slide = workingScene.slides.find((item) => item.slideNumber === slideNumber)!;
  if (recipe === "template-layout" && !layout?.semantic) throw new Error("Choose an installed template layout with semantic regions before applying template-layout.");
  const nodes = activeNodes(slide);
  const title = nodes.find((node) => node.role === "title");
  const footer = nodes.filter(footerNode);
  const eyebrow = nodes.find((node) => node.kind === "text" && node.role === "label" && !footerNode(node) && (!title || node.sourceFrame.y < title.sourceFrame.y));
  const content = nodes.filter((node) => node.id !== title?.id && node.id !== eyebrow?.id && !footerNode(node) && node.role !== "caption" && node.role !== "label");
  const captions = nodes.filter((node) => node.id !== eyebrow?.id && !footerNode(node) && (node.role === "caption" || node.role === "label"));
  const placements = new Map<string, StudioWebFrame>();
  const components = new Map<string, StudioWebNode["component"]>();
  const generatedFigureTreatments: StudioFigureTreatment[] = [];
  const titleFrame = frame(.47, eyebrow ? .58 : .26, 12.39, eyebrow ? .68 : (title?.text?.length ?? 0) > 54 ? 1.14 : .95);
  const contentTop = emuInches(titleFrame.y + titleFrame.height) + .12;
  const contentBottom = 6.64;
  const contentHeight = Math.max(1, contentBottom - contentTop);
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
    const visuals = remaining.filter(meaningfulImage).sort((left, right) => left.sourceFrame.y - right.sourceFrame.y || left.sourceFrame.x - right.sourceFrame.x);
    const connectors = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "connector" && !footerNode(node));
    const annotations = captions.filter(shortFigureAnnotation);
    if (tables.length === 1 && visuals.length >= 2) {
      const relationshipAssignments = assignFigureRelationships(visuals, [...connectors, ...annotations]);
      const relationshipIds = new Set([...connectors, ...annotations].map((node) => node.id));
      const prose = support.filter((node) => !visuals.some((visual) => visual.id === node.id) && !relationshipIds.has(node.id));
      const lead = prose.filter((node) => node.kind === "text" && node.sourceFrame.y < visuals[0].sourceFrame.y);
      const reference = prose.filter((node) => node.kind === "text" && node.sourceFrame.y >= inches(6.35));
      const middle = prose.filter((node) => !lead.includes(node) && !reference.includes(node));
      const leadHeight = lead.length ? narrativeHeightInches(lead, 12.39, .58, .82) : 0;
      const topFigureY = contentTop + leadHeight + (lead.length ? .10 : 0);
      const topFigureHeight = Math.max(1.10, Math.min(1.48, contentHeight * .28));
      const referenceHeight = reference.length ? .48 : 0;
      const mainY = topFigureY + topFigureHeight + .18;
      const mainBottom = contentBottom - referenceHeight - (reference.length ? .10 : 0);
      const mainHeight = Math.max(1.55, mainBottom - mainY);
      const leftWidth = 5.18;
      if (lead.length) for (const [id, value] of stack(lead, frame(.47, contentTop, 12.39, leadHeight), 5)) placements.set(id, value);
      const topVisual = visuals[0];
      const lowerVisuals = visuals.slice(1);
      const placeRelationshipGroup = (visual: StudioWebNode, target: StudioWebFrame, ordinal: number) => {
        const groupNodes = [visual, ...(relationshipAssignments.get(visual.id) ?? [])];
        for (const [id, value] of scaleSourceGroup(groupNodes, target)) placements.set(id, value);
        generatedFigureTreatments.push({
          id: `studio-auto-table-evidence-${slideNumber}-${ordinal}`,
          nodeIds: groupNodes.map((node) => node.id),
          mode: "preserve-as-unit",
          verificationStatus: "source-locked",
          intentSummary: ordinal === 1 ? "Preserve the source overview figure and its callout relationships as one technical evidence unit." : "Preserve the source supporting figure and its callout relationships as one technical evidence unit.",
          informationInventory: ["All source figure pixels, labels, arrows, and relative relationships"],
          invariants: ["Do not alter source labels, values, topology, arrows, or technical meaning."],
          rationale: "The figure carries relationship meaning that is safer to preserve natively while the surrounding slide is recomposed.",
          groupFrame: target,
          lockAspectRatio: true,
          relationshipPolicy: "preserve-internal",
        });
      };
      placeRelationshipGroup(topVisual, frame(.47, topFigureY, 12.39, topFigureHeight), 1);
      const middleHeight = middle.length ? Math.min(1.14, Math.max(.72, narrativeHeightInches(middle, leftWidth, .72, 1.14))) : 0;
      if (middle.length) for (const [id, value] of stack(middle, frame(.47, mainY, leftWidth, middleHeight), 5)) placements.set(id, value);
      placements.set(tables[0].id, frame(.47, mainY + middleHeight + (middle.length ? .12 : 0), leftWidth, Math.max(.72, mainHeight - middleHeight - (middle.length ? .12 : 0))));
      if (lowerVisuals.length === 1) placeRelationshipGroup(lowerVisuals[0], frame(5.93, mainY, 6.93, mainHeight), 2);
      else lowerVisuals.forEach((visual, index) => placeRelationshipGroup(visual, frame(5.93, mainY + index * (mainHeight / lowerVisuals.length), 6.93, mainHeight / lowerVisuals.length - .08), index + 2));
      if (reference.length) for (const [id, value] of stack(reference, frame(5.93, mainBottom + .08, 6.93, referenceHeight), 3)) placements.set(id, value);
    } else if (tables.length === 1 && support.length >= 2) {
      placements.set(tables[0].id, frame(.47, contentTop, 8.05, contentHeight));
      for (const [id, value] of stack(support, frame(8.90, contentTop, 3.96, contentHeight), 12)) placements.set(id, value);
    } else {
      const primary = tables[0];
      const adaptiveHeight = primary?.table ? Math.max(3.15, Math.min(5.48, primary.table.rows * .66 + .44)) : 5.30;
      for (const [id, value] of stack(tables, frame(.47, contentTop, 12.39, support.length ? Math.max(3.15, contentHeight - .72) : Math.min(contentHeight, adaptiveHeight)), 12)) placements.set(id, value);
      for (const [id, value] of stack(support, frame(.47, 5.98, 12.39, .64), 8)) placements.set(id, value);
    }
  } else if (recipe === "ornl-title-question-diagram") {
    const atoms = content.filter((node) => node.kind === "text" && node.sourceBinding === "semantic-atom").sort((left, right) => left.sourceTextOrder - right.sourceTextOrder);
    const intro = atoms.find((node) => !node.text?.trim().endsWith("?"));
    const questions = atoms.filter((node) => node.id !== intro?.id && node.text?.trim().endsWith("?")).slice(0, 6);
    const nativeObject = slide.nodes.find((node) => node.visible && node.kind === "native-object" && !footerNode(node));
    const embeddedNativeText = nativeObject
      ? slide.nodes.filter((node) => node.visible && node.sourceBinding === "catalog-derived" && node.kind === "text" && sourceFrameContains(nativeObject.sourceFrame, node.sourceFrame))
      : [];
    if (intro) {
      placements.set(intro.id, frame(.58, contentTop, 3.18, .82));
      components.set(intro.id, { groupId: `studio-question-${slideNumber}-intro`, role: "question-intro" });
    }
    const questionTop = contentTop + (intro ? 1.00 : .16);
    const questionGap = .18;
    const questionHeight = Math.max(.62, Math.min(.92, (contentBottom - questionTop - questionGap * Math.max(0, questions.length - 1)) / Math.max(1, questions.length)));
    questions.forEach((node, ordinal) => {
      placements.set(node.id, frame(.65, questionTop + ordinal * (questionHeight + questionGap), 3.02, questionHeight));
      components.set(node.id, { groupId: `studio-question-${slideNumber}-${ordinal + 1}`, role: "question-item", ordinal });
    });
    if (nativeObject) {
      const groupNodes = [...new Map([nativeObject, ...embeddedNativeText].map((node) => [node.id, node])).values()];
      const target = frame(4.03, contentTop + .10, 8.83, contentHeight - .10);
      for (const [id, value] of scaleSourceGroup(groupNodes, target)) placements.set(id, value);
      generatedFigureTreatments.push({
        id: `studio-auto-question-diagram-${slideNumber}`,
        nodeIds: groupNodes.map((node) => node.id),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: "Preserve the complete technical diagram and its internal glossary as the evidence that answers the source questions.",
        informationInventory: ["Complete native PowerPoint group", "Every internal label, shape, connector, arrow, glossary entry, and spatial relationship"],
        invariants: ["Do not alter source labels, topology, values, arrows, glossary definitions, or technical relationships."],
        rationale: "The slide's communication job is a question rail paired with a dense relationship-bearing technical figure, so Studio gives each a dedicated region without interpreting or redrawing the source diagram.",
        groupFrame: target,
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
    }
    const reserved = new Set([intro?.id, ...questions.map((node) => node.id), nativeObject?.id, ...embeddedNativeText.map((node) => node.id)].filter((id): id is string => Boolean(id)));
    const leftovers = [...content, ...captions].filter((node) => !reserved.has(node.id) && node.sourceBinding !== "catalog-derived");
    if (leftovers.length) for (const [id, value] of stack(leftovers, frame(.58, 6.18, 3.18, .40), 4)) placements.set(id, value);
  } else if (recipe === "ornl-title-two-column") {
    const nativeObject = slide.nodes.find((node) => node.visible && node.kind === "native-object" && !footerNode(node));
    const technicalVisuals = slide.nodes.filter((node) => node.visible && !footerNode(node) && (node.kind === "native-object" || meaningfulImage(node)));
    const editableNarrative = [...content, ...captions].filter((node) => node.kind === "text" && node.role !== "title" && node.sourceBinding === "editable-object");
    const compositeTechnicalOverview = Boolean(nativeObject && technicalVisuals.length >= 2 && editableNarrative.length === 0);
    const visual = nativeObject ?? content.find((node) => meaningfulImage(node) || node.kind === "table");
    const connectors = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "connector" && !footerNode(node));
    const figureAnnotations = captions.filter(shortFigureAnnotation);
    const embeddedNativeText = nativeObject
      ? slide.nodes.filter((node) => node.visible && node.sourceBinding === "catalog-derived" && node.kind === "text" && sourceFrameContains(nativeObject.sourceFrame, node.sourceFrame))
      : [];
    const relationshipBearingFigure = visual && (visual.kind === "native-object" || (visual.kind === "image" && connectors.length >= 1 && figureAnnotations.length >= 2));
    if (compositeTechnicalOverview) {
      const groupNodes = slide.nodes.filter((node) => node.visible && node.role !== "title" && !footerNode(node));
      const target = frame(.47, contentTop + .06, 12.39, contentHeight - .12);
      for (const [id, value] of scaleSourceGroup(groupNodes, target)) placements.set(id, value);
      generatedFigureTreatments.push({
        id: `studio-auto-technical-overview-${slideNumber}`,
        nodeIds: groupNodes.map((node) => node.id),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: "Preserve the complete multi-part native PowerPoint control overview as one relationship-bearing technical evidence field.",
        informationInventory: ["Every source panel, embedded image, group, label, connector, value, color field, and spatial relationship"],
        invariants: ["Do not separate source panels or alter labels, colors, topology, values, arrows, or technical relationships."],
        rationale: "The slide is one dense technical system assembled from mixed PowerPoint groups and legacy media. Studio improves the title hierarchy and usable scale while keeping the entire technical field intact.",
        groupFrame: target,
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
    } else if (visual && relationshipBearingFigure) {
      const groupNodes = [...new Map([visual, ...connectors, ...figureAnnotations, ...embeddedNativeText].map((node) => [node.id, node])).values()];
      const groupIds = new Set(groupNodes.map((node) => node.id));
      const narrative = [...content, ...captions].filter((node) => !groupIds.has(node.id));
      const narrativeParagraphs = narrative.reduce((sum, node) => sum + Math.max(1, node.sourceParagraphs?.length ?? 1), 0);
      const nativeMinimum = Math.min(2.0, .55 + narrativeParagraphs * .26);
      const narrativeHeight = narrative.length ? narrativeHeightInches(narrative, 12.39, visual.kind === "native-object" ? nativeMinimum : .62, visual.kind === "native-object" ? 2.0 : .92) : 0;
      if (visual.kind === "native-object") narrative.filter((node) => node.kind === "text").forEach((node, ordinal) => components.set(node.id, { groupId: `studio-native-figure-support-${slideNumber}`, role: "supporting-copy", ordinal }));
      if (narrative.length) for (const [id, value] of stack(narrative, frame(.47, contentTop, 12.39, narrativeHeight), 6)) placements.set(id, value);
      const target = frame(.47, contentTop + narrativeHeight + (narrative.length ? .14 : 0), 12.39, Math.max(1.25, contentHeight - narrativeHeight - (narrative.length ? .14 : 0)));
      for (const [id, value] of scaleSourceGroup(groupNodes, target)) placements.set(id, value);
      generatedFigureTreatments.push({
        id: `studio-auto-technical-figure-${slideNumber}`,
        nodeIds: groupNodes.map((node) => node.id),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: visual.kind === "native-object" ? "Preserve the grouped native PowerPoint diagram as one relationship-bearing evidence unit." : "Preserve the technical figure, labels, and connectors as one relationship-bearing evidence unit.",
        informationInventory: visual.kind === "native-object" ? ["Complete native PowerPoint group", "Every internal label, shape, connector, arrow, and spatial relationship"] : ["Source image", "All short figure labels", "All source connectors and arrows"],
        invariants: ["Do not separate labels from their targets or alter arrows, topology, values, or technical relationships."],
        rationale: "The source figure is visually weak but information-dense; Studio improves the page around it without guessing at its internal meaning.",
        groupFrame: target,
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
    } else {
      const left = content.filter((node) => node.id !== visual?.id);
      const wideEvidence = visual?.kind === "image"
        && visual.sourceFrame.width / Math.max(1, visual.sourceFrame.height) >= 1.30
        && left.reduce((sum, node) => sum + (node.text?.length ?? 0), 0) <= 180;
      if (wideEvidence && visual) {
        const narrativeHeight = left.length ? .72 : 0;
        for (const [id, value] of stack(left, frame(.47, contentTop, 12.39, narrativeHeight), 6)) placements.set(id, value);
        const captionHeight = captions.length ? .40 : 0;
        placements.set(visual.id, contained(visual, frame(.47, contentTop + narrativeHeight + (left.length ? .12 : 0), 12.39, contentHeight - narrativeHeight - captionHeight - (left.length ? .12 : 0) - (captions.length ? .10 : 0))));
        for (const [id, value] of stack(captions, frame(.47, contentBottom - captionHeight, 12.39, captionHeight), 4)) placements.set(id, value);
      } else {
        for (const [id, value] of stack(left, frame(.47, contentTop, 5.78, contentHeight), 18)) placements.set(id, value);
        if (visual) placements.set(visual.id, contained(visual, frame(6.63, contentTop, 6.23, contentHeight - (captions.length ? .40 : 0))));
        for (const [id, value] of stack(captions, frame(6.63, contentBottom - .34, 6.23, .34), 4)) placements.set(id, value);
      }
    }
  } else if (recipe === "ornl-title-figure-grid") {
    const visuals = content.filter(meaningfulImage);
    const narrative = [...content.filter((node) => node.kind !== "image"), ...captions];
    const fieldConnectors = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "connector" && !footerNode(node));
    const fieldShapes = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "shape" && /arrow/i.test(node.name) && !footerNode(node));
    const annotations = captions.filter(shortFigureAnnotation);
    if (visuals.length >= 2 && fieldConnectors.length >= 1) {
      const figureNodes = [...visuals, ...fieldConnectors, ...fieldShapes, ...annotations];
      const figureIds = new Set(figureNodes.map((node) => node.id));
      const prose = narrative.filter((node) => !figureIds.has(node.id));
      // Native PowerPoint can wrap dense instructional copy one line taller than
      // the browser estimate. Reserve enough vertical room for that final line
      // while keeping the relationship-bearing figure field intact above it.
      const proseHeight = prose.length ? narrativeHeightInches(prose, 12.39, .72, 2.05) : 0;
      const target = frame(.47, contentTop, 12.39, Math.max(1.35, contentHeight - proseHeight - (prose.length ? .14 : 0)));
      for (const [id, value] of scaleSourceGroup(figureNodes, target)) placements.set(id, value);
      if (prose.length) for (const [id, value] of stack(prose, frame(.47, contentBottom - proseHeight, 12.39, proseHeight), 6)) placements.set(id, value);
      generatedFigureTreatments.push({
        id: `studio-auto-figure-field-${slideNumber}`,
        nodeIds: figureNodes.map((node) => node.id),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: "Preserve the multi-image technical callout field as one relationship-bearing evidence unit.",
        informationInventory: ["All source images", "All short figure annotations", "All source connectors and arrows between the images"],
        invariants: ["Do not detach cross-image arrows or alter source labels, values, sequence, topology, or technical meaning."],
        rationale: "This nominal figure-grid slide contains cross-image connector meaning, so Studio preserves and scales the complete native relationship field instead of exporting unbound lines or separating its evidence.",
        groupFrame: target,
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
    } else if (visuals.length === 2 && narrative.length) {
      const narrativeHeight = narrative.length > 1 ? 2.32 : Math.min(2.25, narrativeHeightInches(narrative, 12.39, 1.35, 2.25) + .30);
      const visualBottom = contentBottom - narrativeHeight - .16;
      for (const [id, value] of grid(visuals, frame(.47, contentTop, 12.39, Math.max(1.0, visualBottom - contentTop)), 2, 18)) placements.set(id, value);
      if (narrative.length === 2) {
        placements.set(narrative[0].id, frame(.47, visualBottom + .16, 12.39, 1.08));
        placements.set(narrative[1].id, frame(.47, visualBottom + 1.36, 12.39, .96));
      } else for (const [id, value] of stack(narrative, frame(.47, visualBottom + .16, 12.39, narrativeHeight), 7)) placements.set(id, value);
    } else {
      for (const [id, value] of grid(visuals, frame(.47, contentTop, 12.39, Math.max(1, contentHeight - (narrative.length ? .84 : 0))), visuals.length <= 2 ? 2 : 3, 18)) placements.set(id, value);
      for (const [id, value] of stack(narrative, frame(.47, contentBottom - .68, 12.39, .68), 8)) placements.set(id, value);
    }
  } else if (recipe === "ornl-title-objective-columns") {
    const objectives = content.filter((node) => node.kind === "text").sort((left, right) => left.zIndex - right.zIndex);
    const count = Math.max(1, objectives.length);
    const gap = .24;
    const width = (12.39 - gap * (count - 1)) / count;
    objectives.forEach((node, ordinal) => {
      placements.set(node.id, frame(.64 + ordinal * (width + gap), 1.69, width - .34, 4.44));
      components.set(node.id, { groupId: `studio-objective-${slideNumber}-${ordinal + 1}`, role: "objective-body", ordinal });
    });
    for (const [id, value] of stack(captions, frame(.47, 6.42, 12.39, .28), 4)) placements.set(id, value);
  } else if (recipe === "ornl-title-challenges-evidence") {
    const atoms = content.filter((node) => node.kind === "text" && node.sourceBinding === "semantic-atom").sort((left, right) => left.sourceTextOrder - right.sourceTextOrder);
    const assertion = atoms[0];
    const intro = atoms[1];
    const challenges = atoms.slice(2, 5);
    if (assertion) {
      placements.set(assertion.id, frame(.47, 1.22, 12.39, .46));
      components.set(assertion.id, { groupId: `studio-challenge-${slideNumber}-assertion`, role: "challenge-assertion" });
    }
    if (intro) {
      placements.set(intro.id, frame(.47, 1.77, 12.39, .28));
      components.set(intro.id, { groupId: `studio-challenge-${slideNumber}-intro`, role: "challenge-intro" });
    }
    const cardGap = .28;
    const cardWidth = (12.39 - cardGap * 2) / 3;
    challenges.forEach((node, ordinal) => {
      placements.set(node.id, frame(.61 + ordinal * (cardWidth + cardGap), 2.25, cardWidth - .28, 1.82));
      components.set(node.id, { groupId: `studio-challenge-${slideNumber}-${ordinal + 1}`, role: "challenge-body", ordinal });
    });
    const sourceDiagramNodes = slide.nodes.filter((node) => node.visible && !footerNode(node) && (
      node.kind === "shape"
      || node.kind === "connector"
      || (node.role === "caption" && node.sourceFrame.x < inches(6.2) && node.sourceFrame.y >= inches(4.1) && node.sourceFrame.y < inches(6.75))
    ));
    if (sourceDiagramNodes.length >= 4) {
      generatedFigureTreatments.push({
        id: `studio-auto-evidence-${slideNumber}`,
        nodeIds: sourceDiagramNodes.map((node) => node.id).slice(0, 30),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: "Preserve the dense source technical diagram and every internal label and relationship as one evidence unit.",
        informationInventory: ["All source diagram nodes, connectors, labels, and relative relationships"],
        invariants: ["Do not change source labels, topology, sequence, arrows, values, or relationships."],
        rationale: "The source figure carries more technical meaning than can be safely inferred from shape geometry alone.",
        groupFrame: frame(.47, 4.48, 4.50, 2.02),
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
    }
    const visual = content.find((node) => meaningfulImage(node));
    if (visual) placements.set(visual.id, frame(5.30, 4.48, 7.56, 1.55));
    const reserved = new Set([assertion?.id, intro?.id, ...challenges.map((node) => node.id), visual?.id, ...generatedFigureTreatments.flatMap((treatment) => treatment.nodeIds)].filter((id): id is string => Boolean(id)));
    const remaining = [...content, ...captions].filter((node) => !reserved.has(node.id));
    const evidenceCaptions = remaining.filter((node) => node.kind === "text").sort((left, right) => left.sourceFrame.x - right.sourceFrame.x);
    const leftCaption = evidenceCaptions[0];
    const rightCaption = evidenceCaptions.at(-1);
    if (leftCaption) placements.set(leftCaption.id, frame(.47, 6.53, 4.50, .30));
    if (rightCaption && rightCaption.id !== leftCaption?.id) placements.set(rightCaption.id, frame(5.30, 6.10, 7.56, .42));
    for (const [id, value] of stack(remaining.filter((node) => !evidenceCaptions.includes(node)), frame(5.30, 6.10, 7.56, .42), 4)) placements.set(id, value);
  } else if (recipe === "ornl-title-process-flow") {
    const visuals = content.filter(meaningfulImage).sort((left, right) => left.sourceFrame.y - right.sourceFrame.y || left.sourceFrame.x - right.sourceFrame.x);
    const labelPool = captions.filter((node) => node.kind === "text").sort((left, right) => left.sourceFrame.y - right.sourceFrame.y || left.sourceFrame.x - right.sourceFrame.x);
    const unusedLabels = new Set(labelPool.map((node) => node.id));
    const pairedLabels = new Map<string, StudioWebNode>();
    for (const visual of [...visuals].sort((left, right) => left.zIndex - right.zIndex)) {
      const label = labelPool.filter((node) => unusedLabels.has(node.id) && node.zIndex > visual.zIndex).sort((left, right) => left.zIndex - right.zIndex)[0]
        ?? labelPool.filter((node) => unusedLabels.has(node.id)).sort((left, right) => Math.abs(left.sourceFrame.y - visual.sourceFrame.y) - Math.abs(right.sourceFrame.y - visual.sourceFrame.y))[0];
      if (label) {
        pairedLabels.set(visual.id, label);
        unusedLabels.delete(label.id);
      }
    }
    const stageText = content.filter((node) => node.kind === "text" && node.sourceBinding !== "semantic-atom").sort((left, right) => left.zIndex - right.zIndex);
    const firstStage = stageText[0];
    const secondStage = stageText[1];
    const outputVisual = secondStage ? visuals.filter((node) => node.zIndex > secondStage.zIndex).sort((left, right) => left.zIndex - right.zIndex)[0] : visuals.at(-1);
    const inputVisuals = visuals.filter((node) => node.id !== outputVisual?.id).sort((left, right) => left.zIndex - right.zIndex).slice(0, 4);
    inputVisuals.forEach((visual, ordinal) => {
      const column = ordinal % 2;
      const row = Math.floor(ordinal / 2);
      const x = .63 + column * 3.33;
      const y = 1.42 + row * 1.12;
      const groupId = `studio-process-input-${slideNumber}-${ordinal + 1}`;
      placements.set(visual.id, contained(visual, frame(x, y + .13, .42, .52)));
      components.set(visual.id, { groupId, role: "process-icon", ordinal });
      const label = pairedLabels.get(visual.id);
      if (label) {
        placements.set(label.id, frame(x + .62, y + .08, 2.50, .66));
        components.set(label.id, { groupId, role: "process-input", ordinal });
      }
    });
    const atomSupport = content.filter((node) => node.kind === "text" && node.sourceBinding === "semantic-atom").sort((left, right) => left.sourceTextOrder - right.sourceTextOrder);
    atomSupport.forEach((node, ordinal) => components.set(node.id, { groupId: `studio-process-support-${slideNumber}-${ordinal + 1}`, role: "supporting-copy", ordinal }));
    for (const [id, value] of stack(atomSupport, frame(7.15, 1.42, 5.71, 5.12), 24)) placements.set(id, value);
    if (firstStage) {
      placements.set(firstStage.id, frame(.47, 3.78, 6.72, .58));
      components.set(firstStage.id, { groupId: `studio-process-stage-${slideNumber}-1`, role: "process-stage", ordinal: 0 });
    }
    if (secondStage) {
      placements.set(secondStage.id, frame(.47, 4.62, 6.72, .72));
      components.set(secondStage.id, { groupId: `studio-process-stage-${slideNumber}-2`, role: "process-stage", ordinal: 1 });
    }
    const outputLabel = outputVisual ? pairedLabels.get(outputVisual.id) : undefined;
    if (outputLabel) {
      placements.set(outputLabel.id, frame(.47, 5.62, 6.72, .66));
      components.set(outputLabel.id, { groupId: `studio-process-output-${slideNumber}`, role: "process-output", ordinal: 0 });
    }
    if (outputVisual) {
      placements.set(outputVisual.id, contained(outputVisual, frame(6.53, 5.72, .42, .46)));
      components.set(outputVisual.id, { groupId: `studio-process-output-${slideNumber}`, role: "process-icon", ordinal: 0 });
    }
  } else if (recipe === "ornl-title-steps-evidence") {
    const visual = content.find((node) => meaningfulImage(node) || node.kind === "table");
    const steps = content.filter((node) => node.id !== visual?.id && node.kind === "text").sort((left, right) => left.zIndex - right.zIndex);
    if (steps.length === 2) {
      placements.set(steps[0].id, frame(.64, 1.42, 4.58, .72));
      placements.set(steps[1].id, frame(.64, 2.28, 4.58, 2.18));
      components.set(steps[0].id, { groupId: `studio-step-${slideNumber}-1`, role: "step-heading", ordinal: 0 });
      components.set(steps[1].id, { groupId: `studio-step-${slideNumber}-2`, role: "step-body", ordinal: 1 });
    } else {
      const gap = .22;
      const rowHeight = Math.max(.78, Math.min(2.15, (5.18 - gap * Math.max(0, steps.length - 1)) / Math.max(1, steps.length)));
      steps.forEach((node, ordinal) => {
        placements.set(node.id, frame(.64, 1.34 + ordinal * (rowHeight + gap), 4.58, rowHeight));
        components.set(node.id, { groupId: `studio-step-${slideNumber}-${ordinal + 1}`, role: ordinal === 0 || node.style.fontWeight >= 600 ? "step-heading" : "step-body", ordinal });
      });
    }
    if (visual) placements.set(visual.id, contained(visual, frame(5.70, 1.20, 7.16, 5.42)));
    for (const [id, value] of stack(captions, frame(5.70, 6.48, 7.16, .24), 4)) placements.set(id, value);
  } else if (recipe === "ornl-title-labeled-figure-grid") {
    const visuals = content.filter(meaningfulImage).sort((left, right) => left.sourceFrame.y - right.sourceFrame.y || left.sourceFrame.x - right.sourceFrame.x);
    const labels = captions.filter((node) => node.kind === "text");
    const fieldConnectors = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "connector" && !footerNode(node));
    const fieldShapes = slide.nodes.filter((node) => node.visible && !node.locked && node.kind === "shape" && /arrow/i.test(node.name) && !footerNode(node));
    const unused = new Set(labels.map((node) => node.id));
    const distance = (node: StudioWebNode, visual: StudioWebNode) => Math.abs((node.sourceFrame.y + node.sourceFrame.height / 2) - (visual.sourceFrame.y + visual.sourceFrame.height / 2));
    const takeNearest = (visual: StudioWebNode, role: "label" | "caption") => {
      const node = labels.filter((candidate) => unused.has(candidate.id) && candidate.role === role).sort((left, right) => distance(left, visual) - distance(right, visual))[0];
      if (node) unused.delete(node.id);
      return node;
    };
    if (visuals.length >= 2 && fieldConnectors.length >= 1) {
      const annotations = labels.filter(shortFigureAnnotation);
      const diagramNodes = [...visuals, ...fieldConnectors, ...fieldShapes, ...annotations];
      const diagramIds = new Set(diagramNodes.map((node) => node.id));
      const prose = [...content, ...captions].filter((node) => !diagramIds.has(node.id) && !meaningfulImage(node));
      const sourceDiagramFrame = unionStudioFrames(diagramNodes.map((node) => node.sourceFrame));
      const lead = prose.filter((node) => node.sourceFrame.y < sourceDiagramFrame.y + inches(.12));
      const tail = prose.filter((node) => !lead.includes(node));
      const leadHeight = lead.length ? narrativeHeightInches(lead, 12.39, .56, .92) : 0;
      const tailHeight = tail.length ? narrativeHeightInches(tail, 12.39, .46, .76) : 0;
      if (lead.length) for (const [id, value] of stack(lead, frame(.47, contentTop, 12.39, leadHeight), 5)) placements.set(id, value);
      const target = frame(.47, contentTop + leadHeight + (lead.length ? .12 : 0), 12.39, Math.max(1.35, contentHeight - leadHeight - tailHeight - (lead.length ? .12 : 0) - (tail.length ? .12 : 0)));
      for (const [id, value] of scaleSourceGroup(diagramNodes, target)) placements.set(id, value);
      annotations.forEach((node, ordinal) => {
        const placed = placements.get(node.id);
        if (placed) {
          const horizontalAllowance = points(5);
          const left = Math.max(target.x, placed.x - horizontalAllowance);
          const right = Math.min(target.x + target.width, placed.x + placed.width + horizontalAllowance);
          placements.set(node.id, { ...placed, x: left, width: Math.max(placed.width, right - left) });
        }
        components.set(node.id, { groupId: `studio-figure-field-${slideNumber}`, role: "technical-annotation", ordinal });
      });
      if (tail.length) for (const [id, value] of stack(tail, frame(.47, contentBottom - tailHeight, 12.39, tailHeight), 5)) placements.set(id, value);
      generatedFigureTreatments.push({
        id: `studio-auto-figure-field-${slideNumber}`,
        nodeIds: diagramNodes.map((node) => node.id),
        mode: "preserve-as-unit",
        verificationStatus: "source-locked",
        intentSummary: "Preserve the multi-image technical callout field as one relationship-bearing evidence unit.",
        informationInventory: ["All source images", "All short figure annotations", "All source connectors and arrows between the images"],
        invariants: ["Do not detach cross-image arrows or alter source labels, values, sequence, topology, or technical meaning."],
        rationale: "The arrows and labels cross between multiple images, so Studio asks PowerPoint to render the isolated native evidence objects as one exact unit. A semantic redraw is allowed only after its labels, sequence, topology, and technical intent have been independently verified.",
        groupFrame: target,
        lockAspectRatio: true,
        relationshipPolicy: "preserve-internal",
      });
      diagramNodes.forEach((node) => unused.delete(node.id));
    } else if (visuals.length === 2) {
      const narrative = content.filter((node) => !meaningfulImage(node));
      const narrativeHeight = narrative.length ? narrativeHeightInches(narrative, 12.39, 1.10, 1.62) : 0;
      const visualBottom = contentBottom - narrativeHeight - (narrative.length ? .16 : 0);
      visuals.forEach((visual, ordinal) => {
        const x = ordinal === 0 ? .47 : 6.78;
        const groupId = `studio-figure-${slideNumber}-${ordinal + 1}`;
        const label = takeNearest(visual, "label");
        const caption = takeNearest(visual, "caption");
        const hasTwoHeaders = Boolean(label && caption);
        const headerHeight = hasTwoHeaders ? .62 : label || caption ? .34 : 0;
        const visualField = frame(x, contentTop + headerHeight + .08, 5.61, Math.max(.85, visualBottom - contentTop - headerHeight - .08));
        const inset = points(8);
        placements.set(visual.id, contained(visual, { ...visualField, x: visualField.x + inset, y: visualField.y + inset, width: visualField.width - inset * 2, height: visualField.height - inset * 2 }));
        components.set(visual.id, { groupId, role: "figure-media", ordinal, frame: visualField });
        if (label) {
          placements.set(label.id, frame(x, contentTop, 5.61, .26));
          components.set(label.id, { groupId, role: "figure-label", ordinal });
        }
        if (caption) {
          placements.set(caption.id, frame(x, contentTop + (label ? .32 : 0), 5.61, label ? .28 : .34));
          components.set(caption.id, { groupId, role: "figure-caption", ordinal });
        }
      });
      if (narrative.length) for (const [id, value] of stack(narrative, frame(.47, visualBottom + .16, 12.39, narrativeHeight), 7)) placements.set(id, value);
    } else {
      const rowGap = .24;
      const rowHeight = Math.max(1.05, (5.42 - rowGap * Math.max(0, visuals.length - 1)) / Math.max(1, visuals.length));
      visuals.forEach((visual, ordinal) => {
        const y = 1.18 + ordinal * (rowHeight + rowGap);
        const groupId = `studio-figure-${slideNumber}-${ordinal + 1}`;
        const label = takeNearest(visual, "label");
        const caption = takeNearest(visual, "caption");
        const visualField = frame(2.16, y, 6.10, rowHeight);
        const inset = points(8);
        placements.set(visual.id, contained(visual, { ...visualField, x: visualField.x + inset, y: visualField.y + inset, width: visualField.width - inset * 2, height: visualField.height - inset * 2 }));
        components.set(visual.id, { groupId, role: "figure-media", ordinal, frame: visualField });
        if (label) {
          placements.set(label.id, frame(.47, y, 1.44, rowHeight));
          components.set(label.id, { groupId, role: "figure-label", ordinal });
        }
        if (caption) {
          placements.set(caption.id, frame(8.55, y, 4.31, rowHeight));
          components.set(caption.id, { groupId, role: "figure-caption", ordinal });
        }
      });
    }
    if (!(visuals.length >= 2 && fieldConnectors.length >= 1)) {
      for (const [id, value] of stack(labels.filter((node) => unused.has(node.id)), frame(8.55, 1.18, 4.31, 5.42), 8)) placements.set(id, value);
      if (visuals.length !== 2) for (const [id, value] of stack(content.filter((node) => !meaningfulImage(node)), frame(.47, 6.28, 12.39, .36), 4)) placements.set(id, value);
    }
  } else if (recipe === "ornl-title-card-grid") {
    const bodies = content.filter((node) => node.kind === "text" && node.role === "body").sort((left, right) => left.zIndex - right.zIndex);
    const groupNodeIds = new Set<string>();
    const semanticSections: Array<{ heading: StudioWebNode; items: StudioWebNode[] }> = [];
    for (const body of bodies.filter((node) => node.sourceBinding === "semantic-atom")) {
      const paragraph = body.sourceParagraphs?.[0];
      if (paragraph && !paragraph.bullet && paragraph.level === 0) semanticSections.push({ heading: body, items: [] });
      else semanticSections.at(-1)?.items.push(body);
    }
    const useSemanticSections = semanticSections.length >= 2 && semanticSections.length <= 4 && semanticSections.every((section) => section.items.length > 0);
    if (useSemanticSections) {
      semanticSections.forEach((section, ordinal) => {
        const groupId = `studio-card-${slideNumber}-${ordinal + 1}`;
        const card = { ...cardFrame(ordinal, semanticSections.length), y: inches(1.46), height: inches(4.82) };
        placements.set(section.heading.id, frame(emuInches(card.x) + .30, emuInches(card.y) + .28, emuInches(card.width) - .60, .48));
        components.set(section.heading.id, { groupId, role: "card-kicker", ordinal, frame: card });
        groupNodeIds.add(section.heading.id);
        const itemRegionHeight = Math.min(emuInches(card.height) - 1.24, Math.max(.72, section.items.length * .78));
        for (const [id, value] of stack(section.items, frame(emuInches(card.x) + .30, emuInches(card.y) + .98, emuInches(card.width) - .60, itemRegionHeight), 14)) placements.set(id, value);
        section.items.forEach((item) => {
          components.set(item.id, { groupId, role: "card-body", ordinal, frame: card });
          groupNodeIds.add(item.id);
        });
      });
    }
    let priorBodyZ = title?.zIndex ?? -Infinity;
    if (!useSemanticSections) bodies.forEach((body, ordinal) => {
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
      const hasHeader = Boolean(kicker || heading);
      placements.set(body.id, frame(emuInches(card.x) + .27, emuInches(card.y) + (hasHeader ? .72 : .38), emuInches(card.width) - .49, emuInches(card.height) - (hasHeader ? .88 : .40)));
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
    figureTreatments: [
      ...(slide.figureTreatments ?? []).filter((treatment) => !treatment.id.startsWith("studio-auto-")),
      ...generatedFigureTreatments,
    ],
    constraints: [],
    nodes: slide.nodes.map((node) => {
      const component = recipe === "source" || recipe === "template-layout" ? undefined : components.get(node.id);
      if (!placements.has(node.id)) return { ...node, component };
      const nextNode = { ...node, component };
      const nextFrame = placements.get(node.id)!;
      const designedStyle = component ? styleForComponent(nextNode) : styleForDesignedNode(nextNode);
      const nextStyle = recipe === "template-layout" ? { ...designedStyle, color: templateTextColor(layout, nextNode, nextFrame) ?? designedStyle.color } : designedStyle;
      return { ...nextNode, frame: nextFrame, style: fittedStyle(nextNode, nextStyle, nextFrame) };
    }),
    updatedAt: now,
  };
  return { ...workingScene, revision: `${workingScene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`, slides: workingScene.slides.map((item) => item.slideNumber === slideNumber ? nextSlide : item) };
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
  const movedNodes = slide.nodes.map((candidate) => candidate.id === nodeId ? { ...candidate, frame: bounded } : candidate);
  const connectedNodes = movedNodes.map((candidate) => candidate.connector ? { ...candidate, frame: studioConnectorFrame(movedNodes, candidate.connector) } : candidate);
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : { ...item, status: "designed", updatedAt: now, designRationale: "Human-adjusted on the shared Studio web canvas.", constraints: (item.constraints ?? []).filter((constraint) => !constraint.nodeIds.includes(nodeId)), nodes: connectedNodes }),
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

function normalizedTableWeights(values: number[] | undefined, count: number): number[] {
  const safeCount = Math.max(1, count);
  if (!values || values.length !== safeCount || values.some((value) => !Number.isFinite(value) || value <= 0)) return Array.from({ length: safeCount }, () => 1 / safeCount);
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(total) || total <= 0) return Array.from({ length: safeCount }, () => 1 / safeCount);
  return values.map((value) => value / total);
}

export function resolvedStudioTableDesign(node: StudioWebNode): StudioTableDesign {
  if (node.kind !== "table" || !node.table) throw new Error("The requested Studio node is not an editable table.");
  const design = node.table.design;
  return {
    headerRows: Math.max(0, Math.min(node.table.rows, Math.round(design?.headerRows ?? (node.table.rows > 0 ? 1 : 0)))),
    columnWidths: normalizedTableWeights(design?.columnWidths, node.table.columns),
    rowHeights: normalizedTableWeights(design?.rowHeights, node.table.rows),
    borderMode: design?.borderMode ?? "subtle",
    borderColor: /^#[0-9a-f]{6}$/i.test(design?.borderColor ?? "") ? design!.borderColor.toUpperCase() : "#DBDCDB",
    borderWidthPt: Math.max(0, Math.min(6, design?.borderWidthPt ?? .75)),
    defaultPaddingPt: design?.defaultPaddingPt ?? { top: 4, right: 7, bottom: 4, left: 7 },
    cellStyles: (design?.cellStyles ?? []).filter((item) => node.table?.cells.some((cell) => cell.id === item.cellId)),
  };
}

function validateTablePadding(value: StudioTableDesign["defaultPaddingPt"]): StudioTableDesign["defaultPaddingPt"] {
  const next = { top: Number(value.top), right: Number(value.right), bottom: Number(value.bottom), left: Number(value.left) };
  if (Object.values(next).some((item) => !Number.isFinite(item) || item < 0 || item > 36)) throw new Error("Studio table cell padding must be between 0 and 36 pt.");
  return next;
}

function validateTableCellBorders(value: StudioTableCellDesign["borders"]): StudioTableCellDesign["borders"] {
  if (!value) return undefined;
  return Object.fromEntries(Object.entries(value).map(([edge, border]) => {
    if (!border || !["none", "solid", "dash"].includes(border.type)) throw new Error(`Studio table ${edge} border must be none, solid, or dash.`);
    if (!/^#[0-9a-f]{6}$/i.test(border.color)) throw new Error(`Studio table ${edge} border color must be a six-digit hex color.`);
    if (!Number.isFinite(border.widthPt) || border.widthPt < 0 || border.widthPt > 6) throw new Error(`Studio table ${edge} border width must be between 0 and 6 pt.`);
    return [edge, { ...border, color: border.color.toUpperCase(), widthPt: border.type === "none" ? 0 : border.widthPt }];
  })) as StudioTableCellDesign["borders"];
}

function updateStudioTableNode(scene: StudioWebScene, slideNumber: number, nodeId: string, updater: (node: StudioWebNode, design: StudioTableDesign) => StudioTableDesign): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const node = slide?.nodes.find((item) => item.id === nodeId);
  if (!slide || !node?.table || node.kind !== "table") throw new Error("The requested Studio table is not present in the current slide revision.");
  if (node.locked) throw new Error(`${node.name} is locked because its PowerPoint representation is not safely editable.`);
  const design = updater(node, resolvedStudioTableDesign(node));
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      status: "designed",
      updatedAt: now,
      designRationale: "Human- or AI-refined with the source-bound Studio table component editor.",
      nodes: item.nodes.map((candidate) => candidate.id === nodeId && candidate.table ? { ...candidate, table: { ...candidate.table, design } } : candidate),
    }),
  };
}

export function updateStudioTableDesign(scene: StudioWebScene, slideNumber: number, nodeId: string, patch: Partial<Pick<StudioTableDesign, "headerRows" | "columnWidths" | "rowHeights" | "borderMode" | "borderColor" | "borderWidthPt" | "defaultPaddingPt">>): StudioWebScene {
  return updateStudioTableNode(scene, slideNumber, nodeId, (node, design) => {
    if (patch.headerRows !== undefined && (!Number.isInteger(patch.headerRows) || patch.headerRows < 0 || patch.headerRows > (node.table?.rows ?? 0))) throw new Error("Studio table header rows must fit within the current table structure.");
    if (patch.borderMode !== undefined && !["none", "subtle", "full"].includes(patch.borderMode)) throw new Error("Studio table borders must be none, subtle, or full.");
    if (patch.borderColor !== undefined && !/^#[0-9a-f]{6}$/i.test(patch.borderColor)) throw new Error("Studio table border color must be a six-digit hex color.");
    if (patch.borderWidthPt !== undefined && (!Number.isFinite(patch.borderWidthPt) || patch.borderWidthPt < 0 || patch.borderWidthPt > 6)) throw new Error("Studio table border width must be between 0 and 6 pt.");
    return {
      ...design,
      ...patch,
      columnWidths: patch.columnWidths ? normalizedTableWeights(patch.columnWidths, node.table!.columns) : design.columnWidths,
      rowHeights: patch.rowHeights ? normalizedTableWeights(patch.rowHeights, node.table!.rows) : design.rowHeights,
      borderColor: patch.borderColor?.toUpperCase() ?? design.borderColor,
      defaultPaddingPt: patch.defaultPaddingPt ? validateTablePadding(patch.defaultPaddingPt) : design.defaultPaddingPt,
    };
  });
}

function resizedTableWeights(current: number[], index: number, requestedEmu: number, totalEmu: number, minimumEmu: number): number[] {
  if (!Number.isInteger(index) || index < 0 || index >= current.length) throw new Error("Choose an existing Studio table row or column.");
  if (!Number.isFinite(requestedEmu) || requestedEmu < minimumEmu) throw new Error("The requested Studio table row or column is below its minimum usable size.");
  const minRatio = Math.min(.2, minimumEmu / Math.max(totalEmu, 1));
  const desired = Math.max(minRatio, Math.min(1 - minRatio * Math.max(0, current.length - 1), requestedEmu / Math.max(totalEmu, 1)));
  const remainder = Math.max(0, 1 - desired);
  const currentRemainder = Math.max(.000001, current.reduce((sum, value, candidate) => candidate === index ? sum : sum + value, 0));
  return current.map((value, candidate) => candidate === index ? desired : value / currentRemainder * remainder);
}

export function resizeStudioTableColumn(scene: StudioWebScene, slideNumber: number, nodeId: string, column: number, widthEmu: number): StudioWebScene {
  return updateStudioTableNode(scene, slideNumber, nodeId, (node, design) => ({ ...design, columnWidths: resizedTableWeights(design.columnWidths, column - 1, widthEmu, node.frame.width, inches(.35)) }));
}

export function resizeStudioTableRow(scene: StudioWebScene, slideNumber: number, nodeId: string, row: number, heightEmu: number): StudioWebScene {
  return updateStudioTableNode(scene, slideNumber, nodeId, (node, design) => ({ ...design, rowHeights: resizedTableWeights(design.rowHeights, row - 1, heightEmu, node.frame.height, inches(.18)) }));
}

export function updateStudioTableCellDesign(scene: StudioWebScene, slideNumber: number, nodeId: string, cellId: string, patch: Omit<StudioTableCellDesign, "cellId">): StudioWebScene {
  return updateStudioTableNode(scene, slideNumber, nodeId, (node, design) => {
    const sourceCell = node.table?.cells.find((cell) => cell.id === cellId);
    if (!sourceCell) throw new Error("The requested source-bound Studio table cell is unavailable.");
    if (patch.fill !== undefined && !/^#[0-9a-f]{6}$/i.test(patch.fill)) throw new Error("Studio table cell fill must be a six-digit hex color.");
    if (patch.color !== undefined && !/^#[0-9a-f]{6}$/i.test(patch.color)) throw new Error("Studio table cell text color must be a six-digit hex color.");
    if (patch.fontSizePt !== undefined && (!Number.isFinite(patch.fontSizePt) || patch.fontSizePt < 10 || patch.fontSizePt > 40)) throw new Error("Studio table cell text size must be between 10 and 40 pt.");
    if (patch.fontWeight !== undefined && ![400, 600, 700].includes(patch.fontWeight)) throw new Error("Studio table cell font weight must be 400, 600, or 700.");
    if (patch.textAlign !== undefined && !["left", "center", "right"].includes(patch.textAlign)) throw new Error("Studio table cell alignment must be left, center, or right.");
    if (patch.verticalAlign !== undefined && !["top", "middle", "bottom"].includes(patch.verticalAlign)) throw new Error("Studio table cell vertical alignment must be top, middle, or bottom.");
    const sourceSemanticFill = /^#([0-9a-f]{6})$/i.exec(sourceCell.fill ?? "")?.[0]?.toUpperCase();
    if (sourceCell.semanticColorRole && sourceSemanticFill && patch.fill !== undefined && patch.fill.toUpperCase() !== sourceSemanticFill) throw new Error(`Cell ${cellId} uses the source-significant ${sourceCell.semanticColorRole} color. Preserve that fill unless a human explicitly changes the source meaning.`);
    const prior = design.cellStyles.find((item) => item.cellId === cellId);
    const nextPatch = {
      ...patch,
      ...(patch.fill !== undefined ? { fill: patch.fill.toUpperCase() } : {}),
      ...(patch.color !== undefined ? { color: patch.color.toUpperCase() } : {}),
      ...(patch.paddingPt !== undefined ? { paddingPt: validateTablePadding(patch.paddingPt) } : {}),
      ...(patch.borders !== undefined ? { borders: { ...prior?.borders, ...validateTableCellBorders(patch.borders) } } : {}),
    };
    const next = { ...prior, ...nextPatch, cellId };
    return { ...design, cellStyles: [...design.cellStyles.filter((item) => item.cellId !== cellId), next] };
  });
}

export function studioConnectorAttachmentPoint(node: StudioWebNode, side: StudioConnectorDesign["fromSide"]): { x: number; y: number } {
  if (side === "top") return { x: node.frame.x + node.frame.width / 2, y: node.frame.y };
  if (side === "right") return { x: node.frame.x + node.frame.width, y: node.frame.y + node.frame.height / 2 };
  if (side === "bottom") return { x: node.frame.x + node.frame.width / 2, y: node.frame.y + node.frame.height };
  if (side === "left") return { x: node.frame.x, y: node.frame.y + node.frame.height / 2 };
  return { x: node.frame.x + node.frame.width / 2, y: node.frame.y + node.frame.height / 2 };
}

export function studioConnectorFrame(nodes: StudioWebNode[], design: StudioConnectorDesign): StudioWebFrame {
  const from = nodes.find((node) => node.id === design.fromNodeId);
  const to = nodes.find((node) => node.id === design.toNodeId);
  if (!from || !to) throw new Error("A verified Studio connector has a stale endpoint binding.");
  const start = studioConnectorAttachmentPoint(from, design.fromSide);
  const end = studioConnectorAttachmentPoint(to, design.toSide);
  return { x: Math.min(start.x, end.x), y: Math.min(start.y, end.y), width: Math.max(points(.25), Math.abs(end.x - start.x)), height: Math.max(points(.25), Math.abs(end.y - start.y)), rotation: 0 };
}

export function updateStudioConnectorDesign(scene: StudioWebScene, slideNumber: number, connectorNodeId: string, design: StudioConnectorDesign): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const connector = slide?.nodes.find((node) => node.id === connectorNodeId);
  if (!slide || !connector || connector.kind !== "connector") throw new Error("Choose a connector node from the current Studio slide.");
  if (connector.locked) throw new Error(`${connector.name} is locked because its PowerPoint representation is not safely editable.`);
  const from = slide.nodes.find((node) => node.id === design.fromNodeId && node.visible);
  const to = slide.nodes.find((node) => node.id === design.toNodeId && node.visible);
  if (!from || !to || from.id === to.id || [from.id, to.id].includes(connector.id)) throw new Error("A verified connector must join two different visible non-connector Studio nodes.");
  const treatment = slide.figureTreatments.find((item) => item.nodeIds.includes(connector.id) && item.nodeIds.includes(from.id) && item.nodeIds.includes(to.id));
  if (!treatment || treatment.verificationStatus !== "verified" || treatment.relationshipPolicy !== "editable-diagram") throw new Error("Direct connector authoring requires one verified editable-diagram figure treatment containing the connector and both endpoint nodes.");
  if (!/^#[0-9a-f]{6}$/i.test(design.stroke)) throw new Error("Studio connector color must be a six-digit hex color.");
  if (!Number.isFinite(design.widthPt) || design.widthPt < .25 || design.widthPt > 8) throw new Error("Studio connector width must be between 0.25 and 8 pt.");
  if (!["solid", "dash", "dashDot"].includes(design.dash)) throw new Error("Studio connector dash must be solid, dash, or dash-dot.");
  if (!["top", "right", "bottom", "left", "center"].includes(design.fromSide) || !["top", "right", "bottom", "left", "center"].includes(design.toSide)) throw new Error("Studio connector endpoints must use a supported node attachment side.");
  const normalized: StudioConnectorDesign = { ...design, stroke: design.stroke.toUpperCase(), verificationStatus: "verified" };
  const connectorFrame = studioConnectorFrame(slide.nodes, normalized);
  const relationships = [
    ...(treatment.relationships ?? []).filter((item) => item.fromNodeId !== connector.id || !["connects-from", "connects-to"].includes(item.kind)),
    { fromNodeId: connector.id, toNodeId: from.id, kind: "connects-from" as const },
    { fromNodeId: connector.id, toNodeId: to.id, kind: "connects-to" as const },
  ];
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      status: "designed",
      updatedAt: now,
      designRationale: "Human- or AI-authored verified connector relationships on the shared Studio canvas.",
      nodes: item.nodes.map((node) => node.id === connector.id ? { ...node, connector: normalized, frame: connectorFrame, style: { ...node.style, borderColor: normalized.stroke, borderWidthPt: normalized.widthPt } } : node),
      figureTreatments: item.figureTreatments.map((candidate) => candidate.id === treatment.id ? { ...candidate, relationships } : candidate),
    }),
  };
}

function unionStudioFrames(frames: StudioWebFrame[]): StudioWebFrame {
  const x = Math.min(...frames.map((frame) => frame.x));
  const y = Math.min(...frames.map((frame) => frame.y));
  const right = Math.max(...frames.map((frame) => frame.x + frame.width));
  const bottom = Math.max(...frames.map((frame) => frame.y + frame.height));
  return { x, y, width: right - x, height: bottom - y, rotation: 0 };
}

export function updateStudioFigureTreatment(scene: StudioWebScene, slideNumber: number, treatment: StudioFigureTreatment): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
  const nodeIds = [...new Set(treatment.nodeIds)];
  if (!treatment.id.trim() || nodeIds.length === 0 || nodeIds.length > 30) throw new Error("A figure treatment requires an ID and 1–30 unique Studio node IDs.");
  if (!["preserve-as-unit", "preserve-and-frame", "hybrid-rebuild", "redraw-candidate"].includes(treatment.mode)) throw new Error("Choose a supported figure-treatment mode.");
  if (!["source-locked", "needs-content-review", "verified"].includes(treatment.verificationStatus)) throw new Error("Choose a supported figure verification status.");
  const nodes = nodeIds.map((id) => slide.nodes.find((node) => node.id === id));
  if (nodes.some((node) => !node)) throw new Error("A figure-treatment node is not present in the current Studio slide revision.");
  if (!nodes.some((node) => node && ["image", "native-object", "shape", "connector"].includes(node.kind))) throw new Error("A figure treatment must contain at least one source visual, native object, shape, or connector.");
  if (!treatment.intentSummary.trim() || treatment.informationInventory.length === 0 || treatment.invariants.length === 0 || !treatment.rationale.trim()) throw new Error("A figure treatment requires an intent summary, information inventory, invariants, and rationale.");
  if (treatment.mode === "redraw-candidate" && treatment.verificationStatus === "source-locked") throw new Error("A redraw candidate must remain needs-content-review until its information and relationships are verified.");
  if (treatment.replacementResourceId && treatment.verificationStatus !== "verified") throw new Error("A replacement Resource may be bound only after the figure treatment is verified.");
  if (treatment.focalPoint && ![treatment.focalPoint.x, treatment.focalPoint.y].every((value) => Number.isFinite(value) && value >= 0 && value <= 1)) throw new Error("A figure focal point must use normalized coordinates from 0 to 1.");
  if (treatment.crop) {
    const { left, top, right, bottom } = treatment.crop;
    if (![left, top, right, bottom].every((value) => Number.isFinite(value) && value >= 0 && value < 1) || left + right >= 1 || top + bottom >= 1) throw new Error("Figure crop values must be normalized from 0 to 1 and leave visible source content.");
  }
  if (treatment.relationshipPolicy === "editable-diagram" && treatment.verificationStatus !== "verified") throw new Error("An editable diagram relationship policy requires verified information and relationships.");
  const relationshipKeys = new Set<string>();
  const relationships = (treatment.relationships ?? []).map((relationship) => {
    if (!nodeIds.includes(relationship.fromNodeId) || !nodeIds.includes(relationship.toNodeId)) throw new Error("Figure relationships must connect nodes inside the same treatment.");
    if (relationship.fromNodeId === relationship.toNodeId) throw new Error("A figure node cannot relate to itself.");
    const key = `${relationship.fromNodeId}:${relationship.toNodeId}:${relationship.kind}`;
    if (relationshipKeys.has(key)) throw new Error("Duplicate figure relationships are not allowed.");
    relationshipKeys.add(key);
    return relationship;
  });
  const resolvedNodes = nodes as StudioWebNode[];
  const visualNodes = resolvedNodes.filter((node) => ["image", "native-object", "shape", "connector"].includes(node.kind));
  const defaultRelationshipPolicy = treatment.mode === "hybrid-rebuild" ? "reflow-annotations" : treatment.mode === "redraw-candidate" && treatment.verificationStatus === "verified" ? "editable-diagram" : "preserve-internal";
  const normalized: StudioFigureTreatment = {
    ...treatment,
    id: treatment.id.trim().slice(0, 180),
    nodeIds,
    intentSummary: treatment.intentSummary.trim().slice(0, 1_000),
    informationInventory: treatment.informationInventory.map((item) => item.trim()).filter(Boolean).slice(0, 40),
    invariants: treatment.invariants.map((item) => item.trim()).filter(Boolean).slice(0, 40),
    rationale: treatment.rationale.trim().slice(0, 1_000),
    relationships,
    groupFrame: treatment.groupFrame ?? unionStudioFrames(visualNodes.map((node) => node.frame)),
    focalPoint: treatment.focalPoint ? { x: treatment.focalPoint.x, y: treatment.focalPoint.y } : { x: .5, y: .5 },
    crop: treatment.crop ? { ...treatment.crop } : { left: 0, top: 0, right: 0, bottom: 0 },
    relationshipPolicy: treatment.relationshipPolicy ?? defaultRelationshipPolicy,
    lockAspectRatio: treatment.lockAspectRatio ?? ["preserve-as-unit", "preserve-and-frame"].includes(treatment.mode),
  };
  const affected = new Set(nodeIds);
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      status: "designed",
      updatedAt: now,
      designRationale: `${item.designRationale} Figure treatment: ${normalized.mode}.`.trim().slice(0, 1_000),
      figureTreatments: [...(item.figureTreatments ?? []).filter((candidate) => candidate.id !== normalized.id && !candidate.nodeIds.some((id) => affected.has(id))), normalized],
    }),
  };
}

export function translateStudioFigureTreatment(scene: StudioWebScene, slideNumber: number, treatmentId: string, dx: number, dy: number): StudioWebScene {
  const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
  const treatment = slide?.figureTreatments.find((item) => item.id === treatmentId);
  if (!slide || !treatment) throw new Error("The requested Studio figure treatment is not present in the current slide revision.");
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) throw new Error("A Studio figure translation requires finite offsets.");
  const sourceFrame = treatment.groupFrame ?? unionStudioFrames(treatment.nodeIds.map((id) => slide.nodes.find((node) => node.id === id)?.frame).filter((frame): frame is StudioWebFrame => Boolean(frame)));
  const targetFrame = { ...sourceFrame, x: Math.round(sourceFrame.x + dx), y: Math.round(sourceFrame.y + dy) };
  if (targetFrame.x < 0 || targetFrame.y < 0 || targetFrame.x + targetFrame.width > scene.slideSize.width || targetFrame.y + targetFrame.height > scene.slideSize.height) throw new Error("The complete Studio figure would leave the slide canvas.");
  const affected = new Set(treatment.nodeIds);
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((item) => item.slideNumber !== slideNumber ? item : {
      ...item,
      status: "designed",
      updatedAt: now,
      designRationale: "Human-adjusted one relationship-preserving Studio figure group on the shared web canvas.",
      constraints: (item.constraints ?? []).filter((constraint) => !constraint.nodeIds.some((id) => affected.has(id))),
      figureTreatments: item.figureTreatments.map((candidate) => candidate.id === treatmentId ? { ...candidate, groupFrame: targetFrame } : candidate),
      nodes: item.nodes.map((node) => affected.has(node.id) ? { ...node, frame: { ...node.frame, x: Math.round(node.frame.x + dx), y: Math.round(node.frame.y + dy) } } : node),
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
  return slide.nodes.filter((node) => node.sourceBinding === "editable-object" && node.visible && !node.locked && frameChanged(node.sourceFrame, node.frame)).map((node) => ({
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
  const figureTreatmentIds = new Set((slide.figureTreatments ?? []).map((treatment) => treatment.id));
  const textStyles = slide.nodes.filter((node) => node.sourceBinding === "editable-object" && node.visible && !node.locked && node.kind === "text").map((node) => ({
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
  const decorations: VisualDesignRequest["decorations"] = studioGeneratedComponents(slide).map((component) => {
    const isFigureTreatment = [...figureTreatmentIds].some((id) => component.id === `${id}-surface` || component.id === `${id}-accent`);
    return {
      id: component.id,
      name: isFigureTreatment ? (component.id.endsWith("-accent") ? "ORNL technical figure accent" : "ORNL technical figure frame") : component.id.includes("title-rule") ? "ORNL title rule" : component.id.includes("footer-rule") ? "ORNL footer rule" : component.id.endsWith("-accent") ? "ORNL card accent" : "ORNL content card",
      geometry: scaleFrame(component.frame, scene.slideSize, scene.sourceSlideSize),
      fillColor: component.fillColor,
      lineColor: component.lineColor,
      lineWidthPt: component.lineWidthPt,
      behindContent: component.behindContent,
      rationale: isFigureTreatment ? "Frame the complete source-locked technical evidence unit without altering its internal pixels, labels, values, code, arrows, or relationships." : "Use the shared square-cornered ORNL web component rather than a slide-specific decoration.",
      author,
    };
  });
  return { slideNumber, textStyles, decorations };
}

export function studioSceneNeedsRebuild(deck: DeckJob): boolean {
  return Boolean(deck.audit && deck.scene) && (!deck.studioScene || deck.studioScene.schema !== STUDIO_WEB_SCENE_SCHEMA || deck.studioScene.version !== STUDIO_WEB_SCENE_VERSION || deck.studioScene.deckId !== deck.id || deck.studioScene.sourceSha256 !== deck.sourceSha256);
}

export function planStudioExportBuild(scene: StudioWebScene): { preservedSourceSlideNumbers: number[]; freshCompositionSlideNumbers: number[]; nativeTemplateSlideNumbers: number[] } {
  const preservedSourceSlideNumbers: number[] = [];
  const freshCompositionSlideNumbers: number[] = [];
  const nativeTemplateSlideNumbers: number[] = [];
  for (const slide of [...scene.slides].sort((left, right) => left.slideNumber - right.slideNumber)) {
    if (slide.status !== "designed" || slide.recipe === "source") preservedSourceSlideNumbers.push(slide.slideNumber);
    else freshCompositionSlideNumbers.push(slide.slideNumber);
  }
  return { preservedSourceSlideNumbers, freshCompositionSlideNumbers, nativeTemplateSlideNumbers };
}
