import PptxGenJS from "pptxgenjs";
import type { SlideRenderCatalog, TemplateCatalog, TemplatePreviewElement } from "./template-catalog";
import type { StudioWebNode, StudioWebScene } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { resolvedStudioTableDesign, studioConnectorAttachmentPoint, studioGeneratedComponents } from "./studio-web-scene";
import { materializeStudioTableContinuationSlides } from "./studio-table-workflow";

const EMU_PER_INCH = 914_400;

export interface StudioCompositionExportResult {
  bytes: Uint8Array;
  slideCount: number;
  outputSlides: StudioCompositionOutputSlide[];
  textNodeCount: number;
  tableCount: number;
  imageCount: number;
  ignoredSourceFurnitureCount: number;
  generatedComponentCount: number;
  warnings: string[];
}

export interface StudioCompositionOutputSlide {
  outputSlideNumber: number;
  sourceSlideNumber: number;
  continuation?: {
    tableNodeId: string;
    segmentOrdinal: number;
    segmentCount: number;
    bodyRowStart: number;
    bodyRowEnd: number;
    repeatedHeaderRows: number;
  };
}

export interface StudioCompositionExportOptions {
  catalog?: SlideRenderCatalog;
  templateCatalog?: TemplateCatalog;
  sourceSlideRasters?: Record<number, { data: string; width: number; height: number }>;
  sourceFigureRasters?: Record<string, { data: string; width: number; height: number }>;
  sourceSlideText?: Record<number, string>;
  templateLayoutRasters?: Record<string, { data: string; width: number; height: number }>;
  /**
   * When present, the caller will attach real ORNL native layouts after this
   * editable composition is written. Use the named neutral layout for shared
   * recipes and suppress slide-local copies of template artwork.
   */
  nativeTemplateLayoutBaseId?: string;
  strict?: boolean;
  title?: string;
}

function inches(value: number): number {
  return value / EMU_PER_INCH;
}

function hex(value: string | undefined, fallback: string): string {
  const normalized = (candidate: string | undefined): string | undefined => {
    const direct = /^#([0-9a-f]{6})$/i.exec(candidate ?? "")?.[1];
    if (direct) return direct.toUpperCase();
    const tagged = /^(?:RGB|srgb):([0-9a-f]{6})$/i.exec(candidate ?? "")?.[1];
    return tagged?.toUpperCase();
  };
  return normalized(value) ?? normalized(fallback) ?? "000000";
}

function margins(node: StudioWebNode): [number, number, number, number] {
  return [node.style.paddingPt.top, node.style.paddingPt.right, node.style.paddingPt.bottom, node.style.paddingPt.left];
}

function editableText(node: StudioWebNode): string | PptxGenJS.TextProps[] {
  const text = node.text ?? "";
  const paragraphs = node.sourceParagraphs?.filter((paragraph) => paragraph.text.length > 0) ?? [];
  if (paragraphs.length <= 1 && !paragraphs[0]?.bullet && !/[\uE000-\uF8FF]/.test(text)) return text;
  const runs: PptxGenJS.TextProps[] = [];
  const source = paragraphs.length ? paragraphs : [{ text, bullet: false, level: 0 }];
  const paragraphSpaceAfter = source.length >= 7 ? 2 : source.some((paragraph) => paragraph.bullet) ? 3 : PRESENTATION_DESIGN_STANDARD.componentSystem.paragraph.bodySpaceAfterPt;
  source.forEach((paragraph, paragraphIndex) => {
    const parts = paragraph.text.split(/([\uE000-\uF8FF])/).filter(Boolean);
    parts.forEach((value, partIndex) => {
      const finalPart = partIndex === parts.length - 1;
      runs.push({
        text: value,
        options: {
          fontFace: /[\uE000-\uF8FF]/.test(value) ? "Wingdings" : "Aptos",
          // PptxGenJS treats an explicit `type: "bullet"` as a numbering
          // variant in rich-text runs. Omitting `type` produces the native
          // a:buChar paragraph PowerPoint expects while retaining our indent.
          bullet: partIndex === 0 && paragraph.bullet ? { indent: 18 + Math.max(0, paragraph.level ?? 0) * 14 } : undefined,
          breakLine: finalPart && paragraphIndex < source.length - 1,
          paraSpaceAfter: finalPart && paragraphIndex < source.length - 1 ? paragraphSpaceAfter : undefined,
        },
      });
    });
  });
  return runs;
}

function editableTableText(cell: NonNullable<StudioWebNode["table"]>["cells"][number]): string | PptxGenJS.TextProps[] {
  const runs = cell.textRuns?.length ? cell.textRuns : undefined;
  if (!runs) return cell.text;
  const paragraphStarts = new Set<number>();
  if (!cell.runBreaksBefore) {
    let cursor = 0;
    for (const count of cell.paragraphRunCounts ?? []) {
      if (cursor > 0 && count > 0) paragraphStarts.add(cursor);
      cursor += count;
    }
  }
  return runs.map((text, index) => ({
    text,
    options: {
      // PptxGenJS applies breakLine after the current run. Our inventory stores
      // a break before the following run, so consult index + 1 here.
      breakLine: cell.runBreaksBefore ? cell.runBreaksBefore[index + 1] !== undefined && cell.runBreaksBefore[index + 1] !== "none" : paragraphStarts.has(index + 1),
      fontFace: /[\uE000-\uF8FF]/.test(text) ? "Wingdings" : "Aptos",
    },
  }));
}

function tableRows(node: StudioWebNode): PptxGenJS.TableRow[] {
  if (!node.table) return [];
  const design = resolvedStudioTableDesign(node);
  const rows: PptxGenJS.TableRow[] = Array.from({ length: node.table.rows }, () => []);
  for (const cell of [...node.table.cells].sort((a, b) => a.row - b.row || a.column - b.column)) {
    const row = Math.max(0, cell.row - 1);
    if (!rows[row]) continue;
    const header = cell.row <= design.headerRows;
    const cellDesign = design.cellStyles.find((item) => item.cellId === cell.id);
    const padding = cellDesign?.paddingPt ?? design.defaultPaddingPt;
    const border = design.borderMode === "none"
      ? { type: "none" as const, color: hex(design.borderColor, "DBDCDB"), pt: 0 }
      : { type: "solid" as const, color: hex(design.borderColor, "DBDCDB"), pt: design.borderMode === "subtle" ? Math.min(.75, design.borderWidthPt) : design.borderWidthPt };
    const cellBorder = (edge: "top" | "right" | "bottom" | "left"): PptxGenJS.BorderProps => {
      const override = cellDesign?.borders?.[edge];
      if (!override) return border;
      return { type: override.type, color: hex(override.color, design.borderColor), pt: override.type === "none" ? 0 : override.widthPt };
    };
    // PptxGenJS synthesizes hMerge/vMerge continuation cells for colspan/rowspan.
    // Supplying placeholder cells for those occupied grid positions creates an
    // extra physical column and silently shifts later content to the right.
    rows[row].push({
      text: editableTableText(cell),
      options: {
        rowspan: Math.max(1, cell.rowSpan),
        colspan: Math.max(1, cell.columnSpan),
        bold: (cellDesign?.fontWeight ?? (header ? 700 : 400)) >= 600,
        fontSize: cellDesign?.fontSizePt ?? node.style.fontSizePt,
        color: hex(cellDesign?.color, header && !cell.semanticColorRole ? "#FFFFFF" : node.style.color),
        // Source table fills are formatting noise unless the audit assigned a
        // meaning-bearing semantic role. Normalize ordinary body cells to the
        // shared ORNL banding system, while preserving an explicitly styled
        // cell or a semantic source color exactly.
        fill: { color: hex(cellDesign?.fill, cell.semanticColorRole ? cell.fill ?? "#FFFFFF" : header ? "#00454D" : (cell.row % 2 === 0 ? "#F0F2F1" : "#FFFFFF")) },
        align: cellDesign?.textAlign ?? node.style.textAlign,
        valign: cellDesign?.verticalAlign ?? node.style.verticalAlign,
        margin: [padding.top, padding.right, padding.bottom, padding.left],
        border: cellDesign?.borders ? [cellBorder("top"), cellBorder("right"), cellBorder("bottom"), cellBorder("left")] : border,
      },
    });
  }
  return rows;
}

function tableDimensions(node: StudioWebNode): Pick<PptxGenJS.TableProps, "colW" | "rowH"> {
  if (!node.table) return {};
  const design = resolvedStudioTableDesign(node);
  return {
    colW: design.columnWidths.map((weight) => inches(node.frame.width) * weight),
    rowH: design.rowHeights.map((weight) => inches(node.frame.height) * weight),
  };
}

function unsupportedContentNode(node: StudioWebNode): boolean {
  return node.visible && node.kind === "native-object";
}

function normalizedTextOrderValue(value: string | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim();
}

function nodeTextOrderValue(node: StudioWebNode): string {
  if (node.kind === "table" && node.table) return normalizedTextOrderValue(node.table.cells.map((cell) => cell.text).join(" "));
  return normalizedTextOrderValue(node.text);
}

function sourceTextOrderIndices(nodes: StudioWebNode[], sourceText: string): Map<string, number> {
  const indices = new Map<string, number>();
  const occupied: Array<{ start: number; end: number }> = [];
  const textual = nodes.map((node) => ({ node, text: nodeTextOrderValue(node) })).filter((item) => item.text);
  // Claim the longest exact source spans first. This prevents a short figure
  // label such as "POI" or "Frequency response" from binding to an earlier
  // occurrence inside a longer explanatory paragraph rather than to its own
  // PowerPoint text box later in the slide XML.
  textual.sort((left, right) => right.text.length - left.text.length || left.node.sourceTextOrder - right.node.sourceTextOrder);
  for (const item of textual) {
    let cursor = 0;
    let match = -1;
    while (cursor <= sourceText.length - item.text.length) {
      const candidate = sourceText.indexOf(item.text, cursor);
      if (candidate < 0) break;
      const end = candidate + item.text.length;
      if (!occupied.some((span) => candidate < span.end && end > span.start)) {
        match = candidate;
        occupied.push({ start: candidate, end });
        break;
      }
      cursor = candidate + 1;
    }
    indices.set(item.node.id, match >= 0 ? match : item.node.sourceTextOrder);
  }
  return indices;
}

function unionFrame(nodes: StudioWebNode[], frameKey: "frame" | "sourceFrame"): StudioWebNode["frame"] {
  const frames = nodes.map((node) => node[frameKey]);
  const left = Math.min(...frames.map((value) => value.x));
  const top = Math.min(...frames.map((value) => value.y));
  const right = Math.max(...frames.map((value) => value.x + value.width));
  const bottom = Math.max(...frames.map((value) => value.y + value.height));
  return { x: left, y: top, width: Math.max(1, right - left), height: Math.max(1, bottom - top), rotation: 0 };
}

function utf8Base64(value: string): string {
  const encoded = new TextEncoder().encode(value);
  let binary = "";
  for (const byte of encoded) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function sourceLockedCropData(
  scene: StudioWebScene,
  sourceFrame: StudioWebNode["sourceFrame"],
  raster: { data: string; width: number; height: number },
): string {
  const x = sourceFrame.x / scene.sourceSlideSize.width * raster.width;
  const y = sourceFrame.y / scene.sourceSlideSize.height * raster.height;
  const width = sourceFrame.width / scene.sourceSlideSize.width * raster.width;
  const height = sourceFrame.height / scene.sourceSlideSize.height * raster.height;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${Math.max(1, width)}" height="${Math.max(1, height)}" viewBox="${x} ${y} ${Math.max(1, width)} ${Math.max(1, height)}"><image x="0" y="0" width="${raster.width}" height="${raster.height}" preserveAspectRatio="none" href="${raster.data}" xlink:href="${raster.data}"/></svg>`;
  return `data:image/svg+xml;base64,${utf8Base64(svg)}`;
}

function croppedSourceFrame(
  frame: StudioWebNode["sourceFrame"],
  crop: { left: number; top: number; right: number; bottom: number } | undefined,
): StudioWebNode["sourceFrame"] {
  if (!crop) return frame;
  return {
    x: frame.x + frame.width * crop.left,
    y: frame.y + frame.height * crop.top,
    width: Math.max(1, frame.width * (1 - crop.left - crop.right)),
    height: Math.max(1, frame.height * (1 - crop.top - crop.bottom)),
    rotation: frame.rotation,
  };
}

function containFigureFrame(
  target: StudioWebNode["frame"],
  source: StudioWebNode["sourceFrame"],
  focalPoint: { x: number; y: number } | undefined,
): StudioWebNode["frame"] {
  const scale = Math.min(target.width / source.width, target.height / source.height);
  const width = Math.max(1, source.width * scale);
  const height = Math.max(1, source.height * scale);
  const focal = focalPoint ?? { x: .5, y: .5 };
  return {
    x: target.x + (target.width - width) * focal.x,
    y: target.y + (target.height - height) * focal.y,
    width,
    height,
    rotation: target.rotation,
  };
}

function addConvertedTemplateArtwork(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  sourceSlide: StudioWebScene["slides"][number],
  templateCatalog: TemplateCatalog | undefined,
  templateLayoutRasters: StudioCompositionExportOptions["templateLayoutRasters"],
  nativeTemplateLayoutBaseId: string | undefined,
  warnings: string[],
): void {
  const layoutId = sourceSlide.recipe === "template-layout" ? sourceSlide.targetLayoutId : nativeTemplateLayoutBaseId;
  if (!layoutId) return;
  const layout = templateCatalog?.layouts.find((item) => item.id === layoutId);
  if (!layout || !templateCatalog) throw new Error(`Slide ${sourceSlide.slideNumber} references a converted ORNL template layout that is not available in the active Template Pack.`);
  slide.background = { color: hex(layout.background, sourceSlide.background) };
  if (nativeTemplateLayoutBaseId) {
    warnings.push(`Slide ${sourceSlide.slideNumber}: editable Studio content is prepared for the native ORNL ${layout.name} layout; the caller must attach its real master/layout dependency graph before native review.`);
    return;
  }
  const xScale = PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches / (templateCatalog.slideWidth / EMU_PER_INCH);
  const yScale = PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches / (templateCatalog.slideHeight / EMU_PER_INCH);
  const frame = (element: TemplatePreviewElement) => ({
    x: inches(element.x) * xScale,
    y: inches(element.y) * yScale,
    w: inches(element.width) * xScale,
    h: inches(element.height) * yScale,
  });
  const addPlaceholderPanels = () => {
    const seen = new Set<string>();
    const occupiedByContent = (bounds: { x: number; y: number; w: number; h: number }) => sourceSlide.nodes.some((node) => {
      if (!node.visible || ["footer", "slide-number", "date", "logo"].includes(node.role)) return false;
      const nodeBounds = { x: inches(node.frame.x), y: inches(node.frame.y), w: inches(node.frame.width), h: inches(node.frame.height) };
      const tolerance = .08;
      return Math.abs(nodeBounds.x - bounds.x) <= tolerance
        && Math.abs(nodeBounds.y - bounds.y) <= tolerance
        && Math.abs(nodeBounds.w - bounds.w) <= tolerance
        && Math.abs(nodeBounds.h - bounds.h) <= tolerance;
    });
    for (const element of layout.elements) {
      if (!element.placeholderType || element.placeholderType === "pic" || !element.fill || element.kind !== "shape") continue;
      const key = [element.placeholderType, element.x, element.y, element.width, element.height, element.fill, element.stroke ?? ""].join(":");
      if (seen.has(key)) continue;
      seen.add(key);
      const bounds = frame(element);
      // Placeholder fills are semantic content furniture, not background
      // artwork. Emit a panel only when a Studio node actually occupies that
      // slot; otherwise layouts such as the three-column green bar produce
      // empty colored boxes that were never part of the requested design.
      if (!occupiedByContent(bounds)) continue;
      slide.addShape(pptx.ShapeType.rect, {
        ...bounds,
        line: element.stroke && (element.strokeWidth ?? 0) > 0 ? { color: hex(element.stroke, "DBDCDB"), width: Math.max(.1, (element.strokeWidth ?? 0) / 12_700) } : { color: "FFFFFF", transparency: 100 },
        fill: { color: hex(element.fill, "FFFFFF"), transparency: Math.round((1 - (element.opacity ?? 1)) * 100) },
        objectName: `Template panel · ${element.name}`,
      });
    }
  };
  const nativeLayoutRaster = templateLayoutRasters?.[layout.id];
  if (nativeLayoutRaster) {
    slide.addImage({
      data: nativeLayoutRaster.data,
      x: 0, y: 0,
      w: PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches,
      h: PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches,
      objectName: `Template background · ${layout.name}`,
      altText: `Authoritative Microsoft PowerPoint render of the approved ORNL ${layout.name} layout artwork. Studio content remains editable above this immutable template base.`,
    });
    warnings.push(`Slide ${sourceSlide.slideNumber}: approved ${layout.name} template artwork is embedded as one PowerPoint-native rendered base so master/layout imagery and strokes remain visually exact; Studio content above it remains editable.`);
    addPlaceholderPanels();
    return;
  }
  for (const element of layout.elements.filter((item) => !item.placeholderType)) {
    const bounds = frame(element);
    if (element.kind === "image") {
      const data = element.mediaId ? templateCatalog.media[element.mediaId] : undefined;
      if (!data) {
        warnings.push(`Slide ${sourceSlide.slideNumber}: converted template image ${element.name} is unavailable.`);
        continue;
      }
      slide.addImage({ data, ...bounds, rotate: element.rotation, objectName: `Template · ${element.name}` });
      continue;
    }
    // Template text is deliberately not copied as slide-local content because
    // it would change the exact source-content inventory. Logos and recurring
    // labels should be represented by approved image/vector artwork instead.
    if (element.kind === "text") continue;
    const line = element.stroke && (element.strokeWidth ?? 0) > 0
      ? { color: hex(element.stroke, "DBDCDB"), width: Math.max(.1, (element.strokeWidth ?? 0) / 12_700), transparency: Math.round((1 - (element.opacity ?? 1)) * 100) }
      : { color: "FFFFFF", transparency: 100 };
    const fill = element.fill
      ? { color: hex(element.fill, "FFFFFF"), transparency: Math.round((1 - (element.opacity ?? 1)) * 100) }
      : { color: "FFFFFF", transparency: 100 };
    const shape = element.geometry === "ellipse" ? pptx.ShapeType.ellipse : element.geometry === "line" ? pptx.ShapeType.line : pptx.ShapeType.rect;
    slide.addShape(shape, { ...bounds, rotate: element.rotation, line, fill, objectName: `Template · ${element.name}` });
  }
  addPlaceholderPanels();
}

/**
 * Builds a genuinely new, editable PowerPoint deck from the Studio Web Scene.
 * This intentionally does not preserve the imported package's masters,
 * animations, transitions, or unsupported native internals. Use the existing
 * source-bound proposal compiler when those properties must survive.
 */
export async function buildStudioCompositionPptx(scene: StudioWebScene, options: StudioCompositionExportOptions = {}): Promise<StudioCompositionExportResult> {
  const strict = options.strict ?? true;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Presentation Studio";
  pptx.company = "Oak Ridge National Laboratory";
  pptx.subject = "Editable presentation rebuilt from a Presentation Studio semantic web scene";
  pptx.title = options.title ?? "Presentation Studio web composition";
  pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos" };
  const warnings: string[] = [];
  let textNodeCount = 0;
  let tableCount = 0;
  let imageCount = 0;
  let ignoredSourceFurnitureCount = 0;
  let generatedComponentCount = 0;
  const outputSlides: StudioCompositionOutputSlide[] = [];

  for (const sourceSceneSlide of [...scene.slides].sort((left, right) => left.slideNumber - right.slideNumber)) {
    const materializedSlides = materializeStudioTableContinuationSlides(scene, sourceSceneSlide);
    for (const materialized of materializedSlides) {
      const sourceSlide = materialized.slide;
      outputSlides.push({
        outputSlideNumber: outputSlides.length + 1,
        sourceSlideNumber: sourceSceneSlide.slideNumber,
        continuation: materialized.continuation,
      });
    if (!sourceSlide.contentCoverage.exactTextMapped) {
      const message = `Slide ${sourceSlide.slideNumber} maps ${sourceSlide.contentCoverage.mappedCharacterCount} of ${sourceSlide.contentCoverage.sourceCharacterCount} normalized source characters into editable Studio nodes. Grouped, inherited, or unsupported text must be atomized before fresh composition.`;
      if (strict) throw new Error(message);
      warnings.push(message);
    }
    const sourceLockedTreatments = sourceSlide.figureTreatments.filter((treatment) =>
      ["preserve-as-unit", "preserve-and-frame"].includes(treatment.mode)
      && ["source-locked", "verified"].includes(treatment.verificationStatus));
    const sourceLockedNodeIds = new Set(sourceLockedTreatments.flatMap((treatment) => treatment.nodeIds));
    const unsupported = sourceSlide.nodes.filter((node) => unsupportedContentNode(node) && !sourceLockedNodeIds.has(node.id));
    if (unsupported.length > 0 && strict) throw new Error(`Slide ${sourceSlide.slideNumber} contains ${unsupported.length} preserved native object${unsupported.length === 1 ? "" : "s"} that the fresh-composition compiler cannot recreate without a disclosed conversion.`);
    if (unsupported.length > 0) warnings.push(`Slide ${sourceSlide.slideNumber}: omitted ${unsupported.length} preserved native object${unsupported.length === 1 ? "" : "s"}.`);
    const slide = pptx.addSlide();
    slide.background = { color: hex(sourceSlide.background, "#FFFFFF") };
    if (sourceSlide.recipe === "source") {
      const raster = options.sourceSlideRasters?.[sourceSlide.slideNumber];
      if (!raster) {
        const message = `Slide ${sourceSlide.slideNumber}: source-preserved composition requires the authoritative PowerPoint source raster.`;
        if (strict) throw new Error(message);
        warnings.push(message);
        continue;
      }
      const sourceText = normalizedTextOrderValue(options.sourceSlideText?.[sourceSlide.slideNumber]);
      const exactNodes = sourceSlide.nodes.filter((node) => node.visible && (node.kind === "text" || node.kind === "table"));
      const sourceOrderIndices = sourceTextOrderIndices(exactNodes, sourceText);
      for (const node of exactNodes.sort((left, right) => (sourceOrderIndices.get(left.id) ?? left.sourceTextOrder) - (sourceOrderIndices.get(right.id) ?? right.sourceTextOrder))) {
        const x = inches(node.sourceFrame.x);
        const y = inches(node.sourceFrame.y);
        const w = inches(node.sourceFrame.width);
        const h = inches(node.sourceFrame.height);
        if (node.kind === "text" && node.text !== undefined) {
          slide.addText(editableText(node), { x, y, w, h, objectName: `Source inventory · ${node.name}`, fontFace: "Aptos", fontSize: 1, color: "FFFFFF", margin: 0, breakLine: false });
          textNodeCount += 1;
        } else if (node.kind === "table" && node.table) {
          slide.addTable(tableRows(node), { x, y, w, h, ...tableDimensions(node), objectName: `Source inventory · ${node.name}`, fontFace: "Aptos", fontSize: 1, color: "FFFFFF", border: { type: "solid", color: "FFFFFF", pt: .1 }, margin: 0, autoPage: false });
          tableCount += 1;
        }
      }
      slide.addImage({
        data: raster.data,
        x: 0,
        y: 0,
        w: PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches,
        h: PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches,
        objectName: "Source-preserved ORNL title slide",
        altText: `Source-preserved PowerPoint slide ${sourceSlide.slideNumber}; approved template composition remains visually unchanged.`,
      });
      imageCount += 1;
      warnings.push(`Slide ${sourceSlide.slideNumber}: preserved as one authoritative PowerPoint-rendered title composition; no template artwork or geometry was altered.`);
      continue;
    }
    addConvertedTemplateArtwork(pptx, slide, sourceSlide, options.templateCatalog, options.templateLayoutRasters, options.nativeTemplateLayoutBaseId, warnings);
    for (const component of studioGeneratedComponents(sourceSlide)) {
      const x = inches(component.frame.x);
      const y = inches(component.frame.y);
      const w = inches(component.frame.width);
      const h = Math.max(.001, inches(component.frame.height));
      slide.addShape(component.kind === "line" ? pptx.ShapeType.line : pptx.ShapeType.rect, {
        x, y, w, h,
        line: component.lineColor && component.lineWidthPt > 0 ? { color: hex(component.lineColor, "DBDCDB"), width: component.lineWidthPt } : { color: hex(component.fillColor, "FFFFFF"), transparency: 100 },
        fill: component.fillColor ? { color: hex(component.fillColor, "FFFFFF") } : { color: "FFFFFF", transparency: 100 },
        objectName: component.id,
      });
      generatedComponentCount += 1;
    }
    const sourceText = normalizedTextOrderValue(options.sourceSlideText?.[sourceSlide.slideNumber]);
    const sourceOrderIndices = sourceTextOrderIndices(sourceSlide.nodes.filter((node) => node.visible), sourceText);
    const sourceIndex = (node: StudioWebNode) => sourceOrderIndices.get(node.id) ?? node.sourceTextOrder;
    const orderedNodes = [...sourceSlide.nodes].sort((left, right) => {
      const layer = (node: StudioWebNode) => node.kind === "image" ? 0 : node.kind === "text" || node.kind === "table" ? 1 : 2;
      return layer(left) - layer(right) || (layer(left) === 1 ? sourceIndex(left) - sourceIndex(right) || left.sourceTextOrder - right.sourceTextOrder : left.zIndex - right.zIndex);
    });
    for (const node of orderedNodes) {
      if (!node.visible || unsupportedContentNode(node)) continue;
      // A source-locked evidence unit is emitted as one authoritative native
      // crop after editable text/table inventory has been written beneath it.
      // Its component images and decorative shapes must not be duplicated.
      if (sourceLockedNodeIds.has(node.id) && !["text", "table"].includes(node.kind)) continue;
      const x = inches(node.frame.x);
      const y = inches(node.frame.y);
      const w = inches(node.frame.width);
      const h = inches(node.frame.height);
      if (node.kind === "connector") {
        if (!node.connector) {
          const message = `Slide ${sourceSlide.slideNumber}: ${node.name} is a relationship-bearing connector without verified Studio endpoint bindings. Preserve its complete figure as one source-locked unit or author it inside a verified editable-diagram treatment.`;
          if (strict) throw new Error(message);
          warnings.push(message);
          continue;
        }
        const from = sourceSlide.nodes.find((candidate) => candidate.id === node.connector?.fromNodeId);
        const to = sourceSlide.nodes.find((candidate) => candidate.id === node.connector?.toNodeId);
        if (!from || !to) {
          const message = `Slide ${sourceSlide.slideNumber}: verified connector ${node.id} has a stale endpoint binding.`;
          if (strict) throw new Error(message);
          warnings.push(message);
          continue;
        }
        const start = studioConnectorAttachmentPoint(from, node.connector.fromSide);
        const end = studioConnectorAttachmentPoint(to, node.connector.toSide);
        slide.addShape(pptx.ShapeType.line, {
          x: inches(Math.min(start.x, end.x)),
          y: inches(Math.min(start.y, end.y)),
          w: Math.max(.001, inches(Math.abs(end.x - start.x))),
          h: Math.max(.001, inches(Math.abs(end.y - start.y))),
          flipH: end.x < start.x,
          flipV: end.y < start.y,
          line: { color: hex(node.connector.stroke, "007833"), width: node.connector.widthPt, dashType: node.connector.dash, beginArrowType: node.connector.beginArrow, endArrowType: node.connector.endArrow },
          objectName: `${node.name} · ${node.id}`.slice(0, 240),
        });
        generatedComponentCount += 1;
        continue;
      }
      if (node.kind === "shape") {
        if (/arrow/i.test(node.name)) {
          slide.addShape(pptx.ShapeType.rightArrow, {
            x, y, w, h,
            line: { color: hex(PRESENTATION_DESIGN_STANDARD.defaults.palette.ornlGreen, "007833"), transparency: 100 },
            fill: { color: hex(PRESENTATION_DESIGN_STANDARD.defaults.palette.ornlGreen, "007833") },
            objectName: `${node.name} · ${node.id}`.slice(0, 240),
          });
          generatedComponentCount += 1;
          continue;
        }
        ignoredSourceFurnitureCount += 1;
        continue;
      }
      if (node.kind === "text" && node.text !== undefined) {
        const sourceLockedInventory = sourceLockedNodeIds.has(node.id);
        slide.addText(editableText(node), {
          x, y, w, h,
          objectName: `${node.name} · ${node.id}`.slice(0, 240),
          fontFace: "Aptos",
          // Source-locked figure copy stays in the editable/auditable package
          // beneath the authoritative crop, but is not a second visible layer.
          fontSize: sourceLockedInventory ? 1 : node.style.fontSizePt,
          bold: node.style.fontWeight >= 600,
          color: hex(node.style.color, PRESENTATION_DESIGN_STANDARD.defaults.palette.darkMatter),
          align: node.style.textAlign,
          valign: node.style.verticalAlign === "middle" ? "middle" : node.style.verticalAlign === "bottom" ? "bottom" : "top",
          margin: margins(node),
          breakLine: false,
          lineSpacingMultiple: node.style.lineHeight,
          rotate: node.frame.rotation,
        });
        textNodeCount += 1;
        continue;
      }
      if (node.kind === "table" && node.table) {
        const sourceLockedInventory = sourceLockedNodeIds.has(node.id);
        slide.addTable(tableRows(node), {
          x, y, w, h,
          ...tableDimensions(node),
          objectName: `${node.name} · ${node.id}`.slice(0, 240),
          fontFace: "Aptos",
          fontSize: sourceLockedInventory ? 1 : node.style.fontSizePt,
          color: hex(node.style.color, PRESENTATION_DESIGN_STANDARD.defaults.palette.darkMatter),
          autoPage: false,
        });
        tableCount += 1;
        continue;
      }
      if (node.kind === "image") {
        const extractedData = node.mediaPart ? options.catalog?.media[node.mediaPart] : undefined;
        const sourceRaster = options.sourceSlideRasters?.[sourceSlide.slideNumber];
        const cropFrame = node.component?.role === "process-icon" ? croppedSourceFrame(node.sourceFrame, { left: .04, top: .04, right: .04, bottom: .04 }) : node.sourceFrame;
        const data = extractedData ?? (sourceRaster ? sourceLockedCropData(scene, cropFrame, sourceRaster) : undefined);
        if (!data) {
          const message = `Slide ${sourceSlide.slideNumber}: ${node.name} has neither an extracted image asset nor an authoritative PowerPoint source raster crop.`;
          if (strict) throw new Error(message);
          warnings.push(message);
          continue;
        }
        const protectedBrandMark = node.component?.role === "footer-logo" || /(?:^|\b)(?:ornl|doe|department of energy|oak ridge|wordmark|logo)(?:\b|$)/i.test(node.name);
        const targetFrame = node.style.objectFit === "contain" || protectedBrandMark
          ? containFigureFrame(node.frame, cropFrame, undefined)
          : node.frame;
        slide.addImage({
          data,
          x: inches(targetFrame.x), y: inches(targetFrame.y), w: inches(targetFrame.width), h: inches(targetFrame.height),
          rotate: node.frame.rotation,
          objectName: `${node.name} · ${node.id}`.slice(0, 240),
          altText: extractedData
            ? node.name
            : `PowerPoint-native crop preserving ${node.name} from source slide ${sourceSlide.slideNumber}.`,
          transparency: 0,
        });
        if (!extractedData) warnings.push(`Slide ${sourceSlide.slideNumber}: ${node.name} was preserved from the authoritative PowerPoint source render because its embedded media part was unavailable.`);
        imageCount += 1;
      }
    }
    for (const treatment of sourceLockedTreatments) {
      const nodes = treatment.nodeIds.map((id) => sourceSlide.nodes.find((node) => node.id === id)).filter((node): node is StudioWebNode => Boolean(node?.visible));
      if (!nodes.length) continue;
      const isolatedRaster = options.sourceFigureRasters?.[treatment.id];
      const raster = isolatedRaster ?? options.sourceSlideRasters?.[sourceSlide.slideNumber];
      if (!raster) {
        const message = `Slide ${sourceSlide.slideNumber}: source-locked figure ${treatment.id} requires the authoritative PowerPoint source raster.`;
        if (strict) throw new Error(message);
        warnings.push(message);
        continue;
      }
      const sourceFrame = croppedSourceFrame(unionFrame(nodes, "sourceFrame"), treatment.crop);
      const requestedTargetFrame = treatment.groupFrame ?? unionFrame(nodes, "frame");
      const targetFrame = treatment.lockAspectRatio === false ? requestedTargetFrame : containFigureFrame(requestedTargetFrame, sourceFrame, treatment.focalPoint);
      const data = sourceLockedCropData(scene, sourceFrame, raster);
      slide.addImage({
        data,
        x: inches(targetFrame.x), y: inches(targetFrame.y), w: inches(targetFrame.width), h: inches(targetFrame.height),
        objectName: `Source-locked · ${treatment.intentSummary} · ${treatment.id}`.slice(0, 240),
        altText: `Source-locked technical evidence preserved from slide ${sourceSlide.slideNumber}. ${treatment.intentSummary}`.slice(0, 500),
      });
      imageCount += 1;
      warnings.push(`Slide ${sourceSlide.slideNumber}: ${treatment.intentSummary} is preserved as one source-locked PowerPoint-rendered evidence unit${isolatedRaster ? " from an object-isolated native render" : ""}.`);
    }
    }
  }
  const output = await pptx.write({ outputType: "uint8array", compression: true });
  const bytes = output instanceof Uint8Array ? output : new Uint8Array(output as ArrayBuffer);
  return { bytes, slideCount: outputSlides.length, outputSlides, textNodeCount, tableCount, imageCount, ignoredSourceFurnitureCount, generatedComponentCount, warnings };
}
