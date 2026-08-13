import PptxGenJS from "pptxgenjs";
import type { SlideRenderCatalog } from "./template-catalog";
import type { StudioWebNode, StudioWebScene } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { studioGeneratedComponents } from "./studio-web-scene";

const EMU_PER_INCH = 914_400;

export interface StudioCompositionExportResult {
  bytes: Uint8Array;
  slideCount: number;
  textNodeCount: number;
  tableCount: number;
  imageCount: number;
  ignoredSourceFurnitureCount: number;
  generatedComponentCount: number;
  warnings: string[];
}

export interface StudioCompositionExportOptions {
  catalog?: SlideRenderCatalog;
  strict?: boolean;
  title?: string;
}

function inches(value: number): number {
  return value / EMU_PER_INCH;
}

function hex(value: string | undefined, fallback: string): string {
  const candidate = /^#[0-9a-f]{6}$/i.test(value ?? "") ? value! : fallback;
  return candidate.slice(1).toUpperCase();
}

function margins(node: StudioWebNode): [number, number, number, number] {
  return [node.style.paddingPt.top, node.style.paddingPt.right, node.style.paddingPt.bottom, node.style.paddingPt.left];
}

function tableRows(node: StudioWebNode): PptxGenJS.TableRow[] {
  if (!node.table) return [];
  const rows: PptxGenJS.TableRow[] = Array.from({ length: node.table.rows }, () => Array.from({ length: node.table!.columns }, () => ({ text: "" })));
  for (const cell of node.table.cells) {
    const row = Math.max(0, cell.row - 1);
    const column = Math.max(0, cell.column - 1);
    if (!rows[row]?.[column]) continue;
    const header = cell.row === 1;
    rows[row][column] = {
      text: cell.text,
      options: {
        rowspan: Math.max(1, cell.rowSpan),
        colspan: Math.max(1, cell.columnSpan),
        bold: header,
        color: header ? "FFFFFF" : hex(node.style.color, PRESENTATION_DESIGN_STANDARD.defaults.palette.darkMatter),
        fill: { color: header ? "00454D" : hex(cell.fill, cell.row % 2 === 0 ? "#F0F2F1" : "#FFFFFF") },
        valign: "middle",
        margin: [4, 6, 4, 6],
      },
    };
  }
  return rows;
}

function unsupportedContentNode(node: StudioWebNode): boolean {
  return node.visible && ["native-object", "connector"].includes(node.kind);
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

  for (const sourceSlide of [...scene.slides].sort((left, right) => left.slideNumber - right.slideNumber)) {
    const unsupported = sourceSlide.nodes.filter(unsupportedContentNode);
    if (unsupported.length > 0 && strict) throw new Error(`Slide ${sourceSlide.slideNumber} contains ${unsupported.length} preserved native object${unsupported.length === 1 ? "" : "s"} that the fresh-composition compiler cannot recreate without a disclosed conversion.`);
    if (unsupported.length > 0) warnings.push(`Slide ${sourceSlide.slideNumber}: omitted ${unsupported.length} preserved native object${unsupported.length === 1 ? "" : "s"}.`);
    const slide = pptx.addSlide();
    slide.background = { color: hex(sourceSlide.background, "#FFFFFF") };
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
    for (const node of [...sourceSlide.nodes].sort((left, right) => left.zIndex - right.zIndex)) {
      if (!node.visible || unsupportedContentNode(node)) continue;
      const x = inches(node.frame.x);
      const y = inches(node.frame.y);
      const w = inches(node.frame.width);
      const h = inches(node.frame.height);
      if (node.kind === "shape") {
        ignoredSourceFurnitureCount += 1;
        continue;
      }
      if (node.kind === "text" && node.text !== undefined) {
        slide.addText(node.text, {
          x, y, w, h,
          objectName: node.name,
          fontFace: "Aptos",
          fontSize: node.style.fontSizePt,
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
        slide.addTable(tableRows(node), {
          x, y, w, h,
          objectName: node.name,
          fontFace: "Aptos",
          fontSize: node.style.fontSizePt,
          color: hex(node.style.color, PRESENTATION_DESIGN_STANDARD.defaults.palette.darkMatter),
          border: { type: "solid", color: "DBDCDB", pt: .75 },
          margin: [4, 6, 4, 6],
          autoPage: false,
          valign: "middle",
        });
        tableCount += 1;
        continue;
      }
      if (node.kind === "image") {
        const data = node.mediaPart ? options.catalog?.media[node.mediaPart] : undefined;
        if (!data) {
          const message = `Slide ${sourceSlide.slideNumber}: ${node.name} has no extracted image asset in the current web-scene catalog.`;
          if (strict) throw new Error(message);
          warnings.push(message);
          continue;
        }
        slide.addImage({ data, x, y, w, h, rotate: node.frame.rotation, objectName: node.name, transparency: 0 });
        imageCount += 1;
      }
    }
  }
  const output = await pptx.write({ outputType: "uint8array", compression: true });
  const bytes = output instanceof Uint8Array ? output : new Uint8Array(output as ArrayBuffer);
  return { bytes, slideCount: scene.slides.length, textNodeCount, tableCount, imageCount, ignoredSourceFurnitureCount, generatedComponentCount, warnings };
}
