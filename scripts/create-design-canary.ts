import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import PptxGenJS from "pptxgenjs";

const ORNL_GREEN = "007A33";
const ORNL_DARK_GREEN = "004C23";
const ORNL_GRAY = "373A36";
const LIGHT_GRAY = "E8ECE9";
const PT = 1 / 72;

function baseSlide(pptx: PptxGenJS, title: string, number: number) {
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 7.5, line: { color: ORNL_GREEN, transparency: 100 }, fill: { color: ORNL_GREEN }, objectName: `canary-${number}-brand-rule` });
  slide.addText(`CANARY ${String(number).padStart(2, "0")}`, { x: 11.75, y: 0.28, w: 0.85, h: 0.22, margin: 0, fontFace: "Aptos", fontSize: 9, bold: true, align: "right", color: ORNL_GREEN, objectName: `canary-${number}-label` });
  slide.addText(title, { x: 0.72, y: 0.48, w: 9.5, h: 0.55, margin: 0, fontFace: "Aptos Display", fontSize: 26, bold: true, color: ORNL_DARK_GREEN, objectName: `canary-${number}-reference-title` });
  slide.addText("Synthetic precision-layout fixture · not client content", { x: 0.72, y: 7.05, w: 5.8, h: 0.2, margin: 0, fontFace: "Aptos", fontSize: 9, color: "5E6B65", objectName: `canary-${number}-footer` });
  return slide;
}

function addPeer(slide: PptxGenJS.Slide, text: string, x: number, y: number, name: string, options: { margin?: number | [number, number, number, number]; bullet?: boolean; height?: number } = {}) {
  slide.addText(options.bullet ? [{ text, options: { bullet: { indent: 18 } } }] : text, { x, y, w: 5.6, h: options.height ?? 0.42, margin: options.margin ?? 0, fontFace: "Aptos", fontSize: 16, color: ORNL_GRAY, breakLine: false, objectName: name });
}

function tableRows() {
  return [
    [{ text: "Parameter" }, { text: "Baseline" }, { text: "Candidate" }],
    [{ text: "Response" }, { text: "14 ms" }, { text: "11 ms" }],
    [{ text: "Margin" }, { text: "7.2%" }, { text: "9.4%" }],
  ];
}

export async function createDesignCanaryDeck(filePath: string) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Presentation Studio synthetic canary";
  pptx.company = "Synthetic test data - not client content";
  pptx.subject = "PowerPoint-native precision layout canaries";
  pptx.title = "Presentation Studio precision-layout canary deck";
  pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos" };

  const slide1 = baseSlide(pptx, "Title is two points left of the intended grid", 1);
  slide1.addText("Measured title defect: 2 pt", { x: 0.72 - 2 * PT, y: 1.35, w: 7.5, h: 0.55, margin: 0, fontFace: "Aptos Display", fontSize: 25, bold: true, color: ORNL_DARK_GREEN, objectName: "canary-1-title-off-2pt" });
  addPeer(slide1, "Reference body begins at the 0.72-inch content grid.", 0.72, 2.25, "canary-1-body-reference");

  const slide2 = baseSlide(pptx, "Title is four points left of the intended grid", 2);
  slide2.addText("Measured title defect: 4 pt", { x: 0.72 - 4 * PT, y: 1.35, w: 7.5, h: 0.55, margin: 0, fontFace: "Aptos Display", fontSize: 25, bold: true, color: ORNL_DARK_GREEN, objectName: "canary-2-title-off-4pt" });
  addPeer(slide2, "Reference body begins at the 0.72-inch content grid.", 0.72, 2.25, "canary-2-body-reference");

  const slide3 = baseSlide(pptx, "Shared shape edges can hide optical misalignment", 3);
  addPeer(slide3, "Visible text begins at the shape edge.", 0.72, 1.55, "canary-3-text-zero-margin");
  // PptxGenJS maps the first text-margin tuple item to PowerPoint's left inset.
  addPeer(slide3, "This text begins 5.4 pt farther right.", 0.72, 2.25, "canary-3-text-margin-5_4pt", { margin: [5.4, 0, 0, 0] });
  slide3.addShape(pptx.ShapeType.line, { x: 0.72, y: 1.35, w: 0, h: 1.55, line: { color: "B8C3BD", width: 1, dashType: "dash" }, objectName: "canary-3-reference-grid" });

  const slide4 = baseSlide(pptx, "Bullet indentation creates a false shape alignment", 4);
  addPeer(slide4, "Plain paragraph visible start", 0.72, 1.55, "canary-4-plain");
  addPeer(slide4, "Bullet visible start", 0.72, 2.25, "canary-4-bullet-indent", { bullet: true });
  slide4.addShape(pptx.ShapeType.line, { x: 0.72, y: 1.35, w: 0, h: 1.55, line: { color: "B8C3BD", width: 1, dashType: "dash" }, objectName: "canary-4-reference-grid" });

  const slide5 = baseSlide(pptx, "Vertical rhythm contains one deliberate 24-point gap", 5);
  const boxHeight = 24 * PT;
  const y1 = 1.45;
  const y2 = y1 + boxHeight + 16 * PT;
  const y3 = y2 + boxHeight + 24 * PT;
  const y4 = y3 + boxHeight + 16 * PT;
  [["First evidence block", y1], ["Second evidence block", y2], ["Third evidence block", y3], ["Fourth evidence block", y4]].forEach(([text, y], index) => {
    slide5.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: Number(y), w: 6.2, h: boxHeight, rectRadius: 0, line: { color: "CCD5D0", width: 1 }, fill: { color: index === 2 ? "F3F6F4" : "FFFFFF" }, objectName: `canary-5-panel-${index + 1}` });
    slide5.addText(String(text), { x: 0.9, y: Number(y) + 0.045, w: 5.8, h: 0.22, margin: 0, fontFace: "Aptos", fontSize: 14, color: ORNL_GRAY, objectName: `canary-5-body-${index + 1}` });
  });

  const slide6 = baseSlide(pptx, "Object enters the safe region by four points", 6);
  slide6.addShape(pptx.ShapeType.rect, { x: 14 * PT, y: 1.55, w: 3.6, h: 1.1, line: { color: "D36B33", width: 1.5 }, fill: { color: "FFF5EF" }, objectName: "canary-6-safe-margin-minus-4pt" });
  slide6.addText("The left edge is 14 pt from the slide edge; the required review margin is 18 pt.", { x: 14 * PT + 0.15, y: 1.77, w: 3.25, h: 0.6, margin: 0, fontFace: "Aptos", fontSize: 15, color: ORNL_GRAY, objectName: "canary-6-safe-margin-copy" });

  const paddingSlides = [
    { number: 7, paddingPt: 2, label: "Two-point cell padding" },
    { number: 8, paddingPt: 4, label: "Four-point cell padding" },
    { number: 9, paddingPt: 6, label: "Six-point cell padding" },
  ];
  for (const item of paddingSlides) {
    const slide = baseSlide(pptx, item.label, item.number);
    slide.addTable(tableRows(), { x: 0.72, y: 1.45, w: 8.8, h: 2.25, margin: item.paddingPt * PT, fontFace: "Aptos", fontSize: 16, color: ORNL_GRAY, border: { type: "solid", color: "AAB8B1", pt: 0.75 }, fill: { color: "FFFFFF" }, rowH: 0.68, objectName: `canary-${item.number}-table-padding-${item.paddingPt}pt` });
  }

  const slide10 = baseSlide(pptx, "Cell text is clipped by approximately two points", 10);
  slide10.addTable([[{ text: "Header" }, { text: "Value" }], [{ text: "Raised type" }, { text: "18 pt" }]], { x: 0.72, y: 1.55, w: 7.1, h: 0.58, rowH: 0.29, margin: 2 * PT, fontFace: "Aptos", fontSize: 18, color: ORNL_GRAY, border: { type: "solid", color: "AAB8B1", pt: 0.75 }, fill: { color: "FFFFFF" }, objectName: "canary-10-table-clipped-2pt" });

  const slide11 = baseSlide(pptx, "A narrow column causes undesirable wrapping", 11);
  slide11.addTable([
    [{ text: "Case" }, { text: "Technical description" }, { text: "Result" }],
    [{ text: "A" }, { text: "Coordinated electromagnetic transient response" }, { text: "Pass" }],
    [{ text: "B" }, { text: "System-level validation boundary condition" }, { text: "Review" }],
  ], { x: 0.72, y: 1.45, w: 8.9, h: 2.4, colW: [1.15, 1.55, 1.25], rowH: 0.68, margin: 5 * PT, fontFace: "Aptos", fontSize: 14, color: ORNL_GRAY, border: { type: "solid", color: "AAB8B1", pt: 0.75 }, fill: { color: "FFFFFF" }, objectName: "canary-11-table-narrow-column" });

  const slide12 = baseSlide(pptx, "Row height is too short after a font-size increase", 12);
  slide12.addTable([[{ text: "Parameter" }, { text: "Raised value" }], [{ text: "Response" }, { text: "Twenty-four point technical value" }]], { x: 0.72, y: 1.55, w: 8.3, h: 0.72, rowH: 0.36, margin: 4 * PT, fontFace: "Aptos", fontSize: 24, color: ORNL_GRAY, border: { type: "solid", color: "AAB8B1", pt: 0.75 }, fill: { color: "FFFFFF" }, objectName: "canary-12-table-short-row" });

  const slide13 = baseSlide(pptx, "Merged cells must remain structurally intact", 13);
  slide13.addTable([
    [{ text: "Merged technical heading", options: { colspan: 3 } }],
    [{ text: "Case" }, { text: "Baseline" }, { text: "Candidate" }],
    [{ text: "Synthetic A" }, { text: "14" }, { text: "18" }],
  ], { x: 0.72, y: 1.45, w: 8.8, h: 2.2, rowH: 0.62, margin: 5 * PT, fontFace: "Aptos", fontSize: 14, color: ORNL_GRAY, border: { type: "solid", color: "AAB8B1", pt: 0.75 }, fill: { color: "FFFFFF" }, objectName: "canary-13-table-merged-cell" });

  const slide14 = baseSlide(pptx, "Mechanical alignment can still feel visually unbalanced", 14);
  slide14.addShape(pptx.ShapeType.rect, { x: 0.72, y: 1.5, w: 2.2, h: 3.9, line: { color: ORNL_GREEN, width: 1 }, fill: { color: "EAF4EE" }, objectName: "canary-14-small-left-panel" });
  slide14.addShape(pptx.ShapeType.rect, { x: 3.15, y: 1.5, w: 8.9, h: 3.9, line: { color: "B8C3BD", width: 1 }, fill: { color: "F7F8F7" }, objectName: "canary-14-large-right-panel" });
  slide14.addText("Aligned", { x: 0.98, y: 1.85, w: 1.65, h: 0.45, margin: 0, fontFace: "Aptos", fontSize: 22, bold: true, color: ORNL_DARK_GREEN, objectName: "canary-14-left-label" });
  slide14.addText("All panel edges and labels sit on a consistent grid, but the composition is deliberately left-light and right-heavy. This slide requires AI design judgment rather than a purely geometric pass.", { x: 3.55, y: 1.85, w: 7.9, h: 1.35, margin: 0, fontFace: "Aptos", fontSize: 18, color: ORNL_GRAY, objectName: "canary-14-right-copy" });

  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await pptx.writeFile({ fileName: filePath });
  return filePath;
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (invokedPath === fileURLToPath(import.meta.url)) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const target = path.join(root, "fixtures", "generated", "precision-layout-canary.pptx");
  await createDesignCanaryDeck(target);
  console.log(`Created ${target}`);
}
