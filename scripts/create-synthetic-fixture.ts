import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import PptxGenJS from "pptxgenjs";

export async function createSyntheticLegacyDeck(filePath: string): Promise<string> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Presentation Studio synthetic fixture";
  pptx.subject = "Synthetic font and table cleanup test";
  pptx.title = "Synthetic legacy formatting fixture";
  pptx.company = "Synthetic test data - not client content";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
  };

  const first = pptx.addSlide();
  first.background = { color: "FFFFFF" };
  first.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.14, h: 7.5, line: { color: "00662C", transparency: 100 }, fill: { color: "00662C" } });
  first.addText("Legacy typography should be auditable", { x: 0.65, y: 0.58, w: 7.4, h: 0.55, fontFace: "Century Gothic", fontSize: 24, bold: true, color: "00454D", margin: 0 });
  first.addText("Every sentence in this file is fictional synthetic fixture copy.", { x: 0.68, y: 1.55, w: 6.6, h: 0.5, fontFace: "Arial", fontSize: 15, color: "373A36", margin: 0 });
  first.addText("Visible text must remain byte-for-byte equivalent after font cleanup.", { x: 0.68, y: 2.18, w: 6.6, h: 0.5, fontFace: "Aptos", fontSize: 15, color: "373A36", margin: 0 });

  const second = pptx.addSlide();
  second.background = { color: "FFFFFF" };
  second.addText("Synthetic table style sample", { x: 0.65, y: 0.5, w: 7.6, h: 0.55, fontFace: "Century Gothic", fontSize: 23, bold: true, color: "00454D", margin: 0 });
  second.addTable([
    [{ text: "Category" }, { text: "Baseline" }, { text: "Candidate" }],
    [{ text: "Alpha" }, { text: "14" }, { text: "18" }],
    [{ text: "Beta" }, { text: "22" }, { text: "27" }],
  ], {
    x: 0.68, y: 1.55, w: 8.6, h: 2.2,
    fontFace: "Arial", fontSize: 13, color: "373A36",
    border: { type: "solid", color: "AAB8B1", pt: 1 },
    fill: { color: "FFFFFF" }, margin: 0.08,
    rowH: 0.62,
    bold: false,
  });
  second.addText("Numbers are synthetic and have no scientific meaning.", { x: 0.68, y: 4.2, w: 7.2, h: 0.4, fontFace: "Aptos", fontSize: 12, color: "5E6B65", margin: 0 });

  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await pptx.writeFile({ fileName: filePath });
  return filePath;
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (invokedPath === fileURLToPath(import.meta.url)) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const target = path.join(root, "fixtures", "generated", "synthetic-legacy-fonts.pptx");
  await createSyntheticLegacyDeck(target);
  console.log(`Created ${target}`);
}
