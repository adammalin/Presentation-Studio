import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildSlideRenderCatalog } from "../src/lib/template-catalog";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { compileStudioWebScene, recomposeStudioWebSlide } from "../src/lib/studio-web-scene";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { bindNativeMeasurement } from "../src/lib/native-measurement";
import { calculateDesignMetrics } from "../src/lib/design-metrics";
import { sha256 } from "../src/lib/hash";
import type { NativeMeasurementResult, NativeRenderResult } from "../src/lib/desktop";
import type { DeckJob, StudioWebScene } from "../src/types";

const require = createRequire(import.meta.url);
const { measurePowerPointNative } = require("../electron/native-measurement.cjs") as { measurePowerPointNative(input: { bytes: Uint8Array; name: string }): Promise<NativeMeasurementResult> };
const { renderPowerPointNative } = require("../electron/native-render.cjs") as { renderPowerPointNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }): Promise<NativeRenderResult> };

function argument(name: string) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function requestedSlides(value: string | undefined) {
  const slides = (value ?? "2,6,21").split(",").map((item) => Number(item.trim())).filter((item) => Number.isInteger(item) && item > 0);
  if (slides.length === 0 || new Set(slides).size !== slides.length) throw new Error("--slides must contain unique positive slide numbers separated by commas.");
  return slides;
}

async function writeRender(render: NativeRenderResult, outputRoot: string, prefix: string, sourceNumbers?: number[], renumber = false) {
  if (render.status !== "ready" || render.renderer !== "powerpoint-native") return [];
  const outputs: string[] = [];
  for (const slide of render.slides) {
    const sourceNumber = renumber ? sourceNumbers?.[slide.number - 1] ?? slide.number : slide.number;
    if (sourceNumbers && !sourceNumbers.includes(sourceNumber)) continue;
    const output = path.join(outputRoot, `${prefix}-slide-${String(sourceNumber).padStart(2, "0")}.png`);
    await fs.writeFile(output, Buffer.from(slide.bytes));
    outputs.push(output);
  }
  return outputs;
}

function exactSelectedText(source: Awaited<ReturnType<typeof auditPptx>>, candidate: Awaited<ReturnType<typeof auditPptx>>, slideNumbers: number[]) {
  return candidate.slideCount === slideNumbers.length && slideNumbers.every((slideNumber, index) => source.slides.find((slide) => slide.number === slideNumber)?.textHash === candidate.slides[index]?.textHash);
}

function exactSelectedTableGrid(source: Awaited<ReturnType<typeof auditPptx>>, candidate: Awaited<ReturnType<typeof auditPptx>>, slideNumbers: number[]) {
  return slideNumbers.every((slideNumber, candidateIndex) => {
    const sourceTables = source.tables.filter((table) => table.slideNumber === slideNumber);
    const candidateTables = candidate.tables.filter((table) => table.slideNumber === candidateIndex + 1);
    if (sourceTables.length !== candidateTables.length) return false;
    return sourceTables.every((table, tableIndex) => {
      const rebuilt = candidateTables[tableIndex];
      return Boolean(rebuilt)
        && table.rowCount === rebuilt.rowCount
        && table.columnCount === rebuilt.columnCount
        && table.contentHash === rebuilt.contentHash
        && table.structureHash === rebuilt.structureHash;
    });
  });
}

function candidateDeck(audit: Awaited<ReturnType<typeof auditPptx>>, bytes: Uint8Array, name: string): Promise<DeckJob> {
  return sha256(bytes).then((sourceSha256) => {
    const deck: DeckJob = { id: "studio-web-benchmark-candidate", name, sourceResourceId: "studio-web-benchmark-candidate-resource", sourceSha256, operationScope: "reflow", templateClassification: audit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit, protectedSlideNumbers: [] };
    deck.scene = compilePresentationScene({ ...deck, audit });
    return deck;
  });
}

export async function qualifyStudioWebBenchmark(sourcePath: string, slideNumbers: number[], outputRoot: string, benchmarkPath?: string) {
  await fs.mkdir(outputRoot, { recursive: true });
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const sourceAudit = await auditPptx(sourceBytes);
  for (const slideNumber of slideNumbers) if (!sourceAudit.slides.some((slide) => slide.number === slideNumber)) throw new Error(`Source slide ${slideNumber} does not exist.`);
  const sourceSha256 = await sha256(sourceBytes);
  const sourceDeck: DeckJob = { id: "studio-web-benchmark-source", name: path.basename(sourcePath), sourceResourceId: "studio-web-benchmark-source-resource", sourceSha256, operationScope: "reflow", templateClassification: sourceAudit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit: sourceAudit, protectedSlideNumbers: [] };
  sourceDeck.scene = compilePresentationScene({ ...sourceDeck, audit: sourceAudit });
  const catalog = await buildSlideRenderCatalog(sourceBytes, path.basename(sourcePath));
  let scene = compileStudioWebScene(sourceDeck, catalog);
  const mappingCompleteBeforeDesign = slideNumbers.every((slideNumber) => scene.slides.find((slide) => slide.slideNumber === slideNumber)?.contentCoverage.exactTextMapped);
  for (const slideNumber of slideNumbers) scene = recomposeStudioWebSlide(scene, slideNumber);
  scene = { ...scene, slides: slideNumbers.map((slideNumber) => scene.slides.find((slide) => slide.slideNumber === slideNumber)!).filter(Boolean) } as StudioWebScene;
  const composition = await buildStudioCompositionPptx(scene, { catalog, strict: true, title: "Presentation Studio native benchmark" });
  const candidatePath = path.join(outputRoot, "studio-web-benchmark.pptx");
  await fs.writeFile(candidatePath, composition.bytes);
  const candidateAudit = await auditPptx(composition.bytes);
  const rebuiltDeck = await candidateDeck(candidateAudit, composition.bytes, path.basename(candidatePath));
  const nativeMeasurement = await measurePowerPointNative({ bytes: composition.bytes, name: path.basename(candidatePath) });
  const measurement = bindNativeMeasurement(rebuiltDeck, nativeMeasurement);
  const metrics = calculateDesignMetrics(rebuiltDeck, measurement);
  const candidateRender = await renderPowerPointNative({ bytes: composition.bytes, name: path.basename(candidatePath), width: 2200, format: "png" });
  const sourceRender = await renderPowerPointNative({ bytes: sourceBytes, name: path.basename(sourcePath), width: 2200, format: "png" });
  const candidateImages = await writeRender(candidateRender, outputRoot, "candidate", slideNumbers, true);
  const sourceImages = await writeRender(sourceRender, outputRoot, "source", slideNumbers);
  let benchmarkImages: string[] = [];
  let benchmarkRenderStatus: NativeRenderResult["status"] | undefined;
  if (benchmarkPath) {
    const benchmarkBytes = new Uint8Array(await fs.readFile(benchmarkPath));
    const benchmarkRender = await renderPowerPointNative({ bytes: benchmarkBytes, name: path.basename(benchmarkPath), width: 2200, format: "png" });
    benchmarkRenderStatus = benchmarkRender.status;
    benchmarkImages = await writeRender(benchmarkRender, outputRoot, "visual-benchmark");
  }
  const fonts = [...new Set(candidateAudit.fonts.map((font) => font.family))];
  const checks = {
    mappingCompleteBeforeDesign,
    exactVisibleTextSequence: exactSelectedText(sourceAudit, candidateAudit, slideNumbers),
    exactNativeTableGrid: exactSelectedTableGrid(sourceAudit, candidateAudit, slideNumbers),
    allEditableTextUsesAptosOrSymbolFont: candidateAudit.textBoxes.every((textBox) => textBox.fontFamilies.every((family) => family === "Aptos" || ["Wingdings", "Symbol"].includes(family))) && candidateAudit.tables.every((table) => table.cellFonts.every((family) => family === "Aptos")),
    nativePowerPointRenderReady: candidateRender.status === "ready" && candidateRender.renderer === "powerpoint-native" && candidateRender.authoritative,
    nativePowerPointMeasurementReady: nativeMeasurement.status === "ready" && nativeMeasurement.authority === "powerpoint-native",
    noNativeTextOverflow: metrics.totals.textOverflowCount === 0,
    noOffSlideObjects: metrics.totals.offSlideObjectCount === 0,
  };
  const report = {
    schema: "presentation-studio/studio-web-benchmark",
    version: 1,
    generatedAt: new Date().toISOString(),
    source: { path: sourcePath, sha256: sourceSha256, slideNumbers },
    benchmark: benchmarkPath ? { path: benchmarkPath, renderStatus: benchmarkRenderStatus } : undefined,
    candidate: { path: candidatePath, sha256: await sha256(composition.bytes), recipes: scene.slides.map((slide) => ({ sourceSlideNumber: slide.slideNumber, recipe: slide.recipe, semanticNodeCount: slide.nodes.filter((node) => node.visible).length })), fonts, ...composition },
    metrics,
    renders: { source: sourceImages, candidate: candidateImages, visualBenchmark: benchmarkImages },
    checks,
    readyForHumanVisualReview: Object.values(checks).every(Boolean),
  };
  const reportPath = path.join(outputRoot, "report.json");
  await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  return { report, reportPath };
}

if (path.resolve(fileURLToPath(import.meta.url)) === path.resolve(process.argv[1] ?? "")) {
  const sourcePath = argument("source");
  if (!sourcePath) throw new Error("Use --source /absolute/path/to/source.pptx.");
  const outputRoot = argument("output") ?? path.join(os.tmpdir(), "presentation-studio-web-benchmark");
  qualifyStudioWebBenchmark(sourcePath, requestedSlides(argument("slides")), outputRoot, argument("benchmark")).then(({ report, reportPath }) => {
    process.stdout.write(`${JSON.stringify({ readyForHumanVisualReview: report.readyForHumanVisualReview, checks: report.checks, candidate: report.candidate.path, report: reportPath, renders: report.renders }, null, 2)}\n`);
    if (!report.readyForHumanVisualReview) process.exitCode = 1;
  }).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
