import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { auditPptx } from "../src/lib/pptx-audit";
import { buildSlideRenderCatalog, buildTemplateCatalog } from "../src/lib/template-catalog";
import { buildTemplatePreviewDeck } from "../src/lib/template-preview-deck";
import { compilePresentationScene } from "../src/lib/scene-graph";
import { compileStudioWebScene, recomposeStudioWebSlide } from "../src/lib/studio-web-scene";
import { buildStudioCompositionPptx } from "../src/lib/studio-composition-export";
import { analyzeStudioDesignImpact } from "../src/lib/studio-design-impact";
import { contentProfileForSlide } from "../src/lib/design-work-order";
import { rankLayoutCompatibility } from "../src/lib/layout-semantics";
import { bindNativeMeasurement } from "../src/lib/native-measurement";
import { calculateDesignMetrics } from "../src/lib/design-metrics";
import { sha256 } from "../src/lib/hash";
import { isolateNativePowerPointObjects } from "../src/lib/native-object-isolation";
import type { NativeMeasurementResult, NativeRenderResult } from "../src/lib/desktop";
import type { DeckJob, StudioWebScene } from "../src/types";

const require = createRequire(import.meta.url);
const { measurePowerPointNative } = require("../electron/native-measurement.cjs") as { measurePowerPointNative(input: { bytes: Uint8Array; name: string }): Promise<NativeMeasurementResult> };
const { renderPowerPointNative } = require("../electron/native-render.cjs") as { renderPowerPointNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }): Promise<NativeRenderResult> };

async function retryNative<T extends { status: string }>(operation: () => Promise<T>, accepted: (result: T) => boolean): Promise<T> {
  let result = await operation();
  for (let attempt = 2; attempt <= 3 && !accepted(result); attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 350 * attempt));
    result = await operation();
  }
  return result;
}

function renderNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }) {
  return retryNative(() => renderPowerPointNative(input), (result) => result.status === "ready" && result.renderer === "powerpoint-native" && result.authoritative);
}

function measureNative(input: { bytes: Uint8Array; name: string }) {
  return retryNative(() => measurePowerPointNative(input), (result) => result.status === "ready" && result.authority === "powerpoint-native");
}

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

function nativeTextOverflowEvidence(native: NativeMeasurementResult, audit: Awaited<ReturnType<typeof auditPptx>>) {
  if (native.status !== "ready" || native.authority !== "powerpoint-native") return [];
  return native.slides.flatMap((slide) => slide.shapes.flatMap((shape) => {
    if (!shape.text?.lineCount) return [];
    const bounds = shape.boundsPt;
    const text = shape.text?.renderedBoundsPt;
    const margins = shape.text?.marginsPt;
    if (!bounds || !text || !margins || shape.text?.coordinateSpace !== "slide") return [];
    const editableObject = audit.editableObjects.find((object) => object.slideNumber === slide.number && object.name === shape.name);
    const textBox = editableObject ? audit.textBoxes.find((item) => item.slideNumber === slide.number && item.shapeId === editableObject.shapeId) : undefined;
    const tolerance = (textBox?.bulletParagraphCount ?? 0) > 0 ? 3 : .5;
    const inner = { left: bounds.left + margins.left, top: bounds.top + margins.top, right: bounds.left + bounds.width - margins.right, bottom: bounds.top + bounds.height - margins.bottom };
    const edges = [
      text.left < inner.left - tolerance ? "left" : undefined,
      text.top < inner.top - tolerance ? "top" : undefined,
      text.left + text.width > inner.right + tolerance ? "right" : undefined,
      text.top + text.height > inner.bottom + tolerance ? "bottom" : undefined,
    ].filter((edge): edge is string => Boolean(edge));
    return edges.length ? [{ slideNumber: slide.number, shapeIndex: shape.shapeIndex, name: shape.name, lineCount: shape.text?.lineCount, edges }] : [];
  }));
}

export async function qualifyStudioWebBenchmark(
  sourcePath: string,
  slideNumbers: number[],
  outputRoot: string,
  benchmarkPath?: string,
  expected?: { sourceSha256?: string; benchmarkSha256?: string },
  options?: { templatePath?: string; designMode?: "shared" | "template" },
) {
  await fs.mkdir(outputRoot, { recursive: true });
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const sourceAudit = await auditPptx(sourceBytes);
  for (const slideNumber of slideNumbers) if (!sourceAudit.slides.some((slide) => slide.number === slideNumber)) throw new Error(`Source slide ${slideNumber} does not exist.`);
  const sourceSha256 = await sha256(sourceBytes);
  if (expected?.sourceSha256 && expected.sourceSha256 !== sourceSha256) throw new Error(`The private source hash changed. Expected ${expected.sourceSha256}, received ${sourceSha256}.`);
  const sourceDeck: DeckJob = { id: "studio-web-benchmark-source", name: path.basename(sourcePath), sourceResourceId: "studio-web-benchmark-source-resource", sourceSha256, operationScope: "reflow", templateClassification: sourceAudit.classification, targetTemplateId: "ornl-16x9-v1", targetTemplateDecisionSource: "automatic-default", targetTemplateConfirmedAt: new Date().toISOString(), status: "ready-for-cleanup", audit: sourceAudit, protectedSlideNumbers: [] };
  sourceDeck.scene = compilePresentationScene({ ...sourceDeck, audit: sourceAudit });
  const catalog = await buildSlideRenderCatalog(sourceBytes, path.basename(sourcePath));
  const sourceRender = await renderNative({ bytes: sourceBytes, name: path.basename(sourcePath), width: 2200, format: "png" });
  if (sourceRender.status !== "ready" || sourceRender.renderer !== "powerpoint-native" || !sourceRender.authoritative) {
    throw new Error(`The private benchmark requires an authoritative Microsoft PowerPoint source render; received ${sourceRender.status} from ${sourceRender.renderer}.`);
  }
  const sourceSlideRasters = Object.fromEntries(sourceRender.slides.map((slide) => [slide.number, {
    data: `data:${slide.mimeType};base64,${Buffer.from(slide.bytes).toString("base64")}`,
    width: slide.width,
    height: slide.height,
  }]));
  const designMode = options?.designMode ?? "shared";
  const templateBytes = options?.templatePath ? new Uint8Array(await fs.readFile(options.templatePath)) : undefined;
  if (designMode === "template" && !templateBytes) throw new Error("Template benchmark mode requires --template /absolute/path/to/authorized-template.potx.");
  const templateCatalog = templateBytes ? await buildTemplateCatalog(templateBytes, path.basename(options!.templatePath!)) : undefined;
  let templateLayoutRasters: Record<string, { data: string; width: number; height: number }> | undefined;
  if (templateBytes && templateCatalog) {
    const preview = await buildTemplatePreviewDeck(templateBytes);
    if (preview.layoutParts.length !== templateCatalog.layouts.length || preview.layoutParts.some((part, index) => part !== templateCatalog.layouts[index]?.sourcePart)) throw new Error("The template benchmark layout order does not match the authorized Template Pack catalog.");
    const templateRender = await renderNative({ bytes: preview.bytes, name: `${path.parse(options!.templatePath!).name}-layout-previews.pptx`, width: 2200, format: "png" });
    if (templateRender.status !== "ready" || templateRender.renderer !== "powerpoint-native" || !templateRender.authoritative || templateRender.slides.length !== templateCatalog.layouts.length) throw new Error("The template benchmark requires an authoritative native render for every authorized Template Pack layout.");
    templateLayoutRasters = Object.fromEntries(templateCatalog.layouts.map((layout, index) => {
      const raster = templateRender.slides.find((slide) => slide.number === index + 1)!;
      return [layout.id, { data: `data:${raster.mimeType};base64,${Buffer.from(raster.bytes).toString("base64")}`, width: raster.width, height: raster.height }];
    }));
  }
  let scene = compileStudioWebScene(sourceDeck, catalog);
  const mappingCompleteBeforeDesign = slideNumbers.every((slideNumber) => scene.slides.find((slide) => slide.slideNumber === slideNumber)?.contentCoverage.exactTextMapped);
  for (const slideNumber of slideNumbers) {
    if (designMode === "template" && templateCatalog) {
      const recommendation = rankLayoutCompatibility(templateCatalog.layouts, contentProfileForSlide(sourceDeck, slideNumber))[0];
      const layout = templateCatalog.layouts.find((item) => item.id === recommendation?.layoutId);
      if (!layout || recommendation?.status === "incompatible") throw new Error(`No compatible authorized Template Pack layout is available for source slide ${slideNumber}.`);
      scene = recomposeStudioWebSlide(scene, slideNumber, "template-layout", layout, `Benchmark the exact source content in the recommended authorized ${layout.name} Template Pack layout.`);
    } else scene = recomposeStudioWebSlide(scene, slideNumber);
  }
  scene = { ...scene, slides: slideNumbers.map((slideNumber) => scene.slides.find((slide) => slide.slideNumber === slideNumber)!).filter(Boolean) } as StudioWebScene;
  const sourceFigureRasters: Record<string, { data: string; width: number; height: number }> = {};
  for (const slide of scene.slides) {
    for (const treatment of slide.figureTreatments.filter((item) => ["source-locked", "verified"].includes(item.verificationStatus))) {
      const shapeIds = treatment.nodeIds.map((id) => slide.nodes.find((node) => node.id === id)).filter((node) => node?.sourceBinding === "editable-object").map((node) => node!.sourceShapeId);
      if (!shapeIds.length) continue;
      const isolated = await isolateNativePowerPointObjects({ sourceBytes, slideNumber: slide.slideNumber, shapeIds });
      const rendered = await renderNative({ bytes: isolated.bytes, name: `source-slide-${slide.slideNumber}-${treatment.id}.pptx`, width: 2200, format: "png" });
      const raster = rendered.status === "ready" ? rendered.slides[0] : undefined;
      if (!raster || rendered.slides.length !== 1) throw new Error(`The private benchmark could not create an object-isolated PowerPoint render for ${treatment.id}.`);
      sourceFigureRasters[treatment.id] = { data: `data:${raster.mimeType};base64,${Buffer.from(raster.bytes).toString("base64")}`, width: raster.width, height: raster.height };
    }
  }
  const composition = await buildStudioCompositionPptx(scene, {
    catalog,
    templateCatalog,
    sourceSlideRasters,
    sourceFigureRasters,
    sourceSlideText: Object.fromEntries(sourceAudit.slides.map((slide) => [slide.number, slide.text])),
    templateLayoutRasters,
    strict: true,
    title: "Presentation Studio native benchmark",
  });
  const candidatePath = path.join(outputRoot, "studio-web-benchmark.pptx");
  await fs.writeFile(candidatePath, composition.bytes);
  const candidateAudit = await auditPptx(composition.bytes);
  const rebuiltDeck = await candidateDeck(candidateAudit, composition.bytes, path.basename(candidatePath));
  const nativeMeasurement = await measureNative({ bytes: composition.bytes, name: path.basename(candidatePath) });
  const measurement = bindNativeMeasurement(rebuiltDeck, nativeMeasurement);
  const metrics = calculateDesignMetrics(rebuiltDeck, measurement);
  const candidateRender = await renderNative({ bytes: composition.bytes, name: path.basename(candidatePath), width: 2200, format: "png" });
  const candidateImages = await writeRender(candidateRender, outputRoot, "candidate", slideNumbers, true);
  const sourceImages = await writeRender(sourceRender, outputRoot, "source", slideNumbers);
  let benchmarkImages: string[] = [];
  let benchmarkRenderStatus: NativeRenderResult["status"] | undefined;
  let benchmarkSha256: string | undefined;
  if (benchmarkPath) {
    const benchmarkBytes = new Uint8Array(await fs.readFile(benchmarkPath));
    benchmarkSha256 = await sha256(benchmarkBytes);
    if (expected?.benchmarkSha256 && expected.benchmarkSha256 !== benchmarkSha256) throw new Error(`The private visual-benchmark hash changed. Expected ${expected.benchmarkSha256}, received ${benchmarkSha256}.`);
    const benchmarkRender = await renderNative({ bytes: benchmarkBytes, name: path.basename(benchmarkPath), width: 2200, format: "png" });
    benchmarkRenderStatus = benchmarkRender.status;
    benchmarkImages = await writeRender(benchmarkRender, outputRoot, "visual-benchmark");
  }
  const fonts = [...new Set(candidateAudit.fonts.map((font) => font.family))];
  const designImpact = scene.slides.map((slide) => ({ sourceSlideNumber: slide.slideNumber, ...analyzeStudioDesignImpact(slide) }));
  const checks = {
    mappingCompleteBeforeDesign,
    exactVisibleTextSequence: exactSelectedText(sourceAudit, candidateAudit, slideNumbers),
    exactNativeTableGrid: exactSelectedTableGrid(sourceAudit, candidateAudit, slideNumbers),
    allEditableTextUsesAptosOrSymbolFont: candidateAudit.textBoxes.every((textBox) => textBox.fontFamilies.every((family) => family === "Aptos" || ["Wingdings", "Symbol"].includes(family))) && candidateAudit.tables.every((table) => table.cellFonts.every((family) => family === "Aptos")),
    nativePowerPointRenderReady: candidateRender.status === "ready" && candidateRender.renderer === "powerpoint-native" && candidateRender.authoritative,
    nativePowerPointMeasurementReady: nativeMeasurement.status === "ready" && nativeMeasurement.authority === "powerpoint-native",
    noNativeTextOverflow: metrics.totals.textOverflowCount === 0,
    noNativeTableCellClearanceViolations: metrics.totals.tableCellClearanceViolationCount === 0,
    noOffSlideObjects: metrics.totals.offSlideObjectCount === 0,
    materialCompositionBeyondTypography: designImpact.every((impact) => ["layout-redesign", "figure-redesign", "full-redesign"].includes(impact.level)),
    ...(designMode === "template" ? { authorizedTemplateArtworkApplied: scene.slides.every((slide) => slide.recipe === "template-layout" && Boolean(slide.targetLayoutId) && composition.warnings.some((warning) => warning.startsWith(`Slide ${slide.slideNumber}: approved `))) } : {}),
  };
  const report = {
    schema: "presentation-studio/studio-web-benchmark",
    version: 1,
    generatedAt: new Date().toISOString(),
    source: { path: sourcePath, sha256: sourceSha256, slideNumbers },
    template: templateCatalog ? { path: options?.templatePath, name: templateCatalog.name, sha256: templateCatalog.sha256, designMode, layoutCount: templateCatalog.layouts.length } : undefined,
    benchmark: benchmarkPath ? { path: benchmarkPath, sha256: benchmarkSha256, renderStatus: benchmarkRenderStatus } : undefined,
    candidate: { path: candidatePath, sha256: await sha256(composition.bytes), recipes: scene.slides.map((slide) => ({ sourceSlideNumber: slide.slideNumber, recipe: slide.recipe, targetLayoutId: slide.targetLayoutId, targetLayoutName: slide.targetLayoutName, semanticNodeCount: slide.nodes.filter((node) => node.visible).length })), fonts, ...composition },
    metrics,
    nativeTextOverflowEvidence: nativeTextOverflowEvidence(nativeMeasurement, candidateAudit),
    designImpact,
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
  const designMode = argument("design-mode") ?? "shared";
  if (!["shared", "template"].includes(designMode)) throw new Error("--design-mode must be shared or template.");
  qualifyStudioWebBenchmark(sourcePath, requestedSlides(argument("slides")), outputRoot, argument("benchmark"), { sourceSha256: argument("source-sha256"), benchmarkSha256: argument("benchmark-sha256") }, { templatePath: argument("template"), designMode: designMode as "shared" | "template" }).then(({ report, reportPath }) => {
    process.stdout.write(`${JSON.stringify({ readyForHumanVisualReview: report.readyForHumanVisualReview, checks: report.checks, candidate: report.candidate.path, report: reportPath, renders: report.renders }, null, 2)}\n`);
    if (!report.readyForHumanVisualReview) process.exitCode = 1;
  }).catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
