import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { buildTemplateCatalog } from "../src/lib/template-catalog";
import { buildTemplatePreviewDeck } from "../src/lib/template-preview-deck";
import { sha256 } from "../src/lib/hash";
import type { NativeRenderResult } from "../src/lib/desktop";

const require = createRequire(import.meta.url);
const { renderPowerPointNative } = require("../electron/native-render.cjs") as { renderPowerPointNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }): Promise<NativeRenderResult> };

function argument(name: string) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

async function renderNative(input: { bytes: Uint8Array; name: string; width: number; format: "png" }) {
  let result = await renderPowerPointNative(input);
  for (let attempt = 2; attempt <= 3 && !(result.status === "ready" && result.renderer === "powerpoint-native" && result.authoritative); attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 350 * attempt));
    result = await renderPowerPointNative(input);
  }
  return result;
}

const templatePath = argument("template");
if (!templatePath) throw new Error("Usage: npm run qualify:ornl-layouts -- --template /absolute/path/to/authorized-template.potx [--output /private/report.json] [--render-dir /private/layout-renders]");
const outputPath = path.resolve(argument("output") ?? "tmp/private-real-deck-tests/ornl-layout-contracts/report.json");
const renderDirectory = argument("render-dir") ? path.resolve(argument("render-dir")!) : undefined;

const templateBytes = new Uint8Array(await fs.readFile(templatePath));
const templateSha256 = await sha256(templateBytes);
const catalog = await buildTemplateCatalog(templateBytes, path.basename(templatePath));
const previewDeck = await buildTemplatePreviewDeck(templateBytes);
const nativeRender = await renderNative({ bytes: previewDeck.bytes, name: `${path.parse(templatePath).name}-layout-contract-preview.pptx`, width: 2200, format: "png" });

const orderMatches = previewDeck.layoutParts.length === catalog.layouts.length && previewDeck.layoutParts.every((part, index) => part === catalog.layouts[index]?.sourcePart);
const nativeReady = nativeRender.status === "ready" && nativeRender.renderer === "powerpoint-native" && nativeRender.authoritative;
const nativeSlideCount = nativeReady ? nativeRender.slides.length : 0;

const layouts = catalog.layouts.map((layout, index) => {
  const semantic = layout.semantic;
  const failures: string[] = [];
  if (!semantic) failures.push("semantic-contract-missing");
  if (semantic && !semantic.contract.nativeAuthority.masterRequired) failures.push("native-master-not-required");
  if (semantic && !semantic.contract.nativeAuthority.layoutRequired) failures.push("native-layout-not-required");
  if (semantic && !semantic.contract.nativeAuthority.preserveInheritedArtwork) failures.push("inherited-artwork-not-protected");
  if (semantic && semantic.contract.family !== "blank" && semantic.contract.compatibleArchetypes.length === 0) failures.push("no-compatible-archetype");
  if (semantic && ["reading", "hero-visual", "comparison", "image-series", "portrait-series"].includes(semantic.contract.family) && !semantic.contract.contentBounds) failures.push("content-bounds-missing");
  if (semantic && ["cover", "conclusion"].includes(semantic.contract.family) && semantic.contract.selectionPolicy !== "sacred") failures.push("sacred-policy-missing");
  if (semantic && ["image-series", "portrait-series"].includes(semantic.contract.family)) {
    const imageSlotIds = semantic.slots.filter((slot) => slot.role === "image").map((slot) => slot.id);
    const coveredImageSlotIds = new Set(semantic.contract.slotGroups.flatMap((group) => group.slotIds));
    if (imageSlotIds.some((slotId) => !coveredImageSlotIds.has(slotId))) failures.push("repeated-image-slot-unbound");
    if (semantic.contract.family === "portrait-series" && semantic.contract.slotGroups.length !== semantic.capabilities.imageSlots) failures.push("portrait-label-group-count-mismatch");
  }
  if (!nativeReady || !nativeRender.slides.some((slide) => slide.number === index + 1)) failures.push("powerpoint-native-render-missing");
  return {
    ordinal: index + 1,
    id: layout.id,
    name: layout.name,
    sourcePart: layout.sourcePart,
    semantic,
    nativeRasterSha256: nativeReady ? nativeRender.slides.find((slide) => slide.number === index + 1)?.sha256 : undefined,
    status: failures.length === 0 ? "qualified" : "failed",
    failures,
  };
});

if (renderDirectory && nativeReady) {
  await fs.mkdir(renderDirectory, { recursive: true });
  for (const slide of nativeRender.slides) {
    const layout = catalog.layouts[slide.number - 1];
    const safeName = (layout?.name ?? `layout-${slide.number}`).replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase();
    await fs.writeFile(path.join(renderDirectory, `${String(slide.number).padStart(2, "0")}-${safeName}.png`), Buffer.from(slide.bytes));
  }
}

const report = {
  schema: "presentation-studio-ornl-layout-contract-qualification-v1",
  generatedAt: new Date().toISOString(),
  template: { name: path.basename(templatePath), sha256: templateSha256, masterCount: catalog.masterCount, layoutCount: catalog.layouts.length },
  nativeRender: {
    ready: nativeReady,
    renderer: nativeRender.renderer,
    authoritative: nativeReady ? nativeRender.authoritative : false,
    powerPointVersion: nativeReady ? nativeRender.powerPointVersion : undefined,
    previewOrderMatchesCatalog: orderMatches,
    slideCount: nativeSlideCount,
  },
  summary: {
    qualified: layouts.filter((layout) => layout.status === "qualified").length,
    failed: layouts.filter((layout) => layout.status === "failed").length,
    families: Object.fromEntries([...new Set(layouts.map((layout) => layout.semantic?.contract.family ?? "missing"))].map((family) => [family, layouts.filter((layout) => (layout.semantic?.contract.family ?? "missing") === family).length])),
    archetypes: Object.fromEntries([...new Set(layouts.flatMap((layout) => layout.semantic?.contract.compatibleArchetypes ?? []))].map((archetype) => [archetype, layouts.filter((layout) => layout.semantic?.contract.compatibleArchetypes.includes(archetype)).length])),
  },
  layouts,
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`);
if (!orderMatches) throw new Error(`Template preview order did not match the semantic catalog. Report: ${outputPath}`);
if (!nativeReady || nativeSlideCount !== catalog.layouts.length) throw new Error(`Microsoft PowerPoint did not authoritatively render every template layout. Report: ${outputPath}`);
if (report.summary.failed > 0) throw new Error(`${report.summary.failed} template layout contract${report.summary.failed === 1 ? "" : "s"} failed qualification. Report: ${outputPath}`);
console.log(JSON.stringify({ outputPath, renderDirectory, template: report.template, nativeRender: report.nativeRender, summary: report.summary }, null, 2));
