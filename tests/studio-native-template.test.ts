import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { createSyntheticLegacyDeck } from "../scripts/create-synthetic-fixture";
import { buildTemplateCatalog } from "../src/lib/template-catalog";
import { applyStudioNativeTemplateLayouts, canonicalOrnlContentLayout } from "../src/lib/studio-native-template";
import { auditPptx } from "../src/lib/pptx-audit";
import type { StudioWebScene } from "../src/types";

test("chooses the neutral one-column ORNL layout as the shared recipe base", () => {
  const layouts = [
    { id: "title", name: "Title | Standard", category: "title" },
    { id: "image", name: "1-Column Key Image", category: "image" },
    { id: "content", name: "1-Column", category: "content" },
  ] as Parameters<typeof canonicalOrnlContentLayout>[0]["layouts"];
  const catalog = { layouts } as Parameters<typeof canonicalOrnlContentLayout>[0];
  assert.equal(canonicalOrnlContentLayout(catalog).id, "content");
});

test("attaches an actual native template layout graph to editable Studio output", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-native-brand-"));
  const sourcePath = path.join(directory, "source.pptx");
  await createSyntheticLegacyDeck(sourcePath);
  const sourceBytes = new Uint8Array(await fs.readFile(sourcePath));
  const template = await JSZip.loadAsync(sourceBytes);
  const layoutPart = Object.keys(template.files).find((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name));
  assert.ok(layoutPart);
  const layoutXml = await template.file(layoutPart)?.async("text");
  assert.ok(layoutXml);
  template.file(layoutPart, layoutXml.replace(/<p:cSld\b/, "<!-- approved ORNL native layout --><p:cSld"));
  const templateBytes = await template.generateAsync({ type: "uint8array" });
  const catalog = await buildTemplateCatalog(templateBytes, "approved-template.potx");
  const target = catalog.layouts.find((layout) => layout.sourcePart === layoutPart);
  assert.ok(target);
  const scene = { slides: [{ slideNumber: 1, recipe: "ornl-title-content" }] } as StudioWebScene;
  const before = await auditPptx(sourceBytes);
  const result = await applyStudioNativeTemplateLayouts({
    bytes: sourceBytes,
    scene,
    outputSlides: [{ outputSlideNumber: 1, sourceSlideNumber: 1 }],
    templateBytes,
    templateCatalog: catalog,
    defaultLayoutId: target.id,
  });
  const after = await auditPptx(result.bytes);
  assert.equal(result.receipts.length, 1);
  assert.equal(result.receipts[0].strategy, "cloned-template-dependency-graph");
  assert.equal(after.slideCount, before.slideCount);
  assert.equal(after.masterCount, before.masterCount + 1);
  assert.deepEqual(after.slides.map((slide) => slide.textHash), before.slides.map((slide) => slide.textHash));
});
