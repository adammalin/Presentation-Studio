import type { TemplateCatalog, TemplateLayoutPreview } from "./template-catalog";
import type { StudioCompositionOutputSlide } from "./studio-composition-export";
import type { StudioWebScene } from "../types";
import { cloneTemplateLayoutForSlide, templateLayoutPartSha256, type NativeLayoutCloneReceipt } from "./native-layout-remap";

export interface StudioNativeTemplateResult {
  bytes: Uint8Array;
  receipts: NativeLayoutCloneReceipt[];
  warnings: string[];
}

function normalizedLayoutName(value: string): string {
  return value.trim().toLowerCase().replace(/[–—]/g, "-").replace(/\s+/g, " ");
}

/**
 * The neutral one-column layout is the canonical native base for Studio's
 * semantic recipes. It supplies the approved ORNL master, theme, logo,
 * footer, and recurring artwork without forcing recipe content into a
 * layout-specific multi-column placeholder system.
 */
export function canonicalOrnlContentLayout(catalog: TemplateCatalog): TemplateLayoutPreview {
  const exact = catalog.layouts.find((layout) => normalizedLayoutName(layout.name) === "1-column");
  if (exact) return exact;
  const content = catalog.layouts.find((layout) => layout.category === "content" && !/key image|cut-out|series|bar/i.test(layout.name));
  if (content) return content;
  const fallback = catalog.layouts.find((layout) => layout.category !== "title" && layout.category !== "conclusion");
  if (!fallback) throw new Error("The active ORNL Template Pack has no neutral content layout for Studio recipes.");
  return fallback;
}

export async function applyStudioNativeTemplateLayouts(input: {
  bytes: Uint8Array;
  scene: StudioWebScene;
  outputSlides: StudioCompositionOutputSlide[];
  templateBytes: Uint8Array;
  templateCatalog: TemplateCatalog;
  defaultLayoutId?: string;
}): Promise<StudioNativeTemplateResult> {
  const fallback = input.templateCatalog.layouts.find((layout) => layout.id === input.defaultLayoutId)
    ?? canonicalOrnlContentLayout(input.templateCatalog);
  const layoutHashes = new Map<string, string>();
  const receipts: NativeLayoutCloneReceipt[] = [];
  const warnings: string[] = [];
  let bytes = input.bytes;

  for (const output of input.outputSlides) {
    const studioSlide = input.scene.slides.find((slide) => slide.slideNumber === output.sourceSlideNumber);
    if (!studioSlide || studioSlide.recipe === "source") continue;
    const layout = studioSlide.recipe === "template-layout"
      ? input.templateCatalog.layouts.find((candidate) => candidate.id === studioSlide.targetLayoutId)
      : fallback;
    if (!layout) throw new Error(`Studio slide ${studioSlide.slideNumber} references an ORNL layout that is not installed.`);
    let layoutSha256 = layoutHashes.get(layout.sourcePart);
    if (!layoutSha256) {
      layoutSha256 = await templateLayoutPartSha256(input.templateBytes, layout.sourcePart);
      layoutHashes.set(layout.sourcePart, layoutSha256);
    }
    const result = await cloneTemplateLayoutForSlide({
      sourceBytes: bytes,
      templateBytes: input.templateBytes,
      command: {
        id: `studio-native-layout-${output.outputSlideNumber}`,
        slideNumber: output.outputSlideNumber,
        templateSha256: input.templateCatalog.sha256,
        templateLayoutPart: layout.sourcePart,
        templateLayoutSha256: layoutSha256,
        templateLayoutName: layout.name,
        rationale: `Attach the actual approved ORNL ${layout.name} master, layout, theme, logo, footer, and recurring artwork to Studio output slide ${output.outputSlideNumber}.`,
        author: "ai",
      },
    });
    bytes = result.bytes;
    receipts.push(result.receipt);
    warnings.push(`Output slide ${output.outputSlideNumber}: attached native ORNL layout ${layout.name} (${result.receipt.strategy}).`);
  }

  return { bytes, receipts, warnings };
}
