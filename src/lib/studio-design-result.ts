import type { NativeRenderResult } from "./desktop";
import type { StudioWebScene } from "../types";

export interface StudioSlideBuildResult {
  deckId: string;
  sourceSlideNumber: number;
  slideUpdatedAt: string;
  nativeRender?: NativeRenderResult;
}

/**
 * Produces the one native render surface used by Slides and Studio.
 * Untouched slides retain their source PowerPoint pixels. Redesigned slides
 * appear only when the build matches that exact Studio slide revision.
 */
export function composeLatestStudioNativeRender(
  scene: StudioWebScene | undefined,
  sourceRender: NativeRenderResult | undefined,
  builds: Record<string, StudioSlideBuildResult>,
): NativeRenderResult | undefined {
  if (!scene || sourceRender?.status !== "ready") return sourceRender;
  const sourceByNumber = new Map(sourceRender.slides.map((slide) => [slide.number, slide]));
  const slides = scene.slides.flatMap((studioSlide) => {
    if (studioSlide.status !== "designed" || studioSlide.recipe === "source") {
      const source = sourceByNumber.get(studioSlide.slideNumber);
      return source ? [source] : [];
    }
    const build = builds[`${scene.deckId}:${studioSlide.slideNumber}`];
    const rendered = build?.slideUpdatedAt === studioSlide.updatedAt && build.nativeRender?.status === "ready" && build.nativeRender.authoritative ? build.nativeRender.slides[0] : undefined;
    return rendered ? [{ ...rendered, number: studioSlide.slideNumber }] : [];
  });
  const held = scene.slides.map((slide) => slide.slideNumber).filter((slideNumber) => !slides.some((slide) => slide.number === slideNumber));
  return {
    ...sourceRender,
    slideCount: slides.length,
    slides,
    warnings: held.length ? [...sourceRender.warnings, `Latest converted design is waiting for a matching PowerPoint build on slide${held.length === 1 ? "" : "s"} ${held.join(", ")}.`] : sourceRender.warnings,
    reason: held.length ? "Some converted slide revisions have not been built yet." : undefined,
  };
}

