import type { NativeRenderResult } from "./desktop";
import type { StudioWebScene } from "../types";

export interface StudioSlideBuildResult {
  deckId: string;
  sourceSlideNumber: number;
  slideUpdatedAt: string;
  nativeRender?: NativeRenderResult;
}

function acceptedRenderedSlide(studioSlide: StudioWebScene["slides"][number], build: StudioSlideBuildResult | undefined) {
  const rendered = build?.slideUpdatedAt === studioSlide.updatedAt && build.nativeRender?.status === "ready" && build.nativeRender.authoritative ? build.nativeRender.slides[0] : undefined;
  const review = studioSlide.qualityReview;
  if (!rendered || review?.recordedVerdict !== "ready" || review.slideUpdatedAt !== studioSlide.updatedAt || review.rasterSha256 !== rendered.sha256) return undefined;
  return rendered;
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
  const waitingForBuild: number[] = [];
  const waitingForAcceptance: number[] = [];
  const missingSource: number[] = [];
  const slides = scene.slides.flatMap((studioSlide) => {
    const source = sourceByNumber.get(studioSlide.slideNumber);
    if (studioSlide.status !== "designed" || studioSlide.recipe === "source") {
      return source ? [source] : [];
    }
    const build = builds[`${scene.deckId}:${studioSlide.slideNumber}`];
    const exactBuild = build?.slideUpdatedAt === studioSlide.updatedAt && build.nativeRender?.status === "ready" && build.nativeRender.authoritative;
    const rendered = acceptedRenderedSlide(studioSlide, build);
    if (rendered) return [{ ...rendered, number: studioSlide.slideNumber }];
    if (exactBuild) waitingForAcceptance.push(studioSlide.slideNumber);
    else waitingForBuild.push(studioSlide.slideNumber);
    if (source) return [source];
    missingSource.push(studioSlide.slideNumber);
    return [];
  });
  const warnings = [...sourceRender.warnings];
  if (waitingForBuild.length) warnings.push(`Candidate design is waiting for a matching PowerPoint build on slide${waitingForBuild.length === 1 ? "" : "s"} ${waitingForBuild.join(", ")}; the faithful source remains visible.`);
  if (waitingForAcceptance.length) warnings.push(`Candidate design is waiting for revision-bound visual acceptance on slide${waitingForAcceptance.length === 1 ? "" : "s"} ${waitingForAcceptance.join(", ")}; the faithful source remains visible.`);
  if (missingSource.length) warnings.push(`No source or accepted candidate render is available for slide${missingSource.length === 1 ? "" : "s"} ${missingSource.join(", ")}.`);
  return {
    ...sourceRender,
    slideCount: slides.length,
    slides,
    warnings,
    reason: waitingForBuild.length || waitingForAcceptance.length || missingSource.length ? "Unaccepted candidate revisions remain held behind faithful source pixels." : undefined,
  };
}
