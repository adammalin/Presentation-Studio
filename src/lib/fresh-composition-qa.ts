import type { NativeMeasurementResult } from "./desktop";

export interface FreshCompositionTextOverflow {
  slideNumber: number;
  name: string;
  edges: Array<"left" | "top" | "right" | "bottom">;
}

/**
 * PowerPoint includes up to roughly 2 pt of glyph overhang in rendered bounds
 * for some aligned runs. That is not clipping; larger excursions are.
 */
export function nativeTextOverflows(measurement: NativeMeasurementResult, tolerancePt = 2.25): FreshCompositionTextOverflow[] {
  const failures: FreshCompositionTextOverflow[] = [];
  for (const slide of measurement.slides) for (const shape of slide.shapes) {
    const box = shape.boundsPt;
    const text = shape.text?.renderedBoundsPt;
    const margins = shape.text?.marginsPt;
    // PowerPoint reports the paragraph terminator in an otherwise-empty shape
    // as one character. Generated rules and other decorative shapes therefore
    // have a text frame with length 1 even though they contain no visible copy.
    if (!box || !text || !margins || shape.text?.coordinateSpace !== "slide" || shape.text.textLength <= 1) continue;
    const inner = { left: box.left + margins.left, top: box.top + margins.top, right: box.left + box.width - margins.right, bottom: box.top + box.height - margins.bottom };
    const edges = [
      text.left < inner.left - tolerancePt ? "left" as const : undefined,
      text.top < inner.top - tolerancePt ? "top" as const : undefined,
      text.left + text.width > inner.right + tolerancePt ? "right" as const : undefined,
      text.top + text.height > inner.bottom + tolerancePt ? "bottom" as const : undefined,
    ].filter((edge): edge is FreshCompositionTextOverflow["edges"][number] => Boolean(edge));
    if (edges.length) failures.push({ slideNumber: slide.number, name: shape.name ?? `slide-${slide.number}-shape-${shape.shapeIndex}`, edges });
  }
  return failures;
}
