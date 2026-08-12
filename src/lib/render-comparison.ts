import type { NativeSlideRender } from "./desktop";

export interface PixelComparisonMetrics {
  exactPixelMatch: boolean;
  changedPixelCount: number;
  changedPixelRatio: number;
  meanAbsoluteChannelDelta: number;
  maximumChannelDelta: number;
  changedBounds?: { x: number; y: number; width: number; height: number; normalized: { x: number; y: number; width: number; height: number } };
}

export function compareRgbaPixels(current: Uint8ClampedArray, proposal: Uint8ClampedArray, width: number, height: number, threshold = 10): PixelComparisonMetrics {
  if (current.length !== proposal.length || current.length !== width * height * 4) throw new Error("Current and Proposal pixel buffers must have identical RGBA dimensions.");
  let changedPixelCount = 0;
  let totalDelta = 0;
  let maximumChannelDelta = 0;
  let minimumX = width;
  let minimumY = height;
  let maximumX = -1;
  let maximumY = -1;
  for (let pixel = 0; pixel < width * height; pixel += 1) {
    const offset = pixel * 4;
    let pixelDelta = 0;
    for (let channel = 0; channel < 3; channel += 1) {
      const delta = Math.abs(current[offset + channel] - proposal[offset + channel]);
      totalDelta += delta;
      pixelDelta = Math.max(pixelDelta, delta);
      maximumChannelDelta = Math.max(maximumChannelDelta, delta);
    }
    if (pixelDelta <= threshold) continue;
    changedPixelCount += 1;
    const x = pixel % width;
    const y = Math.floor(pixel / width);
    minimumX = Math.min(minimumX, x);
    minimumY = Math.min(minimumY, y);
    maximumX = Math.max(maximumX, x);
    maximumY = Math.max(maximumY, y);
  }
  const changedBounds = changedPixelCount > 0 ? {
    x: minimumX,
    y: minimumY,
    width: maximumX - minimumX + 1,
    height: maximumY - minimumY + 1,
    normalized: { x: minimumX / width, y: minimumY / height, width: (maximumX - minimumX + 1) / width, height: (maximumY - minimumY + 1) / height },
  } : undefined;
  return {
    exactPixelMatch: changedPixelCount === 0 && maximumChannelDelta === 0,
    changedPixelCount,
    changedPixelRatio: Number((changedPixelCount / (width * height)).toFixed(6)),
    meanAbsoluteChannelDelta: Number((totalDelta / (width * height * 3)).toFixed(3)),
    maximumChannelDelta,
    changedBounds,
  };
}

async function imageBitmapFor(slide: NativeSlideRender): Promise<ImageBitmap> {
  return createImageBitmap(new Blob([slide.bytes as BlobPart], { type: slide.mimeType }));
}

function canvasBlob(canvas: HTMLCanvasElement, type: "image/jpeg" | "image/png", quality = .9): Promise<Blob> {
  return new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("The native comparison canvas could not be encoded.")), type, quality));
}

export async function compareNativeSlideRenders(current: NativeSlideRender, proposal: NativeSlideRender): Promise<{ metrics: PixelComparisonMetrics; mimeType: "image/jpeg"; bytes: Uint8Array; width: number; height: number }> {
  if (current.width !== proposal.width || current.height !== proposal.height) throw new Error("Current and Proposal native renders have different dimensions.");
  const [currentImage, proposalImage] = await Promise.all([imageBitmapFor(current), imageBitmapFor(proposal)]);
  try {
    const analysisCanvas = document.createElement("canvas");
    analysisCanvas.width = current.width;
    analysisCanvas.height = current.height;
    const analysis = analysisCanvas.getContext("2d", { willReadFrequently: true });
    if (!analysis) throw new Error("The browser could not initialize the native comparison canvas.");
    analysis.drawImage(currentImage, 0, 0);
    const currentPixels = analysis.getImageData(0, 0, current.width, current.height).data;
    analysis.clearRect(0, 0, current.width, current.height);
    analysis.drawImage(proposalImage, 0, 0);
    const proposalPixels = analysis.getImageData(0, 0, current.width, current.height).data;
    const metrics = compareRgbaPixels(currentPixels, proposalPixels, current.width, current.height);

    const outputWidth = 2000;
    const gutter = 24;
    const header = 56;
    const paneWidth = (outputWidth - gutter) / 2;
    const paneHeight = Math.round(paneWidth * current.height / current.width);
    const outputHeight = header + paneHeight;
    const canvas = document.createElement("canvas");
    canvas.width = outputWidth;
    canvas.height = outputHeight;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("The browser could not initialize the comparison output canvas.");
    context.fillStyle = "#F2F5F3";
    context.fillRect(0, 0, outputWidth, outputHeight);
    context.fillStyle = "#00454D";
    context.font = "700 24px Aptos, Arial, sans-serif";
    context.fillText("CURRENT · POWERPOINT NATIVE", 18, 36);
    context.fillText("PROPOSAL · POWERPOINT NATIVE", paneWidth + gutter + 18, 36);
    context.drawImage(currentImage, 0, header, paneWidth, paneHeight);
    context.drawImage(proposalImage, paneWidth + gutter, header, paneWidth, paneHeight);
    if (metrics.changedBounds) {
      const bounds = metrics.changedBounds.normalized;
      context.strokeStyle = "#FF9E1B";
      context.lineWidth = 4;
      context.strokeRect(paneWidth + gutter + bounds.x * paneWidth, header + bounds.y * paneHeight, Math.max(4, bounds.width * paneWidth), Math.max(4, bounds.height * paneHeight));
    }
    const blob = await canvasBlob(canvas, "image/jpeg", .9);
    return { metrics, mimeType: "image/jpeg", bytes: new Uint8Array(await blob.arrayBuffer()), width: outputWidth, height: outputHeight };
  } finally {
    currentImage.close();
    proposalImage.close();
  }
}
