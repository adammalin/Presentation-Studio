import type { NativeRenderResult, NativeSlideRender } from "./desktop";
import { sha256 } from "./hash";

export const CONTACT_SHEET_PAGE_SIZE = 40;

export interface ContactSheetPlacement {
  slideNumber: number;
  source: NativeSlideRender;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ContactSheetPlan {
  page: number;
  pageCount: number;
  pageSize: number;
  totalSlides: number;
  firstSlideNumber: number;
  lastSlideNumber: number;
  width: number;
  height: number;
  placements: ContactSheetPlacement[];
}

export interface NativeContactSheet {
  id: string;
  kind: "deck-contact-sheet";
  mimeType: "image/png";
  bytes: Uint8Array;
  width: number;
  height: number;
  sha256: string;
  page: number;
  pageCount: number;
  pageSize: number;
  totalSlides: number;
  firstSlideNumber: number;
  lastSlideNumber: number;
  reason: string;
}

export function planContactSheet(slides: NativeSlideRender[], page = 1, pageSize = CONTACT_SHEET_PAGE_SIZE): ContactSheetPlan {
  if (!Number.isInteger(pageSize) || pageSize < 1 || pageSize > 50) throw new Error("Contact-sheet page size must be from 1 to 50 slides.");
  const ordered = [...slides].sort((left, right) => left.number - right.number);
  if (ordered.length === 0) throw new Error("A contact sheet requires at least one native slide render.");
  const pageCount = Math.ceil(ordered.length / pageSize);
  if (!Number.isInteger(page) || page < 1 || page > pageCount) throw new Error(`Choose contact-sheet page 1 to ${pageCount}.`);
  const selected = ordered.slice((page - 1) * pageSize, page * pageSize);
  const columns = Math.min(5, selected.length);
  const gutter = 18;
  const outer = 22;
  const headerHeight = 46;
  const labelHeight = 24;
  const thumbnailWidth = 320;
  const maximumAspectHeight = Math.max(...selected.map((slide) => Math.round(thumbnailWidth * slide.height / slide.width)));
  const cellHeight = maximumAspectHeight + labelHeight;
  const rows = Math.ceil(selected.length / columns);
  const width = outer * 2 + columns * thumbnailWidth + Math.max(0, columns - 1) * gutter;
  const height = headerHeight + outer + rows * cellHeight + Math.max(0, rows - 1) * gutter + outer;
  const placements = selected.map((source, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const height = Math.round(thumbnailWidth * source.height / source.width);
    return {
      slideNumber: source.number,
      source,
      x: outer + column * (thumbnailWidth + gutter),
      y: headerHeight + outer + row * (cellHeight + gutter),
      width: thumbnailWidth,
      height,
    };
  });
  return {
    page,
    pageCount,
    pageSize,
    totalSlides: ordered.length,
    firstSlideNumber: selected[0].number,
    lastSlideNumber: selected.at(-1)!.number,
    width,
    height,
    placements,
  };
}

export async function renderNativeContactSheet(render: NativeRenderResult, page = 1): Promise<NativeContactSheet> {
  if (render.status !== "ready" || !render.authoritative || render.renderer !== "powerpoint-native") throw new Error("A deck contact sheet requires an authoritative PowerPoint-native render.");
  const plan = planContactSheet(render.slides, page);
  const canvas = document.createElement("canvas");
  canvas.width = plan.width;
  canvas.height = plan.height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("The contact-sheet canvas is unavailable.");
  context.fillStyle = "#e7ece9";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#153329";
  context.font = "700 18px Aptos, Arial, sans-serif";
  context.textBaseline = "middle";
  context.fillText(`PowerPoint-native deck review · slides ${plan.firstSlideNumber}–${plan.lastSlideNumber}`, 22, 24);
  context.textAlign = "right";
  context.font = "600 14px Aptos, Arial, sans-serif";
  context.fillText(`Page ${plan.page} of ${plan.pageCount}`, canvas.width - 22, 24);
  context.textAlign = "left";
  for (const placement of plan.placements) {
    const source = new Blob([new Uint8Array(placement.source.bytes).slice().buffer], { type: placement.source.mimeType });
    const bitmap = await createImageBitmap(source);
    context.fillStyle = "#ffffff";
    context.fillRect(placement.x - 1, placement.y - 1, placement.width + 2, placement.height + 2);
    context.drawImage(bitmap, placement.x, placement.y, placement.width, placement.height);
    bitmap.close();
    context.fillStyle = "#ffffff";
    context.fillRect(placement.x, placement.y + placement.height, placement.width, 24);
    context.fillStyle = "#153329";
    context.font = "700 13px Aptos, Arial, sans-serif";
    context.textBaseline = "middle";
    context.fillText(`Slide ${placement.slideNumber}`, placement.x + 8, placement.y + placement.height + 12);
  }
  const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error("Presentation Studio could not encode the deck contact sheet.")), "image/png"));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  return {
    id: `deck-contact-sheet-page-${plan.page}`,
    kind: "deck-contact-sheet",
    mimeType: "image/png",
    bytes,
    width: plan.width,
    height: plan.height,
    sha256: await sha256(bytes),
    page: plan.page,
    pageCount: plan.pageCount,
    pageSize: plan.pageSize,
    totalSlides: plan.totalSlides,
    firstSlideNumber: plan.firstSlideNumber,
    lastSlideNumber: plan.lastSlideNumber,
    reason: "PowerPoint-native overview for deck-level hierarchy, repetition, density, pacing, and cross-slide consistency review.",
  };
}
