import type { SlideRenderCatalog, SlideRenderPreview, TemplatePreviewElement } from "./template-catalog";

function escapeXml(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}

function wrapText(text: string, maximumCharacters: number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split(/\r?\n/)) {
    const words = paragraph.trim().split(/\s+/).filter(Boolean);
    if (words.length === 0) { lines.push(""); continue; }
    let line = words[0];
    for (const word of words.slice(1)) {
      if (`${line} ${word}`.length <= maximumCharacters) line += ` ${word}`;
      else { lines.push(line); line = word; }
    }
    lines.push(line);
  }
  return lines;
}

function fontStack(fontFamily?: string): string {
  const requested = (fontFamily ?? "Aptos").replaceAll('"', "").trim() || "Aptos";
  const fallbacks = /^aptos(?:\s|$)/i.test(requested) ? ["Arial", "sans-serif"] : ["Aptos", "Arial", "sans-serif"];
  return [`"${requested}"`, ...fallbacks.map((family) => family === "sans-serif" ? family : `"${family}"`)].join(", ");
}

function transform(element: TemplatePreviewElement, scale: number): string {
  if (!element.rotation) return "";
  return ` transform="rotate(${element.rotation} ${(element.x + element.width / 2) * scale} ${(element.y + element.height / 2) * scale})"`;
}

function elementSvg(element: TemplatePreviewElement, catalog: SlideRenderCatalog, scale: number): string {
  const rotation = transform(element, scale);
  const x = element.x * scale;
  const y = element.y * scale;
  const width = element.width * scale;
  const height = element.height * scale;
  if (element.kind === "image") {
    const href = element.mediaId ? catalog.media[element.mediaId] : undefined;
    return href ? `<image href="${escapeXml(href)}" x="${x}" y="${y}" width="${width}" height="${height}" preserveAspectRatio="xMidYMid slice"${rotation}/>` : "";
  }
  if (element.kind === "text") {
    const sourceFontSize = (element.fontSize ?? 18) * 12700;
    const fontSize = sourceFontSize * scale;
    const maximumCharacters = Math.max(4, Math.floor(element.width / Math.max(1, sourceFontSize * .54)));
    const lines = wrapText(element.text ?? "", maximumCharacters);
    const lineHeight = fontSize * 1.12;
    const textHeight = lines.length * lineHeight;
    const textX = element.textAlign === "center" ? x + width / 2 : element.textAlign === "right" ? x + width : x;
    const textY = element.verticalAlign === "center" ? y + Math.max(0, (height - textHeight) / 2) : element.verticalAlign === "bottom" ? y + Math.max(0, height - textHeight) : y;
    const anchor = element.textAlign === "center" ? "middle" : element.textAlign === "right" ? "end" : "start";
    const spans = lines.map((line, index) => `<tspan x="${textX}" dy="${index === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`).join("");
    return `<text x="${textX}" y="${textY}" fill="${escapeXml(element.textColor ?? "#373A36")}" font-family="${escapeXml(fontStack(element.fontFamily))}" font-size="${fontSize}" font-weight="${element.fontWeight ?? 400}" text-anchor="${anchor}" dominant-baseline="hanging" opacity="${element.opacity ?? 1}"${rotation}>${spans}</text>`;
  }
  if (element.geometry === "ellipse") return `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}" fill="${escapeXml(element.fill ?? "none")}" stroke="${escapeXml(element.stroke ?? "none")}" stroke-width="${(element.strokeWidth ?? 0) * scale}" opacity="${element.opacity ?? 1}"${rotation}/>`;
  if (element.geometry === "line") return `<line x1="${x}" y1="${y}" x2="${x + width}" y2="${y + height}" stroke="${escapeXml(element.stroke ?? element.fill ?? "#373A36")}" stroke-width="${(element.strokeWidth ?? 12700) * scale}" opacity="${element.opacity ?? 1}"${rotation}/>`;
  return `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${escapeXml(element.fill ?? "none")}" stroke="${escapeXml(element.stroke ?? "none")}" stroke-width="${(element.strokeWidth ?? 0) * scale}" opacity="${element.opacity ?? 1}"${rotation}/>`;
}

export function slidePreviewSvg(catalog: SlideRenderCatalog, slide: SlideRenderPreview, viewportWidth = 1200, fontFaceCss = ""): string {
  const scale = viewportWidth / catalog.slideWidth;
  const viewportHeight = catalog.slideHeight * scale;
  const elements = slide.elements.map((element) => elementSvg(element, catalog, scale)).join("");
  const embeddedFonts = fontFaceCss ? `<defs><style>${fontFaceCss}</style></defs>` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${viewportWidth}" height="${viewportHeight}" viewBox="0 0 ${viewportWidth} ${viewportHeight}">${embeddedFonts}<rect width="${viewportWidth}" height="${viewportHeight}" fill="${escapeXml(slide.background)}"/>${elements}</svg>`;
}

export async function slidePreviewJpeg(catalog: SlideRenderCatalog, slide: SlideRenderPreview, width = 1200, fontFaceCss = ""): Promise<{ data: string; width: number; height: number }> {
  const height = Math.max(1, Math.round(width * catalog.slideHeight / catalog.slideWidth));
  const blob = new Blob([slidePreviewSvg(catalog, slide, width, fontFaceCss)], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    image.decoding = "async";
    image.src = url;
    await image.decode();
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("The local slide preview canvas is unavailable.");
    context.fillStyle = slide.background || "#FFFFFF";
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    const dataUrl = canvas.toDataURL("image/jpeg", .88);
    const data = dataUrl.slice(dataUrl.indexOf(",") + 1);
    if (!data) throw new Error("The local slide preview could not be encoded.");
    return { data, width, height };
  } finally {
    URL.revokeObjectURL(url);
  }
}
