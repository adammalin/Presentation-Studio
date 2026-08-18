import type { ProjectResource } from "../types";

export const MAX_RESOURCE_TEXT_PAGE_CHARACTERS = 40_000;

function normalized(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

export function extractedResourceText(resource: ProjectResource): { text: string; derivativeSha256: string; truncated: boolean } {
  const derivative = resource.derivatives?.find((item) => item.kind === "extracted-text" && item.bytes?.byteLength);
  if (!derivative?.bytes) throw new Error(`${resource.name} has no embedded extracted-text derivative.`);
  return { text: new TextDecoder().decode(derivative.bytes), derivativeSha256: derivative.sha256, truncated: derivative.truncated };
}

export function resourceTextPage(resource: ProjectResource, offset = 0, maximumCharacters = 20_000) {
  if (resource.mcpAccess !== "text") throw new Error("This Resource has no extracted text available to the active AI session. Turn on AI access and re-list Resources.");
  const extracted = extractedResourceText(resource);
  const safeOffset = Math.max(0, Math.min(Math.floor(offset), extracted.text.length));
  const size = Math.max(1_000, Math.min(MAX_RESOURCE_TEXT_PAGE_CHARACTERS, Math.floor(maximumCharacters)));
  const text = extracted.text.slice(safeOffset, safeOffset + size);
  const nextOffset = safeOffset + text.length < extracted.text.length ? safeOffset + text.length : undefined;
  return {
    text,
    offset: safeOffset,
    characterCount: text.length,
    totalCharacterCount: extracted.text.length,
    nextOffset,
    derivativeSha256: extracted.derivativeSha256,
    extractionTruncated: extracted.truncated,
  };
}

export function assertExactResourceExcerpt(resource: ProjectResource, exactExcerpt: string): void {
  if (resource.mcpAccess !== "text") throw new Error(`${resource.name} has no extracted text available to the active AI session. Turn on AI access and re-list Resources.`);
  const excerpt = normalized(exactExcerpt);
  if (!excerpt) throw new Error(`A non-empty exact source excerpt is required for ${resource.name}.`);
  const source = normalized(extractedResourceText(resource).text);
  if (!source.includes(excerpt)) throw new Error(`The source excerpt for ${resource.name} was not found in its embedded extracted text. Read the current Resource text and submit an exact excerpt.`);
}
