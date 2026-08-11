import JSZip from "jszip";
import type { CleanupChange, CleanupProposal, DeckJob } from "../types";
import { auditPptx } from "./pptx-audit";

function stableChangeId(from: string, to: string): string {
  return `font-${from.toLowerCase().replace(/[^a-z0-9]+/g, "-")}-${to.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
}

export function createFontCleanupProposal(deck: DeckJob, updatedAt: string): CleanupProposal {
  if (!deck.audit) throw new Error("Audit the deck before staging cleanup.");
  if (!deck.targetTemplateConfirmedAt || !deck.targetTemplateId) throw new Error("Confirm the target template before staging cleanup.");
  if (deck.operationScope !== "cleanup-only") throw new Error("Font cleanup requires cleanup-only operation scope.");
  if (deck.audit.containsMacros || deck.audit.containsOleObjects || deck.audit.containsExternalRelationships) {
    throw new Error("Advanced or externally linked content requires manual review before automated cleanup.");
  }
  const changes: CleanupChange[] = deck.audit.fonts
    .filter((font) => font.directSlideCount > 0 && ["century gothic", "arial"].includes(font.normalizedFamily) && !font.isLikelySymbolFont)
    .map((font) => ({
      id: stableChangeId(font.family, "Aptos"),
      kind: "font-family" as const,
      from: font.family,
      to: "Aptos",
      affectedSlideNumbers: font.slideNumbers,
      affectedRunCount: font.directSlideCount,
      rationale: `Normalize legacy ${font.family} markup to the confirmed ORNL Aptos typography while preserving every text string.`,
      selected: true,
    }));
  if (changes.length === 0) throw new Error("No supported legacy font mappings were found.");
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Normalize ${changes.length} legacy font famil${changes.length === 1 ? "y" : "ies"} without changing text`,
    status: "pending",
    changes,
  };
}

function escapeXmlAttribute(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function replaceTypeface(xml: string, from: string, to: string): { xml: string; replacements: number } {
  let replacements = 0;
  const source = from.toLowerCase();
  const result = xml.replace(/\btypeface=("([^"]*)"|'([^']*)')/gi, (whole, quoted: string, doubleValue: string, singleValue: string) => {
    const current = doubleValue ?? singleValue ?? "";
    if (current.trim().toLowerCase() !== source) return whole;
    replacements += 1;
    const quote = quoted.startsWith("'") ? "'" : '"';
    return `typeface=${quote}${escapeXmlAttribute(to)}${quote}`;
  });
  return { xml: result, replacements };
}

export async function applyCleanupToPptx(sourceBytes: Uint8Array, proposal: CleanupProposal): Promise<{ bytes: Uint8Array; replacementCount: number }> {
  if (proposal.status !== "pending") throw new Error("Only a pending proposal can be exported.");
  const selected = proposal.changes.filter((change) => change.selected);
  if (selected.length === 0) throw new Error("Select at least one cleanup change.");
  const sourceAudit = await auditPptx(sourceBytes);
  const zip = await JSZip.loadAsync(sourceBytes, { checkCRC32: false });
  let replacementCount = 0;

  const slidePaths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path));
  for (const path of slidePaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    let xml = await entry.async("text");
    for (const change of selected) {
      const next = replaceTypeface(xml, change.from, change.to);
      xml = next.xml;
      replacementCount += next.replacements;
    }
    zip.file(path, xml);
  }
  if (replacementCount === 0) throw new Error("The selected font mappings did not match any editable slide markup.");

  const output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const outputAudit = await auditPptx(output);
  if (outputAudit.slideCount !== sourceAudit.slideCount) throw new Error("Cleanup validation failed because the slide count changed.");
  for (let index = 0; index < sourceAudit.slides.length; index += 1) {
    if (sourceAudit.slides[index].textHash !== outputAudit.slides[index]?.textHash) {
      throw new Error(`Cleanup validation failed because visible text changed on slide ${sourceAudit.slides[index].number}.`);
    }
  }
  return { bytes: output, replacementCount };
}
