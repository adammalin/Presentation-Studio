import JSZip from "jszip";
import type {
  CleanupChange,
  CleanupProposal,
  DeckJob,
  SlideDesignDisposition,
  TableInventoryItem,
  TableNormalizationException,
} from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { auditPptx } from "./pptx-audit";

const TABLE_PROFILE = PRESENTATION_DESIGN_STANDARD.tableProfile;
const TABLE_MARGIN_HORIZONTAL_EMU = Math.round(TABLE_PROFILE.cellPaddingPt.left * 12_700);
const TABLE_MARGIN_VERTICAL_EMU = Math.round(TABLE_PROFILE.cellPaddingPt.top * 12_700);
const TABLE_RULE_WIDTH_EMU = Math.round(TABLE_PROFILE.strokes.horizontal.widthPt * 12_700);
const TABLE_FONT_SIZE = Math.round(TABLE_PROFILE.body.fontSizePt * 100);

function stableChangeId(from: string, to: string): string {
  return `font-${from.toLowerCase().replace(/[^a-z0-9]+/g, "-")}-${to.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
}

function assertDeckReady(deck: DeckJob) {
  if (!deck.audit) throw new Error("Audit the deck before staging cleanup.");
  if (!deck.targetTemplateConfirmedAt || !deck.targetTemplateId) throw new Error("Confirm the target template before staging cleanup.");
  if (deck.audit.containsMacros || deck.audit.containsOleObjects || deck.audit.containsExternalRelationships) {
    throw new Error("Advanced or externally linked content requires manual review before automated cleanup.");
  }
}

function fontCleanupChanges(deck: DeckJob): CleanupChange[] {
  return (deck.audit?.fonts ?? [])
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
}

function tableException(table: TableInventoryItem): TableNormalizationException | undefined {
  const semanticTokens = table.colorTokens.filter((token) => /^accent[1-6]$/.test(token));
  if (semanticTokens.length > 0) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "semantic-color",
    reason: `Preserved meaning-bearing theme color (${semanticTokens.join(", ")}); this table needs a designer check before normalization.`,
  };
  const cellCount = table.rowCount * table.columnCount;
  if (table.mergedCellCount > 10 || (cellCount > 0 && table.mergedCellCount / cellCount > 0.35)) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "complex-structure",
    reason: "Preserved a complex merged-cell topology; normalize it only after a designer confirms hierarchy and reading order.",
  };
  const averageCharacters = cellCount > 0 ? table.totalCellCharacterCount / cellCount : 0;
  if (cellCount > 40 || table.maximumCellCharacterCount > 180 || averageCharacters > 45) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "dense-table",
    reason: `Preserved a dense technical table (${table.totalCellCharacterCount} characters; ${table.maximumCellCharacterCount} in its longest cell); it needs measured overflow and possible continuation-slide review.`,
  };
  return undefined;
}

function slideDispositions(deck: DeckJob, changes: CleanupChange[], exceptions: TableNormalizationException[]): SlideDesignDisposition[] {
  return (deck.audit?.slides ?? []).map((slide) => {
    const changeIds = changes.filter((change) => change.affectedSlideNumbers.includes(slide.number)).map((change) => change.id);
    const slideExceptions = exceptions.filter((exception) => exception.slideNumber === slide.number);
    if (slideExceptions.length > 0 || slide.warnings.length > 0) return {
      slideNumber: slide.number,
      status: "needs-review",
      changeIds,
      reasons: [
        ...slideExceptions.map((exception) => exception.reason),
        ...slide.warnings.map((warning) => `Preview warning: ${warning}`),
        ...(changeIds.length > 0 ? ["Safe deterministic changes are still included in the proposal."] : []),
      ],
    };
    if (changeIds.length > 0) return {
      slideNumber: slide.number,
      status: "change-proposed",
      changeIds,
      reasons: ["A deterministic typography or native-table improvement is included in the proposal."],
    };
    return {
      slideNumber: slide.number,
      status: "approved-as-is",
      changeIds: [],
      reasons: ["The deck-wide deterministic pass found no supported change or blocking exception on this slide."],
    };
  });
}

export function createFontCleanupProposal(deck: DeckJob, updatedAt: string): CleanupProposal {
  assertDeckReady(deck);
  if (deck.operationScope !== "cleanup-only") throw new Error("Font cleanup requires cleanup-only operation scope.");
  const changes = fontCleanupChanges(deck);
  if (changes.length === 0) throw new Error("No supported legacy font mappings were found.");
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Normalize ${changes.length} legacy font famil${changes.length === 1 ? "y" : "ies"} without changing text`,
    status: "pending",
    mode: "font-cleanup",
    changes,
    slideDispositions: slideDispositions(deck, changes, []),
    tableExceptions: [],
  };
}

export function createDesignerCleanupProposal(deck: DeckJob, updatedAt: string): CleanupProposal {
  assertDeckReady(deck);
  if (!deck.audit) throw new Error("Audit the deck before staging cleanup.");
  const changes = fontCleanupChanges(deck);
  if (deck.audit.alignmentRepairs.length > 0) changes.push({
    id: "alignment-cover-dominant-left-edge",
    kind: "alignment",
    from: "offset cover text alignment",
    to: "dominant cover content edge",
    affectedSlideNumbers: [...new Set(deck.audit.alignmentRepairs.map((repair) => repair.slideNumber))].sort((left, right) => left - right),
    affectedRunCount: deck.audit.alignmentRepairs.length,
    alignmentRepairs: deck.audit.alignmentRepairs,
    rationale: "Align high-confidence cover text outliers to the dominant peer edge while preserving text, vertical position, size, and reading order.",
    selected: true,
  });
  const tableExceptions = deck.audit.tables.map(tableException).filter((item): item is TableNormalizationException => Boolean(item));
  const compatibleTables = deck.audit.tables.filter((table) => !tableExceptions.some((exception) => exception.tableId === table.id));
  if (compatibleTables.length > 0) changes.push({
    id: `table-${TABLE_PROFILE.id}`,
    kind: "table-style",
    from: "mixed native table formatting",
    to: "ORNL native table profile",
    affectedSlideNumbers: [...new Set(compatibleTables.map((table) => table.slideNumber))].sort((left, right) => left - right),
    affectedRunCount: compatibleTables.reduce((sum, table) => sum + table.rowCount * table.columnCount, 0),
    tableIds: compatibleTables.map((table) => table.id),
    profileId: TABLE_PROFILE.id,
    rationale: "Normalize compatible native tables to consistent Aptos typography, padding, fills, and minimal horizontal rules while preserving exact cell text and merged-cell structure.",
    selected: true,
  });
  const dispositions = slideDispositions(deck, changes, tableExceptions);
  const changedSlideCount = dispositions.filter((item) => item.status === "change-proposed" || item.changeIds.length > 0).length;
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Designer cleanup reviewed all ${deck.audit.slideCount} slides and proposes supported improvements on ${changedSlideCount}`,
    status: "pending",
    mode: "designer-cleanup",
    standardVersion: PRESENTATION_DESIGN_STANDARD.version,
    changes,
    slideDispositions: dispositions,
    tableExceptions,
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

function setAttribute(attributes: string, name: string, value: string): string {
  const expression = new RegExp(`\\s${name}=(?:"[^"]*"|'[^']*')`, "i");
  return expression.test(attributes) ? attributes.replace(expression, ` ${name}="${value}"`) : `${attributes} ${name}="${value}"`;
}

function directSolidFill(color: string) {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

function normalizeRunPropertyTag(tag: string, header: boolean): string {
  const match = tag.match(/^<a:(rPr|defRPr|endParaRPr)\b([^>]*?)(\/?)>([\s\S]*?)(?:<\/a:\1>)?$/);
  if (!match) return tag;
  const name = match[1];
  let attributes = setAttribute(match[2] ?? "", "sz", String(TABLE_FONT_SIZE));
  if (header) attributes = setAttribute(attributes, "b", "1");
  let children = match[4] ?? "";
  children = children.replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>|<a:solidFill\b[^>]*\/>/g, "");
  children = children.replace(/<a:latin\b[^>]*\/>/g, "");
  const color = header ? TABLE_PROFILE.header.textColor.slice(1) : TABLE_PROFILE.body.textColor.slice(1);
  return `<a:${name}${attributes}>${directSolidFill(color)}<a:latin typeface="${escapeXmlAttribute(TABLE_PROFILE.fontFamily)}"/>${children}</a:${name}>`;
}

function normalizeTextProperties(cell: string, header: boolean): string {
  let result = cell.replace(/<a:(rPr|defRPr|endParaRPr)\b[^>]*\/>|<a:(rPr|defRPr|endParaRPr)\b[^>]*>[\s\S]*?<\/a:\2>/g, (tag) => normalizeRunPropertyTag(tag, header));
  result = result.replace(/<a:r>(?!\s*<a:rPr\b)/g, `<a:r><a:rPr sz="${TABLE_FONT_SIZE}"${header ? ' b="1"' : ""}>${directSolidFill(header ? TABLE_PROFILE.header.textColor.slice(1) : TABLE_PROFILE.body.textColor.slice(1))}<a:latin typeface="${escapeXmlAttribute(TABLE_PROFILE.fontFamily)}"/></a:rPr>`);
  return result;
}

function normalizeBodyProperties(cell: string): string {
  return cell.replace(/<a:bodyPr\b([^>]*?)(?:\/?>)/, (_tag, initial: string) => {
    let attributes = initial ?? "";
    attributes = setAttribute(attributes, "lIns", String(TABLE_MARGIN_HORIZONTAL_EMU));
    attributes = setAttribute(attributes, "rIns", String(TABLE_MARGIN_HORIZONTAL_EMU));
    attributes = setAttribute(attributes, "tIns", String(TABLE_MARGIN_VERTICAL_EMU));
    attributes = setAttribute(attributes, "bIns", String(TABLE_MARGIN_VERTICAL_EMU));
    attributes = setAttribute(attributes, "anchor", "ctr");
    return `<a:bodyPr${attributes}/>`;
  });
}

function normalizeCellProperties(cell: string, fill: string, lastRow: boolean): string {
  const rules = `<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB${lastRow ? "" : ` w="${TABLE_RULE_WIDTH_EMU}"`}>${lastRow ? "<a:noFill/>" : directSolidFill(TABLE_PROFILE.strokes.horizontal.color.slice(1))}</a:lnB>`;
  const replacement = (attributes: string, children: string) => {
    let nextAttributes = attributes;
    nextAttributes = setAttribute(nextAttributes, "marL", String(TABLE_MARGIN_HORIZONTAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marR", String(TABLE_MARGIN_HORIZONTAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marT", String(TABLE_MARGIN_VERTICAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marB", String(TABLE_MARGIN_VERTICAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "anchor", "ctr");
    const cleaned = children
      .replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>|<a:solidFill\b[^>]*\/>/g, "")
      .replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>/g, "");
    return `<a:tcPr${nextAttributes}>${directSolidFill(fill)}${rules}${cleaned}</a:tcPr>`;
  };
  if (/<a:tcPr\b[^>]*\/>/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)\/>/, (_tag, attributes) => replacement(attributes, ""));
  if (/<a:tcPr\b/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)>([\s\S]*?)<\/a:tcPr>/, (_tag, attributes, children) => replacement(attributes, children));
  return cell.replace(/<\/a:tc>/, `${replacement("", "")}</a:tc>`);
}

function normalizeTableBlock(table: string): string {
  let result = table.replace(/<a:tblPr\b([^>]*)>/, (_tag, initial: string) => {
    let attributes = setAttribute(initial ?? "", "firstRow", "1");
    attributes = setAttribute(attributes, "bandRow", "0");
    attributes = setAttribute(attributes, "bandCol", "0");
    return `<a:tblPr${attributes}>`;
  });
  const rows = [...result.matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)];
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index][0];
    const header = index === 0;
    const lastRow = index === rows.length - 1;
    const fill = header ? TABLE_PROFILE.header.fill.slice(1) : (index % 2 === 0 ? TABLE_PROFILE.body.alternateFill : TABLE_PROFILE.body.fill).slice(1);
    const normalized = row.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => normalizeCellProperties(normalizeBodyProperties(normalizeTextProperties(cell, header)), fill, lastRow));
    const start = rows[index].index ?? 0;
    result = result.slice(0, start) + normalized + result.slice(start + row.length);
  }
  return result;
}

function normalizeSelectedTables(xml: string, slideNumber: number, selectedIds: Set<string>): { xml: string; count: number } {
  let ordinal = 0;
  let count = 0;
  return {
    xml: xml.replace(/<a:tbl\b[\s\S]*?<\/a:tbl>/g, (table) => {
      ordinal += 1;
      if (!selectedIds.has(`slide-${slideNumber}-table-${ordinal}`)) return table;
      count += 1;
      return normalizeTableBlock(table);
    }),
    get count() { return count; },
  };
}

function applyAlignmentRepairs(xml: string, slideNumber: number, repairs: CleanupChange["alignmentRepairs"]): { xml: string; count: number } {
  const selected = (repairs ?? []).filter((repair) => repair.slideNumber === slideNumber);
  if (selected.length === 0) return { xml, count: 0 };
  let count = 0;
  const next = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    const shapeId = shape.match(/<p:cNvPr\b[^>]*\bid=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean);
    const repair = selected.find((item) => item.shapeId === shapeId);
    if (!repair) return shape;
    return shape.replace(/<a:off\b([^>]*)\/>/, (tag, initial: string) => {
      const currentX = Number(initial.match(/\bx=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean));
      const currentY = Number(initial.match(/\by=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean));
      if (currentX !== repair.source.x || currentY !== repair.source.y) return tag;
      let attributes = setAttribute(initial, "x", String(repair.target.x));
      attributes = setAttribute(attributes, "y", String(repair.target.y));
      count += 1;
      return `<a:off${attributes}/>`;
    });
  });
  return { xml: next, count };
}

async function materializeCleanup(sourceBytes: Uint8Array, proposal: CleanupProposal, requireAccepted: boolean): Promise<{ bytes: Uint8Array; replacementCount: number; tableCount: number; alignmentCount: number; normalizedTableIds: string[] }> {
  if (requireAccepted && proposal.status !== "applied") throw new Error("Accept the cleanup plan before exporting a review copy.");
  const selected = proposal.changes.filter((change) => change.selected);
  if (selected.length === 0) throw new Error("Select at least one cleanup change.");
  const sourceAudit = await auditPptx(sourceBytes);
  const zip = await JSZip.loadAsync(sourceBytes, { checkCRC32: false });
  let replacementCount = 0;
  let tableCount = 0;
  let alignmentCount = 0;
  const normalizedTableIds = selected.flatMap((change) => change.kind === "table-style" ? change.tableIds ?? [] : []);
  const selectedTableIds = new Set(normalizedTableIds);
  const selectedAlignmentRepairs = selected.flatMap((change) => change.kind === "alignment" ? change.alignmentRepairs ?? [] : []);

  const slidePaths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path));
  for (const path of slidePaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    let xml = await entry.async("text");
    const slideNumber = Number(path.match(/slide(\d+)\.xml$/i)?.[1] ?? 0);
    for (const change of selected.filter((item) => item.kind === "font-family")) {
      const next = replaceTypeface(xml, change.from, change.to);
      xml = next.xml;
      replacementCount += next.replacements;
    }
    const tables = normalizeSelectedTables(xml, slideNumber, selectedTableIds);
    xml = tables.xml;
    tableCount += tables.count;
    const alignments = applyAlignmentRepairs(xml, slideNumber, selectedAlignmentRepairs);
    xml = alignments.xml;
    alignmentCount += alignments.count;
    zip.file(path, xml);
  }
  if (alignmentCount !== selectedAlignmentRepairs.length) throw new Error("Alignment validation failed because a staged text box no longer matched its source geometry.");
  if (replacementCount === 0 && tableCount === 0 && alignmentCount === 0) throw new Error("The selected cleanup changes did not match any editable slide markup.");

  const output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const outputAudit = await auditPptx(output);
  if (outputAudit.slideCount !== sourceAudit.slideCount) throw new Error("Cleanup validation failed because the slide count changed.");
  for (let index = 0; index < sourceAudit.slides.length; index += 1) {
    if (sourceAudit.slides[index].textHash !== outputAudit.slides[index]?.textHash) {
      throw new Error(`Cleanup validation failed because visible text changed on slide ${sourceAudit.slides[index].number}.`);
    }
  }
  for (const sourceTable of sourceAudit.tables) {
    const outputTable = outputAudit.tables.find((table) => table.id === sourceTable.id);
    if (!outputTable || sourceTable.contentHash !== outputTable.contentHash) throw new Error(`Cleanup validation failed because table content changed in ${sourceTable.id}.`);
    if (sourceTable.structureHash !== outputTable.structureHash) throw new Error(`Cleanup validation failed because merged-cell structure changed in ${sourceTable.id}.`);
  }
  return { bytes: output, replacementCount, tableCount, alignmentCount, normalizedTableIds };
}

export async function buildCleanupProposalPptx(sourceBytes: Uint8Array, proposal: CleanupProposal) {
  return materializeCleanup(sourceBytes, proposal, false);
}

export async function applyCleanupToPptx(sourceBytes: Uint8Array, proposal: CleanupProposal) {
  return materializeCleanup(sourceBytes, proposal, true);
}
