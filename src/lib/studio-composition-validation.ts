import type { PptxAudit, StudioWebNode, StudioWebScene, StudioWebSlide, TableInventoryItem } from "../types";
import type { StudioCompositionOutputSlide } from "./studio-composition-export";
import { materializeStudioTableContinuationSlides } from "./studio-table-workflow";
import { contentCharacterSignature } from "./content-integrity";

export interface StudioCompositionContentValidation {
  valid: boolean;
  exactSourceContent: boolean;
  exactCandidateContent: boolean;
  exactNativeTableContentAndStructure: boolean;
  errors: string[];
}

function normalized(value: string | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim();
}

function tokens(values: string[]): string[] {
  return values.flatMap((value) => normalized(value).split(" ").filter(Boolean)).sort();
}

function valuesForSceneSlide(slide: StudioWebSlide): string[] {
  return slide.nodes.flatMap((node) => {
    if (!node.visible) return [];
    if (node.kind === "text" && node.text !== undefined) return [node.text];
    if (node.kind === "table" && node.table) return node.table.cells.map((cell) => cell.text);
    return [];
  });
}

function valuesForAuditSlide(audit: PptxAudit, slideNumber: number): string[] {
  return [
    ...audit.textBoxes.filter((box) => box.slideNumber === slideNumber).map((box) => box.text),
    ...audit.tables.filter((table) => table.slideNumber === slideNumber).flatMap((table) => (table.cells ?? [])
      .filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation)
      .map((cell) => cell.text)),
  ];
}

function sceneTableSignature(node: StudioWebNode): string {
  if (!node.table) return "";
  return node.table.cells
    .map((cell) => `${cell.row}:${cell.column}:${cell.rowSpan}:${cell.columnSpan}:${normalized(cell.text)}`)
    .sort()
    .join("|");
}

function auditTableSignature(table: TableInventoryItem): string {
  return (table.cells ?? [])
    .filter((cell) => !cell.horizontalMergeContinuation && !cell.verticalMergeContinuation)
    .map((cell) => `${cell.row}:${cell.column}:${cell.rowSpan}:${cell.columnSpan}:${normalized(cell.text)}`)
    .sort()
    .join("|");
}

function equal(left: string[], right: string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function materializedSlideForOutput(scene: StudioWebScene, output: StudioCompositionOutputSlide): StudioWebSlide | undefined {
  const source = scene.slides.find((slide) => slide.slideNumber === output.sourceSlideNumber);
  if (!source) return undefined;
  const materialized = materializeStudioTableContinuationSlides(scene, source);
  if (!output.continuation) return materialized.length === 1 ? materialized[0]?.slide : undefined;
  return materialized.find((candidate) => candidate.continuation?.segmentOrdinal === output.continuation?.segmentOrdinal)?.slide;
}

/**
 * Validates the intentional source-slide to output-slide mapping instead of
 * assuming that every source slide must always remain exactly one slide.
 * Copy is compared as a token multiset to tolerate semantic atomization while
 * native tables are compared cell-by-cell with row/column/merge topology.
 */
export function validateStudioCompositionContent(input: {
  scene: StudioWebScene;
  sourceAudit: PptxAudit;
  candidateAudit: PptxAudit;
  outputSlides: StudioCompositionOutputSlide[];
}): StudioCompositionContentValidation {
  const errors: string[] = [];
  const outputNumbers = input.outputSlides.map((slide) => slide.outputSlideNumber);
  if (input.candidateAudit.slideCount !== input.outputSlides.length || new Set(outputNumbers).size !== outputNumbers.length) {
    errors.push(`Expected ${input.outputSlides.length} uniquely mapped output slides; candidate contains ${input.candidateAudit.slideCount}.`);
  }

  let exactSourceContent = true;
  for (const sourceSlide of input.scene.slides) {
    const sourceSignature = sourceSlide.contentCoverage.sourceContentSignature;
    const sourceMatches = sourceSignature
      ? sourceSignature === contentCharacterSignature(valuesForSceneSlide(sourceSlide))
      : equal(tokens(valuesForAuditSlide(input.sourceAudit, sourceSlide.slideNumber)), tokens(valuesForSceneSlide(sourceSlide)));
    if (!sourceMatches) {
      exactSourceContent = false;
      errors.push(`Source slide ${sourceSlide.slideNumber}'s Studio scene no longer contains the exact source token inventory.`);
    }
  }

  let exactCandidateContent = true;
  let exactNativeTableContentAndStructure = true;
  for (const output of input.outputSlides) {
    const materialized = materializedSlideForOutput(input.scene, output);
    if (!materialized) {
      exactCandidateContent = false;
      exactNativeTableContentAndStructure = false;
      errors.push(`Output slide ${output.outputSlideNumber} has no materialized Studio source mapping.`);
      continue;
    }
    const expectedValues = valuesForSceneSlide(materialized);
    const candidateValues = valuesForAuditSlide(input.candidateAudit, output.outputSlideNumber);
    if (contentCharacterSignature(expectedValues) !== contentCharacterSignature(candidateValues)) {
      exactCandidateContent = false;
      errors.push(`Output slide ${output.outputSlideNumber} changed or omitted mapped Studio copy from source slide ${output.sourceSlideNumber}.`);
    }
    const expectedTables = materialized.nodes.filter((node) => node.visible && node.kind === "table" && node.table).map(sceneTableSignature).sort();
    const candidateTables = input.candidateAudit.tables.filter((table) => table.slideNumber === output.outputSlideNumber).map(auditTableSignature).sort();
    if (!equal(expectedTables, candidateTables)) {
      exactNativeTableContentAndStructure = false;
      errors.push(`Output slide ${output.outputSlideNumber} changed native table cells, row/column order, or merge topology.`);
    }
  }
  return {
    valid: errors.length === 0,
    exactSourceContent,
    exactCandidateContent,
    exactNativeTableContentAndStructure,
    errors,
  };
}
