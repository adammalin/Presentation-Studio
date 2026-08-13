import type { DeckJob, PresentationStudioProject } from "../types";

export function buildAuditReport(project: PresentationStudioProject, deck: DeckJob): Uint8Array {
  if (!deck.audit) throw new Error("Audit the deck before exporting a report.");
  const audit = deck.audit;
  const report = {
    schema: "presentation-studio/audit-report",
    version: 1,
    generatedAt: new Date().toISOString(),
    contentBoundary: "No slide text, speaker notes, picture names, descriptions, or Resource bytes are reproduced in this report.",
    project: { id: project.project.id, name: project.project.name, updatedAt: project.project.updatedAt },
    deck: {
      id: deck.id,
      name: deck.name,
      sourceSha256: deck.sourceSha256,
      operationScope: deck.operationScope,
      status: deck.status,
      detectedTemplate: deck.templateClassification,
      targetTemplateId: deck.targetTemplateId ?? null,
      targetTemplateConfirmedAt: deck.targetTemplateConfirmedAt ?? null,
    },
    summary: {
      slideCount: audit.slideCount,
      masterCount: audit.masterCount,
      layoutCount: audit.layoutCount,
      notesCount: audit.notesCount,
      modernCommentCount: audit.modernCommentCount,
      mediaCount: audit.mediaCount,
      tableCount: audit.tableCount,
      chartCount: audit.chartCount,
      pictureCount: audit.pictureCount,
      editableTextBoxCount: (audit.textBoxes ?? []).length,
      stagedAlignmentCandidateCount: (audit.alignmentRepairs ?? []).length,
      geometryReviewCount: (audit.layoutReviews ?? []).length,
      supportLevel: audit.supportLevel,
      containsMacros: audit.containsMacros,
      containsOleObjects: audit.containsOleObjects,
      containsExternalRelationships: audit.containsExternalRelationships,
    },
    fonts: audit.fonts.map((font) => ({
      family: font.family,
      directSlideCount: font.directSlideCount,
      slideNumbers: font.slideNumbers,
      partKinds: font.partKinds,
      isLikelySymbolFont: font.isLikelySymbolFont,
    })),
    tables: audit.tables,
    pictures: audit.pictures.map(({ slideNumber, ordinal, widthEmu, heightEmu, cropped, hasOutline, hasEffect, description }) => ({
      slideNumber,
      ordinal,
      widthEmu: widthEmu ?? null,
      heightEmu: heightEmu ?? null,
      cropped,
      hasOutline,
      hasEffect,
      hasStoredDescription: Boolean(description),
    })),
    textBoxes: (audit.textBoxes ?? []).map(({ text: _text, textHash: _textHash, paragraphs, ...textBox }) => ({
      ...textBox,
      paragraphs: (paragraphs ?? []).map(({ text: _paragraphText, textHash: _paragraphTextHash, ...paragraph }) => paragraph),
    })),
    layoutReviews: audit.layoutReviews ?? [],
    alignmentRepairs: (audit.alignmentRepairs ?? []).map(({ textHash: _textHash, ...repair }) => repair),
    findings: audit.findings,
    warnings: audit.warnings,
    proposal: deck.proposal ? {
      id: deck.proposal.id,
      status: deck.proposal.status,
      summary: deck.proposal.summary,
      changes: deck.proposal.changes.map((change) => ({ id: change.id, kind: change.kind, from: change.from, to: change.to, affectedSlideNumbers: change.affectedSlideNumbers, affectedRunCount: change.affectedRunCount, selected: change.selected })),
    } : null,
  };
  return new TextEncoder().encode(`${JSON.stringify(report, null, 2)}\n`);
}
