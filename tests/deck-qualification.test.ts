import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import test from "node:test";
import { buildDeckQualificationReport, recordDeckQualificationReviews, type BuildDeckQualificationReportInput, type QualificationRenderSummary } from "../src/lib/deck-qualification";
import type { NativeMeasurementResult } from "../src/lib/desktop";
import type { DesignMetricsReport } from "../src/lib/design-metrics";
import type { PptxAudit } from "../src/types";

const require = createRequire(import.meta.url);
const { captureDeckQualification, finalizeDeckQualification, qualificationReportHtml, readQualificationEvidence } = require("../electron/deck-qualification.cjs") as {
  captureDeckQualification(input: Record<string, unknown>): Promise<{ outputRoot: string; sourceRender: QualificationRenderSummary; candidateRender: QualificationRenderSummary; sourceMeasurement: NativeMeasurementResult; candidateMeasurement: NativeMeasurementResult }>;
  finalizeDeckQualification(input: Record<string, unknown>): Promise<{ reportPath: string; htmlPath: string }>;
  qualificationReportHtml(report: ReturnType<typeof buildDeckQualificationReport>): string;
  readQualificationEvidence(input: Record<string, unknown>): Promise<{ bytes: Uint8Array; mimeType: string }>;
};

function audit(textHashes = ["text-1", "text-2"], tableStructureHash = "table-structure"): PptxAudit {
  return {
    slideCount: 2,
    slides: textHashes.map((textHash, index) => ({ number: index + 1, textHash })),
    tables: [{ slideNumber: 2, contentHash: "table-content", structureHash: tableStructureHash, cellFonts: ["Aptos"] }],
    textBoxes: [{ slideNumber: 1, fontFamilies: ["Aptos"] }, { slideNumber: 2, fontFamilies: ["Aptos"] }],
  } as unknown as PptxAudit;
}

function render(kind: "source" | "candidate", hashes = ["same-title", kind === "source" ? "source-slide-2" : "candidate-slide-2"]): QualificationRenderSummary {
  return {
    status: "ready",
    renderer: "powerpoint-native",
    authoritative: true,
    sourceSha256: kind.repeat(64).slice(0, 64),
    powerPointVersion: "16.111.2",
    pipeline: "powerpoint-save-as-pdf+local-pdf-raster:png:2200px",
    slideCount: 2,
    slides: hashes.map((sha256, index) => ({ number: index + 1, mimeType: "image/png", width: 2200, height: 1238, sha256, relativePath: `${kind}/slide-${String(index + 1).padStart(3, "0")}.png` })),
    warnings: [],
  };
}

function measurement(): NativeMeasurementResult {
  return { status: "ready", adapter: "macos-powerpoint-applescript", authority: "powerpoint-native", slideCount: 2, slides: [], warnings: [], powerPointVersion: "16.111.2" };
}

function metrics(overrides: Partial<DesignMetricsReport["totals"]> = {}): DesignMetricsReport {
  const totals = { safeRegionViolationCount: 0, offSlideObjectCount: 0, textOverflowCount: 0, tableCellClearanceViolationCount: 0, movementCostPt: 0, changedObjectCount: 0, ...overrides };
  return {
    schema: "presentation-studio/design-metrics",
    version: 2,
    measurementRevision: "native-r1",
    sourceSha256: "b".repeat(64),
    totals,
    slides: [1, 2].map((slideNumber) => ({ slideNumber, authority: "powerpoint-native", safeRegionViolationCount: 0, offSlideObjectCount: slideNumber === 2 ? totals.offSlideObjectCount : 0, offSlideObjectIds: slideNumber === 2 && totals.offSlideObjectCount ? ["off-slide-object"] : [], textOverflowCount: slideNumber === 2 ? totals.textOverflowCount : 0, textOverflowObjectIds: slideNumber === 2 && totals.textOverflowCount ? ["overflow-object"] : [], tableCellClearanceViolationCount: slideNumber === 2 ? totals.tableCellClearanceViolationCount : 0, tableCellFindings: [], warnings: [] })),
  };
}

function reportInput(): BuildDeckQualificationReportInput {
  return {
    id: "qualification-test",
    sourceName: "source.pptx",
    candidateName: "candidate.pptx",
    sourceSha256: "a".repeat(64),
    candidateSha256: "b".repeat(64),
    sourceAudit: audit(),
    candidateAudit: audit(),
    sourceRender: render("source"),
    candidateRender: render("candidate"),
    sourceMeasurement: measurement(),
    candidateMeasurement: measurement(),
    candidateMetrics: metrics(),
    protectedSlideNumbers: [1],
    designImpactBySlide: { 1: "unchanged", 2: "layout-redesign" },
    requireMaterialDesignImpact: true,
  };
}

test("deck qualification passes objective gates but keeps visual acceptance explicitly pending", () => {
  const report = buildDeckQualificationReport(reportInput());
  assert.equal(report.status, "visual-review-required");
  assert.equal(report.totals.changedSlides, 1);
  assert.equal(report.checks.protectedSlidesRemainPixelIdentical, true);
  assert.equal(report.visualAcceptance.status, "pending-human-or-authorized-ai-review");
  assert.equal(report.issues.length, 0);
  assert.equal(report.slides[0].protected, true);
});

test("deck qualification routes integrity defects to code and visible fit defects to MCP design", () => {
  const input = reportInput();
  input.candidateAudit = audit(["changed-text", "text-2"], "changed-table-structure");
  input.candidateRender = render("candidate", ["changed-title", "candidate-slide-2"]);
  input.candidateMetrics = metrics({ textOverflowCount: 1, offSlideObjectCount: 1, tableCellClearanceViolationCount: 1 });
  input.designImpactBySlide = { 1: "unchanged", 2: "typography-only" };
  const report = buildDeckQualificationReport(input);
  assert.equal(report.status, "objective-failure");
  assert.ok(report.issues.some((issue) => issue.code === "visible-text" && issue.repairRoute === "engine-code"));
  assert.ok(report.issues.some((issue) => issue.code === "protected-slide-drift" && issue.repairRoute === "engine-code"));
  assert.ok(report.issues.some((issue) => issue.code === "text-overflow" && issue.repairRoute === "mcp-design"));
  assert.ok(report.issues.some((issue) => issue.code === "material-design-impact" && issue.repairRoute === "mcp-design"));
});

test("deck qualification binds deterministic issues to native object regions", () => {
  const input = reportInput();
  input.candidateMetrics = metrics({ textOverflowCount: 1 });
  input.candidateMeasurementPacket = {
    schema: "presentation-studio/native-measurement-packet",
    version: 1,
    status: "ready",
    revision: "native-region",
    sourceSha256: "b".repeat(64),
    adapter: "test",
    authority: "powerpoint-native",
    generatedAt: new Date(0).toISOString(),
    objects: [{ objectId: "overflow-object", shapeId: "7", slideNumber: 2, sourceGeometryPt: { left: 100, top: 80, width: 220, height: 90 }, measuredGeometryPt: { left: 100, top: 80, width: 220, height: 90 }, binding: { method: "shape-id", confidence: "high" }, provenance: { authority: "powerpoint-native", adapter: "test", confidence: "high", note: "test" } }],
    warnings: [],
  };
  const issue = buildDeckQualificationReport(input).issues.find((candidate) => candidate.code === "text-overflow");
  assert.equal(issue?.evidenceRegion?.objectIds[0], "overflow-object");
  assert.ok((issue?.evidenceRegion?.width ?? 0) > 0);
  assert.ok((issue?.evidenceRegion?.height ?? 0) > 0);
  const html = qualificationReportHtml(buildDeckQualificationReport(input));
  assert.match(html, /Show diagnostic regions/);
  assert.match(html, /Native-pixel evidence region/);
});

test("deck qualification routes an unmet governed visual brief back to image concept work", () => {
  const input = reportInput();
  input.designImpactBySlide = { 1: "unchanged", 2: "typography-only" };
  input.visualNeedBySlide = { 2: { id: "need-2", type: "supporting-visual", status: "brief-ready" } };
  const report = buildDeckQualificationReport(input);
  assert.ok(report.issues.some((issue) => issue.slideNumber === 2 && issue.code === "material-design-impact" && issue.repairRoute === "image-concept"));
});

test("deck qualification rejects duplicate render entries and carries global failures onto every slide", () => {
  const input = reportInput();
  input.candidateRender = render("candidate");
  input.candidateRender.slides = [input.candidateRender.slides[0], input.candidateRender.slides[0]];
  input.candidateAudit = audit();
  input.candidateAudit.textBoxes[0].fontFamilies = ["Arial"];
  const report = buildDeckQualificationReport(input);
  assert.equal(report.checks.everySlideHasSourceAndCandidateImage, false);
  assert.equal(report.status, "objective-failure");
  assert.ok(report.slides.every((slide) => slide.status === "objective-failure"));
  assert.ok(report.slides.every((slide) => slide.issueIds.includes("qualification-font-policy")));
});

test("qualification visual review is raster-bound, withholds ready on serious findings, and completes only after every slide", () => {
  const report = buildDeckQualificationReport(reportInput());
  const first = recordDeckQualificationReviews(report, {
    qualificationId: report.id,
    candidateSha256: report.candidate.sha256,
    reviewer: "authorized-ai",
    reviewedAt: new Date(1).toISOString(),
    reviews: [{ slideNumber: 1, sourceRasterSha256: report.slides[0].sourceImage.sha256, candidateRasterSha256: report.slides[0].candidateImage.sha256, verdict: "ready", rationale: "The protected title composition matches the source exactly." }],
  });
  assert.equal(first.status, "visual-review-required");
  assert.equal(first.visualAcceptance.reviewedSlideCount, 1);
  const complete = recordDeckQualificationReviews(first, {
    qualificationId: report.id,
    candidateSha256: report.candidate.sha256,
    reviewer: "authorized-ai",
    reviewedAt: new Date(2).toISOString(),
    reviews: [{ slideNumber: 2, sourceRasterSha256: report.slides[1].sourceImage.sha256, candidateRasterSha256: report.slides[1].candidateImage.sha256, verdict: "ready", rationale: "The candidate has a clearer hierarchy and balanced evidence while retaining source intent." }],
  });
  assert.equal(complete.status, "review-complete");
  assert.equal(complete.visualAcceptance.readySlideCount, 2);

  const serious = recordDeckQualificationReviews(report, {
    qualificationId: report.id,
    candidateSha256: report.candidate.sha256,
    reviewer: "authorized-ai",
    reviews: [{ slideNumber: 2, sourceRasterSha256: report.slides[1].sourceImage.sha256, candidateRasterSha256: report.slides[1].candidateImage.sha256, verdict: "ready", rationale: "A visible alignment defect remains.", findings: [{ category: "alignment", severity: "major", message: "The evidence block is visibly out of alignment.", repairRoute: "mcp-design" }] }],
  });
  assert.equal(serious.slides[1].review?.requestedVerdict, "ready");
  assert.equal(serious.slides[1].review?.recordedVerdict, "revise");
  assert.equal(serious.status, "revision-required");
  assert.throws(() => recordDeckQualificationReviews(report, { qualificationId: report.id, candidateSha256: report.candidate.sha256, reviewer: "authorized-ai", reviews: [{ slideNumber: 1, sourceRasterSha256: "stale", candidateRasterSha256: report.slides[0].candidateImage.sha256, verdict: "ready", rationale: "Stale evidence must fail." }] }), /pixels changed/i);
});

test("qualification iterations record objective trend without claiming aesthetic improvement", () => {
  const first = buildDeckQualificationReport(reportInput());
  const nextInput = reportInput();
  nextInput.id = "qualification-test-2";
  nextInput.previousReport = first;
  const second = buildDeckQualificationReport(nextInput);
  assert.equal(second.iteration.attempt, 2);
  assert.equal(second.iteration.previousQualificationId, first.id);
  assert.equal(second.iteration.objectiveTrend, "unchanged");
  assert.equal(second.status, "visual-review-required");
});

test("qualification capture writes private per-slide PNG pairs and a local HTML review bundle", async () => {
  const parent = await fs.mkdtemp(path.join(os.tmpdir(), "presentation-studio-qualification-test-"));
  const outputRoot = path.join(parent, "run-1");
  let renderCall = 0;
  const fakeRender = async () => {
    renderCall += 1;
    return { status: "ready", renderer: "powerpoint-native", authoritative: true, sourceSha256: String(renderCall).repeat(64).slice(0, 64), powerPointVersion: "test", pipeline: "test-png", slideCount: 2, slides: [1, 2].map((number) => ({ number, mimeType: "image/png", width: 2200, height: 1238, sha256: `${renderCall}-${number}`, bytes: Uint8Array.of(renderCall, number) })), warnings: [] };
  };
  const fakeMeasure = async () => measurement();
  try {
    const capture = await captureDeckQualification({ source: { name: "source.pptx", bytes: Uint8Array.of(1) }, candidate: { name: "candidate.pptx", bytes: Uint8Array.of(2) }, outputRoot, width: 2200, render: fakeRender, measure: fakeMeasure });
    assert.equal((await fs.readFile(path.join(outputRoot, "source", "slide-001.png")))[0], 1);
    assert.equal((await fs.readFile(path.join(outputRoot, "candidate", "slide-002.png")))[0], 2);
    const report = buildDeckQualificationReport({ ...reportInput(), sourceRender: capture.sourceRender, candidateRender: capture.candidateRender, sourceMeasurement: capture.sourceMeasurement, candidateMeasurement: capture.candidateMeasurement, protectedSlideNumbers: [] });
    const finalized = await finalizeDeckQualification({ outputRoot, report });
    assert.match(await fs.readFile(finalized.htmlPath, "utf8"), /Microsoft PowerPoint-native PNGs/i);
    assert.equal((await readQualificationEvidence({ outputRoot, representation: "candidate", slideNumber: 2 })).mimeType, "image/png");
    await assert.rejects(() => readQualificationEvidence({ outputRoot, representation: "other", slideNumber: 1 }), /source or candidate/i);
  } finally {
    await fs.rm(parent, { recursive: true, force: true });
  }
});
