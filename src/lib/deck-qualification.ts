import type { NativeMeasurementResult } from "./desktop";
import type { DesignMetricsReport } from "./design-metrics";
import type { NativeMeasurementPacket } from "./native-measurement";
import type { PptxAudit } from "../types";

export const DECK_QUALIFICATION_SCHEMA = "presentation-studio/deck-qualification" as const;

export type QualificationRepairRoute = "mcp-design" | "engine-code" | "image-concept" | "human-review";
export type QualificationReviewVerdict = "ready" | "revise" | "hold";

export interface QualificationEvidenceRegion {
  x: number;
  y: number;
  width: number;
  height: number;
  objectIds: string[];
  label: string;
}

export interface DeckQualificationSlideReview {
  slideNumber: number;
  sourceRasterSha256: string;
  candidateRasterSha256: string;
  reviewer: "human" | "authorized-ai";
  requestedVerdict: QualificationReviewVerdict;
  recordedVerdict: QualificationReviewVerdict;
  pass: number;
  rationale: string;
  findings: Array<{
    category: "hierarchy" | "alignment" | "spacing" | "layout-balance" | "table-quality" | "figure-clarity" | "template-fidelity" | "deck-consistency" | "source-intent" | "other";
    severity: "major" | "minor";
    message: string;
    repairRoute: QualificationRepairRoute;
  }>;
  reviewedAt: string;
}

export interface QualificationRenderSlide {
  number: number;
  mimeType: "image/png";
  width: number;
  height: number;
  sha256: string;
  relativePath: string;
}

export interface QualificationRenderSummary {
  status: "ready" | "unavailable" | "permission-required" | "failed";
  renderer: "powerpoint-native" | "studio-approximate";
  authoritative: boolean;
  sourceSha256?: string;
  powerPointVersion?: string;
  pipeline?: string;
  slideCount?: number;
  slides: QualificationRenderSlide[];
  warnings: string[];
  reason?: string;
}

export interface DeckQualificationIssue {
  id: string;
  slideNumber?: number;
  severity: "blocker" | "major" | "minor";
  category: "content" | "render" | "typography" | "geometry" | "table" | "template" | "design-impact";
  code: string;
  message: string;
  evidence: string;
  repairRoute: QualificationRepairRoute;
  evidenceRegion?: QualificationEvidenceRegion;
}

export interface DeckQualificationReport {
  schema: typeof DECK_QUALIFICATION_SCHEMA;
  version: 1;
  id: string;
  generatedAt: string;
  status: "objective-failure" | "visual-review-required" | "review-complete" | "revision-required" | "held";
  sceneRevision?: string;
  source: { name: string; sha256: string; slideCount: number };
  candidate: { name: string; sha256: string; slideCount: number };
  iteration: {
    attempt: number;
    previousQualificationId?: string;
    previousCandidateSha256?: string;
    objectiveTrend: "first-run" | "improved" | "regressed" | "unchanged" | "mixed";
    blockerDelta: number;
    majorDelta: number;
    automaticPassLimit: 3;
  };
  nativeEvidence: {
    renderer: "powerpoint-native";
    powerPointVersion?: string;
    pipeline?: string;
    widthPx: number;
    sourceMeasurementAuthority: NativeMeasurementResult["authority"];
    candidateMeasurementAuthority: NativeMeasurementResult["authority"];
  };
  checks: Record<string, boolean>;
  totals: {
    slides: number;
    changedSlides: number;
    protectedSlides: number;
    blockerIssues: number;
    majorIssues: number;
    minorIssues: number;
    textOverflowCount: number;
    offSlideObjectCount: number;
    tableCellClearanceViolationCount: number;
  };
  issues: DeckQualificationIssue[];
  slides: Array<{
    slideNumber: number;
    sourceImage: QualificationRenderSlide;
    candidateImage: QualificationRenderSlide;
    pixelsChanged: boolean;
    protected: boolean;
    designImpact?: string;
    status: "objective-failure" | "visual-review-required" | "ready" | "revise" | "hold";
    issueIds: string[];
    review?: DeckQualificationSlideReview;
  }>;
  repairRouting: Array<{ route: QualificationRepairRoute; issueIds: string[] }>;
  visualAcceptance: {
    status: "pending-human-or-authorized-ai-review" | "review-complete" | "revision-required" | "held";
    rule: string;
    requiredJudgments: string[];
    reviewedSlideCount: number;
    readySlideCount: number;
    revisionSlideCount: number;
    heldSlideCount: number;
    reviews: DeckQualificationSlideReview[];
  };
}

export interface BuildDeckQualificationReportInput {
  id: string;
  generatedAt?: string;
  sceneRevision?: string;
  sourceName: string;
  candidateName: string;
  sourceSha256: string;
  candidateSha256: string;
  sourceAudit: PptxAudit;
  candidateAudit: PptxAudit;
  sourceRender: QualificationRenderSummary;
  candidateRender: QualificationRenderSummary;
  sourceMeasurement: NativeMeasurementResult;
  candidateMeasurement: NativeMeasurementResult;
  candidateMeasurementPacket?: NativeMeasurementPacket;
  candidateMetrics: DesignMetricsReport;
  protectedSlideNumbers?: number[];
  designImpactBySlide?: Record<number, string>;
  visualNeedBySlide?: Record<number, { id: string; type: string; status: string }>;
  requireMaterialDesignImpact?: boolean;
  previousReport?: DeckQualificationReport;
}

function tableSignature(audit: PptxAudit) {
  return audit.tables.map((table) => `${table.slideNumber}:${table.contentHash}:${table.structureHash}`).sort().join("|");
}

function textSignature(audit: PptxAudit) {
  return audit.slides.map((slide) => `${slide.number}:${slide.textHash}`).join("|");
}

function approvedPresentationFont(family: string) {
  return /^Aptos(?:\s|$)/i.test(family) || ["Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Cambria Math"].includes(family);
}

function safeIssueId(code: string, slideNumber?: number) {
  return `qualification-${code}${slideNumber ? `-slide-${slideNumber}` : ""}`;
}

function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(maximum, Math.max(minimum, value));
}

function evidenceRegion(input: BuildDeckQualificationReportInput, slideNumber: number, objectIds: string[], label: string, paddingPt = 14): QualificationEvidenceRegion | undefined {
  const packet = input.candidateMeasurementPacket;
  const slideWidthPt = input.candidateAudit.slideSize?.width ? input.candidateAudit.slideSize.width / 12_700 : 960;
  const slideHeightPt = input.candidateAudit.slideSize?.height ? input.candidateAudit.slideSize.height / 12_700 : 540;
  if (!packet || packet.authority !== "powerpoint-native" || slideWidthPt <= 0 || slideHeightPt <= 0) return undefined;
  const requested = new Set(objectIds);
  const objects = packet.objects.filter((object) => object.slideNumber === slideNumber && object.measuredGeometryPt && (requested.has(object.objectId) || Boolean(object.tableId && requested.has(object.tableId))));
  if (!objects.length) return undefined;
  const left = Math.min(...objects.map((object) => object.measuredGeometryPt!.left)) - paddingPt;
  const top = Math.min(...objects.map((object) => object.measuredGeometryPt!.top)) - paddingPt;
  const right = Math.max(...objects.map((object) => object.measuredGeometryPt!.left + object.measuredGeometryPt!.width)) + paddingPt;
  const bottom = Math.max(...objects.map((object) => object.measuredGeometryPt!.top + object.measuredGeometryPt!.height)) + paddingPt;
  const clippedLeft = clamp(left, 0, slideWidthPt);
  const clippedTop = clamp(top, 0, slideHeightPt);
  const clippedRight = clamp(right, 0, slideWidthPt);
  const clippedBottom = clamp(bottom, 0, slideHeightPt);
  const minimumWidth = Math.min(slideWidthPt, 48);
  const minimumHeight = Math.min(slideHeightPt, 36);
  const regionWidth = Math.max(minimumWidth, clippedRight - clippedLeft);
  const regionHeight = Math.max(minimumHeight, clippedBottom - clippedTop);
  const normalizedWidth = Math.min(1, regionWidth / slideWidthPt);
  const normalizedHeight = Math.min(1, regionHeight / slideHeightPt);
  return {
    x: clamp(clippedLeft / slideWidthPt, 0, 1 - normalizedWidth),
    y: clamp(clippedTop / slideHeightPt, 0, 1 - normalizedHeight),
    width: normalizedWidth,
    height: normalizedHeight,
    objectIds: objects.map((object) => object.objectId),
    label,
  };
}

function fullSlideRegion(label: string): QualificationEvidenceRegion {
  return { x: 0, y: 0, width: 1, height: 1, objectIds: [], label };
}

function objectiveTrend(previous: DeckQualificationReport | undefined, blockerIssues: number, majorIssues: number) {
  if (!previous) return "first-run" as const;
  const blockerDelta = blockerIssues - previous.totals.blockerIssues;
  const majorDelta = majorIssues - previous.totals.majorIssues;
  if (blockerDelta === 0 && majorDelta === 0) return "unchanged" as const;
  if (blockerDelta <= 0 && majorDelta <= 0) return "improved" as const;
  if (blockerDelta >= 0 && majorDelta >= 0) return "regressed" as const;
  return "mixed" as const;
}

export function buildDeckQualificationReport(input: BuildDeckQualificationReportInput): DeckQualificationReport {
  const sourceRenderReady = input.sourceRender.status === "ready" && input.sourceRender.renderer === "powerpoint-native" && input.sourceRender.authoritative;
  const candidateRenderReady = input.candidateRender.status === "ready" && input.candidateRender.renderer === "powerpoint-native" && input.candidateRender.authoritative;
  const sourceMeasurementReady = input.sourceMeasurement.status === "ready" && input.sourceMeasurement.authority === "powerpoint-native";
  const candidateMeasurementReady = input.candidateMeasurement.status === "ready" && input.candidateMeasurement.authority === "powerpoint-native";
  const protectedSlides = new Set(input.protectedSlideNumbers ?? []);
  const sourceRasters = new Map(input.sourceRender.slides.map((slide) => [slide.number, slide]));
  const candidateRasters = new Map(input.candidateRender.slides.map((slide) => [slide.number, slide]));
  const issues: DeckQualificationIssue[] = [];
  const addIssue = (issue: Omit<DeckQualificationIssue, "id">) => {
    const id = safeIssueId(issue.code, issue.slideNumber);
    if (!issues.some((candidate) => candidate.id === id)) issues.push({ id, ...issue });
  };

  const exactSlideCount = input.sourceAudit.slideCount === input.candidateAudit.slideCount;
  const exactVisibleText = exactSlideCount && textSignature(input.sourceAudit) === textSignature(input.candidateAudit);
  const exactTableStructure = tableSignature(input.sourceAudit) === tableSignature(input.candidateAudit);
  const allSlidesRendered = sourceRenderReady && candidateRenderReady
    && input.sourceRender.slides.length === input.sourceAudit.slideCount
    && input.candidateRender.slides.length === input.candidateAudit.slideCount
    && input.sourceAudit.slides.every((slide) => sourceRasters.has(slide.number))
    && input.candidateAudit.slides.every((slide) => candidateRasters.has(slide.number));
  const candidateUsesApprovedFonts = input.candidateAudit.textBoxes.every((box) => box.fontFamilies.every(approvedPresentationFont))
    && input.candidateAudit.tables.every((table) => table.cellFonts.every(approvedPresentationFont));
  const protectedSlidesUnchanged = [...protectedSlides].every((slideNumber) => {
    const source = sourceRasters.get(slideNumber);
    const candidate = candidateRasters.get(slideNumber);
    return Boolean(source && candidate && source.sha256 === candidate.sha256);
  });

  if (!exactSlideCount) addIssue({ severity: "blocker", category: "content", code: "slide-count", message: "The candidate slide count changed.", evidence: `Source has ${input.sourceAudit.slideCount} slides; candidate has ${input.candidateAudit.slideCount}.`, repairRoute: "engine-code" });
  if (!exactVisibleText) addIssue({ severity: "blocker", category: "content", code: "visible-text", message: "Exact visible source content changed in the candidate.", evidence: "The ordered per-slide visible-text hashes do not match.", repairRoute: "engine-code" });
  if (!exactTableStructure) addIssue({ severity: "blocker", category: "table", code: "table-structure", message: "Native table content or merged structure changed.", evidence: "The ordered table content/structure signatures do not match the source.", repairRoute: "engine-code" });
  if (!sourceRenderReady) addIssue({ severity: "blocker", category: "render", code: "source-render", message: "The source deck did not render authoritatively in Microsoft PowerPoint.", evidence: input.sourceRender.reason ?? (input.sourceRender.warnings.join(" ") || "PowerPoint-native source evidence is unavailable."), repairRoute: "engine-code" });
  if (!candidateRenderReady) addIssue({ severity: "blocker", category: "render", code: "candidate-render", message: "The candidate deck did not render authoritatively in Microsoft PowerPoint.", evidence: input.candidateRender.reason ?? (input.candidateRender.warnings.join(" ") || "PowerPoint-native candidate evidence is unavailable."), repairRoute: "engine-code" });
  if (!sourceMeasurementReady || !candidateMeasurementReady) addIssue({ severity: "blocker", category: "render", code: "native-measurement", message: "PowerPoint-native geometry measurement is incomplete.", evidence: `Source authority: ${input.sourceMeasurement.authority}; candidate authority: ${input.candidateMeasurement.authority}.`, repairRoute: "engine-code" });
  if (!allSlidesRendered) addIssue({ severity: "blocker", category: "render", code: "render-count", message: "The evidence bundle is missing one or more uniquely numbered slide images.", evidence: `Source images: ${input.sourceRender.slides.length} (${sourceRasters.size} unique); candidate images: ${input.candidateRender.slides.length} (${candidateRasters.size} unique).`, repairRoute: "engine-code" });
  if (!candidateUsesApprovedFonts) addIssue({ severity: "major", category: "typography", code: "font-policy", message: "The candidate contains a presentation font outside the Aptos and approved symbol families.", evidence: [...new Set([...input.candidateAudit.textBoxes.flatMap((box) => box.fontFamilies), ...input.candidateAudit.tables.flatMap((table) => table.cellFonts)])].filter((family) => !approvedPresentationFont(family)).join(", ") || "Unapproved font family detected.", repairRoute: "mcp-design" });

  for (const slide of input.candidateMetrics.slides) {
    if (slide.textOverflowCount > 0) addIssue({ slideNumber: slide.slideNumber, severity: "blocker", category: "geometry", code: "text-overflow", message: "PowerPoint measured text outside its frame.", evidence: `${slide.textOverflowCount} native text frame${slide.textOverflowCount === 1 ? "" : "s"} overflow.`, repairRoute: "mcp-design", evidenceRegion: evidenceRegion(input, slide.slideNumber, slide.textOverflowObjectIds, "PowerPoint text-overflow region") });
    if (slide.offSlideObjectCount > 0) addIssue({ slideNumber: slide.slideNumber, severity: "blocker", category: "geometry", code: "off-slide", message: "One or more objects extend beyond the slide canvas.", evidence: `${slide.offSlideObjectCount} measured object${slide.offSlideObjectCount === 1 ? "" : "s"} are off-slide.`, repairRoute: "mcp-design", evidenceRegion: evidenceRegion(input, slide.slideNumber, slide.offSlideObjectIds, "Off-slide object region") });
    if (slide.tableCellClearanceViolationCount > 0) {
      const tableIds = [...new Set(slide.tableCellFindings.filter((finding) => finding.rule === "insufficient-clearance").map((finding) => finding.tableId))];
      addIssue({ slideNumber: slide.slideNumber, severity: "major", category: "table", code: "table-clearance", message: "One or more table cells do not meet the resolved native clearance requirement.", evidence: `${slide.tableCellClearanceViolationCount} measured cell-clearance violation${slide.tableCellClearanceViolationCount === 1 ? "" : "s"}.`, repairRoute: "mcp-design", evidenceRegion: evidenceRegion(input, slide.slideNumber, tableIds, "Native table-clearance region", 18) });
    }
  }

  for (const slideNumber of protectedSlides) {
    const source = sourceRasters.get(slideNumber);
    const candidate = candidateRasters.get(slideNumber);
    if (!source || !candidate || source.sha256 !== candidate.sha256) addIssue({ slideNumber, severity: "blocker", category: "template", code: "protected-slide-drift", message: "A protected template slide changed visually.", evidence: "The source and candidate PowerPoint-native raster hashes differ.", repairRoute: "engine-code", evidenceRegion: fullSlideRegion("Protected template composition") });
  }

  const slides = input.sourceAudit.slides.map((slide) => {
    const sourceImage = sourceRasters.get(slide.number);
    const candidateImage = candidateRasters.get(slide.number);
    if (!sourceImage || !candidateImage) return undefined;
    const protectedSlide = protectedSlides.has(slide.number);
    const designImpact = input.designImpactBySlide?.[slide.number];
    const visualNeed = input.visualNeedBySlide?.[slide.number];
    const pixelsChanged = sourceImage.sha256 !== candidateImage.sha256;
    if (input.requireMaterialDesignImpact && !protectedSlide && (!pixelsChanged || !designImpact || ["unchanged", "typography-only", "cleanup"].includes(designImpact))) {
      const repairRoute: QualificationRepairRoute = visualNeed?.status === "brief-ready" ? "image-concept" : visualNeed?.status === "held" ? "human-review" : "mcp-design";
      addIssue({ slideNumber: slide.number, severity: "major", category: "design-impact", code: "material-design-impact", message: "This redesign does not yet show a material whole-slide composition improvement.", evidence: `PowerPoint pixels changed: ${pixelsChanged ? "yes" : "no"}; recorded impact: ${designImpact ?? "unrecorded"}${visualNeed ? `; visual need: ${visualNeed.type} (${visualNeed.status})` : ""}.`, repairRoute, evidenceRegion: fullSlideRegion("Whole-slide design impact") });
    }
    const slideIssueIds = issues.filter((issue) => issue.slideNumber === slide.number || issue.slideNumber === undefined && issue.severity !== "minor").map((issue) => issue.id);
    return {
      slideNumber: slide.number,
      sourceImage,
      candidateImage,
      pixelsChanged,
      protected: protectedSlide,
      designImpact,
      status: slideIssueIds.length ? "objective-failure" as const : "visual-review-required" as const,
      issueIds: slideIssueIds,
    };
  }).filter((slide): slide is NonNullable<typeof slide> => Boolean(slide));

  const checks = {
    exactSlideCount,
    exactVisibleText,
    exactNativeTableContentAndStructure: exactTableStructure,
    sourcePowerPointRenderReady: sourceRenderReady,
    candidatePowerPointRenderReady: candidateRenderReady,
    sourcePowerPointMeasurementReady: sourceMeasurementReady,
    candidatePowerPointMeasurementReady: candidateMeasurementReady,
    everySlideHasSourceAndCandidateImage: allSlidesRendered,
    candidateUsesAptosOrApprovedSymbolFonts: candidateUsesApprovedFonts,
    noNativeTextOverflow: input.candidateMetrics.totals.textOverflowCount === 0,
    noOffSlideObjects: input.candidateMetrics.totals.offSlideObjectCount === 0,
    noNativeTableCellClearanceViolations: input.candidateMetrics.totals.tableCellClearanceViolationCount === 0,
    protectedSlidesRemainPixelIdentical: protectedSlidesUnchanged,
  };
  const requiredChecksPass = Object.values(checks).every(Boolean);
  const repairRouting = (["mcp-design", "engine-code", "image-concept", "human-review"] as QualificationRepairRoute[]).map((route) => ({ route, issueIds: issues.filter((issue) => issue.repairRoute === route).map((issue) => issue.id) })).filter((route) => route.issueIds.length);
  const totals = {
    slides: input.sourceAudit.slideCount,
    changedSlides: slides.filter((slide) => slide.pixelsChanged).length,
    protectedSlides: protectedSlides.size,
    blockerIssues: issues.filter((issue) => issue.severity === "blocker").length,
    majorIssues: issues.filter((issue) => issue.severity === "major").length,
    minorIssues: issues.filter((issue) => issue.severity === "minor").length,
    textOverflowCount: input.candidateMetrics.totals.textOverflowCount,
    offSlideObjectCount: input.candidateMetrics.totals.offSlideObjectCount,
    tableCellClearanceViolationCount: input.candidateMetrics.totals.tableCellClearanceViolationCount,
  };
  const previous = input.previousReport;
  return {
    schema: DECK_QUALIFICATION_SCHEMA,
    version: 1,
    id: input.id,
    generatedAt: input.generatedAt ?? new Date().toISOString(),
    status: requiredChecksPass && !issues.some((issue) => issue.severity !== "minor") ? "visual-review-required" : "objective-failure",
    sceneRevision: input.sceneRevision,
    source: { name: input.sourceName, sha256: input.sourceSha256, slideCount: input.sourceAudit.slideCount },
    candidate: { name: input.candidateName, sha256: input.candidateSha256, slideCount: input.candidateAudit.slideCount },
    iteration: {
      attempt: (previous?.iteration.attempt ?? 0) + 1,
      previousQualificationId: previous?.id,
      previousCandidateSha256: previous?.candidate.sha256,
      objectiveTrend: objectiveTrend(previous, totals.blockerIssues, totals.majorIssues),
      blockerDelta: totals.blockerIssues - (previous?.totals.blockerIssues ?? totals.blockerIssues),
      majorDelta: totals.majorIssues - (previous?.totals.majorIssues ?? totals.majorIssues),
      automaticPassLimit: 3,
    },
    nativeEvidence: {
      renderer: "powerpoint-native",
      powerPointVersion: input.candidateRender.powerPointVersion ?? input.candidateMeasurement.powerPointVersion,
      pipeline: input.candidateRender.pipeline,
      widthPx: input.candidateRender.slides[0]?.width ?? input.sourceRender.slides[0]?.width ?? 0,
      sourceMeasurementAuthority: input.sourceMeasurement.authority,
      candidateMeasurementAuthority: input.candidateMeasurement.authority,
    },
    checks,
    totals,
    issues,
    slides,
    repairRouting,
    visualAcceptance: {
      status: "pending-human-or-authorized-ai-review",
      rule: "Inspect every candidate PNG at full-slide size and compare it with the source. Objective checks make the deck reviewable; they do not prove that the design is better.",
      requiredJudgments: ["hierarchy", "alignment", "spacing", "layout balance", "table quality", "figure clarity", "template fidelity", "deck consistency", "source intent"],
      reviewedSlideCount: 0,
      readySlideCount: 0,
      revisionSlideCount: 0,
      heldSlideCount: 0,
      reviews: [],
    },
  };
}

export function recordDeckQualificationReviews(report: DeckQualificationReport, input: {
  qualificationId: string;
  candidateSha256: string;
  reviewer: "human" | "authorized-ai";
  reviewedAt?: string;
  reviews: Array<{
    slideNumber: number;
    sourceRasterSha256: string;
    candidateRasterSha256: string;
    verdict: QualificationReviewVerdict;
    rationale: string;
    findings?: DeckQualificationSlideReview["findings"];
  }>;
}): DeckQualificationReport {
  if (report.schema !== DECK_QUALIFICATION_SCHEMA || report.version !== 1) throw new Error("Use a current Presentation Studio deck qualification report.");
  if (input.qualificationId !== report.id || input.candidateSha256 !== report.candidate.sha256) throw new Error("The qualification or candidate changed. Read the current evidence before recording visual review.");
  if (!Array.isArray(input.reviews) || input.reviews.length < 1 || input.reviews.length > 40) throw new Error("Record between 1 and 40 slide reviews in one bounded transaction.");
  if (new Set(input.reviews.map((review) => review.slideNumber)).size !== input.reviews.length) throw new Error("Each slide may appear only once in a visual-review transaction.");
  const existing = new Map(report.visualAcceptance.reviews.map((review) => [review.slideNumber, review]));
  const reviewedAt = input.reviewedAt ?? new Date().toISOString();
  const allowedCategories = new Set<DeckQualificationSlideReview["findings"][number]["category"]>(["hierarchy", "alignment", "spacing", "layout-balance", "table-quality", "figure-clarity", "template-fidelity", "deck-consistency", "source-intent", "other"]);
  const allowedRoutes = new Set<QualificationRepairRoute>(["mcp-design", "engine-code", "image-concept", "human-review"]);
  for (const requested of input.reviews) {
    const slide = report.slides.find((candidate) => candidate.slideNumber === requested.slideNumber);
    if (!slide) throw new Error(`Slide ${requested.slideNumber} is not present in this qualification evidence bundle.`);
    if (requested.sourceRasterSha256 !== slide.sourceImage.sha256 || requested.candidateRasterSha256 !== slide.candidateImage.sha256) throw new Error(`Slide ${requested.slideNumber} pixels changed. Inspect the current source and candidate images before recording review.`);
    const rationale = String(requested.rationale ?? "").trim().slice(0, 2_000);
    if (!rationale) throw new Error(`Slide ${requested.slideNumber} requires a concrete visual-review rationale.`);
    const findings = (requested.findings ?? []).slice(0, 12).map((finding) => {
      if (!allowedCategories.has(finding.category) || !allowedRoutes.has(finding.repairRoute) || !["major", "minor"].includes(finding.severity)) throw new Error(`Slide ${requested.slideNumber} contains an invalid visual finding.`);
      const message = String(finding.message ?? "").trim().slice(0, 1_000);
      if (!message) throw new Error(`Slide ${requested.slideNumber} contains an empty visual finding.`);
      return { ...finding, message };
    });
    const previous = existing.get(slide.slideNumber);
    const pass = (previous?.pass ?? 0) + 1;
    const serious = slide.issueIds.length > 0 || findings.some((finding) => finding.severity === "major");
    let recordedVerdict: QualificationReviewVerdict = requested.verdict === "ready" && serious ? "revise" : requested.verdict;
    if (input.reviewer === "authorized-ai" && pass >= report.iteration.automaticPassLimit && recordedVerdict === "revise") recordedVerdict = "hold";
    existing.set(slide.slideNumber, {
      slideNumber: slide.slideNumber,
      sourceRasterSha256: slide.sourceImage.sha256,
      candidateRasterSha256: slide.candidateImage.sha256,
      reviewer: input.reviewer,
      requestedVerdict: requested.verdict,
      recordedVerdict,
      pass,
      rationale,
      findings,
      reviewedAt,
    });
  }
  const reviews = [...existing.values()].sort((left, right) => left.slideNumber - right.slideNumber);
  const readySlideCount = reviews.filter((review) => review.recordedVerdict === "ready").length;
  const revisionSlideCount = reviews.filter((review) => review.recordedVerdict === "revise").length;
  const heldSlideCount = reviews.filter((review) => review.recordedVerdict === "hold").length;
  const objectiveFailure = report.issues.some((issue) => issue.severity === "blocker" || issue.severity === "major");
  const visualStatus = heldSlideCount > 0 ? "held" as const
    : revisionSlideCount > 0 ? "revision-required" as const
      : readySlideCount === report.slides.length ? "review-complete" as const
        : "pending-human-or-authorized-ai-review" as const;
  const status: DeckQualificationReport["status"] = objectiveFailure ? "objective-failure"
    : visualStatus === "review-complete" ? "review-complete"
      : visualStatus === "revision-required" ? "revision-required"
        : visualStatus === "held" ? "held"
          : "visual-review-required";
  return {
    ...report,
    status,
    slides: report.slides.map((slide) => {
      const review = existing.get(slide.slideNumber);
      return {
        ...slide,
        review,
        status: slide.issueIds.length ? "objective-failure" as const : review?.recordedVerdict ?? "visual-review-required" as const,
      };
    }),
    visualAcceptance: {
      ...report.visualAcceptance,
      status: visualStatus,
      reviewedSlideCount: reviews.length,
      readySlideCount,
      revisionSlideCount,
      heldSlideCount,
      reviews,
    },
  };
}
