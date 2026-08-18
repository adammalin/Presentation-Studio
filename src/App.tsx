import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties, type DragEvent, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import {
  Archive,
  ArrowClockwise,
  ArrowCounterClockwise,
  ArrowLeft,
  ArrowRight,
  ArrowsClockwise,
  CaretRight,
  ChatCircleDots,
  Code,
  Check,
  CheckCircle,
  CirclesThreePlus,
  Crosshair,
  FileArrowDown,
  FileLock,
  FileText,
  Files,
  FolderOpen,
  Images,
  Info,
  ListChecks,
  LockKey,
  MagicWand,
  MagnifyingGlass,
  Monitor,
  PresentationChart,
  PaperPlaneTilt,
  ShieldCheck,
  Slideshow,
  Sparkle,
  SquaresFour,
  Table,
  Trash,
  UploadSimple,
  Warning,
  X,
} from "@phosphor-icons/react";
import type { Icon } from "@phosphor-icons/react";
import OnboardingTour from "./OnboardingTour";
import type {
  CleanupProposal,
  DeckJob,
  DesignThread,
  PresentationStudioProject,
  ProjectResource,
  ResourceKind,
  SceneFidelityState,
  SlideEditableObject,
  StudioConnectorDesign,
  StudioLayoutRecipe,
  StudioFigureTreatment,
  StudioQualityIssue,
  StudioTableCellDesign,
  StudioTableContinuationPlan,
  StudioTableDesign,
  StudioTableExemplarDefinition,
  StudioVisualNeed,
  StudioWebFrame,
  StudioWebNode,
  StudioWebScene,
  TemplateClassification,
} from "./types";
import { applyCleanupToPptx, buildCleanupProposalPptx, createDesignerCleanupProposal, createFontCleanupProposal, createGeometryBatchProposal, createGeometryEditProposal, createNativeLayoutProposal, createNativeLayoutRecompositionProposal, createTableLayoutProposal, createTableStyleProposal, createVisualDesignProposal } from "./lib/cleanup";
import type { LocalPresentationFont, NativeMeasurementResult, NativeRenderResult, NativeSlideRender, PickedBinaryFile } from "./lib/desktop";
import { decryptProjectPackage, encryptProjectPackage, isEncryptedProject } from "./lib/encryption";
import { auditPptx, PPTX_AUDIT_SEMANTIC_VISUAL_VERSION } from "./lib/pptx-audit";
import { createProject, touchProject } from "./lib/project";
import { markSubmittedThreadsForReanchor, removeAddressedDesignThreads, removeAddressedDesignThreadsForSlides, removeCompletedDesignThreads, removeDesignThread } from "./lib/design-threads";
import { compilePresentationScene, sceneNeedsRebuild } from "./lib/scene-graph";
import { semanticRecompositionRequests, type SemanticSlotBinding } from "./lib/recomposition";
import { compareNativeSlideRenders, type PixelComparisonMetrics } from "./lib/render-comparison";
import { buildProjectPackage, openProjectPackage } from "./lib/project-package";
import { projectPackageFromDrop } from "./lib/project-drop";
import { removeResourceFromProject, resourceRemovalImpact } from "./lib/resource-removal";
import { resourceWithAiSessionAccess, resourcesWithAiSessionAccess } from "./lib/resource-ai-access";
import {
  isPowerPointResource,
  MAX_PROJECT_RESOURCE_BYTES,
  MAX_SINGLE_RESOURCE_BYTES,
  processResourceInput,
} from "./lib/resource-ingestion";
import { buildAuditReport } from "./lib/report";
import { createOrnlDesignProfile, designStandardSummary, PRESENTATION_DESIGN_STANDARD } from "./lib/design-standard";
import { buildDeckDesignWorkOrder, buildSlideDesignWorkOrder, contentProfileForSlide } from "./lib/design-work-order";
import { slidePreviewJpeg } from "./lib/slide-preview";
import { buildSlideRenderCatalog, buildTemplateCatalog, type SlideRenderCatalog, type SlideRenderPreview, type TemplateCatalog, type TemplateLayoutPreview, type TemplatePreviewElement } from "./lib/template-catalog";
import { rankLayoutCompatibility } from "./lib/layout-semantics";
import { buildTemplatePreviewDeck } from "./lib/template-preview-deck";
import { templateLayoutPartSha256 } from "./lib/native-layout-remap";
import { isolateNativePowerPointObjects, nativeIsolationShapeIds } from "./lib/native-object-isolation";
import { bindNativeMeasurement, compareNativeMeasurementPackets, type NativeMeasurementPacket } from "./lib/native-measurement";
import { calculateDesignMetrics, metricsImproved } from "./lib/design-metrics";
import { buildInspectionPacket, type InspectionCropRegion } from "./lib/inspection-packet";
import { renderNativeContactSheet } from "./lib/contact-sheet";
import { solveAlignment, solveDistribution, solveGroupLayout, solveSafeRegion, solveSceneToLayout, type AlignmentMode, type DistributionMode, type GroupHierarchyRole, type GroupLayoutAlignment, type GroupLayoutMode, type SceneLayoutRegionRequest } from "./lib/layout-solver";
import { recommendedTableGrowthPlan, solveTableLayout } from "./lib/table-layout-solver";
import { nativeTextFrameOverflows, solveTextFit } from "./lib/text-fit-solver";
import { decideVisualIteration } from "./lib/visual-iteration";
import { buildDesignRepairLedger } from "./lib/design-repair-loop";
import { sha256 } from "./lib/hash";
import { cleanFileStem, projectSaveDefaultName } from "./lib/file-names";
import { compileStudioWebScene, planStudioExportBuild, recommendedStudioRecipe, recomposeStudioWebSlide, resizeStudioTableColumn, resizeStudioTableRow, resolvedStudioTableDesign, studioConnectorAttachmentPoint, studioGeneratedComponents, studioSlideContentSignature, translateStudioFigureTreatment, updateStudioConnectorDesign, updateStudioFigureTreatment, updateStudioTableCellDesign, updateStudioTableDesign, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "./lib/studio-web-scene";
import { applyStudioLayoutConstraints, type StudioConstraintRequest } from "./lib/studio-layout-constraints";
import { buildStudioCompositionPptx, type StudioCompositionExportResult } from "./lib/studio-composition-export";
import { validateStudioCompositionContent, type StudioCompositionContentValidation } from "./lib/studio-composition-validation";
import { nativeTextOverflows } from "./lib/fresh-composition-qa";
import { composeLatestStudioNativeRender } from "./lib/studio-design-result";
import { analyzeStudioDesignImpact } from "./lib/studio-design-impact";
import { critiqueStudioSlide } from "./lib/studio-visual-critic";
import { applyStudioDeterministicRepairPass } from "./lib/studio-repair-pass";
import { analyzeStudioDeckConsistency } from "./lib/studio-deck-consistency";
import { adoptStudioComponentStyle, compatibleStudioComponentInstances } from "./lib/studio-component-library";
import { applyStudioTableExemplar, clearStudioTableContinuation, compatibleStudioTableInstances, planStudioTableContinuation, publishStudioTableExemplar } from "./lib/studio-table-workflow";
import { attachStudioConceptReference, removeStudioConceptReference } from "./lib/studio-concept-reference";
import { reconstructStudioConcept } from "./lib/studio-concept-reconstruction";
import { createStudioVisualNeed, holdStudioVisualNeed, markStudioVisualNeedsReconstructionReady, resolveStudioVisualNeeds } from "./lib/studio-visual-needs";
import { assertSacredOrnlTitleSlideIntegrity, isProtectedOrnlTemplateSlide, unsupportedSourceSlideNumbers } from "./lib/template-guardrails";
import { preserveNativeSlide } from "./lib/native-slide-preservation";
import { buildDeckQualificationReport, qualificationEvidenceSlideNumber, recordDeckQualificationReviews, type DeckQualificationReport } from "./lib/deck-qualification";
import {
  ONBOARDING_TOUR_STORAGE_KEY,
  ONBOARDING_TOUR_VERSION,
  shouldShowOnboardingTour,
} from "./lib/onboarding";
import { isProposalSlideWorkspaceRequest, type SlideWorkspaceRequest } from "./lib/slide-workspace";
import { assertExactResourceExcerpt, resourceTextPage } from "./lib/resource-text-access";
import { bindNewStudioSceneToGeneratedPowerPoint, createNewStudioPresentationScene, type NewStudioPresentationInput } from "./lib/new-studio-presentation";

type ViewId = "batch" | "decks" | "slides" | "studio" | "designs" | "rules" | "review" | "resources";
type McpActivityPhase = "working" | "inspecting" | "found-issues" | "fixing" | "rechecking" | "ready" | "attention";
type McpActivityState = { id: string; operation: string; state: "active" | "completed" | "failed"; phase: McpActivityPhase; issueCount?: number; autoFixableCount?: number };

function mcpPhaseForOperation(operation: string, input: Record<string, unknown>): McpActivityPhase {
  if (operation === "run_deck_qualification" || operation === "get_deck_qualification" || operation === "get_qualification_slide" || operation === "get_qualification_contact_sheet") return "inspecting";
  if (operation === "record_deck_qualification_review") return "rechecking";
  if (operation === "get_slide_inspection_packet") return input.representation === "proposal" ? "rechecking" : "inspecting";
  if (operation === "get_studio_slide_critique") return "inspecting";
  if (operation === "preview_studio_fresh_composition" || operation === "record_proposal_visual_critique" || operation === "record_studio_visual_critique" || operation === "get_slide_render_comparison") return "rechecking";
  if (operation.startsWith("stage_") || operation.startsWith("solve_") || operation === "create_studio_presentation" || operation === "fit_scene_to_layout" || operation === "refine_studio_layout" || operation === "repair_studio_objective_issues" || operation === "reconstruct_studio_concept" || operation === "publish_studio_component_style" || operation === "publish_studio_table_exemplar" || operation === "plan_studio_table_continuation") return "fixing";
  return "working";
}

function mcpActivityCopy(activity: McpActivityState) {
  if (activity.state === "failed") return { title: "AI operation needs attention", detail: `${activity.operation.replaceAll("_", " ")} · stopped` };
  if (activity.phase === "inspecting") return { title: "Inspecting original slide…", detail: "Comparing native pixels, structure, fit, and technical objects" };
  if (activity.phase === "found-issues") return { title: `Found ${activity.issueCount ?? 0} issue${activity.issueCount === 1 ? "" : "s"}`, detail: `${activity.autoFixableCount ?? 0} bounded fix${activity.autoFixableCount === 1 ? "" : "es"} available · original intent locked` };
  if (activity.phase === "fixing") return { title: activity.state === "active" ? "Fixing…" : "Fix staged for review", detail: "Applying one coherent semantic treatment; source remains unchanged" };
  if (activity.phase === "rechecking") return { title: activity.state === "active" ? "Rechecking original intent…" : "Original-intent recheck ready", detail: "Comparing exact wording, visuals, values, and relationships" };
  if (activity.phase === "ready") return { title: "Ready for review", detail: "The proposal remains unapplied until human acceptance" };
  return { title: activity.state === "active" ? "AI is using Presentation Studio" : "AI operation completed", detail: `${activity.operation.replaceAll("_", " ")} · ${activity.state}` };
}
type StudioFreshPreview = Omit<StudioCompositionExportResult, "bytes"> & { bytes: Uint8Array; deckId: string; sourceSlideNumber: number; sceneRevision: string; slideUpdatedAt: string; contentValidation: StudioCompositionContentValidation; nativeRender?: NativeRenderResult; nativeMeasurement?: NativeMeasurementResult };
type StudioDeckBuild = Omit<StudioCompositionExportResult, "bytes"> & { bytes: Uint8Array; deckId: string; sceneRevision: string; slideUpdatedAts: Record<number, string>; contentValidation: StudioCompositionContentValidation; nativeRender: NativeRenderResult; nativeMeasurement: NativeMeasurementResult; candidateAudit: NonNullable<DeckJob["audit"]> };
type StudioDeckQualification = { report: DeckQualificationReport; outputRoot: string; reportPath: string; htmlPath: string; sceneRevision: string; candidateSha256: string };
const MAX_PROJECT_PACKAGE_BYTES = 1_500_000_000;

function requestedAddressedThreadIds(input: { addressedThreadIds?: unknown }): string[] {
  return Array.isArray(input.addressedThreadIds) ? [...new Set(input.addressedThreadIds.map(String))] : [];
}

const navItems: Array<{ id: ViewId; label: string; icon: Icon }> = [
  { id: "batch", label: "Batch", icon: Files },
  { id: "decks", label: "Deck audit", icon: PresentationChart },
  { id: "slides", label: "Slides", icon: Slideshow },
  { id: "studio", label: "Studio", icon: Code },
  { id: "designs", label: "Designs", icon: SquaresFour },
  { id: "rules", label: "Rules", icon: ListChecks },
  { id: "review", label: "Review", icon: MagicWand },
  { id: "resources", label: "Resources", icon: Archive },
];

const classificationLabels: Record<TemplateClassification, string> = {
  "current-ornl": "Current ORNL",
  "older-or-modified-ornl": "Older / modified ORNL",
  sponsor: "Sponsor",
  custom: "Custom",
  mixed: "Mixed",
  unknown: "Unknown",
};

const resourceKindLabels: Record<ResourceKind, string> = {
  presentation: "Presentation",
  document: "Document",
  data: "Data",
  image: "Image",
  audio: "Audio",
  video: "Video",
  other: "Other",
};

function bytesFrom(value: Uint8Array): Uint8Array {
  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

function bytesToBase64(value: Uint8Array): string {
  const bytes = bytesFrom(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) binary += String.fromCharCode(...bytes.subarray(offset, Math.min(bytes.length, offset + chunkSize)));
  return btoa(binary);
}

async function boundedResourceImagePreview(resource: ProjectResource, maximumDimension = 1_600) {
  if (resource.kind !== "image" || !resource.mediaType.startsWith("image/") || !resource.bytes?.byteLength) throw new Error("This Resource is not an embedded previewable image.");
  const bitmap = await createImageBitmap(new Blob([bytesFrom(resource.bytes).slice().buffer], { type: resource.mediaType }));
  const scale = Math.min(1, maximumDimension / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  canvas.getContext("2d")!.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error("Presentation Studio could not encode the Resource preview.")), "image/png"));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  return { bytes, mimeType: "image/png" as const, width: canvas.width, height: canvas.height, sha256: await sha256(bytes) };
}

async function studioResourceMedia(scene: StudioWebScene, resources: ProjectResource[]): Promise<Record<string, string>> {
  const ids = new Set(scene.slides.flatMap((slide) => slide.nodes.flatMap((node) => node.mediaPart?.startsWith("resource:") ? [node.mediaPart.slice("resource:".length)] : [])));
  const media: Record<string, string> = {};
  for (const id of ids) {
    const resource = resources.find((item) => item.id === id);
    if (!resource) throw new Error(`A slide references missing image Resource ${id}.`);
    const preview = await boundedResourceImagePreview(resource, 2_400);
    media[`resource:${id}`] = `data:${preview.mimeType};base64,${bytesToBase64(preview.bytes)}`;
  }
  return media;
}

function catalogWithStudioResources(catalog: SlideRenderCatalog | undefined, media: Record<string, string>, scene: StudioWebScene): SlideRenderCatalog {
  return catalog ? { ...catalog, media: { ...catalog.media, ...media } } : {
    id: `${scene.deckId}:resource-media`,
    name: "Resource-authored Studio presentation",
    sha256: scene.sourceSha256,
    slideWidth: scene.slideSize.width,
    slideHeight: scene.slideSize.height,
    slides: [],
    media,
    generatedAt: new Date().toISOString(),
    renderer: "local-ooxml-preview",
  };
}

async function inspectionRasterEvidence(slide: NativeSlideRender, regions: InspectionCropRegion[]) {
  const sourceBlob = new Blob([bytesFrom(slide.bytes).slice().buffer], { type: slide.mimeType });
  const bitmap = await createImageBitmap(sourceBlob);
  const images: Array<{ id: string; kind: string; mimeType: "image/png"; data: string; width: number; height: number; sha256: string; region?: InspectionCropRegion["normalized"]; reason: string }> = [];
  const addCanvas = async (id: string, kind: string, canvas: HTMLCanvasElement, reason: string, region?: InspectionCropRegion["normalized"]) => {
    const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error("Presentation Studio could not encode inspection evidence.")), "image/png"));
    const bytes = new Uint8Array(await blob.arrayBuffer());
    images.push({ id, kind, mimeType: "image/png", data: bytesToBase64(bytes), width: canvas.width, height: canvas.height, sha256: await sha256(bytes), region, reason });
  };
  const full = document.createElement("canvas");
  full.width = bitmap.width;
  full.height = bitmap.height;
  full.getContext("2d")!.drawImage(bitmap, 0, 0);
  await addCanvas(`slide-${slide.number}-full`, "full-slide", full, "Authoritative full-slide PowerPoint render for gestalt and hierarchy review.");
  for (const region of regions) {
    const sx = Math.max(0, Math.floor(region.normalized.x * bitmap.width));
    const sy = Math.max(0, Math.floor(region.normalized.y * bitmap.height));
    const sw = Math.max(1, Math.min(bitmap.width - sx, Math.ceil(region.normalized.width * bitmap.width)));
    const sh = Math.max(1, Math.min(bitmap.height - sy, Math.ceil(region.normalized.height * bitmap.height)));
    const canvas = document.createElement("canvas");
    canvas.width = sw;
    canvas.height = sh;
    canvas.getContext("2d")!.drawImage(bitmap, sx, sy, sw, sh, 0, 0, sw, sh);
    await addCanvas(region.id, region.kind, canvas, region.reason, region.normalized);
  }
  const overlay = document.createElement("canvas");
  overlay.width = bitmap.width;
  overlay.height = bitmap.height;
  const context = overlay.getContext("2d")!;
  context.drawImage(bitmap, 0, 0);
  context.lineWidth = Math.max(3, bitmap.width / 500);
  context.font = `600 ${Math.max(18, Math.round(bitmap.width / 60))}px Aptos, Arial, sans-serif`;
  regions.forEach((region, index) => {
    const x = region.normalized.x * bitmap.width;
    const y = region.normalized.y * bitmap.height;
    const width = region.normalized.width * bitmap.width;
    const height = region.normalized.height * bitmap.height;
    context.strokeStyle = region.kind === "table" ? "#f2a900" : "#007a33";
    context.fillStyle = "rgba(255,255,255,.92)";
    context.strokeRect(x, y, width, height);
    const label = `${index + 1} · ${region.kind}`;
    const labelWidth = context.measureText(label).width + 18;
    context.fillRect(x, Math.max(0, y - 30), labelWidth, 30);
    context.fillStyle = "#111820";
    context.fillText(label, x + 9, Math.max(21, y - 8));
  });
  await addCanvas(`slide-${slide.number}-diagnostic-overlay`, "diagnostic-overlay", overlay, "Deterministic crop map; colored rectangles are diagnostics and are not presentation artwork.");
  bitmap.close();
  return images;
}

function qualificationInspectionRegions(report: DeckQualificationReport, slideNumber: number, issueId?: string): InspectionCropRegion[] {
  return report.issues.filter((issue) => issue.slideNumber === slideNumber && issue.evidenceRegion && (!issueId || issue.id === issueId)).map((issue) => ({
    id: issue.id,
    kind: "issue" as const,
    objectIds: issue.evidenceRegion!.objectIds,
    normalized: { x: issue.evidenceRegion!.x, y: issue.evidenceRegion!.y, width: issue.evidenceRegion!.width, height: issue.evidenceRegion!.height },
    reason: `${issue.message} ${issue.evidence}`,
  }));
}

function previewFontStack(fontFamily?: string): string {
  const requested = (fontFamily ?? "Aptos").replaceAll('"', "").trim() || "Aptos";
  return /^aptos(?:\s|$)/i.test(requested) ? `"${requested}", Arial, sans-serif` : `"${requested}", Aptos, Arial, sans-serif`;
}

function fontFaceRule(font: LocalPresentationFont): string {
  return `@font-face{font-family:"${font.family}";src:url("data:${font.mediaType};base64,${bytesToBase64(font.bytes)}") format("truetype");font-weight:${font.weight};font-style:${font.style};font-display:block;}`;
}

function formatBytes(value: number): string {
  if (value < 1024) return `${value} B`;
  if (value < 1024 ** 2) return `${(value / 1024).toFixed(1)} KB`;
  return `${(value / 1024 ** 2).toFixed(1)} MB`;
}

function statusLabel(status: DeckJob["status"]): string {
  return status.split("-").map((word) => word[0].toUpperCase() + word.slice(1)).join(" ");
}

function sourceForDeck(project: PresentationStudioProject, deck: DeckJob): ProjectResource | undefined {
  return project.resources.find((resource) => resource.id === deck.sourceResourceId);
}

function withCompiledScene(deck: DeckJob): DeckJob {
  return deck.audit ? { ...deck, scene: compilePresentationScene({ ...deck, audit: deck.audit }) } : deck;
}

const fidelityLabels: Record<SceneFidelityState, string> = {
  "editable-native": "Editable native",
  "preserved-native": "Preserved native",
  "conversion-required": "Conversion required",
  "unsupported-blocking": "Blocking",
};

function StatusPill({ status }: { status: DeckJob["status"] }) {
  const intent = status === "approved" || status === "exported" ? "success" : status === "failed" ? "danger" : status.includes("needs") ? "warning" : "neutral";
  return <span className={`status-pill ${intent}`}><span className="status-dot" />{statusLabel(status)}</span>;
}

function Metric({ value, label }: { value: string | number; label: string }) {
  return <div className="metric"><strong>{value}</strong><span>{label}</span></div>;
}

function EmptyWorkspace({ onAdd }: { onAdd: () => void }) {
  return (
    <section className="empty-workspace">
      <div className="empty-icon"><PresentationChart size={34} weight="light" /></div>
      <p className="eyebrow">Source-faithful design</p>
      <h1>Turn a presentation into one coherent design system.</h1>
      <p className="empty-copy">Add PowerPoint decks to preserve their source content, extract reusable evidence and media, and rebuild each slide in the ORNL design language. Studio keeps the original untouched while you and the AI review the finished, exportable result.</p>
      <button className="button primary large" onClick={onAdd}><CirclesThreePlus size={19} />Add PowerPoint decks</button>
      <div className="promise-row">
        <span><ShieldCheck size={17} />Originals stay untouched</span>
        <span><CheckCircle size={17} />Visible text is hash-checked</span>
        <span><FileLock size={17} />Optional project encryption</span>
      </div>
    </section>
  );
}

function BatchView({ project, selectedId, onSelect, onAdd }: { project: PresentationStudioProject; selectedId?: string; onSelect: (id: string) => void; onAdd: () => void }) {
  if (project.decks.length === 0) return <EmptyWorkspace onAdd={onAdd} />;
  const slides = project.decks.reduce((sum, deck) => sum + (deck.audit?.slideCount ?? 0), 0);
  const decisions = project.decks.filter((deck) => deck.status === "needs-template-decision").length;
  const findings = project.decks.reduce((sum, deck) => sum + (deck.audit?.findings.length ?? 0), 0);
  return (
    <div className="view-stack">
      <header className="view-header">
        <div><p className="eyebrow">Review batch</p><h1>{project.project.name}</h1><p>Independent decks, one conservative review queue.</p></div>
        <button className="button primary" onClick={onAdd}><CirclesThreePlus size={18} />Add decks</button>
      </header>
      <div className="metric-strip">
        <Metric value={project.decks.length} label="Decks" />
        <Metric value={slides} label="Slides scanned" />
        <Metric value={findings} label="Audit findings" />
        <Metric value={decisions} label="Template decisions" />
      </div>
      <section className="panel table-panel">
        <div className="panel-heading"><div><h2>Deck queue</h2><p>Cleanup and export remain isolated per source.</p></div><span className="quiet-label">{project.decks.length} total</span></div>
        <div className="deck-table" role="table" aria-label="Imported presentation decks">
          <div className="deck-row deck-head" role="row"><span>Deck</span><span>Template</span><span>Slides</span><span>Fonts</span><span>Tables</span><span>Status</span><span /></div>
          {project.decks.map((deck) => (
            <button key={deck.id} className={`deck-row ${selectedId === deck.id ? "selected" : ""}`} role="row" onClick={() => onSelect(deck.id)}>
              <span className="deck-name"><PresentationChart size={19} /><span><strong>{deck.name}</strong><small>{formatBytes(sourceForDeck(project, deck)?.byteLength ?? 0)}</small></span></span>
              <span>{classificationLabels[deck.templateClassification]}</span>
              <span>{deck.audit?.slideCount ?? "—"}</span>
              <span>{deck.audit?.fonts.filter((font) => font.directSlideCount > 0).length ?? "—"}</span>
              <span>{deck.audit?.tableCount ?? "—"}</span>
              <span><StatusPill status={deck.status} /></span>
              <span><CaretRight size={16} /></span>
            </button>
          ))}
        </div>
      </section>
    </div>
  );
}

function DeckAuditView({ deck, onConfirm, onStage, onStartOrnlCleanup, onMarkExemplar, onExportReport, isExemplar }: { deck?: DeckJob; onConfirm: (templateId: string) => void; onStage: () => void; onStartOrnlCleanup: () => void; onMarkExemplar: () => void; onExportReport: () => void; isExemplar: boolean }) {
  if (!deck?.audit) return <NoSelection message="Select an audited deck from the Batch workspace." />;
  const audit = deck.audit;
  return (
    <div className="view-stack">
      <header className="view-header compact"><div><p className="eyebrow">Read-only deck audit</p><h1>{deck.name}</h1><p>Scanned directly from PowerPoint package structure; no source bytes were changed.</p></div><div className="header-actions"><button className="button ghost small" onClick={onExportReport}><FileArrowDown size={16} />Audit report</button><button className="button primary small" onClick={onStartOrnlCleanup}><MagicWand size={16} />Clean up with ORNL defaults</button><StatusPill status={deck.status} /></div></header>
      <section className="classification-card">
        <div className="classification-icon"><ShieldCheck size={26} /></div>
        <div><span className="field-label">Detected template family</span><h2>{classificationLabels[audit.classification]}</h2><p>{audit.classificationEvidence.join(" ")}</p></div>
        <div className="template-actions">
          <label htmlFor="target-template">Human-confirmed target</label>
          <select id="target-template" value={deck.targetTemplateId ?? ""} onChange={(event) => event.target.value && onConfirm(event.target.value)}>
            <option value="">Choose target…</option>
            <option value="ornl-16x9-v1">ORNL 16:9 · Aptos</option>
            <option value="sponsor-source">Preserve sponsor template</option>
            <option value="custom-source">Preserve custom/source template</option>
          </select>
          {deck.targetTemplateConfirmedAt && <span className="confirmation"><Check size={14} />{deck.targetTemplateDecisionSource === "automatic-default" ? "ORNL default adopted" : deck.targetTemplateDecisionSource === "automatic-source-preservation" ? "Source template preserved" : "Selected by user"}</span>}
        </div>
      </section>
      {deck.designProfile && <section className="adopted-profile-bar"><ShieldCheck size={20} /><div><span className="field-label">Resolved presentation standard</span><strong>{designStandardSummary()}</strong><small>Version {deck.designProfile.standardVersion} · routine design choices proceed without per-slide approvals · draft pending ORNL review</small></div><span>Active</span></section>}
      {deck.scene && <section className="scene-fidelity-card"><div><span className="field-label">Hybrid PowerPoint scene</span><strong>{deck.scene.objects.length} source-bound objects</strong><small>Native PowerPoint pixels remain authoritative; Studio edits only capabilities declared by this scene revision.</small></div><div className="scene-fidelity-metrics">{(Object.keys(fidelityLabels) as SceneFidelityState[]).map((state) => <span key={state} className={`fidelity-pill ${state}`}><strong>{deck.scene?.fidelityCounts[state] ?? 0}</strong>{fidelityLabels[state]}</span>)}</div>{deck.scene.preservationEnvelope.blockingFeatures.length > 0 && <div className="scene-blocker"><Warning size={16} />Manual review required for {deck.scene.preservationEnvelope.blockingFeatures.join(", ")}.</div>}</section>}
      <div className="metric-strip six">
        <Metric value={audit.slideCount} label="Slides" /><Metric value={audit.layoutCount} label="Layouts" /><Metric value={audit.masterCount} label="Masters" /><Metric value={audit.tableCount} label="Tables" /><Metric value={audit.pictureCount} label="Pictures" /><Metric value={audit.notesCount} label="Notes" />
      </div>
      <div className="split-grid">
        <section className="panel">
          <div className="panel-heading"><div><h2>Font inventory</h2><p>Direct slide use is separated from layout, master, and theme fallbacks.</p></div><button className="button secondary small" disabled={!deck.targetTemplateConfirmedAt || deck.targetTemplateId !== "ornl-16x9-v1"} onClick={onStage}><MagicWand size={16} />Stage designer cleanup</button></div>
          <div className="font-list">
            {audit.fonts.length === 0 && <p className="muted">No explicit typeface references were found.</p>}
            {audit.fonts.map((font) => (
              <div className="font-row" key={font.normalizedFamily}>
                <span className="font-sample">Aa</span><span><strong>{font.family}</strong><small>{font.directSlideCount > 0 ? `${font.directSlideCount} direct slide references` : "Template/theme only"}{font.isLikelySymbolFont ? " · protected symbol font" : ""}</small></span><span className="font-count">{font.count}</span>
              </div>
            ))}
          </div>
        </section>
        <section className="panel">
          <div className="panel-heading"><div><h2>Production support</h2><p>Features that can affect cleanup safety.</p></div><span className={`support-badge ${audit.supportLevel}`}>{audit.supportLevel}</span></div>
          <dl className="check-list">
            <div><dt>Modern comments</dt><dd>{audit.modernCommentCount || "None"}</dd></div>
            <div><dt>External relationships</dt><dd className={audit.containsExternalRelationships ? "bad" : "good"}>{audit.containsExternalRelationships ? "Review" : "None"}</dd></div>
            <div><dt>Macros</dt><dd className={audit.containsMacros ? "bad" : "good"}>{audit.containsMacros ? "Review" : "None"}</dd></div>
            <div><dt>Embedded OLE</dt><dd className={audit.containsOleObjects ? "bad" : "good"}>{audit.containsOleObjects ? "Review" : "None"}</dd></div>
            <div><dt>Expanded package</dt><dd>{formatBytes(audit.expandedByteLength)}</dd></div>
          </dl>
          {audit.warnings.map((warning) => <div className="inline-note" key={warning}><Info size={17} />{warning}</div>)}
        </section>
      </div>
      {audit.tableCount > 0 && <section className="exemplar-card"><div className="exemplar-symbol"><Table size={22} /></div><div><span className="field-label">Approved style exemplar</span><h2>{audit.tableCount === 1 ? "Use this deck's native table as the batch reference" : "Choose one table before registering an exemplar"}</h2><p>The record identifies the exact embedded source, slide, and table ordinal. It never copies sample cell content into another deck.</p></div><button className="button secondary" disabled={audit.tableCount !== 1 || isExemplar} onClick={onMarkExemplar}>{isExemplar ? <><Check size={17} />Registered</> : <><Table size={17} />Register table style</>}</button></section>}
      {audit.tables.length > 0 && <section className="panel"><div className="panel-heading"><div><h2>Native table inventory</h2><p>Structure and style fingerprints; sample cell text is excluded.</p></div><span className="quiet-label">{new Set(audit.tables.map((table) => table.styleId ?? "none")).size} style IDs</span></div><div className="table-inventory"><div className="table-inventory-row table-inventory-head"><span>Location</span><span>Structure</span><span>PowerPoint style</span><span>Direct fonts</span><span>Fingerprint</span></div>{audit.tables.map((table) => <div className="table-inventory-row" key={table.id}><span>Slide {table.slideNumber} · Table {table.ordinal}</span><span>{table.rowCount} × {table.columnCount}{table.mergedCellCount ? ` · ${table.mergedCellCount} merges` : ""}</span><span title={table.styleId}>{table.styleId ? `${table.styleId.slice(0, 8)}…` : "None"}</span><span>{table.cellFonts.join(", ") || "Inherited"}</span><span>{table.styleFingerprint.slice(0, 10)}…</span></div>)}</div></section>}
      {audit.pictures.length > 0 && <section className="panel"><div className="panel-heading"><div><h2>Native figure inventory</h2><p>Picture treatment and stored description status.</p></div><span className="quiet-label">{audit.pictures.filter((picture) => !picture.description).length} descriptions missing</span></div><div className="figure-inventory">{audit.pictures.map((picture) => <div className="figure-row" key={picture.id}><span className="figure-icon"><Images size={16} /></span><span><strong>Slide {picture.slideNumber} · {picture.name}</strong><small>{picture.widthEmu && picture.heightEmu ? `${(picture.widthEmu / 914400).toFixed(1)} × ${(picture.heightEmu / 914400).toFixed(1)} in` : "Size unavailable"}</small></span><span>{picture.cropped ? "Cropped" : "Uncropped"}</span><span>{picture.hasOutline ? "Outline" : "No outline"}</span><span className={picture.description ? "good-text" : "review-text"}>{picture.description ? "Description stored" : "Review description"}</span></div>)}</div></section>}
      <section className="panel">
        <div className="panel-heading"><div><h2>Findings</h2><p>Evidence first; ambiguous differences remain review items.</p></div><span className="quiet-label">{audit.findings.length} findings</span></div>
        <div className="finding-list">
          {audit.findings.map((item, index) => <div className="finding-row" key={`${item.id}-${item.slideNumber ?? "deck"}-${index}`}><span className={`finding-icon ${item.severity}`}>{item.severity === "error" || item.severity === "warning" ? <Warning size={17} /> : <Info size={17} />}</span><div><span className="finding-category">{item.category}</span><strong>{item.message}</strong><p>{item.evidence}</p></div><span className="confidence">{item.confidence}</span></div>)}
        </div>
      </section>
    </div>
  );
}

function SlidesView({ deck, catalog, nativeRender, outputSlides, loading, revision, threads, openRequest, deckBuildReady, resourceCount, onAddDeck, onOpenResources, onSaveThread, onDeleteThread, onStageGeometry, onOpenStudioSlide, onSaveDeck }: { deck?: DeckJob; catalog?: SlideRenderCatalog; nativeRender?: NativeRenderResult; outputSlides?: StudioCompositionExportResult["outputSlides"]; loading: boolean; revision: string; threads: DesignThread[]; openRequest?: SlideWorkspaceRequest; deckBuildReady: boolean; resourceCount: number; onAddDeck: () => void; onOpenResources: () => void; onSaveThread: (slide: SlideRenderPreview, anchor: DesignThread["anchor"], comment: string, submit: boolean, outputSlideNumber?: number) => void; onDeleteThread: (threadId: string) => void; onStageGeometry: (object: SlideEditableObject, target: { x: number; y: number; width: number; height: number }, rationale: string) => void; onOpenStudioSlide: (slideNumber: number) => void; onSaveDeck: () => void }) {
  const [selectedNumber, setSelectedNumber] = useState<number>();
  const [mode, setMode] = useState<"review" | "edit" | "comment">("review");
  const [draftAnchor, setDraftAnchor] = useState<DesignThread["anchor"]>();
  const [draftComment, setDraftComment] = useState("");
  const [selectedObjectId, setSelectedObjectId] = useState<string>();
  const [draftGeometry, setDraftGeometry] = useState<{ x: number; y: number; width: number; height: number }>();
  const [editRationale, setEditRationale] = useState("");
  const dragStart = useRef<{ x: number; y: number } | undefined>(undefined);
  const objectDrag = useRef<{ point: { x: number; y: number }; geometry: { x: number; y: number; width: number; height: number } } | undefined>(undefined);
  const editorCanvas = useRef<HTMLDivElement>(null);
  const commentInput = useRef<HTMLTextAreaElement>(null);
  useEffect(() => { setSelectedNumber(undefined); setMode("review"); setDraftAnchor(undefined); setDraftComment(""); setSelectedObjectId(undefined); setDraftGeometry(undefined); setEditRationale(""); }, [deck?.id]);
  useEffect(() => {
    if (!deck || !openRequest || openRequest.deckId !== deck.id) return;
    setSelectedNumber(outputSlides?.find((slide) => slide.sourceSlideNumber === openRequest.slideNumber)?.outputSlideNumber ?? openRequest.slideNumber);
    setMode(openRequest.mode);
    setDraftAnchor(undefined);
    setDraftComment("");
  }, [deck?.id, openRequest?.id, outputSlides]);
  useEffect(() => {
    if (mode !== "comment") return;
    const frame = window.requestAnimationFrame(() => commentInput.current?.focus());
    return () => window.cancelAnimationFrame(frame);
  }, [mode, selectedNumber, threads.length]);
  useEffect(() => { setSelectedObjectId(undefined); setDraftGeometry(undefined); setEditRationale(""); objectDrag.current = undefined; }, [selectedNumber]);
  if (!deck) return <section className="slides-empty-state">
    <span className="slides-empty-icon"><Slideshow size={34} weight="light" /></span>
    <p className="eyebrow">Slides workspace</p>
    <h1>No presentation has been created yet</h1>
    <p>{resourceCount > 0 ? `This project contains ${resourceCount} Resource${resourceCount === 1 ? "" : "s"}. Turn on AI access once and Studio automatically shares every compatible source, then ask the Presentation Studio MCP to create a source-grounded native presentation—or add a PowerPoint to redesign it.` : "Add source materials or a PowerPoint. Presentation Studio can create a new native JSON presentation from shared Resources or redesign an imported deck in the same central scene."}</p>
    <div><button className="button primary" type="button" onClick={onAddDeck}><PresentationChart size={17} />Add a PowerPoint</button>{resourceCount > 0 && <button className="button secondary" type="button" onClick={onOpenResources}><Archive size={17} />Review Resources</button>}</div>
    <small>New presentations use the installed ORNL Template Pack, Aptos, Resource-hash provenance, editable Studio objects, and PowerPoint-native QA.</small>
  </section>;
  if (!deck.audit) return <NoSelection message="This presentation could not be audited. Review its failure details in Deck audit before opening slide designs." />;

  const proposalWorkspace = isProposalSlideWorkspaceRequest(openRequest, deck.id);
  const centralDesignActive = Boolean(deck.studioScene && !proposalWorkspace);
  const centralOutputSlides: StudioCompositionExportResult["outputSlides"] = centralDesignActive && outputSlides?.length ? outputSlides : deck.audit.slides.map((slide) => ({ outputSlideNumber: slide.number, sourceSlideNumber: slide.number }));
  const selectedOutput = centralDesignActive ? centralOutputSlides.find((slide) => slide.outputSlideNumber === selectedNumber) : undefined;
  const selectedSourceNumber = selectedOutput?.sourceSlideNumber ?? selectedNumber;
  const selected = catalog?.slides.find((slide) => slide.number === selectedSourceNumber);
  const selectedStudioSlide = deck.studioScene?.slides.find((slide) => slide.slideNumber === selectedSourceNumber);
  const selectedSacredTitle = Boolean(selectedSourceNumber && isProtectedOrnlTemplateSlide(deck, selectedSourceNumber));
  const selectedConverted = Boolean(selectedStudioSlide?.status === "designed" && selectedStudioSlide.recipe !== "source");
  const firstOutputForSource = centralOutputSlides.find((slide) => slide.sourceSlideNumber === selectedSourceNumber)?.outputSlideNumber;
  const selectedThreads = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === selectedSourceNumber && (!centralDesignActive || thread.outputSlideNumber === selectedNumber || thread.outputSlideNumber === undefined && selectedNumber === firstOutputForSource));
  const acceptedGeometry = new Map((proposalWorkspace || deck.proposal?.status === "applied") && deck.proposal ? deck.proposal.changes.filter((change) => change.selected && change.kind === "geometry").flatMap((change) => change.geometryCommands ?? []).map((command) => [command.objectId, command.target]) : []);
  const slideObjects = (deck.audit.editableObjects ?? []).filter((object) => object.slideNumber === selectedSourceNumber).map((object) => acceptedGeometry.has(object.id) ? { ...object, geometry: { ...object.geometry, ...acceptedGeometry.get(object.id)! } } : object);
  const slideScene = deck.scene?.slides.find((slide) => slide.number === selectedSourceNumber);
  const sceneObjects = (deck.scene?.objects ?? []).filter((object) => object.slideNumber === selectedSourceNumber);
  const sceneObjectById = new Map(sceneObjects.map((object) => [object.id, object]));
  const selectedObject = slideObjects.find((object) => object.id === selectedObjectId);
  const slideSize = deck.audit.slideSize ?? { width: catalog?.slideWidth ?? 12_192_000, height: catalog?.slideHeight ?? 6_858_000 };
  const commentMode = mode === "comment";
  const editMode = mode === "edit";
  const nativeReady = nativeRender?.status === "ready" && nativeRender.authoritative;
  const selectedNativeReady = nativeReady && Boolean(nativeRender.slides.find((slide) => slide.number === selectedNumber));
  const resultCount = centralDesignActive ? nativeRender?.slides.length ?? 0 : nativeRender?.slideCount ?? nativeRender?.slides.length ?? 0;
  function baseGeometry(object: SlideEditableObject) { const { x, y, width, height } = object.geometry; return { x, y, width, height }; }
  function activeGeometry(object: SlideEditableObject) { return object.id === selectedObjectId && draftGeometry ? draftGeometry : baseGeometry(object); }
  function selectObject(object: SlideEditableObject) { setSelectedObjectId(object.id); setDraftGeometry(baseGeometry(object)); setEditRationale(""); }
  function clampGeometry(geometry: { x: number; y: number; width: number; height: number }) {
    const width = Math.max(91_440, Math.min(slideSize.width, geometry.width));
    const height = Math.max(91_440, Math.min(slideSize.height, geometry.height));
    return { width, height, x: Math.max(0, Math.min(slideSize.width - width, geometry.x)), y: Math.max(0, Math.min(slideSize.height - height, geometry.y)) };
  }
  function updateDraft(update: (current: { x: number; y: number; width: number; height: number }) => { x: number; y: number; width: number; height: number }) {
    if (!selectedObject) return;
    setDraftGeometry((current) => clampGeometry(update(current ?? baseGeometry(selectedObject))));
  }
  function beginObjectDrag(event: ReactPointerEvent<HTMLButtonElement>, object: SlideEditableObject) {
    event.stopPropagation();
    selectObject(object);
    if (!object.canMove || !editorCanvas.current) return;
    const bounds = editorCanvas.current.getBoundingClientRect();
    event.currentTarget.setPointerCapture(event.pointerId);
    objectDrag.current = { point: { x: (event.clientX - bounds.left) / bounds.width, y: (event.clientY - bounds.top) / bounds.height }, geometry: baseGeometry(object) };
  }
  function moveObject(event: ReactPointerEvent<HTMLDivElement>) {
    if (!editMode || !objectDrag.current || !editorCanvas.current || event.buttons === 0) return;
    const bounds = editorCanvas.current.getBoundingClientRect();
    const point = { x: (event.clientX - bounds.left) / bounds.width, y: (event.clientY - bounds.top) / bounds.height };
    const snap = 45_720;
    const next = clampGeometry({ ...objectDrag.current.geometry, x: Math.round((objectDrag.current.geometry.x + (point.x - objectDrag.current.point.x) * slideSize.width) / snap) * snap, y: Math.round((objectDrag.current.geometry.y + (point.y - objectDrag.current.point.y) * slideSize.height) / snap) * snap });
    setDraftGeometry(next);
  }
  function finishObjectDrag() { objectDrag.current = undefined; }
  function setInches(field: "x" | "y" | "width" | "height", value: string) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return;
    updateDraft((current) => {
      const nextValue = Math.round(numeric * 914_400);
      if (selectedObject?.kind === "picture" && field === "width") return { ...current, width: nextValue, height: Math.round(nextValue * selectedObject.geometry.height / selectedObject.geometry.width) };
      if (selectedObject?.kind === "picture" && field === "height") return { ...current, height: nextValue, width: Math.round(nextValue * selectedObject.geometry.width / selectedObject.geometry.height) };
      return { ...current, [field]: nextValue };
    });
  }
  function normalizedPoint(event: ReactPointerEvent<HTMLDivElement>) {
    const bounds = event.currentTarget.getBoundingClientRect();
    return { x: Math.min(1, Math.max(0, (event.clientX - bounds.left) / bounds.width)), y: Math.min(1, Math.max(0, (event.clientY - bounds.top) / bounds.height)) };
  }
  function beginAnchor(event: ReactPointerEvent<HTMLDivElement>) {
    if (!commentMode) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    const point = normalizedPoint(event);
    dragStart.current = point;
    setDraftAnchor({ kind: "region", x: point.x, y: point.y, width: .001, height: .001 });
  }
  function moveAnchor(event: ReactPointerEvent<HTMLDivElement>) {
    if (!commentMode || !dragStart.current || event.buttons === 0) return;
    const point = normalizedPoint(event);
    const start = dragStart.current;
    setDraftAnchor({ kind: "region", x: Math.min(start.x, point.x), y: Math.min(start.y, point.y), width: Math.max(.001, Math.abs(point.x - start.x)), height: Math.max(.001, Math.abs(point.y - start.y)) });
  }
  function finishAnchor(event: ReactPointerEvent<HTMLDivElement>) {
    if (!commentMode || !dragStart.current) return;
    const start = dragStart.current;
    const point = normalizedPoint(event);
    dragStart.current = undefined;
    let x = Math.min(start.x, point.x);
    let y = Math.min(start.y, point.y);
    let width = Math.abs(point.x - start.x);
    let height = Math.abs(point.y - start.y);
    if (width < .02 && height < .02) {
      width = .1; height = .1;
      x = Math.max(0, Math.min(1 - width, start.x - width / 2));
      y = Math.max(0, Math.min(1 - height, start.y - height / 2));
    }
    setDraftAnchor({ kind: "region", x, y, width: Math.min(width, 1 - x), height: Math.min(height, 1 - y) });
  }
  function saveThread(submit: boolean) {
    if (!selected || !draftAnchor || !draftComment.trim()) return;
    onSaveThread(selected, draftAnchor, draftComment.trim(), submit, centralDesignActive ? selectedNumber : undefined);
    setDraftAnchor(undefined);
    setDraftComment("");
    setMode("comment");
  }

  return <div className="view-stack slides-view">
    <header className="view-header compact"><div><p className="eyebrow">{proposalWorkspace ? "Proposed slide designs" : centralDesignActive ? "Latest converted presentation" : "Source presentation"}</p><h1>{deck.name}</h1><p>{centralDesignActive ? "This is the single current Studio design. Source slides remain unchanged until redesigned; recipe, converted ORNL layout, and merge-safe table continuation choices update this same presentation." : nativeReady ? `Faithful, revision-bound ${proposalWorkspace ? "proposal" : "source"} pixels rendered locally by Microsoft PowerPoint.` : "PowerPoint-native rendering is unavailable; the visible slide is a labeled structural approximation."}</p></div><div className="header-actions"><span className="render-status"><span className={loading ? "loading" : nativeReady ? "ready" : catalog ? "fallback" : ""} />{loading ? "Rendering in PowerPoint…" : centralDesignActive ? `${resultCount}/${centralOutputSlides.length} latest output results built` : nativeReady ? `${nativeRender?.slideCount ?? nativeRender?.slides.length ?? 0} PowerPoint-native previews` : catalog ? `${catalog.slides.length} approximate previews` : "Preview unavailable"}</span>{centralDesignActive && <button className="button primary small" disabled={!deckBuildReady} title={deckBuildReady ? "Save the exact central design shown here as editable PowerPoint." : "Convert and build every slide in Studio before exporting the central presentation."} onClick={onSaveDeck}><FileArrowDown size={16} />Export presentation</button>}</div></header>
    {selected && catalog && <section className="slide-review panel">
      <div className="slide-review-toolbar"><div><button className="button ghost small" onClick={() => setSelectedNumber(undefined)}><ArrowLeft size={15} />Gallery</button><span><strong>{centralDesignActive ? `Output slide ${selectedNumber}` : `Slide ${selected.number}`}</strong><small>{proposalWorkspace ? "Proposal" : centralDesignActive ? selectedSacredTitle ? `source slide ${selected.number} · approved ORNL template composition · native source preserved` : `${selectedOutput?.continuation ? `source slide ${selected.number} · table continuation ${selectedOutput.continuation.segmentOrdinal}/${selectedOutput.continuation.segmentCount}` : `source slide ${selected.number}`} · ${selectedStudioSlide?.recipe.replaceAll("-", " ") ?? "converted design"} · ${selectedNativeReady ? "built" : "build required"}` : `Current · revision ${revision.slice(0, 19).replace("T", " ")}`}</small></span></div><div><button className="button ghost small" disabled={(selectedNumber ?? 1) <= 1} onClick={() => setSelectedNumber((selectedNumber ?? 1) - 1)}><ArrowLeft size={15} />Previous</button><button className="button ghost small" disabled={(selectedNumber ?? 1) >= (centralDesignActive ? centralOutputSlides.length : catalog.slides.length)} onClick={() => setSelectedNumber((selectedNumber ?? 0) + 1)}>Next<ArrowRight size={15} /></button>{centralDesignActive ? <button className={`button ${selectedSacredTitle ? "secondary" : "primary"} small`} disabled={selectedSacredTitle} title={selectedSacredTitle ? "This approved ORNL template composition is source-preserved and cannot be redesigned." : undefined} onClick={() => onOpenStudioSlide(selected.number)}>{selectedSacredTitle ? <ShieldCheck size={16} /> : <SquaresFour size={16} />}{selectedSacredTitle ? "ORNL template locked" : "Choose layout or edit"}</button> : <button className={`button small ${editMode ? "primary" : "secondary"}`} onClick={() => { setMode(editMode ? "review" : "edit"); setDraftAnchor(undefined); }}><Crosshair size={16} />{editMode ? "Editing objects" : "Edit objects"}</button>}<button className={`button small ${commentMode ? "primary" : "secondary"}`} disabled={centralDesignActive && !selectedNativeReady} title={centralDesignActive && !selectedNativeReady ? "Build this exact Studio revision before anchoring comments to its finished result." : undefined} onClick={() => { setMode(commentMode ? "review" : "comment"); setDraftAnchor(undefined); }}><ChatCircleDots size={16} />{commentMode ? "Select a region" : "Comment"}</button></div></div>
      <div className="slide-review-body">
        <div className="slide-review-stage">
          <div className="slide-review-canvas" ref={editorCanvas}><SlideDesignCanvas nativeRender={nativeRender} slideNumber={centralDesignActive ? selectedNumber ?? selected.number : selected.number} catalog={catalog} layout={selected} requireNative={centralDesignActive} label={`${proposalWorkspace ? "Proposed" : centralDesignActive ? "Latest converted" : "Current"} design for ${centralDesignActive ? `output slide ${selectedNumber} from source slide ${selected.number}` : `slide ${selected.number}`}: ${selected.title}`} /><div className={`slide-anchor-layer ${commentMode ? "active" : ""}`} aria-label={commentMode ? "Drag over the exact slide region to comment" : "Slide comment anchors"} onPointerDown={beginAnchor} onPointerMove={moveAnchor} onPointerUp={finishAnchor}>
            {selectedThreads.map((thread, index) => <span key={thread.id} className={`thread-anchor ${thread.status}`} style={{ left: `${thread.anchor.x * 100}%`, top: `${thread.anchor.y * 100}%`, width: `${thread.anchor.width * 100}%`, height: `${thread.anchor.height * 100}%` }} title={thread.comment}><i>{index + 1}</i></span>)}
            {draftAnchor && <span className="thread-anchor draft" style={{ left: `${draftAnchor.x * 100}%`, top: `${draftAnchor.y * 100}%`, width: `${draftAnchor.width * 100}%`, height: `${draftAnchor.height * 100}%` }} />}
          </div>{editMode && !centralDesignActive && <div className="slide-object-layer" onPointerMove={moveObject} onPointerUp={finishObjectDrag} onPointerCancel={finishObjectDrag}>{slideObjects.map((object) => { const geometry = activeGeometry(object); const sceneObject = sceneObjectById.get(object.id); const movable = sceneObject?.operations.move ?? object.canMove; return <button key={object.id} className={`slide-object-box ${selectedObjectId === object.id ? "selected" : ""} ${!movable ? "locked" : ""} ${sceneObject?.fidelityState ?? ""}`} style={{ left: `${geometry.x / slideSize.width * 100}%`, top: `${geometry.y / slideSize.height * 100}%`, width: `${geometry.width / slideSize.width * 100}%`, height: `${geometry.height / slideSize.height * 100}%`, zIndex: Math.min(100, (sceneObject?.zIndex ?? 0) + 1) }} onPointerDown={(event) => beginObjectDrag(event, object)} onClick={(event) => { event.stopPropagation(); selectObject(object); }} title={`${object.name} · ${object.kind} · ${sceneObject ? fidelityLabels[sceneObject.fidelityState] : "legacy object"}`}><span>{sceneObject?.semanticRole ?? object.kind} · {sceneObject ? fidelityLabels[sceneObject.fidelityState] : "legacy"}</span></button>; })}</div>}</div>
          <div className={`slide-representation-note ${selectedNativeReady ? "native" : "fallback"}`}><ShieldCheck size={15} /><span><strong>{centralDesignActive ? selectedNativeReady ? "Central Studio design · PowerPoint native" : "Central Studio design · build required" : editMode ? "Non-destructive object editor" : nativeReady ? "PowerPoint-native representation" : "Approximate OOXML fallback"}</strong> {centralDesignActive ? selectedNativeReady ? "This is the latest built result from the shared Studio scene. Comments and layout edits target this exact slide revision." : selectedConverted ? "The design changed. Build the current revision in Studio before judging or commenting on its finished result." : "The untouched source slide remains the current result until a Studio recipe or ORNL layout is chosen." : editMode ? "Drag a highlighted object, use measured controls, then stage it into Current/Proposal review. Imported slide bytes stay untouched." : nativeReady ? "These pixels came from Microsoft PowerPoint. Structured editable-object overlays remain bound to the same source revision." : nativeRender?.warnings[0] ?? "This diagnostic reconstruction is not authoritative for wrapping, masters, tables, equations, or crop behavior."}</span></div>
        </div>
        <aside className="slide-thread-panel">{editMode && !centralDesignActive ? <><div className="thread-panel-heading"><Crosshair size={19} /><div><strong>Object editor</strong><small>{slideObjects.length} source-bound objects{slideScene?.preservationRequired ? " · preservation active" : ""}</small></div></div>{!selectedObject || !draftGeometry ? <div className="thread-empty"><Crosshair size={24} /><strong>Select an object</strong><span>Every source-bound slide object is outlined with its fidelity state.</span></div> : (() => { const sceneObject = sceneObjectById.get(selectedObject.id); return <div className="object-editor-panel"><div className="object-editor-identity"><span>{sceneObject?.semanticRole ?? selectedObject.kind}</span><strong>{selectedObject.name}</strong><small>ID {selectedObject.id}</small>{sceneObject && <em className={`fidelity-badge ${sceneObject.fidelityState}`}>{fidelityLabels[sceneObject.fidelityState]}</em>}</div>{sceneObject && <p className="fidelity-reason">{sceneObject.fidelityReason}</p>}<div className="geometry-grid">{(["x", "y", "width", "height"] as const).map((field) => <label key={field}><span>{field === "width" ? "W" : field === "height" ? "H" : field.toUpperCase()} (in)</span><input type="number" min="0" step="0.05" value={(draftGeometry[field] / 914_400).toFixed(2)} disabled={(field === "width" || field === "height") ? !(sceneObject?.operations.resize ?? selectedObject.canResize) : !(sceneObject?.operations.move ?? selectedObject.canMove)} onChange={(event) => setInches(field, event.target.value)} /></label>)}</div><div className="nudge-grid"><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: value.y - 45_720 }))}>↑</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: value.x - 45_720 }))}>←</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: value.x + 45_720 }))}>→</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: value.y + 45_720 }))}>↓</button></div><span className="field-label">Align to 0.5-inch safe area</span><div className="align-grid"><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: 457_200 }))}>Left</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: Math.round((slideSize.width - value.width) / 2) }))}>Center</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: slideSize.width - 457_200 - value.width }))}>Right</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: 457_200 }))}>Top</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: Math.round((slideSize.height - value.height) / 2) }))}>Middle</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: slideSize.height - 457_200 - value.height }))}>Bottom</button></div><label className="edit-rationale"><span className="field-label">Design intent</span><textarea value={editRationale} maxLength={700} onChange={(event) => setEditRationale(event.target.value)} placeholder="Example: Align the caption to the figure edge and establish a consistent lower margin." /></label><button className="button primary object-stage-button" disabled={!(sceneObject?.operations.move || sceneObject?.operations.resize || !sceneObject) || JSON.stringify(draftGeometry) === JSON.stringify(baseGeometry(selectedObject))} onClick={() => onStageGeometry(selectedObject, draftGeometry, editRationale)}><Sparkle size={16} />Stage in Current / Proposal</button><small className="object-editor-note">Text, table content, slide count, unsupported native internals, and source bytes are validation-locked.</small></div>; })()}</> : <><div className="thread-panel-heading"><ChatCircleDots size={19} /><div><strong>Design comments</strong><small>{selectedThreads.length} on this slide</small></div></div>
          {commentMode && <div className="thread-composer"><span className="field-label">1. Drag or click the exact area</span><div className={`anchor-readout ${draftAnchor ? "ready" : ""}`}>{draftAnchor ? <><Check size={14} />Region selected</> : <><Crosshair size={14} />Waiting for a region</>}</div><label><span className="field-label">2. Describe the adjustment</span><textarea ref={commentInput} autoFocus value={draftComment} maxLength={4000} onChange={(event) => setDraftComment(event.target.value)} placeholder="Example: Align this caption with the image edge and give it more breathing room." /></label><div className="thread-composer-actions"><button className="button ghost small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(false)}>Save note</button><button className="button primary small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(true)}><PaperPlaneTilt size={15} />Submit to AI</button></div><small>Submitting keeps this slide open for additional comments and creates a scoped thread for MCP. It does not apply or export a change.</small></div>}
          <div className="thread-list">{selectedThreads.length === 0 && <div className="thread-empty"><ChatCircleDots size={24} /><strong>No comments yet</strong><span>Select Comment on slide, then point to the exact area.</span></div>}{selectedThreads.map((thread, index) => <article key={thread.id} className="thread-item"><span>{index + 1}</span><div><strong>{thread.status === "submitted" ? "Ready for AI" : thread.status.replaceAll("-", " ")}</strong><p>{thread.comment}</p><small>Region {Math.round(thread.anchor.x * 100)}%, {Math.round(thread.anchor.y * 100)}% · {new Date(thread.createdAt).toLocaleString()}</small></div><button className="thread-delete" type="button" onClick={() => onDeleteThread(thread.id)} aria-label={`Delete comment ${index + 1}`} title="Delete comment"><Trash size={15} /></button></article>)}</div>
        </>}</aside>
      </div>
    </section>}
    {!selected && <section className="current-slide-gallery panel"><div className="panel-heading"><div><h2>{centralDesignActive ? "Central converted design" : "Slides"}</h2><p>{centralDesignActive ? "Every thumbnail is either the preserved source slide or the latest PowerPoint-native build of its shared Studio revision. Choose a slide to review, comment, or select a converted ORNL layout." : `Select any ${proposalWorkspace ? "proposed" : "current"} design for a closer, zoomable review and location-anchored comments.`}</p></div><span className="quiet-label">{proposalWorkspace ? "Proposal" : centralDesignActive ? "One presentation" : "Current"} · {PRESENTATION_DESIGN_STANDARD.defaults.slide.aspectRatio} target</span></div>
      {loading && <div className="slide-gallery-loading"><ArrowsClockwise className="spinner" size={25} /><strong>Building local slide previews</strong><span>Reading the embedded source, master, layouts, media, and native table structure.</span></div>}
      {!loading && !catalog && <div className="slide-gallery-loading"><Warning size={25} /><strong>Current designs could not be rendered</strong><span>The structural audit remains available; review the reported preview limitation before cleanup.</span></div>}
      {catalog && centralDesignActive && <div className="slide-grid">{centralOutputSlides.map((output) => {
        const slide = catalog.slides.find((item) => item.number === output.sourceSlideNumber);
        if (!slide) return null;
        const inventory = deck.audit?.slides.find((item) => item.number === output.sourceSlideNumber);
        const studioSlide = deck.studioScene?.slides.find((item) => item.slideNumber === output.sourceSlideNumber);
        const built = Boolean(nativeRender?.slides.find((item) => item.number === output.outputSlideNumber));
        const count = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === output.sourceSlideNumber && (thread.outputSlideNumber === output.outputSlideNumber || thread.outputSlideNumber === undefined && output.outputSlideNumber === centralOutputSlides.find((candidate) => candidate.sourceSlideNumber === output.sourceSlideNumber)?.outputSlideNumber)).length;
        return <button className="slide-card" key={`output-${output.outputSlideNumber}-source-${output.sourceSlideNumber}`} onClick={() => setSelectedNumber(output.outputSlideNumber)}><span className="slide-canvas actual"><SlideDesignCanvas nativeRender={nativeRender} slideNumber={output.outputSlideNumber} catalog={catalog} layout={slide} requireNative label={`Open central output slide ${output.outputSlideNumber} from source slide ${output.sourceSlideNumber}: ${slide.title}`} />{count > 0 && <span className="slide-comment-count"><ChatCircleDots size={12} />{count}</span>}</span><span className="slide-meta"><span>{output.outputSlideNumber}</span><span><strong>{inventory?.title ?? slide.title}</strong><small>{output.continuation ? `Source ${output.sourceSlideNumber} · table ${output.continuation.segmentOrdinal}/${output.continuation.segmentCount} · rows ${output.continuation.bodyRowStart}–${output.continuation.bodyRowEnd}` : `Source ${output.sourceSlideNumber} · ${studioSlide?.recipe.replaceAll("-", " ") ?? "source"}`} · {built ? "result ready" : "build required"}</small></span></span></button>;
      })}</div>}
      {catalog && !centralDesignActive && <div className="slide-grid">{catalog.slides.map((slide) => { const inventory = deck.audit?.slides.find((item) => item.number === slide.number); const count = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === slide.number).length; return <button className="slide-card" key={slide.id} onClick={() => setSelectedNumber(slide.number)}><span className="slide-canvas actual"><SlideDesignCanvas nativeRender={nativeRender} slideNumber={slide.number} catalog={catalog} layout={slide} label={`Open current slide ${slide.number}: ${slide.title}`} />{count > 0 && <span className="slide-comment-count"><ChatCircleDots size={12} />{count}</span>}</span><span className="slide-meta"><span>{slide.number}</span><span><strong>{inventory?.title ?? slide.title}</strong><small>{inventory?.tableCount ? `${inventory.tableCount} table${inventory.tableCount === 1 ? "" : "s"} · ` : ""}{inventory?.pictureCount ? `${inventory.pictureCount} image${inventory.pictureCount === 1 ? "" : "s"} · ` : ""}{inventory?.fonts.length ?? 0} fonts</small></span></span></button>; })}</div>}
    </section>}
  </div>;
}

function wrapTemplatePreviewText(text: string, maximumCharacters: number): string[] {
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

type PreviewCanvasCatalog = Pick<TemplateCatalog, "slideWidth" | "slideHeight" | "media">;

function TemplatePreviewElementView({ element, catalog, scale }: { element: TemplatePreviewElement; catalog: PreviewCanvasCatalog; scale: number }) {
  const x = element.x * scale;
  const y = element.y * scale;
  const width = element.width * scale;
  const height = element.height * scale;
  const transform = element.rotation ? `rotate(${element.rotation} ${(element.x + element.width / 2) * scale} ${(element.y + element.height / 2) * scale})` : undefined;
  if (element.kind === "image") {
    const href = element.mediaId ? catalog.media[element.mediaId] : undefined;
    if (!href) return null;
    return <image href={href} x={x} y={y} width={width} height={height} preserveAspectRatio="xMidYMid slice" transform={transform} />;
  }
  if (element.kind === "text") {
    const sourceFontSize = (element.fontSize ?? 18) * 12700;
    const fontSize = sourceFontSize * scale;
    const maximumCharacters = Math.max(4, Math.floor(element.width / Math.max(1, sourceFontSize * .54)));
    const lines = wrapTemplatePreviewText(element.text ?? "", maximumCharacters);
    const lineHeight = fontSize * 1.12;
    const textHeight = lines.length * lineHeight;
    const textX = element.textAlign === "center" ? x + width / 2 : element.textAlign === "right" ? x + width : x;
    const textY = element.verticalAlign === "center" ? y + Math.max(0, (height - textHeight) / 2) : element.verticalAlign === "bottom" ? y + Math.max(0, height - textHeight) : y;
    return <g transform={transform}><text x={textX} y={textY} fill={element.textColor ?? "#373A36"} fontFamily={previewFontStack(element.fontFamily)} fontSize={fontSize} fontWeight={element.fontWeight} textAnchor={element.textAlign === "center" ? "middle" : element.textAlign === "right" ? "end" : "start"} dominantBaseline="hanging">{lines.map((line, index) => <tspan key={`${element.id}-${index}`} x={textX} dy={index === 0 ? 0 : lineHeight}>{line}</tspan>)}</text></g>;
  }
  if (element.placeholderType && !element.fill && !element.stroke) {
    const label = ({ ctrTitle: "Title", subTitle: "Subtitle", title: "Title", body: "Content", pic: "Image", obj: "Content", tbl: "Table", chart: "Chart", media: "Media" } as Record<string, string>)[element.placeholderType] ?? element.placeholderType;
    return <g transform={transform} className="template-placeholder-guide"><rect x={x} y={y} width={width} height={height} /><text x={x + 110000 * scale} y={y + 230000 * scale}>{label}</text></g>;
  }
  if (element.geometry === "ellipse") return <ellipse cx={x + width / 2} cy={y + height / 2} rx={width / 2} ry={height / 2} fill={element.fill ?? "none"} stroke={element.stroke ?? "none"} strokeWidth={(element.strokeWidth ?? 0) * scale} opacity={element.opacity ?? 1} transform={transform} />;
  if (element.geometry === "line") return <line x1={x} y1={y} x2={x + width} y2={y + height} stroke={element.stroke ?? element.fill ?? "#373A36"} strokeWidth={(element.strokeWidth ?? 12700) * scale} opacity={element.opacity ?? 1} transform={transform} />;
  return <rect x={x} y={y} width={width} height={height} fill={element.fill ?? "none"} stroke={element.stroke ?? "none"} strokeWidth={(element.strokeWidth ?? 0) * scale} opacity={element.opacity ?? 1} transform={transform} />;
}

function TemplateLayoutCanvas({ catalog, layout, label }: { catalog: PreviewCanvasCatalog; layout: TemplateLayoutPreview; label?: string }) {
  const viewportWidth = 1200;
  const scale = viewportWidth / catalog.slideWidth;
  const viewportHeight = catalog.slideHeight * scale;
  return <svg className="template-layout-canvas" viewBox={`0 0 ${viewportWidth} ${viewportHeight}`} preserveAspectRatio="xMidYMid meet" role="img" aria-label={label ?? `${layout.name} layout preview`}>
    <rect width={viewportWidth} height={viewportHeight} fill={layout.background} />
    {layout.elements.map((element) => <TemplatePreviewElementView key={element.id} element={element} catalog={catalog} scale={scale} />)}
  </svg>;
}

function SlideDesignCanvas({ nativeRender, slideNumber, catalog, layout, label, requireNative = false }: { nativeRender?: NativeRenderResult; slideNumber: number; catalog?: PreviewCanvasCatalog; layout?: TemplateLayoutPreview; label: string; requireNative?: boolean }) {
  const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((slide) => slide.number === slideNumber) : undefined;
  const nativeSource = useMemo(() => nativeSlide ? `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}` : undefined, [nativeSlide]);
  if (nativeSource) return <img className="native-slide-render" src={nativeSource} width={nativeSlide?.width} height={nativeSlide?.height} alt={label} />;
  if (requireNative) return <span className="proposal-preview-wait"><Monitor size={22} />Build this converted slide to see its latest PowerPoint result</span>;
  if (catalog && layout) return <TemplateLayoutCanvas catalog={catalog} layout={layout} label={`${label} · approximate OOXML fallback`} />;
  return <span className="proposal-preview-wait"><Warning size={22} />Slide render unavailable</span>;
}

function NativeTemplateLayoutCanvas({ catalog, layout, layoutNumber, nativeRender, label, showLabels = false }: { catalog: TemplateCatalog; layout: TemplateLayoutPreview; layoutNumber: number; nativeRender?: NativeRenderResult; label: string; showLabels?: boolean }) {
  const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((slide) => slide.number === layoutNumber) : undefined;
  const nativeSource = useMemo(() => nativeSlide ? `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}` : undefined, [nativeSlide]);
  if (!nativeSource) return <TemplateLayoutCanvas catalog={catalog} layout={layout} label={`${label} · structural fallback`} />;
  const slots = layout.semantic?.slots.filter((slot) => !["footer", "date", "slide-number"].includes(slot.role)) ?? [];
  return <span className={`native-template-layout ${showLabels ? "detailed" : "compact"}`} role="img" aria-label={label}>
    <img className="native-slide-render" src={nativeSource} width={nativeSlide?.width} height={nativeSlide?.height} alt="" />
    <svg className="native-template-slot-overlay" viewBox={`0 0 ${catalog.slideWidth} ${catalog.slideHeight}`} preserveAspectRatio="xMidYMid meet" aria-hidden="true">
      {slots.map((slot) => { const visualSlot = ["image", "table", "chart", "media", "content"].includes(slot.role); return <g key={slot.id} className={`semantic-slot slot-${slot.role}`}><rect x={slot.x} y={slot.y} width={slot.width} height={slot.height} />{(showLabels || visualSlot) && <text x={slot.x + slot.width / 2} y={slot.y + slot.height / 2} textAnchor="middle" dominantBaseline="middle">{slot.role}</text>}</g>; })}
    </svg>
  </span>;
}

function StudioTemplateBase({ catalog, layout }: { catalog: PreviewCanvasCatalog; layout: TemplateLayoutPreview }) {
  const viewportWidth = 1200;
  const scale = viewportWidth / catalog.slideWidth;
  const viewportHeight = catalog.slideHeight * scale;
  return <svg className="studio-template-base" viewBox={`0 0 ${viewportWidth} ${viewportHeight}`} preserveAspectRatio="none" aria-hidden="true">
    <rect width={viewportWidth} height={viewportHeight} fill={layout.background} />
    {layout.elements.filter((element) => !element.placeholderType).map((element) => <TemplatePreviewElementView key={element.id} element={element} catalog={catalog} scale={scale} />)}
  </svg>;
}

function studioTableCellBorderStyle(cellDesign: StudioTableCellDesign | undefined, design: StudioTableDesign): CSSProperties {
  const global = design.borderMode === "none"
    ? { type: "none" as const, color: design.borderColor, widthPt: 0 }
    : { type: "solid" as const, color: design.borderColor, widthPt: design.borderMode === "subtle" ? Math.min(.75, design.borderWidthPt) : design.borderWidthPt };
  const edge = (name: "top" | "right" | "bottom" | "left") => cellDesign?.borders?.[name] ?? global;
  const css = (name: "top" | "right" | "bottom" | "left"): { color: string; width: string | number; style: CSSProperties["borderTopStyle"] } => {
    const border = edge(name);
    return { color: border.color, width: border.type === "none" ? 0 : `${border.widthPt}px`, style: border.type === "dash" ? "dashed" : border.type };
  };
  const top = css("top"); const right = css("right"); const bottom = css("bottom"); const left = css("left");
  return { borderTopColor: top.color, borderTopWidth: top.width, borderTopStyle: top.style, borderRightColor: right.color, borderRightWidth: right.width, borderRightStyle: right.style, borderBottomColor: bottom.color, borderBottomWidth: bottom.width, borderBottomStyle: bottom.style, borderLeftColor: left.color, borderLeftWidth: left.width, borderLeftStyle: left.style };
}

function StudioNodeView({ node, scene, slide, catalog, nativeSlideSource, selected, onPointerDown }: { node: StudioWebNode; scene: StudioWebScene; slide: StudioWebScene["slides"][number]; catalog?: SlideRenderCatalog; nativeSlideSource?: string; selected: boolean; onPointerDown?: (event: ReactPointerEvent<HTMLElement>, mode: "move" | "resize") => void }) {
  if (!node.visible) return null;
  const percent = (value: number, total: number) => `${value / total * 100}%`;
  const style = {
    left: percent(node.frame.x, scene.slideSize.width), top: percent(node.frame.y, scene.slideSize.height), width: percent(node.frame.width, scene.slideSize.width), height: percent(node.frame.height, scene.slideSize.height),
    transform: node.frame.rotation ? `rotate(${node.frame.rotation}deg)` : undefined,
    zIndex: node.zIndex + 10,
    color: node.style.color,
    background: node.style.background,
    borderColor: node.kind === "text" && !node.style.background ? "transparent" : node.style.borderColor,
    borderWidth: node.kind === "text" && !node.style.background ? "0" : `${node.style.borderWidthPt}px`,
    fontFamily: previewFontStack(node.style.fontFamily),
    fontSize: `${node.style.fontSizePt / 9.6}cqw`,
    fontWeight: node.style.fontWeight,
    lineHeight: node.style.lineHeight,
    textAlign: node.style.textAlign,
    padding: `${node.style.paddingPt.top / 72}in ${node.style.paddingPt.right / 72}in ${node.style.paddingPt.bottom / 72}in ${node.style.paddingPt.left / 72}in`,
    alignItems: node.style.verticalAlign === "middle" ? "center" : node.style.verticalAlign === "bottom" ? "flex-end" : "flex-start",
  } as const;
  const media = node.mediaPart ? catalog?.media[node.mediaPart] : undefined;
  const nativeCrop = nativeSlideSource && (["native-object", "connector", "shape"].includes(node.kind) || (node.kind === "image" && !media)) ? <svg className="studio-native-crop" viewBox={`${node.sourceFrame.x} ${node.sourceFrame.y} ${node.sourceFrame.width} ${node.sourceFrame.height}`} preserveAspectRatio="none" aria-label={`${node.name} from the PowerPoint-native source render`}><image href={nativeSlideSource} x="0" y="0" width={scene.slideSize.width} height={scene.slideSize.height} preserveAspectRatio="none" /></svg> : undefined;
  const connectorFrom = node.connector ? slide.nodes.find((candidate) => candidate.id === node.connector?.fromNodeId) : undefined;
  const connectorTo = node.connector ? slide.nodes.find((candidate) => candidate.id === node.connector?.toNodeId) : undefined;
  const connectorStart = node.connector && connectorFrom ? studioConnectorAttachmentPoint(connectorFrom, node.connector.fromSide) : undefined;
  const connectorEnd = node.connector && connectorTo ? studioConnectorAttachmentPoint(connectorTo, node.connector.toSide) : undefined;
  const authoredConnector = node.connector && connectorStart && connectorEnd ? <svg className="studio-authored-connector" viewBox={`${node.frame.x} ${node.frame.y} ${node.frame.width} ${node.frame.height}`} preserveAspectRatio="none" aria-label={`${node.name}, verified editable connector`}><line x1={connectorStart.x} y1={connectorStart.y} x2={connectorEnd.x} y2={connectorEnd.y} stroke={node.connector.stroke} strokeWidth={Math.max(1, node.connector.widthPt * 12_700)} strokeDasharray={node.connector.dash === "dash" ? "6 4" : node.connector.dash === "dashDot" ? "7 3 1 3" : undefined} vectorEffect="non-scaling-stroke" /></svg> : undefined;
  const tableDesign = node.kind === "table" && node.table ? resolvedStudioTableDesign(node) : undefined;
  const content = node.kind === "image" && media ? <img src={media} alt={node.name} style={{ objectFit: node.style.objectFit ?? "contain" }} />
    : node.kind === "image" && nativeCrop ? nativeCrop
    : node.kind === "table" && node.table && tableDesign ? <span className={`studio-native-table borders-${tableDesign.borderMode}`} style={{ gridTemplateColumns: tableDesign.columnWidths.map((value) => `${value}fr`).join(" "), gridTemplateRows: tableDesign.rowHeights.map((value) => `${value}fr`).join(" ") }}>{node.table.cells.map((cell) => { const cellDesign = tableDesign.cellStyles.find((item) => item.cellId === cell.id); const padding = cellDesign?.paddingPt ?? tableDesign.defaultPaddingPt; return <span key={cell.id} className={cell.row <= tableDesign.headerRows ? "header" : "body"} style={{ gridColumn: `${cell.column} / span ${cell.columnSpan}`, gridRow: `${cell.row} / span ${cell.rowSpan}`, background: cellDesign?.fill ?? (cell.fill && /^#[0-9a-f]{6}$/i.test(cell.fill) ? cell.fill : undefined), color: cellDesign?.color, fontSize: cellDesign?.fontSizePt ? `${cellDesign.fontSizePt / 9.6}cqw` : undefined, fontWeight: cellDesign?.fontWeight, textAlign: cellDesign?.textAlign, justifyContent: cellDesign?.textAlign === "center" ? "center" : cellDesign?.textAlign === "right" ? "flex-end" : undefined, alignItems: cellDesign?.verticalAlign === "top" ? "flex-start" : cellDesign?.verticalAlign === "bottom" ? "flex-end" : undefined, padding: `${padding.top / 72}in ${padding.right / 72}in ${padding.bottom / 72}in ${padding.left / 72}in`, ...studioTableCellBorderStyle(cellDesign, tableDesign) }}>{cell.text}</span>; })}</span>
      : node.kind === "text" ? <span className="studio-text-content">{node.text}</span>
        : authoredConnector ?? nativeCrop ?? <span className="studio-native-object-label">{node.name}</span>;
  return <div role="button" tabIndex={0} className={`studio-node kind-${node.kind} role-${node.role} ${selected ? "selected" : ""} ${node.locked ? "locked" : ""}`} style={style} onPointerDown={(event) => onPointerDown?.(event, "move")} aria-label={`${node.name}${node.locked ? ", locked" : ""}`}>{content}{selected && !node.locked && <span className="studio-resize-handle" onPointerDown={(event) => { event.stopPropagation(); onPointerDown?.(event, "resize"); }} />}</div>;
}

function StudioWebCanvas({ scene, slide, catalog, templateCatalog, templateNativeBaseSource, nativeSlideSource, selectedNodeIds, onSelectNode, onMoveNodes }: { scene: StudioWebScene; slide: StudioWebScene["slides"][number]; catalog?: SlideRenderCatalog; templateCatalog?: TemplateCatalog; templateNativeBaseSource?: string; nativeSlideSource?: string; selectedNodeIds: string[]; onSelectNode: (nodeId: string, additive: boolean) => void; onMoveNodes: (updates: Array<{ nodeId: string; frame: StudioWebFrame }>) => void }) {
  const canvas = useRef<HTMLDivElement>(null);
  const drag = useRef<{ nodeIds: string[]; mode: "move" | "resize"; startX: number; startY: number; frames: Record<string, StudioWebFrame> } | undefined>(undefined);
  const [draftFrames, setDraftFrames] = useState<Record<string, StudioWebFrame>>({});
  const [activeGuides, setActiveGuides] = useState<{ x?: number; y?: number }>({});
  const layout = slide.targetLayoutId ? templateCatalog?.layouts.find((item) => item.id === slide.targetLayoutId) : undefined;
  const generatedComponents = studioGeneratedComponents(slide);
  const rhythm = scene.rhythm ?? { safeMarginPt: 18, gridPt: 6, compactGapPt: 8, normalGapPt: 12, primaryGapPt: 18, captionGapPt: 8, titleContentGapPt: 18 };
  const selected = new Set(selectedNodeIds);
  const componentStyle = (component: ReturnType<typeof studioGeneratedComponents>[number]) => ({
    left: `${component.frame.x / scene.slideSize.width * 100}%`, top: `${component.frame.y / scene.slideSize.height * 100}%`, width: `${component.frame.width / scene.slideSize.width * 100}%`, height: `${component.frame.height / scene.slideSize.height * 100}%`,
    background: component.fillColor,
    borderColor: component.lineColor,
    borderWidth: `${component.lineWidthPt}px`,
  });
  function begin(event: ReactPointerEvent<HTMLElement>, node: StudioWebNode, mode: "move" | "resize") {
    event.stopPropagation();
    const additive = event.shiftKey && mode === "move";
    onSelectNode(node.id, additive);
    const figureTreatment = mode === "move" ? slide.figureTreatments.find((treatment) => treatment.nodeIds.includes(node.id)) : undefined;
    if ((node.locked && !figureTreatment) || !canvas.current) return;
    const activeIds = figureTreatment ? figureTreatment.nodeIds : mode === "move" && (selected.has(node.id) || additive)
      ? [...new Set([...(additive ? selectedNodeIds : selected.has(node.id) ? selectedNodeIds : []), node.id])]
      : [node.id];
    const frames = Object.fromEntries(activeIds.map((id) => [id, { ...slide.nodes.find((candidate) => candidate.id === id)!.frame }]));
    event.currentTarget.setPointerCapture(event.pointerId);
    drag.current = { nodeIds: activeIds, mode, startX: event.clientX, startY: event.clientY, frames };
  }
  function unionFrames(frames: StudioWebFrame[]) {
    const x = Math.min(...frames.map((frame) => frame.x));
    const y = Math.min(...frames.map((frame) => frame.y));
    const right = Math.max(...frames.map((frame) => frame.x + frame.width));
    const bottom = Math.max(...frames.map((frame) => frame.y + frame.height));
    return { x, y, width: right - x, height: bottom - y };
  }
  function snapGroup(raw: Record<string, StudioWebFrame>, movingIds: string[]) {
    const group = unionFrames(Object.values(raw));
    const threshold = 5 * 12_700;
    const safe = rhythm.safeMarginPt * 12_700;
    const xTargets = [safe, scene.slideSize.width / 2, scene.slideSize.width - safe];
    const yTargets = [safe, scene.slideSize.height / 2, scene.slideSize.height - safe];
    for (const peer of slide.nodes.filter((node) => node.visible && !movingIds.includes(node.id))) {
      xTargets.push(peer.frame.x, peer.frame.x + peer.frame.width / 2, peer.frame.x + peer.frame.width);
      yTargets.push(peer.frame.y, peer.frame.y + peer.frame.height / 2, peer.frame.y + peer.frame.height);
    }
    const xFeatures = [group.x, group.x + group.width / 2, group.x + group.width];
    const yFeatures = [group.y, group.y + group.height / 2, group.y + group.height];
    const nearest = (targets: number[], features: number[]) => targets.flatMap((target) => features.map((feature) => ({ delta: target - feature, guide: target }))).filter((candidate) => Math.abs(candidate.delta) <= threshold).sort((left, right) => Math.abs(left.delta) - Math.abs(right.delta))[0];
    const snapX = nearest(xTargets, xFeatures);
    const snapY = nearest(yTargets, yFeatures);
    const grid = Math.max(1, rhythm.gridPt * 12_700);
    const dx = snapX?.delta ?? Math.round(group.x / grid) * grid - group.x;
    const dy = snapY?.delta ?? Math.round(group.y / grid) * grid - group.y;
    setActiveGuides({ x: snapX?.guide, y: snapY?.guide });
    return Object.fromEntries(Object.entries(raw).map(([id, frame]) => [id, { ...frame, x: frame.x + dx, y: frame.y + dy }]));
  }
  function move(event: ReactPointerEvent<HTMLDivElement>) {
    if (!drag.current || !canvas.current || event.buttons === 0) return;
    const bounds = canvas.current.getBoundingClientRect();
    const deltaX = (event.clientX - drag.current.startX) / bounds.width * scene.slideSize.width;
    const deltaY = (event.clientY - drag.current.startY) / bounds.height * scene.slideSize.height;
    if (drag.current.mode === "resize") {
      const id = drag.current.nodeIds[0];
      const start = drag.current.frames[id];
      const grid = Math.max(1, rhythm.gridPt * 12_700);
      setDraftFrames({ [id]: { ...start, width: Math.max(grid, Math.round((start.width + deltaX) / grid) * grid), height: Math.max(grid, Math.round((start.height + deltaY) / grid) * grid) } });
      setActiveGuides({});
      return;
    }
    const raw = Object.fromEntries(drag.current.nodeIds.map((id) => { const start = drag.current!.frames[id]; return [id, { ...start, x: start.x + deltaX, y: start.y + deltaY }]; }));
    setDraftFrames(snapGroup(raw, drag.current.nodeIds));
  }
  function finish() {
    if (drag.current && Object.keys(draftFrames).length) onMoveNodes(Object.entries(draftFrames).map(([nodeId, frame]) => ({ nodeId, frame })));
    drag.current = undefined;
    setDraftFrames({});
    setActiveGuides({});
  }
  function keyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "a") {
      event.preventDefault();
      onSelectNode("", false);
      slide.nodes.filter((node) => node.visible && !node.locked).forEach((node) => onSelectNode(node.id, true));
      return;
    }
    if (event.key === "Escape") { onSelectNode("", false); return; }
    const direction = event.key === "ArrowLeft" ? { x: -1, y: 0 } : event.key === "ArrowRight" ? { x: 1, y: 0 } : event.key === "ArrowUp" ? { x: 0, y: -1 } : event.key === "ArrowDown" ? { x: 0, y: 1 } : undefined;
    if (!direction || selectedNodeIds.length === 0) return;
    event.preventDefault();
    const step = Math.max(1, rhythm.gridPt * 12_700) * (event.shiftKey ? 3 : 1);
    const updates = selectedNodeIds.map((id) => slide.nodes.find((node) => node.id === id)).filter((node): node is StudioWebNode => Boolean(node && !node.locked)).map((node) => ({ nodeId: node.id, frame: { ...node.frame, x: node.frame.x + direction.x * step, y: node.frame.y + direction.y * step } }));
    if (updates.length) onMoveNodes(updates);
  }
  const safePercentX = rhythm.safeMarginPt * 12_700 / scene.slideSize.width * 100;
  const safePercentY = rhythm.safeMarginPt * 12_700 / scene.slideSize.height * 100;
  return <div ref={canvas} className="studio-web-canvas" tabIndex={0} role="group" aria-label={`Editable Studio canvas for slide ${slide.slideNumber}. Command-A selects all editable elements; arrow keys nudge the selection; Shift-arrow uses a larger step; Escape clears selection.`} style={{ aspectRatio: `${scene.slideSize.width} / ${scene.slideSize.height}`, background: slide.background }} onKeyDown={keyDown} onPointerMove={move} onPointerUp={finish} onPointerCancel={finish} onClick={(event) => { if (event.target === event.currentTarget) onSelectNode("", false); }}>
    {layout && templateNativeBaseSource ? <img className="studio-template-native-base" src={templateNativeBaseSource} alt="" /> : layout && templateCatalog ? <StudioTemplateBase catalog={templateCatalog} layout={layout} /> : null}
    {generatedComponents.map((component) => <span key={component.id} className={`studio-generated-component ${component.kind}`} style={componentStyle(component)} />)}
    <span className="studio-safe-guide" style={{ left: `${safePercentX}%`, top: `${safePercentY}%`, right: `${safePercentX}%`, bottom: `${safePercentY}%` }} />
    <span className="studio-center-guide vertical" />
    <span className="studio-center-guide horizontal" />
    {activeGuides.x !== undefined && <span className="studio-snap-guide vertical active" style={{ left: `${activeGuides.x / scene.slideSize.width * 100}%` }} />}
    {activeGuides.y !== undefined && <span className="studio-snap-guide horizontal active" style={{ top: `${activeGuides.y / scene.slideSize.height * 100}%` }} />}
    {slide.nodes.map((node) => <StudioNodeView key={node.id} node={draftFrames[node.id] ? { ...node, frame: draftFrames[node.id] } : node} scene={scene} slide={slide} catalog={catalog} nativeSlideSource={nativeSlideSource} selected={selected.has(node.id)} onPointerDown={(event, mode) => begin(event, node, mode)} />)}
  </div>;
}

function StudioExportResult({ slideNumber, source, preview, designed, templateLayout, onBuild }: { slideNumber: number; source?: string; preview?: StudioFreshPreview; designed: boolean; templateLayout: boolean; onBuild: () => void }) {
  if (!source) return <div className="studio-export-wait">{designed ? <Monitor size={30} /> : <ArrowsClockwise className="spinner" size={30} />}<strong>{designed ? "PowerPoint export result required" : "Preparing the PowerPoint source result…"}</strong><p>Studio does not display the approximate web scene as the finished slide. This area remains held until authoritative PowerPoint pixels are available.</p>{designed && <button className="button primary small" onClick={onBuild}><Monitor size={16} />Build export result</button>}</div>;
  const renderedSlides = preview?.nativeRender?.status === "ready" ? preview.nativeRender.slides : [];
  return <div className="studio-export-result">{renderedSlides.length > 1 ? <div className="studio-continuation-results">{renderedSlides.map((slide, index) => <figure key={slide.number}><img className="native-slide-render" src={`data:${slide.mimeType};base64,${bytesToBase64(bytesFrom(slide.bytes))}`} alt={`PowerPoint export result ${index + 1} of ${renderedSlides.length} for source slide ${slideNumber}`} /><figcaption>Output {index + 1} of {renderedSlides.length}{preview?.outputSlides[index]?.continuation ? ` · table rows ${preview.outputSlides[index].continuation!.bodyRowStart}–${preview.outputSlides[index].continuation!.bodyRowEnd}` : ""}</figcaption></figure>)}</div> : <img className="native-slide-render" src={source} alt={`PowerPoint export result for slide ${slideNumber}`} />}<div className="studio-export-result-meta"><span><CheckCircle size={16} weight="fill" />PowerPoint-native export result{renderedSlides.length > 1 ? `s · ${renderedSlides.length} slides` : ""}</span><small>{preview ? "Rendered from the exact editable PPTX bytes available to save." : "The source slide is preserved exactly until a designed export candidate passes native QA."}</small></div>{preview && <div className="studio-export-warning"><Warning size={17} /><span><strong>{templateLayout ? "Converted Template Pack artwork included." : "Studio ORNL composition built without the source master."}</strong> {templateLayout ? "The selected layout's non-placeholder images, fills, and vector rules are compiled into this editable slide. Master-level animations and behaviors are not retained." : "This result intentionally uses the Studio design system instead of inherited source furniture. The visible PowerPoint result is the export truth."}</span></div>}</div>;
}

function StudioExportInspector({ slide, preview, designed }: { slide: StudioWebScene["slides"][number]; preview?: StudioFreshPreview; designed: boolean }) {
  const ready = !designed || Boolean(preview?.nativeRender?.status === "ready" && preview.nativeMeasurement?.status === "ready");
  const templateLayout = slide.recipe === "template-layout";
  const impact = analyzeStudioDesignImpact(slide);
  return <aside className="studio-inspector studio-export-inspector">
    <span className="field-label">Export status</span>
    <strong>{ready ? "Export result ready" : "PowerPoint render required"}</strong>
    <p>{!designed ? "This slide is still using its untouched source PowerPoint result." : preview ? "The visible result was rendered from the exact editable PowerPoint bytes available to save." : templateLayout ? "Build the central slide to compile the converted ORNL layout artwork and exact source content into one editable result." : "Build the export result before judging or saving this design."}</p>
    <dl>
      <div><dt>Visible surface</dt><dd>PowerPoint pixels</dd></div>
      <div><dt>Recipe</dt><dd>{slide.recipe.replaceAll("-", " ")}</dd></div>
      <div><dt>Content mapping</dt><dd>{slide.contentCoverage.exactTextMapped ? "Exact" : "Incomplete"}</dd></div>
      <div><dt>Design impact</dt><dd>{impact.level.replaceAll("-", " ")}</dd></div>
      {preview && <div><dt>Output slides</dt><dd>{preview.slideCount}{preview.slideCount > 1 ? " · continued table" : ""}</dd></div>}
      <div><dt>Template design</dt><dd>{templateLayout ? ready ? "Converted artwork built" : "Converted artwork pending" : preview ? "Studio ORNL system" : designed ? "Pending" : "Source preserved"}</dd></div>
    </dl>
    {preview && <div className="inline-note warning"><Warning size={15} />{templateLayout ? "Converted layout artwork is embedded in the editable slide; original master behaviors and animations are not retained." : "Fresh composition rebuilds the slide with the Studio ORNL system instead of inherited source-master furniture. The large PowerPoint preview is the export truth."}</div>}
    {impact.requirements.some((item) => !item.passed) && <div className="inline-note warning"><Warning size={15} />{impact.requirements.find((item) => !item.passed)?.reason}</div>}
    {!designed && <div className="inline-note"><ShieldCheck size={15} />No redesign is being implied. The visible slide is the original PowerPoint result.</div>}
  </aside>;
}

function StudioVisualNeedsPanel({ slide, conceptResources, onCreate, onHold, onAttach, onDetach, onReconstruct }: {
  slide: StudioWebScene["slides"][number];
  conceptResources: ProjectResource[];
  onCreate: (slideNumber: number, type: StudioVisualNeed["type"]) => void;
  onHold: (slideNumber: number, visualNeedId: string) => void;
  onAttach: (slideNumber: number, visualNeedId: string, resourceId: string) => void;
  onDetach: (slideNumber: number, referenceId: string) => void;
  onReconstruct: (slideNumber: number, referenceId: string) => void;
}) {
  const [selectedResources, setSelectedResources] = useState<Record<string, string>>({});
  const recommendedType: StudioVisualNeed["type"] = slide.figureTreatments.length > 0 || slide.nodes.some((node) => ["connector", "native-object"].includes(node.kind))
    ? "figure-concept"
    : slide.nodes.some((node) => node.kind === "image") ? "image-treatment" : "layout-concept";
  const needs = slide.visualNeeds ?? [];
  return <section className="studio-visual-needs">
    <div className="studio-visual-needs-heading"><div><span className="field-label">Visual direction queue</span><small>Local briefs · model independent</small></div><button className="button secondary small" onClick={() => onCreate(slide.slideNumber, recommendedType)}><Sparkle size={14} />Add recommended brief</button></div>
    {needs.length === 0 ? <p>No concept brief is active. Studio can redesign directly, or create a governed brief when art direction would materially help.</p> : needs.map((need) => {
      const linkedReference = need.linkedConceptReferenceId ? slide.conceptReferences?.find((reference) => reference.id === need.linkedConceptReferenceId) : undefined;
      const linkedResource = linkedReference ? conceptResources.find((resource) => resource.id === linkedReference.resourceId && resource.sha256 === linkedReference.resourceSha256) : undefined;
      const linkedSource = linkedResource?.bytes ? `data:${linkedResource.mediaType};base64,${bytesToBase64(linkedResource.bytes)}` : undefined;
      const selectedResourceId = selectedResources[need.id] ?? conceptResources[0]?.id ?? "";
      return <article key={need.id} className={`studio-visual-need ${need.status}`}>
        <div className="studio-visual-need-title"><strong>{need.type.replaceAll("-", " ")}</strong><span>{need.status.replaceAll("-", " ")}</span></div>
        <p>{need.communicationJob}</p>
        <small>{need.expression} · {need.brandExpression.motif.replaceAll("-", " ")} · {need.brandExpression.accent} · {need.disclosurePolicy.replaceAll("-", " ")} · {need.targetSlot.role.replaceAll("-", " ")}</small>
        {linkedReference ? <><div className="studio-visual-need-concept">{linkedSource ? <img src={linkedSource} alt="Attached concept-only visual direction" /> : <span><Images size={18} />Turn on AI access for preview</span>}<small>{linkedReference.blueprint.summary}</small></div><div className="studio-visual-need-actions"><span><Images size={14} />Concept attached · {linkedReference.blueprint.zones.length} zone{linkedReference.blueprint.zones.length === 1 ? "" : "s"}</span><button className="button secondary small" disabled={!linkedReference.blueprint.zones.length || need.status !== "concept-attached"} title={linkedReference.blueprint.zones.length ? "Rebuild approved visual zones with exact editable source content." : "The AI must describe normalized semantic zones before reconstruction."} onClick={() => onReconstruct(slide.slideNumber, linkedReference.id)}><MagicWand size={13} />Reconstruct editable</button><button className="button ghost small" onClick={() => onDetach(slide.slideNumber, linkedReference.id)}><X size={13} />Detach</button></div></> : need.status === "brief-ready" ? <div className="studio-visual-need-actions attach"><select value={selectedResourceId} disabled={!conceptResources.length} onChange={(event) => setSelectedResources((current) => ({ ...current, [need.id]: event.currentTarget.value }))}><option value="">{conceptResources.length ? "Choose concept image" : "Turn on AI access to share images"}</option>{conceptResources.map((resource) => <option key={resource.id} value={resource.id}>{resource.name}</option>)}</select><button className="button secondary small" disabled={!selectedResourceId} onClick={() => onAttach(slide.slideNumber, need.id, selectedResourceId)}>Attach</button></div> : undefined}
        {!['held', 'resolved'].includes(need.status) && <button className="studio-visual-need-hold" onClick={() => onHold(slide.slideNumber, need.id)}>Hold this brief</button>}
      </article>;
    })}
    <div className="inline-note"><ShieldCheck size={15} />Default prompts disclose structure only. Generated wording, logos, data, claims, screenshots, and technical details remain untrusted.</div>
  </section>;
}

type StudioTableDesignPatch = Partial<Pick<StudioTableDesign, "headerRows" | "columnWidths" | "rowHeights" | "borderMode" | "borderColor" | "borderWidthPt" | "defaultPaddingPt">>;
type StudioTableCellDesignPatch = Omit<StudioTableCellDesign, "cellId">;

function StudioTableInspector({ node, exemplars, continuationPlan, compatibleTableCount, onDesign, onResizeColumn, onResizeRow, onCell, onPublishExemplar, onApplyExemplar, onPlanContinuation, onClearContinuation }: {
  node: StudioWebNode;
  exemplars: StudioTableExemplarDefinition[];
  continuationPlan?: StudioTableContinuationPlan;
  compatibleTableCount: number;
  onDesign: (patch: StudioTableDesignPatch) => void;
  onResizeColumn: (column: number, widthInches: number) => void;
  onResizeRow: (row: number, heightInches: number) => void;
  onCell: (cellId: string, patch: StudioTableCellDesignPatch) => void;
  onPublishExemplar: () => void;
  onApplyExemplar: (definitionId: string) => void;
  onPlanContinuation: (maximumBodyRowsPerSlide: number) => void;
  onClearContinuation: () => void;
}) {
  const [selectedCellId, setSelectedCellId] = useState(node.table?.cells[0]?.id ?? "");
  const [selectedExemplarId, setSelectedExemplarId] = useState(exemplars[0]?.id ?? "");
  const [maximumBodyRows, setMaximumBodyRows] = useState(8);
  useEffect(() => { setSelectedCellId(node.table?.cells[0]?.id ?? ""); }, [node.id]);
  useEffect(() => { if (!exemplars.some((item) => item.id === selectedExemplarId)) setSelectedExemplarId(exemplars[0]?.id ?? ""); }, [exemplars, selectedExemplarId]);
  if (!node.table) return null;
  const design = resolvedStudioTableDesign(node);
  const selectedCell = node.table.cells.find((cell) => cell.id === selectedCellId) ?? node.table.cells[0];
  const cellDesign = design.cellStyles.find((item) => item.cellId === selectedCell?.id);
  const uniformPadding = design.defaultPaddingPt.top === design.defaultPaddingPt.right && design.defaultPaddingPt.top === design.defaultPaddingPt.bottom && design.defaultPaddingPt.top === design.defaultPaddingPt.left ? design.defaultPaddingPt.top : undefined;
  const globalCellBorder = design.borderMode === "none" ? { type: "none" as const, color: design.borderColor, widthPt: 0 } : { type: "solid" as const, color: design.borderColor, widthPt: design.borderMode === "subtle" ? Math.min(.75, design.borderWidthPt) : design.borderWidthPt };
  const cellEdge = (edge: "top" | "right" | "bottom" | "left") => cellDesign?.borders?.[edge] ?? globalCellBorder;
  const updateCellEdge = (edge: "top" | "right" | "bottom" | "left", patch: Partial<ReturnType<typeof cellEdge>>) => selectedCell && onCell(selectedCell.id, { borders: { ...cellDesign?.borders, [edge]: { ...cellEdge(edge), ...patch } } });
  return <section className="studio-table-editor">
    <div className="studio-table-editor-heading"><span className="field-label">Native table component</span><strong>{node.table.rows} rows × {node.table.columns} columns</strong><small>Copy, cell order, merged spans, and semantic color roles remain source locked.</small></div>
    <div className="studio-table-workflow-controls">
      <div><span>Approved exemplar</span><strong>{compatibleTableCount} compatible table{compatibleTableCount === 1 ? "" : "s"}</strong><small>Style propagates only when columns, header rows, and merge structure match.</small></div>
      <button type="button" className="button secondary small" onClick={onPublishExemplar}>Publish this style</button>
      <select aria-label="Approved table exemplar" value={selectedExemplarId} disabled={!exemplars.length} onChange={(event) => setSelectedExemplarId(event.currentTarget.value)}><option value="">{exemplars.length ? "Choose exemplar" : "No approved exemplar"}</option>{exemplars.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select>
      <button type="button" className="button ghost small" disabled={!selectedExemplarId} onClick={() => onApplyExemplar(selectedExemplarId)}>Apply compatible</button>
    </div>
    <div className={`studio-table-continuation ${continuationPlan?.status ?? "idle"}`}>
      <div><span>Continuation planning</span><strong>{continuationPlan ? continuationPlan.status === "ready" ? `${continuationPlan.segments.length} planned slides` : "Needs resolution" : "Not planned"}</strong><small>{continuationPlan?.status === "ready" ? `Repeats ${continuationPlan.headerRows} header row${continuationPlan.headerRows === 1 ? "" : "s"}; body rows break only between merge-safe groups.` : continuationPlan?.blockers[0] ?? "Use when the table cannot remain legible in one PowerPoint-native region."}</small></div>
      <label><span>Body rows / slide</span><input type="number" min="1" max="40" value={maximumBodyRows} onChange={(event) => setMaximumBodyRows(Number(event.currentTarget.value))} /></label>
      <button type="button" className="button secondary small" onClick={() => onPlanContinuation(maximumBodyRows)}>{continuationPlan ? "Replan" : "Plan continuation"}</button>
      {continuationPlan && <button type="button" className="button ghost small" onClick={onClearContinuation}>Clear plan</button>}
      {continuationPlan?.segments.map((segment) => <small key={segment.ordinal}>Slide {segment.ordinal}: header + rows {segment.bodyRowStart}–{segment.bodyRowEnd}</small>)}
    </div>
    <div className="studio-table-global-controls">
      <label><span>Header rows</span><input key={`headers-${design.headerRows}`} type="number" min="0" max={node.table.rows} defaultValue={design.headerRows} onBlur={(event) => onDesign({ headerRows: Number(event.currentTarget.value) })} /></label>
      <label><span>Borders</span><select value={design.borderMode} onChange={(event) => onDesign({ borderMode: event.currentTarget.value as StudioTableDesign["borderMode"] })}><option value="none">None</option><option value="subtle">Subtle</option><option value="full">Full grid</option></select></label>
      <label><span>Border</span><input type="color" value={design.borderColor} onChange={(event) => onDesign({ borderColor: event.currentTarget.value })} /></label>
      <label><span>Stroke pt</span><input key={`stroke-${design.borderWidthPt}`} type="number" min="0" max="6" step=".25" defaultValue={design.borderWidthPt} onBlur={(event) => onDesign({ borderWidthPt: Number(event.currentTarget.value) })} /></label>
      <label><span>Padding pt</span><input key={`padding-${uniformPadding ?? "mixed"}`} type="number" min="0" max="36" step="1" defaultValue={uniformPadding ?? 6} onBlur={(event) => { const value = Number(event.currentTarget.value); onDesign({ defaultPaddingPt: { top: value, right: value, bottom: value, left: value } }); }} /></label>
    </div>
    <div className="studio-table-dimensions"><div><span>Column widths</span>{design.columnWidths.map((weight, index) => <label key={`column-${index + 1}-${weight}`}><small>C{index + 1}</small><input type="number" min=".35" max={(node.frame.width / 914_400).toFixed(2)} step=".05" defaultValue={(node.frame.width / 914_400 * weight).toFixed(2)} onBlur={(event) => onResizeColumn(index + 1, Number(event.currentTarget.value))} /><em>in</em></label>)}</div><div><span>Row heights</span>{design.rowHeights.map((weight, index) => <label key={`row-${index + 1}-${weight}`}><small>R{index + 1}</small><input type="number" min=".18" max={(node.frame.height / 914_400).toFixed(2)} step=".05" defaultValue={(node.frame.height / 914_400 * weight).toFixed(2)} onBlur={(event) => onResizeRow(index + 1, Number(event.currentTarget.value))} /><em>in</em></label>)}</div></div>
    <div className="studio-table-cell-grid" style={{ gridTemplateColumns: `repeat(${Math.max(1, node.table.columns)}, minmax(0, 1fr))` }}>{node.table.cells.map((cell) => <button key={cell.id} className={cell.id === selectedCell?.id ? "selected" : ""} style={{ gridColumn: `${cell.column} / span ${cell.columnSpan}`, gridRow: `${cell.row} / span ${cell.rowSpan}` }} onClick={() => setSelectedCellId(cell.id)} title={`Row ${cell.row}, column ${cell.column}${cell.semanticColorRole ? ` · ${cell.semanticColorRole}` : ""}`}><span>R{cell.row} C{cell.column}</span><small>{cell.text || "Empty"}</small></button>)}</div>
    {selectedCell && <div className="studio-table-cell-controls"><div><strong>Cell R{selectedCell.row} C{selectedCell.column}</strong><small>{selectedCell.semanticColorRole ? `Semantic role: ${selectedCell.semanticColorRole}` : "Source-bound content"}</small></div><label><span>Fill</span><input type="color" value={cellDesign?.fill ?? selectedCell.fill ?? (selectedCell.row <= design.headerRows ? "#00454D" : "#FFFFFF")} onChange={(event) => onCell(selectedCell.id, { fill: event.currentTarget.value })} /></label><label><span>Text</span><input type="color" value={cellDesign?.color ?? (selectedCell.row <= design.headerRows && !selectedCell.semanticColorRole ? "#FFFFFF" : node.style.color)} onChange={(event) => onCell(selectedCell.id, { color: event.currentTarget.value })} /></label><label><span>Size</span><input key={`cell-size-${selectedCell.id}-${cellDesign?.fontSizePt ?? node.style.fontSizePt}`} type="number" min="10" max="40" step=".25" defaultValue={cellDesign?.fontSizePt ?? node.style.fontSizePt} onBlur={(event) => onCell(selectedCell.id, { fontSizePt: Number(event.currentTarget.value) })} /></label><label><span>Weight</span><select value={cellDesign?.fontWeight ?? (selectedCell.row <= design.headerRows ? 700 : 400)} onChange={(event) => onCell(selectedCell.id, { fontWeight: Number(event.currentTarget.value) as 400 | 600 | 700 })}><option value="400">Regular</option><option value="600">Semibold</option><option value="700">Bold</option></select></label><label><span>Align</span><select value={cellDesign?.textAlign ?? node.style.textAlign} onChange={(event) => onCell(selectedCell.id, { textAlign: event.currentTarget.value as "left" | "center" | "right" })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label><label><span>Vertical</span><select value={cellDesign?.verticalAlign ?? node.style.verticalAlign} onChange={(event) => onCell(selectedCell.id, { verticalAlign: event.currentTarget.value as "top" | "middle" | "bottom" })}><option value="top">Top</option><option value="middle">Middle</option><option value="bottom">Bottom</option></select></label><div className="studio-table-cell-borders"><span>Cell edge rules</span><button type="button" onClick={() => onCell(selectedCell.id, { borders: { top: globalCellBorder, right: globalCellBorder, bottom: globalCellBorder, left: globalCellBorder } })}>Match global</button>{(["top", "right", "bottom", "left"] as const).map((edge) => { const border = cellEdge(edge); return <div key={edge}><strong>{edge}</strong><select aria-label={`${edge} border type`} value={border.type} onChange={(event) => updateCellEdge(edge, { type: event.currentTarget.value as typeof border.type })}><option value="none">None</option><option value="solid">Solid</option><option value="dash">Dash</option></select><input aria-label={`${edge} border color`} type="color" value={border.color} onChange={(event) => updateCellEdge(edge, { color: event.currentTarget.value })} /><input aria-label={`${edge} border width`} key={`${selectedCell.id}-${edge}-${border.widthPt}`} type="number" min="0" max="6" step=".25" defaultValue={border.widthPt} onBlur={(event) => updateCellEdge(edge, { widthPt: Number(event.currentTarget.value) })} /></div>; })}</div></div>}
  </section>;
}

function StudioConnectorInspector({ node, slide, onUpdate }: { node: StudioWebNode; slide: StudioWebScene["slides"][number]; onUpdate: (design: StudioConnectorDesign) => void }) {
  const treatment = slide.figureTreatments.find((item) => item.nodeIds.includes(node.id));
  const candidates = treatment?.nodeIds.map((id) => slide.nodes.find((item) => item.id === id)).filter((item): item is StudioWebNode => Boolean(item && item.id !== node.id && item.kind !== "connector")) ?? [];
  const eligible = treatment?.verificationStatus === "verified" && treatment.relationshipPolicy === "editable-diagram" && candidates.length >= 2;
  const current: StudioConnectorDesign = node.connector ?? { fromNodeId: candidates[0]?.id ?? "", toNodeId: candidates[1]?.id ?? "", fromSide: "right", toSide: "left", stroke: "#00662C", widthPt: 1.5, dash: "solid", beginArrow: "none", endArrow: "triangle", verificationStatus: "verified" };
  const patch = (value: Partial<StudioConnectorDesign>) => onUpdate({ ...current, ...value, verificationStatus: "verified" });
  return <section className="studio-connector-editor"><span className="field-label">Verified connector</span>{!eligible ? <div className="inline-note warning"><Warning size={15} />Treat the connector and both endpoint objects as one figure, verify its information, then choose Editable diagram before authoring relationships.</div> : <><div className="studio-connector-endpoints"><label><span>From</span><select value={current.fromNodeId} onChange={(event) => patch({ fromNodeId: event.currentTarget.value })}>{candidates.map((candidate) => <option key={candidate.id} value={candidate.id}>{candidate.name}</option>)}</select></label><label><span>Side</span><select value={current.fromSide} onChange={(event) => patch({ fromSide: event.currentTarget.value as StudioConnectorDesign["fromSide"] })}>{["top", "right", "bottom", "left", "center"].map((side) => <option key={side}>{side}</option>)}</select></label><label><span>To</span><select value={current.toNodeId} onChange={(event) => patch({ toNodeId: event.currentTarget.value })}>{candidates.map((candidate) => <option key={candidate.id} value={candidate.id}>{candidate.name}</option>)}</select></label><label><span>Side</span><select value={current.toSide} onChange={(event) => patch({ toSide: event.currentTarget.value as StudioConnectorDesign["toSide"] })}>{["top", "right", "bottom", "left", "center"].map((side) => <option key={side}>{side}</option>)}</select></label></div><div className="studio-connector-style"><label><span>Stroke</span><input type="color" value={current.stroke} onChange={(event) => patch({ stroke: event.currentTarget.value })} /></label><label><span>Width</span><input key={`connector-width-${current.widthPt}`} type="number" min=".25" max="8" step=".25" defaultValue={current.widthPt} onBlur={(event) => patch({ widthPt: Number(event.currentTarget.value) })} /></label><label><span>Line</span><select value={current.dash} onChange={(event) => patch({ dash: event.currentTarget.value as StudioConnectorDesign["dash"] })}><option value="solid">Solid</option><option value="dash">Dash</option><option value="dashDot">Dash-dot</option></select></label><label><span>Arrow</span><select value={current.endArrow} onChange={(event) => patch({ endArrow: event.currentTarget.value as StudioConnectorDesign["endArrow"] })}>{["none", "arrow", "stealth", "triangle", "diamond", "oval"].map((arrow) => <option key={arrow}>{arrow}</option>)}</select></label></div>{!node.connector && <button className="button secondary small" disabled={!current.fromNodeId || !current.toNodeId || current.fromNodeId === current.toNodeId} onClick={() => onUpdate(current)}>Create editable connector</button>}<div className="inline-note"><ShieldCheck size={15} />Endpoints remain attached to stable Studio objects when those objects move.</div></>}</section>;
}

function StudioView({ deck, catalog, nativeRender, freshPreviews, templateCatalog, templateNativeRender, resources, requestedSlideNumber, deckBuildReady, qualification, canUndo, canRedo, onUndo, onRedo, onInitialize, onRecompose, onMoveNode, onMoveNodes, onStyleNode, onPublishComponent, onArrangeSelection, onRepairObjectiveIssues, onUpdateTableDesign, onResizeTableColumn, onResizeTableRow, onUpdateTableCell, onPublishTableExemplar, onApplyTableExemplar, onPlanTableContinuation, onClearTableContinuation, onUpdateConnector, onUpdateFigure, onCreateVisualNeed, onHoldVisualNeed, onAttachConcept, onDetachConcept, onReconstructConcept, onPreviewFresh, onPreviewAll, onQualify, onRevealQualification, onSaveFresh, onSaveDeck }: { deck?: DeckJob; catalog?: SlideRenderCatalog; nativeRender?: NativeRenderResult; freshPreviews?: Record<string, StudioFreshPreview>; templateCatalog?: TemplateCatalog; templateNativeRender?: NativeRenderResult; resources: ProjectResource[]; requestedSlideNumber?: number; deckBuildReady: boolean; qualification?: StudioDeckQualification; canUndo: boolean; canRedo: boolean; onUndo: () => void; onRedo: () => void; onInitialize: () => void; onRecompose: (slideNumber: number, recipe: StudioLayoutRecipe, layoutId?: string) => void; onMoveNode: (slideNumber: number, nodeId: string, frame: StudioWebFrame) => void; onMoveNodes: (slideNumber: number, updates: Array<{ nodeId: string; frame: StudioWebFrame }>) => void; onStyleNode: (slideNumber: number, nodeId: string, patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">>) => void; onPublishComponent: (slideNumber: number, nodeId: string) => void; onArrangeSelection: (slideNumber: number, nodeIds: string[], mode: StudioConstraintRequest["mode"]) => void; onRepairObjectiveIssues: (slideNumber: number) => void; onUpdateTableDesign: (slideNumber: number, nodeId: string, patch: StudioTableDesignPatch) => void; onResizeTableColumn: (slideNumber: number, nodeId: string, column: number, widthInches: number) => void; onResizeTableRow: (slideNumber: number, nodeId: string, row: number, heightInches: number) => void; onUpdateTableCell: (slideNumber: number, nodeId: string, cellId: string, patch: StudioTableCellDesignPatch) => void; onPublishTableExemplar: (slideNumber: number, nodeId: string) => void; onApplyTableExemplar: (definitionId: string) => void; onPlanTableContinuation: (slideNumber: number, nodeId: string, maximumBodyRowsPerSlide: number) => void; onClearTableContinuation: (slideNumber: number, nodeId: string) => void; onUpdateConnector: (slideNumber: number, nodeId: string, design: StudioConnectorDesign) => void; onUpdateFigure: (slideNumber: number, treatment: StudioFigureTreatment) => void; onCreateVisualNeed: (slideNumber: number, type: StudioVisualNeed["type"]) => void; onHoldVisualNeed: (slideNumber: number, visualNeedId: string) => void; onAttachConcept: (slideNumber: number, visualNeedId: string, resourceId: string) => void; onDetachConcept: (slideNumber: number, referenceId: string) => void; onReconstructConcept: (slideNumber: number, referenceId: string) => void; onPreviewFresh: (slideNumber: number) => void; onPreviewAll: () => void; onQualify: () => void; onRevealQualification: () => void; onSaveFresh: (slideNumber: number) => void; onSaveDeck: () => void }) {
  const [selectedNumber, setSelectedNumber] = useState(1);
  const [selectedNodeIds, setSelectedNodeIds] = useState<string[]>([]);
  const [showEditor, setShowEditor] = useState(false);
  const [showConsistency, setShowConsistency] = useState(false);
  const scene = deck?.studioScene;
  const selectedNodeId = selectedNodeIds.at(-1);
  const consistency = useMemo(() => scene ? analyzeStudioDeckConsistency(scene) : undefined, [scene]);
  useEffect(() => { setSelectedNumber(requestedSlideNumber ?? 1); setSelectedNodeIds([]); setShowEditor(Boolean(requestedSlideNumber)); }, [deck?.id, requestedSlideNumber]);
  if (!deck?.audit) return <NoSelection message="Select an audited deck before entering Studio redesign mode." />;
  if (!scene) return <div className="view-stack studio-empty"><header className="view-header compact"><div><p className="eyebrow">HTML-first presentation design</p><h1>Build a Studio Web Scene</h1><p>Extract exact source content into a semantic 16:9 web canvas. The original PowerPoint remains immutable.</p></div></header><section className="designs-empty"><span className="designs-empty-icon"><Code size={34} /></span><h2>Turn this deck into an editable web design system</h2><p>Studio will preserve source bindings while giving the AI and the user one component-based canvas for layout, hierarchy, tables, imagery, and ORNL templates.</p><button className="button primary large" onClick={onInitialize}><Sparkle size={18} />Create Studio scene</button></section></div>;
  const slide = scene.slides.find((item) => item.slideNumber === selectedNumber) ?? scene.slides[0];
  const sacredTitleSlide = isProtectedOrnlTemplateSlide(deck, slide.slideNumber);
  const editorVisible = showEditor && !sacredTitleSlide;
  const selectedNode = slide?.nodes.find((node) => node.id === selectedNodeId);
  const compatibleComponentInstances = selectedNode?.component ? compatibleStudioComponentInstances(scene, slide.slideNumber, selectedNode.id) : [];
  const compatibleTableCount = selectedNode?.table ? compatibleStudioTableInstances(scene, slide.slideNumber, selectedNode.id).length : 0;
  const selectedTableContinuationPlan = selectedNode?.table ? scene.tableContinuationPlans?.find((plan) => plan.sourceSlideNumber === slide.slideNumber && plan.tableNodeId === selectedNode.id) : undefined;
  const selectedFigureTreatment = selectedNodeId ? slide?.figureTreatments.find((treatment) => treatment.nodeIds.includes(selectedNodeId)) : undefined;
  const selectedNodes = slide.nodes.filter((node) => selectedNodeIds.includes(node.id));
  const canCreateFigureTreatment = selectedNodes.some((node) => ["image", "native-object", "shape", "connector"].includes(node.kind));
  function createFigureTreatment() {
    if (!selectedNodes.length || !canCreateFigureTreatment) return;
    onUpdateFigure(slide.slideNumber, {
      id: `figure-${slide.slideNumber}-${crypto.randomUUID()}`,
      nodeIds: selectedNodes.map((node) => node.id),
      mode: "preserve-and-frame",
      verificationStatus: "source-locked",
      intentSummary: `Preserve ${selectedNodes.map((node) => node.name).join(", ")} as one technical evidence unit.`,
      informationInventory: selectedNodes.map((node) => `${node.kind}: ${node.name}`).slice(0, 40),
      invariants: ["Preserve all visible labels, values, arrows, sequence, grouping, and source relationships exactly."],
      rationale: "Human grouped these source elements so Studio and AI design around them as one relationship-preserving figure.",
      relationshipPolicy: "preserve-internal",
      lockAspectRatio: true,
    });
  }
  function patchSelectedFigure(patch: Partial<StudioFigureTreatment>) {
    if (selectedFigureTreatment) onUpdateFigure(slide.slideNumber, { ...selectedFigureTreatment, ...patch });
  }
  function patchFigureCrop(field: "left" | "top" | "right" | "bottom", percentage: number) {
    if (!selectedFigureTreatment || !Number.isFinite(percentage)) return;
    patchSelectedFigure({ crop: { left: 0, top: 0, right: 0, bottom: 0, ...selectedFigureTreatment.crop, [field]: percentage / 100 } });
  }
  const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((item) => item.number === slide?.slideNumber) : undefined;
  const nativeSlideSource = nativeSlide ? `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}` : undefined;
  const templateLayoutIndex = slide.targetLayoutId ? templateCatalog?.layouts.findIndex((layout) => layout.id === slide.targetLayoutId) ?? -1 : -1;
  const templateNativeSlide = templateLayoutIndex >= 0 && templateNativeRender?.status === "ready" ? templateNativeRender.slides.find((item) => item.number === templateLayoutIndex + 1) : undefined;
  const templateNativeBaseSource = templateNativeSlide ? `data:${templateNativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(templateNativeSlide.bytes))}` : undefined;
  const selectedFreshPreview = deck && slide ? freshPreviews?.[`${deck.id}:${slide.slideNumber}`] : undefined;
  const requiresFreshComposition = Boolean(slide?.nodes.some((node) => node.visible && node.sourceBinding !== "editable-object"));
  const currentFreshPreview = selectedFreshPreview?.slideUpdatedAt === slide.updatedAt ? selectedFreshPreview : undefined;
  let currentCritique: ReturnType<typeof critiqueStudioSlide> | undefined;
  if (currentFreshPreview?.slideCount === 1 && currentFreshPreview.nativeMeasurement?.status === "ready" && currentFreshPreview.nativeMeasurement.authority === "powerpoint-native") {
    try { currentCritique = critiqueStudioSlide(scene, slide.slideNumber, currentFreshPreview.nativeMeasurement); } catch { currentCritique = undefined; }
  }
  const freshNativeSlide = currentFreshPreview?.nativeRender?.status === "ready" ? currentFreshPreview.nativeRender.slides[0] : undefined;
  const freshNativeSource = freshNativeSlide ? `data:${freshNativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(freshNativeSlide.bytes))}` : undefined;
  const needsDesignedExport = slide.status === "designed" && slide.recipe !== "source";
  const needsFreshExport = needsDesignedExport;
  const exportResultSource = needsFreshExport ? freshNativeSource : nativeSlideSource;
  const catalogSlide = catalog?.slides.find((item) => item.number === slide?.slideNumber);
  const designedCount = scene.slides.filter((item) => item.status === "designed").length;
  const resultStatus = (item: StudioWebScene["slides"][number]) => {
    if (item.status !== "designed" || item.recipe === "source") return nativeRender?.status === "ready" && Boolean(nativeRender.slides.find((rendered) => rendered.number === item.slideNumber)) ? "ready" : "pending";
    const preview = freshPreviews?.[`${deck.id}:${item.slideNumber}`];
    return preview?.slideUpdatedAt === item.updatedAt && preview.nativeRender?.status === "ready" && preview.nativeMeasurement?.status === "ready" ? "ready" : "pending";
  };
  const exportReadyCount = scene.slides.filter((item) => resultStatus(item) === "ready").length;
  const allExportResultsReady = exportReadyCount === scene.slides.length;
  const conceptResources = resources.filter((resource) => resource.kind === "image" && resource.mcpAccess === "preview" && Boolean(resource.bytes?.byteLength));
  const rememberedDesign = scene.designMemory?.find((entry) => entry.contentSignature === studioSlideContentSignature(slide));
  const recommended = sacredTitleSlide ? "source" : rememberedDesign?.recipe && rememberedDesign.recipe !== "template-layout" ? rememberedDesign.recipe : slide ? recommendedStudioRecipe(slide) : "ornl-title-content";
  const recommendedTemplate = slide && templateCatalog ? rankLayoutCompatibility(templateCatalog.layouts, contentProfileForSlide(deck, slide.slideNumber))[0] : undefined;
  const qualificationHeadline = !qualification ? undefined : qualification.report.status === "objective-failure"
    ? `${qualification.report.totals.blockerIssues} blocker${qualification.report.totals.blockerIssues === 1 ? "" : "s"} · ${qualification.report.totals.majorIssues} major issue${qualification.report.totals.majorIssues === 1 ? "" : "s"}`
    : qualification.report.status === "review-complete" ? "Every slide passed revision-bound visual review"
      : qualification.report.status === "revision-required" ? `${qualification.report.visualAcceptance.revisionSlideCount} slide${qualification.report.visualAcceptance.revisionSlideCount === 1 ? "" : "s"} need another design pass`
        : qualification.report.status === "held" ? `${qualification.report.visualAcceptance.heldSlideCount} slide${qualification.report.visualAcceptance.heldSlideCount === 1 ? "" : "s"} held for human review`
          : "Objective gates passed · visual review required";
  const figureControlPanel = editorVisible && selectedNodes.length > 0 ? <section className="studio-figure-toolbar panel">
    <div><span className="field-label">Relationship-aware figure</span><strong>{selectedFigureTreatment ? selectedFigureTreatment.intentSummary : `${selectedNodes.length} selected element${selectedNodes.length === 1 ? "" : "s"}`}</strong><small>{selectedFigureTreatment ? `${selectedFigureTreatment.mode.replaceAll("-", " ")} · ${selectedFigureTreatment.verificationStatus.replaceAll("-", " ")} · ${selectedFigureTreatment.relationships?.length ?? 0} recorded relationships` : "Shift-click the complete figure, labels, captions, and callouts before grouping."}</small></div>
    {!selectedFigureTreatment ? <button className="button secondary small" disabled={!canCreateFigureTreatment} title={canCreateFigureTreatment ? "Keep this selection together as one technical evidence unit." : "Select at least one image, native object, shape, or connector."} onClick={createFigureTreatment}><CirclesThreePlus size={16} />Treat selection as figure</button> : <>
      <label><span>Relationships</span><select value={selectedFigureTreatment.relationshipPolicy ?? "preserve-internal"} onChange={(event) => patchSelectedFigure({ relationshipPolicy: event.currentTarget.value as StudioFigureTreatment["relationshipPolicy"] })}><option value="preserve-internal">Preserve internal</option><option value="reflow-annotations">Reflow annotations</option><option value="editable-diagram" disabled={selectedFigureTreatment.verificationStatus !== "verified"}>Editable diagram</option></select></label>
      <label className="studio-aspect-toggle"><input type="checkbox" checked={selectedFigureTreatment.lockAspectRatio !== false} onChange={(event) => patchSelectedFigure({ lockAspectRatio: event.currentTarget.checked })} /><span>Lock aspect</span></label>
      <div className="studio-figure-numbers"><span>Crop %</span>{(["left", "top", "right", "bottom"] as const).map((field) => <label key={field}><span>{field[0].toUpperCase()}</span><input type="number" min="0" max="95" step="1" defaultValue={Math.round((selectedFigureTreatment.crop?.[field] ?? 0) * 100)} onBlur={(event) => patchFigureCrop(field, Number(event.currentTarget.value))} /></label>)}</div>
      <div className="studio-figure-numbers focal"><span>Focal %</span>{(["x", "y"] as const).map((axis) => <label key={axis}><span>{axis.toUpperCase()}</span><input type="number" min="0" max="100" step="1" defaultValue={Math.round((selectedFigureTreatment.focalPoint?.[axis] ?? .5) * 100)} onBlur={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) patchSelectedFigure({ focalPoint: { x: selectedFigureTreatment.focalPoint?.x ?? .5, y: selectedFigureTreatment.focalPoint?.y ?? .5, [axis]: value / 100 } }); }} /></label>)}</div>
    </>}
  </section> : undefined;
  return <div className="view-stack studio-view">
    <header className="view-header compact">
      <div><p className="eyebrow">PowerPoint export results</p><h1>{deck.name}</h1><p>Studio shows authoritative Microsoft PowerPoint pixels by default. The semantic web scene is available only as an explicit editing surface and is never presented as the finished slide.</p></div>
      <div className="header-actions studio-header-actions">
        <span className="standard-version">{designedCount}/{scene.slides.length} designed</span>
        <span className={`standard-version ${allExportResultsReady ? "ready" : ""}`}>{exportReadyCount}/{scene.slides.length} results ready</span>
        {requiresFreshComposition && <span className="studio-mode-chip">Fresh composition</span>}
        <span className="studio-history-controls"><button className="button ghost small" disabled={!canUndo} onClick={onUndo} title="Undo the last Studio design transaction"><ArrowCounterClockwise size={15} />Undo</button><button className="button ghost small" disabled={!canRedo} onClick={onRedo} title="Redo the last undone Studio design transaction"><ArrowClockwise size={15} />Redo</button></span>
        <button className="button ghost small" disabled={!consistency?.designedSlideCount} onClick={() => setShowConsistency((value) => !value)} title="Review title grids, repeated component typography, and related table systems across the deck"><ListChecks size={15} />Consistency {consistency?.issueCount ? `(${consistency.issueCount})` : ""}</button>
        <button className="button primary small" disabled={designedCount === 0} title={designedCount === 0 ? "Apply a Studio design recipe before rebuilding export results." : "Sequentially rebuild every designed slide in PowerPoint; untouched source slides remain preserved."} onClick={onPreviewAll}><SquaresFour size={16} />{allExportResultsReady ? "Rebuild all results" : "Build all results"}</button>
        <button className="button secondary small" disabled={!deckBuildReady} title={deckBuildReady ? "Reopen the exact central PPTX in Microsoft PowerPoint and export every source/candidate slide as a private 2,200-pixel PNG evidence bundle." : "Build the central presentation before qualification."} onClick={onQualify}><MagnifyingGlass size={16} />Inspect all</button>
        <button className="button primary small" disabled={!deckBuildReady} title={deckBuildReady ? "Save the current central design as one editable PowerPoint presentation." : "Convert and build every slide before exporting one central presentation."} onClick={onSaveDeck}><FileArrowDown size={16} />Export presentation</button>
        <button className="button secondary small" disabled={sacredTitleSlide} title={sacredTitleSlide ? "This approved ORNL template composition is sacred and source-preserved." : undefined} onClick={() => { setShowEditor((value) => !value); setSelectedNodeIds([]); }}><Code size={16} />{sacredTitleSlide ? "ORNL template locked" : editorVisible ? "Close scene editor" : "Edit design scene"}</button>
        {needsFreshExport && <button className="button secondary small" disabled={!slide.contentCoverage.exactTextMapped} title={!slide.contentCoverage.exactTextMapped ? "Grouped or unsupported source text must be atomized before export rendering." : undefined} onClick={() => onPreviewFresh(slide.slideNumber)}><Monitor size={16} />{currentFreshPreview ? "Rebuild export result" : "Build export result"}</button>}
        {currentFreshPreview && <button className="button secondary small" disabled={!freshNativeSource || currentFreshPreview.nativeMeasurement?.status !== "ready" || currentFreshPreview.nativeMeasurement.authority !== "powerpoint-native"} onClick={() => onSaveFresh(slide.slideNumber)}><FileArrowDown size={16} />Save this result</button>}
      </div>
    </header>
    {qualification && <section className={`studio-qualification-bar ${qualification.report.status}`}><span className="studio-qualification-mark"><ShieldCheck size={20} /></span><div><span className="field-label">Latest PowerPoint-native qualification · attempt {qualification.report.iteration.attempt}</span><strong>{qualificationHeadline}</strong><small>{qualification.report.totals.slides} source/candidate PNG pairs · {qualification.report.visualAcceptance.reviewedSlideCount}/{qualification.report.totals.slides} visually reviewed · objective trend {qualification.report.iteration.objectiveTrend.replaceAll("-", " ")} · exact scene {qualification.sceneRevision.slice(0, 20)}…</small></div><button className="button secondary small" onClick={onRevealQualification}><FolderOpen size={16} />Open evidence</button></section>}
    {!editorVisible && currentCritique && <section className={`studio-repair-pass panel ${currentCritique.verdict}`}><div><span className="field-label">Found issues · PowerPoint-native pass {currentCritique.iteration.currentPass}/3</span><strong>{currentCritique.issues.length ? `${currentCritique.blockerCount} blocker · ${currentCritique.majorCount} major · ${currentCritique.minorCount} minor` : "Objective checks are clear"}</strong><small>{currentCritique.issues.length ? "Fix bounded measurements first, then rebuild and recheck the original message in the new export pixels." : "The deterministic checks pass. Visual quality still needs the source/export comparison and whole-deck qualification."}</small></div>{currentCritique.issues.length > 0 && <div className="studio-repair-issue-list">{currentCritique.issues.slice(0, 4).map((issue) => <span key={issue.id} className={`severity-${issue.severity}`}><strong>{issue.category.replaceAll("-", " ")}</strong>{issue.message}</span>)}</div>}<button className="button secondary small" disabled={!currentCritique.autoFixableCount || sacredTitleSlide} title={sacredTitleSlide ? "The approved ORNL title composition is sacred." : currentCritique.autoFixableCount ? "Apply only collision-checked repairs supported by this exact native measurement." : "These issues need a material design or human decision."} onClick={() => onRepairObjectiveIssues(slide.slideNumber)}><MagicWand size={15} />{currentCritique.autoFixableCount ? `Fix ${currentCritique.autoFixableCount} bounded` : "No automatic fix"}</button></section>}
    {showConsistency && consistency && <section className="studio-consistency-review"><div><span className="field-label">Deck consistency review</span><strong>{consistency.issueCount ? `${consistency.issueCount} system-level difference${consistency.issueCount === 1 ? "" : "s"}` : "Shared systems are consistent"}</strong><small>{consistency.designedSlideCount} designed slides · {consistency.repeatedComponentCount} repeated components · {consistency.tableCount} tables</small></div>{consistency.issues.length ? <div className="studio-consistency-issues">{consistency.issues.map((issue) => <button key={issue.id} onClick={() => { const slideNumber = issue.slideNumbers[0]; setSelectedNumber(slideNumber); setSelectedNodeIds(issue.nodeIds.filter((id) => scene.slides.find((item) => item.slideNumber === slideNumber)?.nodes.some((node) => node.id === id))); setShowEditor(true); }}><span>{issue.category.replaceAll("-", " ")} · {issue.severity}</span><strong>{issue.message}</strong><small>{issue.recommendation}</small></button>)}</div> : <div className="inline-note"><CheckCircle size={15} />No outlier title grids, repeated component styles, or related table systems were detected in the current scene revision.</div>}</section>}
    {editorVisible && selectedNodeIds.length >= 2 && <section className="studio-arrange-toolbar panel" aria-label="Align and distribute selected Studio elements"><span><strong>Arrange {selectedNodeIds.length} elements</strong><small>Keyboard-accessible commands use the same collision-checked solver as the AI.</small></span><div>{([{"mode":"left","label":"Left"},{"mode":"center","label":"Center"},{"mode":"right","label":"Right"},{"mode":"top","label":"Top"},{"mode":"middle","label":"Middle"},{"mode":"bottom","label":"Bottom"}] as Array<{ mode: StudioConstraintRequest["mode"]; label: string }>).map((item) => <button key={item.mode} className="button ghost small" onClick={() => onArrangeSelection(slide.slideNumber, selectedNodeIds, item.mode)}>{item.label}</button>)}<button className="button ghost small" disabled={selectedNodeIds.length < 3} onClick={() => onArrangeSelection(slide.slideNumber, selectedNodeIds, "horizontal-equal-gap")}>Distribute H</button><button className="button ghost small" disabled={selectedNodeIds.length < 3} onClick={() => onArrangeSelection(slide.slideNumber, selectedNodeIds, "vertical-equal-gap")}>Distribute V</button></div></section>}
    {figureControlPanel}
    {editorVisible && selectedNode?.component && <section className="studio-component-instance panel"><div><span className="field-label">Reusable deck component</span><strong>{selectedNode.component.role.replaceAll("-", " ")}</strong><small>{compatibleComponentInstances.length} compatible {compatibleComponentInstances.length === 1 ? "instance" : "instances"} on the same light/dark surface class · exact copy and geometry stay unchanged</small></div><button className="button secondary small" disabled={selectedNode.locked || compatibleComponentInstances.length < 2} onClick={() => onPublishComponent(slide.slideNumber, selectedNode.id)}><SquaresFour size={15} />Save style + update compatible instances</button>{selectedNode.component.definitionId && <em>Following {selectedNode.component.definitionId}</em>}</section>}
    {editorVisible && selectedNode?.kind === "connector" && <StudioConnectorInspector node={selectedNode} slide={slide} onUpdate={(design) => onUpdateConnector(slide.slideNumber, selectedNode.id, design)} />}
    {editorVisible && <StudioVisualNeedsPanel slide={slide} conceptResources={conceptResources} onCreate={onCreateVisualNeed} onHold={onHoldVisualNeed} onAttach={onAttachConcept} onDetach={onDetachConcept} onReconstruct={onReconstructConcept} />}
    {editorVisible && !sacredTitleSlide && <div className="studio-semantic-recipe-shortcuts"><span><strong>Semantic layouts</strong><small>Purpose-built systems for dense technical content</small></span><button className={`button small ${slide.recipe === "ornl-title-question-diagram" ? "primary" : "secondary"}`} onClick={() => onRecompose(slide.slideNumber, "ornl-title-question-diagram")}><CirclesThreePlus size={15} />Questions + diagram</button><button className={`button small ${slide.recipe === "ornl-title-challenges-evidence" ? "primary" : "secondary"}`} onClick={() => onRecompose(slide.slideNumber, "ornl-title-challenges-evidence")}><SquaresFour size={15} />Challenges + evidence</button><button className={`button small ${slide.recipe === "ornl-title-process-flow" ? "primary" : "secondary"}`} onClick={() => onRecompose(slide.slideNumber, "ornl-title-process-flow")}><ArrowRight size={15} />Process flow</button></div>}
    <section className={`studio-shell ${editorVisible ? "editing" : "export-only"}`}>
      <aside className="studio-slide-rail"><span className="field-label">Slides</span>{scene.slides.map((item) => { const status = resultStatus(item); return <button key={item.id} className={item.slideNumber === slide?.slideNumber ? "selected" : ""} onClick={() => { setSelectedNumber(item.slideNumber); setSelectedNodeIds([]); }}><i className={`studio-result-dot ${status}`} title={status === "ready" ? "PowerPoint export result ready" : "Export result not built"} /><span>{item.slideNumber}</span><small>{item.status === "designed" ? item.recipe.replace("ornl-", "") : "source"}</small></button>; })}</aside>
      <section className="studio-stage">
        {editorVisible ? <><div className="studio-editor-warning"><Warning size={16} /><span><strong>Editing scene—not the export result.</strong> Close the scene editor and rebuild the PowerPoint result to judge the finished slide.</span></div><div className="studio-toolbar"><label>Recipe<select value={slide?.recipe ?? "source"} onChange={(event) => onRecompose(slide.slideNumber, event.target.value as StudioLayoutRecipe, slide.targetLayoutId)}><option value="source">Source geometry</option><option value="ornl-title-content">ORNL title + content</option><option value="ornl-title-two-column">ORNL two column</option><option value="ornl-title-objective-columns">ORNL objective columns</option><option value="ornl-title-steps-evidence">ORNL steps + evidence</option><option value="ornl-title-card-grid">ORNL comparison cards</option><option value="ornl-title-table">ORNL table</option><option value="ornl-title-figure-grid">ORNL figure grid</option><option value="ornl-title-labeled-figure-grid">ORNL labeled figures</option><option value="ornl-title-question-diagram">ORNL questions + diagram</option><option value="ornl-title-challenges-evidence">ORNL challenges + evidence</option><option value="ornl-title-process-flow">ORNL process flow</option><option value="template-layout">Converted template layout</option></select></label>{slide?.recipe === "template-layout" && <label>Converted ORNL layout<select value={slide.targetLayoutId ?? ""} onChange={(event) => onRecompose(slide.slideNumber, "template-layout", event.target.value)}><option value="">Choose layout…</option>{templateCatalog?.layouts.map((layout) => <option key={layout.id} value={layout.id}>{layout.name}</option>)}</select></label>}<button className="button secondary small" onClick={() => onRecompose(slide.slideNumber, recommended)}><Sparkle size={15} />Best Studio recipe</button>{recommendedTemplate && <button className="button secondary small" title={`${recommendedTemplate.status}: ${recommendedTemplate.reasons.join(" · ")}`} onClick={() => onRecompose(slide.slideNumber, "template-layout", recommendedTemplate.layoutId)}><SquaresFour size={15} />Best ORNL layout</button>}<span className="studio-mode-chip">Scene editor</span></div><StudioWebCanvas scene={scene} slide={slide} catalog={catalog} templateCatalog={templateCatalog} templateNativeBaseSource={templateNativeBaseSource} nativeSlideSource={nativeSlideSource} selectedNodeIds={selectedNodeIds} onSelectNode={(value, additive) => setSelectedNodeIds((current) => !value ? [] : additive ? current.includes(value) ? current.filter((id) => id !== value) : [...current, value] : [value])} onMoveNodes={(updates) => onMoveNodes(slide.slideNumber, updates)} /><div className="studio-canvas-caption"><span><strong>{slide.recipe.replaceAll("-", " ")}</strong> · {slide.nodes.filter((node) => node.visible).length} semantic nodes · {slide.contentCoverage.exactTextMapped ? "all source text mapped" : `${slide.contentCoverage.mappedCharacterCount}/${slide.contentCoverage.sourceCharacterCount} source characters mapped`}</span><span>{selectedNodeIds.length > 1 ? `${selectedNodeIds.length} elements selected · drag as one relationship-preserving group` : slide.contentCoverage.exactTextMapped ? "Shift-click to select a group · rebuild to see export pixels." : "Export is held until grouped or unsupported text is atomized."}</span></div></> : <><div className="studio-toolbar studio-export-toolbar"><span><strong>Slide {slide.slideNumber} export result</strong><small>{sacredTitleSlide ? "Approved ORNL template composition · source preserved" : needsDesignedExport ? currentFreshPreview ? "Current designed revision" : "Design changed · export render required" : "Source preserved"}</small></span><span className="studio-mode-chip">{sacredTitleSlide ? "Brand locked" : "PowerPoint native"}</span></div><StudioExportResult slideNumber={slide.slideNumber} source={exportResultSource} preview={currentFreshPreview} designed={needsDesignedExport} templateLayout={slide.recipe === "template-layout"} onBuild={() => onPreviewFresh(slide.slideNumber)} /></>}
      </section>
      {!editorVisible && <StudioExportInspector slide={slide} preview={currentFreshPreview} designed={needsDesignedExport} />}
      <aside className="studio-inspector"><span className="field-label">Design inspector</span>{selectedNode ? <><strong>{selectedNode.name}</strong><small>{selectedNode.kind} · {selectedNode.role} · {selectedNode.locked ? "locked native" : "editable"}</small><div className="studio-inspector-grid">{(["x", "y", "width", "height"] as const).map((field) => <label key={`${selectedNode.id}-${field}`}><span>{field === "width" ? "W" : field === "height" ? "H" : field.toUpperCase()}</span><input type="number" min={field === "width" || field === "height" ? .1 : 0} max={20} step={.01} defaultValue={(selectedNode.frame[field] / 914400).toFixed(2)} disabled={selectedNode.locked} onBlur={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) onMoveNode(slide.slideNumber, selectedNode.id, { ...selectedNode.frame, [field]: value * 914400 }); }} /><small>in</small></label>)}</div>{selectedNode.kind === "text" && <div className="studio-type-controls"><label><span>Size</span><input type="number" min={10} max={60} step={.25} defaultValue={selectedNode.style.fontSizePt} disabled={selectedNode.locked} onBlur={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { fontSizePt: Number(event.currentTarget.value) })} /><small>pt</small></label><label><span>Weight</span><select value={selectedNode.style.fontWeight} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { fontWeight: Number(event.currentTarget.value) as 400 | 600 | 700 })}><option value="400">Regular</option><option value="600">Semibold</option><option value="700">Bold</option></select></label><label><span>Align</span><select value={selectedNode.style.textAlign} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { textAlign: event.currentTarget.value as "left" | "center" | "right" })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label><label><span>Color</span><input type="color" value={selectedNode.style.color} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { color: event.currentTarget.value })} /></label></div>}{selectedNode.kind === "table" && selectedNode.table && <StudioTableInspector node={selectedNode} exemplars={scene.tableLibrary ?? []} continuationPlan={selectedTableContinuationPlan} compatibleTableCount={compatibleTableCount} onDesign={(patch) => onUpdateTableDesign(slide.slideNumber, selectedNode.id, patch)} onResizeColumn={(column, width) => onResizeTableColumn(slide.slideNumber, selectedNode.id, column, width)} onResizeRow={(row, height) => onResizeTableRow(slide.slideNumber, selectedNode.id, row, height)} onCell={(cellId, patch) => onUpdateTableCell(slide.slideNumber, selectedNode.id, cellId, patch)} onPublishExemplar={() => onPublishTableExemplar(slide.slideNumber, selectedNode.id)} onApplyExemplar={onApplyTableExemplar} onPlanContinuation={(maximumBodyRowsPerSlide) => onPlanTableContinuation(slide.slideNumber, selectedNode.id, maximumBodyRowsPerSlide)} onClearContinuation={() => onClearTableContinuation(slide.slideNumber, selectedNode.id)} />}{selectedNode.exactContent && <div className="inline-note"><ShieldCheck size={15} />Content remains bound to its source hash.</div>}{selectedFigureTreatment && <section className="studio-figure-treatment"><span className="field-label">Figure treatment</span><div className="studio-treatment-status"><strong>{selectedFigureTreatment.mode.replaceAll("-", " ")}</strong><span>{selectedFigureTreatment.verificationStatus.replaceAll("-", " ")}</span></div><p>{selectedFigureTreatment.intentSummary}</p><small>{selectedFigureTreatment.rationale}</small><ul>{selectedFigureTreatment.invariants.slice(0, 3).map((item) => <li key={item}>{item}</li>)}</ul><div className="inline-note"><ShieldCheck size={15} />{selectedFigureTreatment.mode === "preserve-as-unit" || selectedFigureTreatment.mode === "preserve-and-frame" ? "Original technical pixels stay exact; Studio designs around the evidence unit." : "The original remains visible until a verified replacement passes the intent review."}</div></section>}</> : <><p>Select an element to inspect its semantic role, CSS-computed frame, and PowerPoint binding.</p><div className="inline-note"><Info size={15} />This is the new design authority. Native PowerPoint pixels remain the final export authority.</div></>}{(slide.conceptReferences?.length ?? 0) > 0 && <section className="studio-concept-references"><span className="field-label">Concept references</span>{slide.conceptReferences?.map((reference) => <article key={reference.id}><div><Sparkle size={15} /><strong>{reference.origin === "imagegen" ? "Image Gen concept" : "Visual concept"}</strong><span>Concept only</span></div><p>{reference.blueprint.summary}</p><small>Follow: {reference.approvedInfluences.join(" · ")}</small><em>Text, logos, data, claims, and technical details remain untrusted.</em></article>)}</section>}<div className="studio-source-reference"><span className="field-label">PowerPoint source reference</span><span className="studio-source-thumb">{catalogSlide ? <SlideDesignCanvas nativeRender={nativeRender} slideNumber={slide.slideNumber} catalog={catalog} layout={catalogSlide} label={`PowerPoint source slide ${slide.slideNumber}`} /> : <span className="proposal-preview-wait">Preparing…</span>}</span><small>{nativeSlideSource ? "PowerPoint-native · read only" : "Structural fallback · read only"}</small></div>{currentFreshPreview && <div className="studio-source-reference fresh"><span className="field-label">Fresh-composition PowerPoint{currentFreshPreview.slideCount > 1 ? ` · ${currentFreshPreview.slideCount} slides` : ""}</span><div className="studio-source-thumb-stack">{currentFreshPreview.nativeRender?.status === "ready" ? currentFreshPreview.nativeRender.slides.map((rendered, index) => <span className="studio-source-thumb" key={rendered.number}><img className="native-slide-render" src={`data:${rendered.mimeType};base64,${bytesToBase64(bytesFrom(rendered.bytes))}`} width={rendered.width} height={rendered.height} alt={`PowerPoint-native fresh composition ${index + 1} of ${currentFreshPreview.slideCount} for source slide ${slide.slideNumber}`} /></span>) : <span className="proposal-preview-wait">Native preview unavailable</span>}</div><small>{freshNativeSource ? `PowerPoint-native · ${currentFreshPreview.textNodeCount} text · ${currentFreshPreview.tableCount} table · ${currentFreshPreview.imageCount} image` : "Fresh PPTX built; native visual acceptance is still unavailable."}</small></div>}</aside>
    </section>
  </div>;
}

type DesignFilter = "all" | TemplateLayoutPreview["category"];

function DesignsView({ catalog, installedAt, loading, nativeRender, nativeLoading, onInstall }: { catalog?: TemplateCatalog; installedAt?: string; loading: boolean; nativeRender?: NativeRenderResult; nativeLoading: boolean; onInstall: () => void }) {
  const [filter, setFilter] = useState<DesignFilter>("all");
  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState<string>();
  useEffect(() => {
    if (!catalog?.layouts.length) return;
    if (!catalog.layouts.some((layout) => layout.id === selectedId)) setSelectedId(catalog.layouts[0].id);
  }, [catalog, selectedId]);
  if (!catalog) return <div className="view-stack designs-empty-view"><header className="view-header compact"><div><p className="eyebrow">Authorized template library</p><h1>Slide designs</h1><p>Browse the real masters and layouts available for composition and reflow.</p></div></header><section className="designs-empty"><span className="designs-empty-icon"><SquaresFour size={34} weight="light" /></span><h2>Install an authorized PowerPoint template</h2><p>Presentation Studio will read its master, layouts, media, theme, and placeholder geometry locally. The source template stays outside Git and is never uploaded.</p><button className="button primary large" disabled={loading} onClick={onInstall}><UploadSimple size={18} />{loading ? "Reading template…" : "Choose POTX or PPTX"}</button></section></div>;

  const normalizedQuery = query.trim().toLowerCase();
  const visibleLayouts = catalog.layouts.filter((layout) => (filter === "all" || layout.category === filter) && (!normalizedQuery || layout.name.toLowerCase().includes(normalizedQuery) || layout.placeholderTypes.some((type) => type.toLowerCase().includes(normalizedQuery)) || layout.semantic?.intent.includes(normalizedQuery) || layout.semantic?.summary.toLowerCase().includes(normalizedQuery)));
  const selected = catalog.layouts.find((layout) => layout.id === selectedId) ?? catalog.layouts[0];
  const selectedNumber = Math.max(1, catalog.layouts.findIndex((layout) => layout.id === selected.id) + 1);
  const categoryCounts = catalog.layouts.reduce<Record<string, number>>((counts, layout) => ({ ...counts, [layout.category]: (counts[layout.category] ?? 0) + 1 }), {});
  return <div className="view-stack designs-view">
    <header className="view-header compact"><div><p className="eyebrow">Authorized template library</p><h1>Slide designs</h1><p>Every preview is derived locally from the installed PowerPoint master and layout catalog.</p></div><button className="button secondary" disabled={loading} onClick={onInstall}><ArrowsClockwise size={17} />{loading ? "Reading…" : "Install or update"}</button></header>
    <section className="template-pack-bar"><span className="template-pack-mark"><ShieldCheck size={21} /></span><div><span className="field-label">Active local Template Pack</span><strong>{catalog.name}</strong><small>{catalog.masterCount} master{catalog.masterCount === 1 ? "" : "s"} · {catalog.layouts.length} layouts · SHA-256 {catalog.sha256.slice(0, 12)}…{installedAt ? ` · installed ${new Date(installedAt).toLocaleDateString()}` : " · browser session"}</small></div><span className={`template-local-state ${nativeRender?.status === "ready" ? "native" : ""}`}><span />{nativeLoading ? "Rendering in PowerPoint…" : nativeRender?.status === "ready" ? "PowerPoint-native" : "Local structural view"}</span></section>
    <section className="design-feature">
      <div className="design-feature-preview"><NativeTemplateLayoutCanvas catalog={catalog} layout={selected} layoutNumber={selectedNumber} nativeRender={nativeRender} label={`Selected design: ${selected.name}`} showLabels /></div>
      <div className="design-feature-copy"><span className="field-label">Selected design</span><h2>{selected.name}</h2><p>{nativeRender?.status === "ready" ? "The base artwork below is an authoritative Microsoft PowerPoint render of the installed layout. Studio overlays its semantic editable regions for human and AI composition." : "This local structural view combines the actual template geometry, artwork, and semantic editable regions while the PowerPoint-native preview is unavailable."}</p><dl><div><dt>Design intent</dt><dd>{selected.semantic?.intent ?? selected.category}</dd></div><div><dt>Best fit</dt><dd>{selected.semantic?.summary ?? (selected.placeholderTypes.length ? selected.placeholderTypes.join(" · ") : "Freeform / template furniture")}</dd></div><div><dt>Content constraints</dt><dd>{selected.semantic ? `${selected.semantic.constraints.maxTitleLines} title lines max · ${selected.semantic.capabilities.bodySlots ? `${selected.semantic.constraints.minimumBodyFontPt} pt body minimum` : `${selected.semantic.constraints.minimumCaptionFontPt} pt label minimum`} · ${selected.semantic.constraints.preferredBodyDensity} density` : "Template-defined"}</dd></div><div><dt>PowerPoint structure</dt><dd>Master, custom layout, theme, and editable slots preserved</dd></div></dl><div className="inline-note design-note"><Info size={16} />Green dashed regions are Studio guides, not template artwork. Selecting a design changes nothing; AI recommendations use these same slots and constraints.</div></div>
    </section>
    <section className="design-browser panel">
      <div className="design-toolbar"><div className="design-filters" aria-label="Filter slide designs">{(["all", "title", "content", "image", "conclusion", "other"] as DesignFilter[]).map((value) => <button key={value} className={filter === value ? "active" : ""} onClick={() => setFilter(value)}>{value === "all" ? `All ${catalog.layouts.length}` : `${value} ${categoryCounts[value] ?? 0}`}</button>)}</div><label className="design-search"><MagnifyingGlass size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search layouts or placeholders" aria-label="Search slide designs" /></label></div>
      <div className="design-grid">{visibleLayouts.map((layout) => { const layoutNumber = catalog.layouts.findIndex((candidate) => candidate.id === layout.id) + 1; return <button key={layout.id} className={`design-card ${selected.id === layout.id ? "selected" : ""}`} onClick={() => setSelectedId(layout.id)} aria-pressed={selected.id === layout.id}><span className="design-thumb"><NativeTemplateLayoutCanvas catalog={catalog} layout={layout} layoutNumber={layoutNumber} nativeRender={nativeRender} label={`${layout.name} layout`} /></span><span className="design-card-meta"><span>{layoutNumber}</span><span><strong>{layout.name}</strong><small>{layout.semantic?.summary ?? (layout.placeholderTypes.length ? layout.placeholderTypes.join(" · ") : layout.category)}</small></span></span></button>; })}</div>
      {visibleLayouts.length === 0 && <div className="design-no-results"><MagnifyingGlass size={24} /><strong>No matching designs</strong><span>Try another category or search term.</span></div>}
    </section>
  </div>;
}

function RulesView({ deck, exemplarCount }: { deck?: DeckJob; exemplarCount: number }) {
  const ornlReady = deck?.targetTemplateId === "ornl-16x9-v1" && Boolean(deck.targetTemplateConfirmedAt);
  const standard = PRESENTATION_DESIGN_STANDARD;
  const table = standard.tableProfile;
  return <div className="view-stack"><header className="view-header compact"><div><p className="eyebrow">Presentation Design Standard</p><h1>Fast defaults, deterministic exceptions</h1><p>One versioned ruleset drives the app, preflight, cleanup engine, reports, tests, and MCP contract.</p></div><span className="standard-version">v{standard.version}</span></header>
    <section className="standard-summary panel"><div className="standard-summary-mark"><ShieldCheck size={25} /></div><div><span className="field-label">Adopted defaults</span><h2>{designStandardSummary()}</h2><p>Current authorized ORNL Template Pack · source remains read-only · editable native PowerPoint output · routine design choices proceed without per-slide approvals.</p></div><span className={ornlReady ? "ready" : "waiting"}>{ornlReady ? "Ready" : "Waiting for target"}</span></section>
    <section className="panel table-standard"><div className="panel-heading"><div><h2>Native table profile</h2><p>{table.id} · product fallback when the Template Pack or approved exemplar has no stronger rule.</p></div><span className="quiet-label">Native and editable</span></div><div className="table-token-grid"><div><span>Typography</span><strong>{table.fontFamily} · {table.header.fontSizePt} pt header · {table.body.fontSizePt} pt body</strong><small>No silent shrink below {table.body.minimumFontSizePt} pt</small></div><div><span>Header</span><strong>Hale Navy · Polar · bold</strong><small>Vertically centered with semantic alignment</small></div><div><span>Body</span><strong>Dark Matter · Polar / Graphite banding</strong><small>Meaning-bearing source color is preserved</small></div><div><span>Cell padding</span><strong>{table.cellPaddingPt.left} pt horizontal · {table.cellPaddingPt.top} pt vertical</strong><small>Measured again using export font metrics</small></div><div><span>Strokes</span><strong>No outer box · {table.strokes.horizontal.widthPt} pt Graphite rows</strong><small>Vertical rules only when comprehension requires them</small></div><div><span>Overflow</span><strong>Reflow, widen, or continue</strong><small>Never hide, drop, flatten, or silently miniaturize</small></div></div></section>
    <div className="rule-grid"><RuleCard title="Legacy font normalization" scope="Font family only" status={ornlReady ? "Ready" : "Waiting for template"} detail="Maps direct Century Gothic and Arial slide markup to Aptos. Symbol fonts, theme tokens, text strings, object identity, and geometry remain untouched." /><RuleCard title="Exact-content guard" scope="Every exported slide" status="Always on" detail="Compares the visible-text hash and slide count before and after cleanup. The export is rejected if either changes." /><RuleCard title="Advanced content hold" scope="Deck safety" status="Always on" detail="Prevents automated cleanup when macros, embedded OLE objects, or external relationships require a human decision." /><RuleCard title="Approved table exemplar" scope="Tables only" status={exemplarCount > 0 ? `${exemplarCount} registered` : "Product fallback active"} detail="A registered exemplar overrides compatible table tokens without copying cell content. Otherwise the versioned minimal-rule profile is used." /></div></div>;
}

function RuleCard({ title, scope, status, detail }: { title: string; scope: string; status: string; detail: string }) {
  return <article className="rule-card"><div className="rule-top"><ListChecks size={22} /><span>{status}</span></div><h2>{title}</h2><p>{detail}</p><small>{scope}</small></article>;
}

function initialReviewSlide(proposal: CleanupProposal | undefined): number {
  if (!proposal) return 1;
  return proposal.designReview?.evidence?.slideNumber
    ?? proposal.slideDispositions.find((item) => item.status === "change-proposed" || item.changeIds.length > 0)?.slideNumber
    ?? proposal.slideDispositions.find((item) => item.status === "needs-review")?.slideNumber
    ?? 1;
}

function ReviewView({ deck, projectUpdatedAt, currentCatalog, proposalCatalog, currentNativeRender, proposalNativeRender, previewLoading, threads, onToggle, onReviewSlide, onRequestChanges, onDeleteThread, onOpenSlide, onReject, onApply, onExport }: { deck?: DeckJob; projectUpdatedAt: string; currentCatalog?: SlideRenderCatalog; proposalCatalog?: SlideRenderCatalog; currentNativeRender?: NativeRenderResult; proposalNativeRender?: NativeRenderResult; previewLoading: boolean; threads: DesignThread[]; onToggle: (id: string) => void; onReviewSlide: (slideNumber: number, decision: "approved" | "changes-requested") => void; onRequestChanges: (slideNumber: number, comment: string, submit: boolean) => void; onDeleteThread: (threadId: string) => void; onOpenSlide: (slideNumber: number, mode: SlideWorkspaceRequest["mode"]) => void; onReject: () => void; onApply: (slideNumber: number) => void; onExport: () => void }) {
  const proposal = deck?.proposal;
  const [selectedNumber, setSelectedNumber] = useState(1);
  const [pixelComparison, setPixelComparison] = useState<PixelComparisonMetrics>();
  const [requestOpen, setRequestOpen] = useState(false);
  const [requestComment, setRequestComment] = useState("");
  const [overviewRepresentation, setOverviewRepresentation] = useState<"current" | "proposal">("proposal");
  useEffect(() => { setSelectedNumber(initialReviewSlide(proposal)); }, [deck?.id, proposal?.id, proposal?.designReview?.reviewedAt]);
  useEffect(() => {
    const currentSlide = currentNativeRender?.status === "ready" ? currentNativeRender.slides.find((slide) => slide.number === selectedNumber) : undefined;
    const proposalSlide = proposalNativeRender?.status === "ready" ? proposalNativeRender.slides.find((slide) => slide.number === selectedNumber) : undefined;
    if (!currentSlide || !proposalSlide) { setPixelComparison(undefined); return; }
    let canceled = false;
    void compareNativeSlideRenders(currentSlide, proposalSlide).then((comparison) => { if (!canceled) setPixelComparison(comparison.metrics); }).catch(() => { if (!canceled) setPixelComparison(undefined); });
    return () => { canceled = true; };
  }, [currentNativeRender, proposalNativeRender, selectedNumber]);
  useEffect(() => { setRequestOpen(false); setRequestComment(""); }, [selectedNumber, proposal?.id]);
  if (!deck || !proposal) return <NoSelection message="Stage a cleanup proposal from Deck audit to review exact changes here." icon={MagicWand} />;
  const stale = proposal.baseUpdatedAt !== projectUpdatedAt && proposal.status === "pending";
  const disposition = proposal.slideDispositions.find((item) => item.slideNumber === selectedNumber);
  const currentSlide = currentCatalog?.slides.find((slide) => slide.number === selectedNumber);
  const proposalSlide = proposalCatalog?.slides.find((slide) => slide.number === selectedNumber);
  const changedCount = proposal.slideDispositions.filter((item) => item.status === "change-proposed" || item.changeIds.length > 0).length;
  const approvedCount = proposal.slideDispositions.filter((item) => item.status === "approved-as-is").length;
  const reviewCount = proposal.slideDispositions.filter((item) => item.status === "needs-review").length;
  const tableCount = proposal.changes.filter((change) => change.kind === "table-style" && change.selected).reduce((sum, change) => sum + (change.tableIds?.length ?? 0), 0);
  const alignmentCount = proposal.changes.filter((change) => change.kind === "alignment" && change.selected).reduce((sum, change) => sum + (change.alignmentRepairs?.length ?? 0), 0);
  const geometryCommands = proposal.changes.filter((change) => change.kind === "geometry" && change.selected).flatMap((change) => change.geometryCommands ?? []);
  const geometryCount = geometryCommands.length;
  const geometrySlideCount = new Set(geometryCommands.map((command) => command.slideNumber)).size;
  const geometryWarningCount = geometryCommands.reduce((sum, command) => sum + (command.validation?.warnings.length ?? 0), 0);
  const layoutCommands = proposal.changes.filter((change) => change.kind === "layout-remap" && change.selected).flatMap((change) => change.layoutCommands ?? []);
  const layoutCount = layoutCommands.length;
  const textStyleCount = proposal.changes.filter((change) => change.kind === "text-style" && change.selected).reduce((sum, change) => sum + (change.textStyleCommands?.length ?? 0), 0);
  const decorationCount = proposal.changes.filter((change) => change.kind === "decoration" && change.selected).reduce((sum, change) => sum + (change.decorationCommands?.length ?? 0), 0);
  const layoutExceptions = proposal.layoutExceptions ?? [];
  const selectedLayoutExceptions = layoutExceptions.filter((exception) => exception.slideNumber === selectedNumber);
  const selectedReview = (proposal.slideReviews ?? []).find((review) => review.slideNumber === selectedNumber);
  const approvedSlideNumbers = new Set((proposal.slideReviews ?? []).filter((review) => review.decision === "approved").map((review) => review.slideNumber));
  const reviewedSlideCount = approvedSlideNumbers.size;
  const remainingSlideCount = Math.max(0, proposal.slideDispositions.length - reviewedSlideCount);
  const selectedChangeIds = new Set(proposal.changes.filter((change) => change.selected).map((change) => change.id));
  const selectedHasChanges = Boolean(disposition?.changeIds.some((id) => selectedChangeIds.has(id)));
  const unresolvedRequestCount = (proposal.slideReviews ?? []).filter((review) => review.decision === "changes-requested").length;
  const canApprovePlan = !stale && proposal.changes.some((change) => change.selected) && unresolvedRequestCount === 0;
  const selectedSubmittedThreads = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === selectedNumber && ["submitted", "needs-reanchor"].includes(thread.status));
  const overviewCatalog = overviewRepresentation === "proposal" ? proposalCatalog : currentCatalog;
  const overviewNativeRender = overviewRepresentation === "proposal" ? proposalNativeRender : currentNativeRender;
  const overviewIsNative = overviewNativeRender?.status === "ready";
  const reviewSlideNumbers = proposal.slideDispositions.map((item) => item.slideNumber);
  function nextUnapprovedSlide() {
    return reviewSlideNumbers.find((number) => number > selectedNumber && !approvedSlideNumbers.has(number)) ?? reviewSlideNumbers.find((number) => !approvedSlideNumbers.has(number));
  }
  function approveSelectedSlide() {
    onReviewSlide(selectedNumber, "approved");
    const next = nextUnapprovedSlide();
    if (next && next !== selectedNumber) setSelectedNumber(next);
  }
  function submitRequest(submit: boolean) {
    if (!requestComment.trim()) return;
    onRequestChanges(selectedNumber, requestComment.trim(), submit);
    setRequestComment("");
    setRequestOpen(false);
  }
  return <div className="view-stack review-workspace"><header className="view-header compact review-header"><div><p className="eyebrow">Before / after design review</p><h1>{proposal.summary}</h1><p>{deck.name} · Proposal {proposal.id.slice(0, 8)} · {proposal.mode.replaceAll("-", " ")}</p></div><div className="review-header-actions"><span className={`proposal-state ${proposal.status}`}>{proposal.status}</span>{proposal.status === "pending" && <button className="button primary" disabled={!canApprovePlan} title={unresolvedRequestCount ? "Resolve requested changes before approving the plan." : undefined} onClick={() => onApply(selectedNumber)}><CheckCircle size={18} />Approve all &amp; continue</button>}{proposal.status === "applied" && <button className="button secondary" onClick={() => onOpenSlide(selectedNumber, "edit")}><Crosshair size={17} />Continue editing</button>}</div></header>
    {stale && <div className="warning-banner"><Warning size={18} /><div><strong>This proposal is stale.</strong><span>The project changed after it was staged. Restage before applying.</span></div></div>}
    {proposal.designReview && <div className="warning-banner"><Warning size={18} /><div><strong>{proposal.designReview.actor === "ai" ? "AI rejected this draft after native review." : "This draft was rejected in review."}</strong><span>{proposal.designReview.rationale}{proposal.designReview.evidence ? ` PowerPoint-native comparison changed ${(proposal.designReview.evidence.changedPixelRatio * 100).toFixed(2)}% of pixels on slide ${proposal.designReview.evidence.slideNumber}.` : ""}</span></div></div>}
    {proposal.visualIteration?.history.length ? (() => { const critique = proposal.visualIteration.history.at(-1)!; return <div className={critique.verdict === "better" ? "design-decision-banner" : "warning-banner"}><ArrowsClockwise size={18} /><div><strong>AI visual iteration {critique.attempt}/{proposal.visualIteration.maxAttempts}: {critique.verdict === "better" ? "better—ready for human review" : critique.verdict === "revise" ? "revise again" : "rejected"}</strong><span>{critique.rationale} {critique.metrics.improvements.length ? `Improved: ${critique.metrics.improvements.join(", ")}.` : ""} {critique.metrics.regressions.length ? `Regressed: ${critique.metrics.regressions.join(", ")}.` : ""}</span></div></div>; })() : null}
    {proposal.designDecision && <section className="design-decision-banner"><SquaresFour size={20} /><div><span className="field-label">Semantic recomposition decision</span><strong>{proposal.designDecision.targetLayoutName} · {proposal.designDecision.compatibilityScore}/100 {proposal.designDecision.compatibilityStatus}</strong><small>{proposal.designDecision.rationale} {proposal.designDecision.application === "cloned-native-layout" ? "Presentation Studio first reuses an exact approved native layout already in the deck; otherwise it clones the guarded master/layout/theme/media graph. Placeholder identities are mapped to the target before PowerPoint-native review." : "The source master/layout remains intact in this slice; bound objects are staged into approved semantic zones."}</small></div><span>{proposal.designDecision.application === "cloned-native-layout" ? "Native layout" : `${proposal.designDecision.bindingCount} bindings`}</span></section>}
    <section className="review-progress panel"><div><span className="field-label">Slide review progress</span><strong>{reviewedSlideCount} of {proposal.slideDispositions.length} approved</strong><small>{unresolvedRequestCount > 0 ? `${unresolvedRequestCount} change request${unresolvedRequestCount === 1 ? "" : "s"} waiting for revision.` : remainingSlideCount === 0 ? "Every slide has been reviewed. Approve the plan when ready." : `${remainingSlideCount} slide${remainingSlideCount === 1 ? "" : "s"} remaining · use Approve & next for the fastest pass.`}</small></div><div className="review-progress-track" aria-label={`${reviewedSlideCount} of ${proposal.slideDispositions.length} slides approved`}><span style={{ width: `${proposal.slideDispositions.length ? reviewedSlideCount / proposal.slideDispositions.length * 100 : 0}%` }} /></div><button className="button ghost small" onClick={() => onOpenSlide(selectedNumber, "edit")}><Crosshair size={16} />Edit slide</button><button className="button ghost small" onClick={() => onOpenSlide(selectedNumber, "comment")}><ChatCircleDots size={16} />Point comment</button></section>
    <details className="panel review-deck-overview"><summary><span><SquaresFour size={18} /><span><strong>Visual deck overview</strong><small>Review hierarchy, rhythm, and consistency across every slide</small></span></span><span>{overviewIsNative ? "PowerPoint-native" : "Approximate fallback"}<CaretRight size={16} /></span></summary><div className="review-overview-toolbar"><div><span className="field-label">Representation</span><div className="segmented-control"><button type="button" className={overviewRepresentation === "current" ? "active" : ""} onClick={() => setOverviewRepresentation("current")}>Current</button><button type="button" className={overviewRepresentation === "proposal" ? "active" : ""} onClick={() => setOverviewRepresentation("proposal")}>Proposal</button></div></div><p>{overviewIsNative ? "Every thumbnail below is an authoritative Microsoft PowerPoint render of the exact selected revision." : "These thumbnails use the diagnostic OOXML reconstruction. Use them for navigation only; final visual acceptance requires PowerPoint-native pixels."}</p></div>{overviewCatalog ? <div className="review-overview-grid">{overviewCatalog.slides.map((slide) => { const review = (proposal.slideReviews ?? []).find((item) => item.slideNumber === slide.number); const slideDisposition = proposal.slideDispositions.find((item) => item.slideNumber === slide.number); const commentCount = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === slide.number && ["submitted", "needs-reanchor"].includes(thread.status)).length; const status = review?.decision === "approved" ? "approved" : review?.decision === "changes-requested" || commentCount > 0 ? "revise" : slideDisposition?.status === "change-proposed" || slideDisposition?.changeIds.length ? "changed" : slideDisposition?.status === "needs-review" ? "review" : "as-is"; return <button type="button" key={slide.id} className={`review-overview-slide status-${status} ${slide.number === selectedNumber ? "selected" : ""}`} onClick={() => setSelectedNumber(slide.number)} aria-label={`Open slide ${slide.number}; ${status.replace("-", " ")}${commentCount ? `; ${commentCount} open comment${commentCount === 1 ? "" : "s"}` : ""}`}><span className="review-overview-canvas"><SlideDesignCanvas nativeRender={overviewNativeRender} slideNumber={slide.number} catalog={overviewCatalog} layout={slide} label={`${overviewRepresentation} overview slide ${slide.number}`} />{commentCount > 0 && <i><ChatCircleDots size={11} />{commentCount}</i>}</span><span className="review-overview-meta"><b>{slide.number}</b><small>{status.replace("-", " ")}</small></span></button>; })}</div> : <div className="proposal-preview-wait"><ArrowsClockwise className="spinner" size={22} />Preparing deck overview…</div>}</details>
    <div className="metric-strip review-metrics">{proposal.mode === "slide-geometry" ? <><Metric value={reviewedSlideCount} label="Slides approved" /><Metric value={geometrySlideCount} label="Slides in transaction" /><Metric value={geometryCount} label="Objects adjusted" /><Metric value={geometryCommands.filter((command) => (command.validation?.warnings.length ?? 0) === 0).length} label="Clean validations" /><Metric value={geometryWarningCount} label="Review warnings" /><Metric value={0} label="Content edits" /></> : proposal.mode === "slide-reflow" ? <><Metric value={reviewedSlideCount} label="Slides approved" /><Metric value={geometryCount} label="Objects placed" /><Metric value={textStyleCount} label="Text styles" /><Metric value={decorationCount} label="Brand vectors" /><Metric value={selectedSubmittedThreads.length} label="Open AI notes" /><Metric value={0} label="Content edits" /></> : <><Metric value={reviewedSlideCount} label="Slides approved" /><Metric value={changedCount} label="With changes" /><Metric value={approvedCount} label="Approved as-is" /><Metric value={reviewCount} label="Need review" /><Metric value={alignmentCount + geometryCount} label="Layout edits" /><Metric value={tableCount} label="Tables normalized" /></>}</div>
    <section className="proposal-compare panel"><div className="slide-review-toolbar review-decision-toolbar"><div><span><strong>Slide {selectedNumber}</strong><small className={`disposition-label ${selectedReview?.decision === "approved" ? "review-approved" : selectedReview?.decision === "changes-requested" ? "changes-requested" : disposition?.status ?? "approved-as-is"}`}>{selectedReview?.decision === "approved" ? "approved" : selectedReview?.decision === "changes-requested" ? "changes requested" : disposition?.status.replaceAll("-", " ") ?? "reviewed"}</small></span></div><div><button className="button ghost small" disabled={selectedNumber <= 1} onClick={() => setSelectedNumber((value) => value - 1)}><ArrowLeft size={15} />Previous</button><button className="button ghost small" disabled={selectedNumber >= (deck.audit?.slideCount ?? 1)} onClick={() => setSelectedNumber((value) => value + 1)}>Next<ArrowRight size={15} /></button>{proposal.status === "pending" && <><button className="button secondary small" onClick={() => setRequestOpen((value) => !value)}><ChatCircleDots size={16} />Request changes</button><button className="button primary small" onClick={approveSelectedSlide}><Check size={16} />{selectedReview?.decision === "approved" ? "Approved · next" : "Approve & next"}</button></>}</div></div>
      {requestOpen && proposal.status === "pending" && <div className="review-request-composer"><div><ChatCircleDots size={20} /><span><strong>Request changes on slide {selectedNumber}</strong><small>Describe the adjustment here, or choose Point to exact area to anchor the comment on the slide.</small></span></div><textarea autoFocus value={requestComment} maxLength={4000} onChange={(event) => setRequestComment(event.target.value)} placeholder="Example: Align the diagram labels, increase the spacing above the table, and keep all wording unchanged." /><div><button className="button ghost small" onClick={() => setRequestOpen(false)}>Cancel</button><button className="button secondary small" onClick={() => onOpenSlide(selectedNumber, "comment")}><Crosshair size={15} />Point to exact area</button><button className="button ghost small" disabled={!requestComment.trim()} onClick={() => submitRequest(false)}>Save note</button><button className="button primary small" disabled={!requestComment.trim()} onClick={() => submitRequest(true)}><PaperPlaneTilt size={15} />Submit to AI</button></div></div>}
      {selectedSubmittedThreads.length > 0 && <div className="review-open-comments"><div className="review-open-comments-heading"><ChatCircleDots size={17} /><span><strong>Open comments on slide {selectedNumber}</strong><small>Delete feedback that no longer applies, or leave it for the AI to address explicitly.</small></span></div><div>{selectedSubmittedThreads.map((thread, index) => <article key={thread.id}><span>{index + 1}</span><p>{thread.comment}</p><button type="button" onClick={() => onDeleteThread(thread.id)} aria-label={`Delete open comment ${index + 1}`} title="Delete comment"><Trash size={15} /></button></article>)}</div></div>}
      <div className="proposal-compare-body"><div className="proposal-compare-canvases"><div className="proposal-canvas-pane"><div><strong>Current</strong><small>{currentNativeRender?.status === "ready" ? "PowerPoint-native · embedded source · read only" : "Approximate fallback · embedded source"}</small></div><span className="proposal-slide-canvas">{currentCatalog && currentSlide ? <SlideDesignCanvas nativeRender={currentNativeRender} slideNumber={selectedNumber} catalog={currentCatalog} layout={currentSlide} label={`Current slide ${selectedNumber}`} /> : <span className="proposal-preview-wait"><ArrowsClockwise className="spinner" size={22} />Rendering current design…</span>}</span></div><div className="proposal-canvas-pane proposed"><div><strong>Proposal</strong><small>{proposalNativeRender?.status === "ready" ? proposal.status === "rejected" ? "PowerPoint-native · rejected draft evidence" : selectedHasChanges ? "PowerPoint-native · selected reversible changes" : "PowerPoint-native · unchanged on this slide" : proposal.status === "pending" ? "Approximate fallback · selected changes" : proposal.status === "applied" ? "Accepted changes · export-ready" : "Rejected proposal preview"}</small></div><span className="proposal-slide-canvas">{proposalCatalog && proposalSlide ? <SlideDesignCanvas nativeRender={proposalNativeRender} slideNumber={selectedNumber} catalog={proposalCatalog} layout={proposalSlide} label={`Proposed slide ${selectedNumber}`} /> : <span className="proposal-preview-wait">{previewLoading ? <><ArrowsClockwise className="spinner" size={22} />Rendering proposal…</> : <><Info size={22} />Select at least one supported change</>}</span>}</span></div></div>
        <aside className="proposal-slide-rail"><span className="field-label">{proposal.mode === "slide-geometry" ? "Transaction scope" : "Deck-wide disposition"}</span><div className="disposition-reasons">{disposition?.reasons.map((reason) => <p key={reason}>{reason}</p>)}</div><div className="disposition-list">{proposal.slideDispositions.map((item) => { const review = (proposal.slideReviews ?? []).find((candidate) => candidate.slideNumber === item.slideNumber); return <button key={item.slideNumber} className={`${review?.decision === "approved" ? "review-approved" : review?.decision === "changes-requested" ? "changes-requested" : item.status} ${item.slideNumber === selectedNumber ? "selected" : ""}`} onClick={() => setSelectedNumber(item.slideNumber)}><span>{item.slideNumber}</span><small>{review?.decision === "approved" ? "Approved" : review?.decision === "changes-requested" ? "Revise" : item.status === "change-proposed" ? "Changed" : item.status === "needs-review" ? "Review" : proposal.mode === "slide-geometry" ? "Outside" : "As-is"}</small></button>; })}</div></aside></div>
      <div className={`slide-representation-note ${currentNativeRender?.status === "ready" && proposalNativeRender?.status === "ready" ? "native" : "fallback"}`}><ShieldCheck size={15} /><span><strong>Revision-bound comparison</strong> {currentNativeRender?.status === "ready" && proposalNativeRender?.status === "ready" ? "Both views are authoritative Microsoft PowerPoint renders of their exact PPTX revisions." : "One or both views use the diagnostic OOXML approximation. Do not make final visual acceptance decisions until PowerPoint-native renders are available."}</span></div>
      {pixelComparison && <div className={`native-difference-bar ${pixelComparison.exactPixelMatch ? "unchanged" : "changed"}`}><Monitor size={16} /><span><strong>{pixelComparison.exactPixelMatch ? "No native visual change detected" : `${(pixelComparison.changedPixelRatio * 100).toFixed(2)}% of pixels materially changed`}</strong><small>{pixelComparison.exactPixelMatch ? "This proposal has not produced a visible PowerPoint result on the selected slide." : `Mean channel delta ${pixelComparison.meanAbsoluteChannelDelta.toFixed(2)} · changed region ${pixelComparison.changedBounds ? `${Math.round(pixelComparison.changedBounds.normalized.width * 100)}% × ${Math.round(pixelComparison.changedBounds.normalized.height * 100)}%` : "not localized"}. Pixel difference confirms change, not design quality.`}</small></span></div>}
    </section>
    {proposal.tableExceptions.length > 0 && <section className="panel proposal-exceptions"><div className="panel-heading"><div><h2>Table design exceptions</h2><p>These tables were preserved rather than forced into a generic treatment.</p></div><span className="quiet-label">{proposal.tableExceptions.length} designer check{proposal.tableExceptions.length === 1 ? "" : "s"}</span></div><div>{proposal.tableExceptions.map((exception) => <article key={exception.tableId}><Warning size={17} /><span><strong>Slide {exception.slideNumber} · {exception.rule.replaceAll("-", " ")}</strong><small>{exception.reason}</small></span></article>)}</div></section>}
    <section className="panel proposal-exceptions geometry-exceptions"><div className="panel-heading"><div><h2>Geometry and fit checks</h2><p>Slide {selectedNumber} · uncertain cases stay visible instead of triggering hidden shrinking, clipping, or speculative movement.</p></div><span className="quiet-label">{layoutExceptions.length} across deck</span></div><div>{selectedLayoutExceptions.length === 0 ? <article className="geometry-clear"><CheckCircle size={17} /><span><strong>No geometry exception on this slide</strong><small>Supported high-confidence alignment repairs are shown in the proposal comparison above.</small></span></article> : selectedLayoutExceptions.map((exception) => <article key={exception.id} className={`severity-${exception.severity}`}><Warning size={17} /><span><strong>{exception.rule.replaceAll("-", " ")} · shape {exception.shapeId}</strong><small>{exception.reason}</small></span></article>)}</div></section>
    <details className="panel review-change-details"><summary><span><ListChecks size={18} /><span><strong>Advanced change selection</strong><small>{proposal.changes.filter((change) => change.selected).length} of {proposal.changes.length} change groups selected</small></span></span><CaretRight size={17} /></summary><div className="panel-heading"><div><h2>Proposed changes</h2><p>Use this only when excluding a specific technical change group. The fast slide-review controls above are the normal workflow.</p></div><span className="quiet-label">Text and table structure: locked</span></div><div className="proposal-list">{proposal.changes.length === 0 && <div className="resource-empty"><CheckCircle size={25} /><span><strong>No deterministic changes needed</strong><small>Every slide still received an explicit disposition above.</small></span></div>}{proposal.changes.map((change, index) => <label className="proposal-change" key={`${change.kind}-${change.id}-${index}`}><input type="checkbox" checked={change.selected} disabled={proposal.status !== "pending"} onChange={() => onToggle(change.id)} /><span className="change-route"><b>{change.from}</b><CaretRight size={17} /><b>{change.to}</b></span><span>{change.kind === "table-style" ? `${change.tableIds?.length ?? 0} native tables` : change.kind === "alignment" ? `${change.alignmentRepairs?.length ?? 0} text boxes` : change.kind === "geometry" ? `${change.geometryCommands?.length ?? 0} editable objects` : change.kind === "layout-remap" ? `${change.layoutCommands?.length ?? 0} native layout relationships` : change.kind === "text-style" ? `${change.textStyleCommands?.length ?? 0} text objects` : change.kind === "decoration" ? `${change.decorationCommands?.length ?? 0} native vector shapes` : `${change.affectedRunCount} markup references`} · {change.affectedSlideNumbers.length} slide locations</span><small>{change.rationale}</small>{change.geometryCommands?.flatMap((command) => command.validation?.warnings ?? []).map((warning, warningIndex) => <em className="proposal-validation-warning" key={`${warning}-${warningIndex}`}><Warning size={13} />{warning}</em>)}</label>)}</div></details>
    <section className="panel review-completion"><div className="lock-copy"><LockKey size={20} /><span><strong>Exact-content guard</strong><small>Approving applies the selected design plan to the project only. Export remains a separate new-copy action.</small></span></div><div className="review-actions">{proposal.status === "pending" ? <><button className="button ghost" onClick={onReject}><X size={17} />Reject proposal</button><button className="button secondary" onClick={() => onOpenSlide(selectedNumber, "edit")}><Crosshair size={17} />Edit this slide</button><button className="button primary" disabled={!canApprovePlan} title={unresolvedRequestCount ? "Resolve requested changes before approving the plan." : undefined} onClick={() => onApply(selectedNumber)}><CheckCircle size={18} />Approve all &amp; continue</button></> : proposal.status === "applied" ? <><button className="button secondary" onClick={() => onOpenSlide(selectedNumber, "comment")}><ChatCircleDots size={17} />Comment for AI</button><button className="button secondary" onClick={() => onOpenSlide(selectedNumber, "edit")}><Crosshair size={17} />Continue editing</button><button className="button primary" onClick={onExport}><FileArrowDown size={17} />Export review copy</button></> : <span className="muted">Proposal rejected; the source is unchanged.</span>}</div></section></div>;
}

function ResourcesView({ project, aiSessionEnabled, onAdd, onRemove }: { project: PresentationStudioProject; aiSessionEnabled: boolean; onAdd: () => void; onRemove: (id: string) => void }) {
  const packagedBytes = project.resources.reduce((sum, resource) => sum + resource.byteLength + (resource.derivatives?.reduce((derivativeSum, derivative) => derivativeSum + derivative.byteLength, 0) ?? 0), 0);
  const extractedCount = project.resources.filter((resource) => resource.derivatives?.some((derivative) => derivative.kind === "extracted-text")).length;
  const needsReviewCount = project.resources.filter((resource) => resource.processing?.status === "needs-review").length;
  return (
    <div className="view-stack">
      <header className="view-header compact">
        <div><p className="eyebrow">Self-contained project resources</p><h1>Resources</h1><p>Every accepted file is copied into the project. Its original location is never required again.</p></div>
        <button className="button primary" onClick={onAdd}><UploadSimple size={18} />Add files</button>
      </header>
      <section className="resource-drop-zone" onClick={onAdd} role="button" tabIndex={0} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") onAdd(); }}>
        <span className="resource-drop-icon"><UploadSimple size={25} /></span>
        <span><strong>Drop files anywhere in Presentation Studio</strong><small>Drop one .pstudio package to open it. Documents, data, images, audio, video, SVG, and PowerPoint files are processed locally and embedded in the current project.</small></span>
        <span className="resource-drop-action">Choose files</span>
      </section>
      <div className="metric-strip resource-metrics">
        <Metric value={project.resources.length} label="Resources" />
        <Metric value={formatBytes(packagedBytes)} label="Packaged size" />
        <Metric value={extractedCount} label="Text indexed" />
        <Metric value={needsReviewCount} label="Needs review" />
      </div>
      <section className="panel">
        <div className="resource-list">
          <div className="resource-row resource-head"><span>Resource</span><span>Role</span><span>Processing</span><span>Size</span><span>AI access</span><span>Project</span></div>
          {project.resources.length === 0 && <div className="resource-empty"><FileText size={25} /><span><strong>No project Resources yet</strong><small>Drop files into the app or choose files above. Nothing will remain linked to its original location.</small></span></div>}
          {project.resources.map((resource) => {
            const processingStatus = resource.processing?.status ?? "stored-only";
            const hasWarnings = Boolean(resource.processing?.warnings.length);
            return <div className="resource-row" key={resource.id} title={resource.processing?.summary}>
              <span className="resource-name"><Archive size={20} /><span><strong>{resource.name}</strong><small>{resourceKindLabels[resource.kind ?? "other"]} · {resource.sha256.slice(0, 12)}… · embedded</small></span></span>
              <span className="resource-roles">{resource.roles.join(" · ")}</span>
              <span className={`processing-state ${processingStatus}`}>{hasWarnings && <Warning size={13} />}{processingStatus === "indexed" ? "Indexed" : processingStatus === "needs-review" ? "Needs review" : "Stored only"}</span>
              <span>{formatBytes(resource.byteLength)}</span>
              <span className={`access-toggle status ${resource.mcpAccess !== "none" ? "on" : ""}`} title={aiSessionEnabled ? "Automatically shared at the highest level Presentation Studio supports while AI access is on." : "Turn on AI access to share every compatible project Resource automatically."}>{resource.mcpAccess === "none" ? "Access off" : resource.mcpAccess === "preview" ? "Preview shared" : resource.mcpAccess === "text" ? "Text shared" : "Metadata shared"}</span>
              <button className="resource-remove" onClick={() => onRemove(resource.id)} title="Remove this embedded copy from the project; the original file is never deleted"><Trash size={13} />Remove</button>
            </div>;
          })}
        </div>
      </section>
      <div className="inline-note wide"><ShieldCheck size={18} />Turn on AI access once to share every embedded Resource automatically: extracted document/data text, bounded image previews, and metadata for formats Studio cannot yet read. Turning access off removes all AI Resource access at once. Original files remain local and are never changed.</div>
    </div>
  );
}

function NoSelection({ message, icon: EmptyIcon = PresentationChart }: { message: string; icon?: Icon }) {
  return <section className="no-selection"><EmptyIcon size={34} weight="light" /><h2>Nothing to show yet</h2><p>{message}</p></section>;
}

function Inspector({ deck, onOpenReview }: { deck?: DeckJob; onOpenReview: () => void }) {
  return <aside className="inspector"><div className="inspector-heading"><span>Inspector</span>{deck && <StatusPill status={deck.status} />}</div>{!deck ? <div className="inspector-empty"><Monitor size={26} /><p>Select a deck to inspect its audit and review state.</p></div> : <><div className="inspector-title"><PresentationChart size={25} /><div><strong>{deck.name}</strong><small>{classificationLabels[deck.templateClassification]}</small></div></div>{deck.failureMessage && <div className="inspector-error"><Warning size={17} />{deck.failureMessage}</div>}<dl className="inspector-list"><div><dt>Operation</dt><dd>{deck.operationScope}</dd></div><div><dt>Content</dt><dd>Preserve exact</dd></div><div><dt>Target</dt><dd>{deck.targetTemplateId ?? "Not confirmed"}</dd></div><div><dt>Slides</dt><dd>{deck.audit?.slideCount ?? "—"}</dd></div><div><dt>Findings</dt><dd>{deck.audit?.findings.length ?? "—"}</dd></div></dl>{deck.proposal && <button className="inspector-proposal" onClick={onOpenReview}><Sparkle size={18} /><span><strong>Cleanup proposal</strong><small>{deck.proposal.status} · {deck.proposal.changes.length} changes</small></span><CaretRight size={16} /></button>}<div className="guardrail-card"><ShieldCheck size={20} /><div><strong>Source protection</strong><p>The imported deck is read-only. Export always creates a new file.</p></div></div></>}</aside>;
}

export default function App() {
  const [project, setProject] = useState<PresentationStudioProject>(() => createProject());
  const projectRef = useRef(project);
  const [selectedDeckId, setSelectedDeckId] = useState<string>();
  const [activeView, setActiveView] = useState<ViewId>("batch");
  const [busy, setBusy] = useState<string>();
  const [notice, setNotice] = useState<string>();
  const [error, setError] = useState<string>();
  const [mcpEnabled, setMcpEnabled] = useState(false);
  const [mcpStatus, setMcpStatus] = useState<{ available: boolean; runtimeFile?: string }>({ available: false });
  const [nativeReadiness, setNativeReadiness] = useState<{ ready: boolean; sessionLocked: boolean; reason?: string }>({ ready: false, sessionLocked: false });
  const [mcpActivity, setMcpActivity] = useState<McpActivityState>();
  const [presentationFontCss, setPresentationFontCss] = useState("");
  const [secureAutosavePassword, setSecureAutosavePassword] = useState<string>();
  const [fileDragActive, setFileDragActive] = useState(false);
  const [templateCatalog, setTemplateCatalog] = useState<TemplateCatalog>();
  const [templateSourceBytes, setTemplateSourceBytes] = useState<Uint8Array>();
  const [templateInstalledAt, setTemplateInstalledAt] = useState<string>();
  const [templateLoading, setTemplateLoading] = useState(false);
  const [templateNativeRender, setTemplateNativeRender] = useState<NativeRenderResult>();
  const [templateNativeLoading, setTemplateNativeLoading] = useState(false);
  const [slideCatalogs, setSlideCatalogs] = useState<Record<string, SlideRenderCatalog>>({});
  const [slideCatalogLoadingDeckId, setSlideCatalogLoadingDeckId] = useState<string>();
  const [proposalCatalogs, setProposalCatalogs] = useState<Record<string, SlideRenderCatalog>>({});
  const [proposalCatalogLoadingDeckId, setProposalCatalogLoadingDeckId] = useState<string>();
  const [nativeRenderCatalogs, setNativeRenderCatalogs] = useState<Record<string, NativeRenderResult>>({});
  const [nativeRenderLoadingKey, setNativeRenderLoadingKey] = useState<string>();
  const [studioFreshPreviews, setStudioFreshPreviews] = useState<Record<string, StudioFreshPreview>>({});
  const [studioDeckBuilds, setStudioDeckBuilds] = useState<Record<string, StudioDeckBuild>>({});
  const [studioDeckQualifications, setStudioDeckQualifications] = useState<Record<string, StudioDeckQualification>>({});
  const [, setStudioHistoryVersion] = useState(0);
  const [slideWorkspaceRequest, setSlideWorkspaceRequest] = useState<SlideWorkspaceRequest>();
  const [studioOpenSlideNumber, setStudioOpenSlideNumber] = useState<number>();
  const [tourOpen, setTourOpen] = useState(false);
  const [tourStepIndex, setTourStepIndex] = useState(0);
  const fileDragDepth = useRef(0);
  const slideCatalogsRef = useRef(new Map<string, SlideRenderCatalog>());
  const proposalCatalogsRef = useRef(new Map<string, SlideRenderCatalog>());
  const nativeRenderCatalogsRef = useRef(new Map<string, NativeRenderResult>());
  const nativeMeasurementsRef = useRef(new Map<string, NativeMeasurementResult>());
  const inspectionRendersRef = useRef(new Map<string, NativeRenderResult>());
  const templateNativeRendersRef = useRef(new Map<string, NativeRenderResult>());
  const templateNativeRenderPromisesRef = useRef(new Map<string, Promise<NativeRenderResult>>());
  const sourceFigureRastersRef = useRef(new Map<string, { data: string; width: number; height: number }>());
  const studioFreshPreviewsRef = useRef(studioFreshPreviews);
  const studioDeckBuildsRef = useRef(studioDeckBuilds);
  const studioDeckQualificationsRef = useRef(studioDeckQualifications);
  const studioDeckQualificationHistoryRef = useRef<Record<string, StudioDeckQualification[]>>({});
  const studioEditHistoryRef = useRef<Record<string, { undo: StudioWebScene[]; redo: StudioWebScene[] }>>({});
  const onboardingChecked = useRef(false);
  const auditGeometryUpgradeAttempted = useRef(new Set<string>());
  const desktop = window.presentationStudioDesktop;
  const selectedDeck = project.decks.find((deck) => deck.id === selectedDeckId) ?? project.decks[0];
  const selectedStudioHistory = selectedDeck ? studioEditHistoryRef.current[selectedDeck.id] : undefined;
  const canUndoStudio = Boolean(selectedStudioHistory?.undo.length);
  const canRedoStudio = Boolean(selectedStudioHistory?.redo.length);

  function pushStudioEditHistory(deckId: string, scene: StudioWebScene) {
    const history = studioEditHistoryRef.current[deckId] ?? { undo: [], redo: [] };
    if (history.undo.at(-1)?.revision !== scene.revision) history.undo.push(scene);
    while (history.undo.length > 40) history.undo.shift();
    history.redo = [];
    studioEditHistoryRef.current[deckId] = history;
    setStudioHistoryVersion((value) => value + 1);
  }

  function restoreStudioHistory(direction: "undo" | "redo") {
    const current = projectRef.current;
    const deck = current.decks.find((item) => item.id === selectedDeck?.id);
    if (!deck?.studioScene) return;
    const history = studioEditHistoryRef.current[deck.id] ?? { undo: [], redo: [] };
    const source = direction === "undo" ? history.undo : history.redo;
    const destination = direction === "undo" ? history.redo : history.undo;
    const restored = source.pop();
    if (!restored) return;
    destination.push(deck.studioScene);
    while (destination.length > 40) destination.shift();
    studioEditHistoryRef.current[deck.id] = history;
    const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: restored, proposal: undefined, status: "ready-for-cleanup" as const } : item) }, direction === "undo" ? "studio-edit-undone" : "studio-edit-redone", `${direction === "undo" ? "Undid" : "Redid"} one Studio design transaction in ${deck.name}; source PowerPoint bytes remain unchanged.`);
    projectRef.current = next;
    setProject(next);
    setStudioHistoryVersion((value) => value + 1);
    setNotice(`${direction === "undo" ? "Undid" : "Redid"} the last Studio design change. Rebuild the PowerPoint result when you are ready to review it.`);
  }

  useEffect(() => { projectRef.current = project; }, [project]);
  useEffect(() => { studioEditHistoryRef.current = {}; setStudioHistoryVersion((value) => value + 1); }, [project.project.id]);
  useEffect(() => {
    if (activeView !== "studio") return;
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (target?.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target?.tagName ?? "")) return;
      if (!(event.metaKey || event.ctrlKey)) return;
      const undo = event.key.toLowerCase() === "z" && !event.shiftKey;
      const redo = (event.key.toLowerCase() === "z" && event.shiftKey) || event.key.toLowerCase() === "y";
      if (!undo && !redo) return;
      event.preventDefault();
      restoreStudioHistory(undo ? "undo" : "redo");
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [activeView, selectedDeckId, canUndoStudio, canRedoStudio]);
  useEffect(() => { studioFreshPreviewsRef.current = studioFreshPreviews; }, [studioFreshPreviews]);
  useEffect(() => { studioDeckBuildsRef.current = studioDeckBuilds; }, [studioDeckBuilds]);
  useEffect(() => { studioDeckQualificationsRef.current = studioDeckQualifications; }, [studioDeckQualifications]);
  useEffect(() => {
    if (project.settings.designStandardVersion === PRESENTATION_DESIGN_STANDARD.version) return;
    setProject((current) => current.settings.designStandardVersion === PRESENTATION_DESIGN_STANDARD.version ? current : {
      ...current,
      settings: { ...current.settings, designStandardVersion: PRESENTATION_DESIGN_STANDARD.version },
    });
  }, [project.settings.designStandardVersion]);
  useEffect(() => {
    if (!project.designThreads.some((thread) => ["proposal-ready", "resolved"].includes(thread.status))) return;
    setProject((current) => {
      const designThreads = removeCompletedDesignThreads(current.designThreads);
      return designThreads.length === current.designThreads.length ? current : { ...current, designThreads };
    });
  }, [project.designThreads]);
  useEffect(() => { if (!selectedDeckId && project.decks[0]) setSelectedDeckId(project.decks[0].id); }, [project.decks, selectedDeckId]);
  useEffect(() => { void desktop?.getMcpStatus().then(setMcpStatus).catch(() => undefined); }, [desktop]);
  useEffect(() => {
    if (!desktop) return;
    let canceled = false;
    const refresh = () => { void Promise.all([desktop.getNativeRenderCapabilities(), desktop.getNativeMeasurementCapabilities()]).then(([render, measurement]) => {
      if (!canceled) setNativeReadiness({ ready: render.available && measurement.available, sessionLocked: Boolean(render.sessionLocked || measurement.sessionLocked), reason: render.reason ?? measurement.reason });
    }).catch(() => { if (!canceled) setNativeReadiness({ ready: false, sessionLocked: false, reason: "PowerPoint-native QA status is unavailable." }); }); };
    refresh();
    window.addEventListener("focus", refresh);
    return () => { canceled = true; window.removeEventListener("focus", refresh); };
  }, [desktop]);

  useEffect(() => {
    const auditUpgradeKey = (deck: DeckJob) => `${deck.id}:semantic-${deck.audit?.semanticVisualVersion ?? 0}-to-${PPTX_AUDIT_SEMANTIC_VISUAL_VERSION}`;
    const candidates = project.decks.filter((deck) => deck.audit && (sceneNeedsRebuild(deck) || deck.audit.semanticVisualVersion !== PPTX_AUDIT_SEMANTIC_VISUAL_VERSION || !deck.audit.slideSize || !Array.isArray(deck.audit.editableObjects) || (deck.audit.slideCount > 0 && deck.audit.editableObjects.length === 0) || new Set(deck.audit.editableObjects.map((object) => object.id)).size !== deck.audit.editableObjects.length || deck.audit.slides.some((slide) => !slide.sourcePartSha256) || deck.audit.textBoxes.some((textBox) => typeof textBox.text !== "string" || (textBox.characterCount > 0 && textBox.text.length === 0) || (textBox.characterCount > 0 && textBox.estimatedOpticalLeftEmu <= 0))) && !auditGeometryUpgradeAttempted.current.has(auditUpgradeKey(deck)));
    if (candidates.length === 0) return;
    let canceled = false;
    for (const deck of candidates) {
      auditGeometryUpgradeAttempted.current.add(auditUpgradeKey(deck));
      const source = sourceForDeck(project, deck);
      if (!source?.bytes) continue;
      void auditPptx(source.bytes).then((audit) => {
        if (canceled) return;
        setProject((current) => ({ ...current, decks: current.decks.map((item) => item.id === deck.id ? withCompiledScene({ ...item, audit }) : item) }));
      }).catch(() => undefined);
    }
    return () => { canceled = true; };
  }, [project.decks, project.resources]);

  useEffect(() => {
    if (!desktop) return;
    let canceled = false;
    void desktop.getPresentationFonts().then(async ({ fonts }) => {
      const rules = fonts.map(fontFaceRule).join("");
      await Promise.all(fonts.map(async (font) => {
        const data = bytesFrom(font.bytes).slice().buffer;
        const face = new FontFace(font.family, data, { weight: String(font.weight), style: font.style });
        await face.load();
        document.fonts.add(face);
      }));
      if (!canceled) setPresentationFontCss(rules);
    }).catch(() => undefined);
    return () => { canceled = true; };
  }, [desktop]);

  useEffect(() => {
    if (!desktop) return;
    let canceled = false;
    setTemplateLoading(true);
    void desktop.getActiveTemplate().then(async (installed) => {
      if (!installed.installed || !installed.name || !installed.bytes) return;
      const sourceBytes = bytesFrom(installed.bytes);
      const catalog = await buildTemplateCatalog(sourceBytes, installed.name);
      if (!canceled) {
        setTemplateCatalog(catalog);
        setTemplateSourceBytes(sourceBytes);
        setTemplateInstalledAt(installed.installedAt);
      }
    }).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The installed template could not be opened.");
    }).finally(() => { if (!canceled) setTemplateLoading(false); });
    return () => { canceled = true; };
  }, [desktop]);

  const getOrBuildTemplateNativeRender = useCallback(async () => {
    if (!desktop || !templateCatalog || !templateSourceBytes) throw new Error("Install an authorized PowerPoint Template Pack before requesting native layout renders.");
    const cached = templateNativeRendersRef.current.get(templateCatalog.sha256);
    if (cached) { setTemplateNativeRender(cached); return cached; }
    let renderPromise = templateNativeRenderPromisesRef.current.get(templateCatalog.sha256);
    if (!renderPromise) {
      renderPromise = buildTemplatePreviewDeck(templateSourceBytes).then(async (preview) => {
        if (preview.layoutParts.length !== templateCatalog.layouts.length || preview.layoutParts.some((part, index) => part !== templateCatalog.layouts[index]?.sourcePart)) throw new Error("The native layout preview order does not match the installed template catalog.");
        return desktop.renderPowerPoint({ name: `${cleanFileStem(templateCatalog.name)}_layout-previews.pptx`, bytes: preview.bytes });
      }).finally(() => templateNativeRenderPromisesRef.current.delete(templateCatalog.sha256));
      templateNativeRenderPromisesRef.current.set(templateCatalog.sha256, renderPromise);
    }
    const render = await renderPromise;
    if (render.status === "ready" && render.slideCount !== templateCatalog.layouts.length) throw new Error(`PowerPoint returned ${render.slideCount} images for ${templateCatalog.layouts.length} template layouts.`);
    if (render.status === "ready") {
      templateNativeRendersRef.current.set(templateCatalog.sha256, render);
      while (templateNativeRendersRef.current.size > 3) templateNativeRendersRef.current.delete(templateNativeRendersRef.current.keys().next().value as string);
    }
    setTemplateNativeRender(render);
    return render;
  }, [desktop, templateCatalog, templateSourceBytes]);

  useEffect(() => {
    if (!["designs", "studio"].includes(activeView) || !desktop || !templateCatalog || !templateSourceBytes) return;
    let canceled = false;
    setTemplateNativeLoading(true);
    void getOrBuildTemplateNativeRender().catch((caught) => {
      if (!canceled) setTemplateNativeRender({ status: "failed", renderer: "powerpoint-native", authoritative: false, slides: [], warnings: [caught instanceof Error ? caught.message : "PowerPoint could not render the template layouts."] });
    }).finally(() => { if (!canceled) setTemplateNativeLoading(false); });
    return () => { canceled = true; };
  }, [activeView, desktop, getOrBuildTemplateNativeRender, templateCatalog, templateSourceBytes]);

  const getOrBuildSlideCatalog = useCallback(async (deck: DeckJob, current = projectRef.current) => {
    const cached = slideCatalogsRef.current.get(deck.id);
    if (cached?.sha256 === deck.sourceSha256) return cached;
    const source = sourceForDeck(current, deck);
    if (!source?.bytes) throw new Error("The embedded PowerPoint source is unavailable for slide previews.");
    const catalog = await buildSlideRenderCatalog(source.bytes, deck.name);
    if (catalog.sha256 !== deck.sourceSha256) throw new Error("The rendered slide source no longer matches the embedded deck hash.");
    slideCatalogsRef.current.set(deck.id, catalog);
    setSlideCatalogs((existing) => ({ ...existing, [deck.id]: catalog }));
    return catalog;
  }, []);

  const getOrBuildProposalCatalog = useCallback(async (deck: DeckJob, current = projectRef.current) => {
    if (!deck.proposal) throw new Error("Stage a designer cleanup proposal before requesting its render.");
    const selectedSignature = deck.proposal.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|");
    const key = `${deck.sourceSha256}:${deck.proposal.id}:${selectedSignature}`;
    const cached = proposalCatalogsRef.current.get(key);
    if (cached) return cached;
    const source = sourceForDeck(current, deck);
    if (!source?.bytes) throw new Error("The embedded PowerPoint source is unavailable for proposal previews.");
    const materialized = await buildCleanupProposalPptx(source.bytes, deck.proposal, { templateBytes: templateSourceBytes });
    const catalog = await buildSlideRenderCatalog(materialized.bytes, `${cleanFileStem(deck.name)}_proposal.pptx`);
    proposalCatalogsRef.current.set(key, catalog);
    setProposalCatalogs((existing) => ({ ...existing, [deck.id]: catalog }));
    return catalog;
  }, [templateSourceBytes]);

  const getOrBuildNativeRender = useCallback(async (deck: DeckJob, representation: "source" | "current" | "proposal" | "export", current = projectRef.current) => {
    if (!desktop) return undefined;
    const proposalRepresentation = representation === "proposal" || representation === "export";
    const selectedSignature = proposalRepresentation && deck.proposal ? deck.proposal.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|") : "source";
    const representationFamily = proposalRepresentation ? "proposal-export" : "source-current";
    const cacheKey = `${deck.id}:${representationFamily}:${deck.sourceSha256}:${deck.proposal?.id ?? "none"}:${selectedSignature}`;
    const stateKey = `${deck.id}:${representation}`;
    const cached = nativeRenderCatalogsRef.current.get(cacheKey);
    if (cached) {
      setNativeRenderCatalogs((existing) => ({ ...existing, [stateKey]: cached }));
      return cached;
    }
    const source = sourceForDeck(current, deck);
    if (!source?.bytes) throw new Error("The embedded PowerPoint source is unavailable for native rendering.");
    let bytes = bytesFrom(source.bytes);
    let name = deck.name;
    if (proposalRepresentation) {
      if (!deck.proposal) throw new Error("Stage a designer cleanup proposal before requesting its native render.");
      const materialized = await buildCleanupProposalPptx(bytes, deck.proposal, { templateBytes: templateSourceBytes });
      bytes = materialized.bytes;
      name = `${cleanFileStem(deck.name)}_${representation}.pptx`;
    }
    const result = await desktop.renderPowerPoint({ name, bytes });
    const qualified = result.status === "ready" && result.slideCount !== deck.audit?.slideCount
      ? { ...result, status: "failed" as const, authoritative: false, slides: [], warnings: [`PowerPoint returned ${result.slideCount} images for a ${deck.audit?.slideCount}-slide deck. The native render was rejected.`] }
      : result;
    if (qualified.status === "ready") {
      for (const key of nativeRenderCatalogsRef.current.keys()) {
        if (key.startsWith(`${deck.id}:${representationFamily}:`)) nativeRenderCatalogsRef.current.delete(key);
      }
      nativeRenderCatalogsRef.current.set(cacheKey, qualified);
      while (nativeRenderCatalogsRef.current.size > 8) nativeRenderCatalogsRef.current.delete(nativeRenderCatalogsRef.current.keys().next().value as string);
    }
    setNativeRenderCatalogs((existing) => {
      const next = { ...existing, [stateKey]: qualified };
      const keys = Object.keys(next);
      for (const key of keys.slice(0, Math.max(0, keys.length - 8))) delete next[key];
      return next;
    });
    return qualified;
  }, [desktop, templateSourceBytes]);

  const getOrBuildNativeMeasurement = useCallback(async (deck: DeckJob, representation: "current" | "proposal" = "current", current = projectRef.current): Promise<NativeMeasurementPacket> => {
    if (!desktop) return bindNativeMeasurement(deck);
    const proposalRepresentation = representation === "proposal";
    const selectedSignature = proposalRepresentation && deck.proposal ? deck.proposal.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|") : "source";
    const cacheKey = `${deck.id}:${representation}:${deck.sourceSha256}:${deck.proposal?.id ?? "none"}:${selectedSignature}`;
    const cached = nativeMeasurementsRef.current.get(cacheKey);
    if (cached) return bindNativeMeasurement(deck, cached);
    const source = sourceForDeck(current, deck);
    if (!source?.bytes) throw new Error("The embedded PowerPoint source is unavailable for native measurement.");
    let bytes = bytesFrom(source.bytes);
    let name = deck.name;
    if (proposalRepresentation) {
      if (!deck.proposal) throw new Error("Stage a proposal before requesting proposal measurements.");
      const materialized = await buildCleanupProposalPptx(bytes, deck.proposal, { templateBytes: templateSourceBytes });
      bytes = materialized.bytes;
      name = `${cleanFileStem(deck.name)}_proposal-measurement.pptx`;
    }
    const result = await desktop.measurePowerPoint({ name, bytes });
    const qualified = result.status === "ready" && result.slideCount !== deck.audit?.slideCount
      ? { ...result, status: "failed" as const, authority: "unknown" as const, slides: [], warnings: [`PowerPoint returned measurements for ${result.slideCount} slides in a ${deck.audit?.slideCount}-slide deck. The result was rejected.`] }
      : result;
    if (qualified.status === "ready") {
      for (const key of nativeMeasurementsRef.current.keys()) if (key.startsWith(`${deck.id}:${representation}:`)) nativeMeasurementsRef.current.delete(key);
      nativeMeasurementsRef.current.set(cacheKey, qualified);
      while (nativeMeasurementsRef.current.size > 6) nativeMeasurementsRef.current.delete(nativeMeasurementsRef.current.keys().next().value as string);
    }
    return bindNativeMeasurement(deck, qualified);
  }, [desktop, templateSourceBytes]);

  const getOrBuildInspectionRender = useCallback(async (deck: DeckJob, representation: "current" | "proposal" = "current", current = projectRef.current) => {
    if (!desktop) throw new Error("PowerPoint-native inspection requires the desktop app.");
    const proposalRepresentation = representation === "proposal";
    const selectedSignature = proposalRepresentation && deck.proposal ? deck.proposal.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|") : "source";
    const cacheKey = `${deck.id}:${representation}:${deck.sourceSha256}:${deck.proposal?.id ?? "none"}:${selectedSignature}:png-2200`;
    const cached = inspectionRendersRef.current.get(cacheKey);
    if (cached) return cached;
    const source = sourceForDeck(current, deck);
    if (!source?.bytes) throw new Error("The embedded PowerPoint source is unavailable for inspection rendering.");
    let bytes = bytesFrom(source.bytes);
    let name = deck.name;
    if (proposalRepresentation) {
      if (!deck.proposal) throw new Error("Stage a proposal before requesting proposal inspection evidence.");
      const materialized = await buildCleanupProposalPptx(bytes, deck.proposal, { templateBytes: templateSourceBytes });
      bytes = materialized.bytes;
      name = `${cleanFileStem(deck.name)}_proposal-inspection.pptx`;
    }
    const render = await desktop.renderPowerPoint({ name, bytes, width: 2200, format: "png" });
    if (render.status !== "ready" || render.slideCount !== deck.audit?.slideCount) throw new Error(render.warnings[0] ?? "PowerPoint-native inspection rendering failed qualification.");
    inspectionRendersRef.current.set(cacheKey, render);
    while (inspectionRendersRef.current.size > 4) inspectionRendersRef.current.delete(inspectionRendersRef.current.keys().next().value as string);
    return render;
  }, [desktop, templateSourceBytes]);

  useEffect(() => {
    if (!["slides", "review"].includes(activeView) || !selectedDeck?.audit) return;
    let canceled = false;
    setSlideCatalogLoadingDeckId(selectedDeck.id);
    setNativeRenderLoadingKey(`${selectedDeck.id}:current`);
    void Promise.allSettled([getOrBuildSlideCatalog(selectedDeck), getOrBuildNativeRender(selectedDeck, "current")]).then((results) => {
      const structural = results[0];
      if (!canceled && structural.status === "rejected") setError(structural.reason instanceof Error ? structural.reason.message : "The current slide designs could not be rendered.");
    }).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The current slide designs could not be rendered.");
    }).finally(() => { if (!canceled) { setSlideCatalogLoadingDeckId((value) => value === selectedDeck.id ? undefined : value); setNativeRenderLoadingKey((value) => value === `${selectedDeck.id}:current` ? undefined : value); } });
    return () => { canceled = true; };
  }, [activeView, getOrBuildNativeRender, getOrBuildSlideCatalog, selectedDeck?.id, selectedDeck?.sourceSha256]);

  const selectedProposalSignature = selectedDeck?.proposal?.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|");
  useEffect(() => {
    const proposalWorkspace = activeView === "slides" && isProposalSlideWorkspaceRequest(slideWorkspaceRequest, selectedDeck?.id);
    if (!(activeView === "review" || proposalWorkspace || (activeView === "slides" && selectedDeck?.proposal?.status === "applied")) || !selectedDeck?.proposal || !selectedProposalSignature) return;
    let canceled = false;
    setProposalCatalogLoadingDeckId(selectedDeck.id);
    setNativeRenderLoadingKey(`${selectedDeck.id}:proposal`);
    void Promise.allSettled([getOrBuildProposalCatalog(selectedDeck), getOrBuildNativeRender(selectedDeck, "proposal")]).then((results) => {
      const structural = results[0];
      if (!canceled && structural.status === "rejected") setError(structural.reason instanceof Error ? structural.reason.message : "The proposal slide designs could not be rendered.");
      else if (!canceled) setError(undefined);
    }).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The proposal slide designs could not be rendered.");
    }).finally(() => { if (!canceled) { setProposalCatalogLoadingDeckId((value) => value === selectedDeck.id ? undefined : value); setNativeRenderLoadingKey((value) => value === `${selectedDeck.id}:proposal` ? undefined : value); } });
    return () => { canceled = true; };
  }, [activeView, getOrBuildNativeRender, getOrBuildProposalCatalog, selectedDeck?.id, selectedDeck?.proposal?.id, selectedProposalSignature, slideWorkspaceRequest]);

  useEffect(() => {
    if (onboardingChecked.current) return;
    onboardingChecked.current = true;
    let canceled = false;
    void (async () => {
      let storedVersion: string | null = null;
      try {
        storedVersion = desktop ? (await desktop.getOnboardingTourVersion()).version : window.localStorage.getItem(ONBOARDING_TOUR_STORAGE_KEY);
      } catch {
        try { storedVersion = window.localStorage.getItem(ONBOARDING_TOUR_STORAGE_KEY); } catch { storedVersion = null; }
      }
      if (!canceled && shouldShowOnboardingTour(storedVersion)) {
        setTourStepIndex(0);
        setTourOpen(true);
      }
    })();
    return () => { canceled = true; };
  }, [desktop]);

  const closeOnboardingTour = useCallback((remember: boolean) => {
    setTourOpen(false);
    if (!remember) return;
    void (async () => {
      try {
        if (desktop) await desktop.setOnboardingTourVersion(ONBOARDING_TOUR_VERSION);
        else window.localStorage.setItem(ONBOARDING_TOUR_STORAGE_KEY, ONBOARDING_TOUR_VERSION);
      } catch {
        try { window.localStorage.setItem(ONBOARDING_TOUR_STORAGE_KEY, ONBOARDING_TOUR_VERSION); } catch { /* The tour can still close when preferences are unavailable. */ }
      }
    })();
  }, [desktop]);

  const openOnboardingTour = useCallback(() => {
    setTourStepIndex(0);
    setTourOpen(true);
  }, []);

  useEffect(() => {
    if (!desktop || project.resources.length === 0) return;
    const timer = window.setTimeout(() => {
      void buildProjectPackage(project)
        .then((bytes) => secureAutosavePassword ? encryptProjectPackage(bytes, secureAutosavePassword) : bytes)
        .then((bytes) => desktop.autosaveProject({ bytes, encrypted: Boolean(secureAutosavePassword) }))
        .catch(() => undefined);
    }, 1400);
    return () => window.clearTimeout(timer);
  }, [desktop, project, secureAutosavePassword]);

  useEffect(() => {
    if (!desktop) return;
    return desktop.onMcpCommand(async (request) => {
      const activityId = crypto.randomUUID();
      const startingPhase = mcpPhaseForOperation(request.operation, request.input);
      setMcpActivity({ id: activityId, operation: request.operation, state: "active", phase: startingPhase });
      try {
      const result = await (async () => {
      const current = projectRef.current;
      if (request.operation === "get_app_status") {
        const [renderCapabilities, measurementCapabilities] = await Promise.all([desktop.getNativeRenderCapabilities(), desktop.getNativeMeasurementCapabilities()]);
        return { app: "Presentation Studio", designStandardVersion: PRESENTATION_DESIGN_STANDARD.version, project: { name: current.project.name, type: current.project.type, resourceCount: current.resources.length, deckCount: current.decks.length, slideCount: current.decks.reduce((sum, deck) => sum + (deck.audit?.slideCount ?? 0), 0), submittedDesignThreadCount: current.designThreads.filter((thread) => thread.status === "submitted").length, updatedAt: current.project.updatedAt }, aiSessionAccess: mcpEnabled, nativePowerPoint: { ready: renderCapabilities.available && measurementCapabilities.available, sessionLocked: Boolean(renderCapabilities.sessionLocked || measurementCapabilities.sessionLocked), render: renderCapabilities, measurement: measurementCapabilities } };
      }
      if (!mcpEnabled) throw new Error("Enable AI session access in Presentation Studio before reading project metadata or staging work.");
      if (request.operation === "get_template_layout_catalog") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting its layout catalog.");
        return {
          updatedAt: current.project.updatedAt,
          template: { id: templateCatalog.id, name: templateCatalog.name, sha256: templateCatalog.sha256, slideWidth: templateCatalog.slideWidth, slideHeight: templateCatalog.slideHeight, masterCount: templateCatalog.masterCount, layoutCount: templateCatalog.layouts.length },
          nativePreview: { status: templateNativeRender?.status ?? (templateNativeLoading ? "rendering" : "not-requested"), authoritative: templateNativeRender?.authoritative ?? false, renderer: templateNativeRender?.renderer },
          layouts: templateCatalog.layouts.map((layout, index) => ({ id: layout.id, ordinal: index + 1, name: layout.name, category: layout.category, sourcePart: layout.sourcePart, semantic: layout.semantic })),
          instruction: "Choose layouts by semantic compatibility, then inspect the exact Current and Proposal PowerPoint-native renders. Do not infer fit from a layout name alone.",
        };
      }
      if (request.operation === "get_template_layout_render") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting a layout render.");
        const layoutId = String(request.input.layoutId ?? "");
        const layoutIndex = templateCatalog.layouts.findIndex((layout) => layout.id === layoutId);
        if (layoutIndex < 0) throw new Error("The requested layout ID is not present in the active Template Pack. Read get_template_layout_catalog again.");
        const render = await getOrBuildTemplateNativeRender();
        if (render.status !== "ready") throw new Error(render.warnings[0] ?? "PowerPoint-native template rendering is unavailable.");
        const slide = render.slides.find((item) => item.number === layoutIndex + 1);
        if (!slide) throw new Error("PowerPoint did not return the requested layout image.");
        const layout = templateCatalog.layouts[layoutIndex];
        return { updatedAt: current.project.updatedAt, template: { id: templateCatalog.id, name: templateCatalog.name, sha256: templateCatalog.sha256 }, layout: { id: layout.id, ordinal: layoutIndex + 1, name: layout.name, category: layout.category, sourcePart: layout.sourcePart, semantic: layout.semantic }, renderer: render.renderer, pipeline: render.pipeline, powerPointVersion: render.powerPointVersion, sourceSha256: render.sourceSha256, rasterSha256: slide.sha256, authoritative: true, qaNote: "Authoritative Microsoft PowerPoint render of the empty approved custom layout. Use the semantic slot geometry with this image; empty picture/table/content slots may render blank until content is bound.", mimeType: slide.mimeType, data: bytesToBase64(bytesFrom(slide.bytes)), width: slide.width, height: slide.height };
      }
      if (request.operation === "recommend_slide_layouts") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting layout recommendations.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const limit = Math.max(1, Math.min(12, Number(request.input.limit ?? 6)));
        const profile = contentProfileForSlide(deck, slideNumber);
        const recommendations = rankLayoutCompatibility(templateCatalog.layouts, profile).slice(0, limit).map((result) => {
          const layout = templateCatalog.layouts.find((item) => item.id === result.layoutId)!;
          return { ...result, layout: { id: layout.id, ordinal: templateCatalog.layouts.indexOf(layout) + 1, name: layout.name, category: layout.category, sourcePart: layout.sourcePart, semantic: layout.semantic } };
        });
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slide: { number: slideNumber, title: deck.audit.slides.find((slide) => slide.number === slideNumber)?.title }, profile, recommendations, preservationRequired: true, instruction: "Treat scores as deterministic shortlist evidence, not final visual approval. Inspect the authoritative Current render, select the strongest compatible approved layout, compose without rewriting content, and render the Proposal again." };
      }
      if (request.operation === "get_slide_design_work_order") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting a design work order.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        const slideNumber = Number(request.input.slideNumber);
        const [nativeRender, measurement] = await Promise.all([getOrBuildInspectionRender(deck, "current", current), getOrBuildNativeMeasurement(deck, "current", current)]);
        const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber, projectUpdatedAt: current.project.updatedAt, templateCatalog, currentRender: nativeRender, currentMeasurement: measurement, threads: current.designThreads });
        const image = nativeRender.slides.find((item) => item.number === slideNumber);
        if (!image) throw new Error("PowerPoint did not return the requested work-order image.");
        return { ...workOrder, mimeType: image.mimeType, data: bytesToBase64(bytesFrom(image.bytes)), width: image.width, height: image.height };
      }
      if (request.operation === "get_slide_inspection_packet") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting an inspection packet.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        const slideNumber = Number(request.input.slideNumber);
        const representation = request.input.representation === "proposal" ? "proposal" as const : "current" as const;
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const [nativeRender, measurement] = await Promise.all([getOrBuildInspectionRender(deck, representation, current), getOrBuildNativeMeasurement(deck, representation, current)]);
        const baselineMeasurement = representation === "proposal" ? await getOrBuildNativeMeasurement(deck, "current", current) : undefined;
        const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber, projectUpdatedAt: current.project.updatedAt, templateCatalog, currentRender: nativeRender, currentMeasurement: measurement, threads: current.designThreads });
        const metrics = calculateDesignMetrics(deck, measurement, baselineMeasurement);
        const packet = buildInspectionPacket({ deck, slideNumber, projectUpdatedAt: current.project.updatedAt, workOrder, render: nativeRender, measurement, metrics: metrics.slides.find((item) => item.slideNumber === slideNumber)! });
        const issueLedger = buildDesignRepairLedger({ deck, slideNumber, representation, metrics: metrics.slides.find((item) => item.slideNumber === slideNumber), threads: current.designThreads });
        const slideImage = nativeRender.slides.find((item) => item.number === slideNumber);
        if (!slideImage) throw new Error("PowerPoint did not return the requested inspection image.");
        const images = await inspectionRasterEvidence(slideImage, packet.visualEvidence.crops);
        return { ...packet, representation, issueLedger, images };
      }
      if (request.operation === "get_deck_design_work_order") {
        if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before requesting a design work order.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        const [nativeRender, measurement] = await Promise.all([getOrBuildNativeRender(deck, "current", current), getOrBuildNativeMeasurement(deck, "current", current)]);
        return buildDeckDesignWorkOrder({ deck, projectUpdatedAt: current.project.updatedAt, templateCatalog, currentRender: nativeRender, currentMeasurement: measurement, threads: current.designThreads });
      }
      if (request.operation === "get_deck_contact_sheet") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        const representation = request.input.representation === "proposal" ? "proposal" as const : "current" as const;
        const page = Number(request.input.page ?? 1);
        const nativeRender = await getOrBuildNativeRender(deck, representation, current);
        if (!nativeRender) throw new Error("PowerPoint-native deck rendering is unavailable.");
        const sheet = await renderNativeContactSheet(nativeRender, page);
        const { bytes: sheetBytes, ...sheetMetadata } = sheet;
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name, sourceSha256: deck.sourceSha256 },
          representation,
          proposalId: representation === "proposal" ? deck.proposal?.id : undefined,
          renderer: nativeRender.renderer,
          pipeline: nativeRender.pipeline,
          powerPointVersion: nativeRender.powerPointVersion,
          authoritative: true,
          page: sheet.page,
          pageCount: sheet.pageCount,
          pageSize: sheet.pageSize,
          totalSlides: sheet.totalSlides,
          firstSlideNumber: sheet.firstSlideNumber,
          lastSlideNumber: sheet.lastSlideNumber,
          images: [{ ...sheetMetadata, data: bytesToBase64(sheetBytes) }],
          instruction: "Review every page for cross-slide hierarchy, density, pacing, repeated-component consistency, and visual outliers. Open individual inspection packets for precise diagnosis; do not infer point geometry from this overview.",
        };
      }
      if (request.operation === "list_decks") return { updatedAt: current.project.updatedAt, decks: current.decks.map((deck) => ({ id: deck.id, name: deck.name, status: deck.status, operationScope: deck.operationScope, templateClassification: deck.templateClassification, targetTemplateId: deck.targetTemplateId, slideCount: deck.audit?.slideCount ?? 0, findingCount: deck.audit?.findings.length ?? 0 })) };
      if (request.operation === "get_deck_scene_summary") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current hybrid scene. Reopen or re-audit the source deck.");
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name, status: deck.status, targetTemplateId: deck.targetTemplateId },
          scene: {
            schema: deck.scene.schema,
            version: deck.scene.version,
            revision: deck.scene.revision,
            sourceSha256: deck.scene.sourceSha256,
            slideSize: deck.scene.slideSize,
            templateBinding: deck.scene.templateBinding,
            slideCount: deck.scene.slides.length,
            objectCount: deck.scene.objects.length,
            fidelityCounts: deck.scene.fidelityCounts,
          },
          preservationEnvelope: {
            sourceBytesAuthoritative: deck.scene.preservationEnvelope.sourceBytesAuthoritative,
            nativeRenderAuthoritativeForAppearance: deck.scene.preservationEnvelope.nativeRenderAuthoritativeForAppearance,
            exportStrategy: deck.scene.preservationEnvelope.exportStrategy,
            protectedFeatures: deck.scene.preservationEnvelope.protectedFeatures,
            blockingFeatures: deck.scene.preservationEnvelope.blockingFeatures,
          },
        };
      }
      if (request.operation === "get_slide_scene") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current hybrid scene. Reopen or re-audit the source deck.");
        const slideNumber = Number(request.input.slideNumber);
        const slide = deck.scene.slides.find((item) => item.number === slideNumber);
        if (!slide) throw new Error(`Choose a slide from 1 to ${deck.scene.slides.length}.`);
        const objects = deck.scene.objects.filter((object) => object.slideNumber === slideNumber).map((object) => ({
          ...object,
          geometryInches: { x: object.geometry.x / 914_400, y: object.geometry.y / 914_400, width: object.geometry.width / 914_400, height: object.geometry.height / 914_400 },
        }));
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId },
          sceneRevision: deck.scene.revision,
          slide,
          objects,
          instruction: "Use only operations declared true for each object. Preserved or conversion-required internals must remain intact; native PowerPoint pixels remain authoritative for appearance.",
        };
      }
      if (request.operation === "get_studio_web_scene") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current PowerPoint audit and preservation scene.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const catalog = await getOrBuildSlideCatalog(deck, current);
        const scene = deck.studioScene ?? compileStudioWebScene(deck, catalog);
        const slide = scene.slides.find((item) => item.slideNumber === slideNumber);
        if (!slide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId },
          scene: { schema: scene.schema, version: scene.version, revision: scene.revision, slideSize: scene.slideSize, sourceSlideSize: scene.sourceSlideSize, rhythm: scene.rhythm, designMemory: scene.designMemory ?? [], componentLibrary: scene.componentLibrary ?? [], tableLibrary: scene.tableLibrary ?? [], tableContinuationPlans: scene.tableContinuationPlans ?? [], designSystem: scene.designSystem, persisted: Boolean(deck.studioScene) },
          slide: {
            id: slide.id, slideNumber: slide.slideNumber, recipe: slide.recipe, targetLayoutId: slide.targetLayoutId, targetLayoutName: slide.targetLayoutName, status: slide.status, designRationale: slide.designRationale, sourceTextHash: slide.sourceTextHash,
            sourceText: deck.audit.slides.find((item) => item.number === slideNumber)?.text ?? "",
            contentCoverage: slide.contentCoverage,
            sacredTemplateTitle: slideNumber === 1 && isProtectedOrnlTemplateSlide(deck, slideNumber),
            protectedTemplateSlide: isProtectedOrnlTemplateSlide(deck, slideNumber),
            recommendedRecipe: isProtectedOrnlTemplateSlide(deck, slideNumber) ? "source" : recommendedStudioRecipe(slide),
            figureTreatments: slide.figureTreatments,
            conceptReferences: (slide.conceptReferences ?? []).map((reference) => ({
              ...reference,
              resourceAvailable: current.resources.some((resource) => resource.id === reference.resourceId && resource.sha256 === reference.resourceSha256),
              previewAuthorized: current.resources.some((resource) => resource.id === reference.resourceId && resource.sha256 === reference.resourceSha256 && resource.mcpAccess === "preview"),
            })),
            visualNeeds: (slide.visualNeeds ?? []).map((need) => ({
              id: need.id, type: need.type, status: need.status, sourceTextHash: need.sourceTextHash, reason: need.reason, communicationJob: need.communicationJob, expression: need.expression,
              approvedInfluences: need.approvedInfluences, disclosurePolicy: need.disclosurePolicy, brandExpression: need.brandExpression, targetSlot: need.targetSlot, structureInventory: need.structureInventory,
              linkedConceptReferenceId: need.linkedConceptReferenceId, resolutionNote: need.resolutionNote, promptReady: Boolean(need.promptPackage.prompt), createdAt: need.createdAt, updatedAt: need.updatedAt,
            })),
            constraints: slide.constraints ?? [],
            qualityReview: slide.qualityReview,
            tableContinuationPlans: (scene.tableContinuationPlans ?? []).filter((plan) => plan.sourceSlideNumber === slideNumber),
            designImpact: analyzeStudioDesignImpact(slide),
            nodes: slide.nodes.map((node) => ({
              id: node.id, sourceObjectId: node.sourceObjectId, sourceShapeId: node.sourceShapeId, sourceBinding: node.sourceBinding, sourceTextOrder: node.sourceTextOrder, sourceAtom: node.sourceAtom, name: node.name, kind: node.kind, role: node.role, component: node.component, visible: node.visible, locked: node.locked, exactContent: node.exactContent, text: node.text, textHash: node.textHash, sourceParagraphs: node.sourceParagraphs, tableId: node.tableId, table: node.table, connector: node.connector,
              sourceFrameInches: { x: node.sourceFrame.x / 914_400, y: node.sourceFrame.y / 914_400, width: node.sourceFrame.width / 914_400, height: node.sourceFrame.height / 914_400, rotation: node.sourceFrame.rotation },
              frameInches: { x: node.frame.x / 914_400, y: node.frame.y / 914_400, width: node.frame.width / 914_400, height: node.frame.height / 914_400, rotation: node.frame.rotation },
              style: node.style,
            })),
          },
          instruction: isProtectedOrnlTemplateSlide(deck, slideNumber)
            ? `HARD RULE: the approved ORNL template composition on slide ${slideNumber} is sacred. Keep recipe source and preserve its approved marks, artwork, photography, legal copy, geometry, master, and layout exactly. Do not send nodeFrames, nodeStyles, figure treatments, or another recipe for this slide.`
            : slide.contentCoverage.exactTextMapped
            ? `Design in this semantic HTML/CSS scene, not by preserving weak source coordinates. Reuse matching ready-rated deck designMemory patterns and the shared rhythm before inventing a one-off layout.${(slide.conceptReferences?.length ?? 0) > 0 ? " Read each attached concept-only reference and follow only its approved visual influences; generated text, logos, data, claims, and technical details remain untrusted." : ""} Source paragraphs are exposed as exact semantic atom candidates, but approved wording remains locked. Prefer the recommended recipe: objective columns for 2–4 parallel paragraphs; challenge + evidence for one assertion, three peer challenges, and dense source-locked evidence; process flow for four source inputs, ordered stages, one output, and exact supporting copy; steps + evidence for 2–5 instruction atoms beside one visual; labeled figures for multiple images with labels/captions; cards only for already-separated comparisons; and table for native tables. Pair process icons/labels by source relationship and z-order, not page position, and do not classify technical content as footer furniture solely because it is low on the slide. Use compilerMode fresh-composition whenever any visible node is catalog-derived or semantic-atom; source-bound-overlay cannot represent those nodes and will be rejected. Then call preview_studio_fresh_composition, run get_studio_slide_critique against the original, and complete the bounded visual loop before any save.`
            : "This slide's complete source text is returned, but not every character maps to an editable Studio node because grouped, inherited, or unsupported content still needs semantic atomization. Do not preview or save a fresh composition that would omit it. Preserve the native source slide or defer redesign until atomization support covers the missing content.",
        };
      }
      if (request.operation === "get_studio_deck_consistency") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("Create and persist the Studio Web Scene before reviewing deck consistency.");
        const review = analyzeStudioDeckConsistency(deck.studioScene);
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, review, instruction: review.issueCount ? "Inspect each exact slide/node finding, apply only bounded source-preserving component or layout refinements, rebuild the affected slides, then rerun this review before PowerPoint-native whole-deck qualification." : "The supported repeated systems are internally consistent. Continue with PowerPoint-native whole-deck qualification; this deterministic review does not prove aesthetic quality." };
      }
      if (request.operation === "preview_studio_fresh_composition") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before building a fresh composition.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before building a fresh composition.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const studioSlide = deck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
        if (!studioSlide || studioSlide.status !== "designed") throw new Error("Recompose this slide in the Studio Web Scene before requesting a fresh-composition preview.");
        const preview = await buildFreshStudioPreview(deck, slideNumber, deck.studioScene);
        if (preview.nativeRender?.status !== "ready" || !preview.nativeRender.authoritative) throw new Error(`Microsoft PowerPoint could not render the fresh composition authoritatively: ${preview.nativeRender?.reason ?? preview.nativeRender?.warnings.join(" ") ?? "native rendering is unavailable"}`);
        if (preview.nativeMeasurement?.status !== "ready" || preview.nativeMeasurement.authority !== "powerpoint-native") throw new Error(`Microsoft PowerPoint could not remeasure the fresh composition authoritatively: ${preview.nativeMeasurement?.reason ?? preview.nativeMeasurement?.warnings.join(" ") ?? "native measurement is unavailable"}`);
        const images = preview.nativeRender.slides;
        if (!images.length || images.length !== preview.outputSlides.length) throw new Error("Microsoft PowerPoint did not return every materialized fresh-composition slide image.");
        studioFreshPreviewsRef.current = { ...studioFreshPreviewsRef.current, [`${deck.id}:${slideNumber}`]: preview };
        setStudioFreshPreviews(studioFreshPreviewsRef.current);
        setSelectedDeckId(deck.id);
        setActiveView("studio");
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          slide: { number: slideNumber, recipe: studioSlide.recipe, rationale: studioSlide.designRationale },
          sceneRevision: deck.studioScene.revision,
          compilerMode: "fresh-composition",
          editablePowerPoint: { slideCount: preview.slideCount, textNodeCount: preview.textNodeCount, tableCount: preview.tableCount, imageCount: preview.imageCount, ignoredSourceFurnitureCount: preview.ignoredSourceFurnitureCount, generatedComponentCount: preview.generatedComponentCount },
          renderer: preview.nativeRender.renderer,
          pipeline: preview.nativeRender.pipeline,
          powerPointVersion: preview.nativeRender.powerPointVersion,
          measurementAuthority: preview.nativeMeasurement?.authority,
          measuredTextOverflowCount: preview.nativeMeasurement?.status === "ready" ? nativeTextOverflows(preview.nativeMeasurement).length : undefined,
          outputSlides: preview.outputSlides,
          rasterSha256s: images.map((image) => image.sha256),
          sourceSha256: preview.nativeRender.sourceSha256,
          authoritative: true,
          warnings: preview.warnings,
          preservationTradeoff: studioSlide.recipe === "template-layout" ? "This is a newly composed editable slide with converted non-placeholder artwork from the selected Template Pack layout. It preserves exact source content, but not original master behavior, animations, transitions, or unsupported PowerPoint internals." : "This is a newly composed editable slide in the Studio ORNL system. It preserves exact visible source text and native table content/merged structure, but it does not preserve the imported source master, animations, transitions, or unsupported PowerPoint internals.",
          applied: false,
          saved: false,
          images: images.map((image, index) => ({
            label: preview.outputSlides[index]?.continuation ? `Table continuation ${preview.outputSlides[index].continuation!.segmentOrdinal} of ${preview.outputSlides[index].continuation!.segmentCount}` : `Output slide ${index + 1}`,
            representation: "export",
            outputSlideNumber: preview.outputSlides[index]?.outputSlideNumber,
            sourceSlideNumber: preview.outputSlides[index]?.sourceSlideNumber,
            continuation: preview.outputSlides[index]?.continuation,
            mimeType: image.mimeType,
            data: bytesToBase64(bytesFrom(image.bytes)),
            width: image.width,
            height: image.height,
            sha256: image.sha256,
          })),
        };
      }
      if (request.operation === "get_studio_slide_critique") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create, design, and build the Studio slide before requesting its visual critique.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Build and inspect the exact current revision before critique.");
        const slideNumber = Number(request.input.slideNumber);
        const studioSlide = deck.studioScene.slides.find((slide) => slide.slideNumber === slideNumber);
        if (!studioSlide) throw new Error("The requested Studio slide is unavailable.");
        const preview = studioFreshPreviewsRef.current[`${deck.id}:${slideNumber}`];
        if (!preview || preview.sceneRevision !== deck.studioScene.revision || preview.slideUpdatedAt !== studioSlide.updatedAt || preview.nativeRender?.status !== "ready" || preview.nativeMeasurement?.status !== "ready") throw new Error("Build this exact Studio slide revision in PowerPoint before requesting Found issues.");
        if (preview.slideCount !== 1) throw new Error(`Source slide ${slideNumber} materializes as ${preview.slideCount} continuation slides. Inspect every image returned by preview_studio_fresh_composition, then run Build all and run_deck_qualification so each output receives its own revision-bound visual review.`);
        const exportImage = preview.nativeRender.slides[0];
        const sourceRender = await getOrBuildNativeRender(deck, "current", current);
        const sourceImage = sourceRender?.status === "ready" ? sourceRender.slides.find((slide) => slide.number === slideNumber) : undefined;
        if (!sourceImage || !exportImage) throw new Error("Both authoritative original and export-result images are required for Studio critique.");
        const critique = critiqueStudioSlide(deck.studioScene, slideNumber, preview.nativeMeasurement);
        const conceptImages = [] as Array<{ label: string; representation: "concept"; data: string; mimeType: "image/png"; width: number; height: number; sha256: string; referenceId: string; approvedInfluences: string[] }>;
        for (const reference of (studioSlide.conceptReferences ?? []).slice(0, 2)) {
          const resource = current.resources.find((item) => item.id === reference.resourceId && item.sha256 === reference.resourceSha256 && item.mcpAccess === "preview");
          if (!resource) continue;
          const concept = await boundedResourceImagePreview(resource);
          conceptImages.push({ label: `Concept-only reference · ${reference.origin}`, representation: "concept", data: bytesToBase64(concept.bytes), mimeType: concept.mimeType, width: concept.width, height: concept.height, sha256: concept.sha256, referenceId: reference.id, approvedInfluences: reference.approvedInfluences });
        }
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          slideNumber,
          sceneRevision: deck.studioScene.revision,
          slideUpdatedAt: studioSlide.updatedAt,
          rasterSha256: exportImage.sha256,
          critique,
          issueLedger: { phase: "found-issues", issueCount: critique.issues.length, autoFixableCount: critique.autoFixableCount },
          images: [
            { label: "Original PowerPoint reference", representation: "source", data: bytesToBase64(bytesFrom(sourceImage.bytes)), mimeType: sourceImage.mimeType, width: sourceImage.width, height: sourceImage.height, sha256: sourceImage.sha256 },
            ...conceptImages,
            { label: "Current Studio export result", representation: "export", data: bytesToBase64(bytesFrom(exportImage.bytes)), mimeType: exportImage.mimeType, width: exportImage.width, height: exportImage.height, sha256: exportImage.sha256 },
          ],
          instruction: critique.verdict === "ready" ? `Inspect the original and export images${conceptImages.length ? " plus each concept-only image for only its approved visual influences" : ""} for issues the deterministic critic cannot see, then record a ready or revise verdict for this exact raster.` : "Fix the objective issue ledger with high-level Studio operations, rebuild the exact slide, and request critique again. Do not merely reduce all type sizes.",
        };
      }
      if (request.operation === "repair_studio_objective_issues") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Request a fresh Studio critique before applying its objective repair pass.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create, design, and build the Studio slide before applying objective repairs.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Build and critique the exact current revision before repair.");
        const slideNumber = Number(request.input.slideNumber);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot enter automatic repair.`);
        const studioSlide = deck.studioScene.slides.find((slide) => slide.slideNumber === slideNumber);
        if (!studioSlide) throw new Error("The requested Studio slide is unavailable.");
        const previewKey = `${deck.id}:${slideNumber}`;
        const preview = studioFreshPreviewsRef.current[previewKey];
        if (preview?.slideCount !== 1) throw new Error("A multi-slide table continuation cannot use the single-slide automatic repair transaction. Refine the source table/continuation plan, rebuild every output, and use whole-deck qualification.");
        const exportImage = preview?.nativeRender?.status === "ready" ? preview.nativeRender.slides[0] : undefined;
        if (!preview || preview.sceneRevision !== deck.studioScene.revision || preview.slideUpdatedAt !== studioSlide.updatedAt || preview.nativeMeasurement?.status !== "ready" || !exportImage || request.input.expectedRasterSha256 !== exportImage.sha256) throw new Error("The PowerPoint-native repair evidence is stale. Build and critique the exact slide revision again.");
        const beforeSignature = studioSlideContentSignature(studioSlide);
        const requestedIssueIds = Array.isArray(request.input.issueIds) ? request.input.issueIds.map(String) : undefined;
        const result = applyStudioDeterministicRepairPass(deck.studioScene, slideNumber, preview.nativeMeasurement, requestedIssueIds);
        const resultSlide = result.scene.slides.find((slide) => slide.slideNumber === slideNumber)!;
        if (studioSlideContentSignature(resultSlide) !== beforeSignature) throw new Error("Studio refused an automatic repair that changed exact source content.");
        if (!result.requiresNativeRerender) return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slideNumber, priorSceneRevision: deck.studioScene.revision, sceneRevision: deck.studioScene.revision, issueCount: result.critique.issues.length, fixedIssueIds: result.fixedIssueIds, deferredIssueIds: result.deferredIssueIds, actions: result.actions, changedNodeIds: [], requiresNativeRerender: false, saved: false, exported: false, instruction: "No deterministic geometry change was safe or necessary. Use the deferred issue guidance for a material recipe, table, figure, concept, or human review decision; do not force a coordinate patch." };
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const revisionBoundThreads = markSubmittedThreadsForReanchor(current.designThreads, deck.id, slideNumber, resultSlide.updatedAt, addressedThreadIds);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, [slideNumber], addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-objective-repair", `AI applied ${result.fixedIssueIds.length} deterministic PowerPoint-native repair${result.fixedIssueIds.length === 1 ? "" : "s"} to slide ${slideNumber} of ${deck.name}; exact source content remains unchanged and a fresh native render is required.`);
        delete studioFreshPreviewsRef.current[previewKey];
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return { updatedAt: next.project.updatedAt, deck: { id: deck.id, name: deck.name }, slideNumber, priorSceneRevision: deck.studioScene.revision, sceneRevision: result.scene.revision, issueCount: result.critique.issues.length, fixedIssueIds: result.fixedIssueIds, deferredIssueIds: result.deferredIssueIds, actions: result.actions, changedNodeIds: result.changedNodeIds, requiresNativeRerender: true, saved: false, exported: false, instruction: "Fixing is complete for the bounded deterministic items. Rebuild this exact slide in PowerPoint, request Found issues again, inspect the new original/export pixel pair, and never reuse the stale raster." };
      }
      if (request.operation === "record_studio_visual_critique") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Request a fresh Studio slide critique before recording judgment.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("The requested Studio deck is unavailable.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Build and critique the exact current revision first.");
        const slideNumber = Number(request.input.slideNumber);
        const studioSlide = deck.studioScene.slides.find((slide) => slide.slideNumber === slideNumber);
        if (!studioSlide) throw new Error("The requested Studio slide is unavailable.");
        const preview = studioFreshPreviewsRef.current[`${deck.id}:${slideNumber}`];
        if (preview?.slideCount !== 1) throw new Error("A multi-slide table continuation must be reviewed output-by-output through whole-deck qualification; one source-slide raster verdict cannot approve multiple PowerPoint slides.");
        const exportImage = preview?.nativeRender?.status === "ready" ? preview.nativeRender.slides[0] : undefined;
        if (!preview || preview.sceneRevision !== deck.studioScene.revision || preview.slideUpdatedAt !== studioSlide.updatedAt || preview.nativeMeasurement?.status !== "ready" || !exportImage || request.input.rasterSha256 !== exportImage.sha256) throw new Error("The Studio export raster changed. Request a fresh critique before recording visual judgment.");
        const critique = critiqueStudioSlide(deck.studioScene, slideNumber, preview.nativeMeasurement);
        const rawVisualIssues = Array.isArray(request.input.visualIssues) ? request.input.visualIssues as Array<Record<string, unknown>> : [];
        const visualIssues: StudioQualityIssue[] = rawVisualIssues.map((raw, index) => ({
          id: `ai-visual-${index + 1}`,
          category: String(raw.category ?? "other") as StudioQualityIssue["category"],
          severity: String(raw.severity ?? "minor") as StudioQualityIssue["severity"],
          source: "ai-visual",
          nodeIds: Array.isArray(raw.nodeIds) ? raw.nodeIds.map(String) : [],
          message: String(raw.message ?? "").trim().slice(0, 1_000),
          recommendation: String(raw.recommendation ?? "").trim().slice(0, 1_000),
          autoFixable: raw.autoFixable === true,
        }));
        const requestedVerdict = String(request.input.verdict ?? "revise") as "ready" | "revise" | "hold";
        const pass = critique.iteration.currentPass;
        const serious = critique.issues.some((issue) => issue.severity !== "minor") || visualIssues.some((issue) => issue.severity !== "minor");
        let recordedVerdict: "ready" | "revise" | "hold" = requestedVerdict === "ready" && serious ? "revise" : requestedVerdict;
        if (pass >= 3 && recordedVerdict !== "ready") recordedVerdict = "hold";
        const review = {
          sceneRevision: deck.studioScene.revision,
          slideUpdatedAt: studioSlide.updatedAt,
          rasterSha256: exportImage.sha256,
          pass,
          maxPasses: 3 as const,
          requestedVerdict,
          recordedVerdict,
          rationale: String(request.input.rationale ?? "").trim().slice(0, 1_000),
          objectiveIssues: critique.issues,
          visualIssues,
          recordedAt: new Date().toISOString(),
        };
        const memoryEntry = recordedVerdict === "ready" ? {
          contentSignature: studioSlideContentSignature(studioSlide),
          recipe: studioSlide.recipe,
          targetLayoutId: studioSlide.targetLayoutId,
          targetLayoutName: studioSlide.targetLayoutName,
          rhythm: deck.studioScene.rhythm ?? { safeMarginPt: 18, gridPt: 6, compactGapPt: 8, normalGapPt: 12, primaryGapPt: 18, captionGapPt: 8, titleContentGapPt: 18 },
          adoptedFromSlideNumber: slideNumber,
          qualityRasterSha256: exportImage.sha256,
          recordedAt: review.recordedAt,
        } : undefined;
        const designMemory = memoryEntry ? [...(deck.studioScene.designMemory ?? []).filter((entry) => entry.contentSignature !== memoryEntry.contentSignature), memoryEntry].slice(-80) : deck.studioScene.designMemory;
        let studioScene: StudioWebScene = { ...deck.studioScene, designMemory, slides: deck.studioScene.slides.map((slide) => slide.slideNumber === slideNumber ? { ...slide, qualityReview: review } : slide) };
        if (recordedVerdict === "ready") studioScene = resolveStudioVisualNeeds(studioScene, slideNumber, "The editable reconstruction passed the bounded PowerPoint-native visual review.");
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene } : item) }, "mcp-studio-visual-critique-recorded", `AI recorded Studio visual pass ${pass}/3 as ${recordedVerdict} for slide ${slideNumber} of ${deck.name}; slide design geometry and source bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          slideNumber,
          review,
          issueLedger: { phase: recordedVerdict === "ready" ? "ready" : "found-issues", issueCount: critique.issues.length + visualIssues.length, autoFixableCount: [...critique.issues, ...visualIssues].filter((issue) => issue.autoFixable).length },
          instruction: recordedVerdict === "ready" ? "The exact Studio result is ready for human review. No file was saved or exported." : recordedVerdict === "revise" ? "Use one materially different high-level layout, figure, hierarchy, or spacing operation; rebuild and critique the new exact raster." : "The three-pass loop is held. Keep the exact current result visible and ask the human to resolve the remaining visual or content ambiguity.",
        };
      }
      if (request.operation === "build_studio_presentation") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio scenes again before building the presentation.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and design the Studio Web Scene before building the central presentation.");
        assertSacredOrnlTitleSlideIntegrity(deck, deck.studioScene);
        const unconverted = unsupportedSourceSlideNumbers(deck, deck.studioScene);
        if (unconverted.length) throw new Error(`The central presentation cannot be built until every slide has a Studio or converted-template design. Still source-only: ${unconverted.join(", ")}.`);
        let rebuiltSlideCount = 0;
        let reusedSlideCount = 0;
        for (const studioSlide of deck.studioScene.slides) {
          const key = `${deck.id}:${studioSlide.slideNumber}`;
          const cached = studioFreshPreviewsRef.current[key];
          if (cached?.slideUpdatedAt === studioSlide.updatedAt && cached.nativeRender?.status === "ready" && cached.nativeMeasurement?.status === "ready") {
            reusedSlideCount += 1;
            continue;
          }
          const preview = await buildFreshStudioPreview(deck, studioSlide.slideNumber, deck.studioScene);
          studioFreshPreviewsRef.current = { ...studioFreshPreviewsRef.current, [key]: preview };
          setStudioFreshPreviews(studioFreshPreviewsRef.current);
          rebuiltSlideCount += 1;
        }
        const build = await buildCentralStudioDeck(deck, deck.studioScene);
        invalidateStudioQualification(deck.id);
        studioDeckBuildsRef.current = { ...studioDeckBuildsRef.current, [deck.id]: build };
        setStudioDeckBuilds(studioDeckBuildsRef.current);
        setSelectedDeckId(deck.id);
        setActiveView("slides");
        return {
          projectUpdatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          sceneRevision: deck.studioScene.revision,
          slideCount: build.slideCount,
          editablePowerPoint: { textNodeCount: build.textNodeCount, tableCount: build.tableCount, imageCount: build.imageCount, generatedComponentCount: build.generatedComponentCount },
          renderer: build.nativeRender.renderer,
          powerPointVersion: build.nativeRender.powerPointVersion,
          measurementAuthority: build.nativeMeasurement.authority,
          incrementalBuild: { rebuiltSlideCount, reusedSlideCount },
          readyForInAppReview: true,
          saved: false,
          exported: false,
          instruction: "Review the same central result in Slides. Location comments bind to each exact Studio slide revision. Export remains a separate user action.",
        };
      }
      if (request.operation === "run_deck_qualification") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio deck again before qualification.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("The requested deck does not have a complete central Studio design.");
        const build = studioDeckBuildsRef.current[deck.id];
        if (!build || build.sceneRevision !== deck.studioScene.revision) throw new Error("Build the exact current central Studio presentation before qualification.");
        const qualification = await qualifyStudioDeck(deck, build);
        setSelectedDeckId(deck.id);
        setActiveView("studio");
        return {
          projectUpdatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          sceneRevision: deck.studioScene.revision,
          report: qualification.report,
          issueLedger: { phase: qualification.report.issues.length ? "found-issues" : "ready", issueCount: qualification.report.issues.length, autoFixableCount: qualification.report.issues.filter((issue) => issue.repairRoute === "mcp-design").length },
          savedProject: false,
          exportedPresentation: false,
          instruction: qualification.report.status === "visual-review-required" ? "Open every candidate qualification image at full size, compare it with the source, and use bounded Studio operations for concrete visual issues. Objective checks do not prove the design is better." : "Route mcp-design issues through Studio layout/table/figure operations. Treat engine-code issues as reproducible product defects; do not hide them with slide-specific coordinate patches.",
        };
      }
      if (request.operation === "get_deck_qualification") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested central Studio deck is unavailable.");
        const qualification = studioDeckQualificationsRef.current[deck.id];
        if (!qualification || qualification.sceneRevision !== deck.studioScene.revision) throw new Error("Run qualification for the exact current Studio scene before reading its evidence ledger.");
        const history = (studioDeckQualificationHistoryRef.current[deck.id] ?? []).map((item) => ({ id: item.report.id, generatedAt: item.report.generatedAt, sceneRevision: item.sceneRevision, candidateSha256: item.candidateSha256, status: item.report.status, blockerIssues: item.report.totals.blockerIssues, majorIssues: item.report.totals.majorIssues, visualAcceptance: item.report.visualAcceptance.status }));
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, report: qualification.report, history, instruction: "Read the routed issue ledger and objective trend, inspect the native contact-sheet overview, then request both source/candidate full slides and any precise issue crops. Metadata alone cannot establish visual quality." };
      }
      if (request.operation === "get_qualification_slide") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("The requested central Studio deck is unavailable.");
        const qualification = studioDeckQualificationsRef.current[deck.id];
        if (!qualification || qualification.sceneRevision !== deck.studioScene.revision) throw new Error("Run qualification for the exact current Studio scene before requesting its slide images.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > qualification.report.totals.slides) throw new Error(`Choose an exported slide from 1 to ${qualification.report.totals.slides}.`);
        const representation = request.input.representation === "source" ? "source" as const : "candidate" as const;
        const slide = qualification.report.slides.find((item) => item.slideNumber === slideNumber);
        if (!slide) throw new Error("This qualification run does not contain the requested slide pair.");
        const evidenceSlideNumber = qualificationEvidenceSlideNumber(qualification.report, slideNumber, representation);
        if (!evidenceSlideNumber) throw new Error("This qualification run does not contain the requested evidence image.");
        const evidence = await desktop.readDeckQualificationSlide({ outputRoot: qualification.outputRoot, representation, slideNumber: evidenceSlideNumber });
        const view = request.input.view === "diagnostic-overlay" ? "diagnostic-overlay" as const : request.input.view === "issue-crop" ? "issue-crop" as const : "full" as const;
        const issueId = typeof request.input.issueId === "string" ? request.input.issueId : undefined;
        let image = { mimeType: evidence.mimeType as "image/png", data: bytesToBase64(bytesFrom(evidence.bytes)), width: representation === "source" ? slide.sourceImage.width : slide.candidateImage.width, height: representation === "source" ? slide.sourceImage.height : slide.candidateImage.height, rasterSha256: representation === "source" ? slide.sourceImage.sha256 : slide.candidateImage.sha256 };
        if (view !== "full") {
          if (view === "issue-crop" && !issueId) throw new Error("Choose an issue ID from the qualification ledger before requesting an issue crop.");
          const regions = qualificationInspectionRegions(qualification.report, slideNumber, view === "issue-crop" ? issueId : undefined);
          if (!regions.length) throw new Error("The requested qualification issue does not have a PowerPoint-native evidence region. Inspect the full slide pair instead.");
          const rendered = await inspectionRasterEvidence({ number: slideNumber, mimeType: "image/png", width: image.width, height: image.height, sha256: image.rasterSha256, bytes: bytesFrom(evidence.bytes) }, regions);
          const selected = view === "diagnostic-overlay" ? rendered.find((item) => item.kind === "diagnostic-overlay") : rendered.find((item) => item.id === issueId);
          if (!selected) throw new Error("Presentation Studio could not produce the requested qualification evidence view.");
          image = { mimeType: selected.mimeType, data: selected.data, width: selected.width, height: selected.height, rasterSha256: selected.sha256 };
        }
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          qualificationId: qualification.report.id,
          sceneRevision: qualification.sceneRevision,
          representation,
          view,
          issueId,
          slideNumber,
          rasterSha256: image.rasterSha256,
          fullSlideRasterSha256: representation === "source" ? slide.sourceImage.sha256 : slide.candidateImage.sha256,
          authoritative: true,
          issueIds: slide.issueIds,
          mimeType: image.mimeType,
          data: image.data,
          width: image.width,
          height: image.height,
          instruction: view === "issue-crop" ? "Use this native-pixel crop to diagnose the named issue, then use semantic Studio operations or a deterministic solver rather than estimating PowerPoint coordinates from pixels." : view === "diagnostic-overlay" ? "The rectangles are diagnostics, not slide artwork. Use them to locate issues, then inspect the clean full slide before judging quality." : representation === "candidate" ? "Judge the exact PowerPoint candidate at full size for hierarchy, alignment, spacing, table quality, figure clarity, template fidelity, deck consistency, and source intent." : "Use this exact source image to verify content, visual identity, meaningful encodings, and technical relationships before accepting a redesign.",
        };
      }
      if (request.operation === "get_qualification_contact_sheet") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("The requested central Studio deck is unavailable.");
        const qualification = studioDeckQualificationsRef.current[deck.id];
        const build = studioDeckBuildsRef.current[deck.id];
        if (!qualification || qualification.sceneRevision !== deck.studioScene.revision || !build || build.sceneRevision !== deck.studioScene.revision) throw new Error("Build and qualify the exact current central Studio presentation before requesting its contact sheet.");
        const representation = request.input.representation === "source" ? "source" as const : "candidate" as const;
        const render = representation === "candidate" ? build.nativeRender : await getOrBuildNativeRender(deck, "current", current);
        if (!render || render.status !== "ready" || !render.authoritative) throw new Error("The requested PowerPoint-native qualification overview is unavailable.");
        const expectedSourceSha256 = representation === "candidate" ? qualification.candidateSha256 : deck.sourceSha256;
        if (render.sourceSha256 !== expectedSourceSha256) throw new Error("The PowerPoint-native overview is stale. Rebuild and qualify the exact current candidate before review.");
        const sheet = await renderNativeContactSheet(render, Number(request.input.page ?? 1));
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, qualificationId: qualification.report.id, sceneRevision: qualification.sceneRevision, representation, authoritative: true, page: sheet.page, pageCount: sheet.pageCount, pageSize: sheet.pageSize, totalSlides: sheet.totalSlides, firstSlideNumber: sheet.firstSlideNumber, lastSlideNumber: sheet.lastSlideNumber, rasterSha256: sheet.sha256, mimeType: sheet.mimeType, data: bytesToBase64(sheet.bytes), width: sheet.width, height: sheet.height, instruction: "Use this native overview to find pacing, density, hierarchy, repetition, table, and consistency outliers. Then open each clean full slide pair and any issue crop before recording a visual verdict." };
      }
      if (request.operation === "record_deck_qualification_review") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested central Studio deck is unavailable.");
        const qualification = studioDeckQualificationsRef.current[deck.id];
        if (!qualification || qualification.sceneRevision !== deck.studioScene.revision || request.input.expectedSceneRevision !== qualification.sceneRevision) throw new Error("The Studio scene or qualification changed. Read the current qualification images again before recording review.");
        const report = recordDeckQualificationReviews(qualification.report, {
          qualificationId: String(request.input.qualificationId ?? ""),
          candidateSha256: String(request.input.candidateSha256 ?? ""),
          reviewer: "authorized-ai",
          reviews: request.input.reviews as Parameters<typeof recordDeckQualificationReviews>[1]["reviews"],
        });
        const finalized = await desktop.finalizeDeckQualification({ outputRoot: qualification.outputRoot, report });
        const nextQualification = { ...qualification, ...finalized, report };
        studioDeckQualificationsRef.current = { ...studioDeckQualificationsRef.current, [deck.id]: nextQualification };
        setStudioDeckQualifications(studioDeckQualificationsRef.current);
        setSelectedDeckId(deck.id);
        setActiveView("studio");
        const unresolved = report.visualAcceptance.revisionSlideCount + report.visualAcceptance.heldSlideCount;
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, qualificationId: report.id, sceneRevision: qualification.sceneRevision, candidateSha256: report.candidate.sha256, status: report.status, visualAcceptance: report.visualAcceptance, issueLedger: { phase: unresolved ? "found-issues" : report.status === "review-complete" ? "ready" : "rechecking", issueCount: unresolved, autoFixableCount: report.visualAcceptance.reviews.flatMap((review) => review.findings).filter((finding) => finding.repairRoute === "mcp-design").length }, savedProject: false, exportedPresentation: false, instruction: report.status === "review-complete" ? "Every slide is visually reviewed against its exact source/candidate pixels. This is draft qualification evidence, not formal ORNL approval; return the deck for human review and export only through the separate user action." : unresolved ? "Use the recorded slide findings and repair routes for one bounded design pass, rebuild the central deck, and run a new qualification. Do not patch coordinates from pixels or repeat an unchanged candidate." : "Continue reviewing the remaining exact source/candidate slide pairs." };
      }
      if (request.operation === "list_resources") {
        const allowed = current.resources.filter((resource) => resource.mcpAccess !== "none");
        return {
          updatedAt: current.project.updatedAt,
          totalResourceCount: current.resources.length,
          resources: allowed.map((resource) => ({
            id: resource.id,
            name: resource.name,
            kind: resource.kind ?? "other",
            mediaType: resource.mediaType,
            byteLength: resource.byteLength,
            sha256: resource.sha256,
            roles: resource.roles,
            support: resource.support ?? [],
            processing: resource.processing ? { status: resource.processing.status, summary: resource.processing.summary, warnings: resource.processing.warnings } : undefined,
            extractedTextAvailable: Boolean(resource.derivatives?.some((derivative) => derivative.kind === "extracted-text")),
            permission: resource.mcpAccess,
          })),
        };
      }
      if (request.operation === "get_resource_text") {
        const resource = current.resources.find((item) => item.id === request.input.resourceId);
        if (!resource) throw new Error("The requested Resource is not in this project.");
        const page = resourceTextPage(resource, Number(request.input.offset ?? 0), Number(request.input.maximumCharacters ?? 20_000));
        return {
          updatedAt: current.project.updatedAt,
          resource: { id: resource.id, name: resource.name, sha256: resource.sha256, roles: resource.roles, processing: resource.processing },
          ...page,
          instruction: page.nextOffset === undefined ? "The complete embedded extracted-text derivative has been read. Preserve names, numbers, units, qualifications, and attribution; cite exact excerpts when creating slides." : `Continue reading this Resource at offset ${page.nextOffset} before claiming complete coverage.`,
        };
      }
      if (request.operation === "get_resource_preview") {
        const resource = current.resources.find((item) => item.id === request.input.resourceId);
        if (!resource) throw new Error("The requested Resource is not in this project.");
        if (resource.mcpAccess !== "preview") throw new Error("This image is not available. Turn on AI access and re-list the project Resources.");
        const preview = await boundedResourceImagePreview(resource);
        return {
          updatedAt: current.project.updatedAt,
          resource: { id: resource.id, name: resource.name, sha256: resource.sha256, roles: resource.roles, roleLabel: resource.roles.includes("concept-reference") ? "concept-only" : "reference" },
          mimeType: preview.mimeType,
          data: bytesToBase64(preview.bytes),
          width: preview.width,
          height: preview.height,
          rasterSha256: preview.sha256,
        };
      }
      if (request.operation === "create_studio_presentation") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read app status and the authorized Resources again before creating a presentation.");
        if (!templateCatalog) throw new Error("Install the approved ORNL Template Pack before creating a new presentation.");
        const slides = request.input.slides as NewStudioPresentationInput["slides"];
        for (const [index, slide] of slides.entries()) {
          for (const reference of slide.sourceReferences) {
            const resource = current.resources.find((item) => item.id === reference.resourceId);
            if (!resource) throw new Error(`Slide ${index + 1} references a Resource that is not in the current project.`);
            assertExactResourceExcerpt(resource, reference.exactExcerpt ?? "");
          }
          for (const resourceId of slide.imageResourceIds ?? []) {
            const resource = current.resources.find((item) => item.id === resourceId);
            if (!resource || resource.kind !== "image") throw new Error(`Slide ${index + 1} references an unavailable image Resource.`);
            if (resource.mcpAccess !== "preview") throw new Error(`${resource.name} is not available as an image preview. Turn on AI access and re-list the project Resources.`);
          }
          if (slide.table && slide.table.rows.some((row) => row.length !== slide.table!.headers.length)) throw new Error(`Slide ${index + 1} has a table row whose cell count does not match its ${slide.table.headers.length} headers.`);
        }
        const deckId = crypto.randomUUID();
        const sceneDraft = await createNewStudioPresentationScene({
          deckId,
          name: String(request.input.name ?? "Untitled presentation"),
          communicationJob: String(request.input.communicationJob ?? "").trim(),
          expression: request.input.expression === "restrained" || request.input.expression === "expressive" ? request.input.expression : "balanced",
          slides,
        }, current.resources, templateCatalog);
        const resourceMedia = await studioResourceMedia(sceneDraft, current.resources);
        const templateRender = await getOrBuildTemplateNativeRender();
        if (templateRender.status !== "ready" || !templateRender.authoritative) throw new Error(templateRender.warnings[0] ?? "PowerPoint-native rendering of the approved ORNL Template Pack is required to create a protected title slide.");
        const templateLayoutRasters = Object.fromEntries(templateCatalog.layouts.map((layout, index) => {
          const rendered = templateRender.slides.find((slide) => slide.number === index + 1);
          return [layout.id, rendered ? { data: `data:${rendered.mimeType};base64,${bytesToBase64(bytesFrom(rendered.bytes))}`, width: rendered.width, height: rendered.height } : undefined];
        }).filter((entry): entry is [string, { data: string; width: number; height: number }] => Boolean(entry[1])));
        const initialComposition = await buildStudioCompositionPptx(sceneDraft, {
          catalog: catalogWithStudioResources(undefined, resourceMedia, sceneDraft),
          templateCatalog,
          templateLayoutRasters,
          strict: true,
          title: String(request.input.name ?? "Untitled presentation"),
        });
        const sourceName = `${cleanFileStem(String(request.input.name ?? "Untitled presentation"))}.pptx`;
        const generatedResource = resourceWithAiSessionAccess(await processResourceInput({ name: sourceName, mediaType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", bytes: initialComposition.bytes }), true);
        const audit = await auditPptx(initialComposition.bytes);
        if (audit.slideCount !== slides.length) throw new Error(`The native Studio compiler created ${audit.slideCount} slides for a ${slides.length}-slide plan.`);
        const adoptedAt = new Date().toISOString();
        const deckBase: DeckJob = {
          id: deckId,
          name: sourceName,
          sourceResourceId: generatedResource.id,
          sourceSha256: generatedResource.sha256,
          operationScope: "compose",
          templateClassification: "current-ornl",
          targetTemplateId: PRESENTATION_DESIGN_STANDARD.defaults.template.id,
          targetTemplateConfirmedAt: adoptedAt,
          targetTemplateDecisionSource: "automatic-default",
          designProfile: createOrnlDesignProfile("automatic-default", adoptedAt),
          status: "ready-for-cleanup",
          audit,
          protectedSlideNumbers: [1],
        };
        const scene = compilePresentationScene({ ...deckBase, audit });
        const studioScene = bindNewStudioSceneToGeneratedPowerPoint(sceneDraft, generatedResource.sha256, audit);
        const deck: DeckJob = { ...deckBase, scene, studioScene };
        const next = touchProject({
          ...current,
          project: { ...current.project, type: "new-presentation" as const },
          settings: { ...current.settings, contentPolicy: "source-grounded-generative" as const, defaultOperationScope: "compose" as const },
          resources: [...current.resources, generatedResource],
          decks: [...current.decks, deck],
        }, "studio-presentation-created", `Created ${slides.length} source-grounded ORNL slide${slides.length === 1 ? "" : "s"} in the native Studio scene from Resources automatically shared by the active AI session; no project or presentation was saved outside the app.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(1);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          deck: { id: deck.id, name: deck.name, sourceSha256: deck.sourceSha256 },
          slideCount: slides.length,
          sceneRevision: studioScene.revision,
          nativeJson: { schema: studioScene.schema, version: studioScene.version, designSystem: studioScene.designSystem, resourceBindingCount: studioScene.slides.reduce((sum, slide) => sum + (slide.resourceBindings?.length ?? 0), 0) },
          titleSlide: { slideNumber: 1, protected: true, sourcePreserved: true },
          editablePowerPoint: { embeddedSourceCreated: true, textNodeCount: initialComposition.textNodeCount, tableCount: initialComposition.tableCount, imageCount: initialComposition.imageCount },
          readyForInAppReview: true,
          savedProject: false,
          exportedPresentation: false,
          instruction: "Inspect the central Studio design and the PowerPoint-native slide pixels. Use stage_studio_web_design or refine tools for revisions, then call build_studio_presentation. Saving or exporting remains a separate human action.",
        };
      }
      if (request.operation === "create_studio_visual_need") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio scene again before creating a visual-direction brief.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current Studio source scene.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error("This approved ORNL template slide is sacred and cannot enter the visual-concept queue.");
        const catalog = await getOrBuildSlideCatalog(deck, current);
        const baseScene = deck.studioScene ?? compileStudioWebScene(deck, catalog);
        const rawTarget = request.input.targetSlot && typeof request.input.targetSlot === "object" ? request.input.targetSlot as Record<string, unknown> : undefined;
        const scene = createStudioVisualNeed(baseScene, slideNumber, {
          type: String(request.input.type ?? "layout-concept") as StudioVisualNeed["type"],
          reason: String(request.input.reason ?? ""),
          communicationJob: String(request.input.communicationJob ?? ""),
          expression: String(request.input.expression ?? "balanced") as StudioVisualNeed["expression"],
          approvedInfluences: Array.isArray(request.input.approvedInfluences) ? request.input.approvedInfluences.map(String) as StudioVisualNeed["approvedInfluences"] : undefined,
          disclosurePolicy: String(request.input.disclosurePolicy ?? "abstract-structure-only") as StudioVisualNeed["disclosurePolicy"],
          approvedContentSummary: typeof request.input.approvedContentSummary === "string" ? request.input.approvedContentSummary : undefined,
          targetSlot: rawTarget ? {
            role: typeof rawTarget.role === "string" ? rawTarget.role as StudioVisualNeed["targetSlot"]["role"] : undefined,
            aspectRatio: typeof rawTarget.aspectRatio === "string" ? rawTarget.aspectRatio as StudioVisualNeed["targetSlot"]["aspectRatio"] : undefined,
            placementNotes: typeof rawTarget.placementNotes === "string" ? rawTarget.placementNotes : undefined,
          } : undefined,
        });
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-visual-need-created", `Created a local visual-direction brief for slide ${slideNumber} of ${deck.name}; no model was called and source bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        const need = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.visualNeeds?.at(-1);
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, need, imageModelCalled: false, sourceChanged: false, saved: false, exported: false };
      }
      if (request.operation === "list_studio_visual_needs") {
        const deckId = typeof request.input.deckId === "string" ? request.input.deckId : undefined;
        const needs = current.decks.filter((deck) => !deckId || deck.id === deckId).flatMap((deck) => {
          if (!deck.studioScene) return [];
          return deck.studioScene.slides.flatMap((slide) => (slide.visualNeeds ?? []).map((need) => ({
            deckId: deck.id,
            deckName: deck.name,
            slideNumber: slide.slideNumber,
            id: need.id,
            type: need.type,
            status: need.status,
            expression: need.expression,
            disclosurePolicy: need.disclosurePolicy,
            brandExpression: need.brandExpression,
            targetSlot: need.targetSlot,
            sourceTextHash: need.sourceTextHash,
            linkedConceptReferenceId: need.linkedConceptReferenceId,
            conceptAvailable: Boolean(need.linkedConceptReferenceId && slide.conceptReferences?.some((reference) => reference.id === need.linkedConceptReferenceId)),
            updatedAt: need.updatedAt,
          })));
        });
        return { updatedAt: current.project.updatedAt, needs };
      }
      if (request.operation === "get_studio_visual_need_brief") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested deck does not have a persisted Studio scene.");
        const slideNumber = Number(request.input.slideNumber);
        const slide = deck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
        const need = slide?.visualNeeds?.find((item) => item.id === request.input.visualNeedId);
        if (!slide || !need) throw new Error("The requested visual need is not present on this Studio slide.");
        if (need.sourceTextHash !== slide.sourceTextHash) throw new Error("The visual brief is stale because its source-content binding changed.");
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          slideNumber,
          sceneRevision: deck.studioScene.revision,
          need: {
            id: need.id, type: need.type, status: need.status, sourceTextHash: need.sourceTextHash, expression: need.expression, approvedInfluences: need.approvedInfluences,
            disclosurePolicy: need.disclosurePolicy, approvedContentSummary: need.approvedContentSummary, brandExpression: need.brandExpression, structureInventory: need.structureInventory, targetSlot: need.targetSlot, promptPackage: need.promptPackage,
            linkedConceptReferenceId: need.linkedConceptReferenceId, createdAt: need.createdAt, updatedAt: need.updatedAt,
          },
          instruction: "Use only the promptPackage within its disclosure boundary. Do not add source pixels or exact content that the package excludes. Import any resulting raster into Resources, grant Preview, attach it as concept-only, then reconstruct approved visual characteristics with exact editable source content.",
        };
      }
      if (request.operation === "hold_studio_visual_need") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the visual-needs queue again before holding this brief.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested deck does not have a persisted Studio scene.");
        const slideNumber = Number(request.input.slideNumber);
        const visualNeedId = String(request.input.visualNeedId ?? "");
        const scene = holdStudioVisualNeed(deck.studioScene, slideNumber, visualNeedId, String(request.input.note ?? ""));
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-visual-need-held", `Held a local visual-direction brief on slide ${slideNumber} of ${deck.name}; no Resource or source file was deleted.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, visualNeedId, deleted: false, sourceChanged: false };
      }
      if (request.operation === "attach_studio_concept_reference") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio scene and Resources again before attaching a concept reference.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current Studio source scene.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error("This approved ORNL template slide is sacred and cannot use an Image Gen or external concept reference.");
        const resource = current.resources.find((item) => item.id === request.input.resourceId);
        if (!resource) throw new Error("The requested concept image is not embedded in this project.");
        if (resource.mcpAccess !== "preview") throw new Error("This image is not available as a preview. Turn on AI access and re-list the project Resources.");
        const catalog = await getOrBuildSlideCatalog(deck, current);
        const baseScene = deck.studioScene ?? compileStudioWebScene(deck, catalog);
        const rawBlueprint = request.input.blueprint && typeof request.input.blueprint === "object" ? request.input.blueprint as Record<string, unknown> : {};
        const scene = attachStudioConceptReference(baseScene, slideNumber, resource, {
          origin: String(request.input.origin ?? "other") as "imagegen" | "human-reference" | "other",
          approvedInfluences: Array.isArray(request.input.approvedInfluences) ? request.input.approvedInfluences.map(String) as Array<"composition" | "visual-hierarchy" | "negative-space" | "color-balance" | "figure-concept" | "image-treatment" | "visual-rhythm"> : [],
          blueprint: {
            summary: String(rawBlueprint.summary ?? ""),
            zones: Array.isArray(rawBlueprint.zones) ? rawBlueprint.zones.map((zone) => {
              const item = zone && typeof zone === "object" ? zone as Record<string, unknown> : {};
              return { id: String(item.id ?? ""), role: String(item.role ?? "other") as "title" | "primary-visual" | "supporting-evidence" | "caption" | "footer-safe" | "other", x: Number(item.x), y: Number(item.y), width: Number(item.width), height: Number(item.height) };
            }) : [],
            styleNotes: Array.isArray(rawBlueprint.styleNotes) ? rawBlueprint.styleNotes.map(String) : [],
            reconstructionNotes: Array.isArray(rawBlueprint.reconstructionNotes) ? rawBlueprint.reconstructionNotes.map(String) : [],
          },
          provenance: request.input.provenance && typeof request.input.provenance === "object" ? {
            model: typeof (request.input.provenance as Record<string, unknown>).model === "string" ? String((request.input.provenance as Record<string, unknown>).model) : undefined,
            promptSummary: typeof (request.input.provenance as Record<string, unknown>).promptSummary === "string" ? String((request.input.provenance as Record<string, unknown>).promptSummary) : undefined,
            generatedAt: typeof (request.input.provenance as Record<string, unknown>).generatedAt === "string" ? String((request.input.provenance as Record<string, unknown>).generatedAt) : undefined,
          } : undefined,
          visualNeedId: typeof request.input.visualNeedId === "string" ? request.input.visualNeedId : undefined,
        });
        const next = touchProject({
          ...current,
          resources: current.resources.map((item) => item.id === resource.id ? { ...item, roles: [...new Set([...item.roles, "concept-reference" as const])] } : item),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene, operationScope: "reflow", proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "studio-concept-reference-attached", `Attached a concept-only image Resource to slide ${slideNumber} of ${deck.name}; source content and PowerPoint bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("studio");
        const reference = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.conceptReferences?.find((item) => item.resourceId === resource.id);
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, reference, saved: false, exported: false, instruction: "Read this concept alongside the exact Studio scene and original PowerPoint render. Reconstruct only its approved visual characteristics with exact source content; then build and inspect the editable PowerPoint result." };
      }
      if (request.operation === "get_studio_concept_reference") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested deck does not have a persisted Studio scene.");
        const slideNumber = Number(request.input.slideNumber);
        const slide = deck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
        const reference = slide?.conceptReferences?.find((item) => item.id === request.input.referenceId);
        if (!slide || !reference) throw new Error("The requested concept reference is not attached to this Studio slide.");
        if (reference.sourceTextHash !== slide.sourceTextHash) throw new Error("The concept reference is stale because the source content binding changed.");
        const resource = current.resources.find((item) => item.id === reference.resourceId && item.sha256 === reference.resourceSha256);
        if (!resource) throw new Error("The concept Resource is missing or no longer matches its recorded identity.");
        if (resource.mcpAccess !== "preview") throw new Error("This concept image is not available. Turn on AI access and re-list the project Resources.");
        const preview = await boundedResourceImagePreview(resource);
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slideNumber, sceneRevision: deck.studioScene.revision, reference, mimeType: preview.mimeType, data: bytesToBase64(preview.bytes), width: preview.width, height: preview.height, rasterSha256: preview.sha256 };
      }
      if (request.operation === "reconstruct_studio_concept") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio scene and concept reference again before reconstruction.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio scene before reconstructing a concept.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read the exact current revision before reconstruction.");
        const slideNumber = Number(request.input.slideNumber);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error("This approved ORNL template slide is sacred and cannot be reconstructed from a visual concept.");
        const requestedRecipe = request.input.recipe ? String(request.input.recipe) as StudioLayoutRecipe : undefined;
        const reconstruction = reconstructStudioConcept(deck.studioScene, slideNumber, String(request.input.referenceId ?? ""), requestedRecipe);
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: reconstruction.scene, proposal: undefined } : item) }, "mcp-studio-concept-reconstructed", `AI reconstructed approved concept zones as editable Studio content on slide ${slideNumber} of ${deck.name}; exact source content and source PowerPoint bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          sceneRevision: reconstruction.scene.revision,
          slideNumber,
          referenceId: reconstruction.referenceId,
          recipe: reconstruction.recipe,
          mappedNodeIds: reconstruction.mappedNodeIds,
          unmappedNodeIds: reconstruction.unmappedNodeIds,
          designImpact: analyzeStudioDesignImpact(reconstruction.scene.slides.find((slide) => slide.slideNumber === slideNumber)!),
          diagnostics: reconstruction.diagnostics,
          applied: false,
          saved: false,
          instruction: "The concept zones are now an editable Studio composition with exact source-bound content. Build the export result, inspect original/concept/export pixels together, and record the bounded native visual critique before review or save.",
        };
      }
      if (request.operation === "remove_studio_concept_reference") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the Studio scene again before detaching the concept reference.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.studioScene) throw new Error("The requested deck does not have a persisted Studio scene.");
        const slideNumber = Number(request.input.slideNumber);
        const scene = removeStudioConceptReference(deck.studioScene, slideNumber, String(request.input.referenceId ?? ""));
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-concept-reference-detached", `Detached a concept reference from slide ${slideNumber} of ${deck.name}; the embedded Resource and source PowerPoint remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("studio");
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, referenceId: request.input.referenceId, resourceRemoved: false, sourceChanged: false };
      }
      if (request.operation === "get_deck_audit") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, status: deck.status, targetTemplateId: deck.targetTemplateId }, audit: { ...deck.audit, slides: deck.audit.slides.map(({ text: _text, title: _title, ...slide }) => slide), tables: deck.audit.tables.map((table) => ({ ...table, cells: table.cells?.map(({ text: _text, ...cell }) => cell) })), pictures: deck.audit.pictures.map(({ name: _name, description: _description, relationshipId: _relationshipId, ...picture }) => picture), textBoxes: deck.audit.textBoxes.map(({ text: _text, ...textBox }) => textBox) } };
      }
      if (request.operation === "get_slide_measurements") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        const slideNumber = Number(request.input.slideNumber);
        const representation = request.input.representation === "proposal" ? "proposal" as const : "current" as const;
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const measurement = await getOrBuildNativeMeasurement(deck, representation, current);
        const baseline = representation === "proposal" ? await getOrBuildNativeMeasurement(deck, "current", current) : undefined;
        const metrics = calculateDesignMetrics(deck, measurement, baseline).slides.find((item) => item.slideNumber === slideNumber);
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slideNumber, representation, measurement: { ...measurement, objects: measurement.objects.filter((object) => object.slideNumber === slideNumber) }, metrics, instruction: "Use these PowerPoint-native measurements for point geometry. Use an alignment, distribution, or table solver instead of guessing coordinates from pixels." };
      }
      if (request.operation === "get_slide_design_context") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const start = Number(request.input.startSlide);
        const end = Number(request.input.endSlide);
        if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end < start || end - start > 9) throw new Error("Request 1 to 10 consecutive slides with a valid start and end.");
        if (end > deck.audit.slideCount) throw new Error(`The requested range exceeds this deck's ${deck.audit.slideCount} slides.`);
        const slides = deck.audit.slides.filter((slide) => slide.number >= start && slide.number <= end).map((slide) => ({
          id: slide.id,
          number: slide.number,
          title: slide.title,
          exactVisibleText: slide.text,
          textHash: slide.textHash,
          textRunCount: slide.textRunCount,
          fonts: slide.fonts,
          fontSizes: slide.fontSizes,
          objects: { tables: slide.tableCount, pictures: slide.pictureCount, charts: slide.chartCount, connectors: slide.connectorCount, comments: slide.commentCount },
          warnings: slide.warnings,
          findings: deck.audit?.findings.filter((finding) => finding.slideNumber === slide.number) ?? [],
          textBoxes: (deck.audit?.textBoxes ?? []).filter((textBox) => textBox.slideNumber === slide.number).map((textBox) => ({ ...textBox, exactText: textBox.text })),
          editableObjects: (deck.audit?.editableObjects ?? []).filter((object) => object.slideNumber === slide.number).map((object) => ({
            ...object,
            scene: deck.scene?.objects.find((sceneObject) => sceneObject.id === object.id),
            exactText: (deck.audit?.textBoxes ?? []).find((textBox) => textBox.slideNumber === object.slideNumber && textBox.shapeId === object.shapeId)?.text,
            geometryInches: { x: object.geometry.x / 914_400, y: object.geometry.y / 914_400, width: object.geometry.width / 914_400, height: object.geometry.height / 914_400 },
          })),
          geometryChecks: (deck.audit?.layoutReviews ?? []).filter((review) => review.slideNumber === slide.number),
          alignmentRepairs: (deck.audit?.alignmentRepairs ?? []).filter((repair) => repair.slideNumber === slide.number),
          stagedGeometry: deck.proposal?.changes.flatMap((change) => change.kind === "geometry" ? change.geometryCommands ?? [] : []).filter((command) => command.slideNumber === slide.number) ?? [],
        }));
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId }, range: { start, end }, slides, preservationRequired: true };
      }
      if (request.operation === "get_cleanup_rule_profile") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId, targetTemplateDecisionSource: deck.targetTemplateDecisionSource }, resolvedProfile: deck.designProfile ?? null, standard: PRESENTATION_DESIGN_STANDARD, routineApprovalsRequired: false };
      }
      if (request.operation === "get_pending_proposal_manifest") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.proposal || deck.proposal.status !== "pending") throw new Error("The requested deck does not have a pending proposal.");
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, sourceSha256: deck.sourceSha256 }, proposal: deck.proposal };
      }
      if (request.operation === "get_slide_render") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const requestedRepresentation = typeof request.input.representation === "string" ? request.input.representation : "current";
        const representation: "source" | "current" | "proposal" | "export" = ["source", "current", "proposal", "export"].includes(requestedRepresentation) ? requestedRepresentation as "source" | "current" | "proposal" | "export" : "current";
        const proposalRepresentation = representation === "proposal" || representation === "export";
        const catalog = proposalRepresentation ? await getOrBuildProposalCatalog(deck, current) : await getOrBuildSlideCatalog(deck, current);
        const slide = catalog.slides.find((item) => item.number === slideNumber);
        if (!slide) throw new Error("The requested slide render is unavailable.");
        const nativeRender = await getOrBuildNativeRender(deck, representation, current);
        const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((item) => item.number === slideNumber) : undefined;
        if (nativeSlide && nativeRender) {
          return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slide: { id: deck.audit.slides.find((item) => item.number === slideNumber)?.id ?? slide.id, number: slide.number, title: slide.title }, representation, proposalId: proposalRepresentation ? deck.proposal?.id : undefined, renderer: nativeRender.renderer, pipeline: nativeRender.pipeline, powerPointVersion: nativeRender.powerPointVersion, rasterSha256: nativeSlide.sha256, sourceSha256: nativeRender.sourceSha256, authoritative: nativeRender.authoritative, qaNote: "This image was rendered locally by Microsoft PowerPoint from the exact requested PPTX revision and is authoritative for visual design review.", mimeType: nativeSlide.mimeType, data: bytesToBase64(bytesFrom(nativeSlide.bytes)), width: nativeSlide.width, height: nativeSlide.height };
        }
        const jpeg = await slidePreviewJpeg(catalog, slide, 1200, presentationFontCss);
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slide: { id: deck.audit.slides.find((item) => item.number === slideNumber)?.id ?? slide.id, number: slide.number, title: slide.title }, representation, proposalId: proposalRepresentation ? deck.proposal?.id : undefined, renderer: catalog.renderer, authoritative: false, nativeRenderStatus: nativeRender?.status ?? "unavailable", nativeRenderWarnings: nativeRender?.warnings ?? ["PowerPoint-native rendering is not available in this browser session."], qaNote: "Diagnostic OOXML approximation only. Do not use this image for final visual acceptance; retry when a PowerPoint-native render is available.", mimeType: "image/jpeg", ...jpeg };
      }
      if (request.operation === "get_slide_render_comparison") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.proposal) throw new Error("Stage a proposal before requesting a Current/Proposal native comparison.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const [currentRender, proposalRender] = await Promise.all([getOrBuildNativeRender(deck, "current", current), getOrBuildNativeRender(deck, "proposal", current)]);
        if (currentRender?.status !== "ready" || proposalRender?.status !== "ready") throw new Error("Authoritative Microsoft PowerPoint Current and Proposal renders are both required for comparison.");
        const currentSlide = currentRender.slides.find((slide) => slide.number === slideNumber);
        const proposalSlide = proposalRender.slides.find((slide) => slide.number === slideNumber);
        if (!currentSlide || !proposalSlide) throw new Error("PowerPoint did not return both requested slide images.");
        const comparison = await compareNativeSlideRenders(currentSlide, proposalSlide);
        return {
          updatedAt: current.project.updatedAt,
          deck: { id: deck.id, name: deck.name },
          slide: { id: deck.audit.slides.find((slide) => slide.number === slideNumber)?.id, number: slideNumber },
          proposal: { id: deck.proposal.id, status: deck.proposal.status, designDecision: deck.proposal.designDecision },
          renderer: currentRender.renderer,
          powerPointVersion: currentRender.powerPointVersion,
          authoritative: true,
          currentRasterSha256: currentSlide.sha256,
          proposalRasterSha256: proposalSlide.sha256,
          metrics: comparison.metrics,
          qaNote: "Pixel difference proves that PowerPoint rendered a change; it does not prove the proposal is better. Inspect both compositions and revise or reject visual regressions.",
          mimeType: comparison.mimeType,
          data: bytesToBase64(comparison.bytes),
          width: comparison.width,
          height: comparison.height,
        };
      }
      if (request.operation === "reject_design_proposal") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the current proposal comparison before rejecting this draft.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.proposal || deck.proposal.status !== "pending") throw new Error("The requested deck does not have a pending design proposal.");
        if (deck.proposal.id !== request.input.proposalId) throw new Error("The pending proposal changed. Read a fresh Current/Proposal comparison before rejecting it.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const rationale = String(request.input.rationale ?? "").trim();
        const [currentRender, proposalRender] = await Promise.all([getOrBuildNativeRender(deck, "current", current), getOrBuildNativeRender(deck, "proposal", current)]);
        if (currentRender?.status !== "ready" || proposalRender?.status !== "ready") throw new Error("AI self-rejection requires authoritative Microsoft PowerPoint Current and Proposal renders.");
        const currentSlide = currentRender.slides.find((slide) => slide.number === slideNumber);
        const proposalSlide = proposalRender.slides.find((slide) => slide.number === slideNumber);
        if (!currentSlide || !proposalSlide) throw new Error("PowerPoint did not return both requested slide images.");
        const comparison = await compareNativeSlideRenders(currentSlide, proposalSlide);
        const rejectedProposal: CleanupProposal = {
          ...deck.proposal,
          status: "rejected",
          designReview: {
            decision: "rejected",
            actor: "ai",
            rationale: rationale.slice(0, 1_000),
            reviewedAt: new Date().toISOString(),
            evidence: {
              slideNumber,
              renderer: "powerpoint-native",
              currentRasterSha256: currentSlide.sha256,
              proposalRasterSha256: proposalSlide.sha256,
              changedPixelRatio: comparison.metrics.changedPixelRatio,
            },
          },
        };
        const next = touchProject({
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal: rejectedProposal, status: "audited" as const } : item),
        }, "mcp-design-proposal-rejected", `AI rejected proposal ${deck.proposal.id.slice(0, 8)} for ${deck.name} after PowerPoint-native comparison; source bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return {
          proposal: { id: rejectedProposal.id, status: rejectedProposal.status, designReview: rejectedProposal.designReview },
          projectUpdatedAt: next.project.updatedAt,
          sourceChanged: false,
          applied: false,
          saved: false,
          instruction: "Read a fresh design work order before staging another attempt. The rejected draft and its native comparison evidence remain visible in Review.",
        };
      }
      if (request.operation === "record_proposal_visual_critique") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Inspect the current Proposal again before recording a visual critique.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene || !deck.proposal || deck.proposal.status !== "pending") throw new Error("The requested deck does not have a pending source-bound design proposal.");
        if (deck.proposal.id !== request.input.proposalId) throw new Error("The pending proposal changed. Inspect the fresh Proposal pixels before critiquing it.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const [currentRender, proposalRender, currentMeasurement, proposalMeasurement] = await Promise.all([
          getOrBuildInspectionRender(deck, "current", current),
          getOrBuildInspectionRender(deck, "proposal", current),
          getOrBuildNativeMeasurement(deck, "current", current),
          getOrBuildNativeMeasurement(deck, "proposal", current),
        ]);
        const currentSlide = currentRender.slides.find((slide) => slide.number === slideNumber);
        const proposalSlide = proposalRender.slides.find((slide) => slide.number === slideNumber);
        if (!currentSlide || !proposalSlide || currentMeasurement.authority !== "powerpoint-native" || proposalMeasurement.authority !== "powerpoint-native") throw new Error("AI visual critique requires authoritative Current and Proposal PowerPoint pixels and measurements.");
        const expectedInspectionRevision = `${current.project.updatedAt}:${deck.scene.revision}:slide-${slideNumber}:raster-${proposalSlide.sha256}:measurement-${proposalMeasurement.revision}`;
        if (String(request.input.inspectionRevision ?? "") !== expectedInspectionRevision) throw new Error("The Proposal inspection evidence is stale. Read get_slide_inspection_packet for the current proposal and critique that exact revision.");
        const currentMetric = calculateDesignMetrics(deck, currentMeasurement).slides.find((item) => item.slideNumber === slideNumber)!;
        const proposalMetric = calculateDesignMetrics(deck, proposalMeasurement, currentMeasurement).slides.find((item) => item.slideNumber === slideNumber)!;
        const metricEvaluation = metricsImproved(currentMetric, proposalMetric);
        const comparison = await compareNativeSlideRenders(currentSlide, proposalSlide);
        const requestedVerdict = String(request.input.verdict ?? "revise") as "better" | "revise" | "reject";
        const priorHistory = deck.proposal.visualIteration?.history ?? [];
        const rawIntentReview = request.input.intentReview && typeof request.input.intentReview === "object" ? request.input.intentReview as Record<string, unknown> : {};
        const intentReview = {
          status: String(rawIntentReview.status ?? "needs-review") as "pass" | "needs-review",
          exactTextPreserved: Boolean(rawIntentReview.exactTextPreserved),
          sourceVisualsPreserved: Boolean(rawIntentReview.sourceVisualsPreserved),
          relationshipsPreserved: String(rawIntentReview.relationshipsPreserved ?? "unverified") as "yes" | "not-applicable" | "unverified",
          summary: String(rawIntentReview.summary ?? "Original message-intent review was not supplied."),
        };
        const iteration = decideVisualIteration({ priorHistory, requestedVerdict, rationale: String(request.input.rationale ?? ""), slideNumber, inspectionRevision: expectedInspectionRevision, currentRasterSha256: currentSlide.sha256, proposalRasterSha256: proposalSlide.sha256, changedPixelRatio: comparison.metrics.changedPixelRatio, improvements: metricEvaluation.improvements, regressions: metricEvaluation.regressions, intentReview });
        const { verdict, rejected } = iteration;
        const rationale = iteration.entry.rationale;
        const history = [...priorHistory, iteration.entry];
        const updatedProposal: CleanupProposal = {
          ...deck.proposal,
          status: rejected ? "rejected" : "pending",
          visualIteration: { maxAttempts: 3, history },
          designReview: rejected ? {
            decision: "rejected",
            actor: "ai",
            rationale,
            reviewedAt: new Date().toISOString(),
            evidence: { slideNumber, renderer: "powerpoint-native", currentRasterSha256: currentSlide.sha256, proposalRasterSha256: proposalSlide.sha256, changedPixelRatio: comparison.metrics.changedPixelRatio },
          } : deck.proposal.designReview,
        };
        const next = touchProject({
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal: updatedProposal, status: rejected ? "audited" as const : "proposal-ready" as const } : item),
        }, rejected ? "mcp-visual-loop-rejected" : verdict === "revise" ? "mcp-visual-loop-revision-requested" : "mcp-visual-loop-qualified", `AI recorded visual iteration ${iteration.entry.attempt}/3 for slide ${slideNumber} of ${deck.name}: ${verdict}. Source bytes and accepted state remain unchanged.`);
        updatedProposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return {
          projectUpdatedAt: next.project.updatedAt,
          proposal: { id: updatedProposal.id, status: updatedProposal.status },
          critique: history.at(-1),
          requestedVerdict,
          recordedVerdict: verdict,
          nativeEvidence: { renderer: "powerpoint-native", changedPixelRatio: comparison.metrics.changedPixelRatio, currentMetric, proposalMetric },
          applied: false,
          saved: false,
          instruction: verdict === "better" ? "The draft survived deterministic regression checks and the AI visual critique. It remains pending for human review; it was not applied or exported." : verdict === "revise" ? "Stage a materially different semantic operation, then inspect the new native Proposal revision before attempt 2 or 3." : "The draft is rejected and the source remains unchanged. Begin a fresh design approach or leave it for human review.",
        };
      }
      if (request.operation === "list_design_threads") {
        const deckId = typeof request.input.deckId === "string" ? request.input.deckId : undefined;
        const threads = current.designThreads.filter((thread) => (!deckId || thread.deckId === deckId) && ["submitted", "needs-reanchor"].includes(thread.status));
        return { updatedAt: current.project.updatedAt, threads: threads.map((thread) => ({ id: thread.id, deckId: thread.deckId, slideId: thread.slideId, slideNumber: thread.slideNumber, outputSlideNumber: thread.outputSlideNumber, baseRevision: thread.baseRevision, anchor: thread.anchor, comment: thread.comment, status: thread.status, createdAt: thread.createdAt, submittedAt: thread.submittedAt })) };
      }
      if (request.operation === "get_design_thread") {
        const thread = current.designThreads.find((item) => item.id === request.input.threadId);
        if (!thread || !["submitted", "needs-reanchor"].includes(thread.status)) throw new Error("The requested design thread is no longer active or was not submitted to AI in this project.");
        const deck = current.decks.find((item) => item.id === thread.deckId);
        const slide = deck?.audit?.slides.find((item) => item.id === thread.slideId || item.number === thread.slideNumber);
        return { updatedAt: current.project.updatedAt, thread, deck: deck ? { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId } : null, slide: slide ? { id: slide.id, number: slide.number, outputSlideNumber: thread.outputSlideNumber, title: slide.title, textHash: slide.textHash, objects: { tables: slide.tableCount, pictures: slide.pictureCount, charts: slide.chartCount } } : null, instruction: thread.outputSlideNumber ? "This comment is anchored to one exact materialized output slide. Read the current source-scene revision, rebuild it, and inspect that output number in the PowerPoint-native preview or qualification evidence before staging a bounded fix." : "Read the current revision and get_slide_render before staging a bounded fix. Do not guess if the anchor no longer maps unambiguously." };
      }
      if (request.operation === "stage_studio_web_design") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before staging a Studio design.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current PowerPoint audit and preservation scene.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const allowedRecipes: StudioLayoutRecipe[] = ["source", "ornl-title-content", "ornl-title-two-column", "ornl-title-card-grid", "ornl-title-table", "ornl-title-figure-grid", "ornl-title-objective-columns", "ornl-title-steps-evidence", "ornl-title-labeled-figure-grid", "ornl-title-question-diagram", "ornl-title-challenges-evidence", "ornl-title-process-flow", "template-layout"];
        const recipe = String(request.input.recipe ?? "") as StudioLayoutRecipe;
        if (!allowedRecipes.includes(recipe)) throw new Error("Choose a supported Studio web recipe.");
        if (isProtectedOrnlTemplateSlide(deck, slideNumber) && recipe !== "source") throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred. Keep it on Source geometry; its approved marks, artwork, photography, legal copy, geometry, master, and layout cannot be recomposed or restyled.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber) && ([request.input.nodeFrames, request.input.nodeStyles, request.input.figureTreatments].some((value) => Array.isArray(value) && value.length > 0))) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred. Source preservation cannot include node-frame, style, or figure-treatment overrides.`);
        const layoutId = typeof request.input.layoutId === "string" ? request.input.layoutId : undefined;
        const layout = layoutId ? templateCatalog?.layouts.find((item) => item.id === layoutId) : undefined;
        if (recipe === "template-layout" && !layout) throw new Error("Choose an installed Template Pack layout ID before staging template-layout.");
        const catalog = await getOrBuildSlideCatalog(deck, current);
        const baseScene = deck.studioScene ?? compileStudioWebScene(deck, catalog);
        let studioScene = recomposeStudioWebSlide(baseScene, slideNumber, recipe, layout, String(request.input.rationale ?? "Recompose the exact source content as a coherent ORNL-branded slide using the shared Studio Web Scene design system."));
        const nodeFrames = Array.isArray(request.input.nodeFrames) ? request.input.nodeFrames : [];
        for (const raw of nodeFrames) {
          const nodeId = String(raw?.nodeId ?? "");
          const values = [raw?.xInches, raw?.yInches, raw?.widthInches, raw?.heightInches].map(Number);
          if (!nodeId || values.some((value) => !Number.isFinite(value))) throw new Error("Every Studio node-frame override requires a valid nodeId and finite x, y, width, and height in inches.");
          studioScene = updateStudioWebNodeFrame(studioScene, slideNumber, nodeId, {
            x: values[0] * 914_400,
            y: values[1] * 914_400,
            width: values[2] * 914_400,
            height: values[3] * 914_400,
            rotation: Number.isFinite(Number(raw?.rotation)) ? Number(raw.rotation) : 0,
          });
        }
        const nodeStyles = Array.isArray(request.input.nodeStyles) ? request.input.nodeStyles : [];
        for (const raw of nodeStyles) {
          const nodeId = String(raw?.nodeId ?? "");
          if (!nodeId) throw new Error("Every Studio node-style override requires a valid nodeId.");
          const patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">> = {};
          if (raw?.fontSizePt !== undefined) patch.fontSizePt = Number(raw.fontSizePt);
          if (raw?.fontWeight !== undefined) patch.fontWeight = Number(raw.fontWeight) as 400 | 600 | 700;
          if (raw?.color !== undefined) patch.color = String(raw.color);
          if (raw?.textAlign !== undefined) patch.textAlign = String(raw.textAlign) as "left" | "center" | "right";
          if (raw?.verticalAlign !== undefined) patch.verticalAlign = String(raw.verticalAlign) as "top" | "middle" | "bottom";
          if (raw?.objectFit !== undefined) patch.objectFit = String(raw.objectFit) as "contain" | "cover";
          studioScene = updateStudioWebNodeStyle(studioScene, slideNumber, nodeId, patch);
        }
        const figureTreatments = Array.isArray(request.input.figureTreatments) ? request.input.figureTreatments : [];
        for (let index = 0; index < figureTreatments.length; index += 1) {
          const raw = figureTreatments[index];
          const treatment: StudioFigureTreatment = {
            id: String(raw?.id ?? `studio-figure-${slideNumber}-${index + 1}`),
            nodeIds: Array.isArray(raw?.nodeIds) ? raw.nodeIds.map(String) : [],
            mode: String(raw?.mode ?? "preserve-as-unit") as StudioFigureTreatment["mode"],
            verificationStatus: String(raw?.verificationStatus ?? "source-locked") as StudioFigureTreatment["verificationStatus"],
            intentSummary: String(raw?.intentSummary ?? ""),
            informationInventory: Array.isArray(raw?.informationInventory) ? raw.informationInventory.map(String) : [],
            invariants: Array.isArray(raw?.invariants) ? raw.invariants.map(String) : [],
            rationale: String(raw?.rationale ?? ""),
            relationships: Array.isArray(raw?.relationships) ? raw.relationships.map((relationship: Record<string, unknown>) => ({
              fromNodeId: String(relationship.fromNodeId ?? ""),
              toNodeId: String(relationship.toNodeId ?? ""),
              kind: String(relationship.kind ?? "contained-by") as NonNullable<StudioFigureTreatment["relationships"]>[number]["kind"],
            })) : [],
            groupFrame: raw?.groupFrame && typeof raw.groupFrame === "object" ? {
              x: Number((raw.groupFrame as Record<string, unknown>).xInches) * 914_400,
              y: Number((raw.groupFrame as Record<string, unknown>).yInches) * 914_400,
              width: Number((raw.groupFrame as Record<string, unknown>).widthInches) * 914_400,
              height: Number((raw.groupFrame as Record<string, unknown>).heightInches) * 914_400,
              rotation: Number((raw.groupFrame as Record<string, unknown>).rotation ?? 0),
            } : undefined,
            focalPoint: raw?.focalPoint && typeof raw.focalPoint === "object" ? { x: Number((raw.focalPoint as Record<string, unknown>).x), y: Number((raw.focalPoint as Record<string, unknown>).y) } : undefined,
            crop: raw?.crop && typeof raw.crop === "object" ? {
              left: Number((raw.crop as Record<string, unknown>).left),
              top: Number((raw.crop as Record<string, unknown>).top),
              right: Number((raw.crop as Record<string, unknown>).right),
              bottom: Number((raw.crop as Record<string, unknown>).bottom),
            } : undefined,
            relationshipPolicy: raw?.relationshipPolicy ? String(raw.relationshipPolicy) as StudioFigureTreatment["relationshipPolicy"] : undefined,
            lockAspectRatio: typeof raw?.lockAspectRatio === "boolean" ? raw.lockAspectRatio : undefined,
          };
          studioScene = updateStudioFigureTreatment(studioScene, slideNumber, treatment);
        }
        const visualNeedIds = Array.isArray(request.input.visualNeedIds) ? request.input.visualNeedIds.map(String) : [];
        if (visualNeedIds.length) studioScene = markStudioVisualNeedsReconstructionReady(studioScene, slideNumber, visualNeedIds);
        const adoptedDeck: DeckJob = { ...deck, operationScope: "reflow", studioScene };
        if (request.input.compilerMode !== undefined && request.input.compilerMode !== "fresh-composition") throw new Error("Studio has one central composition path. Use compilerMode fresh-composition; source-bound proposals are separate legacy cleanup evidence and cannot become the Studio design authority.");
        const compilerMode = "fresh-composition" as const;
        {
          const addressedThreadIds = requestedAddressedThreadIds(request.input);
          const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
          const revisionBoundThreads = markSubmittedThreadsForReanchor(current.designThreads, deck.id, slideNumber, nextSlideRevision, addressedThreadIds);
          const next = touchProject({
            ...current,
            designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, [slideNumber], addressedThreadIds),
            decks: current.decks.map((item) => item.id === deck.id ? { ...adoptedDeck, proposal: undefined, status: "ready-for-cleanup" as const } : item),
          }, "mcp-studio-fresh-composition-designed", `AI recomposed slide ${slideNumber} of ${deck.name} with the shared ${recipe} Studio web recipe for the fresh-composition compiler; no source-bound proposal was created and source bytes remain unchanged.`);
          projectRef.current = next;
          setProject(next);
          setSelectedDeckId(deck.id);
          setActiveView("studio");
          return {
            projectUpdatedAt: next.project.updatedAt,
            studioSceneRevision: studioScene.revision,
            compilerMode,
            slide: { number: slideNumber, recipe, layoutId: layout?.id, layoutName: layout?.name, nodeFrameOverrideCount: nodeFrames.length, nodeStyleOverrideCount: nodeStyles.length, figureTreatmentCount: figureTreatments.length, reconstructedVisualNeedCount: visualNeedIds.length },
            applied: false,
            saved: false,
            instruction: "The semantic web composition is visible in Studio and no PowerPoint was changed. Call preview_studio_fresh_composition with projectUpdatedAt to compile exact content into a new editable one-slide PPTX, render and measure it in Microsoft PowerPoint, then critique the returned pixels.",
          };
        }
      }
      if (request.operation === "refine_studio_layout") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before refining the Studio layout.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before refining its layout.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before applying layout constraints.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred. Its geometry cannot enter Studio layout refinement.`);
        const rawConstraints = Array.isArray(request.input.constraints) ? request.input.constraints as Array<Record<string, unknown>> : [];
        const constraints: StudioConstraintRequest[] = rawConstraints.map((raw) => ({
          kind: String(raw.kind ?? "align") as StudioConstraintRequest["kind"],
          mode: String(raw.mode ?? "left") as StudioConstraintRequest["mode"],
          nodeIds: Array.isArray(raw.nodeIds) ? raw.nodeIds.map(String) : [],
          groups: Array.isArray(raw.groups) ? raw.groups.map((group) => Array.isArray(group) ? group.map(String) : []) : undefined,
          anchorNodeId: typeof raw.anchorNodeId === "string" ? raw.anchorNodeId : undefined,
          gridPt: raw.gridPt === undefined ? undefined : Number(raw.gridPt),
          rationale: String(raw.rationale ?? ""),
          author: "ai",
        }));
        const studioSlide = deck.studioScene.slides.find((slide) => slide.slideNumber === slideNumber);
        if (!studioSlide) throw new Error("The requested Studio slide is unavailable.");
        const preview = studioFreshPreviewsRef.current[`${deck.id}:${slideNumber}`];
        const matchingMeasurement = preview?.sceneRevision === deck.studioScene.revision && preview.slideUpdatedAt === studioSlide.updatedAt ? preview.nativeMeasurement : undefined;
        const refined = applyStudioLayoutConstraints(deck.studioScene, slideNumber, constraints, matchingMeasurement);
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const nextSlideRevision = refined.scene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? refined.scene.revision;
        const revisionBoundThreads = markSubmittedThreadsForReanchor(current.designThreads, deck.id, slideNumber, nextSlideRevision, addressedThreadIds);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, [slideNumber], addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: refined.scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-layout-refined", `AI applied ${refined.constraints.length} high-level Studio layout constraints on slide ${slideNumber} of ${deck.name}; source bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          sceneRevision: refined.scene.revision,
          slideNumber,
          changedNodeIds: refined.changedNodeIds,
          constraints: refined.constraints,
          evidenceAuthority: refined.evidenceAuthority,
          diagnostics: refined.diagnostics,
          applied: false,
          saved: false,
          instruction: refined.evidenceAuthority === "powerpoint-native" ? "Build this exact slide revision and inspect its PowerPoint-native pixels. The optical constraints used a matching native measurement but the changed result still requires final rendering." : "Build this exact slide revision, inspect and measure it in Microsoft PowerPoint, then run refine_studio_layout again for any remaining optical mismatch. Do not guess correction coordinates.",
        };
      }
      if (request.operation === "publish_studio_component_style") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before publishing a reusable component style.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before publishing a reusable component style.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before publishing the reusable component style.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot define a reusable Studio component.`);
        const requestedTargets = Array.isArray(request.input.targetSlideNumbers) ? [...new Set(request.input.targetSlideNumbers.map(Number))] : undefined;
        if (requestedTargets?.some((number) => !Number.isInteger(number) || number < 1 || number > deck.audit!.slideCount)) throw new Error(`Target slides must be unique integers from 1 to ${deck.audit.slideCount}.`);
        if (requestedTargets?.some((number) => isProtectedOrnlTemplateSlide(deck, number))) throw new Error("Protected ORNL template slides cannot be component-propagation targets.");
        const targetSlideNumbers = requestedTargets ?? deck.studioScene.slides.filter((slide) => !isProtectedOrnlTemplateSlide(deck, slide.slideNumber)).map((slide) => slide.slideNumber);
        const sourceContentAndGeometry = JSON.stringify(deck.studioScene.slides.map((slide) => ({ slideNumber: slide.slideNumber, nodes: slide.nodes.map((node) => ({ id: node.id, text: node.text, table: node.table?.cells.map((cell) => ({ id: cell.id, text: cell.text, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, semanticColorRole: cell.semanticColorRole })), frame: node.frame })) })));
        const result = adoptStudioComponentStyle(deck.studioScene, { slideNumber, nodeId: String(request.input.nodeId ?? ""), name: typeof request.input.name === "string" ? request.input.name : undefined, targetSlideNumbers });
        const resultContentAndGeometry = JSON.stringify(result.scene.slides.map((slide) => ({ slideNumber: slide.slideNumber, nodes: slide.nodes.map((node) => ({ id: node.id, text: node.text, table: node.table?.cells.map((cell) => ({ id: cell.id, text: cell.text, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, semanticColorRole: cell.semanticColorRole })), frame: node.frame })) })));
        if (sourceContentAndGeometry !== resultContentAndGeometry) throw new Error("Studio refused component propagation that changed exact content, table semantics, source order, or geometry.");
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const revisionBoundThreads = result.affectedSlideNumbers.reduce((threads, affectedSlideNumber) => {
          const revision = result.scene.slides.find((slide) => slide.slideNumber === affectedSlideNumber)?.updatedAt ?? result.scene.revision;
          return markSubmittedThreadsForReanchor(threads, deck.id, affectedSlideNumber, revision, addressedThreadIds);
        }, current.designThreads);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, result.affectedSlideNumbers, addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-component-style-published", `AI published ${result.definition.name} to ${result.affectedNodeIds.length} compatible source-bound component instances across ${result.affectedSlideNumbers.length} slides of ${deck.name}; exact content, geometry, semantic table roles, protected slides, and source bytes remain unchanged.`);
        pushStudioEditHistory(deck.id, deck.studioScene);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          sceneRevision: result.scene.revision,
          component: { id: result.definition.id, name: result.definition.name, role: result.definition.role, surface: result.definition.surface, affectedNodeCount: result.affectedNodeIds.length, affectedSlideNumbers: result.affectedSlideNumbers, skippedNodeIds: result.skippedNodeIds },
          exactContentPreserved: true,
          geometryPreserved: true,
          semanticTableRolesPreserved: true,
          protectedSlidesPreserved: true,
          saved: false,
          exported: false,
          instruction: "Build every affected slide, inspect the exact PowerPoint-native results, then rerun get_studio_deck_consistency. Treat any hierarchy or contrast regression as a component-definition revision, not as a one-off slide patch.",
        };
      }
      if (request.operation === "refine_studio_table") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before refining the Studio table.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before refining a table.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before refining the table.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot enter table refinement.`);
        const tableNodeId = String(request.input.tableNodeId ?? "");
        const sourceSlide = deck.studioScene.slides.find((slide) => slide.slideNumber === slideNumber);
        const sourceNode = sourceSlide?.nodes.find((node) => node.id === tableNodeId);
        if (!sourceNode?.table || sourceNode.kind !== "table") throw new Error("Choose a table node ID from the current Studio slide.");
        const sourceStructure = JSON.stringify(sourceNode.table.cells.map((cell) => ({ id: cell.id, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, text: cell.text, semanticColorRole: cell.semanticColorRole })));
        let scene = deck.studioScene;
        const designPatch: StudioTableDesignPatch = {};
        const columnWidths = Array.isArray(request.input.columnWidthsInches) ? request.input.columnWidthsInches.map(Number) : undefined;
        const rowHeights = Array.isArray(request.input.rowHeightsInches) ? request.input.rowHeightsInches.map(Number) : undefined;
        if (columnWidths) {
          if (columnWidths.length !== sourceNode.table.columns) throw new Error(`Supply exactly ${sourceNode.table.columns} column widths for this table.`);
          const tableWidth = sourceNode.frame.width / 914_400;
          if (Math.abs(columnWidths.reduce((sum, value) => sum + value, 0) - tableWidth) > .08) throw new Error(`Column widths must total the current ${tableWidth.toFixed(2)}-inch table width.`);
          designPatch.columnWidths = columnWidths;
        }
        if (rowHeights) {
          if (rowHeights.length !== sourceNode.table.rows) throw new Error(`Supply exactly ${sourceNode.table.rows} row heights for this table.`);
          const tableHeight = sourceNode.frame.height / 914_400;
          if (Math.abs(rowHeights.reduce((sum, value) => sum + value, 0) - tableHeight) > .08) throw new Error(`Row heights must total the current ${tableHeight.toFixed(2)}-inch table height.`);
          designPatch.rowHeights = rowHeights;
        }
        if (request.input.headerRows !== undefined) designPatch.headerRows = Number(request.input.headerRows);
        if (request.input.borderMode !== undefined) designPatch.borderMode = String(request.input.borderMode) as StudioTableDesign["borderMode"];
        if (request.input.borderColor !== undefined) designPatch.borderColor = String(request.input.borderColor);
        if (request.input.borderWidthPt !== undefined) designPatch.borderWidthPt = Number(request.input.borderWidthPt);
        if (request.input.defaultPaddingPt !== undefined) {
          const value = Number(request.input.defaultPaddingPt);
          designPatch.defaultPaddingPt = { top: value, right: value, bottom: value, left: value };
        }
        if (Object.keys(designPatch).length) scene = updateStudioTableDesign(scene, slideNumber, tableNodeId, designPatch);
        const rawCellStyles = Array.isArray(request.input.cellStyles) ? request.input.cellStyles as Array<Record<string, unknown>> : [];
        if (new Set(rawCellStyles.map((item) => String(item.cellId ?? ""))).size !== rawCellStyles.length) throw new Error("Each Studio table cell may appear only once per refinement transaction.");
        for (const raw of rawCellStyles) {
          const cellId = String(raw.cellId ?? "");
          const patch: StudioTableCellDesignPatch = {};
          if (raw.fill !== undefined) patch.fill = String(raw.fill);
          if (raw.color !== undefined) patch.color = String(raw.color);
          if (raw.fontSizePt !== undefined) patch.fontSizePt = Number(raw.fontSizePt);
          if (raw.fontWeight !== undefined) patch.fontWeight = Number(raw.fontWeight) as 400 | 600 | 700;
          if (raw.textAlign !== undefined) patch.textAlign = String(raw.textAlign) as "left" | "center" | "right";
          if (raw.verticalAlign !== undefined) patch.verticalAlign = String(raw.verticalAlign) as "top" | "middle" | "bottom";
          if (raw.paddingPt !== undefined) { const value = Number(raw.paddingPt); patch.paddingPt = { top: value, right: value, bottom: value, left: value }; }
          if (raw.borders && typeof raw.borders === "object") patch.borders = raw.borders as StudioTableCellDesign["borders"];
          scene = updateStudioTableCellDesign(scene, slideNumber, tableNodeId, cellId, patch);
        }
        if (!Object.keys(designPatch).length && rawCellStyles.length === 0) throw new Error("Supply at least one bounded table design change.");
        const resultNode = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.nodes.find((node) => node.id === tableNodeId);
        const resultStructure = JSON.stringify(resultNode?.table?.cells.map((cell) => ({ id: cell.id, row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan, text: cell.text, semanticColorRole: cell.semanticColorRole })));
        if (sourceStructure !== resultStructure) throw new Error("Studio refused a table refinement that changed exact cell content, merge topology, order, or semantic roles.");
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const nextSlideRevision = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? scene.revision;
        const revisionBoundThreads = markSubmittedThreadsForReanchor(current.designThreads, deck.id, slideNumber, nextSlideRevision, addressedThreadIds);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, [slideNumber], addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-table-refined", `AI refined the source-bound table ${tableNodeId} on slide ${slideNumber} of ${deck.name}; exact cell content, merge topology, semantic roles, and source bytes remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return {
          projectUpdatedAt: next.project.updatedAt,
          sceneRevision: scene.revision,
          slideNumber,
          table: { nodeId: tableNodeId, rows: sourceNode.table.rows, columns: sourceNode.table.columns, cellStyleCount: rawCellStyles.length, design: resultNode ? resolvedStudioTableDesign(resultNode) : undefined },
          exactContentPreserved: true,
          mergeTopologyPreserved: true,
          semanticRolesPreserved: true,
          saved: false,
          exported: false,
          instruction: "Build this exact slide revision, inspect the PowerPoint-native table crop and cell measurements, and revise again if any cell clips or loses hierarchy.",
        };
      }
      if (request.operation === "publish_studio_table_exemplar") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before publishing a table exemplar.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before publishing a table exemplar.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before publishing the table exemplar.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot define a table exemplar.`);
        const requestedTargets = Array.isArray(request.input.targetSlideNumbers) ? [...new Set(request.input.targetSlideNumbers.map(Number))] : undefined;
        if (requestedTargets?.some((number) => !Number.isInteger(number) || number < 1 || number > deck.audit!.slideCount || isProtectedOrnlTemplateSlide(deck, number))) throw new Error("Table-exemplar targets must be unprotected slide numbers in the current deck.");
        const sourceInvariant = JSON.stringify(deck.studioScene.slides.map((slide) => ({ slideNumber: slide.slideNumber, nodes: slide.nodes.map((node) => ({ id: node.id, text: node.text, frame: node.frame, cells: node.table?.cells })) })));
        const targetSlideNumbers = requestedTargets ?? deck.studioScene.slides.filter((slide) => !isProtectedOrnlTemplateSlide(deck, slide.slideNumber)).map((slide) => slide.slideNumber);
        const result = publishStudioTableExemplar(deck.studioScene, { slideNumber, tableNodeId: String(request.input.tableNodeId ?? ""), name: typeof request.input.name === "string" ? request.input.name : undefined, targetSlideNumbers });
        const resultInvariant = JSON.stringify(result.scene.slides.map((slide) => ({ slideNumber: slide.slideNumber, nodes: slide.nodes.map((node) => ({ id: node.id, text: node.text, frame: node.frame, cells: node.table?.cells })) })));
        if (sourceInvariant !== resultInvariant) throw new Error("Studio refused a table exemplar that changed exact content, geometry, source cells, merge topology, or semantic roles.");
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const revisionBoundThreads = result.affectedSlideNumbers.reduce((threads, affectedSlideNumber) => {
          const revision = result.scene.slides.find((slide) => slide.slideNumber === affectedSlideNumber)?.updatedAt ?? result.scene.revision;
          return markSubmittedThreadsForReanchor(threads, deck.id, affectedSlideNumber, revision, addressedThreadIds);
        }, current.designThreads);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, result.affectedSlideNumbers, addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-table-exemplar-published", `AI published ${result.definition.name} to ${result.affectedTableNodeIds.length} structurally compatible native tables across ${result.affectedSlideNumbers.length} slides of ${deck.name}; content, geometry, merged topology, semantic fills, protected slides, and source bytes remain unchanged.`);
        pushStudioEditHistory(deck.id, deck.studioScene);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: result.scene.revision, definition: result.definition, affectedSlideNumbers: result.affectedSlideNumbers, affectedTableNodeIds: result.affectedTableNodeIds, skippedTableNodeIds: result.skippedTableNodeIds, exactContentPreserved: true, mergeTopologyPreserved: true, semanticFillsPreserved: true, saved: false, exported: false, instruction: "Build every affected slide, inspect its native PowerPoint table crop and cell measurements, then rerun get_studio_deck_consistency." };
      }
      if (request.operation === "plan_studio_table_continuation") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before planning a continuation.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before planning a table continuation.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before planning the continuation.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be continued.`);
        const result = planStudioTableContinuation(deck.studioScene, { slideNumber, tableNodeId: String(request.input.tableNodeId ?? ""), maximumBodyRowsPerSlide: Number(request.input.maximumBodyRowsPerSlide ?? 8), rationale: typeof request.input.rationale === "string" ? request.input.rationale : undefined });
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" as const } : item) }, "mcp-studio-table-continuation-planned", `${result.plan.status === "ready" ? `AI planned ${result.plan.segments.length} merge-safe table continuation slides` : "AI recorded a blocked table continuation plan"} for slide ${slideNumber} of ${deck.name}; source content and structure remain unchanged.`);
        pushStudioEditHistory(deck.id, deck.studioScene);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: result.scene.revision, plan: result.plan, exactContentPreserved: true, mergeTopologyPreserved: true, materializesOnNextBuild: result.plan.status === "ready", saved: false, exported: false, instruction: result.plan.status === "ready" ? "Build this slide or the full deck. The compiler will materialize these body-row ranges as editable PowerPoint slides, rerender every output in Microsoft PowerPoint, and reject any copy, table-topology, or fit regression." : "Resolve the reported header, merge, or orchestration blocker before requesting another plan." };
      }
      if (request.operation === "clear_studio_table_continuation") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before clearing the continuation plan.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before clearing a table continuation.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before clearing the continuation plan.");
        const slideNumber = Number(request.input.slideNumber);
        const tableNodeId = String(request.input.tableNodeId ?? "");
        const scene = clearStudioTableContinuation(deck.studioScene, { slideNumber, tableNodeId });
        const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene, proposal: undefined, status: "ready-for-cleanup" as const } : item) }, "mcp-studio-table-continuation-cleared", `AI cleared a table continuation plan on slide ${slideNumber} of ${deck.name}; the native table remains unchanged.`);
        pushStudioEditHistory(deck.id, deck.studioScene);
        projectRef.current = next;
        setProject(next);
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, tableNodeId, cleared: true, saved: false, exported: false };
      }
      if (request.operation === "author_studio_connector") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before authoring the connector.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.studioScene) throw new Error("Create and persist the Studio Web Scene before authoring a connector.");
        if (request.input.expectedSceneRevision !== deck.studioScene.revision) throw new Error("The Studio scene changed. Read get_studio_web_scene again before authoring the connector.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot enter connector authoring.`);
        const connectorNodeId = String(request.input.connectorNodeId ?? "");
        const design: StudioConnectorDesign = {
          fromNodeId: String(request.input.fromNodeId ?? ""),
          toNodeId: String(request.input.toNodeId ?? ""),
          fromSide: String(request.input.fromSide ?? "right") as StudioConnectorDesign["fromSide"],
          toSide: String(request.input.toSide ?? "left") as StudioConnectorDesign["toSide"],
          stroke: String(request.input.stroke ?? "#00662C"),
          widthPt: Number(request.input.widthPt ?? 1.5),
          dash: String(request.input.dash ?? "solid") as StudioConnectorDesign["dash"],
          beginArrow: String(request.input.beginArrow ?? "none") as StudioConnectorDesign["beginArrow"],
          endArrow: String(request.input.endArrow ?? "triangle") as StudioConnectorDesign["endArrow"],
          verificationStatus: "verified",
        };
        const scene = updateStudioConnectorDesign(deck.studioScene, slideNumber, connectorNodeId, design);
        const addressedThreadIds = requestedAddressedThreadIds(request.input);
        const nextSlideRevision = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? scene.revision;
        const revisionBoundThreads = markSubmittedThreadsForReanchor(current.designThreads, deck.id, slideNumber, nextSlideRevision, addressedThreadIds);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreadsForSlides(revisionBoundThreads, deck.id, [slideNumber], addressedThreadIds),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: scene, proposal: undefined, status: "ready-for-cleanup" as const } : item),
        }, "mcp-studio-connector-authored", `AI authored verified connector ${connectorNodeId} on slide ${slideNumber} of ${deck.name}; source bytes and exact content remain unchanged.`);
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setStudioOpenSlideNumber(slideNumber);
        setActiveView("studio");
        return { projectUpdatedAt: next.project.updatedAt, sceneRevision: scene.revision, slideNumber, connector: { nodeId: connectorNodeId, ...design }, exactContentPreserved: true, saved: false, exported: false, instruction: "Build this exact slide revision, inspect the connector endpoints and arrow direction in PowerPoint-native pixels, and revise only against the verified relationship inventory." };
      }
      if (request.operation === "stage_font_cleanup") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck) throw new Error("The requested deck is not open.");
        const proposal = createFontCleanupProposal(deck, current.project.updatedAt);
        const next = touchProject({
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-proposal-staged", `AI staged font cleanup for ${deck.name}; no changes were applied.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, changes: proposal.changes }, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_designer_cleanup") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        if (request.input.designStandardVersion !== PRESENTATION_DESIGN_STANDARD.version) throw new Error("The Presentation Design Standard changed. Read get_design_contract and get_cleanup_rule_profile again before staging Designer Cleanup.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck) throw new Error("The requested deck is not open.");
        const proposal = createDesignerCleanupProposal(deck, current.project.updatedAt);
        const next = touchProject({
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-designer-proposal-staged", `AI staged a deck-wide designer cleanup for ${deck.name}; no changes were applied.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode, changes: proposal.changes, slideDispositions: proposal.slideDispositions, tableExceptions: proposal.tableExceptions, layoutExceptions: proposal.layoutExceptions }, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_table_design_update") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging table design.");
        if (request.input.designStandardVersion !== PRESENTATION_DESIGN_STANDARD.version) throw new Error("The Presentation Design Standard changed. Read get_design_contract and get_cleanup_rule_profile again before staging table design.");
        if (request.input.semanticColorPolicy !== PRESENTATION_DESIGN_STANDARD.semanticVisualPolicy.tableColorPolicy) throw new Error("Table design requires the preserve-source semantic-color policy.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const tableIds = Array.isArray(request.input.tableIds) ? request.input.tableIds.map(String) : [];
        const variant = request.input.variant === "dense-technical" ? "dense-technical" as const : "standard" as const;
        const proposal = createTableStyleProposal(deck, current.project.updatedAt, { tableIds, variant, semanticColorPolicy: "preserve-source" });
        const next = touchProject({ ...current, designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)), decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item) }, "mcp-table-design-staged", `AI staged the shared ${variant} ORNL table component for ${tableIds.length} table${tableIds.length === 1 ? "" : "s"} in ${deck.name}; explicitly addressed comments were cleared and source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode }, tableCount: tableIds.length, tableIds, variant, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_slide_geometry_update") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const object = (deck.audit.editableObjects ?? []).find((item) => item.id === request.input.objectId);
        if (!object) throw new Error("The requested object ID is not present in the current deck revision. Read get_slide_design_context again.");
        const proposal = createGeometryEditProposal(deck, current.project.updatedAt, {
          objectId: object.id,
          target: {
            x: Number(request.input.xInches) * 914_400,
            y: Number(request.input.yInches) * 914_400,
            width: Number(request.input.widthInches) * 914_400,
            height: Number(request.input.heightInches) * 914_400,
          },
          rationale: String(request.input.rationale ?? ""),
          author: "ai",
          constraints: {
            allowIntentionalOverlap: request.input.allowIntentionalOverlap === true,
            allowFitRisk: request.input.allowFitRisk === true,
            allowSafeArea: request.input.allowSafeArea === true,
            allowAspectRatioChange: request.input.allowAspectRatioChange === true,
          },
        });
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-slide-geometry-staged", `AI staged a measured ${object.kind} geometry edit on slide ${object.slideNumber} of ${deck.name}; source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        const geometry = proposal.changes.find((change) => change.id === `geometry-${object.id}`)?.geometryCommands?.[0];
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode }, object: { id: object.id, slideNumber: object.slideNumber, name: object.name, kind: object.kind }, geometry, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (["solve_and_stage_alignment", "solve_and_stage_distribution", "solve_and_stage_safe_region", "solve_and_stage_group_layout", "fit_scene_to_layout"].includes(request.operation)) {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read a fresh inspection packet before solving layout geometry.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Semantic layout solving requires a Designer Cleanup or native composition scope.");
        const slideNumber = Number(request.input.slideNumber);
        const groups = Array.isArray(request.input.groups) ? request.input.groups.map((group) => Array.isArray(group) ? group.map(String) : []) : undefined;
        const sceneRegionInputs = Array.isArray(request.input.regions) ? request.input.regions.filter((region): region is Record<string, unknown> => Boolean(region) && typeof region === "object") : [];
        const sceneRegionObjectIds = sceneRegionInputs.flatMap((region) => Array.isArray(region.groups) ? region.groups.flatMap((group) => Array.isArray(group) ? group.map(String) : []) : []);
        const objectIds = request.operation === "solve_and_stage_group_layout" ? [...new Set(groups?.flat() ?? [])] : request.operation === "fit_scene_to_layout" ? [...new Set(sceneRegionObjectIds)] : Array.isArray(request.input.objectIds) ? [...new Set(request.input.objectIds.map(String))] : [];
        const rationale = String(request.input.rationale ?? "").trim();
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const measurement = await getOrBuildNativeMeasurement(deck, "current", current);
        const result = request.operation === "solve_and_stage_alignment"
          ? solveAlignment({ deck, measurement, slideNumber, objectIds, mode: String(request.input.mode ?? "optical-left") as AlignmentMode, rationale, anchorObjectId: request.input.anchorObjectId ? String(request.input.anchorObjectId) : undefined })
          : request.operation === "solve_and_stage_distribution"
            ? solveDistribution({ deck, slideNumber, objectIds, groups, mode: String(request.input.mode ?? "vertical-equal-gap") as DistributionMode, rationale })
            : request.operation === "solve_and_stage_group_layout"
              ? (() => {
                if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before solving content into an approved region.");
                const layout = templateCatalog.layouts.find((item) => item.id === String(request.input.layoutId ?? ""));
                const slot = layout?.semantic?.slots.find((item) => item.id === String(request.input.slotId ?? ""));
                if (!layout || !slot) throw new Error("The requested approved layout region is stale. Read the current Template Pack catalog and choose a semantic slot again.");
                const slotBounds = slot.preferredBounds ?? { x: slot.x, y: slot.y, width: slot.width, height: slot.height };
                const slotMinimumBounds = slot.minimumBounds ?? slotBounds;
                const padding = slot.paddingIntentPt ?? { top: 0, right: 0, bottom: 0, left: 0 };
                const slotScaleFloor = Math.max(slotMinimumBounds.width / Math.max(1, slotBounds.width), slotMinimumBounds.height / Math.max(1, slotBounds.height));
                const groupRoles = Array.isArray(request.input.groupRoles) ? request.input.groupRoles.map(String) as GroupHierarchyRole[] : undefined;
                return solveGroupLayout({ deck, slideNumber, groups: groups ?? [], groupRoles, regionPt: { left: slotBounds.x / 12_700 + padding.left, top: slotBounds.y / 12_700 + padding.top, width: slotBounds.width / 12_700 - padding.left - padding.right, height: slotBounds.height / 12_700 - padding.top - padding.bottom }, mode: String(request.input.mode ?? "vertical-stack") as GroupLayoutMode, alignment: String(request.input.alignment ?? "start") as GroupLayoutAlignment, preferredGapPt: Number(request.input.preferredGapPt ?? PRESENTATION_DESIGN_STANDARD.componentSystem.spacing.normalPt), allowResponsiveScale: request.input.allowResponsiveScale === true, minimumScale: Math.max(slotScaleFloor, Number(request.input.minimumScale ?? .75)), rationale });
              })()
              : request.operation === "fit_scene_to_layout"
                ? (() => {
                  if (!templateCatalog) throw new Error("Install an authorized PowerPoint Template Pack before fitting a scene to approved regions.");
                  const layout = templateCatalog.layouts.find((item) => item.id === String(request.input.layoutId ?? ""));
                  if (!layout?.semantic) throw new Error("The requested approved layout is stale. Read the current Template Pack catalog and choose it again.");
                  const regions: SceneLayoutRegionRequest[] = sceneRegionInputs.map((region, index) => {
                    const slotId = String(region.slotId ?? "");
                    const slot = layout.semantic!.slots.find((item) => item.id === slotId);
                    if (!slot || ["footer", "date", "slide-number"].includes(slot.role)) throw new Error(`Scene region ${index + 1} does not reference an editable content slot in the selected approved layout.`);
                    const slotBounds = slot.preferredBounds ?? { x: slot.x, y: slot.y, width: slot.width, height: slot.height };
                    const slotMinimumBounds = slot.minimumBounds ?? slotBounds;
                    const padding = slot.paddingIntentPt ?? { top: 0, right: 0, bottom: 0, left: 0 };
                    const slotScaleFloor = Math.max(slotMinimumBounds.width / Math.max(1, slotBounds.width), slotMinimumBounds.height / Math.max(1, slotBounds.height));
                    const regionGroups = Array.isArray(region.groups) ? region.groups.map((group) => Array.isArray(group) ? group.map(String) : []) : [];
                    const groupRoles = Array.isArray(region.groupRoles) ? region.groupRoles.map(String) as GroupHierarchyRole[] : undefined;
                    return {
                      id: slot.id,
                      groups: regionGroups,
                      groupRoles,
                      regionPt: {
                        left: slotBounds.x / 12_700 + padding.left,
                        top: slotBounds.y / 12_700 + padding.top,
                        width: slotBounds.width / 12_700 - padding.left - padding.right,
                        height: slotBounds.height / 12_700 - padding.top - padding.bottom,
                      },
                      mode: String(region.mode ?? "vertical-stack") as GroupLayoutMode,
                      alignment: String(region.alignment ?? (slot.alignmentIntent === "optical-left" ? "start" : "center")) as GroupLayoutAlignment,
                      preferredGapPt: Number(region.preferredGapPt ?? PRESENTATION_DESIGN_STANDARD.componentSystem.spacing.normalPt),
                      allowResponsiveScale: region.allowResponsiveScale === true,
                      minimumScale: Math.max(slotScaleFloor, Number(region.minimumScale ?? .75)),
                    };
                  });
                  return solveSceneToLayout({ deck, slideNumber, regions, rationale });
                })()
              : solveSafeRegion({ deck, slideNumber, objectIds, rationale });
        if (result.status === "infeasible") return { updatedAt: current.project.updatedAt, staged: false, result, instruction: "Do not guess coordinates. Address the reported constraint or select a different semantic operation, then solve again." };
        if (projectRef.current.project.updatedAt !== current.project.updatedAt) throw new Error("The project changed while PowerPoint was measuring. Read a fresh inspection packet and solve again.");
        const proposal = createGeometryBatchProposal(deck, current.project.updatedAt, result.commands.map((command) => ({ objectId: command.objectId, target: command.target, rationale: command.rationale, author: "ai" as const, constraints: command.constraints })));
        const proposedDeck = { ...deck, proposal };
        const proposalMeasurement = await getOrBuildNativeMeasurement(proposedDeck, "proposal", current);
        const currentMetric = calculateDesignMetrics(deck, measurement).slides.find((item) => item.slideNumber === slideNumber)!;
        const proposalMetric = calculateDesignMetrics(deck, proposalMeasurement, measurement).slides.find((item) => item.slideNumber === slideNumber)!;
        const geometryMismatches = result.commands.filter((command) => {
          const measured = proposalMeasurement.objects.find((object) => object.objectId === command.objectId)?.measuredGeometryPt;
          return !measured || Math.abs(measured.left - command.target.x / 12_700) > .2 || Math.abs(measured.top - command.target.y / 12_700) > .2 || Math.abs(measured.width - command.target.width / 12_700) > .2 || Math.abs(measured.height - command.target.height / 12_700) > .2;
        }).map((command) => command.objectId);
        const metricRegressions = [
          proposalMetric.offSlideObjectCount > currentMetric.offSlideObjectCount ? "off-slide object count increased" : undefined,
          proposalMetric.textOverflowCount > currentMetric.textOverflowCount ? "native text overflow increased" : undefined,
        ].filter((value): value is string => Boolean(value));
        if (geometryMismatches.length || metricRegressions.length) return { updatedAt: current.project.updatedAt, staged: false, result: { ...result, diagnostics: [...result.diagnostics, ...metricRegressions, ...(geometryMismatches.length ? [`PowerPoint did not confirm the solved geometry for ${geometryMismatches.join(", ")}.`] : [])] }, currentMetric, proposalMetric, instruction: "Presentation Studio withheld the draft after native PowerPoint remeasurement. Revise the semantic intent or constraints instead of accepting a regression." };
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, request.operation === "solve_and_stage_alignment" ? "mcp-semantic-alignment-staged" : request.operation === "solve_and_stage_distribution" ? "mcp-semantic-distribution-staged" : request.operation === "solve_and_stage_group_layout" ? "mcp-semantic-region-layout-staged" : request.operation === "fit_scene_to_layout" ? "mcp-semantic-scene-layout-staged" : "mcp-safe-region-staged", `AI staged ${result.commands.length} deterministic ${result.operation} edits on slide ${slideNumber} of ${deck.name}; source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { projectUpdatedAt: next.project.updatedAt, staged: true, proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status }, result, nativeVerification: { currentMetric, proposalMetric, geometryConfirmed: true }, applied: false, saved: false, instruction: "Native PowerPoint remeasurement confirmed the solved geometry without overflow or off-slide regression. Inspect the Proposal pixels for aesthetic quality before acceptance." };
      }
      if (request.operation === "solve_and_stage_text_fit") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read a fresh inspection packet before fitting text.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Measured text fitting requires a Designer Cleanup or native composition scope.");
        const objectId = String(request.input.objectId ?? "");
        const object = deck.scene.objects.find((item) => item.id === objectId);
        if (!object) throw new Error("The requested text object is stale. Read a fresh inspection packet and select its current stable object ID.");
        const measurement = await getOrBuildNativeMeasurement(deck, "current", current);
        const result = solveTextFit({ deck, measurement, objectId, rationale: String(request.input.rationale ?? "") });
        if (result.status !== "solved") return { updatedAt: current.project.updatedAt, staged: false, result, instruction: result.status === "already-fit" ? "PowerPoint confirms that this frame already fits at the resolved readability floor. Diagnose another design issue instead of changing it." : "Do not shrink the text or guess geometry. Follow the measured recommendation or recompose the surrounding content." };
        if (projectRef.current.project.updatedAt !== current.project.updatedAt) throw new Error("The project changed while PowerPoint was measuring. Read a fresh inspection packet and fit again.");
        let proposal = result.geometry ? createGeometryBatchProposal(deck, current.project.updatedAt, [result.geometry]) : undefined;
        if (result.textStyle) proposal = createVisualDesignProposal(proposal ? { ...deck, proposal } : deck, current.project.updatedAt, { slideNumber: object.slideNumber, textStyles: [result.textStyle], decorations: [] });
        if (!proposal) throw new Error("The text-fit solver produced no reversible proposal commands.");
        const proposedDeck = { ...deck, proposal };
        const proposalMeasurement = await getOrBuildNativeMeasurement(proposedDeck, "proposal", current);
        const measured = proposalMeasurement.objects.find((item) => item.objectId === objectId);
        const geometryMismatch = result.geometry && (!measured?.measuredGeometryPt || Math.abs(measured.measuredGeometryPt.top - result.geometry.target.y / 12_700) > .2 || Math.abs(measured.measuredGeometryPt.height - result.geometry.target.height / 12_700) > .2);
        const currentMetric = calculateDesignMetrics(deck, measurement).slides.find((item) => item.slideNumber === object.slideNumber)!;
        const proposalMetric = calculateDesignMetrics(deck, proposalMeasurement, measurement).slides.find((item) => item.slideNumber === object.slideNumber)!;
        const nativeVerified = proposalMeasurement.authority === "powerpoint-native" && measured?.provenance.authority === "powerpoint-native";
        const stillOverflows = !measured || nativeTextFrameOverflows(measured);
        const metricRegressions = [
          proposalMetric.offSlideObjectCount > currentMetric.offSlideObjectCount ? "off-slide object count increased" : undefined,
          proposalMetric.safeRegionViolationCount > currentMetric.safeRegionViolationCount ? "safe-region violations increased" : undefined,
          proposalMetric.textOverflowCount > currentMetric.textOverflowCount ? "native text overflow increased" : undefined,
        ].filter((value): value is string => Boolean(value));
        if (!nativeVerified || geometryMismatch || stillOverflows || metricRegressions.length) return {
          updatedAt: current.project.updatedAt,
          staged: false,
          result: { ...result, diagnostics: { ...result.diagnostics, reasons: [...result.diagnostics.reasons, ...metricRegressions, ...(!nativeVerified ? ["PowerPoint-native proposal measurement was unavailable."] : []), ...(geometryMismatch ? ["PowerPoint did not confirm the solved frame geometry."] : []), ...(stillOverflows ? ["PowerPoint still reports rendered text outside the fitted frame."] : [])], recommendations: [...result.diagnostics.recommendations, "Use a wider/taller approved region or recompose related content rather than shrinking below the type floor."] } },
          currentMetric,
          proposalMetric,
          instruction: "Presentation Studio withheld the draft after native remeasurement. Choose a more suitable region or semantic composition operation.",
        };
        const nativeRender = await getOrBuildInspectionRender(proposedDeck, "proposal", current);
        const proposalSlide = nativeRender.slides.find((slide) => slide.number === object.slideNumber);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-text-fit-solved", `AI staged a PowerPoint-measured non-clipping fit for ${object.name} on slide ${object.slideNumber} of ${deck.name}; exact copy and source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { projectUpdatedAt: next.project.updatedAt, staged: true, proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status }, result, nativeVerification: { authority: proposalMeasurement.authority, binding: measured.binding, rasterSha256: proposalSlide?.sha256, geometryConfirmed: true, textOverflow: false, currentMetric, proposalMetric }, applied: false, saved: false, instruction: "PowerPoint confirmed the minimum fitted frame and readable type without clipping, safe-region regression, or copy changes. Inspect the Proposal pixels for hierarchy and surrounding composition before acceptance." };
      }
      if (request.operation === "solve_and_stage_table_layout") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read a fresh inspection packet before solving the table.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and cell-level table scene.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Native table solving requires a Designer Cleanup or native composition scope.");
        const tableId = String(request.input.tableId ?? "");
        const measurement = await getOrBuildNativeMeasurement(deck, "current", current);
        let workingMeasurement = measurement;
        let result = solveTableLayout({ deck, measurement: workingMeasurement, tableId, rationale: String(request.input.rationale ?? ""), variant: request.input.variant === "dense-technical" ? "dense-technical" : "standard" });
        if (result.status === "already-fit") return { updatedAt: current.project.updatedAt, staged: false, tableId, result, instruction: "PowerPoint confirms that the table already satisfies the resolved native type, padding, clearance, and wrap constraints. Preserve it as-is and inspect the surrounding composition." };
        if (projectRef.current.project.updatedAt !== current.project.updatedAt) throw new Error("The project changed while PowerPoint was measuring. Read a fresh inspection packet and solve again.");
        const tableObject = deck.scene.objects.find((item) => item.sourceLocator.tableId === tableId);
        let growthPlan = tableObject ? recommendedTableGrowthPlan(result, tableObject.id, String(request.input.rationale ?? "")) : undefined;
        let proposal = growthPlan ? createGeometryBatchProposal(deck, current.project.updatedAt, [{ objectId: growthPlan.objectId, target: growthPlan.target, rationale: growthPlan.rationale, author: "ai", constraints: { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false } }]) : undefined;
        if (proposal && growthPlan) {
          result = solveTableLayout({
            deck,
            measurement,
            tableId,
            rationale: String(request.input.rationale ?? ""),
            variant: request.input.variant === "dense-technical" ? "dense-technical" : "standard",
            targetBoundsPt: { width: growthPlan.target.width / 12_700, height: growthPlan.target.height / 12_700 },
          });
        }
        if (result.status === "infeasible" || !result.command) return { updatedAt: current.project.updatedAt, staged: false, tableId, result, growthPlan, instruction: growthPlan ? "The minimum safe-region growth was not enough after native PowerPoint remeasurement. Follow the updated space recommendation or use a continuation slide." : "The solver refused to shrink type or padding below the ORNL technical-slide floor. Follow a recommendation, enlarge the region, or continue the table before solving again." };
        proposal = createTableLayoutProposal(proposal ? { ...deck, proposal } : deck, current.project.updatedAt, result.command);
        let proposalMeasurement: NativeMeasurementPacket | undefined;
        const iterations: Array<{ iteration: number; commandId: string; minimumHorizontalClearancePt?: number; minimumVerticalClearancePt?: number; violationCellIds: string[] }> = [];
        let nativeVerified = false;
        for (let iteration = 1; iteration <= 3; iteration += 1) {
          const activeCommand = result.command;
          if (!activeCommand) throw new Error("The table solver lost its active command during bounded iteration.");
          proposalMeasurement = await getOrBuildNativeMeasurement({ ...deck, proposal }, "proposal", current);
          const measuredTable = proposalMeasurement.objects.find((object) => object.tableId === tableId)?.table;
          const horizontal = measuredTable?.cells.flatMap((cell) => cell.clearancesPt ? [cell.clearancesPt.left, cell.clearancesPt.right] : []) ?? [];
          const vertical = measuredTable?.cells.flatMap((cell) => cell.clearancesPt ? [cell.clearancesPt.top, cell.clearancesPt.bottom] : []) ?? [];
          const violationCellIds = measuredTable?.cells.filter((cell) => cell.clearancesPt && (cell.clearancesPt.left < activeCommand.constraints.minimumHorizontalPaddingPt - .5 || cell.clearancesPt.right < activeCommand.constraints.minimumHorizontalPaddingPt - .5 || cell.clearancesPt.top < activeCommand.constraints.minimumVerticalPaddingPt - .5 || cell.clearancesPt.bottom < activeCommand.constraints.minimumVerticalPaddingPt - .5)).map((cell) => cell.cellId) ?? [];
          iterations.push({ iteration, commandId: activeCommand.id, minimumHorizontalClearancePt: horizontal.length ? Math.min(...horizontal) : undefined, minimumVerticalClearancePt: vertical.length ? Math.min(...vertical) : undefined, violationCellIds });
          if (proposalMeasurement.authority === "powerpoint-native" && violationCellIds.length === 0) { nativeVerified = true; break; }
          workingMeasurement = proposalMeasurement;
          const refined = solveTableLayout({ deck, measurement: workingMeasurement, tableId, rationale: String(request.input.rationale ?? ""), variant: request.input.variant === "dense-technical" ? "dense-technical" : "standard" });
          if (refined.status === "infeasible" || !refined.command) return { updatedAt: current.project.updatedAt, staged: false, tableId, result: refined, growthPlan, iterations, instruction: "PowerPoint remeasurement made the resolved region infeasible. Follow the space recommendation rather than shrinking text or padding." };
          result = refined;
          proposal = createTableLayoutProposal({ ...deck, proposal }, current.project.updatedAt, refined.command);
        }
        if (!nativeVerified || !proposalMeasurement) return { updatedAt: current.project.updatedAt, staged: false, tableId, result: { ...result, status: "infeasible", diagnostics: { ...result.diagnostics, reasons: [...result.diagnostics.reasons, "Three bounded PowerPoint remeasurement rounds could not satisfy cell-clearance constraints."], recommendations: [...result.diagnostics.recommendations, "Enlarge the table region or use a continuation slide."] } }, growthPlan, iterations, instruction: "Presentation Studio withheld the draft after bounded native iteration. Do not force it through with smaller type or padding." };
        const nativeRender = await getOrBuildInspectionRender({ ...deck, proposal }, "proposal", current);
        const proposalSlideNumber = deck.audit.tables.find((table) => table.id === tableId)!.slideNumber;
        const proposalSlide = nativeRender.slides.find((slide) => slide.number === proposalSlideNumber);
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-table-layout-solved", `AI staged a deterministic cell-level layout for ${tableId} in ${deck.name}; exact cell text, merged structure, source bytes, and accepted state remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { projectUpdatedAt: next.project.updatedAt, staged: true, tableId, proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status }, result, growthPlan, iterations, nativeVerification: { authority: proposalMeasurement.authority, rasterSha256: proposalSlide?.sha256, cellClearancePassed: true }, applied: false, saved: false, instruction: "The bounded table loop solved, materialized, rerendered, and remeasured the proposal without violating cell-clearance constraints. Inspect the native pixels for hierarchy, wrap quality, and surrounding composition before acceptance." };
      }
      if (request.operation === "stage_slide_layout_update") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const slideNumber = Number(request.input.slideNumber);
        const rawCommands = Array.isArray(request.input.commands) ? request.input.commands as Array<Record<string, unknown>> : [];
        const resetObjectIds = Array.isArray(request.input.resetObjectIds) ? [...new Set(request.input.resetObjectIds.map(String))] : [];
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        if (rawCommands.length + resetObjectIds.length === 0 || rawCommands.length + resetObjectIds.length > 20) throw new Error("Stage between 1 and 20 object edits or source-geometry resets for one slide.");
        for (const objectId of resetObjectIds) {
          const object = (deck.audit.editableObjects ?? []).find((item) => item.id === objectId);
          if (!object || object.slideNumber !== slideNumber) throw new Error(`Reset object ${objectId || "(missing ID)"} does not belong to slide ${slideNumber} in the current revision.`);
        }
        const requests = rawCommands.map((command) => {
          const objectId = String(command.objectId ?? "");
          const object = (deck.audit?.editableObjects ?? []).find((item) => item.id === objectId);
          if (!object || object.slideNumber !== slideNumber) throw new Error(`Object ${objectId || "(missing ID)"} does not belong to slide ${slideNumber} in the current revision.`);
          return {
            objectId,
            target: { x: Number(command.xInches) * 914_400, y: Number(command.yInches) * 914_400, width: Number(command.widthInches) * 914_400, height: Number(command.heightInches) * 914_400 },
            rationale: String(command.rationale ?? ""),
            author: "ai" as const,
            constraints: {
              allowIntentionalOverlap: command.allowIntentionalOverlap === true,
              allowFitRisk: command.allowFitRisk === true,
              allowSafeArea: command.allowSafeArea === true,
              allowAspectRatioChange: command.allowAspectRatioChange === true,
            },
          };
        });
        const proposal = createGeometryBatchProposal(deck, current.project.updatedAt, requests, { resetObjectIds });
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-slide-layout-staged", `AI staged ${requests.length} measured object edits on slide ${slideNumber} of ${deck.name}; source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        const geometry = proposal.changes.flatMap((change) => change.kind === "geometry" ? change.geometryCommands ?? [] : []).filter((command) => command.slideNumber === slideNumber);
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode }, slideNumber, commands: geometry, resetObjectIds, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_slide_visual_design") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging visual design.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Native visual polish requires a Designer Cleanup or native composition scope.");
        const slideNumber = Number(request.input.slideNumber);
        const rawTextStyles = Array.isArray(request.input.textStyles) ? request.input.textStyles as Array<Record<string, unknown>> : [];
        const rawDecorations = Array.isArray(request.input.decorations) ? request.input.decorations as Array<Record<string, unknown>> : [];
        const textStyles = rawTextStyles.map((style) => ({
          objectId: String(style.objectId ?? ""),
          fontSizePt: style.fontSizePt === undefined ? undefined : Number(style.fontSizePt),
          bold: style.bold === undefined ? undefined : style.bold === true,
          italic: style.italic === undefined ? undefined : style.italic === true,
          color: style.color === undefined ? undefined : String(style.color),
          alignment: style.alignment === "left" || style.alignment === "center" || style.alignment === "right" ? style.alignment as "left" | "center" | "right" : undefined,
          verticalAlignment: style.verticalAlignment === "top" || style.verticalAlignment === "middle" || style.verticalAlignment === "bottom" ? style.verticalAlignment as "top" | "middle" | "bottom" : undefined,
          insetsInches: style.insetsInches && typeof style.insetsInches === "object" ? {
            top: Number((style.insetsInches as Record<string, unknown>).top), right: Number((style.insetsInches as Record<string, unknown>).right), bottom: Number((style.insetsInches as Record<string, unknown>).bottom), left: Number((style.insetsInches as Record<string, unknown>).left),
          } : undefined,
          paragraphStyle: style.paragraphStyle && typeof style.paragraphStyle === "object" ? {
            lineSpacingMultiple: (style.paragraphStyle as Record<string, unknown>).lineSpacingMultiple === undefined ? undefined : Number((style.paragraphStyle as Record<string, unknown>).lineSpacingMultiple),
            spaceAfterPt: (style.paragraphStyle as Record<string, unknown>).spaceAfterPt === undefined ? undefined : Number((style.paragraphStyle as Record<string, unknown>).spaceAfterPt),
            bulletLeftMarginInches: (style.paragraphStyle as Record<string, unknown>).bulletLeftMarginInches === undefined ? undefined : Number((style.paragraphStyle as Record<string, unknown>).bulletLeftMarginInches),
            bulletHangingInches: (style.paragraphStyle as Record<string, unknown>).bulletHangingInches === undefined ? undefined : Number((style.paragraphStyle as Record<string, unknown>).bulletHangingInches),
          } : undefined,
          rationale: String(style.rationale ?? ""),
          author: "ai" as const,
        }));
        const decorations = rawDecorations.map((decoration) => ({
          id: String(decoration.id ?? ""),
          name: String(decoration.name ?? ""),
          geometry: { x: Number(decoration.xInches) * 914_400, y: Number(decoration.yInches) * 914_400, width: Number(decoration.widthInches) * 914_400, height: Number(decoration.heightInches) * 914_400 },
          fillColor: decoration.fillColor === undefined ? undefined : String(decoration.fillColor),
          lineColor: decoration.lineColor === undefined ? undefined : String(decoration.lineColor),
          lineWidthPt: Number(decoration.lineWidthPt ?? 0),
          behindContent: decoration.behindContent !== false,
          rationale: String(decoration.rationale ?? ""),
          author: "ai" as const,
        }));
        const clearPendingLayoutRemap = request.input.clearPendingLayoutRemap === true;
        const removeDecorationIds = Array.isArray(request.input.removeDecorationIds) ? request.input.removeDecorationIds.map(String) : [];
        const proposal = createVisualDesignProposal(deck, current.project.updatedAt, { slideNumber, clearPendingLayoutRemap, removeDecorationIds, textStyles, decorations });
        const next = touchProject({ ...current, designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)), decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item) }, "mcp-slide-visual-design-staged", `AI staged editable ORNL visual hierarchy and brand geometry on slide ${slideNumber} of ${deck.name}; explicitly addressed comments were cleared and source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode }, slideNumber, textStyleCount: textStyles.length, decorationCount: decorations.length, removedDecorationCount: removeDecorationIds.length, clearedLayoutRemap: clearPendingLayoutRemap, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_slide_native_layout") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read a fresh design work order before staging a native layout.");
        if (!templateCatalog || !templateSourceBytes) throw new Error("Install an authorized PowerPoint Template Pack before staging a native layout.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Native layout remapping requires a Designer Cleanup or native composition scope; cleanup-only cannot change master/layout relationships.");
        const source = sourceForDeck(current, deck);
        if (!source?.bytes) throw new Error("The embedded source deck is unavailable.");
        const slideNumber = Number(request.input.slideNumber);
        const layoutId = String(request.input.layoutId ?? "");
        const layout = templateCatalog.layouts.find((item) => item.id === layoutId);
        if (!layout?.semantic) throw new Error("The requested approved layout is unavailable or has no semantic contract.");
        const [nativeRender, measurement] = await Promise.all([getOrBuildInspectionRender(deck, "current", current), getOrBuildNativeMeasurement(deck, "current", current)]);
        const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber, projectUpdatedAt: current.project.updatedAt, templateCatalog, currentRender: nativeRender, currentMeasurement: measurement, threads: current.designThreads });
        if (request.input.workOrderRevision !== workOrder.revision) throw new Error("The design work order is stale. Read get_slide_design_work_order again before staging a native layout.");
        const compatibility = rankLayoutCompatibility(templateCatalog.layouts, contentProfileForSlide(deck, slideNumber)).find((item) => item.layoutId === layout.id);
        if (!compatibility || compatibility.status === "incompatible") throw new Error(`${layout.name} is incompatible with the current exact-content profile.`);
        if (compatibility.status === "poor" && request.input.allowPoorLayout !== true) throw new Error(`${layout.name} is a poor deterministic fit. Choose a compatible layout or explicitly allow the poor fit with a concrete rationale.`);
        const rationale = String(request.input.rationale ?? "").trim();
        const command = {
          id: `native-layout-slide-${slideNumber}`,
          slideNumber,
          templateSha256: templateCatalog.sha256,
          templateLayoutPart: layout.sourcePart,
          templateLayoutSha256: await templateLayoutPartSha256(templateSourceBytes, layout.sourcePart),
          templateLayoutName: layout.name,
          rationale: rationale.slice(0, 1_000),
          author: "ai" as const,
        };
        const proposal = createNativeLayoutProposal(deck, current.project.updatedAt, command);
        proposal.designDecision = {
          kind: "semantic-recomposition",
          workOrderRevision: workOrder.revision,
          targetLayoutId: layout.id,
          targetLayoutName: layout.name,
          targetLayoutSourcePart: layout.sourcePart,
          compatibilityScore: compatibility.score,
          compatibilityStatus: compatibility.status,
          rationale: rationale.slice(0, 1_000),
          bindingCount: 1,
          application: "cloned-native-layout",
        };
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-native-layout-staged", `AI staged an approved native ${layout.name} layout remap for slide ${slideNumber} of ${deck.name}; explicitly addressed comments were cleared and source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return {
          proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode, designDecision: proposal.designDecision },
          slideNumber,
          targetLayout: { id: layout.id, name: layout.name, sourcePart: layout.sourcePart, layoutSha256: command.templateLayoutSha256, compatibility },
          projectUpdatedAt: next.project.updatedAt,
          applied: false,
          saved: false,
          instruction: "Presentation Studio will reuse an exact approved native layout already present in the deck when possible; otherwise it clones the guarded master/layout/theme/media dependency graph. It remaps compatible placeholder identities and repoints only this slide. Inspect the authoritative Current/Proposal PowerPoint render and reject any visual regression before acceptance.",
        };
      }
      if (request.operation === "stage_slide_recomposition") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read a fresh design work order before staging recomposition.");
        if (!templateCatalog || !templateSourceBytes) throw new Error("Install an authorized PowerPoint Template Pack before staging recomposition.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current audit and hybrid scene.");
        if (deck.operationScope !== "reflow" && deck.operationScope !== "compose") throw new Error("Semantic recomposition requires a Designer Cleanup or native composition scope; cleanup-only permits only bounded cleanup rules.");
        const slideNumber = Number(request.input.slideNumber);
        const layoutId = String(request.input.layoutId ?? "");
        const layout = templateCatalog.layouts.find((item) => item.id === layoutId);
        if (!layout?.semantic) throw new Error("The requested approved layout is unavailable or has no semantic slot contract.");
        const [nativeRender, measurement] = await Promise.all([getOrBuildInspectionRender(deck, "current", current), getOrBuildNativeMeasurement(deck, "current", current)]);
        const workOrder = buildSlideDesignWorkOrder({ deck, slideNumber, projectUpdatedAt: current.project.updatedAt, templateCatalog, currentRender: nativeRender, currentMeasurement: measurement, threads: current.designThreads });
        if (request.input.workOrderRevision !== workOrder.revision) throw new Error("The design work order is stale. Read get_slide_design_work_order again before staging recomposition.");
        const compatibility = rankLayoutCompatibility(templateCatalog.layouts, contentProfileForSlide(deck, slideNumber)).find((item) => item.layoutId === layout.id);
        if (!compatibility || compatibility.status === "incompatible") throw new Error(`${layout.name} is incompatible with the current exact-content profile.`);
        if (compatibility.status === "poor" && request.input.allowPoorLayout !== true) throw new Error(`${layout.name} is a poor deterministic fit. Choose a compatible layout or explicitly allow the poor fit with a concrete rationale.`);
        const rawBindings = Array.isArray(request.input.bindings) ? request.input.bindings as Array<Record<string, unknown>> : [];
        const bindings: SemanticSlotBinding[] = rawBindings.map((binding) => ({
          objectId: String(binding.objectId ?? ""),
          slotId: String(binding.slotId ?? ""),
          fit: binding.fit === "contain" || binding.fit === "align-horizontal" ? binding.fit : "fill",
          insetInches: Number(binding.insetInches ?? 0),
        }));
        const rationale = String(request.input.rationale ?? "").trim();
        const recomposition = semanticRecompositionRequests({ deck, slideNumber, layout, bindings, rationale });
        const nativeLayoutCommand = {
          id: `native-layout-slide-${slideNumber}`,
          slideNumber,
          templateSha256: templateCatalog.sha256,
          templateLayoutPart: layout.sourcePart,
          templateLayoutSha256: await templateLayoutPartSha256(templateSourceBytes, layout.sourcePart),
          templateLayoutName: layout.name,
          rationale: `${rationale} Apply the real approved native layout before semantic object placement.`.trim().slice(0, 1_000),
          author: "ai" as const,
        };
        const proposal = createNativeLayoutRecompositionProposal(deck, current.project.updatedAt, nativeLayoutCommand, recomposition.requests);
        proposal.designDecision = {
          kind: "semantic-recomposition",
          workOrderRevision: workOrder.revision,
          targetLayoutId: layout.id,
          targetLayoutName: layout.name,
          targetLayoutSourcePart: layout.sourcePart,
          compatibilityScore: compatibility.score,
          compatibilityStatus: compatibility.status,
          rationale: rationale.slice(0, 1_000),
          bindingCount: bindings.length,
          application: "cloned-native-layout",
        };
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-slide-recomposition-staged", `AI staged native ${layout.name} remapping plus ${bindings.length} semantic bindings on slide ${slideNumber} of ${deck.name}; explicitly addressed comments were cleared and source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return {
          proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode, designDecision: proposal.designDecision },
          slideNumber,
          targetLayout: { id: layout.id, name: layout.name, sourcePart: layout.sourcePart, compatibility },
          bindings,
          unboundObjectIds: recomposition.unboundObjectIds,
          projectUpdatedAt: next.project.updatedAt,
          applied: false,
          saved: false,
          instruction: "Inspect the authoritative Current/Proposal comparison. The proposal atomically applies the real approved native layout and places bound objects in its semantic zones. Unbound objects remain preserved in place; reject any visual regression before acceptance.",
        };
      }
      throw new Error(`Unknown Presentation Studio MCP operation: ${request.operation}`);
      })();
      const issueLedger = typeof result === "object" && result && "issueLedger" in result ? (result as { issueLedger?: { issueCount?: number; autoFixableCount?: number; phase?: string } }).issueLedger : undefined;
      const completedPhase: McpActivityPhase = issueLedger?.phase === "found-issues" ? "found-issues" : startingPhase === "rechecking" ? "rechecking" : startingPhase === "fixing" ? "fixing" : "ready";
      setMcpActivity({ id: activityId, operation: request.operation, state: "completed", phase: completedPhase, issueCount: issueLedger?.issueCount, autoFixableCount: issueLedger?.autoFixableCount });
      window.setTimeout(() => setMcpActivity((current) => current?.id === activityId ? undefined : current), 8000);
      return result;
      } catch (caught) {
        setMcpActivity({ id: activityId, operation: request.operation, state: "failed", phase: "attention" });
        window.setTimeout(() => setMcpActivity((current) => current?.id === activityId ? undefined : current), 10000);
        throw caught;
      }
    });
  }, [desktop, getOrBuildNativeRender, getOrBuildProposalCatalog, getOrBuildSlideCatalog, getOrBuildTemplateNativeRender, mcpEnabled, presentationFontCss, templateCatalog, templateNativeLoading, templateNativeRender]);

  function clearMessages() { setNotice(undefined); setError(undefined); }

  function updateAiSessionAccess(enabled: boolean) {
    const current = projectRef.current;
    const resources = resourcesWithAiSessionAccess(current.resources, enabled);
    const next = resources === current.resources ? current : { ...current, resources };
    projectRef.current = next;
    setProject(next);
    setMcpEnabled(enabled);
    setNotice(enabled
      ? `AI access is on. ${resources.length} embedded Resource${resources.length === 1 ? " is" : "s are"} automatically shared at the highest supported level.`
      : "AI access is off. No project Resources are shared with MCP models.");
  }

  function invalidateStudioQualification(deckId: string) {
    const current = studioDeckQualificationsRef.current[deckId];
    if (current) studioDeckQualificationHistoryRef.current = { ...studioDeckQualificationHistoryRef.current, [deckId]: [...(studioDeckQualificationHistoryRef.current[deckId] ?? []), current].slice(-5) };
    const next = { ...studioDeckQualificationsRef.current };
    delete next[deckId];
    studioDeckQualificationsRef.current = next;
    setStudioDeckQualifications(next);
  }

  async function importFiles(files: PickedBinaryFile[], intent: "decks" | "resources") {
    if (files.length === 0) return;
    clearMessages();
    setBusy(`Processing 0 of ${files.length} files…`);
    try {
      const importedResources: ProjectResource[] = [];
      const importedDecks: DeckJob[] = [];
      const knownResources = new Map(projectRef.current.resources.map((resource) => [resource.sha256, resource]));
      const knownDeckHashes = new Set(projectRef.current.decks.map((deck) => deck.sourceSha256));
      const failures: string[] = [];
      let duplicateCount = 0;
      let auditFailureCount = 0;
      let projectedPackageBytes = projectRef.current.resources.reduce((sum, resource) => sum + resource.byteLength + (resource.derivatives?.reduce((derivativeSum, derivative) => derivativeSum + derivative.byteLength, 0) ?? 0), 0);
      for (let index = 0; index < files.length; index += 1) {
        try {
          if (intent === "decks" && !isPowerPointResource(files[index].name)) throw new Error("Only .pptx files can enter the deck audit queue.");
          setBusy(`Processing ${index + 1} of ${files.length}: ${files[index].name}`);
          const bytes = bytesFrom(files[index].bytes);
          const processed = resourceWithAiSessionAccess(await processResourceInput({ name: files[index].name, filePath: files[index].filePath, mediaType: files[index].mediaType, bytes }), mcpEnabled);
          const existing = knownResources.get(processed.sha256);
          const resource = existing ?? processed;
          if (existing) duplicateCount += 1;
          else {
            const addedBytes = resource.byteLength + (resource.derivatives?.reduce((sum, derivative) => sum + derivative.byteLength, 0) ?? 0);
            if (projectedPackageBytes + addedBytes > MAX_PROJECT_RESOURCE_BYTES) throw new Error("Adding this file would exceed the 1.25 GB packaged Resource limit.");
            projectedPackageBytes += addedBytes;
            knownResources.set(resource.sha256, resource);
            importedResources.push(resource);
          }

          if (isPowerPointResource(files[index].name) && !knownDeckHashes.has(resource.sha256)) {
            knownDeckHashes.add(resource.sha256);
            try {
              const audit = await auditPptx(bytes);
              const adoptedAt = new Date().toISOString();
              const isOrnl = audit.classification === "current-ornl" || audit.classification === "older-or-modified-ornl";
              const sourceTemplateId = audit.classification === "sponsor" ? "sponsor-source" : audit.classification === "custom" ? "custom-source" : undefined;
              const importedDeck: DeckJob = {
                id: crypto.randomUUID(),
                name: files[index].name,
                sourceResourceId: resource.id,
                sourceSha256: resource.sha256,
                operationScope: "cleanup-only",
                templateClassification: audit.classification,
                targetTemplateId: isOrnl ? PRESENTATION_DESIGN_STANDARD.defaults.template.id : sourceTemplateId,
                targetTemplateConfirmedAt: isOrnl || sourceTemplateId ? adoptedAt : undefined,
                targetTemplateDecisionSource: isOrnl ? "automatic-default" : sourceTemplateId ? "automatic-source-preservation" : undefined,
                designProfile: isOrnl ? createOrnlDesignProfile("automatic-default", adoptedAt) : undefined,
                status: isOrnl ? "ready-for-cleanup" : sourceTemplateId ? "audited" : "needs-template-decision",
                audit,
                protectedSlideNumbers: [],
              };
              importedDecks.push(withCompiledScene(importedDeck));
            } catch (auditError) {
              auditFailureCount += 1;
              importedDecks.push({ id: crypto.randomUUID(), name: files[index].name, sourceResourceId: resource.id, sourceSha256: resource.sha256, operationScope: "audit-only", templateClassification: "unknown", status: "failed", failureMessage: (auditError instanceof Error ? auditError.message : "The PowerPoint audit failed.").slice(0, 700), protectedSlideNumbers: [] });
            }
          }
        } catch (fileError) {
          failures.push(`${files[index].name}: ${fileError instanceof Error ? fileError.message : "Processing failed."}`);
        }
      }
      if (importedResources.length > 0 || importedDecks.length > 0) {
        setProject((current) => touchProject({ ...current, resources: [...current.resources, ...importedResources], decks: [...current.decks, ...importedDecks] }, "resources-added", `Embedded ${importedResources.length} new Resource${importedResources.length === 1 ? "" : "s"}; ${importedDecks.length} PowerPoint deck${importedDecks.length === 1 ? "" : "s"} entered the audit queue.`));
      }
      if (importedDecks[0]) setSelectedDeckId(importedDecks[0].id);
      if (intent === "resources") setActiveView("resources");
      if (failures.length === files.length) setError(failures.slice(0, 3).join(" "));
      else {
        const parts = [`${importedResources.length} new Resource${importedResources.length === 1 ? "" : "s"} embedded`];
        if (importedDecks.length > 0) parts.push(`${importedDecks.length - auditFailureCount} PowerPoint audit${importedDecks.length - auditFailureCount === 1 ? "" : "s"} completed`);
        if (duplicateCount > 0) parts.push(`${duplicateCount} duplicate${duplicateCount === 1 ? "" : "s"} reused`);
        if (failures.length > 0) parts.push(`${failures.length} file${failures.length === 1 ? "" : "s"} skipped`);
        setNotice(`${parts.join(" · ")}. Originals were not changed.`);
      }
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The files could not be added to Resources."); }
    finally { setBusy(undefined); }
  }

  async function addDecks() {
    if (desktop) {
      const result = await desktop.pickPowerPoints();
      if (!result.canceled) await importFiles(result.files, "decks");
      return;
    }
    document.getElementById("web-deck-picker")?.click();
  }

  async function addResources() {
    if (desktop) {
      const result = await desktop.pickResources();
      if (!result.canceled) await importFiles(result.files, "resources");
      return;
    }
    document.getElementById("web-resource-picker")?.click();
  }

  async function loadTemplate(file: PickedBinaryFile, persist: boolean) {
    clearMessages();
    setTemplateLoading(true);
    setBusy(`Reading ${file.name} master and layouts…`);
    try {
      const bytes = bytesFrom(file.bytes);
      const catalog = await buildTemplateCatalog(bytes, file.name);
      let installedAt: string | undefined;
      if (persist) {
        if (!desktop) throw new Error("Persistent Template Pack installation requires the Electron desktop app.");
        const installed = await desktop.installTemplate({ name: file.name, sha256: catalog.sha256, bytes });
        installedAt = installed.installedAt;
      }
      setTemplateCatalog(catalog);
      setTemplateSourceBytes(bytes);
      setTemplateNativeRender(undefined);
      setTemplateInstalledAt(installedAt);
      setActiveView("designs");
      setNotice(`${catalog.layouts.length} slide designs loaded from ${catalog.name}. The authorized template remains local and no slides were changed.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The template designs could not be loaded.");
    } finally {
      setBusy(undefined);
      setTemplateLoading(false);
    }
  }

  async function installTemplate() {
    if (desktop) {
      const result = await desktop.pickTemplate();
      if (!result.canceled && result.file) await loadTemplate(result.file, true);
      return;
    }
    document.getElementById("web-template-picker")?.click();
  }

  async function importWebFiles(fileList: FileList | null, intent: "decks" | "resources") {
    const files = [...(fileList ?? [])];
    const oversized = files.find((file) => file.size > MAX_SINGLE_RESOURCE_BYTES);
    if (oversized) { setError(`${oversized.name} exceeds the 1 GB per-Resource limit.`); return; }
    const picked: PickedBinaryFile[] = [];
    for (const file of files) picked.push({ name: file.name, filePath: file.name, mediaType: file.type, bytes: new Uint8Array(await file.arrayBuffer()) });
    await importFiles(picked, intent);
  }

  function resetProjectRenderState() {
    slideCatalogsRef.current.clear();
    proposalCatalogsRef.current.clear();
    nativeRenderCatalogsRef.current.clear();
    nativeMeasurementsRef.current.clear();
    inspectionRendersRef.current.clear();
    studioDeckBuildsRef.current = {};
    studioDeckQualificationsRef.current = {};
    studioDeckQualificationHistoryRef.current = {};
    setSlideCatalogs({});
    setProposalCatalogs({});
    setNativeRenderCatalogs({});
    setStudioDeckBuilds({});
    setStudioDeckQualifications({});
    setSlideCatalogLoadingDeckId(undefined);
    setProposalCatalogLoadingDeckId(undefined);
    setNativeRenderLoadingKey(undefined);
    setSlideWorkspaceRequest(undefined);
  }

  function adoptOpenedProject(opened: PresentationStudioProject, password?: string) {
    resetProjectRenderState();
    const accessible = { ...opened, resources: resourcesWithAiSessionAccess(opened.resources, mcpEnabled) };
    projectRef.current = accessible;
    setProject(accessible);
    setSelectedDeckId(accessible.decks[0]?.id);
    setSecureAutosavePassword(password);
    setActiveView("batch");
    setNotice(`Opened ${opened.project.name}; all embedded resource hashes passed validation.`);
  }

  async function openProjectFile(file: PickedBinaryFile) {
    clearMessages();
    setBusy(`Opening and validating ${file.name}…`);
    try {
      let bytes = bytesFrom(file.bytes);
      if (bytes.byteLength > MAX_PROJECT_PACKAGE_BYTES) throw new Error("This Presentation Studio project exceeds the 1.5 GB package limit.");
      let password: string | undefined;
      if (isEncryptedProject(bytes)) {
        password = window.prompt("Enter the password for this encrypted Presentation Studio project.") ?? undefined;
        if (!password) return;
        bytes = await decryptProjectPackage(bytes, password);
      }
      adoptOpenedProject(await openProjectPackage(bytes), password);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The project could not be opened.");
    } finally {
      setBusy(undefined);
    }
  }

  function handleDragEnter(event: DragEvent<HTMLDivElement>) {
    if (![...event.dataTransfer.types].includes("Files")) return;
    event.preventDefault();
    fileDragDepth.current += 1;
    setFileDragActive(true);
  }

  function handleDragOver(event: DragEvent<HTMLDivElement>) {
    if (![...event.dataTransfer.types].includes("Files")) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  }

  function handleDragLeave(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    fileDragDepth.current = Math.max(0, fileDragDepth.current - 1);
    if (fileDragDepth.current === 0) setFileDragActive(false);
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    fileDragDepth.current = 0;
    setFileDragActive(false);
    if (event.dataTransfer.files.length === 0) return;
    const files = [...event.dataTransfer.files];
    try {
      const projectFile = projectPackageFromDrop(files);
      if (!projectFile) {
        void importWebFiles(event.dataTransfer.files, "resources");
        return;
      }
      if (projectFile.size > MAX_PROJECT_PACKAGE_BYTES) throw new Error("This Presentation Studio project exceeds the 1.5 GB package limit.");
      void projectFile.arrayBuffer()
        .then((buffer) => openProjectFile({ name: projectFile.name, filePath: projectFile.name, mediaType: projectFile.type, bytes: new Uint8Array(buffer) }))
        .catch((caught) => setError(caught instanceof Error ? caught.message : "The dropped project could not be read."));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The dropped files could not be opened.");
    }
  }

  function removeResource(resourceId: string) {
    const current = projectRef.current;
    try {
      const impact = resourceRemovalImpact(current, resourceId);
      const dependentSummary = impact.linkedDeckIds.length > 0
        ? `\n\nThis Resource is the embedded source for ${impact.linkedDeckIds.length} deck${impact.linkedDeckIds.length === 1 ? "" : "s"}. Removing it will also remove the linked deck work, ${impact.removedExemplarCount} style exemplar${impact.removedExemplarCount === 1 ? "" : "s"}, and ${impact.removedThreadCount} design comment${impact.removedThreadCount === 1 ? "" : "s"} from this project.`
        : impact.removedConceptReferenceCount > 0 ? `\n\nThis will also detach ${impact.removedConceptReferenceCount} concept reference${impact.removedConceptReferenceCount === 1 ? "" : "s"} from Studio slides. The slides and source PowerPoint remain in the project.` : "";
      const confirmed = window.confirm(`Remove “${impact.resource.name}” from this project?${dependentSummary}\n\nThe original file on disk will not be changed or deleted.`);
      if (!confirmed) return;
      const result = removeResourceFromProject(current, resourceId);
      const removedDeckIds = new Set(result.impact.linkedDeckIds);
      for (const cache of [slideCatalogsRef.current, proposalCatalogsRef.current, nativeRenderCatalogsRef.current, nativeMeasurementsRef.current, inspectionRendersRef.current]) {
        for (const key of [...cache.keys()]) if ([...removedDeckIds].some((deckId) => key === deckId || key.startsWith(`${deckId}:`))) cache.delete(key);
      }
      const withoutRemovedDecks = <T,>(catalog: Record<string, T>) => Object.fromEntries(Object.entries(catalog).filter(([key]) => ![...removedDeckIds].some((deckId) => key === deckId || key.startsWith(`${deckId}:`)))) as Record<string, T>;
      setSlideCatalogs(withoutRemovedDecks);
      setProposalCatalogs(withoutRemovedDecks);
      setNativeRenderCatalogs(withoutRemovedDecks);
      studioDeckBuildsRef.current = withoutRemovedDecks(studioDeckBuildsRef.current);
      studioDeckQualificationsRef.current = withoutRemovedDecks(studioDeckQualificationsRef.current);
      studioDeckQualificationHistoryRef.current = withoutRemovedDecks(studioDeckQualificationHistoryRef.current);
      setStudioDeckBuilds(studioDeckBuildsRef.current);
      setStudioDeckQualifications(studioDeckQualificationsRef.current);
      setSlideCatalogLoadingDeckId((id) => id && removedDeckIds.has(id) ? undefined : id);
      setProposalCatalogLoadingDeckId((id) => id && removedDeckIds.has(id) ? undefined : id);
      setNativeRenderLoadingKey((key) => key && [...removedDeckIds].some((id) => key.startsWith(`${id}:`)) ? undefined : key);
      setSlideWorkspaceRequest((request) => request && removedDeckIds.has(request.deckId) ? undefined : request);
      projectRef.current = result.project;
      setProject(result.project);
      setSelectedDeckId((id) => id && removedDeckIds.has(id) ? result.project.decks[0]?.id : id);
      setNotice(`${result.impact.resource.name} was removed from this project${result.impact.linkedDeckIds.length > 0 ? ` with ${result.impact.linkedDeckIds.length} linked deck${result.impact.linkedDeckIds.length === 1 ? "" : "s"}` : ""}. The original file was not changed or deleted.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Resource could not be removed from this project.");
    }
  }

  function confirmTemplate(templateId: string) {
    if (!selectedDeck) return;
    const now = new Date().toISOString();
    setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? withCompiledScene({ ...deck, targetTemplateId: templateId, targetTemplateConfirmedAt: now, targetTemplateDecisionSource: "user-selected", designProfile: templateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id ? createOrnlDesignProfile("user-selected", now) : undefined, status: templateId === "ornl-16x9-v1" ? "ready-for-cleanup" : "audited", proposal: undefined }) : deck) }, "target-template-confirmed", `Confirmed ${templateId} for ${selectedDeck.name}.`));
    setNotice("Target template confirmed. No slide content was changed.");
  }

  function startOrnlCleanup() {
    if (!selectedDeck?.audit) return;
    clearMessages();
    const current = projectRef.current;
    const sourceDeck = current.decks.find((deck) => deck.id === selectedDeck.id);
    if (!sourceDeck) return;
    const now = new Date().toISOString();
    const adoptedDeck: DeckJob = withCompiledScene({ ...sourceDeck, operationScope: "reflow", targetTemplateId: PRESENTATION_DESIGN_STANDARD.defaults.template.id, targetTemplateConfirmedAt: now, targetTemplateDecisionSource: sourceDeck.designProfile ? sourceDeck.targetTemplateDecisionSource ?? "automatic-default" : "user-selected", designProfile: sourceDeck.designProfile ?? createOrnlDesignProfile("user-selected", now), status: "ready-for-cleanup", proposal: undefined });
    const adoptedProject = touchProject({ ...current, decks: current.decks.map((deck) => deck.id === adoptedDeck.id ? adoptedDeck : deck) }, "ornl-defaults-adopted", `Adopted Presentation Design Standard ${PRESENTATION_DESIGN_STANDARD.version} for ${adoptedDeck.name}.`);
    try {
      const proposal = createDesignerCleanupProposal(adoptedDeck, adoptedProject.project.updatedAt);
      setProject({ ...adoptedProject, decks: adoptedProject.decks.map((deck) => deck.id === adoptedDeck.id ? { ...deck, proposal, status: "proposal-ready" } : deck), activity: [...adoptedProject.activity, { id: crypto.randomUUID(), at: now, action: "designer-cleanup-proposal-staged", detail: `Reviewed every slide and staged supported designer cleanup for ${adoptedDeck.name}; nothing was applied.` }] });
      setActiveView("review");
      setNotice("ORNL defaults adopted. Every slide was reviewed, supported designer improvements were staged, and complex exceptions are identified in one review screen.");
    } catch (caught) {
      setProject(adoptedProject);
      setActiveView("slides");
      const message = caught instanceof Error ? caught.message : "No supported deterministic changes were found.";
      setError(message);
    }
  }

  async function initializeStudioScene() {
    if (!selectedDeck?.audit || !selectedDeck.scene) return;
    clearMessages();
    setBusy("Extracting the PowerPoint into a semantic Studio Web Scene…");
    try {
      const current = projectRef.current;
      const deck = current.decks.find((item) => item.id === selectedDeck.id);
      if (!deck?.audit || !deck.scene) throw new Error("Audit the selected deck before creating its Studio Web Scene.");
      const catalog = await getOrBuildSlideCatalog(deck, current);
      const studioScene = compileStudioWebScene(deck, catalog);
      const now = new Date().toISOString();
      const adopted = { ...deck, operationScope: "reflow" as const, targetTemplateId: deck.targetTemplateId ?? PRESENTATION_DESIGN_STANDARD.defaults.template.id, targetTemplateConfirmedAt: deck.targetTemplateConfirmedAt ?? now, targetTemplateDecisionSource: deck.targetTemplateDecisionSource ?? "automatic-default" as const, designProfile: deck.designProfile ?? createOrnlDesignProfile("automatic-default", now), studioScene, proposal: undefined, status: "ready-for-cleanup" as const };
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? adopted : item) }, "studio-web-scene-created", `Extracted ${studioScene.slides.length} slides into the constrained HTML/CSS Studio design system; source PowerPoint bytes remain unchanged.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`Studio Web Scene created for ${studioScene.slides.length} slides. Choose a shared ORNL recipe or installed template layout, then design on the canvas.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio Web Scene could not be created.");
    } finally {
      setBusy(undefined);
    }
  }

  function recomposeStudioSlide(slideNumber: number, recipe: StudioLayoutRecipe, layoutId?: string) {
    if (!selectedDeck?.studioScene) return;
    clearMessages();
    try {
      const layout = layoutId ? templateCatalog?.layouts.find((item) => item.id === layoutId) : undefined;
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber) && recipe !== "source") throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and remains source-preserved. Choose another slide to redesign.`);
      const studioScene = recomposeStudioWebSlide(selectedDeck.studioScene, slideNumber, recipe, layout, recipe === "source" && isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)
        ? `Preserve the approved ORNL template composition on slide ${slideNumber} exactly.`
        : `Recompose slide ${slideNumber} with the shared ${recipe} web layout while retaining exact source content and editable PowerPoint bindings.`);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
      setProject((current) => touchProject({ ...current, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, operationScope: "reflow", studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-slide-recomposed", `Recomposed slide ${slideNumber} with ${recipe} in the Studio Web Scene; no PowerPoint file was changed.`));
      setNotice(`Slide ${slideNumber} now uses the ${recipe.replaceAll("-", " ")} web composition. Drag elements to refine it or stage it for PowerPoint review.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The slide could not be recomposed.");
    }
  }

  function moveStudioNodes(slideNumber: number, updates: Array<{ nodeId: string; frame: StudioWebFrame }>) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be moved or resized in Studio.`);
      if (!updates.length || updates.length > 60 || new Set(updates.map((update) => update.nodeId)).size !== updates.length) throw new Error("Move between 1 and 60 unique Studio elements in one canvas transaction.");
      let studioScene = selectedDeck.studioScene;
      const slide = studioScene.slides.find((item) => item.slideNumber === slideNumber);
      if (!slide) throw new Error("The requested Studio slide is unavailable.");
      const pending = new Map(updates.map((update) => [update.nodeId, update.frame]));
      for (const treatment of slide.figureTreatments) {
        if (!treatment.nodeIds.every((id) => pending.has(id))) continue;
        const anchor = slide.nodes.find((node) => node.id === treatment.nodeIds[0]);
        const target = pending.get(treatment.nodeIds[0]);
        if (!anchor || !target) continue;
        const dx = target.x - anchor.frame.x;
        const dy = target.y - anchor.frame.y;
        if (treatment.nodeIds.some((id) => { const node = slide.nodes.find((candidate) => candidate.id === id); const frame = pending.get(id); return !node || !frame || Math.abs(frame.x - node.frame.x - dx) > 1_000 || Math.abs(frame.y - node.frame.y - dy) > 1_000; })) throw new Error("A first-class Studio figure must move as one intact relationship group.");
        studioScene = translateStudioFigureTreatment(studioScene, slideNumber, treatment.id, dx, dy);
        treatment.nodeIds.forEach((id) => pending.delete(id));
      }
      for (const [nodeId, frame] of pending) studioScene = updateStudioWebNodeFrame(studioScene, slideNumber, nodeId, frame);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio elements could not be moved.");
    }
  }

  function moveStudioNode(slideNumber: number, nodeId: string, nextFrame: StudioWebFrame) {
    moveStudioNodes(slideNumber, [{ nodeId, frame: nextFrame }]);
  }

  function arrangeStudioSelection(slideNumber: number, nodeIds: string[], mode: StudioConstraintRequest["mode"]) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be rearranged in Studio.`);
      const uniqueIds = [...new Set(nodeIds)];
      const distribution = mode === "horizontal-equal-gap" || mode === "vertical-equal-gap";
      if (uniqueIds.length < (distribution ? 3 : 2)) throw new Error(distribution ? "Select at least three elements or complete figure groups to distribute them." : "Select at least two elements or complete figure groups to align them.");
      const slide = selectedDeck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
      if (!slide) throw new Error("The requested Studio slide is unavailable.");
      const selected = new Set(uniqueIds);
      const grouped = new Set<string>();
      const groups: string[][] = [];
      for (const treatment of slide.figureTreatments) {
        const selectedMembers = treatment.nodeIds.filter((id) => selected.has(id));
        if (!selectedMembers.length) continue;
        if (selectedMembers.length !== treatment.nodeIds.length) throw new Error("Select every member of a first-class figure before aligning or distributing it.");
        groups.push([...treatment.nodeIds]);
        treatment.nodeIds.forEach((id) => grouped.add(id));
      }
      uniqueIds.filter((id) => !grouped.has(id)).forEach((id) => groups.push([id]));
      if (groups.length < (distribution ? 3 : 2)) throw new Error(distribution ? "Select at least three independent elements or complete figure groups to distribute them." : "Select at least two independent elements or complete figure groups to align them.");
      const preview = studioFreshPreviewsRef.current[`${selectedDeck.id}:${slideNumber}`];
      const measurement = preview?.sceneRevision === selectedDeck.studioScene.revision && preview.slideUpdatedAt === slide.updatedAt ? preview.nativeMeasurement : undefined;
      const result = applyStudioLayoutConstraints(selectedDeck.studioScene, slideNumber, [{ kind: distribution ? "distribute" : "align", mode, nodeIds: uniqueIds, groups, rationale: `Human arranged ${uniqueIds.length} selected Studio elements using ${mode.replaceAll("-", " ")} while preserving complete figure relationships.`, author: "human" }], measurement);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = result.scene.slides.find((item) => item.slideNumber === slideNumber)?.updatedAt ?? result.scene.revision;
      setProject((current) => touchProject({ ...current, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-selection-arranged", `Arranged ${uniqueIds.length} Studio elements on slide ${slideNumber} with ${mode}; complete figure groups remained intact and source PowerPoint bytes remain unchanged.`));
      setNotice(`${result.changedNodeIds.length ? `${result.changedNodeIds.length} elements were arranged` : "The selection already satisfied that arrangement"}. Rebuild the PowerPoint result for optical review${result.evidenceAuthority === "scene-estimate" ? "; final acceptance still requires native measurement" : ""}.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The selected Studio elements could not be arranged.");
    }
  }

  function styleStudioNode(slideNumber: number, nodeId: string, patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">>) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be restyled in Studio.`);
      const studioScene = updateStudioWebNodeStyle(selectedDeck.studioScene, slideNumber, nodeId, patch);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio element style could not be updated.");
    }
  }

  function publishStudioComponentStyle(slideNumber: number, nodeId: string) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot define a reusable Studio component.`);
      const result = adoptStudioComponentStyle(selectedDeck.studioScene, { slideNumber, nodeId });
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      setProject((current) => {
        const designThreads = result.affectedSlideNumbers.reduce((threads, affectedSlideNumber) => {
          const revision = result.scene.slides.find((slide) => slide.slideNumber === affectedSlideNumber)?.updatedAt ?? result.scene.revision;
          return markSubmittedThreadsForReanchor(threads, selectedDeck.id, affectedSlideNumber, revision);
        }, current.designThreads);
        return touchProject({ ...current, designThreads, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-component-style-published", `Published ${result.definition.name} and updated ${result.affectedNodeIds.length} compatible source-bound component instances across ${result.affectedSlideNumbers.length} slides; exact content, geometry, and source PowerPoint bytes remain unchanged.`);
      });
      setNotice(`${result.definition.name} now drives ${result.affectedNodeIds.length} compatible instances across ${result.affectedSlideNumbers.length} slides. Rebuild the affected PowerPoint results before judging the change.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The reusable Studio component could not be published.");
    }
  }

  function commitStudioComponentMutation(slideNumber: number, mutation: (scene: StudioWebScene) => StudioWebScene) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be restyled in Studio.`);
      const studioScene = mutation(selectedDeck.studioScene);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio component design could not be updated.");
    }
  }

  function styleStudioTable(slideNumber: number, nodeId: string, patch: StudioTableDesignPatch) {
    commitStudioComponentMutation(slideNumber, (scene) => updateStudioTableDesign(scene, slideNumber, nodeId, patch));
  }

  function resizeStudioTableColumnInches(slideNumber: number, nodeId: string, column: number, widthInches: number) {
    commitStudioComponentMutation(slideNumber, (scene) => resizeStudioTableColumn(scene, slideNumber, nodeId, column, widthInches * 914_400));
  }

  function resizeStudioTableRowInches(slideNumber: number, nodeId: string, row: number, heightInches: number) {
    commitStudioComponentMutation(slideNumber, (scene) => resizeStudioTableRow(scene, slideNumber, nodeId, row, heightInches * 914_400));
  }

  function styleStudioTableCell(slideNumber: number, nodeId: string, cellId: string, patch: StudioTableCellDesignPatch) {
    commitStudioComponentMutation(slideNumber, (scene) => updateStudioTableCellDesign(scene, slideNumber, nodeId, cellId, patch));
  }

  function publishStudioTableExemplarStyle(slideNumber: number, nodeId: string) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot define a table exemplar.`);
      const targetSlideNumbers = selectedDeck.studioScene.slides.filter((slide) => !isProtectedOrnlTemplateSlide(selectedDeck, slide.slideNumber)).map((slide) => slide.slideNumber);
      const result = publishStudioTableExemplar(selectedDeck.studioScene, { slideNumber, tableNodeId: nodeId, targetSlideNumbers });
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      setProject((current) => {
        const designThreads = result.affectedSlideNumbers.reduce((threads, affectedSlideNumber) => {
          const revision = result.scene.slides.find((slide) => slide.slideNumber === affectedSlideNumber)?.updatedAt ?? result.scene.revision;
          return markSubmittedThreadsForReanchor(threads, selectedDeck.id, affectedSlideNumber, revision);
        }, current.designThreads);
        return touchProject({ ...current, designThreads, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-table-exemplar-published", `Published ${result.definition.name} and applied its style to ${result.affectedTableNodeIds.length} structurally compatible tables. Exact cell copy, merge topology, and semantic fills remain unchanged.`);
      });
      setNotice(`${result.definition.name} now styles ${result.affectedTableNodeIds.length} compatible table${result.affectedTableNodeIds.length === 1 ? "" : "s"}. Rebuild their PowerPoint results for native review.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The approved Studio table exemplar could not be published.");
    }
  }

  function applyStudioTableExemplarStyle(definitionId: string) {
    if (!selectedDeck?.studioScene) return;
    try {
      const targetSlideNumbers = selectedDeck.studioScene.slides.filter((slide) => !isProtectedOrnlTemplateSlide(selectedDeck, slide.slideNumber)).map((slide) => slide.slideNumber);
      const result = applyStudioTableExemplar(selectedDeck.studioScene, { definitionId, targetSlideNumbers });
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-table-exemplar-applied", `Applied ${result.definition.name} to ${result.affectedTableNodeIds.length} compatible native tables without copying content or semantic fills.`));
      setNotice(`Applied ${result.definition.name} to ${result.affectedTableNodeIds.length} compatible table${result.affectedTableNodeIds.length === 1 ? "" : "s"}.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The approved Studio table exemplar could not be applied.");
    }
  }

  function planStudioTableContinuationFromUi(slideNumber: number, nodeId: string, maximumBodyRowsPerSlide: number) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be continued.`);
      const result = planStudioTableContinuation(selectedDeck.studioScene, { slideNumber, tableNodeId: nodeId, maximumBodyRowsPerSlide });
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene: result.scene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-table-continuation-planned", `${result.plan.status === "ready" ? `Planned ${result.plan.segments.length} merge-safe table continuation slides` : "Recorded a blocked table continuation plan"} for slide ${slideNumber}; source cell content and structure remain unchanged.`));
      setNotice(result.plan.status === "ready" ? `Continuation plan ready: the next build will create ${result.plan.segments.length} editable slides with ${result.plan.headerRows} repeated header row${result.plan.headerRows === 1 ? "" : "s"}.` : `Continuation held: ${result.plan.blockers.join(" ")}`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio table continuation could not be planned.");
    }
  }

  function clearStudioTableContinuationFromUi(slideNumber: number, nodeId: string) {
    commitStudioComponentMutation(slideNumber, (scene) => clearStudioTableContinuation(scene, { slideNumber, tableNodeId: nodeId }));
    setNotice(`Cleared the continuation plan for slide ${slideNumber}; the source-bound table remains unchanged.`);
  }

  function repairStudioObjectiveIssuesFromUi(slideNumber: number) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot enter automatic repair.`);
      const slide = selectedDeck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
      const preview = studioFreshPreviewsRef.current[`${selectedDeck.id}:${slideNumber}`];
      if (!slide || !preview || preview.slideUpdatedAt !== slide.updatedAt || preview.nativeMeasurement?.status !== "ready" || preview.nativeMeasurement.authority !== "powerpoint-native") throw new Error("Build this exact slide result in PowerPoint before running its bounded repair pass.");
      const beforeSignature = studioSlideContentSignature(slide);
      const result = applyStudioDeterministicRepairPass(selectedDeck.studioScene, slideNumber, preview.nativeMeasurement);
      const resultSlide = result.scene.slides.find((item) => item.slideNumber === slideNumber)!;
      if (studioSlideContentSignature(resultSlide) !== beforeSignature) throw new Error("Studio refused an automatic repair that changed exact source content.");
      if (!result.requiresNativeRerender) {
        setNotice(`${result.deferredIssueIds.length || result.critique.issues.length} issue${(result.deferredIssueIds.length || result.critique.issues.length) === 1 ? "" : "s"} need a material layout, table, figure, concept, or human decision; no unsafe patch was applied.`);
        return;
      }
      commitStudioComponentMutation(slideNumber, () => result.scene);
      setNotice(`Fixed ${result.fixedIssueIds.length} bounded issue${result.fixedIssueIds.length === 1 ? "" : "s"}. Rebuild this slide to rerender, remeasure, and check the new result; ${result.deferredIssueIds.length} issue${result.deferredIssueIds.length === 1 ? "" : "s"} remain for design judgment.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The bounded Studio repair pass could not be applied.");
    }
  }

  function authorStudioConnector(slideNumber: number, nodeId: string, design: StudioConnectorDesign) {
    commitStudioComponentMutation(slideNumber, (scene) => updateStudioConnectorDesign(scene, slideNumber, nodeId, design));
  }

  function setStudioFigureTreatment(slideNumber: number, treatment: StudioFigureTreatment) {
    if (!selectedDeck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(selectedDeck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot receive a figure treatment.`);
      const studioScene = updateStudioFigureTreatment(selectedDeck.studioScene, slideNumber, treatment);
      pushStudioEditHistory(selectedDeck.id, selectedDeck.studioScene);
      const nextSlideRevision = studioScene.slides.find((slide) => slide.slideNumber === slideNumber)?.updatedAt ?? studioScene.revision;
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, designThreads: markSubmittedThreadsForReanchor(current.designThreads, selectedDeck.id, slideNumber, nextSlideRevision), decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio figure treatment could not be updated.");
    }
  }

  function createVisualNeedFromStudio(slideNumber: number, type: StudioVisualNeed["type"]) {
    const current = projectRef.current;
    const deck = selectedDeck ? current.decks.find((item) => item.id === selectedDeck.id) : undefined;
    if (!deck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot enter the visual-concept queue.`);
      const copy: Record<StudioVisualNeed["type"], { reason: string; job: string }> = {
        "layout-concept": { reason: "The slide may benefit from a stronger whole-slide hierarchy and composition direction.", job: "Create clearer hierarchy, pacing, and negative space without changing the approved content." },
        "figure-concept": { reason: "The technical figure may benefit from relationship-preserving art direction.", job: "Clarify the figure's evidence hierarchy while preserving its source meaning and exact technical relationships." },
        "image-treatment": { reason: "The source imagery may benefit from a more deliberate crop, frame, and supporting-content relationship.", job: "Make the source image carry the main evidence while leaving clean editable content zones." },
        "supporting-visual": { reason: "A restrained supporting visual may improve comprehension or pacing.", job: "Add one purposeful visual role without competing with the source assertion and evidence." },
        "diagram-rebuild": { reason: "The source schematic may need an editable relationship-preserving reconstruction concept.", job: "Provide a clearer diagram system while retaining every verified label, value, arrow, group, sequence, and causal relationship." },
      };
      const scene = createStudioVisualNeed(deck.studioScene, slideNumber, { type, reason: copy[type].reason, communicationJob: copy[type].job, expression: "balanced" });
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-visual-need-created", `Created a local ${type} brief for slide ${slideNumber} of ${deck.name}; no image model was called.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`A governed ${type.replaceAll("-", " ")} brief is ready on slide ${slideNumber}. It shares abstract structure only by default.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The visual-direction brief could not be created.");
    }
  }

  function holdVisualNeedFromStudio(slideNumber: number, visualNeedId: string) {
    const current = projectRef.current;
    const deck = selectedDeck ? current.decks.find((item) => item.id === selectedDeck.id) : undefined;
    if (!deck?.studioScene) return;
    try {
      const scene = holdStudioVisualNeed(deck.studioScene, slideNumber, visualNeedId, "Held by the person in Studio; preserve the brief for later review.");
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-visual-need-held", `Held a local visual-direction brief on slide ${slideNumber} of ${deck.name}.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`The visual-direction brief on slide ${slideNumber} is held, not deleted.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The visual-direction brief could not be held.");
    }
  }

  function attachConceptFromStudio(slideNumber: number, visualNeedId: string, resourceId: string) {
    const current = projectRef.current;
    const deck = selectedDeck ? current.decks.find((item) => item.id === selectedDeck.id) : undefined;
    if (!deck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot use a concept reference.`);
      const slide = deck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
      const need = slide?.visualNeeds?.find((item) => item.id === visualNeedId);
      if (!need) throw new Error("The selected visual need is no longer available.");
      const resource = current.resources.find((item) => item.id === resourceId);
      if (!resource || resource.kind !== "image") throw new Error("Choose an embedded image Resource for this concept.");
      if (resource.mcpAccess !== "preview") throw new Error("Turn on AI access before attaching an embedded image for art direction.");
      const scene = attachStudioConceptReference(deck.studioScene, slideNumber, resource, {
        visualNeedId,
        origin: "human-reference",
        approvedInfluences: need.approvedInfluences,
        blueprint: {
          summary: need.communicationJob,
          zones: [],
          styleNotes: [`Use the ${need.expression} ORNL expression recorded in visual need ${need.id}.`],
          reconstructionNotes: ["Reconstruct only approved visual characteristics with exact source-bound Studio content; do not trace or flatten the concept."],
        },
      });
      const next = touchProject({
        ...current,
        resources: current.resources.map((item) => item.id === resource.id ? { ...item, roles: [...new Set([...item.roles, "concept-reference" as const])] } : item),
        decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item),
      }, "studio-concept-reference-attached", `Attached a concept-only image Resource to visual need ${visualNeedId} on slide ${slideNumber} of ${deck.name}.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`${resource.name} is attached as concept-only direction on slide ${slideNumber}. Rebuild the slide from editable Studio objects; do not use generated text or claims.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The concept image could not be attached.");
    }
  }

  function detachConceptFromStudio(slideNumber: number, referenceId: string) {
    const current = projectRef.current;
    const deck = selectedDeck ? current.decks.find((item) => item.id === selectedDeck.id) : undefined;
    if (!deck?.studioScene) return;
    try {
      const scene = removeStudioConceptReference(deck.studioScene, slideNumber, referenceId);
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, studioScene: scene } : item) }, "studio-concept-reference-detached", `Detached a concept-only reference from slide ${slideNumber} of ${deck.name}; the Resource remains embedded.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`The concept was detached from slide ${slideNumber}. Its Resource remains in the project.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The concept image could not be detached.");
    }
  }

  function reconstructConceptFromStudio(slideNumber: number, referenceId: string) {
    const current = projectRef.current;
    const deck = selectedDeck ? current.decks.find((item) => item.id === selectedDeck.id) : undefined;
    if (!deck?.studioScene) return;
    try {
      if (isProtectedOrnlTemplateSlide(deck, slideNumber)) throw new Error(`The approved ORNL template composition on slide ${slideNumber} is sacred and cannot be reconstructed from a visual concept.`);
      const reconstruction = reconstructStudioConcept(deck.studioScene, slideNumber, referenceId);
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, operationScope: "reflow" as const, studioScene: reconstruction.scene, proposal: undefined } : item) }, "studio-concept-reconstructed", `Reconstructed concept-only direction as editable Studio content on slide ${slideNumber} of ${deck.name}; source bytes remain unchanged.`);
      projectRef.current = next;
      setProject(next);
      setNotice(`Slide ${slideNumber} now has a material editable reconstruction. Build its PowerPoint result before judging the design.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The concept could not be reconstructed as editable Studio content.");
    }
  }

  async function sourceLockedFigureRasters(deck: DeckJob, studioScene: StudioWebScene, slideNumbers?: Set<number>) {
    if (!desktop) return undefined;
    const source = sourceForDeck(projectRef.current, deck);
    if (!source?.bytes?.byteLength) throw new Error("The embedded source PowerPoint is required to isolate source-locked technical figures.");
    const result: Record<string, { data: string; width: number; height: number }> = {};
    for (const slide of studioScene.slides) {
      if (slideNumbers && !slideNumbers.has(slide.slideNumber)) continue;
      for (const treatment of slide.figureTreatments.filter((item) => ["preserve-as-unit", "preserve-and-frame"].includes(item.mode) && ["source-locked", "verified"].includes(item.verificationStatus))) {
        const shapeIds = nativeIsolationShapeIds(slide, treatment);
        if (!shapeIds.length) continue;
        const cacheKey = `${deck.sourceSha256}:${slide.slideNumber}:${[...shapeIds].sort().join(",")}`;
        let raster = sourceFigureRastersRef.current.get(cacheKey);
        if (!raster) {
          const isolated = await isolateNativePowerPointObjects({ sourceBytes: bytesFrom(source.bytes), slideNumber: slide.slideNumber, shapeIds });
          const rendered = await desktop.renderPowerPoint({ name: `${cleanFileStem(deck.name)}_slide-${slide.slideNumber}_isolated-figure.pptx`, bytes: isolated.bytes, width: 1600, format: "png" });
          const nativeSlide = rendered.status === "ready" && rendered.authoritative && rendered.slides.length === 1 ? rendered.slides[0] : undefined;
          if (!nativeSlide) throw new Error(rendered.reason ?? `Microsoft PowerPoint could not isolate the technical figure on slide ${slide.slideNumber}.`);
          raster = { data: `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}`, width: nativeSlide.width, height: nativeSlide.height };
          sourceFigureRastersRef.current.set(cacheKey, raster);
          while (sourceFigureRastersRef.current.size > 80) sourceFigureRastersRef.current.delete(sourceFigureRastersRef.current.keys().next().value as string);
        }
        result[treatment.id] = raster;
      }
    }
    return Object.keys(result).length ? result : undefined;
  }

  async function buildFreshStudioPreview(deck: DeckJob, slideNumber: number, studioScene: StudioWebScene): Promise<StudioFreshPreview> {
    if (!deck.audit) throw new Error("Audit the source PowerPoint before building a fresh composition.");
    const sourceSlide = deck.audit.slides.find((item) => item.number === slideNumber);
    const studioSlide = studioScene.slides.find((item) => item.slideNumber === slideNumber);
    if (!sourceSlide || !studioSlide) throw new Error(`Slide ${slideNumber} is not present in the current Studio scene.`);
    const oneSlideScene: StudioWebScene = { ...studioScene, slides: [studioSlide] };
    const sourceCatalog = await getOrBuildSlideCatalog(deck, projectRef.current);
    const catalog = catalogWithStudioResources(sourceCatalog, await studioResourceMedia(oneSlideScene, projectRef.current.resources), oneSlideScene);
    const sourceRender = await getOrBuildNativeRender(deck, "current", projectRef.current);
    const sourceRaster = sourceRender?.status === "ready" ? sourceRender.slides.find((slide) => slide.number === slideNumber) : undefined;
    const templateRender = studioSlide.recipe === "template-layout" ? await getOrBuildTemplateNativeRender() : undefined;
    const templateLayoutIndex = studioSlide.targetLayoutId ? templateCatalog?.layouts.findIndex((layout) => layout.id === studioSlide.targetLayoutId) ?? -1 : -1;
    const templateLayoutRaster = templateRender?.status === "ready" && templateLayoutIndex >= 0 ? templateRender.slides.find((slide) => slide.number === templateLayoutIndex + 1) : undefined;
    const sourceFigureRasters = await sourceLockedFigureRasters(deck, oneSlideScene, new Set([slideNumber]));
    const result = await buildStudioCompositionPptx(oneSlideScene, {
      catalog,
      templateCatalog,
      sourceSlideRasters: sourceRaster ? { [slideNumber]: { data: `data:${sourceRaster.mimeType};base64,${bytesToBase64(bytesFrom(sourceRaster.bytes))}`, width: sourceRaster.width, height: sourceRaster.height } } : undefined,
      sourceFigureRasters,
      sourceSlideText: { [slideNumber]: sourceSlide.text },
      templateLayoutRasters: templateLayoutRaster && studioSlide.targetLayoutId ? { [studioSlide.targetLayoutId]: { data: `data:${templateLayoutRaster.mimeType};base64,${bytesToBase64(bytesFrom(templateLayoutRaster.bytes))}`, width: templateLayoutRaster.width, height: templateLayoutRaster.height } } : undefined,
      strict: true,
      title: `${cleanFileStem(deck.name)} · Studio slide ${slideNumber}`,
    });
    const candidateAudit = await auditPptx(result.bytes);
    const contentValidation = validateStudioCompositionContent({ scene: oneSlideScene, sourceAudit: deck.audit, candidateAudit, outputSlides: result.outputSlides });
    if (!contentValidation.valid) throw new Error(`Fresh-composition validation rejected the candidate: ${contentValidation.errors.join(" ")}`);
    const artifactName = `${cleanFileStem(deck.name)}_slide-${slideNumber}${result.slideCount > 1 ? "_table-continuation" : ""}_studio-rebuild.pptx`;
    const nativeRender = desktop ? await desktop.renderPowerPoint({ name: artifactName, bytes: result.bytes, width: 1600, format: "png" }) : undefined;
    const nativeMeasurement = desktop ? await desktop.measurePowerPoint({ name: artifactName, bytes: result.bytes }) : undefined;
    if (nativeMeasurement?.status === "ready") {
      const overflow = nativeTextOverflows(nativeMeasurement);
      if (overflow.length) throw new Error(`Fresh-composition validation rejected the candidate because Microsoft PowerPoint measured text outside its frame: ${overflow.slice(0, 4).map((item) => `${item.name} (${item.edges.join(", ")})`).join("; ")}.`);
    }
    return { ...result, deckId: deck.id, sourceSlideNumber: slideNumber, sceneRevision: studioScene.revision, slideUpdatedAt: studioSlide.updatedAt, contentValidation, nativeRender, nativeMeasurement };
  }

  async function buildCentralStudioDeck(deck: DeckJob, studioScene: StudioWebScene): Promise<StudioDeckBuild> {
    if (!deck.audit) throw new Error("Audit the source PowerPoint before building the central presentation.");
    assertSacredOrnlTitleSlideIntegrity(deck, studioScene);
    const unconverted = unsupportedSourceSlideNumbers(deck, studioScene);
    if (unconverted.length) throw new Error(`Convert every slide into the Studio design system before exporting one central presentation. Still using source geometry: ${unconverted.join(", ")}.`);
    const sourceCatalog = await getOrBuildSlideCatalog(deck, projectRef.current);
    const catalog = catalogWithStudioResources(sourceCatalog, await studioResourceMedia(studioScene, projectRef.current.resources), studioScene);
    const sourceRender = await getOrBuildNativeRender(deck, "current", projectRef.current);
    const sourceSlideRasters = sourceRender?.status === "ready" ? Object.fromEntries(sourceRender.slides.map((slide) => [slide.number, { data: `data:${slide.mimeType};base64,${bytesToBase64(bytesFrom(slide.bytes))}`, width: slide.width, height: slide.height }])) : undefined;
    const templateRender = studioScene.slides.some((slide) => slide.recipe === "template-layout") ? await getOrBuildTemplateNativeRender() : undefined;
    const templateLayoutRasters = templateRender?.status === "ready" && templateCatalog ? Object.fromEntries(templateCatalog.layouts.map((layout, index) => {
      const slide = templateRender.slides.find((candidate) => candidate.number === index + 1);
      return [layout.id, slide ? { data: `data:${slide.mimeType};base64,${bytesToBase64(bytesFrom(slide.bytes))}`, width: slide.width, height: slide.height } : undefined];
    }).filter((entry): entry is [string, { data: string; width: number; height: number }] => Boolean(entry[1]))) : undefined;
    const sourceFigureRasters = await sourceLockedFigureRasters(deck, studioScene);
    const composition = await buildStudioCompositionPptx(studioScene, { catalog, templateCatalog, sourceSlideRasters, sourceFigureRasters, sourceSlideText: Object.fromEntries(deck.audit.slides.map((slide) => [slide.number, slide.text])), templateLayoutRasters, strict: true, title: `${cleanFileStem(deck.name)} · Presentation Studio redesign` });
    let result = composition;
    for (const sourceSlideNumber of studioScene.slides.filter((slide) => isProtectedOrnlTemplateSlide(deck, slide.slideNumber)).map((slide) => slide.slideNumber)) {
      const sourceBytes = sourceForDeck(projectRef.current, deck)?.bytes;
      if (!sourceBytes) throw new Error("The embedded source PowerPoint is required to preserve approved ORNL template slides natively.");
      const destinationSlideNumber = result.outputSlides.find((slide) => slide.sourceSlideNumber === sourceSlideNumber && !slide.continuation)?.outputSlideNumber;
      if (!destinationSlideNumber) throw new Error(`Protected ORNL source slide ${sourceSlideNumber} has no one-to-one output mapping.`);
      const preserved = await preserveNativeSlide({ destinationBytes: result.bytes, sourceBytes, sourceSlideNumber, destinationSlideNumber });
      result = { ...result, bytes: preserved.bytes, warnings: [...result.warnings, `Source slide ${sourceSlideNumber} → output slide ${destinationSlideNumber}: approved native ORNL template composition, layout, master, theme, and ${preserved.receipt.copiedMediaCount} related media part${preserved.receipt.copiedMediaCount === 1 ? "" : "s"} preserved.`] };
    }
    const candidateAudit = await auditPptx(result.bytes);
    const contentValidation = validateStudioCompositionContent({ scene: studioScene, sourceAudit: deck.audit, candidateAudit, outputSlides: result.outputSlides });
    if (!contentValidation.valid) throw new Error(`Central presentation validation rejected the candidate: ${contentValidation.errors.join(" ")}`);
    if (!desktop) throw new Error("Building the central presentation requires the Electron desktop app and Microsoft PowerPoint validation.");
    const artifactName = `${cleanFileStem(deck.name)}_studio-redesign.pptx`;
    const [nativeRender, nativeMeasurement] = await Promise.all([
      desktop.renderPowerPoint({ name: artifactName, bytes: result.bytes, width: 1600, format: "png" }),
      desktop.measurePowerPoint({ name: artifactName, bytes: result.bytes }),
    ]);
    if (nativeRender.status !== "ready" || !nativeRender.authoritative || nativeRender.slides.length !== result.outputSlides.length) throw new Error(nativeRender.reason ?? "Microsoft PowerPoint could not render every materialized slide in the central presentation.");
    if (nativeMeasurement.status !== "ready" || nativeMeasurement.authority !== "powerpoint-native") throw new Error(nativeMeasurement.reason ?? "Microsoft PowerPoint could not measure the central presentation.");
    const protectedSourceSlides = new Set(studioScene.slides.filter((slide) => isProtectedOrnlTemplateSlide(deck, slide.slideNumber)).map((slide) => slide.slideNumber));
    const protectedSlides = new Set(result.outputSlides.filter((slide) => protectedSourceSlides.has(slide.sourceSlideNumber)).map((slide) => slide.outputSlideNumber));
    const overflow = nativeTextOverflows(nativeMeasurement).filter((item) => !protectedSlides.has(item.slideNumber));
    if (overflow.length) throw new Error(`Central presentation validation found text outside its frame: ${overflow.slice(0, 6).map((item) => `${item.name} (${item.edges.join(", ")})`).join("; ")}.`);
    return {
      ...result,
      deckId: deck.id,
      sceneRevision: studioScene.revision,
      slideUpdatedAts: Object.fromEntries(studioScene.slides.map((slide) => [slide.slideNumber, slide.updatedAt])),
      contentValidation,
      nativeRender,
      nativeMeasurement,
      candidateAudit,
    };
  }

  async function previewFreshStudioSlide(slideNumber: number) {
    if (!selectedDeck?.studioScene) return;
    clearMessages();
    setBusy(`Building a fresh editable PowerPoint for Studio slide ${slideNumber}…`);
    try {
      const deck = projectRef.current.decks.find((item) => item.id === selectedDeck.id);
      if (!deck?.studioScene) throw new Error("Create or reopen the Studio Web Scene before building a fresh composition.");
      const preview = await buildFreshStudioPreview(deck, slideNumber, deck.studioScene);
      studioFreshPreviewsRef.current = { ...studioFreshPreviewsRef.current, [`${deck.id}:${slideNumber}`]: preview };
      setStudioFreshPreviews(studioFreshPreviewsRef.current);
      if (preview.nativeRender?.status === "ready" && preview.nativeMeasurement?.status === "ready") setNotice(`${preview.slideCount} fresh-composition PowerPoint output slide${preview.slideCount === 1 ? "" : "s"} from source slide ${slideNumber} passed explicit source/output copy, native-table, PowerPoint-render, and measured-text-fit guards. Compare every result in Studio before saving.`);
      else setNotice(`Fresh-composition PowerPoint for source slide ${slideNumber} passed exact-copy and native-table guards. Native visual review is still required before it can be saved.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The fresh Studio composition could not be built.");
    } finally {
      setBusy(undefined);
    }
  }

  async function previewAllFreshStudioSlides() {
    if (!selectedDeck?.studioScene) return;
    clearMessages();
    const deck = projectRef.current.decks.find((item) => item.id === selectedDeck.id);
    if (!deck?.studioScene) { setError("Create or reopen the Studio Web Scene before building deck results."); return; }
    const buildPlan = planStudioExportBuild(deck.studioScene);
    const designedCount = buildPlan.freshCompositionSlideNumbers.length;
    if (!designedCount) { setNotice("Every slide is still using its preserved source PowerPoint result. Apply a Studio design recipe before rebuilding designed slides."); return; }
    const freshSlides = buildPlan.freshCompositionSlideNumbers.map((slideNumber) => deck.studioScene!.slides.find((slide) => slide.slideNumber === slideNumber)!);
    const failed: Array<{ slideNumber: number; reason: string }> = [];
    let built = 0;
    let reused = 0;
    try {
      for (const [index, studioSlide] of freshSlides.entries()) {
        const key = `${deck.id}:${studioSlide.slideNumber}`;
        const cached = studioFreshPreviewsRef.current[key];
        if (cached?.slideUpdatedAt === studioSlide.updatedAt && cached.nativeRender?.status === "ready" && cached.nativeMeasurement?.status === "ready") {
          reused += 1;
          continue;
        }
        setBusy(`Building PowerPoint export result ${index + 1} of ${freshSlides.length}: slide ${studioSlide.slideNumber}…`);
        try {
          const preview = await buildFreshStudioPreview(deck, studioSlide.slideNumber, deck.studioScene);
          studioFreshPreviewsRef.current = { ...studioFreshPreviewsRef.current, [key]: preview };
          setStudioFreshPreviews(studioFreshPreviewsRef.current);
          built += 1;
        } catch (caught) {
          failed.push({ slideNumber: studioSlide.slideNumber, reason: caught instanceof Error ? caught.message : "PowerPoint export validation failed" });
        }
      }
      const preserved = buildPlan.preservedSourceSlideNumbers.length;
      if (failed.length) {
        setError(`Built ${built} designed slide result${built === 1 ? "" : "s"}; ${preserved} untouched source slide${preserved === 1 ? " remains" : "s remain"} preserved. Could not refresh slide${failed.length === 1 ? "" : "s"}: ${failed.map((item) => item.slideNumber).join(", ")}. A prior verified result remains visible when it still matches the current design; otherwise the slide stays held. Build a failed slide individually to see its exact PowerPoint validation error.`);
      } else {
        const allConverted = unsupportedSourceSlideNumbers(deck, deck.studioScene).length === 0;
        if (allConverted) {
          setBusy("Assembling and validating the one central editable PowerPoint presentation…");
          try {
            const centralBuild = await buildCentralStudioDeck(deck, deck.studioScene);
            invalidateStudioQualification(deck.id);
            studioDeckBuildsRef.current = { ...studioDeckBuildsRef.current, [deck.id]: centralBuild };
            setStudioDeckBuilds(studioDeckBuildsRef.current);
            setNotice(`All ${deck.studioScene.slides.length} slide results and the central editable PowerPoint presentation are ready. Rebuilt ${built}; reused ${reused} exact current result${reused === 1 ? "" : "s"}. Slides, comments, and export now refer to this same design revision.`);
          } catch (caught) {
            setError(`Every individual slide result is ready, but the central presentation export is held: ${caught instanceof Error ? caught.message : "deck validation failed"}`);
          }
        } else {
          setNotice(`All current PowerPoint results are ready: ${built} designed slide${built === 1 ? "" : "s"} rebuilt, ${reused} exact current result${reused === 1 ? "" : "s"} reused, and ${preserved} untouched source slide${preserved === 1 ? "" : "s"} preserved. Convert the remaining source slides before exporting one central presentation.`);
        }
      }
    } finally {
      setBusy(undefined);
    }
  }

  async function qualifyStudioDeck(deck: DeckJob, build: StudioDeckBuild): Promise<StudioDeckQualification> {
    if (!desktop || !deck.audit || !deck.studioScene) throw new Error("Deck qualification requires the Electron app, a current audit, and the central Studio design.");
    if (build.sceneRevision !== deck.studioScene.revision) throw new Error("The central design changed. Build the exact current Studio revision before qualification.");
    const source = sourceForDeck(projectRef.current, deck);
    if (!source?.bytes) throw new Error("The embedded source PowerPoint is required for qualification.");
    const candidateName = `${cleanFileStem(deck.name)}_studio-redesign.pptx`;
    const runId = `${Date.now()}-${crypto.randomUUID()}`;
    const candidateSha256 = await sha256(build.bytes);
    const currentQualification = studioDeckQualificationsRef.current[deck.id];
    if (currentQualification?.sceneRevision === deck.studioScene.revision && currentQualification.candidateSha256 === candidateSha256) return currentQualification;
    const previousQualification = currentQualification ?? studioDeckQualificationHistoryRef.current[deck.id]?.at(-1);
    const capture = await desktop.captureDeckQualification({
      projectId: projectRef.current.project.id,
      deckId: deck.id,
      runId,
      width: 2200,
      source: { name: deck.name, bytes: source.bytes },
      candidate: { name: candidateName, bytes: build.bytes },
    });
    const measuredDeck: DeckJob = {
      ...deck,
      name: candidateName,
      sourceSha256: candidateSha256,
      audit: build.candidateAudit,
      scene: undefined,
      proposal: undefined,
    };
    measuredDeck.scene = compilePresentationScene({ ...measuredDeck, audit: build.candidateAudit });
    const nativePacket = bindNativeMeasurement(measuredDeck, capture.candidateMeasurement);
    const candidateMetrics = calculateDesignMetrics(measuredDeck, nativePacket);
    const protectedSourceSlideNumbers = new Set([
      ...deck.protectedSlideNumbers,
      ...deck.studioScene.slides.filter((slide) => isProtectedOrnlTemplateSlide(deck, slide.slideNumber)).map((slide) => slide.slideNumber),
    ]);
    const protectedSlideNumbers = build.outputSlides.filter((slide) => protectedSourceSlideNumbers.has(slide.sourceSlideNumber)).map((slide) => slide.outputSlideNumber);
    const designImpactBySlide = Object.fromEntries(build.outputSlides.map((output) => {
      const sourceSlide = deck.studioScene!.slides.find((slide) => slide.slideNumber === output.sourceSlideNumber);
      return [output.outputSlideNumber, sourceSlide ? analyzeStudioDesignImpact(sourceSlide).level : "unrecorded"];
    }));
    const visualNeedBySlide = Object.fromEntries(build.outputSlides.flatMap((output) => {
      const sourceSlide = deck.studioScene!.slides.find((slide) => slide.slideNumber === output.sourceSlideNumber);
      const need = (sourceSlide?.visualNeeds ?? []).find((item) => item.status !== "resolved");
      return need ? [[output.outputSlideNumber, { id: need.id, type: need.type, status: need.status }]] : [];
    }));
    const report = buildDeckQualificationReport({
      id: runId,
      sceneRevision: deck.studioScene.revision,
      sourceName: deck.name,
      candidateName,
      sourceSha256: deck.sourceSha256,
      candidateSha256,
      sourceAudit: deck.audit,
      candidateAudit: build.candidateAudit,
      sourceRender: capture.sourceRender,
      candidateRender: capture.candidateRender,
      sourceMeasurement: capture.sourceMeasurement,
      candidateMeasurement: capture.candidateMeasurement,
      candidateMeasurementPacket: nativePacket,
      candidateMetrics,
      outputSlides: build.outputSlides,
      contentValidation: build.contentValidation,
      protectedSlideNumbers,
      designImpactBySlide,
      visualNeedBySlide,
      requireMaterialDesignImpact: true,
      previousReport: previousQualification?.report,
    });
    const finalized = await desktop.finalizeDeckQualification({ outputRoot: capture.outputRoot, report });
    const qualification = { report, ...finalized, sceneRevision: deck.studioScene.revision, candidateSha256 };
    if (currentQualification) studioDeckQualificationHistoryRef.current = { ...studioDeckQualificationHistoryRef.current, [deck.id]: [...(studioDeckQualificationHistoryRef.current[deck.id] ?? []), currentQualification].slice(-5) };
    studioDeckQualificationsRef.current = { ...studioDeckQualificationsRef.current, [deck.id]: qualification };
    setStudioDeckQualifications(studioDeckQualificationsRef.current);
    return qualification;
  }

  async function runCentralStudioQualification() {
    if (!selectedDeck?.studioScene) { setError("Open a converted Studio presentation before running qualification."); return; }
    const build = studioDeckBuildsRef.current[selectedDeck.id];
    if (!build || build.sceneRevision !== selectedDeck.studioScene.revision) { setError("Run Build all results before inspecting the exact central PowerPoint candidate."); return; }
    clearMessages();
    setBusy("Building the private PowerPoint-native qualification bundle…");
    try {
      const qualification = await qualifyStudioDeck(selectedDeck, build);
      const failures = qualification.report.totals.blockerIssues + qualification.report.totals.majorIssues;
      setNotice(failures
        ? `Qualification found ${failures} objective issue${failures === 1 ? "" : "s"}. Every source and candidate slide was exported as a private full-resolution PNG; open the evidence report for repair routing.`
        : `Qualification passed objective gates for all ${qualification.report.totals.slides} slides. The full-resolution PowerPoint PNG bundle is ready for visual review.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The central presentation could not be qualified.");
    } finally {
      setBusy(undefined);
    }
  }

  async function revealCentralStudioQualification() {
    if (!selectedDeck) return;
    const qualification = studioDeckQualificationsRef.current[selectedDeck.id];
    if (!qualification) { setError("Run PowerPoint-native qualification before opening its evidence folder."); return; }
    try {
      await desktop?.revealDeckQualification({ outputRoot: qualification.outputRoot });
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The qualification evidence folder could not be opened.");
    }
  }

  async function saveFreshStudioSlide(slideNumber: number) {
    if (!selectedDeck?.studioScene || !desktop) { setError("Saving a fresh Studio PowerPoint requires the Electron desktop app."); return; }
    const preview = studioFreshPreviews[`${selectedDeck.id}:${slideNumber}`];
    const studioSlide = selectedDeck.studioScene.slides.find((item) => item.slideNumber === slideNumber);
    if (!preview || !studioSlide || preview.slideUpdatedAt !== studioSlide.updatedAt) { setError("This slide design changed. Rebuild its PowerPoint export result before saving it."); return; }
    if (preview.nativeRender?.status !== "ready" || !preview.nativeRender.authoritative) { setError("Render the fresh composition successfully in Microsoft PowerPoint before saving it."); return; }
    if (preview.nativeMeasurement?.status !== "ready" || preview.nativeMeasurement.authority !== "powerpoint-native" || nativeTextOverflows(preview.nativeMeasurement).length > 0) { setError("Remeasure the fresh composition successfully in Microsoft PowerPoint with no material text overflow before saving it."); return; }
    clearMessages();
    try {
      const result = await desktop.saveBinary({ kind: "pptx", defaultName: `${cleanFileStem(selectedDeck.name)}_slide-${slideNumber}${preview.slideCount > 1 ? "_table-continuation" : ""}_studio-rebuild.pptx`, bytes: preview.bytes });
      if (!result.canceled) setNotice(`Saved ${preview.slideCount === 1 ? "a new one-slide" : `${preview.slideCount} new merge-safe continuation slides in one`} editable PowerPoint rebuilt from Studio source slide ${slideNumber}. The imported deck remains untouched; original masters, animations, transitions, and unsupported native internals are not part of this fresh-composition file.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The fresh Studio PowerPoint could not be saved.");
    }
  }

  async function saveCentralStudioDeck() {
    if (!selectedDeck?.studioScene || !desktop) { setError("Exporting the central presentation requires the Electron desktop app."); return; }
    const build = studioDeckBuilds[selectedDeck.id];
    if (!build || build.sceneRevision !== selectedDeck.studioScene.revision) { setError("The central design changed. Run Build all results before exporting the presentation."); return; }
    clearMessages();
    setBusy("Rerendering the exact central presentation export candidate in Microsoft PowerPoint…");
    try {
      const artifactName = `${cleanFileStem(selectedDeck.name)}_studio-redesign.pptx`;
      const [render, measurement] = await Promise.all([
        desktop.renderPowerPoint({ name: artifactName, bytes: build.bytes, width: 1600, format: "png" }),
        desktop.measurePowerPoint({ name: artifactName, bytes: build.bytes }),
      ]);
      if (render.status !== "ready" || !render.authoritative) throw new Error(render.reason ?? "PowerPoint rerender failed.");
      if (render.slides.map((slide) => slide.sha256).join("|") !== build.nativeRender.slides.map((slide) => slide.sha256).join("|")) throw new Error("Export acceptance failed because the rerendered slide pixels no longer match the central design shown in the app.");
      const protectedSourceSlides = new Set(selectedDeck.studioScene.slides.filter((slide) => isProtectedOrnlTemplateSlide(selectedDeck, slide.slideNumber)).map((slide) => slide.slideNumber));
      const protectedSlides = new Set(build.outputSlides.filter((slide) => protectedSourceSlides.has(slide.sourceSlideNumber)).map((slide) => slide.outputSlideNumber));
      if (measurement.status !== "ready" || measurement.authority !== "powerpoint-native" || nativeTextOverflows(measurement).some((item) => !protectedSlides.has(item.slideNumber))) throw new Error("Export acceptance failed because PowerPoint measurement is unavailable or found text overflow outside source-preserved ORNL template slides.");
      setBusy("Export acceptance passed. Choose where to save the central presentation…");
      const result = await desktop.saveBinary({ kind: "pptx", defaultName: `${cleanFileStem(selectedDeck.name)}_studio-redesign.pptx`, bytes: build.bytes });
      if (!result.canceled) setNotice("Exported the exact central Presentation Studio design as one editable PowerPoint. The imported source file remains untouched.");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The central presentation could not be exported.");
    } finally {
      setBusy(undefined);
    }
  }

  function stageCleanup() {
    if (!selectedDeck) return;
    clearMessages();
    try {
      const proposal = createDesignerCleanupProposal(selectedDeck, project.project.updatedAt);
      setProject((current) => ({
        ...current,
        decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, proposal, status: "proposal-ready" } : deck),
        activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action: "designer-cleanup-proposal-staged", detail: `Reviewed every slide and staged supported designer cleanup for ${selectedDeck.name}; nothing was applied.` }],
      }));
      setActiveView("review");
    } catch (caught) { setError(caught instanceof Error ? caught.message : "Cleanup could not be staged."); }
  }

  function markTableExemplar() {
    if (!selectedDeck?.audit || selectedDeck.audit.tableCount !== 1) { setError("Select a deck with exactly one native table for this first exemplar workflow."); return; }
    const source = sourceForDeck(project, selectedDeck);
    const slide = selectedDeck.audit.slides.find((item) => item.tableCount === 1);
    if (!source || !slide) { setError("The table exemplar location could not be resolved."); return; }
    if (project.styleExemplars.some((item) => item.deckId === selectedDeck.id && item.kind === "table")) return;
    const exemplar = { id: crypto.randomUUID(), name: `${cleanFileStem(selectedDeck.name)} table style`, kind: "table" as const, resourceId: source.id, deckId: selectedDeck.id, slideNumber: slide.number, objectOrdinal: 1, scope: "batch" as const, createdAt: new Date().toISOString() };
    setProject((current) => touchProject({ ...current, styleExemplars: [...current.styleExemplars, exemplar], resources: current.resources.map((resource) => resource.id === source.id && !resource.roles.includes("style-exemplar") ? { ...resource, roles: [...resource.roles, "style-exemplar"] } : resource) }, "style-exemplar-registered", `Registered the sole native table in ${selectedDeck.name} as a batch style exemplar.`));
    setNotice("Approved table style exemplar registered by source hash and slide location. No table content was copied or changed.");
  }

  function saveDesignThread(slide: SlideRenderPreview, anchor: DesignThread["anchor"], comment: string, submit: boolean, outputSlideNumber?: number) {
    if (!selectedDeck?.audit) return;
    const inventory = selectedDeck.audit.slides.find((item) => item.number === slide.number);
    if (!inventory) { setError("The selected slide revision could not be resolved."); return; }
    const now = new Date().toISOString();
    const centralSlideRevision = selectedDeck.studioScene?.slides.find((item) => item.slideNumber === slide.number)?.updatedAt;
    const thread: DesignThread = { id: crypto.randomUUID(), deckId: selectedDeck.id, slideId: inventory.id, slideNumber: slide.number, outputSlideNumber, baseRevision: centralSlideRevision ?? projectRef.current.project.updatedAt, anchor, comment, status: submit ? "submitted" : "note", createdAt: now, updatedAt: now, submittedAt: submit ? now : undefined };
    setProject((current) => {
      const next = touchProject({ ...current, designThreads: [...current.designThreads, thread] }, submit ? "design-thread-submitted" : "design-note-saved", `${submit ? "Submitted" : "Saved"} a location-anchored design comment on slide ${slide.number} of ${selectedDeck.name}.`);
      return { ...next, decks: next.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal && ["pending", "applied"].includes(deck.proposal.status) ? { ...deck, proposal: { ...deck.proposal, baseUpdatedAt: next.project.updatedAt, slideReviews: submit ? [...(deck.proposal.slideReviews ?? []).filter((review) => review.slideNumber !== slide.number), { slideNumber: slide.number, decision: "changes-requested", reviewedAt: now, comment }] : deck.proposal.slideReviews } } : deck) };
    });
    const target = outputSlideNumber && outputSlideNumber !== slide.number ? `output slide ${outputSlideNumber} (source slide ${slide.number})` : `slide ${slide.number}`;
    setNotice(submit ? `Design comment submitted for AI on ${target}. It is anchored to the selected region and no change was applied.` : `Private design note saved on ${target}. It is not available through MCP.`);
  }

  function deleteDesignThread(threadId: string) {
    const thread = projectRef.current.designThreads.find((item) => item.id === threadId);
    if (!thread) return;
    const deckName = projectRef.current.decks.find((deck) => deck.id === thread.deckId)?.name ?? "the deck";
    setProject((current) => {
      const designThreads = removeDesignThread(current.designThreads, threadId);
      const hasOtherSlideComment = designThreads.some((item) => item.deckId === thread.deckId && item.slideNumber === thread.slideNumber);
      const updated = touchProject({
        ...current,
        designThreads,
        decks: current.decks.map((deck) => deck.id === thread.deckId && deck.proposal && !hasOtherSlideComment ? {
          ...deck,
          proposal: { ...deck.proposal, slideReviews: (deck.proposal.slideReviews ?? []).filter((review) => !(review.slideNumber === thread.slideNumber && review.decision === "changes-requested")) },
        } : deck),
      }, "design-thread-deleted", `Deleted a design comment from slide ${thread.slideNumber} of ${deckName}.`);
      return {
        ...updated,
        decks: updated.decks.map((deck) => deck.id === thread.deckId && deck.proposal ? { ...deck, proposal: { ...deck.proposal, baseUpdatedAt: updated.project.updatedAt } } : deck),
      };
    });
    setNotice(`Comment deleted from slide ${thread.slideNumber}. The clean slide is ready for new feedback.`);
  }

  function reviewSlide(slideNumber: number, decision: "approved" | "changes-requested") {
    if (!selectedDeck?.proposal || selectedDeck.proposal.status !== "pending") return;
    const now = new Date().toISOString();
    updateProposal((proposal) => ({ ...proposal, slideReviews: [...(proposal.slideReviews ?? []).filter((review) => review.slideNumber !== slideNumber), { slideNumber, decision, reviewedAt: now }] }), decision === "approved" ? "slide-review-approved" : "slide-review-changes-requested");
    setNotice(decision === "approved" ? `Slide ${slideNumber} approved. The source remains unchanged until export.` : `Slide ${slideNumber} marked for revision.`);
  }

  function requestSlideChanges(slideNumber: number, comment: string, submit: boolean) {
    if (!selectedDeck?.audit || !selectedDeck.proposal || selectedDeck.proposal.status !== "pending") return;
    const inventory = selectedDeck.audit.slides.find((slide) => slide.number === slideNumber);
    if (!inventory) { setError("The selected slide revision could not be resolved."); return; }
    const now = new Date().toISOString();
    const thread: DesignThread = { id: crypto.randomUUID(), deckId: selectedDeck.id, slideId: inventory.id, slideNumber, baseRevision: projectRef.current.project.updatedAt, anchor: { kind: "region", x: 0, y: 0, width: 1, height: 1 }, comment, status: submit ? "submitted" : "note", createdAt: now, updatedAt: now, submittedAt: submit ? now : undefined };
    setProject((current) => ({
      ...current,
      designThreads: [...current.designThreads, thread],
      decks: current.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal ? { ...deck, proposal: { ...deck.proposal, slideReviews: [...(deck.proposal.slideReviews ?? []).filter((review) => review.slideNumber !== slideNumber), { slideNumber, decision: "changes-requested", reviewedAt: now, comment }] } } : deck),
      activity: [...current.activity, { id: crypto.randomUUID(), at: now, action: submit ? "review-change-request-submitted" : "review-change-note-saved", detail: `${submit ? "Submitted" : "Saved"} a whole-slide revision request on slide ${slideNumber} of ${selectedDeck.name}.` }],
    }));
    setNotice(submit ? `Slide ${slideNumber} sent back to AI with your comment. Other slide approvals are preserved.` : `Revision note saved on slide ${slideNumber}.`);
  }

  function openSlideWorkspace(slideNumber: number, mode: SlideWorkspaceRequest["mode"]) {
    if (!selectedDeck) return;
    setSlideWorkspaceRequest({ id: crypto.randomUUID(), deckId: selectedDeck.id, slideNumber, mode, representation: "proposal" });
    setActiveView("slides");
  }

  function stageGeometryEdit(object: SlideEditableObject, target: { x: number; y: number; width: number; height: number }, rationale: string) {
    if (!selectedDeck?.audit) return;
    clearMessages();
    try {
      const proposal = createGeometryEditProposal(selectedDeck, project.project.updatedAt, { objectId: object.id, target, rationale, author: "human" });
      setProject((current) => ({
        ...current,
        decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, proposal, status: "proposal-ready" } : deck),
        activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action: "slide-geometry-proposal-staged", detail: `Staged a measured ${object.kind} geometry edit on slide ${object.slideNumber} of ${selectedDeck.name}; source bytes remain unchanged.` }],
      }));
      setActiveView("review");
      setNotice(`Object edit staged on slide ${object.slideNumber}. Compare Current and Proposal before accepting; no source content was changed.`);
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The object edit could not be staged."); }
  }

  function updateProposal(update: (proposal: CleanupProposal) => CleanupProposal, action: string) {
    if (!selectedDeck?.proposal) return;
    setProject((current) => ({
      ...current,
      decks: current.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal ? { ...deck, proposal: update(deck.proposal) } : deck),
      activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action, detail: `${action} for ${selectedDeck.name}.` }],
    }));
  }

  function toggleChange(id: string) { updateProposal((proposal) => ({ ...proposal, changes: proposal.changes.map((change) => change.id === id ? { ...change, selected: !change.selected } : change) }), "proposal-selection-updated"); }
  function rejectProposal() { updateProposal((proposal) => ({ ...proposal, status: "rejected", designReview: { decision: "rejected", actor: "human", rationale: "Rejected in Presentation Studio review.", reviewedAt: new Date().toISOString() } }), "cleanup-proposal-rejected"); setProject((current) => ({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck?.id ? { ...deck, status: "audited" } : deck) })); }
  function acceptProposal(slideNumber = 1) {
    if (!selectedDeck?.proposal || selectedDeck.proposal.baseUpdatedAt !== project.project.updatedAt) { setError("The proposal is stale. Restage it against the current project."); return; }
    if (selectedDeck.proposal.slideReviews?.some((review) => review.decision === "changes-requested")) { setError("Resolve the requested slide changes before approving the full plan."); return; }
    const now = new Date().toISOString();
    setProject((current) => {
      const next = touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal ? { ...deck, proposal: { ...deck.proposal, status: "applied", slideReviews: (deck.proposal.slideDispositions ?? []).map((item) => ({ slideNumber: item.slideNumber, decision: "approved" as const, reviewedAt: now })) }, status: "approved" } : deck) }, "cleanup-plan-accepted", `Accepted the cleanup plan for ${selectedDeck.name}; source bytes remain unchanged until a new copy is exported.`);
      return { ...next, decks: next.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal ? { ...deck, proposal: { ...deck.proposal, baseUpdatedAt: next.project.updatedAt } } : deck) };
    });
    setSlideWorkspaceRequest({ id: crypto.randomUUID(), deckId: selectedDeck.id, slideNumber, mode: "review", representation: "proposal" });
    setActiveView("slides");
    setNotice("All selected changes approved. You can keep editing or comment for AI; export remains a separate action.");
  }

  async function exportCleaned() {
    if (!selectedDeck?.proposal || selectedDeck.proposal.status !== "applied") return;
    const resource = sourceForDeck(project, selectedDeck);
    if (!resource?.bytes) { setError("The embedded source deck is unavailable."); return; }
    setBusy("Building and validating a new cleaned PowerPoint copy…");
    clearMessages();
    try {
      const output = await applyCleanupToPptx(resource.bytes, selectedDeck.proposal, { templateBytes: templateSourceBytes });
      if (!desktop) throw new Error("PowerPoint export requires the Electron desktop app.");
      setBusy("Rerendering and remeasuring the actual export candidate in PowerPoint…");
      const [stagedRender, stagedMeasurement, exportRender, exportNativeMeasurement] = await Promise.all([
        getOrBuildNativeRender(selectedDeck, "proposal", project),
        getOrBuildNativeMeasurement(selectedDeck, "proposal", project),
        desktop.renderPowerPoint({ name: `${cleanFileStem(selectedDeck.name)}_export-acceptance.pptx`, bytes: output.bytes }),
        desktop.measurePowerPoint({ name: `${cleanFileStem(selectedDeck.name)}_export-acceptance.pptx`, bytes: output.bytes }),
      ]);
      if (!stagedRender || stagedRender.status !== "ready" || exportRender.status !== "ready") throw new Error("Export acceptance failed because both staged and exported-file PowerPoint renders were not authoritative.");
      const stagedHashes = stagedRender.slides.map((slide) => `${slide.number}:${slide.sha256}`);
      const exportHashes = exportRender.slides.map((slide) => `${slide.number}:${slide.sha256}`);
      if (stagedHashes.length !== exportHashes.length || stagedHashes.some((value, index) => value !== exportHashes[index])) throw new Error("Export acceptance failed because the independently rerendered PPTX does not match the staged PowerPoint proposal pixels.");
      const exportMeasurement = bindNativeMeasurement(selectedDeck, exportNativeMeasurement);
      const measurementComparison = compareNativeMeasurementPackets(stagedMeasurement, exportMeasurement);
      if (!measurementComparison.equivalent) throw new Error(`Export acceptance failed because native PowerPoint geometry changed in the written artifact candidate: ${measurementComparison.mismatches.slice(0, 8).join(", ")}.`);
      setBusy("Export acceptance passed. Choose where to save the verified review copy…");
      const result = await desktop.saveBinary({ kind: "pptx", defaultName: `${cleanFileStem(selectedDeck.name)}_designer-cleaned.pptx`, bytes: output.bytes });
      if (!result.canceled) {
        setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, status: "needs-manual-review", exportedAt: new Date().toISOString() } : deck) }, "cleaned-review-copy-exported", `Exported a new review copy of ${selectedDeck.name} after independent PowerPoint rerender and remeasurement acceptance; ${output.replacementCount} font references, ${output.alignmentCount} alignments, ${output.geometryCount} object edits, ${output.textStyleCount} text styles, ${output.decorationCount} brand vectors, ${output.layoutCount} native layout remaps, ${output.tableCount} native table styles, and ${output.tableLayoutCount} solved table grids were updated.`));
        setNotice(`Verified review copy exported. Exact visible text, table content, and merged-cell structure passed validation; the independently rerendered PPTX matched every staged slide raster and native measurement within ${measurementComparison.tolerancePt} pt. Complete the final human design review in PowerPoint.`);
      }
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The cleaned copy could not be exported."); }
    finally { setBusy(undefined); }
  }

  async function exportAuditReport() {
    if (!selectedDeck?.audit) return;
    if (!desktop) { setError("Audit report export requires the Electron desktop app."); return; }
    clearMessages();
    try {
      const bytes = buildAuditReport(project, selectedDeck);
      const result = await desktop.saveBinary({ kind: "report", defaultName: `${cleanFileStem(selectedDeck.name)}_audit-report.json`, bytes });
      if (!result.canceled) setNotice("Audit report exported without slide text, notes, picture names, descriptions, or Resource bytes.");
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The audit report could not be exported."); }
  }

  async function saveProject(encrypted: boolean) {
    if (!desktop) { setError("Project save requires the Electron desktop app."); return; }
    clearMessages(); setBusy(encrypted ? "Encrypting the self-contained project…" : "Packaging the self-contained project…");
    try {
      let bytes = await buildProjectPackage(project);
      let password: string | undefined;
      if (encrypted) {
        password = window.prompt("Create a project password (at least 12 characters). The password cannot be recovered.") ?? undefined;
        if (!password) return;
        const confirmation = window.prompt("Enter the same project password again.");
        if (confirmation !== password) throw new Error("The two project passwords did not match.");
        bytes = await encryptProjectPackage(bytes, password);
      }
      const kind = encrypted ? "secure-project" : "project";
      const result = await desktop.saveBinary({
        kind,
        defaultName: projectSaveDefaultName({
          projectName: project.project.name,
          deckNames: project.decks.map((deck) => deck.name),
          encrypted,
        }),
        bytes,
      });
      if (!result.canceled) {
        setSecureAutosavePassword(password);
        setNotice(encrypted ? "Encrypted project saved. External originals and exported files are not covered by this encryption." : "Self-contained project saved.");
      }
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The project could not be saved."); }
    finally { setBusy(undefined); }
  }

  async function openProject() {
    if (!desktop) return;
    const result = await desktop.openProject();
    if (result.canceled || !result.file) return;
    await openProjectFile(result.file);
  }

  const currentStudioDeckBuild = selectedDeck?.studioScene && studioDeckBuilds[selectedDeck.id]?.sceneRevision === selectedDeck.studioScene.revision
    ? studioDeckBuilds[selectedDeck.id]
    : undefined;
  const currentStudioQualification = selectedDeck?.studioScene && studioDeckQualifications[selectedDeck.id]?.sceneRevision === selectedDeck.studioScene.revision
    ? studioDeckQualifications[selectedDeck.id]
    : undefined;
  const latestStudioNativeRender = currentStudioDeckBuild?.nativeRender ?? (selectedDeck?.studioScene
    ? composeLatestStudioNativeRender(
      selectedDeck.studioScene,
      nativeRenderCatalogs[`${selectedDeck.id}:current`],
      studioFreshPreviews,
    )
    : undefined);

  const mainContent = useMemo(() => {
    if (activeView === "batch") return <BatchView project={project} selectedId={selectedDeck?.id} onSelect={(id) => { setSelectedDeckId(id); setActiveView("decks"); }} onAdd={() => void addDecks()} />;
    if (activeView === "decks") return <DeckAuditView deck={selectedDeck} onConfirm={confirmTemplate} onStage={stageCleanup} onStartOrnlCleanup={startOrnlCleanup} onMarkExemplar={markTableExemplar} onExportReport={() => void exportAuditReport()} isExemplar={Boolean(selectedDeck && project.styleExemplars.some((item) => item.deckId === selectedDeck.id && item.kind === "table"))} />;
    if (activeView === "slides") { const proposalWorkspace = selectedDeck?.proposal?.status === "applied" || isProposalSlideWorkspaceRequest(slideWorkspaceRequest, selectedDeck?.id); return <SlidesView deck={selectedDeck} catalog={selectedDeck ? proposalWorkspace ? proposalCatalogs[selectedDeck.id] : slideCatalogs[selectedDeck.id] : undefined} nativeRender={selectedDeck ? proposalWorkspace ? nativeRenderCatalogs[`${selectedDeck.id}:proposal`] : selectedDeck.studioScene ? latestStudioNativeRender : nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} outputSlides={!proposalWorkspace ? currentStudioDeckBuild?.outputSlides : undefined} loading={Boolean(selectedDeck && (proposalWorkspace ? proposalCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:proposal` : slideCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:current`))} revision={project.project.updatedAt} threads={project.designThreads} openRequest={slideWorkspaceRequest} deckBuildReady={Boolean(currentStudioDeckBuild)} resourceCount={project.resources.length} onAddDeck={() => void addDecks()} onOpenResources={() => setActiveView("resources")} onSaveThread={saveDesignThread} onDeleteThread={deleteDesignThread} onStageGeometry={stageGeometryEdit} onOpenStudioSlide={(slideNumber) => { setStudioOpenSlideNumber(slideNumber); setActiveView("studio"); }} onSaveDeck={() => void saveCentralStudioDeck()} />; }
    if (activeView === "studio") return <StudioView deck={selectedDeck} catalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} nativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} freshPreviews={studioFreshPreviews} templateCatalog={templateCatalog} templateNativeRender={templateNativeRender} resources={project.resources} requestedSlideNumber={studioOpenSlideNumber} deckBuildReady={Boolean(currentStudioDeckBuild)} qualification={currentStudioQualification} canUndo={canUndoStudio} canRedo={canRedoStudio} onUndo={() => restoreStudioHistory("undo")} onRedo={() => restoreStudioHistory("redo")} onInitialize={() => void initializeStudioScene()} onRecompose={recomposeStudioSlide} onMoveNode={moveStudioNode} onMoveNodes={moveStudioNodes} onStyleNode={styleStudioNode} onPublishComponent={publishStudioComponentStyle} onArrangeSelection={arrangeStudioSelection} onRepairObjectiveIssues={repairStudioObjectiveIssuesFromUi} onUpdateTableDesign={styleStudioTable} onResizeTableColumn={resizeStudioTableColumnInches} onResizeTableRow={resizeStudioTableRowInches} onUpdateTableCell={styleStudioTableCell} onPublishTableExemplar={publishStudioTableExemplarStyle} onApplyTableExemplar={applyStudioTableExemplarStyle} onPlanTableContinuation={planStudioTableContinuationFromUi} onClearTableContinuation={clearStudioTableContinuationFromUi} onUpdateConnector={authorStudioConnector} onUpdateFigure={setStudioFigureTreatment} onCreateVisualNeed={createVisualNeedFromStudio} onHoldVisualNeed={holdVisualNeedFromStudio} onAttachConcept={attachConceptFromStudio} onDetachConcept={detachConceptFromStudio} onReconstructConcept={reconstructConceptFromStudio} onPreviewFresh={(slideNumber) => void previewFreshStudioSlide(slideNumber)} onPreviewAll={() => void previewAllFreshStudioSlides()} onQualify={() => void runCentralStudioQualification()} onRevealQualification={() => void revealCentralStudioQualification()} onSaveFresh={(slideNumber) => void saveFreshStudioSlide(slideNumber)} onSaveDeck={() => void saveCentralStudioDeck()} />;
    if (activeView === "designs") return <DesignsView catalog={templateCatalog} installedAt={templateInstalledAt} loading={templateLoading} nativeRender={templateNativeRender} nativeLoading={templateNativeLoading} onInstall={() => void installTemplate()} />;
    if (activeView === "rules") return <RulesView deck={selectedDeck} exemplarCount={project.styleExemplars.filter((item) => item.kind === "table").length} />;
    if (activeView === "review") return <ReviewView deck={selectedDeck} projectUpdatedAt={project.project.updatedAt} currentCatalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} proposalCatalog={selectedDeck ? proposalCatalogs[selectedDeck.id] : undefined} currentNativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} proposalNativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:proposal`] : undefined} previewLoading={Boolean(selectedDeck && (proposalCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:proposal`))} threads={project.designThreads} onToggle={toggleChange} onReviewSlide={reviewSlide} onRequestChanges={requestSlideChanges} onDeleteThread={deleteDesignThread} onOpenSlide={openSlideWorkspace} onReject={rejectProposal} onApply={acceptProposal} onExport={() => void exportCleaned()} />;
    return <ResourcesView project={project} aiSessionEnabled={mcpEnabled} onAdd={() => void addResources()} onRemove={removeResource} />;
  }, [activeView, currentStudioDeckBuild, currentStudioQualification, latestStudioNativeRender, mcpEnabled, nativeRenderCatalogs, nativeRenderLoadingKey, project, proposalCatalogLoadingDeckId, proposalCatalogs, selectedDeck, slideCatalogLoadingDeckId, slideCatalogs, slideWorkspaceRequest, studioFreshPreviews, studioOpenSlideNumber, templateCatalog, templateInstalledAt, templateLoading, templateNativeLoading, templateNativeRender]);
  const mcpActivityMessage = mcpActivity ? mcpActivityCopy(mcpActivity) : undefined;

  return (
    <div className="app-shell" onDragEnter={handleDragEnter} onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}>
      <header className="topbar">
        <div className="brand" data-tour="brand"><span className="brand-mark">PS</span><div><strong>Presentation Studio</strong><small>Audit · Clean · Review</small></div></div>
        <div className="project-title"><span>{project.project.name}</span><small>{project.decks.length} decks · {project.settings.contentPolicy}</small></div>
        <div className="top-actions" data-tour="save">
          <button className="button ghost small" data-tour="tour" aria-haspopup="dialog" onClick={openOnboardingTour}><Info size={17} />Tour</button>
          <button className="button ghost small" onClick={() => void openProject()}><FolderOpen size={17} />Open</button>
          <button className="button ghost small" onClick={() => void saveProject(false)}><FileArrowDown size={17} />Save</button>
          <button className="button secondary small" onClick={() => void saveProject(true)}><FileLock size={17} />Save encrypted</button>
        </div>
      </header>
      <nav className="rail" aria-label="Workspace">
        <div className="rail-items">{navItems.map((item) => { const NavIcon = item.icon; return <button key={item.id} data-tour={`nav-${item.id}`} className={activeView === item.id ? "active" : ""} onClick={() => setActiveView(item.id)}><NavIcon size={20} /><span>{item.label}</span>{item.id === "review" && project.decks.some((deck) => deck.proposal?.status === "pending") && <i />}</button>; })}</div>
        <div className="rail-bottom">
          <button className={`ai-session ${mcpEnabled ? "enabled" : ""}`} data-tour="ai-session" onClick={() => updateAiSessionAccess(!mcpEnabled)}><span className="ai-icon"><Sparkle size={18} /></span><span><strong>AI access</strong><small>{mcpEnabled ? `All ${project.resources.length} Resources shared` : "Access off"}</small></span><span className="toggle-knob" /></button>
          <div className="local-status"><span className={mcpStatus.available ? "online" : ""} /><span>{mcpStatus.available ? "Local MCP ready" : "Browser preview"}</span></div>
          {desktop && <div className="local-status native-qa" title={nativeReadiness.reason}><span className={nativeReadiness.ready ? "online" : nativeReadiness.sessionLocked ? "locked" : ""} /><span>{nativeReadiness.ready ? "PowerPoint QA ready" : nativeReadiness.sessionLocked ? "Unlock Mac for native QA" : "Native QA unavailable"}</span></div>}
        </div>
      </nav>
      <main className="workspace">{mainContent}</main>
      <Inspector deck={selectedDeck} onOpenReview={() => setActiveView("review")} />
      <OnboardingTour open={tourOpen} stepIndex={tourStepIndex} onStepChange={setTourStepIndex} onClose={closeOnboardingTour} />
      {mcpActivity && mcpActivityMessage && <div className={`mcp-activity ${mcpActivity.state} phase-${mcpActivity.phase}`} role="status" aria-live="polite"><span>{mcpActivity.state === "active" ? <ArrowsClockwise className="spinner" size={17} /> : mcpActivity.state === "completed" ? <CheckCircle size={17} /> : <Warning size={17} />}</span><div><strong>{mcpActivityMessage.title}</strong><small>{mcpActivityMessage.detail}</small></div></div>}
      {(notice || error) && <div className={`toast ${error ? "error" : "success"}`}><span>{error ? <Warning size={18} /> : <CheckCircle size={18} />}</span><p>{error ?? notice}</p><button onClick={clearMessages}><X size={16} /></button></div>}
      {fileDragActive && <div className="file-drop-overlay"><div className="file-drop-card"><UploadSimple size={38} /><strong>Drop to open a project or add Resources</strong><span>A single .pstudio package opens as the active project. Other files are processed locally, embedded by hash, and preserved with the current project.</span></div></div>}
      {busy && <div className="busy-overlay"><div className="busy-card"><ArrowsClockwise className="spinner" size={25} /><strong>{busy}</strong><span>Files are processed locally and copied into the project. External originals remain untouched.</span></div></div>}
      <input id="web-deck-picker" hidden type="file" accept=".pptx" multiple onChange={(event) => { void importWebFiles(event.target.files, "decks"); event.currentTarget.value = ""; }} />
      <input id="web-resource-picker" hidden type="file" accept=".pptx,.potx,.docx,.pdf,.md,.markdown,.txt,.csv,.tsv,.json,.xlsx,.png,.jpg,.jpeg,.webp,.tif,.tiff,.svg,.wav,.mp3,.m4a,.mp4,.mov,.doc,.xls" multiple onChange={(event) => { void importWebFiles(event.target.files, "resources"); event.currentTarget.value = ""; }} />
      <input id="web-template-picker" hidden type="file" accept=".pptx,.potx" onChange={(event) => { const file = event.target.files?.[0]; if (file) void file.arrayBuffer().then((buffer) => loadTemplate({ name: file.name, filePath: file.name, mediaType: file.type, bytes: new Uint8Array(buffer) }, false)); event.currentTarget.value = ""; }} />
    </div>
  );
}
