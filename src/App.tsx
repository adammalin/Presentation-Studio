import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent, type PointerEvent as ReactPointerEvent } from "react";
import {
  Archive,
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
  StudioLayoutRecipe,
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
import { removeAddressedDesignThreads, removeCompletedDesignThreads, removeDesignThread } from "./lib/design-threads";
import { compilePresentationScene, sceneNeedsRebuild } from "./lib/scene-graph";
import { semanticRecompositionRequests, type SemanticSlotBinding } from "./lib/recomposition";
import { compareNativeSlideRenders, type PixelComparisonMetrics } from "./lib/render-comparison";
import { buildProjectPackage, openProjectPackage } from "./lib/project-package";
import { projectPackageFromDrop } from "./lib/project-drop";
import { removeResourceFromProject, resourceRemovalImpact } from "./lib/resource-removal";
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
import { bindNativeMeasurement, compareNativeMeasurementPackets, type NativeMeasurementPacket } from "./lib/native-measurement";
import { calculateDesignMetrics, metricsImproved } from "./lib/design-metrics";
import { buildInspectionPacket, type InspectionCropRegion } from "./lib/inspection-packet";
import { renderNativeContactSheet } from "./lib/contact-sheet";
import { solveAlignment, solveDistribution, solveGroupLayout, solveSafeRegion, solveSceneToLayout, type AlignmentMode, type DistributionMode, type GroupHierarchyRole, type GroupLayoutAlignment, type GroupLayoutMode, type SceneLayoutRegionRequest } from "./lib/layout-solver";
import { recommendedTableGrowthPlan, solveTableLayout } from "./lib/table-layout-solver";
import { nativeTextFrameOverflows, solveTextFit } from "./lib/text-fit-solver";
import { decideVisualIteration } from "./lib/visual-iteration";
import { sha256 } from "./lib/hash";
import { cleanFileStem, projectSaveDefaultName } from "./lib/file-names";
import { compileStudioWebScene, recommendedStudioRecipe, recomposeStudioWebSlide, studioGeometryRequests, studioVisualDesignRequest, updateStudioWebNodeFrame, updateStudioWebNodeStyle } from "./lib/studio-web-scene";
import {
  ONBOARDING_TOUR_STORAGE_KEY,
  ONBOARDING_TOUR_VERSION,
  shouldShowOnboardingTour,
} from "./lib/onboarding";

type ViewId = "batch" | "decks" | "slides" | "studio" | "designs" | "rules" | "review" | "resources";
type SlideWorkspaceRequest = { id: string; deckId: string; slideNumber: number; mode: "review" | "edit" | "comment"; representation: "current" | "proposal" };
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
      <p className="eyebrow">Conservative cleanup first</p>
      <h1>Bring order to a large presentation batch.</h1>
      <p className="empty-copy">Add PowerPoint decks for a read-only structural audit. Presentation Studio copies each source into the project, identifies template and font inconsistencies, and waits for human review before changing anything.</p>
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

function SlidesView({ deck, catalog, nativeRender, loading, revision, threads, openRequest, onSaveThread, onDeleteThread, onStageGeometry }: { deck?: DeckJob; catalog?: SlideRenderCatalog; nativeRender?: NativeRenderResult; loading: boolean; revision: string; threads: DesignThread[]; openRequest?: SlideWorkspaceRequest; onSaveThread: (slide: SlideRenderPreview, anchor: DesignThread["anchor"], comment: string, submit: boolean) => void; onDeleteThread: (threadId: string) => void; onStageGeometry: (object: SlideEditableObject, target: { x: number; y: number; width: number; height: number }, rationale: string) => void }) {
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
    setSelectedNumber(openRequest.slideNumber);
    setMode(openRequest.mode);
    setDraftAnchor(undefined);
    setDraftComment("");
  }, [deck?.id, openRequest?.id]);
  useEffect(() => {
    if (mode !== "comment") return;
    const frame = window.requestAnimationFrame(() => commentInput.current?.focus());
    return () => window.cancelAnimationFrame(frame);
  }, [mode, selectedNumber, threads.length]);
  useEffect(() => { setSelectedObjectId(undefined); setDraftGeometry(undefined); setEditRationale(""); objectDrag.current = undefined; }, [selectedNumber]);
  if (!deck?.audit) return <NoSelection message="Select a deck to inspect its current slide designs." />;

  const proposalWorkspace = openRequest?.deckId === deck.id && openRequest.representation === "proposal";
  const selected = catalog?.slides.find((slide) => slide.number === selectedNumber);
  const selectedThreads = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === selectedNumber);
  const acceptedGeometry = new Map((proposalWorkspace || deck.proposal?.status === "applied") && deck.proposal ? deck.proposal.changes.filter((change) => change.selected && change.kind === "geometry").flatMap((change) => change.geometryCommands ?? []).map((command) => [command.objectId, command.target]) : []);
  const slideObjects = (deck.audit.editableObjects ?? []).filter((object) => object.slideNumber === selectedNumber).map((object) => acceptedGeometry.has(object.id) ? { ...object, geometry: { ...object.geometry, ...acceptedGeometry.get(object.id)! } } : object);
  const slideScene = deck.scene?.slides.find((slide) => slide.number === selectedNumber);
  const sceneObjects = (deck.scene?.objects ?? []).filter((object) => object.slideNumber === selectedNumber);
  const sceneObjectById = new Map(sceneObjects.map((object) => [object.id, object]));
  const selectedObject = slideObjects.find((object) => object.id === selectedObjectId);
  const slideSize = deck.audit.slideSize ?? { width: catalog?.slideWidth ?? 12_192_000, height: catalog?.slideHeight ?? 6_858_000 };
  const commentMode = mode === "comment";
  const editMode = mode === "edit";
  const nativeReady = nativeRender?.status === "ready" && nativeRender.authoritative;
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
    onSaveThread(selected, draftAnchor, draftComment.trim(), submit);
    setDraftAnchor(undefined);
    setDraftComment("");
    setMode("comment");
  }

  return <div className="view-stack slides-view">
    <header className="view-header compact"><div><p className="eyebrow">{proposalWorkspace ? "Proposed slide designs" : "Current slide designs"}</p><h1>{deck.name}</h1><p>{nativeReady ? `Faithful, revision-bound ${proposalWorkspace ? "proposal" : "current"} pixels rendered locally by Microsoft PowerPoint.` : "PowerPoint-native rendering is unavailable; the visible slide is a labeled structural approximation."}</p></div><span className="render-status"><span className={loading ? "loading" : nativeReady ? "ready" : catalog ? "fallback" : ""} />{loading ? "Rendering in PowerPoint…" : nativeReady ? `${nativeRender.slideCount} PowerPoint-native previews` : catalog ? `${catalog.slides.length} approximate previews` : "Preview unavailable"}</span></header>
    {selected && catalog && <section className="slide-review panel">
      <div className="slide-review-toolbar"><div><button className="button ghost small" onClick={() => setSelectedNumber(undefined)}><ArrowLeft size={15} />Gallery</button><span><strong>Slide {selected.number}</strong><small>{proposalWorkspace ? "Proposal" : "Current"} · revision {revision.slice(0, 19).replace("T", " ")}</small></span></div><div><button className="button ghost small" disabled={selected.number <= 1} onClick={() => setSelectedNumber(selected.number - 1)}><ArrowLeft size={15} />Previous</button><button className="button ghost small" disabled={selected.number >= catalog.slides.length} onClick={() => setSelectedNumber(selected.number + 1)}>Next<ArrowRight size={15} /></button><button className={`button small ${editMode ? "primary" : "secondary"}`} onClick={() => { setMode(editMode ? "review" : "edit"); setDraftAnchor(undefined); }}><Crosshair size={16} />{editMode ? "Editing objects" : "Edit objects"}</button><button className={`button small ${commentMode ? "primary" : "secondary"}`} onClick={() => { setMode(commentMode ? "review" : "comment"); setDraftAnchor(undefined); }}><ChatCircleDots size={16} />{commentMode ? "Select a region" : "Comment"}</button></div></div>
      <div className="slide-review-body">
        <div className="slide-review-stage">
          <div className="slide-review-canvas" ref={editorCanvas}><SlideDesignCanvas nativeRender={nativeRender} slideNumber={selected.number} catalog={catalog} layout={selected} label={`${proposalWorkspace ? "Proposed" : "Current"} design for slide ${selected.number}: ${selected.title}`} /><div className={`slide-anchor-layer ${commentMode ? "active" : ""}`} aria-label={commentMode ? "Drag over the exact slide region to comment" : "Slide comment anchors"} onPointerDown={beginAnchor} onPointerMove={moveAnchor} onPointerUp={finishAnchor}>
            {selectedThreads.map((thread, index) => <span key={thread.id} className={`thread-anchor ${thread.status}`} style={{ left: `${thread.anchor.x * 100}%`, top: `${thread.anchor.y * 100}%`, width: `${thread.anchor.width * 100}%`, height: `${thread.anchor.height * 100}%` }} title={thread.comment}><i>{index + 1}</i></span>)}
            {draftAnchor && <span className="thread-anchor draft" style={{ left: `${draftAnchor.x * 100}%`, top: `${draftAnchor.y * 100}%`, width: `${draftAnchor.width * 100}%`, height: `${draftAnchor.height * 100}%` }} />}
          </div>{editMode && <div className="slide-object-layer" onPointerMove={moveObject} onPointerUp={finishObjectDrag} onPointerCancel={finishObjectDrag}>{slideObjects.map((object) => { const geometry = activeGeometry(object); const sceneObject = sceneObjectById.get(object.id); const movable = sceneObject?.operations.move ?? object.canMove; return <button key={object.id} className={`slide-object-box ${selectedObjectId === object.id ? "selected" : ""} ${!movable ? "locked" : ""} ${sceneObject?.fidelityState ?? ""}`} style={{ left: `${geometry.x / slideSize.width * 100}%`, top: `${geometry.y / slideSize.height * 100}%`, width: `${geometry.width / slideSize.width * 100}%`, height: `${geometry.height / slideSize.height * 100}%`, zIndex: Math.min(100, (sceneObject?.zIndex ?? 0) + 1) }} onPointerDown={(event) => beginObjectDrag(event, object)} onClick={(event) => { event.stopPropagation(); selectObject(object); }} title={`${object.name} · ${object.kind} · ${sceneObject ? fidelityLabels[sceneObject.fidelityState] : "legacy object"}`}><span>{sceneObject?.semanticRole ?? object.kind} · {sceneObject ? fidelityLabels[sceneObject.fidelityState] : "legacy"}</span></button>; })}</div>}</div>
          <div className={`slide-representation-note ${nativeReady ? "native" : "fallback"}`}><ShieldCheck size={15} /><span><strong>{editMode ? "Non-destructive object editor" : nativeReady ? "PowerPoint-native representation" : "Approximate OOXML fallback"}</strong> {editMode ? "Drag a highlighted object, use measured controls, then stage it into Current/Proposal review. Imported slide bytes stay untouched." : nativeReady ? "These pixels came from Microsoft PowerPoint. Structured editable-object overlays remain bound to the same source revision." : nativeRender?.warnings[0] ?? "This diagnostic reconstruction is not authoritative for wrapping, masters, tables, equations, or crop behavior."}</span></div>
        </div>
        <aside className="slide-thread-panel">{editMode ? <><div className="thread-panel-heading"><Crosshair size={19} /><div><strong>Object editor</strong><small>{slideObjects.length} source-bound objects{slideScene?.preservationRequired ? " · preservation active" : ""}</small></div></div>{!selectedObject || !draftGeometry ? <div className="thread-empty"><Crosshair size={24} /><strong>Select an object</strong><span>Every source-bound slide object is outlined with its fidelity state.</span></div> : (() => { const sceneObject = sceneObjectById.get(selectedObject.id); return <div className="object-editor-panel"><div className="object-editor-identity"><span>{sceneObject?.semanticRole ?? selectedObject.kind}</span><strong>{selectedObject.name}</strong><small>ID {selectedObject.id}</small>{sceneObject && <em className={`fidelity-badge ${sceneObject.fidelityState}`}>{fidelityLabels[sceneObject.fidelityState]}</em>}</div>{sceneObject && <p className="fidelity-reason">{sceneObject.fidelityReason}</p>}<div className="geometry-grid">{(["x", "y", "width", "height"] as const).map((field) => <label key={field}><span>{field === "width" ? "W" : field === "height" ? "H" : field.toUpperCase()} (in)</span><input type="number" min="0" step="0.05" value={(draftGeometry[field] / 914_400).toFixed(2)} disabled={(field === "width" || field === "height") ? !(sceneObject?.operations.resize ?? selectedObject.canResize) : !(sceneObject?.operations.move ?? selectedObject.canMove)} onChange={(event) => setInches(field, event.target.value)} /></label>)}</div><div className="nudge-grid"><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: value.y - 45_720 }))}>↑</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: value.x - 45_720 }))}>←</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: value.x + 45_720 }))}>→</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: value.y + 45_720 }))}>↓</button></div><span className="field-label">Align to 0.5-inch safe area</span><div className="align-grid"><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: 457_200 }))}>Left</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: Math.round((slideSize.width - value.width) / 2) }))}>Center</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, x: slideSize.width - 457_200 - value.width }))}>Right</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: 457_200 }))}>Top</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: Math.round((slideSize.height - value.height) / 2) }))}>Middle</button><button disabled={!(sceneObject?.operations.move ?? selectedObject.canMove)} onClick={() => updateDraft((value) => ({ ...value, y: slideSize.height - 457_200 - value.height }))}>Bottom</button></div><label className="edit-rationale"><span className="field-label">Design intent</span><textarea value={editRationale} maxLength={700} onChange={(event) => setEditRationale(event.target.value)} placeholder="Example: Align the caption to the figure edge and establish a consistent lower margin." /></label><button className="button primary object-stage-button" disabled={!(sceneObject?.operations.move || sceneObject?.operations.resize || !sceneObject) || JSON.stringify(draftGeometry) === JSON.stringify(baseGeometry(selectedObject))} onClick={() => onStageGeometry(selectedObject, draftGeometry, editRationale)}><Sparkle size={16} />Stage in Current / Proposal</button><small className="object-editor-note">Text, table content, slide count, unsupported native internals, and source bytes are validation-locked.</small></div>; })()}</> : <><div className="thread-panel-heading"><ChatCircleDots size={19} /><div><strong>Design comments</strong><small>{selectedThreads.length} on this slide</small></div></div>
          {commentMode && <div className="thread-composer"><span className="field-label">1. Drag or click the exact area</span><div className={`anchor-readout ${draftAnchor ? "ready" : ""}`}>{draftAnchor ? <><Check size={14} />Region selected</> : <><Crosshair size={14} />Waiting for a region</>}</div><label><span className="field-label">2. Describe the adjustment</span><textarea ref={commentInput} autoFocus value={draftComment} maxLength={4000} onChange={(event) => setDraftComment(event.target.value)} placeholder="Example: Align this caption with the image edge and give it more breathing room." /></label><div className="thread-composer-actions"><button className="button ghost small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(false)}>Save note</button><button className="button primary small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(true)}><PaperPlaneTilt size={15} />Submit to AI</button></div><small>Submitting keeps this slide open for additional comments and creates a scoped thread for MCP. It does not apply or export a change.</small></div>}
          <div className="thread-list">{selectedThreads.length === 0 && <div className="thread-empty"><ChatCircleDots size={24} /><strong>No comments yet</strong><span>Select Comment on slide, then point to the exact area.</span></div>}{selectedThreads.map((thread, index) => <article key={thread.id} className="thread-item"><span>{index + 1}</span><div><strong>{thread.status === "submitted" ? "Ready for AI" : thread.status.replaceAll("-", " ")}</strong><p>{thread.comment}</p><small>Region {Math.round(thread.anchor.x * 100)}%, {Math.round(thread.anchor.y * 100)}% · {new Date(thread.createdAt).toLocaleString()}</small></div><button className="thread-delete" type="button" onClick={() => onDeleteThread(thread.id)} aria-label={`Delete comment ${index + 1}`} title="Delete comment"><Trash size={15} /></button></article>)}</div>
        </>}</aside>
      </div>
    </section>}
    {!selected && <section className="current-slide-gallery panel"><div className="panel-heading"><div><h2>Slides</h2><p>Select any {proposalWorkspace ? "proposed" : "current"} design for a closer, zoomable review and location-anchored comments.</p></div><span className="quiet-label">{proposalWorkspace ? "Proposal" : "Current"} · {PRESENTATION_DESIGN_STANDARD.defaults.slide.aspectRatio} target</span></div>
      {loading && <div className="slide-gallery-loading"><ArrowsClockwise className="spinner" size={25} /><strong>Building local slide previews</strong><span>Reading the embedded source, master, layouts, media, and native table structure.</span></div>}
      {!loading && !catalog && <div className="slide-gallery-loading"><Warning size={25} /><strong>Current designs could not be rendered</strong><span>The structural audit remains available; review the reported preview limitation before cleanup.</span></div>}
      {catalog && <div className="slide-grid">{catalog.slides.map((slide) => { const inventory = deck.audit?.slides.find((item) => item.number === slide.number); const count = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === slide.number).length; return <button className="slide-card" key={slide.id} onClick={() => setSelectedNumber(slide.number)}><span className="slide-canvas actual"><SlideDesignCanvas nativeRender={nativeRender} slideNumber={slide.number} catalog={catalog} layout={slide} label={`Open current slide ${slide.number}: ${slide.title}`} />{count > 0 && <span className="slide-comment-count"><ChatCircleDots size={12} />{count}</span>}</span><span className="slide-meta"><span>{slide.number}</span><span><strong>{inventory?.title ?? slide.title}</strong><small>{inventory?.tableCount ? `${inventory.tableCount} table${inventory.tableCount === 1 ? "" : "s"} · ` : ""}{inventory?.pictureCount ? `${inventory.pictureCount} image${inventory.pictureCount === 1 ? "" : "s"} · ` : ""}{inventory?.fonts.length ?? 0} fonts</small></span></span></button>; })}</div>}
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
    return <g transform={transform}>{element.placeholderType && <rect className="template-text-placeholder-guide" x={x} y={y} width={width} height={height} />}<text x={textX} y={textY} fill={element.textColor ?? "#373A36"} fontFamily={previewFontStack(element.fontFamily)} fontSize={fontSize} fontWeight={element.fontWeight} textAnchor={element.textAlign === "center" ? "middle" : element.textAlign === "right" ? "end" : "start"} dominantBaseline="hanging">{lines.map((line, index) => <tspan key={`${element.id}-${index}`} x={textX} dy={index === 0 ? 0 : lineHeight}>{line}</tspan>)}</text></g>;
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

function SlideDesignCanvas({ nativeRender, slideNumber, catalog, layout, label }: { nativeRender?: NativeRenderResult; slideNumber: number; catalog?: PreviewCanvasCatalog; layout?: TemplateLayoutPreview; label: string }) {
  const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((slide) => slide.number === slideNumber) : undefined;
  const nativeSource = useMemo(() => nativeSlide ? `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}` : undefined, [nativeSlide]);
  if (nativeSource) return <img className="native-slide-render" src={nativeSource} width={nativeSlide?.width} height={nativeSlide?.height} alt={label} />;
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

function StudioNodeView({ node, scene, catalog, nativeSlideSource, selected, onPointerDown }: { node: StudioWebNode; scene: StudioWebScene; catalog?: SlideRenderCatalog; nativeSlideSource?: string; selected: boolean; onPointerDown?: (event: ReactPointerEvent<HTMLElement>, mode: "move" | "resize") => void }) {
  if (!node.visible) return null;
  const percent = (value: number, total: number) => `${value / total * 100}%`;
  const style = {
    left: percent(node.frame.x, scene.slideSize.width), top: percent(node.frame.y, scene.slideSize.height), width: percent(node.frame.width, scene.slideSize.width), height: percent(node.frame.height, scene.slideSize.height),
    transform: node.frame.rotation ? `rotate(${node.frame.rotation}deg)` : undefined,
    zIndex: node.zIndex + 10,
    color: node.style.color,
    background: node.style.background,
    borderColor: node.style.borderColor,
    borderWidth: `${node.style.borderWidthPt}px`,
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
  const content = node.kind === "image" && media ? <img src={media} alt={node.name} style={{ objectFit: node.style.objectFit ?? "contain" }} />
    : node.kind === "image" && nativeCrop ? nativeCrop
    : node.kind === "table" && node.table ? <span className="studio-native-table" style={{ gridTemplateColumns: `repeat(${Math.max(1, node.table.columns)}, minmax(0, 1fr))` }}>{node.table.cells.map((cell) => <span key={cell.id} className={cell.row === 1 ? "header" : "body"} style={{ gridColumn: `${cell.column} / span ${cell.columnSpan}`, gridRow: `${cell.row} / span ${cell.rowSpan}`, background: cell.fill && /^#[0-9a-f]{6}$/i.test(cell.fill) ? cell.fill : undefined }}>{cell.text}</span>)}</span>
      : node.kind === "text" ? <span className="studio-text-content">{node.text}</span>
        : nativeCrop ?? <span className="studio-native-object-label">{node.name}</span>;
  return <div role="button" tabIndex={0} className={`studio-node kind-${node.kind} role-${node.role} ${selected ? "selected" : ""} ${node.locked ? "locked" : ""}`} style={style} onPointerDown={(event) => onPointerDown?.(event, "move")} aria-label={`${node.name}${node.locked ? ", locked" : ""}`}>{content}{selected && !node.locked && <span className="studio-resize-handle" onPointerDown={(event) => { event.stopPropagation(); onPointerDown?.(event, "resize"); }} />}</div>;
}

function StudioWebCanvas({ scene, slide, catalog, templateCatalog, nativeSlideSource, selectedNodeId, onSelectNode, onMoveNode }: { scene: StudioWebScene; slide: StudioWebScene["slides"][number]; catalog?: SlideRenderCatalog; templateCatalog?: TemplateCatalog; nativeSlideSource?: string; selectedNodeId?: string; onSelectNode: (nodeId: string) => void; onMoveNode: (nodeId: string, frame: StudioWebFrame) => void }) {
  const canvas = useRef<HTMLDivElement>(null);
  const drag = useRef<{ nodeId: string; mode: "move" | "resize"; startX: number; startY: number; frame: StudioWebFrame } | undefined>(undefined);
  const layout = slide.targetLayoutId ? templateCatalog?.layouts.find((item) => item.id === slide.targetLayoutId) : undefined;
  function begin(event: ReactPointerEvent<HTMLElement>, node: StudioWebNode, mode: "move" | "resize") {
    event.stopPropagation();
    onSelectNode(node.id);
    if (node.locked || !canvas.current) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    drag.current = { nodeId: node.id, mode, startX: event.clientX, startY: event.clientY, frame: { ...node.frame } };
  }
  function move(event: ReactPointerEvent<HTMLDivElement>) {
    if (!drag.current || !canvas.current || event.buttons === 0) return;
    const bounds = canvas.current.getBoundingClientRect();
    const deltaX = (event.clientX - drag.current.startX) / bounds.width * scene.slideSize.width;
    const deltaY = (event.clientY - drag.current.startY) / bounds.height * scene.slideSize.height;
    const next = drag.current.mode === "resize"
      ? { ...drag.current.frame, width: drag.current.frame.width + deltaX, height: drag.current.frame.height + deltaY }
      : { ...drag.current.frame, x: drag.current.frame.x + deltaX, y: drag.current.frame.y + deltaY };
    onMoveNode(drag.current.nodeId, next);
  }
  return <div ref={canvas} className="studio-web-canvas" style={{ aspectRatio: `${scene.slideSize.width} / ${scene.slideSize.height}`, background: slide.background }} onPointerMove={move} onPointerUp={() => { drag.current = undefined; }} onPointerCancel={() => { drag.current = undefined; }} onClick={() => onSelectNode("")}>
    {layout && templateCatalog && <StudioTemplateBase catalog={templateCatalog} layout={layout} />}
    {slide.recipe !== "source" && slide.recipe !== "template-layout" && <span className="studio-title-rule" />}
    {slide.nodes.map((node) => <StudioNodeView key={node.id} node={node} scene={scene} catalog={catalog} nativeSlideSource={nativeSlideSource} selected={node.id === selectedNodeId} onPointerDown={(event, mode) => begin(event, node, mode)} />)}
  </div>;
}

function StudioView({ deck, catalog, nativeRender, templateCatalog, onInitialize, onRecompose, onMoveNode, onStyleNode, onStage }: { deck?: DeckJob; catalog?: SlideRenderCatalog; nativeRender?: NativeRenderResult; templateCatalog?: TemplateCatalog; onInitialize: () => void; onRecompose: (slideNumber: number, recipe: StudioLayoutRecipe, layoutId?: string) => void; onMoveNode: (slideNumber: number, nodeId: string, frame: StudioWebFrame) => void; onStyleNode: (slideNumber: number, nodeId: string, patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">>) => void; onStage: (slideNumber: number) => void }) {
  const [selectedNumber, setSelectedNumber] = useState(1);
  const [selectedNodeId, setSelectedNodeId] = useState<string>();
  const scene = deck?.studioScene;
  useEffect(() => { setSelectedNumber(1); setSelectedNodeId(undefined); }, [deck?.id]);
  if (!deck?.audit) return <NoSelection message="Select an audited deck before entering Studio redesign mode." />;
  if (!scene) return <div className="view-stack studio-empty"><header className="view-header compact"><div><p className="eyebrow">HTML-first presentation design</p><h1>Build a Studio Web Scene</h1><p>Extract exact source content into a semantic 16:9 web canvas. The original PowerPoint remains immutable.</p></div></header><section className="designs-empty"><span className="designs-empty-icon"><Code size={34} /></span><h2>Turn this deck into an editable web design system</h2><p>Studio will preserve source bindings while giving the AI and the user one component-based canvas for layout, hierarchy, tables, imagery, and ORNL templates.</p><button className="button primary large" onClick={onInitialize}><Sparkle size={18} />Create Studio scene</button></section></div>;
  const slide = scene.slides.find((item) => item.slideNumber === selectedNumber) ?? scene.slides[0];
  const selectedNode = slide?.nodes.find((node) => node.id === selectedNodeId);
  const nativeSlide = nativeRender?.status === "ready" ? nativeRender.slides.find((item) => item.number === slide?.slideNumber) : undefined;
  const nativeSlideSource = nativeSlide ? `data:${nativeSlide.mimeType};base64,${bytesToBase64(bytesFrom(nativeSlide.bytes))}` : undefined;
  const catalogSlide = catalog?.slides.find((item) => item.number === slide?.slideNumber);
  const designedCount = scene.slides.filter((item) => item.status === "designed").length;
  const recommended = slide ? (slide.nodes.some((node) => node.kind === "table") ? "ornl-title-table" : slide.nodes.filter((node) => node.kind === "image").length >= 2 ? "ornl-title-figure-grid" : slide.nodes.some((node) => node.kind === "image") ? "ornl-title-two-column" : "ornl-title-content") as StudioLayoutRecipe : "ornl-title-content";
  return <div className="view-stack studio-view">
    <header className="view-header compact"><div><p className="eyebrow">Studio Web Scene · HTML/CSS design authority</p><h1>{deck.name}</h1><p>The user and AI edit the same semantic web canvas; export compiles supported nodes back to editable PowerPoint objects.</p></div><div className="header-actions"><span className="standard-version">{designedCount}/{scene.slides.length} designed</span><button className="button primary small" disabled={!slide || slide.status !== "designed"} onClick={() => onStage(slide.slideNumber)}><MagicWand size={16} />Stage to PowerPoint</button></div></header>
    <section className="studio-shell">
      <aside className="studio-slide-rail"><span className="field-label">Slides</span>{scene.slides.map((item) => <button key={item.id} className={item.slideNumber === slide?.slideNumber ? "selected" : ""} onClick={() => { setSelectedNumber(item.slideNumber); setSelectedNodeId(undefined); }}><span>{item.slideNumber}</span><small>{item.status === "designed" ? item.recipe.replace("ornl-", "") : "source"}</small></button>)}</aside>
      <section className="studio-stage">
        <div className="studio-toolbar"><label>Recipe<select value={slide?.recipe ?? "source"} onChange={(event) => onRecompose(slide.slideNumber, event.target.value as StudioLayoutRecipe, slide.targetLayoutId)}><option value="source">Source geometry</option><option value="ornl-title-content">ORNL title + content</option><option value="ornl-title-two-column">ORNL two column</option><option value="ornl-title-table">ORNL table</option><option value="ornl-title-figure-grid">ORNL figure grid</option><option value="template-layout">Installed template layout</option></select></label>{slide?.recipe === "template-layout" && <label>Template layout<select value={slide.targetLayoutId ?? ""} onChange={(event) => onRecompose(slide.slideNumber, "template-layout", event.target.value)}><option value="">Choose layout…</option>{templateCatalog?.layouts.map((layout) => <option key={layout.id} value={layout.id}>{layout.name}</option>)}</select></label>}<button className="button secondary small" onClick={() => onRecompose(slide.slideNumber, recommended)}><Sparkle size={15} />Use recommended</button><span className="studio-mode-chip">Web canvas</span></div>
        <StudioWebCanvas scene={scene} slide={slide} catalog={catalog} templateCatalog={templateCatalog} nativeSlideSource={nativeSlideSource} selectedNodeId={selectedNodeId} onSelectNode={(value) => setSelectedNodeId(value || undefined)} onMoveNode={(nodeId, next) => onMoveNode(slide.slideNumber, nodeId, next)} />
        <div className="studio-canvas-caption"><span><strong>{slide.recipe.replaceAll("-", " ")}</strong> · {slide.nodes.filter((node) => node.visible).length} semantic nodes · exact source copy locked</span><span>Drag elements; use the lower-right handle to resize.</span></div>
      </section>
      <aside className="studio-inspector"><span className="field-label">Design inspector</span>{selectedNode ? <><strong>{selectedNode.name}</strong><small>{selectedNode.kind} · {selectedNode.role} · {selectedNode.locked ? "locked native" : "editable"}</small><div className="studio-inspector-grid">{(["x", "y", "width", "height"] as const).map((field) => <label key={`${selectedNode.id}-${field}`}><span>{field === "width" ? "W" : field === "height" ? "H" : field.toUpperCase()}</span><input type="number" min={field === "width" || field === "height" ? .1 : 0} max={20} step={.01} defaultValue={(selectedNode.frame[field] / 914400).toFixed(2)} disabled={selectedNode.locked} onBlur={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) onMoveNode(slide.slideNumber, selectedNode.id, { ...selectedNode.frame, [field]: value * 914400 }); }} /><small>in</small></label>)}</div>{selectedNode.kind === "text" && <div className="studio-type-controls"><label><span>Size</span><input type="number" min={10} max={60} step={.25} defaultValue={selectedNode.style.fontSizePt} disabled={selectedNode.locked} onBlur={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { fontSizePt: Number(event.currentTarget.value) })} /><small>pt</small></label><label><span>Weight</span><select value={selectedNode.style.fontWeight} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { fontWeight: Number(event.currentTarget.value) as 400 | 600 | 700 })}><option value="400">Regular</option><option value="600">Semibold</option><option value="700">Bold</option></select></label><label><span>Align</span><select value={selectedNode.style.textAlign} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { textAlign: event.currentTarget.value as "left" | "center" | "right" })}><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label><label><span>Color</span><input type="color" value={selectedNode.style.color} disabled={selectedNode.locked} onChange={(event) => onStyleNode(slide.slideNumber, selectedNode.id, { color: event.currentTarget.value })} /></label></div>}{selectedNode.exactContent && <div className="inline-note"><ShieldCheck size={15} />Content remains bound to its source hash.</div>}</> : <><p>Select an element to inspect its semantic role, CSS-computed frame, and PowerPoint binding.</p><div className="inline-note"><Info size={15} />This is the new design authority. Native PowerPoint pixels remain the final export authority.</div></>}<div className="studio-source-reference"><span className="field-label">PowerPoint source reference</span><span className="studio-source-thumb">{catalogSlide ? <SlideDesignCanvas nativeRender={nativeRender} slideNumber={slide.slideNumber} catalog={catalog} layout={catalogSlide} label={`PowerPoint source slide ${slide.slideNumber}`} /> : <span className="proposal-preview-wait">Preparing…</span>}</span><small>{nativeSlideSource ? "PowerPoint-native · read only" : "Structural fallback · read only"}</small></div></aside>
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

function ResourcesView({ project, onToggleMcp, onAdd, onRemove }: { project: PresentationStudioProject; onToggleMcp: (id: string) => void; onAdd: () => void; onRemove: (id: string) => void }) {
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
          <div className="resource-row resource-head"><span>Resource</span><span>Role</span><span>Processing</span><span>Size</span><span>AI session</span><span>Project</span></div>
          {project.resources.length === 0 && <div className="resource-empty"><FileText size={25} /><span><strong>No project Resources yet</strong><small>Drop files into the app or choose files above. Nothing will remain linked to its original location.</small></span></div>}
          {project.resources.map((resource) => {
            const processingStatus = resource.processing?.status ?? "stored-only";
            const hasWarnings = Boolean(resource.processing?.warnings.length);
            return <div className="resource-row" key={resource.id} title={resource.processing?.summary}>
              <span className="resource-name"><Archive size={20} /><span><strong>{resource.name}</strong><small>{resourceKindLabels[resource.kind ?? "other"]} · {resource.sha256.slice(0, 12)}… · embedded</small></span></span>
              <span className="resource-roles">{resource.roles.join(" · ")}</span>
              <span className={`processing-state ${processingStatus}`}>{hasWarnings && <Warning size={13} />}{processingStatus === "indexed" ? "Indexed" : processingStatus === "needs-review" ? "Needs review" : "Stored only"}</span>
              <span>{formatBytes(resource.byteLength)}</span>
              <button className={`access-toggle ${resource.mcpAccess !== "none" ? "on" : ""}`} onClick={() => onToggleMcp(resource.id)}>{resource.mcpAccess === "none" ? "Not shared" : "Metadata only"}</button>
              <button className="resource-remove" onClick={() => onRemove(resource.id)} title="Remove this embedded copy from the project; the original file is never deleted"><Trash size={13} />Remove</button>
            </div>;
          })}
        </div>
      </section>
      <div className="inline-note wide"><ShieldCheck size={18} />Original bytes and extracted text remain local. Remove affects only the embedded project copy—not the source file. Metadata sharing requires both AI session access and a per-Resource choice; file bytes are never returned through MCP.</div>
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
  const [mcpActivity, setMcpActivity] = useState<{ id: string; operation: string; state: "active" | "completed" | "failed" }>();
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
  const [slideWorkspaceRequest, setSlideWorkspaceRequest] = useState<SlideWorkspaceRequest>();
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
  const onboardingChecked = useRef(false);
  const auditGeometryUpgradeAttempted = useRef(new Set<string>());
  const desktop = window.presentationStudioDesktop;
  const selectedDeck = project.decks.find((deck) => deck.id === selectedDeckId) ?? project.decks[0];

  useEffect(() => { projectRef.current = project; }, [project]);
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
    if (activeView !== "designs" || !desktop || !templateCatalog || !templateSourceBytes) return;
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
    const proposalWorkspace = activeView === "slides" && slideWorkspaceRequest?.deckId === selectedDeck?.id && slideWorkspaceRequest.representation === "proposal";
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
      setMcpActivity({ id: activityId, operation: request.operation, state: "active" });
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
        const slideImage = nativeRender.slides.find((item) => item.number === slideNumber);
        if (!slideImage) throw new Error("PowerPoint did not return the requested inspection image.");
        const images = await inspectionRasterEvidence(slideImage, packet.visualEvidence.crops);
        return { ...packet, representation, images };
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
          scene: { schema: scene.schema, version: scene.version, revision: scene.revision, slideSize: scene.slideSize, designSystem: scene.designSystem, persisted: Boolean(deck.studioScene) },
          slide: {
            id: slide.id, slideNumber: slide.slideNumber, recipe: slide.recipe, targetLayoutId: slide.targetLayoutId, targetLayoutName: slide.targetLayoutName, status: slide.status, designRationale: slide.designRationale, sourceTextHash: slide.sourceTextHash,
            recommendedRecipe: recommendedStudioRecipe(slide),
            nodes: slide.nodes.map((node) => ({
              id: node.id, sourceObjectId: node.sourceObjectId, name: node.name, kind: node.kind, role: node.role, visible: node.visible, locked: node.locked, exactContent: node.exactContent, text: node.text, textHash: node.textHash, tableId: node.tableId, table: node.table,
              sourceFrameInches: { x: node.sourceFrame.x / 914_400, y: node.sourceFrame.y / 914_400, width: node.sourceFrame.width / 914_400, height: node.sourceFrame.height / 914_400, rotation: node.sourceFrame.rotation },
              frameInches: { x: node.frame.x / 914_400, y: node.frame.y / 914_400, width: node.frame.width / 914_400, height: node.frame.height / 914_400, rotation: node.frame.rotation },
              style: node.style,
            })),
          },
          instruction: "Design in this semantic HTML/CSS scene, not by preserving weak source coordinates. Choose one shared ORNL recipe or installed template layout, then use stage_studio_web_design. Exact copy and table cells are locked; PowerPoint-native Current/Proposal renders remain the visual acceptance authority.",
        };
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
        const iteration = decideVisualIteration({ priorHistory, requestedVerdict, rationale: String(request.input.rationale ?? ""), slideNumber, inspectionRevision: expectedInspectionRevision, currentRasterSha256: currentSlide.sha256, proposalRasterSha256: proposalSlide.sha256, changedPixelRatio: comparison.metrics.changedPixelRatio, improvements: metricEvaluation.improvements, regressions: metricEvaluation.regressions });
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
        return { updatedAt: current.project.updatedAt, threads: threads.map((thread) => ({ id: thread.id, deckId: thread.deckId, slideId: thread.slideId, slideNumber: thread.slideNumber, baseRevision: thread.baseRevision, anchor: thread.anchor, comment: thread.comment, status: thread.status, createdAt: thread.createdAt, submittedAt: thread.submittedAt })) };
      }
      if (request.operation === "get_design_thread") {
        const thread = current.designThreads.find((item) => item.id === request.input.threadId);
        if (!thread || !["submitted", "needs-reanchor"].includes(thread.status)) throw new Error("The requested design thread is no longer active or was not submitted to AI in this project.");
        const deck = current.decks.find((item) => item.id === thread.deckId);
        const slide = deck?.audit?.slides.find((item) => item.id === thread.slideId || item.number === thread.slideNumber);
        return { updatedAt: current.project.updatedAt, thread, deck: deck ? { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId } : null, slide: slide ? { id: slide.id, number: slide.number, title: slide.title, textHash: slide.textHash, objects: { tables: slide.tableCount, pictures: slide.pictureCount, charts: slide.chartCount } } : null, instruction: "Read the current revision and get_slide_render before staging a bounded fix. Do not guess if the anchor no longer maps unambiguously." };
      }
      if (request.operation === "stage_studio_web_design") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read get_studio_web_scene again before staging a Studio design.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit || !deck.scene) throw new Error("The requested deck does not have a current PowerPoint audit and preservation scene.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const allowedRecipes: StudioLayoutRecipe[] = ["source", "ornl-title-content", "ornl-title-two-column", "ornl-title-table", "ornl-title-figure-grid", "template-layout"];
        const recipe = String(request.input.recipe ?? "") as StudioLayoutRecipe;
        if (!allowedRecipes.includes(recipe)) throw new Error("Choose a supported Studio web recipe.");
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
        const adoptedDeck: DeckJob = { ...deck, operationScope: "reflow", studioScene };
        const compiled = await compileStudioProposal(adoptedDeck, studioScene, slideNumber, current.project.updatedAt, "ai");
        const proposal = compiled.proposal;
        const next = touchProject({
          ...current,
          designThreads: removeAddressedDesignThreads(current.designThreads, deck.id, proposal, requestedAddressedThreadIds(request.input)),
          decks: current.decks.map((item) => item.id === deck.id ? { ...adoptedDeck, proposal, status: "proposal-ready" as const } : item),
        }, "mcp-studio-web-design-staged", `AI recomposed slide ${slideNumber} of ${deck.name} with the shared ${recipe} Studio web recipe and compiled ${compiled.geometryCount} source-bound geometry edits to editable PowerPoint; source bytes remain unchanged.`);
        proposal.baseUpdatedAt = next.project.updatedAt;
        projectRef.current = next;
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return {
          projectUpdatedAt: next.project.updatedAt,
          studioSceneRevision: studioScene.revision,
          slide: { number: slideNumber, recipe, layoutId: layout?.id, layoutName: layout?.name, nodeFrameOverrideCount: nodeFrames.length, nodeStyleOverrideCount: nodeStyles.length },
          proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, geometryCount: compiled.geometryCount },
          applied: false,
          saved: false,
          instruction: "Presentation Studio is showing the staged design in Review. Inspect the authoritative PowerPoint-native Current/Proposal comparison, then revise, reject, or leave it for human acceptance. The source file was not changed.",
        };
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
        if (deck.operationScope !== "reflow") throw new Error("Semantic layout solving requires the deck's Designer Cleanup reflow scope.");
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
        if (deck.operationScope !== "reflow") throw new Error("Measured text fitting requires the deck's Designer Cleanup reflow scope.");
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
        if (deck.operationScope !== "reflow") throw new Error("Native table solving requires the deck's Designer Cleanup reflow scope.");
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
        if (deck.operationScope !== "reflow") throw new Error("Native visual polish requires the deck's Designer Cleanup reflow scope.");
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
        if (deck.operationScope !== "reflow") throw new Error("Native layout remapping requires the deck's Designer Cleanup reflow scope; cleanup-only cannot change master/layout relationships.");
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
        if (deck.operationScope !== "reflow") throw new Error("Semantic recomposition requires the deck's Designer Cleanup reflow scope; cleanup-only permits only bounded cleanup rules.");
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
      setMcpActivity({ id: activityId, operation: request.operation, state: "completed" });
      window.setTimeout(() => setMcpActivity((current) => current?.id === activityId ? undefined : current), 8000);
      return result;
      } catch (caught) {
        setMcpActivity({ id: activityId, operation: request.operation, state: "failed" });
        window.setTimeout(() => setMcpActivity((current) => current?.id === activityId ? undefined : current), 10000);
        throw caught;
      }
    });
  }, [desktop, getOrBuildNativeRender, getOrBuildProposalCatalog, getOrBuildSlideCatalog, getOrBuildTemplateNativeRender, mcpEnabled, presentationFontCss, templateCatalog, templateNativeLoading, templateNativeRender]);

  function clearMessages() { setNotice(undefined); setError(undefined); }

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
          const processed = await processResourceInput({ name: files[index].name, filePath: files[index].filePath, mediaType: files[index].mediaType, bytes });
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
    setSlideCatalogs({});
    setProposalCatalogs({});
    setNativeRenderCatalogs({});
    setSlideCatalogLoadingDeckId(undefined);
    setProposalCatalogLoadingDeckId(undefined);
    setNativeRenderLoadingKey(undefined);
    setSlideWorkspaceRequest(undefined);
  }

  function adoptOpenedProject(opened: PresentationStudioProject, password?: string) {
    resetProjectRenderState();
    projectRef.current = opened;
    setProject(opened);
    setSelectedDeckId(opened.decks[0]?.id);
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
        : "";
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
      const studioScene = recomposeStudioWebSlide(selectedDeck.studioScene, slideNumber, recipe, layout, `Recompose slide ${slideNumber} with the shared ${recipe} web layout while retaining exact source content and editable PowerPoint bindings.`);
      setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, operationScope: "reflow", studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }, "studio-slide-recomposed", `Recomposed slide ${slideNumber} with ${recipe} in the Studio Web Scene; no PowerPoint file was changed.`));
      setNotice(`Slide ${slideNumber} now uses the ${recipe.replaceAll("-", " ")} web composition. Drag elements to refine it or stage it for PowerPoint review.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The slide could not be recomposed.");
    }
  }

  async function compileStudioProposal(deck: DeckJob, studioScene: StudioWebScene, slideNumber: number, updatedAt: string, author: "human" | "ai") {
    const geometry = studioGeometryRequests(deck, studioScene, slideNumber, author);
    const studioSlide = studioScene.slides.find((item) => item.slideNumber === slideNumber);
    if (!studioSlide) throw new Error(`Slide ${slideNumber} is not present in the Studio Web Scene.`);
    let baseProposal: CleanupProposal;
    if (studioSlide.recipe === "template-layout") {
      if (!templateCatalog || !templateSourceBytes || !studioSlide.targetLayoutId) throw new Error("The installed Template Pack is required to compile this web layout to its real PowerPoint master and custom layout.");
      const layout = templateCatalog.layouts.find((item) => item.id === studioSlide.targetLayoutId);
      if (!layout) throw new Error("The Studio slide references a template layout that is no longer installed.");
      const command = {
        id: `studio-native-layout-slide-${slideNumber}`,
        slideNumber,
        templateSha256: templateCatalog.sha256,
        templateLayoutPart: layout.sourcePart,
        templateLayoutSha256: await templateLayoutPartSha256(templateSourceBytes, layout.sourcePart),
        templateLayoutName: layout.name,
        rationale: `${studioSlide.designRationale} Compile the same installed template layout used by the Studio web canvas into PowerPoint.`.slice(0, 1_000),
        author,
      } as const;
      baseProposal = createNativeLayoutRecompositionProposal({ ...deck, operationScope: "reflow" }, updatedAt, command, geometry);
    } else {
      baseProposal = createGeometryBatchProposal({ ...deck, operationScope: "reflow" }, updatedAt, geometry);
    }
    const proposal = createVisualDesignProposal({ ...deck, proposal: baseProposal, operationScope: "reflow" }, updatedAt, studioVisualDesignRequest(studioScene, slideNumber, author));
    return { proposal, geometryCount: geometry.length };
  }

  function moveStudioNode(slideNumber: number, nodeId: string, nextFrame: StudioWebFrame) {
    if (!selectedDeck?.studioScene) return;
    try {
      const studioScene = updateStudioWebNodeFrame(selectedDeck.studioScene, slideNumber, nodeId, nextFrame);
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio element could not be moved.");
    }
  }

  function styleStudioNode(slideNumber: number, nodeId: string, patch: Partial<Pick<StudioWebNode["style"], "fontSizePt" | "fontWeight" | "color" | "textAlign" | "verticalAlign" | "objectFit">>) {
    if (!selectedDeck?.studioScene) return;
    try {
      const studioScene = updateStudioWebNodeStyle(selectedDeck.studioScene, slideNumber, nodeId, patch);
      setProject((current) => ({ ...current, project: { ...current.project, updatedAt: new Date().toISOString() }, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, studioScene, proposal: undefined, status: "ready-for-cleanup" } : deck) }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio element style could not be updated.");
    }
  }

  async function stageStudioSlide(slideNumber: number) {
    if (!selectedDeck?.studioScene) return;
    clearMessages();
    setBusy(`Compiling Studio slide ${slideNumber} to editable PowerPoint…`);
    try {
      const current = projectRef.current;
      const deck = current.decks.find((item) => item.id === selectedDeck.id);
      if (!deck?.studioScene) throw new Error("Create or reopen the Studio Web Scene before compiling a slide to PowerPoint.");
      const compiled = await compileStudioProposal(deck, deck.studioScene, slideNumber, current.project.updatedAt, "human");
      const proposal = compiled.proposal;
      const next = touchProject({ ...current, decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" } : item) }, "studio-slide-compiled", `Compiled slide ${slideNumber}'s semantic HTML/CSS scene into ${compiled.geometryCount} editable PowerPoint geometry bindings plus shared ORNL typography and components.`);
      proposal.baseUpdatedAt = next.project.updatedAt;
      projectRef.current = next;
      setProject(next);
      setActiveView("review");
      setNotice(`Slide ${slideNumber} was compiled from the web scene into an editable PowerPoint proposal. Compare the PowerPoint-native Current and Proposal renders before accepting it.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The Studio slide could not be compiled to PowerPoint.");
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

  function saveDesignThread(slide: SlideRenderPreview, anchor: DesignThread["anchor"], comment: string, submit: boolean) {
    if (!selectedDeck?.audit) return;
    const inventory = selectedDeck.audit.slides.find((item) => item.number === slide.number);
    if (!inventory) { setError("The selected slide revision could not be resolved."); return; }
    const now = new Date().toISOString();
    const thread: DesignThread = { id: crypto.randomUUID(), deckId: selectedDeck.id, slideId: inventory.id, slideNumber: slide.number, baseRevision: projectRef.current.project.updatedAt, anchor, comment, status: submit ? "submitted" : "note", createdAt: now, updatedAt: now, submittedAt: submit ? now : undefined };
    setProject((current) => {
      const next = touchProject({ ...current, designThreads: [...current.designThreads, thread] }, submit ? "design-thread-submitted" : "design-note-saved", `${submit ? "Submitted" : "Saved"} a location-anchored design comment on slide ${slide.number} of ${selectedDeck.name}.`);
      return { ...next, decks: next.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal && ["pending", "applied"].includes(deck.proposal.status) ? { ...deck, proposal: { ...deck.proposal, baseUpdatedAt: next.project.updatedAt, slideReviews: submit ? [...(deck.proposal.slideReviews ?? []).filter((review) => review.slideNumber !== slide.number), { slideNumber: slide.number, decision: "changes-requested", reviewedAt: now, comment }] : deck.proposal.slideReviews } } : deck) };
    });
    setNotice(submit ? `Design comment submitted for AI on slide ${slide.number}. It is anchored to the selected region and no change was applied.` : `Private design note saved on slide ${slide.number}. It is not available through MCP.`);
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

  const mainContent = useMemo(() => {
    if (activeView === "batch") return <BatchView project={project} selectedId={selectedDeck?.id} onSelect={(id) => { setSelectedDeckId(id); setActiveView("decks"); }} onAdd={() => void addDecks()} />;
    if (activeView === "decks") return <DeckAuditView deck={selectedDeck} onConfirm={confirmTemplate} onStage={stageCleanup} onStartOrnlCleanup={startOrnlCleanup} onMarkExemplar={markTableExemplar} onExportReport={() => void exportAuditReport()} isExemplar={Boolean(selectedDeck && project.styleExemplars.some((item) => item.deckId === selectedDeck.id && item.kind === "table"))} />;
    if (activeView === "slides") { const proposalWorkspace = selectedDeck?.proposal?.status === "applied" || (slideWorkspaceRequest?.deckId === selectedDeck?.id && slideWorkspaceRequest.representation === "proposal"); return <SlidesView deck={selectedDeck} catalog={selectedDeck ? proposalWorkspace ? proposalCatalogs[selectedDeck.id] : slideCatalogs[selectedDeck.id] : undefined} nativeRender={selectedDeck ? proposalWorkspace ? nativeRenderCatalogs[`${selectedDeck.id}:proposal`] : nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} loading={Boolean(selectedDeck && (proposalWorkspace ? proposalCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:proposal` : slideCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:current`))} revision={project.project.updatedAt} threads={project.designThreads} openRequest={slideWorkspaceRequest} onSaveThread={saveDesignThread} onDeleteThread={deleteDesignThread} onStageGeometry={stageGeometryEdit} />; }
    if (activeView === "studio") return <StudioView deck={selectedDeck} catalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} nativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} templateCatalog={templateCatalog} onInitialize={() => void initializeStudioScene()} onRecompose={recomposeStudioSlide} onMoveNode={moveStudioNode} onStyleNode={styleStudioNode} onStage={stageStudioSlide} />;
    if (activeView === "designs") return <DesignsView catalog={templateCatalog} installedAt={templateInstalledAt} loading={templateLoading} nativeRender={templateNativeRender} nativeLoading={templateNativeLoading} onInstall={() => void installTemplate()} />;
    if (activeView === "rules") return <RulesView deck={selectedDeck} exemplarCount={project.styleExemplars.filter((item) => item.kind === "table").length} />;
    if (activeView === "review") return <ReviewView deck={selectedDeck} projectUpdatedAt={project.project.updatedAt} currentCatalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} proposalCatalog={selectedDeck ? proposalCatalogs[selectedDeck.id] : undefined} currentNativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:current`] : undefined} proposalNativeRender={selectedDeck ? nativeRenderCatalogs[`${selectedDeck.id}:proposal`] : undefined} previewLoading={Boolean(selectedDeck && (proposalCatalogLoadingDeckId === selectedDeck.id || nativeRenderLoadingKey === `${selectedDeck.id}:proposal`))} threads={project.designThreads} onToggle={toggleChange} onReviewSlide={reviewSlide} onRequestChanges={requestSlideChanges} onDeleteThread={deleteDesignThread} onOpenSlide={openSlideWorkspace} onReject={rejectProposal} onApply={acceptProposal} onExport={() => void exportCleaned()} />;
    return <ResourcesView project={project} onAdd={() => void addResources()} onRemove={removeResource} onToggleMcp={(id) => setProject((current) => touchProject({ ...current, resources: current.resources.map((resource) => resource.id === id ? { ...resource, mcpAccess: resource.mcpAccess === "none" ? "metadata" : "none" } : resource) }, "resource-access-updated", "Updated session-only Resource metadata permission."))} />;
  }, [activeView, nativeRenderCatalogs, nativeRenderLoadingKey, project, proposalCatalogLoadingDeckId, proposalCatalogs, selectedDeck, slideCatalogLoadingDeckId, slideCatalogs, slideWorkspaceRequest, templateCatalog, templateInstalledAt, templateLoading, templateNativeLoading, templateNativeRender]);

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
          <button className={`ai-session ${mcpEnabled ? "enabled" : ""}`} data-tour="ai-session" onClick={() => setMcpEnabled((value) => !value)}><span className="ai-icon"><Sparkle size={18} /></span><span><strong>AI session</strong><small>{mcpEnabled ? "Audit metadata allowed" : "Access off"}</small></span><span className="toggle-knob" /></button>
          <div className="local-status"><span className={mcpStatus.available ? "online" : ""} /><span>{mcpStatus.available ? "Local MCP ready" : "Browser preview"}</span></div>
          {desktop && <div className="local-status native-qa" title={nativeReadiness.reason}><span className={nativeReadiness.ready ? "online" : nativeReadiness.sessionLocked ? "locked" : ""} /><span>{nativeReadiness.ready ? "PowerPoint QA ready" : nativeReadiness.sessionLocked ? "Unlock Mac for native QA" : "Native QA unavailable"}</span></div>}
        </div>
      </nav>
      <main className="workspace">{mainContent}</main>
      <Inspector deck={selectedDeck} onOpenReview={() => setActiveView("review")} />
      <OnboardingTour open={tourOpen} stepIndex={tourStepIndex} onStepChange={setTourStepIndex} onClose={closeOnboardingTour} />
      {mcpActivity && <div className={`mcp-activity ${mcpActivity.state}`} role="status" aria-live="polite"><span>{mcpActivity.state === "active" ? <ArrowsClockwise className="spinner" size={17} /> : mcpActivity.state === "completed" ? <CheckCircle size={17} /> : <Warning size={17} />}</span><div><strong>{mcpActivity.state === "active" ? "AI is using Presentation Studio" : mcpActivity.state === "completed" ? "AI operation completed" : "AI operation needs attention"}</strong><small>{mcpActivity.operation.replaceAll("_", " ")} · {mcpActivity.state === "active" ? "local operation in progress" : mcpActivity.state}</small></div></div>}
      {(notice || error) && <div className={`toast ${error ? "error" : "success"}`}><span>{error ? <Warning size={18} /> : <CheckCircle size={18} />}</span><p>{error ?? notice}</p><button onClick={clearMessages}><X size={16} /></button></div>}
      {fileDragActive && <div className="file-drop-overlay"><div className="file-drop-card"><UploadSimple size={38} /><strong>Drop to open a project or add Resources</strong><span>A single .pstudio package opens as the active project. Other files are processed locally, embedded by hash, and preserved with the current project.</span></div></div>}
      {busy && <div className="busy-overlay"><div className="busy-card"><ArrowsClockwise className="spinner" size={25} /><strong>{busy}</strong><span>Files are processed locally and copied into the project. External originals remain untouched.</span></div></div>}
      <input id="web-deck-picker" hidden type="file" accept=".pptx" multiple onChange={(event) => { void importWebFiles(event.target.files, "decks"); event.currentTarget.value = ""; }} />
      <input id="web-resource-picker" hidden type="file" accept=".pptx,.potx,.docx,.pdf,.md,.markdown,.txt,.csv,.tsv,.json,.xlsx,.png,.jpg,.jpeg,.webp,.tif,.tiff,.svg,.wav,.mp3,.m4a,.mp4,.mov,.doc,.xls" multiple onChange={(event) => { void importWebFiles(event.target.files, "resources"); event.currentTarget.value = ""; }} />
      <input id="web-template-picker" hidden type="file" accept=".pptx,.potx" onChange={(event) => { const file = event.target.files?.[0]; if (file) void file.arrayBuffer().then((buffer) => loadTemplate({ name: file.name, filePath: file.name, mediaType: file.type, bytes: new Uint8Array(buffer) }, false)); event.currentTarget.value = ""; }} />
    </div>
  );
}
