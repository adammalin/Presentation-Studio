import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent, type PointerEvent as ReactPointerEvent } from "react";
import {
  Archive,
  ArrowLeft,
  ArrowRight,
  ArrowsClockwise,
  CaretRight,
  ChatCircleDots,
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
  TemplateClassification,
} from "./types";
import { applyCleanupToPptx, buildCleanupProposalPptx, createDesignerCleanupProposal, createFontCleanupProposal } from "./lib/cleanup";
import type { LocalPresentationFont, PickedBinaryFile } from "./lib/desktop";
import { decryptProjectPackage, encryptProjectPackage, isEncryptedProject } from "./lib/encryption";
import { auditPptx } from "./lib/pptx-audit";
import { createProject, touchProject } from "./lib/project";
import { buildProjectPackage, openProjectPackage } from "./lib/project-package";
import {
  isPowerPointResource,
  MAX_PROJECT_RESOURCE_BYTES,
  MAX_SINGLE_RESOURCE_BYTES,
  processResourceInput,
} from "./lib/resource-ingestion";
import { buildAuditReport } from "./lib/report";
import { createOrnlDesignProfile, designStandardSummary, PRESENTATION_DESIGN_STANDARD } from "./lib/design-standard";
import { slidePreviewJpeg } from "./lib/slide-preview";
import { buildSlideRenderCatalog, buildTemplateCatalog, type SlideRenderCatalog, type SlideRenderPreview, type TemplateCatalog, type TemplateLayoutPreview, type TemplatePreviewElement } from "./lib/template-catalog";
import {
  ONBOARDING_TOUR_STORAGE_KEY,
  ONBOARDING_TOUR_VERSION,
  shouldShowOnboardingTour,
} from "./lib/onboarding";

type ViewId = "batch" | "decks" | "slides" | "designs" | "rules" | "review" | "resources";

const navItems: Array<{ id: ViewId; label: string; icon: Icon }> = [
  { id: "batch", label: "Batch", icon: Files },
  { id: "decks", label: "Deck audit", icon: PresentationChart },
  { id: "slides", label: "Slides", icon: Slideshow },
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

function previewFontStack(fontFamily?: string): string {
  const requested = (fontFamily ?? "Aptos").replaceAll('"', "").trim() || "Aptos";
  return /^aptos(?:\s|$)/i.test(requested) ? `"${requested}", Arial, sans-serif` : `"${requested}", Aptos, Arial, sans-serif`;
}

function fontFaceRule(font: LocalPresentationFont): string {
  return `@font-face{font-family:"${font.family}";src:url("data:${font.mediaType};base64,${bytesToBase64(font.bytes)}") format("truetype");font-weight:${font.weight};font-style:${font.style};font-display:block;}`;
}

function cleanFileStem(name: string): string {
  return name.replace(/\.pptx$/i, "");
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
          {audit.findings.map((item) => <div className="finding-row" key={item.id}><span className={`finding-icon ${item.severity}`}>{item.severity === "error" || item.severity === "warning" ? <Warning size={17} /> : <Info size={17} />}</span><div><span className="finding-category">{item.category}</span><strong>{item.message}</strong><p>{item.evidence}</p></div><span className="confidence">{item.confidence}</span></div>)}
        </div>
      </section>
    </div>
  );
}

function SlidesView({ deck, catalog, loading, revision, threads, onSaveThread }: { deck?: DeckJob; catalog?: SlideRenderCatalog; loading: boolean; revision: string; threads: DesignThread[]; onSaveThread: (slide: SlideRenderPreview, anchor: DesignThread["anchor"], comment: string, submit: boolean) => void }) {
  const [selectedNumber, setSelectedNumber] = useState<number>();
  const [commentMode, setCommentMode] = useState(false);
  const [draftAnchor, setDraftAnchor] = useState<DesignThread["anchor"]>();
  const [draftComment, setDraftComment] = useState("");
  const dragStart = useRef<{ x: number; y: number } | undefined>(undefined);
  useEffect(() => { setSelectedNumber(undefined); setCommentMode(false); setDraftAnchor(undefined); setDraftComment(""); }, [deck?.id]);
  if (!deck?.audit) return <NoSelection message="Select a deck to inspect its current slide designs." />;

  const selected = catalog?.slides.find((slide) => slide.number === selectedNumber);
  const selectedThreads = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === selectedNumber);
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
    setCommentMode(false);
  }

  return <div className="view-stack slides-view">
    <header className="view-header compact"><div><p className="eyebrow">Current slide designs</p><h1>{deck.name}</h1><p>Revision-bound local previews from the embedded editable PowerPoint. Native PowerPoint remains the final fidelity check.</p></div><span className="render-status"><span className={loading ? "loading" : catalog ? "ready" : ""} />{loading ? "Rendering current slides…" : catalog ? `${catalog.slides.length} current previews` : "Preview unavailable"}</span></header>
    {selected && catalog && <section className="slide-review panel">
      <div className="slide-review-toolbar"><div><button className="button ghost small" onClick={() => setSelectedNumber(undefined)}><ArrowLeft size={15} />Gallery</button><span><strong>Slide {selected.number}</strong><small>Current · revision {revision.slice(0, 19).replace("T", " ")}</small></span></div><div><button className="button ghost small" disabled={selected.number <= 1} onClick={() => setSelectedNumber(selected.number - 1)}><ArrowLeft size={15} />Previous</button><button className="button ghost small" disabled={selected.number >= catalog.slides.length} onClick={() => setSelectedNumber(selected.number + 1)}>Next<ArrowRight size={15} /></button><button className={`button small ${commentMode ? "primary" : "secondary"}`} onClick={() => { setCommentMode((value) => !value); setDraftAnchor(undefined); }}><Crosshair size={16} />{commentMode ? "Select a region" : "Comment on slide"}</button></div></div>
      <div className="slide-review-body">
        <div className="slide-review-stage">
          <div className="slide-review-canvas"><TemplateLayoutCanvas catalog={catalog} layout={selected} label={`Current design for slide ${selected.number}: ${selected.title}`} /><div className={`slide-anchor-layer ${commentMode ? "active" : ""}`} aria-label={commentMode ? "Drag over the exact slide region to comment" : "Slide comment anchors"} onPointerDown={beginAnchor} onPointerMove={moveAnchor} onPointerUp={finishAnchor}>
            {selectedThreads.map((thread, index) => <span key={thread.id} className={`thread-anchor ${thread.status}`} style={{ left: `${thread.anchor.x * 100}%`, top: `${thread.anchor.y * 100}%`, width: `${thread.anchor.width * 100}%`, height: `${thread.anchor.height * 100}%` }} title={thread.comment}><i>{index + 1}</i></span>)}
            {draftAnchor && <span className="thread-anchor draft" style={{ left: `${draftAnchor.x * 100}%`, top: `${draftAnchor.y * 100}%`, width: `${draftAnchor.width * 100}%`, height: `${draftAnchor.height * 100}%` }} />}
          </div></div>
          <div className="slide-representation-note"><ShieldCheck size={15} /><span><strong>Current representation</strong> Generated locally from OOXML, master, layout, editable shapes, text, images, tables, and connectors. Verify final export in PowerPoint.</span></div>
        </div>
        <aside className="slide-thread-panel"><div className="thread-panel-heading"><ChatCircleDots size={19} /><div><strong>Design comments</strong><small>{selectedThreads.length} on this slide</small></div></div>
          {commentMode && <div className="thread-composer"><span className="field-label">1. Drag or click the exact area</span><div className={`anchor-readout ${draftAnchor ? "ready" : ""}`}>{draftAnchor ? <><Check size={14} />Region selected</> : <><Crosshair size={14} />Waiting for a region</>}</div><label><span className="field-label">2. Describe the adjustment</span><textarea value={draftComment} maxLength={4000} onChange={(event) => setDraftComment(event.target.value)} placeholder="Example: Align this caption with the image edge and give it more breathing room." /></label><div className="thread-composer-actions"><button className="button ghost small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(false)}>Save note</button><button className="button primary small" disabled={!draftAnchor || !draftComment.trim()} onClick={() => saveThread(true)}><PaperPlaneTilt size={15} />Submit to AI</button></div><small>Submitting creates a scoped thread for MCP. It does not apply or export a change.</small></div>}
          <div className="thread-list">{selectedThreads.length === 0 && <div className="thread-empty"><ChatCircleDots size={24} /><strong>No comments yet</strong><span>Select Comment on slide, then point to the exact area.</span></div>}{selectedThreads.map((thread, index) => <article key={thread.id} className="thread-item"><span>{index + 1}</span><div><strong>{thread.status === "submitted" ? "Ready for AI" : thread.status.replaceAll("-", " ")}</strong><p>{thread.comment}</p><small>Region {Math.round(thread.anchor.x * 100)}%, {Math.round(thread.anchor.y * 100)}% · {new Date(thread.createdAt).toLocaleString()}</small></div></article>)}</div>
        </aside>
      </div>
    </section>}
    {!selected && <section className="current-slide-gallery panel"><div className="panel-heading"><div><h2>Slides</h2><p>Select any current design for a closer, zoomable review and location-anchored comments.</p></div><span className="quiet-label">Current · {PRESENTATION_DESIGN_STANDARD.defaults.slide.aspectRatio} target</span></div>
      {loading && <div className="slide-gallery-loading"><ArrowsClockwise className="spinner" size={25} /><strong>Building local slide previews</strong><span>Reading the embedded source, master, layouts, media, and native table structure.</span></div>}
      {!loading && !catalog && <div className="slide-gallery-loading"><Warning size={25} /><strong>Current designs could not be rendered</strong><span>The structural audit remains available; review the reported preview limitation before cleanup.</span></div>}
      {catalog && <div className="slide-grid">{catalog.slides.map((slide) => { const inventory = deck.audit?.slides.find((item) => item.number === slide.number); const count = threads.filter((thread) => thread.deckId === deck.id && thread.slideNumber === slide.number).length; return <button className="slide-card" key={slide.id} onClick={() => setSelectedNumber(slide.number)}><span className="slide-canvas actual"><TemplateLayoutCanvas catalog={catalog} layout={slide} label={`Open current slide ${slide.number}: ${slide.title}`} />{count > 0 && <span className="slide-comment-count"><ChatCircleDots size={12} />{count}</span>}</span><span className="slide-meta"><span>{slide.number}</span><span><strong>{inventory?.title ?? slide.title}</strong><small>{inventory?.tableCount ? `${inventory.tableCount} table${inventory.tableCount === 1 ? "" : "s"} · ` : ""}{inventory?.pictureCount ? `${inventory.pictureCount} image${inventory.pictureCount === 1 ? "" : "s"} · ` : ""}{inventory?.fonts.length ?? 0} fonts</small></span></span></button>; })}</div>}
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

type DesignFilter = "all" | TemplateLayoutPreview["category"];

function DesignsView({ catalog, installedAt, loading, onInstall }: { catalog?: TemplateCatalog; installedAt?: string; loading: boolean; onInstall: () => void }) {
  const [filter, setFilter] = useState<DesignFilter>("all");
  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState<string>();
  useEffect(() => {
    if (!catalog?.layouts.length) return;
    if (!catalog.layouts.some((layout) => layout.id === selectedId)) setSelectedId(catalog.layouts[0].id);
  }, [catalog, selectedId]);
  if (!catalog) return <div className="view-stack designs-empty-view"><header className="view-header compact"><div><p className="eyebrow">Authorized template library</p><h1>Slide designs</h1><p>Browse the real masters and layouts available for composition and reflow.</p></div></header><section className="designs-empty"><span className="designs-empty-icon"><SquaresFour size={34} weight="light" /></span><h2>Install an authorized PowerPoint template</h2><p>Presentation Studio will read its master, layouts, media, theme, and placeholder geometry locally. The source template stays outside Git and is never uploaded.</p><button className="button primary large" disabled={loading} onClick={onInstall}><UploadSimple size={18} />{loading ? "Reading template…" : "Choose POTX or PPTX"}</button></section></div>;

  const normalizedQuery = query.trim().toLowerCase();
  const visibleLayouts = catalog.layouts.filter((layout) => (filter === "all" || layout.category === filter) && (!normalizedQuery || layout.name.toLowerCase().includes(normalizedQuery) || layout.placeholderTypes.some((type) => type.toLowerCase().includes(normalizedQuery))));
  const selected = catalog.layouts.find((layout) => layout.id === selectedId) ?? catalog.layouts[0];
  const categoryCounts = catalog.layouts.reduce<Record<string, number>>((counts, layout) => ({ ...counts, [layout.category]: (counts[layout.category] ?? 0) + 1 }), {});
  return <div className="view-stack designs-view">
    <header className="view-header compact"><div><p className="eyebrow">Authorized template library</p><h1>Slide designs</h1><p>Every preview is derived locally from the installed PowerPoint master and layout catalog.</p></div><button className="button secondary" disabled={loading} onClick={onInstall}><ArrowsClockwise size={17} />{loading ? "Reading…" : "Install or update"}</button></header>
    <section className="template-pack-bar"><span className="template-pack-mark"><ShieldCheck size={21} /></span><div><span className="field-label">Active local Template Pack</span><strong>{catalog.name}</strong><small>{catalog.masterCount} master{catalog.masterCount === 1 ? "" : "s"} · {catalog.layouts.length} layouts · SHA-256 {catalog.sha256.slice(0, 12)}…{installedAt ? ` · installed ${new Date(installedAt).toLocaleDateString()}` : " · browser session"}</small></div><span className="template-local-state"><span />Local only</span></section>
    <section className="design-feature">
      <div className="design-feature-preview"><TemplateLayoutCanvas catalog={catalog} layout={selected} label={`Selected design: ${selected.name}`} /></div>
      <div className="design-feature-copy"><span className="field-label">Selected design</span><h2>{selected.name}</h2><p>This is a structural preview of the actual template layout. Template media, native geometry, theme colors, and placeholder prompts are read from the installed file.</p><dl><div><dt>Category</dt><dd>{selected.category}</dd></div><div><dt>Content regions</dt><dd>{selected.placeholderTypes.length ? selected.placeholderTypes.join(" · ") : "Freeform / template furniture"}</dd></div><div><dt>PowerPoint structure</dt><dd>Master and layout preserved</dd></div></dl><div className="inline-note design-note"><Info size={16} />Dashed green regions are app-added guides for editable placeholders; they are not template artwork. PowerPoint-native rendering remains the final fidelity check, and selecting a design here changes nothing.</div></div>
    </section>
    <section className="design-browser panel">
      <div className="design-toolbar"><div className="design-filters" aria-label="Filter slide designs">{(["all", "title", "content", "image", "conclusion", "other"] as DesignFilter[]).map((value) => <button key={value} className={filter === value ? "active" : ""} onClick={() => setFilter(value)}>{value === "all" ? `All ${catalog.layouts.length}` : `${value} ${categoryCounts[value] ?? 0}`}</button>)}</div><label className="design-search"><MagnifyingGlass size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search layouts or placeholders" aria-label="Search slide designs" /></label></div>
      <div className="design-grid">{visibleLayouts.map((layout, index) => <button key={layout.id} className={`design-card ${selected.id === layout.id ? "selected" : ""}`} onClick={() => setSelectedId(layout.id)} aria-pressed={selected.id === layout.id}><span className="design-thumb"><TemplateLayoutCanvas catalog={catalog} layout={layout} /></span><span className="design-card-meta"><span>{index + 1}</span><span><strong>{layout.name}</strong><small>{layout.placeholderTypes.length ? layout.placeholderTypes.join(" · ") : layout.category}</small></span></span></button>)}</div>
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

function ReviewView({ deck, projectUpdatedAt, currentCatalog, proposalCatalog, previewLoading, onToggle, onReject, onApply, onExport }: { deck?: DeckJob; projectUpdatedAt: string; currentCatalog?: SlideRenderCatalog; proposalCatalog?: SlideRenderCatalog; previewLoading: boolean; onToggle: (id: string) => void; onReject: () => void; onApply: () => void; onExport: () => void }) {
  const proposal = deck?.proposal;
  const [selectedNumber, setSelectedNumber] = useState(1);
  useEffect(() => { setSelectedNumber(1); }, [deck?.id, proposal?.id]);
  if (!deck || !proposal) return <NoSelection message="Stage a cleanup proposal from Deck audit to review exact changes here." icon={MagicWand} />;
  const stale = proposal.baseUpdatedAt !== projectUpdatedAt && proposal.status === "pending";
  const disposition = proposal.slideDispositions.find((item) => item.slideNumber === selectedNumber);
  const currentSlide = currentCatalog?.slides.find((slide) => slide.number === selectedNumber);
  const proposalSlide = proposalCatalog?.slides.find((slide) => slide.number === selectedNumber);
  const changedCount = proposal.slideDispositions.filter((item) => item.status === "change-proposed" || item.changeIds.length > 0).length;
  const approvedCount = proposal.slideDispositions.filter((item) => item.status === "approved-as-is").length;
  const reviewCount = proposal.slideDispositions.filter((item) => item.status === "needs-review").length;
  const tableCount = proposal.changes.filter((change) => change.kind === "table-style" && change.selected).reduce((sum, change) => sum + (change.tableIds?.length ?? 0), 0);
  return <div className="view-stack"><header className="view-header compact"><div><p className="eyebrow">Before / after design review</p><h1>{proposal.summary}</h1><p>{deck.name} · Proposal {proposal.id.slice(0, 8)} · {proposal.mode.replaceAll("-", " ")}</p></div><span className={`proposal-state ${proposal.status}`}>{proposal.status}</span></header>
    {stale && <div className="warning-banner"><Warning size={18} /><div><strong>This proposal is stale.</strong><span>The project changed after it was staged. Restage before applying.</span></div></div>}
    <div className="metric-strip review-metrics"><Metric value={proposal.slideDispositions.length} label="Slides reviewed" /><Metric value={changedCount} label="With changes" /><Metric value={approvedCount} label="Approved as-is" /><Metric value={reviewCount} label="Need review" /><Metric value={tableCount} label="Tables normalized" /></div>
    <section className="proposal-compare panel"><div className="slide-review-toolbar"><div><span><strong>Slide {selectedNumber}</strong><small className={`disposition-label ${disposition?.status ?? "approved-as-is"}`}>{disposition?.status.replaceAll("-", " ") ?? "reviewed"}</small></span></div><div><button className="button ghost small" disabled={selectedNumber <= 1} onClick={() => setSelectedNumber((value) => value - 1)}><ArrowLeft size={15} />Previous</button><button className="button ghost small" disabled={selectedNumber >= (deck.audit?.slideCount ?? 1)} onClick={() => setSelectedNumber((value) => value + 1)}>Next<ArrowRight size={15} /></button></div></div>
      <div className="proposal-compare-body"><div className="proposal-compare-canvases"><div className="proposal-canvas-pane"><div><strong>Current</strong><small>Embedded source · read only</small></div><span className="proposal-slide-canvas">{currentCatalog && currentSlide ? <TemplateLayoutCanvas catalog={currentCatalog} layout={currentSlide} label={`Current slide ${selectedNumber}`} /> : <span className="proposal-preview-wait"><ArrowsClockwise className="spinner" size={22} />Rendering current design…</span>}</span></div><div className="proposal-canvas-pane proposed"><div><strong>Proposal</strong><small>Selected changes · not yet applied</small></div><span className="proposal-slide-canvas">{proposalCatalog && proposalSlide ? <TemplateLayoutCanvas catalog={proposalCatalog} layout={proposalSlide} label={`Proposed slide ${selectedNumber}`} /> : <span className="proposal-preview-wait">{previewLoading ? <><ArrowsClockwise className="spinner" size={22} />Rendering proposal…</> : <><Info size={22} />Select at least one supported change</>}</span>}</span></div></div>
        <aside className="proposal-slide-rail"><span className="field-label">Deck-wide disposition</span><div className="disposition-reasons">{disposition?.reasons.map((reason) => <p key={reason}>{reason}</p>)}</div><div className="disposition-list">{proposal.slideDispositions.map((item) => <button key={item.slideNumber} className={`${item.status} ${item.slideNumber === selectedNumber ? "selected" : ""}`} onClick={() => setSelectedNumber(item.slideNumber)}><span>{item.slideNumber}</span><small>{item.status === "change-proposed" ? "Changed" : item.status === "needs-review" ? "Review" : "As-is"}</small></button>)}</div></aside></div>
      <div className="slide-representation-note"><ShieldCheck size={15} /><span><strong>Revision-bound comparison</strong> Both views are rendered locally from native OOXML. Export still requires an independent native PowerPoint visual QA pass.</span></div>
    </section>
    {proposal.tableExceptions.length > 0 && <section className="panel proposal-exceptions"><div className="panel-heading"><div><h2>Table design exceptions</h2><p>These tables were preserved rather than forced into a generic treatment.</p></div><span className="quiet-label">{proposal.tableExceptions.length} designer check{proposal.tableExceptions.length === 1 ? "" : "s"}</span></div><div>{proposal.tableExceptions.map((exception) => <article key={exception.tableId}><Warning size={17} /><span><strong>Slide {exception.slideNumber} · {exception.rule.replaceAll("-", " ")}</strong><small>{exception.reason}</small></span></article>)}</div></section>}
    <section className="panel"><div className="panel-heading"><div><h2>Proposed changes</h2><p>Only checked changes enter the accepted cleanup plan.</p></div><span className="quiet-label">Text and table structure: locked</span></div><div className="proposal-list">{proposal.changes.length === 0 && <div className="resource-empty"><CheckCircle size={25} /><span><strong>No deterministic changes needed</strong><small>Every slide still received an explicit disposition above.</small></span></div>}{proposal.changes.map((change) => <label className="proposal-change" key={change.id}><input type="checkbox" checked={change.selected} disabled={proposal.status !== "pending"} onChange={() => onToggle(change.id)} /><span className="change-route"><b>{change.from}</b><CaretRight size={17} /><b>{change.to}</b></span><span>{change.kind === "table-style" ? `${change.tableIds?.length ?? 0} native tables` : change.kind === "alignment" ? `${change.alignmentRepairs?.length ?? 0} text boxes` : `${change.affectedRunCount} markup references`} · {change.affectedSlideNumbers.length} slide locations</span><small>{change.rationale}</small></label>)}</div><div className="review-footer"><div className="lock-copy"><LockKey size={20} /><span><strong>Exact-content guard</strong><small>Export aborts if slide count, visible text, table cell content, or merged-cell topology changes.</small></span></div><div className="review-actions">{proposal.status === "pending" ? <><button className="button ghost" onClick={onReject}><X size={17} />Reject</button><button className="button primary" disabled={stale || !proposal.changes.some((change) => change.selected)} onClick={onApply}><Check size={17} />Accept plan</button></> : proposal.status === "applied" ? <button className="button primary" onClick={onExport}><FileArrowDown size={17} />Export review copy</button> : <span className="muted">Proposal rejected; the source is unchanged.</span>}</div></div></section></div>;
}

function ResourcesView({ project, onToggleMcp, onAdd }: { project: PresentationStudioProject; onToggleMcp: (id: string) => void; onAdd: () => void }) {
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
        <span><strong>Drop files anywhere in Presentation Studio</strong><small>Documents, data, images, audio, video, SVG, and PowerPoint files are processed locally and embedded in this project.</small></span>
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
          <div className="resource-row resource-head"><span>Resource</span><span>Role</span><span>Processing</span><span>Size</span><span>AI session</span></div>
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
            </div>;
          })}
        </div>
      </section>
      <div className="inline-note wide"><ShieldCheck size={18} />Original bytes and extracted text remain local. Metadata sharing requires both AI session access and a per-Resource choice; file bytes are never returned through MCP.</div>
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
  const [mcpActivity, setMcpActivity] = useState<{ id: string; operation: string; state: "active" | "completed" | "failed" }>();
  const [presentationFontCss, setPresentationFontCss] = useState("");
  const [secureAutosavePassword, setSecureAutosavePassword] = useState<string>();
  const [fileDragActive, setFileDragActive] = useState(false);
  const [templateCatalog, setTemplateCatalog] = useState<TemplateCatalog>();
  const [templateInstalledAt, setTemplateInstalledAt] = useState<string>();
  const [templateLoading, setTemplateLoading] = useState(false);
  const [slideCatalogs, setSlideCatalogs] = useState<Record<string, SlideRenderCatalog>>({});
  const [slideCatalogLoadingDeckId, setSlideCatalogLoadingDeckId] = useState<string>();
  const [proposalCatalogs, setProposalCatalogs] = useState<Record<string, SlideRenderCatalog>>({});
  const [proposalCatalogLoadingDeckId, setProposalCatalogLoadingDeckId] = useState<string>();
  const [tourOpen, setTourOpen] = useState(false);
  const [tourStepIndex, setTourStepIndex] = useState(0);
  const fileDragDepth = useRef(0);
  const slideCatalogsRef = useRef(new Map<string, SlideRenderCatalog>());
  const proposalCatalogsRef = useRef(new Map<string, SlideRenderCatalog>());
  const onboardingChecked = useRef(false);
  const desktop = window.presentationStudioDesktop;
  const selectedDeck = project.decks.find((deck) => deck.id === selectedDeckId) ?? project.decks[0];

  useEffect(() => { projectRef.current = project; }, [project]);
  useEffect(() => { if (!selectedDeckId && project.decks[0]) setSelectedDeckId(project.decks[0].id); }, [project.decks, selectedDeckId]);
  useEffect(() => { void desktop?.getMcpStatus().then(setMcpStatus).catch(() => undefined); }, [desktop]);

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
      const catalog = await buildTemplateCatalog(bytesFrom(installed.bytes), installed.name);
      if (!canceled) {
        setTemplateCatalog(catalog);
        setTemplateInstalledAt(installed.installedAt);
      }
    }).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The installed template could not be opened.");
    }).finally(() => { if (!canceled) setTemplateLoading(false); });
    return () => { canceled = true; };
  }, [desktop]);

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
    const materialized = await buildCleanupProposalPptx(source.bytes, deck.proposal);
    const catalog = await buildSlideRenderCatalog(materialized.bytes, `${cleanFileStem(deck.name)}_proposal.pptx`);
    proposalCatalogsRef.current.set(key, catalog);
    setProposalCatalogs((existing) => ({ ...existing, [deck.id]: catalog }));
    return catalog;
  }, []);

  useEffect(() => {
    if (!["slides", "review"].includes(activeView) || !selectedDeck?.audit) return;
    let canceled = false;
    setSlideCatalogLoadingDeckId(selectedDeck.id);
    void getOrBuildSlideCatalog(selectedDeck).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The current slide designs could not be rendered.");
    }).finally(() => { if (!canceled) setSlideCatalogLoadingDeckId((value) => value === selectedDeck.id ? undefined : value); });
    return () => { canceled = true; };
  }, [activeView, getOrBuildSlideCatalog, selectedDeck?.id, selectedDeck?.sourceSha256]);

  const selectedProposalSignature = selectedDeck?.proposal?.changes.filter((change) => change.selected).map((change) => change.id).sort().join("|");
  useEffect(() => {
    if (activeView !== "review" || !selectedDeck?.proposal || !selectedProposalSignature) return;
    let canceled = false;
    setProposalCatalogLoadingDeckId(selectedDeck.id);
    void getOrBuildProposalCatalog(selectedDeck).then(() => {
      if (!canceled) setError(undefined);
    }).catch((caught) => {
      if (!canceled) setError(caught instanceof Error ? caught.message : "The proposal slide designs could not be rendered.");
    }).finally(() => { if (!canceled) setProposalCatalogLoadingDeckId((value) => value === selectedDeck.id ? undefined : value); });
    return () => { canceled = true; };
  }, [activeView, getOrBuildProposalCatalog, selectedDeck?.id, selectedDeck?.proposal?.id, selectedProposalSignature]);

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
      if (request.operation === "get_app_status") return { app: "Presentation Studio", designStandardVersion: current.settings.designStandardVersion, project: { name: current.project.name, type: current.project.type, resourceCount: current.resources.length, deckCount: current.decks.length, slideCount: current.decks.reduce((sum, deck) => sum + (deck.audit?.slideCount ?? 0), 0), submittedDesignThreadCount: current.designThreads.filter((thread) => thread.status === "submitted").length, updatedAt: current.project.updatedAt }, aiSessionAccess: mcpEnabled };
      if (!mcpEnabled) throw new Error("Enable AI session access in Presentation Studio before reading project metadata or staging work.");
      if (request.operation === "list_decks") return { updatedAt: current.project.updatedAt, decks: current.decks.map((deck) => ({ id: deck.id, name: deck.name, status: deck.status, operationScope: deck.operationScope, templateClassification: deck.templateClassification, targetTemplateId: deck.targetTemplateId, slideCount: deck.audit?.slideCount ?? 0, findingCount: deck.audit?.findings.length ?? 0 })) };
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
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, status: deck.status, targetTemplateId: deck.targetTemplateId }, audit: { ...deck.audit, slides: deck.audit.slides.map(({ text: _text, title: _title, ...slide }) => slide), pictures: deck.audit.pictures.map(({ name: _name, description: _description, relationshipId: _relationshipId, ...picture }) => picture) } };
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
        }));
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId }, range: { start, end }, slides, preservationRequired: true };
      }
      if (request.operation === "get_cleanup_rule_profile") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId, targetTemplateDecisionSource: deck.targetTemplateDecisionSource }, resolvedProfile: deck.designProfile ?? null, standard: PRESENTATION_DESIGN_STANDARD, routineApprovalsRequired: false };
      }
      if (request.operation === "get_slide_render") {
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck?.audit) throw new Error("The requested deck is not open or audited.");
        const slideNumber = Number(request.input.slideNumber);
        if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
        const representation = request.input.representation === "proposal" ? "proposal" : "current";
        const catalog = representation === "proposal" ? await getOrBuildProposalCatalog(deck, current) : await getOrBuildSlideCatalog(deck, current);
        const slide = catalog.slides.find((item) => item.number === slideNumber);
        if (!slide) throw new Error("The requested slide render is unavailable.");
        const jpeg = await slidePreviewJpeg(catalog, slide, 1200, presentationFontCss);
        return { updatedAt: current.project.updatedAt, deck: { id: deck.id, name: deck.name }, slide: { id: deck.audit.slides.find((item) => item.number === slideNumber)?.id ?? slide.id, number: slide.number, title: slide.title }, representation, proposalId: representation === "proposal" ? deck.proposal?.id : undefined, renderer: catalog.renderer, authoritative: false, qaNote: "This local OOXML preview supports design comparison; the exported native PPTX and an independent PowerPoint render remain the final fidelity authority.", mimeType: "image/jpeg", ...jpeg };
      }
      if (request.operation === "list_design_threads") {
        const deckId = typeof request.input.deckId === "string" ? request.input.deckId : undefined;
        const threads = current.designThreads.filter((thread) => (!deckId || thread.deckId === deckId) && ["submitted", "proposal-ready", "needs-reanchor"].includes(thread.status));
        return { updatedAt: current.project.updatedAt, threads: threads.map((thread) => ({ id: thread.id, deckId: thread.deckId, slideId: thread.slideId, slideNumber: thread.slideNumber, baseRevision: thread.baseRevision, anchor: thread.anchor, comment: thread.comment, status: thread.status, createdAt: thread.createdAt, submittedAt: thread.submittedAt })) };
      }
      if (request.operation === "get_design_thread") {
        const thread = current.designThreads.find((item) => item.id === request.input.threadId);
        if (!thread || thread.status === "note") throw new Error("The requested design thread is not submitted to AI in this project.");
        const deck = current.decks.find((item) => item.id === thread.deckId);
        const slide = deck?.audit?.slides.find((item) => item.id === thread.slideId || item.number === thread.slideNumber);
        return { updatedAt: current.project.updatedAt, thread, deck: deck ? { id: deck.id, name: deck.name, targetTemplateId: deck.targetTemplateId } : null, slide: slide ? { id: slide.id, number: slide.number, title: slide.title, textHash: slide.textHash, objects: { tables: slide.tableCount, pictures: slide.pictureCount, charts: slide.chartCount } } : null, instruction: "Read the current revision and get_slide_render before staging a bounded fix. Do not guess if the anchor no longer maps unambiguously." };
      }
      if (request.operation === "stage_font_cleanup") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck) throw new Error("The requested deck is not open.");
        const proposal = createFontCleanupProposal(deck, current.project.updatedAt);
        const next = {
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
          activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action: "mcp-proposal-staged", detail: `AI staged font cleanup for ${deck.name}; no changes were applied.` }],
        };
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, changes: proposal.changes }, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
      }
      if (request.operation === "stage_designer_cleanup") {
        if (request.input.expectedUpdatedAt !== current.project.updatedAt) throw new Error("The project changed. Read the deck list again before staging a proposal.");
        const deck = current.decks.find((item) => item.id === request.input.deckId);
        if (!deck) throw new Error("The requested deck is not open.");
        const proposal = createDesignerCleanupProposal(deck, current.project.updatedAt);
        const next = {
          ...current,
          decks: current.decks.map((item) => item.id === deck.id ? { ...item, proposal, status: "proposal-ready" as const } : item),
          activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action: "mcp-designer-proposal-staged", detail: `AI staged a deck-wide designer cleanup for ${deck.name}; no changes were applied.` }],
        };
        setProject(next);
        setSelectedDeckId(deck.id);
        setActiveView("review");
        return { proposal: { id: proposal.id, summary: proposal.summary, status: proposal.status, mode: proposal.mode, changes: proposal.changes, slideDispositions: proposal.slideDispositions, tableExceptions: proposal.tableExceptions }, projectUpdatedAt: next.project.updatedAt, applied: false, saved: false };
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
  }, [desktop, getOrBuildProposalCatalog, getOrBuildSlideCatalog, mcpEnabled, presentationFontCss]);

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
              importedDecks.push({
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
              });
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
    if (event.dataTransfer.files.length > 0) void importWebFiles(event.dataTransfer.files, "resources");
  }

  function confirmTemplate(templateId: string) {
    if (!selectedDeck) return;
    const now = new Date().toISOString();
    setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, targetTemplateId: templateId, targetTemplateConfirmedAt: now, targetTemplateDecisionSource: "user-selected", designProfile: templateId === PRESENTATION_DESIGN_STANDARD.defaults.template.id ? createOrnlDesignProfile("user-selected", now) : undefined, status: templateId === "ornl-16x9-v1" ? "ready-for-cleanup" : "audited", proposal: undefined } : deck) }, "target-template-confirmed", `Confirmed ${templateId} for ${selectedDeck.name}.`));
    setNotice("Target template confirmed. No slide content was changed.");
  }

  function startOrnlCleanup() {
    if (!selectedDeck?.audit) return;
    clearMessages();
    const current = projectRef.current;
    const sourceDeck = current.decks.find((deck) => deck.id === selectedDeck.id);
    if (!sourceDeck) return;
    const now = new Date().toISOString();
    const adoptedDeck: DeckJob = { ...sourceDeck, operationScope: "reflow", targetTemplateId: PRESENTATION_DESIGN_STANDARD.defaults.template.id, targetTemplateConfirmedAt: now, targetTemplateDecisionSource: sourceDeck.designProfile ? sourceDeck.targetTemplateDecisionSource ?? "automatic-default" : "user-selected", designProfile: sourceDeck.designProfile ?? createOrnlDesignProfile("user-selected", now), status: "ready-for-cleanup", proposal: undefined };
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
    setProject((current) => touchProject({ ...current, designThreads: [...current.designThreads, thread] }, submit ? "design-thread-submitted" : "design-note-saved", `${submit ? "Submitted" : "Saved"} a location-anchored design comment on slide ${slide.number} of ${selectedDeck.name}.`));
    setNotice(submit ? `Design comment submitted for AI on slide ${slide.number}. It is anchored to the selected region and no change was applied.` : `Private design note saved on slide ${slide.number}. It is not available through MCP.`);
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
  function rejectProposal() { updateProposal((proposal) => ({ ...proposal, status: "rejected" }), "cleanup-proposal-rejected"); setProject((current) => ({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck?.id ? { ...deck, status: "audited" } : deck) })); }
  function acceptProposal() {
    if (!selectedDeck?.proposal || selectedDeck.proposal.baseUpdatedAt !== project.project.updatedAt) { setError("The proposal is stale. Restage it against the current project."); return; }
    setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id && deck.proposal ? { ...deck, proposal: { ...deck.proposal, status: "applied" }, status: "approved" } : deck) }, "cleanup-plan-accepted", `Accepted the cleanup plan for ${selectedDeck.name}; source bytes remain unchanged until a new copy is exported.`));
  }

  async function exportCleaned() {
    if (!selectedDeck?.proposal || selectedDeck.proposal.status !== "applied") return;
    const resource = sourceForDeck(project, selectedDeck);
    if (!resource?.bytes) { setError("The embedded source deck is unavailable."); return; }
    setBusy("Building and validating a new cleaned PowerPoint copy…");
    clearMessages();
    try {
      const output = await applyCleanupToPptx(resource.bytes, selectedDeck.proposal);
      if (!desktop) throw new Error("PowerPoint export requires the Electron desktop app.");
      const result = await desktop.saveBinary({ kind: "pptx", defaultName: `${cleanFileStem(selectedDeck.name)}_designer-cleaned.pptx`, bytes: output.bytes });
      if (!result.canceled) {
        setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, status: "needs-manual-review", exportedAt: new Date().toISOString() } : deck) }, "cleaned-review-copy-exported", `Exported a new review copy of ${selectedDeck.name} with ${output.replacementCount} font references, ${output.alignmentCount} alignments, and ${output.tableCount} native tables updated.`));
        setNotice(`Review copy exported with ${output.replacementCount} font references, ${output.alignmentCount} text alignments, and ${output.tableCount} native tables updated. Exact visible text, table content, and merged-cell structure passed validation; complete native visual QA in PowerPoint.`);
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
      const extension = encrypted ? ".pstudio-secure" : ".pstudio";
      const result = await desktop.saveBinary({ kind, defaultName: `${project.project.name.replace(/[^a-z0-9 _-]+/gi, "").trim() || "Presentation Studio project"}${extension}`, bytes });
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
    clearMessages(); setBusy("Opening and validating project package…");
    try {
      let bytes = bytesFrom(result.file.bytes);
      let password: string | undefined;
      if (isEncryptedProject(bytes)) {
        password = window.prompt("Enter the password for this encrypted Presentation Studio project.") ?? undefined;
        if (!password) return;
        bytes = await decryptProjectPackage(bytes, password);
      }
      const opened = await openProjectPackage(bytes);
      setProject(opened); setSelectedDeckId(opened.decks[0]?.id); setSecureAutosavePassword(password); setActiveView("batch"); setNotice(`Opened ${opened.project.name}; all embedded resource hashes passed validation.`);
    } catch (caught) { setError(caught instanceof Error ? caught.message : "The project could not be opened."); }
    finally { setBusy(undefined); }
  }

  const mainContent = useMemo(() => {
    if (activeView === "batch") return <BatchView project={project} selectedId={selectedDeck?.id} onSelect={(id) => { setSelectedDeckId(id); setActiveView("decks"); }} onAdd={() => void addDecks()} />;
    if (activeView === "decks") return <DeckAuditView deck={selectedDeck} onConfirm={confirmTemplate} onStage={stageCleanup} onStartOrnlCleanup={startOrnlCleanup} onMarkExemplar={markTableExemplar} onExportReport={() => void exportAuditReport()} isExemplar={Boolean(selectedDeck && project.styleExemplars.some((item) => item.deckId === selectedDeck.id && item.kind === "table"))} />;
    if (activeView === "slides") return <SlidesView deck={selectedDeck} catalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} loading={Boolean(selectedDeck && slideCatalogLoadingDeckId === selectedDeck.id)} revision={project.project.updatedAt} threads={project.designThreads} onSaveThread={saveDesignThread} />;
    if (activeView === "designs") return <DesignsView catalog={templateCatalog} installedAt={templateInstalledAt} loading={templateLoading} onInstall={() => void installTemplate()} />;
    if (activeView === "rules") return <RulesView deck={selectedDeck} exemplarCount={project.styleExemplars.filter((item) => item.kind === "table").length} />;
    if (activeView === "review") return <ReviewView deck={selectedDeck} projectUpdatedAt={project.project.updatedAt} currentCatalog={selectedDeck ? slideCatalogs[selectedDeck.id] : undefined} proposalCatalog={selectedDeck ? proposalCatalogs[selectedDeck.id] : undefined} previewLoading={Boolean(selectedDeck && proposalCatalogLoadingDeckId === selectedDeck.id)} onToggle={toggleChange} onReject={rejectProposal} onApply={acceptProposal} onExport={() => void exportCleaned()} />;
    return <ResourcesView project={project} onAdd={() => void addResources()} onToggleMcp={(id) => setProject((current) => touchProject({ ...current, resources: current.resources.map((resource) => resource.id === id ? { ...resource, mcpAccess: resource.mcpAccess === "none" ? "metadata" : "none" } : resource) }, "resource-access-updated", "Updated session-only Resource metadata permission."))} />;
  }, [activeView, project, proposalCatalogLoadingDeckId, proposalCatalogs, selectedDeck, slideCatalogLoadingDeckId, slideCatalogs, templateCatalog, templateInstalledAt, templateLoading]);

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
        </div>
      </nav>
      <main className="workspace">{mainContent}</main>
      <Inspector deck={selectedDeck} onOpenReview={() => setActiveView("review")} />
      <OnboardingTour open={tourOpen} stepIndex={tourStepIndex} onStepChange={setTourStepIndex} onClose={closeOnboardingTour} />
      {mcpActivity && <div className={`mcp-activity ${mcpActivity.state}`} role="status" aria-live="polite"><span>{mcpActivity.state === "active" ? <ArrowsClockwise className="spinner" size={17} /> : mcpActivity.state === "completed" ? <CheckCircle size={17} /> : <Warning size={17} />}</span><div><strong>{mcpActivity.state === "active" ? "AI is using Presentation Studio" : mcpActivity.state === "completed" ? "AI operation completed" : "AI operation needs attention"}</strong><small>{mcpActivity.operation.replaceAll("_", " ")} · {mcpActivity.state === "active" ? "local operation in progress" : mcpActivity.state}</small></div></div>}
      {(notice || error) && <div className={`toast ${error ? "error" : "success"}`}><span>{error ? <Warning size={18} /> : <CheckCircle size={18} />}</span><p>{error ?? notice}</p><button onClick={clearMessages}><X size={16} /></button></div>}
      {fileDragActive && <div className="file-drop-overlay"><div className="file-drop-card"><UploadSimple size={38} /><strong>Drop to add project Resources</strong><span>Files will be processed locally, embedded by hash, and preserved with this project.</span></div></div>}
      {busy && <div className="busy-overlay"><div className="busy-card"><ArrowsClockwise className="spinner" size={25} /><strong>{busy}</strong><span>Files are processed locally and copied into the project. External originals remain untouched.</span></div></div>}
      <input id="web-deck-picker" hidden type="file" accept=".pptx" multiple onChange={(event) => { void importWebFiles(event.target.files, "decks"); event.currentTarget.value = ""; }} />
      <input id="web-resource-picker" hidden type="file" accept=".pptx,.potx,.docx,.pdf,.md,.markdown,.txt,.csv,.tsv,.json,.xlsx,.png,.jpg,.jpeg,.webp,.tif,.tiff,.svg,.wav,.mp3,.m4a,.mp4,.mov,.doc,.xls" multiple onChange={(event) => { void importWebFiles(event.target.files, "resources"); event.currentTarget.value = ""; }} />
      <input id="web-template-picker" hidden type="file" accept=".pptx,.potx" onChange={(event) => { const file = event.target.files?.[0]; if (file) void file.arrayBuffer().then((buffer) => loadTemplate({ name: file.name, filePath: file.name, mediaType: file.type, bytes: new Uint8Array(buffer) }, false)); event.currentTarget.value = ""; }} />
    </div>
  );
}
