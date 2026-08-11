import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent } from "react";
import {
  Archive,
  ArrowsClockwise,
  CaretRight,
  Check,
  CheckCircle,
  CirclesThreePlus,
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
  Monitor,
  PresentationChart,
  ShieldCheck,
  Slideshow,
  Sparkle,
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
  PresentationStudioProject,
  ProjectResource,
  ResourceKind,
  TemplateClassification,
} from "./types";
import { applyCleanupToPptx, createFontCleanupProposal } from "./lib/cleanup";
import type { PickedBinaryFile } from "./lib/desktop";
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
import {
  ONBOARDING_TOUR_STORAGE_KEY,
  ONBOARDING_TOUR_VERSION,
  shouldShowOnboardingTour,
} from "./lib/onboarding";

type ViewId = "batch" | "decks" | "slides" | "rules" | "review" | "resources";

const navItems: Array<{ id: ViewId; label: string; icon: Icon }> = [
  { id: "batch", label: "Batch", icon: Files },
  { id: "decks", label: "Deck audit", icon: PresentationChart },
  { id: "slides", label: "Slides", icon: Slideshow },
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

function DeckAuditView({ deck, onConfirm, onStage, onMarkExemplar, onExportReport, isExemplar }: { deck?: DeckJob; onConfirm: (templateId: string) => void; onStage: () => void; onMarkExemplar: () => void; onExportReport: () => void; isExemplar: boolean }) {
  if (!deck?.audit) return <NoSelection message="Select an audited deck from the Batch workspace." />;
  const audit = deck.audit;
  return (
    <div className="view-stack">
      <header className="view-header compact"><div><p className="eyebrow">Read-only deck audit</p><h1>{deck.name}</h1><p>Scanned directly from PowerPoint package structure; no source bytes were changed.</p></div><div className="header-actions"><button className="button ghost small" onClick={onExportReport}><FileArrowDown size={16} />Audit report</button><StatusPill status={deck.status} /></div></header>
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
          {deck.targetTemplateConfirmedAt && <span className="confirmation"><Check size={14} />Confirmed by user</span>}
        </div>
      </section>
      <div className="metric-strip six">
        <Metric value={audit.slideCount} label="Slides" /><Metric value={audit.layoutCount} label="Layouts" /><Metric value={audit.masterCount} label="Masters" /><Metric value={audit.tableCount} label="Tables" /><Metric value={audit.pictureCount} label="Pictures" /><Metric value={audit.notesCount} label="Notes" />
      </div>
      <div className="split-grid">
        <section className="panel">
          <div className="panel-heading"><div><h2>Font inventory</h2><p>Direct slide use is separated from layout, master, and theme fallbacks.</p></div><button className="button secondary small" disabled={!deck.targetTemplateConfirmedAt || deck.targetTemplateId !== "ornl-16x9-v1"} onClick={onStage}><MagicWand size={16} />Stage font cleanup</button></div>
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

function SlidesView({ deck }: { deck?: DeckJob }) {
  if (!deck?.audit) return <NoSelection message="Select a deck to inspect its structural slide inventory." />;
  return <div className="view-stack"><header className="view-header compact"><div><p className="eyebrow">Structural canvas</p><h1>{deck.name}</h1><p>Text and object inventory from the editable PowerPoint source. Visual render comparison is the next renderer layer.</p></div></header><div className="slide-grid">{deck.audit.slides.map((slide) => <article className="slide-card" key={slide.id}><div className="slide-canvas"><div className="slide-title-line" /><div className="slide-text-lines"><i /><i /><i /></div><div className="object-chips">{slide.tableCount > 0 && <span><Table size={13} />{slide.tableCount}</span>}{slide.pictureCount > 0 && <span><Images size={13} />{slide.pictureCount}</span>}</div></div><div className="slide-meta"><span>{slide.number}</span><div><strong>{slide.title}</strong><small>{slide.textRunCount} text runs · {slide.fonts.length} fonts</small></div></div></article>)}</div></div>;
}

function RulesView({ deck, exemplarCount }: { deck?: DeckJob; exemplarCount: number }) {
  const ornlReady = deck?.targetTemplateId === "ornl-16x9-v1" && Boolean(deck.targetTemplateConfirmedAt);
  return <div className="view-stack"><header className="view-header compact"><div><p className="eyebrow">Cleanup rule profile</p><h1>Minimum necessary change</h1><p>Rules are scoped, deterministic, and inactive until the correct template is confirmed.</p></div></header><div className="rule-grid"><RuleCard title="Legacy font normalization" scope="Font family only" status={ornlReady ? "Ready" : "Waiting for template"} detail="Maps direct Century Gothic and Arial slide markup to Aptos. Symbol fonts, theme tokens, text strings, object identity, and geometry remain untouched." /><RuleCard title="Exact-content guard" scope="Every exported slide" status="Always on" detail="Compares the visible-text hash and slide count before and after cleanup. The export is rejected if either changes." /><RuleCard title="Advanced content hold" scope="Deck safety" status="Always on" detail="Prevents automated cleanup when macros, embedded OLE objects, or external relationships require a human decision." /><RuleCard title="Approved table exemplar" scope="Tables only" status={exemplarCount > 0 ? `${exemplarCount} registered` : "Needs exemplar"} detail="Pins a user-designated table by embedded source hash, slide, and object ordinal without copying its content. Style extraction and normalization are the next rule layer." /></div></div>;
}

function RuleCard({ title, scope, status, detail }: { title: string; scope: string; status: string; detail: string }) {
  return <article className="rule-card"><div className="rule-top"><ListChecks size={22} /><span>{status}</span></div><h2>{title}</h2><p>{detail}</p><small>{scope}</small></article>;
}

function ReviewView({ deck, projectUpdatedAt, onToggle, onReject, onApply, onExport }: { deck?: DeckJob; projectUpdatedAt: string; onToggle: (id: string) => void; onReject: () => void; onApply: () => void; onExport: () => void }) {
  const proposal = deck?.proposal;
  if (!deck || !proposal) return <NoSelection message="Stage a cleanup proposal from Deck audit to review exact changes here." icon={MagicWand} />;
  const stale = proposal.baseUpdatedAt !== projectUpdatedAt && proposal.status === "pending";
  return <div className="view-stack"><header className="view-header compact"><div><p className="eyebrow">Human review gate</p><h1>{proposal.summary}</h1><p>{deck.name} · Proposal {proposal.id.slice(0, 8)}</p></div><span className={`proposal-state ${proposal.status}`}>{proposal.status}</span></header>{stale && <div className="warning-banner"><Warning size={18} /><div><strong>This proposal is stale.</strong><span>The project changed after it was staged. Restage before applying.</span></div></div>}<section className="panel"><div className="panel-heading"><div><h2>Proposed changes</h2><p>Only checked changes enter the accepted cleanup plan.</p></div><span className="quiet-label">Visible text: locked</span></div><div className="proposal-list">{proposal.changes.map((change) => <label className="proposal-change" key={change.id}><input type="checkbox" checked={change.selected} disabled={proposal.status !== "pending"} onChange={() => onToggle(change.id)} /><span className="change-route"><b>{change.from}</b><CaretRight size={17} /><b>{change.to}</b></span><span>{change.affectedRunCount} markup references · {change.affectedSlideNumbers.length || "master/layout"} slide locations</span><small>{change.rationale}</small></label>)}</div><div className="review-footer"><div className="lock-copy"><LockKey size={20} /><span><strong>Content lock</strong><small>Export aborts on any visible-text or slide-count change. Native visual comparison is not yet available.</small></span></div><div className="review-actions">{proposal.status === "pending" ? <><button className="button ghost" onClick={onReject}><X size={17} />Reject</button><button className="button primary" disabled={stale || !proposal.changes.some((change) => change.selected)} onClick={onApply}><Check size={17} />Accept plan</button></> : proposal.status === "applied" ? <button className="button primary" onClick={onExport}><FileArrowDown size={17} />Export review copy</button> : <span className="muted">Proposal rejected; the source is unchanged.</span>}</div></div></section></div>;
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
  const [secureAutosavePassword, setSecureAutosavePassword] = useState<string>();
  const [fileDragActive, setFileDragActive] = useState(false);
  const [tourOpen, setTourOpen] = useState(false);
  const [tourStepIndex, setTourStepIndex] = useState(0);
  const fileDragDepth = useRef(0);
  const onboardingChecked = useRef(false);
  const desktop = window.presentationStudioDesktop;
  const selectedDeck = project.decks.find((deck) => deck.id === selectedDeckId) ?? project.decks[0];

  useEffect(() => { projectRef.current = project; }, [project]);
  useEffect(() => { if (!selectedDeckId && project.decks[0]) setSelectedDeckId(project.decks[0].id); }, [project.decks, selectedDeckId]);
  useEffect(() => { void desktop?.getMcpStatus().then(setMcpStatus).catch(() => undefined); }, [desktop]);

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
    return desktop.onMcpCommand((request) => {
      const current = projectRef.current;
      if (request.operation === "get_app_status") return { app: "Presentation Studio", project: { name: current.project.name, type: current.project.type, resourceCount: current.resources.length, deckCount: current.decks.length, slideCount: current.decks.reduce((sum, deck) => sum + (deck.audit?.slideCount ?? 0), 0), updatedAt: current.project.updatedAt }, aiSessionAccess: mcpEnabled };
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
      throw new Error(`Unknown Presentation Studio MCP operation: ${request.operation}`);
    });
  }, [desktop, mcpEnabled]);

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
              importedDecks.push({ id: crypto.randomUUID(), name: files[index].name, sourceResourceId: resource.id, sourceSha256: resource.sha256, operationScope: "cleanup-only", templateClassification: audit.classification, status: audit.classification === "current-ornl" ? "ready-for-cleanup" : "needs-template-decision", audit, protectedSlideNumbers: [] });
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
    setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, targetTemplateId: templateId, targetTemplateConfirmedAt: now, status: templateId === "ornl-16x9-v1" ? "ready-for-cleanup" : "audited", proposal: undefined } : deck) }, "target-template-confirmed", `Confirmed ${templateId} for ${selectedDeck.name}.`));
    setNotice("Target template confirmed. No slide content was changed.");
  }

  function stageCleanup() {
    if (!selectedDeck) return;
    clearMessages();
    try {
      const proposal = createFontCleanupProposal(selectedDeck, project.project.updatedAt);
      setProject((current) => ({
        ...current,
        decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, proposal, status: "proposal-ready" } : deck),
        activity: [...current.activity, { id: crypto.randomUUID(), at: new Date().toISOString(), action: "cleanup-proposal-staged", detail: `Staged font cleanup for ${selectedDeck.name}; nothing was applied.` }],
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
    if (!window.confirm("This build verifies slide count and visible text, but it does not yet compare native before/after renders. Export a draft copy for manual visual review?")) return;
    const resource = sourceForDeck(project, selectedDeck);
    if (!resource?.bytes) { setError("The embedded source deck is unavailable."); return; }
    setBusy("Building and validating a new cleaned PowerPoint copy…");
    clearMessages();
    try {
      const output = await applyCleanupToPptx(resource.bytes, selectedDeck.proposal);
      if (!desktop) throw new Error("PowerPoint export requires the Electron desktop app.");
      const result = await desktop.saveBinary({ kind: "pptx", defaultName: `${cleanFileStem(selectedDeck.name)}_cleaned.pptx`, bytes: output.bytes });
      if (!result.canceled) {
        setProject((current) => touchProject({ ...current, decks: current.decks.map((deck) => deck.id === selectedDeck.id ? { ...deck, status: "needs-manual-review", exportedAt: new Date().toISOString() } : deck) }, "cleaned-review-copy-exported", `Exported a new review copy of ${selectedDeck.name} with ${output.replacementCount} font references updated.`));
        setNotice(`Review copy exported with ${output.replacementCount} font references updated and all visible-text hashes preserved. Complete visual QA in PowerPoint.`);
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
    if (activeView === "decks") return <DeckAuditView deck={selectedDeck} onConfirm={confirmTemplate} onStage={stageCleanup} onMarkExemplar={markTableExemplar} onExportReport={() => void exportAuditReport()} isExemplar={Boolean(selectedDeck && project.styleExemplars.some((item) => item.deckId === selectedDeck.id && item.kind === "table"))} />;
    if (activeView === "slides") return <SlidesView deck={selectedDeck} />;
    if (activeView === "rules") return <RulesView deck={selectedDeck} exemplarCount={project.styleExemplars.filter((item) => item.kind === "table").length} />;
    if (activeView === "review") return <ReviewView deck={selectedDeck} projectUpdatedAt={project.project.updatedAt} onToggle={toggleChange} onReject={rejectProposal} onApply={acceptProposal} onExport={() => void exportCleaned()} />;
    return <ResourcesView project={project} onAdd={() => void addResources()} onToggleMcp={(id) => setProject((current) => touchProject({ ...current, resources: current.resources.map((resource) => resource.id === id ? { ...resource, mcpAccess: resource.mcpAccess === "none" ? "metadata" : "none" } : resource) }, "resource-access-updated", "Updated session-only Resource metadata permission."))} />;
  }, [activeView, project, selectedDeck]);

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
      {(notice || error) && <div className={`toast ${error ? "error" : "success"}`}><span>{error ? <Warning size={18} /> : <CheckCircle size={18} />}</span><p>{error ?? notice}</p><button onClick={clearMessages}><X size={16} /></button></div>}
      {fileDragActive && <div className="file-drop-overlay"><div className="file-drop-card"><UploadSimple size={38} /><strong>Drop to add project Resources</strong><span>Files will be processed locally, embedded by hash, and preserved with this project.</span></div></div>}
      {busy && <div className="busy-overlay"><div className="busy-card"><ArrowsClockwise className="spinner" size={25} /><strong>{busy}</strong><span>Files are processed locally and copied into the project. External originals remain untouched.</span></div></div>}
      <input id="web-deck-picker" hidden type="file" accept=".pptx" multiple onChange={(event) => { void importWebFiles(event.target.files, "decks"); event.currentTarget.value = ""; }} />
      <input id="web-resource-picker" hidden type="file" accept=".pptx,.potx,.docx,.pdf,.md,.markdown,.txt,.csv,.tsv,.json,.xlsx,.png,.jpg,.jpeg,.webp,.tif,.tiff,.svg,.wav,.mp3,.m4a,.mp4,.mov,.doc,.xls" multiple onChange={(event) => { void importWebFiles(event.target.files, "resources"); event.currentTarget.value = ""; }} />
    </div>
  );
}
