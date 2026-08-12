export const PROJECT_SCHEMA = "presentation-studio/project" as const;
export const PROJECT_SCHEMA_VERSION = 1 as const;

export type OperationScope = "audit-only" | "cleanup-only" | "reflow" | "hybrid" | "compose";
export type TemplateClassification =
  | "current-ornl"
  | "older-or-modified-ornl"
  | "sponsor"
  | "custom"
  | "mixed"
  | "unknown";
export type DeckStatus =
  | "not-scanned"
  | "audited"
  | "needs-template-decision"
  | "ready-for-cleanup"
  | "proposal-ready"
  | "needs-manual-review"
  | "approved"
  | "exported"
  | "failed";
export type Severity = "info" | "warning" | "error";
export type ResourceRole =
  | "import-origin"
  | "prior-approved-revision"
  | "style-exemplar"
  | "grounding-source"
  | "slide-media"
  | "chart-data"
  | "template-source"
  | "reference-only";
export type ResourceKind = "presentation" | "document" | "data" | "image" | "audio" | "video" | "other";
export type ResourceProcessingStatus = "indexed" | "stored-only" | "needs-review";
export type ResourceSupportState = "source-readable" | "previewable" | "placeable" | "pptx-preserved" | "unsupported";
export type TemplateDecisionSource = "automatic-default" | "automatic-source-preservation" | "user-selected";
export type DesignThreadStatus = "note" | "submitted" | "proposal-ready" | "resolved" | "needs-reanchor";

export interface ResolvedDesignProfile {
  id: string;
  standardVersion: string;
  templateId: string;
  slideSize: "16:9";
  fontFamily: "Aptos";
  contentPolicy: "preserve-exact";
  adoptedAt: string;
  source: TemplateDecisionSource;
  customized: boolean;
}

export interface DesignThread {
  id: string;
  deckId: string;
  slideId: string;
  slideNumber: number;
  baseRevision: string;
  anchor: {
    kind: "region";
    x: number;
    y: number;
    width: number;
    height: number;
  };
  comment: string;
  status: DesignThreadStatus;
  createdAt: string;
  updatedAt: string;
  submittedAt?: string;
}

export interface ResourceDerivative {
  id: string;
  kind: "extracted-text";
  mediaType: "text/plain";
  byteLength: number;
  sha256: string;
  createdAt: string;
  processor: string;
  truncated: boolean;
  bytes?: Uint8Array;
}

export interface ProjectResource {
  id: string;
  name: string;
  mediaType: string;
  byteLength: number;
  sha256: string;
  roles: ResourceRole[];
  kind?: ResourceKind;
  support?: ResourceSupportState[];
  processing?: {
    status: ResourceProcessingStatus;
    summary: string;
    processedAt: string;
    warnings: string[];
  };
  derivatives?: ResourceDerivative[];
  createdAt: string;
  sourcePath?: string;
  embedded: true;
  bytes?: Uint8Array;
  mcpAccess: "none" | "metadata" | "text" | "preview";
}

export interface FontInventoryItem {
  family: string;
  normalizedFamily: string;
  count: number;
  directSlideCount: number;
  slideNumbers: number[];
  partKinds: string[];
  isThemeFont: boolean;
  isLikelySymbolFont: boolean;
}

export interface SlideInventoryItem {
  id: string;
  number: number;
  title: string;
  text: string;
  textHash: string;
  textRunCount: number;
  tableCount: number;
  pictureCount: number;
  chartCount: number;
  connectorCount: number;
  commentCount: number;
  fonts: string[];
  fontSizes: number[];
  warnings: string[];
}

export interface TableInventoryItem {
  id: string;
  slideNumber: number;
  ordinal: number;
  rowCount: number;
  columnCount: number;
  mergedCellCount: number;
  totalCellCharacterCount: number;
  maximumCellCharacterCount: number;
  styleId?: string;
  styleFlags: string[];
  cellFonts: string[];
  colorTokens: string[];
  marginSignatures: string[];
  styleFingerprint: string;
  contentHash: string;
  structureHash: string;
}

export interface PictureInventoryItem {
  id: string;
  slideNumber: number;
  ordinal: number;
  name: string;
  description?: string;
  relationshipId?: string;
  widthEmu?: number;
  heightEmu?: number;
  cropped: boolean;
  hasOutline: boolean;
  hasEffect: boolean;
}

export interface AlignmentRepairCandidate {
  id: string;
  slideNumber: number;
  shapeId: string;
  textHash: string;
  source: { x: number; y: number; width: number; height: number };
  target: { x: number; y: number; width: number; height: number };
  ruleId: "cover.dominant-left-edge";
  confidence: "high";
  rationale: string;
}

export interface AuditFinding {
  id: string;
  ruleId: string;
  category: "template" | "font" | "table" | "figure" | "layout" | "production" | "technical-review";
  severity: Severity;
  confidence: "high" | "medium" | "low";
  slideNumber?: number;
  message: string;
  evidence: string;
  autoFixable: boolean;
}

export interface PptxAudit {
  scannedAt: string;
  supportLevel: "native-ooxml" | "partial" | "blocked";
  slideCount: number;
  masterCount: number;
  layoutCount: number;
  themeCount: number;
  notesCount: number;
  legacyCommentCount: number;
  modernCommentCount: number;
  mediaCount: number;
  tableCount: number;
  chartCount: number;
  pictureCount: number;
  containsMacros: boolean;
  containsOleObjects: boolean;
  containsExternalRelationships: boolean;
  packageFileCount: number;
  expandedByteLength: number;
  classification: TemplateClassification;
  classificationEvidence: string[];
  fonts: FontInventoryItem[];
  slides: SlideInventoryItem[];
  tables: TableInventoryItem[];
  pictures: PictureInventoryItem[];
  alignmentRepairs: AlignmentRepairCandidate[];
  findings: AuditFinding[];
  warnings: string[];
}

export interface CleanupChange {
  id: string;
  kind: "font-family" | "table-style" | "alignment";
  from: string;
  to: string;
  affectedSlideNumbers: number[];
  affectedRunCount: number;
  tableIds?: string[];
  profileId?: string;
  alignmentRepairs?: AlignmentRepairCandidate[];
  rationale: string;
  selected: boolean;
}

export interface SlideDesignDisposition {
  slideNumber: number;
  status: "change-proposed" | "approved-as-is" | "needs-review";
  reasons: string[];
  changeIds: string[];
}

export interface TableNormalizationException {
  tableId: string;
  slideNumber: number;
  reason: string;
  rule: "semantic-color" | "complex-structure" | "dense-table";
}

export interface CleanupProposal {
  id: string;
  deckId: string;
  baseUpdatedAt: string;
  createdAt: string;
  summary: string;
  status: "pending" | "applied" | "rejected";
  mode: "font-cleanup" | "designer-cleanup";
  standardVersion?: string;
  changes: CleanupChange[];
  slideDispositions: SlideDesignDisposition[];
  tableExceptions: TableNormalizationException[];
}

export interface DeckJob {
  id: string;
  name: string;
  sourceResourceId: string;
  sourceSha256: string;
  operationScope: OperationScope;
  templateClassification: TemplateClassification;
  targetTemplateId?: string;
  targetTemplateConfirmedAt?: string;
  targetTemplateDecisionSource?: TemplateDecisionSource;
  designProfile?: ResolvedDesignProfile;
  status: DeckStatus;
  audit?: PptxAudit;
  proposal?: CleanupProposal;
  protectedSlideNumbers: number[];
  failureMessage?: string;
  exportedAt?: string;
}

export interface PresentationStudioProject {
  schema: typeof PROJECT_SCHEMA;
  schemaVersion: typeof PROJECT_SCHEMA_VERSION;
  project: {
    id: string;
    name: string;
    type: "review-batch" | "single-deck" | "new-presentation";
    createdAt: string;
    updatedAt: string;
  };
  settings: {
    contentPolicy: "preserve-exact" | "source-grounded-generative";
    defaultOperationScope: OperationScope;
    autosave: boolean;
    designStandardVersion: string;
    defaultProfileId: string;
    defaultSlideSize: "16:9";
    defaultFontFamily: "Aptos";
  };
  resources: ProjectResource[];
  styleExemplars: Array<{
    id: string;
    name: string;
    kind: "table" | "figure";
    resourceId: string;
    deckId: string;
    slideNumber: number;
    objectOrdinal: number;
    scope: "deck" | "batch";
    createdAt: string;
  }>;
  designThreads: DesignThread[];
  decks: DeckJob[];
  activity: Array<{ id: string; at: string; action: string; detail: string }>;
}
