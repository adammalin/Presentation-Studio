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
  sourcePart?: string;
  sourcePartSha256?: string;
  relationshipPart?: string;
  relationshipPartSha256?: string;
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
  semanticColorTokens: string[];
  marginSignatures: string[];
  styleFingerprint: string;
  contentHash: string;
  structureHash: string;
  columns?: TableColumnInventoryItem[];
  rows?: TableRowInventoryItem[];
  cells?: TableCellInventoryItem[];
}

export interface TableColumnInventoryItem {
  id: string;
  index: number;
  widthEmu: number;
}

export interface TableRowInventoryItem {
  id: string;
  index: number;
  heightEmu: number;
}

export interface TableCellInventoryItem {
  id: string;
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
  horizontalMergeContinuation: boolean;
  verticalMergeContinuation: boolean;
  text: string;
  textHash: string;
  characterCount: number;
  paragraphCount: number;
  fontFamilies: string[];
  fontSizes: number[];
  fillToken?: string;
  semanticColorRole?: string;
  marginsEmu: { left: number; right: number; top: number; bottom: number };
  horizontalAlignment: "left" | "center" | "right" | "justified" | "mixed";
  verticalAlignment: "top" | "middle" | "bottom";
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

export interface TextBoxInventoryItem {
  id: string;
  slideNumber: number;
  ordinal: number;
  shapeId: string;
  text: string;
  textHash: string;
  characterCount: number;
  paragraphCount: number;
  geometry: { x: number; y: number; width: number; height: number };
  textInsets: { left: number; right: number; top: number; bottom: number };
  paragraphLeftMarginsEmu: number[];
  paragraphIndentsEmu: number[];
  bulletParagraphCount: number;
  opticalLeftOffsetEmu: number;
  estimatedOpticalLeftEmu: number;
  opticalAlignmentConfidence: "direct" | "partial-inheritance";
  fontFamilies: string[];
  fontSizes: number[];
  directFontSizeKnown: boolean;
  paragraphAlignment: "left" | "center" | "right" | "justified" | "mixed";
  verticalAlignment: "top" | "middle" | "bottom";
  role: "title" | "body" | "caption" | "label" | "other";
  autoFit: "none" | "shrink-text" | "resize-shape" | "unspecified";
  estimatedLineCount: number;
  estimatedRequiredHeightEmu: number;
  fitRatio: number;
  safeAreaStatus: "inside" | "near-edge" | "off-slide";
  warnings: string[];
}

export interface LayoutReviewItem {
  id: string;
  slideNumber: number;
  shapeId: string;
  rule: "overflow-risk" | "off-slide" | "safe-area" | "alignment-ambiguous";
  severity: Severity;
  confidence: "high" | "medium" | "low";
  reason: string;
  geometry: { x: number; y: number; width: number; height: number };
  fitRatio?: number;
}

export interface AlignmentRepairCandidate {
  id: string;
  slideNumber: number;
  shapeId: string;
  textHash: string;
  source: { x: number; y: number; width: number; height: number };
  target: { x: number; y: number; width: number; height: number };
  ruleId: "cover.dominant-left-edge" | "peer.dominant-left-edge";
  confidence: "high";
  rationale: string;
}

export type SlideEditableObjectKind = "text" | "shape" | "picture" | "table" | "chart" | "connector" | "group" | "graphic-frame";
export type SlideEditableObjectElement = "p:sp" | "p:pic" | "p:graphicFrame" | "p:cxnSp" | "p:grpSp";

export interface SlideEditableObject {
  id: string;
  slideNumber: number;
  shapeId: string;
  name: string;
  kind: SlideEditableObjectKind;
  sourceElement: SlideEditableObjectElement;
  geometry: { x: number; y: number; width: number; height: number; rotation: number };
  canMove: boolean;
  canResize: boolean;
  textHash?: string;
  tableId?: string;
  pictureId?: string;
}

export const PRESENTATION_SCENE_SCHEMA = "presentation-studio/scene" as const;
export const PRESENTATION_SCENE_VERSION = 3 as const;
export const PRESERVATION_ENVELOPE_SCHEMA = "presentation-studio/preservation-envelope" as const;
export const PRESERVATION_ENVELOPE_VERSION = 1 as const;
export const STUDIO_WEB_SCENE_SCHEMA = "presentation-studio/web-scene" as const;
export const STUDIO_WEB_SCENE_VERSION = 1 as const;

export type SceneFidelityState = "editable-native" | "preserved-native" | "conversion-required" | "unsupported-blocking";
export type SceneSemanticRole = "title" | "body" | "caption" | "label" | "image" | "table" | "chart" | "connector" | "group" | "decoration" | "other";
export type SceneRepresentationState = "native" | "partial" | "preserved" | "none";

export interface SceneFidelityCounts {
  "editable-native": number;
  "preserved-native": number;
  "conversion-required": number;
  "unsupported-blocking": number;
}

export interface PresentationSceneObject {
  id: string;
  slideId: string;
  slideNumber: number;
  shapeId: string;
  name: string;
  kind: SlideEditableObjectKind;
  sourceElement: SlideEditableObjectElement;
  semanticRole: SceneSemanticRole;
  fidelityState: SceneFidelityState;
  fidelityReason: string;
  geometry: { x: number; y: number; width: number; height: number; rotation: number };
  zIndex: number;
  sourceLocator: {
    slidePart: string;
    shapeId: string;
    tableId?: string;
    pictureId?: string;
  };
  representation: {
    geometry: SceneRepresentationState;
    text: SceneRepresentationState;
    style: SceneRepresentationState;
    internalStructure: SceneRepresentationState;
  };
  operations: {
    move: boolean;
    resize: boolean;
    restyle: boolean;
    editText: boolean;
    editTableStyle: boolean;
    replaceMedia: boolean;
    editChartData: boolean;
    editInternalStructure: boolean;
  };
  contentHash?: string;
  protected: boolean;
}

export interface PresentationSceneSlide {
  id: string;
  number: number;
  sourcePart: string;
  sourcePartSha256?: string;
  relationshipPart?: string;
  relationshipPartSha256?: string;
  sourceTextHash: string;
  targetLayoutId?: string;
  objectIds: string[];
  fidelityCounts: SceneFidelityCounts;
  preservationRequired: boolean;
  protected: boolean;
}

export interface PresentationSceneTableColumn {
  id: string;
  index: number;
  width: number;
  x: number;
}

export interface PresentationSceneTableRow {
  id: string;
  index: number;
  height: number;
  y: number;
}

export interface PresentationSceneTableCell {
  id: string;
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
  geometry: { x: number; y: number; width: number; height: number };
  margins: { left: number; right: number; top: number; bottom: number };
  contentHash: string;
  characterCount: number;
  horizontalAlignment: TableCellInventoryItem["horizontalAlignment"];
  verticalAlignment: TableCellInventoryItem["verticalAlignment"];
  mergeContinuation: boolean;
}

export interface PresentationSceneTable {
  id: string;
  objectId: string;
  slideNumber: number;
  rowCount: number;
  columnCount: number;
  geometry: { x: number; y: number; width: number; height: number };
  rows: PresentationSceneTableRow[];
  columns: PresentationSceneTableColumn[];
  cells: PresentationSceneTableCell[];
  contentHash: string;
  structureHash: string;
}

export interface PowerPointPreservationEnvelope {
  schema: typeof PRESERVATION_ENVELOPE_SCHEMA;
  version: typeof PRESERVATION_ENVELOPE_VERSION;
  sourceResourceId: string;
  sourceSha256: string;
  sourceBytesAuthoritative: true;
  nativeRenderAuthoritativeForAppearance: true;
  exportStrategy: "surgical-ooxml-overlay";
  packageFileCount: number;
  expandedByteLength: number;
  protectedFeatures: {
    macros: boolean;
    oleObjects: boolean;
    externalRelationships: boolean;
  };
  blockingFeatures: Array<"macros" | "ole-objects" | "external-relationships">;
  slides: Array<{
    slideId: string;
    slideNumber: number;
    sourcePart: string;
    sourcePartSha256?: string;
    relationshipPart?: string;
    relationshipPartSha256?: string;
    sourceTextHash: string;
    objectIds: string[];
  }>;
}

export interface PresentationScene {
  schema: typeof PRESENTATION_SCENE_SCHEMA;
  version: typeof PRESENTATION_SCENE_VERSION;
  revision: string;
  sourceSha256: string;
  slideSize: { width: number; height: number };
  templateBinding: {
    sourceClassification: TemplateClassification;
    targetTemplateId?: string;
    targetDecisionSource?: TemplateDecisionSource;
  };
  slides: PresentationSceneSlide[];
  objects: PresentationSceneObject[];
  tables?: PresentationSceneTable[];
  fidelityCounts: SceneFidelityCounts;
  preservationEnvelope: PowerPointPreservationEnvelope;
}

export type StudioLayoutRecipe = "source" | "ornl-title-content" | "ornl-title-two-column" | "ornl-title-table" | "ornl-title-figure-grid" | "template-layout";
export type StudioWebNodeKind = "text" | "image" | "table" | "shape" | "connector" | "native-object";

export interface StudioWebFrame {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
}

export interface StudioWebNode {
  id: string;
  sourceObjectId: string;
  sourceShapeId: string;
  name: string;
  kind: StudioWebNodeKind;
  role: SceneSemanticRole;
  sourceFrame: StudioWebFrame;
  frame: StudioWebFrame;
  zIndex: number;
  visible: boolean;
  locked: boolean;
  exactContent: boolean;
  text?: string;
  textHash?: string;
  tableId?: string;
  table?: {
    rows: number;
    columns: number;
    cells: Array<{
      id: string;
      row: number;
      column: number;
      rowSpan: number;
      columnSpan: number;
      text: string;
      fill?: string;
      semanticColorRole?: string;
    }>;
  };
  mediaPart?: string;
  style: {
    fontFamily: "Aptos";
    fontSizePt: number;
    fontWeight: 400 | 600 | 700;
    lineHeight: number;
    color: string;
    background?: string;
    borderColor?: string;
    borderWidthPt: number;
    textAlign: "left" | "center" | "right";
    verticalAlign: "top" | "middle" | "bottom";
    paddingPt: { top: number; right: number; bottom: number; left: number };
    objectFit?: "contain" | "cover";
  };
}

export interface StudioWebSlide {
  id: string;
  slideNumber: number;
  sourceSlideId: string;
  sourceTextHash: string;
  sourceRevision: string;
  recipe: StudioLayoutRecipe;
  targetLayoutId?: string;
  targetLayoutName?: string;
  background: string;
  status: "imported" | "designed";
  designRationale: string;
  nodes: StudioWebNode[];
  updatedAt: string;
}

export interface StudioWebScene {
  schema: typeof STUDIO_WEB_SCENE_SCHEMA;
  version: typeof STUDIO_WEB_SCENE_VERSION;
  revision: string;
  deckId: string;
  sourceSha256: string;
  slideSize: { width: number; height: number };
  designSystem: {
    id: "ornl-presentation-web-v1";
    standardVersion: string;
    unit: "emu";
    renderer: "html-css";
    exportTarget: "editable-powerpoint";
  };
  slides: StudioWebSlide[];
}

export interface GeometryEditCommand {
  id: string;
  slideNumber: number;
  objectId: string;
  shapeId: string;
  sourceElement: SlideEditableObjectElement;
  objectKind: SlideEditableObjectKind;
  operation: "move" | "resize" | "move-and-resize";
  source: { x: number; y: number; width: number; height: number };
  target: { x: number; y: number; width: number; height: number };
  rationale: string;
  author: "human" | "ai";
  constraints: {
    allowIntentionalOverlap: boolean;
    allowFitRisk: boolean;
    allowSafeArea: boolean;
    allowAspectRatioChange: boolean;
  };
  validation: {
    fitRatio?: number;
    safeAreaStatus: "inside" | "near-edge";
    overlapObjectIds: string[];
    warnings: string[];
  };
}

export interface TextStyleCommand {
  id: string;
  slideNumber: number;
  objectId: string;
  shapeId: string;
  typeface: "Aptos";
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  alignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "middle" | "bottom";
  insetsInches?: { top: number; right: number; bottom: number; left: number };
  paragraphStyle?: {
    lineSpacingMultiple?: number;
    spaceAfterPt?: number;
    bulletLeftMarginInches?: number;
    bulletHangingInches?: number;
  };
  rationale: string;
  author: "human" | "ai";
}

export type TableStyleVariant = "standard" | "dense-technical";

export interface TableLayoutCommand {
  id: string;
  slideNumber: number;
  tableId: string;
  objectId: string;
  columnWidthsEmu: number[];
  rowHeightsEmu: number[];
  cellMarginsEmu: { left: number; right: number; top: number; bottom: number };
  rationale: string;
  author: "human" | "ai";
  constraints: {
    minimumFontPt: number;
    minimumHorizontalPaddingPt: number;
    minimumVerticalPaddingPt: number;
    preserveTableBounds: boolean;
  };
  validation: {
    feasible: boolean;
    predictedOverflowCellIds: string[];
    warnings: string[];
  };
}

export interface DecorativeShapeCommand {
  id: string;
  slideNumber: number;
  name: string;
  geometry: { x: number; y: number; width: number; height: number };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt: number;
  behindContent: boolean;
  rationale: string;
  author: "human" | "ai";
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
  semanticVisualVersion: number;
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
  slideSize: { width: number; height: number };
  classification: TemplateClassification;
  classificationEvidence: string[];
  fonts: FontInventoryItem[];
  slides: SlideInventoryItem[];
  tables: TableInventoryItem[];
  pictures: PictureInventoryItem[];
  textBoxes: TextBoxInventoryItem[];
  editableObjects: SlideEditableObject[];
  layoutReviews: LayoutReviewItem[];
  alignmentRepairs: AlignmentRepairCandidate[];
  findings: AuditFinding[];
  warnings: string[];
}

export interface CleanupChange {
  id: string;
  kind: "font-family" | "table-style" | "table-layout" | "alignment" | "geometry" | "layout-remap" | "text-style" | "decoration";
  from: string;
  to: string;
  affectedSlideNumbers: number[];
  affectedRunCount: number;
  tableIds?: string[];
  tableLayoutCommands?: TableLayoutCommand[];
  profileId?: string;
  semanticColorPolicy?: "preserve-source";
  alignmentRepairs?: AlignmentRepairCandidate[];
  geometryCommands?: GeometryEditCommand[];
  layoutCommands?: NativeLayoutRemapCommand[];
  textStyleCommands?: TextStyleCommand[];
  decorationCommands?: DecorativeShapeCommand[];
  rationale: string;
  selected: boolean;
}

export interface NativeLayoutRemapCommand {
  id: string;
  slideNumber: number;
  templateSha256: string;
  templateLayoutPart: string;
  templateLayoutSha256: string;
  templateLayoutName: string;
  rationale: string;
  author: "ai" | "human";
}

export interface SlideDesignDisposition {
  slideNumber: number;
  status: "change-proposed" | "approved-as-is" | "needs-review";
  reasons: string[];
  changeIds: string[];
}

export interface SlideReviewDecision {
  slideNumber: number;
  decision: "approved" | "changes-requested";
  reviewedAt: string;
  comment?: string;
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
  mode: "font-cleanup" | "designer-cleanup" | "slide-geometry" | "slide-reflow";
  standardVersion?: string;
  designDecision?: {
    kind: "semantic-recomposition";
    workOrderRevision: string;
    targetLayoutId: string;
    targetLayoutName: string;
    targetLayoutSourcePart: string;
    compatibilityScore: number;
    compatibilityStatus: "recommended" | "compatible" | "poor" | "incompatible";
    rationale: string;
    bindingCount: number;
    application: "semantic-zones-on-source-layout" | "cloned-native-layout";
  };
  designReview?: {
    decision: "rejected";
    actor: "ai" | "human";
    rationale: string;
    reviewedAt: string;
    evidence?: {
      slideNumber: number;
      renderer: "powerpoint-native";
      currentRasterSha256: string;
      proposalRasterSha256: string;
      changedPixelRatio: number;
    };
  };
  visualIteration?: {
    maxAttempts: 3;
    history: Array<{
      attempt: number;
      slideNumber: number;
      inspectionRevision: string;
      verdict: "better" | "revise" | "reject";
      rationale: string;
      reviewedAt: string;
      currentRasterSha256: string;
      proposalRasterSha256: string;
      metrics: { improvements: string[]; regressions: string[] };
    }>;
  };
  changes: CleanupChange[];
  slideDispositions: SlideDesignDisposition[];
  slideReviews?: SlideReviewDecision[];
  tableExceptions: TableNormalizationException[];
  layoutExceptions: LayoutReviewItem[];
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
  scene?: PresentationScene;
  studioScene?: StudioWebScene;
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
