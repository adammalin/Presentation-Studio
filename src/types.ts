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
  | "concept-reference"
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
  /** Materialized PowerPoint output number when one source slide produces continuations. */
  outputSlideNumber?: number;
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
  textRuns?: string[];
  paragraphRunCounts?: number[];
  runBreaksBefore?: Array<"none" | "line" | "paragraph">;
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
  paragraphs: TextParagraphInventoryItem[];
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

export interface TextParagraphInventoryItem {
  index: number;
  text: string;
  textHash: string;
  characterCount: number;
  bullet: boolean;
  bulletConfidence: "direct" | "inherited-possible";
  level: number;
  fontFamilies: string[];
  fontSizes: number[];
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
export const STUDIO_WEB_SCENE_VERSION = 5 as const;

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

export type StudioLayoutRecipe = "source" | "ornl-title-content" | "ornl-title-two-column" | "ornl-title-card-grid" | "ornl-title-metric-grid" | "ornl-title-table" | "ornl-title-figure-grid" | "ornl-title-objective-columns" | "ornl-title-steps-evidence" | "ornl-title-labeled-figure-grid" | "ornl-title-question-diagram" | "ornl-title-challenges-evidence" | "ornl-title-process-flow" | "template-layout";
export type StudioWebNodeKind = "text" | "image" | "table" | "shape" | "connector" | "native-object";
export type StudioComponentRole = "eyebrow" | "card-kicker" | "card-heading" | "card-body" | "metric-card" | "objective-body" | "step-heading" | "step-body" | "figure-media" | "figure-label" | "figure-caption" | "technical-annotation" | "question-intro" | "question-item" | "challenge-assertion" | "challenge-intro" | "challenge-body" | "process-icon" | "process-input" | "process-stage" | "process-output" | "supporting-copy" | "footer-logo" | "footer-meta";
export type StudioFigureTreatmentMode = "preserve-as-unit" | "preserve-and-frame" | "hybrid-rebuild" | "redraw-candidate";
export type StudioFigureVerificationStatus = "source-locked" | "needs-content-review" | "verified";
export type StudioFigureRelationshipKind = "caption-for" | "label-for" | "callout-for" | "connects-from" | "connects-to" | "contained-by";
export type StudioConceptInfluence = "composition" | "visual-hierarchy" | "negative-space" | "color-balance" | "figure-concept" | "image-treatment" | "visual-rhythm";
export type StudioConceptUntrustedElement = "generated-text" | "generated-logos" | "generated-data" | "generated-technical-details" | "generated-claims";
export type StudioVisualNeedType = "layout-concept" | "figure-concept" | "image-treatment" | "supporting-visual" | "diagram-rebuild";
export type StudioVisualNeedStatus = "brief-ready" | "concept-attached" | "reconstruction-ready" | "resolved" | "held";
export type StudioVisualExpression = "restrained" | "balanced" | "expressive";
export type StudioVisualDisclosurePolicy = "abstract-structure-only" | "exact-content-approved";
export type StudioVisualMotif = "pattern-free" | "modular-square-grid" | "directional-rule" | "editorial-layering" | "green-motion-gradient" | "subordinate-hex-system";
export type StudioVisualAccent = "none" | "Energy" | "Mist" | "Biome" | "Aqua" | "Infinity" | "Hydro" | "Forge" | "Spark" | "Plasma" | "Pulsar";
export type StudioConstraintKind = "align" | "distribute" | "snap-to-grid" | "fit-safe-region";
export type StudioConstraintMode = "left" | "optical-left" | "center" | "right" | "top" | "optical-top" | "middle" | "bottom" | "horizontal-equal-gap" | "vertical-equal-gap" | "both";
export type StudioTableBorderMode = "none" | "subtle" | "full";
export type StudioTableBorderType = "none" | "solid" | "dash";
export type StudioConnectorSide = "top" | "right" | "bottom" | "left" | "center";
export type StudioConnectorArrow = "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
export type StudioComponentSurface = "light" | "dark";

export interface StudioTableCellBorder {
  type: StudioTableBorderType;
  color: string;
  widthPt: number;
}

export interface StudioTableCellBorders {
  top?: StudioTableCellBorder;
  right?: StudioTableCellBorder;
  bottom?: StudioTableCellBorder;
  left?: StudioTableCellBorder;
}

export interface StudioTableCellDesign {
  cellId: string;
  fill?: string;
  color?: string;
  fontSizePt?: number;
  fontWeight?: 400 | 600 | 700;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  paddingPt?: { top: number; right: number; bottom: number; left: number };
  borders?: StudioTableCellBorders;
}

export interface StudioTableDesign {
  headerRows: number;
  columnWidths: number[];
  rowHeights: number[];
  borderMode: StudioTableBorderMode;
  borderColor: string;
  borderWidthPt: number;
  defaultPaddingPt: { top: number; right: number; bottom: number; left: number };
  cellStyles: StudioTableCellDesign[];
}

export interface StudioTableRoleStyle {
  fill?: string;
  color?: string;
  fontSizePt?: number;
  fontWeight?: 400 | 600 | 700;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  paddingPt?: { top: number; right: number; bottom: number; left: number };
  borders?: StudioTableCellBorders;
}

export interface StudioTableExemplarDefinition {
  id: string;
  name: string;
  sourceNodeId: string;
  adoptedFromSlideNumber: number;
  compatibility: {
    columns: number;
    headerRows: number;
    headerStructure: string;
    bodyStructure: string;
  };
  tableStyle: Pick<StudioTableDesign, "columnWidths" | "borderMode" | "borderColor" | "borderWidthPt" | "defaultPaddingPt">;
  roleStyles: {
    header: StudioTableRoleStyle;
    bodyOdd: StudioTableRoleStyle;
    bodyEven: StudioTableRoleStyle;
  };
  updatedAt: string;
}

export interface StudioTableContinuationSegment {
  ordinal: number;
  bodyRowStart: number;
  bodyRowEnd: number;
  repeatedHeaderRows: number;
  sourceCellIds: string[];
}

export interface StudioTableContinuationPlan {
  id: string;
  sourceSlideNumber: number;
  tableNodeId: string;
  headerRows: number;
  maximumBodyRowsPerSlide: number;
  policy: "repeat-header-rows";
  status: "ready" | "blocked";
  segments: StudioTableContinuationSegment[];
  blockers: string[];
  rationale: string;
  createdAt: string;
}

export interface StudioConnectorDesign {
  fromNodeId: string;
  toNodeId: string;
  fromSide: StudioConnectorSide;
  toSide: StudioConnectorSide;
  stroke: string;
  widthPt: number;
  dash: "solid" | "dash" | "dashDot";
  beginArrow: StudioConnectorArrow;
  endArrow: StudioConnectorArrow;
  verificationStatus: "verified";
}

export interface StudioWebFrame {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
}

export interface StudioDeckRhythm {
  safeMarginPt: number;
  gridPt: number;
  compactGapPt: number;
  normalGapPt: number;
  primaryGapPt: number;
  captionGapPt: number;
  titleContentGapPt: number;
}

export interface StudioOpticalInsets {
  left: number;
  top: number;
  right: number;
  bottom: number;
  authority: "scene-frame" | "source-estimate" | "powerpoint-native";
  basis: "shape" | "rendered-text" | "active-image-content";
}

export interface StudioLayoutConstraint {
  id: string;
  kind: StudioConstraintKind;
  mode: StudioConstraintMode;
  nodeIds: string[];
  groups?: string[][];
  anchorNodeId?: string;
  gridPt?: number;
  rationale: string;
  author: "human" | "ai";
  evidenceAuthority: "scene-estimate" | "powerpoint-native";
  appliedAt: string;
}

export interface StudioQualityIssue {
  id: string;
  category: "overflow" | "alignment" | "spacing" | "safe-region" | "hierarchy" | "figure" | "brand" | "legibility" | "consistency" | "other";
  severity: "blocker" | "major" | "minor";
  source: "powerpoint-native" | "scene" | "ai-visual";
  nodeIds: string[];
  message: string;
  recommendation: string;
  autoFixable: boolean;
}

export interface StudioQualityReview {
  sceneRevision: string;
  slideUpdatedAt: string;
  rasterSha256: string;
  pass: number;
  maxPasses: 3;
  requestedVerdict: "ready" | "revise" | "hold";
  recordedVerdict: "ready" | "revise" | "hold";
  rationale: string;
  objectiveIssues: StudioQualityIssue[];
  visualIssues: StudioQualityIssue[];
  recordedAt: string;
}

export interface StudioDesignMemoryEntry {
  contentSignature: string;
  recipe: StudioLayoutRecipe;
  targetLayoutId?: string;
  targetLayoutName?: string;
  rhythm: StudioDeckRhythm;
  adoptedFromSlideNumber: number;
  qualityRasterSha256: string;
  recordedAt: string;
}

export interface StudioResourceBinding {
  resourceId: string;
  resourceSha256: string;
  derivativeSha256?: string;
  kind: "text" | "image" | "table";
  relationship: "grounds" | "supplies-media" | "supplies-data";
  exactExcerpt?: string;
  exactExcerptHash?: string;
}

export interface StudioWebNode {
  id: string;
  sourceObjectId: string;
  sourceShapeId: string;
  sourceBinding: "editable-object" | "catalog-derived" | "semantic-atom";
  name: string;
  kind: StudioWebNodeKind;
  role: SceneSemanticRole;
  sourceFrame: StudioWebFrame;
  frame: StudioWebFrame;
  zIndex: number;
  sourceTextOrder: number;
  visible: boolean;
  locked: boolean;
  exactContent: boolean;
  text?: string;
  textHash?: string;
  sourceParagraphs?: TextParagraphInventoryItem[];
  sourceAtom?: {
    sourceNodeId: string;
    sourceObjectId: string;
    paragraphStart: number;
    paragraphEnd: number;
    ordinal: number;
    count: number;
    aggregateSourceTextHash?: string;
  };
  resourceBindings?: StudioResourceBinding[];
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
      textRuns?: string[];
      paragraphRunCounts?: number[];
      runBreaksBefore?: Array<"none" | "line" | "paragraph">;
      fill?: string;
      semanticColorRole?: string;
    }>;
    design?: StudioTableDesign;
  };
  connector?: StudioConnectorDesign;
  mediaPart?: string;
  component?: { groupId: string; role: StudioComponentRole; ordinal?: number; frame?: StudioWebFrame; definitionId?: string };
  opticalInsets?: StudioOpticalInsets;
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

export interface StudioComponentDefinition {
  id: string;
  name: string;
  role: StudioComponentRole;
  surface: StudioComponentSurface;
  sourceNodeId: string;
  adoptedFromSlideNumber: number;
  style: StudioWebNode["style"];
  updatedAt: string;
}

export interface StudioFigureTreatment {
  id: string;
  nodeIds: string[];
  mode: StudioFigureTreatmentMode;
  verificationStatus: StudioFigureVerificationStatus;
  intentSummary: string;
  informationInventory: string[];
  invariants: string[];
  rationale: string;
  replacementResourceId?: string;
  relationships?: Array<{ fromNodeId: string; toNodeId: string; kind: StudioFigureRelationshipKind }>;
  groupFrame?: StudioWebFrame;
  focalPoint?: { x: number; y: number };
  crop?: { left: number; top: number; right: number; bottom: number };
  relationshipPolicy?: "preserve-internal" | "reflow-annotations" | "editable-diagram";
  lockAspectRatio?: boolean;
}

export interface StudioConceptReference {
  id: string;
  resourceId: string;
  resourceSha256: string;
  sourceTextHash: string;
  status: "concept-only";
  origin: "imagegen" | "human-reference" | "other";
  approvedInfluences: StudioConceptInfluence[];
  untrustedElements: StudioConceptUntrustedElement[];
  blueprint: {
    summary: string;
    zones: Array<{
      id: string;
      role: "title" | "primary-visual" | "supporting-evidence" | "caption" | "footer-safe" | "other";
      x: number;
      y: number;
      width: number;
      height: number;
    }>;
    styleNotes: string[];
    reconstructionNotes: string[];
  };
  provenance?: {
    model?: string;
    promptSummary?: string;
    generatedAt?: string;
  };
  visualNeedId?: string;
  attachedAt: string;
}

export interface StudioVisualNeed {
  id: string;
  type: StudioVisualNeedType;
  status: StudioVisualNeedStatus;
  sourceTextHash: string;
  reason: string;
  communicationJob: string;
  expression: StudioVisualExpression;
  approvedInfluences: StudioConceptInfluence[];
  disclosurePolicy: StudioVisualDisclosurePolicy;
  approvedContentSummary?: string;
  brandExpression: {
    motif: StudioVisualMotif;
    accent: StudioVisualAccent;
    accentRole: string;
    typographyStrategy: "no-generated-type-reserve-editable-aptos-zones";
    rationale: string;
  };
  structureInventory: {
    titleCount: number;
    textGroupCount: number;
    imageCount: number;
    tableCount: number;
    figureCount: number;
    calloutCount: number;
  };
  targetSlot: {
    role: "whole-slide" | "primary-visual" | "supporting-evidence" | "figure" | "background-treatment";
    aspectRatio: "16:9" | "4:3" | "1:1" | "free";
    placementNotes: string;
  };
  promptPackage: {
    prompt: string;
    negativePrompt: string;
    contentSafety: string;
  };
  linkedConceptReferenceId?: string;
  resolutionNote?: string;
  createdAt: string;
  updatedAt: string;
}

export interface StudioWebSlide {
  id: string;
  slideNumber: number;
  sourceSlideId: string;
  sourceTextHash: string;
  contentCoverage: {
    exactTextMapped: boolean;
    sourceContentSignature?: string;
    sourceCharacterCount: number;
    mappedCharacterCount: number;
    sourceTextBoxCount: number;
    mappedTextNodeCount: number;
    groupedOrUnsupportedTextPresent: boolean;
  };
  sourceRevision: string;
  recipe: StudioLayoutRecipe;
  targetLayoutId?: string;
  targetLayoutName?: string;
  background: string;
  status: "imported" | "designed";
  designRationale: string;
  resourceBindings?: StudioResourceBinding[];
  figureTreatments: StudioFigureTreatment[];
  conceptReferences?: StudioConceptReference[];
  visualNeeds?: StudioVisualNeed[];
  constraints?: StudioLayoutConstraint[];
  qualityReview?: StudioQualityReview;
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
  sourceSlideSize: { width: number; height: number };
  rhythm?: StudioDeckRhythm;
  designMemory?: StudioDesignMemoryEntry[];
  componentLibrary?: StudioComponentDefinition[];
  tableLibrary?: StudioTableExemplarDefinition[];
  tableContinuationPlans?: StudioTableContinuationPlan[];
  designSystem: {
    id: "ornl-presentation-web-v1" | "source-template-preservation-web-v1";
    standardVersion: string;
    unit: "emu";
    renderer: "html-css";
    exportTarget: "editable-powerpoint";
    compilerModes: ["source-bound-overlay", "fresh-composition"];
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
  externalHyperlinkCount: number;
  blockingExternalRelationshipCount: number;
  containsBlockingExternalRelationships: boolean;
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
      intentReview: {
        status: "pass" | "needs-review";
        exactTextPreserved: boolean;
        sourceVisualsPreserved: boolean;
        relationshipsPreserved: "yes" | "not-applicable" | "unverified";
        summary: string;
      };
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
