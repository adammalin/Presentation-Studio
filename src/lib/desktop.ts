import type { DeckQualificationReport, QualificationRenderSummary } from "./deck-qualification";

export interface PickedBinaryFile {
  name: string;
  filePath: string;
  mediaType?: string;
  bytes: Uint8Array;
}

export interface LocalPresentationFont {
  family: string;
  weight: number;
  style: "normal" | "italic";
  mediaType: "font/ttf";
  bytes: Uint8Array;
}

export type NativeRenderStatus = "ready" | "unavailable" | "permission-required" | "failed";

export interface NativeSlideRender {
  number: number;
  mimeType: "image/jpeg" | "image/png";
  width: number;
  height: number;
  sha256: string;
  bytes: Uint8Array;
}

export interface NativeRenderResult {
  status: NativeRenderStatus;
  renderer: "powerpoint-native" | "studio-approximate";
  pipeline?: string;
  powerPointVersion?: string;
  authoritative: boolean;
  sourceSha256?: string;
  generatedAt?: string;
  slideCount?: number;
  slides: NativeSlideRender[];
  warnings: string[];
  reason?: string;
  powerPointInstalled?: boolean;
  rasterizerAvailable?: boolean;
}

export interface NativeRenderCapabilities {
  available: boolean;
  renderer: "powerpoint-native" | "studio-approximate";
  reason?: string;
  powerPointInstalled: boolean;
  rasterizerAvailable: boolean;
  sessionLocked?: boolean;
}

export type MeasurementAuthority = "powerpoint-native" | "direct-ooxml" | "derived-ooxml" | "heuristic" | "unknown";

export interface NativeBoundsPt {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface NativeMarginsPt {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

export interface NativeTextMeasurement {
  marginsPt?: NativeMarginsPt;
  renderedBoundsPt?: NativeBoundsPt;
  coordinateSpace: "slide" | "cell-relative";
  textLength: number;
  lineCount: number;
  verticalAnchor: string;
}

export interface NativeTableCellMeasurement {
  row: number;
  column: number;
  boundsPt?: NativeBoundsPt;
  marginsPt?: NativeMarginsPt;
  renderedTextBoundsPt?: NativeBoundsPt;
  textCoordinateSpace: "cell-relative";
  textLength: number;
  lineCount: number;
  verticalAnchor: string;
}

export interface NativeTableMeasurement {
  rowCount: number;
  columnCount: number;
  rowHeightsPt: number[];
  columnWidthsPt: number[];
  cells: NativeTableCellMeasurement[];
}

export interface NativeShapeMeasurement {
  slideNumber: number;
  shapeIndex: number;
  nativeShapeId?: string;
  name?: string;
  zOrder: number;
  boundsPt?: NativeBoundsPt;
  rotation: number;
  hasTextFrame: boolean;
  hasTable: boolean;
  text?: NativeTextMeasurement;
  table?: NativeTableMeasurement;
}

export interface NativeSlideMeasurement {
  number: number;
  shapeCount: number;
  shapes: NativeShapeMeasurement[];
}

export interface NativeMeasurementResult {
  status: NativeRenderStatus;
  adapter: "macos-powerpoint-applescript" | "windows-powerpoint-com-pending" | "ooxml-fallback";
  authority: MeasurementAuthority;
  sourceSha256?: string;
  generatedAt?: string;
  powerPointVersion?: string;
  slideCount?: number;
  slides: NativeSlideMeasurement[];
  warnings: string[];
  reason?: string;
}

export interface NativeMeasurementCapabilities {
  available: boolean;
  adapter: NativeMeasurementResult["adapter"];
  reason?: string;
  sessionLocked?: boolean;
}

export interface DeckQualificationCaptureResult {
  outputRoot: string;
  sourceRender: QualificationRenderSummary;
  candidateRender: QualificationRenderSummary;
  sourceMeasurement: NativeMeasurementResult;
  candidateMeasurement: NativeMeasurementResult;
}

export interface DesktopBridge {
  isDesktop: true;
  platform: string;
  versions: { chrome: string; electron: string };
  pickPowerPoints(): Promise<{ canceled: boolean; files: PickedBinaryFile[] }>;
  pickResources(): Promise<{ canceled: boolean; files: PickedBinaryFile[] }>;
  pickTemplate(): Promise<{ canceled: boolean; file?: PickedBinaryFile }>;
  installTemplate(payload: { name: string; sha256: string; bytes: Uint8Array }): Promise<{ installed: true; name: string; sha256: string; installedAt: string }>;
  getActiveTemplate(): Promise<{ installed: boolean; name?: string; sha256?: string; installedAt?: string; bytes?: Uint8Array }>;
  openProject(): Promise<{ canceled: boolean; file?: PickedBinaryFile }>;
  saveBinary(payload: { kind: "project" | "secure-project" | "pptx" | "report"; defaultName: string; bytes: Uint8Array }): Promise<{ canceled: boolean; filePath?: string }>;
  autosaveProject(payload: { bytes: Uint8Array; encrypted: boolean }): Promise<{ recoveryPath: string }>;
  checkpointProjectState(payload: { bytes: Uint8Array; encrypted: boolean }): Promise<{ recoveryPath: string }>;
  getAutosaveStatus(): Promise<{ available: boolean; latestModifiedAt?: string }>;
  getAutosaveRecovery(): Promise<{ available: boolean; candidates: Array<{ encrypted: boolean; previous: boolean; modifiedAt: string; package: PickedBinaryFile; checkpoint?: PickedBinaryFile }> }>;
  getMcpStatus(): Promise<{ available: boolean; address: string | null; runtimeFile: string }>;
  getPresentationFonts(): Promise<{ fonts: LocalPresentationFont[] }>;
  getNativeRenderCapabilities(): Promise<NativeRenderCapabilities>;
  getNativeMeasurementCapabilities(): Promise<NativeMeasurementCapabilities>;
  renderPowerPoint(payload: { name: string; bytes: Uint8Array; width?: number; format?: "jpeg" | "png" }): Promise<NativeRenderResult>;
  measurePowerPoint(payload: { name: string; bytes: Uint8Array }): Promise<NativeMeasurementResult>;
  captureDeckQualification(payload: { projectId: string; deckId: string; runId: string; width?: number; source: { name: string; bytes: Uint8Array }; candidate: { name: string; bytes: Uint8Array } }): Promise<DeckQualificationCaptureResult>;
  finalizeDeckQualification(payload: { outputRoot: string; report: DeckQualificationReport }): Promise<{ outputRoot: string; reportPath: string; htmlPath: string }>;
  readDeckQualificationSlide(payload: { outputRoot: string; representation: "source" | "candidate"; slideNumber: number }): Promise<{ mimeType: "image/png"; bytes: Uint8Array; filePath: string }>;
  revealDeckQualification(payload: { outputRoot: string }): Promise<{ revealed: boolean; outputRoot: string; htmlPath: string }>;
  getOnboardingTourVersion(): Promise<{ version: string | null }>;
  setOnboardingTourVersion(version: string): Promise<{ saved: boolean; version: string }>;
  openUserGuide(): Promise<{ opened: boolean; path?: string }>;
  onMcpCommand(handler: (request: { id: string; operation: string; input: Record<string, unknown> }) => unknown): () => void;
}
