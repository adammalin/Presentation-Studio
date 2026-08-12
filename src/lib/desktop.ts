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
  getMcpStatus(): Promise<{ available: boolean; address: string | null; runtimeFile: string }>;
  getPresentationFonts(): Promise<{ fonts: LocalPresentationFont[] }>;
  getOnboardingTourVersion(): Promise<{ version: string | null }>;
  setOnboardingTourVersion(version: string): Promise<{ saved: boolean; version: string }>;
  openUserGuide(): Promise<{ opened: boolean; path?: string }>;
  onMcpCommand(handler: (request: { id: string; operation: string; input: Record<string, unknown> }) => unknown): () => void;
}
