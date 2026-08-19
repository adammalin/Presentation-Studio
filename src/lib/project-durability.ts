import type { PresentationStudioProject, ProjectResource } from "../types";
import { projectForJson, projectSchema } from "./project";

export function projectHasRecoverableWork(project: PresentationStudioProject): boolean {
  return project.resources.length > 0
    || project.decks.length > 0
    || project.designThreads.length > 0
    || project.styleExemplars.length > 0
    || project.activity.some((entry) => entry.action !== "project-created");
}

export function projectResourceInventoryKey(project: PresentationStudioProject): string {
  return `${project.project.id}:${project.resources.map((resource) => `${resource.id}:${resource.sha256}:${resource.byteLength}:${(resource.derivatives ?? []).map((derivative) => `${derivative.id}:${derivative.sha256}:${derivative.byteLength}`).join(",")}`).join("|")}`;
}

export const PROJECT_RECOVERY_CHECKPOINT_SCHEMA = "presentation-studio/recovery-checkpoint" as const;
export const PROJECT_RECOVERY_CHECKPOINT_VERSION = 1 as const;

interface RecoveryResourceInventory {
  id: string;
  sha256: string;
  byteLength: number;
  derivatives: Array<{ id: string; sha256: string; byteLength: number }>;
}

interface ProjectRecoveryCheckpoint {
  schema: typeof PROJECT_RECOVERY_CHECKPOINT_SCHEMA;
  version: typeof PROJECT_RECOVERY_CHECKPOINT_VERSION;
  createdAt: string;
  projectId: string;
  projectUpdatedAt: string;
  resources: RecoveryResourceInventory[];
  project: ReturnType<typeof projectForJson>;
}

function inventory(resources: ProjectResource[]): RecoveryResourceInventory[] {
  return resources.map((resource) => ({
    id: resource.id,
    sha256: resource.sha256,
    byteLength: resource.byteLength,
    derivatives: (resource.derivatives ?? []).map((derivative) => ({ id: derivative.id, sha256: derivative.sha256, byteLength: derivative.byteLength })),
  }));
}

function sameInventory(left: RecoveryResourceInventory[], right: RecoveryResourceInventory[]): boolean {
  if (left.length !== right.length) return false;
  const normalized = (items: RecoveryResourceInventory[]) => items
    .map((item) => ({ ...item, derivatives: [...item.derivatives].sort((a, b) => a.id.localeCompare(b.id)) }))
    .sort((a, b) => a.id.localeCompare(b.id));
  return JSON.stringify(normalized(left)) === JSON.stringify(normalized(right));
}

export function buildProjectRecoveryCheckpoint(project: PresentationStudioProject): Uint8Array {
  const checkpoint: ProjectRecoveryCheckpoint = {
    schema: PROJECT_RECOVERY_CHECKPOINT_SCHEMA,
    version: PROJECT_RECOVERY_CHECKPOINT_VERSION,
    createdAt: new Date().toISOString(),
    projectId: project.project.id,
    projectUpdatedAt: project.project.updatedAt,
    resources: inventory(project.resources),
    project: projectForJson(project),
  };
  return new TextEncoder().encode(`${JSON.stringify(checkpoint)}\n`);
}

export function applyProjectRecoveryCheckpoint(base: PresentationStudioProject, bytes: Uint8Array): PresentationStudioProject {
  let decoded: unknown;
  try {
    decoded = JSON.parse(new TextDecoder().decode(bytes));
  } catch {
    throw new Error("The autosave checkpoint is not valid JSON.");
  }
  if (!decoded || typeof decoded !== "object" || Array.isArray(decoded)) throw new Error("The autosave checkpoint is invalid.");
  const checkpoint = decoded as Partial<ProjectRecoveryCheckpoint>;
  if (checkpoint.schema !== PROJECT_RECOVERY_CHECKPOINT_SCHEMA || checkpoint.version !== PROJECT_RECOVERY_CHECKPOINT_VERSION || typeof checkpoint.projectId !== "string" || typeof checkpoint.projectUpdatedAt !== "string" || !Array.isArray(checkpoint.resources)) {
    throw new Error("The autosave checkpoint version is not supported.");
  }
  if (checkpoint.projectId !== base.project.id) throw new Error("The autosave checkpoint belongs to a different project.");
  if (!sameInventory(checkpoint.resources, inventory(base.resources))) throw new Error("The autosave checkpoint does not match the embedded project Resources.");
  const recovered = projectSchema.parse(checkpoint.project) as PresentationStudioProject;
  if (recovered.project.id !== base.project.id || recovered.project.updatedAt !== checkpoint.projectUpdatedAt) throw new Error("The autosave checkpoint project identity is inconsistent.");
  if (Date.parse(recovered.project.updatedAt) < Date.parse(base.project.updatedAt)) return base;
  const baseResources = new Map(base.resources.map((resource) => [resource.id, resource]));
  const resources = recovered.resources.map((resource) => {
    const embedded = baseResources.get(resource.id);
    if (!embedded?.bytes) throw new Error(`The full autosave package is missing ${resource.name}.`);
    const embeddedDerivatives = new Map((embedded.derivatives ?? []).map((derivative) => [derivative.id, derivative]));
    return {
      ...resource,
      sourcePath: undefined,
      mcpAccess: "none" as const,
      bytes: embedded.bytes,
      derivatives: resource.derivatives?.map((derivative) => {
        const embeddedDerivative = embeddedDerivatives.get(derivative.id);
        if (!embeddedDerivative?.bytes) throw new Error(`The full autosave package is missing a derivative for ${resource.name}.`);
        return { ...derivative, bytes: embeddedDerivative.bytes };
      }),
    };
  });
  return { ...recovered, resources };
}

export type AutosavePhase = "pending" | "saving" | "saved" | "error";

export interface AutosaveProgress {
  phase: AutosavePhase;
  revision?: string;
  savedAt?: string;
  error?: string;
}

export interface LatestOnlySaver<T> {
  request(value: T): void;
  flush(): Promise<void>;
  dispose(): Promise<void>;
}

export function createLatestOnlySaver<T>(options: {
  delayMs: number;
  revision(value: T): string;
  save(value: T): Promise<void>;
  onProgress?(progress: AutosaveProgress): void;
}): LatestOnlySaver<T> {
  let pending: T | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let running: Promise<void> | undefined;
  let disposed = false;
  let lastError: Error | undefined;

  const drain = async () => {
    if (running) return running;
    running = (async () => {
      while (pending) {
        const value = pending;
        pending = undefined;
        const revision = options.revision(value);
        options.onProgress?.({ phase: "saving", revision });
        try {
          await options.save(value);
          lastError = undefined;
          options.onProgress?.({ phase: "saved", revision, savedAt: new Date().toISOString() });
        } catch (caught) {
          lastError = caught instanceof Error ? caught : new Error("Autosave failed.");
          options.onProgress?.({ phase: "error", revision, error: lastError.message });
        }
      }
    })().finally(() => { running = undefined; });
    return running;
  };

  const schedule = () => {
    if (timer || disposed) return;
    timer = setTimeout(() => {
      timer = undefined;
      void drain();
    }, Math.max(0, options.delayMs));
  };

  return {
    request(value) {
      if (disposed) return;
      pending = value;
      lastError = undefined;
      options.onProgress?.({ phase: "pending", revision: options.revision(value) });
      schedule();
    },
    async flush() {
      if (timer) { clearTimeout(timer); timer = undefined; }
      await drain();
      if (pending) await drain();
      if (lastError) throw lastError;
    },
    async dispose() {
      disposed = true;
      if (timer) { clearTimeout(timer); timer = undefined; }
      await drain();
    },
  };
}
