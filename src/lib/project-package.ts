import JSZip from "jszip";
import type { PresentationStudioProject, ProjectResource } from "../types";
import { sha256 } from "./hash";
import { projectForJson, projectSchema } from "./project";

const MAX_PACKAGE_FILES = 50_000;
const MAX_PROJECT_JSON_BYTES = 25 * 1024 * 1024;
const MAX_RESOURCE_BYTES = 1_000_000_000;
const MAX_DERIVATIVE_BYTES = 5 * 1024 * 1024;
const MAX_TOTAL_RESOURCE_BYTES = 1_250_000_000;

interface PackageManifest {
  schema: "presentation-studio/package-manifest";
  version: 1;
  createdAt: string;
  projectPath: "project.json";
  resources: Array<{ id: string; path: string; sha256: string; byteLength: number; mediaType: string }>;
  derivatives?: Array<{ resourceId: string; id: string; path: string; sha256: string; byteLength: number; mediaType: string }>;
}

function resourcePath(resource: ProjectResource): string {
  return `resources/blobs/${resource.sha256}`;
}

function derivativePath(resource: ProjectResource, sha256: string): string {
  return `resources/derivatives/${resource.sha256}/${sha256}`;
}

export async function buildProjectPackage(project: PresentationStudioProject): Promise<Uint8Array> {
  const zip = new JSZip();
  const resources: PackageManifest["resources"] = [];
  const derivatives: NonNullable<PackageManifest["derivatives"]> = [];
  for (const resource of project.resources) {
    if (!resource.bytes) throw new Error(`${resource.name} is missing its embedded project bytes.`);
    const digest = await sha256(resource.bytes);
    if (digest !== resource.sha256) throw new Error(`${resource.name} failed the project-package integrity check.`);
    const path = resourcePath(resource);
    if (!zip.file(path)) zip.file(path, resource.bytes, { binary: true });
    resources.push({ id: resource.id, path, sha256: resource.sha256, byteLength: resource.byteLength, mediaType: resource.mediaType });
    for (const derivative of resource.derivatives ?? []) {
      if (!derivative.bytes) throw new Error(`${resource.name} is missing its ${derivative.kind} derivative bytes.`);
      if (derivative.bytes.byteLength > MAX_DERIVATIVE_BYTES) throw new Error(`${resource.name} has an oversized ${derivative.kind} derivative.`);
      const derivativeDigest = await sha256(derivative.bytes);
      if (derivativeDigest !== derivative.sha256 || derivative.bytes.byteLength !== derivative.byteLength) throw new Error(`${resource.name} failed a derivative integrity check.`);
      const derivativeMember = derivativePath(resource, derivative.sha256);
      if (!zip.file(derivativeMember)) zip.file(derivativeMember, derivative.bytes, { binary: true });
      derivatives.push({ resourceId: resource.id, id: derivative.id, path: derivativeMember, sha256: derivative.sha256, byteLength: derivative.byteLength, mediaType: derivative.mediaType });
    }
  }
  const manifest: PackageManifest = {
    schema: "presentation-studio/package-manifest",
    version: 1,
    createdAt: new Date().toISOString(),
    projectPath: "project.json",
    resources,
    derivatives,
  };
  zip.file("manifest.json", `${JSON.stringify(manifest, null, 2)}\n`);
  zip.file("project.json", `${JSON.stringify(projectForJson(project), null, 2)}\n`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

export async function openProjectPackage(bytes: Uint8Array): Promise<PresentationStudioProject> {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  const paths = Object.keys(zip.files).filter((path) => !zip.files[path].dir);
  if (paths.length > MAX_PACKAGE_FILES) throw new Error("This project contains too many package entries.");
  if (paths.some((path) => path.startsWith("/") || path.includes("\\") || path.split("/").includes(".."))) throw new Error("This project contains an unsafe package path.");
  const manifestEntry = zip.file("manifest.json");
  const projectEntry = zip.file("project.json");
  if (!manifestEntry || !projectEntry) throw new Error("This package is missing manifest.json or project.json.");
  const manifestBytes = await manifestEntry.async("uint8array");
  const projectBytes = await projectEntry.async("uint8array");
  if (manifestBytes.byteLength > 5_000_000 || projectBytes.byteLength > MAX_PROJECT_JSON_BYTES) throw new Error("This project contains oversized package metadata.");
  const manifest = JSON.parse(new TextDecoder().decode(manifestBytes)) as PackageManifest;
  if (manifest.schema !== "presentation-studio/package-manifest" || manifest.version !== 1) {
    throw new Error("This Presentation Studio package version is not supported.");
  }
  if (!Array.isArray(manifest.resources) || manifest.resources.length > 10_000) throw new Error("This project has invalid Resource metadata.");
  let declaredTotal = 0;
  for (const item of manifest.resources) {
    if (!item || typeof item.id !== "string" || typeof item.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(item.sha256) || item.path !== `resources/blobs/${item.sha256}` || !Number.isInteger(item.byteLength) || item.byteLength < 0 || item.byteLength > MAX_RESOURCE_BYTES) {
      throw new Error("This project has invalid Resource package metadata.");
    }
    declaredTotal += item.byteLength;
  }
  const manifestDerivatives = manifest.derivatives ?? [];
  if (!Array.isArray(manifestDerivatives) || manifestDerivatives.length > 200_000) throw new Error("This project has invalid Resource derivative metadata.");
  for (const item of manifestDerivatives) {
    const parent = manifest.resources.find((resource) => resource.id === item?.resourceId);
    if (!item || !parent || typeof item.id !== "string" || typeof item.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(item.sha256) || item.path !== `resources/derivatives/${parent.sha256}/${item.sha256}` || !Number.isInteger(item.byteLength) || item.byteLength < 0 || item.byteLength > MAX_DERIVATIVE_BYTES) {
      throw new Error("This project has invalid Resource derivative package metadata.");
    }
    declaredTotal += item.byteLength;
  }
  if (declaredTotal > MAX_TOTAL_RESOURCE_BYTES) throw new Error("This project's embedded Resources exceed the 1.25 GB safety limit.");
  const parsed = projectSchema.parse(JSON.parse(new TextDecoder().decode(projectBytes))) as PresentationStudioProject;
  const manifestIds = new Set(manifest.resources.map((item) => item.id));
  if (manifestIds.size !== manifest.resources.length || parsed.resources.length !== manifest.resources.length) throw new Error("This project has duplicate or inconsistent Resource entries.");
  const parsedDerivativeCount = parsed.resources.reduce((count, resource) => count + (resource.derivatives?.length ?? 0), 0);
  const derivativeIds = new Set(manifestDerivatives.map((item) => `${item.resourceId}:${item.id}`));
  if (derivativeIds.size !== manifestDerivatives.length || parsedDerivativeCount !== manifestDerivatives.length) throw new Error("This project has duplicate or inconsistent Resource derivative entries.");
  const byId = new Map(manifest.resources.map((item) => [item.id, item]));
  const derivativesById = new Map(manifestDerivatives.map((item) => [`${item.resourceId}:${item.id}`, item]));
  const resources: ProjectResource[] = [];
  for (const resource of parsed.resources) {
    const packaged = byId.get(resource.id);
    if (!packaged || packaged.sha256 !== resource.sha256 || packaged.byteLength !== resource.byteLength) {
      throw new Error(`${resource.name} has inconsistent package metadata.`);
    }
    const entry = zip.file(packaged.path);
    if (!entry) throw new Error(`${resource.name} is missing from the package.`);
    const embeddedBytes = await entry.async("uint8array");
    if (embeddedBytes.byteLength !== resource.byteLength || await sha256(embeddedBytes) !== resource.sha256) {
      throw new Error(`${resource.name} failed package integrity validation.`);
    }
    const derivatives = [];
    for (const derivative of resource.derivatives ?? []) {
      const packagedDerivative = derivativesById.get(`${resource.id}:${derivative.id}`);
      if (!packagedDerivative || packagedDerivative.sha256 !== derivative.sha256 || packagedDerivative.byteLength !== derivative.byteLength) {
        throw new Error(`${resource.name} has inconsistent derivative metadata.`);
      }
      const derivativeEntry = zip.file(packagedDerivative.path);
      if (!derivativeEntry) throw new Error(`${resource.name} is missing a packaged derivative.`);
      const derivativeBytes = await derivativeEntry.async("uint8array");
      if (derivativeBytes.byteLength !== derivative.byteLength || await sha256(derivativeBytes) !== derivative.sha256) throw new Error(`${resource.name} failed derivative integrity validation.`);
      derivatives.push({ ...derivative, bytes: derivativeBytes });
    }
    resources.push({ ...resource, derivatives, sourcePath: undefined, bytes: embeddedBytes, mcpAccess: "none" });
  }
  return { ...parsed, resources };
}
