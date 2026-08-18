import type { ProjectResource } from "../types";

export type ResourceAiAccess = ProjectResource["mcpAccess"];

/**
 * Resolve the most useful access Presentation Studio can actually provide for
 * an embedded Resource. The single visible AI-session switch is the access
 * decision; callers should not require a second per-Resource approval.
 */
export function automaticResourceAiAccess(resource: ProjectResource): ResourceAiAccess {
  const extractedText = resource.derivatives?.some((derivative) =>
    derivative.kind === "extracted-text" && Boolean(derivative.bytes?.byteLength),
  );
  if (extractedText) return "text";
  if (resource.kind === "image" && resource.support?.includes("previewable") && Boolean(resource.bytes?.byteLength)) return "preview";
  return "metadata";
}

export function resourceWithAiSessionAccess(resource: ProjectResource, enabled: boolean): ProjectResource {
  const mcpAccess: ResourceAiAccess = enabled ? automaticResourceAiAccess(resource) : "none";
  return resource.mcpAccess === mcpAccess ? resource : { ...resource, mcpAccess };
}

export function resourcesWithAiSessionAccess(resources: ProjectResource[], enabled: boolean): ProjectResource[] {
  let changed = false;
  const next = resources.map((resource) => {
    const resolved = resourceWithAiSessionAccess(resource, enabled);
    if (resolved !== resource) changed = true;
    return resolved;
  });
  return changed ? next : resources;
}
