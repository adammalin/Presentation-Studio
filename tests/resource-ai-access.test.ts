import assert from "node:assert/strict";
import test from "node:test";
import type { ProjectResource } from "../src/types";
import { automaticResourceAiAccess, resourcesWithAiSessionAccess } from "../src/lib/resource-ai-access";

function resource(overrides: Partial<ProjectResource>): ProjectResource {
  return {
    id: "resource",
    name: "resource.bin",
    mediaType: "application/octet-stream",
    byteLength: 3,
    sha256: "a".repeat(64),
    roles: ["reference-only"],
    kind: "other",
    support: ["unsupported"],
    createdAt: "2026-08-17T12:00:00.000Z",
    embedded: true,
    bytes: new Uint8Array([1, 2, 3]),
    mcpAccess: "none",
    ...overrides,
  };
}

test("AI session access automatically exposes every Resource at its highest supported level", () => {
  const textBytes = new TextEncoder().encode("source text");
  const document = resource({
    id: "document",
    name: "source.docx",
    kind: "document",
    support: ["source-readable"],
    derivatives: [{
      id: "text",
      kind: "extracted-text",
      mediaType: "text/plain",
      byteLength: textBytes.byteLength,
      sha256: "b".repeat(64),
      createdAt: "2026-08-17T12:00:00.000Z",
      processor: "test",
      truncated: false,
      bytes: textBytes,
    }],
  });
  const image = resource({ id: "image", name: "figure.png", kind: "image", mediaType: "image/png", support: ["previewable", "placeable"] });
  const unsupported = resource({ id: "legacy", name: "legacy.doc", kind: "document" });

  assert.equal(automaticResourceAiAccess(document), "text");
  assert.equal(automaticResourceAiAccess(image), "preview");
  assert.equal(automaticResourceAiAccess(unsupported), "metadata");
  assert.deepEqual(resourcesWithAiSessionAccess([document, image, unsupported], true).map((item) => item.mcpAccess), ["text", "preview", "metadata"]);
});

test("turning AI session access off removes access from every Resource at once", () => {
  const resources = [
    resource({ id: "text", mcpAccess: "text" }),
    resource({ id: "image", mcpAccess: "preview" }),
    resource({ id: "metadata", mcpAccess: "metadata" }),
  ];
  assert.deepEqual(resourcesWithAiSessionAccess(resources, false).map((item) => item.mcpAccess), ["none", "none", "none"]);
});
