import assert from "node:assert/strict";
import test from "node:test";
import { decryptProjectPackage, encryptProjectPackage, isEncryptedProject } from "../src/lib/encryption";
import { sha256 } from "../src/lib/hash";
import { createProject } from "../src/lib/project";
import { buildProjectPackage, openProjectPackage } from "../src/lib/project-package";

test("self-contained project package round-trips embedded Resource bytes", async () => {
  const project = createProject("Synthetic package test");
  const bytes = new TextEncoder().encode("synthetic resource bytes only");
  const extractedBytes = new TextEncoder().encode("locally extracted synthetic text");
  project.resources.push({
    id: "resource-one",
    name: "synthetic.txt",
    mediaType: "text/plain",
    byteLength: bytes.byteLength,
    sha256: await sha256(bytes),
    roles: ["grounding-source"],
    kind: "document",
    support: ["source-readable"],
    processing: { status: "indexed", summary: "Synthetic extraction complete.", processedAt: new Date().toISOString(), warnings: [] },
    derivatives: [{ id: "derivative-one", kind: "extracted-text", mediaType: "text/plain", byteLength: extractedBytes.byteLength, sha256: await sha256(extractedBytes), createdAt: new Date().toISOString(), processor: "test/text-v1", truncated: false, bytes: extractedBytes }],
    createdAt: new Date().toISOString(),
    sourcePath: "/not/a/dependency/synthetic.txt",
    embedded: true,
    bytes,
    mcpAccess: "none",
  });
  const packed = await buildProjectPackage(project);
  const opened = await openProjectPackage(packed);
  assert.equal(opened.resources.length, 1);
  assert.deepEqual(opened.resources[0].bytes, bytes);
  assert.deepEqual(opened.resources[0].derivatives?.[0].bytes, extractedBytes);
  assert.equal(opened.resources[0].sourcePath, undefined);
  assert.equal(opened.resources[0].mcpAccess, "none");
});

test("encrypted package rejects a wrong password and decrypts with the right one", async () => {
  const plain = new TextEncoder().encode("synthetic package payload");
  const secure = await encryptProjectPackage(plain, "a sufficiently long test password");
  assert.equal(isEncryptedProject(secure), true);
  await assert.rejects(() => decryptProjectPackage(secure, "this password is wrong"), /incorrect|changed/i);
  assert.deepEqual(await decryptProjectPackage(secure, "a sufficiently long test password"), plain);
});

test("encrypted project requires a non-trivial password", async () => {
  await assert.rejects(() => encryptProjectPackage(new Uint8Array([1, 2, 3]), "short"), /at least 12/);
});
