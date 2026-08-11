import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { processResourceInput } from "../src/lib/resource-ingestion";

test("plain text is embedded with a bounded extracted-text derivative", async () => {
  const bytes = new TextEncoder().encode("Assertion\n\nEvidence from a synthetic local source.");
  const resource = await processResourceInput({ name: "synthetic-notes.md", bytes, mediaType: "text/markdown" });
  assert.equal(resource.kind, "document");
  assert.deepEqual(resource.roles, ["grounding-source"]);
  assert.equal(resource.processing?.status, "indexed");
  assert.equal(resource.derivatives?.length, 1);
  assert.equal(new TextDecoder().decode(resource.derivatives?.[0].bytes), "Assertion\n\nEvidence from a synthetic local source.");
});
test("DOCX text is extracted locally without changing the original bytes", async () => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("word/document.xml", "<w:document xmlns:w=\"urn:test\"><w:body><w:p><w:r><w:t>First paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Second &amp; final</w:t></w:r></w:p></w:body></w:document>");
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const resource = await processResourceInput({ name: "synthetic.docx", bytes });
  assert.equal(resource.processing?.status, "indexed");
  assert.equal(new TextDecoder().decode(resource.derivatives?.[0].bytes), "First paragraph\nSecond & final");
  assert.deepEqual(resource.bytes, bytes);
});

test("active SVG is embedded but held for review", async () => {
  const bytes = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script><rect width="10" height="10"/></svg>');
  const resource = await processResourceInput({ name: "unsafe.svg", bytes });
  assert.equal(resource.processing?.status, "needs-review");
  assert.deepEqual(resource.support, ["unsupported"]);
  assert.match(resource.processing?.warnings.join(" ") ?? "", /active|externally linked/i);
});

test("active program and script files are rejected", async () => {
  await assert.rejects(() => processResourceInput({ name: "do-not-run.js", bytes: new TextEncoder().encode("alert(1)") }), /cannot be added/i);
});
