import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function text(relativePath: string) {
  return readFile(path.join(root, relativePath), "utf8");
}

test("macOS one-line installer is self-contained, verified, and policy-safe", async () => {
  const installer = await text("scripts/install-macos.sh");
  assert.match(installer, /^#!\/bin\/zsh/);
  assert.match(installer, /PRESENTATION_STUDIO_NODE_VERSION="v22\.13\.0"/);
  assert.match(installer, /nodejs\.org\/dist/);
  assert.match(installer, /SHASUMS256\.txt/);
  assert.match(installer, /shasum -a 256/);
  assert.match(installer, /npm ci/);
  assert.match(installer, /npm rebuild electron/);
  assert.match(installer, /require\("electron"\)/);
  assert.match(installer, /fs\.constants\.X_OK/);
  assert.match(installer, /npm test/);
  assert.match(installer, /npm run check:data-safety/);
  assert.match(installer, /npm run build/);
  assert.match(installer, /configure-mcp\.mjs" --codex/);
  assert.match(installer, /Restart Codex after installation/);
  assert.match(installer, /\.presentation-studio-managed-install/);
  assert.match(installer, /stop_managed_presentation_studio/);
  assert.match(installer, /mcp-runtime\.json/);
  assert.match(installer, /node_modules\/electron/);
  assert.doesNotMatch(installer, /pkill\s+-f/);
  assert.doesNotMatch(installer, /xattr\s+-d|spctl\s+--disable|csrutil\s+disable|sudo\s/);
});

test("Windows one-line installer is self-contained, verified, and policy-safe", async () => {
  const installer = await text("scripts/install-windows.ps1");
  assert.match(installer, /PRESENTATION_STUDIO_INSTALL_DIR/);
  assert.match(installer, /NodeVersion = "v22\.13\.0"/);
  assert.match(installer, /nodejs\.org\/dist/);
  assert.match(installer, /SHASUMS256\.txt/);
  assert.match(installer, /Get-FileHash -Algorithm SHA256/);
  assert.match(installer, /Invoke-Checked \$NpmCommand @\("ci"\)/);
  assert.match(installer, /Invoke-Checked \$NpmCommand @\("rebuild", "electron"\)/);
  assert.match(installer, /require\("electron"\)/);
  assert.match(installer, /fs\.constants\.X_OK/);
  assert.match(installer, /Invoke-Checked \$NpmCommand @\("test"\)/);
  assert.match(installer, /Invoke-Checked \$NpmCommand @\("run", "check:data-safety"\)/);
  assert.match(installer, /Invoke-Checked \$NpmCommand @\("run", "build"\)/);
  assert.match(installer, /configure-mcp\.mjs"\), "--codex"/);
  assert.match(installer, /Restart Codex after installation/);
  assert.match(installer, /\.presentation-studio-managed-install/);
  assert.match(installer, /Stop-ManagedPresentationStudio/);
  assert.match(installer, /mcp-runtime\.json/);
  assert.match(installer, /StartsWith\(\$InstallRoot/);
  assert.doesNotMatch(installer, /ExecutionPolicy\s+(Bypass|Unrestricted)|Set-ExecutionPolicy|DisableRealtimeMonitoring/);
});

test("installation docs expose complete one-line commands before manual Git setup", async () => {
  const readme = await text("README.md");
  const guide = await text("docs/INSTALLATION.md");
  const macCommand = "curl -fsSL https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-macos.sh | /bin/zsh";
  const windowsCommand = "irm https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-windows.ps1 | iex";
  for (const document of [readme, guide]) {
    assert.match(document, new RegExp(macCommand.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    assert.match(document, new RegExp(windowsCommand.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    assert.ok(document.indexOf(macCommand) < document.indexOf("git clone"));
    assert.match(document, /(does not require Git|Git is not required)/i);
    assert.match(document, /PowerPoint.*required for PowerPoint-native rendering/is);
    assert.match(document, /(PDF\.js|local PDF.*rasterizer)/i);
    assert.match(document, /(?:Poppler|pdftoppm).*not required/i);
  }
});

test("installation guide includes the ChatGPT Desktop starter prompt and native Resource composition workflow", async () => {
  const guide = await text("docs/INSTALLATION.md");
  const pdfBuilder = await text("scripts/build-install-guide.py");
  for (const document of [guide, pdfBuilder]) {
    assert.match(document, /Start a ChatGPT Desktop design session/);
    assert.match(document, /read the Presentation Studio design contract/i);
    assert.match(document, /assertion-evidence slides/);
    assert.match(document, /Found issues -> Fixing -> Rechecking original intent/);
    assert.match(document, /get_resource_text/);
    assert.match(document, /create_studio_presentation/);
    assert.match(document, /without a starter PowerPoint/i);
    assert.match(document, /single (?:\*\*)?AI access(?:\*\*)? switch/i);
  }
});

test("data-safety verification remains effective in a Git-free release archive", async () => {
  const fixtureRoot = await mkdtemp(path.join(os.tmpdir(), "presentation-studio-archive-safety-"));
  try {
    await mkdir(path.join(fixtureRoot, "scripts"));
    await copyFile(path.join(root, "scripts", "check-data-safety.mjs"), path.join(fixtureRoot, "scripts", "check-data-safety.mjs"));
    await writeFile(path.join(fixtureRoot, "README.md"), "safe release archive\n");

    const safe = spawnSync(process.execPath, [path.join(fixtureRoot, "scripts", "check-data-safety.mjs")], { encoding: "utf8" });
    assert.equal(safe.status, 0, safe.stderr);
    assert.match(safe.stdout, /source archive mode/);

    await writeFile(path.join(fixtureRoot, "client-deck.pptx"), "not a real presentation");
    const unsafe = spawnSync(process.execPath, [path.join(fixtureRoot, "scripts", "check-data-safety.mjs")], { encoding: "utf8" });
    assert.equal(unsafe.status, 1);
    assert.match(unsafe.stderr, /client-deck\.pptx/);
  } finally {
    await rm(fixtureRoot, { recursive: true, force: true });
  }
});
