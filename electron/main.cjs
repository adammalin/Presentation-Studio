const { app, BrowserWindow, dialog, ipcMain, session, shell } = require("electron");
const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const http = require("node:http");
const path = require("node:path");
const { createHash, randomBytes, randomUUID, timingSafeEqual } = require("node:crypto");
const { nativeRenderCapabilities, renderPowerPointNative } = require("./native-render.cjs");

app.setName("Presentation Studio");
const projectRoot = path.resolve(__dirname, "..");
const smokeTest = process.env.PRESENTATION_STUDIO_SMOKE_TEST === "1";
const capturePath = process.env.PRESENTATION_STUDIO_CAPTURE_PATH || "";
const captureView = process.env.PRESENTATION_STUDIO_CAPTURE_VIEW || "";
const captureResourceFixture = process.env.PRESENTATION_STUDIO_CAPTURE_RESOURCE_FIXTURE === "1";
const skipFirstRunTour = process.env.PRESENTATION_STUDIO_SKIP_FIRST_RUN_TOUR === "1";
if (smokeTest || capturePath) app.setPath("userData", path.join(app.getPath("temp"), `presentation-studio-isolated-${process.pid}`));

const MCP_RUNTIME_FILE_NAME = "mcp-runtime.json";
const APP_PREFERENCES_FILE_NAME = "preferences.json";
const ACTIVE_TEMPLATE_FILE_NAME = "active-template.json";
const ONBOARDING_TOUR_VERSION = "3";
const MCP_MAX_BODY_BYTES = 5_000_000;
const MAX_BINARY_BYTES = 1_250_000_000;
const mcpToken = randomBytes(32).toString("hex");
let mainWindow = null;
let mcpServer = null;
let mcpAddress = null;
const pendingMcpCommands = new Map();
let autosaveQueue = Promise.resolve();
let preferencesQueue = Promise.resolve();
let nativeRenderQueue = Promise.resolve();

function runtimePath() {
  return path.join(app.getPath("userData"), MCP_RUNTIME_FILE_NAME);
}

function recoveryPath(encrypted) {
  return path.join(app.getPath("userData"), "autosave", encrypted ? "current-project.pstudio-secure" : "current-project.pstudio");
}

function preferencesPath() {
  return path.join(app.getPath("userData"), APP_PREFERENCES_FILE_NAME);
}

function templatePackRoot() {
  return path.join(app.getPath("userData"), "template-packs");
}

function activeTemplateManifestPath() {
  return path.join(templatePackRoot(), ACTIVE_TEMPLATE_FILE_NAME);
}

function templateSourcePath(sha256, extension) {
  return path.join(templatePackRoot(), sha256, `source.${extension}`);
}

function privatePresentationFontRoots() {
  if (process.platform === "darwin") return ["/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"];
  if (process.platform === "win32") {
    const programFiles = process.env.ProgramFiles || "C:\\Program Files";
    return [path.join(programFiles, "Microsoft Office", "root", "vfs", "Fonts", "private")];
  }
  return [];
}

async function readLocalPresentationFonts() {
  const descriptors = [
    { family: "Aptos", fileName: "Aptos.ttf", weight: 400, style: "normal" },
    { family: "Aptos", fileName: "Aptos-Bold.ttf", weight: 700, style: "normal" },
    { family: "Aptos", fileName: "Aptos-Italic.ttf", weight: 400, style: "italic" },
    { family: "Aptos", fileName: "Aptos-Bold-Italic.ttf", weight: 700, style: "italic" },
    { family: "Aptos Light", fileName: "Aptos-Light.ttf", weight: 400, style: "normal" },
    { family: "Aptos Light", fileName: "Aptos-Light-Italic.ttf", weight: 400, style: "italic" },
    { family: "Aptos Narrow", fileName: "Aptos-Narrow.ttf", weight: 400, style: "normal" },
    { family: "Aptos Narrow", fileName: "Aptos-Narrow-Bold.ttf", weight: 700, style: "normal" },
  ];
  const fonts = [];
  for (const descriptor of descriptors) {
    for (const root of privatePresentationFontRoots()) {
      try {
        const filePath = path.join(root, descriptor.fileName);
        const bytes = await fs.readFile(filePath);
        if (bytes.byteLength === 0 || bytes.byteLength > 5 * 1024 * 1024) continue;
        fonts.push({ family: descriptor.family, weight: descriptor.weight, style: descriptor.style, mediaType: "font/ttf", bytes: new Uint8Array(bytes) });
        break;
      } catch (error) {
        if (error?.code !== "ENOENT") break;
      }
    }
  }
  return fonts;
}

async function readPreferences() {
  try {
    const bytes = await fs.readFile(preferencesPath());
    if (bytes.byteLength > 64 * 1024) throw new Error("Presentation Studio preferences are oversized.");
    const parsed = JSON.parse(bytes.toString("utf8"));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch (error) {
    if (error?.code === "ENOENT") return {};
    return {};
  }
}

function tokenMatches(value) {
  if (typeof value !== "string") return false;
  const supplied = Buffer.from(value);
  const expected = Buffer.from(mcpToken);
  return supplied.length === expected.length && timingSafeEqual(supplied, expected);
}

function jsonResponse(response, status, body) {
  const payload = Buffer.from(JSON.stringify(body));
  response.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": payload.length,
    "cache-control": "no-store",
  });
  response.end(payload);
}

async function readJsonBody(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MCP_MAX_BODY_BYTES) throw new Error("The MCP request exceeds the 5 MB limit.");
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
  } catch {
    throw new Error("The MCP request body is not valid JSON.");
  }
}

function dispatchMcpCommand(operation, input) {
  if (!mainWindow || mainWindow.isDestroyed() || mainWindow.webContents.isLoading()) {
    return Promise.reject(new Error("Open Presentation Studio and wait for the project to finish loading."));
  }
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingMcpCommands.delete(id);
      reject(new Error("Presentation Studio did not answer the MCP request in time."));
    }, 15_000);
    pendingMcpCommands.set(id, { resolve, reject, timer });
    mainWindow.webContents.send("mcp:command", { id, operation, input });
  });
}

async function handleMcpRequest(request, response) {
  const remote = request.socket.remoteAddress;
  if (remote !== "127.0.0.1" && remote !== "::ffff:127.0.0.1" && remote !== "::1") {
    jsonResponse(response, 403, { error: "Only loopback MCP requests are accepted." });
    return;
  }
  if (!tokenMatches(request.headers["x-presentation-studio-token"])) {
    jsonResponse(response, 403, { error: "Presentation Studio rejected the desktop session token." });
    return;
  }
  const url = new URL(request.url || "/", "http://127.0.0.1");
  if (request.method !== "POST" || url.pathname !== "/command") {
    jsonResponse(response, 404, { error: "Unknown local MCP endpoint." });
    return;
  }
  try {
    const body = await readJsonBody(request);
    if (typeof body.operation !== "string" || !body.operation.trim()) throw new Error("An MCP operation is required.");
    jsonResponse(response, 200, { result: await dispatchMcpCommand(body.operation, body.input ?? {}) });
  } catch (error) {
    jsonResponse(response, 400, { error: error instanceof Error ? error.message.slice(0, 700) : "The MCP request failed." });
  }
}

async function startMcpBridge() {
  mcpServer = http.createServer((request, response) => void handleMcpRequest(request, response));
  mcpServer.on("clientError", (_error, socket) => socket.destroy());
  await new Promise((resolve, reject) => {
    mcpServer.once("error", reject);
    mcpServer.listen(0, "127.0.0.1", resolve);
  });
  const address = mcpServer.address();
  if (!address || typeof address === "string") throw new Error("The local MCP bridge did not receive a port.");
  mcpAddress = `http://127.0.0.1:${address.port}/`;
  const target = runtimePath();
  const temporary = `${target}.${process.pid}.tmp`;
  await fs.mkdir(path.dirname(target), { recursive: true, mode: 0o700 });
  await fs.writeFile(temporary, `${JSON.stringify({ version: 1, pid: process.pid, baseUrl: mcpAddress, token: mcpToken, startedAt: new Date().toISOString() })}\n`, { encoding: "utf8", mode: 0o600 });
  await fs.rename(temporary, target);
  try { await fs.chmod(target, 0o600); } catch { /* Windows has no POSIX permission bits. */ }
}

function stopMcpBridge() {
  for (const pending of pendingMcpCommands.values()) {
    clearTimeout(pending.timer);
    pending.reject(new Error("Presentation Studio is closing."));
  }
  pendingMcpCommands.clear();
  if (mcpServer) mcpServer.close();
  mcpServer = null;
  try {
    const descriptor = JSON.parse(fsSync.readFileSync(runtimePath(), "utf8"));
    if (descriptor?.pid === process.pid) fsSync.rmSync(runtimePath(), { force: true });
  } catch { /* A future start replaces an unreadable or stale descriptor. */ }
}

function validateBinary(value) {
  const bytes = Buffer.from(value ?? []);
  if (bytes.byteLength === 0) throw new Error("Presentation Studio received an empty file.");
  if (bytes.byteLength > MAX_BINARY_BYTES) throw new Error("The file exceeds the 1.25 GB local safety limit.");
  return bytes;
}

async function atomicWrite(target, bytes) {
  await fs.mkdir(path.dirname(target), { recursive: true });
  const temporary = `${target}.${process.pid}.${randomUUID()}.tmp`;
  await fs.writeFile(temporary, bytes, { mode: 0o600 });
  await fs.rename(temporary, target);
}

function filtersFor(kind) {
  if (kind === "pptx") return [{ name: "PowerPoint presentations", extensions: ["pptx"] }];
  if (kind === "secure-project") return [{ name: "Encrypted Presentation Studio projects", extensions: ["pstudio-secure"] }];
  if (kind === "project") return [{ name: "Presentation Studio projects", extensions: ["pstudio", "pstudio-secure"] }];
  if (kind === "report") return [{ name: "JSON audit reports", extensions: ["json"] }];
  return [{ name: "All files", extensions: ["*"] }];
}

function resourceFilters() {
  return [
    { name: "Presentation Studio resources", extensions: ["pptx", "potx", "docx", "pdf", "md", "txt", "csv", "tsv", "json", "xlsx", "png", "jpg", "jpeg", "webp", "tif", "tiff", "svg", "wav", "mp3", "m4a", "mp4", "mov", "doc", "xls"] },
    { name: "All files", extensions: ["*"] },
  ];
}

async function readPickedFiles(filePaths) {
  return Promise.all(filePaths.map(async (filePath) => ({
    name: path.basename(filePath),
    filePath,
    bytes: new Uint8Array(await fs.readFile(filePath)),
  })));
}

function installNetworkPolicy() {
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    let allowed = false;
    try {
      const url = new URL(details.url);
      allowed = ["file:", "devtools:", "chrome-extension:", "data:", "blob:"].includes(url.protocol)
        || (["http:", "ws:"].includes(url.protocol) && ["127.0.0.1", "localhost"].includes(url.hostname));
    } catch { allowed = false; }
    callback({ cancel: !allowed });
  });
}

function registerIpc() {
  ipcMain.on("mcp:response", (event, payload) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return;
    const pending = pendingMcpCommands.get(payload?.id);
    if (!pending) return;
    pendingMcpCommands.delete(payload.id);
    clearTimeout(pending.timer);
    if (payload.ok) pending.resolve(payload.result);
    else pending.reject(new Error(typeof payload.error === "string" ? payload.error : "Presentation Studio rejected the MCP request."));
  });

  ipcMain.handle("mcp:get-status", () => ({ available: Boolean(mcpAddress), address: mcpAddress, runtimeFile: runtimePath() }));

  ipcMain.handle("fonts:get-presentation-fonts", async () => ({ fonts: await readLocalPresentationFonts() }));

  ipcMain.handle("render:get-capabilities", () => nativeRenderCapabilities());

  ipcMain.handle("render:powerpoint", async (_event, payload) => {
    const name = path.basename(String(payload?.name ?? "presentation.pptx"));
    if (!/\.pptx$/i.test(name)) throw new Error("PowerPoint-native rendering requires a PPTX file.");
    const bytes = validateBinary(payload?.bytes);
    nativeRenderQueue = nativeRenderQueue.catch(() => undefined).then(() => renderPowerPointNative({ bytes, name, homePath: app.getPath("home") }));
    return nativeRenderQueue;
  });

  ipcMain.handle("app:get-onboarding-tour-version", async () => {
    if (skipFirstRunTour) return { version: ONBOARDING_TOUR_VERSION };
    const preferences = await readPreferences();
    return { version: typeof preferences.onboardingTourVersion === "string" ? preferences.onboardingTourVersion : null };
  });

  ipcMain.handle("app:set-onboarding-tour-version", async (_event, version) => {
    if (typeof version !== "string" || !/^[a-z0-9._-]{1,32}$/i.test(version)) throw new Error("The onboarding tour version is invalid.");
    preferencesQueue = preferencesQueue.catch(() => undefined).then(async () => {
      const preferences = await readPreferences();
      await atomicWrite(preferencesPath(), Buffer.from(`${JSON.stringify({ ...preferences, onboardingTourVersion: version }, null, 2)}\n`));
    });
    await preferencesQueue;
    return { saved: true, version };
  });

  ipcMain.handle("file:pick-powerpoints", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Add PowerPoint decks for read-only audit",
      properties: ["openFile", "multiSelections"],
      filters: filtersFor("pptx"),
    });
    if (result.canceled) return { canceled: true, files: [] };
    const files = await readPickedFiles(result.filePaths);
    return { canceled: false, files };
  });

  ipcMain.handle("file:pick-resources", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Add files to project Resources",
      properties: ["openFile", "multiSelections"],
      filters: resourceFilters(),
    });
    if (result.canceled) return { canceled: true, files: [] };
    return { canceled: false, files: await readPickedFiles(result.filePaths) };
  });

  ipcMain.handle("template:pick-source", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Install an authorized PowerPoint template",
      properties: ["openFile"],
      filters: [{ name: "PowerPoint templates", extensions: ["potx", "pptx"] }],
    });
    if (result.canceled || !result.filePaths[0]) return { canceled: true };
    const [file] = await readPickedFiles([result.filePaths[0]]);
    return { canceled: false, file };
  });

  ipcMain.handle("template:install", async (_event, payload) => {
    const name = path.basename(String(payload?.name ?? ""));
    const extension = path.extname(name).slice(1).toLowerCase();
    if (!name || !["potx", "pptx"].includes(extension)) throw new Error("Choose a POTX or PPTX template.");
    const bytes = validateBinary(payload?.bytes);
    if (bytes.byteLength > 250 * 1024 * 1024) throw new Error("The template exceeds the 250 MB installation limit.");
    const digest = createHash("sha256").update(bytes).digest("hex");
    if (digest !== String(payload?.sha256 ?? "")) throw new Error("The installed template hash does not match the previewed template.");
    const installedAt = new Date().toISOString();
    const sourcePath = templateSourcePath(digest, extension);
    await atomicWrite(sourcePath, bytes);
    await atomicWrite(activeTemplateManifestPath(), Buffer.from(`${JSON.stringify({ version: 1, name, sha256: digest, extension, installedAt }, null, 2)}\n`));
    return { installed: true, name, sha256: digest, installedAt };
  });

  ipcMain.handle("template:get-active", async () => {
    try {
      const manifestBytes = await fs.readFile(activeTemplateManifestPath());
      if (manifestBytes.byteLength > 64 * 1024) throw new Error("The active template manifest is oversized.");
      const manifest = JSON.parse(manifestBytes.toString("utf8"));
      if (!manifest || manifest.version !== 1 || !/^[0-9a-f]{64}$/.test(manifest.sha256) || !["potx", "pptx"].includes(manifest.extension)) throw new Error("The active template manifest is invalid.");
      const bytes = await fs.readFile(templateSourcePath(manifest.sha256, manifest.extension));
      if (createHash("sha256").update(bytes).digest("hex") !== manifest.sha256) throw new Error("The installed template no longer matches its manifest.");
      return { installed: true, name: manifest.name, sha256: manifest.sha256, installedAt: manifest.installedAt, bytes: new Uint8Array(bytes) };
    } catch (error) {
      if (error?.code === "ENOENT") return { installed: false };
      throw error;
    }
  });

  ipcMain.handle("file:open-project", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Open Presentation Studio project",
      properties: ["openFile"],
      filters: filtersFor("project"),
    });
    if (result.canceled || !result.filePaths[0]) return { canceled: true };
    const filePath = result.filePaths[0];
    return { canceled: false, file: { name: path.basename(filePath), filePath, bytes: new Uint8Array(await fs.readFile(filePath)) } };
  });

  ipcMain.handle("file:save-binary", async (_event, payload) => {
    const kind = ["project", "secure-project", "pptx", "report"].includes(payload?.kind) ? payload.kind : "project";
    const result = await dialog.showSaveDialog(mainWindow, {
      title: kind === "pptx" ? "Export a new cleaned PowerPoint copy" : "Save Presentation Studio project",
      defaultPath: String(payload?.defaultName || "Presentation Studio project.pstudio"),
      filters: filtersFor(kind),
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    await atomicWrite(result.filePath, validateBinary(payload?.bytes));
    return { canceled: false, filePath: result.filePath };
  });

  ipcMain.handle("project:autosave", async (_event, payload) => {
    const bytes = validateBinary(payload?.bytes);
    const target = recoveryPath(Boolean(payload?.encrypted));
    autosaveQueue = autosaveQueue.catch(() => undefined).then(() => atomicWrite(target, bytes));
    await autosaveQueue;
    return { recoveryPath: target };
  });

  ipcMain.handle("app:open-user-guide", async () => {
    const guidePath = path.join(projectRoot, "output", "pdf", "Presentation-Studio-Installation-Guide.pdf");
    try {
      await fs.access(guidePath);
      const error = await shell.openPath(guidePath);
      if (error) throw new Error(error);
      return { opened: true, path: guidePath };
    } catch {
      return { opened: false };
    }
  });
}

async function qualifyOnboardingTour() {
  const result = await mainWindow.webContents.executeJavaScript(`(async () => {
    const wait = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));
    const nextPaint = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    let tour = null;
    for (let attempt = 0; attempt < 60; attempt += 1) {
      tour = document.querySelector('.onboarding-tour');
      if (tour) break;
      await wait(50);
    }
    if (!tour) return { passed: false, error: 'First-run tour did not open.', layouts: [] };
    const count = Number(tour.dataset.stepCount || 0);
    const layouts = [];
    for (let index = 0; index < count; index += 1) {
      for (let attempt = 0; attempt < 30 && Number(tour.dataset.stepIndex) !== index; attempt += 1) await wait(20);
      await nextPaint();
      if (index === 0) await wait(350);
      const targetName = tour.dataset.target;
      const target = document.querySelector('[data-tour="' + targetName + '"]');
      const spotlight = document.querySelector('[data-tour-spotlight]');
      const card = document.querySelector('.tour-card');
      const targetRect = target?.getBoundingClientRect();
      const spotlightRect = spotlight?.getBoundingClientRect();
      const cardRect = card?.getBoundingClientRect();
      const expectedLeft = targetRect ? Math.max(6, targetRect.left - 8) : null;
      const expectedTop = targetRect ? Math.max(6, targetRect.top - 8) : null;
      const expectedRight = targetRect ? Math.min(innerWidth - 6, targetRect.right + 8) : null;
      const expectedBottom = targetRect ? Math.min(innerHeight - 6, targetRect.bottom + 8) : null;
      const spotlightAligned = Boolean(targetRect && spotlightRect &&
        Math.abs(spotlightRect.left - expectedLeft) <= 6 &&
        Math.abs(spotlightRect.top - expectedTop) <= 6 &&
        Math.abs(spotlightRect.right - expectedRight) <= 6 &&
        Math.abs(spotlightRect.bottom - expectedBottom) <= 6);
      const cardInViewport = Boolean(cardRect && cardRect.left >= 10 && cardRect.top >= 10 && cardRect.right <= innerWidth - 10 && cardRect.bottom <= innerHeight - 10);
      layouts.push({ index, id: tour.dataset.stepId, target: targetName, targetVisible: Boolean(targetRect?.width && targetRect?.height), spotlightAligned, cardInViewport });
      const next = document.querySelector('.tour-actions .button.primary');
      if (!next) return { passed: false, error: 'Tour next button is missing.', layouts };
      next.click();
      await wait(70);
      tour = document.querySelector('.onboarding-tour');
      if (index < count - 1 && !tour) return { passed: false, error: 'Tour closed before its final step.', layouts };
    }
    return { passed: layouts.length === count && layouts.every((item) => item.targetVisible && item.spotlightAligned && item.cardInViewport), layouts };
  })()`);
  console.log(`PRESENTATION_STUDIO_TOUR_SMOKE ${JSON.stringify(result)}`);
  if (!result?.passed) throw new Error(result?.error || "The onboarding tour layout qualification failed.");
  let storedVersion = null;
  for (let attempt = 0; attempt < 30; attempt += 1) {
    storedVersion = (await readPreferences()).onboardingTourVersion ?? null;
    if (storedVersion === ONBOARDING_TOUR_VERSION) break;
    await new Promise((resolve) => setTimeout(resolve, 30));
  }
  if (storedVersion !== ONBOARDING_TOUR_VERSION) throw new Error("The onboarding tour completion preference was not persisted.");
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 940,
    minWidth: 1120,
    minHeight: 720,
    backgroundColor: "#f3f5f4",
    titleBarStyle: process.platform === "darwin" ? "hiddenInset" : "default",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: !smokeTest,
    },
  });
  const devUrl = process.env.PRESENTATION_STUDIO_DEV_URL;
  if (devUrl) await mainWindow.loadURL(devUrl);
  else await mainWindow.loadFile(path.join(projectRoot, "dist", "index.html"));
  if (capturePath || smokeTest) {
    let ready = false;
    for (let attempt = 0; attempt < 50; attempt += 1) {
      ready = await mainWindow.webContents.executeJavaScript("Boolean(document.querySelector('.app-shell') && document.body.innerText.includes('Presentation Studio'))").catch(() => false);
      if (ready) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    if (!ready) throw new Error("The Electron renderer did not paint the Presentation Studio app shell in time.");
  }
  if (smokeTest) await qualifyOnboardingTour();
  if (capturePath) {
    if (captureView) {
      const selected = await mainWindow.webContents.executeJavaScript(`(() => { const label = ${JSON.stringify(captureView)}; const button = [...document.querySelectorAll('.rail-items > button')].find((item) => item.textContent?.trim() === label); if (!button) return false; button.click(); return true; })()`);
      if (!selected) throw new Error(`The capture view ${captureView} was not found.`);
      await new Promise((resolve) => setTimeout(resolve, 150));
    }
    if (captureResourceFixture) {
      const added = await mainWindow.webContents.executeJavaScript(`(() => { const input = document.querySelector('#web-resource-picker'); if (!input) return false; const transfer = new DataTransfer(); transfer.items.add(new File(['Synthetic assertion\\n\\nSynthetic evidence for local UI qualification.'], 'synthetic-source.md', { type: 'text/markdown' })); input.files = transfer.files; input.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
      if (!added) throw new Error("The synthetic Resource capture fixture could not be added.");
      for (let attempt = 0; attempt < 40; attempt += 1) {
        const ready = await mainWindow.webContents.executeJavaScript("document.querySelectorAll('.resource-row:not(.resource-head)').length > 0");
        if (ready) break;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
    await mainWindow.webContents.capturePage().then((image) => fs.writeFile(capturePath, image.toPNG()));
  }
  if (smokeTest || capturePath) setTimeout(() => app.quit(), 900);
  mainWindow.on("closed", () => { mainWindow = null; });
}

app.whenReady().then(async () => {
  installNetworkPolicy();
  registerIpc();
  await startMcpBridge();
  await createWindow();
  app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) void createWindow(); });
}).catch((error) => {
  console.error(error);
  app.exit(1);
});

app.on("before-quit", stopMcpBridge);
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
process.once("SIGINT", () => app.quit());
process.once("SIGTERM", () => app.quit());
