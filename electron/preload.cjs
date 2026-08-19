const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("presentationStudioDesktop", {
  isDesktop: true,
  platform: process.platform,
  versions: Object.freeze({ chrome: process.versions.chrome, electron: process.versions.electron }),
  pickPowerPoints: () => ipcRenderer.invoke("file:pick-powerpoints"),
  pickResources: () => ipcRenderer.invoke("file:pick-resources"),
  pickTemplate: () => ipcRenderer.invoke("template:pick-source"),
  installTemplate: (payload) => ipcRenderer.invoke("template:install", payload),
  getActiveTemplate: () => ipcRenderer.invoke("template:get-active"),
  openProject: () => ipcRenderer.invoke("file:open-project"),
  saveBinary: (payload) => ipcRenderer.invoke("file:save-binary", payload),
  autosaveProject: (payload) => ipcRenderer.invoke("project:autosave", payload),
  checkpointProjectState: (payload) => ipcRenderer.invoke("project:checkpoint-state", payload),
  getAutosaveStatus: () => ipcRenderer.invoke("project:get-autosave-status"),
  getAutosaveRecovery: () => ipcRenderer.invoke("project:get-autosave-recovery"),
  getMcpStatus: () => ipcRenderer.invoke("mcp:get-status"),
  getPresentationFonts: () => ipcRenderer.invoke("fonts:get-presentation-fonts"),
  getNativeRenderCapabilities: () => ipcRenderer.invoke("render:get-capabilities"),
  getNativeMeasurementCapabilities: () => ipcRenderer.invoke("measurement:get-capabilities"),
  renderPowerPoint: (payload) => ipcRenderer.invoke("render:powerpoint", payload),
  measurePowerPoint: (payload) => ipcRenderer.invoke("measurement:powerpoint", payload),
  cancelPowerPointMeasurement: (payload) => ipcRenderer.invoke("measurement:cancel", payload),
  captureDeckQualification: (payload) => ipcRenderer.invoke("qualification:capture", payload),
  finalizeDeckQualification: (payload) => ipcRenderer.invoke("qualification:finalize", payload),
  readDeckQualificationSlide: (payload) => ipcRenderer.invoke("qualification:read-slide", payload),
  revealDeckQualification: (payload) => ipcRenderer.invoke("qualification:reveal", payload),
  getOnboardingTourVersion: () => ipcRenderer.invoke("app:get-onboarding-tour-version"),
  setOnboardingTourVersion: (version) => ipcRenderer.invoke("app:set-onboarding-tour-version", version),
  openUserGuide: () => ipcRenderer.invoke("app:open-user-guide"),
  onMcpCommand: (handler) => {
    const listener = (_event, request) => {
      Promise.resolve()
        .then(() => handler(request))
        .then((result) => ipcRenderer.send("mcp:response", { id: request.id, ok: true, result }))
        .catch((error) => ipcRenderer.send("mcp:response", {
          id: request.id,
          ok: false,
          error: error instanceof Error ? error.message : "Presentation Studio rejected the MCP request.",
        }));
    };
    ipcRenderer.on("mcp:command", listener);
    return () => ipcRenderer.removeListener("mcp:command", listener);
  },
});
