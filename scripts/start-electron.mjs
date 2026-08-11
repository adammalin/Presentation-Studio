import { spawn } from "node:child_process";
import { createServer } from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function openPort() {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") return reject(new Error("Could not allocate a development port."));
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

function run(command, args, env = process.env) {
  return spawn(command, args, { cwd: root, env, stdio: "inherit", shell: process.platform === "win32" });
}

const port = await openPort();
const url = `http://127.0.0.1:${port}`;
const viteBin = path.join(root, "node_modules", "vite", "bin", "vite.js");
const electronBin = path.join(root, "node_modules", "electron", "cli.js");
const vite = run(process.execPath, [viteBin, "--host", "127.0.0.1", "--port", String(port), "--strictPort"]);

async function waitForVite() {
  for (let attempt = 0; attempt < 80; attempt += 1) {
    try { const response = await fetch(url); if (response.ok) return; } catch { /* still starting */ }
    await new Promise((resolve) => setTimeout(resolve, 150));
  }
  throw new Error("The Presentation Studio development server did not start in time.");
}

try {
  await waitForVite();
  const electron = run(process.execPath, [electronBin, "."], { ...process.env, PRESENTATION_STUDIO_DEV_URL: url });
  electron.on("exit", (code) => { vite.kill(); process.exitCode = code ?? 0; });
  process.on("SIGINT", () => { electron.kill(); vite.kill(); });
} catch (error) {
  vite.kill();
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
}
