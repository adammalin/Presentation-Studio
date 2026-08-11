/// <reference types="vite/client" />

import type { DesktopBridge } from "./lib/desktop";

declare global {
  interface Window {
    presentationStudioDesktop?: DesktopBridge;
  }
}

export {};
