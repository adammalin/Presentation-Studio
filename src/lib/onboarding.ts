export const ONBOARDING_TOUR_VERSION = "4";
export const ONBOARDING_TOUR_STORAGE_KEY = "presentation-studio:onboarding-tour";

export interface OnboardingTourStep {
  id: string;
  target: string;
  eyebrow: string;
  title: string;
  body: string;
}
export const ONBOARDING_TOUR_STEPS: readonly OnboardingTourStep[] = Object.freeze([
  Object.freeze({
    id: "welcome",
    target: "brand",
    eyebrow: "Welcome",
    title: "Meet Presentation Studio",
    body: "Presentation Studio keeps source files intact while you audit, clean, review, and eventually compose presentation work inside one portable project.",
  }),
  Object.freeze({
    id: "batch",
    target: "nav-batch",
    eyebrow: "Start here",
    title: "Import and audit PowerPoint decks",
    body: "Use Batch to add one or many PPTX files. Each deck is embedded as a read-only project Resource, structurally audited, and kept separate from every other job.",
  }),
  Object.freeze({
    id: "resources",
    target: "nav-resources",
    eyebrow: "Portable sources",
    title: "Create directly from project Resources",
    body: "Drop documents, data, images, media, SVGs, and presentations anywhere in the app. Accepted files and local derivatives stay embedded by hash. Share document Text or image Preview for the current AI session, and an MCP model can build a new source-grounded native Studio presentation without a starter PowerPoint.",
  }),
  Object.freeze({
    id: "inspect",
    target: "nav-slides",
    eyebrow: "See and collaborate",
    title: "Open the current slide design",
    body: "Slides renders the current embedded PowerPoint locally. Select a thumbnail for a close-up, then point to an exact region to save a private note or submit a scoped design comment to AI.",
  }),
  Object.freeze({
    id: "designs",
    target: "nav-designs",
    eyebrow: "Template library",
    title: "See the available slide designs",
    body: "Designs reads the locally installed PowerPoint template and shows its real masters, layouts, media, and placeholder structure. AI can instantiate these stable layout IDs in the central Studio JSON scene; browsing by itself never changes a slide.",
  }),
  Object.freeze({
    id: "review",
    target: "nav-review",
    eyebrow: "Human approval",
    title: "Review every proposed cleanup",
    body: "Rules can stage bounded changes, but Review is the approval gate. Nothing is accepted, saved, or exported merely because an AI or automated rule proposed it.",
  }),
  Object.freeze({
    id: "ai-session",
    target: "ai-session",
    eyebrow: "Local AI control",
    title: "AI access is session-based",
    body: "The local MCP bridge is ready while the app is open, but project access stays off until you enable AI session access. Each Resource separately cycles through Not shared, Metadata, and—when supported—Text or Preview sharing.",
  }),
  Object.freeze({
    id: "save",
    target: "save",
    eyebrow: "Portable projects",
    title: "Save standard or encrypted projects",
    body: "Save creates a self-contained PSTUDIO project. Save encrypted protects the packaged project; external originals and separately exported presentation files are not covered.",
  }),
  Object.freeze({
    id: "replay",
    target: "tour",
    eyebrow: "You are ready",
    title: "Replay this walkthrough anytime",
    body: "Use Tour in the top bar whenever you want a refresher. Skipping or completing this walkthrough only changes the local onboarding preference.",
  }),
]);

export interface TourRect {
  left: number;
  right: number;
  top: number;
  bottom: number;
  width: number;
  height: number;
}

export function shouldShowOnboardingTour(storedVersion: string | undefined | null): boolean {
  return String(storedVersion ?? "") !== ONBOARDING_TOUR_VERSION;
}

export function clampTourStep(index: number, stepCount = ONBOARDING_TOUR_STEPS.length): number {
  return Math.min(Math.max(0, Math.trunc(stepCount) - 1), Math.max(0, Math.trunc(index) || 0));
}

export function tourCardPosition(
  targetRect: TourRect | null,
  cardSize: { width: number; height: number },
  viewport: { width: number; height: number },
  options: { margin?: number; gap?: number } = {},
): { placement: "center" | "below" | "above" | "right" | "left"; left: number; top: number } {
  const margin = options.margin ?? 16;
  const gap = options.gap ?? 18;
  const viewportWidth = Math.max(0, viewport.width);
  const viewportHeight = Math.max(0, viewport.height);
  const cardWidth = Math.min(Math.max(0, cardSize.width), Math.max(0, viewportWidth - margin * 2));
  const cardHeight = Math.min(Math.max(0, cardSize.height), Math.max(0, viewportHeight - margin * 2));
  const maxLeft = Math.max(margin, viewportWidth - cardWidth - margin);
  const maxTop = Math.max(margin, viewportHeight - cardHeight - margin);
  const clampLeft = (value: number) => Math.min(maxLeft, Math.max(margin, value));
  const clampTop = (value: number) => Math.min(maxTop, Math.max(margin, value));

  if (!targetRect) return { placement: "center", left: clampLeft((viewportWidth - cardWidth) / 2), top: clampTop((viewportHeight - cardHeight) / 2) };
  const centeredLeft = clampLeft(targetRect.left + targetRect.width / 2 - cardWidth / 2);
  const centeredTop = clampTop(targetRect.top + targetRect.height / 2 - cardHeight / 2);
  if (targetRect.bottom + gap + cardHeight <= viewportHeight - margin) return { placement: "below", left: centeredLeft, top: targetRect.bottom + gap };
  if (targetRect.top - gap - cardHeight >= margin) return { placement: "above", left: centeredLeft, top: targetRect.top - gap - cardHeight };
  if (targetRect.right + gap + cardWidth <= viewportWidth - margin) return { placement: "right", left: targetRect.right + gap, top: centeredTop };
  if (targetRect.left - gap - cardWidth >= margin) return { placement: "left", left: targetRect.left - gap - cardWidth, top: centeredTop };
  return { placement: "center", left: clampLeft((viewportWidth - cardWidth) / 2), top: clampTop((viewportHeight - cardHeight) / 2) };
}
