import type { StudioWebNode, StudioWebScene } from "../types";
import { resolvedStudioTableDesign } from "./studio-web-scene";

const EMU_PER_INCH = 914_400;

export interface StudioDeckConsistencyIssue {
  id: string;
  category: "title-grid" | "component-type" | "table-system";
  severity: "major" | "minor";
  slideNumbers: number[];
  nodeIds: string[];
  message: string;
  recommendation: string;
}

export interface StudioDeckConsistencyReview {
  sceneRevision: string;
  designedSlideCount: number;
  repeatedComponentCount: number;
  tableCount: number;
  issueCount: number;
  issues: StudioDeckConsistencyIssue[];
}

function median(values: number[]): number {
  const ordered = [...values].sort((left, right) => left - right);
  const middle = Math.floor(ordered.length / 2);
  return ordered.length % 2 ? ordered[middle] : ((ordered[middle - 1] ?? 0) + (ordered[middle] ?? 0)) / 2;
}

function styleSignature(node: StudioWebNode): string {
  return [node.style.fontSizePt.toFixed(2), node.style.fontWeight, node.style.color.toUpperCase(), node.style.textAlign, node.style.verticalAlign].join("|");
}

function tableSignature(node: StudioWebNode): string {
  const design = resolvedStudioTableDesign(node);
  const padding = design.defaultPaddingPt;
  return [design.headerRows, design.borderMode, design.borderColor.toUpperCase(), design.borderWidthPt.toFixed(2), padding.top, padding.right, padding.bottom, padding.left, node.style.fontSizePt.toFixed(2)].join("|");
}

export function analyzeStudioDeckConsistency(scene: StudioWebScene): StudioDeckConsistencyReview {
  const designed = scene.slides.filter((slide) => slide.status === "designed" && slide.recipe !== "source");
  const issues: StudioDeckConsistencyIssue[] = [];
  const titles = designed
    .filter((slide) => slide.recipe !== "template-layout")
    .flatMap((slide) => slide.nodes.filter((node) => node.visible && node.role === "title" && node.kind === "text").map((node) => ({ slideNumber: slide.slideNumber, node })));
  if (titles.length >= 3) {
    const expectedX = median(titles.map(({ node }) => node.frame.x));
    const expectedY = median(titles.map(({ node }) => node.frame.y));
    const outliers = titles.filter(({ node }) => Math.abs(node.frame.x - expectedX) > .08 * EMU_PER_INCH || Math.abs(node.frame.y - expectedY) > .08 * EMU_PER_INCH);
    if (outliers.length) issues.push({ id: "deck-title-grid", category: "title-grid", severity: "major", slideNumbers: outliers.map((item) => item.slideNumber), nodeIds: outliers.map((item) => item.node.id), message: `${outliers.length} designed title${outliers.length === 1 ? " is" : "s are"} off the deck's dominant title grid.`, recommendation: "Align title optical starts through refine_studio_layout, then rebuild and inspect the native results." });
  }
  const roleGroups = new Map<string, Array<{ slideNumber: number; node: StudioWebNode }>>();
  for (const slide of designed) for (const node of slide.nodes) {
    if (!node.visible || node.kind !== "text" || !node.component?.role) continue;
    const entries = roleGroups.get(node.component.role) ?? [];
    entries.push({ slideNumber: slide.slideNumber, node });
    roleGroups.set(node.component.role, entries);
  }
  for (const [role, entries] of roleGroups) {
    if (entries.length < 3) continue;
    const counts = new Map<string, number>();
    entries.forEach(({ node }) => counts.set(styleSignature(node), (counts.get(styleSignature(node)) ?? 0) + 1));
    const dominant = [...counts.entries()].sort((left, right) => right[1] - left[1])[0]?.[0];
    const outliers = dominant ? entries.filter(({ node }) => styleSignature(node) !== dominant) : [];
    if (outliers.length) issues.push({ id: `component-${role}`, category: "component-type", severity: "minor", slideNumbers: [...new Set(outliers.map((item) => item.slideNumber))], nodeIds: outliers.map((item) => item.node.id), message: `${outliers.length} ${role.replaceAll("-", " ")} component${outliers.length === 1 ? " differs" : "s differ"} from the dominant deck style.`, recommendation: "Apply the established component typography instead of creating a one-off style." });
  }
  const tableGroups = new Map<string, Array<{ slideNumber: number; node: StudioWebNode }>>();
  for (const slide of designed) for (const node of slide.nodes) {
    if (!node.visible || node.kind !== "table" || !node.table) continue;
    const semanticRoles = [...new Set(node.table.cells.map((cell) => cell.semanticColorRole).filter(Boolean))].sort().join(",");
    const key = `${node.table.rows}x${node.table.columns}|roles:${semanticRoles}`;
    const entries = tableGroups.get(key) ?? [];
    entries.push({ slideNumber: slide.slideNumber, node });
    tableGroups.set(key, entries);
  }
  for (const [key, entries] of tableGroups) {
    if (entries.length < 2) continue;
    const signatures = new Map<string, number>();
    entries.forEach(({ node }) => signatures.set(tableSignature(node), (signatures.get(tableSignature(node)) ?? 0) + 1));
    const dominant = [...signatures.entries()].sort((left, right) => right[1] - left[1])[0]?.[0];
    const outliers = dominant ? entries.filter(({ node }) => tableSignature(node) !== dominant) : [];
    if (outliers.length) issues.push({ id: `table-${key}`, category: "table-system", severity: "major", slideNumbers: [...new Set(outliers.map((item) => item.slideNumber))], nodeIds: outliers.map((item) => item.node.id), message: `${outliers.length} related table${outliers.length === 1 ? " does" : "s do"} not use the deck's dominant structural style.`, recommendation: "Match border, padding, header, and type tokens while preserving each cell's semantic fill role." });
  }
  return { sceneRevision: scene.revision, designedSlideCount: designed.length, repeatedComponentCount: [...roleGroups.values()].filter((entries) => entries.length >= 2).reduce((sum, entries) => sum + entries.length, 0), tableCount: designed.flatMap((slide) => slide.nodes).filter((node) => node.visible && node.kind === "table").length, issueCount: issues.length, issues };
}
