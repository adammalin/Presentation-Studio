import type {
  StudioTableCellBorders,
  StudioTableCellDesign,
  StudioTableContinuationPlan,
  StudioTableExemplarDefinition,
  StudioTableRoleStyle,
  StudioWebNode,
  StudioWebScene,
  StudioWebSlide,
} from "../types";
import { STUDIO_WEB_SCENE_VERSION } from "../types";
import { resolvedStudioTableDesign } from "./studio-web-scene";

const HEX = /^#[0-9a-f]{6}$/i;

export interface StudioTableExemplarResult {
  scene: StudioWebScene;
  definition: StudioTableExemplarDefinition;
  affectedSlideNumbers: number[];
  affectedTableNodeIds: string[];
  skippedTableNodeIds: string[];
}

export interface StudioTableContinuationResult {
  scene: StudioWebScene;
  plan: StudioTableContinuationPlan;
}

export interface StudioTableCapacityAssessment {
  required: boolean;
  tableNodeId: string;
  totalCellCharacterCount: number;
  maximumCellCharacterCount: number;
  averageCellCharacterCount: number;
  bodyRowCount: number;
  recommendedMaximumBodyRowsPerSlide?: number;
  reason: string;
  nextTool?: "plan_studio_table_continuation";
}

export interface MaterializedStudioTableSlide {
  slide: StudioWebSlide;
  continuation?: {
    tableNodeId: string;
    segmentOrdinal: number;
    segmentCount: number;
    bodyRowStart: number;
    bodyRowEnd: number;
    repeatedHeaderRows: number;
  };
}

function nextRevision(scene: StudioWebScene, now: string): string {
  return `${scene.sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`;
}

function cloneBorders(value: StudioTableCellBorders | undefined): StudioTableCellBorders | undefined {
  if (!value) return undefined;
  return Object.fromEntries(Object.entries(value).map(([edge, border]) => [edge, border ? { ...border } : border])) as StudioTableCellBorders;
}

function cloneRoleStyle(value: StudioTableRoleStyle): StudioTableRoleStyle {
  return {
    ...value,
    paddingPt: value.paddingPt ? { ...value.paddingPt } : undefined,
    borders: cloneBorders(value.borders),
  };
}

function cloneCellStyle(value: StudioTableCellDesign): StudioTableCellDesign {
  return {
    ...value,
    paddingPt: value.paddingPt ? { ...value.paddingPt } : undefined,
    borders: cloneBorders(value.borders),
  };
}

function normalizedFill(value: string | undefined): string | undefined {
  return HEX.test(value ?? "") ? value!.toUpperCase() : undefined;
}

function rowStructure(node: StudioWebNode, row: number): string {
  return (node.table?.cells ?? [])
    .filter((cell) => cell.row === row)
    .sort((left, right) => left.column - right.column)
    .map((cell) => `${cell.column}:${cell.columnSpan}:${cell.rowSpan}`)
    .join("|");
}

function tableCompatibility(node: StudioWebNode): StudioTableExemplarDefinition["compatibility"] {
  if (node.kind !== "table" || !node.table) throw new Error("Choose a native Studio table before publishing or applying an exemplar.");
  const design = resolvedStudioTableDesign(node);
  return {
    columns: node.table.columns,
    headerRows: design.headerRows,
    headerStructure: Array.from({ length: design.headerRows }, (_, index) => rowStructure(node, index + 1)).join("/"),
    bodyStructure: design.headerRows < node.table.rows ? rowStructure(node, design.headerRows + 1) : "",
  };
}

function compatible(left: StudioTableExemplarDefinition["compatibility"], right: StudioTableExemplarDefinition["compatibility"]): boolean {
  return left.columns === right.columns
    && left.headerRows === right.headerRows
    && left.headerStructure === right.headerStructure
    && left.bodyStructure === right.bodyStructure;
}

function structureKey(value: StudioTableExemplarDefinition["compatibility"]): string {
  const source = `${value.columns}|${value.headerRows}|${value.headerStructure}|${value.bodyStructure}`;
  let hash = 2_166_136_261;
  for (let index = 0; index < source.length; index += 1) hash = Math.imul(hash ^ source.charCodeAt(index), 16_777_619) >>> 0;
  return hash.toString(16).padStart(8, "0");
}

function sourceCellStyle(node: StudioWebNode, row: number, fallback: StudioTableRoleStyle): StudioTableRoleStyle {
  if (!node.table) return fallback;
  const design = resolvedStudioTableDesign(node);
  const cell = node.table.cells
    .filter((candidate) => candidate.row === row && !candidate.semanticColorRole)
    .sort((left, right) => left.column - right.column)[0];
  const override = cell ? design.cellStyles.find((candidate) => candidate.cellId === cell.id) : undefined;
  return {
    fill: normalizedFill(override?.fill ?? cell?.fill ?? fallback.fill),
    color: normalizedFill(override?.color ?? fallback.color),
    fontSizePt: override?.fontSizePt ?? fallback.fontSizePt,
    fontWeight: override?.fontWeight ?? fallback.fontWeight,
    textAlign: override?.textAlign ?? fallback.textAlign,
    verticalAlign: override?.verticalAlign ?? fallback.verticalAlign,
    paddingPt: { ...(override?.paddingPt ?? design.defaultPaddingPt) },
    borders: cloneBorders(override?.borders),
  };
}

function definitionFor(node: StudioWebNode, slideNumber: number, name: string | undefined, now: string): StudioTableExemplarDefinition {
  if (!node.table) throw new Error("Choose a native Studio table before publishing an exemplar.");
  const design = resolvedStudioTableDesign(node);
  const firstBodyRow = Math.min(node.table.rows, design.headerRows + 1);
  const secondBodyRow = Math.min(node.table.rows, firstBodyRow + 1);
  const compatibility = tableCompatibility(node);
  return {
    id: `ornl-table-${compatibility.columns}c-${compatibility.headerRows}h-${structureKey(compatibility)}`,
    name: name?.trim().slice(0, 120) || `${node.table.columns}-column ORNL table`,
    sourceNodeId: node.id,
    adoptedFromSlideNumber: slideNumber,
    compatibility,
    tableStyle: {
      columnWidths: [...design.columnWidths],
      borderMode: design.borderMode,
      borderColor: design.borderColor,
      borderWidthPt: design.borderWidthPt,
      defaultPaddingPt: { ...design.defaultPaddingPt },
    },
    roleStyles: {
      header: sourceCellStyle(node, Math.max(1, design.headerRows), { fill: "#00454D", color: "#FFFFFF", fontSizePt: node.style.fontSizePt, fontWeight: 700, textAlign: node.style.textAlign, verticalAlign: "middle", paddingPt: { ...design.defaultPaddingPt } }),
      bodyOdd: sourceCellStyle(node, firstBodyRow, { fill: "#FFFFFF", color: node.style.color, fontSizePt: node.style.fontSizePt, fontWeight: 400, textAlign: node.style.textAlign, verticalAlign: node.style.verticalAlign, paddingPt: { ...design.defaultPaddingPt } }),
      bodyEven: sourceCellStyle(node, secondBodyRow, { fill: "#F0F2F1", color: node.style.color, fontSizePt: node.style.fontSizePt, fontWeight: 400, textAlign: node.style.textAlign, verticalAlign: node.style.verticalAlign, paddingPt: { ...design.defaultPaddingPt } }),
    },
    updatedAt: now,
  };
}

function tableLockedByFigure(slide: StudioWebSlide, nodeId: string): boolean {
  return slide.figureTreatments.some((treatment) => treatment.nodeIds.includes(nodeId) && treatment.verificationStatus === "source-locked");
}

function applyDefinitionToNode(node: StudioWebNode, definition: StudioTableExemplarDefinition): StudioWebNode {
  if (!node.table) return node;
  const design = resolvedStudioTableDesign(node);
  const existing = new Map(design.cellStyles.map((style) => [style.cellId, cloneCellStyle(style)]));
  for (const cell of node.table.cells) {
    const prior = existing.get(cell.id);
    const role = cell.row <= design.headerRows ? definition.roleStyles.header : (cell.row - design.headerRows) % 2 === 0 ? definition.roleStyles.bodyEven : definition.roleStyles.bodyOdd;
    const roleStyle = cloneRoleStyle(role);
    // A meaning-bearing source fill is data, not decoration. Exemplar reuse may
    // align typography, padding, and edges around it but can never recolor it.
    if (cell.semanticColorRole) delete roleStyle.fill;
    existing.set(cell.id, {
      ...prior,
      ...roleStyle,
      fill: cell.semanticColorRole ? prior?.fill : roleStyle.fill ?? prior?.fill,
      paddingPt: roleStyle.paddingPt ?? prior?.paddingPt,
      borders: roleStyle.borders ?? prior?.borders,
      cellId: cell.id,
    });
  }
  return {
    ...node,
    table: {
      ...node.table,
      design: {
        ...design,
        columnWidths: [...definition.tableStyle.columnWidths],
        borderMode: definition.tableStyle.borderMode,
        borderColor: definition.tableStyle.borderColor,
        borderWidthPt: definition.tableStyle.borderWidthPt,
        defaultPaddingPt: { ...definition.tableStyle.defaultPaddingPt },
        cellStyles: [...existing.values()],
      },
    },
  };
}

export function compatibleStudioTableInstances(scene: StudioWebScene, slideNumber: number, tableNodeId: string): Array<{ slideNumber: number; tableNodeId: string }> {
  const source = scene.slides.find((slide) => slide.slideNumber === slideNumber)?.nodes.find((node) => node.id === tableNodeId);
  if (!source?.table) return [];
  const signature = tableCompatibility(source);
  return scene.slides.flatMap((slide) => slide.nodes.flatMap((node) => node.visible && node.kind === "table" && node.table && compatible(signature, tableCompatibility(node))
    ? [{ slideNumber: slide.slideNumber, tableNodeId: node.id }]
    : []));
}

function applyDefinition(scene: StudioWebScene, definition: StudioTableExemplarDefinition, targetSlideNumbers?: number[]): StudioTableExemplarResult {
  const targetSet = targetSlideNumbers?.length ? new Set(targetSlideNumbers) : undefined;
  const now = new Date().toISOString();
  const affectedSlideNumbers = new Set<number>();
  const affectedTableNodeIds: string[] = [];
  const skippedTableNodeIds: string[] = [];
  const slides = scene.slides.map((slide) => {
    if (targetSet && !targetSet.has(slide.slideNumber)) return slide;
    let changed = false;
    const nodes = slide.nodes.map((node) => {
      if (!node.visible || node.kind !== "table" || !node.table || !compatible(definition.compatibility, tableCompatibility(node))) return node;
      if (node.locked || tableLockedByFigure(slide, node.id)) {
        skippedTableNodeIds.push(node.id);
        return node;
      }
      changed = true;
      affectedTableNodeIds.push(node.id);
      return applyDefinitionToNode(node, definition);
    });
    if (!changed) return slide;
    affectedSlideNumbers.add(slide.slideNumber);
    return { ...slide, nodes, qualityReview: undefined, updatedAt: now, designRationale: `Applied approved table exemplar ${definition.name} without copying source content or semantic fills.` };
  });
  if (!affectedTableNodeIds.length) throw new Error("No compatible unlocked tables were found for this exemplar's column, header, and merge structure.");
  const tableLibrary = [...(scene.tableLibrary ?? []).filter((item) => item.id !== definition.id), definition];
  return {
    scene: { ...scene, revision: nextRevision(scene, now), tableLibrary, slides },
    definition,
    affectedSlideNumbers: [...affectedSlideNumbers].sort((left, right) => left - right),
    affectedTableNodeIds,
    skippedTableNodeIds,
  };
}

export function publishStudioTableExemplar(scene: StudioWebScene, input: { slideNumber: number; tableNodeId: string; name?: string; targetSlideNumbers?: number[] }): StudioTableExemplarResult {
  const slide = scene.slides.find((candidate) => candidate.slideNumber === input.slideNumber);
  const node = slide?.nodes.find((candidate) => candidate.id === input.tableNodeId);
  if (!slide || !node?.table || node.kind !== "table") throw new Error("The exemplar table is not present in the current Studio scene revision.");
  if (node.locked || tableLockedByFigure(slide, node.id)) throw new Error("A locked or source-preserved table cannot define an editable deck exemplar.");
  const now = new Date().toISOString();
  const definition = definitionFor(node, input.slideNumber, input.name, now);
  return applyDefinition(scene, definition, input.targetSlideNumbers);
}

export function applyStudioTableExemplar(scene: StudioWebScene, input: { definitionId: string; targetSlideNumbers?: number[] }): StudioTableExemplarResult {
  const definition = scene.tableLibrary?.find((candidate) => candidate.id === input.definitionId);
  if (!definition) throw new Error("The requested approved table exemplar is not present in this Studio scene.");
  return applyDefinition(scene, definition, input.targetSlideNumbers);
}

type StudioTableCell = NonNullable<StudioWebNode["table"]>["cells"][number];

function repeatableContextCell(cell: StudioTableCell): boolean {
  return cell.column === 1
    && cell.columnSpan === 1
    && cell.rowSpan > 1
    && cell.text.trim().length > 0
    && cell.text.trim().length <= 240;
}

function bodyRowClusters(node: StudioWebNode, headerRows: number): Array<{ start: number; end: number }> {
  if (!node.table) return [];
  const result: Array<{ start: number; end: number }> = [];
  let cursor = headerRows + 1;
  while (cursor <= node.table.rows) {
    let end = cursor;
    let expanded = true;
    while (expanded) {
      expanded = false;
      for (const cell of node.table.cells) {
        // A concise leftmost merged label is repeated as continuation context
        // on each output slide. It must not force every body row into one
        // indivisible cluster; the detailed cells in the other columns remain
        // the actual merge-safe boundary authority.
        if (repeatableContextCell(cell)) continue;
        const cellEnd = cell.row + cell.rowSpan - 1;
        if (cell.row <= end && cellEnd >= cursor && cellEnd > end) {
          end = cellEnd;
          expanded = true;
        }
      }
    }
    result.push({ start: cursor, end: Math.min(node.table.rows, end) });
    cursor = end + 1;
  }
  return result;
}

/**
 * Deterministic capacity guidance for the agent before it spends a native
 * PowerPoint render on a table that cannot remain readable on one slide.
 * This is deliberately conservative: the final authority is still the exact
 * PowerPoint-native measurement after continuation materialization.
 */
export function assessStudioTableCapacity(node: StudioWebNode): StudioTableCapacityAssessment {
  if (!node.table || node.kind !== "table") throw new Error("Choose a native Studio table before assessing slide capacity.");
  const design = resolvedStudioTableDesign(node);
  const cellLengths = node.table.cells.map((cell) => cell.text.trim().length);
  const totalCellCharacterCount = cellLengths.reduce((sum, value) => sum + value, 0);
  const maximumCellCharacterCount = Math.max(0, ...cellLengths);
  const averageCellCharacterCount = cellLengths.length ? totalCellCharacterCount / cellLengths.length : 0;
  const bodyRowCount = Math.max(0, node.table.rows - design.headerRows);
  const required = bodyRowCount >= 2 && (totalCellCharacterCount > 1_600 || maximumCellCharacterCount > 260 || averageCellCharacterCount > 95);
  const recommendedMaximumBodyRowsPerSlide = required ? (averageCellCharacterCount > 80 || maximumCellCharacterCount > 220 ? 1 : Math.min(2, bodyRowCount - 1)) : undefined;
  return {
    required,
    tableNodeId: node.id,
    totalCellCharacterCount,
    maximumCellCharacterCount,
    averageCellCharacterCount: Math.round(averageCellCharacterCount * 10) / 10,
    bodyRowCount,
    recommendedMaximumBodyRowsPerSlide,
    reason: required
      ? `The table contains ${totalCellCharacterCount} characters across ${node.table.rows} rows (${maximumCellCharacterCount} in its longest cell). One-slide composition is not a valid candidate; preserve every cell in a merge-safe continuation before native preview.`
      : `The table's deterministic source inventory does not require continuation before its first native measurement (${totalCellCharacterCount} characters; ${maximumCellCharacterCount} in its longest cell).`,
    nextTool: required ? "plan_studio_table_continuation" : undefined,
  };
}

function planFor(node: StudioWebNode, slideNumber: number, maximumBodyRowsPerSlide: number, rationale: string | undefined, now: string): StudioTableContinuationPlan {
  if (!node.table || node.kind !== "table") throw new Error("Choose a native Studio table before planning a continuation.");
  const design = resolvedStudioTableDesign(node);
  if (!Number.isInteger(maximumBodyRowsPerSlide) || maximumBodyRowsPerSlide < 1 || maximumBodyRowsPerSlide > 40) throw new Error("A continuation segment must contain between 1 and 40 body rows.");
  const blockers: string[] = [];
  if (design.headerRows < 1) blockers.push("Identify at least one repeated header row before continuing this table.");
  for (const cell of node.table.cells) {
    const cellEnd = cell.row + cell.rowSpan - 1;
    if (cell.row <= design.headerRows && cellEnd > design.headerRows && !repeatableContextCell(cell)) blockers.push(`Cell ${cell.id} merges across the header/body boundary.`);
  }
  const clusters = bodyRowClusters(node, design.headerRows);
  const bodyRows = Math.max(0, node.table.rows - design.headerRows);
  if (bodyRows <= maximumBodyRowsPerSlide) blockers.push(`The table already fits within ${maximumBodyRowsPerSlide} body rows; a continuation is not needed.`);
  const segments: StudioTableContinuationPlan["segments"] = [];
  let current: { start: number; end: number } | undefined;
  for (const cluster of clusters) {
    const clusterSize = cluster.end - cluster.start + 1;
    if (clusterSize > maximumBodyRowsPerSlide) {
      blockers.push(`Rows ${cluster.start}-${cluster.end} form one merged unit larger than the requested segment size.`);
      continue;
    }
    if (!current) current = { ...cluster };
    else if (cluster.end - current.start + 1 <= maximumBodyRowsPerSlide) current.end = cluster.end;
    else {
      segments.push({ ordinal: segments.length + 1, bodyRowStart: current.start, bodyRowEnd: current.end, repeatedHeaderRows: design.headerRows, sourceCellIds: [] });
      current = { ...cluster };
    }
  }
  if (current) segments.push({ ordinal: segments.length + 1, bodyRowStart: current.start, bodyRowEnd: current.end, repeatedHeaderRows: design.headerRows, sourceCellIds: [] });
  for (const segment of segments) {
    segment.sourceCellIds = node.table.cells
      .filter((cell) => {
        const cellEnd = cell.row + cell.rowSpan - 1;
        const intersectsBodySegment = cell.row <= segment.bodyRowEnd && cellEnd >= segment.bodyRowStart;
        if (repeatableContextCell(cell) && intersectsBodySegment) return true;
        return cell.row <= design.headerRows && cellEnd <= design.headerRows
          || cell.row >= segment.bodyRowStart && cellEnd <= segment.bodyRowEnd;
      })
      .map((cell) => cell.id);
  }
  return {
    id: `table-continuation-${slideNumber}-${node.id}`,
    sourceSlideNumber: slideNumber,
    tableNodeId: node.id,
    headerRows: design.headerRows,
    maximumBodyRowsPerSlide,
    policy: "repeat-header-rows",
    status: blockers.length || segments.length < 2 ? "blocked" : "ready",
    segments,
    blockers: [...new Set(blockers)],
    rationale: rationale?.trim().slice(0, 1_000) || "Continue the native editable table at safe row boundaries and repeat its approved header semantics instead of shrinking below the ORNL minimum.",
    createdAt: now,
  };
}

export function planStudioTableContinuation(scene: StudioWebScene, input: { slideNumber: number; tableNodeId: string; maximumBodyRowsPerSlide: number; rationale?: string }): StudioTableContinuationResult {
  const slide = scene.slides.find((candidate) => candidate.slideNumber === input.slideNumber);
  const node = slide?.nodes.find((candidate) => candidate.id === input.tableNodeId);
  if (!slide || !node?.table || node.kind !== "table") throw new Error("The continuation table is not present in the current Studio scene revision.");
  if (node.locked || tableLockedByFigure(slide, node.id)) throw new Error("A locked or source-preserved table cannot be split into editable continuation slides.");
  const now = new Date().toISOString();
  let plan = planFor(node, input.slideNumber, input.maximumBodyRowsPerSlide, input.rationale, now);
  const visibleTables = slide.nodes.filter((candidate) => candidate.visible && candidate.kind === "table" && candidate.table);
  const orchestrationBlockers = [
    ...(slide.status !== "designed" || slide.recipe === "source" ? ["Choose and apply a Studio table or approved Template Pack layout before materializing a continuation."] : []),
    ...(visibleTables.length !== 1 ? [`Continuation materialization requires one primary visible table on the source slide; found ${visibleTables.length}.`] : []),
  ];
  if (orchestrationBlockers.length) plan = { ...plan, status: "blocked", blockers: [...new Set([...plan.blockers, ...orchestrationBlockers])] };
  return {
    plan,
    scene: {
      ...scene,
      revision: nextRevision(scene, now),
      tableContinuationPlans: [...(scene.tableContinuationPlans ?? []).filter((candidate) => candidate.tableNodeId !== node.id || candidate.sourceSlideNumber !== input.slideNumber), plan],
      slides: scene.slides.map((candidate) => candidate.slideNumber !== input.slideNumber ? candidate : { ...candidate, qualityReview: undefined, updatedAt: now }),
    },
  };
}

function normalizedWeights(values: number[]): number[] {
  const total = values.reduce((sum, value) => sum + value, 0);
  return total > 0 ? values.map((value) => value / total) : values.map(() => 1 / Math.max(1, values.length));
}

function materializedTableNode(node: StudioWebNode, plan: StudioTableContinuationPlan, segment: StudioTableContinuationPlan["segments"][number]): StudioWebNode {
  if (!node.table) return node;
  const included = new Set(segment.sourceCellIds);
  const design = resolvedStudioTableDesign(node);
  const selectedRows = [
    ...Array.from({ length: plan.headerRows }, (_, index) => index + 1),
    ...Array.from({ length: segment.bodyRowEnd - segment.bodyRowStart + 1 }, (_, index) => segment.bodyRowStart + index),
  ];
  const selectedRowHeights = normalizedWeights(selectedRows.map((row) => design.rowHeights[row - 1] ?? 1));
  const segmentBodyRows = segment.bodyRowEnd - segment.bodyRowStart + 1;
  const cells = node.table.cells.filter((cell) => included.has(cell.id)).map((cell) => {
    if (repeatableContextCell(cell)) {
      const crossesHeader = cell.row <= plan.headerRows;
      return {
        ...cell,
        row: crossesHeader ? cell.row : plan.headerRows + 1,
        rowSpan: crossesHeader ? selectedRows.length - cell.row + 1 : segmentBodyRows,
      };
    }
    return {
      ...cell,
      row: cell.row <= plan.headerRows ? cell.row : plan.headerRows + cell.row - segment.bodyRowStart + 1,
    };
  });
  return {
    ...node,
    table: {
      ...node.table,
      rows: selectedRows.length,
      cells,
      design: {
        ...design,
        rowHeights: selectedRowHeights,
        cellStyles: design.cellStyles.filter((style) => included.has(style.cellId)),
      },
    },
  };
}

export function materializeStudioTableContinuationSlides(scene: StudioWebScene, sourceSlide: StudioWebSlide): MaterializedStudioTableSlide[] {
  const plans = (scene.tableContinuationPlans ?? []).filter((plan) => plan.sourceSlideNumber === sourceSlide.slideNumber);
  if (!plans.length) return [{ slide: sourceSlide }];
  if (plans.some((plan) => plan.status === "blocked")) throw new Error(`Slide ${sourceSlide.slideNumber} has a blocked table continuation plan. Resolve or clear it before building PowerPoint.`);
  if (plans.length !== 1) throw new Error(`Slide ${sourceSlide.slideNumber} has ${plans.length} continuation plans. Materialize one primary table per source slide.`);
  const plan = plans[0];
  const table = sourceSlide.nodes.find((node) => node.id === plan.tableNodeId && node.kind === "table" && node.table);
  if (!table?.table) throw new Error(`Slide ${sourceSlide.slideNumber}'s continuation plan no longer maps to its source-bound table.`);
  if (resolvedStudioTableDesign(table).headerRows !== plan.headerRows) throw new Error(`Slide ${sourceSlide.slideNumber}'s header-row definition changed after continuation planning. Replan the table before building PowerPoint.`);
  return plan.segments.map((segment, index) => {
    const repeatedRoles = new Set(["eyebrow", "footer-logo", "footer-meta"]);
    const nodes = sourceSlide.nodes
      .filter((node) => index === 0 || node.id === table.id || node.role === "title" || Boolean(node.component?.role && repeatedRoles.has(node.component.role)))
      .map((node) => node.id === table.id ? materializedTableNode(node, plan, segment) : node);
    const nodeIds = new Set(nodes.map((node) => node.id));
    return {
      slide: {
        ...sourceSlide,
        id: `${sourceSlide.id}-table-continuation-${segment.ordinal}`,
        nodes,
        figureTreatments: index === 0 ? sourceSlide.figureTreatments : sourceSlide.figureTreatments.filter((treatment) => treatment.nodeIds.every((id) => nodeIds.has(id))),
        qualityReview: undefined,
        designRationale: `${sourceSlide.designRationale} Table segment ${segment.ordinal} of ${plan.segments.length} repeats ${plan.headerRows} header row${plan.headerRows === 1 ? "" : "s"} and carries body rows ${segment.bodyRowStart}-${segment.bodyRowEnd}.`,
      },
      continuation: {
        tableNodeId: table.id,
        segmentOrdinal: segment.ordinal,
        segmentCount: plan.segments.length,
        bodyRowStart: segment.bodyRowStart,
        bodyRowEnd: segment.bodyRowEnd,
        repeatedHeaderRows: segment.repeatedHeaderRows,
      },
    };
  });
}

export function clearStudioTableContinuation(scene: StudioWebScene, input: { slideNumber: number; tableNodeId: string }): StudioWebScene {
  const existing = scene.tableContinuationPlans?.some((plan) => plan.sourceSlideNumber === input.slideNumber && plan.tableNodeId === input.tableNodeId);
  if (!existing) return scene;
  const now = new Date().toISOString();
  return {
    ...scene,
    revision: nextRevision(scene, now),
    tableContinuationPlans: (scene.tableContinuationPlans ?? []).filter((plan) => plan.sourceSlideNumber !== input.slideNumber || plan.tableNodeId !== input.tableNodeId),
    slides: scene.slides.map((slide) => slide.slideNumber !== input.slideNumber ? slide : { ...slide, qualityReview: undefined, updatedAt: now }),
  };
}
