import JSZip from "jszip";
import type {
  CleanupChange,
  CleanupProposal,
  DecorativeShapeCommand,
  DeckJob,
  GeometryEditCommand,
  LayoutReviewItem,
  NativeLayoutRemapCommand,
  SlideEditableObject,
  SlideDesignDisposition,
  TableInventoryItem,
  TableNormalizationException,
  TextStyleCommand,
} from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { auditPptx } from "./pptx-audit";
import { cloneTemplateLayoutForSlide, type NativeLayoutCloneReceipt } from "./native-layout-remap";

const TABLE_PROFILE = PRESENTATION_DESIGN_STANDARD.tableProfile;
const TABLE_MARGIN_HORIZONTAL_EMU = Math.round(TABLE_PROFILE.cellPaddingPt.left * 12_700);
const TABLE_MARGIN_VERTICAL_EMU = Math.round(TABLE_PROFILE.cellPaddingPt.top * 12_700);
const TABLE_RULE_WIDTH_EMU = Math.round(TABLE_PROFILE.strokes.horizontal.widthPt * 12_700);
const TABLE_FONT_SIZE = Math.round(TABLE_PROFILE.body.fontSizePt * 100);

function stableChangeId(from: string, to: string): string {
  return `font-${from.toLowerCase().replace(/[^a-z0-9]+/g, "-")}-${to.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
}

function assertDeckReady(deck: DeckJob) {
  if (!deck.audit) throw new Error("Audit the deck before staging cleanup.");
  if (!deck.targetTemplateConfirmedAt || !deck.targetTemplateId) throw new Error("Confirm the target template before staging cleanup.");
  if (deck.audit.containsMacros || deck.audit.containsOleObjects || deck.audit.containsExternalRelationships) {
    throw new Error("Advanced or externally linked content requires manual review before automated cleanup.");
  }
}

function fontCleanupChanges(deck: DeckJob): CleanupChange[] {
  return (deck.audit?.fonts ?? [])
    .filter((font) => font.directSlideCount > 0 && ["century gothic", "arial"].includes(font.normalizedFamily) && !font.isLikelySymbolFont)
    .map((font) => ({
      id: stableChangeId(font.family, "Aptos"),
      kind: "font-family" as const,
      from: font.family,
      to: "Aptos",
      affectedSlideNumbers: font.slideNumbers,
      affectedRunCount: font.directSlideCount,
      rationale: `Normalize legacy ${font.family} markup to the confirmed ORNL Aptos typography while preserving every text string.`,
      selected: true,
    }));
}

function tableException(table: TableInventoryItem): TableNormalizationException | undefined {
  const semanticTokens = table.colorTokens.filter((token) => /^accent[1-6]$/.test(token));
  if (semanticTokens.length > 0) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "semantic-color",
    reason: `Preserved meaning-bearing theme color (${semanticTokens.join(", ")}); this table needs a designer check before normalization.`,
  };
  const cellCount = table.rowCount * table.columnCount;
  if (table.mergedCellCount > 10 || (cellCount > 0 && table.mergedCellCount / cellCount > 0.35)) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "complex-structure",
    reason: "Preserved a complex merged-cell topology; normalize it only after a designer confirms hierarchy and reading order.",
  };
  const averageCharacters = cellCount > 0 ? table.totalCellCharacterCount / cellCount : 0;
  if (cellCount > 40 || table.maximumCellCharacterCount > 180 || averageCharacters > 45) return {
    tableId: table.id,
    slideNumber: table.slideNumber,
    rule: "dense-table",
    reason: `Preserved a dense technical table (${table.totalCellCharacterCount} characters; ${table.maximumCellCharacterCount} in its longest cell); it needs measured overflow and possible continuation-slide review.`,
  };
  return undefined;
}

function slideDispositions(deck: DeckJob, changes: CleanupChange[], exceptions: TableNormalizationException[], layoutExceptions: LayoutReviewItem[]): SlideDesignDisposition[] {
  return (deck.audit?.slides ?? []).map((slide) => {
    const changeIds = changes.filter((change) => change.affectedSlideNumbers.includes(slide.number)).map((change) => change.id);
    const slideExceptions = exceptions.filter((exception) => exception.slideNumber === slide.number);
    const slideLayoutExceptions = layoutExceptions.filter((exception) => exception.slideNumber === slide.number && exception.severity !== "info");
    if (slideExceptions.length > 0 || slideLayoutExceptions.length > 0 || slide.warnings.length > 0) return {
      slideNumber: slide.number,
      status: "needs-review",
      changeIds,
      reasons: [
        ...slideExceptions.map((exception) => exception.reason),
        ...slideLayoutExceptions.map((exception) => exception.reason),
        ...slide.warnings.map((warning) => `Preview warning: ${warning}`),
        ...(changeIds.length > 0 ? ["Safe deterministic changes are still included in the proposal."] : []),
      ],
    };
    if (changeIds.length > 0) return {
      slideNumber: slide.number,
      status: "change-proposed",
      changeIds,
      reasons: ["A deterministic typography, geometry, or native-table improvement is included in the proposal."],
    };
    return {
      slideNumber: slide.number,
      status: "approved-as-is",
      changeIds: [],
      reasons: ["The deck-wide deterministic pass found no supported change or blocking exception on this slide."],
    };
  });
}

export function createFontCleanupProposal(deck: DeckJob, updatedAt: string): CleanupProposal {
  assertDeckReady(deck);
  if (deck.operationScope !== "cleanup-only") throw new Error("Font cleanup requires cleanup-only operation scope.");
  const changes = fontCleanupChanges(deck);
  if (changes.length === 0) throw new Error("No supported legacy font mappings were found.");
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Normalize ${changes.length} legacy font famil${changes.length === 1 ? "y" : "ies"} without changing text`,
    status: "pending",
    mode: "font-cleanup",
    changes,
    slideDispositions: slideDispositions(deck, changes, [], []),
    tableExceptions: [],
    layoutExceptions: [],
  };
}

export function createDesignerCleanupProposal(deck: DeckJob, updatedAt: string): CleanupProposal {
  assertDeckReady(deck);
  if (!deck.audit) throw new Error("Audit the deck before staging cleanup.");
  const changes = fontCleanupChanges(deck);
  if (deck.audit.alignmentRepairs.length > 0) changes.push({
    id: "alignment-dominant-left-edge",
    kind: "alignment",
    from: "offset text-box alignment",
    to: "dominant peer content edge",
    affectedSlideNumbers: [...new Set(deck.audit.alignmentRepairs.map((repair) => repair.slideNumber))].sort((left, right) => left - right),
    affectedRunCount: deck.audit.alignmentRepairs.length,
    alignmentRepairs: deck.audit.alignmentRepairs,
    rationale: "Align high-confidence cover and peer text-box outliers to a dominant content edge while preserving text, vertical position, size, and reading order.",
    selected: true,
  });
  const tableExceptions = deck.audit.tables.map(tableException).filter((item): item is TableNormalizationException => Boolean(item));
  const compatibleTables = deck.audit.tables.filter((table) => !tableExceptions.some((exception) => exception.tableId === table.id));
  if (compatibleTables.length > 0) changes.push({
    id: `table-${TABLE_PROFILE.id}`,
    kind: "table-style",
    from: "mixed native table formatting",
    to: "ORNL native table profile",
    affectedSlideNumbers: [...new Set(compatibleTables.map((table) => table.slideNumber))].sort((left, right) => left - right),
    affectedRunCount: compatibleTables.reduce((sum, table) => sum + table.rowCount * table.columnCount, 0),
    tableIds: compatibleTables.map((table) => table.id),
    profileId: TABLE_PROFILE.id,
    rationale: "Normalize compatible native tables to consistent Aptos typography, padding, fills, and minimal horizontal rules while preserving exact cell text and merged-cell structure.",
    selected: true,
  });
  const layoutExceptions = deck.audit.layoutReviews ?? [];
  const dispositions = slideDispositions(deck, changes, tableExceptions, layoutExceptions);
  const changedSlideCount = dispositions.filter((item) => item.status === "change-proposed" || item.changeIds.length > 0).length;
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Designer cleanup reviewed all ${deck.audit.slideCount} slides and proposes supported improvements on ${changedSlideCount}`,
    status: "pending",
    mode: "designer-cleanup",
    standardVersion: PRESENTATION_DESIGN_STANDARD.version,
    changes,
    slideDispositions: dispositions,
    tableExceptions,
    layoutExceptions,
  };
}

function objectGeometry(object: SlideEditableObject) {
  const { x, y, width, height } = object.geometry;
  return { x, y, width, height };
}

type GeometryConstraints = GeometryEditCommand["constraints"];

export interface GeometryEditRequest {
  objectId: string;
  target: { x: number; y: number; width: number; height: number };
  rationale: string;
  author: "human" | "ai";
  constraints?: Partial<GeometryConstraints>;
}

export interface VisualDesignRequest {
  slideNumber: number;
  clearPendingLayoutRemap?: boolean;
  textStyles: Array<Omit<TextStyleCommand, "id" | "slideNumber" | "shapeId" | "typeface">>;
  decorations: Array<Omit<DecorativeShapeCommand, "slideNumber">>;
}

const GEOMETRY_DEFAULTS: GeometryConstraints = { allowIntentionalOverlap: false, allowFitRisk: false, allowSafeArea: false, allowAspectRatioChange: false };
const GEOMETRY_SAFE_MARGIN_EMU = Math.round(.25 * 914_400);
const GEOMETRY_MARGIN_TOLERANCE_EMU = Math.round(.02 * 914_400);

function geometryOverlapRatio(left: GeometryEditCommand["source"], right: GeometryEditCommand["source"]): number {
  const width = Math.max(0, Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x));
  const height = Math.max(0, Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y));
  const minimumArea = Math.min(left.width * left.height, right.width * right.height);
  return minimumArea > 0 ? (width * height) / minimumArea : 0;
}

function geometryEdgeMargin(geometry: GeometryEditCommand["source"], slideSize: { width: number; height: number }): number {
  return Math.min(geometry.x, geometry.y, slideSize.width - geometry.x - geometry.width, slideSize.height - geometry.y - geometry.height);
}

function collisionSensitive(object: SlideEditableObject): boolean {
  return ["text", "picture", "table", "chart", "graphic-frame"].includes(object.kind);
}

function validateGeometryCommands(deck: DeckJob, requests: GeometryEditRequest[]): GeometryEditCommand[] {
  if (!deck.audit) throw new Error("Audit the deck before editing slide geometry.");
  if (requests.length === 0 || requests.length > 200) throw new Error("Stage between 1 and 200 object geometry edits in one atomic proposal.");
  if (new Set(requests.map((request) => request.objectId)).size !== requests.length) throw new Error("Each object may appear only once in an atomic geometry proposal.");
  const objects = deck.audit.editableObjects ?? [];
  const slideSize = deck.audit.slideSize ?? { width: 12_192_000, height: 6_858_000 };
  const commands = requests.map((input): GeometryEditCommand => {
    const object = objects.find((item) => item.id === input.objectId);
    if (!object) throw new Error(`Editable object ${input.objectId} is not present in the current deck revision.`);
    if (deck.protectedSlideNumbers.includes(object.slideNumber)) throw new Error(`Slide ${object.slideNumber} is protected from automated or in-app geometry edits.`);
    const source = objectGeometry(object);
    const target = Object.fromEntries(Object.entries(input.target).map(([key, value]) => [key, Math.round(value)])) as typeof input.target;
    if (![target.x, target.y, target.width, target.height].every(Number.isFinite)) throw new Error(`Geometry for ${object.id} must contain finite x, y, width, and height values.`);
    if (target.width < 91_440 || target.height < 91_440) throw new Error(`${object.name} must remain at least 0.1 inches wide and high.`);
    if (target.x < 0 || target.y < 0 || target.x + target.width > slideSize.width || target.y + target.height > slideSize.height) throw new Error(`${object.name} must remain inside the physical slide boundary.`);
    const moved = target.x !== source.x || target.y !== source.y;
    const resized = target.width !== source.width || target.height !== source.height;
    if (!moved && !resized) throw new Error(`The staged geometry for ${object.name} is identical to the current geometry.`);
    if (moved && !object.canMove) throw new Error(`${object.name} cannot be moved safely by the current editor.`);
    if (resized && !object.canResize) throw new Error(`${object.name} cannot be resized safely by the current editor.`);
    const constraints = { ...GEOMETRY_DEFAULTS, ...input.constraints };
    const warnings: string[] = [];
    const sourceMargin = geometryEdgeMargin(source, slideSize);
    const targetMargin = geometryEdgeMargin(target, slideSize);
    const safeAreaStatus = targetMargin < GEOMETRY_SAFE_MARGIN_EMU ? "near-edge" as const : "inside" as const;
    if (safeAreaStatus === "near-edge" && targetMargin < sourceMargin - GEOMETRY_MARGIN_TOLERANCE_EMU) {
      if (!constraints.allowSafeArea) throw new Error(`${object.name} would move farther into the 0.25-inch safe-margin review zone. Set allowSafeArea only when the template or design intent requires it.`);
      warnings.push("The target intentionally moves farther into the 0.25-inch safe-margin review zone.");
    }
    let fitRatio: number | undefined;
    const textBox = deck.audit?.textBoxes.find((item) => item.slideNumber === object.slideNumber && item.shapeId === object.shapeId);
    if (textBox) {
      const widthDemand = Math.max(1, source.width / target.width);
      fitRatio = Number(((textBox.estimatedRequiredHeightEmu * widthDemand) / target.height).toFixed(3));
      if (fitRatio > Math.max(1.18, textBox.fitRatio + .05)) {
        if (!constraints.allowFitRisk) throw new Error(`${object.name} would worsen estimated text fit from ${textBox.fitRatio.toFixed(2)}× to ${fitRatio.toFixed(2)}×. Widen or heighten the box instead of silently risking clipping or shrinkage.`);
        warnings.push(`The target intentionally accepts an estimated ${fitRatio.toFixed(2)}× text-fit risk.`);
      }
    }
    if (object.kind === "picture" && resized) {
      const sourceRatio = source.width / source.height;
      const targetRatio = target.width / target.height;
      const ratioDrift = Math.abs(Math.log(targetRatio / sourceRatio));
      if (ratioDrift > .03) {
        if (!constraints.allowAspectRatioChange) throw new Error(`${object.name} would change picture aspect ratio. Resize proportionally or explicitly authorize a crop/frame-ratio change.`);
        warnings.push("The target intentionally changes the picture frame aspect ratio and requires crop inspection.");
      }
    }
    return {
      id: `geometry-${object.id}`,
      slideNumber: object.slideNumber,
      objectId: object.id,
      shapeId: object.shapeId,
      sourceElement: object.sourceElement,
      objectKind: object.kind,
      operation: moved && resized ? "move-and-resize" : moved ? "move" : "resize",
      source,
      target,
      rationale: input.rationale.trim().slice(0, 700) || `Adjust ${object.name} to improve slide alignment and fit while preserving content.`,
      author: input.author,
      constraints,
      validation: { fitRatio, safeAreaStatus, overlapObjectIds: [], warnings },
    };
  });

  const finalGeometry = new Map(objects.map((object) => [object.id, objectGeometry(object)]));
  for (const command of commands) finalGeometry.set(command.objectId, command.target);
  for (const command of commands) {
    const object = objects.find((item) => item.id === command.objectId);
    if (!object || !collisionSensitive(object)) continue;
    const overlapObjectIds: string[] = [];
    for (const other of objects.filter((item) => item.slideNumber === object.slideNumber && item.id !== object.id && collisionSensitive(item))) {
      const sourceOverlap = geometryOverlapRatio(command.source, objectGeometry(other));
      const targetOverlap = geometryOverlapRatio(command.target, finalGeometry.get(other.id) ?? objectGeometry(other));
      if (targetOverlap > .04 && targetOverlap > sourceOverlap + .02) overlapObjectIds.push(other.id);
    }
    if (overlapObjectIds.length > 0 && !command.constraints.allowIntentionalOverlap) throw new Error(`${object.name} would create or materially increase overlap with ${overlapObjectIds.join(", ")}. Adjust the layout or explicitly authorize an intentional overlay.`);
    if (overlapObjectIds.length > 0) command.validation.warnings.push(`Intentional overlap requires visual review with ${overlapObjectIds.join(", ")}.`);
    command.validation.overlapObjectIds = overlapObjectIds;
  }
  return commands;
}

export function createGeometryBatchProposal(deck: DeckJob, updatedAt: string, inputs: GeometryEditRequest[]): CleanupProposal {
  assertDeckReady(deck);
  if (!deck.audit) throw new Error("Audit the deck before editing slide geometry.");
  const base: CleanupProposal = deck.proposal && ["pending", "applied"].includes(deck.proposal.status) && deck.proposal.baseUpdatedAt === updatedAt ? deck.proposal : {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: "Stage a bounded slide layout transaction",
    status: "pending",
    mode: "slide-geometry",
    standardVersion: PRESENTATION_DESIGN_STANDARD.version,
    changes: [],
    slideDispositions: [],
    tableExceptions: [],
    layoutExceptions: [],
  };
  const existingGeometry = base.changes.flatMap((change) => change.kind === "geometry" ? (change.geometryCommands ?? []).map((command) => ({ command, selected: change.selected })) : []);
  const incomingIds = new Set(inputs.map((input) => input.objectId));
  const combinedInputs: GeometryEditRequest[] = [
    ...existingGeometry.filter(({ command }) => !incomingIds.has(command.objectId)).map(({ command }) => ({ objectId: command.objectId, target: command.target, rationale: command.rationale, author: command.author, constraints: command.constraints ?? GEOMETRY_DEFAULTS })),
    ...inputs,
  ];
  const commands = validateGeometryCommands(deck, combinedInputs);
  const selectedByObject = new Map(existingGeometry.map(({ command, selected }) => [command.objectId, selected]));
  for (const input of inputs) selectedByObject.set(input.objectId, true);
  const targetedShapes = new Set(commands.map((command) => `${command.slideNumber}:${command.shapeId}`));
  const changes = base.changes
    .filter((change) => change.kind !== "geometry")
    .map((change) => change.kind === "alignment" ? { ...change, alignmentRepairs: (change.alignmentRepairs ?? []).filter((repair) => !targetedShapes.has(`${repair.slideNumber}:${repair.shapeId}`)) } : change)
    .filter((change) => change.kind !== "alignment" || (change.alignmentRepairs?.length ?? 0) > 0);
  for (const command of commands) changes.push({
      id: command.id,
      kind: "geometry",
      from: `${(command.source.x / 914_400).toFixed(2)}, ${(command.source.y / 914_400).toFixed(2)} · ${(command.source.width / 914_400).toFixed(2)} × ${(command.source.height / 914_400).toFixed(2)} in`,
      to: `${(command.target.x / 914_400).toFixed(2)}, ${(command.target.y / 914_400).toFixed(2)} · ${(command.target.width / 914_400).toFixed(2)} × ${(command.target.height / 914_400).toFixed(2)} in`,
      affectedSlideNumbers: [command.slideNumber],
      affectedRunCount: 1,
      geometryCommands: [command],
      rationale: command.rationale,
      selected: selectedByObject.get(command.objectId) ?? true,
    });
  const geometryCount = commands.length;
  const revisedSlideNumbers = new Set(commands.filter((command) => incomingIds.has(command.objectId)).map((command) => command.slideNumber));
  return {
    ...base,
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    summary: `Stage ${geometryCount} object-level slide edit${geometryCount === 1 ? "" : "s"} with exact content preserved`,
    status: "pending",
    mode: changes.every((change) => change.kind === "geometry") ? "slide-geometry" : "designer-cleanup",
    changes,
    slideDispositions: slideDispositions(deck, changes, base.tableExceptions, base.layoutExceptions),
    slideReviews: (base.slideReviews ?? []).filter((review) => !revisedSlideNumbers.has(review.slideNumber)),
  };
}

export function createGeometryEditProposal(deck: DeckJob, updatedAt: string, input: GeometryEditRequest): CleanupProposal {
  return createGeometryBatchProposal(deck, updatedAt, [input]);
}

function nativeLayoutChange(command: NativeLayoutRemapCommand): CleanupChange {
  return {
    id: command.id,
    kind: "layout-remap",
    from: "source native master/layout relationship",
    to: `${command.templateLayoutName} · approved native layout remap`,
    affectedSlideNumbers: [command.slideNumber],
    affectedRunCount: 1,
    layoutCommands: [command],
    rationale: command.rationale,
    selected: true,
  };
}

export function createNativeLayoutProposal(deck: DeckJob, updatedAt: string, command: NativeLayoutRemapCommand): CleanupProposal {
  assertDeckReady(deck);
  if (!deck.audit?.slides.some((slide) => slide.number === command.slideNumber)) throw new Error(`Slide ${command.slideNumber} is not present in the audited source deck.`);
  if (deck.protectedSlideNumbers.includes(command.slideNumber)) throw new Error(`Slide ${command.slideNumber} is protected from native layout remapping.`);
  const change = nativeLayoutChange(command);
  const retained = deck.proposal && ["pending", "applied"].includes(deck.proposal.status) && deck.proposal.baseUpdatedAt === updatedAt ? deck.proposal.changes.filter((existing) => existing.kind !== "layout-remap" || !existing.affectedSlideNumbers.includes(command.slideNumber)) : [];
  const changes = [...retained.filter((existing) => existing.kind === "layout-remap"), change, ...retained.filter((existing) => existing.kind !== "layout-remap")];
  return {
    id: crypto.randomUUID(),
    deckId: deck.id,
    baseUpdatedAt: updatedAt,
    createdAt: new Date().toISOString(),
    summary: `Remap slide ${command.slideNumber} to native ${command.templateLayoutName}`,
    status: "pending",
    mode: "slide-reflow",
    standardVersion: PRESENTATION_DESIGN_STANDARD.version,
    changes,
    slideDispositions: slideDispositions(deck, changes, deck.proposal?.tableExceptions ?? [], deck.audit.layoutReviews ?? []),
    slideReviews: (deck.proposal?.slideReviews ?? []).filter((review) => review.slideNumber !== command.slideNumber),
    tableExceptions: deck.proposal?.tableExceptions ?? [],
    layoutExceptions: deck.proposal?.layoutExceptions ?? deck.audit.layoutReviews ?? [],
  };
}

export function createNativeLayoutRecompositionProposal(deck: DeckJob, updatedAt: string, command: NativeLayoutRemapCommand, geometry: GeometryEditRequest[]): CleanupProposal {
  const proposal = createGeometryBatchProposal(deck, updatedAt, geometry);
  const layout = nativeLayoutChange(command);
  const retainedLayouts = proposal.changes.filter((change) => change.kind === "layout-remap" && !change.affectedSlideNumbers.includes(command.slideNumber));
  const changes = [...retainedLayouts, layout, ...proposal.changes.filter((change) => change.kind !== "layout-remap")];
  return {
    ...proposal,
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    summary: `Recompose slide ${command.slideNumber} into native ${command.templateLayoutName}`,
    mode: "slide-reflow",
    changes,
    slideDispositions: slideDispositions(deck, changes, proposal.tableExceptions, proposal.layoutExceptions),
  };
}

function cleanHexColor(value: string | undefined, field: string): string | undefined {
  if (value === undefined) return undefined;
  if (!/^#[0-9a-f]{6}$/i.test(value)) throw new Error(`${field} must be a six-digit hex color such as #00662C.`);
  return value.toUpperCase();
}

export function createVisualDesignProposal(deck: DeckJob, updatedAt: string, input: VisualDesignRequest): CleanupProposal {
  assertDeckReady(deck);
  if (!deck.audit) throw new Error("Audit the deck before staging visual design.");
  if (deck.operationScope !== "reflow") throw new Error("Visual styling and brand geometry require Designer Cleanup reflow scope.");
  if (!Number.isInteger(input.slideNumber) || input.slideNumber < 1 || input.slideNumber > deck.audit.slideCount) throw new Error(`Choose a slide from 1 to ${deck.audit.slideCount}.`);
  if (input.textStyles.length + input.decorations.length === 0) throw new Error("Stage at least one text style or decoration.");
  if (input.textStyles.length > 20 || input.decorations.length > 30) throw new Error("Stage no more than 20 text styles and 30 decorations for one slide.");
  const objects = deck.audit.editableObjects ?? [];
  const slideSize = deck.audit.slideSize ?? { width: 12_192_000, height: 6_858_000 };
  const textStyles: TextStyleCommand[] = input.textStyles.map((style) => {
    const object = objects.find((item) => item.id === style.objectId && item.slideNumber === input.slideNumber);
    if (!object || object.sourceElement !== "p:sp" || !["text", "shape"].includes(object.kind)) throw new Error(`Text object ${style.objectId} is not editable on slide ${input.slideNumber}.`);
    if (style.fontSizePt !== undefined && (style.fontSizePt < 10 || style.fontSizePt > 60)) throw new Error("Text size must be between 10 and 60 pt.");
    const insets = style.insetsInches;
    if (insets && Object.values(insets).some((value) => !Number.isFinite(value) || value < 0 || value > .25)) throw new Error("Text insets must be between 0 and 0.25 inches.");
    return {
      ...style,
      id: `text-style-${object.id}`,
      slideNumber: input.slideNumber,
      shapeId: object.shapeId,
      typeface: "Aptos",
      color: cleanHexColor(style.color, "Text color"),
      rationale: style.rationale.trim().slice(0, 700),
    };
  });
  if (new Set(textStyles.map((style) => style.objectId)).size !== textStyles.length) throw new Error("Each text object may be styled only once per slide transaction.");
  const decorations: DecorativeShapeCommand[] = input.decorations.map((decoration) => {
    const geometry = Object.fromEntries(Object.entries(decoration.geometry).map(([key, value]) => [key, Math.round(value)])) as DecorativeShapeCommand["geometry"];
    if (![geometry.x, geometry.y, geometry.width, geometry.height].every(Number.isFinite) || geometry.width < 9_144 || geometry.height < 9_144) throw new Error(`${decoration.name} must have finite on-slide geometry at least 0.01 inches wide and high.`);
    if (geometry.x < 0 || geometry.y < 0 || geometry.x + geometry.width > slideSize.width || geometry.y + geometry.height > slideSize.height) throw new Error(`${decoration.name} must remain inside the physical slide boundary.`);
    if (decoration.lineWidthPt < 0 || decoration.lineWidthPt > 6) throw new Error("Decoration line width must be between 0 and 6 pt.");
    if (!decoration.fillColor && !decoration.lineColor) throw new Error(`${decoration.name} needs a fill or line color.`);
    return { ...decoration, slideNumber: input.slideNumber, geometry, fillColor: cleanHexColor(decoration.fillColor, "Decoration fill"), lineColor: cleanHexColor(decoration.lineColor, "Decoration line"), rationale: decoration.rationale.trim().slice(0, 700) };
  });
  if (new Set(decorations.map((decoration) => decoration.id)).size !== decorations.length) throw new Error("Each decoration ID must be unique per slide transaction.");

  const base: CleanupProposal = deck.proposal && ["pending", "applied"].includes(deck.proposal.status) && deck.proposal.baseUpdatedAt === updatedAt ? deck.proposal : {
    id: crypto.randomUUID(), deckId: deck.id, baseUpdatedAt: updatedAt, createdAt: new Date().toISOString(), summary: "Stage a bounded visual design transaction", status: "pending", mode: "slide-reflow", standardVersion: PRESENTATION_DESIGN_STANDARD.version, changes: [], slideDispositions: [], tableExceptions: [], layoutExceptions: deck.audit.layoutReviews ?? [],
  };
  const styleObjectIds = new Set(textStyles.map((style) => style.objectId));
  const decorationIds = new Set(decorations.map((decoration) => decoration.id));
  const retained = base.changes.filter((change) => !input.clearPendingLayoutRemap || change.kind !== "layout-remap" || !change.affectedSlideNumbers.includes(input.slideNumber)).map((change) => change.kind === "text-style" ? { ...change, textStyleCommands: (change.textStyleCommands ?? []).filter((style) => !styleObjectIds.has(style.objectId)) } : change.kind === "decoration" ? { ...change, decorationCommands: (change.decorationCommands ?? []).filter((decoration) => !decorationIds.has(decoration.id)) } : change).filter((change) => change.kind !== "text-style" || (change.textStyleCommands?.length ?? 0) > 0).filter((change) => change.kind !== "decoration" || (change.decorationCommands?.length ?? 0) > 0);
  const changes: CleanupChange[] = [...retained];
  if (textStyles.length > 0) changes.push({ id: `text-style-slide-${input.slideNumber}`, kind: "text-style", from: "mixed source text hierarchy", to: "measured Aptos hierarchy", affectedSlideNumbers: [input.slideNumber], affectedRunCount: textStyles.length, textStyleCommands: textStyles, rationale: "Apply deliberate ORNL-aligned hierarchy, readable sizing, optical alignment, and text-frame spacing without changing any wording.", selected: true });
  if (decorations.length > 0) changes.push({ id: `decoration-slide-${input.slideNumber}`, kind: "decoration", from: "unstructured open canvas", to: "restrained ORNL visual grouping", affectedSlideNumbers: [input.slideNumber], affectedRunCount: decorations.length, decorationCommands: decorations, rationale: "Add native editable vector rules, panels, and grouping cues that clarify the existing content without adding claims or decorative clutter.", selected: true });
  return { ...base, id: crypto.randomUUID(), createdAt: new Date().toISOString(), summary: `Polish slide ${input.slideNumber} with native editable hierarchy and brand geometry`, status: "pending", mode: "slide-reflow", changes, slideDispositions: slideDispositions(deck, changes, base.tableExceptions, base.layoutExceptions), slideReviews: (base.slideReviews ?? []).filter((review) => review.slideNumber !== input.slideNumber) };
}

function escapeXmlAttribute(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function replaceTypeface(xml: string, from: string, to: string): { xml: string; replacements: number } {
  let replacements = 0;
  const source = from.toLowerCase();
  const result = xml.replace(/\btypeface=("([^"]*)"|'([^']*)')/gi, (whole, quoted: string, doubleValue: string, singleValue: string) => {
    const current = doubleValue ?? singleValue ?? "";
    if (current.trim().toLowerCase() !== source) return whole;
    replacements += 1;
    const quote = quoted.startsWith("'") ? "'" : '"';
    return `typeface=${quote}${escapeXmlAttribute(to)}${quote}`;
  });
  return { xml: result, replacements };
}

function setAttribute(attributes: string, name: string, value: string): string {
  const expression = new RegExp(`\\s${name}=(?:"[^"]*"|'[^']*')`, "i");
  return expression.test(attributes) ? attributes.replace(expression, ` ${name}="${value}"`) : `${attributes} ${name}="${value}"`;
}

function directSolidFill(color: string) {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

function normalizeRunPropertyTag(tag: string, header: boolean): string {
  const match = tag.match(/^<a:(rPr|defRPr|endParaRPr)\b([^>]*?)(\/?)>([\s\S]*?)(?:<\/a:\1>)?$/);
  if (!match) return tag;
  const name = match[1];
  let attributes = setAttribute(match[2] ?? "", "sz", String(TABLE_FONT_SIZE));
  if (header) attributes = setAttribute(attributes, "b", "1");
  let children = match[4] ?? "";
  children = children.replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>|<a:solidFill\b[^>]*\/>/g, "");
  children = children.replace(/<a:latin\b[^>]*\/>/g, "");
  const color = header ? TABLE_PROFILE.header.textColor.slice(1) : TABLE_PROFILE.body.textColor.slice(1);
  return `<a:${name}${attributes}>${directSolidFill(color)}<a:latin typeface="${escapeXmlAttribute(TABLE_PROFILE.fontFamily)}"/>${children}</a:${name}>`;
}

function normalizeTextProperties(cell: string, header: boolean): string {
  let result = cell.replace(/<a:(rPr|defRPr|endParaRPr)\b[^>]*\/>|<a:(rPr|defRPr|endParaRPr)\b[^>]*>[\s\S]*?<\/a:\2>/g, (tag) => normalizeRunPropertyTag(tag, header));
  result = result.replace(/<a:r>(?!\s*<a:rPr\b)/g, `<a:r><a:rPr sz="${TABLE_FONT_SIZE}"${header ? ' b="1"' : ""}>${directSolidFill(header ? TABLE_PROFILE.header.textColor.slice(1) : TABLE_PROFILE.body.textColor.slice(1))}<a:latin typeface="${escapeXmlAttribute(TABLE_PROFILE.fontFamily)}"/></a:rPr>`);
  return result;
}

function normalizeBodyProperties(cell: string): string {
  return cell.replace(/<a:bodyPr\b([^>]*?)(?:\/?>)/, (_tag, initial: string) => {
    let attributes = initial ?? "";
    attributes = setAttribute(attributes, "lIns", String(TABLE_MARGIN_HORIZONTAL_EMU));
    attributes = setAttribute(attributes, "rIns", String(TABLE_MARGIN_HORIZONTAL_EMU));
    attributes = setAttribute(attributes, "tIns", String(TABLE_MARGIN_VERTICAL_EMU));
    attributes = setAttribute(attributes, "bIns", String(TABLE_MARGIN_VERTICAL_EMU));
    attributes = setAttribute(attributes, "anchor", "ctr");
    return `<a:bodyPr${attributes}/>`;
  });
}

function normalizeCellProperties(cell: string, fill: string, lastRow: boolean): string {
  const rules = `<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB${lastRow ? "" : ` w="${TABLE_RULE_WIDTH_EMU}"`}>${lastRow ? "<a:noFill/>" : directSolidFill(TABLE_PROFILE.strokes.horizontal.color.slice(1))}</a:lnB>`;
  const replacement = (attributes: string, children: string) => {
    let nextAttributes = attributes;
    nextAttributes = setAttribute(nextAttributes, "marL", String(TABLE_MARGIN_HORIZONTAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marR", String(TABLE_MARGIN_HORIZONTAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marT", String(TABLE_MARGIN_VERTICAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "marB", String(TABLE_MARGIN_VERTICAL_EMU));
    nextAttributes = setAttribute(nextAttributes, "anchor", "ctr");
    const cleaned = children
      .replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>|<a:solidFill\b[^>]*\/>/g, "")
      .replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>/g, "");
    return `<a:tcPr${nextAttributes}>${directSolidFill(fill)}${rules}${cleaned}</a:tcPr>`;
  };
  if (/<a:tcPr\b[^>]*\/>/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)\/>/, (_tag, attributes) => replacement(attributes, ""));
  if (/<a:tcPr\b/.test(cell)) return cell.replace(/<a:tcPr\b([^>]*)>([\s\S]*?)<\/a:tcPr>/, (_tag, attributes, children) => replacement(attributes, children));
  return cell.replace(/<\/a:tc>/, `${replacement("", "")}</a:tc>`);
}

function normalizeTableBlock(table: string): string {
  let result = table.replace(/<a:tblPr\b([^>]*)>/, (_tag, initial: string) => {
    let attributes = setAttribute(initial ?? "", "firstRow", "1");
    attributes = setAttribute(attributes, "bandRow", "0");
    attributes = setAttribute(attributes, "bandCol", "0");
    return `<a:tblPr${attributes}>`;
  });
  const rows = [...result.matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)];
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index][0];
    const header = index === 0;
    const lastRow = index === rows.length - 1;
    const fill = header ? TABLE_PROFILE.header.fill.slice(1) : (index % 2 === 0 ? TABLE_PROFILE.body.alternateFill : TABLE_PROFILE.body.fill).slice(1);
    const normalized = row.replace(/<a:tc\b[\s\S]*?<\/a:tc>/g, (cell) => normalizeCellProperties(normalizeBodyProperties(normalizeTextProperties(cell, header)), fill, lastRow));
    const start = rows[index].index ?? 0;
    result = result.slice(0, start) + normalized + result.slice(start + row.length);
  }
  return result;
}

function normalizeSelectedTables(xml: string, slideNumber: number, selectedIds: Set<string>): { xml: string; count: number } {
  let ordinal = 0;
  let count = 0;
  return {
    xml: xml.replace(/<a:tbl\b[\s\S]*?<\/a:tbl>/g, (table) => {
      ordinal += 1;
      if (!selectedIds.has(`slide-${slideNumber}-table-${ordinal}`)) return table;
      count += 1;
      return normalizeTableBlock(table);
    }),
    get count() { return count; },
  };
}

function applyAlignmentRepairs(xml: string, slideNumber: number, repairs: CleanupChange["alignmentRepairs"]): { xml: string; count: number } {
  const selected = (repairs ?? []).filter((repair) => repair.slideNumber === slideNumber);
  if (selected.length === 0) return { xml, count: 0 };
  let count = 0;
  const next = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    const shapeId = shape.match(/<p:cNvPr\b[^>]*\bid=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean);
    const repair = selected.find((item) => item.shapeId === shapeId);
    if (!repair) return shape;
    return shape.replace(/<a:off\b([^>]*)\/>/, (tag, initial: string) => {
      const currentX = Number(initial.match(/\bx=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean));
      const currentY = Number(initial.match(/\by=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean));
      if (currentX !== repair.source.x || currentY !== repair.source.y) return tag;
      let attributes = setAttribute(initial, "x", String(repair.target.x));
      attributes = setAttribute(attributes, "y", String(repair.target.y));
      count += 1;
      return `<a:off${attributes}/>`;
    });
  });
  return { xml: next, count };
}

function applyGeometryCommands(xml: string, slideNumber: number, commands: GeometryEditCommand[]): { xml: string; count: number } {
  const selected = commands.filter((command) => command.slideNumber === slideNumber);
  if (selected.length === 0) return { xml, count: 0 };
  let result = xml;
  let count = 0;
  for (const command of selected) {
    const escapedElement = command.sourceElement.replace(":", "\\:");
    const objectExpression = new RegExp(`<${escapedElement}\\b[\\s\\S]*?<\\/${escapedElement}>`, "g");
    let matched = false;
    result = result.replace(objectExpression, (objectBlock) => {
      const shapeId = objectBlock.match(/<p:cNvPr\b[^>]*\bid=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean);
      if (shapeId !== command.shapeId) return objectBlock;
      const transformElement = command.sourceElement === "p:graphicFrame" ? "p:xfrm" : "a:xfrm";
      const transformExpression = new RegExp(`<${transformElement}\\b([^>]*)>([\\s\\S]*?)<\\/${transformElement}>`);
      return objectBlock.replace(transformExpression, (transformBlock, transformAttributes: string, transformChildren: string) => {
        const offset = transformChildren.match(/<a:off\b([^>]*)\/>/);
        const extent = transformChildren.match(/<a:ext\b([^>]*)\/>/);
        if (!offset || !extent) return transformBlock;
        const currentX = Number(attributeFromXml(offset[1] ?? "", "x"));
        const currentY = Number(attributeFromXml(offset[1] ?? "", "y"));
        const currentWidth = Number(attributeFromXml(extent[1] ?? "", "cx"));
        const currentHeight = Number(attributeFromXml(extent[1] ?? "", "cy"));
        if (currentX !== command.source.x || currentY !== command.source.y || currentWidth !== command.source.width || currentHeight !== command.source.height) return transformBlock;
        let offsetAttributes = setAttribute(offset[1] ?? "", "x", String(command.target.x));
        offsetAttributes = setAttribute(offsetAttributes, "y", String(command.target.y));
        let extentAttributes = setAttribute(extent[1] ?? "", "cx", String(command.target.width));
        extentAttributes = setAttribute(extentAttributes, "cy", String(command.target.height));
        const children = transformChildren
          .replace(offset[0], `<a:off${offsetAttributes}/>`)
          .replace(extent[0], `<a:ext${extentAttributes}/>`);
        matched = true;
        return `<${transformElement}${transformAttributes}>${children}</${transformElement}>`;
      });
    });
    if (matched) count += 1;
  }
  return { xml: result, count };
}

function textAlignmentValue(value: TextStyleCommand["alignment"]): string | undefined {
  return value === "left" ? "l" : value === "center" ? "ctr" : value === "right" ? "r" : undefined;
}

function verticalAlignmentValue(value: TextStyleCommand["verticalAlignment"]): string | undefined {
  return value === "top" ? "t" : value === "middle" ? "ctr" : value === "bottom" ? "b" : undefined;
}

function styleRunProperties(tag: string, command: TextStyleCommand): string {
  const match = tag.match(/^<a:(rPr|defRPr|endParaRPr)\b([^>]*)>/);
  if (!match) return tag;
  const name = match[1];
  let attributes = (match[2] ?? "").replace(/\/\s*$/, "");
  if (command.fontSizePt !== undefined) attributes = setAttribute(attributes, "sz", String(Math.round(command.fontSizePt * 100)));
  if (command.bold !== undefined) attributes = setAttribute(attributes, "b", command.bold ? "1" : "0");
  if (command.italic !== undefined) attributes = setAttribute(attributes, "i", command.italic ? "1" : "0");
  const closingTag = `</a:${name}>`;
  let children = tag.endsWith("/>") ? "" : tag.slice(tag.indexOf(">") + 1, tag.lastIndexOf(closingTag));
  if (command.color) {
    const color = directSolidFill(command.color.slice(1));
    if (/<a:solidFill\b/.test(children)) children = children.replace(/<a:solidFill\b[\s\S]*?<\/a:solidFill>|<a:solidFill\b[^>]*\/>/, color);
    else children = children.replace(/(?=<a:(?:highlight|uLn|uFill|latin|ea|cs|sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b)|$/, color);
  }
  const latin = `<a:latin typeface="${escapeXmlAttribute(command.typeface)}"/>`;
  if (/<a:latin\b/.test(children)) children = children.replace(/<a:latin\b[^>]*\/>/, latin);
  else children = children.replace(/(?=<a:(?:ea|cs|sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b)|$/, latin);
  return `<a:${name}${attributes}>${children}</a:${name}>`;
}

function defaultRunProperties(command: TextStyleCommand): string {
  let attributes = "";
  if (command.fontSizePt !== undefined) attributes = setAttribute(attributes, "sz", String(Math.round(command.fontSizePt * 100)));
  if (command.bold !== undefined) attributes = setAttribute(attributes, "b", command.bold ? "1" : "0");
  if (command.italic !== undefined) attributes = setAttribute(attributes, "i", command.italic ? "1" : "0");
  return `<a:rPr${attributes}>${command.color ? directSolidFill(command.color.slice(1)) : ""}<a:latin typeface="${escapeXmlAttribute(command.typeface)}"/></a:rPr>`;
}

function styleParagraphs(shape: string, command: TextStyleCommand): string {
  const alignment = textAlignmentValue(command.alignment);
  if (!alignment) return shape;
  return shape.replace(/<a:p\b([^>]*)>([\s\S]*?)<\/a:p>/g, (_paragraph, paragraphAttributes: string, children: string) => {
    let nextChildren = children;
    if (/<a:pPr\b/.test(nextChildren)) nextChildren = nextChildren.replace(/<a:pPr\b([^>]*)>/, (tag, initial: string) => {
      const selfClosing = tag.endsWith("/>");
      const attributes = setAttribute(initial.replace(/\/\s*$/, ""), "algn", alignment);
      return `<a:pPr${attributes}${selfClosing ? "/" : ""}>`;
    });
    else nextChildren = `<a:pPr algn="${alignment}"/>${nextChildren}`;
    return `<a:p${paragraphAttributes}>${nextChildren}</a:p>`;
  });
}

function applyTextStyleCommands(xml: string, slideNumber: number, commands: TextStyleCommand[]): { xml: string; count: number; matchedIds: string[] } {
  const selected = commands.filter((command) => command.slideNumber === slideNumber);
  if (selected.length === 0) return { xml, count: 0, matchedIds: [] };
  const matchedIds = new Set<string>();
  const next = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    const shapeId = shape.match(/<p:cNvPr\b[^>]*\bid=(?:"([^"]+)"|'([^']+)')/)?.slice(1).find(Boolean);
    const command = selected.find((item) => item.shapeId === shapeId);
    if (!command) return shape;
    let styled = shape.replace(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\/>|<a:(rPr|defRPr|endParaRPr)\b[^>]*>[\s\S]*?<\/a:\1>/g, (tag) => styleRunProperties(tag, command));
    styled = styled.replace(/<a:r\b([^>]*)>(?!\s*<a:rPr\b)/g, (_tag, attributes: string) => `<a:r${attributes}>${defaultRunProperties(command)}`);
    styled = styleParagraphs(styled, command);
    const anchor = verticalAlignmentValue(command.verticalAlignment);
    if (anchor || command.insetsInches) {
      styled = styled.replace(/<a:bodyPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:bodyPr>)/, (_tag, initial: string, children: string) => {
        let attributes = initial ?? "";
        if (anchor) attributes = setAttribute(attributes, "anchor", anchor);
        if (command.insetsInches) {
          attributes = setAttribute(attributes, "tIns", String(Math.round(command.insetsInches.top * 914_400)));
          attributes = setAttribute(attributes, "rIns", String(Math.round(command.insetsInches.right * 914_400)));
          attributes = setAttribute(attributes, "bIns", String(Math.round(command.insetsInches.bottom * 914_400)));
          attributes = setAttribute(attributes, "lIns", String(Math.round(command.insetsInches.left * 914_400)));
        }
        return children === undefined ? `<a:bodyPr${attributes}/>` : `<a:bodyPr${attributes}>${children}</a:bodyPr>`;
      });
    }
    matchedIds.add(command.id);
    return styled;
  });
  return { xml: next, count: matchedIds.size, matchedIds: [...matchedIds] };
}

function decorationXml(command: DecorativeShapeCommand, shapeId: number): string {
  const fill = command.fillColor ? directSolidFill(command.fillColor.slice(1)) : "<a:noFill/>";
  const line = command.lineColor && command.lineWidthPt > 0 ? `<a:ln w="${Math.round(command.lineWidthPt * 12_700)}">${directSolidFill(command.lineColor.slice(1))}<a:prstDash val="solid"/></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="${escapeXmlAttribute(command.name)}" descr="Presentation Studio native editable decoration"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${command.geometry.x}" y="${command.geometry.y}"/><a:ext cx="${command.geometry.width}" cy="${command.geometry.height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}${line}</p:spPr></p:sp>`;
}

function applyDecorationCommands(xml: string, slideNumber: number, commands: DecorativeShapeCommand[]): { xml: string; count: number } {
  const selected = commands.filter((command) => command.slideNumber === slideNumber);
  if (selected.length === 0) return { xml, count: 0 };
  let nextId = Math.max(1, ...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid=(?:"(\d+)"|'(\d+)')/g)].map((match) => Number(match[1] ?? match[2] ?? 0))) + 1;
  const behind = selected.filter((command) => command.behindContent).map((command) => decorationXml(command, nextId++)).join("");
  const front = selected.filter((command) => !command.behindContent).map((command) => decorationXml(command, nextId++)).join("");
  let result = xml;
  if (behind) result = result.replace(/(<p:grpSpPr\b[\s\S]*?<\/p:grpSpPr>)/, `$1${behind}`);
  if (front) result = result.replace(/<\/p:spTree>/, `${front}</p:spTree>`);
  return { xml: result, count: selected.length };
}

function attributeFromXml(attributes: string, name: string): string | undefined {
  return attributes.match(new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "i"))?.slice(1).find(Boolean);
}

async function materializeCleanup(sourceBytes: Uint8Array, proposal: CleanupProposal, requireAccepted: boolean, options?: { templateBytes?: Uint8Array }): Promise<{ bytes: Uint8Array; replacementCount: number; tableCount: number; alignmentCount: number; geometryCount: number; layoutCount: number; textStyleCount: number; decorationCount: number; layoutReceipts: NativeLayoutCloneReceipt[]; normalizedTableIds: string[] }> {
  if (requireAccepted && proposal.status !== "applied") throw new Error("Accept the cleanup plan before exporting a review copy.");
  const selected = proposal.changes.filter((change) => change.selected);
  if (selected.length === 0) throw new Error("Select at least one cleanup change.");
  const sourceAudit = await auditPptx(sourceBytes);
  const zip = await JSZip.loadAsync(sourceBytes, { checkCRC32: false });
  let replacementCount = 0;
  let tableCount = 0;
  let alignmentCount = 0;
  let geometryCount = 0;
  let textStyleCount = 0;
  let decorationCount = 0;
  const matchedTextStyleIds = new Set<string>();
  const normalizedTableIds = selected.flatMap((change) => change.kind === "table-style" ? change.tableIds ?? [] : []);
  const selectedTableIds = new Set(normalizedTableIds);
  const selectedAlignmentRepairs = selected.flatMap((change) => change.kind === "alignment" ? change.alignmentRepairs ?? [] : []);
  const selectedGeometryCommands = selected.flatMap((change) => change.kind === "geometry" ? change.geometryCommands ?? [] : []);
  const selectedLayoutCommands = selected.flatMap((change) => change.kind === "layout-remap" ? change.layoutCommands ?? [] : []);
  const selectedTextStyleCommands = selected.flatMap((change) => change.kind === "text-style" ? change.textStyleCommands ?? [] : []);
  const selectedDecorationCommands = selected.flatMap((change) => change.kind === "decoration" ? change.decorationCommands ?? [] : []);

  const slidePaths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path));
  for (const path of slidePaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    let xml = await entry.async("text");
    const slideNumber = Number(path.match(/slide(\d+)\.xml$/i)?.[1] ?? 0);
    for (const change of selected.filter((item) => item.kind === "font-family")) {
      const next = replaceTypeface(xml, change.from, change.to);
      xml = next.xml;
      replacementCount += next.replacements;
    }
    const tables = normalizeSelectedTables(xml, slideNumber, selectedTableIds);
    xml = tables.xml;
    tableCount += tables.count;
    const geometry = applyGeometryCommands(xml, slideNumber, selectedGeometryCommands);
    xml = geometry.xml;
    geometryCount += geometry.count;
    const alignments = applyAlignmentRepairs(xml, slideNumber, selectedAlignmentRepairs);
    xml = alignments.xml;
    alignmentCount += alignments.count;
    const textStyles = applyTextStyleCommands(xml, slideNumber, selectedTextStyleCommands);
    xml = textStyles.xml;
    textStyleCount += textStyles.count;
    for (const id of textStyles.matchedIds) matchedTextStyleIds.add(id);
    const decorations = applyDecorationCommands(xml, slideNumber, selectedDecorationCommands);
    xml = decorations.xml;
    decorationCount += decorations.count;
    zip.file(path, xml);
  }
  if (alignmentCount !== selectedAlignmentRepairs.length) throw new Error("Alignment validation failed because a staged text box no longer matched its source geometry.");
  if (geometryCount !== selectedGeometryCommands.length) throw new Error("Geometry validation failed because a staged object no longer matched its source geometry.");
  if (textStyleCount !== selectedTextStyleCommands.length) {
    const missing = selectedTextStyleCommands.filter((command) => !matchedTextStyleIds.has(command.id)).map((command) => `${command.objectId} (slide ${command.slideNumber}, shape ${command.shapeId})`).slice(0, 12);
    throw new Error(`Text-style validation failed because ${selectedTextStyleCommands.length - textStyleCount} staged text object${selectedTextStyleCommands.length - textStyleCount === 1 ? "" : "s"} no longer matched the source revision: ${missing.join(", ")}.`);
  }
  if (decorationCount !== selectedDecorationCommands.length) throw new Error("Decoration validation failed because a staged slide no longer matched the source revision.");
  if (replacementCount === 0 && tableCount === 0 && alignmentCount === 0 && geometryCount === 0 && textStyleCount === 0 && decorationCount === 0 && selectedLayoutCommands.length === 0) throw new Error("The selected cleanup changes did not match any editable slide markup.");

  let output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const layoutReceipts: NativeLayoutCloneReceipt[] = [];
  for (const command of selectedLayoutCommands) {
    if (!options?.templateBytes) throw new Error("The exact active Template Pack bytes are required to materialize this native layout proposal.");
    const remapped = await cloneTemplateLayoutForSlide({ sourceBytes: output, templateBytes: options.templateBytes, command });
    output = remapped.bytes;
    layoutReceipts.push(remapped.receipt);
  }
  const outputAudit = await auditPptx(output);
  if (outputAudit.slideCount !== sourceAudit.slideCount) throw new Error("Cleanup validation failed because the slide count changed.");
  for (let index = 0; index < sourceAudit.slides.length; index += 1) {
    if (sourceAudit.slides[index].textHash !== outputAudit.slides[index]?.textHash) {
      throw new Error(`Cleanup validation failed because visible text changed on slide ${sourceAudit.slides[index].number}.`);
    }
  }
  for (const sourceTable of sourceAudit.tables) {
    const outputTable = outputAudit.tables.find((table) => table.id === sourceTable.id);
    if (!outputTable || sourceTable.contentHash !== outputTable.contentHash) throw new Error(`Cleanup validation failed because table content changed in ${sourceTable.id}.`);
    if (sourceTable.structureHash !== outputTable.structureHash) throw new Error(`Cleanup validation failed because merged-cell structure changed in ${sourceTable.id}.`);
  }
  for (const command of selectedGeometryCommands) {
    const outputObject = (outputAudit.editableObjects ?? []).find((object) => object.id === command.objectId);
    if (!outputObject || outputObject.geometry.x !== command.target.x || outputObject.geometry.y !== command.target.y || outputObject.geometry.width !== command.target.width || outputObject.geometry.height !== command.target.height) throw new Error(`Geometry validation failed for ${command.objectId}.`);
  }
  return { bytes: output, replacementCount, tableCount, alignmentCount, geometryCount, layoutCount: selectedLayoutCommands.length, textStyleCount, decorationCount, layoutReceipts, normalizedTableIds };
}

export async function buildCleanupProposalPptx(sourceBytes: Uint8Array, proposal: CleanupProposal, options?: { templateBytes?: Uint8Array }) {
  return materializeCleanup(sourceBytes, proposal, false, options);
}

export async function applyCleanupToPptx(sourceBytes: Uint8Array, proposal: CleanupProposal, options?: { templateBytes?: Uint8Array }) {
  return materializeCleanup(sourceBytes, proposal, true, options);
}
