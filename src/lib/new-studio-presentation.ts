import type {
  PptxAudit,
  ProjectResource,
  StudioLayoutRecipe,
  StudioResourceBinding,
  StudioTableCellDesign,
  StudioWebFrame,
  StudioWebNode,
  StudioWebScene,
  StudioWebSlide,
} from "../types";
import { STUDIO_WEB_SCENE_SCHEMA, STUDIO_WEB_SCENE_VERSION } from "../types";
import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";
import { sha256Text } from "./hash";
import { defaultStudioDeckRhythm, recomposeStudioWebSlide } from "./studio-web-scene";
import type { TemplateCatalog, TemplateLayoutPreview } from "./template-catalog";

const EMU_PER_INCH = 914_400;

export interface NewStudioSourceReference {
  resourceId: string;
  exactExcerpt?: string;
}

export interface NewStudioTableInput {
  headers: string[];
  rows: string[][];
}

export interface NewStudioSlideInput {
  title: string;
  subtitle?: string;
  body: string[];
  recipe: Exclude<StudioLayoutRecipe, "source">;
  layoutId?: string;
  imageResourceIds?: string[];
  table?: NewStudioTableInput;
  sourceReferences: NewStudioSourceReference[];
  rationale: string;
}

export interface NewStudioPresentationInput {
  deckId: string;
  name: string;
  communicationJob: string;
  expression: "restrained" | "balanced" | "expressive";
  slides: NewStudioSlideInput[];
}

function frame(x: number, y: number, width: number, height: number): StudioWebFrame {
  return { x: Math.round(x * EMU_PER_INCH), y: Math.round(y * EMU_PER_INCH), width: Math.round(width * EMU_PER_INCH), height: Math.round(height * EMU_PER_INCH), rotation: 0 };
}

function normalizedText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function uniqueBindings(bindings: StudioResourceBinding[]): StudioResourceBinding[] {
  const seen = new Set<string>();
  return bindings.filter((binding) => {
    const key = `${binding.resourceId}:${binding.kind}:${binding.exactExcerptHash ?? "media"}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function resourceBinding(resource: ProjectResource, kind: StudioResourceBinding["kind"], exactExcerpt?: string): Promise<StudioResourceBinding> {
  const derivative = kind === "text" || kind === "table" ? resource.derivatives?.find((item) => item.kind === "extracted-text") : undefined;
  const excerpt = exactExcerpt?.trim();
  return {
    resourceId: resource.id,
    resourceSha256: resource.sha256,
    derivativeSha256: derivative?.sha256,
    kind,
    relationship: kind === "image" ? "supplies-media" : kind === "table" ? "supplies-data" : "grounds",
    exactExcerpt: excerpt || undefined,
    exactExcerptHash: excerpt ? await sha256Text(normalizedText(excerpt)) : undefined,
  };
}

function baseStyle(role: StudioWebNode["role"]): StudioWebNode["style"] {
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  const title = role === "title";
  const caption = role === "caption" || role === "label";
  return {
    fontFamily: "Aptos",
    fontSizePt: title ? 32 : caption ? 14 : 18,
    fontWeight: title ? 700 : 400,
    lineHeight: title ? 1.02 : 1.1,
    color: palette.darkMatter,
    borderWidthPt: 0,
    textAlign: "left",
    verticalAlign: "top",
    paddingPt: { top: 2, right: 4, bottom: 2, left: 4 },
  };
}

async function textNode(input: {
  id: string;
  slideNumber: number;
  name: string;
  role: StudioWebNode["role"];
  text: string;
  sourceTextOrder: number;
  zIndex: number;
  frame: StudioWebFrame;
  resourceBindings: StudioResourceBinding[];
  subtitle?: boolean;
}): Promise<StudioWebNode> {
  const text = input.text.trim();
  const textHash = await sha256Text(normalizedText(text));
  const style = baseStyle(input.role);
  if (input.subtitle) {
    style.fontSizePt = 22;
    style.fontWeight = 400;
    style.color = PRESENTATION_DESIGN_STANDARD.defaults.palette.haleNavy;
  }
  return {
    id: input.id,
    sourceObjectId: `resource-authored:${input.id}`,
    sourceShapeId: `new-${input.slideNumber}-${input.id}`,
    sourceBinding: "catalog-derived",
    name: input.name,
    kind: "text",
    role: input.role,
    sourceFrame: { ...input.frame },
    frame: { ...input.frame },
    zIndex: input.zIndex,
    sourceTextOrder: input.sourceTextOrder,
    visible: true,
    locked: false,
    exactContent: true,
    text,
    textHash,
    sourceParagraphs: [{ index: 1, text, textHash, characterCount: text.length, bullet: false, bulletConfidence: "direct", level: 0, fontFamilies: ["Aptos"], fontSizes: [style.fontSizePt] }],
    resourceBindings: input.resourceBindings,
    style,
  };
}

function tableCellStyles(headers: string[], rows: string[][]): StudioTableCellDesign[] {
  const palette = PRESENTATION_DESIGN_STANDARD.defaults.palette;
  const styles: StudioTableCellDesign[] = [];
  const rowCount = rows.length + 1;
  const columnCount = headers.length;
  for (let row = 1; row <= rowCount; row += 1) {
    for (let column = 1; column <= columnCount; column += 1) {
      styles.push({
        cellId: `cell-r${row}-c${column}`,
        fill: row === 1 ? palette.haleNavy : row % 2 === 0 ? "#FFFFFF" : "#F0F2F1",
        color: row === 1 ? "#FFFFFF" : palette.darkMatter,
        fontSizePt: row === 1 ? 16 : 15,
        fontWeight: row === 1 ? 700 : 400,
        textAlign: column === 1 ? "left" : "center",
        verticalAlign: "middle",
        paddingPt: { top: 5, right: 7, bottom: 5, left: 7 },
        borders: { bottom: { type: "solid", color: row === 1 ? palette.ornlGreen : "#DBDCDB", widthPt: row === 1 ? 1.25 : .5 } },
      });
    }
  }
  return styles;
}

async function tableNode(slideNumber: number, input: NewStudioTableInput, bindings: StudioResourceBinding[], zIndex: number): Promise<StudioWebNode> {
  const rows = [input.headers, ...input.rows];
  const cells = rows.flatMap((row, rowIndex) => input.headers.map((_, columnIndex) => ({
    id: `cell-r${rowIndex + 1}-c${columnIndex + 1}`,
    row: rowIndex + 1,
    column: columnIndex + 1,
    rowSpan: 1,
    columnSpan: 1,
    text: row[columnIndex] ?? "",
  })));
  const text = cells.map((cell) => cell.text).join(" ");
  const bounds = frame(.55, 1.55, 12.23, 4.95);
  return {
    id: `new-slide-${slideNumber}-table`,
    sourceObjectId: `resource-authored:slide-${slideNumber}-table`,
    sourceShapeId: `new-${slideNumber}-table`,
    sourceBinding: "catalog-derived",
    name: `Editable source-grounded table ${slideNumber}`,
    kind: "table",
    role: "table",
    sourceFrame: bounds,
    frame: bounds,
    zIndex,
    sourceTextOrder: 10_000,
    visible: true,
    locked: false,
    exactContent: true,
    textHash: await sha256Text(normalizedText(text)),
    tableId: `new-slide-${slideNumber}-table`,
    table: {
      rows: rows.length,
      columns: input.headers.length,
      cells,
      design: {
        headerRows: 1,
        columnWidths: Array.from({ length: input.headers.length }, () => 1 / input.headers.length),
        rowHeights: Array.from({ length: rows.length }, () => 1 / rows.length),
        borderMode: "none",
        borderColor: "#DBDCDB",
        borderWidthPt: .5,
        defaultPaddingPt: { top: 5, right: 7, bottom: 5, left: 7 },
        cellStyles: tableCellStyles(input.headers, input.rows),
      },
    },
    resourceBindings: bindings.map((binding) => ({ ...binding, kind: "table", relationship: "supplies-data" })),
    style: { ...baseStyle("table"), fontSizePt: 15, verticalAlign: "middle" },
  };
}

async function imageNode(slideNumber: number, resource: ProjectResource, ordinal: number, zIndex: number): Promise<StudioWebNode> {
  const bounds = frame(7.15, 1.55 + ordinal * .18, 5.58, 4.82 - ordinal * .18);
  return {
    id: `new-slide-${slideNumber}-image-${ordinal + 1}`,
    sourceObjectId: `resource:${resource.id}`,
    sourceShapeId: `resource-${resource.id}`,
    sourceBinding: "catalog-derived",
    name: resource.name,
    kind: "image",
    role: "image",
    sourceFrame: bounds,
    frame: bounds,
    zIndex,
    sourceTextOrder: 20_000 + ordinal,
    visible: true,
    locked: false,
    exactContent: true,
    mediaPart: `resource:${resource.id}`,
    resourceBindings: [await resourceBinding(resource, "image")],
    style: { ...baseStyle("image"), objectFit: "contain", paddingPt: { top: 0, right: 0, bottom: 0, left: 0 } },
  };
}

function titleLayout(catalog: TemplateCatalog, requestedId?: string): TemplateLayoutPreview {
  const requested = requestedId ? catalog.layouts.find((layout) => layout.id === requestedId) : undefined;
  if (requested && requested.category !== "title") throw new Error(`The first slide must use an approved title layout; ${requested.name} is categorized as ${requested.category}.`);
  const selected = requested ?? catalog.layouts.find((layout) => layout.category === "title" && Boolean(layout.semantic)) ?? catalog.layouts.find((layout) => layout.category === "title");
  if (!selected) throw new Error("The installed Template Pack does not contain an approved title layout.");
  return selected;
}

function layoutForSlide(catalog: TemplateCatalog, slide: NewStudioSlideInput, slideNumber: number): TemplateLayoutPreview | undefined {
  if (slide.recipe !== "template-layout") return undefined;
  if (slideNumber === 1) return titleLayout(catalog, slide.layoutId);
  const layout = slide.layoutId ? catalog.layouts.find((item) => item.id === slide.layoutId) : undefined;
  if (!layout) throw new Error(`Slide ${slideNumber} must name a valid installed layoutId when recipe is template-layout.`);
  return layout;
}

export async function createNewStudioPresentationScene(input: NewStudioPresentationInput, resources: ProjectResource[], templateCatalog: TemplateCatalog): Promise<StudioWebScene> {
  if (!input.slides.length) throw new Error("Create at least one slide.");
  const byId = new Map(resources.map((resource) => [resource.id, resource]));
  const now = new Date().toISOString();
  const slides: StudioWebSlide[] = [];

  for (const [index, planned] of input.slides.entries()) {
    const slideNumber = index + 1;
    if (slideNumber === 1 && planned.recipe !== "template-layout") throw new Error("The first slide of a new ORNL presentation must use an approved Template Pack title layout.");
    const sourceBindings: StudioResourceBinding[] = [];
    for (const reference of planned.sourceReferences) {
      const resource = byId.get(reference.resourceId);
      if (!resource) throw new Error(`Slide ${slideNumber} references a Resource that is not in the project.`);
      sourceBindings.push(await resourceBinding(resource, resource.kind === "data" ? "table" : "text", reference.exactExcerpt));
    }
    const nodes: StudioWebNode[] = [];
    let order = 0;
    nodes.push(await textNode({ id: `new-slide-${slideNumber}-title`, slideNumber, name: `Slide ${slideNumber} title`, role: "title", text: planned.title, sourceTextOrder: order, zIndex: 10, frame: frame(.55, .34, 12.2, .9), resourceBindings: sourceBindings }));
    order += planned.title.length + 1;
    if (planned.subtitle?.trim()) {
      nodes.push(await textNode({ id: `new-slide-${slideNumber}-subtitle`, slideNumber, name: `Slide ${slideNumber} subtitle`, role: "body", text: planned.subtitle, sourceTextOrder: order, zIndex: 11, frame: frame(.58, 1.45, 6.2, .75), resourceBindings: sourceBindings, subtitle: true }));
      order += planned.subtitle.length + 1;
    }
    for (const [bodyIndex, body] of planned.body.entries()) {
      nodes.push(await textNode({ id: `new-slide-${slideNumber}-body-${bodyIndex + 1}`, slideNumber, name: `Slide ${slideNumber} evidence ${bodyIndex + 1}`, role: "body", text: body, sourceTextOrder: order, zIndex: 20 + bodyIndex, frame: frame(.58, 1.55 + bodyIndex * .78, 6.18, .62), resourceBindings: sourceBindings }));
      order += body.length + 1;
    }
    if (planned.table) nodes.push(await tableNode(slideNumber, planned.table, sourceBindings, 60));
    for (const [imageIndex, resourceId] of (planned.imageResourceIds ?? []).entries()) {
      const resource = byId.get(resourceId);
      if (!resource || resource.kind !== "image") throw new Error(`Slide ${slideNumber} references an image Resource that is unavailable.`);
      nodes.push(await imageNode(slideNumber, resource, imageIndex, 40 + imageIndex));
    }
    const allBindings = uniqueBindings([...sourceBindings, ...nodes.flatMap((node) => node.resourceBindings ?? [])]);
    if (!allBindings.length) throw new Error(`Slide ${slideNumber} must retain at least one Resource binding.`);
    const mappedText = nodes.flatMap((node) => node.kind === "table" ? node.table?.cells.map((cell) => cell.text) ?? [] : node.text ? [node.text] : []).join(" ");
    const sourceTextHash = await sha256Text(normalizedText(mappedText));
    slides.push({
      id: `new-studio-slide-${slideNumber}`,
      slideNumber,
      sourceSlideId: `resource-authored-slide-${slideNumber}`,
      sourceTextHash,
      contentCoverage: { exactTextMapped: true, sourceCharacterCount: normalizedText(mappedText).length, mappedCharacterCount: normalizedText(mappedText).length, sourceTextBoxCount: nodes.filter((node) => node.kind === "text").length, mappedTextNodeCount: nodes.filter((node) => node.kind === "text" || node.kind === "table").length, groupedOrUnsupportedTextPresent: false },
      sourceRevision: `resources:${allBindings.map((binding) => binding.resourceSha256).sort().join(":")}`,
      recipe: planned.recipe,
      background: "#FFFFFF",
      status: "designed",
      designRationale: `${planned.rationale} Communication job: ${input.communicationJob}`.slice(0, 1_000),
      resourceBindings: allBindings,
      figureTreatments: [],
      nodes,
      updatedAt: now,
    });
  }

  const resourceDigest = await sha256Text(JSON.stringify({ name: input.name, communicationJob: input.communicationJob, expression: input.expression, slides: input.slides, resources: resources.map((resource) => [resource.id, resource.sha256]) }));
  let scene: StudioWebScene = {
    schema: STUDIO_WEB_SCENE_SCHEMA,
    version: STUDIO_WEB_SCENE_VERSION,
    revision: `${resourceDigest}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    deckId: input.deckId,
    sourceSha256: resourceDigest,
    slideSize: { width: Math.round(PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches * EMU_PER_INCH), height: Math.round(PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches * EMU_PER_INCH) },
    sourceSlideSize: { width: Math.round(PRESENTATION_DESIGN_STANDARD.defaults.slide.widthInches * EMU_PER_INCH), height: Math.round(PRESENTATION_DESIGN_STANDARD.defaults.slide.heightInches * EMU_PER_INCH) },
    rhythm: defaultStudioDeckRhythm(),
    designMemory: [],
    componentLibrary: [],
    tableLibrary: [],
    tableContinuationPlans: [],
    designSystem: { id: "ornl-presentation-web-v1", standardVersion: PRESENTATION_DESIGN_STANDARD.version, unit: "emu", renderer: "html-css", exportTarget: "editable-powerpoint", compilerModes: ["source-bound-overlay", "fresh-composition"] },
    slides,
  };
  for (const [index, planned] of input.slides.entries()) {
    const slideNumber = index + 1;
    scene = recomposeStudioWebSlide(scene, slideNumber, planned.recipe, layoutForSlide(templateCatalog, planned, slideNumber), planned.rationale);
  }
  return { ...scene, slides: scene.slides.map((slide) => ({ ...slide, nodes: slide.nodes.map((node) => ({ ...node, sourceFrame: { ...node.frame } })) })) };
}

export function bindNewStudioSceneToGeneratedPowerPoint(scene: StudioWebScene, sourceSha256: string, audit: PptxAudit): StudioWebScene {
  const now = new Date().toISOString();
  return {
    ...scene,
    sourceSha256,
    sourceSlideSize: { ...audit.slideSize },
    revision: `${sourceSha256}:web-v${STUDIO_WEB_SCENE_VERSION}:${now}`,
    slides: scene.slides.map((slide) => {
      const audited = audit.slides.find((candidate) => candidate.number === slide.slideNumber);
      if (!audited) throw new Error(`The generated PowerPoint is missing Studio slide ${slide.slideNumber}.`);
      return {
        ...slide,
        sourceSlideId: audited.id,
        sourceTextHash: audited.textHash,
        sourceRevision: sourceSha256,
        recipe: slide.slideNumber === 1 ? "source" : slide.recipe,
        status: "designed",
        nodes: slide.nodes.map((node) => ({ ...node, sourceFrame: { ...node.frame } })),
        updatedAt: now,
      };
    }),
  };
}
