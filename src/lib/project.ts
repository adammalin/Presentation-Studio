import { z } from "zod";
import { PROJECT_SCHEMA, PROJECT_SCHEMA_VERSION, type PresentationStudioProject } from "../types";

const isoTimestamp = z.string().datetime({ offset: true });
const fontSchema = z.object({
  family: z.string(),
  normalizedFamily: z.string(),
  count: z.number().int().nonnegative(),
  directSlideCount: z.number().int().nonnegative(),
  slideNumbers: z.array(z.number().int().positive()),
  partKinds: z.array(z.string()),
  isThemeFont: z.boolean(),
  isLikelySymbolFont: z.boolean(),
});
const slideSchema = z.object({
  id: z.string(),
  number: z.number().int().positive(),
  title: z.string(),
  text: z.string(),
  textHash: z.string().regex(/^[0-9a-f]{64}$/),
  textRunCount: z.number().int().nonnegative(),
  tableCount: z.number().int().nonnegative(),
  pictureCount: z.number().int().nonnegative(),
  chartCount: z.number().int().nonnegative(),
  connectorCount: z.number().int().nonnegative(),
  commentCount: z.number().int().nonnegative(),
  fonts: z.array(z.string()),
  fontSizes: z.array(z.number().nonnegative()),
  warnings: z.array(z.string()),
});
const tableInventorySchema = z.object({
  id: z.string(),
  slideNumber: z.number().int().positive(),
  ordinal: z.number().int().positive(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  mergedCellCount: z.number().int().nonnegative(),
  styleId: z.string().optional(),
  styleFlags: z.array(z.string()),
  cellFonts: z.array(z.string()),
  colorTokens: z.array(z.string()),
  marginSignatures: z.array(z.string()),
  styleFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
});
const pictureInventorySchema = z.object({
  id: z.string(),
  slideNumber: z.number().int().positive(),
  ordinal: z.number().int().positive(),
  name: z.string(),
  description: z.string().optional(),
  relationshipId: z.string().optional(),
  widthEmu: z.number().int().nonnegative().optional(),
  heightEmu: z.number().int().nonnegative().optional(),
  cropped: z.boolean(),
  hasOutline: z.boolean(),
  hasEffect: z.boolean(),
});
const findingSchema = z.object({
  id: z.string(),
  ruleId: z.string(),
  category: z.enum(["template", "font", "table", "figure", "layout", "production", "technical-review"]),
  severity: z.enum(["info", "warning", "error"]),
  confidence: z.enum(["high", "medium", "low"]),
  slideNumber: z.number().int().positive().optional(),
  message: z.string(),
  evidence: z.string(),
  autoFixable: z.boolean(),
});
const auditSchema = z.object({
  scannedAt: isoTimestamp,
  supportLevel: z.enum(["native-ooxml", "partial", "blocked"]),
  slideCount: z.number().int().nonnegative(),
  masterCount: z.number().int().nonnegative(),
  layoutCount: z.number().int().nonnegative(),
  themeCount: z.number().int().nonnegative(),
  notesCount: z.number().int().nonnegative(),
  legacyCommentCount: z.number().int().nonnegative(),
  modernCommentCount: z.number().int().nonnegative(),
  mediaCount: z.number().int().nonnegative(),
  tableCount: z.number().int().nonnegative(),
  chartCount: z.number().int().nonnegative(),
  pictureCount: z.number().int().nonnegative(),
  containsMacros: z.boolean(),
  containsOleObjects: z.boolean(),
  containsExternalRelationships: z.boolean(),
  packageFileCount: z.number().int().positive(),
  expandedByteLength: z.number().int().nonnegative(),
  classification: z.enum(["current-ornl", "older-or-modified-ornl", "sponsor", "custom", "mixed", "unknown"]),
  classificationEvidence: z.array(z.string()),
  fonts: z.array(fontSchema),
  slides: z.array(slideSchema),
  tables: z.array(tableInventorySchema),
  pictures: z.array(pictureInventorySchema),
  findings: z.array(findingSchema),
  warnings: z.array(z.string()),
});
const changeSchema = z.object({
  id: z.string(),
  kind: z.literal("font-family"),
  from: z.string(),
  to: z.string(),
  affectedSlideNumbers: z.array(z.number().int().positive()),
  affectedRunCount: z.number().int().nonnegative(),
  rationale: z.string(),
  selected: z.boolean(),
});
const proposalSchema = z.object({
  id: z.string(),
  deckId: z.string(),
  baseUpdatedAt: isoTimestamp,
  createdAt: isoTimestamp,
  summary: z.string(),
  status: z.enum(["pending", "applied", "rejected"]),
  changes: z.array(changeSchema),
});
const deckSchema = z.object({
  id: z.string(),
  name: z.string(),
  sourceResourceId: z.string(),
  sourceSha256: z.string().regex(/^[0-9a-f]{64}$/),
  operationScope: z.enum(["audit-only", "cleanup-only", "reflow", "hybrid", "compose"]),
  templateClassification: z.enum(["current-ornl", "older-or-modified-ornl", "sponsor", "custom", "mixed", "unknown"]),
  targetTemplateId: z.string().optional(),
  targetTemplateConfirmedAt: isoTimestamp.optional(),
  status: z.enum(["not-scanned", "audited", "needs-template-decision", "ready-for-cleanup", "proposal-ready", "needs-manual-review", "approved", "exported", "failed"]),
  audit: auditSchema.optional(),
  proposal: proposalSchema.optional(),
  protectedSlideNumbers: z.array(z.number().int().positive()),
  failureMessage: z.string().max(700).optional(),
  exportedAt: isoTimestamp.optional(),
});

export const projectSchema = z.object({
  schema: z.literal(PROJECT_SCHEMA),
  schemaVersion: z.literal(PROJECT_SCHEMA_VERSION),
  project: z.object({
    id: z.string().min(1),
    name: z.string().min(1).max(240),
    type: z.enum(["review-batch", "single-deck", "new-presentation"]),
    createdAt: isoTimestamp,
    updatedAt: isoTimestamp,
  }),
  settings: z.object({
    contentPolicy: z.enum(["preserve-exact", "source-grounded-generative"]),
    defaultOperationScope: z.enum(["audit-only", "cleanup-only", "reflow", "hybrid", "compose"]),
    autosave: z.boolean(),
  }),
  resources: z.array(z.object({
    id: z.string().min(1),
    name: z.string().min(1),
    mediaType: z.string().min(1),
    byteLength: z.number().int().nonnegative(),
    sha256: z.string().regex(/^[0-9a-f]{64}$/),
    roles: z.array(z.enum(["import-origin", "prior-approved-revision", "style-exemplar", "grounding-source", "slide-media", "chart-data", "template-source", "reference-only"])).min(1),
    kind: z.enum(["presentation", "document", "data", "image", "audio", "video", "other"]).optional(),
    support: z.array(z.enum(["source-readable", "previewable", "placeable", "pptx-preserved", "unsupported"])).optional(),
    processing: z.object({
      status: z.enum(["indexed", "stored-only", "needs-review"]),
      summary: z.string().max(1_000),
      processedAt: isoTimestamp,
      warnings: z.array(z.string().max(1_000)).max(50),
    }).optional(),
    derivatives: z.array(z.object({
      id: z.string().min(1),
      kind: z.literal("extracted-text"),
      mediaType: z.literal("text/plain"),
      byteLength: z.number().int().nonnegative(),
      sha256: z.string().regex(/^[0-9a-f]{64}$/),
      createdAt: isoTimestamp,
      processor: z.string().min(1).max(240),
      truncated: z.boolean(),
      bytes: z.instanceof(Uint8Array).optional(),
    })).max(20).optional(),
    createdAt: isoTimestamp,
    sourcePath: z.string().optional(),
    embedded: z.literal(true),
    bytes: z.instanceof(Uint8Array).optional(),
    mcpAccess: z.enum(["none", "metadata", "text", "preview"]),
  })),
  styleExemplars: z.array(z.object({
    id: z.string(),
    name: z.string(),
    kind: z.enum(["table", "figure"]),
    resourceId: z.string(),
    deckId: z.string(),
    slideNumber: z.number().int().positive(),
    objectOrdinal: z.number().int().positive(),
    scope: z.enum(["deck", "batch"]),
    createdAt: isoTimestamp,
  })),
  decks: z.array(deckSchema),
  activity: z.array(z.object({ id: z.string(), at: isoTimestamp, action: z.string(), detail: z.string() })),
});

export function createProject(name = "Untitled review batch"): PresentationStudioProject {
  const now = new Date().toISOString();
  return {
    schema: PROJECT_SCHEMA,
    schemaVersion: PROJECT_SCHEMA_VERSION,
    project: {
      id: crypto.randomUUID(),
      name,
      type: "review-batch",
      createdAt: now,
      updatedAt: now,
    },
    settings: {
      contentPolicy: "preserve-exact",
      defaultOperationScope: "cleanup-only",
      autosave: true,
    },
    resources: [],
    styleExemplars: [],
    decks: [],
    activity: [{ id: crypto.randomUUID(), at: now, action: "project-created", detail: "Created a conservative cleanup review batch." }],
  };
}

export function touchProject(project: PresentationStudioProject, action: string, detail: string): PresentationStudioProject {
  const now = new Date().toISOString();
  return {
    ...project,
    project: { ...project.project, updatedAt: now },
    activity: [...project.activity, { id: crypto.randomUUID(), at: now, action, detail }],
  };
}

export function projectForJson(project: PresentationStudioProject) {
  return {
    ...project,
    resources: project.resources.map(({ bytes: _bytes, sourcePath: _sourcePath, derivatives, ...resource }) => ({
      ...resource,
      mcpAccess: "none" as const,
      derivatives: derivatives?.map(({ bytes: _derivativeBytes, ...derivative }) => derivative),
    })),
  };
}
