import { PRESENTATION_DESIGN_STANDARD } from "./design-standard";

const roles = PRESENTATION_DESIGN_STANDARD.semanticVisualPolicy.tableColorRoles;

export function normalizeCellFillToken(kind: string, value: string): string {
  const normalizedKind = kind.toLowerCase();
  const normalizedValue = value.trim().toLowerCase();
  return normalizedKind === "schemeclr" ? normalizedValue : `${normalizedKind.replace(/clr$/, "")}:${normalizedValue}`;
}

export function semanticColorRoleForToken(token?: string): string | undefined {
  if (!token) return undefined;
  const normalized = token.toLowerCase();
  if (roles[normalized]) return normalized;
  if (!normalized.startsWith("srgb:")) return undefined;
  const hex = `#${normalized.slice(5)}`.toUpperCase();
  return Object.entries(roles).find(([, role]) => role.base.toUpperCase() === hex || role.tint100.toUpperCase() === hex)?.[0];
}

export function semanticTableTintForToken(token?: string): string | undefined {
  const role = semanticColorRoleForToken(token);
  return role ? roles[role]?.tint100 : undefined;
}

export function semanticTableRoleDefinition(role: string) {
  return roles[role];
}
