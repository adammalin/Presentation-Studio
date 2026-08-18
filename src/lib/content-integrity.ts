/**
 * Exact visible-copy signature that tolerates layout-driven reordering,
 * paragraph atomization, and PowerPoint run-boundary whitespace. Character
 * identity and multiplicity remain strict, including case, punctuation,
 * numbers, and units.
 */
export function contentCharacterSignature(values: string[]): string {
  const counts = new Map<number, number>();
  for (const value of values) {
    for (const character of value.normalize("NFKC")) {
      if (/\s/u.test(character)) continue;
      const codePoint = character.codePointAt(0);
      if (codePoint === undefined) continue;
      counts.set(codePoint, (counts.get(codePoint) ?? 0) + 1);
    }
  }
  return [...counts.entries()]
    .sort(([left], [right]) => left - right)
    .map(([codePoint, count]) => `${codePoint.toString(16)}:${count}`)
    .join("|");
}
