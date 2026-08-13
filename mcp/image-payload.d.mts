export function stripImagePayloads<T extends Record<string, unknown>>(
  images: T[],
): Array<Omit<T, "data" | "bytes">>;
