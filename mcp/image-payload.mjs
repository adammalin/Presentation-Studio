export function stripImagePayloads(images) {
  return images.map(({ data: _data, bytes: _bytes, ...metadata }) => metadata);
}
