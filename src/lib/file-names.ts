const PRESENTATION_EXTENSION = /\.(?:pptx|pptm|potx|ppsx)$/i;

export function cleanFileStem(name: string, fallback = "Presentation"): string {
  const withoutExtension = name.replace(PRESENTATION_EXTENSION, "");
  const cleaned = withoutExtension
    .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "")
    .trim();
  return cleaned || fallback;
}

export function projectSaveDefaultName(input: {
  projectName: string;
  deckNames: string[];
  encrypted: boolean;
}): string {
  const extension = input.encrypted ? ".pstudio-secure" : ".pstudio";
  if (input.deckNames.length === 0) return `${cleanFileStem(input.projectName, "Presentation Studio project")}${extension}`;

  const sourceStem = cleanFileStem(input.deckNames[0], "Presentation");
  const batchSuffix = input.deckNames.length > 1 ? `_and-${input.deckNames.length - 1}-more` : "";
  return `${sourceStem}${batchSuffix}_Presentation-Studio${extension}`;
}
