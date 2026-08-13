import fs from "node:fs";

const standardUrl = new URL("../shared/presentation-design-standard.json", import.meta.url);
const parsed = JSON.parse(fs.readFileSync(standardUrl, "utf8"));

if (parsed?.schema !== "presentation-studio/design-standard" || typeof parsed.version !== "string" || parsed.defaults?.slide?.aspectRatio !== "16:9" || parsed.defaults?.typography?.family !== "Aptos") {
  throw new Error("The packaged Presentation Studio design standard is invalid.");
}

export const DESIGN_CONTRACT_VERSION = parsed.version;
export const DESIGN_CONTRACT = Object.freeze(parsed);

export function designContractMessage() {
  return `Presentation Studio design standard ${DESIGN_CONTRACT_VERSION}: 16:9 and Aptos by default. Inspect native PowerPoint, compose the whole slide in the shared semantic HTML/CSS Studio Web Scene, compile supported nodes to editable PowerPoint, then judge the native Current/Proposal renders. Make a substantive layout decision instead of merely shrinking text; preserve approved content exactly and verify the independently rerendered and remeasured export before completion.`;
}
