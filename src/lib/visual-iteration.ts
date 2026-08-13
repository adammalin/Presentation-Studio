import type { CleanupProposal } from "../types";

type VisualIterationEntry = NonNullable<CleanupProposal["visualIteration"]>["history"][number];

export function decideVisualIteration(input: {
  priorHistory: VisualIterationEntry[];
  requestedVerdict: "better" | "revise" | "reject";
  rationale: string;
  slideNumber: number;
  inspectionRevision: string;
  currentRasterSha256: string;
  proposalRasterSha256: string;
  changedPixelRatio: number;
  improvements: string[];
  regressions: string[];
  reviewedAt?: string;
}) {
  const attempt = input.priorHistory.length + 1;
  if (attempt > 3) throw new Error("The bounded three-attempt AI visual loop is exhausted.");
  const objectiveReasons = [...input.regressions];
  if (input.changedPixelRatio <= 0) objectiveReasons.push("proposal pixels are identical to current");
  let verdict = input.requestedVerdict;
  if (verdict === "better" && objectiveReasons.length) verdict = "revise";
  const exhausted = verdict === "revise" && attempt === 3;
  if (exhausted) verdict = "reject";
  const rationale = `${input.rationale.trim()}${input.requestedVerdict === "better" && objectiveReasons.length ? ` Presentation Studio withheld the requested better verdict because ${objectiveReasons.join("; ")}.` : ""}${exhausted ? " The bounded three-attempt loop is exhausted." : ""}`.trim().slice(0, 1_000);
  const entry: VisualIterationEntry = {
    attempt,
    slideNumber: input.slideNumber,
    inspectionRevision: input.inspectionRevision,
    verdict,
    rationale,
    reviewedAt: input.reviewedAt ?? new Date().toISOString(),
    currentRasterSha256: input.currentRasterSha256,
    proposalRasterSha256: input.proposalRasterSha256,
    metrics: { improvements: [...input.improvements], regressions: objectiveReasons },
  };
  return { entry, verdict, rejected: verdict === "reject", exhausted, objectiveReasons };
}
