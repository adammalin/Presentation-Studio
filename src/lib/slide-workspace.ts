export interface SlideWorkspaceRequest {
  id: string;
  deckId: string;
  slideNumber: number;
  mode: "review" | "edit" | "comment";
  representation: "current" | "proposal";
}

export function isProposalSlideWorkspaceRequest(request: SlideWorkspaceRequest | undefined, deckId: string | undefined): boolean {
  return Boolean(request && deckId && request.deckId === deckId && request.representation === "proposal");
}
