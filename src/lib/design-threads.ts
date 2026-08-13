import type { CleanupProposal, DesignThread } from "../types";

export function removeAddressedDesignThreads(threads: DesignThread[], deckId: string, proposal: CleanupProposal, addressedThreadIds: string[] = []): DesignThread[] {
  const affectedSlides = new Set(proposal.changes.filter((change) => change.selected).flatMap((change) => change.affectedSlideNumbers));
  return removeAddressedDesignThreadsForSlides(threads, deckId, affectedSlides, addressedThreadIds);
}

export function removeAddressedDesignThreadsForSlides(threads: DesignThread[], deckId: string, affectedSlides: Iterable<number>, addressedThreadIds: string[] = []): DesignThread[] {
  if (addressedThreadIds.length === 0) return threads;
  const addressed = new Set(addressedThreadIds);
  const slides = new Set(affectedSlides);
  return threads.filter((thread) => !(addressed.has(thread.id) && thread.deckId === deckId && ["submitted", "proposal-ready"].includes(thread.status) && slides.has(thread.slideNumber)));
}

export function removeCompletedDesignThreads(threads: DesignThread[]): DesignThread[] {
  return threads.filter((thread) => !["proposal-ready", "resolved"].includes(thread.status));
}

export function removeDesignThread(threads: DesignThread[], threadId: string): DesignThread[] {
  return threads.filter((thread) => thread.id !== threadId);
}
