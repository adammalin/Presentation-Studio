import assert from "node:assert/strict";
import test from "node:test";
import { buildDesignRepairLedger } from "../src/lib/design-repair-loop";
import type { DeckJob } from "../src/types";

const technicalDeck = {
  id: "technical-deck",
  audit: {
    slideCount: 1,
    slides: [{ number: 1, text: "Fault ON Resistance 10 [ohm]", textHash: "a".repeat(64), pictureCount: 3, connectorCount: 2 }],
    findings: [],
  },
  scene: {
    objects: [
      { id: "screenshot-a", slideNumber: 1, kind: "picture" },
      { id: "relationship-arrow", slideNumber: 1, kind: "connector" },
    ],
  },
} as unknown as DeckJob;

test("inspection reports found issues and carries an exact original-intent reference", () => {
  const ledger = buildDesignRepairLedger({ deck: technicalDeck, slideNumber: 1, representation: "current" });
  assert.equal(ledger.phase, "found-issues");
  assert.equal(ledger.issues.some((issue) => issue.category === "intent"), true);
  assert.deepEqual(ledger.originalIntentReference.sourceVisualObjectIds, ["screenshot-a", "relationship-arrow"]);
  assert.deepEqual(ledger.originalIntentReference.relationshipObjectIds, ["relationship-arrow"]);
  assert.match(ledger.instruction, /fix the bounded items/i);
});

test("proposal inspection switches to original-message rechecking", () => {
  const ledger = buildDesignRepairLedger({ deck: technicalDeck, slideNumber: 1, representation: "proposal" });
  assert.equal(ledger.phase, "rechecking-original-intent");
  assert.match(ledger.instruction, /prettier slide is not better/i);
});
