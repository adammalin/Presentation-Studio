import assert from "node:assert/strict";
import test from "node:test";
import { singleFlight } from "../src/lib/single-flight";

test("single flight shares one in-progress build and releases the revision afterward", async () => {
  const flights = new Map<string, Promise<number>>();
  let calls = 0;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const operation = async () => { calls += 1; await gate; return calls; };
  const first = singleFlight(flights, "deck:revision", operation);
  const second = singleFlight(flights, "deck:revision", operation);
  assert.equal(first, second);
  assert.equal(calls, 1);
  release();
  assert.equal(await first, 1);
  assert.equal(flights.size, 0);
  assert.equal(await singleFlight(flights, "deck:revision", async () => { calls += 1; return calls; }), 2);
});
