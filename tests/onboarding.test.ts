import assert from "node:assert/strict";
import test from "node:test";
import {
  clampTourStep,
  ONBOARDING_TOUR_STEPS,
  ONBOARDING_TOUR_VERSION,
  shouldShowOnboardingTour,
  tourCardPosition,
} from "../src/lib/onboarding";

test("first-run tour is versioned and every step has a stable unique target", () => {
  assert.equal(shouldShowOnboardingTour(null), true);
  assert.equal(shouldShowOnboardingTour("older"), true);
  assert.equal(shouldShowOnboardingTour(ONBOARDING_TOUR_VERSION), false);
  assert.equal(new Set(ONBOARDING_TOUR_STEPS.map((step) => step.id)).size, ONBOARDING_TOUR_STEPS.length);
  assert.equal(new Set(ONBOARDING_TOUR_STEPS.map((step) => step.target)).size, ONBOARDING_TOUR_STEPS.length);
});
test("tour step navigation stays inside the available range", () => {
  assert.equal(clampTourStep(-8), 0);
  assert.equal(clampTourStep(2.9), 2);
  assert.equal(clampTourStep(10_000), ONBOARDING_TOUR_STEPS.length - 1);
});

test("tour card placement remains inside the viewport", () => {
  const viewport = { width: 1200, height: 800 };
  const card = { width: 390, height: 270 };
  const targets = [
    { left: 20, right: 170, top: 20, bottom: 70, width: 150, height: 50 },
    { left: 20, right: 170, top: 680, bottom: 740, width: 150, height: 60 },
    { left: 1000, right: 1180, top: 20, bottom: 70, width: 180, height: 50 },
    { left: 500, right: 700, top: 350, bottom: 450, width: 200, height: 100 },
  ];
  for (const target of targets) {
    const result = tourCardPosition(target, card, viewport);
    assert.ok(result.left >= 16);
    assert.ok(result.top >= 16);
    assert.ok(result.left + card.width <= viewport.width - 16);
    assert.ok(result.top + card.height <= viewport.height - 16);
  }
});
