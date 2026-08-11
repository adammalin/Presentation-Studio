import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { ArrowLeft, ArrowRight, Check, X } from "@phosphor-icons/react";
import {
  clampTourStep,
  ONBOARDING_TOUR_STEPS,
  tourCardPosition,
  type TourRect,
} from "./lib/onboarding";

interface TourLayout {
  highlight: TourRect | null;
  card: { left: number; top: number; placement: string };
}

function visibleTarget(targetName: string): HTMLElement | null {
  const element = document.querySelector(`[data-tour="${targetName}"]`);
  if (!(element instanceof HTMLElement) || element.getClientRects().length === 0) return null;
  const style = getComputedStyle(element);
  return style.display === "none" || style.visibility === "hidden" ? null : element;
}

function paddedRect(rect: DOMRect, padding = 8, margin = 6): TourRect {
  const left = Math.max(margin, rect.left - padding);
  const top = Math.max(margin, rect.top - padding);
  const right = Math.min(window.innerWidth - margin, rect.right + padding);
  const bottom = Math.min(window.innerHeight - margin, rect.bottom + padding);
  return { left, top, right, bottom, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
}

export default function OnboardingTour({ open, stepIndex, onStepChange, onClose }: {
  open: boolean;
  stepIndex: number;
  onStepChange: (index: number) => void;
  onClose: (remember: boolean) => void;
}) {
  const cardRef = useRef<HTMLElement>(null);
  const nextButtonRef = useRef<HTMLButtonElement>(null);
  const returnFocusRef = useRef<HTMLElement | null>(null);
  const [layout, setLayout] = useState<TourLayout>({ highlight: null, card: { left: 0, top: 0, placement: "center" } });
  const currentIndex = clampTourStep(stepIndex);
  const step = ONBOARDING_TOUR_STEPS[currentIndex];

  const measure = useCallback(() => {
    if (!open) return;
    const target = visibleTarget(step.target);
    target?.scrollIntoView({ block: "center", inline: "nearest" });
    const targetRect = target?.getBoundingClientRect() ?? null;
    const cardRect = cardRef.current?.getBoundingClientRect();
    const card = tourCardPosition(
      targetRect ? { left: targetRect.left, right: targetRect.right, top: targetRect.top, bottom: targetRect.bottom, width: targetRect.width, height: targetRect.height } : null,
      { width: cardRect?.width ?? 380, height: cardRect?.height ?? 280 },
      { width: window.innerWidth, height: window.innerHeight },
    );
    setLayout({ highlight: targetRect ? paddedRect(targetRect) : null, card });
  }, [open, step.target]);

  useLayoutEffect(() => {
    if (!open) return;
    let firstFrame = 0;
    let secondFrame = 0;
    const settleTimer = window.setTimeout(measure, 320);
    measure();
    firstFrame = window.requestAnimationFrame(() => {
      secondFrame = window.requestAnimationFrame(measure);
    });
    const target = visibleTarget(step.target);
    const observer = typeof ResizeObserver === "undefined" ? undefined : new ResizeObserver(measure);
    if (target) observer?.observe(target);
    if (cardRef.current) observer?.observe(cardRef.current);
    window.addEventListener("resize", measure);
    window.addEventListener("scroll", measure, true);
    return () => {
      window.clearTimeout(settleTimer);
      window.cancelAnimationFrame(firstFrame);
      window.cancelAnimationFrame(secondFrame);
      observer?.disconnect();
      window.removeEventListener("resize", measure);
      window.removeEventListener("scroll", measure, true);
    };
  }, [measure, open, step.target]);

  useEffect(() => {
    if (!open) return;
    if (!returnFocusRef.current) returnFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const focusTimer = window.setTimeout(() => nextButtonRef.current?.focus({ preventScroll: true }), 0);
    return () => window.clearTimeout(focusTimer);
  }, [currentIndex, open]);

  useEffect(() => {
    if (open) return;
    const returnTarget = returnFocusRef.current;
    returnFocusRef.current = null;
    if (returnTarget?.isConnected) window.requestAnimationFrame(() => returnTarget.focus({ preventScroll: true }));
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") { event.preventDefault(); onClose(true); return; }
      if (event.key === "ArrowRight") { event.preventDefault(); currentIndex === ONBOARDING_TOUR_STEPS.length - 1 ? onClose(true) : onStepChange(currentIndex + 1); return; }
      if (event.key === "ArrowLeft" && currentIndex > 0) { event.preventDefault(); onStepChange(currentIndex - 1); return; }
      if (event.key !== "Tab") return;
      const controls = [...(cardRef.current?.querySelectorAll<HTMLButtonElement>("button:not(:disabled)") ?? [])];
      if (controls.length === 0) return;
      const activeIndex = controls.indexOf(document.activeElement as HTMLButtonElement);
      const nextIndex = event.shiftKey ? (activeIndex <= 0 ? controls.length - 1 : activeIndex - 1) : (activeIndex < 0 || activeIndex === controls.length - 1 ? 0 : activeIndex + 1);
      event.preventDefault();
      controls[nextIndex].focus();
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, [currentIndex, onClose, onStepChange, open]);

  if (!open) return null;
  const finalStep = currentIndex === ONBOARDING_TOUR_STEPS.length - 1;
  return (
    <div className={`onboarding-tour ${layout.highlight ? "has-target" : "is-centered"}`} data-step-index={currentIndex} data-step-id={step.id} data-step-count={ONBOARDING_TOUR_STEPS.length} data-target={step.target}>
      <div className="tour-interaction-shield" />
      {layout.highlight && <div className="tour-spotlight" data-tour-spotlight style={{ left: layout.highlight.left, top: layout.highlight.top, width: layout.highlight.width, height: layout.highlight.height }} />}
      <section ref={cardRef} className="tour-card" role="dialog" aria-modal="true" aria-labelledby="tour-title" aria-describedby="tour-body" tabIndex={-1} data-placement={layout.card.placement} style={{ left: layout.card.left, top: layout.card.top }}>
        <div className="tour-card-heading">
          <div><span className="tour-eyebrow">{step.eyebrow}</span><span className="tour-progress" aria-live="polite">Step {currentIndex + 1} of {ONBOARDING_TOUR_STEPS.length}</span></div>
          <button className="tour-close" onClick={() => onClose(true)} aria-label="Skip and close tutorial"><X size={17} /></button>
        </div>
        <h2 id="tour-title">{step.title}</h2>
        <p id="tour-body">{step.body}</p>
        <div className="tour-dots" aria-hidden="true">{ONBOARDING_TOUR_STEPS.map((candidate, index) => <i key={candidate.id} className={index === currentIndex ? "current" : index < currentIndex ? "complete" : ""} />)}</div>
        <div className="tour-footer">
          <button className="tour-skip" onClick={() => onClose(true)}>Skip tour</button>
          <div className="tour-actions">
            <button className="button ghost small" disabled={currentIndex === 0} onClick={() => onStepChange(currentIndex - 1)}><ArrowLeft size={15} />Back</button>
            <button ref={nextButtonRef} className="button primary small" onClick={() => finalStep ? onClose(true) : onStepChange(currentIndex + 1)}>{finalStep ? <><Check size={15} />Done</> : <>Next<ArrowRight size={15} /></>}</button>
          </div>
        </div>
        <small className="tour-keyboard-hint">Use ← and → to move · Esc to close</small>
      </section>
    </div>
  );
}
