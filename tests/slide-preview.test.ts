import assert from "node:assert/strict";
import test from "node:test";
import { slidePreviewSvg } from "../src/lib/slide-preview";
import type { SlideRenderCatalog } from "../src/lib/template-catalog";

test("slide previews keep presentation-scale type and never fall back to a serif default", () => {
  const catalog: SlideRenderCatalog = {
    id: "catalog",
    name: "Fixture",
    sha256: "0".repeat(64),
    slideWidth: 12_192_000,
    slideHeight: 6_858_000,
    media: {},
    generatedAt: "2026-08-12T00:00:00.000Z",
    renderer: "local-ooxml-preview",
    slides: [{
      id: "slide-1",
      number: 1,
      name: "Slide 1",
      title: "Training Objectives",
      category: "content",
      background: "#FFFFFF",
      hidden: false,
      placeholderTypes: ["title"],
      sourcePart: "ppt/slides/slide1.xml",
      renderWarnings: [],
      elements: [{
        id: "title",
        kind: "text",
        name: "Title",
        x: 429_767,
        y: 274_320,
        width: 11_430_000,
        height: 424_732,
        rotation: 0,
        geometry: "rect",
        text: "Training Objectives",
        fontFamily: "Aptos",
        fontSize: 18,
        fontWeight: 400,
      }],
    }],
  };

  const svg = slidePreviewSvg(catalog, catalog.slides[0], 1200, '@font-face{font-family:"Aptos";src:url("data:font/ttf;base64,AA==")}');
  assert.match(svg, /font-family="&quot;Aptos&quot;, &quot;Arial&quot;, sans-serif"/);
  assert.match(svg, /font-size="22\.5"/);
  assert.doesNotMatch(svg, /font-size="228600"/);
  assert.match(svg, /@font-face\{font-family:"Aptos"/);
});
