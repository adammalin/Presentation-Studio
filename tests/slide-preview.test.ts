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

test("slide previews preserve imported PowerPoint source crops", () => {
  const catalog: SlideRenderCatalog = {
    id: "crop-catalog",
    name: "Crop fixture",
    sha256: "1".repeat(64),
    slideWidth: 1_000,
    slideHeight: 500,
    media: { "logo.svg": "data:image/svg+xml;base64,PHN2Zy8+" },
    generatedAt: "2026-08-20T00:00:00.000Z",
    renderer: "local-ooxml-preview",
    slides: [{
      id: "slide-1",
      number: 1,
      name: "Slide 1",
      title: "",
      category: "content",
      background: "#FFFFFF",
      hidden: false,
      placeholderTypes: [],
      sourcePart: "ppt/slides/slide1.xml",
      renderWarnings: [],
      elements: [{
        id: "cropped-logo",
        kind: "image",
        name: "Cropped logo",
        x: 100,
        y: 100,
        width: 400,
        height: 200,
        rotation: 0,
        geometry: "rect",
        mediaId: "logo.svg",
        sourceCropped: true,
        sourceCrop: { left: .1, top: .2, right: .1, bottom: .2 },
      }],
    }],
  };

  const svg = slidePreviewSvg(catalog, catalog.slides[0], 1_000);
  assert.match(svg, /clipPath id="crop-cropped-logo"/);
  assert.match(svg, /x="50" y="33\.3333333333333\d" width="500" height="333\.333333333333\d+"/);
  assert.match(svg, /preserveAspectRatio="none"/);
});
