# Presentation Studio quality pipeline

Presentation Studio separates software correctness from PowerPoint-native visual qualification. A green portable build is necessary, but it is not evidence that a presentation looks correct or better.

## Portable lane

Run `npm run quality` locally. GitHub Actions runs the same checks with `npm run quality -- --ci`:

1. TypeScript checking (`npm run lint`)
2. Unit and integration tests (`npm test`)
3. Production renderer build (`npm run build`)
4. Repository data-safety scan (`npm run check:data-safety`)

The lane works without Microsoft PowerPoint. Its report is written to ignored local storage at `tmp/quality-pipeline/latest.json` and explicitly records that no client content is included. It does not prove Office rendering, font metrics, editable-output geometry, visual hierarchy, template fidelity, or aesthetic improvement.

The same hosted workflow also runs the published one-line Windows source installer on `windows-latest` with the runner's Node.js paths removed. That smoke test must download and checksum-verify the portable Node.js prerequisite, install the release archive without Git, pass the staged application checks, and create the managed-install marker and Windows launcher. It uses no presentation or client data.

## PowerPoint-native lane

Run `npm run quality:native` only on an approved workstation with the required local Microsoft PowerPoint integration. It adds the Electron smoke test and the generated synthetic precision-layout canary. The canary covers native rendering and measurement, exact content/table preservation, optical geometry, clipping, safe regions, table clearances, proposal/export raster equivalence, and bounded repair behavior.

The manual GitHub workflow is intentionally limited to a self-hosted macOS runner labeled `presentation-studio-powerpoint`. Repository workflow configuration does not provision, approve, patch, or secure that workstation. The runner owner must keep PowerPoint licensed and current, restrict repository access, prevent forked pull requests from using the runner, and review logs before retention.

No workflow uploads a client presentation, authorized ORNL template, `.pstudio` package, private-golden manifest, native render, qualification image, or HTML/JSON qualification report as an artifact. Private deck qualification stays on the authorized local workstation and is launched from the app or with `npm run qualify:deck`.

## Private production-design benchmarks

Representative client-like decks may be used only as ignored local evidence. A production-design case should include the immutable source slide, the exact editable candidate, PowerPoint-native source/candidate PNGs, native measurements, and a machine-readable report. The report must prove exact visible text, exact native table structure where applicable, approved fonts, native render and measurement readiness, no text overflow, no table-cell clearance defect, no off-slide object, and material composition beyond typography. Those objective gates qualify a slide for full-size visual judgment; they do not establish that it is attractive, on-brand, or approved.

Image generation may provide a layout concept only from a sanitized structural brief. Do not send source slide bytes, wording, logos, figures, screenshots, data, or technical details. Treat the concept as untrusted visual influence and reconstruct the accepted hierarchy, spacing, and grouping with source-bound editable Studio objects. A source-locked PowerPoint group should be rendered through the object-isolation path so neighboring slide objects cannot leak into the fresh composition.

The hash-pinned private-golden runner makes this repeatable without adding client files to Git:

1. Copy `fixtures/private-golden-manifest.example.json` to an ignored local path.
2. Point it at the immutable source deck, a human-cleaned golden deck, and the installed authorized ORNL template. Record each SHA-256 and select 1–12 communication-archetype cases.
3. Run `npm run benchmark:private-golden -- --manifest /absolute/private/manifest.json --output /absolute/private/new-run-folder`.
4. Inspect `private-golden-review.html`. It contains only the selected PowerPoint-native source, Studio candidate, and golden slide pixels—not all slides in the benchmark deck.
5. Give a context-isolated model the generated `fresh-agent-prompt.txt` while the same source deck is open in Presentation Studio. Product or engine failures become regression work; the model must not use hidden developer context or silently work around them.
6. Treat `objectiveReady` as permission to perform visual review, never as an aesthetic pass. Source wins whenever the Studio candidate is weaker, and the human-cleaned deck is a quality reference rather than authority to change the immutable source copy.

Every run writes an editable candidate PPTX, a JSON ledger, an HTML triptych review, selected native PNGs, and a fresh-agent prompt to a new ignored output folder. The source, benchmark, template, and previous runs remain unchanged.

## Deck design loop

For an actual presentation, the required loop is:

1. Build the one current Studio Web Scene into the central editable candidate.
2. Run **Inspect all** or MCP `run_deck_qualification` once for those candidate bytes.
3. Inspect every page of the clean candidate contact sheet.
4. Inspect the clean full-size PowerPoint-native source and candidate image for every slide.
5. Use an issue crop or diagnostic overlay only when a finding needs localization.
6. Record a raster-bound review for every slide. Automated objective success means only ready for visual review.
7. Route slide-design findings through Studio/MCP, reproducible engine defects through synthetic or private regression, governed visual needs through text-free/logo-free concept work, and ambiguity to a person.
8. Rebuild the same central scene and create a new qualification attempt only when candidate bytes change.
9. Compare objective trend. Stop automatic repair after attempt three and hold unresolved work.
10. Before delivery, independently rerender and remeasure the exact export bytes and inspect the written artifact.

`review-complete` is evidence that every slide in one exact candidate was reviewed and no blocker or major finding remained. It is not official ORNL brand approval, technical-owner approval, accessibility certification, rights clearance, or authorization to publish.
