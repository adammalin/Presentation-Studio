const fs = require("node:fs/promises");
const path = require("node:path");
const { renderPowerPointNative } = require("./native-render.cjs");
const { measurePowerPointNative } = require("./native-measurement.cjs");

const QUALIFICATION_SCHEMA = "presentation-studio/deck-qualification";

function safeSegment(value, fallback = "qualification") {
  const result = String(value || "").replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "").slice(0, 120);
  return result || fallback;
}

function assertAbsoluteOutputRoot(outputRoot) {
  if (typeof outputRoot !== "string" || !path.isAbsolute(outputRoot)) throw new Error("The qualification output folder must be an absolute local path.");
  return path.resolve(outputRoot);
}

function relativeImagePath(representation, slideNumber) {
  return `${representation}/slide-${String(slideNumber).padStart(3, "0")}.png`;
}

async function writeRenderImages(outputRoot, representation, render) {
  const folder = path.join(outputRoot, representation);
  await fs.mkdir(folder, { recursive: false, mode: 0o700 });
  const slides = [];
  for (const slide of render.slides || []) {
    const relativePath = relativeImagePath(representation, slide.number);
    await fs.writeFile(path.join(outputRoot, relativePath), Buffer.from(slide.bytes), { mode: 0o600 });
    slides.push({ number: slide.number, mimeType: "image/png", width: slide.width, height: slide.height, sha256: slide.sha256, relativePath });
  }
  return {
    status: render.status,
    renderer: render.renderer,
    authoritative: Boolean(render.authoritative),
    sourceSha256: render.sourceSha256,
    powerPointVersion: render.powerPointVersion,
    pipeline: render.pipeline,
    slideCount: render.slideCount,
    slides,
    warnings: render.warnings || [],
    reason: render.reason,
  };
}

function renderFailureSummary(render) {
  return {
    status: render.status,
    renderer: render.renderer,
    authoritative: Boolean(render.authoritative),
    sourceSha256: render.sourceSha256,
    powerPointVersion: render.powerPointVersion,
    pipeline: render.pipeline,
    slideCount: render.slideCount,
    slides: [],
    warnings: render.warnings || [],
    reason: render.reason,
  };
}

async function captureDeckQualification({ source, candidate, outputRoot, width = 2200, render = renderPowerPointNative, measure = measurePowerPointNative }) {
  const root = assertAbsoluteOutputRoot(outputRoot);
  if (!source?.bytes?.length || !candidate?.bytes?.length) throw new Error("Source and candidate PowerPoint bytes are required for qualification.");
  if (!Number.isInteger(width) || width < 1600 || width > 3000) throw new Error("Qualification PNG width must be an integer from 1,600 to 3,000 pixels.");
  await fs.mkdir(path.dirname(root), { recursive: true, mode: 0o700 });
  try {
    await fs.mkdir(root, { recursive: false, mode: 0o700 });
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error("The qualification output folder already exists. Use a new run ID or output path.");
    throw error;
  }
  const capture = {
    schema: "presentation-studio/deck-qualification-capture",
    version: 1,
    generatedAt: new Date().toISOString(),
    widthPx: width,
    source: undefined,
    candidate: undefined,
  };
  try {
    const sourceRender = await render({ bytes: source.bytes, name: source.name, width, format: "png" });
    capture.source = sourceRender.status === "ready" ? await writeRenderImages(root, "source", sourceRender) : renderFailureSummary(sourceRender);
    const sourceMeasurement = await measure({ bytes: source.bytes, name: source.name });
    await fs.writeFile(path.join(root, "source-measurements.json"), `${JSON.stringify(sourceMeasurement, null, 2)}\n`, { mode: 0o600 });

    const candidateRender = await render({ bytes: candidate.bytes, name: candidate.name, width, format: "png" });
    capture.candidate = candidateRender.status === "ready" ? await writeRenderImages(root, "candidate", candidateRender) : renderFailureSummary(candidateRender);
    const candidateMeasurement = await measure({ bytes: candidate.bytes, name: candidate.name });
    await fs.writeFile(path.join(root, "candidate-measurements.json"), `${JSON.stringify(candidateMeasurement, null, 2)}\n`, { mode: 0o600 });
    await fs.writeFile(path.join(root, "capture.json"), `${JSON.stringify(capture, null, 2)}\n`, { mode: 0o600 });
    return { outputRoot: root, sourceRender: capture.source, candidateRender: capture.candidate, sourceMeasurement, candidateMeasurement };
  } catch (error) {
    await fs.writeFile(path.join(root, "capture-error.txt"), `${error instanceof Error ? error.stack || error.message : String(error)}\n`, { mode: 0o600 }).catch(() => undefined);
    throw error;
  }
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
}

function issueCrop(issue, candidateImage) {
  const region = issue.evidenceRegion;
  if (!region || region.width <= 0 || region.height <= 0) return "";
  const imageWidth = 100 / region.width;
  const imageLeft = -region.x / region.width * 100;
  const imageTop = -region.y / region.height * 100;
  const aspectWidth = Math.max(1, region.width * candidateImage.width);
  const aspectHeight = Math.max(1, region.height * candidateImage.height);
  return `<details class="issue-detail"><summary>Native-pixel evidence region</summary><div class="issue-crop" style="aspect-ratio:${aspectWidth}/${aspectHeight}"><img loading="lazy" decoding="async" src="${escapeHtml(candidateImage.relativePath)}" alt="Candidate crop for ${escapeHtml(issue.id)}" style="width:${imageWidth}%;left:${imageLeft}%;top:${imageTop}%"></div><small>${escapeHtml(region.label)} · ${region.objectIds.length} bound object${region.objectIds.length === 1 ? "" : "s"}</small></details>`;
}

function issueList(report, slideNumber, candidateImage) {
  const issues = report.issues.filter((issue) => issue.slideNumber === slideNumber || issue.slideNumber === undefined);
  if (!issues.length) return '<p class="clear">No deterministic blockers. Full-size visual judgment is still required.</p>';
  return `<ul class="issues">${issues.map((issue) => `<li class="${escapeHtml(issue.severity)}"><strong>${escapeHtml(issue.message)}</strong><span>${escapeHtml(issue.evidence)}</span><small>Route: ${escapeHtml(issue.repairRoute)} · ${escapeHtml(issue.id)}</small>${issueCrop(issue, candidateImage)}</li>`).join("")}</ul>`;
}

function qualificationReportHtml(report) {
  const checks = Object.entries(report.checks).map(([name, passed]) => `<li class="${passed ? "pass" : "fail"}"><b>${passed ? "PASS" : "FAIL"}</b><span>${escapeHtml(name.replace(/([A-Z])/g, " $1"))}</span></li>`).join("");
  const overview = report.slides.map((slide) => `<a class="overview-slide ${slide.status}" href="#slide-${slide.slideNumber}"><img loading="lazy" decoding="async" src="${escapeHtml(slide.candidateImage.relativePath)}" alt="Candidate slide ${slide.slideNumber}"><span>${slide.slideNumber} · ${escapeHtml(slide.status.replaceAll("-", " "))}</span></a>`).join("");
  const slides = report.slides.map((slide) => `<article id="slide-${slide.slideNumber}">
    <header><div><span>Slide ${slide.slideNumber}</span><h2>${slide.protected ? "Protected template composition" : escapeHtml(slide.designImpact || "Visual review")}</h2></div><b class="status ${slide.status}">${escapeHtml(slide.status.replaceAll("-", " "))}</b></header>
    <div class="comparison"><figure><figcaption>Source · clean PowerPoint render</figcaption><a href="${escapeHtml(slide.sourceImage.relativePath)}"><img loading="lazy" decoding="async" src="${escapeHtml(slide.sourceImage.relativePath)}" alt="Source slide ${slide.slideNumber}"></a></figure><figure><figcaption>Candidate · exact export result</figcaption><div class="candidate-frame"><a href="${escapeHtml(slide.candidateImage.relativePath)}"><img loading="lazy" decoding="async" src="${escapeHtml(slide.candidateImage.relativePath)}" alt="Candidate slide ${slide.slideNumber}"></a>${report.issues.filter((issue) => issue.slideNumber === slide.slideNumber && issue.evidenceRegion).map((issue, index) => `<span class="diagnostic-box" style="left:${issue.evidenceRegion.x * 100}%;top:${issue.evidenceRegion.y * 100}%;width:${issue.evidenceRegion.width * 100}%;height:${issue.evidenceRegion.height * 100}%"><b>${index + 1}</b></span>`).join("")}</div></figure></div>
    ${slide.review ? `<section class="review ${slide.review.recordedVerdict}"><strong>Visual pass ${slide.review.pass} · ${escapeHtml(slide.review.recordedVerdict)}</strong><span>${escapeHtml(slide.review.rationale)}</span><small>${escapeHtml(slide.review.reviewer)} · ${escapeHtml(slide.review.reviewedAt)}</small></section>` : ""}
    ${issueList(report, slide.slideNumber, slide.candidateImage)}
  </article>`).join("\n");
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Presentation Studio qualification · ${escapeHtml(report.candidate.name)}</title><style>
  :root{font-family:Aptos,Arial,sans-serif;color:#373A36;background:#DBDCDB;line-height:1.35}*{box-sizing:border-box}body{margin:0}header.hero{padding:32px 4vw;background:#00454D;color:#fff;border-bottom:10px solid #00662C}.hero p{max-width:980px;margin:8px 0 0}.hero-actions{display:flex;gap:12px;align-items:center;margin-top:18px}.hero button{appearance:none;border:1px solid #FFFFFF;background:#FFFFFF;color:#00454D;padding:9px 12px;font:700 13px Aptos,Arial,sans-serif;cursor:pointer;border-radius:0}.summary{display:grid;grid-template-columns:repeat(6,minmax(120px,1fr));gap:1px;background:#DBDCDB;border:1px solid #DBDCDB;margin:24px 4vw}.summary div{background:#fff;padding:18px}.summary b{display:block;font-size:26px;color:#00662C}.overview{margin:24px 4vw;padding:20px;background:#FFFFFF;border:1px solid #DBDCDB}.overview h2{margin-top:0}.overview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px}.overview-slide{display:grid;color:#373A36;text-decoration:none;border:1px solid #DBDCDB;background:#FFFFFF}.overview-slide img{border:0}.overview-slide span{padding:7px;font-size:11px;font-weight:700}.overview-slide.objective-failure,.overview-slide.revise,.overview-slide.hold{border-color:#FE5000}.checks{margin:24px 4vw;padding:0;display:grid;grid-template-columns:repeat(3,minmax(200px,1fr));gap:1px;background:#DBDCDB;border:1px solid #DBDCDB;list-style:none}.checks li{background:#fff;padding:12px;display:flex;gap:10px}.checks b{color:#00662C}.checks .fail b{color:#FE5000}main{margin:0 4vw 60px}article{background:#fff;border:1px solid #DBDCDB;margin:24px 0;padding:20px}article>header{display:flex;justify-content:space-between;align-items:start;border-bottom:4px solid #00662C;padding-bottom:12px}h2{margin:2px 0 0;font-size:22px}.status{text-transform:uppercase;font-size:12px;padding:6px 8px;border:1px solid #DBDCDB}.status.objective-failure,.status.revise,.status.hold{background:#fff;color:#FE5000;border-color:#FE5000}.comparison{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:18px}figure{margin:0}figcaption{font-weight:700;margin-bottom:6px}img{display:block;width:100%;height:auto;border:1px solid #DBDCDB;background:#fff}.candidate-frame{position:relative}.diagnostic-box{display:none;position:absolute;border:4px solid #FE5000;pointer-events:none}.diagnostic-box b{position:absolute;left:-4px;top:-28px;min-width:24px;height:24px;padding:3px;background:#FE5000;color:#FFFFFF;text-align:center}.show-diagnostics .diagnostic-box{display:block}.issues{padding-left:22px}.issues>li{margin:12px 0}.issues li span,.issues li small{display:block}.blocker>strong,.major>strong{color:#FE5000}.issue-detail{margin-top:8px}.issue-detail summary{cursor:pointer;font-weight:700;color:#00454D}.issue-crop{position:relative;overflow:hidden;max-width:720px;max-height:460px;margin-top:8px;border:1px solid #DBDCDB;background:#FFFFFF}.issue-crop img{position:absolute;max-width:none;height:auto;border:0}.review{display:grid;gap:4px;margin-top:14px;padding:12px;border-left:5px solid #00662C;background:#FFFFFF}.review.revise,.review.hold{border-left-color:#FE5000}.clear{color:#00662C;font-weight:600}.footer{margin:30px 4vw;color:#373A36}@media(max-width:1100px){.summary{grid-template-columns:repeat(3,1fr)}}@media(max-width:900px){.summary,.checks,.comparison{grid-template-columns:1fr}.hero{padding:24px}}
  </style></head><body><header class="hero"><h1>Presentation Studio qualification</h1><p>${escapeHtml(report.candidate.name)} · ${escapeHtml(report.status.replaceAll("-", " "))}. Microsoft PowerPoint-native PNGs were rendered for every clean source/candidate slide from the exact artifacts. Objective checks make the deck reviewable; they do not prove the design is better.</p><div class="hero-actions"><button type="button" id="toggle-diagnostics">Show diagnostic regions</button><span>Attempt ${report.iteration?.attempt ?? 1} · objective trend ${escapeHtml(report.iteration?.objectiveTrend ?? "first-run")}</span></div></header>
  <section class="summary"><div><span>Slides</span><b>${report.totals.slides}</b></div><div><span>Changed</span><b>${report.totals.changedSlides}</b></div><div><span>Blockers</span><b>${report.totals.blockerIssues}</b></div><div><span>Major issues</span><b>${report.totals.majorIssues}</b></div><div><span>Reviewed</span><b>${report.visualAcceptance.reviewedSlideCount ?? 0}</b></div><div><span>Ready</span><b>${report.visualAcceptance.readySlideCount ?? 0}</b></div></section>
  <section class="overview"><h2>Candidate deck overview</h2><p>Use this grid for pacing and consistency. Open each slide below for clean full-size source/candidate comparison.</p><div class="overview-grid">${overview}</div></section>
  <ul class="checks">${checks}</ul><main>${slides}</main><p class="footer">Draft qualification evidence—not formal ORNL approval. Inspect every clean slide at full size and obtain the appropriate content and brand review before release.</p><script>document.getElementById("toggle-diagnostics").addEventListener("click",function(){document.body.classList.toggle("show-diagnostics");this.textContent=document.body.classList.contains("show-diagnostics")?"Hide diagnostic regions":"Show diagnostic regions";});</script></body></html>`;
}

async function finalizeDeckQualification({ outputRoot, report }) {
  const root = assertAbsoluteOutputRoot(outputRoot);
  if (report?.schema !== QUALIFICATION_SCHEMA || report?.version !== 1) throw new Error("Use a Presentation Studio deck-qualification version 1 report.");
  await fs.writeFile(path.join(root, "qualification.json"), `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
  await fs.writeFile(path.join(root, "qa-report.html"), qualificationReportHtml(report), { mode: 0o600 });
  return { outputRoot: root, reportPath: path.join(root, "qualification.json"), htmlPath: path.join(root, "qa-report.html") };
}

async function readQualificationEvidence({ outputRoot, representation, slideNumber }) {
  const root = assertAbsoluteOutputRoot(outputRoot);
  if (!['source', 'candidate'].includes(representation)) throw new Error("Qualification evidence must be source or candidate.");
  if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > 1000) throw new Error("A valid qualification slide number is required.");
  const filePath = path.join(root, relativeImagePath(representation, slideNumber));
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(`${root}${path.sep}`)) throw new Error("Qualification evidence escaped the run folder.");
  const bytes = await fs.readFile(resolved);
  return { mimeType: "image/png", bytes: new Uint8Array(bytes), filePath: resolved };
}

module.exports = {
  captureDeckQualification,
  finalizeDeckQualification,
  qualificationReportHtml,
  readQualificationEvidence,
  relativeImagePath,
  safeSegment,
};
