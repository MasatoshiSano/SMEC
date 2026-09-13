// Swiss Grid PowerPoint slide component library (pptxgenjs).
//
// This is the PowerPoint counterpart to docs/slide_template/*.html — the same
// design system (monochrome + red accent,罫線グリッド), reimplemented as
// reusable pptxgenjs functions so each subject's deck (scripts/pptx_slidegen/
// build_<subject>.js) only has to supply content, not re-derive layout math.
//
// See docs/13_pptx_slide_template_spec.md for the usage guide, the layout
// conventions (coordinate grid, column widths), and the "when to add a
// right-side diagram" decision rule.

const pptxgen = require("pptxgenjs");

// ---- palette / fonts ----
const INK = "171717";
const INK_SOFT = "8A8A8A";
const RED = "FF3B1F";
const LINE = "E6E6E6";
const GHOST = "F4F4F4";
const WHITE = "FFFFFF";

const F_HEAD = "Yu Gothic";       // Japanese-safe, ships with modern Office/Windows
const F_BODY = "Yu Gothic";
const F_MONO = "Consolas";        // for kicker/labels/years (latin/numeral only)

// Pre-blended solid grays for "white at N% opacity over the #171717 dark background".
// Google Slides does not reliably render <a:alpha> on text/line fills (it shows fully
// opaque instead), so instead of color:WHITE + transparency:N we bake the blend into a
// solid hex color up front — this renders identically in real PowerPoint and Google
// Slides. Add more only against this same #171717 background; a different dark
// background needs its own set.
const W_10 = "E8E8E8";  // white, 10% transparent (90% opacity)
const W_20 = "D1D1D1";  // white, 20% transparent (80% opacity)
const W_25 = "C5C5C5";  // white, 25% transparent (75% opacity)
const W_40 = "A2A2A2";  // white, 40% transparent (60% opacity)
const W_45 = "7F7F7F";  // white, 45% transparent (55% opacity)
const W_75 = "515151";  // white, 75% transparent (25% opacity)
const W_94 = "252525";  // white, 94% transparent (6% opacity) -- ghost numerals

const PAGE_W = 13.333, PAGE_H = 7.5; // 16:9 widescreen

function newPres() {
  const p = new pptxgen();
  p.defineLayout({ name: "WIDE", width: PAGE_W, height: PAGE_H });
  p.layout = "WIDE";
  return p;
}

// ---------- D1: cover slide (one per subject) ----------
function addCoverSlide(pres, { eyebrow, subjectNo, title, subtitle, stats, tag, notes }) {
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText(eyebrow, {
    x: 0.7, y: 0.55, w: 6, h: 0.35,
    fontFace: F_MONO, fontSize: 12, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText(subjectNo, {
    x: 9.5, y: 0.55, w: 3.1, h: 0.35, align: "right",
    fontFace: F_MONO, fontSize: 12, color: RED, isTextBox: true, margin: 0,
  });
  s.addText([
    { text: title, options: { color: WHITE } },
    { text: "。", options: { color: RED } },
  ], {
    x: 0.65, y: 2.3, w: 11, h: 1.5,
    fontFace: F_HEAD, fontSize: 66, bold: true, isTextBox: true, margin: 0,
  });
  s.addText(subtitle, {
    x: 0.7, y: 3.75, w: 8.4, h: 1.0,
    fontFace: F_BODY, fontSize: 15, color: W_25, isTextBox: true, margin: 0, lineSpacingMultiple: 1.5,
  });
  s.addShape("line", { x: 0.7, y: 6.35, w: 11.9, h: 0, line: { color: W_75, width: 0.75 } });
  stats.forEach(([num, lbl], i) => {
    const x = 0.7 + i * 2.1;
    s.addText(num, { x, y: 6.55, w: 1.9, h: 0.5, fontFace: F_MONO, fontSize: 24, bold: true, color: RED, isTextBox: true, margin: 0 });
    s.addText(lbl, { x, y: 7.05, w: 1.9, h: 0.3, fontFace: F_BODY, fontSize: 10, color: W_25, isTextBox: true, margin: 0 });
  });
  s.addShape("rect", { x: 8.6, y: 6.75, w: 4.0, h: 0.42, fill: { type: "none" }, line: { color: W_45, width: 0.75 } });
  s.addText(tag, {
    x: 8.6, y: 6.75, w: 4.0, h: 0.42, align: "center", valign: "middle",
    fontFace: F_MONO, fontSize: 8.5, color: WHITE, isTextBox: true, margin: 0,
  });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------- D2: section divider slide (one per major grouping, e.g. 経営戦略論/組織論/…) ----------
function addDividerSlide(pres, { ghostNo, partNo, partLabel, title, desc, chips, notes }) {
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText(ghostNo, {
    x: 8.8, y: 1.6, w: 4.3, h: 4.5, align: "right", valign: "middle",
    fontFace: F_MONO, fontSize: 220, bold: true, color: W_94, isTextBox: true, margin: 0,
  });
  s.addText([
    { text: partNo, options: { color: RED, bold: true } },
    { text: "   " + partLabel, options: { color: W_40 } },
  ], {
    x: 0.7, y: 2.1, w: 9, h: 0.4,
    fontFace: F_MONO, fontSize: 13, isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: 0.65, y: 2.55, w: 9, h: 1.1,
    fontFace: F_HEAD, fontSize: 46, bold: true, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText(desc, {
    x: 0.7, y: 3.75, w: 8.2, h: 1.3,
    fontFace: F_BODY, fontSize: 14, color: W_20, isTextBox: true, margin: 0, lineSpacingMultiple: 1.6,
  });
  let cx = 0.7;
  chips.forEach((c) => {
    const cw = 0.35 + c.length * 0.135;
    s.addShape("rect", { x: cx, y: 5.25, w: cw, h: 0.42, fill: { type: "none" }, line: { color: W_45, width: 0.75 } });
    s.addText(c, {
      x: cx, y: 5.25, w: cw, h: 0.42, align: "center", valign: "middle",
      fontFace: F_MONO, fontSize: 10, color: W_10, isTextBox: true, margin: 0,
    });
    cx += cw + 0.25;
  });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------- shared header for content slides ----------
function addHeader(slide, { kicker, title, overview, tag }) {
  slide.background = { color: WHITE };
  slide.addText(kicker, {
    x: 0.55, y: 0.35, w: 9.5, h: 0.3,
    fontFace: F_MONO, fontSize: 11, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  slide.addText(title, {
    x: 0.53, y: 0.62, w: 9.7, h: 0.55,
    fontFace: F_HEAD, fontSize: 25, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  slide.addText(overview, {
    x: 0.55, y: 1.18, w: 11.7, h: 0.4,
    fontFace: F_BODY, fontSize: 13, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  // tag, top right
  slide.addShape("rect", {
    x: 10.6, y: 0.4, w: 2.25, h: 0.36, fill: { color: INK }, line: { type: "none" },
  });
  slide.addText(tag, {
    x: 10.6, y: 0.4, w: 2.25, h: 0.36,
    fontFace: F_MONO, fontSize: 10.5, color: WHITE, align: "center", valign: "middle",
    isTextBox: true, margin: 0,
  });
  // rule under header
  slide.addShape("line", {
    x: 0.53, y: 1.62, w: 12.27, h: 0,
    line: { color: INK, width: 1.25 },
  });
}

// ---------- shared frequency data bar (real 頻出ランク・出題年度 data — never fabricated) ----------
function addFreqBar(slide, { y, rank, rankLabel, related, years }) {
  slide.addShape("line", { x: 0.55, y, w: 12.25, h: 0, line: { color: INK, width: 0.75 } });
  const topY = y + 0.08;
  slide.addShape("ellipse", {
    x: 0.55, y: topY, w: 0.36, h: 0.36, fill: { color: RED }, line: { type: "none" },
  });
  slide.addText(rank, {
    x: 0.55, y: topY, w: 0.36, h: 0.36, align: "center", valign: "middle",
    fontFace: F_MONO, fontSize: 13, bold: true, color: WHITE, isTextBox: true, margin: 0,
  });
  slide.addText(rankLabel, {
    x: 1.0, y: topY, w: 1.5, h: 0.36, valign: "middle",
    fontFace: F_BODY, fontSize: 10.5, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  slide.addShape("line", { x: 2.55, y: topY, w: 0, h: 0.36, line: { color: LINE, width: 1 } });
  slide.addText(related, {
    x: 2.75, y: topY - 0.02, w: 10.05, h: 0.4, valign: "middle",
    fontFace: F_BODY, fontSize: 10, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  const tlY = topY + 0.44;
  slide.addText("出題年度（過去10年）", {
    x: 0.55, y: tlY, w: 1.55, h: 0.24, valign: "middle",
    fontFace: F_BODY, fontSize: 8.5, color: INK_SOFT, isTextBox: true, margin: 0,
  });
  const startX = 2.25, spanW = 10.55, n = years.length, colW = spanW / n;
  years.forEach((yr, i) => {
    const cx = startX + i * colW + colW / 2;
    const on = yr.on;
    slide.addShape("ellipse", {
      x: cx - 0.08, y: tlY, w: 0.16, h: 0.16,
      fill: { color: on ? RED : WHITE },
      line: { color: on ? RED : LINE, width: 1 },
    });
    slide.addText(yr.label, {
      x: cx - colW / 2, y: tlY + 0.17, w: colW, h: 0.19, align: "center",
      fontFace: F_MONO, fontSize: 8, color: on ? INK : INK_SOFT, bold: on,
      isTextBox: true, margin: 0,
    });
  });
}

const YEARS_16_25 = ["'16", "'17", "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25"];
// onSet: a Set of the label strings (e.g. "'19") the topic was tested in — pull this
// straight from problem_sets/1st_stage/<subject>.md's real 出題年度 data.
function mkYears(onSet) {
  return YEARS_16_25.map((label) => ({ label, on: onSet.has(label) }));
}

// ---------- term-row list (bold label | gray description), for 2-4 short definitions ----------
function addTermRows(slide, x, y, w, rows, { fontSize = 11.5, labelW = 1.7, gap = 0.3 } = {}) {
  let cy = y;
  rows.forEach((r) => {
    // Per-row `gap` override: a row whose value text needs 2+ lines at this width
    // needs a taller slot than a short one-line row — see docs/13 for how to size this.
    const rowH = (r.gap || gap) - 0.02;
    slide.addText(r.k, {
      x, y: cy, w: labelW, h: rowH, valign: "top",
      fontFace: F_BODY, fontSize, bold: true, color: INK, isTextBox: true, margin: 0,
    });
    slide.addShape("line", { x: x + labelW - 0.02, y: cy, w: 0, h: 0.28, line: { color: LINE, width: 1 } });
    slide.addText(r.v, {
      x: x + labelW + 0.15, y: cy, w: w - labelW - 0.15, h: rowH, valign: "top",
      fontFace: F_BODY, fontSize, color: INK_SOFT, isTextBox: true, margin: 0,
    });
    cy += (r.gap || gap);
  });
  return cy;
}

// ---------- generic multi-row list: bold name | mono tag | gray description ----------
// (the pattern used for e.g. the Ansoff growth-matrix's 4 strategies; reusable for any
// "N named items, each with a short code and a one-line explanation" content)
function addRowList(slide, x, y, w, rows, { rowH = 0.92, nameW = 2.6, tagW = 2.55 } = {}) {
  let cy = y;
  rows.forEach((r, i) => {
    if (i > 0) slide.addShape("line", { x, y: cy, w, h: 0, line: { color: LINE, width: 1 } });
    slide.addText(r.name, { x, y: cy + 0.12, w: nameW, h: 0.35, fontFace: F_HEAD, fontSize: 15, bold: true, color: INK, isTextBox: true, margin: 0 });
    if (r.tag) {
      slide.addText(r.tag, { x: x + nameW + 0.1, y: cy + 0.14, w: tagW, h: 0.3, fontFace: F_MONO, fontSize: 11, color: RED, isTextBox: true, margin: 0 });
    }
    const descX = x + nameW + 0.1 + (r.tag ? tagW + 0.1 : 0);
    slide.addText(r.desc, {
      x: descX, y: cy + 0.08, w: x + w - descX, h: rowH - 0.17, valign: "middle",
      fontFace: F_BODY, fontSize: 11, color: INK_SOFT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
    });
    cy += rowH;
  });
  return cy;
}

// ---------- exam-check question/answer renderer (実際の過去問をそのまま提示する) ----------
// Two-slide pattern like the HTML deck's .examq: call once with no `correctIndex` for
// the question-only slide, then again with `correctIndex` + explain/relatedNote for the
// answer slide. Never fabricate a question — only use real past_exams content, skip the
// topic's exam-check pair entirely if no matching exercise exists yet.
function addExamQuestion(slide, { stem, choices, correctIndex = -1, stemH = 1.0 }) {
  let cy = 1.9;
  if (stem) {
    slide.addText(stem, {
      x: 0.55, y: cy, w: 12.25, h: stemH,
      fontFace: F_BODY, fontSize: 12.5, color: INK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3,
    });
    cy += stemH + 0.05;
  }
  choices.forEach((c, i) => {
    const isCorrect = i === correctIndex;
    const rowH = 0.52;
    if (isCorrect) {
      slide.addShape("rect", { x: 0.55, y: cy - 0.03, w: 12.25, h: rowH, fill: { color: GHOST }, line: { type: "none" } });
      slide.addShape("line", { x: 0.55, y: cy - 0.03, w: 0, h: rowH, line: { color: RED, width: 3 } });
    }
    slide.addText([
      { text: c.badge + "  ", options: { bold: true, color: isCorrect ? RED : INK_SOFT, fontFace: F_BODY } },
      { text: c.text, options: { color: INK, fontFace: F_BODY } },
    ], {
      x: 0.8, y: cy, w: 11.8, h: rowH - 0.05, valign: "middle",
      fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
    });
    cy += rowH;
  });
  return cy;
}

// ---------- right-column diagrams ----------
// Add a diagram to the right third of a content slide ONLY when the content has
// genuinely spatial/comparative structure that a picture adds information to — not as
// a default decoration on every slide. See docs/13_pptx_slide_template_spec.md for the
// decision rule and worked examples (when C-1/C-2/C-3 each did or didn't get one).

// two bars contrasting a narrow vs. wide scope/definition
function drawScopeBars(slide, x, y, w, h, { narrowLabel, wideLabel, resultLabel, narrowFrac = 0.42 } = {}) {
  const barH = 0.55;
  const labelH = 0.42;
  const gap = 0.55;
  const narrowW = w * narrowFrac;
  let cy = y;
  slide.addText(narrowLabel, {
    x, y: cy, w, h: labelH, fontFace: F_BODY, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  cy += labelH;
  slide.addShape("rect", { x, y: cy, w: narrowW, h: barH, fill: { color: GHOST }, line: { color: INK, width: 1.25 } });
  cy += barH + gap;
  slide.addText(wideLabel, {
    x, y: cy, w, h: labelH, fontFace: F_BODY, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  cy += labelH;
  slide.addShape("rect", { x, y: cy, w, h: barH, fill: { color: GHOST }, line: { color: RED, width: 1.75 } });
  cy += barH + 0.2;
  if (resultLabel) {
    slide.addText(resultLabel, {
      x, y: cy, w, h: 0.35, fontFace: F_BODY, fontSize: 10, bold: true, color: RED, isTextBox: true, margin: 0,
    });
  }
}

// generic 2x2 matrix (PPM, Porter's generic strategies, etc.) — four labeled quadrants
// plus a one-line caption underneath explaining what the axes mean (kept as plain text
// rather than rotated axis labels, which pptxgenjs handles less predictably).
function drawQuadrant(slide, x, y, w, h, { cells, axisCaption }) {
  const gap = 0.06;
  const gridH = h - 0.4;
  const cw = (w - gap) / 2, ch = (gridH - gap) / 2;
  const pos = {
    tl: [x, y], tr: [x + cw + gap, y],
    bl: [x, y + ch + gap], br: [x + cw + gap, y + ch + gap],
  };
  cells.forEach((cell) => {
    const [cx, cy] = pos[cell.pos];
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: GHOST }, line: { color: INK, width: 1 } });
    slide.addText([
      { text: cell.label + "\n", options: { bold: true, fontSize: 12, color: INK, breakLine: true } },
      { text: cell.sublabel || "", options: { fontSize: 8.5, color: INK_SOFT } },
    ], {
      x: cx + 0.08, y: cy, w: cw - 0.16, h: ch, align: "center", valign: "middle",
      fontFace: F_BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
    });
  });
  if (axisCaption) {
    slide.addText(axisCaption, {
      x, y: y + gridH + 0.1, w, h: 0.3, align: "center",
      fontFace: F_BODY, fontSize: 8, color: INK_SOFT, isTextBox: true, margin: 0,
    });
  }
}

// Porter's five forces — a center box (the industry) with four surrounding forces
// connected by thin spokes. Use for any "one central thing pressured from N directions"
// content, not just 5 forces specifically.
function draw5Forces(slide, x, y, w, h, { center, top, bottom, left, right }) {
  // 3x3 grid (like the HTML .forces-wrap CSS grid): side columns narrower than the
  // center column, three equal rows. This mirrors the fixed-size-columns approach
  // rather than centering boxes by half-widths, which previously produced overlapping
  // boxes when the "center" box was wider than the side gaps allowed.
  const gapX = 0.15, gapY = 0.12;
  const colSideW = w * 0.24;
  const colCenterW = w - 2 * colSideW - 2 * gapX;
  const rowH = (h - 2 * gapY) / 3;
  const col1X = x, col2X = x + colSideW + gapX, col3X = col2X + colCenterW + gapX;
  const row1Y = y, row2Y = y + rowH + gapY, row3Y = y + 2 * (rowH + gapY);
  const mkBox = (label, bx, by, bw, fill, textColor) => {
    slide.addShape("rect", { x: bx, y: by, w: bw, h: rowH, fill: { color: fill }, line: { color: INK, width: 1 } });
    slide.addText(label, {
      x: bx, y: by, w: bw, h: rowH, align: "center", valign: "middle",
      fontFace: F_BODY, fontSize: 9, bold: fill !== WHITE, color: textColor, isTextBox: true, margin: 0, lineSpacingMultiple: 1.05,
    });
  };
  mkBox(top, col2X, row1Y, colCenterW, WHITE, INK);
  mkBox(left, col1X, row2Y, colSideW, WHITE, INK);
  mkBox(center, col2X, row2Y, colCenterW, RED, WHITE);
  mkBox(right, col3X, row2Y, colSideW, WHITE, INK);
  mkBox(bottom, col2X, row3Y, colCenterW, WHITE, INK);
}

// two overlapping circles — relatedness → the overlap is the payoff (synergy, shared risk, etc.)
function drawVennOverlap(slide, x, y, w, h, { leftLabel, rightLabel, overlapLabel }) {
  const r = Math.min(w, h) * 0.34;
  const cy = y + h * 0.42;
  const overlap = r * 0.7;
  const c1x = x + w / 2 - r + overlap / 2;
  const c2x = x + w / 2 - r - overlap / 2 + r * 2;
  slide.addShape("ellipse", { x: c1x - r, y: cy - r, w: r * 2, h: r * 2, fill: { type: "none" }, line: { color: INK, width: 1.75 } });
  slide.addShape("ellipse", { x: c2x - r, y: cy - r, w: r * 2, h: r * 2, fill: { type: "none" }, line: { color: RED, width: 1.75 } });
  slide.addText(leftLabel, {
    x: c1x - r - 0.3, y: cy - r - 0.35, w: r * 2, h: 0.3, align: "center",
    fontFace: F_BODY, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  slide.addText(rightLabel, {
    x: c2x - r + 0.3, y: cy - r - 0.35, w: r * 2, h: 0.3, align: "center",
    fontFace: F_BODY, fontSize: 10, bold: true, color: RED, isTextBox: true, margin: 0,
  });
  slide.addText(overlapLabel, {
    x: x + w / 2 - 0.7, y: cy + r + 0.15, w: 1.4, h: 0.3, align: "center",
    fontFace: F_MONO, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0,
  });
}

// N-step horizontal process flow with arrows (e.g. Lewin's unfreeze-change-refreeze).
function drawProcessSteps(slide, x, y, w, h, steps) {
  const gap = 0.35;
  const n = steps.length;
  const stepW = (w - gap * (n - 1)) / n;
  steps.forEach((step, i) => {
    const sx = x + i * (stepW + gap);
    slide.addShape("rect", { x: sx, y, w: stepW, h, fill: { color: GHOST }, line: { color: LINE, width: 1 } });
    slide.addText([
      { text: step.num + "\n", options: { fontFace: F_MONO, fontSize: 10, bold: true, color: RED, breakLine: true } },
      { text: step.label + "\n", options: { fontFace: F_HEAD, fontSize: 14, bold: true, color: INK, breakLine: true } },
      { text: step.desc, options: { fontFace: F_BODY, fontSize: 9.5, color: INK_SOFT } },
    ], {
      x: sx + 0.15, y, w: stepW - 0.3, h, align: "center", valign: "middle",
      isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
    });
    if (i < n - 1) {
      slide.addText("→", {
        x: sx + stepW, y, w: gap, h, align: "center", valign: "middle",
        fontFace: F_BODY, fontSize: 18, color: INK_SOFT, isTextBox: true, margin: 0,
      });
    }
  });
}

// N-step vertical process flow with down-arrows — for the narrow right-column
// diagram slot, where drawProcessSteps' horizontal boxes would be too cramped.
function drawProcessStepsVertical(slide, x, y, w, h, steps) {
  const gap = 0.22;
  const n = steps.length;
  const stepH = (h - gap * (n - 1)) / n;
  steps.forEach((step, i) => {
    const sy = y + i * (stepH + gap);
    slide.addShape("rect", { x, y: sy, w, h: stepH, fill: { color: GHOST }, line: { color: LINE, width: 1 } });
    const runs = [];
    if (step.num) runs.push({ text: step.num + "\n", options: { fontFace: F_MONO, fontSize: 9, bold: true, color: RED, breakLine: true } });
    runs.push({ text: step.label, options: { fontFace: F_HEAD, fontSize: 16, bold: true, color: INK } });
    if (step.desc) runs.push({ text: "\n" + step.desc, options: { fontFace: F_BODY, fontSize: 10, bold: true, color: RED, breakLine: true } });
    slide.addText(runs, {
      x, y: sy, w, h: stepH, align: "center", valign: "middle",
      isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
    });
    if (i < n - 1) {
      slide.addText("↓", {
        x, y: sy + stepH, w, h: gap, align: "center", valign: "middle",
        fontFace: F_BODY, fontSize: 13, color: INK_SOFT, isTextBox: true, margin: 0,
      });
    }
  });
}

// ---------- 経済学専用グラフ（座標軸＋直線で表す需要供給曲線・IS-LM分析・
// 45度線分析など）。pptxgenjsに滑らかな曲線を描く手段がないため、教科書の
// 慣例どおりどれも直線（右下がり/右上がりの直線）で描く。座標は0〜1に正規化
// したチャート内座標(nx,ny)で指定し、原点(0,0)=左下、(1,1)=右上として扱う。 ----------

function drawLineSeg(slide, ax1, ay1, ax2, ay2, opts = {}) {
  const x = Math.min(ax1, ax2), y = Math.min(ay1, ay2);
  const w = Math.abs(ax2 - ax1) || 0.0001, h = Math.abs(ay2 - ay1) || 0.0001;
  const flipV = (ax2 - ax1) * (ay2 - ay1) < 0;
  slide.addShape("line", { x, y, w, h, flipV, line: { color: opts.color || INK, width: opts.width || 1.5, dashType: opts.dashType } });
}

// x,y,w,h: bounding box in inches. axis: {xLabel, yLabel}. series: [{x1,y1,x2,y2 (normalized 0-1),
// color, width, dashType, label, labelDx, labelDy}]. point: {nx, ny, label, guides: true}.
function drawLineChart(slide, x, y, w, h, { xLabel, yLabel, series = [], point, gridPad = 0.35 } = {}) {
  const ax = x + gridPad, ay = y + h, bx = x + w, by = y + gridPad;
  // axis lines with arrowheads
  slide.addShape("line", { x: ax, y: ay, w: bx - ax, h: 0, line: { color: INK, width: 1.6, endArrowType: "triangle" } });
  slide.addShape("line", { x: ax, y: by, w: 0, h: ay - by, line: { color: INK, width: 1.6, endArrowType: "triangle" }, flipV: true });
  if (xLabel) slide.addText(xLabel, { x: bx - 1.1, y: ay - 0.32, w: 1.1, h: 0.3, align: "right", fontFace: F_BODY, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0 });
  if (yLabel) slide.addText(yLabel, { x: ax + 0.05, y: by - 0.05, w: 1.1, h: 0.28, align: "left", fontFace: F_BODY, fontSize: 10, bold: true, color: INK, isTextBox: true, margin: 0 });

  const toX = (nx) => ax + nx * (bx - ax);
  const toY = (ny) => ay - ny * (ay - by);

  series.forEach((s) => {
    const x1 = toX(s.x1), y1 = toY(s.y1), x2 = toX(s.x2), y2 = toY(s.y2);
    drawLineSeg(slide, x1, y1, x2, y2, { color: s.color || INK, width: s.width || 2, dashType: s.dashType });
    if (s.label) {
      slide.addText(s.label, {
        x: x2 + (s.labelDx || 0.05), y: y2 + (s.labelDy || -0.15), w: 0.6, h: 0.3,
        fontFace: F_BODY, fontSize: 11, bold: true, color: s.color || INK, isTextBox: true, margin: 0,
      });
    }
  });

  if (point) {
    const px = toX(point.nx), py = toY(point.ny);
    if (point.guides !== false) {
      drawLineSeg(slide, ax, py, px, py, { color: INK_SOFT, width: 0.75, dashType: "dash" });
      drawLineSeg(slide, px, ay, px, py, { color: INK_SOFT, width: 0.75, dashType: "dash" });
    }
    slide.addShape("ellipse", { x: px - 0.05, y: py - 0.05, w: 0.1, h: 0.1, fill: { color: INK }, line: { type: "none" } });
    if (point.label) {
      slide.addText(point.label, {
        x: px + 0.08, y: py - 0.32, w: 0.5, h: 0.28,
        fontFace: F_BODY, fontSize: 11, bold: true, color: INK, isTextBox: true, margin: 0,
      });
    }
  }
}

module.exports = {
  INK, INK_SOFT, RED, LINE, GHOST, WHITE,
  F_HEAD, F_BODY, F_MONO,
  W_10, W_20, W_25, W_40, W_45, W_75, W_94,
  PAGE_W, PAGE_H,
  newPres,
  addCoverSlide,
  addDividerSlide,
  addHeader,
  addFreqBar,
  mkYears,
  addTermRows,
  addRowList,
  addExamQuestion,
  drawScopeBars,
  drawVennOverlap,
  drawQuadrant,
  draw5Forces,
  drawProcessSteps,
  drawProcessStepsVertical,
  drawLineChart,
};
