#!/usr/bin/env node
/**
 * "Process Timeline" deck - 16 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only. Run: node <thisfile>  ->  writes the .pptx next to it.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const FONT = 'Roboto';

const C = {
  red: 'FD536D',      // accent1
  orange: 'FF8957',   // accent2
  yellow: 'EED054',   // accent3
  lime: 'CAD849',     // accent4
  green: '00C182',    // accent5
  blue: '429EB0',     // accent6
  gray: '5C5C5C',     // tx1 lum75
  dark: '3C3C3C',
  ink: '262626',
  band: 'F2F2F2',
  rule: 'D9D9D9',
  silver: 'BFBFBF',
  steel: 'A6A6A6',
  white: 'FFFFFF',
};
const ACCENTS = [C.red, C.orange, C.yellow, C.lime, C.green, C.blue];

const NOLINE = { type: 'none' };
// pptxgenjs text margin order is [left, right, bottom, top], in points
const PPT_INSET = [7.2, 7.2, 3.6, 3.6];

// drop shadows, in points, matching the source deck's outerShdw settings.
// pptxgenjs rewrites the object it is handed, so every call site gets a fresh one.
const shadow = (blur, offset, opacity) => ({ type: 'outer', blur, offset, angle: 90, color: '000000', opacity });
const SH_CARD = () => shadow(10, 2, 0.10);   // slide 1 white cards
const SH_RING = () => shadow(10, 1, 0.20);   // rings and outlined chevrons
const SH_DISC = () => shadow(10, 2, 0.20);   // slides 14/15 white discs
const SH_BADGE = () => shadow(5, 1, 0.30);   // slide 4 cards, slide 5 inner circles
const SH_PANEL = () => shadow(5, 0, 0.15);   // slide 5 description cards
const SH_GLOW = () => shadow(20, 0, 0.30);   // slide 6 numbered ovals and dots
const SH_PILL = () => shadow(20, 0, 0.23);   // slide 6 pills
const SH_DOT = () => shadow(10, 0, 0.25);    // slide 9 hub dots

/* ---------------------------------------------------------------- helpers */

const T = (o) => Object.assign({ fontFace: FONT, color: C.gray, margin: 0, valign: 'top' }, o);

/** plain text box (no fill, no line) */
function txt(slide, text, o) {
  slide.addText(text, T(o));
}

/** filled shape carrying centered text */
function shapeText(slide, shape, text, o) {
  slide.addText(text, Object.assign({ shape, fontFace: FONT, align: 'center', valign: 'middle', line: NOLINE }, o));
}

/** solid shape without text */
function box(slide, shape, o) {
  slide.addShape(shape, Object.assign({ line: NOLINE }, o));
}

/** custom geometry from normalised [0..1] points */
function poly(slide, x, y, w, h, pts, o) {
  slide.addShape('custGeom', Object.assign(
    { x, y, w, h, line: NOLINE, points: pts.map((p) => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }]) },
    o));
}

/** circle centred on (cx, cy) */
function dot(slide, cx, cy, d, o) {
  box(slide, 'ellipse', Object.assign({ x: cx - d / 2, y: cy - d / 2, w: d, h: d }, o));
}

/** straight connector; pts are [[x1,y1],[x2,y2],...] - last segment may carry an arrow */
function polyline(slide, pts, lineOpts) {
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i];
    const [x2, y2] = pts[i + 1];
    const opt = Object.assign({}, lineOpts);
    if (i < pts.length - 2) delete opt.endArrowType;
    slide.addShape('line', {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      flipH: x2 < x1, flipV: y2 < y1, line: opt,
    });
  }
}

/** stand-in emblem for the deck's decorative vector icons */
function icon(slide, cx, cy, size, color, glyph) {
  box(slide, glyph || 'star5', { x: cx - size / 2, y: cy - size / 2, w: size, h: size, fill: { color } });
}

/** slide title + subtitle, straight from the "Main Title & Subtitle" layout */
function heading(pptx, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  slide.addText('Process Timeline', T({
    x: 0.424, y: 0.171, w: 9.152, h: 0.542, fontSize: 26, align: 'center', valign: 'middle',
  }));
  slide.addText(subtitle, T({
    x: 0.424, y: 0.727, w: 9.152, h: 0.189, fontSize: 10, align: 'center', valign: 'middle',
  }));
  return slide;
}

/* ------------------------------------------------- reusable path fragments */

// zig-zag ribbon that steps down (slide 1)
const RIBBON_DOWN = [[0.90, 0.21], [0.12, 0.28], [0, 0.04], [0, 0.83], [0.02, 0.95],
  [0.08, 1], [0.92, 1], [0.98, 0.95], [1, 0.83], [1, 0]];
// zig-zag ribbon that steps up (slide 1)
const RIBBON_UP = [[0.10, 0.755], [0.883, 0.755], [1, 1], [1, 0.17], [0.98, 0.05],
  [0.92, 0], [0.08, 0], [0.02, 0.05], [0, 0.17], [0, 0.966]];
// broad arrow used as the last timeline segment (slide 3)
const FAT_ARROW = [[0, 0.353], [0.763, 0.353], [0.687, 0.188], [0.723, 0], [0.775, 0], [0.809, 0.032],
  [0.985, 0.424], [1, 0.5], [0.985, 0.578], [0.809, 0.969], [0.775, 1], [0.723, 1],
  [0.687, 0.814], [0.763, 0.648], [0, 0.648]];
// L-shaped arrow: stem rises on the left, head points right at the top (slide 11)
const HOOK_UP = [[0.048, 1], [0.048, 0.094], [0.062, 0.083], [0.731, 0.083], [0.731, 0.143],
  [1, 0.072], [0.731, 0], [0.731, 0.060], [0.062, 0.060], [0, 0.072], [0, 1]];
// same hook mirrored: head points right at the bottom (slide 11)
const HOOK_DOWN = HOOK_UP.map((p) => [p[0], 1 - p[1]]);
// leaning parallelogram chip (slide 13)
const SLANT = [[0, 1], [0.882, 1], [1, 0], [0.119, 0]];
// column with rounded top corners (slide 16)
function columnPts(w, h, r) {
  const rx = r / w; const ry = r / h;
  return [[1, 1], [0, 1], [0, ry], [0.3 * rx, 0.3 * ry], [rx, 0], [1 - rx, 0], [1 - 0.3 * rx, 0.3 * ry], [1, ry]];
}

/* ============================================================== slide  1 */
/* four alternating cards linked by red/orange/yellow/lime ribbons          */

function slide01(pptx) {
  const slide = heading(pptx, 'Insert subtitle here');
  const BODY = 'PLACEHOLDER';

  const ribbons = [
    { x: 0.389, y: 3.189, w: 2.292, h: 1.097, pts: RIBBON_DOWN, color: C.red },
    { x: 2.681, y: 2.137, w: 2.292, h: 1.095, pts: RIBBON_UP, color: C.orange },
    { x: 4.969, y: 3.189, w: 2.292, h: 1.097, pts: RIBBON_DOWN, color: C.yellow },
    { x: 7.260, y: 2.137, w: 2.293, h: 1.095, pts: RIBBON_UP, color: C.lime },
  ];
  ribbons.forEach((r) => poly(slide, r.x, r.y, r.w, r.h, r.pts, { fill: { color: r.color } }));

  const cards = [
    { card: [0.594, 1.509, 1.881, 2.606], pill: [0.845, 2.062], body: [0.778, 2.608], year: [1.138, 3.302, '2020'], circle: [1.535, 4.153], color: C.red, glyph: 'sun' },
    { card: [2.889, 2.285, 1.881, 2.348], pill: [3.137, 3.169], body: [3.070, 3.691], year: [3.433, 2.746, '2021'], circle: [3.827, 2.277], color: C.orange, glyph: 'flowChartProcess' },
    { card: [5.174, 1.509, 1.881, 2.606], pill: [5.425, 2.062], body: [5.358, 2.608], year: [5.718, 3.302, '2022'], circle: [6.115, 4.153], color: C.yellow, glyph: 'gear6' },
    { card: [7.468, 2.285, 1.881, 2.348], pill: [7.718, 3.169], body: [7.650, 3.691], year: [8.011, 2.746, '2023'], circle: [8.408, 2.277], color: C.lime, glyph: 'heart' },
  ];
  cards.forEach((c) => {
    box(slide, 'roundRect', {
      x: c.card[0], y: c.card[1], w: c.card[2], h: c.card[3],
      rectRadius: 0.07, fill: { color: C.white }, shadow: SH_CARD(),
    });
  });
  cards.forEach((c) => {
    dot(slide, c.circle[0], c.circle[1], 0.740, { fill: { color: C.white }, shadow: SH_RING() });
    dot(slide, c.circle[0], c.circle[1], 0.599, { fill: { color: c.color } });
    icon(slide, c.circle[0], c.circle[1], 0.28, C.white, c.glyph);

    shapeText(slide, 'roundRect', 'Insert Text here', {
      x: c.pill[0], y: c.pill[1], w: 1.379, h: 0.378, rectRadius: 0.063,
      fill: { color: c.color }, color: C.white, fontSize: 11, bold: true,
    });
    txt(slide, BODY, { x: c.body[0], y: c.body[1], w: 1.514, h: 0.454, fontSize: 9, align: 'center' });
    txt(slide, c.year[2], {
      x: c.year[0], y: c.year[1], w: 0.793, h: 0.333, fontSize: 14, bold: true,
      align: 'center', valign: 'middle',
    });
  });
}

/* ============================================================== slide  2 */
/* twelve-month gantt grid with task rows                                   */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const GRID_LINE = { type: 'solid', pt: 1, color: C.band };
const NO_LINE = { type: 'none' };
const BOX_BORDER = [GRID_LINE, GRID_LINE, GRID_LINE, GRID_LINE];
// vertical rules only (slide 7 ruler)
const COL_BORDER = [NO_LINE, GRID_LINE, NO_LINE, NO_LINE];
// margin is [top, right, bottom, left] in inches - PowerPoint's default cell insets.
// The 1pt default keeps empty cells from inflating their row.
const CELL = { valign: 'middle', margin: [0.05, 0.1, 0.05, 0.1], fontSize: 1, fontFace: FONT };

function gridCell(text, opts) {
  return { text: text || '', options: Object.assign({ border: BOX_BORDER }, CELL, opts) };
}

/** small coloured duration bar plus its "Resource Name" caption */
function ganttBar(slide, bar, res) {
  shapeText(slide, 'rect', bar[4], {
    x: bar[0], y: bar[1], w: bar[2], h: bar[3], fill: { color: bar[5] },
    color: C.white, fontSize: 8, bold: true, align: 'left', valign: 'middle',
  });
  if (res) txt(slide, 'Resource Name', { x: res[0], y: res[1], w: 0.920, h: 0.135, fontSize: 8, bold: true });
}

function slide02(pptx) {
  const slide = heading(pptx, '12 months timeline');

  const rows = [];
  for (let r = 0; r < 8; r++) {
    const row = [gridCell('Task ' + (r + 1), { color: C.gray, fontSize: 10, bold: true, fontFace: FONT })];
    for (let c = 0; c < 12; c++) row.push(gridCell(''));
    rows.push(row);
  }
  rows.push(new Array(13).fill(0).map(() => gridCell('')));
  rows.push(new Array(13).fill(0).map(() => gridCell('')));
  const monthRow = [gridCell('')];
  MONTHS.forEach((m, i) => monthRow.push(gridCell(m, {
    fill: { color: i === 0 ? C.gray : (i < 6 ? 'C9C9C9' : '929292') },
    color: C.white, fontSize: 10.5, bold: true, fontFace: FONT,
  })));
  rows.push(monthRow);

  slide.addTable(rows, {
    x: 0.25, y: 1.147, w: 9.5, colW: [1.357].concat(new Array(12).fill(0.679)), rowH: 0.373,
  });

  //           x      y      w      h     label       colour
  const bars = [
    [1.729, 1.265, 1.164, 0.170, '30 Days', C.red],
    [2.569, 1.609, 0.855, 0.170, '10 Days', C.orange],
    [3.123, 2.002, 1.355, 0.170, '40 Days', C.yellow],
    [4.198, 2.358, 1.164, 0.170, '30 Days', C.lime],
    [5.039, 2.702, 0.855, 0.170, '10 Days', C.green],
    [5.593, 3.095, 1.355, 0.170, '40 Days', C.blue],
    [6.667, 3.508, 0.855, 0.170, '10 Days', C.red],
    [7.221, 3.862, 1.355, 0.170, '40 Days', C.orange],
  ];
  const caps = [[2.932, 1.283], [3.801, 1.626], [4.561, 2.019], [5.402, 2.376],
    [6.270, 2.720], [7.031, 3.113], [7.585, 3.526], [8.659, 3.880]];
  bars.forEach((b, i) => ganttBar(slide, b, caps[i]));

  // two elbow links between dependent bars
  [[3.424, 1.694], [5.894, 2.787]].forEach(([ex, ey]) => {
    polyline(slide, [[ex, ey], [ex, ey + 0.196], [ex - 0.301, ey + 0.196], [ex - 0.301, ey + 0.393]],
      { color: C.gray, width: 0.75, endArrowType: 'triangle' });
  });

  // "today" marker
  slide.addShape('line', {
    x: 6.417, y: 1.350, w: 0, h: 2.796,
    line: { color: C.gray, width: 2.25, dashType: 'sysDash', endArrowType: 'triangle' },
  });

  [[0.886, 'Project Start', 1.672], [5.483, 'Milestone 2 (Today)', 6.270], [8.171, 'Project End', 8.973]]
    .forEach(([x, label, tx]) => {
      txt(slide, label, { x, y: 4.229, w: 1.866, h: 0.168, fontSize: 10, bold: true, align: 'center' });
      box(slide, 'triangle', { x: tx, y: 4.467, w: 0.293, h: 0.253, rotate: 180, fill: { color: C.gray } });
    });
}

/* ============================================================== slide  3 */
/* numbered bubbles above a five step arrow                                 */

function slide03(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');

  txt(slide, 'In contrast to centralizing power in the hands of one person, shared leadership divides authority within a group.',
    { x: 0.609, y: 1.260, w: 8.641, h: 0.471, fontSize: 14, bold: true });

  const steps = [
    { x: 0.609, num: '01', year: '2020', color: C.red, up: true, label: 'Shared Vision ', labelH: 0.202 },
    { x: 2.318, num: '02', year: '2021', color: C.orange, up: false, label: 'Cooperation', labelH: 0.202 },
    { x: 4.024, num: '03', year: '2022', color: C.yellow, up: true, label: 'Collective Accountability', labelH: 0.404 },
    { x: 5.729, num: '04', year: '2023', color: C.lime, up: false, label: 'Integrated Systems Approach', labelH: 0.404 },
    { x: 7.436, num: '05', year: '2024', color: C.green, up: true, label: 'Widely Distributed Leadership', labelH: 0.404 },
  ];

  steps.forEach((s, i) => {
    const barY = i === 3 ? 3.189 : 3.180;
    const stemX = s.x + 0.799;
    // connector stem + terminal dot
    box(slide, 'rect', { x: stemX, y: s.up ? 2.628 : 2.628, w: 0.031, h: s.up ? 1.684 : 1.167, fill: { color: C.rule } });
    dot(slide, stemX + 0.015, (s.up ? 4.378 : 3.792), 0.132, { fill: { color: s.color } });

    if (i === 4) {
      poly(slide, 7.436, 2.803, 1.955, 1.095, FAT_ARROW, { fill: { color: s.color } });
    } else {
      box(slide, 'roundRect', { x: s.x, y: barY, w: 1.624, h: 0.323, rectRadius: 0.024, fill: { color: s.color } });
    }

    // numbered circle
    const cx = s.x + 0.814;
    const cy = s.up ? 2.601 : 2.284;
    dot(slide, cx, cy, 0.747, { fill: { color: s.color } });
    txt(slide, s.num, { x: cx - 0.247, y: cy - 0.34, w: 0.493, h: 0.404, fontSize: 18, bold: true, color: C.white, align: 'center' });

    txt(slide, s.year, { x: s.x + 0.412, y: 3.177, w: 0.793, h: 0.333, fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle' });

    shapeText(slide, 'rect', s.label, {
      x: s.x + 0.013, y: s.up ? 4.689 : 4.080, w: 1.602, h: s.labelH, margin: 0,
      fill: { color: s.color, transparency: 90 }, color: C.gray, fontSize: 12, bold: true, valign: 'top',
    });
  });
}

/* ============================================================== slide  4 */
/* six chevrons with alternating callout cards                              */

function slide04(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');

  const top = [
    { card: 0.417, y: 1.110, badge: 0.906, title: 'Establish', body: 'an effective corporate governance structure', n: '1', color: C.red, arrow: 1.025, ay: 2.459, ah: 0.434 },
    { card: 3.380, y: 1.062, badge: 3.869, title: 'Strengthen', body: 'corporate management', n: '3', color: C.yellow, arrow: 3.989, ay: 2.411, ah: 0.481 },
    { card: 6.395, y: 1.110, badge: 6.884, title: 'Respect', body: 'the rights and interests of the shareholders', n: '5', color: C.green, arrow: 7.007, ay: 2.459, ah: 0.434 },
  ];
  top.forEach((t) => {
    slide.addShape('line', { x: t.arrow, y: t.ay, w: 0.01, h: t.ah, flipV: true, line: { color: t.color, width: 0.75, endArrowType: 'triangle' } });
    box(slide, 'roundRect', { x: t.card, y: t.y, w: 2.259, h: 1.224, rectRadius: 0.043, fill: { color: C.white }, shadow: SH_BADGE() });
    txt(slide, t.title, { x: t.card + 0.196, y: t.y + 0.128, w: 2.063, h: 0.236, fontSize: 14, bold: true });
    txt(slide, t.body, { x: t.card + 0.196, y: t.y + 0.388, w: 1.863, h: 0.337, fontSize: 10 });
    shapeText(slide, 'ellipse', t.n, { x: t.badge, y: t.y + 1.089, w: 0.26, h: 0.26, fill: { color: t.color }, color: C.white, fontSize: 11, bold: true });
  });

  // this label sits behind the middle bottom card in the original deck
  txt(slide, 'better access to external financing', { x: 4.082, y: 4.627, w: 2.063, h: 0.505, fontSize: 10, align: 'right', lineSpacingMultiple: 1.5 });
  txt(slide, 'Funding', { x: 4.082, y: 4.367, w: 2.063, h: 0.236, fontSize: 14, bold: true, align: 'right' });

  const bottom = [
    { card: 1.121, badge: 2.346, title: 'Warranty', body: ["the shareholders' ", 'rights'], n: '2', color: C.orange, arrow: 2.476 },
    { card: 4.141, badge: 5.367, title: 'Control', body: ['and monitor departments and', 'work processes of the company'], n: '4', color: C.lime, arrow: 5.497 },
    { card: 7.167, badge: 8.392, title: 'Improve', body: ['information and ', 'transparency'], n: '6', color: C.blue, arrow: 8.516 },
  ];
  bottom.forEach((b, i) => {
    slide.addShape('line', { x: b.arrow, y: 3.519, w: 0.006, h: 0.434, flipH: i < 2, line: { color: b.color, width: 0.75, endArrowType: 'triangle' } });
    box(slide, 'roundRect', { x: b.card, y: 4.089, w: 2.259, h: 1.224, rectRadius: 0.043, fill: { color: C.white }, shadow: SH_BADGE() });
    txt(slide, b.body.map((t) => ({ text: t, options: { breakLine: true } })),
      { x: b.card + 0.108, y: 4.627, w: 1.954, h: 0.505, fontSize: 10, align: 'right' });
    txt(slide, b.title, { x: b.card, y: 4.367, w: 2.063, h: 0.236, fontSize: 14, bold: true, align: 'right' });
    shapeText(slide, 'ellipse', b.n, { x: b.badge, y: 3.953, w: 0.26, h: 0.26, fill: { color: b.color }, color: C.white, fontSize: 11, bold: true });
  });

  shapeText(slide, 'homePlate', '2020', { x: 0.421, y: 2.892, w: 1.522, h: 0.627, fill: { color: C.red }, color: C.white, fontSize: 14, bold: true });
  const chevrons = [
    { x: 1.821, year: '2021', fill: C.white, line: C.orange, color: C.gray },
    { x: 3.330, year: '2022', fill: C.yellow, line: null, color: C.white },
    { x: 4.839, year: '2023', fill: C.white, line: C.lime, color: C.gray },
    { x: 6.348, year: '2024', fill: C.green, line: null, color: C.white },
    { x: 7.857, year: '2025', fill: C.white, line: C.blue, color: C.gray },
  ];
  chevrons.forEach((c) => {
    slide.addText(c.year, {
      shape: 'chevron', x: c.x, y: 2.892, w: 1.631, h: 0.627,
      fill: { color: c.fill }, line: c.line ? { color: c.line, width: 0.75 } : NOLINE,
      shadow: c.line ? SH_RING() : undefined,
      fontFace: FONT, color: c.color, fontSize: 14, bold: true, align: 'center', valign: 'middle',
    });
  });
}

/* ============================================================== slide  5 */
/* six icon rings over a year band and description cards                    */

function slide05(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');
  const BODY = 'the rights and interests of the shareholders';
  const GLYPHS = ['sun', 'smileyFace', 'donut', 'star5', 'gear6', 'pie'];

  box(slide, 'rect', { x: 0, y: 2.887, w: 10, h: 0.416, fill: { color: C.band } });

  ACCENTS.forEach((color, i) => {
    const x = 0.503 + i * 1.5076;         // group pitch
    const ringX = 0.621 + i * 1.5078;
    const cardX = 0.477 + i * 1.5076;

    dot(slide, ringX + 0.578, 1.712, 1.156, { fill: { color } });
    dot(slide, ringX + 0.578, 1.712, 0.913, { fill: { color: C.white }, shadow: SH_BADGE() });
    icon(slide, ringX + 0.578, 1.712, 0.42, C.dark, GLYPHS[i]);

    box(slide, 'triangle', { x: x + 0.569, y: 2.562, w: 0.254, h: 0.219, fill: { color } });
    txt(slide, String(2020 + i), { x, y: 2.999, w: 1.393, h: 0.202, fontSize: 12, bold: true, align: 'center' });
    box(slide, 'triangle', { x: x + 0.569, y: 3.419, w: 0.254, h: 0.219, rotate: 180, fill: { color } });

    shapeText(slide, 'roundRect', 'Insert Text here', {
      x: cardX, y: 3.825, w: 1.440, h: 0.320, rectRadius: 0.053,
      fill: { color }, color: C.white, fontSize: 11, bold: true,
    });
    box(slide, 'roundRect', { x: cardX, y: 4.262, w: 1.441, h: 0.968, rectRadius: 0.075, fill: { color: C.white }, shadow: SH_PANEL() });
    txt(slide, BODY, { x: cardX + 0.039, y: 4.391, w: 1.367, h: 0.505, fontSize: 10, align: 'center' });
  });
}

/* ============================================================== slide  6 */
/* six downward banners linked by a dashed path                             */

function slide06(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');

  const items = [
    { color: C.red, n: '1', body: 'Inspire and challenge employees', high: true },
    { color: C.orange, n: '2', body: 'Build trust and act as a role model', high: false },
    { color: C.yellow, n: '3', body: 'Encourage initiative', high: true },
    { color: C.lime, n: '4', body: 'Act entrepreneurially', high: false },
    { color: C.green, n: '5', body: 'Communicate openly and fairly', high: true },
    { color: C.blue, n: '6', body: 'Develop employee skills', high: false },
  ];
  const PITCH = 1.5625;

  items.forEach((it, i) => {
    const px = -0.406 + i * PITCH;       // pre-rotation origin of the 3.0 x 1.34 pentagon
    const left = 0.462 + i * PITCH;
    const pillY = it.high ? 1.621 : 1.287;
    const dotY = it.high ? 1.382 : 2.003;

    box(slide, 'homePlate', { x: px, y: 1.976, w: 3.0, h: 1.340, rotate: 90, fill: { color: it.color } });
    shapeText(slide, 'ellipse', it.n, {
      x: 0.810 + i * PITCH, y: 3.159, w: 0.568, h: 0.568,
      fill: { color: C.white, transparency: 59 }, shadow: SH_GLOW(), color: C.white, fontSize: 14, bold: true,
    });
    shapeText(slide, 'roundRect', 'Insert Text here', {
      x: left, y: pillY, w: 1.253, h: 0.497, rectRadius: 0.045,
      fill: { color: it.color }, shadow: SH_PILL(), color: C.white, fontSize: 10, bold: true,
    });
    shapeText(slide, 'roundRect', it.body, {
      x: left + 0.054, y: 2.247, w: 1.156, h: 0.714, rectRadius: 0.022,
      fill: { color: C.white, transparency: 70 }, color: C.white, fontSize: 8, bold: false,
    });
    dot(slide, left + 0.632, dotY, 0.119, { fill: { color: C.white }, shadow: SH_GLOW() });
  });

  // dashed staircase joining neighbouring dots
  for (let i = 0; i < 5; i++) {
    const x1 = 1.154 + i * PITCH;
    const x2 = x1 + 1.443;
    const yA = items[i].high ? 1.382 : 2.003;
    const yB = items[i + 1].high ? 1.382 : 2.003;
    const mid = (x1 + x2) / 2;
    polyline(slide, [[x1, yA], [mid, yA], [mid, yB], [x2, yB]], { color: C.rule, width: 1, dashType: 'dash' });
  }

  box(slide, 'roundRect', { x: 0.424, y: 4.424, w: 9.152, h: 0.702, rectRadius: 0.064, fill: { color: C.white }, shadow: shadow(20, 0, 0.15) });
  items.forEach((it, i) => {
    shapeText(slide, 'roundRect', String(2020 + i), {
      x: 0.462 + i * PITCH, y: 4.526, w: 1.253, h: 0.497, rectRadius: 0.045,
      fill: { color: it.color }, shadow: SH_PILL(), color: C.white, fontSize: 10, bold: true,
    });
  });
}

/* ============================================================== slide  7 */
/* month ruler with milestone flags above and description bars below        */

function slide07(pptx) {
  const slide = heading(pptx, '12 months timeline');

  const rows = [];
  for (let r = 0; r < 16; r++) {
    const row = [];
    for (let c = 0; c < 12; c++) {
      const isMonthRow = r === 8;
      row.push(gridCell(isMonthRow ? MONTHS[c] : '', Object.assign({ border: COL_BORDER },
        isMonthRow ? { fill: { color: ACCENTS[Math.floor(c / 3)] }, color: C.white, fontSize: 10.5, bold: true } : {})));
    }
    rows.push(row);
  }
  slide.addTable(rows, {
    x: 0.317, y: 1.146, w: 9.367, colW: new Array(12).fill(0.781),
    rowH: [0.232, 0.232, 0.217, 0.217, 0.217, 0.217, 0.217, 0.217,
      0.288, 0.288, 0.288, 0.288, 0.288, 0.288, 0.288, 0.301],
  });

  const bars = [
    [0.583, 3.944, 2.176, 0.202, 'Description', C.red],
    [0.583, 4.192, 3.167, 0.170, 'Description', C.red],
    [3.348, 4.409, 2.159, 0.170, 'Description', C.orange],
    [4.167, 4.626, 1.408, 0.170, 'Description', C.orange],
    [5.500, 4.842, 2.383, 0.170, 'Description', C.yellow],
    [6.909, 5.059, 1.675, 0.170, 'Description', C.lime],
  ];
  const caps = [[2.787, 3.993], [3.818, 4.210], [5.575, 4.427], [5.627, 4.643], [7.930, 4.860], [8.624, 5.077]];
  bars.forEach((b, i) => ganttBar(slide, b, caps[i]));

  // milestone flags: label box, note box, marker triangle (down = flag sits above the ruler)
  const NOTE = "The project's schedule, deadlines and important milestones.";
  const flags = [
    { lx: 0.407, ly: 2.280, lw: 1.312, nx: 0.355, ny: 2.508, nh: 0.151, note: 'Date', tx: 0.967, ty: 2.719, down: true },
    { lx: 1.984, ly: 1.724, lw: 1.312, nx: 1.932, ny: 1.952, nh: 0.454, note: NOTE, tx: 2.595, ty: 2.719, down: true },
    { lx: 3.802, ly: 2.280, lw: 1.312, nx: 3.750, ny: 2.508, nh: 0.151, note: 'Date', tx: 4.363, ty: 2.719, down: true },
    { lx: 5.715, ly: 2.280, lw: 1.312, nx: 5.663, ny: 2.508, nh: 0.151, note: 'Date', tx: 6.276, ty: 2.719, down: true },
    { lx: 7.436, ly: 1.724, lw: 1.312, nx: 7.384, ny: 1.952, nh: 0.454, note: NOTE, tx: 8.047, ty: 2.719, down: true },
    { lx: 8.586, ly: 2.280, lw: 1.312, nx: 8.534, ny: 2.508, nh: 0.151, note: 'Date', tx: 9.147, ty: 2.719, down: true },
    { lx: 0.766, ly: 3.416, lw: 1.866, nx: 0.991, ny: 3.578, nh: 0.151, note: 'Date', tx: 1.604, ty: 3.215, down: false },
    { lx: 4.240, ly: 3.501, lw: 1.866, nx: 4.465, ny: 3.747, nh: 0.151, note: 'Date', tx: 5.078, ty: 3.215, down: false },
    { lx: 5.976, ly: 3.501, lw: 1.866, nx: 6.201, ny: 3.747, nh: 0.151, note: 'Date', tx: 6.814, ty: 3.215, down: false },
    { lx: 7.912, ly: 3.501, lw: 1.866, nx: 8.138, ny: 3.782, nh: 0.454, note: NOTE, tx: 8.750, ty: 3.215, down: false },
  ];
  flags.forEach((f) => {
    txt(slide, 'Milestone', { x: f.lx, y: f.ly, w: f.lw, h: 0.168, fontSize: 10, bold: true, align: 'center' });
    txt(slide, f.note, { x: f.nx, y: f.ny, w: 1.416, h: f.nh, fontSize: 9, align: 'center' });
    box(slide, 'triangle', { x: f.tx, y: f.ty, w: 0.190, h: 0.164, rotate: f.down ? 180 : 0, fill: { color: C.gray } });
  });
}

/* ============================================================== slide  8 */
/* four phase chevrons on an arrow rail with theory cards                   */

function slide08(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');

  box(slide, 'rightArrow', { x: 0, y: 1.954, w: 10, h: 0.458, fill: { color: C.band } });

  const phases = [
    { year: '2023', label: 'Decide', color: C.red, items: ['Expectancy-value Theory', 'Expectancy, Instrumentality, Valence Theory', 'Risk-taking Model'] },
    { year: '2024', label: 'Set Goal', color: C.orange, items: ['Goal-setting Theory'] },
    { year: '2025', label: 'Act', color: C.yellow, items: ['Action Vs. State Orientation', 'Self-regulation Theory'] },
    { year: '2026', label: 'Assess', color: C.lime, items: ['Attribution Theory', 'Equity Theory'] },
  ];
  const PITCH = 2.097;

  phases.forEach((p, i) => {
    shapeText(slide, 'roundRect', p.year, {
      x: 0.838 + i * PITCH, y: 1.259, w: 1.879, h: 0.417, rectRadius: 0.069,
      fill: { color: C.gray }, color: C.white, fontSize: 12, bold: true,
    });
    shapeText(slide, 'ellipse', String(i + 1), {
      x: 1.604 + i * PITCH, y: 2.010, w: 0.347, h: 0.347,
      fill: { color: p.color }, color: C.white, fontSize: 14, bold: true,
    });
    shapeText(slide, 'chevron', p.label, {
      x: 0.665 + i * PITCH, y: 2.553, w: 2.225, h: 0.447,
      fill: { color: p.color }, color: C.white, fontSize: 12, bold: true,
    });
    box(slide, 'roundRect', {
      x: 0.826 + i * PITCH, y: 3.312, w: 1.902, h: 1.750, rectRadius: 0.078,
      fill: { color: p.color, transparency: 85 },
    });
    txt(slide, p.items.map((t) => ({ text: t, options: { bullet: { characterCode: '2022', indent: 13.5 }, breakLine: true } })), {
      x: 0.871 + i * PITCH, y: 3.478, w: 1.813, h: 1.111, fontSize: 8, lineSpacingMultiple: 1.5,
      margin: PPT_INSET,
    });
  });
}

/* ============================================================== slide  9 */
/* overlapping translucent circles chained by white rings                   */

function slide09(pptx) {
  const slide = heading(pptx, 'Enter your sub headline here');
  const BODY = 'It\u2019s now that much easier and more effective to start, Lorem';

  // painted back-to-front exactly as in the source deck
  const circles = [
    [3.776, 2.864, C.blue], [2.079, 2.864, C.green], [7.170, 2.864, C.orange],
    [3.776, 1.114, C.yellow], [0.382, 1.114, C.red], [2.079, 1.114, C.orange],
    [5.473, 2.864, C.red], [5.473, 1.114, C.lime],
  ];
  circles.forEach(([x, y, color]) => box(slide, 'ellipse', { x, y, w: 2.449, h: 2.449, fill: { color, transparency: 30 } }));

  [5.531, 7.251, 3.838, 2.151].forEach((x) => {
    box(slide, 'ellipse', { x, y: 2.910, w: 0.607, h: 0.607, fill: { color: C.white } });
    box(slide, 'ellipse', { x: x + 0.186, y: 3.096, w: 0.235, h: 0.235, fill: { color: C.white }, shadow: SH_DOT() });
  });

  // heading + body per lobe: [titleX, titleY, title, bodyX, bodyY]
  const lobes = [
    [0.831, 1.831, 'Content', 0.880, 2.181], [2.529, 2.593, 'Project', 2.577, 1.831],
    [4.226, 1.831, 'Place', 4.274, 2.181], [5.923, 2.560, 'Hosting', 5.971, 1.831],
    [2.272, 4.308, 'Development', 2.501, 3.551], [4.214, 3.551, 'Business', 4.278, 3.896],
    [5.912, 4.308, 'Plan', 5.975, 3.551], [7.609, 3.551, 'Place', 7.672, 3.896],
  ];
  lobes.forEach(([tx, ty, title, bx, by], i) => {
    txt(slide, title, {
      x: tx, y: ty, w: i === 4 ? 1.903 : 1.571, h: i === 3 ? 0.303 : 0.269,
      fontSize: 16, bold: true, color: C.white, align: 'center',
    });
    txt(slide, BODY, { x: bx, y: by, w: 1.475, h: 0.682, fontSize: 9, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  });

  const years = [[1.226, 1.426, '2020'], [2.847, 1.426, '2021'], [4.607, 1.426, '2022'], [6.306, 1.426, '2023'],
    [2.912, 4.741, '2024'], [4.609, 4.741, '2025'], [6.306, 4.741, '2026'], [8.003, 4.741, '2027']];
  years.forEach(([x, y, t]) => txt(slide, t, { x, y, w: 0.783, h: 0.269, fontSize: 16, bold: true, color: C.white, align: 'center' }));
}

/* ============================================================== slide 10 */
/* project table with a completion column and milestone callouts            */

function slide10(pptx) {
  const slide = heading(pptx, '12 months timeline');

  const blank = () => gridCell('');
  const rows = [];
  const head = [gridCell('My Project', { color: C.gray, fontSize: 9, bold: true, fontFace: FONT }),
    gridCell('% comp.', { color: C.gray, fontSize: 8, bold: true, fontFace: FONT })];
  for (let c = 0; c < 12; c++) head.push(blank());
  rows.push(head);

  const tasks = [['Planning', '80%', C.red], ['-Description Text', '80%', C.ink], ['-Description Text', '80%', C.ink],
    ['-Description Text', '', C.ink], ['', '', C.ink], ['-Description Text', '80%', C.ink]];
  tasks.forEach(([label, pct, color]) => {
    const row = [gridCell(label, { color, fontSize: 8, bold: color === C.red, fontFace: FONT }),
      gridCell(pct, { color: C.ink, fontSize: 8, fontFace: FONT })];
    for (let c = 0; c < 12; c++) row.push(blank());
    rows.push(row);
  });

  const last = [gridCell('Project 2023', { colspan: 2, fill: { color: C.gray }, color: C.white, fontSize: 10, bold: true, fontFace: FONT })];
  MONTHS.forEach((m, i) => last.push(gridCell(m, {
    fill: { color: i === 0 ? C.gray : (i < 6 ? 'C9C9C9' : '929292') },
    color: C.white, fontSize: 10.5, bold: true, fontFace: FONT,
  })));
  rows.push(last);

  slide.addTable(rows, {
    x: 0.25, y: 1.032, w: 9.5, colW: [1.267].concat(new Array(13).fill(0.633)),
    rowH: [0.222, 0.207, 0.207, 0.207, 0.207, 0.207, 0.207, 0.244],
  });

  slide.addShape('line', {
    x: 2.601, y: 1.402, w: 4.8, h: 0,
    line: { color: C.gray, width: 1, beginArrowType: 'triangle', endArrowType: 'triangle' },
  });

  const bars = [
    [2.601, 1.546, 1.164, 0.170, 'Description', C.red],
    [3.351, 1.779, 1.164, 0.170, 'Description', C.red],
    [4.102, 2.012, 1.164, 0.170, 'Description', C.red],
    [4.853, 2.246, 0.895, 0.170, 'Description', C.red],
    [5.356, 2.478, 0.895, 0.170, 'Description', C.red],
  ];
  const caps = [[3.805, 1.563], [4.555, 1.797], [5.306, 2.030], [5.809, 2.264], [6.313, 2.496]];
  bars.forEach((b, i) => ganttBar(slide, b, caps[i]));

  const NOTE = "The project's schedule, deadlines and important milestones.";
  // milestones under the ruler
  [[1.502, 3.243, 2.340], [4.451, 3.243, 5.290], [7.401, 3.243, 8.239],
    [2.976, 4.312, 3.815], [5.926, 4.312, 6.764]].forEach(([lx, ly, tx]) => {
    txt(slide, 'Milestone', { x: lx, y: ly, w: 1.866, h: 0.168, fontSize: 10, bold: true, align: 'center' });
    txt(slide, NOTE, { x: lx + 0.225, y: ly + 0.280, w: 1.416, h: 0.454, fontSize: 9, align: 'center' });
    box(slide, 'triangle', { x: tx, y: 2.957, w: 0.190, h: 0.164, fill: { color: C.gray } });
  });
  [3.910, 6.859].forEach((x) => slide.addShape('line', { x, y: 3.120, w: 0, h: 1.192, line: { color: C.gray, width: 0.75 } }));

  // milestone card floating over the right edge of the table
  box(slide, 'rect', { x: 8.083, y: 1.377, w: 1.533, h: 0.852, fill: { color: C.white } });
  txt(slide, 'Milestone', { x: 8.201, y: 1.461, w: 1.312, h: 0.168, fontSize: 10, bold: true, align: 'center' });
  txt(slide, NOTE, { x: 8.149, y: 1.689, w: 1.416, h: 0.454, fontSize: 9, align: 'center' });
  box(slide, 'triangle', { x: 8.762, y: 2.456, w: 0.190, h: 0.164, rotate: 180, fill: { color: C.gray } });
}

/* ============================================================== slide 11 */
/* seven year segments on a pill rail with hook arrows to cards             */

function slide11(pptx) {
  const slide = heading(pptx, 'Enter your sub headline here');

  const RAIL_Y = 3.028;
  const RAIL_H = 0.516;
  // rounded end caps: the round side stays visible, the square side is covered by the segments
  box(slide, 'roundRect', { x: 0.007, y: RAIL_Y, w: 1.981, h: RAIL_H, rectRadius: 0.258, fill: { color: C.band } });
  box(slide, 'roundRect', { x: 7.578, y: RAIL_Y, w: 2.415, h: RAIL_H, rectRadius: 0.258, fill: { color: C.band } });

  const segs = [
    { x: 1.688, color: C.red, year: '2022', hook: 'up', hx: 1.689 },
    { x: 2.573, color: C.orange, year: '2023', hook: 'down', hx: 2.573 },
    { x: 3.458, color: C.yellow, year: '2024', hook: 'up', hx: 3.458 },
    { x: 4.342, color: C.lime, year: '2025', hook: 'down', hx: 4.342 },
    { x: 5.226, color: C.green, year: '2026', hook: 'up', hx: 5.226 },
    { x: 6.111, color: C.blue, year: '2027', hook: 'down', hx: 6.111 },
    { x: 6.995, color: C.red, year: '2028', hook: 'up', hx: 6.995 },
  ];
  segs.forEach((s) => {
    poly(slide, s.hx, s.hook === 'up' ? 2.524 : 3.135, 0.432, 0.911,
      s.hook === 'up' ? HOOK_UP : HOOK_DOWN, { fill: { color: s.color } });
    box(slide, 'rect', { x: s.x, y: RAIL_Y, w: 0.885, h: RAIL_H, fill: { color: s.color } });
    txt(slide, s.year, { x: s.x + 0.074, y: 3.117, w: 0.736, h: 0.337, fontSize: 14, bold: true, color: C.white, align: 'center' });
  });

  const cards = [
    [1.150, 1.146], [2.923, 1.146], [4.696, 1.146], [6.469, 1.146],
    [2.027, 4.328], [3.774, 4.328], [5.521, 4.328],
  ];
  cards.forEach(([x, y]) => {
    box(slide, 'roundRect', { x, y, w: 1.483, h: 0.984, rectRadius: 0.056, fill: { color: C.band } });
    txt(slide, 'Insert Title Here', { x: x + 0.085, y: y + 0.176, w: 1.314, h: 0.185, fontSize: 11, bold: true, align: 'center' });
    txt(slide, 'It\u2019s now that much easier and more', {
      x: x + 0.153, y: y + 0.405, w: 1.178, h: 0.404, fontSize: 8, align: 'center', lineSpacingMultiple: 1.5,
    });
  });
}

/* ============================================================== slide 12 */
/* before / after bars for three customer groups                            */

function slide12(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');

  box(slide, 'rect', { x: 1.203, y: 1.312, w: 3.797, h: 0.299, fill: { color: C.rule } });
  box(slide, 'rect', { x: 5.000, y: 1.312, w: 3.797, h: 0.299, fill: { color: C.band } });
  txt(slide, '2020', { x: 1.223, y: 1.344, w: 1.252, h: 0.236, fontSize: 14, bold: true, color: C.dark, align: 'center' });
  txt(slide, '2025', { x: 7.331, y: 1.344, w: 1.252, h: 0.236, fontSize: 14, bold: true, color: C.dark, align: 'center' });
  txt(slide, 'CUSTOMER GROUPS', { x: 3.855, y: 1.816, w: 2.290, h: 0.236, fontSize: 14, bold: true, color: C.dark, align: 'center' });

  const groups = [
    { name: 'GROUP - A', y: 2.582, ruleY: 2.404, lx: 2.307, lw: 2.693, lPct: '70%', rw: 2.333, rPct: '50%', rPctX: 6.750, color: C.red },
    { name: 'GROUP - B', y: 3.443, ruleY: 3.265, lx: 3.250, lw: 1.750, lPct: '20%', rw: 1.333, rPct: '15%', rPctX: 5.802, color: C.orange },
    { name: 'GROUP - C', y: 4.304, ruleY: 4.126, lx: 2.750, lw: 2.250, lPct: '55%', rw: 3.250, rPct: '70%', rPctX: 7.709, color: C.yellow },
  ];
  groups.forEach((g) => {
    slide.addShape('line', {
      x: 0.813, y: g.ruleY, w: 8.373, h: 0,
      line: { color: C.steel, width: 0.75, beginArrowType: 'oval', endArrowType: 'oval' },
    });
    shapeText(slide, 'rect', g.name, {
      x: 4.415, y: g.ruleY - 0.123, w: 1.252, h: 0.236, margin: 0,
      fill: { color: C.white }, color: C.steel, fontSize: 14, bold: true, valign: 'top',
    });
    box(slide, 'rect', { x: g.lx, y: g.y, w: g.lw, h: 0.352, fill: { color: g.color } });
    box(slide, 'rect', { x: 5.000, y: g.y, w: g.rw, h: 0.352, fill: { color: g.color } });
    box(slide, 'rect', { x: 5.000, y: g.y, w: g.rw, h: 0.352, fill: { color: C.white, transparency: 50 } });
    txt(slide, g.lPct, { x: g.lx + 0.038, y: g.y + 0.076, w: 0.508, h: 0.202, fontSize: 12, bold: true, color: C.white, align: 'center' });
    txt(slide, g.rPct, { x: g.rPctX, y: g.y + 0.076, w: 0.508, h: 0.202, fontSize: 12, bold: true, color: C.white, align: 'center' });
  });
}

/* ============================================================== slide 13 */
/* slanted year chips with pinned notes above and below                     */

function slide13(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');
  const NOTE = 'variations of passag of lorem ipsum available,';

  const chips = [
    { x: 0.839, color: C.red, year: '2022', up: false, labelX: 1.087 },
    { x: 2.215, color: C.orange, year: '2024', up: true, labelX: 1.529 },
    { x: 3.593, color: C.yellow, year: '2026', up: false, labelX: 3.887 },
    { x: 4.969, color: C.lime, year: '2028', up: true, labelX: 4.000 },
    { x: 6.346, color: C.green, year: '2030', up: false, labelX: 6.592 },
    { x: 7.724, color: C.blue, year: '2032', up: true, labelX: 7.210 },
  ];
  chips.forEach((c, i) => {
    poly(slide, c.x, 2.955, 1.437, 0.418, SLANT, { fill: { color: c.color } });
    txt(slide, c.year, { x: 1.181 + i * 1.3756, y: 3.030, w: 0.767, h: 0.269, fontSize: 16, bold: true, color: C.white, align: 'center' });
  });

  // pins: [stemX, stemY, dotX, dotY, colour]
  const pins = [
    [0.930, 3.370, 0.839, 4.786, C.red], [3.744, 3.370, 3.651, 4.786, C.yellow], [6.437, 3.370, 6.346, 4.786, C.green],
    [3.513, 1.398, 3.423, 1.312, C.orange], [6.208, 1.398, 6.118, 1.312, C.lime], [9.022, 1.398, 8.931, 1.312, C.blue],
  ];
  pins.forEach(([sx, sy, dx, dy, color]) => {
    box(slide, 'rect', { x: sx, y: sy, w: 0.047, h: 1.560, fill: { color: C.rule } });
    box(slide, 'ellipse', { x: dx, y: dy, w: 0.229, h: 0.229, fill: { color } });
  });

  // notes below the rail (left aligned) and above it (right aligned)
  [[1.087, 1.913], [3.887, 1.530], [6.592, 1.705]].forEach(([x, tw]) => {
    txt(slide, [{ text: 'Insert Title ', options: { breakLine: true } }, { text: 'Here' }],
      { x, y: 3.777, w: tw, h: 0.404, fontSize: 12, bold: true });
    txt(slide, NOTE, { x, y: 4.217, w: 1.337, h: 0.454, fontSize: 9, lineSpacingMultiple: 1.5 });
  });
  [[1.529, 1.872, 2.065], [4.000, 2.111, 4.774], [7.210, 1.716, 7.579]].forEach(([x, tw, nx]) => {
    txt(slide, [{ text: 'Insert Title ', options: { breakLine: true } }, { text: 'Here' }],
      { x, y: 1.730, w: tw, h: 0.404, fontSize: 12, bold: true, align: 'right' });
    txt(slide, NOTE, { x: nx, y: 2.169, w: 1.337, h: 0.454, fontSize: 9, align: 'right', lineSpacingMultiple: 1.5 });
  });
}

/* ============================================================== slide 14 */
/* dashed staircase from "A" to "Z" with vertical note columns              */

function slide14(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');
  const NOTE = 'There are many variations of passages of lorem ipsum available, but the majority have suffered alteration in';
  const DASH = { color: C.silver, width: 1.88, dashType: 'dash' };
  const TOP = 1.862;
  const BOTTOM = 4.480;

  // staircase: alternating verticals joined at the top / bottom rails
  const stops = [1.000, 2.329, 3.661, 4.995, 6.328, 7.660];
  stops.forEach((x) => polyline(slide, [[x, TOP], [x, BOTTOM]], DASH));
  [[1.000, 2.329, BOTTOM], [2.329, 3.661, TOP], [3.661, 4.995, BOTTOM],
    [4.995, 6.328, TOP], [6.328, 7.660, BOTTOM], [7.660, 8.984, TOP]]
    .forEach(([x1, x2, y]) => polyline(slide, [[x1, y], [x2, y]], DASH));

  // A / Z end caps
  [[0.539, 'A', 0.837, 0.321], [8.529, 'Z', 8.853, 0.281]].forEach(([x, letter, tx, tw]) => {
    box(slide, 'ellipse', { x, y: 1.396, w: 0.932, h: 0.932, fill: { color: C.white }, shadow: SH_RING() });
    box(slide, 'ellipse', { x: x + 0.064, y: 1.465, w: 0.793, h: 0.793, fill: { color: C.rule } });
    txt(slide, letter, { x: tx, y: 1.566, w: tw, h: 0.606, fontSize: 36, bold: true, color: C.white, align: 'center' });
  });

  const nodes = [
    { x: 2.046, y: 4.197, color: C.red, label: '1 Year', lx: 1.232, ly: 4.618, glyph: 'mathMultiply' },
    { x: 3.380, y: 1.579, color: C.orange, label: '2 Year', lx: 2.607, ly: 1.534, glyph: 'heart' },
    { x: 4.713, y: 4.197, color: C.yellow, label: '3 Year', lx: 3.895, ly: 4.618, glyph: 'sun' },
    { x: 6.045, y: 1.579, color: C.lime, label: '4 Year', lx: 5.270, ly: 1.534, glyph: 'flowChartManualInput' },
    { x: 7.377, y: 4.197, color: C.green, label: '5 Year', lx: 6.602, ly: 4.618, glyph: 'star5' },
  ];
  nodes.forEach((n, i) => {
    box(slide, 'ellipse', { x: n.x, y: n.y, w: 0.566, h: 0.566, fill: { color: C.white }, shadow: SH_DISC() });
    box(slide, 'ellipse', { x: n.x + 0.062, y: n.y + 0.063, w: 0.442, h: 0.442, fill: { color: n.color } });
    icon(slide, n.x + 0.283, n.y + 0.283, 0.21, C.white, n.glyph);
    txt(slide, n.label, {
      x: n.lx - 0.15, y: n.ly, w: 0.856, h: 0.236, fontSize: 14, bold: true, align: 'center',
      color: i === 3 ? C.yellow : n.color,
    });
  });

  // vertical note columns (boxes rotated a quarter turn)
  [[0.520, 'right'], [1.865, 'left'], [3.209, 'right'], [4.554, 'left'], [5.899, 'right']]
    .forEach(([x, align]) => {
      txt(slide, NOTE, { x, y: 2.916, w: 2.234, h: 0.454, fontSize: 9, align, rotate: 90 });
    });
}

/* ============================================================== slide 15 */
/* two phases split by a cross rule, six icon rings                         */

function slide15(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');
  const NOTE = 'There are many variations of passages of m available,';
  const GLYPHS = ['sun', 'lightningBolt', 'donut', 'heart', 'star5', 'gear6'];

  slide.addShape('line', { x: 0, y: 3.528, w: 10, h: 0, line: { color: C.silver, width: 2.25 } });
  slide.addShape('line', { x: 4.988, y: 1.562, w: 0, h: 1.965, line: { color: C.silver, width: 2.25 } });

  ACCENTS.forEach((color, i) => {
    const x = 0.474 + i * 1.5914;
    box(slide, 'ellipse', { x, y: 2.228, w: 0.915, h: 0.915, fill: { color: C.white }, shadow: SH_DISC() });
    box(slide, 'ellipse', { x: x + 0.099, y: 2.327, w: 0.718, h: 0.715, fill: { color } });
    icon(slide, x + 0.457, 2.685, 0.33, C.white, GLYPHS[i]);
    box(slide, 'ellipse', { x: x + 0.361, y: 3.431, w: 0.194, h: 0.193, fill: { color } });
    txt(slide, String(2020 + i), { x: x + 0.167, y: 3.784, w: 0.582, h: 0.303, fontSize: 18, bold: true, color, align: 'center' });
    txt(slide, NOTE, { x: x - 0.181, y: 1.581, w: 1.349, h: 0.454, fontSize: 9, align: 'center' });
  });

  [[1.480, 'PHASES 01 '], [6.254, 'PHASES 02 ']].forEach(([x, label]) => {
    box(slide, 'roundRect', { x, y: 4.375, w: 2.086, h: 0.505, rectRadius: 0.084, fill: { color: C.white }, shadow: SH_DISC() });
    txt(slide, label, { x: x + 0.180, y: 4.408, w: 1.726, h: 0.438, margin: PPT_INSET, fontSize: 20, bold: true, color: C.ink, align: 'center' });
  });
}

/* ============================================================== slide 16 */
/* rising columns on a rail, each with a date and a note                    */

function slide16(pptx) {
  const slide = heading(pptx, 'Type the subtitle of your great here');
  const NOTE = 'There are many variations of passages of lorem ipsum available, but the';

  //        x      y      w      h      date      colour   noteX  noteY  titleX titleY titleW
  const cols = [
    { x: 1.247, y: 3.569, w: 1.606, h: 0.993, date: '01,Jan', color: C.red, dx: 1.609, dy: 3.956, dw: 0.600, nx: 0.914, ny: 1.656, tx: 0.852, ty: 1.418, tw: 1.509, title: 'Insert Text here' },
    { x: 2.330, y: 3.849, w: 1.946, h: 0.714, date: '15,Feb', color: C.orange, dx: 3.028, dy: 4.071, dw: 0.596, nx: 2.281, ny: 2.425, tx: 2.416, ty: 2.187, tw: 1.114, title: 'Insert Text ' },
    { x: 3.582, y: 3.090, w: 1.948, h: 1.472, date: '10,Mar', color: C.yellow, dx: 3.998, dy: 3.345, dw: 0.619, nx: 3.648, ny: 1.550, tx: 3.797, ty: 1.312, tw: 1.086, title: 'Insert Text ' },
    { x: 4.837, y: 3.489, w: 1.946, h: 1.073, date: '05,Apr', color: C.lime, dx: 5.503, dy: 3.878, dw: 0.574, nx: 5.015, ny: 1.981, tx: 5.165, ty: 1.743, tw: 1.086, title: 'Insert Text ' },
    { x: 6.396, y: 4.026, w: 1.335, h: 0.536, date: '28,Dec', color: C.green, dx: 6.752, dy: 4.142, dw: 0.610, nx: 6.382, ny: 2.514, tx: 6.532, ty: 2.276, tw: 1.086, title: 'Insert Text ' },
    { x: 7.649, y: 3.729, w: 1.335, h: 0.833, date: '22,Oct', color: C.blue, dx: 8.019, dy: 3.987, dw: 0.580, nx: 7.749, ny: 2.205, tx: 7.899, ty: 1.967, tw: 1.086, title: 'Insert Text ' },
  ];
  // painted in the source deck's z-order so the right column overlaps its left neighbour
  [0, 2, 3, 5, 1, 4].forEach((i) => {
    const c = cols[i];
    poly(slide, c.x, c.y, c.w, c.h, columnPts(c.w, c.h, 0.06), { fill: { color: c.color } });
  });

  // thin shadow slivers where two columns meet
  [[2.299, 3.911, 0.031, 0.602], [4.276, 3.911, 0.026, 0.602], [4.804, 3.552, 0.033, 0.962],
    [6.373, 4.088, 0.023, 0.425], [7.731, 4.088, 0.017, 0.425]].forEach(([x, y, w, h]) => {
    poly(slide, x, y, w, h, [[1, 0], [1, 1], [0, 1]], { fill: { color: C.ink, transparency: 50 } });
  });

  // base rail with round start cap and arrow head
  box(slide, 'rect', { x: 0.528, y: 4.513, w: 8.837, h: 0.099, fill: { color: C.silver } });
  box(slide, 'ellipse', { x: 0.425, y: 4.404, w: 0.319, h: 0.318, fill: { color: '4D4B4D' } });
  box(slide, 'triangle', { x: 9.253, y: 4.383, w: 0.323, h: 0.359, rotate: 90, fill: { color: '4D4B4D' } });
  ACCENTS.forEach((color, i) => {
    const x = 1.458 + i * 1.3745;
    box(slide, 'roundRect', { x, y: 4.546, w: 0.219, h: 0.219, rectRadius: 0.04, fill: { color: C.white } });
    box(slide, 'roundRect', { x: x + 0.010, y: 4.557, w: 0.200, h: 0.200, rectRadius: 0.036, fill: { color } });
  });

  // grey pointer arrows from the rail up to the notes
  [[1.495, 3.069, 0.564], [2.870, 3.237, 0.682], [4.243, 2.505, 0.684],
    [5.618, 2.902, 0.684], [6.993, 3.416, 0.686], [8.368, 3.123, 0.686]]
    .forEach(([x, y, h]) => box(slide, 'upArrow', { x, y, w: 0.144, h, fill: { color: C.silver } }));

  cols.forEach((c) => {
    // widened a touch so the date never wraps at 14pt
    txt(slide, c.date, { x: c.dx - 0.09, y: c.dy, w: c.dw + 0.18, h: 0.236, fontSize: 14, bold: true, color: C.white, align: 'center' });
    txt(slide, c.title, { x: c.tx, y: c.ty, w: c.tw, h: 0.202, fontSize: 12, bold: true, color: c.color, align: 'center' });
    txt(slide, NOTE, { x: c.nx, y: c.ny, w: 1.385, h: 0.606, fontSize: 9, align: 'center' });
  });
}

/* -------------------------------------------------------------- assembly */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
  pptx.layout = 'DECK';
  pptx.title = 'Process Timeline';
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
  BUILDERS.forEach((fn) => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '1061618d-6239-40ca-abca-78521eac7502_grok_final.pptx') });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => { console.error(e); process.exit(1); });
