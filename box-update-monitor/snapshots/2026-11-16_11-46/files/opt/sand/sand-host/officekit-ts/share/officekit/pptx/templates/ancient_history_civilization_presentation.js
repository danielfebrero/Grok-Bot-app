/*
 * Ancient History – Presentation Template
 * Standalone pptxgenjs re-creation of the 20-slide reference deck.
 *
 *   node 04d3ee0d-e648-46d4-b68b-1efa5c844a42_grok_final.js
 *
 * Raster photography in the original is replaced by native-shape placeholders.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ── Theme ─────────────────────────────────────────────────────────────── */

const BROWN = '573004'; // accent1 – primary dark
const GOLD = 'D5AA6D'; // accent2 – primary light
const CREAM = 'FCF7EE'; // page background (lt2 @ 65%)
const INK = '0D0D0D'; // body copy on light panels
const WHITE = 'FFFFFF';
const BLACK = '000000';
const TRACK = 'D9D9D9'; // progress-bar trough

const NO_LINE = { type: 'none' }; // pptxgenjs draws a default outline unless told otherwise

const SERIF = 'Lora'; // major font
const SANS = 'Manrope'; // minor font

const NAV_LINKS = [
  { label: 'Home', x: 5.331, w: 0.656 },
  { label: 'About Us', x: 6.033, w: 0.872 },
  { label: 'Information', x: 6.951, w: 1.052 },
];

const LOREM = {
  long: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
  hero: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor enean massa. ',
  short: 'Lorem ipsum dolor sit amet, consectetuer adipiscing.',
  tight: 'Lorem ipsum dolor sit consectetuer adipiscing aenean commodo.',
  quote: '\u201CLorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo.\u201D',
};

/* ── Primitive helpers ─────────────────────────────────────────────────── */

// Plain text block: top-anchored, auto-height, like every text box in the deck.
function text(slide, x, y, w, h, runs, o = {}) {
  slide.addText(typeof runs === 'string' ? [{ text: runs }] : runs, {
    x, y, w, h,
    fontFace: o.face || SERIF,
    fontSize: o.size || 18,
    color: o.color || INK,
    italic: !!o.italic,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.lh,
    autoFit: true,
  });
}

// Body copy is always 12pt Manrope on 1.3 line spacing.
function body(slide, x, y, w, h, str, color) {
  text(slide, x, y, w, h, str, { face: SANS, size: 12, color, lh: 1.3 });
}

function box(slide, x, y, w, h, color, o = {}) {
  slide.addShape('rect', { x, y, w, h, fill: { color, transparency: o.transparency }, line: NO_LINE });
}

function hline(slide, x, y, w, color) {
  slide.addShape('line', { x, y, w, h: 0, line: { color, width: 0.5 } });
}

function vline(slide, x, y, h, color) {
  slide.addShape('line', { x, y, w: 0, h, line: { color, width: 0.5 } });
}

// Pill button: rounded rectangle + centred caption.
function button(slide, x, y, label, fill, color, wide) {
  const w = 1.476, h = 0.405;
  slide.addShape('roundRect', { x, y, w, h, rectRadius: h / 2, fill: { color: fill }, line: NO_LINE });
  const pad = wide ? 0 : 0.056;
  text(slide, x + pad, y + 0.034, w - pad * 2, 0.337, label, { size: 14, color, align: 'center' });
}

// Small italic kicker preceded by a rule — the deck's section label.
function eyebrow(slide, x, y, color, label = 'History Of Civilization') {
  hline(slide, x, y, 0.89, color);
  text(slide, x + 0.898, y - 0.171, 2.25, 0.34, label, { size: 14, color, italic: true });
}

// Wide kicker used on the opening / closing slides.
function eyebrowWide(slide, x, y, color) {
  hline(slide, x, y, 2.683, color);
  text(slide, x + 2.688, y - 0.202, 3.0, 0.404, 'History Of Civilization', { color, italic: true });
}

/* ── Icon & ornament helpers ───────────────────────────────────────────── */

// Classical-bank glyph used as the nav logo (replaces the original SVG icon).
function bankIcon(slide, x, y, s, color) {
  const fill = { color }, line = { none: true };
  slide.addShape('triangle', { x: x + s * 0.04, y, w: s * 0.92, h: s * 0.26, fill, line });
  slide.addShape('rect', { x, y: y + s * 0.26, w: s, h: s * 0.08, fill, line });
  slide.addShape('rect', { x, y: y + s * 0.86, w: s, h: s * 0.1, fill, line });
  for (let i = 0; i < 4; i++) {
    slide.addShape('rect', { x: x + s * (0.1 + i * 0.24), y: y + s * 0.38, w: s * 0.1, h: s * 0.46, fill, line });
  }
}

// Three-bar "hamburger" mark at the far right of the nav bar.
function hamburger(slide, x, y, w, h, color) {
  const fill = { color }, line = { none: true }, bar = h * 0.16;
  slide.addShape('roundRect', { x: x + w * 0.28, y, w: w * 0.72, h: bar, rectRadius: bar / 2, fill, line });
  slide.addShape('roundRect', { x: x + w * 0.28, y: y + h * 0.42, w: w * 0.72, h: bar, rectRadius: bar / 2, fill, line });
  slide.addShape('roundRect', { x, y: y + h * 0.84, w: w * 0.19, h: bar, rectRadius: bar / 2, fill, line });
  slide.addShape('roundRect', { x, y, w: w * 0.19, h: bar, rectRadius: bar / 2, fill, line });
  slide.addShape('roundRect', { x, y: y + h * 0.42, w: w * 0.19, h: bar, rectRadius: bar / 2, fill, line });
  slide.addShape('roundRect', { x: x + w * 0.28, y: y + h * 0.84, w: w * 0.72, h: bar, rectRadius: bar / 2, fill, line });
}

// Filled circle with a white tick — the deck's list bullet.
function checkBullet(slide, x, y, color, s = 0.157) {
  slide.addShape('ellipse', { x, y, w: s, h: s, fill: { color }, line: NO_LINE });
  slide.addShape('custGeom', {
    x, y, w: s, h: s, fill: { color: WHITE }, line: NO_LINE,
    points: [
      { x: s * 0.24, y: s * 0.52 }, { x: s * 0.34, y: s * 0.41 }, { x: s * 0.44, y: s * 0.54 },
      { x: s * 0.68, y: s * 0.28 }, { x: s * 0.78, y: s * 0.39 }, { x: s * 0.44, y: s * 0.74 },
      { close: true },
    ],
  });
}

// Fan of hairline curves used as a watermark along the slide edges.
// Each line arcs up to a crest at ~45% of the box, then dives past the bottom.
function swoosh(slide, x, y, w, h, color, alpha, flipH) {
  const N = 20;
  for (let i = 0; i < N; i++) {
    const t = i / (N - 1);
    const y0 = h * (0.78 + 0.52 * t);
    const crest = h * (0.04 + 0.46 * t);
    const y1 = h * (0.96 + 0.42 * t);
    slide.addShape('custGeom', {
      x, y, w, h, flipH: !!flipH,
      line: { color, width: 0.75, transparency: 100 - alpha },
      points: [
        { x: 0, y: y0 },
        { curve: { type: 'cubic', x1: w * 0.22, y1: crest + h * 0.12, x2: w * 0.40, y2: crest }, x: w * 0.62, y: crest + h * 0.16 },
        { curve: { type: 'cubic', x1: w * 0.78, y1: crest + h * 0.42, x2: w * 0.90, y2: y1 - h * 0.18 }, x: w, y: y1 },
      ],
    });
  }
}

// Translucent circle + triangle: the "play" affordance on the dark slides.
function playButton(slide, x, y, d) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: BROWN, transparency: 45 }, line: NO_LINE });
  slide.addShape('triangle', {
    x: x + d * 0.385, y: y + d * 0.375, w: d * 0.29, h: d * 0.25,
    rotate: 90, fill: { color: WHITE }, line: NO_LINE,
  });
}

// Photo stand-in: flat block plus caption, never an embedded bitmap.
function imagePlaceholder(slide, x, y, w, h, o = {}) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: o.radius || 0.08, rotate: o.rotate,
    fill: { color: o.fill || 'BFBFBF' }, line: { color: o.line || '8C8C8C', width: 1 },
  });
  if (o.inner) {
    slide.addShape('roundRect', {
      x: x + w * 0.05, y: y + h * 0.045, w: w * 0.9, h: h * 0.91,
      rectRadius: 0.04, rotate: o.rotate, fill: { color: o.inner }, line: NO_LINE,
    });
  }
  text(slide, x, y + h / 2 - 0.2, w, 0.4, '[image]', { size: 12, face: SANS, color: '595959', align: 'center' });
}

/* ── Composite blocks ──────────────────────────────────────────────────── */

// "01." bullet + heading + description, used in the agenda / feature lists.
function numberedItem(slide, cx, cy, item, o) {
  const d = 0.563;
  slide.addShape('ellipse', { x: cx, y: cy, w: d, h: d, fill: { color: GOLD }, line: NO_LINE });
  text(slide, cx, cy + 0.113, d, 0.34, item.no, { size: 14, color: INK, align: 'center' });
  text(slide, o.titleX, cy - 0.072, o.titleW, 0.707, item.title, { color: o.color });
  body(slide, o.descX, cy + (o.descDY === undefined ? -0.019 : o.descDY), o.descW, o.descH || 0.34, item.desc, o.color);
}

// Gold card: heading, trough, filled bar and a percentage read-out.
function progressCard(slide, x, y, c) {
  const w = c.w || 4.06;
  box(slide, x, y, w, c.h || 1.687, GOLD);
  text(slide, x + 0.262, y + 0.262, c.titleW || 3.08, 0.707, c.title, { color: BLACK });
  const barY = y + 1.19, barW = c.barW || 2.89, barH = 0.126;
  slide.addShape('roundRect', { x: x + 0.262, y: barY, w: barW, h: barH, rectRadius: barH / 2, fill: { color: TRACK }, line: NO_LINE });
  slide.addShape('roundRect', { x: x + 0.262, y: barY, w: c.fill, h: barH, rectRadius: barH / 2, fill: { color: BROWN }, line: NO_LINE });
  text(slide, x + 0.262 + barW + 0.012, barY - 0.1, 0.63, 0.34, c.pct, { size: 14, color: BLACK });
}

// Gold card with a brown disc, open gold ring, percentage and caption.
function donutCard(slide, x, y, c) {
  box(slide, x, y, 2.331, 2.769, GOLD);
  slide.addShape('ellipse', { x: x + 0.475, y: y + 0.276, w: 1.381, h: 1.381, fill: { color: BROWN }, line: NO_LINE });
  slide.addShape('arc', {
    x: x + 0.61, y: y + 0.413, w: 1.108, h: 1.108,
    angleRange: [272.9, 157.2], line: { color: GOLD, width: 3.75 },
  });
  text(slide, x + 0.788, y + 0.766, 0.75, 0.42, c.pct, { color: GOLD, align: 'center' });
  const cw = c.captionW || 1.737;
  text(slide, x + 1.166 - cw / 2, y + 1.902, cw, 0.64, c.caption, { size: 16, color: BLACK, align: 'center' });
}

/* ── Chrome shared by every slide ──────────────────────────────────────── */

function chrome(slide, num, dark) {
  const c = dark ? WHITE : INK;
  hline(slide, 0, 0.601, 13.333, c);
  bankIcon(slide, 0.326, 0.173, 0.214, dark ? GOLD : BROWN);
  text(slide, 0.54, 0.155, 1.375, 0.286, 'Ancient History ', { size: 11, color: c });
  NAV_LINKS.forEach(l => text(slide, l.x, 0.153, l.w, 0.286, l.label, { size: 11, color: c, align: 'center' }));
  hamburger(slide, 12.766, 0.227, 0.241, 0.143, c);
  text(slide, 12.76, 7.06, 0.45, 0.29, String(num), { size: 11, color: c, align: 'center' });
}

/* ── World map (slide 18) ──────────────────────────────────────────────── */

// Coarse land outlines, in inches relative to the slide, drawn at 50% white.
const CONTINENTS = [
  { x: 9.02, y: 3.94, w: 3.87, h: 2.72, pts: [[3.87, 0.57], [3.64, 0.45], [3.35, 0.46], [3.11, 0.36], [3.03, 0.39], [2.97, 0.31], [2.93, 0.4], [2.77, 0.41], [2.72, 0.33], [2.43, 0.32], [2.43, 0.22], [2.28, 0.19], [2.22, 0.08], [2.17, 0.11], [2.11, 0], [2.04, 0.06], [2.22, 0.14], [2.2, 0.21], [1.96, 0.26], [1.83, 0.36], [1.7, 0.39], [1.63, 0.32], [1.59, 0.48], [1.46, 0.45], [1.2, 0.55], [1.13, 0.49], [1.14, 0.56], [1.11, 0.56], [0.89, 0.41], [0.74, 0.41], [0.62, 0.46], [0.54, 0.61], [0.42, 0.69], [0.44, 0.78], [0.54, 0.77], [0.56, 0.86], [0.49, 0.82], [0.49, 0.93], [0.23, 1.05], [0.29, 1.09], [0.29, 1.18], [0.15, 1.17], [0.16, 1.31], [0.21, 1.35], [0, 1.63], [0.01, 1.8], [0.15, 1.95], [0.41, 1.92], [0.52, 1.97], [0.5, 2.07], [0.59, 2.25], [0.55, 2.4], [0.67, 2.72], [0.84, 2.69], [0.97, 2.52], [0.96, 2.42], [1.07, 2.33], [1.04, 2.12], [1.27, 1.87], [1.26, 1.82], [1.13, 1.85], [1.12, 1.81], [1.35, 1.72], [1.43, 1.62], [1.38, 1.55], [1.57, 1.55], [1.62, 1.62], [1.69, 1.64], [1.75, 1.87], [1.82, 1.87], [1.82, 1.94], [1.86, 1.91], [1.81, 1.86], [1.83, 1.75], [2.02, 1.62], [2.08, 1.74], [2.13, 1.73], [2.13, 1.89], [2.2, 1.98], [2.17, 2.01], [2.09, 1.94], [2.14, 2.07], [2.32, 2.19], [2.44, 2.2], [2.28, 2.15], [2.33, 2.1], [2.22, 2.04], [2.24, 1.96], [2.16, 1.81], [2.22, 1.81], [2.27, 1.89], [2.34, 1.84], [2.28, 1.66], [2.36, 1.63], [2.35, 1.7], [2.37, 1.64], [2.54, 1.55], [2.59, 1.46], [2.54, 1.37], [2.6, 1.31], [2.58, 1.27], [2.64, 1.25], [2.68, 1.36], [2.74, 1.35], [2.7, 1.25], [2.88, 1.12], [2.95, 0.97], [2.99, 1.1], [2.99, 0.91], [2.88, 0.92], [2.85, 0.87], [2.99, 0.77], [3.18, 0.77], [3.28, 0.69], [3.38, 0.69], [3.23, 0.85], [3.25, 0.99], [3.35, 0.89], [3.38, 0.76], [3.68, 0.68], [3.7, 0.58], [3.83, 0.63]] },
  { x: 6.18, y: 3.82, w: 2.93, h: 3.32, pts: [[2.93, 0.09], [2.76, 0.1], [2.76, 0.05], [2.58, 0], [1.94, 0.12], [1.97, 0.04], [1.68, 0.03], [1.35, 0.15], [1.46, 0.28], [1.52, 0.24], [1.44, 0.31], [1.32, 0.22], [1.43, 0.37], [1.2, 0.29], [1.2, 0.35], [0.96, 0.27], [0.86, 0.33], [0.98, 0.3], [1.03, 0.39], [1.38, 0.38], [1.36, 0.46], [1.31, 0.4], [1.23, 0.44], [1.37, 0.47], [1.3, 0.62], [1.2, 0.6], [1.26, 0.56], [1.2, 0.42], [1.07, 0.46], [0.88, 0.38], [0.8, 0.47], [0.92, 0.48], [0.97, 0.59], [0.21, 0.5], [0.03, 0.6], [0.08, 0.66], [0, 0.7], [0.13, 0.75], [0.02, 0.86], [0.19, 0.9], [0.13, 0.97], [0.37, 0.83], [0.56, 0.88], [0.82, 1.16], [0.83, 1.38], [1.09, 1.72], [1.07, 1.62], [1.22, 1.81], [1.69, 1.99], [1.67, 2.12], [1.74, 2.18], [1.63, 2.27], [1.85, 2.54], [1.74, 3.11], [1.83, 3.32], [1.92, 3.3], [1.85, 3.2], [1.93, 2.97], [2.08, 2.91], [2.26, 2.64], [2.39, 2.58], [2.51, 2.28], [2.23, 2.16], [2.02, 1.96], [1.81, 1.92], [1.62, 1.98], [1.58, 1.85], [1.5, 1.85], [1.53, 1.74], [1.4, 1.81], [1.32, 1.74], [1.34, 1.6], [1.59, 1.57], [1.64, 1.67], [1.62, 1.54], [1.86, 1.25], [2.02, 1.22], [1.93, 1.2], [1.95, 1.13], [2.07, 1.09], [2.05, 1.18], [2.16, 1.19], [1.94, 0.84], [1.88, 0.91], [1.85, 0.84], [1.7, 0.79], [1.66, 1.08], [1.37, 0.89], [1.47, 0.74], [1.61, 0.74], [1.61, 0.79], [1.57, 0.67], [1.63, 0.55], [1.7, 0.54], [1.72, 0.63], [1.73, 0.58], [1.8, 0.62], [1.69, 0.72], [1.9, 0.8], [1.95, 0.75], [1.89, 0.68], [2, 0.66], [1.72, 0.43], [1.52, 0.38], [1.65, 0.38], [1.66, 0.27], [1.91, 0.12], [1.93, 0.2], [1.81, 0.22], [1.82, 0.29], [2.11, 0.39], [2.17, 0.5], [2.13, 0.56], [2.21, 0.57], [2.16, 0.66], [2.34, 0.86], [2.42, 0.68], [2.76, 0.54], [2.82, 0.28], [2.77, 0.22]] },
  { x: 11.44, y: 6.18, w: 0.75, h: 0.58, pts: [[0.75, 0.31], [0.55, 0], [0.41, 0.33], [0.3, 0.15], [0.43, 0.03], [0.34, 0.02], [0.31, 0.09], [0.22, 0.08], [0, 0.26], [0.05, 0.49], [0.38, 0.42], [0.57, 0.58], [0.7, 0.54]] },
  { x: 11.78, y: 6, w: 0.36, h: 0.2, pts: [[0, 0], [0.04, 0.06], [0.13, 0.09], [0.12, 0.15], [0.17, 0.14], [0.21, 0.16], [0.26, 0.14], [0.36, 0.2], [0.26, 0.07], [0.13, 0.02], [0.06, 0.04], [0.05, 0]] },
  { x: 9.15, y: 4.72, w: 0.23, h: 0.21, pts: [[0.07, 0.01], [0.08, 0.08], [0.01, 0.12], [0, 0.17], [0.02, 0.19], [0.06, 0.17], [0.07, 0.12], [0.12, 0.11], [0.12, 0.14], [0.09, 0.19], [0.11, 0.21], [0.2, 0.21], [0.23, 0.16], [0.2, 0.16], [0.13, 0.08], [0.15, 0.03], [0.12, 0]] },
  { x: 9.54, y: 3.98, w: 0.31, h: 0.16, pts: [[0, 0.07], [0.06, 0.08], [0.06, 0.12], [0.11, 0.16], [0.16, 0.08], [0.19, 0.08], [0.22, 0.13], [0.25, 0.11], [0.15, 0.06], [0.15, 0.02], [0.26, 0.05], [0.31, 0.01], [0.14, 0], [0.14, 0.02], [0.02, 0.02]] },
  { x: 11.36, y: 5.86, w: 0.2, h: 0.19, pts: [[0.19, 0.01], [0.15, 0], [0.12, 0.04], [0.1, 0.04], [0.07, 0.07], [0.05, 0.07], [0.04, 0.1], [0.02, 0.1], [0.02, 0.08], [0, 0.12], [0.03, 0.17], [0.1, 0.18], [0.12, 0.19], [0.16, 0.14], [0.15, 0.12], [0.19, 0.1], [0.16, 0.08], [0.16, 0.05], [0.2, 0.02]] },
  { x: 10.14, y: 6.2, w: 0.13, h: 0.26, pts: [[0.11, 0], [0.1, 0.03], [0.08, 0.03], [0.08, 0.05], [0.06, 0.07], [0.02, 0.08], [0.01, 0.1], [0.01, 0.17], [0, 0.22], [0.01, 0.26], [0.05, 0.26], [0.08, 0.24], [0.12, 0.12], [0.11, 0.08], [0.13, 0.06]] },
  { x: 11.75, y: 5.06, w: 0.31, h: 0.31, pts: [[0.31, 0.04], [0.28, 0.04], [0.24, 0], [0.19, 0.17], [0.14, 0.18], [0.09, 0.23], [0.06, 0.22], [0, 0.28], [0.02, 0.31], [0.15, 0.24], [0.21, 0.23], [0.23, 0.12], [0.22, 0.08]] },
  { x: 10.3, y: 4.12, w: 0.34, h: 0.24, pts: [[0.34, 0.01], [0.31, 0], [0.09, 0.06], [0.09, 0.1], [0.03, 0.18], [0, 0.18], [0, 0.2], [0.11, 0.24], [0.08, 0.18], [0.14, 0.09]] },
  { x: 12.46, y: 6.68, w: 0.22, h: 0.26, pts: [[0.22, 0.06], [0.18, 0.07], [0.17, 0.04], [0.12, 0], [0.15, 0.05], [0.13, 0.12], [0.1, 0.12], [0.08, 0.17], [0.02, 0.21], [0, 0.24], [0.02, 0.26], [0.05, 0.25], [0.07, 0.21], [0.12, 0.18], [0.13, 0.14], [0.18, 0.14], [0.19, 0.09]] },
  { x: 8.89, y: 4.49, w: 0.19, h: 0.09, pts: [[0, 0.02], [0, 0.04], [0.01, 0.02], [0.04, 0.03], [0.04, 0.07], [0.09, 0.09], [0.16, 0.07], [0.19, 0.04], [0.19, 0.02], [0.15, 0], [0.14, 0], [0.13, 0.01], [0.04, 0.01], [0.04, 0], [0.02, 0], [0.01, 0.01]] },
  { x: 11.57, y: 5.63, w: 0.12, h: 0.25, pts: [[0, 0], [0, 0.07], [0.07, 0.11], [0.06, 0.13], [0.01, 0.12], [0.05, 0.16], [0.04, 0.22], [0.07, 0.22], [0.09, 0.25], [0.12, 0.2], [0.09, 0.16], [0.1, 0.12], [0.06, 0.11], [0.05, 0.08], [0.03, 0.08], [0.04, 0.03], [0.04, 0]] },
];

const MAP_PINS = [
  [11.48, 4.38], [10.55, 4.60], [9.62, 5.57], [8.16, 6.01], [7.06, 4.49],
  [7.47, 4.87], [8.48, 3.93], [11.08, 5.07], [11.75, 6.25],
];

function worldMap(slide) {
  CONTINENTS.forEach(c => slide.addShape('custGeom', {
    x: c.x, y: c.y, w: c.w, h: c.h,
    fill: { color: WHITE, transparency: 50 }, line: NO_LINE,
    points: c.pts.map(p => ({ x: p[0], y: p[1] })).concat([{ close: true }]),
  }));
  MAP_PINS.forEach(([x, y]) => {
    // teardrop `adj` is reached through rectRadius: adj = rectRadius / size * 100000
    slide.addShape('teardrop', {
      x, y, w: 0.203, h: 0.203, rotate: 138.2, rectRadius: 0.277,
      fill: { color: WHITE }, line: NO_LINE,
    });
    slide.addShape('ellipse', { x: x + 0.045, y: y + 0.045, w: 0.114, h: 0.114, fill: { color: GOLD }, line: NO_LINE });
  });
}

/* ── Slides ────────────────────────────────────────────────────────────── */

const build = [];

// 1 — Title
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  box(s, 8.559, 3.646, 4.222, 1.663, GOLD);
  text(s, 0.76, 1.396, 8.613, 2.524, [
    { text: 'Ancient History ', options: { color: WHITE } },
    { text: 'Presentation', options: { color: GOLD } },
  ], { size: 80, lh: 0.9 });
  eyebrowWide(s, 0.979, 1.194, WHITE);
  text(s, 0.76, 5.129, 3.877, 0.404, 'Presentation Template', { color: WHITE });
  body(s, 0.76, 5.625, 4.729, 0.605, LOREM.hero, WHITE);
  button(s, 0.872, 6.495, 'Learn More', GOLD, INK);
  button(s, 2.563, 6.495, 'Get Started', GOLD, INK);
  chrome(s, 1, true);
});

// 2 — Statement banner
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 0.76, 1.097, 11.81, 4.524, BROWN);
  swoosh(s, 0.78, 3.91, 6.42, 1.67, GOLD, 10);
  text(s, 1.516, 2.475, 10.3, 2.1, [
    { text: 'From Early Farming To Mighty ', options: { color: GOLD } },
    { text: 'Empires, The Past Built The ', options: { color: WHITE } },
    { text: 'Foundation Of Modern Civilization', options: { color: GOLD } },
  ], { size: 44, align: 'center', lh: 0.9 });
  text(s, 5.37, 1.74, 3.0, 0.404, 'History Of Civilization', { color: WHITE, italic: true, align: 'center' });
  hline(s, 3.67, 1.94, 1.69, WHITE);
  hline(s, 8.38, 1.94, 1.69, WHITE);
  button(s, 5.12, 5.42, 'Learn More', GOLD, INK);
  button(s, 6.74, 5.42, 'Continue', GOLD, INK);
  chrome(s, 2, false);
});

// 3 — Table of content
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 7.705, 0.6, 5.626, 6.897, BROWN);
  swoosh(s, -1.64, 5.88, 7.42, 1.67, BROWN, 5);
  text(s, 0.76, 1.686, 3.909, 1.918, [
    { text: 'Table of ', options: { color: INK } },
    { text: 'Content', options: { color: BROWN } },
  ], { size: 60, lh: 0.9 });
  eyebrow(s, 0.859, 1.358, INK);
  text(s, 0.76, 4.905, 3.568, 0.707, 'Here is a List Of Content Covers Several Topics');
  body(s, 0.76, 5.702, 3.528, 0.605, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit aenean commodo.', INK);
  button(s, 0.871, 6.499, 'Learn More', BROWN, WHITE);
  const agenda = [
    { no: '01.', title: 'Introduction to Ancient History ', w: 2.32 },
    { no: '02.', title: 'Major Civilizations Overview', w: 2.61 },
    { no: '03.', title: 'Rise and Fall of Civilizations', w: 2.09 },
    { no: '04.', title: 'Geography & Natural Resources', w: 2.61 },
  ];
  agenda.forEach((it, i) => numberedItem(s, 9.24, 1.378 + i * 1.466,
    { no: it.no, title: it.title, desc: 'Lorem ipsum dolor consectetuer' },
    { titleX: 9.96, titleW: it.w, descX: 9.96, descW: 2.73, descDY: 0.672, color: WHITE }));
  chrome(s, 3, false);
});

// 4 — Introduction
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, 1.64, 5.88, 7.42, 1.67, GOLD, 10);
  box(s, 0, 0.605, 3.527, 6.891, GOLD);
  text(s, 5.35, 1.813, 5.99, 1.74, [
    { text: 'Introduction to ', options: { color: WHITE } },
    { text: 'Ancient History ', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 5.51, 1.541, WHITE);
  text(s, 0.555, 4.8, 2.37, 0.707, 'Historical Significance');
  body(s, 0.555, 5.6, 2.42, 0.605, LOREM.short, INK);
  text(s, 0.555, 6.32, 2.15, 0.57, '1875 SM', { size: 28, color: BLACK });
  text(s, 5.35, 4.583, 3.38, 0.404, 'What Is Ancient History?', { color: WHITE });
  body(s, 5.35, 5.09, 3.68, 1.13, LOREM.long + 'Aenean massa. Cum sociis natoque penatibus et magnis dis.', WHITE);
  button(s, 5.43, 6.51, 'Learn More', GOLD, INK);
  playButton(s, 10.86, 5.49, 0.723);
  chrome(s, 4, true);
});

// 5 — Major civilizations overview
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  swoosh(s, -1.64, 5.88, 7.42, 1.67, BROWN, 5);
  box(s, 6.106, 4.833, 7.223, 2.079, BROWN);
  text(s, 0.818, 1.42, 6.955, 1.74, [
    { text: 'Major Civilizations ', options: { color: INK } },
    { text: 'Overview', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.98, 1.15, INK);
  text(s, 0.76, 3.507, 3.8, 0.707, 'The Past Built The Foundation Of Modern Civilization');
  body(s, 0.76, 4.306, 4.33, 0.605, LOREM.long, INK);
  hline(s, 0.76, 5.123, 4.25, INK);
  checkBullet(s, 0.818, 5.428, BROWN);
  body(s, 0.985, 5.336, 4.33, 0.34, 'Lorem ipsum dolor amet, consectetuer adipiscing.', INK);
  checkBullet(s, 0.818, 5.913, GOLD);
  body(s, 0.985, 5.816, 4.33, 0.34, 'Lorem ipsum dolor consectetuer adipiscing.', INK);
  button(s, 0.818, 6.5, 'Learn More', BROWN, WHITE);
  progressCard(s, 8.53, 2.146, { title: 'Discover The Origins Of Civilization', titleW: 2.89, pct: '85%', fill: 2.21 });
  chrome(s, 5, false);
});

// 6 — Religion and mythology
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 0, 4.81, 13.333, 2.69, BROWN);
  swoosh(s, 3.68, 5.6, 7.42, 1.67, GOLD, 10, true);
  text(s, 0.76, 1.42, 4.76, 2.56, [
    { text: 'Religion and Mythology ', options: { color: INK } },
    { text: 'Identity', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.14, INK);
  text(s, 0.76, 5.43, 3.57, 0.707, 'The Spiritual Understanding Of Early Civilizations', { color: WHITE });
  body(s, 4.71, 5.48, 3.53, 1.13, LOREM.long + 'Aenean massa. Cum sociis natoque penatibus et magnis.', WHITE);
  button(s, 0.76, 6.5, 'Learn More', GOLD, INK);
  donutCard(s, 8.2, 2.35, { pct: '78%', caption: 'Empires Rose With Innovation', captionW: 1.94 });
  chrome(s, 6, false);
});

// 7 — Ancient architecture
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  box(s, 0, 0.605, 6.27, 6.891, GOLD);
  swoosh(s, -1.64, 5.88, 7.42, 1.67, BROWN, 8);
  text(s, 0.76, 1.512, 5.02, 1.74, 'Ancient Architecture', { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.238, INK);
  text(s, 0.76, 3.856, 3.97, 0.707, 'Ancient Architecture Remains  Testament To Human Creativity');
  body(s, 0.76, 4.655, 4.33, 0.605, LOREM.long, INK);
  checkBullet(s, 0.855, 5.452, BROWN);
  body(s, 1.023, 5.36, 4.33, 0.34, 'Lorem ipsum dolor amet, consectetuer adipiscing.', INK);
  checkBullet(s, 0.855, 5.913, BROWN);
  body(s, 1.023, 5.82, 4.33, 0.34, 'Lorem ipsum dolor amet, consectetuer adipiscing.', INK);
  button(s, 0.855, 6.5, 'Learn More', BROWN, WHITE);

  text(s, 10.56, 0.945, 2.15, 0.57, '1832 SM', { size: 28, color: GOLD });
  hline(s, 10.66, 1.638, 2.21, WHITE);
  body(s, 10.56, 1.766, 2.4, 0.867, LOREM.quote, WHITE);
  button(s, 10.66, 2.9, 'Learn More', GOLD, INK);

  box(s, 9.02, 4.909, 4.31, 1.82, GOLD);
  s.addShape('ellipse', { x: 9.29, y: 5.226, w: 1.175, h: 1.175, fill: { color: BROWN }, line: NO_LINE });
  s.addShape('arc', { x: 9.41, y: 5.35, w: 0.943, h: 0.943, angleRange: [272.9, 157.2], line: { color: GOLD, width: 3.75 } });
  text(s, 9.5, 5.635, 0.761, 0.37, '75%', { size: 16, color: GOLD, align: 'center' });
  text(s, 10.66, 5.184, 2.16, 0.64, 'Ancient Societies Governance', { size: 16, color: BLACK });
  body(s, 10.66, 5.86, 2.41, 0.605, 'Lorem ipsum dolor sit amet, consectur adipiscing.', INK);
  chrome(s, 7, true);
});

// 8 — Preservation of ancient writing
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 0, 3.909, 13.333, 3.59, BROWN);
  swoosh(s, -1.46, 5.77, 7.42, 1.67, GOLD, 10);
  text(s, 6.71, 1.416, 5.86, 1.74, [
    { text: 'Preservation Of ', options: { color: INK } },
    { text: 'Ancient Writing', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 6.87, 1.142, INK);
  ['01.', '02.'].forEach((no, i) => numberedItem(s, 6.75, 4.668 + i * 1.513,
    { no, title: 'Knowledge Of Ancient Writing', desc: LOREM.tight },
    { titleX: 7.49, titleW: 2.17, descX: 9.83, descW: 3.06, descH: 0.605, color: WHITE }));
  hline(s, 6.71, 5.71, 6.05, WHITE);
  donutCard(s, 0.76, 2.37, { pct: '85%', caption: 'Environment To Survive', captionW: 1.74 });
  chrome(s, 8, false);
});

// 9 — Culture reveals ancient civilizations
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 6.95, 4.187, 6.383, 3.312, GOLD);
  swoosh(s, 7.09, 5.88, 7.42, 1.67, BROWN, 8, true);
  box(s, 0, 0.605, 2.319, 6.891, BROWN);
  text(s, 2.79, 1.416, 8.58, 1.74, [
    { text: 'Culture Reveals Ancient ', options: { color: INK } },
    { text: 'Civilizations Of Life', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 2.94, 1.142, INK);
  text(s, 8.95, 4.8, 3.49, 0.707, 'Understanding Ancient History Helps Us Recognize');
  body(s, 8.95, 5.59, 4.07, 0.605, 'Lorem ipsum dolor amet consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ', INK);
  text(s, 8.95, 6.32, 3.69, 0.57, '1835 SM \u2013 1859 SM', { size: 28, color: BROWN });
  progressCard(s, 0, 5.24, { title: 'Ancient Innovations Fuel Agricultural Change', pct: '85%', fill: 2.21 });
  chrome(s, 9, false);
});

// 10 — The legacy of ancient history
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, -1.07, 5.73, 7.42, 1.67, GOLD, 10);
  box(s, 11.212, 0.605, 2.121, 6.891, GOLD);
  text(s, 0.76, 1.562, 5.99, 2.56, [
    { text: 'The Legacy Of Ancient History ', options: { color: WHITE } },
    { text: 'Continues', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.288, WHITE);
  ['01.', '02.'].forEach((no, i) => numberedItem(s, 0.83, 4.74 + i * 1.513,
    { no, title: 'The Legacy Of Ancient History ', desc: LOREM.tight },
    { titleX: 1.56, titleW: 2.17, descX: 3.9, descW: 3.06, descH: 0.605, color: WHITE }));
  hline(s, 0.78, 5.78, 6.05, WHITE);
  donutCard(s, 7.79, 4.13, { pct: '85%', caption: 'Cultural Appreciation', captionW: 1.74 });
  playButton(s, 10.61, 2.47, 1.03);
  chrome(s, 10, true);
});

// 11 — Ancient discoveries of modern science
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  swoosh(s, 6.98, 5.83, 7.42, 1.67, BROWN, 5, true);
  box(s, 3.37, 5.84, 5.91, 1.66, GOLD);
  box(s, 0, 0.605, 3.37, 6.891, BROWN);
  text(s, 5.04, 1.416, 7.39, 1.74, [
    { text: 'Ancient Discoveries ', options: { color: INK } },
    { text: 'of Modern Science', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 5.2, 1.142, INK);
  text(s, 0.5, 4.8, 2.37, 0.707, 'Historical Significance', { color: WHITE });
  body(s, 0.5, 5.6, 2.31, 0.605, LOREM.short, WHITE);
  text(s, 0.5, 6.32, 2.15, 0.57, '1875 SM', { size: 28, color: GOLD });
  text(s, 5.04, 3.61, 3.57, 0.707, 'The Spiritual Understanding Of Early Civilizations');
  body(s, 8.72, 3.61, 3.53, 0.867, LOREM.long + 'Aenean massa. ', INK);
  button(s, 10.77, 6.5, 'Learn More', BROWN, WHITE);
  chrome(s, 11, false);
});

// 12 — Influencing how societies interpreted life
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  swoosh(s, -1.64, 5.88, 7.42, 1.67, BROWN, 5);
  box(s, 4.352, 0.605, 8.981, 6.891, BROWN);
  text(s, 6.56, 1.58, 5.91, 2.56, [
    { text: 'Influencing How Societies ', options: { color: WHITE } },
    { text: 'Interpreted Life', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 6.72, 1.31, WHITE);
  text(s, 6.56, 5.03, 2.52, 0.707, 'Empires Expanded Through Strategy', { color: WHITE });
  body(s, 9.25, 5.04, 3.53, 0.605, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit aenean commodo.', WHITE);
  [5.79, 6.17, 6.56].forEach(y => {
    checkBullet(s, 9.33, y + 0.092, GOLD);
    body(s, 9.49, y, 3.08, 0.34, 'Lorem ipsum dolor sit consectetuer.', WHITE);
  });
  button(s, 6.64, 6.5, 'Learn More', GOLD, INK);
  progressCard(s, 1.13, 4.42, { title: 'Inscriptions Powerful Evidence Of Ancient', titleW: 2.95, pct: '85%', fill: 2.21 });
  chrome(s, 12, false);
});

// 13 — Culture long before globalization
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, -0.99, 5.88, 7.42, 1.67, GOLD, 10);
  box(s, 11.949, 0.605, 1.382, 6.891, GOLD);
  text(s, 0.76, 1.361, 7.99, 1.74, [
    { text: 'Culture Long Before ', options: { color: WHITE } },
    { text: 'Globalization Existed', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.087, WHITE);
  text(s, 0.76, 4.14, 2.14, 0.707, '1862 SM', { size: 36, color: GOLD });
  vline(s, 3.115, 4.165, 0.656, WHITE);
  body(s, 3.33, 4.19, 3.53, 0.605, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo.', WHITE);
  body(s, 9.2, 2.01, 2.44, 0.867, LOREM.quote, WHITE);
  progressCard(s, 0.76, 5.21, { title: 'Ancient Cultures That Influenced World History', titleW: 3.22, pct: '72%', fill: 1.82 });
  progressCard(s, 5.24, 5.21, { title: 'Ancient Cultures That Influenced World History', titleW: 3.33, pct: '85%', fill: 2.21 });
  chrome(s, 13, true);
});

// 14 — From forgotten kingdoms to iconic empires
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  swoosh(s, 6.98, 5.83, 7.42, 1.67, BROWN, 5, true);
  box(s, 0, 0.605, 8.226, 6.891, BROWN);
  text(s, 0.818, 1.531, 6.07, 2.56, [
    { text: 'From Forgotten Kingdoms To ', options: { color: WHITE } },
    { text: 'Iconic Empires', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.98, 1.257, WHITE);
  [{ x: 0.818, year: '1848 SM' }, { x: 4.715, year: '1875 SM' }].forEach(col => {
    text(s, col.x, 4.8, 2.37, 0.707, 'World-famous Kingdom History', { color: WHITE });
    body(s, col.x, 5.6, 2.81, 0.605, 'Lorem ipsum dolor consectetuer adipiscing enean commodo.', WHITE);
    text(s, col.x, 6.32, 2.15, 0.57, col.year, { size: 28, color: WHITE });
  });
  vline(s, 4.174, 4.856, 1.981, WHITE);
  progressCard(s, 9.271, 3.417, { title: 'Ancient Innovations Fuel Agricultural Change', pct: '85%', fill: 2.21 });
  chrome(s, 14, false);
});

// 15 — Historical French military figure
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 5.156, 0.6, 8.177, 5.133, BROWN);
  swoosh(s, 6.93, 3.65, 7.42, 1.67, GOLD, 10, true);
  text(s, 5.81, 1.416, 6.45, 2.56, [
    { text: 'Historical French Military and ', options: { color: WHITE } },
    { text: 'Political Figure', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 5.97, 1.142, WHITE);
  box(s, 4.294, 4.462, 6.784, 2.436, GOLD);
  text(s, 7.267, 4.746, 3.08, 0.707, 'Napoleon Is The Title For Napoleon Bonaparte', { color: BLACK });
  const barY = 5.674, barW = 2.78, barH = 0.126;
  s.addShape('roundRect', { x: 7.367, y: barY, w: barW, h: barH, rectRadius: barH / 2, fill: { color: TRACK }, line: NO_LINE });
  s.addShape('roundRect', { x: 7.367, y: barY, w: 2.1, h: barH, rectRadius: barH / 2, fill: { color: BROWN }, line: NO_LINE });
  text(s, 10.166, 5.574, 0.63, 0.34, '85%', { size: 14, color: BLACK });
  body(s, 7.267, 6.007, 3.53, 0.605, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit aenean commodo.', INK);
  chrome(s, 15, false);
});

// 16 — Gallery of historical places
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, 2.88, 5.83, 7.42, 1.67, GOLD, 10, true);
  box(s, 0, 4.756, 7.521, 2.141, GOLD);
  box(s, 10.711, 0.605, 2.622, 6.891, GOLD);
  text(s, 0.76, 1.361, 7.79, 1.74, [
    { text: 'Gallery Of Historical ', options: { color: WHITE } },
    { text: 'Places In The World', options: { color: GOLD } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.087, WHITE);
  text(s, 0.844, 5.873, 3.22, 0.707, 'Ancient Cultures That Influenced World History', { color: BLACK });
  body(s, 4.232, 5.923, 3.06, 0.605, LOREM.tight, INK);
  chrome(s, 16, true);
});

// 17 — Visit historical places on internet
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 0, 0.6, 3.964, 6.897, BROWN);
  swoosh(s, 6.86, 5.88, 7.42, 1.67, BROWN, 5, true);
  imagePlaceholder(s, 1.5, 1.05, 4.0, 5.05, { rotate: -4, fill: '3A3A3A', line: '2B2B2B', inner: 'FFFFFF', radius: 0.18 });
  text(s, 6.37, 1.416, 6.2, 1.554, [
    { text: 'Visit Historical ', options: { color: INK } },
    { text: 'Places On Internet', options: { color: BROWN } },
  ], { size: 48, lh: 0.9 });
  eyebrow(s, 6.53, 1.142, INK);
  text(s, 7.54, 3.507, 3.8, 0.707, 'Various Types Of Historical Places That Can Be Visited');
  body(s, 7.54, 4.306, 4.55, 0.605, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor aenean massa. ', INK);
  hline(s, 7.54, 5.123, 4.55, INK);
  checkBullet(s, 7.6, 5.428, BROWN);
  body(s, 7.8, 5.336, 4.33, 0.34, LOREM.short, INK);
  checkBullet(s, 7.6, 5.913, GOLD);
  body(s, 7.8, 5.816, 4.33, 0.34, 'Lorem ipsum dolor consectetuer adipiscing.', INK);
  button(s, 7.6, 6.5, 'Learn More', BROWN, WHITE);
  progressCard(s, 0, 5.24, { title: 'Ancient Innovations Fuel Agricultural Change', pct: '85%', fill: 2.21 });
  chrome(s, 17, false);
});

// 18 — Distribution of historical places (world map)
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  box(s, 0, 3.492, 13.333, 4.008, BROWN);
  swoosh(s, -1.0, 5.77, 7.42, 1.67, GOLD, 10);
  text(s, 0.76, 1.33, 9.77, 1.74, [
    { text: 'Distribution Of Historical ', options: { color: INK } },
    { text: 'Places In The World', options: { color: BROWN } },
  ], { size: 54, lh: 0.9 });
  eyebrow(s, 0.92, 1.056, INK);
  worldMap(s);
  text(s, 0.89, 4.07, 3.59, 0.707, 'Belief Systems That Influence Our World Today', { color: GOLD });
  body(s, 0.89, 4.88, 4.41, 0.605, LOREM.long, WHITE);
  text(s, 0.88, 5.63, 3.69, 0.57, '1835 SM \u2013 1859 SM', { size: 28, color: GOLD });
  button(s, 0.96, 6.52, 'Learn More', GOLD, INK);
  chrome(s, 18, false);
});

// 19 — Timeline
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, 6.9, 5.83, 7.42, 1.67, GOLD, 10, true);
  text(s, 0.76, 1.562, 4.98, 1.74, 'Timeline of Ancient Eras', { size: 54, color: WHITE, lh: 0.9 });
  eyebrow(s, 0.92, 1.288, WHITE);
  vline(s, 7.0, 1.288, 6.21, WHITE);
  [0, 1, 2].forEach(i => {
    const y = 1.222 + i * 2.166;
    numberedItem(s, 6.72, y, { no: '01.', title: 'Historical Events Of Ancient Cultures', desc: 'Lorem ipsum dolor sit amet, consectur adipiscing.' },
      { titleX: 7.54, titleW: 2.52, descX: 10.24, descW: 2.33, descH: 0.605, color: WHITE });
    text(s, 7.54, y + 0.748, 2.4, 0.404, '1725 SM \u2013 1748 SM', { color: WHITE });
  });
  text(s, 3.27, 4.84, 2.47, 0.707, 'Redirected History\u2019s Timeline', { color: WHITE });
  body(s, 3.27, 5.69, 2.85, 0.605, 'Lorem ipsum dolor consectetuer adipiscing enean commodo.', WHITE);
  button(s, 3.27, 6.5, 'Learn More', GOLD, INK);
  chrome(s, 19, true);
});

// 20 — Closing
build.push(pptx => {
  const s = pptx.addSlide();
  s.background = { color: BROWN };
  swoosh(s, -1.03, 5.84, 7.42, 1.67, GOLD, 10);
  box(s, 7.52, 3.646, 4.222, 1.663, GOLD);
  text(s, 0.76, 1.416, 7.95, 2.766, [
    { text: 'Get in Touch ', options: { color: WHITE } },
    { text: 'With Us', options: { color: GOLD } },
  ], { size: 88, lh: 0.9 });
  eyebrowWide(s, 0.979, 1.194, WHITE);
  text(s, 0.76, 5.129, 3.877, 0.404, 'Thank You For Your Attention', { color: WHITE });
  body(s, 0.76, 5.625, 4.729, 0.605, LOREM.hero, WHITE);
  button(s, 0.872, 6.495, 'Learn More', GOLD, INK);
  button(s, 2.563, 6.495, 'Get in Touch', GOLD, INK, true);
  chrome(s, 20, true);
});

/* ── Render ────────────────────────────────────────────────────────────── */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: SERIF, bodyFontFace: SANS };
pptx.title = 'Ancient History Presentation';

build.forEach(fn => fn(pptx));

pptx.writeFile({ fileName: path.join(__dirname, '04d3ee0d-e648-46d4-b68b-1efa5c844a42_grok_final.pptx') })
  .then(f => console.log('wrote', f));
