/**
 * "Healence Clinic — Medical Consultations", a 20-slide 16:9 deck,
 * rebuilt from scratch with pptxgenjs.
 *
 *   node 02328832-0106-40d4-b146-2b1cab68cfaf_grok_final.js
 *
 * Photographs in the source deck are replaced by programmatic placeholders.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================== *
 * 1. Design tokens
 * ================================================================== */

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

const C = {
  cyan: '08A6E3',       // theme accent 3 — the brand blue
  cyanMid: '95DFFB',
  cyanPale: 'CAEFFD',
  cyanLight: '5FC7EF',
  ink: '404040',        // headings + body copy
  inkSoft: '3F3F3F',
  muted: 'BFBFBF',      // de-emphasised captions
  white: 'FFFFFF',
  grey05: 'F2F2F2',
  grey15: 'D9D9D9',
  glowBlue: '156BA8',   // colour of the soft radial blobs
  navyLo: '101A38',     // dark background gradient, bottom-left
  navyHi: '123F69',     // dark background gradient, top-right
  deviceEdge: '444444',
  placeholder: 'EDF1F3',
  placeholderEdge: 'D5DDE2',
  placeholderInk: '8FA6B2',
};

const FONT_HEAD = 'Figtree Medium';  // theme major font
const FONT_BODY = 'Inter';           // theme minor font

// PowerPoint's default text-box insets, in points: [left, right, bottom, top].
const INSET = [7.2, 7.2, 3.6, 3.6];

/* ================================================================== *
 * 2. Text helpers
 * ================================================================== */

/** Text box with the deck's standard insets, top-aligned. */
function text(slide, runs, opts) {
  slide.addText(runs, Object.assign(
    { valign: 'top', margin: INSET, fontFace: FONT_BODY, color: C.ink, isTextBox: true },
    opts));
}

/** Turn ['line', ['line', COLOR]] into stacked runs of one paragraph each. */
function lines(list) {
  return list.map((ln, i) => {
    const [t, color] = Array.isArray(ln) ? ln : [ln, null];
    const o = { breakLine: i < list.length - 1 };
    if (color) o.color = color;
    return { text: t, options: o };
  });
}

/** Stacked heading in the display face. */
function heading(slide, list, opts) {
  text(slide, lines(list), Object.assign({ fontFace: FONT_HEAD, color: C.ink }, opts));
}

/** Body copy: 12 pt, 1.3 line spacing, one paragraph per source line. */
function body(slide, list, opts) {
  text(slide, lines(list), Object.assign(
    { fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 }, opts));
}

/* ================================================================== *
 * 3. Brand furniture
 * ================================================================== */

// The "Healence" swirl, traced from the slide master's custom geometry.
// Points are fractions of the mark's bounding box.
const LOGO_PATH = [
  { x: 0.4167, y: 0.0000 },
  { x: 0.4372, y: 0.0007, c: [0.4236, 0.0000, 0.4304, 0.0002] },
  { x: 0.4375, y: 0.0004 },
  { x: 0.4377, y: 0.0000, c: [0.4375, 0.0002, 0.4376, 0.0000] },
  { x: 0.8186, y: 0.0000 },
  { x: 0.8333, y: 0.0208, c: [0.8267, 0.0000, 0.8333, 0.0093] },
  { x: 0.8290, y: 0.0355, c: [0.8333, 0.0263, 0.8318, 0.0316] },
  { x: 0.7291, y: 0.1765 },
  { x: 0.9419, y: 0.1765 },
  { x: 0.9740, y: 0.2069, c: [0.9561, 0.1765, 0.9691, 0.1882] },
  { x: 1.0000, y: 0.4118, c: [0.9908, 0.2707, 1.0000, 0.3397] },
  { x: 0.5833, y: 1.0000, c: [1.0000, 0.7366, 0.8135, 1.0000] },
  { x: 0.5628, y: 0.9993, c: [0.5764, 1.0000, 0.5696, 0.9997] },
  { x: 0.5625, y: 0.9996 },
  { x: 0.5623, y: 1.0000, c: [0.5625, 0.9998, 0.5624, 1.0000] },
  { x: 0.1814, y: 1.0000 },
  { x: 0.1667, y: 0.9792, c: [0.1733, 1.0000, 0.1667, 0.9907] },
  { x: 0.1710, y: 0.9645, c: [0.1667, 0.9737, 0.1682, 0.9684] },
  { x: 0.2708, y: 0.8235 },
  { x: 0.0581, y: 0.8235 },
  { x: 0.0260, y: 0.7931, c: [0.0439, 0.8235, 0.0309, 0.8118] },
  { x: 0.0000, y: 0.5882, c: [0.0092, 0.7293, 0.0000, 0.6603] },
  { x: 0.4167, y: 0.0000, c: [0.0000, 0.2634, 0.1865, 0.0000] },
  { close: true },
  { x: 0.5592, y: 0.4164, move: true },
  { x: 0.4167, y: 0.3235, c: [0.5249, 0.3596, 0.4738, 0.3235] },
  { x: 0.2292, y: 0.5882, c: [0.3131, 0.3235, 0.2292, 0.4420] },
  { x: 0.2399, y: 0.6765, c: [0.2292, 0.6192, 0.2330, 0.6489] },
  { x: 0.3750, y: 0.6765 },
  { x: 0.4408, y: 0.5836 },
  { x: 0.5833, y: 0.6765, c: [0.4751, 0.6404, 0.5262, 0.6765] },
  { x: 0.7708, y: 0.4118, c: [0.6869, 0.6765, 0.7708, 0.5580] },
  { x: 0.7601, y: 0.3235, c: [0.7708, 0.3808, 0.7670, 0.3511] },
  { x: 0.6250, y: 0.3235 },
  { x: 0.5592, y: 0.4164 },
  { close: true },
];

/** Map a normalised path onto a w x h box (pptxgenjs custGeom points). */
function scalePath(pts, w, h) {
  return pts.map(p => {
    if (p.close) return { close: true };
    const o = { x: p.x * w, y: p.y * h };
    if (p.move) o.moveTo = true;
    if (p.c) o.curve = { type: 'cubic', x1: p.c[0] * w, y1: p.c[1] * h, x2: p.c[2] * w, y2: p.c[3] * h };
    return o;
  });
}

/** Swirl mark + "Healence Clinic" wordmark (top-left of nearly every slide). */
function logoMark(slide, opts) {
  const o = Object.assign({ x: 0.8, y: 0.689, scale: 1, color: C.ink, size: 10 }, opts);
  const w = 0.302 * o.scale, h = 0.216 * o.scale;
  slide.addShape('custGeom', { x: o.x, y: o.y, w, h, fill: { color: C.cyan }, points: scalePath(LOGO_PATH, w, h) });
  text(slide, [{ text: 'Healence Clinic ' }], {
    x: o.x + 0.435 * o.scale, y: o.y - 0.027, w: 1.705, h: 0.269,
    fontSize: o.size, fontFace: FONT_HEAD, color: o.color,
  });
}

/** Right-aligned page number in the header band. */
function pageNumber(slide, n, color) {
  text(slide, [{ text: String(n) }], {
    x: 9.641, y: 0.662, w: 3.0, h: 0.269,
    align: 'right', fontSize: 10, fontFace: FONT_HEAD, color: color || C.ink,
  });
}

/* ================================================================== *
 * 4. Icon badges
 * ================================================================== */

/**
 * Pictographs drawn from primitives. Each receives the badge centre, the
 * badge diameter `s`, the glyph colour and the disc colour behind it.
 */
const GLYPHS = {
  // Microscope: eyepiece and body tube on the left, arm curving round to the
  // right, standing on a stage bar and a footed base.
  microscope(sl, cx, cy, s, ink) {
    const g = 0.40 * s;                              // glyph fills ~40 % of the badge
    const X = u => cx + (u - 0.5) * g, Y = v => cy + (v - 0.5) * g, S = k => k * g;
    sl.addShape('roundRect', { x: X(0.34), y: Y(0.00), w: S(0.11), h: S(0.18), rectRadius: S(0.055), fill: { color: ink } });
    sl.addShape('roundRect', { x: X(0.26), y: Y(0.12), w: S(0.26), h: S(0.46), rectRadius: S(0.08), fill: { color: ink } });
    sl.addShape('blockArc', {
      x: X(0.49), y: Y(0.21), w: S(0.52), h: S(0.62),
      angleRange: [270, 90], arcThicknessRatio: 0.34, fill: { color: ink },
    });
    sl.addShape('roundRect', { x: X(0.13), y: Y(0.66), w: S(0.51), h: S(0.09), rectRadius: S(0.045), fill: { color: ink } });
    sl.addShape('rect', { x: X(0.35), y: Y(0.72), w: S(0.09), h: S(0.15), fill: { color: ink } });
    sl.addShape('roundRect', { x: X(0.00), y: Y(0.84), w: S(0.86), h: S(0.09), rectRadius: S(0.045), fill: { color: ink } });
    [0.08, 0.60].forEach(u => sl.addShape('rect', { x: X(u), y: Y(0.90), w: S(0.09), h: S(0.10), fill: { color: ink } }));
  },
  // Stethoscope: binaural tubes joining in a U-bend, a lead running out to the
  // right, and the round chest piece up top.
  stethoscope(sl, cx, cy, s, ink) {
    const g = 0.40 * s;
    const X = u => cx + (u - 0.5) * g, Y = v => cy + (v - 0.5) * g, S = k => k * g;
    [0.02, 0.49].forEach(u => sl.addShape('roundRect', {
      x: X(u), y: Y(0.03), w: S(0.09), h: S(0.30), rectRadius: S(0.045), fill: { color: ink },
    }));
    sl.addShape('blockArc', {
      x: X(0.02), y: Y(0.06), w: S(0.56), h: S(0.56),
      angleRange: [0, 180], arcThicknessRatio: 0.16, fill: { color: ink },
    });
    sl.addShape('rect', { x: X(0.26), y: Y(0.56), w: S(0.08), h: S(0.20), fill: { color: ink } });
    sl.addShape('blockArc', {
      x: X(0.26), y: Y(0.62), w: S(0.60), h: S(0.76),
      angleRange: [0, 180], arcThicknessRatio: 0.13, fill: { color: ink },
    });
    sl.addShape('rect', { x: X(0.82), y: Y(0.34), w: S(0.07), h: S(0.42), fill: { color: ink } });
    sl.addShape('ellipse', {
      x: X(0.73), y: Y(0.13), w: S(0.25), h: S(0.25),
      fill: { type: 'none' }, line: { color: ink, width: 2.2 * (s / 0.485) },
    });
  },
  // Clinic: house silhouette with a rounded cross knocked out of it.
  clinic(sl, cx, cy, s, ink, disc) {
    const g = 0.40 * s;
    const X = u => cx + (u - 0.5) * g, Y = v => cy + (v - 0.5) * g, S = k => k * g;
    sl.addShape('custGeom', {
      x: X(0), y: Y(0), w: S(1), h: S(1), fill: { color: ink },
      points: [
        { x: S(0.50), y: 0 },
        { x: S(0.88), y: S(0.28), curve: { type: 'cubic', x1: S(0.62), y1: 0, x2: S(0.80), y2: S(0.22) } },
        { x: S(1.00), y: S(0.46), curve: { type: 'cubic', x1: S(0.96), y1: S(0.34), x2: S(1.00), y2: S(0.40) } },
        { x: S(1.00), y: S(0.82) },
        { x: S(0.82), y: S(1.00), curve: { type: 'cubic', x1: S(1.00), y1: S(0.92), x2: S(0.92), y2: S(1.00) } },
        { x: S(0.18), y: S(1.00) },
        { x: S(0.00), y: S(0.82), curve: { type: 'cubic', x1: S(0.08), y1: S(1.00), x2: S(0.00), y2: S(0.92) } },
        { x: S(0.00), y: S(0.46) },
        { x: S(0.12), y: S(0.28), curve: { type: 'cubic', x1: S(0.00), y1: S(0.40), x2: S(0.04), y2: S(0.34) } },
        { close: true },
      ],
    });
    // Cross: two crossed rounded bars knocked out in the disc's own colour.
    sl.addShape('roundRect', { x: X(0.22), y: Y(0.46), w: S(0.56), h: S(0.22), rectRadius: S(0.11), fill: { color: disc } });
    sl.addShape('roundRect', { x: X(0.39), y: Y(0.29), w: S(0.22), h: S(0.56), rectRadius: S(0.11), fill: { color: disc } });
  },
  envelope(sl, cx, cy, s, ink) {
    const g = 0.44 * s;
    const X = u => cx + (u - 0.5) * g, Y = v => cy + (v - 0.5) * g, S = k => k * g;
    sl.addShape('roundRect', { x: X(0), y: Y(0.12), w: S(1), h: S(0.76), rectRadius: S(0.14), fill: { color: ink } });
    sl.addShape('custGeom', {
      x: X(0.06), y: Y(0.18), w: S(0.88), h: S(0.44), fill: { color: C.cyan },
      points: [{ x: 0, y: 0 }, { x: S(0.88), y: 0 }, { x: S(0.44), y: S(0.44) }, { close: true }],
    });
  },
  globe(sl, cx, cy, s, ink) {
    const g = 0.46 * s;
    const X = u => cx + (u - 0.5) * g, Y = v => cy + (v - 0.5) * g, S = k => k * g;
    sl.addShape('ellipse', { x: X(0.03), y: Y(0.03), w: S(0.80), h: S(0.80), fill: { type: 'none' }, line: { color: ink, width: 1.4 } });
    sl.addShape('ellipse', { x: X(0.30), y: Y(0.03), w: S(0.26), h: S(0.80), fill: { type: 'none' }, line: { color: ink, width: 1.1 } });
    sl.addShape('rect', { x: X(0.03), y: Y(0.41), w: S(0.80), h: S(0.05), fill: { color: ink } });
    sl.addShape('custGeom', {
      x: X(0.46), y: Y(0.44), w: S(0.54), h: S(0.60), fill: { color: ink },
      points: [{ x: 0, y: 0 }, { x: S(0.54), y: S(0.34) }, { x: S(0.28), y: S(0.38) }, { x: S(0.18), y: S(0.60) }, { close: true }],
    });
  },
  phone(sl, cx, cy, s, ink) {
    const g = 0.44 * s;
    sl.addShape('blockArc', {
      x: cx - 0.5 * g, y: cy - 0.5 * g, w: g, h: g,
      angleRange: [200, 105], arcThicknessRatio: 0.62, fill: { color: ink }, rotate: 25,
    });
  },
};

/** Circular (or squircle) icon badge: coloured disc plus a pictograph. */
function iconBadge(slide, x, y, d, opts) {
  const o = Object.assign({ disc: C.cyan, ink: C.white, glyph: 'microscope', round: true, transparency: 0 }, opts);
  const geo = { x, y, w: d, h: d, fill: { color: o.disc } };
  if (o.transparency) geo.fill.transparency = o.transparency;
  if (!o.round) geo.rectRadius = d / 2;
  slide.addShape(o.round ? 'ellipse' : 'roundRect', geo);
  if (GLYPHS[o.glyph]) GLYPHS[o.glyph](slide, x + d / 2, y + d / 2, d, o.ink, o.disc);
}

/** The recurring trio of badges: solid cyan, 60 % cyan, 20 % cyan. */
function badgeTrio(slide, x, y) {
  const d = 0.485, gap = 0.5995;
  iconBadge(slide, x, y, d, { glyph: 'microscope' });
  iconBadge(slide, x + gap, y, d, { glyph: 'stethoscope', transparency: 40 });
  iconBadge(slide, x + 2 * gap, y, d, { glyph: 'clinic', transparency: 80, disc: C.cyanPale });
}

/* ================================================================== *
 * 5. Repeated composite elements
 * ================================================================== */

/** "01  Health Guidance" — numbered disc plus a label. */
function numberedItem(slide, x, y, d, num, label, opts) {
  const o = Object.assign({ disc: C.cyan, numColor: C.white }, opts);
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: o.disc } });
  text(slide, [{ text: num }], {
    x: x - 0.033, y: y + 0.087, w: d + 0.066, h: 0.269,
    align: 'center', fontSize: 10, fontFace: FONT_HEAD, color: o.numColor,
  });
  text(slide, [{ text: label }], { x: x + 0.567, y: y + 0.07, w: 1.957, h: 0.303, fontSize: 12 });
}

/** "80%   Medical Consultation Connects Patients" stat line. */
function statLine(slide, x, y, value, label, valueColor) {
  text(slide, [{ text: value }], {
    x, y, w: 1.637, h: 0.572, fontSize: 28, fontFace: FONT_HEAD, color: valueColor || C.ink,
  });
  text(slide, [{ text: label }], { x: x + 1.191, y: y + 0.135, w: 5.311, h: 0.303, fontSize: 12 });
}

// Very soft drop shadow on the floating white cards. Returns a fresh object
// each time because pptxgenjs rewrites the options it is handed.
const cardShadow = () => ({ type: 'outer', color: '000000', blur: 26, offset: 0.05, angle: 90, opacity: 0.1 });

/** Donut chart card headed "Grocery" with a "+562" centre label. */
function groceryCard(slide, x, y) {
  slide.addShape('roundRect', { x, y, w: 2.175, h: 2.438, rectRadius: 0.22, fill: { color: C.white }, shadow: cardShadow() });
  text(slide, [{ text: 'Grocery' }], {
    x: x + 0.358, y: y + 0.276, w: 1.493, h: 0.303,
    align: 'center', fontSize: 12, fontFace: FONT_HEAD, color: C.inkSoft,
  });
  // Three arcs, each a slice of the ring: 40 % cyan, 21 % pale, 18 % dark.
  const dx = x + 0.535, dy = y + 0.923, d = 1.105;
  [[205.0, 270.8, C.ink], [270.0, 133.8, C.cyan], [132.9, 206.3, C.cyanPale]]
    .forEach(([from, to, color]) => slide.addShape('arc', {
      x: dx, y: dy, w: d, h: d, angleRange: [from, to], line: { color, width: 20 },
    }));
  text(slide, [{ text: '+562' }], {
    x: x + 0.654, y: y + 1.3, w: 0.867, h: 0.343,
    align: 'center', fontSize: 12, bold: true, color: C.inkSoft, lineSpacingMultiple: 1.2,
  });
}

/* ================================================================== *
 * 6. Backgrounds, glows and image stand-ins
 * ================================================================== */

function mix(a, b, t) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  const v = i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0');
  return (v(0) + v(1) + v(2)).toUpperCase();
}

/**
 * Navy background of the title / break / closing slides: a diagonal ramp
 * that is darkest at bottom-left and lightest at top-right. Painted as a
 * coarse grid of flat tiles because pptxgenjs has no gradient fill.
 */
function darkBackground(slide) {
  const cols = 16, rows = 8;
  const cw = SLIDE_W / cols, rh = SLIDE_H / rows;
  slide.background = { color: C.navyLo };
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = c * cw, y = r * rh;
      const t = ((x + cw / 2) - (y + rh / 2) + 8.2) / 43;
      slide.addShape('rect', {
        x, y, w: cw + 0.02, h: rh + 0.02,
        fill: { color: mix(C.navyLo, C.navyHi, Math.max(0, Math.min(1, t))) },
      });
    }
  }
}

// Opacity of the two source blobs, sampled at r = 0.05 … 1.00 of the radius.
const GLOW_STRONG = [0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80,
  0.79, 0.77, 0.71, 0.59, 0.42, 0.24, 0.11, 0.04, 0.01, 0.00];
const GLOW_SOFT = [0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.49, 0.47, 0.44,
  0.38, 0.31, 0.23, 0.16, 0.09, 0.05, 0.02, 0.01, 0.00, 0.00];

/**
 * Soft radial blue blob. The source deck uses large blurred PNGs; the same
 * falloff is rebuilt from 20 concentric translucent discs painted largest
 * first. Each disc's own opacity is solved so the *composited* coverage at
 * its radius equals the sampled profile above.
 */
function glow(slide, cx, cy, radius, profile) {
  const n = profile.length;
  let carried = 1;                                   // light still getting through
  for (let k = n - 1; k >= 0; k--) {
    const alpha = 1 - (1 - profile[k]) / carried;
    if (alpha <= 0.003) continue;
    carried *= (1 - alpha);
    const rr = radius * ((k + 1) / n);
    slide.addShape('ellipse', {
      x: cx - rr, y: cy - rr, w: rr * 2, h: rr * 2,
      fill: { color: C.glowBlue, transparency: Math.round((1 - alpha) * 1000) / 10 },
      line: { type: 'none' },
    });
  }
}

const glowStrong = (s, cx, cy, r) => glow(s, cx, cy, r, GLOW_STRONG);
const glowSoft = (s, cx, cy, r) => glow(s, cx, cy, r, GLOW_SOFT);

/** Neutral stand-in for a photograph. */
function imagePlaceholder(slide, x, y, w, h, opts) {
  const o = Object.assign({ radius: 0, label: '[image]' }, opts);
  const box = { x, y, w, h, fill: { color: C.placeholder }, line: { color: C.placeholderEdge, width: 1 } };
  if (o.radius) box.rectRadius = o.radius;
  slide.addShape(o.radius ? 'roundRect' : 'rect', box);
  captionPlaceholder(slide, x, y, w, h, o.label);
}

/** Centre the "[image]" caption on the part of a stand-in that is on-slide. */
function captionPlaceholder(slide, x, y, w, h, label) {
  if (!label) return;
  const left = Math.max(x, 0.2), right = Math.min(x + w, SLIDE_W - 0.2);
  const top = Math.max(y, 0.2), bottom = Math.min(y + h, SLIDE_H - 0.2);
  text(slide, [{ text: label }], {
    x: left, y: (top + bottom) / 2 - 0.15, w: Math.max(right - left, 0.6), h: 0.3,
    align: 'center', fontSize: 11, color: C.placeholderInk,
  });
}

/* --- Device mock-ups (stand-ins for the three product photos) ------ */

/** Open laptop seen head-on; the base runs off the bottom of the slide. */
function laptopMockup(slide, x, y, w) {
  const lidH = w * 0.664;
  slide.addShape('roundRect', { x, y, w, h: lidH, rectRadius: 0.07, fill: { color: '2B2B2B' } });
  slide.addShape('rect', { x: x + 0.1, y: y + 0.1, w: w - 0.2, h: lidH - 0.2, fill: { color: C.white } });
  slide.addShape('roundRect', { x: x - 0.05, y: y + lidH, w: w + 0.1, h: 0.11, rectRadius: 0.055, fill: { color: 'B9BDC0' } });
  slide.addShape('roundRect', { x: x + 0.06, y: y + lidH + 0.24, w: w - 0.12, h: 3.4, rectRadius: 0.09, fill: { color: 'C3C6C9' } });
  slide.addShape('roundRect', { x: x + 0.34, y: y + lidH + 0.36, w: w - 0.68, h: 0.22, rectRadius: 0.09, fill: { color: 'D6D9DB' } });
  slide.addShape('roundRect', { x: x + 0.36, y: y + lidH + 0.76, w: w - 0.72, h: 1.65, rectRadius: 0.05, fill: { color: '9EA3A8' } });
  slide.addShape('roundRect', { x: x + 1.6, y: y + lidH + 2.6, w: w - 3.2, h: 0.8, rectRadius: 0.05, fill: { color: 'AEB2B6' } });
  captionPlaceholder(slide, x, y, w, lidH, '[image]');
}

/** Phone standing at the top of the slide, screen showing through. */
function phoneMockup(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.62, fill: { color: C.white, transparency: 35 }, line: { color: '1D1D1D', width: 6 } });
  slide.addShape('roundRect', { x: x + w * 0.36, y: y + 0.12, w: w * 0.28, h: 0.15, rectRadius: 0.075, fill: { color: '1D1D1D' } });
  captionPlaceholder(slide, x, y + h - 1.2, w, 0.6, '[image]');
}

/** Tablet lying in the top-left corner. */
function tabletMockup(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.34, fill: { color: C.white }, line: { color: '1D1D1D', width: 4 } });
  captionPlaceholder(slide, x, y, w, h, '[image]');
}

/* ================================================================== *
 * 7. Shared copy
 * ================================================================== */

const TAG_CARE = 'Medical Consultation Connects Patients Care';
const TAG_PATIENTS = 'Medical Consultation Connects Patients';
const BODY_2 = ['Medical consultations play a vital role in ensuring ',
  'accurate diagnosis and effective treatment.'];
const BODY_3 = ['Medical consultations play a vital role in ensuring ',
  'accurate diagnosis, effective treatment, and ongoing ',
  'continuous patient care.'];
const BODY_ICON = ['Medical consultations play a ', 'vital role in ensuring accurate ', 'medical diagnosis.'];

/* ================================================================== *
 * 8. Slides
 * ================================================================== */

/* 1 — Title */
function slide01(s) {
  darkBackground(s);
  glowStrong(s, 11.207, -0.971, 5.403);
  logoMark(s, { y: 1.476, scale: 1.19, color: C.white, size: 14 });
  text(s, [{ text: TAG_CARE }], { x: 7.332, y: 6.476, w: 5.311, h: 0.303, align: 'right', fontSize: 12, color: C.white });
  text(s, [{ text: 'Medical' }], { x: 0.561, y: 2.395, w: 9.339, h: 2.036, fontSize: 115, fontFace: FONT_HEAD, color: C.white });
  text(s, [{ text: 'Consultations' }], {
    x: 0, y: 4.214, w: 12.709, h: 2.036, align: 'right', fontSize: 115, fontFace: FONT_HEAD, color: C.cyanLight,
  });
  [0.802, 1.244, 1.685].forEach((x, i) => s.addShape('ellipse', {
    x, y: 5.053, w: 0.357, h: 0.357, fill: { color: C.cyan, transparency: [0, 40, 80][i] },
  }));
  text(s, [{ text: 'www.yourwebsite.com' }], { x: 0.691, y: 6.476, w: 5.311, h: 0.303, fontSize: 12, color: C.white });
}

/* 2 — Professional team */
function slide02(s) {
  logoMark(s);
  pageNumber(s, 2);
  // Portrait card: transparent white at the top washing into cyan at the foot.
  const bands = 24, bx = 3.504, by = 1.403, bw = 2.88, bh = 2.88;
  for (let i = 0; i < bands; i++) {
    const t = i / (bands - 1);
    const p = Math.max(0, (t - 0.28) / 0.72);
    s.addShape('rect', {
      x: bx, y: by + t * bh, w: bw, h: bh / bands + 0.02,
      fill: { color: mix(C.white, C.cyan, p), transparency: Math.round(90 - 83 * p) },
    });
  }
  text(s, [{ text: 'Avery Davis' }], { x: 3.886, y: 3.349, w: 2.116, h: 0.404, align: 'center', fontSize: 18, fontFace: FONT_HEAD, color: C.white });
  text(s, [{ text: 'The Medical Consultations' }], { x: 3.516, y: 3.747, w: 2.858, h: 0.269, align: 'center', fontSize: 10, color: C.white });
  text(s, [
    { text: 'Our Professional ', options: { breakLine: true } },
    { text: 'Medical ' },
    { text: 'Team Experts', options: { color: C.cyan } },
  ], { x: 0.659, y: 5.34, w: 9.339, h: 1.447, fontSize: 40, fontFace: FONT_HEAD, color: C.ink, valign: 'top', margin: INSET });
  body(s, ['Medical consultations play a vital role in ensuring accurate',
    ' diagnosis and effective treatment for patients worldwide.'],
    { x: 6.3, y: 4.887, w: 6.342, h: 0.602, align: 'right' });
  badgeTrio(s, 10.849, 6.094);
}

/* 3 — Digital transformation (laptop) */
function slide03(s) {
  glowSoft(s, 13.417, 0.560, 4.764);
  laptopMockup(s, 6.913, 1.397, 5.013);
  logoMark(s);
  pageNumber(s, 3);
  heading(s, ['Digital ', 'Transformation in ', ['Consultations', C.cyan]],
    { x: 0.659, y: 1.9, w: 6.293, h: 2.121, fontSize: 40 });
  body(s, BODY_2, { x: 0.69, y: 4.437, w: 5.247, h: 0.602 });
  numberedItem(s, 0.803, 6.083, 0.442, '01', 'Health Guidance');
  numberedItem(s, 3.465, 6.083, 0.442, '02', 'Accurate Diagnosis', { disc: C.cyanPale, numColor: C.cyan });
}

/* 4 — Collaboration */
function slide04(s) {
  logoMark(s);
  pageNumber(s, 4);
  heading(s, ['Collaboration Between Specialists', ['and Patients Together', C.cyan]],
    { x: 0.668, y: 1.706, w: 8.148, h: 1.178, fontSize: 32 });
  text(s, [{ text: TAG_CARE }], { x: 0.694, y: 3.28, w: 4.534, h: 0.303, fontSize: 12 });
  badgeTrio(s, 0.8, 5.224);
  body(s, BODY_2, { x: 0.69, y: 6.042, w: 5.247, h: 0.602 });
}

/* 5 — Patient satisfaction */
function slide05(s) {
  glowSoft(s, 12.718, -0.989, 4.383);
  pageNumber(s, 5, C.white);
  heading(s, ['Enhancing ', ['Patient Satisfaction ', C.cyan], 'and Outcomes'],
    { x: 6.929, y: 1.764, w: 6.405, h: 1.919, fontSize: 36 });
  body(s, BODY_3, { x: 6.96, y: 4.631, w: 6.373, h: 0.863 });
  statLine(s, 0.681, 6.05, '80%', TAG_PATIENTS, C.cyan);
  badgeTrio(s, 10.849, 6.094);
}

/* 6 — Data-driven decisions (bar chart) */
function slide06(s) {
  logoMark(s);
  pageNumber(s, 6);
  // Three bars rising left to right, all anchored to the slide's bottom edge.
  [[0.8, 1.823, C.cyanPale], [4.727, 2.74, C.cyanMid], [8.655, 3.656, C.cyan]]
    .forEach(([x, h, color]) => s.addShape('rect', { x, y: SLIDE_H - h, w: 3.573, h, fill: { color } }));
  text(s, [{ text: 'Data-Driven Medical ', options: { color: C.cyan, breakLine: true } }, { text: 'Decision Making' }],
    { x: 0.659, y: 1.866, w: 8.148, h: 1.313, fontSize: 36, fontFace: FONT_HEAD, color: C.ink, valign: 'top', margin: INSET });
  text(s, [{ text: '60%' }], { x: 8.957, y: 4.089, w: 2.987, h: 1.01, fontSize: 54, fontFace: FONT_HEAD, color: C.white });
  body(s, ['The consultations improve health outcomes significantly.'],
    { x: 8.957, y: 6.251, w: 2.55, h: 0.517, fontSize: 10, color: C.white });
  body(s, ['Medical consultations play a vital role in ', 'ensuring accurate diagnosis and ongoing ', 'continuous patient care.'],
    { x: 8.66, y: 2.226, w: 6.373, h: 0.863 });
  badgeTrio(s, 0.8, 3.75);
}

/* 7 — Personalized care (phone) */
function slide07(s) {
  glowSoft(s, 9.069, 0.036, 3.737);
  phoneMockup(s, 8.976, -2.328, 3.475, 7.095);
  logoMark(s);
  heading(s, ['Personalized Care ', 'Through Medical Advice'], { x: 0.68, y: 1.65, w: 8.148, h: 1.313, fontSize: 36 });
  statLine(s, 0.681, 3.336, '80%', TAG_PATIENTS);
  iconBadge(s, 4.122, 5.956, 0.643, { glyph: 'microscope' });
  body(s, BODY_ICON, { x: 5.063, y: 5.846, w: 3.275, h: 0.863 });
  iconBadge(s, 8.828, 5.956, 0.643, { disc: C.cyanPale, ink: C.cyan, glyph: 'clinic' });
  body(s, BODY_ICON, { x: 9.768, y: 5.846, w: 3.275, h: 0.863 });
}

/* 8 — Enhancing satisfaction */
function slide08(s) {
  logoMark(s);
  pageNumber(s, 8);
  heading(s, ['Enhancing ', ['Patient Satisfaction ', C.cyan], 'and Outcomes'],
    { x: 0.661, y: 1.838, w: 8.148, h: 1.919, fontSize: 36 });
  text(s, [{ text: TAG_CARE }], { x: 0.685, y: 4.106, w: 5.311, h: 0.303, fontSize: 12, color: C.muted });
  body(s, BODY_3, { x: 0.689, y: 5.6, w: 6.373, h: 0.863 });
}

/* 9 — Importance of consultations (tablet) */
function slide09(s) {
  tabletMockup(s, -0.267, -1.614, 5.012, 7.213);
  pageNumber(s, 9);
  heading(s, ['The Importance of Effective ', 'Medical Consultations for Better ', ['Patient Outcomes', C.cyan]],
    { x: 5.587, y: 1.608, w: 9.339, h: 1.717, fontSize: 32 });
  iconBadge(s, 5.722, 5.039, 0.522, { glyph: 'microscope' });
  body(s, BODY_ICON, { x: 5.614, y: 5.812, w: 3.275, h: 0.863 });
  iconBadge(s, 9.404, 5.039, 0.522, { disc: C.cyanPale, ink: C.cyan, glyph: 'stethoscope' });
  body(s, BODY_ICON, { x: 9.296, y: 5.812, w: 3.275, h: 0.863 });
  text(s, [{ text: TAG_PATIENTS }], { x: 0.685, y: 6.372, w: 5.311, h: 0.303, fontSize: 12 });
}

/* 10 — Break divider */
function slide10(s) {
  darkBackground(s);
  glowSoft(s, 13.325, 3.751, 5.507);
  logoMark(s, { color: C.white });
  pageNumber(s, 10, C.white);
  text(s, lines(['It\u2019s time to Take ', 'a Break']), {
    x: 0.634, y: 1.854, w: 9.339, h: 2.121,
    fontSize: 60, fontFace: FONT_HEAD, bold: true, color: C.white, valign: 'top', margin: INSET,
  });
  iconBadge(s, 0.808, 5.003, 0.643, { glyph: 'microscope' });
  text(s, [{ text: '60 Minutes' }], { x: 1.644, y: 5.039, w: 4.236, h: 0.572, fontSize: 28, fontFace: FONT_HEAD, color: C.cyan });
  body(s, BODY_2, { x: 0.69, y: 6.042, w: 5.247, h: 0.602, color: C.white });
}

/* 11 — Personalized care, 80 % headline */
function slide11(s) {
  logoMark(s);
  pageNumber(s, 11);
  heading(s, ['Personalized Care ', 'Through Medical Advice'], { x: 0.657, y: 1.65, w: 6.844, h: 1.178, fontSize: 32 });
  text(s, [{ text: TAG_CARE }], { x: 0.685, y: 3.169, w: 5.311, h: 0.303, fontSize: 12, color: C.muted });
  text(s, [{ text: '80%' }], { x: 0.678, y: 4.57, w: 4.236, h: 1.582, fontSize: 88, fontFace: FONT_HEAD, color: C.cyan });
  body(s, ['A medical consultations play a vital role in ', 'ensuring accurate medical diagnosis.'],
    { x: 0.699, y: 6.152, w: 4.236, h: 0.601 });
  iconBadge(s, 9.289, 2.124, 0.738, { glyph: 'microscope' });
  text(s, [{ text: 'Health Guidance' }], { x: 10.376, y: 2.176, w: 4.035, h: 0.404, fontSize: 18 });
  body(s, ['Presentation Template'], { x: 10.397, y: 2.524, w: 2.296, h: 0.286, fontSize: 10, lineSpacingMultiple: 1.2 });
}

/* 12 — Ethics (dot-matrix infographic) */
function slide12(s) {
  logoMark(s);
  pageNumber(s, 12);
  heading(s, ['Ethics in Medical Consultations'], { x: 0.657, y: 1.777, w: 7.294, h: 1.582, fontSize: 44 });
  body(s, ['Medical consultations play a vital role in ensuring ', 'accurate diagnosis and the effective treatment.'],
    { x: 0.685, y: 3.704, w: 6.378, h: 0.602 });

  // 6 x 5 grid of discs; 1 = highlighted (cyan), 0 = neutral grey.
  const grid = [
    [0, 0, 0, 0, 0, 0],
    [0, 0, 1, 0, 0, 0],
    [0, 1, 1, 0, 0, 1],
    [0, 1, 1, 0, 1, 1],
    [1, 1, 1, 1, 1, 1],
  ];
  const gx = 6.913, gy = 1.381, step = 0.9435, d = 0.901;
  grid.forEach((row, r) => row.forEach((on, c) => s.addShape('ellipse', {
    x: gx + c * step, y: gy + r * step, w: d, h: d, fill: { color: on ? C.cyan : C.grey05 },
  })));

  // Two floating callouts over the grid.
  [[7.488, 1.913], [10.331, 3.779]].forEach(([x, y]) => {
    s.addShape('roundRect', { x, y, w: 1.76, h: 0.78, rectRadius: 0.1, fill: { color: C.white }, shadow: cardShadow() });
    s.addShape('ellipse', { x: x + 0.214, y: y + 0.323, w: 0.135, h: 0.135, fill: { color: C.cyan } });
    text(s, [{ text: '52,8%' }], { x: x + 0.464, y: y + 0.171, w: 1.362, h: 0.438, fontSize: 20, fontFace: FONT_HEAD });
  });

  statLine(s, 6.913, 6.331, '80%', TAG_CARE);
  [['01', 'Personalized Medical Care', 5.631, C.cyan, C.white],
   ['02', 'Expert Health Guidance', 6.291, C.cyanPale, C.cyan]].forEach(([num, label, y, disc, numColor]) => {
    s.addShape('ellipse', { x: 0.805, y, w: 0.478, h: 0.478, fill: { color: disc } });
    text(s, [{ text: num }], { x: 0.575, y: y + 0.104, w: 0.938, h: 0.269, align: 'center', fontSize: 10, color: numColor });
    text(s, [{ text: label }], { x: 1.439, y: y + 0.054, w: 4.703, h: 0.37, fontSize: 16, valign: 'middle' });
  });
}

/* 13 — Digital transformation with donut */
function slide13(s) {
  logoMark(s);
  pageNumber(s, 13);
  heading(s, ['Digital ', 'Transformation in ', ['Consultations', C.cyan]],
    { x: 0.659, y: 1.681, w: 5.862, h: 1.919, fontSize: 36 });
  [['Diagnosis', 0.8, 1.11, 0.849, C.cyan, C.white], ['Treatment', 2.476, 2.768, 0.886, C.cyanPale, C.cyan]]
    .forEach(([label, x, tx, tw, fill, color]) => {
      s.addShape('roundRect', { x, y: 5.22, w: 1.469, h: 0.412, rectRadius: 0.206, fill: { color: fill } });
      text(s, [{ text: label }], { x: tx, y: 5.291, w: tw, h: 0.269, align: 'center', fontSize: 10, color });
    });
  body(s, BODY_2, { x: 0.69, y: 5.969, w: 5.247, h: 0.602 });
  groceryCard(s, 7.773, 2.726);
}

/* 14 — Satisfaction plus list cards */
function slide14(s) {
  logoMark(s);
  pageNumber(s, 14);
  heading(s, ['Enhancing ', ['Patient Satisfaction ', C.cyan], 'and Outcomes'],
    { x: 0.672, y: 1.764, w: 6.405, h: 1.717, fontSize: 32 });
  badgeTrio(s, 0.8, 4.976);
  body(s, BODY_3, { x: 0.689, y: 5.915, w: 6.373, h: 0.863 });
  [['01', 'Meaningful Impact', 4.34], ['02', 'Innovative Solutions', 5.325]].forEach(([num, label, y]) => {
    s.addShape('roundRect', { x: 7.474, y, w: 2.868, h: 0.736, rectRadius: 0.1, fill: { color: C.white }, shadow: cardShadow() });
    s.addShape('roundRect', { x: 7.639, y: y + 0.136, w: 0.464, h: 0.464, rectRadius: 0.232, fill: { color: C.cyan } });
    text(s, [{ text: num }], { x: 7.639, y: y + 0.234, w: 0.464, h: 0.269, align: 'center', fontSize: 10, fontFace: FONT_HEAD, color: C.white });
    text(s, [{ text: label }], { x: 8.24, y: y + 0.217, w: 2.086, h: 0.303, fontSize: 12 });
  });
}

/* 15 — Satisfaction, right-hand column */
function slide15(s) {
  glowSoft(s, 11.464, -1.391, 5.918);
  logoMark(s);
  pageNumber(s, 15);
  heading(s, ['Enhancing Patient Satisfaction ', ['and Outcomes', C.cyan]],
    { x: 0.666, y: 1.557, w: 6.844, h: 1.178, fontSize: 32 });
  text(s, [{ text: TAG_CARE }], { x: 0.685, y: 3.078, w: 5.311, h: 0.303, fontSize: 12, color: C.muted });
  badgeTrio(s, 9.78, 5.027);
  body(s, ['Medical consultations play a vital ', 'role in ensuring accurate diagnosis ', 'and effective treatment.'],
    { x: 9.667, y: 5.785, w: 5.247, h: 0.865 });
}

/* 16 — Satisfaction with a header tag */
function slide16(s) {
  logoMark(s);
  pageNumber(s, 16);
  heading(s, ['Enhancing ', ['Patient Satisfaction ', C.cyan], 'and Outcomes'],
    { x: 0.661, y: 1.838, w: 8.148, h: 1.919, fontSize: 36 });
  text(s, [{ text: TAG_CARE }], { x: 0.685, y: 4.106, w: 5.311, h: 0.303, fontSize: 12, color: C.muted });
  body(s, ['Medical consultations play a vital role in ensuring accurate ',
    'diagnosis, effective treatment, and ongoing continuous ', 'patient care.'],
    { x: 0.689, y: 5.8, w: 6.373, h: 0.863 });
  text(s, [{ text: 'Medical Consultation ' }, { text: 'Connects Patients', options: { color: C.cyan } }],
    { x: 6.752, y: 1.718, w: 5.311, h: 0.303, align: 'right', fontSize: 12 });
  iconBadge(s, 12.212, 1.709, 0.321, { glyph: 'microscope' });
}

/* 17 — Accessibility / price point */
function slide17(s) {
  pageNumber(s, 17);
  heading(s, ['Improving Accessibility ', 'to Medical Expertise'], { x: 6.547, y: 1.369, w: 11.973, h: 1.313, fontSize: 36 });
  text(s, [{ text: '$25.100' }], { x: 6.54, y: 3.569, w: 5.812, h: 0.841, fontSize: 44, fontFace: FONT_HEAD });
  text(s, [{ text: TAG_CARE }], { x: 6.553, y: 4.515, w: 5.579, h: 0.303, fontSize: 12, color: C.muted });
  [['Expert Consultation', 0.691, 0.706, 5.523], ['Health Guidance', 6.553, 6.562, 5.299]]
    .forEach(([label, lx, bxx, bw]) => {
      text(s, [{ text: label }], { x: lx, y: 5.664, w: 4.658, h: 0.337, fontSize: 14, fontFace: FONT_HEAD, color: C.cyan });
      body(s, ['The medical consultations play a vital role in ensuring',
        'accurate diagnosis and ongoing continuous patient care.'], { x: bxx, y: 6.073, w: bw, h: 0.601 });
    });
}

/* 18 — Data-driven decisions (monitor + donut) */
function slide18(s) {
  logoMark(s);
  pageNumber(s, 18);
  monitorMockup(s, 6.454, 2.666);
  heading(s, ['Data-Driven ', 'Medical Decision ', 'Making'], { x: 0.658, y: 1.843, w: 6.405, h: 1.919, fontSize: 36 });
  body(s, ['A medical consultations play a vital role in ensuring ', 'accurate diagnosis and ongoing continuous patient.'],
    { x: 0.689, y: 4.276, w: 6.373, h: 0.601 });
  numberedItem(s, 0.803, 6.291, 0.442, '01', 'Health Guidance');
  numberedItem(s, 3.465, 6.291, 0.442, '02', 'Accurate Diagnosis', { disc: C.cyanPale, numColor: C.cyan });
  groceryCard(s, 10.358, 1.368);
}

/** Flat-illustration desktop monitor showing an ECG trace. */
function monitorMockup(s, x, y) {
  s.addShape('rect', { x: x + 1.552, y: y + 3.004, w: 1.888, h: 0.699, fill: { color: C.grey15 } });
  s.addShape('roundRect', { x, y, w: 4.992, h: 3.027, rectRadius: 0.17, fill: { color: C.grey05 } });
  s.addShape('roundRect', { x: x + 1.243, y: y + 3.621, w: 2.505, h: 0.225, rectRadius: 0.11, fill: { color: C.grey05 } });
  s.addShape('rect', { x: x + 0.26, y: y + 0.232, w: 4.446, h: 2.502, fill: { color: C.white } });
  [0.482, 1.282].forEach(dx => s.addShape('roundRect', {
    x: x + dx, y: y + 0.359, w: 0.673, h: 0.156, rectRadius: 0.078, fill: { color: C.cyan },
  }));
  s.addShape('rect', { x: x + 0.384, y: y + 1.505, w: 4.185, h: 0.02, fill: { color: C.grey05 } });
  const ex = x + 0.384, ey = y + 1.084, ew = 4.113, eh = 0.908;
  const trace = [0.32, 0.31, 0.30, 0.86, 0.47, 0.72, 0.20, 0.66, 0.44, 0.65, 0.25, 0.64, 0.02, 0.86, 0.46, 0.52, 0.48];
  s.addShape('custGeom', {
    x: ex, y: ey, w: ew, h: eh, line: { color: C.cyan, width: 3 },
    points: trace.map((v, i) => ({ x: (i / (trace.length - 1)) * ew, y: (1 - v) * eh })),
  });
  for (let i = 0; i < 8; i++) s.addShape('roundRect', {
    x: x + 0.449 + i * 0.5735, y: y + 2.386, w: 0.066, h: 0.267, rectRadius: 0.033, fill: { color: C.cyan },
  });
}

/* 19 — Contact */
function slide19(s) {
  logoMark(s);
  pageNumber(s, 19);
  heading(s, ['Our Contact Information'], { x: 0.682, y: 4.953, w: 6.57, h: 0.707, fontSize: 36 });
  body(s, ['Medical consultations play a vital role in ensuring accurate diagnosis and ongoing continuous patient care.'],
    { x: 0.692, y: 6.055, w: 5.526, h: 0.601 });
  [['yourmail@addmail.com', 3.823, 'envelope'], ['www.yourwebsite.com', 4.837, 'globe'], ['+12 345 6789 390', 5.846, 'phone']]
    .forEach(([label, y, glyph]) => {
      s.addShape('roundRect', { x: 8.514, y, w: 3.336, h: 0.736, rectRadius: 0.1, fill: { color: C.white }, shadow: cardShadow() });
      iconBadge(s, 8.68, y + 0.136, 0.464, { glyph, round: false });
      text(s, [{ text: label }], { x: 9.28, y: y + 0.217, w: 2.571, h: 0.303, fontSize: 12 });
    });
}

/* 20 — Closing */
function slide20(s) {
  darkBackground(s);
  glowStrong(s, 10.988, 8.823, 5.403);
  logoMark(s, { color: C.white });
  text(s, [{ text: TAG_CARE }], { x: 7.332, y: 0.662, w: 5.311, h: 0.303, align: 'right', fontSize: 12, color: C.white });
  text(s, lines(['Thanks for Your', ['Attention ', C.cyan]]), {
    x: 0.691, y: 2.219, w: 12.5, h: 3.063, fontSize: 88, fontFace: FONT_HEAD, color: C.white, valign: 'top', margin: INSET,
  });
  text(s, [{ text: 'www.yourwebsite.com' }], { x: 0.691, y: 6.476, w: 5.311, h: 0.303, fontSize: 12, color: C.white });
}

/* ================================================================== *
 * 9. Build
 * ================================================================== */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK_16x9', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK_16x9';
  pptx.author = 'Healence Clinic';
  pptx.title = 'Medical Consultations';

  SLIDES.forEach(fn => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    fn(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '02328832-0106-40d4-b146-2b1cab68cfaf_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
