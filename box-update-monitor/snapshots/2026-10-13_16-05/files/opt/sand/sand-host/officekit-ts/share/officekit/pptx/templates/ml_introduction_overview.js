/**
 * "Machine Learning 1.0" — 20 slide deck rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5in.  Dark theme: near-black canvas that fades to navy
 * toward the bottom-right corner, wispy purple line-art down the right edge and
 * a periwinkle -> pink diagonal sweep used for headline words, pills and cards.
 *
 * Raster photos of the original deck are replaced by labelled placeholder boxes.
 *
 * Run:  node 0beef72e-8775-451f-8957-0e4dcb63593c_grok_final.js
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const CYAN = '83F2F5';
const PERI = '95A7FF';
const PINK = 'FF89C4';
const PINK_LT = 'FFB8DC';
const VIOLET = 'CF84F0';
const NAVY = '000E4C';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const CARD = '262626';   // charcoal card fill
const TRACK = 'F2F2F2';  // light progress-bar track

const HEAD = 'Montserrat';        // theme major font
const BODY = 'Work Sans';         // theme minor font
const NUMF = 'Urbanist ExtraBold'; // display numerals on the square stat cards

const SW = 13.3333333;   // 12192000 EMU
const SH = 7.5;

const LOREM = 'Machine Learning stands at the core of modern innovation, ' +
  'enabling systems to learn from data and make intelligent decisions.';
const LOREM_SHORT = 'Machine Learning stands at the core of modern innovation, ' +
  'enabling systems to learn .';
const LOREM_TINY = 'Machine Learning stands at the core of modern.';

/* --------------------------------------------------------------- utilities */

function hex2rgb(h) { return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
function rgb2hex(c) { return c.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).toUpperCase().padStart(2, '0')).join(''); }
function mix(a, b, t) {
  const x = hex2rgb(a), y = hex2rgb(b);
  return rgb2hex([0, 1, 2].map(i => x[i] + (y[i] - x[i]) * t));
}

/** Colour of the deck's signature sweep at position t (0 = top-left, 1 = bottom-right). */
function ramp(t) { return mix(PERI, PINK, Math.max(0, Math.min(1, t))); }
const RAMP_MID = ramp(0.5);

/**
 * Paint the deck's diagonal sweep across a (rounded) rectangle. pptxgenjs
 * cannot emit gradient fills, so the shape is drawn as a stack of thin slices
 * running along its long axis, each in the ramp colour of its own centre.
 * Slices near a rounded end are pulled in on the cross axis to follow the
 * corner arc, so the silhouette stays true; a full rounded rect in the middle
 * colour sits underneath to smooth the stair-steps.
 */
function sweep(s, o) {
  const vertical = !!o.vertical;
  const len = vertical ? o.h : o.w;
  const cross = vertical ? o.w : o.h;
  const r = Math.min(o.r || 0, Math.min(o.w, o.h) / 2);
  const shade = (lx, ly) => ramp((lx + ly) / (o.w + o.h)); // 45deg ramp in local coords

  panel(s, { x: o.x, y: o.y, w: o.w, h: o.h, r: r, fill: shade(o.w / 2, o.h / 2) });
  [[r, r], [o.w - r, r], [r, o.h - r], [o.w - r, o.h - r]].forEach(function (c) {
    if (r > 0) s.addShape('ellipse', {          // exact colour under each corner arc
      x: o.x + c[0] - r, y: o.y + c[1] - r, w: 2 * r, h: 2 * r,
      fill: { color: shade(c[0], c[1]) }, line: { type: 'none' }
    });
  });

  const n = o.n || Math.min(64, Math.max(6, Math.round(len / 0.06)));
  const step = len / n;
  for (let i = 0; i < n; i++) {
    const a = i * step;
    const edge = Math.max(0, Math.min(r, a, len - a - step)); // gap to the nearer end
    const inset = r - Math.sqrt(r * r - (r - edge) * (r - edge));
    const mid = a + step / 2;
    s.addShape('rect', Object.assign(
      vertical
        ? { x: o.x + inset, y: o.y + a, w: cross - 2 * inset, h: step }
        : { x: o.x + a, y: o.y + inset, w: step, h: cross - 2 * inset },
      {
        fill: { color: vertical ? shade(cross / 2, mid) : shade(mid, cross / 2) },
        line: { type: 'none' }
      }));
  }
}

/** Solid or swept rounded rectangle; fill === 'grad' selects the sweep. */
function panel(s, o) {
  if (o.fill === 'grad') sweep(s, o);
  else s.addShape(o.r > 0 ? 'roundRect' : 'rect',
    { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill }, line: { type: 'none' }, rectRadius: o.r || undefined });
}

/** Text box with the deck defaults (top aligned, white, body font). */
function txt(s, runs, o) {
  s.addText(runs, Object.assign({ fontFace: BODY, color: WHITE, valign: 'top' }, o));
}

/** Headline runs. 'g' marks the words the deck fills with the periwinkle sweep. */
function head(parts, size) {
  const out = [];
  parts.forEach(p => {
    const t = p[0], kind = p[1] || 'n', brk = p[2] === 'br';
    out.push({
      text: t,
      options: {
        bold: kind === 'g' || kind === 'b', color: kind === 'g' ? PERI : WHITE,
        fontSize: p[3] || size || 48, fontFace: HEAD, breakLine: brk
      }
    });
  });
  return out;
}

/* ------------------------------------------------------- shared decoration */

/** Near-black canvas fading to navy toward the bottom-right (45deg bands). */
function backdrop(s) {
  s.background = { color: BLACK };
  const bands = 26, diag = SW + SH;
  for (let i = 0; i < bands; i++) {
    const p = 0.53 + (0.47 * (i + 0.5)) / bands;         // fraction along the diagonal
    const cx = (p * diag) / 2;
    const thick = (0.47 * diag) / Math.SQRT2 / bands;
    s.addShape('rect', {
      x: cx - 16, y: cx - thick, w: 32, h: thick * 2, rotate: -45,
      fill: { color: mix(BLACK, NAVY, (p - 0.53) / 0.45) }, line: { type: 'none' }
    });
  }
}

/**
 * Bundle of hair-thin strands hugging the right edge. They rest just off the
 * canvas and get pulled left by three swells — a sharp one near the top, a
 * gentle one at mid height and a wide fan across the bottom-right corner —
 * pinching back to the edge in between. Each strand takes a fraction of the
 * pull, and the outer ones peak slightly earlier so the fan self-crosses.
 * One custom-geometry shape holds one open sub-path per strand.
 */
const STRAND_REST = 13.42;  // x the bundle relaxes to, just past the right edge

/** How far left the outermost strand is pulled at height y (f = strand index). */
function swell(y, f) {
  return 0.80 * Math.exp(-Math.pow(y / 0.55, 2)) +
         0.90 * Math.exp(-Math.pow((y - 3.35) / 0.90, 2)) +
         4.40 * Math.exp(-Math.pow((y - (8.3 - 0.5 * f)) / 1.65, 2));
}

function strandPoints(flip) {
  const count = 22, steps = 70, pts = [];
  for (let k = 0; k < count; k++) {
    const f = k / (count - 1);
    const pull = 0.04 + 0.96 * Math.pow(f, 1.2);
    for (let i = 0; i <= steps; i++) {
      const y = (SH * i) / steps;
      const waist = 0.35 * Math.exp(-Math.pow((y - 5.0) / 1.2, 2)); // node the bundle pinches through
      const x = STRAND_REST - waist - pull * swell(y, f);
      // the left-hand copy is the same artwork turned through 180 degrees
      const p = flip ? { x: SW - x, y: SH - y } : { x: x, y: y };
      if (i === 0) p.moveTo = true;
      pts.push(p);
    }
  }
  return pts;
}

function strands(s, flip) {
  s.addShape('custGeom', {
    x: 0, y: 0, w: SW, h: SH, points: strandPoints(flip),
    line: { color: '6D4E7E', width: 0.55 }
  });
}

/**
 * Wordmark in the top-left corner of every slide: four overlapping discs
 * stepping through the ramp (the first bitten by a quarter notch), a tiny
 * satellite dot, the descender of the "g" and the "by logoipsum" byline.
 */
function logo(s) {
  const d = 0.207, y = 0.478;
  for (let i = 0; i < 4; i++) {
    s.addShape('ellipse', {
      x: 0.806 + i * 0.2057, y: y, w: d, h: d,
      fill: { color: ramp(0.1 + i * 0.26) }, line: { type: 'none' }
    });
  }
  s.addShape('rect', { x: 0.806 + d / 2, y: y, w: d / 2, h: d / 2, fill: { color: BLACK }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 1.256, y: y + 0.216, w: 0.108, h: 0.055, rectRadius: 0.027, fill: { color: ramp(0.62) }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 1.612, y: y + 0.008, w: 0.036, h: 0.036, fill: { color: ramp(0.9) }, line: { type: 'none' } });
  txt(s, 'by logoipsum', { x: 1.378, y: y + 0.212, w: 0.5, h: 0.09, margin: 0, fontSize: 3, bold: true, color: PINK_LT });
}

/**
 * Every slide: canvas, line-art, photo stand-ins, wordmark, running head and
 * page number. Photos go in early so the corner furniture stays on top of them,
 * exactly as the transparent PNGs behave in the original deck.
 */
function frame(s, num, opts) {
  const o = opts || {};
  backdrop(s);
  strands(s, false);
  if (o.mirrorStrands) strands(s, true);
  (o.photos || []).forEach(function (ph) { imagePlaceholder(s, ph); });
  logo(s);
  const tone = o.headTint || WHITE;
  txt(s, [
    { text: 'Machine Learning  ', options: { fontSize: 11 } },
    { text: '1.0 ', options: { fontSize: 11, bold: true } }
  ], { x: 10.169, y: 0.473, w: 2.364, h: 0.286, align: 'right', valign: 'middle', fontFace: HEAD, color: tone });
  txt(s, String(num), { x: 9.533, y: 6.951, w: 3.0, h: 0.399, align: 'right', fontSize: 14, color: WHITE });
}

/* ------------------------------------------------------- reusable elements */

/** "ML 1.0 - Introduction" eyebrow. */
function kicker(s, x, y, color) {
  txt(s, [
    { text: 'ML 1.0 - ', options: { bold: true, fontSize: 14 } },
    { text: 'Introduction', options: { fontSize: 12 } }
  ], { x: x, y: y, w: 2.364, h: 0.337, fontFace: HEAD, color: color || WHITE });
}

/** Gradient "Learn More" pill with its outlined chevron button. */
function learnMore(s, x, y, ink) {
  const fg = ink || WHITE;
  sweep(s, { x: x, y: y, w: 1.804, h: 0.471, r: 0.2355 });
  txt(s, 'Learn More', { x: x + 0.172, y: y + 0.094, w: 1.595, h: 0.303, bold: true, fontSize: 12, color: fg, lineSpacingMultiple: 1 });
  s.addShape('ellipse', { x: x + 1.394, y: y + 0.144, w: 0.196, h: 0.196, fill: { type: 'none' }, line: { color: fg, width: 1.25 } });
  chevron(s, x + 1.472, y + 0.242, 0.05, fg);
}

/** Small ">" built from two bars, centred on (cx, cy). */
function chevron(s, cx, cy, r, color) {
  [45, -45].forEach(function (rot, i) {
    s.addShape('rect', {
      x: cx - r / 2, y: cy + (i === 0 ? -r / 2 : r / 2) - 0.007, w: r, h: 0.014,
      rotate: rot, fill: { color: color }, line: { type: 'none' }
    });
  });
}

/** Rounded chip: 24pt value over a 12pt caption (1.997 x 1.052). */
function chip(s, o) {
  const w = o.w || 1.997, h = o.h || 1.052;
  panel(s, { x: o.x, y: o.y, w: w, h: h, r: 0.2646, fill: o.fill });
  txt(s, o.value, { x: o.x + 0.218, y: o.y + 0.155, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, o.caption, { x: o.x + 0.218, y: o.y + 0.553, w: o.capW || 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
}

/** Centred badge: 1.93 x 0.995 plate with a 2.21 wide text column. */
function badge(s, o) {
  panel(s, { x: o.x + 0.14, y: o.y, w: 1.93, h: 0.995, r: 0.2436, fill: o.fill });
  txt(s, o.value, { x: o.x, y: o.y + 0.132, w: 2.21, h: 0.505, align: 'center', fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, o.caption, { x: o.x, y: o.y + 0.561, w: 2.21, h: 0.286, align: 'center', fontSize: 10.5 });
}

/** Square stat tile with an oversized numeral (slides 10 and 15). */
function tile(s, o) {
  const side = o.side || 1.659;
  panel(s, { x: o.x, y: o.y, w: side, h: side, r: o.r === undefined ? 0.102 : o.r, fill: o.fill });
  txt(s, o.value, { x: o.x + 0.179, y: o.y + 0.347, w: 1.268, h: 0.707, align: 'center', fontFace: NUMF, fontSize: 36, color: o.valueColor || WHITE });
  txt(s, o.caption, { x: o.x, y: o.y + 0.899, w: 1.63, h: 0.336, align: 'center', fontSize: 10.5, lineSpacingMultiple: 1.5 });
}

/** Feature card: title, blurb, progress rail and a small cyan glyph. */
function featureCard(s, o) {
  panel(s, { x: o.x, y: o.y, w: 2.867, h: 1.749, r: 0.245, fill: o.fill });
  txt(s, o.title, { x: o.x + 0.26, y: o.y + 0.199, w: 2.5, h: 0.37, fontFace: HEAD, fontSize: 16 });
  txt(s, o.body || LOREM_TINY, { x: o.x + 0.26, y: o.y + 0.541, w: 2.373, h: 0.667, fontSize: 12, lineSpacingMultiple: 1.5 });
  rail(s, o.x + 0.674, o.y + 1.428, 1.785, o.progress);
  s.addShape('roundRect', { x: o.x + 0.353, y: o.y + 1.341, w: 0.173, h: 0.173, rectRadius: 0.045, fill: { color: CYAN }, line: { type: 'none' } });
}

/** Rounded 8pt rail with a coloured portion on top. */
function rail(s, x, yMid, len, filled) {
  const t = 0.111;
  s.addShape('roundRect', { x: x - t / 2, y: yMid - t / 2, w: len + t, h: t, rectRadius: t / 2, fill: { color: TRACK }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x - t / 2, y: yMid - t / 2, w: filled + t, h: t, rectRadius: t / 2, fill: { color: CYAN }, line: { type: 'none' } });
}

/** Circle carrying a small megaphone glyph — the deck's accent icon. */
function iconDot(s, o) {
  panel(s, { x: o.x, y: o.y, w: o.d, h: o.d, r: o.d / 2, fill: o.fill });
  megaphone(s, o.x + o.d * 0.28, o.y + o.d * 0.3, o.d * 0.44, o.glyph);
}

/**
 * Megaphone pointing right, drawn inside a d x d box: a rounded body, a horn
 * flaring to the right, a stub handle below and three sound dashes.
 * Parts are placed by centre (rotation in PowerPoint pivots about the centre).
 */
function megaphone(s, x, y, d, color) {
  const put = (shape, cu, cv, w, h, rot) => s.addShape(shape, {
    x: x + (cu - w / 2) * d, y: y + (cv - h / 2) * d, w: w * d, h: h * d,
    rotate: rot || 0, rectRadius: Math.min(w, h) * d * 0.35,
    fill: { color: color }, line: { type: 'none' }
  });
  put('roundRect', 0.24, 0.72, 0.15, 0.30, 20);   // handle
  put('roundRect', 0.16, 0.45, 0.32, 0.30);       // body
  put('trapezoid', 0.38, 0.45, 0.80, 0.44, 270);  // flared horn
  put('roundRect', 0.58, 0.45, 0.09, 0.86);       // horn lip
  put('roundRect', 0.80, 0.22, 0.18, 0.09, -30);  // sound dashes
  put('roundRect', 0.84, 0.45, 0.18, 0.09);
  put('roundRect', 0.80, 0.68, 0.18, 0.09, 30);
}

/** Stand-in for a photo the original deck placed here (deliberately quiet). */
function imagePlaceholder(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.14,
    fill: { color: '0B0B12' }, line: { color: '242433', width: 0.75, dashType: 'dash' }
  });
  txt(s, o.label || '[image]', { x: o.x, y: o.y + o.h / 2 - 0.18, w: o.w, h: 0.36, align: 'center', fontSize: 11, color: '55555F' });
}

/** Body copy defaults: 12pt, 1.5 line spacing. */
function copy(s, x, y, w, h, text, o) {
  txt(s, text, Object.assign({ x: x, y: y, w: w, h: h, fontSize: 12, lineSpacingMultiple: 1.5 }, o || {}));
}

/** Bold 12pt Montserrat sub-heading. */
function subhead(s, x, y, w, text) {
  txt(s, text, { x: x, y: y, w: w, h: 0.303, fontFace: HEAD, bold: true, fontSize: 12 });
}

/* ------------------------------------------------------------ slide bodies */

const slides = [];

slides.push(function (s) { // 1 — title
  frame(s, 1, { mirrorStrands: true, headTint: CYAN });
  kicker(s, 1.085, 2.24, CYAN);
  txt(s, 'MACHINE', { x: 1.0, y: 2.53, w: 8.483, h: 1.447, fontFace: HEAD, bold: true, fontSize: 80, color: PERI });
  txt(s, 'LEARNING', { x: 0.994, y: 3.733, w: 10.3, h: 1.717, fontFace: HEAD, fontSize: 96, color: WHITE });
  learnMore(s, 1.149, 5.623, BLACK);
  copy(s, 3.297, 5.479, 6.0, 0.673, LOREM);
});

slides.push(function (s) { // 2 — understanding machine learning
  frame(s, 2);
  kicker(s, 1.07, 1.766);
  txt(s, head([['Understanding'], [' ', 'b'], ['Machine', 'g', '', 54], [' ', 'b', '', 54], ['Learning', 'g', '', 54]]),
    { x: 0.986, y: 2.179, w: 7.85, h: 1.818 });
  subhead(s, 1.07, 4.799, 4.086, 'From Data to Decisions');
  copy(s, 1.07, 5.229, 3.206, 1.279, LOREM);
  txt(s, '+80%', { x: 9.321, y: 4.299, w: 2.668, h: 0.707, align: 'center', fontFace: HEAD, bold: true, fontSize: 36, color: PINK });
  copy(s, 9.69, 4.891, 1.93, 0.675, 'Machine Learning stands', { align: 'center', fontFace: HEAD });
  learnMore(s, 9.755, 6.031);
});

slides.push(function (s) { // 3 — fueling intelligent systems
  frame(s, 3);
  kicker(s, 8.334, 1.761);
  txt(s, head([['Fueling', 'n', 'br'], ['Intelligent', 'g', 'br'], ['Systems']]),
    { x: 8.334, y: 2.179, w: 7.85, h: 2.524 });
  subhead(s, 8.334, 4.954, 4.086, 'From Data to Decisions');
  copy(s, 8.334, 5.384, 4.18, 0.976, LOREM);
  subhead(s, 1.07, 2.403, 4.086, 'Analytics and Modeling');
  copy(s, 1.07, 2.832, 2.78, 1.582, LOREM, { align: 'justify' });
  chip(s, { x: 1.07, y: 5.384, fill: NAVY, value: '$3,1M', caption: 'Machine Learning' });
  chip(s, { x: 3.251, y: 5.384, fill: 'grad', value: '20%', caption: 'Neural Network' });
});

slides.push(function (s) { // 4 — types of machine learning
  frame(s, 4);
  kicker(s, 4.938, 1.811);
  txt(s, head([['Types of', 'n', 'br'], ['Machine Learning', 'g']]),
    { x: 4.853, y: 2.224, w: 10.562, h: 1.717 });
  copy(s, 10.81, 1.224, 1.745, 0.867, 'Deep Learning Model Development Process',
    { align: 'right', fontSize: 10.5 });
  featureCard(s, { x: 2.785, y: 4.412, fill: CARD, title: 'Model Training', progress: 0.738 });
  featureCard(s, { x: 6.165, y: 4.412, fill: NAVY, title: 'Artificial Intelligence', progress: 1.336 });
  featureCard(s, { x: 9.545, y: 4.412, fill: 'grad', title: 'Learning Process', progress: 0.931 });
});

slides.push(function (s) { // 5 — learning from labeled data
  frame(s, 5);
  kicker(s, 1.07, 3.127);
  txt(s, head([['Learning', 'g'], [' from', 'n', 'br'], ['Labeled '], ['Data', 'g']]),
    { x: 0.986, y: 3.54, w: 7.85, h: 1.717 });
  subhead(s, 1.07, 5.481, 8.895, 'From Data to Decisions');
  copy(s, 1.07, 5.911, 5.011, 0.673, LOREM_SHORT);
  chip(s, { x: 2.489, y: 1.295, fill: NAVY, value: '+22%', caption: 'Intelligent' });
  chip(s, { x: 4.67, y: 1.295, fill: 'grad', value: '+30%', caption: 'Algorithm' });
  badge(s, { x: 7.251, y: 4.267, fill: NAVY, value: '$99K', caption: 'Data Analysis' });
  copy(s, 9.562, 4.406, 2.873, 0.673, LOREM_TINY);
  badge(s, { x: 7.251, y: 5.476, fill: 'grad', value: '$1M', caption: 'Learning Process' });
  copy(s, 9.562, 5.615, 2.873, 0.673, LOREM_TINY);
});

slides.push(function (s) { // 6 — unsupervised learning insights
  frame(s, 6);
  kicker(s, 7.66, 1.494);
  txt(s, head([['Unsupervised'], [' Learning', 'g', 'br'], ['Insights']]),
    { x: 7.575, y: 1.907, w: 7.85, h: 2.524 });
  [
    { y: 5.379, label: 'Artificial Intelligence Framework', pct: '80%', w: 3.044, color: PINK_LT, knob: 10.517 },
    { y: 6.392, label: 'Automated Learning Process', pct: '50%', w: 2.286, color: CYAN, knob: 9.58 }
  ].forEach(function (b) {
    txt(s, b.label, { x: 7.66, y: b.y - 0.515, w: 3.7, h: 0.303, bold: true, fontSize: 12 });
    txt(s, b.pct, { x: 10.22, y: b.y - 0.515, w: 2.286, h: 0.303, align: 'right', fontSize: 12 });
    s.addShape('roundRect', { x: 7.711, y: b.y, w: 4.882, h: 0.188, rectRadius: 0.094, fill: { color: TRACK }, line: { type: 'none' } });
    s.addShape('roundRect', { x: 7.711, y: b.y, w: b.w, h: 0.188, rectRadius: 0.094, fill: { color: b.color }, line: { type: 'none' } });
    s.addShape('ellipse', { x: b.knob, y: b.y - 0.123, w: 0.417, h: 0.435, fill: { color: WHITE }, line: { type: 'none' } });
    s.addShape('ellipse', { x: b.knob + 0.075, y: b.y - 0.046, w: 0.268, h: 0.28, fill: { color: b.color }, line: { type: 'none' } });
  });
  txt(s, '+70%', { x: 1.178, y: 1.774, w: 2.668, h: 0.707, align: 'center', fontFace: HEAD, bold: true, fontSize: 36, color: PINK });
  copy(s, 1.547, 2.367, 1.93, 0.675, 'Machine Learning Model', { align: 'center', fontFace: HEAD });
  chip(s, { x: 3.628, y: 5.742, fill: 'grad', value: '$1M', caption: 'Learning Process' });
});

slides.push(function (s) { // 7 — reinforcement learning basics
  frame(s, 7);
  kicker(s, 1.07, 4.481);
  txt(s, head([['Reinforcement'], [' Learning Basics', 'g']]),
    { x: 0.986, y: 4.895, w: 7.85, h: 1.717 });
  [
    { y: 5.112, fill: 'grad', n: '1', nFill: BLACK, label: 'Automated Insights', ink: BLACK },
    { y: 5.82, fill: NAVY, n: '2', nFill: 'grad', label: 'Intelligent Algorithm', ink: WHITE }
  ].forEach(function (r) {
    panel(s, { x: 8.325, y: r.y, w: 3.621, h: 0.557, r: 0.195, fill: r.fill });
    panel(s, { x: 8.461, y: r.y + 0.111, w: 0.346, h: 0.334, r: 0.122, fill: r.nFill });
    txt(s, r.n, { x: 8.461, y: r.y + 0.111, w: 0.346, h: 0.334, align: 'center', valign: 'middle', bold: true, fontSize: 10.5 });
    txt(s, r.label, { x: 8.886, y: r.y + 0.093, w: 3.5, h: 0.337, valign: 'middle', fontFace: HEAD, bold: true, fontSize: 14, color: r.ink });
  });
  [4.272, 9.08].forEach(function (x) {
    iconDot(s, { x: x, y: 2.234, d: 0.788, fill: BLACK, glyph: VIOLET });
  });
});

slides.push(function (s) { // 8 — healthcare innovations
  frame(s, 8);
  kicker(s, 0.914, 4.592);
  txt(s, head([['Healthcare '], ['Innovations', 'g']]),
    { x: 0.914, y: 4.896, w: 5.361, h: 1.717 });
  copy(s, 0.914, 1.371, 1.724, 0.867, 'Deep Learning Model Development Process', { fontSize: 10.5 });
  chip(s, { x: 3.473, y: 2.238, fill: 'grad', value: '122%', caption: 'Machine Learning' });
  [
    { y: 1.852, fill: 'grad', glyph: WHITE, value: '+70%', caption: 'Learning Process' },
    { y: 2.974, fill: NAVY, glyph: RAMP_MID, value: '+20%', caption: 'Data Analysis' }
  ].forEach(function (r) {
    iconDot(s, { x: 10.088, y: r.y + 0.047, d: 0.607, fill: r.fill, glyph: r.glyph });
    txt(s, r.value, { x: 10.846, y: r.y, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
    txt(s, r.caption, { x: 10.846, y: r.y + 0.398, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  });
  iconDot(s, { x: 7.486, y: 5.221, d: 0.788, fill: BLACK, glyph: VIOLET });
});

slides.push(function (s) { // 9 — smarter decisions, better outcomes
  frame(s, 9);
  kicker(s, 4.773, 4.414);
  txt(s, head([['Smarter Decisions, '], ['Better Outcomes', 'g']]),
    { x: 4.688, y: 4.828, w: 7.85, h: 1.717 });
  sweep(s, { x: 0.853, y: 1.41, w: 5.041, h: 0.557, r: 0.195 });
  panel(s, { x: 0.989, y: 1.522, w: 0.346, h: 0.334, r: 0.122, fill: NAVY });
  txt(s, '1', { x: 0.989, y: 1.522, w: 0.346, h: 0.334, align: 'center', valign: 'middle', bold: true, fontSize: 10.5 });
  txt(s, 'Artificial Intelligence Framework', { x: 1.414, y: 1.504, w: 4.046, h: 0.337, valign: 'middle', fontFace: HEAD, bold: true, fontSize: 14 });
  badge(s, { x: 0.71, y: 2.4, fill: NAVY, value: '$99K', caption: 'Data Analysis' });
  copy(s, 3.021, 2.539, 2.873, 0.673, LOREM_TINY);
});

slides.push(function (s) { // 10 — finance and risk management
  frame(s, 10);
  kicker(s, 0.909, 1.282);
  txt(s, head([['Finance and Risk '], ['Management', 'g']]),
    { x: 0.825, y: 1.696, w: 7.85, h: 1.717 });
  txt(s, 'Learning Model', { x: 0.968, y: 4.192, w: 2.935, h: 0.438, fontFace: HEAD, bold: true, fontSize: 20 });
  ['Machine Learning Development Process', 'Predictive Analytics and Modeling', 'Deep Learning Neural Network']
    .forEach(function (t, i) {
      const y = 4.768 + i * 0.491;
      s.addShape('ellipse', { x: 1.076, y: y + 0.153, w: 0.12, h: 0.12, fill: { color: CYAN }, line: { type: 'none' } });
      copy(s, 1.282, y, 3.614, 0.376, t);
    });
  tile(s, { x: 5.329, y: 4.468, fill: 'grad', value: '$7M', caption: 'Learning Process' });
  tile(s, { x: 7.251, y: 4.468, fill: NAVY, value: '$3B', caption: 'Automated Insights', valueColor: CYAN });
});

slides.push(function (s) { // 11 — engaging customers intelligently
  frame(s, 11);
  kicker(s, 0.909, 1.673);
  txt(s, head([['Engaging '], ['Customers', 'g'], [' Intelligently']]),
    { x: 0.825, y: 2.087, w: 5.163, h: 2.524 });
  featureCard(s, { x: 6.176, y: 1.312, fill: NAVY, title: 'Neural Network', progress: 1.336 });
  featureCard(s, { x: 9.557, y: 1.312, fill: 'grad', title: 'Intelligent', progress: 0.931 });
  featureCard(s, { x: 6.176, y: 3.332, fill: CARD, title: 'Algorithm', progress: 1.336 });
  featureCard(s, { x: 9.557, y: 3.332, fill: CARD, title: 'Computing', progress: 0.931 });
  txt(s, '$3,1M', { x: 0.873, y: 4.966, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24, color: RAMP_MID });
  txt(s, 'AI Systems', { x: 0.873, y: 5.438, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  txt(s, '$1,1M', { x: 2.903, y: 4.966, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, 'Model Training', { x: 2.903, y: 5.438, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  learnMore(s, 0.909, 6.059);
});

slides.push(function (s) { // 12 — natural language processing
  frame(s, 12);
  kicker(s, 0.909, 1.45);
  txt(s, head([['Natural '], ['Language', 'g'], [' Processing']]),
    { x: 0.825, y: 1.863, w: 7.692, h: 1.717 });
  subhead(s, 0.909, 3.768, 3.649, 'From Data to Decisions');
  // two upright pills (the original rotates a 2.488 x 2.019 rounded rect by 90deg)
  sweep(s, { x: 0.909, y: 4.393, w: 2.019, h: 2.488, r: 1.0095, vertical: true });
  panel(s, { x: 3.273, y: 4.393, w: 2.019, h: 2.488, r: 1.0095, fill: CARD });
  txt(s, '+20%', { x: 1.196, y: 5.104, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, 'Machine Learning Model', { x: 1.196, y: 5.584, w: 1.289, h: 0.707, fontFace: HEAD, fontSize: 12 });
  txt(s, '+10%', { x: 3.559, y: 5.104, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, 'Predictive Data Analysis', { x: 3.559, y: 5.584, w: 1.289, h: 0.707, fontFace: HEAD, fontSize: 12 });
  iconDot(s, { x: 8.161, y: 5.684, d: 0.607, fill: 'grad', glyph: WHITE });
  txt(s, '+30%', { x: 9.336, y: 5.637, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24, color: PINK });
  txt(s, 'Automated Insights', { x: 9.336, y: 6.035, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  iconDot(s, { x: 10.862, y: 3.431, d: 0.788, fill: NAVY, glyph: RAMP_MID });
});

slides.push(function (s) { // 13 — computer vision
  frame(s, 13);
  kicker(s, 9.061, 1.62);
  txt(s, head([['Computer '], ['Vision', 'g']]),
    { x: 8.977, y: 2.033, w: 3.859, h: 1.717 });
  subhead(s, 9.061, 4.088, 4.637, 'From Data to Decisions');
  copy(s, 9.061, 4.517, 3.356, 0.976, LOREM_SHORT);
  learnMore(s, 9.102, 5.88);
  [
    { x: 0.718, value: '+80%', color: PINK, caption: 'Machine Learning stands' },
    { x: 3.31, value: '+30%', color: CYAN, caption: 'Automated Learning Process' },
    { x: 5.903, value: '+70%', color: VIOLET, caption: 'Artificial Intelligence Framework' }
  ].forEach(function (c) {
    txt(s, c.value, { x: c.x, y: 5.464, w: 2.668, h: 0.707, align: 'center', fontFace: HEAD, bold: true, fontSize: 36, color: c.color });
    copy(s, c.x + 0.369, 6.056, 1.93, 0.675, c.caption, { align: 'center', fontFace: HEAD });
  });
});

slides.push(function (s) { // 14 — recommendation systems (bar chart)
  frame(s, 14);
  kicker(s, 0.909, 1.255);
  txt(s, head([['Recommendation', 'g'], [' Systems']]),
    { x: 0.825, y: 1.669, w: 7.692, h: 1.717 });
  copy(s, 10.544, 1.423, 1.728, 0.867, 'Deep Learning Model Development Process',
    { align: 'right', fontSize: 10.5 });
  s.addChart('bar', [
    { name: 'Column1', labels: ['Week 1', 'Week 2', 'Week 3'], values: [2, 2.5, 3.5] },
    { name: 'Column2', labels: ['Week 1', 'Week 2', 'Week 3'], values: [3, 4.4, 4] }
  ], {
    x: 3.519, y: 2.899, w: 9.242, h: 4.01,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 117, barOverlapPct: -26,
    chartColors: [VIOLET, PINK],
    showLegend: false, showTitle: false, showValue: true,
    dataLabelColor: WHITE, dataLabelFontFace: 'Be Vietnam Pro Light', dataLabelFontSize: 9,
    dataLabelFormatCode: 'General',
    catAxisLabelColor: WHITE, catAxisLabelFontFace: 'Be Vietnam Pro Light', catAxisLabelFontSize: 9,
    catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' }
  });
  txt(s, '$13,1K', { x: 0.873, y: 3.72, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24, color: PERI });
  txt(s, 'Deep Neural Network', { x: 0.873, y: 4.192, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  txt(s, '$23,1K', { x: 0.909, y: 4.863, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, 'AI Decision System', { x: 0.909, y: 5.335, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  learnMore(s, 0.909, 6.05);
});

slides.push(function (s) { // 15 — predictive analytics (calendar)
  frame(s, 15);
  kicker(s, 0.956, 5.434);
  txt(s, head([['Predictive '], ['Analytics', 'g']]),
    { x: 0.871, y: 5.809, w: 8.713, h: 0.909 });
  learnMore(s, 10.381, 6.05);
  panel(s, { x: 0.956, y: 1.43, w: 9.414, h: 3.679, r: 0.213, fill: CARD });
  sweep(s, { x: 1.176, y: 1.654, w: 9.002, h: 0.445, r: 0.159 });
  const colX = [1.42, 2.77, 4.13, 5.51, 6.93, 8.26, 9.62];
  ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(function (d, i) {
    txt(s, d, { x: colX[i], y: 1.725, w: 0.4, h: 0.303, valign: 'middle', fontFace: HEAD, bold: true, fontSize: 12, color: BLACK });
  });
  // highlighted ranges 12->14 and 21->23
  [{ x: 2.694, y: 3.338, w: 3.19, dots: [2.765, 5.516] }, { x: 5.455, y: 3.89, w: 3.175, dots: [5.516, 8.269] }]
    .forEach(function (h) {
      s.addShape('roundRect', { x: h.x, y: h.y, w: h.w, h: 0.431, rectRadius: 0.2155, fill: { color: TRACK }, line: { type: 'none' } });
      h.dots.forEach(function (dx) {
        s.addShape('ellipse', { x: dx, y: h.y + 0.065, w: 0.3, h: 0.3, fill: { color: CYAN }, line: { type: 'none' } });
      });
    });
  const rowY = [2.309, 2.862, 3.414, 3.967, 4.52];
  for (let d = 1; d <= 31; d++) {
    const cell = d + 4;                       // the 1st falls on a Thursday
    const r = Math.floor((cell - 1) / 7), c = (cell - 1) % 7;
    const dark = d === 12 || d === 14 || d === 21 || d === 23;
    txt(s, String(d), {
      x: colX[c] - 0.012, y: rowY[r], w: 0.4, h: 0.278, valign: 'middle',
      fontSize: dark ? 10 : 10.5, bold: dark || d === 1, color: dark ? BLACK : WHITE
    });
  }
  tile(s, { x: 10.526, y: 1.43, side: 1.774, r: 0.234, fill: 'grad', value: '$7M', caption: 'Learning Process' });
  tile(s, { x: 10.526, y: 3.452, side: 1.659, r: 0.234, fill: NAVY, value: '$3B', caption: 'Automated Insights', valueColor: CYAN });
});

slides.push(function (s) { // 16 — edge machine learning
  frame(s, 16, { photos: [{ x: 6.9, y: 0.23, w: 5.1, h: 7.27, label: '[image] phone' }] });
  kicker(s, 1.07, 1.766);
  txt(s, head([['Edge '], ['Machine', 'g'], [' Learning']]),
    { x: 0.986, y: 2.179, w: 7.85, h: 1.717 });
  learnMore(s, 1.07, 4.187);
  subhead(s, 1.07, 5.481, 8.895, 'From Data to Decisions');
  copy(s, 1.07, 5.911, 5.011, 0.673, LOREM_SHORT);
  sweep(s, { x: 5.867 + (2.488 - 2.019) / 2, y: 3.71 - (2.488 - 2.019) / 2, w: 2.019, h: 2.488, r: 1.0095, vertical: true });
  txt(s, '90%', { x: 6.388, y: 4.187, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24 });
  txt(s, 'Machine Learning Model', { x: 6.388, y: 4.667, w: 1.289, h: 0.707, fontFace: HEAD, fontSize: 12 });
});

slides.push(function (s) { // 17 — the role of cloud platforms (laptop)
  frame(s, 17, { photos: [{ x: 5.88, y: 0.87, w: 6.6, h: 5.4, label: '[image] laptop' }] });
  kicker(s, 0.746, 2.509);
  txt(s, head([['The '], ['Role', 'g'], [' of', 'n', 'br'], ['Cloud', 'g'], [' '], ['Platforms', 'g']]),
    { x: 0.746, y: 2.923, w: 6.321, h: 1.717 });
  subhead(s, 0.746, 4.847, 4.637, 'From Data to Decisions');
  copy(s, 0.746, 5.277, 4.451, 0.673, LOREM_SHORT);
  featureCard(s, { x: 5.583, y: 1.247, fill: NAVY, title: 'Neural Network', progress: 0.931 });
  iconDot(s, { x: 11.805, y: 3.228, d: 0.788, fill: 'grad', glyph: WHITE });
});

slides.push(function (s) { // 18 — the role of cloud platforms (phones)
  frame(s, 18, { photos: [{ x: 8.67, y: 0.0, w: 3.95, h: 5.99, label: '[image] phones' }] });
  kicker(s, 0.746, 1.566);
  txt(s, head([['The '], ['Role', 'g'], [' of', 'n', 'br'], ['Cloud', 'g'], [' '], ['Platforms', 'g']]),
    { x: 0.746, y: 1.98, w: 6.321, h: 1.717 });
  subhead(s, 4.683, 4.147, 4.637, 'From Data to Decisions');
  copy(s, 4.683, 4.577, 4.451, 0.673, LOREM_SHORT);
  learnMore(s, 7.844, 2.166);
  [
    { x: 4.677, fill: NAVY, glyph: RAMP_MID, tx: 5.435, value: '+44%', color: PINK, caption: 'Neural Network' },
    { x: 7.545, fill: 'grad', glyph: WHITE, tx: 8.303, value: '+10%', color: CYAN, caption: 'Automated Learning' }
  ].forEach(function (r) {
    iconDot(s, { x: r.x, y: 5.806, d: 0.607, fill: r.fill, glyph: r.glyph });
    txt(s, r.value, { x: r.tx, y: 5.759, w: 1.561, h: 0.505, fontFace: HEAD, bold: true, fontSize: 24, color: r.color });
    txt(s, r.caption, { x: r.tx, y: 6.157, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  });
});

slides.push(function (s) { // 19 — contact
  frame(s, 19);
  kicker(s, 0.746, 1.566);
  txt(s, head([['Connect', 'n', 'br'], ['with us at:', 'g']]),
    { x: 0.746, y: 1.98, w: 6.321, h: 1.717 });
  txt(s, [
    { text: '"The greatest learning machines are still human minds that never stop exploring.\u201c', options: { fontSize: 14, breakLine: true } },
    { text: '~Unkown~', options: { fontSize: 14, bold: true, italic: true } }
  ], { x: 7.446, y: 1.566, w: 4.831, h: 1.122, align: 'right', lineSpacingMultiple: 1.5 });
  [
    { x: 0.746, w: 2.255, fill: 'grad', title: 'E-Mail', value: 'hello@yourmail.com' },
    { x: 3.176, w: 2.152, fill: NAVY, title: 'Phone', value: '+123-456-7890' },
    { x: 5.484, w: 2.152, fill: CARD, title: 'Website', value: 'yourwebsite.com' },
    { x: 7.791, w: 2.152, fill: CARD, title: 'Location', value: '123 Anywhere St.' }
  ].forEach(function (c) {
    panel(s, { x: c.x, y: 4.2, w: c.w, h: 1.052, r: 0.2646, fill: c.fill });
    txt(s, c.title, { x: c.x + 0.177, y: 4.356, w: 1.561, h: 0.303, fontFace: HEAD, bold: true, fontSize: 12 });
    txt(s, c.value, { x: c.x + 0.177, y: 4.792, w: 2.364, h: 0.303, fontFace: HEAD, fontSize: 12 });
  });
  learnMore(s, 0.746, 6.05);
});

slides.push(function (s) { // 20 — thank you
  frame(s, 20);
  kicker(s, 0.837, 1.62);
  txt(s, head([['Thank', 'g', '', 88], [' You!', 'n', '', 88]]),
    { x: 0.837, y: 2.033, w: 6.666, h: 3.063 });
  copy(s, 0.837, 5.772, 5.011, 0.673,
    'Thank you for exploring the world of Machine Learning \u2014 where data transforms into discovery.');
  learnMore(s, 10.146, 6.05);
});

/* ---------------------------------------------------------------- assemble */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: SW, height: SH });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Machine Learning 1.0';

slides.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '0beef72e-8775-451f-8957-0e4dcb63593c_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
