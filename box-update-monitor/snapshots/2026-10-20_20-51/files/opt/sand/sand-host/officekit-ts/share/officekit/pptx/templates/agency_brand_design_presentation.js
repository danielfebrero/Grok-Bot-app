#!/usr/bin/env node
/**
 * "Whisper Agency Design" — brand-guideline deck rebuilt with pptxgenjs.
 * 30 slides, 13.333 x 7.5 in.  Run: node <this file>
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const C = {
  bg: '0D0D0D', // slide master background (near-black)
  white: 'FFFFFF',
  black: '000000',
  green: '003D29', // accent1
  tan: 'BF8040', // accent2
  sand: 'E4DBB4', // accent3
  cream: 'FFFBE5', // accent4
  red: 'FF2200', // accent5
  ink: '001A11', // accent6
  tanDark: '8F6030', // accent2 @ lum 75%
  tanDeep: '604020', // accent2 @ lum 50%
  tanLite: 'AC733A', // accent2 @ lum 90%
  grey95: 'F2F2F2',
  grey85: 'D9D9D9',
  grey75: 'BFBFBF',
  grey65: 'A6A6A6',
  grey30: '4C4C4C',
};

const HEAD = 'Roboto'; // theme major font
const BODY = 'Inter'; // theme minor font
const POPPINS = 'Poppins'; // used by a few caption blocks

/* Lorem blocks that repeat throughout the deck. */
const LOREM = {
  variations:
    "There are many variations of passages of Lorem Ipsum available, but the majority have suffered " +
    "alteration in some form, by injected humor, or randomized words which don't look even slightly " +
    'believable. If you are',
  variationsUse:
    "There are many variations of passages of Lorem Ipsum available, but the majority have suffered " +
    "alteration in some form, by injected humor, or randomized words which don't look even slightly " +
    'believable. If you are going to use',
  ipsumShort:
    'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum has ' +
    'been type and survived not only five centuries, but also passages, and Lorem Ipsum.',
  dolor:
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies',
  sed:
    'PLACEHOLDER' +
    'laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi ' +
    'architecto beatae vitae dicta sunt explicabo. ',
  caseTail:
    "\u00a0is simply dummy text of the printing and typesetting industry. lorem ipsum has been the " +
    "industry's standard dummy text ever since the 1500s",
};

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
const pptx = new PptxGenJS();
const S = pptx.ShapeType;

/** Blend two hex colours (t = 0 → a, 1 → b). Used for low-contrast tints. */
function mix(a, b, t) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map((i) => Math.round(ch(a, i) * (1 - t) + ch(b, i) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/** Solid rectangle. */
function rect(sl, x, y, w, h, fill, opts) {
  sl.addShape(S.rect, Object.assign({ x, y, w, h, fill: { color: fill } }, opts));
}

/** Rounded rectangle; `r` is the corner radius as a fraction of the short side. */
function roundRect(sl, x, y, w, h, r, opts) {
  sl.addShape(S.roundRect, Object.assign({ x, y, w, h, rectRadius: r * Math.min(w, h) }, opts));
}

/** Circle / ellipse. */
function oval(sl, x, y, w, h, opts) {
  sl.addShape(S.ellipse, Object.assign({ x, y, w, h }, opts));
}

/** Horizontal rule. */
function hline(sl, x, y, w, color, width, transparency) {
  sl.addShape(S.line, { x, y, w, h: 0, line: { color, width: width || 0.75, transparency } });
}

/** Vertical rule. */
function vline(sl, x, y, h, color, width, transparency) {
  sl.addShape(S.line, { x, y, w: 0, h, line: { color, width: width || 0.75, transparency } });
}

/** Text box. Defaults match the source deck: Inter, top-anchored, left-aligned. */
function txt(sl, text, o) {
  sl.addText(text, Object.assign({ fontFace: BODY, isTextBox: true, valign: 'top' }, o));
}

/** Turn an array of strings into stacked paragraphs for addText(). */
function lines(arr, opts) {
  return arr.map((t) => ({ text: t, options: Object.assign({ breakLine: true }, opts) }));
}

/** A stroke of an icon: a bar of length `len`, centred on (cx, cy), rotated. */
function bar(sl, cx, cy, len, thick, angle, color) {
  sl.addShape(S.rect, {
    x: cx - len / 2, y: cy - thick / 2, w: len, h: thick, rotate: angle, fill: { color },
  });
}

/**
 * Tiny pictograms built from native shapes (the reference uses freeform icon
 * art). `s` is the nominal icon size in inches, (cx, cy) its centre.
 */
function icon(sl, kind, cx, cy, s, color) {
  const fill = { color };
  switch (kind) {
    case 'rss': // broadcast waves
      oval(sl, cx - s * 0.42, cy - s * 0.42, s * 0.84, s * 0.84, { fill: { type: 'none' }, line: { color, width: 1.25 } });
      oval(sl, cx - s * 0.2, cy - s * 0.2, s * 0.4, s * 0.4, { fill: { type: 'none' }, line: { color, width: 1.25 } });
      oval(sl, cx - s * 0.07, cy - s * 0.07, s * 0.14, s * 0.14, { fill });
      break;
    case 'cloud':
      sl.addShape(S.cloud, { x: cx - s * 0.5, y: cy - s * 0.36, w: s, h: s * 0.72, fill });
      break;
    case 'asterisk':
      sl.addShape(S.star8, { x: cx - s * 0.5, y: cy - s * 0.5, w: s, h: s, fill });
      break;
    case 'check':
      bar(sl, cx - s * 0.19, cy + s * 0.13, s * 0.38, s * 0.11, 45, color);
      bar(sl, cx + s * 0.11, cy - s * 0.02, s * 0.72, s * 0.11, -48, color);
      break;
    case 'cart':
      sl.addShape(S.trapezoid, { x: cx - s * 0.36, y: cy - s * 0.22, w: s * 0.72, h: s * 0.4, flipV: true, fill });
      bar(sl, cx - s * 0.42, cy - s * 0.3, s * 0.3, s * 0.08, -35, color);
      oval(sl, cx - s * 0.28, cy + s * 0.24, s * 0.16, s * 0.16, { fill });
      oval(sl, cx + s * 0.12, cy + s * 0.24, s * 0.16, s * 0.16, { fill });
      break;
    case 'bars':
      [[-0.34, 0.3], [-0.06, 0.5], [0.22, 0.74]].forEach(([dx, h]) => {
        sl.addShape(S.rect, { x: cx + dx * s, y: cy + s * 0.38 - h * s, w: s * 0.2, h: h * s, fill });
      });
      break;
    case 'plant': // sprout held in a palm
      sl.addShape(S.chord, { x: cx - s * 0.5, y: cy - s * 0.05, w: s, h: s * 0.55, rotate: 180, fill });
      bar(sl, cx, cy - s * 0.22, s * 0.34, s * 0.07, 90, color);
      oval(sl, cx - s * 0.34, cy - s * 0.46, s * 0.3, s * 0.2, { rotate: 20, fill });
      oval(sl, cx + s * 0.04, cy - s * 0.46, s * 0.3, s * 0.2, { rotate: -20, fill });
      break;
    case 'bulb':
      oval(sl, cx - s * 0.3, cy - s * 0.46, s * 0.6, s * 0.6, { fill });
      sl.addShape(S.rect, { x: cx - s * 0.16, y: cy + s * 0.12, w: s * 0.32, h: s * 0.1, fill });
      sl.addShape(S.rect, { x: cx - s * 0.12, y: cy + s * 0.28, w: s * 0.24, h: s * 0.09, fill });
      break;
    case 'bird': // stylised paper-plane / bird mark
      sl.addShape(S.triangle, { x: cx - s * 0.5, y: cy - s * 0.4, w: s, h: s * 0.8, rotate: 105, fill });
      break;
  }
}

/**
 * Stand-in for a photo / picture placeholder: a soft outlined panel with an
 * "[image]" caption. No raster data is embedded anywhere in this deck.
 */
function photo(sl, x, y, w, h, o) {
  o = o || {};
  const base = o.over || C.bg; // colour the placeholder sits on
  const edge = o.edge || mix(base, o.onDark === false ? C.black : C.white, 0.13);
  sl.addShape(S.rect, {
    x, y, w, h,
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: { color: edge, width: 0.75 },
  });
  txt(sl, o.label || '[image]', {
    x, y: o.labelTop ? y + 0.06 : y + h / 2 - 0.16, w, h: 0.32,
    align: 'center', valign: 'middle', fontSize: 10, color: o.labelColor || edge,
  });
}

/* ------------------------------------------------------------------ *
 * Slide builders
 * ------------------------------------------------------------------ */

// 1 — Title
function slide01(sl) {
  rect(sl, 5.189, 2.198, 7.758, 3.104, C.tan);
  rect(sl, 5.575, 2.198, 7.758, 3.104, C.sand);
  txt(sl, lines(['Whisper', 'Agency Design']), {
    x: 6.029, y: 2.887, w: 5.551, h: 1.41,
    fontFace: HEAD, fontSize: 54, color: C.tanDark, lineSpacingMultiple: 0.7,
  });
  txt(sl, 'Agency Brand Design Presentation', {
    x: 6.029, y: 4.439, w: 4.731, h: 0.325, fontSize: 14, color: C.tanDark, lineSpacing: 16,
  });
}

// 2 — Welcome message
function slide02(sl) {
  sl.background = { color: C.tan };
  txt(sl, 'Welcome Message Details', {
    x: 1.892, y: 1.988, w: 4.523, h: 1.447, fontFace: HEAD, fontSize: 40, color: C.white,
  });
  txt(sl, LOREM.variations, {
    x: 1.902, y: 4.111, w: 3.38, h: 1.427, fontSize: 12, color: C.white, lineSpacing: 16,
  });
  txt(sl, 'Your Name', {
    x: 7.387, y: 5.168, w: 1.893, h: 0.37,
    fontFace: HEAD, fontSize: 16, color: C.white, charSpacing: 1,
  });
  txt(sl, 'Position', { x: 7.387, y: 5.519, w: 1.986, h: 0.303, fontSize: 12, color: C.white });
  // three social marks
  icon(sl, 'bird', 7.583, 6.077, 0.2, C.white);
  [['f', 7.895], ['G', 8.249]].forEach(([g, x]) => {
    roundRect(sl, x, 5.99, 0.16, 0.16, 0.18, { fill: { color: C.white } });
    txt(sl, g, {
      x, y: 5.99, w: 0.16, h: 0.16, align: 'center', valign: 'middle',
      fontSize: 8, bold: true, color: C.tan, margin: 0,
    });
  });
}

// 3 — Table of content
function slide03(sl) {
  rect(sl, 0, 2.627, 8.851, 3.104, C.tan);
  rect(sl, 8.013, 2.168, 3.98, 3.792, C.sand);
  const toc = ['The Brand', 'Logo', 'Colour Palette', 'Typography', 'Gallery', 'Application', 'Contact'];
  txt(sl, lines(toc), {
    x: 8.561, y: 3.151, w: 1.638, h: 2.196,
    fontFace: HEAD, fontSize: 10, color: C.tan, lineSpacingMultiple: 1.8,
  });
  txt(sl, lines(['10', '11', '12', '13', '14', '15', '16']), {
    x: 10.8, y: 3.151, w: 0.706, h: 2.196, align: 'right',
    fontFace: HEAD, fontSize: 10, color: C.tan, lineSpacingMultiple: 1.8,
  });
  for (let i = 0; i < 6; i++) hline(sl, 8.632, 3.509 + i * 0.3035, 2.808, mix(C.sand, C.tan, 0.15), 1);
  txt(sl, 'Agency Table of Content', {
    x: 0.892, y: 3.237, w: 5.089, h: 0.614, fontFace: HEAD, fontSize: 32, color: C.white,
  });
  txt(sl, LOREM.variationsUse, {
    x: 0.892, y: 4.472, w: 6.528, h: 0.754,
    fontSize: 12, color: C.white, transparency: 30, lineSpacing: 16,
  });
}

// 4 — Section divider "01 / About Us"
function slide04(sl) {
  sl.addShape(S.rect, { x: 6.667, y: 3.75, w: 6.667, h: 3.75, fill: { color: C.tan, transparency: 10 } });
  txt(sl, '01', {
    x: 7.341, y: 1.323, w: 3.296, h: 1.935,
    fontFace: HEAD, fontSize: 115, bold: true, color: C.sand, margin: 0,
  });
  txt(sl, 'SECTION ONE', {
    x: 10.078, y: 5.598, w: 2.433, h: 0.236, align: 'right',
    fontFace: HEAD, fontSize: 14, color: C.white, margin: 0,
  });
  txt(sl, 'About Us', {
    x: 10.078, y: 5.994, w: 2.433, h: 0.673, align: 'right',
    fontFace: HEAD, fontSize: 40, bold: true, color: C.white, margin: 0,
  });
}

// 5 — Brand overview
function slide05(sl) {
  photo(sl, 7.775, 0, 5.53, 7.5);
  hline(sl, 0, 6.986, 5.735, C.white, 0.75, 65);
  txt(sl, 'BRAND [CLIENT NAME]', {
    x: 0.403, y: 7.129, w: 3.722, h: 0.269, fontSize: 10, color: C.white,
  });
  rect(sl, 0, 3.047, 7.804, 3.395, C.sand);
  txt(sl, lines(['Brand ', 'Overview']), {
    x: 0.776, y: 3.522, w: 5.089, h: 1.557, fontFace: HEAD, fontSize: 44, color: C.tan,
  });
  txt(sl, LOREM.variationsUse, {
    x: 0.776, y: 5.126, w: 6.528, h: 0.983, fontSize: 12, color: C.black, lineSpacing: 16,
  });
}

// 6 — History cards
function slide06(sl) {
  sl.background = { color: C.tan };
  const cards = [
    { x: 2.143, tx: 2.331, bx: 2.313, bw: 1.952, cx: 2.421, year: '2021', body: 'A wonderful serenity has taken possession of my entire soul.', glyph: 'rss', hi: false },
    { x: 4.476, tx: 4.664, bx: 4.688, bw: 1.798, cx: 4.758, year: '2022', body: 'A wonderful serenity has taken of my entire soul.', glyph: 'cloud', hi: true },
    { x: 6.81, tx: 6.976, bx: 7.004, bw: 1.952, cx: 7.072, year: '2023', body: 'A wonderful serenity possession of my entire soul.', glyph: 'asterisk', hi: false },
    { x: 9.143, tx: 9.331, bx: 9.316, bw: 1.952, cx: 9.454, year: '2024', body: 'A wonderful serenity entire soul.', glyph: 'check', hi: false },
  ];
  cards.forEach((c) => {
    const ink = c.hi ? C.tan : C.white;
    roundRect(sl, c.x, 2.991, 2.208, 3.249, 0.0597,
      c.hi ? { fill: { color: C.sand } } : { fill: { type: 'none' }, line: { color: C.white, width: 0.5 } });
    txt(sl, c.year, {
      x: c.tx, y: 3.259, w: 1.833, h: 0.74, fontFace: HEAD, fontSize: 32, bold: true, color: ink,
    });
    txt(sl, c.body, {
      x: c.bx, y: 4.05, w: c.bw, h: 0.867, fontSize: 12, color: ink, lineSpacingMultiple: 1.3,
    });
    oval(sl, c.cx, 5.128, 0.744, 0.742, { fill: { color: c.hi ? C.tan : C.sand } });
    icon(sl, c.glyph, c.cx + 0.372, 5.499, 0.36, c.hi ? C.white : C.tan);
  });
  txt(sl, 'History', {
    x: 2.886, y: 1.596, w: 7.561, h: 0.828, align: 'center',
    fontFace: HEAD, fontSize: 48, bold: true, color: C.white,
    charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
}

// 7 — Keywords
function slide07(sl) {
  txt(sl, 'Keywords', {
    x: 1.148, y: 1.03, w: 4.125, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.tan,
  });
  txt(sl, LOREM.ipsumShort, {
    x: 1.148, y: 2.005, w: 4.525, h: 1.151, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  rect(sl, 6.667, 2.713, 2.968, 3.308, C.tan);
  txt(sl, LOREM.ipsumShort, {
    x: 7.116, y: 3.428, w: 1.992, h: 2.202, fontSize: 10.67, color: C.white, lineSpacingMultiple: 1.3,
  });
  const cols = [
    { x: 1.148, y: 3.967, h: 2.74, words: ['Creative', 'Awesome', 'Clean', 'Stylish', 'Minimal', 'Exclusive', 'Original'] },
    { x: 2.714, y: 4.344, h: 2.363, words: ['Classic', 'Understated', 'Authentic', 'Subtle', 'Elegant', 'Structured'] },
    { x: 4.316, y: 5.475, h: 1.232, words: ['Subtle', 'Elegant', 'Structured'] },
  ];
  cols.forEach((c) => txt(sl, lines(c.words), {
    x: c.x, y: c.y, w: 1.774, h: c.h, fontSize: 14, color: C.white, lineSpacingMultiple: 1.6,
  }));
}

// 8 — Moodboard
function slide08(sl) {
  rect(sl, 5.007, -0.023, 8.326, 7.523, C.sand);
  txt(sl, 'Moodboard', {
    x: 0.662, y: 3.628, w: 4.125, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.tan,
  });
  txt(sl, LOREM.ipsumShort, {
    x: 0.662, y: 4.773, w: 3.667, h: 1.414, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  photo(sl, 5.75, 1.876, 3.289, 4.188, { over: C.sand, onDark: false });
  photo(sl, 9.48, 1.902, 2.962, 1.848, { over: C.sand, onDark: false });
  photo(sl, 9.48, 4.191, 2.962, 1.874, { over: C.sand, onDark: false });
}

// 9 — Main concept
function slide09(sl) {
  rect(sl, -0.01, -0.011, 7.142, 7.511, C.tan);
  txt(sl, 'Main Concept', {
    x: 0.819, y: 1.136, w: 4.125, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.white,
  });
  txt(sl, LOREM.ipsumShort, {
    x: 0.819, y: 2.096, w: 3.667, h: 1.414, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  // the one real photograph in the deck → flat colour stand-in
  photo(sl, 0.819, 4.203, 2.092, 2.056, { fill: '8C99AB', edge: '8C99AB', labelColor: '3C4655' });
  const pills = [
    { px: 4.349, py: 4.865, tx: 4.429, tw: 1.181, label: 'Lorem' },
    { px: 3.41, py: 5.329, tx: 3.468, tw: 1.181, label: 'Dolor' },
    { px: 4.934, py: 5.329, tx: 5.02, tw: 1.181, label: 'Ipsum' },
    { px: 4.06, py: 5.792, tx: 4.06, tw: 1.347, label: 'Consecte' },
  ];
  pills.forEach((p) => {
    roundRect(sl, p.px, p.py, 1.347, 0.329, 0.5, { fill: { color: C.white } });
    txt(sl, p.label, {
      x: p.tx, y: p.py + 0.014, w: p.tw, h: 0.292, align: 'center',
      fontFace: HEAD, fontSize: 10, color: C.tan, lineSpacing: 15,
    });
  });
  photo(sl, 7.757, 0.728, 4.935, 2.737);
  [['Case A', 3.867, 4.374], ['Case B', 5.696, 6.203]].forEach(([label, ty, by]) => {
    txt(sl, label, {
      x: 7.749, y: ty, w: 2.333, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.tan,
    });
    txt(sl, [
      { text: 'Lorem ipsum', options: { bold: true } },
      { text: LOREM.caseTail },
    ], { x: 7.749, y: by, w: 4.94, h: 0.715, align: 'justify', fontSize: 12, color: C.white, lineSpacing: 15 });
  });
  hline(sl, 7.116, 5.327, 6.217, C.green, 0.75);
}

// 10 — Brand persona (venn)
function slide10(sl) {
  sl.background = { color: C.sand };
  const circles = [
    { x: 7.611, y: 1.23, label: 'Elegant' },
    { x: 9.473, y: 1.23, label: 'Charm' },
    { x: 6.732, y: 2.746, label: 'Social' },
    { x: 8.594, y: 2.746, label: 'Professional' },
    { x: 10.456, y: 2.746, label: 'Equality' },
  ];
  circles.forEach((c) => {
    oval(sl, c.x, c.y, 2.245, 2.245, { fill: { type: 'none' }, line: { color: C.white, width: 0.75 } });
    txt(sl, c.label, {
      x: c.x, y: c.y, w: 2.245, h: 2.245, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 14, bold: true, color: C.tan,
    });
  });
  txt(sl, lines(['Brand', 'Persona']), {
    x: 1.148, y: 1.794, w: 3.81, h: 1.555, fontFace: HEAD, fontSize: 48, bold: true,
    color: C.tan, charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
  txt(sl, LOREM.dolor + ' Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies', {
    x: 1.154, y: 4.031, w: 3.724, h: 2.328, fontSize: 12, color: C.white, lineSpacingMultiple: 1.4,
  });
  txt(sl, LOREM.dolor, {
    x: 6.488, y: 5.38, w: 5.799, h: 0.598, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// 11 — Tone and voice
function slide11(sl) {
  rect(sl, -0.01, -0.011, 6.676, 7.511, C.tan);
  txt(sl, lines(['Tone and', 'Voice']), {
    x: 0.819, y: 1.136, w: 4.125, h: 1.447, fontFace: HEAD, fontSize: 40, color: C.white,
  });
  txt(sl, 'Your Tone and Voice Here', {
    x: 0.817, y: 3.056, w: 3.228, h: 0.303,
    fontSize: 12, bold: true, italic: true, color: C.white, charSpacing: 1,
  });
  txt(sl, LOREM.ipsumShort + ' ' + LOREM.ipsumShort, {
    x: 0.811, y: 3.548, w: 4.819, h: 2.201, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  [0.868, 2.373, 3.838].forEach((x) => {
    roundRect(sl, x, 6.026, 1.347, 0.329, 0.5, { fill: { color: C.white } });
    txt(sl, 'Dolor', {
      x: x + 0.059, y: 6.044, w: 1.181, h: 0.292, align: 'center',
      fontFace: HEAD, fontSize: 10, color: C.tan, lineSpacing: 15,
    });
  });
  [1.233, 3.066, 4.899].forEach((y, i) => {
    photo(sl, 7.581, y, 1.562, 1.333);
    hline(sl, 9.529, [1.336, 3.253, 5.196][i], 0.356, C.grey85, 1);
    txt(sl, [
      { text: 'Lorem ipsum' },
      { text: '\u00a0is simply dummy text of the printing and' },
    ], { x: 9.411, y: [1.421, 3.359, 5.297][i], w: 2.078, h: 0.72, fontSize: 12, color: C.white, lineSpacing: 15 });
  });
}

// 12 — Colour palette
function slide12(sl) {
  const swatches = [
    { x: 0.828, fill: C.tan, code: 'ECE8E3' },
    { x: 3.2, fill: C.sand, code: '858C80' },
    { x: 5.572, fill: C.cream, code: 'B0B5AC' },
    { x: 7.944, fill: C.red, code: '404040' },
    { x: 10.315, fill: C.ink, code: 'E7E6E6' },
  ];
  swatches.forEach((s) => {
    rect(sl, s.x, 3.645, 2.182, 3.08, s.fill);
    txt(sl, 'RGB', { x: s.x + 0.13, y: 3.856, w: 0.981, h: 0.281, fontSize: 10.67, color: C.white });
    txt(sl, 'Color Code Text', { x: s.x + 0.146, y: 4.142, w: 1.363, h: 0.281, fontSize: 10.67, color: C.white });
    txt(sl, s.code, { x: s.x + 0.089, y: 6.258, w: 1.067, h: 0.281, fontSize: 10.67, color: C.white });
  });
  txt(sl, 'Color Palette', {
    x: 0.736, y: 0.842, w: 4.125, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.tan,
  });
  txt(sl, LOREM.ipsumShort, {
    x: 0.736, y: 1.816, w: 5.211, h: 0.889, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  txt(sl, 'Vibrant Theme Color', {
    x: 6.55, y: 1.131, w: 3.228, h: 0.438, fontFace: HEAD, fontSize: 20, color: C.tan, charSpacing: 1,
  });
  txt(sl, LOREM.ipsumShort +
    ' Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type', {
    x: 6.55, y: 1.722, w: 5.948, h: 1.17, align: 'justify',
    fontSize: 10, color: C.white, lineSpacingMultiple: 1.3,
  });
  hline(sl, 0.828, 3.175, 11.67, C.green, 0.75);
}

// 13 — Typography
function slide13(sl) {
  const specimens = [
    { x: 3.878, y: 5.102, h: 1.447, size: 80, alpha: 64 },
    { x: 5.925, y: 4.621, h: 2.121, size: 120, alpha: 37 },
    { x: 8.732, y: 3.891, h: 3.13, size: 180, alpha: 0 },
  ];
  specimens.forEach((s) => txt(sl, 'Aa', {
    x: s.x, y: s.y, w: 3.631, h: s.h, align: 'center',
    fontFace: HEAD, fontSize: s.size, color: C.tan, transparency: s.alpha,
  }));
  txt(sl, lines(['Used', 'Typography']), {
    x: 1.237, y: 1.255, w: 3.81, h: 1.434, fontFace: HEAD, fontSize: 44,
    color: C.tan, charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
  hline(sl, 1.343, 3.594, 10.579, C.tan, 0.75);
  txt(sl, 'PLACEHOLDER', {
    x: 5.604, y: 1.255, w: 6.544, h: 0.847, fontFace: HEAD, fontSize: 18,
    color: C.white, lineSpacingMultiple: 1.3,
  });
  txt(sl, '0123456789', {
    x: 5.604, y: 2.354, w: 3.481, h: 0.462, fontFace: HEAD, fontSize: 18,
    color: C.white, lineSpacingMultiple: 1.3,
  });
  txt(sl, 'Tilt Warp', {
    x: 1.343, y: 4.203, w: 2.646, h: 0.438, fontSize: 20, bold: true, color: C.white,
  });
  txt(sl, lines(['Regular. Medium', 'Semi-bold. Bold']), {
    x: 1.343, y: 5.474, w: 2.056, h: 0.769, fontSize: 15, color: C.white, lineSpacing: 25,
  });
}

// 14 — Logo clear space
function slide14(sl) {
  sl.background = { color: C.tan };
  hline(sl, 2.527, 4.615, 8.11, C.white, 0.75, 48);
  rect(sl, 5.469, 4.409, 1.965, 0.465, C.tan);
  // bold construction lines
  hline(sl, 2.395, 2.385, 8.11, C.white, 0.75);
  hline(sl, 2.395, 3.526, 8.11, C.white, 0.75);
  vline(sl, 2.951, 2.078, 1.817, C.white, 0.75);
  vline(sl, 10.034, 2.003, 1.978, C.white, 0.75);
  // faint construction lines
  hline(sl, 2.73, 2.684, 7.497, C.white, 0.75, 58);
  hline(sl, 2.73, 3.268, 7.497, C.white, 0.75, 58);
  [[3.258, 2.078, 1.817], [9.741, 2.003, 1.978], [5.201, 2.078, 1.817], [7.725, 2.003, 1.892]]
    .forEach(([x, y, h]) => vline(sl, x, y, h, C.white, 0.75, 58));
  txt(sl, 'Your Brand Logo', {
    x: 2.901, y: 2.558, w: 7.164, h: 0.919, align: 'center',
    fontFace: HEAD, fontSize: 54, bold: true, color: C.white,
    charSpacing: 3, lineSpacingMultiple: 0.9,
  });
  txt(sl, 'Minimum clear space', {
    x: 5.041, y: 4.436, w: 2.816, h: 0.336, align: 'center',
    fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 1.3,
  });
  vline(sl, 11.022, 2.385, 1.141, C.white, 0.75, 37);
  txt(sl, lines(['X', 'height']), {
    x: 11.219, y: 2.65, w: 0.817, h: 0.598, align: 'center',
    fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 1.3,
  });
  hline(sl, 5.32, 1.587, 2.392, C.white, 0.75, 48);
  rect(sl, 6.109, 1.423, 0.815, 0.465, C.tan);
  txt(sl, '3x', {
    x: 6.206, y: 1.37, w: 0.62, h: 0.336, align: 'center',
    fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 1.3,
  });
  txt(sl, LOREM.dolor + 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ', {
    x: 1.598, y: 5.574, w: 10.029, h: 0.598, align: 'center',
    fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// 15 — Market trends (doughnut)
function slide15(sl) {
  rect(sl, 0, 0, 5.619, 7.5, C.sand);
  photo(sl, 0.952, 1.377, 3.714, 4.746, { over: C.sand, onDark: false });
  txt(sl, 'Market Trends', {
    x: 6.696, y: 1.161, w: 4.272, h: 0.535, fontFace: HEAD, fontSize: 40, color: C.tan,
  });
  txt(sl, 'There variations are many', {
    x: 6.696, y: 2.431, w: 3.304, h: 0.39, fontFace: HEAD, fontSize: 16, color: C.tan, lineSpacing: 22,
  });
  txt(sl, LOREM.variations, {
    x: 6.696, y: 2.861, w: 4.891, h: 0.999,
    fontFace: POPPINS, fontSize: 10, color: C.white, transparency: 30, lineSpacing: 16,
  });
  sl.addChart(pptx.ChartType.doughnut,
    [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] }],
    {
      x: 6.792, y: 4.595, w: 2.283, h: 2.079,
      holeSize: 75, showLegend: false, showValue: false, showPercent: true,
      chartColors: [C.tanLite, C.tanDeep, C.tanDark, C.tanDark],
      dataBorder: { pt: 0.75, color: C.tan },
      dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontFace: BODY,
      chartArea: { fill: { type: 'none' } },
    });
  oval(sl, 7.666, 5.259, 0.471, 0.465, { fill: { color: C.black } });
  txt(sl, 'Trend Use', {
    x: 7.088, y: 5.732, w: 1.69, h: 0.269, align: 'center', fontSize: 10, color: C.white,
  });
  txt(sl, 'Most All', {
    x: 9.926, y: 4.967, w: 1.888, h: 0.39, fontFace: HEAD, fontSize: 20, color: C.tan,
  });
  txt(sl, '8.106M', {
    x: 9.926, y: 5.44, w: 2.944, h: 0.861, fontFace: HEAD, fontSize: 45, color: C.tan,
  });
}

// 16 — Market review
function slide16(sl) {
  sl.background = { color: C.tan };
  txt(sl, 'Market Review', {
    x: 4.531, y: 1.217, w: 4.272, h: 0.535, align: 'center',
    fontFace: HEAD, fontSize: 40, color: C.white,
  });
  sl.addShape(S.rect, {
    x: 0.778, y: 2.741, w: 11.778, h: 1.194, fill: { type: 'none' }, line: { color: C.white, width: 0.75 },
  });
  [3.722, 6.667, 9.611].forEach((x) => vline(sl, x, 2.741, 1.194, C.white, 1));
  const cells = [
    { cx: 1.738, glyph: 'cart', title: 'Market ', tx: 0.949, bx: 0.622 },
    { cx: 4.818, glyph: 'bars', title: 'Needs', tx: 4.04, bx: 3.713 },
    { cx: 7.909, glyph: 'plant', title: 'Risks', tx: 7.13, bx: 6.803 },
    { cx: 11.003, glyph: 'bulb', title: 'Solution', tx: 10.221, bx: 9.894 },
  ];
  cells.forEach((c) => {
    oval(sl, c.cx, 3.035, 0.6, 0.6, { fill: { color: C.sand } });
    icon(sl, c.glyph, c.cx + 0.3, 3.335, 0.33, C.black);
    txt(sl, c.title, {
      x: c.tx, y: 4.648, w: 2.161, h: 0.438, align: 'center',
      fontFace: HEAD, fontSize: 20, color: C.white, charSpacing: 1,
    });
    txt(sl, LOREM.variations.slice(0, 148), {
      x: c.bx, y: 5.207, w: 2.815, h: 1.223, align: 'center',
      fontSize: 10, color: C.white, lineSpacing: 16,
    });
  });
}

// 17 — Market analysis (stacked bar)
function slide17(sl) {
  txt(sl, lines(['Market ', 'Analysis']), {
    x: 1.357, y: 1.454, w: 3.872, h: 1.569, fontFace: HEAD, fontSize: 44, color: C.tan,
    lineSpacingMultiple: 1.0,
  });
  txt(sl, 'There variations are many', {
    x: 1.357, y: 4.77, w: 3.304, h: 0.39, fontFace: HEAD, fontSize: 16, color: C.tan, lineSpacing: 22,
  });
  txt(sl, LOREM.variations, {
    x: 1.357, y: 5.327, w: 3.872, h: 1.209,
    fontFace: POPPINS, fontSize: 10, color: C.white, transparency: 30, lineSpacing: 16,
  });
  const stats = [
    { unit: 'Million', value: '10.42', uy: 2.183, vy: 2.634 },
    { unit: 'Billion', value: '20.5', uy: 3.82, vy: 4.271 },
    { unit: 'Million', value: '22.11', uy: 5.489, vy: 5.939 },
  ];
  stats.forEach((s) => {
    txt(sl, s.unit, { x: 6.667, y: s.uy, w: 1.25, h: 0.333, fontSize: 18, color: C.white });
    txt(sl, s.value, { x: 6.667, y: s.vy, w: 1.7, h: 0.632, fontFace: HEAD, fontSize: 35, color: C.tan });
  });
  [3.461, 5.128].forEach((y) => hline(sl, 6.769, y, 1.924, C.black, 0.5));
  sl.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: ['Category 1'], values: [0.04] },
    { name: 'Series 2', labels: ['Category 1'], values: [0.06] },
  ], {
    x: 8.928, y: 0.686, w: 3.676, h: 6.269,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 400, barOverlapPct: 100,
    chartColors: [C.tanDeep, C.tan],
    showValue: true, dataLabelFormatCode: '0%', dataLabelColor: C.white, dataLabelFontSize: 12,
    showLegend: true, legendPos: 'b', legendColor: C.white, legendFontSize: 12,
    catAxisLabelColor: C.white, catAxisLabelFontSize: 12, catAxisLineShow: true,
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    chartArea: { fill: { type: 'none' } },
  });
}

// 18 — Target growth (four small column charts)
function slide18(sl) {
  txt(sl, 'Target Growth', {
    x: 4.012, y: 0.922, w: 5.309, h: 0.785, align: 'center',
    fontFace: HEAD, fontSize: 44, color: C.tan,
  });
  [1.003, 3.826, 6.649, 9.471].forEach((x) => {
    sl.addChart(pptx.ChartType.bar,
      [{ name: 'Series 1', labels: ['Category 1', 'Category 2'], values: [4, 2] }],
      {
        x, y: 2.456, w: 2.82, h: 3.365,
        barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 100, barOverlapPct: -24,
        chartColors: [C.tan, C.sand],
        showValue: false, showLegend: false,
        catAxisLabelColor: C.white, catAxisLabelFontSize: 12,
        valAxisLabelColor: C.white, valAxisLabelFontSize: 12,
        valGridLine: { color: 'E7E6E6', size: 0.75 }, catGridLine: { style: 'none' },
        valAxisMaxVal: 4.5, valAxisMajorUnit: 0.5,
        chartArea: { fill: { type: 'none' } },
      });
  });
  txt(sl, [
    { text: 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humor, or There' },
    { text: ' are many variations of passages of', options: { transparency: 30 } },
  ], {
    x: 2.711, y: 6.224, w: 7.911, h: 0.526, align: 'center',
    fontSize: 11, color: C.white, lineSpacing: 16,
  });
}

// 19 — Market solution
function slide19(sl) {
  rect(sl, 4.584, 1.442, 8.201, 0.663, C.sand);
  rect(sl, 4.598, 2.391, 8.201, 4.062, C.tan);
  [[4.988, 1.235, 'Solution 01'], [6.496, 1.771, 'Solution 02'],
   [8.725, 1.771, 'Solution 03'], [10.892, 1.771, 'Solution 04']].forEach(([x, w, label]) => {
    txt(sl, label, {
      x, y: 1.652, w, h: 0.303, align: 'center', fontSize: 12, bold: true, color: C.white,
    });
  });
  [[5.369, 3.105], [9.107, 3.105], [5.369, 4.98], [9.107, 4.98]].forEach(([x, y]) => {
    oval(sl, x, y, 0.409, 0.409, { fill: { color: C.white } });
    icon(sl, 'check', x + 0.205, y + 0.205, 0.21, C.tan);
    txt(sl, 'Solution Here', {
      x: x + 0.598, y: y - 0.003, w: 2.161, h: 0.303,
      fontFace: HEAD, fontSize: 12, color: C.white, charSpacing: 1,
    });
    txt(sl, 'There are many variations of passages of Lorem available,', {
      x: x + 0.598, y: y + 0.298, w: 2.484, h: 0.585, fontSize: 10, color: C.white, lineSpacing: 18,
    });
  });
  hline(sl, 4.584, 4.412, 8.216, C.white, 0.75, 48);
  txt(sl, lines(['Market', 'Solution']), {
    x: 0.714, y: 1.17, w: 2.675, h: 1.569, fontFace: HEAD, fontSize: 44, color: C.tan,
  });
  txt(sl, lines(['There variations ', 'are many']), {
    x: 0.737, y: 3.929, w: 3.304, h: 0.692, fontFace: HEAD, fontSize: 14, color: C.tan, lineSpacing: 22,
  });
  txt(sl, LOREM.variations, {
    x: 0.737, y: 4.995, w: 2.932, h: 1.433, align: 'justify',
    fontFace: POPPINS, fontSize: 10, color: C.white, transparency: 30, lineSpacing: 16,
  });
}

// 20 — Perspective target
function slide20(sl) {
  const cols = [
    { x: 0.99, pct: '60%', px: 1.231, ex: 1.334, knob: 2.116, knobFill: C.tan, knobLine: C.black, bar: C.tan, tx: 0.853 },
    { x: 3.156, pct: '30%', px: 3.441, ex: 3.501, knob: 4.343, knobFill: C.green, knobLine: C.tan, bar: C.sand, tx: 3.156 },
    { x: 5.323, pct: '20%', px: 5.565, ex: 5.667, knob: 6.528, knobFill: C.tan, knobLine: C.black, bar: C.tan, tx: 5.323 },
  ];
  cols.forEach((c) => {
    oval(sl, c.x, 3.064, 1.625, 1.625, { fill: { type: 'none' }, line: { color: C.grey85, width: 0.75 } });
    txt(sl, c.pct, {
      x: c.px, y: 3.415, w: 1.142, h: 0.572, align: 'center', fontFace: HEAD, fontSize: 28, color: C.tan,
    });
    txt(sl, lines(['EXAMPLE', 'TEXT']), {
      x: c.ex, y: 3.976, w: 0.937, h: 0.471, align: 'center', fontSize: 11, color: C.white,
    });
    roundRect(sl, c.x, 5.365, 1.711, 0.073, 0.5, { fill: { color: C.grey95 } });
    roundRect(sl, c.x, 5.365, 1.24, 0.073, 0.5, { fill: { color: c.bar } });
    oval(sl, c.knob, 5.325, 0.153, 0.153, {
      fill: { color: c.knobFill }, line: { color: c.knobLine, width: 0.25, transparency: 50 },
    });
    txt(sl, [
      { text: 'Lorem' },
      { text: ' ipsum\u00a0is simply dummy text of the printing typesetting industry. Lorem standard' },
    ], { x: c.tx, y: 5.874, w: 1.923, h: 0.92, fontSize: 10, color: C.white, lineSpacing: 15 });
  });
  txt(sl, lines(['Perspective', 'Target']), {
    x: 0.853, y: 0.86, w: 4.609, h: 1.569, fontFace: HEAD, fontSize: 40, color: C.tan,
  });
  photo(sl, 8.016, 0, 5.318, 7.5);
}

// 21 — Custom table chart (scatter)
function slide21(sl) {
  sl.background = { color: C.tan };
  sl.addChart(pptx.ChartType.scatter, [
    { name: 'X-Axis', values: [0.7, 1.8, 2.6] },
    { name: 'Y-Values', values: [2.7, 3.2, 0.8] },
  ], {
    x: 5.56, y: 1.696, w: 6.857, h: 4.426,
    lineSize: 0, lineDataSymbol: 'circle', lineDataSymbolSize: 5,
    chartColors: [C.green], showLegend: false,
    showValue: true, dataLabelPosition: 't', dataLabelColor: '5A5A5A', dataLabelFontSize: 12,
    catAxisLabelColor: '5A5A5A', valAxisLabelColor: '5A5A5A',
    catGridLine: { color: 'D9D9D9', size: 0.75 }, valGridLine: { color: 'D9D9D9', size: 0.75 },
    catAxisMinVal: 0.5, catAxisMaxVal: 3, valAxisMinVal: 0, valAxisMaxVal: 3.5,
    chartArea: { fill: { color: C.white } }, plotArea: { fill: { color: C.grey95 } },
  });
  txt(sl, lines(['Custom', 'Table Chart']), {
    x: 1.048, y: 1.447, w: 3.606, h: 1.433, fontFace: HEAD, fontSize: 44, color: C.white,
  });
  txt(sl, lines(['Customizable Chart ', 'Function For User']), {
    x: 1.048, y: 3.75, w: 3.606, h: 0.692, fontFace: HEAD, fontSize: 14, color: C.white, lineSpacing: 22,
  });
  txt(sl, LOREM.variations, {
    x: 1.048, y: 4.694, w: 3.606, h: 1.209, align: 'justify',
    fontFace: POPPINS, fontSize: 10, color: C.white, lineSpacing: 16,
  });
}

// 22 — Harvest product
function slide22(sl) {
  photo(sl, 0.021, 0, 5.53, 7.511);
  rect(sl, 5.54, 2.97, 7.794, 4.542, C.tan);
  txt(sl, 'Harvest Product', {
    x: 6.123, y: 0.763, w: 5.165, h: 0.661, fontFace: HEAD, fontSize: 44, color: C.tan,
  });
  txt(sl, LOREM.variations, {
    x: 6.123, y: 1.66, w: 6.301, h: 0.76,
    fontFace: POPPINS, fontSize: 10, color: C.white, transparency: 30, lineSpacing: 16,
  });
  txt(sl, 'There are many variations', {
    x: 0.65, y: 5.25, w: 2.047, h: 0.718,
    fontFace: HEAD, fontSize: 16, color: C.white, underline: { style: 'sng' }, lineSpacing: 22,
  });
  txt(sl, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humor', {
    x: 0.65, y: 6.111, w: 4.26, h: 0.76, fontSize: 11, color: C.white, lineSpacing: 16,
  });
  oval(sl, 6.314, 5.473, 0.379, 0.379, { fill: { type: 'none' }, line: { color: C.white, width: 0.5 } });
  txt(sl, '1', {
    x: 6.123, y: 5.51, w: 0.761, h: 0.337, align: 'center', fontSize: 14, color: C.white, margin: 0,
  });
  txt(sl, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humor', {
    x: 6.215, y: 6.002, w: 4.26, h: 0.76, fontSize: 11, color: C.white, transparency: 30, lineSpacing: 16,
  });
  txt(sl, 'BRAND [CLIENT NAME]', { x: 0.403, y: 7.129, w: 3.722, h: 0.269, fontSize: 10, color: C.white });
}

// 23 — Gallery grid
function slide23(sl) {
  [[0, 0, 4.523, 7.5], [4.552, 0, 4.523, 3.75], [4.552, 3.75, 4.523, 3.75], [9.103, 0, 4.23, 7.5]]
    .forEach(([x, y, w, h]) => photo(sl, x, y, w, h));
  const blocks = [
    { x: 0.65, ty: 5.41, by: 6.271, bh: 0.751, body: 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration' },
    { x: 4.91, ty: 2.101, by: 2.963, bh: 0.526, body: 'There are many variations of passages of Lorem Ipsum available' },
    { x: 4.91, ty: 5.454, by: 6.315, bh: 0.526, body: 'There are many variations of passages of Lorem Ipsum available' },
    { x: 9.432, ty: 5.41, by: 6.271, bh: 0.526, body: 'There are many variations of passages of Lorem Ipsum available' },
  ];
  blocks.forEach((b) => {
    txt(sl, 'There are many variations', {
      x: b.x, y: b.ty, w: 2.047, h: 0.718,
      fontFace: HEAD, fontSize: 16, color: C.white, underline: { style: 'sng' }, lineSpacing: 22,
    });
    txt(sl, b.body, { x: b.x, y: b.by, w: 3.442, h: b.bh, fontSize: 11, color: C.white, lineSpacing: 16 });
  });
}

// 24 — Team
function slide24(sl) {
  sl.background = { color: C.tan };
  txt(sl, lines(['Behind the ', 'Brand Team']), {
    x: 1.277, y: 1.095, w: 6.301, h: 1.447, fontFace: HEAD, fontSize: 40, color: C.white,
  });
  txt(sl, LOREM.sed, {
    x: 1.259, y: 3.154, w: 2.498, h: 2.436, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
  [4.726, 7.518, 10.309].forEach((x) => {
    txt(sl, [
      { text: 'Your Name', options: { breakLine: true, fontSize: 18 } },
      { text: 'Position', options: { fontSize: 14 } },
    ], { x, y: 5.512, w: 1.823, h: 0.64, align: 'right', fontFace: HEAD, color: C.white });
  });
}

// 25 — Tablet + stylus mock-up (dark, near-monochrome)
function slide25(sl) {
  // two big rounded slabs rotated 45°, offset from each other across the slide
  [{ x: -3.10, y: -1.49 }, { x: 6.42, y: -0.09 }].forEach((t) => sl.addShape(S.roundRect, {
    x: t.x, y: t.y, w: 7.4, h: 9.9, rotate: 45, rectRadius: 0.35,
    fill: { color: C.bg }, line: { color: '8E8E8E', width: 1 },
  }));
  // the two white styluses, parallel to the slabs' long edge
  [{ cx: 4.00, cy: 6.08, len: 4.35 }, { cx: 8.52, cy: 2.66, len: 6.24 }].forEach((p) =>
    sl.addShape(S.rect, {
      x: p.cx - p.len / 2, y: p.cy - 0.08, w: p.len, h: 0.16, rotate: 315, fill: { color: C.white },
    }));
}

// 26 — Billboard advert
function slide26(sl) {
  sl.background = { color: C.tan };
  // support pole
  rect(sl, 8.713, 5.075, 0.407, 1.804, C.grey75);
  rect(sl, 8.713, 5.075, 0.124, 1.804, C.grey85);
  rect(sl, 8.996, 5.075, 0.124, 1.804, C.grey65);
  // board
  rect(sl, 5.136, 1.512, 7.565, 3.624, C.green);
  sl.addShape(S.rect, {
    x: 5.136, y: 1.512, w: 7.565, h: 3.624, fill: { type: 'none' }, line: { color: C.grey95, width: 10 },
  });
  // hanging clamps
  [5.951, 7.659, 9.929, 11.637].forEach((x) => {
    sl.addShape(S.rect, { x: x + 0.146, y: 5.144, w: 0.048, h: 0.214, rotate: 32, fill: { color: C.grey65 } });
    oval(sl, x + 0.129, 5.126, 0.119, 0.125, { fill: { color: C.grey95 } });
    roundRect(sl, x, 5.293, 0.164, 0.135, 0.13, { fill: { color: C.white } });
  });
  rect(sl, 0, 6.652, 13.333, 0.848, C.sand);
  txt(sl, 'Get Ads For Your Agraria Product', {
    x: 1.011, y: 1.698, w: 3.851, h: 1.919, fontFace: HEAD, fontSize: 36, bold: true, color: C.white,
  });
  txt(sl, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi', {
    x: 1.011, y: 3.944, w: 3.352, h: 1.384, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// 27 — Mobile app mock-up
function slide27(sl) {
  sl.background = { color: C.sand };
  rect(sl, 0, 6.515, 13.333, 0.985, C.tan);
  const phones = [
    { bx: 7.023, by: 1.68, sx: 7.155, sy: 1.807 },
    { bx: 9.621, by: -0.863, sx: 9.761, sy: -0.736 },
    { bx: 9.621, by: 3.889, sx: 9.761, sy: 4.016 },
  ];
  phones.forEach((p) => {
    roundRect(sl, p.bx, p.by, 2.192, 4.447, 0.09, { fill: { color: '111111' } });
    rect(sl, p.bx - 0.025, p.by + 0.95, 0.035, 0.3, '3A3A3A');
    rect(sl, p.bx - 0.025, p.by + 1.4, 0.035, 0.5, '3A3A3A');
    photo(sl, p.sx, p.sy, 1.932, 4.187, {
      fill: C.grey95, edge: C.grey95, label: 'Image Placeholder', labelColor: C.grey75, labelTop: true,
    });
  });
  txt(sl, 'Mobile Apps to achieve your Goals!', {
    x: 1.248, y: 1.629, w: 3.851, h: 2.121, fontFace: HEAD, fontSize: 40, bold: true, color: C.tan,
  });
  txt(sl, LOREM.sed, {
    x: 1.213, y: 4.118, w: 4.111, h: 1.399, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// 28 — Laptop mock-up
function slide28(sl) {
  sl.background = { color: C.tan };
  rect(sl, 7.306, 0, 4.02, 7.5, C.sand);
  roundRect(sl, 6.11, 1.973, 5.86, 4.0, 0.03, { fill: { color: '141414' } });
  roundRect(sl, 6.257, 2.207, 5.591, 3.493, 0.006, { fill: { color: 'ECEDEF' } });
  roundRect(sl, 5.457, 5.973, 7.19, 0.19, 0.45, { fill: { color: C.grey30 } });
  rect(sl, 8.25, 5.995, 1.6, 0.055, '6E6E6E');
  txt(sl, 'Laptop Mock up', {
    x: 1.449, y: 2.195, w: 3.81, h: 1.555, fontFace: HEAD, fontSize: 48, color: C.white,
    charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
  txt(sl, LOREM.dolor, {
    x: 1.449, y: 4.049, w: 3.81, h: 1.121, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3,
  });
}

/* World map, drawn as a coarse raster of tan cells (78 x 46 grid). */
const WORLD_MAP = [
  '.....................####....######...........................................',
  '..................###################.........................................',
  '.................######.############.......####................##.............',
  '.............#....####.#############........##...................#............',
  '...........####.#####...############...................##......#####......#...',
  '...........#.....####......#########..................#......#######......#...',
  '..........#####.##.####....########..................#...#.#############..###.',
  '.#####...#..####.#.#####....#######..........###.........#####################',
  '#####################.###...######..........######.###########################',
  '#####################.####..####...##......######.############################',
  '.##################.#..##...###...........####################################',
  '##################...###.....#............###.################################',
  '..##...###########....####.............#..###.#############################...',
  '.........###########..####.............#..##.#############################....',
  '.........##################...........###.#################################...',
  '..........################..............####################################..',
  '..........###############...............##################################....',
  '..........##############................#################################..#..',
  '..........#############...............###...############################..#...',
  '...........###########................##.......#######################.#..#...',
  '...........###########.................####......#####################.###....',
  '............#########.................################################........',
  '.............####...#................#################################........',
  '..............###....................##################..############.........',
  '...............##..#................###################...####.####...........',
  '................####................##############.###....##...####..#........',
  '...................##...............################.......#....###...#.......',
  '....................#.####...........################......#..........#.......',
  '......................######..........###.##########.............#..#.........',
  '.....................########..............########.............##.##.........',
  '.....................##########............########..............#.###..###...',
  '.....................###########...........#######.................#......##..',
  '.....................###########............######............................',
  '......................#########............#######..#..................###....',
  '.......................########............#######.##.................######..',
  '.......................########.............#####..#................#########.',
  '.......................######...............#####..#................##########',
  '.......................######...............####....................##########',
  '.......................#####.................###....................##########',
  '.......................####...............................................###.',
  '.......................###.................................................#..',
  '.......................##.....................................................',
  '.......................##.....................................................',
  '......................###.....................................................',
  '......................##......................................................',
  '.......................#......................................................',
];

// 29 — Worldwide target
function slide29(sl) {
  const ox = 5.514, oy = 1.664, cw = 0.1002, ch = 0.1037;
  WORLD_MAP.forEach((row, r) => {
    let c = 0;
    while (c < row.length) {
      if (row[c] === '#') {
        let n = 1;
        while (row[c + n] === '#') n++;
        rect(sl, ox + c * cw, oy + r * ch, n * cw + 0.004, ch + 0.004, C.tan);
        c += n;
      } else c++;
    }
  });
  txt(sl, lines(['Worldwide', 'Target']), {
    x: 1.209, y: 2.195, w: 3.81, h: 1.555, fontFace: HEAD, fontSize: 48, color: C.tan,
    charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
  txt(sl, LOREM.dolor, {
    x: 1.209, y: 4.049, w: 3.81, h: 1.121, fontSize: 12, color: C.tan, lineSpacingMultiple: 1.3,
  });
}

// 30 — Thank you
function slide30(sl) {
  rect(sl, 0.427, 4.003, 7.758, 3.104, C.tan);
  txt(sl, 'Thank You Very Much!', {
    x: 0.949, y: 4.893, w: 5.551, h: 1.168,
    fontFace: HEAD, fontSize: 44, color: C.white, lineSpacingMultiple: 0.7,
  });
  txt(sl, 'For Your Attention', {
    x: 0.949, y: 6.061, w: 3.38, h: 0.305, fontSize: 12, color: C.white, lineSpacing: 16,
  });
}

/* ------------------------------------------------------------------ *
 * Assemble
 * ------------------------------------------------------------------ */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Whisper Agency Design';
pptx.author = 'Whisper Agency';

BUILDERS.forEach((build) => {
  const sl = pptx.addSlide();
  sl.background = { color: C.bg }; // slide master default; individual slides override
  build(sl);
});

pptx
  .writeFile({ fileName: path.join(__dirname, '0c1380ce-d077-48ab-b7a4-649fb30e47cd_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
