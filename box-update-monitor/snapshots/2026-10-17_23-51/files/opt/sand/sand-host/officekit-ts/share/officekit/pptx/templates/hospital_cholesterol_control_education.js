/**
 * "Dr. Stones" medical presentation — rebuilt with pptxgenjs.
 *
 * Run:  node 0fa9c1b9-8095-45f0-94f7-0c923de4672a_grok_final.js
 * Out:  0fa9c1b9-8095-45f0-94f7-0c923de4672a_grok_final.pptx (next to this file)
 *
 * Photographs in the original deck are re-created as flat placeholder
 * rectangles; every other element is a native pptxgenjs shape or text run.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

const C = {
  teal: '169199', // brand primary
  coral: 'EB6655', // brand accent
  dark: '09393D', // deep teal used for small headings
  ink: '000000', // headline black
  name: '0C1307', // near-black used for people's names
  green: '253917', // the two percentage read-outs
  body: '7F7F7F', // body copy grey
  white: 'FFFFFF',
  track: 'D9D9D9', // empty part of a progress bar
  photo: 'C3C3C3', // stand-in for a photograph
  photoDark: '7F7F7F', // stand-in for the dimmed cover photo
  frame: '1F1F1F', // device mock-up bezel
};

const F = { head: 'Oxygen', body: 'Open Sans' };

// Soft drop shadow used by the "feature" cards.
// A fresh object every call — pptxgenjs rewrites the one it is handed.
const softShadow = (opacity, blur) => ({ type: 'outer', color: '000000', opacity, blur, offset: 0, angle: 90 });

/* ------------------------------------------------------------------ *
 * Brand mark
 *
 * The logo is an open arch (a half ring) cradling a smaller solid dome.
 * Both are stored as normalised paths — x/y as a fraction of the shape's
 * own box — so the mark can be scaled and turned freely.
 * ------------------------------------------------------------------ */

const ARCH_RING = [
  { x: 1.0, y: 1.0 },
  { x: 0.8504, y: 1.0 },
  { curve: { type: 'cubic', x1: 0.8282, y1: 0.6196, x2: 0.6798, y2: 0.3256 }, x: 0.5, y: 0.3256 },
  { curve: { type: 'cubic', x1: 0.3202, y1: 0.3256, x2: 0.1717, y2: 0.6197 }, x: 0.1496, y: 1.0 },
  { x: 0.0, y: 1.0 },
  { curve: { type: 'cubic', x1: 0.0228, y1: 0.4395, x2: 0.238, y2: 0.0 }, x: 0.5, y: 0.0 },
  { curve: { type: 'cubic', x1: 0.762, y1: 0.0, x2: 0.9772, y2: 0.4395 }, x: 1.0, y: 1.0 },
  { close: true },
];

const ARCH_DOME = [
  { x: 1.0, y: 1.0 },
  { x: 0.0, y: 1.0 },
  { curve: { type: 'cubic', x1: 0.0449, y1: 0.4304, x2: 0.2517, y2: 0.0 }, x: 0.5, y: 0.0 },
  { curve: { type: 'cubic', x1: 0.7483, y1: 0.0, x2: 0.9551, y2: 0.4304 }, x: 1.0, y: 1.0 },
  { close: true },
];

const ARCH_RATIO = 0.4565; // ring height / ring width
const DOME_W = 0.4694; // dome width  / ring width
const DOME_H = 0.4252; // dome height / ring height
const DOME_DY = 0.2874; // dome centre offset above the ring centre (x ring height)

function scalePath(points, w, h) {
  return points.map(p => {
    if (p.close) return p;
    const out = { x: p.x * w, y: p.y * h };
    if (p.curve) {
      out.curve = {
        type: 'cubic',
        x1: p.curve.x1 * w,
        y1: p.curve.y1 * h,
        x2: p.curve.x2 * w,
        y2: p.curve.y2 * h,
      };
    }
    return out;
  });
}

/**
 * Draw the logo mark centred on (cx, cy).
 * `turn` = extra clockwise quarter-turns on top of the 180 deg flip that
 * turns the arch into an upward-opening bowl.
 */
function brandMark(slide, cx, cy, w, ringColor, coreColor, turn) {
  const h = w * ARCH_RATIO;
  const rotate = (180 + turn * 90) % 360;
  const place = (dx, dy, bw, bh, points, color) => {
    let ox = dx;
    let oy = dy;
    for (let i = 0; i < turn; i++) {
      const t = ox;
      ox = -oy;
      oy = t;
    }
    slide.addShape('custGeom', {
      x: cx + ox - bw / 2,
      y: cy + oy - bh / 2,
      w: bw,
      h: bh,
      points: scalePath(points, bw, bh),
      fill: { color },
      line: { type: 'none' },
      rotate,
    });
  };
  place(0, 0, w, h, ARCH_RING, ringColor);
  place(0, -h * DOME_DY, w * DOME_W, h * DOME_H, ARCH_DOME, coreColor);
}

/** Header lockup: the quarter-turned mark plus the two-line wordmark. */
function logo(slide, x) {
  brandMark(slide, x + 0.074, 0.8095, 0.318, C.teal, C.coral, 1);
  slide.addText('DR. STONES', {
    x: x + 0.278, y: 0.568, w: 1.461, h: 0.303,
    fontFace: F.head, fontSize: 12, bold: true, color: C.dark, valign: 'top',
  });
  slide.addText('best service', {
    x: x + 0.278, y: 0.8, w: 1.461, h: 0.236,
    fontFace: F.head, fontSize: 8, color: C.coral, valign: 'top',
  });
}

/* ------------------------------------------------------------------ *
 * Icons — small pictograms drawn from primitives
 * ------------------------------------------------------------------ */

const CHECK_PATH = [
  { x: 0.0, y: 0.27 },
  { x: 0.31, y: 0.70 },
  { x: 0.93, y: 0.0 },
  { x: 1.0, y: 0.12 },
  { x: 0.33, y: 0.95 },
  { x: 0.0, y: 0.52 },
  { close: true },
];

/** Teal/coral disc with a white tick — the deck's bullet marker. */
function checkBadge(slide, x, y, d, color) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' } });
  slide.addShape('custGeom', {
    x: x + d * 0.25, y: y + d * 0.339, w: d * 0.5, h: d * 0.349,
    points: scalePath(CHECK_PATH, d * 0.5, d * 0.349),
    fill: { color: C.white }, line: { type: 'none' },
  });
}

/** Outlined circle enclosing an outlined cross (the "medical" icon). */
function medicalCrossIcon(slide, x, y, d, color) {
  const stroke = { color, width: 1.5 };
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: stroke });
  slide.addShape('plus', {
    x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52,
    fill: { type: 'none' }, line: stroke,
  });
}

/** Syringe drawn on a 45 deg diagonal. */
function syringeIcon(slide, x, y, d, color) {
  const solid = { fill: { color }, line: { type: 'none' }, rotate: 315 };
  slide.addShape('roundRect', { x: x + d * 0.26, y: y + d * 0.42, w: d * 0.52, h: d * 0.20, rectRadius: d * 0.05, ...solid });
  slide.addShape('rect', { x: x + d * 0.66, y: y + d * 0.15, w: d * 0.34, h: d * 0.07, ...solid });
  slide.addShape('rect', { x: x + d * 0.10, y: y + d * 0.66, w: d * 0.20, h: d * 0.07, ...solid });
  slide.addShape('rect', { x: x + d * 0.44, y: y + d * 0.30, w: d * 0.05, h: d * 0.17, ...solid });
  slide.addShape('rect', { x: x + d * 0.56, y: y + d * 0.42, w: d * 0.05, h: d * 0.17, ...solid });
}

const DOWN_ARROW_PATH = [
  { x: 0.289, y: 0.0 },
  { x: 0.711, y: 0.0 },
  { x: 0.711, y: 0.543 },
  { x: 1.0, y: 0.543 },
  { x: 0.5, y: 1.0 },
  { x: 0.0, y: 0.543 },
  { x: 0.289, y: 0.543 },
  { close: true },
];

/** Arrow dropping into an open tray. `d` is the enclosing disc diameter. */
function downloadIcon(slide, x, y, d, color) {
  const solid = { fill: { color }, line: { type: 'none' } };
  const aw = d * 0.294;
  const ah = d * 0.395;
  slide.addShape('custGeom', {
    x: x + d * 0.356, y: y + d * 0.26, w: aw, h: ah,
    points: scalePath(DOWN_ARROW_PATH, aw, ah), ...solid,
  });
  slide.addShape('rect', { x: x + d * 0.271, y: y + d * 0.576, w: d * 0.04, h: d * 0.164, ...solid });
  slide.addShape('rect', { x: x + d * 0.694, y: y + d * 0.576, w: d * 0.04, h: d * 0.164, ...solid });
  slide.addShape('rect', { x: x + d * 0.271, y: y + d * 0.700, w: d * 0.463, h: d * 0.04, ...solid });
}

/* ------------------------------------------------------------------ *
 * Building blocks
 * ------------------------------------------------------------------ */

/** Placeholder standing in for a photograph in the source deck. */
function photo(slide, x, y, w, h, opts = {}) {
  const o = { x, y, w, h, fill: { color: opts.color || C.photo }, line: { type: 'none' } };
  if (opts.shape === 'ellipse') slide.addShape('ellipse', o);
  else if (opts.radius) slide.addShape('roundRect', { ...o, rectRadius: opts.radius });
  else slide.addShape('rect', o);
}

/** Big two-tone headline. `parts` = [[text, color], ...] on one line. */
function heading(slide, x, y, w, h, parts, size) {
  slide.addText(
    parts.map(([text, color]) => ({ text, options: { color } })),
    { x, y, w, h, fontFace: F.head, fontSize: size, bold: true, color: C.ink, align: 'left', valign: 'top' }
  );
}

/** Small teal kicker above a headline. */
function eyebrow(slide, x, y, w, text, color = C.teal) {
  slide.addText(text, {
    x, y, w, h: 0.372, fontFace: F.body, fontSize: 12, color,
    lineSpacingMultiple: 1.5, align: 'left', valign: 'top',
  });
}

/** Justified grey body copy. */
function paragraph(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: F.body, fontSize: 12, color: opts.color || C.body,
    align: opts.align || 'justify', valign: 'top',
    lineSpacingMultiple: opts.lineSpacing || 1.66666,
  });
}

/** Bold dark-teal sub-heading (14pt Oxygen). */
function subhead(slide, x, y, w, text, opts = {}) {
  slide.addText(text, {
    x, y, w: w, h: opts.h || 0.355,
    fontFace: F.head, fontSize: opts.size || 14, bold: true, color: opts.color || C.dark,
    align: opts.align || 'left', valign: 'top', lineSpacingMultiple: opts.lineSpacing || 1.42857,
  });
}

/** Coral capsule button. */
function pill(slide, x, y, w, h, label) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: h / 2, fill: { color: C.coral }, line: { type: 'none' },
  });
  slide.addText(label, {
    x, y, w, h, fontFace: F.body, fontSize: 12, color: C.white, align: 'center', valign: 'middle',
  });
}

/** Rounded card sitting on a soft, centred shadow. */
function card(slide, x, y, w, h, radius, opts = {}) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: radius,
    fill: { color: opts.color || C.white }, line: { type: 'none' },
    shadow: softShadow(opts.opacity || 0.267, opts.blur || 5),
  });
}

/** Rounded track + coloured fill + label + percentage. */
function progressRow(slide, x, y, label, pct, color) {
  const trackW = 4.435;
  slide.addShape('roundRect', {
    x: x + 0.04, y: y + 0.437, w: trackW, h: 0.123, rectRadius: 0.0615,
    fill: { color: C.track }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: x + 0.04, y: y + 0.437, w: trackW * pct, h: 0.123, rectRadius: 0.0615,
    fill: { color }, line: { type: 'none' },
  });
  subhead(slide, x, y, 2.048, label, { h: 0.362, align: 'justify' });
  subhead(slide, x + 3.793, y, 0.682, `${Math.round(pct * 100)}%`, { h: 0.362, align: 'justify', color: C.green });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/** Slides 1 and 20 share the full-bleed cover treatment. */
function coverSlide(pres, title, subtitle) {
  const s = pres.addSlide();
  photo(s, 0, 0, SLIDE_W, SLIDE_H, { color: C.photoDark });
  s.addText(subtitle, {
    x: 4.222, y: 2.225, w: 4.89, h: 0.362,
    fontFace: F.body, fontSize: 14, color: C.white,
    align: 'center', valign: 'top', lineSpacingMultiple: 1.42857,
  });
  s.addText(title, {
    x: 2.219, y: 2.692, w: 8.895, h: 1.212,
    fontFace: F.head, fontSize: 66, bold: true, color: C.white, align: 'center', valign: 'top',
  });
  brandMark(s, 6.677, 4.491, 1.402, C.white, C.coral, 0);
  [3.7, 7.168].forEach(x => {
    s.addShape('line', { x, y: 4.174, w: 2.489, h: 0, line: { color: C.white, width: 1.5 } });
  });
  return s;
}

const LOREM = {
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore',
  mid: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim',
  card: 'Lorem ipsum dolors situp amet, consectetur adipiscing elitdoei us modt empor',
  num: 'Lorem ipsum dolor sit amet, cons ectetur adipiscing elit do eiusmod tempor incidid untut labor dolore',
};

function slide02(pres) {
  const s = pres.addSlide();
  photo(s, 8.194, 0, 5.139, 7.5);
  logo(s, 1.398);
  eyebrow(s, 1.322, 1.971, 3.196, 'consuming alcoholic beverages');
  heading(s, 1.326, 2.35, 5.704, 1.178,
    [['Difficulty breathing, and short ', C.ink], ['duration of breath', C.teal]], 32);
  paragraph(s, 1.322, 3.749, 5.344, 0.675,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad',
    { lineSpacing: 1.5 });
  subhead(s, 1.322, 4.645, 4.89, 'Controlling cholesterol levels in the blood',
    { size: 16, h: 0.362, align: 'justify', lineSpacing: 1.25 });
  paragraph(s, 1.322, 5.135, 5.344, 0.978,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation',
    { lineSpacing: 1.5 });
}

function slide03(pres) {
  const s = pres.addSlide();
  photo(s, 1.21, 0, 4.306, 7.5);
  logo(s, 7.586);
  eyebrow(s, 7.524, 2.507, 4.287, 'Knowledge of hospital rules');
  heading(s, 7.524, 2.879, 4.589, 1.313,
    [['Our strategy From ', C.dark], ['Dr. Stones', C.teal]], 36);
  paragraph(s, 7.524, 4.374, 4.598, 0.978,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud',
    { lineSpacing: 1.5 });
  paragraph(s, 7.524, 5.485, 4.598, 0.675, LOREM.short, { lineSpacing: 1.5 });
}

function slide04(pres) {
  const s = pres.addSlide();
  logo(s, 1.398);
  pill(s, 1.398, 1.653, 1.806, 0.474, 'About Us');
  heading(s, 1.249, 2.437, 4.646, 1.313,
    [['Improving public health & ', C.dark], ['nutrition', C.teal]], 36);
  paragraph(s, 1.249, 4.06, 4.351, 0.916,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis');
  paragraph(s, 1.249, 5.125, 4.448, 0.635, LOREM.short);
  s.addShape('rect', { x: 6.962, y: 5.536, w: 1.681, h: 1.964, fill: { color: C.coral }, line: { type: 'none' } });
  medicalCrossIcon(s, 7.213, 5.928, 1.179, C.white);
  photo(s, 9.0, 5.536, 3.084, 1.964);
  photo(s, 6.964, 0, 5.12, 5.214);
}

function slide05(pres) {
  const s = pres.addSlide();
  logo(s, 1.398);
  heading(s, 1.128, 2.411, 4.583, 1.919,
    [['Improved disease control quickly & ', C.ink], ['effectively', C.teal]], 36);
  paragraph(s, 1.128, 4.476, 4.583, 1.196,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis incididunt ut labordolore ');
  pill(s, 1.209, 5.931, 1.575, 0.489, 'Explore Us');
  card(s, 7.225, 3.318, 4.981, 1.426, 0, { color: C.coral, opacity: 0.227, blur: 9 });
  subhead(s, 7.691, 3.508, 4.064, 'Provide honest & complete information', { color: C.white });
  paragraph(s, 7.691, 3.897, 4.064, 0.635,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut', { color: C.white });
  photo(s, 7.225, 4.994, 4.981, 2.506);
  photo(s, 7.225, 0, 4.981, 3.068);
}

function slide06(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0.818, y: 0, w: 0.219, h: 2.542, fill: { color: C.coral }, line: { type: 'none' } });
  photo(s, 1.029, 0, 5.276, 6.196);
  logo(s, 7.586);
  heading(s, 7.491, 1.304, 4.071, 1.313,
    [['About Hospital ', C.ink], ['Dr. Stones', C.teal]], 36);
  paragraph(s, 7.498, 2.792, 4.503, 1.757,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation exercitation magna aliquaenim ad minim veniam quis nostrud exercitation eiusmod tempor incididunt ut labordolore');
  pill(s, 7.6, 4.862, 1.575, 0.489, 'Explore Us');
  checkBadge(s, 7.6, 5.741, 0.456, C.teal);
  paragraph(s, 8.37, 5.631, 3.63, 0.675,
    'PLACEHOLDER',
    { color: C.dark, align: 'left', lineSpacing: 1.5 });
}

function slide07(pres) {
  const s = pres.addSlide();
  logo(s, 1.398);
  heading(s, 1.307, 1.453, 4.812, 1.313,
    [['Patient Rights That Must ', C.ink], ['Be Obtained', C.teal]], 36);
  [
    { badgeY: 0.615, titleY: 0.636, bodyY: 1.064, title: 'Get services that are humane' },
    { badgeY: 5.373, titleY: 5.394, bodyY: 5.821, title: 'Obtain quality health services' },
  ].forEach(row => {
    checkBadge(s, 7.6, row.badgeY, 0.456, C.teal);
    subhead(s, 8.345, row.titleY, 3.274, row.title, { h: 0.362, align: 'justify' });
    paragraph(s, 8.345, row.bodyY, 3.682, 0.916, LOREM.mid);
  });
  photo(s, 7.6, 2.495, 4.401, 2.418);
  photo(s, 1.388, 3.355, 4.731, 3.409);
}

function slide08(pres) {
  const s = pres.addSlide();
  photo(s, 1.157, 0.651, 5.093, 6.049);
  logo(s, 7.586);
  heading(s, 7.538, 1.917, 4.148, 1.313,
    [['Our strategy in ', C.ink], ['Dr. Stones', C.teal]], 36);
  [
    { badgeY: 3.643, titleY: 3.653, bodyY: 4.178, title: 'Improved disease control quickly ' },
    { badgeY: 5.211, titleY: 5.258, bodyY: 5.784, title: 'Strengthening the health system' },
  ].forEach(row => {
    checkBadge(s, 7.597, row.badgeY, 0.456, C.teal);
    subhead(s, 8.236, row.titleY, 3.441, row.title, { h: 0.362 });
    paragraph(s, 7.538, row.bodyY, 4.462, 0.916,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud');
  });
}

function slide09(pres) {
  const s = pres.addSlide();
  logo(s, 1.398);
  s.addShape('rect', { x: 0, y: 2.434, w: 1.148, h: 1.295, fill: { color: C.coral }, line: { type: 'none' } });
  heading(s, 1.326, 2.418, 4.265, 1.313, [['The Target for Public Health', C.ink]], 36);
  paragraph(s, 1.326, 3.969, 4.706, 1.196,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation');
  paragraph(s, 1.326, 5.255, 4.706, 0.635, LOREM.short);
  pill(s, 1.398, 6.138, 1.575, 0.489, 'Explore Us');
  photo(s, 7.073, 0, 4.929, 4.05);
  // Teal "Body weight" tile
  s.addShape('rect', { x: 7.073, y: 4.214, w: 2.375, h: 2.413, fill: { color: C.teal }, line: { type: 'none' } });
  syringeIcon(s, 8.015, 4.453, 0.491, C.white);
  subhead(s, 7.073, 5.069, 2.375, 'Body weight', { h: 0.362, align: 'center', color: C.white });
  paragraph(s, 7.08, 5.472, 2.369, 0.916,
    'Lorem ipsum dolor sit consectetur adipiscing elit do eiusmod', { color: C.white, align: 'center' });
  photo(s, 9.633, 4.214, 2.368, 2.413);
}

function slide10(pres) {
  const s = pres.addSlide();
  photo(s, 0, 0, 5.643, 6.116);
  s.addShape('rect', { x: 0, y: 6.0785, w: 2.542, h: 0.219, fill: { color: C.coral }, line: { type: 'none' } });
  logo(s, 7.066);
  heading(s, 6.945, 1.666, 5.413, 1.313, [['Providing service fees for hospital', C.ink]], 36);
  card(s, 6.945, 3.327, 5.413, 1.524, 0.1384);
  checkBadge(s, 7.325, 3.675, 0.723, C.teal);
  subhead(s, 8.338, 3.574, 3.574, 'instructions of doctors and nurses');
  paragraph(s, 8.338, 3.947, 3.667, 0.635, LOREM.card);
  paragraph(s, 6.945, 5.199, 5.413, 0.916,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud');
}

function slide11(pres) {
  const s = pres.addSlide();
  logo(s, 7.495);
  heading(s, 7.375, 1.634, 4.556, 2.322,
    [['Agreements that have ', C.ink], ['been made', C.teal]], 44);
  [
    { y: 1.053, color: C.teal },
    { y: 3.035, color: C.coral },
    { y: 4.923, color: C.teal },
  ].forEach(row => {
    card(s, 1.265, row.y, 5.413, 1.524, 0.1384);
    checkBadge(s, 1.645, row.y + 0.347, 0.723, row.color);
    subhead(s, 2.658, row.y + 0.247, 3.574, 'instructions of doctors and nurses');
    paragraph(s, 2.658, row.y + 0.619, 3.667, 0.635, LOREM.card);
  });
  photo(s, 7.495, 4.611, 5.839, 2.889);
}

function slide12(pres) {
  const s = pres.addSlide();
  photo(s, 0, 0, 5.566, 7.5);
  logo(s, 6.842);
  eyebrow(s, 6.747, 1.351, 3.872, 'Living community movement');
  heading(s, 6.747, 1.797, 4.98, 0.774, [['Dr. Stones Hole', C.ink]], 40);
  paragraph(s, 6.747, 2.647, 5.253, 0.978,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation',
    { lineSpacing: 1.5 });
  paragraph(s, 6.747, 3.775, 5.253, 0.675,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim',
    { lineSpacing: 1.5 });
  card(s, 6.747, 4.826, 5.253, 1.874, 0.1755);
  progressRow(s, 7.091, 5.017, 'Health Resources', 0.6, C.coral);
  progressRow(s, 7.091, 5.835, 'Health Services', 0.85, C.teal);
}

function slide13(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0, y: 0.844, w: SLIDE_W, h: 0.053, fill: { color: C.teal }, line: { type: 'none' } });
  s.addShape('rect', { x: 4.517, y: 0.585, w: 4.298, h: 0.505, fill: { color: C.white }, line: { type: 'none' } });
  s.addText('Best Our Team', {
    x: 4.816, y: 0.585, w: 3.701, h: 0.505,
    fontFace: F.head, fontSize: 24, bold: true, color: C.ink, align: 'center', valign: 'top',
  });
  paragraph(s, 2.974, 1.328, 7.385, 0.635,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim',
    { align: 'center' });
  const team = [
    { x: 1.935, roleX: 1.935, roleW: 2.381, capX: 1.935, role: 'Manager Hospital', name: 'Dr. Stone Mohu' },
    { x: 5.476, roleX: 5.783, roleW: 1.779, capX: 5.476, role: 'CEO / Owner', name: 'Dr. Andreas Moc' },
    { x: 9.018, roleX: 9.319, roleW: 1.779, capX: 9.168, role: 'Doctor Specialist', name: 'Dr. Zenin Gojo' },
  ];
  team.forEach(m => {
    photo(s, m.x, 2.395, 2.381, 2.381, { shape: 'ellipse' });
    paragraph(s, m.roleX, 4.992, m.roleW, 0.355, m.role, { align: 'center' });
    s.addText(m.name, {
      x: m.x === 1.935 ? 1.935 : m.x + 0.064, y: m.x === 9.018 ? 5.306 : 5.31,
      w: m.x === 1.935 ? 2.381 : 2.265, h: 0.362,
      fontFace: F.head, fontSize: 16, bold: true, color: C.name,
      align: 'center', valign: 'top', lineSpacingMultiple: 1.25,
    });
    paragraph(s, m.capX, 5.752, 2.381, 0.635,
      'Lorem ipsum dolor sit amet consectetur dulurs om', { align: 'center' });
  });
}

function slide14(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 9.226, y: 1.857, w: 2.071, h: 2.071, fill: { color: C.coral }, line: { type: 'none' } });
  logo(s, 1.398);
  eyebrow(s, 1.295, 2.266, 4.966, 'Difficulty breathing, and short duration of breath');
  heading(s, 1.295, 2.697, 5.112, 1.919, [['Cultivating healthy living community movement', C.ink]], 36);
  paragraph(s, 1.295, 4.732, 4.537, 1.196,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation exercitation');
  pill(s, 1.379, 6.211, 1.86, 0.489, 'See more');
  photo(s, 7.189, 2.484, 2.904, 4.216);
  photo(s, 10.429, 0.651, 2.904, 4.223);
}

function slide15(pres) {
  const s = pres.addSlide();
  photo(s, 1.093, 0, 5.003, 4.537);
  photo(s, 1.093, 4.737, 3.549, 2.763);
  s.addShape('rect', { x: 4.874, y: 4.737, w: 1.222, h: 2.763, fill: { color: C.coral }, line: { type: 'none' } });
  logo(s, 7.109);
  // Vertical accent strokes above the headline
  s.addShape('line', { x: 7.779, y: 1.305, w: 0, h: 1.802, line: { color: C.coral, width: 2.25 } });
  s.addShape('line', { x: 7.966, y: 1.681, w: 0, h: 1.427, line: { color: C.teal, width: 2.25 } });
  s.addShape('line', { x: 7.596, y: 1.305, w: 0, h: 1.473, line: { color: C.teal, width: 2.25 } });
  s.addShape('rect', { x: 7.596, y: 1.305, w: 0.183, h: 0.236, fill: { color: C.coral }, line: { type: 'none' } });
  s.addShape('rect', { x: 7.784, y: 2.872, w: 0.183, h: 0.236, fill: { color: C.coral }, line: { type: 'none' } });
  eyebrow(s, 7.008, 3.329, 3.747, 'Strengthening the health system');
  heading(s, 7.008, 3.799, 5.233, 1.447, [['7 Ways to Maintain Heart Health', C.ink]], 40);
  paragraph(s, 7.008, 5.416, 5.233, 1.196,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation exercitation magna aliquaenim ad');
}

function slide16(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 10.897, y: -0.004, w: 2.451, h: 2.48, fill: { color: C.coral }, line: { type: 'none' } });
  logo(s, 1.398);
  heading(s, 1.208, 1.901, 5.076, 1.717,
    [['Welcoming young doctors who will serve the community', C.ink]], 32);
  photo(s, 7.75, 0.96, 4.376, 2.706);
  photo(s, 0, 4.345, 6.143, 3.155);
  [
    { y: 4.228, num: '01.', title: 'Improving public health and nutrition', titleW: 3.958 },
    { y: 5.539, num: '02.', title: 'instructions of doctors and nurses', titleW: 3.574 },
  ].forEach(row => {
    s.addText(row.num, {
      x: 6.786, y: row.y, w: 0.857, h: 0.64,
      fontFace: F.head, fontSize: 32, bold: true, color: C.teal, align: 'right', valign: 'top',
    });
    subhead(s, 7.875, row.y + 0.206, row.titleW, row.title);
    paragraph(s, 7.875, row.y + 0.675, 4.25, 0.635, LOREM.num);
  });
}

function slide17(pres) {
  const s = pres.addSlide();
  logo(s, 1.398);
  heading(s, 1.216, 1.507, 3.843, 1.447, [['Our Mockup Dr. Stones', C.ink]], 40);
  paragraph(s, 1.216, 3.138, 4.061, 1.477,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation tempor incididunt ut ');
  paragraph(s, 1.216, 4.799, 4.061, 0.635,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt');
  pill(s, 1.29, 5.729, 1.86, 0.489, 'See more');
  // Laptop mock-up: silver lid, dark bezel, screen placeholder, then the base rail
  s.addShape('roundRect', {
    x: 6.63, y: 1.507, w: 6.692, h: 4.5, rectRadius: 0.13,
    fill: { color: 'CDCFD1' }, line: { type: 'none' },
  });
  s.addShape('roundRect', {
    x: 6.655, y: 1.532, w: 6.642, h: 4.45, rectRadius: 0.11,
    fill: { color: '0C1116' }, line: { type: 'none' },
  });
  photo(s, 6.84, 1.73, 6.307, 4.04);
  s.addShape('ellipse', { x: 9.965, y: 1.6, w: 0.05, h: 0.05, fill: { color: '39414A' }, line: { type: 'none' } });
  // Base: light rail, mid-grey lip, dark front edge
  s.addShape('roundRect', {
    x: 5.886, y: 6.007, w: 7.436, h: 0.165, rectRadius: 0.06,
    fill: { color: 'EDEDED' }, line: { type: 'none' },
  });
  s.addShape('roundRect', {
    x: 8.85, y: 6.008, w: 1.2, h: 0.085, rectRadius: 0.042,
    fill: { color: 'DCDDDE' }, line: { type: 'none' },
  });
  s.addShape('rect', { x: 5.94, y: 6.172, w: 7.382, h: 0.063, fill: { color: '9EA0A2' }, line: { type: 'none' } });
  s.addShape('roundRect', {
    x: 5.96, y: 6.235, w: 7.362, h: 0.045, rectRadius: 0.022,
    fill: { color: '4A4A4A' }, line: { type: 'none' },
  });
}

function slide18(pres) {
  const s = pres.addSlide();
  // Phone mock-up: side buttons, then nested shells down to the screen
  [
    { y: 1.34, h: 0.253 }, { y: 1.807, h: 0.473 }, { y: 2.407, h: 0.466 },
  ].forEach(b => {
    s.addShape('roundRect', {
      x: 1.297, y: b.y, w: 0.045, h: b.h, rectRadius: 0.02,
      fill: { color: '8A8A8A' }, line: { type: 'none' },
    });
  });
  [
    { inset: 0.000, radius: 0.44, color: '2B2B2B' }, // outer shell
    { inset: 0.014, radius: 0.43, color: 'D8D8D8' }, // metallic highlight
    { inset: 0.054, radius: 0.39, color: '141414' }, // hairline
    { inset: 0.080, radius: 0.37, color: 'A8A8A8' }, // bezel body
    { inset: 0.124, radius: 0.33, color: '202020' }, // screen surround
  ].forEach(l => {
    s.addShape('roundRect', {
      x: 1.319 + l.inset, y: 0.46 + l.inset,
      w: 3.225 - 2 * l.inset, h: 6.533 - 2 * l.inset,
      rectRadius: l.radius, fill: { color: l.color }, line: { type: 'none' },
    });
  });
  photo(s, 1.475, 0.589, 2.921, 6.272, { radius: 0.2371 });
  logo(s, 6.415);
  eyebrow(s, 6.316, 1.845, 4.533, 'agreements that have been made.');
  heading(s, 6.316, 2.317, 5.865, 1.313,
    [['Fulfill the things that have been agreed upon', C.ink]], 36);
  paragraph(s, 6.316, 3.766, 5.865, 1.196,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliquaenim ad minim veniam quis nostrud exercitation veniam quis nostrud exercitation exercitation magna aliquaenim ad minim');
  paragraph(s, 6.316, 5.067, 5.865, 0.635,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore magna aliqu');
  pill(s, 6.415, 5.863, 1.84, 0.489, 'Download Now');
  s.addShape('ellipse', { x: 3.707, y: 2.556, w: 2.084, h: 2.084, fill: { color: C.white }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 4.161, y: 3.009, w: 1.178, h: 1.178, fill: { color: C.teal }, line: { type: 'none' } });
  downloadIcon(s, 4.161, 3.009, 1.178, C.white);
}

function slide19(pres) {
  const s = pres.addSlide();
  photo(s, 8.013, 0, 5.32, 7.5);
  eyebrow(s, 1.152, 1.651, 4.626, 'providing service fees for hospital');
  heading(s, 1.152, 2.03, 5.514, 0.774, [['Contact Dr. Stones', C.name]], 40);
  paragraph(s, 1.152, 2.963, 5.126, 0.916,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore mag dolor sit amet, consectetur adipiscing elit do eiusmod tempor');
  paragraph(s, 1.152, 4.061, 5.126, 0.635,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit do eiusmod tempor incididunt ut labordolore mag dolor sit');
  pill(s, 1.249, 5.002, 1.84, 0.489, 'Contact Us');
}

/* ------------------------------------------------------------------ *
 * Assemble
 * ------------------------------------------------------------------ */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DRSTONES', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'DRSTONES';
  pres.title = 'Doctor Stones';

  coverSlide(pres, 'DOCTOR STONES', 'controlling cholesterol levels in the blood');
  [slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19]
    .forEach(fn => fn(pres));
  coverSlide(pres, 'THANK YOU ', 'controlling cholesterol levels in the blood');

  return pres;
}

const OUT = path.join(__dirname, '0fa9c1b9-8095-45f0-94f7-0c923de4672a_grok_final.pptx');
build()
  .writeFile({ fileName: OUT })
  .then(f => console.log('wrote', f))
  .catch(err => {
    console.error(err);
    process.exit(1);
  });
