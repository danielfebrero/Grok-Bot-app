/**
 * Paywise pitch deck — rebuilt with pptxgenjs.
 * Recreation of "07247259-a1d4-473b-b293-96d15c21107a.pptx" (30 slides, 13.333" x 7.5").
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

const C = {
  white: 'FFFFFF',
  black: '000000',
  gray: '808080',          // tx1 lumMod 50% — body copy
  a1: '8474B7',            // light purple
  a2: '7761A9',            // mid purple
  a3: '5F4C8A',            // deep purple
  a4: '5C70C8',            // indigo
  a5: '25336E',            // navy
  a6: '130F3A',            // near-black navy
  band1: 'D8D5E5',         // table banding (accent1 tint 40%)
  band2: 'ECEBF2'          // table banding (accent1 tint 20%)
};

const HEAD = 'Roboto';     // major latin font
const BODY = 'Lato';       // minor latin font
const QUOTE = 'Open Sans';

const SIZE = { w: 13.3333333, h: 7.5 };

// ---------------------------------------------------------------------------
// Primitive helpers
// ---------------------------------------------------------------------------

let PRES = null; // set in build(); gives access to PRES.ShapeType

/** Solid rectangle; `fill` is a hex string or {color, transparency}. */
function rect(s, x, y, w, h, fill, opts) {
  const f = typeof fill === 'string' ? { color: fill } : fill;
  s.addShape(PRES.ShapeType.rect, Object.assign({ x, y, w, h, fill: f }, opts || {}));
}

/** Filled ellipse. */
function oval(s, x, y, w, h, fill) {
  s.addShape(PRES.ShapeType.ellipse, { x, y, w, h, fill: { color: fill } });
}

/** Small bullet dot (0.149" circle). */
function dot(s, x, y, fill) {
  oval(s, x, y, 0.149, 0.149, fill || C.a3);
}

/**
 * Text box. Mirrors the deck's own boxes: top anchored, 0.1"/0.05" insets,
 * word wrap on, "resize shape to fit text" like the originals.
 */
function text(s, body, x, y, w, h, o) {
  o = o || {};
  s.addText(body, {
    x, y, w, h,
    valign: 'top',
    align: o.align || 'left',
    fontFace: o.face || BODY,
    fontSize: o.size || 11,
    color: o.color || C.black,
    bold: o.bold || false,
    italic: o.italic || false,
    lineSpacingMultiple: o.lsp,
    margin: [7.2, 7.2, 3.6, 3.6],
    isTextBox: true,
    fit: 'resize'
  });
}

/** 32pt bold section heading. */
function heading(s, label, x, y, w, h, color, align) {
  text(s, label, x, y, w || 4.3, h || 0.64,
    { size: 32, bold: true, face: HEAD, color: color || C.black, align });
}

/** 14pt bold purple lead-in line. */
function lead(s, body, x, y, w, h, color) {
  text(s, body, x, y, w, h || 0.572, { size: 14, bold: true, face: HEAD, color: color || C.a1 });
}

/** 11pt gray paragraph at 1.5 line spacing. */
function para(s, body, x, y, w, h, o) {
  o = o || {};
  text(s, body, x, y, w, h, {
    size: 11, color: o.color || C.gray, lsp: 1.5, align: o.align, face: o.face
  });
}

/** Dot + one line of text; the dot is centred on the first text line. */
function bullet(s, dotX, textX, y, w, body, o) {
  o = o || {};
  dot(s, dotX, y + 0.069, o.dotColor);
  text(s, body, textX, y, w, o.h || 0.286, {
    size: 11, color: o.color || C.gray, lsp: o.lsp, face: o.face
  });
}

/**
 * Bold lead + soft line break + gray detail — the deck's most common list item.
 * `parts` is an array of {t, b} run descriptors for the second line.
 */
function bulletBlock(s, dotX, textX, y, w, h, title, parts) {
  dot(s, dotX, y + 0.069);
  const runs = [{ text: title, options: { bold: true, color: C.black } }];
  parts.forEach((p, i) => runs.push({
    text: p.t,
    options: { bold: !!p.b, color: C.gray, softBreakBefore: i === 0 }
  }));
  s.addText(runs, {
    x: textX, y, w, h, valign: 'top', fontFace: BODY, fontSize: 11,
    lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
}

/** Mixed-format single line: parts = [{t, b, c}] */
function runs(s, parts, x, y, w, h, o) {
  o = o || {};
  s.addText(parts.map(p => ({
    text: p.t, options: { bold: !!p.b, color: p.c || C.gray }
  })), {
    x, y, w, h, valign: 'top', align: o.align || 'left',
    fontFace: o.face || BODY, fontSize: o.size || 11, color: C.black,
    lineSpacingMultiple: o.lsp, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
}

/** The recurring "thin rule + heavier half rule" motif (horizontal). */
function hRule(s, y, xa, xThin, xThick, color) {
  s.addShape(PRES.ShapeType.line, {
    x: Math.min(xa, xThin), y, w: Math.abs(xThin - xa), h: 0, line: { color, width: 1 }
  });
  s.addShape(PRES.ShapeType.line, {
    x: Math.min(xa, xThick), y, w: Math.abs(xThick - xa), h: 0, line: { color, width: 2.5 }
  });
}

/** Same motif, vertical. */
function vRule(s, x, ya, yThin, yThick, color) {
  s.addShape(PRES.ShapeType.line, {
    x, y: Math.min(ya, yThin), w: 0, h: Math.abs(yThin - ya), line: { color, width: 1 }
  });
  s.addShape(PRES.ShapeType.line, {
    x, y: Math.min(ya, yThick), w: 0, h: Math.abs(yThick - ya), line: { color, width: 2.5 }
  });
}

/**
 * Right triangle — the source uses the isosceles-triangle preset with adj=0
 * plus flips, which degenerates to a right triangle.
 * `corner` names the square corner: 'bl' | 'tl' | 'br' | 'tr'.
 */
const TRI_PTS = {
  bl: [[0, 0], [0, 1], [1, 1]],
  tl: [[0, 0], [1, 0], [0, 1]],
  br: [[1, 0], [1, 1], [0, 1]],
  tr: [[0, 0], [1, 0], [1, 1]]
};
function tri(s, x, y, size, fill, corner) {
  poly(s, x, y, size, size, fill, TRI_PTS[corner]);
}

/** The stacked two-tone corner triangles used as page furniture. */
function triPair(s, a, b, size, corner) {
  tri(s, a[0], a[1], size, a[2], corner);
  tri(s, b[0], b[1], size, b[2], corner);
}

/** Polygon from normalised (0..1) points scaled into a box. */
function poly(s, x, y, w, h, fill, pts) {
  s.addShape(PRES.ShapeType.custGeom, {
    x, y, w, h, fill: { color: fill },
    points: pts.map(p => ({ x: +(p[0] * w).toFixed(4), y: +(p[1] * h).toFixed(4) })).concat([{ close: true }])
  });
}

const CHECK_PTS = [[0.35, 1], [0, 0.526], [0.087, 0.407], [0.35, 0.763], [0.913, 0], [1, 0.119]];
const CROSS_PTS = [[0.1, 1], [0, 0.9], [0.4, 0.5], [0, 0.1], [0.1, 0], [0.5, 0.4],
  [0.9, 0], [1, 0.1], [0.6, 0.5], [1, 0.9], [0.9, 1], [0.5, 0.6]];

/** Placeholder standing in for a raster image in the source deck. */
function imagePlaceholder(s, x, y, w, h, label) {
  s.addText(label || '[image]', {
    x, y, w, h, shape: PRES.ShapeType.rect,
    fill: { color: 'D9D9D9' }, line: { color: 'AFAFAF', width: 1 },
    align: 'center', valign: 'middle', fontFace: BODY, fontSize: 12, color: '6E6E6E'
  });
}

/** Hamburger menu glyph (three white bars). */
function hamburger(s, x, y) {
  for (let i = 0; i < 3; i++) rect(s, x, y + i * 0.075, 0.3, 0.028, C.white);
}

/** Rounded social buttons: LinkedIn, mail, Facebook. */
function socialTiles(s, x, y, vertical) {
  ['in', 'mail', 'f'].forEach((kind, i) => {
    const bx = vertical ? x : x + i * 0.472;
    const by = vertical ? y + i * 0.4645 : y;
    s.addShape(PRES.ShapeType.roundRect, {
      x: bx, y: by, w: 0.304, h: 0.304, rectRadius: 0.03, fill: { color: C.a2 }
    });
    if (kind === 'mail') {
      // envelope: white body with a purple fold across the top
      rect(s, bx + 0.055, by + 0.095, 0.194, 0.125, C.white);
      poly(s, bx + 0.055, by + 0.095, 0.194, 0.075, C.a2, [[0.08, 0], [0.92, 0], [0.5, 1]]);
    } else {
      s.addText(kind, {
        x: bx, y: by, w: 0.304, h: 0.304, align: 'center', valign: 'middle', margin: 0, wrap: false,
        fontFace: BODY, fontSize: 10, bold: true, color: C.white
      });
    }
  });
}

// ---------------------------------------------------------------------------
// Repeated slide furniture
// ---------------------------------------------------------------------------

/** Top nav used on the cover and the closing slide. */
function navBar(s, x0) {
  [['About Us', 0, 0.958], ['Team', 1.404, 0.63], ['Service', 2.481, 0.821],
   ['Portfolio', 3.749, 1.07], ['Contact', 5.266, 0.94]].forEach(([label, dx, w]) => {
    text(s, label, x0 + dx, 0.357, w, 0.286, { size: 11, color: C.white });
  });
}

/**
 * Cover / closing hero band. The source is one rectangle with a 4-stop black
 * gradient (85% → 20% alpha over white); pptxgenjs has no gradient fill, so it
 * is approximated with a strip of solid steps.
 */
const HERO_W = 12.3333;
const HERO_STEPS = 96;
const HERO_FROM = 0.15;   // 85% black over white
const HERO_TO = 0.80;     // 20% black over white

function heroGradient(s, x, reverse) {
  for (let i = 0; i < HERO_STEPS; i++) {
    let t = (i + 0.5) / HERO_STEPS;
    if (reverse) t = 1 - t;
    const g = Math.round(255 * (HERO_FROM + (HERO_TO - HERO_FROM) * t));
    const hex = ((g << 16) | (g << 8) | g).toString(16).padStart(6, '0').toUpperCase();
    rect(s, x + (i / HERO_STEPS) * HERO_W, 0, HERO_W / HERO_STEPS + 0.02, 7.5, hex);
  }
}

// ---------------------------------------------------------------------------
// Slide builders
// ---------------------------------------------------------------------------

function slide01(s) {                                    // Cover — "Paywise"
  heroGradient(s, 1.0);
  text(s, 'Paywise', 1.74, 1.873, 4.772, 1.447, { size: 80, bold: true, face: HEAD, color: C.white });
  text(s, 'Empowering People to Master Their Money', 1.74, 3.361, 4.532, 0.37,
    { size: 16, face: HEAD, color: C.white });

  rect(s, 0, 0, 1.0, 7.5, C.a6);
  hamburger(s, 0.35, 0.4);
  navBar(s, 6.513);

  // Left stat card
  rect(s, 0, 4.96, 3.123, 2.095, C.a5);
  text(s, 'Information', 0.338, 5.283, 2.446, 0.337,
    { size: 14, bold: true, face: HEAD, color: C.white, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, elita consectetur adipiscing, sed do',
    0.364, 5.611, 2.395, 0.62, { size: 11, color: C.white, align: 'center', lsp: 1.5 });
  socialTiles(s, 0.937, 6.472, false);

  // Middle stat card
  rect(s, 3.356, 4.96, 3.171, 2.095, C.a4);
  [[0, 0.30], [0.145, 0.44], [0.29, 0.58]].forEach(([dx, hh]) => {
    rect(s, 3.771 + dx, 5.976 - hh, 0.1, hh, C.white);
  });
  poly(s, 4.206, 5.62, 0.143, 0.356, C.white,
    [[0, 0], [1, 0], [1, 1], [0.62, 1], [0.62, 0.34], [0, 0.34]]);
  text(s, '$23.00', 4.503, 5.367, 1.608, 0.64, { size: 32, bold: true, face: HEAD, color: C.white });
  text(s, 'Lorem ipsum dolor sit amet, consectet adipiscing elit, sed do eiusmod',
    3.613, 6.172, 2.754, 0.62, { size: 11, color: C.white, align: 'center', lsp: 1.5 });

  // Right stat card (translucent black over the gradient + purple outline)
  rect(s, 6.759, 4.96, 3.171, 2.095, { color: C.black, transparency: 50 },
    { line: { color: C.a1, width: 2 } });
  text(s, '4,5K User', 7.235, 5.352, 2.22, 0.64,
    { size: 32, bold: true, face: HEAD, color: C.white, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectet adipiscing elit, sed do eiusmod tempor',
    6.968, 6.044, 2.754, 0.62, { size: 11, color: C.white, align: 'center', lsp: 1.5 });

  vRule(s, 12.985, 1.188, 4.415, 2.778, C.white);
}

const AGENDA = ['About', 'The Problem', 'The Solution', 'Product', 'Market',
  'Business Model', 'Strategy', 'Team', 'Financials', 'Closing'];

function slide02(s) {                                    // Agenda
  rect(s, 9.972, 4.139, 3.361, 3.361, C.a6);
  heading(s, 'Agenda', 1.482, 1.343, 2.305, 0.64);
  rect(s, 0, 2.633, 0.926, 3.415, C.a1);
  rect(s, 0.926, 2.633, 6.382, 3.415, C.a3);

  AGENDA.forEach((label, i) => {
    const col = i < 5 ? 0 : 1;
    const x = col === 0 ? 1.42 : 4.205;
    const y = 3.086 + (i % 5) * 0.5265;
    text(s, String(i + 1).padStart(2, '0'), x, y, 0.516, 0.404,
      { size: 18, bold: true, color: C.white, align: 'center' });
    text(s, label, x + 0.61, y, col === 0 ? 1.645 : 1.999, 0.404,
      { size: 18, bold: true, color: C.white });
  });

  hRule(s, 0.692, 0, 3.25, 1.601, C.a1);
}

function slide03(s) {                                    // About Our Company
  rect(s, 6.638, 0.778, 6.269, 5.944, C.a6);
  heading(s, 'About Out Company', 1.206, 1.755, 4.226, 0.64);
  lead(s, 'We are on a mission to make financial literacy simple, personal, and powerful.',
    1.206, 2.827, 4.226);
  para(s, 'Founded in 2024, [Startup Name] is a fintech startup based in Indonesia, built by a team of finance educators, product designers, and technology enthusiasts. We believe managing money shouldn\u2019t be complicated or intimidating.',
    1.206, 3.535, 4.226, 1.175);
  para(s, 'Our platform empowers individuals\u2014especially Gen Z and Millennials\u2014to take control of their finances through education, habit-forming tools, and personalized guidance.',
    1.206, 4.847, 4.226, 0.897);
  para(s, 'With a user-first approach and a strong community focus, we\u2019re redefining how people build wealth and financial confidence in their daily lives.',
    7.21, 5.793, 5.125, 0.62, { color: C.white, align: 'center' });

  rect(s, 12.907, 0.778, 0.426, 5.944, C.a3);
  vRule(s, 13.12, 2.125, 5.375, 3.726, C.white);
}

function slide04(s) {                                    // Quotes
  tri(s, 1.21, 1.809, 4.481, C.a5, 'bl');
  tri(s, 0, 1.21, 6.29, C.a6, 'bl');
  oval(s, 1.21, 1.21, 5.08, 5.08, C.a1);

  text(s, 'Quotes', 7.698, 1.684, 3.026, 1.111, { size: 60, bold: true, face: HEAD, color: C.a3 });
  text(s, '\u201D', 11.339, 1.709, 0.625, 1.717,
    { size: 96, bold: true, face: HEAD, color: C.a1, align: 'center' });
  text(s, 'The number one problem in today\'s generation and economy is the lack of financial literacy.',
    7.698, 3.232, 4.26, 1.717, { size: 24, italic: true, face: QUOTE });
  text(s, '- Alan Greenspan', 7.698, 5.445, 2.667, 0.37,
    { size: 16, bold: true, face: HEAD, color: C.a5 });
}

/** Slides 5–7 & 12 share this "headline + copy + lead + 3 bullets" layout. */
function problemLayout(s, cfg) {
  heading(s, cfg.title, cfg.x, cfg.titleY, cfg.titleW, cfg.titleH || 0.64);
  if (cfg.para) para(s, cfg.para, cfg.x, cfg.paraY, cfg.paraW, 0.897);
  if (cfg.lead) lead(s, cfg.lead, cfg.x, cfg.leadY, cfg.leadW);
  cfg.bullets.forEach((b, i) => {
    bullet(s, cfg.dotX, cfg.bulletX, cfg.bulletY + i * 0.375, cfg.bulletW,
      typeof b === 'string' ? b : b[0], { h: typeof b === 'string' ? 0.286 : b[1] });
  });
}

function slide05(s) {                                    // The Problem
  rect(s, 0, 0, 3.562, 7.5, C.a6);
  rect(s, 13.12, 0, 0.213, 7.5, C.a3);
  problemLayout(s, {
    title: 'The Problem', x: 7.998, titleY: 1.855, titleW: 2.81,
    para: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim adminim.',
    paraY: 2.694, paraW: 3.572,
    lead: 'Most people don\u2019t know how to manage their money:', leadY: 3.777, leadW: 3.46,
    dotX: 8.18, bulletX: 8.425, bulletY: 4.609, bulletW: 2.383,
    bullets: ['70% live paycheck to paycheck', 'No financial guidance', 'Poor financial habits']
  });
  vRule(s, 0.526, 1.573, 5.927, 3.718, C.white);
}

function slide06(s) {                                    // Current Solutions
  rect(s, 12.533, 0.406, 0.426, 6.688, C.a3);
  problemLayout(s, {
    title: 'Current Solutions', x: 1.362, titleY: 1.818, titleW: 3.783,
    para: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim adminim.',
    paraY: 2.637, paraW: 3.887,
    lead: 'Existing finance apps are fragmented and overly technical:', leadY: 3.918, leadW: 3.596,
    dotX: 1.543, bulletX: 1.789, bulletY: 4.646, bulletW: 2.679,
    bullets: ['Complex interfaces', 'No personalized advice', 'Tracking-focused, not habit-forming']
  });
  triPair(s, [0, 6.365, C.a6], [0.161, 6.203, C.a5], 1.135, 'bl');
  triPair(s, [0, 0, C.a6], [0.161, 0.161, C.a5], 1.135, 'tl');
}

function slide07(s) {                                    // Our Solution
  rect(s, 0, 1.01, 5.385, 5.479, C.a6);
  problemLayout(s, {
    title: 'Our Solution', x: 8.15, titleY: 1.839, titleW: 3.783,
    para: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim adminim.',
    paraY: 2.679, paraW: 3.887,
    lead: 'An integrated personal finance platform that\u2019s smart, simple, and habit-driven.',
    leadY: 3.866, leadW: 3.887,
    dotX: 8.332, bulletX: 8.577, bulletY: 4.625, bulletW: 3.356,
    bullets: ['Financial education + management in one app', 'Personalized financial planning',
      'Built for real people, not just finance geeks']
  });
  hRule(s, 0.505, 0, 3.25, 1.601, C.a1);
  hRule(s, 6.995, 0, 3.25, 1.601, C.a1);
}

/** Stand-in for the laptop product shot: bezel, white screen, silver base. */
function laptopPlaceholder(s, x, y, w, h) {
  const bodyH = h * 0.83;
  rect(s, x + w * 0.075, y, w * 0.85, bodyH, '2B2B2B');
  rect(s, x + w * 0.10, y + bodyH * 0.045, w * 0.80, bodyH * 0.90, C.white);
  rect(s, x, y + bodyH, w, h - bodyH, 'C6C8CC');
  s.addText('[image]', { x: x + w * 0.10, y: y + bodyH * 0.40, w: w * 0.80, h: 0.3,
    align: 'center', valign: 'middle', margin: 0, fontFace: BODY, fontSize: 12, color: '9A9A9A' });
}

/** Stand-in for the phone product shot. */
function phonePlaceholder(s, x, y, w, h) {
  s.addShape(PRES.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.16, fill: { color: '2B2B2B' } });
  s.addShape(PRES.ShapeType.roundRect, {
    x: x + w * 0.06, y: y + h * 0.03, w: w * 0.88, h: h * 0.94, rectRadius: 0.13,
    fill: { color: '1B1B2A' }
  });
}

function slide08(s) {                                    // Product Demo
  tri(s, 7.043, 0, 6.29, C.a5, 'tr');
  tri(s, 7.043, 1.21, 6.29, C.a6, 'br');
  laptopPlaceholder(s, 5.257, 1.772, 6.575, 3.956);
  phonePlaceholder(s, 11.002, 2.361, 1.651, 3.367);
  problemLayout(s, {
    title: 'Product Demo', x: 1.098, titleY: 2.394, titleW: 3.597,
    para: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim adminim.',
    paraY: 3.277, paraW: 3.597,
    dotX: 1.269, bulletX: 1.525, bulletY: 4.445, bulletW: 3.17,
    bullets: ['Easy-to-use dashboard', 'Budgeting, saving, and goal tracking tools']
  });
}

function slide09(s) {                                    // Unique Value Proposition
  rect(s, 0, 7.229, 13.333, 0.271, C.a4);
  rect(s, 0, 0, 13.333, 0.271, C.a4);
  rect(s, 0, 1.323, 3.583, 4.854, C.a3);

  heading(s, 'Unique Value Proposition', 5.367, 1.637, 5.578, 0.64);
  lead(s, 'We don\u2019t just help people track their money \u2014 we help them transform their financial behavior.',
    5.367, 2.566, 5.578);
  text(s, 'Our platform is designed to make financial wellness:', 5.367, 3.428, 5.578, 0.337,
    { size: 14, bold: true, face: HEAD });

  [['Accessible', ' \u2014 Easy-to-understand tools, built for everyone, not just finance experts.'],
   ['Actionable', ' \u2014 Step-by-step plans and nudges that turn insights into daily financial wins.'],
   ['Automatic', ' \u2014 Smart systems that build good habits, reduce manual work, and create momentum.']
  ].forEach(([word, rest], i) => {
    const y = 3.947 + i * 0.375;
    dot(s, 5.548, y + 0.069);
    runs(s, [{ t: word, c: C.black }, { t: rest }], 5.794, y, 6.575, 0.286);
  });

  para(s, 'We meet users where they are \u2014 whether they\u2019re just getting started or trying to level up \u2014 and empower them with a platform that grows with their goals.',
    5.367, 5.243, 7.196, 0.62);

  rect(s, 0, 0, 13.333, 0.223, C.a6);
  rect(s, 0, 7.277, 13.333, 0.223, C.a6);
}

const KEY_FEATURES = [
  ['AI-Powered Financial Insights:', 'Personalized financial recommendations based on user behavior and goals.'],
  ['Budgeting & Expense Tracking:', 'Easy-to-use tools to track spending, set budgets, and manage financial goals.'],
  ['Automated Savings Plans:', 'Automatically set aside money for savings and investments with minimal effort.'],
  ['Debt Management Tools:', 'Track loans, set repayment goals, and create a debt reduction plan.'],
  ['Financial Education Hub:', 'Access to articles, webinars, and tutorials to improve financial literacy.'],
  ['Mobile-First Design:', 'Seamless experience across all devices, with a focus on mobile usability.']
];

function slide10(s) {                                    // Key Features
  rect(s, 9.354, 1.854, 3.583, 5.245, C.a4);
  rect(s, 9.354, 0.401, 3.583, 5.245, C.a6);
  heading(s, 'Key Features', 0.85, 0.955, 6.279, 0.64);
  lead(s, 'Our platform offers innovative features designed to simplify financial wellness and empower users to take control of their financial future.',
    0.85, 1.784, 6.279);
  KEY_FEATURES.forEach(([title, body], i) => {
    bulletBlock(s, 1.031, 1.308, 2.607 + i * 0.6635, 5.477, 0.62, title, [{ t: body }]);
  });
}

function slide11(s) {                                    // Join us
  rect(s, 0, 0, 4.375, 7.5, C.a6);
  rect(s, 0.542, 0.542, 4.708, 6.417, C.a1);
  text(s, 'Join us in building a financially stronger generation.', 6.172, 1.803, 6.24, 2.524,
    { size: 48, bold: true, face: HEAD, color: C.a3 });
  text(s, 'Partner with us to make money management a daily habit.', 6.172, 4.788, 6.24, 0.909,
    { size: 24, face: QUOTE });
}

function slide12(s) {                                    // Market Opportunity
  rect(s, 4.604, 0.955, 2.695, 5.589, C.a6);
  rect(s, 0, 0, 0.955, 0.955, C.a1);
  rect(s, 0, 6.545, 0.955, 0.955, C.a3);
  problemLayout(s, {
    title: 'Market Opportunity', x: 8.399, titleY: 1.44, titleW: 3.783, titleH: 1.178,
    lead: 'Indonesia\u2019s rising middle class = massive demand for financial empowerment.',
    leadY: 2.883, leadW: 3.887,
    dotX: 8.58, bulletX: 8.826, bulletY: 3.715, bulletW: 3.46,
    bullets: ['TAM: $X Billion', '190M+ smartphone users',
      ['70M+ young adults in need of financial education', 0.471]]
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim adminim.',
    8.347, 5.163, 3.939, 0.897);
}

function slide13(s) {                                    // Testimonials
  heading(s, 'Testimonials', 4.554, 0.828, 4.226, 0.64, C.black, 'center');
  rect(s, 0.716, 2.684, 2.151, 2.151, C.a6);
  rect(s, 6.924, 2.684, 2.151, 2.151, C.a5);

  [[3.344, 'Ricard Well', '\u201CThis is the first app that helped me actually understand my money.\u201D', '\u2014 Beta User, 24 y.o., Jakarta'],
   [9.553, 'Albert Grandam', '\u201CI finally saved my first 10M IDR thanks to the challenge system!\u201D', '\u2014 Freelancer, Bandung']
  ].forEach(([x, name, quote, who]) => {
    text(s, name, x, 2.932, 2.731, 0.37, { size: 16, bold: true, face: HEAD, color: C.a5 });
    text(s, quote, x, 3.421, 3.19, 0.761, { size: 14, color: C.gray, lsp: 1.5 });
    text(s, who, x, 4.452, 2.731, 0.303, { size: 12, bold: true, face: HEAD, color: C.a1 });
  });

  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip.',
    2.75, 6.135, 7.833, 0.62, { align: 'center' });

  hRule(s, 2.076, 5.042, 8.292, 6.643, C.a1);
  hRule(s, 5.568, 8.292, 5.042, 6.691, C.a1);
  triPair(s, [0, 0, C.a6], [0.161, 0.161, C.a5], 1.135, 'tl');
  triPair(s, [12.198, 6.365, C.a3], [12.036, 6.203, C.a1], 1.135, 'br');
}

const MARKET_TRENDS = [
  ['Rising Demand for Financial Literacy:', '76% of millennials seek financial education, with increasing interest in digital financial tools.'],
  ['Growth in Personal Finance Apps:', 'The global personal finance app market is expected to grow at a CAGR of 20% over the next 5 years.'],
  ['Shift to Digital Financial Services:', 'Consumers are moving away from traditional banking, with digital-first financial services gaining traction.'],
  ['Automation & AI in Finance:', '65% of users prefer automated savings and investment tools powered by AI for ease and efficiency.'],
  ['Focus on Financial Wellness:', 'Companies and individuals are prioritizing financial well-being, with wellness solutions seeing a surge in adoption.']
];

function slide14(s) {                                    // Market Trends
  rect(s, 0, 3.885, 3.615, 3.615, C.a5);
  heading(s, 'Market Trends', 0.987, 0.97, 3.931, 0.64);
  para(s, 'The financial wellness market is experiencing rapid growth, driven by increasing demand for accessible and personalized financial solutions.',
    0.987, 1.789, 4.279, 0.897);
  MARKET_TRENDS.forEach(([title, body], i) => {
    bulletBlock(s, 6.592, 6.869, 0.97 + i * 1.1655, 5.477, 0.897, title, [{ t: body }]);
  });
  hRule(s, 0.514, 13.333, 10.083, 11.732, C.a1);
}

const COMPARE_ROWS = [
  ['Personalized Financial Planning', 1, 0, 0, 1],
  ['Bahvioral Habit Coaching', 1, 0, 1, 0],
  ['Gamified Financial Challenges', 1, 0, 1, 0],
  ['Micro-learning Education', 1, 0, 1, 0],
  ['AI-Powered Recommendation', 1, 1, 0, 1],
  ['Simplicity & User-Friendliness', 1, 1, 1, 0]
];

function slide15(s) {                                    // Competitive Analysis
  s.addText('Competitive Analysis', {
    x: 3.583, y: 0.641, w: 6.167, h: 0.64, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 32, bold: true, color: C.black,
    margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
  para(s, 'In a crowded market, we stand out by focusing on what truly matters: long-term financial behavior change.',
    4.214, 1.461, 4.905, 0.62, { align: 'center' });
  text(s, 'How We Compare to the Competition', 1.209, 2.662, 5.477, 0.337, { size: 14, bold: true });

  const header = ['Features', 'Startup Name', 'Competitor A', 'Competitor B', 'Competitor C'];
  const rows = [header.map((t, i) => ({
    text: t, options: { bold: true, color: C.white, fill: { color: C.a1 }, fontSize: 14, align: 'center' }
  }))];
  COMPARE_ROWS.forEach((r, ri) => {
    const band = ri % 2 === 0 ? C.band1 : C.band2;
    rows.push([{ text: r[0], options: { fontSize: 11, align: 'left', fill: { color: band } } }]
      .concat([1, 2, 3, 4].map(() => ({ text: '', options: { fill: { color: band } } }))));
  });
  s.addTable(rows, {
    x: 1.209, y: 3.225, w: 10.916, colW: [2.775, 2.032, 2.036, 2.047, 2.026],
    rowH: 0.501, valign: 'middle', fontFace: BODY, fontSize: 11, color: C.black,
    border: { type: 'solid', color: C.white, pt: 1 }
  });

  // Tick / cross marks drawn over the table body.
  COMPARE_ROWS.forEach((r, ri) => {
    [1, 2, 3, 4].forEach(ci => {
      const cx = [4.858, 6.896, 8.933, 10.97][ci - 1];
      const y = 3.88 + ri * 0.5015;
      if (r[ci]) poly(s, cx, y, 0.275, 0.203, C.a4, CHECK_PTS);
      else poly(s, cx + 0.057, y - 0.016, 0.236, 0.236, C.a5, CROSS_PTS);
    });
  });

  triPair(s, [0, 0, C.a6], [0.161, 0.161, C.a5], 1.135, 'tl');
  triPair(s, [12.198, 0, C.a6], [12.036, 0.161, C.a5], 1.135, 'tr');
}

const REVENUE_LEFT = [
  ['Freemium Model', [{ t: 'Offering a ' }, { t: 'free version', b: 1 },
    { t: ' with basic features, and ' }, { t: 'premium subscriptions', b: 1 },
    { t: ' for advanced tools, personalized financial planning, and AI-driven insights.' }]],
  ['Subscription Fees', [{ t: 'Monthly or annual subscriptions for ' }, { t: 'individual users', b: 1 },
    { t: ' and ' }, { t: 'businesses', b: 1 },
    { t: ' (e.g., companies offering the platform to employees as a financial wellness benefit).' }]],
  ['In-App Purchases', [{ t: 'Users can purchase additional features such as ' },
    { t: 'customized financial reports', b: 1 }, { t: ', ' }, { t: 'exclusive content', b: 1 },
    { t: ', and ' }, { t: 'financial coaching services', b: 1 }, { t: '.' }]]
];

const REVENUE_RIGHT = [
  ['Partnerships & B2B Solutions', [{ t: 'Collaborations with ' }, { t: 'banks', b: 1 },
    { t: ', ' }, { t: 'financial institutions', b: 1 }, { t: ', and ' },
    { t: 'educational organizations', b: 1 },
    { t: ' to integrate our platform into their services, generating licensing and partnership fees.' }]],
  ['Affiliate Marketing', [{ t: 'Earning commissions by partnering with financial service providers (e.g., investment platforms, insurance companies) to offer tailored product recommendations within the app.' }]]
];

function slide16(s) {                                    // Business Model
  rect(s, 7.494, 0.5, 5.11, 2.255, C.a6);
  heading(s, 'Business Model', 0.972, 1.127, 5.361, 0.64);
  para(s, 'Our business model combines multiple revenue streams to drive growth, enhance user engagement, and ensure long-term profitability.',
    0.972, 1.857, 5.486, 0.62);
  lead(s, 'Revenue Streams:', 0.972, 2.953, 5.486, 0.337);

  REVENUE_LEFT.forEach(([title, parts], i) => {
    bulletBlock(s, 1.153, 1.43, 3.429 + i * 1.0305, 5.477, 0.897, title + ':', parts);
  });
  REVENUE_RIGHT.forEach(([title, parts], i) => {
    bulletBlock(s, 7.419, 7.696, 3.649 + i * 1.342, 4.665, 1.175, title + ':', parts);
  });

  hRule(s, 0.5, 0, 3.25, 1.601, C.a4);
  rect(s, 1.994, 7.292, 9.346, 0.208, C.a3);
}

function slide17(s) {                                    // Go-to-Market
  rect(s, 0.469, 0.312, 2.021, 6.875, C.a1);
  heading(s, 'Go-to-Market', 6.29, 0.926, 3.092, 0.64);
  para(s, 'Our go-to-market strategy focuses on targeting the right audience with the right message at the right time, ensuring rapid user adoption and sustainable growth.',
    6.29, 1.656, 6.086, 0.62);

  lead(s, 'Target Market:', 6.29, 2.689, 5.486, 0.337);
  [[{ t: 'Gen Z & Millennials', b: 1, c: C.black }, { t: ' ' },
    { t: '(18-35) seeking financial literacy and management tools.' }],
   [{ t: 'Small businesses & freelancers', b: 1, c: C.black }, { t: ' ' },
    { t: 'needing simple financial solutions.' }]
  ].forEach((parts, i) => {
    const y = 3.246 + i * 0.388;
    dot(s, 6.472, y + 0.069);
    runs(s, parts, 6.749, y, 5.477, 0.286);
  });

  lead(s, 'Phases:', 6.29, 4.285, 5.486, 0.337);
  const phases = [
    [4.793, 0.62, [{ t: 'Launch (Phase 1)', b: 1, c: C.black }, { t: ': ', c: C.black },
      { t: 'Organic growth through ' }, { t: 'content marketing', b: 1 }, { t: ', ' },
      { t: 'influencer partnerships', b: 1 }, { t: ', and ' }, { t: 'referral programs', b: 1 }, { t: '.' }]],
    [5.495, 0.342, [{ t: 'User Acquisition (Phase 2): ', b: 1, c: C.black },
      { t: 'Targeted ads', b: 1 }, { t: ' on social media and ' }, { t: 'SEO-driven', b: 1 }, { t: ' content.' }]],
    [5.954, 0.62, [{ t: 'Scaling (Phase 3): ', b: 1, c: C.black },
      { t: 'Expansion to Southeast Asia', b: 1 }, { t: ' and partnerships with ' }, { t: 'banks', b: 1 },
      { t: ' and ' }, { t: 'educational institutions', b: 1 }, { t: '.' }]]
  ];
  phases.forEach(([y, h, parts], i) => {
    text(s, String(i + 1).padStart(2, '0') + '.', 6.329, [4.843, 5.523, 6.004][i], 0.435, 0.286,
      { bold: true, face: HEAD, color: C.a3, align: 'center' });
    runs(s, parts, 6.749, y, 5.653, h, { lsp: 1.5 });
  });

  // Key metrics card
  rect(s, 0.932, 4.396, 4.427, 2.355, C.a6);
  lead(s, 'Key Metrics:', 1.412, 4.694, 3.382, 0.337, C.white);
  [[{ t: '100K active users', b: 1, c: C.white }, { t: ' by end of Year 1.', c: C.white }],
   [{ t: '5-8% paid conversion', b: 1, c: C.white }, { t: ' from free users.', c: C.white }],
   [{ t: '40% increase', b: 1, c: C.white }, { t: ' in social media engagement.', c: C.white }]
  ].forEach((parts, i) => {
    const y = 5.251 + i * 0.398;
    dot(s, 1.531, y + 0.069, C.a1);
    runs(s, parts, 1.808, y, 3.07, 0.286);
  });

  hRule(s, 1.246, 9.691, 12.191, 10.923, C.a1);
}

const MARKETING = [
  ['Digital Marketing', [{ t: 'Focused on ' }, { t: 'performance-driven', b: 1 },
    { t: ' channels such as social media ads, search engine marketing, and influencer partnerships to increase brand visibility and user acquisition.' }]],
  ['Content Marketing', [{ t: 'Creating high-quality content focused on ' }, { t: 'financial education', b: 1 },
    { t: ', including blogs, webinars, and video tutorials, to engage and educate our target audience.' }]],
  ['Referral Program', [{ t: 'Launching a ' }, { t: 'referral program', b: 1 },
    { t: ' to encourage existing users to refer friends and family in exchange for premium features or discounts.' }]],
  ['Partnerships', [{ t: 'Partnering with ' }, { t: 'financial institutions', b: 1 }, { t: ' and ' },
    { t: 'educational organizations', b: 1 }, { t: ' to expand our reach and credibility.' }]],
  ['Community Engagement', [{ t: 'Building a loyal community through ' }, { t: 'online forums', b: 1 },
    { t: ', user support groups, and customer success initiatives, fostering long-term relationships.' }]]
];

function slide18(s) {                                    // Marketing Strategy
  heading(s, 'Marketing Strategy', 0.921, 0.97, 4.138, 0.64);
  para(s, 'Our marketing strategy is designed to build brand awareness, drive user acquisition, and convert users into loyal customers.',
    0.921, 1.835, 4.621, 0.62);
  MARKETING.forEach(([title, parts], i) => {
    bulletBlock(s, 6.218, 6.495, 0.97 + i * 1.1655, 5.964, 0.897, title + ':', parts);
  });
  rect(s, 0, 0, 0.198, 2.26, C.a5);
  rect(s, 0, 7.302, 2.26, 0.198, C.a6);
  rect(s, 11.073, 0, 2.26, 0.198, C.a1);
  rect(s, 13.135, 5.24, 0.198, 2.26, C.a3);
}

const SALES = [
  ['Freemium Model', 2.427, [{ t: 'Offering a ' }, { t: 'free tier', b: 1 },
    { t: ' to attract users, with the option to upgrade to premium plans for advanced features and personalized services.' }]],
  ['B2B Partnerships', 3.531, [{ t: 'Targeting ' }, { t: 'business partnerships', b: 1 },
    { t: ' with financial service providers and educational institutions to offer our platform to their employees/students as a value-add service.' }]],
  ['Upselling & Cross-selling', 4.634, [{ t: 'Focused on converting users to premium subscriptions by showcasing the benefits of personalized financial planning tools and advanced insights.' }]],
  ['Data-Driven Sales', 5.738, [{ t: 'Leveraging ' }, { t: 'customer data', b: 1 },
    { t: ' to segment our audience and deliver tailored sales pitches, ensuring higher conversion rates for our premium offerings.' }]]
];

/** Outline icons beside each sales-strategy item (banknote, coin, chart, checklist). */
function salesIcon(s, x, y, kind) {
  const ln = { color: C.a2, width: 1.5 };
  if (kind === 'Freemium Model') {
    // stacked banknotes
    s.addShape(PRES.ShapeType.rect, { x: x + 0.06, y, w: 0.36, h: 0.24, fill: { color: C.white }, line: ln });
    s.addShape(PRES.ShapeType.rect, { x, y: y + 0.09, w: 0.36, h: 0.24, fill: { color: C.white }, line: ln });
    oval(s, x + 0.14, y + 0.16, 0.09, 0.09, C.a2);
  } else if (kind === 'B2B Partnerships') {
    s.addShape(PRES.ShapeType.ellipse, { x, y, w: 0.38, h: 0.38, fill: { color: C.white }, line: ln });
    s.addText('$', { x, y, w: 0.38, h: 0.38, align: 'center', valign: 'middle', margin: 0,
      fontFace: BODY, fontSize: 11, bold: true, color: C.a2 });
  } else if (kind === 'Upselling & Cross-selling') {
    [[0, 0.12], [0.12, 0.20], [0.24, 0.30], [0.36, 0.38]].forEach(([dx, hh]) => {
      rect(s, x + 0.02 + dx, y + 0.38 - hh, 0.07, hh, C.a2);
    });
  } else {
    // clipboard with a tick
    s.addShape(PRES.ShapeType.rect, { x: x + 0.04, y: y + 0.05, w: 0.31, h: 0.33, fill: { color: C.white }, line: ln });
    rect(s, x + 0.13, y, 0.13, 0.08, C.a2);
    poly(s, x + 0.11, y + 0.16, 0.17, 0.13, C.a2, CHECK_PTS);
  }
}

function slide19(s) {                                    // Sales Strategy
  heading(s, 'Sales Strategy', 0.865, 0.865, 5.361, 0.64);
  para(s, 'Our sales strategy is designed to build brand awareness, drive user acquisition, and convert users into loyal customers.',
    0.865, 1.595, 4.667, 0.62);
  SALES.forEach(([title, y, parts]) => {
    s.addText([{ text: title + ':', options: { bold: true, color: C.black } }].concat(
      parts.map((p, i) => ({ text: p.t, options: { bold: !!p.b, color: C.gray, softBreakBefore: i === 0 } }))
    ), {
      x: 1.729, y, w: 5.799, h: 0.897, valign: 'top', fontFace: BODY, fontSize: 11,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
    });
    salesIcon(s, 1.02, y + 0.24, title);
  });
  rect(s, 8.393, 0.865, 4.94, 5.771, C.a6);
  hRule(s, 7.068, 12.488, 9.238, 10.887, C.a1);
  hRule(s, 0.432, 9.238, 12.488, 10.839, C.a1);
}

const MILESTONES_KEY = [
  [[{ t: 'Q1 2024', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: 'MVP Launch', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Successfully launched to early adopters, gathering valuable feedback.' }]],
  [[{ t: 'Q2 2024', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: '10K Users', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Reached our first 10K active users through targeted campaigns.' }]],
  [[{ t: 'Q4 2024', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: 'Seed Funding', b: 1, c: C.black },
    { t: ':', c: C.black }, { t: ' Raised $X million to scale operations and development.' }]]
];

const MILESTONES_NEXT = [
  [0.62, [{ t: 'Q1 2025', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: '100K Users', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Scaling user base through increased marketing and product enhancements.' }]],
  [0.62, [{ t: 'Q2 2025', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: 'Launch Premium Features', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Adding personalized financial tools and advanced AI insights.' }]],
  [0.62, [{ t: 'Q4 2025', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: 'Expansion to SEA', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Targeting Southeast Asia, aiming for 500K users.' }]],
  [0.286, [{ t: '2026', b: 1, c: C.black }, { t: ' \u2013 ', c: C.black }, { t: 'Profitability', b: 1, c: C.black },
    { t: ': ', c: C.black }, { t: 'Achieving breakeven and scaling to 1M+ users.' }]]
];

function slide20(s) {                                    // Milestones & Roadmap
  s.addText('Milestones & Roadmap', {
    x: 4.073, y: 0.464, w: 5.188, h: 0.64, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 32, bold: true, color: C.black,
    margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
  para(s, 'Our journey is guided by clear milestones, each step bringing us closer to our vision of empowering users to take control of their financial future.',
    0.931, 1.511, 5.308, 0.62);

  lead(s, 'Key Milestones:', 0.931, 2.472, 5.486, 0.337);
  MILESTONES_KEY.forEach(([parts], i) => {
    const y = 2.939 + i * 0.7315;
    dot(s, 1.112, y + 0.235);
    runs(s, parts, 1.464, y, 4.942, 0.62, { lsp: 1.5 });
  });

  lead(s, 'Next Milestones:', 6.916, 2.075, 5.486, 0.337);
  [2.541, 3.272, 4.004, 4.735].forEach((y, i) => {
    const [h, parts] = MILESTONES_NEXT[i];
    dot(s, 7.098, y + (h > 0.4 ? 0.236 : 0.069));
    runs(s, parts, 7.449, y, 4.942, h, { lsp: h > 0.4 ? 1.5 : undefined });
  });

  hRule(s, 0.783, 9.902, 12.402, 11.134, C.a5);
  hRule(s, 0.783, 3.431, 0.931, 2.199, C.a5);
  rect(s, 0, 5.635, 0.876, 1.865, C.a6);
  rect(s, 12.457, 5.635, 0.876, 1.865, C.a5);
}

const TEAM = [
  { x: 0.427, y: 2.321, align: 'right', avatar: [5.042, 2.073, C.a3], name: 'Adam Burkman, ', nameColor: C.a3,
    role: 'CEO & Co-Founder',
    bio: 'Experienced fintech product leader with over 10 years in the finance and tech industries. Previously led product teams at [Company], driving growth and innovation in personal finance apps.' },
  { x: 8.469, y: 3.38, align: 'left', avatar: [6.667, 3.132, C.a4], name: 'Noah Roberts', nameColor: C.a4,
    sep: ', ', role: 'CTO & Co-Founder',
    bio: 'A data scientist and AI expert with a proven track record in building scalable tech solutions. Previously worked at [Company], building AI-driven financial products that improved user engagement by 30%.' },
  { x: 0.427, y: 4.439, align: 'right', avatar: [5.042, 4.191, C.a5], name: 'Liam Anderson, ', nameColor: C.a5,
    role: 'Head of Product',
    bio: 'A seasoned product manager with 8+ years of experience in building user-centric digital products. Specializes in UX/UI design and behavioral product design for financial apps.' },
  { x: 8.469, y: 5.498, align: 'left', avatar: [6.667, 5.25, C.a6], name: 'Amelia Wright', nameColor: C.a6,
    sep: ', ', role: 'Head of Marketing',
    bio: 'A digital marketing leader with 7+ years of experience in growth marketing, brand strategy, and user acquisition. Has worked with top fintech startups to scale user bases rapidly.' }
];

function slide21(s) {                                    // Meet the Team
  heading(s, 'Meet the Team', 2.214, 0.712, 3.219, 0.64);
  para(s, 'Our team is a blend of experienced leaders and passionate innovators, committed to making financial wellness accessible to everyone.',
    5.839, 0.698, 5.281, 0.62);

  TEAM.forEach(m => {
    oval(s, m.avatar[0], m.avatar[1], 1.625, 1.625, m.avatar[2]);
    const parts = [{ text: m.name, options: { bold: true, color: m.nameColor, fontSize: 12, fontFace: HEAD } }];
    if (m.sep) parts.push({ text: m.sep, options: { bold: true, color: C.a1, fontSize: 12, fontFace: HEAD } });
    parts.push({ text: m.role, options: { bold: true, color: C.black, fontSize: 12 } });
    parts.push({ text: m.bio, options: { color: C.gray, fontSize: 10, softBreakBefore: true } });
    s.addText(parts, {
      x: m.x, y: m.y, w: 4.438, h: 1.128, valign: 'top', align: m.align,
      fontFace: BODY, fontSize: 12, lineSpacingMultiple: 1.5,
      margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
    });
  });

  triPair(s, [0, 6.068, C.a6], [0.204, 5.865, C.a5], 1.432, 'bl');
  triPair(s, [11.902, 0, C.a3], [11.698, 0.204, C.a1], 1.432, 'tr');
}

const ADVISORS = [
  { x: 5.281, y: 1.761, w: 4.57, bioH: 0.573, name: 'Adam Burkman', role: 'Senior Advisor (Fintech & Finance)',
    bioY: 2.593, avatar: [3.482, 1.693],
    bio: 'Former CFO at [Big Company], with extensive experience in financial modeling, strategic planning, and scaling businesses in the fintech space.' },
  { x: 2.518, y: 3.737, w: 4.29, bioH: 0.825, name: 'Noah Roberts ', role: 'Technology Advisor',
    bioY: 4.569, avatar: [0.719, 3.796],
    bio: 'Co-Founder of [Successful Tech Company], a renowned expert in AI, blockchain, and mobile fintech solutions. Played a key role in the development of [Tech Product], which was adopted by millions.' },
  { x: 9.175, y: 3.737, w: 3.438, bioH: 0.825, name: 'Liam Anderson', role: 'Business Strategy Advisor',
    bioY: 4.569, avatar: [7.376, 3.796],
    bio: 'Ex-Consultant at [Top Consultancy Firm], specializing in business expansion, market entry strategies, and scaling startups globally.' }
];

function slide22(s) {                                    // Advisory board
  rect(s, 0, 6.0, 13.333, 1.449, C.a4);
  rect(s, 0, 6.051, 13.333, 1.449, C.a6);
  heading(s, 'Advisory board', 2.2, 0.57, 3.418, 0.64);
  para(s, 'We are proud to have an exceptional advisory board that brings deep expertise in finance, technology, and business strategy.',
    5.852, 0.556, 5.281, 0.62);

  ADVISORS.forEach(a => {
    s.addText([
      { text: a.name, options: { bold: true, color: C.a1, fontSize: 14, fontFace: HEAD, breakLine: true } },
      { text: a.role, options: { bold: true, color: C.black, fontSize: 12 } }
    ], {
      x: a.x, y: a.y, w: 3.061, h: 0.717, valign: 'top', fontFace: BODY, fontSize: 14,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
    });
    text(s, a.bio, a.x, a.bioY, a.w, a.bioH, { size: 10, color: C.gray, lsp: 1.5 });
  });

  s.addText([
    { text: 'Strategic Partners:', options: { bold: true, breakLine: true } },
    { text: '[add Bank/Investor Name]' }
  ], {
    x: 0.621, y: 6.489, w: 2.49, h: 0.572, valign: 'top', fontFace: BODY, fontSize: 14,
    color: C.white, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
  text(s, '[add university/institution Name]', 9.535, 6.607, 3.177, 0.337,
    { size: 14, bold: true, color: C.white, align: 'right' });
}

const KEY_METRICS = [
  ['User Growth:', 2.507, [
    [2.923, [{ t: 'Monthly Active Users (MAU)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '100,000 (Projected for Q2 2025)' }]],
    [3.286, [{ t: 'Total Users', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '500,000 (Year-to-date)' }]],
    [3.65, [{ t: 'User Retention Rate', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '85% (Monthly)' }]]]],
  ['Customer Acquisition Cost (CAC):', 4.239, [
    [4.655, [{ t: 'Current CAC', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$10' }]],
    [5.018, [{ t: 'Projected CAC (6 months)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$8' }]]]],
  ['Lifetime Value (LTV):', 5.608, [
    [6.023, [{ t: 'LTV', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$120 per user (avg.)' }]],
    [6.387, [{ t: 'LTV to CAC Ratio', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '12:1' }]]]]
];

function slide23(s) {                                    // Key Metrics
  rect(s, 10.146, 0, 3.187, 7.5, C.a3);
  rect(s, 7.969, 0.391, 4.906, 6.719, C.a6);
  heading(s, 'Key Metrics', 0.903, 0.827, 4.882, 0.64);
  para(s, 'Our growth and success are driven by key metrics that showcase traction, engagement, and the financial health of our business.',
    0.903, 1.537, 4.882, 0.62);
  KEY_METRICS.forEach(([title, ty, items]) => {
    lead(s, title, 0.903, ty, 4.715, 0.337);
    items.forEach(([y, parts]) => {
      dot(s, 1.084, y + 0.069);
      runs(s, parts, 1.361, y, 4.424, 0.286);
    });
  });
}

const PERFORMANCE = [
  ['Revenue Growth:', 2.243, [
    [2.659, 4.424, [{ t: 'Current Revenue (YTD)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$300K' }]],
    [3.022, 4.424, [{ t: 'Projected Revenue (End of Year)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$1M' }]],
    [3.386, 4.424, [{ t: 'ARR (Annual Recurring Revenue)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$500K' }]]]],
  ['Conversion Rates:', 3.975, [
    [4.391, 4.424, [{ t: 'Free to Paid Conversion', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '5% (Target: 8% by end of Q3 2025)' }]],
    [4.754, 4.424, [{ t: 'Trial to Subscription', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '60%' }]]]],
  ['Key Takeaways:', 5.344, [
    [5.759, 4.424, [{ t: 'Strong user retention, with 85% of users actively engaged.' }]],
    [6.123, 5.264, [{ t: 'Exceptional LTV to CAC ratio, demonstrating efficient customer acquisition.' }]],
    [6.486, 5.264, [{ t: 'Solid revenue growth with projections to exceed $1M by year-end.' }]]]]
];

function slide24(s) {                                    // Key Metrics & Business Performance
  rect(s, 0, 0, 0.99, 7.5, C.a6);
  text(s, 'Key Metrics & Business Performance', 6.748, 0.728, 4.764, 1.178,
    { size: 32, bold: true, face: HEAD });
  PERFORMANCE.forEach(([title, ty, items]) => {
    lead(s, title, 6.748, ty, 4.715, 0.337);
    items.forEach(([y, w, parts]) => {
      dot(s, 6.93, y + 0.069);
      runs(s, parts, 7.207, y, w, 0.286);
    });
  });
}

const FORECAST = [
  ['Metric', 'Year 1', 'Year 2', 'Year 3'],
  ['Revenue', '$300K', '$1M', '$3M'],
  ['Expenses', '$400K', '$800K', '$2M'],
  ['EBITDA', '-$100K', '$200K', '$1M'],
  ['Net Profit', '-$150K', '$100K', '$500K']
];

function slide25(s) {                                    // Financial Projections
  [[0.047, C.a4], [0, C.a6]].forEach(([dy, c]) => rect(s, 0, dy, 13.333, 0.401, c));
  [[7.099, C.a4], [7.052, C.a6]].forEach(([y, c]) => rect(s, 0, y, 13.333, 0.401, c));

  s.addText('Financial Projections', {
    x: 3.583, y: 1.253, w: 6.167, h: 0.64, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 32, bold: true, color: C.black,
    margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
  para(s, 'Our financial growth is poised for strong acceleration, driven by user acquisition and a scalable business model.',
    4.042, 2.024, 5.25, 0.62, { align: 'center' });
  text(s, '3- Year Financial Forecast :', 0.76, 3.03, 2.736, 0.337, { size: 14, bold: true });

  const rows = FORECAST.map((r, ri) => r.map((cell, ci) => ({
    text: cell,
    options: ri === 0
      ? { bold: true, color: C.white, fill: { color: C.a1 }, fontSize: 14, align: ci === 0 ? 'left' : 'center' }
      : { fontSize: 11, align: ci === 0 ? 'left' : 'center', fill: { color: ri % 2 === 1 ? C.band1 : C.band2 } }
  })));
  s.addTable(rows, {
    x: 0.896, y: 3.587, w: 6.542, colW: [2.101, 1.475, 1.479, 1.486], rowH: 0.532,
    valign: 'middle', fontFace: BODY, fontSize: 11, color: C.black,
    border: { type: 'solid', color: C.white, pt: 1 }
  });

  lead(s, 'Key Assumptions:', 8.095, 4.036, 4.311, 0.337);
  [[4.451, 0.286, [{ t: 'User Growth', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '100K users by end of Year 1, 500K by Year 3' }]],
   [4.815, 0.286, [{ t: 'Conversion', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '5% free-to-paid conversion rate.' }]],
   [5.178, 0.62, [{ t: 'ARPU (Average Revenue Per User)', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$10 by Year 2, increasing with premium features.' }]]
  ].forEach(([y, h, parts]) => {
    dot(s, 8.277, y + (h > 0.4 ? 0.236 : 0.069));
    runs(s, parts, 8.554, y, 4.311, h, { lsp: h > 0.4 ? 1.5 : undefined });
  });
}

function slide26(s) {                                    // Funding Ask
  rect(s, 7.548, 0.615, 4.882, 6.104, C.a6);
  heading(s, 'Funding Ask', 0.903, 0.882, 4.882, 0.64);
  para(s, 'We are raising $500K in funding to accelerate our growth, drive user acquisition, and scale our platform.',
    0.903, 1.592, 4.882, 0.62);

  lead(s, 'Investment Opportunity:', 0.903, 2.562, 4.715, 0.337);
  [[2.977, 0.286, 4.424, [{ t: 'Amount Sought', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: '$500K' }]],
   [3.341, 0.286, 4.424, [{ t: 'Funding Type', b: 1, c: C.black }, { t: ': ', c: C.black }, { t: 'SAFE note with a $X valuation cap' }]]
  ].forEach(([y, h, w, parts]) => {
    dot(s, 1.084, y + 0.069);
    runs(s, parts, 1.361, y, w, h);
  });
  dot(s, 1.084, 3.773);
  bulletBlockRaw(s, 1.361, 3.704, 4.424, 1.175, 'Projected Impact:',
    [{ t: 'This investment will enable us to reach 500K users by the end of Year 3, scale our platform, and become a market leader in financial wellness.' }]);

  lead(s, 'Why Invest?', 0.903, 5.19, 4.523, 0.337);
  [[5.605, 5.149, [{ t: 'Proven ', b: 1, c: C.black }, { t: 'demand', b: 1 }, { t: ' for accessible financial education and management tools.' }]],
   [5.969, 4.424, [{ t: 'Scalable business model', b: 1, c: C.black }, { t: ' ', c: C.black }, { t: 'with high potential for profitability.' }]],
   [6.332, 4.424, [{ t: 'Strong growth metrics', b: 1, c: C.black }, { t: ' ', c: C.black }, { t: 'and a clear path to market leadership.' }]]
  ].forEach(([y, w, parts]) => {
    dot(s, 1.084, y + 0.069);
    runs(s, parts, 1.361, y, w, 0.286);
  });

  triPair(s, [12.367, 0, C.a3], [12.229, 0.137, C.a1], 0.967, 'tr');
  triPair(s, [12.367, 6.533, C.a6], [12.229, 6.396, C.a5], 0.967, 'br');
}

/** bulletBlock without its own dot (caller places it). */
function bulletBlockRaw(s, x, y, w, h, title, parts) {
  s.addText([{ text: title, options: { bold: true, color: C.black } }].concat(
    parts.map((p, i) => ({ text: p.t, options: { bold: !!p.b, color: C.gray, softBreakBefore: i === 0 } }))
  ), {
    x, y, w, h, valign: 'top', fontFace: BODY, fontSize: 11, lineSpacingMultiple: 1.5,
    margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
}

const USE_OF_FUNDS = [
  ['Product Development (40%):', C.a6, 'Enhancing platform features, optimizing AI algorithms, and scaling infrastructure to support rapid user growth.'],
  ['Marketing & User Acquisition (30%):', C.a5, 'Expanding brand awareness, influencer partnerships, and targeted campaigns to drive user acquisition across key demographics.'],
  ['Team Expansion (15%):', C.a3, 'Hiring additional talent in engineering, product, and customer support to ensure a smooth scaling process.'],
  ['Operations (15%):', C.a1, 'Strengthening internal operations, improving customer success, and ensuring efficient business growth.']
];

function slide27(s) {                                    // Use of Funds (bar chart)
  s.addChart(PRES.ChartType.bar, [{
    name: 'Series 1',
    labels: ['Best Data One', 'Best Data Two', 'Best Data Three', 'Best Data Four'],
    values: [40, 30, 15, 15]
  }], {
    x: 1.036, y: 0.506, w: 4.957, h: 6.488,
    chartColors: [C.a6, C.a5, C.a3, C.a1],
    barDir: 'col', barGapWidthPct: 80,
    showLegend: false, showTitle: false, showValue: false,
    catAxisHidden: true, valAxisMaxVal: 50, valAxisMinVal: 0,
    valAxisLabelFontSize: 12, valAxisLabelColor: '595959', valAxisLabelFontFace: BODY,
    valAxisLineShow: false, catAxisLineShow: false,
    valGridLine: { color: 'D9D9D9', size: 1, style: 'solid' },
    catGridLine: { style: 'none' }
  });

  heading(s, 'Use of Funds', 7.029, 0.818, 4.882, 0.64);
  para(s, 'We are seeking $500K in funding to accelerate our product development, marketing efforts, and team expansion. Here\u2019s how we plan to use the funds:',
    7.029, 1.528, 5.269, 0.62);
  USE_OF_FUNDS.forEach(([title, color, body], i) => {
    const y = 2.48 + i * 1.091;
    lead(s, title, 7.029, y, 4.324, 0.337, color);
    para(s, body, 7.029, y + 0.309, 5.269, 0.62);
  });

  rect(s, 0, 5.24, 0.198, 2.26, C.a5);
  rect(s, 0, 0, 2.26, 0.198, C.a6);
  rect(s, 11.073, 7.302, 2.26, 0.198, C.a1);
  rect(s, 13.135, 0, 0.198, 2.26, C.a3);
}

function slide28(s) {                                    // Closing quote
  triPair(s, [0.47, 0.369, C.a6], [0.631, 0.53, C.a5], 1.135, 'tl');
  triPair(s, [11.728, 5.996, C.a3], [11.567, 5.834, C.a1], 1.135, 'br');
  rect(s, 0.808, 0.707, 11.717, 6.086, { color: C.black, transparency: 30 });

  s.addText([
    { text: '\u201CWe don\u2019t just build tech. ', options: { breakLine: true } },
    { text: 'We build habits that last.\u201D' }
  ], {
    x: 2.98, y: 2.313, w: 7.374, h: 1.582, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 44, bold: true, color: C.white,
    margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true, fit: 'resize'
  });
  text(s, 'Let\u2019s create a future where money empowers, not controls.', 3.131, 4.278, 7.071, 0.909,
    { size: 24, face: QUOTE, color: C.white, align: 'center' });

  hRule(s, 0.369, 10.025, 12.525, 11.257, C.a1);
  hRule(s, 7.131, 0.47, 2.97, 1.701, C.a5);
}

const CONTACT = [
  ['Name', 'Albert Grandam'],
  ['Email', 'office@paywise.com'],
  ['Websites', 'www.paywise.com'],
  ['LinkedIn', 'Paywise Company']
];

function slide29(s) {                                    // Contact
  rect(s, 0, 0, 4.869, 7.5, C.a6);
  rect(s, 0.677, 1.051, 5.386, 5.399, C.a3);

  text(s, 'Contact', 1.552, 1.873, 2.9, 0.64, { size: 32, bold: true, face: HEAD, color: C.white });
  CONTACT.forEach(([label, value], i) => {
    const y = 2.919 + i * 0.4905;
    text(s, label, 1.582, y, 1.082, 0.337, { size: 14, bold: true, face: HEAD, color: C.white });
    text(s, value, 2.754, y + 0.025, 1.648, 0.286, { size: 11, color: C.white });
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod temp',
    1.552, 5.008, 2.9, 0.62, { color: C.white });

  hRule(s, 1.318, 1.091, 4.454, 2.748, C.a1);
  hRule(s, 6.136, 1.091, 4.454, 2.748, C.a1);

  rect(s, 5.03, 3.76, 2.081, 2.377, C.a1);
  text(s, 'QR Code', 5.53, 3.805, 1.082, 0.337,
    { size: 14, bold: true, face: HEAD, color: C.white, align: 'center' });
}

function slide30(s) {                                    // Thank You
  heroGradient(s, 0, true);
  rect(s, 12.333, 0, 1.0, 7.5, C.a6);
  hamburger(s, 12.683, 0.4);
  // Same nav as the cover, but the whole group is mirrored so the order reverses.
  [['Contact', 5.758, 0.94], ['Portfolio', 7.145, 1.07], ['Service', 8.662, 0.821],
   ['Team', 9.930, 0.63], ['About Us', 11.006, 0.958]].forEach(([label, x, w]) => {
    text(s, label, x, 0.357, w, 0.286, { size: 11, color: C.white });
  });

  text(s, 'Thank You', 4.812, 2.816, 6.781, 1.582,
    { size: 88, bold: true, face: HEAD, color: C.white, align: 'right' });
  text(s, 'Ready to change the way people handle money? Let\u2019s talk.', 4.812, 4.314, 6.781, 0.37,
    { size: 16, face: HEAD, color: C.white, align: 'right' });

  vRule(s, 0.349, 6.312, 1.188, 3.788, C.white);
  socialTiles(s, 12.681, 5.922, true);
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19,
  slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pres = new PptxGenJS();
  PRES = pres;
  pres.defineLayout({ name: 'WIDE', width: SIZE.w, height: SIZE.h });
  pres.layout = 'WIDE';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Paywise';

  SLIDES.forEach(fn => {
    const s = pres.addSlide();
    s.background = { color: C.white };
    fn(s);
  });

  return pres.writeFile({ fileName: path.join(__dirname, '07247259-a1d4-473b-b293-96d15c21107a_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
