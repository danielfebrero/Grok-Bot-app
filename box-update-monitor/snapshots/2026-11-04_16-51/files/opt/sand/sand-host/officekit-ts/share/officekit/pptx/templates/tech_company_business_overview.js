/**
 * "Runnn" — simple business & finance presentation (40 slides, 26.66" x 15").
 *
 * Standalone pptxgenjs re-creation of the reference deck. Every slide is built
 * by its own `slideNN(s)` function from plain helper calls; all geometry is in
 * inches and matches the original shape positions.
 *
 * Raster photos in the original are re-created as flat grey placeholder blocks
 * (`photo`) and the icon artwork as small vector `icon()` drawings.
 *
 * Run: node 10bfac58-0ebb-4cfb-aa8d-713976759cd1_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 26.659722;
const SLIDE_H = 15;

/** Deck palette (hex, no leading #). */
const C = {
  DK: '696464',     // dark warm grey — panels, numbers, dark slides
  TAN: 'BDB196',    // accent tan
  GOLD: 'A28E6A',   // filled review stars
  GY: '7F7F7F',     // body copy grey / photo placeholder
  GY4: 'BFBFBF',
  GY6: '595959',
  HEAD: '3F3F3F',   // small caps headings
  INK: '262626',    // big titles
  W: 'FFFFFF',
  LT: 'F2F2F2',     // body copy on dark panels
  PALE: 'D8D8D8',
  RULE: '918485',   // short rule above titles
  PH: '7F7F7F',     // photo placeholder fill
  PH2: 'C3C3C3'     // lighter photo placeholder fill
};

/** Type stack of the template. */
const F = {
  title: 'Raleway ExtraBold',
  semi: 'Raleway SemiBold',
  black: 'Raleway Black',
  body: 'Merriweather',
  head: 'Poppins',
  headSemi: 'Poppins SemiBold'
};

// Text insets, in points, matching the source deck (0.1" sides / 0.05" top-bottom).
const INSET = [7.2, 7.2, 3.6, 3.6];
/** Soft drop shadow used by the infographic shapes on slides 32-38. */
const shadow = () => ({ type: 'outer', color: '000000', opacity: 0.2, blur: 40, offset: 18, angle: 45 });
const SHAPE_INSET = [14.4, 14.4, 7.2, 7.2];
const LEAD = 1.5;            // body-copy line spacing

const T1 = 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada';
const T2 = 'Maecenas porttitor congue massa. Fusce posuere, magna sed';
const T3 = 'Our Amazing Infographics';
const T4 = 'Id aliquet risus feugiat in ante. Eget magna';
const T5 = 'YOUR GREAT TITTLE PLACE HERE';
const T6 = 'Pellentesque habitant morbi tristique senectus';
const T7 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna pulvinar ultricies';
const T8 = 'Lorem ipsum dolor sit amet, consectetueres adipiscing elit congue massa. Fusce posuere, magna sed pulvinar amet';
const T9 = 'Non odio euismod lacinia at quis risus. Diam maecenas sed enim ut sem viverra aliquet';
const T10 = 'Dolor amet consectetur adipiscing. Id aliquet risus';
const T11 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit tellus. Maecenas porttitor congue massa. Fusce posuere, magna sed';
const T12 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.';
const T13 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et';
const T14 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna';
const T15 = 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus  magna amet';
const T16 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit lectus. Maecenas porttitor congue massa. Fusce posuere, magna sed';
const T17 = 'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.';
const T18 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit sed. Maecenas dolor porttitor congue massa. Fusce posuere';
const T19 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa posuere sed';
const T20 = 'YOUR GREAT SERVICES TITTLE';
const T21 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit sed. Maecenas porttitor congue massa. Fusce posuere, magna';
const T22 = 'YOUR GREAT SERVICES TITTLE HERE';
const T23 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa fusced sed';
const T24 = 'Maecenas porttitor congue massa sed. Fusce posuere, ae magna sed pulvinar ultricies, purusted';
const T25 = 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus labore magna';
const T26 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Elite maecenas porttitor congue massa. Fusce posuere, magna sed';
const T27 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing ae elit. Maecenas sed porttitor congue massa magna sed';
const T28 = 'GREAT TITTLE TO IMPRESS CLIENT';
const T29 = 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies';
const T30 = 'Our Amazing Single Tablet Mockup Here';
const T31 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit elite. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet';
const T32 = 'Maecenas porttitor congue massade. Fusce posuere, magna sed pulvinar';
const T33 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna.';
const T34 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elite. Maecenas porttitor congue massa fusced sed magna';
const T35 = 'Pellentesque habitant morbi ameted tristique senectus et netus et pedes malesuada fames ac turpis egestas. ';
const T36 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies';
const T37 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor amet';
const T38 = 'YOUR SERVICE TITTLE HERE';
const T39 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas lectus porttitor congue massa eros. Fusce posuere, magna pulvinar ultricies';
const T40 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elite, sed do eiusmod tempor sit elit incididunt ut labore et';
const T41 = 'PLACEHOLDER';


// ---------------------------------------------------------------- primitives

/** Text box. `o` may carry font/size/bold/color/lead/align/rotate overrides. */
function txt(s, text, x, y, w, h, o) {
  o = o || {};
  s.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: o.font || F.body,
    fontSize: o.size || 20,
    bold: !!o.bold,
    color: o.color || C.GY,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.lead,
    rotate: o.rotate,
    margin: INSET,
    wrap: true
  });
}

/** Big slide title (Raleway ExtraBold 62). */
function title(s, text, x, y, w, h, o) {
  o = o || {};
  txt(s, text, x, y, w, h, { font: F.title, size: 62, bold: true, color: o.color || C.INK, align: o.align, rotate: o.rotate });
}

/** Small caps section heading (Poppins bold 24). */
function head(s, text, x, y, w, h, o) {
  o = o || {};
  txt(s, text, x, y, w, h, { font: F.head, size: 24, bold: true, color: o.color || C.HEAD, align: o.align, rotate: o.rotate });
}

/** Body copy (Merriweather 20, 1.5 line spacing). */
function body(s, text, x, y, w, h, o) {
  o = o || {};
  txt(s, text, x, y, w, h, { font: F.body, size: 20, color: o.color || C.GY, align: o.align, lead: LEAD, rotate: o.rotate });
}

/** Short horizontal (or vertical, via o.h) accent rule; `o.dot` adds an end dot. */
function rule(s, x, y, w, o) {
  o = o || {};
  s.addShape('line', {
    x: x, y: y, w: w, h: o.h || 0,
    line: { color: o.color || C.RULE, width: o.width || 4.5,
            endArrowType: o.dot ? 'oval' : undefined }
  });
}

/** Any preset autoshape, optionally with centered label text. */
function shape(s, kind, x, y, w, h, o) {
  o = o || {};
  const opt = {
    x: x, y: y, w: w, h: h,
    fill: o.fill ? { color: o.fill, transparency: o.alpha } : undefined,
    line: o.line ? { color: o.line, width: o.lineW || 1 } : undefined,
    rotate: o.rotate, flipH: o.flipH, flipV: o.flipV, rectRadius: o.radius,
    shadow: o.shadow ? shadow() : undefined
  };
  if (o.text === undefined) { s.addShape(kind, opt); return; }
  opt.shape = kind;
  opt.align = 'center';
  opt.valign = 'middle';
  opt.margin = SHAPE_INSET;
  opt.fontFace = o.font || F.body;
  opt.fontSize = o.size || 20;
  opt.bold = !!o.bold;
  opt.color = o.color || C.W;
  s.addText(o.text, opt);
}

/** Placeholder standing in for a photograph in the original deck. */
function photo(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.shape || 'rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: o.color || C.PH },
    rectRadius: o.radius
  });
}

/**
 * Page number, drawn first on every slide so full-bleed artwork covers it
 * exactly as the slide master does in the reference deck.
 */
function pageNum(s, n) {
  txt(s, String(n), 19.909, 0.75, 5.998, 0.799,
      { font: F.body, size: 16, color: C.GY, align: 'right', valign: 'middle' });
}

/** Filled polygon; `pts` are [x, y] pairs in inches relative to (x, y). */
function poly(s, x, y, w, h, pts, o) {
  o = o || {};
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h,
    fill: o.fill ? { color: o.fill, transparency: o.alpha } : undefined,
    flipH: o.flipH, rotate: o.rotate,
    shadow: o.shadow ? shadow() : undefined,
    points: pts.map(p => ({ x: p[0], y: p[1] })).concat([{ close: true }])
  });
}

/** Notched block arrow used for the STEP ONE..SIX band on slide 32. */
function stepArrow(s, x, y, w, h, o) {
  const ss = Math.min(w, h);
  const half = h * 0.32164, head = ss * 0.46248, notch = half * head / (h / 2);
  const mid = h / 2, xh = w - head;
  poly(s, x, y, w, h, [[0, mid - half], [xh, mid - half], [xh, 0], [w, mid],
                       [xh, h], [xh, mid + half], [0, mid + half], [notch, mid]], o);
}

/** Chevron step banner used on slide 33. */
function chevronStep(s, x, y, w, h, o) {
  const dx = Math.min(w, h) * 0.33494;
  poly(s, x, y, w, h, [[0, 0], [w - dx, 0], [w, h / 2], [w - dx, h], [0, h], [dx, h / 2]], o);
}

/** Square callout with an arrow leaving its right edge (slide 34). */
function calloutArrow(s, x, y, w, h, o) {
  const ss = Math.min(w, h), stem = ss * 0.16716, head = ss * 0.30869, len = ss * 0.38651;
  const mid = h / 2, xb = w * 0.69913, xh = w - len;
  poly(s, x, y, w, h, [[xb, mid - stem], [xh, mid - stem], [xh, mid - head], [w, mid],
                       [xh, mid + head], [xh, mid + stem], [xb, mid + stem],
                       [xb, h], [0, h], [0, 0], [xb, 0]], o);
}

// ------------------------------------------------------------------ backdrop

/** Full-bleed background treatments taken from the deck's slide layouts. */
function backdrop(s, kind) {
  if (kind === 'dark') shape(s, 'rect', 0, 0, SLIDE_W, SLIDE_H, { fill: C.DK });
  if (kind === 'inset') {
    shape(s, 'rect', 0, 0, SLIDE_W, SLIDE_H, { fill: C.DK });
    shape(s, 'rect', 0.561, 0, 25.519, SLIDE_H, { fill: C.W });
  }
  if (kind === 'rightPanel') shape(s, 'rect', 12.58, 0, 13.5, SLIDE_H, { fill: C.HEAD });
  if (kind === 'leftPanel') shape(s, 'rect', 0.561, 0.75, 13.519, 14.248, { fill: C.DK });
}

// --------------------------------------------------------------------- icons

/**
 * Flat vector stand-ins for the deck's icon PNGs. Each is drawn inside a
 * `size` x `size` box anchored at (x, y) using simple pptxgenjs shapes.
 */
function icon(s, kind, x, y, size, color) {
  const u = size / 12;                                   // 12x12 design grid
  const put = (sh, a, b, c, d, extra) =>
    shape(s, sh, x + a * u, y + b * u, c * u, d * u, Object.assign({ fill: color }, extra || {}));
  const ring = (a, b, c, d, wt) =>
    shape(s, 'ellipse', x + a * u, y + b * u, c * u, d * u, { line: color, lineW: wt });

  if (kind === 'pie') {                                  // pie chart with one slice pulled out
    shape(s, 'pie', x + 0.2 * u, y + 0.8 * u, 10.4 * u, 10.4 * u, { fill: color, angleRange: [86, 356] });
    shape(s, 'pie', x + 1.4 * u, y + 0.2 * u, 10.4 * u, 10.4 * u, { fill: color, angleRange: [4, 84] });
  } else if (kind === 'bulb') {                          // light bulb with filament base
    ring(2.8, 0.5, 6.4, 6.4, 1.1 * u * 72);
    put('rect', 4.2, 5.4, 3.6, 1.1);
    put('rect', 4.4, 7.1, 3.2, 0.75);
    put('rect', 4.4, 8.2, 3.2, 0.75);
    put('rect', 4.4, 9.3, 3.2, 0.75);
    put('pie', 4.6, 9.9, 2.8, 2, { angleRange: [0, 180] });
  } else if (kind === 'network') {                       // hub-and-spoke node diagram
    [[2.6, 2.6, -45], [7.2, 2.6, 45], [2.6, 7.4, 45], [7.2, 7.4, -45], [5.5, 1.4, 0], [5.5, 8.1, 0]]
      .forEach(p => put('rect', p[0] + 0.9, p[1], 0.45, 2.6, { rotate: p[2] }));
    put('ellipse', 4.4, 4.4, 3.2, 3.2);
    put('ellipse', 0.3, 0.9, 2.6, 2.6);
    put('ellipse', 9.1, 0.9, 2.6, 2.6);
    put('ellipse', 0.3, 8.5, 2.6, 2.6);
    put('ellipse', 9.1, 8.5, 2.6, 2.6);
    put('ellipse', 4.9, 0, 2.2, 2.2);
    put('ellipse', 4.9, 9.8, 2.2, 2.2);
  } else if (kind === 'search') {                        // magnifier with heart-beat trace
    ring(0.5, 0.5, 8.6, 8.6, 1.1 * u * 72);
    put('rect', 7.9, 8.1, 3.6, 1.1, { rotate: 45 });
    put('rect', 3, 4.4, 0.5, 1.4);
    put('rect', 4.1, 3.2, 0.5, 3.6);
    put('rect', 5.2, 4.1, 0.5, 2);
    put('rect', 6.3, 4.4, 0.5, 1.4);
  } else if (kind === 'people') {                        // group of three figures
    put('ellipse', 0.2, 2.9, 2.9, 2.9);
    put('round1Rect', 0, 6, 3.6, 4, { radius: size / 12 });
    put('ellipse', 8.9, 2.9, 2.9, 2.9);
    put('round1Rect', 8.4, 6, 3.6, 4, { radius: size / 12 });
    put('ellipse', 3.9, 1.7, 4.2, 4.2);
    put('round1Rect', 2.7, 5.7, 6.6, 4.7, { radius: size / 9 });
  } else if (kind === 'bars' || kind === 'chart') {      // bar chart with rising arrow
    put('rect', 0.7, 1, 0.7, 10.3);
    put('rect', 0.7, 10.6, 10.6, 0.7);
    put('rect', 2.3, 6.9, 1.9, 3.3);
    put('rect', 4.9, 4.7, 1.9, 5.5);
    put('rect', 7.5, 2.6, 1.9, 7.6);
    put('rect', 2.6, 3.4, 5.4, 0.55, { rotate: -32 });
    put('triangle', 8.4, 0.9, 1.9, 1.9, { rotate: 60 });
  } else if (kind === 'puzzle') {                        // single jigsaw piece with knobs
    put('rect', 1.6, 1.6, 8.8, 8.8);
    put('ellipse', 4.1, 0, 3.8, 3.8);                    // top knob
    put('ellipse', 8.6, 4.1, 3.8, 3.8);                  // right knob
    shape(s, 'ellipse', x - 0.4 * u, y + 4.1 * u, 3.8 * u, 3.8 * u, { fill: C.W });
    shape(s, 'ellipse', x + 4.1 * u, y + 8.6 * u, 3.8 * u, 3.8 * u, { fill: C.W });
  } else if (kind === 'calendar') {                      // wall calendar with date grid
    put('rect', 0.6, 1.8, 10.8, 9.4);
    put('rect', 2.4, 0.5, 1, 2.4);
    put('rect', 8.6, 0.5, 1, 2.4);
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 4; c++) {
        shape(s, 'rect', x + (1.7 + c * 2.2) * u, y + (5.2 + r * 1.8) * u, 1.5 * u, 1 * u, { fill: C.DK });
      }
    }
  }
}

// -------------------------------------------------------- device mock-ups

/** Tablet body; the screen is a photo placeholder (slides 27-28). */
function tablet(s, x, y, w, h) {
  shape(s, 'roundRect', x, y, w, h, { fill: 'C7C1CB', radius: 0.42 });
  shape(s, 'roundRect', x + 0.046, y + 0.034, w - 0.09, h - 0.07, { fill: '565757', radius: 0.4 });
  shape(s, 'roundRect', x + 0.05, y + 0.05, w - 0.1, h - 0.1, { fill: '0E0E0E', radius: 0.4 });
  shape(s, 'roundRect', x + 0.36, y + 0.37, w - 0.73, h - 0.74, { fill: C.INK, radius: 0.2 });
  photo(s, x + 0.29, y + 0.323, w - 0.637, h - 0.653);
  shape(s, 'ellipse', x + 3.83, y + 0.14, 0.15, 0.15, { fill: '141414' });
  shape(s, 'rect', x + w, y + 1.2, 0.05, 0.32, { fill: 'C7C1CB' });
}

/** Phone body with pill notch; the screen is a photo placeholder (slides 29-30). */
function phone(s, x, y, w, h) {
  shape(s, 'roundRect', x + 0.04, y, w - 0.08, h, { fill: 'A6A299', radius: 0.7 });
  shape(s, 'roundRect', x + 0.06, y + 0.03, w - 0.12, h - 0.06, { fill: 'B1B2AD', radius: 0.68 });
  shape(s, 'roundRect', x + 0.086, y + 0.06, w - 0.17, h - 0.12, { fill: '010300', radius: 0.66 });
  shape(s, 'roundRect', x + 0.26, y + 0.215, w - 0.52, h - 0.43, { fill: '454442', radius: 0.55 });
  photo(s, x + 0.214, y + 0.195, w - 0.423, h - 0.391, { shape: 'roundRect', radius: 0.5 });
  shape(s, 'roundRect', x + 1.854, y + 0.376, 1.271, 0.39, { fill: '100E11', radius: 0.19 });
  shape(s, 'ellipse', x + 1.932, y + 0.456, 0.238, 0.238, { fill: '000100' });
  shape(s, 'ellipse', x + 2.773, y + 0.427, 0.295, 0.295, { fill: '060A0A' });
  shape(s, 'ellipse', x + 2.813, y + 0.467, 0.216, 0.216, { fill: '0A1519' });
  [[0, 2.306], [0, 8.97], [0, 3.325], [4.882, 0.917], [4.882, 3.495], [4.882, 8.97]]
    .forEach(p => shape(s, 'rect', x + p[0], y + p[1], 0.055, 0.09, { fill: 'A5A6A1' }));
}

/** Open laptop seen head-on; the screen is a photo placeholder (slide 31). */
function laptop(s, x, y, w, h) {
  const lidW = w * 0.841, lidX = x + w * 0.079;
  shape(s, 'roundRect', lidX, y, lidW, h * 0.959, { fill: '707A82', radius: 0.28 });
  shape(s, 'roundRect', lidX + 0.043, y + 0.029, lidW - 0.087, h * 0.952 - 0.06, { fill: '262A2D', radius: 0.26 });
  shape(s, 'roundRect', lidX + 0.075, y + 0.06, lidW - 0.15, h * 0.945 - 0.12, { fill: '000201', radius: 0.24 });
  photo(s, x + 1.42, y + 0.237, 12.333, 8.03);
  shape(s, 'roundRect', x, y + h * 0.951, w, h * 0.041, { fill: '788391', radius: 0.05 });
  shape(s, 'roundRect', x + w * 0.409, y + h * 0.956, w * 0.181, 0.139, { fill: '9EA4B4', radius: 0.06 });
}

function slide01(s) {
  photo(s, 0, 0, 26.66, 15, { color: C.PH2 });
  shape(s, 'ellipse', 12.84, 14.01, 0.19, 0.19, { fill: C.GY, rotate: -90 });
  shape(s, 'ellipse', 13.24, 14.01, 0.19, 0.19, { fill: C.GY4, rotate: -90 });
  shape(s, 'ellipse', 13.63, 14.01, 0.19, 0.19, { fill: C.GY4, rotate: -90 });
  txt(s, [{ text: 'R' }, { text: 'U', options: { italic: true } }, { text: 'NNN' }, { text: '.', options: { color: C.TAN } }], 6.16, 5.75, 14.35, 2.52, { font: F.black, size: 144, bold: true, color: C.INK, align: 'center' });
  txt(s, 'simple business & finance presentation', 6.16, 8.25, 14.35, 0.81, { font: F.semi, size: 42, bold: true, color: C.GY6, align: 'center' });
  txt(s, 'Runnn.', 24.56, 13.81, 1.12, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY, align: 'right' });
  txt(s, 'Template', 0.98, 13.81, 1.49, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY });
  txt(s, 'Finance', 24.43, 0.82, 1.25, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY, align: 'right' });
  txt(s, 'Business', 1.39, 0.82, 1.38, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY });
}

function slide02(s) {
  title(s, 'Hello! Welcome to Runnn Company', 1.33, 1.82, 10.49, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas amet porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. ', 1.33, 5.17, 10.49, 1.56);
  head(s, 'INTRODUCTION', 1.33, 4.6, 10.49, 0.51);
  rule(s, 1.41, 1.33, 1.67);
  body(s, T33, 14.84, 9.48, 4.5, 3.58);
  head(s, 'YOUR TITTLE', 14.84, 8.88, 4.5, 0.51);
  rule(s, 14.91, 8.27, 1.18);
  body(s, T33, 20.09, 9.48, 4.5, 3.58);
  head(s, 'YOUR TITTLE', 20.09, 8.88, 4.5, 0.51);
  rule(s, 20.17, 8.27, 1.18);
  photo(s, 14.83, 0.75, 5.24, 6.75);
  photo(s, 20.09, 0, 5.24, 7.5);
  photo(s, 1.35, 7.5, 5.24, 7.5);
  photo(s, 6.61, 7.5, 5.22, 6.75);
}

function slide03(s) {
  photo(s, 1.33, 3.76, 9.75, 11.24, { color: C.PH2 });
  title(s, 'Runnn Best Tech Company History Here', 13.33, 2.08, 11.25, 2.19);
  body(s, 'Non odio euismod lacinia at quis risus. Diam maecenas sed enim ut sem viverra aliquet eget sit. Integer eget aliquet nibh praesent tristique magna sit amet. Ac mauris pharetra et ultrices neque ornare. Dui id ornare arcu odio. Ullamcorper malesuada proin libero nunc consequat. Integer malesuada nunc vel risusae', 13.33, 5.45, 11.25, 2.07);
  head(s, T5, 13.33, 4.84, 11.25, 0.51);
  rule(s, 13.41, 1.58, 1.67);
  body(s, 'Non odio euismod lacinia at quis risus. Diam dolore maecenas sed enim ut sem viverra aliquet eget sit ae. Integer eget aliquet nibh praesent tristique magna', 3.58, 1.44, 7.5, 1.56);
  head(s, T5, 3.58, 0.83, 7.5, 0.51);
  shape(s, 'rect', 1.34, 2.67, 1.5, 1.83, { fill: C.DK });
  shape(s, 'rect', 15.58, 9, 4.5, 6, { fill: C.GY });
  shape(s, 'rect', 20.08, 9, 4.5, 6, { fill: C.TAN });
  shape(s, 'rect', 11.08, 9, 4.5, 6, { fill: C.DK });
  body(s, T15, 11.6, 11.9, 3.46, 2.07, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 11.6, 11.3, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'pie', 12.81, 10.03, 1.03, C.W);
  body(s, T15, 16.1, 11.95, 3.46, 2.07, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 16.1, 11.34, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'bulb', 17.32, 9.98, 1.02, C.W);
  body(s, T15, 20.6, 11.94, 3.46, 2.07, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 20.6, 11.34, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'people', 21.81, 9.99, 1.03, C.W);
  icon(s, 'puzzle', 1.36, 0.79, 1.46, C.DK);
}

function slide04(s) {
  backdrop(s, 'dark');
  title(s, 'Our Company Focus On Technology', 2.08, 2.23, 9.01, 2.19, { color: C.W });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit lectus. Maecenas porttitor congue massa. Fusce posuere, magna sed  ae pulvinar ultricies, purus lectus malesuada libero, sit amet', 2.08, 5, 9, 1.56, { color: C.PALE });
  head(s, 'FOCUS ON THE NETWORKING', 2.08, 7.14, 9.05, 0.51, { color: C.W });
  body(s, T16, 2.08, 7.74, 9, 1.06, { color: C.PALE });
  rule(s, 2.16, 1.75, 1.67);
  head(s, 'FOCUS ON THE QUALITY', 2.08, 9.36, 9.05, 0.51, { color: C.W });
  body(s, T16, 2.08, 9.96, 9, 1.06, { color: C.PALE });
  head(s, 'FOCUS ON THE QUANTITY', 2.08, 11.57, 9.05, 0.51, { color: C.W });
  body(s, T16, 2.08, 12.18, 9, 1.06, { color: C.PALE });
  body(s, 'Non odio euismod lacinia at quis risus. Diam maecenas sed enim ut sem viverra  proin aliquet eget sit. Integer eget aliquet nibh praesent tristique magna sit amet. Ac mauris pharetra et ultrices neque ornare. Dui id ornare arcu odio. Ullamcorper malesuada.', 13.33, 11.19, 12, 1.56, { color: C.LT });
  head(s, T5, 13.33, 10.58, 12, 0.51, { color: C.W });
  rule(s, 13.41, 13.33, 1.18);
  photo(s, 13.33, 1.5, 12, 8.25);
}

function slide05(s) {
  photo(s, 0, 0.75, 26.66, 13.5, { color: C.PH2 });
  shape(s, 'rect', 6.55, 4.51, 13.56, 5.98, { fill: C.DK, alpha: 5.1 });
  title(s, 'Runnn Amazing Tittle Place Here Please!', 7.71, 5.83, 11.25, 2.19, { color: C.W, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor est congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus dolore', 7.71, 8.59, 11.25, 1.06, { color: C.PALE, align: 'center' });
  rule(s, 12.5, 5.35, 1.67);
}

function slide06(s) {
  shape(s, 'rect', 11.08, 10.5, 9, 3.75, { fill: C.DK, alpha: 5.1 });
  body(s, 'Non odio euismod lacinia at quis risus. Diam ae dolore maecenas sed enim ut sem viverra amet aliquet eget sit ae. Integer eget aliquet nibh', 12.59, 11.86, 6.74, 1.56, { color: C.LT });
  head(s, T5, 12.59, 11.25, 6.74, 0.51, { color: C.W });
  title(s, 'Runnn Company Best Opening Tittle Here', 1.33, 2.15, 10.5, 2.19);
  body(s, 'Non odio euismod lacinia at quis risus. Diam maecenas sed enim ut sem viverra aliquet eget sit. Integer eget aliquet nibh praesent tristique magna sit amet. Ac mauris pharetra et ultrices neque ornare. Dui id ornare arcu', 1.33, 5.52, 10.5, 1.56);
  head(s, T5, 1.33, 4.92, 10.5, 0.51);
  rule(s, 1.41, 1.65, 1.67);
  body(s, T17, 21.6, 8.87, 3.73, 3.08);
  photo(s, 0.56, 8.25, 11.27, 6.75);
  photo(s, 12.58, 0.75, 4.5, 4.49);
  photo(s, 12.58, 6.01, 7.5, 4.49);
  photo(s, 21.58, 0.76, 4.5, 6.75);
  head(s, 'YOUR TITTLE HERE', 21.6, 8.27, 3.73, 0.51);
  body(s, 'Maecenas porttitor congue massa elite. Fusce posuere, sed magna pulvinar', 17.86, 3.18, 2.96, 2.07);
  head(s, 'YOUR TITTLE', 17.86, 2.58, 2.96, 0.51);
}

function slide07(s) {
  title(s, 'Runnns Company Best Visions & Missions', 2.08, 2.4, 10.5, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas viverra porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc ae viverra imperdiet enim. Fusce est. Vivamus a tellus.', 2.08, 5.77, 10.5, 2.07);
  shape(s, 'diamond', 2.13, 9.3, 1.32, 1.32, { fill: C.DK, text: '1', font: F.body, size: 24, color: C.W });
  body(s, T18, 3.83, 9.38, 8.75, 1.06);
  shape(s, 'diamond', 2.13, 10.69, 1.32, 1.32, { fill: C.TAN, text: '2', font: F.body, size: 24, color: C.W });
  body(s, T18, 3.83, 10.77, 8.75, 1.06);
  head(s, 'OUR COMPANY VISION', 2.13, 5.17, 10.42, 0.51);
  head(s, 'OUR COMPANY MISSION', 2.13, 8.54, 10.42, 0.51);
  rule(s, 2.21, 1.92, 1.67);
  shape(s, 'diamond', 2.13, 12.05, 1.32, 1.32, { fill: C.DK, text: '3', font: F.body, size: 24, color: C.W });
  body(s, T18, 3.83, 12.13, 8.75, 1.06);
  body(s, T9, 20.08, 1.86, 4.5, 1.56);
  head(s, 'YOUR TITTLE HERE', 20.08, 1.25, 4.5, 0.51);
  rule(s, 20.16, 4, 1.18);
  body(s, T9, 20.08, 6.77, 4.5, 1.56);
  head(s, 'YOUR TITTLE HERE', 20.08, 6.17, 4.5, 0.51);
  rule(s, 20.16, 8.92, 1.18);
  body(s, T9, 20.08, 12.02, 4.5, 1.56);
  head(s, 'YOUR TITTLE HERE', 20.08, 11.41, 4.5, 0.51);
  rule(s, 20.16, 14.17, 1.18);
  photo(s, 14.1, 0.75, 5.23, 3.75);
  photo(s, 14.1, 5.25, 5.23, 4.5);
  photo(s, 14.1, 10.5, 5.23, 4.5);
}

function slide08(s) {
  backdrop(s, 'inset');
  title(s, 'Runnn Is a Big Tech Company in 2023', 2.12, 1.68, 11.2, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus amet malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra  sed imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi.', 2.11, 5.18, 11.21, 2.07);
  head(s, 'COLLABORATION WITH EXPERT TEAM', 2.11, 4.58, 11.21, 0.51);
  rule(s, 2.21, 1.19, 1.67);
  photo(s, 2.08, 8.25, 11.24, 6);
  photo(s, 14.83, 0.75, 11.24, 14.25);
}

function slide09(s) {
  title(s, 'Runnn Amazing Top Teams', 2.09, 1.59, 22.47, 1.14, { align: 'center' });
  txt(s, 'Sales Manager', 2.02, 11.31, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  photo(s, 2.02, 4.49, 4.72, 5.47, { shape: 'roundRect', radius: 0.79 });
  photo(s, 7.99, 4.49, 4.72, 5.47, { shape: 'roundRect', radius: 0.79 });
  photo(s, 13.95, 4.49, 4.72, 5.47, { shape: 'roundRect', radius: 0.79 });
  photo(s, 19.92, 4.49, 4.72, 5.47, { shape: 'roundRect', radius: 0.79 });
  head(s, 'SARAH ETEON', 2.02, 10.7, 4.72, 0.51, { align: 'center' });
  txt(s, 'Web Designer', 7.99, 11.3, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'BILL SIMOTI', 7.99, 10.7, 4.72, 0.51, { align: 'center' });
  txt(s, 'Copywriting', 19.95, 11.31, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'GARRY CHOCO', 19.92, 10.7, 4.72, 0.51, { align: 'center' });
  body(s, T10, 2.02, 12.27, 4.72, 1.06, { align: 'center' });
  body(s, T10, 7.99, 12.27, 4.72, 1.06, { align: 'center' });
  body(s, T10, 19.92, 12.27, 4.72, 1.06, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
  rule(s, 9.76, 13.83, 1.18);
  rule(s, 3.79, 13.83, 1.18);
  rule(s, 21.69, 13.83, 1.18);
  shape(s, 'triangle', 2.52, 9.48, 0.76, 0.65, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 1.67, 8.58, 1.49, 1.29, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 8.62, 9.48, 0.76, 0.65, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 7.77, 8.58, 1.49, 1.29, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 13.81, 8.58, 1.49, 1.29, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 14.66, 9.48, 0.76, 0.65, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 20.69, 9.48, 0.76, 0.65, { line: C.GY4, rotate: -29.36 });
  shape(s, 'triangle', 19.84, 8.58, 1.49, 1.29, { line: C.GY4, rotate: -29.36 });
  txt(s, 'Marketing Manager', 13.95, 11.3, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'EMIR SHAM', 13.95, 10.7, 4.72, 0.51, { align: 'center' });
  body(s, T10, 13.95, 12.27, 4.72, 1.06, { align: 'center' });
  rule(s, 15.72, 13.83, 1.18);
}

function slide10(s) {
  title(s, 'Runnn Amazing Top Teams', 0.99, 1.59, 24.68, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
  txt(s, 'Sales Manager', 2.02, 12.61, 4.63, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'SARAH ETEON', 2.02, 12, 4.72, 0.51, { align: 'center' });
  txt(s, 'Web Designer', 7.99, 12.61, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'BILL SIMOTI', 7.99, 12, 4.72, 0.51, { align: 'center' });
  txt(s, 'Copywriting', 19.89, 12.61, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'GARRY CHOCO', 19.86, 12, 4.72, 0.51, { align: 'center' });
  rule(s, 9.76, 13.8, 1.18);
  rule(s, 3.79, 13.8, 1.18);
  rule(s, 21.63, 13.8, 1.18);
  txt(s, 'Marketing Manager', 13.95, 12.61, 4.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'EMIR SHAM', 13.95, 12, 4.72, 0.51, { align: 'center' });
  rule(s, 15.72, 13.8, 1.18);
  photo(s, 19.86, 4.49, 4.72, 6.76);
  photo(s, 13.95, 4.49, 4.72, 6.76);
  photo(s, 7.99, 4.49, 4.72, 6.76);
  photo(s, 2.08, 4.49, 4.72, 6.76);
}

function slide11(s) {
  photo(s, 0.56, 0.01, 25.52, 6.74, { color: C.PH2 });
  photo(s, 1.33, 5.25, 6, 7.52);
  title(s, 'Runnn Amazing Team Portfolios Place Here', 8.09, 8.16, 11.24, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus amet malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra  sed', 8.08, 11.66, 11.25, 1.56);
  head(s, 'AMAZING TEAM AMAZING WORK', 8.08, 11.06, 11.25, 0.51);
  rule(s, 8.17, 7.67, 1.67);
  shape(s, 'rect', 2.1, 12, 5.27, 1.5, { fill: C.DK, alpha: 5.1 });
  txt(s, 'Sales Manager', 2.1, 12.75, 5.23, 0.55, { font: F.body, size: 20, bold: true, color: C.W, lead: 1.5, align: 'center' });
  head(s, 'SARAH ETEON', 2.1, 12.23, 5.23, 0.51, { color: C.W, align: 'center' });
  shape(s, 'rect', 20.83, 6.01, 4.5, 4.5, { fill: C.GY });
  shape(s, 'rect', 20.83, 10.5, 4.5, 4.5, { fill: C.TAN });
  shape(s, 'rect', 20.83, 1.52, 5.25, 4.5, { fill: C.DK });
  body(s, T4, 21.35, 4.18, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 21.35, 3.57, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'pie', 22.56, 2.3, 1.03, C.W);
  body(s, T4, 21.35, 8.71, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 21.35, 8.11, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'bulb', 22.57, 6.75, 1.02, C.W);
  body(s, T4, 21.35, 13.2, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 21.35, 12.59, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'people', 22.56, 11.25, 1.03, C.W);
}

function slide12(s) {
  title(s, 'Runnn Great Team', 0.94, 1.54, 24.77, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
  txt(s, 'Sales Manager', 2.51, 11.15, 3.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'SARAH ETEON', 2.51, 10.54, 3.72, 0.51, { align: 'center' });
  txt(s, 'Web Designer', 7.02, 11.14, 3.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'BILL SIMOTI', 7.02, 10.54, 3.72, 0.51, { align: 'center' });
  txt(s, 'Copywriting', 15.98, 11.15, 3.74, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'GARRY CHOCO', 15.98, 10.54, 3.72, 0.51, { align: 'center' });
  body(s, T6, 2.51, 12.11, 3.72, 0.94, { align: 'center' });
  body(s, T6, 7.02, 12.11, 3.72, 0.94, { align: 'center' });
  body(s, T6, 15.98, 12.11, 3.72, 0.94, { align: 'center' });
  rule(s, 8.29, 13.67, 1.18);
  rule(s, 3.78, 13.67, 1.18);
  rule(s, 17.25, 13.67, 1.18);
  txt(s, 'Marketing Manager', 11.5, 11.14, 3.72, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'EMIR SHAM', 11.5, 10.54, 3.72, 0.51, { align: 'center' });
  body(s, T6, 11.5, 12.11, 3.72, 0.94, { align: 'center' });
  rule(s, 12.77, 13.67, 1.18);
  txt(s, 'Costumer Relationship', 20.47, 11.15, 3.74, 0.55, { font: F.body, size: 20, bold: true, color: C.GY, lead: 1.5, align: 'center' });
  head(s, 'MARCUSS', 20.47, 10.54, 3.72, 0.51, { align: 'center' });
  body(s, T6, 20.47, 12.11, 3.72, 0.94, { align: 'center' });
  rule(s, 21.74, 13.67, 1.18);
  photo(s, 2.51, 4.49, 3.72, 5.26);
  photo(s, 7.02, 4.49, 3.72, 5.26);
  photo(s, 11.5, 4.52, 3.72, 5.26);
  photo(s, 15.98, 4.52, 3.72, 5.26);
  photo(s, 20.47, 4.52, 3.72, 5.26);
}

function slide13(s) {
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est.', 11.08, 5.19, 11.99, 1.56);
  head(s, 'SERVICE TITTLE HERE', 11.08, 4.58, 11.99, 0.51);
  title(s, 'Our Runnns Excellent Services Place Here', 11.08, 1.81, 12, 2.19);
  photo(s, 1.33, 0.75, 9, 5.99);
  photo(s, 11.08, 7.5, 12, 7.5);
  rule(s, 11.15, 1.32, 1.67);
  body(s, T19, 2.83, 8.11, 7.5, 1.06);
  head(s, T20, 2.83, 7.5, 7.5, 0.51);
  icon(s, 'pie', 1.17, 7.66, 1.34, C.TAN);
  icon(s, 'bulb', 1.16, 9.84, 1.33, C.DK);
  icon(s, 'people', 1.17, 12.08, 1.34, C.TAN);
  body(s, T19, 2.83, 10.27, 7.5, 1.06);
  head(s, T20, 2.83, 9.67, 7.5, 0.51);
  body(s, T19, 2.83, 12.52, 7.5, 1.06);
  head(s, T20, 2.83, 11.92, 7.5, 0.51);
  shape(s, 'rect', 21.58, 11.16, 3.75, 3.09, { fill: C.DK });
}

function slide14(s) {
  photo(s, 1.33, 0.77, 10.52, 10.48, { shape: 'round1Rect', radius: 2.23 });
  title(s, 'Our Excellent Services', 13.36, 1.74, 11.25, 1.14);
  body(s, [{ text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor sed congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus eros ae malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra enim. ' }, { text: '', options: { breakLine: true } }, { text: 'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.', options: { breakLine: true } }], 13.36, 3.23, 11.25, 3.08);
  rule(s, 13.44, 1.25, 1.67);
  shape(s, 'roundRect', 0, 9.81, 11.85, 4.43, { fill: C.DK, radius: 0 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit tellus. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet', 2.08, 11.1, 9, 1.56, { color: C.LT });
  head(s, 'OUR LATEST CUSTOMER TESTIMONIALS', 2.08, 10.49, 9, 0.51, { color: C.W });
  shape(s, 'star5', 2.13, 13.21, 0.56, 0.57, { fill: C.GOLD });
  shape(s, 'star5', 2.83, 13.21, 0.56, 0.57, { fill: C.GOLD });
  shape(s, 'star5', 3.53, 13.21, 0.56, 0.57, { fill: C.GOLD });
  shape(s, 'star5', 4.24, 13.21, 0.56, 0.57, { fill: C.GOLD });
  shape(s, 'star5', 4.94, 13.21, 0.56, 0.57, { fill: 'E8E6E6' });
  txt(s, 'Martinn Estern – 5min ago', 6.55, 13.16, 4.53, 0.55, { font: F.body, size: 20, bold: true, color: C.W, lead: 1.5 });
  shape(s, 'diamond', 13.16, 6.97, 1.32, 1.32, { fill: C.TAN, text: '1', font: F.body, size: 24, color: C.W });
  body(s, T7, 14.86, 7.05, 9.72, 1.06);
  shape(s, 'diamond', 13.16, 8.37, 1.32, 1.32, { fill: C.DK, text: '2', font: F.body, size: 24, color: C.W });
  body(s, T7, 14.86, 8.45, 9.72, 1.06);
  shape(s, 'diamond', 13.16, 9.73, 1.32, 1.32, { fill: C.TAN, text: '3', font: F.body, size: 24, color: C.W });
  body(s, T7, 14.86, 9.81, 9.72, 1.06);
  shape(s, 'diamond', 13.16, 11.12, 1.32, 1.32, { fill: C.DK, text: '4', font: F.body, size: 24, color: C.W });
  body(s, T7, 14.86, 11.2, 9.72, 1.06);
  shape(s, 'diamond', 13.16, 12.54, 1.32, 1.32, { fill: C.TAN, text: '5', font: F.body, size: 24, color: C.W });
  body(s, T7, 14.86, 12.62, 9.72, 1.06);
}

function slide15(s) {
  backdrop(s, 'dark');
  photo(s, 14.85, 8.25, 11.23, 6.75);
  photo(s, 1.33, 0, 4.5, 6.03);
  photo(s, 6.58, 1.53, 6.75, 4.5);
  title(s, 'Runnn Company Best Service Place Here', 1.33, 7.36, 12, 2.19, { color: C.W });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est pulvina. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.', 1.33, 10.85, 12, 2.57, { color: C.LT });
  head(s, 'MAKE YOUR BUSINESS GROW WITH US', 1.33, 10.25, 12, 0.51, { color: C.W });
  rule(s, 1.41, 6.86, 1.67);
  body(s, T21, 16.66, 2.11, 8.67, 1.06, { color: C.LT });
  head(s, T22, 16.66, 1.6, 8.67, 0.51, { color: C.W });
  icon(s, 'pie', 14.67, 1.68, 1.34, C.TAN);
  icon(s, 'bulb', 14.66, 3.86, 1.33, C.W);
  icon(s, 'people', 14.67, 6.1, 1.34, C.TAN);
  body(s, T21, 16.66, 4.28, 8.67, 1.06, { color: C.LT });
  head(s, T22, 16.66, 3.77, 8.67, 0.51, { color: C.W });
  body(s, T21, 16.66, 6.53, 8.67, 1.06, { color: C.LT });
  head(s, T22, 16.66, 6.02, 8.67, 0.51, { color: C.W });
  shape(s, 'rect', 20.85, 11, 5.81, 3.25, { fill: C.TAN });
  body(s, 'Nunc viverra imperdiet enim. Fusce est. Vivamus', 21.62, 12.44, 3.71, 1.06, { color: C.LT });
  head(s, 'SERVICE TITTLE', 21.62, 11.84, 3.71, 0.51, { color: C.W });
}

function slide16(s) {
  title(s, 'Runnn Company Great Service Here', 2.08, 2.65, 9, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit tellus. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet amet commodo magna eros quis urna. Nunc viverra imperdiet', 2.08, 5.55, 9, 2.07);
  shape(s, 'ellipse', 2, 8.37, 1.38, 1.38, { fill: C.TAN });
  shape(s, 'ellipse', 2, 10.03, 1.38, 1.38, { fill: C.DK });
  body(s, T23, 3.67, 8.53, 7.4, 1.06);
  body(s, T23, 3.67, 10.19, 7.4, 1.06);
  icon(s, 'network', 2.29, 8.66, 0.79, C.W);
  icon(s, 'search', 2.29, 10.33, 0.79, C.W);
  rule(s, 2.16, 2.17, 1.67);
  body(s, T24, 12.58, 9.61, 2.99, 2.57);
  head(s, 'TITTLE HERE', 12.58, 9, 2.99, 0.51);
  rule(s, 12.72, 12.67, 1.18);
  body(s, T24, 17.08, 9.61, 2.99, 2.57);
  head(s, 'TITTLE HERE', 17.08, 9, 2.99, 0.51);
  rule(s, 17.22, 12.67, 1.18);
  body(s, T24, 21.58, 9.61, 2.99, 2.57);
  head(s, 'TITTLE HERE', 21.58, 9, 2.99, 0.51);
  rule(s, 21.72, 12.67, 1.18);
  shape(s, 'ellipse', 2.03, 11.71, 1.38, 1.38, { fill: C.TAN });
  body(s, T23, 3.7, 11.87, 7.4, 1.06);
  icon(s, 'bars', 2.33, 12, 0.79, C.W);
  photo(s, 12.59, 0.75, 4.49, 7.5);
  photo(s, 21.57, 0.75, 4.5, 7.5);
  photo(s, 17.08, 0.75, 4.49, 7.5, { color: C.PH2 });
}

function slide17(s) {
  backdrop(s, 'dark');
  shape(s, 'rect', 2.08, 0.75, 4.5, 6.75, { fill: C.GY });
  shape(s, 'rect', 6.91, 0.75, 4.17, 5.78, { fill: C.TAN });
  shape(s, 'rect', 2.08, 7.88, 4.5, 6.37, { fill: C.TAN });
  shape(s, 'rect', 6.91, 6.79, 4.17, 5.96, { fill: C.GY });
  body(s, 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus labore magna\ntempor non elited\npurus magna sed', 2.6, 3.45, 3.46, 3.08, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 2.6, 2.84, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'pie', 3.81, 1.57, 1.03, C.W);
  body(s, 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus labore magna\ntempor non elited', 7.27, 3.54, 3.46, 2.57, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 7.27, 2.93, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'bulb', 8.48, 1.57, 1.02, C.W);
  body(s, 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus magna amet \nTempor non elited', 7.27, 9.29, 3.46, 2.57, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 7.27, 8.68, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'search', 8.48, 7.33, 1.03, C.W);
  body(s, 'Id aliquet risus feugiat in ante. Eget magna fermentum amet nisl purus labore magna\ntempor non elited\npurus magna sed', 2.6, 10.45, 3.46, 3.08, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 2.6, 9.85, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'people', 3.81, 8.5, 1.03, C.W);
  title(s, [{ text: 'Amazing' }, { text: ' ', options: { italic: true } }, { text: 'Service In Our Company Runnn' }], 13.37, 2.07, 11.21, 2.19, { color: C.W });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor sed congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus liberos malesuada, sit amet commodo magna eros quis urna. Nunc viverra imperdiet', 13.33, 4.96, 11.25, 1.56, { color: C.LT });
  rule(s, 13.5, 1.58, 1.67);
  shape(s, 'ellipse', 13.25, 10.46, 1.38, 1.38, { fill: C.TAN });
  shape(s, 'ellipse', 13.25, 7.12, 1.38, 1.38, { fill: C.TAN });
  shape(s, 'ellipse', 13.25, 8.79, 1.38, 1.38, { fill: C.GY });
  body(s, T11, 14.92, 7.28, 9.54, 1.06, { color: C.LT });
  body(s, T11, 14.92, 8.95, 9.54, 1.06, { color: C.LT });
  icon(s, 'network', 13.54, 7.42, 0.79, C.W);
  icon(s, 'search', 13.54, 9.08, 0.79, C.W);
  body(s, T11, 14.92, 10.62, 9.54, 1.06, { color: C.LT });
  icon(s, 'pie', 13.54, 10.75, 0.79, C.W);
  shape(s, 'ellipse', 13.25, 12.12, 1.38, 1.38, { fill: C.GY });
  body(s, T11, 14.92, 12.28, 9.54, 1.06, { color: C.LT });
  icon(s, 'bars', 13.54, 12.42, 0.79, C.W);
}

function slide18(s) {
  title(s, 'Our Amazing Service In The Last Year Here', 14.08, 2.08, 10.5, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas amet porttitor sed congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus telluse malesuada libero, sit amet commodo magna eros', 14.08, 5.01, 10.48, 1.56);
  body(s, T26, 15.67, 7.92, 8.91, 1.06);
  body(s, T26, 15.67, 10.19, 8.91, 1.06);
  head(s, 'SERVICE TITTLE', 15.67, 7.32, 8.91, 0.51);
  head(s, 'SERVICE TITTLE', 15.67, 9.58, 8.91, 0.51);
  icon(s, 'network', 14, 7.32, 1.29, C.TAN);
  icon(s, 'search', 14, 9.58, 1.29, C.DK);
  rule(s, 14.16, 1.58, 1.67);
  body(s, T26, 15.67, 12.46, 8.91, 1.06);
  head(s, 'SERVICE TITTLE', 15.67, 11.85, 8.91, 0.51);
  icon(s, 'pie', 14, 11.85, 1.29, C.TAN);
  photo(s, 0.58, 7.52, 6, 5.98, { color: C.PH2 });
  photo(s, 6.58, 0.78, 6, 7.46, { color: C.PH2 });
  photo(s, 6.58, 8.25, 6, 6.75);
  photo(s, 0.58, 0, 6, 7.52);
}

function slide19(s) {
  backdrop(s, 'rightPanel');
  body(s, T12, 13.33, 10.36, 6, 2.57, { color: C.PALE });
  head(s, 'YOUR TITTLE PLACE HERE', 13.33, 9.75, 6, 0.51, { color: C.W });
  title(s, [{ text: 'Our Latest Amazing' }, { text: '  ', options: { italic: true } }, { text: 'Service Company Here' }], 1.33, 2.04, 9.75, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas dolore porttitor congue massa. Fusce posuere, magna sed pulvinar ae ultricies, purus lectus malesuada libero, sit amet commodo', 1.35, 4.94, 9.73, 1.56);
  rule(s, 1.44, 1.55, 1.67);
  body(s, 'Lorem ipsum dolor sit amet, lectus consectetuer adipiscing elit. Urna ae maecenas porttitor congue massa. ', 20.08, 10.36, 5.23, 1.56, { color: C.PALE });
  head(s, 'YOUR TITTLE HERE', 20.08, 9.75, 5.23, 0.51, { color: C.W });
  rule(s, 13.49, 13.42, 1.18);
  rule(s, 20.2, 12.47, 1.18);
  shape(s, 'ellipse', 1.25, 7.12, 1.38, 1.38, { fill: C.TAN });
  shape(s, 'ellipse', 1.25, 8.78, 1.38, 1.38, { fill: C.DK });
  body(s, T34, 2.92, 7.28, 8.16, 1.06);
  body(s, T34, 2.92, 8.94, 8.16, 1.06);
  icon(s, 'network', 1.54, 7.41, 0.79, C.W);
  icon(s, 'search', 1.54, 9.08, 0.79, C.W);
  photo(s, 1.33, 11.25, 3.75, 3.75);
  photo(s, 5.08, 11.25, 6, 3.75, { color: C.PH2 });
  photo(s, 13.33, 0.77, 12, 8.23);
}

function slide20(s) {
  photo(s, 0, 0, 26.66, 15, { color: C.PH2 });
  txt(s, [{ text: 'BRE' }, { text: 'A', options: { italic: true } }, { text: 'K SLIDE' }], 6.49, 6.24, 13.69, 2.52, { font: F.black, size: 144, bold: true, color: C.INK, align: 'center' });
  shape(s, 'ellipse', 12.84, 13.99, 0.19, 0.19, { fill: C.GY, rotate: -90 });
  shape(s, 'ellipse', 13.24, 13.99, 0.19, 0.19, { fill: C.INK, rotate: -90 });
  shape(s, 'ellipse', 13.63, 13.99, 0.19, 0.19, { fill: C.GY, rotate: -90 });
  txt(s, 'Finance', 24.43, 0.82, 1.25, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY, align: 'right' });
  txt(s, 'Business', 1.39, 0.82, 1.38, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY });
}

function slide21(s) {
  title(s, 'Our Amazing Portfolios In The Month Place Here', 1.25, 9.59, 11.96, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim lectus elit.', 1.2, 12.48, 12, 1.56);
  body(s, T27, 16.62, 1.94, 7.96, 1.06);
  head(s, T28, 16.65, 1.34, 7.94, 0.51);
  rule(s, 1.33, 9.09, 1.67);
  body(s, T27, 16.62, 4.19, 7.96, 1.06);
  head(s, T28, 16.65, 3.59, 7.94, 0.51);
  body(s, T27, 16.62, 6.44, 7.96, 1.06);
  head(s, T28, 16.65, 5.84, 7.94, 0.51);
  shape(s, 'diamond', 14.75, 1.39, 1.56, 1.56, { fill: C.DK, text: '1', font: F.body, size: 24, color: C.W });
  shape(s, 'diamond', 14.75, 3.64, 1.56, 1.56, { fill: C.TAN, text: '2', font: F.body, size: 24, color: C.W });
  shape(s, 'diamond', 14.75, 5.89, 1.56, 1.56, { fill: C.DK, text: '3', font: F.body, size: 24, color: C.W });
  body(s, T9, 7.33, 1.41, 6.75, 1.06);
  head(s, 'YOUR TITTLE HERE', 7.33, 0.8, 4.5, 0.51);
  rule(s, 7.41, 3, 1.18);
  photo(s, 1.33, 0.75, 5.25, 7.5);
  photo(s, 7.33, 3.75, 6.75, 4.5);
  photo(s, 14.83, 8.25, 11.25, 6.75);
}

function slide22(s) {
  title(s, 'See Our Amazing Portfolios Here', 10.09, 2.23, 9, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore ae magna aliqua. Dolor sit amet consectetur adipiscing. Id aliquet risus feugiat', 10.09, 5.12, 9, 1.56);
  shape(s, 'diamond', 10.08, 7.27, 1.45, 1.45, { fill: C.TAN, text: '1', font: F.body, size: 24, color: C.W });
  shape(s, 'diamond', 10.08, 8.86, 1.45, 1.45, { fill: C.DK, text: '2', font: F.body, size: 24, color: C.W });
  body(s, T13, 11.78, 7.46, 7.31, 1.06);
  shape(s, 'diamond', 10.08, 10.44, 1.45, 1.45, { fill: C.TAN, text: '3', font: F.body, size: 24, color: C.W });
  body(s, T13, 11.78, 9.05, 7.31, 1.06);
  body(s, T13, 11.78, 10.64, 7.31, 1.06);
  rule(s, 10.17, 1.83, 1.67);
  shape(s, 'diamond', 10.08, 12.03, 1.45, 1.45, { fill: C.DK, text: '4', font: F.body, size: 24, color: C.W });
  body(s, T13, 11.78, 12.22, 7.31, 1.06);
  body(s, T1, 20.13, 8.12, 5.2, 1.56);
  head(s, 'YOUR TITTLE HERE', 20.13, 7.5, 5.2, 0.51);
  photo(s, 1.35, 1.5, 7.48, 12);
  photo(s, 20.08, 0, 5.25, 6.75);
  photo(s, 20.08, 10.5, 5.25, 4.5);
}

function slide23(s) {
  title(s, 'Our Best Portfolios', 2.08, 2.51, 9.75, 1.14);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas dolor porttitor congue massa. Fusce posuere, magna sed pulvinar sed ultricies, purus lectus malesuada libero, sit amet commodo', 2.08, 4.37, 9.75, 1.56);
  body(s, T8, 3.52, 6.63, 8.31, 1.06);
  body(s, T8, 3.52, 7.94, 8.31, 1.06);
  shape(s, 'ellipse', 2.03, 6.65, 1.03, 1.03, { fill: C.DK, text: '1', font: F.body, size: 28, bold: true, color: C.W });
  shape(s, 'ellipse', 2.03, 7.95, 1.03, 1.03, { fill: C.TAN, text: '2', font: F.body, size: 28, bold: true, color: C.W });
  rule(s, 2.16, 2.03, 1.67);
  body(s, T8, 3.52, 9.28, 8.31, 1.06);
  shape(s, 'ellipse', 2.03, 9.29, 1.03, 1.03, { fill: C.DK, text: '3', font: F.body, size: 28, bold: true, color: C.W });
  body(s, T35, 13.36, 8.12, 5.23, 1.56);
  head(s, 'YOUR TITTLE HERE', 13.36, 7.5, 5.23, 0.51);
  body(s, T35, 19.33, 12.62, 5.23, 1.56);
  head(s, 'YOUR TITTLE HERE', 19.33, 12, 5.23, 0.51);
  body(s, T8, 3.52, 10.61, 8.31, 1.06);
  shape(s, 'ellipse', 2.03, 10.62, 1.03, 1.03, { fill: C.TAN, text: '4', font: F.body, size: 28, bold: true, color: C.W });
  body(s, T8, 3.52, 11.94, 8.31, 1.06);
  shape(s, 'ellipse', 2.03, 11.95, 1.03, 1.03, { fill: C.DK, text: '5', font: F.body, size: 28, bold: true, color: C.W });
  photo(s, 13.33, -0, 11.25, 6.75);
  photo(s, 19.33, 7.5, 5.25, 3.75);
  photo(s, 13.35, 10.53, 5.23, 4.5);
}

function slide24(s) {
  shape(s, 'rect', 15.41, 0.58, 4.5, 6.6, { fill: C.TAN });
  shape(s, 'rect', 20.33, 2.17, 4.47, 5.37, { fill: C.DK });
  shape(s, 'rect', 15.41, 7.54, 4.5, 5.37, { fill: C.DK });
  shape(s, 'rect', 20.33, 7.88, 4.47, 6.62, { fill: C.TAN });
  body(s, T36, 15.93, 3.13, 3.46, 3.58, { color: C.LT, align: 'center' });
  head(s, 'TITTLE HERE', 15.93, 2.52, 3.46, 0.51, { color: C.W, align: 'center' });
  shape(s, 'diamond', 17.05, 1.06, 1.22, 1.22, { fill: C.W, text: '1', font: F.body, size: 28, bold: true, color: C.GY6 });
  body(s, T29, 20.84, 4.68, 3.46, 2.07, { color: C.PALE, align: 'center' });
  head(s, 'TITTLE HERE', 20.84, 4.08, 3.46, 0.51, { color: C.W, align: 'center' });
  shape(s, 'diamond', 21.95, 2.61, 1.22, 1.22, { fill: C.W, text: '2', font: F.body, size: 28, bold: true, color: C.GY6 });
  body(s, T29, 15.93, 10.1, 3.46, 2.07, { color: C.PALE, align: 'center' });
  head(s, 'TITTLE HERE', 15.93, 9.49, 3.46, 0.51, { color: C.W, align: 'center' });
  shape(s, 'diamond', 17.05, 7.99, 1.22, 1.22, { fill: C.W, text: '3', font: F.body, size: 28, bold: true, color: C.GY6 });
  body(s, T36, 20.84, 10.42, 3.46, 3.58, { color: C.LT, align: 'center' });
  head(s, 'TITTLE HERE', 20.84, 9.81, 3.46, 0.51, { color: C.W, align: 'center' });
  shape(s, 'diamond', 21.95, 8.33, 1.22, 1.22, { fill: C.W, text: '4', font: F.body, size: 28, bold: true, color: C.GY6 });
  title(s, 'Place Your Latest Amazing Portfolios in Here', 1.33, 8.71, 12, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est er orcu. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et', 1.33, 11.6, 12, 2.07);
  rule(s, 1.41, 8.31, 1.67);
  photo(s, 1.33, 0.75, 12, 6.75);
}

function slide25(s) {
  photo(s, 0.56, 0, 6.77, 4.49);
  photo(s, 8.08, 2.99, 6, 8.26);
  photo(s, 2.83, 8.25, 4.5, 6.75);
  title(s, 'Runnn Portfolios', 15.55, 3.58, 9.75, 1.14);
  body(s, T14, 16.98, 5.43, 8.32, 1.06);
  body(s, T14, 16.98, 6.99, 8.32, 1.06);
  shape(s, 'ellipse', 15.5, 5.49, 1.03, 1.03, { fill: C.DK, text: '1', font: F.body, size: 28, bold: true, color: C.W });
  shape(s, 'ellipse', 15.5, 7.05, 1.03, 1.03, { fill: C.TAN, text: '2', font: F.body, size: 28, bold: true, color: C.W });
  body(s, T14, 16.98, 8.57, 8.32, 1.06);
  shape(s, 'ellipse', 15.5, 8.63, 1.03, 1.03, { fill: C.DK, text: '3', font: F.body, size: 28, bold: true, color: C.W });
  rule(s, 15.64, 3.08, 1.67);
  body(s, T29, 8.11, 12, 5.97, 1.06);
  body(s, T37, 1.33, 5.87, 5.97, 1.06);
  head(s, T38, 1.33, 5.25, 5.97, 0.51);
  rule(s, 8.24, 13.61, 1.18);
  shape(s, 'rect', 1.33, 9, 2.25, 4.5, { fill: C.DK });
  body(s, T14, 16.98, 10.15, 8.32, 1.06);
  shape(s, 'ellipse', 15.5, 10.22, 1.03, 1.03, { fill: C.TAN, text: '4', font: F.body, size: 28, bold: true, color: C.W });
  rule(s, 1.4, 7.42, 1.18);
}

function slide26(s) {
  title(s, 'Our Runnn Company Amazing Portfolios Here', 2.08, 1.36, 10.5, 2.19);
  body(s, T39, 2.1, 4.86, 10.49, 1.06);
  head(s, 'YOUR GREAT TITTLE HERE', 2.1, 4.25, 10.5, 0.51);
  rule(s, 2.17, 0.88, 1.67);
  body(s, T39, 2.1, 7.28, 10.49, 1.06);
  head(s, 'YOUR GREAT TITTLE HERE', 2.1, 6.67, 10.5, 0.51);
  shape(s, 'rect', 5.08, 9.75, 4.5, 4.5, { fill: C.GY });
  shape(s, 'rect', 9.58, 9.75, 4.5, 5.25, { fill: C.TAN });
  shape(s, 'rect', 0.58, 9.75, 4.5, 4.5, { fill: C.GY6 });
  body(s, T4, 1.1, 12.41, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 1.1, 11.8, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'pie', 2.31, 10.53, 1.03, C.W);
  body(s, T4, 5.6, 12.45, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 5.6, 11.85, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'bulb', 6.82, 10.49, 1.02, C.W);
  body(s, T4, 10.1, 12.45, 3.46, 1.06, { color: C.LT, align: 'center' });
  head(s, 'SERVICE TITTLE', 10.1, 11.84, 3.46, 0.51, { color: C.W, align: 'center' });
  icon(s, 'people', 11.31, 10.5, 1.03, C.W);
  body(s, 'Lorem ipsum dolor sit amet ae consectetuer adipiscing elit. ', 14.08, 6.61, 4.5, 1.06);
  head(s, 'YOUR TITTLE HERE', 14.08, 6, 4.51, 0.51);
  body(s, T37, 19.36, 11.87, 5.97, 1.06);
  head(s, T38, 19.35, 11.25, 5.97, 0.51);
  rule(s, 19.42, 13.42, 1.18);
  photo(s, 19.33, 0.75, 6.01, 9.75);
  photo(s, 14.08, 0.75, 4.5, 4.5);
  photo(s, 14.08, 8.25, 4.5, 6);
}

function slide27(s) {
  tablet(s, 2.83, 1.68, 8.93, 11.65);
  title(s, T30, 13.33, 2.23, 10.5, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas morbin porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc ae viverra imperdiet enim. Fusce est. Vivamus a tellus.', 13.33, 5.12, 10.5, 2.07);
  rule(s, 13.41, 1.75, 1.4);
  shape(s, 'diamond', 13.33, 8.12, 1.35, 1.35, { fill: C.DK, text: '1', font: F.body, size: 28, color: C.W });
  body(s, T31, 15.03, 7.91, 8.8, 1.56);
  shape(s, 'diamond', 13.33, 9.93, 1.35, 1.35, { fill: C.TAN, text: '2', font: F.body, size: 28, color: C.W });
  body(s, T31, 15.03, 9.82, 8.8, 1.56);
  shape(s, 'diamond', 13.33, 11.88, 1.35, 1.35, { fill: C.DK, text: '3', font: F.body, size: 28, color: C.W });
  body(s, T31, 15.03, 11.77, 8.8, 1.56);
}

function slide28(s) {
  title(s, 'Our Best Double Tablet Mockup', 16.74, 2.9, 8.25, 2.19);
  body(s, T12, 16.74, 5.79, 8.25, 2.07);
  rule(s, 16.82, 2.42, 1.4);
  tablet(s, -3, 1.68, 8.93, 11.65);
  tablet(s, 6.46, 1.68, 8.93, 11.65);
  body(s, T40, 16.74, 9.19, 8.25, 1.06);
  head(s, 'YOUR TITTLE HERE', 16.74, 8.58, 8.26, 0.51);
  body(s, T40, 16.74, 11.61, 8.25, 1.06);
  head(s, 'YOUR TITTLE HERE', 16.74, 11, 8.26, 0.51);
}

function slide29(s) {
  shape(s, 'roundRect', 2.03, 11.08, 3.79, 0.75, { fill: C.TAN, radius: 0.37, text: 'LEARN MORE', font: F.head, size: 24, bold: true, color: C.W });
  phone(s, 14.05, 2.52, 4.98, 9.98);
  phone(s, 19.51, 2.56, 4.98, 9.98);
  title(s, T30, 2.11, 3.9, 10.5, 2.19);
  body(s, [{ text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas dolor porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc ae  viverra imperdiet enim. Fusce est. Vivamus a tellus.' }, { text: '', options: { breakLine: true } }, { text: 'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.', options: { breakLine: true } }], 2.11, 6.79, 10.5, 3.58);
  rule(s, 2.19, 3.42, 1.4);
}

function slide30(s) {
  phone(s, 1.9, 4.25, 4.98, 9.98);
  phone(s, 7.53, 4.25, 4.98, 9.98);
  phone(s, 13.21, 4.24, 4.98, 9.98);
  title(s, 'Our Amazing Phone Mockup Here', 0.94, 1.59, 24.78, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
  body(s, T32, 19.34, 5.52, 5.25, 1.06);
  head(s, 'YOUR TITTLE HERE', 19.34, 4.92, 5.22, 0.51);
  rule(s, 19.42, 7.15, 1.18);
  body(s, T32, 19.34, 8.79, 5.25, 1.06);
  head(s, 'YOUR TITTLE HERE', 19.34, 8.19, 5.22, 0.51);
  rule(s, 19.42, 10.42, 1.18);
  body(s, T32, 19.34, 11.92, 5.25, 1.06);
  head(s, 'YOUR TITTLE HERE', 19.34, 11.32, 5.22, 0.51);
  rule(s, 19.42, 13.55, 1.18);
}

function slide31(s) {
  shape(s, 'roundRect', 1.28, 10.07, 3.79, 0.75, { fill: C.TAN, radius: 0.37, text: 'LEARN MORE', font: F.head, size: 24, bold: true, color: C.W });
  title(s, T30, 1.36, 4.4, 8.97, 2.19);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit netus. Maecenas porttitor congue massa. Fusce posuere, magna sed ae pulvinar ultricies, purus lectus malesuada libero, sit amet eros quis urna. Nunc viverra imperdiet enim. Fusce est vivamus', 1.36, 7.29, 8.97, 2.07);
  rule(s, 1.44, 3.92, 1.4);
  laptop(s, 10.66, 2.9, 15.13, 9.11);
}

function slide32(s) {
  stepArrow(s, 2.81, 4.67, 4.61, 3.71, { fill: C.DK, shadow: true });
  stepArrow(s, 6.51, 4.67, 4.61, 3.71, { fill: C.GY, shadow: true });
  stepArrow(s, 10.23, 4.67, 4.61, 3.71, { fill: C.TAN, shadow: true });
  txt(s, 'STEP ONE', 4.32, 6.22, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5 });
  txt(s, 'STEP TWO', 8.05, 6.22, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5 });
  txt(s, 'STEP THREE', 11.81, 6.22, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5 });
  stepArrow(s, 11.82, 9.63, 4.61, 3.71, { fill: C.TAN, shadow: true, flipH: true });
  stepArrow(s, 15.52, 9.63, 4.61, 3.71, { fill: C.GY, shadow: true, flipH: true });
  stepArrow(s, 19.23, 9.63, 4.61, 3.71, { fill: C.DK, shadow: true, flipH: true });
  txt(s, 'STEP FOUR', 12.59, 11.17, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5, align: 'right' });
  txt(s, 'STEP FIVE', 16.33, 11.17, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5, align: 'right' });
  txt(s, 'STEP SIX', 20.09, 11.17, 2.25, 0.6, { font: F.headSemi, size: 22, color: C.W, lead: 1.5, align: 'right' });
  body(s, T12, 15.59, 5.79, 7.95, 2.07);
  head(s, 'YOUR TITTLE HERE', 15.59, 5.18, 7.91, 0.51);
  body(s, T12, 3.29, 10.76, 7.77, 2.07, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 3.3, 10.15, 7.73, 0.51, { align: 'right' });
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide33(s) {
  rule(s, 3.89, 8.77, 0, { color: C.TAN, width: 3, h: 0.83, dot: true });
  rule(s, 8.38, 8.77, 0, { width: 3, h: 1.58, dot: true });
  rule(s, 12.89, 8.77, 0, { color: C.TAN, width: 3, h: 0.83, dot: true });
  rule(s, 17.38, 8.77, 0, { width: 3, h: 1.58, dot: true });
  rule(s, 21.91, 8.77, 0, { color: C.TAN, width: 3, h: 0.83, dot: true });
  chevronStep(s, 1.69, 4.78, 5.2, 3.32, { fill: C.DK, shadow: true });
  chevronStep(s, 6.21, 4.78, 5.2, 3.32, { fill: C.TAN, shadow: true });
  chevronStep(s, 10.73, 4.78, 5.2, 3.32, { fill: C.DK, shadow: true });
  chevronStep(s, 15.25, 4.78, 5.2, 3.32, { fill: C.TAN, shadow: true });
  chevronStep(s, 19.77, 4.78, 5.2, 3.32, { fill: C.DK, shadow: true });
  txt(s, '01', 3.26, 5.83, 2.28, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '02', 7.76, 5.83, 2.28, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '03', 12.28, 5.83, 2.28, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '04', 16.78, 5.83, 2.28, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '05', 21.27, 5.83, 2.21, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  body(s, T2, 2, 10.8, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 1.99, 10.19, 3.77, 0.51, { align: 'center' });
  body(s, T2, 10.95, 10.8, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 10.95, 10.19, 3.77, 0.51, { align: 'center' });
  body(s, T2, 19.95, 10.8, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 19.95, 10.19, 3.77, 0.51, { align: 'center' });
  body(s, T2, 6.49, 11.47, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 6.49, 10.86, 3.77, 0.51, { align: 'center' });
  body(s, T2, 15.47, 11.47, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 15.47, 10.86, 3.77, 0.51, { align: 'center' });
  rule(s, 3.3, 12.74, 1.18);
  rule(s, 7.79, 13.42, 1.18);
  rule(s, 12.26, 12.74, 1.18);
  rule(s, 21.26, 12.74, 1.18);
  rule(s, 16.78, 13.42, 1.18);
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide34(s) {
  calloutArrow(s, 18.41, 5.23, 6.01, 3.75, { fill: C.DK, shadow: true });
  rule(s, 4.46, 8.63, 0, { color: C.DK, width: 3, h: 0.83, dot: true });
  rule(s, 9.67, 9.37, 0, { width: 3, h: 0.83, dot: true });
  rule(s, 14.96, 8.63, 0, { color: C.DK, width: 3, h: 0.83, dot: true });
  rule(s, 20.14, 9.37, 0, { width: 3, h: 0.83, dot: true });
  body(s, T2, 2.58, 10.56, 3.79, 1.56, { align: 'center' });
  calloutArrow(s, 13.03, 4.5, 6.01, 3.75, { fill: C.TAN, shadow: true });
  head(s, 'YOUR TITTLE', 2.57, 9.95, 3.77, 0.51, { align: 'center' });
  body(s, T2, 13.08, 10.56, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 13.08, 9.95, 3.77, 0.51, { align: 'center' });
  body(s, T2, 7.8, 11.65, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 7.8, 11.05, 3.77, 0.51, { align: 'center' });
  body(s, T2, 18.21, 11.65, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 18.2, 11.05, 3.77, 0.51, { align: 'center' });
  rule(s, 3.88, 12.5, 1.18);
  rule(s, 9.11, 13.58, 1.18);
  rule(s, 14.38, 12.5, 1.18);
  calloutArrow(s, 7.63, 5.23, 6.01, 3.75, { fill: C.DK, shadow: true });
  rule(s, 19.51, 13.58, 1.18);
  txt(s, '03', 13.46, 5.78, 3.75, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '04', 18.91, 6.51, 3.75, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  txt(s, '02', 8.07, 6.51, 3.75, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  calloutArrow(s, 2.24, 4.5, 6.01, 3.75, { fill: C.TAN, shadow: true });
  txt(s, '01', 2.61, 5.78, 3.75, 1.2, { font: F.headSemi, size: 48, color: C.W, lead: 1.5, align: 'center' });
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide35(s) {
  shape(s, 'hexagon', 8.35, 4.61, 5.15, 4.44, { fill: C.W, shadow: true, rotate: 90 });
  shape(s, 'hexagon', 13.09, 4.61, 5.15, 4.44, { fill: C.W, shadow: true, rotate: 90 });
  shape(s, 'hexagon', 5.98, 8.92, 5.15, 4.44, { fill: C.W, shadow: true, rotate: 90 });
  shape(s, 'hexagon', 10.72, 8.92, 5.15, 4.44, { fill: C.TAN, shadow: true, rotate: 90 });
  shape(s, 'hexagon', 15.46, 8.92, 5.15, 4.44, { fill: C.W, shadow: true, rotate: 90 });
  icon(s, 'puzzle', 16.86, 9.96, 2.35, C.TAN);
  txt(s, '567', 9.59, 5.87, 2.66, 1.21, { font: F.headSemi, size: 66, bold: true, color: C.HEAD, align: 'center' });
  body(s, 'Your Tittle Here', 9.1, 7.23, 3.79, 0.55, { align: 'center' });
  txt(s, '180', 14.26, 5.87, 2.66, 1.21, { font: F.headSemi, size: 66, bold: true, color: C.HEAD, align: 'center' });
  body(s, 'Your Tittle Here', 13.77, 7.23, 3.79, 0.55, { align: 'center' });
  txt(s, '456', 7.15, 10.18, 2.66, 1.21, { font: F.headSemi, size: 66, bold: true, color: C.HEAD, align: 'center' });
  body(s, 'Your Tittle Here', 6.66, 11.54, 3.79, 0.55, { align: 'center' });
  txt(s, '456', 11.89, 10.14, 2.66, 1.21, { font: F.head, size: 66, bold: true, color: C.W, align: 'center' });
  txt(s, 'Your Tittle Here', 11.39, 11.49, 3.79, 0.64, { font: F.body, size: 24, color: C.LT, lead: 1.5, align: 'center' });
  body(s, T1, 2.5, 6.35, 5.25, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.5, 5.74, 5.25, 0.51, { align: 'right' });
  body(s, T41, 1.75, 10.66, 3.75, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE', 1.75, 10.05, 3.75, 0.51, { align: 'right' });
  body(s, T1, 18.92, 6.35, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 18.92, 5.74, 5.25, 0.51);
  body(s, T41, 21.17, 10.66, 3.75, 1.56);
  head(s, 'YOUR TITTLE', 21.17, 10.05, 3.75, 0.51);
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide36(s) {
  shape(s, 'hexagon', 3.15, 5.07, 4.35, 3.75, { fill: C.TAN, alpha: 5.1, shadow: true });
  shape(s, 'hexagon', 6.2, 7.37, 1.69, 1.45, { fill: C.W, shadow: true, text: '01', font: F.headSemi, size: 36, bold: true, color: C.HEAD });
  icon(s, 'bulb', 4.95, 6.35, 1.21, C.W);
  shape(s, 'hexagon', 3.61, 5.47, 3.44, 2.96, { line: C.W });
  shape(s, 'hexagon', 8.36, 6.24, 4.35, 3.75, { fill: C.DK, alpha: 5.1, shadow: true });
  icon(s, 'network', 9.97, 7.55, 1.14, C.W);
  shape(s, 'hexagon', 11.41, 8.54, 1.69, 1.45, { fill: C.W, shadow: true, text: '02', font: F.headSemi, size: 36, bold: true, color: C.HEAD });
  shape(s, 'hexagon', 8.82, 6.63, 3.44, 2.96, { line: C.W });
  shape(s, 'hexagon', 13.57, 5.07, 4.35, 3.75, { fill: C.TAN, alpha: 5.1, shadow: true });
  icon(s, 'pie', 15.2, 6.4, 1.09, C.W);
  shape(s, 'hexagon', 16.61, 7.37, 1.69, 1.45, { fill: C.W, shadow: true, text: '03', font: F.headSemi, size: 36, bold: true, color: C.HEAD });
  shape(s, 'hexagon', 14.02, 5.47, 3.44, 2.96, { line: C.W });
  shape(s, 'hexagon', 18.78, 6.24, 4.35, 3.75, { fill: C.DK, alpha: 5.1, shadow: true });
  icon(s, 'chart', 20.46, 7.63, 0.98, C.W);
  shape(s, 'hexagon', 21.82, 8.54, 1.69, 1.45, { fill: C.W, shadow: true, text: '04', font: F.headSemi, size: 36, bold: true, color: C.HEAD });
  shape(s, 'hexagon', 19.23, 6.63, 3.44, 2.96, { line: C.W });
  body(s, T2, 3.43, 10.23, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 3.44, 9.62, 3.77, 0.51, { align: 'center' });
  body(s, T2, 13.85, 10.23, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 13.85, 9.62, 3.77, 0.51, { align: 'center' });
  body(s, T2, 8.64, 11.39, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 8.65, 10.78, 3.77, 0.51, { align: 'center' });
  body(s, T2, 19.05, 11.39, 3.79, 1.56, { align: 'center' });
  head(s, 'YOUR TITTLE', 19.06, 10.78, 3.77, 0.51, { align: 'center' });
  rule(s, 9.94, 13.42, 1.18, { color: C.TAN });
  rule(s, 4.74, 12.25, 1.18, { color: C.TAN });
  rule(s, 20.36, 13.42, 1.18, { color: C.TAN });
  rule(s, 15.15, 12.25, 1.18, { color: C.TAN });
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide37(s) {
  shape(s, 'teardrop', 13.6, 9.13, 4.21, 4.21, { fill: C.TAN, shadow: true, rotate: -90, radius: 4.76 });
  shape(s, 'teardrop', 9.13, 9.13, 3.96, 3.96, { fill: C.DK, shadow: true, radius: 4.47 });
  shape(s, 'teardrop', 13.6, 4.67, 3.96, 3.96, { fill: C.DK, shadow: true, rotate: -90, flipH: true, radius: 4.47 });
  shape(s, 'teardrop', 9.4, 4.94, 3.69, 3.69, { fill: C.TAN, shadow: true, rotate: 180, flipH: true, radius: 4.16 });
  icon(s, 'network', 14.93, 6.03, 1.31, C.W);
  icon(s, 'bars', 15.12, 10.6, 0.95, C.W);
  icon(s, 'calendar', 10.54, 10.5, 1.14, C.W);
  icon(s, 'bulb', 10.85, 6.26, 1.06, C.W);
  body(s, T1, 18.57, 6.27, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 18.58, 5.67, 5.21, 0.51);
  body(s, T1, 18.57, 10.02, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 18.58, 9.42, 5.21, 0.51);
  body(s, T1, 2.84, 6.27, 5.25, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.84, 5.67, 5.25, 0.51, { align: 'right' });
  body(s, T1, 2.84, 10.02, 5.25, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.84, 9.42, 5.25, 0.51, { align: 'right' });
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide38(s) {
  shape(s, 'hexagon', 11.48, 3.95, 3.63, 3.23, { fill: C.TAN, shadow: true });
  shape(s, 'hexagon', 11.48, 10.85, 3.63, 3.23, { fill: C.DK, shadow: true });
  shape(s, 'hexagon', 8.46, 9.13, 3.63, 3.23, { fill: C.TAN, shadow: true });
  shape(s, 'hexagon', 14.56, 9.13, 3.63, 3.23, { fill: C.TAN, shadow: true });
  shape(s, 'hexagon', 8.46, 5.66, 3.63, 3.23, { fill: C.DK, shadow: true });
  shape(s, 'hexagon', 14.56, 5.66, 3.63, 3.23, { fill: C.DK, shadow: true });
  shape(s, 'hexagon', 11.48, 7.43, 3.63, 3.23, { fill: C.W, shadow: true });
  txt(s, '02', 12.59, 5.01, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  txt(s, '01', 9.57, 6.72, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  txt(s, '03', 15.67, 6.72, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  txt(s, '06', 9.57, 10.19, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  txt(s, '04', 15.67, 10.19, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  txt(s, '05', 12.59, 11.91, 1.42, 1.11, { font: F.headSemi, size: 60, bold: true, color: C.W, align: 'center', rotate: -21.6 });
  icon(s, 'network', 12.73, 8.51, 1.14, C.TAN);
  body(s, T1, 19.36, 5.36, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 19.36, 4.76, 5.25, 0.52);
  body(s, T1, 19.32, 8.54, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 19.36, 7.94, 5.22, 0.51);
  body(s, T1, 19.36, 11.71, 5.25, 1.56);
  head(s, 'YOUR TITTLE HERE', 19.36, 11.1, 5.25, 0.52);
  body(s, T1, 2.06, 5.36, 5.27, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.08, 4.76, 5.25, 0.51, { align: 'right' });
  body(s, T1, 2.08, 8.54, 5.25, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.08, 7.94, 5.25, 0.51, { align: 'right' });
  body(s, T1, 2.06, 11.71, 5.27, 1.56, { align: 'right' });
  head(s, 'YOUR TITTLE HERE', 2.08, 11.1, 5.25, 0.51, { align: 'right' });
  title(s, T3, 1.71, 1.59, 23.25, 1.14, { align: 'center' });
  rule(s, 12.5, 1.09, 1.67);
}

function slide39(s) {
  backdrop(s, 'leftPanel');
  txt(s, 'Please Contact us', 2.08, 2.92, 7.53, 1.11, { font: F.title, size: 60, bold: true, color: C.W });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas proin porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', 2.11, 4.91, 10.47, 1.56, { color: C.PALE });
  head(s, 'OFFICE HOURS', 2.11, 7.35, 5.21, 0.51, { color: C.W });
  body(s, 'Monday – Friday\n08:00 – 16:00\nSaturday – Sunday\n11:00 – 14:00', 2.11, 7.98, 5.22, 2.07, { color: C.PALE });
  head(s, 'GET IN TOUCH', 2.11, 10.9, 5.25, 0.51, { color: C.W });
  body(s, '(+00) 1234 5678 999\n(0987) 654321', 2.11, 11.53, 5.22, 1.06, { color: C.PALE });
  body(s, '13 Petter Saint Street ,\nMother Road, EUR 12345\nEurope Eastern', 8.08, 7.99, 4.5, 1.56, { color: C.PALE });
  head(s, 'OUR ADDRESS', 8.08, 7.37, 4.49, 0.71, { color: C.W });
  body(s, 'www.Runnn.com\nhello@Runnn.com', 8.08, 11.53, 4.5, 1.06, { color: C.PALE });
  head(s, 'FOLLOW US', 8.08, 10.9, 4.48, 0.66, { color: C.W });
  rule(s, 2.16, 2.4, 1.67);
  photo(s, 14.08, 0.75, 11.25, 14.25, { shape: 'round1Rect', radius: 2.22 });
}

function slide40(s) {
  photo(s, 0, 0, 26.66, 15, { color: C.PH2 });
  txt(s, [{ text: 'THANK Y' }, { text: 'O', options: { italic: true } }, { text: 'U' }, { text: '.', options: { color: C.TAN } }], 6.49, 5.75, 13.69, 2.52, { font: F.black, size: 144, bold: true, color: C.INK, align: 'center' });
  txt(s, '“runnn” simple business & finance template', 6.49, 8.25, 13.69, 0.81, { font: F.semi, size: 42, color: C.GY6, align: 'center' });
  shape(s, 'ellipse', 12.84, 14.01, 0.19, 0.19, { fill: C.GY4, rotate: -90 });
  shape(s, 'ellipse', 13.24, 14.01, 0.19, 0.19, { fill: C.GY4, rotate: -90 });
  shape(s, 'ellipse', 13.63, 14.01, 0.19, 0.19, { fill: C.GY, rotate: -90 });
  txt(s, 'Finance', 24.43, 0.82, 1.25, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY, align: 'right' });
  txt(s, 'Runnn.', 24.56, 13.81, 1.12, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY, align: 'right' });
  txt(s, 'Template', 0.98, 13.81, 1.49, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY });
  txt(s, 'Business', 1.39, 0.82, 1.38, 0.44, { font: F.semi, size: 20, bold: true, color: C.GY });
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40];

// ------------------------------------------------------------------- assemble

function buildDeck() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'RUNNN', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'RUNNN';
  pptx.title = 'Runnn — simple business & finance presentation';
  pptx.defineSlideMaster({ title: 'RUNNN_MASTER', background: { color: C.W } });
  SLIDES.forEach((fn, i) => {
    const s = pptx.addSlide({ masterName: 'RUNNN_MASTER' });
    pageNum(s, i + 1);
    fn(s);
  });
  return pptx;
}

buildDeck()
  .writeFile({ fileName: path.join(__dirname, '10bfac58-0ebb-4cfb-aa8d-713976759cd1_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
