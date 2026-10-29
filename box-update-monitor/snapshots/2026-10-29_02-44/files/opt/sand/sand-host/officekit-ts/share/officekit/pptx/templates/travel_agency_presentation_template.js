/**
 * Travia - Travel Agency Presentation Template (30 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs from native shapes and text only.
 *
 * The source template ships its picture frames EMPTY (they are `pic`
 * placeholders carrying a 5% dot pattern that renders as plain background), so
 * no raster art is reproduced here; the device mock-ups on slides 24-26 are
 * redrawn with `phoneMock` / `tabletMock` / `laptopMock`.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ----------------------------------------------------------------- palette */

const BLUE = '3594ED'; // accent
const NAVY = '001439'; // headings
const BODY = '000000'; // body copy
const WHITE = 'FFFFFF';
const TINT = 'DFEFFE'; // background tint

const HEAD = 'Inter SemiBold';
const TEXT = 'Lato';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------------- boilerplate */

const LOREM = {
  long: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros, vitae tempor metus rhoncus vitae. Integer sodales augue velit, at euismod mauris accumsan a. Sed erat leo, eleifend eu pulvinar nec.',
  med: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros, vitae tempor metus rhoncus vitae. Integer.',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros, vitae tempor metus rhoncus vitae.',
  sodales: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros, vitae tempor metus rhoncus vitae. Integer sodales.',
  euismod: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros, vitae tempor metus rhoncus vitae. Integer sodales augue velit, at euismod.',
  eros: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa. Fusce gravida orci eros.',
  tiny: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed.',
  posu: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posu.',
  bullet: 'Lorem ipsum dolor sit amet.',
  hero: 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.',
};
LOREM.xlSuscipit = LOREM.long + ' Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit.';
LOREM.xlMassa = LOREM.long + ' Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam eget enim posuere, suscipit nulla sed, consequat massa.';

/* ------------------------------------------------------------------ helpers */

const mix = (a, b, t) => Math.round(a + (b - a) * t);
const hex = n => n.toString(16).padStart(2, '0').toUpperCase();

/** Master background: vertical DFEFFE -> white ramp, painted as flat bands. */
function background(slide) {
  const BANDS = 30;
  for (let i = 0; i < BANDS; i++) {
    const t = i / (BANDS - 1);
    const color = hex(mix(0xdf, 0xff, t)) + hex(mix(0xef, 0xff, t)) + hex(mix(0xfe, 0xff, t));
    slide.addShape('rect', {
      x: 0, y: (SLIDE_H / BANDS) * i, w: SLIDE_W, h: SLIDE_H / BANDS + 0.01,
      fill: { color }, line: { type: 'none' },
    });
  }
}

/** Title-slide overlay: solid tint on the left fading out towards x = 9.25". */
function tintWash(slide) {
  const BANDS = 40;
  const W = 9.25;
  for (let i = 0; i < BANDS; i++) {
    const t = i / (BANDS - 1);
    const transparency = t <= 0.23 ? 0 : Math.round(((t - 0.23) / 0.77) * 100);
    slide.addShape('rect', {
      x: (W / BANDS) * i, y: 0, w: W / BANDS, h: SLIDE_H,
      fill: { color: TINT, transparency }, line: { type: 'none' },
    });
  }
}

/** Top-left "Travia" wordmark with its rounded badge and plane glyph. */
function brand(slide, color) {
  slide.addShape('roundRect', { x: 0.366, y: 0.365, w: 0.291, h: 0.291, rectRadius: 0.07, fill: { color: BLUE }, line: { type: 'none' } });
  planeIcon(slide, 0.366, 0.365, 0.291);
  slide.addText('Travia', { x: 0.74, y: 0.375, w: 0.818, h: 0.269, fontSize: 16, fontFace: HEAD, color, margin: 0, valign: 'top' });
}

/** Left-edge page rail: five numbers, the current one inside a blue dot. */
function rail(slide, current, allWhite) {
  const first = Math.floor((current - 1) / 5) * 5 + 1;
  const idx = (current - 1) % 5;              // 0 = bottom entry
  const tops = [6.912, 6.599, 6.285, 5.972, 5.659];
  slide.addShape('ellipse', { x: 0.21, y: tops[idx] - 0.09, w: 0.313, h: 0.313, fill: { color: BLUE }, line: { type: 'none' } });
  tops.forEach((y, i) => {
    slide.addText(String(first + i).padStart(2, '0'), {
      x: 0.245, y, w: 0.243, h: 0.135, fontSize: 8, bold: true, fontFace: TEXT,
      color: (i === idx || allWhite) ? WHITE : NAVY, align: 'center', valign: 'top', margin: 0,
    });
  });
}

/** Short accent rule that sits under every section heading. */
function rule(slide, x, y, color) {
  slide.addShape('line', { x, y, w: 0.648, h: 0, line: { color: color || BLUE, width: 1.5 } });
}

/** Heading line, optionally split into differently coloured runs. */
function heading(slide, x, y, w, runs, opts) {
  const o = opts || {};
  const parts = (typeof runs === 'string' ? [[runs, NAVY]] : runs)
    .map(([text, color]) => ({ text, options: { color } }));
  slide.addText(parts, {
    x, y, w, h: o.h || 0.74, fontSize: o.size || 44, fontFace: HEAD,
    align: o.align || 'left', valign: 'top', margin: 0,
  });
}

/** Justified 10pt body paragraph with 1.5 line spacing. */
function body(slide, x, y, w, h, text, opts) {
  const o = opts || {};
  slide.addText(text, {
    x, y, w, h, fontSize: o.size || 10, fontFace: TEXT, color: o.color || BODY,
    align: o.align || 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: 0,
  });
}

/** Small bold blue kicker above a paragraph. */
function kicker(slide, x, y, text, color) {
  slide.addText(text, { x, y, w: 1.6, h: 0.202, fontSize: 12, bold: true, fontFace: TEXT, color: color || BLUE, valign: 'top', margin: 0 });
}

/** Pill button: rounded blue plate with a centred white caption. */
function button(slide, x, y, w, h, label) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: h / 2, fill: { color: BLUE }, line: { type: 'none' } });
  slide.addText(label, { x, y, w, h, fontSize: 12, fontFace: TEXT, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
}

/** Flat blue block - the decorative slabs used throughout the deck. */
function slab(slide, x, y, w, h, shape) {
  slide.addShape(shape || 'rect', { x, y, w, h, fill: { color: BLUE }, line: { type: 'none' } });
}

/** Blue rounded badge - the plate behind each "Travel Tips" pictogram. */
function badge(slide, x, y, size) {
  slide.addShape('roundRect', { x, y, w: size, h: size, rectRadius: size * 0.24, fill: { color: BLUE }, line: { type: 'none' } });
}

/** Document pictogram: white page with a folded corner and four ruled lines. */
function docIcon(slide, x, y, size) {
  slide.addShape('snip1Rect', { x: x + size * 0.28, y: y + size * 0.18, w: size * 0.44, h: size * 0.6, fill: { color: WHITE }, line: { type: 'none' } });
  for (let i = 0; i < 4; i++) {
    slide.addShape('rect', { x: x + size * 0.35, y: y + size * 0.42 + i * size * 0.09, w: size * 0.3, h: size * 0.035, fill: { color: BLUE }, line: { type: 'none' } });
  }
}

/** Suitcase pictogram: handle, body and two clasp bands. */
function suitcase(slide, x, y, size) {
  slide.addShape('rect', { x: x + size * 0.42, y: y + size * 0.2, w: size * 0.16, h: size * 0.1, fill: { type: 'none' }, line: { color: WHITE, width: 1.6 } });
  slide.addShape('roundRect', { x: x + size * 0.24, y: y + size * 0.3, w: size * 0.52, h: size * 0.48, rectRadius: size * 0.04, fill: { color: WHITE }, line: { type: 'none' } });
  [0.33, 0.59].forEach(f => {
    slide.addShape('rect', { x: x + size * f, y: y + size * 0.3, w: size * 0.08, h: size * 0.48, fill: { color: BLUE }, line: { type: 'none' } });
  });
}

/** Aeroplane pictogram: the font glyph points right, so rotate it upright. */
function planeIcon(slide, x, y, size) {
  slide.addText('\u2708', { x, y, w: size, h: size, rotate: -90, fontSize: Math.round(size * 52), fontFace: 'DejaVu Sans', color: WHITE, align: 'center', valign: 'middle', margin: 0 });
}

/** Support-agent pictogram: head with headset over a shoulder block. */
function personIcon(slide, x, y, size) {
  slide.addShape('ellipse', { x: x + size * 0.28, y: y + size * 0.06, w: size * 0.38, h: size * 0.38, fill: { color: WHITE }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + size * 0.62, y: y + size * 0.16, w: size * 0.2, h: size * 0.2, fill: { color: WHITE }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + size * 0.12, y: y + size * 0.52, w: size * 0.76, h: size * 0.62, fill: { color: WHITE }, line: { type: 'none' } });
}

/** Map-pin pictogram: teardrop rotated point-down with a hollow centre. */
function mapPin(slide, x, y, size) {
  slide.addShape('teardrop', { x: x + size * 0.2, y, w: size * 0.62, h: size * 0.62, rotate: 135, fill: { color: WHITE }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + size * 0.37, y: y + size * 0.15, w: size * 0.26, h: size * 0.26, fill: { color: BLUE }, line: { type: 'none' } });
}

/** Twitter-style bird: a crescent wing plus a teardrop head. */
function birdIcon(slide, x, y, size) {
  slide.addShape('moon', { x, y, w: size, h: size, rotate: 250, fill: { color: WHITE }, line: { type: 'none' } });
  slide.addShape('teardrop', { x: x + size * 0.42, y: y + size * 0.02, w: size * 0.42, h: size * 0.42, rotate: 225, fill: { color: WHITE }, line: { type: 'none' } });
}

/** Instagram-style camera: rounded square outline, lens ring and flash dot. */
function cameraIcon(slide, x, y, size) {
  slide.addShape('roundRect', { x, y, w: size, h: size, rectRadius: size * 0.28, fill: { type: 'none' }, line: { color: WHITE, width: 1.6 } });
  slide.addShape('ellipse', { x: x + size * 0.29, y: y + size * 0.29, w: size * 0.42, h: size * 0.42, fill: { type: 'none' }, line: { color: WHITE, width: 1.6 } });
  slide.addShape('ellipse', { x: x + size * 0.72, y: y + size * 0.14, w: size * 0.12, h: size * 0.12, fill: { color: WHITE }, line: { type: 'none' } });
}

/**
 * Device mock-ups are drawn as hollow frames (stroke only) so the slide art
 * shows through the screen, exactly as the transparent PNG mock-ups did.
 * `inset` shifts the frame in from the nominal picture box.
 */
function deviceFrame(slide, x, y, w, h, thickIn, radius, inset) {
  const m = (inset || 0) + thickIn / 2;
  slide.addShape('roundRect', {
    x: x + m, y: y + m, w: w - 2 * m, h: h - 2 * m,
    rectRadius: radius, fill: { type: 'none' }, line: { color: '1C1C1E', width: thickIn * 72 },
  });
}

function phoneMock(slide, x, y, w, h) {
  deviceFrame(slide, x, y, w, h, w * 0.033, w * 0.135, w * 0.017);
  slide.addShape('roundRect', { x: x + w * 0.262, y: y + h * 0.006, w: w * 0.47, h: h * 0.05, rectRadius: h * 0.02, fill: { color: '1C1C1E' }, line: { type: 'none' } });
}

function tabletMock(slide, x, y, w, h) {
  deviceFrame(slide, x, y, w, h, w * 0.028, w * 0.028, w * 0.002);
  slide.addShape('rect', { x: x + w * 0.42, y: y + h * 0.955, w: w * 0.16, h: h * 0.008, fill: { color: '1C1C1E' }, line: { type: 'none' } });
}

function laptopMock(slide, x, y, w, h) {
  deviceFrame(slide, x + w * 0.052, y, w * 0.89, h * 0.92, w * 0.013, w * 0.012, 0);
  slide.addShape('roundRect', { x, y: y + h * 0.92, w, h: h * 0.078, rectRadius: h * 0.018, fill: { color: '4A4A4E' }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: x + w * 0.43, y: y + h * 0.928, w: w * 0.14, h: h * 0.026, rectRadius: h * 0.011, fill: { color: '2E2E32' }, line: { type: 'none' } });
}

/* -------------------------------------------------------------- slide parts */

/** Heading + rule + one or two body paragraphs - the deck's workhorse block. */
function textBlock(slide, cfg) {
  let ty = cfg.y;
  (cfg.titles || []).forEach(runs => { heading(slide, cfg.x, ty, cfg.titleW, runs, { align: cfg.align }); ty += 0.634; });
  rule(slide, cfg.ruleX !== undefined ? cfg.ruleX : cfg.x, cfg.ruleY);
  (cfg.paras || []).forEach(p => body(slide, p.x !== undefined ? p.x : cfg.x, p.y, p.w, p.h, p.text, p));
}

/* ------------------------------------------------------------ slide builders */

const slides = [];

slides.push(function title(s) {                                        // 1
  tintWash(s);
  s.addText('Travel Agency Presentation Template', { x: 1.026, y: 2.054, w: 3.62, h: 0.236, fontSize: 14, fontFace: TEXT, color: NAVY, valign: 'top', margin: 0 });
  s.addText('Travia', { x: 1.026, y: 2.171, w: 5.464, h: 1.935, fontSize: 115, fontFace: HEAD, color: NAVY, valign: 'top', margin: 0 });
  body(s, 1.026, 4.11, 4.8, 0.472, LOREM.hero, { color: NAVY });
  button(s, 1.026, 5.057, 1.645, 0.354, 'Get Started');
});

slides.push(function introduction(s) {                                 // 2
  slab(s, 9.306, 1.278, 4.028, 6.222);
  textBlock(s, {
    x: 1.026, y: 1.95, titleW: 4.097, titles: [[['Introduction', NAVY]]], ruleY: 2.878,
    paras: [{ y: 3.323, w: 4.842, h: 0.977, text: LOREM.long }, { y: 4.826, w: 4.842, h: 0.724, text: LOREM.med }],
  });
});

slides.push(function journey(s) {                                      // 3
  slab(s, 5.381, 7.135, 5.229, 0.365);
  heading(s, 6.667, 1.738, 5.229, [['The ', NAVY], ['Journey', BLUE], [' Not', NAVY]]);
  heading(s, 6.667, 2.372, 5.971, 'The Arrival Matters');
  rule(s, 6.667, 3.3);
  body(s, 6.667, 3.745, 4.842, 0.977, LOREM.long);
  body(s, 6.667, 5.248, 4.842, 0.724, LOREM.med);
});

slides.push(function investment(s) {                                   // 4
  slab(s, 12.967, 1.215, 0.371, 1.794);
  heading(s, 0.74, 1.446, 6.677, [['Investment in ', NAVY], ['Travel', BLUE]]);
  heading(s, 0.74, 2.08, 9.079, [['Is An ', NAVY], ['Investment', BLUE], [' in Yourself', NAVY]]);
  rule(s, 0.74, 3.008);
  body(s, 10.171, 4.552, 2.621, 1.734, LOREM.long);
});

slides.push(function wiseMan(s) {                                      // 5
  slab(s, 4.854, 7.135, 5.097, 0.365);
  heading(s, 0.74, 1.488, 8.01, [['Travel', BLUE], [' Makes A Wise Man', NAVY]]);
  heading(s, 0.74, 2.122, 7.802, 'Better But A Fool Worse');
  rule(s, 0.74, 3.05);
  body(s, 5.816, 4.296, 3.173, 1.481, LOREM.long);
  body(s, 5.816, 6.074, 3.173, 0.472, LOREM.tiny);
});

slides.push(function welcome(s) {                                      // 6
  slab(s, 10.431, 0, 2.903, 0.365);
  heading(s, 6.102, 1.393, 6.114, [['Welcome ', NAVY], ['Message', BLUE]]);
  rule(s, 6.102, 2.321);
  body(s, 6.102, 5.444, 6.437, 0.724, LOREM.long);
  body(s, 6.102, 6.405, 5.776, 0.472, LOREM.short);
});

slides.push(function letsTalk(s) {                                     // 7
  slab(s, 2.5, -0.009, 5.724, 1.731);
  heading(s, 1.026, 5.367, 7.199, [['Let\u2019s Talk About ', NAVY], ['Travia', BLUE]]);
  rule(s, 1.026, 6.295);
  body(s, 9.14, 5.493, 3.626, 1.229, LOREM.long);
});

slides.push(function about(s) {                                        // 8
  slab(s, 6.963, -0.239, 8.311, 8.311, 'ellipse');
  heading(s, 1.026, 1.99, 4.301, [['About ', NAVY], ['Travia', BLUE]]);
  rule(s, 1.026, 2.918);
  body(s, 1.026, 3.364, 4.842, 0.977, LOREM.long);
  body(s, 1.026, 4.786, 4.842, 0.724, LOREM.med);
});

slides.push(function purposeA(s) {                                     // 9
  slab(s, 4.034, 6.42, 2.952, 1.08);
  heading(s, 7.708, 1.738, 4.222, [['The ', NAVY], ['Purpose', BLUE]]);
  heading(s, 7.708, 2.372, 3.861, 'Of Our Trip');
  rule(s, 7.708, 3.3);
  body(s, 7.708, 3.745, 4.833, 0.977, LOREM.long);
  body(s, 7.704, 5.131, 4.222, 0.724, LOREM.short);
});

slides.push(function purposeB(s) {                                     // 10
  slab(s, 7.431, 0, 5.903, 0.365);
  heading(s, 1.026, 1.335, 4.222, [['The ', NAVY], ['Purpose', BLUE]]);
  heading(s, 1.026, 1.969, 3.861, 'Of Our Trip');
  rule(s, 1.026, 2.897);
  ['Home', 'Service', 'Gallery'].forEach((label, i) => {
    const x = 7.431 + i * 1.193;
    slide_pill(s, x, label);
  });
  body(s, 7.431, 1.859, 4.833, 0.977, LOREM.long);
});

function slide_pill(s, x, label) {
  s.addShape('roundRect', { x, y: 1.447, w: 1.09, h: 0.278, rectRadius: 0.06, fill: { color: BLUE }, line: { type: 'none' } });
  s.addText(label, { x: x + 0.102, y: 1.502, w: 0.884, h: 0.168, fontSize: 10, fontFace: TEXT, color: WHITE, align: 'center', valign: 'top', margin: 0 });
}

slides.push(function beach(s) {                                        // 11
  slab(s, 7.481, 1.222, 5.852, 6.278);
  heading(s, 1.026, 2.037, 5.852, [['Trip On ', NAVY], ['The Beach', BLUE]]);
  rule(s, 1.026, 2.965);
  body(s, 1.026, 3.411, 5.641, 0.724, LOREM.long);
  body(s, 1.026, 4.487, 4.842, 0.977, LOREM.long);
});

slides.push(function travelTips(s) {                                   // 12
  slab(s, 12.967, 0, 0.366, 7.5);
  s.addShape('roundRect', { x: 0.533, y: 0.531, w: 0.291, h: 0.291, rectRadius: 0.07, fill: { color: BLUE }, line: { type: 'none' } });
  heading(s, 1.026, 4.501, 3.903, [['Travel ', NAVY], ['Tips', BLUE]]);
  rule(s, 1.026, 5.429);
  body(s, 1.026, 5.874, 3.903, 0.724, LOREM.short);
  badge(s, 7.097, 4.557, 0.751);
  docIcon(s, 7.097, 4.557, 0.751);
  badge(s, 8.325, 4.563, 0.751);
  suitcase(s, 8.325, 4.563, 0.751);
  badge(s, 9.552, 4.563, 0.751);
  planeIcon(s, 9.552, 4.563, 0.751);
  s.addText([
    { text: 'Documentation', options: { breakLine: true } },
    { text: 'Packing', options: { breakLine: true } },
    { text: 'Safety' },
  ], { x: 10.72, y: 4.635, w: 1.441, h: 0.606, fontSize: 12, bold: true, fontFace: TEXT, color: NAVY, valign: 'top', margin: 0 });
  body(s, 7.097, 5.619, 4.842, 0.977, LOREM.long);
});

slides.push(function adventureA(s) {                                   // 13
  slab(s, 10.527, 0, 2.806, 7.5);
  heading(s, 1.026, 1.738, 3.577, 'Adventure');
  heading(s, 1.026, 2.372, 4.101, [['Experiences', BLUE]]);
  rule(s, 1.026, 3.3);
  body(s, 1.026, 3.745, 4.833, 0.977, LOREM.long);
  kicker(s, 1.026, 5.075, 'Scuba Diving');
  body(s, 1.026, 5.355, 4.363, 0.724, LOREM.sodales);
});

slides.push(function adventureB(s) {                                   // 14
  slab(s, 6.667, 0, 3.097, 0.365);
  heading(s, 4.794, 1.985, 3.577, 'Adventure');
  heading(s, 4.794, 2.619, 4.04, [['Experiences', BLUE]]);
  rule(s, 4.794, 3.547);
  kicker(s, 4.794, 4.008, 'Best Experience');
  body(s, 4.794, 4.312, 3.714, 1.229, LOREM.long);
  body(s, 4.794, 5.777, 3.714, 0.472, LOREM.tiny);
});

slides.push(function adventureC(s) {                                   // 15
  slab(s, 3.639, 7.135, 9.694, 0.365);
  heading(s, 4.549, 4.756, 3.577, 'Adventure');
  heading(s, 4.549, 5.39, 4.04, [['Experiences', BLUE]]);
  rule(s, 4.549, 6.318);
  kicker(s, 9.66, 4.873, 'Best Experience');
  body(s, 9.66, 5.178, 2.711, 1.229, LOREM.short);
});

slides.push(function breakSlide(s) {                                   // 16
  tintWash(s);
  s.addText('Travel Agency Presentation Template', { x: 1.149, y: 2.054, w: 3.538, h: 0.236, fontSize: 14, fontFace: TEXT, color: NAVY, valign: 'top', margin: 0 });
  s.addText('Break Slide', { x: 1.087, y: 2.217, w: 6.98, h: 1.481, fontSize: 88, fontFace: HEAD, color: NAVY, valign: 'top', margin: 0 });
  body(s, 1.149, 5.972, 3.784, 0.724, LOREM.short, { color: NAVY });
});

slides.push(function meetTeam(s) {                                     // 17
  slab(s, 10.931, 0, 2.403, 0.365);
  heading(s, 8.866, 2.132, 3.277, 'Meet Our');
  heading(s, 8.866, 2.765, 4.101, [['Great Team', BLUE]]);
  rule(s, 8.866, 3.694);
  body(s, 8.866, 4.139, 3.577, 1.229, LOREM.long);
  body(s, 8.866, 5.659, 3.577, 0.724, LOREM.eros);
});

slides.push(function ourTeam(s) {                                      // 18
  heading(s, 3.846, 1.108, 5.641, [['Our ', NAVY], ['Great Team', BLUE]], { align: 'center' });
  rule(s, 6.343, 2.036);
  body(s, 2.045, 6.03, 9.244, 0.724, LOREM.xlMassa, { align: 'center' });
});

slides.push(function travelGuide(s) {                                  // 19
  slab(s, 8.597, 0, 4.736, 1.806);
  heading(s, 1.026, 2.575, 5.325, [['Our', NAVY], [' Travel Guide', BLUE]]);
  rule(s, 1.026, 3.503);
  body(s, 1.026, 3.949, 5.325, 1.229, LOREM.xlSuscipit);
});

slides.push(function greatGallery(s) {                                 // 20
  slab(s, 5.337, 7.135, 7.997, 0.365);
  heading(s, 7.078, 2.575, 5.641, [['Our', NAVY], [' Great Gallery', BLUE]]);
  rule(s, 7.078, 3.503);
  body(s, 7.078, 3.949, 5.325, 1.229, LOREM.xlSuscipit);
});

slides.push(function travelGallery(s) {                                // 21
  slab(s, 12.967, 4.892, 0.366, 1.663);
  heading(s, 6.667, 1.607, 5.641, [['Our', NAVY], [' Travel Gallery', BLUE]]);
  rule(s, 6.667, 2.535);
  body(s, 6.667, 2.98, 5.325, 1.229, LOREM.xlSuscipit);
});

slides.push(function kyoto(s) {                                        // 22
  slab(s, 7.699, 1.227, 0.515, 5.046);
  heading(s, 1.026, 1.738, 5.362, 'Discover Kyoto\u2019s');
  heading(s, 1.026, 2.372, 5.056, [['Tranquil Beauty', BLUE]]);
  rule(s, 1.026, 3.3);
  body(s, 1.026, 3.745, 4.833, 0.977, LOREM.long);
  kicker(s, 1.026, 5.075, 'Kyoto Japan');
  body(s, 1.026, 5.355, 4.363, 0.724, LOREM.sodales);
});

slides.push(function packages(s) {                                     // 23
  heading(s, 0.968, 1.372, 5.699, [['Travel ', NAVY], ['Packages', BLUE]]);
  rule(s, 0.968, 2.3);
  body(s, 0.968, 2.693, 7.063, 0.472, LOREM.euismod);
  kicker(s, 9.903, 5.0, 'Kyoto Japan');
  body(s, 9.903, 5.28, 2.515, 1.229, LOREM.sodales);
});

slides.push(function culturalTours(s) {                                // 24
  slab(s, 10.398, 0, 2.935, 7.5);
  heading(s, 4.204, 1.108, 4.925, [['Cultural ', BLUE], ['Tours', NAVY]], { align: 'center' });
  rule(s, 6.343, 2.036);
  kicker(s, 1.335, 3.104, 'Heritages Tour');
  body(s, 1.335, 3.401, 4.479, 0.977, LOREM.long);
  body(s, 1.335, 4.624, 3.988, 0.724, LOREM.short);
  button(s, 1.333, 5.678, 1.645, 0.354, 'Read More');
  tabletMock(s, 6.991, 2.47, 5.733, 4.196);
});

slides.push(function phoneMockup(s) {                                  // 25
  phoneMock(s, 1.072, 1.077, 3.008, 6.081);
  phoneMock(s, 4.416, 1.842, 2.251, 4.552);
  slab(s, 7.782, 7.158, 2.523, 0.342);
  heading(s, 7.782, 1.521, 3.577, 'Our Device');
  heading(s, 7.782, 2.154, 2.936, [['Mockup', BLUE]]);
  rule(s, 7.782, 3.083);
  body(s, 7.782, 3.528, 4.479, 0.977, LOREM.long);
  kicker(s, 7.782, 4.951, 'Best Experience');
  body(s, 7.782, 5.255, 3.988, 0.724, LOREM.short);
});

slides.push(function laptopMockup(s) {                                 // 26
  slab(s, 7.539, 0, 4.243, 7.5);
  laptopMock(s, 6.667, 1.24, 5.987, 3.768);
  heading(s, 1.026, 1.521, 3.577, 'Our Device');
  heading(s, 1.026, 2.154, 2.936, [['Mockup', BLUE]]);
  rule(s, 1.026, 3.083);
  body(s, 1.026, 3.528, 4.479, 0.977, LOREM.long);
  kicker(s, 1.026, 4.951, 'Best Experience');
  body(s, 1.026, 5.255, 3.988, 0.724, LOREM.short);
  body(s, 7.975, 5.622, 3.371, 0.977, LOREM.euismod, { color: WHITE });
});

slides.push(function testimonials(s) {                                 // 27
  slab(s, 6.062, 0, 0.605, 1.728);
  slab(s, 6.468, 4.542, 3.438, 2.958);
  heading(s, 1.026, 2.132, 3.422, 'Customer');
  heading(s, 1.026, 2.765, 4.131, [['Testimonials', BLUE]]);
  rule(s, 1.026, 3.694);
  body(s, 1.026, 4.139, 3.808, 1.229, LOREM.long);
  s.addText('Logan Jr.', { x: 7.311, y: 5.032, w: 1.732, h: 0.337, fontSize: 20, bold: true, fontFace: TEXT, color: WHITE, align: 'center', valign: 'top', margin: 0 });
  body(s, 6.938, 5.603, 2.481, 1.229, LOREM.sodales, { color: WHITE, align: 'center' });
  s.addText('Maria Sriya', { x: 10.749, y: 5.032, w: 1.732, h: 0.337, fontSize: 20, bold: true, fontFace: TEXT, color: BLUE, align: 'center', valign: 'top', margin: 0 });
  body(s, 10.374, 5.603, 2.481, 1.229, LOREM.sodales, { color: NAVY, align: 'center' });
});

const PRICING = [
  { x: 1.597, name: 'Reguler Package', price: '$99', term: '/2 Days', priceX: 2.032, priceW: 1.98 },
  { x: 5.242, name: 'Couple Package', price: '$149', term: '/3 Days', priceX: 5.54, priceW: 2.253 },
  { x: 8.886, name: 'VIP Package', price: '$199', term: '/5 Days', priceX: 9.102, priceW: 2.419 },
];

slides.push(function pricing(s) {                                      // 28
  heading(s, 4.204, 1.108, 4.925, [['Pricing', BLUE], [' Table', NAVY]], { align: 'center' });
  rule(s, 6.343, 2.036);
  body(s, 2.045, 2.222, 9.244, 0.472, LOREM.long, { align: 'center' });
  PRICING.forEach(card => {
    s.addShape('roundRect', { x: card.x, y: 3.202, w: 2.85, h: 3.637, rectRadius: 0.204, fill: { color: BLUE }, line: { type: 'none' } });
    s.addShape('round2SameRect', { x: card.x, y: 3.202, w: 2.85, h: 0.579, fill: { color: NAVY }, line: { type: 'none' } });
    s.addText(card.name, { x: card.x + 0.397, y: 3.357, w: 2.056, h: 0.269, fontSize: 16, bold: true, fontFace: TEXT, color: WHITE, align: 'center', valign: 'top', margin: 0 });
    s.addText([
      { text: card.price, options: { fontSize: 44 } },
      { text: card.term, options: { fontSize: 12 } },
    ], { x: card.priceX, y: 3.979, w: card.priceW, h: 0.74, bold: true, fontFace: TEXT, color: WHITE, align: 'center', valign: 'top', margin: 0 });
    s.addText(LOREM.posu, { x: card.x + 0.298, y: 4.743, w: 2.253, h: 0.377, fontSize: 8, fontFace: TEXT, color: WHITE, align: 'center', lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
    [5.379, 5.681, 5.983, 6.286].forEach((y, i) => {
      s.addShape('ellipse', { x: card.x + 0.396, y, w: 0.224, h: 0.224, fill: { color: NAVY }, line: { type: 'none' } });
      s.addText('\u2713', { x: card.x + 0.396, y, w: 0.224, h: 0.224, fontSize: 10, fontFace: 'DejaVu Sans', color: WHITE, align: 'center', valign: 'middle', margin: 0 });
      s.addText(LOREM.bullet, { x: card.x + 0.763, y: y + 0.027, w: 1.691, h: 0.168, fontSize: 10, fontFace: TEXT, color: WHITE, valign: 'top', margin: 0 });
    });
  });
});

slides.push(function contact(s) {                                      // 29
  slab(s, 6.468, 0, 6.499, 7.135);
  s.addShape('line', { x: 7.208, y: 0.427, w: 5.019, h: 0, line: { color: WHITE, width: 1.5 } });
  heading(s, 1.257, 5.091, 3.577, 'Contact');
  heading(s, 1.257, 5.725, 3.988, [['Information', BLUE]]);
  rule(s, 1.257, 6.653);
  s.addText('Social Media', { x: 7.289, y: 5.109, w: 1.104, h: 0.202, fontSize: 12, bold: true, fontFace: TEXT, color: WHITE, valign: 'top', margin: 0 });
  s.addText([
    { text: '@Travia', options: { breakLine: true } },
    { text: 'Travia Marketing', options: { breakLine: true } },
    { text: 'The Travia' },
  ], { x: 7.289, y: 5.479, w: 1.642, h: 0.733, fontSize: 10, fontFace: TEXT, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  s.addText('Contact', { x: 9.677, y: 5.097, w: 0.82, h: 0.202, fontSize: 12, bold: true, fontFace: TEXT, color: WHITE, valign: 'top', margin: 0 });
  s.addText([
    { text: 'info@traviamarket.com', options: { breakLine: true } },
    { text: '+22 123 456 7890', options: { breakLine: true } },
    { text: '325 East 38th Street, New York, NY, 10019, USA' },
  ], { x: 9.962, y: 5.479, w: 2.346, h: 0.986, fontSize: 10, fontFace: TEXT, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  s.addText('\u2709', { x: 9.665, y: 5.536, w: 0.167, h: 0.167, fontSize: 9, fontFace: 'DejaVu Sans', color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  personIcon(s, 9.665, 5.78, 0.167);
  mapPin(s, 9.667, 6.046, 0.167);
});

slides.push(function thankYou(s) {                                     // 30
  s.addText('Travel Agency Presentation Template', { x: 4.857, y: 2.583, w: 3.62, h: 0.236, fontSize: 14, fontFace: TEXT, color: WHITE, align: 'center', valign: 'top', margin: 0 });
  s.addText('Thank You', { x: 2.797, y: 2.7, w: 7.739, h: 1.616, fontSize: 96, fontFace: HEAD, color: WHITE, align: 'center', valign: 'top', margin: 0 });
  s.addShape('roundRect', { x: 5.542, y: 4.434, w: 2.25, h: 0.484, rectRadius: 0.08, fill: { color: BLUE }, line: { type: 'none' } });
  birdIcon(s, 5.92, 4.55, 0.3);
  s.addText('f', { x: 6.53, y: 4.5, w: 0.3, h: 0.35, fontSize: 20, bold: true, fontFace: 'DejaVu Sans', color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  cameraIcon(s, 7.16, 4.53, 0.28);
});

/* ------------------------------------------------------------------- render */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'TRAVIA', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'TRAVIA';
  pptx.author = 'Travia';
  pptx.title = 'Travia - Travel Agency Presentation Template';

  slides.forEach((buildSlide, i) => {
    const slide = pptx.addSlide();
    background(slide);
    buildSlide(slide);
    const n = i + 1;
    brand(slide, n === 30 ? WHITE : NAVY);
    rail(slide, n, n === 30);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '1950fb45-730b-438c-9738-f756affd946c_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
