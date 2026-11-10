/**
 * ESSENTIAL Presentation — "Business Plan" deck (30 slides, 16:9 / 13.333" x 7.5").
 *
 * Recreated with pptxgenjs only. Photographic content in the source deck lives in
 * empty picture frames; each one is redrawn here as a light-gray `panel()` rectangle
 * of the same position and size. Icon artwork is redrawn with native shapes.
 *
 * Run: node 17a8c2ed-540e-4449-85f7-fefd9ee85b1b_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '0C0C0C';    // near-black used for headlines and the logo
const NAVY = '061320';    // dark blue used for some headlines / labels
const BODY = '262626';    // paragraph gray
const GRAY = 'BFBFBF';    // hairline rules, muted chart bars, "Presentation"
const MID = '7F7F7F';    // mid-gray infographic drops
const LIGHT = 'F2F2F2';    // image panels, chart tracks, quote cards
const WHITE = 'FFFFFF';

const DISPLAY = 'Bebas Neue';  // condensed all-caps headlines
const SANS = 'Rubik';       // body copy and labels
const ALT = 'Poppins';     // side tab + a few sub-headings

/* ------------------------------------------------------- text helpers
 * Every text frame in the deck is top-anchored with zero internal insets,
 * so the (x, y, w, h) values below are the raw frame rectangles.
 */

function text(s, str, x, y, w, h, o) {
  o = o || {};
  s.addText(str, {
    x: x, y: y, w: w, h: h,
    fontFace: o.font || SANS,
    fontSize: o.size,
    bold: !!o.bold,
    color: o.color || BODY,
    align: o.align || 'left',
    valign: 'top',
    margin: 0,
    lineSpacingMultiple: o.lineSpacing
  });
}

// Big Bebas Neue headline ("BUSINESS", "OUR HISTORY", ...).
function display(s, str, x, y, w, h, size, o) {
  o = o || {};
  text(s, str, x, y, w, h, { font: DISPLAY, size: size, color: o.color || INK, align: o.align });
}

// 14pt bold eyebrow above a headline ("About Us", "People", ...).
function kicker(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, { size: 14, bold: true, color: o.color || INK, align: o.align, font: o.font });
}

// 16pt bold section sub-heading ("Your Awesome Title Here", ...).
function subhead(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, { size: 16, bold: true, color: o.color || INK, align: o.align, font: o.font });
}

// 10pt paragraph copy, 150% leading.
function body(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, { size: 10, color: o.color || BODY, align: o.align, lineSpacing: 1.5 });
}

// 12pt single-line copy (contact details, chart value labels).
function small(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, { size: 12, color: o.color || BODY, align: o.align });
}

// Oversized opening quotation mark above a pull quote.
function quoteMark(s, x, y, w, color) {
  text(s, '\u201C', x, y, w, 0.471, { size: 28, bold: true, color: color });
}

/* --------------------------------------------------------- shape helpers */

// Stand-in for a photo frame: flat light-gray rectangle.
function panel(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color || LIGHT } });
}

function roundPanel(s, x, y, w, h, radius, color) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: radius, fill: { color: color || LIGHT } });
}

// Pull-quote card: light panel under a very soft drop shadow.
function quoteCard(s, x, y, w, h) {
  s.addShape('rect', {
    x: x, y: y, w: w, h: h, fill: { color: LIGHT },
    shadow: { type: 'outer', color: '000000', opacity: 0.09, blur: 14, offset: 4, angle: 90 }
  });
}

function rule(s, x, y, w, h) {
  s.addShape('line', { x: x, y: y, w: w, h: h, line: { color: GRAY, width: 0.75 } });
}

// Half-disc: flat edge on one side, semicircular bulge on the other.
function halfDisc(s, x, y, w, h, facing, color) {
  const k = 0.5523; // circular bezier constant
  const pts = facing === 'down'
    ? [{ x: 0, y: 0, moveTo: true },
       { x: w / 2, y: h, curve: { type: 'cubic', x1: 0, y1: h * k, x2: w / 2 - (w / 2) * k, y2: h } },
       { x: w, y: 0, curve: { type: 'cubic', x1: w / 2 + (w / 2) * k, y1: h, x2: w, y2: h * k } }]
    : [{ x: 0, y: 0, moveTo: true },
       { x: w, y: h / 2, curve: { type: 'cubic', x1: w * k, y1: 0, x2: w, y2: h / 2 - (h / 2) * k } },
       { x: 0, y: h, curve: { type: 'cubic', x1: w, y1: h / 2 + (h / 2) * k, x2: w * k, y2: h } }];
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts.concat([{ close: true }]) });
}

// Logo: three nested half-discs, shallowest first.
const LOGO_STEPS = [0.211, 0.348, 0.5];

function logoMark(s, x, y, w, h, facing) {
  const total = LOGO_STEPS.reduce(function (a, b) { return a + b; }, 0);
  let cursor = 0;
  LOGO_STEPS.forEach(function (step) {
    const span = (step / total) * (facing === 'down' ? h : w);
    if (facing === 'down') halfDisc(s, x, y + cursor, w, span, 'down', INK);
    else halfDisc(s, x + cursor, y, span, h, 'right', INK);
    cursor += span;
  });
}

// Furniture repeated on every slide: rotated side tab, logo, vertical hairline.
function chrome(s, font) {
  s.addText([
    { text: 'ESSENTIAL ', options: { bold: true, color: NAVY } },
    { text: 'Presentation', options: { color: GRAY } }
  ], {
    x: -0.802, y: 5.805, w: 2.24, h: 0.185, rotate: -90,
    fontFace: font || ALT, fontSize: 11, valign: 'top', margin: 0
  });
  logoMark(s, 0.145, 0.483, 0.347, 0.368, 'down');
  rule(s, 0.642, 0, 0, 7.5);
}

// Pictogram stand-ins for the deck's icon PNGs.
// `ink` is the icon color, `hole` the color showing through cut-outs.
function icon(s, kind, x, y, d, ink, hole) {
  const put = function (shape, ix, iy, iw, ih, color, extra) {
    s.addShape(shape, Object.assign({
      x: x + ix * d, y: y + iy * d, w: iw * d, h: ih * d, fill: { color: color }
    }, extra || {}));
  };
  if (kind === 'gear') {
    put('gear6', 0, 0.05, 1, 0.9, ink);
    put('ellipse', 0.36, 0.38, 0.28, 0.28, hole);
  } else if (kind === 'bulb') {
    put('ellipse', 0.14, 0.0, 0.72, 0.72, ink);        // glass
    put('ellipse', 0.24, 0.1, 0.52, 0.52, hole);
    put('rect', 0.42, 0.5, 0.16, 0.16, ink);           // neck
    put('rect', 0.34, 0.7, 0.32, 0.07, ink);           // screw threads
    put('rect', 0.34, 0.82, 0.32, 0.07, ink);
    put('trapezoid', 0.4, 0.94, 0.2, 0.06, ink);
  } else if (kind === 'photo') {
    put('rect', 0.18, 0.06, 0.82, 0.6, ink);           // back frame
    put('rect', 0.26, 0.14, 0.66, 0.44, hole);
    put('rect', 0.0, 0.26, 0.82, 0.62, ink);           // front frame
    put('rect', 0.08, 0.34, 0.66, 0.46, hole);
    put('triangle', 0.14, 0.46, 0.4, 0.32, ink);       // mountain
    put('ellipse', 0.54, 0.4, 0.12, 0.12, ink);        // sun
  } else if (kind === 'target') {
    put('donut', 0.02, 0.02, 0.96, 0.96, ink);
    put('donut', 0.26, 0.26, 0.48, 0.48, ink);
    put('ellipse', 0.42, 0.42, 0.16, 0.16, ink);
  } else if (kind === 'gem') {
    s.addShape('custGeom', {
      x: x, y: y, w: d, h: d, fill: { color: ink },
      points: [
        { x: 0.30 * d, y: 0.14 * d, moveTo: true },
        { x: 0.70 * d, y: 0.14 * d },
        { x: 0.98 * d, y: 0.42 * d },
        { x: 0.50 * d, y: 0.94 * d },
        { x: 0.02 * d, y: 0.42 * d },
        { close: true }
      ]
    });
  }
}

// Phone mockup: dark shell, white screen, notch.
function phoneMockup(s, x, y, w, h) {
  const edge = 0.04 * w;               // black bezel thickness
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.13 * w, fill: { color: INK } });
  s.addShape('roundRect', {
    x: x + edge, y: y + edge, w: w - 2 * edge, h: h - 2 * edge,
    rectRadius: 0.11 * w, fill: { color: WHITE }
  });
  s.addShape('roundRect', {
    x: x + 0.26 * w, y: y + edge, w: w * 0.48, h: 0.075 * w,
    rectRadius: 0.03 * w, fill: { color: INK }
  });
}

// Desktop monitor mockup: bezel + screen, light chin, stand, base.
function monitorMockup(s, x, y, w, h) {
  const bezelH = h * 0.717;
  const bezel = w * 0.038;             // dark border around the screen
  s.addShape('roundRect', { x: x, y: y, w: w, h: bezelH, rectRadius: 0.03, fill: { color: '1D1E20' } });
  s.addShape('rect', { x: x + bezel, y: y + bezel, w: w - 2 * bezel, h: bezelH - 2 * bezel, fill: { color: WHITE } });
  s.addShape('rect', { x: x, y: y + bezelH, w: w, h: h * 0.11, fill: { color: 'D3D3D3' } });
  s.addShape('trapezoid', { x: x + 0.396 * w, y: y + h * 0.827, w: w * 0.208, h: h * 0.155, fill: { color: 'A6A6A6' } });
  s.addShape('ellipse', { x: x + 0.31 * w, y: y + h * 0.965, w: w * 0.39, h: h * 0.035, fill: { color: 'D3D3D3' } });
}

/* ------------------------------------------------------ slide 27 bar chart */

const CHART_TOP = 0.995;
const CHART_BOTTOM = 3.988;
const BAR_W = 0.272;
// [left edge, filled height, fill, printed value]
const CHART_BARS = [
  [1.476, 1.510, GRAY, '51'], [2.344, 1.102, GRAY, '37'], [3.220, 1.102, INK, '37'],
  [4.096, 0.761, GRAY, '18'], [4.973, 2.594, INK, '87'], [5.855, 1.922, INK, '54'],
  [6.729, 1.922, GRAY, '62'], [7.610, 1.510, GRAY, '51'], [8.478, 1.102, GRAY, '37'],
  [9.355, 0.761, GRAY, '18'], [10.231, 2.594, INK, '87'], [11.113, 1.922, INK, '54'],
  [11.986, 1.660, GRAY, '49']
];

function barChart(s) {
  CHART_BARS.forEach(function (bar) {
    const x = bar[0], value = bar[1], color = bar[2], label = bar[3];
    const radius = BAR_W / 2;
    s.addShape('roundRect', { x: x, y: CHART_TOP, w: BAR_W, h: CHART_BOTTOM - CHART_TOP, rectRadius: radius, fill: { color: LIGHT } });
    s.addShape('roundRect', { x: x, y: CHART_BOTTOM - value, w: BAR_W, h: value, rectRadius: radius, fill: { color: color } });
    s.addText([{ text: label }, { text: '+', options: { superscript: true } }], {
      x: x - 0.122, y: 4.155, w: 0.5, h: 0.202,
      fontFace: SANS, fontSize: 12, color: color, align: 'center', valign: 'top', margin: 0
    });
  });
}

/* ------------------------------------------------------------ slide builders */

// 1. title slide
function slide01(s) {
  chrome(s, SANS);
  display(s, 'Business', 1.432, 2.557, 6.599, 2.323, 138);
  display(s, 'Plan', 1.366, 3.802, 6.776, 4.46, 265);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet.', 1.494, 0.851, 5.314, 0.757);
  panel(s, 8.587, 0.483, 4.3, 6.535);
}

// 2. introduction
function slide02(s) {
  chrome(s);
  display(s, 'Introduction', 1.319, 1.563, 4.899, 1.111, 66);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis.', 1.319, 4.422, 4.233, 1.515);
  subhead(s, 'WELCOME MESSAGE', 1.319, 3.993, 2.649, 0.269);
  panel(s, 6.667, 0, 6.667, 6.535);
}

// 3. welcome
function slide03(s) {
  chrome(s);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida.', 7.288, 3.75, 5.198, 1.01);
  display(s, 'Welcome To Our', 7.288, 1.318, 5.403, 1.111, 66);
  display(s, 'Company', 7.288, 2.195, 3.302, 1.111, 66);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum.', 7.288, 5.575, 5.198, 0.757);
  subhead(s, 'Your  Awesome Title Here', 7.288, 5.152, 3.163, 0.269);
  panel(s, 0.656, 0.483, 3.703, 6.535);
  panel(s, 4.462, 0.483, 1.955, 6.535);
}

// 4. company profile
function slide04(s) {
  chrome(s);
  display(s, 'Company Profile', 1.319, 1.348, 5.091, 1.01, 60);
  kicker(s, 'About Us', 1.319, 1.032, 1.083, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet.', 1.319, 5.395, 4.381, 1.01);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet.', 7.633, 5.395, 4.692, 1.01);
  subhead(s, 'Your Title Here', 7.633, 4.965, 1.864, 0.269, { color: NAVY });
  panel(s, 0.642, 2.443, 5.664, 2.539);
  panel(s, 7.633, 0.493, 4.692, 3.796);
}

// 5. our history
function slide05(s) {
  panel(s, 0.653, 3.562, 12.234, 3.455);
  panel(s, 0.653, 4.778, 5.368, 2.722, WHITE);
  chrome(s);
  display(s, 'Our History', 1.487, 5.622, 4.711, 1.212, 72);
  kicker(s, 'About Us', 1.487, 5.343, 1.063, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis.', 1.487, 1.621, 4.944, 1.262);
  subhead(s, 'Your Awesome Title Here', 1.487, 1.191, 3.08, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis.', 7.612, 1.621, 4.944, 1.262);
  rule(s, 0.642, 7.007, 5.222, 0);
}

// 6. who started the company
function slide06(s) {
  chrome(s);
  display(s, 'Who Started', 7.786, 4.925, 4.312, 0.909, 54, { color: NAVY });
  display(s, 'The Company', 8.479, 5.65, 4.212, 0.909, 54, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis.', 7.786, 1.548, 4.312, 1.515);
  subhead(s, 'Your Awesome Title Here', 7.786, 1.119, 3.056, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 7.786, 3.283, 4.312, 0.505);
  kicker(s, 'About Us', 7.786, 4.588, 1.025, 0.236);
  logoMark(s, 7.805, 5.794, 0.539, 0.509, 'right');
  panel(s, 0.652, 0, 6.015, 6.558);
}

// 7. people portrait
function slide07(s) {
  chrome(s);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis. Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar.', 7.341, 3.911, 4.944, 2.02);
  display(s, 'Briana Lovie', 7.341, 1.847, 4.375, 1.111, 66, { color: NAVY });
  subhead(s, 'About Company', 7.341, 3.482, 2.482, 0.269);
  kicker(s, 'People', 7.341, 1.569, 0.918, 0.236);
  roundPanel(s, 1.048, 0.661, 5.281, 6.177, 2.64);
}

// 8. our founder
function slide08(s) {
  chrome(s);
  display(s, 'Our Founder', 1.55, 1.408, 4.419, 0.909, 54);
  kicker(s, 'People', 1.55, 1.13, 0.801, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus urna neque viverra justo nec. Eget felis eget nunc lobortis mattis. Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris.', 8.696, 1.838, 3.902, 2.02);
  subhead(s, 'About Company', 8.696, 1.408, 2.108, 0.269, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris.', 8.696, 4.067, 3.902, 0.505);
  body(s, 'You can\'t connect the dots looking forward; you can only connect them looking backward. So you have to trust that the dots will somehow connect in your future.', 8.696, 6.056, 3.902, 0.757, { color: '000000' });
  kicker(s, 'Steve Jobs', 8.696, 5.358, 1.237, 0.236, { color: NAVY });
  quoteMark(s, 8.696, 5.711, 0.325, NAVY);
  display(s, 'Amanda Stevie', 1.55, 5.694, 2.367, 0.471, 28, { color: NAVY });
  kicker(s, 'Designer', 1.55, 5.416, 1.02, 0.236, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisqu.', 1.55, 6.308, 2.367, 0.505);
  display(s, 'Mike Kabrone', 4.982, 5.694, 2.367, 0.471, 28, { color: NAVY });
  kicker(s, 'Designer', 4.982, 5.416, 1.02, 0.236, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisqu.', 4.982, 6.308, 2.367, 0.505);
  panel(s, 1.547, 2.622, 2.578, 2.488);
  panel(s, 4.981, 2.622, 2.578, 2.488);
}

// 9. our amazing people
function slide09(s) {
  chrome(s);
  display(s, 'Our Amazing People', 1.362, 1.409, 4.934, 0.808, 48);
  kicker(s, 'People', 1.362, 1.131, 0.95, 0.236);
  display(s, 'Kita Serinna', 1.477, 5.646, 2.446, 0.471, 28, { align: 'center' });
  kicker(s, 'Designer', 2.13, 5.388, 1.139, 0.236, { align: 'center' });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam.', 1.429, 6.26, 2.542, 0.757, { align: 'center' });
  display(s, 'David Powell', 5.672, 5.646, 2.446, 0.471, 28, { align: 'center' });
  kicker(s, 'Designer', 6.325, 5.388, 1.139, 0.236, { align: 'center' });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam.', 5.624, 6.26, 2.542, 0.757, { align: 'center' });
  display(s, 'Lilith Tiana', 9.868, 5.646, 2.446, 0.471, 28, { align: 'center' });
  kicker(s, 'Designer', 10.521, 5.388, 1.139, 0.236, { align: 'center' });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam.', 9.82, 6.26, 2.542, 0.757, { align: 'center' });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 8.621, 1.409, 3.815, 0.505);
  panel(s, 1.362, 2.488, 2.681, 2.641);
  panel(s, 5.556, 2.488, 2.681, 2.641);
  panel(s, 9.754, 2.488, 2.681, 2.641);
}

// 10. our vision
function slide10(s) {
  chrome(s);
  display(s, 'Our Vision', 8.672, 3.537, 3.624, 0.909, 54, { color: NAVY });
  kicker(s, 'About Us', 8.672, 3.249, 1.017, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida.', 8.672, 5.332, 4.01, 1.262);
  subhead(s, 'Your Awesome Title Here', 8.672, 4.903, 3.046, 0.269);
  panel(s, 0.652, 0.942, 7.141, 5.653);
}

// 11. our mission
function slide11(s) {
  panel(s, 0.652, 0, 6.763, 7.5);
  chrome(s);
  display(s, 'Our Mission', 8.48, 1.578, 3.605, 0.909, 54);
  kicker(s, 'About Us', 8.48, 1.3, 1.034, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum.', 8.48, 3.325, 4.01, 1.01);
  subhead(s, 'Mission One', 8.48, 2.896, 1.584, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum.', 8.48, 5.191, 4.01, 1.01);
  subhead(s, 'Mission Two', 8.48, 4.761, 1.518, 0.269);
}

// 12. branding
function slide12(s) {
  chrome(s);
  display(s, 'Branding', 1.677, 1.617, 3.251, 1.111, 66);
  kicker(s, 'Our Product', 1.677, 1.28, 1.429, 0.236);
  display(s, 'New Concept', 1.677, 2.49, 4.549, 1.111, 66);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus.', 7.271, 5.164, 5.396, 1.01);
  subhead(s, 'Brand Concept', 7.271, 4.741, 1.945, 0.269, { color: NAVY, font: ALT });
  panel(s, 0.652, 4.189, 5.726, 2.538);
  panel(s, 7.271, 0, 5.396, 3.75);
}

// 13. best style product
function slide13(s) {
  chrome(s);
  display(s, 'Best Style', 9.075, 2.058, 3.296, 0.909, 54);
  kicker(s, 'Our Product', 9.075, 1.722, 1.39, 0.236);
  display(s, 'Product', 9.075, 2.733, 2.868, 0.909, 54);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet.', 9.075, 4.769, 3.296, 1.01);
  display(s, '2021', 9.075, 4.108, 1.151, 0.606, 36);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus.', 1.127, 1.724, 2.958, 1.01);
  subhead(s, 'Style Product', 1.127, 1.305, 2.156, 0.269);
  panel(s, 1.127, 3.189, 3.296, 3.829);
  panel(s, 4.908, 0, 3.296, 7.5);
}

// 14. best service (three rows)
function slide14(s) {
  chrome(s);
  display(s, 'Our Best Service', 1.348, 1.29, 4.917, 0.909, 54);
  kicker(s, 'Our Service', 1.348, 0.953, 1.483, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.992, 2.92, 3.605, 0.757);
  subhead(s, 'Service One', 8.992, 2.496, 1.595, 0.269, { color: NAVY, font: ALT });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.996, 4.566, 3.605, 0.757);
  subhead(s, 'Service Two', 8.996, 4.142, 1.595, 0.269, { color: NAVY, font: ALT });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.992, 6.285, 3.605, 0.757);
  subhead(s, 'Service Three', 8.992, 5.862, 1.763, 0.269, { color: NAVY, font: ALT });
  icon(s, 'gear', 8.191, 2.453, 0.441, INK, WHITE);
  icon(s, 'bulb', 8.205, 4.136, 0.412, INK, WHITE);
  icon(s, 'photo', 8.205, 5.856, 0.412, INK, WHITE);
  panel(s, 0.652, 2.496, 6.839, 4.521);
}

// 15. best service (three columns)
function slide15(s) {
  chrome(s);
  display(s, 'Our Best Service', 1.348, 1.588, 4.917, 0.909, 54);
  kicker(s, 'Our Service', 1.348, 1.252, 1.483, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.348, 5.239, 2.837, 1.01);
  subhead(s, 'Service One', 1.348, 4.815, 1.595, 0.269);
  icon(s, 'gear', 1.348, 3.613, 1.04, INK, WHITE);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 5.588, 5.239, 2.732, 1.01);
  subhead(s, 'Service Two', 5.588, 4.815, 1.595, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 9.723, 5.239, 2.732, 1.01);
  subhead(s, 'Service Three', 9.723, 4.815, 1.763, 0.269);
  icon(s, 'bulb', 5.337, 3.647, 0.971, INK, WHITE);
  icon(s, 'photo', 9.723, 3.647, 0.971, INK, WHITE);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Nisl rhoncus mattis rhoncus.', 7.271, 1.487, 5.184, 1.01);
}

// 16. best service (two columns)
function slide16(s) {
  chrome(s);
  display(s, 'Our Best Service', 8.065, 4.779, 4.917, 0.909, 54);
  kicker(s, 'Our Service', 8.065, 4.443, 1.483, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.065, 5.789, 4.237, 0.757);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 5.375, 2.415, 2.561, 0.757);
  subhead(s, 'Service One', 5.375, 1.992, 1.595, 0.269);
  icon(s, 'gear', 5.34, 1.242, 0.587, INK, WHITE);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 8.855, 2.415, 2.561, 0.757);
  subhead(s, 'Service Two', 8.855, 1.992, 1.595, 0.269);
  icon(s, 'bulb', 8.762, 1.247, 0.549, INK, WHITE);
  panel(s, 0.652, 0.492, 3.919, 3.258);
  panel(s, 0.652, 3.981, 6.537, 3.027);
}

// 17. best services (icon list)
function slide17(s) {
  chrome(s);
  display(s, 'Our Best', 9.093, 1.377, 3.398, 1.01, 60, { color: NAVY });
  kicker(s, 'Our Service', 9.093, 1.04, 1.349, 0.236);
  display(s, 'Services', 9.086, 2.154, 3.253, 1.01, 60, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit.', 9.093, 4.253, 3.427, 1.262);
  subhead(s, 'Your Awesome Title Here', 9.093, 3.824, 3.032, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa.', 9.083, 5.702, 3.427, 0.757);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.93, 1.32, 2.929, 1.01);
  subhead(s, 'Service One', 1.93, 0.897, 1.595, 0.269);
  icon(s, 'gear', 1.129, 0.854, 0.441, INK, WHITE);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.934, 3.479, 2.929, 1.01);
  subhead(s, 'Service Two', 1.934, 3.056, 1.595, 0.269);
  icon(s, 'bulb', 1.143, 3.049, 0.412, INK, WHITE);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.93, 5.637, 2.929, 1.01);
  subhead(s, 'Service Three', 1.93, 5.213, 1.763, 0.269);
  icon(s, 'photo', 1.143, 5.208, 0.412, INK, WHITE);
  panel(s, 5.638, 0, 2.626, 7.5);
}

// 18. break slide
function slide18(s) {
  panel(s, 0.652, 0, 12.682, 7.5);
  chrome(s);
  display(s, 'Break Slide', 1.616, 1.325, 4.948, 1.346, 80, { color: WHITE });
  kicker(s, 'Breaktime', 1.616, 0.988, 1.233, 0.236, { color: WHITE });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc sceleque viverra mauris in aliquam sem fringilla. Lorem donec massa.', 1.616, 2.671, 4.453, 0.505, { color: WHITE });
}

// 19. our portfolio (banner)
function slide19(s) {
  chrome(s);
  display(s, 'Our Portfolio', 1.548, 4.703, 4.574, 1.01, 60);
  kicker(s, 'Our Portfolio', 1.548, 4.367, 1.387, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.549, 6.26, 4.01, 0.757);
  kicker(s, 'About Portfolio', 1.549, 5.915, 1.617, 0.236);
  quoteCard(s, 7.426, 4.601, 5.063, 1.938);
  body(s, 'You can\'t connect the dots looking forward; you can only connect them looking backward. So you have to trust that the dots will somehow connect in your future.', 8.006, 5.54, 3.902, 0.757, { color: '000000' });
  kicker(s, 'Steve Jobs', 8.006, 4.842, 1.237, 0.236);
  quoteMark(s, 8.006, 5.195, 0.182, INK);
  panel(s, 0.652, 0, 12.682, 3.75);
}

// 20. our portfolio (icons + quote)
function slide20(s) {
  panel(s, 2.459, 3.759, 10.874, 3.731);
  chrome(s);
  display(s, 'Our Portfolio', 8.452, 1.373, 4.574, 1.01, 60);
  kicker(s, 'Our Portfolio', 8.452, 1.036, 1.387, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.452, 2.484, 4.01, 0.757);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 1.292, 2.205, 2.561, 0.757);
  subhead(s, 'Service One', 1.292, 1.782, 1.595, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 4.772, 2.205, 2.561, 0.757);
  subhead(s, 'Service Two', 4.772, 1.782, 1.595, 0.269);
  icon(s, 'bulb', 1.222, 1.036, 0.549, INK, WHITE);
  icon(s, 'photo', 4.772, 1.071, 0.549, INK, WHITE);
  quoteCard(s, 0.642, 4.601, 5.063, 1.938);
  body(s, 'You can\'t connect the dots looking forward; you can only connect them looking backward. So you have to trust that the dots will somehow connect in your future.', 1.222, 5.54, 3.902, 0.757, { color: '000000' });
  kicker(s, 'Steve Jobs', 1.222, 4.842, 1.237, 0.236);
  quoteMark(s, 1.222, 5.195, 0.182, INK);
}

// 21. our portfolio (two blocks)
function slide21(s) {
  chrome(s);
  display(s, 'Our Portfolio', 8.452, 1.373, 4.574, 1.01, 60);
  kicker(s, 'Our Portfolio', 8.452, 1.036, 1.387, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida.', 1.357, 5.701, 5.25, 1.01);
  subhead(s, 'Portfolio Two', 1.357, 5.271, 1.611, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida.', 1.357, 1.466, 5.25, 1.01);
  subhead(s, 'Portfolio One', 1.357, 1.036, 1.694, 0.269);
  panel(s, 1.357, 2.811, 5.5, 2.104);
  panel(s, 8.452, 2.811, 4.881, 4.689);
}

// 22. our galery
function slide22(s) {
  chrome(s);
  display(s, 'Our Galery', 1.542, 5.38, 4.666, 1.346, 80);
  kicker(s, 'Our Galery', 1.542, 5.043, 1.207, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 1.542, 2.235, 4.364, 1.515);
  subhead(s, 'Best Of 2021', 1.542, 1.806, 2.213, 0.269);
  panel(s, 7.107, 0, 6.226, 4.557);
  panel(s, 7.107, 4.778, 6.226, 2.722);
}

// 23. best style product (two panels)
function slide23(s) {
  chrome(s);
  display(s, 'Best Style Product', 1.375, 1.397, 5.559, 0.909, 54);
  kicker(s, 'Our Product', 1.375, 1.061, 1.559, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerique viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet.', 1.375, 6.26, 4.568, 0.757);
  subhead(s, 'Best Product', 1.375, 5.83, 1.55, 0.269, { color: NAVY, font: ALT });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse.', 6.667, 6.26, 6.379, 0.757);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 8.352, 1.397, 3.804, 0.757);
  panel(s, 0.652, 2.619, 5.792, 2.843);
  panel(s, 6.667, 2.619, 6.22, 2.843);
}

// 24. your style
function slide24(s) {
  chrome(s);
  display(s, 'Your Style', 1.375, 1.56, 4.238, 1.111, 66);
  kicker(s, 'Our Product', 1.375, 1.224, 1.37, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida.', 7.774, 5.786, 4.917, 1.01);
  subhead(s, 'Modern Style', 7.774, 5.357, 1.635, 0.269);
  panel(s, 1.375, 2.93, 5.292, 4.57);
  panel(s, 7.802, 0.851, 5.085, 4.083);
}

// 25. mobile mockup
function slide25(s) {
  phoneMockup(s, 1.409, 0.482, 3.242, 6.556);
  chrome(s);
  display(s, 'Mockup Device', 7.247, 1.903, 5.017, 1.01, 60, { color: NAVY });
  kicker(s, 'Our Mockup', 7.247, 1.566, 1.536, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 7.247, 3.886, 5.146, 1.262);
  subhead(s, 'Mobile Mockup', 7.247, 3.456, 2.161, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 7.247, 5.429, 5.146, 0.505);
  phoneMockup(s, 3.722, 1.369, 2.364, 4.78);
}

// 26. desktop mockup
function slide26(s) {
  monitorMockup(s, 7.723, 1.96, 4.868, 4.07);
  chrome(s);
  display(s, 'Mockup Device', 1.39, 1.903, 5.017, 1.01, 60, { color: NAVY });
  kicker(s, 'Our Mockup', 1.39, 1.566, 1.536, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum. Praesent tristique magna sit amet purus gravida. Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla.', 1.39, 3.886, 5.146, 1.262);
  subhead(s, 'Desktop Mockup', 1.39, 3.456, 2.161, 0.269);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et.', 1.39, 5.429, 5.146, 0.505);
}

// 27. column chart
function slide27(s) {
  chrome(s);
  barChart(s);
  display(s, 'Column Chart', 1.482, 5.495, 4.692, 1.01, 60);
  kicker(s, 'Our Chart', 1.482, 5.158, 1.4, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum.', 8.021, 5.276, 4.125, 1.01);
}

// 28. infographic
function slide28(s) {
  s.addShape('teardrop', { x: 4.338, y: 3.095, w: 1.725, h: 1.725, rotate: 45, fill: { color: INK } });
  s.addShape('teardrop', { x: 7.327, y: 3.095, w: 1.725, h: 1.725, rotate: 45, fill: { color: MID } });
  s.addShape('teardrop', { x: 10.323, y: 3.095, w: 1.725, h: 1.725, rotate: 45, fill: { color: INK } });
  chrome(s);
  display(s, 'Infographic', 1.371, 1.352, 4.713, 1.212, 72);
  kicker(s, 'Our Infographic', 1.371, 1.016, 1.725, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque laoreet suspendisse interdum.', 7.767, 1.254, 4.125, 1.01);
  s.addShape('teardrop', { x: 1.371, y: 3.095, w: 1.725, h: 1.725, rotate: 45, fill: { color: MID } });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem.', 1.371, 5.727, 2.375, 0.757);
  subhead(s, 'About One', 1.371, 5.297, 1.641, 0.269, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem.', 4.352, 5.727, 2.375, 0.757);
  subhead(s, 'About Two', 4.352, 5.297, 1.641, 0.269, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem.', 7.334, 5.727, 2.375, 0.757);
  subhead(s, 'About Three', 7.334, 5.297, 1.792, 0.269, { color: NAVY });
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem.', 10.316, 5.727, 2.375, 0.757);
  subhead(s, 'About Four', 10.316, 5.297, 1.641, 0.269, { color: NAVY });
  icon(s, 'target', 7.818, 3.578, 0.757, WHITE, MID);
  icon(s, 'gear', 1.855, 3.578, 0.757, WHITE, MID);
  icon(s, 'bulb', 4.836, 3.578, 0.757, WHITE, INK);
  icon(s, 'gem', 10.8, 3.578, 0.757, WHITE, INK);
}

// 29. contact us
function slide29(s) {
  chrome(s);
  display(s, 'Contact Us', 7.434, 1.439, 4.121, 1.111, 66, { color: NAVY });
  kicker(s, 'Visit Us', 7.434, 1.103, 0.836, 0.236);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta non pulvinar neque.', 7.434, 2.651, 4.594, 0.757);
  subhead(s, 'Office Hours', 7.434, 4.013, 1.487, 0.269);
  small(s, 'Monday - Saturday', 7.434, 4.43, 1.626, 0.202, { color: '000000' });
  small(s, '08.00 – 16.00', 7.434, 4.724, 1.125, 0.202, { color: '000000' });
  subhead(s, 'Get In Touch', 10.547, 4.013, 1.498, 0.269);
  small(s, '+214-704-76532', 10.547, 4.43, 1.335, 0.202, { color: '000000' });
  small(s, '+214-704-76533', 10.547, 4.724, 1.335, 0.202, { color: '000000' });
  subhead(s, 'Follow Us', 7.434, 5.484, 1.16, 0.269);
  small(s, 'www.companyinfo.com', 7.434, 5.901, 1.84, 0.202, { color: '000000' });
  small(s, 'company@info.com', 7.434, 6.195, 1.626, 0.202, { color: '000000' });
  subhead(s, 'Our Address', 10.547, 5.484, 1.498, 0.269);
  s.addText([{ text: 'Koyambed – 1' }, { text: 'st', options: { superscript: true } }, { text: '  Street 019' }], { x: 10.547, y: 5.901, w: 2.144, h: 0.202, fontFace: SANS, fontSize: 12, color: '000000', valign: 'top', margin: 0 });
  small(s, 'Agassiz City - 4784', 10.547, 6.195, 1.564, 0.202, { color: '000000' });
  panel(s, 0.969, 0.483, 5.697, 6.535);
}

// 30. thank you
function slide30(s) {
  panel(s, 0.652, 0, 12.682, 7.5);
  chrome(s);
  body(s, 'Sit amet consectetur adipiscing elit duis. Cursus eget nunc scelerisque viverra mauris in aliquam sem fringilla. Lorem donec massa sapien faucibus et. Vel orci porta.', 9.382, 5.115, 3.094, 1.01, { color: WHITE, align: 'right' });
  display(s, 'Thank', 7.474, 0.677, 5.002, 2.794, 166, { color: WHITE, align: 'right' });
  display(s, 'You', 7.686, 2.619, 4.789, 2.794, 166, { color: WHITE, align: 'right' });
}

/* ------------------------------------------------------------------- build */

const BUILDERS = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
  slide27,
  slide28,
  slide29,
  slide30
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.3333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.title = 'ESSENTIAL Presentation - Business Plan';

BUILDERS.forEach(function (build) {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '17a8c2ed-540e-4449-85f7-fefd9ee85b1b_grok_final.pptx') })
  .then(function (name) { console.log('wrote ' + name); })
  .catch(function (err) { console.error(err); process.exit(1); });
