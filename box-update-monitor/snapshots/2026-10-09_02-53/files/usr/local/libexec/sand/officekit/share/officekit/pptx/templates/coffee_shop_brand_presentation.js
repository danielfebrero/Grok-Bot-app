/*
 * "Cupfe - Coffee" deck, rebuilt with pptxgenjs.
 * Raster photos in the original are replaced by flat placeholder rectangles.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  brown: '6C503A',        // accent1
  green: '576E65',        // accent2
  tan: 'D19566',          // accent3
  cream: 'F5EFE0',        // accent4 / slide background
  creamDeep: 'EADDBD',    // accent4 lumMod 90%
  white: 'FFFFFF',
  ink: '111111',          // tx1
  body: '353535',         // tx1 lumMod 85%
  bodyMid: '4C4C4C',      // tx1 lumMod 75%
  navMuted: 'BFBFBF',     // bg1 lumMod 75%
  brownDark: '36281D',
  brownMid: '513C2C',
  brownLight: 'B79479',
  brownPale: 'CFB8A5',
  ph: 'CCCCCC',           // image placeholder body
  phText: '9B9B9B'
};

const F = { sans: 'Work Sans', display: 'Barlow Condensed SemiBold', script: 'Sacramento' };

// pptxgenjs rewrites shadow option objects in place, so hand out a fresh one every time
function shadow(angle) {
  return { type: 'outer', blur: 50, offset: 20, angle: angle === undefined ? 45 : angle, color: '000000', opacity: 0.2 };
}

// roundRect "adj" values are a fraction of the shorter side
const R = (frac, w, h) => +(frac * Math.min(w, h)).toFixed(3);

// ------------------------------------------------------------ text bodies ---
const T = {
  lorem: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud',
  loremShort: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et',
  loremCard: 'Lorem ipsum dolor sit consectetu',
  loremMenu: 'Lorem ipsum dolor sit consectetu ipsum',
  loremNew: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore labore et dolore magna aliqua. Ut enim',
  bullet: 'Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor congue. ipsum dolor sit amet',
  bulletLong: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris',
  swotW: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. porttitor congue massa congue massa. porttitor congue massa congue massa. porttitor congue massa. ',
  swotT: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. porttitor congue massa congue massa. congue massa congue massa. ',
  swotMini: 'Lorem ipsum dolor sit amet, adipiscing elit, sed do ipsum',
  service: 'Lorem ipsum dolor sit conse',
  price: 'Lorem ipsum dolor sit amet, adipiscing elit porttitor adipi scing elit porttitor',
  arrowStep: 'Lorem ipsum dolor sit amet, conse ctetuer apiscing',
  ringItem: 'Lorem ipsum dolor sit amet, consectetuer adipiscing',
  wideIntro: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere porttitor congue massa. ',
  bigPct: 'Lorem ipsum dolor sit amet, conse ctetuer adipiscing elit. Maecenas porttitor congue massa ipsum dolor sit amet ipsum dolor',
  priceRow: 'Lorem ipsum dolor sit amet, consect etuer adipiscing elit. dolor sit etuer',
  circleItem: 'Lorem ipsum dolor sit amet, elit porttitor congue',
  timeline: 'Lorem ipsum dolor amet, consectetuer massa. Fusce posuere, magna',
  yearCard: 'Lorem ipsum dolor sit amet, consect ip sum dolor ',
  learn: 'Lorem ipsum dolor sitam et, consec tetuer',
  donutItem: 'Lorem ipsum dolor ame consect adipiscing',
  cycleItem: 'Lorem ipsum dolor sit amet, consect etuer adipiscing elit. etuer adipiscing etu er adipiscing',
  cardItem: 'Lorem ipsum dolor sit amet, adip dolor sit',
  checkItem: 'Lorem ipsum dolor sit amet, consect etuer adipiscing elit. etuer adipiscing etuer',
  stepItem: 'Lorem ipsum dolor sit amet, consect',
  contact: 'Lorem ipsum dolor amet ipsum amet, consectetuer adip iscing elit. Maecenas porttitor congue. ',
  thanks: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore'
};

// ---------------------------------------------------------------- helpers ---
// paragraph of running body copy (10.5pt Work Sans, 150% leading)
function body(s, text, x, y, w, h, o) {
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: h, fontFace: F.sans, fontSize: 10.5, color: C.body,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top'
  }, o));
}

// small bold label (card / section titles)
function label(s, text, x, y, w, o) {
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: C.body,
    align: 'justify', valign: 'top'
  }, o));
}

function shape(s, kind, o) { s.addShape(kind, o); }

function textShape(s, kind, text, o) {
  s.addText(text, Object.assign({ shape: kind, fontFace: F.sans, align: 'center', valign: 'middle' }, o));
}

// grey block standing in for a photo in the source deck.
// `flat` squares off one side, matching the half-rounded crops used in the deck.
function imagePlaceholder(s, x, y, w, h, o) {
  o = o || {};
  const opt = { x: x, y: y, w: w, h: h, fill: { color: C.ph }, line: { color: C.ph, width: 0.5 } };
  if (o.rectRadius) opt.rectRadius = o.rectRadius;
  if (o.rotate) opt.rotate = o.rotate;
  if (o.flipH) opt.flipH = true;
  if (o.flipV) opt.flipV = true;
  s.addShape(o.shape || 'roundRect', opt);
  const r = o.rectRadius || 0;
  if (o.flat === 'left') shape(s, 'rect', { x: x, y: y, w: r, h: h, fill: { color: C.ph } });
  if (o.flat === 'right') shape(s, 'rect', { x: x + w - r, y: y, w: r, h: h, fill: { color: C.ph } });
  if (w > 1.1 && h > 0.5) {
    s.addText('[image]', {
      x: x, y: y + h / 2 - 0.18, w: w, h: 0.36, rotate: o.rotate,
      fontFace: F.sans, fontSize: 11, color: C.phText, align: 'center', valign: 'middle'
    });
  }
}

// site-style chrome repeated on nearly every slide: logo, menu, corner tabs
function nav(s, o) {
  o = o || {};
  const menu = o.menuColor || C.navMuted;
  const brand = o.brandColor || C.body;
  const mark = o.markColor || C.tan;
  const tab = o.tabColor || C.brown;
  const items = [
    ['Home', 9.122, 0.565, menu],
    ['About', 9.887, 0.570, o.menuColor || C.brown],
    ['Service', 10.648, 0.649, menu],
    ['Contact', 11.582, 0.689, menu]
  ];
  items.forEach(function (it) {
    s.addText(it[0], {
      x: it[1], y: 0.372, w: it[2], h: 0.252, fontFace: F.sans, fontSize: 9,
      color: it[3], align: 'right', valign: 'top'
    });
  });
  s.addText('CUPFE', {
    x: 0.724, y: 0.355, w: 0.714, h: 0.286, fontFace: F.sans, fontSize: 11, bold: true,
    color: brand, valign: 'top'
  });
  // cup-and-steam glyph left of the wordmark
  shape(s, 'roundRect', { x: 0.545, y: 0.43, w: 0.15, h: 0.10, fill: { color: mark }, rectRadius: 0.025 });
  shape(s, 'rect', { x: 0.679, y: 0.452, w: 0.045, h: 0.022, fill: { color: mark } });
  shape(s, 'rect', { x: 0.541, y: 0.548, w: 0.19, h: 0.024, fill: { color: mark } });
  shape(s, 'rect', { x: 0.582, y: 0.372, w: 0.02, h: 0.045, fill: { color: mark } });
  shape(s, 'rect', { x: 0.632, y: 0.372, w: 0.02, h: 0.045, fill: { color: mark } });
  // top-right notch + right-edge tabs
  shape(s, 'round1Rect', { x: 12.729, y: 0, w: 0.604, h: 0.604, fill: { color: tab }, flipH: true, flipV: true });
  shape(s, 'round1Rect', { x: 13.236, y: 6.118, w: 0.097, h: 0.391, fill: { color: C.green }, rectRadius: R(0.5, 0.097, 0.391), flipH: true });
  shape(s, 'round1Rect', { x: 13.236, y: 6.509, w: 0.097, h: 0.391, fill: { color: C.tan }, rectRadius: R(0.5, 0.097, 0.391), flipH: true, flipV: true });
}

// "Delicious Coffee" script kicker + big display headline
function heading(s, text, x, y, w, o) {
  o = o || {};
  const align = o.align || 'left';
  s.addText('Delicious Coffee', {
    x: o.subX !== undefined ? o.subX : x, y: y - 0.219, w: o.subW || 2.31, h: 0.438,
    fontFace: F.script, fontSize: 20, bold: true, color: C.tan, align: align,
    valign: 'top'
  });
  s.addText(text, {
    x: x, y: y, w: w, h: 0.841, fontFace: F.display, fontSize: 44, bold: true, color: C.brown,
    align: align, valign: 'top'
  });
}

// centred headline used by every "Creative Infographic" style slide
function headingCenter(s, text) {
  heading(s, text, 3.354, 1.427, 6.332, { align: 'center', subX: 5.319, subW: 2.403 });
}

// number / title / paragraph trio that sits inside the coloured cards
function numberBlock(s, x, numY, num, title, desc, o) {
  o = o || {};
  s.addText(num, {
    x: x, y: numY, w: o.numW || 1.049, h: 0.64, fontFace: F.sans, fontSize: o.numSize || 32,
    bold: true, color: o.color || C.white, align: 'justify', valign: 'top'
  });
  label(s, title, o.textX !== undefined ? o.textX : x, numY + 0.574, o.titleW || 1.612, { color: o.color || C.white });
  body(s, desc, o.textX !== undefined ? o.textX : x, numY + 0.813, o.descW || 1.612, 0.601, { color: o.color || C.white });
}

// ------------------------------------------------------------ icon glyphs ---
function iconTag(s, x, y, w, h, color) {          // price tag (slide 2)
  shape(s, 'homePlate', { x: x, y: y, w: w, h: h, fill: { color: color } });
  shape(s, 'ellipse', { x: x + w * 0.14, y: y + h * 0.24, w: h * 0.22, h: h * 0.22, fill: { color: C.green } });
}
function iconThumb(s, x, y, w, h, color) {        // thumbs-up (slides 6, 12)
  shape(s, 'roundRect', { x: x + w * 0.30, y: y + h * 0.28, w: w * 0.70, h: h * 0.72, fill: { color: color }, rectRadius: 0.05 });
  shape(s, 'roundRect', { x: x, y: y + h * 0.42, w: w * 0.26, h: h * 0.58, fill: { color: color }, rectRadius: 0.03 });
  shape(s, 'roundRect', { x: x + w * 0.42, y: y, w: w * 0.20, h: h * 0.45, fill: { color: color }, rectRadius: 0.05 });
}
function iconHeart(s, x, y, w, h, color) { shape(s, 'heart', { x: x, y: y, w: w, h: h, fill: { color: color } }); }
function iconTrophy(s, x, y, w, h, color) {
  shape(s, 'trapezoid', { x: x + w * 0.12, y: y, w: w * 0.76, h: h * 0.55, fill: { color: color }, rotate: 180 });
  shape(s, 'rect', { x: x + w * 0.42, y: y + h * 0.5, w: w * 0.16, h: h * 0.28, fill: { color: color } });
  shape(s, 'roundRect', { x: x + w * 0.18, y: y + h * 0.78, w: w * 0.64, h: h * 0.22, fill: { color: color }, rectRadius: 0.03 });
}
function iconPerson(s, x, y, w, h, color) {
  shape(s, 'ellipse', { x: x + w * 0.28, y: y, w: w * 0.44, h: h * 0.42, fill: { color: color } });
  shape(s, 'ellipse', { x: x, y: y + h * 0.45, w: w, h: h * 0.9, fill: { color: color } });
}
function iconPhoto(s, x, y, w, h, color) {
  shape(s, 'roundRect', { x: x, y: y, w: w, h: h, fill: { color: color }, rectRadius: 0.04 });
  shape(s, 'ellipse', { x: x + w * 0.18, y: y + h * 0.18, w: w * 0.22, h: h * 0.19, fill: { color: C.green } });
  shape(s, 'triangle', { x: x + w * 0.22, y: y + h * 0.45, w: w * 0.6, h: h * 0.4, fill: { color: C.green } });
}
function iconCheck(s, x, y, size, color) {
  shape(s, 'ellipse', { x: x, y: y, w: size, h: size, fill: { color: color } });
  shape(s, 'rect', { x: x + size * 0.24, y: y + size * 0.47, w: size * 0.26, h: size * 0.10, fill: { color: C.white }, rotate: 45 });
  shape(s, 'rect', { x: x + size * 0.40, y: y + size * 0.42, w: size * 0.42, h: size * 0.10, fill: { color: C.white }, rotate: -45 });
}
function iconCloud(s, x, y, w, h, color) {        // cloud-with-download-arrow (slide 25)
  shape(s, 'ellipse', { x: x, y: y + h * 0.30, w: w * 0.42, h: h * 0.46, fill: { color: color } });
  shape(s, 'ellipse', { x: x + w * 0.56, y: y + h * 0.30, w: w * 0.44, h: h * 0.46, fill: { color: color } });
  shape(s, 'ellipse', { x: x + w * 0.24, y: y + h * 0.06, w: w * 0.52, h: h * 0.56, fill: { color: color } });
  shape(s, 'rect', { x: x + w * 0.20, y: y + h * 0.50, w: w * 0.60, h: h * 0.26, fill: { color: color } });
  shape(s, 'downArrow', { x: x + w * 0.36, y: y + h * 0.34, w: w * 0.28, h: h * 0.52, fill: { color: C.cream } });
}
function iconGlobe(s, x, y, size, color) {
  shape(s, 'ellipse', { x: x, y: y, w: size, h: size, fill: { color: color } });
  shape(s, 'ellipse', { x: x + size * 0.30, y: y, w: size * 0.40, h: size, fill: { color: C.cream }, line: { color: color, width: 0.75 } });
  shape(s, 'rect', { x: x, y: y + size * 0.44, w: size, h: size * 0.1, fill: { color: C.cream } });
}
function iconHome(s, x, y, w, h, color) {
  shape(s, 'triangle', { x: x, y: y, w: w, h: h * 0.55, fill: { color: color } });
  shape(s, 'rect', { x: x + w * 0.16, y: y + h * 0.5, w: w * 0.68, h: h * 0.5, fill: { color: color } });
}
function iconMail(s, x, y, w, h, color) {
  shape(s, 'rect', { x: x, y: y, w: w, h: h, fill: { color: color } });
  shape(s, 'triangle', { x: x + w * 0.1, y: y + h * 0.15, w: w * 0.8, h: h * 0.6, fill: { color: C.cream }, rotate: 180 });
}
function iconPhone(s, x, y, w, h, color) {
  shape(s, 'ellipse', { x: x, y: y, w: w, h: h, fill: { color: color } });
  shape(s, 'roundRect', { x: x + w * 0.28, y: y + h * 0.22, w: w * 0.44, h: h * 0.56, fill: { color: C.cream }, rectRadius: 0.03, rotate: 25 });
}

// C-shaped band used by the snake diagram on slide 15
function band(s, x, y, w, h, color, openRight) {
  const t = h * 0.3;
  shape(s, 'roundRect', { x: x, y: y, w: w, h: h, fill: { color: color }, rectRadius: h / 2 });
  shape(s, 'roundRect', {
    x: openRight ? x + t : x - t, y: y + t, w: w, h: h - 2 * t,
    fill: { color: C.cream }, rectRadius: (h - 2 * t) / 2
  });
}

// triangular arrow head centred on (cx, cy); `len` runs along the pointing
// direction, `span` across it. deg 0 = up, 90 = right, 270 = left.
function arrowHead(s, cx, cy, len, span, deg, color) {
  s.addShape('triangle', {
    x: cx - span / 2, y: cy - len / 2, w: span, h: len,
    fill: { color: color }, rotate: deg
  });
}

// ================================================================== slides ==
const slides = [];

// 1 — cover
slides.push(function (s) {
  imagePlaceholder(s, 0, 0, 13.333, 7.5, { shape: 'rect' });
  shape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.cream, transparency: 7 } });
  imagePlaceholder(s, 7.773, 1.433, 4.669, 4.669, { shape: 'rect' });
  nav(s);
  s.addText('CUPFE', {
    x: 0.968, y: 1.571, w: 6.806, h: 3.635, fontFace: F.display, fontSize: 210, bold: true,
    color: C.brown, valign: 'top', wrap: false
  });
  s.addText('Coffee', {
    x: 1.032, y: 3.789, w: 5.811, h: 2.457, fontFace: F.script, fontSize: 140, bold: true,
    color: C.tan, valign: 'top'
  });
  [[7.911, 1.175, 0.321, C.tan], [12.034, 6.04, 0.321, C.tan],
   [11.881, 1.352, 0.502, C.green], [7.821, 5.688, 0.502, C.green]].forEach(function (d) {
    shape(s, 'donut', { x: d[0], y: d[1], w: d[2], h: d[2], fill: { color: d[3] } });
  });
});

// 2 — welcome
slides.push(function (s) {
  imagePlaceholder(s, 0, 1.417, 5.875, 4.688, { rectRadius: 0.365, flat: 'left' });
  nav(s);
  heading(s, 'Welcome To Cupfe', 7.463, 1.815, 4.623, { subW: 2.423 });
  body(s, T.lorem, 7.463, 3.181, 4.623, 0.896);
  label(s, 'Description About Us', 7.463, 4.603, 2.215, { color: C.green, align: 'left' });
  body(s, T.lorem, 7.463, 5.046, 4.623, 0.896);
  shape(s, 'roundRect', { x: 4.286, y: 4.628, w: 2.383, h: 2.383, fill: { color: C.green }, rectRadius: R(0.078, 2.383, 2.383), shadow: shadow() });
  iconTag(s, 5.256, 5.259, 0.444, 0.348, C.white);
  s.addText('$150', { x: 4.808, y: 5.621, w: 1.339, h: 0.572, fontFace: F.sans, fontSize: 28, bold: true, color: C.white, align: 'center', valign: 'top' });
  s.addText('Starting Price', { x: 4.808, y: 6.102, w: 1.339, h: 0.278, fontFace: F.sans, fontSize: 10, color: C.white, align: 'center', valign: 'top' });
});

// 3 — about vision (three cards)
slides.push(function (s) {
  imagePlaceholder(s, 7.479, 1.417, 5.854, 6.083, { shape: 'round1Rect', rectRadius: 0.975, flipH: true });
  nav(s);
  heading(s, 'About Vision', 1.19, 1.621, 4.028);
  body(s, T.loremShort, 1.19, 2.613, 4.704, 0.601);
  [[1.303, C.brown, '01.'], [3.907, C.green, '02.'], [6.511, C.tan, '03.']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: 3.75, w: 2.362, h: 2.362, fill: { color: d[1] }, rectRadius: R(0.078, 2.362, 2.362), shadow: shadow() });
    numberBlock(s, d[0] + 0.375, 4.224, d[2], 'Vision Tittle', T.loremCard);
  });
});

// 4 — about mission (four cards)
slides.push(function (s) {
  imagePlaceholder(s, 7.081, 4.224, 4.966, 1.888, { rectRadius: 0.147 });
  nav(s);
  heading(s, 'About Mission', 1.19, 1.621, 4.27, { subW: 2.268 });
  body(s, T.loremShort, 1.19, 2.613, 4.704, 0.601);
  [[1.303, 3.75, C.brown, '01.'], [3.907, 3.75, C.green, '02.'],
   [7.081, 1.418, C.brown, '03.'], [9.686, 1.418, C.tan, '04.']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: d[1], w: 2.362, h: 2.362, fill: { color: d[2] }, rectRadius: R(0.078, 2.362, 2.362), shadow: shadow() });
    numberBlock(s, d[0] + 0.375, d[1] + 0.474, d[3], 'Mission Tittle', T.loremCard);
  });
});

// 5 — coffee menu
slides.push(function (s) {
  // the three crops are cut on their bottom corners: right only, both, left only
  imagePlaceholder(s, 0, 0, 4.331, 2.683, { shape: 'round1Rect', flipV: true });
  imagePlaceholder(s, 4.501, 0, 4.331, 2.683, { shape: 'round2SameRect', flipV: true });
  imagePlaceholder(s, 9.003, 0, 4.331, 2.683, { shape: 'round1Rect', flipH: true, flipV: true });
  nav(s, { menuColor: C.white, brandColor: C.white, markColor: C.white, tabColor: C.white });
  heading(s, 'Variants Coffee Menu', 3.354, 3.524, 6.332, { align: 'center', subX: 5.319, subW: 2.403 });
  [[1.177, C.brown, '01.'], [4.961, C.green, '02.'], [8.745, C.tan, '03.']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: 4.742, w: 3.432, h: 1.37, fill: { color: d[1] }, rectRadius: R(0.078, 3.432, 1.37), shadow: shadow() });
    s.addText(d[2], { x: d[0] + 0.316, y: 4.956, w: 0.911, h: 0.64, fontFace: F.sans, fontSize: 32, bold: true, color: C.white, align: 'justify', valign: 'top' });
    label(s, 'Coffee Tittle', d[0] + 1.279, 5.028, 1.612, { color: C.white });
    body(s, T.loremMenu, d[0] + 1.279, 5.267, 1.796, 0.601, { color: C.white });
  });
});

// 6 — new coffee
slides.push(function (s) {
  // photo area is two panels: wide one rounded at the top, narrow one rounded at the bottom
  imagePlaceholder(s, 1.167, 1.417, 4.708, 6.083, { shape: 'round2SameRect', rectRadius: 0.451 });
  imagePlaceholder(s, 6.074, 0, 1.737, 6.083, { shape: 'round2SameRect', rectRadius: 0.427, flipV: true });
  nav(s);
  shape(s, 'round2SameRect', { x: 0, y: 3.742, w: 2.362, h: 2.362, fill: { color: C.green }, rotate: 90, rectRadius: R(0.123, 2.362, 2.362), shadow: shadow() });
  shape(s, 'roundRect', { x: 5.288, y: 2.389, w: 1.37, h: 1.37, fill: { color: C.tan }, rectRadius: R(0.108, 1.37, 1.37), shadow: shadow() });
  heading(s, 'New Coffee', 9.047, 1.944, 3.185, { subW: 2.423 });
  iconThumb(s, 1.006, 4.267, 0.351, 0.338, C.white);
  label(s, 'Best Service', 0.375, 4.709, 1.612, { color: C.white, align: 'center' });
  body(s, T.loremCard, 0.375, 4.948, 1.612, 0.601, { color: C.white, align: 'center' });
  s.addText('10+', { x: 5.448, y: 2.675, w: 1.051, h: 0.572, fontFace: F.sans, fontSize: 28, bold: true, color: C.white, align: 'center', valign: 'top' });
  s.addText('New Coffee', { x: 5.448, y: 3.157, w: 1.051, h: 0.252, fontFace: F.sans, fontSize: 9, color: C.white, align: 'center', valign: 'top' });
  label(s, 'About Coffee', 9.047, 3.263, 2.215, { color: C.green, align: 'left' });
  body(s, T.loremNew, 9.047, 3.707, 3.185, 1.161);
  textShape(s, 'roundRect', 'Learn More', {
    x: 9.122, y: 5.189, w: 1.685, h: 0.552, fill: { color: C.tan }, rectRadius: R(0.108, 1.685, 0.552),
    shadow: shadow(), fontSize: 12, bold: true, color: C.white
  });
});

// 7 — barista team
slides.push(function (s) {
  [[1.16, 2.196, 2.186], [3.875, 2.529, 2.518], [6.923, 2.196, 2.186], [9.637, 2.529, 2.518]].forEach(function (d) {
    imagePlaceholder(s, d[0], 2.616, d[1], d[2], { rectRadius: 0.196 });
  });
  nav(s);
  [['Ndemi Otieno', 3.875, 5.271, 2.529, 0.847, C.tan], ['Olivia Wilson', 9.637, 5.271, 2.529, 0.847, C.tan],
   ['Lars Peeters', 1.16, 4.948, 2.196, 0.748, C.green], ['Avery Davis', 6.923, 4.948, 2.196, 0.748, C.green]].forEach(function (d) {
    textShape(s, 'roundRect', d[0], {
      x: d[1], y: d[2], w: d[3], h: d[4], fill: { color: d[5] }, rectRadius: R(0.2, d[3], d[4]),
      shadow: shadow(), fontSize: 14, bold: true, color: C.white
    });
  });
  headingCenter(s, 'Meet Our Barista');
});

// 8-11 — SWOT family. Each has one photo, one big letter tile and a body area.
function swotBase(s, opts) {
  imagePlaceholder(s, opts.imgX, 1.417, 3.896, 4.688, { rectRadius: 0.303 });
  nav(s);
  textShape(s, 'roundRect', opts.letter, {
    x: opts.tileX, y: 2.885, w: 1.75, h: 1.75, fill: { color: opts.tileColor },
    rectRadius: R(0.108, 1.75, 1.75), shadow: shadow(), fontSize: 72, bold: true, color: C.white
  });
  heading(s, 'Our Opportunities', opts.textX, 1.621, 5.004);
}

function bulletList(s, x, ys, text, h) {
  ys.forEach(function (y) {
    shape(s, 'ellipse', { x: x, y: y, w: 0.156, h: 0.156, fill: { color: C.brown } });
    body(s, text, x + 0.315, y - 0.139, 4.287, h, { color: C.bodyMid });
  });
}

slides.push(function (s) {                                        // 8 — S
  swotBase(s, { imgX: 8.292, tileX: 7.413, tileColor: C.brown, letter: 'S', textX: 1.19 });
  bulletList(s, 1.317, [2.928, 3.843, 4.757, 5.672], T.bullet, 0.601);
});

slides.push(function (s) {                                        // 9 — W
  swotBase(s, { imgX: 1.167, tileX: 4.184, tileColor: C.green, letter: 'W', textX: 7.351 });
  body(s, T.swotW, 7.351, 2.934, 4.742, 1.161, { color: C.bodyMid });
  [[7.351, '01.', C.green], [10.225, '02.', C.tan]].forEach(function (d) {
    s.addText(d[1], { x: d[0], y: 4.589, w: 0.875, h: 0.572, fontFace: F.sans, fontSize: 28, bold: true, color: d[2], align: 'justify', valign: 'top' });
    body(s, T.swotMini, d[0], 5.253, 1.804, 0.896, { color: C.bodyMid });
  });
});

slides.push(function (s) {                                        // 10 — O
  swotBase(s, { imgX: 8.292, tileX: 7.413, tileColor: C.green, letter: 'O', textX: 1.19 });
  bulletList(s, 1.317, [3.204, 4.968], T.bulletLong, 1.161);
});

slides.push(function (s) {                                        // 11 — T (bar chart)
  swotBase(s, { imgX: 1.167, tileX: 4.184, tileColor: C.brown, letter: 'T', textX: 7.351 });
  [[3.078, '2021', '64%', 1.274, C.brown], [3.739, '2022', '89%', 2.081, C.green],
   [4.401, '2023', '78%', 1.788, C.tan]].forEach(function (d) {
    s.addText(d[1], { x: 7.369, y: d[0], w: 0.819, h: 0.286, fontFace: F.sans, fontSize: 11, color: C.bodyMid, align: 'justify', valign: 'top' });
    shape(s, 'roundRect', { x: 8.081, y: d[0] + 0.045, w: 2.494, h: 0.195, fill: { color: C.creamDeep }, rectRadius: 0.0975 });
    shape(s, 'roundRect', { x: 8.08, y: d[0] + 0.045, w: d[3], h: 0.195, fill: { color: d[4] }, rectRadius: 0.0975 });
    s.addText(d[2], { x: 10.679, y: d[0], w: 0.683, h: 0.286, fontFace: F.sans, fontSize: 11, color: C.bodyMid, align: 'justify', valign: 'top' });
  });
  body(s, T.swotT, 7.369, 5.141, 4.903, 0.896, { color: C.bodyMid });
});

// 12 — service grid
slides.push(function (s) {
  [[1.958, 1.417], [4.563, 3.87], [7.168, 1.417], [9.774, 3.87]].forEach(function (p) {
    imagePlaceholder(s, p[0], p[1], 2.414, 2.242, { rectRadius: R(0.078, 2.414, 2.242) });
  });
  nav(s);
  [[1.958, 3.87, C.brown, iconPerson, 3.013, 4.35, 0.305],
   [4.563, 1.417, C.green, iconHeart, 5.574, 1.887, 0.394],
   [7.168, 3.87, C.tan, iconPhoto, 8.233, 4.35, 0.286],
   [9.774, 1.417, C.brown, iconTrophy, 10.798, 1.889, 0.366]].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: d[1], w: 2.414, h: 2.242, fill: { color: d[2] }, rectRadius: R(0.078, 2.414, 2.242), shadow: shadow() });
    d[3](s, d[4], d[5], d[6], 0.338, C.white);
    label(s, 'Service Tittle', d[0] + 0.401, d[1] + 0.976, 1.612, { color: C.white, align: 'center' });
    body(s, T.service, d[0] + 0.401, d[1] + 1.215, 1.612, 0.601, { color: C.white, align: 'center' });
  });
  s.addText('SERVICE', {
    x: -1.13, y: 2.959, w: 4.315, h: 1.582, rotate: 270, fontFace: F.display, fontSize: 88,
    bold: true, color: C.creamDeep, align: 'center', valign: 'top'
  });
});

// 13 — portfolio
slides.push(function (s) {
  [1.938, 5.417, 8.896].forEach(function (x) {
    imagePlaceholder(s, x, 1.417, 3.292, 4.688, { rectRadius: 0.256 });
  });
  nav(s);
  s.addText('PORTOFOLIO', {
    x: -1.826, y: 3.144, w: 5.708, h: 1.212, rotate: 270, fontFace: F.display, fontSize: 66,
    bold: true, color: C.creamDeep, align: 'center', valign: 'top'
  });
});

// 14 — pricing table
slides.push(function (s) {
  nav(s);
  [[1.582, C.brown, '$2.000 '], [5.162, C.green, '$5.000 '], [8.741, C.tan, '$3.000 ']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: 2.67, w: 3.01, h: 3.378, fill: { color: d[1] }, rectRadius: R(0.082, 3.01, 3.378), shadow: shadow(114) });
    shape(s, 'rect', { x: d[0], y: 3.085, w: 3.01, h: 0.692, fill: { color: C.white, transparency: 90 } });
    s.addText([
      { text: d[2], options: { fontSize: 20, bold: true, color: C.white, fontFace: F.sans } },
      { text: '/Month', options: { fontSize: 12, bold: true, color: C.white, fontFace: F.sans } }
    ], { x: d[0] + 0.262, y: 3.213, w: 2.493, h: 0.438, align: 'center', valign: 'top' });
    body(s, T.price, d[0] + 0.385, 4.011, 2.234, 1.161, { color: C.white, align: 'center' });
    textShape(s, 'roundRect', 'More Information', {
      x: d[0] + 0.385, y: 5.148, w: 2.234, h: 0.499, fill: { color: C.white },
      rectRadius: R(0.5, 2.234, 0.499), fontSize: 10.5, bold: true, color: d[1]
    });
  });
  headingCenter(s, 'Pricing Table');
});

// 15 — snake arrow infographic
slides.push(function (s) {
  band(s, 8.166, 2.945, 2.518, 1.013, C.tan, true);
  band(s, 9.779, 3.651, 2.142, 1.008, C.brownDark, false);
  band(s, 8.166, 4.352, 2.285, 1.01, C.green, true);
  band(s, 9.779, 5.055, 2.142, 1.008, C.brown, false);
  // heads: odd rows point right, even rows point left
  [[10.758, 3.10, C.tan, 90], [10.237, 3.80, C.brownDark, 270],
   [9.996, 4.505, C.green, 90], [10.237, 5.21, C.brown, 270]].forEach(function (d) {
    arrowHead(s, d[0], d[1], 0.38, 0.695, d[3], d[2]);
  });
  s.addText('Location', { x: 11.043, y: 2.964, w: 1.152, h: 0.37, fontFace: F.sans, fontSize: 16, color: C.ink, valign: 'top', wrap: false });
  s.addText('Start', { x: 8.923, y: 5.768, w: 0.752, h: 0.37, fontFace: F.sans, fontSize: 16, color: C.ink, align: 'right', valign: 'top', wrap: false });
  [[1.203, 2.822, '01.', C.brown], [4.544, 2.822, '02.', C.green],
   [1.203, 4.649, '03.', C.green], [4.544, 4.649, '04.', C.brown]].forEach(function (d) {
    s.addText(d[2], { x: d[0], y: d[1], w: 0.908, h: 0.64, fontFace: F.sans, fontSize: 32, bold: true, color: d[3], align: 'justify', valign: 'top' });
    label(s, 'Tittle Here', d[0], d[1] + 0.593, 2.333, { color: d[3] });
    body(s, T.arrowStep, d[0], d[1] + 0.904, 2.601, 0.601);
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 16 — split cards with percentage rings
slides.push(function (s) {
  [[2.171, 4.831, 5.143, 1.72, 1.742, 1.732, 2.563, 3.032, 3.034, 3.311, C.brown, '64%', C.brown],
   [2.171, 4.831, 5.143, 1.72, 1.742, 3.549, 4.38, 4.849, 4.851, 5.128, C.green, '76%', C.green],
   [7.674, 10.333, 10.645, 7.223, 7.245, 1.732, 2.563, 3.032, 3.034, 3.311, C.green, '89%', C.green],
   [7.674, 10.333, 10.645, 7.223, 7.245, 3.549, 4.38, 4.849, 4.851, 5.128, C.brown, '42%', C.brown]
  ].forEach(function (d) {
    shape(s, 'round2SameRect', { x: d[0], y: d[5], w: 1.66, h: 3.491, fill: { color: C.creamDeep }, rotate: 270, rectRadius: R(0.167, 1.66, 3.491) });
    shape(s, 'round2SameRect', { x: d[1], y: d[6], w: 1.66, h: 1.829, fill: { color: d[10] }, rotate: 90, rectRadius: R(0.167, 1.66, 1.829) });
    label(s, 'Tittle Here', d[3], d[8], 2.539, { color: d[12] });
    body(s, T.ringItem, d[4], d[9], 2.539, 0.601);
    textShape(s, 'ellipse', d[11], {
      x: d[2], y: d[7], w: 0.936, h: 0.936, fill: { color: d[10] }, line: { color: C.white, width: 4 },
      fontSize: 16, bold: true, color: C.white, margin: 0
    });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 17 — two big-percentage cards
slides.push(function (s) {
  [[1.28, C.brown, '64%', 1.693, 3.641, 3.281], [6.867, C.green, '78%', 7.365, 9.314, 8.953]].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: 3.764, w: 5.169, h: 2.208, fill: { color: d[1] }, rectRadius: R(0.167, 5.169, 2.208) });
    s.addText('Tittle Here', { x: d[3], y: 4.151, w: 1.53, h: 0.37, fontFace: F.sans, fontSize: 12, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText(d[2], { x: d[3], y: 4.334, w: 1.588, h: 1.111, fontFace: F.sans, fontSize: 40, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
    body(s, T.bigPct, d[4], 4.172, 2.254, 1.426, { color: C.white });
    shape(s, 'line', { x: d[5], y: 4.184, w: 0, h: 1.389, line: { color: C.white, width: 1 } });
  });
  body(s, T.wideIntro, 2.072, 2.665, 8.96, 0.601, { align: 'center' });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 18 — price list
slides.push(function (s) {
  [[1.175, 2.664, C.brown, '$56'], [1.175, 3.943, C.green, '$70'], [1.175, 5.222, C.brown, '$63'],
   [7.197, 2.664, C.green, '$88'], [7.197, 3.943, C.brown, '$98'], [7.197, 5.222, C.green, '$76']
  ].forEach(function (d) {
    textShape(s, 'roundRect', d[3], {
      x: d[0], y: d[1], w: 1.066, h: 0.917, fill: { color: d[2] }, rectRadius: R(0.1667, 1.066, 0.917),
      flipH: true, fontSize: 18, bold: true, color: C.white
    });
    shape(s, 'roundRect', { x: d[0] + 1.318, y: d[1] + 0.071, w: 0.088, h: 0.77, fill: { color: d[2] }, rectRadius: R(0.1667, 0.088, 0.77) });
    s.addText('Tittle Here', { x: d[0] + 1.576, y: d[1] + 0.013, w: 1.098, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: d[2], valign: 'top', wrap: false });
    body(s, T.priceRow, d[0] + 1.576, d[1] + 0.286, 3.445, 0.601);
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 19 — twelve chips
slides.push(function (s) {
  const cols = [1.443, 5.465, 9.486];
  const rows = [2.648, 3.586, 4.524, 5.463];
  rows.forEach(function (y, r) {
    cols.forEach(function (x, c) {
      const raised = (r + c) % 2 === 0;
      const n = r * 3 + c + 1;
      shape(s, 'roundRect', Object.assign({
        x: x, y: y, w: 2.67, h: 0.655, fill: { color: raised ? C.cream : C.creamDeep },
        rectRadius: R(0.082, 2.67, 0.655)
      }, raised ? { shadow: shadow() } : {}));
      textShape(s, 'ellipse', (n < 10 ? '0' : '') + n, {
        x: x - 0.261, y: y + 0.061, w: 0.532, h: 0.532, fill: { color: raised ? C.brown : C.green },
        shadow: shadow(90), fontSize: 9, bold: true, color: C.white
      });
      s.addText('Your Tittle Here', { x: x + 0.602, y: y + 0.171, w: 1.503, h: 0.303, fontFace: F.sans, fontSize: 12, color: C.body, align: 'justify', valign: 'top' });
    });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 20 — circles over labelled cards
slides.push(function (s) {
  [[1.793, C.tan, '01', 1.171, C.tan, '1'], [4.695, C.brown, '02', 4.036, C.brown, '2'],
   [7.598, C.tan, '03', 6.901, C.tan, '3'], [10.5, C.green, '04', 9.765, C.green, '4']].forEach(function (d) {
    textShape(s, 'ellipse', d[2], {
      x: d[0], y: 2.646, w: 1.068, h: 1.068, fill: { color: d[1] }, shadow: shadow(),
      fontSize: 18, bold: true, color: C.white
    });
    label(s, 'Tittle Here', d[0] - 0.87, 3.822, 2.808, { align: 'center' });
    body(s, T.circleItem, d[0] - 0.87, 4.08, 2.808, 0.601, { align: 'center' });
    shape(s, 'roundRect', { x: d[3], y: 5.014, w: 2.392, h: 1.103, fill: { color: d[4] }, rectRadius: R(0.099, 2.392, 1.103) });
    textShape(s, 'ellipse', d[5], {
      x: d[3] + 0.276, y: 5.285, w: 0.463, h: 0.463, fill: { color: C.white, transparency: 90 },
      fontSize: 12, bold: true, color: C.white
    });
    label(s, 'Tittle Here', d[3] + 0.881, 5.285, 1.191, { color: C.white });
    body(s, 'Lorem ipsum', d[3] + 0.881, 5.504, 1.191, 0.336, { color: C.white, italic: true });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 21 — year bars
slides.push(function (s) {
  [[1.275, 2.686, C.brown, '2021'], [1.275, 3.944, C.green, '2023'], [1.275, 5.203, C.brown, '2025'],
   [6.844, 2.686, C.green, '2022'], [6.844, 3.944, C.brown, '2024'], [6.844, 5.203, C.green, '2026']
  ].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: d[1], w: 5.225, h: 0.894, fill: { color: d[2] }, rectRadius: R(0.082, 5.225, 0.894) });
    shape(s, 'roundRect', { x: d[0] + 0.194, y: d[1] + 0.15, w: 0.115, h: 0.594, fill: { color: C.white, transparency: 80 }, rectRadius: R(0.5, 0.115, 0.594) });
    s.addText(d[3], { x: d[0] + 0.576, y: d[1] + 0.195, w: 1.21, h: 0.505, fontFace: F.sans, fontSize: 24, bold: true, color: C.white, valign: 'top' });
    body(s, T.timeline, d[0] + 1.816, d[1] + 0.123, 3.061, 0.601, { color: C.white });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 22 — year cards
slides.push(function (s) {
  [[1.181, 2.648, C.brown, '2021'], [1.181, 4.51, C.green, '2024'],
   [5.007, 2.648, C.green, '2022'], [5.007, 4.51, C.brown, '2025'],
   [8.833, 2.648, C.brown, '2023'], [8.833, 4.51, C.green, '2026']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: d[1], w: 3.347, h: 1.583, fill: { color: d[2] }, rectRadius: R(0.082, 3.347, 1.583) });
    shape(s, 'roundRect', { x: d[0] + 0.681, y: d[1] - 0.027, w: 0.078, h: 0.662, fill: { color: C.tan, transparency: 80 }, rotate: 270, rectRadius: R(0.5, 0.078, 0.662) });
    s.addText(d[3], { x: d[0] + 0.278, y: d[1] + 0.426, w: 1.03, h: 0.505, fontFace: F.sans, fontSize: 24, bold: true, color: C.white, valign: 'top' });
    body(s, T.yearCard, d[0] + 1.308, d[1] + 0.39, 1.678, 0.896, { color: C.white });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 23 — sliders + year pills
slides.push(function (s) {
  shape(s, 'roundRect', { x: 1.178, y: 2.648, w: 5.489, h: 3.443, fill: { color: C.cream }, rectRadius: R(0.082, 5.489, 3.443), shadow: shadow() });
  [[3.106, 3.463, 3.119, 4.669, '01', '83%', C.brown],
   [4.044, 4.401, 2.519, 4.07, '02', '64%', C.green],
   [4.982, 5.339, 3.525, 5.081, '03', '91%', C.brown]].forEach(function (d) {
    s.addText('Your Tittle Here ' + d[4], { x: 1.572, y: d[0], w: 2.596, h: 0.34, fontFace: F.sans, fontSize: 12, color: C.ink, align: 'justify', lineSpacingMultiple: 1.3, valign: 'top' });
    shape(s, 'roundRect', { x: 1.649, y: d[1], w: 3.919, h: 0.102, fill: { color: 'D9D9D9', transparency: 50 }, rectRadius: 0.051 });
    shape(s, 'roundRect', { x: 1.649, y: d[1] - 0.008, w: d[2], h: 0.118, fill: { color: d[6] }, rectRadius: 0.059 });
    shape(s, 'ellipse', { x: d[3], y: d[1] - 0.053, w: 0.197, h: 0.197, fill: { color: d[6] } });
    s.addText(d[5], { x: 5.568, y: d[0] + 0.305, w: 0.704, h: 0.34, fontFace: F.sans, fontSize: 12, color: C.ink, align: 'justify', lineSpacingMultiple: 1.3, valign: 'top' });
  });
  [[2.686, '2021', C.green, false], [3.944, '2022', C.brown, true], [5.203, '2023', C.green, false]].forEach(function (d) {
    shape(s, 'roundRect', Object.assign({
      x: 6.957, y: d[0], w: 5.225, h: 0.894, fill: { color: d[3] ? C.cream : C.creamDeep }, rectRadius: R(0.5, 5.225, 0.894)
    }, d[3] ? { shadow: shadow() } : {}));
    shape(s, 'roundRect', { x: 7.328, y: d[0] + 0.287, w: 0.077, h: 0.32, fill: { color: d[2] }, rectRadius: R(0.5, 0.077, 0.32) });
    s.addText(d[1], { x: 7.534, y: d[0] + 0.195, w: 1.21, h: 0.505, fontFace: F.sans, fontSize: 24, bold: true, color: C.ink, valign: 'top' });
    body(s, T.timeline, 8.774, d[0] + 0.123, 3.061, 0.601);
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 24 — badge cards with learn-more buttons
slides.push(function (s) {
  [[1.178, C.brown, '01'], [4.119, C.green, '02'], [7.06, C.brown, '03'], [10.001, C.green, '04']].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: 2.659, w: 2.18, h: 1.73, fill: { color: C.creamDeep }, rectRadius: 0.174 });
    shape(s, 'ellipse', { x: d[0] + 0.425, y: 2.857, w: 1.334, h: 1.334, fill: { color: C.cream } });
    textShape(s, 'roundRect', d[2], {
      x: d[0] + 0.658, y: 3.09, w: 0.869, h: 0.869, fill: { color: d[1] }, rectRadius: R(0.5, 0.869, 0.869),
      shadow: shadow(), fontSize: 18, bold: true, color: C.white
    });
    s.addText('Tittle Here', { x: d[0], y: 4.591, w: 2.18, h: 0.337, fontFace: F.sans, fontSize: 14, bold: true, color: C.body, align: 'center', valign: 'top' });
    body(s, T.learn, d[0], 4.883, 2.18, 0.601, { align: 'center' });
    textShape(s, 'roundRect', 'Learn More', {
      x: d[0] + 0.43, y: 5.661, w: 1.32, h: 0.435, fill: { color: d[1] }, rectRadius: R(0.5, 1.32, 0.435),
      shadow: shadow(), fontSize: 10, bold: true, color: C.white
    });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 25 — donut of percentage bubbles
slides.push(function (s) {
  // ring: outer disc minus a cream hole (63% of the diameter)
  shape(s, 'ellipse', { x: 1.461, y: 2.66, w: 3.442, h: 3.442, fill: { color: C.brown } });
  shape(s, 'ellipse', { x: 2.094, y: 3.293, w: 2.176, h: 2.176, fill: { color: C.cream } });
  [[1.99, 2.668, C.tan, '50%'], [3.364, 2.668, C.green, '98%'], [4.076, 3.906, C.tan, '65%'],
   [1.303, 3.906, C.green, '77%'], [1.99, 5.122, C.tan, '41%'], [3.364, 5.122, C.green, '29%']].forEach(function (d) {
    textShape(s, 'ellipse', d[3], {
      x: d[0], y: d[1], w: 0.987, h: 0.987, fill: { color: d[2] }, fontSize: 16, bold: true, color: C.white
    });
  });
  iconCloud(s, 2.821, 4.151, 0.685, 0.498, C.brown);
  [[6.015, 6.823, 2.815, '01', C.brown], [9.46, 10.268, 2.815, '02', C.brown],
   [6.015, 6.823, 4.067, '03', C.green], [9.46, 10.268, 4.067, '04', C.brown],
   [6.015, 6.823, 5.32, '05', C.brown], [9.46, 10.268, 5.32, '06', C.brown]].forEach(function (d) {
    textShape(s, 'ellipse', d[3], {
      x: d[0], y: d[2], w: 0.634, h: 0.634, fill: { color: d[4] }, fontSize: 14, bold: true, color: C.white
    });
    body(s, T.donutItem, d[1], d[2] + 0.002, 1.919, 0.601);
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 26 — segmented cycle
slides.push(function (s) {
  // six arc segments running clockwise, each capped with an arrow head.
  // angles follow OOXML: 0 = 3 o'clock, growing clockwise.
  const cyc = { cx: 6.65, cy: 4.395, rx: 1.79, ry: 1.715, band: 0.36 };
  const mid = 1 - cyc.band / 2;                       // radius factor of the band centre line
  [[205, C.brownPale], [265, C.brownLight], [325, C.brown],
   [25, C.brownMid], [85, C.brownDark], [145, C.brown]].forEach(function (d) {
    shape(s, 'blockArc', {
      x: cyc.cx - cyc.rx, y: cyc.cy - cyc.ry, w: cyc.rx * 2, h: cyc.ry * 2,
      fill: { color: d[1] }, angleRange: [d[0], d[0] + 47], arcThicknessRatio: cyc.band
    });
    const tip = d[0] + 55;                            // head sits just past the arc end
    const a = tip * Math.PI / 180;
    arrowHead(s, cyc.cx + cyc.rx * mid * Math.cos(a), cyc.cy + cyc.ry * mid * Math.sin(a),
      0.44, 0.88, tip + 180, d[1]);
  });
  [[5.687, 3.068, 0.395, '01'], [7.145, 3.037, 0.411, '02'], [7.829, 4.25, 0.411, '03'],
   [7.139, 5.458, 0.419, '04'], [5.749, 5.475, 0.409, '05'], [5.038, 4.247, 0.414, '06']].forEach(function (d) {
    s.addText(d[3], { x: d[0], y: d[1], w: d[2], h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'top', wrap: false });
  });
  s.addText('95%', { x: 5.947, y: 4.009, w: 1.401, h: 0.774, fontFace: F.sans, fontSize: 40, bold: true, color: C.brown, align: 'center', valign: 'top', wrap: false });
  [[3.681, 2.97, '01', C.brown, 0.938, 2.891, 2.441, 'right'],
   [3.681, 4.819, '03', C.green, 0.938, 4.74, 2.441, 'right'],
   [9.054, 2.97, '02', C.green, 9.991, 2.891, 2.28, 'left'],
   [9.054, 4.819, '04', C.brown, 9.991, 4.74, 2.28, 'left']].forEach(function (d) {
    textShape(s, 'ellipse', d[2], { x: d[0], y: d[1], w: 0.634, h: 0.634, fill: { color: d[3] }, fontSize: 14, bold: true, color: C.white });
    body(s, T.cycleItem, d[4], d[5], d[6], 1.161, { align: d[7] });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 27 — four cards + checklist
slides.push(function (s) {
  [[1.173, 2.667, C.brown, '01', 2.986, 3.286], [1.173, 4.556, C.green, '03', 4.867, 5.167],
   [5.115, 2.667, C.green, '02', 2.986, 3.286], [5.115, 4.556, C.brown, '04', 4.867, 5.167]].forEach(function (d) {
    shape(s, 'roundRect', { x: d[0], y: d[1], w: 3.585, h: 1.555, fill: { color: d[2] }, rectRadius: R(0.187, 3.585, 1.555) });
    textShape(s, 'ellipse', d[3], {
      x: d[0] + 0.322, y: d[1] + 0.342, w: 0.854, h: 0.854, fill: { color: C.white, transparency: 85 },
      fontSize: 20, bold: true, color: C.white
    });
    s.addText('Tittle Here', { x: d[0] + 1.388, y: d[4], w: 1.777, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: C.white, valign: 'top' });
    body(s, T.cardItem, d[0] + 1.388, d[5], 1.877, 0.601, { color: C.white });
  });
  [2.746, 3.937, 5.127].forEach(function (y) {
    iconCheck(s, 9.646, y + 0.09, 0.239, C.brown);
    body(s, T.checkItem, 9.986, y, 2.263, 0.896);
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 28 — year timeline
slides.push(function (s) {
  [[1.177, C.brown, '2020', C.brown, '01'], [4.194, C.green, '2021', C.green, '02'],
   [7.222, C.brown, '2022', C.brown, '03'], [10.25, C.green, '2023', C.green, '04']].forEach(function (d, i) {
    shape(s, 'ellipse', { x: d[0], y: 2.667, w: 1.917, h: 1.917, fill: { color: d[1] } });
    if (i < 3) shape(s, 'downArrow', { x: d[0] + 2.245, y: 3.422, w: 0.433, h: 0.452, fill: { color: C.creamDeep }, rotate: 270 });
    s.addText(d[2], { x: d[0] + 0.172, y: 3.158, w: 1.573, h: 0.707, fontFace: F.sans, fontSize: 36, bold: true, color: C.white, align: 'center', valign: 'top' });
    s.addText('Subtittle', { x: d[0] + 0.302, y: 3.691, w: 1.313, h: 0.37, fontFace: F.sans, fontSize: 12, color: C.white, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
    textShape(s, 'ellipse', d[4], { x: d[0] + 0.641, y: 4.785, w: 0.634, h: 0.634, fill: { color: d[3] }, fontSize: 14, bold: true, color: C.white });
    body(s, T.stepItem, d[0] - 0.144, 5.5, 2.206, 0.631, { align: 'center' });
  });
  nav(s);
  headingCenter(s, 'Creative Infographic');
});

// 29 — contact
slides.push(function (s) {
  shape(s, 'rect', { x: 0, y: 0, w: 3.729, h: 7.5, fill: { color: C.brown } });
  // tilted tablet mock-up: tan back plate peeking out bottom-left, black bezel, grey screen
  shape(s, 'roundRect', { x: 1.569, y: 1.932, w: 4.72, h: 3.68, fill: { color: C.tan }, rectRadius: 0.26, rotate: 328 });
  shape(s, 'roundRect', { x: 1.649, y: 1.922, w: 4.66, h: 3.62, fill: { color: '111111' }, rectRadius: 0.24, rotate: 328 });
  imagePlaceholder(s, 1.726, 1.999, 4.506, 3.466, { rotate: 328, rectRadius: 0.16 });
  nav(s, { brandColor: C.white });
  heading(s, 'Our Contact', 8.206, 1.621, 4.231);
  body(s, T.contact, 8.206, 2.92, 4.044, 0.601, { color: C.bodyMid });
  iconGlobe(s, 8.305, 4.127, 0.244, C.brown);
  iconHome(s, 8.295, 4.675, 0.264, 0.213, C.green);
  iconMail(s, 8.299, 5.232, 0.257, 0.159, C.brown);
  iconPhone(s, 8.316, 5.733, 0.222, 0.22, C.green);
  [[4.107, 'www.yourwebsite.com', 2.416], [4.638, '0123 Street, Country 0123', 2.532],
   [5.169, 'nametittle@email.com', 2.223], [5.7, '+123 0000 0000 0000', 2.223]].forEach(function (d) {
    s.addText(d[1], { x: 8.947, y: d[0], w: d[2], h: 0.286, fontFace: F.sans, fontSize: 11, color: C.bodyMid, valign: 'top' });
  });
});

// 30 — thank you
slides.push(function (s) {
  shape(s, 'rect', { x: 6.083, y: 2.417, w: 7.25, h: 5.083, fill: { color: C.brown } });
  nav(s);
  heading(s, 'Thank You', 1.093, 1.808, 3.326);
  body(s, T.thanks, 1.092, 3.094, 2.554, 1.161);
  body(s, T.thanks, 1.092, 4.699, 2.554, 1.161);
  imagePlaceholder(s, 7.479, 1.417, 4.708, 4.688, { rectRadius: 0.365 });
  imagePlaceholder(s, 5.027, 1.417, 2.241, 2.231, { rectRadius: 0.382 });
  imagePlaceholder(s, 5.027, 3.873, 2.241, 2.231, { rectRadius: 0.382 });
});

// =================================================================== build ==
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'CUPFE', width: 13.333, height: 7.5 });
pptx.layout = 'CUPFE';
pptx.title = 'Cupfe Coffee';

slides.forEach(function (build) {
  const slide = pptx.addSlide();
  slide.background = { color: C.cream };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '02cc4297-c864-43f8-9580-25691b151b5a_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
