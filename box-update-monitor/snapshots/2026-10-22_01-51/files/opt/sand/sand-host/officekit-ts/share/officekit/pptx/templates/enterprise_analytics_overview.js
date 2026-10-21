/**
 * Enterprise Analytics — 20 slide deck rebuilt with pptxgenjs.
 * Run: node 08245316-ad10-4892-b91d-8b657be24cb6_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const C = {
  green: 'BFE462',      // accent1
  greenMid: 'A3D525',   // accent2
  greenDark: '668D05',  // accent3
  greenDeep: '9FD124',  // gradient partner of accent1
  olive: '7AA01C',
  dark: '404040',       // tx1 lum75
  gray: '595959',
  grayLt: 'D9D9D9',
  panel: '262626',
  white: 'FFFFFF',
  black: '000000',
  ink: '1A1A1A',
  paleGreen: 'F2FAE0',
  star: 'FFC000',
  imgFill: '7E7D7E',
};

const F_HEAD = 'Fira Sans';   // +mj-lt
const F_BODY = 'Roboto';      // +mn-lt
const F_CARD = 'Be Vietnam Pro SemiBold';

const NO_LINE = { type: 'none' };

// Shadows measured from the source deck (blurRad / dist / alpha).
// pptxgenjs rewrites the options object in place, so hand it a fresh copy every time.
const shadow = (blur, offset, opacity, angle) => () =>
  ({ type: 'outer', angle: angle || 45, blur, offset, color: C.black, opacity });
const SH_SOFT = shadow(19, 12, 0.06);
const SH_TINY = shadow(16, 6, 0.08);
const SH_WIDE = shadow(24, 21, 0.05);
const SH_FLAT = shadow(41, 0, 0.03, 90);
const SH_CHIP = shadow(17, 12, 0.12, 90);

// Repeated body copy in the source deck.
const LOREM_FULL =
  'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
  'By integrating advanced analytics across departments, businesses can make better decisions.';
const LOREM_MED =
  'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
  'By integrating advanced analytics.';
const LOREM_SHORT = 'Enterprise Analytics empowers organizations to transform data into actionable intelligence.';
const QUOTE = '"Without data, you\u2019re just another person with an opinion." ';
const DEMING = 'W. Edwards Deming';

/* ------------------------------------------------------------------ *
 * Small building blocks
 * ------------------------------------------------------------------ */

function rect(s, x, y, w, h, fill, opt) {
  const o = Object.assign({ rectRadius: 0 }, opt || {});
  const shape = o.rectRadius > 0 ? 'roundRect' : 'rect';
  s.addShape(shape, Object.assign(
    { x, y, w, h, fill: typeof fill === 'string' ? { color: fill } : fill, line: NO_LINE }, o));
}

/**
 * Corner-to-corner black wash (dense top-left, lighter bottom-right).
 * pptxgenjs has no gradient fill, so the ramp is painted as a stack of thin
 * rectangles laid perpendicular to the box diagonal, each a solid step of the
 * black-over-white ramp. `clip` masks the bleed for washes that cover a band.
 */
function diagWash(s, x0, y0, w, h, alphaTL, alphaBR, clip) {
  const n = 30;
  const span = (w + h) / Math.SQRT2;                 // box diagonal projected on the 45-deg axis
  const step = span / n;
  for (let i = 0; i < n; i++) {
    const f = (i + 0.5) / n;
    const level = Math.round(255 * (1 - (alphaTL + (alphaBR - alphaTL) * f)));
    const hex = level.toString(16).toUpperCase().padStart(2, '0').repeat(3);
    const d = (span * f) / Math.SQRT2;               // offset along x and y
    const bh = step * 1.4;                           // slight overlap hides the seams
    s.addShape('rect', {
      x: x0 + d - 12, y: y0 + d - bh / 2, w: 24, h: bh,
      fill: { color: hex }, line: NO_LINE, rotate: 315,
    });
  }
  if (clip) {                                        // trim the bleed above and below a band
    rect(s, 0, 0, 13.333, y0, C.white);
    rect(s, 0, y0 + h, 13.333, 7.5 - y0 - h, C.white);
  }
}

/** Rounded square rotated 45 deg — the deck's signature diamond motif. */
function diamond(s, cx, cy, size, fill, opt) {
  const o = Object.assign({ rot: 315, radius: size * 0.24 }, opt || {});
  s.addShape('roundRect', {
    x: cx - size / 2, y: cy - size / 2, w: size, h: size,
    fill: typeof fill === 'string' ? { color: fill } : fill,
    line: NO_LINE, rectRadius: o.radius, rotate: o.rot,
  });
}

function oval(s, x, y, d, fill, opt) {
  s.addShape('ellipse', Object.assign(
    { x, y, w: d, h: d, fill: typeof fill === 'string' ? { color: fill } : fill, line: NO_LINE },
    opt || {}));
}

// pptxgenjs omits the inset attributes when `margin` is absent, which lands on
// PowerPoint's defaults (0.1" sides / 0.05" top-bottom) — same as the source deck.
function txt(s, body, o) {
  s.addText(body, Object.assign({ fontFace: F_BODY, color: C.dark, valign: 'top' }, o));
}

/** Big two-tone slide heading (36 pt, Fira Sans bold). */
function heading(s, x, y, w, lines, opt) {
  const o = Object.assign({ size: 36, h: 0.71 * lines.length, align: 'left' }, opt || {});
  const body = [];
  lines.forEach((line, li) => {
    line.forEach((run, ri) => {
      body.push({
        text: run[0],
        options: {
          color: run[1] || C.dark, bold: true, fontSize: o.size, fontFace: F_HEAD,
          breakLine: li < lines.length - 1 && ri === line.length - 1,
        },
      });
    });
  });
  txt(s, body, { x, y, w, h: o.h, align: o.align });
}

function subtitle(s, x, y, text, opt) {
  txt(s, text, Object.assign({ x, y, w: 5.321, h: 0.337, fontSize: 14 }, opt || {}));
}

function bodyText(s, x, y, w, h, text, opt) {
  txt(s, text, Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.5 }, opt || {}));
}

/** "+287" over a caption — used on almost every slide. */
function stat(s, x, y, value, caption, opt) {
  const o = Object.assign(
    { size: 28, valueColor: C.greenMid, capColor: C.dark, align: 'left', w: 1.674, capW: 1.961, capBold: false },
    opt || {});
  txt(s, value, { x, y, w: o.w, h: 0.572, fontSize: o.size, bold: true, fontFace: F_HEAD, color: o.valueColor, align: o.align });
  txt(s, caption, {
    x: o.align === 'right' ? x - (o.capW - o.w) : x, y: y + 0.46, w: o.capW, h: 0.376,
    fontSize: 12, bold: o.capBold, color: o.capColor, align: o.align, lineSpacingMultiple: 1.5,
  });
}

/** "235+" over a small grey caption (36 pt variant). */
function statBig(s, x, y, value, caption, opt) {
  const o = Object.assign({ valueColor: C.green, capColor: C.gray, capW: 1.405 }, opt || {});
  txt(s, value, { x, y, w: 2.003, h: 0.707, fontSize: 36, bold: true, fontFace: F_HEAD, color: o.valueColor, valign: 'middle' });
  txt(s, caption, { x, y: y + 0.619, w: o.capW, h: 0.303, fontSize: 12, fontFace: F_HEAD, color: o.capColor, valign: 'middle' });
}

/** White pill with centred label (e.g. "Enterprise Data Analytics"). */
function pillLabel(s, x, y, w, parts, opt) {
  const o = Object.assign({ h: 0.5, fill: C.white, color: C.dark }, opt || {});
  rect(s, x, y, w, o.h, o.fill, { rectRadius: o.h / 2, shadow: SH_SOFT() });
  txt(s, parts, { x: x + 0.05, y: y + 0.098, w: w - 0.1, h: 0.303, fontSize: 12, align: 'center', color: o.color });
}

/** Rounded card with a label, a percentage and a green progress track. */
function progressCard(s, x, y, parts, pct, pctLabel) {
  rect(s, x, y, 4.236, 0.956, C.white, { rectRadius: 0.315, shadow: SH_WIDE() });
  txt(s, parts, { x: x + 0.245, y: y + 0.185, w: 2.454, h: 0.303, fontSize: 12, fontFace: F_HEAD, valign: 'middle' });
  txt(s, pctLabel, { x: x + 2.47, y: y + 0.168, w: 1.518, h: 0.337, fontSize: 14, fontFace: F_HEAD, align: 'right', valign: 'middle' });
  const trackX = x + 0.836, trackW = 2.983;
  s.addShape('line', { x: trackX, y: y + 0.655, w: trackW, h: 0, line: { color: C.green, width: 6.5, transparency: 80 } });
  s.addShape('line', { x: x + 0.394, y: y + 0.655, w: 0.442 + trackW * pct, h: 0, line: { color: C.green, width: 6.5 } });
  oval(s, trackX + trackW * pct - 0.098, y + 0.557, 0.196, C.green, { line: { color: C.white, width: 3.5 }, shadow: SH_TINY() });
}

/* --- pictograms ---------------------------------------------------- */

/** Shopping-cart glyph. */
function iconCart(s, x, y, color) {
  s.addShape('roundRect', { x: x + 0.03, y, w: 0.20, h: 0.15, fill: { color }, line: NO_LINE, rectRadius: 0.03 });
  s.addShape('rect', { x, y: y + 0.005, w: 0.06, h: 0.03, fill: { color }, line: NO_LINE });
  oval(s, x + 0.055, y + 0.22, 0.044, color);
  oval(s, x + 0.165, y + 0.22, 0.044, color);
  oval(s, x + 0.154, y + 0.022, 0.088, color);
}

/** Monitor / display glyph. */
function iconScreen(s, x, y, color) {
  s.addShape('roundRect', { x, y, w: 0.25, h: 0.175, fill: { color }, line: NO_LINE, rectRadius: 0.03 });
  s.addShape('rect', { x: x + 0.105, y: y + 0.175, w: 0.04, h: 0.045, fill: { color }, line: NO_LINE });
  s.addShape('rect', { x: x + 0.05, y: y + 0.21, w: 0.15, h: 0.019, fill: { color }, line: NO_LINE });
}

/** Bar-chart glyph. */
function iconChart(s, x, y, color) {
  [[0.0, 0.09], [0.055, 0.15], [0.11, 0.21]].forEach(([dx, h]) => {
    s.addShape('rect', { x: x + dx, y: y + 0.25 - h, w: 0.036, h, fill: { color }, line: NO_LINE });
  });
  s.addShape('rect', { x: x + 0.175, y: y + 0.02, w: 0.032, h: 0.23, fill: { color }, line: NO_LINE });
  s.addShape('rect', { x: x + 0.155, y: y + 0.09, w: 0.072, h: 0.028, fill: { color }, line: NO_LINE });
}

/** Rounded blob with a notch — the deck's little "seed" mark. */
function iconLeaf(s, x, y, d, color, notch) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: NO_LINE });
  s.addShape('rect', { x: x + d * 0.46, y: y + d * 0.44, w: d * 0.5, h: d * 0.1, fill: { color: notch || C.white }, line: NO_LINE });
  s.addShape('rect', { x: x + d * 0.6, y: y + d * 0.56, w: d * 0.36, h: d * 0.14, fill: { color: notch || C.white }, line: NO_LINE });
}

/** Bank / institution glyph. */
function iconBank(s, x, y, d, color) {
  s.addShape('triangle', { x, y, w: d, h: d * 0.3, fill: { color }, line: NO_LINE });
  [0.12, 0.42, 0.72].forEach((f) => {
    s.addShape('rect', { x: x + d * f, y: y + d * 0.35, w: d * 0.14, h: d * 0.4, fill: { color }, line: NO_LINE });
  });
  s.addShape('rect', { x, y: y + d * 0.8, w: d, h: d * 0.16, fill: { color }, line: NO_LINE });
}

/** Compass glyph. */
function iconCompass(s, x, y, d, color) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color, width: 1.5 } });
  s.addShape('triangle', { x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, fill: { color }, line: NO_LINE, rotate: 135 });
}

/** Three rounded tiles holding the cart / screen / chart glyphs. */
function iconTiles(s, x, y, opt) {
  const o = Object.assign({ size: 0.821, gap: 1.038, activeFill: C.dark, activeIcon: C.green, icon: C.dark }, opt || {});
  const tiles = [
    { fill: o.activeFill, color: o.activeIcon, draw: iconCart, shadow: null },
    { fill: C.white, color: o.icon, draw: iconScreen, shadow: SH_SOFT },
    { fill: C.white, color: o.icon, draw: iconChart, shadow: SH_SOFT },
  ];
  tiles.forEach((t, i) => {
    const tx = x + i * o.gap;
    rect(s, tx, y, o.size, o.size, t.fill, { rectRadius: 0.2, shadow: t.shadow && t.shadow() });
    t.draw(s, tx + (o.size - 0.25) / 2, y + (o.size - 0.25) / 2, t.color);
  });
}

/** Small white circle with the leaf mark. */
function leafBadge(s, x, y, opt) {
  const o = Object.assign({ d: 0.378, disc: C.white, leaf: C.green }, opt || {});
  oval(s, x, y, o.d, o.disc, { shadow: SH_TINY() });
  iconLeaf(s, x + 0.116, y + 0.116, 0.148, o.leaf, o.disc);
}

/** Big white circle with the leaf mark (0.807"). */
function leafDisc(s, x, y, opt) {
  const o = Object.assign({ disc: C.white, leaf: C.green }, opt || {});
  oval(s, x, y, 0.807, o.disc);
  iconLeaf(s, x + 0.243, y + 0.243, 0.322, o.leaf, o.disc);
}

/** Five gold rating stars. */
function stars(s, x, y, opt) {
  const o = Object.assign({ w: 0.153, h: 0.136, gap: 0.1847 }, opt || {});
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', { x: x + i * o.gap, y, w: o.w, h: o.h, fill: { color: C.star }, line: NO_LINE });
  }
}

/* --- page chrome (repeated from the slide master) ------------------- */

const CHROME_LIGHT = { logo: C.dark, pill: { color: C.black, transparency: 75 }, disc: C.dark, icon: C.green };
const CHROME_DARK = { logo: C.green, pill: { color: C.paleGreen, transparency: 80 }, disc: C.green, icon: C.white };

function chrome(s, style, opt) {
  const o = Object.assign({ logo: true }, opt || {});
  if (o.logo) {
    // "logoipsum" mark: a bracket, a half disc and the wordmark
    s.addShape('rect', { x: 0.473, y: 0.437, w: 0.03, h: 0.2, fill: { color: style.logo }, line: NO_LINE });
    s.addShape('rect', { x: 0.473, y: 0.437, w: 0.092, h: 0.028, fill: { color: style.logo }, line: NO_LINE });
    s.addShape('rect', { x: 0.473, y: 0.609, w: 0.092, h: 0.028, fill: { color: style.logo }, line: NO_LINE });
    oval(s, 0.512, 0.466, 0.142, style.logo);
    txt(s, 'logoipsum', {
      x: 0.66, y: 0.415, w: 0.9, h: 0.25, fontSize: 11, bold: true,
      fontFace: F_HEAD, color: style.logo, valign: 'middle', margin: 0,
    });
  }
  rect(s, 10.202, 0.388, 2.223, 0.389, style.pill, { rectRadius: 0.194, shadow: SH_FLAT() });
  txt(s, 'Search Here', { x: 10.36, y: 0.42, w: 0.872, h: 0.28, fontSize: 8, italic: true, color: C.paleGreen, lineSpacingMultiple: 1.5 });
  oval(s, 12.505, 0.391, 0.382, style.disc);
  oval(s, 12.665, 0.519, 0.063, style.icon);                                       // head
  s.addShape('roundRect', { x: 12.649, y: 0.593, w: 0.094, h: 0.052, fill: { color: style.icon }, line: NO_LINE, rectRadius: 0.026 }); // shoulders
}

function pageNum(s, n) {
  txt(s, 'Page', { x: 12.231, y: 7.0, w: 0.846, h: 0.303, fontSize: 12, bold: true, fontFace: F_HEAD, color: C.green, valign: 'middle' });
  txt(s, String(n), { x: 12.589, y: 6.951, w: 0.517, h: 0.399, fontSize: 12, bold: true, color: C.grayLt, align: 'center', valign: 'middle' });
}

/** Stand-in for a raster image from the source deck. */
function imagePlaceholder(s, x, y, w, h, opt) {
  const o = Object.assign({ fill: C.imgFill, radius: 0.25, label: '[image]', text: C.white, rotate: 0 }, opt || {});
  rect(s, x, y, w, h, o.fill, { rectRadius: o.radius, rotate: o.rotate });
  if (o.label) txt(s, o.label, { x, y: y + h / 2 - 0.2, w, h: 0.4, fontSize: 12, color: o.text, align: 'center', valign: 'middle' });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(s) {
  s.background = { color: C.black };
  pageNum(s, 1);                                     // master pager, dimmed by the band below
  rect(s, 0, 0, 13.333, 7.5, C.black);
  heading(s, 0.976, 2.008, 7.846, [[['ENTERPRISE', C.greenMid], [' ', C.green], ['ANALYTICS', C.white]]], { size: 80, h: 2.794 });
  txt(s, '\u201cDriving Smarter Decisions with Data Power\u201d', { x: 1.001, y: 6.028, w: 5.321, h: 0.37, fontSize: 16, color: C.green });

  const glow = { color: C.greenDeep, transparency: 82 };
  diamond(s, 10.202, 4.485, 1.496, glow);
  diamond(s, 11.748, 3.372, 1.496, glow);
  diamond(s, 10.202, 2.185, 1.496, glow);
  diamond(s, 0.043, 3.39, 2.969, { color: C.greenDeep, transparency: 88 });

  bodyText(s, 8.596, 5.614, 4.163, 0.976, LOREM_MED, { color: C.white });
  leafBadge(s, 8.596, 5.103);
  chrome(s, CHROME_DARK);
}

function slide02(s) {
  diamond(s, -0.106, 2.329, 2.083, { color: C.greenDeep, transparency: 88 });
  chrome(s, CHROME_LIGHT);
  heading(s, 0.976, 5.348, 9.957, [[['The Power of ', C.dark], ['Enterprise', C.greenMid], [' Analytics', C.dark]]]);
  subtitle(s, 0.976, 6.241, 'Driving Smarter Decisions with Data Power');
  bodyText(s, 9.922, 2.948, 2.808, 1.279, LOREM_MED);
  leafBadge(s, 1.084, 3.987);
  txt(s, 'Driving Smarter', { x: 1.578, y: 4.025, w: 5.321, h: 0.303, fontSize: 12, bold: true, color: C.white });
  leafDisc(s, 8.872, 1.836);
  stat(s, 9.922, 1.966, '+287', 'Enterprise Analytics', { capBold: true });
  pillLabel(s, 0.38, 2.097, 2.484, [
    { text: 'Enterprise', options: { bold: true } }, { text: ' Data Analytics' },
  ]);
  diamond(s, 13.415, 6.442, 1.496, C.green);
  diamond(s, 12.278, 7.632, 1.496, C.dark);
  pageNum(s, 2);
}

function slide03(s) {
  chrome(s, CHROME_LIGHT);
  rect(s, 3.481, 5.425, 1.545, 0.98, C.white, { rectRadius: 0.24, shadow: SH_SOFT() });
  diamond(s, 1.058, 0.326, 1.496, C.green);
  diamond(s, -0.079, 1.516, 1.496, C.dark);
  heading(s, 1.77, 1.78, 5.949, [
    [['Connecting', C.dark]],
    [['Multiple', C.greenMid], [' Sources', C.dark]],
  ]);
  subtitle(s, 1.77, 3.182, 'Connecting Multiple Sources');
  bodyText(s, 1.77, 4.027, 5.949, 0.976, LOREM_FULL);
  statBig(s, 1.77, 5.425, '30K', [{ text: 'Data ' }, { text: 'Science', options: { bold: true } }], { capW: 1.679 });
  statBig(s, 3.621, 5.425, '235+', [{ text: 'Data ' }, { text: 'Insights', options: { bold: true } }]);
  diamond(s, 8.859, 7.584, 1.793, C.dark);

  rect(s, 7.718, 1.745, 2.771, 1.605, C.white, { rectRadius: 0.28, shadow: SH_FLAT() });
  rect(s, 8.04, 1.962, 0.372, 0.372, C.green, { rectRadius: 0.113, shadow: SH_TINY() });
  iconLeaf(s, 8.152, 2.074, 0.148, C.white, C.green);
  txt(s, 'Dashboard', { x: 7.952, y: 2.458, w: 2.059, h: 0.337, fontSize: 14, bold: true, fontFace: F_HEAD });
  bodyText(s, 7.952, 2.744, 2.624, 0.37, 'Predictive Models');
  pageNum(s, 3);
}

function slide04(s) {
  chrome(s, CHROME_LIGHT);
  heading(s, 4.402, 1.83, 9.957, [[['Why ', C.dark], ['Analytics', C.greenMid], [' Matters', C.dark]]]);
  subtitle(s, 4.402, 2.722, 'Driving Data-Driven Cultures');
  bodyText(s, 4.402, 3.463, 5.18, 1.279,
    'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
    'By integrating advanced analytics across departments, businesses can make better decisions, ' +
    'enhance efficiency, and unlock hidden opportunities.');

  rect(s, 2.968, 5.2, 4.488, 1.1, C.white, { rectRadius: 0.192, shadow: SH_SOFT() });
  rect(s, 7.839, 5.2, 1.969, 1.1, C.white, { rectRadius: 0.192, shadow: SH_SOFT() });
  oval(s, 3.35, 5.497, 0.506, C.green);
  txt(s, '$', { x: 3.35, y: 5.497, w: 0.506, h: 0.506, fontSize: 18, bold: true, color: C.white, align: 'center', valign: 'middle' });
  txt(s, '$367,000+', { x: 4.348, y: 5.363, w: 4.306, h: 0.774, fontSize: 40, bold: true, fontFace: F_HEAD });
  txt(s, '90%', { x: 8.044, y: 5.363, w: 2.201, h: 0.774, fontSize: 40, bold: true, fontFace: F_HEAD, color: C.greenMid });

  rect(s, 11.466, 2.012, 2.083, 1.836, C.green, { rectRadius: 0.313 });
  rect(s, 11.744, 4.147, 2.134, 2.134, C.dark, { rectRadius: 0.32 });
  stat(s, 11.77, 2.517, '+287', 'AI Models', { valueColor: C.white, capColor: C.white });
  stat(s, 12.036, 4.791, '+287', 'Decisions', { valueColor: C.white, capColor: C.white });
  rect(s, -1.138, 4.103, 2.134, 2.134, C.dark, { rectRadius: 0.32 });
  pageNum(s, 4);
}

function slide05(s) {
  chrome(s, CHROME_LIGHT);
  rect(s, 0, 5.296, 13.333, 2.204, '0D0D0D');
  heading(s, 1.044, 1.83, 7.397, [[['Real-Time ', C.dark], ['Decision', C.greenMid], [' Making', C.dark]]]);
  subtitle(s, 1.044, 2.651, 'Faster Insights, Better Actions');
  [C.white, C.white, C.green, C.white].forEach((fill, i) => {
    rect(s, 1.044 + i * 2.837, 3.469, 2.735, 2.735, fill, { rectRadius: 0.53, shadow: SH_SOFT() });
  });
  stat(s, 9.922, 1.965, '+287', 'Cloud Analytics', { capBold: true });
  bodyText(s, 9.801, 4.527, 2.081, 1.279, [
    { text: 'Enterprise', options: { bold: true } },
    { text: ' Analytics empowers organizations to transform data into actionable intelligence.' },
  ], { color: C.gray });
  leafBadge(s, 9.801, 4.016, { disc: C.green, leaf: C.white });
  diamond(s, -0.05, 2.204, 0.684, C.green, { radius: 0.16 });
  pageNum(s, 5);
}

function slide06(s) {
  heading(s, 6.936, 1.83, 5.949, [
    [['Core ', C.dark], ['Components', C.greenMid]],
    [['of Analytics', C.dark]],
  ]);
  subtitle(s, 6.936, 3.232, 'Core Components of Analytics');
  bodyText(s, 6.936, 5.417, 5.949, 0.976, LOREM_FULL);
  iconTiles(s, 7.023, 4.313, { activeIcon: C.green });
  txt(s, '+287', { x: 10.205, y: 4.517, w: 1.674, h: 0.572, fontSize: 28, bold: true, fontFace: F_HEAD, color: C.greenMid });
  chrome(s, CHROME_LIGHT, { logo: false });          // this slide hides the master logo
  pageNum(s, 6);
  pillLabel(s, 1.897, 5.991, 2.484, [
    { text: 'Enterprise ', options: { bold: true } }, { text: 'Data' },
    { text: ' ', options: { bold: true } }, { text: 'Analytics' },
  ]);
  leafDisc(s, 4.415, 1.955);
}

function slide07(s) {
  chrome(s, CHROME_LIGHT);
  diamond(s, 1.929, 6.933, 1.377, C.green);
  heading(s, 6.816, 1.646, 5.949, [
    [['Turning ', C.dark], ['Insights', C.greenMid]],
    [['into Products', C.dark]],
  ]);
  subtitle(s, 6.816, 3.048, 'Recommendations for Success');
  bodyText(s, 6.816, 3.75, 5.949, 0.976, LOREM_FULL);
  progressCard(s, 3.764, 5.206, 'Advanced Business Insights', 0.72, '80%');
  rect(s, 8.264, 5.134, 1.969, 1.1, C.white, { rectRadius: 0.43, shadow: SH_SOFT() });
  txt(s, '90%', { x: 8.524, y: 5.296, w: 2.201, h: 0.774, fontSize: 40, bold: true, fontFace: F_HEAD, color: C.greenMid });
  leafDisc(s, 1.363, 2.141);
  pageNum(s, 7);
}

function slide08(s) {
  chrome(s, CHROME_LIGHT);
  heading(s, 1.17, 1.64, 5.949, [
    [['Recommendations', C.dark]],
    [['for ', C.dark], ['Success', C.greenMid]],
  ]);
  subtitle(s, 1.17, 3.042, 'Staying Ahead of Competition');
  rect(s, -0.369, 4.424, 2.315, 2.315, C.green, { rectRadius: 0.35 });
  rect(s, 11.438, 1.406, 2.195, 2.195, C.dark, { rectRadius: 0.33 });
  stat(s, 11.833, 2.077, '+287', 'Advanced', { capColor: C.white });
  iconTiles(s, 7.461, 3.173, { activeIcon: C.green });

  rect(s, 6.311, 4.779, 2.293, 1.605, C.white, { rectRadius: 0.28, shadow: SH_SOFT() });
  rect(s, 6.633, 4.996, 0.372, 0.372, C.green, { rectRadius: 0.113, shadow: SH_TINY() });
  iconLeaf(s, 6.745, 5.108, 0.148, C.white, C.green);
  txt(s, 'Monitoring System', { x: 6.544, y: 5.492, w: 2.059, h: 0.337, fontSize: 14, bold: true, fontFace: F_HEAD });
  bodyText(s, 6.544, 5.777, 2.624, 0.37, [
    { text: 'Automated', options: { bold: true } }, { text: ' Reporting' },
  ]);
  bodyText(s, 9.035, 4.936, 3.254, 1.279, LOREM_MED);
  stat(s, -0.089, 5.18, '+126', 'Monitoring', { align: 'right', valueColor: C.white, capColor: C.white });
  pageNum(s, 8);
}

function slide09(s) {
  chrome(s, CHROME_LIGHT);
  diamond(s, 13.733, 4.054, 4.538, C.green, { rot: 45, radius: 0.41 });
  heading(s, 1.236, 1.797, 4.691, [
    [['Reducing', C.dark]],
    [['Waste', C.greenMid], [' and Costs', C.dark]],
  ]);
  subtitle(s, 1.236, 3.198, 'Managing Massive Information');
  bodyText(s, 1.236, 3.813, 5.035, 0.976, LOREM_FULL);
  iconTiles(s, 1.236, 5.249, { activeIcon: C.green });
  statBig(s, 4.554, 5.209, '235+', [
    { text: 'Intelligent', options: { bold: true } }, { text: ' Data' },
  ]);
  leafDisc(s, 10.12, 3.59);
  stat(s, 11.159, 3.593, '+287', 'Strategic', { valueColor: C.white, capColor: C.white, capBold: true });
  pageNum(s, 9);
}

function slide10(s) {
  for (let i = 0; i < 5; i++) diamond(s, 0.42 + i * 3.18, 0.33, 2.008, C.dark, { rot: 45, radius: 0.36 });
  heading(s, 0.927, 5.571, 6.347, [[['Let\u2019s ', C.dark], ['Meet', C.greenMid], [' Our Team', C.dark]]]);
  subtitle(s, 0.927, 6.33, 'Driving Smarter Decisions with Data Power');

  const team = [
    { x: 0.78, cx: 0.895, cw: 2.205, name: 'Rachael Lee', fill: C.dark, color: C.green },
    { x: 3.97, cx: 3.97, cw: 2.534, name: 'Rebeca Naseer', fill: C.white, color: C.dark },
    { x: 7.102, cx: 7.123, cw: 2.534, name: 'Damian Maximo', fill: C.white, color: C.dark },
    { x: 10.28, cx: 10.33, cw: 2.534, name: 'Donnie Albert', fill: C.white, color: C.dark },
  ];
  team.forEach((m) => {
    pillLabel(s, m.x, 3.008, 2.484, [{ text: m.name, options: { bold: true } }], { fill: m.fill, color: m.color });
    txt(s, 'Enterprise Analytics empowers.', {
      x: m.cx, y: 3.622, w: m.cw, h: 0.673, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5,
    });
  });

  [{ x: 9.14, v: '+21', c: 'Support', col: C.dark }, { x: 10.701, v: '97%', c: 'Strategic', col: C.greenMid }].forEach((k) => {
    rect(s, k.x, 5.69, 1.432, 0.961, C.white, { rectRadius: 0.16, shadow: SH_SOFT() });
    txt(s, k.v, { x: k.x + 0.051, y: 5.69, w: 1.33, h: 0.64, fontSize: 32, bold: true, fontFace: F_HEAD, color: k.col, align: 'center' });
    txt(s, k.c, { x: k.x + 0.051, y: 6.185, w: 1.33, h: 0.376, fontSize: 12, bold: true, align: 'center', lineSpacingMultiple: 1.5 });
  });
  pageNum(s, 10);
}

function slide11(s) {
  chrome(s, CHROME_LIGHT);
  heading(s, 0.897, 1.347, 3.382, [
    [['Identifying ', C.dark]], [['Hidden', C.greenMid]], [[' Threats', C.dark]],
  ], { h: 1.919 });
  bodyText(s, 0.897, 3.403, 2.593, 0.769, 'Personalizing Shopping Experiences', { fontSize: 14 });
  progressCard(s, 0.598, 5.338, [
    { text: 'KPI ' }, { text: 'Monitoring', options: { bold: true } }, { text: ' System' },
  ], 0.5, '60%');
  progressCard(s, 8.126, 2.598, [
    { text: 'Enterprise', options: { bold: true } }, { text: ' Analytics Platform' },
  ], 0.72, '80%');
  leafBadge(s, 8.292, 4.657);
  txt(s, 'Risk Management', { x: 8.319, y: 5.198, w: 2.059, h: 0.337, fontSize: 14, bold: true, fontFace: F_HEAD });
  bodyText(s, 8.32, 5.543, 4.163, 0.976, LOREM_MED);
  pageNum(s, 11);
}

function slide12(s) {
  chrome(s, CHROME_LIGHT);
  rect(s, 1.014, 5.489, 1.564, 1.136, C.white, { rectRadius: 0.24, shadow: SH_SOFT() });
  heading(s, 8.322, 1.788, 4.691, [[['Streamlining ', C.dark], ['Supply', C.greenMid], [' Chains', C.dark]]]);
  subtitle(s, 8.322, 3.19, 'Case Study: Manufacturing', { w: 3.622 });
  bodyText(s, 8.322, 3.805, 4.384, 0.976,
    'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
    'By integrating advanced analytics across.');

  oval(s, 2.54, 2.134, 0.807, C.green);
  iconBank(s, 2.794, 2.389, 0.298, C.white);
  stat(s, 0.606, 2.1, '+234', 'Data Pipeline', { align: 'right', valueColor: C.dark });
  oval(s, 2.54, 3.636, 0.807, C.greenMid);
  iconCompass(s, 2.794, 3.891, 0.298, C.white);
  stat(s, 0.606, 3.602, '+87', 'Learning', { align: 'right', valueColor: C.dark });

  bodyText(s, 2.794, 5.72, 4.93, 0.673, LOREM_MED);
  iconTiles(s, 8.322, 5.647, { activeIcon: C.green });
  statBig(s, 1.114, 5.556, '235+', 'Cloud Analytics', { capW: 1.395 });
  pageNum(s, 12);
}

function slide13(s) {
  chrome(s, CHROME_LIGHT);
  diagWash(s, 0, 5.434, 13.333, 2.066, 0.9, 0.7, true);   // dark footer band, lighter to the right
  heading(s, 4.321, 4.437, 4.691, [[['Data Security', C.dark]]], { align: 'center' });
  txt(s, 'Safeguarding Information', { x: 4.006, y: 5.717, w: 5.321, h: 0.337, fontSize: 14, color: C.green, align: 'center' });

  txt(s, 'Machine Learning', { x: 10.266, y: 1.694, w: 2.094, h: 0.37, fontSize: 16, color: C.ink, fontFace: F_HEAD });
  bodyText(s, 10.266, 2.043, 2.78, 0.675, 'Quis nostrum exercitationem ullam corporis suscipit');
  txt(s, 'Smart Decisions', { x: 0.998, y: 1.634, w: 2.094, h: 0.37, fontSize: 16, color: C.ink, fontFace: F_HEAD, align: 'right' });
  bodyText(s, 0.312, 1.984, 2.78, 0.675, 'Quis nostrum exercitationem ullam corporis suscipit', { align: 'right' });

  diamond(s, 10.75, 3.701, 0.904, C.green, { radius: 0.16 });
  diamond(s, 2.766, 3.701, 0.904, C.gray, { radius: 0.16 });
  leafBadge(s, 8.847, 1.634);
  leafBadge(s, 4.175, 1.634);
  txt(s, 'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
    'By integrating advanced analytics across departments, businesses can make better decisions, ' +
    'enhance efficiency, and unlock hidden opportunities.',
  { x: 1.789, y: 6.11, w: 9.755, h: 0.673, fontSize: 12, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  pageNum(s, 13);
  oval(s, 6.342, 3.128, 0.807, C.green);
  iconBank(s, 6.596, 3.383, 0.298, C.white);
}

function slide14(s) {
  chrome(s, CHROME_LIGHT);
  heading(s, 1.046, 5.021, 3.106, [
    [['Challenges', C.greenMid]], [['in Analytics', C.dark]],
  ], { h: 1.313 });
  subtitle(s, 1.046, 6.372, 'Barriers to Overcome');
  bodyText(s, 6.274, 6.035, 6.3, 0.673,
    'Enterprise Analytics empowers organizations to transform data into actionable intelligence. ' +
    'By integrating advanced analytics across departments.');
  iconTiles(s, 6.274, 4.986, { activeIcon: C.green });
  statBig(s, 9.592, 4.947, '235+', [
    { text: 'Digital', options: { bold: true } }, { text: ' Billing' },
  ]);

  // KPI panel
  rect(s, 0.76, 1.51, 3.554, 2.787, C.dark, { rectRadius: 0.375 });
  const kpis = [
    { label: 'Average Revenue Per Unit', value: '570.00$', fill: C.green, lx: 1.186, lw: 2.701, ly: 1.727, by: 2.041 },
    { label: 'Customer Lifetime Value', value: '429.00$', fill: C.greenMid, lx: 1.245, lw: 2.583, ly: 2.56, by: 2.875 },
    { label: 'Customer Acquisition Cost', value: '483.00$', fill: C.olive, lx: 1.136, lw: 2.802, ly: 3.393, by: 3.708 },
  ];
  kpis.forEach((k) => {
    txt(s, k.label, { x: k.lx, y: k.ly - 0.03, w: k.lw, h: 0.26, fontSize: 12, bold: true, fontFace: F_CARD, color: C.white, align: 'center', margin: 0 });
    rect(s, 1.875, k.by, 1.323, 0.297, k.fill, { rectRadius: 0.148, shadow: SH_CHIP() });
    txt(s, k.value, { x: 1.875, y: k.by, w: 1.323, h: 0.297, fontSize: 11, bold: true, fontFace: F_CARD, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  });

  // Chart panel
  rect(s, 4.519, 1.51, 8.314, 2.787, C.white, { rectRadius: 0.308, shadow: SH_SOFT() });
  txt(s, 'Intelligent Enterprise Data Insights', { x: 5.42, y: 1.652, w: 6.576, h: 0.3, fontSize: 14, color: C.black, align: 'center', margin: 0 });
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  s.addChart('bar', [
    { name: 'Column1', labels: months, values: [4.3, 2.5, 3.5, 4.5, 4.1, 4.26, 4.42, 3.5, 6, 3, 7, 5] },
    { name: 'Column2', labels: months, values: [2.4, 4.4, 1.8, 2.8, 2.5, 2.36, 2.22, 1.8, 3, 7, 5.5, 3] },
    { name: 'Column3', labels: months, values: [2, 2, 1, 2, 1, 2, 4, 1, 3, 5.5, 3, 4] },
  ], {
    x: 4.798, y: 2.153, w: 7.775, h: 1.957,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 219, barOverlapPct: 100,
    chartColors: [C.green, C.greenMid, C.greenDark],
    showLegend: false, showTitle: false,
    valAxisMinVal: 0, valAxisMaxVal: 20, valAxisMajorUnit: 5,
    valAxisLineShow: false, valGridLine: { style: 'solid', size: 0.75, color: 'E8E8E8' },
    catAxisLineShow: true, catAxisLineColor: 'D9D9D9',
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    catAxisLabelFontFace: 'Be Vietnam Pro Light', valAxisLabelFontFace: 'Be Vietnam Pro Light',
    catAxisLabelColor: C.dark, valAxisLabelColor: C.dark,
  });
  pageNum(s, 14);
}

function slide15(s) {
  chrome(s, CHROME_LIGHT);
  diamond(s, 6.508, 2.888, 1.383, C.dark, { rot: 45, radius: 0.33 });
  imagePlaceholder(s, 2.03, 1.44, 1.48, 3.55, { fill: '8A8A8C', radius: 0.3, label: '', rotate: 9 });  // phone, back view
  imagePlaceholder(s, 3.90, 1.50, 2.40, 5.18, { fill: '2E2E2E', radius: 0.42, label: '' });            // phone bezel
  imagePlaceholder(s, 3.98, 1.58, 2.24, 5.02, { fill: 'FDFDFD', radius: 0.36, label: '', text: C.gray });
  oval(s, 5.03, 1.73, 0.06, '2E2E2E');                                                                 // front camera
  diamond(s, 3.862, 0.707, 0.37, C.dark, { rot: 45, radius: 0.09 });
  diamond(s, 2.166, 3.53, 1.263, C.green, { rot: 45, radius: 0.3 });
  heading(s, 8.197, 1.964, 4.691, [
    [['Building', C.dark]], [['Data-First', C.greenMid]], [['Mindsets', C.dark]],
  ], { h: 1.919 });
  subtitle(s, 8.197, 3.956, 'Cultural Transformation', { w: 3.622 });
  bodyText(s, 10.37, 4.774, 2.085, 0.976, 'Enterprise Analytics empowers organizations to transform data.');
  oval(s, 3.236, 5.785, 0.807, C.greenMid);
  iconCompass(s, 3.49, 6.04, 0.298, C.white);
  stat(s, 1.303, 5.75, '+287', [
    { text: 'Automated', options: { bold: true } }, { text: ' Data' },
  ], { align: 'right', valueColor: C.dark });
  pillLabel(s, 1.816, 3.28, 2.484, [
    { text: 'Big Data ' }, { text: 'Processing', options: { bold: true } }, { text: ' Platform' },
  ]);
  progressCard(s, 5.595, 4.795, 'Real-Time Data analytic', 0.72, '80%');
  pageNum(s, 15);
}

function slide16(s) {
  chrome(s, CHROME_LIGHT);
  diamond(s, 11.43, 5.468, 2.508, { color: C.greenDeep, transparency: 88 });
  txt(s, 'Take a moment. Insights are just beginning.', { x: 9.279, y: 5.206, w: 2.9, h: 0.64, fontSize: 16 });
  txt(s, 'Slide', { x: 5.989, y: 3.309, w: 4.691, h: 2.036, fontSize: 115, bold: true, fontFace: F_HEAD, color: C.dark });
  txt(s, 'Break.', { x: 2.257, y: 1.528, w: 6.332, h: 2.423, fontSize: 138, bold: true, fontFace: F_HEAD, color: C.greenMid });
  txt(s, '45min', { x: 7.096, y: 5.13, w: 2.079, h: 0.774, fontSize: 40, bold: true, fontFace: F_HEAD, color: C.greenMid, align: 'right' });
  diamond(s, 2.39, 2.586, 2.969, { color: C.greenDeep, transparency: 88 });
  diamond(s, 12.087, 2.435, 1.126, C.green, { rot: 45, radius: 0.101 });
  diamond(s, 13.334, 2.954, 1.126, C.gray, { rot: 45, radius: 0.101 });
  diamond(s, 1.282, 5.361, 1.126, C.green, { rot: 45, radius: 0.101 });
  diamond(s, 0.035, 4.842, 1.126, C.gray, { rot: 45, radius: 0.101 });
  txt(s, [
    { text: QUOTE + '\u2013 ' },
    { text: DEMING, options: { bold: true, italic: true } },
  ], { x: 2.39, y: 3.87, w: 3.421, h: 0.909, fontSize: 16, align: 'right' });
  pageNum(s, 16);
}

function slide17(s) {
  chrome(s, CHROME_LIGHT);
  imagePlaceholder(s, 6.56, 2.15, 5.39, 4.18, { fill: '2E2E2E', radius: 0.34, label: '' });      // tablet bezel
  imagePlaceholder(s, 6.76, 2.35, 4.99, 3.78, { fill: 'FDFDFD', radius: 0.14, label: '[tablet mockup]', text: C.gray });
  imagePlaceholder(s, 7.95, 6.35, 3.25, 0.2, { fill: 'D8D8D8', radius: 0.1, label: '', rotate: 353 });   // stylus
  diamond(s, 11.897, 2.278, 1.148, C.dark, { rot: 45, radius: 0.27 });
  progressCard(s, 2.961, 3.841, [
    { text: 'Enterprise', options: { bold: true } }, { text: ' Data Driven' },
  ], 0.72, '80%');
  progressCard(s, 1.171, 5.221, [
    { text: 'trategic Data ' }, { text: 'Governance', options: { bold: true } },
  ], 0.72, '80%');
  heading(s, 1.213, 1.777, 4.718, [
    [['Challenges', C.greenMid]], [['in Analytics', C.dark]],
  ], { h: 1.313 });
  subtitle(s, 1.213, 3.127, 'Barriers to Overcome');
  stat(s, 1.287, 3.883, '+287', 'Data Insights', { valueColor: C.dark });
  pageNum(s, 17);
}

function slide18(s) {
  s.background = { color: '2C2C2C' };
  diagWash(s, 0, 0, 13.333, 7.5, 0.9, 0.7);          // charcoal gradient, lighter to the bottom-right
  diamond(s, 12.886, 2.292, 1.496, { color: C.greenDeep, transparency: 70 });
  diamond(s, 0.535, 5.743, 1.496, { color: C.greenDeep, transparency: 70 });

  // header row
  const headers = [
    { x: 0.8, w: 2.558, label: 'Customer', fill: C.white, color: C.gray, tx: 0.935 },
    { x: 3.55, w: 2.462, label: 'Project', fill: C.white, color: C.gray, tx: 3.751 },
    { x: 6.182, w: 1.812, label: 'Rating', fill: C.white, color: C.gray, tx: 6.382 },
    { x: 8.164, w: 4.217, label: 'feedback', fill: C.greenMid, color: C.white, tx: 8.279 },
  ];
  headers.forEach((h) => {
    rect(s, h.x, 2.978, h.w, 0.48, h.fill, { rectRadius: 0.084, shadow: SH_WIDE() });
    txt(s, h.label, { x: h.tx, y: 3.045, w: 2.722, h: 0.337, fontSize: 14, fontFace: F_HEAD, color: h.color, valign: 'middle' });
  });

  // testimonial rows
  const rows = [
    { y: 3.86, name: 'Robert Luis', sel: false },
    { y: 4.63, name: 'Aaron Loeb', sel: false },
    { y: 5.422, name: 'Samira Hadid', sel: true },
    { y: 6.215, name: 'Amanda Wong', sel: false },
  ];
  rows.forEach((r) => {
    const fg = r.sel ? C.green : C.white;
    rect(s, 0.8, r.y, 11.58, 0.611, C.panel, {
      rectRadius: 0.305,
      line: r.sel ? { color: C.green, width: 1 } : NO_LINE,
      shadow: r.sel ? SH_WIDE() : null,
    });
    txt(s, r.name, { x: 0.935, y: r.y + 0.154, w: 2.722, h: 0.303, fontSize: 12, bold: r.sel, color: fg });
    txt(s, 'Option project', { x: 3.745, y: r.y + 0.154, w: 2.722, h: 0.303, fontSize: 12, bold: r.sel, color: fg });
    if (r.sel) stars(s, 6.34, r.y + 0.218, { w: 0.198, h: 0.176, gap: 0.24 });
    else stars(s, 6.473, r.y + 0.238);
    txt(s, LOREM_SHORT, { x: 8.279, y: r.y + 0.07, w: 4.146, h: 0.471, fontSize: 11, color: fg });
  });

  heading(s, 0.8, 1.865, 5.636, [[['Testimonial', C.greenMid], [' ', C.dark], ['Data', C.white]]]);
  rect(s, 9.64, 1.815, 0.821, 0.821, C.white, { rectRadius: 0.2, shadow: SH_SOFT() });
  iconChart(s, 9.926, 2.1, C.dark);
  statBig(s, 10.883, 1.775, '235+', 'Data Insights', { capColor: 'F2F2F2' });
  chrome(s, CHROME_DARK);
  pageNum(s, 18);
}

function slide19(s) {
  s.background = { color: '2C2C2C' };
  diagWash(s, 0, 0, 13.333, 7.5, 0.9, 0.7);          // charcoal gradient, lighter to the bottom-right
  diamond(s, 12.054, 2.954, 1.781, { color: C.greenDeep, transparency: 70 });
  diamond(s, 0.203, 6.611, 1.876, { color: C.greenDeep, transparency: 70 });
  txt(s, [
    { text: 'Contact', options: { bold: true, color: C.green } },
    { text: ' Slide', options: { bold: true, color: C.white } },
  ], { x: 7.566, y: 2.583, w: 5.321, h: 0.438, fontSize: 20 });
  chrome(s, CHROME_DARK);

  rect(s, -1.114, 1.872, 7.367, 4.728, C.panel, { rectRadius: 0.825 });
  const contacts = [
    { y: 2.523, icon: 'mail', text: 'loremipsum@mail.com', color: C.green },
    { y: 3.456, icon: 'phone', text: '+0123 4567 890 000', color: C.white, bar: true },
    { y: 4.41, icon: 'globe', text: 'www.yourwebsite.com', color: C.green },
  ];
  contacts.forEach((c) => {
    if (c.bar) rect(s, 1.53, 3.297, 4.037, 0.708, C.green, { rectRadius: 0.15 });
    const ic = c.bar ? C.white : C.green;
    if (c.icon === 'mail') {
      s.addShape('rect', { x: 1.843, y: 2.61, w: 0.242, h: 0.195, fill: { color: ic }, line: NO_LINE });
      s.addShape('triangle', { x: 1.861, y: 2.632, w: 0.206, h: 0.13, fill: { color: C.panel }, line: NO_LINE, rotate: 180 });
    } else if (c.icon === 'phone') {
      s.addShape('roundRect', { x: 1.858, y: 3.534, w: 0.212, h: 0.212, fill: { color: ic }, line: NO_LINE, rectRadius: 0.075, rotate: 20 });
    } else {
      s.addShape('ellipse', { x: 1.858, y: 4.489, w: 0.212, h: 0.212, fill: { type: 'none' }, line: { color: ic, width: 1.5 } });
      s.addShape('ellipse', { x: 1.925, y: 4.489, w: 0.078, h: 0.212, fill: { type: 'none' }, line: { color: ic, width: 1 } });
    }
    txt(s, c.text, { x: 2.236, y: c.y, w: 3.204, h: 0.37, fontSize: 16, color: c.color, valign: 'middle' });
  });

  heading(s, 7.566, 3.158, 4.061, [
    [['Enterprise ', C.white]], [['Analytics', C.greenMid]], [['Solutions', C.white]],
  ], { size: 48, h: 2.524 });
  txt(s, [
    { text: QUOTE, options: { color: C.white } },
    { text: '\u2013 ', options: { color: C.green } },
    { text: DEMING, options: { bold: true, italic: true, color: C.green } },
  ], { x: 1.355, y: 5.37, w: 4.556, h: 0.64, fontSize: 16 });
  diamond(s, 13.141, 4.579, 1.478, { color: C.greenDeep, transparency: 70 });
  pageNum(s, 19);
}

function slide20(s) {
  s.background = { color: '2C2C2C' };
  pageNum(s, 20);                                    // master pager, dimmed by the wash above it
  diagWash(s, 0, 0, 13.333, 7.5, 0.9, 0.7);          // charcoal gradient, lighter to the bottom-right
  diamond(s, 11.217, 5.46, 2.508, { color: C.greenDeep, transparency: 88 });
  txt(s, 'You!', { x: 6.667, y: 4.13, w: 4.691, h: 2.036, fontSize: 115, bold: true, fontFace: F_HEAD, color: C.white });
  txt(s, 'Thank', { x: 2.935, y: 2.349, w: 6.332, h: 2.423, fontSize: 138, bold: true, fontFace: F_HEAD, color: C.greenMid });
  diamond(s, 2.39, 2.777, 2.969, { color: C.greenDeep, transparency: 88 });
  diamond(s, 13.334, 2.954, 1.126, C.gray, { rot: 45, radius: 0.101 });
  diamond(s, 0.036, 4.843, 1.126, C.gray, { rot: 45, radius: 0.101 });
  txt(s, 'Thank you for exploring Enterprise Analytics with us. Let\u2019s build a smarter future together.', {
    x: 3.068, y: 4.772, w: 3.421, h: 0.976, fontSize: 12, color: C.white, align: 'right', lineSpacingMultiple: 1.5,
  });
  chrome(s, CHROME_DARK);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'W16x9';
  pptx.theme = { headFontFace: F_HEAD, bodyFontFace: F_BODY };
  BUILDERS.forEach((fn) => fn(pptx.addSlide()));
  return pptx.writeFile({
    fileName: path.join(__dirname, '08245316-ad10-4892-b91d-8b657be24cb6_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
