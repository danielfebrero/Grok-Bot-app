/**
 * SWOT & Strategy Business Planning  --  20 slide deck, rebuilt with pptxgenjs.
 *
 * Run:  node 022902c4-cc7b-48fa-acaf-632b7b3ee216_grok_final.js
 * Out:  022902c4-cc7b-48fa-acaf-632b7b3ee216_grok_final.pptx  (next to this file)
 *
 * Raster art from the source deck is replaced by programmatic placeholders
 * (see `imagePlaceholder` / `iconGlyph`); everything else is native shapes.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme: palette, typography, reusable effects
 * ------------------------------------------------------------------ */

const NAVY      = '1B232E'; // theme accent1 -- the deck's signature colour
const NAVY_LT   = '2B384A'; // accent1 lumMod 90%
const NAVY_DK   = '141A22'; // accent1 lumMod 75%
const NAVY_SOFT = '435773'; // accent1 lumMod 75% / lumOff 25%
const ICE       = 'E3E8EF'; // accent1 lumMod 10% / lumOff 90%
const ICE_DEEP  = 'C8D2DF'; // accent1 lumMod 20% / lumOff 80%
const WHITE     = 'FFFFFF';
const BLACK     = '000000';
const GREY_95   = 'F2F2F2';
const GREY_85   = 'D9D9D9';
const GREY_25   = '404040';
const GREY_35   = '595959';
const GREY_50   = '808080';
const GREY_DK   = '3B3838';
const LT2       = 'E7E6E6'; // theme lt2
const SLATE     = '44546A'; // theme tx2
const SLATE_MID = '8497B0';
const SLATE_DK  = '333F50';
const GREEN     = '70AD47'; // theme accent6, one slice of the slide-15 gauge

const HEAD     = 'Montserrat';       // theme major latin font
const HEAD_MED = 'Montserrat Medium'; // used for a few captions
const BODY     = 'Poppins Light';    // theme minor latin font

const DECK_TITLE = 'SWOT & Strategy Business Planning';
const SW = 40 / 3; // 13.333in -- exactly 12192000 EMU, as in the source deck
const SH = 7.5;

// Drop shadows used across the deck (blur / offset in pt, angle in deg).
// pptxgenjs rewrites shadow objects in place while writing XML, so each shape
// needs its own copy -- hence a factory rather than shared constants.
const SHADOW_PRESETS = {
  card:   { blur: 18, offset: 3,  angle: 90, color: BLACK,   opacity: 0.11 },
  soft:   { blur: 35, offset: 10, angle: 50, color: BLACK,   opacity: 0.05 },
  gear:   { blur: 30, offset: 10, angle: 50, color: BLACK,   opacity: 0.15 },
  panel:  { blur: 35, offset: 10, angle: 50, color: BLACK,   opacity: 0.16 },
  chart:  { blur: 35, offset: 10, angle: 50, color: BLACK,   opacity: 0.10 },
  ring:   { blur: 15, offset: 3,  angle: 90, color: GREY_50, opacity: 0.22 },
  desk:   { blur: 34, offset: 12, angle: 50, color: BLACK,   opacity: 0.05 },
  bubble: { blur: 18, offset: 6,  angle: 45, color: BLACK,   opacity: 0.12 },
};
function shadow(name) {
  return Object.assign({ type: 'outer' }, SHADOW_PRESETS[name]);
}

/* ------------------------------------------------------------------ *
 * Small drawing helpers
 * ------------------------------------------------------------------ */

// Text box: OOXML defaults (top anchored, minor font, 12pt, black).
function txt(slide, content, o) {
  slide.addText(content, Object.assign({ fontFace: BODY, fontSize: 12, color: BLACK, valign: 'top' }, o));
}

// Heading text (theme major font).
function head(slide, content, o) {
  txt(slide, content, Object.assign({ fontFace: HEAD }, o));
}

function rect(slide, o)      { slide.addShape('rect', o); }
function roundRect(slide, o) { slide.addShape('roundRect', o); }
function ellipse(slide, o)   { slide.addShape('ellipse', o); }
function line(slide, o)      { slide.addShape('line', o); }

function mix(a, b, t) {
  const p = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map(i => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

// pptxgenjs has no gradient fill, so smooth ramps are painted as thin bands.
function gradientBands(slide, o) {
  const steps = o.steps || 48;
  for (let i = 0; i < steps; i++) {
    const t = (i + 0.5) / steps;
    const box = o.dir === 'h'
      ? { x: o.x + (o.w * i) / steps, y: o.y, w: o.w / steps + 0.01, h: o.h }
      : { x: o.x, y: o.y + (o.h * i) / steps, w: o.w, h: o.h / steps + 0.01 };
    rect(slide, Object.assign(box, { fill: { color: o.colorAt(t) } }));
  }
}

// Free-form path described with fractions of the shape box: ['M'|'L', fx, fy],
// ['C', fx1, fy1, fx2, fy2, fx, fy] or ['Z'].
function fracPath(slide, o) {
  const points = [];
  o.path.forEach(seg => {
    const [k] = seg;
    const fx = (i) => o.w * seg[i], fy = (i) => o.h * seg[i];
    if (k === 'M') points.push({ x: fx(1), y: fy(2), moveTo: true });
    else if (k === 'L') points.push({ x: fx(1), y: fy(2) });
    else if (k === 'C') points.push({ x: fx(5), y: fy(6), curve: { type: 'cubic', x1: fx(1), y1: fy(2), x2: fx(3), y2: fy(4) } });
    else points.push({ close: true });
  });
  slide.addShape('custGeom', Object.assign({}, o, { points: points, path: undefined }));
}

// Rectangle with only the two top corners rounded (round2SameRect look-alike).
function topRoundedRect(slide, o) {
  const { w, h, r } = o;
  slide.addShape('custGeom', Object.assign({}, o, {
    r: undefined,
    points: [
      { x: 0, y: h, moveTo: true },
      { x: 0, y: r },
      { curve: { type: 'arc', wR: r, hR: r, stAng: 180, swAng: 90 } },
      { x: w - r, y: 0 },
      { curve: { type: 'arc', wR: r, hR: r, stAng: 270, swAng: 90 } },
      { x: w, y: h },
      { close: true },
    ],
  }));
}

// Folder-style tab: rectangle whose top-right corner is cut on the diagonal.
const TAB_PATH = [['M', 0, 0], ['L', 0.8136, 0], ['L', 1, 0.996], ['L', 1, 1], ['L', 0, 1], ['Z']];

/**
 * Cog wheel. `teeth` trapezoidal teeth whose tips reach `tipR` and whose roots
 * sit at `rootR` (both as a fraction of the box half-width); `phase` is the
 * angle of the first tooth centre. Tip and root arcs each span a fixed share
 * of one tooth period, matching the proportions of the source artwork.
 */
function gear(slide, o) {
  const n = o.teeth, half = o.w / 2, cx = half, cy = o.h / 2;
  const tip = half * o.tipR, root = half * o.rootR;
  const period = (2 * Math.PI) / n, phase = ((o.phase || 0) * Math.PI) / 180;
  const pts = [];
  for (let i = 0; i < n; i++) {
    const c = phase + period * i;
    [[c - period * 0.11, tip], [c + period * 0.11, tip], [c + period * 0.25, root], [c + period * 0.75, root]]
      .forEach(([ang, rad], k) => pts.push({
        x: cx + rad * Math.cos(ang), y: cy + rad * Math.sin(ang),
        moveTo: i === 0 && k === 0 ? true : undefined,
      }));
  }
  pts.push({ close: true });
  slide.addShape('custGeom', Object.assign({}, o, { teeth: undefined, tipR: undefined, rootR: undefined, phase: undefined, points: pts }));
}

/**
 * Stand-in for the small vector pictograms of the original deck.
 * kind 0 = presentation board, 1 = colour wheel, 2 = target, 3 = bar chart.
 */
function iconGlyph(slide, x, y, s, color, kind) {
  const k = (kind || 0) % 4;
  if (k === 0) {
    slide.addShape('roundRect', { x: x + s * 0.04, y: y + s * 0.10, w: s * 0.92, h: s * 0.56, rectRadius: s * 0.06, fill: { type: 'none' }, line: { color: color, width: 1.5 } });
    [[0.22, 0.30], [0.44, 0.20], [0.66, 0.36]].forEach(([bx, by]) =>
      rect(slide, { x: x + s * bx, y: y + s * by, w: s * 0.12, h: s * (0.66 - by), fill: { color: color } }));
    line(slide, { x: x + s * 0.5, y: y + s * 0.66, w: 0, h: s * 0.24, line: { color: color, width: 1.5 } });
    line(slide, { x: x + s * 0.28, y: y + s * 0.90, w: s * 0.22, h: -s * 0.24, line: { color: color, width: 1.5 } });
    line(slide, { x: x + s * 0.50, y: y + s * 0.66, w: s * 0.22, h: s * 0.24, line: { color: color, width: 1.5 } });
  } else if (k === 1) {
    [[0.30, 0.10], [0.10, 0.46], [0.50, 0.46]].forEach(([cx, cy]) =>
      ellipse(slide, { x: x + s * cx, y: y + s * cy, w: s * 0.40, h: s * 0.40, fill: { color: color } }));
  } else if (k === 2) {
    [0.06, 0.24, 0.40].forEach(inset =>
      ellipse(slide, { x: x + s * inset, y: y + s * inset, w: s * (1 - 2 * inset), h: s * (1 - 2 * inset), fill: { type: 'none' }, line: { color: color, width: 1.5 } }));
    line(slide, { x: x + s * 0.5, y: y + s * 0.5, w: s * 0.46, h: -s * 0.46, line: { color: color, width: 1.5, endArrowType: 'arrow' } });
  } else {
    [[0.16, 0.30], [0.38, 0.14], [0.60, 0.44]].forEach(([bx, by]) =>
      rect(slide, { x: x + s * bx, y: y + s * by, w: s * 0.14, h: s * (0.82 - by), fill: { color: color } }));
    line(slide, { x: x + s * 0.84, y: y + s * 0.10, w: 0, h: s * 0.80, line: { color: color, width: 1.5 } });
  }
}

// Stand-in for the raster illustrations of the original deck.
function imagePlaceholder(slide, o) {
  roundRect(slide, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.12, fill: { color: o.fill || ICE }, line: { color: o.stroke || ICE_DEEP, width: 1, dashType: 'dash' } });
  txt(slide, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.16, w: o.w, h: 0.32, align: 'center', valign: 'middle', fontSize: 11, color: o.text || SLATE_MID });
}

// Running header + slide number that the slide master paints on every page.
function chrome(slide, num, o) {
  o = o || {};
  head(slide, DECK_TITLE, { x: 0.36, y: 0.34, w: 3.41, h: 0.286, fontSize: 11, color: o.titleColor || BLACK });
  head(slide, String(num), { x: 11.894, y: 0.34, w: 1.079, h: 0.286, fontSize: 11, bold: true, align: 'right', color: o.numberColor || BLACK });
}

// "Learn More >" pill button (dark or light variant).
function learnMore(slide, x, y, dark) {
  const fg = dark ? WHITE : NAVY, bg = dark ? NAVY : WHITE;
  roundRect(slide, { x: x, y: y, w: 1.739, h: 0.434, rectRadius: 0.058, fill: { color: bg } });
  head(slide, 'Learn More', { x: x + 0.126, y: y + 0.066, w: 1.14, h: 0.286, fontSize: 11, align: 'center', color: fg });
  ellipse(slide, { x: x + 1.31, y: y + 0.075, w: 0.285, h: 0.285, fill: { color: fg } });
  slide.addShape('chevron', { x: x + 1.421, y: y + 0.16, w: 0.071, h: 0.114, fill: { color: bg } });
}

/* ------------------------------------------------------------------ *
 * Shared content
 * ------------------------------------------------------------------ */

const LOREM_SHORT = 'Lorem sit ligula ipsum dolor amet sit amet, consectetuer amet';
const LOREM_CARD  = 'Lorem ligula ipsum dolor sit amet, consectetuer amet';
const LOREM_RING  = 'Lorem ligula ipsum dolor sit amet, consectetuer';
const LOREM_BAR   = 'Lorem ligula sit ipsum dolor';
const LOREM_LONG  = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus';
const LOREM_MED   = 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget ligula';
const LOREM_XL    = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. '
                  + 'Cum sociis natoque penatibus ipsum dolor sit amet, consectetuer adipiscing elit. Aenean ipsum dolor sit amet, '
                  + 'consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis';
const SWOT        = ['Strengths', 'Weaknesses', 'Opportunities', 'Threats'];

/* ------------------------------------------------------------------ *
 * Slide builders
 * ------------------------------------------------------------------ */

// 1 -- Cover: horizontal navy gradient with oversized title.
function slide01(s) {
  gradientBands(s, { x: 0, y: 0, w: SW, h: SH, dir: 'h', steps: 64, colorAt: t => mix(NAVY, NAVY_LT, t) });
  head(s, DECK_TITLE, { x: 0.959, y: 4.061, w: 10.931, h: 2.524, fontSize: 72, color: WHITE });
  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget',
    { x: 9.268, y: 0.878, w: 3.107, h: 0.87, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3 });
}

// 2 -- Section opener "Overview".
function slide02(s) {
  s.background = { color: NAVY };
  fracPath(s, { x: 0.727, y: 4.505, w: 3.444, h: 0.707, path: TAB_PATH, fill: { color: NAVY } });
  chrome(s, 2, { titleColor: WHITE, numberColor: WHITE });
  head(s, 'Overview', { x: 0.828, y: 4.695, w: 2.73, h: 0.707, fontSize: 36, color: WHITE });
  txt(s, LOREM_XL, { x: 0.727, y: 5.768, w: 9.239, h: 0.87, color: WHITE, lineSpacingMultiple: 1.3 });
  learnMore(s, 10.324, 5.986, false);
}

// 3 & 4 share the same four SWOT cards, only the vertical offsets differ.
const SWOT_CARDS = [
  { x: 1.075, icon: 1.305, iconSize: 0.691, dark: true },
  { x: 3.904, icon: 4.146, iconSize: 0.666, dark: false },
  { x: 6.732, icon: 7.007, iconSize: 0.599, dark: true },
  { x: 9.560, icon: 9.856, iconSize: 0.558, dark: false },
];

function swotCard(s, c, i, top) {
  const fg = c.dark ? WHITE : BLACK;
  roundRect(s, { x: c.x, y: top, w: 2.698, h: 2.556, rectRadius: 0.113, fill: { color: c.dark ? NAVY : WHITE }, shadow: shadow('card') });
  iconGlyph(s, c.icon, top + 0.221, c.iconSize, fg, i);
  head(s, SWOT[i], { x: c.icon, y: top + 1.116, w: 1.7, h: 0.337, fontSize: 14, color: fg });
  txt(s, LOREM_SHORT, { x: c.icon, y: top + 1.453, w: 2.221, h: 0.87, color: fg, align: 'justify', lineSpacingMultiple: 1.3 });
}

// 3 -- SWOT cards on a "SWOT" watermark.
function slide03(s) {
  chrome(s, 3);
  head(s, 'Key Strengths Driving Our Competitive Advantage', { x: 2.107, y: 1.117, w: 9.12, h: 1.582, fontSize: 44, align: 'center' });
  [['S', 1.824], ['W', 4.556], ['O', 7.384], ['T', 10.212]].forEach(([ch, x]) =>
    head(s, ch, { x: x, y: 2.935, w: 1.394, h: 2.036, fontSize: 115, bold: true, align: 'center' }));
  SWOT_CARDS.forEach((c, i) => swotCard(s, c, i, 4.06));
}

// 4 -- Same cards, staggered, joined by curved arrows.
function slide04(s) {
  chrome(s, 4);
  head(s, 'Identifying Internal Weaknesses to Improve Performance', { x: 1.371, y: 1.117, w: 10.591, h: 1.582, fontSize: 44, align: 'center' });
  SWOT_CARDS.forEach((c, i) => swotCard(s, c, i, i % 2 === 0 ? 3.414 : 4.3));
  const arc = { w: 0.981, h: 0.981, angleRange: [270, 0], fill: { type: 'none' }, line: { color: NAVY, width: 1.5, endArrowType: 'arrow' } };
  s.addShape('arc', Object.assign({ x: 3.681, y: 3.483 }, arc));
  s.addShape('arc', Object.assign({ x: 9.360, y: 3.483 }, arc));
  s.addShape('arc', Object.assign({ x: 6.393, y: 5.779, rotate: 176.9, flipH: true }, arc));
}

// 5 -- Four-petal SWOT wheel with labels on both sides.
const PETAL = [['M', 1, 0], ['L', 1, 1], ['L', 0, 1], ['L', 0.0183, 0.8154], ['C', 0.1117, 0.3501, 0.5157, 0, 1, 0], ['Z']];

function slide05(s) {
  s.background = { color: WHITE };
  chrome(s, 5);
  line(s, { x: 6.667, y: 1.218, w: 0, h: 5.387, line: { color: NAVY, width: 1.25 } });
  line(s, { x: 1.596, y: 3.912, w: 10.141, h: 0, line: { color: NAVY, width: 1.25 } });

  // petals: top-left, top-right, bottom-right, bottom-left
  [
    { x: 4.692, y: 1.978, fill: NAVY,  flipH: false, flipV: false },
    { x: 6.732, y: 1.978, fill: WHITE, flipH: true,  flipV: false },
    { x: 6.732, y: 3.972, fill: NAVY,  flipH: true,  flipV: true  },
    { x: 4.692, y: 3.972, fill: WHITE, flipH: false, flipV: true  },
  ].forEach(p => fracPath(s, { x: p.x, y: p.y, w: 1.909, h: 1.873, path: PETAL, fill: { color: p.fill }, flipH: p.flipH, flipV: p.flipV, shadow: shadow('soft') }));

  ellipse(s, { x: 5.631, y: 2.907, w: 2.072, h: 2.072, fill: { color: WHITE }, shadow: shadow('soft') });
  [['S ', 5.047, 2.528, WHITE], ['W ', 7.088, 2.528, BLACK], ['O ', 7.088, 4.522, WHITE], ['T ', 5.047, 4.522, BLACK]]
    .forEach(([ch, x, y, col]) => head(s, ch, { x: x, y: y, w: 1.198, h: 0.774, fontSize: 40, bold: true, align: 'center', color: col }));
  head(s, 'SWOT Analysis', { x: 5.924, y: 3.623, w: 1.486, h: 0.64, fontSize: 16, align: 'center', color: GREY_DK });

  // corner label blocks: icon tile, heading and body, mirrored left/right
  [
    { label: 'Strengths',     icon: 0, tile: 1.69,  dark: true,  tx: 1.591, align: 'justify', body: 'Lorem ligula ipsum dolor sit amet, sit consectetuer amet', ty: 1.808, w: 2.748 },
    { label: 'Weaknesses',    icon: 1, tile: 11.04, dark: false, tx: 8.994, align: 'right',   body: 'Lorem ligula ipsum dolor sit amet, sit consectetuer amet', ty: 1.808, w: 2.748 },
    { label: 'Threats',       icon: 3, tile: 1.69,  dark: false, tx: 1.591, align: 'justify', body: 'Lorem ligula ipsum dolor sit amet, sit consectetuer amet adipiscing elit. ', ty: 4.333, w: 3.456 },
    { label: 'Opportunities', icon: 2, tile: 11.04, dark: true,  tx: 8.439, align: 'right',   body: 'Lorem ligula ipsum dolor sit amet, sit consectetuer amet adipiscing elit. ', ty: 4.333, w: 3.304 },
  ].forEach(b => {
    const right = b.align === 'right';
    roundRect(s, { x: b.tile, y: b.ty, w: 0.604, h: 0.604, rectRadius: 0.101, fill: { color: b.dark ? NAVY : WHITE }, shadow: shadow('card') });
    iconGlyph(s, b.tile + 0.093, b.ty + 0.1, 0.42, b.dark ? WHITE : NAVY, b.icon);
    head(s, b.label, { x: right ? b.tx + b.w - 1.95 : b.tx, y: b.ty + 0.738, w: 1.95, h: 0.337, fontSize: 14, align: right ? 'right' : 'left' });
    txt(s, b.body, { x: b.tx, y: b.ty + 1.075, w: b.w, h: 0.608, align: b.align, lineSpacingMultiple: 1.3 });
  });
}

// 6 -- Segmented donut wheel with four satellite labels.
const INNER_RING = [['M', 0, 1], ['C', 0.0327, 0.2616, 0.6244, -0.0088, 1, 0.0002],
  ['C', 0.9987, 0.1524, 1.0008, -0.0308, 0.9996, 0.1214],
  ['C', 0.8178, 0.1039, 0.1333, 0.2514, 0.1064, 1], ['L', 0, 1], ['Z']];

function slide06(s) {
  chrome(s, 6);
  head(s, 'External Threats to Business Stability', { x: 0.788, y: 1.117, w: 11.758, h: 0.841, fontSize: 44, align: 'center' });

  const blockArc = { w: 3.939, h: 3.868, angleRange: [180, 269.9], arcThicknessRatio: 0.344, shadow: shadow('card') };
  s.addShape('blockArc', Object.assign({ x: 4.661, y: 2.633, rotate: 270, fill: { color: WHITE } }, blockArc));
  s.addShape('blockArc', Object.assign({ x: 4.661, y: 2.582, rotate: 90, flipV: true, fill: { color: NAVY } }, blockArc));
  s.addShape('blockArc', Object.assign({ x: 4.734, y: 2.633, rotate: 90, flipH: true, fill: { color: NAVY } }, blockArc));
  s.addShape('blockArc', Object.assign({ x: 4.734, y: 2.582, rotate: 270, flipH: true, flipV: true, fill: { color: WHITE } }, blockArc));

  [
    { x: 6.709, y: 3.072, w: 1.443, h: 1.447, rotate: 270, flipH: true, flipV: true, color: GREY_95 },
    { x: 5.182, y: 3.072, w: 1.443, h: 1.448, rotate: 90, flipV: true, color: NAVY_DK },
    { x: 6.703, y: 4.566, w: 1.447, h: 1.456, rotate: 90, flipH: true, color: NAVY_DK },
    { x: 5.178, y: 4.570, w: 1.449, h: 1.445, rotate: 270, color: GREY_95 },
  ].forEach(p => fracPath(s, Object.assign({}, p, { path: INNER_RING, fill: { color: p.color } })));

  ellipse(s, { x: 4.531, y: 2.406, w: 4.272, h: 4.272, fill: { type: 'none' }, line: { color: NAVY, width: 2 } });
  head(s, 'SWOT', { x: 5.846, y: 4.256, w: 1.641, h: 0.572, fontSize: 28, align: 'center', color: GREY_DK });

  // node bubbles sitting on the ring
  [
    { ch: 'S', x: 4.948, y: 2.773, dark: true,  tx: 5.064, ty: 2.888 },
    { ch: 'W', x: 7.747, y: 2.773, dark: false, tx: 7.835, ty: 2.888 },
    { ch: 'O', x: 4.948, y: 5.676, dark: false, tx: 5.075, ty: 5.790 },
    { ch: 'T', x: 7.747, y: 5.676, dark: true,  tx: 7.868, ty: 5.790 },
  ].forEach(n => {
    ellipse(s, { x: n.x - 0.035, y: n.y - 0.034, w: 0.741, h: 0.741, fill: { color: WHITE } });
    ellipse(s, { x: n.x, y: n.y, w: 0.671, h: 0.671, fill: { color: n.dark ? NAVY : WHITE }, line: n.dark ? undefined : { color: NAVY, width: 2 } });
    head(s, n.ch, { x: n.x, y: n.y + 0.16, w: 0.671, h: 0.4, fontSize: 18, bold: true, align: 'center', color: n.dark ? WHITE : BLACK });
  });

  // satellite labels + icon tiles
  [
    { label: 'Strengths ',     icon: 0, lx: 2.037, lw: 1.501, tx: 1.131, align: 'right', ty: 2.723, tile: 3.792, dark: true },
    { label: 'weaknesses',     icon: 1, lx: 9.766, lw: 1.707, tx: 9.766, align: 'left',  ty: 2.723, tile: 8.997, dark: false },
    { label: 'Opportunities ', icon: 3, lx: 1.679, lw: 1.859, tx: 1.131, align: 'right', ty: 5.589, tile: 3.792, dark: false },
    { label: 'Threats ',       icon: 2, lx: 9.766, lw: 1.501, tx: 9.766, align: 'left',  ty: 5.589, tile: 8.997, dark: true },
  ].forEach(b => {
    head(s, b.label, { x: b.lx, y: b.ty, w: b.lw, h: 0.337, fontSize: 14, align: b.align });
    txt(s, [{ text: 'Lorem ipsum dolor sit amet,', options: { breakLine: true } }, { text: 'Consectetur adipiscing elit' }],
      { x: b.tx, y: b.ty + 0.35, w: 2.436, h: 0.533, fontSize: 11, align: b.align, lineSpacingMultiple: 1.2 });
    roundRect(s, { x: b.tile, y: b.ty + 0.14, w: 0.604, h: 0.604, rectRadius: 0.101, fill: { color: b.dark ? NAVY : WHITE }, shadow: shadow('card') });
    iconGlyph(s, b.tile + 0.093, b.ty + 0.235, 0.42, b.dark ? WHITE : NAVY, b.icon);
  });
}

// 7 -- Split layout: white text column on the left, dark stage on the right.
function slide07(s) {
  s.background = { color: NAVY };
  rect(s, { x: 0, y: 0, w: 4.337, h: 7.5, fill: { color: WHITE } });
  line(s, { x: 4.337, y: 0.838, w: 8.997, h: 0, line: { color: WHITE, width: 1 } });
  line(s, { x: 12.343, y: 0.044, w: 0, h: 7.412, line: { color: WHITE, width: 1 } });
  chrome(s, 7, { numberColor: WHITE });

  roundRect(s, { x: 0.605, y: 2.994, w: 1.047, h: 1.047, rectRadius: 0.175, fill: { color: NAVY }, shadow: shadow('card') });
  iconGlyph(s, 0.773, 3.163, 0.709, WHITE, 0);
  txt(s, 'Lorem amet ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus',
    { x: 0.446, y: 4.407, w: 3.546, h: 1.133, align: 'justify', lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem amet ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean',
    { x: 0.446, y: 5.593, w: 3.546, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  line(s, { x: 5.323, y: 4.518, w: 1.975, h: 0, line: { color: WHITE, width: 1 } });
  head(s, 'Key Milestones for Implementation', { x: 7.435, y: 4.333, w: 4.191, h: 0.37, fontSize: 16, color: WHITE });
  head(s, 'Execution Timeline Overview', { x: 5.138, y: 4.746, w: 6.488, h: 1.717, fontSize: 48, color: WHITE });
  learnMore(s, 9.092, 5.032, false);
}

// 8 -- Folder panel holding a three-step line chart.
const FOLDER_PANEL = [['M', 0, 0], ['L', 0.3194, 0], ['L', 0.3603, 0.0506], ['L', 0.3603, 0.0509], ['L', 0.9826, 0.0509],
  ['C', 0.9922, 0.0509, 1, 0.0606, 1, 0.0725], ['L', 1, 0.9784],
  ['C', 1, 0.9903, 0.9922, 1, 0.9826, 1], ['L', 0.0174, 1],
  ['C', 0.0078, 1, 0, 0.9903, 0, 0.9784], ['L', 0, 0.0725], ['Z']];

function slide08(s) {
  chrome(s, 8);
  fracPath(s, { x: 5.929, y: 1.027, w: 6.717, h: 5.417, path: FOLDER_PANEL, fill: { color: NAVY } });
  head(s, 'Early-Stage Initiatives and Preparations', { x: 0.788, y: 1.506, w: 4.818, h: 2.322, fontSize: 44 });
  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget ligula ipsum dolor sit amet, consectetuer',
    { x: 0.788, y: 5.124, w: 4.445, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  // highlighted column behind step 02 fades from white (bottom) into the panel
  gradientBands(s, {
    x: 8.454, y: 2.503, w: 1.823, h: 3.495, dir: 'v', steps: 60,
    colorAt: t => { const p = 1 - t; return p >= 0.78 ? NAVY : mix(WHITE, NAVY, p / 0.78); },
  });
  [6.480, 8.454, 10.277, 12.100].forEach(x => line(s, { x: x, y: 2.503, w: 0, h: 3.495, line: { color: WHITE, width: 1 } }));

  head(s, 'Foundation Steps to Begin the Strategic', { x: 6.48, y: 1.748, w: 4.962, h: 0.377, fontSize: 16, color: WHITE });
  [
    { label: 'Step 01', x: 6.658, y: 2.503 },
    { label: 'Step 02', x: 8.556, y: 2.503 },
    { label: 'Step 03', x: 10.379, y: 5.119 },
  ].forEach(st => {
    head(s, st.label, { x: st.x, y: st.y, w: 1.178, h: 0.314, fontSize: 12, color: WHITE });
    txt(s, 'Lorem sit ligula ipsum dolor', { x: st.x, y: st.y + 0.314, w: 1.619, h: 0.567, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.3 });
  });

  // rising trend line + markers
  s.addShape('custGeom', {
    x: 6.462, y: 2.945, w: 5.634, h: 2.759, fill: { type: 'none' }, line: { color: WHITE, width: 1 },
    points: [{ x: 0, y: 2.759, moveTo: true }, { x: 2.050, y: 2.128 }, { x: 3.827, y: 0.971 }, { x: 5.634, y: 0 }],
  });
  [[6.380, 5.604], [8.349, 4.975], [10.180, 3.826], [11.996, 2.845]].forEach(([x, y]) =>
    ellipse(s, { x: x, y: y, w: 0.2, h: 0.2, fill: { color: WHITE }, line: { color: ICE, width: 3.75 } }));
}

// 9 -- Horizontal milestone timeline on navy.
function slide09(s) {
  s.background = { color: NAVY };
  chrome(s, 9, { titleColor: WHITE, numberColor: WHITE });
  head(s, 'Mid-Year Development and Optimization', { x: 0.951, y: 1.256, w: 7.566, h: 1.582, fontSize: 44, color: WHITE });
  txt(s, LOREM_MED, { x: 8.951, y: 1.612, w: 3.431, h: 0.87, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3 });
  line(s, { x: 0.417, y: 4.365, w: 12.5, h: 0, line: { color: WHITE, width: 1, beginArrowType: 'oval', endArrowType: 'oval' } });

  [
    { year: '2023', label: 'Implementation', px: 0.978, dot: 1.458, tx: 1.735, tw: 2.573 },
    { year: '2024', label: 'Development',    px: 5.072, dot: 5.554, tx: 5.833, tw: 2.587 },
    { year: '2025', label: 'Strategy',       px: 9.185, dot: 9.667, tx: 9.948, tw: 2.485 },
  ].forEach(m => {
    roundRect(s, { x: m.px, y: 3.578, w: 1.098, h: 0.466, rectRadius: 0.05, fill: { color: WHITE }, shadow: shadow('card') });
    head(s, m.year, { x: m.px + 0.094, y: 3.643, w: 0.909, h: 0.337, fontSize: 14, bold: true, align: 'center' });
    ellipse(s, { x: m.dot, y: 4.298, w: 0.134, h: 0.134, fill: { color: WHITE }, line: { color: ICE, width: 3.75 } });
    line(s, { x: m.dot + 0.068, y: 4.432, w: 0, h: 2.013, line: { color: WHITE, width: 1 } });
    head(s, m.label, { x: m.tx, y: 5.203, w: 2.087, h: 0.337, fontSize: 14, color: WHITE });
    txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean',
      { x: m.tx, y: 5.574, w: m.tw, h: 0.87, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3 });
  });
}

// 10 -- Six quarter columns of alternating height.
function slide10(s) {
  chrome(s, 10);
  head(s, 'Strategic Expansion and Acceleration', { x: 2.944, y: 1.062, w: 7.445, h: 1.582, fontSize: 44, align: 'center' });
  head(s, '2022', { x: 0.43, y: 6.467, w: 1.056, h: 0.37, fontSize: 16 });
  head(s, '2025', { x: 11.847, y: 6.467, w: 1.056, h: 0.37, fontSize: 16, align: 'right' });
  line(s, { x: 0.573, y: 6.224, w: 12.188, h: 0, line: { color: BLACK, transparency: 67, width: 1, beginArrowType: 'oval', endArrowType: 'oval' } });

  [
    { q: 'Q1', x: 0.722,  y: 2.273, h: 3.952, dark: true  },
    { q: 'Q2', x: 2.736,  y: 3.043, h: 3.181, dark: false },
    { q: 'Q3', x: 4.751,  y: 3.660, h: 2.564, dark: true  },
    { q: 'Q4', x: 6.766,  y: 3.660, h: 2.564, dark: false },
    { q: 'Q5', x: 8.781,  y: 3.043, h: 3.181, dark: true  },
    { q: 'Q6', x: 10.796, y: 2.273, h: 3.952, dark: false },
  ].forEach(c => {
    const fg = c.dark ? WHITE : BLACK;
    topRoundedRect(s, { x: c.x, y: c.y, w: 1.816, h: c.h, r: 0.176, fill: { color: c.dark ? NAVY : WHITE }, shadow: shadow('card') });
    head(s, c.q, { x: c.x + 0.297, y: c.y + 0.171, w: 1.222, h: 0.812, fontSize: 44, align: 'center', color: fg });
    txt(s, LOREM_BAR, { x: c.x + 0.188, y: 5.186, w: 1.438, h: 0.608, align: 'center', color: fg, lineSpacingMultiple: 1.3 });
  });
  txt(s, LOREM_MED, { x: 1.853, y: 6.479, w: 9.628, h: 0.345, align: 'center', lineSpacingMultiple: 1.3 });
}

// 11 -- Ascending percentage columns beside a headline figure.
function slide11(s) {
  chrome(s, 11);
  head(s, 'Review Evaluate Future Planning', { x: 0.569, y: 1.212, w: 5.451, h: 1.582, fontSize: 44 });
  head(s, '73.228+', { x: 0.569, y: 4.542, w: 2.742, h: 0.841, fontSize: 44, color: NAVY });
  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean',
    { x: 0.569, y: 5.465, w: 2.742, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  line(s, { x: 0.736, y: 6.733, w: 12.028, h: 0, line: { color: NAVY, width: 1, beginArrowType: 'oval', endArrowType: 'oval' } });

  [
    { pct: '74%', label: 'Review',   x: 3.931,  y: 3.972, h: 2.760, dark: false },
    { pct: '84%', label: 'Evaluate', x: 6.144,  y: 3.003, h: 3.729, dark: false },
    { pct: '87%', label: 'Future ',  x: 8.364,  y: 2.233, h: 4.500, dark: false },
    { pct: '94%', label: 'Planning', x: 10.584, y: 1.212, h: 5.521, dark: true  },
  ].forEach(c => {
    const fg = c.dark ? WHITE : BLACK;
    if (c.dark) {
      // gradient column: transparent at the bottom, solid navy from 70% up
      topRoundedRect(s, { x: c.x, y: c.y, w: 2.034, h: c.h, r: 0.127, fill: { color: NAVY } });
      gradientBands(s, {
        x: c.x, y: c.y + c.h * 0.3, w: 2.034, h: c.h * 0.7, dir: 'v', steps: 60,
        colorAt: t => mix(NAVY, WHITE, t),
      });
    } else {
      topRoundedRect(s, { x: c.x, y: c.y, w: 2.034, h: c.h, r: 0.127, fill: { color: WHITE }, shadow: shadow('card') });
    }
    head(s, c.pct, { x: c.x + 0.468, y: c.y + 0.25, w: 1.098, h: 0.505, fontSize: 24, bold: true, align: 'center', color: fg });
    head(s, c.label, { x: c.x + 0.297, y: c.y + 0.696, w: 1.438, h: 0.419, fontSize: 16, align: 'center', color: fg, lineSpacingMultiple: 1.3 });
    txt(s, LOREM_BAR, { x: c.x + 0.297, y: 5.789, w: 1.438, h: 0.608, align: 'center', color: fg, lineSpacingMultiple: 1.3 });
  });
}

// 12 -- Two-column section divider with a numbered framework entry.
function slide12(s) {
  chrome(s, 12);
  rect(s, { x: 4.927, y: 0, w: 8.406, h: 7.5, fill: { color: NAVY } });
  line(s, { x: 4.927, y: 0.548, w: 8.406, h: 0, line: { color: WHITE, width: 1 } });
  line(s, { x: 4.927, y: 6.952, w: 8.406, h: 0, line: { color: WHITE, width: 1 } });
  line(s, { x: 12.766, y: 0.05, w: 0, h: 7.401, line: { color: WHITE, width: 1 } });
  head(s, 'Key Business', { x: 0.673, y: 1.197, w: 3.535, h: 1.582, fontSize: 44 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis',
    { x: 0.673, y: 5.349, w: 3.639, h: 0.87, lineSpacingMultiple: 1.3 });
  fracPath(s, { x: 5.508, y: 3.668, w: 3.444, h: 0.707, path: TAB_PATH, fill: { color: NAVY } });
  head(s, 'Strategy Model', { x: 5.557, y: 3.86, w: 3.027, h: 0.505, fontSize: 24, color: WHITE });
  head(s, '01.', { x: 5.557, y: 5.108, w: 0.771, h: 0.505, fontSize: 24, color: WHITE });
  head(s, 'Essential Strategy Framework', { x: 6.186, y: 5.236, w: 4.107, h: 0.337, fontSize: 14, color: WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus ipsum',
    { x: 5.557, y: 5.611, w: 6.595, h: 0.608, color: WHITE, lineSpacingMultiple: 1.3 });
}

// 13 -- Interlocking cogs carrying KPI badges.
function slide13(s) {
  chrome(s, 13);
  head(s, 'Market Expansion & Customer Reach', { x: 0.71, y: 1.498, w: 5.451, h: 2.322, fontSize: 44 });
  head(s, 'Approaches to Increase Market Penetration', { x: 0.71, y: 4.914, w: 5.09, h: 0.419, fontSize: 16, lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget ligula ipsum dolor sit amet, ',
    { x: 0.71, y: 5.395, w: 5.587, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });

  const SMALL_COG = { w: 2.056, h: 2.056, teeth: 6, tipR: 0.983, rootR: 0.700, phase: 30, fill: { color: NAVY } };
  gear(s, Object.assign({ x: 6.706, y: 2.836 }, SMALL_COG));
  gear(s, Object.assign({ x: 10.346, y: 1.360 }, SMALL_COG));
  gear(s, { x: 8.239, y: 3.074, w: 3.666, h: 3.666, teeth: 9, tipR: 0.995, rootR: 0.800, phase: 30, rotate: 354.5, fill: { color: NAVY } });
  gear(s, { x: 7.646, y: 0.760, w: 3.028, h: 3.028, teeth: 6, tipR: 0.802, rootR: 0.572, phase: 15, fill: { color: NAVY } });

  ellipse(s, { x: 7.255, y: 3.385, w: 0.958, h: 0.958, fill: { color: WHITE }, shadow: shadow('gear') });
  iconGlyph(s, 7.53, 3.66, 0.41, NAVY, 2);
  ellipse(s, { x: 10.894, y: 1.909, w: 0.958, h: 0.958, fill: { color: WHITE }, shadow: shadow('gear') });
  iconGlyph(s, 11.16, 2.175, 0.43, NAVY, 3);

  ellipse(s, { x: 9.053, y: 3.887, w: 2.04, h: 2.04, fill: { color: WHITE }, shadow: shadow('gear') });
  head(s, 'Value One', { x: 9.187, y: 4.578, w: 1.771, h: 0.386, fontSize: 20, align: 'center' });
  txt(s, [{ text: '5.204+ ', options: { fontFace: HEAD, fontSize: 14 } }, { text: 'Profit', options: { fontSize: 12 } }],
    { x: 9.187, y: 4.948, w: 1.771, h: 0.337, align: 'center', color: NAVY });

  ellipse(s, { x: 8.592, y: 1.706, w: 1.137, h: 1.137, fill: { color: WHITE }, shadow: shadow('gear') });
  head(s, '64%', { x: 8.529, y: 1.887, w: 1.261, h: 0.501, fontSize: 20, bold: true, align: 'center', lineSpacingMultiple: 1.4 });
  head(s, 'Successful', { x: 8.597, y: 2.255, w: 1.126, h: 0.333, fontSize: 11, align: 'center', color: NAVY, lineSpacingMultiple: 1.4 });
}

// 14 -- Three metric cards linked to icon tiles, desk illustration bottom-right.
function slide14(s) {
  chrome(s, 14);
  head(s, 'Improving Internal Efficiency', { x: 2.105, y: 1.183, w: 9.124, h: 0.841, fontSize: 44, align: 'center' });

  [
    { card: 1.333, cardY: 2.529, radius: 0.131, value: '58+', vx: 4.179, bar: 5.591, barW: 3.011, barY: 3.108, tile: 8.574, tileY: 2.586, tx: 1.302, icon: 0 },
    { card: 2.179, cardY: 3.943, radius: 0.081, value: '32+', vx: 5.071, bar: 6.528, barW: 0.912, barY: 4.514, tile: 7.432, tileY: 4.026, tx: 2.290, icon: 1 },
    { card: 1.333, cardY: 5.356, radius: 0.081, value: '70+', vx: 4.115, bar: 5.591, barW: 1.020, barY: 5.925, tile: 6.609, tileY: 5.423, tx: 1.333, icon: 2 },
  ].forEach(r => {
    roundRect(s, { x: r.card, y: r.cardY, w: 4.368, h: 1.206, rectRadius: r.radius, fill: { color: WHITE }, shadow: shadow('card') });
    txt(s, LOREM_CARD, { x: r.tx, y: r.cardY + 0.27, w: 2.767, h: 0.608, align: 'right', lineSpacingMultiple: 1.3 });
    head(s, r.value, { x: r.vx, y: r.cardY + 0.25, w: 1.17, h: 0.707, fontSize: 36, align: 'center' });
    roundRect(s, { x: r.bar - 0.05, y: r.barY - 0.191, w: 0.05, h: 0.412, rectRadius: 0.025, fill: { color: NAVY } });
    rect(s, { x: r.bar, y: r.barY, w: r.barW, h: 0.104, fill: { color: ICE } });
    roundRect(s, { x: r.tile, y: r.tileY, w: 1.082, h: 1.08, rectRadius: 0.18, fill: { color: ICE }, shadow: shadow('desk') });
    roundRect(s, { x: r.tile + 0.097, y: r.tileY + 0.095, w: 0.89, h: 0.89, rectRadius: 0.148, fill: { color: NAVY } });
    iconGlyph(s, r.tile + 0.29, r.tileY + 0.29, 0.5, WHITE, r.icon);
  });

  imagePlaceholder(s, { x: 10.036, y: 4.056, w: 2.288, h: 2.218 });
  topRoundedRect(s, { x: 7.842, y: 6.048, w: 5.491, h: 1.452, r: 0.72, fill: { color: NAVY_LT } });
}

// 15 -- Dark slide: figure illustration, KPI gauge and two progress bars.
function slide15(s) {
  s.background = { color: NAVY };
  chrome(s, 15, { titleColor: WHITE, numberColor: WHITE });
  head(s, 'Advancing Product & Service Value', { x: 7.818, y: 0.995, w: 4.616, h: 2.322, fontSize: 44, color: WHITE });

  imagePlaceholder(s, { x: 3.701, y: 1.922, w: 3.136, h: 4.848, fill: NAVY_LT, stroke: NAVY_SOFT, text: ICE_DEEP });

  [
    { value: '29k', label: 'Point #01', y: 3.314 },
    { value: '153k', label: 'Point #02', y: 4.753 },
  ].forEach(p => {
    head(s, p.value, { x: 1.483, y: p.y, w: 1.493, h: 0.78, fontSize: 36, color: WHITE, lineSpacingMultiple: 1.2 });
    txt(s, p.label, { x: 1.483, y: p.y + 0.682, w: 1.516, h: 0.365, fontSize: 14, fontFace: HEAD_MED, color: WHITE, lineSpacingMultiple: 1.2 });
    s.addShape('triangle', { x: 2.75, y: p.y + 0.295, w: 0.219, h: 0.189, fill: { color: WHITE } });
  });

  [
    { pct: '65%', y: 3.649, barY: 4.238, fillW: 1.120, textW: 4.033 },
    { pct: '85%', y: 5.209, barY: 5.798, fillW: 1.399, textW: 3.838 },
  ].forEach(p => {
    head(s, p.pct, { x: 7.818, y: p.y, w: 1.001, h: 0.572, fontSize: 28, color: WHITE, wrap: false });
    roundRect(s, { x: 7.935, y: p.barY, w: 1.81, h: 0.05, rectRadius: 0.025, fill: { color: LT2 } });
    roundRect(s, { x: 7.935, y: p.barY, w: p.fillW, h: 0.05, rectRadius: 0.025, fill: { color: NAVY_SOFT } });
    txt(s, 'Lorem ligula ipsum dolor sit amet, Aenean consectetuer amet adipiscing elit. ',
      { x: 7.818, y: p.barY + 0.1, w: p.textW, h: 0.608, color: WHITE, lineSpacingMultiple: 1.3 });
  });

  // magnifier bubble with a segmented value ring
  ellipse(s, { x: 2.146, y: 1.371, w: 1.691, h: 1.698, fill: { color: WHITE }, shadow: shadow('bubble') });
  s.addShape('triangle', { x: 3.09, y: 2.83, w: 0.42, h: 0.55, rotate: 205, fill: { color: WHITE } });
  [
    { range: [183.89, 266.33], color: GREY_50, transparency: 80 },
    { range: [36.21, 222.52], color: GREY_25 },
    { range: [356.22, 31.03], color: GREY_35, transparency: 80 },
    { range: [321.21, 11.78], color: GREY_25 },
    { range: [291.77, 315.28], color: GREEN, transparency: 80 },
    { range: [271.89, 314.94], color: GREY_25 },
  ].forEach(p => s.addShape('pie', { x: 2.3, y: 1.548, w: 1.324, h: 1.324, angleRange: p.range, fill: { color: p.color, transparency: p.transparency } }));
  ellipse(s, { x: 2.368, y: 1.616, w: 1.187, h: 1.187, fill: { color: WHITE } });
  head(s, '65%', { x: 2.387, y: 1.918, w: 1.15, h: 0.438, fontSize: 20, bold: true, align: 'center', color: NAVY });
  txt(s, 'Value', { x: 2.387, y: 2.273, w: 1.15, h: 0.286, fontSize: 11, align: 'center', color: NAVY });
}

// 16 -- Head-to-head "VS" comparison.
function slide16(s) {
  chrome(s, 16);
  roundRect(s, { x: 3.118, y: 1.497, w: 7.034, h: 2.393, rectRadius: 0.399, fill: { color: WHITE }, shadow: shadow('panel') });
  imagePlaceholder(s, { x: 3.177, y: 1.359, w: 2.142, h: 2.248, fill: SLATE_MID, stroke: SLATE, text: WHITE });
  imagePlaceholder(s, { x: 7.915, y: 1.359, w: 2.142, h: 2.248, fill: SLATE_MID, stroke: SLATE, text: WHITE });
  imagePlaceholder(s, { x: 5.421, y: 1.779, w: 2.226, h: 2.130, fill: NAVY, stroke: SLATE_DK, text: ICE_DEEP });
  head(s, 'VS', { x: 5.227, y: 3.035, w: 2.752, h: 2.366, fontSize: 138, align: 'center', valign: 'middle', color: WHITE, wrap: false });
  line(s, { x: 6.603, y: 5.045, w: 0, h: 1.096, line: { color: NAVY, width: 1.5 } });

  txt(s, LOREM_CARD, { x: 0.428, y: 2.463, w: 2.505, h: 0.572, align: 'right', lineSpacingMultiple: 1.2 });
  txt(s, LOREM_CARD, { x: 10.400, y: 2.463, w: 2.505, h: 0.572, lineSpacingMultiple: 1.2 });

  [
    { icon: 3, tile: 8.654, tileY: 4.215, title: 'Project Strategy #1', tx: 9.481, align: 'left',  rule: 8.607, track: 8.654, fillX: 8.728, fillW: 1.711, sub: 8.570, badge: 11.612, pct: '56%' },
    { icon: 2, tile: 3.912, tileY: 4.218, title: 'Project Strategy #2', tx: 1.429, align: 'right', rule: 1.057, track: 1.883, fillX: 2.685, fillW: 1.882, sub: 3.002, badge: 1.214, pct: '74%' },
  ].forEach(c => {
    roundRect(s, { x: c.tile, y: c.tileY, w: 0.705, h: 0.705, rectRadius: 0.118, fill: { color: NAVY }, line: { color: ICE, width: 8 } });
    iconGlyph(s, c.tile + 0.19, c.tileY + 0.2, 0.32, WHITE, c.icon);
    head(s, c.title, { x: c.tx, y: c.tileY - 0.073, w: 2.36, h: 0.361, fontSize: 16, align: c.align, wrap: false });
    txt(s, LOREM_CARD, { x: c.align === 'right' ? 0.846 : 9.481, y: c.tileY + 0.27, w: 2.944, h: 0.572, align: c.align, lineSpacingMultiple: 1.2 });
    line(s, { x: c.rule, y: c.tileY + 1.185, w: 3.585, h: 0, line: { color: NAVY, width: 1 } });
    txt(s, 'Subtitle Text Here', { x: c.sub, y: c.tileY + 1.35, w: 1.698, h: 0.296, fontSize: 12, fontFace: HEAD_MED, align: c.align, wrap: false });
    roundRect(s, { x: c.track, y: 5.903, w: 2.733, h: 0.185, rectRadius: 0.0925, fill: { color: WHITE }, shadow: shadow('card') });
    roundRect(s, { x: c.fillX, y: 5.937, w: c.fillW, h: 0.117, rectRadius: 0.0585, fill: { color: NAVY } });
    roundRect(s, { x: c.badge, y: 5.821, w: 0.519, h: 0.312, rectRadius: 0.071, fill: { color: NAVY } });
    s.addShape('triangle', { x: c.badge - 0.105, y: 5.935, w: 0.217, h: 0.156, rotate: 270, fill: { color: NAVY } });
    head(s, c.pct, { x: c.badge - 0.006, y: 5.837, w: 0.534, h: 0.279, fontSize: 11, align: 'center', color: WHITE, wrap: false });
  });
}

// 17 -- Dark sidebar with a tabbed card next to a text column.
function slide17(s) {
  rect(s, { x: 0, y: 0, w: 5.172, h: 7.5, fill: { color: NAVY } });
  chrome(s, 17, { titleColor: WHITE });
  line(s, { x: 0, y: 0.775, w: 5.161, h: 0, line: { color: WHITE, width: 1 } });
  line(s, { x: 0, y: 6.725, w: 5.161, h: 0, line: { color: WHITE, width: 1 } });
  roundRect(s, { x: 0.714, y: 1.687, w: 3.626, h: 4.528, rectRadius: 0.099, fill: { color: WHITE } });
  fracPath(s, { x: 0.714, y: 1.285, w: 2.361, h: 0.511, path: TAB_PATH, fill: { color: WHITE } });
  head(s, 'External Risks', { x: 0.896, y: 1.393, w: 1.853, h: 0.37, fontSize: 16, color: NAVY });

  head(s, 'Business Performance Data Overview', { x: 5.688, y: 1.553, w: 7.358, h: 1.582, fontSize: 44 });
  learnMore(s, 10.664, 2.538, true);
  txt(s, LOREM_LONG, { x: 5.688, y: 3.24, w: 6.66, h: 0.608, lineSpacingMultiple: 1.3 });
  head(s, 'Approaches to Increase Market Penetration', { x: 5.688, y: 4.859, w: 5.09, h: 0.419, fontSize: 16, lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget ligula ipsum dolor sit amet, ',
    { x: 5.688, y: 5.34, w: 5.587, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
}

// 18 -- Four doughnut charts on two cards, each with a value callout.
function slide18(s) {
  chrome(s, 18);
  head(s, 'Revenue & Growth Trends', { x: 2.098, y: 1.344, w: 9.124, h: 0.841, fontSize: 44, align: 'center' });
  roundRect(s, { x: 0.867, y: 2.736, w: 5.66, h: 2.885, rectRadius: 0.314, fill: { color: WHITE }, shadow: shadow('chart') });
  roundRect(s, { x: 6.827, y: 2.751, w: 5.66, h: 2.885, rectRadius: 0.314, fill: { color: WHITE }, shadow: shadow('chart') });

  const donut = { holeSize: 64, firstSliceAng: 90, showLegend: false, showTitle: false, showValue: false, chartArea: { fill: { type: 'none' } }, layout: { x: 0.0948, y: 0.0912, w: 0.8104, h: 0.8176 } };
  [
    { x: 0.846, values: [50, 50], colors: [LT2, NAVY], pct: '50%', px: 1.875 },
    { x: 3.830, values: [20, 80], colors: [LT2, NAVY], pct: '80%', px: 4.796 },
    { x: 6.814, values: [25, 75], colors: [GREY_95, NAVY], pct: '75%', px: 7.801 },
    { x: 9.798, values: [75, 25], colors: [NAVY, GREY_95], pct: '25%', px: 10.790 },
  ].forEach(d => {
    s.addChart('doughnut', [{ name: 'Sales', labels: ['a', 'b'], values: d.values }],
      Object.assign({ x: d.x, y: 2.861, w: 2.676, h: 2.652, chartColors: d.colors }, donut));
    head(s, d.pct, { x: d.px, y: 3.985, w: 0.745, h: 0.404, fontSize: 24, bold: true, align: 'center', color: NAVY, wrap: false, margin: 0 });
  });

  // value callout pills
  [
    { x: 2.262, y: 2.954, label: '15.255+' },
    { x: 3.371, y: 4.843, label: '34.465+' },
    { x: 8.230, y: 2.954, label: '27.565+' },
    { x: 9.324, y: 4.843, label: '7.835+' },
  ].forEach(c => {
    roundRect(s, { x: c.x, y: c.y, w: 1.449, h: 0.424, rectRadius: 0.212, fill: { color: ICE_DEEP }, shadow: shadow('card') });
    ellipse(s, { x: c.x + 0.087, y: c.y + 0.065, w: 0.285, h: 0.284, fill: { color: NAVY } });
    s.addShape('chevron', { x: c.x + 0.195, y: c.y + 0.14, w: 0.062, h: 0.13, fill: { color: WHITE } });
    head(s, c.label, { x: c.x + 0.414, y: c.y + 0.04, w: 1.15, h: 0.344, fontSize: 14, valign: 'middle' });
  });

  txt(s, 'Lorem ligula ipsum dolor sit amet, consectetuer amet adipiscing elit. Aenean commodo ligula eget ligula ipsum dolor sit amet, ligula ipsum dolor sit amet, consectetuer amet',
    { x: 2.521, y: 5.973, w: 8.279, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 });
}

// 19 -- Four progress rings, the second one highlighted on a dark card.
function slide19(s) {
  chrome(s, 19);
  head(s, 'Revenue & Growth Trends', { x: 2.098, y: 0.883, w: 9.124, h: 0.841, fontSize: 44, align: 'center' });
  roundRect(s, { x: 3.864, y: 2.172, w: 2.789, h: 4.571, rectRadius: 0.174, fill: { color: NAVY }, shadow: shadow('card') });

  [
    { pct: '75%', end: 60.556,  x: 1.550, dark: false, pctColor: NAVY,  label: 'External Risks',   ly: 4.899, by: 5.262 },
    { pct: '43%', end: 274.372, x: 4.366, dark: true,  pctColor: BLACK, label: 'Outside Threats',  ly: 5.364, by: 5.734 },
    { pct: '61%', end: 11.210,  x: 7.182, dark: false, pctColor: BLACK, label: 'External Threats', ly: 4.899, by: 5.262 },
    { pct: '50%', end: 331.215, x: 9.998, dark: false, pctColor: BLACK, label: 'Outside Risks',    ly: 4.899, by: 5.262 },
  ].forEach(r => {
    const y = r.dark ? 2.566 : 2.568;
    ellipse(s, { x: r.x, y: y, w: 1.785, h: 1.785, fill: { color: r.dark ? ICE : WHITE }, shadow: shadow('ring') });
    head(s, r.pct, { x: r.x, y: y + 0.5, w: 1.785, h: 0.78, fontSize: 32, bold: true, align: 'center', color: r.pctColor, margin: 0 });
    s.addShape('arc', {
      x: r.x + 0.053, y: y + 0.053, w: 1.68, h: 1.68, rotate: 30, angleRange: [145.621, r.end],
      fill: { type: 'none' }, line: { color: r.dark ? WHITE : NAVY, width: 8.5, endArrowType: 'oval' },
    });
    head(s, r.label, { x: r.x - 0.328, y: r.ly, w: 2.441, h: 0.37, fontSize: 16, align: 'center', color: r.dark ? WHITE : NAVY });
    txt(s, LOREM_RING, { x: r.x - 0.255, y: r.by, w: 2.295, h: 0.608, align: 'center', color: r.dark ? WHITE : BLACK, lineSpacingMultiple: 1.3 });
  });
  iconGlyph(s, 4.982, 4.69, 0.53, WHITE, 1);
}

// 20 -- Four mini bar-chart groups with captions.
function slide20(s) {
  chrome(s, 20);
  const GROUPS = [
    { x: 1.172, base: 1.183, label: 'External Risks',   tx: 1.029, muted: true  },
    { x: 4.161, base: 4.166, label: 'Outside Threats',  tx: 3.971, muted: false },
    { x: 7.150, base: 7.150, label: 'External Threats', tx: 6.912, muted: false },
    { x: 10.138, base: 10.138, label: 'Outside Risks',  tx: 9.854, muted: false },
  ];
  // bar heights per group (inches), tracks are always 2.462 tall ending at y=3.971
  const HEIGHTS = [
    [1.164, 1.904, 1.589, 2.264],
    [1.904, 2.046, 1.589, 0.614],
    [1.904, 2.046, 1.589, 2.173],
    [1.904, 2.046, 1.589, 2.173],
  ];

  GROUPS.forEach((g, gi) => {
    HEIGHTS[gi].forEach((h, bi) => {
      const x = g.x + [0, 0.575, 1.150, 1.729][bi];
      roundRect(s, { x: x, y: 1.509, w: 0.294, h: 2.462, rectRadius: 0.049, fill: { color: GREY_95 } });
      roundRect(s, { x: x, y: 3.971 - h, w: 0.294, h: h, rectRadius: 0.049, fill: { color: NAVY } });
    });
    rect(s, { x: g.base, y: 4.145, w: 2.013, h: 0.116, fill: { color: GREY_85 } });
    ellipse(s, { x: g.x + 0.75, y: 4.573, w: 0.799, h: 0.799, fill: { color: NAVY } });
    iconGlyph(s, g.x + 1.00, 4.822, 0.30, WHITE, gi);
    head(s, g.label, { x: g.tx + 0.136, y: 5.511, w: 2.315, h: 0.37, fontSize: 16, align: 'center', color: NAVY });
    txt(s, LOREM_RING, { x: g.tx, y: 5.882, w: 2.585, h: 0.608, align: 'center', color: g.muted ? GREY_25 : BLACK, lineSpacingMultiple: 1.3 });
  });
}

/* ------------------------------------------------------------------ *
 * Build & save
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SW, height: SH });
  pptx.layout = 'DECK';
  pptx.title = DECK_TITLE;

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(builder => builder(pptx.addSlide()));

  return pptx.writeFile({ fileName: path.join(__dirname, '022902c4-cc7b-48fa-acaf-632b7b3ee216_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
