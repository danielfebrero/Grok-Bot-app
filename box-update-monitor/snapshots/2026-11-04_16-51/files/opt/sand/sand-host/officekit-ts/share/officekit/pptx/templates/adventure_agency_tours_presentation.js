/*
 * "Adventrail" travel-agency deck (31 slides, 16:9 / 13.333 x 7.5 in)
 * rebuilt with pptxgenjs. Photographs in the original are replaced by
 * flat grey placeholder rectangles labelled "[image]".
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const TEAL = '01626F'; // primary dark teal
const CYAN = '028A9C'; // secondary teal
const LIME = 'B9DF41'; // primary lime
const GREEN = '92D050'; // logo green
const GRASS = '00B050';
const OLIVE = '9BC121';
const DARK = '404040'; // headings  (tx1 lumMod 75%)
const BODY = '262626'; // body copy (tx1 lumMod 85%)
const SOFT = 'D9D9D9'; // bg1 lumMod 95%
const SMOKE = 'F2F2F2';
const WHITE = 'FFFFFF';
const PHOTO = '6F6F6F'; // stand-in colour for photographs
const PHOTO_DARK = '626262';
const PHOTO_LIGHT = '868686';
const CAPTION = 'BFBFBF';

const HEAD = 'Montserrat';
const LIGHT = 'Montserrat Light';
const BLACK = 'Montserrat Black';

const NOLINE = { type: 'none' };
const TM = [5.67, 5.67, 0, 0]; // lIns/rIns 0.079", no top/bottom inset  [l, r, b, t]

/* ------------------------------------------------------------------- copy   */
const L1 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut euismod, est eget maximus venenatis, ' +
  'augue justo tincidunt massa, a rhoncus tortor lectus ac tortor. ';
const L2 = 'Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ';
const L3 = 'Nulla condimentum dolor at tristique dapibus. ';
const LOREM_MED = L1 + L2;
const LOREM_FULL = L1 + L2 + L3;
const LOREM_SHORT =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut euismod, est eget maximus venenatis, ' +
  'augue justo tincidunt massa, a rhoncus tortor lectus ac tortor.';
const LOREM_HALF = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut euismod, est eget maximus venenatis, ';
const UTEU = 'Ut euismod, est eget maximus venenatis, augue justo tincidunt massa, a rhoncus tortor lectus ac tortor.';

const F1 = 'Fusce a felis porta, ultrices arcu eu, ultricies magna. Aliquam sed ipsum maximus est euismod dignissim. ';
const F2 = 'Maecenas consequat elementum eros, ut pharetra velit convallis nec. ';
const F3 = 'Vivamus interdum porta varius. Nam quis libero ipsum. ';
const FUSCE_MED = F1 + F2;
const FUSCE_LONG = F1 + F2 + F3 + 'Quisque ut nibh ante. Vestibulum sit amet nisi at leo tempor facilisis.';
const FUSCE_PANEL = F1 + F2 + F3;
const FUSCE_CARD = F1 + 'Maecenas consequat elementum eros, ';

const TAGLINE =
  'Company Adventure Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut euismod,  est eget ' +
  'maximus venenatis, augue justo tincidunt massa, a rhoncus tortor lectus ac tortor. ';

/* --------------------------------------------------------------- primitives */
const rect = (s, x, y, w, h, color, opt) =>
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NOLINE }, opt));

const text = (s, body, opt) => s.addText(body, Object.assign({ fontFace: HEAD }, opt));

/** Body copy: Montserrat Light, 150% leading. */
const para = (s, body, x, y, w, h, opt) =>
  text(s, body, Object.assign(
    { x, y, w, h, fontFace: LIGHT, fontSize: 11, color: BODY, lineSpacingMultiple: 1.5, valign: 'top', margin: TM },
    opt));

/** Section heading: bold Montserrat 36pt, vertically centred. */
const title = (s, body, x, y, w, h, opt) =>
  text(s, body, Object.assign(
    { x, y, w, h, fontSize: 36, bold: true, color: DARK, valign: 'middle', margin: TM }, opt));

/** Free-form polygon; points are inches relative to the shape box. */
const poly = (s, x, y, w, h, points, fill, opt) =>
  s.addShape('custGeom', Object.assign({ x, y, w, h, points, fill, line: NOLINE }, opt));

const mix = (a, b, t) => {
  const ch = (c, i) => parseInt(c.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0'))
    .join('').toUpperCase();
};

/**
 * Vertical fade used over every photo in the deck: transparent lime at the top
 * blending to 50%-opaque teal at the bottom.  Rendered as a stack of solid
 * bands because pptxgenjs has no gradient fill.  `slantTop`/`slantBottom`
 * inset the left edge so the fade can follow an angled photo.
 */
function shade(s, x, y, w, h, opts) {
  const o = Object.assign({ from: LIME, to: TEAL, top: 0, bottom: 50, steps: 10, slantTop: 0, slantBottom: 0 }, opts);
  for (let i = 0; i < o.steps; i++) {
    const t = (i + 0.5) / o.steps;
    const ya = (h * i) / o.steps;
    const bh = h / o.steps + 0.01;
    const fill = { color: mix(o.from, o.to, t), transparency: Math.round(100 - (o.top + (o.bottom - o.top) * t)) };
    if (o.slantTop || o.slantBottom) {
      const edge = (yy) => o.slantTop + (o.slantBottom - o.slantTop) * (Math.min(yy, h) / h);
      poly(s, x, y + ya, w, bh, [
        { x: edge(ya), y: 0 }, { x: w, y: 0 }, { x: w, y: bh }, { x: edge(ya + bh), y: bh }, { close: true },
      ], fill);
    } else {
      rect(s, x, y + ya, w, bh, o.color, { fill });
    }
  }
}

/**
 * Bottom "wave" band: a V-notch across the top edge.  When two colours are
 * given the band is sliced into vertical strips to fake a diagonal gradient.
 */
function chevron(s, y, h, dip, from, to, transparency) {
  const W = 13.333;
  if (!to) {
    poly(s, 0, y, W, h, [
      { x: 0, y: 0 }, { x: W / 2, y: dip }, { x: W, y: 0 }, { x: W, y: h }, { x: 0, y: h }, { close: true },
    ], { color: from, transparency });
    return;
  }
  const n = 16;
  const top = (x) => dip * (1 - Math.abs(1 - (2 * x) / W));
  for (let i = 0; i < n; i++) {
    const xa = (W * i) / n;
    const bw = W / n + 0.01;
    poly(s, xa, y, bw, h, [
      { x: 0, y: top(xa) }, { x: bw, y: top(xa + bw) }, { x: bw, y: h }, { x: 0, y: h }, { close: true },
    ], { color: mix(from, to, (i + 0.5) / n) });
  }
}

/**
 * Grey block standing in for a photograph.  A gentle dark-to-light wash from
 * the top-left keeps the flat placeholder from looking like a solid swatch.
 */
function photo(s, x, y, w, h, opt) {
  const o = opt || {};
  if (o.points || o.shape) {
    const fill = { color: PHOTO };
    if (o.points) poly(s, x, y, w, h, o.points, fill);
    else s.addShape(o.shape, { x, y, w, h, fill, line: NOLINE });
  } else {
    const n = Math.max(6, Math.round(w * 4));
    for (let i = 0; i < n; i++)
      rect(s, x + (w * i) / n, y, w / n + 0.02, h, mix(PHOTO_DARK, PHOTO_LIGHT, (i + 0.5) / n));
  }
  if (o.label !== false && w >= 2 && h >= 1.2)
    text(s, '[image]', { x, y, w, h, align: 'center', valign: 'middle', fontSize: 12, color: CAPTION });
}

/** Two stacked triangles = the "Adventure Agency" mountain mark. */
function mark(s, x, y, w, h, opt) {
  const o = opt || {};
  s.addShape('triangle', { x, y, w, h, fill: { color: TEAL, transparency: o.transparency || 0 }, line: NOLINE });
  s.addShape('triangle', {
    x, y: y + h * 0.5535, w, h: h * 0.4442,
    fill: { color: LIME, transparency: o.limeTransparency || o.transparency || 0 }, line: NOLINE,
  });
}

/** Header lock-up: "Adventure Agency" wordmark + mountain mark. */
function logo(s, x, y) {
  text(s, [
    { text: 'Adventure', options: { bold: true, color: GREEN } },
    { text: ' ', options: { color: GREEN } },
    { text: 'Agency', options: { color: TEAL } },
  ], { x, y, w: 1.393, h: 0.248, fontSize: 10, align: 'center', valign: 'middle', margin: 2.83 });
  mark(s, x + 1.393, y + 0.072, 0.129, 0.107);
}

/* ------------------------------------------------------------------- icons  */
const stroke = (color) => ({ color, width: 0.75 });

const ICONS = {
  // two overlapping peaks under a pair of little "cloud" zigzags
  mountain(s, x, y, u, ink) {
    s.addShape('triangle', { x: x + 0.06 * u, y: y + 0.42 * u, w: 0.44 * u, h: 0.5 * u, fill: { type: 'none' }, line: stroke(ink) });
    s.addShape('triangle', { x: x + 0.32 * u, y: y + 0.3 * u, w: 0.62 * u, h: 0.62 * u, fill: { type: 'none' }, line: stroke(ink) });
    [[0.04, 0.14, 0.5], [0.04, 0.24, 0.34]].forEach(([dx, dy, len]) => {
      const n = 3;
      for (let i = 0; i < n; i++)
        s.addShape('triangle', { x: x + (dx + (len * i) / n) * u, y: y + dy * u, w: (len / n) * u, h: 0.05 * u,
          fill: { type: 'none' }, line: stroke(ink) });
    });
  },
  // hourglass: two hex bowls, flat sides meeting in the middle
  tree(s, x, y, u, ink) {
    s.addShape('hexagon', { x: x + 0.17 * u, y: y + 0.02 * u, w: 0.66 * u, h: 0.42 * u, fill: { type: 'none' }, line: stroke(ink), rotate: 90 });
    s.addShape('hexagon', { x: x + 0.17 * u, y: y + 0.5 * u, w: 0.66 * u, h: 0.42 * u, fill: { type: 'none' }, line: stroke(ink), rotate: 90 });
    s.addShape('line', { x: x + 0.24 * u, y: y + 0.16 * u, w: 0.52 * u, h: 0, line: stroke(ink) });
    s.addShape('line', { x: x + 0.24 * u, y: y + 0.8 * u, w: 0.52 * u, h: 0, line: stroke(ink) });
  },
  pins(s, x, y, u, ink, opt) {
    const o = opt || {};
    const two = [[0.02, 0, ink], [0.44, 0.16, o.second || ink]];
    two.forEach(([dx, dy, col]) => {
      s.addShape('teardrop', { x: x + dx * u, y: y + dy * u, w: 0.44 * u, h: 0.44 * u,
        fill: { type: 'none' }, line: stroke(col), rotate: 135 });
      s.addShape('hexagon', { x: x + (dx + 0.13) * u, y: y + (dy + 0.09) * u, w: 0.18 * u, h: 0.2 * u,
        fill: o.solid ? { color: col } : { type: 'none' }, line: stroke(col), rotate: 90 });
    });
    s.addShape('line', { x: x + 0.24 * u, y: y + 0.62 * u, w: 0.42 * u, h: 0,
      line: { color: o.second || ink, width: 0.75, dashType: 'dash' } });
  },
  sun(s, x, y, u, ink) {
    s.addShape('ellipse', { x: x + 0.3 * u, y: y + 0.3 * u, w: 0.4 * u, h: 0.4 * u, fill: { type: 'none' }, line: stroke(ink) });
    for (let i = 0; i < 8; i++) {
      const a = (i * Math.PI) / 4;
      const r0 = 0.3 * u;
      const r1 = 0.46 * u;
      const cx = x + 0.5 * u;
      const cy = y + 0.5 * u;
      s.addShape('line', {
        x: cx + Math.min(Math.cos(a) * r0, Math.cos(a) * r1),
        y: cy + Math.min(Math.sin(a) * r0, Math.sin(a) * r1),
        w: Math.abs(Math.cos(a)) * (r1 - r0), h: Math.abs(Math.sin(a)) * (r1 - r0),
        flipH: Math.cos(a) * Math.sin(a) < 0, line: stroke(ink),
      });
    }
  },
  smile(s, x, y, u, ink) {
    s.addShape('hexagon', { x: x + 0.05 * u, y: y + 0.05 * u, w: 0.9 * u, h: 0.9 * u, fill: { type: 'none' }, line: stroke(ink), rotate: 90 });
    s.addShape('blockArc', { x: x + 0.26 * u, y: y + 0.3 * u, w: 0.18 * u, h: 0.14 * u, angleRange: [0, 180], arcThicknessRatio: 0.2, fill: { color: ink }, line: NOLINE });
    s.addShape('blockArc', { x: x + 0.56 * u, y: y + 0.3 * u, w: 0.18 * u, h: 0.14 * u, angleRange: [0, 180], arcThicknessRatio: 0.2, fill: { color: ink }, line: NOLINE });
    s.addShape('blockArc', { x: x + 0.28 * u, y: y + 0.4 * u, w: 0.44 * u, h: 0.3 * u, angleRange: [0, 180], arcThicknessRatio: 0.14, fill: { color: ink }, line: NOLINE });
  },
  frown(s, x, y, u, ink) {
    s.addShape('hexagon', { x: x + 0.05 * u, y: y + 0.05 * u, w: 0.9 * u, h: 0.9 * u, fill: { type: 'none' }, line: stroke(ink), rotate: 90 });
    s.addShape('line', { x: x + 0.26 * u, y: y + 0.36 * u, w: 0.14 * u, h: 0, line: stroke(ink) });
    s.addShape('line', { x: x + 0.6 * u, y: y + 0.36 * u, w: 0.14 * u, h: 0, line: stroke(ink) });
    s.addShape('blockArc', { x: x + 0.28 * u, y: y + 0.52 * u, w: 0.44 * u, h: 0.3 * u, angleRange: [180, 0], arcThicknessRatio: 0.14, fill: { color: ink }, line: NOLINE });
  },
  hand(s, x, y, u, ink) {
    s.addShape('roundRect', { x: x + 0.24 * u, y: y + 0.1 * u, w: 0.52 * u, h: 0.62 * u, rectRadius: 0.02, fill: { type: 'none' }, line: stroke(ink) });
    [0.34, 0.46, 0.58].forEach((dx) =>
      s.addShape('line', { x: x + dx * u, y: y + 0.1 * u, w: 0, h: 0.3 * u, line: stroke(ink) }));
    s.addShape('line', { x: x + 0.24 * u, y: y + 0.72 * u, w: 0.52 * u, h: 0.16 * u, line: stroke(ink) });
  },
  camp(s, x, y, u, ink) {
    s.addShape('triangle', { x: x + 0.06 * u, y: y + 0.12 * u, w: 0.88 * u, h: 0.7 * u, fill: { type: 'none' }, line: stroke(ink) });
    s.addShape('line', { x: x + 0.5 * u, y: y + 0.12 * u, w: 0, h: 0.7 * u, line: stroke(ink) });
    s.addShape('line', { x: x + 0.06 * u, y: y + 0.86 * u, w: 0.88 * u, h: 0, line: stroke(ink) });
  },
  flag(s, x, y, u, ink) {
    s.addShape('line', { x: x + 0.24 * u, y: y + 0.06 * u, w: 0, h: 0.88 * u, line: stroke(ink) });
    s.addShape('triangle', { x: x + 0.24 * u, y: y + 0.1 * u, w: 0.52 * u, h: 0.36 * u, fill: { type: 'none' }, line: stroke(ink), rotate: 90 });
  },
};

/** Filled square tile holding a line icon (the deck's recurring feature chip). */
function tile(s, x, y, size, bg, glyph, inkOverride) {
  rect(s, x, y, size, size, bg);
  const ink = inkOverride || (bg === LIME ? TEAL : WHITE);
  ICONS[glyph](s, x + size * 0.2, y + size * 0.2, size * 0.6, ink);
}

/** The four-chip row (mountain / tree / pins / sun) used on several slides. */
function chipRow(s, x, y, size, gap) {
  const cfg = [[TEAL, 'mountain'], [LIME, 'tree'], [TEAL, 'pins'], [LIME, 'sun']];
  cfg.forEach(([bg, glyph], i) => tile(s, x + i * (size + gap), y, size, bg, glyph));
}

/* --------------------------------------------------- recurring composites   */

/** Smart-phone mockup (slides 11 & 27) - black body, grey screen. */
function phoneMockup(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.42, fill: { color: '111111' }, line: NOLINE });
  s.addShape('roundRect', { x: x + 0.02, y: y + 0.02, w: w - 0.04, h: h - 0.04, rectRadius: 0.4, fill: { color: '333333' }, line: NOLINE });
  s.addShape('roundRect', { x: x + 0.1, y: y + 0.1, w: w - 0.2, h: h - 0.2, rectRadius: 0.34, fill: { color: '000000' }, line: NOLINE });
  photo(s, x + 0.204, y + 0.183, w - 0.41, h - 0.366);
  s.addShape('roundRect', { x: x + w / 2 - 0.6, y: y + 0.1, w: 1.2, h: 0.22, rectRadius: 0.11, fill: { color: '000000' }, line: NOLINE });
  s.addShape('roundRect', { x: x + w / 2 - 0.21, y: y + 0.155, w: 0.42, h: 0.04, rectRadius: 0.02, fill: { color: '171012' }, line: NOLINE });
  // volume + power keys
  [[0.85, 0.24], [1.31, 0.45], [1.86, 0.45]].forEach(([dy, kh]) =>
    s.addShape('roundRect', { x: x - 0.014, y: y + dy, w: 0.031, h: kh, rectRadius: 0.015, fill: { color: '868686' }, line: NOLINE }));
  s.addShape('roundRect', { x: x + w - 0.017, y: y + 1.456, w: 0.024, h: 0.73, rectRadius: 0.012, fill: { color: '868686' }, line: NOLINE });
}

/** Laptop mockup (slides 12 & 28) - rounded lid, screen, tapered base. */
function laptopMockup(s, x, y, w, h) {
  const baseH = h * 0.035;
  const lidW = w * 0.816;
  const lidX = x + (w - lidW) / 2;
  const lidH = h - baseH - h * 0.008;
  s.addShape('roundRect', { x: lidX, y, w: lidW, h: lidH, rectRadius: 0.09, fill: { color: 'B0B2B4' }, line: NOLINE });
  s.addShape('roundRect', { x: lidX + 0.02, y: y + 0.02, w: lidW - 0.04, h: lidH - 0.04, rectRadius: 0.08, fill: { color: '111111' }, line: NOLINE });
  const sx = lidX + w * 0.026;
  const sy = y + h * 0.026;
  const sw = lidW - w * 0.052;
  const sh = lidH - h * 0.052;
  photo(s, sx, sy, sw, sh);
  shade(s, sx, sy + sh * 0.55, sw, sh * 0.45);
  poly(s, x, y + lidH + h * 0.008, w, baseH,
    [{ x: w * 0.05, y: 0 }, { x: w * 0.95, y: 0 }, { x: w, y: baseH }, { x: 0, y: baseH }, { close: true }],
    { color: 'C9CACB' });
  rect(s, x + w * 0.405, y + lidH + h * 0.008, w * 0.19, baseH * 0.35, '96999A');
}

/** Bell-shaped "mountain" area used by the hand-drawn chart on slides 15 & 29. */
function bell(s, x, y, w, h, color, transparency) {
  poly(s, x, y, w, h, [
    { x: 0, y: h },
    { x: w * 0.505, y: 0, curve: { type: 'cubic', x1: w * 0.167, y1: h * 0.777, x2: w * 0.369, y2: 0 } },
    { x: w, y: h, curve: { type: 'cubic', x1: w * 0.641, y1: 0, x2: w * 0.833, y2: h * 0.777 } },
    { close: true },
  ], { color, transparency });
}

/** Full "mountain chart" block: 0-100% axis labels + five overlapping bells. */
function mountainChart(s, x, y) {
  ['100%', '80%', '60%', '40%', '20%', '0%'].forEach((lbl, i) =>
    text(s, lbl, { x, y: y + i * 0.616, w: 0.659, h: 0.325, fontSize: 10, color: '808080', align: 'right', valign: 'middle', margin: TM }));
  const bells = [
    [0.659, 0.326, 3.081, LIME, 20],
    [1.6, 1.353, 2.054, TEAL, 20],
    [2.541, 0.0, 3.406, 'BFBFBF', 20],
    [3.483, 2.132, 1.275, TEAL, 20],
    [4.424, 1.37, 2.037, 'BFBFBF', 20],
  ];
  bells.forEach(([dx, dy, bh, color, tr]) => bell(s, x + dx, y + dy, 1.718, bh, color, tr));
}

/** Two-column "The future of marketing / Embrace the change" cards. */
function markingCards(s, x) {
  const cards = [
    [3.025, 'The future of marketing'],
    [4.676, 'Embrace the change'],
  ];
  cards.forEach(([cy, heading]) => {
    rect(s, x, cy, 5.013, 1.549, LIME);
    text(s, heading, { x: x + 0.25, y: cy + 0.273, w: 4.512, h: 0.343, fontSize: 16, bold: true, color: CYAN });
    text(s, FUSCE_CARD, { x: x + 0.25, y: cy + 0.533, w: 4.512, h: 0.828, fontSize: 12, color: DARK, fontFace: LIGHT, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
  });
  text(s, 'Efficiency In Reach', { x, y: 1.035, w: 5.013, h: 0.404, fontSize: 24, bold: true, color: CYAN, valign: 'middle', margin: TM });
  para(s, LOREM_SHORT + ' ', x, 2.071, 5.013, 0.833);
}

/* -------------------------------------------------------------- the slides  */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: HEAD, bodyFontFace: LIGHT };

/* 1 - cover ---------------------------------------------------------------- */
function slide01() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 13.333, 7.5, { label: false });
  rect(s, 0, 0, 13.333, 7.5, '0B2036', { fill: { color: '0B2036', transparency: 75 } });
  poly(s, 0, -0.004, 13.333, 5.437, [
    { x: 0, y: 0 }, { x: 13.333, y: 0 }, { x: 13.333, y: 4.162 }, { x: 0, y: 4.162 }, { close: true },
    { x: 0, y: 4.164, moveTo: true }, { x: 13.333, y: 4.164 }, { x: 6.687, y: 5.437 }, { close: true },
  ], { color: TEAL, transparency: 75 });
  chevron(s, 4.717, 2.783, 1.289, LIME, null, 72);
  mark(s, 6.285, 2.364, 0.764, 0.632);
  text(s, 'A D V E N T R A I L', { x: 3.02, y: 3.148, w: 7.294, h: 1.01, fontSize: 54, bold: true, color: LIME, align: 'center', valign: 'middle', margin: TM });
  text(s, TAGLINE, { x: 3.75, y: 4.28, w: 5.833, h: 0.656, fontSize: 11, color: WHITE, align: 'center', valign: 'middle', margin: TM });
  logo(s, 5.906, 0.316);
}

/* 2 / 3 - "photo on one side" intro slides ---------------------------------- */
function introSlide(headLines, photoX, textX, logoX, chipY) {
  const s = pptx.addSlide();
  photo(s, photoX, 0, 5.486, 7.5);
  shade(s, photoX, 3.931, 5.486, 3.569);
  title(s, headLines, textX, chipY.title, 6.021, 1.212);
  para(s, LOREM_FULL, textX, chipY.body, 6.021, 1.079);
  logo(s, logoX, 0.461);
  return s;
}

function slide02() {
  const s = introSlide('Great Adventures\nBegin Here', 7.847, 1.003, 0.574, { title: 1.821, body: 3.085 });
  tile(s, 1.082, 4.444, 0.65, TEAL, 'pins', LIME);
  text(s, 'Short Title Here', { x: 1.868, y: 4.355, w: 3.519, h: 0.269, fontSize: 12, bold: true, color: TEAL, lineSpacingMultiple: 1.5, valign: 'top', margin: TM });
  para(s, LOREM_SHORT, 1.868, 4.656, 3.909, 0.757, { fontSize: 10 });
}

function slide03() {
  const s = introSlide('Good Experiences\nStart Here', 0, 6.288, 11.181, { title: 1.472, body: 2.736 });
  chipRow(s, 6.405, 4.125, 0.65, 0.229);
  text(s, 'Short Title Here', { x: 6.288, y: 5.053, w: 3.519, h: 0.269, fontSize: 12, bold: true, color: TEAL, lineSpacingMultiple: 1.5, valign: 'top', margin: TM });
  para(s, LOREM_SHORT, 6.288, 5.353, 3.909, 0.757, { fontSize: 10 });
}

/* numbered 1-4 list, reused by slides 4 and 23 ------------------------------ */
const NUMBERED = [L2, UTEU, L2, UTEU];
function numberedList(s, boxX, textX, y0, step) {
  NUMBERED.forEach((body, i) => {
    const yy = y0 + i * step;
    rect(s, boxX, yy, 0.581, 0.581, i % 2 ? LIME : TEAL);
    text(s, String(i + 1), { x: boxX, y: yy, w: 0.581, h: 0.581, fontSize: 16, bold: true, align: 'center', valign: 'middle', color: i % 2 ? TEAL : WHITE });
    para(s, body, textX, yy + (i ? 0.012 : 0), 4.489, 0.505, { fontSize: 10 });
  });
}

/* 4 - Explore The Wild + numbered list -------------------------------------- */
function slide04() {
  const s = pptx.addSlide();
  rect(s, 10.146, 0, 3.188, 7.5, LIME);
  rect(s, 10.028, 0, 0.118, 7.5, TEAL);
  title(s, 'Explore The Wild', 1.003, 1.572, 6.021, 0.606);
  para(s, LOREM_FULL, 1.003, 2.214, 6.021, 1.079);
  numberedList(s, 1.067, 1.795, 3.641, 0.7165);
  photo(s, 8.052, 0.64, 4.187, 6.221);
  shade(s, 8.052, 3.972, 4.187, 2.886);
  logo(s, 0.574, 0.461);
}

/* 5 - Team Adventure -------------------------------------------------------- */
function slide05() {
  const s = pptx.addSlide();
  photo(s, 0, 3.615, 13.333, 3.885);
  rect(s, 0, 3.615, 13.333, 3.885, TEAL, { fill: { color: TEAL, transparency: 55 } });
  [[8.549, 'Team Leader', 8.237], [10.413, 'Co-Leader', 10.102]].forEach(([cx, role, lx]) => {
    s.addShape('ellipse', { x: cx, y: 1.179, w: 1.434, h: 1.434, fill: { color: LIME }, line: NOLINE });
    s.addShape('ellipse', { x: cx + 0.05, y: 1.229, w: 1.333, h: 1.333, fill: { color: PHOTO }, line: NOLINE });
    text(s, role, { x: lx, y: 2.752, w: 2.058, h: 0.236, fontSize: 14, color: TEAL, align: 'center', valign: 'middle', margin: TM });
  });
  title(s, 'Team Adventure', 0.968, 1.269, 5.171, 0.606);
  para(s, LOREM_MED, 0.968, 1.91, 5.921, 1.111);
  text(s, 'E X P E D I T I O N S', { x: 1.948, y: 5.137, w: 9.438, h: 0.841, fontFace: BLACK, fontSize: 50, bold: true, color: LIME, align: 'center', valign: 'middle', margin: TM });
  logo(s, 0.574, 0.461);
}

/* 6 - Team Experiences ------------------------------------------------------ */
function slide06() {
  const s = pptx.addSlide();
  rect(s, 0.003, 1.956, 1.0, 3.754, LIME);
  photo(s, 1.003, 1.956, 5.629, 3.754);
  shade(s, 1.003, 4.264, 5.629, 1.447);
  title(s, 'Team\nExperiences', 7.253, 1.889, 4.775, 1.212);
  para(s, LOREM_FULL, 7.253, 3.189, 4.775, 1.388);
  tile(s, 7.336, 4.809, 0.902, TEAL, 'smile');
  tile(s, 8.384, 4.809, 0.902, LIME, 'frown');
  logo(s, 0.574, 0.461);
}

/* 7 - photo grid ------------------------------------------------------------ */
function slide07() {
  const s = pptx.addSlide();
  chevron(s, 5.491, 2.009, 0.93, LIME, CYAN);
  title(s, 'Enjoy The Adventure Experience', 0.968, 0.73, 9.096, 0.539, { fontSize: 32 });
  photo(s, 0.968, 1.589, 7.574, 5.052);
  shade(s, 0.968, 5.491, 7.574, 1.15);
  photo(s, 8.613, 1.589, 3.787, 2.491);
  shade(s, 8.613, 2.931, 3.787, 1.15);
  photo(s, 8.613, 4.15, 3.787, 2.491);
  shade(s, 8.613, 5.491, 3.787, 1.15);
  logo(s, 11.181, 0.461);
}

/* 8 - 2020 ------------------------------------------------------------------ */
function slide08() {
  const s = pptx.addSlide();
  rect(s, 0.003, 1.643, 4.355, 4.798, LIME);
  text(s, '2020', { x: 0.937, y: 2.098, w: 2.843, h: 1.01, fontSize: 60, bold: true, color: TEAL, align: 'right', valign: 'middle', margin: TM });
  para(s, LOREM_FULL, 0.937, 3.107, 2.843, 2.777, { align: 'right' });
  [4.49, 8.978].forEach((x) => {
    photo(s, x, 1.643, 4.355, 4.798);
    shade(s, x, 4.514, 4.355, 1.927);
  });
  logo(s, 5.906, 0.461);
}

/* 9 - Outdoor Life ---------------------------------------------------------- */
function slide09() {
  const s = pptx.addSlide();
  const slash = [{ x: 0, y: 0 }, { x: 1.038, y: 0 }, { x: 6.205, y: 7.5 }, { x: 1.038, y: 7.5 }, { x: 1.038, y: 7.5 }, { x: 0, y: 7.5 }, { close: true }];
  poly(s, 0.38, 0, 6.205, 7.5, slash, { color: TEAL });
  poly(s, 0, 0, 6.205, 7.5, slash, { color: LIME });
  title(s, 'Outdoor Life', 6.516, 1.081, 4.089, 0.606);
  para(s, '"Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, ' +
    'totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta ' +
    'sunt explicabo. ' + L2 + L3, 6.516, 1.757, 4.089, 1.912);
  [[1.848, 0.981], [1.856, 4.035], [6.334, 4.034]].forEach(([x, y]) => {
    photo(s, x, y, 4.35, 2.902);
    shade(s, x, y + 1.342, 4.35, 1.56);
  });
  logo(s, 10.936, 0.461);
}

/* 10 - Rumi quote ----------------------------------------------------------- */
function slide10() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 13.333, 7.5, { label: false });
  shade(s, 0, 0, 13.333, 6.486, { from: '0F8FA0', to: '0BD0D9', top: 70, bottom: 0, steps: 14 });
  poly(s, 0, -0.004, 13.333, 5.437, [
    { x: 0, y: 0 }, { x: 13.333, y: 0 }, { x: 13.333, y: 4.162 }, { x: 0, y: 4.162 }, { close: true },
    { x: 0, y: 4.164, moveTo: true }, { x: 13.333, y: 4.164 }, { x: 6.687, y: 5.437 }, { close: true },
  ], { color: TEAL, transparency: 90 });
  text(s, ' \u201CRun from what\u2019s comfortable. Forget safety.\nLive where you fear to live. Destroy your reputation.\n' +
    'Be notorious. I have tried prudent planning long enough.\nFrom now on I\u2019ll be mad.\u201D',
    { x: 1.125, y: 1.978, w: 11.083, h: 1.703, fontSize: 26, bold: true, color: LIME, align: 'center' });
  s.addShape('roundRect', { x: 5.986, y: 4.019, w: 1.361, h: 0.471, rectRadius: 0.235, fill: { color: TEAL, transparency: 25 }, line: NOLINE });
  text(s, '\u2013 Rumi', { x: 5.986, y: 4.019, w: 1.361, h: 0.471, fontSize: 16, bold: true, italic: true, color: LIME, align: 'center', valign: 'middle' });
}

/* 11 - phone mockup + lime cards -------------------------------------------- */
function slide11() {
  const s = pptx.addSlide();
  rect(s, 0, 1.495, 5.542, 4.51, TEAL);
  [1.495, 3.782].forEach((y) => rect(s, 7.792, y, 5.542, 2.223, LIME));
  text(s, 'Explore The Wild', { x: 0.847, y: 1.88, w: 3.833, h: 1.616, fontSize: 48, bold: true, color: LIME, align: 'right', valign: 'middle', margin: TM });
  para(s, LOREM_FULL, 0.847, 3.542, 3.833, 1.944, { align: 'right', color: SOFT });
  [1.718, 3.931].forEach((y) => {
    text(s, 'Maintain performance', { x: 8.629, y, w: 4.299, h: 0.374, fontSize: 18, bold: true, color: TEAL });
    text(s, LOREM_MED, { x: 8.629, y: y + 0.344, w: 4.299, h: 1.212, fontSize: 11, fontFace: LIGHT, color: DARK, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
  });
  phoneMockup(s, 5.085, 0.598, 3.163, 6.304);
  logo(s, 10.936, 0.461);
}

/* 12 - laptop hero ---------------------------------------------------------- */
function slide12() {
  const s = pptx.addSlide();
  chevron(s, 4.681, 2.734, 1.266, LIME);
  chevron(s, 4.766, 2.734, 1.266, '5CA058');
  title(s, 'Great Adventures Begin Here', 1.67, 1.011, 9.993, 0.606, { align: 'center' });
  para(s, LOREM_MED, 1.67, 1.695, 9.993, 0.555, { align: 'center' });
  laptopMockup(s, 3.194, 2.539, 6.944, 4.009);
  logo(s, 5.906, 0.461);
}

/* 13 / 14 - vertical timelines ---------------------------------------------- */
const TIMELINE = [
  ['6.00 AM', 'START', 'pins'],
  ['12.00 AM', 'PLACEHOLDER', 'tree'],
  ['3.00 AM', 'Ut enim ad minima veniam, quis nostrum exercitationem ullam corporis suscipi', 'camp'],
  ['6.00 PM', 'REST', 'hand'],
  ['8.00 PM', 'PLACEHOLDER', 'mountain'],
  ['9.00 PM', 'CHECKPOINT', 'flag'],
];

function timelineSlide(dayLabel, dayColor, lineColor, ringColor, headLines) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 5.019, y: 0, w: 0, h: 0.709, line: { color: lineColor, width: 1 } });
  s.addShape('line', { x: 5.019, y: 1.17, w: 0, h: 6.33, line: { color: lineColor, width: 1 } });
  text(s, dayLabel, { x: 4.589, y: 0.773, w: 0.861, h: 0.334, fontSize: 14, bold: true, color: dayColor, align: 'center', valign: 'middle', margin: TM });
  TIMELINE.forEach(([time, label, glyph], i) => {
    const y = 1.572 + i * 0.913;
    s.addShape('ellipse', { x: 4.769, y, w: 0.499, h: 0.499, fill: { color: WHITE }, line: { color: ringColor, width: 0.75 } });
    ICONS[glyph](s, 4.889, y + 0.12, 0.26, TEAL);
    text(s, time, { x: 5.345, y: y + 0.073, w: 1.08, h: 0.353, fontSize: 14, color: TEAL, align: 'right', lineSpacingMultiple: 1.5, valign: 'middle', margin: TM });
    para(s, label, 1.196, y + 0.073, 3.37, 0.37, { align: 'right', valign: 'middle', lineSpacingMultiple: 1 });
  });
  title(s, headLines, 7.67, 2.65, 4.299, 1.212);
  para(s, L1, 7.67, 3.977, 4.299, 0.833);
  ICONS.pins(s, 10.774, 1.583, 1.063, TEAL, { solid: true, second: LIME });
  logo(s, 0.574, 0.461);
}

const slide13 = () => timelineSlide('Day 1', GREEN, GREEN, TEAL, 'Trip\nSchedules');
const slide14 = () => timelineSlide('Day 2', TEAL, TEAL, GREEN, 'Time\nManagement');

/* 15 / 16 - teal panel + chart ---------------------------------------------- */
function panelSlide(heading, drawChart) {
  const s = pptx.addSlide();
  // 315-degree gradient: teal at the top, cyan towards the bottom
  for (let i = 0; i < 14; i++) rect(s, 0, (7.5 * i) / 14, 5.542, 7.5 / 14 + 0.01, mix(TEAL, CYAN, (i + 0.5) / 14));
  text(s, 'Efficiency In Reach', { x: 0.801, y: 1.688, w: 4.059, h: 0.337, fontSize: 20, bold: true, color: LIME, valign: 'middle', margin: TM });
  text(s, heading, { x: 0.801, y: 2.122, w: 4.059, h: 0.471, fontSize: 28, bold: true, color: WHITE, valign: 'middle', margin: TM });
  para(s, FUSCE_LONG, 0.801, 2.691, 4.059, 2.121, { fontSize: 12, color: WHITE });
  chipRow(s, 0.801, 5.431, 0.65, 0.228);
  drawChart(s);
  text(s, F1, { x: 6.647, y: 5.702, w: 5.875, h: 0.545, fontSize: 11, fontFace: LIGHT, color: TEAL, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
  logo(s, 11.181, 0.461);
}

const slide15 = () => panelSlide('Team Management', (s) => mountainChart(s, 6.394, 2.024));
const slide16 = () => panelSlide('Time Management', (s) => barChart(s, 6.633, 2.016));

/** Clustered column chart shared by slides 16 and 30. */
function barChart(s, x, y) {
  const data = [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2, 2, 3, 5] },
  ];
  s.addChart(pptx.ChartType.bar, data, {
    x, y, w: 5.875, h: 3.526,
    barDir: 'col', barGapWidthPct: 80, barOverlapPct: 25,
    chartColors: [TEAL, CYAN, LIME],
    showLegend: false, showTitle: false,
    catAxisHidden: true, catAxisLineShow: false,
    valAxisLineShow: false, valAxisLabelColor: '808080', valAxisLabelFontSize: 12,
    valAxisLabelFontFace: LIGHT, valGridLine: { color: 'F2F2F2', size: 0.75 }, catGridLine: { style: 'none' },
  });
}

/* 17 - Estimation donut + progress bars ------------------------------------- */
function slide17() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 3.188, 7.5, LIME);
  s.addShape('ellipse', { x: 1.104, y: 1.667, w: 4.167, h: 4.167, fill: { color: WHITE }, line: NOLINE });
  s.addShape('donut', { x: 1.333, y: 1.896, w: 3.708, h: 3.708, fill: { color: TEAL }, line: NOLINE });
  [[CYAN, [145, 216]], [GRASS, [90, 146]], [GREEN, [34, 90]]].forEach(([color, range]) =>
    s.addShape('blockArc', { x: 1.333, y: 1.896, w: 3.708, h: 3.708, angleRange: range, arcThicknessRatio: 0.4, fill: { color }, line: NOLINE }));
  s.addShape('ellipse', { x: 2.355, y: 2.918, w: 1.664, h: 1.664, fill: { color: LIME }, line: { color: TEAL, width: 1, dashType: 'dash' } });
  ICONS.tree(s, 2.833, 3.361, 0.709, TEAL);
  [['50% ', 3.815, 2.531], ['15% ', 3.477, 4.947], ['15% ', 2.23, 4.947], ['20% ', 1.405, 3.632]].forEach(([t, x, y]) =>
    text(s, t, { x, y, w: 0.643, h: 0.236, fontSize: 14, color: WHITE, align: 'center', valign: 'middle', margin: TM }));

  title(s, 'Estimation', 5.586, 2.161, 6.021, 0.606);
  para(s, L1, 5.586, 2.834, 6.424, 0.555);
  const bars = [[3.78, 3.746, TEAL], [4.227, 2.853, CYAN], [4.673, 1.97, GREEN], [5.12, 1.97, GRASS]];
  bars.forEach(([y, len, color], i) => {
    text(s, 'Day ' + (i + 1), { x: 5.586, y: y - 0.009, w: 0.738, h: 0.202, fontSize: 12, color: DARK, valign: 'middle', margin: TM });
    s.addShape('roundRect', { x: 6.253, y: y - 0.004, w: 5.512, h: 0.176, rectRadius: 0.088, fill: { color: SMOKE }, line: NOLINE });
    s.addShape('roundRect', { x: 6.253, y, w: len, h: 0.176, rectRadius: 0.088, fill: { color }, line: NOLINE });
  });
  logo(s, 11.181, 0.461);
}

/* 18 / 31 - schedule table -------------------------------------------------- */
const TABLE_ROWS = [
  'Aliquam sed ipsum maximus ',
  'Maecenas consequat elementum ',
  'Vivamus interdum porta ',
  'Quisque ut nibh ante',
  'Orci varius natoque penatibus ',
  'Aenean rhoncus tempus erat',
];

function scheduleTable(s, ruleColor, headFill) {
  const rule = { type: 'solid', color: ruleColor, pt: 0.5 };
  const none = { type: 'none' };
  const head = (txt, fill, color, side) => ({
    text: txt,
    options: { fill: fill ? { color: fill } : undefined, color, fontSize: 12, align: 'center', valign: 'middle',
      border: [none, side === 'l' ? rule : none, rule, side === 'l' ? none : rule] },
  });
  const rows = [[
    head('SCHEDULES', headFill && headFill[0], headFill ? WHITE && (headFill[0] === LIME ? TEAL : WHITE) : TEAL, 'l'),
    head('DAY', headFill && headFill[1], headFill ? WHITE : TEAL, 'r'),
  ]];
  TABLE_ROWS.forEach((label, i) => {
    const last = i === TABLE_ROWS.length - 1;
    rows.push([
      { text: label, options: { color: DARK, fontSize: 11, valign: 'middle', border: [rule, rule, last ? none : rule, none] } },
      { text: 'Passed', options: { color: DARK, fontSize: 11, align: 'center', valign: 'middle', border: [rule, none, last ? none : rule, rule] } },
    ]);
  });
  s.addTable(rows, {
    x: 1.125, y: 2.671, colW: [3.268, 1.482], rowH: 0.559,
    fontFace: LIGHT, margin: [3.6, 7.2, 3.6, 7.2],
  });
}

function tableSlide(kicker, heading, headFill, ruleColor, headY, bodyY, capBold, capSize) {
  const s = pptx.addSlide();
  if (kicker) text(s, kicker, { x: 1.056, y: 0.774, w: 5.013, h: 0.404, fontSize: 24, bold: true, color: CYAN, valign: 'middle', margin: TM });
  title(s, heading, 1.056, headY, 4.819, 0.606);
  para(s, LOREM_HALF, 1.056, bodyY, 4.819, 0.555);
  scheduleTable(s, ruleColor, headFill);
  photo(s, 6.667, 1.198, 5.542, 4.062);
  shade(s, 6.667, 3.814, 5.533, 1.447);
  text(s, 'Short Title Here', { x: 6.708, y: capBold, w: 5.492, h: 0.314, fontSize: capSize, bold: true, color: TEAL, lineSpacingMultiple: 1.5, valign: 'top', margin: TM });
  para(s, LOREM_SHORT, 6.667, capBold + 0.385, 5.542, 0.757, { fontSize: capSize === 14 ? 10.5 : 10 });
  logo(s, 11.181, 0.461);
}

const slide18 = () => tableSlide(null, 'Enjoy The Trip', null, LIME, 1.198, 1.843, 5.447, 12);
const slide31 = () => tableSlide('Efficiency In Reach', 'Explore The Wild', [LIME, TEAL], TEAL, 1.248, 1.923, 5.437, 14);

/* 19 - four grey subtitle panels around a centre photo ---------------------- */
function slide19() {
  const s = pptx.addSlide();
  rect(s, 0, 1.147, 13.333, 2.814, SMOKE);
  rect(s, 0, 4.073, 13.333, 2.814, SMOKE);
  photo(s, 4.333, 0, 4.667, 7.5);
  shade(s, 4.333, 3.889, 4.667, 3.611);
  const panels = [
    { x: 0.674, w: 3.187, align: 'right', titleY: 2.019, bodyY: 2.319, color: OLIVE, tileX: 3.275, tileY: 1.434, bg: LIME, glyph: 'tree' },
    { x: 9.472, w: 3.23, align: 'left', titleY: 2.019, bodyY: 2.319, color: TEAL, tileX: 9.569, tileY: 1.434, bg: TEAL, glyph: 'pins' },
    { x: 0.674, w: 3.187, align: 'right', titleY: 4.916, bodyY: 5.204, color: TEAL, tileX: 3.275, tileY: 4.333, bg: TEAL, glyph: 'camp' },
    { x: 9.472, w: 3.23, align: 'left', titleY: 4.916, bodyY: 5.204, color: OLIVE, tileX: 9.569, tileY: 4.333, bg: LIME, glyph: 'sun' },
  ];
  panels.forEach((p) => {
    text(s, 'Subtitle Here', { x: p.x, y: p.titleY, w: p.w, h: 0.343, fontSize: 16, bold: true, color: p.color, align: p.align });
    text(s, FUSCE_MED, { x: p.x, y: p.bodyY, w: p.w, h: 1.212, fontSize: 11, fontFace: LIGHT, color: DARK, align: p.align, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
    tile(s, p.tileX, p.tileY, 0.489, p.bg, p.glyph);
  });
  logo(s, 11.181, 0.461);
}

/* 20 - Our Team ------------------------------------------------------------- */
const TEAM = [
  ['Don Zach', 'Team Leader'],
  ['Diana Marz', 'Co-Leader'],
  ['Lily Starr', 'Co Leader'],
  ['Jack Mode', 'Planner'],
];

function slide20() {
  const s = pptx.addSlide();
  title(s, 'Our Team', 1.794, 0.796, 9.858, 0.606, { align: 'center' });
  para(s, LOREM_MED, 1.794, 1.438, 9.858, 0.555, { align: 'center' });
  TEAM.forEach(([name, role], i) => {
    const x = 1.512 + i * 2.694;
    s.addShape('pie', { x, y: 2.363, w: 2.201, h: 2.201, angleRange: [0, 315], fill: { color: LIME }, line: NOLINE });
    s.addShape('ellipse', { x: x + 0.143, y: 2.506, w: 1.915, h: 1.915, fill: { color: PHOTO }, line: NOLINE });
    text(s, name, { x: x + 0.143, y: 4.817, w: 1.915, h: 0.269, fontSize: 16, bold: true, color: TEAL, align: 'center', valign: 'middle', margin: TM });
    text(s, role, { x: x + 0.143, y: 5.132, w: 1.915, h: 0.236, fontSize: 14, color: CYAN, align: 'center', valign: 'middle', margin: TM });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.',
      { x, y: 5.412, w: 2.201, h: 0.909, fontSize: 10, fontFace: LIGHT, color: DARK, align: 'center', lineSpacingMultiple: 1.2, margin: [0, 0, 3.6, 3.6] });
  });
  logo(s, 11.181, 0.461);
}

/* 21 - Exciting The Wild (teal photo panel left) ---------------------------- */
function slide21() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 6.667, 7.5, TEAL, { fill: { color: TEAL, transparency: 25 } });
  text(s, 'Exciting The Wild', { x: 1.181, y: 1.88, w: 4.069, h: 1.616, fontSize: 48, bold: true, color: LIME, align: 'right', valign: 'middle', margin: TM });
  para(s, LOREM_FULL, 1.181, 3.542, 4.069, 2.121, { fontSize: 12, color: SOFT, align: 'right' });
  title(s, 'Good Experiences\nStart Here', 7.427, 1.166, 4.99, 1.212);
  para(s, LOREM_FULL, 7.427, 2.43, 4.99, 1.388);
  chipRow(s, 7.544, 4.125, 0.65, 0.229);
  text(s, 'Subtitle Here', { x: 7.427, y: 5.053, w: 3.519, h: 0.303, fontSize: 12, bold: true, color: TEAL, lineSpacingMultiple: 1.5, valign: 'top', margin: TM });
  para(s, LOREM_SHORT, 7.427, 5.353, 3.909, 0.757, { fontSize: 10 });
  logo(s, 11.181, 0.461);
}

/* 22 - two photos with captions --------------------------------------------- */
function slide22() {
  const s = pptx.addSlide();
  title(s, 'Enjoy The Adventure Experience', 1.139, 0.935, 9.681, 0.606);
  para(s, LOREM_MED, 1.139, 1.579, 9.681, 0.555);
  [[1.139, TEAL], [6.847, OLIVE]].forEach(([x, color]) => {
    photo(s, x, 2.361, 5.431, 3.07);
    shade(s, x, 3.986, 5.431, 1.445);
    text(s, 'Subtitle Here', { x, y: 5.628, w: 5.431, h: 0.343, fontSize: 16, bold: true, color });
    text(s, FUSCE_MED, { x, y: 5.972, w: 5.431, h: 0.767, fontSize: 11, fontFace: LIGHT, color: DARK, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
  });
  logo(s, 11.181, 0.461);
}

/* 23 - Enjoy The Wild (diagonal photo right) -------------------------------- */
function slide23() {
  const s = pptx.addSlide();
  photo(s, 6.125, 0, 7.208, 7.5, {
    points: [{ x: 0, y: 0 }, { x: 7.208, y: 0 }, { x: 7.208, y: 7.5 }, { x: 1.444, y: 7.5 }, { close: true }],
  });
  shade(s, 6.962, 4.255, 5.846, 3.245, { slantTop: 0, slantBottom: 0.618 });
  mark(s, 9.107, 2.81, 2.273, 1.881, { transparency: 35, limeTransparency: 45 });
  title(s, 'Enjoy The Wild', 0.871, 1.299, 4.934, 0.606);
  para(s, LOREM_FULL, 0.871, 2.029, 4.934, 1.388, { color: TEAL });
  numberedList(s, 0.936, 1.664, 3.75, 0.7165);
  logo(s, 0.574, 0.461);
}

/* 24 - photo banner + two columns + footer ---------------------------------- */
function slide24() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 13.333, 3.07, { label: false });
  shade(s, 0, 0.714, 13.333, 2.356);
  rect(s, 0, 3.07, 13.333, 0.05, LIME);
  rect(s, 0, 6.27, 13.333, 1.23, TEAL);
  text(s, 'Exciting The Wild', { x: 1.968, y: 1.098, w: 9.397, h: 0.404, fontSize: 24, bold: true, color: LIME, align: 'center', valign: 'middle', margin: TM });
  text(s, 'Enjoy The Adventure Experience', { x: 1.968, y: 1.502, w: 9.397, h: 0.673, fontSize: 40, color: WHITE, align: 'center', valign: 'middle', margin: TM });
  [['Team Experiences', 1.141], ['Tour Experiences', 7.443]].forEach(([heading, x]) => {
    text(s, heading, { x, y: 3.859, w: 4.748, h: 0.404, fontSize: 24, bold: true, color: DARK, valign: 'middle', margin: TM });
    para(s, L1 + 'Orci varius natoque penatibus et magnis dis parturient montes, ', x, 4.399, 4.748, 1.111);
  });
  s.addShape('line', { x: 6.667, y: 4.01, w: 0, h: 1.5, line: { color: LIME, width: 1 } });
  text(s, 'Company Adventure Lorem ipsum dolor sit amet, consectetur adipiscing elit.\n' +
    'Ut euismod,  est eget maximus venenatis, augue justo tincidunt massa, a rhoncus tortor lectus ac tortor. ',
    { x: 1.316, y: 6.649, w: 10.717, h: 0.471, fontSize: 14, fontFace: LIGHT, color: WHITE, align: 'center', valign: 'top', margin: TM });
  logo(s, 5.906, 0.461);
}

/* 25 - Enjoying The Wild ---------------------------------------------------- */
function slide25() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 13.333, 7.5, { label: false });
  poly(s, 0, 0.592, 13.333, 6.908, [
    { x: 0, y: 0 }, { x: 13.333, y: 6.908 }, { x: 0, y: 6.908 }, { close: true },
  ], { color: LIME, transparency: 40 });
  rect(s, 1.125, 1.167, 4.903, 5.167, TEAL, { fill: { color: TEAL, transparency: 20 } });
  text(s, 'Enjoying\nThe Wild', { x: 1.556, y: 1.88, w: 3.875, h: 1.616, fontSize: 48, bold: true, color: LIME, align: 'right', valign: 'middle', margin: TM });
  para(s, LOREM_FULL, 1.556, 3.542, 3.875, 2.121, { fontSize: 12, color: SOFT, align: 'right' });
  logo(s, 11.181, 0.461);
}

/* 26 - Great Adventures text card ------------------------------------------- */
function slide26() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 13.333, 7.5, { label: false });
  rect(s, 0, 0, 13.333, 7.5, TEAL, { fill: { color: TEAL, transparency: 50 } });
  rect(s, 0.571, 1.587, 12.131, 4.325, LIME, { fill: { color: LIME, transparency: 65 } });
  rect(s, 0, 1.756, 8.588, 3.987, WHITE, { fill: { color: WHITE, transparency: 25 } });
  text(s, 'Great Adventures', { x: 1.145, y: 2.259, w: 6.904, h: 0.74, fontSize: 44, bold: true, color: TEAL, valign: 'middle', margin: TM });
  para(s, [
    { text: FUSCE_LONG, options: { breakLine: true } },
    { text: L2 + 'Vivamus in lorem in odio tristique ultricies id vitae odio. Vestibulum sed sem suscipit dui luctus consequat quis sed justo. ' },
  ], 1.191, 3.126, 6.858, 2.121, { fontSize: 12, color: DARK });
  logo(s, 11.181, 0.461);
}

/* 27 - review bars beside a phone ------------------------------------------- */
const REVIEWS = [
  [1.273, CYAN, 'Expeditiion Review', SOFT, TEAL, 'mountain'],
  [2.968, TEAL, 'Guide Review', LIME, LIME, 'tree'],
  [4.664, LIME, 'Tour Review', TEAL, TEAL, 'pins'],
];

function slide27() {
  const s = pptx.addSlide();
  REVIEWS.forEach(([y, band, heading, ink, tileBg, glyph]) => {
    rect(s, 4.042, y, 8.416, 1.658, band);
    text(s, heading, { x: 5.424, y: y + 0.273, w: 6.746, h: 0.374, fontSize: 18, bold: true, color: ink });
    text(s, LOREM_MED, { x: 5.424, y: y + 0.618, w: 6.746, h: 0.767, fontSize: 11, fontFace: LIGHT, color: ink, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
    tile(s, 4.578, y + 0.504, 0.65, tileBg, glyph);
  });
  phoneMockup(s, 1.086, 0.598, 3.163, 6.304);
  logo(s, 10.936, 0.461);
}

/* 28 - laptop + teal footer band -------------------------------------------- */
function slide28() {
  const s = pptx.addSlide();
  for (let i = 0; i < 16; i++) rect(s, (13.333 * i) / 16, 4.828, 13.333 / 16 + 0.01, 2.672, mix(TEAL, CYAN, (i + 0.5) / 16));
  rect(s, 0, 4.828, 13.333, 0.234, LIME);
  laptopMockup(s, 0.835, 1.312, 6.313, 3.644);
  title(s, 'Great Adventures Begin Here', 7.299, 1.564, 4.701, 1.212);
  para(s, LOREM_MED, 7.299, 2.94, 4.701, 1.388);
  [['Embrace the change', 1.368], ['Maintain performance', 7.299]].forEach(([heading, x]) => {
    text(s, heading, { x, y: 5.505, w: x < 4 ? 3.065 : 3.474, h: 0.343, fontSize: 16, bold: true, color: WHITE });
    text(s, FUSCE_PANEL, { x, y: 5.85, w: 5.222, h: 0.99, fontSize: 11, fontFace: LIGHT, color: WHITE, lineSpacingMultiple: 1 });
  });
  logo(s, 5.906, 0.461);
}

/* 29 / 30 - chart left, lime cards right ------------------------------------ */
function chartCardsSlide(heading, drawChart) {
  const s = pptx.addSlide();
  drawChart(s);
  text(s, F1, { x: 1.125, y: 5.702, w: 5.875, h: 0.545, fontSize: 11, fontFace: LIGHT, color: TEAL, lineSpacingMultiple: 1.2, paraSpaceAfter: 6 });
  markingCards(s, 7.67);
  text(s, heading, { x: 7.67, y: 1.486, w: 5.013, h: 0.539, fontSize: 32, bold: true, color: DARK, valign: 'middle', margin: TM });
  logo(s, 0.574, 0.461);
}

const slide29 = () => chartCardsSlide('Time Management', (s) => mountainChart(s, 0.858, 2.024));
const slide30 = () => chartCardsSlide('Team Management', (s) => barChart(s, 1.125, 2.016));

/* ------------------------------------------------------------------- build  */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31].forEach((build) => build());

pptx.writeFile({ fileName: path.join(__dirname, '05431a0f-22fd-4d16-a81d-07ce5b09c325_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f))
  .catch((e) => { console.error(e); process.exit(1); });
