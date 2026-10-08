/**
 * "Winter Solstice" deck -- rebuilt with pptxgenjs.
 *
 * Recreated from a 20-slide 13.333x7.5in reference deck.
 * Raster photos in the original are empty picture placeholders or device
 * mock-ups; they are redrawn here with native shapes / flat placeholders.
 *
 *   node 00b2cbc8-bbff-44f2-9fef-648d5416581b_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'pptxgenjs';
pptx.title = 'Winter Solstice';
pptx.theme = { headFontFace: 'Raleway', bodyFontFace: 'Roboto' };

const S = pptx.shapes;
const HAIRLINE = 1.25;   // the pale flakes on the gradient slides are drawn much finer
const HEAD = 'Raleway';       // theme major font (+mj-lt)
const BODY = 'Roboto';        // theme minor font (+mn-lt)

/* ------------------------------------------------------------------ *
 * Theme palette (from the reference theme "Kustom 312")
 * ------------------------------------------------------------------ */
const C = {
  accent1: '0A80D1',
  accent2: '0460B7',
  accent3: '89CBE2',
  accent4: '0F9ED5',
  accent5: 'D7E2F1',
  accent6: '416F9A',
  dk2: '0E2841',
  lt2: 'E8E8E8',
  white: 'FFFFFF',
  black: '000000',
  // pre-computed luminance-modified variants used throughout the deck
  navy: '034889',      // accent2 lumMod 75%
  navyDeep: '02305C',  // accent2 lumMod 50%
  iceBlue: 'CAEEFB',   // accent4 lumMod 20% / lumOff 80%
  paleBlue: 'DCEAF7',  // dk2 lumMod 10% / lumOff 90%
  grey40: '404040',    // tx1 lumMod 75% / lumOff 25%
  grey59: '595959',    // tx1 lumMod 65% / lumOff 35%
  greyD9: 'D9D9D9',    // bg1 lumMod 85%
  greyF2: 'F2F2F2',
  frost: '9BC5E8',
};

/* Blended stand-ins for the reference's linear gradients ------------- */
const G = {
  // snowflake gradients: [start, end] shaded top -> bottom
  card: mix(C.accent2, C.accent4, 0.5),   // accent2 -> accent4
  flakeIce: [C.accent2, C.iceBlue],
  flakeAqua: [C.accent3, C.accent4],
  flakePale: ['EAF4FB', C.iceBlue],
  pill: '0B4C8E',
};

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
function mix(a, b, t) {
  const p = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  const v = i => Math.round(p(a, i) * (1 - t) + p(b, i) * t);
  return [0, 1, 2].map(i => v(i).toString(16).toUpperCase().padStart(2, '0')).join('');
}

function rect(sl, x, y, w, h, color, opts) {
  sl.addShape(S.RECTANGLE, Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, opts));
}

function roundRect(sl, x, y, w, h, radius, color, opts) {
  sl.addShape(S.ROUNDED_RECTANGLE, Object.assign(
    { x, y, w, h, rectRadius: radius, fill: { color }, line: { type: 'none' } }, opts));
}

function oval(sl, x, y, w, h, color, opts) {
  sl.addShape(S.OVAL, Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, opts));
}

function ring(sl, cx, cy, d, color, lw) {
  sl.addShape(S.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { type: 'none' }, line: { color, width: lw } });
}

/** Thick rounded bar joining two points (used for connector "tentacles"). */
function capsule(sl, x1, y1, x2, y2, thick, color) {
  const len = Math.hypot(x2 - x1, y2 - y1);
  const ang = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
  roundRect(sl, (x1 + x2) / 2 - len / 2, (y1 + y2) / 2 - thick / 2, len, thick, thick / 2, color, { rotate: ang });
}

/** Straight segment between two arbitrary points. */
function seg(sl, x1, y1, x2, y2, color, w) {
  sl.addShape(S.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    line: { color, width: w }, flipH: x2 < x1, flipV: y2 < y1,
  });
}

/** Hollow rounded frame -- device bezels keep the slide showing through. */
function frame(sl, x, y, w, h, radius, thick, color) {
  const t = thick / 2, r = Math.max(0, radius - t);
  sl.addShape(r > 0.01 ? S.ROUNDED_RECTANGLE : S.RECTANGLE, {
    x: x + t, y: y + t, w: w - thick, h: h - thick, rectRadius: r,
    fill: { type: 'none' }, line: { color: color || '111111', width: thick * 72 },
  });
}

/** Text box matching the reference defaults (top anchored, no autofit). */
function txt(sl, body, o) {
  sl.addText(body, Object.assign({ valign: 'top', fontFace: BODY, color: C.black }, o));
}

/* Linear gradient faked with a strip of banded rectangles ------------ */
function gradBandsH(sl, x, y, w, h, c0, c1, bands) {
  const n = bands || 24, bw = w / n;
  for (let i = 0; i < n; i++) {
    rect(sl, x + i * bw, y, bw + 0.012, h, mix(c0, c1, (i + 0.5) / n));
  }
}

/** Horizontal gradient inside a rounded rectangle (caps keep end colors). */
function gradRoundRectH(sl, x, y, w, h, r, c0, c1) {
  roundRect(sl, x, y, w, h, r, c0);
  gradBandsH(sl, x + r, y, w - 2 * r, h, mix(c0, c1, r / w), mix(c0, c1, 1 - r / w), 20);
  if (r > 0) roundRect(sl, x + w - 2 * r, y, 2 * r, h, r, c1);
}

/** Full-bleed background gradient, drawn as rotated bands. */
function gradientBackground(sl, angleDeg, stops) {
  const L = 15.4, N = 60, cx = 13.333 / 2, cy = 7.5 / 2;
  const a = angleDeg * Math.PI / 180, bw = L / N;
  const at = t => {
    for (let i = 1; i < stops.length; i++) {
      if (t <= stops[i][0] || i === stops.length - 1) {
        const [p0, k0] = stops[i - 1], [p1, k1] = stops[i];
        return mix(k0, k1, Math.min(1, Math.max(0, (t - p0) / (p1 - p0))));
      }
    }
    return stops[0][1];
  };
  for (let i = 0; i < N; i++) {
    const t = (i + 0.5) / N, d = (t - 0.5) * L;
    sl.addShape(S.RECTANGLE, {
      x: cx + d * Math.cos(a) - bw / 2, y: cy + d * Math.sin(a) - L / 2,
      w: bw + 0.02, h: L, rotate: angleDeg,
      fill: { color: at(t) }, line: { type: 'none' },
    });
  }
}

/* ------------------------------------------------------------------ *
 * Tiny line-art glyph set. Each icon is drawn inside a `size` box
 * centred on (cx, cy) using strokes / rings only.
 * ------------------------------------------------------------------ */
function icon(sl, kind, cx, cy, size, color) {
  const r = size / 2, lw = Math.max(0.75, size * 4);
  const box = (dx, dy, w, h) => sl.addShape(S.RECTANGLE, {
    x: cx + dx * r, y: cy + dy * r, w: w * r, h: h * r,
    fill: { type: 'none' }, line: { color, width: lw },
  });
  const line = (x1, y1, x2, y2) => seg(sl, cx + x1 * r, cy + y1 * r, cx + x2 * r, cy + y2 * r, color, lw);
  const dot = (dx, dy, d) => oval(sl, cx + dx * r - d * r / 2, cy + dy * r - d * r / 2, d * r, d * r, color);

  switch (kind) {
    case 'clock':
      ring(sl, cx, cy, size, color, lw);
      line(0, 0, 0, -0.55); line(0, 0, 0.4, 0.18);
      break;
    case 'phone':
      line(-0.45, -0.45, 0.1, 0.35); line(0.1, 0.35, 0.45, 0.45);
      line(-0.45, -0.45, -0.15, -0.55);
      break;
    case 'mail':
      box(-0.6, -0.42, 1.2, 0.84);
      line(-0.6, -0.42, 0, 0.1); line(0.6, -0.42, 0, 0.1);
      break;
    case 'globe':
      ring(sl, cx, cy, size, color, lw);
      line(-0.7, 0, 0.7, 0); ring(sl, cx, cy, size * 0.45, color, lw);
      break;
    case 'home':
      line(-0.65, 0, 0, -0.6); line(0.65, 0, 0, -0.6);
      box(-0.45, 0, 0.9, 0.6);
      break;
    case 'gear':
      ring(sl, cx, cy, size * 0.82, color, lw);
      ring(sl, cx, cy, size * 0.34, color, lw);
      for (let i = 0; i < 8; i++) {
        const a = i * Math.PI / 4;
        line(0.42 * Math.cos(a), 0.42 * Math.sin(a), 0.72 * Math.cos(a), 0.72 * Math.sin(a));
      }
      break;
    case 'shuffle':
      line(-0.7, -0.35, 0.7, 0.35); line(-0.7, 0.35, 0.7, -0.35);
      line(0.7, -0.35, 0.3, -0.45); line(0.7, 0.35, 0.3, 0.45);
      break;
    case 'bag':
      box(-0.55, -0.25, 1.1, 0.85); line(-0.25, -0.25, -0.15, -0.6); line(0.25, -0.25, 0.15, -0.6);
      break;
    case 'brush':
      line(-0.5, 0.5, 0.15, -0.15); box(0.1, -0.6, 0.5, 0.5);
      break;
    case 'shield':
      line(-0.5, -0.5, 0, -0.65); line(0.5, -0.5, 0, -0.65);
      line(-0.5, -0.5, 0, 0.65); line(0.5, -0.5, 0, 0.65);
      break;
    case 'power':
      ring(sl, cx, cy + 0.1 * r, size * 0.85, color, lw); line(0, -0.65, 0, 0);
      break;
    case 'print':
      box(-0.45, -0.6, 0.9, 0.35); box(-0.65, -0.25, 1.3, 0.5); box(-0.45, 0.25, 0.9, 0.4);
      break;
    case 'code':
      line(-0.25, -0.45, -0.65, 0); line(-0.65, 0, -0.25, 0.45);
      line(0.25, -0.45, 0.65, 0); line(0.65, 0, 0.25, 0.45);
      break;
    case 'ban':
      ring(sl, cx, cy, size, color, lw); line(-0.5, 0.5, 0.5, -0.5);
      break;
    case 'radiation':
      dot(0, 0, 0.28);
      for (let i = 0; i < 3; i++) {
        const a = (i * 120 - 90) * Math.PI / 180;
        dot(0.6 * Math.cos(a), 0.6 * Math.sin(a), 0.5);
      }
      break;
    case 'candle':
      box(-0.28, -0.1, 0.56, 0.8); line(0, -0.1, 0, -0.55);
      line(0.3, 0.3, 0.75, 0.55); line(0.3, 0.55, 0.75, 0.3);
      break;
    case 'snowman':
      ring(sl, cx, cy - 0.42 * r, size * 0.45, color, lw);
      ring(sl, cx, cy + 0.05 * r, size * 0.42, color, lw);
      ring(sl, cx, cy + 0.5 * r, size * 0.62, color, lw);
      line(-0.42, -0.4, -0.85, -0.4); line(0.42, -0.4, 0.85, -0.4);
      break;
    case 'route':
      dot(-0.45, -0.35, 0.42); dot(0.35, -0.45, 0.42);
      line(-0.6, 0.5, 0.1, 0.5); line(0.1, 0.5, 0.55, 0.15);
      break;
    case 'search':
      ring(sl, cx - 0.1 * r, cy - 0.1 * r, size * 0.7, color, lw);
      line(0.2, 0.2, 0.55, 0.55);
      break;
    default:
      ring(sl, cx, cy, size, color, lw);
  }
}

/* ------------------------------------------------------------------ *
 * Snowflake -- six-fold star used as the deck's signature ornament.
 * Box aspect is 0.866 (flat-to-flat / point-to-point of a hexagon).
 * ------------------------------------------------------------------ */
function snowflake(sl, x, y, w, h, color, lw, rotDeg) {
  const cx = x + w / 2, cy = y + h / 2;
  const phase = (rotDeg || 0) * Math.PI / 180;
  // shrink a rotated crystal so its arms still land inside the given box
  const R = (() => {
    let mc = 0, ms = 0;
    for (let i = 0; i < 6; i++) {
      const a = (90 + i * 60) * Math.PI / 180 - phase;
      mc = Math.max(mc, Math.abs(Math.cos(a))); ms = Math.max(ms, Math.abs(Math.sin(a)));
    }
    return Math.min(w / (2 * mc), h / (2 * ms));
  })();
  const width = lw || R * 3;   // stroke scales with the crystal
  // `color` may be a pair -- the reference fills these with a linear gradient,
  // approximated here by shading each stroke by its vertical position.
  const shade = my => Array.isArray(color)
    ? mix(color[0], color[1], Math.min(1, Math.max(0, (my - y) / h)))
    : color;
  const stroke = (x1, y1, x2, y2) => seg(sl, x1, y1, x2, y2, shade((y1 + y2) / 2), width);
  for (let i = 0; i < 6; i++) {
    const a = (90 + i * 60) * Math.PI / 180 - phase;
    const ux = Math.cos(a), uy = -Math.sin(a);
    stroke(cx, cy, cx + R * ux, cy + R * uy);
    [[0.33, 0.24], [0.54, 0.21], [0.73, 0.16], [0.90, 0.10]].forEach(([f, bl]) => {
      const bx = cx + R * f * ux, by = cy + R * f * uy;
      [-1, 1].forEach(sgn => {
        const b = a + sgn * 0.96;   // barbs splay ~55 deg off the arm
        stroke(bx, by, bx + R * bl * Math.cos(b), by - R * bl * Math.sin(b));
        // half-way tick: the secondary branching that makes the crystal lacy
        const mx = bx + R * bl * 0.5 * Math.cos(b), my = by - R * bl * 0.5 * Math.sin(b);
        const c2 = b + sgn * 0.96;
        stroke(mx, my, mx + R * bl * 0.36 * Math.cos(c2), my - R * bl * 0.36 * Math.sin(c2));
      });
    });
  }
}

/* ------------------------------------------------------------------ *
 * Repeated chrome: nav bar and the "More Detail >>" affordance
 * ------------------------------------------------------------------ */
function navBar(sl, opts) {
  const o = opts || {};
  if (o.logo !== false) {
    snowflake(sl, 0.226, 0.178, 0.313, 0.361, o.mark || C.accent4, 0.9);
    txt(sl, 'Winter', { x: 0.626, y: 0.195, w: 1.106, h: 0.303, fontSize: 12, bold: true, charSpacing: 1, fontFace: HEAD, color: o.word || C.accent2 });
  }
  if (o.search === false) return;
  txt(sl, 'Search', { x: 12.028, y: 0.22, w: 0.759, h: 0.278, fontSize: 10, align: 'center', color: o.search || C.accent4 });
  oval(sl, 12.848, 0.233, 0.252, 0.252, o.dot || C.navy);
  icon(sl, 'search', 12.974, 0.359, 0.13, C.white);
}

/** Two ">" chevrons drawn from four short strokes. */
function chevrons(sl, x, y, w, h, color, lw) {
  const half = w * 0.51;
  [0, w - half].forEach(dx => {
    seg(sl, x + dx, y, x + dx + half, y + h / 2, color, lw);
    seg(sl, x + dx, y + h, x + dx + half, y + h / 2, color, lw);
  });
}

function moreDetail(sl, x, y, cx, cy, color, size, w) {
  txt(sl, 'More Detail', { x, y, w: w || 1.503, h: 0.37, fontSize: size || 16, charSpacing: 1, color, valign: 'middle' });
  chevrons(sl, cx, cy, 0.16, 0.164, color, 1);
}

/* ================================================================== *
 * Slide 1 -- title
 * ================================================================== */
function slide01() {
  const sl = pptx.addSlide();
  gradientBackground(sl, 45, [[0, '54687F'], [0.5, '46829F'], [1, '3E9AD8']]);
  roundRect(sl, 7.812, 0.887, 4.281, 5.895, 2.1405, C.navy);
  snowflake(sl, 7.0, -0.24, 3.131, 3.597, 'EFF6FC', HAIRLINE);
  snowflake(sl, -0.709, 5.813, 2.671, 3.081, 'E6F1FA', HAIRLINE);
  snowflake(sl, 12.742, 5.552, 1.066, 1.229, 'E6F1FA', HAIRLINE);
  roundRect(sl, 6.176, 6.474, 0.981, 2.052, 0.4905, G.pill);

  txt(sl, 'Winter', { x: 1.064, y: 1.719, w: 6.665, h: 1.717, fontSize: 96, bold: true, fontFace: HEAD, color: C.accent3 });
  txt(sl, 'Solstice', { x: 1.064, y: 3.001, w: 6.665, h: 1.717, fontSize: 96, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, "simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the,  ",
    { x: 1.064, y: 4.65, w: 5.025, h: 0.625, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
  txt(sl, 'Celebration Of Light And Renewal',
    { x: 1.064, y: 5.478, w: 2.951, h: 0.303, fontSize: 12, bold: true, fontFace: HEAD, color: C.white, valign: 'middle' });

  navBar(sl, { mark: C.white, word: C.white, dot: C.accent2, search: C.white });
}

/* ================================================================== *
 * Slide 2 -- "A Celebration Of Light And Renewal"
 * ================================================================== */
function slide02() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 5.827, 0.799, 1.336, 1.54, G.flakeAqua);
  snowflake(sl, 11.27, 5.199, 1.367, 1.578, G.flakeIce);

  txt(sl, [
    { text: 'A Celebration ', options: { color: C.accent4 } },
    { text: 'Of Light And Renewal', options: { color: C.black } },
  ], { x: 7.805, y: 1.571, w: 4.84, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD });

  txt(sl, 'Embracing The Longest',
    { x: 1.189, y: 4.797, w: 3.865, h: 0.37, fontSize: 16, bold: true, fontFace: HEAD, color: C.accent4, align: 'right' });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
    { x: 0.821, y: 5.248, w: 4.232, h: 1.132, fontSize: 10.5, color: C.grey40, align: 'right', lineSpacingMultiple: 1.5 });

  gradRoundRectH(sl, 6.229, 4.376, 2.657, 2.423, 0.392, C.accent2, C.accent4);
  txt(sl, '231,4+', { x: 6.495, y: 4.815, w: 2.177, h: 0.707, fontSize: 36, bold: true, color: C.white });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor',
    { x: 6.568, y: 5.494, w: 2.028, h: 0.867, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
}

/* ================================================================== *
 * Slide 3 -- "Traditions As The Sun Begins Its Return"
 * ================================================================== */
function slide03() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 4.795, 0.43, 0.919, 1.059, G.flakeAqua);
  snowflake(sl, 2.713, 5.984, 2.44, 2.815, G.flakeIce);

  txt(sl, [
    { text: 'Traditions As ', options: { color: C.accent4 } },
    { text: 'The Sun Begins Its Return', options: { color: C.black } },
  ], { x: 1.034, y: 1.836, w: 4.232, h: 1.919, fontSize: 36, bold: true, fontFace: HEAD });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
    { x: 1.034, y: 3.916, w: 4.232, h: 1.132, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
  moreDetail(sl, 1.034, 5.294, 2.647, 5.393, C.accent4);

  gradRoundRectH(sl, 8.097, 1.194, 4.009, 1.879, 0.304, C.accent2, C.accent4);
  txt(sl, 'The Longest Night',
    { x: 8.404, y: 1.718, w: 2.587, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, 'adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, ',
    { x: 8.414, y: 1.977, w: 3.541, h: 0.602, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });

  txt(sl, '12,3K', { x: 6.321, y: 5.437, w: 2.177, h: 0.707, fontSize: 36, bold: true, color: C.accent1 });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
    { x: 6.394, y: 6.116, w: 2.44, h: 0.602, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
}

/* ================================================================== *
 * Slide 4 -- "The Depth Of Winter's Darkness"
 * ================================================================== */
function slide04() {
  const sl = pptx.addSlide();
  navBar(sl, { logo: false });
  snowflake(sl, 4.293, 2.572, 2.208, 2.545, G.flakeAqua);
  snowflake(sl, 11.624, 6.541, 1.566, 1.807, G.flakeIce);
  snowflake(sl, -0.234, 6.004, 0.754, 0.867, G.flakePale);

  txt(sl, [
    { text: 'The Depth Of ', options: { color: C.accent4 } },
    { text: 'Winter\u2019s Darkness', options: { color: C.black } },
  ], { x: 7.251, y: 1.203, w: 4.806, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD });
  moreDetail(sl, 7.257, 2.572, 8.87, 2.671, C.accent4);

  [['62,7%', 7.278, 7.524, 1.717], ['15,3%', 9.85, 10.167, 1.574]].forEach(([v, bx, tx, tw]) => {
    gradRoundRectH(sl, bx, 3.575, 2.208, 0.9, 0.1456, C.accent2, C.accent4);
    txt(sl, v, { x: tx, y: 3.739, w: tw, h: 0.572, fontSize: 28, bold: true, color: C.white, align: 'center' });
  });

  [1.301, 7.167].forEach(x => {
    txt(sl, 'Heritage Sites Worth ',
      { x, y: 5.093, w: 3.865, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4 });
    txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
      { x, y: 5.43, w: 5.366, h: 0.867, fontSize: 10.5, color: C.black, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================== *
 * Slide 5 -- "The Longest Night With"
 * ================================================================== */
function slide05() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 11.745, 1.013, 0.788, 0.908, G.flakeAqua);
  snowflake(sl, -0.892, 2.959, 1.877, 2.165, G.flakeIce);

  gradRoundRectH(sl, 10.13, 2.826, 1.906, 1.906, 0.21, C.accent2, C.accent4);
  icon(sl, 'clock', 11.083, 3.779, 0.62, C.white);

  seg(sl, 5.463, 5.125, 5.463, 7.5, C.greyD9, 1);

  txt(sl, [
    { text: 'The Longest ', options: { color: C.black } },
    { text: 'Night With', options: { color: C.accent4 } },
  ], { x: 1.355, y: 5.169, w: 4.168, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD });
  txt(sl, 'From Cities To Countrysides',
    { x: 6.266, y: 5.169, w: 3.865, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4 });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. ',
    { x: 6.266, y: 5.62, w: 5.823, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
}

/* ================================================================== *
 * Slide 6 -- "Longest Night With Hope Light"
 * ================================================================== */
function slide06() {
  const sl = pptx.addSlide();
  roundRect(sl, 0.999, -3.008, 2.823, 4.705, 1.4115, G.card);
  navBar(sl, { logo: false });
  snowflake(sl, 3.428, 3.093, 0.788, 0.908, G.flakeAqua);
  snowflake(sl, 10.579, 5.515, 2.362, 2.725, G.flakeIce);

  txt(sl, [
    { text: 'Longest Night ', options: { breakLine: true } },
    { text: 'With ' },
    { text: 'Hope Light', options: { color: C.accent4 } },
  ], { x: 7.741, y: 2.008, w: 4.54, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'Honoring Ancient Solstice',
    { x: 7.746, y: 3.636, w: 3.865, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4 });
  txt(sl, 'Maecenas congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Lorem ipsum dolor sit amet, consectetuer',
    { x: 7.739, y: 3.973, w: 4.54, h: 0.867, fontSize: 10.5, color: C.black, lineSpacingMultiple: 1.5 });
  moreDetail(sl, 7.746, 5.122, 9.359, 5.22, C.accent4);
}

/* ================================================================== *
 * Slide 7 -- "Welcoming The Return Of The Sun"
 * ================================================================== */
function slide07() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 4.304, -1.277, 2.362, 2.725, G.flakeIce);
  snowflake(sl, 8.047, 5.529, 1.524, 1.756, G.flakeAqua);

  txt(sl, [
    { text: 'Welcoming The Return Of ' },
    { text: 'The Sun', options: { color: C.accent4 } },
  ], { x: 1.04, y: 1.233, w: 4.626, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros',
    { x: 1.025, y: 2.625, w: 5.025, h: 0.602, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
  moreDetail(sl, 1.04, 3.417, 2.653, 3.515, C.accent4);

  gradRoundRectH(sl, 1.062, 4.202, 3.375, 2.256, 0.2484, C.accent2, C.accent4);
  const body = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor';
  txt(sl, 'Option hare ', { x: 1.736, y: 4.719, w: 1.693, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, body, { x: 1.737, y: 5.073, w: 2.326, h: 0.867, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
  txt(sl, 'Option hare ', { x: 5.366, y: 4.719, w: 1.693, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent1 });
  txt(sl, body, { x: 5.368, y: 5.073, w: 2.102, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
}

/* ================================================================== *
 * Slide 8 -- "Moment To Reflect, Restore, And Renew"
 * ================================================================== */
function slide08() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 5.494, 2.341, 1.524, 1.756, G.flakeAqua);
  snowflake(sl, 11.494, 6.351, 1.307, 1.507, G.flakeIce);

  txt(sl, [
    { text: 'Moment To Reflect, Restore, ' },
    { text: 'And Renew', options: { color: C.accent4 } },
  ], { x: 7.462, y: 1.043, w: 5.49, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'From Cities To Countrysides',
    { x: 7.462, y: 2.476, w: 3.865, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4 });
  txt(sl, 'congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. elit. ',
    { x: 7.462, y: 2.83, w: 5.071, h: 0.602, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });

  [7.873, 10.352].forEach(x => seg(sl, x, 4.382, x, 6.327, C.greyD9, 1));
  [5.638, 8.102, 10.581].forEach((x, i) => {
    txt(sl, 'Option hare ', { x, y: 4.719, w: 1.693, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent1 });
    txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor',
      { x: [5.64, 8.103, 10.583][i], y: 5.073, w: 2.326, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================== *
 * Slide 9 -- "Meet Our Best Team"
 * ================================================================== */
function slide09() {
  const sl = pptx.addSlide();
  navBar(sl);
  txt(sl, [
    { text: 'Meet Our ' },
    { text: 'Best Team', options: { color: C.accent4 } },
  ], { x: 3.332, y: 0.394, w: 6.67, h: 0.707, fontSize: 36, bold: true, fontFace: HEAD, color: C.black, align: 'center' });

  const team = [
    [1.099, 1.349, 'Archibald Winston'],
    [4.944, 5.29, 'Charlotte Sophia'],
    [8.789, 9.232, 'Gabriel Thomas'],
  ];
  team.forEach(([cardX, textX, name]) => {
    roundRect(sl, cardX, 1.55, 3.445, 4.401, 1.7225, C.navy);
    txt(sl, name, { x: textX, y: 6.227, w: 2.752, h: 0.404, fontSize: 18, bold: true, fontFace: HEAD, color: C.accent6, align: 'center' });
    txt(sl, 'Lorem ipsum dolor', { x: textX, y: 6.699, w: 2.752, h: 0.303, fontSize: 12, color: C.grey40, align: 'center' });
  });

  snowflake(sl, 0.311, 3.467, 2.077, 2.394, G.flakeAqua, null, 18.4);
  snowflake(sl, 10.858, 1.316, 2.075, 2.394, [C.accent3, 'BFE0EE'], null, -9.1);
}

/* ================================================================== *
 * Slide 10 -- "Break Slide"
 * ================================================================== */
function slide10() {
  const sl = pptx.addSlide();
  gradientBackground(sl, 0, [[0, '3982C7'], [1, '386EA2']]);
  snowflake(sl, -0.854, 5.507, 2.671, 3.081, 'E6F1FA', HAIRLINE);
  snowflake(sl, 9.753, 1.115, 2.109, 2.423, 'EFF6FC', HAIRLINE);
  snowflake(sl, 11.329, 5.461, 1.066, 1.229, 'E6F1FA', HAIRLINE);
  roundRect(sl, 7.557, -0.936, 0.981, 2.052, 0.4905, G.pill);

  txt(sl, 'Break', { x: 1.441, y: 1.318, w: 7.458, h: 2.895, fontSize: 166, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, 'Slide', { x: 3.634, y: 3.758, w: 5.238, h: 2.423, fontSize: 138, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, 'Celebration Of Light And Renewal',
    { x: 1.742, y: 4.617, w: 1.255, h: 0.707, fontSize: 12, bold: true, fontFace: HEAD, color: C.white, align: 'right', valign: 'middle' });

  roundRect(sl, 9.235, 3.758, 2.657, 2.423, 0.392, C.navyDeep);
  txt(sl, '12.15+', { x: 9.501, y: 4.198, w: 2.177, h: 0.707, fontSize: 36, bold: true, color: C.white });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor',
    { x: 9.574, y: 4.876, w: 2.028, h: 0.867, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });

  navBar(sl, { mark: C.white, word: C.white, dot: C.accent2, search: C.white });
}

/* ================================================================== *
 * Slide 11 -- progress bars
 * ================================================================== */
function slide11() {
  const sl = pptx.addSlide();
  navBar(sl);
  roundRect(sl, 9.71, 5.806, 2.823, 4.705, 1.4115, G.card);
  snowflake(sl, 5.486, -1.35, 2.362, 2.725, G.flakeIce);
  snowflake(sl, 12.822, 4.889, 1.024, 1.18, G.flakeAqua);
  snowflake(sl, 0.308, 6.466, 0.754, 0.867, G.flakePale);

  txt(sl, [
    { text: 'Welcoming The Return Of ' },
    { text: 'The Sun', options: { color: C.accent4 } },
  ], { x: 0.691, y: 1.571, w: 4.626, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros',
    { x: 0.686, y: 3.059, w: 5.025, h: 0.602, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });

  // track: 4.225in wide starting at 1.062 ; fill starts at 0.780
  [[3.949, 4.392, 3.652, '80%'], [4.677, 5.12, 3.145, '70%']].forEach(([ty, by, fillW, pct]) => {
    roundRect(sl, 1.062, by, 4.225, 0.084, 0.042, C.greyF2);
    roundRect(sl, 0.78, by, fillW, 0.084, 0.042, C.accent2);
    txt(sl, 'Renewing Strength', { x: 0.691, y: ty, w: 1.989, h: 0.454, fontSize: 14, fontFace: HEAD, color: C.black, lineSpacingMultiple: 1.5 });
    txt(sl, pct, { x: 4.534, y: ty - 0.015, w: 0.987, h: 0.454, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4, align: 'center', lineSpacingMultiple: 1.5 });
  });

  txt(sl, 'Finding Warmth and Meaning Amid Winter\u2019s',
    { x: 0.707, y: 5.627, w: 4.378, h: 0.303, fontSize: 12, bold: true, fontFace: HEAD, color: C.accent4 });
}

/* ================================================================== *
 * Slide 12 -- circular infographic
 * ================================================================== */
function slide12() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 1.119, 0.835, 1.683, 1.94, G.flakeAqua);
  snowflake(sl, 11.729, 5.088, 1.119, 1.291, G.flakeIce);
  txt(sl, [
    { text: 'Infographic', options: { color: C.accent4 } },
    { text: ' Slide Design' },
  ], { x: 2.976, y: 0.496, w: 7.063, h: 0.653, fontSize: 36, bold: true, fontFace: HEAD, color: C.black, align: 'center' });

  // main ring, plus the two organic arms that reach out to the satellites
  const cx = 7.089, cy = 4.28;
  capsule(sl, cx, cy, 4.677, 3.912, 0.64, C.accent1);
  capsule(sl, cx, cy, 4.61, 5.584, 0.72, C.accent1);
  oval(sl, cx - 1.583, cy - 1.583, 3.167, 3.167, C.accent1);
  oval(sl, 5.834, 3.025, 2.51, 2.51, C.white);
  txt(sl, 'Write Something Here',
    { x: 6.027, y: 3.623, w: 2.123, h: 1.313, fontSize: 24, bold: true, fontFace: HEAD, color: C.dk2, align: 'center' });

  // upper-right lobe: two bubbles fused by a tapered neck, over a dark shadow
  [[C.navyDeep, 0.05], [C.accent3, 0]].forEach(([color, dy]) => {
    oval(sl, 7.475, 2.275 + dy, 1.33, 1.33, color);
    oval(sl, 8.541, 3.643 + dy, 0.82, 0.82, color);
    capsule(sl, 8.14, 2.94 + dy, 8.951, 4.053 + dy, 0.82, color);
  });
  oval(sl, 7.673, 2.484, 0.936, 0.937, C.lt2);
  icon(sl, 'shield', 8.141, 2.952, 0.36, C.accent3);
  oval(sl, 8.682, 3.784, 0.539, 0.54, C.lt2);
  icon(sl, 'bag', 8.951, 4.054, 0.22, C.accent3);

  // satellites
  oval(sl, 3.999, 4.972, 1.223, 1.224, C.lt2);
  icon(sl, 'gear', 4.610, 5.584, 0.38, C.accent1);
  oval(sl, 4.146, 3.382, 1.062, 1.059, C.accent2);
  icon(sl, 'brush', 4.677, 3.912, 0.38, C.white);
  oval(sl, 8.733, 4.827, 0.799, 0.795, C.accent4);
  icon(sl, 'shuffle', 9.132, 5.224, 0.36, C.white);

  const body = 'eligendi optio cumque nihil impedit  nihil impedit quo minus id quod';
  const labels = [
    [9.035, 1.86, 9.035, 2.221, 2.477, 'l'],
    [9.916, 3.387, 9.922, 3.802, 2.477, 'l'],
    [8.733, 5.687, 8.733, 6.104, 2.612, 'l'],
    [2.072, 3.373, 1.263, 3.79, 2.455, 'r'],
    [1.71, 5.519, 0.92, 5.901, 2.455, 'r'],
  ];
  labels.forEach(([hx, hy, bx, by, bw, al]) => {
    txt(sl, 'Option Here', { x: hx, y: hy, w: 1.666, h: 0.413, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent1, align: al, lineSpacingMultiple: 1.5 });
    txt(sl, body, { x: bx, y: by, w: bw, h: 0.625, fontSize: 10.5, color: C.grey40, align: al, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================== *
 * Slide 13 -- radial "pie" infographic
 * ================================================================== */
function slide13() {
  const sl = pptx.addSlide();
  navBar(sl);
  snowflake(sl, 12.028, 2.094, 1.683, 1.94, G.flakeAqua);
  snowflake(sl, 0.475, 5.814, 1.119, 1.291, G.flakeIce);
  txt(sl, [
    { text: 'Infographic', options: { color: C.accent4 } },
    { text: ' Slide Design' },
  ], { x: 2.976, y: 0.496, w: 7.063, h: 0.653, fontSize: 36, bold: true, fontFace: HEAD, color: C.black, align: 'center' });

  // sectors: [startAngle, endAngle] clockwise from East, radius, fill
  const cx = 6.677, cy = 4.495;
  const sectors = [
    [270, 360, 2.218, C.accent4],
    [0, 45, 1.777, C.accent5],
    [45, 90, 1.221, C.accent6],
    [90, 135, 1.516, C.paleBlue],
    [135, 180, 2.021, C.accent1],
    [180, 225, 1.512, C.accent2],
    [225, 270, 1.899, C.accent3],
  ];
  sectors.forEach(([a0, a1, r, color]) => {
    sl.addShape(S.PIE, {
      x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, angleRange: [a0, a1],
      fill: { color }, line: { color: C.white, width: 1 },
    });
  });
  // white glyphs sitting on each sector
  [['power', 6.323, 3.332], ['print', 5.579, 4.854], ['print', 5.901, 4.161],
   ['code', 6.925, 5.094], ['ban', 7.662, 4.842], ['radiation', 7.435, 3.431],
   ['code', 6.230, 5.292]].forEach(([kind, ix, iy]) => icon(sl, kind, ix, iy, 0.3, C.white));

  // elbow leaders running from each sector out to its caption
  [[[5.853, 3.041], [5.196, 3.041], [5.196, 2.375], [4.723, 2.375]],
   [[6.151, 5.493], [6.151, 6.138], [5.077, 6.138]],
   [[4.208, 5.055], [5.241, 5.055]],
   [[4.161, 3.479], [5.288, 3.479], [5.288, 3.884], [5.664, 3.884]],
   [[6.980, 5.308], [6.980, 6.076], [9.036, 6.076]],
   [[7.540, 3.704], [9.266, 3.704], [9.266, 3.190]],
   [[7.902, 5.160], [9.367, 5.160], [9.367, 4.596]]].forEach(pathPts => {
    for (let i = 1; i < pathPts.length; i++) {
      seg(sl, pathPts[i - 1][0], pathPts[i - 1][1], pathPts[i][0], pathPts[i][1], C.paleBlue, 0.5);
    }
  });

  const body = 'eligendi optio cumque nihil impedit  nihil impedit quo minus id quod';
  const labels = [
    [8.892, 2.046, 8.892, 2.401, 2.573, 'l'],
    [9.511, 4.063, 9.511, 4.418, 2.571, 'l'],
    [9.071, 5.721, 9.071, 6.076, 2.654, 'l'],
    [2.817, 1.956, 1.829, 2.311, 2.654, 'r'],
    [2.182, 3.289, 1.194, 3.644, 2.654, 'r'],
    [2.434, 4.741, 1.446, 5.096, 2.654, 'r'],
    [3.26, 5.884, 2.272, 6.238, 2.654, 'r'],
  ];
  labels.forEach(([hx, hy, bx, by, bw, al]) => {
    txt(sl, 'Option Here', { x: hx, y: hy, w: 1.666, h: 0.413, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent1, align: al, lineSpacingMultiple: 1.5 });
    txt(sl, body, { x: bx, y: by, w: bw, h: 0.602, fontSize: 10.5, color: C.grey59, align: al, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================== *
 * Slide 14 -- two phone mock-ups + KPI figures
 * ================================================================== */
/** Landscape phone shell: bezel, side notch, top buttons, speaker slot. */
function phoneMockup(sl, x, y, w, h) {
  frame(sl, x, y, w, h, 0.42, 0.18);
  roundRect(sl, x + 0.09, y + h * 0.24, 0.22, h * 0.36, 0.1, '111111');   // notch
  [0.36, 0.45, 0.54].forEach(f => rect(sl, x + w * f, y + 0.02, w * 0.06, 0.05, '111111'));
  rect(sl, x + w * 0.44, y + h - 0.07, w * 0.12, 0.05, '111111');         // speaker
}

function slide14() {
  const sl = pptx.addSlide();
  phoneMockup(sl, -1.339, 3.792, 6.28, 3.111);
  phoneMockup(sl, 7.187, 0.403, 6.146, 3.111);
  navBar(sl, { search: false });
  snowflake(sl, 6.265, 1.951, 2.059, 2.375, G.flakeIce);
  snowflake(sl, 1.814, 6.416, 0.781, 0.9, G.flakeAqua);
  snowflake(sl, 11.241, 6.574, 1.292, 1.484, G.flakePale);

  txt(sl, [
    { text: 'Celebrating', options: { color: C.accent4 } },
    { text: ' the Light That Emerges from the Shadows' },
  ], { x: 0.733, y: 1.165, w: 5.073, h: 1.919, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });

  [['+273,5', 5.775, 3.227], ['+162,3', 9.396, 3.512]].forEach(([v, x, bw]) => {
    txt(sl, v, { x, y: 4.449, w: 2.059, h: 0.64, fontSize: 32, bold: true, color: C.accent2 });
    txt(sl, 'who fail in their duty through weakness of will, which is the same as saying through shrinking from toil and pain. cases are perfectly',
      { x, y: 5.092, w: bw, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================== *
 * Slide 15 -- desktop monitor + sliders
 * ================================================================== */
function slide15() {
  const sl = pptx.addSlide();
  gradBandsH(sl, -0.008, 3.224, 4.043, 4.302, C.accent2, C.accent4, 20);
  // monitor: hollow bezel, trapezoid neck, base bar
  frame(sl, 0.93, 1.306, 5.979, 4.083, 0.06, 0.26);
  rect(sl, 0.93, 4.98, 5.979, 0.41, '111111');
  sl.addShape(S.CUSTOM_GEOMETRY, {
    x: 2.8, y: 5.389, w: 2.18, h: 0.74, fill: { color: '111111' }, line: { type: 'none' },
    points: [{ x: 0.349, y: 0 }, { x: 1.876, y: 0 }, { x: 2.18, y: 0.74 }, { x: 0, y: 0.74 }, { close: true }],
  });
  roundRect(sl, 2.8, 6.05, 2.18, 0.13, 0.065, '111111');
  navBar(sl);
  snowflake(sl, 6.046, 0.56, 1.501, 1.731, G.flakeIce);
  snowflake(sl, 10.13, 6.728, 0.781, 0.9, G.flakeAqua);

  roundRect(sl, 1.593, 0.707, 1.438, 1.438, 0.3065, C.accent2);
  icon(sl, 'candle', 2.312, 1.421, 0.68, C.white);

  txt(sl, [
    { text: 'Lighting Candles To ' },
    { text: 'Symbolize', options: { color: C.accent4 } },
  ], { x: 7.906, y: 1.652, w: 4.448, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which many variations of passages ',
    { x: 7.906, y: 3.189, w: 4.627, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });

  [[4.538, 4.968, 3.233, C.accent2, '76%', 4.549, 0.662],
   [5.398, 5.848, 3.795, C.accent4, '87%', 5.373, 0.759]].forEach(([ly, by, fw, col, pct, py, pw]) => {
    txt(sl, 'Insert Subtitle Here', { x: 8.039, y: ly, w: 1.893, h: 0.286, fontSize: 11, bold: true, fontFace: HEAD, color: C.black });
    roundRect(sl, 8.072, by - 0.052, 4.281, 0.104, 0.052, C.accent5);
    roundRect(sl, 8.039, by - 0.052, fw, 0.104, 0.052, col);
    oval(sl, 8.039 + fw - 0.115, by - 0.115, 0.23, 0.23, col);
    txt(sl, pct, { x: 11.803, y: py, w: pw, h: 0.337, fontSize: 14, color: C.black });
  });
}

/* ================================================================== *
 * Slide 16 -- tablet mock-up + wide gradient band
 * ================================================================== */
function slide16() {
  const sl = pptx.addSlide();
  gradRoundRectH(sl, -0.497, 3.842, 13.181, 2.027, 0.3804, C.accent2, C.accent4);
  navBar(sl);
  snowflake(sl, 11.737, 0.945, 0.781, 0.9, G.flakeAqua);
  snowflake(sl, 6.485, 6.402, 1.501, 1.731, G.flakeIce, null, 18);

  // tablet: hollow bezel, deeper chins top and bottom, home button
  frame(sl, 1.179, 0.806, 4.246, 5.902, 0.3, 0.222);
  rect(sl, 1.401, 1.028, 3.802, 0.194, '111111');
  rect(sl, 1.401, 6.292, 3.802, 0.194, '111111');
  oval(sl, 3.24, 6.44, 0.12, 0.12, '111111');
  ring(sl, 3.3, 6.5, 0.15, '111111', 1);

  roundRect(sl, 4.761, 1.751, 1.23, 1.209, 0.18, C.accent2);
  icon(sl, 'snowman', 5.376, 2.356, 0.62, C.white);

  txt(sl, [
    { text: 'Celebrating', options: { color: C.accent4 } },
    { text: ' The ', options: { breakLine: true } },
    { text: 'Light Emerges From' },
  ], { x: 6.656, y: 1.632, w: 5.094, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  moreDetail(sl, 6.667, 3.08, 8.279, 3.178, C.accent4);

  txt(sl, 'Welcome The New Cycle',
    { x: 6.656, y: 4.254, w: 5.708, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, 'Consectetur adipiscing elit. Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urnadui suscipit purus, ',
    { x: 6.656, y: 4.591, w: 5.472, h: 0.867, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
}

/* ================================================================== *
 * Slide 17 / 18 -- map slides
 * ================================================================== */
/** Draw a polygon given normalised (0..1) points inside a box. */
function landmass(sl, box, pts, color) {
  const [bx, by, bw, bh] = box;
  sl.addShape(S.CUSTOM_GEOMETRY, {
    x: bx, y: by, w: bw, h: bh,
    fill: { color: color || C.accent5 }, line: { color: C.white, width: 1 },
    points: pts.map(([u, v], i) => ({ x: bw * u, y: bh * v, moveTo: i === 0 })).concat([{ close: true }]),
  });
}

/** Map pin: teardrop + hole, matching the reference "Teardrop + Oval" pair. */
function pin(sl, x, y, size, color, holeColor) {
  sl.addShape(S.TEAR, { x, y, w: size, h: size, rotate: 135, fill: { color }, line: { type: 'none' } });
  oval(sl, x + size * 0.196, y + size * 0.196, size * 0.608, size * 0.608, holeColor);
}

/* Simplified coastlines, normalised 0..1 inside each map's bounding box. */
const WORLD = [
  [[0.996,0.208],[0.866,0.14],[0.829,0.159],[0.816,0.131],[0.76,0.135],[0.782,0.106],[0.758,0.084],
   [0.698,0.129],[0.699,0.149],[0.679,0.14],[0.676,0.166],[0.685,0.187],[0.662,0.201],[0.672,0.191],
   [0.67,0.138],[0.656,0.16],[0.661,0.182],[0.637,0.17],[0.596,0.197],[0.594,0.18],[0.574,0.222],[0.56,0.198],
   [0.578,0.204],[0.58,0.186],[0.531,0.16],[0.511,0.179],[0.484,0.263],[0.514,0.291],[0.517,0.246],
   [0.537,0.207],[0.528,0.247],[0.549,0.257],[0.524,0.302],[0.501,0.309],[0.498,0.278],[0.491,0.311],
   [0.456,0.35],[0.465,0.384],[0.448,0.384],[0.445,0.427],[0.461,0.43],[0.493,0.381],[0.513,0.42],[0.519,0.407],
   [0.505,0.371],[0.532,0.429],[0.53,0.407],[0.548,0.399],[0.553,0.365],[0.576,0.36],[0.571,0.376],
   [0.583,0.395],[0.541,0.407],[0.546,0.429],[0.568,0.437],[0.563,0.463],[0.521,0.469],[0.499,0.451],
   [0.497,0.427],[0.452,0.44],[0.425,0.514],[0.426,0.582],[0.455,0.621],[0.494,0.621],[0.513,0.812],
   [0.518,0.843],[0.532,0.848],[0.55,0.829],[0.58,0.729],[0.58,0.662],[0.607,0.599],[0.609,0.579],[0.59,0.588],
   [0.579,0.562],[0.564,0.483],[0.592,0.575],[0.632,0.527],[0.624,0.495],[0.612,0.509],[0.603,0.474],
   [0.652,0.502],[0.671,0.526],[0.685,0.602],[0.717,0.522],[0.735,0.553],[0.754,0.639],[0.744,0.572],
   [0.757,0.597],[0.769,0.583],[0.765,0.527],[0.803,0.484],[0.798,0.439],[0.806,0.428],[0.793,0.417],
   [0.813,0.412],[0.819,0.447],[0.822,0.41],[0.854,0.354],[0.859,0.322],[0.842,0.303],[0.862,0.27],[0.926,0.24],
   [0.903,0.281],[0.903,0.332],[0.924,0.261],[0.964,0.24],[0.958,0.218],[0.971,0.206],[0.987,0.224]],
  [[0.367,0.677],[0.333,0.657],[0.298,0.588],[0.24,0.593],[0.222,0.524],[0.206,0.541],[0.197,0.523],
   [0.208,0.476],[0.234,0.476],[0.245,0.5],[0.256,0.419],[0.301,0.367],[0.284,0.344],[0.311,0.327],
   [0.304,0.355],[0.321,0.36],[0.291,0.26],[0.28,0.271],[0.265,0.24],[0.251,0.246],[0.245,0.325],[0.205,0.265],
   [0.225,0.212],[0.244,0.226],[0.228,0.207],[0.247,0.166],[0.275,0.202],[0.251,0.219],[0.283,0.243],
   [0.28,0.203],[0.297,0.2],[0.281,0.167],[0.238,0.127],[0.23,0.15],[0.232,0.129],[0.218,0.152],[0.236,0.169],
   [0.227,0.189],[0.205,0.148],[0.215,0.13],[0.206,0.123],[0.196,0.181],[0.168,0.197],[0.15,0.175],
   [0.185,0.175],[0.174,0.128],[0.169,0.149],[0.131,0.121],[0.119,0.143],[0.127,0.156],[0.139,0.139],
   [0.139,0.158],[0.156,0.165],[0.147,0.175],[0.037,0.157],[0.005,0.185],[0.015,0.205],[0.003,0.218],
   [0.018,0.221],[0.005,0.248],[0.027,0.276],[0.021,0.295],[0.056,0.255],[0.098,0.29],[0.122,0.361],
   [0.122,0.412],[0.155,0.501],[0.188,0.552],[0.249,0.599],[0.258,0.653],[0.242,0.674],[0.273,0.77],
   [0.258,0.945],[0.269,0.98],[0.263,0.987],[0.286,0.994],[0.276,0.972],[0.287,0.896],[0.355,0.773]],
  [[0.402,0.166],[0.408,0.14],[0.401,0.133],[0.416,0.12],[0.407,0.093],[0.419,0.09],[0.409,0.082],[0.43,0.029],
   [0.403,0.043],[0.407,0.027],[0.378,0.032],[0.406,0.021],[0.398,0.009],[0.34,0.012],[0.303,0.028],
   [0.267,0.075],[0.311,0.119],[0.326,0.161],[0.316,0.171],[0.327,0.172],[0.32,0.206],[0.344,0.258],
   [0.358,0.212]],
  [[0.88,0.77],[0.862,0.71],[0.857,0.746],[0.845,0.718],[0.816,0.728],[0.781,0.784],[0.785,0.848],[0.837,0.838],
   [0.873,0.881],[0.893,0.815]],
  [[0.235,0.07],[0.219,0.095],[0.251,0.096],[0.24,0.089],[0.258,0.076],[0.253,0.063],[0.298,0.021],[0.218,0.029]],
];

const USA = [
  [[0.592,0.639],[0.579,0.556],[0.522,0.554],[0.469,0.526],[0.469,0.463],[0.414,0.461],[0.411,0.611],
   [0.343,0.607],[0.403,0.71],[0.424,0.682],[0.442,0.686],[0.487,0.799],[0.522,0.816],[0.518,0.749],
   [0.539,0.73],[0.534,0.72],[0.56,0.712],[0.564,0.685],[0.57,0.697],[0.587,0.685]],
  [[0.21,0.51],[0.122,0.332],[0.133,0.235],[0.064,0.214],[0.053,0.263],[0.057,0.315],[0.065,0.35],[0.092,0.357],
   [0.072,0.366],[0.092,0.474],[0.129,0.502],[0.149,0.557],[0.198,0.561]],
  [[0.404,0.151],[0.404,0.182],[0.289,0.17],[0.288,0.186],[0.259,0.185],[0.247,0.143],[0.234,0.142],
   [0.241,0.107],[0.221,0.079],[0.221,0.029],[0.406,0.054]],
  [[0.352,0.873],[0.308,0.82],[0.299,0.838],[0.277,0.821],[0.257,0.675],[0.2,0.654],[0.158,0.695],[0.182,0.732],
   [0.146,0.735],[0.154,0.756],[0.178,0.758],[0.148,0.799],[0.16,0.82],[0.149,0.821],[0.157,0.835],
   [0.167,0.822],[0.163,0.855],[0.193,0.851],[0.189,0.876],[0.151,0.914],[0.182,0.9],[0.232,0.811],[0.221,0.85],
   [0.238,0.836],[0.241,0.814],[0.26,0.831],[0.285,0.823],[0.305,0.847],[0.301,0.834],[0.314,0.843],
   [0.308,0.824],[0.338,0.876]],
  [[0.345,0.613],[0.297,0.624],[0.308,0.436],[0.415,0.445],[0.411,0.611]],
  [[0.306,0.465],[0.297,0.624],[0.259,0.619],[0.192,0.568],[0.21,0.511],[0.203,0.452],[0.214,0.453],
   [0.219,0.422],[0.308,0.436]],
  [[0.216,0.449],[0.204,0.449],[0.203,0.485],[0.122,0.332],[0.133,0.235],[0.233,0.258]],
  [[0.434,0.346],[0.433,0.445],[0.308,0.436],[0.315,0.304],[0.434,0.313]],
  [[0.402,0.246],[0.4,0.312],[0.281,0.299],[0.289,0.17],[0.403,0.181]],
  [[0.1,0.087],[0.111,0.111],[0.203,0.128],[0.188,0.169],[0.183,0.247],[0.062,0.212],[0.087,0.082]],
  [[0.283,0.267],[0.281,0.299],[0.315,0.304],[0.308,0.436],[0.219,0.422],[0.233,0.258]],
  [[0.632,0.075],[0.566,0.067],[0.547,0.042],[0.512,0.055],[0.528,0.23],[0.614,0.224],[0.586,0.186],
   [0.584,0.155],[0.595,0.118]],
  [[0.288,0.186],[0.283,0.267],[0.183,0.247],[0.188,0.169],[0.204,0.135],[0.205,0.025],[0.221,0.029],
   [0.221,0.079],[0.241,0.107],[0.234,0.142],[0.247,0.143],[0.259,0.185]],
  [[0.565,0.442],[0.433,0.445],[0.434,0.346],[0.556,0.347]],
  [[0.528,0.263],[0.55,0.344],[0.434,0.346],[0.434,0.313],[0.4,0.312],[0.402,0.246]],
  [[0.525,0.151],[0.528,0.263],[0.402,0.246],[0.404,0.151]],
  [[0.525,0.151],[0.404,0.151],[0.406,0.054],[0.512,0.055]],
  [[0.616,0.325],[0.663,0.445],[0.646,0.468],[0.646,0.451],[0.565,0.459],[0.563,0.371],[0.541,0.325]],
  [[0.565,0.459],[0.571,0.554],[0.484,0.539],[0.469,0.526],[0.469,0.463],[0.415,0.461],[0.415,0.445],
   [0.561,0.442]],
  [[0.204,0.025],[0.118,0.001],[0.123,0.034],[0.108,0.061],[0.115,0.029],[0.086,0.01],[0.086,0.08],
   [0.111,0.111],[0.198,0.12]],
  [[0.828,0.564],[0.776,0.477],[0.731,0.487],[0.757,0.625],[0.811,0.625]],
  [[0.687,0.366],[0.673,0.248],[0.626,0.254],[0.635,0.275],[0.615,0.332],[0.641,0.373],[0.638,0.394],
   [0.659,0.432],[0.674,0.428]],
  [[0.866,0.756],[0.822,0.611],[0.812,0.609],[0.811,0.625],[0.705,0.626],[0.708,0.649],[0.733,0.637],
   [0.752,0.661],[0.776,0.642],[0.802,0.668],[0.809,0.723],[0.85,0.792],[0.863,0.787]],
  [[0.675,0.171],[0.665,0.182],[0.665,0.144],[0.613,0.123],[0.614,0.111],[0.595,0.12],[0.584,0.155],
   [0.586,0.186],[0.612,0.209],[0.626,0.254],[0.673,0.248]],
  [[0.626,0.254],[0.635,0.275],[0.616,0.325],[0.541,0.325],[0.525,0.23],[0.614,0.224]],
  [[0.654,0.466],[0.633,0.569],[0.58,0.574],[0.565,0.459],[0.646,0.451],[0.643,0.468]],
  [[0.754,0.616],[0.731,0.487],[0.684,0.494],[0.691,0.649],[0.697,0.636],[0.699,0.653],[0.709,0.647],
   [0.705,0.626]],
  [[0.903,0.421],[0.885,0.409],[0.903,0.405],[0.896,0.391],[0.797,0.418],[0.754,0.481],[0.813,0.463],
   [0.872,0.489],[0.879,0.457],[0.899,0.444],[0.887,0.447],[0.894,0.435],[0.883,0.43]],
  [[0.691,0.649],[0.686,0.498],[0.646,0.501],[0.631,0.546],[0.64,0.592],[0.628,0.637],[0.665,0.632],
   [0.669,0.658]],
  [[0.91,0.235],[0.898,0.105],[0.869,0.12],[0.849,0.18],[0.815,0.19],[0.82,0.205],[0.807,0.225],[0.882,0.211]],
  [[0.681,0.69],[0.667,0.677],[0.674,0.663],[0.653,0.662],[0.67,0.658],[0.665,0.632],[0.628,0.637],[0.64,0.592],
   [0.634,0.568],[0.58,0.574],[0.592,0.639],[0.587,0.683]],
  [[0.9,0.266],[0.882,0.211],[0.796,0.238],[0.805,0.311],[0.885,0.287]],
  [[0.761,0.426],[0.797,0.427],[0.755,0.482],[0.646,0.501],[0.656,0.449]],
  [[0.761,0.214],[0.751,0.182],[0.733,0.195],[0.738,0.144],[0.71,0.133],[0.704,0.168],[0.703,0.153],
   [0.693,0.171],[0.691,0.269],[0.748,0.258]],
  [[0.799,0.283],[0.796,0.238],[0.766,0.265],[0.725,0.264],[0.733,0.349],[0.778,0.361],[0.799,0.316]],
  [[0.787,0.389],[0.768,0.353],[0.734,0.347],[0.715,0.39],[0.682,0.397],[0.659,0.449],[0.76,0.427]],
  [[0.896,0.382],[0.87,0.373],[0.889,0.378],[0.869,0.345],[0.886,0.349],[0.843,0.309],[0.817,0.382],
   [0.788,0.389],[0.761,0.426]],
  [[0.734,0.357],[0.725,0.262],[0.68,0.272],[0.679,0.402],[0.715,0.39]],
  [[1,0.069],[0.981,0.048],[0.973,0.006],[0.949,0.003],[0.933,0.082],[0.949,0.148],[0.973,0.089],[0.979,0.098]],
  [[0.862,0.493],[0.813,0.463],[0.773,0.486],[0.827,0.564]],
  [[0.822,0.306],[0.824,0.323],[0.844,0.301],[0.853,0.317],[0.843,0.309],[0.817,0.382],[0.791,0.397],
   [0.774,0.363],[0.799,0.316],[0.799,0.283],[0.805,0.311]],
];

function slide17() {
  const sl = pptx.addSlide();
  navBar(sl);
  const box = [1.565, 1.749, 10.441, 5.174];
  WORLD.forEach(poly => landmass(sl, box, poly));

  [[8.455, 4.03, C.accent4], [9.959, 3.063, C.accent2], [6.334, 3.75, C.accent2],
   [10.172, 5.422, C.accent4], [8.326, 2.219, C.accent4], [5.104, 1.819, C.accent2],
   [4.534, 4.953, C.accent2], [2.815, 2.841, C.accent4]].forEach(([x, y, c]) => pin(sl, x, y, 0.311, c, C.lt2));

  txt(sl, [
    { text: 'World ' }, { text: 'Map', options: { color: C.accent4 } }, { text: ' Design Slide' },
  ], { x: 2.548, y: 0.586, w: 8.237, h: 0.707, fontSize: 36, bold: true, fontFace: HEAD, color: C.black, align: 'center' });

  txt(sl, 'passages Lorem Ipsum available, but the majority suffered alteration',
    { x: 0.743, y: 4.065, w: 1.993, h: 0.867, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });
  gradRoundRectH(sl, 0.743, 5.411, 2.946, 0.796, 0.1758, C.accent2, C.accent4);
  moreDetail(sl, 1.124, 5.59, 3.055, 5.675, C.white, 20, 1.9);

  txt(sl, '4,67%', { x: 11.341, y: 4.179, w: 1.634, h: 0.64, fontSize: 32, bold: true, color: C.accent4 });
  seg(sl, 11.125, 4.932, 13.333, 4.932, C.accent5, 1);

  snowflake(sl, 1.473, 0.923, 0.751, 0.867, G.flakeIce);
  snowflake(sl, 7.972, 6.239, 1.283, 1.479, G.flakeAqua);
}

function slide18() {
  const sl = pptx.addSlide();
  navBar(sl);
  const box = [0.539, 1.584, 7.299, 4.87];
  USA.forEach(poly => landmass(sl, box, poly));

  [[2.598, 1.555, C.accent2], [1.404, 2.513, C.accent4], [3.887, 3.916, C.accent4],
   [5.033, 2.4, C.accent2], [1.832, 4.418, C.accent2], [6.415, 3.085, C.accent4]]
    .forEach(([x, y, c]) => pin(sl, x, y, 0.431, c, C.white));

  gradRoundRectH(sl, 4.979, 5.504, 1.151, 1.151, 0.1, C.accent2, C.accent4);
  icon(sl, 'route', 5.554, 6.080, 0.58, C.white);

  txt(sl, [
    { text: 'Australia ' }, { text: 'Map ', options: { color: C.accent4 } }, { text: 'Design Slide' },
  ], { x: 8.43, y: 1.996, w: 3.826, h: 1.313, fontSize: 36, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'Emerges From The Shadows',
    { x: 8.438, y: 3.649, w: 3.865, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD, color: C.accent4 });
  txt(sl, 'Maecenas congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Lorem ipsum',
    { x: 8.43, y: 3.985, w: 3.899, h: 0.867, fontSize: 10.5, color: C.black, lineSpacingMultiple: 1.5 });
  moreDetail(sl, 8.438, 5.134, 10.05, 5.232, C.accent4);

  snowflake(sl, 5.894, 0.687, 0.846, 0.975, G.flakeAqua);
  snowflake(sl, 10.596, 5.478, 2.26, 2.607, G.flakeIce);
}

/* ================================================================== *
 * Slide 19 -- contact
 * ================================================================== */
function slide19() {
  const sl = pptx.addSlide();
  snowflake(sl, 6.597, -0.097, 6.669, 7.694, [C.navy, C.frost], null, 32);
  navBar(sl);
  snowflake(sl, 3.242, -1.622, 2.26, 2.607, G.flakeIce);
  snowflake(sl, -0.197, 6.346, 0.846, 0.975, G.flakeAqua);

  txt(sl, [
    { text: 'Our Contact ', options: { color: C.accent4 } },
    { text: 'Information' },
  ], { x: 0.912, y: 1.362, w: 5.114, h: 2.121, fontSize: 60, bold: true, fontFace: HEAD, color: C.black });
  txt(sl, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque non elit mauris. Cras euismod, metus ac finibus finibuspurus, ',
    { x: 0.926, y: 3.622, w: 4.643, h: 0.602, fontSize: 10.5, color: C.grey40, lineSpacingMultiple: 1.5 });

  const contacts = [
    ['phone', 1.022, 4.890, 1.528, 4.970, 1.714, '+123-456-7890'],
    ['mail', 3.481, 4.894, 4.004, 4.976, 1.894, 'hello@reallygreatsite.com'],
    ['globe', 1.022, 5.725, 1.528, 5.793, 1.788, 'www.reallygreatsite.com'],
    ['home', 3.481, 5.725, 4.004, 5.793, 1.894, '123 Anywhere St., Any City'],
  ];
  contacts.forEach(([kind, bx, by, tx, ty, tw, label]) => {
    roundRect(sl, bx, by, 0.413, 0.413, 0.0938, C.accent2);
    icon(sl, kind, bx + 0.2065, by + 0.2065, 0.2, C.white);
    txt(sl, label, { x: tx, y: ty, w: tw, h: 0.252, fontSize: 9, bold: true, fontFace: HEAD, color: C.black, valign: 'middle' });
  });
}

/* ================================================================== *
 * Slide 20 -- thanks
 * ================================================================== */
function slide20() {
  const sl = pptx.addSlide();
  gradientBackground(sl, 315, [[0, '3E9AD9'], [0.5, '46829F'], [1, '4C6684']]);
  roundRect(sl, 1.151, 0.887, 4.281, 5.895, 2.1405, C.navy);
  snowflake(sl, 3.521, 0.184, 2.435, 2.798, 'EFF6FC', HAIRLINE);
  snowflake(sl, 10.237, 5.342, 2.671, 3.081, 'E6F1FA', HAIRLINE);
  snowflake(sl, -0.44, 5.85, 1.066, 1.229, 'E6F1FA', HAIRLINE);

  txt(sl, 'Thanks For', { x: 6.359, y: 1.886, w: 7.542, h: 1.447, fontSize: 80, bold: true, fontFace: HEAD, color: C.accent3 });
  txt(sl, 'Attention ', { x: 6.359, y: 3.046, w: 6.664, h: 1.447, fontSize: 80, bold: true, fontFace: HEAD, color: C.white });
  txt(sl, "simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the,  ",
    { x: 6.359, y: 4.483, w: 5.024, h: 0.626, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
  txt(sl, 'Celebration Of Light And Renewal',
    { x: 6.359, y: 5.312, w: 2.951, h: 0.303, fontSize: 12, bold: true, fontFace: HEAD, color: C.white, valign: 'middle' });

  navBar(sl, { mark: C.white, word: C.white, dot: C.accent2, search: C.white });
}

/* ------------------------------------------------------------------ */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach(fn => fn());

pptx.writeFile({ fileName: path.join(__dirname, '00b2cbc8-bbff-44f2-9fef-648d5416581b_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
