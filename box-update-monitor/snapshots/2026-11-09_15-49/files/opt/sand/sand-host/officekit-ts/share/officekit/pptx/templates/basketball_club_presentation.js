/*
 * Dribbly — Basketball Club Presentation.
 * 33 slides at 20" x 11.25", rebuilt from scratch with pptxgenjs.
 *
 * The reference deck's picture placeholders ship empty, so they contribute no
 * pixels; the two raster mock-ups (phone, laptop) are redrawn here as native
 * shapes rather than embedded bitmaps.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  bg: '191919',        // theme background
  magenta: 'CC00FF',   // brand accent
  white: 'FFFFFF',
  gray: '595959',      // tx1 lumMod 65% + lumOff 35% — card fill
  track: 'BFBFBF',     // white lumMod 75% — progress-bar track
  arrow: '808080',     // white lumMod 50%
  ribbonDk: '404040',  // tx1 lumMod 75% + lumOff 25%
  subtle: 'D1D5DB',
  silhouette: '9607BA',// flat colour of the player-silhouette artwork
  bezel: 'B8BCC0',
};

// -------------------------------------------------------------- typography --
const F = {
  reg: 'Poppins',
  med: 'Poppins Medium',
  semi: 'Poppins SemiBold',
  xbold: 'Poppins ExtraBold',
  script: 'Lobster',
  brush: 'Yellowtail',
};
const SZ = { small: 18, body: 20, lead: 24, head: 48, huge: 150, script: 200, brush: 450, stat: 40 };

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE20', width: 20, height: 11.25 });
pptx.layout = 'WIDE20';
pptx.title = 'Basketball Club Presentation';

// ---------------------------------------------------------------- helpers ---

/** Filled rectangle, no outline. */
function rect(s, x, y, w, h, fill, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color: fill }, line: { type: 'none' } }, opts));
}

/** Outlined rectangle, no fill. */
function outline(s, x, y, w, h, color, width) {
  s.addShape('rect', { x, y, w, h, fill: { type: 'none' }, line: { color, width: width || 3 } });
}

/**
 * Text block matching the source text boxes. PowerPoint stacks lines from the
 * top of the box, so `valign` defaults to top. Text-frame insets are left
 * unset, exactly as in the reference, so the full box width is available.
 */
function text(s, str, x, y, w, h, o) {
  o = o || {};
  s.addText(str, {
    x, y, w, h,
    fontFace: o.font || F.reg,
    fontSize: o.size || SZ.body,
    color: o.color || C.white,
    bold: o.bold || false,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.ls,
    margin: o.margin,
    wrap: o.wrap === undefined ? true : o.wrap,
    rotate: o.rotate,
    isTextBox: true,
    fit: 'none',
  });
}

/** Text centred both ways inside a box — used for buttons and badges. */
function centered(s, str, x, y, w, h, o) {
  o = o || {};
  s.addText(str, {
    x, y, w, h, align: 'center', valign: 'middle', margin: 0,
    fontFace: o.font || F.semi, fontSize: o.size || SZ.body, color: o.color || C.white,
  });
}

/** Bold heading + body paragraph — the deck's ubiquitous "Section Content". */
function labelBlock(s, x, y, w, label, body, o) {
  o = o || {};
  const big = (o.labelSize || SZ.body) >= SZ.lead;
  text(s, label, x, y, w, big ? 0.505 : 0.438,
    { font: F.semi, size: o.labelSize || SZ.body, color: o.labelColor || C.white, align: o.align });
  text(s, body, x, y + (big ? 0.512 : 0.364), w, o.bodyH || 1.063,
    { size: o.bodySize || SZ.body, ls: 1.5, align: o.align });
}

// ------------------------------------------------------------------ icons ---
/*
 * Line-art icons are drawn from primitives so no bitmaps are needed:
 *   ['o'|'O', cx, cy, r]        circle — outline / filled
 *   ['e', x, y, w, h]           ellipse outline
 *   ['r'|'R', x, y, w, h]       rectangle — outline / filled
 *   ['l', x1, y1, x2, y2]       straight line
 *   ['p'|'P', [[x,y], ...]]     polygon — filled / outline
 * All coordinates are fractions of the icon's bounding box.
 */
const ICONS = {
  ball: [['o', 0.5, 0.5, 0.47], ['l', 0.5, 0.03, 0.5, 0.97], ['l', 0.03, 0.5, 0.97, 0.5], ['l', 0.15, 0.15, 0.85, 0.85]],
  hoop: [['o', 0.33, 0.30, 0.28], ['l', 0.10, 0.56, 0.95, 0.56], ['l', 0.18, 0.56, 0.36, 0.98],
    ['l', 0.88, 0.56, 0.70, 0.98], ['l', 0.36, 0.98, 0.70, 0.98], ['l', 0.40, 0.56, 0.48, 0.98],
    ['l', 0.66, 0.56, 0.58, 0.98], ['l', 0.26, 0.77, 0.80, 0.77]],
  trophy: [['P', [[0.20, 0.04], [0.80, 0.04], [0.70, 0.55], [0.30, 0.55]]], ['l', 0.5, 0.55, 0.5, 0.74],
    ['r', 0.28, 0.74, 0.44, 0.10], ['r', 0.18, 0.88, 0.64, 0.10],
    ['l', 0.20, 0.10, 0.06, 0.10], ['l', 0.06, 0.10, 0.10, 0.34], ['l', 0.10, 0.34, 0.26, 0.40],
    ['l', 0.80, 0.10, 0.94, 0.10], ['l', 0.94, 0.10, 0.90, 0.34], ['l', 0.90, 0.34, 0.74, 0.40]],
  medal: [['o', 0.5, 0.33, 0.31], ['o', 0.5, 0.33, 0.16],
    ['P', [[0.28, 0.58], [0.72, 0.58], [0.72, 0.99], [0.50, 0.86], [0.28, 0.99]]]],
  court: [['r', 0.10, 0.03, 0.80, 0.94], ['l', 0.10, 0.5, 0.90, 0.5], ['o', 0.5, 0.5, 0.13],
    ['r', 0.28, 0.03, 0.44, 0.16], ['r', 0.28, 0.81, 0.44, 0.16]],
  pin: [['o', 0.5, 0.36, 0.30], ['o', 0.5, 0.36, 0.11], ['l', 0.22, 0.52, 0.5, 0.99], ['l', 0.78, 0.52, 0.5, 0.99]],
  mail: [['r', 0.04, 0.20, 0.92, 0.60], ['l', 0.04, 0.20, 0.50, 0.58], ['l', 0.96, 0.20, 0.50, 0.58]],
  phone: [['P', [[0.12, 0.06], [0.34, 0.04], [0.44, 0.30], [0.30, 0.42], [0.60, 0.72], [0.72, 0.58],
    [0.97, 0.68], [0.95, 0.90], [0.72, 0.98], [0.36, 0.80], [0.14, 0.50], [0.04, 0.24]]]],
  globe: [['o', 0.5, 0.5, 0.47], ['e', 0.28, 0.03, 0.44, 0.94], ['l', 0.03, 0.5, 0.97, 0.5],
    ['l', 0.13, 0.22, 0.87, 0.22], ['l', 0.13, 0.78, 0.87, 0.78]],
  people: [['O', 0.5, 0.26, 0.17], ['O', 0.13, 0.34, 0.13], ['O', 0.87, 0.34, 0.13],
    ['p', [[0.24, 0.78], [0.28, 0.52], [0.72, 0.52], [0.76, 0.78]]],
    ['p', [[0.0, 0.78], [0.02, 0.60], [0.22, 0.60], [0.20, 0.78]]],
    ['p', [[0.78, 0.60], [0.98, 0.60], [1.0, 0.78], [0.80, 0.78]]]],
  bulb: [['o', 0.5, 0.35, 0.30], ['l', 0.38, 0.60, 0.38, 0.70], ['l', 0.62, 0.60, 0.62, 0.70],
    ['R', 0.34, 0.72, 0.32, 0.055], ['R', 0.36, 0.82, 0.28, 0.055], ['R', 0.40, 0.92, 0.20, 0.055]],
  rocket: [['p', [[0.98, 0.02], [0.96, 0.26], [0.86, 0.40], [0.44, 0.76], [0.22, 0.54], [0.58, 0.12], [0.72, 0.03]]],
    ['p', [[0.34, 0.28], [0.52, 0.30], [0.26, 0.56], [0.10, 0.48], [0.16, 0.36]]],
    ['p', [[0.44, 0.76], [0.70, 0.82], [0.60, 0.94], [0.44, 0.90], [0.34, 0.94]]],
    ['l', 0.24, 0.78, 0.02, 0.98], ['l', 0.16, 0.66, 0.02, 0.80], ['l', 0.36, 0.90, 0.20, 0.99]],
  puzzle: [['p', [[0.06, 0.06], [0.40, 0.06], [0.36, 0.18], [0.44, 0.28], [0.56, 0.28], [0.62, 0.18],
    [0.58, 0.06], [0.94, 0.06], [0.94, 0.40], [0.82, 0.36], [0.72, 0.44], [0.72, 0.56], [0.82, 0.62],
    [0.94, 0.58], [0.94, 0.94], [0.58, 0.94], [0.62, 0.82], [0.56, 0.72], [0.44, 0.72], [0.36, 0.82],
    [0.40, 0.94], [0.06, 0.94]]]],
  twitter: [['p', [[0.98, 0.14], [0.80, 0.24], [0.86, 0.10], [0.66, 0.16], [0.44, 0.14], [0.28, 0.28],
    [0.30, 0.44], [0.06, 0.28], [0.04, 0.46], [0.16, 0.58], [0.08, 0.58], [0.20, 0.74], [0.30, 0.76],
    [0.10, 0.86], [0.38, 0.92], [0.68, 0.82], [0.84, 0.54], [0.84, 0.32]]]],
  instagram: [['r', 0.05, 0.05, 0.90, 0.90], ['o', 0.5, 0.5, 0.27], ['O', 0.76, 0.24, 0.06]],
  facebook: [['p', [[0.66, 0.02], [0.92, 0.02], [0.92, 0.24], [0.74, 0.24], [0.68, 0.30], [0.68, 0.40],
    [0.92, 0.40], [0.88, 0.62], [0.68, 0.62], [0.68, 0.99], [0.42, 0.99], [0.42, 0.62], [0.22, 0.62],
    [0.22, 0.40], [0.42, 0.40], [0.42, 0.26], [0.50, 0.09]]]],
};

/** Renders one icon from ICONS at (x, y) inside a square of `size` inches. */
function icon(s, name, x, y, size, color) {
  const stroke = Math.max(0.9, size * 2.2);
  const col = color || C.white;
  (ICONS[name] || []).forEach(part => {
    const kind = part[0];
    if (kind === 'o' || kind === 'O') {
      const r = part[3];
      s.addShape('ellipse', {
        x: x + (part[1] - r) * size, y: y + (part[2] - r) * size, w: 2 * r * size, h: 2 * r * size,
        fill: kind === 'O' ? { color: col } : { type: 'none' },
        line: kind === 'O' ? { type: 'none' } : { color: col, width: stroke },
      });
    } else if (kind === 'e') {
      s.addShape('ellipse', {
        x: x + part[1] * size, y: y + part[2] * size, w: part[3] * size, h: part[4] * size,
        fill: { type: 'none' }, line: { color: col, width: stroke },
      });
    } else if (kind === 'r' || kind === 'R') {
      s.addShape('rect', {
        x: x + part[1] * size, y: y + part[2] * size, w: part[3] * size, h: part[4] * size,
        fill: kind === 'R' ? { color: col } : { type: 'none' },
        line: kind === 'R' ? { type: 'none' } : { color: col, width: stroke },
      });
    } else if (kind === 'l') {
      const [, x1, y1, x2, y2] = part;
      s.addShape('line', {
        x: x + Math.min(x1, x2) * size, y: y + Math.min(y1, y2) * size,
        w: Math.abs(x2 - x1) * size, h: Math.abs(y2 - y1) * size,
        line: { color: col, width: stroke }, flipH: (x2 - x1) * (y2 - y1) < 0,
      });
    } else {
      // Polygon points are relative to the shape's own bounding box.
      const pts = part[1].map((p, i) => ({ x: +(p[0] * size).toFixed(3), y: +(p[1] * size).toFixed(3), moveTo: i === 0 }));
      pts.push({ close: true });
      s.addShape('custGeom', {
        x, y, w: size, h: size, points: pts,
        fill: kind === 'p' ? { color: col } : { type: 'none' },
        line: kind === 'p' ? { type: 'none' } : { color: col, width: stroke },
      });
    }
  });
}

/** Magenta square tile with a white icon centred on it. */
function iconTile(s, x, y, size, name) {
  rect(s, x, y, size, size, C.magenta);
  icon(s, name, x + size * 0.19, y + size * 0.19, size * 0.62);
}

/** Free-form polygon from normalised [0..1] coordinates. */
function poly(s, x, y, w, h, pts, fill, opts) {
  const points = pts.map((p, i) => ({ x: +(p[0] * w).toFixed(3), y: +(p[1] * h).toFixed(3), moveTo: i === 0 }));
  points.push({ close: true });
  s.addShape('custGeom', Object.assign({ x, y, w, h, points, fill: { color: fill }, line: { type: 'none' } }, opts));
}

/** Black scrim laid over artwork; `opacity` is the black coverage percentage. */
function scrim(s, x, y, w, h, opacity) {
  s.addShape('rect', { x, y, w, h, fill: { color: '000000', transparency: 100 - opacity }, line: { type: 'none' } });
}

/**
 * Parallelogram matching the OOXML preset: `adj` is the horizontal offset as a
 * fraction of the shape's shorter side (the preset default is 0.25).
 */
function para(s, x, y, w, h, adj, fill) {
  const k = Math.min(w, h) * adj / w;
  poly(s, x, y, w, h, [[k, 0], [1, 0], [1 - k, 1], [0, 1]], fill);
}

/** The 1.3" magenta logo badge that anchors a corner of most slides. */
function logoBadge(s, x, y) {
  rect(s, x, y, 1.3, 1.3, C.magenta);
  icon(s, 'hoop', x + 0.3, y + 0.3, 0.7);
}

/** Vertical rule plus stacked social glyphs — the deck's left/right side rail. */
const SOCIAL = ['twitter', 'instagram', 'facebook'];
function sideRail(s, side) {
  const x = side === 'left' ? 0 : 18.7;
  logoBadge(s, x, 0);
  s.addShape('line', { x: x + 0.65, y: 1.838, w: 0, h: 6.559, line: { color: C.white, width: 3 } });
  SOCIAL.forEach((g, i) => icon(s, g, x + 0.475, 8.929 + i * 0.71, 0.35));
}

/** "Section progress ... 80%" label pair over a two-tone bar. */
function progressBar(s, x, y, w, label, pct, o) {
  const font = (o && o.font) || F.semi;
  text(s, label, x, y, 2.6, 0.438, { font, wrap: false });
  text(s, pct + '%', x + w - 0.8, y, 0.8, 0.438, { font, align: 'right', wrap: false });
  rect(s, x + 0.06, y + 0.538, w - 0.06, 0.15, C.track);
  rect(s, x + 0.06, y + 0.538, (w - 0.06) * pct / 100, 0.15, C.magenta);
}

/** Oversized magenta figure with a caption below. */
function statBlock(s, x, y, value, caption) {
  text(s, value, x, y, 2.057, 0.62, { font: F.xbold, size: SZ.stat, color: C.magenta, align: 'center', wrap: false });
  text(s, caption, x, y + 0.62, 2.057, 0.44, { font: F.semi, align: 'center', wrap: false });
}

/** Filled call-to-action button. */
function buttonSolid(s, x, y, label) {
  rect(s, x, y, 2.633, 0.8, C.magenta);
  centered(s, label, x, y, 2.633, 0.8, { font: F.med });
}

/** Outlined call-to-action button. */
function buttonGhost(s, x, y, label) {
  outline(s, x, y, 2.633, 0.8, C.magenta);
  centered(s, label, x, y, 2.633, 0.8, { font: F.med });
}

// --------------------------------------------------------- shared geometry --
// Two overlapping diagonal "slash" ribbons form the deck's decorative motif.
const SLASH_A = [[0.191, 0.713], [0.525, 0.713], [0.335, 1.0], [0.0, 1.0]];
const SLASH_A2 = [[0.665, 0.0], [1.0, 0.0], [0.586, 0.622], [0.252, 0.622]];
const SLASH_B = [[0.515, 0.226], [0.850, 0.226], [0.335, 1.0], [0.0, 1.0]];
const SLASH_B2 = [[0.665, 0.0], [1.0, 0.0], [0.911, 0.134], [0.576, 0.134]];

/** Paired slash motif; `x` is the left edge of the first ribbon. */
function slashPair(s, x) {
  poly(s, x, 0, 4.9, 11.25, SLASH_A, C.magenta);
  poly(s, x, 0, 4.9, 11.25, SLASH_A2, C.magenta);
  poly(s, x + 2.01, 0.016, 4.9, 11.207, SLASH_B, C.magenta);
  poly(s, x + 2.01, 0.016, 4.9, 11.207, SLASH_B2, C.magenta);
}

/**
 * Device mock-up: a hollow rounded-rectangle bezel whose screen stays
 * transparent, so slide artwork shows through it just as in the reference.
 * Drawn as a thick black stroke sandwiched between two thin silver highlights.
 */
function deviceBezel(s, x, y, w, h, thick, radius) {
  const t = thick / 2;
  s.addShape('roundRect', { x: x + t, y: y + t, w: w - thick, h: h - thick, rectRadius: radius,
    fill: { type: 'none' }, line: { color: '060606', width: thick * 72 } });
  s.addShape('roundRect', { x, y, w, h, rectRadius: radius,
    fill: { type: 'none' }, line: { color: C.bezel, width: 2 } });
  s.addShape('roundRect', { x: x + thick, y: y + thick, w: w - 2 * thick, h: h - 2 * thick,
    rectRadius: Math.max(0.01, radius - thick), fill: { type: 'none' }, line: { color: '4A4A4C', width: 1 } });
}

// A rounded chevron arrow-head, repeated in rows on the timeline slide.
const CHEVRON = [[0.222, 0.0], [0.338, 0.011], [0.409, 0.043], [0.933, 0.391], [0.970, 0.424], [1.0, 0.5],
  [0.970, 0.576], [0.933, 0.609], [0.409, 0.957], [0.283, 0.999], [0.132, 0.987], [0.062, 0.954],
  [0.017, 0.905], [0.0, 0.849], [0.016, 0.792], [0.068, 0.739], [0.264, 0.609], [0.324, 0.539],
  [0.324, 0.461], [0.301, 0.424], [0.068, 0.261], [0.016, 0.208], [0.0, 0.151], [0.017, 0.095],
  [0.062, 0.046], [0.132, 0.013]];

/**
 * Samples a left-to-right gradient at position `t` (0..1). `stops` is an
 * ordered list of [position, hexColour] pairs; colours hold flat outside them.
 */
function gradientAt(stops, t) {
  let lo = stops[0];
  let hi = stops[stops.length - 1];
  for (let i = 0; i < stops.length - 1; i++) {
    if (t >= stops[i][0] && t <= stops[i + 1][0]) { lo = stops[i]; hi = stops[i + 1]; break; }
  }
  if (t <= stops[0][0]) return stops[0][1];
  if (t >= hi[0] && hi === stops[stops.length - 1] && t > hi[0]) return hi[1];
  const k = hi[0] === lo[0] ? 0 : (t - lo[0]) / (hi[0] - lo[0]);
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const av = parseInt(lo[1].substr(i, 2), 16);
    const bv = parseInt(hi[1].substr(i, 2), 16);
    out += Math.round(av + (bv - av) * k).toString(16).padStart(2, '0').toUpperCase();
  }
  return out;
}

// Player silhouettes traced from the reference artwork (normalised outlines).
const SIL = {
  DRIBBLE: [[0.902,0.751],[0.828,0.705],[0.761,0.704],[0.657,0.660],[0.569,0.648],[0.593,0.623],[0.582,0.593],[0.589,0.483],[0.576,0.447],[0.519,0.381],[0.470,0.356],[0.446,0.246],[0.542,0.267],[0.744,0.269],[0.786,0.322],[0.872,0.343],[0.937,0.332],[0.982,0.302],[1.000,0.260],[0.982,0.218],[0.937,0.188],[0.872,0.176],[0.805,0.189],[0.758,0.221],[0.610,0.202],[0.505,0.178],[0.479,0.161],[0.471,0.125],[0.426,0.096],[0.300,0.091],[0.296,0.074],[0.283,0.075],[0.280,0.034],[0.240,0.008],[0.178,0.002],[0.119,0.029],[0.109,0.073],[0.123,0.145],[0.053,0.181],[0.007,0.245],[0.000,0.352],[0.012,0.459],[0.042,0.539],[0.182,0.692],[0.220,0.821],[0.251,0.830],[0.250,0.765],[0.256,0.818],[0.273,0.818],[0.276,0.762],[0.298,0.765],[0.287,0.704],[0.382,0.779],[0.488,0.836],[0.676,0.891],[0.694,0.908],[0.698,0.935],[0.681,0.994],[0.766,0.999],[0.800,0.979],[0.795,0.944],[0.815,0.918],[0.816,0.883],[0.762,0.854],[0.619,0.809],[0.567,0.773],[0.489,0.742],[0.451,0.700],[0.546,0.740],[0.720,0.746],[0.937,0.872],[0.976,0.861],[0.977,0.821],[0.964,0.796],[0.902,0.751]],
  LAYUP: [[1.000,0.651],[0.932,0.506],[0.869,0.454],[0.831,0.392],[0.809,0.327],[0.728,0.281],[0.664,0.262],[0.692,0.220],[0.723,0.204],[0.727,0.175],[0.700,0.148],[0.647,0.135],[0.595,0.143],[0.561,0.173],[0.558,0.228],[0.533,0.242],[0.374,0.231],[0.217,0.253],[0.087,0.149],[0.096,0.119],[0.164,0.120],[0.231,0.092],[0.246,0.061],[0.231,0.030],[0.191,0.008],[0.113,0.001],[0.059,0.016],[0.007,0.048],[0.009,0.099],[0.095,0.241],[0.149,0.295],[0.361,0.309],[0.392,0.419],[0.366,0.439],[0.379,0.457],[0.337,0.472],[0.330,0.558],[0.366,0.618],[0.440,0.688],[0.388,0.730],[0.375,0.780],[0.240,0.789],[0.173,0.760],[0.130,0.764],[0.015,0.862],[0.003,0.895],[0.042,0.903],[0.104,0.873],[0.212,0.858],[0.241,0.839],[0.375,0.830],[0.362,0.840],[0.348,0.904],[0.286,0.956],[0.511,0.998],[0.452,0.923],[0.460,0.907],[0.442,0.885],[0.451,0.840],[0.439,0.835],[0.533,0.830],[0.576,0.815],[0.626,0.689],[0.641,0.685],[0.637,0.610],[0.661,0.535],[0.657,0.494],[0.680,0.474],[0.690,0.437],[0.744,0.383],[0.832,0.527],[0.938,0.614],[0.930,0.665],[0.943,0.706],[0.958,0.702],[0.979,0.721],[0.997,0.713],[0.987,0.682],[1.000,0.651]],
  GUARD: [[0.489,0.000],[0.529,0.004],[0.568,0.027],[0.580,0.043],[0.584,0.112],[0.689,0.148],[0.861,0.324],[0.894,0.378],[0.924,0.484],[0.996,0.563],[0.992,0.602],[0.960,0.555],[0.992,0.605],[0.985,0.627],[0.949,0.610],[0.906,0.567],[0.902,0.607],[0.889,0.610],[0.870,0.474],[0.794,0.365],[0.692,0.306],[0.703,0.350],[0.773,0.435],[0.817,0.553],[0.798,0.574],[0.801,0.597],[0.877,0.719],[0.897,0.824],[0.924,0.858],[0.934,0.966],[0.908,0.990],[0.866,0.999],[0.849,0.984],[0.832,0.829],[0.777,0.766],[0.699,0.622],[0.616,0.643],[0.550,0.565],[0.468,0.625],[0.365,0.597],[0.290,0.676],[0.272,0.744],[0.226,0.817],[0.230,0.856],[0.219,0.876],[0.104,0.943],[0.077,0.943],[0.059,0.921],[0.059,0.902],[0.081,0.879],[0.147,0.829],[0.170,0.825],[0.181,0.698],[0.208,0.616],[0.250,0.537],[0.373,0.388],[0.380,0.337],[0.358,0.296],[0.305,0.323],[0.177,0.316],[0.151,0.361],[0.112,0.376],[0.065,0.376],[0.012,0.348],[0.003,0.302],[0.064,0.252],[0.193,0.269],[0.286,0.248],[0.318,0.163],[0.437,0.127],[0.416,0.083],[0.415,0.043],[0.443,0.010],[0.489,0.000]],
  DRIVE: [[0.832,0.000],[0.887,0.028],[0.902,0.087],[0.918,0.106],[0.910,0.156],[0.841,0.178],[0.898,0.209],[0.864,0.319],[0.960,0.342],[0.999,0.394],[0.989,0.437],[0.948,0.471],[0.872,0.488],[0.792,0.470],[0.752,0.534],[0.835,0.686],[0.711,0.912],[0.775,0.956],[0.839,0.971],[0.837,0.991],[0.789,1.000],[0.593,0.974],[0.594,0.943],[0.651,0.844],[0.663,0.772],[0.721,0.706],[0.714,0.699],[0.661,0.723],[0.620,0.670],[0.580,0.764],[0.559,0.768],[0.519,0.811],[0.108,0.864],[0.089,0.898],[0.081,0.954],[0.030,0.947],[0.015,0.925],[0.000,0.822],[0.015,0.799],[0.047,0.794],[0.096,0.814],[0.291,0.758],[0.428,0.742],[0.370,0.708],[0.456,0.574],[0.456,0.504],[0.488,0.386],[0.435,0.370],[0.428,0.341],[0.574,0.183],[0.632,0.155],[0.724,0.141],[0.698,0.069],[0.705,0.047],[0.745,0.022],[0.832,0.000]],};

// ------------------------------------------------------------- copy blocks --
const KICKER = 'Basketball Club Presentation';
const LOREM = 'Donec purus felis, mauris volutpat in auctor sit qui amet, euismod ut.';
const LOREM_SHORT = 'Donec purus felis, mauris volutpat in auctor sit qui.';
const LOREM_TINY = 'Donec purus felis, qui mauris volutpat in sit.';
const LOREM_LONG = 'Donec purus felis, mauris volutpat in auctor sit qui amet, euismod ut lorem.';
const LOREM_NODOT = 'Donec purus felis, mauris volutpat in auctor sit qui';
const SWOT_BODY = 'Donec purus felis, mauris volutpat in auctor sit qui. Torto Nullam volutpat libero ut tristique condimentum.';
const VESTIBULUM = 'Vestibulum pellentesque magna sapien, in porttitor arcu volutpat quis. Mauris gravida, mi eu commodo.';
const NUNC = 'Nunc ut nibh gravida, varius metus sed, finibus ligula. dolor venenatis lacus, ac pulvinar massa orci vel diam.';
const TINCIDUNT = 'Donec tincidunt id lectus eget finibus. Nam egestas libero vitae dui tincidunt, eu malesuada nisl dictum. Sed sollicitudin leo ne morbi a id arcu. ';
const INTEGER = 'Integer magna ex, volutpat non nibh at';
const ALIQUAM = 'Aliquam erat volutpat. Nam varius, lorem id blandit rutrum, risus metus luctus massa.';
const SECTION = 'Section Content';

function newSlide() {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  return s;
}

// ================================================================ SLIDE 1 ===
// Title: slash motif behind a 70% scrim, script wordmark, left rail.
function slide01() {
  const s = newSlide();
  slashPair(s, 6.529);
  scrim(s, 0, 0, 20, 11.25, 70);
  sideRail(s, 'left');
  text(s, 'Dribbly', 5.786, 2.931, 8.427, 3.45, { font: F.script, size: SZ.script, wrap: false });
  text(s, 'Donec nulla lacus, ornare nec pharetra', 5.839, 5.799, 6.875, 0.438, { font: F.med });
  text(s, KICKER, 17.187, 8.311, 4.322, 0.438, { font: F.semi, rotate: 270, wrap: false });
}

// ================================================================ SLIDE 2 ===
// Hero: magenta portrait panel left, headline + CTA cluster right.
function slide02() {
  const s = newSlide();
  rect(s, 1.271, 1.445, 5.317, 7.365, C.magenta);
  sideRail(s, 'right');

  text(s, KICKER, 9.086, 1.654, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'First Dribble To Start My Basketball Journey', 9.086, 2.092, 8.833, 1.717, { font: F.semi, size: SZ.head });
  text(s, 'Sed sollicitudin leo nec sollicitudin efficitur. Nulla sapien tellus, porta in rutrum eu, pretium id arcu. ',
    9.086, 4.124, 8.833, 1.063, { ls: 1.5 });

  iconTile(s, 9.15, 6.003, 1.0, 'ball');
  labelBlock(s, 10.378, 5.907, 5.161, SECTION, LOREM);

  rect(s, 3.521, 7.622, 4.617, 2.35, C.gray);
  labelBlock(s, 3.988, 7.985, 3.683, SECTION, LOREM_SHORT);

  buttonSolid(s, 9.086, 8.257, 'Join Now');
  buttonGhost(s, 11.953, 8.251, 'Learn More');
}

// ================================================================ SLIDE 3 ===
// Captain profile with three progress bars; slash motif far right.
function slide03() {
  const s = newSlide();
  slashPair(s, 11.449);
  sideRail(s, 'left');

  text(s, 'Captain Team', 2.15, 1.449, 2.243, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'Jonathan Gurley', 2.15, 1.928, 6.698, 0.909, { font: F.semi, size: SZ.head });
  text(s, 'Aenean eget metus varius, euismo sit justo a, volutpat augue. Morbi euismod eros justo, sit amet euismod nibh vitae.',
    2.15, 3.556, 7.972, 1.568, { ls: 1.5 });

  [5.905, 6.872, 7.839].forEach(y => progressBar(s, 2.204, y, 7.919, 'Section progress', 80));
}

// ================================================================ SLIDE 4 ===
// Team roster: four parallelogram name plates with roles and captions.
const ROSTER = [
  { plate: 0.682, name: 'Samuel Mitchell', nx: 1.351, nw: 2.935, role: 'Shooting Guard', rx: 1.589, rw: 2.437, bx: 1.043 },
  { plate: 5.099, name: 'Matthew Davis', nx: 5.798, nw: 2.735, role: 'Small Forward', rx: 6.113, rw: 2.248, bx: 5.461 },
  { plate: 9.523, name: 'Zanele Nkosi', nx: 10.457, nw: 2.369, role: 'Center', rx: 11.053, rw: 1.173, bx: 9.878 },
  { plate: 13.934, name: 'Satoshi Mori', nx: 14.922, nw: 2.306, role: 'Power Forward', rx: 14.918, rw: 2.313, bx: 14.296 },
];
function slide04() {
  const s = newSlide();
  text(s, 'Basketball Team', 6.667, 1.171, 6.667, 0.909, { font: F.semi, size: SZ.head, align: 'center' });
  text(s, 'Cras vel nisl rutrum, porta libero id, pretium lorem. Ut ornare ex id felis maximus viverra. Aliquam vitae tristique mauris, at pellentesque eros. ',
    4.395, 2.184, 11.209, 0.774, { align: 'center' });

  ROSTER.forEach(p => {
    para(s, p.plate, 7.936, 4.275, 1.0, 0.25, C.magenta);
    text(s, p.name, p.nx, 7.998, p.nw, 0.505, { font: F.semi, size: SZ.lead, align: 'center', wrap: false });
    text(s, p.role, p.rx, 8.45, p.rw, 0.438, { font: F.semi, align: 'center', wrap: false });
    text(s, 'Quisque hendrerit odio purus, nec vestibulum.', p.bx, 9.119, 3.551, 0.774, { align: 'center' });
  });
}

// ================================================================ SLIDE 5 ===
// About: headline, two paragraphs, four icon captions on a grey strip.
const SECTIONS_5 = [
  { gx: 2.434, gw: 1.902, ix: 2.985, label: 'Section One' },
  { gx: 5.309, gw: 1.895, ix: 5.857, label: 'Section Two' },
  { gx: 8.067, gw: 2.122, ix: 8.728, label: 'Section Three' },
  { gx: 11.027, gw: 1.939, ix: 11.596, label: 'Section Four' },
];
function slide05() {
  const s = newSlide();
  para(s, 15.647, 0, 3.917, 11.25, 0.547, C.magenta);
  sideRail(s, 'left');

  text(s, KICKER, 1.835, 1.466, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'Community, Camaraderie, Basketball Club', 1.835, 1.904, 10.517, 1.717, { font: F.semi, size: SZ.head });
  text(s, 'Pellentesque laoreet nulla eget hendrerit ultricies. Duis lobortis finibus imperdiet. Nullam eros quam, blandit sit amet bibendum.',
    1.835, 4.113, 9.317, 1.063, { ls: 1.5 });
  text(s, VESTIBULUM, 1.835, 5.421, 9.317, 1.063, { ls: 1.5 });

  rect(s, 1.835, 7.429, 11.733, 1.851, C.gray);
  SECTIONS_5.forEach(c => {
    icon(s, 'ball', c.ix, 7.659, 0.8);
    text(s, c.label, c.gx, 8.612, c.gw, 0.438, { font: F.med, align: 'center', wrap: false });
  });
}

// ================================================================ SLIDE 6 ===
// "Why Join": three feature cards, the last one highlighted magenta.
const CARDS_6 = [
  { x: 10.197, y: 3.081, fill: C.gray, icon: 'trophy' },
  { x: 2.688, y: 5.819, fill: C.gray, icon: 'medal' },
  { x: 10.197, y: 5.819, fill: C.magenta, icon: 'ball' },
];
function slide06() {
  const s = newSlide();
  scrim(s, 0, 0, 20, 11.25, 70);
  icon(s, 'hoop', 1.083, 0.617, 0.8);
  text(s, 'Why Join Our Basketball Club', 3.044, 3.403, 6.403, 1.717, { font: F.semi, size: SZ.head });

  CARDS_6.forEach(c => {
    rect(s, c.x, c.y, 7.066, 2.35, c.fill);
    icon(s, c.icon, c.x + 0.306, c.y + 0.699, 1.0);
    labelBlock(s, c.x + 1.598, c.y + 0.388, 5.161, SECTION, LOREM);
  });
}

// ================================================================ SLIDE 7 ===
// Event promo: date card top-left, copy + stat + CTAs on the right.
function slide07() {
  const s = newSlide();
  sideRail(s, 'right');

  rect(s, 4.732, 1.152, 6.04, 2.333, C.gray);
  rect(s, 5.003, 1.468, 1.7, 1.7, C.magenta);
  s.addText([
    { text: '22', options: { fontSize: 32, breakLine: true } },
    { text: 'August', options: { fontSize: SZ.body } },
  ], { x: 5.003, y: 1.468, w: 1.7, h: 1.7, align: 'center', valign: 'middle', fontFace: F.semi, color: C.white, margin: 0 });
  s.addText([
    { text: 'Los Angeles', options: { fontFace: F.med, fontSize: SZ.small, breakLine: true } },
    { text: 'Basketball Champion', options: { fontFace: F.semi, fontSize: SZ.body } },
  ], { x: 6.854, y: 1.468, w: 3.688, h: 0.74, valign: 'top', color: C.white, margin: 0.05 });
  text(s, 'Nulla mollis dolor iaculis ad posuere molestie.', 6.854, 2.21, 3.688, 0.967, { size: SZ.small, ls: 1.5 });

  text(s, KICKER, 11.718, 1.537, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'Basketball Event', 11.718, 1.975, 6.852, 0.909, { font: F.semi, size: SZ.head });
  text(s, "HoopsFest: A Slammin' Basketball Extravaganza", 11.718, 2.985, 6.852, 1.255,
    { font: F.med, size: SZ.lead, color: C.subtle, ls: 1.5 });
  text(s, 'Mauris id odio imperdiet, pellentesque magna sit amet, efficitur mi. Suspendisse ipsum dolor, faucibus vel dolor facilisis.',
    11.718, 4.654, 6.852, 1.568, { ls: 1.5 });

  statBlock(s, 11.718, 6.872, '+26%', 'Lorem Ipsum');
  text(s, LOREM_SHORT, 14.392, 6.828, 4.179, 1.063, { font: F.semi, ls: 1.5 });
  buttonSolid(s, 11.833, 8.547, 'Read More');
  buttonGhost(s, 14.7, 8.541, 'Learn More');
}

// ================================================================ SLIDE 8 ===
// About Dribbly: copy left, magenta parallelogram + two grey cards right.
function slide08() {
  const s = newSlide();
  para(s, 15.072, 1.313, 3.928, 8.132, 0.355, C.magenta);
  sideRail(s, 'left');

  text(s, KICKER, 1.995, 1.607, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'About Dribbly', 1.995, 2.106, 8.129, 0.909, { font: F.semi, size: SZ.head });
  text(s, INTEGER, 1.995, 3.658, 5.717, 0.438, { font: F.semi, wrap: false });
  text(s, VESTIBULUM, 1.995, 4.098, 7.7, 1.063, { ls: 1.5 });

  [[2.108, 3.161], [6.779, 7.833]].forEach(pair => {
    iconTile(s, pair[0], 6.204, 0.8, 'ball');
    labelBlock(s, pair[1], 6.109, 3.267, SECTION, LOREM_TINY, { bodyH: 1.568 });
  });

  statBlock(s, 2.108, 8.337, '+26%', 'Lorem Ipsum');
  text(s, ALIQUAM, 4.391, 8.361, 6.642, 1.063, { font: F.semi, ls: 1.5 });

  [[11.291, 2.04, 11.758, 2.404], [14.336, 7.35, 14.803, 7.714]].forEach(c => {
    rect(s, c[0], c[1], 4.617, 2.35, C.gray);
    labelBlock(s, c[2], c[3], 3.683, 'Section About', LOREM_SHORT, { labelSize: SZ.lead });
  });
}

// ================================================================ SLIDE 9 ===
// Offerings: full-width magenta band with two icon blocks over it.
function slide09() {
  const s = newSlide();
  rect(s, 1.718, 5.633, 16.443, 2.411, C.magenta);
  [[2.433, 3.725, 'trophy'], [11.035, 12.327, 'ball']].forEach(b => {
    icon(s, b[2], b[0], 6.362, 1.0);
    labelBlock(s, b[1], 6.051, 5.161, SECTION, LOREM, { labelSize: SZ.lead });
  });
  text(s, 'Our Offerings Service', 1.718, 8.718, 6.179, 1.717, { font: F.semi, size: SZ.head });
  text(s, 'Pellentesque laoreet nulla eget hendrerit ultricies. Duis lobortis      tortor finibus imperdiet. Nullam eros quam, faucibus sit amet bibendum eget, ornare ac quam. In iaculis placerat blandit.',
    9.449, 8.893, 8.833, 1.568, { align: 'right', ls: 1.5 });
}

// =============================================================== SLIDE 10 ===
// Offerings grid: 2x2 icon tiles left, 2x2 magenta-titled pairs right.
const TILES_10 = [
  { x: 1.446, y: 2.125, ix: 2.696, iy: 2.569, isz: 1.0, lx: 1.882, ly: 3.891, bx: 1.782, by: 4.328, icon: 'ball' },
  { x: 1.446, y: 5.848, ix: 2.596, iy: 6.198, isz: 1.2, lx: 1.882, ly: 7.620, bx: 1.782, by: 8.057, icon: 'court' },
  { x: 5.324, y: 2.125, ix: 6.474, iy: 2.475, isz: 1.2, lx: 5.760, ly: 3.898, bx: 5.660, by: 4.335, icon: 'trophy' },
  { x: 5.287, y: 5.867, ix: 6.437, iy: 6.217, isz: 1.2, lx: 5.723, ly: 7.639, bx: 5.623, by: 8.076, icon: 'medal' },
];
function slide10() {
  const s = newSlide();
  slashPair(s, 1.354);
  sideRail(s, 'right');

  TILES_10.forEach(t => {
    rect(s, t.x, t.y, 3.5, 3.5, C.gray);
    icon(s, t.icon, t.ix, t.iy, t.isz);
    text(s, SECTION, t.lx, t.ly, 2.629, 0.438, { font: F.semi, align: 'center' });
    text(s, 'Donec purus felis ad, mauris volutpat qui.', t.bx, t.by, 2.828, 0.967,
      { size: SZ.small, align: 'center', ls: 1.5 });
  });

  text(s, 'Our Offerings Service', 9.667, 1.413, 7.301, 1.717, { font: F.semi, size: SZ.head });
  text(s, 'Sed consectetur, risus ac egestas maximus, ante nunc fringilla lacus, sed suscipit lorem felis at leo. Cras augue libero.',
    9.667, 3.596, 8.129, 1.568, { ls: 1.5 });
  [[9.667, 5.939], [14.113, 5.939], [9.667, 8.043], [14.113, 8.043]].forEach(p =>
    labelBlock(s, p[0], p[1], 3.683, SECTION, LOREM_SHORT, { labelSize: SZ.lead, labelColor: C.magenta }));
}

// =============================================================== SLIDE 11 ===
// Achievements: three text+button clusters around a dark right-hand panel.
const ACHIEVE_11 = [
  { x: 7.15, y: 0.957, align: 'left', btn: 7.25 },
  { x: 0.917, y: 4.361, align: 'right', btn: 3.833 },
  { x: 7.15, y: 7.765, align: 'left', btn: 7.25 },
];
function slide11() {
  const s = newSlide();
  scrim(s, 13.583, 0.821, 5.5, 9.608, 70);

  ACHIEVE_11.forEach(a => {
    labelBlock(s, a.x, a.y, 5.7, SECTION, 'Duis mattis ipsum eget quam placerat imperdiet eleifend. ',
      { labelSize: SZ.lead, align: a.align });
    buttonSolid(s, a.btn, a.y + 1.728, 'Read More');
  });

  text(s, 'Our Achievement', 13.851, 7.325, 4.965, 1.717, { font: F.semi, size: SZ.head });
  text(s, 'Ut leo ex, consequat a ornare a, bibendum quis erat.', 13.851, 9.041, 4.965, 0.774);
}

// =============================================================== SLIDE 12 ===
// Gallery intro: magenta hero panel plus one grey caption tile bottom-right.
function slide12() {
  const s = newSlide();
  rect(s, 1.104, 0.602, 8.8, 4.759, C.magenta);
  text(s, 'Gallery & Portofolio', 1.527, 1.362, 4.886, 1.717, { font: F.semi, size: SZ.head });
  text(s, VESTIBULUM, 1.527, 3.407, 7.955, 1.063, { ls: 1.5 });

  rect(s, 14.668, 5.625, 4.267, 4.958, C.gray);
  icon(s, 'ball', 16.301, 6.514, 1.0);
  text(s, SECTION, 15.486, 7.836, 2.629, 0.438, { font: F.semi, align: 'center' });
  text(s, 'PLACEHOLDER',
    14.912, 8.273, 3.778, 1.421, { size: SZ.small, align: 'center', ls: 1.5 });
}

// =============================================================== SLIDE 13 ===
// Gallery detail: magenta top band, icon rows right, stat strips below.
function slide13() {
  const s = newSlide();
  rect(s, 0, 0, 20, 5.625, C.magenta);
  text(s, 'Gallery & Portofolio', 11.479, 1.283, 4.886, 1.717, { font: F.semi, size: SZ.head });
  text(s, VESTIBULUM, 11.479, 3.328, 7.955, 1.063, { ls: 1.5 });

  [6.276, 8.543].forEach(y => {
    iconTile(s, 11.337, y + 0.095, 0.8, 'ball');
    labelBlock(s, 12.39, y, 3.781, SECTION, LOREM_NODOT);
  });

  rect(s, 16.58, 6.048, 2.277, 4.199, C.gray);
  statBlock(s, 16.69, 6.587, '+26%', 'Lorem Ipsum');
  statBlock(s, 16.69, 8.596, '+26%', 'Lorem Ipsum');

  rect(s, 1.104, 8.395, 9.823, 1.851, C.gray);
  statBlock(s, 1.555, 8.765, '+26%', 'Lorem Ipsum');
  text(s, ALIQUAM, 3.866, 8.789, 6.642, 1.063, { font: F.semi, ls: 1.5 });
}

// =============================================================== SLIDE 14 ===
// Skills: copy + CTAs left, a vertical column of four stats on the right.
function slide14() {
  const s = newSlide();
  sideRail(s, 'left');

  text(s, KICKER, 1.864, 1.37, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'Refine Your Skills at Our Basketball Club', 1.864, 1.808, 8.129, 1.717, { font: F.semi, size: SZ.head });
  text(s, VESTIBULUM, 1.864, 3.974, 8.129, 1.063, { ls: 1.5 });
  [1.864, 6.309].forEach(x => labelBlock(s, x, 5.728, 3.683, SECTION, LOREM_SHORT,
    { labelSize: SZ.lead, labelColor: C.magenta }));
  buttonSolid(s, 1.977, 7.997, 'Join Now');
  buttonGhost(s, 4.843, 7.991, 'Learn More');

  rect(s, 11.236, 1.385, 2.651, 8.48, C.gray);
  [2.056, 4.065, 6.074, 8.083].forEach(y => statBlock(s, 11.533, y, '+26%', 'Lorem Ipsum'));
}

// =============================================================== SLIDE 15 ===
// Split hero: magenta panel bottom-right, icon rows and stats elsewhere.
function slide15() {
  const s = newSlide();
  [1.265, 3.241].forEach(y => {
    iconTile(s, 14.229, y + 0.095, 0.8, 'ball');
    labelBlock(s, 15.283, y, 3.781, SECTION, LOREM_NODOT);
  });

  text(s, 'Donec purus felis, mauris volutpat auctor.', 0.936, 7.325, 6.175, 0.558, { font: F.semi, ls: 1.5 });
  iconTile(s, 1.047, 8.456, 0.8, 'ball');
  labelBlock(s, 2.1, 8.361, 5.01, SECTION, LOREM_NODOT);
  statBlock(s, 7.527, 7.301, '+26%', 'Lorem Ipsum');
  statBlock(s, 7.527, 8.868, '+26%', 'Lorem Ipsum');

  rect(s, 10.0, 5.625, 9.064, 4.648, C.magenta);
  text(s, 'Dribbling Towards Greatness', 10.353, 6.206, 8.152, 1.717, { font: F.semi, size: SZ.head });
  text(s, TINCIDUNT, 10.349, 8.124, 8.152, 1.568, { ls: 1.5 });
}

// =============================================================== SLIDE 16 ===
// Schedule: four date chips along the bottom, the second one highlighted.
const SCHEDULE = [
  { x: 1.159, tx: 1.345, w: 4.056, y: 8.395, date: '23 August', fill: C.gray },
  { x: 5.698, tx: 5.887, w: 4.071, y: 8.395, date: '25 August', fill: C.magenta },
  { x: 10.242, tx: 10.429, w: 4.056, y: 8.381, date: '27 August', fill: C.gray },
  { x: 14.783, tx: 14.971, w: 4.056, y: 8.395, date: '29 August', fill: C.gray },
];
function slide16() {
  const s = newSlide();
  text(s, KICKER, 1.159, 1.275, 4.322, 0.438, { font: F.semi, color: C.magenta, wrap: false });
  text(s, 'Basketball Club Schedule', 1.159, 1.713, 6.667, 1.717, { font: F.semi, size: SZ.head });
  text(s, TINCIDUNT + 'Integer sodales turpis rhoncus sollicitudin facilisis. ',
    9.352, 1.55, 9.489, 1.568, { align: 'right', ls: 1.5 });

  SCHEDULE.forEach(c => {
    rect(s, c.x, c.y, c.w, 1.365, c.fill);
    text(s, c.date, c.tx, 8.64, 3.683, 0.438, { font: F.semi, align: 'center' });
    text(s, 'Schedule Section', c.tx, 8.956, 3.683, 0.558, { align: 'center', ls: 1.5 });
  });
}

// =============================================================== SLIDE 17 ===
// Contact: slanted grey panel carrying four parallelogram icon chips.
const CONTACT = [
  { px: 6.362, py: 5.466, tx: 8.002, ty: 5.466, glyph: 'pin', label: 'Address', body: '567 AR Prawiranegara St, Metro City, Indonesia' },
  { px: 12.215, py: 5.466, tx: 13.856, ty: 5.479, glyph: 'mail', label: 'Email', body: 'dribblyclub@mail.com\nteam@contact.com' },
  { px: 5.687, py: 8.203, tx: 7.367, ty: 8.203, glyph: 'phone', label: 'Phone', body: '+12 345 678 900\n+12 098 765 432' },
  { px: 11.511, py: 8.203, tx: 13.180, ty: 8.216, glyph: 'globe', label: 'Website', body: 'www.basket-dribbly.com' },
];
function slide17() {
  const s = newSlide();
  poly(s, 0.011, 4.208, 2.295, 7.042, [[0.606, 0], [1, 0], [0.223, 1], [0, 1], [0, 0.78]], C.magenta);
  poly(s, 9.844, 0, 1.6, 2.747, [[0.435, 0], [1, 0], [0.565, 1], [0, 1]], C.magenta);

  text(s, 'Connect With Our Club', 10.39, 3.019, 6.033, 1.717, { font: F.semi, size: SZ.head });
  para(s, 4.889, 5.008, 13.734, 5.119, 0.25, C.gray);

  CONTACT.forEach(c => {
    para(s, c.px, c.py, 1.5, 1.0, 0.25, C.magenta);
    icon(s, c.glyph, c.px + 0.45, c.py + 0.2, 0.6);
    labelBlock(s, c.tx, c.ty, 3.683, c.label, c.body);
  });
}

// =============================================================== SLIDE 18 ===
// Break slide: horizontal gradient scrim behind an oversized wordmark.
// pptxgenjs has no gradient fill, so the wash is banded into 20 strips.
const BREAK_BAND = ['1D1D1D', '202021', '242425', '282828', '2A2A2B', '2C2C2D', '2D2E2E', '2F2F2F', '2F2F30',
  '2F2F30', '2E2F2F', '2D2D2E', '2B2B2C', '292A2A', '262626', '232324', '1F201F', '1A1A1A', '161616', '101010'];
function slide18() {
  const s = newSlide();
  BREAK_BAND.forEach((c, i) => rect(s, i, 0, 1.001, 11.25, c));
  text(s, KICKER, 1.083, 0.625, 4.322, 0.438, { font: F.semi, wrap: false });
  SOCIAL.forEach((g, i) => icon(s, g, 17.149 + i * 0.696, 0.742, 0.35));
  text(s, 'BREAK SLIDE', 10.883, 3.05, 8.833, 5.15, { font: F.semi, size: SZ.huge });
}

// ======================================================== SLIDES 19 - 22 ===
// S.W.O.T. quartet. Each carries a giant brush-script letter on a dark panel.
const SWOT = [
  { // 19 — Strength: numbered list left, letter panel right
    word: 'Strength', letter: 'S', panel: [11.915, 0.875, 7.117, 9.5], letterBox: [12.65, 1.788, 4.42],
    rail: 'left', textX: 2.19, textW: 8.936,
    items: [4.31, 6.223, 8.221], numX: 2.271, bodyX: 3.186, bodyW: 7.941,
  },
  { // 20 — Weakness: letter panel left, copy plus three stat chips right
    word: 'Weakness', letter: 'W', panel: [0.931, 0.875, 7.212, 9.5], letterBox: [0.74, 1.788, 5.857],
    rail: 'right', textX: 8.977, textW: 9.656, stats: [8.977, 11.915, 14.852],
    paras: [[4.078, 1.063, NUNC],
      [5.573, 1.568, 'Curabitur nec sollicitudin nunc. Fusce iaculis eros lobortis nibh pharetra feugiat. Donec quis est consequat, venenatis orci in, accumsan lacus.']],
  },
  { // 21 — Opportunities: copy + two numbered items left, letter panel right
    word: 'Opportunities', letter: 'O', panel: [11.915, 0.875, 7.117, 9.5], letterBox: [12.824, 1.788, 4.227],
    textX: 1.794, textW: 9.656, paras: [[3.971, 1.063, NUNC]],
    items: [5.889, 7.887], numX: 1.875, bodyX: 2.796, bodyW: 8.0,
  },
  { // 22 — Threats: letter panel left, two progress bars right
    word: 'Threats', letter: 'T', panel: [0.931, 0.875, 7.117, 9.5], letterBox: [1.351, 1.983, 3.78],
    textX: 8.977, textW: 9.656, paras: [[3.689, 1.063, NUNC]],
    bars: [5.602, 7.985], barX: 9.035, barW: 8.033,
  },
];
function swotSlide(cfg) {
  const s = newSlide();
  scrim(s, cfg.panel[0], cfg.panel[1], cfg.panel[2], cfg.panel[3], 50);
  text(s, cfg.letter, cfg.letterBox[0], cfg.letterBox[1], cfg.letterBox[2], 7.674,
    { font: F.brush, size: SZ.brush, color: C.magenta, wrap: false });
  if (cfg.rail) sideRail(s, cfg.rail);

  text(s, 'S.W.O.T', cfg.textX, 1.15, 1.475, 0.505, { font: F.semi, size: SZ.lead, wrap: false });
  s.addText([
    { text: cfg.word, options: { color: C.magenta, breakLine: true } },
    { text: 'Analysis Slide', options: { color: C.white } },
  ], { x: cfg.textX, y: 1.655, w: cfg.textW, h: 1.717, fontFace: F.semi, fontSize: SZ.head, valign: 'top', margin: 0.05 });

  (cfg.paras || []).forEach(p => text(s, p[2], cfg.textX, p[0], cfg.textW, p[1], { ls: 1.5 }));

  (cfg.items || []).forEach((y, i) => {
    rect(s, cfg.numX, y + 0.117, 0.6, 0.6, C.gray);
    centered(s, String(i + 1), cfg.numX, y + 0.117, 0.6, 0.6, { size: SZ.lead });
    labelBlock(s, cfg.bodyX, y, cfg.bodyW, SECTION, SWOT_BODY, { labelSize: SZ.lead });
  });

  (cfg.bars || []).forEach(y => {
    progressBar(s, cfg.barX, y, cfg.barW, 'Section progress', 80, { font: F.med });
    text(s, SWOT_BODY, cfg.barX, y + 0.788, cfg.barW, 1.063, { ls: 1.5 });
  });

  (cfg.stats || []).forEach(x => {
    rect(s, x, 7.866, 2.651, 1.275, C.gray);
    statBlock(s, x + 0.297, 7.948, '+26%', 'Lorem Ipsum');
  });
}

// ======================================================== SLIDES 23 - 26 ===
// "Unique & Creative" quartet — mirrored layouts around a player silhouette.
const CREATIVE = [
  { sil: SIL.DRIBBLE, box: [13.083, 1.582, 5.26, 8.09], rail: 'left', textX: 3.529, icons: [3.529, 8.2], footX: 14.338 },
  { sil: SIL.LAYUP, box: [1.535, 1.579, 4.4, 8.09], rail: 'right', textX: 8.148, icons: [8.148, 12.819], footX: 0.45 },
  { sil: SIL.GUARD, box: [12.929, 1.58, 5.63, 8.09], rail: 'left', textX: 3.529, icons: [3.529, 8.2], footX: 14.338 },
  { sil: SIL.DRIVE, box: [1.683, 1.58, 5.29, 8.09], rail: 'right', textX: 8.148, icons: [8.148, 12.819], footX: 0.45 },
];
function creativeSlide(cfg) {
  const s = newSlide();
  poly(s, cfg.box[0], cfg.box[1], cfg.box[2], cfg.box[3], cfg.sil, C.silhouette);
  sideRail(s, cfg.rail);

  s.addText([
    { text: 'Unique & ', options: { color: C.magenta, breakLine: true } },
    { text: 'Creative Slide', options: { color: C.white } },
  ], { x: cfg.textX, y: 2.352, w: 7.7, h: 1.717, fontFace: F.semi, fontSize: SZ.head, valign: 'top', margin: 0.05 });
  text(s, INTEGER, cfg.textX, 4.583, 5.717, 0.438, { font: F.semi, wrap: false });
  text(s, VESTIBULUM, cfg.textX, 5.024, 7.7, 1.063, { ls: 1.5 });

  cfg.icons.forEach(x => {
    iconTile(s, x, 6.919, 0.8, 'ball');
    labelBlock(s, x + 1.053, 6.823, 3.267, SECTION, LOREM_TINY, { bodyH: 1.568 });
  });

  text(s, KICKER + ' Template', cfg.footX, 10.295, 5.212, 0.404, { font: F.med, size: SZ.small, wrap: false });
}

// =============================================================== SLIDE 27 ===
// App promo: phone mock-up right over the slash motif, copy + CTAs left.
function slide27() {
  const s = newSlide();
  slashPair(s, 11.449);
  // Phone: hollow bezel over the slash motif, with a notch bar at the top.
  deviceBezel(s, 12.57, 1.02, 4.71, 9.11, 0.30, 0.52);
  s.addShape('roundRect', { x: 14.22, y: 1.32, w: 1.42, h: 0.28, rectRadius: 0.4,
    fill: { color: '060606' }, line: { type: 'none' } });
  logoBadge(s, 0, 0);

  s.addText([
    { text: 'Watch ', options: { color: C.white } },
    { text: 'Basketball Matches ', options: { color: C.magenta } },
    { text: 'on Your Smartphone', options: { color: C.white } },
  ], { x: 1.696, y: 1.962, w: 9.469, h: 1.717, fontFace: F.semi, fontSize: SZ.head, valign: 'top', margin: 0.05 });
  text(s, 'Cras eu sem vitae massa sollicitudin commodo id ac lectus. In vehicula cursus elit, facilisis commodo ligula suscipit quis enim. Pellentesque sed sem at massa faucibus tristique sit amet quis.',
    1.696, 4.127, 9.469, 1.568, { ls: 1.5 });
  statBlock(s, 1.76, 6.104, '+26%', 'Lorem Ipsum');
  text(s, LOREM_SHORT, 4.971, 6.128, 4.179, 1.063, { font: F.semi, ls: 1.5 });
  buttonSolid(s, 1.76, 7.844, 'Download');
  buttonGhost(s, 4.626, 7.838, 'Live Demo');
}

// =============================================================== SLIDE 28 ===
// Laptop mock-up left over a big magenta slab, copy + CTAs right.
function slide28() {
  const s = newSlide();
  // Magenta slab: a diagonal cut running from the top edge down to the left.
  poly(s, 0, 0, 7.94, 11.25, [[0, 0], [1, 0], [0.5155, 1], [0, 1]], C.magenta);
  // Laptop: hollow bezel so the slab shows through the screen, plus a base bar.
  deviceBezel(s, 1.517, 2.700, 7.866, 5.040, 0.19, 0.04);
  poly(s, 0.617, 7.833, 9.650, 0.267, [[0, 0], [1, 0], [0.981, 1], [0.019, 1]], 'DFE1E4');
  rect(s, 4.60, 7.833, 1.6, 0.09, 'B6B8BC');
  logoBadge(s, 18.7, 0);

  text(s, 'Check Your Favorite Team\u2019s Match Schedule', 11.355, 1.476, 6.813, 2.524, { font: F.semi, size: SZ.head });
  text(s, 'Duis efficitur id enim non tincidunt. Donec molestie leo quam, ac vestibulum massa sollicitudin eget.',
    11.355, 4.48, 7.345, 1.063, { ls: 1.5 });
  iconTile(s, 11.435, 6.118, 1.0, 'ball');
  labelBlock(s, 12.663, 6.023, 5.161, SECTION, LOREM);
  buttonSolid(s, 11.435, 8.168, 'Read More');
  buttonGhost(s, 14.301, 8.162, 'Learn More');
}

// =============================================================== SLIDE 29 ===
// Chevron process timeline — three runs of arrow-heads that grow in size.
// Each run carries its own left-to-right gradient, mirroring the source fills.
const CHEV_RUNS = [
  { x: 1.259, y: 6.159, w: 0.435, h: 0.765, gap: 0.4375, n: 10, stops: [[0, 'F0AFFF'], [1.0, 'E989FF']] },
  { x: 6.117, y: 5.909, w: 0.722, h: 1.272, gap: 0.7265, n: 8, stops: [[0, 'E989FF'], [0.81, 'DA3BFF']] },
  { x: 12.175, y: 5.550, w: 1.133, h: 1.994, gap: 1.1465, n: 6, stops: [[0, 'DA3BFF'], [0.14, 'CC00FF']] },
];
const YEAR_LABELS = [['2021', 2.609, 8.355, 1.65], ['2022', 8.108, 3.686, 1.803], ['2023', 14.668, 8.355, 1.813]];
const CHEV_CAPTIONS = [{ x: 1.025, y: 3.213 }, { x: 6.6, y: 7.882 }, { x: 13.166, y: 3.213 }];
function slide29() {
  const s = newSlide();
  text(s, 'Chevron Process Diagram for PowerPoint', 4.44, 0.708, 11.121, 1.212,
    { font: F.semi, size: SZ.head, align: 'center' });

  CHEV_RUNS.forEach(run => {
    const span = (run.n - 1) * run.gap + run.w;
    for (let i = 0; i < run.n; i++) {
      const centre = (i * run.gap + run.w / 2) / span;
      poly(s, run.x + i * run.gap, run.y, run.w, run.h, CHEVRON, gradientAt(run.stops, centre));
    }
  });

  YEAR_LABELS.forEach(y => text(s, y[0], y[1], y[2], y[3], 0.909,
    { font: F.med, size: SZ.head, bold: true, align: 'center', wrap: false }));

  CHEV_CAPTIONS.forEach(c => {
    text(s, SECTION, c.x, c.y, 4.818, 0.438, { font: F.med, bold: true, align: 'center', margin: 0 });
    text(s, 'Lorem ipsum dolor sit amet, nibh est. A magna maecenas, quam magna nec quis, lorem nunc.',
      c.x + 0.013, c.y + 0.434, 4.805, 1.421, { size: SZ.small, align: 'center', ls: 1.5, margin: 0 });
  });
}

// =============================================================== SLIDE 30 ===
// Vertical chevron list: four interleaved chevrons with icons and captions.
const VCHEV = [
  { cx: 8.471, cy: 2.232, flip: false, icon: 'rocket', ix: 7.115, iy: 2.988, tx: 2.667, ty: 2.927, align: 'right' },
  { cx: 10.0, cy: 3.913, flip: true, icon: 'people', ix: 11.632, iy: 4.670, tx: 13.981, ty: 4.608, align: 'left' },
  { cx: 8.471, cy: 5.595, flip: false, icon: 'bulb', ix: 7.119, iy: 6.351, tx: 2.743, ty: 6.290, align: 'right' },
  { cx: 10.0, cy: 7.276, flip: true, icon: 'puzzle', ix: 11.632, iy: 8.033, tx: 13.981, ty: 7.971, align: 'left' },
];
function slide30() {
  const s = newSlide();
  text(s, 'Vertical Chevron List for PowerPoint', 1.375, 0.768, 17.25, 1.212,
    { font: F.semi, size: SZ.head, align: 'center' });

  VCHEV.forEach(v => {
    // Alternate rows point left, mirroring the reference's 180-degree rotation.
    const pts = v.flip ? [[1, 0], [0.334, 0], [0, 0.5], [0.334, 1], [1, 1], [0.666, 0.5]]
      : [[0, 0], [0.666, 0], [1, 0.5], [0.666, 1], [0, 1], [0.334, 0.5]];
    poly(s, v.cx, v.cy, 1.81, 3.013, pts, C.magenta);
    icon(s, v.icon, v.ix, v.iy, 1.5);
    text(s, SECTION, v.tx, v.ty, 3.683, 0.438, { font: F.semi, align: v.align });
    text(s, LOREM_LONG, v.tx, v.ty + 0.512, 3.683, 1.111, { align: v.align });
  });
}

// =============================================================== SLIDE 31 ===
// Process & ribbons: grey arrow behind four slanted magenta ribbons.
const RIBBON_X = [1.463, 4.350, 7.240, 10.134];
const RIBBON_TAB_R = [[0.705, 0.635], [1.0, 0.635], [1.0, 1.0], [0.705, 1.0]];
const RIBBON_TAB_L = [[0.0, 0.0], [0.140, 0.0], [0.140, 0.234], [0.0, 0.234]];
const RIBBON_BODY = [[0.330, 0.0], [1.0, 0.811], [0.670, 1.0], [0.0, 0.189]];
const RINGS = [[15.332, 3.161, 4.199], [15.974, 3.802, 2.916], [16.724, 4.553, 1.415]];
function slide31() {
  const s = newSlide();
  text(s, 'Process and Ribbons Diagram for PowerPoint', 3.929, 0.615, 12.142, 1.212,
    { font: F.semi, size: SZ.head, align: 'center' });

  RINGS.forEach(r => s.addShape('ellipse', { x: r[0], y: r[1], w: r[2], h: r[2],
    fill: { color: '000000', transparency: 95 }, line: { type: 'none' } }));

  RIBBON_X.forEach(x => {
    poly(s, x, 4.063, 4.311, 2.395, RIBBON_TAB_R, C.ribbonDk);
    poly(s, x, 4.063, 4.311, 2.395, RIBBON_TAB_L, C.ribbonDk);
  });
  s.addShape('rightArrow', { x: 0, y: 3.69, w: 17.432, h: 3.141, fill: { color: C.arrow }, line: { type: 'none' } });
  RIBBON_X.forEach(x => poly(s, x, 3.338, 4.311, 3.845, RIBBON_BODY, C.magenta));

  [1.331, 5.864, 10.398, 14.931].forEach(x => {
    text(s, SECTION, x, 8.405, 3.683, 0.438, { font: F.semi });
    text(s, LOREM_LONG, x, 8.918, 3.683, 1.111);
  });
}

// =============================================================== SLIDE 32 ===
// Thanks (slash variant): slash motif right, wordmark centre-left.
function slide32() {
  const s = newSlide();
  slashPair(s, 11.449);
  sideRail(s, 'left');
  text(s, 'THANKS', 1.77, 4.312, 8.849, 2.625, { font: F.semi, size: SZ.huge, align: 'center' });
  text(s, KICKER + ' Template', 3.317, 6.458, 5.756, 0.438, { font: F.semi, wrap: false });
  text(s, KICKER, 17.187, 8.311, 4.322, 0.438, { font: F.semi, rotate: 270, wrap: false });
}

// =============================================================== SLIDE 33 ===
// Thanks (dark variant): full-bleed scrim, centred wordmark, rails both sides.
function slide33() {
  const s = newSlide();
  scrim(s, 0, 0, 20, 11.25, 70);
  sideRail(s, 'left');
  s.addShape('line', { x: 19.369, y: 5.245, w: 0, h: 5.374, line: { color: C.white, width: 3 } });
  text(s, KICKER, 17.187, 2.573, 4.322, 0.438, { font: F.semi, rotate: 270, wrap: false });
  text(s, 'THANKS', 5.575, 4.312, 8.849, 2.625, { font: F.semi, size: SZ.huge, align: 'center' });
  text(s, KICKER, 7.839, 6.438, 4.322, 0.438, { font: F.semi, wrap: false });
}

// ------------------------------------------------------------------ build --
slide01(); slide02(); slide03(); slide04(); slide05(); slide06();
slide07(); slide08(); slide09(); slide10(); slide11(); slide12();
slide13(); slide14(); slide15(); slide16(); slide17(); slide18();
SWOT.forEach(swotSlide);          // 19 - 22
CREATIVE.forEach(creativeSlide);  // 23 - 26
slide27(); slide28(); slide29(); slide30(); slide31(); slide32(); slide33();

pptx.writeFile({ fileName: path.join(__dirname, '052e1e67-3db0-4c61-bb98-db10cbd41f25_grok_final.pptx') })
  .then(f => console.log('wrote', f));
