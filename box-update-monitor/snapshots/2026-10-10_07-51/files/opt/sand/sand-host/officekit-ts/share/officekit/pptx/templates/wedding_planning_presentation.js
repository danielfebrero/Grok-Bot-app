/*
 * Eternal Vows - "Wedding Plan Made Elegant" deck (20 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs. Raster photos in the original are replaced by
 * light-grey "[image]" placeholders that keep the original silhouettes.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const DARK1 = '3E0710'; // theme accent1 - deepest wine (title backgrounds)
const DARK2 = '5D0B18'; // theme accent2
const WINE  = '7C1020'; // theme accent3
const ROSE  = 'B7494C'; // theme accent4
const BLUSH = 'FFE4E1'; // theme accent5
const WHITE = 'FFFFFF';
const TX85  = '262626'; // tx1 lumMod 85%
const TX75  = '404040'; // tx1 lumMod 75%
const TX95  = '0D0D0D'; // tx1 lumMod 95%
const GRAY95 = 'F2F2F2'; // bg1 lumMod 95%
const GRAY75 = 'BFBFBF'; // bg1 lumMod 75%
const BLUSH75 = 'FF7869'; // accent5 lumMod 75%
const BLUSH90 = 'FFB9B1'; // accent5 lumMod 90%
const IMG = 'F2F2F2';     // photo placeholder fill
const IMG_TXT = 'C9C9C9'; // photo placeholder caption

/* ------------------------------------------------------------------- fonts */
const HEAD = 'Domine';      // theme major latin
const BODY = 'Outfit';      // theme minor latin
const SCRIPT = 'Parisienne';

/* PowerPoint default text-box insets, as [left, right, bottom, top] in points */
const INS = [7.2, 7.2, 3.6, 3.6];

/* the two shadow presets used throughout the original deck */
const SH_DOT  = { type: 'outer', color: '000000', opacity: 0.05, blur: 30, offset: 10, angle: 45 };
const SH_CARD = { type: 'outer', color: '000000', opacity: 0.05, blur: 41, offset: 10, angle: 100 };

/* ------------------------------------------------------------- geometry kit */
const K = 0.5523; // circular-arc -> cubic Bezier constant

/* Points for a rectangle with an individual radius per corner (inches). */
function rounded(w, h, r) {
  const [a, b, c, d] = (Array.isArray(r) ? r : [r, r, r, r]).map(v => v || 0);
  const p = [];
  p.push({ x: a, y: 0 });
  p.push({ x: w - b, y: 0 });
  if (b) p.push({ x: w, y: b, curve: { type: 'cubic', x1: w - b * (1 - K), y1: 0, x2: w, y2: b * (1 - K) } });
  p.push({ x: w, y: h - c });
  if (c) p.push({ x: w - c, y: h, curve: { type: 'cubic', x1: w, y1: h - c * (1 - K), x2: w - c * (1 - K), y2: h } });
  p.push({ x: d, y: h });
  if (d) p.push({ x: 0, y: h - d, curve: { type: 'cubic', x1: d * (1 - K), y1: h, x2: 0, y2: h - d * (1 - K) } });
  p.push({ x: 0, y: a });
  if (a) p.push({ x: a, y: 0, curve: { type: 'cubic', x1: 0, y1: a * (1 - K), x2: a * (1 - K), y2: 0 } });
  p.push({ close: true });
  return p;
}

/* Points from a normalised outline: [x,y] = corner, [x,y,x1,y1,x2,y2] = cubic. */
function outline(w, h, segs) {
  return segs.map(s => s.length === 2
    ? { x: s[0] * w, y: s[1] * h }
    : { x: s[0] * w, y: s[1] * h, curve: { type: 'cubic', x1: s[2] * w, y1: s[3] * h, x2: s[4] * w, y2: s[5] * h } }
  ).concat([{ close: true }]);
}

/* Irregular card / photo silhouettes lifted from the original custom geometry. */
const SHAPES = {
  tiltCard:  [[0.002, 1], [1, 0.871], [0.921, 0], [0.116, 0], [0, 0.139, 0.052, 0, 0, 0.062], [0, 0.982]],
  tiltPhoto: [[0.004, 0], [1, 0], [1, 0.888], [0.173, 0.999, 0.124, 1.006, 0.08, 0.962], [0.001, 0.044, 0.075, 0.901, 0, 0.014]],
  petal:     [[0, 0], [1, 0], [0.991, 0.089], [0.5, 1, 0.926, 0.617, 0.731, 1], [0.009, 0.089, 0.269, 1, 0.074, 0.617]],
  quarter:   [[0.561, 0], [0.993, 0.269, 0.735, 0, 0.89, 0.105], [1, 0.281], [1, 1], [0.037, 1], [0.025, 0.958],
              [0, 0.738, 0.009, 0.889, 0, 0.815], [0.561, 0, 0, 0.331, 0.251, 0]],
  leanCard:  [[0.139, 0.003], [0, 0.145, 0.06, 0.016, 0, 0.075], [0, 0.855], [0.175, 1, 0, 0.935, 0.078, 1],
              [1, 1], [0.843, 0], [0.175, 0], [0.139, 0.003, 0.163, 0, 0.151, 0.001]],
  leanPhoto: [[0.761, 0], [0.879, 0.087, 0.821, 0.004, 0.871, 0.039], [0.999, 0.798],
              [0.887, 0.91, 1.008, 0.853, 0.958, 0.903], [0, 1], [0, 0.075], [0.734, 0.001], [0.761, 0, 0.743, 0, 0.752, 0]],
  scoopTR:   [[0.094, 0], [1, 0], [1, 0.862], [0.967, 0.888], [0.62, 1, 0.868, 0.959, 0.749, 1],
              [0, 0.345, 0.278, 1, 0, 0.707], [0.075, 0.033, 0, 0.232, 0.027, 0.126]],
  scoopBL:   [[0.41, 0], [1, 0.732, 0.736, 0, 1, 0.328], [0.973, 0.95, 1, 0.808, 0.991, 0.881], [0.959, 1],
              [0, 1], [0, 0.207], [0.035, 0.167], [0.41, 0, 0.137, 0.063, 0.268, 0]],
  scoopBL2:  [[0.436, 0], [1, 0.711, 0.747, 0, 1, 0.318], [0.956, 0.988, 1, 0.809, 0.984, 0.903], [0.951, 1],
              [0, 1], [0, 0.26], [0.001, 0.259], [0.436, 0, 0.104, 0.101, 0.261, 0]],
  scoopTR2:  [[0.113, 0], [1, 0], [1, 0.831], [0.963, 0.865], [0.127, 0.758, 0.704, 1.081, 0.33, 1.033],
              [0.105, 0.011, -0.038, 0.536, -0.039, 0.232]],
};

/* ------------------------------------------------------------ tiny wrappers */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'Eternal Vows';
pptx.title = 'Wedding Plan Made Elegant';

/* A text box that behaves like PowerPoint's auto-fit boxes: top anchored,
   default insets, no bullet. `o` may add align / lineSpacing / etc. */
function tx(s, body, x, y, w, h, o = {}) {
  s.addText(body, Object.assign({
    x, y, w, h, margin: INS, valign: 'top', fontFace: BODY,
    fontSize: 12, color: TX75,
  }, o));
}

/* Solid rectangle / rounded rectangle / ellipse shorthands. */
function rect(s, x, y, w, h, fill, o = {}) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color: fill } }, o));
}
function roundRect(s, x, y, w, h, fill, radius, o = {}) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, fill: { color: fill }, rectRadius: radius }, o));
}
function oval(s, x, y, d, fill, o = {}) {
  s.addShape('ellipse', Object.assign({ x, y, w: d, h: d, fill: { color: fill } }, o));
}
/* Free shape from a normalised outline of SHAPES. */
function blob(s, key, x, y, w, h, fill, o = {}) {
  s.addShape('custGeom', Object.assign({ x, y, w, h, fill: { color: fill }, points: outline(w, h, SHAPES[key]) }, o));
}
/* Rectangle with per-corner radii. */
function card(s, x, y, w, h, fill, r, o = {}) {
  s.addShape('custGeom', Object.assign({ x, y, w, h, fill: { color: fill }, points: rounded(w, h, r) }, o));
}

/* ------------------------------------------------------------- photo stand-in
   Every raster photo of the source deck becomes a flat grey silhouette with a
   small "[image]" caption, keeping the original position, size and corners. */
function photo(s, x, y, w, h, r) {
  card(s, x, y, w, h, IMG, r);
  const ch = Math.min(0.32, h);
  s.addText('[image]', {
    x, y: y + (h - ch) / 2, w, h: ch, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 11, color: IMG_TXT, margin: 0,
  });
}
function photoBlob(s, key, x, y, w, h) {
  blob(s, key, x, y, w, h, IMG);
  s.addText('[image]', {
    x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 11, color: IMG_TXT, margin: 0,
  });
}

/* ------------------------------------------------------- repeated furniture */
/* Small heart mark + "Eternal Vows" wordmark that the master paints top-left. */
function brandMark(s, color) {
  s.addShape('heart', { x: 0.635, y: 0.44, w: 0.2, h: 0.185, fill: { color } });
  tx(s, 'Eternal Vows', 0.906, 0.39, 1.597, 0.269, { fontFace: HEAD, fontSize: 10, color });
}
/* "Page N" marker top-right (master placeholder). */
function pageMark(s, n, base, accent) {
  tx(s, [{ text: 'Page ', options: { color: base } }, { text: String(n), options: { color: accent } }],
    11.854, 0.39, 0.939, 0.269, { align: 'right', fontFace: HEAD, fontSize: 10 });
}
function header(s, n, opts = {}) {
  const light = !!opts.light;
  brandMark(s, light ? WHITE : WINE);
  pageMark(s, n, light ? WHITE : TX85, light ? WHITE : ROSE);
}

/* ---------------------------------------------------------------- pictograms
   The source deck uses flat one-colour icon art. Each icon here is rebuilt
   from a handful of native shapes at unit scale: draw(slide, x, y, size, color)
   where the icon is centred inside a size x size box. */
const ICONS = {
  /* two interlocking wedding bands */
  rings(s, x, y, u, c) {
    const r = u * 0.34, t = Math.max(0.012, u * 0.075);
    s.addShape('donut', { x: x + u * 0.04, y: y + u * 0.28, w: r, h: r, fill: { color: c }, rectRadius: t });
    s.addShape('donut', { x: x + u * 0.62 - r * 0.62, y: y + u * 0.28, w: r, h: r, fill: { color: c } });
    [0.30, 0.42, 0.54].forEach((fx, i) => s.addShape('rect', {
      x: x + u * fx, y: y + u * (0.04 + 0.03 * Math.abs(i - 1)), w: t * 0.7, h: u * 0.16, fill: { color: c },
      rotate: (i - 1) * 22,
    }));
  },
  /* padlock shaped like a heart */
  lock(s, x, y, u, c, bg) {
    s.addShape('heart', { x: x + u * 0.05, y: y + u * 0.12, w: u * 0.9, h: u * 0.86, fill: { color: c } });
    s.addShape('ellipse', { x: x + u * 0.4, y: y + u * 0.4, w: u * 0.2, h: u * 0.2, fill: { color: bg } });
    s.addShape('rect', { x: x + u * 0.45, y: y + u * 0.5, w: u * 0.1, h: u * 0.2, fill: { color: bg } });
  },
  /* picture frame with swap arrows */
  swap(s, x, y, u, c) {
    s.addShape('round2DiagRect', {
      x: x + u * 0.06, y: y + u * 0.14, w: u * 0.88, h: u * 0.72,
      fill: { type: 'none' }, line: { color: c, width: Math.max(1, u * 12) }, rectRadius: u * 0.18,
    });
    s.addShape('heart', { x: x + u * 0.34, y: y + u * 0.38, w: u * 0.32, h: u * 0.3, fill: { color: c } });
  },
  /* chapel with a cross */
  church(s, x, y, u, c) {
    s.addShape('rect', { x: x + u * 0.44, y: y + u * 0.02, w: u * 0.12, h: u * 0.16, fill: { color: c } });
    s.addShape('rect', { x: x + u * 0.38, y: y + u * 0.07, w: u * 0.24, h: u * 0.06, fill: { color: c } });
    s.addShape('triangle', { x: x + u * 0.3, y: y + u * 0.16, w: u * 0.4, h: u * 0.3, fill: { color: c } });
    s.addShape('rect', { x: x + u * 0.3, y: y + u * 0.42, w: u * 0.4, h: u * 0.5, fill: { color: c } });
    s.addShape('triangle', { x: x + u * 0.02, y: y + u * 0.46, w: u * 0.3, h: u * 0.2, fill: { color: c } });
    s.addShape('rect', { x: x + u * 0.02, y: y + u * 0.62, w: u * 0.3, h: u * 0.3, fill: { color: c } });
    s.addShape('triangle', { x: x + u * 0.68, y: y + u * 0.46, w: u * 0.3, h: u * 0.2, fill: { color: c } });
    s.addShape('rect', { x: x + u * 0.68, y: y + u * 0.62, w: u * 0.3, h: u * 0.3, fill: { color: c } });
  },
  /* tiered wedding cake */
  cake(s, x, y, u, c) {
    [0.06, 0.2, 0.34].forEach((fx, i) => s.addShape('rect', {
      x: x + u * (0.36 + i * 0.14), y: y + u * 0.04, w: u * 0.03, h: u * 0.12, fill: { color: c },
    }));
    [0.36, 0.5, 0.64].forEach(fx => s.addShape('ellipse', {
      x: x + u * (fx - 0.035), y: y + u * 0.0, w: u * 0.07, h: u * 0.07, fill: { color: c },
    }));
    s.addShape('roundRect', { x: x + u * 0.3, y: y + u * 0.18, w: u * 0.4, h: u * 0.2, fill: { color: c }, rectRadius: u * 0.05 });
    s.addShape('roundRect', { x: x + u * 0.18, y: y + u * 0.4, w: u * 0.64, h: u * 0.22, fill: { color: c }, rectRadius: u * 0.05 });
    s.addShape('roundRect', { x: x + u * 0.04, y: y + u * 0.64, w: u * 0.92, h: u * 0.28, fill: { color: c }, rectRadius: u * 0.06 });
  },
  /* two champagne flutes toasting */
  toast(s, x, y, u, c) {
    [[0.05, -18], [0.55, 18]].forEach(([fx, rot]) => {
      s.addShape('ellipse', { x: x + u * fx, y: y + u * 0.2, w: u * 0.4, h: u * 0.44, fill: { color: c }, rotate: rot });
      s.addShape('rect', { x: x + u * (fx + 0.17), y: y + u * 0.58, w: u * 0.06, h: u * 0.22, fill: { color: c }, rotate: rot });
      s.addShape('rect', { x: x + u * (fx + 0.06), y: y + u * 0.8, w: u * 0.28, h: u * 0.06, fill: { color: c }, rotate: rot });
    });
    [[0.22, -20], [0.46, 0], [0.7, 20]].forEach(([fx, rot]) => s.addShape('rect', {
      x: x + u * fx, y: y + u * 0.02, w: u * 0.05, h: u * 0.14, fill: { color: c }, rotate: rot,
    }));
  },
  /* the proposal - one figure kneeling before another */
  propose(s, x, y, u, c) {
    s.addShape('ellipse', { x: x + u * 0.1, y: y + u * 0.06, w: u * 0.22, h: u * 0.22, fill: { color: c } });
    s.addShape('roundRect', { x: x + u * 0.08, y: y + u * 0.32, w: u * 0.26, h: u * 0.6, fill: { color: c }, rectRadius: u * 0.1 });
    s.addShape('ellipse', { x: x + u * 0.62, y: y + u * 0.2, w: u * 0.2, h: u * 0.2, fill: { color: c } });
    s.addShape('roundRect', { x: x + u * 0.6, y: y + u * 0.44, w: u * 0.24, h: u * 0.3, fill: { color: c }, rectRadius: u * 0.08 });
    s.addShape('roundRect', { x: x + u * 0.52, y: y + u * 0.74, w: u * 0.44, h: u * 0.16, fill: { color: c }, rectRadius: u * 0.06 });
    s.addShape('rect', { x: x + u * 0.32, y: y + u * 0.44, w: u * 0.3, h: u * 0.07, fill: { color: c }, rotate: -8 });
  },
};

/* A pictogram inside a filled circle - the deck's most repeated motif. */
function iconDot(s, x, y, d, circle, icon, glyphColor = WHITE) {
  oval(s, x, y, d, circle, { shadow: SH_DOT });
  const u = d * 0.44;
  ICONS[icon](s, x + (d - u) / 2, y + (d - u) / 2, u, glyphColor, circle);
}
/* Evenly spaced row of icon dots: list entries are [circleColor, iconName]. */
function iconRow(s, x, y, d, gap, list) {
  list.forEach((it, i) => iconDot(s, x + i * (d + gap), y, d, it[0], it[1]));
}

/* Small eyebrow label used above nearly every headline. */
function eyebrow(s, text, x, y, color = ROSE, w = 2.661) {
  tx(s, text, x, y, w, 0.286, { fontFace: HEAD, fontSize: 11, color });
}
/* Multi-line headline; every array entry is its own paragraph. */
function headline(s, lines, x, y, w, h, size, color, o = {}) {
  tx(s, lines.map((t, i) => ({ text: t, options: i < lines.length - 1 ? { breakLine: true } : {} })),
    x, y, w, h, Object.assign({ fontFace: HEAD, fontSize: size, color, lineSpacingMultiple: 1.3 }, o));
}
/* Body paragraph block, one entry per hard line break. */
function para(s, lines, x, y, w, h, o = {}) {
  tx(s, lines.map((t, i) => ({ text: t, options: i < lines.length - 1 ? { breakLine: true } : {} })),
    x, y, w, h, Object.assign({ lineSpacingMultiple: 1.3 }, o));
}
/* "First Love / A wedding plan makes dreams reality." caption pairs. */
function captionPair(s, title, lines, x, y, o = {}) {
  tx(s, title, x, y, 1.763, 0.38, Object.assign({ fontFace: HEAD, fontSize: 14, color: ROSE, lineSpacingMultiple: 1.3 }, o));
  para(s, lines, x, y + 0.447, o.w || 2.654, 0.602, Object.assign({ color: TX85 }, o));
}

/* ---------------- slides ------------------------------------------------- */

/* 1 & 20 - full-bleed wine covers with a script/serif mixed headline. */
function coverSlide(s, lines, subtitle, big, script) {
  rect(s, 0, 0, 13.333, 7.5, DARK1);
  photo(s, 0, 0, 13.333, 7.5, 0);
  rect(s, 0, 0, 13.333, 7.5, DARK2, { fill: { color: DARK2, transparency: 10 } });
  brandMark(s, WHITE);
  tx(s, [{ text: 'Presentation Template ' }, { text: '2030', options: { bold: true } }],
    9.193, 0.416, 3.601, 0.236, { align: 'right', fontFace: HEAD, fontSize: 8, color: WHITE });
  s.addText(lines.map(ln => ln.runs.map((r, i) => ({
    text: r[0],
    options: Object.assign(
      { fontFace: r[1] ? SCRIPT : HEAD, fontSize: r[1] ? script : big, color: WHITE },
      i === ln.runs.length - 1 ? { breakLine: true, align: ln.align } : { align: ln.align }),
  }))).flat(), {
    x: lines.x, y: lines.y, w: lines.w, h: lines.h, margin: INS, valign: 'top', lineSpacing: 115,
  });
  para(s, subtitle, 1.422, lines.subY, 4.59, 0.602, { color: BLUSH90 });
  iconRow(s, 10.449, 5.841, 0.409, 0.165, [[ROSE, 'rings'], [ROSE, 'lock'], [ROSE, 'swap']]);
}

function slide1(s) {
  const L = [{ align: 'left', runs: [['W', 1], ['edding Plan', 0]] },
             { align: 'right', runs: [['Made Elegan', 0], ['t', 1]] }];
  L.x = 1.376; L.y = 2.498; L.w = 10.582; L.h = 3.627; L.subY = 4.659;
  coverSlide(s, L, ['Dreams Become Weddings ', 'Filled with Love.'], 80, 138);
}

function slide20(s) {
  const L = [{ align: 'left', runs: [['T', 1], ['hanks', 0]] },
             { align: 'right', runs: [['for ', 0], ['L', 1], ['ove', 0]] }];
  L.x = 1.175; L.y = 2.752; L.w = 10.984; L.h = 3.761; L.subY = 4.833;
  coverSlide(s, L, ['A Wedding Plan Turns Dreams ', 'Into Cherished Memories.'], 96, 199);
}

/* 2 - "Choosing the Ideal Wedding Venue" with two portrait cards. */
function slide2(s) {
  header(s, 2);
  [[9.175, 0.938, 'Anna Grace', 9.495, 1.227], [5.581, 2.423, 'Lily Rose', 5.901, 2.712]].forEach(
    ([cx, cy, name, px, py]) => {
      roundRect(s, cx, cy, 3.274, 4.24, WHITE, 0.38, { shadow: SH_CARD });
      photo(s, px, py, 2.634, 2.679, 0.21);
      tx(s, name, cx + 0.394, cy + 3.254, 2.486, 0.37, { align: 'center', fontFace: HEAD, fontSize: 16, color: ROSE });
      tx(s, 'The Wedding Plan', cx + 0.317, cy + 3.606, 2.64, 0.341,
        { align: 'center', color: GRAY75, lineSpacingMultiple: 1.3 });
    });
  eyebrow(s, 'The Wedding Planner', 0.541, 1.398);
  headline(s, ['Choosing the', 'Ideal Wedding', 'Venue'], 0.533, 1.762, 5.648, 2.494, 44, TX75, { lineSpacing: 58 });
  tx(s, 'Wedding Plan is A Journey Their Special Day', 0.538, 5.294, 4.84, 0.34,
    { fontFace: HEAD, color: ROSE, lineSpacingMultiple: 1.3 });
  para(s, ['The Wedding plan is a thoughtful journey, guiding ',
           'couples through love, and ensuring their special day.'], 0.538, 5.784, 4.864, 0.602);
  iconDot(s, 9.569, 5.757, 0.619, WINE, 'lock');
  tx(s, '$12,580', 10.45, 5.747, 3.584, 0.64, { fontFace: HEAD, fontSize: 32, color: ROSE });
}

/* 3 - "Defining Your Dream Wedding Vision" + three date chips. */
function slide3(s) {
  header(s, 3);
  card(s, 10.554, 4.809, 2.779, 2.691, WHITE, [0.41, 0, 0, 0], { shadow: SH_CARD });
  photo(s, 10.756, 5.011, 2.577, 2.489, [0.28, 0.03, 0, 0]);
  card(s, 5.358, 0, 3.161, 2.923, WHITE, [0, 0, 0.41, 0.41], { shadow: SH_CARD });
  photo(s, 5.557, 0, 2.757, 2.726, [0, 0, 0.28, 0.28]);
  roundRect(s, 7.908, 2.351, 3.161, 3.161, WHITE, 0.41, { shadow: SH_CARD });
  photo(s, 8.11, 2.553, 2.757, 2.757, 0.28);
  eyebrow(s, 'The Wedding Planner', 0.541, 2.41);
  headline(s, ['Defining Your ', 'Dream Wedding Vision'], 0.533, 2.774, 7.009, 1.44, 32, TX85);
  tx(s, 'A wedding plan turns love into lasting memories.', 0.541, 4.455, 4.923, 0.303);
  [[0.646, DARK2, '25'], [1.769, WINE, '04'], [2.892, ROSE, '30']].forEach(([x, c, n]) => {
    roundRect(s, x, 5.61, 0.905, 0.92, c, 0.11);
    tx(s, n, x + 0.068, 5.851, 0.771, 0.438, { align: 'center', fontFace: HEAD, fontSize: 20, color: WHITE });
  });
  para(s, ['The wedding plan organizes every detail ', 'carefully, blending love, vision, creativity, ',
           'and celebration.'], 4.066, 5.638, 4.095, 0.865, { color: TX85 });
  captionPair(s, 'First Love', ['A wedding plan makes ', 'dreams reality.'], 8.853, 0.823);
  captionPair(s, 'True Love', ['Weddings begin ', 'with love.'], 11.406, 3.407, { w: 1.763 });
  iconDot(s, 5.133, 1.017, 0.661, WINE, 'lock');
  iconDot(s, 7.678, 3.601, 0.661, ROSE, 'lock');
}

/* 4 - "Understanding the Essence of Wedding Planning" (photo pair left). */
function slide4(s) {
  header(s, 4);
  card(s, 0, 1.059, 5.778, 5.664, WHITE, [0, 0.46, 0.46, 0], { shadow: SH_CARD });
  photo(s, 0, 1.246, 2.128, 5.289, [0, 0.35, 0.35, 0]);
  photo(s, 2.128, 1.246, 3.455, 5.289, [0, 0.35, 0.35, 0]);
  eyebrow(s, 'The Wedding Planner', 6.332, 1.565, ROSE, 3.208);
  headline(s, ['Understanding the Essence ', 'of Wedding Planning'], 6.324, 1.929, 7.009, 1.272, 28, TX85);
  iconDot(s, 6.428, 4.499, 0.475, WINE, 'rings');
  iconDot(s, 7.095, 4.499, 0.475, ROSE, 'lock');
  tx(s, 'Timeless Moments', 7.762, 4.535, 5.084, 0.404, { fontSize: 18, color: TX85 });
  para(s, ['A wedding plan is a thoughtful journey, guiding couples through love, ',
           'vision, organization, and creativity, ensuring their special day becomes ',
           'unforgettable and timeless.'], 6.332, 5.353, 6.672, 0.865);
  oval(s, 4.746, 1.565, 0.558, WHITE, { shadow: SH_DOT });
  ICONS.swap(s, 4.746 + 0.153, 1.565 + 0.153, 0.252, WINE);
}

/* 5 - "Music and Entertainment Planning Ideas" photo mosaic + date chip. */
function slide5(s) {
  header(s, 5);
  photo(s, 8.609, 0, 3.094, 1.212, [0, 0, 0.3, 0.3]);
  photo(s, 5.304, 1.429, 3.094, 3.094, 0.3);
  photo(s, 8.609, 1.429, 3.094, 3.094, 0.3);
  photo(s, 8.609, 4.74, 3.094, 2.76, [0.3, 0.3, 0, 0]);
  photo(s, 11.914, 4.74, 1.419, 2.76, [0.3, 0, 0, 0]);
  tx(s, '10 Years', 0.533, 1.429, 7.009, 0.64, { fontFace: HEAD, fontSize: 32, color: ROSE });
  para(s, ['A wedding plan organizes every detail carefully, blending love, vision, creativity, and celebration ',
           'with lasting joy.'], 0.541, 2.248, 4.376, 0.865);
  eyebrow(s, 'The Wedding Planner', 0.541, 4.819);
  headline(s, ['Music and Entertainment ', 'Planning Ideas'], 0.533, 5.184, 7.009, 1.439, 32, TX75);
  roundRect(s, 7.665, 3.908, 1.677, 1.723, WHITE, 0.21, { shadow: SH_CARD });
  tx(s, '25', 8.011, 4.319, 0.984, 0.774, { align: 'center', fontFace: HEAD, fontSize: 40, color: ROSE });
  tx(s, 'September', 7.784, 4.971, 1.438, 0.303, { align: 'center', color: ROSE });
  iconDot(s, 8.223, 3.632, 0.56, ROSE, 'lock');
}

/* 6 - "Curating the Perfect Wedding Menu" with two phone mock-ups. */
function slide6(s) {
  header(s, 6);
  photo(s, 9.788, 3.022, 2.507, 4.478, [0.35, 0.35, 0, 0]);
  photo(s, 6.294, 1.253, 2.497, 5.368, 0.37);
  phoneFrame(s, 6.11, 1.06, 2.87, 5.75);
  phoneFrame(s, 9.6, 2.83, 2.88, 4.67);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.588);
  headline(s, ['Curating the Perfect ', 'Wedding Menu'], 0.533, 1.952, 7.009, 1.439, 32, TX85);
  para(s, ['A wedding plan is a thoughtful journey, guiding couples ',
           'through love, vision, organization, and creativity, ',
           'ensuring their special day.'], 0.533, 5.561, 6.672, 0.865);
  captionPair(s, 'First Love', ['A wedding plan makes ', 'dreams reality.'], 9.788, 1.451);
  iconRow(s, 0.645, 4.654, 0.558, 0.227, [[DARK2, 'rings'], [WINE, 'lock'], [ROSE, 'swap']]);
}

/* Outline of a smartphone shell drawn around a placeholder screen. */
function phoneFrame(s, x, y, w, h) {
  s.addShape('roundRect', {
    x, y, w, h, fill: { type: 'none' }, rectRadius: 0.42,
    line: { color: TX95, width: 7 }, shadow: SH_CARD,
  });
  s.addShape('roundRect', { x: x + w / 2 - 0.42, y: y + 0.1, w: 0.84, h: 0.19, fill: { color: TX95 }, rectRadius: 0.09 });
}

/* 7 - dark "Designing a Unique Wedding Theme" with tilted photo cards. */
function slide7(s) {
  rect(s, 0, 0, 13.333, 7.5, DARK2);
  brandMark(s, WHITE);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.779, BLUSH75);
  headline(s, ['Designing a Unique ', 'Wedding Theme'], 0.533, 2.144, 7.009, 1.606, 36, WHITE);
  tx(s, 'Wedding plan is a thoughtful journey their special day.', 0.54, 4.043, 6.639, 0.34,
    { color: WHITE, lineSpacingMultiple: 1.3 });
  oval(s, 0.646, 5.865, 0.522, BLUSH75);
  tx(s, '01', 0.646, 5.974, 0.522, 0.303, { align: 'center', fontFace: HEAD, color: WHITE });
  para(s, ['A wedding plan creates harmony and lasting memories with true love and happiness.'],
    1.389, 5.825, 4.178, 0.602, { color: WHITE });
  blob(s, 'tiltCard', 9.342, -0.232, 4.141, 3.462, WHITE, { flipV: true, rotate: -6.16, shadow: SH_CARD });
  photoBlob(s, 'tiltPhoto', 9.419, 0, 3.914, 3.149);
  roundRect(s, 6.171, 2.532, 4.9, 3.881, WHITE, 0.48, { rotate: -6.16, shadow: SH_CARD });
  photo(s, 6.269, 2.574, 4.703, 3.796, 0.44);
  para(s, ['A wedding plan makes ', 'dreams reality.'], 5.935, 1.463, 2.654, 0.602, { align: 'right', color: WHITE });
  tx(s, 'First Love', 6.826, 1.016, 1.763, 0.38,
    { align: 'right', fontFace: HEAD, fontSize: 14, color: BLUSH75, lineSpacingMultiple: 1.3 });
  iconDot(s, 8.935, 1.143, 0.865, ROSE, 'lock');
}

/* 8 - "Creating a Memorable Guest Experience" with two date tiles. */
function slide8(s) {
  header(s, 8);
  card(s, 6.694, 0, 5.111, 4.496, WHITE, [0, 0, 0.43, 0.43], { shadow: SH_CARD });
  photo(s, 6.995, 0, 4.51, 4.208, [0, 0, 0.35, 0.35]);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.588);
  headline(s, ['Creating a Memorable ', 'Guest Experience'], 0.533, 1.952, 7.009, 1.439, 32, TX75);
  [[0.644, WINE, '22', 'September', ROSE, 'lock'], [3.156, ROSE, '10', 'November', WINE, 'swap']].forEach(
    ([x, fill, day, month, dot, icon]) => {
      roundRect(s, x, 4.478, 2.087, 2.067, fill, 0.26, { shadow: SH_DOT });
      tx(s, day, x + 0.257, 5.136, 1.459, 0.774, { fontFace: HEAD, fontSize: 40, color: WHITE });
      tx(s, month, x + 0.272, 5.91, 1.735, 0.37, { fontSize: 16, color: WHITE });
      iconDot(s, x + 1.412, 4.77, 0.398, dot, icon);
    });
  iconRow(s, 6.965, 5.192, 0.475, 0.192, [[WINE, 'lock'], [ROSE, 'swap']]);
  tx(s, 'Timeless Moments', 8.299, 5.227, 5.084, 0.404, { fontSize: 18, color: TX85 });
  para(s, ['The wedding plan is a thoughtful journey, guiding couples through ',
           'love, vision, organization, and creativity, ensuring their special day.'],
    6.854, 5.929, 6.639, 0.602);
}

/* 9 - "Floral Arrangements and Wedding Decor" - circular photo cluster. */
function slide9(s) {
  header(s, 9);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.618);
  headline(s, ['Floral Arrangements ', 'and Wedding Decor'], 0.533, 1.982, 7.009, 1.607, 36, TX85);
  para(s, ['A Wedding plan is a thoughtful journey, guiding couples through ',
           'love, vision, organization, and creativity, ensuring their special ',
           'day becomes timeless.'], 0.54, 5.569, 6.672, 0.865);
  photoBlob(s, 'petal', 8.298, 0, 3.613, 1.432);
  photo(s, 6.667, 1.796, 3.908, 3.908, 1.954);
  photoBlob(s, 'quarter', 10.021, 4.985, 3.313, 2.515);
  oval(s, 11.071, 2.94, 1.62, WINE, { shadow: SH_DOT });
  tx(s, '$95', 11.071, 3.464, 1.62, 0.572, { align: 'center', fontFace: HEAD, fontSize: 28, color: WHITE });
  iconRow(s, 0.645, 4.654, 0.558, 0.227, [[DARK2, 'rings'], [WINE, 'lock'], [ROSE, 'swap']]);
}

/* 10 - mirror of slide 9: photos left, copy right, two stat rows. */
function slide10(s) {
  header(s, 10);
  blob(s, 'leanCard', -0.284, 2.627, 3.449, 4.156, WHITE, { flipH: true, rotate: -7.44, shadow: SH_CARD });
  photoBlob(s, 'leanPhoto', 0, 2.754, 3.087, 3.978);
  roundRect(s, 2.26, 0.628, 4.156, 4.156, WHITE, 0.6, { rotate: 7.44, shadow: SH_CARD });
  photo(s, 2.341, 0.71, 3.992, 3.992, 0.58);
  oval(s, 1.596, 1.957, 1.238, WINE, { shadow: SH_DOT });
  tx(s, '$95', 1.405, 2.324, 1.62, 0.505, { align: 'center', fontFace: HEAD, fontSize: 24, color: WHITE });
  para(s, ['Love makes weddings ', 'unforgettable'], 3.979, 5.768, 2.687, 0.603);
  eyebrow(s, 'The Wedding Planner', 7.408, 1.492);
  headline(s, ['Floral Arrangements ', 'and Wedding Décor'], 7.408, 1.857, 5.648, 1.44, 32, TX75, { lineSpacingMultiple: null });
  tx(s, 'Wedding Plan is A Journey Their Special Day', 7.408, 4.256, 6.639, 0.34,
    { fontFace: HEAD, color: ROSE, lineSpacingMultiple: 1.3 });
  para(s, ['Wedding plan is a thoughtful journey, guiding couples through ',
           'love, organization, and creativity, ensuring their special day.'], 7.408, 4.746, 6.672, 0.602);
  iconDot(s, 7.517, 5.828, 0.482, WINE, 'lock');
  tx(s, '82,5%', 8.128, 5.783, 1.879, 0.572, { fontFace: HEAD, fontSize: 28 });
  iconDot(s, 10.007, 5.828, 0.482, ROSE, 'swap');
  tx(s, '82,5%', 10.618, 5.783, 1.636, 0.572, { fontFace: HEAD, fontSize: 28 });
}

/* 11 - "Selecting Colors and Style Elements": four feature cards. */
function slide11(s) {
  header(s, 11);
  eyebrow(s, 'The Wedding Planner', 5.336, 1.125, ROSE);
  tx(s, 'Selecting Colors and Style Elements', 1.919, 1.629, 9.495, 0.572,
    { align: 'center', fontFace: HEAD, fontSize: 28, color: TX75 });
  const CARDS = [
    { x: 0.997, active: false, icon: 'church' },
    { x: 4.016, active: true,  icon: 'cake' },
    { x: 7.034, active: false, icon: 'toast' },
    { x: 10.053, active: false, icon: 'propose' },
  ];
  CARDS.forEach(c => {
    roundRect(s, c.x, 2.722, 2.283, 2.724, c.active ? WINE : WHITE, 0.68, { shadow: SH_CARD });
    oval(s, c.x + 0.61, 3.075, 1.062, c.active ? DARK2 : GRAY95);
    ICONS[c.icon](s, c.x + 0.904, 3.368, 0.476, c.active ? WHITE : WINE);
    tx(s, [{ text: 'Organize Details ', options: { breakLine: true } }, { text: 'with Love' }],
      c.x + 0.227, 4.591, 1.829, 0.505, { align: 'center', color: c.active ? WHITE : TX75 });
  });
  iconRow(s, 0.642, 6.208, 0.475, 0.193, [[WINE, 'rings'], [ROSE, 'lock']]);
  para(s, ['The wedding plan organizes every detail carefully, blending love, vision, and celebration.'],
    2.246, 6.144, 4.058, 0.602, { color: TX85 });
  tx(s, '82,5%', 7.221, 6.126, 1.662, 0.64, { fontFace: HEAD, fontSize: 32, color: ROSE });
  para(s, ['The wedding plan organizes every detail carefully, ',
           'blending love, vision, creativity, and celebration.'], 8.917, 6.144, 4.852, 0.602, { color: TX85 });
}

/* 12 - "Take a Short and Refreshing Break" section divider. */
function slide12(s) {
  rect(s, 0, 0, 13.333, 7.5, DARK1);
  photo(s, 0, 0, 13.333, 7.5, 0);
  s.addShape('rect', {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { type: 'solid', color: DARK2, transparency: 0 },
    gradient: { type: 'linear', angle: 90, stops: [{ color: WINE, position: 0, transparency: 20 }, { color: DARK2, position: 100 }] },
  });
  s.addShape('heart', { x: 0.63, y: 3.06, w: 0.45, h: 0.42, fill: { color: BLUSH75 } });
  tx(s, 'Eternal Vows', 1.307, 3.061, 3.216, 0.404, { fontFace: HEAD, fontSize: 18, color: WHITE });
  pageMark(s, 12, WHITE, WHITE);
  tx(s, '60 Minutes', 6.669, 1.501, 6.164, 0.774, { align: 'right', fontFace: HEAD, fontSize: 40, color: WHITE });
  para(s, ['A wedding plan is a thoughtful journey, guiding ', 'couples through love, organization, and ',
           'ensuring their special day.'], 7.754, 2.477, 5.056, 0.865,
    { align: 'right', color: WHITE, transparency: 50 });
  s.addText([{ text: 'T', options: { fontFace: SCRIPT, fontSize: 88 } },
             { text: 'ake a Short ', options: { fontFace: HEAD, fontSize: 54 } }],
    { x: 0.574, y: 4.312, w: 9.993, h: 1.582, margin: INS, valign: 'top', color: WHITE });
  s.addText([{ text: 'and Refreshing ', options: { fontFace: BODY, fontSize: 54 } },
             { text: 'B', options: { fontFace: SCRIPT, fontSize: 88 } },
             { text: 'reak', options: { fontFace: BODY, fontSize: 54 } }],
    { x: 0.519, y: 5.36, w: 10.049, h: 1.582, margin: INS, valign: 'top', color: WHITE });
  para(s, ['A wedding plan makes ', 'dreams reality.'], 7.614, 5.881, 5.17, 0.602,
    { align: 'right', color: BLUSH });
}

/* 13 - "Choosing the Ideal Wedding Venue" with arch-topped photos. */
function slide13(s) {
  brandMark(s, WINE);
  pageMark(s, 13, TX85, ROSE);
  card(s, 5.658, 2.148, 4.363, 4.67, WHITE, [2.16, 2.16, 0.31, 0.31], { shadow: SH_CARD });
  card(s, 8.478, 0, 4.366, 2.191, WHITE, [0, 0, 2.18, 2.18], { shadow: SH_CARD });
  photo(s, 8.631, 0, 4.06, 2.016, [0, 0, 2.0, 2.0]);
  photo(s, 5.81, 2.332, 4.06, 4.303, [2.0, 2.0, 0.28, 0.28]);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.486);
  headline(s, ['Choosing the Ideal ', 'Wedding Venue'], 0.533, 1.851, 5.648, 1.607, 36, TX75);
  para(s, ['A wedding plan is a thoughtful journey, guiding couples ',
           'through love, vision, organization, and creativity, ',
           'ensuring their special day.'], 0.533, 5.392, 6.672, 0.865);
  iconRow(s, 0.645, 4.483, 0.558, 0.227, [[DARK2, 'rings'], [WINE, 'lock'], [ROSE, 'swap']]);
  [[4.004, '01', 3.964], [5.432, '02', 5.392]].forEach(([cy, num, ty]) => {
    oval(s, 9.713, cy, 0.522, WINE);
    tx(s, num, 9.713, cy + 0.11, 0.522, 0.303, { align: 'center', fontFace: HEAD, color: WHITE });
    para(s, ['Wedding planning makes ', 'dreams real.'], 10.483, ty, 2.85, 0.602);
  });
}

/* 14 - "Bridal Fashion and Groom's Attire" with tilted calendar strip. */
function slide14(s) {
  header(s, 14);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.398);
  headline(s, ['Bridal Fashion and Groom\u2019s Attire'], 0.533, 1.762, 5.648, 1.439, 32, TX75);
  tx(s, 'Wedding plan is a thoughtful journey their special day.', 0.533, 3.526, 6.639, 0.34,
    { lineSpacingMultiple: 1.3 });
  photo(s, 0, 4.531, 13.333, 2.969, 0);
  para(s, ['The wedding plan organizes every detail carefully, blending love, vision, and celebration beautifully',
           'together.'], 8.861, 1.569, 4.058, 0.865);
  iconDot(s, 7.109, 1.723, 0.558, WINE, 'lock');
  iconDot(s, 7.894, 1.723, 0.558, ROSE, 'swap');
  calendarStrip(s, 6.109, 3.136, -4.05, 'September 2030', 6.352, 3.523,
    [[6.489, 4.147, '18'], [7.554, 4.072, '19'], [8.695, 4.221, '20'], [9.835, 3.91, '21'], [10.901, 3.834, '22']],
    { card: [5.889, 2.125, 0.29], heart: [8.578, 3.989, 0.982, 0.863], dot: [11.251, 3.241, 0.338], boxH: 0.865 });
}

/* Rotated white calendar card carrying five day chips, the middle one a heart. */
function calendarStrip(s, x, y, rot, monthLabel, mx, my, days, o) {
  const [w, h, r] = o.card;
  roundRect(s, x, y, w, h, WHITE, r, { rotate: rot, shadow: SH_CARD });
  s.addShape('heart', { x: o.heart[0], y: o.heart[1], w: o.heart[2], h: o.heart[3], fill: { color: WINE }, rotate: rot });
  tx(s, monthLabel, mx, my, 2.661, 0.303, { rotate: rot, fontFace: HEAD, color: ROSE });
  const dw = 0.748 * (w / 5.889), dh = o.boxH;
  days.forEach(([dx, dy, label], i) => {
    const mid = i === Math.floor(days.length / 2);
    if (!mid) roundRect(s, dx, dy, dw, dh, GRAY95, 0.1, { rotate: rot });
    tx(s, label, dx, dy + dh * 0.266, dw, 0.404 * (h / 2.125),
      { rotate: rot, align: 'center', fontFace: HEAD, fontSize: 18, color: mid ? WHITE : GRAY75 });
  });
  iconDot(s, o.dot[0], o.dot[1], o.dot[2], WINE, 'lock');
}

/* 15 - "Photography and Videography Essentials" with laptop mock-ups. */
function slide15(s) {
  header(s, 15);
  laptop(s, 6.08, -2.244, 4.789, 6.005, true);
  laptop(s, 8.098, 3.75, 4.789, 6.005, false);
  photo(s, 6.762, 1.075, 3.491, 2.391, 0);
  photo(s, 8.779, 4.017, 3.493, 2.391, 0);
  iconDot(s, 0.642, 1.537, 0.558, WINE, 'lock');
  tx(s, '2030', 1.465, 1.429, 7.009, 0.774, { fontFace: HEAD, fontSize: 40, color: ROSE });
  para(s, ['The wedding plan carefully organizes every detail, combining love, ',
           'vision, creativity, and celebration to ensure a joyful, unforgettable, ',
           'and timeless day.'], 0.541, 2.402, 6.126, 0.865);
  eyebrow(s, 'The Wedding Planner', 0.541, 4.454);
  headline(s, ['Photography and Videography Essentials'], 0.533, 4.818, 7.009, 1.606, 36, TX85);
  captionPair(s, 'First Love', ['A wedding plan makes ', 'dreams reality.'], 10.68, 1.877);
  oval(s, 7.691, 3.162, 1.62, WINE, { shadow: SH_DOT });
  tx(s, '$95', 7.691, 3.686, 1.62, 0.572, { align: 'center', fontFace: HEAD, fontSize: 28, color: WHITE });
}

/* Open laptop silhouette: screen shell above, keyboard deck below. */
function laptop(s, x, y, w, h, flipped) {
  const lid = h * 0.62, deck = y + (flipped ? -0.02 : lid);
  s.addShape('roundRect', { x: x + w * 0.08, y: flipped ? y + h - lid : y, w: w * 0.84, h: lid,
    fill: { color: GRAY95 }, line: { color: TX95, width: 4 }, rectRadius: 0.08 });
  s.addShape('roundRect', { x, y: flipped ? deck - h * 0.06 : deck, w, h: h * 0.06,
    fill: { color: GRAY75 }, rectRadius: 0.04 });
}

/* 16 - dark "Bridal Fashion and Groom's Attire" with circular photos. */
function slide16(s) {
  rect(s, 0, 0, 13.333, 7.5, DARK2);
  oval(s, 4.249, 1.331, 4.834, WHITE, { shadow: SH_DOT });
  blob(s, 'scoopTR2', 9.437, 0, 3.896, 3.691, WHITE, { shadow: SH_DOT });
  blob(s, 'scoopBL', 0, 4.198, 4.1, 3.302, WHITE, { shadow: SH_DOT });
  brandMark(s, WHITE);
  pageMark(s, 16, WHITE, WHITE);
  tx(s, '16', 11.667, 6.827, 1.365, 0.399,
    { align: 'right', valign: 'middle', fontFace: HEAD, fontSize: 14, color: WINE, margin: [7.2, 3.6, 7.2, 3.6] });
  eyebrow(s, 'The Wedding Planner', 0.541, 1.398, BLUSH75);
  headline(s, ['Bridal Fashion and ', 'Groom\u2019s Attire'], 0.533, 1.762, 5.648, 1.272, 28, WHITE);
  photoBlob(s, 'scoopBL2', 0, 4.438, 3.861, 3.062);
  photo(s, 4.488, 1.569, 4.358, 4.358, 2.179);
  photoBlob(s, 'scoopTR2', 9.676, 0, 3.657, 3.452);
  tx(s, 'The Wedding Planner', 6.122, 5.432, 6.672, 0.42,
    { align: 'right', fontFace: HEAD, fontSize: 16, color: BLUSH75, lineSpacingMultiple: 1.3 });
  para(s, ['A wedding plan is a thoughtful journey, ',
           'guiding couples through love, vision, organization.'], 6.122, 6.108, 6.672, 0.602,
    { align: 'right', color: WHITE });
  oval(s, 9.7, 2.446, 1.87, ROSE, { shadow: SH_CARD });
  tx(s, '50%', 9.98, 2.985, 1.31, 0.572, { align: 'center', fontFace: HEAD, fontSize: 28, color: WHITE });
  tx(s, 'The Memories', 9.82, 3.473, 1.63, 0.303, { align: 'center', color: WHITE });
  iconDot(s, 2.485, 4.087, 1.165, ROSE, 'lock');
}

/* 17 - "Choosing the Ideal Wedding Venue" with the bride & groom illustration. */
function slide17(s) {
  header(s, 17);
  coupleArt(s, 4.302, 2.139, 4.729, 4.258);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.51);
  headline(s, ['Choosing the ', 'Ideal Wedding', 'Venue'], 0.533, 1.875, 5.648, 2.48, 40, TX75, { lineSpacing: 58 });
  para(s, ['The Wedding plan is a thoughtful journey, ', 'guiding couples through love, and ensuring ',
           'their special day.'], 0.538, 5.327, 4.864, 0.865);
  calendarStrip(s, 8.916, 1.726, 0, 'September 2030', 9.396, 1.907,
    [[9.481, 2.397, '19'], [10.474, 2.397, '20'], [11.467, 2.397, '21']],
    { card: [3.765, 1.702, 0.23], heart: [10.373, 2.395, 0.853, 0.749], dot: [11.83, 1.892, 0.293], boxH: 0.751 });
  iconDot(s, 8.862, 4.909, 0.546, WINE, 'lock');
  tx(s, '2030', 9.684, 4.829, 1.879, 0.707, { fontFace: HEAD, fontSize: 36 });
  para(s, ['Wedding planning creates love, and', 'unforgettable memories.'], 9.705, 5.586, 3.519, 0.602);
}

/* Flat-vector bride & groom, assembled from primitive shapes. */
function coupleArt(s, x, y, w, h) {
  const SKIN = 'FFB8AB', SKIN2 = 'FFDFD6', SUIT = '0D0D0D', GOWN = 'FEF0F3', GOWN2 = 'FFE4E1';
  s.addShape('ellipse', { x: x + w * 0.21, y: y + h * 0.85, w: w * 0.62, h: h * 0.15, fill: { color: 'F2EDED' } });
  /* bride: veil, gown, bouquet, head */
  s.addShape('ellipse', { x: x + w * 0.25, y: y + h * 0.04, w: w * 0.4, h: h * 0.55, fill: { color: 'FEF0F3', transparency: 45 } });
  s.addShape('triangle', { x: x + w * 0.21, y: y + h * 0.42, w: w * 0.45, h: h * 0.47, fill: { color: GOWN2 } });
  s.addShape('triangle', { x: x + w * 0.25, y: y + h * 0.45, w: w * 0.36, h: h * 0.44, fill: { color: GOWN } });
  s.addShape('roundRect', { x: x + w * 0.36, y: y + h * 0.2, w: w * 0.14, h: h * 0.26, fill: { color: GOWN2 }, rectRadius: 0.12 });
  s.addShape('ellipse', { x: x + w * 0.37, y: y + h * 0.08, w: w * 0.12, h: h * 0.13, fill: { color: SKIN } });
  s.addShape('ellipse', { x: x + w * 0.38, y: y + h * 0.03, w: w * 0.1, h: h * 0.08, fill: { color: SUIT } });
  s.addShape('ellipse', { x: x + w * 0.33, y: y + h * 0.42, w: w * 0.14, h: h * 0.13, fill: { color: WINE } });
  /* groom: suit, shirt, hat, head */
  s.addShape('roundRect', { x: x + w * 0.55, y: y + h * 0.19, w: w * 0.2, h: h * 0.33, fill: { color: SUIT }, rectRadius: 0.1 });
  s.addShape('triangle', { x: x + w * 0.61, y: y + h * 0.2, w: w * 0.08, h: h * 0.19, fill: { color: GOWN } });
  s.addShape('roundRect', { x: x + w * 0.58, y: y + h * 0.5, w: w * 0.16, h: h * 0.42, fill: { color: SUIT }, rectRadius: 0.06 });
  s.addShape('ellipse', { x: x + w * 0.6, y: y + h * 0.06, w: w * 0.11, h: h * 0.12, fill: { color: SKIN2 } });
  s.addShape('rect', { x: x + w * 0.575, y: y + h * 0.045, w: w * 0.16, h: h * 0.016, fill: { color: SUIT } });
  s.addShape('rect', { x: x + w * 0.6, y: y + h * 0.0, w: w * 0.11, h: h * 0.05, fill: { color: SUIT } });
  s.addShape('rect', { x: x + w * 0.46, y: y + h * 0.44, w: w * 0.14, h: h * 0.03, fill: { color: SKIN } });
}

/* 18 - full-width photo band under "Bridal Fashion and Groom's Attire". */
function slide18(s) {
  header(s, 18);
  card(s, 0, 2.736, 11.95, 2.802, WHITE, [0, 0.31, 0.31, 0], { shadow: SH_CARD });
  eyebrow(s, 'The Wedding Planner', 0.541, 1.32);
  tx(s, 'Bridal Fashion and Groom\u2019s Attire', 0.533, 1.742, 9.314, 0.572,
    { fontFace: HEAD, fontSize: 28, color: TX85 });
  photo(s, 0, 2.888, 1.123, 2.499, [0, 0.26, 0.26, 0]);
  photo(s, 1.283, 2.888, 2.499, 2.499, 0.26);
  photo(s, 3.941, 2.888, 5.157, 2.499, 0.26);
  photo(s, 9.257, 2.888, 2.499, 2.499, 0.26);
  iconDot(s, 8.794, 3.753, 0.768, ROSE, 'lock');
  iconRow(s, 0.642, 6.208, 0.475, 0.193, [[WINE, 'rings'], [ROSE, 'church']]);
  para(s, ['The wedding plan organizes every detail ', 'carefully, blending love, and celebration.'],
    2.246, 6.144, 4.058, 0.602, { color: TX85 });
  tx(s, '82,5%', 6.313, 6.126, 1.662, 0.64, { fontFace: HEAD, fontSize: 32, color: ROSE });
  para(s, ['The wedding plan organizes every detail carefully, ',
           'blending love, vision, creativity, and celebration.'], 8.009, 6.144, 4.852, 0.602, { color: TX85 });
}

/* 19 - "Contact Our Wedding Team" closing details. */
function slide19(s) {
  header(s, 19);
  card(s, 5.75, 0, 5.917, 5.434, WHITE, [0, 0, 2.96, 2.96], { shadow: SH_CARD });
  photo(s, 6.053, 0, 5.311, 5.079, [0, 0, 2.655, 2.655]);
  eyebrow(s, 'The Wedding Planner', 0.541, 1.51);
  headline(s, ['Contact Our Wedding Team'], 0.533, 1.875, 5.648, 1.666, 40, TX75, { lineSpacing: 58 });
  iconDot(s, 0.642, 4.524, 0.475, ROSE, 'toast');
  [['Phone Number', '+123-4567-8900', 5.368], ['Email Address', 'yourmail@addmail.com', 5.802],
   ['Website Link', 'www.yourwebsite.com', 6.236]].forEach(([label, value, y]) => {
    tx(s, label, 0.528, y, 2.2, 0.303);
    tx(s, ': ' + value, 2.55, y, 3.7, 0.303);
  });
  oval(s, 6.262, 3.965, 1.87, WINE, { shadow: SH_CARD });
  tx(s, '50%', 6.542, 4.504, 1.31, 0.572, { align: 'center', fontFace: HEAD, fontSize: 28, color: WHITE });
  tx(s, 'The Memories', 6.382, 4.993, 1.63, 0.303, { align: 'center', color: WHITE });
  iconRow(s, 10.564, 6.054, 0.558, 0.227, [[DARK2, 'rings'], [WINE, 'lock'], [ROSE, 'swap']]);
}

/* ------------------------------------------------------------------ compose */
const BUILDERS = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
                  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

BUILDERS.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '0cbf3396-75ec-485e-bee9-fb358b246bb8_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
