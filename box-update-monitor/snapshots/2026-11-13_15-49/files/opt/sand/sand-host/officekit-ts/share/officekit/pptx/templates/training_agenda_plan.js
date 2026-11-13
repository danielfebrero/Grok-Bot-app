/**
 * "Training Agenda" deck - rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9), 21 slides.
 *
 * Raster artwork in the source deck is replaced by native shape placeholders
 * (see picturePlaceholder / icon helpers below).
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const FONT = 'Space Grotesk';

const THEME = {
  pink: 'D9326F', // accent1
  cyan: '0388A6', // accent2
  yellow: 'D9A404', // accent3
  gold: 'BF7E04', // accent4
  red: 'D90404', // accent5
  white: 'FFFFFF',
  black: '000000'
};

/** OOXML lumMod/lumOff applied in HSL space (how the deck derives tints). */
function lum(hex, lumMod = 1, lumOff = 0) {
  const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let l = (max + min) / 2;
  const d = max - min;
  const denom = 1 - Math.abs(2 * l - 1);
  const s = d === 0 || denom === 0 ? 0 : d / denom;
  if (d !== 0) {
    if (max === r) h = (g - b) / d;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h = ((h * 60) % 360 + 360) % 360;
  }
  l = Math.min(1, Math.max(0, l * lumMod + lumOff));
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  const rgb = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x]
    : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return rgb.map(v => Math.round((v + m) * 255).toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** Blend two hex colours (t = 0 -> a, t = 1 -> b). */
function mix(a, b, t = 0.5) {
  return [0, 2, 4]
    .map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t))
    .map(v => v.toString(16).padStart(2, '0').toUpperCase())
    .join('');
}

const light = hex => lum(hex, 0.6, 0.4); // the "lumMod 60 / lumOff 40" tint used by every card gradient
/** Flat stand-in for the deck's 45 degree accent gradient. */
const grad = hex => mix(hex, light(hex));

const C = {
  pink: THEME.pink,
  cyan: THEME.cyan,
  yellow: THEME.yellow,
  gold: THEME.gold,
  red: THEME.red,
  white: 'FFFFFF',
  dark: lum(THEME.black, 0.85, 0.15), // 262626 - body headings
  gray: lum(THEME.black, 0.5, 0.5), // 808080 - body copy
  mute: lum(THEME.black, 0.65, 0.35), // 595959 - calendar digits
  silver: lum('FFFFFF', 0.75), // BFBFBF - captions, page numbers
  hair: lum('FFFFFF', 0.85), // D9D9D9 - rules
  wash: lum('FFFFFF', 0.95), // F2F2F2 - calendar cell fill
  pinkPale: lum(THEME.pink, 0.2, 0.8), // Sunday column tint
  pinkSoft: lum(THEME.pink, 0.4, 0.6),
  pinkDeep: lum(THEME.pink, 0.75)
};

const ACCENTS = [C.pink, C.cyan, C.yellow, C.gold, C.red];

/* Shadows used throughout the deck. pptxgenjs rewrites the object it is handed,
   so every shape needs its own copy. */
const shadow = (blur, offset, opacity = 0.15, angle = 90) =>
  () => ({ type: 'outer', blur, offset, angle, color: '000000', opacity });

const SH_CARD = shadow(80, 10); // wide, very soft card drop
const SH_SOFT = shadow(30, 8); // tighter shadow on inner panels and badges

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */

/** Text box; defaults mirror PowerPoint (top anchored, left aligned). */
function T(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, fontSize: 12, color: C.dark, valign: 'top' }, o));
}

/** Rounded rectangle card. `r` is the corner radius in inches. */
function card(slide, o) {
  const { x, y, w, h, r = 0.15, fill, line, shadow, flipH, shape = 'roundRect' } = o;
  slide.addShape(shape, {
    x, y, w, h,
    rectRadius: r,
    fill: fill ? (typeof fill === 'string' ? { color: fill } : fill) : { type: 'none' },
    line: line || { type: 'none' },
    ...(shadow ? { shadow } : {}),
    ...(flipH ? { flipH: true } : {})
  });
}

/** Card + centred label, the deck's most repeated element. */
function labelCard(slide, o) {
  card(slide, o);
  if (o.text) {
    T(slide, o.text, {
      x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle',
      fontSize: o.fontSize || 12, bold: o.bold, color: o.color || C.white, margin: 0
    });
  }
}

/**
 * Fakes a linear gradient with plain strips (square-cornered areas only).
 * `dir`: 'h' left-to-right, 'v' top-to-bottom, 'diag' top-left to bottom-right.
 */
function gradBand(slide, o) {
  const { x, y, w, h, from, to, steps = 28, dir = 'h' } = o;
  if (dir === 'diag') {
    const cols = 24;
    const rows = 14;
    const cw = w / cols;
    const ch = h / rows;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const t = (c / (cols - 1) + r / (rows - 1)) / 2;
        slide.addShape('rect', {
          x: x + c * cw, y: y + r * ch, w: cw + 0.012, h: ch + 0.012,
          fill: { color: mix(from, to, t) }, line: { type: 'none' }
        });
      }
    }
    return;
  }
  for (let i = 0; i < steps; i++) {
    const c = mix(from, to, steps === 1 ? 0 : i / (steps - 1));
    const span = (dir === 'h' ? w : h) / steps;
    slide.addShape('rect', {
      x: dir === 'h' ? x + i * span : x,
      y: dir === 'h' ? y : y + i * span,
      w: dir === 'h' ? span + 0.012 : w,
      h: dir === 'h' ? h : span + 0.012,
      fill: { color: c }, line: { type: 'none' }
    });
  }
}

/* --- pictogram icons (drawn from primitives, no bitmaps) ----------- */

function iconMail(slide, x, y, s, color) {
  slide.addShape('rect', { x, y: y + 0.1 * s, w: s, h: 0.62 * s, fill: { color }, line: { type: 'none' } });
  slide.addShape('triangle', {
    x: x + 0.06 * s, y: y + 0.16 * s, w: 0.88 * s, h: 0.42 * s,
    flipV: true, fill: { color: C.white }, line: { color, width: 1.5 }
  });
}

function iconCalendar(slide, x, y, s, color) {
  slide.addShape('roundRect', {
    x, y: y + 0.12 * s, w: s, h: 0.88 * s, rectRadius: 0.03 * s,
    fill: { color }, line: { type: 'none' }
  });
  slide.addShape('rect', { x: x + 0.1 * s, y: y + 0.44 * s, w: 0.8 * s, h: 0.44 * s, fill: { color: C.white }, line: { type: 'none' } });
  [0.2, 0.72].forEach(f => slide.addShape('rect', { x: x + f * s, y, w: 0.08 * s, h: 0.2 * s, fill: { color }, line: { type: 'none' } }));
}

function iconClipboard(slide, x, y, s, color) {
  slide.addShape('roundRect', {
    x, y: y + 0.11 * s, w: s, h: 0.89 * s, rectRadius: 0.06 * s,
    fill: { color }, line: { type: 'none' }
  });
  slide.addShape('rect', { x: x + 0.25 * s, y, w: 0.5 * s, h: 0.22 * s, fill: { color }, line: { color: C.white, width: 1 } });
}

function iconPin(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x, y, w, h: h * 0.68, fill: { color }, line: { type: 'none' } });
  slide.addShape('triangle', {
    x: x + w * 0.2, y: y + h * 0.42, w: w * 0.6, h: h * 0.58,
    flipV: true, fill: { color }, line: { type: 'none' }
  });
  slide.addShape('ellipse', {
    x: x + w * 0.3, y: y + h * 0.18, w: w * 0.4, h: h * 0.26,
    fill: { color: C.white }, line: { type: 'none' }
  });
}

const ICONS = { mail: iconMail, calendar: iconCalendar, clipboard: iconClipboard };

/** ">" / "<" chevron stroke. `dir` = 1 points right, -1 points left. */
function chevron(slide, x, y, w, h, color, dir = 1) {
  const a = dir > 0 ? 0 : w;
  const b = dir > 0 ? w : 0;
  slide.addShape('custGeom', {
    x, y, w, h,
    points: [{ x: a, y: 0 }, { x: b, y: h / 2 }, { x: a, y: h }],
    fill: { type: 'none' }, line: { color, width: 1.25 }
  });
}

/** White (or tinted) circle badge holding a pictogram. */
function iconBadge(slide, o) {
  const { x, y, d = 0.669, bg = C.white, fg, kind, drop = SH_SOFT } = o;
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: bg }, line: { type: 'none' }, shadow: drop() });
  const s = d * 0.43;
  ICONS[kind](slide, x + (d - s) / 2, y + (d - s) / 2, s, fg);
}

/**
 * Empty picture placeholder: translucent panel + prompt caption + a small
 * image pictogram. Matches the layout's "bg1 lumMod 95% / alpha 75%" frame.
 */
function picturePlaceholder(slide, o) {
  const { x, y, w, h, label = 'Image Placeholder', panel } = o;
  if (panel) {
    slide.addShape('rect', { x, y, w, h, fill: { color: C.wash, transparency: 25 }, line: { type: 'none' } });
  }
  T(slide, label, { x, y: y + 0.04, w, h: 0.33, align: 'center', fontSize: 14, color: C.silver });
  const gw = 0.62;
  const gh = 0.52;
  const gx = x + (w - gw) / 2;
  const gy = y + (h - gh) / 2;
  slide.addShape('rect', { x: gx, y: gy, w: gw, h: gh, fill: { color: C.white }, line: { color: '404040', width: 0.75 } });
  slide.addShape('triangle', { x: gx + 0.1, y: gy + 0.2, w: 0.42, h: 0.24, fill: { color: '8EA9DB' }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: gx + 0.12, y: gy + 0.08, w: 0.1, h: 0.1, fill: { color: 'FFC000' }, line: { type: 'none' } });
}

/** "2024 | Agenda Plan" strap-line + rule + page number drawn by the slide master. */
function masterChrome(slide, pageNo, o = {}) {
  const tint = o.tint || C.silver;
  T(slide, [
    { text: '2024 ', options: { bold: true, color: o.brand || C.pink } },
    { text: '|', options: { color: tint } },
    { text: ' Agenda Plan', options: { bold: true, color: tint } }
  ], { x: 0.581, y: 0.169, w: 3.461, h: 0.286, fontSize: 11, valign: 'middle' });
  slide.addShape('line', { x: 0, y: 0.312, w: 0.438, h: 0, line: { color: tint, width: 1 } });
  T(slide, String(pageNo), {
    x: 11.564, y: 6.915, w: 1.586, h: 0.438, align: 'right', fontSize: 20, color: tint
  });
}

/* --- calendars ----------------------------------------------------- */

/**
 * Day numbers for one calendar page.
 * Returns {n, out} where `out` marks leading/trailing days of adjacent months.
 */
function monthDays(lead, days, prevLast, count) {
  const cells = [];
  for (let i = 0; i < count; i++) {
    if (i < lead) cells.push({ n: prevLast - lead + 1 + i, out: true });
    else if (i < lead + days) cells.push({ n: i - lead + 1, out: false });
    else cells.push({ n: i - lead - days + 1, out: true });
  }
  return cells;
}

/** Grid runner shared by every calendar on the deck. */
function grid(o, draw) {
  const { x, y, dx, dy, cols = 7, count } = o;
  for (let i = 0; i < count; i++) {
    draw(i, x + (i % cols) * dx, y + Math.floor(i / cols) * dy);
  }
}

/* ================================================================== *
 * Slides
 * ================================================================== */

/* 1 - Title slide over the pink radial background. */
function slide01(s) {
  gradBand(s, { x: 0, y: 0, w: 13.333, h: 7.5, from: C.pinkSoft, to: C.pink, dir: 'diag' });
  T(s, '2024', {
    x: 1.59, y: 1.284, w: 10.152, h: 4.931, align: 'center', valign: 'middle',
    fontSize: 287, bold: true, color: C.white, transparency: 92, wrap: false
  });
  T(s, 'Training Agenda', {
    x: 1.691, y: 2.572, w: 9.951, h: 1.582, align: 'center', valign: 'middle',
    fontSize: 88, bold: true, color: C.white, wrap: false
  });
  T(s, 'Presentation Template', {
    x: 3.459, y: 4.154, w: 6.415, h: 0.774, align: 'center', valign: 'middle',
    fontSize: 40, color: C.white, wrap: false
  });
}

/* 2 - Photo grid with five agenda captions. */
function slide02(s) {
  const GREY_TOP = 'F7F7F7';
  const GREY_BOT = 'BEBEBE';
  const tiles = [
    { x: 1.349, y: 2.194, w: 2.488, h: 2.062, label: 'First Agenda', text: 'Lorem ipsum dolor sit amet, consectetur', tx: 1.535, ty: 3.07, tw: 2.283 },
    { x: 4.079, y: 2.194, w: 2.488, h: 2.062, label: 'Second Agenda', text: 'Lorem ipsum dolor sit amet, consectetur', tx: 4.269, ty: 3.07, tw: 2.283 },
    { x: 6.775, y: 2.194, w: 2.504, h: 4.341, label: 'Third Agenda', text: 'Lorem ipsum dolor sit amet, consectetur', tx: 6.918, ty: 5.298, tw: 2.283 },
    { x: 9.493, y: 2.194, w: 2.488, h: 2.062, label: 'Fourth Agenda', text: 'Lorem ipsum dolor sit amet, consectetur', tx: 9.689, ty: 3.07, tw: 2.283 },
    { x: 1.349, y: 4.467, w: 5.195, h: 2.062, label: 'Fifth Agenda', text: LOREM, tx: 1.535, ty: 5.304, tw: 4.097 }
  ];

  tiles.forEach(t => gradBand(s, { x: t.x, y: t.y, w: t.w, h: t.h, from: GREY_TOP, to: GREY_BOT, steps: 22, dir: 'v' }));

  T(s, 'Our Agenda Next Month', {
    x: 2.738, y: 0.855, w: 7.88, h: 0.841, align: 'center', fontSize: 44, color: C.dark
  });

  /* pink "2024 / Oct - Nov" tile */
  card(s, { x: 9.493, y: 4.467, w: 2.488, h: 2.062, r: 0.217, fill: grad(C.pink), shadow: SH_CARD() });
  T(s, '2024', { x: 9.972, y: 4.907, w: 1.53, h: 0.681, align: 'center', fontSize: 32, color: C.white });
  T(s, 'Oct - Nov', {
    x: 9.859, y: 5.598, w: 1.756, h: 0.391, align: 'center', fontSize: 16, color: C.white, charSpacing: 3
  });
  chevron(s, 9.715, 5.729, 0.069, 0.124, C.white, -1);
  chevron(s, 11.69, 5.729, 0.069, 0.124, C.white, 1);

  tiles.forEach(t => {
    T(s, t.label, { x: t.tx, y: t.ty, w: 2.109, h: 0.37, fontSize: 16, color: C.white });
    T(s, t.text, { x: t.tx, y: t.ty + 0.376, w: t.tw, h: 0.607, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });

  tiles.forEach(t => picturePlaceholder(s, t));
  masterChrome(s, 2);
}

/* 3 - Four dated chips with descriptions. */
function slide03(s) {
  const items = [
    { x: 1.072, y: 0.646, day: '12', month: 'August', color: C.pink, tx: 2.636, title: 'Agenda 01' },
    { x: 5.983, y: 0.646, day: '19', month: 'August', color: C.cyan, tx: 7.547, title: 'Agenda 02' },
    { x: 3.583, y: 5.555, day: '26', month: 'August', color: C.yellow, tx: 5.147, title: 'Agenda 03' },
    { x: 8.494, y: 5.555, day: '4', month: 'September', color: C.gold, tx: 10.058, title: 'Agenda 04' }
  ];
  items.forEach(it => {
    card(s, { x: it.x, y: it.y, w: 1.378, h: 1.299, r: 0.145, fill: grad(it.color), shadow: SH_CARD() });
    const t = { x: it.x, w: 1.378, align: 'center', color: C.white, lineSpacingMultiple: 1.2 };
    T(s, 'Monday', Object.assign({ y: it.y + 0.056, h: 0.367, fontSize: 14 }, t));
    T(s, it.day, Object.assign({ y: it.y + 0.353, h: 0.557, fontSize: 24, bold: true }, t));
    T(s, it.month, Object.assign({ y: it.y + 0.83, h: 0.367, fontSize: 14 }, t));
    T(s, it.title, { x: it.tx, y: it.y, w: 2.375, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    T(s, LOREM, {
      x: it.tx, y: it.y + 0.384, w: 2.511, h: 0.867, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 3);
}

/* 4 - Week grid: five day rows x five time chips. */
function slide04(s) {
  const rows = [
    { label: 'Mon, 18', color: C.pink, slots: ['08.00 AM', '09.00 AM', '10.00 AM', '11.00 AM', '12.00 AM'] },
    { label: 'Tue, 19', color: C.cyan, slots: ['09.00 AM', null, '11.00 AM', '12.00 AM', '13.00 AM'] },
    { label: 'Wed, 20', color: C.yellow, slots: ['08.00 AM', '09.00 AM', '10.00 AM', null, '12.00 AM'] },
    { label: 'Thu, 21', color: C.gold, slots: ['08.00 AM', '09.00 AM', null, '10.00 AM', '11.00 AM'] },
    { label: 'Fri, 22', color: C.red, slots: ['08.00 AM', null, '10.00 AM', '10.00 AM', null] }
  ];
  const COLS = [2.694, 4.55, 6.407, 8.263, 10.119];

  T(s, 'Training Agenda This Week', {
    x: 2.231, y: 0.601, w: 8.902, h: 0.841, align: 'center', fontSize: 44, color: C.dark
  });
  [1.73, 2.764, 3.798, 4.831, 5.865, 6.899].forEach(y =>
    s.addShape('line', { x: 0.792, y, w: 11.781, h: 0, line: { color: C.hair, width: 1 } }));

  rows.forEach((row, r) => {
    const y = 1.93 + r * 1.034;
    T(s, row.label, { x: 0.775, y: y + 0.132, w: 1.676, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    row.slots.forEach((slot, c) => {
      if (!slot) return;
      labelCard(s, {
        x: COLS[c] + 0.362, y, w: 0.952, h: 0.27, r: 0.054, fill: grad(row.color),
        text: slot, fontSize: 11
      });
      T(s, 'Agenda title here', {
        x: COLS[c], y: y + 0.297, w: 1.676, h: 0.336,
        align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
      });
    });
  });
  masterChrome(s, 4);
}

/* Calendar used on slides 5 and 7: pink card, white sheet, 7 x 6 tiled days. */
function tiledCalendar(s, o) {
  const { x, y, title, highlights } = o;
  card(s, { x, y, w: 4.409, h: 4.528, r: 0.197, fill: grad(C.pink), shadow: SH_CARD() });
  card(s, { x: x + 0.062, y: y + 0.56, w: 4.285, h: 3.885, r: 0.181, fill: C.white, shadow: SH_SOFT() });
  T(s, title, {
    x: x + 0.987, y: y + 0.101, w: 2.435, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.white
  });

  const x0 = x + 0.169;
  const y0 = y + 0.759;
  ['M', 'T', 'W', 'T', 'F', 'S', 'S'].forEach((d, i) =>
    T(s, d, {
      x: x0 + i * 0.582, y: y0, w: 0.582, h: 0.511, align: 'center', valign: 'middle',
      fontSize: 12, color: C.dark, margin: 0
    }));

  const days = monthDays(6, 31, 30, 42);
  grid({ x: x0, y: y0 + 0.511, dx: 0.582, dy: 0.511, count: 42 }, (i, cx, cy) => {
    const d = days[i];
    const hit = !d.out && highlights[d.n];
    const sunday = i % 7 === 6;
    const fill = d.out ? { type: 'none' } : hit ? { color: grad(hit) } : { color: sunday ? C.pinkPale : C.wash };
    labelCard(s, {
      x: cx, y: cy, w: 0.582, h: 0.511, r: 0.077, fill,
      line: { color: C.white, width: 0.5 },
      text: String(d.n), fontSize: 12,
      color: hit ? C.white : d.out ? C.silver : C.mute
    });
  });
}

/* 5 - September calendar + meeting strip. */
function slide05(s) {
  card(s, { x: 4.921, y: -0.002, w: 1.753, h: 7.5, r: 0, shape: 'rect', fill: C.white, shadow: { type: 'outer', blur: 40, offset: 30, angle: 0, color: '000000', opacity: 0.09 } });
  tiledCalendar(s, { x: 1.401, y: 1.485, title: 'September 2024', highlights: { 20: C.pink } });

  T(s, 'Our Agenda ', { x: 7.353, y: 1.61, w: 4.394, h: 0.841, align: 'right', fontSize: 44, color: C.dark });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis.', {
    x: 7.353, y: 2.466, w: 4.394, h: 0.861, align: 'right', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });

  card(s, { x: 6.369, y: 3.822, w: 5.408, h: 1.523, r: 0.152, fill: grad(C.pink), shadow: SH_CARD() });
  T(s, 'Internal Meeting', { x: 7.232, y: 4.147, w: 2.435, h: 0.37, fontSize: 16, color: C.white });
  T(s, '10:00 \u2013 11:30', { x: 6.89, y: 4.629, w: 2.778, h: 0.37, fontSize: 16, color: C.white });
  s.addShape('line', { x: 10.011, y: 4.111, w: 0, h: 0.945, line: { color: C.white, width: 4.5 } });
  T(s, 'Friday', { x: 9.922, y: 4.031, w: 1.439, h: 0.37, align: 'center', fontSize: 16, color: C.white });
  T(s, '10', { x: 9.887, y: 4.409, w: 1.525, h: 0.707, align: 'center', fontSize: 36, bold: true, color: C.white, charSpacing: -1.5 });
  iconBadge(s, { x: 11.442, y: 4.992, kind: 'calendar', fg: C.pink });

  iconPin(s, 11.425, 6.012, 0.322, 0.513, grad(C.pink));
  T(s, 'Jakarta, Indonesia', {
    x: 7.707, y: 6.117, w: 3.538, h: 0.37, align: 'right', valign: 'bottom', fontSize: 16, bold: true, color: C.dark
  });
  masterChrome(s, 5);
}

/* 6 - Timeline with three dated cards. */
function slide06(s) {
  const items = [
    { x: 1.049, title: 'Agenda One', date: '23 February', color: C.pink, up: false, icon: 'mail', ix: 3.509 },
    { x: 5.234, title: 'Agenda Two', date: '29 February', color: C.cyan, up: false, icon: 'calendar', ix: 7.701 },
    { x: 9.419, title: 'Agenda Three', date: '7 March', color: C.yellow, up: false, icon: 'clipboard', ix: 11.876 }
  ];
  s.addShape('line', { x: 0, y: 4.785, w: 13.333, h: 0, line: { color: C.hair, width: 1.25 } });

  T(s, '2024', {
    x: 2.67, y: 0.009, w: 7.998, h: 3.45, align: 'center', fontSize: 199, bold: true,
    color: lum('FFFFFF', 0.65, 0), transparency: 90, charSpacing: -3
  });
  T(s, 'Our Agenda', { x: 3.101, y: 1.256, w: 7.202, h: 0.841, align: 'center', fontSize: 44, color: C.dark });

  items.forEach(it => {
    T(s, it.title, { x: it.x, y: 3.417, w: 2.834, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.dark });
    card(s, { x: it.x, y: 4.1, w: 2.834, h: 1.361, r: 0.145, fill: grad(it.color), shadow: SH_CARD() });
    T(s, it.date, { x: it.x, y: 4.4, w: 2.834, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.white });
    T(s, '09am \u2013 1pm', { x: it.x, y: 4.747, w: 2.834, h: 0.415, align: 'center', fontSize: 16, color: C.white });
    T(s, LOREM, { x: it.x, y: 5.678, w: 2.834, h: 0.861, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3 });
    iconBadge(s, { x: it.ix, y: 3.737, kind: it.icon, fg: it.color });
  });
  masterChrome(s, 6);
}

/* 7 - August calendar with a three-item agenda list. */
function slide07(s) {
  tiledCalendar(s, { x: 7.494, y: 1.485, title: 'August 2024', highlights: { 10: C.pink, 18: C.cyan, 26: C.yellow } });

  T(s, 'Training Agenda This Month', { x: 0.952, y: 1.296, w: 5.095, h: 1.582, fontSize: 44, color: C.dark });
  const rows = [
    { y: 3.078, title: 'Agenda 01', color: C.pink, icon: 'mail' },
    { y: 4.168, title: 'Agenda 02', color: C.cyan, icon: 'calendar' },
    { y: 5.257, title: 'Agenda 03', color: C.yellow, icon: 'clipboard' }
  ];
  rows.forEach(r => {
    iconBadge(s, { x: 0.952, y: r.y + 0.003, kind: r.icon, bg: grad(r.color), fg: C.white, drop: SH_CARD });
    T(s, r.title, { x: 1.719, y: r.y, w: 4.12, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt', {
      x: 1.719, y: r.y + 0.379, w: 4.704, h: 0.607, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 7);
}

/* 8 - Zig-zag timeline, three alternating cards. */
function slide08(s) {
  const items = [
    { title: 'Agenda 01', date: '26 August 2024', color: C.pink, icon: 'mail', low: true, cx: 1.635, tx: 1.697, dot: 2.911, bx: 3.971 },
    { title: 'Agenda 02', date: '27 August 2024', color: C.cyan, icon: 'calendar', low: false, cx: 5.335, tx: 5.399, dot: 6.608, bx: 7.67 },
    { title: 'Agenda 03', date: '28 August 2024', color: C.yellow, icon: 'clipboard', low: true, cx: 8.833, tx: 8.894, dot: 10.109, bx: 11.169 }
  ];
  s.addShape('line', { x: 0, y: 2.383, w: 13.333, h: 0, line: { color: C.hair, width: 1.25 } });

  items.forEach(it => {
    const cardY = it.low ? 2.616 : 0.847;
    const titleY = it.low ? 0.88 : 2.676;
    const bodyY = it.low ? 1.267 : 3.056;
    card(s, { x: it.cx, y: cardY, w: 2.671, h: 1.233, r: 0.139, fill: grad(it.color), shadow: SH_CARD() });
    T(s, it.date, { x: it.cx + 0.067, y: cardY + 0.415, w: 2.536, h: 0.404, align: 'center', fontSize: 18, color: C.white });
    T(s, it.title, { x: it.tx, y: titleY, w: 2.521, h: 0.37, align: 'center', fontSize: 16, bold: true, color: it.color });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', {
      x: it.tx + 0.14, y: bodyY, w: 2.234, h: 0.869, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
    s.addShape('ellipse', { x: it.dot, y: 2.322, w: 0.118, h: 0.118, fill: { color: it.color }, line: { type: 'none' } });
    iconBadge(s, { x: it.bx, y: it.low ? 3.484 : 0.513, kind: it.icon, fg: it.color });
  });
  masterChrome(s, 8);
}

/* 9 - October calendar (borderless grid) + two meeting cards. */
function slide09(s) {
  card(s, { x: 1.256, y: 1.189, w: 5.135, h: 5.126, r: 0.212, fill: C.white, shadow: { type: 'outer', blur: 45, offset: 30, angle: 90, color: '000000', opacity: 0.15 } });
  card(s, { x: 1.256, y: 1.171, w: 5.135, h: 0.77, r: 0.231, shape: 'round2SameRect', fill: grad(C.pink) });
  T(s, 'October 2024', { x: 2.439, y: 1.337, w: 2.77, h: 0.438, align: 'center', fontSize: 20, bold: true, color: C.white });

  ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].forEach((d, i) =>
    T(s, d, {
      x: 1.562 + i * 0.6835, y: 2.121, w: 0.431, h: 0.248, align: 'center', valign: 'middle',
      fontSize: 12, color: C.dark, transparency: 20, margin: 0, lineSpacingMultiple: 1.3
    }));

  const days = monthDays(5, 28, 30, 42);
  const marks = { 6: C.pink, 21: C.cyan };
  grid({ x: 1.587, y: 2.61, dx: 0.6835, dy: 0.5985, count: 42 }, (i, cx, cy) => {
    const d = days[i];
    const hit = !d.out && marks[d.n];
    if (hit) card(s, { x: cx, y: cy, w: 0.384, h: 0.418, r: 0.07, fill: grad(hit) });
    T(s, String(d.n), {
      x: cx, y: cy, w: 0.384, h: 0.418, align: 'center', valign: 'middle', margin: 0,
      fontSize: 12, color: hit ? C.white : d.out ? C.silver : C.dark
    });
  });

  T(s, 'Training Agenda Section', {
    x: 6.808, y: 1.146, w: 5.894, h: 1.582, fontSize: 44, color: C.dark, charSpacing: -1.1
  });
  [
    { y: 2.996, day: 'Wednesday', num: '6', label: 'Internal Meeting', color: C.pink },
    { y: 4.758, day: 'Thursday', num: '21', label: 'Monthly Report', color: C.cyan }
  ].forEach(m => {
    card(s, { x: 7.012, y: m.y, w: 4.822, h: 1.523, r: 0.186, fill: C.white, shadow: SH_SOFT() });
    T(s, m.day, { x: 7.156, y: m.y + 0.209, w: 1.525, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.dark });
    T(s, m.num, { x: 7.156, y: m.y + 0.587, w: 1.525, h: 0.707, align: 'center', fontSize: 36, bold: true, color: m.color, charSpacing: -1.5 });
    s.addShape('line', { x: 8.778, y: m.y + 0.289, w: 0, h: 0.945, line: { color: m.color, width: 4.5 } });
    T(s, m.label, { x: 9.121, y: m.y + 0.325, w: 2.435, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    T(s, '10:00 \u2013 11:30', { x: 9.121, y: m.y + 0.806, w: 2.778, h: 0.37, fontSize: 16, color: m.color });
  });
  masterChrome(s, 9);
}

/* 10 - Connected zig-zag of four day tiles. */
function slide10(s) {
  const items = [
    { x: 1.474, y: 1.918, num: '12', dow: 'Friday', color: C.pink, title: 'First Agenda ', ty: 3.681, tx: 1.052 },
    { x: 4.504, y: 3.762, num: '19', dow: 'Friday', color: C.cyan, title: 'Second Agenda', ty: 5.54, tx: 3.968 },
    { x: 7.373, y: 3.768, num: '20', dow: 'Saturday', color: C.yellow, title: 'Third Agenda', ty: 5.531, tx: 6.883, flip: true },
    { x: 10.362, y: 1.918, num: '23', dow: 'Tuesday', color: C.gold, title: 'Fourth Agenda', ty: 3.697, tx: 9.799, flip: true }
  ];
  [
    { x: -0.328, y: 2.677, w: 1.803, h: 0.007, flipV: true },
    { x: 2.992, y: 2.677, w: 1.512, h: 1.843 },
    { x: 6.021, y: 4.521, w: 1.351, h: 0.006 },
    { x: 8.89, y: 2.677, w: 1.472, h: 1.85, flipV: true },
    { x: 11.879, y: 2.677, w: 1.978, h: 0 }
  ].forEach(l => s.addShape('line', { ...l, line: { color: C.silver, width: 2 } }));

  items.forEach(it => {
    card(s, { x: it.x + 0.072, y: it.y + 0.072, w: 1.374, h: 1.374, r: 0.205, fill: grad(it.color), shadow: SH_CARD() });
    card(s, { x: it.x, y: it.y, w: 1.518, h: 1.518, r: 0.226, line: { color: C.silver, width: 2 } });
    /* "flip" tiles show the weekday above the number instead of below */
    const numY = it.y + (it.flip ? 0.59 : 0.302);
    const dowY = it.y + (it.flip ? 0.302 : 0.86);
    T(s, it.num, { x: it.x + 0.178, y: numY, w: 1.161, h: 0.64, align: 'center', fontSize: 32, bold: true, color: C.white });
    T(s, it.dow, { x: it.x + 0.148, y: dowY, w: 1.222, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.white });
    T(s, it.title, { x: it.tx, y: it.ty, w: 2.498, h: 0.37, align: 'center', fontSize: 16, bold: true, color: it.color });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,', {
      x: it.tx, y: it.ty + 0.365, w: 2.498, h: 0.604, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
  });

  T(s, 'April 2024 Agenda', { x: 4.659, y: 1.046, w: 4.016, h: 1.582, align: 'center', fontSize: 44, bold: true, color: C.dark });
  T(s, 'www.yourwebsite.com', { x: 9.634, y: 6.942, w: 2.435, h: 0.303, align: 'right', fontSize: 12, color: C.mute });
  masterChrome(s, 10);
}

/* 11 - Two checklist cards beside a text column. */
function slide11(s) {
  T(s, 'This Month Agenda', { x: 0.76, y: 1.886, w: 4.262, h: 1.582, fontSize: 44, color: C.dark });
  T(s, 'Training Agenda', { x: 0.76, y: 3.709, w: 2.295, h: 0.37, fontSize: 16, bold: true, color: C.pink });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis fringilla lectus sed scelerisque commodo. Praesent imperdiet augue ac neque posuere bibendum. Praesent a gravida dolor. Pellentesque sit amet semper elit', {
    x: 0.76, y: 4.088, w: 3.678, h: 1.657, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });

  [
    { x: 5.444, title: 'Agenda 01', date: '12 February 2024', color: C.pink, icon: 'mail', bx: 8.278 },
    { x: 9.298, title: 'Agenda 02', date: '15 February 2024', color: C.cyan, icon: 'calendar', bx: 12.144 }
  ].forEach(cardo => {
    card(s, { x: cardo.x, y: 1.578, w: 3.18, h: 3.072, r: 0.237, fill: grad(cardo.color), shadow: SH_CARD() });
    T(s, cardo.title, { x: cardo.x + 0.271, y: 1.921, w: 2.104, h: 0.438, fontSize: 20, bold: true, color: C.white });
    T(s, cardo.date, { x: cardo.x + 0.271, y: 2.341, w: 2.104, h: 0.336, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
    [
      { y: 3.195, label: 'Lorem ipsum dolor sit' },
      { y: 3.766, label: 'Consectetur adipiscing ' }
    ].forEach(row => {
      card(s, { x: cardo.x + 0.271, y: row.y, w: 0.395, h: 0.435, r: 0.115, fill: C.white });
      T(s, '\u2713', {
        x: cardo.x + 0.271, y: row.y, w: 0.395, h: 0.435, align: 'center', valign: 'middle',
        fontSize: 16, bold: true, color: C.dark, margin: 0
      });
      T(s, row.label, { x: cardo.x + 0.806, y: row.y + 0.049, w: 2.104, h: 0.303, fontSize: 12, color: C.white });
    });
    iconBadge(s, { x: cardo.bx, y: 1.31, kind: cardo.icon, fg: cardo.color });
    T(s, 'Meeting Agenda', { x: cardo.x + 0.443, y: 4.986, w: 2.295, h: 0.37, align: 'center', fontSize: 16, bold: true, color: cardo.color });
    T(s, LOREM, { x: cardo.x, y: 5.364, w: 3.18, h: 0.869, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3 });
  });
  masterChrome(s, 11);
}

/* 12 - July calendar card + three numbered agenda rows. */
function slide12(s) {
  card(s, { x: 1.275, y: 1.387, w: 4.917, h: 4.745, r: 0.202, fill: C.white, shadow: SH_CARD() });
  card(s, { x: 1.572, y: 1.609, w: 1.693, h: 0.512, r: 0.096, fill: grad(C.pink), shadow: SH_SOFT() });
  T(s, 'July 2024', {
    x: 1.661, y: 1.664, w: 1.514, h: 0.401, align: 'center', valign: 'middle',
    fontSize: 16, bold: true, color: C.white, lineSpacingMultiple: 1.2
  });

  const COLX = [1.452, 2.108, 2.764, 3.42, 4.076, 4.731, 5.387];
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach((d, i) =>
    T(s, d, {
      x: COLX[i], y: 2.29, w: 0.623, h: 0.364, align: 'center', valign: 'middle',
      fontSize: 14, bold: true, color: C.dark, lineSpacingMultiple: 1.2
    }));

  const days = monthDays(1, 31, 28, 35); // July 2024 starts Monday
  const marks = { 9: C.pink, 18: C.cyan, 23: C.yellow };
  grid({ x: COLX[0], y: 2.831, dx: 0.6558, dy: 0.5997, count: 35 }, (i, cx, cy) => {
    const d = days[i];
    const hit = !d.out && marks[d.n];
    const sunday = i % 7 === 0;
    const color = hit || (d.out ? (sunday ? C.pinkSoft : C.silver) : sunday ? C.pinkDeep : C.dark);
    T(s, String(d.n), {
      x: cx, y: cy, w: 0.623, h: 0.355, align: 'center', valign: 'middle',
      fontSize: 14, bold: Boolean(hit), color, lineSpacingMultiple: 1.2
    });
  });

  T(s, 'Agenda This Month', { x: 6.81, y: 1.228, w: 6.329, h: 0.841, fontSize: 44, bold: true, color: C.dark });
  [
    { y: 2.193, num: '09', title: 'Agenda 01', color: C.pink },
    { y: 3.548, num: '18', title: 'Agenda 02', color: C.cyan },
    { y: 4.903, num: '23', title: 'Agenda 03', color: C.yellow }
  ].forEach(r => {
    T(s, [
      { text: r.num, options: { fontSize: 40, bold: true, color: r.color } },
      { text: 'th ', options: { fontSize: 40, bold: true, color: r.color, superscript: true } }
    ], { x: 6.81, y: r.y, w: 1.938, h: 0.911, valign: 'bottom', fontFace: FONT, lineSpacingMultiple: 1.3 });
    T(s, r.title, { x: 8.08, y: r.y + 0.22, w: 4.12, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et', {
      x: 8.08, y: r.y + 0.599, w: 4.12, h: 0.599, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 12);
}

/* 13 - Four wide day cards, mirrored left/right. */
function slide13(s) {
  T(s, 'Our Agenda This Week', { x: 2.753, y: 0.601, w: 7.827, h: 0.841, align: 'center', fontSize: 44, color: C.dark });
  const cards = [
    { x: 0.76, y: 2.685, num: '15', color: C.pink, right: true },
    { x: 0.76, y: 4.587, num: '16', color: C.yellow, right: true },
    { x: 8.787, y: 2.685, num: '17', color: C.cyan, right: false },
    { x: 8.787, y: 4.587, num: '18', color: C.gold, right: false }
  ];
  /* The left pair is a mirror image of the right pair: date block hugs the
     inner edge and all copy is right-aligned. */
  cards.forEach(c => {
    card(s, { x: c.x, y: c.y, w: 3.783, h: 1.653, r: 0.179, fill: grad(c.color), shadow: SH_CARD() });
    const align = c.right ? 'right' : 'left';
    const dateX = c.right ? c.x + 2.571 : c.x + 0.229;
    const bodyX = c.right ? c.x + 0.101 : c.x + 1.248;
    T(s, c.num, { x: dateX, y: c.y + 0.243, w: 0.928, h: 0.774, align, valign: 'middle', fontSize: 40, bold: true, color: C.white });
    T(s, 'Aug 2024', { x: dateX - 0.25, y: c.y + 0.985, w: 1.178, h: 0.344, align, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
    T(s, LOREM, { x: bodyX, y: c.y + 0.309, w: 2.283, h: 1.0, align, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });
  masterChrome(s, 13);
}

/* 14 - Single wide panel: month banner and five weekday columns. */
function slide14(s) {
  T(s, 'This Month Agenda', { x: 1.873, y: 0.797, w: 9.587, h: 0.841, align: 'center', fontSize: 44, color: C.dark });
  card(s, { x: 0.748, y: 2.235, w: 11.812, h: 4.328, r: 0.282, fill: C.white, shadow: SH_SOFT() });
  card(s, { x: 3.79, y: 1.954, w: 5.753, h: 0.838, r: 0.09, fill: grad(C.pink), shadow: SH_CARD() });
  T(s, 'March 2024', { x: 5.528, y: 2.087, w: 4.394, h: 0.572, fontSize: 28, bold: true, color: C.white });

  const cols = [
    { x: 0.962, dow: 'Monday', num: '15', color: C.pink, label: 'Agenda 01' },
    { x: 3.259, dow: 'Tuesday', num: '16', color: C.cyan, label: 'Agenda 02' },
    { x: 5.555, dow: 'Wednesday', num: '17', color: C.yellow, label: 'Agenda 03' },
    { x: 7.852, dow: 'Thursday', num: '18', color: C.gold, label: 'Agenda 04' },
    { x: 10.148, dow: 'Friday', num: '19', color: C.red, label: 'Agenda 05' }
  ];
  [3.21, 5.506, 7.803, 10.099].forEach(x =>
    s.addShape('line', { x, y: 3.343, w: 0, h: 1.525, line: { color: C.mute, width: 0.75 } }));
  cols.forEach(c => {
    T(s, c.dow, { x: c.x + 0.294, y: 3.203, w: 1.61, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.dark });
    T(s, c.num, { x: c.x + 0.294, y: 3.537, w: 1.61, h: 1.111, align: 'center', fontSize: 60, bold: true, color: c.color, charSpacing: -1.1 });
    T(s, c.label, { x: c.x, y: 4.564, w: 2.199, h: 0.37, align: 'center', fontSize: 16, bold: true, color: c.color });
  });

  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis fringilla lectus sed scelerisque commodo. Praesent imperdiet augue ac neque posuere bibendum. Praesent a gravida dolor. Pellentesque sit amet semper elit', {
    x: 1.283, y: 5.418, w: 10.769, h: 0.607, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });
  masterChrome(s, 14);
}

/* 15 - Full-bleed pink panel with a translucent October calendar. */
function slide15(s) {
  gradBand(s, { x: 0, y: 0.601, w: 13.333, h: 6.299, from: C.pink, to: light(C.pink), dir: 'diag' });

  card(s, { x: 0.76, y: 1.839, w: 5.618, h: 4.241, r: 0.235, fill: C.white, shadow: SH_CARD() });
  card(s, { x: 0.95, y: 2.003, w: 5.238, h: 0.478, r: 0.146, shape: 'round2SameRect', fill: grad(C.pink), flipH: true });
  T(s, 'October 2024', { x: 2.159, y: 1.292, w: 2.821, h: 0.505, align: 'center', fontSize: 24, bold: true, color: C.white });
  ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].forEach((d, i) =>
    T(s, d, {
      x: 0.95 + i * 0.755, y: 1.994, w: 0.708, h: 0.512, align: 'center', valign: 'middle',
      fontSize: 12, bold: true, color: C.white, margin: 0
    }));

  const days = monthDays(2, 31, 30, 42); // Oct 2024 starts Tuesday
  const marks = { 7: C.pink, 16: C.cyan };
  grid({ x: 0.95, y: 2.659, dx: 0.755, dy: 0.5432, count: 42 }, (i, cx, cy) => {
    const d = days[i];
    if (d.out) {
      card(s, { x: cx, y: cy, w: 0.708, h: 0.512, r: 0.077, fill: { color: C.hair, transparency: 70 } });
      return;
    }
    const hit = marks[d.n];
    labelCard(s, {
      x: cx, y: cy, w: 0.708, h: 0.512, r: 0.077,
      fill: hit ? grad(hit) : { color: C.white, transparency: 100 },
      text: String(d.n), fontSize: 12, color: hit ? C.white : C.dark
    });
  });

  T(s, 'Our Agenda Schedules', { x: 7.184, y: 1.659, w: 4.277, h: 1.582, fontSize: 44, color: C.white });
  s.addShape('line', { x: 7.298, y: 4.833, w: 4.724, h: 0, line: { color: C.hair, width: 0.75, dashType: 'sysDot' } });
  [
    { y: 3.792, title: 'Agenda 01', icon: 'mail', color: C.pink },
    { y: 5.194, title: 'Agenda 02', icon: 'calendar', color: C.cyan }
  ].forEach(r => {
    iconBadge(s, { x: 7.298, y: r.y, kind: r.icon, fg: r.color });
    T(s, r.title, { x: 8.244, y: r.y, w: 3.217, h: 0.37, fontSize: 16, bold: true, color: C.white });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing.', {
      x: 8.244, y: r.y + 0.379, w: 4.163, h: 0.336, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 15);
}

/* 16 - Split layout: three month cards left, pink arrow panel right. */
function slide16(s) {
  /* the right panel is a rectangle with a left-pointing notch */
  s.addShape('custGeom', {
    x: 6.667, y: -0.012, w: 6.664, h: 7.5,
    points: [
      { x: 6.664, y: 0 }, { x: 0.659, y: 0 }, { x: 0.659, y: 0.891 }, { x: 0, y: 1.2 },
      { x: 0.659, y: 1.51 }, { x: 0.659, y: 7.5 }, { x: 6.664, y: 7.5 }, { close: true }
    ],
    fill: { color: grad(C.pink) }, line: { type: 'none' }, shadow: SH_CARD()
  });

  [
    { y: 0.971, label: 'Agenda ', month: 'September', color: C.white, x: 8.72, w: 3.356 },
    { y: 3.228, label: 'Agenda ', month: 'October', color: C.pinkSoft, x: 8.29, w: 3.034 },
    { y: 5.455, label: 'Agenda', month: 'November', color: C.pinkSoft, x: 8.29, w: 3.356 }
  ].forEach(m => {
    T(s, m.label, { x: m.x, y: m.y, w: m.w, h: 0.303, fontSize: 12, color: m.color });
    T(s, m.month, { x: m.x, y: m.y + 0.337, w: m.w, h: 0.707, fontSize: 36, bold: true, color: m.color });
  });

  T(s, 'Training Agenda', { x: 1.1, y: 0.621, w: 5.202, h: 0.841, fontSize: 44, color: C.dark });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore', {
    x: 1.1, y: 1.474, w: 4.768, h: 0.607, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });
  [
    { y: 2.352, date: '10 September', color: C.pink },
    { y: 3.905, date: '15 October', color: C.cyan },
    { y: 5.458, date: '20 November', color: C.yellow }
  ].forEach(b => {
    card(s, { x: 1.152, y: b.y, w: 4.716, h: 1.433, r: 0.142, fill: grad(b.color), shadow: SH_CARD() });
    T(s, b.date, { x: 1.481, y: b.y + 0.113, w: 2.633, h: 0.505, fontSize: 24, bold: true, color: C.white });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipi scing elit, sed do eiusmod tempor', {
      x: 1.502, y: b.y + 0.62, w: 4.165, h: 0.607, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 16);
}

/* 17 - Pink title banner over four time slots. */
function slide17(s) {
  card(s, { x: 2.593, y: 0.609, w: 8.147, h: 1.334, r: 0.167, fill: grad(C.pink), shadow: SH_CARD() });
  T(s, 'Agenda This Day', { x: 3.04, y: 0.855, w: 7.253, h: 0.841, align: 'center', fontSize: 44, color: C.white });

  const slots = [
    { x: 0.942, head: 'Opening', time: '08-09 AM' },
    { x: 3.965, head: 'Speech', time: '09-10 AM' },
    { x: 6.987, head: 'Main Event', time: '10-11 AM' },
    { x: 10.01, head: 'Closing', time: '11-12 AM' }
  ];
  [3.613, 6.635, 9.658].forEach(x =>
    s.addShape('ellipse', { x, y: 3.269, w: 0.083, h: 0.083, fill: { color: C.silver }, line: { type: 'none' } }));

  slots.forEach(sl => {
    T(s, sl.head, { x: sl.x + 0.278, y: 2.285, w: 1.847, h: 0.401, align: 'center', fontSize: 16, bold: true, color: C.dark, lineSpacingMultiple: 1.2 });
    T(s, sl.time, { x: sl.x, y: 2.698, w: 2.402, h: 0.637, align: 'center', fontSize: 28, bold: true, color: C.pink, lineSpacingMultiple: 1.2 });
    T(s, LOREM, { x: sl.x, y: 3.439, w: 2.402, h: 1.132, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3 });
  });
  masterChrome(s, 17);
}

/* 18 - Four staggered agenda cards, first one filled pink. */
function slide18(s) {
  T(s, 'Our Training Agenda', { x: 3.406, y: 0.855, w: 6.553, h: 0.841, align: 'center', fontSize: 44, bold: true, color: C.dark });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim.', {
    x: 3.406, y: 1.705, w: 6.553, h: 0.596, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });

  const cards = [
    { x: 0.76, y: 2.859, title: 'Agenda 01', date: 'Friday, 09 Feb 2024', solid: true },
    { x: 6.854, y: 2.859, title: 'Agenda 02', date: 'Saturday, 10 Feb 2024', solid: false },
    { x: 2.075, y: 4.801, title: 'Agenda 03', date: 'Tuesday 13 Feb 2024', solid: false },
    { x: 8.198, y: 4.801, title: 'Agenda 04', date: 'Thursday, 15 Feb 2024', solid: false }
  ];
  cards.forEach(c => {
    card(s, {
      x: c.x, y: c.y, w: 4.375, h: 1.622, r: 0.177,
      fill: c.solid ? grad(C.pink) : C.white, shadow: SH_CARD()
    });
    const fg = c.solid ? C.white : C.dark;
    T(s, c.title, { x: c.x + 0.204, y: c.y + 0.181, w: 1.7, h: 0.37, valign: 'bottom', fontSize: 16, bold: true, color: fg });
    T(s, 'Lorem ipsum dolor sit amet, consectetuer', {
      x: c.x + 0.204, y: c.y + 0.492, w: 3.937, h: 0.344, fontSize: 12, color: c.solid ? C.white : C.gray, lineSpacingMultiple: 1.3
    });
    T(s, c.date, { x: c.x + 0.245, y: c.y + 0.971, w: 2.866, h: 0.37, fontSize: 16, color: c.solid ? C.white : C.pink });
    /* circular chevron button in the lower-right corner */
    const bx = c.x + 3.634;
    const by = c.y + 0.943;
    card(s, { x: bx, y: by, w: 0.394, h: 0.394, r: 0.197, fill: c.solid ? C.white : grad(C.pink) });
    chevron(s, bx + 0.183, by + 0.141, 0.068, 0.132, c.solid ? C.pink : C.white);
  });
  masterChrome(s, 18);
}

/* 19 - Pink header band, two gallery cards, two numbered rows. */
function slide19(s) {
  gradBand(s, { x: 0, y: 0, w: 13.324, h: 3.75, from: C.pink, to: light(C.pink), dir: 'diag' });

  T(s, 'Our Training Agenda', { x: 3.406, y: 0.855, w: 6.553, h: 0.841, align: 'center', fontSize: 44, bold: true, color: C.white });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim.', {
    x: 3.406, y: 1.705, w: 6.553, h: 0.596, align: 'center', fontSize: 12, color: C.white, lineSpacingMultiple: 1.3
  });

  [
    { x: 0.76, cx: 0.928, caption: 'Agenda 01 Gallery' },
    { x: 3.885, cx: 4.045, caption: 'Agenda 02 Gallery' }
  ].forEach(g => {
    card(s, { x: g.x, y: 3.387, w: 2.918, h: 3.244, r: 0.204, fill: C.white, shadow: SH_CARD() });
    picturePlaceholder(s, { x: g.x + 0.1, y: 3.5, w: 2.718, h: 2.4 });
    T(s, g.caption, { x: g.cx, y: 5.99, w: 2.583, h: 0.34, align: 'center', fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3 });
  });

  [
    { y: 3.399, num: '01', date: '28 Sep 2024', color: C.cyan },
    { y: 5.322, num: '02', date: '03 Oct 2024', color: C.yellow }
  ].forEach(r => {
    card(s, { x: 7.479, y: r.y, w: 4.352, h: 1.185, r: 0.14, fill: C.white, shadow: SH_CARD() });
    T(s, r.num, { x: 7.734, y: r.y + 0.233, w: 0.913, h: 0.64, fontSize: 32, bold: true, color: C.dark });
    T(s, 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit, sed do', {
      x: 8.534, y: r.y + 0.27, w: 3.18, h: 0.607, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
    labelCard(s, {
      x: 9.942, y: r.y + 1.022, w: 1.646, h: 0.334, r: 0.06, fill: grad(r.color),
      text: r.date, fontSize: 12
    });
  });
  masterChrome(s, 19, { brand: C.pinkSoft });
}

/* 20 - Left photo panel, white sheet, three day tiles. */
function slide20(s) {
  card(s, { x: 1.688, y: 2.017, w: 11.645, h: 3.467, r: 0.311, shape: 'round1Rect', fill: C.white, flipH: true, shadow: { type: 'outer', blur: 100, offset: 21, angle: 90, color: '000000', opacity: 0.15 } });

  T(s, '2024', {
    x: 2.264, y: 1.958, w: 4.954, h: 2.036, fontSize: 115, bold: true,
    color: lum('FFFFFF', 0.65, 0), transparency: 90, charSpacing: -3
  });
  T(s, 'Training Agenda', { x: 2.415, y: 2.512, w: 5.139, h: 0.841, fontSize: 44, color: C.dark });
  T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit', {
    x: 2.343, y: 3.57, w: 4.874, h: 0.64, fontSize: 16, bold: true, color: lum('E7E6E6', 0.85, 0.15)
  });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore', {
    x: 2.343, y: 4.212, w: 4.874, h: 0.607, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
  });
  card(s, { x: 1.688, y: 5.043, w: 2.592, h: 0.44, r: 0.164, shape: 'round1Rect', fill: grad(C.pink) });
  T(s, '7 days \u2013 11 days', { x: 1.935, y: 5.095, w: 2.265, h: 0.344, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });

  [
    { x: 8.037, y: 2.842, num: '18', color: C.cyan },
    { x: 10.007, y: 1.442, num: '15', color: C.pink },
    { x: 10.007, y: 4.283, num: '25', color: C.yellow }
  ].forEach(t => {
    card(s, { x: t.x, y: t.y, w: 1.806, h: 1.806, r: 0.176, fill: grad(t.color), shadow: SH_CARD() });
    T(s, 'August', { x: t.x + 0.081, y: t.y + 0.187, w: 0.817, h: 0.344, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
    T(s, t.num, { x: t.x + 0.802, y: t.y + 0.077, w: 1.0, h: 0.6, fontSize: 44, bold: true, color: C.white });
    T(s, 'Lorem ipsum dolor sit amet', {
      x: t.x + 0.064, y: t.y + 0.995, w: 1.635, h: 0.599, fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.3
    });
  });

  /* the layout's picture frame sits on top and washes out everything left of it */
  picturePlaceholder(s, { x: 0, y: 0, w: 5.653, h: 7.5, label: 'Image placeholder', panel: true });
  masterChrome(s, 20);
}

/* 21 - Grey hero band with three icon-led agenda columns. */
function slide21(s) {
  gradBand(s, { x: 0.76, y: 1.09, w: 11.812, h: 2.66, from: 'F7F7F7', to: 'BEBEBE', steps: 24, dir: 'v' });
  picturePlaceholder(s, { x: 0.76, y: 1.09, w: 11.812, h: 2.66, label: 'Image placeholder' });
  T(s, 'Agenda Plan Today', { x: 3.083, y: 2.052, w: 7.521, h: 0.774, align: 'center', fontSize: 40, color: C.white });

  [
    { x: 0.813, title: 'Agenda Number One', time: '08 PM \u2013 10 PM', color: C.pink, icon: 'calendar' },
    { x: 5.016, title: 'Agenda Number Two', time: '10 AM \u2013 12 AM', color: C.cyan, icon: 'clipboard' },
    { x: 9.219, title: 'Agenda Number Three', time: '08 AM \u2013 12 AM', color: C.yellow, icon: 'mail' }
  ].forEach(col => {
    iconBadge(s, { x: col.x, y: 4.386, kind: col.icon, bg: grad(col.color), fg: C.white, drop: SH_CARD });
    const tx = col.x + 0.768;
    T(s, col.title, { x: tx, y: 4.386, w: 2.586, h: 0.37, fontSize: 16, bold: true, color: C.dark });
    T(s, col.time, { x: tx, y: 4.773, w: 2.534, h: 0.415, fontSize: 16, color: col.color, lineSpacingMultiple: 1.3 });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', {
      x: tx, y: 5.198, w: 2.534, h: 1.132, fontSize: 12, color: C.gray, lineSpacingMultiple: 1.3
    });
  });
  masterChrome(s, 21);
}

/* ================================================================== *
 * Build
 * ================================================================== */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14,
  slide15, slide16, slide17, slide18, slide19, slide20, slide21
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pptx.title = 'Training Agenda';

  BUILDERS.forEach(fn => fn(pptx.addSlide()));
  return pptx;
}

build().writeFile({ fileName: path.join(__dirname, '037baac3-9556-4eea-a680-22b5e863a007_grok_final.pptx') })
  .then(f => console.log('wrote', f))
  .catch(e => { console.error(e); process.exit(1); });
