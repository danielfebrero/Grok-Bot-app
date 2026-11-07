/*
 * "Medicine Product" — medical presentation template, 20 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only.
 *
 * Substitutions vs. the source deck:
 *   - photographs      -> flat grey rectangles labelled "[image]"
 *   - pictogram icons  -> small compositions of native pptxgenjs shapes
 *   - radial gradients -> flat blends of their two stops (every gradient in the
 *                         deck runs between near-identical pastel tints)
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  ink: '262626',      // headline black
  ink2: '0C0C0C',     // darkest headline
  body: '595959',     // body copy
  body2: '3F3F3F',    // secondary body copy
  muted: '7F7F7F',    // captions / page numbers
  white: 'FFFFFF',
  photo: 'F2F2F2',    // image-placeholder grey
  aqua: 'CBEFED',     // theme accent 1
  blue: 'C9E3FC',     // theme accent 2
  lime: 'CCFB9C',     // theme accent 3
  mint: 'A6FBA9',     // theme accent 4
  charcoal: '323137', // chart line / device bezel
  rule: 'A5A5A5',
  grid: 'D9D9D9',
  navy: '2D3847',
  tagBlue: '9ECCF9',
  tagAqua: 'A9E4E1',
  tagLime: '7EF982',
};

/** blend two hex colours (t = 0 -> a, 1 -> b) */
function mix(a, b, t) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

const AQUA_BLUE = mix(C.aqua, C.blue, 0.55);  // accent1 -> accent2 wash
const LIME_MINT = mix(C.lime, C.mint, 0.40);  // accent3 -> accent4 wash
const AQUA_LIME = mix(C.aqua, C.lime, 0.50);  // accent1 -> accent3 wash (mission badges)
const LIME_WASH = mix(C.aqua, C.lime, 0.61);  // same stops, lime-heavy (mission pills)

const F = { head: 'Manrope SemiBold', body: 'Inter' };
const INSET = [5.4, 5.4, 2.7, 2.7];           // lIns/rIns/bIns/tIns, points

/* the two drop shadows the template reuses everywhere */
const CARD_SHADOW = { type: 'outer', angle: 45, blur: 62, offset: 20, color: C.body2, opacity: 0.15 };
const TILE_SHADOW = { type: 'outer', angle: 90, blur: 11, offset: 22, color: '98A7C5', opacity: 0.24 };

/* ------------------------------------------------------------------ helpers */
/** body text box using the deck's Google-Slides insets */
function txt(s, text, o) {
  s.addText(text, Object.assign(
    { fontFace: F.body, fontSize: 11, color: C.body, margin: INSET, valign: 'top', align: 'left' },
    o));
}
/** headline text box */
function head(s, text, o) {
  txt(s, text, Object.assign({ fontFace: F.head, fontSize: 41, color: C.ink }, o));
}
/** rounded rectangle card (`r` = corner radius, inches; `shadow` = soft drop shadow) */
function card(s, x, y, w, h, fill, r, shadow) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: r === undefined ? 0.16 : r, fill: { color: fill },
    shadow: shadow ? CARD_SHADOW : undefined,
  });
}
/** soft decorative square that fades into the background */
function glow(s, x, y, size, color, transparency) {
  s.addShape('roundRect', {
    x, y, w: size, h: size, rectRadius: size * 0.16667,
    fill: { color: color || C.lime, transparency: transparency === undefined ? 65 : transparency },
  });
}
function glows(s, list, color, transparency) {
  list.forEach(([x, y, size]) => glow(s, x, y, size, color, transparency));
}
/** placeholder standing in for a photograph in the original deck */
function photo(s, x, y, w, h, r) {
  s.addShape(r ? 'roundRect' : 'rect', { x, y, w, h, rectRadius: r || 0, fill: { color: C.photo } });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle', margin: 0,
    fontFace: F.body, fontSize: Math.max(8, Math.min(12, Math.round(w * 3))), color: '8A8A8A',
  });
}

/* ---------------------------------------------------------------- pictograms
 * The source deck uses custom-path icons. Each is rebuilt from native shapes.
 * Parts are [shape, x, y, w, h, rotate, 'ko'] in tile-relative units (0..1);
 * 'ko' knocks the part out in the tile's own fill colour.
 */
const ICONS = {
  tooth: [['ellipse', 0.17, 0.08, 0.66, 0.58], ['rect', 0.26, 0.40, 0.16, 0.50], ['rect', 0.58, 0.40, 0.16, 0.50]],
  chatPlus: [['ellipse', 0.36, 0.34, 0.58, 0.58], ['wedgeRoundRectCallout', 0.04, 0.06, 0.62, 0.60],
    ['plus', 0.20, 0.20, 0.30, 0.30, 0, 'ko']],
  dropPlus: [['teardrop', 0.04, 0.10, 0.52, 0.52, 225], ['teardrop', 0.38, 0.26, 0.58, 0.58, 225],
    ['plus', 0.56, 0.48, 0.24, 0.24, 0, 'ko']],
  heartPulse: [['heart', 0.08, 0.12, 0.84, 0.74]],
  bandage: [['roundRect', 0.10, 0.36, 0.80, 0.28, 45], ['roundRect', 0.10, 0.36, 0.80, 0.28, 315]],
  heartPlus: [['heart', 0.06, 0.10, 0.88, 0.76], ['plus', 0.54, 0.42, 0.30, 0.30, 0, 'ko']],
  flask: [['ellipse', 0.06, 0.10, 0.34, 0.34], ['trapezoid', 0.00, 0.52, 0.52, 0.34],
    ['triangle', 0.42, 0.30, 0.50, 0.58]],
  hospital: [['rect', 0.16, 0.14, 0.68, 0.76], ['plus', 0.36, 0.22, 0.28, 0.26, 0, 'ko'],
    ['rect', 0.28, 0.58, 0.12, 0.26, 0, 'ko'], ['rect', 0.60, 0.58, 0.12, 0.26, 0, 'ko']],
  smiley: [['smileyFace', 0.10, 0.10, 0.80, 0.80]],
  monitor: [['rect', 0.08, 0.16, 0.84, 0.56], ['rect', 0.18, 0.26, 0.64, 0.36, 0, 'ko'],
    ['rect', 0.44, 0.72, 0.12, 0.14]],
  doc: [['rect', 0.20, 0.08, 0.60, 0.84], ['rect', 0.30, 0.24, 0.40, 0.08, 0, 'ko'],
    ['rect', 0.30, 0.42, 0.40, 0.08, 0, 'ko'], ['rect', 0.30, 0.60, 0.26, 0.08, 0, 'ko']],
  home: [['triangle', 0.04, 0.10, 0.92, 0.44], ['rect', 0.22, 0.46, 0.56, 0.44]],
  phone: [['moon', 0.16, 0.08, 0.62, 0.84, 315]],
  mail: [['rect', 0.06, 0.28, 0.88, 0.44], ['triangle', 0.06, 0.24, 0.88, 0.30, 180]],
  globe: [['donut', 0.08, 0.08, 0.84, 0.84], ['rect', 0.08, 0.44, 0.84, 0.12]],
  instagram: [['frame', 0.10, 0.10, 0.80, 0.80], ['ellipse', 0.32, 0.32, 0.36, 0.36]],
  twitter: [['teardrop', 0.10, 0.16, 0.74, 0.70, 225]],
  youtube: [['roundRect', 0.04, 0.24, 0.92, 0.52], ['triangle', 0.40, 0.34, 0.26, 0.32, 90, 'ko']],
  check: [['rect', 0.16, 0.50, 0.30, 0.12, 45], ['rect', 0.34, 0.30, 0.52, 0.12, 315]],
  pill: [['flowChartTerminator', 0.10, 0.34, 0.80, 0.32, 315]],
  calendar: [['rect', 0.12, 0.20, 0.76, 0.70], ['rect', 0.26, 0.06, 0.09, 0.20], ['rect', 0.66, 0.06, 0.09, 0.20],
    ['rect', 0.22, 0.46, 0.16, 0.14, 0, 'ko'], ['rect', 0.46, 0.46, 0.16, 0.14, 0, 'ko'],
    ['rect', 0.22, 0.68, 0.16, 0.14, 0, 'ko']],
  person: [['ellipse', 0.34, 0.10, 0.32, 0.32], ['trapezoid', 0.14, 0.50, 0.72, 0.36]],
  plus: [['plus', 0.16, 0.16, 0.68, 0.68]],
};

/** draw a pictogram inside the box (x, y, w, h) */
function icon(s, name, x, y, w, h, inkColor, tileColor) {
  ICONS[name].forEach(([shape, rx, ry, rw, rh, rotate, ko]) => {
    s.addShape(shape, {
      x: x + rx * w, y: y + ry * h, w: rw * w, h: rh * h, rotate: rotate || 0,
      fill: { color: ko ? tileColor : inkColor },
    });
  });
}

/** rounded icon tile: coloured square + centred pictogram (~60% of the tile) */
function tile(s, x, y, w, h, radius, fill, name, inkColor, shadow) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: radius, fill: { color: fill }, shadow: shadow ? CARD_SHADOW : undefined,
  });
  const g = Math.min(w, h) * 0.62;
  icon(s, name, x + (w - g) / 2, y + (h - g) / 2, g, g, inkColor || C.ink, fill);
}

/** master furniture: brand tag top-right + page number bottom-right */
function chrome(s, page, onDark) {
  txt(s, 'Medicine Product', {
    x: 8.258, y: 0.26, w: 1.414, h: 0.189, align: 'right',
    fontFace: F.head, fontSize: 8, color: onDark ? C.white : C.tagBlue, lineSpacingMultiple: 0.9,
  });
  txt(s, String(page), {
    x: 8.779, y: 5.226, w: 0.975, h: 0.215, align: 'right', fontSize: 8,
    color: onDark ? C.white : C.muted,
  });
}

/* ==================================================================== slides */

/* 1 — title (layout hides the master furniture) */
function slide01(p) {
  const s = p.addSlide();
  s.addShape('ellipse', { x: 4.111, y: -0.159, w: 5.06, h: 5.06, fill: { color: C.aqua, transparency: 82 } });
  card(s, 0.406, 0.344, 9.188, 4.969, AQUA_BLUE, 0.164);
  card(s, 0.756, 0.669, 4.494, 2.144, C.white, 0.148);
  txt(s, 'Medicine Product', { x: 1.164, y: 0.945, w: 3.013, h: 1.59, fontFace: F.head, fontSize: 45, color: C.ink });
  s.addShape('plus', { x: 4.111, y: 1.409, w: 0.691, h: 0.664, fill: { color: LIME_MINT } });
  txt(s, 'Medical Presentation Template', {
    x: 0.756, y: 4.159, w: 3.837, h: 0.353, fontSize: 15, bold: true, color: C.body2, lineSpacingMultiple: 1.2,
  });
  glow(s, 5.765, 1.018, 1.445, C.white, 35);
  glow(s, 8.014, 2.536, 1.194, C.white, 35);
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
    x: 0.756, y: 4.508, w: 3.599, h: 0.482, lineSpacingMultiple: 1.2,
  });
  photo(s, 4.021, 0.469, 5.573, 4.844);
}

/* 2 — two-value intro */
function slide02(p) {
  const s = p.addSlide();
  chrome(s, 2);
  photo(s, 0.841, 0.785, 5.492, 1.85, 0.308);
  photo(s, 6.528, 0.785, 2.636, 1.85, 0.308);
  head(s, 'Where do ideas come from?', { x: 0.788, y: 2.911, w: 3.679, h: 2.121 });
  [
    { x: 4.949, tx: 5.047, fill: AQUA_BLUE, title: 'Value one', icon: 'tooth' },
    { x: 7.353, tx: 7.437, fill: LIME_MINT, title: 'Value two', icon: 'chatPlus' },
  ].forEach((c) => {
    tile(s, c.tx, 3.053, 0.622, 0.623, 0.104, c.fill, c.icon);
    txt(s, c.title, {
      x: c.x, y: 3.822, w: 1.751, h: 0.33, fontFace: F.head, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.2,
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: c.x, y: 4.186, w: 1.954, h: 0.694, lineSpacingMultiple: 1.2,
    });
  });
}

/* 3 — quote band */
function slide03(p) {
  const s = p.addSlide();
  chrome(s, 3);
  card(s, 1.25, 2.594, 7.875, 2.22, AQUA_BLUE, 0.37);
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. '
    + 'Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.', {
    x: 2.425, y: 3.685, w: 6.421, h: 0.694, lineSpacingMultiple: 1.2,
  });
  head(s, 'How Treat \nOur Patient', { x: 5.538, y: 0.975, w: 3.587, h: 1.439, color: C.ink2 });
  ['heartPulse', 'bandage', 'heartPlus'].forEach((name, i) => {
    tile(s, 6.094 + i * 0.832, 2.958, 0.586, 0.553, 0.098, C.white, name);
  });
  txt(s, '\u201C', { x: 1.761, y: 3.318, w: 0.928, h: 1.672, fontSize: 86, color: C.ink, lineSpacingMultiple: 1.2 });
  glows(s, [[5.112, -0.17, 0.777], [-0.358, 4.208, 1.061], [9.182, 2.312, 0.429]]);
  photo(s, 0.703, 1.061, 4.455, 2.22, 0.37);
}

/* 4 — event details */
function slide04(p) {
  const s = p.addSlide();
  chrome(s, 4);
  card(s, 0.896, 3.308, 8.208, 1.403, AQUA_BLUE, 0.234);
  head(s, 'Health Test & Medication Check', {
    x: 0.847, y: 0.771, w: 4.153, h: 2.121, color: C.ink2, lineSpacingMultiple: 1,
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. '
    + 'Aenean massa. Cum sociis natoque penatibus et magnis', {
    x: 5.388, y: 1.006, w: 3.992, h: 1.101, fontSize: 12, color: C.body2, lineSpacingMultiple: 1.3,
  });
  txt(s, '-Bring to the table win-win survival strategies to ensure proactive domination.-', {
    x: 5.388, y: 2.271, w: 3.676, h: 0.48, fontSize: 12, italic: true,
  });
  [
    { ix: 1.483, x: 2.191, w: 0.861, w2: 0.994, icon: 'calendar', title: 'Date', value: '18 May 2020' },
    { ix: 4.046, x: 4.716, w: 1.037, w2: 1.283, icon: 'pill', title: 'Location', value: 'Donec aliquet' },
    { ix: 6.640, x: 7.311, w: 1.206, w2: 0.973, icon: 'person', title: 'Registration', value: 'Free' },
  ].forEach((f) => {
    tile(s, f.ix, 3.704, 0.586, 0.553, 0.098, C.white, f.icon);
    txt(s, f.title, {
      x: f.x, y: 3.664, w: f.w, h: 0.315, fontFace: F.head, fontSize: 12, bold: true, color: C.ink,
      lineSpacingMultiple: 1.3,
    });
    txt(s, f.value, { x: f.x, y: 3.98, w: f.w2, h: 0.259, fontSize: 9, color: C.body2, lineSpacingMultiple: 1.3 });
  });
  glows(s, [[5.074, -0.258, 0.777], [8.221, 5.128, 0.883], [-0.253, 1.841, 0.604]]);
}

/* 5 — three photo cards */
function slide05(p) {
  const s = p.addSlide();
  chrome(s, 5);
  glow(s, 0.537, 4.462, 0.777, C.aqua, 88);
  glow(s, 2.752, 4.713, 0.142, C.aqua, 76);
  const cols = [
    { x: 1.182, tx: 1.381, bx: 1.381, gx: 1.457, fill: C.white, n: '01', nf: AQUA_BLUE, nc: C.ink, tw: 2.004, title: 'The Inflammation Connection ' },
    { x: 3.905, tx: 4.104, bx: 4.121, gx: 4.180, fill: AQUA_BLUE, n: '02', nf: C.white, nc: C.ink, tw: 2.202, title: 'Lifestyle Factors Can Play a Role ' },
    { x: 6.629, tx: 6.828, bx: 6.828, gx: 6.904, fill: C.white, n: '03', nf: LIME_MINT, nc: C.white, tw: 2.402, title: 'A Word From\nTherapist' },
  ];
  cols.forEach((c) => card(s, c.x, 0.89, 2.402, 3.572, c.fill, 0.126, true));
  cols.forEach((c) => photo(s, c.x, 0.889, 2.402, 1.67));
  glow(s, 8.249, -0.352, 2.207);
  cols.forEach((c) => {
    txt(s, c.title, { x: c.tx, y: 2.946, w: c.tw, h: 0.53, fontFace: F.head, fontSize: 14, color: C.ink });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', {
      x: c.bx, y: 3.54, w: 1.94, h: 0.564, fontSize: 9, color: C.body2, lineSpacingMultiple: 1.1,
    });
  });
  glows(s, [[0.759, 4.756, 1.088], [0.406, 4.538, 0.618], [2.496, 4.664, 0.142]]);
  cols.forEach((c) => {
    s.addShape('roundRect', { x: c.gx, y: 2.238, w: 0.522, h: 0.522, rectRadius: 0.087, fill: { color: c.nf } });
    txt(s, c.n, {
      x: c.gx, y: 2.238, w: 0.522, h: 0.522, align: 'center', valign: 'middle',
      fontFace: F.head, fontSize: 12, color: c.nc,
    });
  });
}

/* 6 — leadership cards */
function slide06(p) {
  const s = p.addSlide();
  chrome(s, 6);
  glows(s, [[5.367, 0.223, 0.777], [3.998, 5.165, 1.369], [-0.337, 1.0, 0.933]], C.aqua, 85);
  glow(s, 0.215, 0.874, 0.251, C.aqua, 76);
  photo(s, 6.477, 0.862, 2.628, 3.9);
  card(s, 3.388, 2.573, 2.338, 1.903, AQUA_BLUE, 0.139, true);
  card(s, 5.880, 2.759, 2.338, 1.903, C.white, 0.139, true);
  card(s, 0.896, 2.759, 2.338, 1.903, C.white, 0.139, true);
  head(s, 'Corporate Leadership Motion', {
    x: 0.896, y: 0.795, w: 4.996, h: 1.303, valign: 'bottom', lineSpacingMultiple: 0.9,
  });
  const cols = [
    { x: 1.083, y: 3.543, tw: 1.974, title: 'Anxiety', bx: 1.158, by: 3.014, bf: AQUA_BLUE, n: '01' },
    { x: 3.614, y: 3.368, tw: 1.974, title: 'Bipolar Disorder', bx: 3.693, by: 2.855, bf: C.white, n: '02' },
    { x: 6.111, y: 3.543, tw: 2.124, title: 'Panic Disorder', bx: 6.144, by: 3.014, bf: LIME_MINT, n: '01' },
  ];
  cols.forEach((c) => {
    txt(s, c.title, {
      x: c.x, y: c.y, w: c.tw, h: 0.345, valign: 'bottom', fontFace: F.head, fontSize: 14, color: C.ink,
      lineSpacingMultiple: 1.3,
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer', {
      x: c.x, y: c.y + 0.347, w: 1.974, h: 0.514, color: C.body2, lineSpacingMultiple: 1.3,
    });
  });
  glows(s, [[5.367, 0.223, 0.777], [3.998, 5.165, 1.369], [-0.337, 1.0, 0.933]]);
  cols.forEach((c) => {
    s.addShape('roundRect', { x: c.bx, y: c.by, w: 0.437, h: 0.437, rectRadius: 0.073, fill: { color: c.bf } });
    txt(s, c.n, {
      x: c.bx, y: c.by, w: 0.437, h: 0.437, align: 'center', valign: 'middle',
      fontFace: F.head, fontSize: 12, color: C.ink,
    });
  });
}

/* 7 — agenda (layout has a full-bleed gradient background) */
function slide07(p) {
  const s = p.addSlide();
  s.background = { color: AQUA_BLUE };
  chrome(s, 7, true);
  photo(s, 3.479, 0.917, 5.914, 2.479, 0.32);
  glow(s, 7.295, 4.550, 0.777, C.aqua, 88);
  glow(s, 7.017, 4.708, 1.369, C.white, 35);
  glow(s, 8.386, 4.384, 0.142, C.white, 35);
  glow(s, 8.956, -0.081, 0.933, C.aqua, 88);
  head(s, 'Agenda\nConsul', { x: 0.871, y: 1.073, w: 2.129, h: 1.303, valign: 'bottom', lineSpacingMultiple: 0.9 });
  glow(s, 9.385, 4.677, 0.142, C.aqua, 76);
  glow(s, 8.072, -0.206, 0.933, C.white, 35);
  glow(s, -0.249, 4.747, 1.119, C.white, 35);
  card(s, 2.612, 2.883, 4.740, 1.429, C.white, 0.238, true);
  card(s, 0.975, 2.887, 1.369, 1.436, C.white, 0.183, true);
  s.addShape('round2SameRect', { x: 0.975, y: 2.887, w: 1.369, h: 0.343, fill: { color: AQUA_BLUE } });
  txt(s, 'Wednesday', {
    x: 0.975, y: 2.947, w: 1.369, h: 0.257, align: 'center', fontFace: F.head, fontSize: 12, color: C.ink,
    lineSpacingMultiple: 0.9,
  });
  txt(s, '26', { x: 0.975, y: 3.181, w: 1.369, h: 0.833, align: 'center', fontFace: F.head, fontSize: 45, color: C.ink });
  txt(s, 'June  2023', { x: 0.975, y: 3.944, w: 1.369, h: 0.219, align: 'center', fontFace: F.head, color: C.body2 });
  txt(s, 'Schedule Meet Therapist', { x: 2.916, y: 3.166, w: 3.076, h: 0.303, fontFace: F.head, fontSize: 14, color: C.ink });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa', {
    x: 2.916, y: 3.562, w: 4.254, h: 0.453, lineSpacingMultiple: 1.1,
  });
}

/* 8 — anxiety facts */
function slide08(p) {
  const s = p.addSlide();
  chrome(s, 8);
  txt(s, 'Is Anxiety a Mental Illness?', {
    x: 0.875, y: 0.622, w: 8.25, h: 0.682, fontFace: F.head, fontSize: 36, color: C.ink,
  });
  glows(s, [[0.418, 3.347, 1.135], [7.996, -0.318, 1.135], [3.133, 5.106, 0.718]]);
  glow(s, 8.842, 0.675, 0.283, C.aqua, 76);
  [
    { y: 1.745, h: 0.577, w: 2.475, fill: AQUA_BLUE, icon: 'smiley', text: 'Written by Mental Health Experts and Journalists' },
    { y: 2.867, h: 0.315, w: 2.747, fill: LIME_MINT, icon: 'monitor', text: 'Depression, bipolar disorder' },
    { y: 3.727, h: 0.577, w: 2.852, fill: AQUA_BLUE, icon: 'doc', text: 'Fact-checked with Science-backed  and Research' },
  ].forEach((r) => {
    txt(s, r.text, { x: 6.835, y: r.y, w: r.w, h: r.h, fontFace: F.head, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 });
    tile(s, 6.031, r.y + 0.012, 0.586, 0.553, 0.098, r.fill, r.icon, C.ink, true);
  });
  photo(s, 0.875, 1.703, 2.219, 2.695, 0.28);
  photo(s, 3.453, 1.703, 2.219, 2.695, 0.28);
}

/* 9 — feature specialties */
function slide09(p) {
  const s = p.addSlide();
  chrome(s, 9);
  head(s, 'Our feature specialty', { x: 1.503, y: 0.54, w: 6.994, h: 0.757, align: 'center' });
  const specialty = (c, y, ty, by, fill) => {
    tile(s, c.x, y, 0.976, 0.978, 0.278, fill, c.icon, C.ink, true);
    txt(s, c.title, { x: c.tx, y: ty, w: 1.953, h: 0.278, align: 'center', fontFace: F.head, fontSize: 12, color: C.ink });
    txt(s, 'Lorem ipsum dolor sit amet, ', {
      x: c.bx, y: by, w: c.bw, h: 0.256, align: 'center', fontSize: 9, lineSpacingMultiple: 1.3,
    });
  };
  [
    { x: 1.979, tx: 1.490, bx: 1.504, bw: 1.927, title: 'Dentistry', icon: 'tooth' },
    { x: 4.513, tx: 4.024, bx: 3.860, bw: 2.282, title: 'General diagnosis', icon: 'chatPlus' },
    { x: 7.047, tx: 6.558, bx: 6.571, bw: 1.927, title: 'Neuro surgery', icon: 'dropPlus' },
  ].forEach((c) => specialty(c, 1.599, 2.656, 2.912, AQUA_BLUE));
  [
    { x: 3.199, tx: 2.711, bx: 2.724, bw: 1.927, title: 'Pharmacy', icon: 'flask' },
    { x: 5.824, tx: 5.336, bx: 5.349, bw: 1.927, title: 'Emergency medical', icon: 'hospital' },
  ].forEach((c) => specialty(c, 3.501, 4.564, 4.819, LIME_MINT));
  glows(s, [[-0.161, 0.697, 1.135], [9.091, 1.599, 1.135], [1.061, 4.002, 0.718]]);
}

/* 10 — facilities (layout has a full-bleed gradient background) */
function slide10(p) {
  const s = p.addSlide();
  s.background = { color: AQUA_BLUE };
  chrome(s, 10, true);
  head(s, 'Your Partner in Healthy Living', { x: 0.766, y: 1.411, w: 2.92, h: 2.802, color: C.ink2 });
  card(s, 4.195, 1.206, 1.850, 3.212, C.white, 0.18, true);
  card(s, 6.206, 1.206, 2.958, 1.508, C.white, 0.161, true);
  card(s, 6.206, 2.910, 2.958, 1.508, C.white, 0.218, true);
  txt(s, 'Facility One', {
    x: 4.321, y: 2.502, w: 1.598, h: 0.278, align: 'center', fontFace: F.head, fontSize: 12, color: C.body2,
  });
  txt(s, 'Lorem ipsum \ndolor sit amet, consectetuer', {
    x: 4.372, y: 2.85, w: 1.495, h: 0.694, align: 'center', lineSpacingMultiple: 1.2,
  });
  tile(s, 4.881, 1.868, 0.477, 0.479, 0.136, AQUA_BLUE, 'tooth', C.ink, true);
  [
    { y: 1.501, ty: 1.759, fill: LIME_MINT, icon: 'dropPlus', title: 'Facility Two' },
    { y: 3.179, ty: 3.477, fill: AQUA_BLUE, icon: 'chatPlus', title: 'Facility Three' },
  ].forEach((r) => {
    txt(s, r.title, { x: 7.045, y: r.y, w: 1.598, h: 0.278, fontFace: F.head, fontSize: 12, color: C.body2 });
    tile(s, 6.449, r.y - 0.013, 0.477, 0.479, 0.136, r.fill, r.icon, C.ink, true);
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: 7.045, y: r.ty, w: 1.876, h: 0.694, lineSpacingMultiple: 1.2,
    });
  });
  glows(s, [[2.364, -0.231, 1.135], [9.464, 2.192, 1.135], [2.931, 4.479, 0.718]], C.white, 35);
}

/* 11 — problems */
function slide11(p) {
  const s = p.addSlide();
  chrome(s, 11);
  photo(s, 0, 0, 10, 3.062);
  head(s, 'Take Care of the Product', { x: 0.783, y: 3.453, w: 4.107, h: 1.439 });
  glow(s, 4.602, 5.110, 0.718);
  [
    { x: 5.321, n: '01', color: C.tagAqua },
    { x: 6.708, n: '02', color: C.tagBlue },
    { x: 8.095, n: '03', color: C.tagLime },
  ].forEach((c) => {
    card(s, c.x, 2.507, 1.324, 0.872, C.white, 0.106, true);
    txt(s, 'Problem', {
      x: c.x + 0.113, y: 2.569, w: 1.14, h: 0.315, fontFace: F.head, fontSize: 12, color: c.color,
      lineSpacingMultiple: 1.3,
    });
    txt(s, c.n, { x: c.x + 0.113, y: 2.777, w: 1.014, h: 0.581, fontFace: F.head, fontSize: 30, color: C.body2 });
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore '
    + 'et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut '
    + 'aliquip ex ea commodo consequat. ', {
    x: 5.321, y: 3.610, w: 4.098, h: 1.202, lineSpacingMultiple: 1.3,
  });
  glows(s, [[1.858, -0.294, 1.213], [1.467, 0.543, 0.2]]);
}

/* 12 — our mission */
function slide12(p) {
  const s = p.addSlide();
  chrome(s, 12);
  const cols = [
    { x: 1.288, cardFill: C.white, pillFill: LIME_WASH, circleFill: AQUA_LIME, label: 'Mission 01', n: '01', bodyColor: C.body, py: 2.939, cy: 2.162 },
    { x: 3.985, cardFill: AQUA_BLUE, pillFill: C.white, circleFill: C.white, label: 'Mission 02', n: '02', bodyColor: C.body2, py: 2.897, cy: 2.160 },
    { x: 6.682, cardFill: C.white, pillFill: LIME_WASH, circleFill: AQUA_LIME, label: 'Mission 03', n: '03', bodyColor: C.body, py: 2.939, cy: 2.160 },
  ];
  cols.forEach((c) => {
    card(s, c.x, 1.914, 2.03, 3.062, c.cardFill, 0.175, true);
    s.addShape('roundRect', { x: c.x - 0.133, y: c.py, w: 2.297, h: 0.552, rectRadius: 0.106, fill: { color: c.pillFill } });
    txt(s, c.label, {
      x: c.x, y: c.py + 0.086, w: 2.03, h: 0.379, align: 'center', fontFace: F.head, fontSize: 18, bold: true, color: C.ink,
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor', {
      x: c.x + 0.148, y: 3.715, w: 1.734, h: 0.846, align: 'center', fontSize: 9, color: c.bodyColor,
      lineSpacingMultiple: 1.3,
    });
  });
  cols.forEach((c) => {
    const cx = c.x + 0.731;
    s.addShape('ellipse', {
      x: cx, y: c.cy, w: 0.568, h: 0.568, fill: { color: c.circleFill },
      shadow: { type: 'outer', angle: 45, blur: 45, offset: 42, color: '000000', opacity: 0.15 },
    });
    txt(s, c.n, {
      x: cx, y: c.cy, w: 0.568, h: 0.568, align: 'center', valign: 'middle',
      fontFace: F.head, fontSize: 12, bold: true, color: C.ink,
    });
  });
  txt(s, 'Our Mission', {
    x: 1.31, y: 0.444, w: 7.381, h: 0.631, align: 'center', fontFace: F.head, fontSize: 33, bold: true, color: C.ink2,
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. '
    + 'Fusce posuere, magna', {
    x: 2.075, y: 1.073, w: 5.85, h: 0.519, align: 'center', lineSpacingMultiple: 1.3,
  });
  glows(s, [[1.05, 0.658, 0.519], [9.245, 1.151, 1.135], [-0.123, 3.715, 0.718]]);
}

/* 13 — profile / timeline */
function slide13(p) {
  const s = p.addSlide();
  chrome(s, 13);
  glow(s, 1.011, 4.330, 0.718);
  photo(s, 1.372, 0, 3.439, 5.625);
  tile(s, 0.944, 1.581, 0.785, 0.784, 0.181, AQUA_BLUE, 'heartPulse');
  txt(s, 'Medic volunteer', { x: 5.24, y: 0.834, w: 2.241, h: 0.252, color: C.muted });
  txt(s, 'John Dew', { x: 5.24, y: 1.113, w: 3.581, h: 0.757, fontFace: F.head, fontSize: 41, color: C.ink2 });
  s.addShape('line', { x: 8.217, y: 1.505, w: 1.783, h: 0, line: { color: 'D8D8D8', width: 1 } });
  txt(s, 'Job Tittle Here', { x: 5.24, y: 1.807, w: 2.056, h: 0.345, fontSize: 12, color: C.muted, lineSpacingMultiple: 1.5 });
  [
    { year: '2024', yy: 2.506, ly: 2.590, by: 2.930, w: 0.964 },
    { year: '2025', yy: 3.632, ly: 3.733, by: 4.072, w: 1.037 },
  ].forEach((r) => {
    txt(s, r.year, { x: 5.24, y: r.yy, w: r.w, h: 0.483, fontFace: F.head, fontSize: 18, color: C.ink, lineSpacingMultiple: 1.5 });
    txt(s, 'Work Experience', {
      x: 6.038, y: r.ly, w: 1.861, h: 0.346, fontFace: F.head, fontSize: 12, color: C.body2, lineSpacingMultiple: 1.5,
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', {
      x: 6.038, y: r.by, w: 3.304, h: 0.513, lineSpacingMultiple: 1.3,
    });
  });
  glows(s, [[-0.259, 0.697, 0.519], [9.342, 1.782, 0.84]]);
}

/* 14 — team */
function slide14(p) {
  const s = p.addSlide();
  chrome(s, 14);
  photo(s, 4.844, 2.480, 1.895, 2.352, 0.2);
  photo(s, 6.990, 0.664, 1.895, 2.352, 0.2);
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing ', {
    x: 4.653, y: 1.598, w: 2.171, h: 0.513, align: 'right', lineSpacingMultiple: 1.3,
  });
  txt(s, 'James Smith', {
    x: 4.972, y: 1.029, w: 1.852, h: 0.328, align: 'right', fontFace: F.head, fontSize: 14, color: C.ink,
    lineSpacingMultiple: 1.2,
  });
  txt(s, 'Job Position', {
    x: 4.972, y: 1.288, w: 1.852, h: 0.255, align: 'right', fontSize: 9, italic: true, lineSpacingMultiple: 1.3,
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing ', {
    x: 6.967, y: 4.160, w: 2.194, h: 0.513, lineSpacingMultiple: 1.3,
  });
  txt(s, 'Violet Sims', {
    x: 6.967, y: 3.558, w: 1.852, h: 0.328, fontFace: F.head, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.2,
  });
  txt(s, 'Job Position', { x: 6.967, y: 3.833, w: 1.852, h: 0.255, fontSize: 9, italic: true, lineSpacingMultiple: 1.3 });
  head(s, 'Corporate Legacy in the Making', { x: 0.944, y: 0.904, w: 3.459, h: 2.121 });
  s.addShape('line', { x: 1.045, y: 3.222, w: 0.737, h: 0, line: { color: C.rule, width: 1.5 } });
  txt(s, 'What would be your perfect weekend?', {
    x: 0.944, y: 3.479, w: 3.005, h: 0.638, fontFace: F.head, fontSize: 14, lineSpacingMultiple: 1.3,
  });
  glows(s, [[1.522, -0.077, 0.519], [9.325, 3.181, 0.84], [0.406, 4.594, 0.718]]);
  [[8.726, 1.024], [4.648, 2.840]].forEach(([x, y]) => {
    ['instagram', 'twitter', 'youtube'].forEach((name, i) => tile(s, x, y + i * 0.607, 0.416, 0.416, 0.069, AQUA_BLUE, name, C.ink, true));
  });
}

/* 15 — positioning matrix (layout has a full-bleed gradient background) */
function slide15(p) {
  const s = p.addSlide();
  s.background = { color: AQUA_BLUE };
  chrome(s, 15, true);
  glows(s, [[0.406, -0.4, 1.135], [8.951, 3.303, 0.577], [-0.284, 4.26, 0.718]], C.white, 35);
  [[4.946, C.lime], [3.572, C.blue], [2.259, C.aqua]].forEach(([x, color]) => {
    s.addShape('line', { x, y: 1.619, w: 0, h: 2.999, line: { color, width: 1 } });
  });
  s.addShape('line', { x: 1.91, y: 4.619, w: 6.573, h: 0, line: { color: 'BFBFBF', width: 1.25, dashType: 'dash' } });
  s.addShape('line', { x: 1.91, y: 1.432, w: 0, h: 3.187, line: { color: 'BFBFBF', width: 1.25, dashType: 'dash' } });
  txt(s, 'Horizontal value', { x: 4.322, y: 4.69, w: 1.75, h: 0.303, align: 'center', fontSize: 14 });
  txt(s, 'Vertical value', { x: 0.769, y: 2.874, w: 1.75, h: 0.303, align: 'center', fontSize: 14, rotate: 270 });
  [
    { x: 4.889, y: 1.704, tx: 5.271, ty: 1.790, tw: 2.475, text: 'Written by Mental Health Experts and Journalists' },
    { x: 3.546, y: 2.600, tx: 3.824, ty: 2.703, tw: 2.747, text: 'Shaping the Future of Medicine Together' },
    { x: 2.234, y: 3.494, tx: 2.420, ty: 3.598, tw: 2.852, text: 'Fact-checked with Science-backed  and Research' },
  ].forEach((b) => {
    card(s, b.x, b.y, 3.182, 0.812, C.white, 0.135, true);
    txt(s, b.text, {
      x: b.tx, y: b.ty, w: b.tw, h: 0.58, fontFace: F.head, fontSize: 12, color: C.body2, lineSpacingMultiple: 1.3,
    });
  });
  head(s, 'What is Life Medicine', { x: 1.744, y: 0.485, w: 6.512, h: 0.757, align: 'center' });
}

/* 16 — three feature tables */
function slide16(p) {
  const s = p.addSlide();
  chrome(s, 16);
  glows(s, [[-0.259, 0.910, 0.519], [9.293, 1.170, 0.840], [0.476, 4.459, 0.718]]);
  const ROWS = ['10 Book Medical', 'Installed CMS', '100 Infographic Sets', '5 Hospital',
    '100 Infographic Sets', '5 Medicine', '5 Patient'];
  [
    { x: 1.350, rx: 1.581, check: C.aqua },
    { x: 3.914, rx: 4.146, check: C.blue },
    { x: 6.479, rx: 6.710, check: C.lime },
  ].forEach((col) => {
    s.addShape('roundRect', { x: col.x, y: 2.018, w: 2.172, h: 0.516, rectRadius: 0.258, fill: { color: AQUA_BLUE } });
    txt(s, 'Table', {
      x: col.x, y: 2.018, w: 2.172, h: 0.516, align: 'center', valign: 'middle',
      fontFace: F.head, fontSize: 15, bold: true, color: C.ink,
    });
    ROWS.forEach((label, i) => {
      const y = 2.698 + i * 0.328;
      icon(s, 'check', col.rx, y + 0.055, 0.145, 0.145, col.check);
      txt(s, label, { x: col.rx + 0.188, y, w: 2.2, h: 0.253 });
    });
  });
  head(s, 'Three Table Medicine', { x: 1.744, y: 0.485, w: 6.512, h: 0.757, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ', {
    x: 1.533, y: 1.269, w: 6.934, h: 0.252, align: 'center',
  });
}

/* 17 — combo bar + line chart */
function slide17(p) {
  const s = p.addSlide();
  chrome(s, 17);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Des'];
  const bars = [4.3, 2.5, 3.5, 4.5, 2.7, 6, 4.3, 2.5, 3.5, 4.5, 2.7, 6];
  const dark = [4.3, 2.5, 3.5, 4.5, 2.7, 6, 4.3, 2.5, 3.5, 4.5, 2.7, 6];
  const green = [2.3, 1.2, 3, 5, 2, 4, 2.3, 1.2, 3, 5, 2, 4];
  s.addChart([
    {
      type: p.ChartType.bar,
      data: [{ name: 'Volume', labels: months, values: bars }],
      options: { chartColors: [C.aqua], barGapWidthPct: 155 },
    },
    {
      type: p.ChartType.line,
      data: [{ name: '2025', labels: months, values: dark }, { name: '2024', labels: months, values: green }],
      options: {
        chartColors: [C.charcoal, C.lime], lineSize: 3, lineDataSymbol: 'circle',
        lineDataSymbolSize: 7, lineDataSymbolLineSize: 2.25, lineSmooth: false,
      },
    },
  ], {
    x: 0.980, y: 2.097, w: 8.190, h: 2.364,
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: F.body, catAxisLabelFontSize: 9, catAxisLabelColor: C.body,
    valAxisLabelFontFace: F.body, valAxisLabelFontSize: 9, valAxisLabelColor: C.body,
    valAxisMinVal: 0, valAxisMaxVal: 7, valAxisMajorUnit: 1,
    valGridLine: { color: C.grid, size: 0.75 }, catGridLine: { style: 'none' },
    valAxisLineShow: false, catAxisLineColor: C.grid,
  });
  // the deck mirrors the value scale down the right-hand edge
  for (let v = 0; v <= 7; v += 1) {
    txt(s, String(v), { x: 9.003, y: 4.032 - v * 0.2805, w: 0.321, h: 0.227, align: 'right', fontSize: 9 });
  }
  card(s, 6.797, 2.165, 0.959, 0.413, C.white, 0.158, true);
  txt(s, '2025', {
    x: 6.797, y: 2.165, w: 0.959, h: 0.413, align: 'center', valign: 'middle',
    fontFace: F.head, fontSize: 12, color: C.ink2,
  });
  glow(s, 0.373, 0.390, 0.777, C.aqua, 88);
  glow(s, -0.487, 0.779, 1.369);
  glow(s, 9.612, 1.342, 0.142, C.aqua, 76);
  glow(s, 9.441, 1.645, 0.933);
  head(s, 'Medical Condition Result', { x: 0.859, y: 0.573, w: 8.282, h: 0.757, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce ', {
    x: 1.199, y: 1.289, w: 7.603, h: 0.283, align: 'center', lineSpacingMultiple: 1.3,
  });
}

/* 18 — pricing */
function slide18(p) {
  const s = p.addSlide();
  chrome(s, 18);
  [
    { y: 0.686, fill: AQUA_BLUE, price: '42', color: C.ink, name: 'VIP Plan', nameSize: 18, ny: 0.856, by: 1.224, bodyColor: C.body2, monthColor: C.ink, monthFace: F.head },
    { y: 2.186, fill: C.white, price: '25', color: C.tagAqua, name: 'Basic Plan', nameSize: 14, ny: 2.422, by: 2.725, bodyColor: C.body, monthColor: '283848', monthFace: 'IBM Plex Sans' },
    { y: 3.687, fill: C.white, price: '57', color: '99F73A', name: 'Exclusive Plan ', nameSize: 14, ny: 3.923, by: 4.250, bodyColor: C.body, monthColor: '283848', monthFace: 'IBM Plex Sans' },
  ].forEach((pl) => {
    card(s, 1.065, pl.y, 4.091, 1.252, pl.fill, 0.159, true);
    txt(s, [
      { text: '$', options: { fontSize: 30 } },
      { text: pl.price, options: { fontSize: 30 } },
      { text: '.00', options: { fontSize: 15 } },
    ], {
      x: 1.302, y: pl.y + 0.287, w: 1.181, h: 0.53, fontFace: F.head, bold: true, color: pl.color,
      lineSpacingMultiple: 0.9,
    });
    txt(s, '/Month', {
      x: 1.324, y: pl.y + 0.738, w: 1.137, h: 0.227, align: 'center', fontFace: pl.monthFace, fontSize: 9,
      color: pl.monthColor,
    });
    txt(s, pl.name, {
      x: 2.605, y: pl.ny, w: 2.125, h: pl.nameSize === 18 ? 0.379 : 0.303,
      valign: pl.nameSize === 18 ? 'top' : 'middle',
      fontFace: F.head, fontSize: pl.nameSize, bold: true, color: C.ink,
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', {
      x: 2.605, y: pl.by, w: 2.239, h: 0.453, color: pl.bodyColor, lineSpacingMultiple: 1.1,
    });
  });
  head(s, 'Changing Lives Through Science', { x: 5.801, y: 0.908, w: 3.454, h: 2.802 });
  s.addShape('line', { x: 5.926, y: 3.878, w: 0.737, h: 0, line: { color: C.rule, width: 1.5 } });
  txt(s, 'What would be your perfect weekend?', {
    x: 5.801, y: 4.077, w: 3.005, h: 0.638, fontFace: F.head, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.3,
  });
  glow(s, 9.194, 0.342, 0.933, C.aqua, 88);
  glows(s, [[-0.259, 0.910, 0.519], [9.479, 1.118, 0.840], [7.448, 5.198, 0.718]]);
}

/* 19 — phone mockups */
function slide19(p) {
  const s = p.addSlide();
  chrome(s, 19);
  glow(s, 9.479, 1.118, 0.840);
  glow(s, 4.711, 4.107, 0.718);
  head(s, 'Where do Patient Come From?', { x: 0.764, y: 0.618, w: 3.812, h: 2.121 });
  glow(s, -0.259, 1.058, 0.519);
  [
    { title: 'Schedule Vaccine?', ty: 3.048, by: 3.407, bw: 2.556, tx: 1.275, cx: 0.873, cy: 3.143, cs: 0.264 },
    { title: 'Health consultation', ty: 4.144, by: 4.503, bw: 2.759, tx: 1.294, cx: 0.842, cy: 4.253, cs: 0.326 },
  ].forEach((r) => {
    txt(s, r.title, {
      x: r.tx, y: r.ty, w: 3.036, h: 0.344, fontFace: F.head, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.3,
    });
    tile(s, r.cx, r.cy, r.cs, r.cs, r.cs * 0.32, AQUA_BLUE, 'check', C.white);
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: r.tx, y: r.by, w: r.bw, h: 0.515, color: C.navy, lineSpacingMultiple: 1.3,
    });
  });
  // device mock-ups: bezel outline wrapping a screen placeholder
  [[4.940, 0.480], [7.120, 1.033]].forEach(([bx, by]) => {
    s.addShape('roundRect', {
      x: bx, y: by, w: 1.858, h: 3.988, rectRadius: 0.22, fill: { color: C.photo },
      line: { color: C.charcoal, width: 4 },
    });
    txt(s, '[image]', {
      x: bx, y: by, w: 1.858, h: 3.988, align: 'center', valign: 'middle', margin: 0,
      fontSize: 8, color: '8A8A8A',
    });
  });
}

/* 20 — thank you (layout hides the master furniture) */
function slide20(p) {
  const s = p.addSlide();
  s.addShape('rect', { x: 6.038, y: 0, w: 3.962, h: 5.625, fill: { color: AQUA_BLUE } });
  photo(s, 0.937, 2.128, 9.063, 2.187);
  glow(s, 2.769, 4.127, 0.718);
  txt(s, 'Thank You!', { x: 0.773, y: 0.72, w: 5.006, h: 1.085, fontFace: F.head, fontSize: 60, color: C.ink });
  glow(s, -0.259, 0.858, 0.519);
  glow(s, 7.925, -0.408, 1.135, C.white, 35);
  card(s, 6.475, 1.705, 2.752, 1.910, C.white, 0.132, true);
  [
    { y: 1.957, name: 'home', text: 'Company address 1234, A' },
    { y: 2.331, name: 'phone', text: '+0782 4022 XXXX XXX' },
    { y: 2.691, name: 'mail', text: 'mail@companysite.com' },
    { y: 3.087, name: 'globe', text: 'www.companysite.com' },
  ].forEach((r) => {
    txt(s, r.text, { x: 6.978, y: r.y, w: 2.367, h: 0.272, color: C.body2, lineSpacingMultiple: 1.2 });
    icon(s, r.name, 6.758, r.y + 0.048, 0.146, 0.146, AQUA_BLUE, C.white);
  });
}

/* ------------------------------------------------------------------- build */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.author = 'pptxgenjs';
pptx.title = 'Medicine Product';
pptx.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
pptx.layout = 'DECK';
BUILDERS.forEach((build) => build(pptx));
pptx.writeFile({ fileName: path.join(__dirname, '0847c9f5-f283-4e4f-84c8-ef782a5c354e_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
