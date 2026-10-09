/**
 * Agriculture Presentation Template - 30 slides, 13.333 x 7.5 in
 * Rebuilt with pptxgenjs only. Raster photos in the source deck are replaced
 * with labelled placeholder rectangles; empty picture placeholders (which render
 * blank in the source) are left blank.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GREEN = '129B54';
const LIME = '7CBC11';
const GOLD = 'F3C332';
const WHITE = 'FFFFFF';
const DARK = '404040'; // headings
const BODY = '595959'; // body copy
const MUTED = 'A6A6A6';
const LIGHT = 'F2F2F2'; // card / panel grey
const RULE = 'ADACAC'; // dashed rules
const SILVER = 'D9D9D9'; // empty progress track
const STAR = 'FFC000';
const PHOTO = 'F0F0F0'; // image placeholder body

const HEAD = 'Lato Black';
const SANS = 'Montserrat';
const ICONF = 'DejaVu Sans'; // font that carries the pictogram glyphs

/* Gradients used throughout the deck (OOXML angle, degrees clockwise from +x) */
const G_BUTTON = { stops: [[0.20, LIME], [1.00, GOLD]], ang: 315 };
const G_CARD = { stops: [[0.35, LIME], [1.00, GOLD]], ang: 315 };
const G_TILE = { stops: [[0.20, LIME], [1.00, GOLD]], ang: 135 };
const G_BADGE = { stops: [[0.30, LIME], [1.00, GOLD]], ang: 225 };
const G_CIRCLE = { stops: [[0.00, LIME], [1.00, GOLD]], ang: 270 };
const G_DISC = { stops: [[0.10, LIME], [1.00, GOLD]], ang: 315 };
const G_DOT = { stops: [[0.10, LIME], [1.00, GOLD]], ang: 90 };
const G_BAR = { stops: [[0.00, LIME], [1.00, GOLD]], ang: 45 };
// full-bleed hero washes: translucent grey fading into green over the white page
const G_HERO = { stops: [[0.00, 'A6A6A6', 0.56], [0.90, GREEN, 1]], ang: 90 };
const G_HERO_SOFT = { stops: [[0.00, 'A6A6A6', 0.80], [0.90, GREEN, 0.70]], ang: 90 };
const G_HERO_DIAG = { stops: [[0.00, 'A6A6A6', 0.45], [0.90, GREEN, 0.95]], ang: 135 };

/* Pictograms - single characters drawn on top of the gradient chips */
const ICON = {
  check: '\u2713', wheat: '\u2698', plant: '\u273F', farmer: '\u263A',
  barn: '\u2302', tractor: '\u2699', drone: '\u2708', bulb: '\u2600',
  pencil: '\u270E', download: '\u2B07', monitor: '\u25A3', camera: '\u25C9',
  sound: '\u266B', play: '\u25B6', home: '\u2302', mail: '\u2709', phone: '\u260E'
};

/* Boilerplate copy. Every long paragraph in the deck is a prefix of this. */
const LOREM = 'Lorem Ipsum\u00A0is simply dummy text of the printing and typesetting industry. '
  + 'Lorem Ipsum has been the industry\u2019s standard dummy text ever since the 1500s, when an '
  + 'unknown printer took a galley of type, Lorem Ipsum\u00A0is simply dummy text of the printing. '
  + 'Lorem Ipsum\u00A0is simply dummy text of the printing and typesetting industry. Lorem Ipsum '
  + 'has been the industry\u2019s standard';
const lorem = n => LOREM.slice(0, n);
const POREM = 'Porem ipsum dolor sit amet, cot adipicing elit. ';

/* ------------------------------------------------------------------ helpers */
const mix = (a, b, t) => [0, 2, 4]
  .map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t))
  .map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();

const flatten = (c, a) => (a === undefined || a >= 1) ? c : mix('FFFFFF', c, a);

/**
 * Colour of a multi-stop gradient at fraction t (0..1). A stop is
 * [position, hex] or [position, hex, alpha]; alpha is flattened over white.
 */
function stopColor(stops, t) {
  const first = stops[0], last = stops[stops.length - 1];
  if (t <= first[0]) return flatten(first[1], first[2]);
  if (t >= last[0]) return flatten(last[1], last[2]);
  for (let i = 1; i < stops.length; i++) {
    const [p0, c0, a0] = stops[i - 1], [p1, c1, a1] = stops[i];
    if (t > p1) continue;
    const k = (t - p0) / (p1 - p0);
    // PowerPoint interpolates colour and opacity independently, then composites
    return flatten(mix(c0, c1, k), (a0 === undefined ? 1 : a0) * (1 - k) + (a1 === undefined ? 1 : a1) * k);
  }
  return flatten(last[1], last[2]);
}
const midColor = g => stopColor(g.stops, 0.5);

/* --- gradient painting ------------------------------------------------------
 * pptxgenjs has no gradient fill, so a gradient is drawn as a fan of flat
 * polygons: take the shape outline, slice it with lines perpendicular to the
 * gradient direction, and fill each slice with the colour at that position.
 * ------------------------------------------------------------------------- */

/** Outline polygons (arrays of [x, y] inch pairs), all convex. */
const polyRect = (x, y, w, h) => [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];

function polyRound(x, y, w, h, r) {
  const seg = 6, pts = [];
  const corners = [[x + w - r, y + h - r, 0], [x + r, y + h - r, 90], [x + r, y + r, 180], [x + w - r, y + r, 270]];
  corners.forEach(([cx, cy, a0]) => {
    for (let i = 0; i <= seg; i++) {
      const a = (a0 + (i / seg) * 90) * Math.PI / 180;
      pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
    }
  });
  return pts;
}

function polyEllipse(x, y, w, h) {
  const seg = 64, rx = w / 2, ry = h / 2, pts = [];
  for (let i = 0; i < seg; i++) {
    const a = (i / seg) * 2 * Math.PI;
    pts.push([x + rx + rx * Math.cos(a), y + ry + ry * Math.sin(a)]);
  }
  return pts;
}

/** Clip a convex polygon against the half-plane  X*nx + Y*ny <= limit. */
function clipHalf(poly, nx, ny, limit) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i], b = poly[(i + 1) % poly.length];
    const da = a[0] * nx + a[1] * ny - limit, db = b[0] * nx + b[1] * ny - limit;
    if (da <= 0) out.push(a);
    if ((da < 0 && db > 0) || (da > 0 && db < 0)) {
      const k = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * k, a[1] + (b[1] - a[1]) * k]);
    }
  }
  return out;
}

/**
 * Fill `poly` with gradient `g`, one flat polygon per colour step. PowerPoint
 * ramps a gradient across the shape's bounding box, so `bounds` (defaulting to
 * the outline itself) is what sets the 0..1 range.
 */
function gradPoly(s, g, poly, steps, bounds) {
  const nx = Math.cos(g.ang * Math.PI / 180), ny = Math.sin(g.ang * Math.PI / 180);
  const proj = (bounds || poly).map(p => p[0] * nx + p[1] * ny);
  const lo = Math.min.apply(null, proj), hi = Math.max.apply(null, proj);
  const n = steps || Math.max(10, Math.min(96, Math.round((hi - lo) * 34)));
  const cell = (hi - lo) / n, bleed = cell * 0.4;
  for (let i = 0; i < n; i++) {
    let slice = clipHalf(poly, nx, ny, i === n - 1 ? Infinity : lo + (i + 1) * cell + bleed);
    slice = clipHalf(slice, -nx, -ny, -(lo + i * cell - bleed));
    if (slice.length < 3) continue;
    const xs = slice.map(p => p[0]), ys = slice.map(p => p[1]);
    const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
    const pw = Math.max.apply(null, xs) - x0, ph = Math.max.apply(null, ys) - y0;
    if (pw < 0.004 || ph < 0.004) continue;
    s.addShape('custGeom', {
      x: x0, y: y0, w: pw, h: ph,
      points: slice.map(p => ({ x: +(p[0] - x0).toFixed(4), y: +(p[1] - y0).toFixed(4) })).concat([{ close: true }]),
      fill: { color: stopColor(g.stops, (i + 0.5) / n) }, line: { type: 'none' }
    });
  }
}

/** Gradient-filled rectangle. */
const gradRect = (s, o) => gradPoly(s, o.grad, polyRect(o.x, o.y, o.w, o.h), o.steps);

/** Gradient-filled rounded rectangle (optionally outlined / shadowed). */
function gradRoundRect(s, o) {
  const r = o.rad || 0;
  if (o.shadow) s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: r,
    fill: { color: midColor(o.grad) }, line: { type: 'none' }, shadow: o.shadow
  });
  gradPoly(s, o.grad, polyRound(o.x, o.y, o.w, o.h, r), o.steps, polyRect(o.x, o.y, o.w, o.h));
  if (o.line) s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: r, fill: { type: 'none' }, line: o.line
  });
}

/** Gradient-filled ellipse (optionally outlined / shadowed). */
function gradOval(s, o) {
  if (o.shadow) s.addShape('ellipse', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: midColor(o.grad) }, line: { type: 'none' }, shadow: o.shadow
  });
  gradPoly(s, o.grad, polyEllipse(o.x, o.y, o.w, o.h), o.steps, polyRect(o.x, o.y, o.w, o.h));
  if (o.line) s.addShape('ellipse', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: { type: 'none' }, line: o.line
  });
}

/** Plain text box; the source deck anchors text boxes to the top. */
function txt(s, str, o) {
  s.addText(str, Object.assign({ fontFace: SANS, valign: 'top', color: BODY, isTextBox: true }, o));
}

const eyebrow = (s, x, y, color) =>
  txt(s, 'Agriculture Template', { x, y, w: 2.602, h: 0.269, fontSize: 10, color: color || GREEN });

const eyebrowC = (s, x, y, color) =>
  txt(s, 'Agriculture Template', { x, y, w: 2.602, h: 0.269, fontSize: 10, color: color || GREEN, align: 'center' });

/** Big 32pt Lato Black heading; `lines` may hold one or more paragraphs. */
function heading(s, lines, o) {
  const runs = [].concat(lines).map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } }));
  s.addText(runs, Object.assign({
    fontFace: HEAD, fontSize: 32, bold: true, color: DARK, valign: 'top', isTextBox: true
  }, o));
}

/** Justified 10pt body paragraph with 1.5 line spacing. */
function para(s, str, o) {
  txt(s, str, Object.assign({ fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 }, o));
}

/** Small bold caption above a body paragraph. */
const caption = (s, str, o) =>
  txt(s, str, Object.assign({ fontSize: 12, bold: true, color: DARK }, o));

/** Pill shaped gradient "More Info" call to action. */
function moreInfo(s, x, y) {
  gradRoundRect(s, { x, y, w: 1.534, h: 0.289, rad: 0.0586, grad: G_BUTTON, shadow: SHADOW_BUTTON });
  txt(s, 'More Info', {
    x, y, w: 1.534, h: 0.289, fontSize: 10, color: WHITE, align: 'center',
    valign: 'middle', charSpacing: 3
  });
}

/* Drop shadows, transcribed from the source (blur/offset in points, angle in degrees). */
const SHADOW_HALO = { type: 'outer', color: '000000', opacity: 0.20, blur: 25, offset: 0, angle: 90 };
const SHADOW_CARD = { type: 'outer', color: '000000', opacity: 0.20, blur: 25, offset: 20, angle: 30 };
const SHADOW_BUTTON = { type: 'outer', color: '000000', opacity: 0.25, blur: 45, offset: 0, angle: 90 };
const SHADOW_TILE = { type: 'outer', color: '000000', opacity: 0.40, blur: 15, offset: 0, angle: 90 };
const SHADOW_CHIP = { type: 'outer', color: '000000', opacity: 0.20, blur: 20, offset: 5, angle: 80 };
const SHADOW_STAT = { type: 'outer', color: '000000', opacity: 0.25, blur: 15, offset: 10, angle: 45 };
const SHADOW_QUOTE = { type: 'outer', color: '000000', opacity: 0.08, blur: 26, offset: 26, angle: 68 };
const SHADOW_ROUND = { type: 'outer', color: '000000', opacity: 0.43, blur: 15, offset: 0, angle: 50 };

/** White pictogram centred on a chip. */
const glyph = (s, ch, x, y, w, h, size, color) =>
  txt(s, ch, {
    x, y, w, h, fontSize: size, color: color || WHITE, fontFace: ICONF,
    align: 'center', valign: 'middle'
  });

/** Row of three small gradient social buttons (twitter / facebook / pinterest). */
function socialRow(s, x, y, w) {
  const d = 0.234, gap = (w - d) / 2;
  ['t', 'f', 'p'].forEach((ch, i) => {
    gradOval(s, { x: x + i * gap, y, w: d, h: d, grad: G_CIRCLE });
    txt(s, ch, {
      x: x + i * gap, y, w: d, h: d, fontSize: 8, bold: true, color: WHITE,
      align: 'center', valign: 'middle'
    });
  });
}

/**
 * Stand-in for a raster photograph. The source images are product mockups on a
 * transparent background, so the panel is translucent and lets artwork behind
 * it show through, as the real photo does.
 */
function photo(s, x, y, w, h, label) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.06, fill: { color: PHOTO, transparency: 70 },
    line: { color: 'D0D0D0', width: 1, dashType: 'dash' }
  });
  txt(s, label || '[image]', {
    x, y, w, h, fontSize: 12, color: 'A6A6A6', align: 'center', valign: 'middle'
  });
}

/** Free-form outline; `pts` are inches relative to the shape origin. */
const freeform = (s, o) =>
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: o.points,
    fill: { color: o.color }, line: { type: 'none' }
  });

/* ------------------------------------------------------------- slide bodies */

function slide01(s) { // hero / cover
  gradRect(s, { x: 0, y: 0, w: 13.333, h: 7.5, grad: G_HERO, steps: 70 });
  txt(s, 'Agriculture', {
    x: 1.747, y: 2.033, w: 9.84, h: 1.717, fontFace: HEAD, fontSize: 96, bold: true,
    color: WHITE, align: 'center', charSpacing: 3
  });
  txt(s, 'Agriculture Presentation Template', {
    x: 3.53, y: 3.91, w: 6.273, h: 0.337, fontSize: 14, color: WHITE, align: 'center', charSpacing: 3
  });
  moreInfo(s, 5.899, 4.765);
}

function slide02(s) { // welcome
  s.addShape('rect', { x: 0, y: 0, w: 3.939, h: 4.485, fill: { color: GREEN }, line: { type: 'none' } });
  eyebrow(s, 7.026, 1.707);
  heading(s, ['Welcome To', 'Agriculture Slide'], { x: 7.026, y: 2.066, w: 5.227, h: 1.178 });
  para(s, lorem(369), { x: 7.026, y: 3.559, w: 5.031, h: 1.616 });
  moreInfo(s, 7.164, 5.641);
}

function slide03(s) { // about, with green info card
  s.addShape('roundRect', {
    x: 1.217, y: 4.354, w: 7.767, h: 2.467, rectRadius: 0.218,
    fill: { color: GREEN }, line: { type: 'none' }, shadow: SHADOW_CARD
  });
  eyebrow(s, 1.217, 1.188);
  heading(s, ['About Agriculture', 'Company'], { x: 1.217, y: 1.548, w: 5.227, h: 1.178 });
  para(s, lorem(227), { x: 1.217, y: 3.013, w: 6.192, h: 0.858 });
  caption(s, 'Description Agriculture', { x: 4.358, y: 4.878, w: 3.147, h: 0.303, color: WHITE });
  para(s, lorem(141), { x: 4.361, y: 5.25, w: 3.145, h: 1.111, color: WHITE });
}

function slide04(s) { // our history
  eyebrow(s, 1.177, 0.858);
  heading(s, ['Our History', 'Agriculture Slide'], { x: 1.177, y: 1.218, w: 4.989, h: 1.178 });
  moreInfo(s, 1.285, 2.822);
  caption(s, 'Description Agriculture', { x: 6.861, y: 1.294, w: 2.836, h: 0.303, color: BODY });
  para(s, lorem(250), { x: 6.864, y: 1.711, w: 5.031, h: 1.111 });
}

function slide05(s) { // our vision
  eyebrow(s, 1.133, 1.434);
  heading(s, ['Our Vision', 'Agriculture Slide'], { x: 1.133, y: 1.793, w: 4.989, h: 1.178 });
  caption(s, 'Description Vision', { x: 1.133, y: 3.348, w: 2.836, h: 0.303, color: BODY });
  para(s, lorem(250), { x: 1.136, y: 3.766, w: 5.031, h: 1.111 });
  para(s, lorem(181), { x: 1.133, y: 4.948, w: 5.031, h: 0.858 });
  s.addShape('rect', { x: 6.868, y: 0, w: 0.646, h: 2.971, fill: { color: GREEN }, line: { type: 'none' } });
}

function slide06(s) { // 2 x 2 mission cards
  const CARDS = [
    { x: 5.847, y: 3.026, title: 'Our Mission One', grad: true },
    { x: 9.335, y: 3.026, title: 'Our Mission Two' },
    { x: 5.847, y: 4.903, title: 'Our Mission Three' },
    { x: 9.335, y: 4.903, title: 'Our Mission Four' }
  ];
  CARDS.forEach(c => {
    const box = { x: c.x, y: c.y, w: 3.228, h: 1.654, rad: 0.092 };
    const line = { color: WHITE, width: 2.25 };
    if (c.grad) gradRoundRect(s, Object.assign({ grad: G_CARD, line, shadow: SHADOW_HALO }, box));
    else s.addShape('roundRect', {
      x: box.x, y: box.y, w: box.w, h: box.h, rectRadius: box.rad,
      fill: { color: GREEN }, line, shadow: SHADOW_HALO
    });
    caption(s, c.title, { x: c.x + 0.288, y: c.y + 0.223, w: 2.204, h: 0.303, color: WHITE });
    para(s, lorem(86), { x: c.x + 0.288, y: c.y + 0.58, w: 2.652, h: 0.858, color: WHITE });
  });
  eyebrow(s, 6.135, 1.092);
  heading(s, ['Our Mission', 'Agriculture Slide'], { x: 6.135, y: 1.452, w: 4.989, h: 1.178 });
}

function slide07(s) { // facilities checklist
  s.addShape('rect', { x: 9.939, y: 0, w: 3.394, h: 7.5, fill: { color: GREEN }, line: { type: 'none' } });
  [3.339, 4.973].forEach(y => {
    gradRoundRect(s, { x: 1.223, y, w: 0.591, h: 0.591, rad: 0.099, grad: G_BADGE, shadow: SHADOW_HALO });
    glyph(s, ICON.check, 1.223, y, 0.591, 0.591, 22);
    caption(s, 'Description Agriculture', { x: 2.104, y, w: 2.836, h: 0.303, color: BODY });
    para(s, lorem(151), { x: 2.106, y: y + 0.411, w: 4.324, h: 0.858 });
  });
  eyebrow(s, 1.047, 1.261);
  heading(s, ['Our Company', 'Agriculture Facilities'], { x: 1.047, y: 1.621, w: 5.62, h: 1.178 });
}

function slide08(s) { // two KPI cards on a green band
  s.addShape('rect', { x: 0, y: 5.323, w: 13.333, h: 2.177, fill: { color: GREEN }, line: { type: 'none' } });
  const KPI = [
    { x: 4.598, value: '4M+', icon: ICON.wheat, iconX: 4.795 },
    { x: 8.803, value: '550K', icon: ICON.plant, iconX: 9.048 }
  ];
  KPI.forEach(k => {
    s.addShape('roundRect', {
      x: k.x, y: 4.553, w: 3.74, h: 1.575, rectRadius: 0.13,
      fill: { color: WHITE }, line: { type: 'none' }, shadow: SHADOW_HALO
    });
    txt(s, k.value, {
      x: k.x + 1.035, y: 4.748, w: 1.669, h: 0.505, fontSize: 24, bold: true, color: '000000'
    });
    gradRoundRect(s, { x: k.x, y: 4.553, w: 0.787, h: 0.787, rad: 0.145, grad: G_BADGE, shadow: SHADOW_HALO });
    glyph(s, k.icon, k.x, 4.553, 0.787, 0.787, 24);
    txt(s, lorem(48), { x: k.x + 1.043, y: 5.299, w: 2.374, h: 0.601, fontSize: 10, lineSpacingMultiple: 1.5 });
  });
  eyebrow(s, 6.682, 1.448);
  heading(s, ['We Are Provide', 'Agriculture Slide'], { x: 6.682, y: 1.807, w: 5.303, h: 1.178 });
  para(s, lorem(190), { x: 6.682, y: 3.294, w: 5.409, h: 0.858 });
}

function slide09(s) { // text left, gradient callout right
  eyebrow(s, 1.194, 1.564);
  heading(s, ['We Are Provide', 'Agriculture Slide'], { x: 1.194, y: 1.924, w: 4.981, h: 1.178 });
  para(s, lorem(325), { x: 1.194, y: 3.416, w: 4.52, h: 1.616 });
  para(s, lorem(151), { x: 1.194, y: 5.127, w: 4.52, h: 0.858 });
  gradRoundRect(s, {
    x: 9.823, y: 4.921, w: 2.864, h: 1.731, rad: 0.251, grad: G_CARD,
    line: { color: WHITE, width: 2.25 }, shadow: SHADOW_HALO
  });
  glyph(s, ICON.plant, 10.973, 5.202, 0.554, 0.422, 20);
  txt(s, lorem(48), {
    x: 10.084, y: 5.768, w: 2.333, h: 0.606, fontSize: 10, color: WHITE,
    align: 'center', lineSpacingMultiple: 1.5
  });
}

function slide10(s) { // three service cards under a gradient header
  gradRect(s, { x: 0, y: 0, w: 13.333, h: 4.524, grad: G_HERO_DIAG, steps: 80 });
  const SERVICES = [
    { x: 0.76, title: 'Services One', icon: ICON.farmer, iconX: 1.349, iconY: 4.268, iconW: 0.365 },
    { x: 4.823, title: 'Services Two', icon: ICON.barn, iconX: 5.405, iconY: 4.261, iconW: 0.38 },
    { x: 8.882, title: 'Services Three', icon: ICON.plant, iconX: 9.472, iconY: 4.298, iconW: 0.372 }
  ];
  SERVICES.forEach(c => {
    s.addShape('roundRect', {
      x: c.x, y: 3.75, w: 3.688, h: 2.948, rectRadius: 0.211, fill: { color: WHITE },
      line: { color: LIME, width: 2.25 }, shadow: SHADOW_CARD
    });
    caption(s, c.title, { x: c.x + 0.387, y: 5.084, w: 2.349, h: 0.337, fontSize: 14 });
    gradRoundRect(s, { x: c.x + 0.419, y: 4.128, w: 0.709, h: 0.669, rad: 0.112, grad: G_TILE, shadow: SHADOW_TILE });
    glyph(s, c.icon, c.iconX - 0.1, c.iconY - 0.06, c.iconW + 0.2, c.iconW + 0.12, 20);
    para(s, lorem(95), { x: c.x + 0.387, y: 5.482, w: 2.875, h: 0.858 });
  });
  eyebrow(s, 1.179, 1.016, WHITE);
  heading(s, 'Our Best Services', { x: 1.179, y: 1.376, w: 5.568, h: 0.64, color: WHITE });
  para(s, lorem(227), { x: 1.179, y: 2.294, w: 6.38, h: 0.858, color: WHITE });
}

function slide11(s) { // green sidebar with two circular icons
  s.addShape('rect', { x: 0, y: 0, w: 3.475, h: 7.5, fill: { color: GREEN }, line: { type: 'none' } });
  [{ y: 1.128, icon: ICON.tractor }, { y: 4.177, icon: ICON.drone }].forEach(b => {
    gradOval(s, { x: 1.344, y: b.y, w: 0.787, h: 0.787, grad: G_CIRCLE, shadow: SHADOW_CHIP });
    glyph(s, b.icon, 1.344, b.y, 0.787, 0.787, 24);
    caption(s, 'Subtitle Services', { x: 0.362, y: b.y + 1.171, w: 2.75, h: 0.303, color: WHITE, align: 'center' });
    txt(s, lorem(75), {
      x: 0.601, y: b.y + 1.523, w: 2.273, h: 0.858, fontSize: 10, color: WHITE,
      align: 'center', lineSpacingMultiple: 1.5
    });
  });
  eyebrow(s, 8.109, 1.4);
  heading(s, ['Amazing', 'Best Services', 'For You'], { x: 8.109, y: 1.759, w: 4.18, h: 1.717 });
  para(s, lorem(360), { x: 8.109, y: 3.77, w: 4.291, h: 1.868 });
  moreInfo(s, 8.247, 6.058);
}

function slide12(s) { // three photo columns
  [0.804, 4.836, 8.868].forEach(x => {
    caption(s, 'Description Services', { x, y: 5.465, w: 2.836, h: 0.337, fontSize: 14, color: GREEN });
    para(s, lorem(110), { x: x + 0.003, y: 5.876, w: 3.339, h: 0.858 });
  });
  eyebrowC(s, 5.382, 0.773);
  heading(s, 'Our Professional Services', { x: 3.193, y: 1.133, w: 6.979, h: 0.64, align: 'center' });
}

function slide13(s) { // team member with skill bars
  freeform(s, {
    x: 0, y: 0, w: 5.646, h: 4.607, color: GREEN,
    points: [
      { x: 0, y: 0 }, { x: 5.096, y: 0 }, { x: 5.216, y: 0.081 },
      { x: 5.161, y: 1.657, curve: { type: 'cubic', x1: 5.746, y1: 0.475, x2: 5.851, y2: 0.991 } },
      { x: 0.351, y: 4.607, curve: { type: 'cubic', x1: 2.953, y1: 3.789, x2: 2.349, y2: 4.612 } },
      { x: 0.168, y: 4.602, curve: { type: 'cubic', x1: 0.289, y1: 4.607, x2: 0.228, y2: 4.606 } },
      { x: 0, y: 4.588 }, { close: true }
    ]
  });
  gradOval(s, { x: 0.621, y: 0.963, w: 5.591, h: 5.591, grad: G_CIRCLE });
  const BARS = [
    { y: 5.258, labelY: 4.887, pctY: 5.115, done: 2.756, color: GREEN, pct: '70%' },
    { y: 6.018, labelY: 5.656, pctY: 5.875, done: 3.410, color: LIME, pct: '85%' }
  ];
  BARS.forEach(b => {
    s.addShape('line', { x: 7.293, y: b.y, w: 3.995, h: 0, line: { color: SILVER, width: 6 } });
    s.addShape('line', { x: 7.293, y: b.y, w: b.done, h: 0, line: { color: b.color, width: 6 } });
    txt(s, 'Text Here', { x: 7.135, y: b.labelY, w: 1.169, h: 0.269, fontSize: 10, bold: true, color: DARK });
    txt(s, b.pct, { x: 11.578, y: b.pctY, w: 0.68, h: 0.286, fontSize: 11, bold: true, color: DARK });
  });
  eyebrow(s, 7.135, 1.436);
  heading(s, ['Our Amazing', 'Team Agriculture'], { x: 7.135, y: 1.796, w: 4.981, h: 1.178 });
  caption(s, 'MIRANDA', { x: 7.135, y: 3.348, w: 2.642, h: 0.337, fontSize: 14, color: GREEN });
  para(s, lorem(190), { x: 7.138, y: 3.766, w: 5.117, h: 0.858 });
}

function slide14(s) { // two profile cards
  s.addShape('rect', { x: 7.69, y: 0, w: 3.409, h: 7.5, fill: { color: LIGHT }, line: { type: 'none' } });
  const CARDS = [
    { x: 6.383, bg: GREEN, fg: WHITE, name: 'FERNANDES', handle: '@fernandes', social: 7.206, textX: 6.726, textY: 4.486 },
    { x: 9.624, bg: WHITE, fg: BODY, name: 'MIRANDA', handle: '@miranda', social: 10.475, textX: 9.963, textY: 4.48 }
  ];
  CARDS.forEach(c => {
    s.addShape('roundRect', {
      x: c.x, y: 2.402, w: 2.928, h: 3.825, rectRadius: 0.195,
      fill: { color: c.bg }, line: { type: 'none' }, shadow: SHADOW_HALO
    });
    caption(s, c.name, { x: c.x + 0.436, y: 3.768, w: 1.975, h: 0.337, fontSize: 14, color: c.fg === WHITE ? WHITE : GREEN, align: 'center' });
    txt(s, c.handle, {
      x: c.x + 0.469, y: 4.131, w: 1.91, h: 0.278, fontSize: 10, italic: true, color: c.fg, align: 'center'
    });
    txt(s, lorem(75), {
      x: c.textX, y: c.textY, w: 2.193, h: 0.858, fontSize: 10, color: c.fg,
      align: 'center', lineSpacingMultiple: 1.5
    });
    socialRow(s, c.social, 5.613, 1.202);
  });
  eyebrow(s, 1.119, 1.538);
  heading(s, ['Our Amazing', 'Team Agriculture'], { x: 1.119, y: 1.897, w: 4.816, h: 1.178 });
  para(s, lorem(360), { x: 1.119, y: 3.39, w: 4.291, h: 1.868 });
  moreInfo(s, 1.256, 5.709);
}

function slide15(s) { // three circular team portraits
  const TEAM = [
    { ring: 1.773, ringY: 2.298, name: 'NAMIRA SMITH', nameX: 2.06, nameY: 5.37, roleY: 5.752, social: 2.447, socialY: 6.275 },
    { ring: 5.377, ringY: 2.286, name: 'FERNANDO', nameX: 5.684, nameY: 5.358, roleY: 5.74, social: 6.071, socialY: 6.245 },
    { ring: 8.962, ringY: 2.286, name: 'JESSICE JANE', nameX: 9.291, nameY: 5.358, roleY: 5.74, social: 9.677, socialY: 6.245 }
  ];
  TEAM.forEach(t => s.addShape('ellipse', {
    x: t.ring, y: t.ringY, w: 2.598, h: 2.598, fill: { color: GREEN }, line: { type: 'none' }
  }));
  TEAM.forEach(t => {
    caption(s, t.name, { x: t.nameX, y: t.nameY, w: 1.975, h: 0.303, color: BODY, align: 'center' });
    txt(s, 'Lorem Ipsum Dolor', {
      x: t.nameX + 0.032, y: t.roleY, w: 1.91, h: 0.278, fontSize: 10, italic: true, align: 'center'
    });
    socialRow(s, t.social, t.socialY, 1.202);
  });
  eyebrowC(s, 5.382, 0.698);
  heading(s, 'Meet Our Best Team', { x: 3.602, y: 1.057, w: 6.13, h: 0.64, align: 'center' });
}

function slide16(s) { // four team bios flanking a photo grid
  const PEOPLE = [
    { name: 'Kim Jennie', nameX: 1.438, y: 2.387, roleX: 1.303, bodyX: 0.909, social: 2.257, align: 'right' },
    { name: 'Abdullah', nameX: 1.441, y: 4.794, roleX: 1.307, bodyX: 0.913, social: 2.261, align: 'right' },
    { name: 'John Doe', nameX: 9.747, y: 2.387, roleX: 9.747, bodyX: 9.747, social: 9.889, align: 'left' },
    { name: 'Jessica', nameX: 9.747, y: 4.794, roleX: 9.747, bodyX: 9.747, social: 9.889, align: 'left' }
  ];
  PEOPLE.forEach(p => {
    caption(s, p.name, { x: p.nameX, y: p.y, w: 2.161, h: 0.337, fontSize: 14, color: GREEN, align: p.align });
    caption(s, 'Professional Petani', { x: p.roleX, y: p.y + 0.386, w: 2.295, h: 0.286, fontSize: 11, color: MUTED, align: p.align });
    txt(s, lorem(64), {
      x: p.bodyX, y: p.y + 0.757, w: 2.689, h: 0.577, fontSize: 10,
      align: p.align, lineSpacingMultiple: 1.5
    });
    socialRow(s, p.social, p.y + 1.619, 1.202);
  });
  eyebrowC(s, 5.382, 0.698);
  heading(s, 'Meet Our Best Team', { x: 3.602, y: 1.057, w: 6.13, h: 0.64, align: 'center' });
}

function slide17(s) { // section break
  gradRect(s, { x: 0, y: 0, w: 13.333, h: 7.5, grad: G_HERO, steps: 70 });
  txt(s, 'Break Slide', {
    x: 1.604, y: 2.413, w: 10.142, h: 1.717, fontFace: HEAD, fontSize: 96, bold: true,
    color: WHITE, align: 'center', charSpacing: 3
  });
  txt(s, lorem(191), {
    x: 2.431, y: 4.272, w: 8.472, h: 0.606, fontSize: 10, color: WHITE,
    align: 'center', lineSpacingMultiple: 1.5
  });
}

function slide18(s) { // portfolio intro with two stats
  const STATS = [
    { ovalX: 1.156, ovalY: 5.126, iconX: 1.34, textX: 1.97, value: '240.2K', icon: ICON.download },
    { ovalX: 3.729, ovalY: 5.124, iconX: 3.913, textX: 4.543, value: '56.2K', icon: ICON.play }
  ];
  STATS.forEach(k => {
    gradOval(s, { x: k.ovalX, y: k.ovalY, w: 0.63, h: 0.63, grad: G_DOT, shadow: SHADOW_ROUND });
    glyph(s, k.icon, k.ovalX, k.ovalY, 0.63, 0.63, 16);
    caption(s, k.value, { x: k.textX, y: 5.119, w: 1.155, h: 0.337, fontSize: 14 });
    para(s, lorem(12), { x: k.textX, y: 5.532, w: 1.49, h: 0.303, fontSize: 9 });
  });
  eyebrow(s, 0.976, 1.68);
  heading(s, 'Our Portfolio', { x: 0.976, y: 2.039, w: 4.753, h: 0.64 });
  para(s, lorem(360), { x: 0.976, y: 2.952, w: 4.897, h: 1.616 });
}

function slide19(s) { // portfolio grid header card
  s.addShape('roundRect', {
    x: 0.681, y: 0.62, w: 6.583, h: 3.129, rectRadius: 0.219,
    fill: { color: GREEN }, line: { type: 'none' }
  });
  eyebrow(s, 1.341, 1.147, WHITE);
  heading(s, 'Our Portfolio', { x: 1.341, y: 1.506, w: 4.753, h: 0.64, color: WHITE });
  para(s, lorem(190), { x: 1.341, y: 2.419, w: 5.294, h: 0.858, color: WHITE });
}

function slide20(s) { // phone mockup
  s.addShape('rect', { x: 9.873, y: 0, w: 3.46, h: 7.5, fill: { color: GREEN }, line: { type: 'none' } });
  photo(s, 6.079, 0.492, 6.73, 6.508, '[image] phone mockup');
  eyebrow(s, 1.048, 1.609);
  heading(s, 'Phone Mockup', { x: 1.048, y: 1.969, w: 4.753, h: 0.64 });
  para(s, lorem(360), { x: 1.048, y: 2.952, w: 4.891, h: 1.616 });
  [
    { x: 1.048, y: 4.906, value: '+345', label: 'TOTAL DOWLOAD', labelW: 2.185 },
    { x: 3.449, y: 4.88, value: '+120', label: 'GOD BLES YOU', labelW: 2.058 }
  ].forEach(k => {
    txt(s, k.value, { x: k.x, y: k.y, w: 1.948, h: 0.64, fontSize: 32, color: GREEN });
    txt(s, k.label, { x: k.x, y: 5.623, w: k.labelW, h: 0.269, fontSize: 10, bold: true });
  });
}

function slide21(s) { // laptop mockup
  freeform(s, {
    x: 0, y: 2.847, w: 5.092, h: 4.653, color: GREEN,
    points: [
      { x: 1.243, y: 0 },
      { x: 5.092, y: 3.849, curve: { type: 'cubic', x1: 3.369, y1: 0, x2: 5.092, y2: 1.723 } },
      { x: 5.014, y: 4.625, curve: { type: 'cubic', x1: 5.092, y1: 4.115, x2: 5.065, y2: 4.374 } },
      { x: 5.007, y: 4.653 }, { x: 0, y: 4.653 }, { x: 0, y: 0.209 }, { x: 0.098, y: 0.173 },
      { x: 1.243, y: 0, curve: { type: 'cubic', x1: 0.460, y1: 0.061, x2: 0.844, y2: 0 } },
      { close: true }
    ]
  });
  photo(s, 0.381, 1.731, 6.331, 4.038, '[image] laptop mockup');
  eyebrow(s, 7.446, 1.731);
  heading(s, 'Laptop Mockup', { x: 7.446, y: 2.091, w: 4.753, h: 0.64 });
  para(s, lorem(298), { x: 7.443, y: 3.102, w: 4.891, h: 1.363 });
  gradOval(s, { x: 7.507, y: 4.93, w: 0.787, h: 0.787, grad: G_DISC, shadow: SHADOW_CHIP });
  glyph(s, ICON.plant, 7.507, 4.93, 0.787, 0.787, 22);
  caption(s, 'Subtitle Agriculture', { x: 8.561, y: 4.93, w: 2.924, h: 0.303 });
  para(s, lorem(86), { x: 8.561, y: 5.302, w: 3.773, h: 0.577 });
}

function slide22(s) { // desktop mockup with feature list
  photo(s, 0.412, 1.051, 5.868, 5.868, '[image] desktop mockup');
  const ROWS = [
    { tileX: 7.07, y: 2.761, icon: ICON.farmer, bodyH: 0.577 },
    { tileX: 7.07, y: 4.054, icon: ICON.barn, bodyH: 0.606 },
    { tileX: 7.087, y: 5.347, icon: ICON.plant, bodyH: 0.606 }
  ];
  ROWS.forEach(r => {
    gradRoundRect(s, { x: r.tileX, y: r.y, w: 0.709, h: 0.669, rad: 0.112, grad: G_TILE, shadow: SHADOW_TILE });
    glyph(s, r.icon, r.tileX, r.y, 0.709, 0.669, 20);
    caption(s, 'Description Agriculture', { x: 8.097, y: r.y, w: 3.147, h: 0.303 });
    para(s, lorem(95), { x: 8.099, y: r.y + 0.393, w: 4.083, h: r.bodyH });
  });
  eyebrow(s, 6.886, 1.297);
  heading(s, 'Desktop Mockup', { x: 6.886, y: 1.657, w: 4.753, h: 0.64 });
}

function slide23(s, pptx) { // bar chart infographic
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2, 2, 3, 5] }
  ], {
    x: 6.955, y: 2.292, w: 5.458, h: 4.042,
    barDir: 'bar', barGapWidthPct: 182, showLegend: false, showTitle: false,
    chartColors: ['E7E6E6', LIME, GREEN],
    catAxisLineColor: 'D9D9D9', catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
    catAxisLabelFontFace: SANS, catAxisLabelFontSize: 12, catAxisLabelColor: BODY,
    valAxisLineShow: false, valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    valAxisLabelFontFace: SANS, valAxisLabelFontSize: 12, valAxisLabelColor: BODY,
    valGridLine: { color: 'D9D9D9', size: 1 }, catGridLine: { style: 'none' }
  });
  [
    { x: 1.21, value: '90%', textX: 1.34, capX: 1.292 },
    { x: 3.562, value: '80+', textX: 3.692, capX: 3.644 }
  ].forEach(k => {
    gradRoundRect(s, { x: k.x, y: 4.515, w: 1.782, h: 1.233, rad: 0.112, grad: G_DISC, shadow: SHADOW_STAT });
    txt(s, k.value, { x: k.textX, y: 4.725, w: 1.52, h: 0.438, fontSize: 20, color: WHITE, align: 'center' });
    txt(s, 'Graphic Second', { x: k.capX, y: 5.205, w: 1.618, h: 0.286, fontSize: 11, color: WHITE, align: 'center' });
  });
  caption(s, 'Description Agriculture', { x: 1.088, y: 2.584, w: 2.836, h: 0.303, color: BODY });
  para(s, lorem(250), { x: 1.091, y: 3.002, w: 5.031, h: 1.111 });
  eyebrowC(s, 5.367, 0.698);
  heading(s, 'Our Infographic', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center' });
}

function slide24(s, pptx) { // doughnut chart infographic
  s.addChart(pptx.ChartType.doughnut, [{
    name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2]
  }], {
    x: 1.198, y: 2.126, w: 5.469, h: 4.318,
    holeSize: 45, firstSliceAng: 0, showLegend: false, showTitle: false, showValue: false,
    chartColors: [GREEN, LIME, LIGHT, LIGHT], dataBorder: { pt: 1.5, color: WHITE }
  });
  [{ y: 2.531, color: GREEN }, { y: 4.51, color: LIME }].forEach(k => {
    s.addShape('rect', { x: 7.509, y: k.y, w: 0.316, h: 0.316, fill: { color: k.color }, line: { type: 'none' } });
    caption(s, '+98,987,124', { x: 7.869, y: k.y - 0.018, w: 1.697, h: 0.337, fontSize: 14 });
    para(s, lorem(151), { x: 7.419, y: k.y + 0.554, w: 4.55, h: 0.858 });
  });
  eyebrowC(s, 5.367, 0.698);
  heading(s, 'Our Infographic', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center' });
}

function slide25(s) { // five overlapping circles
  const RING = { color: WHITE, width: 4.5 };
  const SMALL = [
    { x: 1.364, icon: ICON.bulb, iconX: 2.349, iconY: 3.535, iconW: 0.382, iconH: 0.559, value: '220M', valueX: 2.13, label: 1.948, labelY: 5.511 },
    { x: 5.518, icon: ICON.download, iconX: 6.389, iconY: 3.535, iconW: 0.555, iconH: 0.559, value: '550M', valueX: 6.281, label: 6.068, labelY: 5.511 },
    { x: 9.673, icon: ICON.camera, iconX: 10.502, iconY: 3.535, iconW: 0.639, iconH: 0.54, value: '950M', valueX: 10.403, label: 10.223, labelY: 5.511 }
  ];
  const BIG = [
    { x: 3.132, icon: ICON.pencil, iconX: 4.232, iconY: 3.299, iconW: 0.84, iconH: 0.849, value: '585M', valueX: 4.157, label: 4.053, labelY: 5.915 },
    { x: 7.225, icon: ICON.monitor, iconX: 8.294, iconY: 3.299, iconW: 0.902, iconH: 0.849, value: '660M', valueX: 8.238, label: 8.145, labelY: 5.915 }
  ];
  // interleaved z-order: each disc overlaps the one to its left
  [0, 1, 2].forEach(i => {
    gradOval(s, { x: SMALL[i].x, y: 2.913, w: 2.296, h: 2.296, grad: G_DISC, line: RING });
    if (BIG[i]) s.addShape('ellipse', { x: BIG[i].x, y: 2.541, w: 3.04, h: 3.04, fill: { color: GREEN }, line: RING });
  });
  SMALL.forEach(c => {
    glyph(s, c.icon, c.iconX, c.iconY, c.iconW, c.iconH, 26);
    txt(s, c.value, { x: c.valueX, y: 4.182, w: 0.845, h: 0.37, fontSize: 16, bold: true, color: WHITE, align: 'center' });
    caption(s, 'Text Here', { x: c.label, y: c.labelY, w: 1.198, h: 0.337, fontSize: 14, align: 'center' });
  });
  BIG.forEach(c => {
    glyph(s, c.icon, c.iconX, c.iconY, c.iconW, c.iconH, 34);
    txt(s, c.value, { x: c.valueX, y: 4.315, w: 1.019, h: 0.438, fontSize: 20, bold: true, color: WHITE, align: 'center' });
    caption(s, 'Text Here', { x: c.label, y: c.labelY, w: 1.198, h: 0.337, fontSize: 14, align: 'center' });
  });
  eyebrowC(s, 5.367, 0.698);
  heading(s, 'Our Infographic', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center' });
}

function slide26(s) { // isometric layer stack + bullet list
  // One hollow slab: an outer isometric rectangle with a smaller one punched out.
  // The ring is drawn as four trapezoid strips (pptxgenjs paths cannot hold a hole).
  const OUTER = [[1.301, 0], [0, 0.847], [2.832, 1.242], [4.135, 0.392]];
  const INNER = [[1.331, 0.122], [0.330, 0.775], [2.807, 1.116], [3.808, 0.465]];
  const ring = (x, y) => OUTER.forEach((o, i) => {
    const q = [o, OUTER[(i + 1) % 4], INNER[(i + 1) % 4], INNER[i]];
    const x0 = Math.min.apply(null, q.map(p => p[0])), y0 = Math.min.apply(null, q.map(p => p[1]));
    freeform(s, {
      x: x + x0, y: y + y0,
      w: Math.max.apply(null, q.map(p => p[0])) - x0,
      h: Math.max.apply(null, q.map(p => p[1])) - y0,
      color: GREEN,
      points: q.map(p => ({ x: +(p[0] - x0).toFixed(3), y: +(p[1] - y0).toFixed(3) })).concat([{ close: true }])
    });
  });
  [5.399, 4.403, 3.408, 2.412].forEach((y, idx) => {
    freeform(s, { // back-right wall
      x: 3.075, y, w: 2.477, h: 0.738, color: GREEN,
      points: [{ x: 2.477, y: 0.343 }, { x: 0, y: 0 }, { x: 0, y: 0.477 }, { x: 1.873, y: 0.738 }, { close: true }]
    });
    freeform(s, { // right wall
      x: 4.577, y: y + 0.270, w: 1.303, h: 1.446, color: GREEN,
      points: [{ x: 1.303, y: 0 }, { x: 0, y: 0.850 }, { x: 0, y: 1.446 }, { x: 1.303, y: 0.597 }, { close: true }]
    });
    freeform(s, { // back-left wall
      x: 2.075, y, w: 1.0, h: 0.738, color: GREEN,
      points: [{ x: 1.0, y: 0 }, { x: 0, y: 0.653 }, { x: 0.603, y: 0.738 }, { x: 1.0, y: 0.477 }, { close: true }]
    });
    freeform(s, { // front-left wall
      x: 1.745, y: y + 0.725, w: 2.832, h: 0.991, color: GREEN,
      points: [{ x: 2.832, y: 0.991 }, { x: 0, y: 0.599 }, { x: 0, y: 0 }, { x: 2.832, y: 0.395 }, { close: true }]
    });
    ring(1.745, y - 0.122);                       // top face
    if (idx < 3) ring(1.745, y + 0.270);          // underside, hidden on the top slab
  });
  const STEPS = [
    { y: 2.547, dotY: 2.581, title: 'Executive summary', titleX: 8.460, titleW: 4.100, bodyY: 2.890, bodyW: 3.859, icon: ICON.sound },
    { y: 3.603, dotY: 3.720, title: 'Service', titleX: 8.474, titleW: 3.647, bodyY: 3.974, bodyW: 3.859, icon: ICON.sound },
    { y: 4.762, dotY: 4.844, title: 'Promotion', titleX: 8.474, titleW: 3.647, bodyY: 5.133, bodyW: 3.647, icon: ICON.sound },
    { y: 5.877, dotY: 5.997, title: 'Product Brand', titleX: 8.474, titleW: 3.647, bodyY: 6.248, bodyW: 4.085, icon: ICON.play }
  ];
  STEPS.forEach(t => {
    gradOval(s, { x: 7.484, y: t.dotY, w: 0.512, h: 0.512, grad: G_DISC });
    glyph(s, t.icon, 7.484, t.dotY, 0.512, 0.512, 15);
    txt(s, t.title, {
      x: t.titleX, y: t.y, w: t.titleW, h: 0.286,
      fontSize: 14, bold: true, color: DARK, valign: 'middle', margin: 0
    });
    txt(s, POREM, { x: 8.474, y: t.bodyY, w: t.bodyW, h: 0.252, fontSize: 10, lineSpacingMultiple: 1.5, margin: 0 });
  });
  eyebrowC(s, 5.367, 0.698);
  heading(s, 'Our Infographic', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center' });
}

function slide27(s) { // pricing tables
  const PLANS = [
    { x: 2.165, card: LIGHT, name: 'Basic', nameX: 2.897, nameW: 1.183, nameC: BODY, price: '$49/mo', priceX: 2.382, priceW: 2.214, priceC: GREEN, lineX: 2.531, lineC: RULE, itemX: 2.233, itemC: BODY, bullet: GREEN, btnX: 2.801, btnY: 6.013, cta: 2.897 },
    { x: 5.361, card: GREEN, name: 'Medium', nameX: 5.972, nameW: 1.414, nameC: WHITE, price: '$99/mo', priceX: 5.664, priceW: 2.030, priceC: WHITE, lineX: 5.720, lineC: WHITE, itemX: 5.423, itemC: WHITE, bullet: WHITE, btnX: 5.991, btnY: 6.017, cta: 6.087, dy: 0.004 },
    { x: 8.561, card: LIGHT, name: 'Premium', nameX: 9.043, nameW: 1.651, nameC: BODY, price: '$150/mo', priceX: 8.752, priceW: 2.234, priceC: GREEN, lineX: 8.910, lineC: RULE, itemX: 8.613, itemC: BODY, bullet: GREEN, btnX: 9.181, btnY: 6.013, cta: 9.277 }
  ];
  PLANS.forEach(p => {
    const dy = p.dy || 0;
    s.addShape('roundRect', {
      x: p.x, y: 2.276, w: 2.636, h: 4.376, rectRadius: 0.265,
      fill: { color: p.card }, line: { type: 'none' }, shadow: SHADOW_CARD
    });
    caption(s, p.name, { x: p.nameX, y: 2.599 + dy, w: p.nameW, h: 0.348, fontSize: 14, color: p.nameC, align: 'center' });
    txt(s, p.price, {
      x: p.priceX, y: 2.987 + dy, w: p.priceW, h: 0.505, fontSize: 24, bold: true,
      color: p.priceC, align: 'center'
    });
    [3.649, 4.185, 4.685, 5.185, 5.685].forEach((ly, i) => s.addShape('line', {
      x: p.lineX, y: ly + dy, w: 1.917, h: 0,
      line: { color: p.lineC, width: 1, dashType: i === 0 ? 'solid' : 'dash' }
    }));
    // bullet is drawn at 200% size like the source; lineSpacing pins the baseline
    [3.834, 4.304, 4.828, 5.302].forEach(iy => s.addText([
      { text: '\u2022 ', options: { fontSize: 20, color: p.bullet } },
      { text: 'Lorem Ipsum\u00A0is simply', options: { fontSize: 10, color: p.itemC } }
    ], {
      x: p.itemX, y: iy + dy - 0.055, w: 2.512, h: 0.269, fontFace: SANS,
      align: 'center', valign: 'top', lineSpacing: 13, isTextBox: true
    }));
    gradRect(s, { x: p.btnX, y: p.btnY, w: 1.375, h: 0.368, grad: G_DISC });
    txt(s, 'TRY NOW', {
      x: p.cta, y: p.btnY + 0.056, w: 1.183, h: 0.286, fontSize: 11, bold: true,
      color: WHITE, align: 'center'
    });
  });
  eyebrowC(s, 5.367, 0.698);
  heading(s, 'Pricing Tables', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center' });
}

function slide28(s) { // testimonials over a soft wash
  gradRect(s, { x: 0, y: 0, w: 13.333, h: 7.5, grad: G_HERO_SOFT, steps: 70 });
  const CARDS = [
    { x: 2.164, barX: 2.152, y: 2.341, barY: 2.316, name: 'Son Smith', nameX: 2.470, dateX: 2.607, bodyX: 2.317, stars: 2.961 },
    { x: 5.323, barX: 5.323, y: 2.316, barY: 2.300, name: 'Fernandes', nameX: 5.630, dateX: 5.767, bodyX: 5.477, stars: 6.121 },
    { x: 8.481, barX: 8.494, y: 2.300, barY: 2.300, name: 'Firman', nameX: 8.811, dateX: 8.948, bodyX: 8.658, stars: 9.302 }
  ];
  CARDS.forEach(c => {
    s.addShape('roundRect', {
      x: c.x, y: c.y, w: 2.714, h: 4.139, rectRadius: 0.171,
      fill: { color: WHITE }, line: { type: 'none' }, shadow: SHADOW_QUOTE
    });
    gradRoundRect(s, { x: c.barX, y: c.barY, w: 2.714, h: 0.131, rad: 0.0655, grad: G_BAR });
    caption(s, c.name, { x: c.nameX, y: 4.36, w: 2.077, h: 0.337, fontSize: 14, color: BODY, align: 'center' });
    txt(s, 'London,05 April 1988', {
      x: c.dateX, y: 4.765, w: 1.806, h: 0.269, fontSize: 10, italic: true, align: 'center'
    });
    txt(s, lorem(48), {
      x: c.bodyX, y: 5.086, w: 2.384, h: 0.577, fontSize: 10, align: 'center', lineSpacingMultiple: 1.5
    });
    for (let i = 0; i < 5; i++) s.addShape('star5', {
      x: c.stars + i * 0.234, y: 5.919, w: 0.173, h: 0.173,
      fill: { color: i < 4 ? STAR : SILVER }, line: { type: 'none' }
    });
  });
  eyebrowC(s, 5.367, 0.698, WHITE);
  heading(s, 'Our Customer Here', { x: 4.03, y: 1.057, w: 5.293, h: 0.64, align: 'center', color: WHITE });
}

function slide29(s) { // contact details
  s.addShape('rect', { x: 6.667, y: 2.985, w: 0.773, h: 4.515, fill: { color: GREEN }, line: { type: 'none' } });
  const CONTACT = [
    { icon: ICON.home, iconX: 1.245, iconY: 2.933, iconW: 0.434, iconH: 0.412, label: 'Address', y: 2.896, valueY: 3.287, value: 'Lorem Ipsum\u00A0is simply dummy text. ', valueW: 3.546 },
    { icon: ICON.mail, iconX: 1.276, iconY: 4.014, iconW: 0.372, iconH: 0.26, label: 'Email', y: 3.938, valueY: 4.329, value: 'www.agriculturepresentation.com', valueW: 3.546 },
    { icon: ICON.phone, iconX: 1.294, iconY: 4.957, iconW: 0.386, iconH: 0.383, label: 'Telephone', y: 4.932, valueY: 5.323, value: '08-0089-0977-4546', valueW: 2.543 }
  ];
  CONTACT.forEach(c => {
    glyph(s, c.icon, c.iconX - 0.1, c.iconY - 0.08, c.iconW + 0.2, c.iconH + 0.16, 26, midColor(G_DOT));
    caption(s, c.label, { x: 2.021, y: c.y, w: 1.979, h: 0.337, fontSize: 14 });
    txt(s, c.value, { x: 2.021, y: c.valueY, w: c.valueW, h: 0.347, fontSize: 11, lineSpacingMultiple: 1.5 });
  });
  txt(s, '08-0089-0977-0908', { x: 2.021, y: 5.733, w: 2.543, h: 0.347, fontSize: 11, lineSpacingMultiple: 1.5 });
  eyebrow(s, 1.245, 1.409);
  heading(s, 'Get It Touch', { x: 1.245, y: 1.768, w: 4.322, h: 0.64 });
}

function slide30(s) { // thank you
  gradRect(s, { x: 0, y: 0, w: 13.333, h: 7.5, grad: G_HERO, steps: 70 });
  txt(s, 'Thank You', {
    x: 1.747, y: 2.907, w: 9.84, h: 1.717, fontFace: HEAD, fontSize: 96, bold: true,
    color: WHITE, align: 'center', charSpacing: 3
  });
  txt(s, 'Agriculture Presentation Template', {
    x: 3.53, y: 4.783, w: 6.273, h: 0.337, fontSize: 14, color: WHITE, align: 'center', charSpacing: 3
  });
}

/* ---------------------------------------------------------------- assemble */
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'AGRI_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'AGRI_16x9';
  pptx.title = 'Agriculture';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: SANS };
  SLIDES.forEach(fn => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    fn(slide, pptx);
  });
  return pptx.writeFile({
    fileName: path.join(__dirname, '02e24a99-b934-41b0-b2dc-a6ec3b07930d_grok_final.pptx')
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
