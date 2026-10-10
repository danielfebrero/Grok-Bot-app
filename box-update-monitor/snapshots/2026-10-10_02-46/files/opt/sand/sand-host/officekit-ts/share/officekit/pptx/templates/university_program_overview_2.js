/**
 * "University Program" template deck — rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9).  Photographs are replaced by grey
 * placeholder shapes; icons are replaced by small rounded-square marks.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / fonts / shared metrics
 * ------------------------------------------------------------------ */

const C = {
  orange: 'FF8825',      // accent1
  orangeLt: 'FFB87C',    // accent1 lum 60/40
  orangeAlt: 'FE7A36',   // accent6
  sand: 'FFDD95',        // accent2
  blue: '3468C0',        // accent4
  blueLt: '86A7FC',      // accent3
  blueDk: '274E90',
  white: 'FFFFFF',
  black: '000000',
  grey: '7F7F7F',
  darkGrey: '404040',
  ph: 'E5E5E5',          // empty picture-placeholder grey
  phDark: 'CCCCCC',
  arch: 'E6E6E6'
};

const F = { head: 'Poppins SemiBold', body: 'Open Sans' };

const SUN = [[0.45, C.orangeLt], [1, 'FFEDDE']];                    // pale decorative dot
const CARD_GRAD = [[0.37, C.orange], [1, C.orangeLt]];              // orange panels
const PANEL_BANDS = 34;                                             // gradient resolution
const CARD_SHADOW = { type: 'outer', blur: 20, offset: 4, angle: 90, color: C.black, opacity: 0.43 };
const SOFT_SHADOW = { type: 'outer', blur: 60, offset: 0, angle: 45, color: C.black, opacity: 0.14 };

const LOREM = 'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_SHORT = 'PLACEHOLDER';

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

const mix = (a, b, t) => [0, 2, 4]
  .map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t))
  .map(v => v.toString(16).padStart(2, '0').toUpperCase())
  .join('');

/** Colour of a multi-stop linear gradient at position t (0..1). */
function colorAt(stops, t) {
  if (t <= stops[0][0]) return stops[0][1];
  if (t >= stops[stops.length - 1][0]) return stops[stops.length - 1][1];
  for (let i = 1; i < stops.length; i++) {
    const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
    if (t <= p1) return mix(c0, c1, (t - p0) / (p1 - p0));
  }
  return stops[stops.length - 1][1];
}

/**
 * Linear gradient fills are emulated with flat-coloured pieces, because
 * pptxgenjs can only write solid fills.
 *
 *   ellipse   -> the disc, then nested "chord" caps perpendicular to the axis
 *   rectangle -> the shape itself (silhouette + shadow) covered by a mosaic of
 *                plain rectangles; every rounded corner is covered instead by a
 *                quarter-disc so that nothing spills outside the outline.
 *
 * roundRect rounds all four corners; round1Rect rounds the top-right one,
 * moved by flipH / flipV.
 */
function gradShape(slide, shape, box, stops, angleDeg, extra) {
  const o = Object.assign({}, extra);
  const bands = o.bands || 18;
  delete o.bands;
  const one = (o.flipV ? 'b' : 't') + (o.flipH ? 'l' : 'r');
  const corners = shape === 'roundRect' ? ['tl', 'tr', 'br', 'bl'] : shape === 'round1Rect' ? [one] : [];
  const rad = angleDeg * Math.PI / 180;
  const cx = Math.cos(rad), cy = Math.sin(rad);

  if (shape === 'ellipse') {
    slide.addShape('ellipse', Object.assign({}, box, o, { fill: { color: colorAt(stops, 0.5 / bands) } }));
    for (let i = 1; i < bands; i++) {
      const half = Math.acos(2 * (i / bands) - 1) * 180 / Math.PI;
      slide.addShape('chord', Object.assign({}, box, o, {
        fill: { color: colorAt(stops, (i + 0.5) / bands) },
        angleRange: [(angleDeg - half + 720) % 360, (angleDeg + half + 720) % 360]
      }));
    }
    return;
  }

  const span = box.w * Math.abs(cx) + box.h * Math.abs(cy);
  const r = Math.min(o.rectRadius || 0, box.w / 2, box.h / 2);
  // gradient position of a point given in box-local inches
  const tAt = (u, v) => ((cx >= 0 ? u : box.w - u) * Math.abs(cx) +
                         (cy >= 0 ? v : box.h - v) * Math.abs(cy)) / span;
  const step = span / bands;

  slide.addShape(shape, Object.assign({}, box, o, { fill: { color: colorAt(stops, tAt(box.w / 2, box.h / 2)) } }));

  const tile = (u0, v0, u1, v1) => {
    if (u1 - u0 < 1e-4 || v1 - v0 < 1e-4) return;
    const nu = Math.min(40, Math.max(1, Math.round((u1 - u0) * Math.abs(cx) / step)));
    const nv = Math.min(40, Math.max(1, Math.round((v1 - v0) * Math.abs(cy) / step)));
    for (let i = 0; i < nu; i++) {
      for (let j = 0; j < nv; j++) {
        const cu = u0 + (u1 - u0) * (i + 0.5) / nu, cv = v0 + (v1 - v0) * (j + 0.5) / nv;
        slide.addShape('rect', {
          x: box.x + u0 + (u1 - u0) * i / nu, y: box.y + v0 + (v1 - v0) * j / nv,
          w: (u1 - u0) / nu + 0.004, h: (v1 - v0) / nv + 0.004,
          fill: { color: colorAt(stops, tAt(cu, cv)) }
        });
      }
    }
  };
  const has = c => corners.indexOf(c) >= 0;
  tile(0, r, box.w, box.h - r);                                       // straight middle
  tile(has('tl') ? r : 0, 0, has('tr') ? box.w - r : box.w, r);       // top edge
  tile(has('bl') ? r : 0, box.h - r, has('br') ? box.w - r : box.w, box.h);

  const quads = { tl: [0, 0, 180], tr: [box.w - 2 * r, 0, 270], br: [box.w - 2 * r, box.h - 2 * r, 0], bl: [0, box.h - 2 * r, 90] };
  corners.forEach(c => {
    const [u, v, a] = quads[c];
    slide.addShape('pie', {
      x: box.x + u, y: box.y + v, w: 2 * r, h: 2 * r, angleRange: [a, a + 90],
      fill: { color: colorAt(stops, tAt(u + r + (c[1] === 'l' ? -0.4 : 0.4) * r, v + r + (c[0] === 't' ? -0.4 : 0.4) * r)) }
    });
  });
}

/** Text box; the reference deck anchors all of its text boxes to the top. */
function T(slide, runs, opts) {
  slide.addText(runs, Object.assign({ fontFace: F.body, fontSize: 12, color: C.black, valign: 'top' }, opts));
}

/** Body copy: 12 pt, 150 % leading. */
function body(slide, txt, x, y, w, h, opts) {
  T(slide, txt, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.5 }, opts));
}

/** Small "University Template" lock-up in the top-left corner. */
function logo(slide, x, y, light) {
  x = x === undefined ? 0.823 : x;
  y = y === undefined ? 0.46 : y;
  slide.addShape('roundRect', { x, y, w: 0.295, h: 0.295, rectRadius: 0.055, fill: { color: light ? C.white : C.orange } });
  T(slide, [
    { text: 'University ', options: { bold: true, italic: true, color: light ? C.white : C.orange } },
    { text: 'Template', options: { bold: true, color: light ? C.white : C.black } }
  ], { x: x + 0.233, y: y + 0.021, w: 1.749, h: 0.268, fontSize: 11, lineSpacingMultiple: 0.9 });
}

/** Kicker line + two-tone display title. */
function heading(slide, o) {
  T(slide, o.kicker, {
    x: o.kickerX === undefined ? o.x + 0.03 : o.kickerX, y: o.y, w: 3.812, h: 0.337,
    fontSize: 14, color: o.light ? C.white : C.black, align: o.kickerAlign || o.align || 'left'
  });
  T(slide, o.parts.map(p => ({
    text: p[0],
    options: Object.assign({ color: p[1] || (o.light ? C.white : C.black) }, p[2] || {})
  })), {
    x: o.x, y: o.y + (o.gap === undefined ? 0.37 : o.gap), w: o.w, h: o.h,
    fontFace: F.head, fontSize: o.size, bold: true, align: o.align || 'left',
    color: o.light ? C.white : C.black, lineSpacingMultiple: 0.9
  });
}

/** Icon stand-in: an outlined rounded square, matching the line-art icons' weight. */
function icon(slide, x, y, d, color) {
  const s = d * 0.74, off = (d - s) / 2;
  slide.addShape('roundRect', {
    x: x + off, y: y + off, w: s, h: s, rectRadius: s * 0.2,
    line: { color, width: Math.max(1, s * 7) }
  });
}

/** Flat colour disc + icon. */
function iconDisc(slide, x, y, d, fill, iconColor) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill } });
  icon(slide, x + d * 0.193, y + d * 0.193, d * 0.614, iconColor);
}

/** Orange badge with a white ring (d = 0.852 everywhere in the deck). */
function badge(slide, x, y, d) {
  d = d || 0.852;
  slide.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color: C.orangeAlt }, line: { color: C.white, width: 4.75 }
  });
  icon(slide, x + d * 0.2, y + d * 0.2, d * 0.6, C.white);
}

/** Decorative gradient dot. */
function sunDot(slide, x, y, d) {
  gradShape(slide, 'ellipse', { x, y, w: d, h: d }, SUN, 50, { bands: 10 });
}

/** Short orange rule used as a page mark. */
function rule(slide, x, y) {
  slide.addShape('line', { x, y, w: 0.301, h: 0, line: { color: C.orange, width: 3.25 } });
}

/** Grey stand-in for a photograph. */
function photo(slide, o) {
  slide.addShape(o.shape || 'rect', Object.assign(
    { fill: { color: o.color || C.ph } },
    o, { shape: undefined, label: undefined, color: undefined }
  ));
  if (o.label !== false) {
    T(slide, '[image]', {
      x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
      fontSize: 12, color: C.grey, align: 'center', valign: 'middle'
    });
  }
}

/** Blue rounded card. */
function card(slide, x, y, w, h, r, opts) {
  slide.addShape('roundRect', Object.assign({
    x, y, w, h, rectRadius: r, fill: { color: C.blue }, shadow: CARD_SHADOW
  }, opts));
}

/** Card with a white icon disc, bold heading and body copy (blue tiles). */
function iconCard(slide, o) {
  card(slide, o.x, o.y, o.w, o.h, o.r === undefined ? 0.2235 : o.r, o.card);
  if (o.disc !== false) iconDisc(slide, o.x + (o.dx || 0.31), o.y + (o.dy || 0.243), 0.683, C.white, C.blue);
  T(slide, o.title, {
    x: o.tx, y: o.y + (o.ty || 0.243), w: o.tw || 2.749, h: 0.364,
    fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.2
  });
  body(slide, o.body, o.tx, o.y + (o.by || 0.617), o.bw || 2.2, o.bh || 0.675, { color: C.white });
}

/** Big number + caption + body, as used inside the blue "stat" cards. */
function statBlock(slide, o) {
  T(slide, o.value, {
    x: o.x, y: o.y, w: o.vw || 2.231, h: 0.59, fontFace: F.head, fontSize: o.size || 32,
    bold: true, color: o.color || C.white, align: o.align || 'left', lineSpacingMultiple: 0.9
  });
  T(slide, o.caption === undefined ? 'Text Here' : o.caption, {
    x: o.cx === undefined ? o.x + 0.01 : o.cx, y: o.cy, w: o.cw || 2.508, h: 0.367,
    fontSize: o.capSize || 14, italic: o.italic !== false, bold: !!o.capBold,
    color: o.capColor || o.color || C.white, align: o.align || 'left', lineSpacingMultiple: 1.2
  });
  if (o.body) body(slide, o.body, o.bx === undefined ? o.x + 0.01 : o.bx, o.by, o.bw || 2.231, o.bh || 0.978, { color: o.color || C.white });
}

/** "01." / "02." numbered paragraph pairs. */
function numberedList(slide, x, y, items) {
  items.forEach((it, i) => {
    T(slide, it.num, {
      x, y: y + i * 1.383, w: 0.9, h: 0.529,
      fontFace: F.head, fontSize: 28, bold: true, color: C.orange, lineSpacingMultiple: 0.9
    });
    body(slide, it.lines, x + 0.738, y + i * 1.383, it.w || 3.72, 0.978);
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

const slides = [];

/* 1 — cover -------------------------------------------------------- */
slides.push(s => {
  gradShape(s, 'ellipse', { x: 6.055, y: 0.46, w: 6.547, h: 6.547 }, SUN, 50, { bands: 22 });
  logo(s);
  heading(s, {
    x: 0.702, y: 1.553, w: 4.124, h: 1.744, size: 54, kicker: 'Specializing in Different Areas',
    parts: [['University '], ['Program', C.orange]]
  });
  body(s, LOREM, 1.727, 4.306, 3.439, 1.28);
  T(s, 'WWW. YOURWEBSITE. COM', {
    x: 0.731, y: 6.082, w: 3.29, h: 0.372, color: C.orange, charSpacing: 3, lineSpacingMultiple: 1.5
  });
  iconDisc(s, 0.835, 4.306, 0.673, C.orange, C.white);
  badge(s, 6.156, 1.589);
  rule(s, 0.845, 6.809);
  s.addShape('donut', { x: 11.489, y: 5.199, w: 0.938, h: 0.938, fill: { color: C.orange } });

  card(s, 10.803, 2.015, 2.059, 1.058, 0.1514);
  icon(s, 11.035, 2.163, 0.413, C.white);
  statBlock(s, { x: 11.43, y: 2.196, vw: 1.299, size: 24, value: '+11,8%', cx: 11.439, cy: 2.616, cw: 1.29, capSize: 12, caption: 'Text Tittle' });

  card(s, 5.463, 4.964, 2.059, 1.058, 0.1514);
  icon(s, 5.657, 5.2, 0.413, C.white);
  body(s, 'At vero eos et accusamus ', 6.063, 5.2, 1.327, 0.627, { fontSize: 11, color: C.white });
});

/* 2 — table of content --------------------------------------------- */
slides.push(s => {
  const cards = [
    { x: 0.839, n: '01', grad: true },
    { x: 3.772, n: '02', grad: false },
    { x: 6.705, n: '03', grad: true },
    { x: 9.638, n: '04', grad: true }
  ];
  cards.forEach(c => {
    const box = { x: c.x, y: 3.015, w: 2.856, h: 3.075 };
    if (c.grad) gradShape(s, 'roundRect', box, CARD_GRAD, 50, { rectRadius: 0.3154, bands: PANEL_BANDS });
    else s.addShape('roundRect', Object.assign({}, box, {
      rectRadius: 0.3154, fill: { color: C.blue },
      shadow: { type: 'outer', blur: 21, offset: 4, angle: 90, color: C.black, opacity: 0.56 }
    }));
    T(s, c.n, { x: c.x + 0.344, y: 3.468, w: 0.912, h: 0.64, fontSize: 32, bold: true, color: C.white });
    T(s, 'Heading text ', { x: c.x + 0.344, y: 4.07, w: 1.941, h: 0.337, fontSize: 14, bold: true, italic: true, color: C.white });
    body(s, 'At vero eos et accusamus et iusto odio dignissimos', c.x + 0.344, 4.849, 2.27, 0.678, { color: C.white });
  });
  logo(s);
  badge(s, 11.213, 2.548);
  heading(s, {
    x: 1.599, y: 1.162, w: 10.136, h: 0.774, size: 44, align: 'center',
    kicker: 'Discover Your Path to Success', kickerX: 4.761, gap: 0.478,
    parts: [['Table Of '], ['Content', C.orange]]
  });
});

/* 3 — about --------------------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 5.225, y: 0, w: 2.622, h: 4.756 }, CARD_GRAD, 230,
    { rectRadius: 0.7216, flipH: true, flipV: true, bands: PANEL_BANDS });
  logo(s);
  heading(s, {
    x: 0.702, y: 1.609, w: 4.124, h: 2.106, size: 44, kicker: 'Specializing in Different Areas',
    parts: [['About University '], ['Program', C.orange]]
  });
  T(s, 'Text Here', { x: 1.675, y: 4.594, w: 2.993, h: 0.337, fontSize: 14, bold: true, color: C.orange });
  body(s, LOREM, 1.675, 5.011, 2.993, 1.583);
  iconDisc(s, 0.835, 4.594, 0.673, C.orange, C.white);
  rule(s, 12.194, 6.727);

  card(s, 5.84, 4.052, 2.993, 2.809, 0.3266);
  statBlock(s, { x: 6.239, y: 4.497, value: '+46,8%', cx: 6.249, cy: 5.094, body: LOREM_SHORT, bx: 6.249, by: 5.438 });
  badge(s, 10.769, 1.467);
  sunDot(s, 12.194, 4.377, 0.601);
});

/* 4 — service ------------------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 3.344, y: 2.263, w: 2.888, h: 4.463 }, CARD_GRAD, 310,
    { rectRadius: 0.7948, flipV: true, bands: PANEL_BANDS });
  logo(s);
  heading(s, {
    x: 8.35, y: 1.08, w: 4.124, h: 2.106, size: 44, kicker: 'Discover Your Path to Success',
    parts: [['Service University '], ['Program', C.orange]]
  });
  numberedList(s, 8.361, 4.186, [
    { num: '01.', lines: 'PLACEHOLDER', w: 3.092 },
    { num: '02.', lines: 'PLACEHOLDER', w: 2.911 }
  ]);
  sunDot(s, 12.186, 0.947, 0.601);
  rule(s, 12.194, 6.727);
  badge(s, 2.98, 1.006);

  card(s, 3.096, 4.567, 3.728, 1.827, 0.2688);
  iconDisc(s, 3.472, 4.841, 0.683, C.white, C.blue);
  T(s, 'Text Here', { x: 4.317, y: 4.841, w: 2.508, h: 0.367, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  body(s, LOREM_SHORT, 4.317, 5.184, 2.231, 0.978, { color: C.white });
});

/* 5 — teachers ------------------------------------------------------ */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 5.84, y: 0, w: 7.493, h: 7.5 }, CARD_GRAD, 50,
    { rectRadius: 1.1471, flipH: true, flipV: true, bands: PANEL_BANDS });
  logo(s);
  heading(s, {
    x: 0.704, y: 1.469, w: 4.124, h: 2.106, size: 44, kicker: 'Discover Your Path to Success',
    parts: [['Teachers University '], ['Program', C.orange]]
  });
  T(s, 'Text Here', { x: 1.675, y: 4.841, w: 2.993, h: 0.337, fontSize: 14, bold: true, color: C.orange });
  body(s, 'PLACEHOLDER',
    1.675, 5.23, 2.993, 1.28);
  iconDisc(s, 0.835, 4.702, 0.673, C.orange, C.white);

  [1.084, 4.114].forEach(rowY => {
    [6.699, 8.838, 10.946].forEach(colX => badge(s, colX, rowY, 0.553));
  });
  [{ y: 2.794, ty: 3.052 }, { y: 5.801, ty: 6.059 }].forEach(row => {
    [6.699, 8.797, 10.895].forEach(colX => {
      T(s, 'Name Here', { x: colX, y: row.y, w: 1.58, h: 0.337, fontSize: 14, bold: true, color: C.white, align: 'center' });
      T(s, 'Text Tittle Here', {
        x: colX + 0.108, y: row.ty, w: 1.364, h: 0.349, fontSize: 10.5, italic: true,
        color: C.white, align: 'center', lineSpacingMultiple: 1.5
      });
    });
  });
});

/* 6 — benefits ------------------------------------------------------ */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 9.816, y: 0.521, w: 2.913, h: 4.463 }, CARD_GRAD, 310,
    { rectRadius: 0.8017, flipV: true, bands: PANEL_BANDS });
  logo(s);
  heading(s, {
    x: 0.704, y: 1.187, w: 4.124, h: 2.106, size: 44, kicker: 'Discover Your Path to Success',
    parts: [['Benefits', null, { breakLine: true }], ['of Higher '], ['Education', C.orange]]
  });
  badge(s, 10.762, 1.723);
  body(s, LOREM, 0.76, 4.948, 3.439, 1.28);
  iconDisc(s, 0.835, 4.171, 0.673, C.orange, C.white);
  rule(s, 0.845, 6.809);

  [2.119, 4.276].forEach(y => {
    card(s, 4.853, y, 3.624, 1.827, 0.2688);
    iconDisc(s, 5.229, y + 0.273, 0.683, C.white, C.blue);
    T(s, 'Text Here', { x: 6.074, y: y + 0.273, w: 2.508, h: 0.367, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
    body(s, LOREM_SHORT, 6.074, y + 0.617, 2.231, 0.978, { color: C.white });
  });

  card(s, 10.488, 4.37, 1.656, 1.611, 0.1758);
  statBlock(s, { x: 10.555, y: 4.744, vw: 1.521, value: '+46%', align: 'center', cx: 10.66, cy: 5.281, cw: 1.312, capSize: 12 });
  sunDot(s, 6.771, 0.853, 0.601);
});

/* 7 — why choose ---------------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 0, y: 0, w: 5.208, h: 3.588 }, CARD_GRAD, 310,
    { rectRadius: 0.5993, flipV: true, bands: PANEL_BANDS });
  logo(s, 0.823, 0.46, true);
  heading(s, {
    x: 7.715, y: 1.177, w: 5.0, h: 2.106, size: 44, kicker: 'Discover Your Path to Success',
    parts: [['Why Choose', null, { breakLine: true }], ['a '], ['University', C.orange], [' Program?']]
  });
  sunDot(s, 12.178, 0.908, 0.537);

  statBlock(s, {
    x: 1.414, y: 5.132, vw: 1.607, size: 24, value: '+21,5%', color: C.blue,
    cx: 1.407, cy: 5.555, cw: 1.428, capSize: 12, capColor: C.black
  });
  body(s, [
    { text: 'At vero eos et accusamus et iusto', options: { breakLine: true } },
    { text: 'odio dignissimos' }
  ], 1.407, 5.841, 2.08, 0.978);
  iconDisc(s, 0.863, 5.182, 0.395, C.blue, C.white);

  numberedList(s, 7.726, 4.046, [
    {
      num: '01.', w: 3.714, lines: [
        { text: 'PLACEHOLDER', options: { breakLine: true } },
        { text: 'voluptatum deleniti atque corrupti quos' }]
    },
    {
      num: '02.', w: 3.731, lines: [
        { text: 'PLACEHOLDER', options: { breakLine: true } },
        { text: 'voluptatum deleniti atque corrupti quos' }]
    }
  ]);
  rule(s, 12.194, 6.727);

  card(s, 3.375, 3.001, 1.656, 1.611, 0.1758);
  icon(s, 3.89, 3.344, 0.627, C.white);
  T(s, 'Text Here', { x: 3.547, y: 3.943, w: 1.312, h: 0.326, fontSize: 12, italic: true, color: C.white, align: 'center', lineSpacingMultiple: 1.2 });
  badge(s, 5.692, 5.691);
});

/* 8 — a guide to ---------------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 0, y: 3.762, w: 3.832, h: 3.738 }, CARD_GRAD, 130,
    { rectRadius: 0.6439, flipH: true, bands: PANEL_BANDS });
  T(s, [
    { text: 'Public Health', options: { breakLine: true } },
    { text: 'Awareness' }
  ], { x: 8.35, y: 1.834, w: 4.017, h: 1.447, fontFace: F.head, fontSize: 40, color: C.white, lineSpacingMultiple: 1.0 });
  T(s, 'Encouraging Nutritional Education', { x: 8.358, y: 1.43, w: 4.009, h: 0.37, fontSize: 16, color: C.white });
  logo(s);
  heading(s, {
    x: 8.348, y: 1.279, w: 4.124, h: 2.106, size: 44, gap: 0.434, kicker: 'Discover Your Path to Success',
    parts: [['A Guide to '], ['University', C.orange], [' Programs']]
  });
  body(s, 'Choosing a university program is a significant step in shaping your future. This presentation ' +
    'will provide you with a comprehensive overview of different university programs, helping you make an informed decision.',
    8.38, 5.172, 3.798, 1.583);
  iconDisc(s, 8.489, 4.267, 0.673, C.orange, C.white);
  sunDot(s, 12.178, 0.908, 0.537);
  rule(s, 12.194, 6.727);

  body(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui', 1.532, 2.077, 2.202, 0.978);
  T(s, 'Text Here', { x: 1.532, y: 1.716, w: 2.147, h: 0.337, fontSize: 14, bold: true });
  iconDisc(s, 0.692, 1.716, 0.673, C.orange, C.white);

  card(s, 3.489, 3.543, 3.993, 1.342, 0.2237, { shadow: undefined });
  [['86%', 3.903, 3.877], ['35k', 5.638, 5.612]].forEach(([v, vx, cx]) => {
    T(s, v, { x: vx, y: 3.758, w: 1.431, h: 0.64, fontSize: 32, bold: true, color: C.white, align: 'center' });
    T(s, 'Text here ', { x: cx, y: 4.333, w: 1.483, h: 0.337, fontSize: 14, italic: true, color: C.white, align: 'center' });
  });
  sunDot(s, 6.664, 5.799, 0.537);
});

/* 9 — why choose (cards) -------------------------------------------- */
slides.push(s => {
  gradShape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 7.5 }, [[0, C.white], [1, C.orange]], 0, { bands: 54 });
  logo(s);
  heading(s, {
    x: 0.704, y: 1.431, w: 4.124, h: 1.924, size: 40, gap: 0.435, kicker: 'Benefits of Higher Education',
    parts: [['Why Choose a '], ['University', C.orange], [' Program?']]
  });
  [
    { y: 1.184, title: 'Expand your knowledge', body: 'Delve deeper into your chosen field of study.', bw: 2.195 },
    { y: 3.001, title: 'Develop critical skills', body: 'Learn to analyze information and solve problems', bw: 2.518 },
    { y: 4.818, title: 'Enhance career prospects', body: 'Gain the qualifications needed for your desired career.', bw: 2.615 }
  ].forEach(c => iconCard(s, {
    x: 8.514, y: c.y, w: 3.977, h: 1.519, tx: 9.642,
    title: c.title, body: c.body, bw: c.bw
  }));
  body(s, LOREM, 0.748, 4.41, 3.074, 1.583);
  rule(s, 0.845, 6.809);
  badge(s, 4.488, 3.391);
});

/* 10 — types of programs -------------------------------------------- */
slides.push(s => {
  logo(s);
  heading(s, {
    x: 0.704, y: 1.431, w: 4.124, h: 1.5, size: 40, gap: 0.435, kicker: 'Find the Right Fit for You',
    parts: [['Types of '], ['University', C.orange], [' Programs']]
  });
  badge(s, 11.955, 0.906);
  [
    { y: 0.819, title: 'Undergraduate programs', body: "Bachelor's degrees (e.g., BA, BSc)", bw: 2.872 },
    { y: 2.34, title: 'Graduate programs', body: "Master's and doctoral degrees", bw: 2.872 },
    { y: 3.86, title: 'Professional programs', body: 'Law, medicine, engineering', bw: 2.872 },
    { y: 5.381, title: 'Vocational programs', body: 'Certificates and diplomas for  skills', bw: 2.972 }
  ].forEach(c => iconCard(s, {
    x: 4.511, y: c.y, w: 4.314, h: 1.295, tx: 5.639, dy: 0.307, ty: 0.275, by: 0.649, bh: 0.372,
    title: c.title, body: c.body, bw: c.bw
  }));
  body(s, LOREM, 0.748, 4.41, 3.074, 1.583);
  rule(s, 0.845, 6.809);
});

/* 11 — price table --------------------------------------------------- */
slides.push(s => {
  s.addShape('roundRect', {
    x: 0.856, y: 0.797, w: 6.323, h: 5.906, rectRadius: 0.3229, fill: { color: C.white },
    shadow: { type: 'outer', blur: 20, offset: 3, angle: 270, color: C.grey, opacity: 0.1 }
  });

  const cell = (t, o) => Object.assign({ text: t, options: Object.assign({ fontSize: 12, align: 'center', valign: 'middle', fontFace: F.body }, o) });
  const line = (c, pt) => ({ type: 'solid', color: c, pt: pt || 1 });
  const none = { type: 'none' };
  const rows = [
    [cell('Text Here', { color: C.white, fill: { color: C.blue }, border: [line(C.grey, 0.5), none, line(C.white), none] }),
     cell('Text Here', { color: C.white, fill: { color: C.blue }, border: [line(C.grey, 0.5), none, line(C.white), none] })]
  ];
  [['Text Here', '$1,500'], ['Text Here', '$2,500'], ['Text Here', '$1,200'],
   ['Text Here', '$500'], ['Text Here', '$800'], ['Text Here', '$300']].forEach(([a, b], i) => {
    const top = i === 0 ? line(C.white) : line(C.blue);
    const bottom = i === 5 ? line(C.sand) : line(C.blue);
    rows.push([
      cell(a, { color: C.black, fill: { color: C.white }, border: [top, none, bottom, none] }),
      cell(b, { color: C.black, fill: { color: C.white }, border: [top, none, bottom, none] })
    ]);
  });
  rows.push([
    cell('Total', { color: C.white, bold: true, fill: { color: C.orange }, border: [line(C.sand), none, none, none] }),
    cell('$6,300', { color: C.white, fill: { color: C.orange }, border: [line(C.sand), none, none, none] })
  ]);
  s.addTable(rows, {
    x: 1.193, y: 1.118, w: 5.67, colW: [2.939, 2.731],
    rowH: [0.563, 0.563, 0.686, 0.722, 0.722, 0.722, 0.722, 0.563],
    margin: [0.032, 0.064, 0.032, 0.064]
  });

  logo(s, 10.658, 0.521);
  heading(s, {
    x: 7.718, y: 1.25, w: 4.124, h: 1.924, size: 40, gap: 0.434, kicker: 'Benefits of Higher Education',
    parts: [['Why Choose a '], ['University', C.orange], [' Program?']]
  });
  body(s, LOREM + ' excepturi sint occaecati', 7.764, 5.228, 4.061, 1.28);
  sunDot(s, 12.208, 1.684, 0.601);
  rule(s, 12.194, 6.727);
  [['+21,5%', 8.175, 8.168, 7.848, 7.9], ['+41,8%', 10.588, 10.581, 10.261, 10.313]].forEach(([v, vx, cx, ox, ix]) => {
    statBlock(s, { x: vx, y: 4.185, vw: 1.607, size: 24, value: v, color: C.blue, cx, cy: 4.608, cw: 1.428, capSize: 12, capColor: C.black });
    s.addShape('ellipse', { x: ox, y: 4.273, w: 0.292, h: 0.292, fill: { color: C.blue } });
    icon(s, ix, 4.325, 0.188, C.white);
  });
});

/* 12 — donut statistics --------------------------------------------- */
slides.push(s => {
  s.addShape('roundRect', { x: 0.681, y: 3.75, w: 11.656, h: 3.131, rectRadius: 0.3448, fill: { color: C.blue } });
  logo(s);
  [
    { vx: 4.158, cx: 4.151, ox: 3.832, ix: 3.884, bx: 4.153, value: '+21,5%' },
    { vx: 9.747, cx: 9.74, ox: 9.42, ix: 9.472, bx: 9.742, value: '+41,5%' }
  ].forEach(g => {
    statBlock(s, {
      x: g.vx, y: 4.452, vw: 1.607, size: 24, value: g.value, cx: g.cx, cy: 4.875, cw: 1.428,
      capSize: 12, body: LOREM_SHORT, bx: g.bx, by: 5.201, bw: 2.205
    });
    s.addShape('ellipse', { x: g.ox, y: 4.54, w: 0.292, h: 0.292, fill: { color: C.white } });
    icon(s, g.ix, 4.592, 0.188, C.white === C.white ? C.blue : C.blue);
  });

  // ring charts: orange disc, white pie mask, blue hub
  [{ x: 1.426, sweepStart: 1.97, label: '25%', tx: 1.967 },
   { x: 6.938, sweepStart: 68.5, label: '45%', tx: 7.479 }].forEach(r => {
    s.addShape('ellipse', { x: r.x + 0.013, y: 4.294, w: 2.039, h: 2.039, fill: { color: C.orange } });
    s.addShape('pie', { x: r.x, y: 4.29, w: 2.051, h: 2.051, fill: { color: C.white }, angleRange: [r.sweepStart, 270] });
    s.addShape('ellipse', { x: r.x + 0.195, y: 4.485, w: 1.663, h: 1.663, fill: { color: C.blue } });
    T(s, r.label, { x: r.tx, y: 5.128, w: 0.969, h: 0.375, fontSize: 18, bold: true, color: C.white, align: 'center' });
  });

  sunDot(s, 11.96, 0.299, 0.601);
  [[1.84, 2.782, 3.435], [7.162, 8.062, 3.432]].forEach(([ix, tx, tw]) => {
    icon(s, ix, 2.485, 0.711, C.blue);
    body(s, LOREM_SHORT, tx, 2.502, tw, 0.678);
  });
  heading(s, {
    x: 1.599, y: 0.932, w: 10.136, h: 0.774, size: 44, align: 'center', gap: 0.438,
    kicker: 'Discover Your Path to Success', kickerX: 4.761,
    parts: [['Benefits of Higher '], ['Education', C.orange]]
  });
});

/* 13 — application process ------------------------------------------ */
slides.push(s => {
  heading(s, {
    x: 9.017, y: 1.348, w: 3.61, h: 1.924, size: 40, gap: 0.434, kicker: 'Steps to Take',
    parts: [['The Application '], ['Process', C.orange]]
  });
  [
    { y: 5.074, w: 4.01, title: 'Research universities', tw: 2.749, body: 'Explore different institutions and their programs.', bw: 2.749 },
    { y: 1.442, w: 4.994, title: 'Prepare application materials', tw: 3.25, body: 'Write essays, gather transcripts, and request letters of recommendation.', bw: 3.696 },
    { y: 3.258, w: 4.64, title: 'Submit your application', tw: 3.12, body: 'Follow the specific guidelines for each university', bw: 3.414 }
  ].forEach(c => iconCard(s, {
    x: 0.846, y: c.y, w: c.w, h: 1.519, tx: 1.974, title: c.title, tw: c.tw, body: c.body, bw: c.bw
  }));
  body(s, LOREM, 9.059, 4.882, 3.439, 1.28);
  iconDisc(s, 9.134, 4.105, 0.673, C.orange, C.white);
  rule(s, 9.059, 6.632);
  sunDot(s, 12.178, 0.908, 0.537);
  badge(s, 7.576, 1.09);
  logo(s);
});

/* 14 — start your journey ------------------------------------------- */
slides.push(s => {
  gradShape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 7.5 },
    [[0, C.orange], [1, C.white]], 0, { bands: 54 });

  /** white disc + two pie wedges + white hub. */
  const ring = (o) => {
    s.addShape('ellipse', { x: o.x, y: o.y, w: o.d, h: o.d, fill: { color: C.white }, shadow: SOFT_SHADOW });
    const p = (o.d - o.pd) / 2;
    s.addShape('pie', { x: o.x + p, y: o.y + p, w: o.pd, h: o.pd, fill: { color: C.blue }, angleRange: o.blue });
    s.addShape('pie', { x: o.x + p, y: o.y + p, w: o.pd, h: o.pd, fill: { color: C.orange }, angleRange: o.orange });
    const q = (o.d - o.hd) / 2;
    s.addShape('ellipse', { x: o.x + q, y: o.y + q, w: o.hd, h: o.hd, fill: { color: C.white } });
    T(s, o.label, {
      x: o.x + o.d / 2 - 0.6755, y: o.y + o.d / 2 - o.lh / 2, w: 1.351, h: o.lh,
      fontSize: o.ls, bold: true, color: C.darkGrey, align: 'center', valign: 'middle'
    });
  };
  ring({ x: 4.875, y: 1.181, d: 3.121, pd: 2.562, hd: 1.599, blue: [270.3, 352.4], orange: [352.1, 270.8], label: '23%', ls: 24, lh: 0.505 });
  ring({ x: 3.408, y: 3.943, d: 2.275, pd: 1.911, hd: 1.192, blue: [271.3, 112.9], orange: [112.75, 274.2], label: '57%', ls: 20, lh: 0.438 });
  ring({ x: 5.992, y: 4.626, d: 1.791, pd: 1.409, hd: 0.88, blue: [271.3, 112.9], orange: [339.6, 274.2], label: '20%', ls: 16, lh: 0.37 });

  [
    { x: 1.332, y: 1.75, w: 3.0, h: 1.673, r: 0.2231, vx: 1.86, vy: 2.131, vw: 1.549, value: '335+', bx: 1.86, by: 2.568, bw: 2.224, bh: 0.579, text: 'At vero eos et accusamus et iusto odio dignissimos' },
    { x: 0.998, y: 4.092, w: 2.101, h: 2.028, r: 0.2134, vx: 1.382, vy: 4.507, vw: 1.323, value: '565+', bx: 1.382, by: 4.923, bw: 1.475, bh: 0.831, text: 'At vero eos et accusamus et iusto odio dignissimos' },
    { x: 8.334, y: 4.417, w: 3.334, h: 2.0, r: 0.2957, vx: 8.88, vy: 4.824, vw: 1.606, value: '2335+', bx: 8.88, by: 5.262, bw: 2.454, bh: 0.834, text: LOREM_SHORT }
  ].forEach(b => {
    card(s, b.x, b.y, b.w, b.h, b.r, { shadow: SOFT_SHADOW });
    T(s, b.value, { x: b.vx, y: b.vy, w: b.vw, h: 0.457, fontFace: F.head, fontSize: 28, color: C.white, lineSpacingMultiple: 0.7 });
    body(s, b.text, b.bx, b.by, b.bw, b.bh, { fontSize: 10, color: C.white });
  });

  heading(s, {
    x: 8.712, y: 1.348, w: 3.61, h: 2.106, size: 44, gap: 0.434, kicker: 'Steps to Take',
    parts: [['Start Your Journey Today']]
  });
  logo(s, 0.823, 0.46, true);
  sunDot(s, 12.268, 0.879, 0.537);
  badge(s, 6.887, 1.21);
});

/* 15 — four-quadrant diagram ---------------------------------------- */
slides.push(s => {
  heading(s, {
    x: 1.333, y: 1.063, w: 10.667, h: 0.774, size: 44, align: 'center', gap: 0.434,
    kicker: 'Steps to Take', kickerX: 4.761, parts: [['Discover Your Path to Success']]
  });
  logo(s);
  s.addShape('ellipse', { x: 5.619, y: 3.646, w: 2.095, h: 2.095, fill: { color: C.ph } });

  const petals = [
    { x: 6.844, y: 2.803, fill: C.blueLt },
    { x: 6.844, y: 4.713, fill: C.blue, flipV: true },
    { x: 4.685, y: 4.713, fill: C.blueDk, flipH: true, flipV: true },
    { x: 4.685, y: 2.803, fill: C.orange, flipH: true }
  ];
  petals.forEach(p => s.addShape('teardrop', {
    x: p.x, y: p.y, w: 1.804, h: 1.804, fill: { color: p.fill }, flipH: !!p.flipH, flipV: !!p.flipV
  }));
  [[7.056, 3.015], [7.056, 5.023], [4.897, 3.015], [4.897, 5.023]].forEach(([x, y]) => {
    s.addShape('ellipse', { x, y, w: 1.38, h: 1.38, fill: { color: C.white } });
    icon(s, x + 0.444, y + 0.443, 0.493, '231F20');
  });

  [
    { hx: 9.102, hy: 3.029, bx: 9.102, by: 3.257, align: 'left' },
    { hx: 9.102, hy: 5.102, bx: 9.102, by: 5.329, align: 'left' },
    { hx: 2.464, hy: 3.029, bx: 0.734, by: 3.257, align: 'right' },
    { hx: 2.464, hy: 5.102, bx: 0.734, by: 5.329, align: 'right' }
  ].forEach(t => {
    T(s, 'Text Here', { x: t.hx, y: t.hy, w: 1.748, h: 0.337, fontFace: F.head, fontSize: 14, bold: true, align: t.align });
    body(s, LOREM_SHORT, t.bx, t.by, 3.477, 0.675, { align: t.align });
  });

  [
    { x: 9.225, y: 4.051, w: 1.205, fill: C.blueLt, lx: 9.347, ly: 4.068 },
    { x: 9.213, y: 6.108, w: 1.216, fill: C.blue, lx: 9.342, ly: 6.125 },
    { x: 2.904, y: 6.117, w: 1.216, fill: C.blueDk, lx: 3.032, ly: 6.134 },
    { x: 2.904, y: 4.048, w: 1.205, fill: C.orange, lx: 3.026, ly: 4.064 }
  ].forEach(p => {
    s.addShape('roundRect', { x: p.x, y: p.y, w: p.w, h: 0.337, rectRadius: 0.0562, fill: { color: p.fill } });
    T(s, 'Text Here', { x: p.lx, y: p.ly, w: 0.96, h: 0.303, fontSize: 12, color: C.white, align: 'center' });
  });
});

/* 16 — find the right fit ------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 0.839, y: 1.877, w: 3.266, h: 3.186 }, CARD_GRAD, 130,
    { rectRadius: 0.5352, flipH: true, bands: PANEL_BANDS });
  // laptop photo: dark bezel around a white screen, plus the base strip
  s.addShape('rect', { x: 1.781, y: 2.65, w: 5.493, h: 3.75, fill: { color: '2A2A2C' } });
  photo(s, { x: 1.948, y: 2.883, w: 5.16, h: 3.25, color: C.white });
  s.addShape('roundRect', { x: 1.132, y: 6.46, w: 6.79, h: 0.15, rectRadius: 0.075, fill: { color: 'B4B4B6' } });
  logo(s);
  heading(s, {
    x: 8.714, y: 1.39, w: 3.841, h: 1.924, size: 40, gap: 0.434, kicker: 'The Application Process',
    parts: [['Find the '], ['Right', C.orange], [' Fit for You']]
  });
  sunDot(s, 11.825, 0.723, 0.537);
  body(s, LOREM, 8.735, 3.789, 3.942, 1.28);
  T(s, '+21,5%', {
    x: 9.158, y: 5.312, w: 1.607, h: 0.468, fontFace: F.head, fontSize: 24,
    bold: true, color: C.blue, lineSpacingMultiple: 0.9
  });
  body(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui', 9.151, 5.735, 2.942, 0.675);
  s.addShape('ellipse', { x: 8.832, y: 5.4, w: 0.292, h: 0.292, fill: { color: C.blue } });
  icon(s, 8.884, 5.452, 0.188, C.white);

  card(s, 4.505, 1.38, 3.624, 1.827, 0.2688);
  iconDisc(s, 4.881, 1.653, 0.683, C.white, C.blue);
  T(s, 'Text Here', { x: 5.725, y: 1.653, w: 2.508, h: 0.367, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  body(s, LOREM_SHORT, 5.725, 1.996, 2.231, 0.978, { color: C.white });
});

/* 17 — start your journey (freeform) --------------------------------- */
slides.push(s => {
  // large orange free-form panel: straight right edge, bowed left edge
  s.addShape('custGeom', {
    x: 5.865, y: 0, w: 7.469, h: 7.5, fill: { color: C.orange },
    points: [
      { x: 1.389, y: 0, moveTo: true },
      { x: 7.469, y: 0 },
      { x: 7.469, y: 7.5 },
      { x: 1.322, y: 7.5 },
      { x: 1.166, y: 7.301 },
      { x: 0, y: 3.79, curve: { type: 'cubic', x1: 0.434, y1: 6.322, x2: 0, y2: 5.106 } },
      { x: 1.34, y: 0.057, curve: { type: 'cubic', x1: 0, y1: 2.371, x2: 0.503, y2: 1.071 } },
      { close: true }
    ]
  });
  logo(s);
  heading(s, {
    x: 8.361, y: 1.406, w: 3.841, h: 2.106, size: 44, gap: 0.434, light: true,
    kicker: 'Take the First Step', parts: [['Start Your Journey Today']]
  });
  body(s, 'Explore your options, create a shortlist for a better research, if possible talk to advisor ' +
    'and visits university to make good comparison.', 8.361, 5.148, 3.841, 0.978, { color: C.white });
  icon(s, 8.469, 4.459, 0.413, C.white);
  badge(s, 5.942, 0.798);
  rule(s, 12.194, 6.727);

  // arch-shaped image placeholder inherited from the layout (semicircular top)
  s.addShape('ellipse', { x: 0.854, y: 1.13, w: 4.632, h: 4.632, fill: { color: C.arch } });
  s.addShape('rect', { x: 0.854, y: 3.446, w: 4.632, h: 2.426, fill: { color: C.arch } });
  T(s, 'image place holder', { x: 0.854, y: 3.3, w: 4.632, h: 0.4, fontSize: 18, align: 'center', valign: 'middle' });

  card(s, 1.859, 3.93, 2.993, 2.809, 0.3266);
  statBlock(s, { x: 2.258, y: 4.375, value: '+46,8%', cx: 2.267, cy: 4.972, body: LOREM_SHORT, bx: 2.267, by: 5.316 });
  sunDot(s, 1.222, 1.442, 0.601);
});

/* 18 — service (phone) ----------------------------------------------- */
slides.push(s => {
  gradShape(s, 'round1Rect', { x: 3.462, y: 1.567, w: 3.393, h: 3.31 }, CARD_GRAD, 50, { rectRadius: 0.5562, bands: PANEL_BANDS });
  // phone photo: only the dark bezel is opaque, the screen shows the panel behind
  s.addShape('roundRect', { x: 2.47, y: 2.46, w: 3.3, h: 5.04, rectRadius: 0.42, line: { color: '202020', width: 5 } });
  s.addShape('roundRect', { x: 3.65, y: 2.46, w: 0.95, h: 0.28, rectRadius: 0.11, fill: { color: '202020' } });
  T(s, '[image]', { x: 2.47, y: 6.4, w: 3.3, h: 0.4, fontSize: 12, color: C.grey, align: 'center' });
  card(s, 4.853, 0.788, 1.656, 1.611, 0.1758);
  icon(s, 5.368, 1.131, 0.627, C.white);
  T(s, 'Text Here', { x: 5.026, y: 1.73, w: 1.312, h: 0.326, fontSize: 12, italic: true, color: C.white, align: 'center', lineSpacingMultiple: 1.2 });
  logo(s);
  heading(s, {
    x: 7.718, y: 1.094, w: 4.124, h: 2.106, size: 44, kicker: 'Discover Your Path to Success',
    parts: [['Service University '], ['Program', C.orange]]
  });
  badge(s, 1.426, 1.729);
  sunDot(s, 12.186, 0.947, 0.601);
  numberedList(s, 7.726, 4.046, [
    {
      num: '01.', w: 3.714, lines: [
        { text: 'PLACEHOLDER', options: { breakLine: true } },
        { text: 'voluptatum deleniti atque corrupti quos' }]
    },
    {
      num: '02.', w: 3.731, lines: [
        { text: 'PLACEHOLDER', options: { breakLine: true } },
        { text: 'voluptatum deleniti atque corrupti quos' }]
    }
  ]);
  rule(s, 12.194, 6.727);

  card(s, 0.838, 3.778, 2.993, 2.45, 0.2673);
  statBlock(s, {
    x: 1.237, y: 4.196, value: '+46,8%', cx: 1.246, cy: 4.793,
    body: 'At vero eos et accusamus et iusto odio dignissimos', bx: 1.246, by: 5.137, bh: 0.675
  });
});

/* 19 — contact -------------------------------------------------------- */
slides.push(s => {
  s.addShape('roundRect', { x: 1.52, y: 4.807, w: 9.321, h: 2.17, rectRadius: 1.085, fill: { color: C.blue } });
  heading(s, {
    x: 1.333, y: 0.924, w: 10.667, h: 0.774, size: 44, align: 'center', gap: 0.434,
    kicker: "Let's Create a Better Future Together", kickerX: 4.761, kickerAlign: 'left',
    parts: [['Contact us For More '], ['Information', C.orange]]
  });
  logo(s);
  [
    { title: 'Address', hx: 2.355, bx: 2.38, bw: 2.535, ix: 2.476, iy: 5.22, iw: 0.249, ih: 0.257, lines: ['1234, Name Street, Building, Your City, Your Country.'] },
    { title: 'Phone', hx: 5.248, bx: 5.273, bw: 1.72, ix: 5.315, iy: 5.167, iw: 0.363, ih: 0.363, lines: ['+00 123 4567 890', '+12 345 6789'] },
    { title: 'Mail', hx: 7.272, bx: 7.297, bw: 1.941, ix: 7.393, iy: 5.254, iw: 0.249, ih: 0.19, lines: ['yourinfo@email.com', 'yourmail@mail.co '] }
  ].forEach(col => {
    s.addShape('roundRect', { x: col.ix, y: col.iy, w: col.iw, h: col.ih, rectRadius: 0.04, fill: { color: C.white } });
    T(s, col.title, { x: col.hx, y: 5.647, w: 1.583, h: 0.37, fontFace: F.head, fontSize: 16, bold: true, color: C.white });
    body(s, col.lines.map((l, i) => ({ text: l, options: { breakLine: i < col.lines.length - 1 } })),
      col.bx, 5.938, col.bw, 0.678, { italic: true, color: C.white });
  });
  badge(s, 1.091, 5.465);
  sunDot(s, 11.805, 0.473, 0.301);
  rule(s, 12.194, 6.727);

  s.addShape('roundRect', { x: 9.313, y: 4.006, w: 3.334, h: 2.0, rectRadius: 0.2957, fill: { color: C.orange }, shadow: SOFT_SHADOW });
  T(s, '2335/', { x: 9.859, y: 4.414, w: 1.606, h: 0.457, fontFace: F.head, fontSize: 28, color: C.white, lineSpacingMultiple: 0.7 });
  body(s, LOREM_SHORT, 9.859, 4.851, 2.454, 0.834, { fontSize: 10, color: C.white });
  T(s, 'Text Here', { x: 11.044, y: 4.409, w: 1.012, h: 0.367, fontSize: 14, italic: true, color: C.white, lineSpacingMultiple: 1.2 });
});

/* 20 — thank you ------------------------------------------------------ */
slides.push(s => {
  gradShape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 7.5 },
    [[0, 'FFF7F1'], [0.5, C.orange]], 0, { bands: 54 });
  heading(s, {
    x: 7.731, y: 1.219, w: 3.123, h: 1.927, size: 60, gap: 0.384, light: true,
    kicker: "Let's Create a Better Future Together", parts: [['Thank You']]
  });
  body(s, LOREM + ' excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia',
    7.772, 3.849, 4.482, 1.583, { color: C.white });
  T(s, 'WWW. YOURWEBSITE. COM', {
    x: 7.77, y: 6.312, w: 3.29, h: 0.372, color: C.white, charSpacing: 3, lineSpacingMultiple: 1.5
  });
  s.addShape('donut', { x: 6.028, y: 5.286, w: 1.287, h: 1.287, fill: { color: C.blue } });
  gradShape(s, 'ellipse', { x: 1.351, y: 0.756, w: 5.956, h: 5.956 },
    [[0, 'FF8F31'], [1, 'FFB170']], 50, { bands: 24 });
  logo(s);
  rule(s, 0.845, 6.809);
  badge(s, 1.458, 1.812);
});

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
pptx.author = 'pptxgenjs';
pptx.title = 'University Program';

slides.forEach(build => {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '03206ac7-d0a2-4cc2-8a0c-0a5d590c9431_grok_final.pptx') })
  .then(f => console.log('wrote', f));
