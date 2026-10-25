/**
 * "Pharma Medical" deck — rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9).  Run: node <thisfile>.js
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette
   Theme colours plus the luminance-shifted tones the original deck derives
   from them (tx1 lumMod/lumOff, accentN lumMod, bg1 lumMod ...).           */
const C = {
  black: '000000',
  white: 'FFFFFF',
  ink05: '0D0D0D', // tx1 lum 95%
  ink15: '262626', // tx1 lum 85%
  ink25: '404040', // tx1 lum 75%
  gray35: '595959', // tx1 lum 65%
  gray50: '808080', // tx1 lum 50%
  line15: 'D9D9D9', // bg1 lum 85%
  line25: 'BFBFBF', // bg1 lum 75%
  line35: 'A6A6A6', // bg1 lum 65%
  wash: 'F2F2F2', // bg1 lum 95%
  violet: 'AEA5F4', // accent1
  violetDk: '8D81F0', // accent1 lum 90%
  violetInk: '2816B6', // accent1 lum 50%
  sand: 'F8D98C', // accent2
  blue: 'BCD1FC', // accent3
  blueDk: '5389F7', // accent3 lum 75%
  blueInk: '094BD3', // accent3 lum 50%
  brown: '964A27', // accent4
  lime: 'E8EE62', // accent5
  limeInk: '929810', // accent5 lum 50%
  peach: 'F3AE80', // accent6
  limeWash: 'F2F6A7', // soft accent5 area fill
  violetWash: 'D3CEF9' // soft accent1 area fill
};

const MAJOR = 'Bricolage Grotesque 24pt Medium';
const MAJOR_LIGHT = 'Bricolage Grotesque Light';
const MINOR = 'Inter';

/* Filler copy reused all over the original template. */
const L = {
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  mid: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
  long: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
  full: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis',
  praesent: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent commodo cursus magna, vel scelerisque nisl consectetur et. Donec sed odio dui. Nulla vitae elit libero, a pharetra augue. Nullam id',
  praesentShort: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent commodo cursus magna, vel scelerisque nisl',
  dot: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '
};

const SOFT_SHADOW = { type: 'outer', blur: 40, offset: 10, angle: 90, color: C.black, opacity: 0.15, rotateWithShape: false };
const WIDE_SHADOW = { type: 'outer', blur: 76, offset: 13, angle: 90, color: C.black, opacity: 0.15, rotateWithShape: false };

/* ----------------------------------------------------------------- helpers */

// `roundRect` corner radius in inches from the OOXML `adj` value of the source.
const radius = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/** Plain text box: top anchored, no fill, like every TextBox in the deck. */
function tx(slide, text, o) {
  slide.addText(text, Object.assign({ valign: 'top', fontFace: MINOR, fontSize: 12, color: C.gray50 }, o));
}

/** Display type: major face, 80% line spacing. */
function head(slide, text, o) {
  tx(slide, text, Object.assign({ fontFace: MAJOR, color: C.ink15, lineSpacingMultiple: 0.8 }, o));
}

/** Body copy: minor face, 130% line spacing. */
function body(slide, text, o) {
  tx(slide, text, Object.assign({ fontSize: 12, color: C.gray50, lineSpacingMultiple: 1.3 }, o));
}

/** Rounded "pill" button/tag. */
function pill(slide, text, o) {
  const w = o.w, h = o.h;
  slide.addText(text, Object.assign(
    {
      shape: 'roundRect',
      rectRadius: radius(50000, w, h),
      fill: { color: o.fill || C.white },
      line: o.line === null ? { type: 'none' } : { color: o.line || C.ink25, width: 0.5 },
      valign: 'middle',
      align: o.align || 'left',
      fontFace: o.fontFace || MINOR,
      fontSize: o.fontSize || 12,
      color: o.color || C.ink25
    },
    o
  ));
}

/** Small solid dot used as a bullet inside the tag pills. */
function dot(slide, x, y, d, color) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' } });
}

/**
 * Arrow glyph (the deck's recurring "任意形状" freeform): a straight shaft that
 * ends in an L-shaped head, drawn as three strokes inside a square box.
 * dir: 'ne' | 'se' | 'sw' | 'nw' | 'e'
 */
const ARROW_DIRS = {
  ne: { tail: [0, 1], tip: [1, 0], arms: [[-1, 0], [0, 1]], head: 0.9 },
  nw: { tail: [1, 1], tip: [0, 0], arms: [[1, 0], [0, 1]], head: 0.9 },
  sw: { tail: [1, 0], tip: [0, 1], arms: [[1, 0], [0, -1]], head: 0.9 },
  se: { tail: [0, 0], tip: [1, 1], arms: [[-1, 0], [0, -1]], head: 0.9 },
  e: { tail: [0, 0.5], tip: [1, 0.5], arms: [[-0.42, -0.42], [-0.42, 0.42]], head: 1 }
};
function arrow(slide, o) {
  const g = ARROW_DIRS[o.dir || 'ne'];
  const s = o.w, head = s * g.head;
  const stroke = { color: o.color || C.ink15, width: o.width || Math.max(0.75, 11 * s) };
  const seg = (x1, y1, x2, y2) => slide.addShape('line', {
    x: o.x + Math.min(x1, x2), y: o.y + Math.min(y1, y2),
    w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1, line: stroke
  });
  const [tx0, ty0] = [g.tail[0] * s, g.tail[1] * s];
  const [px, py] = [g.tip[0] * s, g.tip[1] * s];
  seg(tx0, ty0, px, py);
  g.arms.forEach(a => seg(px, py, px + a[0] * head, py + a[1] * head));
}

/** Circle + arrow badge (coloured disc with the arrow glyph inside). */
function arrowBadge(slide, o) {
  slide.addShape('ellipse', {
    x: o.x, y: o.y, w: o.d, h: o.d,
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lineWidth || 0.5 } : { type: 'none' },
    shadow: o.shadow
  });
  const s = o.d * 0.356;
  arrow(slide, { x: o.x + (o.d - s) / 2, y: o.y + (o.d - s) / 2, w: s, dir: o.dir, color: o.arrow || C.ink15, width: o.arrowWidth });
}

/** Stethoscope mark used in the dashed medallions. */
function stethoscope(slide, x, y, s, color) {
  const stroke = { color, width: Math.max(1, s * 5) };
  slide.addShape('arc', { x, y: y - s * 0.1, w: s * 0.62, h: s * 0.75, line: stroke, rotate: 180 });
  slide.addShape('line', { x: x + s * 0.31, y: y + s * 0.5, w: s * 0.3, h: s * 0.28, line: stroke });
  slide.addShape('ellipse', { x: x + s * 0.5, y: y + s * 0.72, w: s * 0.3, h: s * 0.3, fill: { type: 'none' }, line: stroke });
  slide.addShape('ellipse', { x: x + s * 0.72, y: y + s * 0.06, w: s * 0.2, h: s * 0.2, fill: { color }, line: { type: 'none' } });
}

/**
 * Stand-in for a raster picture: soft grey rounded rectangle with a caption.
 * Empty picture placeholders of the original template render as nothing, so
 * only real pictures are substituted here.
 */
function imagePlaceholder(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate || 0,
    rectRadius: o.rectRadius || 0.15,
    fill: { color: o.fill || 'E9E9EC' },
    line: { color: o.line || C.line15, width: 1 },
    shadow: o.shadow
  });
  slide.addText(o.label || '[image]', {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate || 0,
    align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 12, color: o.labelColor || C.line35
  });
}

/** Master furniture: header wordmark, copyright line and page number. */
function chrome(slide, pageNo, numColor) {
  slide.addText(
    [
      { text: 'Pharma ', options: { fontFace: MAJOR, color: C.ink15 } },
      { text: 'Presentation', options: { fontFace: MAJOR_LIGHT, color: C.line35 } }
    ],
    { x: 0.431, y: 0.331, w: 1.833, h: 0.268, valign: 'top', fontSize: 11, lineSpacingMultiple: 0.9 }
  );
  slide.addText('Copyright © 2024', {
    x: 0.482, y: 7.021, w: 2.583, h: 0.278, valign: 'bottom', fontFace: MINOR, fontSize: 10, color: C.line25
  });
  slide.addText(String(pageNo), {
    x: 11.705, y: 7.065, w: 1.3, h: 0.269, valign: 'top', align: 'right',
    fontFace: MINOR, fontSize: 10, color: numColor || C.line35
  });
}

/* ------------------------------------------------------------------ slides */

// 1 — cover: oversized PHARMA / MEDICAL wordmarks and three tag pills.
function slide01(s) {
  head(s, 'PHARMA', { x: 0.511, y: 0.596, w: 9.938, h: 2.406, fontSize: 166 });
  head(s, 'MEDICAL', { x: 3.066, y: 5.333, w: 9.938, h: 2.017, fontSize: 138, align: 'right' });
  body(s, L.short.replace('tempor', 'tempor incididunt ut labore et'), { x: 8.201, y: 3.255, w: 3.989, h: 0.99, fontSize: 14 });

  [
    { x: 0.822, y: 5.658, w: 1.573, text: '      Healthier', fill: C.white, dotX: 0.968, dotY: 5.805, dotColor: C.lime },
    { x: 0.822, y: 6.117, w: 1.044, text: '     Drugs', fill: C.white, dotX: 0.968, dotY: 6.255, dotColor: C.violet },
    { x: 1.901, y: 6.117, w: 1.044, text: '     Medic', fill: C.blue, dotX: 2.048, dotY: 6.255, dotColor: C.ink15 }
  ].forEach(t => {
    pill(s, t.text, { x: t.x, y: t.y, w: t.w, h: 0.396, fill: t.fill });
    dot(s, t.dotX, t.dotY, 0.13, t.dotColor);
  });

  arrowBadge(s, { x: 10.624, y: 0.955, d: 1.166, line: C.ink15, lineWidth: 1, dir: 'sw', arrow: C.ink15 });
}

// 2 — split intro: big statement left, three plus-marked features right.
function slide02(s) {
  head(s, 'Innovate Science Pharma', { x: 0.379, y: 2.787, w: 4.051, h: 2.795, fontSize: 66 });

  [
    { x: 0.542, w: 1.019, label: 'Drugs', dotX: 0.725, dotColor: C.brown },
    { x: 1.558, w: 1.253, label: 'Medicine', dotX: 1.742, dotColor: C.violet }
  ].forEach(t => {
    pill(s, [
      { text: '     ', options: { fontSize: 12 } },
      { text: t.label, options: { fontSize: 11 } }
    ], { x: t.x, y: 1.918, w: t.w, h: 0.388 });
    dot(s, t.dotX, 2.059, 0.106, t.dotColor);
  });

  [
    { y: 1.523, title: 'Medical Capsule' },
    { y: 3.281, title: 'Pharma Healthier' },
    { y: 5.025, title: 'Medic Science ' }
  ].forEach(f => {
    s.addShape('mathPlus', { x: 8.746, y: f.y + 0.052, w: 0.314, h: 0.314, fill: { color: C.white }, line: { type: 'none' } });
    tx(s, f.title, { x: 9.206, y: f.y, w: 3.444, h: 0.464, fontFace: MAJOR, fontSize: 24, color: C.ink05, lineSpacingMultiple: 0.9 });
    body(s, L.short, { x: 9.231, y: f.y + 0.475, w: 3.153, h: 0.559, fontSize: 11, color: C.gray35 });
  });
}

// 3 — lime capsule panel with headline, feature block and CTA.
function slide03(s) {
  s.addShape('roundRect', {
    x: 1.677, y: 1.089, w: 3.585, h: 5.322, rectRadius: radius(50000, 3.585, 5.322),
    fill: { color: C.lime }, line: { type: 'none' }
  });
  head(s, 'Building a Healthy Tomorrow', { x: 5.945, y: 1.639, w: 6.416, h: 1.578, fontSize: 52, color: C.ink05 });
  s.addShape('mathPlus', { x: 6.019, y: 3.676, w: 0.314, h: 0.314, fill: { color: C.blue }, line: { type: 'none' } });
  tx(s, 'Pharma Healthier', { x: 6.479, y: 3.624, w: 3.444, h: 0.464, fontFace: MAJOR, fontSize: 24, color: C.ink25, lineSpacingMultiple: 0.9 });
  body(s, L.full, { x: 6.479, y: 4.113, w: 5.66, h: 0.99, fontSize: 14 });
  pill(s, ' Get Appointment', { x: 6.495, y: 5.382, w: 1.763, h: 0.479, fill: C.blue });

  // dashed medallion with the stethoscope mark
  s.addShape('ellipse', { x: 1.135, y: 4.65, w: 1.333, h: 1.333, fill: { color: C.white }, line: { type: 'none' }, shadow: WIDE_SHADOW });
  s.addShape('ellipse', { x: 1.258, y: 4.774, w: 1.086, h: 1.086, fill: { type: 'none' }, line: { color: C.blueDk, width: 1.25, dashType: 'dash' } });
  stethoscope(s, 1.56, 5.06, 0.45, C.blueDk);
}

// 4 — three service bands, the middle one inverted on violet.
function slide04(s) {
  head(s, 'Our Service Pharma', { x: 2.59, y: 0.603, w: 8.154, h: 0.767, fontSize: 48, align: 'center', color: C.ink05 });

  [
    { y: -2.523, fill: C.wash },
    { y: -0.788, fill: C.violet },
    { y: 0.946, fill: C.wash }
  ].forEach(b => {
    s.addShape('roundRect', {
      x: 5.835, y: b.y, w: 1.582, h: 10.143, rotate: 90,
      rectRadius: radius(10622, 1.582, 10.143), fill: { color: b.fill }, line: { type: 'none' }
    });
  });

  const rows = [
    { num: '01/', numColor: C.line35, title: 'Pharma Healthier', titleColor: C.ink15, copyColor: C.gray35,
      numX: 1.743, numY: 1.998, titleX: 2.656, titleY: 2.193, titleW: 1.79, copyX: 5.954, copyY: 2.106, copyW: 2.888, align: 'left' },
    { num: '02/', numColor: C.violetInk, title: 'Medical Make Drug', titleColor: C.violetInk, copyColor: C.white,
      numX: 10.565, numY: 3.731, titleX: 8.289, titleY: 3.905, titleW: 2.002, copyX: 4.511, copyY: 3.852, copyW: 2.708, align: 'right' },
    { num: '03/', numColor: C.line35, title: 'Innovation Medicine ', titleColor: C.ink15, copyColor: C.gray35,
      numX: 1.743, numY: 5.464, titleX: 2.656, titleY: 5.66, titleW: 1.999, copyX: 5.954, copyY: 5.586, copyW: 2.888, align: 'left' }
  ];
  rows.forEach(r => {
    head(s, r.num, { x: r.numX, y: r.numY, w: 0.884, h: 0.434, fontSize: 24, color: r.numColor, align: r.align });
    head(s, r.title, { x: r.titleX, y: r.titleY, w: r.titleW, h: 0.757, fontSize: 24, color: r.titleColor, align: r.align });
    body(s, L.short, { x: r.copyX, y: r.copyY, w: r.copyW, h: 0.863, color: r.copyColor, align: r.align });
  });
}

// 5 — "Medic Pharma" gallery row with two captions and a Learn More CTA.
function slide05(s) {
  head(s, 'Medic Pharma', { x: 0.408, y: 0.87, w: 3.662, h: 1.742, fontSize: 60 });
  s.addShape('roundRect', {
    x: 0.542, y: 3.86, w: 2.769, h: 2.949, rectRadius: radius(2489, 2.769, 2.949),
    fill: { color: C.blue }, line: { type: 'none' }
  });
  head(s, 'Doctor Approved', { x: 0.703, y: 4.07, w: 1.244, h: 0.539, fontSize: 16, color: C.blueDk });
  arrowBadge(s, { x: 0.736, y: 6.289, d: 0.344, fill: C.white, dir: 'ne', arrow: C.blueDk });

  [
    { x: 4.588, copyX: 4.613, title: 'Global Health' },
    { x: 9.496, copyX: 9.521, title: 'Medical Capsule' }
  ].forEach(c => {
    tx(s, c.title, { x: c.x, y: 5.712, w: 1.917, h: 0.343, fontFace: MAJOR, fontSize: 16, color: C.ink15, lineSpacingMultiple: 0.9 });
    body(s, L.mid, { x: c.copyX, y: 6.087, w: 3.271, h: 0.8, fontSize: 11 });
  });

  pill(s, [
    { text: '     ', options: { fontSize: 12, color: C.ink25 } },
    { text: 'Learn More', options: { fontSize: 11, color: C.ink15 } }
  ], { x: 10.866, y: 4.274, w: 1.459, h: 0.388, fill: C.blue });
  dot(s, 11.049, 4.415, 0.106, C.white);
  arrowBadge(s, { x: 12.414, y: 4.279, d: 0.378, fill: C.blue, line: C.ink25, dir: 'ne', arrow: C.white });
}

// 6 — stat + headline layout ("It's More Drug Therapies").
function slide06(s) {
  head(s, 'It’s More Drug Therapies', { x: 1.283, y: 3.875, w: 6.5, h: 1.791, fontSize: 62 });
  tx(s, '250K', { x: 8.496, y: 0.949, w: 1.762, h: 0.707, fontFace: MAJOR, fontSize: 40, color: C.violetDk, lineSpacingMultiple: 0.9 });
  body(s, 'Lives with Innovative Healthcare Solutions', { x: 8.496, y: 1.579, w: 3.896, h: 0.338 });
  tx(s, 'Efficacy in Every Dose', { x: 8.708, y: 2.65, w: 1.509, h: 1.101, fontFace: MAJOR, fontSize: 22, color: C.ink15, lineSpacingMultiple: 0.9 });
  body(s, L.long, { x: 2.089, y: 5.879, w: 5.378, h: 0.601 });
  s.addShape('mathPlus', { x: 1.461, y: 6.017, w: 0.311, h: 0.325, fill: { color: C.violetDk }, line: { type: 'none' } });
  arrowBadge(s, { x: 11.505, y: 5.812, d: 0.509, fill: C.violetDk, dir: 'nw', arrow: C.white });
}

// 7 — 2x2 product grid, each cell captioned with a "Capsule" tag.
function slide07(s) {
  head(s, 'The Role of Pharma Today', { x: 0.8, y: 0.953, w: 11.733, h: 0.934, fontSize: 60, align: 'center', color: C.ink05 });
  [
    { title: 'Immune Boost', tx: 1.26, ty: 2.45, px: 4.977, py: 3.885 },
    { title: 'Probiotic Acid', tx: 7.012, ty: 2.45, px: 10.717, py: 3.885 },
    { title: 'Collagen Glow', tx: 1.26, ty: 4.822, px: 4.977, py: 6.257 },
    { title: 'Ginseng Vital', tx: 7.012, ty: 4.822, px: 10.717, py: 6.257 }
  ].forEach(c => {
    head(s, c.title, { x: c.tx, y: c.ty, w: 1.398, h: 0.648, fontSize: 20 });
    pill(s, 'Capsule', { x: c.px, y: c.py, w: 1.456, h: 0.356, fill: C.white, line: null, align: 'center', fontFace: MAJOR, color: C.ink15 });
  });
}

// 8 — hero card over a two-row KPI table.
function slide08(s) {
  head(s, 'Enhance Pharma Efficiency', { x: 0.542, y: 0.862, w: 3.939, h: 2.06, fontSize: 48 });
  s.addShape('roundRect', {
    x: 4.987, y: 0.931, w: 3.146, h: 2.107, rectRadius: radius(7271, 3.146, 2.107),
    fill: { color: C.blue }, line: { type: 'none' }
  });
  body(s, L.full.replace(' veniam, quis', ''), { x: 9.132, y: 0.964, w: 3.457, h: 1.126 });
  arrowBadge(s, { x: 7.701, y: 2.693, d: 0.688, fill: C.white, dir: 'ne', arrow: C.ink25, shadow: SOFT_SHADOW });

  [3.892, 5.203, 6.503].forEach(y => {
    s.addShape('line', { x: 0, y, w: 13.333, h: 0, line: { color: C.line25, width: 1, transparency: 27 } });
  });

  [
    { kpi: '200K', kpiW: 1.958, kpiY: 4.329, label: 'Healthier Lives, Brighter Futures', labelW: 2.944, labelY: 4.213, badgeY: 4.375, dotY: 4.5, fill: C.violet },
    { kpi: '150K', kpiW: 1.903, kpiY: 5.64, label: 'Quality Control in Pharma', labelW: 2.329, labelY: 5.524, badgeY: 5.719, dotY: 5.845, fill: C.lime }
  ].forEach(r => {
    head(s, r.kpi, { x: 0.542, y: r.kpiY, w: r.kpiW, h: 0.49, fontSize: 28 });
    tx(s, r.label, { x: 6.001, y: r.labelY, w: r.labelW, h: 0.707, fontFace: MAJOR, fontSize: 20, color: C.ink15, lineSpacingMultiple: 0.9 });
    s.addShape('roundRect', {
      x: 12.001, y: r.badgeY, w: 0.79, h: 0.324, rectRadius: radius(50000, 0.79, 0.324),
      fill: { color: r.fill }, line: { color: C.ink25, width: 0.5 }
    });
    [12.21, 12.36, 12.51].forEach(x => dot(s, x, r.dotY, 0.074, C.white));
  });
}

// 9 — three solution cards with pill captions.
function slide09(s) {
  head(s, 'Pharmaceutical Our Solutions', { x: 0.542, y: 0.951, w: 11.247, h: 0.851, fontSize: 54 });
  [
    { x: 0.633, fill: C.blue, title: 'Immune Boost', titleX: 0.96 },
    { x: 8.785, fill: C.lime, title: 'Probiotic Acid', titleX: 9.018 }
  ].forEach(card => {
    s.addShape('roundRect', {
      x: card.x, y: 2.501, w: 3.823, h: 4.076, rectRadius: radius(6697, 3.823, 4.076),
      fill: { color: card.fill }, line: { type: 'none' }
    });
    head(s, card.title, { x: card.titleX, y: 2.806, w: 1.723, h: 0.757, fontSize: 24 });
  });
  [
    { x: 0.872, text: 'Healthier Lives, Brighter Futures' },
    { x: 4.928, text: 'Quality Control in Pharma' },
    { x: 9.004, text: 'Enhance Pharma Efficiency' }
  ].forEach(p => {
    pill(s, p.text, { x: p.x, y: 5.933, w: 3.384, h: 0.478, fill: C.white, line: null, align: 'center', fontFace: MAJOR, color: C.ink15 });
  });
}

// 10 — violet feature card plus two bulleted list rows.
function slide10(s) {
  s.addShape('roundRect', {
    x: 0.822, y: 0.867, w: 3.244, h: 2.767, rectRadius: radius(7067, 3.244, 2.767),
    fill: { color: C.violet }, line: { type: 'none' }
  });
  tx(s, 'Best Capsule', { x: 0.991, y: 1.07, w: 1.292, h: 0.646, fontFace: MAJOR, fontSize: 18, color: C.white, lineSpacingMultiple: 0.9 });
  arrowBadge(s, { x: 1.095, y: 2.982, d: 0.415, fill: C.white, line: C.ink25, dir: 'ne', arrow: C.ink15 });
  body(s, L.long, { x: 4.522, y: 0.99, w: 2.917, h: 1.126, color: C.gray35 });

  const bullets = [
    { text: 'Lorem ipsum dolor sit amet,', options: { bullet: { characterCode: '25B8', indent: 18.25 }, paraSpaceAfter: 4 } },
    { text: 'consectetur adipiscing elit, ', options: { bullet: { characterCode: '25B8', indent: 18.25 }, paraSpaceAfter: 4, breakLine: true } }
  ];
  [
    { y: 4.28, ty: 4.3, title: 'Global Health', titleW: 1.917, by: 4.225, disc: C.peach },
    { y: 5.736, ty: 5.756, title: 'Medical Capsule', titleW: 2.416, by: 5.681, disc: C.ink15 }
  ].forEach(r => {
    arrowBadge(s, { x: 0.754, y: r.y, d: 0.377, line: C.ink15, lineWidth: 1, dir: 'e', arrow: C.ink15 });
    tx(s, r.title, { x: 1.295, y: r.ty, w: r.titleW, h: 0.374, fontFace: MAJOR, fontSize: 18, color: C.ink15, lineSpacingMultiple: 0.9 });
    tx(s, bullets, { x: 4.503, y: r.by, w: 2.708, h: 0.657, fontSize: 12, color: C.gray35, lineSpacingMultiple: 1.3 });
  });

  s.addShape('line', { x: 0.756, y: 5.322, w: 7.336, h: 0, line: { color: C.line35, width: 1, transparency: 64 } });
}

// 11 — right-hand copy column over a 2x2 product tag grid.
function slide11(s) {
  head(s, 'Life-Saving Medications', { x: 5.608, y: 1.229, w: 6.747, h: 0.656, fontSize: 40 });
  body(s, L.praesent, { x: 5.608, y: 1.974, w: 6.481, h: 0.863 });
  head(s, 'Healthier Lives, Brighter Futures', { x: 8.568, y: 4.091, w: 3.966, h: 0.757, fontSize: 24 });
  [
    { x: 8.662, y: 5.615, text: 'EnergyMax' },
    { x: 10.528, y: 5.615, text: 'PainRelief ' },
    { x: 8.662, y: 6.204, text: 'LiverCare' },
    { x: 10.528, y: 6.204, text: 'HerbaMax ' }
  ].forEach(p => pill(s, p.text, { x: p.x, y: p.y, w: 1.763, h: 0.479, fill: C.blue, align: 'center', fontSize: 14 }));
}

// 12 — industry stat page: 85% figure with supporting copy.
function slide12(s) {
  head(s, 'Pharmaceutical Industry', { x: 1.053, y: 1.338, w: 6.242, h: 0.545, fontSize: 32 });
  body(s, L.praesent, { x: 1.053, y: 1.997, w: 6.097, h: 0.863 });
  s.addShape('roundRect', {
    x: 8.235, y: 0.985, w: 2.356, h: 2.357, rectRadius: radius(7599, 2.356, 2.357),
    fill: { color: C.lime }, line: { type: 'none' }
  });
  head(s, '85%', { x: 4.087, y: 3.906, w: 3.936, h: 1.698, fontSize: 115 });
  body(s, 'Biotech Meets Pharma: Transforming Medicine Together', { x: 8.178, y: 4.136, w: 4.14, h: 0.851, fontSize: 18 });
  [
    { x: 4.253, text: 'EnergyMax' },
    { x: 6.162, text: 'PainRelief ' },
    { x: 8.059, text: 'LiverCare' }
  ].forEach(p => pill(s, p.text, { x: p.x, y: 6.07, w: 1.763, h: 0.408, fill: C.white, align: 'center', fontSize: 14 }));
}

// 13 — doctor profile with a 50% stat and product tags.
function slide13(s) {
  s.addShape('roundRect', {
    x: 5.395, y: 2.032, w: 3.519, h: 4.861, rectRadius: radius(6697, 3.519, 4.861),
    fill: { color: C.blue }, line: { type: 'none' }
  });
  head(s, 'Meet Our Dr. Axelius', { x: 0.701, y: 1.241, w: 4.392, h: 1.632, fontSize: 54 });
  body(s, L.praesentShort, { x: 0.701, y: 3.157, w: 3.973, h: 0.863 });
  pill(s, 'Let’s Talk Together', { x: 0.788, y: 4.344, w: 2.279, h: 0.408, fill: C.blue, align: 'center', fontSize: 14 });
  arrowBadge(s, { x: 3.151, y: 4.334, d: 0.415, fill: C.blue, line: C.ink25, dir: 'ne', arrow: C.ink15 });

  head(s, [{ text: '50' }, { text: '%', options: { superscript: true } }], { x: 10.013, y: 1.992, w: 1.973, h: 1.011, fontSize: 66 });
  head(s, 'Healthier Lives', { x: 10.013, y: 2.947, w: 1.98, h: 0.323, fontSize: 16, color: C.blueDk });
  body(s, L.short, { x: 10.013, y: 3.381, w: 2.545, h: 0.863 });
  [
    { x: 10.037, y: 6.13, text: 'EnergyMax' },
    { x: 11.23, y: 6.13, text: 'PainRelief ' },
    { x: 10.037, y: 6.539, text: 'LiverCare' },
    { x: 11.23, y: 6.539, text: 'HerbaMax ' }
  ].forEach(p => pill(s, p.text, { x: p.x, y: p.y, w: 1.13, h: 0.329, fill: C.white, align: 'center', fontSize: 11 }));
}

// 14 — team page: two doctor cards, a stat and a side note.
function slide14(s) {
  head(s, 'Our Best Doctor', { x: 0.726, y: 0.99, w: 5.001, h: 2.289, fontSize: 80 });
  head(s, '85%', { x: 0.792, y: 4.311, w: 1.837, h: 0.656, fontSize: 40 });
  body(s, L.praesentShort, { x: 0.792, y: 4.986, w: 3.973, h: 0.863 });
  pill(s, 'Let’s Talk Together', { x: 0.878, y: 6.174, w: 2.279, h: 0.408, fill: C.white, align: 'center', fontSize: 14 });

  s.addShape('roundRect', {
    x: 6.696, y: 0.932, w: 3.197, h: 3.273, rectRadius: radius(7625, 3.197, 3.273),
    fill: { color: C.lime }, line: { type: 'none' }
  });
  tx(s, 'Dr. Sarah Andreas', { x: 6.886, y: 1.222, w: 1.325, h: 0.586, fontFace: MAJOR, fontSize: 16, color: C.ink15, lineSpacingMultiple: 0.9 });
  tx(s, 'Specialist Vitamin', { x: 6.886, y: 1.758, w: 0.839, h: 0.404, fontSize: 10, color: C.limeInk, lineSpacingMultiple: 0.9 });

  s.addShape('roundRect', {
    x: 5.571, y: 4.443, w: 4.14, h: 2.067, rectRadius: radius(8834, 4.14, 2.067),
    fill: { color: C.blue }, line: { type: 'none' }
  });
  tx(s, 'Dr. Mateo Alexius', { x: 5.811, y: 4.655, w: 1.553, h: 0.586, fontFace: MAJOR, fontSize: 16, color: C.ink15, lineSpacingMultiple: 0.9 });
  tx(s, 'Specialist Capsule', { x: 5.799, y: 5.226, w: 1.421, h: 0.26, fontSize: 10, color: C.blueInk, lineSpacingMultiple: 0.9 });

  // stomach glyph
  s.addShape('teardrop', { x: 10.45, y: 1.14, w: 0.42, h: 0.42, rotate: 200, fill: { color: C.violet }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 10.585, y: 1.017, w: 0.085, h: 0.3, rectRadius: 0.042, fill: { color: C.violet }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 10.398, y: 1.312, w: 0.091, h: 0.262, rectRadius: 0.045, fill: { color: C.violet }, line: { type: 'none' } });

  body(s, L.short, { x: 10.287, y: 1.806, w: 2.177, h: 1.126 });
  pill(s, 'LiverCare', { x: 10.326, y: 3.383, w: 1.468, h: 0.355, fill: C.white, align: 'center' });
  arrowBadge(s, { x: 11.851, y: 3.383, d: 0.354, fill: C.white, line: C.ink25, dir: 'sw', arrow: C.ink15 });
  tx(s, 'Dr. Zara Adhisty', { x: 10.145, y: 4.33, w: 0.896, h: 0.464, fontFace: MAJOR, fontSize: 12, color: C.ink15, lineSpacingMultiple: 0.9 });
}

// 15 — donut breakdown of market data.
function slide15(s) {
  s.addShape('ellipse', { x: 4.363, y: 1.477, w: 4.546, h: 4.546, fill: { color: C.white }, line: { type: 'none' }, shadow: SOFT_SHADOW });
  s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [0.7, 0.2, 0.1] }], {
    x: 4.072, y: 1.621, w: 5.061, h: 4.281,
    chartColors: [C.blue, C.lime, C.violet],
    showLegend: false, showTitle: false, showValue: true,
    dataLabelFormatCode: '0%', dataLabelFontFace: MAJOR, dataLabelFontSize: 18, dataLabelColor: C.ink15,
    dataLabelPosition: 'bestFit', dataBorder: { pt: 7, color: C.white }
  });
  s.addShape('ellipse', {
    x: 6.035, y: 3.194, w: 1.135, h: 1.135, fill: { color: C.white },
    line: { color: C.line15, width: 0.5 },
    shadow: { type: 'outer', blur: 40, offset: 35, angle: 45, color: C.black, opacity: 0.1, rotateWithShape: false }
  });

  head(s, 'Recap Pharma Data', { x: 0.882, y: 0.99, w: 3.463, h: 2.55, fontSize: 60 });

  s.addShape('roundRect', {
    x: 1.315, y: 4.803, w: 3.745, h: 1.423, rectRadius: radius(13186, 3.745, 1.423),
    fill: { color: C.white }, line: { type: 'none' }, shadow: SOFT_SHADOW
  });
  tx(s, 'Market Capitalization', { x: 1.579, y: 5.071, w: 2.468, h: 0.303 });
  tx(s, '$ 393,654,541,00', { x: 1.579, y: 5.372, w: 3.372, h: 0.572, fontFace: MAJOR, fontSize: 28, color: C.ink15 });

  s.addShape('roundRect', {
    x: 9.433, y: 3.921, w: 2.772, h: 2.555, rectRadius: radius(11447, 2.772, 2.555),
    fill: { color: C.blue }, line: { type: 'none' }
  });
  s.addText([{ text: '70' }, { text: '%', options: { superscript: true } }], {
    x: 9.695, y: 4.306, w: 2.29, h: 0.707, valign: 'middle', fontFace: MAJOR, fontSize: 36, color: C.ink05, charSpacing: -3
  });
  tx(s, 'Energy Max', { x: 9.695, y: 5.017, w: 1.958, h: 0.352, fontFace: MAJOR, fontSize: 18, color: C.ink05, lineSpacingMultiple: 0.8 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', { x: 9.695, y: 5.489, w: 2.249, h: 0.565, color: C.gray35, lineSpacingMultiple: 1.2 });

  body(s, L.praesentShort, { x: 9.416, y: 1.421, w: 3.029, h: 1.126 });
  pill(s, 'LiverCare', { x: 9.493, y: 0.906, w: 1.227, h: 0.355, fill: C.white, align: 'center' });
}

// 16 — four staggered capsule bars with arrow leaders.
function slide16(s) {
  const bars = [
    { x: 8.435, y: -1.969, h: 6.865, fill: C.blue, label: 'Immune Boost', labelX: 9.075, labelW: 2.878, labelY: 1.275, labelColor: C.ink15, arrowX: 5.95, arrowY: 1.464 },
    { x: 8.856, y: -0.043, h: 6.024, fill: C.wash, label: 'Collagen Glow', labelX: 9.075, labelW: 2.878, labelY: 2.775, labelColor: C.black, arrowX: 6.862, arrowY: 2.969 },
    { x: 9.241, y: 1.848, h: 5.254, fill: C.violet, label: 'Probiotic Acid', labelX: 8.79, labelW: 3.162, labelY: 4.293, labelColor: C.black, arrowX: 7.606, arrowY: 4.475 },
    { x: 9.679, y: 3.791, h: 4.378, fill: C.lime, label: 'Ginseng Vital', labelX: 9.075, labelW: 2.878, labelY: 5.815, labelColor: C.black, arrowX: 8.473, arrowY: 5.98 }
  ];
  bars.forEach(b => {
    s.addShape('roundRect', {
      x: b.x, y: b.y, w: 1.108, h: b.h, rotate: 270,
      rectRadius: radius(50000, 1.108, b.h), fill: { color: b.fill }, line: { type: 'none' }
    });
    tx(s, b.label, {
      x: b.labelX, y: b.labelY, w: b.labelW, h: 0.404, align: 'right',
      fontFace: MAJOR, fontSize: 20, color: b.labelColor, lineSpacingMultiple: 0.9
    });
    s.addShape('line', { x: b.arrowX, y: b.arrowY, w: 0.614, h: 0, line: { color: b.labelColor, width: 0.75, endArrowType: 'triangle' } });
  });

  head(s, 'Healthier Lives, Brighter Futures', { x: 0.882, y: 1.606, w: 4.118, h: 3.357, fontSize: 60 });
  body(s, L.mid, { x: 0.882, y: 5.03, w: 3.723, h: 0.863 });
}

// 17 — node statistics dashboard: two KPI banners over an area/line chart.
function slide17(s) {
  const seriesA = [5, 5, 8, 5, 9, 14, 11, 16, 14, 11, 7, 14, 20, 22, 28, 33, 56, 144, 134, 104, 110, 60, 45, 50, 40, 22, 12, 8, 10];
  const seriesB = [10, 2, 2, 8, 13, 20, 19, 34, 12, 40, 22, 35, 35, 30, 33, 40, 80, 105, 80, 70, 40, 30, 35, 39, 37, 45, 30, 25, 20];
  const labels = seriesA.map((_, i) => String(i + 1));

  s.addChart(
    [
      { type: 'area', data: [{ name: 'Category A', labels, values: seriesA }, { name: 'Category B Fill', labels, values: seriesB }],
        options: { chartColors: [C.limeWash, C.violetWash] } },
      { type: 'line', data: [{ name: 'Category A Fill', labels, values: seriesA }, { name: 'Category B', labels, values: seriesB }],
        options: { chartColors: [C.sand, C.violet], lineSize: 1, lineDataSymbol: 'circle', lineDataSymbolSize: 5 } }
    ],
    {
      x: 1.422, y: 2.611, w: 10.092, h: 3.368,
      showLegend: false, showTitle: false,
      catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 160, valAxisMinVal: 0,
      catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
      plotArea: { fill: { type: 'none' } }, chartArea: { fill: { type: 'none' } }
    }
  );

  [
    { x: 1.5, w: 0.743, text: 'Apr 21' }, { x: 2.432, w: 0.818, text: 'May 19' }, { x: 3.363, w: 0.654, text: 'Jun 3' },
    { x: 4.295, w: 0.694, text: 'Jul 23' }, { x: 5.226, w: 0.778, text: 'Aug 17' }, { x: 6.157, w: 0.78, text: 'Sep 12' },
    { x: 7.089, w: 0.647, text: 'Oct 2' }, { x: 8.02, w: 0.703, text: 'Nov 8' }, { x: 8.951, w: 0.784, text: 'Dec 24' },
    { x: 9.883, w: 0.767, text: 'Jan 15' }, { x: 10.814, w: 0.774, text: 'Feb 20' }
  ].forEach(t => tx(s, t.text, { x: t.x, y: 6.231, w: t.w, h: 0.391, wrap: false, color: C.ink05 }));

  const banners = [
    { x: 1.636, fill: C.violet, labelColor: C.white, value: '1.42%', valueColor: C.white, triColor: C.white, copyColor: C.white, gx: 1.917, copyX: 3.643 },
    { x: 6.836, fill: C.lime, labelColor: C.ink25, value: '2.68%', valueColor: C.limeInk, triColor: C.limeInk, copyColor: C.gray35, gx: 7.117, copyX: 8.843 }
  ];
  banners.forEach(b => {
    s.addShape('roundRect', {
      x: b.x, y: 0.925, w: 4.886, h: 1.148, rectRadius: radius(13186, 4.886, 1.148),
      fill: { color: b.fill }, line: { type: 'none' }
    });
    tx(s, 'Last 24 Hours', { x: b.gx, y: 1.13, w: 1.442, h: 0.303, color: b.labelColor });
    tx(s, b.value, { x: b.gx + 0.293, y: 1.431, w: 1.22, h: 0.438, fontFace: MAJOR, fontSize: 20, color: b.valueColor });
    s.addShape('triangle', { x: b.gx + 0.105, y: 1.585, w: 0.137, h: 0.118, fill: { color: b.triColor }, line: { type: 'none' } });
    tx(s, L.dot, { x: b.copyX, y: 1.22, w: 2.732, h: 0.559, fontSize: 11, color: b.copyColor, lineSpacingMultiple: 1.3 });
  });

  s.addShape('line', { x: 1.602, y: 2.476, w: 10.076, h: 0, line: { color: C.line15, width: 1 } });
  pill(s, '  Node Statistic', { x: 1.602, y: 2.994, w: 2.566, h: 0.596, fill: C.white, fontFace: MAJOR, fontSize: 16, color: C.ink15 });
  arrow(s, { x: 3.602, y: 3.188, w: 0.209, dir: 'se', color: C.ink15, width: 2 });

  s.addShape('wedgeRoundRectCallout', {
    x: 9.449, y: 2.948, w: 2.258, h: 1.06, fill: { color: C.white }, line: { color: C.ink25, width: 0.5 }
  });
  tx(s, 'Highest bid', { x: 9.656, y: 3.145, w: 1.186, h: 0.303, color: C.gray35 });
  tx(s, '8.023 WETH', { x: 9.656, y: 3.417, w: 1.86, h: 0.438, fontFace: MAJOR, fontSize: 20, color: C.violetDk });
}

// 18 — four circular thumbnails in a row.
function slide18(s) {
  head(s, 'Quality Control in Pharma', { x: 0.871, y: 1.153, w: 8.663, h: 2.289, fontSize: 80 });
  [
    { x: 0.782, fill: C.blue, shadow: undefined, text: 'Brain Art One', textX: 0.977, color: C.ink15 },
    { x: 2.715, fill: C.white, shadow: SOFT_SHADOW, text: 'Brain Art Two', textX: 2.909, color: C.gray50 },
    { x: 4.647, fill: C.lime, shadow: undefined, text: 'Brain Art Three', textX: 4.842, color: C.limeInk },
    { x: 6.58, fill: C.white, shadow: SOFT_SHADOW, text: 'Brain Art Four', textX: 6.774, color: C.gray50 }
  ].forEach(c => {
    s.addShape('ellipse', { x: c.x, y: 4.038, w: 1.682, h: 1.682, fill: { color: c.fill }, line: { type: 'none' }, shadow: c.shadow });
    tx(s, c.text, { x: c.textX, y: 4.533, w: 1.293, h: 0.747, align: 'center', fontSize: 16, color: c.color, lineSpacingMultiple: 1.2 });
  });
  arrowBadge(s, { x: 8.526, y: 4.563, d: 0.556, fill: C.ink15, dir: 'e', arrow: C.white });
}

// 19 — product shot (replaced by an image placeholder) with tags and copy.
function slide19(s) {
  imagePlaceholder(s, {
    x: 7.713, y: 1.114, w: 5.022, h: 5.285, rotate: 345, rectRadius: 0.3,
    fill: 'EDEDED', line: 'E4E4E4', label: '[image]', labelColor: C.line35,
    shadow: { type: 'outer', blur: 50, offset: 28, angle: 55, color: C.black, opacity: 0.08, rotateWithShape: false }
  });
  head(s, 'Pharma in Smartwatch ', { x: 0.871, y: 1.575, w: 6.413, h: 2.07, fontSize: 72 });
  [
    { x: 0.969, fill: C.white, text: 'EnergyMax' },
    { x: 2.878, fill: C.lime, text: 'PainRelief ' },
    { x: 4.776, fill: C.white, text: 'LiverCare' }
  ].forEach(p => pill(s, p.text, { x: p.x, y: 4.288, w: 1.763, h: 0.408, fill: p.fill, align: 'center', fontSize: 14 }));
  body(s, L.praesent, { x: 0.871, y: 4.992, w: 6.315, h: 0.863 });

  s.addShape('ellipse', { x: 7.912, y: 1.22, w: 1.1, h: 1.1, fill: { color: C.white }, line: { type: 'none' }, shadow: WIDE_SHADOW });
  s.addShape('ellipse', { x: 8.014, y: 1.322, w: 0.896, h: 0.896, fill: { type: 'none' }, line: { color: C.violet, width: 1.25, dashType: 'dash' } });
  stethoscope(s, 8.28, 1.55, 0.37, C.violet);

  // "drop an image here" prompt inherited from the slide layout
  s.addShape('roundRect', {
    x: 8.23, y: 2.141, w: 1.64, h: 2.21, rectRadius: 0.11,
    fill: { color: '909090' }, line: { type: 'none' }
  });
  s.addText('Drag and Drop Image Here', {
    x: 8.23, y: 3.021, w: 1.64, h: 0.45, align: 'center', valign: 'middle', wrap: false,
    fontFace: MINOR, fontSize: 18, color: C.white
  });
}

// 20 — closing "THANK YOU" with contact details.
function slide20(s) {
  s.addText('BUSINESS PRESENTATION TEMPLATE', {
    shape: 'roundRect', x: 0.756, y: 0.416, w: 3.353, h: 0.351, rectRadius: radius(50000, 3.353, 0.351),
    fill: { color: C.blue }, line: { color: C.ink15, width: 0.5 }, shadow: { type: 'outer', blur: 95, offset: 69, angle: 135, color: C.black, opacity: 0.1, rotateWithShape: false },
    valign: 'middle', wrap: false, margin: [14.4, 7.2, 3.6, 7.2], fontFace: MINOR, fontSize: 11, color: C.ink15
  });
  s.addText('BY JHON DHOE', {
    shape: 'roundRect', x: 11.144, y: 0.416, w: 1.647, h: 0.351, rectRadius: radius(50000, 1.647, 0.351),
    fill: { color: C.lime }, line: { color: C.ink15, width: 0.5 }, shadow: { type: 'outer', blur: 95, offset: 69, angle: 135, color: C.black, opacity: 0.1, rotateWithShape: false },
    valign: 'middle', wrap: false, align: 'right', margin: [14.4, 14.4, 3.6, 7.2], fontFace: MINOR, fontSize: 11, color: C.ink15
  });

  head(s, 'THANK YOU', { x: 0.542, y: 2.171, w: 11.742, h: 1.986, fontSize: 140, color: C.ink05, charSpacing: -1.5 });
  arrowBadge(s, { x: 0.916, y: 4.154, d: 0.397, fill: C.violet, line: C.ink15, lineWidth: 1, dir: 'se', arrow: C.ink15 });
  body(s, 'See You Next Time Everyone', { x: 1.532, y: 4.116, w: 8.812, h: 0.472, fontSize: 18 });

  [
    { x: 0.848, label: 'Address:', value: 'London city, no 1234' },
    { x: 4.257, label: 'Website:', value: 'WWW.Websiter.Com' },
    { x: 7.665, label: 'Number:', value: '+1234 5678 91012' }
  ].forEach(g => {
    tx(s, g.label, { x: g.x, y: 5.889, w: 2.513, h: 0.404, fontFace: MAJOR, fontSize: 18, color: C.blueDk });
    tx(s, g.value, { x: g.x, y: 6.289, w: 2.303, h: 0.378, fontSize: 14, lineSpacingMultiple: 1.3 });
  });
}

/* -------------------------------------------------------------------- main */

const SLIDES = [
  { build: slide01, chrome: false },
  { build: slide02 }, { build: slide03 }, { build: slide04 }, { build: slide05 },
  { build: slide06 }, { build: slide07, pageColor: C.limeInk }, { build: slide08 },
  { build: slide09 }, { build: slide10 }, { build: slide11 }, { build: slide12 },
  { build: slide13 }, { build: slide14 }, { build: slide15 }, { build: slide16 },
  { build: slide17 }, { build: slide18 }, { build: slide19, pageColor: C.limeInk },
  { build: slide20, chrome: false }
];

function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.title = 'Pharma Medical';

  SLIDES.forEach((def, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    if (def.chrome !== false) chrome(slide, i + 1, def.pageColor);
    def.build(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0c0b63ae-7107-4248-a766-c03a81375b42_grok_final.pptx') });
}

main().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
