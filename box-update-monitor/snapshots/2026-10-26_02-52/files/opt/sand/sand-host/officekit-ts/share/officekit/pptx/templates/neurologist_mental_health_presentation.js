/**
 * "Neurologist — Mental Health Presentation Template" (20 slides, 13.333 x 7.5 in)
 * rebuilt with pptxgenjs only.
 *
 * Raster artwork from the source deck (icon glyphs, the body silhouette) is not
 * embedded; each one becomes a flat colour block of the same size and position.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
// Theme colours of the original deck ("Custom 1" scheme).
const BG = 'F0F5F3'; // slide background (master)
const LIGHT = 'F9F9F9'; // lt1 / bg1
const DARK = '000000'; // tx1
const MINT = '86BCAA'; // accent1
const SAGE = 'AECAC2'; // accent2
const ORANGE = 'E85B2D'; // accent3
const PALE = 'F0F0F0'; // accent5
const GREY = '707070'; // accent6
const DEEP_MINT = '569B84'; // accent1 @ 75% luminance — slide 5 background
const NEAR_WHITE = 'FBFBFB'; // lt2 lightened, used for body copy on mint cards

const HEAD = 'Montserrat Medium'; // major font
const BODY = 'Open Sans'; // minor font

// ------------------------------------------------------------- lorem text ---
const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.';
const LOREM_SHORT =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies.';
const LOREM_NUNC = 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const LOREM_PELL =
  'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ' +
  'Proin pharetra nonummy pede. Mauris et orci.';
const QUOTE =
  '\u201CLorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas ultricies, purus lectus malesuada libero.\u201D';
const FORMULA_LINES = [
  'Lorem ipsum dolor sit ',
  'amet, consectetur adiiscing elit, sed do eiusmod temporincidi',
  'Dunt adiiscing elit.',
];
const CHECK_ITEM = 'Lorem ipsum dolor sit amet, consectetur';

// ------------------------------------------------------------------ utils ---

/** Turn a list of paragraphs into a pptxgenjs run array (blank strings = empty lines). */
function paras(list) {
  return list.map((t, i) => ({ text: t, options: { breakLine: i < list.length - 1 } }));
}

/** Plain top-aligned text box. */
function text(slide, str, opts) {
  slide.addText(str, Object.assign({ valign: 'top', fontFace: BODY, color: DARK }, opts));
}

/**
 * Hanging-indent list where the marker glyph is a coloured run of its own,
 * because pptxgenjs cannot tint `buChar` independently of the run text.
 * A non-breaking-space bullet supplies the hang; `align` must be set on the
 * box so pptxgenjs keeps marker and text inside one paragraph.
 */
function markerList(slide, x, y, w, h, items, marker, markerColor, textColor, opts) {
  const bullet = { characterCode: '00A0', indent: 14.4 };
  const runs = [];
  items.forEach((t, i) => {
    runs.push({ text: marker, options: { color: markerColor, bullet } });
    runs.push({ text: t, options: { color: textColor, bullet, breakLine: true } });
    if (i < items.length - 1 && opts && opts.blankBetween) {
      runs.push({ text: ' ', options: { bullet, breakLine: true } });
    }
  });
  slide.addText(runs, Object.assign(
    { x, y, w, h, valign: 'top', align: 'left', fontFace: BODY, fontSize: 12 }, opts));
}

/** Multi-paragraph body copy. */
function block(slide, list, opts) {
  text(slide, paras(list), opts);
}

/** Rounded rectangle; `r` is the corner radius in inches. */
function card(slide, x, y, w, h, r, fill, extra) {
  slide.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: r, fill: { color: fill } }, extra));
}

/** Pill-shaped button with a centred caption. */
function pill(slide, x, y, w, h, fill, caption, captionColor, fontSize) {
  slide.addText(caption, {
    shape: 'roundRect', x, y, w, h,
    rectRadius: h / 2,
    fill: { color: fill },
    fontFace: BODY,
    fontSize: fontSize || 12,
    color: captionColor,
    align: 'center',
    valign: 'middle',
  });
}

/** Small orange bullet dot used next to inline headings. */
function dot(slide, x, y) {
  slide.addShape('ellipse', { x, y, w: 0.157, h: 0.157, fill: { color: ORANGE } });
}

/** The recurring "long rule + short orange tick" divider. */
function divider(slide, x, y, longColor) {
  slide.addShape('line', { x, y, w: 2.411, h: 0, line: { color: longColor || MINT, width: 1.25 } });
  slide.addShape('line', { x: x + 2.482, y, w: 0.27, h: 0, line: { color: ORANGE, width: 1.25 } });
}

/** Five concentric hairline circles, fading outwards — the deck's watermark motif. */
const RING_SIZES = [3.933, 3.106, 2.405, 1.725, 0.985];
const RING_ALPHA = [10, 20, 40, 60, 80];
function rings(slide, cx, cy, color) {
  RING_SIZES.forEach((d, i) => {
    slide.addShape('ellipse', {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { type: 'none' },
      line: { color, width: 1, transparency: 100 - RING_ALPHA[i] },
    });
  });
}

// Stand-in colours for the deck's raster artwork. Each value is the average
// colour the original bitmap contributes over its own background, so a flat
// block of it reads at roughly the same weight as the artwork it replaces.
const IMG_BRAIN_ON_ORANGE = 'D96A46';
const IMG_NERVE_ON_WHITE = 'DEE5E3';
const IMG_NERVE_ON_MINT = '8DB9AA';
const IMG_BRAIN_ON_GREY = 'BCC0BF';
const IMG_BODY_SILHOUETTE = 'BFD9CF';
const IMG_CONTACT_GLYPH = '8DB8A9';

/** Raster artwork stand-in: a flat block in place of the bitmap. */
function imagePlaceholder(slide, x, y, w, h, color, radius) {
  slide.addShape(radius ? 'roundRect' : 'rect', {
    x, y, w, h,
    rectRadius: radius || undefined,
    fill: { color },
    objectName: '[image]',
  });
}

/** Linear gradient faked with slim bands (pptxgenjs has no gradient fill). */
function gradientBands(slide, x, y, w, h, angleDeg, stops, bands) {
  const a = (angleDeg * Math.PI) / 180;
  const ux = Math.cos(a), uy = Math.sin(a);
  // Length of the box measured along the gradient axis, and the axis origin.
  const len = Math.abs(w * ux) + Math.abs(h * uy);
  const ox = x + (ux < 0 ? w : 0), oy = y + (uy < 0 ? h : 0);
  const mix = (t) => {
    let lo = stops[0], hi = stops[stops.length - 1];
    for (let i = 0; i < stops.length - 1; i++) {
      if (t >= stops[i][0] && t <= stops[i + 1][0]) { lo = stops[i]; hi = stops[i + 1]; break; }
    }
    const f = hi[0] === lo[0] ? 0 : (t - lo[0]) / (hi[0] - lo[0]);
    const ch = (k) => {
      const c0 = parseInt(lo[1].substr(k * 2, 2), 16), c1 = parseInt(hi[1].substr(k * 2, 2), 16);
      return Math.round(c0 + (c1 - c0) * f).toString(16).padStart(2, '0');
    };
    return (ch(0) + ch(1) + ch(2)).toUpperCase();
  };
  // Each band is a thin strip perpendicular to the axis, over-sized so the
  // rotated strips still cover the corners; the whole set is clipped by the
  // slide edge, which is what the original gradient looks like.
  const step = len / bands;
  const long = Math.hypot(w, h) * 1.2;
  for (let i = 0; i < bands; i++) {
    const t = (i + 0.5) / bands;
    const cx = ox + ux * step * (i + 0.5), cy = oy + uy * step * (i + 0.5);
    slide.addShape('rect', {
      x: cx - step, y: cy - long / 2, w: step * 2, h: long,
      fill: { color: mix(t) }, rotate: angleDeg,
    });
  }
}

/** Master furniture: corner label, footer, page number and watermark rings. */
function chrome(slide, pageNo) {
  rings(slide, 12.9995, 7.2275, MINT);
  text(slide, 'Neurologist', { x: 0.167, y: 0.121, w: 1.5, h: 0.286, fontFace: HEAD, fontSize: 11, color: MINT });
  text(slide, 'www.yourwebsite.com', { x: 0.167, y: 7.073, w: 2.0, h: 0.269, fontSize: 10, color: MINT });
  text(slide, String(pageNo), {
    x: 11.39, y: 7.096, w: 1.783, h: 0.303, fontFace: HEAD, fontSize: 12, color: DARK,
    transparency: 70, align: 'right',
  });
}

// ------------------------------------------------------------ slide 1..20 ---

function slide1(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  // Mint frame with a background-coloured window punched out of the top half.
  card(s, 0.156, 0.194, 13.022, 7.111, 0.363, MINT);
  card(s, 0.428, 0.466, 12.45, 3.166, 0.194, BG);
  card(s, 0.789, 1.499, 2.915, 2.83, 0.2155, ORANGE);
  card(s, 3.84, 1.499, 2.915, 2.83, 0.2155, ORANGE);
  text(s, 'Neurologist', { x: 0.685, y: 4.638, w: 8.015, h: 1.481, fontFace: HEAD, fontSize: 82, color: LIGHT });
  text(s, 'Mental Health Presentation Template', { x: 0.719, y: 6.237, w: 7.037, h: 0.404, fontSize: 18, color: LIGHT });
  card(s, 9.222, 4.6, 3.367, 2.156, 0.2704, SAGE);
  block(s, ['Lorem ipsum dolor sit amet, consectetur adiiscing elit. ', '',
    'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '],
    { x: 9.611, y: 4.88, w: 2.722, h: 1.538, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
}

function slide2(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  gradientBands(s, 0, 0, 13.333, 7.5, 45, [[0, 'BDD3CD'], [0.659, '88BEAB'], [1, '88BEAB']], 48);
  rings(s, 0.4335, 0.0, MINT);
  rings(s, 13.3335, 7.4995, LIGHT);
  card(s, 1.043, 3.152, 11.247, 3.35, 0.3822, SAGE);
  divider(s, 1.222, 2.691, LIGHT);
  text(s, paras(['About', 'Neurologist']), {
    x: 1.093, y: 0.872, w: 5.908, h: 1.723, fontFace: HEAD, fontSize: 60, color: LIGHT, lineSpacingMultiple: 0.8,
  });
  pill(s, 8.97, 2.415, 1.596, 0.418, LIGHT, 'General', DARK);
  pill(s, 10.625, 2.415, 1.596, 0.418, ORANGE, 'Medical', LIGHT);
  const copy = [LOREM_LONG, '', LOREM_NUNC, LOREM_PELL];
  block(s, copy, { x: 1.433, y: 3.667, w: 4.94, h: 2.507, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
  block(s, copy, { x: 6.767, y: 3.667, w: 5.217, h: 2.265, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
}

function slide3(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 3);
  card(s, 6.089, 4.189, 2.744, 2.678, 0.2039, ORANGE);
  card(s, 8.933, 4.189, 3.722, 2.678, 0.2039, SAGE);
  text(s, 'Introduction', { x: 0.63, y: 4.129, w: 4.601, h: 0.37, fontSize: 16 });
  text(s, 'A Neurologist\u2019s Perspective',
    { x: 0.63, y: 4.483, w: 4.924, h: 1.582, fontFace: HEAD, fontSize: 44 });
  divider(s, 0.711, 6.313);
  text(s, '+12.31K', { x: 6.31, y: 4.422, w: 2.268, h: 0.707, fontFace: HEAD, fontSize: 36, color: LIGHT });
  text(s, 'Emotional Balance, And Brain Repair.',
    { x: 6.31, y: 5.685, w: 2.168, h: 0.909, fontSize: 16, color: LIGHT });
  block(s, [LOREM_LONG], { x: 9.333, y: 4.667, w: 3.033, h: 1.78, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
}

function slide4(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 4);
  text(s, 'Advances in Neurology', { x: 0.63, y: 0.938, w: 4.081, h: 1.582, fontSize: 44 });
  divider(s, 0.711, 2.768);
  dot(s, 0.697, 3.476);
  text(s, 'Brain Health Matters', { x: 0.99, y: 3.252, w: 2.048, h: 0.64, fontSize: 16 });
  dot(s, 3.837, 3.476);
  text(s, 'Unlocking the Mind', { x: 4.13, y: 3.252, w: 2.048, h: 0.64, fontSize: 16 });
  block(s, [LOREM_SHORT], { x: 0.63, y: 4.169, w: 5.937, h: 0.569, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  block(s, [LOREM_LONG, '', LOREM_PELL],
    { x: 0.63, y: 5.104, w: 5.948, h: 1.538, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  pill(s, 7.792, 6.248, 1.819, 0.418, ORANGE, 'Neurologist', LIGHT);
}

function slide5(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  s.background = { color: DEEP_MINT };
  chrome(s, 5);
  card(s, 2.958, 2.744, 4.0, 1.016, 0.1635, ORANGE);
  imagePlaceholder(s, 3.203, 2.989, 0.517, 0.517, IMG_BRAIN_ON_ORANGE);
  text(s, 'Emotional Balance, And Brain Repair.',
    { x: 3.918, y: 2.926, w: 3.019, h: 0.64, fontSize: 16, color: LIGHT });
  card(s, 0.768, 5.385, 4.0, 1.016, 0.1762, LIGHT);
  imagePlaceholder(s, 1.034, 5.628, 0.533, 0.533, IMG_NERVE_ON_WHITE);
  text(s, 'Neurological Care for a Better Life', { x: 1.664, y: 5.559, w: 2.842, h: 0.64, fontSize: 16 });
  text(s, 'Caring for the Brain and Mind',
    { x: 7.43, y: 1.216, w: 5.359, h: 1.447, fontFace: HEAD, fontSize: 40, color: LIGHT });
  divider(s, 7.511, 3.046, LIGHT);
  block(s, [LOREM_LONG, '', LOREM_NUNC, LOREM_PELL],
    { x: 7.389, y: 3.744, w: 4.778, h: 2.507, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
}

function slide6(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 6);
  text(s, 'Modern Approaches to Brain, Nerve, and Mental Health',
    { x: 0.625, y: 0.734, w: 7.957, h: 2.322, fontSize: 44 });
  divider(s, 3.23, 2.667);
  card(s, 4.848, 3.523, 7.819, 2.833, 0.3436, SAGE);
  // 60 / 40 progress ring
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [60, 40] }], {
    x: 4.974, y: 3.831, w: 2.425, h: 2.223,
    chartColors: [ORANGE, 'C4D8D2'], holeSize: 84,
    showLegend: false, showValue: false, showTitle: false,
    chartArea: { fill: { color: SAGE } }, plotArea: { fill: { color: SAGE } },
  });
  text(s, '60%', { x: 5.613, y: 4.617, w: 1.157, h: 0.64, fontFace: HEAD, fontSize: 32, color: LIGHT, align: 'center' });
  text(s, 'Protecting Your Most Vital Organ',
    { x: 7.743, y: 3.94, w: 4.476, h: 0.37, fontFace: HEAD, fontSize: 16, color: LIGHT });
  block(s, ['Lorem ipsum dolor sit amet, consectetur adiiscing elit, sed do eiusmod tempor incididunt'],
    { x: 7.743, y: 4.403, w: 4.435, h: 0.565, fontSize: 12, color: LIGHT, lineSpacingMultiple: 1.2 });
  const stats = [
    { x: 7.777, big: [{ text: '+293' }], label: 'Neurology', lw: 1.504 },
    { x: 9.281, big: [{ text: '100' }, { text: '%', options: { fontSize: 24 } }], label: 'Mental', lw: 1.402 },
    { x: 11.009, big: [{ text: '80&' }], label: 'Lifestyle', lw: 1.169 },
  ];
  stats.forEach((st) => {
    text(s, st.big, { x: st.x, y: 5.095, w: st.lw, h: 0.64, fontFace: HEAD, fontSize: 32, color: LIGHT });
    text(s, st.label, { x: st.x, y: 5.687, w: 1.402, h: 0.303, fontFace: HEAD, fontSize: 12, color: LIGHT });
  });
}

function slide7(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 7);
  text(s, 'Neurological Health Strategies',
    { x: 0.597, y: 1.349, w: 5.592, h: 1.582, fontFace: HEAD, fontSize: 44 });
  divider(s, 0.711, 3.235);
  text(s, 'Quantitative', { x: 0.63, y: 4.041, w: 2.136, h: 0.37, fontFace: HEAD, fontSize: 16 });
  block(s, [LOREM_SHORT], { x: 0.63, y: 4.391, w: 4.47, h: 0.811, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  pill(s, 0.714, 5.682, 1.596, 0.418, LIGHT, 'Mental', DARK);
  pill(s, 2.37, 5.682, 1.596, 0.418, ORANGE, 'Health', LIGHT);

  const cards = [
    { x: 6.683, y: 1.044, fill: LIGHT, title: 'Formula A', ink: DARK, body: GREY, tx: 7.011 },
    { x: 9.697, y: 1.044, fill: MINT, title: 'Formula B', ink: LIGHT, body: LIGHT, tx: 10.014 },
    { x: 6.683, y: 3.852, fill: SAGE, title: 'Formula C', ink: LIGHT, body: LIGHT, tx: 7.011 },
    { x: 9.697, y: 3.852, fill: ORANGE, title: 'Formula D', ink: LIGHT, body: LIGHT, tx: 10.014 },
  ];
  cards.forEach((c) => {
    card(s, c.x, c.y, 2.631, 2.389, 0.1977, c.fill);
    text(s, c.title, { x: c.x + 0.328, y: c.y + 0.291, w: 2.077, h: 0.337, fontFace: HEAD, fontSize: 14, color: c.ink });
    block(s, FORMULA_LINES,
      { x: c.tx, y: c.y + 0.784, w: 2.106, h: 1.215, fontSize: 11, color: c.body, lineSpacingMultiple: 1.2 });
  });
}

function slide8(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 8);
  card(s, 0.667, 3.061, 5.746, 2.917, 0.276, LIGHT);
  card(s, 6.921, 3.061, 5.746, 2.917, 0.2344, MINT);
  text(s, 'From Stress to Strength Neurology in Daily Life',
    { x: 2.688, y: 0.768, w: 7.957, h: 1.313, fontFace: HEAD, fontSize: 36, align: 'center' });
  divider(s, 5.291, 2.489);
  const bullets = [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venen atis tristique.',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus.',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
  ];
  [{ x: 1.068, ink: DARK, body: GREY, bul: MINT }, { x: 7.423, ink: LIGHT, body: LIGHT, bul: ORANGE }].forEach((col) => {
    text(s, 'Study Case One', { x: col.x, y: 3.335, w: 4.823, h: 0.37, fontFace: HEAD, fontSize: 16, color: col.ink });
    markerList(s, col.x, 3.879, 4.71, 1.78, bullets, '\u2022  ', col.bul, col.body,
      { lineSpacingMultiple: 1.2, blankBetween: true });
  });
}

function slide9(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 9);
  text(s, 'Choosing The Right Method For Our Research',
    { x: 2.941, y: 0.376, w: 7.451, h: 1.178, fontFace: HEAD, fontSize: 32, align: 'center' });
  divider(s, 5.291, 1.823);
  block(s, [LOREM_SHORT],
    { x: 2.259, y: 1.991, w: 8.815, h: 0.569, fontSize: 12, color: GREY, align: 'center', lineSpacingMultiple: 1.2 });
  const methods = [
    { x: 1.623, fill: LIGHT, ink: GREY, tag: 'Metode 01', tagFill: MINT, tagInk: LIGHT, tagX: 3.396 },
    { x: 5.173, fill: SAGE, ink: LIGHT, tag: 'Metode 02', tagFill: ORANGE, tagInk: LIGHT, tagX: 6.871 },
    { x: 8.681, fill: ORANGE, ink: LIGHT, tag: 'Metode 02', tagFill: LIGHT, tagInk: DARK, tagX: 10.416 },
  ];
  methods.forEach((m) => {
    card(s, m.x, 4.44, 3.025, 2.293, 0.2489, m.fill);
    block(s, [LOREM_NUNC],
      { x: m.x + 0.276, y: 5.29, w: 2.388, h: 0.811, fontSize: 12, color: m.ink, lineSpacingMultiple: 1.2 });
    pill(s, m.tagX, 6.307, 1.104, 0.273, m.tagFill, m.tag, m.tagInk, 9);
  });
}

function slide10(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 10);
  rings(s, 2.8165, 2.6225, MINT);
  text(s, 'Modern Approaches to Brain', { x: 6.874, y: 0.769, w: 5.784, h: 1.447, fontSize: 40 });
  divider(s, 6.955, 2.637);
  text(s, '25K+', { x: 6.813, y: 3.105, w: 1.301, h: 0.64, fontFace: HEAD, fontSize: 32, color: MINT });
  text(s, 'Brain Health Matters', { x: 8.123, y: 3.173, w: 1.632, h: 0.505, fontSize: 12 });
  text(s, '150+', { x: 9.724, y: 3.105, w: 1.362, h: 0.64, fontFace: HEAD, fontSize: 32, color: ORANGE, align: 'center' });
  text(s, 'Unlocking the Mind', { x: 11.108, y: 3.173, w: 1.632, h: 0.505, fontSize: 12 });
  block(s, [LOREM_LONG, '', LOREM_PELL],
    { x: 6.797, y: 4.095, w: 5.431, h: 2.023, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  card(s, 0.723, 4.758, 2.001, 1.95, 0.286, ORANGE);
  imagePlaceholder(s, 0.999, 5.07, 0.415, 0.415, IMG_BRAIN_ON_ORANGE);
  text(s, 'Emotional Balance, And Brain Repair.',
    { x: 0.877, y: 5.734, w: 1.674, h: 0.707, fontSize: 12, color: LIGHT });
}

function slide11(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 11);
  rings(s, 8.5835, 1.2995, MINT);
  text(s, 'Neurological Health Strategies',
    { x: 0.597, y: 0.838, w: 5.592, h: 1.582, fontFace: HEAD, fontSize: 44 });
  divider(s, 0.711, 2.724);
  text(s, 'Frontiers in Brain Research', { x: 0.63, y: 3.196, w: 4.015, h: 0.337, fontFace: HEAD, fontSize: 14 });
  block(s, [LOREM_SHORT], { x: 0.63, y: 3.591, w: 4.47, h: 0.811, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  card(s, 0.644, 4.839, 8.439, 1.509, 0.1721, MINT);
  [1.05, 5.05].forEach((x) => {
    text(s, 'Subtitle Here',
      { x, y: 5.004, w: 3.5, h: 0.428, fontFace: HEAD, fontSize: 16, color: 'F2F2F2', lineSpacingMultiple: 1.3 });
    block(s, ['Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce consequat'],
      { x, y: 5.497, w: 3.5, h: 0.566, fontSize: 12, color: NEAR_WHITE, lineSpacingMultiple: 1.2 });
  });
}

function slide12(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 12);
  text(s, 'Bridging Neuroscience and Patient Care',
    { x: 0.597, y: 0.772, w: 5.592, h: 2.322, fontFace: HEAD, fontSize: 44 });
  divider(s, 0.711, 3.368);
  dot(s, 0.752, 4.32);
  text(s, 'Brain Health Matters', { x: 1.046, y: 4.096, w: 2.048, h: 0.64, fontSize: 16 });
  dot(s, 3.892, 4.32);
  text(s, 'Unlocking the Mind', { x: 4.186, y: 4.096, w: 2.048, h: 0.64, fontSize: 16 });
  card(s, 7.212, 3.929, 4.0, 1.016, 0.1762, LIGHT);
  imagePlaceholder(s, 7.478, 4.172, 0.533, 0.533, IMG_NERVE_ON_WHITE);
  text(s, 'Neurological Care for a Better Life', { x: 8.109, y: 4.103, w: 2.842, h: 0.64, fontSize: 16 });
  card(s, 0.726, 5.233, 5.796, 1.222, 0.1482, MINT);
  block(s, [QUOTE],
    { x: 1.137, y: 5.507, w: 5.097, h: 0.646, fontSize: 14, italic: true, color: LIGHT, lineSpacingMultiple: 1.2 });
  text(s, '+12.31K', { x: 7.716, y: 5.729, w: 1.772, h: 0.64, fontFace: HEAD, fontSize: 32 });
  text(s, 'About Neurologist', { x: 9.654, y: 5.49, w: 2.726, h: 0.353, fontFace: HEAD, fontSize: 15 });
  block(s, ['PLACEHOLDER'],
    { x: 9.654, y: 5.844, w: 2.817, h: 0.811, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
}

/** Wingdings tick bullets of the original become coloured "\u2713" markers. */
function checkList(slide, x, y, w, h, items) {
  markerList(slide, x, y, w, h, items, '\u2713 ', MINT, GREY, { lineSpacingMultiple: 1.5 });
}

function slide13(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 13);
  text(s, 'Our Specialist Doctor', { x: 0.919, y: 0.769, w: 5.784, h: 0.774, fontSize: 40 });
  divider(s, 1.0, 1.726);
  const doctors = [
    { y: 2.167, name: 'Dr. Devin Oller', badgeX: 7.763, badgeY: 2.828 },
    { y: 4.5, name: 'Dr. Beckett Binn', badgeX: 9.671, badgeY: 5.817 },
  ];
  doctors.forEach((d) => {
    card(s, 0.927, d.y, 5.517, 2.122, 0.1756, LIGHT);
    text(s, d.name, { x: 1.396, y: d.y + 0.29, w: 4.928, h: 0.401, fontSize: 16, lineSpacingMultiple: 1.2 });
    checkList(s, 1.396, d.y + 0.767, 4.553, 0.978, [CHECK_ITEM, CHECK_ITEM, CHECK_ITEM]);
    card(s, d.badgeX, d.badgeY, 2.487, 0.695, 0.3475, LIGHT);
    text(s, d.name, {
      x: d.badgeX + 0.278, y: d.badgeY + 0.083, w: 1.897, h: 0.337,
      fontFace: HEAD, fontSize: 14, color: '404040', align: 'center',
    });
    text(s, 'Job Position', {
      x: d.badgeX + 0.288, y: d.badgeY + 0.343, w: 1.897, h: 0.269,
      fontSize: 10, color: ORANGE, align: 'center',
    });
  });
}

function slide14(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 14);
  text(s, 'Modern Approaches to Brain', { x: 0.852, y: 0.769, w: 5.784, h: 1.447, fontSize: 40 });
  divider(s, 0.933, 2.637);
  card(s, 4.412, 3.073, 1.982, 1.949, 0.2288, ORANGE);
  imagePlaceholder(s, 4.685, 3.4, 0.412, 0.412, IMG_BRAIN_ON_ORANGE);
  text(s, 'Emotional Balance, And Brain Repair.',
    { x: 4.564, y: 4.058, w: 1.658, h: 0.7, fontSize: 12, color: LIGHT });
  [{ dx: 2.108, tx: 1.18, label: 'Our Services 01' }, { dx: 4.697, tx: 3.737, label: 'Our Services 02' }]
    .forEach((c) => {
      dot(s, c.dx, 5.632);
      text(s, c.label, { x: c.tx, y: 5.974, w: 2.007, h: 0.353, fontFace: HEAD, fontSize: 15, align: 'center' });
    });
  text(s, 'Efficient Taxes. Effective Growth.',
    { x: 6.937, y: 3.727, w: 4.928, h: 0.407, fontSize: 16, lineSpacingMultiple: 1.2 });
  checkList(s, 6.937, 4.17, 4.553, 0.978, [CHECK_ITEM, CHECK_ITEM, CHECK_ITEM]);
  block(s, [LOREM_SHORT], { x: 6.852, y: 5.502, w: 4.803, h: 0.811, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
}

/** Three-arc progress ring (orange / white / sage) used on slide 15. */
function progressRing(slide, x, y, size, percent, label) {
  const arcs = [
    { color: ORANGE, range: [270, 105.8], flipH: false },
    { color: LIGHT, range: [269.8, 1.5], flipH: true },
    { color: SAGE, range: [358.8, 76.2], flipH: true },
  ];
  arcs.forEach((a) => {
    slide.addShape('arc', {
      x, y, w: size, h: size, flipH: a.flipH,
      angleRange: a.range, line: { color: a.color, width: 11.25 },
    });
  });
  slide.addText(percent, {
    x, y: y + 0.466, w: size, h: 0.6,
    fontFace: HEAD, fontSize: 28, color: 'F2F2F2', align: 'center', valign: 'top',
  });
  slide.addText(label, {
    x: x - 0.329, y: y + 1.765, w: 2.177, h: 0.337,
    fontFace: HEAD, fontSize: 14, color: LIGHT, align: 'center', valign: 'top',
  });
}

function slide15(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 15);
  card(s, 0.667, 3.421, 12.0, 2.995, 0.3036, MINT);
  [{ x: 1.098, pct: '73%', label: 'Data One' },
   { x: 3.976, pct: '63%', label: 'Data Two' },
   { x: 6.843, pct: '40%', label: 'Data Three' }]
    .forEach((g) => progressRing(s, g.x + 0.329, 3.901, 1.519, g.pct, g.label));
  imagePlaceholder(s, 9.844, 3.817, 0.435, 0.435, IMG_NERVE_ON_MINT);
  text(s, 'Healthy Brain, Healthy Mind', { x: 9.747, y: 4.518, w: 2.015, h: 0.64, fontSize: 16, color: LIGHT });
  block(s, ['Lorem ipsum dolor sit ', 'amet, consectetur adiiscing elit, sed do eiusmod temporincidi.'],
    { x: 9.747, y: 5.188, w: 2.498, h: 0.752, fontSize: 11, color: LIGHT, lineSpacingMultiple: 1.2 });
  text(s, 'From Stress to Strength Neurology in Daily Life',
    { x: 0.944, y: 0.99, w: 7.957, h: 1.313, fontFace: HEAD, fontSize: 36 });
  divider(s, 1.124, 2.711);
}

function slide16(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 16);
  text(s, 'Keeping Your Nerves Strong',
    { x: 0.597, y: 0.638, w: 5.592, h: 1.447, fontFace: HEAD, fontSize: 40 });
  divider(s, 0.711, 2.368);
  [{ x: 7.658, fill: LIGHT, big: '17.23K', bigInk: DARK, cap: 'Neurologist', capInk: GREY },
   { x: 10.26, fill: ORANGE, big: '2,500+', bigInk: LIGHT, cap: 'Protecting', capInk: LIGHT }]
    .forEach((k) => {
      card(s, k.x, 0.629, 2.444, 1.666, 0.1571, k.fill);
      text(s, k.big, {
        x: k.x + 0.225, y: 0.932, w: 1.993, h: 0.707,
        fontFace: HEAD, fontSize: 36, color: k.bigInk, align: 'center',
      });
      text(s, k.cap, { x: k.x + 0.302, y: 1.639, w: 1.84, h: 0.303, fontSize: 12, color: k.capInk, align: 'center' });
    });
  block(s, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere'],
    { x: 8.765, y: 2.453, w: 3.938, h: 0.811, fontSize: 12, color: GREY, align: 'right', lineSpacingMultiple: 1.2 });
  dot(s, 0.743, 2.681);
  text(s, 'Brain Wellness for a Better Life', { x: 0.63, y: 2.931, w: 3.753, h: 0.37, fontFace: HEAD, fontSize: 14 });
  dot(s, 4.811, 2.681);
  text(s, 'Protecting Memory, Focus, and Mood', { x: 4.698, y: 2.931, w: 4.334, h: 0.337, fontFace: HEAD, fontSize: 14 });
}

function slide17(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 17);
  imagePlaceholder(s, 10.414, 0.562, 2.469, 6.375, IMG_BODY_SILHOUETTE, 0.3);
  text(s, 'HUMAN BIOLOGY',
    { x: 8.201, y: 1.052, w: 2.874, h: 0.37, fontFace: 'Open Sans Medium', fontSize: 16, color: PALE });
  const rows = [
    { y: 0.966, lineY: 0.826, lineW: 4.216, dotX: 11.52, label: 'Data One' },
    { y: 2.445, lineY: 2.175, lineW: 4.216, dotX: 11.52, label: 'Data Two' },
    { y: 3.768, lineY: 3.6, lineW: 4.216, dotX: 11.52, label: 'Data Three' },
    { y: 5.393, lineY: 6.529, lineW: 3.717, dotX: 10.934, label: 'Data Four' },
  ];
  rows.forEach((r) => {
    s.addShape('line', { x: 7.329, y: r.lineY, w: r.lineW, h: 0, line: { color: DARK, width: 1.5 } });
    s.addShape('ellipse', {
      x: r.dotX, y: r.lineY - 0.112, w: 0.224, h: 0.224,
      fill: { color: MINT }, line: { color: ORANGE, width: 3, transparency: 42 },
    });
    text(s, r.label, { x: 7.261, y: r.y, w: 3.304, h: 0.337, fontSize: 14 });
    block(s, ['Lorem ipsum dolor sit  amet, consectetur adipiscing elit, sed'],
      { x: 7.261, y: r.y + 0.335, w: 3.304, h: 0.681, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  });
  text(s, 'Neurological Health Data',
    { x: 0.597, y: 1.067, w: 4.758, h: 1.582, fontFace: HEAD, fontSize: 44 });
  divider(s, 0.711, 2.953);
  text(s, 'Frontiers in Brain Research', { x: 0.63, y: 3.425, w: 4.015, h: 0.337, fontFace: HEAD, fontSize: 14 });
  block(s, [LOREM_SHORT], { x: 0.63, y: 3.82, w: 4.47, h: 0.811, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });
  pill(s, 0.714, 5.682, 1.596, 0.418, LIGHT, 'Mental', DARK);
  pill(s, 2.37, 5.682, 1.596, 0.418, ORANGE, 'Health', LIGHT);
}

function slide18(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 18);
  card(s, 0.76, 2.896, 7.042, 3.865, 0.2313, LIGHT);
  text(s, 'Visualizing Research Achievements',
    { x: 0.718, y: 0.772, w: 6.334, h: 1.447, fontFace: HEAD, fontSize: 40 });
  divider(s, 0.784, 2.483);
  block(s, [LOREM_LONG], { x: 7.683, y: 1.025, w: 4.669, h: 1.053, fontSize: 12, color: GREY, lineSpacingMultiple: 1.2 });

  const years = ['2025', '2026', '2027', '2028'];
  s.addChart('line', [
    { name: 'Series 1', labels: years, values: [4.3, 2, 3.5, 4.5] },
    { name: 'Series 2', labels: years, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: years, values: [2, 2, 3, 0] },
  ], {
    x: 1.125, y: 3.064, w: 6.631, h: 3.413,
    chartColors: [MINT, PALE, LIGHT], lineSize: 1.5, lineSmooth: true,
    lineDataSymbol: 'circle', lineDataSymbolSize: 6, lineDataSymbolLineSize: 3, lineDataSymbolLineColor: SAGE,
    showLegend: false, showTitle: false, showValue: false,
    catAxisLineShow: true, catAxisLineColor: 'EDEDED', catAxisLabelColor: DARK,
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 11,
    valAxisLineShow: true, valAxisLineColor: 'EDEDED', valAxisLabelColor: DARK,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 11,
    valAxisLabelFormatCode: '#,##0.0', valAxisMinVal: 0, valAxisMaxVal: 5, valAxisMajorUnit: 0.5,
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
    chartArea: { fill: { color: LIGHT } }, plotArea: { fill: { color: LIGHT } },
  });

  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [40, 25, 35] }], {
    x: 8.052, y: 2.365, w: 4.374, h: 4.947,
    chartColors: [GREY, '75A597', ORANGE], holeSize: 45,
    showLegend: false, showTitle: false, showValue: true,
    dataLabelColor: LIGHT, dataLabelFontFace: HEAD, dataLabelFontSize: 16, dataLabelFormatCode: '0"%"',
    chartArea: { fill: { color: BG } }, plotArea: { fill: { color: BG } },
  });
  imagePlaceholder(s, 9.967, 4.56, 0.575, 0.575, IMG_BRAIN_ON_GREY);
}

function slide19(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 19);
  text(s, 'Questions and discussion', { x: 4.997, y: 0.416, w: 3.339, h: 0.37, fontSize: 16 });
  text(s, 'Q&A and Interactive Conversation',
    { x: 3.22, y: 0.774, w: 6.894, h: 1.447, fontFace: HEAD, fontSize: 40, align: 'center' });
  const testimonial =
    '\u201Cadipiscing elit. amet, , sed. sit amet, , consectetur adipiscing elit, sed do eiusmod tempor. ' +
    'incididunt dolor sit amet, consectetur do\u201D';
  [{ card: 1.968, name: 'Maura Nadine', nameX: 3.385, subX: 3.352, quoteCard: 2.385, quoteX: 2.694 },
   { card: 7.843, name: 'Sofyan Dore', nameX: 9.26, subX: 9.248, quoteCard: 8.26, quoteX: 8.569 }]
    .forEach((c, i) => {
      card(s, c.quoteCard, 4.833, 4.282, 1.583, i === 0 ? 0.1502 : 0.1827, LIGHT);
      card(s, c.card, 3.146, 4.282, 1.094, 0.1823, MINT);
      text(s, c.name, { x: c.nameX, y: 3.322, w: 2.174, h: 0.37, fontSize: 16, color: LIGHT });
      text(s, 'Donec usaeltrices tincidunt',
        { x: c.subX, y: 3.687, w: 2.817, h: 0.326, fontSize: 12, italic: true, color: LIGHT });
      block(s, [testimonial],
        { x: c.quoteX, y: 5.051, w: 3.664, h: 1.053, fontSize: 12, italic: true, color: GREY, lineSpacingMultiple: 1.2 });
    });
}

function slide20(p) {
  const s = p.addSlide({ masterName: 'BASE' });
  chrome(s, 20);
  rings(s, 2.7835, 2.9105, MINT);
  card(s, 6.648, 3.154, 4.571, 2.669, 0.2705, SAGE);
  text(s, 'Contact Us', { x: 6.497, y: 0.649, w: 5.592, h: 0.841, fontFace: HEAD, fontSize: 44 });
  divider(s, 6.611, 1.813);
  block(s, [QUOTE],
    { x: 6.481, y: 2.096, w: 6.038, h: 0.646, fontSize: 14, italic: true, color: MINT, lineSpacingMultiple: 1.2 });
  [{ y: 3.404, iy: 3.545, is: 0.219, ix: 7.012, label: 'loremipsum@yourmail.com', ty: 3.425 },
   { y: 4.207, iy: 4.324, is: 0.267, ix: 6.988, label: '+0123 4567 890', ty: 4.227 },
   { y: 5.008, iy: 5.122, is: 0.272, ix: 6.985, label: 'www.yourwebsite.com', ty: 5.029 }]
    .forEach((r) => {
      s.addShape('ellipse', { x: 6.871, y: r.y, w: 0.5, h: 0.5, fill: { color: MINT } });
      imagePlaceholder(s, r.ix, r.iy, r.is, r.is, IMG_CONTACT_GLYPH);
      text(s, r.label, { x: 7.661, y: r.ty, w: 3.473, h: 0.459, fontSize: 16, color: LIGHT, lineSpacingMultiple: 1.5 });
    });
  pill(s, 6.614, 6.203, 1.596, 0.418, LIGHT, 'Mental', DARK);
  pill(s, 8.32, 6.203, 1.596, 0.418, ORANGE, 'Health', LIGHT);
}

// ------------------------------------------------------------------- build ---

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.defineSlideMaster({ title: 'BASE', background: { color: BG } });
  pptx.author = 'pptxgenjs';
  pptx.title = 'Neurologist — Mental Health Presentation Template';

  [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0feb1c70-d4be-4f33-9a1f-b43fe44f0185_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
