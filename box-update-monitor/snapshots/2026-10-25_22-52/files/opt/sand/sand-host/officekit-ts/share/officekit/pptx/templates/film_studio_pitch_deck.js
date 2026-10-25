/**
 * "Studidy" film-studio pitch deck — 20 slides, 20" x 11.25" (widescreen @ 96 DPI x2).
 * Rebuilt with pptxgenjs only. Photographs in the original are empty picture
 * placeholders, so nothing raster is embedded here.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const INK = '000000';
const PURPLE = '6C37FA'; // theme accent3
const BLUE = '4800FF'; // theme accent1
const GREY = '404040'; // theme tx2, used for body copy
const WHITE = 'FFFFFF';

const HEAD = 'Karla'; // theme major font
const BODY = 'Open Sans'; // theme minor font
const ALT = 'Inter';

/* ------------------------------------------------------- typography presets */
// Most text boxes in the deck sit flush against their frame (zero inset); a few
// small labels keep PowerPoint's default 3.6pt top/bottom padding.
const PAD_V = [0, 0, 3.6, 3.6];

const S = {
  hero: { fontFace: HEAD, fontSize: 96, bold: true, color: INK, lineSpacingMultiple: 0.85 },
  title: { fontFace: HEAD, fontSize: 56, bold: true, color: INK, lineSpacingMultiple: 0.85 },
  toc: { fontFace: HEAD, fontSize: 46, bold: false, color: INK, lineSpacingMultiple: 0.85 },
  step: { fontFace: HEAD, fontSize: 42, bold: true, color: INK, lineSpacingMultiple: 0.85 },
  lead: { fontFace: HEAD, fontSize: 32, color: INK, charSpacing: -0.89, lineSpacingMultiple: 1.1 },
  h3: { fontFace: HEAD, fontSize: 28, bold: true, color: INK, charSpacing: -0.52, lineSpacingMultiple: 1.0 },
  h4: { fontFace: HEAD, fontSize: 24, bold: true, color: INK, margin: PAD_V },
  eyebrow: { fontFace: HEAD, fontSize: 21, bold: true, color: PURPLE, charSpacing: -0.52, lineSpacingMultiple: 1.0 },
  body: { fontFace: BODY, fontSize: 18, color: GREY, charSpacing: -0.44, lineSpacingMultiple: 1.0 },
  stat: { fontFace: HEAD, fontSize: 80, bold: true, color: PURPLE, lineSpacingMultiple: 0.85 },
};

const purple = (t) => ({ text: t, options: { color: PURPLE } });
const ink = (t) => ({ text: t, options: { color: INK } });

/* ------------------------------------------------------------- primitives */
function txt(slide, x, y, w, h, text, opts) {
  slide.addText(text, Object.assign({ x, y, w, h, margin: 0, valign: 'top', isTextBox: true }, opts));
}

// Outlined "card" — the deck's signature rounded rectangle. `r` is the corner radius in inches.
function card(slide, x, y, w, h, r, fill) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: r,
    fill: fill ? { color: fill } : undefined,
    line: { color: INK, width: 1.5 },
  });
}

function rule(slide, x, y, w, width, color) {
  slide.addShape('line', { x, y, w, h: 0, line: { color: color || INK, width: width || 1 } });
}

/* ------------------------------------------------------------------- logo */
// The mark is the union of three overlapping discs, washed with a yellow→orange
// gradient and sliced into 13 horizontal bands. Each band is painted as a few
// rectangles clipped to the disc outline, which gives the scalloped silhouette.
const MARK_W = 0.829, MARK_H = 0.416; // at scale 1
const MARK_GRADIENT = [[0, 'FFFFF0'], [0.06, 'FFF7BC'], [0.24, 'FFDA34'], [0.30, 'FFC614'], [1, 'FF7917']];
const BANDS = 13, BAND_GAP = 0.38; // white slit height as a fraction of the band pitch
const STRIPS = 44; // vertical slices used to paint the gradient

function markColor(t) {
  let i = 1;
  while (i < MARK_GRADIENT.length - 1 && MARK_GRADIENT[i][0] < t) i++;
  const [t0, c0] = MARK_GRADIENT[i - 1], [t1, c1] = MARK_GRADIENT[i];
  const f = (t - t0) / (t1 - t0);
  let out = '';
  for (let ch = 0; ch < 3; ch++) {
    const a = parseInt(c0.substr(ch * 2, 2), 16), b = parseInt(c1.substr(ch * 2, 2), 16);
    out += Math.round(a + (b - a) * f).toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

function logoMark(slide, x, y, k, mono) {
  const r = MARK_H * k / 2;
  const cy = y + r;
  const centers = [0, 1, 2].map((i) => x + r + i * (MARK_W - MARK_H) * k / 2);
  // Paint the silhouette of the three merged discs as vertical gradient slices…
  const sw = MARK_W * k / STRIPS;
  for (let i = 0; i < STRIPS; i++) {
    const t = (i + 0.5) / STRIPS;
    const px = x + t * MARK_W * k;
    const half = Math.max(...centers.map((c) => {
      const dx = Math.abs(px - c);
      return dx < r ? Math.sqrt(r * r - dx * dx) : 0;
    }));
    slide.addShape('rect', {
      x: x + i * sw, y: cy - half, w: sw * 1.03, h: half * 2,
      fill: { color: mono || markColor(t) },
    });
  }
  // …then cut it into bands with white slits.
  const pitch = MARK_H * k / BANDS;
  for (let b = 1; b < BANDS; b++) {
    slide.addShape('rect', {
      x, y: y + b * pitch - pitch * BAND_GAP / 2, w: MARK_W * k, h: pitch * BAND_GAP,
      fill: { color: WHITE },
    });
  }
}

// Horizontal lockup: mark on the left, "Studidy" on the right.
function logo(slide, x, y, scale, mono) {
  const k = scale || 1;
  logoMark(slide, x, y, k, mono);
  txt(slide, x + 0.965 * k, y + 0.060 * k, 1.409 * k, 0.388 * k, 'Studidy', {
    fontFace: HEAD, fontSize: 27 * k, bold: true, color: mono || INK,
    lineSpacingMultiple: 0.85, wrap: false,
  });
}

/* ------------------------------------------------------------- shared copy */
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ';
const ULLAM = 'ullam tempor felis est, quis pellentesque diam tincidunt sed. Nam purus dui, hendrerit sed magna ut, egestas porttitor velit. In hac habitasse platea.';
const ASCONE = 'Ascone proprietary fintech platform of ambiguity. Montfort, as a public and fines company.';
const PELLENTESQUE = 'Pellentesque scelerisque malesuada libero a pellentesque. Morbi orci dui, fermentum eget lectus ornare, viverr viverra ctetur adipiscing elit. ';
const MAURIS = 'Mauris ultrices turpis mi, a ultricies leo fringilla non. Nulla sodales ullamcorper diam vel maximus. Nunc euismod risus eu nibh bibendum, at condimentum dolor interdum.';
const CURABITUR = 'Curabitur suscipit tellus enim, id aliquet nisi maximus eu. Aliquam erat volutpat. Sed feugiat semper lacus, eget blandit neque vehicula ac.';
const UT_VEHICULA = 'Ut vehicula enim eu dui volutpat efficitur. Maecenas scelerisque eget risus sed convallis. Proin venenatis accumsan augue, vel rhoncus elit consectetur eu.';
const UT_VEHICULA_SHORT = 'Ut vehicula enim eu dui volutpat efficitur. Maecenas scelerisque eget risus sed convallis proin venenatis.';
const QUOTE = '"Great films don\u2019t just happen\u2014they are crafted with creativity, vision, and innovation."';
const TAGLINE = 'The Next Big Name in Film Production';
const CINEMA = 'Investing in the Future of Cinema';

/* =================================================================== slides */

// 1 — Title
function slide01(s) {
  txt(s, 0.727, 2.663, 9.568, 5.498,
    [ink('A Studio That Transforms Ideas into '), purple('Blockbusters')], S.hero);
  txt(s, 0.727, 8.319, 7.382, 0.56, TAGLINE, S.lead);
  logo(s, 0.802, 0.731);
  // Two decorative pill outlines bleeding off the right edge.
  card(s, 16.083, 5.947, 3.188, 5.865, 1.594);
  card(s, 11.891, -1.487, 3.188, 5.865, 1.594);
}

// 2 — Table of contents
const TOC = [
  { n: '01', nx: 1.144, ny: 3.704, x: 1.124, y: 4.614, w: 4.977, h: 1.318, t: 'Problem and Solution Client' },
  { n: '02', nx: 7.419, ny: 3.765, x: 7.398, y: 4.648, w: 5.158, h: 1.318, t: 'Market Research Studio Team' },
  { n: '03', nx: 13.873, ny: 3.764, x: 13.853, y: 4.480, w: 5.024, h: 1.318, t: 'Product and Services Studio' },
  { n: '04', nx: 1.118, ny: 7.888, x: 1.098, y: 8.877, w: 5.168, h: 1.318, t: 'Go to  Film Market Strategy' },
  { n: '05', nx: 7.450, ny: 7.888, x: 7.429, y: 8.864, w: 4.951, h: 1.318, t: 'Financial Breakdown' },
  { n: '06', nx: 13.875, ny: 7.888, x: 13.854, y: 8.760, w: 4.415, h: 0.660, t: 'Conclusion' },
];

function slide02(s) {
  txt(s, 0.724, 1.824, 7.385, 0.801, [ink('Table of '), purple('Contents')], S.title);
  TOC.forEach((it, i) => {
    card(s, [0.745, 7.087, 13.483][i % 3], i < 3 ? 3.498 : 7.426, 5.792, 3.1, 0.337);
    txt(s, it.nx, it.ny, 0.988, 0.64, it.n, { fontFace: HEAD, fontSize: 32, color: PURPLE, margin: PAD_V });
    txt(s, it.x, it.y, it.w, it.h, it.t, S.toc);
  });
  logo(s, 0.802, 0.731);
}

// 3 — Three-step process
const STEPS = [
  { cx: 5.510, ex: 6.057, ey: 3.815, ew: 1.761, hx: 6.057, hy: 4.489, hw: 3.436, e: 'STEP 01.', h: 'Pre \u2013 Production.', px: 5.919, py: 7.217, pw: 3.191, ph: 1.899, sx: 6.012, sw: 3.191 },
  { cx: 10.321, ex: 10.769, ey: 3.853, ew: 1.884, hx: 10.762, hy: 4.552, hw: 3.436, e: 'STEP 02.', h: 'Shooting Income.', px: 10.729, py: 7.061, pw: 3.338, ph: 1.418, sx: 10.769, sw: 3.429 },
  { cx: 15.097, ex: 15.541, ey: 3.831, ew: 2.268, hx: 15.399, hy: 4.530, hw: 3.650, e: 'STEP 03.', h: 'Grow Film Gains.', px: 15.506, py: 7.131, pw: 3.436, ph: 1.418, sx: 15.506, sw: 3.328 },
];

function slide03(s) {
  txt(s, 5.51, 0.724, 12.167, 0.801,
    [ink('A Film Studio Built for '), purple('Success')], S.title);
  txt(s, 5.512, 1.793, 7.382, 0.56, TAGLINE, S.lead);
  STEPS.forEach((st) => {
    card(s, st.cx, 3.283, 4.196, 7.243, 0.351);
    rule(s, st.cx + 0.409, 6.395, 3.436, 1);
    txt(s, st.ex, st.ey, st.ew, 0.354, st.e, S.eyebrow);
    txt(s, st.hx, st.hy, st.hw, 1.204, st.h, S.step);
    txt(s, st.px, st.py, st.pw, st.ph, 'Aenean interdum mattis congue donec dignissim congue.',
      Object.assign({}, S.lead, { fontSize: 26 }));
    txt(s, st.sx, 9.408, st.sw, 0.625, LOREM_SHORT, S.body);
  });
  logo(s, 0.802, 0.731);
}

// 4 — Creativity meets business
const FEATURES = [
  { y: 0.749, hy: 1.456, hw: 4.489, by: 2.166, h: 'Bringing Stories to Life', b: 'We provide to unique private credit investments, a rare but valuable part of a investment portfolio.' },
  { y: 4.048, hy: 4.732, hw: 4.486, by: 5.421, h: 'Creative Thinking', b: ASCONE },
  { y: 7.371, hy: 8.122, hw: 3.894, by: 8.811, h: 'Ambience Film', b: ASCONE },
];

function slide04(s) {
  txt(s, 0.727, 1.506, 5.789, 2.406,
    [ink('Where '), purple('Creativity '), ink('Meets Business')], S.title);
  txt(s, 0.727, 8.646, 5.841, 1.963,
    '"Every great film starts with a vision\u2014our studio brings it to life."',
    Object.assign({}, S.lead, { fontSize: 36 }));
  FEATURES.forEach((f, i) => {
    card(s, i === 2 ? 8.457 : 8.445, f.y, 10.818, 3.1, 0.337);
    txt(s, i === 0 ? 13.780 : 13.783, f.hy, f.hw, 0.471, f.h, S.h3);
    txt(s, 13.783, f.by, 4.486, 0.909, f.b, S.body);
  });
  s.addShape('triangle', { x: 1.929, y: 9.723, w: 0.196, h: 0.169, rotate: 30, fill: { color: WHITE } });
  logo(s, 0.802, 0.731);
}

// 5 — Change style
function slide05(s) {
  txt(s, 0.727, 1.643, 5.787, 0.804, [purple('Change'), ink(' Style')], S.title);
  txt(s, 7.108, 1.525, 7.382, 2.072, QUOTE, Object.assign({}, S.lead, { fontSize: 38 }));
  txt(s, 0.727, 8.858, 5.789, 0.471, 'Films is to generate profits', S.h3);
  txt(s, 0.728, 9.617, 5.787, 0.909, ULLAM, S.body);
  txt(s, 15.1, 9.617, 4.174, 0.909,
    'Almost all of our participants report significant improvements, both physically and mentally.', S.body);
  logo(s, 0.802, 0.731);
}

// 6 — Services
const SERVICES = [
  { cy: 3.298, ch: 2.200, r: 0.239, lx: 7.610, ly: 4.145, lw: 3.684, l: 'Design, Ideation & Play', by: 3.731, bw: 6.628, bh: 1.212,
    b: 'Lorem ipsum dolor sit amet, conse lectus ornare, viverra ctetur adipiscing elit. Pellentesque scelerisque malesuada libero a pellentesque. Morbi orci dui, fermentum eget lectus ornare, viverr viverra ctetur adipiscing elit. ' },
  { cy: 6.125, ch: 1.889, r: 0.206, lx: 7.673, ly: 6.877, lw: 3.602, l: 'Film Brainstorming', by: 6.603, bw: 6.465, bh: 0.909, b: PELLENTESQUE },
  { cy: 8.642, ch: 1.889, r: 0.206, lx: 7.673, ly: 9.394, lw: 3.602, l: 'Film Execution', by: 9.119, bw: 6.465, bh: 0.909, b: PELLENTESQUE },
];

function slide06(s) {
  txt(s, 7.108, 0.661, 12.167, 0.785,
    [ink('Our '), purple('Experted Services '), ink('for You')],
    { fontFace: HEAD, fontSize: 56, bold: true, color: INK, lineSpacing: 56 });
  txt(s, 7.108, 1.622, 11.161, 0.56, 'Brand New Vision and Mission in Filmmaking', S.lead);
  SERVICES.forEach((v, i) => {
    card(s, i === 0 ? 7.112 : 7.132, v.cy, 12.167, v.ch, v.r);
    txt(s, v.lx, v.ly, v.lw, 0.353, v.l, S.eyebrow);
    txt(s, 11.891, v.by, v.bw, v.bh, v.b, S.body);
  });
  txt(s, 0.728, 9.278, 4.782, 1.212, ULLAM, S.body);
  logo(s, 0.802, 0.731);
}

// 7 — Next big film project
function slide07(s) {
  txt(s, 0.727, 2.993, 6.007, 2.406,
    [ink('The Next '), purple('Big Film Project '), ink('in Film Industry')], S.title);
  txt(s, 0.727, 5.642, 4.489, 1.153, CINEMA, S.lead);
  txt(s, 0.728, 8.400, 5.787, 0.909, ULLAM, S.body);
  txt(s, 0.727, 9.696, 5.787, 0.909, ULLAM, S.body);
  card(s, 7.108, -1.745, 4.187, 6.933, 2.094);
  card(s, 18.241, 5.551, 4.187, 7.715, 2.094);
  logo(s, 0.802, 0.731);
}

// 8 — About / contact mosaic
function slide08(s) {
  card(s, 0.727, 0.724, 5.809, 2.955, 0.290);
  card(s, 3.917, 4.286, 5.788, 2.864, 0.317);
  card(s, 15.078, 4.286, 4.195, 2.864, 0.300);
  card(s, 0.728, 7.783, 8.979, 2.772, 0.278);
  card(s, 15.077, 7.782, 4.196, 2.773, 0.278);
  logo(s, 0.982, 0.921, 1.678 / 2.374);

  const bigLabel = { fontFace: HEAD, fontSize: 27.99, color: INK, charSpacing: -0.69, lineSpacingMultiple: 1.0 };
  const tinyLabel = { fontFace: HEAD, fontSize: 18, color: INK, charSpacing: -0.44, lineSpacingMultiple: 1.0 };

  txt(s, 1.700, 1.688, 3.324, 0.473, 'Hi, we are Studidy', Object.assign({}, bigLabel, { bold: true }));
  txt(s, 1.700, 2.305, 3.862, 0.625, 'Lead fil production, currently working at mano based in Cairo', S.body);
  txt(s, 4.639, 4.749, 3.175, 0.302, 'About', tinyLabel);
  txt(s, 4.639, 5.204, 4.343, 1.413, 'Passionate about film and enjoy solving problems.', bigLabel);
  txt(s, 15.477, 5.002, 2.941, 0.302, 'Resources', tinyLabel);
  txt(s, 15.477, 5.457, 3.398, 0.977, 'Resource to speed your workflow.', bigLabel);
  txt(s, 1.099, 8.437, 8.237, 1.464, 'Get shooting tips & guides straight to your inbox for free!',
    { fontFace: ALT, fontSize: 40.5, color: INK, margin: [7.2, 7.2, 3.6, 3.6] });
  txt(s, 15.759, 8.680, 2.833, 0.977, 'Have a Project in Mind?', bigLabel);
}

// 9 — Cinematic excellence, four notes
const NOTES = [
  { x: 8.701, hy: 3.889, by: 4.611, h: 'Educational impact' },
  { x: 13.486, hy: 3.892, by: 4.614, h: 'Culinary inspiration' },
  { x: 3.915, hy: 8.207, by: 8.930, h: 'Building community' },
  { x: 8.700, hy: 8.210, by: 8.933, h: 'Celebrating diversity' },
];

function slide09(s) {
  txt(s, 8.701, 0.744, 10.562, 1.605,
    [purple('Cinematic Excellence, '), ink('Unmatched Innovation')], S.title);
  NOTES.forEach((n) => {
    txt(s, n.x, n.hy, 4.194, 0.505, n.h, S.h4);
    txt(s, n.x, n.by, 4.194, 1.515, MAURIS, S.body);
  });
  txt(s, 0.737, 8.207, 2.156, 1.745, 'Academy of film production', S.lead);
  logo(s, 0.802, 0.731);
}

// 10 — Previous project (full-bleed photo grid in the original)
function slide10(s) {
  txt(s, 0.727, 5.949, 7.975, 1.605,
    [purple('Our Previous '), ink('Film Production Project')], S.title);
  txt(s, 0.727, 7.860, 7.382, 0.56, CINEMA, S.lead);
  txt(s, 0.727, 9.579, 5.789, 0.947,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    S.body);
}

// 11 — Pull quote
function slide11(s) {
  txt(s, 8.701, 1.625, 9.762, 5.158, QUOTE,
    { fontFace: HEAD, fontSize: 65, color: INK, charSpacing: -1.4, lineSpacingMultiple: 1.2 });
  txt(s, 8.701, 8.202, 5.883, 0.307, 'Tantina Amira, Bazaarvoice Webinars, 2022',
    { fontFace: HEAD, fontSize: 18, color: INK, charSpacing: -0.44, lineSpacing: 23.4 });
}

// 12 — Editing tips
function slide12(s) {
  txt(s, 0.691, 1.766, 8.875, 0.804, [ink('Editing '), purple('Tips and Trick')], S.title);
  txt(s, 0.727, 2.825, 9.568, 0.56, 'The Future of Film Starts Here: Join the Movement', S.lead);
  const head = Object.assign({}, S.h3, { fontSize: 27.99, charSpacing: -0.69 });
  txt(s, 0.738, 4.352, 4.149, 0.473, 'Sound', head);
  txt(s, 0.738, 5.084, 4.149, 2.238,
    'Avoid using sounds that convey cultural cliches or that reinforce stereotypes of a culture, ethnicity, or region. Instead, consider recording sounds or using music that is locally made, culturally relevant or appropriate to the mood.',
    S.body);
  txt(s, 5.510, 4.352, 4.196, 0.473, 'Colour', head);
  txt(s, 5.510, 5.084, 4.196, 2.243,
    'Low \u2013 and middle \u2013 income countries are often depicted with warm, tropical climates, and bright colours, and edited with filters that convey this. However, these filters also communicate a difference and excotisim that doesn\u2019t reflect reality.',
    S.body);
  txt(s, 0.727, 7.960, 10.573, 2.566,
    'While this is an image photographed in China. It could almost be anywhere in the world. It has an a ambigulty about it that allows the viewer to think about it more as an image of the home, as opposed to an image about China.',
    { fontFace: ALT, fontSize: 28, color: INK, charSpacing: -0.89, lineSpacing: 46.8 });
  logo(s, 0.802, 0.731);
}

// 13 — Two stat cards over a full-bleed photo
const STATS = [
  { cx: 15.082, cy: 0.724, tx: 15.709, tw: 3.424, ly: 2.309, vy: 2.904, vh: 0.861, size: 60,
    label: 'Audience Reach', value: '1.35 m', by: 3.878, bw: 1.968,
    body: [{ text: 'crashes involve', options: { breakLine: true } }, { text: 'human error in the US' }] },
  { cx: 10.295, cy: 5.920, tx: 10.923, tw: 2.274, ly: 7.433, vy: 7.917, vh: 1.149, size: 80,
    label: 'Percentage', value: '94%', by: 9.074, bw: 2.274,
    body: 'deaths worldwide due to vehicle crashes every year' },
];

function slide13(s) {
  txt(s, 0.727, 7.605, 5.788, 2.921,
    'Sed feugiat semper lacus, eget blandit neque vehicula ac. Vivamus vestibulum eleifend aliquam. Quisque hendrerit diam quam.',
    Object.assign({}, S.lead, { color: WHITE }));
  STATS.forEach((st) => {
    card(s, st.cx, st.cy, 4.193, 4.606, 0.482, WHITE);
    txt(s, st.tx, st.ly, st.tw, 0.345, st.label,
      Object.assign({}, S.h4, { margin: 0, lineSpacingMultiple: 0.85 }));
    txt(s, st.tx, st.vy, st.tw, st.vh, st.value, Object.assign({}, S.stat, { fontSize: st.size }));
    txt(s, st.tx, st.by, st.bw, 0.909, st.body, S.body);
  });
  logo(s, 0.727, 0.724, 1, WHITE); // sits on a full-bleed photo in the original
}

// 14 — Future of the team
function slide14(s) {
  txt(s, 10.293, 5.749, 8.977, 1.706,
    [ink('The '), purple('Future '), ink('of Our Film Studio Team')],
    Object.assign({}, S.title, { margin: PAD_V }));
  [['Principle 1', 10.295], ['Principle 2', 15.078]].forEach(([label, x]) => {
    txt(s, x, 8.361, 2.599, 0.505, label, S.h4);
    txt(s, x, 9.011, 4.194, 1.515, UT_VEHICULA, S.body);
  });
  logo(s, 1.136, 1.078, 1, WHITE);
}

// 15 — Join the project
function slide15(s) {
  txt(s, 0.717, 1.978, 4.196, 1.605, [purple('Join'), ink(' the Our Project')], S.title);
  txt(s, 7.108, 1.978, 2.625, 1.414, 'Cras imperdiet vel metus ac condiment.',
    { fontFace: HEAD, fontSize: 28, color: INK, charSpacing: -0.44, lineSpacingMultiple: 1.0 });
  txt(s, 15.078, 1.275, 0.983, 0.471, '03',
    { fontFace: HEAD, fontSize: 28, color: INK, charSpacing: -0.44, lineSpacingMultiple: 1.0 });
  card(s, 10.295, 6.242, 4.220, 4.284, 0.342, WHITE);
  txt(s, 10.880, 6.879, 2.677, 1.144, '89%', S.stat);
  rule(s, 10.880, 8.310, 0.777, 2.25, GREY);
  txt(s, 10.880, 8.704, 2.677, 0.909, 'Promotion Management', S.h4);
  txt(s, 15.078, 7.795, 4.220, 1.818,
    'Praesent convallis ligula eu diam porttitor sodales. Cras imperdiet vel metus ac condimentum. Curabitur fermentum sapien diam, sed volutpat est scelerisque vitae. Suspendisse in sagittis nunc.',
    S.body);
  logo(s, 0.802, 0.731);
}

// 16 — Closing statement with a three-column strip
function slide16(s) {
  card(s, 1.324, 7.983, 17.353, 1.962, 0.327, WHITE);
  txt(s, 0.727, 1.677, 7.383, 2.403,
    [ink('A Studio That Transforms Ideas into '), purple('Blockbusters')], S.title);
  txt(s, 15.290, 1.677, 3.983, 0.606, '@2024 STUDIDY STUDIO ALL RIGHT RESERVED',
    Object.assign({}, S.body, { bold: true, align: 'right' }));
  txt(s, 13.484, 3.409, 5.789, 0.625,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed et leo feugiat, posuere dui ac, laoreet velit. ',
    Object.assign({}, S.body, { align: 'right' }));
  ['Principle 1', 'Principle 2', 'Principle 3'].forEach((label, i) => {
    const x = [1.911, 7.655, 13.398][i];
    txt(s, x, 8.417, 4.691, 0.385, label,
      { fontFace: HEAD, fontSize: 28, bold: true, color: INK, charSpacing: -0.52, lineSpacing: 27.3 });
    txt(s, x, 8.937, i === 0 ? 4.690 : 4.691, 0.625,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.', S.body);
  });
  logo(s, 0.802, 0.731);
}

// 17 — Phone mockup
function phoneMockup(s) {
  s.addShape('roundRect', { x: 8.178, y: 0.785, w: 4.715, h: 9.671, rectRadius: 0.86, fill: { color: 'B2B2B2' } });
  s.addShape('roundRect', { x: 8.276, y: 0.888, w: 4.519, h: 9.465, rectRadius: 0.753, fill: { color: '191919' } });
  s.addShape('roundRect', { x: 8.391, y: 1.002, w: 4.290, h: 9.214, rectRadius: 0.634, fill: { color: WHITE } });
  [[2.789, 0.464], [3.590, 0.719], [4.504, 0.719]].forEach(([y, h]) =>
    s.addShape('rect', { x: 8.110, y, w: 0.068, h, fill: { color: BLUE } }));
}

function slide17(s) {
  phoneMockup(s);
  // Stacked white lockup sitting on the phone screen.
  logoMark(s, 10.100, 5.009, 1, WHITE);
  txt(s, 9.891, 5.533, 1.409, 0.388, 'Studidy',
    { fontFace: HEAD, fontSize: 27, bold: true, color: WHITE, lineSpacingMultiple: 0.85, wrap: false });

  txt(s, 0.727, 1.847, 5.788, 1.736,
    'Sed feugiat semper lacus, eget blandit neque vehicula ac vivamus vestibulum.', S.lead);
  [[0.727, 4.286, 4.504, 5.161, 'Principle 1'], [0.685, 7.263, 7.474, 8.131, 'Principle 2']]
    .forEach(([cx, cy, hy, by, label]) => {
      card(s, cx, cy, 5.355, 2.381, 0.263);
      txt(s, 1.284, hy, 4.194, 0.505, label, Object.assign({}, S.h4, { color: PURPLE }));
      txt(s, 1.284, by, 4.194, 1.212, CURABITUR, S.body);
    });
  logo(s, 0.802, 0.731);
}

// 18 — Locations
const LOCATIONS = [
  { x: 0.740, w: 7.369, h: 'Location A', b: UT_VEHICULA },
  { x: 8.701, w: 4.193, h: 'Location B', b: UT_VEHICULA_SHORT },
  { x: 13.484, w: 4.193, h: 'Location C', b: UT_VEHICULA_SHORT },
];

function slide18(s) {
  txt(s, 8.701, 0.724, 8.977, 1.736,
    'Maecenas scelerisque eget risus sed convallis. Proin venenatis accumsan augue, vel rhoncus elit consectetur eu.',
    S.lead);
  LOCATIONS.forEach((l) => {
    txt(s, l.x, 8.967, 2.599, 0.505, l.h, Object.assign({}, S.h4, { color: PURPLE }));
    txt(s, l.x, 9.617, l.w, 0.909, l.b, S.body);
  });
  logo(s, 1.136, 1.078, 1, WHITE);
}

// 19 — Roadmap
const MILESTONES = [
  { cx: 0.727, cy: 3.703, qx: 1.246, x: 1.229, qy: 3.988, ty: 4.620, by: 5.345, q: 'Q1 2025', t: 'Milestone One' },
  { cx: 7.108, cy: 3.703, qx: 7.626, x: 7.609, qy: 3.988, ty: 4.620, by: 5.345, q: 'Q3 2025', t: 'Milestone Three' },
  { cx: 0.726, cy: 7.282, qx: 1.246, x: 1.229, qy: 7.601, ty: 8.233, by: 8.958, q: 'Q2 2025', t: 'Milestone Two' },
  { cx: 7.106, cy: 7.282, qx: 7.626, x: 7.609, qy: 7.601, ty: 8.233, by: 8.958, q: 'Q4 2025', t: 'Milestone Four' },
];

function slide19(s) {
  txt(s, 0.726, 2.038, 8.808, 0.804,
    [ink('Bringing '), purple('Stories to '), ink('Life')], S.title);
  MILESTONES.forEach((m) => {
    card(s, m.cx, m.cy, 5.788, 3.233, 0.320);
    txt(s, m.qx, m.qy, 2.599, 0.505, m.q, { fontFace: HEAD, fontSize: 24, color: PURPLE, margin: PAD_V });
    txt(s, m.x, m.ty, 4.194, 0.471, m.t,
      { fontFace: HEAD, fontSize: 28, bold: true, color: INK, charSpacing: -0.52, lineSpacingMultiple: 1.0 });
    txt(s, m.x, m.by, 4.194, 1.212, CURABITUR, S.body);
  });
  card(s, 13.484, -0.850, 6.516, 4.067, 0.573, WHITE);
  txt(s, 14.180, 0.883, 5.124, 1.001, 'Sed feugiat semper lacus, eget blandit neque vehicular.',
    Object.assign({}, S.lead, { fontSize: 28 }));
  logo(s, 0.802, 0.731);
}

// 20 — Thanks
function slide20(s) {
  txt(s, 0.727, 1.943, 8.979, 3.439, [purple('Thanks'), ink(' for Watching!')],
    Object.assign({}, S.hero, { fontSize: 120 }));
  txt(s, 2.323, 6.790, 7.382, 1.153,
    '"Movies are more than entertainment; they are experiences that last forever."', S.lead);
  txt(s, 6.516, 9.748, 4.182, 0.774, 'Let,s get in touch together soon as possible',
    Object.assign({}, S.body, { fontSize: 23, paraSpaceAfter: 12 }));
  card(s, 15.157, 4.867, 2.809, 2.809, 0.337);
  logo(s, 0.802, 0.731);
}

/* ==================================================================== build */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'STUDIO', width: 20, height: 11.25 });
pptx.layout = 'STUDIO';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

BUILDERS.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '17e2a2ea-ed48-4cc4-b2bc-c2410a7c665b_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
