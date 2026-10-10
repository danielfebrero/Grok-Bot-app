/**
 * Zenbizz – Business Consulting deck (20 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs only. Picture frames of the source deck become plain
 * tinted boxes and its small vector glyphs become outlined tiles; everything
 * else is a native shape or text run.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */
const C = {
  accent: '4F53EB', // theme accent1 – headlines
  brand: '5F63ED', // theme tx2 – buttons, sub headings
  navy: '181DD3',
  body: '5C5C5C', // tx1 lum 75%
  grey: '929292', // tx1 lum 50%
  l20: 'DFE0FB', // brand lum 20%
  l40: 'BFC1F8', // brand lum 40%
  l60: '9FA1F4', // brand lum 60%
  t1: 'DCDDFB',
  t2: 'B9BAF7',
  t3: '9598F3',
  hair: 'F2F2F2',
  white: 'FFFFFF',
  // picture frames in the source are empty placeholders filled with a 5%
  // accent dot pattern over white, which averages out to this pale tint
  photoFill: 'FAFAFE',
  photoEdge: 'E9EAFB',
};
const HEAD = 'Inter SemiBold'; // theme major font
const BODY = 'Inter'; // theme minor font

/* ---------------------------------------------------------------- helpers */
const txt = (s, text, o) => s.addText(text, Object.assign({ valign: 'top' }, o));

/** big headline (major font, accent blue) */
const headline = (s, text, o) =>
  txt(s, text, Object.assign({ fontFace: HEAD, fontSize: 36, color: C.accent }, o));

/** grey running text, 11pt / 150% leading */
const para = (s, text, o) =>
  txt(s, text, Object.assign(
    { fontFace: BODY, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 }, o));

/** small blue caption above a paragraph */
const label = (s, text, o) =>
  txt(s, text, Object.assign({ fontFace: HEAD, fontSize: 14, color: C.brand }, o));

/** thin decorative rule */
const rule = (s, x, y, w, color, width) =>
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color || C.l40, width: width || 0.5 } });

/** solid rounded card; r = corner radius in inches */
const card = (s, x, y, w, h, fill, r) =>
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: r });

/** filled circle carrying a short label (numbers "01", "1", ...) */
const bubble = (s, x, y, d, text, o) =>
  s.addText(text || '', Object.assign({
    shape: 'ellipse', x: x, y: y, w: d, h: d,
    fill: { color: C.l20 }, line: { color: C.l20, width: 0 },
    align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 14, color: C.brand,
  }, o));

/** round "task done" check badge used next to intro paragraphs */
const checkBadge = (s, x, y, d) =>
  s.addText('\u2713', {
    shape: 'ellipse', x: x, y: y, w: d, h: d,
    fill: { color: C.l60 }, line: { color: C.l60, width: 0 },
    align: 'center', valign: 'middle', fontFace: BODY, fontSize: 12, bold: true, color: C.white,
  });

/** stand-in for a picture frame of the original deck */
const photo = (s, x, y, w, h, r) =>
  s.addShape(r ? 'roundRect' : 'rect', {
    x: x, y: y, w: w, h: h, rectRadius: r || undefined,
    fill: { color: C.photoFill }, line: { color: C.photoEdge, width: 0.5 },
  });

/** placeholder standing in for one of the small vector glyphs */
const iconTile = (s, x, y, d) =>
  s.addShape('roundRect', { x: x, y: y, w: d, h: d, fill: { type: 'none' }, line: { color: C.brand, width: 1.25 }, rectRadius: d * 0.22 });

/** purple block with one rounded corner that bleeds off the lower-left edge */
const cornerBlock = (s) =>
  s.addShape('round1Rect', { x: 0, y: 5.118, w: 2.701, h: 2.382, fill: { color: C.brand }, line: { color: C.brand, width: 0 }, rectRadius: 0.201 });

/* -------------------------------------------------------- shared fragments */
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, ';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.';
const LOREM_CONGUE = LOREM_SHORT + ', ';

/** "Business Consulting" pill used on the two opening slides */
function pill(s, x, y) {
  s.addShape('roundRect', { x: x, y: y, w: 2.175, h: 0.433, fill: { color: C.brand }, line: { color: C.brand, width: 0 }, rectRadius: 0.2165 });
  txt(s, 'Business Consulting', {
    x: x + 0.146, y: y + 0.073, w: 1.884, h: 0.286,
    align: 'center', fontFace: HEAD, fontSize: 11, color: C.l20,
  });
}

/** headline + rule + caption + paragraph block reused on several slides */
function zenbizzTitle(s, x, y, w, h) {
  headline(s, [
    { text: 'Start Improving Your Business Today with ' },
    { text: 'Zenbizz', options: { italic: true } },
  ], { x: x, y: y, w: w, h: h });
}

/* ----------------------------------------------------------------- slides */
function slide01(s) {
  photo(s, 1.268, 0.965, 2.995, 5.544, 0.289);
  txt(s, 'w w w . w e b s i t e . c o m', { x: 9.201, y: 0.864, w: 2.934, h: 0.303, align: 'right', fontFace: BODY, fontSize: 12, color: C.grey });
  pill(s, 9.941, 2.937);
  txt(s, 'ZENBIZZ', { x: 5.276, y: 3.505, w: 7.037, h: 2.036, align: 'right', wrap: false, fontFace: HEAD, fontSize: 115, color: C.accent });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna', { x: 6.945, y: 5.495, w: 5.174, h: 0.624, align: 'right' });
}

function slide02(s) {
  photo(s, 0, 0, 13.333, 4.175, 0);
  pill(s, 1.219, 4.765);
  txt(s, 'Consulting Your Business with Us', { x: 1.132, y: 5.333, w: 5.614, h: 1.447, fontFace: HEAD, fontSize: 40, color: C.brand });
  para(s, LOREM_LONG + 'purus lectus malesuada libero, sit amet commodo magna eros quis urna.', { x: 7.331, y: 5.465, w: 4.948, h: 1.179 });
}

function slide03(s) {
  const rows = [
    { n: '01', t: 'Introduction', ny: 1.894, nx: 1.331, nw: 0.491, ty: 1.814, tw: 2.535, oy: 1.789 },
    { n: '02', t: 'Our Vision', ny: 2.854, nx: 1.313, nw: 0.526, ty: 2.774, tw: 2.181, oy: 2.740 },
    { n: '03', t: 'Our Mission', ny: 3.813, nx: 1.309, nw: 0.535, ty: 3.733, tw: 2.469, oy: 3.713 },
    { n: '04', t: 'Revenue Growth', ny: 4.772, nx: 1.307, nw: 0.539, ty: 4.692, tw: 3.370, oy: 4.665 },
    { n: '05', t: 'Conclusions', ny: 5.731, nx: 1.312, nw: 0.530, ty: 5.651, tw: 2.551, oy: 5.614 },
  ];
  // vertical spine behind the numbered bullets (comes from the layout)
  s.addShape('line', { x: 1.588, y: 2.0, w: 0, h: 3.912, line: { color: C.hair, width: 3.25 } });
  rows.forEach((r) => {
    s.addShape('ellipse', { x: 1.274, y: r.oy, w: 0.627, h: 0.627, fill: { color: C.t1 }, line: { color: C.t1, width: 0 } });
    txt(s, r.n, { x: r.nx, y: r.ny, w: r.nw, h: 0.404, align: 'center', wrap: false, fontFace: HEAD, fontSize: 18, color: C.brand });
    txt(s, r.t, { x: 2.114, y: r.ty, w: r.tw, h: 0.572, wrap: false, fontFace: HEAD, fontSize: 28, color: C.brand });
  });
  photo(s, 7.332, 2.842, 4.69, 3.316, 0.32);
  headline(s, 'Section Lists', { x: 6.481, y: 0.986, w: 5.613, h: 0.909, align: 'right', fontSize: 48 });
  rule(s, 6.847, 2.167, 5.175);
}

function slide04(s) {
  photo(s, 1.268, 0.965, 2.995, 5.544, 0.289);
  [
    { n: '01', oy: 1.036, ty: 1.826, t: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.' },
    { n: '02', oy: 3.043, ty: 3.833, t: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' },
    { n: '03', oy: 5.051, ty: 5.840, t: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' },
  ].forEach((it) => {
    bubble(s, 4.875, it.oy, 0.667, it.n);
    para(s, it.t, { x: 4.834, y: it.ty, w: 2.683, h: 0.624 });
  });
  headline(s, 'A Business Consultancy that Can Produce Anything', { x: 8.229, y: 1.04, w: 3.836, h: 3.13, align: 'left' });
  rule(s, 8.229, 4.498, 3.194);
  label(s, 'Consultation Business', { x: 8.229, y: 5.047, w: 2.538, h: 0.337 });
  para(s, LOREM_LONG, { x: 8.229, y: 5.562, w: 4.011, h: 0.902 });
}

function slide05(s) {
  card(s, 1.19, 3.242, 2.89, 3.281, C.brand, 0.398);
  headline(s, "We're Proud to Knowing Our History", { x: 1.099, y: 0.882, w: 6.178, h: 1.178, fontSize: 32 });
  checkBadge(s, 7.325, 1.304, 0.353);
  para(s, LOREM_LONG, { x: 7.901, y: 1.03, w: 4.144, h: 0.902 });
  rule(s, 1.051, 2.389, 3.194);
  const cols = [
    { x: 1.464, oy: 3.535, ty: 4.664, by: 5.050, n: '1', t: 'Good Business', numColor: C.brand, titleColor: C.l20 },
    { x: 4.512, oy: 3.703, ty: 4.832, by: 5.218, n: '2', t: 'Planning Ideas', numColor: C.l20, titleColor: C.l60 },
    { x: 7.085, oy: 3.703, ty: 4.832, by: 5.218, n: '3', t: 'Creative Process', numColor: C.l20, titleColor: C.l60 },
    { x: 9.606, oy: 3.703, ty: 4.832, by: 5.218, n: '4', t: 'Financial Planning', numColor: C.l20, titleColor: C.l60 },
  ];
  cols.forEach((c) => {
    bubble(s, c.x, c.oy, 0.851, c.n, { fill: { color: C.l40 }, line: { color: C.l40, width: 0 }, fontSize: 24, color: c.numColor });
    label(s, c.t, { x: c.x, y: c.ty, w: 2.538, h: 0.337, color: c.titleColor });
    para(s, LOREM_MED + '.', { x: c.x, y: c.by, w: 2.308, h: 1.179, color: C.l40 });
  });
}

function slide06(s) {
  card(s, 1.248, 0.895, 4.243, 1.73, C.l20, 0.238);
  photo(s, 1.248, 2.842, 4.69, 3.316, 0.32);
  cornerBlock(s);
  txt(s, '+88%', { x: 1.436, y: 1.049, w: 1.934, h: 0.841, wrap: false, fontFace: HEAD, fontSize: 44, color: C.brand });
  txt(s, LOREM_TINY, { x: 1.516, y: 1.946, w: 3.686, h: 0.505, fontFace: BODY, fontSize: 12, color: C.brand });
  zenbizzTitle(s, 6.93, 1.743, 5.303, 1.919);
  rule(s, 6.842, 4.389, 3.194);
  checkBadge(s, 6.976, 5.011, 0.353);
  para(s, LOREM_LONG, { x: 7.512, y: 4.913, w: 4.614, h: 0.902 });
}

function slide07(s) {
  photo(s, 1.248, 3.75, 4.278, 2.689, 0.29);
  photo(s, 8.655, 0.881, 3.321, 2.533, 0.244);
  cornerBlock(s);
  headline(s, 'Plan A Marketing Strategy With Targeted Goals', { x: 1.125, y: 1.056, w: 7.026, h: 1.313 });
  rule(s, 1.086, 2.775, 3.194);
  [
    { t: 'Marketing Strategy', cy: 4.018, ty: 3.923, by: 4.323 },
    { t: 'Financial Planning', cy: 5.336, ty: 5.241, by: 5.642 },
  ].forEach((it) => {
    checkBadge(s, 6.569, it.cy, 0.353);
    label(s, it.t, { x: 7.106, y: it.ty, w: 2.538, h: 0.337 });
    para(s, LOREM_MED + ', ', { x: 7.106, y: it.by, w: 4.614, h: 0.624 });
  });
}

function slide08(s) {
  card(s, 1.225, 3.794, 2.559, 2.525, C.l60, 0.246);
  card(s, 4.034, 3.794, 2.559, 2.525, C.l60, 0.246);
  card(s, 6.781, 2.695, 5.243, 3.624, C.brand, 0.282);
  headline(s, 'Discover What We do to Take Care of Your Future', { x: 1.11, y: 1.033, w: 8.067, h: 1.313 });
  rule(s, 1.086, 2.775, 3.194);
  [
    { x: 1.412, ox: 1.524, n: '01', t: 'Strategy Operation', b: LOREM_TINY + ' ' },
    { x: 4.221, ox: 4.281, n: '02', t: 'Taxes Efficiency', b: LOREM_TINY + ' ' },
  ].forEach((cd) => {
    bubble(s, cd.ox, 4.007, 0.667, cd.n);
    label(s, cd.t, { x: cd.x, y: 4.881, w: 2.203, h: 0.337, fontFace: BODY, color: C.white });
    para(s, cd.b, { x: cd.x, y: 5.205, w: 2.203, h: 0.902, color: C.l20 });
  });
  s.addShape('ellipse', { x: 7.132, y: 2.9, w: 0.608, h: 0.606, fill: { color: C.l20 }, line: { color: C.l20, width: 0 } });
  iconTile(s, 7.29, 3.058, 0.292);
  txt(s, 'Financial & Restructruring', { x: 7.072, y: 3.775, w: 4.841, h: 0.438, fontFace: BODY, fontSize: 20, color: C.white });
  para(s, LOREM_SHORT, { x: 7.072, y: 4.311, w: 4.348, h: 0.624, color: C.l20 });
  [7.072, 9.412].forEach((x) => {
    s.addText([1, 2, 3].map(() => ({
      text: 'Lorem ipsum dolor sit amet, ',
      options: { bullet: { characterCode: '2022', indent: 13.5 }, paraSpaceAfter: 6 },
    })), {
      x: x, y: 5.096, w: 2.613, h: 0.997, valign: 'top',
      fontFace: BODY, fontSize: 10, color: C.l20, lineSpacingMultiple: 1.5,
    });
  });
}

function slide09(s) {
  photo(s, 0, 2.947, 13.333, 1.834, 0);
  headline(s, 'Unveiling the Impact Business Growth', { x: 1.11, y: 1.033, w: 6.293, h: 1.313 });
  checkBadge(s, 7.512, 1.578, 0.353);
  para(s, LOREM_LONG, { x: 8.089, y: 1.303, w: 4.144, h: 0.902 });
  [
    { cx: 1.224, ox: 1.506, vx: 1.363, vw: 1.95, lx: 1.394, lw: 1.744, v: '146,987', t: 'Success Business' },
    { cx: 4.180, ox: 4.461, vx: 4.318, vw: 2.46, lx: 4.350, lw: 1.744, v: '$12,000', t: 'Average Revenue' },
    { cx: 6.953, ox: 7.235, vx: 7.092, vw: 1.95, lx: 7.123, lw: 1.950, v: '85%', t: 'Financial Growth' },
    { cx: 9.817, ox: 10.098, vx: 9.955, vw: 1.95, lx: 9.986, lw: 1.744, v: '50K++', t: 'Active Clients' },
  ].forEach((k) => {
    card(s, k.cx, 4.033, 2.184, 2.479, C.brand, 0.301);
    bubble(s, k.ox, 4.207, 0.667, '01');
    txt(s, k.v, { x: k.vx, y: 5.183, w: k.vw, h: 0.64, fontFace: BODY, fontSize: 32, color: C.white });
    para(s, k.t, { x: k.lx, y: 5.765, w: k.lw, h: 0.346, color: C.l20 });
  });
}

function slide10(s) {
  photo(s, 1.248, 1.267, 4.278, 3.502, 0.337);
  para(s, LOREM_LONG, { x: 1.373, y: 5.394, w: 4.144, h: 0.902 });
  headline(s, 'PLACEHOLDER', { x: 7.134, y: 1.365, w: 4.931, h: 2.524, align: 'left' });
  rule(s, 7.081, 4.498, 3.194);
  label(s, 'Consultation Business', { x: 7.134, y: 4.995, w: 2.538, h: 0.337 });
  para(s, LOREM_LONG, { x: 7.134, y: 5.422, w: 4.931, h: 0.902 });
}

function slide11(s) {
  photo(s, 0, 0, 13.333, 2.474, 0);
  headline(s, 'Decoding the Secrets of Our Consultancy', { x: 2.292, y: 3.031, w: 8.749, h: 1.313, align: 'center' });
  [
    { x: 1.735, lx: 1.733, lw: 2.05, v: '320+', t: 'Objective Business Consulting' },
    { x: 4.356, lx: 4.354, lw: 2.159, v: '$22,00', t: 'Start-Up Funding Projects' },
  ].forEach((k) => {
    txt(s, k.v, { x: k.x, y: 5.159, w: 2.324, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.brand });
    txt(s, k.t, { x: k.lx, y: 5.97, w: k.lw, h: 0.572, fontFace: BODY, fontSize: 14, color: C.brand });
  });
  rule(s, 7.303, 5.254, 3.194);
  para(s, LOREM_LONG, { x: 7.25, y: 5.607, w: 4.348, h: 0.902 });
}

function slide12(s) {
  s.addShape('line', { x: 1.568, y: 1.299, w: 0, h: 4.386, line: { color: C.hair, width: 3.25 } });
  [
    { y: 1.127, by: 1.464, t: 'Branding Problems' },
    { y: 2.519, by: 2.855, t: 'Branding Solutions' },
    { y: 3.868, by: 4.205, t: 'Campaign Brief' },
    { y: 5.315, by: 5.652, t: 'Campaign Execution' },
  ].forEach((it) => {
    s.addShape('ellipse', { x: 1.274, y: it.y, w: 0.589, h: 0.589, fill: { color: C.t1 }, line: { color: C.t1, width: 0 } });
    iconTile(s, 1.394, it.y + 0.12, 0.348);
    label(s, it.t, { x: 2.071, y: it.y, w: 3.69, h: 0.337 });
    para(s, LOREM_SHORT, { x: 2.075, y: it.by, w: 3.69, h: 0.624 });
  });
  zenbizzTitle(s, 6.93, 1.743, 5.303, 1.919);
  rule(s, 6.842, 4.389, 3.194);
  checkBadge(s, 6.976, 5.011, 0.353);
  para(s, LOREM_LONG, { x: 7.512, y: 4.913, w: 4.614, h: 0.902 });
}

function slide13(s) {
  headline(s, 'Masters of Strategy : A Showcase of Consulting Brilliance', { x: 1.097, y: 1.075, w: 8.75, h: 1.313 });
  rule(s, 1.611, 3.318, 8.905, C.l20);
  [
    { ox: 1.220, tx: 1.122, tw: 2.137, nx: 1.081, n: '01' },
    { ox: 4.319, tx: 4.111, tw: 1.977, nx: 4.070, n: '02' },
    { ox: 7.417, tx: 7.276, tw: 2.035, nx: 7.235, n: '03' },
    { ox: 10.516, tx: 10.339, tw: 1.801, nx: 10.298, n: '04' },
  ].forEach((it) => {
    s.addShape('ellipse', { x: it.ox, y: 3.123, w: 0.391, h: 0.391, fill: { color: C.t2 }, line: { color: C.t2, width: 0 } });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer.', { x: it.tx, y: 3.855, w: it.tw, h: 0.909, fontFace: BODY, fontSize: 16, color: C.body });
    txt(s, it.n, { x: it.nx, y: 4.978, w: 1.863, h: 1.447, fontFace: HEAD, fontSize: 80, color: C.accent });
  });
}

function slide14(s) {
  photo(s, 1.268, 0.965, 2.023, 2.547, 0.306);
  photo(s, 9.236, 4.401, 2.74, 2.076, 0.200);
  headline(s, 'A Business Consultancy that Can Produce Anything', { x: 4.295, y: 1.37, w: 7.205, h: 1.313, align: 'left' });
  rule(s, 4.433, 3.106, 3.194);
  [
    { x: 1.124, ox: 1.223, n: '01', t: 'Business Planning' },
    { x: 3.690, ox: 3.789, n: '02', t: 'Marketing Strategy' },
    { x: 6.256, ox: 6.355, n: '03', t: 'Financial Consulting' },
  ].forEach((it) => {
    bubble(s, it.ox, 4.293, 0.537, it.n, { fontSize: 9 });
    label(s, it.t, { x: it.x, y: 5.161, w: 2.538, h: 0.337 });
    para(s, LOREM_CONGUE, { x: it.x, y: 5.561, w: 2.495, h: 0.902 });
  });
}

function slide15(s) {
  s.addShape('rect', { x: 0, y: 0, w: 3.73, h: 7.5, fill: { color: C.brand }, line: { color: C.brand, width: 0 } });
  [
    { x: 1.254, y: 0.757, d: 5.986, fill: C.t1, tx: 3.608, ty: 0.946, tw: 1.278, v: '$90k', tc: C.accent },
    { x: 1.763, y: 1.775, d: 4.968, fill: C.t2, tx: 3.643, ty: 1.963, tw: 1.208, v: '$57k', tc: C.white },
    { x: 2.272, y: 2.792, d: 3.951, fill: C.t3, tx: 3.632, ty: 2.981, tw: 1.231, v: '$23k', tc: C.white },
    { x: 2.751, y: 3.750, d: 2.993, fill: C.accent, tx: 3.625, ty: 4.927, tw: 1.243, v: '$12k', tc: C.white },
  ].forEach((r) => {
    s.addShape('ellipse', { x: r.x, y: r.y, w: r.d, h: r.d, fill: { color: r.fill }, line: { color: r.fill, width: 0 } });
    s.addText(r.v, { x: r.tx, y: r.ty, w: r.tw, h: 0.64, align: 'center', valign: 'middle', wrap: false, fontFace: HEAD, fontSize: 32, color: r.tc });
  });
  headline(s, 'Unveiling the Impact Business Growth', { x: 8.315, y: 1.516, w: 4.282, h: 1.616, fontSize: 30 });
  rule(s, 8.315, 3.62, 2.789);
  txt(s, '318K', { x: 8.315, y: 4.368, w: 2.979, h: 1.447, wrap: false, fontFace: HEAD, fontSize: 80, color: C.accent });
  txt(s, 'Target Completed', { x: 8.315, y: 5.776, w: 3.668, h: 0.572, wrap: false, fontFace: BODY, fontSize: 28, color: C.brand });
}

function slide16(s) {
  headline(s, 'Discover What We do to Take Care of Your Future', { x: 1.11, y: 0.936, w: 8.067, h: 1.313 });
  rule(s, 1.086, 2.677, 3.194);
  [
    { x: 1.235, y: 3.367, d: 3.086, fill: C.l40, vx: 1.427, vy: 4.288, vw: 2.702, v: '$12,000', vc: C.brand, lx: 1.580, ly: 5.211, lw: 2.395, l: 'Transaction per Year', lc: C.brand },
    { x: 4.385, y: 2.972, d: 3.481, fill: C.brand, vx: 5.247, vy: 4.073, vw: 1.757, v: '85k+', vc: C.white, lx: 5.311, ly: 5.054, lw: 1.629, l: 'Repeat Client', lc: C.l40 },
    { x: 7.939, y: 2.414, d: 4.039, fill: C.accent, vx: 8.602, vy: 3.771, vw: 2.712, v: '120,000', vc: C.white, lx: 9.243, ly: 4.833, lw: 1.429, l: 'Active User', lc: C.l40 },
  ].forEach((b) => {
    s.addShape('ellipse', { x: b.x, y: b.y, w: b.d, h: b.d, fill: { color: b.fill }, line: { color: b.fill, width: 0 } });
    s.addText(b.v, { x: b.vx, y: b.vy, w: b.vw, h: 0.841, align: 'center', valign: 'middle', wrap: false, fontFace: HEAD, fontSize: 44, color: b.vc });
    s.addText(b.l, { x: b.lx, y: b.ly, w: b.lw, h: 0.37, align: 'center', valign: 'middle', wrap: false, fontFace: BODY, fontSize: 16, color: b.lc });
  });
}

function slide17(s) {
  // 14 column bars, right-to-left years, baseline at y = 6.234
  const BASE = 6.234, BARW = 0.385, PITCH = 0.79245, X0 = 1.224;
  const bars = [
    { y: '2029', h: 4.605, c: C.accent }, { y: '2028', h: 4.414, c: C.l60 },
    { y: '2027', h: 3.975, c: C.t1 }, { y: '2026', h: 3.642, c: C.accent },
    { y: '2025', h: 3.060, c: C.l60 }, { y: '2024', h: 3.308, c: C.t1 },
    { y: '2023', h: 2.935, c: C.accent }, { y: '2022', h: 2.726, c: C.l60 },
    { y: '2021', h: 2.506, c: C.t1 }, { y: '2020', h: 1.867, c: C.accent },
    { y: '2019', h: 2.125, c: C.l60 }, { y: '2018', h: 1.867, c: C.accent },
    { y: '2017', h: 1.145, c: C.l60 }, { y: '2016', h: 1.145, c: C.t1 },
  ];
  bars.forEach((b, i) => {
    const x = X0 + i * PITCH;
    s.addShape('rect', { x: x, y: BASE - b.h, w: BARW, h: b.h, fill: { color: b.c }, line: { color: b.c, width: 0 } });
    txt(s, b.y, { x: x + BARW / 2 - 0.29, y: 6.272, w: 0.58, h: 0.286, align: 'center', wrap: false, fontFace: BODY, fontSize: 11, color: C.brand });
  });
  rule(s, 1.106, BASE, 10.891);
  headline(s, 'Unveiling the Impact Business Growth', { x: 7.863, y: 0.931, w: 3.984, h: 1.616, fontSize: 30 });
  rule(s, 7.754, 2.808, 3.194);
}

function slide18(s) {
  // five rounded panels come from the slide layout
  card(s, 1.225, 2.069, 3.253, 2.342, C.brand, 0.104);
  card(s, 4.662, 2.069, 3.758, 2.342, C.accent, 0.104);
  card(s, 8.604, 2.069, 3.396, 2.342, C.brand, 0.104);
  card(s, 1.204, 4.571, 7.467, 2.342, C.l60, 0.104);
  card(s, 8.855, 4.571, 3.184, 2.342, C.accent, 0.104);
  headline(s, 'Decoding the Secrets of Our Consultancy', { x: 1.165, y: 0.616, w: 5.893, h: 1.178, fontSize: 32 });

  // panel 1 – copy
  label(s, 'Let\u2019s do AI Marketing Tools', { x: 1.4, y: 2.299, w: 3.337, h: 0.337, color: C.l20 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.', { x: 1.4, y: 2.739, w: 2.995, h: 0.624, color: C.l40 });
  s.addText([1, 2].map(() => ({
    text: 'Lorem ipsum dolor sit amet consec',
    options: { bullet: { characterCode: '2022', indent: 13.5 }, paraSpaceAfter: 6 },
  })), { x: 1.4, y: 3.421, w: 2.992, h: 0.708, valign: 'top', fontFace: BODY, fontSize: 11, color: C.l40, lineSpacingMultiple: 1.5 });

  // panel 2 – stacked progress bars + percent scale
  txt(s, '2024 KPI for AI Marketing Tools', { x: 4.812, y: 2.299, w: 3.461, h: 0.303, fontFace: HEAD, fontSize: 12, color: C.l20 });
  const segColors = [C.l20, C.l40, C.l60];
  [
    { y: 2.817, w: [1.363, 0.699, 1.168] },
    { y: 3.225, w: [0.902, 1.610, 0.722] },
    { y: 3.644, w: [1.597, 0.886, 0.748] },
  ].forEach((row) => {
    let x = 4.917;
    row.w.forEach((w, i) => {
      s.addShape('rect', { x: x, y: row.y, w: w, h: 0.162, fill: { color: segColors[i] }, line: { color: segColors[i], width: 0 } });
      x += w;
    });
  });
  ['0%', '10%', '20%', '30%', '40%', '50%', '60%', '70%', '80%', '90%', '100%'].forEach((t, i) => {
    txt(s, t, { x: 4.779 + i * 0.3083, y: 3.918, w: 0.44, h: 0.202, align: 'center', fontFace: BODY, fontSize: 6, color: C.l20 });
  });

  // panel 3 – KPI figure
  s.addShape('ellipse', { x: 8.912, y: 2.374, w: 0.499, h: 0.5, fill: { color: C.l40 }, line: { color: C.l40, width: 0 } });
  iconTile(s, 9.078, 2.54, 0.167);
  txt(s, '$12,000', { x: 8.912, y: 3.16, w: 2.55, h: 0.774, fontFace: HEAD, fontSize: 40, color: C.l20 });
  txt(s, 'Average Revenue', { x: 8.912, y: 3.912, w: 2.55, h: 0.286, fontFace: BODY, fontSize: 11, color: C.l40 });

  // panel 4 – grouped column chart drawn as plain rectangles, 0..6 scale
  const GBASE = 6.437, UNIT = 0.2518;
  for (let i = 0; i <= 6; i++) rule(s, 1.803, GBASE - i * UNIT, 6.479, C.l40, 0.72);
  const groups = [
    [4.28, 2.41, 1.98], [2.51, 4.39, 1.98], [3.50, 1.80, 3.01], [4.49, 2.80, 4.99], [2.51, 4.39, 1.98],
  ];
  const barFill = [C.l20, C.brand, C.navy];
  groups.forEach((vals, g) => {
    vals.forEach((v, i) => {
      const x = 2.079 + g * 1.2945 + i * 0.2485;
      const h = v * UNIT;
      s.addShape('rect', { x: x, y: GBASE - h, w: 0.248, h: h, fill: { color: barFill[i] }, line: { color: barFill[i], width: 0 } });
    });
    txt(s, 'Category ' + (g + 1), { x: 2.04 + g * 1.2925, y: 6.476, w: 0.87, h: 0.252, align: 'center', wrap: false, fontFace: BODY, fontSize: 9, color: C.l20 });
  });
  for (let v = 0; v <= 6; v++) {
    txt(s, String(v), { x: 1.547, y: GBASE - v * UNIT - 0.126, w: 0.281, h: 0.252, align: 'center', wrap: false, fontFace: BODY, fontSize: 9, color: C.l20 });
  }

  // panel 5 – partnership call-out
  txt(s, [{ text: '2026', options: { breakLine: true } }, { text: 'January' }], { x: 8.967, y: 4.746, w: 0.858, h: 0.505, wrap: false, fontFace: HEAD, fontSize: 12, color: C.l20 });
  txt(s, 'Business Partnership', { x: 9.709, y: 5.816, w: 2.189, h: 0.909, align: 'right', fontFace: HEAD, fontSize: 24, color: C.l20 });
}

function slide19(s) {
  [
    { x: 1.231, y: 6.104, h: 1.396, fill: C.l40, r: 0.116, v: '$20M', vc: C.accent, vy: 6.223, yr: '2023', yy: 5.566, yx: 1.350, vx: 1.273 },
    { x: 5.122, y: 3.968, h: 3.532, fill: C.t3, r: 0.106, v: '$80M', vc: C.l20, vy: 4.087, yr: '2024', yy: 3.461, yx: 5.241, vx: 5.164 },
    { x: 9.014, y: 2.167, h: 5.333, fill: C.accent, r: 0.106, v: '$120M', vc: C.l20, vy: 2.285, yr: '2025', yy: 1.660, yx: 9.133, vx: 9.056 },
  ].forEach((b) => {
    s.addShape('round2SameRect', { x: b.x, y: b.y, w: 3.019, h: b.h, fill: { color: b.fill }, line: { color: b.fill, width: 0 }, rectRadius: b.r });
    txt(s, b.v, { x: b.vx, y: b.vy, w: 2.935, h: 0.909, align: 'center', fontFace: HEAD, fontSize: 48, color: b.vc });
    txt(s, b.yr, { x: b.yx, y: b.yy, w: 2.781, h: 0.438, align: 'center', fontFace: HEAD, fontSize: 20, color: C.brand });
  });
  headline(s, 'Plan A Marketing Strategy With Targeted Goals', { x: 1.125, y: 1.056, w: 7.026, h: 1.313 });
  rule(s, 1.086, 3.056, 3.194);
  para(s, LOREM_CONGUE, { x: 1.124, y: 3.461, w: 3.157, h: 0.902 });
}

function slide20(s) {
  photo(s, 1.282, 2.179, 10.77, 3.143, 0.136);
  txt(s, [
    { text: 'THANK ', options: { bold: true } },
    { text: 'YOU' },
  ], { x: 1.282, y: 0.979, w: 4.44, h: 0.909, fontFace: HEAD, fontSize: 48, color: C.accent });
  rule(s, 7.611, 1.583, 4.226);
  txt(s, 'For Watching Our Presentation', { x: 7.847, y: 1.058, w: 3.848, h: 0.404, align: 'right', wrap: false, fontFace: BODY, fontSize: 18, color: C.brand });
  [
    { x: 1.282, dx: 1.407, t: 'Mail' },
    { x: 4.111, dx: 4.236, t: 'Website' },
    { x: 7.065, dx: 7.190, t: 'Phone Number' },
    { x: 10.145, dx: 10.270, t: 'Address' },
  ].forEach((c) => {
    s.addShape('ellipse', { x: c.dx, y: 5.867, w: 0.13, h: 0.13, fill: { color: C.brand }, line: { color: C.brand, width: 0 } });
    s.addText(c.t, { x: c.x, y: 6.108, w: 2.265, h: 0.404, valign: 'bottom', fontFace: HEAD, fontSize: 18, bold: true, color: C.brand });
    para(s, 'Type here to edit', { x: c.x, y: 6.472, w: 2.265, h: 0.346 });
  });
}

/* ------------------------------------------------------------------- build */
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'ZENBIZZ', width: 13.333, height: 7.5 });
pptx.layout = 'ZENBIZZ';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Zenbizz Business Consulting';

SLIDES.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0b30923a-2961-43d9-a820-082fbcb0b34a_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f));
