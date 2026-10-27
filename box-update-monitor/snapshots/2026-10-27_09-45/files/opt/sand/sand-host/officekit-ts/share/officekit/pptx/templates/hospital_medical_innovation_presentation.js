/**
 * "Lifecare - Medical Presentation Template" — 20 slide deck rebuilt with pptxgenjs.
 * Slide size 20 x 11.25 in (16:9 @ 1920px-equivalent).
 *
 * Run: node 053270e4-dd3a-417b-b64c-1ff564821bf6_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const TEAL = '16BCA4';
const ORANGE = 'FAA239';
const INK = '222424';
const WHITE = 'FFFFFF';
const OFFWHITE = 'FAFAFA';

const F_HEAD = 'Inter Tight SemiBold';
const F_REG = 'Inter Tight';
const F_BODY = 'Inter Light';

const SLIDE_W = 20;
const SLIDE_H = 11.25;

/* ------------------------------------------------------- text style helpers */

/** Big teal section heading (44pt unless overridden). */
const title = (t, x, y, w, h, o = {}) => ({
  text: t,
  opts: Object.assign(
    { x, y, w, h, fontFace: F_HEAD, fontSize: 44, color: TEAL, valign: 'top' },
    o
  ),
});

/** 20pt semibold sub-heading / list line. */
const label = (t, x, y, w, o = {}) => ({
  text: t,
  opts: Object.assign(
    { x, y, w, h: 0.438, fontFace: F_HEAD, fontSize: 20, color: INK, valign: 'top' },
    o
  ),
});

/** 18pt light body copy, 150% line spacing. */
const body = (t, x, y, w, h, o = {}) => ({
  text: t,
  opts: Object.assign(
    {
      x, y, w, h,
      fontFace: F_BODY, fontSize: 18, color: INK,
      lineSpacingMultiple: 1.5, valign: 'top',
    },
    o
  ),
});

function addTexts(slide, items) {
  items.forEach((it) => slide.addText(it.text, it.opts));
}

/* ------------------------------------------------------------ shape helpers */

/**
 * Drop-shadow presets, in points. Returned fresh each call because pptxgenjs
 * rewrites the object it is handed (pt -> EMU) while rendering.
 */
const SHADOW_GREY = () => ({ type: 'outer', color: '808080', opacity: 0.4, blur: 32, offset: 36, angle: 135 });
const SHADOW_SOFT = () => ({ type: 'outer', color: 'D9D9D9', opacity: 0.3, blur: 50, offset: 55, angle: 135 });
const SHADOW_WIDE = () => ({ type: 'outer', color: 'BFBFBF', opacity: 0.3, blur: 60, offset: 35, angle: 135 });
const SHADOW_LEFT = () => ({ type: 'outer', color: 'D9D9D9', opacity: 0.3, blur: 42, offset: 50, angle: 45 });

/** White card with the deck's soft drop shadow. */
function card(slide, x, y, w, h, shadow) {
  slide.addShape('rect', { x, y, w, h, fill: { color: WHITE }, shadow: shadow() });
}

/** Rounded orange tile that hosts an icon (1.087in in the source deck). */
function iconTile(slide, x, y, s = 1.087) {
  slide.addShape('roundRect', { x, y, w: s, h: s, fill: { color: ORANGE }, rectRadius: s * 0.26889 });
}

/**
 * White pictogram standing in for the deck's small raster/SVG icons.
 * Drawn from native shapes so nothing has to be embedded.
 */
function icon(slide, kind, x, y, s = 0.591) {
  if (kind === 'medical') {
    slide.addShape('donut', {
      x: x + s * 0.105, y: y + s * 0.105, w: s * 0.79, h: s * 0.79,
      fill: { color: WHITE }, rectRadius: s * 0.0625,
    });
    slide.addShape('plus', {
      x: x + s * 0.25, y: y + s * 0.25, w: s * 0.5, h: s * 0.5,
      fill: { color: WHITE }, rectRadius: s * 0.15,
    });
  } else if (kind === 'database') {
    [0.08, 0.35, 0.62].forEach((dy) =>
      slide.addShape('ellipse', {
        x: x + s * 0.16, y: y + s * dy, w: s * 0.68, h: s * 0.22, fill: { color: WHITE },
      })
    );
  } else if (kind === 'care') {
    slide.addShape('heart', { x: x + s * 0.28, y: y + s * 0.08, w: s * 0.44, h: s * 0.38, fill: { color: WHITE } });
    slide.addShape('blockArc', {
      x: x + s * 0.02, y: y + s * 0.14, w: s * 0.96, h: s * 0.8,
      fill: { color: WHITE }, angleRange: [20, 160], arcThicknessRatio: 0.14,
    });
  } else if (kind === 'stethoscope') {
    slide.addShape('blockArc', {
      x: x + s * 0.06, y: y - s * 0.2, w: s * 0.7, h: s * 0.92,
      fill: { color: WHITE }, angleRange: [15, 165], arcThicknessRatio: 0.13,
    });
    slide.addShape('ellipse', { x: x + s * 0.62, y: y + s * 0.2, w: s * 0.28, h: s * 0.28, fill: { color: WHITE } });
    slide.addShape('ellipse', { x: x + s * 0.29, y: y + s * 0.64, w: s * 0.26, h: s * 0.26, fill: { color: WHITE } });
  }
}

/* -------------------------------------------------------------- custom paths
 * Paths are stored in a normalised 0..1 box and scaled to the shape size.
 */

function scalePath(pts, w, h) {
  return pts.map((p) => {
    if (p.close) return p;
    if (p.c) {
      return {
        x: p.x * w, y: p.y * h,
        curve: { type: 'cubic', x1: p.c[0] * w, y1: p.c[1] * h, x2: p.c[2] * w, y2: p.c[3] * h },
      };
    }
    return { x: p.x * w, y: p.y * h };
  });
}

function addPath(slide, pts, x, y, w, h, o = {}) {
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points: scalePath(pts, w, h) }, o));
}

/** Pill-with-a-folded-corner used for every "Explore Now" button. */
const CTA_PATH = [
  { x: 0.1202, y: 0 },
  { x: 0.8648, y: 0 },
  { x: 1, y: 0.3971, c: [0.9395, 0, 1, 0.1778] },
  { x: 0.8648, y: 0.7942, c: [1, 0.6164, 0.9395, 0.7942] },
  { x: 0.0964, y: 0.7942 },
  { x: 0.0818, y: 1 },
  { x: 0, y: 1 },
  { x: 0.0494, y: 0.3043 },
  { x: 0.0993, y: 0.3043 },
  { close: true },
];

/** Small numbered badge (same folded-corner motif) used for 01/02/03 markers. */
const TAG_PATH = [
  { x: 0.3013, y: 0 },
  { x: 0.6675, y: 0 },
  { x: 1, y: 0.38, c: [0.8511, 0, 1, 0.1702] },
  { x: 1, y: 0.4153 },
  { x: 0.6675, y: 0.7953, c: [1, 0.6251, 0.8511, 0.7953] },
  { x: 0.2485, y: 0.7953 },
  { x: 0.2115, y: 1 },
  { x: 0, y: 1 },
  { x: 0.144, y: 0.2047 },
  { x: 0.2642, y: 0.2047 },
  { close: true },
];

const CTA_W = 3.299;
const CTA_H = 1.123;
const TAG_W = 0.945;
const TAG_H = 0.827;

/** Orange "Explore Now" button: x/y is the top-left of the banner shape. */
function cta(slide, x, y) {
  addPath(slide, CTA_PATH, x, y, CTA_W, CTA_H, { fill: { color: ORANGE } });
  slide.addText('Explore Now', {
    x: x + 0.86, y: y + 0.227, w: 1.795, h: 0.438,
    fontFace: F_HEAD, fontSize: 20, color: WHITE,
    align: 'center', valign: 'top', wrap: false,
  });
}

/** Orange numbered badge; nx/nw position the digits inside it. */
function tag(slide, num, x, y, nx, nw) {
  addPath(slide, TAG_PATH, x, y, TAG_W, TAG_H, { fill: { color: ORANGE } });
  slide.addText(num, {
    x: nx, y: y + 0.138, w: nw, h: 0.404,
    fontFace: F_HEAD, fontSize: 18, color: WHITE,
    align: 'center', valign: 'top', wrap: false,
  });
}

/* ------------------------------------------------------- repeated copy text */

const LOREM_LONG =
  'PLACEHOLDER' +
  'etunus dolore magna aliqua. Nibh sit amet commodo nulla facilisi nullam. Mattis pellentesque id ' +
  'nibuan tortor tellus integer feugiat scelerisque. Nunc velua risuta commodo viverra maecenas accumsan. ';
const LOREM_SHORT =
  'PLACEHOLDER' +
  'dolore magna aliqua. ';
const LOREM_MED =
  'PLACEHOLDER' +
  'etunus dolore magna aliqua nibua amet.';
const LOREM_COMFORT =
  'PLACEHOLDER' +
  'etunus dolore magna aliqua. Nibuan situa amet commodo nulla facilisi nullam mattis tesque nibuan comfort.';
const LOREM_ROWS =
  'Lorem ipsum dolor conse dipiscing sulita sedun donas eiusmod tempor incidunt labore etunus dolore.';
const LOREM_CARD = 'Lorem ipsum dolor amet here conse adipiscing sulita sed donas';

const WRITE_ANY = 'Write Anything Here';
const FEATURES = [
  'Save Space and Speed Up Patient Service',
  'Patient Data is Easily Accessible',
  'Medical Devices Based on Augmented Reality',
];

/* ---------------------------------------------------------------- the slides */

/** Cover — full-bleed teal with the product name. */
function slide01(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: TEAL, transparency: 15 } });
  s.addText('Lifecare', {
    x: 4.937, y: 3.705, w: 9.392, h: 3.13,
    fontFace: F_HEAD, fontSize: 180, color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
  s.addText('Medical Presentation Template', {
    x: 6.535, y: 6.639, w: 6.93, h: 0.707,
    fontFace: F_REG, fontSize: 36, color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
  s.addShape('plus', { x: 13.959, y: 3.456, w: 1.104, h: 1.104, fill: { color: ORANGE }, rectRadius: 0.3885 });
}

/** Intro statement with CTA (art placeholders on the right are empty in the source). */
function slide02(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [
    title('Discover the Latest Innovations in Hospital Services', 2.088, 2.01, 5.74, 2.322),
    body(LOREM_LONG, 2.088, 4.763, 6.381, 2.774),
  ]);
  cta(s, 2.088, 8.117);
}

/** Two numbered talking points over the layout's pale grey chevron. */
function slide03(pptx) {
  const s = pptx.addSlide();
  addPath(s, [{ x: 0.9111, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { close: true }],
    9.952, 5.625, 2.548, 5.625, { fill: { color: OFFWHITE } });
  addPath(s, [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0.9151, y: 1 }, { close: true }],
    9.964, 0, 2.536, 5.625, { fill: { color: OFFWHITE } });

  addTexts(s, [title('Future of Health Services in Hospitals', 2.047, 2.157, 6.412, 1.582)]);
  [
    { n: '01', y: 4.445, nx: 2.313, nw: 0.458 },
    { n: '02', y: 7.135, nx: 2.288, nw: 0.509 },
  ].forEach((it) => {
    tag(s, it.n, 2.006, it.y, it.nx, it.nw);
    addTexts(s, [
      label(WRITE_ANY, 3.452, it.y - 0.003, 5.277),
      body(LOREM_SHORT, 3.452, it.y + 0.547, 5.277, 1.411),
    ]);
  });
}

/** Three floating feature cards on a teal band. */
function slide04(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 13.688, y: 0, w: 6.312, h: 11.25, fill: { color: TEAL } });

  [
    { x: 12.042, y: 6.717, icon: 'medical', text: FEATURES[2] },
    { x: 11.232, y: 4.535, icon: 'database', text: FEATURES[1] },
    { x: 12.042, y: 2.345, icon: 'care', text: FEATURES[0] },
  ].forEach((c) => {
    card(s, c.x, c.y, 6.729, 2.188, SHADOW_GREY);
    addTexts(s, [label(c.text, c.x + 0.795, c.y + 0.707, 3.956, { h: 0.774 })]);
    iconTile(s, c.x + 5.081, c.y + 0.551);
    icon(s, c.icon, c.x + 5.329, c.y + 0.799);
  });

  addTexts(s, [
    title('Role of Technology in Modern Hospitals', 1.776, 2.485, 6.037, 1.582),
    body(
      'PLACEHOLDER' +
      'labore etunus dolore magna aliqua. Nibuhan sit ametu commodo nulla facilisi nullam. Mattis ' +
      'pellente id nibuan tortor tellus integer feugiat sceleris ques velua risuta commodo viver rasuna.',
      1.776, 4.377, 6.037, 2.774
    ),
  ]);
  cta(s, 1.776, 7.642);
}

/** Single highlighted card, right side. */
function slide05(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [
    title('How Hospitals Are Redefining Access to Care Here', 1.922, 2.762, 5.495, 2.322),
    label('01. ' + WRITE_ANY, 1.922, 5.584, 5.277),
    body(
      'PLACEHOLDER' +
      'etunus dolore magna aliqua. Nibua amet codonus nulla facilisi nullam. Mattis pellen tesque ' +
      'nibuan tortor tellus integer feugiat lerisque. ',
      1.922, 6.168, 6.12, 2.32
    ),
  ]);

  card(s, 10.521, 7.238, 6.729, 2.188, SHADOW_SOFT);
  addTexts(s, [label(FEATURES[2], 11.316, 7.945, 3.956, { h: 0.774 })]);
  iconTile(s, 15.602, 7.788);
  icon(s, 'medical', 15.85, 8.037);
  s.addShape('round1Rect', {
    x: 14.874, y: 6.955, w: 2.376, h: 0.283, fill: { color: ORANGE }, rectRadius: 0.1415,
  });
}

/** Statistic block plus a numbered checklist. */
function slide06(pptx) {
  const s = pptx.addSlide();
  s.addShape('round1Rect', {
    x: 3.854, y: 6.958, w: 4.312, h: 4.292, fill: { color: TEAL }, rectRadius: 1.5834, flipH: true,
  });
  s.addShape('round1Rect', {
    x: 8.167, y: 5.611, w: 1.354, h: 1.348, fill: { color: ORANGE }, rectRadius: 0.4973,
  });
  icon(s, 'medical', 8.548, 5.989);
  s.addText('50K+', {
    x: 4.807, y: 8.042, w: 2.407, h: 0.774,
    fontFace: F_HEAD, fontSize: 40, color: WHITE, valign: 'top',
  });
  s.addText('Patients are Satisfied with the Service', {
    x: 4.807, y: 8.853, w: 2.407, h: 1.313,
    fontFace: F_HEAD, fontSize: 24, color: WHITE, valign: 'top',
  });

  addTexts(s, [
    title('Designing Hospitals for Comfort and Care', 11.234, 2.653, 6.849, 1.582),
    body(LOREM_COMFORT, 11.234, 4.618, 6.849, 1.865),
  ]);
  FEATURES.forEach((t, i) =>
    addTexts(s, [label(`${i + 1}. ${t}`, 11.234, 6.896 + i * 0.6315, 6.541)])
  );
}

/** Text-only statement with the layout's teal diagonal. */
function slide07(pptx) {
  const s = pptx.addSlide();
  addPath(s, [{ x: 0.382, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { close: true }],
    13.208, 0, 6.792, 11.25, { fill: { color: TEAL } });
  addTexts(s, [
    title('Hospital Responses to Public Health Emergencies', 2.088, 2.01, 5.766, 2.322),
    body(LOREM_LONG, 2.088, 4.763, 6.381, 2.774),
  ]);
  cta(s, 2.088, 8.117);
}

/** Case studies: heading bottom-left, two icon rows on the right. */
function slide08(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [
    title('Case Studies of Successful Hospital Campaigns', 2.067, 5.782, 6.037, 2.322),
    body(
      'PLACEHOLDER' +
      'etunus dolore magna aliqua sit amet commodo.',
      2.067, 8.39, 6.381, 1.411
    ),
  ]);
  [
    { y: 5.353, icon: 'care' },
    { y: 8.138, icon: 'medical' },
  ].forEach((r) => {
    iconTile(s, 10.878, r.y);
    icon(s, r.icon, 11.126, r.y + 0.248);
    addTexts(s, [
      label(WRITE_ANY, 12.656, r.y - 0.121, 5.277),
      body(LOREM_SHORT, 12.656, r.y + 0.429, 5.277, 1.411),
    ]);
  });
}

/** Two numbered notes, right column. */
function slide09(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [title('Culturally Competent Care in Hospitals', 11.107, 2.33, 6.393, 1.582)]);
  ['01', '02'].forEach((n, i) => {
    const y = 4.435 + i * 2.491;
    addTexts(s, [
      label(`${n}. ${WRITE_ANY}`, 11.107, y, 5.277),
      body(LOREM_MED, 11.107, y + 0.583, 6.12, 1.411),
    ]);
  });
}

/** Three icon rows beside a heading and CTA. */
function slide10(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [
    title('How Hospitals Are Addressing Environmental Health', 2.047, 1.994, 5.141, 3.063),
    body(
      'PLACEHOLDER' +
      'magna aliqua Nibu sit amet commodo nulla facilisi nullam mattis pellen tesque nibuan.',
      2.047, 5.312, 5.141, 2.32
    ),
  ]);
  cta(s, 2.047, 8.133);

  ['care', 'stethoscope', 'medical'].forEach((kind, i) => {
    const y = 1.84 + i * 2.925;
    iconTile(s, 8.545, y);
    icon(s, kind, 8.793, y + 0.248);
    addTexts(s, [
      label(WRITE_ANY, 10.156, y - 0.12, 4.615),
      body(LOREM_ROWS, 10.156, y + 0.43, 4.615, 1.411),
    ]);
  });
}

/** Team grid — four outlined portrait cards. */
function slide11(pptx) {
  const s = pptx.addSlide();
  s.addText('Meet Our Specialist', {
    x: 4.711, y: 1.468, w: 10.578, h: 0.909,
    fontFace: F_HEAD, fontSize: 48, color: TEAL, align: 'center', valign: 'top',
  });
  [1.558, 5.839, 10.12, 14.4].forEach((x) => {
    s.addShape('roundRect', {
      x, y: 3.625, w: 4.042, h: 5.521,
      fill: { color: OFFWHITE }, line: { color: TEAL, width: 1 }, rectRadius: 0.3977,
    });
    addTexts(s, [
      label('Write Name Here', x, 7.77, 4.042, { align: 'center' }),
      body('Job Position', x, 8.207, 4.042, 0.502, { align: 'center' }),
    ]);
  });
}

/** Doctor cards staggered over a teal diagonal. */
function slide12(pptx) {
  const s = pptx.addSlide();
  addPath(s, [{ x: 0, y: 1 }, { x: 0.3535, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { close: true }],
    13.458, 0, 6.542, 11.25, { fill: { color: TEAL } });

  [
    { x: 10.0, y: 1.198, tx: 10.714 },
    { x: 11.25, y: 4.214, tx: 11.964 },
    { x: 10.0, y: 7.229, tx: 10.714 },
  ].forEach((c) => {
    card(s, c.x, c.y, 6.406, 2.823, SHADOW_WIDE);
    addTexts(s, [
      label('Write Name Here', c.tx, c.y + 0.912, 2.911),
      body('Job Position', c.tx, c.y + 1.409, 2.911, 0.502),
    ]);
  });

  addTexts(s, [
    title('Certified and Experienced Medical Doctors are here', 2.13, 2.515, 6.037, 2.322),
    body(
      'PLACEHOLDER' +
      'etunus dolore magna aliqua. Nibh sit amet commodo nulla facilisi nullam mattis pellen tesque nibuan',
      2.13, 5.24, 6.381, 1.865
    ),
  ]);
  cta(s, 2.13, 7.612);
}

/** Staircase of three notes above a teal arc band. */
function slide13(pptx) {
  const s = pptx.addSlide();
  addPath(s, [
    { x: 0.6504, y: 0 },
    { x: 0.9814, y: 0.1858, c: [0.7703, 0, 0.883, 0.0673] },
    { x: 1, y: 0.2095 },
    { x: 1, y: 1 },
    { x: 0.9621, y: 1 },
    { x: 0.9556, y: 0.9817 },
    { x: 0.6504, y: 0.6627, c: [0.8831, 0.7869, 0.7733, 0.6627] },
    { x: 0.3453, y: 0.9817, c: [0.5276, 0.6627, 0.4178, 0.7869] },
    { x: 0.3388, y: 1 },
    { x: 0, y: 1 },
    { x: 0.0106, y: 0.9402 },
    { x: 0.6504, y: 0, c: [0.116, 0.3877, 0.3628, 0] },
    { close: true },
  ], 8.665, 6.137, 11.335, 5.113, { fill: { color: TEAL, transparency: 20 } });

  s.addText('Health Check Procedures in Hospitals', {
    x: 2.13, y: 1.943, w: 6.037, h: 2.524,
    fontFace: F_HEAD, fontSize: 48, color: TEAL, valign: 'top',
  });
  [
    { n: '01', x: 11.036, y: 1.938 },
    { n: '02', x: 7.407, y: 4.644 },
    { n: '03', x: 2.13, y: 7.35 },
  ].forEach((it) => {
    addTexts(s, [
      label(`${it.n}. ${WRITE_ANY}`, it.x, it.y, 5.277),
      body(LOREM_SHORT, it.x, it.y + 0.551, 5.277, 1.411),
    ]);
  });

  [{ x: 9.062, y: 3.182 }, { x: 5.373, y: 5.85 }].forEach((a) =>
    s.addShape('arc', {
      x: a.x, y: a.y, w: 2.73, h: 2.73, rotate: 280.4,
      line: { color: ORANGE, width: 0.5, beginArrowType: 'arrow' },
    })
  );
}

/** Five reasons laid out in two columns. */
function slide14(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [title('Why trust us with your health?', 2.224, 1.879, 6.037, 1.582)]);
  [
    { n: '01', x: 2.192, y: 4.633, nx: 2.5, nw: 0.458, tx: 3.639, ty: 4.629 },
    { n: '02', x: 2.192, y: 7.24, nx: 2.474, nw: 0.509, tx: 3.639, ty: 7.236 },
    { n: '03', x: 11.085, y: 2.021, nx: 11.363, nw: 0.516, tx: 12.531, ty: 2.017 },
    { n: '04', x: 11.085, y: 4.637, nx: 11.361, nw: 0.519, tx: 12.531, ty: 4.633 },
    { n: '05', x: 11.085, y: 7.234, nx: 11.365, nw: 0.512, tx: 12.531, ty: 7.231 },
  ].forEach((it) => {
    tag(s, it.n, it.x, it.y, it.nx, it.nw);
    addTexts(s, [
      label(WRITE_ANY, it.tx, it.ty, 5.277),
      body(LOREM_SHORT, it.tx, it.ty + 0.55, 5.277, 1.411),
    ]);
  });
}

/** Six service cards with teal headers, on an off-white background. */
function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: OFFWHITE };
  [6.667, 12.802].forEach((x) => {
    [1.098, 4.198, 7.297].forEach((y) => {
      s.addShape('rect', { x, y, w: 5.896, h: 2.854, fill: { color: WHITE } });
      s.addShape('rect', { x, y, w: 5.896, h: 1.114, fill: { color: TEAL } });
      icon(s, 'medical', x + 0.495, y + 0.262);
      addTexts(s, [
        label(WRITE_ANY, x + 1.447, y + 0.338, 3.289, { color: WHITE }),
        body(LOREM_CARD, x + 0.736, y + 1.506, 4.423, 0.957),
      ]);
    });
  });
  s.addText('Providing the best service for your health', {
    x: 1.536, y: 3.119, w: 3.762, h: 3.332,
    fontFace: F_HEAD, fontSize: 48, color: TEAL, valign: 'top',
  });
  cta(s, 1.536, 7.008);
}

/** Teal full-bleed with a white content panel. */
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  s.addShape('rect', { x: 7.958, y: 1.083, w: 10.188, h: 9.083, fill: { color: WHITE } });
  addTexts(s, [
    title('Specialized Care in Geriatric Hospitals', 9.627, 2.653, 6.849, 1.582),
    body(LOREM_COMFORT, 9.627, 4.618, 6.849, 1.865),
  ]);
  FEATURES.forEach((t, i) =>
    addTexts(s, [label(`${i + 1}. ${t}`, 9.627, 6.896 + i * 0.6315, 6.541)])
  );
}

/** Centred column of two numbered notes. */
function slide17(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [title('Maternity Care in Modern Hospitals', 6.94, 2.311, 6.037, 1.582)]);
  ['01', '02'].forEach((n, i) => {
    const y = 4.444 + i * 2.5;
    addTexts(s, [
      label(`${n}. ${WRITE_ANY}`, 6.94, y, 5.277),
      body(LOREM_MED, 6.94, y + 0.584, 6.12, 1.411),
    ]);
  });
}

/** Three feature cards with left-side shadows. */
function slide18(pptx) {
  const s = pptx.addSlide();
  addTexts(s, [
    title('Your Health is our Main Goal, come Check it Now', 1.963, 2.471, 5.766, 2.322),
    label('01. ' + WRITE_ANY, 1.963, 5.421, 5.277),
    body(
      'PLACEHOLDER' +
      'etunus dolore magna aliqua. Nibh sit amet commodo nulla facilisi nullam. Mattis pellen tesque ' +
      'nibuan tortorius tellus integer feugiat risque velua commodo.',
      1.963, 6.004, 5.766, 2.774
    ),
  ]);
  [
    { y: 1.908, icon: 'care', text: FEATURES[0] },
    { y: 4.531, icon: 'database', text: FEATURES[1] },
    { y: 7.155, icon: 'medical', text: FEATURES[2] },
  ].forEach((c) => {
    card(s, 8.542, c.y, 6.729, 2.188, SHADOW_LEFT);
    addTexts(s, [label(c.text, 9.337, c.y + 0.706, 3.956, { h: 0.774 })]);
    iconTile(s, 13.623, c.y + 0.55);
    icon(s, c.icon, 13.871, c.y + 0.798);
  });
}

/** Big teal quote panel plus a footnote row. */
function slide19(pptx) {
  const s = pptx.addSlide();
  s.background = { color: OFFWHITE };
  s.addShape('roundRect', {
    x: 1.75, y: 1.438, w: 16.5, h: 6.125, fill: { color: TEAL }, rectRadius: 0.7084,
  });
  addTexts(s, [
    title('Specialized Care in Geriatric Hospitals', 3.476, 2.643, 6.849, 1.582, { color: WHITE }),
    body(
      'PLACEHOLDER' +
      'etunus dolore magna aliqua. Nibuan sitau amet commodo nulla facilisi nullam. Mattis ' +
      'pellentesque id nibuan tortor.',
      3.476, 4.491, 6.849, 1.865, { color: WHITE }
    ),
  ]);
  s.addShape('roundRect', {
    x: 2.913, y: 8.142, w: 0.733, h: 0.733, fill: { color: ORANGE }, rectRadius: 0.1971,
  });
  icon(s, 'medical', 3.083, 8.308, 0.394);
  addTexts(s, [
    body(
      'PLACEHOLDER',
      4.172, 7.995, 7.099, 0.957
    ),
  ]);
}

/** Closing slide. */
function slide20(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: TEAL, transparency: 15 } });
  s.addText('Medical Presentation Template', {
    x: 6.905, y: 3.697, w: 6.19, h: 0.64,
    fontFace: F_REG, fontSize: 32, color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
  s.addText('Thanks for Watching', {
    x: 3.301, y: 4.337, w: 13.397, h: 1.784,
    fontFace: F_HEAD, fontSize: 100, color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
  addPath(s, CTA_PATH, 8.35, 6.43, CTA_W, CTA_H, { fill: { color: ORANGE } });
  s.addText('Lifecare', {
    x: 9.394, y: 6.605, w: 1.427, h: 0.505,
    fontFace: F_HEAD, fontSize: 24, color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
}

/* ---------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'LIFECARE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'LIFECARE';
  pptx.title = 'Lifecare - Medical Presentation Template';

  [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  ].forEach((fn) => fn(pptx));

  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '053270e4-dd3a-417b-b64c-1ff564821bf6_grok_final.pptx'),
});
