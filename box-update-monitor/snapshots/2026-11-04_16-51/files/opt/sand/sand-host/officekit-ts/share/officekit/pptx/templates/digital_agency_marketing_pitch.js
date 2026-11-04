/**
 * "DIGITAL AGENCY" — 20-slide deck rebuilt with pptxgenjs.
 *
 * The source deck contains no raster media at all: every photo slot is an
 * *empty* picture placeholder, so `pictureFrame()` reproduces each one as a
 * plain rectangle at the same position/size. All other artwork (arrows, check
 * marks, pictograms, the hand-drawn line chart, the gantt grid) is recreated
 * with native pptxgenjs shapes, and slide 15 uses a real pptxgenjs bar chart.
 *
 * Run:  node 0518e166-bb4f-4675-a097-ad93b3c71dba_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  bg: '140731',          // dk2 — deck background
  purpleDeep: '3F0E5B',  // lt2
  violet: '7158E6',      // accent1
  magenta: 'C065FF',     // accent2
  white: 'FFFFFF',
  black: '000000',
  magentaDark: '9C0CFF',    // accent2 lumMod 75%
  magentaPale: 'D9A3FF',    // accent2 lumMod 60% / lumOff 40%
  magentaLight: 'E6C1FF',   // accent2 lumMod 40% / lumOff 60%
  magentaMist: 'F2E0FF',    // accent2 lumMod 20% / lumOff 80%
  violetPale: 'AA9BF0',     // accent1 lumMod 60% / lumOff 40%
  violetMist: 'E3DEFA',     // accent1 lumMod 20% / lumOff 80%
  greyLine: 'D9D9D9',       // bg1 lumMod 85%
  cardEdge: '312664',       // accent1 shade 15% — the theme outline on filled cards
};

const FONT_MAJOR = 'SUSE';           // theme major latin face
const FONT_MINOR = 'Noto Sans';      // theme minor latin face
const FONT_LIGHT = 'Space Grotesk Light';

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

/* -------------------------------------------------- shared body copy */

const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.\u00a0';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim.';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit';
const LOREM_SED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed.';
const UT_ENIM = 'Ut enim ad minim veniam, quis nostrud exercitation ullamco.';

/* ------------------------------------------------- reusable text presets */

// Big headline: theme major font, white, 54pt, single-spaced.
const heading = (over) => Object.assign({
  fontFace: FONT_MAJOR, fontSize: 54, color: C.white,
  valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], fit: 'resize',
}, over);

// Headline variant used on slides 3/4/5/7/17 — 115% leading, 8pt after.
const headingLoose = (over) => heading(Object.assign({
  lineSpacingMultiple: 1.15, paraSpaceAfter: 8,
}, over));

// 14pt body copy, 150% leading — the deck's default paragraph.
const body = (over) => Object.assign({
  fontFace: FONT_MINOR, fontSize: 14, color: C.white,
  lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], fit: 'resize',
}, over);

// Plain textbox with no leading override (numbers, section labels, captions).
const label = (over) => Object.assign({
  fontFace: FONT_MINOR, color: C.white, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], fit: 'resize',
}, over);

/* ------------------------------------------------------- vector icon art */

// Each icon is a custGeom outline in its own EMU coordinate box; `icon()`
// scales the outline into an inch-sized rectangle on the slide.

// long "->" arrow used by the LEARN MORE links (drawn rotated 180deg)
const ICON_ARROW = { w: 152400, h: 152400, d: [['M',36433,85725], ['L',89773,139065], ['L',76200,152400],
  ['L',0,76200], ['L',76200,0], ['L',89773,13335], ['L',36433,66675], ['L',152400,66675], ['L',152400,85725],
  ['L',36433,85725], ['Z']] };

// "v" chevron bulleting the list on slide 6 (drawn rotated 270deg)
const ICON_CHEVRON = { w: 114300, h: 70485, d: [['M',57150,70485], ['L',0,13335], ['L',13335,0], ['L',57150,43815],
  ['L',100965,0], ['L',114300,13335], ['L',57150,70485], ['Z']] };

// circled check mark, slides 7 and 18
const ICON_CHECK = { w: 190500, h: 190500, d: [['M',81915,139065], ['L',149066,71914], ['L',135731,58579],
  ['L',81915,112395], ['L',54769,85249], ['L',41434,98584], ['L',81915,139065], ['Z'], ['M',95250,190500],
  ['C',82074,190500,69691,188000,58103,182999], ['C',46514,177998,36433,171212,27861,162639],
  ['C',19288,154067,12502,143986,7501,132398], ['C',2500,120809,0,108426,0,95250],
  ['C',0,82074,2500,69691,7501,58103], ['C',12502,46514,19288,36433,27861,27861],
  ['C',36433,19288,46514,12502,58103,7501], ['C',69691,2500,82074,0,95250,0], ['C',108426,0,120809,2500,132398,7501],
  ['C',143986,12502,154067,19288,162639,27861], ['C',171212,36433,177998,46514,182999,58103],
  ['C',188000,69691,190500,82074,190500,95250], ['C',190500,108426,188000,120809,182999,132398],
  ['C',177998,143986,171212,154067,162639,162639], ['C',154067,171212,143986,177998,132398,182999],
  ['C',120809,188000,108426,190500,95250,190500], ['Z'], ['M',95250,171450],
  ['C',116522,171450,134541,164068,149304,149304], ['C',164068,134541,171450,116523,171450,95250],
  ['C',171450,73978,164068,55959,149304,41196], ['C',134541,26432,116522,19050,95250,19050],
  ['C',73978,19050,55959,26432,41196,41196], ['C',26432,55959,19050,73978,19050,95250],
  ['C',19050,116523,26432,134541,41196,149304], ['C',55959,164068,73978,171450,95250,171450], ['Z']] };

// eight-point sparkle sitting on each gantt bar, slide 14
const ICON_SPARK = { w: 171450, h: 171450, d: [['M',76200,171450], ['L',76200,108585], ['L',31909,153114],
  ['L',18336,139541], ['L',62865,95250], ['L',0,95250], ['L',0,76200], ['L',62865,76200], ['L',18336,31909],
  ['L',31909,18336], ['L',76200,62865], ['L',76200,0], ['L',95250,0], ['L',95250,62865], ['L',139541,18336],
  ['L',153114,31909], ['L',108585,76200], ['L',171450,76200], ['L',171450,95250], ['L',108585,95250],
  ['L',153114,139541], ['L',139541,153114], ['L',95250,108585], ['L',95250,171450], ['L',76200,171450], ['Z']] };

// map pin, slide 16
const ICON_PIN = { w: 349036, h: 436293, d: [['M',174518,218147], ['C',186516,218147,196787,213875,205331,205331],
  ['C',213876,196786,218148,186515,218148,174517], ['C',218148,162519,213876,152248,205331,143704],
  ['C',196787,135160,186516,130888,174518,130888], ['C',162520,130888,152249,135160,143705,143704],
  ['C',135160,152248,130889,162519,130889,174517], ['C',130889,186515,135160,196786,143705,205331],
  ['C',152249,213875,162520,218147,174518,218147], ['Z'], ['M',174518,378484],
  ['C',218874,337764,251779,300770,273230,267502], ['C',294681,234235,305407,204694,305407,178880],
  ['C',305407,139250,292772,106801,267503,81532], ['C',242235,56264,211239,43629,174518,43629],
  ['C',137797,43629,106801,56264,81533,81532], ['C',56264,106801,43630,139250,43630,178880],
  ['C',43630,204694,54355,234235,75806,267502], ['C',97257,300770,130162,337764,174518,378484], ['Z'],
  ['M',174518,436293], ['C',115982,386483,72261,340218,43357,297498], ['C',14452,254777,0,215238,0,178880],
  ['C',0,124344,17543,80896,52628,48538], ['C',87713,16179,128343,0,174518,0],
  ['C',220693,0,261323,16179,296408,48538], ['C',331493,80896,349036,124344,349036,178880],
  ['C',349036,215238,334584,254777,305679,297498], ['C',276775,340218,233054,386483,174518,436293], ['Z']] };

// person pictogram (slide 11): a circular head plus a skirt / trousers torso
const ICON_HEAD = { w: 212234, h: 212234, d: [['M',106117,0], ['C',135295,0,160278,10389,181061,31172],
  ['C',201839,51956,212234,76934,212234,106117], ['C',212234,135301,201839,160279,181061,181062],
  ['C',160278,201845,135295,212234,106117,212234], ['C',76933,212234,51950,201845,31166,181062],
  ['C',10388,160279,0,135301,0,106117], ['C',0,76934,10388,51956,31166,31172],
  ['C',51950,10389,76933,0,106117,0], ['Z']] };
const BODY_SKIRT = { w: 477527, h: 795878, d: [['M',238767,0], ['C',262642,0,284082,6189,303100,18572],
  ['C',322112,30949,335152,48636,342227,71630], ['L',477527,477527], ['L',318352,477527], ['L',318352,795878],
  ['L',159176,795878], ['L',159176,477527], ['L',0,477527], ['L',135301,71630],
  ['C',142376,48636,155416,30949,174428,18572], ['C',193445,6189,214886,0,238767,0], ['Z']] };
const BODY_TROUSERS = { w: 371410, h: 795878, d: [['M',106117,0], ['L',265293,0],
  ['C',294476,0,319454,10389,340238,31172], ['C',361021,51955,371410,76933,371410,106117], ['L',371410,397942],
  ['L',291825,397942], ['L',291825,795878], ['L',79591,795878], ['L',79591,397942], ['L',0,397942],
  ['L',0,106117], ['C',0,76933,10389,51955,31172,31172], ['C',51956,10389,76934,0,106117,0], ['Z']] };

/* -------------------------------------------------------------- helpers */

/** Draw a vector icon scaled into the box (x, y, w, h), in inches. */
function icon(slide, art, x, y, w, h, opts = {}) {
  const sx = w / art.w;
  const sy = h / art.h;
  const points = art.d.map((seg) => {
    if (seg[0] === 'Z') return { close: true };
    if (seg[0] === 'M') return { x: seg[1] * sx, y: seg[2] * sy, moveTo: true };
    if (seg[0] === 'L') return { x: seg[1] * sx, y: seg[2] * sy };
    return {
      x: seg[5] * sx, y: seg[6] * sy,
      curve: { type: 'cubic', x1: seg[1] * sx, y1: seg[2] * sy, x2: seg[3] * sx, y2: seg[4] * sy },
    };
  });
  slide.addShape('custGeom', {
    x, y, w, h, points, rotate: opts.rotate || 0, line: { type: 'none' },
    fill: { color: opts.color || C.white, transparency: opts.transparency || 0 },
  });
}

/** Empty picture placeholder from the source deck, reproduced 1:1 in size. */
function pictureFrame(slide, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h, fill: { color: C.bg }, line: { type: 'none' } });
}

// Filled cards inherit the theme's 1pt outline; band/panel rectangles override it to none.
const cardOutline = { color: C.cardEdge, width: 1 };

/** Straight rule; pptxgenjs draws `line` shapes from (x, y) with size (w, h). */
function rule(slide, x, y, w, h, color, pt) {
  slide.addShape('line', { x, y, w, h, line: { color, width: pt } });
}

/** "-> LEARN MORE" call to action; the label wraps when the box is narrow. */
function learnMore(slide, x, y, w, arrowColor) {
  icon(slide, ICON_ARROW, x, y + 0.166, 0.185, 0.185, { rotate: 180, color: arrowColor || C.white });
  slide.addText('LEARN MORE', body({ x: x + 0.185, y, w, h: 0.454 }));
}

/** Stat block: optional eyebrow, a big number, then a caption (slides 2, 17). */
function statBlock(slide, o) {
  if (o.eyebrow) slide.addText(o.eyebrow, body({ x: o.x, y: o.y, w: o.w, h: 0.454 }));
  slide.addText(o.value, label({ x: o.x, y: o.valueY, w: o.w, h: 0.909, fontSize: 48, fontFace: o.valueFont || FONT_MINOR }));
  slide.addText(o.caption, body({ x: o.x, y: o.captionY, w: o.w, h: 0.808 }));
}

/** Master furniture: "DIGITAL AGENCY" top-right and "PAGE n / 20" bottom-left. */
function chrome(slide, pageNo) {
  slide.addText('DIGITAL AGENCY', label({
    x: 10.755, y: 0.431, w: 1.896, h: 0.303, align: 'right', fontSize: 12, fontFace: FONT_LIGHT,
  }));
  slide.addText(` PAGE ${pageNo} / 20`, label({
    x: 0.594, y: 6.799, w: 2.168, h: 0.303, fontSize: 12, fontFace: FONT_LIGHT,
  }));
}

/* ------------------------------------------------------- slide builders */

// 1 — Cover: two oversized words with a caption sandwiched between them.
function slide01(s) {
  pictureFrame(s, 0.682, 1.063, 3.891, 2.038);
  pictureFrame(s, 8.75, 4.484, 3.901, 2.038);
  s.addText('DIGITAL', label({ x: 4.905, y: 1.046, w: 6.422, h: 2.036, fontSize: 115, fontFace: FONT_MAJOR, wrap: false }));
  s.addText('PRESENTATION TEMPLATE', label({ x: 3.356, y: 3.469, w: 6.793, h: 0.505, fontSize: 24, fontFace: FONT_LIGHT, align: 'center' }));
  s.addText('AGENCY', label({ x: 1.546, y: 4.399, w: 6.506, h: 2.036, fontSize: 115, fontFace: FONT_MAJOR, align: 'right', wrap: false }));
}

// 2 — Section header with a headline stat and a LEARN MORE link.
function slide02(s) {
  pictureFrame(s, 6.667, 2.99, 6.667, 3.753);
  s.addText('YOUR DIGITAL GROWTH PARTNER', heading({ x: 0.76, y: 0.878, w: 11.812, h: 1.919, align: 'center' }));
  s.addText(LOREM_LONG, body({ x: 0.594, y: 2.822, w: 5.977, h: 1.515 }));
  statBlock(s, { x: 1.594, y: 4.55, w: 3.083, value: '32,4K', valueY: 4.55, captionY: 5.402, caption: LOREM_TINY });
  learnMore(s, 5.299, 5.755, 1.374);
}

// 3 — Quadrant layout: two "CREATIVE" blurbs left, big headline right.
function slide03(s) {
  rule(s, 5.75, 0.04, 0, 7.46, C.white, 1.5);
  rule(s, 0, 3.75, 5.75, 0, C.white, 1.5);
  [['CREATIVE 01', 1.226], ['CREATIVE 02', 4.585]].forEach(([title, y]) => {
    s.addText(title, label({ x: 1.168, y, w: 4.457, h: 0.707, fontSize: 36 }));
    s.addText(LOREM_MED, body({ x: 1.168, y: y + 0.599, w: 4.457, h: 1.161 }));
  });
  s.addText('CREATIVE SOLUTIONS FOR THE DIGITAL AGE', headingLoose({ x: 6.3, y: 1.067, w: 7.25, h: 3.177 }));
  s.addText(LOREM_LONG, body({ x: 6.494, y: 4.5, w: 5.79, h: 1.515 }));
}

// 4 — Two photos + two objectives, right-aligned headline with a big arrow.
function slide04(s) {
  pictureFrame(s, 0.909, 1.233, 2.954, 3.317);
  pictureFrame(s, 4.479, 1.233, 2.954, 3.317);
  icon(s, ICON_ARROW, 6.274, 3.404, 0.785, 0.785, { rotate: 180 });
  s.addText('DIGITAL MARKETING THAT DELIVERS', headingLoose({ x: 7.095, y: 1.233, w: 5.556, h: 2.835, fontSize: 48, align: 'right' }));
  s.addText(LOREM_MED.trim() + '.\u00a0', body({ x: 8.758, y: 4.233, w: 3.893, h: 1.515, align: 'right' }));
  [['OBJECTIVES ONE', 0.76, 4.804], ['OBJECTIVES TWO', 4.313, 4.821]].forEach(([title, x, y]) => {
    s.addText(title, label({ x, y, w: 3.918, h: 0.438, fontSize: 20, fontFace: FONT_MAJOR }));
    s.addText('Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt',
      body({ x, y: 5.21, w: 3.392, h: 1.161 }));
  });
}

// 5 — Full-bleed photo top-left, headline right, 89% stat along the bottom.
function slide05(s) {
  pictureFrame(s, 0, 0, 6.667, 4.267);
  s.addText('PERFORMANCE-DRIVEN DIGITAL CAMPAIGNS', headingLoose({ x: 6.915, y: 1.174, w: 6.75, h: 3.177 }));
  s.addText('89%', label({ x: -0.164, y: 4.62, w: 2.225, h: 0.909, fontSize: 48, align: 'right' }));
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim.',
    body({ x: 2.526, y: 4.682, w: 3.99, h: 1.515 }));
  s.addText('Lorem. ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt',
    body({ x: 7.04, y: 4.482, w: 5.2, h: 1.515 }));
}

// 6 — Three photo panels left, chevron list right.
function slide06(s) {
  pictureFrame(s, 0, 0.757, 3.678, 5.986);
  pictureFrame(s, 4.131, 2.555, 2.802, 1.942);
  pictureFrame(s, 4.131, 4.801, 4.319, 1.942);
  s.addText('DIGITAL STRATEGY PROPOSAL', heading({ x: 4.365, y: 0.933, w: 8.286, h: 1.919, align: 'right' }));
  s.addText(LOREM_SHORT, body({ x: 7.387, y: 2.909, w: 5.264, h: 1.161, align: 'right' }));
  ['ACADEMIC', 'SOCIETAL', 'INDUSTRY'].forEach((txt, i) => {
    const y = 4.358 + i * 0.8605;
    icon(s, ICON_CHEVRON, 9.012, y + 0.213, 0.238, 0.147, { rotate: 270 });
    s.addText(txt, label({ x: 9.409, y, w: 2.447, h: 0.572, fontSize: 28 }));
  });
}

// 7 — Two check-list columns split by a vertical rule.
function slide07(s) {
  s.addText('DIGITAL AGENCY OVERVIEW', headingLoose({ x: 2.862, y: 0.785, w: 7.608, h: 2.132, align: 'center' }));
  rule(s, 6.667, 2.46, 0, 4.543, C.black, 1.5);
  const columns = [
    { title: 'STRATEGY', titleX: 1.018, titleY: 3.166, iconX: 1.112,
      rows: [[3.883, 4.074, 1.56, 4.04], [4.835, 5.027, 1.56, 4.04], [5.844, 5.979, 1.56, 4.04]] },
    { title: 'CREATIVITY', titleX: 7.009, titleY: 3.198, iconX: 7.145,
      rows: [[3.915, 4.084, 7.711, 4.063], [4.892, 5.036, 7.591, 3.959], [5.844, 5.989, 7.591, 3.959]] },
  ];
  columns.forEach((col) => {
    s.addText(col.title, label({ x: col.titleX, y: col.titleY, w: 5.124, h: 0.572, fontSize: 28 }));
    col.rows.forEach(([textY, iconY, textX, textW]) => {
      icon(s, ICON_CHECK, col.iconX, iconY, 0.269, 0.269);
      s.addText(LOREM_SED, body({ x: textX, y: textY, w: textW, h: 0.808 }));
    });
  });
}

// 8 — Timeline: a vertical rule with one year + blurb per row.
function slide08(s) {
  rule(s, 1.06, 0.757, 0, 5.993, C.white, 1.5);
  rule(s, 0, 6.699, 13.333, 0, C.black, 1.5);
  s.addText('TIMELINE INFOGRAPHIC', heading({ x: 6.113, y: 1.314, w: 5.906, h: 1.919 }));
  ['2022', '2023', '2024', '2025'].forEach((year, i) => {
    const y = 1.165 + i * 1.426;
    s.addText(year, label({ x: 1.314, y, w: 1.686, h: 0.774, fontSize: 40, fontFace: FONT_MAJOR }));
    s.addText('Lorem ipsum dolor sit amet, adipiscing elit sed.', body({ x: 3.0, y, w: 2.573, h: 0.808 }));
  });
  pictureFrame(s, 6.231, 3.75, 7.103, 3.75);
}

// 9 — Hand-drawn line chart on a violet card (ovals + connectors, not a chart).
function slide09(s) {
  s.addShape('rect', { x: 0.682, y: 0.909, w: 5.984, h: 5.62, fill: { color: C.violet }, line: { type: 'none' } });
  s.addText('Income Quarter 1-2', label({ x: 1.083, y: 1.255, w: 4.837, h: 0.404, fontSize: 18, fontFace: FONT_MAJOR }));
  for (let i = 0; i < 6; i++) rule(s, 1.185, 2.084 + i * 0.6556, 4.708, 0, C.greyLine, 1);
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].forEach((m, i) => {
    s.addText(m, label({ x: 1.003 + i * 0.6905, y: 5.553, w: 0.852, h: 0.337, fontSize: 14, align: 'center' }));
  });
  // polyline joining the markers: [x, y, w, h, flipH]
  [[1.369, 3.485, 0.607, 1.158, true], [2.128, 3.485, 0.538, 0.775, false],
   [2.819, 2.11, 0.505, 2.15, true], [3.54, 2.11, 0.508, 1.248, false],
   [4.232, 2.835, 0.433, 0.587, true]].forEach(([x, y, w, h, flipH]) => {
    s.addShape('line', { x, y, w, h, flipH, line: { color: C.magentaMist, width: 2.25 } });
  });
  [[1.185, 4.617], [1.945, 3.331], [2.635, 4.233], [3.324, 2.019], [4.016, 3.332], [4.633, 2.681]]
    .forEach(([x, y]) => s.addShape('ellipse', {
      x, y, w: 0.215, h: 0.181, fill: { color: C.white }, line: { color: C.magentaLight, width: 2.25 },
    }));
  s.addText('TURNING TRAFFIC INTO LOYALTY', heading({ x: 7.143, y: 0.787, w: 5.477, h: 2.827 }));
  s.addText('2025', label({ x: 7.169, y: 4.051, w: 5.477, h: 0.909, fontSize: 48, fontFace: FONT_MAJOR }));
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    body({ x: 7.169, y: 4.96, w: 5.477, h: 1.161 }));
}

// 10 — Three tinted cards: eyebrow / title / copy / LEARN MORE.
function slide10(s) {
  s.addText('WHY CHOOSE US AS YOUR DIGITAL AGENCY', heading({ x: 2.138, y: 0.797, w: 9.057, h: 1.919, align: 'center' }));
  const cards = [
    { cardX: 0.562, cardY: 3.07,  fill: C.magentaDark, x: 0.959, y: 3.403, eyebrow: 'FIRST',  title: 'QUALITATIVE', linkW: 2.185 },
    { cardX: 4.723, cardY: 3.07,  fill: C.magenta,     x: 5.134, y: 3.433, eyebrow: 'SECOND', title: 'QUANTTATIVE', linkW: 2.003 },
    { cardX: 8.881, cardY: 3.064, fill: C.magentaPale, x: 9.235, y: 3.418, eyebrow: 'THIRD',  title: 'MIXED',       linkW: 2.323 },
  ];
  cards.forEach((c) => {
    s.addShape('rect', { x: c.cardX, y: c.cardY, w: 3.89, h: 3.633, fill: { color: c.fill }, line: cardOutline });
    s.addText(c.eyebrow, body({ x: c.x, y: c.y, w: 3.065, h: 0.454 }));
    s.addText(c.title, label({ x: c.x, y: c.y + 0.455, w: 3.065, h: 0.505, fontSize: 24, fontFace: FONT_MAJOR }));
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
      body({ x: c.x, y: c.y + 1.167, w: 3.065, h: 1.161 }));
    learnMore(s, c.x + 0.006, c.y + 2.466, c.linkW);
  });
}

// 11 — Pictogram infographic: ten figures per row, 60% / 40% split.
function slide11(s) {
  s.addText('OUR USER DIGITAL', heading({ x: 1.445, y: 1.16, w: 10.562, h: 1.01, align: 'center' }));
  s.addText(LOREM_SHORT, body({ x: 2.143, y: 2.374, w: 9.051, h: 0.808, align: 'center' }));
  s.addShape('rect', { x: 0.682, y: 3.535, w: 8.07, h: 2.953, fill: { color: C.magenta }, line: cardOutline });
  const rows = [
    { x0: 1.289, y: 3.9,  litUpTo: 7, torso: BODY_SKIRT,    headOff: 0.096, torsoW: 0.345 },
    { x0: 1.311, y: 5.15, litUpTo: 4, torso: BODY_TROUSERS, headOff: 0.058, torsoW: 0.268 },
  ];
  rows.forEach((row) => {
    for (let i = 0; i < 10; i++) {
      const x = row.x0 + i * 0.4893;
      // figures past the split are solid dark; the rest are translucent white
      const style = i >= row.litUpTo ? { color: C.purpleDeep } : { color: C.white, transparency: 69 };
      icon(s, ICON_HEAD, x + row.headOff, row.y, 0.153, 0.153, style);
      icon(s, row.torso, x, row.y + 0.191, row.torsoW, 0.575, style);
    }
  });
  s.addText('60%', label({ x: 6.799, y: 3.806, w: 1.87, h: 1.01, fontSize: 54, fontFace: FONT_MAJOR }));
  s.addText('40%', label({ x: 6.799, y: 5.038, w: 1.87, h: 1.01, fontSize: 54, fontFace: FONT_MAJOR }));
  s.addText(LOREM_LONG, body({ x: 9.058, y: 3.573, w: 3.501, h: 2.928 }));
}

// 12 — Photo strip left, 2x2 grid of coloured service cards right.
function slide12(s) {
  pictureFrame(s, 0, 0, 3.354, 6.743);
  s.addText('FULL-SERVICE DIGITAL AGENCY', heading({ x: 3.605, y: 0.75, w: 8.319, h: 1.919 }));
  const cards = [
    { title: 'SURVEYS',     cardX: 3.63,  cardY: 2.753, fill: C.magenta, x: 4.011, y: 2.895, w: 4.344 },
    { title: 'INTERVIEWS',  cardX: 8.355, cardY: 2.771, fill: C.violet,  x: 8.736, y: 2.909, w: 3.859 },
    { title: 'EXPERIMENTS', cardX: 3.63,  cardY: 4.821, fill: C.violet,  x: 4.011, y: 5.038, w: 4.344 },
    { title: 'ETC',         cardX: 8.355, cardY: 4.839, fill: C.magenta, x: 8.736, y: 5.002, w: 4.344 },
  ];
  cards.forEach((c) => {
    s.addShape('rect', { x: c.cardX, y: c.cardY, w: 4.344, h: 1.863, fill: { color: c.fill }, line: cardOutline });
    s.addText(c.title, label({ x: c.x, y: c.y, w: c.w, h: 0.64, fontSize: 32, fontFace: FONT_MAJOR }));
    s.addText('Lorem ipsum dolor sit amet, consetur adipiscing elit, sed do eiusmod tempor',
      body({ x: c.x, y: c.y + 0.64, w: c.w, h: 0.808 }));
  });
}

// 13 — Magenta band listing tools and analytical methods.
function slide13(s) {
  s.addText('REDEFINING DIGITAL EXPERIENCES', heading({ x: 2.141, y: 0.815, w: 9.051, h: 1.919, align: 'center' }));
  s.addText(LOREM_SHORT, body({ x: 2.141, y: 2.753, w: 9.051, h: 0.808, align: 'center' }));
  s.addShape('rect', { x: 0, y: 3.75, w: 13.333, h: 2.993, fill: { color: C.magenta }, line: cardOutline });
  const rows = [
    { title: 'TOOLS/SOFTWARE', titleW: 4.397, y: 4.3, itemY: 4.3, items: ['01. SPSS', '02. NVivo', '03. EXCEL'] },
    { title: 'ANALYTICAL METHODS', titleW: 5.002, y: 5.374, itemY: 5.413, items: ['01. CODING', '02. STATISTICS', '03. ETC'] },
  ];
  rows.forEach((r) => {
    s.addText(r.title, label({ x: 0.76, y: r.y, w: r.titleW, h: 0.572, fontSize: 28, fontFace: FONT_MAJOR }));
    r.items.forEach((it, i) => s.addText(it, body({ x: 6.366 + i * 2.2705, y: r.itemY, w: 1.666, h: 0.454 })));
  });
}

// 14 — Gantt chart drawn as a grid of rules plus phase bars.
function slide14(s) {
  s.addText('OUR PROJECT TIMELINE', heading({ x: 0.75, y: 0.851, w: 11.812, h: 1.01, align: 'center' }));
  s.addShape('rect', { x: 0, y: 2.316, w: 13.333, h: 4.387, fill: { color: C.violet }, line: { type: 'none' } });
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Dec'].forEach((m, i) => {
    s.addText(m, label({ x: 3.2 + i * 0.8316, y: 2.544, w: 0.657, h: 0.37, fontSize: 16 }));
  });
  for (let i = 0; i < 11; i++) rule(s, 3.074 + i * 0.8337, 2.544, 0, 3.782, C.violetPale, 1);
  [3.019, 3.642, 4.317, 4.991, 5.675, 6.326].forEach((y) => rule(s, 1.16, y, 11.117, 0, C.violetPale, 1));
  const bars = [
    { project: 'Project 01', labelY: 3.192, barX: 3.287, barY: 3.185, barW: 3.626, phase: 'Phase One' },
    { project: 'Project 02', labelY: 3.81,  barX: 4.306, barY: 3.825, barW: 4.069, phase: 'Phase Two' },
    { project: 'Project 03', labelY: 4.557, barX: 5.789, barY: 4.517, barW: 4.223, phase: 'Phase Three' },
    { project: 'Project 04', labelY: 5.186, barX: 8.083, barY: 5.21,  barW: 2.494, phase: 'Phase Four' },
    { project: 'Project 05', labelY: 5.88,  barX: 9.437, barY: 5.892, barW: 2.494, phase: 'Phase Four' },
  ];
  bars.forEach((b) => {
    s.addText(b.project, label({ x: 1.16, y: b.labelY, w: 2.04, h: 0.37, fontSize: 16 }));
    s.addShape('rect', { x: b.barX, y: b.barY, w: b.barW, h: 0.281, fill: { color: C.violetPale }, line: { type: 'none' } });
    icon(s, ICON_SPARK, b.barX + 0.058, b.barY + 0.048, 0.184, 0.184, { color: C.violetMist });
    s.addText(b.phase, label({ x: b.barX + 0.24, y: b.barY - 0.001, w: 2.04, h: 0.286, fontSize: 11 }));
  });
}

// 15 — Native stacked bar chart on a violet panel, copy on the right.
function slide15(s) {
  s.addShape('rect', { x: 0.682, y: 0, w: 6.151, h: 7.5, fill: { color: C.violet }, line: { type: 'none' } });
  s.addText('REPORT INCOME IN 2025', label({ x: 1.111, y: 1.164, w: 4.583, h: 0.438, fontSize: 20, fontFace: FONT_MAJOR }));
  s.addChart('bar', [{
    name: 'Market B',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Des'],
    values: [1.3, 2.8, 2, 3.4, 3, 3.2, 3.6, 5.7, 7.9, 8.7, 6, 9],
  }], {
    x: 0.998, y: 1.961, w: 5.52, h: 3.06,
    // the source chart has gapWidth 0; pptxgenjs rewrites a falsy 0 to 50 on stacked bars
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 1,
    chartColors: [C.magenta], dataBorder: { pt: 0.75, color: C.violetMist },
    valAxisMaxVal: 10, showLegend: false, showTitle: false,
    catAxisLabelColor: C.white, catAxisLabelFontFace: FONT_MINOR, catAxisLabelFontSize: 12,
    valAxisLabelColor: C.white, valAxisLabelFontFace: FONT_MINOR, valAxisLabelFontSize: 12,
    catAxisLineColor: C.greyLine, valAxisLineShow: false,
    valGridLine: { color: C.violetPale, size: 1 }, catGridLine: { style: 'none' },
  });
  s.addText('FIRST QUARTER', body({ x: 1.111, y: 5.168, w: 1.911, h: 0.454, fontFace: FONT_MAJOR }));
  s.addText('87,88K', label({ x: 1.111, y: 5.636, w: 1.911, h: 0.64, fontSize: 32, fontFace: FONT_MAJOR }));
  s.addText('LAST QUARTER', body({ x: 4.454, y: 5.168, w: 1.911, h: 0.454, align: 'right', fontFace: FONT_MAJOR }));
  s.addText('253,65K', label({ x: 4.454, y: 5.636, w: 1.911, h: 0.64, fontSize: 32, fontFace: FONT_MAJOR, align: 'right' }));
  s.addText('BAR CHART SLIDE', heading({ x: 7.387, y: 1.282, w: 4.949, h: 1.919 }));
  s.addText([
    { text: LOREM_MED, options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: UT_ENIM },
  ], body({ x: 7.456, y: 3.557, w: 4.745, h: 2.221 }));
  learnMore(s, 7.552, 6.056, 2.091);
  // a duplicate CTA is parked below the slide edge in the source deck
  learnMore(s, 6.304, 7.689, 1.374, C.black);
}

// 16 — Contact slide: copy left, address with a pin icon, photo right.
function slide16(s) {
  s.addText('LET\u2019S GROW TOGETHER', heading({ x: 2.38, y: 1.175, w: 8.573, h: 1.01, align: 'center' }));
  s.addText(LOREM_MED, body({ x: 0.898, y: 2.849, w: 5.138, h: 1.161 }));
  icon(s, ICON_PIN, 1.218, 4.651, 0.364, 0.455);
  s.addText('ADDRESS', label({ x: 1.81, y: 4.651, w: 3.903, h: 0.572, fontSize: 28 }));
  s.addText('Consectetur Road, Cambridge G1578K, United Kingdom', body({ x: 1.81, y: 5.223, w: 3.906, h: 0.808 }));
  pictureFrame(s, 6.84, 2.883, 5.811, 3.442);
}

// 17 — Violet band with four quarter figures.
function slide17(s) {
  s.addText('CREATING DIGITAL IMPACT THAT MATTERS', headingLoose({ x: 2.638, y: 0.958, w: 10.07, h: 2.132, align: 'right' }));
  s.addShape('rect', { x: 0, y: 3.319, w: 13.333, h: 3.0, fill: { color: C.violet }, line: { type: 'none' } });
  [['QUARTER 01', '45,4K', 0.625], ['QUARTER 02', '660,8K', 3.742],
   ['QUARTER 03', '868,1K', 6.858], ['QUARTER 04', '998,1K', 9.975]].forEach(([eyebrow, value, x]) => {
    statBlock(s, { x, y: 3.762, w: 3.083, eyebrow, value, valueY: 4.217, valueFont: FONT_MAJOR,
      captionY: 5.069, caption: LOREM_TINY });
  });
}

// 18 — Two columns: prose left, checklist right.
function slide18(s) {
  s.addText('YOUR NEXT DIGITAL PARTNER', heading({ x: 1.955, y: 0.781, w: 9.423, h: 1.919, align: 'center' }));
  s.addText('FOR PRACTICE', label({ x: 0.941, y: 3.098, w: 3.267, h: 0.505, fontSize: 24, fontFace: FONT_MAJOR }));
  s.addText([
    { text: LOREM_MED, options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: UT_ENIM },
  ], body({ x: 0.941, y: 3.721, w: 4.676, h: 2.221 }));
  s.addText('FURTHER RESEARCH', label({ x: 6.083, y: 3.098, w: 5.906, h: 0.505, fontSize: 24, fontFace: FONT_MAJOR }));
  [3.652, 4.735, 5.795].forEach((y) => {
    icon(s, ICON_CHECK, 6.206, y + 0.148, 0.269, 0.269);
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
      body({ x: 6.567, y, w: 5.423, h: 0.808 }));
  });
}

// 19 — Photo left, headline plus two labelled rows right.
function slide19(s) {
  pictureFrame(s, 0, 0.757, 6.117, 6.743);
  s.addText('BOLD IDEAS, REAL RESULTS.', heading({ x: 6.426, y: 0.986, w: 6.225, h: 1.919 }));
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    body({ x: 6.492, y: 3.097, w: 5.974, h: 1.161 }));
  [['TYPE ONE', 4.576, 4.312], ['TYPE TWO', 5.806, 4.178]].forEach(([title, y, w]) => {
    s.addText(title, label({ x: 6.492, y: y + 0.07, w: 1.663, h: 0.438, fontSize: 20, fontFace: FONT_MAJOR }));
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do.', body({ x: 8.155, y, w, h: 0.808 }));
  });
}

// 20 — Closing slide, "THANK YOU" set at 138pt.
function slide20(s) {
  s.addText('THANK', label({ x: 0.882, y: 1.33, w: 9.084, h: 2.423, fontSize: 138, fontFace: FONT_MAJOR }));
  s.addText('FOR YOUR ATTENTION', label({ x: 0.882, y: 4.471, w: 5.334, h: 0.707, fontSize: 36, fontFace: FONT_LIGHT }));
  s.addText('YOU', label({ x: 5.996, y: 3.259, w: 4.132, h: 2.423, fontSize: 138, fontFace: FONT_MAJOR, align: 'right', wrap: false }));
}

/* ---------------------------------------------------------------- build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'WIDE';
  pptx.theme = { headFontFace: FONT_MAJOR, bodyFontFace: FONT_MINOR };

  BUILDERS.forEach((builder, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.bg };
    builder(slide);
    if (i > 0) chrome(slide, i + 1); // the cover hides the master furniture
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0518e166-bb4f-4675-a097-ad93b3c71dba_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
