/*
 * Executive Business — 15-slide deck rebuilt with pptxgenjs.
 *
 * Run:  node 0473ed37-274a-4f1d-8b37-217b701f02b4_grok_final.js
 * Out:  0473ed37-274a-4f1d-8b37-217b701f02b4_grok_final.pptx  (next to this file)
 *
 * Slide size 13.333 x 7.5 in (16:9). Every slide inherits the "CHROME" master:
 * a light left rail with the rotated wordmark, the slide number and the year.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  blue: '072BF2', // primary brand blue (backgrounds, accents)
  blueMid: '1C46F2', // secondary blue (headline highlights, bars)
  blueSoft: '4B75F2', // tertiary blue (pills, rules)
  blueTint: 'B3BDF2', // pale blue (tinted backgrounds, muted headline)
  paper: 'EEEEEE', // page background
  white: 'FFFFFF',
  ink: '262626', // tx1 @85%  — headlines
  inkMid: '404040', // tx1 @75%  — labels
  inkSoft: '595959', // tx1 @65%  — small print
  grey: 'A6A6A6', // bg1 @65%  — body copy
  greyLt: 'BFBFBF', // bg1 @75%
  hair: 'D9D9D9', // bg1 @85%  — hairlines
  slate: '6C757D', // card titles
  deep: '092A38', // darkest skill bar
};

const HEAD = 'Space Grotesk Medium';
const HEAD_ALT = 'Space Grotesk';
const BODY = 'DM Sans';

const LOREM =
  'PLACEHOLDER' +
  'doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore ' +
  'veritatis et quasi architecto beatae vitae dicta sunt';
const LOREM_LONG = LOREM + ' explicabo. Nemo enim ipsam';
const LOREM_VOL = LOREM + ' explicabo. Nemo enim ipsam voluptatem';
const LOREM_ODIT = LOREM + ' explicabo. Nemo enim ipsam voluptatem quia voluptas odit';
const LOREM_SHORT = 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem ' +
  'accusantium doloremque laudantium, totam rem aperiam, ';
const LOREM_TINY = 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem ' +
  'accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab';

/* ------------------------------------------------------------------ helpers */

// Plain paragraph of copy. Reference text boxes are top-anchored, so make that
// the default (pptxgenjs otherwise centres text vertically).
function copy(slide, text, o) {
  slide.addText(text, Object.assign(
    { fontFace: BODY, fontSize: 9, color: C.grey, valign: 'top', isTextBox: true },
    o
  ));
}

// Multi-colour headline: pass [[text, color], ...] plus shared options.
function headline(slide, runs, o) {
  const size = o.fontSize || 40;
  slide.addText(
    runs.map(([text, color]) => ({ text, options: { color, fontFace: o.fontFace || HEAD, fontSize: size } })),
    Object.assign({ valign: 'top', isTextBox: true }, o)
  );
}

// roundRect adj values in the source are fractions of the shorter side.
const radius = (w, h, adj) => Math.min(w, h) * adj;

// Concentric-ring logo mark. Each ring is a pair of half-annuli whose centres
// sit 0.1 box-widths apart, so the two halves are offset across a horizontal
// seam; four slanted slivers ride on that seam. All fractions are of the mark
// box and come straight from the source artwork's freeform geometry.
const LOGO_HALVES = [
  // [x, y, w, h, startAngle, sweepEnd, thicknessRatio]
  [0.1009, 0.0000, 0.8984, 0.9480, 180, 360, 0.2805], // outer, upper
  [0.0008, 0.0520, 0.8984, 0.9480, 0, 180, 0.2805], // outer, lower
  [0.3016, 0.2105, 0.4968, 0.5242, 180, 360, 0.5120], // inner, upper
  [0.2016, 0.2652, 0.4968, 0.5242, 0, 180, 0.5120], // inner, lower
];
const LOGO_SLIVERS = [0.0, 0.2001, 0.5746, 0.7749]; // left edges, each 0.2251 wide

function logoMark(slide, x, y, w, h, ringColor, sliverColor, transparency) {
  const t = transparency || 0;
  LOGO_HALVES.forEach(([fx, fy, fw, fh, a0, a1, thick]) => {
    slide.addShape('blockArc', {
      x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h,
      angleRange: [a0, a1], arcThicknessRatio: thick,
      fill: { color: ringColor, transparency: t }, line: { type: 'none' },
    });
  });
  const sw = 0.2251 * w, sh = 0.0741 * h, slant = 0.4443 * sw;
  LOGO_SLIVERS.forEach((fx) => {
    slide.addShape('custGeom', {
      x: x + fx * w, y: y + 0.4629 * h, w: sw, h: sh,
      fill: { color: sliverColor, transparency: t }, line: { type: 'none' },
      points: [{ x: slant, y: 0 }, { x: sw, y: 0 }, { x: sw - slant, y: sh }, { x: 0, y: sh }, { close: true }],
    });
  });
}

// Rounded blue tile with the logo mark inside it (badgeW drives everything).
function brandBadge(slide, x, y, badgeW) {
  const badgeH = badgeW * 1.0111;
  slide.addShape('roundRect', {
    x, y, w: badgeW, h: badgeH, fill: { color: C.blue },
    rectRadius: radius(badgeW, badgeH, 0.33986),
    shadow: { type: 'outer', blur: 40, offset: 24, angle: 45, color: '000000', opacity: 0.25 },
  });
  logoMark(slide, x + 0.2270 * badgeW, y + 0.2442 * badgeH,
    0.5459 * badgeW, 0.5116 * badgeH, C.paper, C.blueMid);
}

// "[badge] Executive." lock-up used as the running header.
function brandLockup(slide, x, y, color) {
  brandBadge(slide, x, y, 0.27);
  slide.addText('Executive.', {
    x: x + 0.286, y: y - 0.032, w: 2.671, h: 0.337,
    fontFace: HEAD, fontSize: 14, color: color || C.ink, valign: 'top', isTextBox: true,
  });
}

// Fan of hairline arcs (the "Graphic 70" swoosh). Each arc is one cubic whose
// control points are a linear blend between the innermost and outermost curve.
const SWOOSH_IN = [[0.4006, 0.1111], [0.8498, 0.2015], [0.8348, 0.7609], [0.0, 1.0]];
const SWOOSH_OUT = [[0.9436, 0.0], [1.1702, 0.4503], [0.6859, 0.9297], [0.0, 1.0]];

function swoosh(slide, x, y, w, h, rotate, count) {
  for (let i = 0; i < count; i++) {
    const t = i / (count - 1);
    const p = SWOOSH_IN.map((a, k) => [
      (a[0] + (SWOOSH_OUT[k][0] - a[0]) * t) * w,
      (a[1] + (SWOOSH_OUT[k][1] - a[1]) * t) * h,
    ]);
    slide.addShape('custGeom', {
      x, y, w, h, rotate,
      fill: { type: 'none' },
      line: { color: C.white, width: 1.1, transparency: 62 },
      points: [
        { x: p[0][0], y: p[0][1] },
        { curve: { type: 'cubic', x1: p[1][0], y1: p[1][1], x2: p[2][0], y2: p[2][1] }, x: p[3][0], y: p[3][1] },
      ],
    });
  }
}

// Scalloped band used for the top / bottom edge of the pricing cards.
// Upper boundary follows a cosine wave; the band is half the box tall.
function scallopBand(slide, x, y, w, h, bumps, fill, invert) {
  const amp = h / 2;
  const steps = bumps * 2;
  const waveY = (i) => {
    const v = (amp / 2) * (1 - Math.cos((Math.PI * i) / 1));
    return invert ? amp - v : v;
  };
  const yAt = (i) => (invert ? amp : 0) + (invert ? -1 : 1) * (amp / 2) * (1 - Math.cos(Math.PI * i));
  const pts = [{ x: 0, y: yAt(0) }];
  for (let i = 1; i <= steps; i++) {
    const x0 = ((i - 1) / steps) * w, x1 = (i / steps) * w;
    pts.push({ curve: { type: 'cubic', x1: x0 + (x1 - x0) / 3, y1: yAt(i - 1), x2: x1 - (x1 - x0) / 3, y2: yAt(i) }, x: x1, y: yAt(i) });
  }
  pts.push({ x: w, y: yAt(steps) + amp });
  for (let i = steps - 1; i >= 0; i--) {
    const x0 = ((i + 1) / steps) * w, x1 = (i / steps) * w;
    pts.push({ curve: { type: 'cubic', x1: x0 - (x0 - x1) / 3, y1: yAt(i + 1) + amp, x2: x1 + (x0 - x1) / 3, y2: yAt(i) + amp }, x: x1, y: yAt(i) + amp });
  }
  pts.push({ close: true });
  slide.addShape('custGeom', { x, y, w, h, fill: { color: fill }, points: pts, line: { type: 'none' } });
  void waveY;
}

// Callout: rounded stat card with a triangular tail and a dot on the curve.
function statPin(slide, cardX, cardY, tailX, tailY, dotX, dotY, tailDown) {
  slide.addShape('triangle', {
    x: tailX, y: tailY, w: 1.272, h: 0.771,
    fill: { color: C.white, transparency: 80 }, rotate: tailDown ? 180 : 0,
  });
  slide.addShape('ellipse', {
    x: dotX, y: dotY, w: 0.214, h: 0.214,
    fill: { color: C.blueTint }, line: { color: C.white, width: 2.5 },
    shadow: { type: 'outer', blur: 15, offset: 3, angle: 90, color: '000000', opacity: 0.2 },
  });
  slide.addShape('roundRect', {
    x: cardX, y: cardY, w: 1.634, h: 1.628, fill: { color: C.paper },
    line: { color: C.white, width: 0.5 }, rectRadius: radius(1.634, 1.628, 0.13186),
    shadow: { type: 'outer', blur: 60, offset: 18, angle: 90, color: '000000', opacity: 0.1 },
  });
  slide.addText('72%', {
    x: cardX + 0.181, y: cardY + 0.328, w: 1.272, h: 0.649, align: 'center', valign: 'top',
    fontFace: HEAD_ALT, fontSize: 36, bold: true, color: C.blueMid,
    charSpacing: -1.5, lineSpacingMultiple: 0.9, isTextBox: true,
  });
  slide.addText('Tittle text here', {
    x: cardX + 0.181, y: cardY + 1.053, w: 1.272, h: 0.253, align: 'center', valign: 'top',
    fontFace: HEAD_ALT, fontSize: 10, color: C.inkSoft, lineSpacingMultiple: 0.9, isTextBox: true,
  });
}

// Small footer pair: "www.company.tld" on the left, a year on the right.
function footerPair(slide, x1, x2, y) {
  copy(slide, 'www.company.tld', { x: x1, y, w: 1.594, h: 0.303, fontSize: 12, color: C.inkSoft, wrap: false });
  copy(slide, '2029', { x: x2, y, w: 0.616, h: 0.303, fontSize: 12, color: C.inkSoft, wrap: false });
}

// Stand-in for a raster asset from the source deck.
function imagePlaceholder(slide, x, y, w, h, color) {
  slide.addShape('roundRect', {
    x, y, w, h, fill: { color: color || C.blueMid }, rectRadius: Math.min(w, h) * 0.16,
  });
}

/* ------------------------------------------------------------------- slides */

function slide01(pptx) { // Cover
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blue };
  swoosh(s, 0.343, 0.493, 6.148, 7.431, 285, 22);
  headline(s, [
    ['Executive\n', C.white],
    ['Business', C.blueTint],
    ['\u00ae', C.white],
  ], { x: 5.747, y: 1.204, w: 6.919, h: 2.794, fontSize: 80 });
  s.addShape('roundRect', {
    x: 5.963, y: 4.28, w: 1.673, h: 0.614, fill: { color: C.blueSoft },
    rectRadius: radius(1.673, 0.614, 0.18727),
    shadow: { type: 'outer', blur: 40, offset: 24, angle: 45, color: '000000', opacity: 0.25 },
  });
  copy(s, '2080', {
    x: 6.214, y: 4.264, w: 1.17, h: 0.56, fontSize: 20, color: C.blue,
    align: 'center', charSpacing: 3, lineSpacingMultiple: 1.5,
  });
  copy(s, 'Navigating Disruption How C-Suite Leaders Drive Transformation', {
    x: 5.901, y: 6.089, w: 4.134, h: 0.572, fontSize: 14, color: C.greyLt,
  });
}

function slide02(pptx) { // Section opener with blue panel
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.addShape('rect', { x: 9.946, y: 0, w: 3.387, h: 7.5, fill: { color: C.blue } });
  swoosh(s, 0.334, 0.216, 6.148, 7.431, 285, 22);
  [[0.672, 0.391], [0.756, 0.319], [0.839, 0.391]].forEach(([y, w]) => {
    s.addShape('line', { x: 1.321, y, w, h: 0, line: { color: C.grey, width: 1.5 } });
  });
  headline(s, [['Executive', C.ink], ['.', C.blue]], { x: 1.321, y: 3.026, w: 6.651, h: 1.447, fontSize: 80 });
  copy(s, 'Presented by.', { x: 1.321, y: 6.072, w: 1.008, h: 0.252, color: C.grey, wrap: false });
  copy(s, 'Company Name', { x: 1.321, y: 6.375, w: 1.352, h: 0.286, fontSize: 11, color: C.inkSoft, wrap: false });
  copy(s, 'Website.', { x: 4.441, y: 6.072, w: 0.714, h: 0.252, color: C.grey, wrap: false });
  copy(s, 'www.company.tld', { x: 4.441, y: 6.375, w: 1.476, h: 0.286, fontSize: 11, color: C.inkSoft, wrap: false });
}

function slide03(pptx) { // Title + oversized watermark
  const s = pptx.addSlide({ masterName: 'CHROME' });
  logoMark(s, 8.099, 2.601, 4.72, 4.473, C.blueTint, C.blueTint, 76);
  brandLockup(s, 1.383, 0.908);
  headline(s, [
    ['The Executive\u2019s ', C.ink],
    ['Playbook Strategies ', C.blueMid],
    ['for Leadership Excellence', C.ink],
  ], { x: 1.383, y: 3.637, w: 6.651, h: 2.121 });
  footerPair(s, 1.507, 5.526, 6.624);
}

function slide04(pptx) { // Pale blue statement page
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blueTint };
  brandLockup(s, 1.383, 0.908);
  headline(s, [['Boardroom to ', C.ink], ['Bottom Line', C.blue]], { x: 1.175, y: 2.658, w: 4.471, h: 1.447 });
  copy(s, LOREM_LONG, { x: 1.175, y: 4.417, w: 3.833, h: 0.989, color: C.inkSoft, lineSpacingMultiple: 1.5 });
  copy(s, LOREM, { x: 10.236, y: 4.912, w: 2.548, h: 1.443, color: C.inkSoft, lineSpacingMultiple: 1.5 });
}

function slide05(pptx) { // Blue page, big badge, two copy columns
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blueMid };
  swoosh(s, 10.464, -0.893, 2.908, 3.515, 105, 22);
  headline(s, [
    ['Psychology of Power Executive', C.blueTint],
    [' Leadership in Complex ', C.blueSoft],
    ['Organizations', C.blueTint],
  ], { x: 5.08, y: 1.97, w: 7.165, h: 2.121 });
  s.addShape('line', { x: 4.111, y: 1.97, w: 0, h: 4.03, line: { color: C.blueSoft, width: 1 } });
  [5.08, 8.863].forEach((x) => copy(s, LOREM_VOL, { x, y: 4.617, w: 3.395, h: 1.216, lineSpacingMultiple: 1.5 }));
  brandBadge(s, 1.453, 2.84, 1.746);
  copy(s, 'Executive.', {
    x: 1.719, y: 4.888, w: 1.214, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.paper, align: 'center',
  });
}

function slide06(pptx) { // Big number statistic
  const s = pptx.addSlide({ masterName: 'CHROME' });
  headline(s, [['670+', C.ink]], { x: 6.194, y: 1.246, w: 4.867, h: 2.036, fontSize: 115 });
  headline(s, [['Boardroom to ', C.ink], ['Bottom Line', C.blue]], { x: 6.159, y: 3.094, w: 4.471, h: 0.438, fontSize: 20 });
  [6.194, 9.598].forEach((x) => copy(s, LOREM_LONG, { x, y: 4.217, w: 2.743, h: 1.443, lineSpacingMultiple: 1.5 }));
  footerPair(s, 6.194, 10.213, 6.346);
}

function slide07(pptx) { // Vision & mission, numbered list
  const s = pptx.addSlide({ masterName: 'CHROME' });
  brandLockup(s, 1.009, 0.908);
  headline(s, [['Our ', C.ink], ['Vision and ', C.blueMid], ['Mission', C.ink]], { x: 1.009, y: 2.694, w: 3.951, h: 1.447 });
  copy(s, LOREM_ODIT, { x: 1.009, y: 4.482, w: 3.575, h: 1.216, lineSpacingMultiple: 1.5 });
  [['01', 2.258, 2.315, 2.729, 8.626, 0.628],
   ['02', 3.748, 3.805, 4.219, 8.595, 0.691],
   ['03', 5.238, 5.295, 5.709, 8.593, 0.695]].forEach(([num, numY, titleY, bodyY, numX, numW]) => {
    copy(s, num, {
      x: numX, y: numY, w: numW, h: 0.572, fontFace: HEAD, fontSize: 28,
      color: C.inkMid, align: 'center', wrap: false,
    });
    copy(s, 'Tittle text here', {
      x: 9.301, y: titleY, w: 2.065, h: 0.404, fontFace: HEAD, fontSize: 18, color: C.slate, wrap: false,
    });
    copy(s, LOREM_TINY, { x: 9.301, y: bodyY, w: 3.19, h: 0.761, lineSpacingMultiple: 1.5 });
  });
}

function slide08(pptx) { // Team grid
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blueTint };
  headline(s, [['Meet our ', C.blueMid], ['team', C.ink]], { x: 1.238, y: 1.653, w: 4.491, h: 0.774 });
  copy(s, LOREM.replace(' beatae vitae dicta sunt', ''), {
    x: 7.998, y: 1.659, w: 4.345, h: 0.762, color: C.inkMid, lineSpacingMultiple: 1.5,
  });
  [3.104, 6.155, 9.386].forEach((x) => {
    copy(s, 'Name Text here', {
      x, y: 6.025, w: 2.118, h: 0.404, fontFace: HEAD, fontSize: 18, color: C.inkMid, wrap: false,
    });
    copy(s, 'Position here', { x, y: 6.439, w: 1.774, h: 0.307, color: C.inkMid, lineSpacingMultiple: 1.5 });
  });
}

function slide09(pptx) { // Profile with skill bars
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.addShape('rect', { x: 0.521, y: 0, w: 3.688, h: 7.5, fill: { color: C.blueMid } });
  s.addText('Alexander Merwin', {
    x: 7.271, y: 1.577, w: 5.575, h: 0.83, fontFace: HEAD_ALT, fontSize: 36, bold: true,
    color: C.ink, valign: 'top', lineSpacingMultiple: 1.3, isTextBox: true,
  });
  s.addText('Position Name', {
    x: 7.271, y: 2.276, w: 3.847, h: 0.425, fontFace: BODY, fontSize: 16, color: C.grey,
    valign: 'bottom', lineSpacingMultiple: 1.3, isTextBox: true,
  });
  copy(s, LOREM_LONG, { x: 7.271, y: 2.817, w: 4.346, h: 0.989, lineSpacingMultiple: 1.5 });
  [[4.1, 90, C.blueMid, 3.123], [4.724, 85, C.blueSoft, 2.92], [5.347, 75, C.deep, 2.379]]
    .forEach(([rowY, score, barColor, barW]) => {
      s.addText('Skill description', {
        x: 7.28, y: rowY, w: 2.478, h: 0.345, fontFace: HEAD_ALT, fontSize: 12,
        color: C.ink, valign: 'bottom', isTextBox: true,
      });
      s.addText(String(score), {
        x: 10.149, y: rowY, w: 0.978, h: 0.345, fontFace: HEAD_ALT, fontSize: 12, bold: true,
        color: C.ink, align: 'right', valign: 'bottom', isTextBox: true,
      });
      s.addShape('line', {
        x: 7.4, y: rowY + 0.417, w: 3.597, h: 0,
        line: { color: '000000', width: 5, transparency: 87 },
      });
      s.addShape('line', {
        x: 7.4, y: rowY + 0.417, w: barW, h: 0,
        line: { color: barColor, width: 5, endArrowType: 'oval' },
      });
    });
}

function slide10(pptx) { // Service grid
  const s = pptx.addSlide({ masterName: 'CHROME' });
  brandLockup(s, 1.154, 1.51);
  headline(s, [['Our Service', C.ink]], { x: 1.009, y: 2.008, w: 3.951, h: 0.774 });
  copy(s, LOREM_ODIT, { x: 1.009, y: 4.469, w: 3.575, h: 1.216, lineSpacingMultiple: 1.5 });
  [['Service First Title', 6.12, 2.337, 1.901, 2.999, LOREM_SHORT],
   ['Service Second Title', 9.629, 2.337, 2.214, 2.999, LOREM_SHORT + 'e'],
   ['Service Third Title', 6.12, 4.888, 1.909, 2.881, LOREM_SHORT],
   ['Service Fourth Title', 9.629, 4.888, 2.122, 2.881, LOREM_SHORT]]
    .forEach(([title, iconX, iconY, titleW, bodyW, body]) => {
      imagePlaceholder(s, iconX, iconY, 0.461, 0.461, C.blueMid);
      const textX = iconX - 0.055;
      copy(s, title, {
        x: textX, y: iconY + 0.679, w: titleW, h: 0.337, fontFace: HEAD_ALT, fontSize: 14,
        bold: true, color: C.slate, wrap: false,
      });
      copy(s, body, { x: textX, y: iconY + 1.016, w: bodyW, h: 0.762, lineSpacingMultiple: 1.5 });
    });
}

function slide11(pptx) { // Revenue chart
  const s = pptx.addSlide({ masterName: 'CHROME' });
  brandLockup(s, 1.644, 1.51);
  headline(s, [['Estimated Annual Revenue', C.ink]], { x: 1.499, y: 2.008, w: 3.951, h: 2.121 });
  copy(s, LOREM_ODIT, { x: 1.499, y: 4.469, w: 3.575, h: 1.216, lineSpacingMultiple: 1.5 });
  const years = ['2017', '2018', '2019', '2020', '2021', '2022', '2023'];
  s.addChart([
    {
      type: pptx.ChartType.bar,
      data: [{ name: 'Series 3', labels: years, values: [13, 18, 38, 35, 31, 49, 33] }],
      options: { chartColors: [C.blueMid], barDir: 'col', barGapWidthPct: 150 },
    },
    {
      type: pptx.ChartType.line,
      data: [{ name: 'Series 2', labels: years, values: [47, 14, 33, 16, 28, 37, 49] }],
      options: {
        chartColors: [C.blueTint], lineSize: 2.75, lineSmooth: true,
        lineDataSymbol: 'circle', lineDataSymbolSize: 6,
        lineDataSymbolFillColor: 'ADB5BD', lineDataSymbolLineColor: '6C757D',
      },
    },
  ], {
    x: 5.701, y: 0.656, w: 6.714, h: 6.188,
    layout: { x: 0.0266, y: 0.1284, w: 0.9468, h: 0.6749 },
    showLegend: false, valAxisHidden: true,
    valGridLine: { color: 'E9ECEF', size: 0.5 },
    catAxisLineColor: C.hair, catAxisLabelColor: '172B4D',
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 10,
    chartArea: { fill: { color: C.paper } },
  });
  statPin(s, 7.334, 0.902, 7.515, 2.395, 8.044, 3.181, true);
}

function slide12(pptx) { // Blue page with area wave
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blue };

  // Silhouette of the area chart. The vertical fade is built by laying bands of
  // the page colour over the wave, which leaves the sky above it untouched.
  const wx = 0.513, wy = 1.54, ww = 12.82, wh = 5.96;
  const wave = [[0, 0.6520], [0.04008, 0.56210, 0.17237, 0.69478, 0.25896, 0.63460],
    [0.34553, 0.57443, 0.42594, 0.36780, 0.51945, 0.25615],
    [0.61296, 0.14450, 0.67524, 0.34238, 0.75533, 0.28366],
    [0.83543, 0.22493, 0.96846, -0.12520, 0.99846, 0.04735]];
  const pts = [{ x: wave[0][0] * ww, y: wave[0][1] * wh }];
  wave.slice(1).forEach(([x1, y1, x2, y2, x, y]) => pts.push({
    curve: { type: 'cubic', x1: x1 * ww, y1: y1 * wh, x2: x2 * ww, y2: y2 * wh }, x: x * ww, y: y * wh,
  }));
  pts.push({ x: ww, y: wh }, { x: 0, y: wh }, { close: true });
  s.addShape('custGeom', { x: wx, y: wy, w: ww, h: wh, fill: { color: '4A73F1' }, points: pts, line: { type: 'none' } });

  // Vertical fade: stacked translucent sheets of the page colour, each running
  // to the bottom edge. Sheet i is placed where the accumulated opacity
  // 1-(1-STEP)^i reaches the linear ramp, so the blend is even top to bottom.
  const STEP = 0.09;
  for (let i = 0, cover = 0; cover < 0.995; i++) {
    const next = 1 - Math.pow(1 - STEP, i + 1);
    const yTop = wy + wh * ((cover + next) / 2);
    s.addShape('rect', {
      x: 0.528, y: yTop, w: 12.805, h: 7.5 - yTop,
      fill: { color: C.blue, transparency: Math.round(100 - STEP * 100) },
    });
    cover = next;
  }

  brandLockup(s, 1.17, 0.821, C.blueTint);
  headline(s, [
    ['Executive Approaches ', C.blueTint],
    ['to Retention and ', C.blueSoft],
    ['Development', C.blueTint],
  ], { x: 1.17, y: 1.384, w: 6.601, h: 2.121 });
  s.addText('195%', {
    x: 10.008, y: 4.406, w: 2.548, h: 1.105, fontFace: HEAD, fontSize: 66, color: C.blueTint,
    align: 'center', valign: 'top', charSpacing: -1.5, lineSpacingMultiple: 0.9, isTextBox: true,
  });
  s.addShape('line', { x: 7.385, y: 5.655, w: 4.949, h: 0, line: { color: C.blueSoft, width: 1 } });
  copy(s, LOREM_SHORT, { x: 7.771, y: 6.017, w: 4.562, h: 0.535, align: 'right', lineSpacingMultiple: 1.5 });
  statPin(s, 9.025, 0.917, 9.206, 2.41, 9.735, 3.196, true);
  statPin(s, 4.546, 5.18, 4.727, 4.426, 5.255, 4.197, false);
}

function slide13(pptx) { // Pricing table
  const s = pptx.addSlide({ masterName: 'CHROME' });
  headline(s, [['Our Pricing Plan', C.ink]], { x: 3.377, y: 0.715, w: 6.579, h: 0.774, align: 'center' });

  const plans = [
    { x: 1.139, name: 'Business', price: '$123', dark: false },
    { x: 4.025, name: 'Advanced', price: '$213', dark: true },
    { x: 6.928, name: 'Plus Package', price: '$313', dark: false },
    { x: 9.831, name: 'Premium ', price: '$563', dark: false },
  ];
  const bullets = ['Sed ut perspiciatis ', 'Unde omnis iste natus ', 'Error sit voluptatem ', 'Remque laudantium, totam'];

  plans.forEach(({ x, name, price, dark }) => {
    const card = dark ? C.blue : C.white;
    const heading = dark ? C.blueTint : C.ink;
    const note = dark ? C.blueTint : C.grey;
    s.addShape('rect', {
      x, y: 2.134, w: 2.641, h: 4.227, fill: { color: card },
      shadow: { type: 'outer', blur: 60, offset: 30, angle: 90, color: '000000', opacity: 0.15 },
    });
    scallopBand(s, x, 2.049, 2.641, 0.172, 8, card, false);
    scallopBand(s, x, 6.272, 2.641, 0.172, 8, card, true);

    const p = x + 0.195; // text column inside the card
    s.addText(name, {
      x: p, y: 2.259, w: 2.148, h: 0.386, fontFace: HEAD, fontSize: 14, color: heading,
      valign: 'bottom', lineSpacingMultiple: 1.3, isTextBox: true,
    });
    copy(s, 'For solo entrerpreaner', { x: p, y: 2.587, w: 2.152, h: 0.285, color: note, lineSpacingMultiple: 1.3 });
    s.addText(price, {
      x: p, y: 2.931, w: 2.148, h: 0.749, fontFace: HEAD, fontSize: 32,
      color: dark ? C.white : C.blue, valign: 'bottom', lineSpacingMultiple: 1.3, isTextBox: true,
    });
    copy(s, '/month', { x: p + 1.122, y: 3.315, w: 0.815, h: 0.285, color: note, lineSpacingMultiple: 1.3 });
    s.addText('Subcribe', {
      x: p + 0.05, y: 3.766, w: 2.148, h: 0.389, shape: 'roundRect', fill: { color: C.blueMid },
      line: { color: C.white, width: 0.5 }, rectRadius: radius(2.148, 0.389, 0.21755),
      fontFace: HEAD, fontSize: 12, color: C.white, align: 'center', valign: 'middle',
      shadow: { type: 'outer', blur: 40, offset: 12, angle: 45, color: '000000', opacity: 0.1 },
    });
    s.addText('Free features', {
      x: p + 0.052, y: 4.328, w: 2.148, h: 0.39, fontFace: HEAD, fontSize: 14, color: heading,
      valign: 'bottom', lineSpacingMultiple: 1.3, isTextBox: true,
    });
    const bulletRuns = [];
    bullets.forEach((t) => {
      bulletRuns.push({ text: '\u25B8   ', options: { color: dark ? C.blueTint : C.blue } });
      bulletRuns.push({ text: t, options: { breakLine: true } });
    });
    s.addText(bulletRuns, {
      x: p + 0.05, y: 4.815, w: 2.152, h: 1.128, fontFace: BODY, fontSize: 9, color: note,
      valign: 'top', lineSpacingMultiple: 1.3, paraSpaceAfter: 6, isTextBox: true,
    });
  });

  s.addText('Popular', {
    x: 5.842, y: 2.277, w: 0.776, h: 0.211, shape: 'roundRect', rotate: 14.8,
    fill: { color: C.blueMid }, line: { color: C.white, width: 0.5 },
    rectRadius: radius(0.776, 0.211, 0.21755),
    fontFace: HEAD, fontSize: 8, color: C.white, align: 'center', valign: 'middle',
  });
}

function slide14(pptx) { // Thank you
  const s = pptx.addSlide({ masterName: 'CHROME' });
  s.background = { color: C.blue };
  swoosh(s, 0.343, 0.493, 6.148, 7.431, 285, 22);
  headline(s, [['Thank You\u00ae', C.white]], { x: 5.622, y: 2.934, w: 6.919, h: 1.447, fontSize: 80 });
  s.addShape('roundRect', {
    x: 5.838, y: 4.451, w: 1.673, h: 0.614, fill: { color: C.blueSoft },
    rectRadius: radius(1.673, 0.614, 0.18727),
    shadow: { type: 'outer', blur: 40, offset: 24, angle: 45, color: '000000', opacity: 0.25 },
  });
  copy(s, '2080', {
    x: 6.089, y: 4.464, w: 1.17, h: 0.56, fontSize: 20, color: C.blue,
    align: 'center', charSpacing: 3, lineSpacingMultiple: 1.5,
  });
  copy(s, 'Presented by.', { x: 5.838, y: 6.072, w: 1.008, h: 0.252, color: C.blueSoft, wrap: false });
  copy(s, 'Company Name', { x: 5.838, y: 6.375, w: 1.352, h: 0.286, fontSize: 11, color: C.blueTint, wrap: false });
  copy(s, 'Website.', { x: 8.958, y: 6.072, w: 0.714, h: 0.252, color: C.blueSoft, wrap: false });
  copy(s, 'www.company.tld', { x: 8.958, y: 6.375, w: 1.476, h: 0.286, fontSize: 11, color: C.blueTint, wrap: false });
}

function slide15(pptx) { // Brand sheet
  const s = pptx.addSlide({ masterName: 'CHROME' });
  [[0.537, 4.309, 4.33, 0], [4.867, 3.776, 8.467, 0], [4.875, 0, 0, 7.5]].forEach(([x, y, w, h]) => {
    s.addShape('line', { x, y, w, h, line: { color: C.hair, width: 0.75 } });
  });

  brandBadge(s, 1.893, 1.12, 0.806);
  headline(s, [['Executive', C.ink]], { x: 0.961, y: 1.998, w: 2.671, h: 0.64, fontSize: 32, align: 'center' });
  copy(s, 'Logo \u2013 Space Grotesk', {
    x: 0.959, y: 2.633, w: 2.675, h: 0.352, fontSize: 11, color: C.blueSoft, align: 'center', lineSpacingMultiple: 1.5,
  });

  [[C.deep === 'x' ? '' : '1E1005', 1.663, 5.214], [C.blue, 2.122, 5.214], [C.blueMid, 2.581, 5.214],
   [C.blueTint, 1.663, 5.633], [C.blueSoft, 2.122, 5.633], ['C5D0D9', 2.581, 5.633]]
    .forEach(([color, x, y]) => s.addShape('ellipse', { x, y, w: 0.348, h: 0.348, fill: { color } }));
  copy(s, 'Pallete', {
    x: 1.703, y: 6.053, w: 1.189, h: 0.352, fontSize: 11, color: C.blueSoft, align: 'center', lineSpacingMultiple: 1.5,
  });

  copy(s, 'Tittle  \u2013 Space Grotesk', {
    x: 5.691, y: 1.049, w: 4.24, h: 0.352, fontSize: 11, color: C.blueSoft, lineSpacingMultiple: 1.5,
  });
  headline(s, [['Sed Ut Perspiciatis Unde Omnis Iste Natus Error Sit', C.ink]], {
    x: 5.691, y: 1.414, w: 7.071, h: 1.313, fontSize: 36,
  });
  copy(s, 'Long Text - DM Sans', {
    x: 5.691, y: 4.637, w: 2.188, h: 0.352, fontSize: 11, color: C.blueSoft, lineSpacingMultiple: 1.5,
  });
  copy(s,
    LOREM + ' explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut ' +
    'fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. ' +
    'Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci ' +
    'velit, sed quia non numquam eius modi tempora incidunt ut labore et dolore magnam',
    { x: 5.691, y: 5.099, w: 6.432, h: 1.442, lineSpacingMultiple: 1.5 });
}

/* --------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'Executive';
  pptx.title = 'Executive Business';

  pptx.defineSlideMaster({
    title: 'CHROME',
    background: { color: C.paper },
    objects: [
      { rect: { x: 0, y: 0, w: 0.528, h: 7.5, fill: { color: C.paper },
        shadow: { type: 'outer', blur: 26, offset: 3, angle: 180, color: '000000', opacity: 0.1 } } },
      { text: { text: 'Executive.', options: {
        x: -0.723, y: 3.607, w: 1.974, h: 0.286, rotate: 270, align: 'center', valign: 'middle',
        fontFace: HEAD, fontSize: 11, color: C.inkMid } } },
      { text: { text: '2026', options: {
        x: -0.106, y: 6.735, w: 0.739, h: 0.269, rotate: 270, align: 'left', valign: 'middle',
        fontFace: BODY, fontSize: 10, color: C.grey } } },
    ],
    slideNumber: { x: 0.03, y: 0.261, w: 0.469, h: 0.252, align: 'center',
      fontFace: BODY, fontSize: 9, color: C.grey },
  });

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0473ed37-274a-4f1d-8b37-217b701f02b4_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
