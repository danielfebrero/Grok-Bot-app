/**
 * DataSphere — Data Science Presentation Template
 * Standalone pptxgenjs recreation of the 20-slide reference deck (13.333in x 7.5in).
 *
 * Run: node 17a6a07b-d3c9-4fbb-b5b6-b8c3a443ac57_grok_final.js
 */
'use strict'

const path = require('path')
const PptxGenJS = require('pptxgenjs')

// ---------------------------------------------------------------------------
// Theme palette (resolved from the deck theme + lumMod/lumOff derivations)
// ---------------------------------------------------------------------------
const C = {
  blue: '0000FE', // accent1
  blueDark: '0000BE', // accent1 lumMod 75%
  mint: '9DEDB9', // accent2
  periwinkle: '9899FF', // accent3
  charcoal: '323335', // accent4
  grey: 'BBBBBB', // accent5
  white: 'FFFFFF',
  ink: '181717', // bg2 lumMod 10%  — headline near-black
  slate: '3B3838', // bg2 lumMod 25%  — sub-heading grey
  mute: '767171', // bg2 lumMod 50%  — body copy grey
  ink85: '262626', // tx1 85/15
  ink75: '404040', // tx1 75/25
  ink50: '808080', // tx1 50/50
  wash: 'F2F2F2', // bg1 lumMod 95% — decorative blobs
  navy: '171C24',
  pill: 'F1F1F1',
}

const FONT_H = 'DM Sans 36pt SemiBold' // theme major font
const FONT_B = 'DM Sans 18pt Light' // theme minor font
const BLOB = { color: C.wash, transparency: 65 } // decorative petal wash (alpha 35%)

// ---------------------------------------------------------------------------
// Custom geometry, normalised to the 0..1 unit box of each shape.
// Encoded as flat [cmd, x, y, ...] runs: 'M' move, 'L' line, 'C' cubic, 'Z' close.
// ---------------------------------------------------------------------------

// The DataSphere "petal" mark: four rounded arms around a centre.
const PETAL_MARK = [
  'M',0.4996,0.6578, 'C',0.5391,0.6594,0.5845,0.6645,0.6276,0.6862,
  'C',0.6707,0.7079,0.7395,0.7591,0.758,0.7879, 'C',0.7766,0.8167,0.7529,0.8458,0.7391,0.8588,
  'C',0.7252,0.8718,0.7059,0.8845,0.6751,0.8659, 'C',0.6519,0.852,0.6126,0.8053,0.5815,0.7732,
  'L',0.5646,0.7575, 'L',0.5646,0.9357, 'C',0.5646,0.9712,0.5357,1,0.5001,1,
  'C',0.4644,1,0.4356,0.9712,0.4356,0.9357, 'L',0.4356,0.7662, 'L',0.4265,0.775,
  'C',0.3964,0.8062,0.3558,0.8512,0.3312,0.8636, 'C',0.2984,0.8801,0.2696,0.8644,0.2553,0.8494,
  'C',0.2411,0.8344,0.2233,0.8025,0.2458,0.7737, 'C',0.2684,0.7449,0.3482,0.6961,0.3905,0.6768,
  'C',0.4328,0.6575,0.4601,0.6563,0.4996,0.6578, 'Z', 'M',0.795,0.234,
  'C',0.8169,0.2302,0.8377,0.2428,0.849,0.2534, 'C',0.864,0.2676,0.8798,0.2964,0.8632,0.3291,
  'C',0.8507,0.3536,0.8056,0.3941,0.7743,0.4242, 'L',0.7655,0.4331, 'L',0.9355,0.4331,
  'C',0.9711,0.4331,1,0.462,1,0.4975, 'C',1,0.533,0.9711,0.5618,0.9355,0.5618, 'L',0.7568,0.5618,
  'L',0.7726,0.5787, 'C',0.8048,0.6097,0.8516,0.6489,0.8656,0.672,
  'C',0.8841,0.7027,0.8715,0.722,0.8584,0.7358, 'C',0.8454,0.7496,0.8162,0.7733,0.7873,0.7547,
  'C',0.7584,0.7362,0.7071,0.6676,0.6853,0.6247, 'C',0.6636,0.5817,0.6585,0.5364,0.6569,0.497,
  'C',0.6553,0.4576,0.6565,0.4304,0.6758,0.3882, 'C',0.6952,0.3461,0.7442,0.2664,0.7731,0.244,
  'C',0.7803,0.2384,0.7877,0.2353,0.795,0.234, 'Z', 'M',0.1889,0.2339,
  'C',0.1941,0.2331,0.1995,0.233,0.205,0.234, 'C',0.2123,0.2353,0.2197,0.2384,0.2269,0.244,
  'C',0.2558,0.2664,0.3048,0.3461,0.3242,0.3882, 'C',0.3435,0.4304,0.3447,0.4576,0.3431,0.497,
  'C',0.3415,0.5364,0.3364,0.5817,0.3147,0.6247, 'C',0.2929,0.6676,0.2416,0.7362,0.2127,0.7547,
  'C',0.1838,0.7733,0.1546,0.7496,0.1416,0.7358, 'C',0.1285,0.722,0.1159,0.7027,0.1344,0.672,
  'C',0.1484,0.6489,0.1952,0.6097,0.2274,0.5787, 'L',0.2432,0.5618, 'L',0.0645,0.5618,
  'C',0.0289,0.5618,0,0.533,0,0.4975, 'C',0,0.462,0.0289,0.4331,0.0645,0.4331, 'L',0.2345,0.4331,
  'L',0.2257,0.4242, 'C',0.1944,0.3941,0.1493,0.3536,0.1368,0.3291,
  'C',0.1202,0.2964,0.136,0.2676,0.151,0.2534, 'C',0.1595,0.2455,0.1733,0.2364,0.1889,0.2339, 'Z',
  'M',0.4999,0, 'C',0.5356,0,0.5644,0.0288,0.5644,0.0643, 'L',0.5644,0.2338, 'L',0.5735,0.225,
  'C',0.6036,0.1938,0.6442,0.1488,0.6688,0.1364, 'C',0.7016,0.1199,0.7304,0.1356,0.7447,0.1506,
  'C',0.7589,0.1656,0.7767,0.1975,0.7542,0.2263, 'C',0.7316,0.2551,0.6518,0.3039,0.6095,0.3232,
  'C',0.5672,0.3425,0.5399,0.3437,0.5004,0.3422, 'C',0.4609,0.3406,0.4155,0.3355,0.3724,0.3138,
  'C',0.3293,0.2921,0.2605,0.2409,0.242,0.2121, 'C',0.2234,0.1833,0.2471,0.1542,0.2609,0.1412,
  'C',0.2748,0.1282,0.2941,0.1155,0.3249,0.1341, 'C',0.3481,0.148,0.3874,0.1947,0.4185,0.2268,
  'L',0.4354,0.2425, 'L',0.4354,0.0643, 'C',0.4354,0.0288,0.4643,0,0.4999,0, 'Z',
]

// A single petal arm, used in mirrored pairs for the big background blobs.
const LEAF_BLOB = [
  'M',0.4999,0, 'C',0.5671,0,0.6217,0.0841,0.6217,0.1878, 'L',0.6217,0.6824, 'L',0.6387,0.6568,
  'C',0.6955,0.5657,0.7722,0.4344,0.8186,0.3982, 'C',0.8805,0.3499,0.935,0.3959,0.9619,0.4396,
  'C',0.9887,0.4833,1.0223,0.5764,0.9798,0.6604, 'C',0.9372,0.7444,0.7865,0.887,0.7067,0.9433,
  'C',0.6269,0.9997,0.5754,1.0032,0.5008,0.9986, 'C',0.4262,0.994,0.3404,0.979,0.2591,0.9157,
  'C',0.1778,0.8525,0.048,0.703,0.0129,0.619, 'C',-0.0222,0.535,0.0226,0.4499,0.0487,0.412,
  'C',0.0748,0.374,0.1114,0.3372,0.1696,0.3913, 'C',0.2132,0.4318,0.2875,0.5681,0.3461,0.6619,
  'L',0.3781,0.7078, 'L',0.3781,0.1878, 'C',0.3781,0.0841,0.4326,0,0.4999,0, 'Z',
]

// Two-arm variant used on the closing slide.
const SPLIT_BLOB = [
  'M',0.3119,0.3363, 'C',0.3498,0.377,0.4016,0.4635,0.4238,0.5122,
  'C',0.4491,0.5678,0.4507,0.6037,0.4486,0.6557, 'C',0.4465,0.7077,0.4398,0.7675,0.4114,0.8241,
  'L',0.4092,0.8282, 'L',0.2262,1, 'L',0.2238,0.9992, 'C',0.2074,0.9922,0.1936,0.9799,0.1851,0.9708,
  'C',0.168,0.9526,0.1515,0.9271,0.1758,0.8865, 'C',0.194,0.8561,0.2552,0.8043,0.2973,0.7635,
  'L',0.318,0.7412, 'L',0.0844,0.7412, 'C',0.0378,0.7412,0,0.7032,0,0.6563,
  'C',0,0.6095,0.0378,0.5714,0.0844,0.5714, 'L',0.3066,0.5714, 'L',0.2951,0.5596,
  'C',0.2542,0.52,0.1952,0.4666,0.1789,0.4342, 'C',0.1572,0.391,0.1778,0.3531,0.1975,0.3344,
  'C',0.2085,0.3238,0.2266,0.3118,0.247,0.3086, 'C',0.2538,0.3075,0.2609,0.3074,0.268,0.3087,
  'C',0.2776,0.3104,0.2873,0.3145,0.2967,0.3219, 'C',0.3014,0.3256,0.3065,0.3305,0.3119,0.3363, 'Z',
  'M',0.9736,0.1987, 'C',0.9875,0.2135,1.0041,0.2409,0.9991,0.2697, 'L',0.9971,0.2761, 'L',0.8982,0.369,
  'L',0.8732,0.3848, 'C',0.8452,0.4018,0.8176,0.4169,0.7969,0.4264,
  'C',0.7416,0.4519,0.7059,0.4535,0.6543,0.4514, 'C',0.6026,0.4493,0.5432,0.4426,0.4868,0.414,
  'C',0.4305,0.3854,0.3406,0.3178,0.3163,0.2798, 'C',0.292,0.2419,0.323,0.2034,0.3411,0.1862,
  'C',0.3592,0.1691,0.3845,0.1524,0.4248,0.1769, 'C',0.4551,0.1952,0.5065,0.2568,0.5471,0.2992,
  'L',0.5692,0.32, 'L',0.5692,0.0849, 'C',0.5692,0.038,0.607,0,0.6536,0,
  'C',0.7002,0,0.738,0.038,0.738,0.0849, 'L',0.738,0.3085, 'L',0.7498,0.2969,
  'C',0.7891,0.2557,0.8422,0.1964,0.8744,0.18, 'C',0.9173,0.1582,0.955,0.179,0.9736,0.1987, 'Z',
]

// Outlined hexagon pair behind the doughnut chart on slide 13.
const HEX_OUTER = [
  'M',0.5,0, 'L',0.9008,0.1979, 'L',1,0.6429, 'L',0.7225,1, 'L',0.2775,1, 'L',0,0.6429, 'L',0.0992,0.1979, 'Z',
]
const HEX_INNER = [
  'M',0.5655,0, 'L',0.6966,0.3471, 'L',1,0.4981, 'L',0.7589,0.7742, 'L',0.2759,1, 'L',0.2759,0.4981,
  'L',0,0.089, 'Z',
]

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Expand a flat [cmd,x,y,...] run into pptxgenjs custom-geometry points. */
function toPoints (flat, w, h) {
  const pts = []
  let i = 0
  while (i < flat.length) {
    const cmd = flat[i++]
    if (cmd === 'Z') { pts.push({ close: true }); continue }
    if (cmd === 'C') {
      const [x1, y1, x2, y2, x, y] = flat.slice(i, i + 6); i += 6
      pts.push({ x: x * w, y: y * h, curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h } })
    } else {
      const [x, y] = flat.slice(i, i + 2); i += 2
      pts.push({ x: x * w, y: y * h, moveTo: cmd === 'M' })
    }
  }
  return pts
}

/** Draw one of the custom-geometry marks above at x/y/w/h with an optional rotation. */
function mark (slide, flat, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: toPoints(flat, o.w, o.h),
    fill: o.fill, line: o.line || { type: 'none' },
    rotate: o.rotate || 0, flipH: o.flipH || false, flipV: o.flipV || false,
  })
}

/** Big wash-coloured petal that bleeds off the slide edges. */
function petalWash (slide, x, y, w, h, rotate) {
  mark(slide, PETAL_MARK, { x, y, w, h, fill: BLOB, rotate })
}

/**
 * Mirrored leaf pair used as an alternate background motif.
 * Both arms are 69.26% x 44.99% of the motif box; the second is rotated a quarter turn.
 */
function leafWash (slide, x, y, w, h) {
  const aw = w * 0.6926
  const ah = h * 0.4499
  mark(slide, LEAF_BLOB, { x: x + w * 0.3074, y, w: aw, h: ah, fill: BLOB })
  mark(slide, LEAF_BLOB, { x: x - w * 0.1217, y: y + h * 0.4282, w: aw, h: ah, fill: BLOB, rotate: 90, flipH: true, flipV: true })
}

/** Horizontal tab whose left or right end is a half-round cap (`cap`: 'l' | 'r' | null). */
function tab (slide, o) {
  const r = o.h / 2
  const pts = o.cap === 'l'
    ? [{ x: r, y: 0, moveTo: true }, { x: o.w, y: 0 }, { x: o.w, y: o.h }, { x: r, y: o.h },
      { x: r, y: 0, curve: { type: 'arc', hR: r, wR: r, stAng: 90, swAng: 180 } }, { close: true }]
    : o.cap === 'r'
      ? [{ x: 0, y: 0, moveTo: true }, { x: o.w - r, y: 0 },
        { x: o.w - r, y: o.h, curve: { type: 'arc', hR: r, wR: r, stAng: 270, swAng: 180 } },
        { x: 0, y: o.h }, { close: true }]
      : [{ x: 0, y: 0, moveTo: true }, { x: o.w, y: 0 }, { x: o.w, y: o.h }, { x: 0, y: o.h }, { close: true }]
  slide.addShape('custGeom', { x: o.x, y: o.y, w: o.w, h: o.h, points: pts, fill: { color: o.fill }, line: { type: 'none' } })
}

/** Rounded rectangle. `r` is the OOXML adj value (fraction of the shorter side). */
function roundRect (slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill }, line: { type: 'none' },
    rectRadius: (o.r === undefined ? 0.16667 : o.r) * Math.min(o.w, o.h),
  })
}

/**
 * Text box. `runs` is a string or an array of { text, color, bold, size, ... } run specs.
 * Positions come straight from the reference deck (inches, top-left origin).
 */
function text (slide, runs, o) {
  const body = typeof runs === 'string' ? [{ text: runs }] : runs
  slide.addText(body.map(r => ({ text: r.text, options: r })), Object.assign({
    fontFace: FONT_B, fontSize: 18, color: C.slate,
    margin: [7.2, 7.2, 3.6, 3.6], valign: 'top', isTextBox: true,
  }, o))
}

/** Headline: mixed-colour runs at 90% line spacing in the major font. */
function headline (slide, runs, o) {
  return text(slide, runs.map(r => Object.assign({ fontFace: FONT_H, fontSize: o.fontSize || 54 }, r)),
    Object.assign({ lineSpacingMultiple: 0.9, color: C.ink }, o))
}

/** Body paragraph in the minor font. */
function body (slide, str, o) {
  return text(slide, str, Object.assign({ fontSize: 12, color: C.mute, lineSpacingMultiple: 1.2 }, o))
}

/** Master furniture: petal glyph + "DataSphere" wordmark + running head + page number. */
function chrome (slide, n, opts) {
  const o = opts || {}
  mark(slide, PETAL_MARK, { x: 0.4703, y: 0.425, w: 0.2381, h: 0.2388, fill: { color: C.mint }, rotate: 46.34 })
  text(slide, [
    { text: 'Data' }, { text: 'Sphere', bold: true },
  ], { x: 0.6395, y: 0.3929, w: 1.105, h: 0.303, fontSize: 12, color: C.ink85, wrap: false })
  text(slide, 'Data Science Presentation',
    { x: 10.682, y: 0.3929, w: 2.23, h: 0.303, fontSize: 12, color: o.headColor || C.slate, align: 'right', wrap: false })
  text(slide, String(n),
    { x: 12.1936, y: 6.803, w: 0.7168, h: 0.303, fontSize: 12, fontFace: FONT_H, color: o.numColor || C.slate, align: 'right' })
}

// ---------------------------------------------------------------------------
// Shared copy
// ---------------------------------------------------------------------------
const LOREM_SHORT = 'Cras at tellus in mauris varius lacinia. Sed erat velit, porta vitae tortor sed, convallis gravida orci. Proin imperdiet tellus ac felis molestie, nec vehicula ante. '
const LOREM_QA = 'Lorem ipsum dolor amet, elit, sed do eiusmod tempor incididunt ut laboree dolore magna aliqua. '
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt'
const FAR_AWAY = 'Far far away, behind the word'

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

// 1 — Cover: oversized wordmark over a rotated petal wash.
function slide01 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 6.3425, 1.0434, 8.0692, 8.0921, 6.79)
  text(s, 'Data Science Presentation Template',
    { x: 1.4381, y: 1.3987, w: 5.4826, h: 0.4042, fontSize: 18, fontFace: FONT_H, color: C.slate })
  headline(s, [{ text: 'DataSphere Presentation Template', color: C.blueDark }],
    { x: 1.4381, y: 1.8817, w: 6.4517, h: 3.3878, fontSize: 70 })
  mark(s, PETAL_MARK, { x: 1.3369, y: 5.5497, w: 0.597, h: 0.5987, fill: { color: C.mint }, rotate: 45.16 })
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit euismoodyantr amutie consectetur.',
    { x: 2.2917, y: 5.5967, w: 3.4157, h: 0.5053, lineSpacingMultiple: 1 })
}

// 2 — "Why Data Matters": two stat chips on the right.
function slide02 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -2.4542, -3.1647, 7.8979, 7.9203, 0)
  chrome(s, 2)
  headline(s, [
    { text: 'Why ' }, { text: 'Data', color: C.blue }, { text: ' Matters More Than Ever' },
  ], { x: 1.4948, y: 1.8918, w: 4.9311, h: 2.6119, fontSize: 55 })
  body(s, [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt labore dolore ',
    'magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi commodo consequat.\neiusmod tempor incididunt labore et dolore',
  ].join('\n'), { x: 1.4948, y: 4.7561, w: 4.7409, h: 1.4692, paraSpaceAfter: 12 })

  roundRect(s, { x: 6.9075, y: 1.9474, w: 1.2552, h: 0.9788, fill: C.mint })
  text(s, '29K', { x: 7.0516, y: 2.1339, w: 0.9669, h: 0.6058, fontSize: 30, fontFace: FONT_H, color: C.ink, align: 'center' })
  roundRect(s, { x: 10.583, y: 4.9717, w: 1.2552, h: 0.9788, fill: C.blue })
  text(s, '29K', { x: 10.7272, y: 5.1582, w: 0.9669, h: 0.6058, fontSize: 30, fontFace: FONT_H, color: C.white, align: 'center' })
}

// 3 — "Exploring the World Data Science" + mint stat panel bleeding off the right edge.
function slide03 (pptx) {
  const s = pptx.addSlide()
  mark(s, LEAF_BLOB, { x: 5.9142, y: 0.4227, w: 4.9589, h: 3.2166, fill: { color: 'E7E6E6', transparency: 70 }, rotate: 316.16 })
  chrome(s, 3)
  headline(s, [
    { text: 'Exploring the World' }, { text: ' ', color: C.navy }, { text: 'Data Science', color: C.blue },
  ], { x: 1.1766, y: 1.7031, w: 7.3162, h: 1.7484 })
  roundRect(s, { x: 8.8731, y: 1.7971, w: 5.1112, h: 4.4566, fill: C.mint, r: 0.09188 })
  text(s, '150+', { x: 9.6059, y: 2.3752, w: 2.3813, h: 0.9416, fontSize: 50, fontFace: FONT_H, color: C.ink })
  text(s, 'Data Professionals', { x: 9.6807, y: 3.4569, w: 2.7134, h: 0.4042, fontSize: 18, fontFace: FONT_H, color: C.slate })
  body(s, 'Cras at tellus in torto mauris varius lacinia. Sed erat velit, porta vitae tortor sed, ronan convallis gravida orci. Proin imperdiet tellus convallis.',
    { x: 9.6807, y: 4.0426, w: 3.0146, h: 1.6321, fontSize: 14, color: C.slate, lineSpacingMultiple: 1.3 })
}

// 4 — "What Is Data Science?" with a mint definition card.
function slide04 (pptx) {
  const s = pptx.addSlide()
  leafWash(s, 4.8157, 1.8432, 5.7259, 5.717)
  chrome(s, 4)
  headline(s, [
    { text: 'What Is ' }, { text: 'Data Science?', color: C.blue },
  ], { x: 1.1006, y: 1.8679, w: 5.2364, h: 1.7484 })
  roundRect(s, { x: 7.2561, y: 4.3169, w: 4.9609, h: 2.0672, fill: C.mint, r: 0.11292 })
  text(s, 'Data Science', { x: 7.6788, y: 4.6807, w: 1.9425, h: 0.4681, fontSize: 18, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 1.3 })
  body(s, LOREM_SHORT, { x: 7.6788, y: 5.1488, w: 4.1163, h: 0.8713, color: C.slate, lineSpacingMultiple: 1.3 })
}

// 5 — "Core Activities": two numbered chips plus a bulleted feature list.
function slide05 (pptx) {
  const s = pptx.addSlide()
  leafWash(s, 6.2138, 0.1429, 4.7182, 4.7108)
  chrome(s, 5)
  // headline / intro sit left; feature list and chips stack down the page
  headline(s, [
    { text: 'Core Activities in ' }, { text: 'Data Science', color: C.blue },
  ], { x: 1.3141, y: 1.9472, w: 5.9104, h: 1.7484 })
  body(s, 'Far far away, behind the word mountains, far from the countries Vokalia and Consonantia, there live the blind texts separated',
    { x: 1.3141, y: 3.7236, w: 5.2402, h: 0.5731 })

  // Two feature rows: petal glyph + "Data Science," lead-in + grey remainder.
  const rows = [
    { y: 4.9106, gy: 4.9653, glyph: C.blue, tail: 'behind the word mountains, far from the countries Vokalia 2026' },
    { y: 5.6625, gy: 5.6995, glyph: C.mint, tail: 'behind the word mountains, far from the countries ' },
  ]
  rows.forEach(r => {
    mark(s, PETAL_MARK, { x: 1.4179, y: r.gy, w: 0.4987, h: 0.5001, fill: { color: r.glyph }, rotate: 46.34 })
    text(s, [
      { text: 'Data Science', color: C.slate }, { text: ', ', color: C.charcoal }, { text: r.tail, color: C.mute },
    ], { x: 2.1301, y: r.y, w: 3.5077, h: 0.5731, fontSize: 12, lineSpacingMultiple: 1.2 })
  })

  // Numbered chips (02 on top, 01 below — as in the reference).
  const chips = [
    { y: 3.9159, fill: C.blue, num: '02', fg: C.white },
    { y: 4.8543, fill: C.mint, num: '01', fg: C.slate },
  ]
  chips.forEach(c => {
    roundRect(s, { x: 7.3039, y: c.y, w: 2.5387, h: 0.8101, fill: c.fill, r: 0.15845 })
    text(s, c.num, { x: 7.4996, y: c.y + 0.0471, w: 0.8446, h: 0.7137, fontSize: 32, fontFace: FONT_H, color: c.fg, lineSpacingMultiple: 1.2 })
    text(s, FAR_AWAY, { x: 8.1885, y: c.y + 0.1176, w: 1.4592, h: 0.5731, fontSize: 12, color: c.fg, lineSpacingMultiple: 1.2 })
  })
}

// 6 — "Insights from Big Data": mint note card and a 78% metric block.
function slide06 (pptx) {
  const s = pptx.addSlide()
  leafWash(s, 10.3891, 4.3581, 3.8415, 3.8358)
  chrome(s, 6)
  headline(s, [
    { text: 'Insights from ' }, { text: 'Big Data', color: C.blue },
  ], { x: 1.0235, y: 1.302, w: 4.7028, h: 1.7484 })
  body(s, 'Lorem ipsum dolor sit amet, consetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua sed do ',
    { x: 1.0235, y: 3.1587, w: 3.9527, h: 0.8154 })

  roundRect(s, { x: 6.2808, y: 1.3017, w: 2.4117, h: 2.0363, fill: C.mint, r: 0.11661 })
  text(s, 'Data collection', { x: 6.4478, y: 1.6296, w: 2.0766, h: 0.4462, fontSize: 18, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 1.2 })
  text(s, 'Far far away, behind the word mountains, from the countries.',
    { x: 6.489, y: 2.0762, w: 2.0372, h: 0.9345, fontSize: 14, color: C.slate, lineSpacingMultiple: 1.2 })

  text(s, 'Data collection', { x: 1.0235, y: 5.1103, w: 2.0766, h: 0.4462, fontSize: 18, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 1.2 })
  text(s, 'Far far away, behind the word mountains.',
    { x: 1.0235, y: 5.6238, w: 2.2777, h: 0.6521, fontSize: 14, color: C.mute, lineSpacingMultiple: 1.2 })
  text(s, '78%', { x: 3.6069, y: 4.9024, w: 1.7663, h: 0.7783, fontSize: 48, fontFace: 'DM Sans ExtraBold', color: C.slate, valign: 'middle', lineSpacingMultiple: 0.8 })
  text(s, FAR_AWAY, { x: 3.6069, y: 5.6238, w: 1.9672, h: 0.6521, fontSize: 14, color: C.mute, lineSpacingMultiple: 1.2 })
}

// 7 — "The Future of Data Science": tall stacked headline, two body paragraphs.
function slide07 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 6.2276, -1.0194, 7.8979, 7.9203, 0)
  chrome(s, 7)
  headline(s, [
    { text: 'The Future of ' }, { text: 'Data Science', color: C.blue },
  ], { x: 1.1267, y: 1.2485, w: 3.2059, h: 3.3845 })
  body(s, [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud.',
    'nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non. Duis aute irure dolor in reprehenderit in voluptate velit esse. ',
  ].join('\n'), { x: 1.1267, y: 4.8677, w: 7.2306, h: 1.4685, paraSpaceAfter: 12 })

  roundRect(s, { x: 8.9236, y: 1.5689, w: 2.9965, h: 1.5445, fill: C.mint, r: 0.11292 })
  text(s, 'Data collection', { x: 9.1379, y: 1.7573, w: 2.0766, h: 0.4462, fontSize: 18, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 1.2 })
  text(s, 'Far far away, behind the\nword mountains, countries',
    { x: 9.1379, y: 2.2325, w: 2.5679, h: 0.6929, fontSize: 14, color: C.slate, lineSpacingMultiple: 1.3 })
}

// 8 — "The Stages of Data Science": three numbered bands.
function slide08 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 5.7541, -0.9689, 7.8979, 7.9203, 338.42)
  chrome(s, 8)
  headline(s, [
    { text: 'The' }, { text: ' ' }, { text: 'Stages', color: C.blue }, { text: ' ' }, { text: 'of Data Science' },
  ], { x: 1.2772, y: 2.2971, w: 5.1119, h: 1.7369 })
  text(s, 'Place Your Sub Heading',
    { x: 1.2772, y: 4.3528, w: 3.0954, h: 0.4278, fontSize: 16, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 1.3 })
  body(s, [
    'Cras at tellus in mauris varius lacinia. Sed erat velit, porta vitae tortor sed, convallis gravida orci. Proin imperdiet tellus ac felis. ',
    'Anec vehicula ante convallis. Aenean varius finibwuus pulvinar tellus fringilla sed. Fusce quis auctor mi. Phasellus teli euismod, justo sit amet molestie imperdiet varius finibus .',
  ].join('\n'), { x: 1.2772, y: 4.8296, w: 5.0011, h: 1.5645, lineSpacingMultiple: 1.3, paraSpaceAfter: 12 })

  const bands = [
    { y: 1.6733, fill: C.blue, num: '01', fg: C.white, numW: 1.3423 },
    { y: 3.3229, fill: C.mint, num: '02', fg: C.ink, numW: 1.5851 },
    { y: 4.9724, fill: C.blue, num: '03', fg: C.white, numW: 1.5851 },
  ]
  bands.forEach(b => {
    roundRect(s, { x: 7.3499, y: b.y, w: 4.7058, h: 1.4224, fill: b.fill })
    text(s, b.num, { x: 7.6539, y: b.y + 0.245, w: b.numW, h: 0.9416, fontSize: 50, fontFace: FONT_H, color: b.fg, charSpacing: -0.7 })
    text(s, 'Cras at tellus in mauris varius lacinia. Sed erat velit, porta vit.',
      { x: 8.9713, y: b.y + 0.4188, w: 2.8383, h: 0.5857, fontSize: 12, color: b.fg === C.white ? C.white : C.slate, lineSpacingMultiple: 1.2, paraSpaceAfter: 12 })
  })
}

// 9 — "Skills Every Data Science Needs": stacked mint + blue stat bands.
function slide09 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 6.9301, -1.3083, 7.8979, 7.9203, 338.42)
  chrome(s, 9)
  headline(s, [
    { text: 'Skills Every ' }, { text: 'Data Science', color: C.blue }, { text: ' Needs' },
  ], { x: 1.0798, y: 1.4306, w: 6.3924, h: 1.7369 })

  roundRect(s, { x: 1.0798, y: 3.4536, w: 6.8367, h: 1.2212, fill: C.mint, r: 0.12757 })
  text(s, '80%', { x: 1.4118, y: 3.5265, w: 2.9603, h: 1.0197, fontSize: 45, fontFace: FONT_H, bold: true, color: C.slate, lineSpacingMultiple: 1.3 })
  text(s, 'Cras at in mauris varius lacinia convallis gravida orci. Proin imperdiet tellus. Cras at in mauris varius lacinia. Sed erat velit, porta vitae tortor sed, convallis.',
    { x: 3.2337, y: 3.6564, w: 4.3527, h: 0.8154, fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 })

  roundRect(s, { x: 1.0798, y: 4.848, w: 6.8367, h: 1.2212, fill: C.blue, r: 0.12757 })
  text(s, 'Data Visualization', { x: 1.4118, y: 5.0876, w: 2.9603, h: 0.4258, fontSize: 16, fontFace: FONT_H, bold: true, color: C.white, lineSpacingMultiple: 1.3 })
  text(s, 'Cras at in mauris varius lacinia. Sed erat velit, porta vitae tortor sed, convallis.',
    { x: 1.4118, y: 5.4836, w: 6.1762, h: 0.3462, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 })
}

// 10 — "Where Today's Data Comes From" + a 100+ mint tile.
function slide10 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -1.7156, -2.4113, 5.7345, 5.7508, 323.34)
  chrome(s, 10)
  headline(s, [
    { text: 'Where Today’s ' }, { text: 'Data ', color: C.blue }, { text: 'Comes From' },
  ], { x: 1.3392, y: 2.7212, w: 4.8127, h: 2.5674 })
  body(s, 'Cras at tellus in mauris varius lacinia. Sed velit, portavi\nsed, convallis gravida orci. Proin imperdiet tellus afelis\nmolestie, nec vehicula ante convallis. ',
    { x: 1.3392, y: 5.3733, w: 5.3273, h: 0.8713, lineSpacingMultiple: 1.3 })
  roundRect(s, { x: 9.2849, y: 4.3506, w: 2.7089, h: 1.6386, fill: C.mint, r: 0.11807 })
  text(s, '100+', { x: 9.5081, y: 4.4827, w: 2.2626, h: 1.1113, fontSize: 60, fontFace: FONT_H, color: C.ink, align: 'center' })
  text(s, 'Databases', { x: 9.879, y: 5.4306, w: 1.5224, h: 0.4258, fontSize: 16, fontFace: FONT_H, color: C.slate, align: 'center', lineSpacingMultiple: 1.3 })
}

// 11 — "The Business Power of Big Data": two stat chips + clustered column chart.
function slide11 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -2.6143, 3.4511, 6.5936, 6.6123, 270)
  chrome(s, 11)
  headline(s, [
    { text: 'The Business Power of ' }, { text: 'Big Data', color: C.blue },
  ], { x: 1.3208, y: 1.7699, w: 6.3642, h: 1.7484 })
  body(s, 'Cras at tellus in mauris varius lacinia. Sed velit, portavi sed, convallis gravida orci. Proin imperdiet tellus afelis molestie, nec vehicula ante convallis. ',
    { x: 1.3208, y: 3.5763, w: 6.0527, h: 0.6091, lineSpacingMultiple: 1.3 })

  const chips = [
    { x: 1.3231, w: 2.9636, fill: C.blue, big: '192+', fg: C.white, numX: 1.7276, pctX: 3.104 },
    { x: 4.5154, w: 2.858, fill: C.mint, big: '130+', fg: C.ink, numX: 4.867, pctX: 6.2428 },
  ]
  chips.forEach(c => {
    roundRect(s, { x: c.x, y: 4.6109, w: c.w, h: 1.2588, fill: c.fill, r: 0.12678 })
    text(s, c.big, { x: c.numX, y: 4.8532, w: 2.0409, h: 0.7743, fontSize: 40, fontFace: FONT_H, color: c.fg })
    text(s, '5%', { x: c.pctX, y: 4.9539, w: 0.7794, h: 0.5723, fontSize: 28, fontFace: FONT_H, color: c.fg === C.white ? C.white : C.slate, align: 'center' })
  })

  s.addChart(pptx.ChartType.bar, [
    { name: 'Income', labels: ['Category 1'], values: [4.6] },
    { name: 'Outcome', labels: ['Category 1'], values: [7] },
    { name: 'Supply', labels: ['Category 1'], values: [2.4] },
  ], {
    x: 7.913, y: 1.3272, w: 4.0988, h: 4.8467,
    layout: { x: 0.0421, y: 0.031, w: 0.9248, h: 0.8403 },
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 173, barOverlapPct: -18,
    chartColors: [C.periwinkle, C.charcoal, C.grey],
    catAxisHidden: true, valAxisHidden: true,
    valGridLine: { color: C.wash, size: 0.75 }, catGridLine: { style: 'none' },
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: 'General',
    dataLabelFontSize: 11, dataLabelColor: C.ink50, dataLabelFontFace: FONT_B,
    showLegend: true, legendPos: 'b', legendFontSize: 11, legendColor: C.ink50, legendFontFace: FONT_B,
    chartArea: { fill: { type: 'none' } },
  })
}

// 12 — "Data Science in Business Intelligence": mint bar with four percentages.
function slide12 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -2.6635, 0.989, 5.5063, 5.5219, 44.63)
  petalWash(s, 10.4628, 0.989, 5.5063, 5.5219, 44.63)
  chrome(s, 12)
  headline(s, [
    { text: 'Data Science ' }, { text: 'in Business Intelligence', color: C.blue },
  ], { x: 2.8421, y: 1.0379, w: 7.6488, h: 1.7484, align: 'center' })

  roundRect(s, { x: 1.6427, y: 3.39, w: 10.0472, h: 2.1526, fill: C.mint, r: 0.09816 })
  roundRect(s, { x: 4.0325, y: 3.096, w: 2.9128, h: 2.711, fill: C.blue, r: 0.09816 })

  // Highlighted centre figure, then the three mint-panel figures.
  text(s, '68%', { x: 4.3966, y: 3.8091, w: 2.1839, h: 1.0389, fontSize: 66, fontFace: FONT_H, color: C.white, align: 'center', valign: 'middle', charSpacing: -0.7 })
  text(s, 'Data Science', { x: 4.5191, y: 4.6775, w: 1.9398, h: 0.4147, fontSize: 20, fontFace: FONT_H, color: C.white, align: 'center', lineSpacingMultiple: 1.2 })
  const figures = [
    { x: 2.1289, lx: 2.1662, lw: 1.3554, v: '10%' },
    { x: 7.4813, lx: 7.4846, lw: 1.4239, v: '22%' },
    { x: 9.5613, lx: 9.5646, lw: 1.4239, v: '43%' },
  ]
  figures.forEach(f => {
    text(s, f.v, { x: f.x, y: 3.9907, w: 1.4297, h: 0.7929, fontSize: 44, fontFace: FONT_H, color: C.slate, align: 'center', valign: 'middle', charSpacing: -0.7 })
    text(s, 'Data Science', { x: f.lx, y: 4.6311, w: f.lw, h: 0.3311, fontSize: 12, color: C.slate, align: 'center', lineSpacingMultiple: 1.2 })
  })
  body(s, 'Cras at tellus in mauris varius lacinia. Sed velit, portavi sed',
    { x: 3.7152, y: 6.1157, w: 5.9024, h: 0.3459, align: 'center', lineSpacingMultiple: 1.3 })
}

// 13 — "Challenges & Solutions": coloured data table + exploded doughnut.
function slide13 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -0.1246, 2.5537, 5.1098, 5.1243, 48.23)
  chrome(s, 13)
  headline(s, [
    { text: 'Challenges & ', color: C.blue, bold: true }, { text: 'Solutions ', color: C.slate, bold: true },
  ], { x: 1.6987, y: 1.359, w: 5.1428, h: 1.5657, fontSize: 48 })

  // Column chrome sits behind the table: a header pill and a tall body pill per column.
  const cols = [
    { x: 1.8527, fill: C.blue, headR: 0.27828, bodyR: 0.11183, bodyY: 3.808, bodyH: 2.4219, fg: C.white, head: 'Value A', vals: ['100', '20', '55', '35', '100'] },
    { x: 3.3555, fill: C.mint, headR: 0.27828, bodyR: 0.10399, bodyY: 3.8161, bodyH: 2.4138, fg: C.slate, head: 'Value B', vals: ['50', '80', '55', '80', '20'] },
    { x: 4.8734, fill: C.periwinkle, headR: 0.25595, bodyR: 0.11183, bodyY: 3.8161, bodyH: 2.4138, fg: C.slate, head: 'Value A', vals: ['150', '100', '130', '125', '120'] },
  ]
  cols.forEach(c => {
    roundRect(s, { x: c.x, y: 3.2156, w: 1.3295, h: 0.4666, fill: c.fill, r: c.headR })
    roundRect(s, { x: c.x, y: c.bodyY, w: 1.3295, h: c.bodyH, fill: c.fill, r: c.bodyR })
  })

  const headRow = cols.map(c => ({ text: c.head, options: { fontSize: 14, fontFace: FONT_H, color: c.fg, align: 'center', valign: 'middle' } }))
  const dataRows = [0, 1, 2, 3, 4].map(i => cols.map((c, ci) => ({
    text: c.vals[i],
    options: { fontSize: 12, fontFace: FONT_B, color: ci === 0 ? C.white : (ci === 1 ? C.slate : C.ink), align: 'center', valign: 'middle' },
  })))
  s.addTable([headRow].concat(dataRows), {
    x: 1.6987, y: 3.1088, w: 4.6029, colW: [1.5343, 1.5343, 1.5343],
    rowH: [0.7091, 0.4839, 0.4839, 0.4839, 0.4839, 0.4839],
    margin: 0, border: { type: 'none' },
  })

  s.addChart(pptx.ChartType.doughnut, [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] },
  ], {
    x: 7.7192, y: 1.359, w: 3.6575, h: 4.0049,
    holeSize: 75, firstSliceAng: 0, dataNoEffects: true,
    chartColors: [C.blue, 'B5F2CA', C.mint, C.periwinkle],
    showLegend: true, legendPos: 'b', legendFontSize: 11, legendColor: C.slate, legendFontFace: FONT_B,
    chartArea: { fill: { type: 'none' } },
  })
  body(s, 'Lorem ipsum dolor sit, consectetur adipiscing elit, sed do eiusmod',
    { x: 7.6314, y: 5.4886, w: 3.8337, h: 0.6521, fontSize: 14, align: 'center' })

  // Hexagon outline motif over the doughnut hole plus a mint petal accent.
  mark(s, HEX_OUTER, { x: 9.3941, y: 3.1087, w: 1.0048, h: 1.0048, fill: { type: 'none' } })
  mark(s, HEX_INNER, { x: 9.5702, y: 3.3596, w: 0.5779, h: 0.6486, fill: { type: 'none' } })
  mark(s, PETAL_MARK, { x: 9.0699, y: 2.6292, w: 0.9563, h: 0.959, fill: { color: C.mint }, rotate: 45.16 })
}

// 14 — Section cover: "Case Studies" with three outcome cards.
function slide14 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 10.4373, 4.8331, 4.7014, 4.3521, 65.03)
  petalWash(s, -1.1739, -1.5544, 4.3418, 4.3541, 48.23)
  headline(s, [
    { text: 'Case Studies From Business, ', bold: true }, { text: 'Healthcare, And Finance', color: C.blue, bold: true },
  ], { x: 2.4249, y: 2.6067, w: 8.2354, h: 1.5819, fontSize: 44, align: 'center', lineSpacingMultiple: 1 })

  const cards = [
    { x: 1.1487, fill: C.blue, fg: C.white, big: '+5%', subColor: C.white },
    { x: 5.0285, fill: C.mint, fg: C.slate, big: '1,500,000', subColor: C.charcoal },
    { x: 8.9075, fill: C.blue, fg: C.white, big: '-20%', subColor: C.white },
  ]
  cards.forEach(c => {
    roundRect(s, { x: c.x, y: 4.4054, w: 3.2758, h: 2.1691, fill: c.fill, r: 0.09816 })
    text(s, c.big, { x: c.x + 0.2946, y: 4.6083, w: 2.53, h: 0.5219, fontSize: 25, fontFace: FONT_H, bold: true, color: c.fg, valign: 'bottom' })
    text(s, 'Smart Data Science', { x: c.x + 0.2946, y: 5.1041, w: 2.53, h: 0.4042, fontSize: 18, fontFace: FONT_H, bold: true, color: c.subColor, valign: 'bottom' })
    text(s, LOREM_CARD, { x: c.x + 0.2946, y: 5.5561, w: 2.6, h: 0.8154, fontSize: 12, color: c.fg, lineSpacingMultiple: 1.2 })
  })

  // Cover variant of the master chrome: both corner labels sit on white pills.
  roundRect(s, { x: 10.4999, y: 0.2777, w: 2.5622, h: 0.5333, fill: C.white })
  roundRect(s, { x: 0.2711, y: 0.2777, w: 1.6119, h: 0.5333, fill: C.white })
  chrome(s, 14)
}

// 15 — "Let's Dive Into Your Questions": four labelled columns.
function slide15 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 10.9032, 1.1161, 5.253, 5.2679, 344.09)
  petalWash(s, -1.7652, -2.2068, 6.0706, 6.0879, 344.09)
  chrome(s, 15)
  headline(s, [
    { text: 'Let’s Dive Into ' }, { text: 'Your Questions', color: C.blue },
  ], { x: 3.4066, y: 1.4497, w: 6.5195, h: 1.7484, align: 'center' })
  body(s, LOREM_QA, { x: 4.3277, y: 3.3287, w: 4.6778, h: 0.5731, align: 'center' })

  // Four tabs across the page; only the outer ends of the run are rounded.
  const tabs = [
    { x: 1.0128, tx: 0.8757, cap: 'l', fill: C.blue, fg: C.white, label: 'Introduction' },
    { x: 3.8263, tx: 3.8257, cap: null, fill: C.mint, fg: C.slate, label: 'Data Source' },
    { x: 6.9126, tx: 6.7757, cap: null, fill: C.blue, fg: C.white, label: 'Understanding' },
    { x: 9.8626, tx: 9.7257, cap: 'r', fill: C.periwinkle, fg: C.slate, label: 'Future Trends' },
  ]
  tabs.forEach(t => {
    tab(s, { x: t.tx, y: 4.2769, w: 2.7325, h: 0.6392, fill: t.fill, cap: t.cap })
    text(s, t.label, { x: t.x + 0.2569, y: 4.3623, w: 1.9425, h: 0.4681, fontSize: 18, fontFace: FONT_H, color: t.fg, align: 'center', lineSpacingMultiple: 1.3 })
    body(s, 'Cras at tellus in mauris varius lacinia. Proin imperdiet tellus felis molestie, vehicula ante. ',
      { x: t.x, y: 5.1789, w: 2.4581, h: 0.8713, align: 'center', lineSpacingMultiple: 1.3 })
  })
}

// 16 — "Machine Learning Models": Supervised vs Unsupervised checklists.
function slide16 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, -1.2887, -1.863, 4.6622, 4.6755, 344.09)
  chrome(s, 16)
  headline(s, [
    { text: 'Machine Learning Models in ' }, { text: 'Data Science', color: C.blue },
  ], { x: 1.2527, y: 1.4894, w: 4.9028, h: 2.1098, fontSize: 44 })
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua.',
    { x: 1.2527, y: 3.5998, w: 4.9028, h: 0.5731, paraSpaceAfter: 12 })

  const chips = [{ x: 1.3601, fill: C.blue, v: '192+', fg: C.white }, { x: 3.8236, fill: C.mint, v: '130+', fg: C.ink }]
  chips.forEach(c => {
    roundRect(s, { x: c.x, y: 5.1301, w: 2.0122, h: 1.2588, fill: c.fill, r: 0.12678 })
    text(s, c.v, { x: c.x + 0.2937, y: 5.3723, w: 1.4247, h: 0.7743, fontSize: 40, fontFace: FONT_H, color: c.fg, align: 'center' })
  })

  const panels = [
    { y: 1.1116, h: 2.4661, fill: C.mint, fg: C.slate, title: 'Supervised', last: 'Hendrerit Augue Semper.' },
    { y: 3.9028, h: 2.4855, fill: C.blue, fg: C.white, title: 'Unsupervised', last: 'Hendrerit Augue Semper Sed.' },
  ]
  panels.forEach(p => {
    roundRect(s, { x: 7.6809, y: p.y, w: 4.4084, h: p.h, fill: p.fill, r: 0.06219 })
    text(s, p.title, { x: 8.1321, y: p.y + 0.364, w: 3.0472, h: 0.4699, fontSize: 24, fontFace: FONT_H, color: p.fg, lineSpacingMultiple: 0.9 })
    text(s, ['Lorem Ipsum Dolor Sit Amet', 'Consectetur Adipiscing Elit', 'Sed Iaculis Mattis Augue', p.last].map(t => ({
      text: t, fontSize: 12, color: p.fg, bullet: { characterCode: '2713', indent: 22.5 }, lineSpacingMultiple: 1.3, breakLine: true,
    })), { x: 8.1321, y: p.y + 0.9694, w: 3.5063, h: 1.1327 })
  })
  s.addShape('ellipse', { x: 10.5148, y: 3.037, w: 1.2229, h: 1.2229, fill: { color: C.periwinkle }, line: { type: 'none' } })
  text(s, 'VS', { x: 10.7176, y: 3.3623, w: 0.8186, h: 0.5723, fontSize: 28, fontFace: FONT_H, color: C.ink, align: 'center', valign: 'middle' })
}

// 17 — "The Cost of Data Breaches": three metric tiles + smoothed line chart on a white card.
function slide17 (pptx) {
  const s = pptx.addSlide()
  s.addShape('rect', { x: 10.2415, y: 0, w: 3.0921, h: 7.5, fill: { color: C.mint }, line: { type: 'none' } })
  petalWash(s, 5.9716, -0.4674, 4.4229, 4.4354, 311.11)
  chrome(s, 17, { headColor: C.white, numColor: C.white })
  headline(s, [
    { text: 'The Cost of ' }, { text: 'Data Breaches', color: C.blue },
  ], { x: 1.0425, y: 1.5865, w: 5.1456, h: 1.7484 })
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing lobortis sapien at efficitur. Sed  tellus vel enim vehicula pharetra non est.',
    { x: 1.0425, y: 3.349, w: 5.1456, h: 0.6085, lineSpacingMultiple: 1.3, paraSpaceAfter: 12 })

  const tiles = [
    { x: 1.0425, fill: C.blue, big: '+81', sup: '+', bigFg: C.white, pillFill: C.pill, pillFg: C.slate, label: 'Data 1', numX: 1.2136, numW: 1.4082 },
    { x: 3.1, fill: C.mint, big: '123', sup: '', bigFg: C.ink75, pillFill: C.blue, pillFg: C.white, label: 'Data 2', numX: 3.2718, numW: 1.4082 },
    { x: 5.1583, fill: C.blue, big: '10k', sup: 'k', bigFg: C.white, pillFill: C.periwinkle, pillFg: C.white, label: 'Data 3', numX: 5.1583, numW: 1.7519 },
  ]
  tiles.forEach(t => {
    roundRect(s, { x: t.x, y: 4.7212, w: 1.7661, h: 1.4294, fill: t.fill, r: 0.09932 })
    const runs = t.sup === '+'
      ? [{ text: '+', superscript: true }, { text: '81' }]
      : (t.sup === 'k' ? [{ text: '10' }, { text: 'k', superscript: true }] : [{ text: t.big }])
    text(s, runs.map(r => Object.assign({ fontSize: 40, fontFace: FONT_H, color: t.bigFg }, r)),
      { x: t.numX, y: t.x === 5.1583 ? 4.9328 : 4.9199, w: t.numW, h: 0.7743, align: 'center', margin: [0, 7.2, 3.6, 3.6], paraSpaceAfter: 3 })
    roundRect(s, { x: t.x + 0.4655, y: 5.6597, w: 0.8352, h: 0.2317, fill: t.pillFill, r: 0.5 })
    text(s, t.label, { x: t.x + 0.4655, y: 5.6597, w: 0.8352, h: 0.2317, fontSize: 9, color: t.pillFg, align: 'center', valign: 'middle' })
  })
  text(s, 'Data 3', { x: 5.6997, y: 5.7078, w: 0.6819, h: 0.303, fontSize: 12, color: C.ink50, align: 'center', margin: [0, 7.2, 3.6, 3.6] })

  roundRect(s, { x: 7.4655, y: 1.3496, w: 4.8247, h: 4.7999, fill: C.white, r: 0.04223 })
  s.addChart(pptx.ChartType.line, [
    { name: 'A', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [90, 50, 95, 70] },
    { name: 'B', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [40, 95, 60, 80] },
    { name: 'C', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [80, 70, 80, 40] },
  ], {
    x: 7.6294, y: 1.7349, w: 4.4981, h: 4.0292,
    chartColors: [C.periwinkle, C.blue, C.charcoal],
    lineSize: 3, lineSmooth: true, lineDataSymbol: 'none',
    catAxisHidden: true, valAxisHidden: true, valAxisMinVal: 40, valAxisMaxVal: 100,
    valGridLine: { color: C.wash, size: 0.75 }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'b', legendFontSize: 12, legendColor: '000000', legendFontFace: FONT_B,
    chartArea: { fill: { color: C.white } },
  })
}

// 18 — "Accurate Are Data Predictions": glyph list + smoothed monthly line chart.
function slide18 (pptx) {
  const s = pptx.addSlide()
  chrome(s, 18)
  headline(s, [
    { text: 'Accurate Are ' }, { text: 'Data Predictions', color: C.blue },
  ], { x: 1.2618, y: 1.7815, w: 5.9667, h: 1.7484, color: '000000' })

  const items = [
    { y: 3.8501, fill: C.blue, label: 'Lorem Ipsum Dolor Sit Amet' },
    { y: 4.3633, fill: C.mint, label: 'Cras At Tellus In Mauris Varius' },
    { y: 4.8765, fill: C.blue, label: 'Lacinia Sed Erat Velit, Porta Vitae' },
    { y: 5.3897, fill: C.mint, label: 'Convallis Gravida Imperdiet Tellus.' },
  ]
  items.forEach(it => {
    mark(s, PETAL_MARK, { x: 1.3369, y: it.y, w: 0.3856, h: 0.3867, fill: { color: it.fill } })
    text(s, it.label, { x: 1.8318, y: it.y, w: 3.6579, h: 0.3868, fontSize: 14, color: C.mute, lineSpacingMultiple: 1.3, paraSpaceAfter: 12 })
  })

  roundRect(s, { x: 7.3888, y: 1.4426, w: 4.6825, h: 4.6146, fill: C.mint, r: 0.03355 })
  s.addChart(pptx.ChartType.line, [{
    name: 'Series 1',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    values: [1.5, 4.6, 2.4, 5.4, 4.23, 5.13, 3.9, 6.5, 8.5, 7.32, 4.7, 8.56],
  }], {
    x: 7.6908, y: 1.8509, w: 4.0785, h: 3.7981,
    layout: { x: 0.1096, y: 0.0829, w: 0.8769, h: 0.7577 },
    chartColors: [C.blue], lineSize: 3.5, lineSmooth: true, lineDataSymbol: 'none',
    catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT_B, catAxisLabelColor: C.blue,
    catAxisLineColor: C.wash, catAxisMajorTickMark: 'none',
    valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT_B, valAxisLabelColor: C.mute,
    valAxisLabelFormatCode: '"$"#,##0.00', valAxisMajorUnit: 2, valAxisLineShow: false,
    valGridLine: { color: 'E8F8EE', size: 0.5 }, catGridLine: { style: 'none' },
    showLegend: false, chartArea: { fill: { type: 'none' } },
  })
}

// 19 — Q&A divider.
function slide19 (pptx) {
  const s = pptx.addSlide()
  petalWash(s, 5.2714, 0.2051, 7.7889, 7.811, 314.31)
  chrome(s, 19)
  text(s, 'Text for Sub Heading', { x: 1.5028, y: 2.0768, w: 2.9846, h: 0.3782, fontSize: 18, fontFace: FONT_H, color: C.slate, lineSpacingMultiple: 0.9 })
  headline(s, [
    { text: 'Let’s Dive Into ' }, { text: 'Your Questions', color: C.blue },
  ], { x: 1.5028, y: 2.6274, w: 4.8306, h: 2.5674 })
  body(s, LOREM_QA, { x: 1.5028, y: 5.5709, w: 4.2224, h: 0.5731 })
}

// 20 — Closing slide.
function slide20 (pptx) {
  const s = pptx.addSlide()
  mark(s, SPLIT_BLOB, { x: 0.1843, y: 0.6396, w: 6.1616, h: 6.1234, fill: BLOB, rotate: 335.9 })
  text(s, 'Data Science Presentation Template',
    { x: 6.0011, y: 1.4181, w: 5.4826, h: 0.4042, fontSize: 18, fontFace: FONT_H, color: C.slate })
  headline(s, [{ text: 'Thank You\nSo Much!', color: C.blueDark }],
    { x: 6.0011, y: 1.9202, w: 6.4517, h: 2.7858, fontSize: 88 })
  mark(s, PETAL_MARK, { x: 7.2657, y: 4.8033, w: 0.597, h: 0.5987, fill: { color: C.mint }, rotate: 45.16 })
  body(s, 'Lorem ipsum dolor sit amet, elit euismoodyantr amutie consectur.',
    { x: 8.0641, y: 4.8503, w: 2.8467, h: 0.5053, lineSpacingMultiple: 1 })
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------
function build () {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'WIDE', width: 13.3333, height: 7.5 })
  pptx.layout = 'WIDE'
  pptx.theme = { headFontFace: FONT_H, bodyFontFace: FONT_B }
  pptx.title = 'DataSphere Presentation Template'

  const builders = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  builders.forEach(fn => fn(pptx))

  return pptx.writeFile({ fileName: path.join(__dirname, '17a6a07b-d3c9-4fbb-b5b6-b8c3a443ac57_grok_final.pptx') })
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1) })
