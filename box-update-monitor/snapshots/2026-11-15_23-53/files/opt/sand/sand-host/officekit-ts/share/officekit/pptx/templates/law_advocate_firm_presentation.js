/*
 * Advocatez - Law Advocate Presentation
 * 24 slides, 13.3333 x 7.5 in (16:9).
 *
 * Standalone pptxgenjs recreation of the reference deck.
 *   node 089b79f3-61c4-4f54-a7fe-60b9653295d5_grok_final.js
 *
 * The original deck contains no embedded raster images - every picture frame is
 * an unfilled placeholder that renders as plain white - so no image stand-ins
 * are required here.
 */

const path = require('path')
const PptxGenJS = require('pptxgenjs')

// ---------------------------------------------------------------- palette & type
const NAVY = '293040' // theme accent1
const GOLD = 'D3AC2B' // theme accent2
const WHITE = 'FFFFFF'
const GREY = '808080' // tx1 at 50% luminance - every paragraph of body copy
const HEAD = 'Poppins' // theme major latin
const BODY = 'Open Sans' // theme minor latin

const NONE = { type: 'none' }

// Each white card carries the same very soft, wide drop shadow. pptxgenjs
// mutates the shadow object it is handed, so hand out a fresh copy every time.
const cardShadow = () => ({ type: 'outer', blur: 15, offset: 0, angle: 90, color: '000000', opacity: 0.16 })

// ---------------------------------------------------------------- geometry

const KAPPA = 0.5523 // cubic-bezier control ratio approximating a quarter circle

// Point list (shape-local inches) for a rectangle with rounded corners.
// radii = [topLeft, topRight, bottomRight, bottomLeft].
function roundedRectPoints (w, h, [tl, tr, br, bl]) {
  const pts = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }]
  if (tr) pts.push({ x: w, y: tr, curve: { type: 'cubic', x1: w - tr * (1 - KAPPA), y1: 0, x2: w, y2: tr * (1 - KAPPA) } })
  pts.push({ x: w, y: h - br })
  if (br) pts.push({ x: w - br, y: h, curve: { type: 'cubic', x1: w, y1: h - br * (1 - KAPPA), x2: w - br * (1 - KAPPA), y2: h } })
  pts.push({ x: bl, y: h })
  if (bl) pts.push({ x: 0, y: h - bl, curve: { type: 'cubic', x1: bl * (1 - KAPPA), y1: h, x2: 0, y2: h - bl * (1 - KAPPA) } })
  pts.push({ x: 0, y: tl })
  if (tl) pts.push({ x: tl, y: 0, curve: { type: 'cubic', x1: 0, y1: tl * (1 - KAPPA), x2: tl * (1 - KAPPA), y2: 0 } })
  pts.push({ close: true })
  return pts
}

// ---------------------------------------------------------------- panels

// Plain white card with the deck's signature soft shadow.
function card (s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: WHITE }, line: NONE, shadow: cardShadow() })
}

// White card rounded on one side only. `side` is 'l' or 'r'.
function roundCard (s, x, y, w, h, r, side) {
  const radii = side === 'r' ? [0, r, r, 0] : [r, 0, 0, r]
  s.addShape('custGeom', {
    x, y, w, h, points: roundedRectPoints(w, h, radii),
    fill: { color: WHITE }, line: NONE, shadow: cardShadow()
  })
}

// White card rounded on all four corners (slide 19's tall portrait panel).
function pillCard (s, x, y, w, h, r) {
  s.addShape('custGeom', {
    x, y, w, h, points: roundedRectPoints(w, h, [r, r, r, r]),
    fill: { color: WHITE }, line: NONE, shadow: cardShadow()
  })
}

// The "ticket" tab used around the frame: a bar with two rounded ends and a
// white half-circle bitten out of one side. x/y/w/h describe the tab as it
// appears on the slide; `side` ('l' or 'r') is the rounded/bitten edge.
function ticket (s, x, y, w, h, color, side) {
  // round2SameRect rounds its top corners, so the bar is authored upright and
  // spun a quarter turn to put those corners on the requested edge.
  s.addShape('round2SameRect', {
    x: x + w / 2 - h / 2, y: y + h / 2 - w / 2, w: h, h: w,
    rotate: side === 'r' ? 90 : 270, fill: { color }, line: NONE
  })
  const d = h * 0.4555 // notch diameter, proportional to the tab thickness
  const cx = side === 'r' ? x + w : x
  s.addShape('ellipse', { x: cx - d / 2, y: y + h / 2 - d / 2, w: d, h: d, fill: { color: WHITE }, line: NONE })
}

// ---------------------------------------------------------------- accents

// Gold mark: a large right triangle with two quarter-size ones tucked beside
// it, reading as a square sliced by a diagonal. Placement is expressed as
// fractions of the mark's bounding box so it can be scaled and spun freely.
const MARK_PARTS = [
  { fx: 0, fy: 0, f: 1, rot: 90 },
  { fx: 0.5599, fy: 0.1244, f: 0.4401, rot: 270 },
  { fx: 0.1198, fy: 0.5634, f: 0.4401, rot: 270 }
]

// `spin` is 0 or 90 - roughly half the slides use the quarter-turned version.
function mark (s, x, y, size, spin) {
  MARK_PARTS.forEach(part => {
    let { fx, fy, rot } = part
    const f = part.f
    if (spin === 90) { // rotate the whole group a quarter turn about its centre
      const dx = fx + f / 2 - 0.5
      const dy = fy + f / 2 - 0.5
      fx = 0.5 - dy - f / 2
      fy = 0.5 + dx - f / 2
      rot = (rot + 90) % 360
    }
    s.addShape('rtTriangle', {
      x: x + fx * size, y: y + fy * size, w: f * size, h: f * size,
      rotate: rot, fill: { color: GOLD }, line: NONE
    })
  })
}

// Thin gold rule ending in a triangular head - the deck's list marker.
function arrow (s, x, y, w) {
  s.addShape('line', { x, y, w: w || 0.4476, h: 0, line: { color: GOLD, width: 0.5, endArrowType: 'triangle' } })
}

// Navy pill button, always the same size and always reading "Learn More".
function learnMore (s, x, y) {
  s.addShape('roundRect', { x, y, w: 1.5528, h: 0.4045, rectRadius: 0.1, fill: { color: NAVY }, line: NONE })
  s.addText('Learn More', {
    x, y, w: 1.5528, h: 0.4045,
    fontFace: BODY, fontSize: 11, color: WHITE, charSpacing: 1.5, align: 'center', valign: 'middle'
  })
}

// Small solid dot (slide 2 puts one gold one on the edge of the wide card).
function dot (s, x, y, color) {
  s.addShape('ellipse', { x, y, w: 0.253, h: 0.253, fill: { color }, line: NONE })
}

// Filled circle carrying a short bold label ("01", "$15", ...).
function badge (s, x, y, fill, text, color) {
  const d = 1.0086
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, line: NONE })
  s.addText(text, { x, y, w: d, h: d, fontFace: HEAD, fontSize: 24, bold: true, color, align: 'center', valign: 'middle' })
}

// ---------------------------------------------------------------- text

// Two-tone section title: navy phrase followed by a gold one.
function heading (s, x, y, w, h, dark, light, size, align) {
  s.addText([
    { text: dark, options: { color: NAVY } },
    { text: light, options: { color: GOLD } }
  ], {
    x, y, w, h, fontFace: HEAD, fontSize: size, bold: true, align: align || 'left', valign: 'top'
  })
}

// Letter-spaced grey strapline sitting under the two cover titles.
function tagline (s, x, y, w, text) {
  s.addText(text, { x, y, w, h: 0.3366, fontFace: BODY, fontSize: 14, color: GREY, charSpacing: 3, valign: 'top', wrap: false })
}

// Bold navy 14pt caption.
function label (s, x, y, w, text, align) {
  s.addText(text, { x, y, w, h: 0.3366, fontFace: HEAD, fontSize: 14, bold: true, color: NAVY, align: align || 'left', valign: 'top', wrap: false })
}

// Grey 11pt paragraph at 1.5 line spacing - used for every block of copy.
function body (s, x, y, w, h, text, align) {
  s.addText(text, { x, y, w, h, fontFace: BODY, fontSize: 11, color: GREY, lineSpacingMultiple: 1.5, align: align || 'left', valign: 'top' })
}

// Oversized bold figure such as "90%" or a list index.
function stat (s, x, y, w, text, color, align) {
  s.addText(text, { x, y, w, h: 0.505, fontFace: HEAD, fontSize: 24, bold: true, color, align: align || 'left', valign: 'top', wrap: false })
}

// Caption + trailing arrow + paragraph: the deck's repeating list item.
function feature (s, x, y, title, arrowX, text, textW, textH, opts) {
  const o = opts || {}
  label(s, x, y, 2.25, title, o.align)
  arrow(s, arrowX, y + 0.176, o.arrowW)
  body(s, x, y + 0.3345, textW, textH, text, o.align)
}

// ---------------------------------------------------------------- page frame

const NAV_ITEMS = [
  { text: 'About', x: 7.9236, w: 0.6367 },
  { text: 'Service', x: 9.3298, w: 0.7103 },
  { text: 'Team', x: 10.8096, w: 0.6034 },
  { text: 'Portfolio', x: 12.1825, w: 0.8208 }
]
const NAV_DIVIDERS = [8.945, 10.4249, 11.7977]

// Furniture repeated on all 24 slides: navy nav bar top right, brand tab top
// left, gold website tab bottom left and page-number tab bottom right.
// `logoX` shifts the brand tab (slide 5 centres it); `pageR` is the right edge
// of the page label, which a few slides pull inward.
function chrome (s, num, logoX, pageR) {
  const right = pageR || 12.8724

  // top-right navy navigation bar
  ticket(s, 7.4671, 0.2743, 5.8663, 0.5556, NAVY, 'l')
  NAV_ITEMS.forEach(it => s.addText(it.text, {
    x: it.x, y: 0.409, w: it.w, h: 0.2861,
    fontFace: BODY, fontSize: 11, color: WHITE, align: 'right', valign: 'top', wrap: false
  }))
  NAV_DIVIDERS.forEach(x => s.addShape('line', { x, y: 0.4199, w: 0, h: 0.2596, line: { color: WHITE, width: 0.5 } }))

  // top-left brand tab
  ticket(s, logoX, 0.4216, 0.2579, 0.2579, NAVY, 'r')
  s.addText('Advocatez', {
    x: logoX + 0.4609, y: 0.409, w: 1.0487, h: 0.2861,
    fontFace: HEAD, fontSize: 11, bold: true, color: NAVY, valign: 'top', wrap: false
  })

  // bottom-left gold website tab
  ticket(s, 0, 6.6701, 2.6354, 0.5556, GOLD, 'r')
  s.addText('Your Website Here', {
    x: 0.1886, y: 6.8049, w: 2.2583, h: 0.2861,
    fontFace: BODY, fontSize: 11, color: WHITE, charSpacing: 3, align: 'center', valign: 'top', wrap: false
  })

  // bottom-right page-number tab
  ticket(s, right + 0.203, 6.8175, 0.2579, 0.2579, NAVY, 'l')
  s.addText('Page ' + String(num).padStart(2, '0'), {
    x: right - 1.2, y: 6.8049, w: 1.2, h: 0.2861,
    fontFace: HEAD, fontSize: 11, bold: true, color: NAVY, align: 'right', valign: 'top', wrap: false
  })
}

// ---------------------------------------------------------------- slides

function slide01 (s) {
  chrome(s, 1, 0)
  card(s, 0, 1.6032, 8.9206, 4.2857)
  heading(s, 1.0964, 2.4723, 4.8213, 1.1107, 'Advoca', 'tez', 60)
  tagline(s, 1.0964, 3.4148, 3.7344, 'Law Advocate Presentation')
  body(s, 1.0964, 4.1232, 5.5368, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation commodo consequat. ')
  mark(s, 0.2579, 1.8632, 0.4458)
  arrow(s, 0.2562, 5.6351)
}

function slide02 (s) {
  chrome(s, 2, 0)
  heading(s, 5.705, 1.7034, 3.5766, 0.5722, 'Table Of ', 'Content', 28)
  body(s, 5.705, 2.4442, 6.6667, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  card(s, 0.2579, 4.3121, 12.5211, 1.8246)
  feature(s, 0.7484, 4.7439, 'About', 2.3858, 'Lorem ipsum dolor sit amet, consectetur adipiscing', 2.2498, 0.6268)
  feature(s, 3.2951, 4.7439, 'Service', 4.9326, 'Lorem ipsum dolor sit amet, consectetur adipiscing', 2.2498, 0.6268)
  feature(s, 5.8419, 4.7439, 'Team', 7.4794, 'Lorem ipsum dolor sit amet, consectetur adipiscing', 2.2498, 0.6268)
  feature(s, 8.3887, 4.7439, 'Portfolio', 10.0261, 'Lorem ipsum dolor sit amet, consectetur adipiscing', 2.2498, 0.6268)
  dot(s, 0.1314, 5.0979, 'D3AC2B')
}

function slide03 (s) {
  chrome(s, 3, 0)
  card(s, 0, 1.303, 6.2032, 2.0464)
  heading(s, 9.708, 1.9834, 2.3824, 1.0434, 'Welcome ', 'Message', 28)
  body(s, 9.708, 3.3495, 2.6743, 1.7375, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt laboreet dolore magna aliqua. Ut enim aminim veniam, quis nostrud exercitation ullamco laboris nisi aliquip')
  learnMore(s, 9.822, 5.5969)
  feature(s, 1.3735, 1.7068, 'Happy Client', 3.0109, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  stat(s, 0.4105, 2.0738, 0.8909, '90%', 'D3AC2B')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
}

function slide04 (s) {
  chrome(s, 4, 0)
  card(s, 7.1301, 1.6614, 5.054, 2.0464)
  heading(s, 1.1492, 5.2981, 2.3504, 1.0434, 'About Our  ', 'Company', 28)
  learnMore(s, 10.6308, 5.6175)
  body(s, 4.0317, 5.2287, 6, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo aute dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  feature(s, 9.4761, 2.0636, 'Happy Client', 11.1135, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  stat(s, 8.5131, 2.4306, 0.8909, '90%', 'D3AC2B')
}

function slide05 (s) {
  chrome(s, 5, 4.6061)
  card(s, 4.6061, 1.5671, 8.7273, 2.1829)
  heading(s, 5.3101, 2.1368, 3.169, 1.0434, 'Our Company  ', 'History', 28)
  body(s, 8.8388, 2.0675, 3.8891, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim venia, quis nostrud exercitation ullamco laboris')
  body(s, 6.0423, 4.2992, 2.685, 0.9044, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore')
  stat(s, 5.3101, 4.499, 0.5385, '01', '293040')
  mark(s, 4.7973, 1.7594, 0.4458)
  body(s, 10.0429, 4.2992, 2.685, 0.9044, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore')
  stat(s, 9.3107, 4.499, 0.6227, '03', '293040')
  body(s, 6.0423, 5.4533, 2.685, 0.9044, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore')
  stat(s, 5.3101, 5.6531, 0.6104, '02', '293040')
  body(s, 10.0429, 5.4533, 2.685, 0.9044, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore')
  stat(s, 9.3107, 5.6531, 0.6455, '04', '293040')
}

function slide06 (s) {
  chrome(s, 6, 0)
  card(s, 1.2701, 3.8246, 7.6505, 2.0534)
  heading(s, 1.1492, 1.4547, 2.8564, 1.0434, 'The Best Law ', 'Advocate', 28)
  learnMore(s, 10.8183, 1.7742)
  body(s, 4.3817, 1.3854, 6, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo aute dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  mark(s, 1.4349, 3.9908, 0.281)
  feature(s, 2.7234, 4.2319, 'Happy Client', 6.1364, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed eiusmod tempor incididunt labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud', 4.3882, 0.9044, { arrowW: 0.8104 })
  stat(s, 1.7604, 4.5989, 0.8909, '90%', 'D3AC2B')
}

function slide07 (s) {
  chrome(s, 7, 0)
  card(s, 7.0831, 1.8875, 6.2502, 2.6047)
  heading(s, 1.1492, 1.6736, 2.8356, 1.0434, 'Our Vision ', 'and Mission', 28)
  body(s, 1.1492, 3.0397, 3.1235, 1.7375, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim aminim veniam, quis nostrud exercitation ullamco laboris nisut aliquip ex ea commodo consequat. ')
  learnMore(s, 1.2632, 5.2871)
  label(s, 9.2308, 3.4895, 1.3835, 'Our Mission')
  body(s, 10.6913, 3.2056, 2.3253, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed eiusmod tempor')
  label(s, 9.2308, 2.5536, 1.238, 'Our Vision')
  body(s, 10.6913, 2.2696, 2.3253, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed eiusmod tempor')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
}

function slide08 (s) {
  chrome(s, 8, 0, 7.5033)
  card(s, 2.127, 4.0663, 6.0066, 1.8967)
  heading(s, 1.1492, 1.2881, 3.2144, 0.5722, 'More About ', 'Us', 28)
  body(s, 1.1492, 2.0125, 11.4114, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  feature(s, 4.1957, 4.5341, 'Happy Client', 6.739, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  stat(s, 3.2327, 4.7622, 0.8909, '90%', 'D3AC2B')
}

function slide09 (s) {
  chrome(s, 9, 0)
  card(s, 0, 1.4982, 3.9597, 1.2388)
  heading(s, 0.3726, 1.8241, 3.2144, 0.5722, 'What We ', 'Do?', 28, 'center')
  feature(s, 6.228, 1.4982, 'Consultation', 7.8654, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  badge(s, 4.9784, 1.6133, 'D3AC2B', '01', 'FFFFFF')
  feature(s, 10.2881, 1.4982, 'Legal Aid', 11.9255, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  badge(s, 9.0385, 1.6133, 'D3AC2B', '02', 'FFFFFF')
  body(s, 9.9025, 3.7589, 2.6354, 2.2929, 'Lorem ipsum dolorsit consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit')
  mark(s, 0.1886, 1.6733, 0.2723)
}

function slide10 (s) {
  chrome(s, 10, 0)
  card(s, 1.5096, 1.174, 3.5942, 3.5664)
  heading(s, 1.0787, 5.7252, 3.2144, 0.5722, 'Our ', 'Services', 28)
  feature(s, 2.3342, 1.62, 'Consultation', 3.9717, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  feature(s, 2.3342, 3.0556, 'Legal Aid', 3.9717, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  badge(s, 1.0053, 1.7351, '293040', '01', 'D3AC2B')
  badge(s, 1.0053, 3.1707, '293040', '02', 'D3AC2B')
  body(s, 4.3537, 5.4199, 5.9064, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum')
  learnMore(s, 10.7013, 5.8087)
  mark(s, 12.4852, 5.1438, 0.4458, 90)
}

function slide11 (s) {
  chrome(s, 11, 0, 9.1413)
  heading(s, 1.1492, 1.6736, 3.1235, 1.0434, 'Our Excellent ', 'Services', 28)
  card(s, 5.8728, 3.8936, 3.457, 1.8702)
  card(s, 5.8728, 1.7363, 3.457, 1.8702)
  feature(s, 6.7225, 2.052, 'Consultation', 8.3599, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  feature(s, 6.7225, 4.2092, 'Legal Aid', 8.3599, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  badge(s, 5.3685, 2.1671, 'D3AC2B', '01', '293040')
  badge(s, 5.3685, 4.3244, 'D3AC2B', '02', '293040')
  body(s, 1.1492, 3.0397, 3.1235, 1.7375, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim aminim veniam, quis nostrud exercitation ullamco laboris nisut aliquip ex ea commodo consequat. ')
  learnMore(s, 1.2632, 5.2871)
}

function slide12 (s) {
  chrome(s, 12, 0)
  card(s, 3.4848, 1.4298, 9.8485, 4.6405)
  heading(s, 4.4808, 2.3376, 2.6354, 1.0434, 'Service We ', 'Provide', 28)
  feature(s, 5.9156, 4.1318, 'Consultation', 7.553, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  feature(s, 10.2249, 4.1318, 'Legal Aid', 11.8623, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  badge(s, 4.5867, 4.2469, '293040', '01', 'D3AC2B')
  badge(s, 8.8959, 4.2469, '293040', '02', 'D3AC2B')
  body(s, 7.4745, 2.1294, 5.0001, 1.4598, 'Lorem ipsum dolor sit amet, consectetur adipiscing seddo eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit voluptate velit esse cillum dolore eu fugiat nulla pariatur. ')
  mark(s, 3.7428, 1.6897, 0.4458)
}

function slide13 (s) {
  chrome(s, 13, 0)
  card(s, 5.6032, 1.6032, 7.7302, 4.2857)
  heading(s, 6.9685, 2.4723, 5.0106, 1.1107, 'Break ', 'Slide', 60)
  tagline(s, 6.9685, 3.4148, 3.7344, 'Law Advocate Presentation')
  body(s, 6.9685, 4.1232, 5.5368, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation commodo consequat. ')
  arrow(s, 12.6287, 5.6351)
  mark(s, 12.6278, 1.8632, 0.4458, 90)
}

function slide14 (s) {
  chrome(s, 14, 0)
  heading(s, 1.1727, 1.2749, 2.6354, 1.0434, 'Meet Our ', 'Best Team', 28)
  body(s, 1.1727, 2.4793, 3.1784, 0.6268, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod')
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  roundCard(s, 7.0392, 3.8542, 6.2941, 2.2784, 0.3797, 'l')
  feature(s, 7.6976, 4.374, 'Sally Howell', 9.3351, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  feature(s, 10.4082, 4.374, 'David Ferreira', 12.0456, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
}

function slide15 (s) {
  chrome(s, 15, 0)
  card(s, 6.6667, 1.8998, 6.6667, 2.0061)
  heading(s, 1.1492, 5.4107, 3.7162, 1.0434, 'Our Professional ', 'Team', 28)
  learnMore(s, 10.6291, 5.7301)
  body(s, 5.381, 5.3413, 4.725, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt labore dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ')
  feature(s, 7.9488, 2.2834, 'Erika Cobb', 9.5862, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  feature(s, 10.5469, 2.2834, 'Michael Hogue', 12.1843, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044)
  mark(s, 0.6601, 1.1903, 0.4458)
}

function slide16 (s) {
  chrome(s, 16, 0)
  card(s, 1.3207, 1.2304, 6.6029, 2.0061)
  heading(s, 1.2331, 4.1288, 3.0577, 0.5722, 'David ', 'Ferreira', 28)
  body(s, 1.2331, 4.8357, 5.9336, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse')
  feature(s, 2.8256, 1.7529, 'Happy Client', 6.8734, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod tempor incididunt ut labore ', 4.5967, 0.6268)
  stat(s, 1.8626, 1.981, 0.8909, '90%', 'D3AC2B')
  mark(s, 1.5505, 1.4615, 0.3484)
}

function slide17 (s) {
  chrome(s, 17, 0)
  roundCard(s, 3.5969, 1.2784, 8.8923, 4.9432, 0.8239, 'r')
  heading(s, 7.0447, 1.9539, 2.6142, 0.5722, 'Sally ', 'Howell', 28)
  feature(s, 9.3651, 4.3073, 'Public Speaking', 11.2552, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 2.2498, 0.9044, { arrowW: 0.1949 })
  learnMore(s, 7.1487, 4.7244)
  body(s, 7.0447, 2.7662, 4.5702, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sedo eiusmod tempor incididu labore dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ')
  mark(s, 0.6601, 1.1903, 0.4458)
}

function slide18 (s) {
  chrome(s, 18, 0, 8.811)
  card(s, 3.3652, 3.8251, 6.6029, 2.0061)
  heading(s, 1.2441, 1.8624, 3.0542, 0.5722, 'Our Best ', 'Work', 28)
  body(s, 4.6878, 1.5574, 7.4947, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisi aliquip commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore fugiat sint occaecat cupidatat non officia deserunt mollit anim id est laborum.')
  mark(s, 0.6601, 1.1903, 0.4458)
  feature(s, 4.8022, 4.3476, 'Our Best Work', 7.3456, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  learnMore(s, 8.4952, 4.6258)
}

function slide19 (s) {
  chrome(s, 19, 0)
  heading(s, 5.6117, 4.839, 6.5182, 0.5722, 'Highlighting Our ', 'Notable Cases', 28)
  body(s, 5.6117, 5.4962, 6.6667, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ')
  pillCard(s, 1.3637, 1.2341, 3.1663, 5.0318, 0.4193)
  learnMore(s, 2.1702, 5.5057)
  label(s, 2.1946, 4.283, 1.5045, 'Happy Client', 'center')
  body(s, 1.5463, 4.6174, 2.8012, 0.6268, 'Lorem ipsum dolor sit consectetur adipiscing elites sed do', 'center')
  stat(s, 2.5014, 3.7831, 0.8909, '90%', 'D3AC2B', 'center')
  mark(s, 12.4852, 4.2812, 0.4458, 90)
}

function slide20 (s) {
  chrome(s, 20, 0)
  heading(s, 1.1492, 1.6736, 3.408, 1.0434, 'Legal Expertise ', 'in Action', 28)
  body(s, 1.1492, 3.0397, 3.1235, 1.7375, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim aminim veniam, quis nostrud exercitation ullamco laboris nisut aliquip ex ea commodo consequat. ')
  learnMore(s, 1.2632, 5.2871)
  mark(s, 0.6601, 1.1903, 0.4458)
}

function slide21 (s) {
  chrome(s, 21, 0)
  card(s, 3.6908, 1.6263, 8.7273, 2.1829)
  body(s, 7.9236, 2.1267, 3.8891, 1.1821, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim venia, quis nostrud exercitation ullamco laboris')
  mark(s, 3.882, 1.8186, 0.4458)
  heading(s, 4.3948, 2.196, 3.408, 1.0434, 'Proven Track ', 'Record', 28)
  mark(s, 12.4852, 1.1903, 0.4458, 90)
  feature(s, 9.1571, 4.4162, 'Our Best Work', 11.7005, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  feature(s, 9.1571, 5.5194, 'Our Best Work', 11.7005, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
}

function slide22 (s) {
  chrome(s, 22, 0)
  learnMore(s, 9.22, 5.3545)
  body(s, 9.1059, 3.107, 3.1235, 1.7375, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore dolore magna aliqua. Ut enim aminim veniam, quis nostrud exercitation ullamco laboris nisut aliquip ex ea commodo consequat. ')
  heading(s, 9.1059, 1.7409, 3.2052, 1.0434, 'Our Excellent ', 'Pricing Plan', 28)
  roundCard(s, 1.7173, 2.9443, 5.3519, 1.6115, 0.2686, 'l')
  roundCard(s, 1.7173, 1.1359, 5.3519, 1.6115, 0.2686, 'l')
  roundCard(s, 1.7173, 4.7585, 5.3519, 1.6115, 0.2686, 'l')
  badge(s, 1.2084, 3.2457, 'D3AC2B', '$25', '293040')
  badge(s, 1.2084, 1.4373, 'D3AC2B', '$15', '293040')
  badge(s, 1.2084, 5.0599, 'D3AC2B', '$35', '293040')
  learnMore(s, 6.2969, 3.5477)
  learnMore(s, 6.2969, 1.7393)
  learnMore(s, 6.2969, 5.3619)
  feature(s, 2.7259, 3.2694, 'Package Two', 5.2692, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  feature(s, 2.7259, 1.461, 'Package One', 5.2692, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  feature(s, 2.7259, 5.0836, 'Package Three', 5.2692, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  mark(s, 12.4852, 1.1903, 0.4458, 90)
}

function slide23 (s) {
  chrome(s, 23, 0)
  body(s, 1.255, 2.4608, 3.0902, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit seddo eiusmod incididunt ut labore et dolore magna aliqua. ')
  heading(s, 1.255, 1.7115, 3.2052, 0.5722, 'Contact ', 'Us', 28)
  feature(s, 1.255, 3.7238, 'Phone Number', 3.7984, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  feature(s, 1.255, 4.8274, 'Our Address', 3.7984, 'Lorem ipsum dolor sit amet, consectetur adipiscing elites sed do eiusmod ', 3.0922, 0.6268)
  mark(s, 0.6601, 1.1903, 0.4458)
}

function slide24 (s) {
  chrome(s, 24, 0)
  card(s, 4.4394, 1.6032, 8.8939, 4.2857)
  heading(s, 6.9685, 2.4723, 4.7477, 1.1107, 'Thank ', 'You', 60)
  tagline(s, 6.9685, 3.4148, 4.3129, 'For Watching This Presentation')
  body(s, 6.9685, 4.1232, 5.5368, 0.9044, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation commodo consequat. ')
  arrow(s, 12.6287, 5.6351)
  mark(s, 12.6278, 1.8632, 0.4458, 90)
}

// ---------------------------------------------------------------- assemble

const pptx = new PptxGenJS()
pptx.defineLayout({ name: 'WIDE', width: 13.3333333, height: 7.5 }) // 12192000 EMU exactly
pptx.layout = 'WIDE'
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY }

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24
]

BUILDERS.forEach(build => {
  const s = pptx.addSlide()
  s.background = { color: WHITE }
  build(s)
})

pptx.writeFile({ fileName: path.join(__dirname, '089b79f3-61c4-4f54-a7fe-60b9653295d5_grok_final.pptx') })
  .then(f => console.log('wrote ' + f))
