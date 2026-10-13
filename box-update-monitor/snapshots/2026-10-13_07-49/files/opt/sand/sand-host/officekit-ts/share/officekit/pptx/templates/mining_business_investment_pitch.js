/**
 * "Mining Business" pitch deck - rebuilt with pptxgenjs.
 *
 * 16 slides, 13.333in x 7.5in.  Every position/size below is in inches and is
 * taken straight from the source deck.  Raster artwork in the original is
 * replaced by native shapes (icons) or flat placeholder blocks (photos).
 *
 * Run: node 02af2d7e-a4b0-45fb-9aa3-5faec8c4f944_grok_final.js
 */

const PptxGenJS = require('pptxgenjs')
const path = require('path')

/* ------------------------------------------------------------------ theme */

const ORANGE = 'E14D0D' // accent1
const WHITE = 'FFFFFF'
const INK = '0D0D0D' // tx1
const MIST = 'F2F2F2' // bg1 lumMod 95%
const PHOTO = 'D9D9D9' // stand-in for the photographs in the original
const PHONE = 'B87B8B' // average tint of the phone mock-up photo on slide 8

const HEAD = 'Raleway SemiBold' // theme major latin
const BODY = 'Poppins' // theme minor latin
const ALT = 'Urbanist SemiBold' // used on a few slides

// Shadow presets, translated from the source `outerShdw` elements.
const SOFT = { type: 'outer', blur: 42, offset: 0, angle: 0, color: '000000', opacity: 0.12 }
const CARD = { type: 'outer', blur: 35, offset: 10, angle: 50, color: '000000', opacity: 0.15 }
const LIFT = { type: 'outer', blur: 18, offset: 3, angle: 90, color: '000000', opacity: 0.11 }
const BAR = { type: 'outer', blur: 25, offset: 3, angle: 45, color: '000000', opacity: 0.15 }

/* ---------------------------------------------------------------- helpers */

// PowerPoint stores a roundRect corner radius as a fraction of the short side.
const corner = (adj, w, h) => (adj / 100000) * Math.min(w, h)

function tx (slide, text, opts) {
  // margin left unset so PowerPoint's default insets (0.1in / 0.05in) apply,
  // which is what every text box in the source deck uses.
  // fit:'resize' == PowerPoint's "resize shape to fit text", which every text
  // box in the source deck has switched on.
  slide.addText(text, Object.assign(
    { fontFace: BODY, fontSize: 12, color: INK, valign: 'top', isTextBox: true, fit: 'resize' },
    opts
  ))
}

function box (slide, opts) {
  const { shape = 'rect', shadow, ...rest } = opts
  // pptxgenjs rewrites the shadow object in place while emitting XML, so the
  // shared presets above must be copied before they are handed over.
  slide.addShape(shape, shadow ? { ...rest, shadow: { ...shadow } } : rest)
}

function rule (slide, x, y, len, color, width = 1, opts = {}) {
  slide.addShape('line', Object.assign({ x, y, w: len, h: 0, line: { color, width } }, opts))
}

// Stand-in for a photograph: a flat block tinted to the picture's average
// colour, captioned so it reads as a placeholder rather than real artwork.
function photo (slide, x, y, w, h, tint = PHOTO, rounding = 0) {
  box(slide, { shape: rounding ? 'roundRect' : 'rect', x, y, w, h, fill: { color: tint }, rectRadius: rounding })
  tx(slide, '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', color: WHITE, fontSize: 12 })
}

/* ------------------------------------------------- icons (native shapes) */

// Concentric-ring target. `cross` draws the crosshair ticks of the "Target"
// icon (bookend slides); otherwise a dart is added for the "Bullseye" icon.
function bullseye (slide, x, y, d, color, cross) {
  const ring = k => ({ x: x + (d * (1 - k)) / 2, y: y + (d * (1 - k)) / 2, w: d * k, h: d * k })
  const stroke = { color, width: d * 3.6 }
  box(slide, { shape: 'ellipse', ...ring(0.9), fill: { type: 'none' }, line: stroke })
  box(slide, { shape: 'ellipse', ...ring(0.52), fill: { type: 'none' }, line: stroke })
  box(slide, { shape: 'ellipse', ...ring(0.16), fill: { color } })
  if (cross) {
    rule(slide, x, y + d / 2, d, color, d * 3.6)
    slide.addShape('line', { x: x + d / 2, y, w: 0, h: d, line: stroke })
  } else {
    slide.addShape('line', {
      x: x + d * 0.5, y: y + d * 0.06, w: d * 0.44, h: d * 0.44, flipV: true,
      line: { color, width: d * 3.6, endArrowType: 'triangle' }
    })
  }
}

// Three interlocking discs, standing in for the "Venn diagram" graphic.
function venn (slide, x, y, d, color) {
  const r = d * 0.58
  const seats = [[(d - r) / 2, 0], [0, d - r], [d - r, d - r]]
  seats.forEach(([dx, dy]) => box(slide, {
    shape: 'ellipse', x: x + dx, y: y + dy, w: r, h: r, fill: { color, transparency: 25 }
  }))
}

// Filled disc with an arrow knocked out of it.
const ARROW = [[0, 0.38], [0.52, 0.38], [0.52, 0.13], [1, 0.5], [0.52, 0.87], [0.52, 0.62], [0, 0.62], 'close']

function arrowDisc (slide, x, y, d, disc, arrow) {
  box(slide, { shape: 'ellipse', x, y, w: d, h: d, fill: { color: disc } })
  freeform(slide, ARROW, x + d * 0.25, y + d * 0.32, d * 0.5, d * 0.36, { fill: { color: arrow } })
}

// Micro-pictograms for the five badges on the timeline slide (each ~0.3in).
const BADGES = {
  coins: (s, x, y, d, c) => {
    [0, 1, 2].forEach(i => box(s, { shape: 'can', x, y: y + d * (0.04 + i * 0.22), w: d * 0.56, h: d * 0.22, fill: { color: c } }))
    box(s, { shape: 'ellipse', x: x + d * 0.42, y: y + d * 0.42, w: d * 0.58, h: d * 0.58, fill: { color: c } })
  },
  percent: (s, x, y, d, c) => {
    box(s, { shape: 'ellipse', x, y, w: d * 0.34, h: d * 0.34, fill: { color: c } })
    box(s, { shape: 'ellipse', x: x + d * 0.66, y: y + d * 0.66, w: d * 0.34, h: d * 0.34, fill: { color: c } })
    s.addShape('line', { x: x + d * 0.14, y: y + d * 0.06, w: d * 0.72, h: d * 0.88, flipH: true, line: { color: c, width: d * 8 } })
  },
  chart: (s, x, y, d, c) => {
    [0.34, 0.56, 0.82].forEach((hh, i) => box(s, {
      x: x + d * (0.04 + i * 0.28), y: y + d * (1 - hh), w: d * 0.19, h: d * hh, fill: { color: c }
    }))
    s.addShape('line', { x: x + d * 0.1, y: y + d * 0.08, w: d * 0.85, h: d * 0.42, flipV: true, line: { color: c, width: d * 6, endArrowType: 'triangle' } })
  },
  people: (s, x, y, d, c) => {
    [[0, 0.8], [0.36, 1], [0.72, 0.8]].forEach(([fx, k]) => {
      const w = d * 0.28
      box(s, { shape: 'ellipse', x: x + d * fx, y: y + d * (0.32 - 0.28 * k), w, h: w, fill: { color: c } })
      box(s, { shape: 'roundRect', x: x + d * fx, y: y + d * 0.4, w, h: d * 0.58 * k, fill: { color: c }, rectRadius: d * 0.09 })
    })
  },
  // A cupped hand (half-disc) holding a rising bar chart and a coin.
  hand: (s, x, y, d, c) => {
    box(s, { shape: 'pie', x, y: y + d * 0.56, w: d, h: d * 0.88, fill: { color: c }, angleRange: [180, 360] })
    ;[0.22, 0.4].forEach((hh, i) => box(s, {
      x: x + d * (0.1 + i * 0.24), y: y + d * (0.54 - hh), w: d * 0.16, h: d * hh, fill: { color: c }
    }))
    box(s, { shape: 'ellipse', x: x + d * 0.58, y: y + d * 0.06, w: d * 0.34, h: d * 0.34, fill: { color: c } })
  }
}

/* ------------------------------------------------------------- freeforms */

// Points below are normalised to the shape's own box; `poly` scales them.
const poly = (pts, x, y, w, h) => pts.map(p =>
  p === 'close' ? { close: true }
    : p.length === 6
      ? { x: x + p[4] * w, y: y + p[5] * h, curve: { type: 'cubic', x1: x + p[0] * w, y1: y + p[1] * h, x2: x + p[2] * w, y2: y + p[3] * h } }
      : { x: x + p[0] * w, y: y + p[1] * h }
)

// Title-slide left tab: rounded panel with a circular bump on its right edge.
const SIDE_TAB = [
  [0, 0], [0.5855, 0],
  [0.6502, 0, 0.7027, 0.0128, 0.7027, 0.0286],
  [0.7027, 0.4648], [0.7547, 0.4623],
  [0.8902, 0.4623, 1, 0.4891, 1, 0.5221],
  [1, 0.5552, 0.8902, 0.582, 0.7547, 0.582],
  [0.7027, 0.5794], [0.7027, 0.9714],
  [0.7027, 0.9872, 0.6502, 1, 0.5855, 1],
  [0, 1], 'close'
]

// Half-disc used for the two centre panels on the "90+" slide.
const HALF_DISC = [
  [0.0009, 0],
  [0.5527, 0, 1, 0.2239, 1, 0.5],
  [1, 0.7762, 0.5527, 1, 0.0009, 1],
  [0, 1], [0, 0], 'close'
]

// Typographic quote mark (two slabs) used on the testimonial slides.
const QUOTE = [
  [0.7223, 0], [0.9482, 0], [0.8704, 0.4876], [0.9963, 0.4876], [1, 1], [0.6037, 1], [0.6, 0.4834], 'close',
  [0.1222, 0], [0.3482, 0], [0.2704, 0.4876], [0.3963, 0.4876], [0.4, 1], [0.0037, 1], [0, 0.4834], 'close'
]

function freeform (slide, pts, x, y, w, h, opts = {}) {
  box(slide, Object.assign({ shape: 'custGeom', x, y, w, h, points: poly(pts, 0, 0, w, h) }, opts))
}

/* ------------------------------------------------- master chrome (nav bar) */

const NAV = [['Home', 6.755], ['About Us', 8.459], ['Service', 10.162], ['Contact', 11.865]]

// `c.nav` is either one colour for all four links or a per-link array.
function chrome (slide, page, c = {}) {
  tx(slide, 'Mining Business Pitch', { x: 0.364, y: 0.307, w: 3.844, h: 0.303, fontFace: HEAD, color: c.brand || INK })
  NAV.forEach(([label, x], i) => tx(slide, label, {
    x, y: 0.241, w: 1.104, h: 0.345, align: 'center', lineSpacingMultiple: 1.3,
    fontFace: HEAD, color: (Array.isArray(c.nav) ? c.nav[i] : c.nav) || INK
  }))
  tx(slide, 'www.MiningBusiness.com', { x: 0.323, y: 6.956, w: 4.083, h: 0.303, color: c.www || INK, lineSpacingMultiple: 1 })
  // Slide 11 hides the page number behind a full-height panel.
  if (page) tx(slide, 'Page ' + page, { x: 11.375, y: 6.956, w: 1.635, h: 0.303, align: 'right', fontFace: HEAD, color: c.page || INK })
}

/* ------------------------------------------------------------ copy blocks */

// Slides 9 and 10 carry a bold Raleway list-style override on their title.
const SECTION_TITLE = { fontSize: 44, bold: true, align: 'center', fontFace: HEAD, lineSpacingMultiple: 0.8 }

const L = {
  hero: 'Lorem ipsum dolor sit amet, eget consectetuer adipiscing. Aenean commodo ligula dolor. Aenean ipsum dolor sit',
  short: 'Lorem ipsum dolor sit amet, eget consectetuer adipiscing. Aenean commodo ligula dolor. ',
  trim: 'Lorem ipsum dolor sit amet, eget consectetuer adipiscing. Aenean',
  elit: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
  elitLong: 'Lorem ipsum dolor sit amet, consectetuer dolor adipiscing elit. Aenean commodo ligula eget dolor. Aenean. ',
  sociis: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque',
  tiny: 'Lorem ipsum dolor sit amet, consectetuer'
}

/* ----------------------------------------------------------- slide 1 / 16 */

// Both bookend slides share the orange field, left tab and rotated side nav.
function bookend (pres, opts) {
  const s = pres.addSlide()
  s.background = { color: ORANGE }
  freeform(s, SIDE_TAB, 0, 0, 1.83, 7.5, { fill: { color: WHITE }, shadow: SOFT })
  tx(s, 'Pitch', {
    x: 2.033, y: 2.894, w: 11.897, h: 5.89, fontSize: 344, fontFace: HEAD, color: WHITE, transparency: 90
  })
  ;[['Home', 2.979, 'center'], ['About Us', 4.745, 'center'], ['Contact', 6.51, 'left']].forEach(([label, y, align]) =>
    tx(s, label, { x: 0.091, y, w: 1.104, h: 0.345, rotate: 270, align, lineSpacingMultiple: 1.3, fontFace: HEAD })
  )
  tx(s, 'Mining Business', { x: -0.194, y: 0.951, w: 1.673, h: 0.303, rotate: 270, align: 'right', fontFace: HEAD })
  bullseye(s, 1.136, 3.67, 0.491, INK, true)

  box(s, {
    shape: 'roundRect', x: opts.pillX, y: opts.pillY, w: 3.897, h: 0.465,
    fill: { color: WHITE }, rectRadius: corner(50000, 3.897, 0.465), shadow: SOFT
  })
  rule(s, opts.lineX, opts.pillY + 0.233, opts.lineW, WHITE, 1, { line: { color: WHITE, width: 1, endArrowType: 'oval' } })
  tx(s, 'TEMPLATE PRESENTATION', {
    x: opts.pillX + 0.094, y: opts.pillY + 0.065, w: 3.71, h: 0.337,
    fontSize: 14, align: 'center', charSpacing: 3, fontFace: HEAD
  })
  return s
}

function slide01 (pres) {
  const s = bookend(pres, { pillX: 2.686, pillY: 1.778, lineX: 7.028, lineW: 4.044 })
  tx(s, 'Mining Business', { x: 2.506, y: 2.456, w: 9.371, h: 1.447, fontSize: 80, fontFace: HEAD, color: WHITE })
  tx(s, L.hero, { x: 2.506, y: 3.99, w: 5.218, h: 0.608, color: WHITE, lineSpacingMultiple: 1.3 })
}

function slide16 (pres) {
  const s = bookend(pres, { pillX: 2.48, pillY: 1.493, lineX: 6.822, lineW: 1.542 })
  tx(s, 'Thank You!!', { x: 2.3, y: 2.17, w: 6.83, h: 1.447, fontSize: 80, fontFace: HEAD, color: WHITE })
  ;[['+214-704-76532', 5.469, 1.691], ['thesisone@gmail.com', 5.88, 2.089], ['www.companyinfo.com', 6.252, 2.231]]
    .forEach(([t, y, w]) => tx(s, t, { x: 10.284, y, w, h: 0.345, color: WHITE, lineSpacingMultiple: 1.3 }))
}

/* --------------------------------------------------------------- slide 2 */

function slide02 (pres) {
  const s = pres.addSlide()
  chrome(s, 2)
  box(s, {
    shape: 'roundRect', x: 0.534, y: 3.15, w: 3.246, h: 3.362,
    fill: { color: ORANGE }, rectRadius: corner(3234, 3.246, 3.362), shadow: SOFT
  })
  tx(s, 'Investment Opportunity', { x: 0.534, y: 1.147, w: 9.759, h: 1.01, fontSize: 54, fontFace: HEAD })
  tx(s, 'Introducing Our Mining Vision', { x: 9.701, y: 3.034, w: 3.098, h: 0.337, fontSize: 14, align: 'right', fontFace: HEAD })
  tx(s, '+68M', { x: 0.534, y: 2.447, w: 1.101, h: 0.575, fontSize: 24, fontFace: HEAD, lineSpacingMultiple: 1.3 })
  tx(s, 'Market Potential', { x: 0.771, y: 3.458, w: 2.246, h: 0.37, fontSize: 16, fontFace: HEAD, color: WHITE })
  rule(s, 0.816, 3.954, 2.676, WHITE)
  tx(s, [
    { text: 'Lorem ipsum dolor sit amet, eget ' },
    { text: 'consectetuer adipiscing. Aenean sit', options: { bold: true, italic: true, underline: { style: 'sng' } } },
    { text: ' commodo ligula dolor. Aenean ipsum dolor sit' }
  ], { x: 0.771, y: 5.071, w: 2.772, h: 1.133, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3 })
  tx(s, 'Lorem ipsum dolor sit amet', { x: 1.58, y: 2.646, w: 2.655, h: 0.304, fontSize: 10, lineSpacingMultiple: 1.3 })
  box(s, { shape: 'roundRect', x: 4.841, y: 2.314, w: 1.414, h: 0.088, fill: { color: ORANGE }, rectRadius: corner(50000, 1.414, 0.088) })
  arrowDisc(s, 3.14, 3.514, 0.259, WHITE, ORANGE)
}

/* --------------------------------------------------------------- slide 3 */

function slide03 (pres) {
  const s = pres.addSlide()
  s.background = { color: ORANGE }
  box(s, { x: 8.224, y: 0, w: 5.109, h: 7.5, fill: { color: WHITE }, shadow: SOFT })
  chrome(s, 3, { brand: WHITE, nav: [WHITE, INK, INK, INK], www: WHITE })

  // Right-hand headline column.
  box(s, { shape: 'roundRect', x: 9.188, y: 1.688, w: 0.645, h: 0.088, fill: { color: ORANGE }, rectRadius: corner(50000, 0.645, 0.088) })
  tx(s, 'Mining Location', { x: 9.128, y: 1.844, w: 3.476, h: 1.919, fontSize: 54, fontFace: HEAD })
  tx(s, 'Strategic Sites with High Resource Potential', { x: 9.128, y: 4.288, w: 2.597, h: 0.572, fontSize: 14, fontFace: HEAD })
  rule(s, 9.188, 4.979, 2.676, ORANGE)
  tx(s, L.trim, { x: 9.128, y: 5.205, w: 3.138, h: 0.608, lineSpacingMultiple: 1.3 })

  // Two stacked white cards on the orange field.
  const cards = [
    { y: 0.965, icon: bullseye, iconY: 1.242, ruleY: 1.724, titleY: 1.286, bodyY: 2.527, arrowY: 1.325,
      body: 'Lorem ipsum dolor sit amet, eget sit consectetuer ligula adipiscing. Aenean amet' },
    { y: 3.825, icon: venn, iconY: 4.158, ruleY: 4.628, titleY: 4.201, bodyY: 5.387, arrowY: 4.24,
      body: 'Lorem ipsum dolor sit amet, eget consectetuer sit ligula adipiscing. Aenean amet' }
  ]
  cards.forEach(c => {
    box(s, {
      shape: 'roundRect', x: 0.935, y: c.y, w: 3.153, h: 2.709,
      fill: { color: WHITE }, rectRadius: corner(4773, 3.153, 2.709), shadow: SOFT
    })
    c.icon(s, 1.174, c.iconY, 0.423, ORANGE)
    rule(s, 1.174, c.ruleY, 2.676, INK)
    tx(s, 'Your Title Here', { x: 1.666, y: c.titleY, w: 1.691, h: 0.337, fontSize: 14, fontFace: HEAD })
    tx(s, c.body, { x: 1.174, y: c.bodyY, w: 2.623, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 })
    arrowDisc(s, 3.578, c.arrowY, 0.259, ORANGE, WHITE)
  })
}

/* --------------------------------------------------------------- slide 4 */

function slide04 (pres) {
  const s = pres.addSlide()
  box(s, { x: 0, y: 3.639, w: 13.333, h: 3.861, fill: { color: ORANGE }, shadow: SOFT })
  chrome(s, 4, { www: WHITE, page: WHITE })

  tx(s, 'Mineral Reserves Resource Quality', { x: 0.81, y: 4.349, w: 7.23, h: 1.919, fontSize: 54, fontFace: HEAD, color: WHITE })
  ;[4.737, 5.742].forEach(y => {
    arrowDisc(s, 7.639, y, 0.401, WHITE, ORANGE)
    tx(s, L.short, { x: 8.251, y, w: 4.272, h: 0.608, color: WHITE, lineSpacingMultiple: 1.3 })
  })

  box(s, {
    shape: 'roundRect', x: 0.81, y: 1.084, w: 3.21, h: 2.314,
    fill: { color: ORANGE }, rectRadius: corner(5200, 3.21, 2.314), shadow: SOFT
  })
  bullseye(s, 1.03, 1.358, 0.615, WHITE)
  tx(s, '1200M+', { x: 1.726, y: 1.339, w: 1.588, h: 0.654, fontSize: 28, fontFace: HEAD, color: WHITE, lineSpacingMultiple: 1.3 })
  tx(s, 'Lorem ipsum dolor sit amet, eget consectetuer adipiscing. ', {
    x: 1.03, y: 2.536, w: 2.772, h: 0.608, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3,
    bold: true, italic: true, underline: { style: 'sng' }
  })
  box(s, { shape: 'roundRect', x: 0.95, y: 4.262, w: 1.414, h: 0.088, fill: { color: WHITE }, rectRadius: corner(50000, 1.414, 0.088) })
}

/* --------------------------------------------------------------- slide 5 */

function slide05 (pres) {
  const s = pres.addSlide()
  chrome(s, 5)
  tx(s, 'Mining Operations Production Process', { x: 0.896, y: 1.202, w: 7.593, h: 1.919, fontSize: 54, fontFace: HEAD })
  box(s, { shape: 'roundRect', x: 1.044, y: 1.107, w: 1.414, h: 0.088, fill: { color: ORANGE }, rectRadius: corner(50000, 1.414, 0.088) })

  box(s, {
    shape: 'roundRect', x: 5.854, y: 3.424, w: 2.542, h: 3.26,
    fill: { color: ORANGE }, rectRadius: corner(6068, 2.542, 3.26), shadow: SOFT
  })
  tx(s, '89+', { x: 6.107, y: 3.691, w: 1.247, h: 0.812, fontSize: 36, fontFace: HEAD, color: WHITE, lineSpacingMultiple: 1.3 })
  tx(s, 'Efficient, Safe, and Scalable Extraction Methods', {
    x: 6.107, y: 4.596, w: 2.143, h: 0.99, fontSize: 14, fontFace: HEAD, color: WHITE,
    underline: { style: 'sng' }, lineSpacingMultiple: 1.3
  })
  rule(s, 6.194, 6.312, 1.861, WHITE)

  tx(s, 'Proven Assets Backed by Geological Data', { x: 0.896, y: 4.218, w: 4.17, h: 0.378, fontSize: 14, fontFace: HEAD, lineSpacingMultiple: 1.3 })
  tx(s, 'Lorem ipsum dolor sit amet, eget consectetuer adipiscing. Aenean commodo ligula dolor. Aenean ipsum dolor sit ipsum dolor sit ipsum', {
    x: 0.896, y: 4.625, w: 4.17, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3
  })
  tx(s, L.short, { x: 0.896, y: 5.586, w: 4.17, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 })
}

/* --------------------------------------------------------------- slide 6 */

function slide06 (pres) {
  const s = pres.addSlide()
  chrome(s, 6)
  tx(s, 'Market Demand', { x: 3.157, y: 1.162, w: 7.02, h: 1.01, fontSize: 54, align: 'center', fontFace: HEAD })
  box(s, { shape: 'roundRect', x: 5.914, y: 1.074, w: 1.505, h: 0.088, fill: { color: ORANGE }, rectRadius: corner(50000, 1.505, 0.088) })
  tx(s, L.short, { x: 4.582, y: 2.195, w: 4.17, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 })

  const tiles = [
    { x: 0.915, label: 'Environmental ', lx: 1.712, lw: 2.114 },
    { x: 4.812, label: 'Governance Commitment', lx: 5.253, lw: 2.828 },
    { x: 8.71, label: 'Environmental ', lx: 9.507, lw: 2.114 }
  ]
  tiles.forEach(t => box(s, {
    shape: 'roundRect', x: t.x, y: 3.197, w: 3.708, h: 3.229,
    fill: { color: ORANGE }, rectRadius: corner(4409, 3.708, 3.229), shadow: SOFT
  }))
  tiles.forEach(t => tx(s, t.label, {
    x: t.lx, y: 5.92, w: t.lw, h: 0.337, fontSize: 14, align: 'center', fontFace: HEAD, color: WHITE
  }))
}

/* --------------------------------------------------------------- slide 7 */

function slide07 (pres) {
  const s = pres.addSlide()
  chrome(s, 7)
  // Highlight blocks sit behind the two knocked-out phrases in the pull quote.
  box(s, { shape: 'roundRect', x: 3.815, y: 4.395, w: 2.493, h: 0.597, fill: { color: ORANGE }, rectRadius: corner(16667, 2.493, 0.597), shadow: SOFT })
  box(s, { shape: 'roundRect', x: 7.255, y: 2.943, w: 4.416, h: 0.597, fill: { color: ORANGE }, rectRadius: corner(16667, 4.416, 0.597), shadow: SOFT })

  tx(s, [
    { text: 'Experienced ' },
    { text: 'Management', options: { color: ORANGE } }
  ], { x: 3.425, y: 1.162, w: 6.484, h: 0.64, fontSize: 32, align: 'center', fontFace: HEAD })
  box(s, { shape: 'roundRect', x: 6.667, y: 1.074, w: 0.732, h: 0.088, fill: { color: ORANGE }, rectRadius: corner(50000, 0.732, 0.088) })

  tx(s, [
    { text: 'A profitable mining ' },
    { text: 'business starts', options: { color: WHITE, breakLine: true } },
    { text: 'with reliable resources and ' },
    { text: 'efficient', options: { color: WHITE } },
    { text: ' operations' }
  ], { x: 1.547, y: 2.79, w: 10.239, h: 2.322, fontSize: 44, align: 'center', fontFace: HEAD })
  tx(s, 'Industry Experts Driving Operational Excellence', { x: 4.582, y: 2.519, w: 4.17, h: 0.345, align: 'center', lineSpacingMultiple: 1.3 })

  tx(s, 'Mark Medison', { x: 6.206, y: 5.388, w: 1.688, h: 0.37, fontSize: 16, fontFace: HEAD })
  tx(s, 'Your Position', { x: 6.244, y: 5.666, w: 1.397, h: 0.345, italic: true, lineSpacingMultiple: 1.3 })

  freeform(s, QUOTE, 1.239, 2.059, 0.548, 0.491, { fill: { color: INK } })
  freeform(s, QUOTE, 11.546, 5.004, 0.548, 0.491, { fill: { color: ORANGE }, flipH: true, flipV: true })
}

/* --------------------------------------------------------------- slide 8 */

function slide08 (pres) {
  const s = pres.addSlide()
  box(s, { x: 0, y: 0, w: 5.26, h: 7.5, fill: { color: ORANGE }, shadow: SOFT })
  chrome(s, 8, { brand: WHITE, www: WHITE })

  tx(s, 'Equipment Overview', { x: 0.651, y: 1.591, w: 4.375, h: 1.919, fontSize: 54, fontFace: HEAD, color: WHITE })
  tx(s, 'Modern Assets Supporting Operational Efficiency', { x: 0.651, y: 4.121, w: 2.737, h: 0.572, fontSize: 14, fontFace: HEAD, color: WHITE })
  rule(s, 0.711, 4.812, 2.676, WHITE)
  tx(s, 'Lorem ipsum dolor Aenean sit amet, eget consectetuer Aenean adipiscing. Aenean commodo ligula dolor. Aenean ipsum', {
    x: 0.651, y: 5.038, w: 3.755, h: 0.87, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3
  })

  // Two phone mock-up photos in the original; PHONE is their average tint.
  photo(s, 6.174, 0.736, 3.097, 6.028, PHONE, 0.3)
  photo(s, 9.414, 0.736, 3.097, 6.028, PHONE, 0.3)
}

/* --------------------------------------------------------------- slide 9 */

function slide09 (pres) {
  const s = pres.addSlide()
  chrome(s, 9)
  tx(s, 'Section Infographic', { x: 2.652, y: 1.067, w: 7.787, h: 0.693, ...SECTION_TITLE })
  tx(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus Lorem ipsum dolor sit ame sociis', {
    x: 2.704, y: 1.829, w: 7.683, h: 0.608, align: 'center', lineSpacingMultiple: 1.3
  })

  // Date card.
  box(s, { x: 0.839, y: 2.93, w: 3.329, h: 3.541, fill: { color: MIST } })
  tx(s, 'Tuesday', { x: 1.946, y: 3.166, w: 1.114, h: 0.37, fontSize: 16, align: 'center', fontFace: HEAD, wrap: false })
  tx(s, '03', { x: 1.588, y: 3.813, w: 1.832, h: 1.784, fontSize: 100, align: 'center', fontFace: HEAD, wrap: false })
  tx(s, 'June / 2025', { x: 1.75, y: 5.865, w: 1.506, h: 0.37, fontSize: 16, align: 'center', fontFace: HEAD, wrap: false })

  // Highlighted second row.
  box(s, { x: 3.892, y: 4.191, w: 7.637, h: 1.019, fill: { color: ORANGE }, shadow: SOFT })

  const rows = [
    { n: '01', nx: 4.698, ny: 3.019, nw: 0.886, tx: 5.706, ty: 3.089, tw: 6.507, color: INK,
      body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa cum adispacing et' },
    { n: '02', nx: 4.055, ny: 4.28, nw: 0.915, tx: 5.141, ty: 4.351, tw: 6.008, color: WHITE,
      body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing et elit. Aenean commodo ligula eget dolor. Aenean massa cum et ' },
    { n: '03', nx: 4.702, ny: 5.549, nw: 0.919, tx: 5.81, ty: 5.62, tw: 6.403, color: INK,
      body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing et elit cum Aenean commodo ligula eget dolor. Aenean massa adispacin' }
  ]
  rows.forEach(r => {
    tx(s, r.n, { x: r.nx, y: r.ny, w: r.nw, h: 0.841, fontSize: 44, fontFace: HEAD, color: r.color, wrap: false })
    tx(s, r.body, { x: r.tx, y: r.ty, w: r.tw, h: 0.692, fontSize: 14, color: r.color, lineSpacingMultiple: 1.3 })
  })
}

/* -------------------------------------------------------------- slide 10 */

function slide10 (pres) {
  const s = pres.addSlide()
  chrome(s, 10)
  tx(s, 'Section Infographic', { x: 2.773, y: 1.456, w: 7.787, h: 0.693, ...SECTION_TITLE })
  tx(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus Lorem ipsum dolor sit', {
    x: 2.825, y: 2.288, w: 7.683, h: 0.608, align: 'center', lineSpacingMultiple: 1.3
  })

  // Connector track behind the four milestone cards.
  ;[2.604, 5.744, 8.767].forEach(x => rule(s, x, 4.229, 1.958, INK, 1.5))

  const steps = [
    { x: 0.683, dot: 2.129, on: false, year: '2025', yx: 0.857, yw: 1.219, label: 'Project One', lw: 1.448 },
    { x: 3.706, dot: 5.053, on: true, year: '2026', yx: 3.885, yw: 1.243, label: 'Project Two', lw: 1.49 },
    { x: 6.728, dot: 8.143, on: false, year: '2027', yx: 6.893, yw: 1.217, label: 'Project Three', lw: 1.638 },
    { x: 9.75, dot: 11.174, on: true, year: '2028', yx: 9.93, yw: 1.242, label: 'Project Four', lw: 1.524 }
  ]
  steps.forEach(t => {
    box(s, {
      x: t.x, y: 4.604, w: 2.9, h: 1.522,
      fill: { color: t.on ? ORANGE : MIST }, ...(t.on ? { shadow: SOFT } : {})
    })
    box(s, t.on
      ? { shape: 'ellipse', x: t.dot, y: 4.13, w: 0.198, h: 0.198, fill: { color: ORANGE }, shadow: SOFT }
      : { shape: 'ellipse', x: t.dot, y: 4.13, w: 0.198, h: 0.198, fill: { color: WHITE }, line: { color: INK, width: 0.75, transparency: 50 }, shadow: SOFT })
    const color = t.on ? WHITE : INK
    tx(s, t.year, { x: t.yx, y: 4.724, w: t.yw, h: 0.64, fontSize: 32, fontFace: HEAD, color, wrap: false })
    tx(s, t.label, { x: t.yx, y: 5.585, w: t.lw, h: 0.37, fontSize: 16, fontFace: HEAD, color, wrap: false })
  })
}

/* -------------------------------------------------------------- slide 11 */

function slide11 (pres) {
  const s = pres.addSlide()
  box(s, { x: 7.862, y: 0, w: 5.471, h: 7.5, fill: { color: ORANGE }, shadow: SOFT })
  chrome(s, 0, { nav: [INK, WHITE, WHITE, WHITE] })

  // Three nested block arcs, largest first so the smaller ones stay visible.
  const arcs = [
    { x: 1.397, y: 1.685, d: 5.461, range: [268.528, 1.98817], thick: 0.24066, pct: '25%', px: 1.498, py: 1.219, ty: 1.921 },
    { x: 2.18, y: 2.468, d: 3.896, range: [178.673, 1.19755], thick: 0.32128, pct: '55%', px: 1.38, py: 4.706, ty: 5.407 },
    { x: 3.003, y: 3.29, d: 2.25, range: [89.602, 0.2616], thick: 0.46672, pct: '75%', px: 4.571, py: 4.535, ty: 5.237 }
  ]
  // Source draws the small arc + hub first, then the larger rings on top.
  const order = [2, 1, 0]
  order.forEach(i => {
    const a = arcs[i]
    box(s, { shape: 'blockArc', x: a.x, y: a.y, w: a.d, h: a.d, fill: { color: ORANGE }, angleRange: a.range, arcThicknessRatio: a.thick })
    if (i === 2) box(s, { shape: 'ellipse', x: 3.722, y: 4.009, w: 0.812, h: 0.812, fill: { color: ORANGE } })
  })
  arcs.forEach(a => {
    tx(s, a.pct, { x: a.px, y: a.py, w: 1.861, h: 0.707, fontSize: 36, fontFace: HEAD })
    tx(s, L.tiny, { x: a.px, y: a.ty, w: 2.187, h: 0.565, fontSize: 11, lineSpacingMultiple: 1.3 })
  })

  tx(s, 'Section Infographic', { x: 8.287, y: 1.576, w: 4.554, h: 1.42, fontSize: 48, fontFace: HEAD, color: WHITE, lineSpacingMultiple: 0.8 })
  tx(s, '500+', { x: 8.44, y: 4.084, w: 2.433, h: 0.977, fontSize: 44, fontFace: HEAD, color: WHITE, lineSpacingMultiple: 1.3 })
  tx(s, L.sociis, { x: 8.44, y: 5.054, w: 4.144, h: 0.87, color: WHITE, align: 'justify', lineSpacingMultiple: 1.3 })
}

/* -------------------------------------------------------------- slide 12 */

function slide12 (pres) {
  const s = pres.addSlide()
  chrome(s, 12)
  tx(s, 'Section Infographic', { x: 2.257, y: 1.322, w: 8.82, h: 0.747, fontSize: 48, align: 'center', fontFace: HEAD, lineSpacingMultiple: 0.8 })

  // Centre "split disc": grey right half, orange left half (mirrored).
  freeform(s, HALF_DISC, 6.73, 2.424, 2.113, 4.223, { fill: { color: MIST } })
  freeform(s, HALF_DISC, 4.49, 2.424, 2.113, 4.223, { fill: { color: ORANGE }, flipH: true, shadow: CARD })

  // Outer stat cards.
  const outer = [
    { x: 0.984, adj: 6826, fill: ORANGE, color: WHITE, tx: 1.216, nx: 1.365, lx: 1.565 },
    { x: 9.446, adj: 7178, fill: MIST, color: INK, tx: 9.677, nx: 9.826, lx: 10.026 }
  ]
  outer.forEach(c => {
    box(s, {
      shape: 'roundRect', x: c.x, y: 2.644, w: 2.903, h: 3.782,
      fill: { color: c.fill }, rectRadius: corner(c.adj, 2.903, 3.782), ...(c.fill === ORANGE ? { shadow: CARD } : {})
    })
    tx(s, '90+', { x: c.nx, y: 3.053, w: 2.142, h: 0.909, fontSize: 48, align: 'center', fontFace: ALT, color: c.color })
    tx(s, 'Your Tile Here', { x: c.lx, y: 3.867, w: 1.742, h: 0.37, fontSize: 16, align: 'center', fontFace: ALT, color: c.color })
    tx(s, L.elitLong, { x: c.tx, y: 4.541, w: 2.44, h: 1.395, align: 'center', color: c.color, lineSpacingMultiple: 1.3 })
  })

  // Inner labels, mirrored around the split.
  tx(s, 'Title Here 01', { x: 5.075, y: 4.395, w: 1.35, h: 0.357, bold: true, fontFace: HEAD, color: WHITE, align: 'right', valign: 'bottom', lineSpacingMultiple: 1.3 })
  tx(s, 'Title Here 02', { x: 6.931, y: 4.413, w: 1.35, h: 0.34, bold: true, fontFace: HEAD, valign: 'bottom', lineSpacingMultiple: 1.3 })
  tx(s, L.elit, { x: 4.898, y: 4.762, w: 1.528, h: 1.047, fontSize: 11, align: 'right', color: WHITE, lineSpacingMultiple: 1.3 })
  tx(s, L.elit, { x: 6.931, y: 4.762, w: 1.528, h: 1.047, fontSize: 11, align: 'left', color: INK, lineSpacingMultiple: 1.3 })

  box(s, { shape: 'roundRect', x: 5.545, y: 3.261, w: 0.82, h: 0.766, fill: { color: WHITE }, rectRadius: corner(8258, 0.82, 0.766), shadow: CARD })
  box(s, { shape: 'roundRect', x: 6.968, y: 3.261, w: 0.82, h: 0.766, fill: { color: ORANGE }, rectRadius: corner(8258, 0.82, 0.766), shadow: LIFT })
  venn(s, 5.696, 3.386, 0.518, ORANGE)
  bullseye(s, 7.119, 3.386, 0.518, WHITE)
}

/* -------------------------------------------------------------- slide 13 */

function slide13 (pres) {
  const s = pres.addSlide()
  chrome(s, 13)
  tx(s, 'Section Infographic', { x: 1.023, y: 1.104, w: 11.287, h: 0.828, fontSize: 48, align: 'center', fontFace: HEAD, lineSpacingMultiple: 0.9 })
  tx(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus dolor sit amet, ', {
    x: 2.976, y: 1.997, w: 7.382, h: 0.608, align: 'center', lineSpacingMultiple: 1.3
  })

  const cols = [
    { x: 0.917, on: true, year: '2024', yx: 1.104, bx: 1.087, bw: 2.08, badge: 1.244, ix: 1.381, icon: 'coins' },
    { x: 3.25, on: false, year: '2025', yx: 3.438, bx: 3.462, bw: 1.981, badge: 3.648, ix: 3.785, icon: 'percent' },
    { x: 5.583, on: true, year: '2026', yx: 5.75, bx: 5.778, bw: 2.149, badge: 5.867, ix: 6.004, icon: 'chart' },
    { x: 7.917, on: false, year: '2027', yx: 8.104, bx: 8.09, bw: 2.149, badge: 8.194, ix: 8.33, icon: 'people' },
    { x: 10.25, on: true, year: '2028', yx: 10.438, bx: 10.434, bw: 2.149, badge: 10.517, ix: 10.653, icon: 'hand' }
  ]
  cols.forEach(c => {
    const color = c.on ? WHITE : INK
    box(s, { x: c.x, y: 3.264, w: 2.208, h: 3.249, fill: { color: c.on ? ORANGE : MIST } })
    // The year boxes inherit a 12pt space-before from their list style.
    tx(s, c.year, { x: c.yx, y: 3.532, w: 1.833, h: 0.74, fontSize: 32, fontFace: HEAD, color, lineSpacingMultiple: 1.3, paraSpaceBefore: 12 })
    tx(s, L.elit, { x: c.bx, y: 4.323, w: c.bw, h: 0.87, color, lineSpacingMultiple: 1.3 })
    // Badge is filled with the opposite colour so the pictogram reads.
    box(s, { shape: 'ellipse', x: c.badge, y: 5.664, w: 0.591, h: 0.591, fill: { color: c.on ? MIST : ORANGE } })
    BADGES[c.icon](s, c.ix, 5.801, 0.317, c.on ? ORANGE : WHITE)
  })
}

/* -------------------------------------------------------------- slide 14 */

function slide14 (pres) {
  const s = pres.addSlide()
  chrome(s, 14)

  const bars = [
    { x: 0.406, y: 5.915, h: 1.585, on: false, pct: '49%', px: 0.677, py: 6.018, pw: 1.441, align: 'center' },
    { x: 2.514, y: 4.826, h: 2.674, on: true, pct: '67%', px: 2.732, py: 4.934, pw: 1.547, align: 'center' },
    { x: 4.621, y: 5.389, h: 2.111, on: false, pct: '58%', px: 4.908, py: 5.516, pw: 1.427, align: 'center' },
    { x: 6.729, y: 3.532, h: 3.968, on: true, pct: '86%', px: 6.924, py: 3.678, pw: 1.594, align: 'left' },
    { x: 8.837, y: 2.496, h: 5.004, on: true, pct: '99%', px: 9.04, py: 2.601, pw: 1.576, align: 'left' },
    { x: 10.945, y: 4.105, h: 3.395, on: false, pct: '74%', px: 11.218, py: 4.207, pw: 1.436, align: 'center' }
  ]
  bars.forEach(b => box(s, {
    x: b.x, y: b.y, w: 1.983, h: b.h, fill: { color: b.on ? ORANGE : MIST }, ...(b.on ? { shadow: BAR } : {})
  }))

  tx(s, 'Selection Infographic', { x: 0.717, y: 1.393, w: 8.491, h: 0.828, fontSize: 48, fontFace: HEAD, lineSpacingMultiple: 0.9 })
  tx(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis', {
    x: 0.717, y: 2.226, w: 5.618, h: 0.608, lineSpacingMultiple: 1.3
  })

  bars.forEach(b => tx(s, b.pct, {
    x: b.px, y: b.py, w: b.pw, h: 0.841, fontSize: 44, fontFace: HEAD,
    color: b.on ? WHITE : INK, align: b.align, wrap: false
  }))
}

/* -------------------------------------------------------------- slide 15 */

function slide15 (pres) {
  const s = pres.addSlide()
  chrome(s, 15)
  tx(s, 'Customer Testimonials', { x: 2.947, y: 1.112, w: 7.438, h: 0.693, fontSize: 44, bold: true, align: 'center', fontFace: HEAD, lineSpacingMultiple: 0.8 })

  const cards = [
    { x: 1.091, adj: 4099, fill: WHITE, color: INK, pillFill: ORANGE, pillText: WHITE, disc: WHITE, arrow: ORANGE,
      quoteX: 1.445, bar: 1.393, barColor: ORANGE, headX: 1.335, headW: 2.711, nameX: 1.445, pillX: 2.803, learnX: 2.918, iconX: 3.942 },
    { x: 4.885, adj: 3806, fill: ORANGE, color: WHITE, pillFill: WHITE, pillText: INK, disc: ORANGE, arrow: WHITE,
      quoteX: 5.238, bar: 5.167, barColor: WHITE, headX: 5.129, headW: 2.551, nameX: 5.238, pillX: 6.595, learnX: 6.71, iconX: 7.734 },
    { x: 8.678, adj: 3806, fill: WHITE, color: INK, pillFill: ORANGE, pillText: WHITE, disc: WHITE, arrow: ORANGE,
      quoteX: 9.032, bar: 8.94, barColor: ORANGE, headX: 8.913, headW: 2.688, nameX: 9.032, pillX: 10.447, learnX: 10.563, iconX: 11.601 }
  ]
  cards.forEach(c => {
    box(s, {
      shape: 'roundRect', x: c.x, y: 2.193, w: 3.564, h: 4.389,
      fill: { color: c.fill }, rectRadius: corner(c.adj, 3.564, 4.389), shadow: LIFT
    })
    box(s, {
      shape: 'roundRect', x: c.pillX, y: 2.5, w: 1.567, h: 0.474,
      fill: { color: c.pillFill }, rectRadius: corner(50000, 1.567, 0.474), shadow: LIFT
    })
    tx(s, 'Learn Now', { x: c.learnX, y: 2.585, w: 1.024, h: 0.303, align: 'center', fontFace: ALT, color: c.pillText })
    arrowDisc(s, c.iconX, 2.585, 0.303, c.disc, c.arrow)
    freeform(s, QUOTE, c.quoteX, 3.397, 0.394, 0.353, { fill: { color: c.color } })
    tx(s, 'Profit starts with reliable resources', { x: c.headX, y: 3.945, w: c.headW, h: 1.313, fontSize: 24, fontFace: HEAD, color: c.color })
    slide15Rule(s, c)
    tx(s, 'Deni Sakira', { x: c.nameX, y: 5.555, w: 1.391, h: 0.37, fontSize: 16, fontFace: HEAD, color: c.color, wrap: false })
    tx(s, 'Job Position', { x: c.nameX, y: 5.925, w: 1.215, h: 0.303, italic: true, color: c.color, wrap: false })
  })
}

function slide15Rule (slide, c) {
  slide.addShape('line', { x: c.bar, y: 5.59, w: 0, h: 0.578, line: { color: c.barColor, width: 1.25 } })
}

/* -------------------------------------------------------------------- go */

function build () {
  const pres = new PptxGenJS()
  pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 })
  pres.layout = 'WIDE'
  pres.author = 'Mining Business'
  pres.title = 'Mining Business Pitch'

  ;[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
    .forEach(fn => fn(pres))

  return pres.writeFile({ fileName: path.join(__dirname, '02af2d7e-a4b0-45fb-9aa3-5faec8c4f944_grok_final.pptx') })
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1) })
