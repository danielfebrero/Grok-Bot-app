/**
 * Investment Pitch Deck — 25 slides, 13.333 x 7.5 in (16:9), dark theme.
 *
 * Rebuilt with PptxGenJS.  `node <this file>` writes the .pptx beside itself.
 * The deck's photographs are empty picture frames in the source, so they are
 * drawn here as flat placeholder panels (see `imageBox`); everything else is
 * native PptxGenJS geometry.
 */
'use strict'

const path = require('path')
const PptxGenJS = require('pptxgenjs')

/* ------------------------------------------------------------------ palette */
const W = 'FFFFFF' // theme bg2 — white
const K = '000000' // theme tx1 — the slide stage
const P = '6C36FE' // theme accent1 — electric violet
const HEAD = 'Sora SemiBold' // theme major font
const BODY = ' Plus Jakarta Sans' // theme minor font

const NOLINE = { type: 'none' }
const GLASS_VIOLET = 59 // opacity % of the violet stop in every frosted card

/* ---------------------------------------------------------------- utilities */
const hex = c => [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)]
const mix = (a, b, t) => hex(a)
  .map((v, i) => v + (hex(b)[i] - v) * Math.max(0, Math.min(1, t)))
  .map(v => Math.round(v).toString(16).toUpperCase().padStart(2, '0'))
  .join('')

/* -------------------------------------------------------- geometry helpers */

/** Rounded rectangle. `r` = corner radius in inches, `clear` = transparency %. */
function rr (s, x, y, w, h, r, color, clear) {
  s.addShape(r > 0 ? 'roundRect' : 'rect', {
    x, y, w, h, rectRadius: r || undefined,
    fill: { color, transparency: clear || 0 }, line: NOLINE
  })
}

/** Ellipse. */
function el (s, x, y, w, h, color, clear) {
  s.addShape('ellipse', { x, y, w, h, fill: { color, transparency: clear || 0 }, line: NOLINE })
}

/** Hairline rule / connector; a zero width or height gives a straight line. */
function ln (s, x, y, w, h, o) {
  o = o || {}
  s.addShape('line', {
    x, y, w, h,
    line: { color: o.color || W, transparency: o.alpha || 0, width: o.width || 0.25 }
  })
}

/** Small filled circle used as a rule end-cap. */
function tick (s, cx, cy, color) {
  el(s, cx - 0.0405, cy - 0.0405, 0.081, 0.081, color || W)
}

/**
 * Paint a linear ramp inside a rounded rectangle.  PptxGenJS has no gradient
 * fill, so the shape is stacked from nested rounded rectangles that all share
 * the bright edge; each is a little shorter and a little brighter, so the
 * silhouette stays exact while the colour walks along the ramp.  The last band
 * shrinks to a `2r` cap, never to a sliver.  `at(f)` gives the fill at f (0..1).
 */
function ramp (s, x, y, w, h, r, at, bands) {
  const vert = h >= w
  const len = vert ? h : w
  const cap = Math.min(len, 2 * r || 0.02)
  const n = bands || Math.max(6, Math.min(64, Math.round(len * 14)))
  for (let i = 0; i < n; i++) {
    const size = len - (len - cap) * i / (n - 1)
    s.addShape(r > 0 ? 'roundRect' : 'rect', {
      x: vert ? x : x + w - size,
      y,
      w: vert ? w : size,
      h: vert ? size : h,
      rectRadius: r || undefined,
      fill: at(i / (n - 1)),
      line: NOLINE
    })
  }
}

/**
 * Every gradient in the deck runs at the same 310deg angle: bright violet at
 * the top-right corner, black at the bottom-left.  `axis` collapses that
 * diagonal onto the shape's long side and returns the gradient position (0..1)
 * of the band at `f` (0 = dim end of that side, 1 = bright end).
 */
const COS = Math.cos(50 * Math.PI / 180)
const SIN = Math.sin(50 * Math.PI / 180)
function axis (w, h, f) {
  const long = h >= w ? SIN * h : COS * w
  const short = h >= w ? COS * w : SIN * h
  return (long * f + short / 2) / (long + short)
}

/** Opaque violet->black ramp: buttons, pills, chart bars, phone bodies. */
function gradRect (s, x, y, w, h, r, hi, lo, stops) {
  const [p0, p1] = stops || [0, 1]
  ramp(s, x, y, w, h, r, f => ({ color: mix(lo || K, hi || P, (axis(w, h, f) - p0) / (p1 - p0)) }))
}

/**
 * Frosted-glass card: a white wash (`wash`% opaque) brightening into violet
 * (59% opaque) along the ramp.  Butting translucent bands would leave seams,
 * so the blend against the backdrop is resolved here instead: `behind` is the
 * colour the card sits on — the black stage unless it overlaps a phone body.
 */
function glass (s, x, y, w, h, r, wash, behind) {
  const lo = wash === undefined ? 7 : wash
  ramp(s, x, y, w, h, r, f => {
    const g = axis(w, h, f)
    return { color: mix(behind || K, mix(W, P, g), (lo + (GLASS_VIOLET - lo) * g) / 100) }
  })
}

/**
 * Soft radial corner glow.  The source is one radial gradient fading from a
 * 64%-opaque violet core to fully clear at the rim; here it is nested
 * translucent discs, painted outside-in.  Every disc also covers the ones
 * inside it, so each one's own alpha is backed out of the cumulative alpha the
 * glow needs at that radius — which keeps the glow see-through, as artwork
 * regularly sits underneath it.
 */
const ORB_CORE = '2A00A4' // tint at the centre
const ORB_RIM = '3D1A93' // tint at the rim
const ORB_ALPHA = [0.78, 0.47] // cumulative opacity, centre -> rim
function orb (s, x, y, d) {
  const n = 10
  let covered = 0
  for (let i = n - 1; i >= 0; i--) {
    const f = i / (n - 1) // 0 = centre, 1 = rim
    const want = ORB_ALPHA[1] + (ORB_ALPHA[0] - ORB_ALPHA[1]) * (1 - f)
    const band = 1 - (1 - want) / (1 - covered)
    covered = want
    const size = d * (i + 1) / n
    el(s, x + (d - size) / 2, y + (d - size) / 2, size, size, mix(ORB_CORE, ORB_RIM, f), 100 - 100 * band)
  }
}

/** Concentric "radar" rings centred on (cx, cy) — a recurring backdrop motif. */
function rings (s, cx, cy) {
  const sizes = [1.865, 1.561, 1.287, 0.997, 0.705]
  const fade = [5, 21, 44, 68, 100]
  sizes.forEach((d, i) => {
    s.addShape('ellipse', {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { type: 'none' }, line: { color: W, width: 0.75, transparency: 100 - fade[i] }
    })
  })
  el(s, cx - 0.2265, cy - 0.2265, 0.453, 0.453, W) // solid hub
}

/** Vertical three-dot "more" affordance. */
function dotsV (s, x, y, color) {
  for (let i = 0; i < 3; i++) el(s, x, y + i * 0.0725, 0.05, 0.05, color)
}

/** Round "more" button: filled disc with three horizontal dots. */
function dotsBadge (s, x, y, d, bg, dot) {
  el(s, x, y, d, d, bg)
  ;[0.208, 0.434, 0.666].forEach(f => el(s, x + d * f, y + d * 0.434, d * 0.133, d * 0.133, dot))
}

/** Round brand badge: filled disc carrying the flame mark. */
function flameBadge (s, x, y, d, bg, fg) {
  el(s, x, y, d, d, bg)
  s.addShape('teardrop', {
    x: x + d * 0.26, y: y + d * 0.22, w: d * 0.48, h: d * 0.48,
    rotate: 315, fill: { color: fg }, line: NOLINE
  })
}

/**
 * Placeholder standing in for a photograph.  The originals are empty picture
 * frames filled with a 5% dot pattern, so these stay deliberately faint and
 * never hide the artwork they overlap.
 */
function imageBox (s, x, y, w, h, r) {
  s.addShape(r >= Math.min(w, h) / 2 ? 'ellipse' : 'roundRect', {
    x, y, w, h, rectRadius: r || undefined,
    fill: { color: P, transparency: 96 }, line: { color: P, width: 0.5, transparency: 85 }
  })
}

/**
 * Smartphone mock-up: notched rounded body carrying the house ramp.  The body
 * is far taller than the slide, so the visible band only spans part of the
 * gradient — the end colours below are read off the reference render.
 * `pale` selects the washed-out variant used on the two section-break slides.
 */
const PHONE_TOP = '5D2FDC'
const PHONE_BTM = '0F0723'
const PALE_TOP = '6C36FE'
const PALE_BTM = '838188'
function phone (s, x, y, w, h, pale) {
  const top = pale ? PALE_TOP : PHONE_TOP
  const btm = pale ? PALE_BTM : PHONE_BTM
  ramp(s, x, y, w, h, w * 0.11, f => ({ color: mix(btm, top, f) }))
  s.addShape('roundRect', {
    x: x + w * 0.249, y: y - 0.08, w: w * 0.509, h: h * 0.041 + 0.08,
    rectRadius: 0.07, fill: { color: K }, line: NOLINE
  })
}

/** Text run. `o`: sz, font, color, alpha, align, lh (line-height multiple). */
function t (s, x, y, w, h, text, o) {
  o = o || {}
  s.addText(text, {
    x, y, w, h,
    fontSize: o.sz || 18,
    fontFace: o.font,
    color: o.color || W,
    transparency: o.alpha,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.lh,
    fit: 'resize',
    isTextBox: true
  })
}

/**
 * Mini statistic chart. `violet` = how many of the four slices are accent
 * coloured; `hole` = doughnut hole size, or 0 for a solid pie.
 */
function pie (s, x, y, w, h, violet, hole) {
  const data = [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] }]
  s.addChart(hole ? 'doughnut' : 'pie', data, {
    x, y, w, h,
    chartColors: [0, 1, 2, 3].map(i => (i < violet ? P : W)),
    holeSize: hole || undefined,
    showLegend: false,
    showTitle: false,
    showValue: false,
    chartArea: { fill: { color: K, transparency: 100 } },
    plotArea: { fill: { color: K, transparency: 100 } }
  })
}

/** Navigation bar, wordmark and search glyph — repeated on every slide. */
function chrome (s) {
  s.addShape('moon', { x: 0.578, y: 0.582, w: 0.194, h: 0.246, rotate: 336, fill: { color: P }, line: NOLINE })
  el(s, 0.75, 0.652, 0.107, 0.107, P)
  t(s, 1.079, 0.585, 1.47, 0.303, 'Your Company', { sz: 12, font: HEAD })
  t(s, 5.384, 0.61, 0.705, 0.252, 'Home ', { sz: 9 })
  t(s, 6.807, 0.61, 0.705, 0.252, 'About', { sz: 9, align: 'center' })
  t(s, 8.23, 0.61, 0.705, 0.252, 'Contact', { sz: 9, align: 'right' })
  rr(s, 11.221, 0.61, 0.913, 0.28, 0.14, W, 78)
  t(s, 11.326, 0.61, 0.705, 0.252, '2040', { sz: 9, align: 'center' })
  ln(s, 12.44, 0.681, 0, 0.121, { width: 0.5 })
  s.addShape('ellipse', { x: 12.696, y: 0.691, w: 0.073, h: 0.073, fill: { type: 'none' }, line: { color: W, width: 0.5 } })
  ln(s, 12.669, 0.755, 0.037, 0.037, { width: 0.5 })
}

/* ------------------------------------------------------------------ slides */

/** Slide 1 — Cover */
function slide01 (s) {
  rings(s, 1.214, 5.304)
  orb(s, -2.145, -2.085, 5.905)
  ln(s, 3.23, -0.393, 0, 4.074, { alpha: 51 })
  phone(s, 1.258, 2.32, 3.734, 7.722)
  imageBox(s, 1.681, 3.241, 2.889, 2.184, 0.385)
  glass(s, 1.68, 5.69, 2.889, 1.093, 0.27, 7, '3D1E8F')
  pie(s, 1.565, 5.741, 1.487, 0.992, 2, 0)
  t(s, 2.851, 5.939, 1.218, 0.286, 'Data Statistic', { sz: 11 })
  t(s, 2.851, 6.209, 1.117, 0.438, '5.980', { sz: 20, font: HEAD })
  dotsV(s, 4.362, 5.91, W)
  dotsBadge(s, 1.855, 3.463, 0.332, K, W)
  glass(s, 4.759, 3.241, 2.026, 3.134, 0.255, 34)
  rr(s, 5.065, 4.644, 0.268, 1.148, 0.134, W)
  rr(s, 5.446, 4.469, 0.268, 1.404, 0.134, W)
  rr(s, 5.829, 4.951, 0.268, 0.921, 0.134, W)
  rr(s, 6.211, 4.389, 0.268, 1.483, 0.134, W)
  gradRect(s, 5.065, 5.052, 0.268, 0.82, 0.134)
  gradRect(s, 5.446, 4.823, 0.268, 1.056, 0.134)
  gradRect(s, 5.829, 5.279, 0.268, 0.593, 0.134)
  gradRect(s, 6.211, 4.637, 0.268, 1.235, 0.134)
  t(s, 4.876, 3.431, 1.336, 0.252, 'Customer Target', { sz: 9 })
  t(s, 4.848, 3.724, 1.741, 0.404, '+75.002%', { sz: 18, font: HEAD })
  ln(s, 5.385, 6.187, 0.774, 0, { width: 2.375 })
  dotsV(s, 6.564, 3.558, W)
  t(s, 7.66, 2.308, 4.057, 0.64, 'Comprehensive', { sz: 32, font: HEAD })
  flameBadge(s, 7.809, 1.775, 0.259, P, W)
  t(s, 8.243, 1.775, 2.253, 0.269, 'Securing Future Grwoth: ', { sz: 10 })
  t(s, 7.701, 3.925, 4.779, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 7.66, 2.936, 2.715, 0.64, 'Invesment', { sz: 32, font: HEAD, color: P })
  t(s, 10.237, 2.936, 2.715, 0.64, 'Pitch Deck', { sz: 32, font: HEAD })
  rr(s, 7.809, 5.553, 1.547, 0.375, 0.188, W)
  gradRect(s, 9.639, 5.553, 1.715, 0.375, 0.188)
  t(s, 8.053, 5.615, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 9.985, 5.615, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  orb(s, 10.815, 4.894, 4.475)
  chrome(s)
  tick(s, 3.23, 2.628)
  ln(s, 12.219, 2.627, 3.612, 0, { alpha: 51 })
  tick(s, 12.211, 2.628)
}

/** Slide 2 — Table Of Content */
function slide02 (s) {
  orb(s, 10.344, -1.584, 4.475)
  ln(s, -1.159, 4.344, 13.065, 0, { alpha: 51 })
  orb(s, -3.356, 4.873, 5.905)
  t(s, 1.027, 2.152, 4.328, 0.64, 'Table Of Content', { sz: 32, font: HEAD })
  rr(s, 1.139, 4.156, 1.547, 0.375, 0.188, W)
  t(s, 1.16, 4.191, 1.507, 0.303, 'Content One', { sz: 12, font: HEAD, color: K, align: 'center' })
  tick(s, 2.687, 4.342, P)
  gradRect(s, 4.208, 4.156, 1.556, 0.375, 0.188)
  t(s, 4.298, 4.191, 1.376, 0.303, 'Content Two', { sz: 12, font: HEAD, align: 'center' })
  tick(s, 4.207, 4.342)
  tick(s, 5.76, 4.342)
  rr(s, 7.281, 4.156, 1.547, 0.375, 0.188, W)
  t(s, 7.333, 4.191, 1.443, 0.303, 'Content Three', { sz: 12, font: HEAD, color: K, align: 'center' })
  tick(s, 7.28, 4.342, P)
  tick(s, 8.822, 4.342, P)
  gradRect(s, 10.349, 4.156, 1.556, 0.375, 0.188)
  t(s, 10.444, 4.191, 1.367, 0.303, 'Content Four', { sz: 12, font: HEAD, align: 'center' })
  tick(s, 10.344, 4.342)
  tick(s, 11.905, 4.342)
  t(s, 1.027, 5.017, 2.556, 0.62, 'Est dignissim sollicitudin vitae semper ultricies dolor', { sz: 11, lh: 1.5 })
  t(s, 4.11, 5.031, 2.556, 0.62, 'Est dignissim sollicitudin vitae semper ultricies dolor', { sz: 11, lh: 1.5 })
  t(s, 7.154, 5.017, 2.556, 0.62, 'Est dignissim sollicitudin vitae semper ultricies dolor', { sz: 11, lh: 1.5 })
  t(s, 10.262, 5.017, 2.556, 0.62, 'Est dignissim sollicitudin vitae semper ultricies dolor', { sz: 11, lh: 1.5 })
  t(s, 6.684, 2.152, 4.13, 0.62, 'Aliquam cras imperdiet egestas, dui duis natoque Aliquam tincidunt risus nostra vulputate', { sz: 11, lh: 1.5 })
  flameBadge(s, 6.26, 2.332, 0.259, P, W)
  chrome(s)
}

/** Slide 3 — Introduction */
function slide03 (s) {
  rings(s, 11.499, 4.218)
  phone(s, 8.805, 1.371, 2.692, 5.568, true)
  imageBox(s, 8.098, 2.119, 2.937, 2.213, 0.279)
  orb(s, 10.815, 4.894, 4.475)
  orb(s, -2.145, -2.085, 5.905)
  t(s, 1.409, 1.866, 3.3, 0.64, 'Introduction', { sz: 32, font: HEAD })
  chrome(s)
  t(s, 1.44, 2.79, 4.362, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 1.44, 3.566, 4.649, 0.62, 'Aquam porttitor dignissim id placerat tempus Amet euis\nnunc quisque commodo sapien cursus', { sz: 11, lh: 1.5 })
  rr(s, 1.541, 4.929, 1.547, 0.375, 0.188, W)
  gradRect(s, 3.371, 4.929, 1.715, 0.375, 0.188)
  t(s, 1.785, 4.991, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 3.717, 4.991, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  ln(s, 9.43, 6.679, 1.441, 0, { width: 2.375 })
  dotsBadge(s, 10.521, 2.312, 0.332, K, W)
  t(s, 1.44, 6.439, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 1.44, 6.019, 1.117, 0.438, '175K', { sz: 20, font: HEAD })
  t(s, 3.371, 6.439, 1.743, 0.269, 'Best Company', { sz: 10 })
  t(s, 3.371, 6.019, 1.715, 0.438, '90.0%', { sz: 20, font: HEAD })
  glass(s, 8.098, 4.677, 2.976, 1.439, 0.28, 7, '7D6CA9')
  ln(s, 6.716, 3.41, 0, 4.852, { alpha: 51 })
  tick(s, 6.715, 3.41)
  pie(s, 9.651, 4.884, 1.487, 0.992, 2, 0)
  t(s, 8.539, 5.088, 1.336, 0.252, 'Total Income', { sz: 9 })
  t(s, 8.512, 5.381, 1.741, 0.404, '$27.005', { sz: 18, font: HEAD })
  dotsV(s, 8.361, 5.015, W)
}

/** Slide 4 — Problem Statement */
function slide04 (s) {
  orb(s, -2.145, -2.085, 5.905)
  ln(s, 2.543, -0.135, 0, 3.563, { alpha: 51 })
  orb(s, 10.815, 4.894, 4.475)
  rings(s, 5.611, 2.647)
  phone(s, 0.54, 2.832, 4.007, 8.286)
  glass(s, 0.817, 4.156, 4.245, 2.65, 0.302, 7, '4E27B8')
  chrome(s)
  t(s, 8.944, 1.438, 3.216, 1.178, 'Problem Statement', { sz: 32, font: HEAD })
  t(s, 8.965, 2.832, 3.92, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  rr(s, 1.714, 4.912, 2.645, 0.192, 0.096, W)
  rr(s, 1.714, 5.366, 2.645, 0.192, 0.096, W)
  rr(s, 1.714, 5.821, 2.645, 0.192, 0.096, W)
  rr(s, 1.714, 6.287, 2.645, 0.192, 0.096, W)
  gradRect(s, 1.715, 4.91, 1.235, 0.192, 0.096)
  gradRect(s, 1.715, 5.365, 1.997, 0.192, 0.096)
  gradRect(s, 1.715, 5.821, 1.362, 0.192, 0.096)
  gradRect(s, 1.714, 6.287, 2.273, 0.192, 0.096)
  rr(s, 1.714, 4.508, 2.645, 0.192, 0.096, W)
  gradRect(s, 1.715, 4.509, 1.782, 0.192, 0.096)
  t(s, 1.079, 4.45, 0.509, 0.252, 'May', { sz: 9, font: HEAD })
  t(s, 1.079, 4.871, 0.509, 0.252, 'Apr', { sz: 9, font: HEAD })
  t(s, 1.079, 5.336, 0.509, 0.252, 'Mar', { sz: 9, font: HEAD })
  t(s, 1.079, 5.791, 0.509, 0.252, 'Feb', { sz: 9, font: HEAD })
  t(s, 1.079, 6.26, 0.509, 0.252, 'Jan', { sz: 9, font: HEAD })
  ln(s, 1.715, 4.531, 0, 1.948, { color: P, alpha: 51 })
  t(s, 1.51, 3.699, 2.221, 0.252, 'Monthly Downtime Percentage', { sz: 9, font: HEAD, align: 'center' })
  t(s, 4.404, 4.478, 0.509, 0.252, '30%', { sz: 9, alpha: 31 })
  t(s, 4.404, 4.922, 0.509, 0.252, '60%', { sz: 9, alpha: 31 })
  t(s, 4.404, 5.365, 0.509, 0.252, '30%', { sz: 9, alpha: 31 })
  t(s, 4.404, 5.808, 0.509, 0.252, '50%', { sz: 9, alpha: 31 })
  t(s, 4.404, 6.251, 0.509, 0.252, '80%', { sz: 9, alpha: 31 })
  flameBadge(s, 1.173, 3.696, 0.259, K, W)
  rr(s, 5.59, 4.922, 2.681, 1.885, 0.304, W)
  t(s, 5.839, 5.285, 1.937, 0.303, 'Ideal', { sz: 12, font: HEAD, color: K })
  t(s, 5.839, 5.633, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  rr(s, 8.798, 4.922, 2.681, 1.885, 0.304, P)
  t(s, 9.047, 5.285, 2.391, 0.303, 'Consequences', { sz: 12, font: HEAD })
  t(s, 9.047, 5.633, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  rr(s, 5.553, 2.496, 2.681, 1.885, 0.304, P)
  t(s, 5.802, 2.86, 1.345, 0.303, 'Proposal', { sz: 12, font: HEAD })
  t(s, 5.802, 3.207, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  dotsBadge(s, 9.075, 3.899, 0.332, W, K)
  t(s, 9.583, 4.281, 1.22, 0.269, 'There Are Many', { sz: 10 })
  t(s, 9.583, 3.861, 1.22, 0.438, '275K', { sz: 20, font: HEAD })
  ln(s, 12.213, 4.118, 0, 3.563, { alpha: 51 })
  tick(s, 12.213, 4.143)
  tick(s, 2.542, 3.167)
}

/** Slide 5 — Best Solution */
function slide05 (s) {
  rings(s, 5.603, 5.01)
  imageBox(s, 5.653, 1.864, 2.937, 4.917, 0.342)
  orb(s, -2.196, 5.511, 4.475)
  gradRect(s, 2.706, 3.999, 1.4, 0.375, 0.188)
  orb(s, 10.495, -1.772, 4.475)
  chrome(s)
  t(s, 0.866, 1.864, 3.314, 0.64, 'Best Solution', { sz: 32, font: HEAD })
  glass(s, 5.859, 5.034, 2.525, 1.543, 0.314, 69)
  t(s, 5.991, 5.238, 1.312, 0.269, 'Income', { sz: 10 })
  t(s, 5.991, 5.508, 1.713, 0.438, '$234.002', { sz: 20, font: HEAD })
  dotsV(s, 8.132, 5.312, W)
  rr(s, 6.061, 6.112, 2.121, 0.308, 0.154, W)
  t(s, 6.352, 6.14, 1.539, 0.252, 'Experience More', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 0.897, 2.703, 4.116, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 1.488, 6.14, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 1.488, 5.72, 1.117, 0.438, '250K', { sz: 20, font: HEAD })
  t(s, 3.084, 6.14, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 3.084, 5.72, 1.715, 0.438, '12.9K', { sz: 20, font: HEAD })
  pie(s, 0.685, 5.615, 0.944, 0.629, 2, 0)
  rr(s, 8.935, 1.865, 3.402, 1.457, 0.235, P)
  t(s, 9.152, 2.115, 2.391, 0.303, 'Challenges', { sz: 12, font: HEAD })
  t(s, 9.152, 2.462, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  rr(s, 8.935, 3.595, 3.402, 1.457, 0.235, W)
  t(s, 9.152, 3.844, 2.391, 0.303, 'Solutions', { sz: 12, font: HEAD, color: K })
  t(s, 9.152, 4.192, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  rr(s, 8.935, 5.324, 3.402, 1.457, 0.235, P)
  t(s, 9.152, 5.574, 2.391, 0.303, 'Benefits', { sz: 12, font: HEAD })
  t(s, 9.152, 5.921, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  dotsBadge(s, 5.919, 2.13, 0.332, K, W)
  ln(s, 12.102, 2.403, 0, 0.382, { width: 2.375 })
  ln(s, 12.102, 3.942, 0, 0.763, { color: P, width: 2.375 })
  ln(s, 12.102, 5.862, 0, 0.382, { width: 2.375 })
  rr(s, 0.997, 3.999, 1.4, 0.375, 0.188, W)
  t(s, 1.185, 4.06, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 2.893, 4.06, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
}

/** Slide 6 — Unique Value Proposition */
function slide06 (s) {
  orb(s, 10.815, 4.894, 4.475)
  rings(s, 10.745, 5.378)
  imageBox(s, 4.908, 4.419, 5.837, 2.12, 0.379)
  orb(s, -2.145, -2.085, 5.905)
  chrome(s)
  t(s, 9.088, 2.303, 3.314, 1.178, 'Unique Value Proposition', { sz: 32, font: HEAD })
  rr(s, 2.702, 2.031, 2.681, 1.885, 0.304, W)
  t(s, 2.951, 2.395, 1.682, 0.303, 'Features', { sz: 12, font: HEAD, color: K })
  t(s, 2.951, 2.743, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  el(s, 4.671, 1.854, 0.385, 0.385, P)
  t(s, 4.633, 1.904, 0.462, 0.286, '01', { sz: 11, font: HEAD, align: 'center' })
  rr(s, 5.857, 2.031, 2.681, 1.885, 0.304, P)
  t(s, 6.106, 2.395, 1.682, 0.303, 'Benefits', { sz: 12, font: HEAD })
  t(s, 6.106, 2.743, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  el(s, 7.826, 1.854, 0.385, 0.385, W)
  t(s, 7.788, 1.904, 0.462, 0.286, '02', { sz: 11, font: HEAD, color: K, align: 'center' })
  dotsBadge(s, 5.091, 4.685, 0.332, K, W)
  t(s, 1.396, 4.594, 2.958, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 1.967, 6.106, 1.743, 0.269, 'Total Balance', { sz: 10 })
  t(s, 1.967, 5.686, 1.967, 0.438, '$328.90K', { sz: 20, font: HEAD })
  flameBadge(s, 1.485, 5.694, 0.332, P, W)
  rr(s, -0.9, 4.419, 1.743, 2.12, 0.294, W)
  rr(s, 12.769, 2.031, 2.681, 1.885, 0.304, P)
  ln(s, -1.943, 2.933, 3.612, 0, { alpha: 51 })
  tick(s, 1.676, 2.933)
}

/** Slide 7 — Market Opportunity */
function slide07 (s) {
  orb(s, 10.431, 5.604, 4.475)
  orb(s, -2.145, -2.085, 5.905)
  rr(s, 6.374, 5.125, 1.4, 0.375, 0.188, W)
  gradRect(s, 4.666, 5.125, 1.4, 0.375, 0.188)
  glass(s, 6.043, 1.989, 6.294, 2.178, 0.424)
  chrome(s)
  t(s, 0.868, 2.038, 3.314, 1.178, 'Market Opportunity', { sz: 32, font: HEAD })
  pie(s, 6.107, 2.202, 2.151, 1.434, 1, 75)
  t(s, 6.847, 3.628, 0.671, 0.252, '2020', { sz: 9, alpha: 40, align: 'center' })
  t(s, 6.847, 2.784, 0.671, 0.303, '30%', { sz: 12, font: HEAD, align: 'center' })
  pie(s, 7.928, 2.202, 2.151, 1.434, 2, 75)
  t(s, 8.668, 3.628, 0.671, 0.252, '2030', { sz: 9, alpha: 40, align: 'center' })
  t(s, 8.668, 2.784, 0.671, 0.303, '70%', { sz: 12, font: HEAD, align: 'center' })
  pie(s, 9.748, 2.202, 2.151, 1.434, 3, 75)
  t(s, 10.488, 3.628, 0.671, 0.252, '2040', { sz: 9, alpha: 40, align: 'center' })
  t(s, 10.488, 2.784, 0.671, 0.303, '80%', { sz: 12, font: HEAD, align: 'center' })
  dotsV(s, 11.929, 2.432, W)
  t(s, 1.331, 3.636, 3.92, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  rr(s, 0.997, 5.125, 3.144, 1.469, 0.257, W)
  t(s, 1.167, 5.359, 1.682, 0.303, 'Income', { sz: 12, font: HEAD, color: K })
  t(s, 1.167, 5.707, 3.014, 0.572, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  flameBadge(s, 0.949, 3.798, 0.259, P, W)
  dotsBadge(s, 3.662, 5.25, 0.332, K, W)
  t(s, 4.854, 5.186, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, align: 'center' })
  t(s, 6.563, 5.186, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 4.733, 6.324, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 4.733, 5.904, 1.117, 0.438, '200K', { sz: 20, font: HEAD })
  t(s, 6.476, 6.324, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 6.476, 5.904, 1.715, 0.438, '75.0K', { sz: 20, font: HEAD })
  rings(s, 9.377, 5.904)
  rr(s, 12.769, 1.368, 1.341, 1.064, 0.171, P)
  rr(s, 12.769, 2.882, 1.341, 1.892, 0.216, W)
  ln(s, 4.593, -0.748, 0, 3.612, { alpha: 51 })
  tick(s, 4.592, 2.872)
  imageBox(s, 9.339, 4.774, 2.51, 2.007, 0.312)
}

/** Slide 8 — Target Audience */
function slide08 (s) {
  imageBox(s, 6.872, 6.433, 2.433, 1.771, 0.275)
  orb(s, 10.495, -1.772, 4.475)
  orb(s, -2.196, 5.511, 4.475)
  rings(s, 2.947, 4.467)
  chrome(s)
  t(s, 10.005, 2.205, 3.249, 1.178, 'Target \nAudience', { sz: 32, font: HEAD })
  rr(s, 3.739, 5.273, 2.433, 1.508, 0.243, W)
  t(s, 3.916, 5.528, 1.682, 0.303, 'Psychographic', { sz: 12, font: HEAD, color: K })
  t(s, 3.916, 5.875, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 6.872, 4.29, 2.433, 1.508, 0.243, P)
  t(s, 7.05, 4.544, 1.682, 0.303, 'Demographic', { sz: 12, font: HEAD })
  t(s, 7.05, 4.892, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 10.005, 5.273, 2.433, 1.508, 0.243, W)
  t(s, 10.183, 5.528, 1.682, 0.303, 'Geographic', { sz: 12, font: HEAD, color: K })
  t(s, 10.183, 5.875, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 0.589, 4.322, 2.433, 1.508, 0.243, P)
  t(s, 0.767, 4.577, 1.682, 0.303, 'Behavioral', { sz: 12, font: HEAD })
  t(s, 0.767, 4.925, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  dotsBadge(s, 7.05, 6.615, 0.332, K, W)
  t(s, 5.467, 2.205, 3.92, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  flameBadge(s, 5.086, 2.379, 0.259, P, W)
  rr(s, 5.549, 3.094, 1.598, 0.375, 0.188, W)
  t(s, 5.836, 3.156, 1.024, 0.252, 'Learn More', { sz: 9, font: HEAD, color: K, align: 'center' })
  ln(s, 11.072, 4.204, 3.612, 0, { alpha: 51 })
  tick(s, 11.064, 4.204)
  t(s, 2.432, 2.756, 1.743, 0.438, 'Lorem Ipsum Dolor Sit amet Conicquenc', { sz: 10 })
  t(s, 2.432, 2.271, 1.117, 0.438, '+175K', { sz: 20, font: HEAD })
}

/** Slide 9 — Product Overview */
function slide09 (s) {
  ln(s, 6.667, 4.737, 0, 3.612, { alpha: 51 })
  orb(s, -3.087, -2.628, 5.905)
  orb(s, 10.431, 5.163, 4.475)
  rr(s, 5.373, 4.423, 2.587, 2.068, 0.349, W)
  rr(s, 2.292, 4.713, 2.587, 2.068, 0.349, P)
  rr(s, 8.455, 4.713, 2.587, 2.068, 0.349, P)
  chrome(s)
  t(s, 4.399, 1.595, 4.574, 0.64, 'Product Overview', { sz: 32, font: HEAD, align: 'center' })
  t(s, 3.727, 2.492, 5.918, 0.62, 'PLACEHOLDER', { sz: 11, align: 'center', lh: 1.5 })
  t(s, 2.744, 6.317, 1.682, 0.303, 'Product One', { sz: 12, font: HEAD, align: 'center' })
  t(s, 5.826, 6.032, 1.682, 0.303, 'Product Two', { sz: 12, font: HEAD, color: K, align: 'center' })
  t(s, 8.907, 6.317, 1.682, 0.303, 'Product Three', { sz: 12, font: HEAD, align: 'center' })
  rings(s, 2.449, 4.218)
  flameBadge(s, 10.702, 2.591, 0.259, P, W)
  t(s, 11.153, 2.852, 1.743, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 11.137, 2.586, 2.016, 0.269, 'The Details', { sz: 10, font: HEAD })
  t(s, 1.482, 2.91, 1.743, 0.269, 'Best Quality', { sz: 10 })
  t(s, 1.482, 2.49, 1.715, 0.438, '90.0%', { sz: 20, font: HEAD })
  dotsBadge(s, 0.943, 2.597, 0.303, W, K)
  tick(s, 6.667, 6.489)
  imageBox(s, 2.292, 4.082, 2.587, 2.068, 0.342)
  imageBox(s, 5.373, 3.807, 2.587, 2.068, 0.342)
  imageBox(s, 8.455, 4.082, 2.587, 2.068, 0.342)
}

/** Slide 10 — Traction */
function slide10 (s) {
  orb(s, 10.495, -1.772, 4.475)
  orb(s, -2.196, 5.511, 4.475)
  rings(s, 2.985, 2.674)
  ln(s, 3.465, 3.987, 0, 1.984, { alpha: 51 })
  ln(s, 3.465, 3.987, 3.09, 0, { alpha: 51 })
  ln(s, 6.305, 2.478, 0, 1.92, { alpha: 51 })
  ln(s, 6.305, 2.478, 7.35, 0, { alpha: 51 })
  ln(s, 0, 5.55, 3.612, 0, { alpha: 51 })
  chrome(s)
  t(s, 8.23, 3.996, 2.525, 0.64, 'Traction', { sz: 32, font: HEAD })
  rr(s, 2.249, 4.796, 2.433, 1.508, 0.243, P)
  t(s, 2.427, 5.051, 1.682, 0.303, '2020', { sz: 12, font: HEAD })
  t(s, 2.427, 5.398, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 5.089, 3.232, 2.433, 1.508, 0.243, W)
  t(s, 5.267, 3.487, 1.682, 0.303, '2030', { sz: 12, font: HEAD, color: K })
  t(s, 5.267, 3.835, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 7.943, 1.724, 2.433, 1.508, 0.243, P)
  t(s, 8.121, 1.979, 1.682, 0.303, '2040', { sz: 12, font: HEAD })
  t(s, 8.121, 2.327, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  tick(s, 3.466, 3.991)
  tick(s, 6.305, 2.479)
  t(s, 8.241, 4.881, 3.92, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  flameBadge(s, 5.347, 5.687, 0.259, P, W)
  t(s, 5.798, 5.948, 1.743, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 5.782, 5.682, 2.016, 0.269, 'The Details', { sz: 10, font: HEAD })
  rr(s, 8.337, 5.902, 1.598, 0.375, 0.188, W)
  t(s, 8.624, 5.964, 1.024, 0.252, 'Learn More', { sz: 9, font: HEAD, color: K, align: 'center' })
  imageBox(s, -0.381, 1.682, 3.339, 1.988, 0.329)
}

/** Slide 11 — Business Model */
function slide11 (s) {
  orb(s, -1.72, -1.77, 4.475)
  orb(s, 11.365, 5.723, 4.475)
  ln(s, 1.625, -0.534, 0, 3.612, { alpha: 51 })
  tick(s, 1.625, 3.086)
  chrome(s)
  t(s, 9.328, 1.899, 3.276, 1.178, 'Business Model', { sz: 32, font: HEAD })
  rr(s, 2.731, 2.556, 2.802, 1.719, 0.246, P)
  t(s, 2.985, 2.942, 1.682, 0.303, 'Distribution', { sz: 12, font: HEAD })
  t(s, 2.985, 3.348, 2.255, 0.572, 'Est dignissim sollicitudin vitaeul semper ultricies maximus dol', { sz: 10, lh: 1.5 })
  rr(s, 5.854, 2.556, 2.802, 1.719, 0.246, W)
  t(s, 6.108, 2.942, 1.682, 0.303, 'Customer', { sz: 12, font: HEAD, color: K })
  t(s, 6.108, 3.348, 2.255, 0.572, 'Est dignissim sollicitudin vitaeul semper ultricies maximus dol', { sz: 10, color: K, lh: 1.5 })
  rr(s, 5.866, 4.661, 2.802, 1.719, 0.246, P)
  t(s, 6.12, 5.048, 1.682, 0.303, 'Revenue', { sz: 12, font: HEAD })
  t(s, 6.12, 5.453, 2.255, 0.572, 'Est dignissim sollicitudin vitaeul semper ultricies maximus dol', { sz: 10, lh: 1.5 })
  t(s, 9.361, 3.282, 3.579, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  rings(s, 1.645, 4.643)
  gradRect(s, 11.178, 4.62, 1.4, 0.375, 0.188)
  rr(s, 9.468, 4.62, 1.4, 0.375, 0.188, W)
  t(s, 9.656, 4.681, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 11.365, 4.681, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  t(s, 9.887, 6.143, 1.241, 0.269, 'There Are Many', { sz: 10 })
  t(s, 9.887, 5.723, 1.117, 0.438, '175K', { sz: 20, font: HEAD })
  t(s, 11.483, 6.143, 1.241, 0.269, 'Best Company', { sz: 10 })
  t(s, 11.483, 5.723, 1.241, 0.438, '200K', { sz: 20, font: HEAD })
  flameBadge(s, 9.499, 5.812, 0.259, P, W)
  imageBox(s, 0.957, 4.643, 4.581, 2.337, 0.283)
}

/** Slide 12 — Go-to-Market Strategy */
function slide12 (s) {
  phone(s, 8.897, 4.424, 3.808, 7.875)
  imageBox(s, 9.322, 5.253, 2.959, 1.804, 0.286)
  orb(s, 10.354, -1.436, 4.475)
  orb(s, -2.196, 5.511, 4.475)
  ln(s, 5.106, -0.252, 0, 2.816, { alpha: 51 })
  rings(s, 1.888, 4.768)
  chrome(s)
  t(s, 7.243, 1.772, 4.232, 1.178, 'Go-to-Market Strategy', { sz: 32, font: HEAD })
  rr(s, 0.564, 1.983, 2.637, 1.736, 0.28, P)
  t(s, 0.742, 2.262, 1.682, 0.505, 'In-Depth Market Research', { sz: 12, font: HEAD })
  t(s, 0.742, 2.851, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 2.549, 4.587, 2.637, 1.736, 0.28, W)
  t(s, 2.727, 4.866, 1.682, 0.505, 'SEO and SEM Strategy', { sz: 12, font: HEAD, color: K })
  t(s, 2.727, 5.455, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 5.713, 4.573, 2.637, 1.736, 0.28, P)
  t(s, 5.89, 4.852, 1.682, 0.505, 'Events and Trade Fairs', { sz: 12, font: HEAD })
  t(s, 5.89, 5.441, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 3.787, 1.983, 2.637, 1.736, 0.28, W)
  t(s, 3.965, 2.262, 1.682, 0.505, 'Social Media Marketing', { sz: 12, font: HEAD, color: K })
  t(s, 3.965, 2.851, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  t(s, 7.243, 3.236, 4.429, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  tick(s, 5.106, 1.982)
  dotsBadge(s, 9.498, 5.432, 0.332, K, W)
  imageBox(s, -0.41, 4.587, 2.433, 1.736, 0.262)
}

/** Slide 13 — Customer Acquisition */
function slide13 (s) {
  rings(s, 5.917, 3.127)
  orb(s, 10.431, -1.811, 4.475)
  orb(s, -2.047, 5.416, 4.475)
  chrome(s)
  t(s, 0.852, 1.852, 3.269, 1.178, 'Customer Acquisition', { sz: 32, font: HEAD })
  rr(s, 8.789, 1.977, 2.433, 1.508, 0.243, W)
  t(s, 8.966, 2.232, 1.682, 0.303, 'Purchase', { sz: 12, font: HEAD, color: K })
  t(s, 8.966, 2.58, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 8.789, 4.014, 2.433, 1.508, 0.243, P)
  t(s, 8.966, 4.269, 1.682, 0.303, 'Awareness', { sz: 12, font: HEAD })
  t(s, 8.966, 4.617, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 5.847, 3.027, 2.433, 1.508, 0.243, P)
  t(s, 6.024, 3.281, 1.682, 0.303, 'Nurturing', { sz: 12, font: HEAD })
  t(s, 6.024, 3.629, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 5.847, 5.063, 2.433, 1.508, 0.243, W)
  t(s, 6.024, 5.318, 1.682, 0.303, 'Interest', { sz: 12, font: HEAD, color: K })
  t(s, 6.024, 5.666, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  t(s, 9.407, 6.437, 1.241, 0.269, 'There Are Many', { sz: 10 })
  t(s, 9.407, 6.017, 1.117, 0.438, '25K', { sz: 20, font: HEAD })
  flameBadge(s, 9.018, 6.106, 0.259, P, W)
  ln(s, 12.175, 3.121, 0, 4.601, { alpha: 51 })
  tick(s, 12.175, 3.08)
  t(s, 0.895, 3.407, 3.579, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  glass(s, 2.38, 5.063, 2.964, 1.508, 0.293)
  pie(s, 2.189, 5.218, 1.798, 1.198, 1, 75)
  t(s, 2.752, 5.657, 0.671, 0.286, '30%', { sz: 11, font: HEAD, align: 'center' })
  t(s, 3.719, 5.581, 1.197, 0.525, 'There Are Many\nVariat ', { sz: 9, lh: 1.5 })
  dotsV(s, 5.07, 5.327, W)
}

/** Slide 14 — Competitive Analysis */
function slide14 (s) {
  rings(s, 8.466, 3.823)
  imageBox(s, 5.594, 3.685, 2.959, 3.186, 0.316)
  orb(s, 10.841, -1.982, 4.475)
  orb(s, -1.659, 5.354, 4.475)
  chrome(s)
  t(s, 5.498, 1.918, 3.269, 1.178, 'Competitive Analysis', { sz: 32, font: HEAD })
  rr(s, 0.957, 1.955, 3.402, 1.457, 0.235, P)
  t(s, 1.173, 2.204, 2.391, 0.303, 'Operations', { sz: 12, font: HEAD })
  t(s, 1.173, 2.552, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  rr(s, 1.369, 3.685, 3.402, 1.457, 0.235, W)
  t(s, 1.586, 3.934, 2.391, 0.303, 'Product', { sz: 12, font: HEAD, color: K })
  t(s, 1.586, 4.282, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  rr(s, 0.957, 5.414, 3.402, 1.457, 0.235, P)
  t(s, 1.173, 5.663, 2.391, 0.303, 'Marketing', { sz: 12, font: HEAD })
  t(s, 1.173, 6.011, 3.152, 0.572, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  ln(s, 4.124, 2.493, 0, 0.382, { width: 2.375 })
  ln(s, 4.535, 4.032, 0, 0.763, { color: P, width: 2.375 })
  ln(s, 4.124, 5.952, 0, 0.382, { width: 2.375 })
  rr(s, -0.571, 3.685, 1.644, 1.457, 0.235, W)
  dotsBadge(s, 5.78, 3.904, 0.332, K, W)
  t(s, 9.048, 4.854, 3.586, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  flameBadge(s, 9.132, 6.063, 0.259, P, W)
  t(s, 9.583, 6.324, 1.743, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 9.567, 6.058, 2.016, 0.269, 'The Details', { sz: 10, font: HEAD })
  ln(s, 10.535, -1.005, 0, 4.601, { alpha: 51 })
  tick(s, 10.535, 3.637)
}

/** Slide 15 — Let’s Take A Break */
function slide15 (s) {
  rings(s, 3.419, 4.29)
  phone(s, 3.419, 1.691, 2.514, 5.198, true)
  imageBox(s, 3.856, 2.37, 2.651, 3.842, 0.276)
  orb(s, -1.72, -1.77, 4.475)
  orb(s, 11.221, 5.263, 4.475)
  chrome(s)
  ln(s, 3.955, 6.597, 1.441, 0, { width: 2.375 })
  t(s, 7.306, 2.253, 3.269, 1.313, 'Let\u2019s Take A Break', { sz: 36, font: HEAD })
  t(s, 7.355, 4.037, 3.586, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  gradRect(s, 9.147, 5.456, 1.4, 0.375, 0.188)
  rr(s, 7.437, 5.456, 1.4, 0.375, 0.188, W)
  t(s, 7.625, 5.517, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 9.334, 5.517, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  dotsBadge(s, 6.017, 2.566, 0.332, K, W)
}

/** Slide 16 — Market Strategy */
function slide16 (s) {
  rings(s, 6.669, 4.656)
  ln(s, 8.042, -1.005, 0, 4.601, { alpha: 51 })
  imageBox(s, 6.569, 2.411, 3.313, 1.77, 0.226)
  orb(s, 10.595, -1.818, 4.475)
  orb(s, -1.713, 5.663, 4.475)
  rr(s, 12.188, 4.572, 2.433, 1.77, 0.244, P)
  chrome(s)
  t(s, 0.817, 2.411, 4.21, 0.64, 'Market Strategy', { sz: 32, font: HEAD })
  rr(s, 10.298, 2.411, 2.433, 1.77, 0.244, P)
  t(s, 10.476, 2.988, 1.682, 0.303, 'Promotion', { sz: 12, font: HEAD })
  t(s, 10.476, 3.336, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 6.569, 4.572, 2.433, 1.77, 0.244, P)
  t(s, 6.747, 5.149, 1.392, 0.303, 'Product', { sz: 12, font: HEAD })
  t(s, 6.747, 5.497, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 9.379, 4.572, 2.433, 1.77, 0.244, W)
  t(s, 9.557, 5.149, 1.682, 0.303, 'Procces', { sz: 12, font: HEAD, color: K })
  t(s, 9.557, 5.497, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  dotsBadge(s, 6.747, 2.565, 0.332, K, W)
  t(s, 0.874, 3.336, 4.21, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 1.354, 5.8, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 1.354, 5.38, 1.117, 0.438, '+147%', { sz: 20, font: HEAD })
  t(s, 3.284, 5.8, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 3.284, 5.38, 1.715, 0.438, '170K', { sz: 20, font: HEAD })
  flameBadge(s, 0.957, 5.46, 0.259, P, W)
  tick(s, 8.041, 2.412)
}

/** Slide 17 — Sales Strategy */
function slide17 (s) {
  imageBox(s, 0.56, 3.652, 3.616, 1.508, 0.234)
  orb(s, -1.524, -1.834, 4.475)
  orb(s, 11.027, 5.346, 4.475)
  rings(s, 12.62, 3.779)
  ln(s, 11.553, 2.623, 0, 4.891, { alpha: 51 })
  chrome(s)
  t(s, 0.866, 2.146, 4.21, 0.64, 'Sales Strategy', { sz: 32, font: HEAD })
  t(s, 0.866, 5.967, 4.21, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  rr(s, 4.626, 3.652, 2.433, 1.508, 0.243, P)
  t(s, 4.804, 3.907, 1.682, 0.303, 'Production', { sz: 12, font: HEAD })
  t(s, 4.804, 4.255, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 7.51, 3.652, 2.433, 1.508, 0.243, W)
  t(s, 7.687, 3.907, 1.682, 0.303, 'Promotion', { sz: 12, font: HEAD, color: K })
  t(s, 7.687, 4.255, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 10.336, 3.652, 2.433, 1.508, 0.243, P)
  t(s, 10.514, 3.907, 1.682, 0.303, 'Ideantion', { sz: 12, font: HEAD })
  t(s, 10.514, 4.255, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  gradRect(s, 7.794, 6.117, 1.4, 0.375, 0.188)
  rr(s, 6.084, 6.117, 1.4, 0.375, 0.188, W)
  t(s, 6.273, 6.179, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 7.981, 6.179, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  dotsBadge(s, 3.721, 3.791, 0.332, K, W)
  t(s, 6.133, 2.623, 1.743, 0.269, 'Customer Target', { sz: 10 })
  t(s, 6.133, 2.203, 1.312, 0.438, '+75.9%', { sz: 20, font: HEAD })
  t(s, 8.064, 2.623, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 8.064, 2.203, 1.715, 0.438, '+270%', { sz: 20, font: HEAD })
  flameBadge(s, 5.736, 2.284, 0.259, P, W)
  tick(s, 11.553, 2.601)
}

/** Slide 18 — Partnerships */
function slide18 (s) {
  rings(s, 2.795, 4.54)
  imageBox(s, 0.56, 2.19, 2.293, 3.922, 0.307)
  ln(s, 8.603, -1.877, 0, 4.891, { alpha: 51 })
  orb(s, 10.595, -1.818, 4.475)
  orb(s, -1.713, 5.663, 4.475)
  chrome(s)
  t(s, 3.502, 2.494, 3.472, 0.64, 'Partnerships', { sz: 32, font: HEAD })
  rr(s, 7.387, 2.411, 2.433, 1.508, 0.243, W)
  t(s, 7.565, 2.666, 1.682, 0.303, 'Partner', { sz: 12, font: HEAD, color: K })
  t(s, 7.565, 3.014, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 10.298, 2.411, 2.433, 1.508, 0.243, P)
  t(s, 10.476, 2.666, 1.682, 0.303, 'Finance', { sz: 12, font: HEAD })
  t(s, 10.476, 3.014, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 7.387, 4.414, 2.433, 1.508, 0.243, P)
  t(s, 7.565, 4.669, 1.682, 0.303, 'Agreement', { sz: 12, font: HEAD })
  t(s, 7.565, 5.016, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  glass(s, 10.298, 4.414, 2.433, 1.508, 0.293)
  t(s, 3.532, 3.44, 3.176, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  gradRect(s, 5.308, 5.016, 1.4, 0.375, 0.188)
  rr(s, 3.599, 5.016, 1.4, 0.375, 0.188, W)
  t(s, 3.787, 5.078, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 5.495, 5.078, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
  dotsBadge(s, 0.7, 2.334, 0.332, K, W)
  rr(s, 10.595, 4.931, 0.162, 0.754, 0.081, W)
  gradRect(s, 10.595, 5.155, 0.162, 0.53, 0.081)
  rr(s, 10.837, 5.04, 0.162, 0.64, 0.081, W)
  gradRect(s, 10.837, 5.252, 0.162, 0.434, 0.081)
  rr(s, 11.079, 4.777, 0.162, 0.899, 0.081, W)
  gradRect(s, 11.079, 4.927, 0.162, 0.749, 0.081)
  t(s, 11.405, 5.344, 1.364, 0.252, 'There Are Many', { sz: 9 })
  t(s, 11.405, 4.924, 1.117, 0.438, '175K', { sz: 20, font: HEAD })
  dotsV(s, 12.497, 4.709, W)
  tick(s, 8.604, 2.412)
}

/** Slide 19 — Product Roadmap */
function slide19 (s) {
  rings(s, 3.3, 2.28)
  orb(s, -1.524, -1.834, 4.475)
  orb(s, 11.027, 5.346, 4.475)
  ln(s, 9.396, 5.989, 3.859, 0, { alpha: 51 })
  ln(s, -0.194, 2.972, 5.621, 0, { alpha: 51 })
  ln(s, 8.68, 4.481, 0, 1.788, { alpha: 51 })
  ln(s, 4.821, 4.481, 3.859, 0, { alpha: 51 })
  chrome(s)
  t(s, 7.464, 2.035, 3.472, 1.178, 'Product Roadmap', { sz: 32, font: HEAD })
  rr(s, 0.957, 2.218, 2.433, 1.508, 0.243, P)
  t(s, 1.134, 2.473, 1.682, 0.303, '2040', { sz: 12, font: HEAD })
  t(s, 1.134, 2.821, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  rr(s, 4.21, 3.726, 2.433, 1.508, 0.243, W)
  t(s, 4.388, 3.981, 1.682, 0.303, '2030', { sz: 12, font: HEAD, color: K })
  t(s, 4.388, 4.329, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  rr(s, 7.464, 5.235, 2.433, 1.508, 0.243, P)
  t(s, 7.642, 5.49, 1.682, 0.303, '2020', { sz: 12, font: HEAD })
  t(s, 7.642, 5.837, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  ln(s, 5.427, 2.972, 0, 1.723, { alpha: 51 })
  tick(s, 5.426, 2.973)
  tick(s, 8.679, 4.481)
  t(s, 10.442, 2.312, 2.53, 0.62, 'Est dignissim sollicitudin vitae semperin ultrics maximus ', { sz: 11, lh: 1.5 })
  t(s, 10.909, 3.947, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 10.909, 3.526, 1.117, 0.438, '135K', { sz: 20, font: HEAD })
  flameBadge(s, 10.512, 3.607, 0.259, P, W)
  gradRect(s, 2.666, 6.113, 1.4, 0.375, 0.188)
  rr(s, 0.957, 6.113, 1.4, 0.375, 0.188, W)
  t(s, 1.145, 6.174, 1.024, 0.252, 'Get Started', { sz: 9, font: HEAD, color: K, align: 'center' })
  t(s, 2.853, 6.174, 1.024, 0.252, 'Learn More ', { sz: 9, font: HEAD, align: 'center' })
}

/** Slide 20 — Technology Stack */
function slide20 (s) {
  rings(s, 1.081, 3.981)
  orb(s, -1.524, -1.834, 4.475)
  orb(s, 11.027, 5.346, 4.475)
  chrome(s)
  t(s, 3.818, 4.864, 4.567, 0.64, 'Technology Stack', { sz: 32, font: HEAD })
  rr(s, 0.957, 1.803, 7.979, 2.328, 0.375, P)
  t(s, 1.301, 2.536, 1.682, 0.303, 'Lifestyle Change', { sz: 12, font: HEAD })
  t(s, 1.301, 2.884, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  t(s, 3.846, 2.536, 1.682, 0.303, 'Exhaustion', { sz: 12, font: HEAD })
  t(s, 3.846, 2.884, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  t(s, 6.391, 2.536, 2.358, 0.505, 'Market Uncertainty', { sz: 12, font: HEAD })
  t(s, 6.391, 2.884, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  dotsBadge(s, 8.416, 2.011, 0.332, W, K)
  t(s, 3.846, 5.803, 4.835, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  glass(s, 9.612, 2.156, 2.765, 1.976, 0.329)
  rr(s, 9.948, 2.8, 0.23, 1.069, 0.115, W)
  gradRect(s, 9.948, 3.118, 0.23, 0.752, 0.115)
  rr(s, 10.292, 2.953, 0.23, 0.907, 0.115, W)
  gradRect(s, 10.292, 3.255, 0.23, 0.615, 0.115)
  rr(s, 10.635, 2.581, 0.23, 1.274, 0.115, W)
  gradRect(s, 10.635, 2.793, 0.23, 1.061, 0.115)
  t(s, 11.088, 3.417, 1.364, 0.252, 'There Are Many', { sz: 9 })
  t(s, 11.088, 2.997, 1.117, 0.438, '170K', { sz: 20, font: HEAD })
  dotsV(s, 12.11, 2.479, W)
  ln(s, 2.269, 5.352, 0, 4.891, { alpha: 51 })
  tick(s, 2.27, 5.351)
  flameBadge(s, 9.96, 5.947, 0.259, P, W)
  t(s, 10.411, 6.208, 1.743, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 10.395, 5.942, 2.016, 0.269, 'The Details', { sz: 10, font: HEAD })
}

/** Slide 21 — Financial Projections */
function slide21 (s) {
  rings(s, 1.704, 5.449)
  imageBox(s, 1.689, 4.379, 4.978, 2.145, 0.287)
  orb(s, -1.524, -1.834, 4.475)
  orb(s, 11.027, 5.346, 4.475)
  chrome(s)
  t(s, 0.817, 2.166, 3.061, 1.178, 'Financial Projections', { sz: 32, font: HEAD })
  t(s, 7.612, 5.986, 0.446, 0.33, '0', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 5.663, 0.446, 0.33, '1', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 5.34, 0.446, 0.33, '2', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 5.017, 0.446, 0.33, '3', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 4.695, 0.446, 0.33, '4', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 4.372, 0.446, 0.33, '5', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.612, 4.049, 0.446, 0.33, '6', { sz: 10, font: BODY, lh: 1.5 })
  t(s, 7.849, 6.357, 1.472, 0.304, '2020', { sz: 12, font: HEAD, align: 'center' })
  t(s, 9.488, 6.357, 1.472, 0.304, '2030', { sz: 12, font: HEAD, align: 'center' })
  t(s, 11.129, 6.357, 1.474, 0.304, '2040', { sz: 12, font: HEAD, align: 'center' })
  rr(s, 8.44, 5.359, 0.293, 0.851, 0.146, W)
  rr(s, 10.078, 4.59, 0.293, 1.62, 0.146, W)
  rr(s, 11.719, 4.846, 0.293, 1.363, 0.146, W)
  gradRect(s, 8.072, 4.914, 0.293, 1.293, 0.146)
  gradRect(s, 8.804, 5.173, 0.293, 1.043, 0.146)
  gradRect(s, 9.71, 5.571, 0.293, 0.636, 0.146)
  gradRect(s, 10.442, 5.077, 0.293, 1.131, 0.146)
  gradRect(s, 11.351, 5.26, 0.293, 0.948, 0.146)
  gradRect(s, 12.083, 5.561, 0.293, 0.647, 0.146)
  t(s, 4.789, 2.518, 4.13, 0.62, 'Aliquam cras imperdiet egestas, dui duis natoque Aliquam tincidunt risus nostra vulputate', { sz: 11, lh: 1.5 })
  flameBadge(s, 4.364, 2.698, 0.259, P, W)
  rr(s, 9.519, 2.156, 2.433, 1.508, 0.243, W)
  t(s, 9.697, 2.411, 1.845, 0.303, 'Operating Income', { sz: 12, font: HEAD, color: K })
  t(s, 9.697, 2.759, 2.255, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, color: K, lh: 1.5 })
  dotsBadge(s, 6.171, 4.531, 0.332, K, W)
  ln(s, 10.735, -0.843, 0, 3.98, { alpha: 51 })
  tick(s, 10.736, 2.157)
}

/** Slide 22 — Our Team */
function slide22 (s) {
  orb(s, 10.595, -1.818, 4.475)
  orb(s, -1.713, 5.663, 4.475)
  rings(s, 9.594, 5.008)
  ln(s, 7.879, 5.463, 0, 3.98, { alpha: 51 })
  ln(s, 3.495, -0.843, 0, 3.98, { alpha: 51 })
  chrome(s)
  t(s, 6.058, 2.52, 2.773, 0.64, 'Our Team', { sz: 32, font: HEAD })
  rr(s, 1.705, 2.115, 3.58, 1.441, 0.28, P)
  t(s, 3.04, 2.545, 1.363, 0.303, 'Jhon Smiths', { sz: 12, font: HEAD })
  t(s, 3.04, 2.818, 0.861, 0.32, 'Your Job ', { sz: 10, lh: 1.5 })
  flameBadge(s, 4.711, 2.705, 0.259, W, P)
  rr(s, 1.705, 4.362, 3.58, 1.441, 0.28, W)
  t(s, 3.04, 4.792, 1.363, 0.303, 'Jhon Smiths', { sz: 12, font: HEAD, color: K })
  t(s, 3.04, 5.065, 0.861, 0.32, 'Your Job ', { sz: 10, color: K, lh: 1.5 })
  flameBadge(s, 4.711, 4.953, 0.259, P, W)
  rr(s, 6.049, 4.872, 3.58, 1.441, 0.28, P)
  t(s, 7.424, 5.303, 1.363, 0.303, 'Jhon Smiths', { sz: 12, font: HEAD })
  t(s, 7.424, 5.576, 0.861, 0.32, 'Your Job ', { sz: 10, lh: 1.5 })
  flameBadge(s, 9.096, 5.463, 0.259, W, P)
  t(s, 6.058, 3.253, 4.13, 0.62, 'Aliquam cras imperdiet egestas, dui duis natoque Aliquam tincidunt risus nostra vulputate', { sz: 11, lh: 1.5 })
  tick(s, 7.88, 6.313)
  tick(s, 3.495, 2.111)
  flameBadge(s, 10.526, 3.327, 0.259, P, W)
  t(s, 10.977, 3.589, 1.445, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 10.961, 3.322, 1.258, 0.269, 'The Details', { sz: 10, font: HEAD })
  imageBox(s, 1.941, 2.405, 0.861, 0.861, 0.43)
  imageBox(s, 1.941, 4.652, 0.861, 0.861, 0.43)
  imageBox(s, 6.326, 5.162, 0.861, 0.861, 0.43)
}

/** Slide 23 — Risk And Mitigation */
function slide23 (s) {
  rings(s, 10.979, 2.796)
  imageBox(s, 5.753, 1.854, 5.227, 1.884, 0.279)
  orb(s, 10.595, -1.818, 4.475)
  orb(s, -1.713, 5.663, 4.475)
  chrome(s)
  rr(s, 1.924, 4.602, 2.681, 1.885, 0.304, W)
  t(s, 2.173, 4.966, 1.682, 0.303, 'Identifty Risk', { sz: 12, font: HEAD, color: K })
  t(s, 2.173, 5.314, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, color: K, lh: 1.5 })
  el(s, 3.893, 4.425, 0.385, 0.385, P)
  t(s, 3.855, 4.475, 0.462, 0.286, '01', { sz: 11, font: HEAD, align: 'center' })
  rr(s, 5.079, 4.602, 2.681, 1.885, 0.304, P)
  t(s, 5.328, 4.966, 1.682, 0.303, 'Control Risk', { sz: 12, font: HEAD })
  t(s, 5.328, 5.314, 2.391, 0.825, 'PLACEHOLDER', { sz: 10, lh: 1.5 })
  el(s, 7.049, 4.425, 0.385, 0.385, W)
  t(s, 7.01, 4.475, 0.462, 0.286, '02', { sz: 11, font: HEAD, color: K, align: 'center' })
  rr(s, 12.084, 4.602, 2.681, 1.885, 0.304, W)
  t(s, 8.551, 4.904, 3.1, 1.178, 'Risk And Mitigation', { sz: 32, font: HEAD })
  t(s, 2.166, 1.945, 2.958, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  t(s, 2.737, 3.457, 1.743, 0.269, 'Active Creator', { sz: 10 })
  t(s, 2.737, 3.037, 1.743, 0.438, '$659.041', { sz: 20, font: HEAD })
  flameBadge(s, 2.254, 3.045, 0.332, P, W)
  rr(s, -1.407, 1.842, 2.681, 1.885, 0.304, P)
  dotsBadge(s, 5.923, 2.039, 0.332, K, W)
}

/** Slide 24 — Product Vision */
function slide24 (s) {
  orb(s, -1.571, -1.638, 4.475)
  orb(s, 10.967, 4.998, 4.475)
  rings(s, 6.568, 5.778)
  chrome(s)
  t(s, 0.781, 2.108, 3.721, 0.64, 'Product Vision', { sz: 32, font: HEAD })
  rr(s, 6.403, 2.108, 6.328, 3.856, 0.396, P)
  t(s, 6.877, 2.804, 1.682, 0.303, 'First Vision ', { sz: 12, font: HEAD })
  t(s, 6.877, 3.152, 2.391, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  t(s, 9.884, 2.804, 1.682, 0.303, 'Second Vision', { sz: 12, font: HEAD })
  t(s, 9.884, 3.152, 2.391, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  t(s, 6.877, 4.322, 1.682, 0.303, 'Third Vision', { sz: 12, font: HEAD })
  t(s, 6.877, 4.67, 2.391, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  t(s, 9.884, 4.322, 1.682, 0.303, 'Fourth Vision', { sz: 12, font: HEAD })
  t(s, 9.884, 4.67, 2.391, 0.572, 'Est dignissim sollicitudin vitae semper ultricies maximus', { sz: 10, lh: 1.5 })
  dotsBadge(s, 12.207, 2.314, 0.332, W, K)
  t(s, 0.811, 2.952, 4.779, 0.62, 'PLACEHOLDER', { sz: 11, lh: 1.5 })
  flameBadge(s, 0.912, 4.156, 0.259, P, W)
  t(s, 1.362, 4.417, 1.445, 0.252, 'Lorem Ipsum Dolor', { sz: 9 })
  t(s, 1.347, 4.151, 1.258, 0.269, 'The Details', { sz: 10, font: HEAD })
  t(s, 0.839, 5.674, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 0.839, 5.254, 1.117, 0.438, '100K', { sz: 20, font: HEAD })
  t(s, 2.77, 5.674, 1.743, 0.269, 'There Are Many', { sz: 10 })
  t(s, 2.77, 5.254, 1.715, 0.438, '103K', { sz: 20, font: HEAD })
}

/** Slide 25 — Thank You */
function slide25 (s) {
  rings(s, 7.819, 2.595)
  ln(s, -1.159, 5.137, 13.535, 0, { alpha: 51 })
  phone(s, 7.714, 2.437, 3.808, 7.875)
  imageBox(s, 8.132, 3.243, 2.973, 3.218, 0.337)
  orb(s, 10.595, -1.818, 4.475)
  orb(s, -1.713, 5.663, 4.475)
  gradRect(s, 3.629, 4.945, 1.884, 0.38, 0.19)
  chrome(s)
  t(s, 0.827, 2.437, 5.351, 1.313, 'Thank You For Your Attention', { sz: 36, font: HEAD })
  rr(s, 0.918, 4.949, 1.884, 0.375, 0.188, W)
  t(s, 1.1, 4.949, 1.52, 0.32, '@socialmedia', { sz: 10, color: K, align: 'center', lh: 1.5 })
  t(s, 3.814, 4.949, 1.52, 0.32, '+(11) 123-456-789', { sz: 10, align: 'center', lh: 1.5 })
  rr(s, 6.347, 4.949, 1.884, 0.375, 0.188, W)
  t(s, 6.528, 4.949, 1.52, 0.32, 'www.example.com', { sz: 10, color: K, align: 'center', lh: 1.5 })
  tick(s, 2.802, 5.132, P)
  tick(s, 3.633, 5.132)
  tick(s, 5.515, 5.132)
  tick(s, 6.346, 5.132, P)
  tick(s, 12.418, 5.132)
  dotsBadge(s, 10.594, 3.418, 0.332, K, W)
}

/* ----------------------------------------------------------------- assembly */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25]

const pptx = new PptxGenJS()
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 })
pptx.layout = 'DECK'
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY }

BUILDERS.forEach(build => {
  const s = pptx.addSlide()
  s.background = { color: K }
  build(s)
})

pptx.writeFile({ fileName: path.join(__dirname, '01496c1e-c24a-46a6-b82c-26c684121673_grok_final.pptx') })
  .then(f => console.log('wrote', f))
