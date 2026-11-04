/**
 * ARTSTRAX — "Pop Art & Graffiti" deck (30 slides, 13.333in x 7.5in).
 *
 * Every slide is built by a `slideNN()` function below. Geometry is written as
 * `[x, y, w, h]` inch tuples, matching the source deck's shape offsets.
 * Photographs in the original are replaced by flat grey placeholder rectangles.
 */
'use strict'

const path = require('path')
const PptxGenJS = require('pptxgenjs')

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

const ORANGE = 'F1652D' // accent2
const YELLOW = 'FFDB01' // accent1
const DARK = '262626' // accent3 / text
const WHITE = 'FFFFFF'
const LIGHT = 'F2F2F2' // pale card fill
const GREY = '404040' // muted label text
const TRACK = 'D9D9D9' // progress-ring track
const PHOTO = 'BFBFBF' // stand-in for the deck's photographs

const HEAD = 'Montserrat' // display headings
const BODY = 'Open Sans' // body copy
const ALT = 'Work Sans' // pricing table / SWOT captions

const SLIDE_W = 13.333
const SLIDE_H = 7.5

/** Solid fill at a reduced opacity (OOXML `<a:alpha>`). */
const fade = (color, opacityPct) => ({ color, transparency: 100 - opacityPct })

/** 11pt justified body paragraph with 150% line spacing — used all over the deck. */
const P11 = { fontSize: 11, lineSpacingMultiple: 1.5, align: 'justify' }
/** 12pt bold sub-heading sitting above a P11 paragraph. */
const P12B = { fontSize: 12, bold: true, lineSpacingMultiple: 1.5, align: 'justify' }
/** 11pt centered caption with 150% line spacing. */
const P11C = { fontSize: 11, lineSpacingMultiple: 1.5, align: 'center' }

// Filler copy reused across the deck.
const LOREM = {
  short: 'Lorem ipsum dolor. ',
  vision: 'Lorem ipsum dolor sit amet, consectetuer elit. Maecenas porttmassa. Fusce posuere, ',
  mission: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed magna sed pulvinar osuere, magna pulvinar osuere, magna magna posuere, ',
  card: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. ',
  advantage: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar',
  compare: 'Lorem ipsum dolor sit amet, consecte tuer adipiscing elit. sed dolor sit amet, consecte tuer adipiscing elit. sed dolor sit amet, consecte tuer adipiscing elit. sed',
  strength: 'Lorem ipsum dolor sit amet, consecte tuer adipiscing elit. sed pulvinar amet, consecte tuer elit. sed pulvinar amet, consecte tuer elit. sed pulvinar amet, consecte tuer',
  threat: 'Lorem ipsum dolor dolor amet, consecte tuer',
  rating: 'Lorem ipsum dolor dolor ipsum dolor dolor ipsum dolor dolor amet, consecte tuer',
  service: 'Lorem ipsum dolor sita met, adipiscing elit. ',
  step: 'Lorem ipsum dolor sit amet, adipiscing elit. porttitor gue, ',
  creation: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, ',
  about: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere, magna sed pulvinar'
}

// ---------------------------------------------------------------------------
// Drawing helpers — every one takes a `[x, y, w, h]` inch tuple
// ---------------------------------------------------------------------------

const solid = fill => (typeof fill === 'string' ? { color: fill } : fill)
const at = geo => ({ x: geo[0], y: geo[1], w: geo[2], h: geo[3] })

function shape (slide, kind, geo, fill, opts = {}) {
  slide.addShape(kind, { ...at(geo), fill: solid(fill), ...opts })
}

function rect (slide, geo, fill, opts = {}) {
  shape(slide, 'rect', geo, fill, opts)
}

function text (slide, runs, geo, opts = {}) {
  slide.addText(runs, { ...at(geo), fontFace: BODY, color: DARK, valign: 'top', ...opts })
}

/** Filled rectangle with centered text inside (buttons, number chips, big letters). */
function badge (slide, geo, fill, runs, opts = {}) {
  text(slide, runs, geo, { fill: solid(fill), align: 'center', valign: 'middle', ...opts })
}

/**
 * The recurring two-tone heading: a light word followed by a bold accent word.
 * `stacked: true` puts the accent word on its own paragraph.
 */
function heading (slide, geo, light, accent, opts = {}) {
  const { light: lightColor = DARK, accent: accentColor = ORANGE, stacked = false, ...rest } = opts
  text(slide, [
    { text: light, options: { color: lightColor, breakLine: stacked } },
    { text: accent, options: { color: accentColor, bold: true } }
  ], geo, { fontFace: HEAD, fontSize: 36, ...rest })
}

/** The yellow "+" ornament (prstGeom `plus` with adj = 43791). */
const PLUS_ADJ = 0.43791
function plus (slide, x, y, size) {
  shape(slide, 'plus', [x, y, size, size], YELLOW, { rectRadius: PLUS_ADJ * size })
}

/**
 * Grey block standing in for a photograph. The source deck leaves its picture
 * placeholders empty, and PowerPoint paints those with the layout's BFBFBF fill,
 * so a plain rectangle reproduces the original appearance.
 */
function photo (slide, geo, opts = {}) {
  rect(slide, geo, PHOTO, opts)
}

/**
 * Page furniture shared by every slide: the "Artstrax" tag top-left and the two
 * colour tabs bleeding off the right edge.
 */
function chrome (slide, opts = {}) {
  const { tag = ORANGE, swapTabs = false } = opts
  text(slide, 'Artstrax', [0.427, 0.366, 0.812, 0.278], {
    fontSize: 10.5, bold: true, color: WHITE, fill: solid(tag), wrap: false
  })
  rect(slide, [13.188, 5.396, 0.146, 0.708], swapTabs ? ORANGE : YELLOW)
  rect(slide, [13.188, 6.104, 0.146, 0.708], swapTabs ? YELLOW : ORANGE)
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function slide01 (pptx) { // Cover
  const s = pptx.addSlide()
  photo(s, [0, 0, SLIDE_W, SLIDE_H])
  rect(s, [0, 0, SLIDE_W, SLIDE_H], fade(DARK, 35))
  rect(s, [3.188, 2.885, 6.958, 1.73], ORANGE)
  text(s, 'ARTSTRAX', [3.745, 3.089, 5.843, 1.313], {
    fontFace: HEAD, fontSize: 72, bold: true, color: WHITE, align: 'center', wrap: false
  })
  chrome(s)
  plus(s, 10.562, 1.555, 1.094)
  plus(s, 1.708, 4.862, 1.094)
  text(s, 'POP ART & GRAFFITI PRESENTATION', [3.979, 4.918, 5.416, 0.286], {
    fontSize: 11, color: WHITE, charSpacing: 6, align: 'center', wrap: false
  })
}

function slide02 (pptx) { // Welcome
  const s = pptx.addSlide()
  photo(s, [7.479, 0, 5.854, 7.521])
  chrome(s)
  heading(s, [1.162, 1.389, 3.852, 1.313], 'WELCOME TO', 'ARTSTRAX', { stacked: true, wrap: false })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere, magna sed pulvinar magna, magna sed pulvinar magna sed, magna posuere, magna sed pulvinar osuere, magna posuere, magna',
    [1.162, 3.781, 4.901, 2.322], P11)
  plus(s, 6.43, 0.996, 0.787)
  text(s, [
    { text: 'Description ', options: {} },
    { text: 'Here', options: { color: ORANGE } }
  ], [1.162, 3.271, 1.838, 0.404], P12B)
}

function slide03 (pptx) { // Agenda
  const s = pptx.addSlide()
  photo(s, [0, 0, 5.875, 6.104])
  chrome(s)
  heading(s, [7.412, 1.41, 4.742, 0.707], 'CONTENT ', 'TODAY', { wrap: false })
  plus(s, 6.184, 0.644, 0.787)
  rect(s, [5.875, 2.75, 6.375, 4.75], ORANGE)

  // [label, label width, column x, row y]
  const agenda = [
    ['01. About Us', 1.491, 0, 0], ['02. Content', 1.491, 1, 0],
    ['03. Our Target', 1.491, 0, 1], ['04. Portfolio', 1.647, 1, 1],
    ['05. Pricing Table', 1.73, 0, 2], ['06. Contact Us', 1.647, 1, 2]
  ]
  const colX = [[6.977, 6.971], [9.424, 9.418]]
  const rowY = [[3.505, 3.818], [4.779, 5.092], [6.054, 6.367]]
  agenda.forEach(([label, w, col, row]) => {
    text(s, label, [colX[col][0], rowY[row][0], w, 0.372], { ...P12B, color: WHITE })
    text(s, LOREM.short, [colX[col][1], rowY[row][1], 1.736, 0.349], { ...P11, color: WHITE })
  })
}

function slide04 (pptx) { // About Artstrax
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [1.162, 1.41, 3.028, 1.313], 'ABOUT', 'ARTSTRAX', { stacked: true, wrap: false })
  text(s, LOREM.about, [5.282, 1.393, 4.312, 2.045], P11)
  plus(s, 10.688, 5.317, 0.787)
  text(s, LOREM.about, [5.282, 4.053, 4.312, 2.045], P11)
  photo(s, [0, 3.489, 4.189, 4.011])
  photo(s, [10.688, 0, 2.646, 4.688])
}

function slide05 (pptx) { // Fast-spreading culture
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [7.474, 1.389, 4.943, 3.736], 'GRAFFITI ART AND POP ART HAVE BECOME A ', 'FAST-SPREADING CULTURE AND CULTURE')
  plus(s, 6.281, 0.623, 0.787)
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce',
    [7.474, 5.455, 4.713, 0.656], P11)
  rect(s, [0, 5.021, 2.479, 2.479], ORANGE)
  text(s, 'Lorem ipsum dolor sit', [0.508, 6.208, 1.464, 0.625], { ...P11C, color: WHITE })
  text(s, '678+', [0.508, 5.688, 1.464, 0.64], { fontSize: 32, bold: true, color: WHITE, align: 'center' })
  photo(s, [0, 1.41, 2.479, 3.361])
  photo(s, [2.75, 0, 3.292, 7.521])
}

function slide06 (pptx) { // Difference between graffiti and pop art
  const s = pptx.addSlide()
  photo(s, [0, 0, SLIDE_W, 4.062])
  rect(s, [0, 0, SLIDE_W, 4.062], fade(DARK, 80))
  chrome(s)
  heading(s, [2.774, 1.116, 7.81, 1.313], 'DIFFERENCE BETWEEN ', 'GRAFFITI AND POP ART', {
    light: WHITE, align: 'center'
  })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere',
    [2.051, 2.64, 9.241, 0.656], { ...P11C, color: WHITE })

  // Two comparison cards: [card x, card fill, number x, copy x, ink]
  const cards = [[1.198, LIGHT, 1.694, 2.67, DARK], [6.672, ORANGE, 7.169, 8.144, WHITE]]
  cards.forEach(([cardX, fill, numX, copyX, ink], i) => {
    rect(s, [cardX, 4.562, 5.474, 1.917], fill)
    text(s, i === 0 ? '01' : '02', [numX, 4.907, 1.015, 0.774], {
      fontSize: 40, bold: true, color: i === 0 ? ORANGE : WHITE
    })
    text(s, LOREM.compare, [copyX, 4.923, 3.505, 1.212], { ...P11, color: ink })
  })
  plus(s, 11.359, 1.232, 0.787)
  plus(s, 1.193, 1.232, 0.787)
}

/** The orange "MISSION" banner rotated to read bottom-up / top-down. */
function missionBanner (slide, geo, rotate) {
  badge(slide, geo, ORANGE, 'MISSION', {
    fontSize: 44, bold: true, color: WHITE, charSpacing: 6, rotate
  })
}

function slide07 (pptx) { // Our vision
  const s = pptx.addSlide()
  photo(s, [1.583, 0, 4.292, 6.104])
  chrome(s)
  heading(s, [7.453, 1.264, 3.386, 0.707], 'OUR ', 'VISION', { wrap: false })
  plus(s, 5.825, 6.104, 0.787)
  missionBanner(s, [-1.635, 4.281, 4.854, 1.583], 270)

  const cols = [7.474, 10.203]
  const rows = [[2.417, 2.805], [4.53, 4.918]]
  let n = 0
  rows.forEach(([labelY, copyY]) => cols.forEach(x => {
    n += 1
    text(s, `Vision 0${n}`, [x, labelY, 1.964, n === 1 ? 0.404 : 0.372], P12B)
    text(s, LOREM.vision, [x, copyY, 1.964, 1.212], P11)
  }))
}

function slide08 (pptx) { // Our mission
  const s = pptx.addSlide()
  photo(s, [6.667, 1.417, 5.083, 6.083])
  chrome(s)
  heading(s, [1.162, 1.368, 3.813, 0.707], 'OUR ', 'MISSION', { wrap: false })
  plus(s, 6.713, 1.079, 0.787)
  missionBanner(s, [10.115, 1.635, 4.854, 1.583], 90)

  ;[[2.495, 2.883, 0.404], [4.514, 4.902, 0.417]].forEach(([labelY, copyY, labelH], i) => {
    text(s, `Mission 0${i + 1}`, [1.239, labelY, 1.964, labelH], P12B)
    text(s, LOREM.mission, [1.239, copyY, 3.99, 1.212], P11)
  })
}

function slide09 (pptx) { // About street art
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [3.545, 1.137, 6.268, 0.707], 'ABOUT ', 'STREET ART', { align: 'center' })

  // [card x, fill, ink, title, title x, title w, copy x, copy h]
  const panels = [
    [1.375, LIGHT, DARK, 'Street Art Description', 2.555, 0.404, 1.918, 0.656],
    [6.904, ORANGE, WHITE, 'Graffity Description', 8.083, 0.372, 7.447, 0.627]
  ]
  panels.forEach(([x, fill, ink, title, titleX, titleH, copyX, copyH]) => {
    rect(s, [x, 4.5, 5.075, 1.667], fill)
    text(s, title, [titleX, 4.796, 2.716, titleH], { ...P12B, color: ink, align: 'center' })
    text(s, LOREM.card, [copyX, 5.173, 3.99, copyH], { ...P11C, color: ink })
  })
  plus(s, 0.648, 1.596, 0.787)
  plus(s, 11.901, 1.596, 0.787)
  photo(s, [1.375, 2.271, 5.075, 2.229])
  photo(s, [6.904, 2.271, 5.075, 2.229])
}

function slide10 (pptx) { // Favorite type of pop art in 2022
  const s = pptx.addSlide()
  rect(s, [7.543, 3.75, 4.913, 3.75], ORANGE)
  heading(s, [1.057, 1.347, 3.88, 1.919], 'FAVORITE TYPE OF POP ', 'ART IN 2022')
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere massa. Fusce, magna sed pulvinar osuere, ',
    [1.057, 4.947, 4.165, 1.212], P11)
  plus(s, 5.898, 0.644, 0.787)

  // [x, value, value w, caption, caption w, value colour]
  const stats = [
    [1.057, '98k+', 1.334, 'Art Totals', 1.334, DARK],
    [2.508, '7.96', 1.066, 'Rate Point', 1.185, DARK],
    [3.871, '93%', 1.066, 'Percentage', 1.185, ORANGE]
  ]
  stats.forEach(([x, value, vw, caption, cw, color]) => {
    text(s, value, [x, 3.892, vw, 0.572], { fontSize: 28, bold: true, color, align: 'justify' })
    text(s, caption, [x, 4.417, cw, 0.286], { fontSize: 11, align: 'justify' })
  })
  photo(s, [6.667, 1.417, 2.155, 4.688])
  photo(s, [8.922, 0.729, 2.155, 4.688])
  photo(s, [11.178, 1.417, 2.155, 4.688])
  chrome(s)
}

function slide11 (pptx) { // Advantages of our graffiti works
  const s = pptx.addSlide()
  photo(s, [0, 0, 7.474, SLIDE_H])
  rect(s, [0, 0, 7.474, SLIDE_H], fade(DARK, 80))
  chrome(s)
  heading(s, [8.474, 1.347, 3.852, 3.13], 'THE ADVANTAGES OF OUR ', 'GRAFFITI WORKS')
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue',
    [8.474, 4.859, 3.852, 1.212], P11)
  plus(s, 6.3, 0.387, 0.787)

  // [chip y, copy y, chip fill, chip ink]
  const items = [
    [1.411, 1.368, YELLOW, DARK], [2.704, 2.662, ORANGE, WHITE],
    [3.998, 3.955, YELLOW, DARK], [5.291, 5.249, ORANGE, WHITE]
  ]
  items.forEach(([chipY, copyY, fill, ink], i) => {
    badge(s, [1.177, chipY, 0.819, 0.819], fill, `0${i + 1}`, { fontSize: 20, bold: true, color: ink })
    text(s, LOREM.advantage, [2.344, copyY, 3.852, 0.904], { ...P11, color: WHITE })
  })
}

function slide12 (pptx) { // Creative and innovative graffiti
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [1.078, 1.41, 3.588, 2.524], 'CREATIVE AND ', 'INNOVATIVE GRAFFITI')
  text(s, LOREM.creation, [8.948, 5.125, 3.219, 0.934], P11)
  plus(s, 12.4, 1.016, 0.787)
  text(s, LOREM.creation, [5.25, 5.125, 3.219, 0.934], P11)
  photo(s, [0, 4.562, 4.229, 2.938])
  photo(s, [5.25, 0, 6.917, 4.562])
}

function slide13 (pptx) { // Meet greatest artist
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [3.003, 1.241, 7.351, 0.707], 'MEET GREATEST ', 'ARTIST', { align: 'center' })

  // [name, name x, role x, photo x]
  const artists = [
    ['Richard Sanchez', 1.975, 2.17, 1.604],
    ['Rachelle Beaudry', 5.684, 5.879, 5.313],
    ['Pedro Fernandes', 9.392, 9.588, 9.022]
  ]
  artists.forEach(([name, nameX, roleX, photoX]) => {
    text(s, 'Graffiti Artist', [roleX, 5.818, 1.575, 0.286], { fontSize: 11, align: 'center' })
    text(s, name, [nameX, 5.512, 1.966, 0.303], { fontSize: 12, bold: true, color: ORANGE, align: 'center' })
    photo(s, [photoX, 2.594, 2.707, 2.707])
  })
  plus(s, 0.648, 1.846, 0.787)
  plus(s, 11.901, 1.846, 0.787)
}

function slide14 (pptx) { // Meet our artist
  const s = pptx.addSlide()

  // [tab x, tab fill, name, name colour, role colour, photo x, photo y]
  const artists = [
    [5.292, ORANGE, 'Noah Schumacher', WHITE, WHITE, 5.292, 1.423],
    [8.016, LIGHT, 'Margarita Perez', ORANGE, DARK, 8.016, 1.415],
    [10.741, LIGHT, 'Taylor Alonso', ORANGE, DARK, 10.741, 1.406]
  ]
  artists.forEach(([x, fill]) => rect(s, [x, 5.292, 2.592, 0.812], fill))
  chrome(s)
  heading(s, [1.162, 1.41, 3.213, 1.313], 'MEET OUR ', 'ARTIST')
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue',
    [1.162, 3.197, 3.213, 2.045], P11)
  plus(s, 4.505, 0.619, 0.787)
  artists.forEach(([x, , name, nameColor, roleColor, photoX, photoY]) => {
    text(s, 'Graffiti Artist', [x + 0.508, 5.702, 1.575, 0.286], { fontSize: 11, color: roleColor, align: 'center' })
    text(s, name, [x + 0.313, 5.438, 1.966, 0.303], { fontSize: 12, bold: true, color: nameColor, align: 'center' })
    photo(s, [photoX, photoY, 2.592, 3.698])
  })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer, magna posuere congue', [1.162, 5.436, 3.213, 0.656], P11)
}

function slide15 (pptx) { // Founder profile
  const s = pptx.addSlide()
  rect(s, [0, 1.079, 3.721, 4.225], ORANGE)
  photo(s, [1.188, 1.417, 4.688, 6.083])
  chrome(s)
  heading(s, [7.474, 1.41, 4.681, 0.707], 'JAMIE ', 'CHASTAIN', { wrap: false })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar osue',
    [7.474, 3.843, 4.755, 1.212], P11)
  plus(s, 5.817, 0.686, 0.787)
  badge(s, [7.578, 2.254, 2.943, 0.559], LIGHT, 'FOUNDER ARTSTRAX', { fontSize: 12, bold: true, charSpacing: 3 })
  badge(s, [7.578, 5.396, 2.068, 0.688], ORANGE, 'MORE PROFILE', { fontSize: 12, bold: true, color: WHITE })
  text(s, [
    { text: 'Profile ', options: {} },
    { text: 'Description', options: { color: ORANGE } }
  ], [7.474, 3.369, 4.755, 0.372], P12B)
}

function slide16 (pptx) { // Break session
  const s = pptx.addSlide()
  photo(s, [0, 0, SLIDE_W, SLIDE_H])
  rect(s, [0, 0, SLIDE_W, SLIDE_H], fade(DARK, 35))
  rect(s, [0, 2.622, SLIDE_W, 2.256], ORANGE)
  text(s, 'BREAK SESSION', [1.306, 2.959, 10.72, 1.582], {
    fontFace: HEAD, fontSize: 88, bold: true, color: WHITE, align: 'center', wrap: false
  })
  chrome(s)
  plus(s, 11.041, 0.947, 1.094)
  plus(s, 1.224, 5.494, 1.094)
  text(s, 'BREAK UNTIL 07.00 PM', [3.959, 5.223, 5.416, 0.37], {
    fontSize: 16, color: WHITE, charSpacing: 10, align: 'center', wrap: false
  })
}

/** SWOT slides 17-20 all sit on the same orange band + dark photo panel + giant letter. */
function swotBackdrop (slide, letter, side) {
  const left = side === 'left'
  rect(slide, [left ? 0 : 9.021, 0, 4.312, SLIDE_H], ORANGE)
  const panel = [left ? 0 : 7.458, 1.417, 5.875, 4.688]
  photo(slide, panel)
  badge(slide, panel, fade(DARK, 80), letter, { fontSize: 199, bold: true, color: YELLOW })
  plus(slide, left ? 5.817 : 6.729, 0.686, 0.787)
}

function slide17 (pptx) { // Strengths
  const s = pptx.addSlide()
  swotBackdrop(s, 'S', 'left')
  chrome(s, { tag: fade(WHITE, 15) })
  heading(s, [7.474, 1.41, 4.194, 1.313], 'OUR ARTSTRAX', 'STRENGTHS', { stacked: true, wrap: false })
  ;[[3.204, 3.22, YELLOW, 1.212], [4.825, 4.842, ORANGE, 1.182]].forEach(([numY, copyY, color, h], i) => {
    text(s, `0${i + 1}`, [7.474, numY, 1.015, 0.707], { fontSize: 36, bold: true, color })
    text(s, LOREM.strength, [8.454, copyY, 3.73, h], P11)
  })
}

function slide18 (pptx) { // Weaknesses
  const s = pptx.addSlide()
  swotBackdrop(s, 'W', 'right')
  chrome(s, { swapTabs: true })
  heading(s, [1.078, 1.41, 4.194, 1.313], 'OUR ARTSTRAX', 'WEAKNESSES', { stacked: true, wrap: false })

  // Donut gauges: [ring x, ring y, arc colour, arc start angle, value, value x, value y, value w, label x, label y]
  const gauges = [
    [1.197, 3.345, ORANGE, 334.741, '85%', 1.345, 3.634, 0.617, 1.078, 4.386],
    [2.536, 3.34, YELLOW, 29.658, '71%', 2.703, 3.628, 0.579, 2.396, 4.382],
    [3.888, 3.335, ORANGE, 353.72, '76%', 4.036, 3.623, 0.617, 3.727, 4.378]
  ]
  gauges.forEach(([x, y, color, start, value, vx, vy, vw, lx, ly]) => {
    shape(s, 'ellipse', [x, y, 0.913, 0.913], null, { line: { color: TRACK, width: 3 }, flipH: true })
    shape(s, 'arc', [x, y, 0.913, 0.913], null, {
      line: { color, width: 3 }, angleRange: [start, 270.207], flipH: true
    })
    text(s, value, [vx, vy, vw, 0.337], {
      fontFace: 'Work Sans SemiBold', fontSize: 14, color: GREY, align: 'center', valign: 'middle', wrap: false
    })
    text(s, 'Weaknesses', [lx, ly, 1.193, 0.325], { fontFace: ALT, ...P11C, fontSize: 10, color: GREY })
  })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed pulvinar osuere, magna magna posuere, magna',
    [1.078, 5.067, 4.715, 0.934], { ...P11, fontFace: ALT })
}

function slide19 (pptx) { // Opportunities
  const s = pptx.addSlide()
  swotBackdrop(s, 'O', 'left')
  chrome(s, { tag: fade(WHITE, 15) })
  heading(s, [7.474, 1.41, 4.586, 1.313], 'OUR ARTSTRAX', 'OPPORTUNITIES', { stacked: true, wrap: false })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed pulvinar osuere, magna magna posuere, magna sed pulvinar osuere posuere, magna sed pulvinar osuere posuere, magna sed, magna magna sed pulvinar posuere, ',
    [7.48, 3.253, 4.715, 1.489], P11)

  // Horizontal bars: [label, percent text, filled width, bar colour, label y, bar y]
  const bars = [
    ['In 2021', '89%', 2.797, YELLOW, 5.148, 5.239],
    ['In 2022', '78%', 2.259, ORANGE, 5.82, 5.911]
  ]
  bars.forEach(([label, pct, filled, color, labelY, barY]) => {
    text(s, pct, [11.596, labelY + 0.005, 0.821, 0.286], { fontSize: 11, align: 'justify' })
    rect(s, [8.381, barY, 3.151, 0.103], fade(PHOTO, 20))
    rect(s, [8.38, barY, filled, 0.103], color)
    text(s, label, [7.48, labelY, 0.821, 0.286], { fontSize: 11, align: 'justify' })
  })
}

function slide20 (pptx) { // Threats
  const s = pptx.addSlide()
  swotBackdrop(s, 'T', 'right')
  chrome(s, { swapTabs: true })
  heading(s, [1.078, 1.41, 4.194, 1.313], 'OUR ARTSTRAX', 'THREATS', { stacked: true, wrap: false })

  // [number x, number y, colour, copy h]
  const threats = [
    [1.078, 3.052, YELLOW, 0.656], [3.845, 3.052, ORANGE, 0.627],
    [1.078, 4.766, ORANGE, 0.627], [3.845, 4.766, YELLOW, 0.627]
  ]
  threats.forEach(([x, y, color, copyH], i) => {
    text(s, `0${i + 1}`, [x, y, 1.015, 0.707], { fontSize: 36, bold: true, color })
    text(s, LOREM.threat, [x + 0.009, y + 0.654, 2.038, copyH], P11)
  })
}

// --- Small pictogram stand-ins used on slides 21 and 28 ---------------------

function iconPerson (slide, x, y, w, h, color) {
  shape(slide, 'ellipse', [x + w * 0.29, y, w * 0.42, w * 0.42], color)
  shape(slide, 'roundRect', [x, y + h * 0.5, w, h * 0.5], color, { rectRadius: 0.05 })
}

function iconBulb (slide, x, y, w, h, color) {
  shape(slide, 'ellipse', [x, y, w, w * 1.05], null, { line: { color, width: 2.25 } })
  rect(slide, [x + w * 0.32, y + h * 0.78, w * 0.36, h * 0.22], color)
}

function iconPicture (slide, x, y, w, h, color) {
  rect(slide, [x, y, w, h], color)
  rect(slide, [x + w * 0.16, y + h * 0.38, w * 0.68, h * 0.36], WHITE)
  shape(slide, 'triangle', [x + w * 0.24, y + h * 0.48, w * 0.4, h * 0.26], color)
}

function iconGears (slide, x, y, w, h, color) {
  shape(slide, 'donut', [x, y, w * 0.63, w * 0.63], color)
  shape(slide, 'donut', [x + w * 0.5, y + h * 0.4, w * 0.48, w * 0.48], color)
}

function iconInbox (slide, x, y, w, h, color) {
  shape(slide, 'trapezoid', [x, y + h * 0.44, w, h * 0.56], color, { flipV: true })
  shape(slide, 'downArrow', [x + w * 0.35, y, w * 0.3, h * 0.6], color)
}

function slide21 (pptx) { // Best service for client
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [3.003, 1.427, 7.351, 0.707], 'BEST SERVICE ', 'FOR CLIENT', { align: 'center' })
  plus(s, 0.648, 1.387, 0.787)
  plus(s, 11.901, 1.387, 0.787)

  // [card x, card fill, tile fill, ink, icon, icon geometry]
  const services = [
    [1.156, LIGHT, YELLOW, DARK, iconPerson, [2.384, 3.61, 0.305, 0.338]],
    [3.917, ORANGE, fade(WHITE, 15), WHITE, iconBulb, [5.166, 3.586, 0.262, 0.386]],
    [6.677, LIGHT, YELLOW, DARK, iconPicture, [7.891, 3.582, 0.333, 0.394]],
    [9.438, ORANGE, fade(WHITE, 15), WHITE, iconGears, [10.605, 3.585, 0.426, 0.389]]
  ]
  services.forEach(([x, cardFill, tileFill, ink, icon, geo], i) => {
    rect(s, [x, 2.812, 2.76, 3.155], cardFill)
    rect(s, [x + 0.91, 3.309, 0.941, 0.941], tileFill)
    text(s, `Service 0${i + 1}`, [x + 0.397, 4.472, 1.966, 0.303], {
      fontSize: 12, bold: true, color: ink, align: 'center'
    })
    text(s, LOREM.service, [x + 0.397, 4.814, 1.966, 0.627], { ...P11C, color: ink })
    icon(s, geo[0], geo[1], geo[2], geo[3], i % 2 === 0 ? DARK : WHITE)
  })
}

function slide22 (pptx) { // The best service for graffiti creation
  const s = pptx.addSlide()
  chrome(s)
  plus(s, 7.081, 0.561, 0.787)

  // [band y, band width, band fill, chip fill, chip ink, copy width]
  const steps = [
    [1.348, 3.875, LIGHT, YELLOW, DARK, 1.772],
    [2.958, 3.875, ORANGE, fade(WHITE, 15), WHITE, 1.772],
    [4.569, 6.667, LIGHT, YELLOW, DARK, 4.603]
  ]
  steps.forEach(([y, bandW, bandFill, chipFill, ink, copyW], i) => {
    rect(s, [0, y, bandW, 1.61], bandFill)
    badge(s, [0.407, y + 0.335, 0.941, 0.941], chipFill, `0${i + 1}`, { fontSize: 18, bold: true, color: ink })
    text(s, i === 2
      ? 'Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor con gue massa. Fusce pos uere pos uere, magna massa. Fusce pos uere, magna sed pul'
      : LOREM.step, [1.689, y + 0.338, copyW, 0.934], { ...P11, color: ink })
  })
  heading(s, [8.232, 1.372, 4.078, 2.524], 'THE BEST SERVICE FOR ', 'GRAFFITI CREATION')
  text(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor con gue massa. Fusce pos uere, magna sed pul vinar osuere, ',
    [10.271, 4.428, 1.979, 1.767], P11)
  photo(s, [3.875, 0, 2.792, 4.569])
  photo(s, [6.667, 4.569, 2.792, 2.96])
}

function slide23 (pptx) { // Our client ratings
  const s = pptx.addSlide()
  photo(s, [9.042, 0, 4.292, SLIDE_H])
  rect(s, [9.042, 0, 4.292, SLIDE_H], fade(DARK, 80))
  chrome(s)
  heading(s, [1.037, 1.41, 3.526, 1.313], 'OUR CLIENT ', 'RATINGS')
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce massa. Fusce posuere, magna sed',
    [1.037, 3.322, 3.234, 1.767], P11)
  ;[['89%', YELLOW, 0.927], ['77%', ORANGE, 2.956], ['94%', YELLOW, 4.984]].forEach(([pct, color, y]) => {
    text(s, pct, [9.938, y, 1.45, 0.707], { fontSize: 36, bold: true, color })
    text(s, LOREM.rating, [9.947, y + 0.655, 2.491, 0.934], { ...P11, color: WHITE })
  })
  badge(s, [1.114, 5.417, 2.068, 0.688], ORANGE, 'RATE NOW', { fontSize: 12, bold: true, color: WHITE })
  photo(s, [5.167, 0, 3.875, SLIDE_H])
}

/**
 * The pair of yellow slanted slabs used as an opening quotation mark on slide 24.
 * Point coordinates are inches relative to the shape's top-left corner.
 */
const QUOTE_MARK = [
  { x: 0.8278, y: 0 },
  { x: 1.092, y: 0 },
  { x: 0.864, y: 0.586, curve: { type: 'cubic', x1: 1.0164, y1: 0.177, x2: 0.9404, y2: 0.3723 } },
  { x: 0.5308, y: 0.586 },
  { x: 0.5215, y: 0.5603 },
  { x: 0.8278, y: 0, curve: { type: 'cubic', x1: 0.6017, y1: 0.3833, x2: 0.7038, y2: 0.1965 } },
  { close: true },
  { x: 0.3063, y: 0, moveTo: true },
  { x: 0.5694, y: 0 },
  { x: 0.3426, y: 0.586, curve: { type: 'cubic', x1: 0.4985, y1: 0.1661, x2: 0.4228, y2: 0.3614 } },
  { x: 0.0094, y: 0.586 },
  { x: 0, y: 0.5603 },
  { x: 0.3063, y: 0, curve: { type: 'cubic', x1: 0.0803, y1: 0.3833, x2: 0.1824, y2: 0.1965 } },
  { close: true }
]

function slide24 (pptx) { // Pull quote
  const s = pptx.addSlide()
  photo(s, [0, 0, SLIDE_W, SLIDE_H])
  rect(s, [0, 0, SLIDE_W, SLIDE_H], fade(DARK, 80))
  chrome(s)
  plus(s, 11.654, 1.254, 1.094)
  plus(s, 0.624, 5.164, 1.094)
  text(s, [
    { text: 'Pop art and graffiti are abstract works that contain ', options: { color: WHITE } },
    { text: 'criticism, expression, and beauty.', options: { color: YELLOW, bold: true } }
  ], [2.646, 2.492, 8.042, 2.524], {
    fontSize: 32, italic: true, align: 'center', lineSpacingMultiple: 1.5
  })
  text(s, '- Alexander Aronowitz -', [4.531, 5.458, 4.271, 0.404], {
    fontSize: 18, color: WHITE, align: 'center'
  })
  shape(s, 'custGeom', [5.513, 1.456, 1.092, 0.586], YELLOW, { points: QUOTE_MARK })
  shape(s, 'custGeom', [6.729, 1.456, 1.092, 0.586], YELLOW, { points: QUOTE_MARK, rotate: 180 })
}

function slide25 (pptx) { // Portfolio divider
  const s = pptx.addSlide()
  const tiles = [[1.083, 1.312], [6.771, 1.312], [1.083, 3.854], [6.771, 3.854]]
  tiles.forEach(([x, y]) => photo(s, [x, y, 5.479, 2.333]))
  chrome(s)
  rect(s, [3.188, 2.833, 6.958, 1.833], ORANGE)
  heading(s, [3.827, 3.239, 5.68, 0.707], 'PORTFOLIO ', 'SECTION', {
    light: WHITE, accent: WHITE, align: 'center', wrap: false
  })
  text(s, '238K+ TOTAL ART GRAFFITI', [3.625, 3.921, 6.083, 0.379], {
    ...P11C, color: WHITE, charSpacing: 9
  })
}

function slide26 (pptx) { // Portfolio section
  const s = pptx.addSlide()
  const tiles = [[5.8, 0.01], [9.7, 0], [5.8, 3.867], [9.7, 3.856]]
  tiles.forEach(([x, y]) => photo(s, [x, y, 3.633, 3.633]))
  chrome(s)
  heading(s, [1.057, 1.41, 3.261, 1.313], 'PORTFOLIO', 'SECTION', { stacked: true, wrap: false })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar, magna sed pulvinar osuere, ',
    [1.057, 3.239, 3.638, 1.767], P11)
  plus(s, 4.695, 0.623, 0.787)
  badge(s, [1.158, 5.375, 2.068, 0.688], ORANGE, 'MORE PICTURE', { fontSize: 12, bold: true, color: WHITE })
}

function slide27 (pptx) { // Pricing table
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [3.003, 1.408, 7.351, 0.707], 'OUR PRICING ', 'TABLE', { align: 'center' })
  plus(s, 0.648, 1.262, 0.787)
  plus(s, 11.901, 1.262, 0.787)

  // [card x, card w, card fill, ink, ribbon label, ribbon x, ribbon w, accent, price]
  const plans = [
    [0.997, 2.69, LIGHT, DARK, 'Regular', 1.001, 0.847, YELLOW, '$1500'],
    [3.889, 2.69, ORANGE, WHITE, 'Medium', 3.88, 0.894, fade(WHITE, 15), '$2500'],
    [6.781, 2.69, LIGHT, DARK, 'Premium', 6.711, 0.975, YELLOW, '$3000'],
    [9.72, 2.642, ORANGE, WHITE, 'Platinum', 9.633, 0.97, fade(WHITE, 15), '$4500']
  ]
  plans.forEach(([x, w, cardFill, ink, ribbon, ribbonX, ribbonW, accent, price]) => {
    rect(s, [x, 2.728, w, 3.384], cardFill)
    badge(s, [x + 0.25, 5.25, 2.198, 0.542], accent, 'REGISTER NOW', {
      fontFace: ALT, fontSize: 11, bold: true, color: ink
    })
    text(s, price, [x + 0.606, 3.555, 1.465, 0.572], {
      fontFace: ALT, fontSize: 28, bold: true, color: ink, align: 'center'
    })
    ;['001', '002'].forEach((n, row) => {
      const y = 4.305 + row * 0.331
      text(s, `Get Service Class ${n}`, [x + 0.565, y, 1.883, 0.278], { fontFace: ALT, fontSize: 10.5, color: ink })
      rect(s, [x + 0.46, y + 0.095, 0.087, 0.087], ink === DARK ? YELLOW : WHITE)
    })
    shape(s, 'diagStripe', [x, 2.73, 1.241, 1.241], accent, { rectRadius: 0.3842 * 1.241 })
    text(s, ribbon, [ribbonX, 3.009, ribbonW, 0.303], {
      fontSize: 12, bold: true, color: ink, rotate: 315, wrap: false
    })
    text(s, 'Get Price', [x + 0.404, 3.344, 1.883, 0.278], {
      fontFace: ALT, fontSize: 10.5, color: ink, align: 'center'
    })
  })
}

function slide28 (pptx) { // Infographic about sales history
  const s = pptx.addSlide()
  chrome(s)
  heading(s, [1.078, 1.327, 4.443, 1.919], 'INFOGRAPHIC ABOUT SALES', 'HISTORY', { stacked: true })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar sed pulvinar osuere, posuere, magna sed pulvinar',
    [1.078, 3.781, 3.338, 2.322], P11)

  // Drawn back-to-front so the longest (2022) bar sits underneath.
  // [bar x, bar w, bar y, fill, ink, percent, year]
  const bars = [
    [5.193, 6.527, 4.993, DARK, WHITE, '  99%', '2022'],
    [6.414, 5.306, 3.309, ORANGE, WHITE, '  71%', '2021'],
    [7.435, 4.284, 1.625, YELLOW, DARK, '  43%', '2020']
  ]
  const flushMargin = [7.2, 7.2, 0, 0] // lIns/rIns 0.1in, no vertical inset
  bars.forEach(([x, w, y, fill, ink, pct]) => {
    text(s, pct, [x, y, w, 0.89], {
      fill: solid(fill), fontSize: 18, bold: true, color: ink, valign: 'middle', margin: flushMargin
    })
  })
  bars.forEach(([, , y, , ink, , year]) => {
    text(s, `${year} Percentage`, [8.675, y + 0.327, 1.731, 0.236], {
      fontSize: 14, color: ink, align: 'right', valign: 'middle', wrap: false, margin: flushMargin
    })
  })
  // Outlined squares with an inbox pictogram, one per year.
  ;[[1.341, YELLOW], [3.025, ORANGE], [4.709, DARK]].forEach(([y, color]) => {
    rect(s, [10.678, y, 1.483, 1.457], WHITE, { line: { color, width: 3 } })
    iconInbox(s, 11.153, y + 0.495, 0.533, 0.46, color)
  })
}

function slide29 (pptx) { // Our contact
  const s = pptx.addSlide()
  photo(s, [0, 0, 4.708, 7.521])
  rect(s, [2.896, 2.583, 9.333, 3.5], ORANGE)
  chrome(s)
  heading(s, [6.662, 1.41, 4.401, 0.707], 'OUR ', 'CONTACT')
  plus(s, 5.127, 0.644, 0.787)

  // [column x, heading width, value width, [top heading, top value], [bottom heading, bottom value]]
  const columns = [
    [3.739, 1.05, 2.338, ['Website', 'www.nameyourwebsite.com'], ['Office Address', 'City 001122, Country 001122'], 1.697],
    [6.683, 1.455, 1.959, ['Telephone', '+999 0000 0000 0000'], ['Social Media', '@nameyoursosmed'], 1.455],
    [9.248, 1.455, 1.959, ['Email', 'yourmail@email.com'], ['Portfolio', 'www.yourportfolio.com'], 1.455]
  ]
  columns.forEach(([x, headW, valW, top, bottom, bottomHeadW]) => {
    ;[[top, 3.307, headW], [bottom, 4.533, bottomHeadW]].forEach(([[title, value], y, hw]) => {
      text(s, title, [x, y, hw, 0.303], { fontSize: 12, bold: true, color: WHITE })
      text(s, value, [x, y + 0.27, valW, 0.286], { fontSize: 11, color: WHITE })
      text(s, value, [x, y + 0.54, valW, 0.286], { fontSize: 11, color: WHITE })
    })
  })
}

function slide30 (pptx) { // Thank you
  const s = pptx.addSlide()
  photo(s, [0, 4.215, SLIDE_W, 3.285])
  chrome(s)
  plus(s, 11.442, 0.623, 0.787)
  heading(s, [2.461, 1.241, 8.435, 1.582], 'THANK ', 'YOU', { fontSize: 88, align: 'center' })
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna sed, magna posuere congue massa. Fusce posuere, magna sed pulvinar',
    [1.891, 2.822, 9.551, 0.656], P11C)
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30
]

function build () {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'ARTSTRAX_16x9', width: SLIDE_W, height: SLIDE_H })
  pptx.layout = 'ARTSTRAX_16x9'
  pptx.title = 'ARTSTRAX — Pop Art & Graffiti Presentation'
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY }

  BUILDERS.forEach(fn => fn(pptx))

  const outFile = path.join(__dirname, '51714e88-fadf-4413-a71a-1e4a1a27dc39_grok_final.pptx')
  return pptx.writeFile({ fileName: outFile }).then(() => console.log('wrote', outFile))
}

build().catch(err => {
  console.error(err)
  process.exit(1)
})
