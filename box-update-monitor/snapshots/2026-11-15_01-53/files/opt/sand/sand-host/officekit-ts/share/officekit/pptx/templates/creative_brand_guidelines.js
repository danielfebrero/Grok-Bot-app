#!/usr/bin/env node
/*
 * Babel — Brand Guidelines (42 slides, 26.667in x 15in)
 * Standalone pptxgenjs recreation of the reference deck.
 * Run: node <thisfile>.js   ->  writes <thisfile>.pptx next to itself.
 */
'use strict'

const path = require('path')
const PptxGenJS = require('pptxgenjs')

/* ------------------------------------------------------------------ palette */
const NAVY = '19323C'
const CORAL = 'F2545B'
const CREAM = 'F3F7F0'
const WINE = 'A93F55'
const WHITE = 'FFFFFF'
const MINT = 'F4FFFD'
const SILVER = 'D5D5D5'
const INK = '000000'

/* -------------------------------------------------------------------- fonts */
const MONT_B = 'Montserrat Bold'
const MONT_L = 'Montserrat Light'
const MONT_R = 'Montserrat Regular'
const ROBO = 'Roboto'
const ROBO_L = 'Roboto Light'

/* ------------------------------------------------------------- boilerplate text */
const BRO =
  'Bro ipsum dolor sit amet trail swag crank bro, bomb hole bomb pow pow cruiser face plant line single ' +
  'track spread eagle. White room afterbang stunt stoked ripping switch, carbon manny sucker hole caballerial ' +
  'reverse camber wack core shot gapers lid. Snake bite dust on crust pinner chain suck trucks shuttle laps ' +
  'reverse camber bomb hole pipe. '
const BRO_A = BRO + 'Snake bite dust on crust pinner chain suck trucks shuttle'
const BRO_B = BRO_A + ' laps reverse camber bomb hole pipe.'
const BRO_C = BRO_B + ' Shuttle bus.'
const BRO_D = BRO +
  'Japan air flowy euro sucker hole ollie giblets avie. DH sharkbite smear dust on crust bro presta deck ' +
  'scream ski bum skid air. '
const BRO_E = BRO_D + 'Dust on crust bro presta'
const BRO_F = BRO_D +
  'Epic cornice McTwist back country death cookies bro nose chain ring flow pillow popping. Back country ' +
  'roadie steed piste avie huckfest whip ripper T-bar big ring. Whistler apres schwag, granny gear huck taco ' +
  'mitt beater crunchy gapers hurl carcass bro daffy park rat betty switch. Newschooler lid gear jammer, 360 ' +
  'phat frozen chicken heads death cookies gapers euro 360 trucks Skate north shore grab. Stunt apres DH, ' +
  'bonk rock-ectomy chowder poaching Bike face plant clipless.'
const SRI_SHORT =
  'Sriracha glossier paleo, four loko chillwave polaroid post-ironic leggings kale chips disrupt 8-bit. ' +
  'Kogi migas fingerstache roof'
const SRI = SRI_SHORT + ' party hell of XOXO.'
const PITCH =
  'Pitchfork DIY chia chambray marfa shoreditch 8-bit unicorn, pork belly poke. Post-ironic hella meh, vinyl ' +
  'cornhole succulents cloud bread DIY mlkshk af hashtag humblebrag. Asymmetrical pitchfork fam brooklyn, etsy ' +
  'chambray raw denim cliche banjo PBR&B pop-up pour-over wolf. Pop-up 3 wolf moon scenester, swag ramps ' +
  'semiotics hella pork belly brunch kale chips. Pitchfork DIY chia chambray marfa shoreditch 8-bit unicorn, ' +
  'pork belly poke. Post-ironic hella meh, vinyl cornhole succulents cloud bread '

/* ------------------------------------------------------------------ helpers */
// Every text box in the source deck is centre-anchored with a ~0.078in inset.
// `wrap: false` boxes additionally use "resize shape to fit text", which is how
// the original centres a single line inside its (larger) declared frame.
function txt (s, body, o) {
  s.addText(body, Object.assign({
    fontFace: ROBO_L, fontSize: 18, color: NAVY, align: 'left', valign: 'middle',
    margin: 5.625, isTextBox: true
  }, o))
}

function box (s, x, y, w, h, color, rotate) {
  s.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' }, rotate: rotate || 0 })
}

function roundBox (s, x, y, w, h, color, radius) {
  s.addShape('roundRect', { x, y, w, h, fill: { color }, line: { type: 'none' }, rectRadius: radius })
}

function outline (s, shape, x, y, w, h, color, thickness, radius) {
  s.addShape(shape, {
    x, y, w, h, fill: { type: 'none' }, line: { color, width: thickness || 1 },
    rectRadius: shape === 'roundRect' ? (radius || 0.12) : undefined
  })
}

/* The mark, hairline rule, section number and rotated wordmark that appear on
   every interior slide.  `ink` is the colour of the rail for that background. */
const LOGO_MARK = [
  'M',[18116,16723], [18116,20903], [20903,20903], [20903,16723], 'Z', 'M',[14632,16723],
  [14632,20903], [17419,20903], [17419,16723], 'Z', 'M',[11148,16723], [11148,20903],
  [13935,20903], [13935,16723], 'Z', 'M',[7665,16723], [7665,20903], [10452,20903], [10452,16723],
  'Z', 'M',[4181,16723], [4181,20903], [6968,20903], [6968,16723], 'Z', 'M',[697,16723],
  [697,20903], [3484,20903], [3484,16723], 'Z', 'M',[8603,13587],
  [9226,14025,9982,14284,10800,14284], [11617,14284,12374,14025,12996,13587], 'Z',
  'M',[9058,11989], [8157,12890], [9959,12890], 'Z', 'M',[11497,10944], [10247,12194],
  [10944,12890], [13443,12890], 'Z', 'M',[8710,9406], [8518,9406,8361,9563,8361,9755],
  [8361,9947,8518,10103,8710,10103], [8902,10103,9058,9947,9058,9755],
  [9058,9563,8902,9406,8710,9406], 'Z', 'M',[8710,8710], [9286,8710,9755,9179,9755,9755],
  [9755,10331,9286,10800,8710,10800], [8133,10800,7665,10331,7665,9755],
  [7665,9179,8133,8710,8710,8710], 'Z', 'M',[10800,6619], [8687,6619,6968,8339,6968,10452],
  [6968,11204,7189,11904,7565,12497], [8812,11250], [8948,11114,9168,11114,9304,11250],
  [9755,11701], [11250,10205], [11387,10069,11607,10069,11743,10205], [14035,12497],
  [14411,11904,14632,11204,14632,10452], [14632,8339,12913,6619,10800,6619], 'Z', 'M',[10800,5923],
  [13297,5923,15329,7954,15329,10452], [15329,12949,13297,14981,10800,14981],
  [8303,14981,6271,12949,6271,10452], [6271,7954,8303,5923,10800,5923], 'Z', 'M',[10800,2090],
  [5613,2090,1394,6310,1394,11497], [1394,16026], [20206,16026], [20206,11497],
  [20206,6310,15987,2090,10800,2090], 'Z', 'M',[10800,697], [10328,697,9933,1013,9804,1443],
  [10131,1411,10464,1394,10800,1394], [11136,1394,11469,1411,11796,1443],
  [11667,1013,11272,697,10800,697], 'Z', 'M',[10800,0], [11693,0,12422,677,12522,1543],
  [17275,2363,20903,6513,20903,11497], [20903,16026], [21252,16026],
  [21444,16026,21600,16182,21600,16374], [21600,21252], [21600,21444,21444,21600,21252,21600],
  [348,21600], [156,21600,0,21444,0,21252], [0,16374], [0,16182,156,16026,348,16026], [697,16026],
  [697,11497], [697,6513,4325,2363,9078,1543], [9178,677,9907,0,10800,0], 'Z',
]

function chrome (s, ink, sectionNo, showWordmark) {
  s.addShape('custGeom', {
    x: 1.002, y: 1.937, w: 1.145, h: 1.145, fill: { color: ink }, line: { type: 'none' },
    points: markPoints(1.145)
  })
  box(s, 3.15, 1.937, 0.098, 11.126, ink) // vertical hairline rule
  if (showWordmark === false) return
  txt(s, [
    { text: 'Babel', options: { fontFace: ROBO, bold: true, breakLine: true } },
    { text: 'Creative Design' }
  ], { x: 0.638, y: 11.756, w: 1.874, h: 0.74, color: ink, wrap: false, fit: 'resize', rotate: 270 })
  if (sectionNo) {
    txt(s, sectionNo, {
      x: 1.265, y: 9.422, w: 0.649, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, color: ink,
      lineSpacingMultiple: 1.5, wrap: false, fit: 'resize'
    })
  }
}

// LOGO_MARK is stored in the original 21600x21600 design grid: 'M' starts a
// sub-path, 'Z' closes it, a 2-number array is a line-to and a 6-number array
// is a cubic bezier (two control points then the end point).
function markPoints (size) {
  const k = size / 21600
  const pts = []
  let move = false
  for (const seg of LOGO_MARK) {
    if (seg === 'M') { move = true; continue }
    if (seg === 'Z') { pts.push({ close: true }); continue }
    if (seg.length === 6) {
      pts.push({ x: k * seg[4], y: k * seg[5], curve: { type: 'cubic', x1: k * seg[0], y1: k * seg[1], x2: k * seg[2], y2: k * seg[3] } })
    } else {
      pts.push({ x: k * seg[0], y: k * seg[1], moveTo: move })
    }
    move = false
  }
  return pts
}

/* --------------------------------------------------------- section dividers */
// Full-bleed coral covers: "Section <Word>" plus the list of sub-sections.
const SECTIONS = [
  { no: '0.0', word: 'Content', w: 12.476, ink: CREAM, listY: 12.286, listW: 4.699, listH: 0.74, items: ['0.1 Table of Content'] },
  { no: '1.0', word: 'Brand', w: 10.921, ink: WHITE, listY: 12.286, listW: 2.884, listH: 0.74, items: ['1.1 Overview'] },
  { no: '2.0', word: 'Logo', w: 10.042, ink: WHITE, listY: 10.536, listW: 4.527, listH: 2.49, items: ['2.1 Logo Usage', '2.2 Primary Logo', '2.3 Logo Variations'] },
  { no: '3.0', word: 'Colour', w: 11.275, ink: WHITE, listY: 10.536, listW: 4.232, listH: 2.49, items: ['3.1 Introduction', '3.2 Colour Palette', '3.3 Logo Colour'] },
  { no: '4.0', word: 'Typography', w: 15.5, ink: WHITE, listY: 8.823, listW: 6.106, listH: 4.24, items: ['4.1 Introduction', '4.2 Primary Typeface', '4.3 Secondary Typeface', '4.4 Use of Type', '4.5 Typography Hierarchy'] },
  { no: '5.0', word: 'Print', w: 9.958, ink: WHITE, listY: 10.536, listW: 4.446, listH: 2.49, items: ['5.1 Introduction', '5.2 Business Cards', '5.3 Stationery'] },
  { no: '6.0', word: 'Digital', w: 11.303, ink: WHITE, listY: 10.536, listW: 4.317, listH: 2.49, items: ['6.1 Introduction', '6.2 Web Design', '6.3 Mobile Design'] },
  { no: '7.0', word: 'Imagery', w: 12.694, ink: WHITE, listY: 10.536, listW: 3.984, listH: 2.49, items: ['7.1 Introduction', '7.2 Inpiration', '7.3 Photography'] }
]

function sectionCover (s, sec) {
  s.background = { color: CORAL }
  chrome(s, CREAM, sec.no)
  txt(s, [
    { text: 'Section ', options: { fontFace: MONT_L } },
    { text: sec.word }
  ], {
    x: 4.453, y: 1.937, w: sec.w, h: 2.045, fontSize: 112, fontFace: MONT_B, color: sec.ink,
    charSpacing: -2.24, lineSpacingMultiple: 0.9, wrap: false, fit: 'resize'
  })
  txt(s, sec.items.map((t, i) => ({ text: t, options: { breakLine: i < sec.items.length - 1 } })), {
    x: 4.453, y: sec.listY, w: sec.listW, h: sec.listH, fontSize: 34, fontFace: MONT_L, color: sec.ink,
    lineSpacingMultiple: sec.items.length > 1 ? 1.5 : 1, wrap: false, fit: 'resize'
  })
}

/* ============================================================ slide bodies */

function slide01 (s) {
  s.background = { color: CREAM }
  txt(s, 'Babel', {
    x: 1.647, y: 1.657, w: 4.687, h: 2.045, fontSize: 112, fontFace: MONT_B, color: WINE, charSpacing: -2.24,
    lineSpacingMultiple: 0.9, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Leof. Vasileos Konstantinou', options: { breakLine: true } },
    { text: 'Athina 116 35, ', options: { breakLine: true } },
    { text: 'Greece / Europe' }
  ], {
    x: 12.623, y: 2.457, w: 3.213, h: 1.031, color: WINE, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Brand', options: { breakLine: true } },
    { text: 'Guidelines', options: { fontFace: MONT_L } }
  ], {
    x: 17, y: 9.841, w: 8.019, h: 3.745, fontSize: 112, fontFace: MONT_B, color: WINE, charSpacing: -2.24,
    lineSpacingMultiple: 0.9, wrap: false, fit: 'resize'
  })
  txt(s, 'CONTACT', {
    x: 12.623, y: 1.87, w: 1.475, h: 0.448, fontFace: ROBO, bold: true, color: WINE, charSpacing: 1.8,
    wrap: false, fit: 'resize'
  })
  txt(s, 'PHONE', {
    x: 17.482, y: 1.87, w: 1.123, h: 0.448, fontFace: ROBO, bold: true, color: WINE, charSpacing: 1.8,
    wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Phone +30 21 0752 2984', options: { breakLine: true } },
    { text: 'Free toll +30 123 456', options: { breakLine: true } },
    { text: 'Fax +123 456' }
  ], {
    x: 17.482, y: 2.457, w: 2.915, h: 1.031, color: WINE, wrap: false, fit: 'resize'
  })
  txt(s, 'ONLINE', {
    x: 22.044, y: 1.87, w: 1.195, h: 0.448, fontFace: ROBO, bold: true, color: WINE, charSpacing: 1.8,
    wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'kyloscreative@gmail.com', options: { breakLine: true } },
    { text: 'service@kylos.com', options: { breakLine: true } },
    { text: 'www.kyloscreative.com' }
  ], {
    x: 22.044, y: 2.457, w: 2.976, h: 1.031, color: WINE, wrap: false, fit: 'resize'
  })
}

function slide03 (s) {
  s.background = { color: NAVY }
  chrome(s, CREAM, '0.1')
  txt(s, [
    { text: 'Table', options: { breakLine: true } },
    { text: 'of Content' }
  ], {
    x: 4.513, y: 1.937, w: 2.796, h: 1.323, fontSize: 34, fontFace: MONT_B, color: CREAM, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: '1.0' },
    { text: ' Brand', options: { fontFace: MONT_B, breakLine: true } },
    { text: '2.0' },
    { text: ' Logo', options: { fontFace: MONT_B, breakLine: true } },
    { text: '3.0' },
    { text: ' Colour', options: { fontFace: MONT_B, breakLine: true } },
    { text: '4.0' },
    { text: ' Typography', options: { fontFace: MONT_B, breakLine: true } },
    { text: '5.0' },
    { text: ' Print', options: { fontFace: MONT_B, breakLine: true } },
    { text: '6.0' },
    { text: ' Digital', options: { fontFace: MONT_B, breakLine: true } },
    { text: '7.0' },
    { text: ' Imagery', options: { fontFace: MONT_B } }
  ], {
    x: 4.513, y: 7.036, w: 3.968, h: 5.99, fontSize: 34, fontFace: MONT_L, color: CREAM, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
}

function slide05 (s) {
  chrome(s, NAVY, '1.1')
  txt(s, [
    { text: 'Brand', options: { bold: true, breakLine: true } },
    { text: 'Overview', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.5, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Babel Creative', options: { breakLine: true } },
    { text: 'Since 2015', options: { fontFace: MONT_L } }
  ], {
    x: 4.513, y: 3.875, w: 3.771, h: 1.323, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, PITCH, { x: 13.333, y: 3.875, w: 10.084, h: 2.635, lineSpacingMultiple: 1.5 })
}

function slide07 (s) {
  chrome(s, NAVY, '2.1')
  txt(s, [
    { text: 'Logo', options: { bold: true, breakLine: true } },
    { text: 'Logo Usage', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.909, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, BRO_F, { x: 4.513, y: 6.053, w: 13.26, h: 6.972, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
}

function slide08 (s) {
  chrome(s, NAVY, '2.2')
  txt(s, [
    { text: 'Logo', options: { bold: true, breakLine: true } },
    { text: 'Primary Logo', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.125, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  box(s, 8.401, 4.432, 0.098, 2.929, WINE)
  box(s, 9.915, 3.017, 0.098, 2.929, WINE, 90)
  box(s, 12.874, 5.374, 0.098, 3.491, WINE, 90)
  txt(s, 'Circle of Unity', {
    x: 11.706, y: 4.202, w: 2.632, h: 0.559, fontSize: 24, fontFace: MONT_B, color: WINE, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, 'Unlimited Power', {
    x: 14.947, y: 6.84, w: 3.103, h: 0.559, fontSize: 24, fontFace: MONT_B, color: WINE, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, PITCH, { x: 13.333, y: 10.39, w: 10.084, h: 2.635, lineSpacingMultiple: 1.5 })
  txt(s, SRI_SHORT, { x: 11.706, y: 5.038, w: 7.539, h: 0.885, lineSpacingMultiple: 1.5 })
  txt(s, SRI_SHORT, { x: 14.947, y: 7.677, w: 7.539, h: 0.885, lineSpacingMultiple: 1.5 })
}

function slide09 (s) {
  chrome(s, NAVY, '2.3')
  txt(s, [
    { text: 'Logo', options: { bold: true, breakLine: true } },
    { text: 'Logo Variations', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.447, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Marquee', {
    x: 4.513, y: 4.053, w: 0.943, h: 0.378, fontSize: 14, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 4.709, w: 3.743, h: 1.378, fontSize: 14, lineSpacingMultiple: 1.5 })
  txt(s, 'Logotype', {
    x: 4.513, y: 6.365, w: 0.984, h: 0.378, fontSize: 14, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 7.022, w: 3.743, h: 1.378, fontSize: 14, lineSpacingMultiple: 1.5 })
  txt(s, 'Monogram', {
    x: 4.513, y: 8.678, w: 1.121, h: 0.378, fontSize: 14, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 9.334, w: 3.743, h: 1.378, fontSize: 14, lineSpacingMultiple: 1.5 })
  txt(s, 'Secondary Logo', {
    x: 4.513, y: 10.99, w: 1.584, h: 0.378, fontSize: 14, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 11.647, w: 3.743, h: 1.378, fontSize: 14, lineSpacingMultiple: 1.5 })
}

function slide11 (s) {
  chrome(s, NAVY, '3.1')
  txt(s, [
    { text: 'Colour', options: { bold: true, breakLine: true } },
    { text: 'Introduction', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.936, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Behind The Scene :', options: { breakLine: true } },
    { text: 'Babel Colour' }
  ], {
    x: 4.513, y: 5.181, w: 4.899, h: 1.323, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, BRO_E, { x: 13.333, y: 5.181, w: 10.084, h: 4.639, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
}

function slide12 (s) {
  chrome(s, NAVY, '3.2')
  txt(s, PITCH, { x: 13.333, y: 1.937, w: 10.084, h: 2.635, lineSpacingMultiple: 1.5 })
  box(s, 4.423, 6.286, 5.906, 1.969, WINE)
  box(s, 4.423, 8.672, 5.906, 1.969, CORAL)
  box(s, 4.423, 11.057, 5.906, 1.969, NAVY)
  txt(s, [
    { text: 'Colour', options: { bold: true, breakLine: true } },
    { text: 'Colour Palette', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.226, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  box(s, 13.333, 6.286, 2.521, 1.969, 'B86173')
  box(s, 15.854, 6.286, 2.521, 1.969, 'C88492')
  box(s, 18.375, 6.286, 2.521, 1.969, 'D7A7B1')
  box(s, 20.896, 6.286, 2.521, 1.969, 'E7CAD0')
  box(s, 13.333, 8.672, 2.521, 1.969, 'F47378')
  box(s, 15.854, 8.672, 2.521, 1.969, 'F69296')
  box(s, 18.375, 8.672, 2.521, 1.969, 'F9B1B4')
  box(s, 20.896, 8.672, 2.521, 1.969, 'FBD0D2')
  box(s, 13.333, 11.057, 2.521, 1.969, '42575F')
  box(s, 15.854, 11.057, 2.521, 1.969, '6C7C82')
  box(s, 18.375, 11.057, 2.521, 1.969, '96A1A6')
  box(s, 20.896, 11.057, 2.521, 1.969, 'C0C7C9')
}

function slide13 (s) {
  chrome(s, NAVY, '3.3')
  box(s, 17.11, 4.848, 6.307, 2.458, CORAL)
  box(s, 17.11, 7.726, 6.307, 2.458, NAVY)
  box(s, 17.11, 10.605, 6.307, 2.458, CREAM)
  txt(s, [
    { text: 'Colour', options: { bold: true, breakLine: true } },
    { text: 'Logo Colour', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.945, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Marquee', {
    x: 4.513, y: 4.848, w: 1.163, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 5.573, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
  txt(s, 'Logotype', {
    x: 4.513, y: 7.945, w: 1.217, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 8.671, w: 5.924, h: 1.323, lineSpacingMultiple: 1.5 })
  txt(s, 'Monogram', {
    x: 4.513, y: 10.605, w: 1.393, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 11.331, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
}

function slide15 (s) {
  chrome(s, NAVY, '4.1')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Introduction', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.936, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, BRO_B, { x: 4.513, y: 5.759, w: 8.82, h: 4.639, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, 'Montserrat', {
    x: 17.881, y: 6.714, w: 4.619, h: 1.17, fontSize: 60, fontFace: MONT_R, charSpacing: -1.19,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, 'Primary Typeface', {
    x: 17.881, y: 5.591, w: 2.808, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 17.881, 6.414, 5.536, 0.022, NAVY)
  txt(s, 'Roboto', {
    x: 17.881, y: 9.423, w: 2.741, h: 1.142, fontSize: 60, fontFace: ROBO, charSpacing: -1.19,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, 'Secondary Typeface', {
    x: 17.881, y: 8.3, w: 3.216, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 17.881, 9.123, 5.536, 0.022, NAVY)
  txt(s, 'Typeface We Used', {
    x: 4.513, y: 4.602, w: 4.735, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
}

function slide16 (s) {
  chrome(s, NAVY, '4.2')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Primary Typeface', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.72, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Aa', {
    x: 4.513, y: 4.755, w: 5.097, h: 4.892, fontSize: 280, fontFace: MONT_R, charSpacing: -5.6,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', options: { breakLine: true } },
    { text: 'abcdefghijklmnopqrstuvwxyz', options: { breakLine: true } },
    { text: '1234567890,.?!*&' }
  ], {
    x: 4.513, y: 11.258, w: 7.25, h: 1.767, fontSize: 24, fontFace: MONT_R, charSpacing: 2.4,
    lineSpacingMultiple: 1.5, wrap: false, fit: 'resize'
  })
  txt(s, 'Montserrat', {
    x: 4.513, y: 8.607, w: 8.475, h: 2.045, fontSize: 112, fontFace: MONT_R, charSpacing: -2.24,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, BRO_A, { x: 15.587, y: 1.937, w: 7.83, h: 2.646, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, 'Regular', {
    x: 15.587, y: 7.662, w: 3.458, h: 1.17, fontSize: 60, fontFace: MONT_R, wrap: false, fit: 'resize'
  })
  txt(s, 'Thin', {
    x: 15.587, y: 5.566, w: 1.948, h: 1.17, fontSize: 60, fontFace: 'Montserrat Thin', wrap: false, fit: 'resize'
  })
  txt(s, 'Light', {
    x: 15.587, y: 6.614, w: 2.336, h: 1.17, fontSize: 60, fontFace: MONT_L, wrap: false, fit: 'resize'
  })
  txt(s, 'Medium', {
    x: 15.587, y: 8.711, w: 3.722, h: 1.17, fontSize: 60, fontFace: 'Montserrat Medium', wrap: false, fit: 'resize'
  })
  txt(s, 'Bold', {
    x: 15.587, y: 9.759, w: 2.181, h: 1.17, fontSize: 60, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, 'Extrabold', {
    x: 15.587, y: 10.807, w: 4.48, h: 1.17, fontSize: 60, fontFace: 'Montserrat ExtraBold',
    wrap: false, fit: 'resize'
  })
  txt(s, 'Black', {
    x: 15.587, y: 11.855, w: 2.733, h: 1.17, fontSize: 60, fontFace: 'Montserrat Black', wrap: false, fit: 'resize'
  })
}

function slide17 (s) {
  chrome(s, NAVY, '4.3')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Secondary Typeface', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 3.128, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Aa', {
    x: 4.513, y: 8.351, w: 4.667, h: 4.712, fontSize: 280, fontFace: ROBO, charSpacing: -5.6,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, 'Aa', { x: 12.47, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Bb', { x: 14.438, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Cc', { x: 16.407, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Dd', { x: 18.375, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ee', { x: 20.344, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ff', { x: 22.312, y: 6.623, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Gg', { x: 12.47, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Hh', { x: 14.438, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ii', { x: 16.407, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Jj', { x: 18.375, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Kk', { x: 20.344, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ll', { x: 22.312, y: 7.765, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Mm', { x: 12.47, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Nn', { x: 14.438, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Oo', { x: 16.407, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Pp', { x: 18.375, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Qq', { x: 20.344, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Rr', { x: 22.312, y: 8.907, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ss', { x: 12.47, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Tt', { x: 14.438, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Uu', { x: 16.407, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Vv', { x: 18.375, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Ww', { x: 20.344, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Xx', { x: 22.312, y: 10.05, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, 'Zz', { x: 12.47, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, '01', { x: 14.438, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, '02', { x: 16.407, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, '03', { x: 18.375, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, '04', { x: 20.344, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, '05', { x: 22.312, y: 11.192, w: 1.969, h: 1.142, fontSize: 60, fontFace: ROBO })
  txt(s, BRO_A, { x: 12.47, y: 1.978, w: 10.947, h: 2.208, lineSpacingMultiple: 1.5, margin: 6 })
  box(s, 4.513, 7.14, 5.906, 0.039, NAVY)
  txt(s, [
    { text: 'Secondary Typeface : ', options: { fontFace: ROBO_L } },
    { text: 'Roboto', options: { bold: true } }
  ], {
    x: 4.513, y: 6.414, w: 3.937, h: 0.448, fontFace: ROBO, lineSpacingMultiple: 1.5
  })
}

function slide18 (s) {
  chrome(s, NAVY, '4.3')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Secondary Typeface', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 3.128, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Secondary Typeface : Roboto Font', {
    x: 4.513, y: 6.257, w: 4.017, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 6.983, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
  txt(s, 'Regular', {
    x: 11.745, y: 8.233, w: 5.906, h: 1.976, fontSize: 112, fontFace: ROBO, charSpacing: -2.24,
    lineSpacingMultiple: 0.9
  })
  txt(s, 'Thin', {
    x: 11.745, y: 6.257, w: 5.906, h: 1.976, fontSize: 112, fontFace: 'Roboto Thin', charSpacing: -2.24,
    lineSpacingMultiple: 0.9
  })
  txt(s, 'Light', {
    x: 15.173, y: 6.257, w: 5.906, h: 1.976, fontSize: 112, charSpacing: -2.24, lineSpacingMultiple: 0.9
  })
  txt(s, 'Medium', {
    x: 17.512, y: 8.233, w: 5.906, h: 1.976, fontSize: 112, fontFace: 'Roboto Medium', charSpacing: -2.24,
    lineSpacingMultiple: 0.9
  })
  txt(s, 'Bold', {
    x: 11.745, y: 10.208, w: 5.906, h: 1.976, fontSize: 112, fontFace: ROBO, bold: true, charSpacing: -2.24,
    lineSpacingMultiple: 0.9
  })
  txt(s, 'Black', {
    x: 15.488, y: 10.208, w: 5.906, h: 1.976, fontSize: 112, fontFace: 'Roboto Black', charSpacing: -2.24,
    lineSpacingMultiple: 0.9
  })
}

function slide19 (s) {
  chrome(s, NAVY, '4.4')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Use of Typeface', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.518, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, BRO_B, { x: 4.513, y: 6.536, w: 7.487, h: 3.083, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, 'Description', {
    x: 4.513, y: 5.38, w: 3.041, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, 'We use Montserrat - Bold for Headline', {
    x: 14.599, y: 4.127, w: 8.82, h: 2.184, fontSize: 60, fontFace: MONT_B
  })
  txt(s, [
    { text: 'We use Montserrat - Medium and ', options: { breakLine: true } },
    { text: 'Regular  for body text' }
  ], {
    x: 14.599, y: 7.019, w: 8.573, h: 1.323, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, 'We use Roboto - Light font weight also for captions and small bodies of text', {
    x: 14.599, y: 8.976, w: 5.319, h: 0.885, lineSpacingMultiple: 1.5
  })
  txt(s, 'Also we use Roboto - Light font weight but in smaller size. Because of balance the content', {
    x: 14.599, y: 10.495, w: 7.961, h: 0.378, fontSize: 14, lineSpacingMultiple: 1.5
  })
}

function slide20 (s) {
  chrome(s, NAVY, '4.4')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Use of Typeface', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.518, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'AaBb', {
    x: 4.513, y: 6.132, w: 5.906, h: 2.045, fontSize: 112, fontFace: MONT_B, charSpacing: -2.24,
    lineSpacingMultiple: 0.8
  })
  box(s, 4.513, 5.124, 5.905, 0.039, NAVY)
  txt(s, [
    { text: 'Primary Typeface : ', options: { fontFace: ROBO_L } },
    { text: 'Montserrat', options: { bold: true } }
  ], {
    x: 4.513, y: 4.399, w: 3.937, h: 0.448, fontFace: ROBO, lineSpacingMultiple: 1.5
  })
  txt(s, 'I\'m baby poke vape fashion axe readymade raw denim skateboard adaptogen locavore affogato. Unicorn selfies la croix, shabby chic hella', {
    x: 13.333, y: 6.028, w: 5.006, h: 1.837, fontFace: MONT_L, lineSpacingMultiple: 1.5
  })
  txt(s, 'Bold 112pt', {
    x: 4.513, y: 5.441, w: 3.937, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5
  })
  txt(s, 'Light 18pt', {
    x: 13.333, y: 5.441, w: 3.937, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5
  })
  txt(s, 'AaBb', {
    x: 4.513, y: 11.05, w: 5.906, h: 1.976, fontSize: 112, fontFace: ROBO, bold: true, charSpacing: -2.24,
    lineSpacingMultiple: 0.8
  })
  box(s, 4.513, 10.007, 5.906, 0.039, NAVY)
  txt(s, [
    { text: 'Secondary Typeface : ', options: { fontFace: ROBO_L } },
    { text: 'Roboto', options: { bold: true } }
  ], {
    x: 4.513, y: 9.281, w: 3.937, h: 0.448, fontFace: ROBO, lineSpacingMultiple: 1.5
  })
  txt(s, 'I\'m baby poke vape fashion axe readymade raw denim skateboard adaptogen locavore affogato. Unicorn selfies la croix, shabby chic hella shabby jajan.', {
    x: 13.333, y: 10.949, w: 5.006, h: 1.76, lineSpacingMultiple: 1.5
  })
  txt(s, 'Bold 112pt', {
    x: 4.513, y: 10.324, w: 3.937, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5
  })
  txt(s, 'Light 18pt', {
    x: 13.333, y: 10.324, w: 3.937, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5
  })
  txt(s, BRO_A, { x: 13.333, y: 1.937, w: 10.084, h: 1.389, fontSize: 14, lineSpacingMultiple: 1.5, margin: 6 })
}

function slide21 (s) {
  chrome(s, NAVY, '4.5')
  txt(s, [
    { text: 'Typography', options: { bold: true, breakLine: true } },
    { text: 'Typography Hierarchy', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 3.352, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'AaBbCcDd', {
    x: 12.356, y: 4.055, w: 10.867, h: 2.531, fontSize: 140, fontFace: MONT_B, charSpacing: -2.8,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, 'Title', {
    x: 12.35, y: 3.487, w: 0.846, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 12.35, 4.033, 5.536, 0.022, NAVY)
  txt(s, 'AaBbCcDd', {
    x: 12.35, y: 7.153, w: 8.727, h: 2.045, fontSize: 112, fontFace: MONT_B, charSpacing: -2.24,
    lineSpacingMultiple: 0.9, wrap: false, fit: 'resize'
  })
  txt(s, 'Title Small', {
    x: 12.35, y: 6.586, w: 1.778, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 12.35, 7.131, 5.536, 0.022, NAVY)
  txt(s, 'AaBbCcDd', {
    x: 12.35, y: 9.904, w: 2.843, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, 'Subtitle', {
    x: 12.356, y: 9.337, w: 1.331, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 12.35, 9.882, 5.536, 0.022, NAVY)
  txt(s, 'AaBbCc', {
    x: 12.356, y: 11.489, w: 1.334, h: 0.545, fontSize: 24, lineSpacingMultiple: 1.5, wrap: false, fit: 'resize'
  })
  txt(s, 'Caption', {
    x: 12.35, y: 10.922, w: 1.331, h: 0.545, fontSize: 24, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  box(s, 12.35, 11.467, 5.536, 0.022, NAVY)
  txt(s, BRO_C, { x: 4.513, y: 5.88, w: 5.706, h: 4.396, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, 'Description', {
    x: 4.513, y: 4.724, w: 3.041, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
}

function slide23 (s) {
  chrome(s, NAVY, '5.1')
  txt(s, [
    { text: 'Print', options: { bold: true, breakLine: true } },
    { text: 'Introduction', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.936, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Description', {
    x: 4.513, y: 5.181, w: 3.041, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, BRO_E, { x: 13.333, y: 5.181, w: 10.084, h: 4.639, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
}

function slide24 (s) {
  chrome(s, NAVY, '5.2')
  txt(s, [
    { text: 'Print', options: { bold: true, breakLine: true } },
    { text: 'Business Cards', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.433, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Business Card 01', {
    x: 4.513, y: 7.945, w: 2.135, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 8.671, w: 5.924, h: 1.323, lineSpacingMultiple: 1.5 })
  txt(s, 'Business Card 02', {
    x: 4.513, y: 10.605, w: 2.135, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 11.331, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
  txt(s, BRO_A, { x: 13.036, y: 1.978, w: 10.381, h: 2.208, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, '01', {
    x: 12.44, y: 8.884, w: 0.457, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, '02', {
    x: 12.44, y: 12.577, w: 0.457, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
}

function slide25 (s) {
  chrome(s, NAVY, '5.3')
  txt(s, [
    { text: 'Print', options: { bold: true, breakLine: true } },
    { text: 'Business Cards', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.433, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Letterhead', {
    x: 4.513, y: 4.23, w: 1.386, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 4.955, w: 5.924, h: 1.323, lineSpacingMultiple: 1.5 })
  txt(s, 'Envelope', {
    x: 13.333, y: 4.23, w: 1.194, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 13.333, y: 4.955, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
}

function slide27 (s) {
  chrome(s, NAVY, '6.1')
  txt(s, [
    { text: 'Digital', options: { bold: true, breakLine: true } },
    { text: 'Introduction', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.936, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, BRO_F.slice(0, 505), { x: 4.513, y: 6.053, w: 6.588, h: 6.972, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, BRO_F.slice(505), { x: 11.794, y: 6.053, w: 6.588, h: 6.972, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6 })
  txt(s, 'Description', {
    x: 4.513, y: 4.897, w: 3.041, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
}

function slide28 (s) {
  chrome(s, NAVY, '6.2')
  roundBox(s, 11.606, 5.551, 11.811, 7.475, SILVER, 0.028)
  txt(s, 'Web Design', {
    x: 4.513, y: 5.259, w: 3.189, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Digital', options: { bold: true, breakLine: true } },
    { text: 'Web Design', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.904, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 6.276, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
}

function slide29 (s) {
  chrome(s, NAVY, '6.3')
  txt(s, [
    { text: 'Digital', options: { bold: true, breakLine: true } },
    { text: 'Mobile Design', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.235, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Mockup Device 01', {
    x: 4.513, y: 7.945, w: 2.241, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 8.671, w: 5.924, h: 1.323, lineSpacingMultiple: 1.5 })
  txt(s, 'Mockup Device 02', {
    x: 4.513, y: 10.605, w: 2.241, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 11.331, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
  outline(s, 'roundRect', 13.384, 6.808, 3.225, 6.217, INK, 1, 0.15)
  outline(s, 'ellipse', 14.697, 12.178, 0.599, 0.599, INK)
  outline(s, 'roundRect', 17.032, 3.982, 6.385, 9.109, INK, 1, 0.111)
  outline(s, 'ellipse', 19.855, 12.046, 0.74, 0.74, INK)
}

function slide31 (s) {
  chrome(s, NAVY, '7.1')
  txt(s, [
    { text: 'Q1 critical mass, yet\u00a0prioritize these line items\u00a0ensure to follow requirements when developing solutions\u00a0drill down. Feature creep run it up the flagpole moving the goalposts\u00a0going forward, nor after I ran into Helen at a restaurant, I realized she was just office pretty, and take five, punch the tree, and come back in here with a clear head, and price point.\u00a0Back-end of third quarter\u00a0one-sheet. Shotgun approach\u00a0ping me, organic growth horsehead offer.\u00a0Talk to the slides\u00a0back-end of third quarter.\u00a0In this space. ', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'We just need to put these last issues to bed\u00a0high-level\u00a0pixel pushing, we need to future-proof this.\u00a0I dont care if you got some copy, why you dont use officeipsumcom or something like that ?\u00a0table the discussion price point.\u00a0Define the underlying principles that drive decisions and strategy for your design language\u00a0increase the resolution, scale it up we need a larger print\u00a0talk to the slides\u00a0and\u00a0let\'s not solutionize this right now parking lot it\u00a0so\u00a0keep it lean\u00a0yet\u00a0' }
  ], {
    x: 4.513, y: 6.053, w: 13.869, h: 6.972, fontSize: 24, lineSpacingMultiple: 1.5, margin: 6
  })
  txt(s, [
    { text: 'Imagery', options: { bold: true, breakLine: true } },
    { text: 'Introduction', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.936, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Description', {
    x: 4.513, y: 4.897, w: 3.041, h: 0.74, fontSize: 34, fontFace: MONT_B, wrap: false, fit: 'resize'
  })
}

function slide32 (s) {
  chrome(s, NAVY, '7.2')
  txt(s, [
    { text: 'Imagery', options: { bold: true, breakLine: true } },
    { text: 'Inspiration', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 1.706, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Our Inspiration', {
    x: 4.513, y: 5.551, w: 1.838, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 6.276, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
}

function slide33 (s) {
  chrome(s, NAVY, '7.3')
  txt(s, [
    { text: 'Imagery', options: { bold: true, breakLine: true } },
    { text: 'Photography', options: { fontFace: ROBO_L } }
  ], {
    x: 4.513, y: 1.937, w: 2.031, h: 0.934, fontSize: 24, fontFace: ROBO, wrap: false, fit: 'resize'
  })
  txt(s, 'Minimalist Style', {
    x: 4.513, y: 5.551, w: 1.985, h: 0.448, fontFace: ROBO, bold: true, lineSpacingMultiple: 1.5,
    wrap: false, fit: 'resize'
  })
  txt(s, SRI, { x: 4.513, y: 6.276, w: 5.163, h: 1.76, lineSpacingMultiple: 1.5 })
}

function slide34 (s) {
  s.background = { color: CORAL }
  chrome(s, CREAM, null, false)
  txt(s, 'Thank You', {
    x: 4.453, y: 1.937, w: 10.55, h: 2.531, fontSize: 140, fontFace: MONT_B, color: WHITE, charSpacing: -2.8,
    lineSpacingMultiple: 0.8, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Leof. Vasileos Konstantinou', options: { breakLine: true } },
    { text: 'Athina 116 35, ', options: { breakLine: true } },
    { text: 'Greece / Europe' }
  ], {
    x: 4.453, y: 10.22, w: 3.213, h: 1.031, color: MINT, wrap: false, fit: 'resize'
  })
  txt(s, 'CONTACT', {
    x: 4.453, y: 9.633, w: 1.475, h: 0.448, fontFace: ROBO, bold: true, color: MINT, charSpacing: 1.8,
    wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Phone', options: { bold: true } },
    { text: ' - +022 664 0636', options: { fontFace: ROBO_L } }
  ], {
    x: 4.453, y: 11.681, w: 2.729, h: 0.448, fontFace: ROBO, color: MINT, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Email', options: { bold: true } },
    { text: ' - creative@gmail.com', options: { fontFace: ROBO_L } }
  ], {
    x: 4.453, y: 12.129, w: 3.229, h: 0.448, fontFace: ROBO, color: MINT, wrap: false, fit: 'resize'
  })
  txt(s, [
    { text: 'Website', options: { bold: true } },
    { text: '- www.greaterservice.com', options: { fontFace: ROBO_L } }
  ], {
    x: 4.453, y: 12.577, w: 3.921, h: 0.448, fontFace: ROBO, color: MINT, wrap: false, fit: 'resize'
  })
}


/* -------------------------------------------------------------- icon sheets */
// Slides 35-42 are contact sheets of line-art pictograms.  Each pictogram is
// rebuilt from simple outlined primitives on the same grid as the original.
const ICON_SHEETS = [
  { cols: 5, rows: 5, x0: 5.411, y0: 3.308, dx: 3.93, dy: 2.0625, size: 1.28, last: 5, seed: 0 },
  { cols: 10, rows: 5, x0: 5.073, y0: 3.386, dx: 1.836, dy: 2.057, size: 1.2, last: 10, seed: 4 },
  { cols: 10, rows: 3, x0: 5.19, y0: 4.682, dx: 1.81, dy: 2.8175, size: 1.145, last: 8, seed: 8 },
  { cols: 10, rows: 5, x0: 5.019, y0: 4.6, dx: 1.844, dy: 1.8075, size: 1.296, last: 10, seed: 1 },
  { cols: 10, rows: 5, x0: 4.978, y0: 4.65, dx: 1.862, dy: 1.856, size: 1.377, last: 10, seed: 5 },
  { cols: 10, rows: 5, x0: 4.872, y0: 4.465, dx: 1.876, dy: 2.054, size: 1.115, last: 10, seed: 9 },
  { cols: 10, rows: 4, x0: 5.585, y0: 4.568, dx: 1.718, dy: 2.323, size: 1.546, last: 6, seed: 2 },
  { cols: 10, rows: 5, x0: 4.754, y0: 4.34, dx: 1.906, dy: 1.836, size: 1.368, last: 10, seed: 6 }
]

// A pictogram is drawn as an outlined body plus a few inner marks, laid out on
// an 8x8 unit grid.  Cycling through these recipes reproduces the line-art
// texture of the original contact sheets.
let strokePt = 5 // set per sheet so the weight scales with the icon size

function frame (s, shape, x, y, w, h) { outline(s, shape, x, y, w, h, INK, strokePt) }
function rule (s, x, y, w) { box(s, x, y, w, strokePt / 72, INK) }
function bar (s, x, y, w, h) { box(s, x, y, w, h, INK) }

const ICON_RECIPES = [
  // framed document: header block + text rules
  (s, x, y, u) => {
    frame(s, 'rect', x + 1.4 * u, y + 0.3 * u, 5.2 * u, 7.4 * u)
    bar(s, x + 2.3 * u, y + 1.3 * u, 3.4 * u, 1.1 * u)
    rule(s, x + 2.3 * u, y + 3.4 * u, 3.4 * u)
    rule(s, x + 2.3 * u, y + 4.4 * u, 3.4 * u)
    rule(s, x + 2.3 * u, y + 5.4 * u, 3.4 * u)
    rule(s, x + 2.3 * u, y + 6.4 * u, 2 * u)
  },
  // picture frame: horizon, sun and hills
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.3 * u, y + 1.2 * u, 7.4 * u, 5.6 * u)
    frame(s, 'ellipse', x + 1.3 * u, y + 2.1 * u, 1.5 * u, 1.5 * u)
    frame(s, 'triangle', x + 2.4 * u, y + 3 * u, 3 * u, 2.6 * u)
    frame(s, 'triangle', x + 4.4 * u, y + 3.8 * u, 2.6 * u, 1.8 * u)
    rule(s, x + 0.3 * u, y + 5.6 * u, 7.4 * u)
  },
  // badge: dial inside a rounded square
  (s, x, y, u) => {
    frame(s, 'roundRect', x + 0.6 * u, y + 0.6 * u, 6.8 * u, 6.8 * u)
    frame(s, 'ellipse', x + 2 * u, y + 2 * u, 4 * u, 4 * u)
    bar(s, x + 3.8 * u, y + 2.6 * u, 0.5 * u, 1.6 * u)
    bar(s, x + 4 * u, y + 3.8 * u, 1.4 * u, 0.5 * u)
  },
  // tag / label with a punch hole
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.8 * u, y + 0.8 * u, 6.4 * u, 6.4 * u)
    frame(s, 'ellipse', x + 1.8 * u, y + 1.8 * u, 1.3 * u, 1.3 * u)
    rule(s, x + 1.8 * u, y + 4.2 * u, 4.4 * u)
    rule(s, x + 1.8 * u, y + 5.2 * u, 4.4 * u)
    rule(s, x + 1.8 * u, y + 6.2 * u, 2.6 * u)
  },
  // stacked cards
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.4 * u, y + 0.9 * u, 5.6 * u, 4.2 * u)
    frame(s, 'rect', x + 2 * u, y + 2.9 * u, 5.6 * u, 4.2 * u)
    rule(s, x + 2.7 * u, y + 4.3 * u, 4 * u)
    rule(s, x + 2.7 * u, y + 5.4 * u, 4 * u)
  },
  // clipboard / checklist
  (s, x, y, u) => {
    frame(s, 'roundRect', x + 1.2 * u, y + 0.9 * u, 5.6 * u, 6.7 * u)
    bar(s, x + 2.7 * u, y + 0.2 * u, 2.6 * u, 1.3 * u)
    frame(s, 'rect', x + 2.1 * u, y + 3 * u, 0.9 * u, 0.9 * u)
    rule(s, x + 3.4 * u, y + 3.4 * u, 2.4 * u)
    frame(s, 'rect', x + 2.1 * u, y + 4.6 * u, 0.9 * u, 0.9 * u)
    rule(s, x + 3.4 * u, y + 5 * u, 2.4 * u)
  },
  // monitor on a stand
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.3 * u, y + 0.9 * u, 7.4 * u, 5 * u)
    frame(s, 'rect', x + 1.5 * u, y + 2 * u, 5 * u, 2.8 * u)
    bar(s, x + 3.4 * u, y + 5.9 * u, 1.2 * u, 1.2 * u)
    bar(s, x + 1.8 * u, y + 7 * u, 4.4 * u, 0.6 * u)
  },
  // palette / brush
  (s, x, y, u) => {
    frame(s, 'ellipse', x + 0.5 * u, y + 0.5 * u, 7 * u, 7 * u)
    frame(s, 'ellipse', x + 1.9 * u, y + 1.7 * u, 1.2 * u, 1.2 * u)
    frame(s, 'ellipse', x + 4.9 * u, y + 1.7 * u, 1.2 * u, 1.2 * u)
    frame(s, 'ellipse', x + 1.9 * u, y + 4.6 * u, 1.2 * u, 1.2 * u)
    frame(s, 'ellipse', x + 4.9 * u, y + 4.6 * u, 1.2 * u, 1.2 * u)
  },
  // hand-held device
  (s, x, y, u) => {
    frame(s, 'roundRect', x + 2 * u, y + 0.3 * u, 4 * u, 7.4 * u)
    bar(s, x + 3.1 * u, y + 1.1 * u, 1.8 * u, 0.4 * u)
    frame(s, 'rect', x + 2.7 * u, y + 2.2 * u, 2.6 * u, 3.4 * u)
    frame(s, 'ellipse', x + 3.5 * u, y + 6.2 * u, 1 * u, 1 * u)
  },
  // shopping bag
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.9 * u, y + 2.4 * u, 6.2 * u, 5.3 * u)
    frame(s, 'ellipse', x + 2.3 * u, y + 0.6 * u, 3.4 * u, 3.4 * u)
    rule(s, x + 0.9 * u, y + 4.3 * u, 6.2 * u)
    frame(s, 'rect', x + 2.7 * u, y + 5 * u, 2.6 * u, 1.8 * u)
  },
  // pencil pointing down
  (s, x, y, u) => {
    frame(s, 'rect', x + 2.6 * u, y + 0.4 * u, 2.8 * u, 5 * u)
    frame(s, 'triangle', x + 2.6 * u, y + 5.4 * u, 2.8 * u, 2.2 * u)
    rule(s, x + 2.6 * u, y + 1.6 * u, 2.8 * u)
    rule(s, x + 2.6 * u, y + 4.4 * u, 2.8 * u)
  },
  // ring / pie chart
  (s, x, y, u) => {
    frame(s, 'ellipse', x + 0.4 * u, y + 0.4 * u, 7.2 * u, 7.2 * u)
    frame(s, 'ellipse', x + 2.4 * u, y + 2.4 * u, 3.2 * u, 3.2 * u)
    bar(s, x + 3.8 * u, y + 0.4 * u, 0.4 * u, 2 * u)
    bar(s, x + 3.8 * u, y + 5.6 * u, 0.4 * u, 2 * u)
    bar(s, x + 5.6 * u, y + 3.8 * u, 2 * u, 0.4 * u)
  },
  // envelope
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.4 * u, y + 1.6 * u, 7.2 * u, 4.8 * u)
    frame(s, 'triangle', x + 0.4 * u, y + 1.6 * u, 7.2 * u, 3 * u)
    rule(s, x + 1.4 * u, y + 5.4 * u, 5.2 * u)
  },
  // camera
  (s, x, y, u) => {
    frame(s, 'roundRect', x + 0.4 * u, y + 1.8 * u, 7.2 * u, 5 * u)
    bar(s, x + 2.4 * u, y + 1 * u, 2.4 * u, 0.9 * u)
    frame(s, 'ellipse', x + 2.6 * u, y + 2.8 * u, 2.8 * u, 2.8 * u)
    bar(s, x + 5.8 * u, y + 2.5 * u, 1 * u, 0.6 * u)
  },
  // bar chart
  (s, x, y, u) => {
    frame(s, 'rect', x + 0.5 * u, y + 0.6 * u, 7 * u, 6.8 * u)
    bar(s, x + 1.8 * u, y + 4.4 * u, 1 * u, 2 * u)
    bar(s, x + 3.5 * u, y + 3 * u, 1 * u, 3.4 * u)
    bar(s, x + 5.2 * u, y + 1.9 * u, 1 * u, 4.5 * u)
  },
  // folder
  (s, x, y, u) => {
    bar(s, x + 0.5 * u, y + 1.2 * u, 3 * u, 0.9 * u)
    frame(s, 'rect', x + 0.5 * u, y + 2 * u, 7 * u, 5.2 * u)
    rule(s, x + 1.6 * u, y + 4 * u, 4.8 * u)
    rule(s, x + 1.6 * u, y + 5.4 * u, 4.8 * u)
  }
]

function icon (s, cx, cy, size, variant) {
  const u = size / 8
  ICON_RECIPES[variant % ICON_RECIPES.length](s, cx - size / 2, cy - size / 2, u)
}

function iconSheet (s, sheet) {
  strokePt = sheet.size * 4 // ~0.055in of stroke per inch of icon, as in the source
  for (let r = 0; r < sheet.rows; r++) {
    const count = r === sheet.rows - 1 ? sheet.last : sheet.cols
    for (let c = 0; c < count; c++) {
      icon(s, sheet.x0 + c * sheet.dx, sheet.y0 + r * sheet.dy, sheet.size, sheet.seed + r * 7 + c * 5)
    }
  }
}

/* --------------------------------------------------------------- assemble */
const CONTENT_SLIDES = {
  1: slide01, 3: slide03, 5: slide05, 7: slide07, 8: slide08, 9: slide09, 11: slide11, 12: slide12,
  13: slide13, 15: slide15, 16: slide16, 17: slide17, 18: slide18, 19: slide19, 20: slide20, 21: slide21,
  23: slide23, 24: slide24, 25: slide25, 27: slide27, 28: slide28, 29: slide29, 31: slide31, 32: slide32,
  33: slide33, 34: slide34
}
const SECTION_SLIDES = { 2: 0, 4: 1, 6: 2, 10: 3, 14: 4, 22: 5, 26: 6, 30: 7 }

function build () {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'BABEL', width: 26.667, height: 15 })
  pptx.layout = 'BABEL'
  pptx.author = 'Babel Creative Design'
  pptx.title = 'Brand Guidelines'

  for (let n = 1; n <= 42; n++) {
    const s = pptx.addSlide()
    if (n in SECTION_SLIDES) sectionCover(s, SECTIONS[SECTION_SLIDES[n]])
    else if (n <= 34) CONTENT_SLIDES[n](s)
    else iconSheet(s, ICON_SHEETS[n - 35])
  }
  return pptx
}

build().writeFile({ fileName: path.join(__dirname, '0de72eaa-c595-4ba0-95bb-50ea4b731b6a_grok_final.pptx') })
  .then(f => console.log('wrote ' + f))
  .catch(e => { console.error(e); process.exit(1) })
