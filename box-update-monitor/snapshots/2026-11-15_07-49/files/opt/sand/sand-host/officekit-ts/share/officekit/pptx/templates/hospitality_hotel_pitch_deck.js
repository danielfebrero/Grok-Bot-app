/**
 * Modern Hotel — 30-slide pitch deck, rebuilt with pptxgenjs.
 *
 *   node <this file>            ->  writes the .pptx next to this script
 *
 * Raster artwork in the source deck (phone / tablet / laptop / watch mockups and
 * the app-store badges) is redrawn here with native pptxgenjs shapes; the deck's
 * empty picture placeholders are intentionally left blank, exactly as they render
 * in the original.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design tokens
const C = {
  dark:  '1D1C12',   // theme accent2 — near-black olive
  olive: '8F8C76',   // theme accent1 — brand olive
  slate: '51646F',   // theme accent3
  cream: 'F6F5EF',
  sand:  'EDE4D6',
  taupe: 'C0B7A9',
  white: 'FFFFFF',
  ink:   '0D0D0D',
  gray:  '808080',
  mist:  'F2F2F2',
  black: '000000',
  shell: 'E9E8E4',
  lime:  '92D050',
  teal:  '0C3740',
};

const F = { title: 'Playfair Display', body: 'Poppins' };

const noLine = { type: 'none' };
const noFill = { color: 'FFFFFF', transparency: 100 };   // fully transparent

// Boilerplate copy that recurs across the deck.
const T = {
  lead:  'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau',
  lead2: 'deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul',
  body:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Et dolor diam ultricies sed quisque. Tortor cursus sed blandit..',
  caps:  'Lorep  ipsum duis aute irure dolor in kauselih oilue reprehend',
  caps2: 'esse cill inure dolorlaboru sit amet. Duis aute irusitakus reprei',
  caps3: 'Voluptate lorem kuisais.',
  rate:  'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderiti volui ptates esse cill inure dolorlasue boru sit amet. Duis aute',
};

// ------------------------------------------------------------- device mockups
// The source deck uses photographic device mockups whose screens are cut out, so
// the slide artwork shows straight through them. Each helper below redraws that
// silhouette natively: the frame is a rounded-rectangle *outline* (a thick line,
// so the middle stays see-through) sized from the original artwork.

// A rounded frame drawn as a thick outline, leaving the interior transparent.
function frame(s, x, y, w, h, thick, radius, color, extra) {
  s.addShape('roundRect', {
    x: x + thick / 2, y: y + thick / 2, w: w - thick, h: h - thick,
    rectRadius: Math.max(radius - thick / 2, 0.01),
    fill: noFill, line: { color, width: thick * 72 },
    ...(extra || {}),
  });
}

// The deck ships two phone mockups: a thick-bezel one (slide 24) and a slimmer
// one whose rim reaches the edge of its box (slide 25).
const PHONE_THICK = { inset: 0.055, top: 0.039, rim: 0.046, radius: 0.155 };
const PHONE_SLIM  = { inset: 0.006, top: 0.000, rim: 0.058, radius: 0.135 };

function phoneMock(s, x, y, w, h, style) {
  const p = style || PHONE_THICK;
  const bw = w * (1 - p.inset * 2), bx = x + w * p.inset;
  const bh = h * (1 - p.top * 2), by = y + h * p.top;
  frame(s, bx, by, bw, bh, bw * p.rim, bw * p.radius, '0B0B0C');
  s.addShape('roundRect', {                    // dynamic-island pill
    x: bx + bw * 0.29, y: by + bh * 0.030, w: bw * 0.29, h: bh * 0.040,
    rectRadius: bh * 0.02, fill: { color: '0B0B0C' }, line: noLine,
  });
  const btn = { w: bw * 0.014, fill: { color: '3B3B45' }, line: noLine };
  s.addShape('rect', { x: bx - btn.w * 0.5, y: by + bh * 0.17, h: bh * 0.05, ...btn });
  s.addShape('rect', { x: bx - btn.w * 0.5, y: by + bh * 0.26, h: bh * 0.09, ...btn });
  s.addShape('rect', { x: bx + bw - btn.w * 0.5, y: by + bh * 0.23, h: bh * 0.11, ...btn });
}

function tabletMock(s, x, y, w, h) {
  const bw = w * 0.949, bx = x + w * 0.0425, by = y + h * 0.031, bh = h * 0.936;
  frame(s, bx, by, bw, bh, bw * 0.041, bw * 0.055, '131313');
  [0.44, 0.48, 0.52, 0.56].forEach(f => {      // camera dots in the top bezel
    s.addShape('ellipse', { x: bx + bw * f, y: by + bw * 0.014, w: bw * 0.007, h: bw * 0.007, fill: { color: '4A4A4A' }, line: noLine });
  });
  s.addShape('roundRect', {                    // stylus clipped to the right edge
    x: x + w * 0.960, y: y + h * 0.238, w: w * 0.024, h: h * 0.551,
    rectRadius: w * 0.012, fill: { color: 'EFEFEF' }, line: { color: 'C4C4C4', width: 0.5 },
  });
}

function laptopMock(s, x, y, w, h) {
  // the lid bezel is much thicker top/bottom than at the sides, so it is drawn as
  // four bars around the (transparent) screen rather than as a single outline
  const L = x + w * 0.1065, R = x + w * 1.0,  T = y + h * 0.0485, B = y + h * 0.9040;
  const vb = h * 0.0605,    sb = w * 0.0185;
  const lid = { fill: { color: '0A0A0B' }, line: noLine };
  s.addShape('rect', { x: L, y: T, w: R - L, h: vb, ...lid });
  s.addShape('rect', { x: L, y: B - vb, w: R - L, h: vb, ...lid });
  s.addShape('rect', { x: L, y: T, w: sb, h: B - T, ...lid });
  s.addShape('rect', { x: R - sb, y: T, w: sb, h: B - T, ...lid });
  // aluminium base with its trackpad notch
  s.addShape('rect',      { x: x + w * 0.011, y: y + h * 0.9040, w: w * 0.977, h: h * 0.0430, fill: { color: '9C9DA2' }, line: noLine });
  s.addShape('roundRect', { x: x + w * 0.42,  y: y + h * 0.9040, w: w * 0.16,  h: h * 0.0130, rectRadius: h * 0.006, fill: { color: 'B4B5B9' }, line: noLine });
  s.addShape('rect',      { x: x + w * 0.011, y: y + h * 0.9470, w: w * 0.977, h: h * 0.0170, fill: { color: '46474C' }, line: noLine });
}

function watchMock(s, x, y, w, h, rotate) {
  const rot = { rotate }, solid = { rotate, line: noLine };
  // the woven band is two straps; the case sits between them
  s.addShape('roundRect', { x: x + w * 0.22, y, w: w * 0.56, h: h * 0.24, rectRadius: w * 0.07, fill: { color: 'D8CFC5' }, ...solid });
  s.addShape('roundRect', { x: x + w * 0.20, y: y + h * 0.76, w: w * 0.60, h: h * 0.24, rectRadius: w * 0.07, fill: { color: 'D8CFC5' }, ...solid });
  frame(s, x + w * 0.060, y + h * 0.195, w * 0.917, h * 0.610, w * 0.098, w * 0.27, 'CDC4BC', rot);   // titanium case
  frame(s, x + w * 0.118, y + h * 0.212, w * 0.764, h * 0.576, w * 0.062, w * 0.21, '060606', rot);   // black bezel
  s.addShape('roundRect', { x: x + w * 0.955, y: y + h * 0.395, w: w * 0.048, h: h * 0.085, rectRadius: w * 0.02, fill: { color: 'A9A096' }, ...solid });
  s.addShape('rect', { x: x + w * 0.045, y: y + h * 0.470, w: w * 0.022, h: h * 0.065, fill: { color: 'D2691E' }, ...solid });
}

function storeBadge(s, x, y, w, h, small, big) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: h * 0.13, fill: { color: C.black }, line: { color: '8A8A8A', width: 0.5 } });
  s.addText([
    { text: small, options: { fontSize: 5, breakLine: true } },
    { text: big, options: { fontSize: 11 } },
  ], { x: x + w * 0.2, y, w: w * 0.78, h, fontFace: F.body, color: C.white, valign: 'middle', margin: 0, lineSpacingMultiple: 1 });
}

// ------------------------------------------------------------ slide builders
const SLIDES = [
  function slide01(s) {
    s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 2.569, fill: { color: C.dark }, line: noLine });
    s.addText('Modern Hotel', { x: 3.198, y: 0.403, w: 8.665, h: 1.447, fontSize: 80, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('PRESENTATION', {
      x: 0.425,
      y: 6.788,
      w: 1.303,
      h: 0.252,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      align: 'center',
      valign: 'top'
    });
    s.addText('.01', {
      x: 11.033,
      y: 1.199,
      w: 1.912,
      h: 1.313,
      fontSize: 72,
      fontFace: F.title,
      color: C.white,
      align: 'right',
      valign: 'top'
    });
    s.addText([
      { text: T.lead, options: { breakLine: true } },
      { text: T.lead2 },
    ], { x: 3.198, y: 6.136, w: 4.666, h: 0.761, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Modern Design', { x: 3.198, y: 5.85, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('DESIGN', {
      x: 11.073,
      y: 6.803,
      w: 1.303,
      h: 0.252,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      align: 'center',
      valign: 'top'
    });
    s.addText('PRESENTATION', {
      x: 0.425,
      y: 0.323,
      w: 1.303,
      h: 0.252,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      align: 'center',
      valign: 'top'
    });
  },
  function slide02(s) {
    s.addShape('rect', { x: 0.328, y: 0.334, w: 5.276, h: 4.794, fill: { color: C.slate }, line: noLine });
    s.addShape('rect', { x: 0.328, y: 5.128, w: 5.276, h: 2.057, fill: { color: C.cream }, line: noLine });
    s.addText('01', { x: 1.115, y: 5.948, w: 0.557, h: 0.337, fontSize: 14, fontFace: F.body, color: C.ink, valign: 'top' });
    s.addText('Great Copy Text', { x: 1.952, y: 5.971, w: 2.245, h: 0.286, fontSize: 11, fontFace: F.body, color: C.ink, valign: 'top' });
    s.addText(T.lead, {
      x: 1.115,
      y: 1.291,
      w: 2.901,
      h: 0.533,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('Modern Design', { x: 1.115, y: 3.945, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('INSERT YOUR TAGLINE HERE', {
      x: 10.715,
      y: 5.948,
      w: 2.291,
      h: 0.648,
      fontFace: F.body,
      color: C.white,
      charSpacing: 3,
      align: 'right',
      lineSpacingMultiple: 0.9,
      valign: 'top'
    });
    s.addText([
      { text: 'Cozy Room', options: { breakLine: true } },
      { text: 'For Two' },
    ], { x: 6.106, y: 0.68, w: 5.533, h: 2.524, fontSize: 72, fontFace: F.title, color: C.white, valign: 'top' });
  },
  function slide03(s) {
    s.addText('Deluxe Room', {
      x: 1.387,
      y: 2.179,
      w: 10.559,
      h: 2.036,
      fontSize: 115,
      fontFace: F.title,
      color: C.white,
      bold: true,
      valign: 'top'
    });
    s.addText(T.lead, {
      x: 0.353,
      y: 5.412,
      w: 2.901,
      h: 0.533,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('MODERN DESIGN', {
      x: 4.344,
      y: 0.554,
      w: 4.646,
      h: 0.286,
      fontSize: 11,
      fontFace: F.title,
      color: C.ink,
      charSpacing: 6,
      align: 'center',
      valign: 'top'
    });
    s.addText('COVER', {
      x: 11.076,
      y: 5.536,
      w: 1.904,
      h: 0.286,
      fontSize: 11,
      fontFace: F.title,
      color: C.ink,
      align: 'right',
      valign: 'top'
    });
  },
  function slide04(s) {
    s.addShape('rect', { x: 5.437, y: 0, w: 7.896, h: 3.946, fill: { color: C.olive }, line: noLine });
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul', {
      x: 7.786,
      y: 4.843,
      w: 4.42,
      h: 0.76,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText([
      { text: 'Mini Pool', options: { breakLine: true } },
      { text: 'Compartment' },
    ], { x: 5.958, y: 0.711, w: 6.844, h: 2.524, fontSize: 72, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('Modern Design', { x: 7.786, y: 4.557, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addShape('ellipse', { x: 7.874, y: 6.137, w: 0.516, h: 0.516, fill: { color: C.olive }, line: noLine });
    s.addText('+33', {
      x: 7.881,
      y: 6.307,
      w: 0.5,
      h: 0.172,
      fontSize: 10,
      fontFace: F.body,
      color: C.white,
      valign: 'middle',
      margin: [7.2, 7.2, 7.2, 7.2]
    });
    s.addText('YOUR TEXT HERE', { x: 8.514, y: 6.135, w: 1.557, h: 0.286, fontSize: 11, fontFace: F.body, color: C.ink, valign: 'top' });
    s.addText('Lorem ipsum dolor ', {
      x: 8.514,
      y: 6.346,
      w: 1.821,
      h: 0.283,
      fontSize: 8,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('ellipse', { x: 10.118, y: 6.137, w: 0.516, h: 0.516, fill: { color: C.dark }, line: noLine });
    s.addText('+23', {
      x: 10.125,
      y: 6.307,
      w: 0.5,
      h: 0.172,
      fontSize: 10,
      fontFace: F.body,
      color: C.white,
      valign: 'middle',
      margin: [7.2, 7.2, 7.2, 7.2]
    });
    s.addText('YOUR TEXT HERE', { x: 10.758, y: 6.135, w: 1.557, h: 0.286, fontSize: 11, fontFace: F.body, color: C.ink, valign: 'top' });
    s.addText('Lorem ipsum dolor ', {
      x: 10.758,
      y: 6.346,
      w: 1.821,
      h: 0.283,
      fontSize: 8,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
  },
  function slide05(s) {
    s.addShape('rect', { x: 0, y: 0, w: 7.485, h: 7.5, fill: { color: C.white }, line: noLine });
    s.addText('COVER', { x: 0.353, y: 0.58, w: 1.904, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Modern Design', {
      x: 11.076,
      y: 0.554,
      w: 1.904,
      h: 0.286,
      fontSize: 11,
      fontFace: F.title,
      color: C.white,
      align: 'right',
      valign: 'top'
    });
    s.addText(T.lead, {
      x: 1.564,
      y: 0.431,
      w: 2.901,
      h: 0.533,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('rect', { x: 0, y: 3.828, w: 5.446, h: 2.541, fill: { color: C.white }, line: noLine });
    s.addShape('rect', { x: 5.604, y: 4.377, w: 7.559, h: 2.956, fill: { color: C.white }, line: noLine });
    s.addShape('rect', { x: -0.013, y: 6.369, w: 1.124, h: 1.131, fill: { color: C.slate }, line: noLine });
    s.addShape('rect', { x: -0.013, y: 0, w: 1.124, h: 3.829, fill: { color: C.olive }, line: noLine });
    s.addText([
      { text: 'Lorep  ipsum duis ' },
      { text: 'aute irure ', options: { bold: true } },
      { text: 'dolor in kauselih oilue reprehend', options: { breakLine: true } },
      { text: 'esse cill inure dolorlaboru sit amet. Duis aute' },
      { text: ' irusitakus ', options: { bold: true } },
      { text: 'reprei', options: { breakLine: true } },
      { text: T.caps3 },
    ], { x: 0.353, y: 5.168, w: 4.332, h: 0.688, fontSize: 8, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Presentation Design', {
      x: 0.353,
      y: 4.135,
      w: 2.947,
      h: 0.337,
      fontSize: 14,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addText([
      { text: 'Lorep  ipsum duis ' },
      { text: 'aute irure ', options: { bold: true } },
      { text: 'dolor in kauselih oilue reprehend', options: { breakLine: true } },
      { text: 'esse cill inure dolorlaboru sit amet. Duis aute' },
      { text: ' irusitakus ', options: { bold: true } },
      { text: 'reprei', options: { breakLine: true } },
      { text: T.caps3 },
    ], { x: 6.691, y: 6.164, w: 4.372, h: 0.687, fontSize: 8, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Presentation Design', {
      x: 6.691,
      y: 5.131,
      w: 2.947,
      h: 0.337,
      fontSize: 14,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addShape('rect', { x: 3.425, y: 4.116, w: 1.26, h: 0.413, fill: { color: C.olive }, line: noLine });
    s.addText('Button Here', {
      x: 3.425,
      y: 4.116,
      w: 1.26,
      h: 0.413,
      fontSize: 11,
      fontFace: F.body,
      color: C.white,
      underline: { style: 'sng' },
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 11.116, y: 5.14, w: 1.26, h: 0.413, fill: { color: C.olive }, line: noLine });
    s.addText('Button Here', {
      x: 11.116,
      y: 5.14,
      w: 1.26,
      h: 0.413,
      fontSize: 11,
      fontFace: F.body,
      color: C.white,
      underline: { style: 'sng' },
      align: 'center',
      valign: 'middle'
    });
    s.addText('Elevate your experience in Travelling', {
      x: 5.614,
      y: 1.772,
      w: 7.391,
      h: 1.582,
      fontSize: 44,
      fontFace: F.title,
      color: C.white,
      valign: 'top'
    });
  },
  function slide06(s) {
    s.addText('Modern', { x: 1.276, y: 0.647, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Design', { x: 1.276, y: 0.968, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Pitch', { x: 1.276, y: 1.289, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Modern Hotel in Sanur', {
      x: 7.647,
      y: 3.442,
      w: 4.979,
      h: 3.13,
      fontSize: 60,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addText('01', { x: 11.367, y: 0.455, w: 0.734, h: 0.656, fontSize: 33, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide07(s) {
    s.addText('Room Management', {
      x: 1.115,
      y: 1.618,
      w: 11.104,
      h: 1.111,
      fontSize: 60,
      fontFace: F.title,
      color: C.ink,
      align: 'center',
      valign: 'top'
    });
    s.addText('MODERN DESIGN', {
      x: 4.344,
      y: 1.086,
      w: 4.646,
      h: 0.286,
      fontSize: 11,
      fontFace: F.title,
      color: C.ink,
      charSpacing: 6,
      align: 'center',
      valign: 'top'
    });
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul', {
      x: 1.556,
      y: 2.975,
      w: 10.222,
      h: 0.6,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.gray,
      align: 'center',
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
  },
  function slide08(s) {
    s.addShape('rect', { x: 4.462, y: 0.321, w: 8.543, h: 2.047, fill: { color: C.cream }, line: noLine });
    s.addText('Modern', { x: 11.508, y: 0.881, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Design', { x: 11.508, y: 1.202, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Pitch', { x: 11.508, y: 1.523, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Comforting Bed', { x: 5.604, y: 0.958, w: 5.63, h: 0.774, fontSize: 40, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide09(s) {
    s.addShape('custGeom', { x: 0, y: 0, w: 1.293, h: 2.225, fill: { color: C.white }, line: noLine, points: [
      { x: 0, y: 0, moveTo: true },
      { x: 1.221, y: 0 },
      { x: 1.293, y: 0 },
      { x: 1.293, y: 2.225 },
      { close: true },
    ] });
    s.addText([
      { text: 'Cozy ', options: { breakLine: true } },
      { text: 'Rooms' },
    ], { x: 1.405, y: 1.049, w: 4.463, h: 1.919, fontSize: 54, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addShape('custGeom', { x: 0, y: 3.771, w: 11.042, h: 1.826, fill: { color: C.mist }, line: noLine, points: [
      { x: 0, y: 1.826, moveTo: true },
      { x: 6.667, y: 1.826 },
      { x: 7.306, y: 1.826 },
      { x: 9.439, y: 1.826 },
      { x: 11.042, y: 0 },
      { x: 7.306, y: 0 },
      { x: 6.667, y: 0 },
      { x: 0, y: 0 },
      { close: true },
    ] });
    s.addText('Text Tittle Here', { x: 1.391, y: 4.082, w: 1.593, h: 0.286, fontSize: 14, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. ', {
      x: 1.391,
      y: 4.368,
      w: 3.262,
      h: 0.865,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('roundRect', { x: 1.407, y: 5.839, w: 1.672, h: 0.462, fill: { color: C.olive }, line: noLine, rectRadius: 0.231 });
    s.addText('Subtitle Here', {
      x: 1.407,
      y: 5.839,
      w: 1.672,
      h: 0.462,
      fontSize: 14,
      fontFace: F.title,
      color: C.white,
      align: 'center',
      valign: 'middle'
    });
    s.addText('Text Tittle Here', { x: 5.722, y: 4.082, w: 1.593, h: 0.286, fontSize: 14, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. ', {
      x: 5.722,
      y: 4.368,
      w: 3.262,
      h: 0.865,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('roundRect', { x: 5.738, y: 5.839, w: 1.672, h: 0.462, fill: { color: C.olive }, line: noLine, rectRadius: 0.231 });
    s.addText('Subtitle Here', {
      x: 5.738,
      y: 5.839,
      w: 1.672,
      h: 0.462,
      fontSize: 14,
      fontFace: F.title,
      color: C.white,
      align: 'center',
      valign: 'middle'
    });
  },
  function slide10(s) {
    s.addText(T.rate, {
      x: 7.545,
      y: 3.733,
      w: 4.674,
      h: 0.529,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('Satisfied Rate Room', {
      x: 7.545,
      y: 1.581,
      w: 4.249,
      h: 1.582,
      fontSize: 44,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addShape('roundRect', {
      x: 7.691,
      y: 4.931,
      w: 3.975,
      h: 0.261,
      fill: { color: C.teal, transparency: 92 },
      line: noLine,
      rectRadius: 0.131
    });
    s.addText('60%', {
      x: 7.691,
      y: 4.931,
      w: 3.975,
      h: 0.261,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.black,
      align: 'right',
      valign: 'middle'
    });
    s.addShape('roundRect', { x: 7.691, y: 4.931, w: 2.331, h: 0.261, fill: { color: C.dark }, line: noLine, rectRadius: 0.131 });
    s.addShape('roundRect', {
      x: 7.699,
      y: 5.78,
      w: 3.975,
      h: 0.261,
      fill: { color: C.teal, transparency: 92 },
      line: noLine,
      rectRadius: 0.131
    });
    s.addText('80%', {
      x: 7.699,
      y: 5.78,
      w: 3.975,
      h: 0.261,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.black,
      align: 'right',
      valign: 'middle'
    });
    s.addShape('roundRect', { x: 7.699, y: 5.78, w: 3.022, h: 0.261, fill: { color: C.olive }, line: noLine, rectRadius: 0.131 });
    s.addText('Leverage agile frameworks to provide', {
      x: 7.635,
      y: 5.244,
      w: 4.244,
      h: 0.342,
      fontSize: 12,
      fontFace: F.body,
      lineSpacingMultiple: 1.3,
      valign: 'top'
    });
    s.addText('Leverage agile frameworks to provide', {
      x: 7.635,
      y: 6.104,
      w: 4.244,
      h: 0.342,
      fontSize: 12,
      fontFace: F.body,
      lineSpacingMultiple: 1.3,
      valign: 'top'
    });
  },
  function slide11(s) {
    s.addShape('rect', { x: 9.384, y: 0, w: 3.621, h: 2.5, fill: { color: C.cream }, line: noLine });
    s.addText('22.000', { x: 9.661, y: 0.448, w: 3.125, h: 0.672, fontSize: 16, fontFace: F.title, valign: 'top' });
    s.addText(T.body, { x: 9.661, y: 1.204, w: 3.125, h: 0.881, fontSize: 9, fontFace: F.body, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addShape('rect', { x: 9.384, y: 2.5, w: 3.621, h: 2.5, fill: { color: C.dark }, line: noLine });
    s.addText('44.000', { x: 9.661, y: 2.948, w: 3.125, h: 0.672, fontSize: 16, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText(T.body, {
      x: 9.661,
      y: 3.704,
      w: 3.125,
      h: 0.881,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('rect', { x: 9.384, y: 5, w: 3.621, h: 2.5, fill: { color: C.cream }, line: noLine });
    s.addText('22.000', { x: 9.661, y: 5.448, w: 3.125, h: 0.672, fontSize: 16, fontFace: F.title, valign: 'top' });
    s.addText(T.body, { x: 9.661, y: 6.204, w: 3.125, h: 0.881, fontSize: 9, fontFace: F.body, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Modern', { x: 5.765, y: 0.647, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Design', { x: 5.765, y: 0.968, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Pitch', { x: 5.765, y: 1.289, w: 0.941, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Customer in 2024', { x: 5.604, y: 4.414, w: 3.267, h: 1.717, fontSize: 48, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide12(s) {
    s.addShape('rect', { x: 0.328, y: 2.569, w: 7.363, h: 3.622, fill: { color: C.olive }, line: noLine });
    s.addText('Our Room Rates', {
      x: 2.388,
      y: 0.397,
      w: 8.558,
      h: 0.656,
      fontSize: 33,
      fontFace: F.title,
      color: C.ink,
      align: 'center',
      valign: 'top'
    });
    s.addText([
      { text: T.lead, options: { breakLine: true } },
      { text: T.lead2 },
    ], { x: 1.62, y: 1.28, w: 10.093, h: 0.532, fontSize: 9, fontFace: F.body, color: C.gray, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Your Text Here', { x: 0.939, y: 4.286, w: 2.211, h: 0.286, fontSize: 11, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText(T.body, {
      x: 0.939,
      y: 4.611,
      w: 2.106,
      h: 1.215,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('Image', {
      x: 5.985,
      y: 2.811,
      w: 1.364,
      h: 0.37,
      fontSize: 16,
      fontFace: F.title,
      color: C.white,
      bold: true,
      valign: 'top'
    });
    s.addText([
      { text: 'Lorep  ipsum duis ' },
      { text: 'aute irure ', options: { bold: true } },
      { text: 'dolor in kauselih oilue reprehend', options: { breakLine: true } },
      { text: 'esse cill inure dolorlaboru sit amet. Duis aute' },
      { text: ' irusitakus ', options: { bold: true } },
      { text: 'reprei', options: { breakLine: true } },
      { text: T.caps3 },
    ], { x: 7.925, y: 5.343, w: 4.217, h: 0.686, fontSize: 8, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addShape('rect', { x: 10.958, y: 6.273, w: 1.26, h: 0.413, fill: { color: C.olive }, line: noLine });
    s.addText('Button Here', {
      x: 10.958,
      y: 6.273,
      w: 1.26,
      h: 0.413,
      fontSize: 11,
      fontFace: F.body,
      color: C.white,
      underline: { style: 'sng' },
      align: 'center',
      valign: 'middle'
    });
    s.addText('200M', {
      x: 0.938,
      y: 2.811,
      w: 1.364,
      h: 0.37,
      fontSize: 16,
      fontFace: F.title,
      color: C.white,
      bold: true,
      valign: 'top'
    });
  },
  function slide13(s) {
    s.background = { color: C.olive };
    s.addText('Suite Room', { x: 5.431, y: 0.962, w: 10.441, h: 1.111, fontSize: 60, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText([
      { text: 'Lorep  ipsum duis ' },
      { text: 'aute irure ', options: { bold: true } },
      { text: 'dolor in kauselih oilue reprehend esse cill inure dolorlaboru sit amet. Duis aute' },
      { text: ' irusitakus ', options: { bold: true } },
      { text: 'reprei', options: { breakLine: true } },
      { text: T.caps3 },
    ], { x: 5.431, y: 2.189, w: 7.915, h: 0.485, fontSize: 8, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Gallery Photo', { x: 0.678, y: 6.512, w: 2.947, h: 0.337, fontSize: 14, fontFace: F.title, color: C.white, valign: 'top' });
  },
  function slide14(s) {
    s.addText('One Stop Place', { x: 1.128, y: 4.264, w: 10.441, h: 1.111, fontSize: 60, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText([
      { text: 'Lorep  ipsum duis ' },
      { text: 'aute irure ', options: { bold: true } },
      { text: 'dolor in kauselih oilue reprehend esse cill inure dolorlaboru sit amet. Duis aute' },
      { text: ' irusitakus ', options: { bold: true } },
      { text: 'reprei', options: { breakLine: true } },
      { text: T.caps3 },
    ], { x: 4.475, y: 5.976, w: 7.915, h: 0.485, fontSize: 8, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Gallery Photo', { x: 1.128, y: 5.976, w: 2.947, h: 0.337, fontSize: 14, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide15(s) {
    s.addShape('roundRect', { x: 6.781, y: 3.814, w: 6.302, h: 3.436, fill: { color: C.dark }, line: noLine, rectRadius: 0.164 });
    s.addShape('roundRect', { x: 0.25, y: 0.25, w: 6.351, h: 3.436, fill: { color: C.olive }, line: noLine, rectRadius: 0.164 });
    s.addText('Basic Room', { x: 0.509, y: 0.543, w: 5.225, h: 0.774, fontSize: 40, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('Deluxe Room', { x: 6.986, y: 4.117, w: 3.663, h: 0.774, fontSize: 40, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText([
      { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing ', options: { bullet: { characterCode: '2022', indent: 13.536 }, breakLine: true } },
      { text: 'elit. Et dolor  diam ultricies sed quisque.', options: { bullet: { characterCode: '2022', indent: 13.536 }, breakLine: true } },
      { text: 'Tortor cursus sed blandit..', options: { bullet: { characterCode: '2022', indent: 13.536 } } },
    ], { x: 6.986, y: 6.216, w: 3.835, h: 0.894, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText([
      { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing ', options: { bullet: { characterCode: '2022', indent: 13.536 }, breakLine: true } },
      { text: 'elit. Et dolor  diam ultricies sed quisque.', options: { bullet: { characterCode: '2022', indent: 13.536 }, breakLine: true } },
      { text: 'Tortor cursus sed blandit..', options: { bullet: { characterCode: '2022', indent: 13.536 } } },
    ], { x: 0.509, y: 2.371, w: 3.835, h: 0.894, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('$25', {
      x: 5.204,
      y: 0.712,
      w: 1.189,
      h: 0.438,
      fontSize: 20,
      fontFace: F.title,
      color: C.white,
      align: 'right',
      valign: 'top'
    });
    s.addText('$25', {
      x: 11.557,
      y: 4.281,
      w: 1.189,
      h: 0.438,
      fontSize: 20,
      fontFace: F.title,
      color: C.white,
      align: 'right',
      valign: 'top'
    });
  },
  function slide16(s) {
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec tempor quam et mollis consequat. Vestibulum efficitur urna sed ante sagittis.', {
      x: 4.423,
      y: 6.336,
      w: 6.667,
      h: 0.603,
      fontSize: 10.5,
      fontFace: F.body,
      color: C.black,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('Nulla vulputate diam gravida tempus', {
      x: 2.244,
      y: 6.398,
      w: 2.053,
      h: 0.505,
      fontSize: 12,
      fontFace: F.body,
      bold: true,
      valign: 'top'
    });
    s.addText('Take a look around', {
      x: 1.536,
      y: 0.957,
      w: 10.441,
      h: 1.111,
      fontSize: 60,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addText('01', { x: 1.789, y: 5.098, w: 1.203, h: 1.111, fontSize: 60, fontFace: F.title, color: C.taupe, valign: 'top' });
    s.addText('02', { x: 5.296, y: 5.098, w: 1.203, h: 1.111, fontSize: 60, fontFace: F.title, color: C.taupe, valign: 'top' });
    s.addText('03', { x: 8.866, y: 5.098, w: 1.203, h: 1.111, fontSize: 60, fontFace: F.title, color: C.taupe, valign: 'top' });
  },
  function slide17(s) {
    s.addShape('rect', { x: 0.17, y: 2.569, w: 12.993, h: 4.764, fill: { color: C.mist }, line: noLine });
    s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 2.569, fill: { color: C.dark }, line: noLine });
    s.addShape('roundRect', { x: 0.328, y: 6.389, w: 12.677, h: 0.749, fill: { color: C.white }, line: noLine });
    s.addShape('line', { x: 0.329, y: 6.407, w: 12.675, h: 0, line: { color: C.olive, width: 3 } });
    s.addText('09:00am', {
      x: 0.633,
      y: 6.561,
      w: 1.203,
      h: 0.41,
      fontSize: 14,
      fontFace: F.body,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
    s.addText('09:30am', {
      x: 3.556,
      y: 6.561,
      w: 1.203,
      h: 0.41,
      fontSize: 14,
      fontFace: F.body,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
    s.addText('09:40am', {
      x: 6.972,
      y: 6.561,
      w: 1.203,
      h: 0.41,
      fontSize: 14,
      fontFace: F.body,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
    s.addText([
      { text: 'Lorem Ipsum  duist', options: { breakLine: true } },
      { text: 'Aure inure dolor' },
    ], { x: 1.579, y: 5.13, w: 1.37, h: 0.529, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Text Tittle Here', { x: 1.579, y: 4.885, w: 1.454, h: 0.252, fontSize: 9, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText([
      { text: 'Lorem Ipsum  duist', options: { breakLine: true } },
      { text: 'Aure inure dolor' },
    ], { x: 5.094, y: 5.13, w: 1.37, h: 0.529, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Text Tittle Here', { x: 5.094, y: 4.885, w: 1.454, h: 0.252, fontSize: 9, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText([
      { text: 'Lorem Ipsum  duist', options: { breakLine: true } },
      { text: 'Aure inure dolor' },
    ], { x: 11.357, y: 5.13, w: 1.37, h: 0.529, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('Text Tittle Here', {
      x: 11.357,
      y: 4.885,
      w: 1.454,
      h: 0.252,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      valign: 'top'
    });
  },
  function slide18(s) {
    s.addText([
      { text: 'Partnership', options: { breakLine: true } },
      { text: 'Program ' },
    ], { x: 0.958, y: 1.001, w: 7.829, h: 3.063, fontSize: 88, fontFace: F.title, valign: 'top' });
    s.addText('Section Here', {
      x: 1.162,
      y: 5.833,
      w: 1.651,
      h: 0.337,
      fontSize: 14,
      fontFace: F.title,
      color: C.ink,
      underline: { style: 'sng' },
      valign: 'top'
    });
    s.addShape('custGeom', { x: 5.912, y: 6.671, w: 0.354, h: 0.354, fill: { color: C.white }, line: noLine, points: [
      { x: 0.354, y: 0, moveTo: true },
      { x: 0.012, y: 0 },
      { x: 0.012, y: 0.058 },
      { x: 0.255, y: 0.058 },
      { x: 0, y: 0.312 },
      { x: 0.042, y: 0.354 },
      { x: 0.296, y: 0.099 },
      { x: 0.296, y: 0.342 },
      { x: 0.354, y: 0.342 },
      { close: true },
    ] });
  },
  function slide19(s) {
    s.addShape('rect', { x: 6.81, y: 0, w: 6.524, h: 7.5, fill: { color: C.mist }, line: noLine });
    s.addShape('ellipse', { x: 8.214, y: 2.327, w: 4.548, h: 4.548, fill: { color: C.dark }, line: noLine });
    s.addText('$1 Million or More Annually', {
      x: 10.249,
      y: 4.219,
      w: 2.251,
      h: 0.774,
      fontSize: 20,
      fontFace: F.title,
      color: C.white,
      align: 'center',
      valign: 'top'
    });
    s.addShape('ellipse', { x: 6.416, y: 2.05, w: 3.778, h: 3.778, fill: { color: C.olive }, line: noLine });
    s.addText([
      { text: '80% ', options: { color: C.shell, breakLine: true } },
      { text: 'Business Trip', options: { color: C.white } },
    ], { x: 7.033, y: 3.197, w: 2.545, h: 1.515, fontSize: 28, fontFace: F.title, align: 'center', valign: 'top' });
    s.addText([
      { text: '80%of Customer ', options: { breakLine: true } },
      { text: 'in Our Hotel' },
    ], { x: 0.575, y: 0.333, w: 5.405, h: 3.433, fontSize: 66, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide20(s) {
    s.addText('Customer Satisfied Rate', {
      x: 0.984,
      y: 0.564,
      w: 8.558,
      h: 0.656,
      fontSize: 33,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    s.addText([
      { text: T.lead, options: { breakLine: true } },
      { text: T.lead2 },
    ], { x: 0.984, y: 1.446, w: 5.761, h: 0.759, fontSize: 9, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    const cats = ['Data 1', 'Data 2', 'Data 3', 'Data 4', 'Data 5', 'Data 6', 'Data 7', 'Data 8', 'Data 9', 'Data 10'];
    s.addChart(pptx.ChartType.bar, [
      { name: 'Series 1', labels: cats, values: [54, 31, 58, 85, 85, 78, 14, 30, 17, 20] },
      { name: 'Series 2', labels: cats, values: [28, 68, 54, 58, 44, 38, 84, 29, 13, 54] },
      { name: 'Series 3', labels: cats, values: [41, 24, 80, 20, 57, 41, 33, 29, 14, 71] },
    ], { x: 1.115, y: 2.735, w: 6.531, h: 2.639, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 150, chartColors: [C.olive, C.dark, C.slate], showLegend: false, showTitle: false, catAxisHidden: true, valAxisLineShow: false, valAxisLabelFontSize: 8, valAxisLabelFontFace: F.body, valAxisLabelColor: '595959', valGridLine: { style: 'none' } });
    s.addText('2024', { x: 1.115, y: 5.704, w: 2.21, h: 0.285, fontSize: 11, fontFace: F.body, valign: 'top' });
    s.addText(T.body, { x: 1.115, y: 6.107, w: 4.156, h: 0.532, fontSize: 9, fontFace: F.body, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('2025', { x: 15.051, y: 4.602, w: 2.21, h: 0.285, fontSize: 11, fontFace: F.body, valign: 'top' });
    s.addText(T.body, { x: 15.051, y: 5.005, w: 4.156, h: 0.532, fontSize: 9, fontFace: F.body, lineSpacingMultiple: 1.5, valign: 'top' });
  },
  function slide21(s) {
    s.addShape('rect', { x: 0, y: 2.569, w: 9.806, h: 4.931, fill: { color: C.olive }, line: noLine });
    s.addShape('rect', { x: 1.115, y: 2.569, w: 7.717, h: 4.072, fill: { color: C.cream }, line: { color: C.cream, width: 0.75 } });
    s.addText('Room Amenities', { x: 0.984, y: 0.397, w: 8.558, h: 0.656, fontSize: 33, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText([
      { text: T.lead, options: { breakLine: true } },
      { text: T.lead2 },
    ], { x: 0.984, y: 1.28, w: 5.761, h: 0.759, fontSize: 9, fontFace: F.body, color: C.gray, lineSpacingMultiple: 1.5, valign: 'top' });
    const cats = ['Data 1', 'Data 2', 'Data 3', 'Data 4', 'Data 5'];
    s.addChart(pptx.ChartType.bar, [
      { name: 'Series 1', labels: cats, values: [54, 31, 58, 85, 85] },
      { name: 'Series 2', labels: cats, values: [28, 68, 54, 58, 44] },
    ], { x: 1.349, y: 3.101, w: 7.446, h: 3.008, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 150, chartColors: [C.olive, C.dark], showLegend: false, showTitle: false, catAxisHidden: true, valAxisLineShow: false, valAxisLabelFontSize: 8, valAxisLabelFontFace: F.body, valAxisLabelColor: '595959', valGridLine: { style: 'none' } });
  },
  function slide22(s) {
    s.addText('Testimonials Section', {
      x: 2.388,
      y: 0.936,
      w: 8.558,
      h: 0.656,
      fontSize: 33,
      fontFace: F.title,
      color: C.ink,
      align: 'center',
      valign: 'top'
    });
    s.addShape('rect', { x: 4.304, y: 4.633, w: 4.686, h: 1.499, fill: { color: C.white }, line: noLine });
    s.addText([
      { text: '“Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau', options: { breakLine: true } },
      { text: 'deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul”' },
    ], { x: 2.257, y: 5.016, w: 8.819, h: 1.127, fontSize: 14, fontFace: F.body, color: C.gray, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  },
  function slide23(s) {
    s.addShape('rect', { x: 0.957, y: 0.346, w: 12.049, h: 6.829, fill: { color: C.cream }, line: noLine });
    s.addText('Photographer', {
      x: 1.364,
      y: 5.543,
      w: 1.746,
      h: 0.306,
      fontSize: 9,
      fontFace: F.body,
      align: 'center',
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('Hoselar Joni', {
      x: 1.245,
      y: 0.696,
      w: 1.983,
      h: 0.37,
      fontSize: 16,
      fontFace: F.title,
      bold: true,
      align: 'center',
      valign: 'top'
    });
    s.addShape('rect', { x: 1.555, y: 5.62, w: 1.804, h: 0.48, fill: { color: C.olive }, line: noLine });
    s.addText('@accountname ', {
      x: 1.555,
      y: 5.62,
      w: 1.804,
      h: 0.48,
      fontSize: 12,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 4.94, y: 5.62, w: 1.804, h: 0.48, fill: { color: C.dark }, line: noLine });
    s.addText('@accountname ', {
      x: 4.94,
      y: 5.62,
      w: 1.804,
      h: 0.48,
      fontSize: 12,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 8.309, y: 5.62, w: 1.804, h: 0.48, fill: { color: C.slate }, line: noLine });
    s.addText('@accountname ', {
      x: 8.309,
      y: 5.62,
      w: 1.804,
      h: 0.48,
      fontSize: 12,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 0.937, y: 6.079, w: 1.28, h: 0.306, fill: { color: C.olive }, line: noLine });
    s.addText('Back office', {
      x: 0.937,
      y: 6.079,
      w: 1.28,
      h: 0.306,
      fontSize: 10,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 4.322, y: 6.079, w: 1.28, h: 0.306, fill: { color: C.dark }, line: noLine });
    s.addText('Front office', {
      x: 4.322,
      y: 6.079,
      w: 1.28,
      h: 0.306,
      fontSize: 10,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addShape('rect', { x: 7.691, y: 6.079, w: 1.28, h: 0.306, fill: { color: C.slate }, line: noLine });
    s.addText('Housekeeper', {
      x: 7.691,
      y: 6.079,
      w: 1.28,
      h: 0.306,
      fontSize: 10,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'center',
      valign: 'middle'
    });
    s.addText('Rebecca', {
      x: 4.514,
      y: 0.696,
      w: 1.983,
      h: 0.37,
      fontSize: 16,
      fontFace: F.title,
      bold: true,
      align: 'center',
      valign: 'top'
    });
    s.addText('Sheiola', {
      x: 7.919,
      y: 0.696,
      w: 1.983,
      h: 0.37,
      fontSize: 16,
      fontFace: F.title,
      bold: true,
      align: 'center',
      valign: 'top'
    });
    s.addShape('rect', { x: 11.234, y: 0.346, w: 1.791, h: 6.829, fill: { color: C.dark }, line: noLine });
    s.addText('Meet our talented team', {
      x: 10.098,
      y: 1.915,
      w: 4.063,
      h: 1.212,
      rotate: 270,
      fontSize: 33,
      fontFace: F.title,
      color: C.white,
      valign: 'top'
    });
    s.addText('TEXT TITTLE HERE', {
      x: 11.404,
      y: 5.771,
      w: 1.549,
      h: 0.286,
      fontSize: 11,
      fontFace: F.body,
      color: C.white,
      valign: 'top'
    });
    s.addText('Lorep  ipsum duis aute irure d', {
      x: 11.404,
      y: 6.058,
      w: 1.681,
      h: 0.532,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addShape('line', { x: 11.404, y: 5.507, w: 0.731, h: 0, line: { color: C.white, width: 1 } });
  },
  function slide24(s) {
    s.addShape('rect', { x: 11.565, y: 0, w: 1.768, h: 7.5, fill: { color: C.dark }, line: noLine });
    s.addShape('rect', { x: 6.048, y: 0, w: 5.549, h: 7.5, fill: { color: C.olive }, line: noLine });
    phoneMock(s, 4.831, 1.111, 2.651, 5.278);
    phoneMock(s, 7.481, 1.111, 2.651, 5.278);
    phoneMock(s, 10.132, 1.111, 2.651, 5.278);
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul', {
      x: 1.027,
      y: 5.415,
      w: 3.435,
      h: 0.986,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText([
      { text: 'Compact', options: { breakLine: true } },
      { text: 'Room', options: { breakLine: true } },
      { text: 'made for you' },
    ], { x: 0.984, y: 1.256, w: 4.463, h: 1.767, fontSize: 33, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Modern Design', { x: 1.027, y: 5.128, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
  },
  function slide25(s) {
    s.addShape('rect', { x: 6.048, y: 0, w: 6.757, h: 7.5, fill: { color: C.olive }, line: noLine });
    phoneMock(s, 6.667, 5.085, 2.601, 5.202, PHONE_SLIM);
    phoneMock(s, 9.618, 2.627, 2.601, 5.202, PHONE_SLIM);
    phoneMock(s, 9.618, -2.911, 2.601, 5.202, PHONE_SLIM);
    phoneMock(s, 6.667, -0.453, 2.601, 5.202, PHONE_SLIM);
    s.addText('Cozy Room made for you', {
      x: 0.984,
      y: 2.103,
      w: 4.463,
      h: 1.582,
      fontSize: 44,
      fontFace: F.title,
      color: C.ink,
      valign: 'top'
    });
    storeBadge(s, 1.115, 4.894, 1.608, 0.5, 'GET IT ON', 'Google Play');
    storeBadge(s, 2.854, 4.87, 1.608, 0.534, 'Download on the', 'App Store');
  },
  function slide26(s) {
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul', {
      x: 1.027,
      y: 4.152,
      w: 3.435,
      h: 0.986,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText([
      { text: 'Our Dedicated ', options: { breakLine: true } },
      { text: 'Room made for you' },
    ], { x: 0.984, y: 1.256, w: 4.463, h: 1.212, fontSize: 33, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Modern Design', { x: 1.027, y: 3.866, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addText('Button Here', {
      x: 1.027,
      y: 5.854,
      w: 1.651,
      h: 0.337,
      fontSize: 14,
      fontFace: F.body,
      color: C.ink,
      underline: { style: 'sng' },
      valign: 'top'
    });
    laptopMock(s, 4.333, -0.116, 12.125, 7.5);
  },
  function slide27(s) {
    s.addShape('rect', { x: 0, y: 0, w: 13.324, h: 3.414, fill: { color: C.dark }, line: noLine });
    s.addShape('rect', { x: 0, y: 3.414, w: 13.324, h: 4.086, fill: { color: C.olive }, line: noLine });
    s.addText('Lorep  ipsum duis aute irure dolor in kauselih oilue', {
      x: 1.027,
      y: 5.503,
      w: 3.435,
      h: 0.305,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText([
      { text: 'Swimming ', options: { breakLine: true } },
      { text: 'Pool' },
    ], { x: 0.984, y: 0.797, w: 4.463, h: 1.919, fontSize: 54, fontFace: F.title, color: C.white, valign: 'top' });
    s.addText('Modern Design', { x: 1.027, y: 5.217, w: 3.478, h: 0.286, fontSize: 11, fontFace: F.title, color: C.white, valign: 'top' });
    s.addShape('roundRect', { x: 1.115, y: 3.908, w: 0.844, h: 0.844, fill: { color: C.dark }, line: noLine });
    s.addShape('custGeom', { x: 1.319, y: 4.122, w: 0.436, h: 0.417, fill: { color: C.white }, line: noLine, points: [
      { x: 0.372, y: 0.273, moveTo: true },
      { x: 0.115, y: 0.027 },
      { x: 0.436, y: 0.31 },
      { x: 0.349, y: 0.356 },
      { x: 0.346, y: 0.258 },
      { x: 0.372, y: 0.273 },
      { close: true },
      { x: 0.064, y: 0.144, moveTo: true },
      { x: 0.321, y: 0.39 },
      { x: 0, y: 0.107 },
      { x: 0.087, y: 0.061 },
      { x: 0.09, y: 0.159 },
      { x: 0.064, y: 0.144 },
      { close: true },
      { x: 0.145, y: 0.25, moveTo: true },
      { x: 0.26, y: 0.25 },
      { x: 0.176, y: 0.229 },
      { x: 0.197, y: 0.125 },
      { x: 0.197, y: 0.104 },
      { x: 0.239, y: 0.104 },
      { x: 0.239, y: 0.125 },
      { x: 0.291, y: 0.125 },
      { x: 0.291, y: 0.167 },
      { x: 0.176, y: 0.167 },
      { x: 0.26, y: 0.188 },
      { x: 0.239, y: 0.292 },
      { x: 0.239, y: 0.313 },
      { x: 0.197, y: 0.313 },
      { x: 0.197, y: 0.292 },
      { x: 0.145, y: 0.292 },
      { x: 0.145, y: 0.25 },
      { close: true },
    ] });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec ut dolor non ipsum tincidunt luctus. Integer auctor, ', {
      x: 2.25,
      y: 3.797,
      w: 2.701,
      h: 0.988,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    tabletMock(s, 5.608, 0.15, 6.613, 8.461);
  },
  function slide28(s) {
    watchMock(s, 7.85, -1.026, 5.711, 8.943, 345.05);
    s.addText(T.rate, {
      x: 1.115,
      y: 3.733,
      w: 5.788,
      h: 0.529,
      fontSize: 9,
      fontFace: F.body,
      color: C.gray,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText([
      { text: 'Specialized', options: { breakLine: true } },
      { text: 'Sport Center', options: { breakLine: true } },
      { text: '', options: { breakLine: true } },
    ], { x: 1.115, y: 1.256, w: 4.463, h: 1.767, fontSize: 33, fontFace: F.title, color: C.ink, valign: 'top' });
    s.addShape('rect', { x: 1.115, y: 4.812, w: 3.387, h: 0.303, fill: { color: C.cream }, line: noLine });
    s.addShape('rect', { x: 1.115, y: 4.812, w: 1.854, h: 0.303, fill: { color: C.dark }, line: noLine });
    s.addText('45%', {
      x: 1.115,
      y: 4.812,
      w: 1.854,
      h: 0.303,
      fontSize: 9,
      fontFace: F.title,
      color: C.black,
      bold: true,
      align: 'right',
      valign: 'middle',
      margin: [11.34, 11.34, 3.6, 3.6]
    });
    s.addText('Your Text Here', {
      x: 1.255,
      y: 4.833,
      w: 1.141,
      h: 0.26,
      fontSize: 9,
      fontFace: F.title,
      color: C.black,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
    s.addShape('rect', { x: 1.115, y: 5.747, w: 3.387, h: 0.303, fill: { color: C.cream }, line: noLine });
    s.addShape('rect', { x: 1.115, y: 5.747, w: 3.124, h: 0.303, fill: { color: C.slate }, line: noLine });
    s.addText('90%', {
      x: 1.115,
      y: 5.747,
      w: 3.124,
      h: 0.303,
      fontSize: 9,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'right',
      valign: 'middle',
      margin: [11.34, 11.34, 3.6, 3.6]
    });
    s.addText('Your Text Here', {
      x: 1.255,
      y: 5.768,
      w: 1.189,
      h: 0.262,
      fontSize: 9,
      fontFace: F.title,
      color: C.white,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
    s.addShape('rect', { x: 1.115, y: 5.28, w: 3.387, h: 0.303, fill: { color: C.white }, line: noLine });
    s.addShape('rect', { x: 1.115, y: 5.28, w: 2.38, h: 0.303, fill: { color: C.olive }, line: noLine });
    s.addText('60%', {
      x: 1.115,
      y: 5.28,
      w: 2.38,
      h: 0.303,
      fontSize: 9,
      fontFace: F.title,
      color: C.white,
      bold: true,
      align: 'right',
      valign: 'middle',
      margin: [11.34, 11.34, 3.6, 3.6]
    });
    s.addText('Your Text Here', {
      x: 1.255,
      y: 5.302,
      w: 1.141,
      h: 0.26,
      fontSize: 9,
      fontFace: F.title,
      color: C.white,
      bold: true,
      lineSpacingMultiple: 1.3,
      paraSpaceBefore: 10,
      valign: 'top',
      wrap: false,
      margin: [0, 17.01, 2.83, 2.83]
    });
  },
  function slide29(s) {
    s.background = { color: C.olive };
    s.addText('BOOK A CALL', { x: 3.563, y: 5.658, w: 1.593, h: 0.286, fontSize: 11, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText('(+12) 345 678 90123', {
      x: 3.563,
      y: 5.944,
      w: 1.593,
      h: 0.309,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('EMAIL US', { x: 3.563, y: 4.375, w: 2.458, h: 0.286, fontSize: 11, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText('yourawesome@mail.com', {
      x: 3.563,
      y: 4.661,
      w: 2.458,
      h: 0.309,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('MEET US', { x: 1.274, y: 5.658, w: 1.593, h: 0.286, fontSize: 11, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText('123, Sesame street, Malaysia', {
      x: 1.274,
      y: 5.944,
      w: 1.593,
      h: 0.532,
      fontSize: 9,
      fontFace: F.body,
      color: C.white,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('FOLLOW US', { x: 1.274, y: 4.375, w: 1.593, h: 0.286, fontSize: 11, fontFace: F.body, color: C.white, valign: 'top' });
    s.addText([
      { text: '@yourbrandhere', options: { breakLine: true } },
      { text: '@anotherbrand' },
    ], { x: 1.274, y: 4.661, w: 1.593, h: 0.532, fontSize: 9, fontFace: F.body, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addShape('rect', { x: 6.745, y: 5.632, w: 2.018, h: 0.772, fill: { color: C.lime }, line: noLine });
    s.addText('Location', {
      x: 6.745,
      y: 5.632,
      w: 2.018,
      h: 0.772,
      fontSize: 11,
      fontFace: F.body,
      color: C.ink,
      align: 'center',
      valign: 'middle'
    });
    s.addText('Get In Touch', { x: 0.938, y: 1.242, w: 8.224, h: 1.717, fontSize: 96, fontFace: F.title, color: C.white, valign: 'top' });
  },
  function slide30(s) {
    s.addText('Thank You.', { x: 0.576, y: 3.092, w: 8.665, h: 2.036, fontSize: 115, fontFace: F.title, color: C.slate, valign: 'top' });
    s.addText('BOOK A CALL', { x: 10.734, y: 5.598, w: 1.593, h: 0.286, fontSize: 11, fontFace: F.body, color: C.slate, valign: 'top' });
    s.addText('(+12) 345 678 90123', {
      x: 10.734,
      y: 5.884,
      w: 1.593,
      h: 0.309,
      fontSize: 9,
      fontFace: F.body,
      color: C.slate,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
    s.addText('MEET US', { x: 8.444, y: 5.598, w: 1.593, h: 0.286, fontSize: 11, fontFace: F.body, color: C.slate, valign: 'top' });
    s.addText('123, Sesame street, Malaysia', {
      x: 8.444,
      y: 5.884,
      w: 1.593,
      h: 0.532,
      fontSize: 9,
      fontFace: F.body,
      color: C.slate,
      lineSpacingMultiple: 1.5,
      valign: 'top'
    });
  },
];

// ------------------------------------------------------------------- assemble
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'HD', width: 13.333, height: 7.5 });
pptx.layout = 'HD';
pptx.theme = { headFontFace: F.title, bodyFontFace: F.body };

SLIDES.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '142c80b9-b4fb-41ac-b1ab-3b7d4827a9af_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
