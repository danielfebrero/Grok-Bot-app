/**
 * Blifloe - Number One Botanical Garden
 * A 30-slide deck rebuilt with pptxgenjs. Run with `node <this file>` to write
 * the .pptx next to the script.
 *
 * Photographs in the original are replaced by flat grey placeholder blocks.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT = '0910b93d-cfa3-4b79-be34-bcde9a93f0b6_grok_final.pptx';

// -------------------------------------------------------------------- theme
const GREEN  = '85AD45';   // accent1 - leaf green
const OLIVE  = '5D7830';   // accent3 - deep green
const ORANGE = 'C98B31';   // accent2 - amber
const INK    = '3F3F3F';   // tx1
const SLATE  = '595959';   // tx2
const GREY   = '7F7F7F';   // accent6 - body copy
const SILVER = 'CCCCCC';   // page-chrome captions
const TRACK  = 'E5E5E5';   // progress-bar trough / hairline buttons
const GHOST  = 'F2F2F2';   // oversized watermark numeral
const WHITE  = 'FFFFFF';
const PHOTO  = 'D9D9D9';   // stand-in for the deck's photography

const HEAD = 'Raleway Bold';
const BODY = 'Roboto';

// Soft halo behind the white/coloured cards. pptxgenjs rewrites the shadow
// object in place, so every card gets its own copy via cardShadow().
function cardShadow(opacity) {
  return { type: 'outer', color: '000000', opacity: opacity, blur: 20, offset: 0.01, angle: 90 };
}

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'Blifloe';
pptx.title = 'Blifloe Green Garden';

const rect = pptx.ShapeType.rect;
const ellipse = pptx.ShapeType.ellipse;
const roundRect = pptx.ShapeType.roundRect;
const triangle = pptx.ShapeType.triangle;
const line = pptx.ShapeType.line;
const round2SameRect = pptx.ShapeType.round2SameRect;
const teardrop = pptx.ShapeType.teardrop;
const chevron = pptx.ShapeType.chevron;
const donut = pptx.ShapeType.donut;
const gear6 = pptx.ShapeType.gear6;
const star5 = pptx.ShapeType.star5;
const star12 = pptx.ShapeType.star12;

// ------------------------------------------------------------------ helpers

/** Text box. Mirrors the source deck's default insets and top anchoring. */
function text(s, body, o) {
  const opts = {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.fontFace || BODY,
    fontSize: o.fontSize === undefined ? 18 : o.fontSize,
    color: o.color || INK,
    bold: o.bold || false,
    italic: o.italic || false,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: [7.2, 7.2, 3.6, 3.6],   // lIns/rIns 0.1", tIns/bIns 0.05" as in the source deck
    wrap: o.wrap !== false,
    fit: o.fit || 'resize',
  };
  if (o.lineSpacing) opts.lineSpacingMultiple = o.lineSpacing;
  if (o.rotate) opts.rotate = o.rotate;
  if (o.fill) opts.fill = { color: o.fill };
  if (o.line) opts.line = o.line;
  s.addText(body, opts);
}

/** Plain geometric shape. */
function shape(s, kind, o) {
  const opts = { x: o.x, y: o.y, w: o.w, h: o.h };
  if (o.fill) opts.fill = o.fill;
  if (o.line) opts.line = o.line;
  if (o.rotate) opts.rotate = o.rotate;
  if (o.rectRadius) opts.rectRadius = o.rectRadius;
  if (o.shadow) opts.shadow = cardShadow(o.shadow);
  if (!o.fill && !o.line) opts.fill = { type: 'none' };
  s.addShape(kind, opts);
}

/** Grey block standing in for one of the deck's product photographs. */
function photo(s, x, y, w, h) {
  s.addShape(rect, { x: x, y: y, w: w, h: h, fill: { color: PHOTO }, line: { color: SILVER, width: 1 } });
  s.addText('[image]', {
    x: x, y: y + h / 2 - 0.2, w: w, h: 0.4,
    fontFace: BODY, fontSize: 11, color: GREY, align: 'center', valign: 'middle',
  });
}

/** 11pt justified body copy - the deck's standard paragraph style. */
function body(s, txt, x, y, w, h, color) {
  text(s, txt, { x: x, y: y, w: w, h: h, fontSize: 11, color: color,
                 align: 'justify', lineSpacing: 1.5 });
}

/** The amber strapline that sits under every section heading. */
function tagline(s, x, y, align) {
  text(s, 'NUMBER ONE BOTANICAL GARDEN',
    { x: x, y: y, w: 2.924, h: 0.303, fontSize: 12, color: ORANGE, bold: true,
      wrap: false, align: align || 'left' });
}

/** Section heading: a row of four dots above a two-tone Raleway title. */
function heading(s, runs, x, y, w, h, size) {
  dotRow(s, x + 0.121, y - 0.171, 0.12);
  text(s, runs, { x: x, y: y, w: w, h: h, fontSize: size || 36, fontFace: HEAD });
}

// ------------------------------------------------------- repeated page furniture

/** Amber rule + rotated caption down the left edge. */
function sideRule(s) {
  s.addShape(line, { x: 0.353, y: 0.385, w: 0, h: 1.167, line: { color: ORANGE, width: 0.5 } });
}
function sideLabel(s) {
  text(s, 'GREEN GARDEN PRESENTATION',
    { x: -0.471, y: 2.327, w: 1.648, h: 0.219, fontSize: 7, color: SILVER, wrap: false, rotate: 90 });
}
/** Short amber tab bleeding off the bottom-left corner. */
function edgeTab(s) {
  shape(s, rect, { x: 0, y: 5.143, w: 0.05, h: 2.175, fill: { color: ORANGE } });
}
/** Amber rule down the right edge. */
function pageRail(s) {
  s.addShape(line, { x: 12.999, y: 0.769, w: 0, h: 3.621, line: { color: ORANGE, width: 0.5 } });
}
/** Four stacked amber dots at the bottom-right, the second one haloed. */
function pageDots(s) {
  shape(s, ellipse, { x: 12.94, y: 6.667, w: 0.117, h: 0.117, fill: { color: ORANGE, transparency: 75 } });
  [6.431, 6.695, 6.96, 7.225].forEach(function (y) {
    shape(s, ellipse, { x: 12.968, y: y, w: 0.061, h: 0.061, fill: { color: ORANGE } });
  });
}
/** "PAGE n" / "OF 30" markers on the right rail. */
function pageNumber(s, n) {
  text(s, 'PAGE ' + n, { x: 12.734, y: 0.321, w: 0.53, h: 0.219, fontSize: 7, color: SILVER, align: 'center', wrap: false });
  text(s, 'OF 30', { x: 12.77, y: 4.619, w: 0.458, h: 0.219, fontSize: 7, color: SILVER, align: 'center', wrap: false });
}
/** Row of four dots above every section heading; the third one is amber. */
function dotRow(s, x, y, d) {
  const gap = d * 1.923;
  for (let i = 0; i < 4; i++) {
    shape(s, ellipse, { x: x + i * gap, y: y, w: d, h: d, fill: { color: i === 2 ? ORANGE : GREEN } });
  }
}

// ------------------------------------------------------------------- icons
// The source deck draws its small pictograms as freeform paths. Each one is
// redrawn here from native shapes, positioned in fractions of the icon's
// bounding box. `c` is the icon colour, `bg` the colour it sits on (used to
// knock details out of a solid silhouette).
function icon(s, kind, x, y, w, h, c, bg) {
  // put(shapeType, fx, fy, fw, fh, extra) - fractions of the icon box
  const put = function (kind_, fx, fy, fw, fh, colour, extra) {
    const o = { x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, fill: { color: colour } };
    if (extra) Object.keys(extra).forEach(function (k) { o[k] = extra[k]; });
    shape(s, kind_, o);
  };
  // Head-and-shoulders figure, shared by the people/team icons.
  const figure = function (fx, fy, fw, fh) {
    put(ellipse, fx + fw * 0.24, fy, fw * 0.52, fh * 0.40, c);
    put(round2SameRect, fx, fy + fh * 0.46, fw, fh * 0.54, c, { rectRadius: w * fw * 0.45 });
  };

  switch (kind) {
    case 'people':          // two figures side by side
      figure(0.02, 0.00, 0.44, 1.00);
      figure(0.54, 0.00, 0.44, 1.00);
      break;
    case 'team':            // globe above a row of three figures
      put(donut, 0.26, 0.00, 0.48, 0.52, c);
      figure(0.00, 0.46, 0.28, 0.54);
      figure(0.36, 0.34, 0.30, 0.66);
      figure(0.73, 0.46, 0.28, 0.54);
      break;
    case 'sprout':          // two leaves rising from a mound of soil
      put(teardrop, 0.00, 0.06, 0.44, 0.40, c, { rotate: 225 });
      put(teardrop, 0.56, 0.06, 0.44, 0.40, c, { rotate: 315 });
      put(rect, 0.46, 0.30, 0.08, 0.45, c);
      put(ellipse, 0.20, 0.66, 0.60, 0.34, c);
      break;
    case 'segments':        // segmented ring chart
      put(donut, 0.00, 0.00, 1.00, 1.00, c);
      put(rect, 0.44, 0.00, 0.12, 0.30, bg);
      put(rect, 0.44, 0.70, 0.12, 0.30, bg);
      break;
    case 'mail':            // open envelope with a bar chart on the letter
      put(rect, 0.16, 0.06, 0.68, 0.52, c);
      put(rect, 0.28, 0.16, 0.10, 0.30, bg);
      put(rect, 0.45, 0.24, 0.10, 0.22, bg);
      put(rect, 0.62, 0.12, 0.10, 0.34, bg);
      put(rect, 0.00, 0.36, 1.00, 0.64, c);
      put(triangle, 0.06, 0.40, 0.88, 0.42, bg, { rotate: 180 });
      break;
    case 'star':
      put(star5, 0.00, 0.00, 1.00, 1.00, c);
      break;
    case 'play':            // outlined circle with a chevron arrowhead
      shape(s, ellipse, { x: x, y: y, w: w, h: h, line: { color: c, width: 1.5 } });
      put(chevron, 0.28, 0.26, 0.46, 0.48, c);
      break;
    case 'process':         // gear and a figure inside a cycle of arrows
      put(gear6, 0.00, 0.42, 0.58, 0.58, c);
      figure(0.50, 0.04, 0.50, 0.60);
      break;
    case 'social':          // twitter / facebook / linkedin buttons
      put(ellipse, 0.00, 0.00, 0.27, 1.00, c);          // twitter bird
      put(teardrop, 0.05, 0.30, 0.17, 0.34, bg, { rotate: 250 });
      put(ellipse, 0.37, 0.00, 0.27, 1.00, c);          // facebook 'f'
      put(rect, 0.505, 0.22, 0.05, 0.56, bg);
      put(rect, 0.455, 0.40, 0.10, 0.10, bg);
      put(rect, 0.545, 0.22, 0.04, 0.14, bg);
      put(roundRect, 0.73, 0.00, 0.27, 1.00, c, { rectRadius: h * 0.10 });  // linkedin
      put(rect, 0.785, 0.42, 0.045, 0.36, bg);
      put(rect, 0.865, 0.46, 0.055, 0.32, bg);
      put(ellipse, 0.782, 0.22, 0.05, 0.12, bg);
      break;
    case 'badge':           // rosette with two notched ribbon tails
      put(rect, 0.18, 0.60, 0.22, 0.40, c);
      put(rect, 0.60, 0.60, 0.22, 0.40, c);
      put(triangle, 0.18, 0.88, 0.22, 0.12, bg, { rotate: 180 });
      put(triangle, 0.60, 0.88, 0.22, 0.12, bg, { rotate: 180 });
      put(star12, 0.00, 0.00, 1.00, 0.72, c);
      put(ellipse, 0.22, 0.13, 0.56, 0.42, bg);
      put(star5, 0.30, 0.18, 0.40, 0.31, c);
      break;
    case 'gear':
      put(gear6, 0.00, 0.00, 1.00, 1.00, c);
      put(ellipse, 0.30, 0.30, 0.40, 0.40, bg);
      break;
    default:
      put(ellipse, 0.00, 0.00, 1.00, 1.00, c);
  }
}

// ---------------------------------------------------------------- shared copy
const L1 =
  'Lorem ipsum dolor sit amet, consectetur adi pis cingelit. Viva mus vel euismod leo. Don ec Lorem ipsum dolor sit amet, consecte tur Lorem ipsum dolor sit amet, ';
const L2 =
  'Lorem ipsum dolor sit amet, con sectetur adipis cing elit. Vivam';
const L3 =
  'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Donec Lorem ipsum dolor sit amet, ';
const L4 =
  'dolor sit consectetur adi pis cing elit. ';
const L5 =
  'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com';
const L6 =
  'Lorem ipsum dolor sit amet, consectetur adip Donec';
const L7 =
  'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit. Viva mus vel euismod ipsum dolor sit';
const L8 =
  'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Viva mus vel euismod';
const L9 =
  'Lorem ipsum dolor sit amet, cons ectetur adi pis cing elit. ';
const L10 =
  'Lorem ipsum dolor sit amet, consectetur adip isci ';
const L11 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate consectetur adipiscing elit. Vivamus vel euismod leo. Donec';
const L12 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod';
const L13 =
  'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel eui smod leo. ';
const L14 =
  'Lorem ipsum dolor sit amet, conse ctetur adi pis';

// -------------------------------------------------------------------- slides
const SLIDES = [
  //  1. Cover
  function (s) {
  shape(s, rect, { x: 0, y: 4.801, w: 9.646, h: 2.699, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 1);
  text(s, L1, { x: 6.028, y: 5.515, w: 3.315, h: 1.212, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [
    { text: 'Bli', options: { color: ORANGE } },
    { text: 'floe', options: { color: INK } },
  ], { x: 6.374, y: 1.767, w: 5.938, h: 2.036, fontSize: 115, fontFace: HEAD });
  dotRow(s, 6.555, 1.578, 0.221);
  text(s, 'NUMBER ONE BOTANICAL GARDEN', { x: 6.47, y: 3.588, w: 4.29, h: 0.404, fontSize: 18, color: ORANGE, bold: true, wrap: false });
  },
  //  2. About Blifloe
  function (s) {
  shape(s, rect, { x: 6.667, y: 0, w: 4.021, h: 4.854, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  shape(s, rect, { x: 1.177, y: 3.75, w: 7.494, h: 3.014, fill: { color: WHITE }, shadow: 0.1 });
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Green Garden.', options: { color: ORANGE } },
  ], 1.548, 1.455, 4.055, 1.313);
  tagline(s, 1.548, 2.811);
  text(s, L1, { x: 1.623, y: 4.944, w: 3.315, h: 1.212, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  text(s, L1, { x: 5.139, y: 4.944, w: 3.315, h: 1.212, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  icon(s, 'people', 5.258, 4.454, 0.318, 0.371, GREEN, WHITE);
  text(s, 'Let’s Know About Us', { x: 5.633, y: 4.358, w: 2.214, h: 0.337, fontSize: 14, color: GREEN, bold: true });
  text(s, 'We Have Discussion Program', { x: 5.633, y: 4.543, w: 2.526, h: 0.379, fontSize: 11, color: GREY, italic: true, align: 'justify', lineSpacing: 1.5 });
  text(s, [
    { text: '2022 ', options: { fontSize: 24, bold: true } },
    { text: 'Welcome Message', options: { fontSize: 14 } },
  ], { x: 1.633, y: 4.358, w: 2.981, h: 0.505, color: GREEN });
  },
  //  3. What We Do
  function (s) {
  shape(s, rect, { x: 0, y: 0, w: 3.667, h: 5.781, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 3);
  shape(s, rect, { x: 4.656, y: 4.619, w: 3.488, h: 2.145, fill: { color: ORANGE }, shadow: 0.1 });
  shape(s, rect, { x: 8.385, y: 4.619, w: 3.488, h: 2.145, fill: { color: WHITE }, shadow: 0.1 });
  heading(s, [
    { text: 'What We Do ?', options: { color: INK, breakLine: true } },
    { text: 'Details Here.', options: { color: ORANGE } },
  ], 6.173, 1.769, 5.386, 1.313);
  tagline(s, 6.173, 3.125);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cingelit. Viva mus vel euismod leo. Don ec Lorem ipsum dolor sit amet, consecte tur Lorem ipsum dolor', 6.173, 3.554, 5.684, 0.656, GREY);
  body(s, L7, 5.305, 5.447, 2.549, 0.9, WHITE);
  text(s, 'Provides The Best ', { x: 5.305, y: 5.036, w: 2.549, h: 0.411, fontSize: 14, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  icon(s, 'segments', 4.947, 5.182, 0.285, 0.285, WHITE, ORANGE);
  body(s, L7, 9.034, 5.43, 2.549, 0.934, GREY);
  text(s, 'Giving Good Impact', { x: 9.034, y: 5.019, w: 2.549, h: 0.454, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  icon(s, 'sprout', 8.624, 5.182, 0.351, 0.22, GREEN, WHITE);
  },
  //  4. Vision
  function (s) {
  shape(s, rect, { x: 7.646, y: 6.771, w: 5.043, h: 0.729, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Vision Here.', options: { color: ORANGE } },
  ], 1.547, 1.587, 4.055, 1.313);
  tagline(s, 1.547, 2.943);
  text(s, 'Detail Vision Here.', { x: 1.547, y: 4.658, w: 2.214, h: 0.37, fontSize: 16, color: GREEN });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euism od leo. Donec Lorem ipsum dolor sit amet, cons ecte turLorem ipsum dolor sit amet, consectetur adi. ctetur adi pis cing', 1.547, 3.485, 5.429, 0.934, GREY);
  body(s, L8, 1.972, 5.15, 2.138, 0.934, GREY);
  text(s, '01', { x: 1.547, y: 5.124, w: 0.502, h: 0.411, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L8, 4.839, 5.15, 2.138, 0.934, GREY);
  text(s, '02', { x: 4.414, y: 5.124, w: 0.502, h: 0.411, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  //  5. Mission
  function (s) {
  shape(s, rect, { x: 9.708, y: 3.75, w: 3.625, h: 3.75, fill: { color: GREEN } });
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Mission Here.', options: { color: ORANGE } },
  ], 4.314, 1.701, 4.055, 1.313);
  tagline(s, 4.331, 3.057);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Donec Lorem ipsum dolor sit amet, cons ecte turLorem ipsum dolor sit amet, consectetur adi sit amet', 4.314, 4.281, 4.706, 0.934, GREY);
  text(s, 'Description Here', { x: 4.314, y: 3.834, w: 2.214, h: 0.37, fontSize: 16, color: SLATE, bold: true });
  body(s, L3, 4.314, 5.314, 4.706, 0.656, GREY);
  text(s, [
    { text: '78.01% ', options: { fontSize: 12, bold: true, breakLine: true } },
    { text: L9, options: { fontSize: 11 } },
  ], { x: 10.414, y: 4.83, w: 2.122, h: 0.959, color: WHITE, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Details Product', { x: 10.414, y: 4.397, w: 2.214, h: 0.337, fontSize: 14, color: WHITE, bold: true });
  text(s, [
    { text: '81.19% ', options: { fontSize: 12, bold: true, breakLine: true } },
    { text: L9, options: { fontSize: 11 } },
  ], { x: 10.414, y: 5.894, w: 2.122, h: 0.959, color: WHITE, align: 'justify', lineSpacing: 1.5 });
  },
  //  6. Company journey
  function (s) {
  shape(s, rect, { x: 5.917, y: 0.76, w: 7.417, h: 6.74, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  heading(s, [
    { text: 'This Is Our', options: { color: INK, breakLine: true } },
    { text: 'Company Journey.', options: { color: ORANGE } },
  ], 1.202, 1.164, 4.456, 1.313);
  tagline(s, 1.202, 2.52);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euism od leo. Donec Lorem ipsum dolor sit amet, cons ecte', 7.333, 5.617, 5.016, 0.656, WHITE);
  body(s, L10, 7.333, 4.919, 2.104, 0.656, WHITE);
  text(s, 'Blifloe Established', { x: 7.333, y: 4.599, w: 1.978, h: 0.349, fontSize: 11, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L10, 10.156, 4.919, 2.104, 0.656, WHITE);
  text(s, 'Go International', { x: 10.156, y: 4.599, w: 1.978, h: 0.344, fontSize: 11, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, '2021', { x: 7.333, y: 4.233, w: 1.283, h: 0.499, fontSize: 18, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, '2022', { x: 10.156, y: 4.216, w: 1.283, h: 0.499, fontSize: 18, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 9.959, y: 4.48, w: 0.152, h: 0.131, fill: { color: WHITE }, rotate: 90 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euism od leo. Donec Lorem', 1.202, 2.888, 4.456, 0.656, GREY);
  },
  //  7. Overview
  function (s) {
  shape(s, rect, { x: 0.708, y: 0.76, w: 3.968, h: 6.74, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 7);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Overview Here.', options: { color: ORANGE } },
  ], 7.303, 1.408, 4.055, 1.313);
  tagline(s, 7.303, 2.764);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euis mod leo. Donec Lorem ipsum dolor sit amet, consecte tur Lorem ipsum dolor sit amet, consectetur', 7.303, 4.025, 4.697, 0.934, GREY);
  icon(s, 'mail', 7.424, 3.561, 0.385, 0.383, GREEN, WHITE);
  text(s, 'Description Here', { x: 7.929, y: 3.547, w: 4.57, h: 0.411, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euis mod leo. Donec Lorem ipsum dolor sit amet, ', 7.303, 5.607, 4.697, 0.656, GREY);
  text(s, [
    { text: '$ 100.23 ', options: { bold: true } },
    { text: 'Monthly Report' },
  ], { x: 7.303, y: 5.152, w: 4.57, h: 0.411, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  },
  //  8. Be smart
  function (s) {
  shape(s, rect, { x: 7.646, y: 1.75, w: 5.688, h: 4.986, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  shape(s, rect, { x: 0, y: 4.771, w: 5.603, h: 2.722, fill: { color: ORANGE } });
  text(s, [
    { text: 'Let’s Be Smart ', options: { color: INK } },
    { text: 'Green Garden.', options: { color: ORANGE } },
  ], { x: 1.548, y: 1.455, w: 4.055, h: 1.313, fontSize: 36 });
  dotRow(s, 1.669, 1.284, 0.12);
  tagline(s, 1.548, 2.811);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euis mod leo. Donec Lorem ipsum dolor sit amet, consecte tur Lorem ipsum', 1.548, 3.718, 5.563, 0.656, GREY);
  text(s, 'Description Here', { x: 1.548, y: 3.261, w: 4.57, h: 0.411, fontSize: 14, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cin gelit. Viva mus vel euis mod leo. Donec Lorem ip sum dolor sit amet, dolor sit amet', 1.548, 5.842, 3.635, 0.934, WHITE);
  text(s, '+ 12 M ', { x: 1.548, y: 5.182, w: 2.018, h: 0.632, fontSize: 24, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  //  9. Service details
  function (s) {
  shape(s, rect, { x: 0.708, y: 1.771, w: 3.968, h: 5.729, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 9);
  shape(s, rect, { x: 1.679, y: 4.737, w: 3.968, h: 2.048, fill: { color: WHITE }, shadow: 0.1 });
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Service Details.', options: { color: ORANGE } },
  ], 7.322, 1.635, 4.055, 1.313);
  tagline(s, 7.322, 2.991);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate leo. Donec commodo et urna ac semper. Mauris finibus Donec commodo et semper. ', 7.322, 4.18, 4.564, 1.212, GREY);
  text(s, 'Our Details Service.', { x: 7.322, y: 3.661, w: 4.25, h: 0.411, fontSize: 14, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, [{ text: 'Show More', options: { underline: { style: 'sng' } } }], { x: 7.322, y: 5.67, w: 4.564, h: 0.367, fontSize: 12, color: GREEN, align: 'justify', lineSpacing: 1.5 });
  shape(s, ellipse, { x: 1.998, y: 5.167, w: 1.187, h: 1.187, fill: { color: ORANGE } });
  icon(s, 'star', 2.321, 5.49, 0.542, 0.542, WHITE, ORANGE);
  text(s, [
    { text: 'Best Service', options: { bold: true } },
    { text: '. ', options: { breakLine: true } },
    { text: 'Lorem ipsu mdolor sit am et, consectetur adipis cing elit. Vivamus adipis' },
  ], { x: 3.337, y: 5.155, w: 1.992, h: 1.212, fontSize: 11, color: GREY, italic: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 10. Plant characteristics
  function (s) {
  shape(s, rect, { x: 7.658, y: 0.76, w: 5.053, h: 6.74, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  text(s, '22', { x: 1.548, y: 3.424, w: 5.194, h: 1.561, fontSize: 66, color: GHOST, bold: true, align: 'justify', lineSpacing: 1.5 });
  heading(s, [
    { text: 'Mention Plants', options: { color: INK, breakLine: true } },
    { text: 'Characteristic', options: { color: ORANGE } },
  ], 1.548, 1.832, 6.327, 1.313);
  tagline(s, 1.548, 3.188);
  body(s, L11, 1.997, 4.628, 4.212, 1.212, GREY);
  text(s, 'Description Here', { x: 1.997, y: 4.217, w: 4.212, h: 0.411, fontSize: 14, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L12, 10.02, 2.66, 2.22, 0.934, WHITE);
  text(s, '01 Plants', { x: 10.02, y: 2.25, w: 2.02, h: 0.411, fontSize: 14, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, line, { x: 9.908, y: 2.446, w: 0, h: 0.951, line: { color: WHITE, width: 1.5 } });
  body(s, L12, 9.976, 5.327, 2.22, 0.934, WHITE);
  text(s, '02 Plants', { x: 9.976, y: 4.916, w: 2.02, h: 0.411, fontSize: 14, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, line, { x: 9.864, y: 5.113, w: 0, h: 0.951, line: { color: WHITE, width: 1.5 } });
  },
  // 11. Best choice
  function (s) {
  shape(s, rect, { x: 0, y: 3.75, w: 5.598, h: 3.062, fill: { color: GREEN } });
  text(s, [
    { text: 'Best Choice For ', options: { color: INK } },
    { text: 'Beginner Gardeners.', options: { color: ORANGE } },
  ], { x: 1.592, y: 1.074, w: 10.679, h: 0.707, fontSize: 36, align: 'center' });
  tagline(s, 5.204, 1.765, 'center');
  text(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Donec Lorem ipsum dolor sit amet, cons ecte turLorem ipsum dolor sit amet, consectetur adi pis cing elit. Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Lorem ipsum dolor it amset, cons ecte turLorem ipsum dolor sit amet, consectetur adi pis cing elit. Lorem ipsum dolor sit amet', { x: 1.857, y: 2.366, w: 9.619, h: 0.934, fontSize: 11, color: GREY, align: 'center', lineSpacing: 1.5 });
  text(s, [
    { text: '23.4', options: { fontSize: 20, bold: true } },
    { text: ' ', options: { fontSize: 24, bold: true } },
    { text: 'Client Reviews', options: { fontSize: 14, italic: true } },
  ], { x: 1.837, y: 4.195, w: 2.813, h: 0.505, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Donecconsectetur adi pis cing elit. Viva mus vel euismod leo. ', 1.33, 4.756, 3.762, 0.9, WHITE);
  icon(s, 'play', 1.431, 4.288, 0.323, 0.32, WHITE, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Doneccon', 1.33, 5.711, 3.762, 0.656, WHITE);
  },
  // 12. Easy plant life
  function (s) {
  shape(s, rect, { x: 1.625, y: 0, w: 4.042, h: 5.792, fill: { color: GREEN } });
  shape(s, rect, { x: 0.615, y: 3.821, w: 2.951, h: 2.971, fill: { color: ORANGE } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 12);
  heading(s, [
    { text: 'Recommend For', options: { color: INK, breakLine: true } },
    { text: 'Easy Plant Life.', options: { color: ORANGE } },
  ], 7.274, 1.369, 5.19, 1.313);
  tagline(s, 7.274, 2.725);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel eui smod leo. Donec Lorem ipsum dolor sit amet, cons', 7.274, 3.292, 4.693, 0.656, GREY);
  icon(s, 'process', 7.393, 4.293, 0.398, 0.426, GREEN, WHITE);
  icon(s, 'team', 7.393, 5.456, 0.365, 0.357, GREEN, WHITE);
  body(s, L13, 7.952, 4.464, 3.681, 0.656, GREY);
  text(s, 'Insert Text Title', { x: 7.968, y: 4.138, w: 4.693, h: 0.367, fontSize: 12, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L13, 7.952, 5.646, 3.681, 0.656, GREY);
  text(s, 'Insert Text Title', { x: 7.968, y: 5.32, w: 4.693, h: 0.367, fontSize: 12, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, '10 – 30 %', { x: 0.982, y: 4.325, w: 1.47, h: 0.555, fontSize: 18, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel eui smod leo. Donec Lorem ipsum', 0.982, 5.076, 2.218, 1.212, WHITE);
  body(s, 'Insert Title Here', 0.982, 4.677, 2.218, 0.344, WHITE);
  },
  // 13. Meet our team
  function (s) {
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  pageDots(s);
  pageRail(s);
  pageNumber(s, 13);
  text(s, [
    { text: 'Let’s Meet ', options: { color: INK } },
    { text: 'Our Team', options: { color: ORANGE } },
  ], { x: 4.156, y: 0.984, w: 5.228, h: 0.707, fontSize: 36, fontFace: HEAD, align: 'center', wrap: false });
  tagline(s, 5.307, 1.675, 'center');
  shape(s, rect, { x: 2.487, y: 4.515, w: 1.431, h: 0.514, fill: { color: ORANGE } });
  icon(s, 'social', 2.749, 4.648, 0.906, 0.248, WHITE, ORANGE);
  shape(s, rect, { x: 5.936, y: 4.515, w: 1.431, h: 0.514, fill: { color: GREEN } });
  icon(s, 'social', 6.199, 4.648, 0.906, 0.248, WHITE, GREEN);
  shape(s, rect, { x: 9.398, y: 4.515, w: 1.431, h: 0.514, fill: { color: ORANGE } });
  icon(s, 'social', 9.66, 4.648, 0.906, 0.248, WHITE, ORANGE);
  text(s, 'Samira Elfana', { x: 2.161, y: 5.251, w: 2.118, h: 0.454, fontSize: 14, color: GREEN, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, [{ text: 'CEO & Founder', options: { underline: { style: 'sng' } } }], { x: 2.272, y: 6.256, w: 1.895, h: 0.344, fontSize: 11, color: GREY, italic: true, align: 'center', lineSpacing: 1.5 });
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, Viva', options: { breakLine: true } },
    { text: L4 },
  ], { x: 1.601, y: 5.641, w: 3.238, h: 0.656, fontSize: 11, color: GREY, align: 'center', lineSpacing: 1.5 });
  text(s, 'Joy Seungja', { x: 5.593, y: 5.295, w: 2.118, h: 0.411, fontSize: 14, color: GREEN, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, [{ text: 'Plant Expert', options: { underline: { style: 'sng' } } }], { x: 5.704, y: 6.3, w: 1.895, h: 0.344, fontSize: 11, color: GREY, italic: true, align: 'center', lineSpacing: 1.5 });
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, Viva', options: { breakLine: true } },
    { text: L4 },
  ], { x: 5.033, y: 5.685, w: 3.238, h: 0.656, fontSize: 11, color: GREY, align: 'center', lineSpacing: 1.5 });
  text(s, 'Elaine Sivana', { x: 9.054, y: 5.338, w: 2.118, h: 0.411, fontSize: 14, color: GREEN, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, [{ text: 'Marketing Business', options: { underline: { style: 'sng' } } }], { x: 9.166, y: 6.343, w: 1.895, h: 0.344, fontSize: 11, color: GREY, italic: true, align: 'center', lineSpacing: 1.5 });
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, Viva', options: { breakLine: true } },
    { text: L4 },
  ], { x: 8.494, y: 5.729, w: 3.238, h: 0.656, fontSize: 11, color: GREY, align: 'center', lineSpacing: 1.5 });
  },
  // 14. CEO & founder
  function (s) {
  shape(s, rect, { x: 0, y: 2.75, w: 3.657, h: 4.75, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 14);
  heading(s, [
    { text: 'Meet Our ', options: { color: INK, breakLine: true } },
    { text: 'CEO & Founder.', options: { color: ORANGE } },
  ], 6.236, 1.487, 4.055, 1.313);
  tagline(s, 6.236, 2.843);
  text(s, [
    { text: 'Samira Elfana', options: { fontSize: 14, color: SLATE, bold: true } },
    { text: '. ', options: { fontSize: 14, color: GREY, bold: true } },
    { text: L5, options: { fontSize: 11, color: GREY } },
  ], { x: 6.236, y: 3.512, w: 5.504, h: 1.01, align: 'justify', lineSpacing: 1.5 });
  shape(s, roundRect, { x: 6.331, y: 5.183, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.331, y: 5.966, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.337, y: 5.183, w: 4.42, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.337, y: 5.966, w: 4.931, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  text(s, 'About Skill Ratio ', { x: 6.236, y: 4.758, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Work Experience', { x: 6.258, y: 5.574, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 10.607, y: 4.91, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '90%', { x: 10.657, y: 4.657, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 11.166, y: 5.757, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '95%', { x: 11.217, y: 5.505, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 15. Plant expert
  function (s) {
  shape(s, rect, { x: 8.625, y: 2.75, w: 4.708, h: 4.75, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'Meet Our', options: { color: INK, breakLine: true } },
    { text: 'Plant Expert.', options: { color: ORANGE } },
  ], 1.384, 1.487, 4.055, 1.313);
  tagline(s, 1.384, 2.843);
  text(s, [
    { text: 'Joy Seungja. ', options: { fontSize: 14, color: SLATE, bold: true } },
    { text: L5, options: { fontSize: 11, color: GREY } },
  ], { x: 1.384, y: 3.512, w: 5.504, h: 1.01, align: 'justify', lineSpacing: 1.5 });
  shape(s, roundRect, { x: 1.479, y: 5.183, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 1.479, y: 5.966, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 1.485, y: 5.183, w: 4.463, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 1.485, y: 5.966, w: 4.931, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  text(s, 'About Skill Ratio ', { x: 1.384, y: 4.758, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Work Experience', { x: 1.406, y: 5.574, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 5.897, y: 4.91, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '91%', { x: 5.948, y: 4.657, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 6.314, y: 5.757, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '95%', { x: 6.365, y: 5.505, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 16. Marketing business
  function (s) {
  shape(s, rect, { x: 1.708, y: 0, w: 3.979, h: 1.771, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 16);
  heading(s, [
    { text: 'Last, Meet Our', options: { color: INK, breakLine: true } },
    { text: 'Marketing Business.', options: { color: ORANGE } },
  ], 6.405, 1.487, 6.364, 1.313);
  tagline(s, 6.405, 2.843);
  text(s, [
    { text: 'Elaine Sivana', options: { fontSize: 14, color: SLATE, bold: true } },
    { text: '. ', options: { fontSize: 14, color: GREY, bold: true } },
    { text: L5, options: { fontSize: 11, color: GREY } },
  ], { x: 6.405, y: 3.512, w: 5.504, h: 1.01, align: 'justify', lineSpacing: 1.5 });
  shape(s, roundRect, { x: 6.5, y: 5.183, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.5, y: 5.966, w: 5.458, h: 0.219, fill: { color: TRACK }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.506, y: 5.183, w: 4.42, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  shape(s, roundRect, { x: 6.506, y: 5.966, w: 5.133, h: 0.219, fill: { color: GREEN }, rectRadius: 0.109 });
  text(s, 'About Skill Ratio ', { x: 6.405, y: 4.758, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Work Experience', { x: 6.427, y: 5.574, w: 5.483, h: 0.415, fontSize: 14, color: SLATE, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 10.775, y: 4.91, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '90%', { x: 10.826, y: 4.657, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 11.588, y: 5.757, w: 0.102, h: 0.088, fill: { color: GREEN }, rotate: 180 });
  text(s, '98%', { x: 11.639, y: 5.505, w: 0.865, h: 0.415, fontSize: 14, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 17. Pricing table
  function (s) {
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  pageDots(s);
  pageRail(s);
  pageNumber(s, 17);
  text(s, [
    { text: 'Blifloe’s ', options: { color: INK } },
    { text: 'Pricing Table', options: { color: ORANGE } },
  ], { x: 3.936, y: 1.036, w: 5.461, h: 0.707, fontSize: 36, fontFace: HEAD, align: 'center', wrap: false });
  tagline(s, 5.204, 1.726, 'center');
  shape(s, rect, { x: 1.982, y: 2.555, w: 2.737, h: 3.847, fill: { color: GREEN }, shadow: 0.05 });
  shape(s, rect, { x: 8.614, y: 2.555, w: 2.737, h: 3.847, fill: { color: OLIVE }, shadow: 0.05 });
  shape(s, rect, { x: 5.298, y: 2.555, w: 2.737, h: 3.847, fill: { color: ORANGE }, shadow: 0.05 });
  text(s, [{ text: 'Enim nec dui nunc', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 2.189, y: 4.5, w: 1.669, h: 0.348, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Magna sit amet risus pro', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 2.189, y: 4.799, w: 2.359, h: 0.348, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Risus pretium quam', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 2.189, y: 5.098, w: 2.201, h: 0.348, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [
    { text: '$800', options: { fontSize: 32 } },
    { text: '/services', options: { fontSize: 9 } },
  ], { x: 2.25, y: 3.548, w: 2.201, h: 0.823, color: WHITE, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, [{ text: 'Enim nec dui nunc', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 5.505, y: 4.5, w: 1.669, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Magna sit amet risus pro', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 5.505, y: 4.799, w: 2.359, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Risus pretium quam', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 5.505, y: 5.098, w: 2.201, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [
    { text: '$876', options: { fontSize: 32 } },
    { text: '/services', options: { fontSize: 9 } },
  ], { x: 5.566, y: 3.548, w: 2.201, h: 0.823, color: WHITE, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, [{ text: 'Enim nec dui nunc', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 8.82, y: 4.5, w: 1.669, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Magna sit amet risus pro', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 8.82, y: 4.799, w: 2.359, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [{ text: 'Risus pretium quam', options: { bullet: { characterCode: '2022', indent: 13.5 } } }], { x: 8.82, y: 5.098, w: 2.201, h: 0.349, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [
    { text: '$970', options: { fontSize: 32 } },
    { text: '/services', options: { fontSize: 9 } },
  ], { x: 8.882, y: 3.548, w: 2.201, h: 0.823, color: WHITE, bold: true, align: 'center', lineSpacing: 1.5 });
  text(s, 'Get It Now', { x: 2.606, y: 5.677, w: 1.49, h: 0.471, fontSize: 9, color: WHITE, bold: true, align: 'center', valign: 'middle', fit: 'none', line: { color: TRACK, width: 1 } });
  text(s, 'Get It Now', { x: 5.922, y: 5.677, w: 1.49, h: 0.471, fontSize: 9, color: WHITE, bold: true, align: 'center', valign: 'middle', fit: 'none', line: { color: TRACK, width: 1 } });
  text(s, 'Get It Now', { x: 9.237, y: 5.677, w: 1.49, h: 0.471, fontSize: 9, color: WHITE, bold: true, align: 'center', valign: 'middle', fit: 'none', line: { color: TRACK, width: 1 } });
  text(s, 'Pricing One', { x: 2.589, y: 2.948, w: 1.524, h: 0.404, fontSize: 18, color: WHITE, bold: true, align: 'center', wrap: false });
  text(s, 'Update Per Mount', { x: 2.589, y: 3.272, w: 1.524, h: 0.303, fontSize: 9, color: WHITE, italic: true, align: 'center', lineSpacing: 1.5 });
  text(s, 'Pricing Two', { x: 5.889, y: 2.948, w: 1.555, h: 0.404, fontSize: 18, color: WHITE, bold: true, align: 'center', wrap: false });
  text(s, 'Update Per Mount', { x: 5.905, y: 3.272, w: 1.524, h: 0.303, fontSize: 9, color: WHITE, italic: true, align: 'center', lineSpacing: 1.5 });
  text(s, 'Pricing Three', { x: 9.114, y: 2.948, w: 1.736, h: 0.404, fontSize: 18, color: WHITE, bold: true, align: 'center', wrap: false });
  text(s, 'Update Per Mount', { x: 9.221, y: 3.272, w: 1.524, h: 0.303, fontSize: 9, color: WHITE, italic: true, align: 'center', lineSpacing: 1.5 });
  },
  // 18. Our products
  function (s) {
  shape(s, rect, { x: 0, y: 4.729, w: 13.333, h: 2.771, fill: { color: GREEN } });
  shape(s, rect, { x: 6.774, y: 2.792, w: 2.849, h: 4, fill: { color: ORANGE } });
  text(s, [
    { text: 'Let’s See ', options: { color: INK } },
    { text: 'Our Products.', options: { color: ORANGE } },
  ], { x: 3.995, y: 1.156, w: 5.55, h: 0.707, fontSize: 36, fontFace: HEAD, align: 'center', wrap: false });
  tagline(s, 5.307, 1.847, 'center');
  body(s, L6, 7.004, 3.527, 2.389, 0.622, WHITE);
  text(s, 'Product 01', { x: 7.004, y: 3.184, w: 2.227, h: 0.404, fontSize: 12, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L6, 7.004, 4.652, 2.389, 0.622, WHITE);
  text(s, 'Product 02', { x: 7.004, y: 4.309, w: 2.227, h: 0.367, fontSize: 12, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L6, 7.004, 5.777, 2.389, 0.622, WHITE);
  text(s, 'Product 03', { x: 7.004, y: 5.434, w: 2.227, h: 0.367, fontSize: 12, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 19. Many kinds of plants
  function (s) {
  shape(s, rect, { x: 1.678, y: 0, w: 3.968, h: 7.5, fill: { color: GREEN } });
  shape(s, rect, { x: 7.892, y: 4.737, w: 3.968, h: 2.048, fill: { color: ORANGE }, shadow: 0.1 });
  shape(s, rect, { x: 3.715, y: 4.737, w: 3.968, h: 2.048, fill: { color: WHITE }, shadow: 0.1 });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 19);
  heading(s, [
    { text: 'We Have Many', options: { color: INK, breakLine: true } },
    { text: 'Kind Of Plants.', options: { color: ORANGE } },
  ], 6.207, 1.668, 4.055, 1.313);
  tagline(s, 6.207, 3.024);
  text(s, '223.12 K', { x: 4.04, y: 5.022, w: 1.521, h: 0.544, fontSize: 20, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L3, 4.04, 5.565, 3.319, 0.934, GREY);
  text(s, '+129.1 %', { x: 8.217, y: 5.022, w: 1.521, h: 0.544, fontSize: 20, color: WHITE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, L3, 8.217, 5.565, 3.319, 0.934, WHITE);
  text(s, [
    { text: '90% - ', options: { fontSize: 14, bold: true } },
    { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id , ', options: { fontSize: 11 } },
  ], { x: 6.207, y: 3.536, w: 5.654, h: 0.732, color: GREY, align: 'justify', lineSpacing: 1.5 });
  },
  // 20. The equipment
  function (s) {
  shape(s, rect, { x: 6.661, y: 0, w: 3.964, h: 5.833, fill: { color: GREEN } });
  shape(s, rect, { x: 1.469, y: 3.709, w: 7.21, h: 3.034, fill: { color: WHITE }, shadow: 0.1 });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'Blifloe Also Provides', options: { color: INK, breakLine: true } },
    { text: 'The Equipment.', options: { color: ORANGE } },
  ], 1.548, 1.455, 6.53, 1.313);
  tagline(s, 1.548, 2.811);
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel eui smod leo. Donec com modo et', 4.329, 4.835, 3.85, 0.625, GREY);
  text(s, 'Hanna Wilson ', { x: 4.329, y: 4.297, w: 1.722, h: 0.455, fontSize: 16, color: GREEN, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, '91 %', { x: 6.878, y: 4.275, w: 1.226, h: 0.499, fontSize: 18, color: GREEN, bold: true, align: 'right', lineSpacing: 1.5 });
  text(s, [{ text: L14, options: { bullet: { characterCode: '00FC', indent: 13.5 } } }], { x: 4.329, y: 5.479, w: 3.775, h: 0.344, fontSize: 11, color: GREY, align: 'justify', lineSpacing: 1.5 });
  text(s, [{ text: L14, options: { bullet: { characterCode: '00FC', indent: 13.5 } } }], { x: 4.329, y: 5.817, w: 3.775, h: 0.344, fontSize: 11, color: GREY, align: 'justify', lineSpacing: 1.5 });
  },
  // 21. Achievement
  function (s) {
  shape(s, rect, { x: 5.583, y: 0.75, w: 7.75, h: 4.047, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'Blifloe', options: { color: INK, breakLine: true } },
    { text: 'Achievement.', options: { color: ORANGE } },
  ], 1.298, 1.455, 4.055, 1.313);
  tagline(s, 1.298, 2.811);
  shape(s, rect, { x: 8.669, y: 4.274, w: 4.018, h: 1.976, fill: { color: WHITE }, shadow: 0.1 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Viva mus vel euismod leo. Donec Lorem ipsum dolor sit amet, cons ecte turLorem ipsum dolor sit amet, consectetur adi pis cing elit. Lorem ipsum dolor sit amet, ', 6.219, 2.396, 5.615, 0.934, WHITE);
  text(s, [
    { text: 'THE ARE SEVERALS FACTORS, WHY OUR ', options: { breakLine: true } },
    { text: 'PRODUCT ARE THE BEST IN OUR COUNTRY !' },
  ], { x: 6.696, y: 1.522, w: 5.615, h: 0.764, fontSize: 14, color: WHITE, align: 'justify', lineSpacing: 1.5 });
  icon(s, 'badge', 6.219, 1.648, 0.381, 0.513, WHITE, GREEN);
  shape(s, ellipse, { x: 9.04, y: 4.783, w: 1.005, h: 1.005, fill: { color: ORANGE } });
  icon(s, 'gear', 9.337, 5.081, 0.41, 0.41, WHITE, ORANGE);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', { x: 10.224, y: 5.212, w: 2.093, h: 0.656, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  text(s, '+6175', { x: 10.224, y: 4.656, w: 1.946, h: 0.632, fontSize: 24, color: SLATE, bold: true, lineSpacing: 1.5 });
  },
  // 22. Knowledge
  function (s) {
  shape(s, rect, { x: 1.667, y: 3.75, w: 4.062, h: 3.75, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 22);
  heading(s, [
    { text: 'Knowladge About', options: { color: INK, breakLine: true } },
    { text: 'Gardeners Is Here.', options: { color: ORANGE } },
  ], 6.303, 1.221, 5.76, 1.313);
  tagline(s, 6.303, 2.577);
  text(s, 'Show More', { x: 10.661, y: 3.369, w: 1.092, h: 0.348, fontSize: 12, color: GREEN, bold: true, align: 'center', valign: 'middle', fit: 'none', line: { color: GREEN, width: 1 } });
  body(s, L11, 6.303, 3.87, 5.516, 0.934, GREY);
  text(s, 'For Beginner Gardeners', { x: 6.303, y: 3.316, w: 4.212, h: 0.455, fontSize: 16, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Detail Here', { x: 10.661, y: 5.189, w: 1.092, h: 0.348, fontSize: 12, color: GREEN, bold: true, align: 'center', valign: 'middle', fit: 'none', line: { color: GREEN, width: 1 } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id', 6.303, 5.667, 5.516, 0.656, GREY);
  text(s, 'Become An Expert Gardeners', { x: 6.303, y: 5.11, w: 4.212, h: 0.455, fontSize: 16, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 23. Innovation & discussion
  function (s) {
  shape(s, rect, { x: 6.667, y: 0.708, w: 2.989, h: 4.062, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'Our Innovation &', options: { color: INK, breakLine: true } },
    { text: 'Project Discussion', options: { color: ORANGE } },
  ], 1.548, 1.447, 4.631, 1.313);
  tagline(s, 1.548, 2.803);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. ', 1.548, 3.924, 4.614, 0.656, GREY);
  text(s, 'March 10, New Innovation', { x: 1.548, y: 3.47, w: 4.212, h: 0.411, fontSize: 14, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris consec tetur adipiscing elit. Vivamus vel', 1.548, 5.29, 4.614, 0.934, GREY);
  text(s, 'August 22, Project Discussion', { x: 1.548, y: 4.836, w: 4.212, h: 0.411, fontSize: 14, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 24. Take a break
  function (s) {
  shape(s, rect, { x: 2.698, y: 6.771, w: 7.969, h: 0.729, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 24);
  heading(s, [
    { text: 'It’s Time To', options: { color: INK, breakLine: true } },
    { text: 'Take A Break.', options: { color: ORANGE } },
  ], 6.838, 1.556, 4.989, 1.582, 44);
  tagline(s, 6.838, 3.138);
  text(s, 'Break Estimations', { x: 7.598, y: 4.021, w: 3.178, h: 0.505, fontSize: 16, color: SLATE, bold: true, align: 'justify', lineSpacing: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viv amus vel euismod leo. Donec commodo et urna ac. Donec commodo et urna ac semper. semper', 7.598, 4.556, 4.229, 0.934, GREY);
  text(s, '30’', { x: 6.838, y: 3.93, w: 0.932, h: 0.729, fontSize: 28, color: ORANGE, bold: true, align: 'justify', lineSpacing: 1.5 });
  },
  // 25. Phone device
  function (s) {
  shape(s, rect, { x: 8.717, y: 0.75, w: 3.995, h: 6.75, fill: { color: GREEN } });
  photo(s, 5.462, 0.999, 5.903, 5.903);
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Phone Device.', options: { color: ORANGE } },
  ], 1.548, 1.76, 4.055, 1.313);
  tagline(s, 1.548, 3.117);
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec ', options: { color: GREY } },
    { text: 'commodo', options: { color: GREY } },
    { text: ' et urna ac semper. Mauris finibus augue id , consectetur adipiscing elit. Vivamus vel euismod leo. Donec amet, consectetur adipiscing elit. Vivamus vel', options: { color: GREY } },
  ], { x: 1.548, y: 3.829, w: 4.747, h: 1.212, fontSize: 11, align: 'justify', lineSpacing: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et', { x: 2.01, y: 5.284, w: 4.206, h: 0.627, fontSize: 11, color: GREY, italic: true, align: 'justify', lineSpacing: 1.5 });
  shape(s, triangle, { x: 1.653, y: 5.496, w: 0.236, h: 0.203, fill: { color: GREEN }, rotate: 90 });
  },
  // 26. Laptop device
  function (s) {
  shape(s, rect, { x: 0.608, y: 0.769, w: 4.142, h: 6.731, fill: { color: GREEN } });
  photo(s, 1.016, 2.442, 5.894, 3.559);
  pageDots(s);
  pageRail(s);
  pageNumber(s, 26);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Laptop Device.', options: { color: ORANGE } },
  ], 7.1, 1.452, 4.055, 1.313);
  tagline(s, 7.1, 2.809);
  text(s, [
    { text: '1' },
    { text: 'st', options: { superscript: true } },
    { text: '  Website' },
  ], { x: 7.1, y: 3.543, w: 1.603, h: 0.337, fontSize: 14, color: GREEN, bold: true });
  body(s, L2, 7.1, 4.086, 2.478, 0.656, GREY);
  text(s, [
    { text: '2' },
    { text: 'nd', options: { superscript: true } },
    { text: ' Instagram' },
  ], { x: 9.769, y: 3.543, w: 1.603, h: 0.337, fontSize: 14, color: GREEN, bold: true });
  body(s, L2, 9.769, 4.086, 2.478, 0.656, GREY);
  text(s, [
    { text: '3' },
    { text: 'rd', options: { superscript: true } },
    { text: ' Facebook' },
  ], { x: 7.1, y: 5.019, w: 1.603, h: 0.337, fontSize: 14, color: GREEN, bold: true });
  body(s, L2, 7.1, 5.563, 2.478, 0.656, GREY);
  text(s, [
    { text: '4' },
    { text: 'th', options: { superscript: true } },
    { text: ' Youtube' },
  ], { x: 9.769, y: 5.019, w: 1.603, h: 0.337, fontSize: 14, color: GREEN, bold: true });
  body(s, L2, 9.769, 5.563, 2.478, 0.656, GREY);
  body(s, 'Blifloegreengarden.domain', 7.1, 3.707, 2.478, 0.379, GREY);
  body(s, '@blifloe.garden', 9.769, 3.707, 2.478, 0.344, GREY);
  body(s, 'm.facebook/blifloegarden', 7.1, 5.196, 2.478, 0.379, GREY);
  body(s, 'Blifloe green garden', 9.769, 5.196, 2.478, 0.344, GREY);
  },
  // 27. Computer device
  function (s) {
  shape(s, rect, { x: 6.667, y: 0.75, w: 3.617, h: 6.75, fill: { color: GREEN } });
  photo(s, 7.457, 1.867, 5.632, 4.667);
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Computer Device.', options: { color: ORANGE } },
  ], 1.463, 1.734, 4.706, 1.313);
  tagline(s, 1.463, 3.09);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivam us vel euismod leo. Donec commodo et urna ac semp er. ipsum dolor sit amet, adipis cing elit. Vivam us vel euismod leo. Donec commodo et urna ac sadipis cing elit. Vivam us vel', 1.463, 4.726, 4.706, 1.212, GREY);
  text(s, 'Subtitle Here', { x: 1.463, y: 4.274, w: 2.603, h: 0.411, fontSize: 14, color: SLATE, bold: true, lineSpacing: 1.5 });
  text(s, '102,13 K', { x: 1.463, y: 3.706, w: 1.934, h: 0.707, fontSize: 24, color: GREEN, bold: true, lineSpacing: 1.5 });
  },
  // 28. Quote
  function (s) {
  shape(s, rect, { x: 1.74, y: 3.75, w: 3.969, h: 3.75, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 28);
  text(s, '“Happiness will grow if you plant the seeds of love in the garden of hope with compassion and care”', { x: 6.495, y: 2.407, w: 5.1, h: 1.553, fontSize: 20, color: INK, bold: true, italic: true, fontFace: HEAD, lineSpacing: 1.5 });
  text(s, 'Debasish Mridha', { x: 8.746, y: 4.593, w: 2.849, h: 0.499, fontSize: 18, color: GREEN, bold: true, align: 'right', lineSpacing: 1.5 });
  },
  // 29. Contact person
  function (s) {
  shape(s, rect, { x: 8.656, y: 2.688, w: 4.677, h: 4.812, fill: { color: GREEN } });
  sideRule(s);
  sideLabel(s);
  edgeTab(s);
  heading(s, [
    { text: 'About Blifloe ', options: { color: INK } },
    { text: 'Contact Person.', options: { color: ORANGE } },
  ], 1.548, 1.705, 4.055, 1.313);
  tagline(s, 1.548, 3.061);
  text(s, [
    { text: '+01 2333 1222 21', options: { breakLine: true } },
    { text: '+01 2333 1777 21' },
  ], { x: 3.915, y: 5.31, w: 2.143, h: 0.656, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  text(s, 'Offices Location', { x: 3.915, y: 3.798, w: 2.296, h: 0.46, fontSize: 16, color: SLATE, bold: true, lineSpacing: 1.5 });
  text(s, [
    { text: 'Green Garden Company, 1' },
    { text: 'st', options: { superscript: true } },
    { text: ' Floor, San Francsico, USA' },
  ], { x: 3.915, y: 4.169, w: 2.143, h: 0.656, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  text(s, 'E-mail & Media', { x: 1.548, y: 4.94, w: 2.296, h: 0.46, fontSize: 16, color: SLATE, bold: true, lineSpacing: 1.5 });
  text(s, [
    { text: 'greengarden@domain.com  ', options: { breakLine: true } },
    { text: '@greengarden.official' },
  ], { x: 1.548, y: 5.31, w: 2.491, h: 0.656, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  text(s, 'Phone Number', { x: 3.915, y: 4.94, w: 2.296, h: 0.46, fontSize: 16, color: SLATE, bold: true, lineSpacing: 1.5 });
  text(s, 'Office Hours', { x: 1.548, y: 3.798, w: 2.296, h: 0.46, fontSize: 16, color: SLATE, bold: true, lineSpacing: 1.5 });
  text(s, [
    { text: 'Sunday – Friday', options: { breakLine: true } },
    { text: '08.00 AM – 05.00 PM' },
  ], { x: 1.548, y: 4.169, w: 2.491, h: 0.625, fontSize: 11, color: GREY, lineSpacing: 1.5 });
  },
  // 30. Thanks
  function (s) {
  shape(s, rect, { x: 0, y: 4.801, w: 9.646, h: 2.699, fill: { color: GREEN } });
  pageDots(s);
  pageRail(s);
  pageNumber(s, 30);
  text(s, L1, { x: 6.028, y: 5.515, w: 3.315, h: 1.212, fontSize: 11, color: WHITE, lineSpacing: 1.5 });
  text(s, [
    { text: 'Tha', options: { color: ORANGE } },
    { text: 'nks', options: { color: INK } },
  ], { x: 6.374, y: 1.767, w: 5.938, h: 2.036, fontSize: 115, fontFace: HEAD });
  dotRow(s, 6.555, 1.578, 0.221);
  text(s, 'AND SEE YOU NEXT TIME', { x: 6.47, y: 3.63, w: 4.111, h: 0.505, fontSize: 24, color: ORANGE, bold: true, wrap: false });
  },
];

SLIDES.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, OUT) })
  .then(function (f) { console.log('wrote ' + f); });
