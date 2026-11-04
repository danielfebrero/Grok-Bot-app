/**
 * "Maklaar" real-estate deck (30 slides, 13.333" x 7.5") rebuilt with pptxgenjs.
 *
 * Photographs in the source deck are replaced by flat grey placeholder rectangles.
 * Gradient fills (not supported by pptxgenjs) are approximated with stacked
 * translucent bands - see gradientFade().
 *
 *   node 13e1080b-3796-4f11-aca5-fd9770901510_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ tokens */

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const NAVY = '224775';
const RED = 'BE1300';
const INK = '404040'; // body copy (black @ 75% lum)
const INK_DARK = '262626';
const WHITE = 'FFFFFF';
const CREAM = 'EEECE1';
const BG = 'F5F5F5';
const PHOTO = '6F6F6F'; // stand-in for every photograph
const RULE = 'D6D4CB'; // cream @ 90% lum - thin card outlines
const DASH = 'BFBFBF';

const HEAD = 'Open Sans'; // theme major font
const BODY = 'Open Sans Light'; // theme minor font

/* ------------------------------------------------------------- copy blocks */

const LOREM_LONG =
  '"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore ' +
  'et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum ' +
  'dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui ' +
  'officia deserunt mollit anim id est laborum."';

const LOREM_MED =
  '"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore ' +
  'et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum ' +
  'dolore eu fugiat nulla pariatur. ';

const LOREM_SHORT =
  '"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore ' +
  'et dolore magna aliqua. ';

const LOREM_PLAIN =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ' +
  'ex ea commodo consequat. ';

const FUSCE_LONG =
  'Fusce a felis porta, ultrices arcu eu, ultricies magna. Aliquam sed ipsum maximus est euismod dignissim. ' +
  'Maecenas consequat elementum eros, ut pharetra velit convallis nec. Vivamus interdum porta varius. ' +
  'Nam quis libero ipsum. Quisque ut nibh ante. Vestibulum sit amet nisi at leo tempor facilisis.';

const FUSCE_SHORT =
  'Fusce a felis porta, ultrices arcu eu, ultricies magna. Aliquam sed ipsum maximus est euismod dignissim. ';

const ADDRESS = '11740 River Hills Pkwy, Rockton, IL 61072';

/* ----------------------------------------------------------- text helpers */

function text(slide, body, opts) {
  slide.addText(
    body,
    Object.assign(
      { fontFace: BODY, color: INK, fontSize: 12, margin: 0, valign: 'top', wrap: true, isTextBox: true },
      opts
    )
  );
}

/** Small red letter-spaced kicker ("Subtitle texts" / "Title texts"). */
function kicker(slide, x, y, w, label, opts) {
  text(slide, label || 'Subtitle texts', Object.assign({
    x, y, w, h: 0.283, fontSize: 14, color: RED, charSpacing: 4, lineSpacingMultiple: 1.2,
  }, opts));
}

/** Navy section headline set in the major font. */
function heading(slide, x, y, w, h, lines, opts) {
  const runs = [].concat(lines).map((t, i, all) => ({ text: t, options: { breakLine: i < all.length - 1 } }));
  text(slide, runs, Object.assign({ x, y, w, h, fontFace: HEAD, fontSize: 28, color: NAVY }, opts));
}

/** Grey paragraph copy at 150% leading. */
function paragraph(slide, x, y, w, h, body, opts) {
  text(slide, body, Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.5 }, opts));
}

/** "Maklaar" two-tone wordmark; the header instance is always in the same spot. */
function wordmark(slide, x, y, w, size) {
  text(slide, [
    { text: 'Ma', options: { color: NAVY } },
    { text: 'klaar', options: { color: RED } },
  ], { x, y, w, h: size / 54, fontSize: size, charSpacing: 2, align: 'center', valign: 'middle' });
}

function header(slide) {
  wordmark(slide, 5.838, 0.402, 1.657, 12);
}

/** Red / navy / red rounded-square triad used as a section marker. */
function dots(slide, x, y, w) {
  const k = (w || 0.352) / 0.352;
  [[0, 0.0035, 0.073, RED], [0.135, 0, 0.08, NAVY], [0.279, 0.0035, 0.073, RED]].forEach(([dx, dy, s, c]) => {
    slide.addShape('roundRect', {
      x: x + dx * k, y: y + dy * k, w: s * k, h: s * k, fill: { color: c }, rectRadius: s * k * 0.28,
    });
  });
}

/* --------------------------------------------------------- shape helpers */

function rect(slide, x, y, w, h, fill, opts) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: typeof fill === 'string' ? { color: fill } : fill }, opts));
}

/** Grey stand-in for a photograph. */
function photo(slide, x, y, w, h, rounded) {
  if (rounded) {
    slide.addShape('roundRect', { x, y, w, h, fill: { color: PHOTO }, rectRadius: w * 0.12 });
  } else {
    rect(slide, x, y, w, h, PHOTO);
  }
  if (w >= 2 && h >= 1.1) {
    text(slide, '[image]', { x, y: y + h / 2 - 0.13, w, h: 0.26, align: 'center', color: '7C7C7C', fontSize: 9 });
  }
}

/**
 * pptxgenjs cannot emit <a:gradFill>, so a gradient is drawn as N stacked
 * slices. `stops` are [position 0..1, hex colour, opacity 0..100]; each slice is
 * pre-composited against `under` (the flat colour it sits on) and painted
 * opaque, which keeps the ramp free of seams.
 */
function gradientFade(slide, x, y, w, h, stops, dir, steps, under) {
  const n = steps || 16;
  const backdrop = under || PHOTO;
  const horizontal = dir === 'h';
  const span = horizontal ? w : h;
  for (let i = 0; i < n; i++) {
    const t = (i + 0.5) / n;
    let a = stops[0], b = stops[stops.length - 1];
    for (let s = 0; s < stops.length - 1; s++) {
      if (t >= stops[s][0] && t <= stops[s + 1][0]) { a = stops[s]; b = stops[s + 1]; }
    }
    const f = b[0] === a[0] ? 0 : Math.max(0, Math.min(1, (t - a[0]) / (b[0] - a[0])));
    const opacity = a[2] + (b[2] - a[2]) * f;
    const color = mixHex(backdrop, mixHex(a[1], b[1], f), opacity / 100);
    const lo = (span * i) / n;
    const hi = Math.min(span, (span * (i + 1)) / n + span / (n * 4));
    rect(slide,
      horizontal ? x + lo : x,
      horizontal ? y : y + lo,
      horizontal ? hi - lo : w,
      horizontal ? h : hi - lo,
      color);
  }
}

function mixHex(c1, c2, f) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const v = Math.round(parseInt(c1.substr(i, 2), 16) * (1 - f) + parseInt(c2.substr(i, 2), 16) * f);
    out += v.toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

/** Navy scrim that fades in towards the bottom of a photo. */
function scrim(slide, x, y, w, h, opacity) {
  gradientFade(slide, x, y, w, h, [[0, RED, 0], [1, NAVY, opacity]], 'v', 12);
}

/** Filled house silhouette (roof with eaves + body) - the Maklaar mark. */
function houseGlyph(slide, x, y, w, h, color, outline) {
  const P = [[0.5, 0], [1, 0.42], [0.86, 0.42], [0.86, 1], [0.14, 1], [0.14, 0.42], [0, 0.42]];
  slide.addShape('custGeom', {
    x, y, w, h,
    points: P.map(([px, py]) => ({ x: px * w, y: py * h })).concat([{ close: true }]),
    fill: outline ? { type: 'none' } : { color },
    line: outline ? { color, width: 1.25 } : undefined,
  });
}

/* -------------------------------------------------------- device mock-ups */

/** Smartphone body; `screen` is where the (placeholder) photo goes. */
function phoneMock(slide, x, y, w, h, screen) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: '1A1A1A' }, rectRadius: w * 0.155 });
  slide.addShape('roundRect', {
    x: x + w * 0.024, y: y + h * 0.011, w: w * 0.952, h: h * 0.978,
    fill: { color: '000000' }, rectRadius: w * 0.145,
  });
  photo(slide, screen[0], screen[1], screen[2], screen[3], true);
  // notch: black bar with a rounded lower edge, plus speaker slot and lens
  slide.addShape('round2SameRect', {
    x: x + w * 0.29, y, w: w * 0.42, h: h * 0.05, fill: { color: '000000' }, rectRadius: w * 0.03, rotate: 180,
  });
  slide.addShape('roundRect', {
    x: x + w * 0.4, y: y + h * 0.023, w: w * 0.12, h: h * 0.006, fill: { color: '3A3033' }, rectRadius: 0.01,
  });
  slide.addShape('ellipse', { x: x + w * 0.6, y: y + h * 0.019, w: w * 0.028, h: w * 0.028, fill: { color: '223349' } });
}

/** Open laptop drawn from the group's bounding box. */
function laptopMock(slide, x, y, w, h) {
  const box = (fx, fy, fw, fh) => [x + fx * w, y + fy * h, fw * w, fh * h];
  const lid = box(0.0916, 0, 0.8181, 0.945);
  slide.addShape('roundRect', { x: lid[0], y: lid[1], w: lid[2], h: lid[3], fill: { color: '000000' }, rectRadius: w * 0.014 });
  rect(slide, ...box(0.0968, 0.9252, 0.805, 0.0408), '181818');
  rect(slide, ...box(0.1245, 0.0721, 0.7519, 0.8159), '262626');
  rect(slide, ...box(0, 0.966, 1.0, 0.034), 'C4C5C7');
  rect(slide, ...box(0.4045, 0.966, 0.1911, 0.012), 'A5A6A8');
}

/* -------------------------------------------------------------- line icons */

/**
 * Simplified single-colour pictograms built from native shapes.
 * Every icon is drawn inside the unit box (x, y, w, h).
 */
function icon(slide, kind, x, y, w, h, color) {
  const stroke = { color, width: 1 };
  const none = { type: 'none' };
  /** outlined shape placed by unit fractions of the icon box */
  const sh = (shape, fx, fy, fw, fh, opts) => slide.addShape(shape, Object.assign(
    { x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, fill: none, line: stroke }, opts));
  /** solid bar (also used for the thin rules inside documents) */
  const bar = (fx, fy, fw, fh) => sh('rect', fx, fy, fw, fh, { fill: { color }, line: none });
  /** head-and-shoulders bust */
  const bust = (fx, fw, c) => {
    const ln = { color: c || color, width: 1 };
    sh('ellipse', fx + fw * 0.2, 0.0, fw * 0.6, 0.46, { line: ln });
    sh('round2SameRect', fx, 0.5, fw, 0.5, { line: ln, rectRadius: fw * w * 0.3 });
  };

  switch (kind) {
    case 'signpost': // two direction boards on a post
      bar(0.48, 0, 0.04, 1);
      sh('homePlate', 0.22, 0.06, 0.72, 0.26);
      sh('homePlate', 0.06, 0.4, 0.72, 0.26, { flipH: true });
      break;
    case 'people':
      bust(0.0, 0.6);
      bust(0.42, 0.58, NAVY);
      break;
    case 'personCheck':
      bust(0.0, 0.66);
      sh('ellipse', 0.52, 0.5, 0.48, 0.48, { line: { color: NAVY, width: 1 } });
      sh('line', 0.62, 0.72, 0.09, 0.08, { line: { color: NAVY, width: 1 } });
      sh('line', 0.71, 0.62, 0.16, 0.18, { line: { color: NAVY, width: 1 }, flipV: true });
      break;
    case 'personSearch':
      bust(0.0, 0.66);
      sh('ellipse', 0.5, 0.48, 0.34, 0.34, { line: { color: NAVY, width: 1 } });
      sh('line', 0.8, 0.78, 0.16, 0.2, { line: { color: NAVY, width: 1.3 } });
      break;
    case 'houseLine': // outlined house with chimney, door and window
      houseGlyph(slide, x, y + h * 0.14, w, h * 0.86, color, true);
      bar(0.7, 0.05, 0.08, 0.25);
      sh('rect', 0.24, 0.66, 0.16, 0.34, { line: { color, width: 0.75 } });
      sh('rect', 0.52, 0.66, 0.2, 0.18, { line: { color, width: 0.75 } });
      break;
    case 'houseSolid':
      houseGlyph(slide, x, y, w, h, color, false);
      break;
    case 'list': // notepad with ruled lines and a bound top edge
      sh('rect', 0.06, 0.1, 0.88, 0.9);
      for (let i = 0; i < 5; i++) bar(0.2, 0.32 + i * 0.13, 0.6, 0.035);
      for (let i = 0; i < 3; i++) bar(0.24 + i * 0.24, 0.0, 0.08, 0.16);
      break;
    case 'doc': // page with a folded corner
      sh('snip1Rect', 0.16, 0.02, 0.68, 0.96);
      for (let i = 0; i < 4; i++) bar(0.28, 0.36 + i * 0.14, 0.44, 0.035);
      break;
    case 'calc':
      sh('roundRect', 0.16, 0.0, 0.68, 1.0, { rectRadius: w * 0.05 });
      for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) bar(0.27 + c * 0.16, 0.22 + r * 0.16, 0.09, 0.1);
      break;
    case 'speech':
      sh('wedgeRoundRectCallout', 0.0, 0.0, 0.66, 0.58, { rectRadius: w * 0.12 });
      sh('wedgeRoundRectCallout', 0.34, 0.42, 0.62, 0.5, { rectRadius: w * 0.12, flipH: true });
      break;
    case 'pen': // signed document with a pen resting across it
      sh('snip1Rect', 0.38, 0.0, 0.6, 0.66);
      sh('rect', 0.0, 0.62, 0.14, 0.3);
      sh('homePlate', 0.14, 0.66, 0.56, 0.22);
      break;
    case 'check':
      sh('ellipse', 0.06, 0.06, 0.88, 0.88);
      sh('line', 0.25, 0.5, 0.16, 0.16);
      sh('line', 0.41, 0.34, 0.32, 0.32, { flipV: true });
      break;
    case 'money': // curved arrow dropping onto a banknote
      sh('blockArc', 0.2, 0.0, 0.8, 0.86, { angleRange: [200, 100], arcThicknessRatio: 0.18, fill: { color }, line: none });
      sh('triangle', 0.14, 0.16, 0.24, 0.24, { rotate: 180, fill: { color }, line: none });
      sh('parallelogram', 0.0, 0.52, 0.62, 0.4, { line: { color: NAVY, width: 1 } });
      sh('ellipse', 0.24, 0.63, 0.14, 0.18, { line: { color: NAVY, width: 0.75 } });
      break;
    case 'drop':
      sh('teardrop', 0.1, 0.02, 0.8, 0.96, { rotate: 315 });
      break;
    case 'expand': // square with a diagonal out-arrow
      sh('rect', 0.0, 0.46, 0.5, 0.52, { line: { color, width: 0.75, dashType: 'dash' } });
      sh('rect', 0.32, 0.0, 0.68, 0.68);
      sh('line', 0.44, 0.12, 0.42, 0.42, { line: { color, width: 1, endArrowType: 'triangle' }, flipV: true });
      break;
    case 'door':
      sh('rect', 0.06, 0.0, 0.88, 1.0);
      bar(0.68, 0.48, 0.12, 0.1);
      break;
    case 'sofa':
      sh('round2SameRect', 0.08, 0.2, 0.84, 0.5, { rectRadius: w * 0.1 });
      sh('rect', 0.0, 0.44, 1.0, 0.42);
      break;
    case 'building':
      sh('rect', 0.14, 0.06, 0.72, 0.94);
      for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) bar(0.28 + c * 0.26, 0.24 + r * 0.2, 0.13, 0.11);
      break;
    default:
      sh('rect', 0.05, 0.05, 0.9, 0.9);
  }
}

/* ---------------------------------------------------------- slide builders */

/** 1 - Cover: full-bleed photo washed out by a pale gradient, logo centred. */
function slide01(s) {
  photo(s, 0, 0, SLIDE_W, SLIDE_H);
  gradientFade(s, 0, 0, SLIDE_W, SLIDE_H, [[0, WHITE, 97], [0.2, WHITE, 97], [1, 'F8F7F4', 75]], 'h', 40);
  houseGlyph(s, 2.861, 2.454, 1.453, 1.315, RED);
  houseGlyph(s, 3.118, 2.894, 0.938, 0.875, NAVY);
  wordmark(s, 1.611, 3.845, 3.952, 54);
  dots(s, 3.323, 4.926, 0.528);
}

/** 2 - Section divider. */
function slide02(s) {
  photo(s, 0, 0, SLIDE_W, SLIDE_H);
  rect(s, 0, 0, SLIDE_W, SLIDE_H, { color: WHITE, transparency: 10 });
  dots(s, 6.376, 3.022, 0.581);
  text(s, 'Smart Home Listing', {
    x: 2.542, y: 3.363, w: 8.25, h: 0.774, fontSize: 40, color: NAVY, charSpacing: 6,
    align: 'center', valign: 'middle',
  });
}

/** 3 - Text left / portrait photo right. */
function slide03(s) {
  photo(s, 6.156, 1.503, 6.224, 4.727);
  scrim(s, 6.156, 4.714, 6.224, 1.516, 75);
  dots(s, 0.979, 1.498);
  kicker(s, 0.979, 1.793, 4.285);
  heading(s, 0.979, 2.177, 4.285, 0.942, 'Buying a home is just like finding a new love');
  paragraph(s, 0.979, 3.22, 4.285, 2.553, FUSCE_LONG, { fontSize: 14 });
  header(s);
}

/** 4 - Portrait photo left / text right. */
function slide04(s) {
  photo(s, 0.979, 1.094, 4.406, 5.656);
  scrim(s, 0.979, 4.792, 4.406, 1.958, 75);
  dots(s, 6.198, 1.623);
  kicker(s, 6.198, 1.908, 6.16);
  heading(s, 6.198, 2.291, 6.16, 0.942, ['Why You Should Use an Agent', 'to Sell Your Home']);
  paragraph(s, 6.198, 3.387, 6.16, 1.818, LOREM_LONG);
  paragraph(s, 6.198, 5.363, 6.16, 0.606, FUSCE_SHORT, { bold: true, color: RED });
  header(s);
}

/** 5 - Three property categories. */
function slide05(s) {
  const cols = [
    [1.152, 1.056, 'Townhomes & Condos'],
    [5.0, 4.904, 'Residential'],
    [8.848, 8.752, 'New Construction '],
  ];
  cols.forEach(([px, tx, label]) => {
    photo(s, px, 1.8, 3.333, 4.394);
    gradientFade(s, px, 5.125, 3.333, 1.069, [[0.1, RED, 0], [1, NAVY, 85]], 'v', 10);
    text(s, label, {
      x: tx, y: 5.766, w: 3.526, h: 0.329, fontFace: HEAD, fontSize: 14, color: WHITE,
      charSpacing: 2, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2,
    });
    text(s, FUSCE_SHORT, {
      x: tx, y: 6.394, w: 3.526, h: 0.496, fontSize: 10, align: 'center', lineSpacingMultiple: 1.25,
    });
  });
  heading(s, 0.979, 0.914, 11.375, 0.539, 'You\u2019re Buying a Home!', { fontSize: 32, align: 'center' });
  header(s);
}

/** 6 - Full-bleed photo with a navy information card. */
function slide06(s) {
  photo(s, 0, 0, SLIDE_W, SLIDE_H);
  gradientFade(s, 0, 0, SLIDE_W, SLIDE_H, [[0, RED, 0], [1, NAVY, 65]], 'v', 14);
  rect(s, 0.667, 1.0, 4.941, 5.847, { color: NAVY, transparency: 15 });
  const W = 3.965, X = 1.178;
  kicker(s, X, 1.428, W, 'Title texts', { color: WHITE, fontSize: 12, h: 0.242 });
  heading(s, X, 1.749, W, 0.404, 'Search For Homes:', { fontSize: 24, color: WHITE });
  paragraph(s, X, 2.196, W, 0.909, LOREM_SHORT, { color: WHITE });
  heading(s, X, 3.544, W, 0.269, 'Purchase a Home with a REALTOR', { fontSize: 16, color: WHITE });
  paragraph(s, X, 3.842, W, 0.983, LOREM_PLAIN, { fontSize: 10, color: WHITE });
  heading(s, X, 5.119, W, 0.269, 'Connect to a Professional', { fontSize: 16, color: WHITE });
  paragraph(s, X, 5.418, W, 0.983, LOREM_PLAIN, { fontSize: 10, color: WHITE });
  header(s);
}

/** 7 - Photo bleeding off the left edge + three sign-post features. */
function slide07(s) {
  photo(s, 0, 0.943, 4.443, 5.904);
  gradientFade(s, 0, 4.597, 4.443, 2.251, [[0.1, RED, 0], [1, NAVY, 70]], 'v', 10);
  dots(s, 5.178, 1.256);
  kicker(s, 5.178, 1.52, 7.207, 'Title texts');
  heading(s, 5.178, 1.904, 7.207, 0.942, 'Real Estate Agents Possess Overall Market Knowledge');
  paragraph(s, 5.178, 2.947, 7.207, 1.515, LOREM_LONG);
  [5.178, 7.701, 10.224].forEach((x, i) => {
    icon(s, 'signpost', [5.198, 7.731, 10.265][i], 4.952, 0.312, 0.314, NAVY);
    text(s, 'Subtitle texts', {
      x, y: 5.312, w: 2.161, h: 0.305, fontSize: 12, bold: true, color: RED,
      charSpacing: 4, lineSpacingMultiple: 1.2, margin: 2.8,
    });
    paragraph(s, x, 5.602, 2.161, 1.089, LOREM_SHORT, { fontSize: 10, margin: 2.8 });
  });
  header(s);
}

/** 8 - Text left, photo mosaic right. */
function slide08(s) {
  photo(s, 5.802, 0.943, 4.198, 5.904);
  photo(s, 10.142, 0.943, 3.192, 2.886);
  photo(s, 10.142, 3.961, 3.192, 2.886);
  dots(s, 0.979, 1.498);
  kicker(s, 0.979, 1.793, 4.285);
  heading(s, 0.979, 2.177, 4.285, 0.942, 'Understanding the Home Buying Process');
  paragraph(s, 0.979, 3.228, 4.285, 2.694, LOREM_LONG);
  header(s);
  [[5.802, 5.646, 4.198], [10.142, 5.646, 3.192], [10.142, 2.628, 3.192]].forEach(([x, y, w]) => {
    gradientFade(s, x, y, w, 1.201, [[0.1, RED, 0], [1, NAVY, 52]], 'v', 10);
  });
}

/** 9 - Wide photo above three outlined feature cards. */
function slide09(s) {
  photo(s, 1.005, 1.191, 6.663, 3.189);
  dots(s, 8.135, 1.488);
  kicker(s, 8.135, 1.736, 4.193);
  heading(s, 8.135, 2.119, 4.193, 0.942, ['Purchase a Home', 'with a REALTOR']);
  paragraph(s, 8.135, 3.215, 4.193, 0.909, LOREM_SHORT);
  header(s);
  const cards = [
    [1.005, 1.239, 2.601, 'people'],
    [4.834, 5.068, 6.443, 'houseLine'],
    [8.663, 8.897, 10.338, 'list'],
  ];
  cards.forEach(([cx, tx, ix, kind]) => {
    rect(s, cx, 4.686, 3.671, 2.007, { type: 'none' }, { line: { color: RULE, width: 0.25 } });
    icon(s, kind, ix, 5.046, kind === 'list' ? 0.315 : 0.46, 0.373, RED);
    paragraph(s, tx, 5.586, 3.203, 0.804, LOREM_SHORT, { fontSize: 10.5, align: 'center' });
  });
  scrim(s, 1.005, 1.565, 6.663, 2.816, 75);
}

/** 10 - Phone mock-up with three numbered feature rows. */
function slide10(s) {
  [2.016, 3.513, 5.009].forEach((y) => {
    s.addShape('roundRect', { x: 7.797, y, w: 0.875, h: 0.59, fill: { color: RED }, rectRadius: 0.295 });
  });
  phoneMock(s, 5.194, 0.995, 2.944, 5.869, [5.388, 1.168, 2.558, 5.521]);
  [2.185, 3.682, 5.178].forEach((y) => icon(s, 'houseSolid', 8.219, y, 0.27, 0.252, BG));
  dots(s, 0.979, 1.686);
  kicker(s, 0.979, 1.981, 3.674, 'Title texts', { bold: true, h: 0.263 });
  heading(s, 0.979, 2.344, 3.674, 1.414, ['Full Loan Commitment &', 'Are Clear To Close!']);
  paragraph(s, 0.979, 3.859, 3.674, 2.121, FUSCE_LONG);
  header(s);
  [2.013, 3.512, 5.009].forEach((y) => {
    text(s, 'Subtitle texts', {
      x: 8.931, y, w: 3.351, h: 0.362, fontFace: HEAD, fontSize: 14, color: NAVY,
      charSpacing: 2, lineSpacingMultiple: 1.2, margin: 2.8,
    });
    paragraph(s, 8.931, y + 0.288, 3.351, 0.837, LOREM_SHORT, { fontSize: 10, margin: 2.8 });
  });
  gradientFade(s, 5.39, 4.33, 2.556, 2.36, [[0.1, RED, 0], [1, NAVY, 64]], 'v', 10);
}

/** 11 - Laptop mock-up with a red-to-navy footer band. */
function slide11(s) {
  gradientFade(s, 0, 5.593, SLIDE_W, 1.907, [[0, RED, 100], [0.67, NAVY, 100], [1, NAVY, 100]], 'h', 24, BG);
  laptopMock(s, 0.73, 1.5, 7.64, 4.41);
  photo(s, 1.674, 1.823, 5.751, 3.597);
  gradientFade(s, 1.674, 4.14, 5.751, 1.288, [[0.1, RED, 0], [1, NAVY, 65]], 'v', 10);
  dots(s, 8.42, 1.888);
  kicker(s, 8.42, 2.144, 3.977);
  heading(s, 8.42, 2.524, 3.977, 0.942, 'Get Your Home Market Ready!!!');
  paragraph(s, 8.42, 3.616, 3.977, 0.909, LOREM_SHORT);
  header(s);
  text(s, [
    { text: 'Get More Money When You Sell:', options: { fontSize: 18, breakLine: true } },
    { text: '\u201cHome Staging Tips that Will Make Your Home More Profitable\u201d', options: { fontSize: 16, bold: true, italic: true } },
  ], { x: 1.286, y: 6.258, w: 10.762, h: 0.664, color: WHITE, charSpacing: 4, align: 'center', valign: 'middle' });
}

/** 12 - Three team members. */
function slide12(s) {
  const people = [
    [0.979, 'Marissa Valley', 'BUYER SPECIALIST', 2.609],
    [4.859, 'Jhonny Beuckering', 'TEAM LEADER', 6.489],
    [8.74, 'Greg Thompson', 'LISTING SPECIALIST', 10.369],
  ];
  people.forEach(([x, name, role, dx]) => {
    photo(s, x, 2.127, 3.615, 2.0);
    gradientFade(s, x, 3.348, 3.615, 0.779, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 8);
    text(s, name, { x, y: 4.271, w: 3.615, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true, color: NAVY, align: 'center', margin: 3.6 });
    text(s, role, { x, y: 4.6, w: 3.615, h: 0.3, fontSize: 12, align: 'center', margin: 3.6 });
    dots(s, dx, 4.988);
    paragraph(s, x, 5.181, 3.615, 1.59, FUSCE_LONG, { fontSize: 10, margin: 3.6 });
  });
  heading(s, 0.979, 0.99, 11.375, 0.539, 'Our Boards', { fontSize: 32, align: 'center' });
  text(s, 'We have the most productive, service\u2013oriented real estate team in the state line area!', {
    x: 0.979, y: 1.505, w: 11.375, h: 0.34, fontSize: 14, align: 'center', margin: 3.6,
  });
  header(s);
}

/** 13 - 3x3 team directory. */
function slide13(s) {
  const roles = ['TEAM LEADER', 'Co LEADER', 'SUPPORT'];
  [1.044, 4.927, 8.811].forEach((px) => {
    [1.799, 3.596, 5.394].forEach((py, r) => {
      photo(s, px, py, 1.29, 1.316);
      const tx = px + 1.348;
      text(s, 'Team Name', { x: tx, y: py + 0.008, w: 2.026, h: 0.3, fontFace: HEAD, fontSize: 12, bold: true, color: NAVY, margin: 3.6 });
      text(s, roles[r], { x: tx, y: py + 0.289, w: 2.026, h: 0.271, fontSize: 10, color: RED, margin: 3.6 });
      paragraph(s, tx, py + 0.514, 2.026, 0.712, FUSCE_SHORT, { fontSize: 8, margin: 3.6 });
    });
  });
  text(s, [
    { text: 'Get To Know ', options: { color: NAVY } },
    { text: 'Our Teams', options: { color: RED } },
  ], { x: 0.979, y: 0.864, w: 11.375, h: 0.539, fontFace: HEAD, fontSize: 32, align: 'center' });
  header(s);
}

/** 14 - Listing grid: photo + price bar + address bar, six times. */
function slide14(s) {
  [1.773, 4.418].forEach((row) => {
    [0.979, 4.878, 8.776].forEach((col) => {
      photo(s, col, row, 3.578, 1.977);
      rect(s, col, row + 1.65, 3.578, 0.336, { color: RED, transparency: 25 });
      rect(s, col, row + 1.977, 3.578, 0.421, NAVY);
      text(s, '$ 215,000', { x: col + 0.12, y: row + 1.66, w: 1.484, h: 0.296, fontFace: HEAD, fontSize: 12, color: WHITE, valign: 'bottom', margin: 3.6 });
      text(s, '4 Beds    3 Baths', { x: col + 1.92, y: row + 1.676, w: 1.531, h: 0.271, fontFace: HEAD, fontSize: 10, color: WHITE, align: 'right', valign: 'bottom', margin: 3.6 });
      text(s, ADDRESS, { x: col + 0.12, y: row + 2.045, w: 3.334, h: 0.28, fontSize: 10.5, color: WHITE, align: 'center', margin: 3.6 });
    });
  });
  heading(s, 0.979, 0.864, 11.375, 0.539, 'Home Listings', { fontSize: 32, align: 'center' });
  header(s);
}

/** 15 - Hero photo over a property detail strip. */
function slide15(s) {
  photo(s, 0, 0, SLIDE_W, 4.26);
  gradientFade(s, 0, 2.32, SLIDE_W, 1.94, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 10);
  dots(s, 0.972, 4.664);
  heading(s, 0.972, 4.948, 7.284, 0.34, 'About This Home', { fontSize: 20, valign: 'bottom' });
  paragraph(s, 0.972, 5.31, 6.641, 1.213, LOREM_MED);
  icon(s, 'money', 7.98, 4.87, 0.42, 0.4, 'C00000');
  rect(s, 8.497, 4.871, 1.551, 0.4, RED);
  rect(s, 10.048, 4.871, 2.334, 0.4, NAVY);
  text(s, '$ 215,000', { x: 8.559, y: 4.884, w: 1.412, h: 0.371, fontFace: HEAD, fontSize: 16, color: WHITE, align: 'right', valign: 'middle', margin: 3.6 });
  text(s, 'Estimate your mortgage', { x: 10.141, y: 4.928, w: 2.156, h: 0.28, fontSize: 12, italic: true, color: WHITE, align: 'center', valign: 'bottom', margin: 3.6 });
  rect(s, 8.497, 5.263, 3.885, 0.4, CREAM);
  text(s, ADDRESS, { x: 8.497, y: 5.263, w: 3.885, h: 0.4, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5, margin: 3.6 });
  const facts = [[8.497, 8.9, 0.35, 'houseLine', '4 Beds'], [9.857, 10.35, 0.19, 'drop', '3 Baths'], [11.216, 11.66, 0.29, 'expand', '3,974 sqft']];
  facts.forEach(([tx, ix, iw, kind, label]) => {
    icon(s, kind, ix, 5.84, iw, 0.28, RED);
    text(s, label, { x: tx, y: 6.21, w: 1.166, h: 0.3, fontFace: HEAD, fontSize: 11, color: NAVY, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

/** 16 - 3x3 benefit matrix separated by dashed rules. */
function slide16(s) {
  const items = [
    ['Best collection', NAVY], ['Easy to buy', RED], ['Good quality', NAVY],
    ['Easy to buy', RED], ['Affordable price', NAVY], ['Convenient', RED],
    ['Convenient', NAVY], ['Best collection', RED], ['Easy to buy', NAVY],
  ];
  const colX = [1.21, 5.1, 9.14];
  const rowY = [1.45, 3.34, 5.33];
  items.forEach((it, i) => {
    const x = colX[i % 3], y = rowY[Math.floor(i / 3)];
    const [label, color] = it;
    houseGlyph(s, x, y, 0.49, 0.45, color);
    text(s, label, { x: x + 0.69, y: y + 0.2, w: 2.47, h: 0.3, fontFace: HEAD, fontSize: 18, color });
    paragraph(s, x + 0.04, y + 0.6, 3.12, 0.67, FUSCE_SHORT.trim(), { fontSize: 11, lineSpacingMultiple: 1.2 });
  });
  [4.64, 8.69].forEach((x) => s.addShape('line', { x, y: 1.65, w: 0, h: 4.94, line: { color: DASH, dashType: 'dash', width: 0.75 } }));
  [3.03, 5.0].forEach((y) => s.addShape('line', { x: 1.03, y, w: 11.27, h: 0, line: { color: DASH, dashType: 'dash', width: 0.75 } }));
  header(s);
}

/** 17 - "Mountain" comparable-sales chart made of translucent peaks. */
function slide17(s) {
  const peaks = [
    [10.169, 2.926, 2.248, '3F3F3F'], [8.971, 3.675, 1.499, 'C00000'],
    [7.774, 2.688, 2.485, 'B2B2B2'], [6.577, 4.243, 0.93, '17365D'],
    [5.38, 3.687, 1.486, '3F3F3F'], [4.182, 3.425, 1.748, 'C00000'],
    [2.985, 4.674, 0.5, 'B2B2B2'], [1.788, 2.891, 2.282, '17365D'],
  ];
  peaks.forEach(([x, y, h, color]) => {
    const w = 2.186;
    s.addShape('custGeom', {
      x, y, w, h,
      points: [
        { x: 0, y: h },
        { curve: { type: 'cubic', x1: 0.1667 * w, y1: 0.7772 * h, x2: 0.4169 * w, y2: 0.0046 * h }, x: 0.5052 * w, y: 0 },
        { curve: { type: 'cubic', x1: 0.5936 * w, y1: 0, x2: 0.8333 * w, y2: 0.7772 * h }, x: w, y: h },
        { close: true },
      ],
      fill: { color, transparency: 20 },
    });
  });
  comparableSalesFrame(s);
}

/** 18 - Same chart frame, drawn as stacked area silhouettes. */
function slide18(s) {
  const bands = [
    [2.996, 2.178, '595959', 5, [[0.0005, 0.5735], [0.3056, 0.5413], [0.5429, 0], [0.7827, 0.5185], [1, 0.6432]]],
    [3.635, 1.538, 'C00000', 5, [[0.001, 0], [0.3096, 0.5824], [0.5757, 0.4605], [0.7837, 0.088], [1, 0.2641]]],
    [4.042, 1.132, NAVY, 5, [[0.0007, 0.0178], [0.281, 0.4202], [0.555, 0], [0.763, 0.1565], [1, 0.5668]]],
    [4.704, 0.469, 'D9D9D9', 5, [[0.0001, 0.1134], [0.2954, 0.6112], [0.6835, 0], [1, 0.5567]]],
  ];
  bands.forEach(([y, h, color, transparency, pts]) => {
    const w = 10.566;
    s.addShape('custGeom', {
      x: 1.788, y, w, h,
      points: pts.map(([px, py], i) => (i === 0 ? { x: px * w, y: py * h, moveTo: true } : { x: px * w, y: py * h }))
        .concat([{ x: w, y: h }, { x: 0, y: h }, { close: true }]),
      fill: { color, transparency },
    });
  });
  comparableSalesFrame(s);
}

/** Shared axis labels + copy for slides 17 and 18. */
function comparableSalesFrame(s) {
  ['100%', '80%', '60%', '40%', '20%', '0%'].forEach((label, i) => {
    text(s, label, { x: 0.979, y: 2.543 + i * 0.5, w: 0.646, h: 0.24, fontSize: 10, color: '7F7F7F', align: 'center', valign: 'middle' });
  });
  dots(s, 0.979, 0.908);
  kicker(s, 0.979, 1.144, 6.479, 'Title texts');
  heading(s, 0.979, 1.489, 6.479, 0.539, 'Examine Comparable Sales', { fontFace: HEAD, fontSize: 32 });
  text(s, 'Subtitle Texts', { x: 0.972, y: 5.531, w: 11.375, h: 0.34, fontSize: 14, bold: true, color: NAVY, valign: 'bottom', margin: 3.6 });
  paragraph(s, 0.979, 5.864, 11.375, 1.007, LOREM_MED, { margin: 3.6 });
  header(s);
}

/** 19 - Five-step listing process. */
function slide19(s) {
  const steps = [
    [0.91, 1.39, 2.83, 2.42, 'speech', 0.41, 'Meet & greet'],
    [4.95, 5.43, 2.83, 6.52, 'doc', 0.29, 'Pre-list appointment'],
    [8.99, 9.47, 2.83, 10.55, 'calc', 0.31, 'Active listing'],
    [2.9, 3.38, 5.45, 4.41, 'pen', 0.41, 'Accepted offer'],
    [6.96, 7.48, 5.45, 8.52, 'check', 0.4, 'Closing process'],
  ];
  steps.forEach(([cardX, textX, textY, iconX, kind, iconW, label]) => {
    rect(s, cardX, textY + 0.24, 3.44, 1.1, { type: 'none' }, { line: { color: RULE, width: 0.75 } });
    icon(s, kind, iconX, textY - 0.88, iconW, 0.4, RED);
    text(s, label, { x: textX, y: textY - 0.38, w: 2.48, h: 0.3, fontFace: HEAD, fontSize: 18, color: NAVY, align: 'center' });
    paragraph(s, textX, textY, 2.48, 1.11, LOREM_SHORT, { fontSize: 11, align: 'center' });
  });
  [[4.32, 3.06], [8.36, 3.06], [6.31, 5.67]].forEach(([x, y]) => {
    s.addShape('triangle', { x, y, w: 0.43, h: 0.39, fill: { color: RED }, rotate: 90 });
  });
  heading(s, 0.979, 0.864, 11.375, 0.539, 'The Listing Process', { fontSize: 32, align: 'center' });
  header(s);
}

/** 20 - Pull quote. */
function slide20(s) {
  rect(s, 0, 0, SLIDE_W, SLIDE_H, { color: BG, transparency: 20 });
  dots(s, 6.45, 2.792, 0.43);
  text(s, [
    { text: '\u201cHome staging is no longer optional', options: { color: NAVY, breakLine: true } },
    { text: 'In the real estate market, ', options: { color: NAVY } },
    { text: 'it is a must\u201d', options: { color: RED } },
  ], { x: 1.556, y: 3.049, w: 10.222, h: 1.207, fontFace: HEAD, fontSize: 36, align: 'center' });
  text(s, '- Barbara Corcoran', { x: 2.174, y: 4.402, w: 8.976, h: 0.27, fontFace: HEAD, fontSize: 16, align: 'center' });
  header(s);
}

/** 21 - Two photos with a caption block. */
function slide21(s) {
  photo(s, 0, 1.306, 6.632, 4.889);
  photo(s, 6.722, 1.306, 5.653, 2.444);
  gradientFade(s, 6.722, 2.65, 5.653, 1.1, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 8);
  gradientFade(s, 0, 4.985, 6.632, 1.21, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 8);
  dots(s, 7.071, 4.213, 0.43);
  heading(s, 7.03, 4.448, 5.331, 0.4, 'Get Your Home Market Ready!!!', { fontSize: 24, valign: 'bottom' });
  paragraph(s, 7.03, 4.9, 5.331, 1.213, LOREM_PLAIN.replace(/^Lorem/, '"Lorem').trimEnd());
  header(s);
}

/** 22 - Photo collage. */
function slide22(s) {
  photo(s, 0.78, 1.139, 8.232, 5.667);
  photo(s, 9.084, 1.139, 3.498, 2.795);
  photo(s, 9.084, 4.01, 3.498, 2.795);
  [[0.78, 5.573, 8.232], [9.084, 5.573, 3.498], [9.084, 2.693, 3.498]].forEach(([x, y, w]) => {
    gradientFade(s, x, y, w, 1.233, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 8);
  });
  header(s);
}

/** 23 - Comparable-homes diagram. */
function slide23(s) {
  heading(s, 1.806, 1.271, 9.722, 0.47, 'Real Estate Agents Possess Overall Market Knowledge', { align: 'center' });
  s.addShape('ellipse', { x: 1.06, y: 3.05, w: 6.21, h: 2.66, fill: { type: 'none' }, line: { color: 'A6A6A6', width: 1, dashType: 'dash' } });
  s.addShape('ellipse', { x: 1.32, y: 3.17, w: 5.66, h: 2.26, fill: { color: 'EDEDED' } });
  s.addShape('line', { x: 1.68, y: 3.36, w: 4.46, h: 1.82, line: { color: 'A6A6A6', width: 1, dashType: 'dash' }, flipH: true });
  s.addShape('ellipse', { x: 2.46, y: 3.6, w: 2.72, h: 1.11, fill: { color: 'D9D9D9' } });
  photo(s, 2.99, 2.62, 2.19, 1.91);
  [['Age', 5.58, 3.0, 0.55], ['Size', 5.58, 3.34, 0.55], ['Location', 5.58, 3.69, 0.88], ['Improvements', 5.58, 4.04, 1.32]].forEach(([label, x, y, w]) => {
    rect(s, x, y, w, 0.32, NAVY);
    text(s, label, { x, y, w, h: 0.32, fontSize: 11, color: WHITE, align: 'center', valign: 'middle', margin: 5.7 });
  });
  [['1 Mile Radius', 3.52, 5.55, 1.31], ['Similar Homes', 1.78, 3.42, 1.36]].forEach(([label, x, y, w]) => {
    rect(s, x, y, w, 0.34, RED);
    text(s, label, { x, y, w, h: 0.34, fontSize: 11, color: WHITE, align: 'center', valign: 'middle', margin: 3.6 });
  });
  dots(s, 7.99, 2.64);
  kicker(s, 7.96, 2.88, 4.44, 'Title texts', { fontFace: HEAD, fontSize: 16, h: 0.3 });
  paragraph(s, 7.96, 3.29, 4.44, 2.42, LOREM_LONG);
  icon(s, 'money', 11.62, 2.65, 0.5, 0.48, 'C00000');
  header(s);
}

/** 24 - Three agent types, photo on the right. */
function slide24(s) {
  photo(s, 9.907, 1.063, 3.426, 5.733);
  gradientFade(s, 9.907, 4.24, 3.426, 2.556, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 10);
  const rows = [
    [1.291, 'Seller\u2019s Agent', 'personCheck', 1.45, 0.38],
    [3.163, 'Buyer\u2019s Agent', 'personSearch', 3.3, 0.38],
    [5.036, 'Dual Agent ', 'people', 5.19, 0.33],
  ];
  rows.forEach(([y, title, kind, iy, ih]) => {
    icon(s, kind, 0.979, iy, 0.38, ih, RED);
    heading(s, 1.653, y, 7.9, 0.34, title, { fontSize: 20 });
    kicker(s, 1.653, y + 0.362, 7.9, 'Subtitle texts', { bold: true });
    paragraph(s, 1.663, y + 0.669, 7.9, 0.909, FUSCE_LONG);
  });
  header(s);
}

/** 25 - Text column plus a photo strip. */
function slide25(s) {
  photo(s, 8.833, 1.287, 3.519, 5.509);
  [1.287, 3.137, 4.986].forEach((y) => photo(s, 5.016, y, 3.782, 1.81));
  dots(s, 0.979, 1.291);
  kicker(s, 0.979, 1.524, 3.478);
  heading(s, 0.979, 1.898, 3.478, 1.414, 'Real Estate Agents Possess Overall Market Knowledge');
  paragraph(s, 0.979, 3.404, 3.478, 2.42, LOREM_MED);
  gradientFade(s, 8.833, 4.987, 3.519, 1.809, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 10);
  [2.132, 3.982, 5.831].forEach((y) => gradientFade(s, 5.016, y, 3.782, 0.965, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 8));
  header(s);
}

/** 26 - Text with two small features, photo bleeding off the right. */
function slide26(s) {
  photo(s, 7.286, 0, 5.127, 7.5);
  gradientFade(s, 7.286, 4.51, 5.127, 2.99, [[0.1, RED, 0], [1, NAVY, 60]], 'v', 12);
  dots(s, 0.979, 1.291);
  kicker(s, 0.979, 1.524, 5.466);
  heading(s, 0.979, 1.898, 5.466, 0.942, 'Real Estate Agents Possess Overall Market Knowledge');
  paragraph(s, 0.979, 2.987, 5.466, 1.515, LOREM_MED);
  [[1.05, 0.979, 'door', 0.25, 0.31, 4.87], [3.895, 3.859, 'sofa', 0.33, 0.21, 4.98]].forEach(([ix, tx, kind, iw, ih, iy]) => {
    icon(s, kind, ix, iy, iw, ih, RED);
    text(s, 'Subtitle Texts', { x: tx, y: 5.324, w: 2.577, h: 0.2, fontSize: 12, bold: true, color: NAVY, valign: 'bottom' });
    paragraph(s, tx, 5.573, 2.577, 0.933, LOREM_SHORT, { fontSize: 11, lineSpacingMultiple: 1.25 });
  });
}

/** 27 - Two phone mock-ups. */
function slide27(s) {
  phoneMock(s, 1.64, 0.99, 2.94, 5.869, [1.802, 1.168, 2.593, 5.521]);
  gradientFade(s, 1.84, 4.94, 2.556, 1.75, [[0.1, RED, 0], [1, NAVY, 64]], 'v', 8);
  phoneMock(s, 3.77, 2.14, 2.44, 4.86, [3.926, 2.28, 2.116, 4.568]);
  gradientFade(s, 3.93, 5.54, 2.116, 1.31, [[0.1, RED, 0], [1, NAVY, 64]], 'v', 8);
  dots(s, 6.93, 1.424);
  kicker(s, 6.93, 1.717, 5.462);
  heading(s, 6.93, 2.098, 5.462, 0.942, ['Full Loan Commitment', '& Are Clear To Close!']);
  paragraph(s, 6.93, 3.152, 5.462, 1.515, FUSCE_LONG);
  icon(s, 'building', 6.99, 4.99, 0.31, 0.31, RED);
  text(s, 'Subtitle texts', { x: 6.93, y: 5.39, w: 5.462, h: 0.32, fontSize: 12, bold: true, color: RED, charSpacing: 4, lineSpacingMultiple: 1.2, margin: 2.8 });
  paragraph(s, 6.93, 5.712, 5.462, 0.58, LOREM_SHORT, { fontSize: 10, margin: 2.8 });
  header(s);
}

/** 28 - Laptop + tablet mock-ups on a tinted panel. */
function slide28(s) {
  photo(s, 5.462, 2.174, 6.969, 4.359);
  rect(s, 5.462, 2.174, 6.969, 4.366, { color: 'EFC3C2', transparency: 50 });
  laptopMock(s, 0.73, 2.139, 7.64, 4.41);
  photo(s, 1.674, 2.467, 5.751, 3.597);
  gradientFade(s, 1.674, 5.07, 5.751, 1.0, [[0.1, RED, 0], [1, NAVY, 65]], 'v', 8);
  heading(s, 1.876, 1.008, 9.581, 0.47, 'Get Your Home Market Ready!!!', { align: 'center' });
  paragraph(s, 1.876, 1.542, 9.581, 0.28, LOREM_SHORT, { fontSize: 11, align: 'center' });
  laptopMock(s, 5.13, 3.58, 5.35, 3.09);
  photo(s, 5.799, 3.802, 4.022, 2.516);
  gradientFade(s, 5.799, 5.532, 4.022, 0.786, [[0.1, RED, 0], [1, NAVY, 65]], 'v', 8);
  header(s);
}

/** 29 - Native clustered-bar chart. */
function slide29(s) {
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 1.0, y: 2.081, w: 6.049, h: 3.543,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 50,
    chartColors: ['17365D', 'C00000', 'A5A5A5'],
    showLegend: false, showTitle: false, showValue: false,
    catAxisHidden: true, catGridLine: { style: 'none' },
    valGridLine: { color: 'F2F2F2', size: 0.75 },
    valAxisLineShow: false, valAxisLabelColor: INK, valAxisLabelFontFace: BODY, valAxisLabelFontSize: 10.5,
  });
  dots(s, 7.82, 2.008);
  kicker(s, 7.82, 2.297, 4.535);
  heading(s, 7.82, 2.685, 4.535, 0.942, ['Full Loan Commitment', '& Are Clear To Close!']);
  paragraph(s, 7.82, 3.732, 4.535, 1.818, LOREM_MED);
  header(s);
}

/** 30 - Reference table. */
function slide30(s) {
  const rows = [
    ['Aliquam sed ipsum maximus ', 'Passed'],
    ['Maecenas consequat elementum ', 'Passed'],
    ['Vivamus interdum porta ', 'Passed'],
    ['Quisque ut nibh ante', 'Passed'],
    ['Orci varius natoque penatibus ', 'Passed'],
    ['Aenean rhoncus tempus erat', 'Passed'],
  ];
  const head = [
    { text: 'REFERENCES', options: { fill: { color: NAVY }, color: WHITE, align: 'center' } },
    { text: 'PRICE', options: { fill: { color: RED }, color: WHITE, align: 'center' } },
  ];
  const body = rows.map(([label, price], i) => {
    const shade = { color: i % 2 === 0 ? 'F2F2F2' : 'F8F7F7' };
    return [
      { text: label, options: { fill: shade, bold: true, color: INK_DARK } },
      { text: price, options: { fill: shade, color: INK_DARK, align: 'center' } },
    ];
  });
  s.addTable([head].concat(body), {
    x: 6.679, y: 2.148, w: 5.713, colW: [3.93, 1.783], rowH: 0.529,
    fontFace: HEAD, fontSize: 11, valign: 'middle', margin: [2, 8, 2, 8],
    border: { type: 'none' },
  });
  dots(s, 1.008, 2.008);
  kicker(s, 1.008, 2.297, 5.462);
  heading(s, 1.008, 2.685, 5.462, 0.942, ['Full Loan Commitment', '& Are Clear To Close!']);
  paragraph(s, 1.008, 3.732, 5.103, 2.121, LOREM_LONG);
  header(s);
}

/* ------------------------------------------------------------------- build */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'MAKLAAR', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'MAKLAAR';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Smart Home Listing';

  BUILDERS.forEach((builder) => {
    const slide = pres.addSlide();
    slide.background = { color: BG };
    builder(slide);
  });

  return pres.writeFile({ fileName: path.join(__dirname, '13e1080b-3796-4f11-aca5-fd9770901510_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
