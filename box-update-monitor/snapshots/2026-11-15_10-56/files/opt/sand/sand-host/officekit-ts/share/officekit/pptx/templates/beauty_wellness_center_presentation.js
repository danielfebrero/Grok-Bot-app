/**
 * Beaute. — Beauty Care Presentation Template (30 slides, 16:9 / 13.333in x 7.5in)
 *
 * Standalone re-creation of the reference deck with pptxgenjs only.
 * Run:  node 084033d8-f1da-4bc0-8cea-c2f1e00991b2_grok_final.js
 *
 * Picture placeholders in the original carry a decorative pattern fill and no
 * bitmap; here they are drawn as flat tinted rectangles (see `photo`).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
// Palette / typography (from the template theme "Custom 92" / "Custom 3")
// ---------------------------------------------------------------------------
const PURPLE = '9246EC';     // accent1
const PURPLE_LT = 'BE90F4';  // accent1, lighter
const FRAME = 'D3B5F7';      // accent1 @ 40% luminance — the thin page frames
const PHOTO = 'F7F1FD';      // stand-in tint for the picture placeholders
const INK = '2B2B2B';        // tx1
const GREY_M = '404040';     // tx1 @ 75%
const GREY_D = '606060';     // tx1 @ 62%
const GREY = '959595';       // tx1 @ 50%
const RULE = 'D5D5D5';
const CARD_EDGE = 'AAAAAA';
const RED = '890710';
const WHITE = 'FFFFFF';

const HEAD = 'Raleway';      // +mj-lt
const BODY = 'Lato';         // +mn-lt
const SCRIPT = 'Satisfy';

// The two text styles the template reuses everywhere.
const COPY = { fontFace: BODY, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 };
const LEAD = { fontFace: HEAD, color: GREY_M, bold: true, lineSpacingMultiple: 1.5 };

// Coloured runs used inside the two-tone slide titles.
const P = (text, fontSize, breakLine) => ({ text, options: { fontSize, color: PURPLE, breakLine } });
const K = (text, fontSize, breakLine) => ({ text, options: { fontSize, color: INK, breakLine } });

// ---------------------------------------------------------------------------
// Recurring furniture
// ---------------------------------------------------------------------------

// Plain auto-sized text box — the template's default (Lato, top anchored, the
// box hugs its text exactly as PowerPoint's "resize shape to fit text" does).
function box(s, body, o) {
  s.addText(body, Object.assign({ fontFace: BODY, valign: 'top', fit: 'resize' }, o));
}

// Top-right + bottom-left corner frames, running header and page number.
function chrome(s, n) {
  s.addShape('rect', { x: 9.962, y: -0.108, w: 3.404, h: 1.1, line: { color: FRAME, transparency: 18 } });
  s.addShape('rect', { x: -0.078, y: 6.75, w: 1.756, h: 0.875, line: { color: FRAME, transparency: 18 } });
  s.addText('Beauty Care Presentation Template', {
    x: 10.579, y: 0.556, w: 2.181, h: 0.252,
    fontFace: BODY, fontSize: 9, color: GREY, align: 'right', wrap: false, valign: 'top', fit: 'resize',
  });
  s.addText(n + '.', {
    x: 0.638, y: 6.929, w: 0.323, h: 0.269, margin: [7.2, 7.2, 3.6, 3.6],
    fontFace: HEAD, fontSize: 10, bold: true, color: GREY, align: 'center', wrap: false, valign: 'top', fit: 'resize',
  });
}

// Hairline layout frame.
function frame(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, line: { color: FRAME, transparency: 18 } });
}

// Picture placeholder stand-in: the original shapes are empty `pic` placeholders
// carrying a faint purple pattern fill, so a flat tint stands in for them.
function photo(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: PHOTO });
}

// Small round "+" bullet mark. `invert` = white disc with a purple cross.
// rectRadius drives the `plus` adjust value (0.42805 -> thin arms).
function badge(s, x, y, invert) {
  s.addShape('ellipse', { x, y, w: 0.266, h: 0.266, fill: invert ? WHITE : PURPLE });
  s.addShape('plus', { x: x + 0.078, y: y + 0.078, w: 0.11, h: 0.11,
    rectRadius: 0.42805 * 0.11, fill: invert ? PURPLE : WHITE });
}

// Tiny purple play arrow; rot 270 points left, rot 90 points right.
function arrow(s, x, y, rotate) {
  s.addShape('triangle', { x, y, w: 0.146, h: 0.126, fill: PURPLE, rotate });
}

// Diagonally striped purple square (a decorative freeform in the original).
// Stripes run down-left at slope 1.585; both pitch and bar width are fractions
// of the square, so the motif scales with `d`.
function hatch(s, x, y, d) {
  const slope = 1.585, pitch = 0.1046 * d, bar = 0.0521 * d, run = d / slope;
  for (let a = 0.0578 * d - pitch; a < d + run; a += pitch) {
    const stripe = clipToSquare([
      { x: a, y: 0 }, { x: a + bar, y: 0 },
      { x: a + bar - run, y: d }, { x: a - run, y: d },
    ], d);
    if (stripe.length < 3) continue;
    s.addShape('custGeom', { x, y, w: d, h: d, fill: PURPLE_LT, points: [...stripe, { close: true }] });
  }
}

// Sutherland-Hodgman clip of a convex polygon against the square [0,d] x [0,d].
function clipToSquare(poly, d) {
  const planes = [p => p.x, p => d - p.x, p => p.y, p => d - p.y];
  return planes.reduce((pts, dist) => {
    const out = [];
    pts.forEach((cur, i) => {
      const prev = pts[(i + pts.length - 1) % pts.length];
      const dp = dist(prev), dc = dist(cur);
      if ((dp < 0) !== (dc < 0)) {
        const t = dp / (dp - dc);
        out.push({ x: prev.x + (cur.x - prev.x) * t, y: prev.y + (cur.y - prev.y) * t });
      }
      if (dc >= 0) out.push(cur);
    });
    return out;
  }, poly);
}

// Two-tone slide title with its handwritten tagline.
function titleBlock(s, t) {
  box(s, t.head, {
    x: t.x, y: t.y, w: t.w, h: t.h,
    fontFace: HEAD, bold: true, align: t.align || 'left',
  });
  box(s, t.tag, {
    x: t.x + t.dx, y: t.y + t.dy, w: t.tw, h: t.th,
    fontFace: SCRIPT, fontSize: t.tagSize, color: GREY_M, align: t.tagAlign || 'right',
  });
}

// Purple number disc + heading + paragraph.
function numItem(s, it) {
  s.addShape('ellipse', { x: it.x, y: it.y, w: 0.741, h: 0.741, fill: PURPLE });
  s.addText(it.num, {
    x: it.x, y: it.y, w: 0.741, h: 0.741,
    fontFace: BODY, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle',
  });
  box(s, it.title, { ...LEAD, fontSize: 14, x: it.x + 1.143, y: it.y, w: it.tw, h: 0.413 });
  box(s, it.body, { ...COPY, fontSize: 10, x: it.x + 1.148, y: it.y + 0.538, w: it.bw, h: 0.825 });
}
// ---------------------------------------------------------------------------
// Body copy that the template repeats verbatim on several slides
// ---------------------------------------------------------------------------
const LOREM_A = 'Lorem ipsum dolor sit amet, consecter adipiscing tortor. Donec elit lorem, finibusas imperdiet orcia inter lorem finibusa';
const LOREM_B = 'PLACEHOLDER';
const LOREM_C = 'Lorem ipsum dolor sit amet, consecter adipiscing tortor. Donec elit lorem, finibusas imperdiet orciace inter lorem finibusa';
const LOREM_D = 'Lorem ipsum dolor sit amet, consecter hana at adipiscing tortor. Donecas elit lorem, finibusas imperdiet orcia inter';
const LOREM_E = 'Lorem ipsum dolor sit amet at consectera  tortor finisa imperdiet';
const LOREM_F = 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magna dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orci ace, interdumas tortor. Donec elit lorem, finibusas imperdiet orci ac, interdu lorem, finibusa';
const LOREM_G = 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magna dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orci ace, interdumas tortor. Donec elit lorem, finibusas imperdiet orci ac, interdusan lorem, finibus magna dapibus ultriesi';

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function slide01(s, n) {
  frame(s, 7.656, 2.74, 5.01, 4.01);
  frame(s, 0.667, 0.75, 5.01, 4.01);
  photo(s, 1.678, 1.77, 9.989, 3.96);
  chrome(s, n);
  s.addShape('rect', { fill: WHITE, x: 4.667, y: 3.323, w: 4, h: 0.873 });
  box(s, [
    { text: 'Be', options: { fontFace: HEAD, fontSize: 54, color: INK, bold: true, align: 'center' } },
    { text: 'aute.', options: { fontFace: HEAD, fontSize: 54, color: PURPLE, bold: true, align: 'center' } },
  ], { x: 5.048, y: 3.255, w: 3.394, h: 1.01 });
  badge(s, 1.057, 1.135);
  box(s, 'Welcome To', { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, x: 1.63, y: 1.103, w: 2.609, h: 0.337 });
  arrow(s, 12.072, 6.165, 270);
}

function slide02(s, n) {
  frame(s, 5.677, 3.75, 6.99, 3);
  frame(s, 0.667, 0.75, 5.01, 4.01);
  photo(s, 1.682, 1.742, 3, 4.015);
  titleBlock(s, { x: 6.471, y: 1.848, w: 4.626, h: 1.447, head: [P('Our Beauty Care ', 40), K('Center.', 40)], tag: 'beauty & wellness center', tagSize: 17, dx: 2.061, dy: 1.009, tw: 2.48, th: 0.387 });
  box(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magna dapibus ultrices quis eu orci. Donec eu tellus comildo, imperdiet orci ac, interdum tortor. Donec elit lorem, finibus imperdiet orci ac, interdu lorem, finibusa', { ...COPY, fontSize: 10, x: 6.513, y: 4.22, w: 5.286, h: 1.078 });
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 6.513, y: 5.865, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: GREY_D, x: 7.996, y: 5.944, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: GREY_D, align: 'right', x: 10.038, y: 5.943, w: 1.01, h: 0.269 });
  arrow(s, 11.528, 6.017, 270);
  chrome(s, n);
  s.addShape('rect', { fill: PURPLE, x: 4.682, y: 1.742, w: 0.995, h: 3.018 });
  badge(s, 5.046, 2.081, true);
  box(s, ' The Story ', { fontSize: 10, color: WHITE, align: 'right', lineSpacingMultiple: 1.5, rotate: 270, x: 4.725, y: 4.017, w: 0.84, h: 0.32 });
}

function slide03(s, n) {
  frame(s, 0.673, 0.764, 6.994, 3);
  photo(s, 5.667, 1.75, 6.016, 3.021);
  titleBlock(s, { x: 1.691, y: 1.867, w: 4.626, h: 1.447, head: [P('Our Beauty Care ', 40), K('Center.', 40)], tag: 'beauty & wellness center', tagSize: 17, dx: 2.061, dy: 1.009, tw: 2.48, th: 0.387 });
  chrome(s, n);
  frame(s, 5.677, 5.242, 6.997, 1.508);
  box(s, LOREM_F, { ...COPY, fontSize: 10, x: 1.756, y: 4.245, w: 3.129, h: 1.582 });
  hatch(s, 11.293, 5.625, 0.742);
  badge(s, 1.057, 1.135);
  box(s, 'Lorem ipsum dolor sit amet, hain consectetu adipiscing elit. Maecenas at faucibusi asnnal. Phasellus vitae libero.', { ...COPY, fontSize: 9, x: 8.055, y: 5.6, w: 2.61, h: 0.753 });
  box(s, [
    { text: '25', options: { fontSize: 54, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 28, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 6.038, y: 5.477, w: 1.509, h: 1.01 });
  box(s, 'off', { fontSize: 14, color: PURPLE, align: 'justify', lineSpacingMultiple: 1.5, x: 7.068, y: 5.549, w: 0.446, h: 0.408 });
}

function slide04(s, n) {
  frame(s, 6.667, 3.472, 6, 3.281);
  photo(s, 4.653, 1.764, 4.028, 4);
  frame(s, 0.673, 0.764, 2.994, 3);
  chrome(s, n);
  titleBlock(s, { x: 7.146, y: 4.02, w: 4.626, h: 1.447, head: [P('Our Beauty Care ', 40), K('Center.', 40)], tag: 'beauty & wellness center', tagSize: 17, dx: 2.061, dy: 1.009, tw: 2.48, th: 0.387 });
  box(s, 'Lorem ipsum dolor sitas amet, consectetur nasadipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae hanis libero ut orci. Donec eu tellus comildo, interdum tortor. ', { ...COPY, fontSize: 10, x: 1.152, y: 1.777, w: 2.035, h: 1.582 });
  box(s, 'The Story ', { ...LEAD, fontSize: 14, x: 1.152, y: 1.169, w: 2.035, h: 0.413 });
  hatch(s, 9.22, 1.777, 0.742);
  badge(s, 1.243, 4.27);
  box(s, 'Our List :', { ...LEAD, fontSize: 12, rotate: 90, x: 0.912, y: 5.081, w: 1.001, h: 0.368 });
  box(s, [
    { text: '  Washing Hair', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '   ', options: { fontSize: 10.5, color: GREY, breakLine: true } },
    { text: '  Haircut ', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Pedicure Manicure', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Spa ', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 } } },
  ], { x: 1.979, y: 4.265, w: 1.779, h: 1.338 });
  arrow(s, 12.072, 6.165, 270);
}

function slide05(s, n) {
  frame(s, 0.673, 0.764, 7.007, 4.972);
  frame(s, 8.679, 4.714, 3.987, 2.04);
  photo(s, 8.683, 1.778, 2.982, 3.958);
  photo(s, 1.669, 3.785, 1.979, 1.951);
  chrome(s, n);
  titleBlock(s, { x: 1.587, y: 1.827, w: 4.626, h: 1.447, head: [P('Our Beauty Care ', 40), K('Center.', 40)], tag: 'beauty & wellness center', tagSize: 17, dx: 2.061, dy: 1.009, tw: 2.48, th: 0.387 });
  box(s, 'Lorem ipsum dolor sit amet, han consecter adipiscing elit. Maecenas at dasal faucibus asnnal. Phasellus vitaesan libero ut magna dapibus ultrices qieu orci. Doneci eu tellus comildo, imperdiet', { ...COPY, fontSize: 10, x: 4.291, y: 3.978, w: 2.741, h: 1.33 });
  s.addShape('rect', { fill: PURPLE, x: 7.675, y: 1.778, w: 1.006, h: 3.958 });
  badge(s, 8.047, 2.081, true);
  box(s, 'The Story About Us', { fontSize: 10, color: WHITE, align: 'center', lineSpacingMultiple: 1.5, rotate: 270, x: 7.379, y: 4.482, w: 1.534, h: 0.32 });
  box(s, 'ADDRES :', { ...LEAD, fontSize: 12, x: 4.27, y: 6.148, w: 1.23, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: GREY_D, x: 5.5, y: 6.226, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: GREY_D, align: 'right', x: 10.366, y: 6.225, w: 1.01, h: 0.269 });
  arrow(s, 12.069, 6.299, 270);
  box(s, 'NUMBER :', { ...LEAD, fontSize: 12, x: 9.139, y: 6.148, w: 1.01, h: 0.368 });
}

function slide06(s, n) {
  frame(s, 0.673, 0.764, 5.994, 3);
  photo(s, 1.678, 1.762, 4.989, 2.003);
  chrome(s, n);
  titleBlock(s, { x: 1.838, y: 4.5, w: 4.626, h: 1.447, head: [P('Our Beauty Care ', 40), K('Center.', 40)], tag: 'beauty & wellness center', tagSize: 17, dx: 2.061, dy: 1.009, tw: 2.48, th: 0.387 });
  frame(s, 7.671, 3.75, 4.995, 3.003);
  box(s, LOREM_F, { ...COPY, fontSize: 10, x: 8.505, y: 4.59, w: 3.329, h: 1.582 });
  badge(s, 7.723, 1.824);
  box(s, 'Our List :', { ...LEAD, fontSize: 12, rotate: 90, x: 7.391, y: 2.634, w: 1.001, h: 0.368 });
  box(s, [
    { text: '  Washing Hair', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '   ', options: { fontSize: 10.5, color: GREY, breakLine: true } },
    { text: '  Haircut ', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Pedicure Manicure', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Spa ', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 } } },
  ], { x: 8.458, y: 1.819, w: 1.779, h: 1.338 });
  box(s, [
    { text: '  Creambath', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '   ', options: { fontSize: 10.5, color: GREY, breakLine: true } },
    { text: '  Coloring', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Hair loss treatment', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: '  Scrub, Facial, Toning', options: { fontSize: 10.5, color: GREY, bullet: { characterCode: '2713', indent: 13.5 } } },
  ], { x: 10.478, y: 1.819, w: 1.779, h: 1.338 });
  arrow(s, 12.111, 6.299, 270);
}

function slide07(s, n) {
  chrome(s, n);
  frame(s, 6.667, 2.712, 6, 4.041);
  frame(s, 0.673, 0.764, 4.989, 1.948);
  titleBlock(s, { x: 1.228, y: 1.149, w: 3.898, h: 1.178, head: [P('Taking Care Your ', 32), K('Beauty.', 32)], tag: 'beauty care & wellness', tagSize: 16, dx: 1.797, dy: 0.775, tw: 2.074, th: 0.37 });
  numItem(s, { x: 7.46, y: 5.043, num: '04', title: 'Reflexology', tw: 3.463, body: LOREM_C, bw: 3.329 });
  numItem(s, { x: 7.46, y: 3.274, num: '02', title: 'Body Scrub and Spa', tw: 3.178, body: LOREM_C, bw: 3.329 });
  numItem(s, { x: 1.267, y: 5.043, num: '03', title: 'Pedicure Manicure', tw: 2.035, body: LOREM_C, bw: 3.329 });
  numItem(s, { x: 1.267, y: 3.274, num: '01', title: 'Skin & Hair Care', tw: 2.035, body: LOREM_C, bw: 3.329 });
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 6.683, y: 1.85, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: GREY_D, x: 8.166, y: 1.928, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: GREY_D, align: 'right', x: 10.208, y: 1.927, w: 1.01, h: 0.269 });
  arrow(s, 11.74, 2.001, 270);
}

function slide08(s, n) {
  chrome(s, n);
  numItem(s, { x: 7.406, y: 4.546, num: '02', title: 'Body Scrub and Spa', tw: 3.178, body: LOREM_C, bw: 3.329 });
  numItem(s, { x: 1.674, y: 4.546, num: '01', title: 'Skin & Hair Care', tw: 2.035, body: LOREM_A, bw: 3.224 });
  frame(s, 6.667, 3.75, 6, 3.003);
  frame(s, 0.673, 0.764, 5.994, 2.986);
  titleBlock(s, { x: 1.719, y: 1.852, w: 4.431, h: 1.313, head: [P('Taking Care Your ', 36), K('Beauty.', 36)], tag: 'beauty care & wellness', tagSize: 17, dx: 2.119, dy: 0.885, tw: 2.207, th: 0.387 });
  badge(s, 1.057, 1.135);
  box(s, 'Lorem ipsum dolor sit amet, consecter atil adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magnasa dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orcisa ace, interdumas tortor. Donec elit lorem, finibusas impendiora', { ...COPY, fontSize: 10, x: 7.406, y: 1.991, w: 4.477, h: 1.078 });
  arrow(s, 12.072, 6.165, 270);
}

function slide09(s, n) {
  chrome(s, n);
  numItem(s, { x: 1.484, y: 3.737, num: '02', title: 'Skin & Hair Care', tw: 2.035, body: LOREM_A, bw: 3.224 });
  numItem(s, { x: 1.484, y: 1.819, num: '01', title: 'Body Scrub and Spa', tw: 2.751, body: LOREM_A, bw: 3.224 });
  frame(s, 0.673, 0.764, 5.994, 4.998);
  titleBlock(s, { x: 7.54, y: 1.718, w: 4.431, h: 1.313, head: [P('Taking Care Your ', 36), K('Beauty.', 36)], tag: 'beauty care & wellness', tagSize: 17, dx: 2.119, dy: 0.885, tw: 2.207, th: 0.387 });
  box(s, 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magna dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orci ace, interdumas tortor. Donec elit lorem, finibusas imperdiet orci ac, interdu lorem, finibus magna dapibus ultriesi', { ...COPY, fontSize: 10, x: 7.54, y: 3.747, w: 4.309, h: 1.33 });
  frame(s, 6.667, 5.762, 6, 0.992);
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 7.568, y: 6.073, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: GREY_D, x: 8.798, y: 6.152, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: GREY_D, align: 'right', x: 10.839, y: 6.151, w: 1.01, h: 0.269 });
  arrow(s, 1.095, 1.15, 90);
  badge(s, 6.534, 6.151);
}

function slide10(s, n) {
  chrome(s, n);
  numItem(s, { x: 4.623, y: 3.22, num: '02', title: 'Skin & Hair Care', tw: 2.035, body: LOREM_A, bw: 3.224 });
  numItem(s, { x: 4.623, y: 1.376, num: '01', title: 'Body Scrub and Spa', tw: 2.751, body: LOREM_A, bw: 3.224 });
  numItem(s, { x: 4.623, y: 5.065, num: '03', title: 'Pedicure Manicure', tw: 2.035, body: LOREM_A, bw: 3.224 });
  frame(s, 9.962, 2.738, 2.704, 4.015);
  frame(s, 0.673, 0.764, 2.978, 3.998);
  titleBlock(s, { x: 1.273, y: 1.376, w: 1.842, h: 2.255, head: [P('Taking Care Your ', 32), K('Beauty.', 32)], tag: 'beauty care & wellness', tagSize: 14, dx: 0.01, dy: 2.627, tw: 1.842, th: 0.337, tagAlign: 'left' });
  box(s, [
    { text: 'Lorem ipsum dolor sitas amet, consecti adipiscing elit. Maecenas at faucibu libero ut magna dapibus imperdiet orcisa ani aneli ', options: { ...COPY, fontSize: 10, breakLine: true } },
    { text: 'Interdu tortor', options: { fontSize: 10, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 } },
  ], { x: 10.481, y: 4.207, w: 1.709, h: 1.582 });
  box(s, [
    { text: 'We Care ', options: { fontFace: HEAD, fontSize: 16, color: GREY_M, bold: true, breakLine: true } },
    { text: 'About You', options: { fontFace: HEAD, fontSize: 16, color: GREY_M, bold: true } },
  ], { x: 10.481, y: 3.266, w: 1.709, h: 0.64 });
  badge(s, 12.119, 6.196);
}

function slide11(s, n) {
  frame(s, 0.673, 0.764, 4.18, 2.986);
  chrome(s, n);
  numItem(s, { x: 1.674, y: 4.955, num: '01', title: 'Body Scrub and Spa', tw: 3.224, body: LOREM_A, bw: 3.224 });
  numItem(s, { x: 4.481, y: 3.356, num: '02', title: 'Skin & Hair Care', tw: 2.035, body: LOREM_A, bw: 3.224 });
  numItem(s, { x: 7.287, y: 1.757, num: '03', title: 'Pedicure Manicure', tw: 2.035, body: LOREM_A, bw: 3.224 });
  titleBlock(s, { x: 8.214, y: 5.272, w: 3.898, h: 1.178, head: [P('Taking Care Your ', 32), K('Beauty.', 32)], tag: 'beauty care & wellness', tagSize: 16, dx: 1.797, dy: 0.775, tw: 2.074, th: 0.37 });
  frame(s, 7.659, 4.972, 5.007, 1.778);
  badge(s, 12.534, 4.838);
  box(s, 'Lorem ipsum dolor sita amet, consecter adipiscing elit. Maecenas hasat faucibus asnnal. Phasellus vitae liberos ut magna dapibus an ultrices qieusani orci. Donec daibus ultriesi hanisal', { ...COPY, fontSize: 10, x: 1.634, y: 1.83, w: 2.609, h: 1.33 });
  box(s, 'We Care About You', { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, x: 1.634, y: 1.215, w: 2.609, h: 0.337 });
  arrow(s, 1.137, 1.31, 90);
}

function slide12(s, n) {
  frame(s, 7.677, 4.75, 4.983, 2);
  photo(s, 6.667, 1.76, 5, 3.99);
  chrome(s, n);
  titleBlock(s, { x: 1.12, y: 1.77, w: 4.106, h: 0.909, head: [P('Break ', 48), K('Slide.', 48)], tag: 'beauty care & wellness', tagSize: 16, dx: 0.068, dy: 0.971, tw: 2.074, th: 0.37, align: 'center' });
  frame(s, 0.673, 0.764, 5, 2.986);
  box(s, LOREM_G, { ...COPY, fontSize: 10, x: 1.2, y: 4.279, w: 4.458, h: 1.33 });
  badge(s, 1.057, 1.135);
  arrow(s, 12.092, 6.298, 270);
  hatch(s, 6.667, 6.011, 0.742);
  box(s, 'The Story ', { ...LEAD, fontSize: 14, align: 'right', x: 3.638, y: 6.114, w: 2.035, h: 0.413 });
}

function slide13(s, n) {
  photo(s, 7.003, 1.75, 1.406, 1.406);
  photo(s, 8.628, 1.75, 1.406, 1.406);
  photo(s, 10.253, 1.75, 1.406, 1.406);
  photo(s, 7.003, 3.99, 1.406, 1.406);
  photo(s, 8.628, 3.99, 1.406, 1.406);
  photo(s, 10.253, 3.99, 1.406, 1.406);
  frame(s, 5.987, 6.233, 6.688, 0.517);
  chrome(s, n);
  titleBlock(s, { x: 1.449, y: 1.623, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  frame(s, 0.673, 0.764, 5.314, 2.392);
  badge(s, 1.05, 1.107);
  box(s, 'Keziah House', { ...LEAD, fontSize: 10.5, align: 'center', x: 7.003, y: 3.257, w: 1.406, h: 0.346 });
  box(s, 'Huma Greer', { ...LEAD, fontSize: 10.5, align: 'center', x: 8.628, y: 3.257, w: 1.406, h: 0.346 });
  box(s, 'Silas Burn', { ...LEAD, fontSize: 10.5, align: 'center', x: 10.253, y: 3.257, w: 1.406, h: 0.346 });
  box(s, 'Lexie Salas', { ...LEAD, fontSize: 10.5, align: 'center', x: 7.003, y: 5.497, w: 1.406, h: 0.346 });
  box(s, 'Kiya Bauer', { ...LEAD, fontSize: 10.5, align: 'center', x: 8.628, y: 5.497, w: 1.406, h: 0.346 });
  box(s, 'Nada Chandler', { ...LEAD, fontSize: 10.5, align: 'center', x: 10.253, y: 5.497, w: 1.406, h: 0.346 });
  box(s, 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas hanis at faucibus asnnal. Phasellus vitae libero ut magnasan dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orcisanu ace, interdumas tortor. Donec elit lorem, finibusas imperdiet orci aces, interdusan lorem, finibus magna dapibus ultriesi', { ...COPY, fontSize: 10, x: 1.449, y: 4.386, w: 4.537, h: 1.33 });
  box(s, 'Most Experienced Team', { ...LEAD, fontSize: 14, x: 1.449, y: 3.622, w: 3.231, h: 0.413 });
  arrow(s, 12.277, 6.429, 270);
}

function slide14(s, n) {
  frame(s, 0.673, 0.764, 5.994, 4.985);
  photo(s, 1.678, 1.751, 1.336, 1.336);
  photo(s, 3.014, 3.088, 1.336, 1.336);
  photo(s, 4.351, 4.424, 1.336, 1.336);
  chrome(s, n);
  frame(s, 7.677, 4.424, 4.983, 2.326);
  titleBlock(s, { x: 8.288, y: 4.791, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, [
    { text: '25', options: { fontSize: 32, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 18, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 1.656, y: 3.311, w: 1.063, h: 0.64 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, align: 'center', x: 1.586, y: 3.858, w: 1.336, h: 0.298 });
  badge(s, 1.057, 1.144);
  box(s, 'Nada Chandler', { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, x: 3.35, y: 2.068, w: 1.093, h: 0.572 });
  box(s, 'Huma Greer', { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, x: 4.668, y: 3.464, w: 0.773, h: 0.572 });
  box(s, [
    { text: 'Keziah ', options: { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, breakLine: true } },
    { text: 'House', options: { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true } },
  ], { x: 1.829, y: 4.753, w: 0.961, h: 0.572 });
  box(s, 'Lorem ipsi dolor sit amet, consec maecenas', { ...COPY, fontSize: 9, x: 3.014, y: 4.742, w: 1.093, h: 0.753 });
  box(s, 'Lorem ipsi dolor adipiscing al elit. maecenas', { ...COPY, fontSize: 9, x: 4.673, y: 2.051, w: 1.093, h: 0.753 });
  arrow(s, 12.092, 6.298, 270);
  box(s, LOREM_G, { ...COPY, fontSize: 10, x: 7.94, y: 2.508, w: 4.458, h: 1.33 });
  box(s, 'Most Experienced Team', { ...LEAD, fontSize: 14, x: 7.94, y: 1.77, w: 3.231, h: 0.413 });
}

function slide15(s, n) {
  frame(s, 6.88, 3.75, 5.787, 3.003);
  frame(s, 0.673, 0.764, 5.804, 2.986);
  photo(s, 1.678, 2.647, 2.198, 2.206);
  photo(s, 4.279, 2.647, 2.198, 2.206);
  photo(s, 6.88, 2.647, 2.198, 2.206);
  photo(s, 9.481, 2.647, 2.198, 2.206);
  chrome(s, n);
  badge(s, 1.057, 1.144);
  titleBlock(s, { x: 1.929, y: 1.167, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, LOREM_B, { ...COPY, fontSize: 9, x: 1.937, y: 5.608, w: 1.68, h: 0.753 });
  box(s, 'Keziah House', { ...LEAD, fontSize: 12, x: 1.937, y: 5.14, w: 1.68, h: 0.368 });
  box(s, LOREM_B, { ...COPY, fontSize: 9, x: 4.538, y: 5.608, w: 1.68, h: 0.753 });
  box(s, 'Huma Greer', { ...LEAD, fontSize: 12, x: 4.538, y: 5.14, w: 1.68, h: 0.368 });
  box(s, LOREM_B, { ...COPY, fontSize: 9, x: 7.398, y: 5.608, w: 1.68, h: 0.753 });
  box(s, 'Lexie Salas', { ...LEAD, fontSize: 12, x: 7.398, y: 5.14, w: 1.68, h: 0.368 });
  box(s, LOREM_B, { ...COPY, fontSize: 9, x: 9.773, y: 5.608, w: 1.68, h: 0.753 });
  box(s, 'Nada Chandler', { ...LEAD, fontSize: 12, x: 9.773, y: 5.14, w: 1.68, h: 0.368 });
  arrow(s, 12.092, 6.256, 270);
  box(s, 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitaesanu libero ut magna dapibus ultrices qieu orcieu tellus comildo ', { ...COPY, fontSize: 10, x: 7.398, y: 1.377, w: 4.056, h: 0.825 });
}

function slide16(s, n) {
  frame(s, 6.667, 3.75, 6, 3.003);
  photo(s, 6.661, 3.75, 2.005, 2.005);
  photo(s, 9.667, 3.75, 2.005, 2.005);
  chrome(s, n);
  frame(s, 0.673, 0.764, 4.999, 4.985);
  arrow(s, 12.092, 6.204, 270);
  box(s, [
    { text: '81', options: { fontSize: 32, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 18, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 7.064, y: 1.9, w: 1.063, h: 0.64 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, align: 'center', x: 6.993, y: 2.539, w: 1.336, h: 0.298 });
  titleBlock(s, { x: 1.449, y: 1.817, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  badge(s, 1.05, 1.107);
  box(s, 'Lorem ipsum dolor sit amet, hani consecter adipiscing elit. Maecenas hanis at faucibus alasnnal. Phasellus vitae libero ut magnasan dapibus ultrices qieusali orci. Donec eu tellus comildo, imperdiet orcisanu acesanusa, interdumas tortor. Donec, finibus magna ultriesi', { ...COPY, fontSize: 10, x: 1.465, y: 3.593, w: 3.709, h: 1.33 });
  hatch(s, 5.042, 5.118, 0.625);
  box(s, [
    { text: '77', options: { fontSize: 32, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 18, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 10.071, y: 1.904, w: 1.063, h: 0.64 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, align: 'center', x: 10.001, y: 2.543, w: 1.336, h: 0.298 });
  s.addShape('line', { line: { color: PURPLE }, x: 9.168, y: 1.993, w: 0, h: 0.767 });
  box(s, 'Keziah House', { ...LEAD, fontSize: 12, align: 'center', x: 6.84, y: 6.069, w: 1.826, h: 0.368 });
  box(s, 'Nada Chandler', { ...LEAD, fontSize: 12, align: 'center', x: 9.661, y: 6.069, w: 2.01, h: 0.368 });
}

function slide17(s, n) {
  frame(s, 0.673, 0.764, 4.999, 4.985);
  photo(s, 1.678, 1.743, 1.909, 1.926);
  photo(s, 3.763, 1.743, 1.909, 1.926);
  photo(s, 1.678, 3.825, 1.909, 1.926);
  photo(s, 3.763, 3.825, 1.909, 1.926);
  chrome(s, n);
  badge(s, 1.057, 1.144);
  frame(s, 6.667, 3.825, 6, 2.929);
  titleBlock(s, { x: 6.567, y: 1.797, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  hatch(s, 10.869, 1.958, 0.795);
  box(s, 'Lorem ipsum dolor sit amet, hani consecter adipiscing elit. Maecenas hanis at faucibus alasnnal. Phasellus vitae libero ut magnasan han dapibus ultrices qieusali orci. Donec eu tellus comildo, imperdiet nusan orcisanu acesanusa, interdumas tortor. Donec, finibus magnasan ultriesi magnasan han dapibus ultrices hanisi', { ...COPY, fontSize: 10, x: 7.17, y: 4.249, w: 4.611, h: 1.33 });
  arrow(s, 12.092, 6.159, 270);
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 7.17, y: 6.003, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: GREY_D, x: 8.653, y: 6.082, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: GREY_D, align: 'right', x: 10.771, y: 6.08, w: 1.01, h: 0.269 });
}

function slide18(s, n) {
  frame(s, 6.667, 3.75, 6, 3.003);
  frame(s, 0.673, 0.764, 5.994, 2.986);
  photo(s, 4.667, 1.722, 4, 4.056);
  chrome(s, n);
  badge(s, 1.05, 1.107);
  box(s, 'Keziah House', { ...LEAD, fontSize: 20, x: 9.342, y: 2.553, w: 2.328, h: 0.547 });
  box(s, 'Lorem ipsum dolor sit nasanil amet, at faucibus asnnal. Phasellus vitae qieu orci. Donec eu tellus comildo, tortor. Donec elit lorem, finibusas imperdiet orci acesanu, interdusan lorem, finibus magna hanis dapibus ultriesi finibusas ', { ...COPY, fontSize: 10, x: 9.342, y: 4.298, w: 2.322, h: 1.835 });
  arrow(s, 12.092, 6.204, 270);
  box(s, [
    { text: '81', options: { fontSize: 24, color: PURPLE, bold: true } },
    { text: '%', options: { fontSize: 14, color: PURPLE, bold: true } },
  ], { x: 1.434, y: 4.407, w: 0.77, h: 0.505 });
  box(s, 'Lorem ipsum dolor assit amet, hanisa nuna dasat', { ...COPY, fontSize: 9, x: 2.484, y: 4.381, w: 1.509, h: 0.525 });
  box(s, [
    { text: '76', options: { fontSize: 24, color: PURPLE, bold: true } },
    { text: '%', options: { fontSize: 14, color: PURPLE, bold: true } },
  ], { x: 1.434, y: 5.329, w: 0.77, h: 0.505 });
  box(s, 'Lorem ipsum dolor assit amet, hanisa nuna dasat', { ...COPY, fontSize: 9, x: 2.484, y: 5.304, w: 1.509, h: 0.525 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, x: 9.34, y: 3.127, w: 1.509, h: 0.298 });
  titleBlock(s, { x: 1.434, y: 2.077, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
}

function slide19(s, n) {
  frame(s, 0.673, 0.764, 4.999, 4.985);
  photo(s, 1.678, 1.76, 4.001, 3.997);
  s.addShape('rect', { fill: PURPLE, x: 5.679, y: 1.76, w: 0.995, h: 3.018 });
  badge(s, 6.044, 2.099, true);
  chrome(s, n);
  titleBlock(s, { x: 7.445, y: 1.976, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, 'Lorem ipsum dolor sit amet, hani consecter adipiscing elit. Maecenas hanis at faucibus hanas alasnnal. Phasellus vitae libero ut magnasan dapibus nisa ultrices qieusali orci. Donec eu tellus comildo, imperdiet acesanusa, interdumas tortor. ', { ...COPY, fontSize: 10, x: 7.445, y: 3.628, w: 4.326, h: 1.078 });
  frame(s, 6.667, 5.74, 6, 1.014);
  box(s, [
    { text: '81', options: { fontSize: 24, color: PURPLE, bold: true } },
    { text: '%', options: { fontSize: 14, color: PURPLE, bold: true } },
  ], { x: 7.445, y: 5.994, w: 0.77, h: 0.505 });
  box(s, 'Lorem ipsum dolor sit amet, hanisanun', { ...COPY, fontSize: 9, x: 8.278, y: 5.969, w: 1.248, h: 0.525 });
  box(s, [
    { text: '76', options: { fontSize: 24, color: PURPLE, bold: true } },
    { text: '%', options: { fontSize: 14, color: PURPLE, bold: true } },
  ], { x: 9.69, y: 5.992, w: 0.77, h: 0.505 });
  box(s, 'Lorem ipsum dolor sit amet, hanisanun', { ...COPY, fontSize: 9, x: 10.522, y: 5.966, w: 1.248, h: 0.525 });
  box(s, 'Nada Chandler', { fontFace: HEAD, fontSize: 12, color: WHITE, bold: true, lineSpacingMultiple: 1.5, rotate: 270, x: 5.31, y: 3.311, w: 1.67, h: 0.368 });
  arrow(s, 12.092, 6.204, 270);
}

function slide20(s, n) {
  frame(s, 6.667, 3.75, 6, 3.003);
  photo(s, 7.664, 1.764, 4, 4);
  chrome(s, n);
  frame(s, 0.673, 0.764, 5.989, 1);
  titleBlock(s, { x: 1.66, y: 2.285, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, 'Lorem ipsum dolor sitasa amet, hani consecter adipiscingasa elit Maecenas hanis at faucibus alasnnal. Phasellus vitae libero ut magnasan han dapibus ultrices qieusali orci. Doneasce eu tellus comildo, imperdiet nusan orcisanu acesanusa, interdumas tortor Donec, finibus magnasan ultriesi magnasan han dapibus ultrices hanisi magnasan ultriesi magnasana. Phasellus vitae libero ut magnasan han dapibus ultrices', { ...COPY, fontSize: 10, x: 1.669, y: 3.949, w: 4.14, h: 1.835 });
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 1.092, y: 1.047, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: RED, x: 2.575, y: 1.125, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: RED, align: 'right', x: 4.617, y: 1.124, w: 1.01, h: 0.269 });
  arrow(s, 6.107, 1.198, 270);
  box(s, 'Hanifa Pritchard', { ...LEAD, fontSize: 16, align: 'center', x: 7.664, y: 6.037, w: 4, h: 0.458 });
  badge(s, 12.019, 6.146);
}

function slide21(s, n) {
  frame(s, 0.673, 0.764, 4.999, 1.986);
  photo(s, 1.678, 1.758, 2.997, 3.992);
  chrome(s, n);
  badge(s, 5.054, 1.989);
  titleBlock(s, { x: 6.431, y: 1.811, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, 'Lorem ipsum dolor sit amet, hani consecter adipiscing elit. Maecenas hanis at faucibus alasnnal. Phasellus vitae libero ut magnasan han dapibus ultrices qieusali orci. Donec eu tellus comildo, imperdiet nusan orcisanu acesanusa, interdumas tortor. Donec, finibus magnasan ultriesi magnasan han dapibus ultrices hanisi at faucibus alasnnal. Phasellus vitae libero ut magnasan hane', { ...COPY, fontSize: 10, x: 6.462, y: 4.287, w: 5.333, h: 1.33 });
  frame(s, 5.68, 3.75, 6.987, 3.003);
  hatch(s, 10.88, 1.971, 0.795);
  arrow(s, 12.092, 6.204, 270);
  box(s, 'Kaitlan Atkinson', { ...LEAD, fontSize: 14, rotate: 270, x: -0.306, y: 4.323, w: 2.675, h: 0.413 });
  box(s, 'Most Experienced Team', { ...LEAD, fontSize: 12, align: 'right', x: 8.564, y: 6.052, w: 3.231, h: 0.368 });
}

function slide22(s, n) {
  frame(s, 0.673, 0.764, 5.994, 2.986);
  frame(s, 6.667, 3.75, 6, 3.003);
  photo(s, 5.677, 1.762, 3.99, 3.984);
  chrome(s, n);
  box(s, [
    { text: '81', options: { fontSize: 32, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 18, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 10.569, y: 2.083, w: 1.063, h: 0.64 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, align: 'center', x: 10.498, y: 2.721, w: 1.336, h: 0.298 });
  box(s, [
    { text: '76', options: { fontSize: 32, color: PURPLE, bold: true, align: 'right' } },
    { text: '%', options: { fontSize: 18, color: PURPLE, bold: true, align: 'right' } },
  ], { x: 10.569, y: 4.452, w: 1.063, h: 0.64 });
  box(s, 'Lorem ipsum dolor', { ...COPY, fontSize: 9, align: 'center', x: 10.498, y: 5.091, w: 1.336, h: 0.298 });
  arrow(s, 12.092, 6.204, 270);
  badge(s, 1.05, 1.107);
  titleBlock(s, { x: 1.339, y: 1.995, w: 3.761, h: 1.111, head: [P('The Beauty Care ', 32), K('Professional.', 28)], tag: 'beauty expert', tagSize: 12, dx: 2.494, dy: 0.764, tw: 1.188, th: 0.303 });
  box(s, 'Lorem ipsum dolor sitasa amet, hani consecter Phasellus vitae libero ut magnasan han dapibus ultrices qieusali orci. Doneasce eu tellus comildo, imperdiet at nusan orcisanu acesanusa, interdumas tortor Donec, finibus magnasan ultriesi magnasan hanes dapibus ultrices hanisi magnasan han dapibus ultrices', { ...COPY, fontSize: 10, x: 1.333, y: 4.243, w: 3.719, h: 1.582 });
  s.addShape('rect', { fill: PURPLE, x: 6.238, y: 5.458, w: 2.867, h: 0.645 });
  box(s, 'Niamh Medina', { fontFace: HEAD, fontSize: 14, color: WHITE, bold: true, align: 'center', lineSpacingMultiple: 1.5, x: 6.228, y: 5.546, w: 2.877, h: 0.413 });
}

function slide23(s, n) {
  frame(s, 5.682, 4.708, 6.985, 2.042);
  photo(s, 5.682, 0.992, 1.97, 1.657);
  photo(s, 5.682, 2.85, 1.97, 1.657);
  photo(s, 5.682, 4.708, 1.97, 1.657);
  chrome(s, n);
  frame(s, 0.673, 0.764, 4.18, 3.743);
  titleBlock(s, { x: 8.218, y: 5.141, w: 3.898, h: 1.043, head: [P('The Photograph Of ', 28), K('Our Work.', 28)], tag: 'beauty care & wellness', tagSize: 13, dx: 2.021, dy: 0.691, tw: 1.74, th: 0.32 });
  box(s, LOREM_D, { ...COPY, fontSize: 10, align: 'right', x: 1.094, y: 1.677, w: 2.98, h: 0.825 });
  box(s, 'Body Scrub and Spa', { ...LEAD, fontSize: 13, align: 'right', x: 1.089, y: 1.139, w: 2.98, h: 0.391 });
  s.addText('01', { fontFace: BODY, fontSize: 11, color: INK, bold: true, align: 'center', fill: PURPLE, shape: 'ellipse', valign: 'middle', x: 4.565, y: 1.532, w: 0.576, h: 0.576 });
  s.addText('02', { fontFace: BODY, fontSize: 11, color: INK, bold: true, align: 'center', fill: PURPLE, shape: 'ellipse', valign: 'middle', x: 4.565, y: 3.391, w: 0.576, h: 0.576 });
  s.addText('03', { fontFace: BODY, fontSize: 11, color: INK, bold: true, align: 'center', fill: PURPLE, shape: 'ellipse', valign: 'middle', x: 4.565, y: 5.249, w: 0.576, h: 0.576 });
  box(s, LOREM_D, { ...COPY, fontSize: 10, align: 'right', x: 1.094, y: 3.535, w: 2.98, h: 0.825 });
  box(s, 'Skin & Hair Care', { ...LEAD, fontSize: 13, align: 'right', x: 1.089, y: 2.997, w: 2.98, h: 0.391 });
  box(s, LOREM_D, { ...COPY, fontSize: 10, align: 'right', x: 1.094, y: 5.393, w: 2.98, h: 0.825 });
  box(s, 'Pedicure Manicure', { ...LEAD, fontSize: 13, align: 'right', x: 1.089, y: 4.856, w: 2.98, h: 0.391 });
  box(s, 'Lorem ipsum dolor sit amet, consecter adipiscing elit. Maecenas at faucibus asnnal. Phasellus vitae libero ut magna dapibus ultrices qieu orci. Donec eu tellus comildo, imperdiet orci ace, interdumas tortor. Donec elit lorem, finibusas imperdiet orci ac, interdu lorem finibus magna dapibus ultriesi', { ...COPY, fontSize: 10, x: 8.218, y: 2.745, w: 4.309, h: 1.33 });
  box(s, 'ABOUT US :', { ...LEAD, fontSize: 12, x: 8.218, y: 1.576, w: 1.504, h: 0.368 });
  box(s, 'J. Maecenas at faucis No.11', { fontSize: 10.5, color: RED, x: 9.448, y: 1.654, w: 2.097, h: 0.278 });
  box(s, '(88776221x)', { fontSize: 10, color: RED, align: 'right', x: 11.49, y: 1.653, w: 1.01, h: 0.269 });
  arrow(s, 0.989, 1.059, 90);
}

function slide24(s, n) {
  frame(s, 6.667, 3.75, 6, 3.003);
  frame(s, 0.673, 0.764, 5.994, 2.986);
  photo(s, 1.678, 0.992, 3.779, 3.779);
  photo(s, 7.876, 2.971, 3.779, 3.779);
  chrome(s, n);
  box(s, LOREM_A, { ...COPY, fontSize: 10, x: 2.233, y: 5.683, w: 3.224, h: 0.825 });
  box(s, 'Body Scrub and Spa', { ...LEAD, fontSize: 14, x: 2.229, y: 5.146, w: 3.224, h: 0.413 });
  box(s, 'Lorem ipsum dolor sit amet, consecter adipiscing tortor. Donec elit lorem, finibusas imperdiet orcia inter lorem finibusa hansaun', { ...COPY, fontSize: 10, x: 8.427, y: 1.701, w: 3.224, h: 0.825 });
  box(s, 'Reflexology', { ...LEAD, fontSize: 14, x: 8.423, y: 1.163, w: 3.224, h: 0.413 });
  badge(s, 1.052, 1.184);
  box(s, 'The Satisfy Clients', { fontSize: 10, color: PURPLE, lineSpacingMultiple: 1.5, rotate: 270, x: 0.418, y: 2.466, w: 1.45, h: 0.32 });
  badge(s, 12.062, 4.2);
  box(s, 'The Satisfy Clients', { fontSize: 10, color: PURPLE, lineSpacingMultiple: 1.5, rotate: 270, x: 11.428, y: 5.482, w: 1.45, h: 0.32 });
  arrow(s, 1.668, 5.326, 90);
  arrow(s, 7.866, 1.339, 90);
}

function slide25(s, n) {
  frame(s, 6.667, 4.772, 6, 1.978);
  frame(s, 0.673, 0.764, 5.994, 2.868);
  photo(s, 1.678, 1.654, 1.978, 1.978);
  photo(s, 1.669, 3.868, 1.978, 1.978);
  photo(s, 7.471, 3.868, 1.978, 1.978);
  photo(s, 9.686, 3.877, 1.978, 1.978);
  chrome(s, n);
  titleBlock(s, { x: 7.677, y: 1.569, w: 3.898, h: 1.043, head: [P('The Photograph Of ', 28), K('Our Work.', 28)], tag: 'beauty care & wellness', tagSize: 13, dx: 2.021, dy: 0.691, tw: 1.74, th: 0.32 });
  box(s, 'Lorem ipsum dolor sit amet, consecter asilas adipiscing elit Maecen at faucibus asnnal. ', { ...COPY, fontSize: 10, x: 7.714, y: 2.878, w: 3.725, h: 0.573 });
  box(s, LOREM_B, { ...COPY, fontSize: 10, x: 4.246, y: 2.421, w: 1.871, h: 0.825 });
  box(s, 'Reflexology', { ...LEAD, fontSize: 14, x: 4.243, y: 1.883, w: 1.871, h: 0.413 });
  box(s, LOREM_B, { ...COPY, fontSize: 10, x: 4.248, y: 4.713, w: 1.871, h: 0.825 });
  box(s, 'Skin & Hair Care', { ...LEAD, fontSize: 14, x: 4.246, y: 4.176, w: 1.871, h: 0.413 });
  badge(s, 1.052, 1.12);
  box(s, 'The Satisfy Clients', { fontSize: 10, color: PURPLE, lineSpacingMultiple: 1.5, rotate: 270, x: 0.418, y: 2.402, w: 1.45, h: 0.32 });
  badge(s, 12.043, 6.175);
  box(s, 'Pedicure Manicure', { ...LEAD, fontSize: 11, align: 'center', x: 7.471, y: 6.105, w: 1.978, h: 0.346 });
  box(s, 'Body Scrub and Spa', { ...LEAD, fontSize: 11, align: 'center', x: 9.727, y: 6.105, w: 1.895, h: 0.346 });
}

function slide26(s, n) {
  frame(s, 6.778, 3.851, 5.889, 2.899);
  frame(s, 0.666, 0.764, 3.359, 2.884);
  photo(s, 1.678, 1.753, 2.354, 3.994);
  photo(s, 6.778, 1.753, 2.354, 1.895);
  photo(s, 6.778, 3.851, 2.354, 1.895);
  photo(s, 9.328, 1.753, 2.354, 3.994);
  chrome(s, n);
  badge(s, 1.052, 1.125);
  box(s, 'Body Scrub and Spa', { ...LEAD, fontSize: 11, align: 'center', rotate: 270, x: 0.198, y: 2.36, w: 1.895, h: 0.346 });
  box(s, 'Skin & Hair Care', { ...LEAD, fontSize: 11, align: 'center', x: 9.328, y: 6.104, w: 2.354, h: 0.346 });
  arrow(s, 12.072, 6.227, 270);
  titleBlock(s, { x: 4.545, y: 2.268, w: 1.842, h: 2.255, head: [P('Taking Care Your ', 32), K('Beauty.', 32)], tag: 'beauty care & wellness', tagSize: 14, dx: 0.01, dy: 2.627, tw: 1.842, th: 0.337, tagAlign: 'left' });
  box(s, 'Pedicure Manicure', { ...LEAD, fontSize: 11, align: 'center', x: 6.778, y: 1.085, w: 2.354, h: 0.346 });
  box(s, 'Reflexology', { ...LEAD, fontSize: 11, align: 'center', x: 6.781, y: 6.104, w: 2.354, h: 0.346 });
}

function slide27(s, n) {
  frame(s, 8.457, 3.77, 4.207, 2.986);
  frame(s, 0.666, 0.764, 3.359, 2.986);
  photo(s, 1.678, 1.762, 3.225, 3.495);
  photo(s, 5.068, 1.75, 3.225, 3.495);
  photo(s, 8.457, 1.75, 3.225, 3.495);
  chrome(s, n);
  box(s, LOREM_E, { ...COPY, fontSize: 10, align: 'center', x: 1.91, y: 5.97, w: 2.765, h: 0.573 });
  box(s, 'Reflexology', { ...LEAD, fontSize: 12, align: 'center', x: 1.906, y: 5.496, w: 2.765, h: 0.368 });
  box(s, LOREM_E, { ...COPY, fontSize: 10, align: 'center', x: 5.3, y: 5.97, w: 2.765, h: 0.573 });
  box(s, 'Pedicure Manicure', { ...LEAD, fontSize: 12, align: 'center', x: 5.296, y: 5.496, w: 2.765, h: 0.368 });
  box(s, LOREM_E, { ...COPY, fontSize: 10, align: 'center', x: 8.816, y: 5.97, w: 2.765, h: 0.573 });
  box(s, 'Skin & Hair Care', { ...LEAD, fontSize: 12, align: 'center', x: 8.812, y: 5.496, w: 2.765, h: 0.368 });
  arrow(s, 12.072, 6.165, 270);
  badge(s, 1.068, 1.161);
  box(s, 'The Satisfy Clients', { fontSize: 10, color: PURPLE, lineSpacingMultiple: 1.5, rotate: 270, x: 0.434, y: 2.418, w: 1.45, h: 0.32 });
}

function slide28(s, n) {
  chrome(s, n);
  s.addShape('line', { line: { color: RULE }, x: 4.195, y: 3.528, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 4.195, y: 3.956, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 4.195, y: 4.387, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 4.195, y: 4.815, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 4.195, y: 5.244, w: 2.013, h: 0 });
  s.addShape('roundRect', { rectRadius: 0.072, line: { color: RULE }, x: 4.092, y: 2.617, w: 2.248, h: 3.265 });
  s.addText('Order now', { fontFace: HEAD, fontSize: 9, color: WHITE, align: 'center', shape: 'roundRect', rectRadius: 0.157, fill: INK, margin: 0, valign: 'middle', x: 4.482, y: 5.406, w: 1.425, h: 0.339 });
  s.addShape('roundRect', { rectRadius: 0.114, line: { color: CARD_EDGE }, x: 3.947, y: 1.734, w: 2.535, h: 4.316 });
  box(s, [
    { text: '$', options: { fontFace: HEAD, fontSize: 32, color: INK, superscript: true, align: 'center' } },
    { text: '199', options: { fontFace: HEAD, fontSize: 32, color: INK, align: 'center' } },
  ], { wrap: false, x: 4.628, y: 2.769, w: 1.107, h: 0.64 });
  s.addShape('line', { line: { color: RULE }, x: 7.062, y: 3.528, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 7.062, y: 3.956, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 7.062, y: 4.387, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 7.062, y: 4.815, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 7.062, y: 5.244, w: 2.013, h: 0 });
  s.addShape('roundRect', { rectRadius: 0.072, line: { color: RULE }, x: 6.959, y: 2.617, w: 2.248, h: 3.265 });
  s.addText('Order now', { fontFace: HEAD, fontSize: 9, color: WHITE, align: 'center', shape: 'roundRect', rectRadius: 0.157, fill: PURPLE, margin: 0, valign: 'middle', x: 7.35, y: 5.406, w: 1.425, h: 0.339 });
  s.addShape('roundRect', { rectRadius: 0.114, line: { color: CARD_EDGE }, x: 6.815, y: 1.734, w: 2.535, h: 4.316 });
  box(s, [
    { text: '$', options: { fontFace: HEAD, fontSize: 32, color: PURPLE, superscript: true, align: 'center' } },
    { text: '299', options: { fontFace: HEAD, fontSize: 32, color: PURPLE, align: 'center' } },
  ], { wrap: false, x: 7.433, y: 2.769, w: 1.232, h: 0.687 });
  s.addShape('line', { line: { color: RULE }, x: 9.93, y: 3.528, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 9.93, y: 3.956, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 9.93, y: 4.387, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 9.93, y: 4.815, w: 2.013, h: 0 });
  s.addShape('line', { line: { color: RULE }, x: 9.93, y: 5.244, w: 2.013, h: 0 });
  s.addShape('roundRect', { rectRadius: 0.072, line: { color: RULE }, x: 9.827, y: 2.617, w: 2.248, h: 3.265 });
  s.addText('Order now', { fontFace: HEAD, fontSize: 9, color: WHITE, align: 'center', shape: 'roundRect', rectRadius: 0.157, fill: PURPLE_LT, margin: 0, valign: 'middle', x: 10.217, y: 5.406, w: 1.425, h: 0.339 });
  s.addShape('roundRect', { rectRadius: 0.114, line: { color: CARD_EDGE }, x: 9.682, y: 1.734, w: 2.535, h: 4.316 });
  box(s, [
    { text: '$', options: { fontFace: HEAD, fontSize: 32, color: PURPLE_LT, superscript: true, align: 'center' } },
    { text: '399', options: { fontFace: HEAD, fontSize: 32, color: PURPLE_LT, align: 'center' } },
  ], { wrap: false, x: 10.339, y: 2.769, w: 1.154, h: 0.64 });
  box(s, 'First  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 4.69, y: 3.61, w: 1.073, h: 0.286 });
  box(s, 'First  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 7.558, y: 3.61, w: 1.073, h: 0.286 });
  box(s, 'First  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 10.425, y: 3.61, w: 1.073, h: 0.286 });
  box(s, 'Second  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 4.601, y: 4.039, w: 1.252, h: 0.286 });
  box(s, 'Second  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 7.468, y: 4.039, w: 1.252, h: 0.286 });
  box(s, 'Second  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 10.336, y: 4.039, w: 1.252, h: 0.286 });
  box(s, 'Third  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 4.665, y: 4.469, w: 1.124, h: 0.286 });
  box(s, 'Third  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 7.532, y: 4.469, w: 1.124, h: 0.286 });
  box(s, 'Third  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 10.4, y: 4.469, w: 1.124, h: 0.286 });
  box(s, 'Forth  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 4.658, y: 4.897, w: 1.138, h: 0.286 });
  box(s, 'Forth  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 7.525, y: 4.897, w: 1.138, h: 0.286 });
  box(s, 'Forth  feature', { fontSize: 11, color: GREY, align: 'center', wrap: false, x: 10.393, y: 4.897, w: 1.138, h: 0.286 });
  s.addText('Advance', { fontFace: HEAD, fontSize: 16, color: WHITE, align: 'center', shape: 'round2SameRect', fill: INK, line: { color: INK }, margin: [7.2, 7.2, 0, 0], valign: 'middle', x: 3.947, y: 1.734, w: 2.535, h: 0.728 });
  s.addText('Professional', { fontFace: HEAD, fontSize: 16, color: WHITE, align: 'center', shape: 'round2SameRect', fill: PURPLE, line: { color: PURPLE }, margin: [7.2, 7.2, 0, 0], valign: 'middle', x: 6.815, y: 1.734, w: 2.535, h: 0.728 });
  s.addText('Business', { fontFace: HEAD, fontSize: 16, color: WHITE, align: 'center', shape: 'round2SameRect', fill: PURPLE_LT, line: { color: PURPLE_LT }, margin: [7.2, 7.2, 0, 0], valign: 'middle', x: 9.682, y: 1.734, w: 2.535, h: 0.728 });
  titleBlock(s, { x: 1.21, y: 2.269, w: 1.842, h: 2.457, head: [P('About', 35, true), P('Our', 35, true), K('Pricing Plan.', 35)], tag: 'beauty care & wellness', tagSize: 14, dx: 0.01, dy: 2.807, tw: 1.842, th: 0.337, tagAlign: 'left' });
  frame(s, 0.666, -0.474, 5.816, 1.466);
  frame(s, 6.815, 6.748, 7.019, 1.1);
  badge(s, 0.705, 1.783);
  arrow(s, 12.884, 7.051, 270);
}

function slide29(s, n) {
  frame(s, 6.667, 3.77, 5.997, 2.986);
  photo(s, 6.667, 1.755, 4.999, 4.009);
  chrome(s, n);
  frame(s, 0.666, 0.764, 4.999, 2.986);
  titleBlock(s, { x: 1.359, y: 1.766, w: 3.898, h: 1.043, head: [P('Get In Touch With ', 28), K('Us Now.', 28)], tag: 'beauty care & wellness', tagSize: 13, dx: 1.789, dy: 0.691, tw: 1.74, th: 0.32 });
  badge(s, 1.052, 1.12);
  box(s, [
    { text: 'J. Maecenas at ', options: { fontSize: 12, color: GREY_D, breakLine: true } },
    { text: 'Faucis, nandasa 111', options: { fontSize: 12, color: GREY_D, breakLine: true } },
    { text: 'No.11', options: { fontSize: 12, color: GREY_D } },
  ], { wrap: false, x: 1.331, y: 5.121, w: 1.68, h: 0.707 });
  box(s, 'ADDRESS ONE :', { fontFace: HEAD, fontSize: 12, color: INK, bold: true, x: 1.331, y: 4.566, w: 1.839, h: 0.303 });
  box(s, [
    { text: 'J. Maecenas at ', options: { fontSize: 12, color: GREY_D, breakLine: true } },
    { text: 'Faucis, nandasa 111', options: { fontSize: 12, color: GREY_D, breakLine: true } },
    { text: 'No.11', options: { fontSize: 12, color: GREY_D } },
  ], { wrap: false, x: 3.418, y: 5.121, w: 1.68, h: 0.707 });
  box(s, 'SECOND OFFICE:', { fontFace: HEAD, fontSize: 12, color: INK, bold: true, x: 3.418, y: 4.566, w: 1.839, h: 0.303 });
  arrow(s, 12.072, 6.165, 270);
  box(s, 'yourname.com/follow', { fontSize: 12, color: INK, bullet: { characterCode: '2022', indent: 13.5 }, wrap: false, x: 7.126, y: 6.086, w: 1.999, h: 0.303 });
  box(s, 'yourname.com/follow', { fontSize: 12, color: INK, bullet: { characterCode: '2022', indent: 13.5 }, wrap: false, x: 9.37, y: 6.086, w: 1.999, h: 0.303 });
}

function slide30(s, n) {
  frame(s, 7.656, 2.74, 5.01, 4.01);
  frame(s, 0.667, 0.75, 5.01, 4.01);
  photo(s, 1.678, 1.77, 9.989, 3.96);
  chrome(s, n);
  s.addShape('rect', { fill: WHITE, x: 4.158, y: 3.323, w: 5.017, h: 0.873 });
  badge(s, 1.057, 1.135);
  box(s, 'Thank You For Your Time', { fontFace: HEAD, fontSize: 14, color: GREY_M, bold: true, x: 1.63, y: 1.103, w: 2.609, h: 0.337 });
  arrow(s, 12.072, 6.165, 270);
  box(s, [
    { text: 'Thank ', options: { fontFace: HEAD, fontSize: 44, color: INK, bold: true, align: 'center' } },
    { text: 'You.', options: { fontFace: HEAD, fontSize: 44, color: PURPLE, bold: true, align: 'center' } },
  ], { x: 4.031, y: 3.329, w: 5.427, h: 0.841 });
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------
const deck = new PptxGenJS();
deck.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
deck.layout = 'WIDE';
deck.theme = { headFontFace: HEAD, bodyFontFace: BODY };

SLIDES.forEach((build, i) => build(deck.addSlide(), i + 1));

deck.writeFile({ fileName: path.join(__dirname, '084033d8-f1da-4bc0-8cea-c2f1e00991b2_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
