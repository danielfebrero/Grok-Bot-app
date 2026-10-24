#!/usr/bin/env node
/*
 * "Academize" university template - 36 slide deck rebuilt with pptxgenjs.
 * Raster photos in the source deck are replaced by flat [image] placeholder
 * rectangles that keep the original position / size.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / typography
 * ------------------------------------------------------------------ */
const RED = 'C70039';   // brand crimson
const ORANGE = 'FF5733';
const LIME = 'DAF7A6';
const GREEN = '274B23';
const GREEN_M = '4F9845';
const GREEN_L = '9DCF96';
const CYAN = '00B0F0';
const MAGENTA = 'FF33FB';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const DARK = '3F3F3F';
const GREY = '7F7F7F';   // body copy
const GREY_D = '666666';
const GREY_M = 'A5A5A5';
const GREY_BOX = '7F7F7F';   // icon tile on slide 2
const GREY_CARD = 'F2F2F2';
const GREY_TRACK = 'D8D8D8';
const GREY_RULE = 'BFBFBF';
const SKIN = 'F0C8A4';

const PHOTO_BG = 'E4E4E4';   // stand-in for a photograph
const PHOTO_FG = '9A9A9A';

const HEAD = 'Roboto';
const BODY = 'Open Sans';

/* Filler copy used all over the template. */
const L_SHORT = 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.';
const L_MED = L_SHORT + ' I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine.';
const L_FULL = L_MED + ' I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents.';
const L_TINY = 'A wonderful serenity has taken possession of my entire soul,';
const UP = (s) => s.toUpperCase();

/* ------------------------------------------------------------------ *
 * Small drawing helpers
 * ------------------------------------------------------------------ */
function box(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, opts));
}

function outline(s, shape, x, y, w, h, color, pt, opts) {
  s.addShape(shape, Object.assign(
    { x, y, w, h, fill: { type: 'none' }, line: { color, width: pt } }, opts));
}

/* The source deck uses 0.1" side / 0.05" top-bottom text insets everywhere.
 * pptxgenjs reads the array as [left, right, bottom, top] in points. */
const INSET = [7.2, 7.2, 3.6, 3.6];

function txt(s, str, o) {
  s.addText(str, Object.assign({
    fontFace: BODY, fontSize: 11, color: GREY, align: 'left', valign: 'top',
    margin: INSET, wrap: true,
  }, o));
}

/* 36pt Roboto bold crimson headline (the deck's standard title). */
function title(s, str, x, y, w, h, o) {
  txt(s, str, Object.assign({ x, y, w, h, fontFace: HEAD, fontSize: 36, bold: true, color: RED }, o));
}

/* 11pt bold crimson eyebrow label. */
function eyebrow(s, str, x, y, w, h, o) {
  txt(s, str, Object.assign({ x, y, w, h, fontSize: 11, bold: true, color: RED }, o));
}

/* Flat placeholder standing in for a photograph. The caption is parked in the
 * top-left corner so it never collides with copy laid over the picture. */
function photo(s, x, y, w, h, shape) {
  s.addShape(shape || 'rect', { x, y, w, h, fill: { color: PHOTO_BG } });
  s.addText('[image]', {
    x: x + (shape === 'ellipse' ? w * 0.3 : 0.06), y: y + 0.05,
    w: 0.9, h: 0.25, align: 'left', valign: 'top', margin: 0,
    fontFace: BODY, fontSize: 10, color: PHOTO_FG,
  });
}

/* Phone / monitor mock-ups on slides 25 and 26. */
function phoneMock(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.16, fill: { color: '1A1A1A' } });
  photo(s, x + 0.08, y + 0.09, w - 0.16, h - 0.18);
}

function monitorMock(s, x, y, w, h) {
  const bezel = 0.1;
  box(s, x, y, w, h, '1A1A1A');
  photo(s, x + bezel, y + bezel, w - 2 * bezel, h - 2 * bezel);
  s.addShape('trapezoid', { x: x + w * 0.4, y: y + h, w: w * 0.2, h: 0.22,
    fill: { color: 'C9C9C9' } });
  s.addShape('roundRect', { x: x + w * 0.22, y: y + h + 0.2, w: w * 0.56, h: 0.1,
    rectRadius: 0.05, fill: { color: 'B0B0B0' } });
}

/* ------------------------------------------------------------------ *
 * Line-art icon set. Every icon is a list of primitives expressed in
 * fractions of its bounding box:
 *   [shapeName, x, y, w, h, mode, rotation]
 * mode: 'o' outline, 'w' knock-out in white, 'd' upper half-disc.
 * ------------------------------------------------------------------ */
const ICONS = {
  monitor: [['rect', 0, 0, .78, .7, 'o'], ['rect', .16, .16, .46, .3],
            ['rect', .3, .7, .18, .14], ['rect', .16, .84, .46, .12],
            ['roundRect', .84, .12, .16, .56, 'o']],
  doc: [['rect', .1, 0, .8, 1, 'o'], ['rect', .26, .18, .48, .1],
        ['rect', .26, .4, .48, .1], ['rect', .26, .62, .3, .1]],
  chart: [['rect', 0, 0, 1, 1, 'o'], ['rect', .18, .56, .14, .26],
          ['rect', .43, .34, .14, .48], ['rect', .68, .18, .14, .64]],
  user: [['ellipse', .32, .0, .36, .38], ['pie', .1, .38, .8, 1.1, 'd']],
  users: [['ellipse', .0, .12, .27, .28], ['pie', -.12, .38, .5, .8, 'd'],
          ['ellipse', .73, .12, .27, .28], ['pie', .62, .38, .5, .8, 'd'],
          ['ellipse', .36, .0, .3, .32], ['pie', .2, .3, .6, 1.0, 'd']],
  chat: [['roundRect', 0, 0, 1, .7], ['triangle', .16, .62, .24, .28, 'v'],
         ['ellipse', .17, .28, .14, .14, 'w'], ['ellipse', .43, .28, .14, .14, 'w'],
         ['ellipse', .69, .28, .14, .14, 'w']],
  tag: [['homePlate', .05, .1, .9, .8, null, 225], ['ellipse', .24, .28, .16, .16, 'w']],
  phone: [['roundRect', .3, .02, .4, .96, 'o'], ['rect', .42, .1, .16, .05],
          ['ellipse', .44, .82, .12, .12]],
  pencil: [['rect', .38, .06, .24, .72, null, 45], ['triangle', .06, .6, .32, .32, null, 225]],
  laptop: [['trapezoid', .02, .1, .96, .58], ['rect', 0, .72, 1, .12],
           ['rect', .2, .2, .6, .38, 'w']],
  cap: [['triangle', 0, .06, 1, .44], ['trapezoid', .2, .5, .6, .36, 'v'],
        ['rect', .9, .34, .04, .4]],
  lock: [['roundRect', .12, .4, .76, .6], ['blockArc', .24, .02, .52, .76, 'o']],
  gear: [['star8', 0, 0, 1, 1], ['ellipse', .32, .32, .36, .36, 'w']],
  clock: [['ellipse', 0, 0, 1, 1, 'o'], ['rect', .46, .2, .08, .34], ['rect', .5, .46, .26, .08]],
  book: [['rect', .04, .1, .42, .8, 'o'], ['rect', .54, .1, .42, .8, 'o']],
  trophy: [['trapezoid', .18, .05, .64, .48, 'v'], ['rect', .44, .5, .12, .26],
           ['rect', .24, .76, .52, .13],
           ['blockArc', -.06, .04, .3, .36, 'o'], ['blockArc', .76, .04, .3, .36, 'o']],
  medal: [['ellipse', .18, .42, .64, .58, 'o'], ['rect', .28, 0, .14, .44],
          ['rect', .58, 0, .14, .44], ['ellipse', .36, .56, .28, .28]],
  play: [['rect', 0, .18, 1, .64, 'o'], ['triangle', .38, .34, .32, .32, 'r']],
  bulb: [['ellipse', .14, 0, .72, .74], ['rect', .34, .72, .32, .14], ['rect', .38, .88, .24, .12]],
};

function icon(s, name, x, y, w, h, color, weight) {
  (ICONS[name] || []).forEach(function (p) {
    const o = { x: x + p[1] * w, y: y + p[2] * h, w: p[3] * w, h: p[4] * h };
    const mode = p[5];
    if (mode === 'o') { o.fill = { type: 'none' }; o.line = { color: color, width: weight || 1.5 }; }
    else if (mode === 'w') { o.fill = { color: WHITE }; }
    else { o.fill = { color: color }; }
    if (mode === 'v') o.rotate = 180;
    if (mode === 'r') o.rotate = 90;
    if (mode === 'd') o.angleRange = [180, 360];   // upper half of a disc
    if (p[6]) o.rotate = p[6];
    s.addShape(p[0], o);
  });
}

/* Circled social glyph (slides 18 and 35). */
function social(s, glyph, x, y, d, color) {
  outline(s, 'ellipse', x, y, d, d, color, 1.75);
  s.addText(glyph, { x, y, w: d, h: d, align: 'center', valign: 'middle', margin: 0,
    fontFace: HEAD, fontSize: 15, bold: true, color: color });
}

/* Icon + "Title Here" stack used on slides 10, 19, 20 and 21. */
function iconStack(s, name, ix, iy, iw, ih, tx, ty, color) {
  icon(s, name, ix, iy, iw, ih, color);
  txt(s, 'Title Here', { x: tx, y: ty, w: 1.654, h: 0.337, fontFace: HEAD,
    fontSize: 14, bold: true, color: color, align: 'center' });
}

/* Heading + paragraph pair, the deck's most repeated text block. */
function captionPair(s, head, bodyText, x, y, w, o) {
  o = o || {};
  txt(s, head, { x, y, w, h: o.hh || 0.303, fontFace: o.headFace || BODY,
    fontSize: o.headSize || 12, bold: true, color: o.headColor || RED, align: o.align || 'left' });
  txt(s, bodyText, { x, y: y + (o.gap || 0.237), w, h: o.bh || 0.656,
    fontSize: o.size || 11, color: o.color || GREY, align: o.align || 'left',
    lineSpacingMultiple: o.ls });
}

/* "Great Service" card: icon above a centred title and paragraph. */
function serviceCard(s, name, ix, iy, iw, ih, tx, ty, bx, by, bw, bh, str, color) {
  icon(s, name, ix, iy, iw, ih, color);
  txt(s, 'Great Service', { x: tx, y: ty, w: 1.821, h: 0.337, fontFace: HEAD,
    fontSize: 14, bold: true, color: color, align: 'center' });
  txt(s, str, { x: bx, y: by, w: bw, h: bh, color: color === WHITE ? WHITE : GREY,
    align: 'center', lineSpacingMultiple: 1.5 });
}

/* The crimson banner + white headline shared by the infographic slides. */
function infoHeader(s) {
  box(s, 0, 0, 13.333, 1.853, RED);
  txt(s, 'Infographic Of Cleva', { x: 2.566, y: 0.573, w: 8.202, h: 0.707,
    fontFace: HEAD, fontSize: 36, bold: true, color: WHITE, align: 'center' });
}

/* Dotted leader line used by the infographic call-outs. */
function leader(s, x, y, w, h, o) {
  s.addShape('line', Object.assign({ x, y, w, h,
    line: { color: DARK, width: 2, dashType: 'sysDot' } }, o));
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */
const slides = [];

/* 1 - cover */
slides.push(function (s) {
  photo(s, 0, 0, 13.333, 7.5);
  box(s, 3.323, 2.841, 7.129, 2.237, BLACK, { fill: { color: BLACK, transparency: 64 } });
  box(s, 3.323, 5.078, 7.129, 0.127, RED);
  txt(s, 'ACADEMIZE', { x: 3.578, y: 3.291, w: 6.708, h: 1.313, fontFace: HEAD,
    fontSize: 72, bold: true, color: RED, align: 'center' });
  txt(s, 'BEST UNIVERSITY & COLLEGE', { x: 4.888, y: 4.437, w: 4.088, h: 0.404,
    fontSize: 18, color: WHITE, align: 'center' });
  outline(s, 'rect', 0.458, 0.417, 12.417, 6.667, RED, 2.25);
});

/* 2 - four feature cards */
slides.push(function (s) {
  box(s, 9.115, 0, 4.218, 3.001, RED);
  const cards = [
    [0.788, 0.765, 'monitor', 'ONLINE\nEDUCATION'],
    [3.967, 0.765, 'doc', 'CLUBS &\nSTUDYING'],
    [0.788, 4.034, 'chart', 'INTERNATIONAL\nSTUDENTS'],
    [3.967, 4.034, 'user', 'PROGRESSIVE\nPROGRAM'],
  ];
  cards.forEach(function (c) {
    const x = c[0], y = c[1];
    box(s, x, y, 2.568, 0.93, RED);
    box(s, x, y + 0.93, 2.568, 1.772, GREY_CARD);
    box(s, x + 0.095, y + 0.081, 0.846, 0.767, GREY_BOX);
    icon(s, c[2], x + 0.27, y + 0.24, 0.5, 0.45, WHITE);
    txt(s, c[3], { x: x + 1.02, y: y + 0.212, w: 1.5, h: 0.505,
      fontFace: HEAD, fontSize: 12, color: WHITE });
    txt(s, L_SHORT, { x: x + 0.2, y: y + 1.085, w: 2.167, h: 1.336,
      fontSize: 10, lineSpacingMultiple: 1.5 });
  });
  title(s, 'About Our Content In This Template', 7.558, 4.882, 4.892, 1.313);
  photo(s, 7.558, 0.757, 4.286, 3.367);
});

/* 3 - welcome */
slides.push(function (s) {
  photo(s, 0, 0, 6.206, 7.5);
  box(s, 1.277, 0.27, 3.255, 2.775, RED, { fill: { color: RED, transparency: 38 } });
  txt(s, UP(L_SHORT), { x: 1.749, y: 0.789, w: 2.312, h: 1.737, color: WHITE,
    lineSpacingMultiple: 1.5 });
  box(s, 6.206, 0, 0.46, 1.384, RED);
  title(s, 'Welcome to Academize University', 7.158, 1.062, 5.415, 1.313);
  eyebrow(s, 'INTRODUCTION', 7.188, 2.759, 1.58, 0.286);
  txt(s, L_FULL, { x: 7.188, y: 3.093, w: 5.46, h: 1.489, lineSpacingMultiple: 1.5 });
  eyebrow(s, 'OUR SPECIALIZATION', 7.188, 4.919, 2.083, 0.286);
  txt(s, 'I should be incapable of drawing a single stroke at the present moment; and yet I feel that I never was a greater artist than now. . I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine.',
    { x: 7.188, y: 5.256, w: 5.46, h: 1.182, lineSpacingMultiple: 1.5 });
});

/* 4 - founder message */
slides.push(function (s) {
  photo(s, 7.921, 0, 5.413, 7.5);
  title(s, 'This Is Welcome Message From Our Founder', 0.736, 0.714, 6.278, 1.313);
  box(s, 4.997, 3.381, 5.171, 3.382, RED);
  txt(s, UP(L_FULL), { x: 5.417, y: 3.925, w: 4.332, h: 2.293, color: WHITE,
    align: 'justify', lineSpacingMultiple: 1.5 });
  s.addShape('line', { x: 4.997, y: 2.42, w: 1.67, h: 0,
    line: { color: RED, width: 1.5, dashType: 'dash' } });
  photo(s, 0.794, 2.345, 2.841, 2.841, 'ellipse');
  eyebrow(s, 'MESSAGE FROM CEO', 0.736, 5.561, 2.957, 0.286, { align: 'center' });
  txt(s, '\u201CA brand for a company is like a reputation for a person. You earn reputation by trying to do hard things well\u201D',
    { x: 0.568, y: 5.912, w: 3.292, h: 0.904, color: GREY_D, align: 'center',
      lineSpacingMultiple: 1.5 });
});

/* 5 - banner + centred intro */
slides.push(function (s) {
  box(s, 0, 2.694, 13.333, 1.389, RED);
  txt(s, 'About Our University', { x: 3.826, y: 4.715, w: 5.681, h: 0.64,
    fontFace: HEAD, fontSize: 32, bold: true, color: RED, align: 'center' });
  txt(s, L_FULL, { x: 1.288, y: 5.542, w: 10.758, h: 0.904, align: 'center',
    lineSpacingMultiple: 1.5 });
  photo(s, 0, 0, 13.333, 3.75);
});

/* 6 - story */
slides.push(function (s) {
  title(s, 'Story About How We Start The Collage', 7.433, 3.005, 5.021, 1.313);
  box(s, 0, 2.292, 6.667, 1.458, RED);
  icon(s, 'monitor', 0.44, 2.826, 0.427, 0.311, WHITE);
  txt(s, 'ONLINE\nEDUCATION', { x: 1.152, y: 2.729, w: 1.426, h: 0.505,
    fontSize: 12, color: WHITE });
  txt(s, UP(L_SHORT), { x: 2.663, y: 2.605, w: 3.584, h: 0.831, fontSize: 10,
    color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, L_FULL, { x: 7.433, y: 5.307, w: 5.345, h: 1.489, lineSpacingMultiple: 1.5 });
  eyebrow(s, 'UNIVERSITY & COLLAGE', 7.433, 4.972, 2.283, 0.286);
  box(s, 10.984, 0, 2.349, 0.254, RED);
  photo(s, 0, 3.75, 6.667, 3.75);
});

/* 7 - about whom we are */
slides.push(function (s) {
  photo(s, 3.708, 3.75, 9.625, 3.75);
  box(s, 0, 3.75, 6.667, 2.993, RED, { fill: { color: RED, transparency: 20 } });
  txt(s, UP(L_FULL), { x: 0.75, y: 4.239, w: 5.167, h: 2.015, color: WHITE,
    align: 'justify', lineSpacingMultiple: 1.5 });
  title(s, 'About Whom We Are', 0.75, 1.684, 5.008, 0.707);
  txt(s, L_MED + '.', { x: 6.257, y: 1.585, w: 6.478, h: 0.904, lineSpacingMultiple: 1.5 });
});

/* 8 - vision & mission */
slides.push(function (s) {
  box(s, 0, 0, 7.165, 2.708, RED);
  photo(s, 0.76, 0.757, 9.385, 3.326);
  box(s, 9.375, 1.832, 3.198, 2.938, RED, { fill: { color: RED, transparency: 28 } });
  txt(s, UP(L_SHORT) + ' I AM ALONE, ', { x: 9.592, y: 2.166, w: 2.764, h: 2.269,
    color: WHITE, align: 'justify', lineSpacingMultiple: 2 });
  title(s, 'About Vision & Mision Of Academize', 0.681, 5.345, 5.029, 1.313);
  txt(s, L_SHORT + ' ', { x: 6.437, y: 5.688, w: 5.198, h: 0.627, color: BLACK,
    lineSpacingMultiple: 1.5 });
});

/* 9 - we don't stop going to school */
slides.push(function (s) {
  photo(s, 6.667, 2.765, 6.667, 4.735);
  txt(s, 'We Don\u2019t Stop Going To School When We Graduate', { x: 0.735, y: 0.968,
    w: 5.612, h: 2.121, fontFace: HEAD, fontSize: 40, bold: true, color: RED });
  box(s, 4.397, 3.75, 3.198, 2.993, RED, { fill: { color: RED, transparency: 22 } });
  txt(s, UP(L_MED), { x: 4.614, y: 3.961, w: 2.764, h: 2.57, color: WHITE,
    align: 'justify', lineSpacingMultiple: 1.5 });
  txt(s, L_FULL, { x: 0.76, y: 3.822, w: 3.048, h: 2.848, lineSpacingMultiple: 1.5 });
  box(s, 12.598, 0, 0.735, 2.765, RED);
  txt(s, 'A WONDERFUL SERENITY HAS TAKEN POSSESSION OF MY ENTIRE SOUL, LIKE THESE SWEET MORNINGS',
    { x: 7.112, y: 1.381, w: 5.167, h: 0.627, color: RED, align: 'center',
      lineSpacingMultiple: 1.5 });
});

/* 10 - don't stop to studying */
slides.push(function (s) {
  box(s, 0, 5.103, 13.333, 2.397, RED);
  box(s, 6.309, 0, 1.302, 4.042, RED);
  title(s, 'Don\u2019t Stop To Studying When You Graduate', 1.0, 1.146, 4.349, 1.919);
  txt(s, L_MED, { x: 8.923, y: 1.062, w: 3.663, h: 1.886, fontSize: 12,
    lineSpacingMultiple: 1.5 });
  iconStack(s, 'users', 6.725, 0.72, 0.471, 0.428, 6.134, 1.279, WHITE);
  photo(s, 0.322, 4.042, 12.69, 3.167);
  iconStack(s, 'chat', 6.725, 2.426, 0.471, 0.428, 6.134, 2.985, WHITE);
  box(s, 9.038, 3.46, 3.433, 1.351, RED, { fill: { color: RED, transparency: 46 } });
  txt(s, 'A WONDERFUL SERENITY HAS TAKEN POSSESSION OF MY ENTIRE SOUL, LIKE THESE SWEET MORNINGS OF SPRING',
    { x: 9.194, y: 3.72, w: 3.163, h: 0.831, fontSize: 10, color: WHITE,
      align: 'justify', lineSpacingMultiple: 1.5 });
});

/* 11 - section break */
slides.push(function (s) {
  photo(s, 0, 0, 13.333, 7.5);
  box(s, 0.987, 0.723, 11.617, 6.054, RED, { fill: { color: RED, transparency: 67 } });
  outline(s, 'rect', 1.596, 1.382, 10.142, 4.735, WHITE, 2.25);
  txt(s, 'We All Take Break Here In This Slide ', { x: 3.485, y: 2.334, w: 6.363,
    h: 1.447, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE, align: 'center' });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which i enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. ',
    { x: 2.854, y: 4.262, w: 7.625, h: 0.904, color: WHITE, align: 'center',
      lineSpacingMultiple: 1.5 });
});

/* 12 - our great services (2 x 2) */
slides.push(function (s) {
  box(s, 9.346, 0, 3.988, 7.5, RED);
  const cards = [
    ['doc', 2.579, 2.309, 0.53, 0.53, 1.934, 3.126, 1.164, 3.528],
    ['laptop', 6.609, 2.305, 0.595, 0.541, 5.996, 3.126, 5.226, 3.528],
    ['pencil', 2.579, 4.714, 0.53, 0.53, 1.934, 5.53, 1.164, 5.933],
    ['cap', 6.539, 4.71, 0.734, 0.534, 5.996, 5.53, 5.226, 5.933],
  ];
  cards.forEach(function (c) {
    serviceCard(s, c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], 3.36, 0.627, L_TINY + ' ', RED);
  });
  txt(s, 'Our Great Services', { x: 0.433, y: 0.94, w: 4.85, h: 0.707,
    fontFace: HEAD, fontSize: 36, bold: true, color: RED, align: 'center' });
  photo(s, 10.127, 0.783, 3.206, 5.934);
});

/* 13 - services on crimson */
slides.push(function (s) {
  box(s, 0, 1.024, 12.135, 5.451, RED);
  [['users', 1.833, 1.713], ['chat', 3.358, 3.237], ['tag', 4.882, 4.757]].forEach(function (r) {
    outline(s, 'rect', 7.453, r[1], 0.79, 0.79, WHITE, 2.25);
    icon(s, r[0], 7.63, r[1] + 0.19, 0.43, 0.41, WHITE);
    txt(s, 'Great Service', { x: 8.638, y: r[2], w: 1.654, h: 0.337,
      fontFace: HEAD, fontSize: 14, bold: true, color: WHITE });
    txt(s, 'a ' + L_TINY.slice(2) + ' ', { x: 8.638, y: r[2] + 0.356, w: 2.575, h: 0.626,
      color: WHITE, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  photo(s, 1.198, 0, 5.333, 7.5);
});

/* 14 - awesome about our services (3 up) */
slides.push(function (s) {
  box(s, 0, 5.905, 13.333, 1.595, RED);
  txt(s, UP(L_MED), { x: 0.965, y: 6.389, w: 11.403, h: 0.626, color: WHITE,
    align: 'center', lineSpacingMultiple: 1.5 });
  txt(s, 'Awesome About Our Services', { x: 2.566, y: 1.055, w: 8.202, h: 0.707,
    fontFace: HEAD, fontSize: 36, bold: true, color: RED, align: 'center' });
  const copy = 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy';
  [['doc', 2.339, 2.694, 0.53, 0.53, 1.694, 3.51, 0.924, 3.913],
   ['laptop', 6.369, 2.69, 0.595, 0.541, 5.756, 3.51, 4.986, 3.913],
   ['pencil', 10.464, 2.69, 0.53, 0.53, 9.819, 3.506, 9.049, 3.909]].forEach(function (c) {
    serviceCard(s, c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], 3.36, 0.904, copy, RED);
  });
});

/* 15 - crimson column */
slides.push(function (s) {
  box(s, 4.307, 0.7, 4.719, 6.099, RED);
  [['doc', 6.402, 1.432, 0.53, 0.53, 5.756, 2.248, 4.986, 2.651],
   ['laptop', 6.369, 4.17, 0.595, 0.541, 5.756, 4.991, 4.986, 5.393]].forEach(function (c) {
    serviceCard(s, c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], 3.36, 0.626, 'a ' + L_TINY.slice(2) + ' ', WHITE);
  });
  const side = 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings';
  txt(s, side, { x: 9.85, y: 0.744, w: 2.66, h: 0.904, lineSpacingMultiple: 1.5 });
  txt(s, side, { x: 9.85, y: 5.857, w: 2.66, h: 0.904, lineSpacingMultiple: 1.5 });
  title(s, 'Awesome About Our Services', 0.887, 2.791, 2.77, 1.919);
  photo(s, 9.026, 2.137, 4.307, 3.207);
});

/* 16 - lecturer grid: [photo x, photo y, panel x, panel y, text x, text y, align] */
slides.push(function (s) {
  const cells = [
    [0.521, 2.604, 3.396, 2.604, 3.875, 3.116, 'left'],
    [9.938, 2.604, 6.833, 2.604, 7.146, 3.116, 'right'],
    [3.625, 4.969, 0.521, 4.969, 0.917, 5.481, 'right'],
    [6.833, 4.969, 9.708, 4.969, 10.187, 5.481, 'left'],
  ];
  cells.forEach(function (c) { box(s, c[2], c[3], 3.104, 2.052, RED); });
  cells.forEach(function (c) {
    const right = c[6] === 'right';
    txt(s, 'Awesome Team', { x: right ? c[4] + 0.659 : c[4], y: c[5], w: 1.654, h: 0.337,
      fontFace: HEAD, fontSize: 14, bold: true, color: WHITE, align: c[6] });
    txt(s, 'a wonderful serenity has taken possession', { x: c[4], y: c[5] + 0.354,
      w: 2.313, h: 0.626, color: WHITE, align: c[6], lineSpacingMultiple: 1.5 });
  });
  txt(s, 'Awesome About Our Lecturer', { x: 2.566, y: 1.111, w: 8.202, h: 0.707,
    fontFace: HEAD, fontSize: 36, bold: true, color: RED, align: 'center' });
  cells.forEach(function (c) { photo(s, c[0], c[1], 2.875, 2.052); });
});

/* 17 - professors */
slides.push(function (s) {
  box(s, 5.048, 0, 8.286, 7.5, RED);
  const copy = 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which i enjoy with my whole heart';
  [[0.766, 1.122, 11], [3.083, 3.439, 12], [5.437, 5.793, 12]].forEach(function (r) {
    txt(s, 'Awesome Team', { x: 8.773, y: r[0], w: 1.654, h: 0.337, fontFace: HEAD,
      fontSize: 14, bold: true, color: WHITE });
    txt(s, copy, { x: 8.773, y: r[1], w: 3.8, h: 0.978, fontSize: r[2],
      color: WHITE, lineSpacingMultiple: 1.5 });
  });
  title(s, 'Our Awesome Professor', 0.687, 2.791, 2.495, 1.919);
  [1.396, 3.75, 6.104].forEach(function (y) {
    s.addShape('line', { x: 7.492, y: y, w: 1.057, h: 0, line: { color: WHITE, width: 1.5 } });
  });
  [0.352, 2.706, 5.06].forEach(function (y) { photo(s, 3.792, y, 4.229, 2.089); });
});

/* 18 - personal profile with skill bars */
slides.push(function (s) {
  box(s, 0, 0, 8.889, 1.23, RED);
  ['t', 'f', 'P'].forEach(function (g, i) {
    social(s, g, 7.014, 3.0 + i * 1.268, 0.527, RED);
  });
  [[5.658, 3.477, '80%', 'Teamwork'], [6.387, 4.101, '94%', 'Working Times']].forEach(function (b) {
    s.addShape('roundRect', { x: 8.081, y: b[0], w: 4.492, h: 0.317, rectRadius: 0.158,
      fill: { color: GREY_TRACK } });
    s.addShape('roundRect', { x: 8.081, y: b[0], w: b[1], h: 0.317, rectRadius: 0.158,
      fill: { color: RED } });
    s.addText(b[2], { x: 8.081, y: b[0], w: b[1] - 0.1, h: 0.317, align: 'right',
      valign: 'middle', margin: 0, fontFace: HEAD, fontSize: 12, bold: true, color: WHITE });
    s.addText(b[3], { x: 8.17, y: b[0], w: 1.9, h: 0.317, align: 'left', valign: 'middle',
      margin: 0, fontFace: HEAD, fontSize: 12, bold: true, color: WHITE });
  });
  txt(s, L_MED, { x: 8.081, y: 3.458, w: 4.74, h: 1.182, lineSpacingMultiple: 1.5 });
  title(s, 'About Personal Team Profile', 8.081, 1.841, 4.423, 1.313);
  photo(s, 1.198, 0, 5.333, 7.5);
});

/* 19 - portfolio with two vertical tabs */
slides.push(function (s) {
  box(s, 0, 2.925, 11.917, 1.649, RED);
  [1.708, 5.042].forEach(function (y) {
    s.addShape('rect', { x: 10.76, y: y, w: 3.062, h: 0.75, fill: { color: RED }, rotate: 270 });
    s.addText('Great Portfolio', { x: 10.76, y: y, w: 3.062, h: 0.75, rotate: 270,
      align: 'center', valign: 'middle', margin: 0, fontFace: HEAD, fontSize: 16,
      bold: true, color: WHITE });
  });
  title(s, 'Awesome Portfolio In Academize Presentation', 0.798, 0.949, 5.869, 1.313);
  txt(s, L_MED, { x: 0.798, y: 5.344, w: 5.265, h: 1.182, lineSpacingMultiple: 1.5 });
  [['users', 1.276, 0.684], ['chat', 3.328, 2.737], ['tag', 5.402, 4.789]].forEach(function (g) {
    iconStack(s, g[0], g[1], 3.302, 0.471, 0.428, g[2], 3.861, WHITE);
  });
  photo(s, 7.127, 0.552, 4.79, 3.062);
  photo(s, 7.127, 3.885, 4.79, 3.062);
});

/* 20 - portfolio, right rail */
slides.push(function (s) {
  box(s, 0, 0, 8.889, 0.397, RED);
  txt(s, 'Awesome Portfolio In Academize Presentation', { x: 0.5, y: 1.25, w: 6.403,
    h: 1.313, fontSize: 36, bold: true, color: RED });
  txt(s, L_MED, { x: 4.747, y: 5.034, w: 5.348, h: 1.182, lineSpacingMultiple: 1.5 });
  box(s, 11.429, 3.75, 1.905, 3.75, RED);
  iconStack(s, 'users', 12.146, 4.42, 0.471, 0.428, 11.554, 4.979, WHITE);
  iconStack(s, 'chat', 12.146, 5.935, 0.471, 0.428, 11.554, 6.494, WHITE);
  photo(s, 0, 3.75, 3.857, 3.75);
  photo(s, 7.19, 0, 6.143, 3.75);
});

/* 21 - portfolio, left rail */
slides.push(function (s) {
  box(s, 9.429, 3.46, 3.905, 4.04, RED);
  txt(s, 'Awesome Portfolio In Academize Presentation', { x: 6.825, y: 0.916, w: 6.379,
    h: 1.313, fontSize: 36, bold: true, color: RED });
  box(s, 0, 0, 1.905, 7.5, RED);
  txt(s, L_MED, { x: 3.048, y: 4.969, w: 5.413, h: 1.182, lineSpacingMultiple: 1.5 });
  [['users', 0.72, 0.733], ['chat', 0.72, 2.446], ['tag', 0.741, 4.159],
   ['phone', 0.741, 5.872]].forEach(function (g) {
    iconStack(s, g[0], g[1], g[2], 0.45, 0.428, 0.129, g[2] + 0.559, WHITE);
  });
  photo(s, 1.905, 0, 4.603, 3.75);
  photo(s, 9.726, 3.75, 3.608, 3.75);
});

/* 22 - portfolio, vertical title */
slides.push(function (s) {
  box(s, 0, 0, 0.76, 7.5, RED);
  [0.496, 2.75, 5.004].forEach(function (y) {
    box(s, 9.603, y, 3.413, 2.0, RED);
    captionPair(s, 'Great Portfolio', L_TINY, 10.162, y + 0.325, 2.396,
      { headFace: HEAD, headSize: 14, headColor: WHITE, color: WHITE, size: 12,
        gap: 0.341, bh: 0.978, ls: 1.5 });
    photo(s, 7.27, y, 2.333, 2.0);
  });
  s.addText('Awseome Portfolio', { x: -0.538, y: 3.397, w: 4.553, h: 0.707, rotate: 90,
    align: 'center', valign: 'top', margin: 0, fontFace: HEAD, fontSize: 36,
    bold: true, color: RED });
  photo(s, 2.651, 0, 4.349, 7.5);
});

/* 23 - two mirrored portfolio cards */
slides.push(function (s) {
  txt(s, 'Awesome Portfolio In Academize Presentation', { x: 2.566, y: 0.871, w: 8.202,
    h: 1.313, fontSize: 36, bold: true, color: RED, align: 'center' });
  box(s, 9.448, 2.881, 3.413, 2.479, RED);
  captionPair(s, 'Great Portfolio', L_TINY, 10.007, 3.385, 2.396,
    { headFace: HEAD, headSize: 14, headColor: WHITE, color: WHITE, size: 12,
      gap: 0.461, bh: 1.01, ls: 1.5 });
  box(s, 0.473, 2.881, 3.413, 2.479, RED);
  captionPair(s, 'Great Portfolio', L_TINY, 0.93, 3.385, 2.396,
    { headFace: HEAD, headSize: 14, headColor: WHITE, color: WHITE, size: 12,
      gap: 0.461, bh: 1.01, ls: 1.5, align: 'right' });
  txt(s, L_MED, { x: 1.246, y: 6.057, w: 10.841, h: 0.627, align: 'center',
    lineSpacingMultiple: 1.5 });
  photo(s, 3.885, 2.881, 2.604, 2.479);
  photo(s, 6.844, 2.881, 2.604, 2.479);
});

/* 24 - four captioned tiles */
slides.push(function (s) {
  box(s, 0, 6.0, 13.333, 1.5, RED);
  [0.625, 3.69, 6.754, 9.819].forEach(function (x) {
    box(s, x, 3.75, 2.889, 0.646, RED);
    s.addText('Great Portfolio', { x: x, y: 3.75, w: 2.889, h: 0.646, align: 'center',
      valign: 'middle', margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: WHITE });
    photo(s, x, 4.396, 2.889, 2.5);
  });
  txt(s, 'Awesome Portfolio In Academize Presentation', { x: 2.566, y: 0.871, w: 8.202,
    h: 1.313, fontSize: 36, bold: true, color: RED, align: 'center' });
  txt(s, L_MED, { x: 1.246, y: 2.519, w: 10.841, h: 0.627, align: 'center',
    lineSpacingMultiple: 1.5 });
});

/* 25 - five phone mock-ups */
slides.push(function (s) {
  box(s, 0, 5.229, 13.333, 2.271, RED);
  txt(s, 'Awesome Portfolio In Academize Presentation', { x: 2.566, y: 0.636, w: 8.202,
    h: 1.313, fontSize: 36, bold: true, color: RED, align: 'center' });
  txt(s, L_MED, { x: 1.246, y: 2.284, w: 10.841, h: 0.627, align: 'center',
    lineSpacingMultiple: 1.5 });
  [1.038, 3.393, 5.747, 8.101, 10.455].forEach(function (x) {
    phoneMock(s, x, 3.34, 1.84, 3.712);
  });
});

/* 26 - desktop mock-ups */
slides.push(function (s) {
  box(s, 0, 0, 3.439, 7.5, RED);
  monitorMock(s, 0.76, 2.15, 4.85, 3.05);
  monitorMock(s, 4.58, 3.54, 3.45, 2.0);
  title(s, 'About Mockup', 8.416, 3.293, 3.727, 0.707);
  icon(s, 'lock', 8.99, 4.5, 0.4, 0.5, RED);
  icon(s, 'laptop', 11.55, 4.55, 0.5, 0.36, RED);
  [8.58, 11.094].forEach(function (x) {
    txt(s, 'NULLAM EU TEMPOR PURUS. NUNC A LEO MAGNA, SIT AMET.', { x: x, y: 5.116,
      w: 1.856, h: 1.063, fontSize: 10, color: RED, lineSpacingMultiple: 2 });
  });
  txt(s, L_MED, { x: 6.083, y: 1.922, w: 6.49, h: 0.904, lineSpacingMultiple: 1.5 });
});

/* Shared chrome for the two chart slides. */
function chartChrome(s, tx) {
  box(s, 0, 4.952, 13.333, 2.548, RED);
  txt(s, UP(L_SHORT) + ' ', { x: 2.991, y: 5.913, w: 7.652, h: 0.627, color: WHITE,
    align: 'center', lineSpacingMultiple: 1.5 });
  title(s, 'Chart Graphic', tx, tx === 1.068 ? 1.176 : 1.354, 3.8, 0.707);
  txt(s, 'University & Collage', { x: tx, y: tx === 1.068 ? 2.487 : 2.577, w: 2.361,
    h: 0.337, fontFace: HEAD, fontSize: 14, bold: true, color: RED });
  txt(s, L_MED + ' ', { x: tx, y: tx === 1.068 ? 2.827 : 2.917, w: 4.697, h: 1.182,
    lineSpacingMultiple: 1.5 });
}

/* Axis grid + tick labels shared by both charts. */
function chartGrid(s, gx, gw, y0, dy, labels, lx, lw, cats, cx0, cdx, cy) {
  labels.forEach(function (t, i) {
    s.addShape('line', { x: gx, y: y0 + i * dy, w: gw, h: 0,
      line: { color: GREY_RULE, width: 0.75 } });
    s.addText(t, { x: lx, y: y0 + i * dy - 0.17, w: lw, h: 0.35, align: 'right',
      margin: 0, valign: 'top', fontFace: BODY, fontSize: 12, color: BLACK });
  });
  cats.forEach(function (t, i) {
    s.addText(t, { x: cx0 + i * cdx, y: cy, w: 0.95, h: 0.35, align: 'center',
      margin: 0, valign: 'top', fontFace: BODY, fontSize: 12, color: BLACK });
  });
}

/* 27 - stacked column chart */
slides.push(function (s) {
  chartChrome(s, 1.068);
  chartGrid(s, 7.123, 5.45, 0.932, 0.39, ['14', '12', '10', '8', '6', '4', '2', '0'],
    6.667, 0.378, ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
    7.423, 1.316, 3.873);
  /* [x, red-top, red-h, lime-top, lime-h, orange-top, orange-h] */
  const cols = [
    [7.614, 2.372, 1.292, 1.982, 0.451, 1.591, 0.39],
    [8.93, 2.839, 0.824, 2.109, 0.772, 1.756, 0.39],
    [10.245, 2.614, 1.05, 2.372, 0.266, 1.969, 0.401],
    [11.561, 2.372, 1.292, 1.982, 0.396, 1.322, 0.708],
  ];
  cols.forEach(function (c) {
    box(s, c[0], c[1], 0.564, c[2], RED);
    box(s, c[0], c[3], 0.564, c[4], LIME);
    box(s, c[0], c[5], 0.564, c[6], ORANGE);
  });
});

/* 28 - candlestick chart */
slides.push(function (s) {
  chartChrome(s, 7.45);
  chartGrid(s, 1.368, 5.042, 0.927, 0.379, ['70', '60', '50', '40', '30', '20', '10', '0'],
    0.946, 0.35, ['01/01/2023', '01/02/2023', '01/03/2023', '01/04/2023'],
    1.606, 1.217, 3.781);
  /* [wick-x, wick-y, wick-h, body-x, body-y, body-h, colour] */
  const candles = [
    [2.083, 1.484, 1.715, 1.823, 1.873, 0.816, ORANGE],
    [3.300, 1.482, 1.715, 3.039, 2.180, 0.538, RED],
    [4.511, 1.428, 1.715, 4.256, 1.702, 0.463, ORANGE],
    [5.734, 1.352, 1.792, 5.473, 1.684, 0.655, RED],
  ];
  candles.forEach(function (c) {
    s.addShape('line', { x: c[0], y: c[1], w: 0, h: c[2], line: { color: BLACK, width: 0.75 } });
    box(s, c[3], c[4], 0.522, c[5], c[6]);
  });
});

/* 29 - graduation cap infographic */
slides.push(function (s) {
  infoHeader(s);
  /* book spines behind the cap */
  [[5.909, 4.549, 0.168, 1.566, LIME], [6.160, 4.509, 0.173, 1.610, ORANGE],
   [7.792, 4.663, 0.172, 1.444, LIME], [6.641, 4.543, 0.169, 1.573, RED],
   [6.411, 4.712, 0.150, 1.392, RED], [6.892, 4.465, 0.178, 1.657, LIME],
   [5.640, 4.382, 0.187, 1.745, ORANGE], [5.409, 4.704, 0.150, 1.401, RED],
   [7.542, 4.686, 0.169, 1.420, RED], [7.161, 4.641, 0.290, 1.468, ORANGE],
  ].forEach(function (b) { box(s, b[0], b[1], b[2], b[3], b[4]); });
  /* mortarboard: crown block, dark brim band, then the diamond top */
  box(s, 5.33, 3.028, 2.689, 1.825, DARK);
  box(s, 5.33, 3.383, 2.689, 0.878, BLACK);
  s.addShape('diamond', { x: 3.526, y: 2.373, w: 6.282, h: 1.606, fill: { color: DARK } });
  s.addShape('line', { x: 6.772, y: 2.918, w: 1.641, h: 0.552, line: { color: GREEN, width: 1.5 } });
  s.addShape('line', { x: 8.413, y: 3.47, w: 0, h: 0.9, line: { color: ORANGE, width: 1.5 } });
  s.addShape('triangle', { x: 8.175, y: 4.013, w: 0.477, h: 0.827, fill: { color: RED }, rotate: 180 });
  s.addShape('ellipse', { x: 8.291, y: 3.95, w: 0.244, h: 0.244, fill: { color: RED } });
  /* six call-outs */
  [[2.246, 5.948, 'right'], [1.598, 5.083, 'right'], [1.878, 4.184, 'right'],
   [8.808, 5.948, 'left'], [9.456, 5.083, 'left'], [9.176, 4.184, 'left'],
  ].forEach(function (c) {
    captionPair(s, 'Content  Here', L_TINY, c[0], c[1], 2.279,
      { headFace: HEAD, align: c[2] });
  });
  leader(s, 4.158, 4.749, 1.172, 0.602);
  leader(s, 3.877, 5.642, 1.865, 0.006, { flipV: true });
  leader(s, 4.526, 5.878, 1.867, 0.635, { flipV: true });
  leader(s, 7.687, 5.647, 1.769, 0.449, { flipH: true });
  leader(s, 8.100, 4.749, 1.076, 0.587, { flipH: true });
  leader(s, 7.393, 5.938, 1.414, 0.575);
});

/* 30 - three thinking heads */
slides.push(function (s) {
  infoHeader(s);
  /* halo of tiny study icons behind each head */
  const halo = [[.04, .04], [.20, -.06], [.40, .00], [.62, -.08], [.80, .04], [.94, .18],
                [.00, .26], [.96, .40], [.06, .52], [.90, .58], [.16, .74], [.82, .74],
                [.30, .16], [.72, .14], [.12, .38], [.88, .30], [.34, -.14], [.56, .12],
                [.02, .66], [.94, .68], [.26, .90], [.70, .90], [.46, -.20], [.50, .84]];
  [[1.42, 2.35], [5.05, 2.35], [8.30, 2.35]].forEach(function (o, k) {
    halo.forEach(function (c, i) {
      const tint = [GREEN, GREEN_M, GREEN_L][(i + k) % 3];
      s.addShape(i % 3 === 0 ? 'ellipse' : 'roundRect',
        { x: o[0] + c[0] * 3.5, y: o[1] + c[1] * 2.6, w: 0.2, h: 0.2, fill: { color: tint } });
    });
  });
  /* profile heads: skull, nose wedge and neck */
  [[1.94, '?', GREY_M], [5.13, '!', GREEN], [8.34, '', GREEN]].forEach(function (h) {
    const L = h[0];
    s.addShape('ellipse', { x: L + 0.22, y: 3.15, w: 1.93, h: 2.1, fill: { color: RED } });
    s.addShape('triangle', { x: L - 0.15, y: 4.08, w: 0.72, h: 0.62,
      fill: { color: RED }, rotate: 270 });
    box(s, L + 0.91, 5.1, 1.05, 0.56, RED);
    if (h[1]) {
      s.addText(h[1], { x: L + 0.5, y: 3.3, w: 1.3, h: 1.5, align: 'center',
        valign: 'middle', margin: 0, fontFace: HEAD, fontSize: 60, bold: true, color: h[2] });
    }
  });
  icon(s, 'bulb', 9.15, 3.5, 1.0, 1.3, GREEN);
  /* time-line rule with three markers */
  s.addShape('line', { x: 1.06, y: 5.63, w: 10.9, h: 0, line: { color: ORANGE, width: 2.25 } });
  [[3.11, GREEN_L, 'book'], [6.31, GREEN, 'pencil'], [9.51, GREEN_M, 'gear']].forEach(function (m) {
    s.addShape('ellipse', { x: m[0], y: 5.32, w: 0.6, h: 0.6, fill: { color: m[1] } });
    icon(s, m[2], m[0] + 0.15, 5.46, 0.32, 0.34, WHITE);
  });
  [2.001, 5.243, 8.451].forEach(function (x) {
    captionPair(s, 'Your Text Here', L_TINY, x, 6.066, 2.79,
      { align: 'center', bh: 0.411, gap: 0.265 });
  });
});

/* 31 - layered graduate silhouette */
slides.push(function (s) {
  infoHeader(s);
  s.addShape('trapezoid', { x: 4.96, y: 5.25, w: 1.96, h: 1.49, fill: { color: GREY_M } });
  s.addShape('trapezoid', { x: 6.38, y: 4.69, w: 1.75, h: 2.05, fill: { color: MAGENTA } });
  s.addShape('trapezoid', { x: 5.30, y: 4.05, w: 2.6, h: 1.3, fill: { color: RED } });
  s.addShape('ellipse', { x: 6.23, y: 3.05, w: 1.18, h: 1.4, fill: { color: ORANGE } });
  /* mortarboard: crown, brim and tassel */
  box(s, 6.36, 3.02, 0.95, 0.35, RED);
  s.addShape('diamond', { x: 5.90, y: 2.24, w: 1.75, h: 0.86, fill: { color: RED } });
  s.addShape('line', { x: 7.52, y: 2.62, w: 0, h: 0.6, line: { color: RED, width: 1.5 } });
  s.addShape('triangle', { x: 7.40, y: 3.16, w: 0.24, h: 0.42, fill: { color: RED }, rotate: 180 });
  [['01', 6.612, 2.791], ['02', 6.684, 3.575], ['03', 6.318, 4.736],
   ['04', 6.997, 5.635], ['05', 5.748, 5.817]].forEach(function (n) {
    s.addText(n[0], { x: n[1], y: n[2], w: 0.405, h: 0.364, align: 'right', margin: 0,
      valign: 'top', fontFace: BODY, fontSize: 20, bold: true, color: WHITE });
  });
  [[1.722, 2.333, 'right'], [1.720, 3.899, 'right'], [1.720, 5.465, 'right'],
   [8.738, 3.116, 'left'], [8.738, 4.682, 'left']].forEach(function (c) {
    captionPair(s, 'Your Text  Here', L_TINY, c[0], c[1], 2.549,
      { align: c[2], headColor: DARK, hh: 0.252, gap: 0.197, bh: 0.471 });
  });
  [[4.27, 2.46, 2.23, 0.59], [4.27, 4.02, 2.24, 0.70], [4.27, 5.59, 0.95, 0.43],
   [7.21, 3.24, 1.53, 0.61], [7.61, 4.81, 1.12, 1.01]].forEach(function (l) {
    s.addShape('line', { x: l[0], y: l[1], w: l[2], h: l[3],
      line: { color: GREEN, width: 1.25, dashType: 'sysDot' } });
  });
});

/* 32 - circular process */
slides.push(function (s) {
  infoHeader(s);
  s.addShape('arc', { x: 4.61, y: 3.65, w: 2.09, h: 2.09,
    fill: { type: 'none' }, line: { color: ORANGE, width: 2.5 }, angleRange: [140, 20] });
  s.addShape('arc', { x: 6.65, y: 3.57, w: 2.09, h: 2.09,
    fill: { type: 'none' }, line: { color: ORANGE, width: 2.5 }, angleRange: [320, 200] });
  [[4.84, 3.90, 'trophy'], [6.82, 3.84, 'medal']].forEach(function (c) {
    outline(s, 'ellipse', c[0], c[1], 1.61, 1.61, RED, 2.25);
    icon(s, c[2], c[0] + 0.45, c[1] + 0.4, 0.7, 0.8, RED);
  });
  const nodes = [
    [4.61, 2.69, 'gear', 2.518, 2.618, 'right'],
    [3.51, 4.28, 'clock', 1.482, 4.243, 'right'],
    [6.29, 5.80, 'play', 4.202, 5.745, 'right'],
    [6.34, 2.82, 'book', 7.148, 2.750, 'left'],
    [9.12, 4.31, 'book', 9.917, 4.243, 'left'],
    [8.03, 5.94, 'medal', 8.867, 5.876, 'left'],
  ];
  nodes.forEach(function (n) {
    outline(s, 'ellipse', n[0], n[1], 0.7, 0.7, RED, 1.75);
    icon(s, n[2], n[0] + 0.18, n[1] + 0.18, 0.34, 0.34, RED);
    captionPair(s, 'Add Text', L_TINY, n[3], n[4], 1.924, { align: n[5], gap: 0.211 });
  });
  [[4.22, 4.65, 0.41, 0.04], [5.09, 3.35, 0.15, 0.41], [6.27, 3.47, 0.24, 0.36],
   [8.72, 4.65, 0.40, 0.00], [8.03, 5.56, 0.18, 0.42], [6.80, 5.43, 0.28, 0.38],
  ].forEach(function (l) {
    s.addShape('line', { x: l[0], y: l[1], w: l[2], h: l[3], line: { color: ORANGE, width: 2 } });
  });
});

/* 33 - students around the globe */
slides.push(function (s) {
  infoHeader(s);
  /* four students, drawn first so the wedges overlap their shoulders:
   * [head x, head y, hair, shirt] */
  [[3.42, 4.25, '2E150B', MAGENTA], [5.02, 2.50, '262421', LIME],
   [7.22, 2.55, '3D1B0F', LIME], [8.72, 4.35, '262421', RED],
  ].forEach(function (st) {
    const cx = st[0], cy = st[1];
    s.addShape('trapezoid', { x: cx - 0.3, y: cy + 0.72, w: 1.4, h: 0.78,
      fill: { color: st[3] }, rotate: 180 });
    s.addShape('ellipse', { x: cx, y: cy, w: 0.78, h: 0.86, fill: { color: SKIN } });
    s.addShape('pie', { x: cx - 0.07, y: cy - 0.1, w: 0.92, h: 0.92,
      fill: { color: st[2] }, angleRange: [180, 360] });
  });
  /* four wedges fanned above the globe, all sharing its centre */
  const HUB_X = 6.66, HUB_Y = 5.74, HUB_R = 2.4;
  [[RED, 186, 224], [ORANGE, 229, 267], [RED, 273, 311], [ORANGE, 316, 354]]
    .forEach(function (w) {
      s.addShape('pie', { x: HUB_X - HUB_R, y: HUB_Y - HUB_R, w: HUB_R * 2, h: HUB_R * 2,
        fill: { color: w[0] }, angleRange: [w[1], w[2]] });
    });
  /* globe */
  s.addShape('ellipse', { x: 5.70, y: 4.78, w: 1.92, h: 1.92, fill: { color: CYAN } });
  [[6.0, 5.1, 0.5, 0.4], [6.7, 5.5, 0.6, 0.5], [6.1, 5.9, 0.4, 0.35]].forEach(function (c) {
    s.addShape('ellipse', { x: c[0], y: c[1], w: c[2], h: c[3], fill: { color: GREEN_L } });
  });
  /* two stacks of books + a mouse */
  [5.04, 7.11].forEach(function (bx) {
    [[5.82, CYAN], [5.99, RED], [6.16, GREEN], [6.33, GREEN_M], [6.50, CYAN]]
      .forEach(function (b, i) {
        box(s, bx + (i % 2) * 0.05, b[0], 1.1 - (i % 2) * 0.06, 0.16, b[1]);
      });
  });
  s.addShape('ellipse', { x: 8.41, y: 5.95, w: 1.0, h: 0.8, fill: { color: GREY_TRACK } });
  s.addShape('ellipse', { x: 8.66, y: 5.91, w: 0.5, h: 0.42, fill: { color: RED } });
  /* glyph on each wedge */
  [['gear', 5.72, 3.95, 0.42, 0.42], ['doc', 7.05, 3.95, 0.36, 0.42],
   ['tag', 4.86, 4.83, 0.42, 0.42]]
    .forEach(function (g) { icon(s, g[0], g[1], g[2], g[3], g[4], WHITE); });
  s.addText('\u03C0', { x: 8.02, y: 4.83, w: 0.5, h: 0.5, align: 'center', valign: 'middle',
    margin: 0, fontFace: BODY, fontSize: 20, bold: true, color: WHITE });
  [[1.382, 2.415, 'right', RED], [9.028, 2.415, 'left', RED],
   [0.964, 5.315, 'right', '595959'], [9.446, 5.315, 'left', '595959'],
  ].forEach(function (c) {
    captionPair(s, 'Your Text Here', L_TINY, c[0], c[1], 2.9,
      { align: c[2], headColor: c[3], gap: 0.272, bh: 0.458 });
  });
});

/* 34 - pull quote */
slides.push(function (s) {
  txt(s, 'Everyone breathing is broken. Keep breathing light into them until the stained-glass collage takes your breath away.',
    { x: 1.542, y: 1.814, w: 3.521, h: 2.572, fontSize: 20, color: BLACK,
      lineSpacingMultiple: 1.5 });
  s.addText('Cleva Have A Good Quote', { x: 2.104, y: 5.623, w: 4.029, h: 1.316,
    align: 'right', valign: 'middle', margin: [2, 4, 2, 4], fontFace: HEAD,
    fontSize: 36, bold: true, color: RED, lineSpacingMultiple: 1.1 });
  s.addText('\u201C', { x: 0.864, y: 0.75, w: 0.9, h: 1.2, margin: 0, align: 'left',
    valign: 'middle', fontFace: HEAD, fontSize: 88, bold: true, color: RED });
  s.addText('\u201D', { x: 4.55, y: 4.05, w: 0.9, h: 1.2, margin: 0, align: 'right',
    valign: 'middle', fontFace: HEAD, fontSize: 88, bold: true, color: RED });
  box(s, 8.25, 2.603, 5.083, 4.14, RED);
  photo(s, 6.562, 0.757, 4.724, 4.682);
});

/* 35 - contact */
slides.push(function (s) {
  photo(s, 0, 0, 4.111, 7.5);
  title(s, 'This Our Contact For Find Us', 6.972, 1.598, 4.746, 1.313);
  txt(s, 'PLEASE FEEL FREE TO CALL US ON (PHONE CELL) OR CONTACT US BY (EMAIL), IF YOU REQUIRE ANY FURTHER INFORMATION.',
    { x: 6.988, y: 3.011, w: 4.873, h: 0.627, color: RED, lineSpacingMultiple: 1.5 });
  const blocks = [
    [6.988, 3.974, 'Office Hours', 'Monday \u2013 Thursday\n08:00 \u2013 17:00'],
    [9.599, 3.974, 'Our Address', '1234 Amphitheatre Parkway\nMountain View, LA 12345\nUnited States'],
    [6.988, 5.520, 'Get In Touch', '(+62) 8123 4567 890\n(+62) 908 4567 123\n(0725) 40123'],
    [9.599, 5.520, 'Follow Us', 'www.geaulmanfashion.com\noffice@myexample.com\n@GeaulmanCompany'],
  ];
  blocks.forEach(function (b) {
    txt(s, b[2], { x: b[0], y: b[1], w: 1.327, h: 0.306, fontSize: 12, bold: true, color: RED });
    txt(s, b[3], { x: b[0], y: b[1] + 0.289, w: 2.262, h: 0.934, lineSpacingMultiple: 1.5 });
  });
  ['t', 'f', 'P'].forEach(function (g, i) {
    social(s, g, 5.041, 2.473 + i * 1.4055, 0.584, RED);
  });
  box(s, 2.484, 0, 5.913, 1.128, RED);
});

/* 36 - closing */
slides.push(function (s) {
  photo(s, 0, 0, 13.333, 7.5);
  box(s, 0, 0, 13.333, 7.5, RED, { fill: { color: RED, transparency: 63 } });
  box(s, 3.485, 3.426, 6.413, 2.237, BLACK, { fill: { color: BLACK, transparency: 64 } });
  box(s, 3.485, 5.663, 6.413, 0.127, RED);
  outline(s, 'rect', 0.458, 0.417, 12.417, 6.667, WHITE, 2.25);
  txt(s, 'THANKS', { x: 3.787, y: 3.825, w: 5.76, h: 1.313, fontFace: HEAD,
    fontSize: 72, bold: true, color: CYAN, align: 'center' });
  txt(s, 'BEST UNIVERSITY & COLLEGE', { x: 4.647, y: 4.873, w: 4.088, h: 0.404,
    fontSize: 18, color: CYAN, align: 'center' });
});

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
const pptx = new PptxGenJS();
pptx.author = 'pptxgenjs';
pptx.title = 'Academize - Best University & College';
pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE_16x9';

slides.forEach(function (build) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '196cae3e-5e81-4107-99e8-0909f1a14998_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
