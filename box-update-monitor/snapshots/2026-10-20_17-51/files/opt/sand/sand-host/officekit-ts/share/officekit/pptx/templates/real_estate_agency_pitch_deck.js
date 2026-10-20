#!/usr/bin/env node
/**
 * "Wism - Real Estate PowerPoint" -- 30 slides, 13.333 x 7.5 in (16:9).
 *
 * Rebuilt from scratch with pptxgenjs only. Photo areas of the original deck are
 * empty picture placeholders; they are reproduced here as labelled frames.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme "FH - Real Estate"                                            *
 * ------------------------------------------------------------------ */
const C = {
  teal: '14657A', tealD: '0F4C5C', tealL: '41BFE0', tealXL: 'C0EAF5',
  orange: 'FC8134', orangeD: 'E15903', orangeDD: '963B02', orangeL: 'FDB385',
  deep: '0B4656', deepD: '083541', deepL: '23BAE3',
  peach: 'F8A36C', peachD: 'F46E17', peachDD: 'AA4808',
  mint: '9AD7D0',
  white: 'FFFFFF', ink: '262626', ink05: '0D0D0D', ink40: '404040', ink59: '595959',
  gray: '808080', gray66: 'A6A6A6', gray75: 'BFBFBF', gray85: 'D9D9D9', gray95: 'F2F2F2',
  shell: 'EAEAEA', skin: 'FCD2C5', skin2: 'F7C7AC', skinD: 'F5B591', navy: '203A60',
  gold: 'FFDD4D', frame: 'C3CED3', frameTx: '9AA9AF'
};
const MAJOR = 'Albert Sans';          // theme major latin font
const MINOR = 'DM Sans 14pt Light';   // theme minor latin font

const NONE = { type: 'none' };

/* Boiler-plate copy that repeats across the deck */
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.';
const LOREM_MID = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer.';
const LOREM_CUT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo';
const TITLE = 'From home buying to investment consulting.';
const GROWN = ' has grown into a leading real estate agency with a portfolio spanning residential, commercial, and luxury properties.';

/* ------------------------------------------------------------------ *
 * Primitive helpers                                                   *
 * ------------------------------------------------------------------ */
const box = (x, y, w, h) => ({ x: x, y: y, w: w, h: h });

/** Flat filled shape, no outline (the deck's default). */
function shape(s, kind, x, y, w, h, fill, extra) {
  s.addShape(kind, Object.assign(box(x, y, w, h), { fill: fill ? { color: fill } : NONE, line: NONE }, extra));
}
function rect(s, x, y, w, h, fill, extra) { shape(s, 'rect', x, y, w, h, fill, extra); }
function oval(s, x, y, w, h, fill, extra) { shape(s, 'ellipse', x, y, w, h, fill, extra); }
function roundRect(s, x, y, w, h, fill, radius, extra) {
  shape(s, 'roundRect', x, y, w, h, fill, Object.assign({ rectRadius: radius || 0 }, extra));
}
/** Hairline-outlined, unfilled rectangle. */
function outline(s, x, y, w, h, color, width) {
  s.addShape('rect', Object.assign(box(x, y, w, h), { fill: NONE, line: { color: color, width: width || 0.5 } }));
}

/* Text ------------------------------------------------------------- */
function text(s, body, x, y, w, h, opts) {
  s.addText(body, Object.assign(box(x, y, w, h), { fontFace: MINOR, fontSize: 18, color: C.ink, valign: 'top' }, opts));
}
/** 40 pt display heading (theme major font, tight tracking). */
function heading(s, body, x, y, w, h, color, opts) {
  text(s, body, x, y, w, h, Object.assign({ fontFace: MAJOR, fontSize: 40, charSpacing: -3, color: color }, opts));
}
/** 12 pt paragraph copy at 120 % leading. */
function para(s, body, x, y, w, h, color, opts) {
  text(s, body, x, y, w, h, Object.assign({ fontSize: 12, color: color, lineSpacingMultiple: 1.2 }, opts));
}
/** 16 pt card / section title. */
function subhead(s, body, x, y, w, h, color, opts) {
  text(s, body, x, y, w, h, Object.assign({ fontFace: MAJOR, fontSize: 16, color: color, lineSpacingMultiple: 1.2 }, opts));
}
/** Oversize statistic. */
function stat(s, body, x, y, w, h, color, size, opts) {
  text(s, body, x, y, w, h, Object.assign({ fontFace: MAJOR, fontSize: size, charSpacing: size >= 80 ? -8 : -3, color: color }, opts));
}
/** Pill button (square-cornered rounded rect in this template). */
function button(s, label, x, y, w, h, fill, color) {
  s.addText(label, Object.assign(box(x, y, w, h), {
    shape: 'roundRect', rectRadius: 0, fill: { color: fill }, line: NONE,
    align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 12, color: color || C.white
  }));
}

/** Straight stroke between two points (limbs, poles, arrow shafts). */
function stroke(s, x1, y1, x2, y2, weight, color) {
  const dx = x2 - x1, dy = y2 - y1;
  const len = Math.sqrt(dx * dx + dy * dy);
  const deg = Math.atan2(dy, dx) * 180 / Math.PI;
  s.addShape('rect', {
    x: x1 + dx / 2 - len / 2, y: y1 + dy / 2 - weight / 2, w: len, h: weight,
    fill: { color: color }, line: NONE, rotate: (deg + 360) % 360
  });
}

/**
 * Quarter-turned shape. `w`/`h` describe the on-screen bounding box; the shape is
 * authored with those swapped and rotated 90 degrees about (cx, cy).
 */
function turned(s, kind, cx, cy, w, h, fill) {
  shape(s, kind, cx - h / 2, cy - w / 2, h, w, fill, { rotate: 90, flipH: true });
}

/** Photo slot of the original template, drawn as a labelled frame. */
function imageSlot(s, x, y, w, h, caption) {
  s.addShape('rect', Object.assign(box(x, y, w, h), { fill: NONE, line: { color: C.frame, width: 0.75, dashType: 'dash' } }));
  if (caption === 'Picture') {                       // template slots that carry a visible prompt
    s.addText('Picture', { x: x, y: y + 0.06, w: w, h: 0.26, fontFace: MINOR, fontSize: 12, color: C.ink, align: 'center' });
    glyph(s, 'photo', x + w / 2 - 0.34, y + h / 2 - 0.26, 0.68, 0.52, C.ink59);
  } else {
    s.addText('[image]', { x: x + 0.04, y: y + 0.04, w: Math.min(w - 0.08, 1.2), h: 0.24, fontFace: MINOR, fontSize: 9, color: C.frameTx });
  }
}

/* Master chrome: word-mark, url and page number sit on every slide. */
const CHROME = {
  dark: { mark: C.teal, url: C.gray66, num: C.teal },        // white / light slides
  brand: { mark: C.orange, url: C.gray66, num: C.orange },   // teal slides
  light: { mark: C.white, url: C.gray95, num: C.gray95 }     // orange slides
};
function chrome(s, n, kind) {
  const k = CHROME[kind];
  text(s, 'Wism', 0.366, 0.117, 0.889, 0.337, { fontFace: MAJOR, fontSize: 14, color: k.mark, valign: 'middle' });
  text(s, 'www.wism.com', 11.558, 0.133, 1.405, 0.303, { fontSize: 12, color: k.url, align: 'right', valign: 'middle' });
  text(s, String(n), 6.222, 7.070, 0.889, 0.337, { fontFace: MAJOR, fontSize: 14, color: k.num, align: 'center', valign: 'middle' });
}

/* ------------------------------------------------------------------ *
 * Small line-art glyphs used inside the icon tiles                     *
 * ------------------------------------------------------------------ */
function glyph(s, kind, x, y, w, h, color) {
  const u = (fx, fy, fw, fh, k) => shape(s, k || 'rect', x + fx * w, y + fy * h, fw * w, fh * h, color);
  const ring = (fx, fy, fw, fh) => s.addShape('ellipse', Object.assign(
    box(x + fx * w, y + fy * h, fw * w, fh * h), { fill: NONE, line: { color: color, width: 1.25 } }));
  switch (kind) {
    case 'home':                                     // roof + walls
      shape(s, 'triangle', x, y, w, h * 0.55, color);
      s.addShape('rect', Object.assign(box(x + 0.14 * w, y + 0.5 * h, w * 0.72, h * 0.5),
        { fill: NONE, line: { color: color, width: 1.5 } }));
      break;
    case 'building':                                 // tower with windows
      s.addShape('rect', Object.assign(box(x, y, w, h), { fill: NONE, line: { color: color, width: 1.5 } }));
      u(0.20, 0.18, 0.16, 0.14); u(0.58, 0.18, 0.16, 0.14);
      u(0.20, 0.45, 0.16, 0.14); u(0.58, 0.45, 0.16, 0.14);
      u(0.36, 0.72, 0.24, 0.28);
      break;
    case 'badge':                                    // award medal
      ring(0.15, 0, 0.7, 0.68);
      u(0.22, 0.6, 0.18, 0.4, 'triangle'); u(0.58, 0.6, 0.18, 0.4, 'triangle');
      break;
    case 'globe':
      ring(0, 0, 1, 1); ring(0.32, 0, 0.36, 1); u(0.02, 0.47, 0.96, 0.06);
      break;
    case 'store':                                    // awning + body
      shape(s, 'trapezoid', x, y, w, h * 0.34, color);
      s.addShape('rect', Object.assign(box(x + 0.08 * w, y + 0.34 * h, w * 0.84, h * 0.66),
        { fill: NONE, line: { color: color, width: 1.5 } }));
      break;
    case 'hand':                                     // pointing hand
      shape(s, 'roundRect', x + 0.34 * w, y, w * 0.2, h * 0.55, color, { rectRadius: 0.02 });
      shape(s, 'roundRect', x + 0.2 * w, y + 0.4 * h, w * 0.62, h * 0.6, color, { rectRadius: 0.04 });
      break;
    case 'photo':                                    // framed landscape thumbnail
      s.addShape('rect', Object.assign(box(x, y, w, h), { fill: { color: C.white }, line: { color: color, width: 1 } }));
      shape(s, 'triangle', x + 0.18 * w, y + 0.34 * h, w * 0.6, h * 0.6, '8FBEE8');
      oval(s, x + 0.14 * w, y + 0.16 * h, w * 0.16, h * 0.2, C.gold);
      break;
  }
}
/** Coloured square tile with a glyph centred inside it. */
function iconTile(s, x, y, size, tile, glyphKind, glyphColor, gx, gy, gw, gh) {
  rect(s, x, y, size, size, tile);
  glyph(s, glyphKind, gx, gy, gw, gh, glyphColor);
}

/* ------------------------------------------------------------------ *
 * Repeating composites                                                *
 * ------------------------------------------------------------------ */
/** "Option One / lorem / [Contact Now]" card body. */
function optionText(s, x, y, title, copy, color, btn) {
  subhead(s, title, x, y, 2.203, 0.399, color);
  para(s, copy, x, y + 0.44, 2.394, 0.573, color);
  if (btn) button(s, btn.label, x + 0.112, y + 1.282, 1.361, 0.438, btn.fill);
}
/** Numbered column: big "01", title, copy. */
function numberedItem(s, x, y, num, title, copy, numColor, titleColor, copyColor, numSize) {
  stat(s, num, x, y, 0.896, 0.64, numColor, numSize || 32);
  subhead(s, title, x, y + 0.68, 2.203, 0.399, titleColor);
  para(s, copy, x, y + 1.121, 2.394, 0.573, copyColor);
}
/** Price row: "$ 388.99  $ 420.00  [Buy Now]". */
function priceRow(s, x, y, price, was, panel) {
  rect(s, x, y, 4.785, 1.172, panel);
  stat(s, price, x + 0.324, y + 0.3, 2.028, 0.572, C.white, 28, { charSpacing: 0 });
  text(s, was, x + 2.022, y + 0.434, 1.479, 0.303, { fontFace: MAJOR, fontSize: 12, color: C.gray66, strike: true });
  button(s, 'Buy Now', x + 3.1, y + 0.367, 1.361, 0.438, C.orange);
}

/* ------------------------------------------------------------------ *
 * Slides                                                              *
 * ------------------------------------------------------------------ */
const slides = [];
const slide = fn => slides.push(fn);

/* 1 - Title ---------------------------------------------------------- */
slide(function (s) {
  s.background = { color: C.teal };
  chrome(s, 1, 'brand');
  rect(s, 5.227, 2.181, 1.439, 4.024, C.orange);
  rect(s, 11.911, 3.167, 2.200, 3.038, C.orange);
  stat(s, 'Wism', 6.972, 4.193, 3.361, 1.582, C.white, 88);
  text(s, 'Real Estate PowerPoint', 6.972, 5.633, 4.125, 0.572, { fontSize: 28, color: C.white, charSpacing: -1.5 });
  text(s, 'New York City\nJune, 2025', 9.561, 3.167, 1.967, 0.640, { fontSize: 16, color: C.white, align: 'right' });
  imageSlot(s, 0.366, 0.521, 5.870, 6.458);
});

/* 2 - Intro + stat --------------------------------------------------- */
slide(function (s) {
  chrome(s, 2, 'dark');
  rect(s, 4.652, 0.521, 8.315, 3.996, C.teal);
  stat(s, '$760B', 9.335, 4.910, 3.257, 1.447, C.orange, 80);
  para(s, LOREM_SHORT, 9.335, 6.161, 2.426, 0.505, C.ink, { lineSpacingMultiple: 1 });
  heading(s, [
    { text: 'Discover seamless property solutions with ' },
    { text: 'Wism.', options: { bold: true } }
  ], 5.633, 1.266, 6.564, 1.447, C.orange);
  para(s, LOREM_LONG, 5.633, 2.840, 4.171, 0.816, C.white);
  button(s, 'Contact Now', 11.231, 2.975, 1.361, 0.438, C.orange);
  imageSlot(s, 0.366, 2.547, 4.819, 4.432);
});

/* 3 - Two photos + checklist ----------------------------------------- */
slide(function (s) {
  chrome(s, 3, 'dark');
  rect(s, 0.366, 3.698, 6.078, 2.925, C.orange);
  heading(s, 'We\u2019re here to guide you every step of the way.', 8.352, 2.372, 4.296, 2.121, C.teal);
  para(s, LOREM_MID, 0.679, 5.484, 2.530, 0.573, C.gray95);
  iconTile(s, 0.742, 4.737, 0.558, C.teal, 'home', C.white, 0.821, 4.816, 0.399, 0.399);
  para(s, ['Lorem ipsum dolor sit amet', 'Lorem ipsum dolor sit amet', 'Lorem ipsum dolor sit amet'].join('\n'),
    8.430, 4.737, 2.530, 0.816, C.gray, { bullet: { characterCode: '2713', indent: 13.5 } });
  imageSlot(s, 0.742, 0.877, 3.573, 3.281);
  imageSlot(s, 4.495, 2.371, 3.573, 3.281);
});

/* 4 - Sectors + free consultation ------------------------------------ */
slide(function (s) {
  chrome(s, 4, 'dark');
  heading(s, TITLE, 3.477, 2.372, 4.792, 2.121, C.ink);
  para(s, LOREM_LONG, 3.477, 4.660, 4.171, 0.816, C.gray);
  [['Barcelona, Spain', 1.250, C.gray75], ['Tokyo, Japan', 2.556, C.teal], ['Rio de Janeiro, Brazil', 5.803, C.gray75]]
    .forEach(function (r) { text(s, 'Sector 02\n' + r[0], 0.839, r[1], 1.967, 0.505, { fontSize: 12, color: r[2] }); });
  rect(s, 8.542, 4.258, 3.089, 2.721, C.orange);
  subhead(s, 'Free Consultation', 8.889, 4.759, 2.203, 0.399, C.white);
  para(s, LOREM_MID, 8.889, 5.199, 2.394, 0.573, C.white);
  button(s, 'Contact Now', 9.001, 6.041, 1.361, 0.438, C.teal);
  imageSlot(s, 8.542, 0.000, 4.792, 7.500);
});

/* 5 - Bar columns ---------------------------------------------------- */
slide(function (s) {
  chrome(s, 5, 'dark');
  const cols = [
    { x: 6.804, top: 4.382, h: 1.958, fill: C.teal, label: '$40B' },
    { x: 8.660, top: 3.424, h: 2.917, fill: C.orange, label: '$60B' },
    { x: 10.516, top: 2.424, h: 3.917, fill: C.deep, label: '$80B' }
  ];
  cols.forEach(function (c) { rect(s, c.x, c.top, 1.856, c.h, c.fill); });
  cols.forEach(function (c) { outline(s, c.x, 1.160, 1.856, 5.181, C.gray95, 0.5); });
  cols.forEach(function (c) { stat(s, c.label, c.x + 0.152, 1.367, 1.552, 0.774, C.orange, 40); });
  text(s, [
    { text: 'Wism', options: { color: C.teal } },
    { text: GROWN, options: { color: C.ink } }
  ], 0.961, 1.263, 5.065, 2.794, { fontFace: MAJOR, fontSize: 32, charSpacing: -3 });
  text(s, 'Data on 2025\nBy Wism Agency', 0.961, 5.732, 1.967, 0.505, { fontSize: 12, color: C.teal });
});

/* 6 - Orange statement ----------------------------------------------- */
slide(function (s) {
  s.background = { color: C.orange };
  chrome(s, 6, 'light');
  heading(s, 'From home buying to investment consulting \u2013 we\u2019re here to guide you every step of the way.',
    5.833, 3.480, 6.814, 2.794, C.white);
  stat(s, '$1250M', 5.833, 1.226, 3.953, 1.447, C.teal, 80);
  para(s, LOREM_SHORT, 5.833, 2.477, 2.426, 0.505, C.white, { lineSpacingMultiple: 1 });
  para(s, LOREM_MID, 0.686, 5.624, 2.530, 0.573, C.gray95);
  iconTile(s, 0.749, 4.877, 0.558, C.teal, 'hand', C.white, 0.829, 4.956, 0.399, 0.399);
  subhead(s, 'Establish in 2015', 1.460, 4.956, 2.203, 0.399, C.white);
  imageSlot(s, 0.686, 0.840, 3.814, 3.680);
});

/* 7 - Mission + three stacked options -------------------------------- */
slide(function (s) {
  chrome(s, 7, 'dark');
  const rows = [
    { y: 0.521, fill: C.orange, num: '01', title: 'Option One' },
    { y: 2.674, fill: C.deep, num: '02', title: 'Option Two' },
    { y: 4.826, fill: C.orange, num: '03', title: 'Option Three' }
  ];
  rows.forEach(function (r) { rect(s, 6.667, r.y, 6.300, 2.153, r.fill); });
  rect(s, 0.366, 0.521, 6.300, 6.458, C.teal);
  heading(s, 'To deliver outstanding real estate experiences through innovation and integrity.',
    0.807, 2.579, 5.419, 2.794, C.white);
  rows.forEach(function (r) {
    stat(s, r.num, 7.017, r.y + 0.653, 0.802, 0.847, C.white, 40, { lineSpacingMultiple: 1.2 });
    subhead(s, r.title, 8.023, r.y + 0.546, 2.864, 0.437, C.white);
    para(s, LOREM_LONG, 8.023, r.y + 1.034, 4.594, 0.573, C.white);
  });
  button(s, 'Contact Now', 0.807, 5.671, 1.361, 0.438, C.orange);
});

/* 8 - Photo trio ----------------------------------------------------- */
slide(function (s) {
  chrome(s, 8, 'dark');
  heading(s, TITLE, 7.014, 2.372, 4.792, 2.121, C.ink);
  para(s, LOREM_LONG, 7.014, 4.660, 4.171, 0.816, C.gray);
  text(s, 'Sector 02\nTokyo, Japan', 0.961, 5.488, 1.967, 0.505, { fontSize: 12, color: C.teal });
  text(s, 'Sector 02\nBarcelona, Spain', 4.292, 1.507, 1.967, 0.505, { fontSize: 12, color: C.teal });
  imageSlot(s, 0.750, 1.507, 3.194, 3.542);
  imageSlot(s, 4.292, 3.743, 2.375, 2.250);
  imageSlot(s, 11.264, 4.743, 1.319, 1.250);
});

/* 9 - Three photo columns with cards --------------------------------- */
slide(function (s) {
  chrome(s, 9, 'dark');
  const cards = [
    { x: 0.366, y: 4.258, title: 'Option One' },
    { x: 4.566, y: 0.521, title: 'Option Two' },
    { x: 8.767, y: 4.258, title: 'Option Three' }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, c.y, 3.089, 2.721, C.orange);
    optionText(s, c.x + 0.347, c.y + 0.501, c.title, LOREM_MID, C.white, { label: 'Contact Now', fill: C.teal });
  });
  [0.366, 4.567, 8.767].forEach(function (x) { imageSlot(s, x, 0.521, 4.200, 6.458); });
});

/* 10 - Horizontal bars ----------------------------------------------- */
slide(function (s) {
  chrome(s, 10, 'dark');
  const bars = [
    { y: 1.477, w: 6.665, fill: C.teal, value: '24K', vx: 7.175, year: '2022' },
    { y: 2.699, w: 4.729, fill: C.orange, value: '18K', vx: 5.242, year: '2023' },
    { y: 3.922, w: 3.467, fill: C.teal, value: '12K', vx: 3.987, year: '2024' },
    { y: 5.144, w: 5.280, fill: C.orange, value: '19K', vx: 5.797, year: '2025' }
  ];
  bars.forEach(function (b) { rect(s, 0.366, b.y, b.w, 1.051, b.fill); });
  bars.forEach(function (b) {
    stat(s, b.value, b.vx, b.y + 0.138, 1.552, 0.774, b.fill, 40);
    subhead(s, 'Data on ' + b.year, 0.705, b.y + 0.307, 2.864, 0.437, C.white, { fontSize: 18 });
  });
  text(s, [
    { text: 'Wism', options: { color: C.teal } },
    { text: GROWN, options: { color: C.ink } }
  ], 8.032, 3.401, 4.596, 2.794, { fontFace: MAJOR, fontSize: 32, charSpacing: -3 });
});

/* 11 - Two icon cards + full-bleed photo ------------------------------ */
slide(function (s) {
  chrome(s, 11, 'dark');
  const cards = [
    { x: 0.965, fill: C.teal, title: 'Option One', gx: 1.554, gy: 4.224, gw: 0.417, gh: 0.375 },
    { x: 4.471, fill: C.orange, title: 'Option Two', gx: 5.038, gy: 4.203, gw: 0.458, gh: 0.417 }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, 3.756, 3.384, 2.623, c.fill);
    subhead(s, c.title, c.x + 0.495, 5.141, 2.203, 0.399, C.white);
    para(s, LOREM_MID, c.x + 0.495, 5.581, 2.394, 0.573, C.white);
    glyph(s, 'building', c.gx, c.gy, c.gw, c.gh, C.white);
  });
  heading(s, TITLE, 0.848, 1.121, 4.648, 2.121, C.ink);
  imageSlot(s, 6.667, 0.000, 6.667, 7.500);
});

/* 12 - Teal slide, three numbered items ------------------------------- */
slide(function (s) {
  s.background = { color: C.teal };
  chrome(s, 12, 'brand');
  rect(s, 0.000, 0.000, 4.132, 5.698, C.orange);
  heading(s, TITLE, 0.557, 2.353, 3.114, 2.794, C.white);
  [['01', 'Option One', 4.540], ['02', 'Option Two', 7.461], ['03', 'Option Three', 10.382]]
    .forEach(function (it) { numberedItem(s, it[2], 3.278, it[0], it[1], LOREM_MID, C.white, C.white, C.white); });
  para(s, LOREM_LONG, 4.993, 5.807, 4.171, 0.816, C.white);
  glyph(s, 'badge', 4.540, 5.990, 0.333, 0.450, C.orange);
  button(s, 'Contact Now', 11.224, 5.990, 1.361, 0.438, C.orange);
  imageSlot(s, 3.132, 0.000, 10.201, 2.443);
});

/* 13 - Percentage split ----------------------------------------------- */
slide(function (s) {
  chrome(s, 13, 'dark');
  rect(s, 4.958, 5.417, 7.357, 1.562, C.orange);
  heading(s, TITLE, 5.990, 1.031, 4.635, 2.121, C.ink);
  para(s, LOREM_LONG, 5.990, 3.320, 4.171, 0.816, C.gray);
  subhead(s, 'Option One', 7.367, 5.710, 2.203, 0.399, C.white);
  para(s, LOREM_MID, 7.367, 6.150, 2.394, 0.573, C.white);
  stat(s, '30.15%', 5.146, 5.830, 1.928, 0.774, C.white, 40, { align: 'right' });
  stat(s, '69.75%', 2.726, 5.847, 1.928, 0.774, C.teal, 40, { align: 'right' });
  rect(s, 0.366, 3.498, 3.113, 1.082, C.teal);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer', 1.196, 3.752, 1.988, 0.573, C.white);
  glyph(s, 'globe', 0.662, 3.830, 0.417, 0.417, C.orange);
  imageSlot(s, 0.366, 0.521, 5.093, 4.058);
  imageSlot(s, 9.955, 4.579, 3.012, 2.400);
});

/* 14 - Two price cards ------------------------------------------------ */
slide(function (s) {
  chrome(s, 14, 'dark');
  heading(s, TITLE, 0.649, 1.553, 4.792, 2.121, C.ink);
  para(s, LOREM_LONG, 0.649, 3.842, 4.171, 0.816, C.gray);
  priceRow(s, 0.656, 5.271, '$ 388.99', '$ 420.00', C.teal);
  priceRow(s, 5.720, 5.271, '$ 667.88', '$ 820.00', C.teal);
  imageSlot(s, 5.720, 0.521, 3.792, 4.271);
  imageSlot(s, 9.792, 3.354, 3.175, 3.625);
});

/* 15 - Tinted photo panel + offer ------------------------------------- */
slide(function (s) {
  chrome(s, 15, 'dark');
  s.addShape('rect', Object.assign(box(0.366, 0.521, 8.842, 5.125), { fill: { color: C.teal, transparency: 58 }, line: NONE }));
  rect(s, 9.208, 1.981, 1.244, 1.082, C.teal);
  glyph(s, 'home', 9.573, 2.292, 0.515, 0.460, C.orange);
  rect(s, 6.667, 3.062, 6.300, 3.917, C.orange);
  heading(s, TITLE, 7.368, 3.525, 4.792, 2.121, C.white);
  para(s, LOREM_LONG, 7.368, 5.814, 4.171, 0.816, C.white);
  priceRow(s, 0.786, 3.062, '$ 388.99', '$ 420.00', C.teal);
  para(s, LOREM_LONG, 0.786, 4.481, 4.171, 0.816, C.white);
  rect(s, 5.422, 5.646, 1.244, 1.082, C.teal);
  glyph(s, 'store', 5.815, 5.978, 0.458, 0.417, C.orange);
  imageSlot(s, 0.366, 0.521, 8.842, 5.125);
});

/* 16 - Full-bleed tint with three numbered items ---------------------- */
slide(function (s) {
  chrome(s, 16, 'dark');
  s.addShape('rect', Object.assign(box(0, 0, 13.333, 7.5), { fill: { color: C.teal, transparency: 55 }, line: NONE }));
  text(s, 'Wism' + GROWN, 0.896, 2.308, 7.816, 1.717, { fontFace: MAJOR, fontSize: 32, charSpacing: -3, color: C.white });
  [['01', 'Option One', 0.896], ['02', 'Option Two', 4.205], ['03', 'Option Three', 7.515]]
    .forEach(function (it) { numberedItem(s, it[2], 4.700, it[0], it[1], LOREM_MID, C.white, C.white, C.white); });
  imageSlot(s, 0.000, 0.000, 13.333, 7.500);
});

/* 17 - Product cards -------------------------------------------------- */
slide(function (s) {
  chrome(s, 17, 'dark');
  rect(s, 0.800, 1.374, 2.695, 4.804, C.teal);
  [3.813, 6.825, 9.838].forEach(function (x, i) {
    para(s, LOREM_CUT, x, 3.750, 2.521, 0.816, C.gray);
    button(s, 'Buy Now', x, 4.729, 1.361, 0.438, C.orange);
    text(s, ['$ 410.899', '$ 299.999', '$ 767.899'][i], [3.803, 6.825, 9.779][i], 5.742, 2.028, 0.438,
      { fontFace: MAJOR, fontSize: 20, color: C.ink });
    imageSlot(s, x, 1.374, 2.695, 1.973, 'Picture');
  });
  heading(s, 'Wism Real Estate', 1.107, 3.698, 1.874, 2.121, C.white);
  rect(s, 0.800, 1.374, 1.244, 1.082, C.orange);
  glyph(s, 'home', 1.165, 1.685, 0.515, 0.460, C.white);
});

/* 18 - Split panel with badges ---------------------------------------- */
slide(function (s) {
  chrome(s, 18, 'dark');
  s.addShape('rect', Object.assign(box(0, 0, 4.689, 7.5), { fill: { color: C.teal, transparency: 55 }, line: NONE }));
  rect(s, 2.443, 1.180, 3.444, 4.961, C.teal);
  heading(s, TITLE, 1.198, 1.597, 4.556, 2.121, C.white);
  para(s, LOREM_LONG, 2.795, 4.422, 2.741, 1.058, C.white);
  button(s, 'Contact Now', 2.889, 5.882, 1.361, 0.438, C.orange);
  [
    { y: 1.421, num: '01', title: 'Option One', badge: C.teal, tx: 6.874, tw: 0.562 },
    { y: 3.211, num: '02', title: 'Option Two', badge: C.orange, tx: 6.770, tw: 0.768 },
    { y: 5.000, num: '03', title: 'Option Three', badge: C.teal, tx: 6.811, tw: 0.686 }
  ].forEach(function (r) {
    subhead(s, r.title, 7.911, r.y, 2.203, 0.399, C.ink);
    para(s, LOREM_LONG, 7.911, r.y + 0.441, 4.423, 0.573, C.gray);
    rect(s, 6.667, r.y + 0.095, 0.975, 0.824, r.badge);
    stat(s, r.num, r.tx, r.y + 0.187, r.tw, 0.640, C.white, 32, { align: 'center', valign: 'middle' });
  });
  imageSlot(s, 0.000, 0.000, 4.690, 7.500);
});

/* 19 - Buildings ------------------------------------------------------ */
slide(function (s) {
  chrome(s, 19, 'dark');
  heading(s, TITLE, 0.649, 1.138, 4.620, 2.121, C.ink);
  para(s, LOREM_LONG, 0.649, 3.427, 3.748, 0.816, C.gray);
  [
    { x: 5.269, title: 'Building One', gx: 5.378, gy: 4.859, gw: 0.458, gh: 0.412 },
    { x: 8.382, title: 'Building Two', gx: 8.485, gy: 4.854, gw: 0.375, gh: 0.417 }
  ].forEach(function (b) {
    subhead(s, b.title, b.x, 5.531, 2.203, 0.399, C.ink);
    para(s, LOREM_MID, b.x, 5.972, 2.394, 0.573, C.gray);
    glyph(s, 'building', b.gx, b.gy, b.gw, b.gh, C.orange);
  });
  imageSlot(s, 5.269, 0.000, 8.064, 4.345);
  imageSlot(s, 0.000, 4.854, 4.398, 2.646);
});

/* 20 - Team ----------------------------------------------------------- */
slide(function (s) {
  s.background = { color: C.orange };
  chrome(s, 20, 'light');
  const people = [
    { card: [0.800, 2.389, 2.769, 3.514], fill: C.white, tx: 1.084, mail: 1.135, ph: [1.568, 3.032, 1.235], name: C.ink, mailC: C.gray },
    { card: [3.569, 2.111, 3.207, 4.070], fill: C.teal, tx: 4.071, mail: 4.122, ph: [4.361, 2.644, 1.623], name: C.white, mailC: C.gray85 },
    { card: [6.776, 2.389, 2.769, 3.514], fill: C.white, tx: 7.059, mail: 7.110, ph: [7.543, 3.032, 1.235], name: C.ink, mailC: C.gray },
    { card: [9.764, 2.389, 2.769, 3.514], fill: C.white, tx: 10.047, mail: 10.098, ph: [10.531, 3.032, 1.235], name: C.ink, mailC: C.gray }
  ];
  people.forEach(function (p) { rect(s, p.card[0], p.card[1], p.card[2], p.card[3], p.fill); });
  people.forEach(function (p) {
    subhead(s, 'Alex Pastoor', p.tx, 4.741, 2.203, 0.399, p.name, { align: 'center' });
    para(s, 'account@mail.com', p.mail, 5.098, 2.102, 0.331, p.mailC, { align: 'center' });
  });
  heading(s, 'Amazing Our Team', 3.050, 0.863, 7.233, 0.774, C.white, { align: 'center' });
  button(s, 'Contact Now', 4.388, 6.012, 1.570, 0.438, C.white, C.teal);
  people.forEach(function (p) { imageSlot(s, p.ph[0], p.ph[1], p.ph[2], p.ph[2], 'Picture'); });
});

/* 21 - Funnel infographic --------------------------------------------- */
slide(function (s) {
  chrome(s, 21, 'dark');
  const layers = [
    { x: 4.892, rimY: 2.716, rimH: 0.299, y: 2.865, w: 3.552, h: 0.970, fill: C.teal, rim: '0A323D', label: '01', size: 28 },
    { x: 5.357, rimY: 3.752, rimH: 0.225, y: 3.864, w: 2.621, h: 0.947, fill: C.orange, rim: C.orangeDD, label: '02', size: 28 },
    { x: 5.821, rimY: 4.735, rimH: 0.256, y: 4.864, w: 1.693, h: 0.943, fill: C.deep, rim: '06232B', label: '03', size: 28 },
    { x: 6.286, rimY: 5.782, rimH: 0.160, y: 5.862, w: 0.764, h: 0.822, fill: C.peach, rim: C.peachDD, label: '04', size: 12 }
  ];
  layers.forEach(function (L) {
    oval(s, L.x, L.rimY, L.w, L.rimH * 1.5, L.rim);                       // dark elliptical rim
    shape(s, 'trapezoid', L.x, L.y, L.w, L.h, L.fill, { flipV: true });   // tapering body
    text(s, L.label, L.x, L.y, L.w, L.h * 0.8, { fontFace: MAJOR, fontSize: L.size, color: C.white, align: 'center', valign: 'middle' });
  });
  const cards = [
    { x: 0.816, y: 2.865, title: 'Awareness', pct: '40%', color: C.teal },
    { x: 3.099, y: 3.916, title: 'Collaboration', pct: '56%', color: C.orange },
    { x: 8.067, y: 3.916, title: 'Discussion', pct: '12%', color: C.deep },
    { x: 10.350, y: 2.865, title: 'Test', pct: '24%', color: C.peach }
  ];
  cards.forEach(function (c) {
    s.addShape('roundRect', Object.assign(box(c.x, c.y, 2.168, 1.830), {
      fill: { color: C.white }, line: NONE, rectRadius: 0.196,
      shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 40, offset: 12, angle: 72 }
    }));
    text(s, c.title, c.x + 0.234, c.y + 0.318, 1.701, 0.345, { fontFace: MAJOR, fontSize: 16, color: C.ink05, align: 'center' });
    text(s, c.pct, c.x + 0.234, c.y + 0.649, 1.701, 0.640, { fontFace: MAJOR, fontSize: 32, bold: true, color: c.color, align: 'center' });
    text(s, 'Growth Results', c.x + 0.234, c.y + 1.249, 1.701, 0.303, { fontSize: 12, color: C.ink40, align: 'center' });
  });
  heading(s, 'Funnel Customer Infographic', 2.540, 0.816, 8.254, 0.774, C.ink, { align: 'center' });
  para(s, LOREM_LONG, 3.790, 1.675, 5.753, 0.573, C.gray, { align: 'center' });
});

/* 22 - Three cone funnels --------------------------------------------- */
slide(function (s) {
  chrome(s, 22, 'dark');
  const cones = [
    { x: 2.208, color: C.teal, tint: C.tealL, pct: '34%', tag: 'Awareness', cap: 1.780, capW: 2.564, capY: 5.900, capH: 0.610 },
    { x: 5.811, color: C.orange, tint: C.orangeL, pct: '50%', tag: 'Collab', cap: 5.383, capW: 2.564, capY: 5.950, capH: 0.510 },
    { x: 9.414, color: C.deep, tint: C.deepL, pct: '85%', tag: 'Discuss', cap: 8.799, capW: 2.944, capY: 5.910, capH: 0.570 }
  ];
  cones.forEach(function (c) {
    const cx = c.x - 0.427;                                   // ground ring block origin
    s.addShape('ellipse', Object.assign(box(cx, 4.200, 2.564, 1.550), { fill: NONE, line: { color: C.gray85, width: 0.5 } }));
    s.addShape('ellipse', Object.assign(box(c.x, 4.440, 1.720, 1.040), { fill: NONE, line: { color: C.gray85, width: 0.5 } }));
    oval(s, c.x + 0.530, 4.840, 0.630, 0.380, C.gray66);
    oval(s, c.x + 0.540, 4.930, 0.610, 0.290, C.ink59);
    shape(s, 'triangle', c.x, 3.300, 1.720, 1.390, c.tint, { flipV: true });            // cone body
    oval(s, c.x, 2.870, 1.720, 1.010, c.tint);                                          // outer rim
    oval(s, c.x + 0.050, 2.920, 1.610, 0.920, c.color);                                 // rim face
    oval(s, c.x + 0.230, 3.020, 1.260, 0.560, c.color === C.teal ? C.tealD : (c.color === C.orange ? C.orangeD : C.deepD));
    rect(s, c.x + 0.740, 4.630, 0.250, 0.260, c.tint);                                  // stem
    oval(s, c.x + 0.220, 2.100, 1.260, 1.260, C.white);                                 // badge
    text(s, c.pct, c.x + 0.060, 2.360, 1.590, 0.570, { fontFace: MAJOR, fontSize: 28, bold: true, color: C.ink, align: 'center' });
    text(s, c.tag, c.x + 0.060, 2.810, 1.590, 0.300, { fontSize: 12, color: C.gray, align: 'center' });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing', c.cap, c.capY, c.capW, c.capH, C.gray, { align: 'center' });
  });
  heading(s, 'Funnel Customer Infographic', 2.322, 0.905, 8.254, 0.774, C.ink, { align: 'center' });
});

/* 23 - Painters illustration ------------------------------------------ */
slide(function (s) {
  chrome(s, 23, 'dark');
  shape(s, 'round1Rect', 6.106, -0.005, 7.227, 7.505, C.teal, { rectRadius: 1.8, flipH: true });
  /* Three painters rolling paint on the wall: figure + long pole + roller head. */
  const painters = [
    { hx: 6.59, hy: 1.95, hair: C.ink, body: C.deepD, leg: C.skin, legW: 0.20, legH: 0.95,
      bx: 6.36, by: 2.32, bw: 0.80, bh: 1.32, shX: 5.38, shY: 5.37, shW: 3.08,
      handX: 7.35, handY: 2.95, rollX: 8.56, rollY: 1.04, rollFill: C.gold, ladder: true },
    { hx: 8.96, hy: 3.41, hair: C.ink, body: C.orange, leg: C.orangeD, legW: 0.32, legH: 1.66,
      bx: 8.70, by: 3.88, bw: 0.65, bh: 1.09, shX: 7.53, shY: 6.15, shW: 2.14,
      handX: 9.66, handY: 4.05, rollX: 11.05, rollY: 2.12, rollFill: C.orange, ladder: false },
    { hx: 10.61, hy: 4.72, hair: C.navy, body: C.deep, leg: C.deepD, legW: 0.28, legH: 1.51,
      bx: 9.79, by: 4.98, bw: 0.89, bh: 0.83, shX: 8.75, shY: 6.88, shW: 2.14,
      handX: 11.30, handY: 5.10, rollX: 11.67, rollY: 4.25, rollFill: C.gold, ladder: false }
  ];
  painters.forEach(function (p) {
    oval(s, p.shX, p.shY, p.shW, 0.41, C.gray85);                                      // ground shadow
    if (p.ladder) {                                                                    // step-ladder
      stroke(s, 6.52, 3.08, 6.78, 5.63, 0.07, C.ink);
      stroke(s, 7.38, 3.08, 7.12, 5.63, 0.07, C.ink);
      [3.65, 4.02, 4.39, 4.77, 5.14].forEach(function (ry) { stroke(s, 6.62, ry, 7.28, ry, 0.045, C.ink); });
    }
    shape(s, 'rect', p.rollX + 0.16, p.rollY + 0.44, 0.48, 0.25, p.rollFill);          // roller pin
    shape(s, 'roundRect', p.rollX, p.rollY, 0.63, 0.58, C.white, { rectRadius: 0.12 });// roller sleeve
    stroke(s, p.handX, p.handY, p.rollX + 0.28, p.rollY + 0.56, 0.055, C.ink);         // pole
    shape(s, 'rect', p.bx + 0.06, p.by + p.bh - 0.05, p.legW, p.legH, p.leg);          // legs
    shape(s, 'rect', p.bx + p.bw - p.legW - 0.06, p.by + p.bh - 0.05, p.legW, p.legH, p.leg);
    shape(s, 'roundRect', p.bx, p.by, p.bw, p.bh, p.body, { rectRadius: 0.08 });       // torso / dress
    oval(s, p.hx, p.hy, 0.38, 0.44, p.hair);                                           // hair
    oval(s, p.hx + 0.10, p.hy + 0.09, 0.24, 0.32, C.skin);                             // face
    stroke(s, p.bx + p.bw * 0.75, p.by + 0.20, p.handX, p.handY, 0.12, C.skin);        // arm
  });
  const stats = [
    { x: 0.918, fill: C.teal, value: '500.987', vc: C.white, cc: C.white },
    { x: 4.132, fill: C.white, value: '275.292', vc: C.ink, cc: C.gray }
  ];
  stats.forEach(function (st) {
    s.addShape('rect', Object.assign(box(st.x, 4.560, 2.940, 1.400), {
      fill: { color: st.fill }, line: NONE,
      shadow: { type: 'outer', color: '000000', opacity: 0.15, blur: 60, offset: 8, angle: 90 }
    }));
    text(s, st.value, st.x + 0.31, 4.770, 2.300, 0.640, { fontFace: MAJOR, fontSize: 32, bold: true, color: st.vc, align: 'center' });
    text(s, 'Insert Your Creative Text', st.x + 0.39, 5.380, 2.150, 0.340, { fontSize: 12, color: st.cc, align: 'center' });
  });
  heading(s, TITLE, 0.918, 1.544, 5.370, 1.450, C.ink);
  para(s, LOREM_LONG, 0.950, 3.170, 4.670, 0.610, C.gray);
});

/* 24 - Rising arrow --------------------------------------------------- */
slide(function (s) {
  chrome(s, 24, 'dark');
  rect(s, 0.010, 6.630, 7.680, 0.320, C.teal);                       // ground rail
  rect(s, 0.010, 6.950, 7.680, 0.060, C.tealXL);
  /* Rising arrow: authored upright then quarter-turned, so the pre-rotation box
     has width and height swapped around the same centre point. */
  turned(s, 'bentArrow', 10.285, 4.685, 5.630, 4.650, C.tealXL);   // soft outer glow
  turned(s, 'bentArrow', 10.285, 4.715, 5.530, 4.520, C.teal);
  [[11.94, 2.71], [12.32, 4.02], [12.01, 5.31], [11.03, 6.30], [8.96, 6.69]].forEach(function (p) {
    s.addShape('ellipse', Object.assign(box(p[0], p[1], 0.23, 0.23), { fill: { color: C.white }, line: { color: C.tealD, width: 3 } }));
  });
  /* Two business people sprinting up the rail, briefcase in the trailing hand.
     Each runner is laid out relative to the top-left of its hair oval. */
  [{ hx: 8.65, hy: 4.00, body: C.tealD }, { hx: 7.15, hy: 4.25, body: C.orange }].forEach(function (r) {
    stroke(s, r.hx + 0.50, r.hy + 1.90, r.hx + 0.10, r.hy + 2.55, 0.20, r.body);          // trailing leg
    stroke(s, r.hx + 0.90, r.hy + 1.85, r.hx + 1.55, r.hy + 2.45, 0.20, r.body);          // leading leg
    shape(s, 'roundRect', r.hx - 0.14, r.hy + 2.48, 0.42, 0.20, C.ink, { rectRadius: 0.04 });  // shoes
    shape(s, 'roundRect', r.hx + 1.40, r.hy + 2.38, 0.42, 0.20, C.ink, { rectRadius: 0.04 });
    shape(s, 'roundRect', r.hx + 0.20, r.hy + 1.32, 0.82, 0.72, r.body, { rectRadius: 0.08 }); // hips
    shape(s, 'roundRect', r.hx + 0.14, r.hy + 0.72, 0.88, 0.74, C.gray85, { rectRadius: 0.1 });// shirt
    oval(s, r.hx, r.hy, 0.92, 0.84, C.ink);                                                // hair
    oval(s, r.hx + 0.18, r.hy + 0.14, 0.68, 0.66, C.skin2);                                // face
    stroke(s, r.hx + 0.20, r.hy + 0.90, r.hx - 0.40, r.hy + 1.22, 0.14, C.skin2);          // trailing arm
    shape(s, 'roundRect', r.hx - 0.80, r.hy + 1.18, 0.72, 0.60, C.ink, { rectRadius: 0.05 });  // briefcase
  });
  const legend = [
    { x: 1.62, y: 3.84, c: C.orange, label: 'Product Owner', lx: 1.84 },
    { x: 1.62, y: 4.22, c: C.deep, label: 'Product Backing', lx: 1.84 },
    { x: 4.13, y: 3.84, c: C.peach, label: 'Sprint Backing', lx: 4.34 },
    { x: 4.13, y: 4.22, c: C.teal, label: 'Product', lx: 4.34 }
  ];
  s.addShape('roundRect', Object.assign(box(1.30, 3.50, 5.10, 1.28), {
    fill: { color: C.white }, line: NONE, rectRadius: 0.14,
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 40, offset: 6, angle: 90 }
  }));
  legend.forEach(function (g) {
    oval(s, g.x, g.y, 0.17, 0.17, g.c);
    text(s, g.label, g.lx, g.y - 0.10, 1.92, 0.37, { fontFace: MAJOR, fontSize: 14, color: C.ink05 });
  });
  [
    { t: 'Goals!!!', x: 9.747, y: 1.334, w: 2.475, h: 0.774, sz: 40, c: C.teal, b: true, rot: 342, al: 'center' },
    { t: 'Goals!!!', x: 9.047, y: 0.957, w: 2.131, h: 0.572, sz: 28, c: C.gray75, rot: 331 },
    { t: 'Goals!!!', x: 9.578, y: 2.401, w: 2.131, h: 0.572, sz: 28, c: C.gray85 },
    { t: 'Start!!!', x: 8.551, y: 2.050, w: 1.250, h: 0.404, sz: 18, c: C.orange, rot: 345, al: 'center' },
    { t: 'Success ', x: 10.500, y: 0.511, w: 2.475, h: 0.640, sz: 32, c: C.peach, b: true, al: 'center' },
    { t: 'Money', x: 11.926, y: 1.967, w: 1.250, h: 0.404, sz: 18, c: 'F0F0F0', rot: 12 }
  ].forEach(function (g) {
    text(s, g.t, g.x, g.y, g.w, g.h, {
      fontFace: MAJOR, fontSize: g.sz, bold: !!g.b, color: g.c, rotate: g.rot || 0,
      align: g.al || 'left', valign: 'middle'
    });
  });
  heading(s, TITLE, 1.300, 1.109, 4.828, 2.121, C.ink);
  para(s, LOREM_LONG, 1.300, 5.220, 4.655, 0.573, C.gray);
});

/* 25 - Gantt ---------------------------------------------------------- */
slide(function (s) {
  chrome(s, 25, 'dark');
  const gridX = 3.210, colW = 0.4543, days = 21, rowH = 0.9635, rowTop = 2.880;
  const rows = ['Awareness', 'Discussion', 'Collaboration', 'Text'];
  rect(s, 0.580, 1.990, 12.160, 0.050, C.gray95);
  s.addText('Project', Object.assign(box(0.580, 2.070, 2.630, 0.810), {
    shape: 'rect', fill: { color: C.gray95 }, line: { color: C.gray95, width: 2.25 },
    fontFace: MAJOR, fontSize: 14, color: C.gray, align: 'center', valign: 'middle'
  }));
  s.addText('August 2024', Object.assign(box(gridX, 2.040, 9.540, 0.370), {
    shape: 'rect', fill: { color: C.white }, line: { color: C.gray95, width: 2.25 },
    fontFace: MAJOR, fontSize: 10.5, bold: true, color: C.ink59, align: 'center', valign: 'middle'
  }));
  for (var d = 0; d < days; d++) {
    s.addText(String(d + 1), Object.assign(box(gridX + d * colW, 2.410, colW, 0.470), {
      shape: 'rect', fill: { color: d % 2 ? C.gray95 : C.shell }, line: { color: C.gray95, width: 2.25 },
      fontFace: MAJOR, fontSize: 14, color: C.gray66, align: 'center', valign: 'middle'
    }));
  }
  rows.forEach(function (name, r) {
    const y = rowTop + r * rowH;
    for (var d2 = 0; d2 < days; d2++) outline(s, gridX + d2 * colW, y, colW, rowH, C.gray95, 2.25);
    s.addText(name, Object.assign(box(0.580, y, 2.630, rowH), {
      shape: 'rect', fill: { color: C.white }, line: { color: C.gray95, width: 2.25 },
      fontFace: MAJOR, fontSize: 14, color: C.gray, align: 'center', valign: 'middle'
    }));
  });
  [
    { x: 7.250, w: 3.230, y: 3.220, fill: C.deep },
    { x: 3.670, w: 3.230, y: 4.210, fill: C.orange },
    { x: 9.110, w: 1.390, y: 4.210, fill: C.orange },
    { x: 4.270, w: 4.850, y: 5.190, fill: C.deep },
    { x: 3.670, w: 3.230, y: 6.140, fill: C.peach },
    { x: 8.190, w: 4.150, y: 6.140, fill: C.peach }
  ].forEach(function (b) { rect(s, b.x, b.y, b.w, 0.270, b.fill); });
  s.addShape('roundRect', Object.assign(box(8.660, 3.380, 3.570, 1.630), {
    fill: { color: C.white }, line: NONE, rectRadius: 0.05,
    shadow: { type: 'outer', color: '595959', opacity: 0.25, blur: 45, offset: 10, angle: 45 }
  }));
  subhead(s, 'Awareness', 8.960, 3.540, 2.980, 0.420, C.ink05);
  text(s, '90%', 8.990, 3.920, 3.040, 0.340, { fontFace: MAJOR, fontSize: 12, bold: true, color: C.teal, align: 'right' });
  roundRect(s, 9.100, 4.330, 2.810, 0.080, C.gray66, 0.04);
  roundRect(s, 9.100, 4.330, 2.460, 0.080, C.teal, 0.04);
  text(s, 'Impact', 8.990, 4.480, 3.040, 0.350, { fontSize: 12, color: C.ink });
  heading(s, TITLE, 1.600, 0.760, 10.050, 0.774, C.ink, { align: 'center' });
});

/* 26 - Project table --------------------------------------------------- */
slide(function (s) {
  chrome(s, 26, 'dark');
  const cols = [
    { key: 'no', x: 1.36, w: 0.72 }, { key: 'date', x: 2.08, w: 1.70 },
    { key: 'proj', x: 3.93, w: 1.88 }, { key: 'status', x: 6.17, w: 1.70 },
    { key: 'value', x: 8.28, w: 1.70 }, { key: 'loc', x: 10.35, w: 1.70 }
  ];
  const head = { no: 'No', date: 'Date', proj: 'Projects', status: 'Status', value: 'Value', loc: 'Location' };
  const rows = [
    { no: '01', date: '16 Feb. 2024', proj: 'Real Estate', pct: '80% ', value: '$420.600,00', loc: 'London, UK' },
    { no: '02', date: '22 Mar. 2024', proj: 'Real Estate', pct: '45% ', value: '$33.400,00', loc: 'Paris, FR' },
    { no: '03', date: '22 Mar. 2024', proj: 'Real Estate', pct: '45% ', value: '$33.400,00', loc: 'Paris, FR', on: true },
    { no: '04', date: '16 Feb. 2024', proj: 'Real Estate', pct: '80% ', value: '$420.600,00', loc: 'London, UK' },
    { no: '05', date: '22 Mar. 2024', proj: 'Real Estate', pct: '45% ', value: '$33.400,00', loc: 'Paris, FR' }
  ];
  s.addShape('roundRect', Object.assign(box(0.99, 2.54, 11.35, 3.85), {
    fill: { color: C.gray95 }, line: NONE, rectRadius: 0.16
  }));
  cols.forEach(function (c) { text(s, head[c.key], c.x, 2.72, c.w, 0.34, { fontFace: MAJOR, fontSize: 14, color: C.ink40 }); });
  rows.forEach(function (r, i) {
    const y = 3.21 + i * 0.6015;
    s.addShape('roundRect', Object.assign(box(1.20, y, 10.94, 0.53), {
      fill: { color: r.on ? C.teal : C.white }, line: NONE, rectRadius: 0.116,
      shadow: { type: 'outer', color: '000000', opacity: 0.15, blur: 36, offset: 0, angle: 90 }
    }));
    const plain = r.on ? C.white : C.gray, accent = r.on ? C.white : C.teal;
    const ty = y + 0.11;
    text(s, r.no, 1.36, ty, 0.72, 0.30, { fontFace: MAJOR, fontSize: 12, color: plain });
    text(s, r.date, 2.08, ty, 1.70, 0.30, { fontFace: MAJOR, fontSize: 12, color: plain });
    text(s, r.proj, 3.93, ty, 1.88, 0.30, { fontFace: MAJOR, fontSize: 12, color: plain });
    text(s, [{ text: r.pct, options: { color: accent } }, { text: '(progress)', options: { color: plain } }],
      6.17, ty, 1.70, 0.30, { fontFace: MAJOR, fontSize: 12 });
    text(s, r.value, 8.28, ty, 1.70, 0.30, { fontFace: MAJOR, fontSize: 12, color: accent });
    text(s, r.loc, 10.35, ty, 1.70, 0.30, { fontFace: MAJOR, fontSize: 12, color: plain });
  });
  heading(s, TITLE, 1.64, 1.11, 10.05, 0.774, C.ink, { align: 'center' });
});

/* 27 - Arrow columns --------------------------------------------------- */
slide(function (s) {
  chrome(s, 27, 'dark');
  const arrows = [
    { col: [6.91, 2.88, 0.50, 4.62], bar: [6.91, 2.38, 1.23, 0.50], head: [8.14, 2.14, 0.54, 0.99], dir: 'right', fill: C.deep, shade: C.deepD },
    { col: [6.30, 3.97, 0.51, 3.53], bar: [5.63, 3.46, 1.18, 0.51], head: [5.09, 3.21, 0.54, 0.99], dir: 'left', fill: C.orange, shade: C.orangeD },
    { col: [7.53, 4.51, 0.50, 2.99], bar: [7.53, 4.01, 1.08, 0.50], head: [8.61, 3.77, 0.54, 0.99], dir: 'right', fill: C.peach, shade: C.peachD },
    { col: [5.68, 5.52, 0.51, 1.98], bar: [5.27, 5.01, 0.92, 0.51], head: [4.72, 4.77, 0.55, 0.99], dir: 'left', fill: C.teal, shade: C.tealD }
  ];
  arrows.forEach(function (a) {
    rect(s, a.col[0], a.col[1], a.col[2], a.col[3], a.shade);
    rect(s, a.bar[0], a.bar[1], a.bar[2], a.bar[3], a.fill);
    shape(s, a.dir === 'right' ? 'rightArrow' : 'leftArrow', a.head[0] - (a.dir === 'right' ? 0.3 : 0.0), a.head[1], a.head[2] + 0.3, a.head[3], a.fill);
  });
  const labels = [
    { title: 'Valuable', tx: 2.92, tw: 1.48, ty: 3.46, bx: 1.94, by: 3.92, tri: [4.537, 3.565], triC: C.orange, dir: 'left', al: 'right' },
    { title: 'Company', tx: 9.44, tw: 1.87, ty: 2.42, bx: 9.44, by: 2.87, tri: [9.010, 2.517], triC: C.deep, dir: 'right', al: 'left' },
    { title: 'Organized', tx: 2.14, tw: 1.85, ty: 5.04, bx: 1.54, by: 5.51, tri: [4.134, 5.143], triC: C.teal, dir: 'left', al: 'right' },
    { title: 'Rate', tx: 9.91, tw: 1.38, ty: 4.09, bx: 9.91, by: 4.56, tri: [9.490, 4.192], triC: C.peach, dir: 'right', al: 'left' }
  ];
  labels.forEach(function (L) {
    shape(s, 'triangle', L.tri[0], L.tri[1], 0.25, 0.25, L.triC,
      { rotate: L.dir === 'left' ? 270 : 90, flipH: L.dir === 'left' });
    text(s, L.title, L.tx, L.ty, L.tw, 0.47, { fontFace: MAJOR, fontSize: 18, bold: true, color: C.ink, align: L.al });
    text(s, LOREM_MID, L.bx, L.by, 2.45, 0.68, { fontSize: 12, color: C.gray, align: 'right', lineSpacingMultiple: 1.5 });
  });
  heading(s, TITLE, 0.97, 1.11, 5.46, 1.45, C.ink);
});

/* 28 - Target ---------------------------------------------------------- */
slide(function (s) {
  chrome(s, 28, 'dark');
  [[5.55, 3.26, 2.22, C.deep], [5.79, 3.50, 1.75, 'F0F0F0'], [6.03, 3.73, 1.28, C.deep],
  [6.26, 3.97, 0.80, 'F0F0F0'], [6.44, 4.14, 0.45, C.deep]].forEach(function (r) {
    oval(s, r[0], r[1], r[2], r[2], r[3]);
  });
  s.addShape('rect', Object.assign(box(6.61, 3.55, 1.50, 0.09), { fill: { color: C.ink }, line: NONE, rotate: 315 }));
  shape(s, 'triangle', 7.40, 2.64, 0.61, 0.80, C.deepD, { rotate: 135 });
  shape(s, 'triangle', 7.59, 3.02, 0.80, 0.61, C.deep, { rotate: 135 });
  /* Four swoosh arrows circling the target, each tagged with a numbered dot.
     Quarter-turned arrows are authored upright then rotated about their centre,
     so their pre-rotation box has width and height swapped. */
  const arrows = [
    { cx: 5.09, cy: 3.78, w: 0.60, h: 1.36, rot: 0, flipH: false, c: C.teal, bx: 4.64, by: 4.39, n: '1' },
    { cx: 7.52, cy: 2.85, w: 0.53, h: 0.85, rot: 90, flipH: false, c: C.orange, bx: 6.68, by: 2.35, n: '2' },
    { cx: 8.24, cy: 4.98, w: 0.60, h: 1.35, rot: 180, flipH: false, c: C.peach, bx: 8.18, by: 3.86, n: '3' },
    { cx: 5.80, cy: 5.92, w: 0.52, h: 0.81, rot: 270, flipH: false, c: C.mint, bx: 6.14, by: 5.90, n: '4' }
  ];
  arrows.forEach(function (a) {
    shape(s, 'swooshArrow', a.cx - a.w / 2, a.cy - a.h / 2, a.w, a.h, a.c, { rotate: a.rot, flipH: a.flipH });
    oval(s, a.bx, a.by, 0.51, 0.51, a.c);
    text(s, a.n, a.bx, a.by, 0.51, 0.51, { fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle' });
  });
  [
    { n: '01', c: C.teal, bx: 0.99, by: 2.66, nx: 0.89, ny: 2.70, tx: 1.64, ty: 2.63, title: 'Valuable' },
    { n: '04', c: C.mint, bx: 0.99, by: 5.19, nx: 0.89, ny: 5.23, tx: 1.64, ty: 5.16, title: 'Organized' },
    { n: '02', c: C.orange, bx: 8.84, by: 2.66, nx: 8.74, ny: 2.70, tx: 9.49, ty: 2.63, title: 'Rate' },
    { n: '03', c: C.peach, bx: 8.84, by: 5.19, nx: 8.74, ny: 5.23, tx: 9.49, ty: 5.16, title: 'Company' }
  ].forEach(function (L) {
    oval(s, L.bx, L.by, 0.48, 0.48, L.c);
    text(s, L.n, L.nx, L.ny, 0.69, 0.37, { fontSize: 16, bold: true, color: C.white, align: 'center' });
    subhead(s, L.title, L.tx, L.ty, 2.95, 0.42, C.ink, { lineSpacingMultiple: 1 });
    para(s, LOREM_MID, L.tx, L.ty + 0.47, 2.85, 0.68, C.gray);
  });
  heading(s, TITLE, 1.05, 1.09, 11.24, 0.774, C.ink, { align: 'center' });
});

/* 29 - Contact --------------------------------------------------------- */
slide(function (s) {
  chrome(s, 29, 'dark');
  rect(s, 0.366, 0.521, 7.860, 6.458, C.teal);
  [
    { w: 4.26, tw: 3.53, tx: 1.46, y: 3.58, label: 'contact@wismrealestate.com' },
    { w: 3.60, tw: 2.90, tx: 1.44, y: 4.48, label: 'www.wismrealestate.com' },
    { w: 2.68, tw: 2.00, tx: 1.43, y: 5.38, label: '+1 234 567 8901' }
  ].forEach(function (p) {
    roundRect(s, 1.09, p.y, p.w, 0.71, C.white, 0.355);
    text(s, p.label, p.tx, p.y + 0.13, p.tw, 0.47, { fontSize: 16, color: C.ink });
  });
  heading(s, 'Get in Touch', 1.09, 1.41, 5.37, 0.774, C.white);
  para(s, LOREM_LONG, 1.09, 2.27, 4.49, 0.573, C.white);
  rect(s, 6.667, 0.521, 1.561, 6.458, C.orange);
  imageSlot(s, 8.227, 0.521, 4.738, 6.458);
});

/* 30 - Thank you -------------------------------------------------------- */
slide(function (s) {
  s.background = { color: C.teal };
  chrome(s, 30, 'brand');
  rect(s, -0.448, 2.181, 1.439, 4.024, C.orange);
  rect(s, 5.395, 3.167, 2.199, 3.038, C.orange);
  stat(s, 'Thank You', 1.297, 4.193, 5.597, 1.582, C.white, 88);
  text(s, 'Real Estate PowerPoint', 1.297, 5.633, 4.125, 0.572, { fontSize: 28, color: C.white, charSpacing: -1.5 });
  imageSlot(s, 7.096, 0.521, 5.870, 6.458);
});

/* ------------------------------------------------------------------ *
 * Build                                                               *
 * ------------------------------------------------------------------ */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WISM_16x9', width: 13.333, height: 7.5 });
pptx.layout = 'WISM_16x9';
pptx.title = 'Wism - Real Estate PowerPoint';
pptx.author = 'Wism';
pptx.company = 'Wism';
pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };

slides.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '0d45b59a-1512-4d6d-9f82-17c5f7ed8508_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
