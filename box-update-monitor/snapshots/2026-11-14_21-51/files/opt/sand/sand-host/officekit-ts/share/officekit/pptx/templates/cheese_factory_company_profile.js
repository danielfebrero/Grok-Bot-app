#!/usr/bin/env node
/**
 * "Cheese - Food Presentation Template" — 10 slides, 13.333 x 7.5 in (16:9).
 *
 * Recreated from scratch with pptxgenjs. The reference deck's photographs are
 * stand-in stock images, so each one is drawn here as a flat grey rectangle
 * labelled "[image]" at the original position and size, and each freeform
 * vector icon is rebuilt out of native pptxgenjs shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const RED = 'F84620'; // theme accent1 / accent3
const YELLOW = 'FFC154'; // theme accent2
const DARK = '262626'; // theme tx1 @ lumMod 85% / lumOff 15%
const WHITE = 'FFFFFF'; // theme bg1

const PHOTO_FILL = 'CCCCCC';
const PHOTO_LABEL = 'B0B0B0';

const HEAD = 'Poppins'; // headings + logo
const BODY = 'Work Sans'; // everything else

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';

const NO_LINE = { type: 'none' };

/* ---------------------------------------------------------------- helpers */

/** Flat shape, no outline. */
function shape(slide, kind, x, y, w, h, color, opts) {
  slide.addShape(kind, Object.assign({ x, y, w, h, fill: { color }, line: NO_LINE }, opts));
}

function rect(slide, x, y, w, h, color, opts) {
  shape(slide, 'rect', x, y, w, h, color, opts);
}

function oval(slide, x, y, w, h, color) {
  shape(slide, 'ellipse', x, y, w, h, color);
}

/** Outlined shape with no fill — used by a few of the line-art icons. */
function stroke(slide, kind, x, y, w, h, color, pt) {
  slide.addShape(kind, {
    x, y, w, h, fill: { type: 'none' }, line: { color, width: pt || 1.25 },
  });
}

/** A straight stroke from (x1,y1) to (x2,y2), drawn as a rotated bar. */
function segment(slide, x1, y1, x2, y2, thickness, color) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const len = Math.sqrt(dx * dx + dy * dy);
  rect(slide, (x1 + x2) / 2 - len / 2, (y1 + y2) / 2 - thickness / 2, len, thickness, color,
    { rotate: (Math.atan2(dy, dx) * 180) / Math.PI });
}

/** "<" / ">" carousel arrow: two strokes meeting at an apex. */
function chevron(slide, x, y, w, h, color, pointsRight) {
  const apexX = pointsRight ? x + w : x;
  const backX = pointsRight ? x : x + w;
  segment(slide, backX, y, apexX, y + h / 2, 0.042, color);
  segment(slide, backX, y + h, apexX, y + h / 2, 0.042, color);
}

/** Placeholder standing in for a photograph in the reference deck. */
function photo(slide, x, y, w, h, labelX, labelW) {
  rect(slide, x, y, w, h, PHOTO_FILL);
  slide.addText('[image]', {
    x: labelX === undefined ? x : labelX, y, w: labelW === undefined ? w : labelW, h,
    fontFace: BODY, fontSize: 12, color: PHOTO_LABEL, align: 'center', valign: 'middle',
  });
}

/**
 * Title-slide hero: the photo carries a "white -> transparent" gradient veil,
 * reproduced as a left-to-right ramp of solid grey-to-white bands.
 */
function photoFade(slide, x, y, w, h, solidFrac) {
  const bands = 26;
  const x0 = x + w * solidFrac;
  const bw = (x + w - x0) / bands;
  rect(slide, x, y, w * solidFrac, h, PHOTO_FILL);
  for (let i = 0; i < bands; i++) {
    const t = (i + 1) / bands; // 0 = photo grey, 1 = pure white
    const v = Math.round(0xcc + (0xff - 0xcc) * t);
    rect(slide, x0 + i * bw, y, bw + 0.008, h, v.toString(16).repeat(3));
  }
  slide.addText('[image]', {
    x, y, w: w * solidFrac, h, fontFace: BODY, fontSize: 12,
    color: PHOTO_LABEL, align: 'center', valign: 'middle',
  });
}

/** Body copy: 10.5pt Work Sans, justified, 1.5 line spacing. */
function body(slide, x, y, w, h, text, color) {
  slide.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: 10.5, color: color || DARK,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
}

/** Numbered sub-heading: 18pt bold Work Sans, single line. */
function label(slide, x, y, w, h, text, color) {
  slide.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: 18, bold: true, color, valign: 'top', wrap: false,
  });
}

/** Slide heading: bold Poppins built from [text, color] run pairs. */
function heading(slide, x, y, w, h, runs, fontSize) {
  slide.addText(runs.map(([text, color]) => ({ text, options: { color } })), {
    x, y, w, h, fontFace: HEAD, fontSize: fontSize || 40, bold: true, valign: 'top',
  });
}

/* ------------------------------------------------------------------ icons */

/**
 * Each icon is drawn from native shapes inside a 1"-square badge, positioned
 * relative to the badge centre (cx, cy). Sizes are in inches.
 */
const ICONS = {
  // Slide 2 — binoculars
  binoculars(s, cx, cy, c) {
    rect(s, cx - 0.20, cy - 0.16, 0.11, 0.10, c);
    rect(s, cx + 0.09, cy - 0.16, 0.11, 0.10, c);
    shape(s, 'roundRect', cx - 0.22, cy - 0.08, 0.16, 0.26, c, { rectRadius: 0.05 });
    shape(s, 'roundRect', cx + 0.06, cy - 0.08, 0.16, 0.26, c, { rectRadius: 0.05 });
    rect(s, cx - 0.05, cy - 0.11, 0.10, 0.22, c);
  },

  // Slide 3 — rising chart: L-shaped axis plus a zig-zagging trend arrow
  chart(s, cx, cy, c) {
    rect(s, cx - 0.19, cy - 0.17, 0.026, 0.34, c);
    rect(s, cx - 0.19, cy + 0.147, 0.38, 0.026, c);
    const trend = [
      [cx - 0.13, cy + 0.07], [cx - 0.06, cy - 0.01],
      [cx - 0.005, cy + 0.04], [cx + 0.10, cy - 0.09],
    ];
    trend.slice(1).forEach((pt, i) => segment(s, trend[i][0], trend[i][1], pt[0], pt[1], 0.032, c));
    shape(s, 'triangle', cx + 0.055, cy - 0.155, 0.13, 0.13, c, { rotate: 45 });
  },

  // Slide 4 — facility building
  bank(s, cx, cy, c) {
    shape(s, 'triangle', cx - 0.21, cy - 0.20, 0.42, 0.13, c);
    [-0.16, -0.03, 0.10].forEach((dx) => rect(s, cx + dx, cy - 0.05, 0.06, 0.17, c));
    rect(s, cx - 0.21, cy + 0.14, 0.42, 0.06, c);
  },

  // Slide 6 — team
  people(s, cx, cy, c) {
    [-0.17, 0.17].forEach((dx) => {
      oval(s, cx + dx - 0.06, cy - 0.15, 0.12, 0.12, c);
      shape(s, 'roundRect', cx + dx - 0.09, cy - 0.02, 0.18, 0.16, c, { rectRadius: 0.05 });
    });
    oval(s, cx - 0.09, cy - 0.16, 0.18, 0.18, c);
    shape(s, 'roundRect', cx - 0.15, cy + 0.01, 0.30, 0.19, c, { rectRadius: 0.06 });
  },

  // Slide 7 — coffee cup
  coffee(s, cx, cy, c) {
    shape(s, 'donut', cx + 0.06, cy - 0.13, 0.20, 0.20, c);
    shape(s, 'roundRect', cx - 0.20, cy - 0.15, 0.28, 0.22, c, { rectRadius: 0.06 });
    shape(s, 'roundRect', cx - 0.22, cy + 0.10, 0.44, 0.07, c, { rectRadius: 0.03 });
  },

  // Slide 7 — fork and knife
  cutlery(s, cx, cy, c) {
    [-0.14, -0.08, -0.02].forEach((dx) => rect(s, cx + dx, cy - 0.20, 0.03, 0.13, c));
    rect(s, cx - 0.14, cy - 0.09, 0.15, 0.05, c);
    rect(s, cx - 0.08, cy - 0.06, 0.04, 0.26, c);
    shape(s, 'roundRect', cx + 0.08, cy - 0.20, 0.09, 0.20, c, { rectRadius: 0.04 });
    rect(s, cx + 0.11, cy - 0.04, 0.04, 0.24, c);
  },

  // Slide 8 — graduation cap: mortarboard over a tapered crown, plus a tassel
  cap(s, cx, cy, c) {
    shape(s, 'trapezoid', cx - 0.11, cy - 0.03, 0.22, 0.15, c, { rotate: 180 });
    shape(s, 'diamond', cx - 0.24, cy - 0.16, 0.48, 0.22, c);
    rect(s, cx + 0.13, cy - 0.05, 0.025, 0.15, c);
    oval(s, cx + 0.115, cy + 0.09, 0.055, 0.075, c);
  },

  // Slide 9 — chat bubble
  chat(s, cx, cy, c) {
    shape(s, 'wedgeEllipseCallout', cx - 0.21, cy - 0.17, 0.42, 0.34, c);
  },

  // Slide 10 — heart
  heart(s, cx, cy, c) {
    shape(s, 'heart', cx - 0.19, cy - 0.17, 0.38, 0.34, c);
  },

  // Slide 9 contact list — small line-art marks
  phone(s, cx, cy, c) {
    stroke(s, 'roundRect', cx - 0.07, cy - 0.11, 0.14, 0.22, c, 1.5);
    rect(s, cx - 0.03, cy + 0.05, 0.06, 0.02, c);
  },
  mail(s, cx, cy, c) {
    stroke(s, 'roundRect', cx - 0.11, cy - 0.08, 0.22, 0.16, c, 1.5);
    segment(s, cx - 0.09, cy - 0.06, cx, cy + 0.01, 0.022, c);
    segment(s, cx, cy + 0.01, cx + 0.09, cy - 0.06, 0.022, c);
  },
  globe(s, cx, cy, c) {
    stroke(s, 'ellipse', cx - 0.10, cy - 0.10, 0.20, 0.20, c, 1.5);
    stroke(s, 'ellipse', cx - 0.04, cy - 0.10, 0.08, 0.20, c, 1);
    rect(s, cx - 0.10, cy - 0.01, 0.20, 0.02, c);
  },
  // Hollow map pin — the "hole" is punched with the slide's own background colour
  pin(s, cx, cy, c, hole) {
    shape(s, 'teardrop', cx - 0.09, cy - 0.12, 0.19, 0.19, c, { rotate: 135 });
    oval(s, cx - 0.031, cy - 0.073, 0.062, 0.062, hole);
  },
};

/** The recurring 1"-wide circular icon badge. */
function badge(slide, x, y, circleColor, iconName, iconColor) {
  oval(slide, x, y, 1, 1, circleColor);
  ICONS[iconName](slide, x + 0.5, y + 0.5, iconColor);
}

/* --------------------------------------------------------------- nav bar  */

/**
 * Website-style navigation bar repeated on every slide.
 *  dy    – the group sits 0.071" lower on slides 2-10 than on the title slide
 *  links – x positions of "Home" / "Services" / "About Us"
 */
function navBar(slide, o) {
  const dy = o.dy;
  const linkW = [0.623, 0.8, 0.863];

  oval(slide, 1.178, 0.293 + dy, 0.422, 0.422, o.oval);
  slide.addText('Cheese Factory', {
    x: 1.24, y: 0.251 + dy, w: 1.19, h: 0.505,
    fontFace: HEAD, fontSize: 12, bold: true, color: o.brand, valign: 'top',
  });

  ['Home', 'Services', 'About Us'].forEach((text, i) => {
    slide.addText(text, {
      x: o.links[i], y: 0.365 + dy, w: linkW[i], h: 0.278,
      fontFace: BODY, fontSize: 10.5, color: DARK, valign: 'top', wrap: false,
    });
  });
  slide.addText('Log in', {
    x: 10.007, y: 0.365 + dy, w: 0.637, h: 0.278,
    fontFace: BODY, fontSize: 10.5, color: o.login || DARK, valign: 'top', wrap: false,
  });

  rect(slide, 10.8, 0.342 + dy, 1.393, 0.325, o.button);
  slide.addText('Order Now', {
    x: 10.8, y: 0.342 + dy, w: 1.393, h: 0.325,
    fontFace: BODY, fontSize: 10.5, color: o.buttonText,
    align: 'center', valign: 'middle', lineSpacingMultiple: 1,
  });
}

/** Inner-page nav: identical on most slides, only the link x positions move. */
function pageNav(slide, links, extra) {
  navBar(slide, Object.assign({
    dy: 0.071, links: links || [3.492, 4.466, 5.616],
    oval: YELLOW, brand: DARK, button: YELLOW, buttonText: DARK,
  }, extra));
}

/* ----------------------------------------------------------------- slides */

// 1 — Title / hero
function slide1(pres) {
  const s = pres.addSlide();
  photo(s, 6.667, 5.344, 6.667, 2.156);
  photoFade(s, 0, 0.944, 9.011, 2.806, 0.37);

  rect(s, 0.9, 3.438, 1.074, 1.421, RED);
  rect(s, 1.178, 3.133, 5.489, 1.421, YELLOW);
  chevron(s, 1.894, 3.729, 0.14, 0.23, DARK, false);
  chevron(s, 5.81, 3.729, 0.14, 0.23, DARK, true);

  heading(s, 7.534, 1.296, 4.781, 1.531, [['Cheese', RED]], 85);
  s.addText('Food Presentation Template', {
    x: 7.644, y: 2.655, w: 4.669, h: 0.404,
    fontFace: BODY, fontSize: 18, color: DARK, valign: 'top',
  });

  rect(s, 2.874, 0, 10.459, 0.944, YELLOW);
  rect(s, 0, 0, 2.874, 0.944, RED);
  navBar(s, {
    dy: 0, links: [3.492, 4.466, 5.616],
    oval: YELLOW, brand: WHITE, button: RED, buttonText: WHITE,
  });

  rect(s, 11.328, 4.878, 1.074, 1.421, RED);
  rect(s, 7.763, 4.634, 4.43, 1.421, YELLOW);
  rect(s, 0, 5.344, 6.667, 2.156, RED);

  body(s, 0.712, 5.792, 5.243, 1.132,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut eiusmod tempor incididunt ut consectetur adipiscing adipiscing elit, ' +
    'PLACEHOLDER', WHITE);
  body(s, 8.304, 5.044, 3.347, 0.601,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod');

  photo(s, 2.861, 3.133, 2.122, 1.421);
  s.addNotes('1');
}

// 2 — Welcome To Cheese
function slide2(pres) {
  const s = pres.addSlide();
  photo(s, 1.182, 1.391, 4.716, 4.703);
  pageNav(s);

  heading(s, 7.831, 1.259, 4.548, 1.447, [['Welcome To ', RED], ['Cheese', YELLOW]]);
  label(s, 7.881, 3.424, 2.621, 0.404, 'About Our Planning', YELLOW);
  body(s, 7.848, 4.496, 4.345, 1.662,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'PLACEHOLDER' +
    'PLACEHOLDER' +
    'eiusmod tempor incididunt ut eiusmod tempor');

  badge(s, 5.365, 4.243, YELLOW, 'binoculars', DARK);
  s.addNotes('2');
}

// 3 — Vision & Mission
function slide3(pres) {
  const s = pres.addSlide();
  photo(s, 7.683, 1.292, 5.65, 6.208);

  heading(s, 1.107, 1.292, 4.977, 0.774, [['Vision & ', RED], ['Mission', YELLOW]]);

  [
    { y: 2.454, title: '1.  Vision', titleX: 1.569, titleY: 2.737, titleW: 1.289, textX: 1.6, textY: 3.196 },
    { y: 4.472, title: '2. Mission', titleX: 1.551, titleY: 4.771, titleW: 1.452, textX: 1.582, textY: 5.23 },
  ].forEach((card) => {
    rect(s, 1.178, card.y, 4.731, 1.626, RED);
    label(s, card.titleX, card.titleY, card.titleW, 0.404, card.title, WHITE);
    body(s, card.textX, card.textY, 4.016, 0.601, LOREM, WHITE);
  });

  pageNav(s);
  badge(s, 7.183, 3.896, YELLOW, 'chart', DARK);
  s.addNotes('3');
}

// 4 — Our Facilities
function slide4(pres) {
  const s = pres.addSlide();
  photo(s, 0, 4.598, 13.333, 2.902, 7.0, 6.3);
  rect(s, 2.537, 0, 4.13, 7.517, YELLOW);

  [
    { title: '1.  Facilities', titleY: 1.428, titleW: 1.666, textY: 2.106 },
    { title: '2.  Facilities', titleY: 3.165, titleW: 1.708, textY: 3.783 },
    { title: '3.  Facilities', titleY: 4.842, titleW: 1.706, textY: 5.52 },
  ].forEach((item) => {
    label(s, 3.424, item.titleY, item.titleW, 0.404, item.title, RED);
    body(s, 3.424, item.textY, 2.449, 0.601,
      'Lorem ipsum dolor sit amet, consectetur adipiscing, ');
  });

  heading(s, 7.66, 1.259, 4.806, 0.774, [['Our ', RED], ['Facilities', YELLOW]]);
  body(s, 7.672, 2.586, 4.592, 1.132,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'PLACEHOLDER' +
    'tempor incididunt ut eiusmod tempor incididunt');

  pageNav(s, [3.112, 4.086, 5.237]);
  badge(s, 0.678, 4.08, RED, 'bank', WHITE);
  s.addNotes('4');
}

// 5 — Our Best Project
function slide5(pres) {
  const s = pres.addSlide();
  rect(s, 9.204, 0, 4.13, 7.517, YELLOW);
  pageNav(s, null, { button: RED, buttonText: WHITE });

  heading(s, 1.091, 1.25, 4.977, 0.774, [['Our Best ', RED], ['Project', YELLOW]]);

  [
    { title: '1.  Project Title', titleX: 1.08, titleW: 2.078, titleY: 2.674, textX: 1.111, textY: 3.241 },
    { title: '2. Project Title', titleX: 1.062, titleW: 2.048, titleY: 4.726, textX: 1.094, textY: 5.294 },
  ].forEach((item) => {
    label(s, item.titleX, item.titleY, item.titleW, 0.404, item.title, RED);
    body(s, item.textX, item.textY, 4.802, 0.867,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'PLACEHOLDER');
  });

  photo(s, 7.417, 1.391, 4.716, 4.703);
  s.addNotes('5');
}

// 6 — Meet Our Team
function slide6(pres) {
  const s = pres.addSlide();
  photo(s, 0, 4.556, 5.111, 2.944);
  photo(s, 10.138, 3.829, 2.052, 2.222);
  photo(s, 6.667, 1.449, 2.052, 2.222);
  pageNav(s);

  heading(s, 1.087, 1.285, 4.548, 0.774, [['Meet Our ', RED], ['Team', YELLOW]]);
  body(s, 1.143, 2.688, 4.345, 0.867,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'PLACEHOLDER');

  rect(s, 8.859, 1.449, 3.331, 2.222, RED);
  rect(s, 6.667, 3.829, 3.331, 2.222, RED, { flipH: true });

  badge(s, 4.681, 5.528, YELLOW, 'people', DARK);

  [
    { name: 'Hannah Morales', x: 6.985, nameY: 4.245, nameW: 2.218, textY: 4.756 },
    { name: 'Claudia Alves', x: 9.151, nameY: 1.845, nameW: 1.881, textY: 2.356 },
  ].forEach((member) => {
    label(s, member.x, member.nameY, member.nameW, 0.404, member.name, YELLOW);
    body(s, member.x, member.textY, 2.753, 0.867, LOREM, WHITE);
  });
  s.addNotes('6');
}

// 7 — Our Best Service
function slide7(pres) {
  const s = pres.addSlide();
  photo(s, 9.03, 3.75, 4.303, 3.75);
  photo(s, 0, 0, 3.492, 7.5);
  rect(s, 9.03, 0, 4.303, 3.75, RED);
  pageNav(s, [4.832, 5.806, 6.956], { login: WHITE });

  heading(s, 9.59, 1.286, 2.848, 1.447, [['Our Best ', WHITE], ['Service', YELLOW]]);

  [
    { icon: 'coffee', badgeX: 2.979, badgeY: 1.85, cardY: 1.399, title: '1. Service', titleY: 1.618, titleW: 1.378, textY: 2.108 },
    { icon: 'cutlery', badgeX: 2.962, badgeY: 4.65, cardY: 4.199, title: '2. Service', titleY: 4.439, titleW: 1.42, textY: 4.928 },
  ].forEach((row) => {
    badge(s, row.badgeX, row.badgeY, RED, row.icon, WHITE);
    rect(s, 4.3, row.cardY, 4.042, 1.902, YELLOW, { flipH: true });
    label(s, 4.616, row.titleY, row.titleW, 0.404, row.title, RED);
    body(s, 4.616, row.textY, 3.415, 0.867,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'incididunt ut eiusmod tempor');
  });
  s.addNotes('7');
}

// 8 — Our Portfolio
function slide8(pres) {
  const s = pres.addSlide();
  photo(s, 7.335, 3.928, 4.847, 2.169);
  photo(s, 1.166, 3.928, 2.326, 2.169);
  photo(s, 3.799, 1.418, 3.229, 4.679);
  photo(s, 1.166, 1.418, 2.326, 2.169);
  pageNav(s);

  heading(s, 8.135, 1.244, 4.317, 0.774, [['Our Portfolio', RED]]);
  body(s, 8.135, 2.487, 3.999, 0.867,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut eiusmod tempor incididunt ut consectetur');

  badge(s, 3.115, 3.25, RED, 'cap', WHITE);
  s.addNotes('8');
}

// 9 — Contact Us
function slide9(pres) {
  const s = pres.addSlide();
  rect(s, 0, 3.75, 13.333, 3.75, RED);
  photo(s, 9.729, 2.821, 3, 4.679);
  photo(s, 6.188, 1.41, 3, 4.679);
  pageNav(s);

  heading(s, 1.086, 1.248, 4.362, 0.774,
    [['Contact', RED], [' ', DARK], ['Us', YELLOW]]);
  body(s, 1.068, 2.389, 3.397, 0.601,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, eiusmod incididunt');

  [
    { icon: 'phone', iconY: 4.655, text: '+12345678910', textY: 4.473, w: 1.633 },
    { icon: 'mail', iconY: 5.323, text: 'yourmail@gmail.com', textY: 5.192, w: 1.953 },
    { icon: 'globe', iconY: 5.986, text: 'www.yourwebsite.com', textY: 5.883, w: 1.953 },
    { icon: 'pin', iconY: 6.664, text: '123 Anywhere ST,, Any City', textY: 6.581, w: 2.259 },
  ].forEach((row) => {
    ICONS[row.icon](s, 1.38, row.iconY, YELLOW, RED);
    s.addText(row.text, {
      x: 1.756, y: row.textY, w: row.w, h: 0.278,
      fontFace: BODY, fontSize: 10.5, color: WHITE, align: 'justify', valign: 'top',
    });
  });

  badge(s, 5.719, 3.25, YELLOW, 'chat', DARK);
  s.addNotes('9');
}

// 10 — Thank You
function slide10(pres) {
  const s = pres.addSlide();
  rect(s, 0, 0, 4.13, 7.5, YELLOW);
  photo(s, 1.178, 1.431, 4.697, 6.069);
  pageNav(s, [5.448, 6.421, 7.572], { oval: RED });

  heading(s, 7.335, 1.213, 5.115, 2.962, [['THANK ', RED], ['YOU!', YELLOW]], 85);
  body(s, 7.361, 4.983, 4.832, 1.132,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut eiusmod tempor incididunt consectetur adipiscing elit, sed do eiusmod ' +
    'tempor incididunt ut eiusmod tempor incididunt euismod');

  badge(s, 5.375, 4.175, RED, 'heart', WHITE);
  s.addNotes('10');
}

/* ------------------------------------------------------------------- main */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'CHEESE_16x9', width: 13.333333, height: 7.5 });
  pres.layout = 'CHEESE_16x9';
  pres.author = 'Cheese Factory';
  pres.title = 'Cheese - Food Presentation Template';

  [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
    .forEach((fn) => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '08db4cd8-750f-41b6-887b-2a66a4205811_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
