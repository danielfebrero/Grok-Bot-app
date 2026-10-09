/**
 * "Travelur" travel presentation template - 35 slides, 20 x 11.25 in.
 * Rebuilt with pptxgenjs only. Photographs / raster mock-ups from the original
 * deck are replaced by flat colour placeholders; vector icons are redrawn with
 * native pptx shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 20;
const SLIDE_H = 11.25;

/* ------------------------------------------------------------------ palette */
const C = {
  teal: '037673',        // primary brand green
  dark: '033A3B',        // deep green used for washes / overlays
  orange: 'F8930F',      // accent
  white: 'FFFFFF',
  grey: 'A6A6A6',        // body copy (white @ 65% luminance in the original)
  greyLt: 'F2F2F2',      // body copy on dark slides (white @ 95%)
  photo: 'F7F9F9',       // stand-in fill for photo placeholders
  photoTx: 'CBD3D3',     // caption colour inside a photo placeholder
  ink: '111111'          // icon-sheet glyphs
};

const F = { head: 'Raleway ExtraBold', semi: 'Poppins SemiBold', body: 'Poppins' };

/* Big soft card shadow used throughout the original deck.
   Returns a fresh object each call: pptxgenjs rewrites shadow options in place
   when it renders, so a shared literal would compound across shapes. */
const shadow = () => ({ type: 'outer', color: '000000', opacity: 0.1, blur: 42, offset: 9, angle: 128 });

const NOLINE = { type: 'none' };

/** roundRect "adj" value (0-100000, as stored in the source deck) -> radius in inches. */
const rr = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/* ------------------------------------------------------------------ helpers */

/** Text box. Defaults mirror PowerPoint: top aligned, 18pt Poppins. */
function text(s, str, o) {
  s.addText(str, Object.assign({ fontFace: F.body, fontSize: 18, color: C.grey, valign: 'top' }, o));
}

/** Small orange kicker above a headline. */
function eyebrow(s, x, y, str, opt) {
  text(s, str || 'TRAVEL TEMPLATE', Object.assign(
    { x, y, w: 3.365, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.orange, charSpacing: 3 }, opt));
}

/** Raleway ExtraBold headline. */
function title(s, o) {
  text(s, o.text, { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: F.head, fontSize: o.size || 44, color: o.color || C.teal, charSpacing: o.spc });
}

/** Justified 150% body paragraph. */
function para(s, o) {
  text(s, o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontSize: o.size || 14, color: o.color || C.grey,
    align: o.align || 'justify', lineSpacingMultiple: 1.5
  });
}

/** Rounded pill button with centred white label. */
function button(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: rr(11007, o.w, o.h),
    fill: { color: o.fill || C.orange }, line: NOLINE
  });
  text(s, o.text, { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: F.semi, fontSize: o.size || 16, color: C.white, align: 'center', valign: 'middle' });
}

/** Composite `fg` over `bg` at the given transparency (0-100) -> opaque hex. */
function blend(bg, fg, transparency) {
  const a = 1 - transparency / 100;
  const ch = (i) => {
    const b = parseInt(bg.substr(i * 2, 2), 16), f = parseInt(fg.substr(i * 2, 2), 16);
    return Math.round(b + (f - b) * a).toString(16).padStart(2, '0');
  };
  return (ch(0) + ch(1) + ch(2)).toUpperCase();
}

/**
 * Linear fade, approximated with a stack of flat slices (pptxgenjs has no
 * gradient fill). Slices are pre-blended against `bg` so they stay opaque and
 * no seams show between them. `stops` are [position 0..1 along `dir`,
 * transparency %]; dir 'r' = left to right, 'u' = bottom to top.
 */
function fade(s, o) {
  const stops = o.stops;
  const steps = o.steps || 26;
  const at = (p) => {
    for (let i = 1; i < stops.length; i++) {
      if (p <= stops[i][0]) {
        const [p0, t0] = stops[i - 1], [p1, t1] = stops[i];
        return t0 + (t1 - t0) * (p1 === p0 ? 0 : (p - p0) / (p1 - p0));
      }
    }
    return stops[stops.length - 1][1];
  };
  for (let i = 0; i < steps; i++) {
    const p = (i + 0.5) / steps;
    const box = o.dir === 'r'
      ? { x: o.x + (o.w * i) / steps, y: o.y, w: (o.w / steps) * 1.05, h: o.h }
      : { x: o.x, y: o.y + o.h - (o.h * (i + 1)) / steps, w: o.w, h: (o.h / steps) * 1.05 };
    s.addShape('rect', Object.assign(box, { fill: { color: blend(o.bg || C.photo, o.color || C.dark, at(p)) }, line: NOLINE }));
  }
}

/** Dark-to-transparent wash used at the foot of every photo card. */
function photoWash(s, o) {
  fade(s, { x: o.x, y: o.y, w: o.w, h: o.h, dir: 'u', stops: [[0, 0], [0.06, 0], [1, 100]], steps: o.steps || 22 });
}

/** Flat stand-in for a photograph. */
function photo(s, o) {
  s.addShape(o.shape || 'rect', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill || C.photo }, line: NOLINE,
    rectRadius: o.radius, rotate: o.rotate
  });
  if (o.label !== false) {
    text(s, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.3, w: o.w, h: 0.6, fontFace: F.semi, fontSize: 16, color: o.labelColor || C.photoTx, align: 'center', valign: 'middle' });
  }
}

/* ------------------------------------------------------------------- icons */
/* The original uses SVG line icons; these are compact native-shape stand-ins. */

const ring = (s, x, y, d, c, w) => s.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: c, width: w || 1.25 } });
const dot = (s, x, y, d, c) => s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: c }, line: NOLINE });

const ICONS = {
  /* map pin + dotted route */
  pin(s, x, y, d, c) {
    ring(s, x + 0.30 * d, y, 0.40 * d, c);
    dot(s, x + 0.44 * d, y + 0.12 * d, 0.12 * d, c);
    s.addShape('triangle', { x: x + 0.40 * d, y: y + 0.36 * d, w: 0.20 * d, h: 0.16 * d, flipV: true, fill: { color: c }, line: NOLINE });
    ring(s, x, y + 0.30 * d, 0.34 * d, c);
    for (let i = 0; i < 4; i++) dot(s, x + (0.40 + 0.15 * i) * d, y + 0.74 * d, 0.07 * d, c);
  },
  /* pair of travel tickets */
  ticket(s, x, y, d, c) {
    s.addShape('parallelogram', { x, y: y + 0.16 * d, w: 0.9 * d, h: 0.46 * d, fill: { type: 'none' }, line: { color: c, width: 1.25 } });
    s.addShape('parallelogram', { x: x + 0.1 * d, y: y + 0.40 * d, w: 0.9 * d, h: 0.40 * d, fill: { type: 'none' }, line: { color: c, width: 1.25 } });
    s.addShape('rightArrow', { x: x + 0.34 * d, y: y + 0.50 * d, w: 0.34 * d, h: 0.18 * d, fill: { color: c }, line: NOLINE });
  },
  /* cruise ship */
  ship(s, x, y, d, c) {
    s.addShape('trapezoid', { x: x + 0.16 * d, y: y + 0.14 * d, w: 0.68 * d, h: 0.26 * d, fill: { type: 'none' }, line: { color: c, width: 1.25 } });
    s.addShape('trapezoid', { x, y: y + 0.44 * d, w: d, h: 0.34 * d, flipV: true, fill: { type: 'none' }, line: { color: c, width: 1.25 } });
    s.addShape('rect', { x: x + 0.08 * d, y: y + 0.40 * d, w: 0.84 * d, h: 0.06 * d, fill: { color: c }, line: NOLINE });
  },
  /* hot air balloon */
  balloon(s, x, y, d, c) {
    ring(s, x + 0.12 * d, y, 0.76 * d, c);
    s.addShape('trapezoid', { x: x + 0.36 * d, y: y + 0.74 * d, w: 0.28 * d, h: 0.22 * d, flipV: true, fill: { type: 'none' }, line: { color: c, width: 1.25 } });
  },
  /* filled tick used by the feature lists */
  check(s, x, y, d, c) {
    dot(s, x, y, d, c);
    text(s, '\u2713', { x, y, w: d, h: d, fontFace: F.semi, fontSize: Math.round(d * 46), bold: true, color: C.white, align: 'center', valign: 'middle' });
  },
  /* small clock / eye markers on the pricing cards */
  clock(s, x, y, d, c) {
    ring(s, x, y, d, c, 1);
    s.addShape('rect', { x: x + 0.48 * d, y: y + 0.22 * d, w: 0.05 * d, h: 0.3 * d, fill: { color: c }, line: NOLINE });
  },
  eye(s, x, y, d, c) {
    s.addShape('moon', { x, y: y + 0.15 * d, w: 0.55 * d, h: 0.7 * d, rotate: 270, fill: { color: c }, line: NOLINE });
  },
  /* envelope + phone tiles on the contact slide */
  mail(s, x, y, d, c) {
    s.addShape('rect', { x, y: y + 0.12 * d, w: d, h: 0.72 * d, fill: { color: c }, line: { color: C.orange, width: 1 } });
    s.addShape('triangle', { x: x + 0.06 * d, y: y + 0.16 * d, w: 0.88 * d, h: 0.5 * d, flipV: true, fill: { color: C.orange }, line: { color: c, width: 1 } });
  },
  phone(s, x, y, d, c) {
    ring(s, x + 0.2 * d, y, 0.6 * d, c, 1.5);
    s.addShape('trapezoid', { x, y: y + 0.6 * d, w: d, h: 0.34 * d, fill: { color: c }, line: NOLINE });
  }
};

function icon(s, kind, x, y, d, color) { ICONS[kind](s, x, y, d, color || C.orange); }

/* ------------------------------------------------- repeated content blocks */

/** Booking bar fields: bold label, grey value and a small orange caret. */
function bookingFields(s, o) {
  o.fields.forEach((f) => {
    text(s, f.label, { x: f.x, y: o.y, w: 1.75, h: 0.4, fontFace: F.semi, fontSize: o.labelSize || 18, color: o.labelColor || C.teal });
    text(s, f.value, { x: f.x, y: o.y + o.gap, w: 1.75, h: 0.31, fontFace: F.semi, fontSize: o.valueSize || 12, color: C.grey });
    s.addShape('triangle', { x: f.x + f.caret, y: o.y + o.caretDy, w: 0.09, h: 0.069, flipV: true, fill: { color: C.orange }, line: NOLINE });
  });
}

/** White booking bar (3 fields + Book Now) used on the cover slides. */
function bookingBar(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: 9.455, h: 1.011, rectRadius: rr(4473, 9.455, 1.011),
    fill: { color: C.white, transparency: o.barTransparency || 0 }, line: NOLINE
  });
  bookingFields(s, {
    y: o.y + 0.2127, gap: 0.3089, caretDy: 0.4185, labelSize: 16, valueSize: 10.5, labelColor: o.labelColor || C.dark,
    fields: [
      { label: 'Location', value: 'Your Location', x: o.x + 0.463, caret: 1.225 },
      { label: 'Check In', value: 'Add Date', x: o.x + 2.821, caret: 0.956 },
      { label: 'Check Out', value: 'Add Date', x: o.x + 5.178, caret: 1.140 }
    ]
  });
  button(s, { x: o.x + 7.4, y: o.y + 0.249, w: 1.706, h: 0.484, text: 'Book Now', size: 14 });
}

/** Icon + bold title + paragraph, laid out in a column (white cards). */
function featureCard(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: rr(5043, o.w, o.h),
    fill: { color: o.fill || C.white }, line: NOLINE, shadow: shadow()
  });
  icon(s, o.icon, o.x + o.iconDx, o.y + o.iconDy, o.iconSize || 0.744, o.iconColor || C.orange);
  text(s, o.title, { x: o.x + o.padX, y: o.titleY, w: 2.265, h: 0.404, fontFace: F.semi, fontSize: 18, color: o.titleColor || C.teal });
  para(s, { x: o.x + o.padX, y: o.bodyY, w: o.bodyW, h: o.bodyH, text: o.body, size: 12, color: o.bodyColor || C.grey, align: o.bodyAlign || 'justify' });
}

/** Coloured tile + heading + copy, stacked horizontally. */
function iconFeature(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.tileW || 1.104, h: 0.982, rectRadius: rr(11831, o.tileW || 1.104, 0.982),
    fill: o.tileFill === 'none' ? { type: 'none' } : { color: o.tileFill }, line: o.tileLine ? { color: o.tileLine, width: 3 } : NOLINE,
    shadow: o.tileFill === 'none' ? undefined : shadow()
  });
  icon(s, o.icon, o.x + (o.tileW || 1.104) / 2 - 0.375, o.y + 0.116, 0.75, C.white);
  text(s, o.title, { x: o.textX, y: o.y - 0.002, w: 3.41, h: 0.404, fontFace: F.semi, fontSize: 18, color: o.titleColor || C.teal });
  para(s, { x: o.textX, y: o.y + 0.4, w: o.textW, h: 0.774, text: o.body, color: o.bodyColor || C.grey, align: o.bodyAlign || 'justify' });
}

/** Orange tick + label. */
function checkItem(s, x, y, str) {
  icon(s, 'check', x, y, 0.4376, C.orange);
  text(s, str, { x: x + 0.555, y: y + 0.047, w: 3.1, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
}

/** Photo tile with the dark wash, a white place name and a caption. */
function placeCard(s, o) {
  photo(s, { x: o.x, y: o.y, w: o.w, h: o.h, radius: o.radius, label: o.label });
  photoWash(s, { x: o.x, y: o.y, w: o.w, h: o.h });
  icon(s, 'pin', o.x + o.px, o.y + o.h - o.iconUp, 0.744, C.white);
  text(s, o.name, { x: o.x + o.px, y: o.y + o.h - o.titleUp, w: 3.41, h: 0.438, fontFace: F.semi, fontSize: 20, color: C.white });
  para(s, { x: o.x + o.px, y: o.y + o.h - o.subUp, w: 3.55, h: 0.421, text: o.sub, color: C.white });
}

/* --------------------------------------------------------------- copy deck */
const L = {
  full: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in',
  mid: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim',
  card: 'Lorem ipsum dolor sitaw consectetur adipis cing elit, sed do eiusmod',
  cardSm: 'Lorem ipsum dolor sitaw cons ectetur adipis cing',
  feat: 'Lorem ipsum dolor sit amet, con ctetur adipiscing elit, ',
  nostrud: 'nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in',
  nostrudA: 'nostrud exercitation ullamco laboris nisi ut aliquip exea commodo consequat. Duis aute irure dolor in',
  maecenas: 'Lorem ipsum dolor sit amet, conse ctetuer adipiscing elit. Maecenas',
  maecenas2: 'Lorem ipsum dolor sit amet, conse ctetuer adipi scing elit. Maecenas',
  aliqua: 'Lorem ipsum dolor sit amet, consectetur adipis cing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut',
  service: 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit, sed do eiu smod tempor incididunt',
  serviceSm: 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit, sed do eiu',
  price: 'Lorem ipsum dolor sit amet, conse ctetuer adipi scing',
  visit: '187+ Lorem ipsum dolor sit amet, ',
  visitSit: '187+ Lorem ipsum dolor sit, ',
  massa: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.',
  massa2: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.Nunc viverra imperdiet enim. Fusce est. '
};
const T = {
  brand: 'TRAVELUR - TRAVEL PRESENTATION TEMPLATE',
  power: 'Travel brings power and love back into your life.',
  wherever: 'Wherever you go becomes a part of you somehow.',
  labore: 'Labore et dolore',
  rabore: 'Rabore et dolore',
  raboreC: 'Rabore et dolore ctetuer',
  porttitor: 'Porttitor congue massa',
  aliquaTitle: 'Labore et dolore magna aliqua'
};
const CHECKS = ['Suitable for vacation', 'Free Food & Drink', 'Affordable prices ', 'Beautiful View'];

/* =====================================================================
   SLIDES
   ===================================================================== */

/** Slides 1 / 11 / 13 / 34 - full-bleed photo cover with a booking bar. */
function heroCover(s, o) {
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, label: false });
  if (o.wash) fade(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, dir: 'r', stops: o.wash, steps: 30 });
  bookingBar(s, { x: 1.7603, y: o.barY, barTransparency: o.barTransparency, labelColor: C.dark });
  title(s, { x: 1.7603, y: o.titleY, w: 9.7318, h: o.titleH, text: o.title, size: o.titleSize, color: C.white, spc: o.titleSpc });
  text(s, T.brand, { x: 1.7603, y: o.brandY, w: 9.0672, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white, charSpacing: 4.5 });
  para(s, { x: 1.7603, y: o.paraY, w: 9.4551, h: 1.1279, text: L.full, color: C.greyLt });
  button(s, { x: 1.8319, y: o.btnY, w: 1.8434, h: 0.5861, text: o.button, size: o.buttonSize || 16 });
}

function slide1(s) {
  heroCover(s, {
    wash: [[0, 69], [0.35, 69], [1, 100]],
    barY: 1.0246, titleY: 3.0602, titleH: 2.5244, titleSize: 144, titleSpc: 3, title: 'Travelur',
    brandY: 5.9932, paraY: 6.8394, btnY: 8.3937, button: 'Read More'
  });
}

function slide34(s) {
  heroCover(s, {
    wash: [[0, 69], [0.35, 69], [1, 100]],
    barY: 1.0246, titleY: 3.0602, titleH: 3.0125, titleSize: 173, titleSpc: 3, title: 'Thanks',
    brandY: 6.1235, paraY: 6.9697, btnY: 8.4789, button: 'Read More'
  });
}

function slide11(s) {
  heroCover(s, {
    wash: [[0, 22], [0.35, 22], [1, 100]],
    barY: 5.9509, barTransparency: 5, titleY: 2.6287, titleH: 2.7937, titleSize: 80, title: "Let's take a short vacation!",
    brandY: 1.9795, paraY: 7.4622, btnY: 8.9689, button: 'Join Us', buttonSize: 18
  });
}

/** Slide 13 - same cover art, flat dark overlay, centred "Our Services". */
function slide13(s) {
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, label: false });
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: C.dark, transparency: 40 }, line: NOLINE });
  text(s, T.brand, { x: 5.4664, y: 1.9795, w: 9.0672, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white, charSpacing: 4.5, align: 'center' });
  text(s, 'Our Services Travelur', { x: 3.8643, y: 2.6287, w: 12.2715, h: 1.4473, fontFace: F.head, fontSize: 80, color: C.white, align: 'center' });
  para(s, { x: 2.7651, y: 4.3965, w: 14.4697, h: 0.7744, text: L.full, color: C.white, align: 'center' });
  [
    { x: 1.6642, icon: 'pin', ix: 2.1039, iy: 6.062, d: 1.3852, bx: 0.9618 },
    { x: 6.4508, icon: 'ticket', ix: 6.9885, iy: 6.2955, d: 1.1894, bx: 5.7485 },
    { x: 11.202, icon: 'ship', ix: 11.6362, iy: 6.1261, d: 1.3964, bx: 10.5022 },
    { x: 16.0655, icon: 'balloon', ix: 16.4997, iy: 6.0632, d: 1.3964, bx: 15.3681 }
  ].forEach((c) => {
    icon(s, c.icon, c.ix, c.iy, c.d, C.white);
    text(s, T.labore, { x: c.x, y: 7.5977, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.white });
    para(s, { x: c.bx, y: 8.1427, w: 3.6694, h: 1.1279, text: L.service, color: C.white, align: 'center' });
  });
}

/** Slides 2 / 35 - dashed circle + aeroplane, headline on the right. */
function circleCover(s, o) {
  s.addShape('ellipse', { x: 0.963, y: 1.788, w: 7.733, h: 7.733, fill: { type: 'none' }, line: { color: C.teal, width: 5.25, dashType: 'dash' } });
  // two teal brush-stroke crescents that sit inside the dashed circle
  s.addShape('custGeom', {
    x: 1.6, y: 2.5, w: 1.55, h: 1.85, fill: { color: C.teal }, line: NOLINE,
    points: [
      { x: 1.50, y: 0.02 },
      { x: 0.75, y: 0.80, curve: { type: 'cubic', x1: 1.35, y1: 0.30, x2: 1.02, y2: 0.52 } },
      { x: 0.02, y: 1.74, curve: { type: 'cubic', x1: 0.50, y1: 1.12, x2: 0.18, y2: 1.42 } },
      { x: 0.38, y: 1.82 },
      { x: 1.10, y: 0.80, curve: { type: 'cubic', x1: 0.68, y1: 1.55, x2: 0.88, y2: 1.18 } },
      { x: 1.50, y: 0.02, curve: { type: 'cubic', x1: 1.32, y1: 0.48, x2: 1.45, y2: 0.22 } },
      { close: true }
    ]
  });
  s.addShape('custGeom', {
    x: 4.46, y: 5.7, w: 4.0, h: 3.85, fill: { color: C.teal }, line: NOLINE,
    points: [
      { x: 3.96, y: 0.02 },
      { x: 3.09, y: 2.32, curve: { type: 'cubic', x1: 4.09, y1: 1.15, x2: 3.74, y2: 1.75 } },
      { x: 0.04, y: 3.82, curve: { type: 'cubic', x1: 2.39, y1: 2.92, x2: 1.24, y2: 3.60 } },
      { x: 0.00, y: 3.82 },
      { x: 2.24, y: 2.55, curve: { type: 'cubic', x1: 0.64, y1: 3.40, x2: 1.49, y2: 3.05 } },
      { x: 3.64, y: 0.40, curve: { type: 'cubic', x1: 3.14, y1: 2.05, x2: 3.54, y2: 1.30 } },
      { close: true }
    ]
  });
  // aeroplane silhouette (was a raster + svg pair in the source deck)
  s.addShape('custGeom', {
    x: 6.13, y: 1.56, w: 2.32, h: 2.32, rotate: 62, fill: { color: C.teal }, line: { color: C.white, width: 1.5 },
    points: [
      { x: 1.16, y: 0.00 }, { x: 1.34, y: 0.20 }, { x: 1.34, y: 0.92 }, { x: 2.32, y: 1.62 },
      { x: 2.32, y: 1.84 }, { x: 1.34, y: 1.50 }, { x: 1.34, y: 1.98 }, { x: 1.62, y: 2.22 },
      { x: 1.62, y: 2.32 }, { x: 1.16, y: 2.18 }, { x: 0.70, y: 2.32 }, { x: 0.70, y: 2.22 },
      { x: 0.98, y: 1.98 }, { x: 0.98, y: 1.50 }, { x: 0.00, y: 1.84 }, { x: 0.00, y: 1.62 },
      { x: 0.98, y: 0.92 }, { x: 0.98, y: 0.20 }, { close: true }
    ]
  });
  text(s, T.brand, { x: 10, y: 5.8159, w: 8.117, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.teal, charSpacing: 3 });
  title(s, { x: 10, y: 3.1475, w: 8.8779, h: o.titleH, text: o.title, size: o.titleSize });
  para(s, { x: 10, y: 6.8333, w: 8.7934, h: 1.4813, text: L.full });
  button(s, { x: 10.0716, y: 8.5962, w: 1.8434, h: 0.5861, text: 'Read More', size: 14 });
  bookingFields(s, {
    y: 1.8472, gap: 0.3406, caretDy: 0.4502,
    fields: [
      { label: 'Location', value: 'Your Location', x: 9.9864, caret: 1.3358 },
      { label: 'Check In', value: 'Add Date', x: 12.3439, caret: 1.0667 },
      { label: 'Check Out', value: 'Add Date', x: 14.7015, caret: 1.1397 }
    ]
  });
  button(s, { x: 16.9234, y: 1.8472, w: 1.7064, h: 0.6018, text: 'Book Now', size: 16 });
}

const slide2 = (s) => circleCover(s, { title: 'Travelur', titleSize: 144, titleH: 2.5244 });
const slide35 = (s) => circleCover(s, { title: 'Thank You', titleSize: 120, titleH: 2.1205 });

/** Slide 3 - tall left photo, discount badge, two mini features. */
function slide3(s) {
  photo(s, { x: 0, y: 0, w: 9.031, h: SLIDE_H, label: false });
  s.addShape('ellipse', { x: 3.5395, y: 6.0096, w: 4.1077, h: 4.1077, fill: { color: C.orange }, line: { color: C.white, width: 8.5 }, shadow: shadow() });
  s.addText([
    { text: 'Discount', options: { breakLine: true } },
    { text: 'Up To', options: { breakLine: true } },
    { text: '45%', options: { fontSize: 88 } }
  ], { x: 3.5395, y: 6.0096, w: 4.1077, h: 4.1077, fontFace: F.semi, fontSize: 28, italic: true, color: C.white, align: 'center', valign: 'middle' });
  s.addShape('ellipse', { x: 6.2182, y: 6.96, w: 0.7881, h: 0.7881, fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  text(s, 'Best Offer', { x: 6.1661, y: 7.0622, w: 0.8923, h: 0.5722, fontFace: F.semi, fontSize: 14, color: C.orange, align: 'center' });

  eyebrow(s, 10.4619, 2.2715);
  title(s, { x: 10.4619, y: 2.7643, w: 8.5556, h: 1.7166, text: T.power, size: 48 });
  para(s, { x: 10.4619, y: 4.9741, w: 8.2325, h: 1.1279, text: L.mid });
  [{ x: 10.4619, ix: 10.5504, iy: 6.8387, d: 0.744, ic: 'pin' },
   { x: 15.1255, ix: 15.1716, iy: 6.9438, d: 0.6388, ic: 'ticket' }].forEach((c) => {
    icon(s, c.ic, c.ix, c.iy, c.d);
    text(s, T.rabore, { x: c.x, y: 7.7267, w: 2.7569, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.x, y: 8.1289, w: 3.3651, h: 0.7744, text: L.feat });
  });
}

/** Slide 4 - copy left with booking bar, photo right, note card. */
function slide4(s) {
  photo(s, { x: 12.778, y: 1.619, w: 6.175, h: 7.302, radius: 0.2 });
  s.addShape('roundRect', { x: 1.6627, y: 7.4182, w: 8.2325, h: 1.1879, rectRadius: rr(6338, 8.2325, 1.1879), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  para(s, { x: 1.6627, y: 5.0998, w: 8.2325, h: 1.4813, text: L.full });
  title(s, { x: 1.6627, y: 2.8901, w: 8.5556, h: 1.582, text: T.wherever });
  eyebrow(s, 1.6627, 2.3972);

  s.addShape('roundRect', { x: 11.0805, y: 7.2769, w: 5.2633, h: 2.6392, rectRadius: rr(5043, 5.2633, 2.6392), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  text(s, T.aliquaTitle, { x: 11.5661, y: 7.8298, w: 4.1639, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
  para(s, { x: 11.5661, y: 8.3091, w: 4.2955, h: 0.9812, text: L.aliqua, size: 12 });

  bookingFields(s, {
    y: 7.7337, gap: 0.3089, caretDy: 0.4185, labelSize: 16, valueSize: 10.5,
    fields: [
      { label: 'Location', value: 'Your Location', x: 2.1573, caret: 1.2246 },
      { label: 'Check In', value: 'Add Date', x: 3.9655, caret: 0.9555 },
      { label: 'Check Out', value: 'Add Date', x: 5.7736, caret: 1.1397 }
    ]
  });
  button(s, { x: 7.7072, y: 7.7701, w: 1.7064, h: 0.4841, text: 'Book Now', size: 14 });
}

/** Slide 5 - banner photo with five service cards straddling the edge. */
function slide5(s) {
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: 7.079, label: false });
  fade(s, { x: 0, y: 0, w: SLIDE_W, h: 7.0794, dir: 'u', stops: [[0, 0], [1, 81]], steps: 24 });
  text(s, T.brand, { x: 0.9757, y: 1.7378, w: 9.0672, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white, charSpacing: 4.5 });
  title(s, { x: 0.9757, y: 2.387, w: 9.7635, h: 2.7937, text: "Let's take a short vacation!", size: 80, color: C.white });
  text(s, T.aliquaTitle, { x: 12.8081, y: 2.403, w: 4.1639, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.white });
  para(s, { x: 12.8081, y: 2.8823, w: 5.7612, h: 1.587, text: L.massa, size: 12, color: C.white });

  [
    { x: 0.7408, y: 6.319, w: 3.1964, h: 3.3249, icon: 'pin', ix: 1.3799, iy: 6.7033, d: 0.744 },
    { x: 4.4807, y: 6.319, w: 3.1964, h: 3.3249, icon: 'ticket', ix: 5.0552, iy: 6.8461, d: 0.6388 },
    { x: 8.09, y: 6.0693, w: 3.8201, h: 3.9736, icon: 'pin', ix: 9.0409, iy: 6.7033, d: 0.744, hot: true },
    { x: 12.3225, y: 6.319, w: 3.1964, h: 3.3249, icon: 'ship', ix: 12.9512, iy: 6.7725, d: 0.75 },
    { x: 16.0624, y: 6.319, w: 3.1964, h: 3.3249, icon: 'balloon', ix: 16.5948, iy: 6.7096, d: 0.75 }
  ].forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: c.y, w: c.w, h: c.h, rectRadius: rr(5043, c.w, c.h), fill: { color: c.hot ? C.orange : C.white }, line: NOLINE, shadow: shadow() });
    icon(s, c.icon, c.ix, c.iy, c.d, c.hot ? C.white : C.orange);
    const tx = c.x + (c.hot ? 0.7973 : 0.4855);
    text(s, T.labore, { x: tx, y: 7.5977, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: c.hot ? C.white : C.teal });
    para(s, { x: tx, y: 8.077, w: 2.2647, h: 0.9812, text: L.card, size: 12, color: c.hot ? C.white : C.grey });
  });
}

/** Slide 6 - photo top-left, headline right, two cards, footnote. */
function slide6(s) {
  photo(s, { x: 0.831, y: 0, w: 6.385, h: 7.585 });
  eyebrow(s, 9.7062, 1.9042);
  title(s, { x: 9.7062, y: 2.397, w: 8.5556, h: 1.582, text: T.wherever });
  para(s, { x: 9.7062, y: 4.6068, w: 8.2325, h: 1.1279, text: L.mid });
  [{ x: 9.7524, fill: C.white, icon: 'pin', ix: 10.3915, iy: 6.6236, d: 0.744, tc: C.teal, bc: C.grey, oc: C.orange },
   { x: 14.3144, fill: C.teal, icon: 'ticket', ix: 14.8889, iy: 6.7664, d: 0.6388, tc: C.white, bc: C.white, oc: C.white }].forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: 6.3624, w: 3.5935, h: 2.663, rectRadius: rr(5043, 3.5935, 2.663), fill: { color: c.fill }, line: NOLINE, shadow: shadow() });
    icon(s, c.icon, c.ix, c.iy, c.d, c.oc);
    text(s, T.labore, { x: c.x + 0.4855, y: 7.5181, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: c.tc });
    para(s, { x: c.x + 0.4855, y: 7.9973, w: 2.6979, h: 0.6782, text: L.cardSm, size: 12, color: c.bc, align: 'left' });
  });
  text(s, T.aliquaTitle, { x: 0.8316, y: 8.3461, w: 4.1639, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
  para(s, { x: 0.8316, y: 8.8254, w: 5.8455, h: 0.6782, text: L.aliqua, size: 12 });
}

/** Slide 7 - checklist left, photo + white feature panel right. */
function slide7(s) {
  photo(s, { x: 10.571, y: 1.698, w: 8.048, h: 4.905, radius: 0.2 });
  s.addShape('roundRect', { x: 9.2615, y: 6.1881, w: 8.9692, h: 3.6275, rectRadius: rr(5043, 8.9692, 3.6275), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  [{ x: 10.1416, tile: C.orange, icon: 'ship', ix: 10.415, iy: 6.8671, d: 0.75, tx: 10.0531, title: T.raboreC },
   { x: 14.186, tile: C.teal, icon: 'ticket', ix: 14.5119, iy: 6.8486, d: 0.6388, tx: 14.0974, title: T.porttitor }].forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: 6.7133, w: 1.2907, h: 0.9821, rectRadius: rr(11831, 1.2907, 0.9821), fill: { color: c.tile }, line: NOLINE, shadow: shadow() });
    icon(s, c.icon, c.ix, c.iy, c.d, C.white);
    text(s, c.title, { x: c.tx, y: 7.9714, w: 3.4102, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.tx, y: 8.3736, w: 3.5484, h: 0.7744, text: L.maecenas });
  });
  eyebrow(s, 1.0015, 2.3217);
  title(s, { x: 1.0015, y: 2.8145, w: 7.537, h: 1.582, text: T.power });
  para(s, { x: 1.0015, y: 4.6883, w: 7.537, h: 1.4813, text: L.mid });
  checkItem(s, 1.1115, 6.9007, CHECKS[0]);
  checkItem(s, 1.1115, 7.7368, CHECKS[2]);
  checkItem(s, 5.2749, 6.8902, CHECKS[1]);
  checkItem(s, 5.2749, 7.7368, CHECKS[3]);
}

/** Slide 8 - tall photo left with a "37%" badge, two feature rows right. */
function slide8(s) {
  photo(s, { x: 0.892, y: 0.8, w: 7.046, h: 10.45, shape: 'round2SameRect' });
  s.addShape('roundRect', { x: 6.3161, y: 1.29, w: 2.322, h: 2.1135, rectRadius: rr(8660, 2.322, 2.1135), fill: { color: C.orange }, line: { color: C.white, width: 6.25 }, shadow: shadow() });
  s.addText([
    { text: '  Discount', options: { breakLine: true } },
    { text: '  Up To', options: { breakLine: true } },
    { text: ' ', options: { fontSize: 36 } },
    { text: '37%', options: { fontSize: 60 } }
  ], { x: 6.3161, y: 1.29, w: 2.322, h: 2.1135, fontFace: F.semi, fontSize: 16, italic: true, color: C.white, valign: 'middle' });
  s.addShape('roundRect', { x: 7.7794, y: 1.4558, w: 0.6433, h: 0.6433, rectRadius: 0.1, fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  text(s, 'New', { x: 7.6535, y: 1.5887, w: 0.8923, h: 0.3366, fontFace: F.semi, fontSize: 14, color: C.orange, align: 'center' });

  eyebrow(s, 10.1014, 1.7299);
  title(s, { x: 10.1014, y: 2.2227, w: 8.5556, h: 1.582, text: T.wherever });
  para(s, { x: 10.1014, y: 4.2174, w: 8.2325, h: 1.4813, text: L.full });
  iconFeature(s, { x: 10.163, y: 6.316, tileFill: C.orange, icon: 'ship', textX: 11.4999, textW: 6.834, title: T.raboreC, body: L.nostrud });
  iconFeature(s, { x: 10.163, y: 7.9712, tileFill: C.teal, icon: 'ticket', textX: 11.4999, textW: 6.834, title: T.porttitor, body: L.nostrud });
}

/** Slide 9 - three destination bands across the foot of the slide. */
function slide9(s) {
  [{ x: 0, w: 5.127 }, { x: 5.4286, w: 9.1429 }, { x: 14.873, w: 5.127 }].forEach((b) => {
    photo(s, { x: b.x, y: 5.625, w: b.w, h: 4.778, label: false });
    photoWash(s, { x: b.x, y: 5.8966, w: b.w, h: 4.5062 });
  });
  [{ x: 0.6069, ix: 0.73, name: 'Danau Ranau', sub: L.visit },
   { x: 6.1647, ix: 6.2877, name: 'Samudroo Juosh', sub: '992+ Lorem ipsum dolor sit amet, ' },
   { x: 15.5998, ix: 15.7228, name: 'Pantai Asyik', sub: '473+ Lorem ipsum dolor sit amet, ' }].forEach((c) => {
    icon(s, 'pin', c.ix, 8.1677, 0.744, C.white);
    text(s, c.name, { x: c.x, y: 9.0679, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
    para(s, { x: c.x, y: 9.4701, w: 3.5484, h: 0.421, text: c.sub, color: C.white });
  });
  eyebrow(s, 1.0169, 1.3588);
  title(s, { x: 1.0169, y: 1.8517, w: 8.4687, h: 1.7166, text: T.power, size: 47 });
  para(s, { x: 1.0015, y: 3.8448, w: 8.0139, h: 0.7744, text: L.short, align: 'left' });
  [{ x: 10.8276, fill: C.white, icon: 'pin', ix: 11.4513, iy: 1.9091, d: 0.744, tc: C.teal, bc: C.grey, oc: C.orange },
   { x: 15.3896, fill: C.teal, icon: 'ticket', ix: 15.9487, iy: 2.0519, d: 0.6388, tc: C.white, bc: C.white, oc: C.white }].forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: 1.6172, w: 3.5935, h: 2.8814, rectRadius: rr(5043, 3.5935, 2.8814), fill: { color: c.fill }, line: NOLINE, shadow: shadow() });
    icon(s, c.icon, c.ix, c.iy, c.d, c.oc);
    text(s, T.labore, { x: c.x + 0.4701, y: 2.8036, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: c.tc });
    para(s, { x: c.x + 0.4701, y: 3.2828, w: 2.6979, h: 0.6782, text: L.cardSm, size: 12, color: c.bc, align: 'left' });
  });
}

/** Slide 10 - "Find Your Location" with six ticks and three photo cards. */
function slide10(s) {
  [{ x: 1.2, y: 6.9219, w: 5.0308, h: 3.268 },
   { x: 7.0462, y: 6.9219, w: 5.0308, h: 3.268 },
   { x: 12.8923, y: 2.8446, w: 5.9077, h: 7.3453 }].forEach((b) => {
    photo(s, { x: b.x, y: b.y, w: b.w, h: b.h, radius: rr(4015, b.w, b.h), label: false });
    photoWash(s, { x: b.x, y: b.y, w: b.w, h: b.h });
  });
  [{ x: 1.7236, ix: 1.8466, iy: 7.9549, ty: 8.8551, sy: 9.2573, name: 'Danau Ranau', sub: L.visit },
   { x: 7.5612, ix: 7.6843, iy: 7.9509, ty: 8.8511, sy: 9.2533, name: 'Samudroo Juosh', sub: '992+ Lorem ipsum dolor sit amet, ' },
   { x: 13.5363, ix: 13.6593, iy: 7.9592, ty: 8.8594, sy: 9.2616, name: 'Pantai Asyik', sub: '473+ Lorem ipsum dolor sit amet, ' }].forEach((c) => {
    icon(s, 'pin', c.ix, c.iy, 0.744, C.white);
    text(s, c.name, { x: c.x, y: c.ty, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
    para(s, { x: c.x, y: c.sy, w: 3.5484, h: 0.421, text: c.sub, color: C.white });
  });
  eyebrow(s, 1.4024, 1.1901);
  title(s, { x: 1.4024, y: 1.6829, w: 9.97, h: 1.2117, text: 'Find Your Location', size: 66 });
  para(s, { x: 1.4024, y: 3.0098, w: 9.97, h: 1.1279, text: L.mid });
  [[1.5124, 4.6354, CHECKS[0]], [1.5124, 5.4715, CHECKS[2]], [5.2776, 4.625, CHECKS[1]],
   [5.2776, 5.4715, CHECKS[3]], [8.6387, 4.6346, CHECKS[3]], [8.6387, 5.4715, CHECKS[1]]]
    .forEach((c) => checkItem(s, c[0], c[1], c[2]));
}

/** Slides 12 / 29 - section break with a five-field search bar. */
function breakSlide(s, headline) {
  bookingFields(s, {
    y: 1.3458, gap: 0.3407, caretDy: 0.4503,
    fields: [
      { label: 'Location', value: 'Your Location', x: 3.2063, caret: 1.3358 },
      { label: 'Check In', value: 'Add Date', x: 6.1583, caret: 1.0667 },
      { label: 'Check Out', value: 'Add Date', x: 9.1102, caret: 1.1397 },
      { label: 'Price', value: 'Add Price', x: 12.0622, caret: 1.1397 },
      { label: 'Short By', value: 'Name Location', x: 15.0142, caret: 1.5756 }
    ]
  });
  text(s, headline, { x: 3.2615, y: 3.7558, w: 13.4769, h: 2.5244, fontFace: F.head, fontSize: 144, color: C.teal, align: 'center' });
  s.addShape('roundRect', { x: 5.2462, y: 6.3563, w: 9.5077, h: 0.6537, rectRadius: rr(11831, 9.5077, 0.6537), fill: { color: C.orange }, line: NOLINE, shadow: shadow() });
  text(s, T.brand, { x: 5.9415, y: 6.4799, w: 8.117, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white, charSpacing: 3, align: 'center' });
  para(s, { x: 4.4192, y: 8.7763, w: 11.1617, h: 1.1279, text: L.full, align: 'center' });
}

const slide12 = (s) => breakSlide(s, 'Break Slides');
const slide29 = (s) => breakSlide(s, 'Icons Slides');

/** Slide 14 - "Our Services Travelur" with a 2x2 card grid. */
function slide14(s) {
  eyebrow(s, 1.4444, 2.1706);
  title(s, { x: 1.4444, y: 2.6634, w: 8.5556, h: 3.063, text: 'Our Services Travelur', size: 88 });
  para(s, { x: 1.4444, y: 6.4894, w: 8.2325, h: 1.4813, text: L.full });
  [{ x: 11.3056, y: 1.7594, fill: C.teal, icon: 'ship', ix: 11.9343, iy: 2.2128, d: 0.75, dark: true },
   { x: 15.0455, y: 1.7594, fill: C.white, icon: 'balloon', ix: 15.5778, iy: 2.15, d: 0.75 },
   { x: 11.3056, y: 5.7263, fill: C.white, icon: 'pin', ix: 11.9447, iy: 6.1106, d: 0.744 },
   { x: 15.0455, y: 5.7263, fill: C.white, icon: 'ticket', ix: 15.62, iy: 6.2534, d: 0.6388 }].forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: c.y, w: 3.1964, h: 3.3249, rectRadius: rr(5043, 3.1964, 3.3249), fill: { color: c.fill }, line: NOLINE, shadow: shadow() });
    icon(s, c.icon, c.ix, c.iy, c.d, c.dark ? C.white : C.orange);
    text(s, T.labore, { x: c.x + 0.4855, y: c.y + 1.2787, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: c.dark ? C.white : C.teal });
    para(s, { x: c.x + 0.4855, y: c.y + 1.758, w: 2.2647, h: 0.9812, text: L.card, size: 12, color: c.dark ? C.white : C.grey });
  });
}

/** Slide 15 - white left column, solid teal right column with four rows. */
function slide15(s) {
  photo(s, { x: 0, y: 0, w: 10, h: 5.159 });
  s.addShape('rect', { x: 10, y: 0, w: 10, h: SLIDE_H, fill: { color: C.teal }, line: NOLINE });
  eyebrow(s, 1.0015, 5.8101);
  title(s, { x: 1.0015, y: 6.303, w: 7.537, h: 1.582, text: T.power });
  checkItem(s, 1.1115, 8.4894, CHECKS[0]);
  checkItem(s, 1.1115, 9.3255, CHECKS[2]);
  checkItem(s, 5.2749, 8.479, CHECKS[1]);
  checkItem(s, 5.2749, 9.3255, CHECKS[3]);
  [{ y: 1.659, icon: 'ship', ix: 11.6282, iy: 1.7751, d: 0.75, title: T.raboreC, body: L.nostrudA },
   { y: 3.9428, icon: 'ticket', ix: 11.6838, iy: 4.1144, d: 0.6388, title: T.porttitor, body: L.nostrudA + 'A' },
   { y: 6.2312, icon: 'pin', ix: 11.6853, iy: 6.3885, d: 0.6396, title: T.raboreC, body: L.nostrudA },
   { y: 8.5149, icon: 'balloon', ix: 11.6736, iy: 8.6656, d: 0.6388, title: T.porttitor, body: L.nostrudA + 'A' }].forEach((r) => {
    s.addShape('roundRect', { x: 11.4511, y: r.y, w: 1.1042, h: 0.9821, rectRadius: rr(11831, 1.1042, 0.9821), fill: { type: 'none' }, line: { color: C.white, width: 3 } });
    icon(s, r.icon, r.ix, r.iy, r.d, C.white);
    text(s, r.title, { x: 12.7881, y: r.y - 0.1379, w: 3.4102, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.white });
    para(s, { x: 12.7881, y: r.y + 0.2643, w: 5.7882, h: 0.7744, text: r.body, color: C.white, align: 'left' });
  });
}

/** Slide 16 - full-width quote card. */
function slide16(s) {
  photo(s, { x: 0, y: 0, w: 10.983, h: SLIDE_H, label: false });
  s.addShape('roundRect', { x: 3.6645, y: 1.3142, w: 15.166, h: 8.6215, rectRadius: rr(862, 15.166, 8.6215), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
  eyebrow(s, 5.3218, 2.3909);
  title(s, { x: 5.3218, y: 2.8837, w: 12.6613, h: 3.4332, text: 'Travel makes one modest. You see what a tiny place you occupy in the world.', size: 66 });
  text(s, '\u201D', { x: 15.8862, y: 5.4254, w: 2.89, h: 5.8903, fontFace: F.head, fontSize: 344, color: C.teal });
  para(s, { x: 5.3218, y: 6.6977, w: 10.102, h: 0.7744, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut' });
  button(s, { x: 5.3934, y: 7.8333, w: 1.8434, h: 0.5861, text: 'Read More', size: 14 });
}

/** Slide 17 - mosaic of three photos plus a teal note box. */
function slide17(s) {
  [{ x: 0.787, y: 3.015, w: 4.843, h: 7.382 },
   { x: 5.631, y: 3.015, w: 4.843, h: 4.103 },
   { x: 10.474, y: 3.908, w: 8.738, h: 3.211 }].forEach((p) => photo(s, { x: p.x, y: p.y, w: p.w, h: p.h, label: false }));
  s.addShape('rect', { x: 10.477, y: 0.923, w: 8.736, h: 2.985, fill: { color: C.teal }, line: NOLINE });
  photoWash(s, { x: 0.7874, y: 5.3175, w: 4.8433, h: 5.0797 });
  photoWash(s, { x: 5.6308, y: 4.0331, w: 4.8433, h: 3.085 });
  photoWash(s, { x: 10.4741, y: 4.7027, w: 8.7385, h: 2.4154 });
  [{ x: 1.3111, ix: 1.4341, iy: 8.1621, ty: 9.0623, name: 'Danau Ranau', sub: L.visit },
   { x: 6.1459, ix: 6.2689, iy: 4.8791, ty: 5.7793, name: 'Samudroo Juosh', sub: '992+ Lorem ipsum dolor sit amet, ' },
   { x: 12.2927, ix: 11.2411, iy: 5.7876, ty: 5.7876, name: 'Pantai Asyik', sub: '473+ Lorem ipsum dolor sit amet, ' }].forEach((c) => {
    icon(s, 'pin', c.ix, c.iy, 0.744, C.white);
    text(s, c.name, { x: c.x, y: c.ty, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
    para(s, { x: c.x, y: c.ty + 0.4022, w: 3.5484, h: 0.421, text: c.sub, color: C.white });
  });
  eyebrow(s, 0.7874, 0.8242);
  title(s, { x: 0.7874, y: 1.317, w: 8.4687, h: 1.3127, text: T.wherever, size: 36 });
  text(s, T.aliquaTitle, { x: 11.2411, y: 1.6477, w: 4.1639, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.white });
  para(s, { x: 11.2411, y: 2.1573, w: 7.2936, h: 0.9812, text: L.massa2, size: 12, color: C.white });
  [{ x: 6.8677, ix: 6.9361, icon: 'pin', d: 0.744, iy: 7.8288, title: T.raboreC },
   { x: 11.1874, ix: 11.3325, icon: 'ship', d: 0.75, iy: 7.8288, title: T.raboreC },
   { x: 15.369, ix: 15.5666, icon: 'ticket', d: 0.6388, iy: 7.8102, title: T.porttitor }].forEach((c) => {
    icon(s, c.icon, c.ix, c.iy, c.d);
    text(s, c.title, { x: c.x, y: 8.7902, w: 3.4102, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.x, y: 9.1924, w: 3.5484, h: 0.7744, text: L.maecenas });
  });
}

/** Slide 18 - two discounted photo cards left, checklist right. */
function slide18(s) {
  [{ y: 1.099, gy: 1.6903, oy: 2.3979, pct: '45%', name: 'Danau Ranau', ny: 3.6235 },
   { y: 6.104, gy: 6.6955, oy: 7.4031, pct: '37%', name: 'Pantai Juosh', ny: 8.6287 }].forEach((c) => {
    photo(s, { x: 1.172, y: c.y, w: 6.797, h: 3.859, radius: 0.27, label: false });
    photoWash(s, { x: 1.1719, y: c.gy, w: 6.7969, h: 3.268 });
    s.addShape('ellipse', { x: 7.338, y: c.oy, w: 1.2616, h: 1.2616, fill: { color: C.orange }, line: { color: C.white, width: 4.25 }, shadow: shadow() });
    s.addText([{ text: 'Discount', options: { breakLine: true } }, { text: c.pct, options: { fontSize: 20 } }],
      { x: 7.338, y: c.oy, w: 1.2616, h: 1.2616, fontFace: F.semi, fontSize: 11, italic: true, color: C.white, align: 'center', valign: 'middle' });
    text(s, c.name, { x: 1.6955, y: c.ny, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
    para(s, { x: 1.6955, y: c.ny + 0.4022, w: 3.5484, h: 0.421, text: L.visit, color: C.white });
  });
  eyebrow(s, 10.4971, 2.1342);
  title(s, { x: 10.4971, y: 2.627, w: 7.537, h: 1.582, text: T.power });
  para(s, { x: 10.4971, y: 4.5007, w: 7.537, h: 1.4813, text: L.mid });
  checkItem(s, 10.6071, 6.7132, CHECKS[0]);
  checkItem(s, 10.6071, 7.5492, CHECKS[2]);
  checkItem(s, 14.7706, 6.7027, CHECKS[1]);
  checkItem(s, 14.7706, 7.5492, CHECKS[3]);
}

/** Small "267 Liked / 80 Days / 123+" meta row used on the pricing cards. */
function metaLiked(s, x, y) {
  s.addShape('heart', { x: x - 0.1613, y: y + 0.1709, w: 0.2, h: 0.1552, fill: { type: 'none' }, line: { color: C.grey, width: 1 } });
  text(s, '267 Liked', { x, y, w: 1.372, h: 0.421, fontFace: F.semi, fontSize: 14, color: C.grey, lineSpacingMultiple: 1.5 });
}
function metaDays(s, x, y) {
  icon(s, 'clock', x - 0.1916, y + 0.1516, 0.1916, C.grey);
  text(s, '80 Days', { x, y, w: 1.0469, h: 0.421, fontFace: F.semi, fontSize: 14, color: C.grey, lineSpacingMultiple: 1.5 });
}

/** Slide 19 - three price cards under a "Find Your Favorite Location" title. */
function slide19(s) {
  ['Rabore et dolore', 'Larno Kolosium', 'Puantai Juash'].forEach((name, i) => {
    const dx = i * 6.2086;
    const cx = 0.85 + i * 6.2445;
    s.addShape('roundRect', { x: cx, y: 3.794, w: 5.811, h: 6.142, rectRadius: rr(2840, 5.811, 6.142), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
    photo(s, { x: cx, y: 3.794, w: 5.811, h: 3.165, shape: 'round2SameRect' });
    text(s, name, { x: 1.2914 + dx, y: 7.5177, w: 2.4273, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    text(s, ['$892', '$752', '$956'][i], { x: 5.0208 + dx, y: 7.5177, w: 1.1771, h: 0.5049, fontFace: F.semi, fontSize: 24, color: C.orange, align: 'right' });
    metaLiked(s, 1.5628 + dx, 7.8121);
    para(s, { x: 1.2914 + dx, y: 8.317, w: 4.9065, h: 0.7744, text: L.maecenas2, align: 'left' });
    metaDays(s, 4.1662 + dx, 9.0326);
    icon(s, 'eye', 5.3366 + dx, 9.1842, 0.2222, C.grey);
    text(s, '123+', { x: 5.5588 + dx, y: 9.0326, w: 0.6669, h: 0.421, fontFace: F.semi, fontSize: 14, color: C.grey, lineSpacingMultiple: 1.5 });
  });
  eyebrow(s, 1.0465, 0.7302);
  title(s, { x: 1.0465, y: 1.223, w: 7.7501, h: 1.9186, text: 'Find Your Favorite Location ', size: 54 });
  text(s, 'Exercitation ullamco laboris ', { x: 8.9949, y: 1.2345, w: 4.7137, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.teal });
  para(s, { x: 8.9949, y: 1.7964, w: 9.97, h: 1.1279, text: L.mid });
}

/** Slide 20 - two discounted price cards on the right. */
function slide20(s) {
  [{ y: 0.915, oy: 1.2481, pct: '15%', was: '$956', now: '$867', dy: 0 },
   { y: 5.996, oy: 6.2379, pct: '20%', was: '$978', now: '$854', dy: 5.1234 }].forEach((c) => {
    s.addShape('roundRect', { x: 10.695, y: c.y, w: 7.726, h: 4.339, rectRadius: rr(2840, 7.726, 4.339), fill: { color: C.white }, line: NOLINE, shadow: shadow() });
    photo(s, { x: 10.695, y: c.y, w: 7.726, h: 2.46, shape: 'round2SameRect' });
    s.addShape('ellipse', { x: 11.0136, y: c.oy, w: 1.2616, h: 1.2616, fill: { color: C.orange }, line: { color: C.white, width: 4.25 }, shadow: shadow() });
    s.addText([{ text: 'Discount', options: { breakLine: true } }, { text: c.pct, options: { fontSize: 20 } }],
      { x: 11.0136, y: c.oy, w: 1.2616, h: 1.2616, fontFace: F.semi, fontSize: 11, italic: true, color: C.white, align: 'center', valign: 'middle' });
    text(s, 'Puantai Juash', { x: 11.0136, y: 3.6891 + c.dy, w: 2.4273, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.orange });
    s.addText([
      { text: c.was, options: { color: C.grey, strike: 'sngStrike' } },
      { text: '  ', options: { color: C.grey } },
      { text: c.now, options: { color: C.teal } }
    ], { x: 15.1015, y: 3.6866 + c.dy, w: 2.9203, h: 0.6395, fontFace: F.semi, fontSize: 32, align: 'right', valign: 'top' });
    para(s, { x: 11.0136, y: 4.1041 + c.dy, w: 2.9203, h: 0.7744, text: L.price, align: 'left' });
    metaLiked(s, 15.5934, 4.4408 + c.dy);
    metaDays(s, 17.0918, 4.4236 + c.dy);
  });
  eyebrow(s, 1.1901, 1.7299);
  title(s, { x: 1.1901, y: 2.2227, w: 8.5556, h: 1.582, text: T.wherever });
  para(s, { x: 1.1901, y: 4.2174, w: 8.2325, h: 1.4813, text: L.full });
  iconFeature(s, { x: 1.2516, y: 6.316, tileFill: C.orange, icon: 'ship', textX: 2.5886, textW: 6.834, title: T.raboreC, body: L.nostrud });
  iconFeature(s, { x: 1.2516, y: 7.9712, tileFill: C.teal, icon: 'ticket', textX: 2.5886, textW: 6.834, title: T.porttitor, body: L.nostrud });
}

/** Slide 21 - "Our Porfolio" with a three-photo column. */
function slide21(s) {
  photo(s, { x: 0.337, y: 0, w: 4.044, h: SLIDE_H, label: false });
  photo(s, { x: 4.812, y: 0, w: 4.044, h: 4.446, label: false });
  photo(s, { x: 4.812, y: 4.908, w: 4.044, h: 6.342, label: false });
  photoWash(s, { x: 4.8117, y: 0.322, w: 4.0437, h: 4.1241 });
  text(s, 'View Photo', { x: 5.3353, y: 3.1113, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
  para(s, { x: 5.3353, y: 3.5135, w: 2.9697, h: 0.421, text: L.visitSit, color: C.white });
  eyebrow(s, 10.9534, 1.7095);
  title(s, { x: 10.9534, y: 2.2024, w: 7.1991, h: 2.7937, text: 'Our Porfolio Travelur', size: 80 });
  para(s, { x: 10.9534, y: 5.3145, w: 7.4205, h: 1.4813, text: L.mid });
  [{ x: 10.9534, ix: 11.0419, iy: 7.1791, d: 0.744, ic: 'pin' },
   { x: 15.4355, ix: 15.4816, iy: 7.2842, d: 0.6388, ic: 'ticket' }].forEach((c) => {
    icon(s, c.ic, c.ix, c.iy, c.d);
    text(s, T.rabore, { x: c.x, y: 8.0671, w: 2.7569, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.x, y: 8.4693, w: 3.1991, h: 0.7744, text: L.feat, align: 'left' });
  });
}

/** Slide 22 - gallery mosaic of six captioned photos. */
function slide22(s) {
  [{ x: 0.593, y: 3.797, w: 6.441, h: 2.678 }, { x: 0.593, y: 6.763, w: 6.441, h: 3.78 },
   { x: 7.396, y: 3.797, w: 3.655, h: 6.746 }, { x: 11.413, y: 3.797, w: 7.994, h: 4.083 },
   { x: 11.413, y: 8.136, w: 3.948, h: 2.407 }, { x: 15.661, y: 8.136, w: 3.746, h: 2.407 }]
    .forEach((p) => photo(s, { x: p.x, y: p.y, w: p.w, h: p.h, radius: 0.14, label: false }));
  [{ x: 0.5932, y: 4.2203, w: 6.4407, h: 2.2542, name: 'Pantai Juohh', ny: 5.2567 },
   { x: 0.5932, y: 7.8644, w: 6.4407, h: 2.678, name: 'Danau Keren', ny: 9.3245 },
   { x: 7.4063, y: 6.8983, w: 3.6446, h: 3.6441, name: 'Hiling Asyik', ny: 9.3245 },
   { x: 11.4128, y: 5.6102, w: 7.994, h: 2.2542, name: 'Asiknya Bermain', ny: 6.6465 },
   { x: 11.4128, y: 8.7119, w: 3.9483, h: 1.8305, name: 'Pulau Hantu', ny: 9.3245 },
   { x: 15.661, y: 8.7119, w: 3.7458, h: 1.8305, name: 'Lakosta', ny: 9.3245 }].forEach((c) => {
    photoWash(s, { x: c.x, y: c.y, w: c.w, h: c.h });
    text(s, c.name, { x: c.x + 0.3711, y: c.ny, w: 3.4102, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.white });
    para(s, { x: c.x + 0.3711, y: c.ny + 0.4022, w: 2.9697, h: 0.421, text: L.visitSit, color: C.white });
  });
  eyebrow(s, 1.0465, 0.7302);
  title(s, { x: 1.0465, y: 1.223, w: 7.7501, h: 1.9186, text: 'View All Gallery Porfolio Travelur', size: 54 });
  text(s, 'Exercitation ullamco laboris ', { x: 8.9949, y: 1.2345, w: 4.7137, h: 0.4376, fontFace: F.semi, fontSize: 20, color: C.teal });
  para(s, { x: 8.9949, y: 1.7964, w: 9.97, h: 1.1279, text: L.mid });
}

/** Slide 23 - desktop monitor mock-up. */
function slide23(s) {
  monitor(s, { x: 10.5774, y: 2.0894, w: 8.0861, h: 7.8089 });
  eyebrow(s, 1.1901, 2.3339);
  title(s, { x: 1.1901, y: 2.8267, w: 8.5556, h: 1.3127, text: 'Mockup Desktop', size: 70 });
  para(s, { x: 1.1901, y: 4.5597, w: 8.2325, h: 1.4813, text: L.full });
  iconFeature(s, { x: 1.2516, y: 6.8821, tileFill: C.orange, icon: 'ship', textX: 2.5886, textW: 6.834, title: T.raboreC, body: L.nostrud });
}

/** Desktop monitor drawn from native shapes (replaces a raster mock-up). */
function monitor(s, o) {
  const bez = o.h * 0.596;   // dark bezel
  const chin = o.h * 0.116;  // light grey chin below the screen
  s.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: bez, rectRadius: 0.09, fill: { color: '3B2F44' }, line: NOLINE });
  s.addShape('rect', { x: o.x + o.w * 0.035, y: o.y + o.h * 0.037, w: o.w * 0.93, h: bez - o.h * 0.081, fill: { color: C.white }, line: NOLINE });
  s.addShape('rect', { x: o.x, y: o.y + bez, w: o.w, h: chin, fill: { color: 'B9BCBE' }, line: NOLINE });
  s.addShape('trapezoid', { x: o.x + o.w * 0.355, y: o.y + bez + chin, w: o.w * 0.29, h: o.h * 0.123, flipV: true, fill: { color: 'A8ACAE' }, line: NOLINE });
  s.addShape('trapezoid', { x: o.x + o.w * 0.30, y: o.y + o.h * 0.835, w: o.w * 0.40, h: o.h * 0.05, fill: { color: 'D2D5D7' }, line: NOLINE });
}

/** Slide 24 - laptop mock-up plus a four-icon feature row. */
function slide24(s) {
  laptop(s, { x: 0.801, y: 1.1281, w: 8.2979, h: 5.1257 });
  eyebrow(s, 9.9981, 1.4427);
  title(s, { x: 9.9981, y: 1.9355, w: 8.7934, h: 2.0532, text: 'Eassy Access Travelur With Laptop', size: 56 });
  para(s, { x: 10, y: 4.3699, w: 8.7934, h: 1.1279, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure' });
  button(s, { x: 10.0716, y: 5.879, w: 1.8434, h: 0.5861, text: 'Read More' });
  [{ x: 1.6642, bx: 0.9618, icon: 'pin', ix: 2.2395, iy: 7.4226, d: 0.93 },
   { x: 6.4508, bx: 5.7485, icon: 'ticket', ix: 7.1241, iy: 7.5918, d: 0.7985 },
   { x: 11.202, bx: 10.5022, icon: 'ship', ix: 11.7718, iy: 7.4904, d: 0.9375 },
   { x: 16.0655, bx: 15.3681, icon: 'balloon', ix: 16.6353, iy: 7.4275, d: 0.9375 }].forEach((c) => {
    icon(s, c.icon, c.ix, c.iy, c.d);
    text(s, T.labore, { x: c.x, y: 8.5031, w: 2.2647, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.bx, y: 9.0481, w: 3.6694, h: 0.7744, text: L.serviceSm, align: 'center' });
  });
}

/** Open laptop drawn from native shapes. */
function laptop(s, o) {
  s.addShape('roundRect', { x: o.x + o.w * 0.096, y: o.y, w: o.w * 0.81, h: o.h * 0.955, rectRadius: 0.1, fill: { color: '2B2B2D' }, line: NOLINE });
  s.addShape('rect', { x: o.x + o.w * 0.122, y: o.y + o.h * 0.045, w: o.w * 0.758, h: o.h * 0.865, fill: { color: C.white }, line: NOLINE });
  s.addShape('trapezoid', { x: o.x + o.w * 0.02, y: o.y + o.h * 0.955, w: o.w * 0.96, h: o.h * 0.04, fill: { color: 'C9CCCE' }, line: NOLINE });
  s.addShape('rect', { x: o.x + o.w * 0.42, y: o.y + o.h * 0.96, w: o.w * 0.17, h: o.h * 0.014, fill: { color: '9EA2A4' }, line: NOLINE });
}

/** Slide 25 - large angled monitor mock-up + checklist. */
function slide25(s) {
  // tilted desktop monitor (the source deck used a photographic mock-up)
  const tilt = 351;
  s.addShape('trapezoid', { x: 14.4, y: 7.9, w: 1.9, h: 1.6, flipV: true, fill: { color: 'AFAFAF' }, line: NOLINE });
  s.addShape('rect', { x: 11.85, y: 7.15, w: 7.6, h: 0.44, rotate: tilt, fill: { color: 'D4D4D2' }, line: NOLINE });
  s.addShape('roundRect', { x: 11.72, y: 2.2, w: 7.6, h: 5.0, rectRadius: 0.1, rotate: tilt, fill: { color: '232323' }, line: NOLINE });
  s.addShape('rect', { x: 11.95, y: 2.44, w: 7.15, h: 4.5, rotate: tilt, fill: { color: C.white }, line: NOLINE });
  eyebrow(s, 1.5511, 2.6031);
  title(s, { x: 1.5511, y: 3.096, w: 7.9307, h: 2.3225, text: 'Preview Travelur With Desktop', size: 66 });
  para(s, { x: 1.5511, y: 5.597, w: 7.9307, h: 1.4813, text: L.mid });
  checkItem(s, 1.6611, 7.1821, CHECKS[0]);
  checkItem(s, 1.6611, 8.0182, CHECKS[2]);
  checkItem(s, 5.8246, 7.1717, CHECKS[1]);
  checkItem(s, 5.8246, 8.0182, CHECKS[3]);
}

/**
 * Phone mock-up drawn from native shapes. `crop` says which edge runs off the
 * slide: 'bottom' rounds only the top corners, 'top' only the bottom ones.
 * The screen is see-through in the source deck, so it is painted with `screen`
 * and, past `splitX`, with white - matching whatever sits behind the handset.
 */
function phoneMock(s, o) {
  const flip = o.crop === 'top';
  const f = o.frame || 0.16;   // bezel thickness
  const sx = o.x + f, sw = o.w - 2 * f, sy = o.y + (flip ? 0 : f), sh = o.h - f;
  s.addShape('round2SameRect', { x: o.x, y: o.y, w: o.w, h: o.h, flipV: flip, fill: { color: '1F1F21' }, line: NOLINE });
  s.addShape('round2SameRect', { x: sx, y: sy, w: sw, h: sh, flipV: flip, fill: { color: o.screen || C.white }, line: NOLINE });
  if (o.splitX) {
    s.addShape('round1Rect', {
      x: o.splitX, y: sy, w: sx + sw - o.splitX, h: sh, rectRadius: 0.16667 * Math.min(sw, sh),
      fill: { color: C.white }, line: NOLINE
    });
  }
  if (!flip) {
    s.addShape('roundRect', { x: o.x + o.w * 0.26, y: o.y + f, w: o.w * 0.485, h: o.w * 0.074, rectRadius: 0.08, fill: { color: '1F1F21' }, line: NOLINE });
  }
}

/** Slide 26 - single phone mock-up on a teal panel. */
function slide26(s) {
  s.addShape('rect', { x: 0, y: 0, w: 4.258, h: SLIDE_H, fill: { color: C.teal }, line: NOLINE });
  phoneMock(s, { x: 1.12, y: 1.36, w: 6.1, h: 9.89, crop: 'bottom', frame: 0.2, screen: C.teal, splitX: 4.258 });
  eyebrow(s, 9.6584, 2.0201);
  title(s, { x: 9.6584, y: 2.5129, w: 8.5707, h: 2.7937, text: 'Easy Access call & message', size: 80 });
  para(s, { x: 9.6584, y: 5.625, w: 8.5707, h: 1.1279, text: L.mid });
  [{ x: 9.6584, ix: 9.747, iy: 7.4896, d: 0.744, ic: 'pin' },
   { x: 14.386, ix: 14.4321, iy: 7.5948, d: 0.6388, ic: 'ticket' }].forEach((c) => {
    icon(s, c.ic, c.ix, c.iy, c.d);
    text(s, T.rabore, { x: c.x, y: 8.3776, w: 2.7569, h: 0.404, fontFace: F.semi, fontSize: 18, color: C.teal });
    para(s, { x: c.x, y: 8.7798, w: 3.97, h: 0.7744, text: L.feat, align: 'left' });
  });
}

/** Slide 27 - two phone mock-ups, search bar and two buttons. */
function slide27(s) {
  phoneMock(s, { x: 15.6, y: -1.5, w: 4.34, h: 7.9, crop: 'top', frame: 0.18 });
  phoneMock(s, { x: 10.58, y: 3.84, w: 4.42, h: 7.41, crop: 'bottom', frame: 0.18 });
  bookingFields(s, {
    y: 1.0629, gap: 0.3407, caretDy: 0.4503,
    fields: [
      { label: 'Location', value: 'Your Location', x: 1.557, caret: 1.3357 },
      { label: 'Check In', value: 'Add Date', x: 4.5089, caret: 1.0667 },
      { label: 'Check Out', value: 'Add Date', x: 7.4609, caret: 1.1397 },
      { label: 'Price', value: 'Add Price', x: 10.2474, caret: 1.0603 },
      { label: 'Short By', value: 'Name Location', x: 12.5324, caret: 1.5757 }
    ]
  });
  eyebrow(s, 1.557, 2.7735);
  title(s, { x: 1.557, y: 3.2663, w: 7.7972, h: 3.1303, text: 'Easy access call & message with App mobile', size: 60 });
  para(s, { x: 1.5589, y: 6.7778, w: 7.7953, h: 1.4813, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure' });
  button(s, { x: 1.6305, y: 8.6402, w: 1.8434, h: 0.5861, text: 'View Demo' });
  button(s, { x: 4.0003, y: 8.6402, w: 2.3122, h: 0.5861, text: 'Download App', fill: C.teal });
}

/** Slide 28 - "Get In Touch" contact details on a white rounded panel. */
function slide28(s) {
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, label: false });
  s.addShape('round2SameRect', { x: 2.8495, y: -0.6568, w: 7.1823, h: 12.8814, rotate: 90, fill: { color: C.white }, line: NOLINE });
  eyebrow(s, 1.4463, 3.0362);
  title(s, { x: 1.4463, y: 3.529, w: 9.4912, h: 1.1107, text: 'Get In Touch Traverlur', size: 60 });
  para(s, { x: 1.4463, y: 4.825, w: 9.3037, h: 0.7744, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco' });
  [{ y: 5.9534, adj: 9635, label: 'Email\t\t:', ly: 5.9876, icon: 'mail', iy: 6.193, lines: ['Hello.taaravelurrrrrr@mail.com', 'travelue213455@mail.com'], ty: 6.0075 },
   { y: 7.1279, adj: 7510, label: 'Phone\t\t:', ly: 7.1862, icon: 'phone', iy: 7.2953, lines: ['+012 \u2013 3884 -7221', '+321 \u2013 5432 -5323'], ty: 7.1614 }].forEach((r) => {
    s.addShape('roundRect', { x: 1.4749, y: r.y, w: 0.9503, h: 0.7845, rectRadius: rr(r.adj, 0.9503, 0.7845), fill: { color: C.orange }, line: NOLINE });
    icon(s, r.icon, 1.7539, r.iy, 0.39, C.white);
    text(s, r.label, { x: 2.6107, y: r.ly, w: 2.5121, h: 0.6395, fontFace: F.semi, fontSize: 32, color: C.teal });
    r.lines.forEach((ln, i) => text(s, ln, { x: 5.2307, y: r.ty + i * 0.4721, w: 4.7941, h: 0.4376, fontSize: 20, color: C.teal }));
  });
  button(s, { x: 8.92, y: 7.4588, w: 1.8434, h: 0.5861, text: 'Contact Us' });
}

/* ------------------------------------------------------- icon sheets 30-33 */

/** Glyph vocabulary for the four icon-library slides (black line/solid marks). */
const GLYPHS = [
  (s, x, y, d) => dot(s, x, y, d, C.ink),
  (s, x, y, d) => ring(s, x, y, d, C.ink, 2),
  (s, x, y, d) => s.addShape('star5', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('star5', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.ink, width: 1.5 } }),
  (s, x, y, d) => s.addShape('heart', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('rect', { x, y: y + 0.18 * d, w: d, h: 0.64 * d, fill: { type: 'none' }, line: { color: C.ink, width: 1.5 } }),
  (s, x, y, d) => s.addShape('triangle', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('diamond', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('cloud', { x, y: y + 0.12 * d, w: d, h: 0.76 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('rightArrow', { x, y: y + 0.22 * d, w: d, h: 0.56 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('plus', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('hexagon', { x, y: y + 0.14 * d, w: d, h: 0.72 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('moon', { x: x + 0.2 * d, y, w: 0.6 * d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('sun', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('pie', { x, y, w: d, h: d, angleRange: [270, 200], fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('donut', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('flowChartMagneticDisk', { x, y: y + 0.1 * d, w: d, h: 0.8 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('lightningBolt', { x: x + 0.2 * d, y, w: 0.6 * d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('flowChartDocument', { x, y: y + 0.14 * d, w: d, h: 0.72 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('smileyFace', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.ink, width: 1.5 } }),
  (s, x, y, d) => s.addShape('gear6', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('teardrop', { x, y, w: d, h: d, rotate: 225, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('chevron', { x, y: y + 0.2 * d, w: d, h: 0.6 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('can', { x: x + 0.15 * d, y, w: 0.7 * d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('bracketPair', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.ink, width: 2 } }),
  (s, x, y, d) => s.addShape('star8', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('leftRightArrow', { x, y: y + 0.28 * d, w: d, h: 0.44 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('frame', { x, y: y + 0.12 * d, w: d, h: 0.76 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('cube', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('trapezoid', { x, y: y + 0.22 * d, w: d, h: 0.56 * d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('bevel', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('flowChartConnector', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.ink, width: 2.5 } }),
  (s, x, y, d) => s.addShape('upDownArrow', { x: x + 0.28 * d, y, w: 0.44 * d, h: d, fill: { color: C.ink }, line: NOLINE }),
  (s, x, y, d) => s.addShape('pentagon', { x, y, w: d, h: d, fill: { color: C.ink }, line: NOLINE })
];

/**
 * Grid of pictogram tiles - stands in for the deck's icon-library pages.
 * `cols` x `rows` marks, evenly spaced between the given margins.
 */
function iconSheet(s, o) {
  for (let r = 0; r < o.rows; r++) {
    for (let c = 0; c < o.cols; c++) {
      const cx = o.x0 + (c * (o.x1 - o.x0)) / (o.cols - 1);
      const cy = o.y0 + (r * (o.y1 - o.y0)) / (o.rows - 1);
      // stride by 7 (co-prime with the glyph count) so neighbours stay distinct
      GLYPHS[((r * o.cols + c) * 7 + o.seed) % GLYPHS.length](s, cx - o.d / 2, cy - o.d / 2, o.d);
    }
  }
}

const slide30 = (s) => iconSheet(s, { cols: 12, rows: 8, x0: 1.124, x1: 18.838, y0: 0.948, y1: 10.263, d: 0.57, seed: 0 });
const slide31 = (s) => iconSheet(s, { cols: 12, rows: 8, x0: 0.869, x1: 19.084, y0: 0.94, y1: 10.174, d: 0.56, seed: 5 });
const slide32 = (s) => iconSheet(s, { cols: 11, rows: 6, x0: 1.541, x1: 18.312, y0: 1.198, y1: 9.482, d: 0.6, seed: 9 });
const slide33 = (s) => iconSheet(s, { cols: 11, rows: 6, x0: 1.716, x1: 18.274, y0: 1.48, y1: 9.52, d: 0.62, seed: 13 });

/* ---------------------------------------------------------------- assemble */

const BUILDERS = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'TRAVELUR', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'TRAVELUR';
  pptx.title = 'Travelur - Travel Presentation Template';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  BUILDERS.forEach((fn) => fn(pptx.addSlide()));
  return pptx.writeFile({ fileName: path.join(__dirname, '06978ade-e540-4b47-8c8e-7b37964d16f7_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
