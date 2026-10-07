/**
 * Fundeluze — finance pitch deck (15 slides, 13.333in x 7.5in)
 * Rebuilt with pptxgenjs.  Photographs in the source deck are stand-in
 * "replace your image here" plates, so every raster image is redrawn here as a
 * flat grey plate of the same size/geometry, captioned `[image]`.
 *
 *   node 031f74dd-a8d5-4104-8b38-3d0d1d2ab830_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------------------ theme */

const C = {
  ink: '181818', // page background
  purple: '8361F5', // accent
  navy: '13165C', // deep accent / gradient end
  white: 'FFFFFF',
  w95: 'F2F2F2',
  w85: 'D9D9D9',
  w75: 'BFBFBF',
  w50: '808080',
  plate: 'CCCCCC', // image stand-in
  plateInk: 'AFAFAF', // caption on the image stand-in
  green: '92D050',
};

const F = {
  head: 'Manrope SemiBold',
  sub: 'Manrope Medium',
  body: 'Roboto',
  step: 'Poppins Medium',
};

const LOREM_S = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla auctor nisi velit. ';
const LOREM_M = LOREM_S + 'Phasellus id elementum neque, nec facilisis arcu. Sed euismod';
const LOREM_L = LOREM_M + ' faucibus nisl in tristique. Etiam sagittis tempor dapibus. ';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla auctor nisi velit. Phasellus id elementum neque, nec';
const BLURB = 'Empowering seamless digital payments with smart technology that adapts to your needs\u2014secure, scalable, and simple.';
const FOOT_NOTE = 'Choose your plan, preview dashboard modules, and personalize your transaction flow\u2014all from your mobile screen.';

/* ---------------------------------------------------------------- helpers */

const hex = (v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0').toUpperCase();

/** linear blend of two "RRGGBB" strings; t=0 -> a, t=1 -> b */
function mix(a, b, t) {
  const ch = (s, i) => parseInt(s.substr(i * 2, 2), 16);
  return hex(ch(a, 0) + (ch(b, 0) - ch(a, 0)) * t) + hex(ch(a, 1) + (ch(b, 1) - ch(a, 1)) * t) + hex(ch(a, 2) + (ch(b, 2) - ch(a, 2)) * t);
}

/** solid shape; omit `fill` for an outline-only shape */
function shape(slide, kind, o) {
  const opts = { x: o.x, y: o.y, w: o.w, h: o.h };
  if (o.fill) opts.fill = { color: o.fill };
  if (o.r) opts.rectRadius = o.r;
  if (o.line) opts.line = o.line;
  if (o.rotate) opts.rotate = o.rotate;
  if (o.flipV) opts.flipV = true;
  slide.addShape(kind, opts);
}

/**
 * Text box.  Every box in the deck is top-anchored and set to "resize shape to
 * fit text", which is also what keeps `wrap: false` boxes centred on their
 * declared frame instead of hanging off the left edge.
 */
function text(slide, body, o) {
  slide.addText(body, Object.assign({ valign: 'top', fit: 'resize', fontFace: F.body, color: C.white, fontSize: 11 }, o));
}

/**
 * pptxgenjs cannot emit a gradient fill, so every `<a:gradFill>` of the source
 * deck is painted as a mosaic of flat tiles.  GRAD_ANGLE is the deck's
 * `<a:lin ang="3600000">`, i.e. 60 degrees clockwise from the x axis.
 */
const GRAD_ANGLE = 60;
const GX = Math.cos((GRAD_ANGLE * Math.PI) / 180);
const GY = Math.sin((GRAD_ANGLE * Math.PI) / 180);

/** gradient position (0..1) of a point inside a w x h box */
const gradT = (u, v, w, h) => (GX * u + GY * v) / (GX * w + GY * h);

/** full-bleed slide backdrop, painted as bands perpendicular to the axis */
function gradientBackdrop(slide, from, to, steps = 44) {
  const axis = GX * SLIDE_W + GY * SLIDE_H;
  const band = axis / steps;
  const long = Math.sqrt(SLIDE_W * SLIDE_W + SLIDE_H * SLIDE_H);
  for (let i = 0; i < steps; i++) {
    const t = (i + 0.5) / steps;
    const d = (t - 0.5) * axis;
    slide.addShape('rect', {
      x: SLIDE_W / 2 + d * GX - long / 2,
      y: SLIDE_H / 2 + d * GY - band / 2,
      w: long,
      h: band * 1.4,
      rotate: GRAD_ANGLE + 90,
      fill: { color: mix(from, to, t) },
    });
  }
}

/**
 * Gradient card: a grid of flat tiles, then the rounded corners are carved
 * back out with a `back`-coloured square plus a disc of the local tile colour.
 * `corners` lists which corners are rounded as [xSide, ySide] with 0 = left/top.
 */
function gradientPanel(slide, o) {
  const nx = o.nx || 8;
  const ny = o.ny || 16;
  const tw = o.w / nx;
  const th = o.h / ny;
  const at = (u, v) => mix(o.from, o.to, gradT(u, v, o.w, o.h));
  for (let ix = 0; ix < nx; ix++) {
    for (let iy = 0; iy < ny; iy++) {
      shape(slide, 'rect', {
        x: o.x + ix * tw, y: o.y + iy * th, w: tw * 1.03, h: th * 1.03,
        fill: at((ix + 0.5) * tw, (iy + 0.5) * th),
      });
    }
  }
  const r = o.r || 0;
  const corners = o.corners || [[0, 0], [1, 0], [0, 1], [1, 1]];
  if (r) {
    corners.forEach(([sx, sy]) => {
      const u = sx ? o.w - r : r;
      const v = sy ? o.h - r : r;
      shape(slide, 'rect', { x: o.x + u - (sx ? 0 : r), y: o.y + v - (sy ? 0 : r), w: r, h: r, fill: o.back || C.ink });
      shape(slide, 'ellipse', { x: o.x + u - r, y: o.y + v - r, w: 2 * r, h: 2 * r, fill: at(u, v) });
    });
  }
}

/** gradient disc — small enough that a vertical ramp reads as the diagonal one */
function gradientDisc(slide, o, steps = 8) {
  for (let i = steps - 1; i >= 0; i--) {
    const t = i / (steps - 1);
    shape(slide, 'ellipse', { x: o.x, y: o.y, w: o.w, h: o.h * (0.25 + 0.75 * t), fill: mix(o.from, o.to, t) });
  }
}

/** grey stand-in for a photograph */
function imagePlate(slide, o) {
  shape(slide, o.kind || 'roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, r: o.r, fill: C.plate });
  if (o.w >= 1.3 && o.h >= 0.9) {
    text(slide, '[image]', {
      x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
      align: 'center', valign: 'middle', fontSize: 12, color: C.plateInk,
    });
  }
}

/** 3x3 dot cluster in the top-right corner */
function dotGrid(slide, x, y, w, color) {
  const d = w * 0.2275;
  const gap = w * 0.3863;
  for (let col = 0; col < 3; col++) {
    for (let row = 0; row < 3; row++) {
      shape(slide, 'ellipse', { x: x + col * gap, y: y + row * gap * 0.966, w: d, h: d, fill: color });
    }
  }
}

/** header repeated on every slide; `light` is the cover/closing variant */
function navBar(slide, light) {
  const brand = light ? C.white : C.purple;
  const off = light ? C.white : C.w75;
  const nav = { fontFace: F.head, fontSize: 9, wrap: false, h: 0.252, y: 0.501 };
  text(slide, 'FUNDELUZE', Object.assign({ x: 0.56, w: 0.919, color: brand }, nav));
  text(slide, 'About Us', Object.assign({ x: 5.141, w: 0.751, align: 'center', color: off }, nav));
  text(slide, 'Dashboard', Object.assign({ x: 6.237, w: 0.859, align: 'center', color: C.white }, nav));
  text(slide, 'Finance', Object.assign({ x: 7.442, w: 0.674, align: 'center', color: off }, nav));
  dotGrid(slide, 12.499, 0.536, 0.187, light ? C.w95 : C.purple);
}

/* text presets --------------------------------------------------------- */

const title28 = (slide, body, o) => text(slide, body, Object.assign({ fontFace: F.head, fontSize: 28, color: C.white }, o));
const sub18 = (slide, body, o) => text(slide, body, Object.assign({ fontFace: F.sub, fontSize: 18, color: C.w95, lineSpacingMultiple: 1.5 }, o));
const para = (slide, body, o) => text(slide, body, Object.assign({ fontFace: F.body, fontSize: 11, color: C.w50, lineSpacingMultiple: 1.5 }, o));
const stat54 = (slide, body, o) => text(slide, body, Object.assign({ fontFace: F.sub, fontSize: 54, color: C.purple }, o));

/* --------------------------------------------------- drawn icon glyphs */

/** rocket flying up-right: teardrop hull, two fins, exhaust swoosh, porthole */
function rocketGlyph(slide, x, y, d, body, hole) {
  shape(slide, 'teardrop', { x: x + d * 0.2, y, w: d * 0.8, h: d * 0.8, fill: body });
  shape(slide, 'triangle', { x, y: y + d * 0.26, w: d * 0.28, h: d * 0.2, fill: body, rotate: 315 });
  shape(slide, 'triangle', { x: x + d * 0.52, y: y + d * 0.72, w: d * 0.2, h: d * 0.28, fill: body, rotate: 135 });
  shape(slide, 'moon', { x, y: y + d * 0.56, w: d * 0.36, h: d * 0.36, fill: body, rotate: 45 });
  shape(slide, 'ellipse', { x: x + d * 0.49, y: y + d * 0.21, w: d * 0.24, h: d * 0.24, fill: hole });
}

/** jigsaw piece: knob on top and left, socket on right and bottom */
function puzzleGlyph(slide, x, y, d, ink, back) {
  const k = d * 0.31; // knob / socket diameter
  const disc = (cx, cy, fill) => shape(slide, 'ellipse', { x: x + d * cx - k / 2, y: y + d * cy - k / 2, w: k, h: k, fill });
  disc(0.617, 0.139, ink); // top knob
  disc(0.157, 0.61, ink); // left knob
  shape(slide, 'roundRect', { x: x + d * 0.24, y: y + d * 0.285, w: d * 0.75, h: d * 0.675, r: d * 0.07, fill: ink });
  disc(0.918, 0.61, back); // right socket
  disc(0.617, 0.955, back); // bottom socket
}

/** stack of coins: four ellipses separated by background-coloured slivers */
function coinsGlyph(slide, x, y, w, h, body, back) {
  const eh = w * 0.42;
  const step = (h - eh) / 3;
  for (let i = 3; i >= 0; i--) {
    shape(slide, 'ellipse', { x, y: y + i * step, w, h: eh, fill: body });
    if (i > 0) shape(slide, 'ellipse', { x, y: y + i * step - step * 0.45, w, h: eh, fill: back });
  }
}

/* device mock-ups (drawn, not embedded) -------------------------------- */

function phoneMock(slide) {
  shape(slide, 'roundRect', { x: 4.86, y: 1.336, w: 4.45, h: 6.164, r: 0.43, fill: '060606', line: { color: 'C6C7C4', width: 3.2 } });
  imagePlate(slide, { x: 5.108, y: 1.534, w: 4.023, h: 5.966, r: 0.427 });
  shape(slide, 'roundRect', { x: 6.271, y: 1.44, w: 1.637, h: 0.42, r: 0.2, fill: '000000' });
  shape(slide, 'ellipse', { x: 7.03, y: 1.63, w: 0.1, h: 0.1, fill: '18242C' });
}

function laptopMock(slide) {
  shape(slide, 'roundRect', { x: 0.368, y: 3.578, w: 12.6, h: 3.922, r: 0.28, fill: '070707', line: { color: 'CFCFCF', width: 3.2 } });
  shape(slide, 'rect', { x: 6.7, y: 3.72, w: 6.12, h: 0.36, fill: '3F3E3C' }); // bezel reflection
  imagePlate(slide, { kind: 'rect', x: 0.72, y: 4.055, w: 11.914, h: 3.445 });
}

/* ------------------------------------------------------------- 1. cover */

function slide01(slide) {
  gradientBackdrop(slide, C.purple, C.navy);
  imagePlate(slide, { x: 0.875, y: 1.612, w: 1.677, h: 1.148, r: 0.156 });
  imagePlate(slide, { x: 1.656, y: 3.219, w: 4.235, h: 2.594, r: 0.189 });
  navBar(slide, true);
  text(slide, 'Your Trusted Gateway to Transactions', {
    x: 6.593, y: 1.921, w: 7.605, h: 3.13, fontFace: F.head, fontSize: 60, color: C.white,
  });
  text(slide, FOOT_NOTE, {
    x: 1.156, y: 6.125, w: 4.677, h: 0.625, fontSize: 11, color: C.w85, align: 'right', lineSpacingMultiple: 1.5,
  });
  text(slide, 'Start Presentation', {
    x: 6.593, y: 6.125, w: 2.612, h: 0.463, fontFace: F.head, fontSize: 16, color: C.w95, lineSpacingMultiple: 1.5,
  });
  slide.addShape('line', { x: 9.094, y: 6.406, w: 0.912, h: 0, line: { color: C.white, width: 1, endArrowType: 'triangle' } });
}

/* ------------------------------------------------------- 2. smart finance */

function slide02(slide) {
  imagePlate(slide, { x: 0.667, y: 3.436, w: 8.076, h: 3.439, r: 0.164 });
  navBar(slide);
  text(slide, 'Smart Finance', { x: 0.56, y: 1.628, w: 3.126, h: 0.572, fontFace: F.sub, fontSize: 28, color: C.purple });
  sub18(slide, BLURB, { x: 4.591, y: 1.628, w: 8.076, h: 0.963 });

  shape(slide, 'roundRect', { x: 9.069, y: 3.436, w: 3.598, h: 1.731, r: 0.145, fill: C.purple });
  text(slide, '#1', { x: 9.362, y: 3.659, w: 1.662, h: 1.01, fontFace: F.sub, fontSize: 54, color: C.white });
  sub18(slide, 'Payment Helper', { x: 10.639, y: 3.83, w: 1.662, h: 0.707, lineSpacingMultiple: null });
  text(slide, 'No paperwork just ID verification and access', {
    x: 9.357, y: 4.536, w: 3.309, h: 0.348, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5,
  });

  sub18(slide, 'Modern', { x: 9.362, y: 5.527, w: 1.245, h: 0.508 });
  para(slide, LOREM_S, { x: 10.712, y: 5.614, w: 2.174, h: 0.903 });
}

/* --------------------------------------------------------- 3. about cards */

function slide03(slide) {
  imagePlate(slide, { x: 0.667, y: 1.333, w: 5.419, h: 5.542, r: 0.165 });
  gradientPanel(slide, { x: 9.567, y: 1.333, w: 3.1, h: 5.542, r: 0.198, from: C.navy, to: C.ink, nx: 6, ny: 14 });
  shape(slide, 'roundRect', { x: 6.398, y: 1.333, w: 2.857, h: 5.542, r: 0.182, fill: C.purple });
  navBar(slide);

  title28(slide, 'About Fundeluze', { x: 6.757, y: 1.837, w: 3.126, h: 1.043 });
  sub18(slide, 'A tech-forward finance solution reimagining how payments work', { x: 6.692, y: 4.215, w: 2.563, h: 1.872 });
  text(slide, 'Brief Detail', { x: 6.692, y: 6.146, w: 1.581, h: 0.348, fontSize: 11, color: C.w75, lineSpacingMultiple: 1.5 });

  rocketGlyph(slide, 10.249, 3.093, 0.383, C.purple, C.navy);
  para(slide, LOREM_S, { x: 10.108, y: 3.637, w: 2.391, h: 0.903 });
  sub18(slide, 'Modern', { x: 10.108, y: 5.085, w: 1.245, h: 0.508 });
  para(slide, LOREM_S, { x: 10.108, y: 5.586, w: 2.391, h: 0.903 });
}

/* ------------------------------------------------------- 4. core features */

function slide04(slide) {
  const cards = [
    { x: 0.667, w: 3.352, tx: 0.783, tw: 2.174, mx: 0.768, copy: LOREM_S },
    { x: 4.331, w: 4.762, tx: 4.563, tw: 3.727, mx: 4.563, copy: LOREM_M },
    { x: 9.405, w: 3.262, tx: 9.597, tw: 3.069, mx: 9.597, copy: LOREM_S + 'Phasellus id elementum neque, ' },
  ];
  cards.forEach((c) => imagePlate(slide, { x: c.x, y: 2.223, w: c.w, h: 2.791, r: 0.126 }));
  navBar(slide);
  title28(slide, 'Core Features We Provided', { x: 2.076, y: 1.429, w: 9.181, h: 0.572, align: 'center' });
  cards.forEach((c) => {
    sub18(slide, 'Modern', { x: c.mx, y: 5.277, w: 1.245, h: 0.508 });
    para(slide, c.copy, { x: c.tx, y: 5.914, w: c.tw, h: 0.903 });
  });
}

/* ------------------------------------------------------ 5. virtual cards */

function slide05(slide) {
  shape(slide, 'roundRect', { x: 5.326, y: 3.606, w: 7.341, h: 3.256, r: 0.188, fill: C.purple });
  imagePlate(slide, { x: 7.659, y: 3.785, w: 4.839, h: 2.899, r: 0.222 });
  navBar(slide);

  title28(slide, 'Virtual Card Issuance', { x: 0.56, y: 1.583, w: 3.126, h: 1.043 });
  sub18(slide, 'Modern', { x: 5.631, y: 1.583, w: 1.245, h: 0.508 });
  para(slide, LOREM_L + 'Etiam ut pretium massa.', { x: 5.631, y: 2.156, w: 6.492, h: 0.903 });

  stat54(slide, '20K+', { x: 0.592, y: 4.949, w: 2.882, h: 1.01 });
  para(slide, LOREM_M, { x: 0.56, y: 5.959, w: 3.727, h: 0.903 });
  sub18(slide, 'Based on the Current Security', { x: 5.6, y: 5.108, w: 1.651, h: 1.417 });
}

/* ---------------------------------------------------------- 6. visual flow */

function slide06(slide) {
  imagePlate(slide, { x: 0.667, y: 1.333, w: 6.158, h: 5.542, r: 0.21 });
  navBar(slide);
  title28(slide, 'Visual Flow', { x: 7.294, y: 1.583, w: 3.126, h: 0.572 });
  para(slide, LOREM_L, { x: 7.294, y: 2.313, w: 5.119, h: 0.903 });

  const steps = [
    { x: 7.174, r: 0.123, label: 'Step 1', cap: ['Account ', 'Log In'], capX: 7.255, dot: 7.958, textX: 7.15 },
    { x: 9.122, r: 0.129, label: 'Step 2', cap: ['Setup ', 'Payment'], capX: 9.203, dot: 9.883, textX: 9.122 },
    { x: 11.07, r: 0.129, label: 'Step 3', cap: ['Finishing ', 'Setup'], capX: 11.151, dot: 11.831, textX: 11.07, dark: true },
  ];
  steps.forEach((s) => {
    if (s.dark) gradientPanel(slide, { x: s.x, y: 4.631, w: 1.622, h: 2.244, r: s.r, from: C.navy, to: C.ink, nx: 5, ny: 8 });
    else shape(slide, 'roundRect', { x: s.x, y: 4.631, w: 1.622, h: 2.244, r: s.r, fill: C.purple });
    text(slide, s.label, { x: s.textX, y: 4.998, w: 1.669, h: 0.572, fontFace: F.step, fontSize: 28, align: 'center' });
    shape(slide, 'ellipse', { x: s.dot, y: 5.728, w: 0.054, h: 0.052, fill: C.w85 });
    text(slide, [{ text: s.cap[0], options: { breakLine: true } }, { text: s.cap[1] }], {
      x: s.capX, y: 6.026, w: 1.46, h: 0.572, fontFace: F.step, fontSize: 14, align: 'center', valign: 'top', color: C.white,
    });
  });
}

/* -------------------------------------------------------- 7. payment guide */

function slide07(slide) {
  navBar(slide);
  title28(slide, 'Payment Guide', { x: 0.85, y: 1.583, w: 3.126, h: 0.572 });
  para(slide, LOREM_L, { x: 5.629, y: 1.628, w: 6.649, h: 0.903 });

  slide.addShape('line', { x: 0, y: 3.319, w: SLIDE_W, h: 0, line: { color: C.purple, width: 1.5 } });

  const cards = [
    { x: 0.683, num: '01', numX: 0.997, head: 'Access Your Account', headX: 2.153, copyX: 1.007, dot: 2.278, solid: true, copyColor: C.white },
    { x: 4.662, num: '02', numX: 4.976, head: 'Choose Pay Methods', headX: 6.102, copyX: 4.955, dot: 6.428, copyColor: C.w75 },
    { x: 8.641, num: '03', numX: 9.025, head: 'Click Confirm Payment', headX: 10.171, copyX: 9.025, dot: 10.378, copyColor: C.w75 },
  ];
  cards.forEach((c) => {
    if (c.solid) shape(slide, 'roundRect', { x: c.x, y: 3.987, w: 3.667, h: 2.715, r: 0.173, fill: C.purple });
    else gradientPanel(slide, { x: c.x, y: 3.987, w: 3.667, h: 2.715, r: 0.173, from: C.navy, to: C.ink, nx: 9, ny: 8 });
  });
  cards.forEach((c) => {
    shape(slide, 'ellipse', { x: c.dot, y: 3.271, w: 0.135, h: 0.135, fill: C.white });
    text(slide, c.num, { x: c.numX, y: 4.335, w: 1.519, h: 1.01, fontFace: F.sub, fontSize: 54, color: C.white });
    sub18(slide, c.head, { x: c.headX, y: 4.466, w: 1.902, h: 0.707, color: C.white, lineSpacingMultiple: null });
    para(slide, LOREM_CARD, { x: c.copyX, y: 5.298, w: 3.16, h: 0.903, color: c.copyColor });
  });
}

/* ----------------------------------------------------------- 8. the team */

function slide08(slide) {
  imagePlate(slide, { x: 2.667, y: 1.333, w: 3.094, h: 4.486, r: 0.166 });
  [0.311, 6.074, 8.431, 10.788].forEach((x) => imagePlate(slide, { x, y: 2.544, w: 2.045, h: 2.654, r: 0.103 }));
  navBar(slide);
  title28(slide, 'Meet Our Finance UX Pros', { x: 6.667, y: 1.583, w: 6.374, h: 0.572 });

  // the lead portrait is larger, so its caption sits lower and is set bigger
  text(slide, 'Morgan Hayes', { x: 2.894, y: 6.036, w: 2.932, h: 0.505, fontFace: F.sub, fontSize: 24, color: C.purple });
  text(slide, 'Threat Intelligence Analyst', { x: 2.894, y: 6.495, w: 3.35, h: 0.286, fontFace: F.sub, fontSize: 11, color: C.white });

  const team = [
    { x: 0.56, roleX: 0.593, name: 'Rizelle Sue', role: 'Marketing' },
    { x: 6.244, roleX: 6.244, name: 'Alexander Jo', role: 'Assistant' },
    { x: 8.673, roleX: 8.673, name: 'Fjord Radish', role: 'Lead Security ' },
    { x: 10.957, roleX: 10.957, name: 'Noah Wilson', role: 'UX Expertise' },
  ];
  team.forEach((m) => {
    text(slide, m.name, { x: m.x, y: 5.416, w: 1.876, h: 0.404, fontFace: F.sub, fontSize: 18, color: C.white });
    text(slide, m.role, { x: m.roleX, y: 5.78, w: 1.257, h: 0.269, fontFace: F.sub, fontSize: 10, color: C.purple });
  });
}

/* -------------------------------------------------------------- 9. pricing */

function slide09(slide) {
  navBar(slide);
  title28(slide, 'Flexible Pricing Tiers', { x: 0.824, y: 1.488, w: 3.126, h: 1.043 });
  para(slide, LOREM_L, { x: 5.619, y: 1.628, w: 6.649, h: 0.903 });

  // cards run off the bottom edge, so only their top corners are rounded
  const tiers = [
    { x: 0.667, solid: true, bullet: C.navy, name: 'Premium Pack', nameX: 1.684, nameW: 1.689, price: '99.99', priceX: 1.575, priceW: 2.053, dollarX: 1.317, copyX: 0.974, trialX: 0.985, bulletX: 1.121, featX: 1.439, feats: [['Unlock All Paid Features', 2.393], ['Unlimited Transactions', 2.344]] },
    { x: 4.702, bullet: C.purple, name: 'Starter Pack', nameX: 5.763, nameW: 1.508, price: '19.99', priceX: 5.627, priceW: 1.909, dollarX: 5.37, copyX: 4.962, trialX: 4.974, bulletX: 5.109, featX: 5.428, feats: [['Multi-currency Wallet', 2.206], ['500 Transactions/Month', 2.497]] },
    { x: 8.727, bullet: C.purple, name: 'Pro Pack', nameX: 10.049, nameW: 1.11, price: '49.99', priceX: 9.714, priceW: 2.037, dollarX: 9.457, copyX: 9.049, trialX: 9.061, bulletX: 9.196, featX: 9.515, feats: [['Basic Fraud Monitoring', 2.304], ['5000 Transactions/Month', 2.619]] },
  ];
  tiers.forEach((t) => {
    if (t.solid) shape(slide, 'roundRect', { x: t.x, y: 3.021, w: 3.723, h: 4.86, r: 0.178, fill: C.purple });
    else gradientPanel(slide, { x: t.x, y: 3.021, w: 3.723, h: 4.86, r: 0.178, from: C.navy, to: C.ink, nx: 9, ny: 14, corners: [[0, 0], [1, 0]] });
  });
  tiers.forEach((t) => {
    text(slide, t.name, { x: t.nameX, y: 3.456, w: t.nameW, h: 0.37, fontFace: F.sub, fontSize: 16, align: 'center', wrap: false });
    text(slide, t.price, { x: t.priceX, y: 3.732, w: t.priceW, h: 0.909, fontFace: F.sub, fontSize: 48, wrap: false });
    text(slide, '$', { x: t.dollarX, y: 4.031, w: 0.405, h: 0.505, fontFace: F.sub, fontSize: 24, wrap: false });
    text(slide, 'Extra One Month Trial', { x: t.trialX, y: 4.511, w: 3.086, h: 0.303, fontFace: F.sub, fontSize: 12, align: 'center' });
    para(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla auctor nisi velit. Phasellus id elementum neque', {
      x: t.copyX, y: 4.94, w: 3.109, h: 0.903, color: C.white, align: 'center',
    });
    t.feats.forEach(([label, width], i) => {
      shape(slide, 'donut', { x: t.bulletX, y: 6.22 + i * 0.367, w: 0.15, h: 0.15, r: 0.02, fill: t.bullet });
      text(slide, label, { x: t.featX, y: 6.131 + i * 0.38, w: width, h: 0.337, fontFace: F.sub, fontSize: 14, wrap: false });
    });
  });
}

/* --------------------------------------------------------- 10. case studies */

function slide10(slide) {
  // full-bleed panel from the layout: only its top-right corner is rounded
  gradientPanel(slide, { x: 0.002, y: 1.333, w: 9.67, h: 6.167, r: 0.256, from: C.navy, to: C.ink, nx: 20, ny: 16, corners: [[1, 0]] });
  imagePlate(slide, { x: 0.667, y: 3.75, w: 3.668, h: 2.646, r: 0.212 });
  imagePlate(slide, { x: 4.647, y: 3.75, w: 3.668, h: 2.646, r: 0.212 });
  navBar(slide);

  sub18(slide, BLURB, { x: 0.823, y: 1.907, w: 7.812, h: 0.963 });
  text(slide, 'Rapid Growing can Adapt to the Environment Almost Instantly', {
    x: 0.823, y: 2.916, w: 5.338, h: 0.348, fontSize: 11, color: C.w75, lineSpacingMultiple: 1.5,
  });

  title28(slide, 'Case Studies', { x: 10.455, y: 1.583, w: 2.231, h: 1.043 });
  stat54(slide, '200K', { x: 10.451, y: 4.561, w: 2.882, h: 1.01 });
  para(slide, LOREM_S, { x: 10.464, y: 5.571, w: 2.221, h: 0.903 });
}

/* --------------------------------------------------------- 11. trusted tech */

function slide11(slide) {
  imagePlate(slide, { x: 0.667, y: 3.23, w: 6.0, h: 3.645, r: 0.174 });
  imagePlate(slide, { x: 6.938, y: 3.23, w: 2.755, h: 3.645, r: 0.132 });
  navBar(slide);

  title28(slide, 'Trusted Tech', { x: 0.777, y: 1.507, w: 3.126, h: 0.572 });
  para(slide, LOREM_L + 'Etiam ut pretium massa.', { x: 5.572, y: 1.504, w: 6.649, h: 0.903 });

  shape(slide, 'roundRect', { x: 9.964, y: 3.238, w: 2.702, h: 3.645, r: 0.129, fill: C.purple });
  puzzleGlyph(slide, 11.952, 3.543, 0.339, C.white, C.purple);
  sub18(slide, 'Trusted By Payment Users and Industry', { x: 10.239, y: 4.556, w: 2.153, h: 1.01, color: C.white, lineSpacingMultiple: null });
  text(slide, '213+', { x: 10.205, y: 5.637, w: 2.882, h: 1.01, fontFace: F.sub, fontSize: 54, color: C.white });
}

/* -------------------------------------------------------- 12. user benefits */

function slide12(slide) {
  gradientPanel(slide, { x: 0.648, y: 2.843, w: 3.797, h: 1.641, r: 0.186, from: C.navy, to: C.ink, nx: 10, ny: 6 });
  phoneMock(slide);
  for (let i = 0; i < 5; i++) {
    imagePlate(slide, { kind: 'ellipse', x: 0.844 + i * 0.5105, y: 3.539, w: 0.708, h: 0.708 });
  }
  imagePlate(slide, { x: 0.648, y: 4.833, w: 3.797, h: 2.042, r: 0.175 });
  navBar(slide);

  // wallet tile
  shape(slide, 'roundRect', { x: 0.648, y: 1.31, w: 2.422, h: 1.211, r: 0.113, fill: C.purple });
  text(slide, 'Wallet (USD)', { x: 0.823, y: 1.504, w: 1.963, h: 0.348, fontSize: 11, color: C.w75, lineSpacingMultiple: 1.5 });
  sub18(slide, '$2,345.67', { x: 0.823, y: 1.792, w: 1.371, h: 0.508 });
  shape(slide, 'triangle', { x: 2.166, y: 2.115, w: 0.098, h: 0.05, fill: C.green });
  text(slide, '2.7%', { x: 2.246, y: 1.917, w: 0.715, h: 0.348, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });

  // send tile
  shape(slide, 'roundRect', { x: 3.367, y: 1.31, w: 1.077, h: 1.211, r: 0.113, fill: C.purple });
  slide.addShape('line', { x: 3.874, y: 1.567, w: 0, h: 0.277, line: { color: C.white, width: 2.25, beginArrowType: 'triangle' } });
  sub18(slide, 'Send', { x: 3.336, y: 1.792, w: 1.077, h: 0.508, align: 'center' });

  // quick transaction row
  sub18(slide, 'Quick Transaction', { x: 0.823, y: 2.937, w: 3.51, h: 0.508 });
  gradientDisc(slide, { x: 3.397, y: 3.539, w: 0.708, h: 0.708, from: C.navy, to: C.ink });
  shape(slide, 'ellipse', { x: 3.397, y: 3.539, w: 0.708, h: 0.708, line: { color: C.white, width: 0.75 } });
  shape(slide, 'mathPlus', { x: 3.649, y: 3.784, w: 0.201, h: 0.201, fill: C.purple });

  title28(slide, 'User Benefits', { x: 9.936, y: 1.795, w: 2.202, h: 1.043 });
  coinsGlyph(slide, 10.123, 5.02, 0.472, 0.415, C.purple, C.ink);
  para(slide, LOREM_S, { x: 10.044, y: 5.571, w: 2.221, h: 0.903 });
}

/* ------------------------------------------------------- 13. instant access */

function slide13(slide) {
  laptopMock(slide);
  navBar(slide);
  title28(slide, '                          Access Comprehensive Digital Finance Services That Instantly Arrived with Fast Setup and Seamless Interface', {
    x: 0.567, y: 1.507, w: 12.118, h: 1.441, lineSpacingMultiple: 1.5,
  });
  sub18(slide, 'Instant Access', { x: 0.567, y: 1.663, w: 2.4, h: 0.508, color: C.purple });
}

/* -------------------------------------------------------------- 14. contact */

function slide14(slide) {
  imagePlate(slide, { x: 3.667, y: 1.333, w: 6.0, h: 5.542, r: 0.169 });
  navBar(slide);
  title28(slide, [{ text: 'Reach Us ', options: { breakLine: true } }, { text: 'For Further Information' }], {
    x: 0.56, y: 1.583, w: 3.308, h: 1.515,
  });
  para(slide, LOREM_CARD, { x: 0.584, y: 5.628, w: 2.43, h: 1.181 });

  text(slide, 'Phone', { x: 10.21, y: 4.102, w: 1.755, h: 0.337, fontFace: F.sub, fontSize: 14, color: C.white });
  text(slide, 'Customer Support', { x: 10.711, y: 4.469, w: 2.065, h: 0.278, fontSize: 10.5, color: C.white });
  text(slide, '+1 555-123-4567', { x: 10.711, y: 4.71, w: 1.656, h: 0.303, fontSize: 12, color: C.w75 });
  text(slide, 'Company Email', { x: 10.711, y: 5.084, w: 2.241, h: 0.278, fontSize: 10.5, color: C.white });
  text(slide, 'name@company.mail', { x: 10.711, y: 5.325, w: 2.694, h: 0.303, fontSize: 12, color: C.w75 });
  text(slide, 'Address', { x: 10.169, y: 5.844, w: 1.755, h: 0.337, fontFace: F.sub, fontSize: 14, color: C.white });
  text(slide, '123 Street Name, Anywhere 400, Any City, State 12345', { x: 10.711, y: 6.279, w: 2.12, h: 0.471, fontSize: 11, color: C.w75 });

  // contact glyphs: handset, envelope, map pin
  shape(slide, 'moon', { x: 10.319, y: 4.559, w: 0.239, h: 0.237, fill: C.purple, rotate: 225 });
  shape(slide, 'roundRect', { x: 10.315, y: 5.18, w: 0.248, h: 0.153, r: 0.018, fill: C.purple });
  shape(slide, 'triangle', { x: 10.331, y: 5.185, w: 0.216, h: 0.082, fill: C.ink, rotate: 180 });
  shape(slide, 'teardrop', { x: 10.315, y: 6.348, w: 0.243, h: 0.283, fill: C.purple, rotate: 135 });
  shape(slide, 'ellipse', { x: 10.387, y: 6.418, w: 0.098, h: 0.098, fill: C.ink });
}

/* --------------------------------------------------------------- 15. outro */

function slide15(slide) {
  gradientBackdrop(slide, C.purple, C.navy);
  imagePlate(slide, { x: 0.667, y: 1.333, w: 5.458, h: 2.938, r: 0.196 });
  navBar(slide, true);
  text(slide, [{ text: 'See You Again ', options: { breakLine: true } }, { text: 'Next Time!' }], {
    x: 6.578, y: 3.819, w: 8.667, h: 2.121, fontFace: F.head, fontSize: 60, color: C.white,
  });
  text(slide, FOOT_NOTE, {
    x: 2.926, y: 5.847, w: 3.125, h: 0.903, fontSize: 11, color: C.white, align: 'right', lineSpacingMultiple: 1.5,
  });
  text(slide, 'End Presentation', {
    x: 6.596, y: 6.287, w: 2.612, h: 0.463, fontFace: F.head, fontSize: 16, color: C.w95, lineSpacingMultiple: 1.5,
  });
  slide.addShape('line', { x: 9.097, y: 6.569, w: 0.912, h: 0, line: { color: C.white, width: 1, endArrowType: 'triangle' } });
}

/* ------------------------------------------------------------------ build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'FUNDELUZE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'FUNDELUZE';
  pptx.author = 'Fundeluze';
  pptx.title = 'Fundeluze — Your Trusted Gateway to Transactions';

  BUILDERS.forEach((buildSlide) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.ink };
    buildSlide(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '031f74dd-a8d5-4104-8b38-3d0d1d2ab830_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f));
