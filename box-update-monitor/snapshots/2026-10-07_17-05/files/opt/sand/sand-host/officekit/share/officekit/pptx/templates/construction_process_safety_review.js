/**
 * "Best Corporate" construction deck — rebuilt with pptxgenjs.
 *
 * Run:  node 065f8849-9b33-405e-b0af-8bb25a10537f_grok_final.js
 * Out:  065f8849-9b33-405e-b0af-8bb25a10537f_grok_final.pptx (next to this file)
 *
 * Raster artwork from the original deck (photo placeholders, icon PNGs, device
 * mockups, oversized SWOT letters) is re-created here with native pptxgenjs
 * shapes and text instead of embedded bitmaps.
 */

'use strict';

const path = require('path');
const pptxgen = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const TEAL = '17726D'; // brand green
const CREAM = 'EEECDF'; // deck background
const INK = '0D0D0D'; // headline black (tx1 lumMod 95%)
const BODY = '404040'; // body grey  (tx1 lumMod 75% / lumOff 25%)
const WHITE = 'FFFFFF';
const WHITE95 = 'F2F2F2'; // bg1 lumMod 95%
const WHITE85 = 'D9D9D9'; // bg1 lumMod 85%
const ORANGE = 'F0A028'; // accent4
const GREY25 = '3C3C3C'; // accent3 lumMod 25%
const NAVY = '00153E';
const DEVICE = '2B2B2B'; // mock-up chassis

const HEAD = 'Montserrat'; // major/theme heading face
const TEXT = 'Open Sans'; // minor/theme body face

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------------------ *
 * Small building blocks
 * ------------------------------------------------------------------ */

/** Solid rectangle. */
function rect(slide, x, y, w, h, color, transparency) {
  slide.addShape('rect', { x, y, w, h, fill: { color, transparency }, line: { type: 'none' } });
}

/** Solid ellipse. */
function oval(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/** Straight rule. */
function line(slide, x, y, w, h, color, width) {
  slide.addShape('line', { x, y, w, h, line: { color, width } });
}

/** Small caps label sitting above every headline. */
function eyebrow(slide, x, y, w, text, color, opts) {
  slide.addText(text, Object.assign(
    { x, y, w, h: 0.337, fontFace: TEXT, fontSize: 14, color, valign: 'top' }, opts));
}

/** Big Montserrat headline; `text` may be a run array for two-tone titles. */
function headline(slide, x, y, w, h, text, opts) {
  slide.addText(text, Object.assign(
    { x, y, w, h, fontFace: HEAD, fontSize: 36, bold: true, color: INK, valign: 'top' }, opts));
}

/** Justified paragraph of body copy. */
function paragraph(slide, x, y, w, h, text, color, opts) {
  slide.addText(text, Object.assign(
    {
      x, y, w, h, fontFace: TEXT, fontSize: 12, color,
      align: 'justify', lineSpacingMultiple: 1.35, valign: 'top',
    }, opts));
}

/** Teal pill button with white caps. */
function button(slide, x, y, text, opts) {
  const o = Object.assign({ w: 1.889, h: 0.501, fill: TEAL, color: WHITE, fontSize: 12 }, opts);
  slide.addText(text, {
    x, y, w: o.w, h: o.h,
    shape: 'roundRect', rectRadius: o.h * 0.16667,
    fill: { color: o.fill }, line: { type: 'none' },
    fontFace: TEXT, fontSize: o.fontSize, bold: true, color: o.color,
    align: 'center', valign: 'middle',
  });
}

/**
 * Facebook / Instagram / Twitter roundels: three outlined circles, 0.412"
 * across, on 0.56" centres, all drawn in one flat colour.
 */
function socialIcons(slide, x, y, color) {
  const D = 0.412;
  const STEP = 0.56;
  const none = { type: 'none' };
  for (let i = 0; i < 3; i++) {
    const cx = x + i * STEP;
    slide.addShape('ellipse', { x: cx, y, w: D, h: D, fill: none, line: { color, width: 1.1 } });
    if (i === 0) {
      slide.addText('f', {
        x: cx, y: y - 0.01, w: D, h: D,
        fontFace: HEAD, fontSize: D * 44, bold: true, color,
        align: 'center', valign: 'middle',
      });
    } else if (i === 1) {
      slide.addShape('roundRect', {
        x: cx + 0.235 * D, y: y + 0.235 * D, w: 0.53 * D, h: 0.53 * D, rectRadius: 0.09 * D,
        fill: none, line: { color, width: 0.9 },
      });
      slide.addShape('ellipse', {
        x: cx + 0.355 * D, y: y + 0.355 * D, w: 0.29 * D, h: 0.29 * D,
        fill: none, line: { color, width: 0.9 },
      });
      slide.addShape('ellipse', {
        x: cx + 0.615 * D, y: y + 0.315 * D, w: 0.075 * D, h: 0.075 * D,
        fill: { color }, line: none,
      });
    } else {
      slide.addShape('moon', {
        x: cx + 0.20 * D, y: y + 0.28 * D, w: 0.58 * D, h: 0.44 * D, rotate: 200,
        fill: { color }, line: none,
      });
      slide.addShape('ellipse', {
        x: cx + 0.55 * D, y: y + 0.28 * D, w: 0.20 * D, h: 0.16 * D, fill: { color }, line: none,
      });
    }
  }
}

/**
 * Hard-hat worker pictogram: helmet, brim, face and hi-vis vest.
 * All measurements are fractions of the icon box `z`, taken off the original.
 */
function workerIcon(slide, x, y, z, color, bg) {
  const f = { color };
  const b = { color: bg };
  const none = { type: 'none' };
  const box = (X, Y, W, H, fill, radius) => slide.addShape(radius ? 'roundRect' : 'rect', {
    x: x + X * z, y: y + Y * z, w: W * z, h: H * z, rectRadius: (radius || 0) * z, fill, line: none,
  });
  slide.addShape('ellipse', { x: x + 0.30 * z, y: y + 0.075 * z, w: 0.40 * z, h: 0.24 * z, fill: f, line: none });
  box(0.290, 0.225, 0.420, 0.055, f, 0.027); // helmet brim
  box(0.325, 0.250, 0.350, 0.290, f, 0.145); // head
  box(0.375, 0.200, 0.250, 0.260, b, 0.115); // face cut-out
  box(0.440, 0.470, 0.120, 0.120, f);        // neck
  box(0.167, 0.550, 0.666, 0.340, f, 0.130); // shoulders
  box(0.167, 0.750, 0.666, 0.140, f);        // torso
  slide.addShape('triangle', { x: x + 0.30 * z, y: y + 0.54 * z, w: 0.18 * z, h: 0.13 * z, fill: b, line: none });
  slide.addShape('triangle', { x: x + 0.52 * z, y: y + 0.54 * z, w: 0.18 * z, h: 0.13 * z, fill: b, line: none });
  box(0.300, 0.660, 0.180, 0.240, b);        // vest opening, left
  box(0.520, 0.660, 0.180, 0.240, b);        // vest opening, right
  box(0.167, 0.790, 0.666, 0.045, f);        // reflective band
}

/** Tracked excavator pictogram: A-frame boom, cab, hull and roller track. */
function excavatorIcon(slide, x, y, z, color, bg) {
  const f = { color };
  const b = { color: bg };
  const none = { type: 'none' };
  const box = (X, Y, W, H, fill, radius) => slide.addShape(radius ? 'roundRect' : 'rect', {
    x: x + X * z, y: y + Y * z, w: W * z, h: H * z, rectRadius: (radius || 0) * z, fill, line: none,
  });
  slide.addShape('line', { x: x + 0.155 * z, y: y + 0.55 * z, w: 0.39 * z, h: -0.40 * z, line: { color, width: z * 4.3 } });
  slide.addShape('line', { x: x + 0.545 * z, y: y + 0.15 * z, w: 0.30 * z, h: 0.33 * z, line: { color, width: z * 4.3 } });
  slide.addShape('pie', { x: x + 0.745 * z, y: y + 0.40 * z, w: 0.16 * z, h: 0.19 * z, rotate: 330, fill: f, line: none });
  box(0.755, 0.440, 0.130, 0.160, f, 0.060); // bucket
  box(0.385, 0.400, 0.255, 0.160, f, 0.035); // cab
  box(0.420, 0.435, 0.185, 0.100, b);        // cab window
  box(0.125, 0.530, 0.535, 0.130, f, 0.035); // hull
  box(0.100, 0.735, 0.585, 0.145, f, 0.072); // track
  for (let i = 0; i < 6; i++) {
    slide.addShape('ellipse', {
      x: x + (0.145 + i * 0.088) * z, y: y + 0.775 * z, w: 0.055 * z, h: 0.06 * z, fill: b, line: none,
    });
  }
}

/** Bulldozer pictogram: cab, arched hood, roller track and angled blade. */
function dozerIcon(slide, x, y, z, color, bg) {
  const f = { color };
  const b = { color: bg };
  const none = { type: 'none' };
  const box = (X, Y, W, H, fill, radius) => slide.addShape(radius ? 'roundRect' : 'rect', {
    x: x + X * z, y: y + Y * z, w: W * z, h: H * z, rectRadius: (radius || 0) * z, fill, line: none,
  });
  box(0.135, 0.215, 0.215, 0.190, f, 0.030); // cab
  box(0.175, 0.255, 0.135, 0.150, b);        // cab window
  box(0.073, 0.385, 0.555, 0.230, f, 0.030); // hood and body
  slide.addShape('chord', { x: x + 0.11 * z, y: y + 0.40 * z, w: 0.50 * z, h: 0.30 * z, rotate: 180, fill: b, line: none });
  box(0.060, 0.545, 0.620, 0.145, f, 0.072); // track
  for (let i = 0; i < 3; i++) {
    slide.addShape('ellipse', { x: x + (0.14 + i * 0.19) * z, y: y + 0.575 * z, w: 0.07 * z, h: 0.075 * z, fill: b, line: none });
  }
  slide.addShape('parallelogram', { x: x + 0.665 * z, y: y + 0.40 * z, w: 0.29 * z, h: 0.29 * z, rotate: 12, fill: f, line: none });
}

/**
 * Outsized SWOT initial (S / W / O / T) used as background artwork.
 * "T" and "O" are pure geometry in the original vector, so they are rebuilt
 * exactly; "S" and "W" fall back to a heavy Montserrat glyph filling the box.
 */
function swotLetter(slide, ch, x, y, w, h, color) {
  const f = { color };
  const none = { type: 'none' };
  if (ch === 'T') {
    const barH = h * 0.195; // crossbar depth, per the original outline
    const stemW = w * 0.315;
    slide.addShape('rect', { x, y, w, h: barH, fill: f, line: none });
    slide.addShape('rect', { x: x + (w - stemW) / 2, y: y + barH, w: stemW, h: h - barH, fill: f, line: none });
  } else if (ch === 'O') {
    slide.addShape('donut', { x, y, w, h, rectRadius: Math.min(w, h) * 0.242, fill: f, line: none });
  } else {
    slide.addText(ch, {
      x: x - w * 0.5, y: y - h * 0.6, w: w * 2, h: h * 2.2,
      fontFace: HEAD, fontSize: Math.round((h * 72) / 0.7), bold: true, color,
      align: 'center', valign: 'middle', wrap: false,
    });
  }
}

/**
 * Phone / tablet mock-up standing in for the photographic device images.
 * The original artwork has a see-through screen, so `screen` lists the colour
 * bands (left to right) that show through it.  Bezels are given in inches,
 * measured off the source artwork.
 */
function deviceMockup(d) {
  const slide = d.slide;
  const none = { type: 'none' };
  const chassis = { color: d.chassis || DEVICE };
  const rim = d.rim || 0;
  if (rim) {
    // Brushed-metal edge visible around the phone bodies.
    slide.addShape('roundRect', {
      x: d.x, y: d.y, w: d.w, h: d.h, rectRadius: d.radius,
      fill: { color: '9E9E9E' }, line: { color: '242424', width: 0.75 },
    });
  }
  slide.addShape('roundRect', {
    x: d.x + rim, y: d.y + rim, w: d.w - 2 * rim, h: d.h - 2 * rim,
    rectRadius: d.radius - rim, fill: chassis, line: none,
  });

  // Screen: the last colour band fills the whole rounded screen; a band in
  // front of it is painted over the left side, squared off where they meet.
  const r = d.screenRadius;
  const sx = d.x + d.bezel.l;
  const sy = d.y + d.bezel.t;
  const sw = d.w - d.bezel.l - d.bezel.r;
  const sh = d.h - d.bezel.t - d.bezel.b;
  const bands = d.screen;
  slide.addShape('roundRect', {
    x: sx, y: sy, w: sw, h: sh, rectRadius: r,
    fill: { color: bands[bands.length - 1].color }, line: none,
  });
  if (bands.length > 1) {
    const bw = Math.min(bands[0].w, sw);
    slide.addShape('roundRect', { x: sx, y: sy, w: bw, h: sh, rectRadius: r, fill: { color: bands[0].color }, line: none });
    slide.addShape('rect', { x: sx + bw - r, y: sy, w: r, h: sh, fill: { color: bands[0].color }, line: none });
  }

  if (d.notch) {
    slide.addShape('roundRect', {
      x: d.x + d.w / 2 - d.notch.w / 2, y: sy - 0.1, w: d.notch.w, h: 0.1 + d.notch.h,
      rectRadius: 0.07, fill: chassis, line: none,
    });
  }
  if (d.homeButton) {
    slide.addShape('ellipse', {
      x: d.x + d.w / 2 - 0.16, y: d.y + d.h - d.bezel.b / 2 - 0.16, w: 0.32, h: 0.32,
      fill: { color: '1A1A1A' }, line: { color: '4A4A4A', width: 1 },
    });
    slide.addShape('ellipse', {
      x: d.x + d.w / 2 - 0.05, y: d.y + d.bezel.t / 2 - 0.025, w: 0.05, h: 0.05,
      fill: { color: '4A4A4A' }, line: none,
    });
  }
  if (d.glare) {
    // Soft diagonal reflection across the upper-right half of the device.
    slide.addShape('rtTriangle', {
      x: d.x, y: d.y, w: d.w, h: d.h, rotate: 180,
      fill: { color: 'FFFFFF', transparency: 92 }, line: none,
    });
  }
}

/* ------------------------------------------------------------------ *
 * Repeated copy
 * ------------------------------------------------------------------ */

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
  'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis ' +
  'nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. ' +
  'Duis autem vel eum iriure dolor in';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
  'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad';
const LOREM_SHORT =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
  'tincidunt ut laoreet';
const LOREM_TINY =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod';
const SWOT_SHORT =
  'Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt ut labore et ' +
  'dolore aliqua.';
const SWOT_LONG =
  'Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt ut labore et ' +
  'dolore aliqua. Ut enim ad minim veniam, quis nstrud exercitation';
const SWOT_TINY =
  'Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt ut labore et';

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Title
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 11.061, 3.75, 2.272, 3.75, TEAL);
  headline(s, 0.778, 2.008, 5.888, 2.322, [
    { text: 'Best ', options: { color: INK } },
    { text: 'Corporate', options: { color: TEAL } },
  ], { fontSize: 66 });
  eyebrow(s, 0.796, 4.408, 4.588, 'Review in ongoing process improvement', '000000', { charSpacing: 0.9 });
  eyebrow(s, 0.796, 0.414, 3.203, 'Construction Presentation', '000000', { charSpacing: 0.9 });
  socialIcons(s, 0.796, 6.69, '000000');
}

// 2 — Full-bleed teal intro
function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: TEAL };
  eyebrow(s, 0.905, 0.949, 3.946, 'Welcome To Best Corporate', WHITE95);
  headline(s, 0.905, 1.325, 5.213, 1.919, 'Transportation routes other than highways', { color: WHITE });
  paragraph(s, 0.905, 3.443, 5.367, 1.439, LOREM_LONG, WHITE95);
  paragraph(s, 0.905, 4.886, 5.367, 0.622,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna', WHITE95);
  socialIcons(s, 1.035, 5.992, CREAM);
}

// 3 — Worker roundel + CTA
function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  oval(s, 1.429, 3.583, 3.143, 3.169, CREAM);
  eyebrow(s, 6.52, 1.508, 4.386, 'Design development period', TEAL);
  headline(s, 6.52, 1.884, 5.713, 1.313, 'Final design period and preparation');
  paragraph(s, 6.525, 3.247, 5.713, 1.439, LOREM_LONG, BODY);
  button(s, 6.667, 4.933, 'CLICK HERE');
  oval(s, 1.736, 3.861, 2.528, 2.528, TEAL);
  workerIcon(s, 2.383, 4.508, 1.233, CREAM, TEAL);
}

// 4 — Headline with sub-block
function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  eyebrow(s, 0.999, 0.807, 4.071, 'Increased production capacity', TEAL);
  headline(s, 0.999, 1.183, 5.126, 1.919, 'Keep ignoring safety and God at work');
  socialIcons(s, 1.068, 3.409, '000000');
  s.addText('Your Tittle Here', {
    x: 3.034, y: 4.123, w: 3.549, h: 0.368,
    fontFace: TEXT, fontSize: 16, color: TEAL, valign: 'top',
  });
  paragraph(s, 3.034, 4.492, 3.091, 1.439,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi', BODY);
  rect(s, 7.458, 5.879, 4.651, 1.621, TEAL);
}

// 5 — Two safety call-outs on teal
function slide05(pres) {
  const s = pres.addSlide();
  s.background = { color: TEAL };
  oval(s, 6.449, 1.021, 1.932, 1.833, TEAL);
  eyebrow(s, 1.073, 1.243, 4.079, 'Maintain quality and reputation', WHITE);
  headline(s, 1.073, 1.689, 5.297, 2.524,
    'Build good relationships with other parties and customers', { color: WHITE });
  dozerIcon(s, 1.082, 4.336, 0.793, WHITE, TEAL);
  s.addText('The floor surface of the location is slippery', {
    x: 2.074, y: 4.391, w: 2.233, h: 0.622,
    fontFace: TEXT, fontSize: 12, bold: true, color: WHITE85,
    lineSpacingMultiple: 1.35, valign: 'top',
  });
  paragraph(s, 1.073, 5.179, 4.881, 0.894, LOREM_MED, WHITE85);

  rect(s, 7.38, 0.936, 4.881, 2.021, TEAL);
  oval(s, 6.69, 1.212, 1.452, 1.452, WHITE);
  excavatorIcon(s, 6.931, 1.426, 1.0, TEAL, WHITE);
  s.addText('Protects from fire, boiler, steam, and corrosives', {
    x: 8.423, y: 1.141, w: 3.018, h: 0.622,
    fontFace: TEXT, fontSize: 12, bold: true, color: WHITE85,
    lineSpacingMultiple: 1.35, valign: 'top',
  });
  paragraph(s, 8.423, 1.786, 3.495, 0.894, LOREM_SHORT, WHITE85);
}

// 6 — Teal column divider
function slide06(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 3.75, 0, 1.347, 7.5, TEAL);
  eyebrow(s, 6.667, 0.883, 2.688, 'Our Team', TEAL, { bold: true });
  headline(s, 6.667, 1.28, 5.275, 2.524, 'protects the head from impact or blows and falling objects');
  paragraph(s, 6.667, 3.865, 4.986, 1.439,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis ' +
    'nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat. ' +
    'Duis autem vel eum ', BODY);
  button(s, 6.778, 5.544, 'CLICK HERE');
}

// 7 — Team of three
function slide07(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  const team = [
    { x: 1.133, eyebrowX: 1.133, eyebrowY: 5.359, bodyX: 1.133, name: 'ISAAC PHOENIX' },
    { x: 5.106, eyebrowX: 5.145, eyebrowY: 5.356, bodyX: 5.133, name: 'HENRY STEELE' },
    { x: 9.11, eyebrowX: 9.11, eyebrowY: 5.356, bodyX: 9.098, name: 'LIAM CREST' },
  ];
  team.forEach((m) => rect(s, m.x, 1.991, 3.02, 1.44, TEAL));
  rect(s, 0, 3.236, 13.333, 4.264, TEAL);

  s.addText('Our Team', {
    x: 4.04, y: 0.624, w: 5.254, h: 0.337,
    fontFace: HEAD, fontSize: 14, bold: true, color: TEAL, align: 'center', valign: 'top',
  });
  headline(s, 4.04, 1.0, 5.254, 0.707, [
    { text: 'Our ', options: { color: INK } },
    { text: 'Constructor', options: { color: GREY25 } },
  ], { align: 'center' });

  team.forEach((m) => {
    s.addText('CONTRACTOR', {
      x: m.eyebrowX, y: m.eyebrowY, w: 2.704, h: 0.303,
      fontFace: TEXT, fontSize: 12, bold: true, charSpacing: 0.8, color: ORANGE, valign: 'top',
    });
    s.addText(m.name, {
      x: m.eyebrowX, y: 5.633, w: 2.704, h: 0.353,
      fontFace: TEXT, fontSize: 15, bold: true, charSpacing: 0.4, color: WHITE, valign: 'top',
    });
    paragraph(s, m.bodyX, 6.012, 2.851, 0.894, LOREM_TINY, CREAM);
  });
}

// 8 — Services list beside teal panel
function slide08(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 8.379, 0, 4.954, 7.5, TEAL);
  eyebrow(s, 1.058, 1.179, 2.688, 'Our Services', TEAL, { bold: true });
  headline(s, 1.001, 1.576, 3.613, 1.313, 'Safety hat or helmet');
  paragraph(s, 0.997, 2.949, 4.881, 0.622, LOREM_SHORT, BODY);
  [
    { titleY: 3.69, bodyY: 4.027 },
    { titleY: 5.089, bodyY: 5.426 },
  ].forEach((blk) => {
    s.addText('Your Text Here', {
      x: 1.001, y: blk.titleY, w: 4.876, h: 0.337,
      fontFace: TEXT, fontSize: 14, bold: true, color: INK, align: 'justify', valign: 'top',
    });
    paragraph(s, 1.001, blk.bodyY, 4.876, 0.894, LOREM_MED, BODY);
  });
}

// 9 — Four image rules + CTA
function slide09(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  [[1.23, 6.496], [4.175, 6.496], [1.23, 3.51], [4.175, 3.51]]
    .forEach(([x, y]) => rect(s, x, y, 2.591, 0.082, TEAL));
  eyebrow(s, 7.671, 1.59, 2.688, 'Our Services', TEAL, { bold: true });
  headline(s, 7.614, 1.987, 4.912, 1.919, 'protects from fire, boiler, steam, and corrosives');
  paragraph(s, 7.671, 4.009, 4.74, 1.167,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis ' +
    'nostrud exerci tation ullamcor', BODY);
  button(s, 7.772, 5.409, 'CLICK HERE');
}

// 10 — SWOT / Strength with progress bars
function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, -0.004, 0, 7.115, 6.516, TEAL);
  s.addText('Strength Company Profile Analysis SWOT', {
    x: 1.077, y: 1.119, w: 4.359, h: 1.717,
    fontFace: HEAD, fontSize: 32, bold: true, charSpacing: 0.3, color: CREAM, valign: 'top',
  });
  s.addText(SWOT_LONG, {
    x: 1.081, y: 2.883, w: 4.704, h: 0.916,
    fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
  });

  const bars = [
    { label: 'Communication', labelY: 3.946, labelW: 1.792, barY: 4.044, barW: 2.168, barH: 0.167, pct: '100', pctX: 5.186, pctY: 3.968, pctW: 0.652, pctAlign: 'right' },
    { label: 'Analysis SWOT', labelY: 4.369, labelW: 1.642, barY: 4.457, barW: 1.642, barH: 0.177, pct: '60', pctX: 5.186, pctY: 4.369, pctW: 0.581, pctAlign: 'left' },
  ];
  bars.forEach((b) => {
    s.addText(b.label, {
      x: 1.077, y: b.labelY, w: b.labelW, h: 0.32,
      fontFace: TEXT, fontSize: 13, bold: true, charSpacing: 0.4, color: CREAM, valign: 'top',
    });
    rect(s, 3.018, b.barY, b.barW, b.barH, CREAM);
    s.addText([
      { text: b.pct, options: { fontSize: 13 } },
      { text: '%', options: { fontSize: 11 } },
    ], {
      x: b.pctX, y: b.pctY, w: b.pctW, h: 0.32,
      fontFace: TEXT, bold: true, color: CREAM, align: b.pctAlign, valign: 'top',
    });
  });

  swotLetter(s, 'S', 8.298, 1.133, 2.76, 3.75, CREAM);
  s.addText(SWOT_SHORT, {
    x: 1.081, y: 4.836, w: 4.704, h: 0.635,
    fontFace: TEXT, fontSize: 12, color: WHITE, align: 'justify', lineSpacing: 20, valign: 'top',
  });
}

// 11 — SWOT / Weakness
function slide11(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 6.667, 0, 6.663, 7.5, TEAL);
  swotLetter(s, 'W', 1.038, 2.387, 4.591, 3.21, CREAM);
  line(s, 7.935, 1.485, 0.91, 0, WHITE, 3);
  s.addText('Get Started', {
    x: 8.928, y: 1.333, w: 1.45, h: 0.303,
    fontFace: HEAD, fontSize: 12, bold: true, charSpacing: 0.3, color: CREAM, valign: 'top',
  });
  s.addText('Weakness Company Profile Analysis SWOT', {
    x: 7.824, y: 1.703, w: 4.648, h: 1.717,
    fontFace: HEAD, fontSize: 32, bold: true, charSpacing: 0.3, color: CREAM, valign: 'top',
  });
  s.addText(SWOT_SHORT, {
    x: 7.828, y: 3.476, w: 4.644, h: 0.635,
    fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
  });
  s.addText(SWOT_LONG, {
    x: 7.824, y: 4.167, w: 4.704, h: 0.916,
    fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
  });
  button(s, 7.935, 5.328, 'Get Start', { w: 1.985, h: 0.538, fill: CREAM, color: TEAL, fontSize: 14 });
}

// 12 — SWOT / Opportunity
function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  swotLetter(s, 'O', 8.288, 1.828, 3.843, 3.843, CREAM);
  rect(s, 0, 0.848, 7.085, 5.803, TEAL);
  s.addShape('roundRect', {
    x: 1.275, y: 1.339, w: 0.091, h: 2.105, rectRadius: 0.045,
    fill: { color: CREAM }, line: { type: 'none' },
  });
  s.addText('Opportunity Company Profile Analysis SWOT', {
    x: 1.818, y: 1.246, w: 4.131, h: 2.255,
    fontFace: HEAD, fontSize: 32, bold: true, charSpacing: 0.3, color: CREAM, valign: 'top',
  });
  [
    { n: '01', title: 'Factory Construction', ovalY: 3.845, titleY: 3.691, bodyY: 4.091 },
    { n: '02', title: 'Profile Construction', ovalY: 5.157, titleY: 5.003, bodyY: 5.403 },
  ].forEach((step) => {
    s.addText(step.n, {
      x: 1.018, y: step.ovalY, w: 0.606, h: 0.606,
      shape: 'ellipse', fill: { type: 'none' }, line: { color: CREAM, width: 2.25 },
      fontFace: TEXT, fontSize: 12, bold: true, color: CREAM, align: 'center', valign: 'middle',
    });
    s.addText(step.title, {
      x: 1.818, y: step.titleY, w: 2.896, h: 0.362,
      fontFace: TEXT, fontSize: 14, bold: true, charSpacing: 0.4, color: CREAM,
      align: 'justify', lineSpacing: 20, valign: 'top',
    });
    s.addText(SWOT_TINY, {
      x: 1.818, y: step.bodyY, w: 3.932, h: 0.635,
      fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
    });
  });
}

// 13 — SWOT / Threat
function slide13(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, -0.004, 0, 13.333, 6.516, TEAL, 33);
  swotLetter(s, 'T', 1.626, 1.227, 2.894, 3.734, CREAM);
  s.addText('Threat Company Profile Analysis SWOT', {
    x: 7.602, y: 1.054, w: 4.429, h: 1.717,
    fontFace: HEAD, fontSize: 32, bold: true, charSpacing: 0.3, color: CREAM, valign: 'top',
  });
  [
    { title: 'Bridge Construction', titleY: 2.903, bodyY: 3.303 },
    { title: 'Mechanical Works', titleY: 4.222, bodyY: 4.622 },
  ].forEach((blk) => {
    s.addText(blk.title, {
      x: 7.602, y: blk.titleY, w: 2.896, h: 0.362,
      fontFace: TEXT, fontSize: 14, bold: true, color: CREAM, align: 'justify',
      lineSpacing: 20, valign: 'top',
    });
    s.addText(SWOT_TINY, {
      x: 7.602, y: blk.bodyY, w: 4.429, h: 0.634,
      fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
    });
  });
}

// 14 — Portfolio intro
function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 8.254, 0, 5.079, 2.369, TEAL);
  eyebrow(s, 1.145, 1.331, 2.688, 'Our Portfolio', TEAL, { bold: true });
  headline(s, 1.088, 1.728, 3.772, 1.313, [
    { text: 'Best ', options: { color: INK } },
    { text: 'Industry Our ', options: { color: TEAL } },
    { text: 'Portfolio', options: { color: INK } },
  ]);
  paragraph(s, 1.147, 3.101, 3.944, 0.894,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore', BODY);
  [4.056, 5.182].forEach((titleY) => {
    s.addText('Your Text Here', {
      x: 1.15, y: titleY, w: 3.168, h: 0.337,
      fontFace: TEXT, fontSize: 14, bold: true, color: INK, align: 'justify', valign: 'top',
    });
    paragraph(s, 1.15, titleY + 0.336, 3.94, 0.622,
      'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh', BODY);
  });
}

// 15 — Portfolio section divider
function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  s.addText('Our Best Portfolio', {
    x: 5.323, y: 0.921, w: 2.688, h: 0.337,
    fontFace: TEXT, fontSize: 14, bold: true, color: TEAL, align: 'center', valign: 'top',
  });
  headline(s, 2.913, 1.318, 7.507, 0.707, [
    { text: 'Our Great ', options: { color: NAVY } },
    { text: 'Portfolio', options: { color: TEAL } },
  ], { align: 'center' });
}

// 16 — Company portfolio
function slide16(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  eyebrow(s, 6.917, 1.051, 2.688, 'Our Company', TEAL, { bold: true });
  headline(s, 6.917, 1.388, 5.705, 1.313, [
    { text: 'Our Great ', options: { color: INK } },
    { text: 'Company Portfolio', options: { color: TEAL } },
  ]);
  [2.7, 5.65].forEach((y) => paragraph(s, 6.917, y, 5.25, 0.622,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore', BODY));
}

// 17 — Phone mock-ups
function slide17(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 4.905, 0, 8.429, 7.5, TEAL);
  // The mock-up screens are transparent: cream to the left of the teal panel,
  // teal to the right of it.
  // Bezels/notch measured off the source artwork; the small phone re-uses the
  // same drawing scaled to 55.6%.
  const phone = (x, y, k) => ({
    slide: s, x, y, w: 3.213 * k, h: 6.378 * k, radius: 0.50 * k,
    bezel: { l: 0.220 * k, r: 0.220 * k, t: 0.205 * k, b: 0.253 * k },
    screenRadius: 0.28 * k, rim: 0.030 * k, chassis: '070707',
    notch: { w: 1.543 * k, h: 0.204 * k },
  });
  deviceMockup(Object.assign(phone(2.982, 0.561, 1), {
    screen: [{ w: 4.905 - 3.202, color: CREAM }, { w: 3.213, color: TEAL }],
  }));
  deviceMockup(Object.assign(phone(0.711, 1.869, 0.556), {
    screen: [{ w: 1.787, color: CREAM }],
  }));
  eyebrow(s, 7.375, 1.136, 2.688, 'Our Mockup', CREAM, { bold: true });
  headline(s, 7.318, 1.533, 4.622, 1.313, 'Business Service Our Mockup', { color: CREAM });
  paragraph(s, 7.318, 2.989, 4.892, 1.167,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad adipiscing elit, ' +
    'sed diam nonummy nibh euismod', CREAM);
  [
    { x: 7.319, w: 2.002, big: '3.74K', small: 'Completed Projects' },
    { x: 9.473, w: 1.59, big: '97%', small: 'Good Reviews' },
  ].forEach((stat) => {
    s.addText(stat.big, {
      x: stat.x, y: 4.442, w: stat.w, h: 0.524,
      fontFace: TEXT, fontSize: 28, bold: true, color: CREAM, align: 'justify',
      lineSpacing: 29, valign: 'top',
    });
    s.addText(stat.small, {
      x: stat.x, y: 4.789, w: stat.w, h: 0.353,
      fontFace: TEXT, fontSize: 12, color: CREAM, align: 'justify', lineSpacing: 20, valign: 'top',
    });
  });
}

// 18 — Tablet mock-up
function slide18(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 9.804, 0, 3.529, 7.5, TEAL);
  deviceMockup({
    slide: s, x: 8.187, y: 0.763, w: 3.997, h: 5.975, radius: 0.24,
    bezel: { l: 0.287, r: 0.295, t: 0.687, b: 0.671 },
    screenRadius: 0.03, homeButton: true, glare: true,
    screen: [{ w: 9.804 - (8.187 + 0.287), color: CREAM }, { w: 3.997, color: TEAL }],
  });
  eyebrow(s, 1.263, 1.241, 2.688, 'Our Mockup', TEAL, { bold: true });
  headline(s, 1.206, 1.638, 6.108, 1.919, [
    { text: 'Utilize Our Services to Foster ', options: { color: INK } },
    { text: 'Your Business Growth', options: { color: TEAL } },
  ]);
  paragraph(s, 1.263, 3.616, 6.052, 0.911,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod ' +
    'tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, ',
    BODY, { lineSpacingMultiple: 1.38 });
  [
    { titleX: 1.258, title: 'Factory Business', bodyX: 1.252, bodyW: 2.914 },
    { titleX: 4.524, title: 'Growth Factory', bodyX: 4.519, bodyW: 2.796 },
  ].forEach((col) => {
    s.addText(col.title, {
      x: col.titleX, y: 4.767, w: 2.567, h: 0.353,
      fontFace: TEXT, fontSize: 15, bold: true, charSpacing: 0.4, color: INK, valign: 'top',
    });
    paragraph(s, col.bodyX, 5.143, col.bodyW, 0.911,
      'Lorem ipsum dolor sitamet consectet adipiscing elit, sed diam nonummy aseui',
      BODY, { lineSpacingMultiple: 1.38 });
  });
}

// 19 — Contact
function slide19(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  // Open-sided bracket: full top/right/bottom rules plus two short left stubs.
  const bx = 1.474, by = 1.242, bw = 4.441, bh = 4.979;
  line(s, bx, by, bw, 0, TEAL, 3);
  line(s, bx + bw, by, 0, bh, TEAL, 3);
  line(s, bx, by + bh, bw, 0, TEAL, 3);
  line(s, bx, by, 0, 0.587, TEAL, 3);
  line(s, bx, by + 4.276, 0, bh - 4.276, TEAL, 3);

  headline(s, 1.211, 1.965, 4.201, 1.919, [
    { text: 'Contact Us ', options: { color: '000000' } },
    { text: 'Immediately Here', options: { color: TEAL } },
  ], { charSpacing: 0.3 });
  s.addText('Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt ut', {
    x: 1.211, y: 3.855, w: 3.694, h: 0.635,
    fontFace: TEXT, fontSize: 12, color: BODY, align: 'justify', lineSpacing: 20, valign: 'top',
  });
  button(s, 1.33, 4.62, 'CLICK HERE');
}

// 20 — Thank you
function slide20(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  rect(s, 0, 0, 13.333, 7.5, TEAL, 20);
  rect(s, 0, 2.377, 13.336, 2.745, CREAM);
  s.addText([
    { text: 'Thank', options: { color: TEAL } },
    { text: ' ', options: {} },
    { text: 'You', options: { color: INK } },
  ], {
    x: 3.124, y: 2.656, w: 7.084, h: 1.582,
    fontFace: HEAD, fontSize: 88, bold: true, align: 'center', valign: 'top', wrap: false,
  });
  s.addText('For Your Attention', {
    x: 2.719, y: 4.238, w: 7.894, h: 0.404,
    fontFace: TEXT, fontSize: 18, bold: true, charSpacing: 0.8, color: '000000',
    align: 'center', valign: 'top',
  });
  s.addText('www.bestcorporate.com', {
    x: 2.719, y: 6.351, w: 7.894, h: 0.303,
    fontFace: TEXT, fontSize: 12, charSpacing: 3, color: WHITE, align: 'center', valign: 'top',
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'WIDE_16x9', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'WIDE_16x9';
  pres.theme = { headFontFace: HEAD, bodyFontFace: TEXT };
  pres.author = 'Best Corporate';
  pres.title = 'Best Corporate — Construction Presentation';
  BUILDERS.forEach((fn) => fn(pres));
  return pres.writeFile({
    fileName: path.join(__dirname, '065f8849-9b33-405e-b0af-8bb25a10537f_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
