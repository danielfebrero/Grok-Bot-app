/**
 * Home-Based Healthcare — 30-slide deck rebuilt with pptxgenjs.
 *
 * Each slide is one builder function below; coordinates are inches on a
 * 13.333 x 7.5 in stage. The reference deck's photo frames are redrawn as
 * `photo()` placeholders and its pictogram graphics as `icon()` marks; the
 * shared header/footer is applied to every page by `chrome()`.
 *
 *   node 17c6fb16-a3bb-46f6-b643-cae1512df098_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const W = 13.333;   // slide width  (in)
const H = 7.5;      // slide height (in)

// ------------------------------------------------------------------ palette
const C = {
  BLUE:   '0078EC',  // primary brand blue
  SKY:    '279EFF',  // lighter accent blue
  NAVY:   '0B356A',  // dark navy panels
  CYAN:   '3FF9FF',  // slider knobs, cover gradient highlight
  ULTRA:  '0116D8',  // slide 11 edge bar
  WHITE:  'FFFFFF',
  BLACK:  '000000',
  INK:    '171717',  // body copy on the infographic slides
  COAL:   '0D0D0D',
  CHAR:   '262626',
  DARK:   '3A3A3A',
  GREY40: '404040',
  GREY74: '747474',
  GREYA:  'AEAEAE',
  GREYA6: 'A6A6A6',
  SILVER: 'BFBFBF',
  SMOKE:  'F2F2F2',  // pricing card panels
  GREY80: '808080',  // muted captions
  FRAME:  'EDEDED',  // hairline of an empty photo frame
};

// --------------------------------------------------------------- copy deck
// The template ships with lorem filler; the longer runs are built by
// appending sentences so the repeated prefixes stay visible in one place.
const T = {
  BRAND: 'Home Based Healthcare',
  SITE:  'www.websitecompany.com',
  FOOT:  'Presentation Template',
  TITLE: 'Your Title Here',
  LI:    'Lorem ipsum',
  LD:    'Lorem ipsum dolor',
  L1:    'Lorem ipsum dolor sit amet',
  L1c:   'Lorem ipsum dolor sit amet, ',
  L2:    'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
};
T.L2s  = T.L2 + ' ';
T.L3   = T.L2 + ' Nullam tincidunt posuere ex in imperdiet.';
T.L3s  = T.L3 + ' ';
T.L4   = T.L3s + 'Donec accumsan commodo elit id posuere. ';
T.L5   = T.L4 + 'Nulla at mattis ex. ';
T.L6   = T.L5 + 'Donec ut nibh eget mauris condimentum finibus sit amet quis nulla. ';
T.L6d  = T.L6 + '.';
T.L7   = T.L6 + 'Curabitur tincidunt magna a augue interdum gravida. ';
T.L8   = T.L7 + 'Quisque id arcu est. ';
T.L9   = T.L8 + 'Aliquam erat volutpat. Ut venenatis elit viverra mi imperdiet sagittis. ';
T.CRAS = T.L2s + 'Cras nunc mi, sollicitudin quis ipsum a, mollis mollis tortor. '
       + 'Ut sodales felis id tellus porttitor sodales. ';

const SANS = 'Open Sans';

// ---------------------------------------------------------------- helpers

// PowerPoint's default text-frame insets (points), which the source deck uses.
const INSET = [7.2, 7.2, 3.6, 3.6];   // left, right, bottom, top

/** Text box. Source text frames are top-anchored. */
function tx(slide, text, opts) {
  slide.addText(text, Object.assign({ fontFace: SANS, valign: 'top', margin: INSET }, opts));
}

/** Straight connector. */
function rule(slide, o) {
  slide.addShape('line', {
    x: o.x, y: o.y, w: o.w, h: o.h, flipV: !!o.flipV,
    line: { color: o.color, width: o.width },
  });
}

/**
 * Stand-in for a photo frame. The reference deck ships these frames unfilled,
 * so the placeholder is only a hairline outline — enough to show where the
 * artwork belongs without painting over the page.
 */
function photo(slide, o) {
  slide.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { type: 'none' }, line: { color: C.FRAME, width: 0.5 },
  });
}

/** Stand-in for a small pictogram: a solid rounded glyph mark. */
function icon(slide, o, color) {
  const d = Math.min(o.w, o.h) * 0.62;
  slide.addShape('roundRect', {
    x: o.x + (o.w - d) / 2, y: o.y + (o.h - d) / 2, w: d, h: d,
    fill: color, rectRadius: d * 0.22,
  });
}

/** Progress slider: filled track, cyan knob and a right-aligned percentage. */
function slider(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: 3.064, h: 0.085, fill: o.color, rectRadius: 0.043,
  });
  slide.addShape('ellipse', { x: o.x + 2.728, y: o.y - 0.054, w: 0.188, h: 0.188, fill: C.CYAN });
  tx(slide, '96%', {
    x: o.x + 2.681, y: o.y - 0.292, w: 0.513, h: 0.278,
    fontSize: 10.5, bold: true, align: 'right',
  });
}

/**
 * Slide 23: horizontal pill bar. `cap` says which end carries the round caps;
 * the underlying shape is a quarter-turned round2SameRect.
 */
function pill(slide, o, color, cap) {
  slide.addShape('round2SameRect', {
    x: o.x + o.w / 2 - o.h / 2, y: o.y + o.h / 2 - o.w / 2, w: o.h, h: o.w,
    fill: color, rotate: cap === 'left' ? 270 : 90, rectRadius: o.h * 0.5,
  });
}

/** Slide 12: navy card whose bottom corners are rounded. */
function card(slide, o, color) {
  slide.addShape('round2SameRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: color, rotate: 180, rectRadius: 0.26,
  });
}

/** Slide 24: ribbon — a bar that ends in a rightward chevron point. */
function banner(slide, o, color) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: color,
    points: [
      { x: 0, y: 0 }, { x: 0.733 * o.w, y: 0 }, { x: o.w, y: 0.5 * o.h },
      { x: 0.733 * o.w, y: o.h }, { x: 0, y: o.h }, { close: true },
    ],
  });
}

/** Slide 24: folded wedge joining a number block to its ribbon. */
function wedge(slide, o, color) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: color, rotate: o.rotate,
    points: [
      { x: 0, y: o.h }, { x: 0.23 * o.w, y: 0 }, { x: o.w, y: 0 },
      { x: 0.64 * o.w, y: o.h }, { close: true },
    ],
  });
}

/** Slide 24: thin cast-shadow sliver under a ribbon. */
function drop(slide, o, color) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: color,
    points: [
      { x: o.w, y: 0.645 * o.h }, { x: 0.389 * o.w, y: 0 }, { x: 0, y: o.h },
      { x: 0.51 * o.w, y: o.h }, { close: true },
    ],
  });
}

/**
 * Slide 29: SWOT quadrant — a thick ring whose inward-facing corner is squared
 * off, built as disc + corner square + white core.
 */
const QUAD_CORNER = { tl: [0, 0], tr: [0.5, 0], bl: [0, 0.5], br: [0.5, 0.5] };
function quadRing(slide, o, corner, color) {
  const [fx, fy] = QUAD_CORNER[corner];
  slide.addShape('ellipse', { x: o.x, y: o.y, w: o.w, h: o.h, fill: color });
  slide.addShape('rect', {
    x: o.x + fx * o.w, y: o.y + fy * o.h, w: o.w * 0.5, h: o.h * 0.5, fill: color,
  });
  slide.addShape('ellipse', {
    x: o.x + o.w * 0.105, y: o.y + o.h * 0.105,
    w: o.w * 0.79, h: o.h * 0.79, fill: C.WHITE,
  });
}

/**
 * Cover background. The source fills the page with a blue-to-cyan gradient;
 * pptxgenjs has no gradient fill, so it is approximated by bilinearly
 * interpolating the four measured corner colours across a grid of tiles.
 */
const COVER_CORNERS = {                 // sampled from the reference deck
  tl: [0x01, 0x76, 0xea], tr: [0x39, 0xcd, 0xf7],
  bl: [0x21, 0xa9, 0xf3], br: [0x56, 0xf8, 0xfe],
};
function coverGradient(slide) {
  const COLS = 40, ROWS = 18;
  const k = COVER_CORNERS;
  for (let cx = 0; cx < COLS; cx++) {
    for (let cy = 0; cy < ROWS; cy++) {
      const u = (cx + 0.5) / COLS, v = (cy + 0.5) / ROWS;
      const hex = [0, 1, 2].map(i => {
        const top = k.tl[i] + (k.tr[i] - k.tl[i]) * u;
        const bot = k.bl[i] + (k.br[i] - k.bl[i]) * u;
        return Math.round(top + (bot - top) * v).toString(16).padStart(2, '0');
      }).join('').toUpperCase();
      slide.addShape('rect', {
        x: (cx * W) / COLS, y: (cy * H) / ROWS,
        w: W / COLS + 0.02, h: H / ROWS + 0.02, fill: hex,
      });
    }
  }
}

/** Running header, footer and page number, repeated on every slide. */
function chrome(slide, page, ink) {
  tx(slide, T.BRAND, { x: 0.44, y: 0.214, w: 2.187, h: 0.286, fontSize: 11, color: ink });
  tx(slide, T.SITE, { x: 10.837, y: 0.214, w: 2.187, h: 0.286, fontSize: 11, color: ink, align: 'right' });
  tx(slide, T.FOOT, { x: 0.329, y: 6.988, w: 2.597, h: 0.286, fontSize: 11, color: ink });
  tx(slide, String(page), {
    x: 12.62, y: 7.039, w: 0.272, h: 0.185, fontSize: 11, color: ink,
    align: 'center', margin: 0,
  });
}

/** Start a slide. The cover is the only page that sits on the blue gradient. */
function newSlide(pres, onCover) {
  const s = pres.addSlide();
  if (onCover) coverGradient(s);
  return s;
}

// Slide 1 — Cover — Home-Based Healthcare
function slide01(p) {
  const s = newSlide(p, true);
  tx(s, "HOME-BASED", { x:0.55, y:3.494, w:7.375, h:0.707, fontSize:36, bold:true, color:C.WHITE, fontFace:"Montserrat SemiBold" });
  tx(s, "PRESENTATION TEMPLATE", { x:0.612, y:4.913, w:3.587, h:0.337, fontSize:14, color:C.WHITE });
  tx(s, T.L3s, { x:0.624, y:5.535, w:4.204, h:0.454, fontSize:10.5, color:C.WHITE });
  tx(s, "HEALTHCARE", { x:0.55, y:3.887, w:7.375, h:1.01, fontSize:54, bold:true, color:C.WHITE, fontFace:"Montserrat SemiBold" });
  photo(s, { x:6.976, y:0.657, w:4.71, h:6.843 });
  return s;
}

// Slide 2 — About Our Service
function slide02(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:7.01, y:5.139, w:5.446, h:1.722, fill:C.NAVY });
  tx(s, "About Our Service", { x:7.006, y:2.019, w:3.98, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L2s, { x:7.006, y:3.235, w:4.282, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L9, { x:7.006, y:3.646, w:5.059, h:1.161, fontSize:10.5 });
  tx(s, T.L5, { x:8.677, y:5.732, w:3.166, h:0.808, fontSize:10.5, color:C.WHITE });
  tx(s, "100%", { x:7.175, y:5.309, w:1.378, h:0.572, fontSize:28, bold:true, color:C.WHITE });
  tx(s, T.L1, { x:8.677, y:5.309, w:3.943, h:0.303, fontSize:12, color:C.WHITE });
  photo(s, { x:0.539, y:0.562, w:3.151, h:6.299 });
  photo(s, { x:3.885, y:0.563, w:2.318, h:3.799 });
  photo(s, { x:3.885, y:4.519, w:2.318, h:2.342 });
  return s;
}

// Slide 3 — Homecare Company Profile
function slide03(p) {
  const s = newSlide(p);
  tx(s, "Homecare Company Profile", { x:0.518, y:1.807, w:5.407, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L3s, { x:0.518, y:3.051, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L5, { x:0.518, y:3.6, w:4.102, h:0.631, fontSize:10.5, align:"justify" });
  tx(s, T.L1, { x:4.848, y:4.65, w:2.277, h:0.286, fontSize:10.5 });
  tx(s, T.L1, { x:4.848, y:5.044, w:2.277, h:0.286, fontSize:10.5 });
  s.addShape("rect", { x:4.73, y:5.517, w:4.785, h:1.387, fill:C.BLUE });
  tx(s, T.L5, { x:5.252, y:5.896, w:3.884, h:0.808, fontSize:10.5, color:C.WHITE, align:"justify" });
  photo(s, { x:9.803, y:0.595, w:3.132, h:6.31 });
  photo(s, { x:0.683, y:4.563, w:3.759, h:2.342 });
  return s;
}

// Slide 4 — Vision and Mission
function slide04(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:6.897, y:5.399, w:2.651, h:1.074, fill:C.SKY });
  s.addShape("rect", { x:9.813, y:2.961, w:2.901, h:0.533, fill:C.BLUE });
  tx(s, [{ text:"Vision", options:{ breakLine:true } }, { text:"and Mission" }], { x:4.121, y:2.193, w:3.541, h:1.447, fontSize:40, fontFace:"Montserrat" });
  tx(s, T.L8, { x:4.121, y:3.64, w:5.063, h:0.985, fontSize:10.5 });
  tx(s, T.L2s, { x:7.133, y:5.708, w:2.18, h:0.454, fontSize:10.5, color:C.WHITE });
  rule(s, { x:11.865, y:3.064, w:0.583, h:0.158, color:C.WHITE, width:1 });
  rule(s, { x:11.865, y:3.22, w:0.583, h:0.127, color:C.WHITE, width:1, flipV:true });
  rule(s, { x:10.665, y:3.219, w:1.783, h:0, color:C.WHITE, width:1 });
  tx(s, T.L5, { x:9.78, y:1.509, w:2.496, h:1.161, fontSize:10.5 });
  s.addShape("rect", { x:4.278, y:5.034, w:1.307, h:0.286, fill:{ type:"none" }, line:{ color:C.BLACK, width:1 } });
  tx(s, T.LI, { x:4.348, y:5.034, w:1.168, h:0.286, fontSize:10.5, align:"center" });
  tx(s, T.LI, { x:5.666, y:6.303, w:1.36, h:0.286, fontSize:10.5, color:C.WHITE, align:"justify" });
  tx(s, T.L5, { x:4.217, y:5.488, w:2.518, h:1.161, fontSize:10.5 });
  photo(s, { x:0.737, y:0.623, w:3.18, h:6.242 });
  photo(s, { x:9.78, y:3.741, w:2.933, h:3.124 });
  return s;
}

// Slide 5 — Our Core Values
function slide05(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:9.261, y:2.817, w:3.788, h:1.575, fill:C.BLUE });
  tx(s, "Our Core Values", { x:0.54, y:2.565, w:4.802, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L2, { x:0.54, y:3.302, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L9, { x:0.54, y:3.726, w:5.066, h:1.161, fontSize:10.5 });
  tx(s, T.L4, { x:9.568, y:3.176, w:3.225, h:0.808, fontSize:10.5, color:C.WHITE });
  s.addShape("rect", { x:0.669, y:5.107, w:3.658, h:1.575, fill:C.SKY });
  tx(s, T.L3s, { x:0.907, y:5.772, w:3.182, h:0.631, fontSize:10.5, color:C.WHITE });
  s.addShape("rect", { x:1.001, y:5.393, w:1.307, h:0.286, fill:C.NAVY });
  tx(s, T.LI, { x:1.07, y:5.397, w:1.168, h:0.286, fontSize:10.5, color:C.WHITE, align:"center" });
  s.addShape("rect", { x:4.729, y:5.698, w:4.162, h:0.985, fill:C.NAVY });
  tx(s, T.L3s, { x:5.606, y:5.895, w:2.957, h:0.631, fontSize:10.5, color:C.WHITE });
  s.addShape("rect", { x:4.727, y:5.251, w:2.428, h:0.286, fill:C.BLUE });
  tx(s, T.L1, { x:4.861, y:5.254, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  icon(s, { x:4.935, y:5.934, w:0.469, h:0.469 }, C.WHITE);
  photo(s, { x:9.174, y:4.717, w:3.788, h:2.18 });
  photo(s, { x:6.239, y:0.565, w:2.651, h:4.152 });
  return s;
}

// Slide 6 — What We Offer
function slide06(p) {
  const s = newSlide(p);
  tx(s, "What We Offer", { x:0.637, y:2.455, w:4.212, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L3s, { x:0.539, y:3.406, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L6, { x:0.539, y:4.099, w:4.508, h:0.808, fontSize:10.5 });
  tx(s, T.L2s, { x:3.349, y:5.274, w:1.862, h:0.631, fontSize:10.5 });
  tx(s, T.L1, { x:5.8, y:3.406, w:2.525, h:0.286, fontSize:10.5, bold:true });
  tx(s, T.L5, { x:5.8, y:3.747, w:2.679, h:0.985, fontSize:10.5 });
  s.addShape("rect", { x:5.683, y:4.977, w:2.941, h:1.919, fill:C.BLUE });
  s.addShape("rect", { x:5.649, y:2.577, w:2.941, h:0.533, fill:C.SKY });
  tx(s, T.L5, { x:6.05, y:5.444, w:2.578, h:0.985, fontSize:10.5, color:C.WHITE });
  rule(s, { x:7.742, y:2.68, w:0.583, h:0.158, color:C.WHITE, width:1 });
  rule(s, { x:7.742, y:2.836, w:0.583, h:0.127, color:C.WHITE, width:1, flipV:true });
  rule(s, { x:6.542, y:2.835, w:1.783, h:0, color:C.WHITE, width:1 });
  s.addShape("rect", { x:0.706, y:5.375, w:2.428, h:0.286, fill:C.NAVY });
  tx(s, T.L1, { x:0.84, y:5.378, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  s.addShape("rect", { x:0.715, y:5.905, w:2.428, h:0.286, fill:C.NAVY });
  tx(s, T.L1, { x:0.849, y:5.908, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  photo(s, { x:8.83, y:0.605, w:3.896, h:6.291 });
  return s;
}

// Slide 7 — How It Works
function slide07(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:0.424, y:3.755, w:0.776, h:0.816, fill:C.NAVY });
  tx(s, T.L3s, { x:0.343, y:2.965, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, "How It Works", { x:0.343, y:1.74, w:4.801, h:0.707, fontSize:36, fontFace:"Montserrat" });
  icon(s, { x:0.646, y:3.984, w:0.358, h:0.358 }, C.WHITE);
  s.addShape("rect", { x:0.424, y:5.218, w:0.776, h:0.816, fill:C.NAVY });
  tx(s, T.L5, { x:1.331, y:5.943, w:4.801, h:0.631, fontSize:10.5 });
  slider(s, { x:1.424, y:5.671, color:C.BLUE });
  icon(s, { x:0.624, y:5.434, w:0.403, h:0.403 }, C.WHITE);
  s.addShape("rect", { x:6.667, y:5.379, w:5.404, h:1.08, fill:C.BLUE });
  tx(s, T.L3s, { x:7.068, y:5.636, w:4.694, h:0.454, fontSize:10.5, color:C.WHITE, align:"center" });
  s.addShape("rect", { x:1.407, y:5.225, w:1.307, h:0.286, fill:C.SKY, line:{ color:C.WHITE, width:1 } });
  tx(s, T.LI, { x:1.476, y:5.228, w:1.168, h:0.286, fontSize:10.5, color:C.WHITE, align:"center" });
  tx(s, T.L5, { x:1.331, y:4.449, w:4.801, h:0.631, fontSize:10.5 });
  slider(s, { x:1.424, y:4.178, color:C.BLUE });
  s.addShape("rect", { x:1.407, y:3.731, w:1.307, h:0.286, fill:C.SKY, line:{ color:C.WHITE, width:1 } });
  tx(s, T.LI, { x:1.476, y:3.735, w:1.168, h:0.286, fontSize:10.5, color:C.WHITE, align:"center" });
  photo(s, { x:6.667, y:0.919, w:5.404, h:4.352 });
  return s;
}

// Slide 8 — Antolia Carla — CEO
function slide08(p) {
  const s = newSlide(p);
  tx(s, T.L3s, { x:7.155, y:2.613, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L6d, { x:7.155, y:3.145, w:4.45, h:0.808, fontSize:10.5 });
  tx(s, "CEO-COMPANY", { x:7.155, y:2.102, w:4.212, h:0.337, fontSize:14 });
  tx(s, "Antolia Carla", { x:7.155, y:1.395, w:5.249, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L5, { x:7.155, y:5.074, w:4.623, h:0.555, fontSize:9 });
  s.addShape("rect", { x:7.249, y:5.826, w:4.529, h:0.808, fill:C.BLUE });
  tx(s, T.L5, { x:7.346, y:5.959, w:4.301, h:0.555, fontSize:9, bold:true, color:C.WHITE });
  slider(s, { x:7.249, y:4.66, color:C.BLUE });
  tx(s, T.LI, { x:7.301, y:4.217, w:1.168, h:0.286, fontSize:10.5, align:"center" });
  photo(s, { x:1.639, y:0.854, w:4.484, h:5.791 });
  return s;
}

// Slide 9 — Meet Our Team
function slide09(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:7.363, y:5.504, w:0.617, h:0.631, fill:C.BLUE });
  s.addShape("rect", { x:0.236, y:5.222, w:3.158, h:1.179, fill:C.NAVY });
  s.addShape("rect", { x:7.363, y:4.482, w:0.617, h:0.631, fill:C.BLUE });
  tx(s, "Meet Our Team", { x:7.244, y:2.709, w:5.366, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L5, { x:7.244, y:3.459, w:3.989, h:0.631, fontSize:10.5, fontFace:"Lato" });
  tx(s, T.L5, { x:8.04, y:4.466, w:3.698, h:0.631, fontSize:10.5, fontFace:"Lato", align:"justify" });
  tx(s, T.L5, { x:8.04, y:5.531, w:3.698, h:0.631, fontSize:10.5, fontFace:"Lato", align:"justify" });
  icon(s, { x:7.452, y:4.578, w:0.438, h:0.438 }, C.WHITE);
  icon(s, { x:7.496, y:5.607, w:0.438, h:0.438 }, C.WHITE);
  tx(s, T.L1, { x:0.723, y:5.472, w:2.277, h:0.286, fontSize:10.5, color:C.WHITE, fontFace:"Lato" });
  tx(s, T.L1, { x:0.723, y:5.866, w:2.277, h:0.286, fontSize:10.5, color:C.WHITE, fontFace:"Lato" });
  photo(s, { x:3.613, y:2.551, w:2.943, h:4.332 });
  photo(s, { x:0.236, y:0.595, w:3.132, h:3.986 });
  return s;
}

// Slide 10 — Care Service Types
function slide10(p) {
  const s = newSlide(p);
  tx(s, T.L3s, { x:0.849, y:5.697, w:3.047, h:0.631, fontSize:10.5, fontFace:"Lato" });
  tx(s, T.L3s, { x:3.032, y:3.696, w:2.749, h:0.631, fontSize:10.5, bold:true, fontFace:"Lato" });
  tx(s, "Care Service Types", { x:3.032, y:1.722, w:3.41, h:1.313, fontSize:36, fontFace:"Montserrat" });
  s.addShape("rect", { x:10.3, y:3.756, w:2.701, h:0.631, fill:C.BLUE });
  rule(s, { x:11.929, y:4.094, w:0.775, h:0, color:C.WHITE, width:1 });
  rule(s, { x:12.154, y:3.936, w:0.55, h:0.161, color:C.WHITE, width:1 });
  rule(s, { x:12.154, y:4.096, w:0.55, h:0.148, color:C.WHITE, width:1, flipV:true });
  rule(s, { x:10.597, y:4.086, w:0.775, h:0, color:C.WHITE, width:1 });
  rule(s, { x:10.597, y:4.083, w:0.55, h:0.161, color:C.WHITE, width:1 });
  rule(s, { x:10.597, y:3.936, w:0.55, h:0.148, color:C.WHITE, width:1, flipV:true });
  tx(s, T.L3s, { x:3.032, y:2.989, w:3.147, h:0.631, fontSize:10.5, fontFace:"Lato" });
  s.addShape("rect", { x:4.904, y:4.772, w:5.501, h:2.122, fill:C.NAVY });
  tx(s, T.L6, { x:5.448, y:5.375, w:4.597, h:0.808, fontSize:10.5, color:C.WHITE, fontFace:"Lato" });
  s.addShape("rect", { x:0.932, y:5.336, w:2.428, h:0.286, fill:C.SKY });
  tx(s, T.L1, { x:1.066, y:5.34, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, fontFace:"Lato", align:"center" });
  photo(s, { x:6.667, y:1.397, w:3.171, h:3.136 });
  photo(s, { x:0.388, y:2.22, w:2.332, h:2.313 });
  photo(s, { x:10.638, y:4.533, w:2.398, h:2.361 });
  return s;
}

// Slide 11 — Patient-Centered Approach
function slide11(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:0.908, y:4.732, w:2.701, h:2.175, fill:C.BLUE });
  tx(s, T.L1, { x:1.073, y:5.633, w:2.647, h:0.278, fontSize:10.5, bold:true, color:C.WHITE });
  tx(s, T.L3s, { x:1.073, y:5.911, w:2.613, h:0.631, fontSize:10.5, color:C.WHITE });
  s.addShape("rect", { x:3.779, y:4.732, w:2.701, h:2.175, fill:C.BLUE });
  tx(s, T.L1, { x:3.944, y:5.633, w:2.647, h:0.278, fontSize:10.5, bold:true, color:C.WHITE });
  tx(s, T.L3s, { x:3.944, y:5.911, w:2.613, h:0.631, fontSize:10.5, color:C.WHITE });
  tx(s, "Patient-Centered Approach", { x:0.678, y:1.717, w:5.335, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L2, { x:0.678, y:2.984, w:3.368, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L6d, { x:0.678, y:3.489, w:4.493, h:0.808, fontSize:10.5 });
  icon(s, { x:1.102, y:4.867, w:0.631, h:0.631 }, C.WHITE);
  icon(s, { x:4.153, y:4.867, w:0.631, h:0.631 }, C.WHITE);
  s.addShape("rect", { x:12.321, y:4.531, w:0.624, h:2.356, fill:C.ULTRA });
  photo(s, { x:6.667, y:2.95, w:2.978, h:3.956 });
  photo(s, { x:9.815, y:4.55, w:3.143, h:2.356 });
  photo(s, { x:9.815, y:0.594, w:3.143, h:3.794 });
  return s;
}

// Slide 12 — Expert Medical Staff
function slide12(p) {
  const s = newSlide(p);
  card(s, { x:0.415, y:4.668, w:3.127, h:1.717 }, C.NAVY);
  tx(s, "Expert Medical Staff", { x:4.386, y:1.763, w:5.417, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L3s, { x:4.386, y:2.582, w:3.814, h:0.631, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:0.836, y:5.034, w:2.429, h:0.808, fontSize:10.5, bold:true, color:C.WHITE });
  tx(s, T.L6, { x:7.023, y:5.244, w:3.221, h:1.161, fontSize:10.5 });
  tx(s, T.L7, { x:4.386, y:3.261, w:5.102, h:0.985, fontSize:10.5 });
  photo(s, { x:10.6, y:2.715, w:2.331, h:4.214 });
  photo(s, { x:3.725, y:4.543, w:2.942, h:2.386 });
  photo(s, { x:0.403, y:0.629, w:3.127, h:3.914 });
  return s;
}

// Slide 13 — Personalized Care Plans
function slide13(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:8.562, y:3.472, w:4.331, h:1.361, fill:C.NAVY });
  s.addShape("rect", { x:8.562, y:5.054, w:4.331, h:1.361, fill:C.NAVY });
  s.addShape("rect", { x:8.562, y:1.864, w:4.331, h:1.361, fill:C.NAVY });
  tx(s, "Personalized Care Plans", { x:0.44, y:2.405, w:4.212, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L9, { x:0.44, y:3.864, w:4.331, h:1.338, fontSize:10.5 });
  tx(s, T.L3s, { x:9.485, y:2.405, w:3.177, h:0.631, fontSize:10.5, color:C.WHITE });
  tx(s, T.L5, { x:0.44, y:5.77, w:4.331, h:0.555, fontSize:9 });
  slider(s, { x:0.534, y:5.555, color:C.BLUE });
  tx(s, T.TITLE, { x:0.44, y:5.254, w:2.143, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L1, { x:9.469, y:2.135, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:9.485, y:3.975, w:3.177, h:0.631, fontSize:10.5, color:C.WHITE });
  tx(s, T.L1, { x:9.469, y:3.705, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:9.485, y:5.57, w:3.177, h:0.631, fontSize:10.5, color:C.WHITE });
  tx(s, T.L1, { x:9.469, y:5.3, w:2.16, h:0.278, fontSize:10.5, color:C.WHITE, align:"center" });
  icon(s, { x:8.944, y:5.445, w:0.381, h:0.381 }, C.WHITE);
  icon(s, { x:8.946, y:3.882, w:0.381, h:0.381 }, C.WHITE);
  icon(s, { x:8.958, y:2.215, w:0.381, h:0.381 }, C.WHITE);
  photo(s, { x:5.121, y:1.144, w:3.285, h:5.212 });
  return s;
}

// Slide 14 — Daily Health Monitoring
function slide14(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:8.586, y:5.243, w:0.7, h:0.739, fill:C.BLUE });
  s.addShape("rect", { x:8.586, y:6.167, w:0.7, h:0.739, fill:C.BLUE });
  tx(s, "Daily Health Monitoring", { x:0.373, y:1.866, w:4.743, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L5, { x:2.926, y:5.649, w:2.789, h:0.985, fontSize:10.5 });
  tx(s, T.L9, { x:0.373, y:3.178, w:5.948, h:0.985, fontSize:10.5 });
  s.addShape("rect", { x:8.586, y:4.349, w:0.7, h:0.739, fill:C.BLUE });
  tx(s, T.TITLE, { x:9.491, y:4.32, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:9.491, y:4.579, w:3.47, h:0.631, fontSize:10.5 });
  tx(s, T.TITLE, { x:9.491, y:5.295, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:9.491, y:5.554, w:3.47, h:0.631, fontSize:10.5 });
  tx(s, T.TITLE, { x:9.491, y:6.095, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:9.491, y:6.354, w:3.47, h:0.631, fontSize:10.5 });
  s.addShape("rect", { x:3, y:5.24, w:1.307, h:0.286, fill:C.BLUE });
  tx(s, T.LI, { x:3.07, y:5.243, w:1.168, h:0.286, fontSize:10.5, color:C.WHITE, align:"center" });
  icon(s, { x:8.75, y:6.321, w:0.381, h:0.381 }, C.WHITE);
  icon(s, { x:8.752, y:5.432, w:0.381, h:0.381 }, C.WHITE);
  icon(s, { x:8.763, y:4.503, w:0.381, h:0.381 }, C.WHITE);
  photo(s, { x:8.524, y:0.589, w:4.437, h:3.605 });
  photo(s, { x:5.885, y:4.551, w:2.333, h:2.353 });
  photo(s, { x:0.396, y:4.551, w:2.333, h:2.353 });
  return s;
}

// Slide 15 — Elderly Support Services
function slide15(p) {
  const s = newSlide(p);
  tx(s, "Elderly Support Services", { x:2.342, y:2.293, w:4.212, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6d, { x:2.383, y:5.733, w:4.019, h:0.985, fontSize:10.5 });
  tx(s, T.L5, { x:2.383, y:3.672, w:3.429, h:0.707, fontSize:9, bold:true });
  s.addShape("rect", { x:6.695, y:3.131, w:2.98, h:1.23, fill:C.NAVY });
  tx(s, T.TITLE, { x:6.857, y:3.295, w:1.703, h:0.278, fontSize:10.5, bold:true, color:C.WHITE });
  tx(s, T.L3s, { x:6.857, y:3.554, w:2.656, h:0.631, fontSize:10.5, color:C.WHITE });
  slider(s, { x:2.515, y:4.72, color:C.NAVY });
  tx(s, T.TITLE, { x:2.422, y:4.419, w:2.143, h:0.278, fontSize:10.5, bold:true });
  slider(s, { x:2.515, y:5.355, color:C.NAVY });
  tx(s, T.TITLE, { x:2.422, y:5.054, w:2.143, h:0.278, fontSize:10.5, bold:true });
  photo(s, { x:9.816, y:0.643, w:3.162, h:6.243 });
  photo(s, { x:6.667, y:4.521, w:3.009, h:2.364 });
  photo(s, { x:0.356, y:3.75, w:1.758, h:1.614 });
  return s;
}

// Slide 16 — Post-Surgery Care
function slide16(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:3.885, y:5.928, w:3.361, h:0.96, fill:C.NAVY });
  tx(s, "Post-Surgery Care", { x:3.791, y:1.805, w:4.212, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L2s, { x:3.811, y:3.261, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L6, { x:3.791, y:3.754, w:3.832, h:0.985, fontSize:10.5 });
  tx(s, T.L3s, { x:4.074, y:6.081, w:3.193, h:0.631, fontSize:10.5, color:C.WHITE });
  tx(s, T.TITLE, { x:7.556, y:4.904, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L2s, { x:7.556, y:5.198, w:2.365, h:0.454, fontSize:10.5 });
  tx(s, T.TITLE, { x:7.556, y:5.964, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L2s, { x:7.556, y:6.258, w:2.365, h:0.454, fontSize:10.5 });
  slider(s, { x:3.904, y:5.378, color:C.BLUE });
  tx(s, T.TITLE, { x:3.811, y:5.077, w:2.143, h:0.278, fontSize:10.5, bold:true });
  photo(s, { x:0.375, y:1.682, w:3.139, h:5.204 });
  photo(s, { x:9.82, y:4.536, w:3.139, h:2.349 });
  photo(s, { x:10.619, y:0.615, w:2.331, h:3.749 });
  return s;
}

// Slide 17 — Emergency Care Access
function slide17(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:10, y:4.227, w:3.021, h:2.717, fill:C.NAVY });
  tx(s, "Emergency Care Access", { x:0.769, y:1.843, w:4.231, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L2s, { x:0.788, y:3.093, w:4.883, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L9, { x:0.769, y:3.444, w:4.231, h:1.338, fontSize:10.5 });
  tx(s, T.L3s, { x:10.311, y:5.141, w:2.305, h:0.808, fontSize:10.5, color:C.WHITE });
  s.addShape("rect", { x:6.677, y:3.444, w:3.186, h:0.631, fill:C.BLUE });
  rule(s, { x:8.562, y:3.782, w:0.775, h:0, color:C.WHITE, width:1 });
  rule(s, { x:8.787, y:3.624, w:0.55, h:0.161, color:C.WHITE, width:1 });
  rule(s, { x:8.787, y:3.783, w:0.55, h:0.148, color:C.WHITE, width:1, flipV:true });
  rule(s, { x:7.23, y:3.774, w:0.775, h:0, color:C.WHITE, width:1 });
  rule(s, { x:7.23, y:3.771, w:0.55, h:0.161, color:C.WHITE, width:1 });
  rule(s, { x:7.23, y:3.624, w:0.55, h:0.148, color:C.WHITE, width:1, flipV:true });
  s.addShape("rect", { x:0.801, y:5.788, w:0.7, h:0.739, fill:C.NAVY });
  s.addShape("rect", { x:0.801, y:4.893, w:0.7, h:0.739, fill:C.NAVY });
  tx(s, T.TITLE, { x:1.807, y:4.865, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:1.807, y:5.124, w:3.47, h:0.631, fontSize:10.5 });
  tx(s, T.TITLE, { x:1.807, y:5.781, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:1.807, y:6.04, w:3.47, h:0.631, fontSize:10.5 });
  icon(s, { x:0.965, y:5.962, w:0.381, h:0.381 }, C.WHITE);
  icon(s, { x:0.966, y:5.073, w:0.381, h:0.381 }, C.WHITE);
  photo(s, { x:6.667, y:4.227, w:3.197, h:2.718 });
  photo(s, { x:10, y:0.635, w:2.969, h:3.592 });
  return s;
}

// Slide 18 — Mental Health Support
function slide18(p) {
  const s = newSlide(p);
  tx(s, "Mental Health Support", { x:0.275, y:3.094, w:4.212, h:1.313, fontSize:36, fontFace:"Montserrat" });
  s.addShape("rect", { x:4.311, y:3.6, w:2.356, h:0.748, fill:C.DARK });
  icon(s, { x:4.815, y:3.841, w:0.227, h:0.212 }, C.WHITE);
  icon(s, { x:5.154, y:3.841, w:0.228, h:0.212 }, C.WHITE);
  icon(s, { x:5.488, y:3.856, w:0.236, h:0.205 }, C.WHITE);
  icon(s, { x:5.826, y:3.826, w:0.236, h:0.235 }, C.WHITE);
  tx(s, "Industrial Design", { x:7.636, y:4.799, w:1.19, h:0.168, fontSize:10, bold:true, align:"right", wrap:false, margin:0 });
  icon(s, { x:7.324, y:4.795, w:0.235, h:0.175 }, C.BLACK);
  tx(s, T.L2s, { x:7.24, y:5.027, w:2.235, h:0.454, fontSize:10.5 });
  s.addShape("rect", { x:6.829, y:5.715, w:2.671, h:1.192, fill:C.BLUE });
  tx(s, "Industrial Design", { x:7.602, y:5.999, w:1.19, h:0.168, fontSize:10, bold:true, color:C.WHITE, align:"right", wrap:false, margin:0 });
  icon(s, { x:7.29, y:5.995, w:0.235, h:0.175 }, C.WHITE);
  tx(s, T.L2s, { x:7.206, y:6.228, w:2.235, h:0.454, fontSize:10.5, color:C.WHITE });
  tx(s, T.L1, { x:0.269, y:4.502, w:2.945, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L6, { x:0.283, y:4.89, w:3.371, h:1.161, fontSize:10.5 });
  photo(s, { x:9.663, y:1.402, w:3.155, h:5.505 });
  photo(s, { x:4.31, y:4.529, w:2.357, h:2.378 });
  return s;
}

// Slide 19 — Technology and Tools
function slide19(p) {
  const s = newSlide(p);
  photo(s, { x:7, y:0, w:6.333, h:7.5 });
  s.addShape("rect", { x:3.306, y:4.707, w:3.36, h:2.204, fill:C.NAVY });
  tx(s, T.L2s, { x:0.425, y:3.279, w:5.258, h:0.278, fontSize:10.5, bold:true });
  tx(s, [{ text:"Technology", options:{ breakLine:true } }, { text:"and Tools" }], { x:0.425, y:1.884, w:5.589, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L5, { x:3.256, y:3.75, w:3.328, h:0.808, fontSize:10.5 });
  tx(s, "87$", { x:3.637, y:5.95, w:1.276, h:0.404, fontSize:24, bold:true, color:C.WHITE, margin:0 });
  tx(s, T.L3, { x:4.471, y:5.939, w:2.099, h:0.707, fontSize:9, color:C.WHITE });
  tx(s, "100%", { x:3.5, y:4.971, w:1.276, h:0.404, fontSize:24, bold:true, color:C.WHITE, margin:0 });
  tx(s, T.L3, { x:4.484, y:4.96, w:2.099, h:0.707, fontSize:9, color:C.WHITE });
  photo(s, { x:0.571, y:3.75, w:2.443, h:2.443 });
  return s;
}

// Slide 20 — Home Visit Process
function slide20(p) {
  const s = newSlide(p);
  tx(s, T.L2s, { x:0.44, y:3.05, w:4.783, h:0.278, fontSize:10.5, bold:true });
  tx(s, [{ text:"Home", options:{ breakLine:true } }, { text:"Visit Process" }], { x:0.44, y:1.666, w:4.639, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:0.44, y:3.415, w:4.639, h:0.808, fontSize:10.5 });
  s.addShape("rect", { x:3.673, y:4.608, w:5.148, h:2.292, fill:C.NAVY });
  tx(s, T.L3s, { x:6.559, y:2.533, w:2.902, h:0.631, fontSize:10.5 });
  tx(s, T.L3s, { x:6.559, y:3.471, w:2.902, h:0.631, fontSize:10.5 });
  tx(s, "97/100", { x:3.958, y:4.999, w:2.246, h:0.707, fontSize:36, bold:true, color:C.WHITE });
  tx(s, T.L5, { x:3.958, y:5.769, w:4.406, h:0.631, fontSize:10.5, color:C.WHITE });
  photo(s, { x:9.804, y:0.601, w:3.168, h:3.808 });
  photo(s, { x:9.04, y:4.609, w:3.933, h:2.291 });
  photo(s, { x:0.753, y:4.609, w:2.626, h:2.291 });
  return s;
}

// Slide 21 — Safety and Hygiene
function slide21(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:3.757, y:5.668, w:4.327, h:1.211, fill:C.NAVY });
  tx(s, T.L3s, { x:3.872, y:4.151, w:4.212, h:0.454, fontSize:10.5, bold:true });
  tx(s, [{ text:"Safety", options:{ breakLine:true } }, { text:"and Hygiene" }], { x:3.872, y:2.792, w:4.212, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.TITLE, { x:3.872, y:4.795, w:1.703, h:0.278, fontSize:10.5, bold:true });
  tx(s, T.L3s, { x:3.872, y:5.054, w:3.813, h:0.454, fontSize:10.5 });
  tx(s, T.TITLE, { x:3.872, y:5.839, w:1.703, h:0.278, fontSize:10.5, bold:true, color:C.WHITE });
  tx(s, T.L3s, { x:3.872, y:6.097, w:3.813, h:0.454, fontSize:10.5, color:C.WHITE });
  photo(s, { x:0.4, y:0.621, w:3.124, h:6.259 });
  photo(s, { x:8.317, y:4.54, w:4.616, h:2.34 });
  photo(s, { x:9.81, y:0.621, w:3.124, h:3.698 });
  return s;
}

// Slide 22 — Price and List
function slide22(p) {
  const s = newSlide(p);
  s.addShape("rect", { x:5.092, y:1.728, w:2.244, h:4.578, fill:C.SMOKE });
  rule(s, { x:5.092, y:2.652, w:2.233, h:0, color:C.BLUE, width:4.5 });
  tx(s, "Most Recent", { x:5.591, y:2.29, w:1.247, h:0.252, fontSize:9, italic:true, color:C.GREY80, align:"center" });
  tx(s, "Free entrance", { x:5.418, y:3.739, w:1.593, h:0.303, fontSize:12, align:"center" });
  tx(s, "Free membership", { x:5.314, y:4.1, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "1x speech entrance", { x:5.314, y:4.46, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "Online tutorial guide", { x:5.262, y:4.821, w:1.913, h:0.303, fontSize:12, align:"center" });
  tx(s, "Frivent dashboard", { x:5.314, y:5.181, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "19", { x:5.716, y:2.744, w:1.043, h:0.841, fontSize:44, align:"center" });
  tx(s, "$", { x:5.587, y:2.868, w:0.317, h:0.37, fontSize:16, bold:true });
  tx(s, "/month", { x:5.851, y:3.419, w:0.774, h:0.269, fontSize:10, italic:true, color:C.GREY80, align:"center" });
  s.addShape("roundRect", { x:5.716, y:5.698, w:1.234, h:0.296, fill:C.WHITE, rectRadius:0.148 });
  tx(s, "Text Here", { x:5.716, y:5.698, w:1.234, h:0.296, fontSize:12, align:"center", valign:"middle" });
  tx(s, "Standart", { x:5.284, y:1.92, w:1.861, h:0.404, align:"center" });
  s.addShape("rect", { x:7.783, y:1.728, w:2.244, h:4.578, fill:C.SMOKE });
  rule(s, { x:7.783, y:2.652, w:2.233, h:0, color:C.NAVY, width:4.5 });
  tx(s, "For Single Speaker", { x:8.008, y:2.29, w:1.783, h:0.252, fontSize:9, italic:true, color:C.GREY80, align:"center" });
  tx(s, "Free entrance", { x:8.108, y:3.739, w:1.593, h:0.303, fontSize:12, align:"center" });
  tx(s, "Free membership", { x:8.005, y:4.1, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "1x speech entrance", { x:8.005, y:4.46, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "Online tutorial guide", { x:7.953, y:4.821, w:1.913, h:0.303, fontSize:12, align:"center" });
  tx(s, "Frivent dashboard", { x:8.005, y:5.181, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "39", { x:8.407, y:2.744, w:1.043, h:0.841, fontSize:44, align:"center" });
  tx(s, "$", { x:8.278, y:2.868, w:0.317, h:0.37, fontSize:16, bold:true });
  tx(s, "/month", { x:8.542, y:3.419, w:0.774, h:0.269, fontSize:10, italic:true, color:C.GREY80, align:"center" });
  s.addShape("roundRect", { x:8.407, y:5.698, w:1.234, h:0.296, fill:C.WHITE, rectRadius:0.148 });
  tx(s, "Text Here", { x:8.407, y:5.698, w:1.234, h:0.296, fontSize:12, align:"center", valign:"middle" });
  tx(s, "Business", { x:7.974, y:1.92, w:1.861, h:0.404, align:"center" });
  s.addShape("rect", { x:10.473, y:1.728, w:2.244, h:4.578, fill:C.SMOKE });
  rule(s, { x:10.473, y:2.652, w:2.233, h:0, color:C.SKY, width:4.5 });
  tx(s, "Most Popular", { x:10.972, y:2.29, w:1.247, h:0.252, fontSize:9, italic:true, color:C.GREY80, align:"center" });
  tx(s, "Free entrance", { x:10.799, y:3.739, w:1.593, h:0.303, fontSize:12, align:"center" });
  tx(s, "Free membership", { x:10.696, y:4.1, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "1x speech entrance", { x:10.696, y:4.46, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "Online tutorial guide", { x:10.644, y:4.821, w:1.913, h:0.303, fontSize:12, align:"center" });
  tx(s, "Frivent dashboard", { x:10.696, y:5.181, w:1.8, h:0.303, fontSize:12, align:"center" });
  tx(s, "59", { x:11.098, y:2.744, w:1.043, h:0.841, fontSize:44, align:"center" });
  tx(s, "$", { x:10.969, y:2.868, w:0.317, h:0.37, fontSize:16, bold:true });
  tx(s, "/month", { x:11.233, y:3.419, w:0.774, h:0.269, fontSize:10, italic:true, color:C.GREY80, align:"center" });
  s.addShape("roundRect", { x:11.098, y:5.698, w:1.234, h:0.296, fill:C.WHITE, rectRadius:0.148 });
  tx(s, "Text Here", { x:11.098, y:5.698, w:1.234, h:0.296, fontSize:12, align:"center", valign:"middle" });
  tx(s, "Premium", { x:10.665, y:1.92, w:1.861, h:0.404, align:"center" });
  tx(s, "PRICE AND LIST", { x:0.567, y:3.753, w:4.273, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.CRAS, { x:0.567, y:4.568, w:3.99, h:0.606, fontSize:10, color:C.INK });
  return s;
}

// Slide 23 — Infographic — pill bars
function slide23(p) {
  const s = newSlide(p);
  pill(s, { x:2.479, y:2.608, w:3.101, h:0.427 }, C.SKY, "left");
  pill(s, { x:2.211, y:3.268, w:3.369, h:0.427 }, C.DARK, "left");
  pill(s, { x:1.707, y:3.929, w:3.873, h:0.427 }, C.SKY, "left");
  pill(s, { x:2.463, y:4.59, w:3.116, h:0.427 }, C.DARK, "left");
  pill(s, { x:1.339, y:5.251, w:4.24, h:0.427 }, C.SKY, "left");
  pill(s, { x:5.579, y:2.608, w:1.338, h:0.427 }, C.DARK, "right");
  pill(s, { x:5.579, y:3.269, w:1.338, h:0.427 }, C.SKY, "right");
  pill(s, { x:5.579, y:3.93, w:1.338, h:0.427 }, C.DARK, "right");
  pill(s, { x:5.579, y:4.59, w:1.338, h:0.427 }, C.SKY, "right");
  pill(s, { x:5.579, y:5.251, w:1.338, h:0.427 }, C.DARK, "right");
  tx(s, "LOREM IPSUM", { x:5.772, y:2.722, w:0.961, h:0.236, fontSize:8, color:C.WHITE, align:"right", wrap:false });
  tx(s, "LOREM IPSUM", { x:5.772, y:3.403, w:0.961, h:0.236, fontSize:8, color:C.WHITE, align:"right", wrap:false });
  tx(s, "LOREM IPSUM", { x:5.772, y:4.063, w:0.961, h:0.236, fontSize:8, color:C.WHITE, align:"right", wrap:false });
  tx(s, "LOREM IPSUM", { x:5.772, y:4.724, w:0.961, h:0.236, fontSize:8, color:C.WHITE, align:"right", wrap:false });
  tx(s, "LOREM IPSUM", { x:5.772, y:5.385, w:0.961, h:0.236, fontSize:8, color:C.WHITE, align:"right", wrap:false });
  icon(s, { x:2.698, y:4.692, w:0.214, h:0.211 }, C.WHITE);
  icon(s, { x:1.613, y:5.356, w:0.214, h:0.229 }, C.WHITE);
  icon(s, { x:1.891, y:4.049, w:0.214, h:0.152 }, C.WHITE);
  icon(s, { x:2.372, y:3.385, w:0.214, h:0.189 }, C.WHITE);
  icon(s, { x:2.698, y:2.721, w:0.214, h:0.195 }, C.WHITE);
  tx(s, "INFOGRAPHIC ELEMENTS", { x:7.832, y:3.142, w:3.889, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.CRAS, { x:7.832, y:4.769, w:3.889, h:0.606, fontSize:10, color:C.INK });
  return s;
}

// Slide 24 — Infographic — ribbon steps
function slide24(p) {
  const s = newSlide(p);
  drop(s, { x:6.19, y:2.796, w:1.256, h:0.404 }, C.DARK);
  wedge(s, { x:6.708, y:2.463, w:0.985, h:0.489, rotate:90 }, C.SKY);
  s.addShape("rect", { x:6.19, y:2.211, w:0.767, h:0.767, fill:C.BLUE });
  drop(s, { x:6.19, y:3.738, w:1.253, h:0.218 }, C.GREYA);
  wedge(s, { x:6.781, y:3.292, w:0.84, h:0.489, rotate:90 }, C.COAL);
  s.addShape("rect", { x:6.19, y:3.112, w:0.767, h:0.767, fill:C.CHAR });
  drop(s, { x:6.19, y:4.845, w:1.253, h:0.218 }, C.DARK);
  wedge(s, { x:6.781, y:5.02, w:0.84, h:0.489, rotate:270 }, C.SKY);
  s.addShape("rect", { x:6.19, y:4.922, w:0.767, h:0.767, fill:C.BLUE });
  s.addShape("rect", { x:6.192, y:4.021, w:0.767, h:0.767, fill:C.BLUE });
  wedge(s, { x:6.818, y:4.16, w:0.767, h:0.487, rotate:90 }, C.SKY);
  banner(s, { x:7.438, y:2.568, w:3.418, h:0.632 }, C.BLUE);
  banner(s, { x:7.438, y:4.088, w:3.418, h:0.631 }, C.BLUE);
  banner(s, { x:7.438, y:4.847, w:3.418, h:0.631 }, C.BLUE);
  banner(s, { x:7.438, y:3.329, w:3.418, h:0.631 }, C.CHAR);
  tx(s, "01", { x:6.3, y:2.456, w:0.546, h:0.269, fontSize:10, color:C.WHITE, charSpacing:3, align:"center" });
  tx(s, "02", { x:6.3, y:3.333, w:0.546, h:0.269, fontSize:10, color:C.WHITE, charSpacing:3, align:"center" });
  tx(s, "03", { x:6.3, y:4.266, w:0.546, h:0.269, fontSize:10, color:C.WHITE, charSpacing:3, align:"center" });
  tx(s, "04", { x:6.3, y:5.13, w:0.546, h:0.269, fontSize:10, color:C.WHITE, charSpacing:3, align:"center" });
  icon(s, { x:10.971, y:4.217, w:0.365, h:0.273 }, C.WHITE);
  icon(s, { x:11.074, y:4.394, w:0.158, h:0.156 }, C.WHITE);
  icon(s, { x:11.005, y:2.707, w:0.296, h:0.333 }, C.COAL);
  icon(s, { x:10.971, y:3.462, w:0.365, h:0.333 }, C.COAL);
  icon(s, { x:10.985, y:4.972, w:0.337, h:0.333 }, C.COAL);
  tx(s, "Lorem Ipsum", { x:7.683, y:2.683, w:1.662, h:0.326, fontSize:10, color:C.WHITE, charSpacing:3, valign:"middle" });
  tx(s, "Dolor", { x:7.683, y:3.467, w:1.662, h:0.326, fontSize:10, color:C.WHITE, charSpacing:3, valign:"middle" });
  tx(s, "Ipsum", { x:7.683, y:4.223, w:1.662, h:0.326, fontSize:10, color:C.WHITE, charSpacing:3, valign:"middle" });
  tx(s, "Dolor Lorem", { x:7.683, y:4.98, w:1.662, h:0.326, fontSize:10, color:C.WHITE, charSpacing:3, valign:"middle" });
  tx(s, "INFOGRAPHIC ELEMENTS", { x:0.574, y:3.242, w:3.999, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.CRAS, { x:0.574, y:4.69, w:4.214, h:0.606, fontSize:10, color:C.INK });
  return s;
}

// Slide 25 — Our Infographic — circles
function slide25(p) {
  const s = newSlide(p);
  s.addShape("ellipse", { x:4.478, y:3.189, w:1.714, h:1.714, fill:C.BLUE });
  s.addShape("ellipse", { x:7.101, y:3.179, w:1.714, h:1.714, fill:C.SKY });
  s.addShape("ellipse", { x:1.861, y:3.169, w:1.714, h:1.714, fill:C.NAVY });
  icon(s, { x:7.613, y:3.678, w:0.689, h:0.689 }, C.WHITE);
  icon(s, { x:2.24, y:3.551, w:0.915, h:0.915 }, C.WHITE);
  icon(s, { x:4.986, y:3.777, w:0.689, h:0.689 }, C.WHITE);
  s.addShape("ellipse", { x:9.706, y:3.189, w:1.714, h:1.714, fill:C.CYAN });
  icon(s, { x:10.254, y:3.68, w:0.731, h:0.731 }, C.WHITE);
  tx(s, T.L3s, { x:9.408, y:5.728, w:2.424, h:0.808, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:9.674, y:5.362, w:1.891, h:0.303, fontSize:12, bold:true, align:"center" });
  tx(s, T.L1c, { x:6.675, y:1.443, w:3.279, h:0.278, fontSize:10.5, bold:true });
  tx(s, "OUR INFOGRAPHIC", { x:1.467, y:1.432, w:4.208, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:6.667, y:1.764, w:4.481, h:0.808, fontSize:10.5 });
  tx(s, T.L3s, { x:6.836, y:5.728, w:2.424, h:0.808, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:7.103, y:5.362, w:1.891, h:0.303, fontSize:12, bold:true, align:"center" });
  tx(s, T.L3s, { x:4.154, y:5.728, w:2.424, h:0.808, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:4.42, y:5.362, w:1.891, h:0.303, fontSize:12, bold:true, align:"center" });
  tx(s, T.L3s, { x:1.519, y:5.728, w:2.424, h:0.808, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:1.786, y:5.362, w:1.891, h:0.303, fontSize:12, bold:true, align:"center" });
  s.addShape("arc", { x:4.177, y:2.947, w:2.311, h:2.55, fill:{ type:"none" }, line:{ color:C.NAVY, width:1 }, angleRange:[195.945, 344.129] });
  s.addShape("arc", { x:9.43, y:2.947, w:2.311, h:2.55, fill:{ type:"none" }, line:{ color:C.NAVY, width:1 }, angleRange:[195.945, 344.129] });
  s.addShape("arc", { x:6.819, y:2.577, w:2.311, h:2.55, fill:{ type:"none" }, line:{ color:C.NAVY, width:1 }, angleRange:[195.945, 344.129], rotate:180 });
  s.addShape("arc", { x:1.525, y:2.577, w:2.311, h:2.55, fill:{ type:"none" }, line:{ color:C.NAVY, width:1 }, angleRange:[195.945, 344.129], rotate:180 });
  return s;
}

// Slide 26 — Our Infographic — venn
function slide26(p) {
  const s = newSlide(p);
  s.addShape("ellipse", { x:6.228, y:3.636, w:1.833, h:1.833, fill:C.SKY });
  s.addShape("ellipse", { x:5.066, y:3.636, w:1.833, h:1.833, fill:C.BLUE });
  s.addShape("ellipse", { x:5.647, y:2.341, w:1.833, h:1.833, fill:C.NAVY });
  icon(s, { x:6.898, y:4.239, w:0.58, h:0.58 }, C.WHITE);
  icon(s, { x:6.295, y:2.884, w:0.58, h:0.58 }, C.WHITE);
  icon(s, { x:5.645, y:4.347, w:0.519, h:0.519 }, C.WHITE);
  tx(s, T.L3s, { x:9.634, y:2.51, w:2.424, h:0.808, fontSize:10.5, align:"justify" });
  tx(s, T.LD, { x:9.635, y:2.235, w:1.891, h:0.278, fontSize:10.5 });
  icon(s, { x:8.984, y:2.469, w:0.58, h:0.58 }, C.WHITE);
  tx(s, T.L3s, { x:9.653, y:3.699, w:2.424, h:0.808, fontSize:10.5, align:"justify" });
  tx(s, T.LD, { x:9.654, y:3.424, w:1.891, h:0.278, fontSize:10.5 });
  tx(s, T.L3s, { x:9.674, y:4.889, w:2.424, h:0.808, fontSize:10.5, align:"justify" });
  tx(s, T.LD, { x:9.675, y:4.614, w:1.891, h:0.278, fontSize:10.5 });
  icon(s, { x:8.943, y:3.617, w:0.58, h:0.58 }, C.WHITE);
  icon(s, { x:9.044, y:4.813, w:0.519, h:0.519 }, C.WHITE);
  tx(s, T.L2s, { x:0.338, y:3.972, w:3.279, h:0.454, fontSize:10.5, bold:true });
  tx(s, "OUR INFOGRAPHIC", { x:0.338, y:2.587, w:4.297, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:0.338, y:4.535, w:3.402, h:1.161, fontSize:10.5 });
  return s;
}

// Slide 27 — Our Infographic — percentages
function slide27(p) {
  const s = newSlide(p);
  tx(s, T.L2s, { x:0.338, y:3.972, w:3.279, h:0.454, fontSize:10.5, bold:true });
  tx(s, "OUR INFOGRAPHIC", { x:0.338, y:2.587, w:4.295, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:0.338, y:4.535, w:3.402, h:1.161, fontSize:10.5 });
  s.addShape("ellipse", { x:10.197, y:2.068, w:1.833, h:1.833, fill:C.NAVY });
  icon(s, { x:10.84, y:2.694, w:0.546, h:0.546 }, C.WHITE);
  s.addShape("ellipse", { x:7.776, y:2.068, w:1.833, h:1.833, fill:C.DARK });
  icon(s, { x:8.347, y:2.639, w:0.69, h:0.69 }, C.WHITE);
  tx(s, T.L3s, { x:5.407, y:5.116, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  s.addShape("roundRect", { x:5.678, y:4.038, w:1.529, h:0.58, fill:C.BLACK, rectRadius:0.162 });
  s.addShape("ellipse", { x:5.526, y:2.068, w:1.833, h:1.833, fill:C.GREY74 });
  icon(s, { x:6.152, y:2.694, w:0.58, h:0.58 }, C.WHITE);
  tx(s, "100%", { x:5.886, y:4.095, w:1.112, h:0.404, color:C.WHITE, align:"center" });
  tx(s, T.LD, { x:5.497, y:4.841, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, T.L3s, { x:7.657, y:5.116, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:7.747, y:4.841, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, T.L3s, { x:10.078, y:5.116, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:10.168, y:4.841, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  s.addShape("roundRect", { x:7.9, y:4.038, w:1.529, h:0.58, fill:C.BLACK, rectRadius:0.162 });
  tx(s, "100%", { x:8.108, y:4.095, w:1.112, h:0.404, color:C.WHITE, align:"center" });
  s.addShape("roundRect", { x:10.313, y:4.038, w:1.529, h:0.58, fill:C.BLACK, rectRadius:0.162 });
  tx(s, "100%", { x:10.521, y:4.095, w:1.112, h:0.404, color:C.WHITE, align:"center" });
  return s;
}

// Slide 28 — SWOT
function slide28(p) {
  const s = newSlide(p);
  s.addShape("roundRect", { x:4.515, y:2.829, w:1.47, h:1.47, fill:C.BLUE, rectRadius:0.146, rotate:45 });
  tx(s, "S", { x:4.792, y:2.88, w:0.851, h:1.447, fontSize:80, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:4.143, y:5.044, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:4.233, y:4.769, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  s.addShape("roundRect", { x:6.676, y:2.838, w:1.47, h:1.47, fill:C.BLUE, rectRadius:0.146, rotate:45 });
  tx(s, "W", { x:6.996, y:2.889, w:0.851, h:1.447, fontSize:80, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:6.304, y:5.053, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:6.394, y:4.778, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  s.addShape("roundRect", { x:8.836, y:2.838, w:1.47, h:1.47, fill:C.BLUE, rectRadius:0.146, rotate:45 });
  tx(s, "O", { x:9.146, y:2.889, w:0.851, h:1.447, fontSize:80, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:8.464, y:5.053, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:8.554, y:4.778, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  s.addShape("roundRect", { x:10.922, y:2.838, w:1.47, h:1.47, fill:C.BLUE, rectRadius:0.146, rotate:45 });
  tx(s, "T", { x:11.238, y:2.889, w:0.851, h:1.447, fontSize:80, color:C.WHITE, align:"center" });
  tx(s, T.L3s, { x:10.549, y:5.053, w:2.071, h:0.985, fontSize:10.5, align:"center" });
  tx(s, T.LD, { x:10.639, y:4.778, w:1.891, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, T.L2s, { x:0.338, y:4.099, w:3.279, h:0.454, fontSize:10.5, bold:true });
  tx(s, "SWOT", { x:0.338, y:3.27, w:3.737, h:0.707, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:0.338, y:4.663, w:3.402, h:1.161, fontSize:10.5 });
  return s;
}

// Slide 29 — SWOT Progress
function slide29(p) {
  const s = newSlide(p);
  tx(s, "SWOT PROGRESS", { x:0.459, y:2.594, w:4.244, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, "Weaknesses", { x:9.474, y:4.149, w:2.464, h:0.337, fontSize:14 });
  tx(s, T.L3s, { x:9.476, y:4.465, w:2.805, h:0.631, fontSize:10.5 });
  tx(s, " Threats", { x:9.474, y:5.229, w:2.464, h:0.337, fontSize:14 });
  tx(s, T.L3s, { x:9.476, y:5.544, w:2.805, h:0.631, fontSize:10.5 });
  quadRing(s, { x:5.51, y:2.594, w:1.556, h:1.556 }, "br", C.NAVY);
  quadRing(s, { x:7.114, y:2.594, w:1.556, h:1.556 }, "bl", C.BLUE);
  quadRing(s, { x:5.512, y:4.201, w:1.556, h:1.556 }, "tr", C.GREY40);
  quadRing(s, { x:7.116, y:4.201, w:1.556, h:1.556 }, "tl", C.BLACK);
  tx(s, "Strengths", { x:9.476, y:1.901, w:1.187, h:0.337, fontSize:14 });
  tx(s, T.L3s, { x:9.476, y:2.216, w:2.805, h:0.631, fontSize:10.5 });
  tx(s, "Opportunities", { x:9.476, y:3.074, w:2.464, h:0.337, fontSize:14 });
  tx(s, T.L3s, { x:9.476, y:3.389, w:2.805, h:0.631, fontSize:10.5 });
  tx(s, "S", { x:5.795, y:2.846, w:0.996, h:1.01, fontSize:54, align:"center" });
  tx(s, "W", { x:7.394, y:2.846, w:0.996, h:1.01, fontSize:54, align:"center" });
  tx(s, "O", { x:5.795, y:4.502, w:0.996, h:1.01, fontSize:54, align:"center" });
  tx(s, "T", { x:7.394, y:4.502, w:0.996, h:1.01, fontSize:54, align:"center" });
  tx(s, T.L2s, { x:0.459, y:4.179, w:3.279, h:0.454, fontSize:10.5, bold:true });
  tx(s, T.L6, { x:0.459, y:4.743, w:3.402, h:1.161, fontSize:10.5 });
  return s;
}

// Slide 30 — Our Infographic — bar chart
function slide30(p) {
  const s = newSlide(p);
  rule(s, { x:8.238, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  rule(s, { x:7.048, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  rule(s, { x:5.907, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  rule(s, { x:9.458, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  rule(s, { x:10.632, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  rule(s, { x:11.821, y:1.695, w:0, h:4.18, color:C.GREYA6, width:1 });
  s.addShape("rect", { x:5.903, y:2.124, w:5.387, h:0.174, fill:C.NAVY });
  s.addShape("rect", { x:5.903, y:2.368, w:3.383, h:0.174, fill:C.CHAR });
  s.addShape("rect", { x:5.903, y:1.88, w:5.135, h:0.174, fill:C.SILVER });
  s.addShape("rect", { x:5.903, y:3.149, w:5.387, h:0.174, fill:C.NAVY });
  s.addShape("rect", { x:5.903, y:3.392, w:3.383, h:0.174, fill:C.CHAR });
  s.addShape("rect", { x:5.903, y:2.905, w:5.135, h:0.174, fill:C.SILVER });
  s.addShape("rect", { x:5.903, y:4.243, w:5.387, h:0.174, fill:C.NAVY });
  s.addShape("rect", { x:5.903, y:4.486, w:3.383, h:0.174, fill:C.CHAR });
  s.addShape("rect", { x:5.903, y:3.999, w:5.135, h:0.174, fill:C.SILVER });
  s.addShape("rect", { x:5.903, y:5.337, w:5.387, h:0.174, fill:C.NAVY });
  s.addShape("rect", { x:5.903, y:5.581, w:3.383, h:0.174, fill:C.CHAR });
  s.addShape("rect", { x:5.903, y:5.094, w:5.135, h:0.174, fill:C.SILVER });
  tx(s, "Chart 1", { x:4.859, y:2.09, w:0.936, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "Chart 2", { x:4.868, y:3.107, w:0.936, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "Chart 3", { x:4.868, y:4.17, w:0.936, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "Chart 4", { x:4.868, y:5.275, w:0.936, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "0", { x:5.669, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "1", { x:6.816, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "2", { x:7.995, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "3", { x:9.205, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "4", { x:10.363, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, "5", { x:11.569, y:6.03, w:0.468, h:0.278, fontSize:10.5, bold:true, align:"center" });
  tx(s, T.L2s, { x:0.338, y:3.972, w:3.279, h:0.454, fontSize:10.5, bold:true });
  tx(s, "OUR INFOGRAPHIC", { x:0.338, y:2.587, w:4.124, h:1.313, fontSize:36, fontFace:"Montserrat" });
  tx(s, T.L6, { x:0.338, y:4.535, w:3.402, h:1.161, fontSize:10.5 });
  return s;
}
// ------------------------------------------------------------------- build
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'HD16x9', width: W, height: H });
  pres.layout = 'HD16x9';
  pres.title = 'Home Based Healthcare';
  SLIDES.forEach((fn, i) => chrome(fn(pres), i + 1, i === 0 ? C.WHITE : C.BLACK));
  return pres;
}

build()
  .writeFile({ fileName: path.join(__dirname, '17c6fb16-a3bb-46f6-b643-cae1512df098_grok_final.pptx') })
  .then(f => console.log('wrote', f));
