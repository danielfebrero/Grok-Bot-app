/**
 * "SUSANA — Food for the Future" business-proposal deck (42 slides, 13.333 x 7.5 in).
 *
 * Rebuilt with pptxgenjs only.  Every position, colour, font and string below is a
 * plain literal so the deck's design can be read straight out of this file.
 * Photographs in the original are stand-ins ("REPLACE THIS IMAGE" tiles); they are
 * reproduced here as flat colour blocks tagged `[image]` via the `photo()` helper.
 *
 * Run:  node 004c9d90-862b-4339-8990-c1d6263098d8_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette */
// Theme "Susana": dk1=FF7537, lt1=F8F8F8, dk2=000000, accent1=2FC399, accent2=F8CD48
const C = {
  ink: '000000',        // theme tx2 — body copy & rules
  ink85: '262626',      // tx2 85% — soft headings
  ink75: '404040',      // tx2 75%
  ink65: '595959',      // tx2 65% — milestone titles
  grey50: '7C7C7C',     // bg1 50% — muted paragraph copy
  grey75: 'BABABA',     // bg1 75% — chart axes
  paper: 'F8F8F8',      // theme bg1 — slide background
  orange: 'FF7537',     // theme tx1 — brand accent
  green: '2FC399',      // theme accent1
  yellow: 'F8CD48',     // theme accent2
  sand: 'EAD7C0',       // theme accent6
  photoBlue: '1531BC',  // stand-in tone of the blue photo tiles
  photoRed: 'FC3514',   // stand-in tone of the red photo tiles
  device: 'C8C9C9',     // laptop / monitor mock-up grey
};

const F = { sans: 'PT Sans', display: 'Raleway Black', source: 'Source Sans Pro' };

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------- shared fragments */
const WWW = 'W   W   W   .   S   U   S   A   N   A   .   C   O   M';
const TAGLINE = '"Empowering Tomorrow\'s Nutrition:\nFood for the Future."';
const TAGLINE3 = '"Empowering\nTomorrow\'s Nutrition:\nFood for the Future."';

/* ----------------------------------------------------------- primitives */
// Text box. Defaults mirror the deck: PT Sans, black, zero inset, top aligned.
function T(slide, text, opts) {
  slide.addText(text, Object.assign({
    fontFace: F.sans, fontSize: 18, color: C.ink, margin: 0,
    align: 'left', valign: 'top',
  }, opts));
}

// Plain shape (rect / ellipse / line / preset silhouette).
function S(slide, kind, opts) {
  slide.addShape(kind, Object.assign({ line: { type: 'none' } }, opts));
}

// Stand-in for a photograph: flat colour block with a small caption.
function photo(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
  slide.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle', margin: 0,
    fontFace: F.sans, fontSize: 11, bold: true, color: 'FFFFFF', transparency: 25,
  });
}

// Left rail present on every content slide: colour bar, hairline, rotated captions.
function chrome(slide, barColor, pageLabel, sectionLabel) {
  const cover = sectionLabel === 'Food For The Future';
  S(slide, 'rect', { x: 0, y: 0, w: 0.708, h: 7.5, fill: { color: barColor } });
  S(slide, 'line', { x: 0.708, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
  S(slide, 'line', {
    x: 0.354, y: cover ? 3.035 : 2.042, w: 0, h: cover ? 2.256 : 3.22,
    line: { color: C.ink, width: 1.5 },
  });
  T(slide, pageLabel, {
    x: cover ? -0.394 : -0.453, y: cover ? 6.272 : 6.213, w: cover ? 1.496 : 1.614,
    h: 0.202, rotate: 270, flipH: true, fontSize: 12, bold: true, charSpacing: 3,
  });
  T(slide, sectionLabel, {
    x: cover ? -0.824 : -0.355, y: cover ? 1.456 : 0.987, w: cover ? 2.356 : 1.417,
    h: 0.202, rotate: 270, flipH: true, align: 'right',
    fontSize: 12, bold: true, charSpacing: 3,
  });
}

// Small brand pictograms (section markers on slides 8, 18 and 29), drawn from
// two or three primitives each rather than embedded artwork.
function icon(slide, kind, x, y, w, h, color) {
  const box = (dx, dy, dw, dh, c, radius) => S(slide, radius ? 'roundRect' : 'rect', {
    x: x + w * dx, y: y + h * dy, w: w * dw, h: h * dh, fill: { color: c }, rectRadius: radius,
  });
  switch (kind) {
    case 'devices':                            // tablet beside a phone
      box(0, 0, 0.6, 1, color, 0.04); box(0.07, 0.16, 0.46, 0.66, C.paper);
      box(0.68, 0.2, 0.32, 0.8, color, 0.03); break;
    case 'envelope':                           // sealed envelope with a flap
      box(0, 0.1, 1, 0.8, color);
      S(slide, 'triangle', { x: x + w * 0.06, y: y + h * 0.16, w: w * 0.88, h: h * 0.55,
        flipV: true, fill: { color: C.paper } }); break;
    case 'screen':                             // monitor / cast surface
      box(0, 0, 1, 0.78, color, 0.04); box(0.1, 0.13, 0.8, 0.5, C.paper);
      box(0.38, 0.78, 0.24, 0.22, color); break;
    case 'briefcase':                          // case with a handle
      box(0.3, 0, 0.4, 0.18, color, 0.03); box(0, 0.18, 1, 0.82, color, 0.04); break;
    case 'building':                           // office block with windows
      box(0, 0, 1, 1, color);
      [0.15, 0.45, 0.75].forEach((r) => [0.15, 0.55].forEach((c) => box(c, r, 0.3, 0.16, C.paper)));
      break;
    case 'phone':                              // handset
      box(0.1, 0, 0.8, 1, color, 0.04); box(0.24, 0.14, 0.52, 0.62, C.paper); break;
    default: box(0, 0, 1, 1, color);
  }
}

// Cog used by the risk-analysis illustration on slide 26.
function gear(slide, x, y, w, h, color) {
  S(slide, 'gear9', { x, y, w, h, fill: { color } });
  S(slide, 'donut', {
    x: x + w * 0.2, y: y + h * 0.2, w: w * 0.6, h: h * 0.6, fill: { color: C.paper },
  });
  S(slide, 'roundRect', {
    x: x + w * 0.38, y: y + h * 0.38, w: w * 0.24, h: h * 0.24,
    fill: { color }, rectRadius: 0.03,
  });
}

// Laptop mock-up (slide 18): dark bezel, photo screen, light grey base.
function laptop(slide, x, y, w, h, sx, sy, sw, sh, screenColor) {
  S(slide, 'rect', { x: sx - 0.11, y, w: sw + 0.22, h: sh + 0.24, fill: { color: C.ink85 } });
  photo(slide, sx, sy, sw, sh, screenColor);
  S(slide, 'trapezoid', { x, y: y + h - 0.19, w, h: 0.19, fill: { color: C.device } });
}

// Desktop-monitor mock-up (slide 29): bezel, photo screen, neck and foot.
function monitor(slide, x, y, w, h, sx, sy, sw, sh, screenColor) {
  S(slide, 'roundRect', { x, y, w, h: sh + 0.42, fill: { color: C.ink85 }, rectRadius: 0.08 });
  photo(slide, sx, sy, sw, sh, screenColor);
  S(slide, 'trapezoid', {
    x: x + w * 0.38, y: y + sh + 0.42, w: w * 0.24, h: h - sh - 0.68,
    fill: { color: C.device },
  });
  S(slide, 'roundRect', {
    x: x + w * 0.2, y: y + h - 0.26, w: w * 0.6, h: 0.26,
    fill: { color: C.device }, rectRadius: 0.05,
  });
}

// Flat figure holding a pie chart — the illustration on slide 27.
function person(slide) {
  S(slide, 'pie', { x: 5.833, y: 3.399, w: 1.187, h: 1.29, fill: { color: C.yellow }, angleRange: [40, 350] });
  S(slide, 'rtTriangle', { x: 6.618, y: 3.38, w: 0.64, h: 0.802, flipH: true, fill: { color: C.orange } });
  S(slide, 'roundRect', { x: 7.388, y: 4.349, w: 0.746, h: 2.181, fill: { color: C.yellow }, rectRadius: 0.12 });
  S(slide, 'roundRect', { x: 7.139, y: 6.392, w: 1.0, h: 0.423, fill: { color: C.ink85 }, rectRadius: 0.12 });
  S(slide, 'roundRect', { x: 7.272, y: 2.778, w: 0.936, h: 1.833, fill: { color: C.orange }, rectRadius: 0.16 });
  S(slide, 'roundRect', { x: 7.143, y: 2.241, w: 0.575, h: 0.529, fill: { color: C.ink85 }, rectRadius: 0.16 });
  S(slide, 'roundRect', { x: 7.194, y: 2.443, w: 0.484, h: 0.557, fill: { color: C.sand }, rectRadius: 0.16 });
  S(slide, 'roundRect', { x: 7.612, y: 2.931, w: 0.41, h: 0.803, fill: { color: C.orange }, rectRadius: 0.1 });
  S(slide, 'roundRect', { x: 6.903, y: 3.597, w: 1.135, h: 0.294, fill: { color: C.sand }, rectRadius: 0.08 });
  S(slide, 'roundRect', { x: 6.829, y: 3.861, w: 0.553, h: 0.227, fill: { color: C.sand }, rectRadius: 0.08 });
}

/* ------------------------------------------------------------- slides */
// 1. Presentation Susana
function slide01(s) {
  S(s, 'rect', { x: 0, y: 0, w: 5.885, h: 7.504, fill: { color: C.green } });
  T(s, TAGLINE, { x: 9.588, y: 6.245, w: 2.958, h: 0.471, align: 'right', fontSize: 14, bold: true });
  T(s, WWW, { x: 10.341, y: 0.783, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  chrome(s, C.yellow, 'Page 01 / 30', 'Food For The Future');
  T(s, 'Proactively envisioned for multimedia based expertise and cross-media into books good growth strategies. Seamlessly visualize on quality intellectual capital Seamlessly base visualize intellectual capital administrates making empowered markets', { x: 6.667, y: 4.468, w: 4.74, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'PRESENTATION', { x: 6.667, y: 3.569, w: 4.957, h: 0.74, flipH: true, fontSize: 44, fontFace: F.display });
  T(s, 'SUSANA', { x: 6.667, y: 2.374, w: 4.957, h: 1.346, flipH: true, fontSize: 80, fontFace: F.display });
  photo(s, 1.494, 0.783, 4.391, 5.933, C.photoBlue);
  S(s, 'line', { x: 5.885, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
}

// 2. Greetings We Are Susana
function slide02(s) {
  T(s, 'GREETINGS', { x: 7.519, y: 0.747, w: 4.696, h: 0.673, flipH: true, align: 'right', fontSize: 40, fontFace: F.display, color: C.ink85 });
  T(s, 'WE ARE SUSANA', { x: 7.519, y: 1.42, w: 4.696, h: 0.337, flipH: true, align: 'right', fontSize: 20, fontFace: F.display, color: C.ink85 });
  T(s, 'Collaboratively administrates making the empowered markets plug playing networks procrastinate user installed globally standards channels scalable benefits.', { x: 1.826, y: 1.037, w: 4.575, h: 0.43, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, color: C.ink85 });
  T(s, 'Freya Schmidt', { x: 7.519, y: 5.183, w: 4.696, h: 0.404, align: 'right', valign: 'middle', fontSize: 24, bold: true });
  T(s, 'Proactively envisioned based expertise media growth into made it quality collaboration leverage agile for frameworks the provide make market global thinking pursue scalable good collaboratively administrates making empowered markets.', { x: 7.519, y: 6.095, w: 4.696, h: 0.657, align: 'right', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Head of Susana Corporate', { x: 7.519, y: 5.587, w: 4.696, h: 0.269, align: 'right', valign: 'middle', fontSize: 16 });
  T(s, WWW, { x: 0.737, y: 5.529, w: 2.313, h: 0.135, rotate: 270, fontSize: 8 });
  chrome(s, C.orange, 'Page 02 / 30', 'Food For The Future');
  photo(s, 2.661, 2.504, 3.74, 4.248, C.photoRed);
}

// 3. Background
function slide03(s) {
  S(s, 'rect', { x: 7.428, y: 0, w: 5.906, h: 7.5, fill: { color: C.green } });
  T(s, WWW, { x: 6.693, y: 4.926, w: 2.651, h: 0.135, rotate: 270, fontSize: 8 });
  chrome(s, C.green, 'Page 03 / 30', 'Chapter 01');
  T(s, 'The global food industry is undergoing a paradigm shift, driven by the increasing demand for sustainable and nutritious food sources. This chapter provides an overview of the evolving landscape and sets the stage for our business venture, "Food for the Future."', { x: 1.666, y: 5.661, w: 4.803, h: 0.657, flipH: true, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Your Title Name Here', { x: 1.666, y: 5.289, w: 4.803, h: 0.185, fontSize: 11, bold: true });
  T(s, 'BACKGROUND', { x: 1.666, y: 1.434, w: 4.803, h: 0.74, valign: 'middle', fontSize: 44, fontFace: F.display });
  T(s, 'Introduction to "Food for the Future" Business', { x: 1.666, y: 1.182, w: 4.803, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, '“To be fully seen by somebody, then, be loved anyhow this is a human offering”', { x: 8.609, y: 1.187, w: 3.543, h: 0.471, valign: 'middle', fontSize: 14, italic: true });
  S(s, 'line', { x: 7.428, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
  photo(s, 8.609, 2.776, 3.543, 3.543, C.photoBlue);
}

// 4. Purpose Of The Business
function slide04(s) {
  chrome(s, C.yellow, 'Page 04 / 30', 'Chapter 01');
  T(s, 'Our mission is to contribute to a sustainable and resilient future by revolutionizing the way we produce, distribute, and consume food. This chapter outlines the core objectives and values that guide our business in addressing the challenges and opportunities in the food industry.', { x: 10.238, y: 4.449, w: 1.969, h: 1.793, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'PURPOSE OF THE BUSINESS', { x: 7.794, y: 1.533, w: 4.412, h: 1.481, valign: 'middle', fontSize: 44, fontFace: F.display });
  T(s, 'Introduction to "Food for the Future" Business', { x: 7.794, y: 1.258, w: 4.412, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, TAGLINE3, { x: 7.794, y: 5.535, w: 1.964, h: 0.707, flipH: true, fontSize: 14, bold: true });
  photo(s, 1.317, 0.61, 2.368, 3.682, C.photoRed);
  photo(s, 1.317, 4.49, 2.368, 2.399, C.photoBlue);
  photo(s, 3.859, 1.258, 2.808, 4.984, C.photoRed);
}

// 5. Vision And Impact
function slide05(s) {
  photo(s, 0.708, 3.284, 7.054, 3.079, C.photoBlue);
  T(s, 'We articulate our vision for "Food for the Future" and the impact we aim to make on global food security, environmental sustainability, and the health and well-being of consumers.', { x: 8.775, y: 5.705, w: 3.546, h: 0.658, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  chrome(s, C.orange, 'Page 05 / 30', 'Chapter 01');
  T(s, 'VISION AND IMPACT', { x: 1.738, y: 1.406, w: 6.024, h: 0.74, align: 'right', valign: 'middle', fontSize: 44, fontFace: F.display });
  T(s, 'Introduction to "Food for the Future" Business', { x: 1.738, y: 1.137, w: 6.024, h: 0.269, align: 'right', valign: 'middle', fontSize: 16 });
  photo(s, 8.775, 1.137, 3.546, 3.432, C.photoRed);
}

// 6. Industry Overview
function slide06(s) {
  photo(s, 0.708, 0.908, 3.797, 6.592, C.photoRed);
  chrome(s, C.green, 'Page 06 / 30', 'Chapter 02');
  T(s, 'This chapter provides an in-depth analysis of the current state of the food industry, highlighting key trends, challenges, and opportunities. Understanding the dynamics of the industry is crucial for positioning "Food for the Future" strategically.', { x: 5.492, y: 5.08, w: 3.331, h: 0.884, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Text Title Goes Here', { x: 5.492, y: 4.755, w: 3.331, h: 0.202, align: 'justify', fontSize: 12, bold: true });
  T(s, WWW, { x: 9.81, y: 6.665, w: 2.402, h: 0.135, fontSize: 8 });
  T(s, 'INDUSTRY OVERVIEW', { x: 5.569, y: 1.177, w: 6.535, h: 0.74, valign: 'middle', fontSize: 44, fontFace: F.display });
  T(s, 'The Current State of the Food Industry', { x: 5.569, y: 0.908, w: 6.535, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, TAGLINE, { x: 5.492, y: 3.194, w: 3.331, h: 0.471, flipH: true, fontSize: 14, bold: true });
  photo(s, 9.81, 3.194, 3.523, 2.771, C.photoBlue);
}

// 7. Challenges In The Current Food System
function slide07(s) {
  chrome(s, C.yellow, 'Page 07 / 30', 'Chapter 02');
  T(s, WWW, { x: 10.555, y: 1.48, w: 2.385, h: 0.135, align: 'right', fontSize: 8 });
  T(s, 'CHALLENGES IN THE CURRENT FOOD SYSTEM', { x: 1.099, y: 1.076, w: 6.414, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'The Current State of the Food Industry', { x: 1.099, y: 0.807, w: 6.414, h: 0.269, valign: 'middle', fontSize: 16 });
  S(s, 'rect', { x: 3.916, y: 3.095, w: 3.596, h: 3.598, fill: { color: C.yellow } });
  T(s, 'Your Text Goes Here', { x: 4.563, y: 4.149, w: 2.304, h: 0.236, valign: 'middle', fontSize: 14, bold: true });
  T(s, 'We delve into the challenges faced by the existing food system, including issues related to food waste, environmental impact, and the need for more sustainable and nutritious food options.', { x: 4.563, y: 4.528, w: 2.304, h: 1.112, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  photo(s, 1.099, 3.095, 2.426, 3.598, C.photoBlue);
  photo(s, 7.904, 3.095, 5.43, 3.598, C.photoRed);
}

// 9. Plant-Based Alternatives
function slide09(s) {
  photo(s, 0.708, 0, 6.72, 7.5, C.photoRed);
  chrome(s, C.green, 'Page 09 / 30', 'Chapter 04');
  T(s, 'We detail our plant-based food offerings, emphasizing the benefits of plant-centric diets for both human health and the environment.', { x: 8.339, y: 5.774, w: 4.084, h: 0.43, flipH: true, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Your Title Name Here', { x: 8.339, y: 5.402, w: 4.084, h: 0.185, fontSize: 11, bold: true });
  T(s, 'PLANT-BASED ALTERNATIVES', { x: 8.339, y: 1.564, w: 4.084, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Product Line and Offerings', { x: 8.339, y: 1.295, w: 4.084, h: 0.269, valign: 'middle', fontSize: 16 });
  S(s, 'line', { x: 7.428, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
}

// 10. Cultivated And Lab-Grown Products
function slide10(s) {
  photo(s, 8.018, 1.195, 4.134, 6.319, C.photoBlue);
  T(s, '“ To be fully seen by somebody, then, and be loved anyhow this is human offering ”', { x: 5.157, y: 3.82, w: 5.079, h: 0.539, flipH: true, fontSize: 16, bold: true });
  chrome(s, C.yellow, 'Page 10 / 30', 'Chapter 04');
  T(s, 'This section explores our commitment to exploring alternative protein sources, including cultivated and lab-grown meat, addressing the challenges of traditional livestock farming.', { x: 1.351, y: 5.775, w: 6.024, h: 0.43, flipH: true, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Your Title Name Here', { x: 1.351, y: 5.402, w: 6.024, h: 0.185, fontSize: 11, bold: true });
  T(s, 'CULTIVATED AND\nLAB-GROWN PRODUCTS', { x: 1.351, y: 1.565, w: 6.024, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Product Line and Offerings', { x: 1.351, y: 1.295, w: 6.024, h: 0.269, valign: 'middle', fontSize: 16 });
}

// 11. Functional Foods
function slide11(s) {
  S(s, 'rect', { x: 0.708, y: 0, w: 5.853, h: 7.5, fill: { color: C.yellow } });
  chrome(s, C.orange, 'Page 11 / 30', 'Chapter 04');
  T(s, 'We introduce our range of functional foods designed to promote health and well-being, incorporating ingredients with proven nutritional benefits.', { x: 9.277, y: 1.378, w: 3.228, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, '“ To be fully seen by somebody, then, and be loved anyhow this is human offering ”', { x: 7.388, y: 3.372, w: 5.118, h: 0.539, flipH: true, align: 'justify', fontSize: 16, bold: true });
  T(s, 'FUNCTIONAL FOODS', { x: 7.388, y: 5.516, w: 5.118, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Product Line and Offerings', { x: 7.388, y: 5.247, w: 5.118, h: 0.269, valign: 'middle', fontSize: 16 });
  S(s, 'line', { x: 6.561, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
  photo(s, 2.086, 1.378, 3.097, 4.744, C.photoRed);
}

// 12. Market Analysis And Target Audience
function slide12(s) {
  chrome(s, C.green, 'Page 12 / 30', 'Chapter 05');
  S(s, 'rect', { x: 1.221, y: 4.046, w: 4.645, h: 2.942, fill: { color: C.yellow } });
  T(s, 'Proactively envisioned based expertise cross into media growths strategies. Interactively into made visualize quality collaboration. Leverage agile for frameworks into the provide for make market global thinking pursue scalable administrates.', { x: 1.811, y: 5.282, w: 3.465, h: 0.885, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, fontFace: F.source });
  T(s, 'Your Text Here', { x: 1.811, y: 4.867, w: 3.465, h: 0.236, valign: 'middle', fontSize: 14, bold: true, fontFace: F.source, color: C.ink85 });
  T(s, 'Market Trends and Opportunities', { x: 6.379, y: 1.434, w: 2.525, h: 0.202, align: 'justify', fontSize: 12, bold: true, fontFace: F.source });
  T(s, 'This chapter provides a comprehensive analysis of market trends, identifying areas of growth and innovation in the food industry that "Food for the Future" aims to capitalize on.', { x: 6.379, y: 1.727, w: 2.525, h: 0.885, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, fontFace: F.source });
  T(s, 'Target Audience', { x: 10.295, y: 1.434, w: 2.525, h: 0.202, align: 'justify', fontSize: 12, bold: true, fontFace: F.source });
  T(s, 'We define our target audience, considering demographics, psychographics, and consumer behaviors, ensuring a strategic approach to product positioning and marketing.', { x: 10.295, y: 1.727, w: 2.525, h: 0.885, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, fontFace: F.source });
  T(s, 'MARKET ANALYSIS\nAND\nTARGET AUDIENCE', { x: 1.221, y: 1.215, w: 4.645, h: 1.616, valign: 'middle', fontSize: 32, fontFace: F.display });
  photo(s, 6.379, 4.046, 6.44, 2.942, C.photoBlue);
}

// 13. Business Model Overview
function slide13(s) {
  chrome(s, C.yellow, 'Page 13 / 30', 'Chapter 06');
  T(s, TAGLINE.replace('\n', ' '), { x: 1.798, y: 6.369, w: 3.046, h: 0.471, fontSize: 14, bold: true });
  T(s, 'This section outlines the fundamental aspects of our business model, encompassing the production process, distribution channels, and customer engagement strategies.', { x: 1.798, y: 5.279, w: 4.736, h: 0.43, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'BUSINESS\nMODEL OVERVIEW', { x: 1.798, y: 2.383, w: 4.736, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Business Model and Revenue Streams', { x: 1.798, y: 2.114, w: 4.736, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, WWW, { x: 9.592, y: 6.537, w: 2.651, h: 0.135, align: 'right', fontSize: 8 });
  photo(s, 7.625, 0, 5.709, 5.709, C.photoRed);
}

// 14. Revenue Streams
function slide14(s) {
  chrome(s, C.orange, 'Page 14 / 30', 'Chapter 06');
  T(s, 'REVENUE STREAMS', { x: 1.827, y: 1.493, w: 5.337, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Business Model and Revenue Streams', { x: 1.827, y: 1.224, w: 5.337, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, WWW, { x: 1.83, y: 6.139, w: 2.402, h: 0.135, fontSize: 8 });
  T(s, 'We explore the various revenue streams for "Food for the Future," including product sales, partnerships, and potential licensing agreements for our innovative food technologies.', { x: 1.827, y: 4.609, w: 3.425, h: 0.658, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, fontFace: F.source });
  T(s, TAGLINE, { x: 1.827, y: 3.266, w: 3.425, h: 0.471, flipH: true, fontSize: 14, bold: true });
  photo(s, 6.371, 3.266, 3.219, 4.234, C.photoBlue);
  photo(s, 10.119, 1.224, 3.214, 6.276, C.photoRed);
}

// 16. Ethical Supply Chain
function slide16(s) {
  photo(s, 9.021, -0.001, 4.312, 7.504, C.photoRed);
  photo(s, 0.708, -0.001, 3.275, 7.505, C.photoBlue);
  S(s, 'rect', { x: 3.983, y: 0, w: 5.038, h: 7.504, fill: { color: C.yellow } });
  chrome(s, C.yellow, 'Page 16 / 30', 'Chapter 07');
  T(s, 'This chapter discusses our dedication to ethical business practices, fair labor conditions, and transparency throughout our supply chain.', { x: 4.823, y: 4.139, w: 3.031, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, WWW, { x: 4.823, y: 6.212, w: 2.205, h: 0.135, flipH: true, fontSize: 8 });
  T(s, 'ETHICAL SUPPLY CHAIN', { x: 4.823, y: 1.421, w: 3.358, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, 'Sustainability and Ethical Practices', { x: 4.823, y: 1.157, w: 3.358, h: 0.269, valign: 'middle', fontSize: 16 });
  S(s, 'line', { x: 9.021, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
  S(s, 'line', { x: 3.983, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
}

// 17. Brand Identity
function slide17(s) {
  chrome(s, C.orange, 'Page 17 / 30', 'Chapter 08');
  T(s, WWW, { x: 5.052, y: 5.705, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  T(s, TAGLINE, { x: 1.587, y: 5.536, w: 3.028, h: 0.471, flipH: true, fontSize: 14, bold: true });
  T(s, 'Collaboratively administrate empowered markets plug play networks dynamic procrastinate users. After installed base benefits dramatic. for multimedia based expertise and cross-media into books good growth strategies. Seamlessly visualize on quality. After installed base benefits dramatic.', { x: 1.587, y: 3.656, w: 5.669, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'BRAND IDENTITY', { x: 1.587, y: 1.759, w: 5.669, h: 0.673, valign: 'middle', fontSize: 40, fontFace: F.display });
  T(s, 'Marketing and Branding Strategy', { x: 1.587, y: 1.493, w: 5.669, h: 0.269, valign: 'middle', fontSize: 16 });
  photo(s, 8.136, 1.493, 4.291, 6.007, C.photoBlue);
}

// 20. Regulatory Landscape
function slide20(s) {
  photo(s, 0.708, 2.5, 4.012, 5, C.photoBlue);
  chrome(s, C.orange, 'Page 20 / 30', 'Chapter 09');
  T(s, 'REGULATORY LANDSCAPE', { x: 6.26, y: 2.769, w: 6.137, h: 0.539, align: 'right', valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, 'Regulatory Compliance and Certifications', { x: 6.26, y: 2.5, w: 6.137, h: 0.269, align: 'right', valign: 'middle', fontSize: 16 });
  T(s, 'We address the regulatory aspects of the food industry, discussing compliance requirements, quality standards, and certifications necessary for ensuring the safety and legitimacy of our products.', { x: 9.405, y: 5.838, w: 2.992, h: 0.884, align: 'right', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, WWW, { x: 2.515, y: 1.183, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  T(s, TAGLINE, { x: 9.444, y: 1.014, w: 2.953, h: 0.471, flipH: true, align: 'right', fontSize: 14, bold: true });
  photo(s, 4.973, 5.329, 3.496, 2.171, C.photoRed);
}

// 21. Ethical And Health Certifications
function slide21(s) {
  photo(s, 0.708, 0, 4.854, 4.854, C.photoRed);
  T(s, 'W   W   W   .   B  A  D  O  G  A   .   C   O   M', { x: -0.502, y: 3.676, w: 2.205, h: 0.151, rotate: 270, flipH: true, align: 'center', fontSize: 9, color: C.paper });
  chrome(s, C.green, 'Page 21 / 30', 'Chapter 09');
  T(s, 'ETHICAL AND HEALTH CERTIFICATIONS', { x: 6.952, y: 1.318, w: 4.992, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, 'Regulatory Compliance and Certifications', { x: 6.952, y: 0.977, w: 4.992, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'We explore the potential certifications and endorsements that align with our commitment to ethical practices, health consciousness, and sustainability.', { x: 1.301, y: 5.849, w: 3.161, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, WWW, { x: 6.952, y: 3.978, w: 2.205, h: 0.135, flipH: true, fontSize: 8 });
  T(s, TAGLINE, { x: 6.952, y: 5.942, w: 2.953, h: 0.471, flipH: true, fontSize: 14, bold: true });
  photo(s, 10.501, 3.371, 2.832, 4.129, C.photoBlue);
}

// 28. Conclusion And Future Outlook
function slide28(s) {
  chrome(s, C.yellow, 'Page 28 / 30', 'Chapter 13');
  T(s, 'Future Growth and Innovation', { x: 8.702, y: 5.399, w: 3.6, h: 0.236, fontSize: 14, bold: true, color: C.ink85 });
  T(s, 'We conclude by outlining the future growth strategies and ongoing innovation initiatives that will keep "Food for the Future" at the forefront of the evolving food industry.', { x: 8.702, y: 5.734, w: 3.6, h: 0.656, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, color: C.ink85 });
  T(s, 'This chapter summarizes the key components of the business proposal, emphasizing the innovative approach, sustainability focus, and potential impact of "Food for the Future."', { x: 1.739, y: 2.891, w: 3.898, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, color: C.ink85 });
  T(s, 'Summary of the Proposal', { x: 1.739, y: 2.56, w: 3.898, h: 0.236, fontSize: 14, bold: true, color: C.ink85 });
  T(s, 'Text Title Here', { x: 6.538, y: 2.953, w: 1.132, h: 0.202, align: 'right', fontSize: 12, bold: true, color: C.ink85 });
  T(s, WWW, { x: 10.137, y: 1.213, w: 2.402, h: 0.135, align: 'right', fontSize: 8 });
  T(s, 'CONCLUSION AND FUTURE OUTLOOK', { x: 1.739, y: 0.742, w: 4.331, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  photo(s, 1.739, 4.29, 5.931, 3.21, C.photoBlue);
}

// 30. Thank You For Watching
function slide30(s) {
  chrome(s, C.green, 'Page 30 / 30', 'Food For The Future');
  T(s, [
      { text: 'Objectively integrate emerging competencies before to process communities great dramatically holistic innovation rather than client-centric data ', options: { fontSize: 9 } },
      { text: 'envisioned based expertise and cross media growth ', options: { fontSize: 9, color: C.ink75 } },
      { text: 'impact of "Food for the Future."', options: { fontSize: 9, color: C.ink85 } },
    ], { x: 8.159, y: 5.228, w: 4.409, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 'THANK YOU\nFOR WATCHING', { x: 8.159, y: 1.615, w: 4.409, h: 1.346, valign: 'middle', fontSize: 40, fontFace: F.display });
  T(s, WWW, { x: 0.437, y: 2.65, w: 2.205, h: 0.135, rotate: 270, flipH: true, align: 'right', fontSize: 8 });
  T(s, TAGLINE3, { x: 1.472, y: 5.178, w: 1.964, h: 0.707, flipH: true, fontSize: 14, bold: true });
  photo(s, 4.201, 1.615, 3.194, 4.269, C.photoBlue);
}

// 8. The Concept Of "Food For The Future"
function slide08(s) {
  chrome(s, C.orange, 'Page 08 / 30', 'Chapter 03');
  // three feature columns
  [
    ['devices', 'Sustainable Sourcing', 1.412,
     'This section outlines our commitment to sustainable sourcing practices, including responsible agricultural methods, ethical supply chains, and reduced environmental impact.'],
    ['envelope', 'Nutrient-Rich Products', 5.601,
     'We introduce our focus on developing and offering innovative, nutrient-rich food products that cater to changing consumer preferences for healthier and more sustainable dietary choices.'],
    ['screen', 'Technology Integration', 9.791,
     'Embraces cutting-edge technologies such as precision agriculture, artificial intelligence, and advanced food processing techniques to enhance efficiency, and reduce waste.'],
  ].forEach(([mark, title, x, body]) => {
    icon(s, mark, x, 1.02, 0.5, 0.37, C.ink85);
    T(s, title, { x, y: 1.755, w: 2.835, h: 0.303, bold: true });
    T(s, body, { x, y: 2.294, w: 2.835, h: 0.884, fontSize: 9, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  T(s, 'THE CONCEPT OF\n"FOOD FOR THE FUTURE"', { x: 6.412, y: 4.2, w: 6.214, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, WWW, { x: 9.975, y: 6.391, w: 2.651, h: 0.135, align: 'right', fontSize: 8 });
  T(s, TAGLINE, { x: 6.412, y: 6.054, w: 3.088, h: 0.471, flipH: true, fontSize: 14, bold: true });
  photo(s, 0.708, 4.2, 4.996, 3.3, C.photoBlue);
}

// 15. Environmental Impact
function slide15(s) {
  S(s, 'rect', { x: 5.851, y: 0, w: 7.483, h: 7.5, fill: { color: C.orange } });
  chrome(s, C.green, 'Page 15 / 30', 'Chapter 07');
  T(s, WWW, { x: 11.502, y: 3.683, w: 2.402, h: 0.135, rotate: 270, align: 'center', fontSize: 8 });
  T(s, 'We detail our commitment to minimizing the environmental footprint of our operations, from sustainable sourcing to eco-friendly packaging and waste reduction initiatives.', { x: 1.35, y: 3.282, w: 3.277, h: 0.658, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'ENVIRONMENTAL\nIMPACT', { x: 1.35, y: 1.523, w: 3.858, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, 'Sustainability and Ethical Practices', { x: 1.35, y: 1.259, w: 3.858, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, TAGLINE, { x: 1.35, y: 5.77, w: 3.425, h: 0.471, flipH: true, fontSize: 14, bold: true });
  S(s, 'line', { x: 5.851, y: 0, w: 0, h: 7.5, line: { color: C.ink, width: 1.5 } });
  // 2 x 2 photo grid, alternating brand tones
  [[7.112, 1.258, C.photoBlue], [9.711, 1.258, C.photoRed],
   [7.112, 3.88, C.photoRed], [9.711, 3.88, C.photoBlue]]
    .forEach(([x, y, color]) => photo(s, x, y, 2.362, 2.362, color));
}

// 18. Marketing Channels
function slide18(s) {
  chrome(s, C.green, 'Page 18 / 30', 'Chapter 08');
  T(s, 'MARKETING CHANNELS', { x: 1.304, y: 1.035, w: 5.964, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Marketing and Branding Strategy', { x: 1.304, y: 0.768, w: 5.964, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'Your Subtitle Here', { x: 1.304, y: 2.522, w: 2.638, h: 0.236, fontSize: 14, bold: true });
  T(s, 'This section outlines our multichannel marketing approach, leveraging digital platforms, social media, and strategic partnerships to reach and engage our target audience.', { x: 1.304, y: 2.99, w: 2.638, h: 0.884, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, WWW, { x: 10.035, y: 1.137, w: 2.402, h: 0.135, align: 'right', fontSize: 8 });
  T(s, 'W   W   W   .   N   E   D   O   L   A   .   C   O   M', { x: 1.304, y: 6.598, w: 2.402, h: 0.135, fontSize: 8 });
  // two laptop mock-ups, each with a caption block beneath
  [[4.067, 4.537, C.photoRed, 4.546, 0.458, C.orange, 'briefcase'],
   [8.808, 9.279, C.photoBlue, 9.255, 0.507, C.yellow, 'screen']]
    .forEach(([shellX, screenX, screenColor, iconX, iconW, iconColor, mark]) => {
      laptop(s, shellX, 2.409, 4.525, 2.583, screenX, 2.522, 3.568, 2.252, screenColor);
      icon(s, mark, iconX, 5.76, iconW, 0.42, iconColor);
      T(s, 'Text Title Here', { x: iconX + 0.908, y: 5.76, w: 1.132, h: 0.202, fontSize: 12, bold: true });
      T(s, 'Empowered markets dynamically procrastinate B2C users after. Interactively a commerce makes within process centric outside making.',
        { x: iconX + 0.908, y: 6.075, w: 2.659, h: 0.658, align: 'justify', valign: 'middle', lineSpacingMultiple: 1.5, fontSize: 9 });
    });
}

// 19. Pricing Table
function slide19(s) {
  chrome(s, C.yellow, 'Page 19 / 30', 'Chapter 08');
  T(s, 'PRICING TABLE', { x: 1.951, y: 1.026, w: 10.138, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Marketing and Branding Strategy', { x: 1.951, y: 0.726, w: 10.138, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'Collaboratively administrate empowered markets plug play networks dynamic procrastinate users. After installed base benefits dramatic. for multimedia based expertise and cross-media into books good growth strategies. Seamlessly visualize on quality. After installed base benefits dramatic.', { x: 1.951, y: 1.814, w: 10.138, h: 0.43, lineSpacingMultiple: 1.5, fontSize: 9 });
  // four price cards joined by short connectors
  const CARDS = [
    ['Basic Price', '$ 25 / Pack ', 1.951, C.green],
    ['Medium Price', '$ 50 / Pack', 4.557, C.orange],
    ['Super Price', '$ 80 / Pack', 7.145, C.green],
    ['Special Price', '$ 120 / Pack', 9.757, C.orange],
  ];
  CARDS.forEach(([label, price, x, color], i) => {
    if (i > 0) {
      const prev = CARDS[i - 1];
      S(s, 'line', { x: prev[2] + 2.333, y: 4.14, w: x - prev[2] - 2.333, h: 0, line: { color: i === 2 ? C.green : C.orange, width: 4.5 } });
    }
    S(s, 'rect', { x, y: 2.97, w: 2.333, h: 2.341, fill: { color } });
    S(s, 'rect', { x: x + 0.143, y: 3.117, w: 2.047, h: 2.047, fill: { color: C.paper } });
    T(s, label, { x: x + 0.143, y: 3.473, w: 2.047, h: 0.202, align: 'center', fontSize: 12, bold: true });
    T(s, price, { x: x + 0.143, y: 3.827, w: 2.047, h: 0.404, align: 'center', fontSize: 24, bold: true });
    T(s, 'Leverage agile framework provide. ', { x: x + 0.454, y: 4.378, w: 1.426, h: 0.43, align: 'center', lineSpacingMultiple: 1.5, fontSize: 9 });
  });
  T(s, WWW, { x: 9.885, y: 6.338, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  T(s, TAGLINE, { x: 1.951, y: 6.17, w: 3.425, h: 0.471, flipH: true, fontSize: 14, bold: true });
}

// 22. Phase-Wise Rollout
function slide22(s) {
  chrome(s, C.yellow, 'Page 22 / 30', 'Chapter 10');
  T(s, 'PHASE-WISE ROLLOUT', { x: 1.801, y: 1.075, w: 5.964, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Implementation Plan and Timeline', { x: 1.801, y: 0.808, w: 5.964, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'This chapter outlines the phased implementation plan for "Food for the Future," including key milestones, timelines, and resource allocation for each phase of the business launch.', { x: 8.735, y: 0.916, w: 3.506, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  // staggered phase bars
  [[2.165, 2.489, 4.855], [9.448, 2.489, 2.428], [4.593, 3.164, 4.855],
   [2.165, 3.838, 2.428], [7.02, 3.838, 4.855]].forEach(([x, y, w]) => {
    S(s, 'rect', { x, y, w, h: 0.445, fill: { color: C.orange } });
    T(s, 'Your Title Here', { x, y: y + 0.121, w, h: 0.202, align: 'center', fontSize: 12, bold: true, color: C.paper });
  });
  // milestone columns
  [[1.801, '12 - MAR - 2025', 'Milestone One'], [3.933, '26 - MAY - 2025', 'Milestone Two'],
   [6.066, '03 - AUG - 2025', 'Milestone Three'], [8.199, '17 - OCT - 2025', 'Milestone Four'],
   [10.332, '01 - DEC - 2025', 'Milestone Five']].forEach(([x, date, title]) => {
    S(s, 'rect', { x, y: 4.651, w: 1.909, h: 0.6, fill: { color: C.green }, line: { color: C.green, width: 2.25 } });
    T(s, date, { x: x + 0.171, y: 4.833, w: 1.565, h: 0.236, align: 'center', fontSize: 14, bold: true, color: C.paper });
    S(s, 'line', { x, y: 5.086, w: 0, h: 1.606, line: { color: C.green, width: 2.25 } });
    T(s, title, { x: x + 0.176, y: 5.619, w: 1.732, h: 0.269, fontSize: 16, bold: true, color: C.ink65 });
    T(s, 'Leverage agile frameworks with provide robust synopsis for high level views new normal.', { x: x + 0.176, y: 5.961, w: 1.732, h: 0.656, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9, color: C.grey50 });
  });
}

// 23. Collaboration And Partnerships
function slide23(s) {
  chrome(s, C.orange, 'Page 23 / 30', 'Chapter 10');
  T(s, 'COLLABORATION AND PARTNERSHIPS', { x: 1.313, y: 0.992, w: 5.964, h: 1.212, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Implementation Plan and Timeline', { x: 1.313, y: 0.723, w: 5.964, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'We discuss potential collaborations with research institutions, agricultural organizations, and technology partners to strengthen our position in the market and foster innovation.', { x: 9.135, y: 1.135, w: 3.49, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  // five partner logos, framed and sized in an arc towards the centre
  [['Advantric', '72%', 1.313, 3.208, 1.791, 1.344, C.green, 4.946, C.photoBlue],
   ['Crowd Outreach', '86%', 3.276, 3.036, 2.25, 1.689, C.yellow, 5.081, C.photoRed],
   ['Client Bright', '93%', 5.699, 2.927, 2.54, 1.906, C.orange, 5.227, C.photoBlue],
   ['Fundraiseiq', '88%', 8.412, 3.036, 2.25, 1.689, C.yellow, 5.081, C.photoRed],
   ['Charity Spirit', '77%', 10.835, 3.208, 1.791, 1.344, C.green, 4.946, C.photoBlue],
  ].forEach(([name, pct, x, y, w, h, frame, ty, tone]) => {
    S(s, 'rect', { x, y, w, h, fill: { color: frame } });
    photo(s, x + w * 0.11, y + h * 0.145, w * 0.78, h * 0.705, tone);
    const tx = x + w / 2 - 0.827;
    T(s, name, { x: tx, y: ty, w: 1.654, h: 0.202, align: 'center', fontSize: 12, bold: true });
    T(s, 'Section outlines to multichannel marketing approach, leveraging good digital platforms.', { x: tx, y: ty + 0.245, w: 1.654, h: 0.584, align: 'center', lineSpacingMultiple: 1.5, fontSize: 8 });
    T(s, pct, { x: tx, y: ty + 1.041, w: 1.654, h: 0.303, align: 'center', bold: true });
    T(s, 'Potential Success', { x: tx, y: ty + 1.365, w: 1.654, h: 0.185, align: 'center', fontSize: 11 });
  });
}

// 24. Financial Forecast
function slide24(s) {
  chrome(s, C.green, 'Page 24 / 30', 'Chapter 11');
  T(s, 'FINANCIAL FORECAST', { x: 1.472, y: 1.225, w: 5.964, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Financial Projections and Funding', { x: 1.472, y: 0.958, w: 5.964, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, WWW, { x: 10.365, y: 1.327, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  T(s, 'This chapter presents detailed financial projections, including revenue forecasts, expense estimates, and profitability analysis over the short and long term.', { x: 1.472, y: 2.789, w: 3.394, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, '2025 Income', { x: 1.472, y: 4.364, w: 1.813, h: 0.237, fontSize: 14, bold: true });
  T(s, '+ $500.000', { x: 3.053, y: 4.364, w: 1.813, h: 0.237, align: 'right', fontSize: 14, bold: true });
  T(s, 'Leverage agile frameworks to provide a robust synopsis to high level for views new normal views new normal.', { x: 1.472, y: 4.683, w: 3.394, h: 0.43, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, TAGLINE, { x: 1.472, y: 6.071, w: 3.394, h: 0.471, flipH: true, fontSize: 14, bold: true });
  T(s, '93%', { x: 5.63, y: 6.088, w: 0.772, h: 0.438, flipH: true, fontSize: 26, bold: true });
  T(s, 'Revenue in\nthe last year', { x: 6.755, y: 6.122, w: 0.991, h: 0.37, flipH: true, fontSize: 11, bold: true });
  T(s, 'Interactively coordinate proactive centric outside into the box thinking pursue scalable empowered market. ', { x: 8.15, y: 6.092, w: 4.42, h: 0.43, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });

  // hand-drawn revenue column chart: baseline y=4.905, 200 units = 0.2915 in
  const BASE_Y = 4.905, UNIT = 0.2915 / 200, BAR_W = 0.296, PITCH = 0.8922, X0 = 6.333;
  S(s, 'line', { x: 6.084, y: 2.859, w: 0, h: 2.087, line: { color: C.grey75, width: 0.25 } });
  S(s, 'line', { x: 6.04, y: BASE_Y, w: 6.529, h: 0, line: { color: C.grey75, width: 0.25 } });
  ['1,400', '1,200', '1,000', '800', '600', '400', '200', '0'].forEach((label, i) => {
    T(s, label, { x: 5.63, y: 2.789 + i * 0.2915, w: 0.338, h: 0.135, align: 'right', fontSize: 8 });
  });
  [['2019', 759, 954], ['2020', 1093, 1297], ['2021', 1020, 897], ['2022', 603, 458],
   ['2023', 1020, 238], ['2024', 1203, 804], ['2025', 579, 1092],
  ].forEach(([year, left, right], i) => {
    const x = X0 + i * PITCH;
    T(s, year, { x: x + 0.064, y: 4.979, w: 0.506, h: 0.135, align: 'center', fontSize: 8 });
    [[x, left, C.orange], [x + 0.338, right, C.yellow]].forEach(([bx, value, color]) => {
      const h = value * UNIT;
      S(s, 'rect', { x: bx, y: BASE_Y - h, w: BAR_W, h, fill: { color } });
      T(s, String(value), { x: bx, y: BASE_Y - h - 0.205, w: BAR_W, h: 0.135, align: 'center', fontSize: 8 });
    });
  });
}

// 25. Funding Requirements
function slide25(s) {
  chrome(s, C.yellow, 'Page 25 / 30', 'Chapter 11');
  T(s, 'FUNDING REQUIREMENTS', { x: 1.95, y: 0.934, w: 4.331, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, 'Financial Projections and Funding', { x: 1.95, y: 0.669, w: 4.331, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'We outline the funding requirements for the implementation of "Food for the Future," considering initial setup costs, research and development, marketing expenses, and working capital needs.', { x: 8.538, y: 1.012, w: 3.553, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, WWW, { x: 1.95, y: 6.696, w: 2.205, h: 0.135, flipH: true, align: 'right', fontSize: 8 });
  // two income rows, each led by an "info" badge
  [['2024 Income', '+ $415.000', 2.68, C.orange], ['2025 Income', '+ $500.000', 4.178, C.green]]
    .forEach(([label, amount, y, color]) => {
      S(s, 'ellipse', { x: 1.95, y: y + 0.062, w: 0.678, h: 0.678, fill: { color } });
      S(s, 'donut', { x: 2.06, y: y + 0.172, w: 0.458, h: 0.458, fill: { color: C.paper } });
      T(s, label, { x: 3.007, y, w: 1.132, h: 0.236, fontSize: 14, bold: true });
      T(s, amount, { x: 4.186, y, w: 1.813, h: 0.237, align: 'right', fontSize: 14, bold: true });
      T(s, 'Leverage agile frameworks to provide a robust synopsis for high level for views new normal views new normal on quality after installed base benefits dramatic.', { x: 3.007, y: y + 0.32, w: 2.992, h: 0.657, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
    });
  // price / quality matrix
  S(s, 'line', { x: 8.089, y: 2.68, w: 0, h: 3.77, line: { color: C.grey75, width: 1 } });
  S(s, 'line', { x: 8.089, y: 6.45, w: 3.77, h: 0, line: { color: C.grey75, width: 1 } });
  [[8.442, 3.04, C.green], [10.045, 3.04, C.orange],
   [8.442, 4.629, C.orange], [10.045, 4.629, C.green]].forEach(([x, y, color]) => {
    S(s, 'rect', { x, y, w: 1.461, h: 1.461, fill: { color } });
    T(s, 'Your Title\nName Here', { x: x + 0.217, y: y + 0.528, w: 1.028, h: 0.37, align: 'center', fontSize: 11, bold: true });
  });
  [['High', 7.523, 2.68, 0.475], ['Price', 7.523, 4.663, 0.475], ['Low', 7.523, 6.646, 0.475],
   ['Quality', 9.727, 6.646, 0.634], ['High', 11.616, 6.646, 0.475]].forEach(([label, x, y, w]) => {
    T(s, label, { x, y, w, h: 0.185, align: 'center', fontSize: 11, bold: true });
  });
}

// 26. Risk Analysis
function slide26(s) {
  chrome(s, C.orange, 'Page 26 / 30', 'Chapter 12');
  T(s, 'RISK ANALYSIS', { x: 6.584, y: 1.287, w: 5.964, h: 0.606, align: 'right', valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Risks and Contingency Plans', { x: 6.584, y: 1.02, w: 5.964, h: 0.269, align: 'right', valign: 'middle', fontSize: 16 });
  T(s, 'A comprehensive risk analysis is provided, identifying potential challenges and uncertainties that may impact the success of "Food for the Future."', { x: 1.493, y: 1.242, w: 4.281, h: 0.43, align: 'justify', lineSpacingMultiple: 1.5, fontSize: 9 });
  [2.909, 3.287, 3.669].forEach((y, i) => {
    T(s, 'Subtitle For Overview', { x: i === 2 ? 1.502 : 1.493, y, w: 1.575, h: 0.202, align: 'justify', fontSize: 12, bold: true });
  });
  T(s, 'Building brands\nand establish\nyour identity', { x: 1.493, y: 5.446, w: 1.182, h: 0.454, fontSize: 9 });
  T(s, '6M', { x: 2.438, y: 5.403, w: 0.63, h: 0.539, fontSize: 32, bold: true });
  T(s, 'Markets via plug and play for networks users good.', { x: 1.493, y: 6.054, w: 1.535, h: 0.43, lineSpacingMultiple: 1.5, fontSize: 9 });
  // numbered call-outs down the right edge
  [[2.909, 'Credibly pontificate highly efficient manufactured products and enabled data efficiently unleash.', '01'],
   [4.364, 'Proactively envisioned multimedia based expert and cross media to growth strategies quality.', '02'],
   [5.827, 'Credibly pontificate highly efficient manufactured products and enabled data efficiently unleash.', '03'],
  ].forEach(([y, body, num]) => {
    T(s, body, { x: 9.893, y, w: 1.97, h: 0.657, align: 'right', lineSpacingMultiple: 1.5, fontSize: 9 });
    T(s, num, { x: 11.949, y: y + 0.132, w: 0.599, h: 0.471, flipH: true, align: 'right', valign: 'middle', fontSize: 28, bold: true });
  });
  // interlocking gears illustration
  [[4.032, 3.184, 2.184, 2.142, C.orange], [7.008, 3.444, 1.809, 1.774, C.green],
   [5.774, 4.557, 1.674, 1.641, C.yellow]].forEach(([x, y, w, h, color]) => gear(s, x, y, w, h, color));
}

// 27. Contingency Plans
function slide27(s) {
  chrome(s, C.green, 'Page 27 / 30', 'Chapter 12');
  T(s, 'CONTINGENCY PLANS', { x: 1.602, y: 0.95, w: 5.964, h: 0.606, valign: 'middle', fontSize: 36, fontFace: F.display });
  T(s, 'Risks and Contingency Plans', { x: 1.602, y: 0.683, w: 5.964, h: 0.269, valign: 'middle', fontSize: 16 });
  T(s, 'We present contingency plans and risk mitigation strategies to address unforeseen challenges, ensuring adaptability and resilience in the face of changing market dynamics.', { x: 8.614, y: 0.791, w: 3.825, h: 0.658, align: 'right', lineSpacingMultiple: 1.5, fontSize: 9 });
  T(s, 'Our Contingency Plans\nHeading Text', { x: 1.602, y: 2.241, w: 2.559, h: 0.551, fontSize: 14, bold: true });
  T(s, 'Our Risk Mitigation Strategies Heading Text', { x: 9.477, y: 2.241, w: 2.962, h: 0.471, align: 'right', fontSize: 14, bold: true });
  // mirrored 01/02/03 lists either side of the illustration
  const BODY = 'Credibly pontificate highly without efficient manufactured to products and enabled data efficiently unleash customers.';
  [[2.204, 1.602, 'left'], [9.477, 11.84, 'right']].forEach(([bodyX, numX, side]) => {
    ['01', '02', '03'].forEach((num, i) => {
      const y = 3.432 + i * 1.2835;
      T(s, BODY, { x: bodyX, y, w: 2.359, h: 0.657, align: side, lineSpacingMultiple: 1.5, fontSize: 9 });
      T(s, num, { x: numX, y: y + 0.13, w: 0.599, h: 0.471, flipH: true, align: side, valign: 'middle', fontSize: 28, bold: true });
    });
  });
  person(s);
}

// 29. Our Detail Contact And About Us
function slide29(s) {
  chrome(s, C.orange, 'Page 29 / 30', 'Food For The Future');
  T(s, 'OUR DETAIL CONTACT AND ABOUT US', { x: 1.541, y: 3.147, w: 4.882, h: 1.077, valign: 'middle', fontSize: 32, fontFace: F.display });
  T(s, [
    { text: 'Collaboration provide a robust to visualize quality intellectual capital without superior  and idea sharing. Proactively envisioned multimedia based expertise and cross media growth ', options: { color: C.ink75 } },
    { text: 'impact of "Food for the Future."', options: { color: C.ink85 } },
  ], { x: 1.541, y: 5.641, w: 4.183, h: 0.658, lineSpacingMultiple: 1.5, fontSize: 9 });
  // contact strip along the top
  [['Email', 'food_marketing@susana.com', 'envelope', 1.541, 0.301, 0.287, C.yellow, 2.095, 1.811],
   ['Address', '22 West , Arizona 1910, United States', 'building', 4.734, 0.301, 0.271, C.green, 5.288, 2.283],
   ['Phone', '+1927 0294 1920', 'phone', 8.399, 0.181, 0.271, C.yellow, 8.953, 1.102],
   ['Website', 'www.susana.com', 'screen', 10.883, 0.333, 0.301, C.green, 11.437, 1.063],
  ].forEach(([label, value, mark, ix, iw, ih, color, tx, tw]) => {
    icon(s, mark, ix, 1.338, iw, ih, color);
    T(s, label, { x: tx, y: 1.338, w: tw, h: 0.236, fontSize: 14, bold: true, color: C.ink75 });
    T(s, value, { x: tx, y: 1.641, w: tw, h: 0.168, fontSize: 10, bold: true, color: C.ink75 });
  });
  monitor(s, 7.257, 3.147, 5.243, 4.347, 7.423, 3.351, 4.91, 2.942, C.photoRed);
}

/* ------------------------------------------------------------ icon sheets */
// Slides 31-42 of the original are library pages: a 10-column grid of small
// solid glyphs in tx2-85%.  Reproduced with preset silhouettes on the same grid.
const ICON_COL_X = [1.32, 2.50, 3.68, 4.86, 6.04, 7.22, 8.40, 9.58, 10.76, 11.94];
const ICON_SIZE = 0.38;
const ICON_GLYPHS = [
  'ellipse', 'rect', 'triangle', 'diamond', 'pentagon', 'hexagon', 'heptagon', 'octagon',
  'star4', 'star5', 'star6', 'star7', 'star8', 'heart', 'sun', 'moon', 'cloud', 'plus',
  'mathPlus', 'donut', 'pie', 'pieWedge', 'blockArc', 'arc', 'teardrop', 'lightningBolt',
  'noSmoking', 'can', 'cube', 'bevel', 'plaque', 'frame', 'halfFrame', 'corner',
  'trapezoid', 'parallelogram', 'chevron', 'homePlate', 'wave', 'snip2DiagRect',
];

// [first row centre, row pitch, row count, indices used on the final row]
const ICON_SHEETS = [
  [0.899, 0.949, 7, null], [0.879, 0.956, 7, null], [0.882, 0.952, 7, null],
  [0.893, 0.950, 7, null], [0.904, 0.948, 7, null], [0.880, 0.953, 7, null],
  [0.892, 0.952, 7, null], [0.905, 0.952, 7, null], [0.879, 0.814, 8, null],
  [0.904, 0.948, 7, null], [0.883, 0.952, 7, null],
  [0.904, 0.957, 7, [0, 1, 2, 3, 4, 8, 9]],
];

function iconSheet(slide, sheetIndex) {
  const [y0, pitch, rows, lastRow] = ICON_SHEETS[sheetIndex];
  for (let r = 0; r < rows; r++) {
    const cols = r === rows - 1 && lastRow ? lastRow : ICON_COL_X.map((_, i) => i);
    cols.forEach((c) => {
      const glyph = ICON_GLYPHS[(sheetIndex * 17 + r * 13 + c * 3) % ICON_GLYPHS.length];
      S(slide, glyph, {
        x: ICON_COL_X[c] - ICON_SIZE / 2, y: y0 + r * pitch - ICON_SIZE / 2,
        w: ICON_SIZE, h: ICON_SIZE, fill: { color: C.ink85 },
      });
    });
  }
}

/* -------------------------------------------------------------- assemble */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27,
  slide28, slide29, slide30,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'SUSANA', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'SUSANA';
  pptx.author = 'Susana';
  pptx.title = 'Food for the Future';

  BUILDERS.forEach((builder) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.paper };
    builder(slide);
  });

  ICON_SHEETS.forEach((_, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.paper };
    iconSheet(slide, i);
  });

  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '004c9d90-862b-4339-8990-c1d6263098d8_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
