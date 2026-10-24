/**
 * "Workshop." presentation template — rebuilt with pptxgenjs.
 * Run: node 162e2d46-c3a0-40d3-a3e5-f789cf666e65_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette + typography (theme accent2 = ED7D31 with lum tints)
 * ------------------------------------------------------------------ */
const C = {
  orange: 'F9761A', // brand orange (slide backgrounds)
  cream: 'FFF6F0', // light backgrounds
  dark: 'C55A11', // accent2 lumMod 75%
  brown: '843C0B', // accent2 lumMod 50%  (headline colour on cream)
  peach1: 'FBE5D6', // accent2 20/80  (headline colour on orange)
  peach2: 'F8CBAD', // accent2 40/60
  peach3: 'F4B183', // accent2 60/40
  white: 'FFFFFF',
  ring: 'B4C7E7', // thin decorative outline on the big circle (slide 7)
  hatch: 'F5F7FB', // "Replace Image" placeholder body
  hatchLine: 'D3DEF1',
};

const FONT = {
  serif: 'Lora',
  sans: 'Raleway',
  sansMed: 'Raleway Medium',
  sansSemi: 'Raleway SemiBold',
  ui: 'Calibri', // placeholder prompt text
};

const BULLET = { characterCode: '2022', indent: 13.5 };
const NO_LINE = { type: 'none' };

/* ------------------------------------------------------------------ *
 * Helpers
 * ------------------------------------------------------------------ */

/** Text box. PowerPoint anchors text boxes at the top, pptxgenjs centres them. */
function text(slide, body, opts) {
  slide.addText(body, Object.assign({ valign: 'top', fontFace: FONT.sans }, opts));
}

/** Heading in Lora bold. */
function heading(slide, body, opts) {
  text(slide, body, Object.assign({ fontFace: FONT.serif, bold: true }, opts));
}

/** Paragraph of body copy (Raleway 12pt, 150% leading). */
function para(slide, body, opts) {
  text(slide, body, Object.assign({ fontSize: 12, lineSpacingMultiple: 1.5 }, opts));
}

/** Bulleted single-line entry used by the agenda/feature lists. */
function bulletItem(slide, body, opts) {
  para(slide, body, Object.assign({ h: 0.368, bullet: BULLET }, opts));
}

/**
 * Rectangle whose corners may be rounded, given as fractions of w/h:
 *   { tl:[fx,fy], tr:[fx,fy], br:[fx,fy], bl:[fx,fy] }
 * Returns a pptxgenjs custGeom point list (inches, relative to the shape).
 */
function roundedRectPoints(w, h, radii) {
  const K = 0.5523; // cubic-bezier circle approximation
  const r = (corner) => {
    const f = radii[corner];
    return f ? [f[0] * w, f[1] * h] : [0, 0];
  };
  const [tlx, tly] = r('tl');
  const [trx, tryy] = r('tr');
  const [brx, bry] = r('br');
  const [blx, bly] = r('bl');
  const pts = [{ x: tlx, y: 0, moveTo: true }, { x: w - trx, y: 0 }];
  if (trx || tryy) {
    pts.push({ x: w, y: tryy, curve: { type: 'cubic', x1: w - trx * (1 - K), y1: 0, x2: w, y2: tryy * (1 - K) } });
  }
  pts.push({ x: w, y: h - bry });
  if (brx || bry) {
    pts.push({ x: w - brx, y: h, curve: { type: 'cubic', x1: w, y1: h - bry * (1 - K), x2: w - brx * (1 - K), y2: h } });
  }
  pts.push({ x: blx, y: h });
  if (blx || bly) {
    pts.push({ x: 0, y: h - bly, curve: { type: 'cubic', x1: blx * (1 - K), y1: h, x2: 0, y2: h - bly * (1 - K) } });
  }
  pts.push({ x: 0, y: tly });
  if (tlx || tly) {
    pts.push({ x: tlx, y: 0, curve: { type: 'cubic', x1: 0, y1: tly * (1 - K), x2: tlx * (1 - K), y2: 0 } });
  }
  pts.push({ close: true });
  return pts;
}

/** Decorative freeform panel: `o` = {x,y,w,h,fill,rotate,flipH,flipV}, `radii` per corner. */
function panel(slide, o, radii) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rotate: o.rotate, flipH: o.flipH, flipV: o.flipV,
    fill: { color: o.fill }, line: NO_LINE,
    points: roundedRectPoints(o.w, o.h, radii),
  });
}

/* Corner recipes reused across the deck (fractions of the shape box). */
const CORNER = {
  bigSquare: { tl: [0.1966, 0.1966] }, // slide 1 tiles
  wideBanner: { tl: [0.1445, 0.2787] }, // slides 2/3/5 banners
  softPanel: { br: [0.1944, 0.1746] }, // slides 2/4 big panels
  card: { tl: [0.0578, 0.1193] }, // slide 6 stat cards
  cardBottom: { bl: [0.0578, 0.1193] },
  cardLeft: { tl: [0.0524, 0.1081], bl: [0.0524, 0.1081] },
  pill: { tl: [0.2256, 0.5], bl: [0.2256, 0.5] }, // slide 3 tab
  slabTop: { tl: [0.1072, 0.1966] }, // slide 1 wide slab
};

/**
 * Stand-in for a photo. The reference deck uses "picture placeholder" shapes filled
 * with a 5% diagonal pattern, so the hatch is drawn as a fan of thin "/" lines.
 */
const HATCH_STEP = 0.139;
const ICON_W = 0.874;
const ICON_H = 0.68;

function imagePlaceholder(slide, x, y, w, h, opts) {
  const o = opts || {};
  slide.addShape('rect', { x, y, w, h, fill: { color: C.white }, line: NO_LINE, shadow: o.shadow });
  for (let c = HATCH_STEP; c < w + h; c += HATCH_STEP) {
    const x1 = Math.min(c, w); // start on the top (or right) edge
    const y1 = c - x1;
    const y2 = Math.min(c, h); // end on the left (or bottom) edge
    const x2 = c - y2;
    slide.addShape('line', {
      x: x + x2, y: y + y1, w: x1 - x2, h: y2 - y1, flipV: true,
      line: { color: C.hatchLine, width: 0.5 },
    });
  }
  slide.addText('Replace Image', {
    x, y: y + 0.05, w, h: 0.3, align: 'center', valign: 'top', wrap: false,
    margin: 0, fontFace: FONT.ui, fontSize: o.promptSize || 18, color: '000000',
  });
  // the small "picture" glyph PowerPoint centres inside an empty placeholder
  const ix = x + (w - ICON_W) / 2;
  const iy = y + (h - ICON_H) / 2;
  slide.addShape('rect', { x: ix, y: iy, w: ICON_W, h: ICON_H, fill: { color: 'FAFAFA' }, line: { color: '545452', width: 1.25 } });
  slide.addShape('rect', { x: ix + 0.06, y: iy + 0.05, w: ICON_W - 0.12, h: ICON_H - 0.1, fill: { color: 'FAFAFA' }, line: { color: 'CECECE', width: 0.75 } });
  slide.addShape('ellipse', { x: ix + 0.143, y: iy + 0.114, w: 0.153, h: 0.153, fill: { color: 'F8DB8F' }, line: { color: 'ED9B33', width: 1 } });
  slide.addShape('triangle', { x: ix + 0.30, y: iy + 0.26, w: 0.50, h: 0.36, fill: { color: '83BEEC' }, line: { color: '2E75B6', width: 0.5 } });
  slide.addShape('triangle', { x: ix + 0.10, y: iy + 0.40, w: 0.44, h: 0.22, fill: { color: '83BEEC' }, line: { color: '2E75B6', width: 0.5 } });
}

/**
 * Chevron arrow drawn as custGeom so the notch depth matches the original
 * (adjust value 34788/100000 of the short side, vs. pptxgenjs' fixed 50%).
 */
function chevron(slide, o) {
  const notch = 0.34788 * Math.min(o.w, o.h);
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate,
    fill: { color: o.fill }, line: NO_LINE,
    points: [
      { x: 0, y: 0, moveTo: true },
      { x: o.w - notch, y: 0 },
      { x: o.w, y: o.h / 2 },
      { x: o.w - notch, y: o.h },
      { x: 0, y: o.h },
      { x: notch, y: o.h / 2 },
      { close: true },
    ],
  });
}

/** Horizontal progress bar: grey track + coloured value + % caption. */
function skillBar(slide, bar) {
  slide.addShape('roundRect', {
    x: bar.x, y: bar.y, w: bar.w, h: bar.h,
    fill: { color: bar.trackColor }, line: NO_LINE, rectRadius: bar.h / 2,
  });
  slide.addShape('roundRect', {
    x: bar.valueX !== undefined ? bar.valueX : bar.x, y: bar.valueY !== undefined ? bar.valueY : bar.y,
    w: bar.valueW, h: bar.valueH !== undefined ? bar.valueH : bar.h,
    fill: { color: bar.valueColor }, line: NO_LINE,
    rectRadius: (bar.valueH !== undefined ? bar.valueH : bar.h) / 2,
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  panel(s, { x: 8.441, y: 5.249, w: 2.445, h: 2.266, fill: C.dark }, CORNER.bigSquare);
  panel(s, { x: 2.019, y: 4.63, w: 3.114, h: 2.886, fill: C.dark, flipH: true }, CORNER.bigSquare);
  panel(s, { x: 10.886, y: 2.882, w: 2.445, h: 2.266, fill: C.dark }, CORNER.bigSquare);
  panel(s, { x: 0, y: 4.63, w: 5.708, h: 2.886, fill: C.dark, flipH: true }, CORNER.slabTop);

  heading(s, 'Workshop.', { x: 1.224, y: 2.158, w: 5.991, h: 1.447, fontSize: 80, color: C.peach1, wrap: false });
  text(s, 'Presentation Template', {
    x: 1.224, y: 1.632, w: 2.285, h: 0.337, fontFace: FONT.serif, italic: true,
    fontSize: 14, color: C.peach1, wrap: false,
  });
  s.addShape('line', { x: 1.312, y: 5.382, w: 0.635, h: 0, line: { color: C.white, width: 1 } });
  text(s, [
    { text: 'Lorem ipsum dolor zasit amet consectetur coluore', options: { breakLine: true } },
    { text: 'adipiscing elit sed eiusmod tempor incididunt labore et dolore magna aliqua utam enim.' },
  ], {
    x: 1.231, y: 5.553, w: 3.717, h: 0.903, fontFace: FONT.serif, italic: true,
    fontSize: 11, color: C.peach1, lineSpacingMultiple: 1.5,
  });
  text(s, 'December, 2021', {
    x: 11.118, y: 0.757, w: 1.573, h: 0.337, fontFace: FONT.serif, italic: true,
    fontSize: 14, color: C.peach1, wrap: false,
  });

  imagePlaceholder(s, 9.333, 3.765, 4.0, 3.75);
}

function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  panel(s, { x: 9.886, y: 3.636, w: 3.447, h: 1.875, fill: C.peach3 }, CORNER.wideBanner);
  panel(s, { x: 0, y: 0, w: 5.055, h: 5.286, fill: C.orange }, CORNER.softPanel);

  heading(s, 'Welcome to Studio', { x: 5.922, y: 0.818, w: 6.514, h: 0.909, fontSize: 48, color: C.brown });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing, and get ready for a successful and productive team meeting dolore is cool and lorem ipsum is best element.',
    { x: 5.922, y: 1.869, w: 5.974, h: 0.974, color: C.brown });

  // white-on-orange column
  text(s, 'Interactive Workshops', { x: 0.695, y: 0.675, w: 3.647, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.peach1 });
  para(s, 'Use this template to run interactive and workshops remotely. This is a collection of our favorite workshop activities that you can run digitally, alongside a base of essential materials and (including the all-important virtual sticky-note!) lorem ipsum.',
    { x: 0.695, y: 0.992, w: 3.698, h: 1.883, color: C.peach1 });
  text(s, 'Description Text', { x: 0.695, y: 3.192, w: 3.752, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing, and get ready for a successful and productive team meeting dolore is cool.',
    { x: 0.7, y: 3.475, w: 3.695, h: 1.277, color: C.peach1 });

  panel(s, { x: 6.262, y: 5.809, w: 3.109, h: 1.691, fill: C.peach2 }, CORNER.wideBanner);

  text(s, 'Lorem ipsum dolore', { x: 0.677, y: 5.921, w: 3.876, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.brown });
  para(s, 'Join using the live video collaboration feature getting those creative juices flowing, and get ready for.',
    { x: 0.677, y: 6.225, w: 4.378, h: 0.671, color: C.brown });

  imagePlaceholder(s, 6.821, 4.125, 6.512, 3.375);
}

function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  panel(s, { x: 10.514, y: 2.447, w: 2.82, h: 1.381, fill: C.peach3 }, CORNER.wideBanner);

  heading(s, 'Brief Information About Our Company', { x: 0.832, y: 0.836, w: 7.343, h: 1.582, fontSize: 44, color: C.peach1 });
  text(s, 'Interactive Workshops', { x: 0.832, y: 2.765, w: 3.647, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.peach1 });
  para(s, 'Use this template to run interactive and workshops remotely. This is a collection of our favorite workshop activities that you can run digitally, alongside a base of essential.',
    { x: 0.832, y: 3.081, w: 6.828, h: 0.671, color: C.peach1, italic: true });
  s.addShape('line', { x: 0.929, y: 4.018, w: 6.375, h: 0, line: { color: C.peach1, width: 1 } });
  text(s, 'Category Text', { x: 0.832, y: 4.283, w: 3.752, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing, and get ready for a successful and productive team meeting dolore is cool.',
    { x: 0.837, y: 4.566, w: 6.823, h: 0.671, color: C.peach1 });

  const list = [
    ['Join using the live video collaboration.', 0.832, 5.386, 3.418],
    ['Join using the live video.', 0.832, 5.732, 3.186],
    ['Join using the live video collab.', 0.832, 6.101, 3.186],
    ['Join using the live video.', 4.698, 5.386, 2.963],
    ['Join using the live video collab.', 4.698, 5.732, 2.963],
    ['Join using the live.', 4.698, 6.101, 2.963],
  ];
  list.forEach(([label, x, y, w]) => bulletItem(s, label, { x, y, w, color: C.peach1 }));

  panel(s, { x: 7.741, y: 6.351, w: 1.583, h: 0.715, fill: C.peach3, rotate: 90 }, CORNER.pill);
}

function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  panel(s, { x: 7.295, y: 1.187, w: 6.038, h: 6.313, fill: C.peach1, flipH: true, flipV: true }, CORNER.softPanel);
  panel(s, { x: 8.786, y: 2.745, w: 4.548, h: 4.755, fill: C.orange, flipH: true, flipV: true }, CORNER.softPanel);

  heading(s, 'Agenda List', { x: 1.15, y: 1.176, w: 5.742, h: 1.111, fontSize: 60, color: C.brown });

  const agenda = [
    ['Welcome', 1.15, 2.659, 1.9],
    ['About Us', 1.15, 3.098, 1.471],
    ['History of Company', 1.15, 3.535, 2.275],
    ['Our Problem', 1.15, 3.972, 1.92],
    ['Our Solution For Company', 1.15, 4.371, 2.618],
    ['Our Vision & Mission', 1.15, 4.808, 2.441],
    ['Our Portofoliio', 1.15, 5.245, 2.441],
    ['Our Big Team', 4.351, 2.659, 1.471],
    ['Our Core Team', 4.351, 3.096, 1.9],
    ['Data Statistic', 4.351, 3.533, 1.471],
    ['Infographics Process', 4.351, 3.97, 2.169],
    ['Thank You', 4.351, 4.407, 1.471],
  ];
  agenda.forEach(([label, x, y, w]) => bulletItem(s, label, { x, y, w, color: C.brown }));

  imagePlaceholder(s, 7.326, 2.273, 3.886, 5.227);
}

function slide05(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  panel(s, { x: 9.662, y: 0.937, w: 3.732, h: 1.895, fill: C.dark, rotate: 90, flipH: true }, CORNER.wideBanner);

  heading(s, 'About Us', { x: 0.977, y: 1.515, w: 4.755, h: 1.313, fontSize: 72, color: C.peach1 });
  para(s, [
    { text: 'Join using the live video collaboration feature, get those creative juices flowing, and get ready for a successful and lol', options: { breakLine: true } },
    { text: 'Lorem ipsum dolore colore.' },
  ], { x: 0.977, y: 3.165, w: 3.414, h: 1.277, color: C.peach1 });

  s.addShape('roundRect', {
    x: 1.081, y: 4.982, w: 1.442, h: 0.443, rectRadius: 0.2215,
    fill: { color: C.peach1 }, line: NO_LINE,
  });
  s.addText('Learn More', {
    x: 1.081, y: 4.982, w: 1.442, h: 0.443, align: 'center', valign: 'middle',
    fontFace: FONT.serif, fontSize: 12, color: C.orange,
  });

  text(s, 'Lorem ipsum dolore', { x: 6.667, y: 5.921, w: 3.876, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature getting those creative juices flowing, and get ready for vocabulary .',
    { x: 6.667, y: 6.275, w: 5.033, h: 0.671, color: C.peach1 });

  imagePlaceholder(s, 6.667, 0, 5.033, 5.425);
}

function slide06(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  panel(s, { x: 5.84, y: 5.509, w: 2.938, h: 1.988, fill: C.peach3 }, CORNER.card);
  panel(s, { x: 5.595, y: 0.783, w: 3.992, h: 2.426, fill: C.peach3, rotate: 270 }, CORNER.card);
  panel(s, { x: 5.169, y: 5.512, w: 2.938, h: 1.988, fill: C.peach3 }, CORNER.card);
  panel(s, { x: 9.342, y: 0, w: 3.992, h: 2.426, fill: C.peach1 }, CORNER.cardBottom);
  panel(s, { x: 9.342, y: 2.537, w: 3.992, h: 2.426, fill: C.peach2 }, CORNER.cardLeft);
  panel(s, { x: 9.342, y: 5.074, w: 3.992, h: 2.426, fill: C.peach3 }, CORNER.card);

  heading(s, 'History of Business', { x: 0.797, y: 1.006, w: 5.103, h: 2.524, fontSize: 72, color: C.brown });
  text(s, 'Interactive Workshops', { x: 0.797, y: 4.455, w: 2.719, h: 0.337, fontFace: FONT.sansMed, fontSize: 14, color: C.brown });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.',
    { x: 0.797, y: 4.806, w: 3.203, h: 0.974, color: C.brown });

  // three stacked stat cards on the right
  const stats = [
    { y: 0.404, value: '+500 K', note: 'Join using the live video collab.' },
    { y: 2.891, value: '$900 M', note: 'Join using the live video.' },
    { y: 5.428, value: '$69 B', note: 'Join using the live collaboration.' },
  ];
  stats.forEach((st) => {
    text(s, 'Lorem ipsum dolore', { x: 10.036, y: st.y, w: 2.533, h: 0.37, fontFace: FONT.sansSemi, bold: true, fontSize: 16, color: C.brown });
    heading(s, st.value, { x: 10.036, y: st.y + 0.478, w: 2.533, h: 0.774, fontSize: 40, color: C.brown });
    para(s, st.note, { x: 10.036, y: st.y + 1.317, w: 2.76, h: 0.37, color: C.brown });
  });

  s.addShape('line', { x: 0, y: 3.933, w: 6.378, h: 0, line: { color: C.peach3, width: 0.5 } });
  // tall photograph -> picture placeholder, with the original's soft left-cast shadow
  imagePlaceholder(s, 5.758, 1.62, 3.261, 5.88, {
    shadow: { type: 'outer', blur: 100, offset: 20, angle: 360, color: C.brown, opacity: 0.4 },
  });
}

function slide07(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  s.addShape('ellipse', { x: 5.899, y: 2.59, w: 4.09, h: 4.09, fill: { color: C.peach1 }, line: { color: C.ring, width: 1 } });
  s.addShape('ellipse', { x: 9.603, y: 0.821, w: 2.417, h: 2.417, fill: { color: C.peach1 }, line: NO_LINE });

  heading(s, 'Our Problem', { x: 1.053, y: 1.006, w: 5.103, h: 2.524, fontSize: 72, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.',
    { x: 1.053, y: 3.75, w: 3.203, h: 0.974, color: C.peach1 });

  // big circle
  heading(s, '$900 M', { x: 6.971, y: 4.212, w: 1.947, h: 0.64, fontSize: 32, color: C.orange, align: 'center' });
  para(s, 'Join using the live video', { x: 6.884, y: 4.848, w: 2.132, h: 0.368, color: C.orange, align: 'center' });
  s.addShape('ellipse', { x: 6.065, y: 2.771, w: 3.771, h: 3.771, fill: { type: 'none' }, line: { color: C.orange, width: 10 } });

  // lower-right circle
  s.addShape('ellipse', { x: 9.353, y: 3.41, w: 3.269, h: 3.269, fill: { color: C.peach1 }, line: { color: C.orange, width: 10 } });
  heading(s, '50%', { x: 10.198, y: 4.588, w: 1.579, h: 0.64, fontSize: 32, color: C.orange, align: 'center' });
  para(s, 'Join using the live', { x: 10.13, y: 5.227, w: 1.729, h: 0.368, color: C.orange, align: 'center' });
  s.addShape('ellipse', { x: 9.603, y: 3.669, w: 2.776, h: 2.776, fill: { type: 'none' }, line: { color: C.orange, width: 10 } });

  // upper-right circle
  s.addShape('ellipse', { x: 9.793, y: 1.006, w: 2.035, h: 2.035, fill: { type: 'none' }, line: { color: C.orange, width: 10 } });
  heading(s, '99 K', { x: 10.021, y: 1.591, w: 1.579, h: 0.64, fontSize: 32, color: C.orange, align: 'center' });
  para(s, 'Join using', { x: 9.946, y: 2.14, w: 1.729, h: 0.368, color: C.orange, align: 'center' });
}

function slide08(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  imagePlaceholder(s, 4.152, 1.07, 4.12, 6.43);

  heading(s, 'Our Solution', { x: 0.919, y: 0.842, w: 5.103, h: 2.524, fontSize: 72, color: C.brown });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.',
    { x: 0.919, y: 3.586, w: 3.203, h: 0.974, color: C.brown });

  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna.';
  const features = [
    { num: '01/', numW: 0.691, title: 'Interactive Workshops', y: 1.07 },
    { num: '02/', numW: 0.726, title: 'Collaboration', y: 3.093 },
    { num: '03/', numW: 0.735, title: 'Creative Juices', y: 5.085 },
  ];
  features.forEach((f, i) => {
    heading(s, f.num, { x: i === 2 ? 8.358 : 8.354, y: f.y, w: f.numW, h: 0.438, fontSize: 20, color: C.brown, wrap: false });
    heading(s, f.title, { x: 9.097, y: f.y, w: 2.897, h: 0.37, fontSize: 16, color: C.brown });
    para(s, LOREM, { x: 9.097, y: f.y + 0.37, w: 3.381, h: 0.974, color: C.brown });
  });

  const cards = [
    { x: 1.002, textX: 1.384, fill: C.peach3 },
    { x: 4.293, textX: 4.687, fill: C.peach2 },
  ];
  cards.forEach((card) => {
    s.addShape('roundRect', {
      x: card.x, y: 5.085, w: 2.792, h: 1.866, rectRadius: 0.171,
      fill: { color: card.fill }, line: NO_LINE,
    });
    heading(s, '$900 M', { x: card.textX, y: 5.503, w: 1.329, h: 0.404, fontSize: 18, color: C.brown });
    para(s, [
      { text: 'Join using the live video', options: { breakLine: true } },
      { text: 'collaboration feature.' },
    ], { x: card.textX, y: 5.907, w: 2.055, h: 0.671, color: C.brown });
  });
}

function slide09(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  heading(s, 'Our Vision In The Company For The Next 5 Years',
    { x: 0.967, y: 4.9, w: 8.459, h: 1.582, fontSize: 44, color: C.peach1 });

  heading(s, 'Interactive Workshops', { x: 5.852, y: 1.174, w: 2.897, h: 0.37, fontSize: 16, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing, and get ready for a successful and lorem.',
    { x: 5.852, y: 1.574, w: 5.381, h: 0.671, color: C.peach1 });

  const links = [
    ['Welcome', 5.852, 2.506, 1.9],
    ['About Us', 5.852, 2.945, 1.471],
    ['History of Company', 5.852, 3.382, 2.275],
    ['Our Big Team', 8.402, 2.499, 1.471],
    ['Our Core Team', 8.402, 2.936, 1.9],
    ['Data Statistic', 8.402, 3.373, 1.471],
  ];
  links.forEach(([label, x, y, w]) => bulletItem(s, label, { x, y, w, color: C.peach1 }));

  s.addShape('line', { x: 9.962, y: 4.971, w: 0, h: 1.597, line: { color: C.peach1, width: 1 } });
  heading(s, 'Notes', { x: 10.206, y: 5.063, w: 1.116, h: 0.368, fontSize: 16, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature get those creative juices.',
    { x: 10.206, y: 5.451, w: 2.275, h: 0.974, color: C.peach1 });

  imagePlaceholder(s, 1.186, 0, 3.826, 4.269);
}

function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  heading(s, 'Our Mission', { x: 1.476, y: 0.925, w: 3.399, h: 1.582, fontSize: 44, color: C.brown });

  const intro = [
    { title: 'Description Text', x: 4.919, bodyX: 4.916, bodyY: 1.364, body: 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.' },
    { title: 'Tittle Text', x: 8.725, bodyX: 8.725, bodyY: 1.36, body: 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for lorem.' },
  ];
  intro.forEach((col) => {
    heading(s, col.title, { x: col.x, y: 0.977, w: 2.117, h: 0.37, fontSize: 16, color: C.brown });
    para(s, col.body, { x: col.bodyX, y: col.bodyY, w: 3.12, h: 0.974, color: C.brown });
  });

  const MISSION_BODY = 'Join using the live video collaborat ion feature those creative lorem ipsum dol.';
  const columns = [
    { picX: 1.476, titleX: 1.434, bodyX: 1.421, title: 'Mission First' },
    { picX: 5.494, titleX: 5.444, bodyX: 5.431, title: 'Mission Second' },
    { picX: 9.512, titleX: 9.426, bodyX: 9.413, title: 'Mission Third' },
  ];
  columns.forEach((col) => {
    imagePlaceholder(s, col.picX, 3.037, 2.345, 1.866);
    heading(s, col.title, { x: col.titleX, y: 5.218, w: 1.952, h: 0.383, fontSize: 16, color: C.brown });
    para(s, MISSION_BODY, { x: col.bodyX, y: 5.601, w: 2.472, h: 0.974, color: C.brown });
  });
}

function slide11(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  heading(s, 'Our Big Team', { x: 3.646, y: 0.886, w: 6.042, h: 1.111, fontSize: 60, color: C.peach1, align: 'center' });

  const QUOTE = '\u201cLorem ipsum dolor sit amet la consectetur adipiscing eiusmod tempor incid.\u201d';
  const people = [
    { name: 'Lucas Anderson', nameX: 1.405, roleX: 2.119, quoteX: 1.405, quoteY: 3.499, roleY: 3.171, align: 'right' },
    { name: 'Castiliion Andrew', nameX: 9.437, roleX: 9.437, quoteX: 9.437, quoteY: 3.507, roleY: 3.163, align: 'left' },
  ];
  people.forEach((p) => {
    heading(s, p.name, { x: p.nameX, y: 2.689, w: 2.492, h: 0.404, fontSize: 18, color: C.peach1, align: p.align });
    text(s, 'Ceo & Founder', { x: p.roleX, y: p.roleY, w: 1.778, h: 0.337, fontFace: FONT.serif, fontSize: 14, color: C.peach1, align: p.align });
    para(s, QUOTE, { x: p.quoteX, y: p.quoteY, w: 2.492, h: 0.974, color: C.peach1, align: p.align });
  });

  imagePlaceholder(s, 4.202, 2.514, 2.102, 2.046);
  imagePlaceholder(s, 7.03, 2.514, 2.102, 2.046);

  // skill bars: dark track, light value fill (left column fills from the right)
  const rows = [
    { label: 'Stength', labelY: 4.985, trackY: 5.127, leftPct: '90%', rightPct: '70%',
      left: { valueX: 2.906, valueY: 5.12, valueW: 2.814, valueH: 0.109 },
      right: { h: 0.111, valueY: 5.12, valueW: 2.267, valueH: 0.125, valueColor: C.peach2 } },
    { label: 'Defance', labelY: 5.572, trackY: 5.699, leftPct: '70%', rightPct: '80%',
      left: { valueX: 3.464, valueY: 5.692, valueW: 2.256, valueH: 0.109 },
      right: { y: 5.692, h: 0.115, valueY: 5.692, valueW: 2.569, valueH: 0.115 } },
    { label: 'Healthy', labelY: 6.234, trackY: 6.361, leftPct: '80%', rightPct: '78%',
      left: { valueX: 3.085, valueY: 6.361, valueW: 2.635, valueH: 0.103 },
      right: { valueW: 2.43 } },
  ];
  const pctY = [5.011, 5.578, 6.254];
  const pctYRight = [5.011, 5.597, 6.254];
  rows.forEach((row, i) => {
    skillBar(s, Object.assign({
      x: 2.513, y: row.trackY, w: 3.207, h: 0.103, trackColor: C.brown, valueColor: C.peach1,
    }, row.left));
    heading(s, row.label, { x: 6.044, y: row.labelY, w: 1.246, h: 0.37, fontSize: 16, color: C.peach1, align: 'center' });
    skillBar(s, Object.assign({
      x: 7.614, y: row.trackY, w: 3.207, h: 0.103, trackColor: C.brown, valueColor: C.peach1,
    }, row.right));
    heading(s, row.leftPct, { x: 1.729, y: pctY[i], w: 0.766, h: 0.303, fontSize: 12, color: C.peach1, align: 'center' });
    heading(s, row.rightPct, { x: 10.82, y: pctYRight[i], w: 0.766, h: 0.303, fontSize: 12, color: C.peach1, align: 'center' });
  });
}

function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  heading(s, 'Our Core Team', { x: 1.097, y: 0.886, w: 4.064, h: 2.121, fontSize: 60, color: C.brown });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.',
    { x: 1.097, y: 3.263, w: 3.203, h: 0.974, color: C.brown });

  const QUOTE = '\u201cLorem ipsum dolor sit amet la consectetur adipiscing eiusmod tempor incid.\u201d';
  const people = [
    { name: 'Jacky Cmiw', tileY: 0.803, picY: 1.047, x: 8.398, nameY: 1.039, roleY: 1.48, quoteY: 1.825 },
    { name: 'Sterly Kascek', tileY: 2.777, picY: 2.992, x: 8.398, nameY: 3.051, roleY: 3.492, quoteY: 3.837 },
    { name: 'Dany Woods', tileY: 4.718, picY: 4.937, x: 8.403, nameY: 4.955, roleY: 5.395, quoteY: 5.74 },
  ];
  people.forEach((p) => {
    s.addShape('roundRect', {
      x: 6.278, y: p.tileY, w: 1.014, h: 1.014, rectRadius: 0.114,
      fill: { color: C.peach1 }, line: NO_LINE,
    });
  });
  people.forEach((p) => {
    heading(s, p.name, { x: p.x, y: p.nameY, w: 2.492, h: 0.404, fontSize: 18, color: C.brown });
    text(s, 'Ceo & Founder', { x: p.x, y: p.roleY, w: 1.778, h: 0.337, fontFace: FONT.serif, fontSize: 14, color: C.brown });
    para(s, QUOTE, { x: p.x, y: p.quoteY, w: 3.57, h: 0.671, color: C.brown });
  });

  const rows = [
    { label: 'Stength', labelY: 4.782, trackY: 4.887, trackH: 0.117, valueW: 1.975, valueH: 0.125, pct: '70%', pctX: 4.984, pctW: 0.557, pctY: 4.795 },
    { label: 'Defance', labelY: 5.457, trackY: 5.578, trackH: 0.115, valueY: 5.575, valueW: 2.192, valueH: 0.117, pct: '80%', pctX: 4.979, pctW: 0.692, pctY: 5.483 },
    { label: 'Healthy', labelY: 6.148, trackY: 6.248, trackH: 0.103, valueW: 1.886, valueH: 0.103, pct: '78%', pctX: 4.979, pctW: 0.557, pctY: 6.154 },
  ];
  rows.forEach((row) => {
    heading(s, row.label, { x: 1.101, y: row.labelY, w: 1.246, h: 0.303, fontSize: 12, color: C.brown });
    skillBar(s, {
      x: 2.078, y: row.trackY, w: 2.779, h: row.trackH, trackColor: C.peach2,
      valueY: row.valueY, valueW: row.valueW, valueH: row.valueH, valueColor: C.brown,
    });
    heading(s, row.pct, { x: row.pctX, y: row.pctY, w: row.pctW, h: 0.303, fontSize: 12, color: C.brown });
  });

  // this layout overrides the placeholder prompt to 24pt
  people.forEach((p) => imagePlaceholder(s, 6.527, p.picY, 1.557, 1.516, { promptSize: 24 }));
}

function slide13(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  s.addShape('rect', { x: 6.667, y: 0, w: 6.667, h: 7.5, fill: { color: C.peach1 }, line: NO_LINE });

  heading(s, 'Data Statistics', { x: 1.128, y: 0.932, w: 4.064, h: 2.121, fontSize: 60, color: C.peach1 });
  para(s, 'Join using the live video collaboration feature, get those creative juices flowing and geting ready for.',
    { x: 1.128, y: 3.212, w: 3.203, h: 0.974, color: C.peach1 });

  const CATS = ['Category 1', 'Category 2', 'Category 3'];
  const chartData = [
    { name: 'Series 1', labels: CATS, values: [3.3, 2.5, 3.5] },
    { name: 'Series 2', labels: CATS, values: [2.4, 4.4, 1.8] },
    { name: 'Series 3', labels: CATS, values: [2, 2, 3] },
  ];
  s.addChart('bar', chartData, {
    x: 7.527, y: 0.963, w: 5.067, h: 5.575,
    barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: 113,
    chartColors: [C.orange, C.peach3, C.peach2],
    showLegend: false, showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: 'General',
    dataLabelColor: C.brown, dataLabelFontFace: FONT.serif, dataLabelFontSize: 12,
    valAxisHidden: true, valGridLine: { style: 'none' },
    catAxisLabelColor: C.brown, catAxisLabelFontFace: FONT.serif, catAxisLabelFontSize: 12,
    catAxisLineColor: 'D9D9D9',
    catGridLine: { color: C.brown, size: 0.75, style: 'solid' },
  });

  const notes = [
    { num: '01/', numW: 0.542, numY: 4.499, bodyY: 4.425, body: 'Join using the live video collaboration feature get those creative juices colore.' },
    { num: '02/', numW: 0.568, numY: 5.266, bodyY: 5.191, body: 'Join using the live video collaboration feature get those creative flowing and geting.' },
    { num: '03/', numW: 0.574, numY: 6.015, bodyY: 5.953, body: 'Join using the live video collaboration lorem ipsum dolore is amet consectet.' },
  ];
  notes.forEach((n) => {
    heading(s, n.num, { x: 1.128, y: n.numY, w: n.numW, h: 0.337, fontSize: 14, color: C.peach1, wrap: false });
    para(s, n.body, { x: 1.67, y: n.bodyY, w: 3.73, h: 0.671, color: C.peach1 });
  });
}

function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  heading(s, 'Infographics Process', { x: 2.081, y: 0.903, w: 9.172, h: 1.111, fontSize: 60, color: C.brown, align: 'center' });

  const BODY = 'Join using the live video collaborat ion feature those creative lorem ipsum.';
  const steps = [
    { title: 'Process First', chevronX: 1.134, chevronY: 2.793, ovalX: 1.493, ovalY: 3.006, iconX: 1.926, iconY: 3.434, titleX: 1.296, titleY: 5.312, bodyX: 1.036, bodyY: 5.618 },
    { title: 'Process Second', chevronX: 4.075, chevronY: 2.793, ovalX: 4.434, ovalY: 3.006, iconX: 4.867, iconY: 3.434, titleX: 4.237, titleY: 5.312, bodyX: 3.977, bodyY: 5.618 },
    { title: 'Process Third', chevronX: 7.037, chevronY: 2.797, ovalX: 7.396, ovalY: 3.01, iconX: 7.829, iconY: 3.443, titleX: 7.199, titleY: 5.317, bodyX: 6.939, bodyY: 5.622 },
    { title: 'Process Fouth', chevronX: 9.978, chevronY: 2.793, ovalX: 10.337, ovalY: 3.006, iconX: 10.77, iconY: 3.434, titleX: 10.14, titleY: 5.312, bodyX: 9.88, bodyY: 5.618 },
  ];
  steps.forEach((step) => {
    chevron(s, { x: step.chevronX, y: step.chevronY, w: 2.215, h: 1.984, rotate: 90, fill: C.brown });
    s.addShape('ellipse', {
      x: step.ovalX, y: step.ovalY, w: 1.496, h: 1.496, rotate: 90,
      fill: { color: C.cream }, line: { color: C.brown, width: 0.5 },
    });
    // line-art icon graphic -> simple brown ring of similar visual weight
    s.addShape('donut', {
      x: step.iconX, y: step.iconY, w: 0.632, h: 0.632,
      fill: { color: C.brown }, line: NO_LINE,
    });
    heading(s, step.title, { x: step.titleX, y: step.titleY, w: 1.952, h: 0.383, fontSize: 16, color: C.brown, align: 'center' });
    para(s, BODY, { x: step.bodyX, y: step.bodyY, w: 2.472, h: 0.974, color: C.brown, align: 'center' });
  });
}

function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  s.addShape('ellipse', { x: 1.262, y: 2.216, w: 3.067, h: 3.067, fill: { color: C.peach1 }, line: NO_LINE });
  heading(s, 'Thank You', {
    x: 2.065, y: 2.732, w: 8.743, h: 2.036, fontSize: 115, color: C.brown,
    align: 'center', charSpacing: 2, wrap: false,
  });
  s.addShape('roundRect', {
    x: 9.78, y: 4.712, w: 0.847, h: 0.141, rectRadius: 0.0705,
    fill: { color: C.brown }, line: NO_LINE,
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE_13x7.5', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_13x7.5';
  pres.theme = { headFontFace: FONT.serif, bodyFontFace: FONT.sans };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15]
    .forEach((fn) => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '162e2d46-c3a0-40d3-a3e5-f789cf666e65_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
