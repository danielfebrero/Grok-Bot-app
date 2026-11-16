/**
 * SafeWork — workplace safety deck (18 slides, 13.333in x 7.5in).
 * Recreated with pptxgenjs only. Photographic content in the original is
 * replaced by labelled placeholder rectangles.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  orange: 'FFBB33', // accent1
  deep: 'FF6106', // accent2
  amberDark: 'E59900',
  amberLite: 'FFD685',
  star: 'FFC000',
  ink: '262626',
  dark: '404040', // tx1 @ 75% lum
  gray: '595959',
  mid: '808080',
  white: 'FFFFFF',
  black: '000000',
  page: 'F7F7F7', // master background
  grid: 'F2F2F2',
  ph: 'D9D9D9' // image placeholder
};

const MAJOR = 'Public Sans';
const MINOR = 'Work Sans';

// The deck uses three recurring outer shadows.
const GLOW = { type: 'outer', color: C.orange, opacity: 0.3, blur: 58, offset: 24, angle: 90 };
const SOFT = { type: 'outer', color: C.black, opacity: 0.1, blur: 54, offset: 13, angle: 135 };
const WIDE = { type: 'outer', color: C.black, opacity: 0.03, blur: 29, offset: 18, angle: 60 };
const LIFT = { type: 'outer', color: C.black, opacity: 0.07, blur: 58, offset: 24, angle: 90 };
const HALO = { type: 'outer', color: C.orange, opacity: 0.07, blur: 41, offset: 0, angle: 90 };

/* ---------------------------------------------------------------- copy deck */

const BODY_FULL =
  'SafeWork focuses on building a culture of safety where every worker understands, practices, and values safe procedures. ';
const BODY_SHORT = 'SafeWork focuses on building a culture of safety where every worker understands, practices.';
const BODY_MED = 'SafeWork focuses on building a culture of safety where every worker understands.';
const BODY_TINY = 'SafeWork focuses on building a culture.';
const PROMO = 'Promoting safety-first mindsets and reliable procedures to ensure secure';

/* ------------------------------------------------------------- primitives */

function txt(slide, content, o) {
  slide.addText(content, Object.assign({ fontFace: MINOR, fontSize: 18, color: C.black, valign: 'top' }, o));
}

/** Eyebrow line above every section title. */
function eyebrow(slide, s, x, y, w, color) {
  txt(slide, s, { x: x, y: y, w: w || 2.831, h: 0.37, fontSize: 16, bold: true, color: color || C.orange });
}

/** Big two-tone section headline (one paragraph per line). */
function headline(slide, lines, o) {
  txt(
    slide,
    lines.map(function (t) { return { text: t, options: { breakLine: true } }; }),
    Object.assign({ fontFace: MAJOR, fontSize: 40, bold: true, color: C.dark }, o)
  );
}

/** Grey body paragraph at 150% leading. */
function body(slide, s, o) {
  txt(slide, s, Object.assign({ fontSize: 12, color: C.dark, lineSpacingMultiple: 1.5 }, o));
}

function rect(slide, o) {
  slide.addShape('rect', o);
}

function roundRect(slide, x, y, w, h, radius, o) {
  slide.addShape('roundRect', Object.assign({ x: x, y: y, w: w, h: h, rectRadius: radius }, o));
}

function ellipse(slide, x, y, w, h, o) {
  slide.addShape('ellipse', Object.assign({ x: x, y: y, w: w, h: h }, o));
}

function hline(slide, x, y, w, color, width) {
  slide.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: width } });
}

/** Rounded bar of length `len` centred on (cx, cy) and rotated `ang` degrees clockwise. */
function bar(slide, cx, cy, len, thick, ang, color) {
  slide.addShape('roundRect', {
    x: cx - len / 2, y: cy - thick / 2, w: len, h: thick,
    rectRadius: thick / 2, rotate: ang, fill: { color: color }
  });
}

/** End point of such a bar, measured from its centre. */
function barEnd(cx, cy, len, ang, sign) {
  const r = (ang * Math.PI) / 180;
  return { x: cx + sign * (len / 2) * Math.cos(r), y: cy + sign * (len / 2) * Math.sin(r) };
}

/* ------------------------------------------------------- design components */

/**
 * "SafeWork" lockup: an orange disc crossed by a diagonal wedge, with a
 * contrasting notch at the wedge's tip, followed by the wordmark.
 * `tone` is 'dark' (on light pages) or 'light' (on the orange cover).
 */
function logo(slide, tone) {
  const behind = tone === 'light' ? C.orange : C.page; // page colour showing through the cut
  const notch = tone === 'light' ? C.white : C.orange;
  const word = tone === 'light' ? C.white : C.dark;
  ellipse(slide, 0.498, 0.317, 0.147, 0.147, { fill: { color: C.deep } });
  slide.addShape('rtTriangle', { x: 0.484, y: 0.358, w: 0.121, h: 0.121, fill: { color: C.deep } });
  bar(slide, 0.5445, 0.4185, 0.26, 0.014, 45, behind); // diagonal cut between disc and wedge
  slide.addShape('rtTriangle', { x: 0.484, y: 0.435, w: 0.043, h: 0.043, fill: { color: notch } });
  txt(slide, 'SafeWork', {
    x: 0.678, y: 0.286, w: 1.35, h: 0.24,
    margin: 0, fontFace: MAJOR, fontSize: 12.5, bold: true, color: word, valign: 'middle'
  });
}

/** Master furniture: logo, right-hand strapline and page number. */
function chrome(slide, opts) {
  const o = opts || {};
  logo(slide, 'dark');
  if (o.footer !== false) footer(slide);
  if (o.number !== false) pageNumber(slide, o.number);
}

function footer(slide) {
  txt(slide, 'Safety Work Presentation', {
    x: 9.587, y: 6.951, w: 3.351, h: 0.302,
    fontSize: 9, color: C.orange, align: 'right', lineSpacingMultiple: 1.5
  });
}

function pageNumber(slide, n) {
  txt(slide, String(n), {
    x: 0.418, y: 6.908, w: 0.682, h: 0.399,
    fontSize: 12, bold: true, color: C.orange, valign: 'middle'
  });
}

/** Number + caption pair, centred on a common axis (used with and without a disc). */
function statLabels(slide, x, y, value, label, color, o) {
  const s = o || {};
  const width = s.w || 2.307;
  txt(slide, value, {
    x: x, y: y, w: width, h: 0.505,
    fontFace: MAJOR, fontSize: 24, bold: !!s.bold, color: color, align: 'center', valign: 'middle'
  });
  txt(slide, label, {
    x: x + (width - 1.166) / 2, y: y + 0.43, w: 1.166, h: 0.236,
    fontFace: MAJOR, fontSize: 8, color: s.labelColor || color, align: 'center', valign: 'middle'
  });
}

/** 1.31" disc badge with a number and caption underneath. */
function statCircle(slide, x, y, value, label, o) {
  const s = o || {};
  ellipse(slide, x, y, 1.31, 1.31, { fill: { color: s.fill || C.orange }, shadow: s.shadow || GLOW });
  statLabels(slide, x - 0.499, y + 0.335, value, label, s.color || C.white, s);
}

/** Rounded icon chip. */
function iconTile(slide, x, y, size, fill, o) {
  roundRect(slide, x, y, size, size, size * 0.29167, Object.assign({ fill: fill }, o || {}));
}

/**
 * Simplified pictograms drawn from native shapes — the original deck uses
 * vector freeforms of the same silhouettes.
 */
function glyph(slide, kind, x, y, w, h, color, behind) {
  const fill = { color: color };
  const bg = { color: behind || C.page }; // colour of whatever sits under the glyph
  const cx = x + w / 2;
  const cy = y + h / 2;
  switch (kind) {
    case 'globe': { // snow globe: dome with a tree and snowflakes, on a plinth
      slide.addShape('chord', { x: x + w * 0.1, y: y + h * 0.02, w: w * 0.8, h: h * 0.78, fill: fill, angleRange: [180, 0] });
      ellipse(slide, x + w * 0.2, y + h * 0.3, w * 0.12, w * 0.12, { fill: bg });
      ellipse(slide, x + w * 0.52, y + h * 0.13, w * 0.14, w * 0.14, { fill: bg });
      ellipse(slide, x + w * 0.66, y + h * 0.36, w * 0.11, w * 0.11, { fill: bg });
      slide.addShape('triangle', { x: x + w * 0.32, y: y + h * 0.25, w: w * 0.26, h: h * 0.3, fill: bg });
      rect(slide, { x: x + w * 0.42, y: y + h * 0.5, w: w * 0.06, h: h * 0.16, fill: bg });
      slide.addShape('trapezoid', { x: x + w * 0.2, y: y + h * 0.68, w: w * 0.6, h: h * 0.14, fill: fill });
      roundRect(slide, x, y + h * 0.84, w, h * 0.16, h * 0.08, { fill: fill });
      break;
    }
    case 'hammers': { // two crossed hammers, heads splayed at the top
      const L = h * 0.86;
      bar(slide, cx, cy, L, w * 0.14, 45, color);
      bar(slide, cx, cy, L, w * 0.14, -45, color);
      const l = barEnd(cx, cy, L, 45, -1);
      const r = barEnd(cx, cy, L, -45, -1);
      bar(slide, l.x, l.y, w * 0.44, h * 0.24, -45, color);
      bar(slide, r.x, r.y, w * 0.44, h * 0.24, 45, color);
      break;
    }
    case 'tools': { // wrench crossed with a screwdriver
      const L = h * 0.9;
      bar(slide, cx, cy, L, w * 0.13, 45, color);
      bar(slide, cx, cy, L, w * 0.13, -45, color);
      const ring = barEnd(cx, cy, L, -45, -1);
      ellipse(slide, ring.x - w * 0.16, ring.y - w * 0.16, w * 0.32, w * 0.32, { fill: fill });
      ellipse(slide, ring.x - w * 0.09, ring.y - w * 0.09, w * 0.18, w * 0.18, { fill: bg });
      const tip = barEnd(cx, cy, L, 45, -1);
      bar(slide, tip.x, tip.y, w * 0.26, h * 0.2, 45, color);
      break;
    }
    case 'backpack':
      roundRect(slide, x, y + h * 0.18, w, h * 0.82, w * 0.28, { fill: fill });
      slide.addShape('chord', { x: x + w * 0.22, y: y, w: w * 0.56, h: h * 0.4, fill: fill, angleRange: [180, 0] });
      roundRect(slide, x + w * 0.24, y + h * 0.5, w * 0.52, h * 0.34, w * 0.08, { fill: bg });
      break;
    case 'clipboard':
      roundRect(slide, x, y + h * 0.08, w, h * 0.92, w * 0.16, { fill: fill });
      roundRect(slide, x + w * 0.24, y, w * 0.52, h * 0.15, w * 0.07, { fill: fill });
      for (var i = 0; i < 3; i++) {
        rect(slide, { x: x + w * 0.42, y: y + h * (0.32 + i * 0.18), w: w * 0.38, h: h * 0.07, fill: bg });
        rect(slide, { x: x + w * 0.2, y: y + h * (0.32 + i * 0.18), w: w * 0.12, h: h * 0.07, fill: bg });
      }
      break;
    case 'sock': // Christmas stocking
      rect(slide, { x: x + w * 0.4, y: y + h * 0.1, w: w * 0.38, h: h * 0.46, fill: fill });
      ellipse(slide, x + w * 0.02, y + h * 0.42, w * 0.58, h * 0.58, { fill: fill });
      rect(slide, { x: x + w * 0.3, y: y + h * 0.42, w: w * 0.48, h: h * 0.58, fill: fill });
      rect(slide, { x: x + w * 0.28, y: y, w: w * 0.66, h: h * 0.11, fill: fill });
      break;
    case 'bauble': // ornament with a scalloped waistline
      ellipse(slide, x, y + h * 0.2, w, h * 0.8, { fill: fill });
      rect(slide, { x: x + w * 0.38, y: y + h * 0.06, w: w * 0.24, h: h * 0.2, fill: fill });
      ellipse(slide, x + w * 0.36, y, w * 0.28, w * 0.28, { fill: fill });
      slide.addShape('wave', { x: x - w * 0.02, y: y + h * 0.42, w: w * 1.04, h: h * 0.2, fill: bg });
      break;
    case 'worker': // hard-hatted figure in a hi-vis vest
      ellipse(slide, x + w * 0.28, y + h * 0.2, w * 0.44, h * 0.38, { fill: fill });
      slide.addShape('chord', { x: x + w * 0.2, y: y + h * 0.02, w: w * 0.6, h: h * 0.32, fill: fill, angleRange: [180, 0] });
      rect(slide, { x: x + w * 0.06, y: y + h * 0.26, w: w * 0.88, h: h * 0.09, fill: fill });
      slide.addShape('trapezoid', { x: x, y: y + h * 0.56, w: w, h: h * 0.44, fill: fill });
      slide.addShape('triangle', { x: x + w * 0.36, y: y + h * 0.56, w: w * 0.28, h: h * 0.3, fill: bg, rotate: 180 });
      break;
    case 'mail':
      roundRect(slide, x, y, w, h, h * 0.14, { fill: fill });
      slide.addShape('triangle', {
        x: x + w * 0.08, y: y + h * 0.06, w: w * 0.84, h: h * 0.62,
        fill: bg, rotate: 180
      });
      break;
    case 'phone': // handset with a signal arc
      slide.addShape('blockArc', {
        x: x - w * 0.05, y: y - h * 0.05, w: w * 1.1, h: h * 1.1,
        fill: fill, angleRange: [125, 25], arcThicknessRatio: 0.55
      });
      ellipse(slide, x + w * 0.02, y + h * 0.52, w * 0.44, h * 0.44, { fill: fill });
      ellipse(slide, x + w * 0.52, y + h * 0.02, w * 0.44, h * 0.44, { fill: fill });
      break;
    case 'pin':
      ellipse(slide, x, y, w, w, { fill: fill });
      rect(slide, { x: cx - w * 0.16, y: y + w * 0.5, w: w * 0.32, h: h - w * 0.5, fill: fill });
      break;
  }
}

/** White pill link used on several slides. */
function readMore(slide, x, y) {
  roundRect(slide, x + 0.04, y, 1.439, 0.429, 0.2145, { fill: { color: C.white }, shadow: SOFT });
  txt(slide, 'Read More', {
    x: x, y: y + 0.063, w: 1.518, h: 0.303,
    fontFace: MAJOR, fontSize: 12, color: C.orange, align: 'center'
  });
}

/** "$44.568 ▲ / -134 for last month" KPI block. */
function moneyStat(slide, o) {
  txt(slide, [
    { text: '$', options: { color: C.orange } },
    { text: o.amount, options: { color: o.color || C.ink } }
  ], { x: o.x, y: o.y, w: 2.317, h: 0.539, margin: 0, fontFace: MAJOR, fontSize: 32 });
  txt(slide, '-134 for last month', {
    x: o.x, y: o.y + o.capDy, w: 1.822, h: 0.185, margin: 0, fontSize: 11, color: o.color || C.ink
  });
  slide.addShape('triangle', {
    x: o.x + 1.949, y: o.y + 0.185, w: 0.139, h: 0.12,
    fill: { color: C.orange }, flipV: !!o.down
  });
}

/** Oversized opening quote mark. */
function quoteMark(slide, x, y) {
  txt(slide, '\u201C', {
    x: x - 0.06, y: y - 0.16, w: 0.5, h: 0.4, margin: 0,
    fontFace: MAJOR, fontSize: 34, bold: true, color: C.orange, valign: 'top'
  });
}

/** Stand-in for a photograph in the source deck. */
function imagePlaceholder(slide, x, y, w, h, radius) {
  roundRect(slide, x, y, w, h, radius || 0.12, { fill: { color: C.ph } });
  txt(slide, '[image]', {
    x: x, y: y + h / 2 - 0.2, w: w, h: 0.4,
    fontSize: 12, color: C.mid, align: 'center', valign: 'middle'
  });
}

/* ------------------------------------------------------------------ slides */

// 1 — full-bleed orange cover.
function slide01(p) {
  const s = p.addSlide();
  s.background = { color: C.orange };
  chrome(s, { number: 1 });
  rect(s, { x: 0, y: -0.049, w: 13.333, h: 7.549, fill: { color: C.orange, transparency: 8 } });
  txt(s, 'Building a responsible, prepared, and secure workplace for all.', {
    x: 3.004, y: 2.178, w: 9.224, h: 0.404, color: C.white, align: 'right'
  });
  txt(s, [
    { text: 'Safe', options: { color: C.white } },
    { text: 'Work', options: { color: C.ink } }
  ], { x: 0.733, y: 2.278, w: 11.867, h: 2.895, fontFace: MAJOR, fontSize: 166, bold: true, align: 'center', valign: 'middle' });
  txt(s, 'Promoting safety-first mindsets and reliable procedures to ensure secure, compliant, and healthy work environments for all.', {
    x: 1.143, y: 5.066, w: 9.224, h: 0.64, fontSize: 16, italic: true, color: C.white
  });
  logo(s, 'light');
}

// 2 — "Safety is Everyone's Job".
function slide02(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 2 });
  roundRect(s, 6.952, 2.54, 4.678, 4.466, 0.398, { fill: { color: C.orange }, shadow: GLOW });
  eyebrow(s, 'SafeWork', 1.193, 1.242);
  headline(s, ['Safety is', 'Everyone\u2019s Job'], { x: 1.193, y: 1.735, w: 7.446, h: 1.447 });
  txt(s, PROMO, { x: 1.193, y: 3.427, w: 4.685, h: 0.572, fontSize: 14, bold: true, color: C.orange });
  body(s, BODY_FULL, { x: 1.193, y: 4.168, w: 5.474, h: 0.673, color: C.black });
  iconTile(s, 1.295, 5.293, 1.0, { color: C.orange }, { shadow: LIFT });
  glyph(s, 'hammers', 1.575, 5.573, 0.44, 0.44, C.white, C.orange);
  statLabels(s, 2.479, 5.472, '5,9K', 'Work Safety', C.dark, { w: 1.49 });
  statLabels(s, 3.991, 5.472, '7,3K', 'Safety Protocol', C.dark, { w: 1.49 });
  statCircle(s, 10.828, 0.839, '5,9K', 'Emergency');
}

// 3 — agenda / roadmap.
function slide03(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  logo(s, 'dark');
  eyebrow(s, 'Agenda', 8.48, 1.715);
  headline(s, ['Our Safety Roadmap'], { x: 8.48, y: 2.204, w: 4.706, h: 1.447 });
  body(s, BODY_FULL, { x: 8.48, y: 3.92, w: 4.052, h: 0.976 });

  // Feature card
  roundRect(s, 4.199, 1.41, 3.252, 3.252, 0.742, { fill: { color: C.orange }, shadow: GLOW });
  iconTile(s, 4.771, 2.151, 0.548, { color: C.white });
  glyph(s, 'globe', 4.935, 2.314, 0.221, 0.221, C.orange, C.white);
  txt(s, 'Risk Management', { x: 4.671, y: 2.847, w: 3.236, h: 0.404, color: C.white });
  body(s, BODY_TINY, { x: 4.671, y: 3.251, w: 2.043, h: 0.673, color: C.white });

  statCircle(s, 1.678, 1.87, '5,9K', 'Safety Protocol');

  // Icon row
  iconTile(s, 8.566, 5.169, 0.811, { color: C.orange });
  glyph(s, 'tools', 8.819, 5.422, 0.305, 0.306, C.white, C.orange);
  glyph(s, 'worker', 9.85, 5.407, 0.248, 0.322, C.orange);
  glyph(s, 'bauble', 10.602, 5.388, 0.31, 0.372, C.orange);

  footer(s);
  pageNumber(s, 3);
}

// 4 — orange band across the lower half.
function slide04(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 4 });
  rect(s, { x: 0, y: 4.034, w: 13.333, h: 2.772, fill: { color: C.orange }, shadow: GLOW });
  eyebrow(s, 'Protecting People and Productivity', 7.466, 1.415, 4.712);
  headline(s, ['Why Workplace Safety Matters'], { x: 7.466, y: 1.991, w: 5.383, h: 1.447 });
  txt(s, PROMO, { x: 7.494, y: 4.62, w: 4.685, h: 0.572, fontSize: 14, bold: true, color: C.white });
  body(s, BODY_FULL, { x: 7.466, y: 5.439, w: 4.712, h: 0.974, color: C.white });
  statCircle(s, 5.476, 2.157, '5,9K', 'Work Safety');
}

// 5 — "Know the Risks".
function slide05(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 5 });
  iconTile(s, 1.098, 1.057, 0.548, { type: 'solid', color: C.orange });
  glyph(s, 'globe', 1.246, 1.204, 0.252, 0.252, C.white, C.orange);
  eyebrow(s, 'Risk Management', 0.998, 1.753, 3.236);
  body(s, BODY_MED, { x: 0.998, y: 2.157, w: 3.904, h: 0.673 });
  statCircle(s, 4.902, 2.83, '5,9K', 'Safety Protocol');
  eyebrow(s, 'Common Workplace Hazards', 6.548, 4.239, 4.395);
  headline(s, ['Know the Risks'], { x: 6.548, y: 4.736, w: 5.999, h: 0.774 });
  body(s, BODY_FULL, { x: 6.548, y: 5.677, w: 5.378, h: 0.673 });
}

// 6 — "A Mindset, Not a Manual".
function slide06(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 6 });
  roundRect(s, -0.206, 1.184, 5.235, 2.501, 0.417, { fill: { color: C.orange } });
  roundRect(s, 7.406, 1.67, 6.316, 4.161, 0.367, { fill: { color: C.white }, shadow: SOFT });

  iconTile(s, 0.737, 1.552, 0.548, { color: C.white });
  glyph(s, 'globe', 0.885, 1.7, 0.252, 0.252, C.orange, C.white);
  eyebrow(s, 'Safety Compliance', 0.637, 2.249, 3.236, C.white);
  body(s, BODY_MED, { x: 0.637, y: 2.653, w: 3.904, h: 0.673, color: C.white });

  iconTile(s, 0.737, 4.176, 0.548, { color: C.orange });
  glyph(s, 'backpack', 0.922, 4.335, 0.179, 0.209, C.white, C.orange);
  eyebrow(s, 'Secure Environment', 0.637, 4.872, 3.236);
  body(s, BODY_MED, { x: 0.637, y: 5.276, w: 3.904, h: 0.671 });

  eyebrow(s, 'Safety Culture', 7.96, 2.531);
  headline(s, ['A Mindset,', 'Not a Manual'], { x: 7.96, y: 3.012, w: 6.442, h: 1.447 });
  txt(s, 'Brief overview of what Christmas is and why it\u2019s celebrated', {
    x: 7.988, y: 4.599, w: 4.685, h: 0.572, fontSize: 14, color: C.orange
  });
}

// 7 — stacked discs beside "Employee Responsibility".
function slide07(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 7 });
  rect(s, { x: 0, y: 1.367, w: 3.195, h: 5.106, fill: { color: C.orange }, shadow: GLOW });
  eyebrow(s, 'Health Monitoring', 7.412, 1.606);
  headline(s, ['Employee Responsibility'], { x: 7.412, y: 2.105, w: 6.442, h: 1.447 });
  body(s, BODY_FULL, { x: 7.412, y: 3.687, w: 4.528, h: 0.974 });

  iconTile(s, 7.453, 5.078, 1.0, { color: C.orange });
  glyph(s, 'globe', 7.723, 5.347, 0.459, 0.46, C.white, C.orange);
  glyph(s, 'bauble', 8.847, 5.397, 0.31, 0.372, C.orange);
  glyph(s, 'hammers', 9.615, 5.425, 0.317, 0.317, C.orange);

  statCircle(s, 4.846, 1.596, '5,9K', 'Compliance');
  statCircle(s, 4.84, 3.182, '73,2K', 'Environment', { fill: C.white, color: C.orange, labelColor: C.dark, shadow: SOFT });
  statCircle(s, 4.84, 4.768, '+235', 'Equipment', { fill: C.ink });
}

// 8 — "Lead by Example" with progress bars.
function slide08(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 8 });
  rect(s, { x: 0, y: 4.163, w: 13.333, h: 2.589, fill: { color: C.orange }, shadow: GLOW });
  eyebrow(s, 'Management\u2019s Role', 1.002, 1.441);
  headline(s, ['Lead by Example'], { x: 1.002, y: 1.994, w: 6.442, h: 0.841, fontSize: 44 });
  body(s, BODY_FULL, { x: 1.002, y: 2.951, w: 5.754, h: 0.673 });

  [
    { y: 4.518, label: 'Experience', pct: '90%', filled: 2.634 },
    { y: 5.608, label: 'Skill', pct: '60%', filled: 1.836 }
  ].forEach(function (bar) {
    roundRect(s, 0.865, bar.y, 3.362, 0.789, 0.198, { fill: { color: C.ink }, shadow: GLOW });
    txt(s, bar.label, {
      x: 1.003, y: bar.y + 0.085, w: 1.752, h: 0.376,
      fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.5
    });
    hline(s, 1.107, bar.y + 0.531, 2.88, C.white, 4);
    hline(s, 1.103, bar.y + 0.531, bar.filled, C.amberLite, 4);
    txt(s, bar.pct, {
      x: 4.418, y: bar.y + 0.189, w: 1.0, h: 0.505,
      fontFace: MAJOR, fontSize: 24, bold: true, color: C.white, valign: 'middle'
    });
  });

  slide08Quote(s);
  statCircle(s, 11.525, 3.34, '5,9K', 'Equipment', {
    fill: C.white, color: C.orange, labelColor: C.gray, bold: true, shadow: SOFT
  });
}

function slide08Quote(s) {
  slide08Bubble(s, 6.951, 0.559, 3.164, 1.525);
  quoteMark(s, 7.309, 0.809);
  txt(s, [
    { text: '\u201CAn ounce of prevention is worth a pound of cure.\u201D', options: { color: C.mid, breakLine: true } },
    { text: '\u2013 ', options: { color: C.mid } },
    { text: 'Benjamin Franklin', options: { color: C.dark } }
  ], { x: 7.224, y: 1.046, w: 2.723, h: 0.707, fontSize: 12, italic: true });
}

/** Rounded speech balloon with a tail on the bottom edge. */
function slide08Bubble(s, x, y, w, h) {
  roundRect(s, x, y, w, h, 0.25, { fill: { color: C.white }, shadow: SOFT });
  s.addShape('triangle', { x: x + w * 0.44, y: y + h - 0.02, w: 0.42, h: 0.42, fill: { color: C.white }, flipV: true });
}

// 9 — three PPE cards under an orange band.
function slide09(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 9 });
  rect(s, { x: 0, y: 0, w: 13.333, h: 2.495, fill: { color: C.orange }, shadow: GLOW });
  [
    { x: 1.245, w: 2.888, tx: 1.228, title: 'Safe Equipment', cx: 1.628 },
    { x: 5.161, w: 3.082, tx: 5.27, title: 'Safety Culture', cx: 5.683 },
    { x: 9.237, w: 3.006, tx: 9.312, title: 'Safe Operations', cx: 9.725 }
  ].forEach(function (card) {
    roundRect(s, card.x, 0.87, card.w, 3.53, card.w * 0.18115, { fill: { color: C.white }, shadow: SOFT });
    txt(s, card.title, {
      x: card.tx, y: 3.485, w: 2.838, h: 0.37,
      fontSize: 16, bold: true, color: C.orange, align: 'center'
    });
    txt(s, 'Work Safety', {
      x: card.cx, y: 3.798, w: 2.038, h: 0.37,
      fontSize: 12, color: C.black, align: 'center', lineSpacingMultiple: 1.5
    });
  });
  txt(s, 'Personal Protective Equipment (PPE)', {
    x: 4.021, y: 4.905, w: 5.291, h: 0.37, fontSize: 16, bold: true, color: C.orange, align: 'center'
  });
  headline(s, ['Gear That Saves Lives'], { x: 3.049, y: 5.314, w: 7.234, h: 0.841, fontSize: 44, align: 'center' });
  body(s, BODY_FULL, { x: 3.049, y: 6.172, w: 7.234, h: 0.673, align: 'center' });
}

// 10 — "Safe Equipment Handling" (master furniture suppressed).
function slide10(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  logo(s, 'dark');
  eyebrow(s, 'Tools with Care', 7.795, 0.901);
  headline(s, ['Safe Equipment Handling'], { x: 7.795, y: 1.276, w: 5.56, h: 1.582, fontSize: 44 });
  body(s, BODY_SHORT, { x: 7.795, y: 2.858, w: 4.528, h: 0.673 });
  readMore(s, 7.795, 3.849);
  eyebrow(s, 'Safety Compliance', 0.585, 5.643, 3.236);
  body(s, 'SafeWork focuses on building a culture of safety where every.', { x: 0.585, y: 6.047, w: 3.144, h: 0.673 });
}

// 11 — "Comfort Prevents Injury".
function slide11(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 11 });
  // Panel bleeding off the left edge, rounded only on its right end.
  roundRect(s, -0.141, 2.054, 5.677, 3.392, 1.405, { fill: { color: C.orange }, shadow: GLOW });
  rect(s, { x: -0.141, y: 2.054, w: 1.5, h: 3.392, fill: { color: C.orange } });

  [0.74, 4.531].forEach(function (cardY) {
    roundRect(s, 3.122, cardY, 2.153, 2.308, 0.142, { fill: { color: C.white }, shadow: HALO });
    for (var i = 0; i < 5; i++) {
      s.addShape('star5', {
        x: 3.801 + i * 0.1645, y: cardY + 2.031, w: 0.136, h: 0.136, fill: { color: C.star }
      });
    }
  });

  [
    { y: 1.465, value: '5,9K', label: 'Emergency' },
    { y: 5.256, value: '5,9K', label: 'Equipment' }
  ].forEach(function (m) {
    txt(s, m.value, {
      x: 5.512, y: m.y, w: 1.49, h: 0.64,
      fontFace: MAJOR, fontSize: 32, bold: true, color: C.orange, valign: 'middle'
    });
    txt(s, m.label, {
      x: 5.512, y: m.y + 0.589, w: 1.166, h: 0.278,
      fontFace: MAJOR, fontSize: 10.5, color: C.dark, valign: 'middle'
    });
  });

  iconTile(s, 0.661, 2.773, 0.665, { color: C.white });
  glyph(s, 'globe', 0.841, 2.953, 0.305, 0.306, C.orange, C.white);
  eyebrow(s, 'Workplace Safety Standards', 0.58, 3.615, 4.23, C.white);
  body(s, BODY_TINY, { x: 0.58, y: 3.899, w: 3.536, h: 0.37, color: C.white });

  eyebrow(s, 'Workplace Signage', 7.821, 1.546);
  headline(s, ['Comfort ', 'Prevents Injury'], { x: 7.821, y: 1.971, w: 5.56, h: 1.582, fontSize: 44 });
  body(s, BODY_SHORT, { x: 7.821, y: 3.553, w: 4.528, h: 0.673 });

  iconTile(s, 7.899, 4.507, 0.811, { color: C.orange });
  glyph(s, 'clipboard', 8.184, 4.757, 0.24, 0.31, C.white, C.orange);
  glyph(s, 'sock', 9.12, 4.726, 0.372, 0.372, C.orange);
  glyph(s, 'bauble', 9.934, 4.726, 0.31, 0.372, C.orange);
  readMore(s, 7.821, 5.693);
}

// 12 — "Ready for Anything" with a three-tone sidebar.
function slide12(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 12 });
  [
    { y: -0.002, fill: C.amberDark, tileFill: C.orange, icon: 'sock', title: 'Work Safety' },
    { y: 2.499, fill: C.orange, tileFill: C.amberDark, icon: 'globe', title: 'Safety Protocol' },
    { y: 4.999, fill: C.amberLite, tileFill: C.orange, icon: 'bauble', title: 'Risk Management' }
  ].forEach(function (band) {
    rect(s, { x: 10.318, y: band.y, w: 3.024, h: 2.501, fill: { color: band.fill } });
    iconTile(s, 10.755, band.y + 0.653, 0.404, { color: band.tileFill });
    glyph(s, band.icon, 10.864, band.y + 0.762, 0.185, 0.186, C.white, band.tileFill);
    eyebrow(s, band.title, 10.665, band.y + 1.139, 3.236, C.white);
    body(s, BODY_TINY, { x: 10.665, y: band.y + 1.474, w: 2.387, h: 0.671, color: C.white });
  });

  eyebrow(s, 'Emergency Preparedness', 0.87, 1.706, 3.957);
  headline(s, ['Ready', 'for Anything'], { x: 0.87, y: 2.212, w: 5.172, h: 1.582, fontSize: 44 });
  body(s, BODY_SHORT, { x: 0.87, y: 3.902, w: 4.528, h: 0.671 });
  moneyStat(s, { x: 0.978, y: 5.101, amount: '44.568', capDy: 0.693 });
  moneyStat(s, { x: 3.574, y: 5.101, amount: '64.890', capDy: 0.693 });
}

// 13 — "Speak Up, Stay Safe".
function slide13(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 13 });
  rect(s, { x: 0, y: 0, w: 5.172, h: 7.5, fill: { color: C.orange }, shadow: GLOW });

  [
    { x: 1.15, pill: 1.435, tx: 1.348, body: 1.387, title: 'Inspection', tw: 2.58 },
    { x: 4.132, pill: 4.374, tx: 4.287, body: 4.326, title: 'Compliance', tw: 2.249 }
  ].forEach(function (card) {
    roundRect(s, card.x, 1.71, 2.683, 4.08, 0.192, { fill: { color: C.white }, shadow: SOFT });
    roundRect(s, card.pill, 1.997, 1.343, 0.386, 0.193, { fill: { color: C.orange }, shadow: SOFT });
    txt(s, 'Secure ', {
      x: card.tx, y: 2.038, w: 1.518, h: 0.303,
      fontFace: MAJOR, fontSize: 12, color: C.white, align: 'center'
    });
    eyebrow(s, card.title, card.body, 2.559, card.tw);
    body(s, 'SafeWork focuses on building a culture of.', { x: card.body, y: 2.865, w: 2.415, h: 0.671 });
  });

  eyebrow(s, 'Incident Reporting', 7.497, 1.945);
  headline(s, ['Speak Up,', 'Stay Safe'], { x: 7.497, y: 2.32, w: 5.56, h: 1.582, fontSize: 44 });
  body(s, BODY_SHORT, { x: 7.497, y: 3.902, w: 4.528, h: 0.671 });
  moneyStat(s, { x: 7.604, y: 5.101, amount: '44.568', capDy: 0.577 });
  moneyStat(s, { x: 10.197, y: 5.101, amount: '64.890', capDy: 0.577, down: true });
}

// 14 — "Psychological Safety Counts" with the market-analysis line chart.
function slide14(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 14 });
  rect(s, { x: 9.36, y: 0, w: 4.024, h: 7.5, fill: { color: C.orange }, shadow: WIDE });

  eyebrow(s, 'Mental Well-Being', 0.684, 2.008);
  headline(s, ['Psychological Safety Counts'], { x: 0.684, y: 2.477, w: 5.56, h: 1.447 });
  body(s, BODY_SHORT, { x: 0.684, y: 4.059, w: 4.427, h: 0.673 });
  readMore(s, 0.684, 5.267);

  roundRect(s, 5.624, 1.02, 7.959, 5.46, 0.256, { fill: { color: C.white }, shadow: WIDE });
  txt(s, 'Market Analysis', { x: 6.097, y: 1.327, w: 3.938, h: 0.37, fontFace: MAJOR, fontSize: 16, bold: true, color: C.orange });
  marketChart(p, s);

  // Annotation layer over the plot
  s.addShape('line', { x: 9.367, y: 1.963, w: 0, h: 3.171, line: { color: C.dark, width: 1, dashType: 'dash' } });
  ellipse(s, 9.289, 4.014, 0.156, 0.156, { fill: { color: C.white }, line: { color: C.orange, width: 2 } });
  roundRect(s, 8.85, 3.215, 1.08, 0.299, 0.1495, { fill: { color: C.ink } });
  txt(s, [
    { text: '+', options: { color: C.orange } },
    { text: '657%', options: { color: C.white } }
  ], { x: 8.783, y: 3.221, w: 1.213, h: 0.286, fontFace: MAJOR, fontSize: 11, align: 'center', valign: 'middle' });

  roundRect(s, 9.879, 4.52, 1.098, 0.438, 0.1, { fill: { color: C.white }, shadow: SOFT });
  s.addShape('triangle', { x: 9.75, y: 4.63, w: 0.22, h: 0.22, fill: { color: C.white }, rotate: 270 });
  txt(s, '230K', { x: 9.925, y: 4.553, w: 1.247, h: 0.337, fontFace: MAJOR, fontSize: 14, color: C.orange, valign: 'middle' });
  txt(s, 'Total', { x: 10.508, y: 4.612, w: 0.626, h: 0.219, fontFace: MAJOR, fontSize: 7, color: C.orange, valign: 'middle' });

  roundRect(s, 6.035, 5.757, 3.458, 0.441, 0.2205, { fill: { color: C.orange, transparency: 85 } });
  ellipse(s, 6.202, 5.931, 0.092, 0.092, { fill: { color: C.orange } });
  txt(s, 'Risk Prevention Strategy', { x: 6.424, y: 5.882, w: 2.068, h: 0.185, margin: 0, fontSize: 11, color: C.dark });
  roundRect(s, 8.715, 5.875, 0.645, 0.218, 0.109, { fill: { color: C.orange } });
  txt(s, '+657%', {
    x: 8.634, y: 5.865, w: 0.774, h: 0.236,
    fontFace: MAJOR, fontSize: 8, color: C.white, align: 'center', valign: 'middle'
  });
}

function marketChart(p, s) {
  const data = [{
    name: '2020',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    values: [4, 7, 3.5, 2, 2, 3, 4, 5.2, 4, 3, 4, 1]
  }];
  s.addChart(p.ChartType.line, data, {
    x: 6.035, y: 2.014, w: 7.182, h: 3.545,
    showLegend: false,
    chartColors: [C.orange],
    lineSize: 2,
    lineDataSymbol: 'circle',
    lineDataSymbolSize: 5,
    lineDataSymbolLineColor: C.orange,
    catAxisLabelColor: C.dark,
    catAxisLabelFontFace: MINOR,
    catAxisLabelFontSize: 8,
    catAxisLineColor: C.grid,
    catGridLine: { style: 'none' },
    valAxisLabelColor: C.dark,
    valAxisLabelFontFace: MINOR,
    valAxisLabelFontSize: 8,
    valAxisLineShow: false,
    valAxisMaxVal: 8,
    valAxisMinVal: 0,
    valAxisMajorUnit: 1,
    valGridLine: { color: C.grid, size: 1 },
    plotArea: { fill: { color: C.white } },
    chartArea: { fill: { color: C.white } }
  });
}

// 15 — "Safety Training Programs" (tablet photo → placeholder).
function slide15(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  rect(s, { x: 0, y: 0, w: 5.433, h: 7.5, fill: { color: C.orange }, shadow: GLOW });
  imagePlaceholder(s, 0.348, 1.443, 6.842, 4.844, 0.3);
  logo(s, 'dark');

  eyebrow(s, 'Learn to Protect', 7.867, 1.999);
  headline(s, ['Safety Training Programs'], { x: 7.867, y: 2.557, w: 5.56, h: 1.447 });
  body(s, BODY_SHORT, { x: 7.867, y: 4.163, w: 4.528, h: 0.671 });
  moneyStat(s, { x: 7.974, y: 5.362, amount: '44.568', capDy: 0.577, color: C.dark });
  moneyStat(s, { x: 10.605, y: 5.362, amount: '64.890', capDy: 0.577, color: C.dark });

  statCircle(s, 0.9, 1.665, '5,9K', 'Guidelines', {
    fill: C.white, color: C.orange, labelColor: C.dark, shadow: SOFT
  });

  roundRect(s, 4.661, 5.095, 2.53, 1.308, 0.294, { fill: { color: C.white }, shadow: SOFT });
  quoteMark(s, 5.037, 5.365);
  txt(s, 'Risk Management Strategy Outline', {
    x: 4.952, y: 5.602, w: 2.223, h: 0.572, fontSize: 14, color: C.orange
  });
}

// 16 — "Smarter Safety Management" (phone photo → placeholder).
function slide16(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 16 });
  imagePlaceholder(s, 8.217, 0.524, 3.406, 6.452, 0.4);

  eyebrow(s, 'Digital Safety Tools', 0.937, 1.297);
  headline(s, ['Smarter Safety Management'], { x: 0.937, y: 1.802, w: 6.612, h: 1.582, fontSize: 44 });
  body(s, BODY_FULL, { x: 0.937, y: 3.565, w: 5.979, h: 0.671 });

  iconTile(s, 0.993, 4.865, 1.0, { color: C.orange });
  glyph(s, 'globe', 1.264, 5.134, 0.459, 0.46, C.white, C.orange);
  statLabels(s, 1.993, 5.044, '5,9K', 'Equipment', C.dark, { w: 1.49 });
  statLabels(s, 3.295, 5.044, '236K', 'Protective ', C.dark, { w: 1.49 });
  statLabels(s, 4.596, 5.044, '+489', 'Workplace ', C.dark, { w: 1.49 });

  roundRect(s, 7.403, 4.662, 2.53, 1.308, 0.294, { fill: { color: C.white }, shadow: SOFT });
  quoteMark(s, 7.78, 4.932);
  txt(s, 'Christmas is a time for togetherness', {
    x: 7.695, y: 5.168, w: 2.223, h: 0.572, fontSize: 14, color: C.orange
  });
  statCircle(s, 10.67, 1.359, '5,9K', 'Reporting ');
}

// 17 — contact card on an orange page.
function slide17(p) {
  const s = p.addSlide();
  s.background = { color: C.orange };
  chrome(s, { number: 17 });
  roundRect(s, -0.517, 2.344, 12.884, 4.307, 0.379, { fill: { color: C.white }, shadow: SOFT });
  txt(s, 'Let\u2019s Build Safer Workplaces Together', {
    x: 0.451, y: 0.894, w: 12.431, h: 0.841,
    fontFace: MAJOR, fontSize: 44, bold: true, color: C.white, align: 'center'
  });
  txt(s, 'Contact Us', { x: 1.413, y: 2.919, w: 4.641, h: 0.404, fontFace: MAJOR, bold: true, color: C.dark });
  body(s, BODY_MED, { x: 1.413, y: 3.323, w: 3.996, h: 0.673 });
  [
    { icon: 'mail', ix: 1.506, iy: 4.421, iw: 0.242, ih: 0.155, tx: 1.849, ty: 4.31, tw: 3.204, text: 'loremipsum@mail.com' },
    { icon: 'phone', ix: 1.521, iy: 5.117, iw: 0.212, ih: 0.212, tx: 1.849, ty: 5.055, tw: 2.794, text: '+0123 4567 890 000' },
    { icon: 'pin', ix: 1.559, iy: 5.801, iw: 0.106, ih: 0.212, tx: 1.834, ty: 5.739, tw: 3.204, text: 'Your Location Here' }
  ].forEach(function (row) {
    glyph(s, row.icon, row.ix, row.iy, row.iw, row.ih, C.orange, C.white);
    txt(s, row.text, { x: row.tx, y: row.ty, w: row.tw, h: 0.337, fontSize: 14, color: C.orange, valign: 'middle' });
  });
}

// 18 — closing slide.
function slide18(p) {
  const s = p.addSlide();
  s.background = { color: C.page };
  chrome(s, { number: 18 });
  rect(s, { x: 9.36, y: 0, w: 4.024, h: 7.5, fill: { color: C.orange }, shadow: WIDE });
  eyebrow(s, 'Stay Safe, Work Smart', 0.68, 2.513);
  headline(s, ['Thank You.'], { x: 0.68, y: 3.102, w: 7.247, h: 1.447, fontSize: 80 });
  body(s, 'Thank you for prioritizing safety\u2014together, we create better environments for everyone.', {
    x: 0.68, y: 4.767, w: 4.528, h: 0.673, color: C.gray
  });
}

/* -------------------------------------------------------------------- main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_13x75', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_13x75';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };
  pptx.title = 'SafeWork';
  BUILDERS.forEach(function (fn) { fn(pptx); });
  return pptx.writeFile({
    fileName: path.join(__dirname, '024b6816-e76e-4355-ac11-96f269dcde96_grok_final.pptx')
  });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) { console.error(e); process.exit(1); });
