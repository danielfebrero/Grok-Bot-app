/**
 * "Restaurant" — Food and Beverage Pitch Presentation (10 slides, 20" x 11.25").
 *
 * Standalone pptxgenjs re-creation of the reference deck. Run with:
 *   node 1095f3d4-368b-44e8-8a0b-8d12a459d5c7_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const SLIDE_W = 20;
const SLIDE_H = 11.25;

const C = {
  ink: '1B1B1B',        // theme dk1 - headlines
  ink2: '242424',       // theme dk2 - subtitles
  white: 'FDFDFD',      // theme lt1 / accent5 - page + reversed text
  paper: 'FAFAFA',      // theme lt2 - card fill
  brand: '661606',      // theme accent1 - deep maroon (gradient start)
  brandLite: 'AA250A',  // theme accent2 - warm red   (gradient end)
  orange: 'D92F0D',     // theme accent3 - triangle wedge
  body: '474747',       // theme accent6 - body copy
  link: '0000FF',
  badge: 'E4EDF4'       // step-number text inside the maroon badges
};

const SERIF = 'Playfair Display';
const SANS = 'Poppins';

const NO_LINE = { type: 'none' };

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

const hex = (n) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0').toUpperCase();
const rgb = (c) => [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)];

/** The deck's signature fill: accent1 held for the first 4%, then a linear ramp to accent2. */
function brandRamp(t) {
  const u = Math.max(0, Math.min(1, (t - 0.04) / 0.96));
  const a = rgb(C.brand);
  const b = rgb(C.brandLite);
  return hex(a[0] + (b[0] - a[0]) * u) + hex(a[1] + (b[1] - a[1]) * u) + hex(a[2] + (b[2] - a[2]) * u);
}

/**
 * How many solid bands to spend on a gradient run: fine enough that neither the
 * colour steps (the whole ramp is only ~68 levels wide) nor the physical steps
 * (one band per 0.15") are visible, and never more than that.
 */
const bandCount = (travel, extent, levels) =>
  Math.max(1, Math.min(Math.round(Math.abs(travel) * (levels || 68)), Math.round(extent / 0.15)));

/**
 * Banded linear gradient inside a circle: a base disc, then chord segments that
 * each slice off a shallower cap. `rotate` turns the band direction (0 = top to
 * bottom, 90 = left to right, 315 = the 45-degree diagonal used on slide 8).
 */
function gradientDisc(slide, o) {
  const from = o.from === undefined ? 0 : o.from;
  const to = o.to === undefined ? 1 : o.to;
  const steps = Math.max(4, bandCount(to - from, o.d));
  const box = { x: o.x, y: o.y, w: o.d, h: o.d, line: NO_LINE, rotate: o.rotate || 0 };
  const shade = (k) => ({ color: brandRamp(from + (to - from) * ((k + 0.5) / steps)) });

  slide.addShape('ellipse', Object.assign({}, box, { fill: shade(0) }));
  for (let k = 1; k < steps; k++) {
    const deg = (Math.asin((2 * k) / steps - 1) * 180) / Math.PI; // chord line at depth k/steps
    slide.addShape('chord', Object.assign({}, box, {
      fill: shade(k),
      angleRange: [(deg + 360) % 360, (180 - deg + 360) % 360]
    }));
  }
}

/**
 * Banded linear gradient inside a (optionally rounded) rectangle.
 *   axis 'v' - dark top       -> light bottom
 *   axis 'h' - dark left      -> light right
 *   axis 'd' - dark top-right -> light bottom-left (the 45-degree panel on slide 4)
 *
 * The outline is split into four corner discs plus three straight regions: the wide
 * middle band (which also trims each disc back to the quadrant that forms its
 * corner) and the two short caps between the discs at either end. Every region is
 * then tiled with solid rectangles, grown by a hair so no seam shows between them.
 */
const SEAM = 0.02; // inches

function gradientShape(slide, o) {
  const r = Math.min(o.rectRadius || 0, o.w / 2, o.h / 2);
  const from = o.from === undefined ? 0 : o.from;
  const to = o.to === undefined ? 1 : o.to;
  const diag = o.axis === 'd';
  const horiz = o.axis === 'h';
  // Local point -> position along the ramp.
  const at = (x, y) => from + (to - from) *
    (diag ? ((o.w - x) + y) / (o.w + o.h) : horiz ? x / o.w : y / o.h);

  if (r > 0) {
    // Each corner disc carries the slice of the ramp that crosses it.
    const half = (to - from) * (diag ? (r * Math.SQRT2) / (o.w + o.h) : horiz ? r / o.w : r / o.h);
    const spin = diag ? 45 : horiz ? 90 : 0;
    [0, o.w - 2 * r].forEach((cx) => [0, o.h - 2 * r].forEach((cy) => {
      const t = at(cx + r, cy + r);
      gradientDisc(slide, { x: o.x + cx, y: o.y + cy, d: 2 * r, rotate: spin, from: t - half, to: t + half });
    }));
  }

  const tile = (x0, y0, x1, y1) => {
    // A diagonal ramp needs a 2-D grid, so trade a little colour resolution for a
    // shape count in the hundreds rather than the thousands.
    const levels = diag ? 34 : 68;
    const nx = diag || horiz ? bandCount(at(x1, y0) - at(x0, y0), x1 - x0, levels) : 1;
    const ny = diag || !horiz ? bandCount(at(x0, y1) - at(x0, y0), y1 - y0, levels) : 1;
    for (let i = 0; i < nx; i++) {
      for (let j = 0; j < ny; j++) {
        const xa = x0 + ((x1 - x0) * i) / nx;
        const xb = Math.min(x1, x0 + ((x1 - x0) * (i + 1)) / nx + SEAM);
        const ya = y0 + ((y1 - y0) * j) / ny;
        const yb = Math.min(y1, y0 + ((y1 - y0) * (j + 1)) / ny + SEAM);
        slide.addShape('rect', {
          x: o.x + xa, y: o.y + ya, w: xb - xa, h: yb - ya,
          fill: { color: brandRamp(at((xa + xb) / 2, (ya + yb) / 2)) }, line: NO_LINE
        });
      }
    }
  };
  // Middle first (full width across the ramp), then the two end caps.
  if (horiz) {
    tile(r, 0, o.w - r, o.h);
    if (r > 0) { tile(0, r, r + SEAM, o.h - r); tile(o.w - r - SEAM, r, o.w, o.h - r); }
  } else {
    tile(0, r, o.w, o.h - r);
    if (r > 0) { tile(r, 0, o.w - r, r + SEAM); tile(r, o.h - r - SEAM, o.w - r, o.h); }
  }
}

/** Soft drop shadow used by the light "card" panels. */
const cardShadow = (opacity) => ({ type: 'outer', color: '000000', opacity, blur: 20, offset: 0.001, angle: 90 });

/** Text box with the deck's defaults: real text box, top anchored, left aligned. */
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({ valign: 'top', align: 'left', isTextBox: true }, o));
}

/* ------------------------------------------------------------------ *
 * Repeated ornaments
 * ------------------------------------------------------------------ */

/** Four page dots in the top-right corner; the third one is the "active" solid dot. */
function pageDots(slide) {
  [17.89, 18.168, 18.446, 18.724].forEach((x, i) => {
    slide.addShape('ellipse', {
      x, y: 0.768, w: 0.183, h: 0.183,
      fill: { color: C.brand, transparency: i === 2 ? 0 : 90 },
      line: NO_LINE
    });
  });
}

const WHITE = { fill: { color: C.white }, line: NO_LINE };

/** Rounded maroon tile holding a translucent disc and a white taco mark. */
function logoMark(slide, x) {
  const y = 0.542;
  slide.addShape('roundRect', { x, y, w: 0.642, h: 0.659, rectRadius: 0.206, fill: { color: C.brand }, line: NO_LINE });
  slide.addShape('ellipse', { x: x + 0.087, y: y + 0.095, w: 0.468, h: 0.468, fill: { color: C.white, transparency: 75 }, line: NO_LINE });
  // Taco = the lower half of a disc (hence the double-height box), tipped
  // counter-clockwise about the centre of its flat top edge.
  slide.addShape('chord', Object.assign({
    x: x + 0.204, y: y + 0.020, w: 0.234, h: 0.412, angleRange: [0, 180], rotate: 336
  }, WHITE));
}

/** Filled disc bearing a white check mark, drawn as two rotated bars. */
function checkDot(slide, x, y, d, fill) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill, line: NO_LINE });
  slide.addShape('rect', Object.assign({ x: x + 0.20 * d, y: y + 0.52 * d, w: 0.30 * d, h: 0.10 * d, rotate: 45 }, WHITE));
  slide.addShape('rect', Object.assign({ x: x + 0.34 * d, y: y + 0.50 * d, w: 0.52 * d, h: 0.10 * d, rotate: -45 }, WHITE));
}

/** Crown pictogram: a flared body, three balled peaks and a detached base bar. */
function crown(slide, x, y, w, h) {
  const ball = 0.135 * w;
  slide.addShape('trapezoid', Object.assign({ x: x + 0.06 * w, y: y + 0.44 * h, w: 0.86 * w, h: 0.30 * h, flipV: true }, WHITE));
  [[0.067, 0.244], [0.492, 0.090], [0.916, 0.244]].forEach(([cx, cy]) => {
    slide.addShape('ellipse', Object.assign({ x: x + cx * w - ball / 2, y: y + cy * h - ball / 2, w: ball, h: ball }, WHITE));
    slide.addShape('triangle', Object.assign({ x: x + (cx - 0.13) * w, y: y + cy * h, w: 0.26 * w, h: (0.74 - cy) * h }, WHITE));
  });
  slide.addShape('roundRect', Object.assign({ x: x + 0.118 * w, y: y + 0.795 * h, w: 0.739 * w, h: 0.205 * h, rectRadius: 0.02 * h }, WHITE));
}

/* ------------------------------------------------------------------ *
 * Shared copy
 * ------------------------------------------------------------------ */

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce consequat quam sit amet pellentesque ' +
  'rutrum. Curabitur ullamcorper maximus mi, vel blandit orci viverra in. Praesent suscipit felis sem, ' +
  'PLACEHOLDER';
const LOREM_CARD =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce cons equat quam sit amet pellentesque ' +
  'rutru rabitur ullam corpe';
const SITE = 'www.insertyoursitehere.com';

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Title: giant serif wordmark over a full-bleed maroon panel.
function slide1(pptx) {
  const s = pptx.addSlide();
  gradientShape(s, { x: 7.0, y: 6.625, w: 11.906, h: 4.625, axis: 'v' });
  logoMark(s, 8.0);
  pageDots(s);

  text(s, 'Restaurant', { x: 2.447, y: 1.042, w: 16.719, h: 4.123, fontSize: 239, fontFace: SERIF, color: C.ink });

  text(s, [
    { text: 'FOOD AND BEVERAGE PITCH ', options: { color: C.ink } },
    { text: 'PRESENTATION', options: { color: C.brandLite, bold: true } }
  ], { x: 7.917, y: 4.875, w: 5.75, h: 1.313, fontSize: 36, fontFace: SANS });

  text(s, '278+', { x: 7.883, y: 7.317, w: 2.105, h: 0.808, fontSize: 42, fontFace: SERIF, color: C.white });
  text(s, 'Your Topic Here', { x: 7.883, y: 8.162, w: 3.201, h: 0.508, fontSize: 20, fontFace: SANS, color: C.white, lineSpacingMultiple: 1.3 });

  [9.374, 10.124].forEach((y) => {
    checkDot(s, 8.0, y, 0.381, { color: C.brandLite });
    text(s, 'Short Detail Description Here', {
      x: 8.403, y: y - 0.082, w: 4.582, h: 0.453,
      fontSize: 18, fontFace: SERIF, color: C.white, lineSpacingMultiple: 1.3
    });
  });
}

// 2 — About Us: split layout, maroon square with an orange right-triangle wedge.
function slide2(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 10, h: SLIDE_H, fill: { color: C.brand }, line: NO_LINE });
  s.addShape('rtTriangle', { x: 0, y: 0, w: 10, h: SLIDE_H, fill: { color: C.orange }, line: NO_LINE });
  pageDots(s);

  text(s, 'About Us', { x: 10.833, y: 2.042, w: 5.659, h: 1.717, fontSize: 96, fontFace: SERIF, color: C.ink });
  text(s, 'Type Your Subtitle Here', { x: 10.825, y: 4.127, w: 7.525, h: 0.571, fontSize: 24, fontFace: SERIF, color: C.ink2, lineSpacingMultiple: 1.3 });
  text(s, LOREM_LONG, { x: 10.825, y: 4.865, w: 7.658, h: 1.88, fontSize: 16.5, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.3 });
  text(s, SITE, { x: 10.833, y: 7.492, w: 7.658, h: 0.437, fontSize: 16.5, fontFace: SANS, color: C.link, underline: { style: 'sng' }, lineSpacingMultiple: 1.3 });
}

// 3 — Our Targets and Goals: checklist on the left, numbered badges on the right.
function slide3(pptx) {
  const s = pptx.addSlide();
  gradientShape(s, { x: 7.607, y: 7.292, w: 5.307, h: 3.359, rectRadius: 0.563, axis: 'v' });

  ['01', '02', '03'].forEach((label, i) => {
    const y = 1.672 + i * 3.271;
    gradientShape(s, { x: 14.885, y, w: 1.365, h: 1.365, rectRadius: 0.497, axis: 'v' });
    s.addText(label, {
      x: 14.885, y, w: 1.365, h: 1.365, align: 'center', valign: 'middle',
      fontSize: 27, fontFace: SANS, color: C.badge, margin: 0.075 * 72
    });
  });

  text(s, 'Our Targets and Goals', { x: 0.833, y: 1.349, w: 5.801, h: 2.524, fontSize: 72, fontFace: SERIF, color: C.ink });
  text(s, 'Nulla et erat mauris. Praesent finibus tincidunt libero sagit.',
    { x: 0.833, y: 4.626, w: 5.584, h: 0.967, fontSize: 18, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.5 });

  // Three identical "Your Solution Here" rows.
  [5.859, 6.941, 8.024].forEach((y) => {
    checkDot(s, 0.994, y + 0.111, 0.483, { color: brandRamp(0.5) });
    text(s, [
      { text: 'Your Solution Here: ', options: { bold: true } },
      { text: 'consectetur adipiscing elit Aliquam sed ullamcorper leo. Maecenas eget placerat odio' }
    ], { x: 1.591, y, w: 4.826, h: 0.934, fontSize: 16.5, fontFace: SANS, color: C.body });
  });

  text(s, 'Your Topic Here', { x: 8.236, y: 8.109, w: 3.201, h: 0.649, fontSize: 28, fontFace: SERIF, color: C.white, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet co ns ectetur adipiscing elit',
    { x: 8.236, y: 8.973, w: 4.049, h: 0.861, fontSize: 18, fontFace: SANS, color: C.white, lineSpacingMultiple: 1.3 });
}

// 4 — Managing reputation: headline above a wide diagonal-gradient panel.
function slide4(pptx) {
  const s = pptx.addSlide();
  pageDots(s);
  gradientShape(s, { x: 7.0, y: 5.625, w: 12.013, h: 5.167, rectRadius: 0.774, axis: 'd' });

  text(s, 'Managing reputation', { x: 9.613, y: 1.439, w: 8.369, h: 3.332, fontSize: 96, fontFace: SERIF, color: C.ink });
  text(s, 'Type Your Subtitle Here', { x: 7.952, y: 6.292, w: 9.934, h: 0.727, fontSize: 32, fontFace: SERIF, color: C.white, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce consequat quam sit amet pellentesque rutrum. Curabitur ullamcorper maximus mi, vel blandit orci viverra in',
    { x: 7.952, y: 7.067, w: 10.109, h: 1.383, fontSize: 20, fontFace: SANS, color: C.white, lineSpacingMultiple: 1.3 });

  s.addText(
    ['Consectetur adipiscing elit', 'Fusce consequat quam', 'Sit amet pellentesque rutrum'].map((t) => ({
      text: t,
      options: { breakLine: true, bullet: { characterCode: '2022', indent: 20.25 } }
    })),
    {
      x: 7.952, y: 8.757, w: 6.004, h: 1.368, valign: 'top', isTextBox: true,
      fontSize: 20, fontFace: SERIF, color: C.white, lineSpacingMultiple: 1.3
    }
  );
}

// 5 — Health and Safety Standards: stat rail inside a tall rounded maroon panel.
function slide5(pptx) {
  const s = pptx.addSlide();
  gradientShape(s, { x: 14.654, y: 1.184, w: 4.312, h: 9.0, rectRadius: 0.677, axis: 'v' });
  logoMark(s, 1.0);

  text(s, 'Health and Safety Standards', { x: 1.0, y: 1.7, w: 8.369, h: 4.544, fontSize: 88, fontFace: SERIF, color: C.ink });
  text(s, 'Type Your Subtitle Here', { x: 1.0, y: 6.744, w: 6.214, h: 0.571, fontSize: 24, fontFace: SERIF, color: C.ink2, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce consequat quam sit amet pellentesque rutrum. Curabitur ullamcorper maximus mi, vel blandit orci viverra in. Praesent suscipit felis sem, sed tempus sapien cursus dapibus. Nullam mi odio, aliquet eget tortor a, commodo. ',
    { x: 1.0, y: 7.52, w: 6.324, h: 2.241, fontSize: 16.5, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.3 });

  // Three "278+ / Type Your Achievements" pairs; the reference nudges each pair slightly.
  [[15.762, 1.875, 15.580, 2.809], [15.758, 4.697, 15.576, 5.630], [15.758, 7.520, 15.576, 8.454]]
    .forEach(([numX, numY, capX, capY]) => {
      s.addText('278+', { x: numX, y: numY, w: 2.105, h: 0.909, align: 'center', valign: 'top', fontSize: 48, fontFace: SERIF, color: C.white });
      s.addText('Type Your Achievements', { x: capX, y: capY, w: 2.469, h: 0.909, align: 'center', valign: 'top', fontSize: 24, fontFace: SERIF, color: C.white });
    });
}

// 6 — Ingredients and Food Quality: centred headline over two soft stat cards.
function slide6(pptx) {
  const s = pptx.addSlide();
  const card = { w: 5.763, h: 5.526, fill: { color: C.paper }, line: NO_LINE, flipH: true };
  s.addShape('roundRect', Object.assign({ x: 13.276, y: 4.612, rectRadius: 0.682, shadow: cardShadow(0.098) }, card));
  s.addShape('roundRect', Object.assign({ x: 0.974, y: 4.625, rectRadius: 0.632, shadow: cardShadow(0.098) }, card));
  logoMark(s, 1.0);
  pageDots(s);

  s.addText('Ingredients and Food Quality', {
    x: 1.555, y: 1.7, w: 16.89, h: 1.582, align: 'center', valign: 'top', isTextBox: true,
    fontSize: 88, fontFace: SERIF, color: C.ink
  });

  [['278+', 1.522, 5.389, 6.838, 8.442], ['90%', 13.824, 5.375, 6.824, 8.429]].forEach(([stat, x, yStat, yTopic, yBody]) => {
    text(s, stat, { x, y: yStat, w: 3.207, h: 1.582, fontSize: 88, fontFace: SERIF, color: C.brand });
    text(s, 'Your Topic Here', { x, y: yTopic, w: 3.689, h: 0.716, fontSize: 32, fontFace: SERIF, color: C.ink2, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet con sectetur adipiscing elit usce',
      { x, y: yBody, w: 4.667, h: 0.945, fontSize: 20, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.3 });
  });
}

// 7 — Menus and Promotions: single oversized card with two numbered blurbs.
function slide7(pptx) {
  const s = pptx.addSlide();
  s.addShape('roundRect', {
    x: 6.974, y: 0.792, w: 12.052, h: 9.667, rectRadius: 0.676,
    fill: { color: C.paper }, line: NO_LINE, flipH: true, shadow: cardShadow(0.098)
  });

  text(s, 'Menus and Promotions', { x: 10.914, y: 1.208, w: 8.112, h: 3.063, fontSize: 88, fontFace: SERIF, color: C.ink });

  [['01. Type Your Subtitle Here', 5.015, 5.659], ['02. Type Your Subtitle Here', 7.148, 7.792]].forEach(([head, yHead, yBody]) => {
    text(s, head, { x: 10.914, y: yHead, w: 6.26, h: 0.571, fontSize: 24, fontFace: SERIF, color: C.ink2, lineSpacingMultiple: 1.3 });
    text(s, LOREM_CARD, { x: 10.914, y: yBody, w: 6.982, h: 1.158, fontSize: 16.5, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.3 });
  });
}

// 8 — Reasons to Choose Us: crown badge, check list and a floating caption card.
function slide8(pptx) {
  const s = pptx.addSlide();
  gradientShape(s, { x: 1.5, y: 5.768, w: 4.31, h: 4.307, rectRadius: 0.552, axis: 'v' });
  s.addShape('roundRect', {
    x: 8.904, y: 7.333, w: 9.19, h: 2.292, rectRadius: 0.504,
    fill: { color: C.white }, line: NO_LINE, shadow: cardShadow(0.149)
  });
  pageDots(s);

  // Circular crown badge straddling the top edge of the maroon card (45-degree ramp).
  gradientDisc(s, { x: 2.804, y: 5.391, d: 1.714, rotate: 315 });
  crown(s, 3.346, 6.036, 0.631, 0.425);

  text(s, 'Reasons to Choose Us', { x: 10.914, y: 1.042, w: 8.112, h: 3.063, fontSize: 88, fontFace: SERIF, color: C.ink });

  [5.086, 5.651, 6.216].forEach((y) => {
    checkDot(s, 10.766, y, 0.315, { color: C.brandLite });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing',
      { x: 11.086, y: y - 0.011, w: 6.771, h: 0.404, fontSize: 18, fontFace: SANS, color: C.ink2 });
  });

  s.addText('Your Subtitle Here', { x: 2.237, y: 8.208, w: 2.848, h: 0.438, align: 'center', valign: 'top', fontSize: 20, fontFace: SERIF, color: C.white });
  s.addText('Lorem ipsum dolor sit amet co nsec tetur adipisci', { x: 1.987, y: 8.765, w: 3.349, h: 0.64, align: 'center', valign: 'top', fontSize: 16, fontFace: SANS, color: C.white });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce consequat quam sit amet pellentesque rutrum. Curabitur ullamcorper maximus mi, vel blandit orci viverra in. ',
    { x: 9.502, y: 7.974, w: 7.993, h: 1.01, fontSize: 18, fontFace: SANS, color: C.body });
}

// 9 — Eco-friendly waste management: full-height maroon column on the left.
function slide9(pptx) {
  const s = pptx.addSlide();
  gradientShape(s, { x: 0, y: 0, w: 9.083, h: SLIDE_H, axis: 'v' });
  pageDots(s);

  text(s, 'Eco-friendly waste management', { x: 10.454, y: 1.625, w: 9.546, h: 2.524, fontSize: 72, fontFace: SERIF, color: C.ink });
  text(s, 'Type Your Subtitle Here', { x: 10.454, y: 5.34, w: 7.525, h: 0.571, fontSize: 24, fontFace: SERIF, color: C.ink2, lineSpacingMultiple: 1.3 });
  text(s, LOREM_LONG, { x: 10.454, y: 6.078, w: 7.658, h: 1.88, fontSize: 16.5, fontFace: SANS, color: C.body, lineSpacingMultiple: 1.3 });
  text(s, SITE, { x: 10.462, y: 8.704, w: 7.658, h: 0.437, fontSize: 16.5, fontFace: SANS, color: C.link, underline: { style: 'sng' }, lineSpacingMultiple: 1.3 });
}

// 10 — Thanks: oversized serif sign-off beside a huge maroon disc.
function slide10(pptx) {
  const s = pptx.addSlide();
  // The reference blob is a circle (r = 7.152") cropped by the top and right slide edges;
  // its gradient is measured over y = 0 .. 10.708, hence the negative start fraction.
  gradientDisc(s, { x: 9.497, y: -3.596, d: 14.304, from: -3.596 / 10.708, to: 1 });
  logoMark(s, 1.0);

  text(s, 'Thanks', { x: 0.833, y: 5.458, w: 9.416, h: 3.45, fontSize: 199, fontFace: SERIF, color: C.ink });
  text(s, 'SEE YOU NEXT TIME', { x: 0.833, y: 8.984, w: 8.178, h: 0.841, fontSize: 44, fontFace: SANS, color: C.ink });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM_20x11_25', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CUSTOM_20x11_25';
  pptx.theme = { headFontFace: SERIF, bodyFontFace: SANS };
  pptx.defineSlideMaster({ title: 'BASE', background: { color: C.white } });

  [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '1095f3d4-368b-44e8-8a0b-8d12a459d5c7_grok_final.pptx')
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
