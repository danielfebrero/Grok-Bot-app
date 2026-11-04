/**
 * "Tessa Consult" — Strategic Business Consulting deck, rebuilt with pptxgenjs.
 *
 * 20 slides, 13.333in x 7.5in. Raster photos in the original are represented here
 * by flat grey rectangles labelled "[image]"; vector pictograms are stood in for
 * by simple native shapes.
 *
 *   node 19e935f3-6273-4d3b-9160-88d91507d5f5_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  cream: 'FEF2E8', // page tint / accent1
  ice: 'DCF2FA', // pale blue / accent2
  coral: 'FF9D73', // accent3
  pink: 'FFBEF0', // accent4
  sky: 'B2E2F4',
  peach: 'FBD7B9',
  cyan: '75CBEB',
  tan: 'F8B073',
  ink: '191919',
  near: '262626',
  dark: '3F3F3F',
  grey: '7F7F7F',
  mid: 'A5A5A5',
  rule: 'D8D8D8',
  rule2: 'BFBFBF',
  wash: 'F2F2F2',
  white: 'FFFFFF',
  photo: 'CFCFCF',
};

const F = {
  body: 'Poppins',
  medium: 'Poppins Medium',
  semi: 'Poppins SemiBold',
  display: 'Epilogue ExtraBold',
};

// Soft drop shadows used by the cards. pptxgenjs rewrites the shadow object it
// is handed, so each shape needs its own copy.
function shadowCard() { return { type: 'outer', blur: 15, offset: 3, angle: 45, color: '000000', opacity: 0.2 }; }
function shadowWide() { return { type: 'outer', blur: 20, offset: 10, angle: 45, color: '000000', opacity: 0.05 }; }

const DASH = { color: C.rule, width: 1, dashType: 'lgDash' };

// ---------------------------------------------------------------- helpers ---
function text(slide, content, opts) {
  slide.addText(content, Object.assign({ fontFace: F.body, fontSize: 11, color: C.grey, align: 'left', valign: 'top' }, opts));
}

/** Big two-tone section heading: dark words then a coral tail. */
function heading(slide, dark, accent, opts) {
  text(slide, [{ text: dark, options: { color: C.ink } }, { text: accent, options: { color: C.coral } }],
    Object.assign({ fontFace: F.display, fontSize: 36, color: C.ink }, opts));
}

/** Small bold caption above a paragraph. */
function caption(slide, str, opts) {
  text(slide, str, Object.assign({ fontFace: F.semi, fontSize: 12, color: C.ink }, opts));
}

/** Body paragraph at 150% leading. */
function para(slide, str, opts) {
  text(slide, str, Object.assign({ fontSize: 11, color: C.grey, lineSpacingMultiple: 1.5 }, opts));
}

function hRule(slide, x, y, w, opts) {
  slide.addShape('line', { x: x, y: y, w: w, h: 0, line: Object.assign({}, DASH, opts) });
}

function vRule(slide, x, y, h, opts) {
  slide.addShape('line', { x: x, y: y, w: 0, h: h, line: Object.assign({}, DASH, opts) });
}

/** Dashed elbow connector: across, down, across. */
function elbow(slide, x1, y1, x2, y2) {
  const xm = (x1 + x2) / 2;
  hRule(slide, Math.min(x1, xm), y1, Math.abs(xm - x1));
  vRule(slide, xm, Math.min(y1, y2), Math.abs(y2 - y1));
  hRule(slide, Math.min(xm, x2), y2, Math.abs(x2 - xm));
}

/** The deck's recurring "target" marker: translucent ring with a solid core. */
function dot(slide, x, y, color) {
  slide.addShape('ellipse', { x: x, y: y, w: 0.1927, h: 0.1927, fill: { color: color || C.pink, transparency: 65 } });
  slide.addShape('ellipse', { x: x + 0.0347, y: y + 0.0347, w: 0.1233, h: 0.1233, fill: { color: color || C.pink } });
}

/** Stand-in for a photo/raster asset. */
function photo(slide, x, y, w, h) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: C.photo } });
  slide.addText('[image]', { x: x, y: y + h / 2 - 0.2, w: w, h: 0.4, align: 'center', valign: 'middle', fontFace: F.body, fontSize: 12, color: C.white });
}

/** Stand-in for one of the deck's small vector pictograms. */
function pictogram(slide, x, y, w, h, color, shape) {
  slide.addShape(shape || 'donut', { x: x, y: y, w: w, h: h === undefined ? w : h, fill: { color: color } });
}

/** The little up-right arrow used next to "Contact Us" and "Learn More". */
function arrowMark(slide, x, y, size, color) {
  slide.addShape('line', {
    x: x + size * 0.18, y: y + size * 0.2, w: size * 0.62, h: size * 0.62, flipV: true,
    line: { color: color, width: 1.5, endArrowType: 'triangle' },
  });
}

/**
 * Nav bar and footer. Slides 8 and 11 fill the canvas edge-to-edge and drop
 * the footer URL, so callers pass `false` there.
 */
function chrome(slide, withFooter) {
  text(slide, [{ text: 'Tessa ', options: { color: C.ink } }, { text: 'Consult.', options: { color: C.coral } }],
    { x: 0.5723, y: 0.3206, w: 2.0799, h: 0.2693, fontFace: F.semi, fontSize: 10 });
  text(slide, 'Services', { x: 10.5827, y: 0.3206, w: 0.8548, h: 0.2693, fontFace: F.medium, fontSize: 10, color: C.ink, align: 'center' });
  text(slide, 'Contact Us', { x: 11.513, y: 0.3206, w: 1.0573, h: 0.2693, fontFace: F.medium, fontSize: 10, color: C.ink, align: 'center' });
  arrowMark(slide, 12.4977, 0.3306, 0.2425, C.coral);
  if (withFooter !== false) {
    text(slide, 'www.yourgreatsite.com', { x: 0.5745, y: 6.8958, w: 2.084, h: 0.2693, fontFace: F.medium, fontSize: 10, color: C.grey });
  }
}

/** Hairline that separates the nav bar from the body on slides 2-19. */
function topRule(slide) {
  hRule(slide, 0, 0.8871, 13.3333);
}

/** Centred "Infographic Section" heading used by slides 6-15. */
function infographicHeading(slide) {
  text(slide, [{ text: 'Infographic ', options: { color: C.ink } }, { text: 'Section', options: { color: C.coral } }],
    { x: 3.75, y: 1.0663, w: 5.8333, h: 0.7068, fontFace: F.display, fontSize: 36, align: 'center' });
  dot(slide, 2.6522, 0.7907, C.pink);
}

// Lorem strings that recur verbatim throughout the deck.
const L = {
  challenge: 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eius mod Ut. Fusce posuere, magna sed.',
  approach: 'Lorem ipsum dolor sit amet, Fusce posuere, mag sed pulvinar  consectetur purus minim pos.',
  service: 'Lorem ipsum dolor sit amet, Fusce posuere, mag sed pulvin consectetur purus minim pos dolore sed minim dolp.',
  stat: 'Lorem ipsum dolor li amet consetetur',
  quote: '\u201cLorem ipsum dolor sit am, sed do ut labore et dol magna aliqua. Ut\u201d',
  timeline: 'Lorem ipsum dolor sit amet, consectetur',
  year: 'Lorem ipsum dolor sit amet, consectetur adipisicing',
  tag: 'Lorem ipsum dolor sit amet, adipisicing elit, sed laborum.',
  dataValue: 'Lorem ipsum dolor sit amet, elit. Sed vestibule eros eget adipiscing, sed dictum nibh nec',
  pill: 'Lorem ipsum dolor sit amet, elit. Sed vestibule ',
  project: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,',
  banner: 'Lorem ipsum dolor amet, elit. vestibule eros eget amous e sit amet',
  phase: 'Lorem ipsum dolor sit ame, Fusce posuere, magna',
  milestone: 'Lorem ipsum dolor sit amet, Fusce posuere, mag sed pulvinar  consectetur purus minim pos.',
  purusMin: 'Lorem ipsum dolor sit amet, consectetuer purus min.',
};

// ============================================================== SLIDE  1 ====
// Title slide: cream page, big orange arch, two floating stat cards.
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  // Orange arch. The shape is drawn taller than the page so its bottom
  // corner radius falls off-slide, leaving only the rounded top visible.
  s.addShape('roundRect', { x: 4.475, y: 3.8363, w: 4.3834, h: 5.5, rectRadius: 1.8319, fill: { color: C.coral } });

  s.addShape('rect', { x: 6.7604, y: 1.1949, w: 2.9948, h: 0.8415, fill: { color: C.ice }, line: { color: C.cyan, width: 1.75, dashType: 'dash' } });
  text(s, 'Strategic Business', { x: 2.9792, y: 1.1949, w: 7.375, h: 0.8415, fontFace: F.display, fontSize: 44, color: C.ink, align: 'center' });
  text(s, 'Consulting for Growth', { x: 2.9792, y: 2.077, w: 7.375, h: 0.8415, fontFace: F.display, fontSize: 44, color: C.ink, align: 'center' });
  dot(s, 9.6589, 1.1011, C.cyan);

  // Left "35K" card with its dashed leader line.
  elbow(s, -0.2936, 2.5443, 2.6441, 5.4636);
  dot(s, 2.5892, 5.3673, C.pink);
  s.addShape('rect', { x: 3.1349, y: 4.443, w: 2.2113, h: 2.0412, fill: { color: C.white }, shadow: shadowWide() });
  text(s, '35K', { x: 3.4389, y: 4.7949, w: 1.4928, h: 0.5722, fontSize: 28, bold: true, color: C.coral });
  caption(s, 'Financial Market', { x: 3.4389, y: 5.2816, w: 1.7025, h: 0.3029 });
  para(s, L.stat, { x: 3.4406, y: 5.5022, w: 1.7025, h: 0.6301, fontSize: 10.5 });

  // Right pink card + "Learn More" strip.
  elbow(s, 13.3333, 2.2062, 10.8865, 3.4279);
  dot(s, 10.7902, 3.3315, C.pink);
  s.addShape('rect', { x: 9.3811, y: 3.8363, w: 3.0664, h: 1.5657, fill: { color: C.pink }, shadow: shadowWide() });
  caption(s, 'Consulting For Efficiency', { x: 9.7005, y: 4.0651, w: 2.4783, h: 0.3029 });
  para(s, 'Lorem ipsum dolor sit amet, co adipiscing elit. nec posuere purus. Proin minim.',
    { x: 9.7022, y: 4.3176, w: 2.4783, h: 0.8711, fontSize: 10.5, color: C.dark });
  s.addShape('rect', { x: 9.3811, y: 5.5022, w: 3.0664, h: 0.5917, fill: { color: C.white }, shadow: shadowWide() });
  caption(s, 'Learn More', { x: 10.2735, y: 5.6372, w: 1.1861, h: 0.3029 });
  arrowMark(s, 11.3126, 5.666, 0.2425, C.pink);
  s.addShape('triangle', { x: 12.5263, y: 6.2221, w: 0.2218, h: 0.2218, fill: { color: C.ink }, rotate: 200 });

  // Vertical brand tab on the left edge.
  s.addShape('rect', { x: 0.8942, y: 3.19, w: 0.7145, h: 1.6042, fill: { color: C.white }, shadow: shadowWide() });
  text(s, [{ text: 'Tessa ', options: { color: C.ink } }, { text: 'Consult', options: { color: C.sky } }],
    { x: 0.5051, y: 3.849, w: 1.4928, h: 0.2861, fontFace: F.semi, fontSize: 10.5, align: 'center', rotate: -90 });

  chrome(s);
}

// ============================================================== SLIDE  2 ====
// "Current Business Challenges" — two icon/copy blocks plus a photo panel.
function slide02(pres) {
  const s = pres.addSlide();
  topRule(s);
  hRule(s, 0, 4.3387, 6.6667);
  vRule(s, 7.2197, 0, 3.75);

  heading(s, 'Current Business ', 'Challenges', { x: 0.8681, y: 1.3127, w: 4.8333, h: 1.3127 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma porttitor congue massa. Fusce posuere, magna sed pul ultricies, purus lectus malesuada libero.',
    { x: 0.8681, y: 2.7435, w: 4.8333, h: 0.9078 });

  const blocks = [
    { icon: [8.2022, 2.0497, C.sky], cap: [9.2196, 1.7468], body: [9.2196, 2.0069], label: 'Business Challenges 01' },
    { icon: [1.0431, 5.3601, C.pink], cap: [2.0606, 5.0571], body: [2.0606, 5.3172], label: 'Business Challenges 02' },
  ];
  blocks.forEach(function (b) {
    pictogram(s, b.icon[0], b.icon[1], 0.5794, 0.5794, b.icon[2], 'pie');
    caption(s, b.label, { x: b.cap[0], y: b.cap[1], w: 3.0222, h: 0.3029 });
    para(s, L.challenge, { x: b.body[0], y: b.body[1], w: 3.2773, h: 0.9078 });
  });

  dot(s, 7.1233, 0.7824, C.pink);
  photo(s, 6.6667, 3.75, 6.6667, 3.75);
  chrome(s);
}

// ============================================================== SLIDE  3 ====
// "Our Consulting Approach" — stat block left, bulleted copy right.
function slide03(pres) {
  const s = pres.addSlide();
  topRule(s);
  vRule(s, 6.68, 0, 7.5);
  vRule(s, 3.2949, 0.8871, 2.9962);

  heading(s, 'Our Consulting ', 'Approach', { x: 7.9762, y: 1.5447, w: 4.8333, h: 1.3127 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma porttitor congue massa. Fusce posuere, magna sed pul ultricies, purus lectus libero.',
    { x: 7.9762, y: 2.9755, w: 4.3604, h: 0.9078 });

  [{ y: 0, color: C.sky, n: '01' }, { y: 1.2149, color: C.pink, n: '02' }].forEach(function (b) {
    s.addShape('ellipse', { x: 8.1003, y: 4.7589 + b.y, w: 0.1458, h: 0.1458, fill: { color: b.color } });
    caption(s, 'Consulting Approach ' + b.n, { x: 8.4913, y: 4.2734 + b.y, w: 2.6476, h: 0.3753, lineSpacingMultiple: 1.5 });
    para(s, L.approach, { x: 8.4913, y: 4.6182 + b.y, w: 3.8452, h: 0.6061, fontSize: 10.5 });
  });

  dot(s, 6.5836, 0.7942, C.pink);
  pictogram(s, 4.4059, 1.6968, 1.1805, 1.1805, 'E9E9E9', 'donut');
  text(s, '85%', { x: 0.852, y: 1.5725, w: 1.4928, h: 0.7068, fontSize: 36, bold: true, color: C.pink });
  caption(s, 'Financial Market', { x: 0.852, y: 2.2113, w: 1.7025, h: 0.3029 });
  para(s, L.stat, { x: 0.8537, y: 2.4319, w: 1.7025, h: 0.6301, fontSize: 10.5 });
  photo(s, 0, 3.75, 6.68, 2.75);
  chrome(s);
}

// ============================================================== SLIDE  4 ====
// "Our Business Consulting Services" — numbered list left, photos right.
function slide04(pres) {
  const s = pres.addSlide();
  topRule(s);
  vRule(s, 7.2445, 0.8871, 6.6129);

  heading(s, 'Our Business ', 'Consulting Services', { x: 0.9125, y: 1.5447, w: 5.3571, h: 1.3127 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma porttitor congue massa. Fusce posuere, magna sed pul ultricies, purus',
    { x: 0.9125, y: 2.9755, w: 5.507, h: 0.6301 });

  [{ n: '01.', color: C.pink, y: 0 }, { n: '02.', color: C.sky, y: 1.2543 }].forEach(function (b, i) {
    text(s, b.n, { x: 0.9166, y: 4.1229 + b.y, w: 0.8508, h: 0.6496, fontSize: 24, bold: true, color: b.color, lineSpacingMultiple: 1.5 });
    caption(s, 'Consulting Services 0' + (i + 1), { x: 1.7453, y: 3.9974 + b.y, w: 2.6476, h: 0.3753, lineSpacingMultiple: 1.5 });
    para(s, L.service, { x: 1.7453, y: 4.3423 + b.y, w: 4.5244, h: 0.6061, fontSize: 10.5 });
  });

  hRule(s, 7.2445, 4.0022, 6.0888);
  hRule(s, 7.2445, 5.5148, 6.0888);
  dot(s, 7.1481, 3.9067, C.pink);
  text(s, '562K+', { x: 8.1939, y: 4.4365, w: 1.9457, h: 0.7068, fontSize: 36, bold: true, color: C.sky });
  para(s, L.purusMin, { x: 10.1416, y: 4.4015, w: 2.6746, h: 0.6301 });
  photo(s, 7.5484, 1.1394, 5.7849, 2.5608);
  photo(s, 7.5484, 5.7892, 5.7849, 1.7108);
  chrome(s);
}

// ============================================================== SLIDE  5 ====
// "A Success Story" — testimonials, star ratings and two progress bars.
function slide05(pres) {
  const s = pres.addSlide();
  topRule(s);
  vRule(s, 6.746, 0.8871, 6.6129);
  vRule(s, 12.255, 0.8871, 2.8629);
  hRule(s, 6.746, 3.7462, 6.5873);
  hRule(s, 6.746, 2.0637, 5.509);

  const reviews = [
    { name: 'Cecilia Agnaer', x: 0.5723, starX: 0.6899, filled: 4 },
    { name: 'Mark Anderson', x: 3.8659, starX: 3.9836, filled: 5 },
  ];
  reviews.forEach(function (r) {
    caption(s, r.name, { x: r.x, y: 1.4473, w: 1.8834, h: 0.3029, color: C.near });
    para(s, L.quote, { x: r.x, y: 1.7484, w: 2.2944, h: 0.9078, italic: true, color: C.mid });
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', { x: r.starX + i * 0.2857, y: 2.8253, w: 0.2032, h: 0.2032, fill: { color: i < r.filled ? C.sky : 'D9D9D9' } });
    }
  });

  heading(s, 'A Success Story ', 'with Tessa Consult', { x: 7.5042, y: 4.4835, w: 5.3571, h: 1.3127 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma port congue massa. Fusce posuere, magna sed pul ultricies, purus',
    { x: 7.5042, y: 5.8191, w: 5.1764, h: 0.6301 });
  dot(s, 6.6488, 3.6512, C.pink);

  text(s, 'Satisfied Customer', { x: 7.535, y: 2.4294, w: 2.2792, h: 0.3524, fontFace: F.semi, fontSize: 10.5, color: C.ink, lineSpacingMultiple: 1.5 });
  [{ y: 2.8398, w: 2.2475, color: C.coral, pct: '89%' }, { y: 3.1753, w: 1.6163, color: C.pink, pct: '45%' }].forEach(function (b) {
    s.addShape('roundRect', { x: 7.6219, y: b.y, w: 3.1112, h: 0.0931, rectRadius: 0.0466, fill: { color: C.rule, transparency: 40 } });
    s.addShape('roundRect', { x: 7.6238, y: b.y + 0.0003, w: b.w, h: 0.0931, rectRadius: 0.0466, fill: { color: b.color } });
    text(s, b.pct, { x: 10.9132, y: b.y - 0.1477, w: 0.6814, h: 0.3524, fontFace: F.medium, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  });

  text(s, '425K+', { x: 7.4317, y: 1.2176, w: 1.9457, h: 0.6395, fontSize: 32, bold: true, color: C.sky });
  para(s, L.purusMin, { x: 9.3794, y: 1.1742, w: 2.6746, h: 0.6301 });
  photo(s, 0, 3.7222, 3.0167, 2.7778);
  photo(s, 3.4048, 3.7222, 3.0167, 2.7778);
  chrome(s);
}

// ============================================================== SLIDE  6 ====
// Infographic: one horizontal bar chart card plus two "sparkline" cards.
function slide06(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  s.addShape('roundRect', { x: 0.6561, y: 2.1953, w: 5.9077, h: 4.3333, rectRadius: 0.1498, fill: { color: C.white }, shadow: shadowCard() });
  text(s, 'Development Company', { x: 2.294, y: 2.7429, w: 2.6319, h: 0.3366, fontSize: 14, bold: true, color: C.near, align: 'center' });

  const rows = [
    { label: 'Month 1', pct: '87%', w: 3.0039, color: C.peach },
    { label: 'Month 2', pct: '64%', w: 2.4518, color: C.sky },
    { label: 'Month 3', pct: '73%', w: 2.8206, color: C.coral },
    { label: 'Month 4', pct: '57%', w: 2.2227, color: C.pink },
  ];
  rows.forEach(function (r, i) {
    const dy = i * 0.6951;
    text(s, r.label, { x: 0.9117, y: 3.5166 + dy, w: 1.1257, h: 0.3366, fontSize: 14, bold: true, color: C.near });
    s.addShape('rect', { x: 2.1053, y: 3.4741 + dy, w: 3.591, h: 0.4215, fill: { color: C.wash } });
    s.addShape('rect', { x: 2.1053, y: 3.4741 + dy, w: r.w, h: 0.4215, fill: { color: r.color } });
    text(s, r.pct, { x: 5.7642, y: 3.5334 + dy, w: 0.5441, h: 0.3029, fontSize: 12, align: 'right' });
  });

  // Two narrow cards, each holding a stack of rounded bars (rotated upright bars).
  const stacks = [
    {
      card: 6.7931, groups: [
        { label: 'Month 1', y: 2.6026, color: C.peach, bars: [[8.1431, 2.1306, 2.1069], [7.907, 2.6911, 1.6347], [8.0436, 2.8789, 1.9079], [7.98, 3.2669, 1.7809]] },
        { label: 'Month 2', y: 4.4565, color: C.sky, bars: [[7.812, 4.3155, 1.4448], [7.812, 4.6398, 1.4448], [7.98, 4.7962, 1.7809], [7.907, 5.1937, 1.6347]] },
      ],
    },
    {
      card: 9.8497, groups: [
        { label: 'Month 3', y: 2.6026, color: C.coral, bars: [[11.1002, 2.23, 1.9079], [10.9945, 2.6601, 1.6965], [10.8686, 3.1104, 1.4448], [10.9636, 3.3399, 1.6347]] },
        { label: 'Month 4', y: 4.4565, color: C.pink, bars: [[10.7897, 4.3944, 1.2868], [10.7619, 4.7466, 1.2312], [10.8222, 5.0107, 1.3519], [10.9636, 5.1937, 1.6347]] },
      ],
    },
  ];
  stacks.forEach(function (st) {
    s.addShape('roundRect', { x: st.card, y: 2.1953, w: 2.8275, h: 4.3333, rectRadius: 0.1498, fill: { color: C.white }, shadow: shadowCard() });
    st.groups.forEach(function (g) {
      text(s, g.label, { x: st.card + 0.3138, y: g.y, w: 1.4448, h: 0.3366, fontSize: 14, bold: true, color: C.near });
      g.bars.forEach(function (b) {
        s.addShape('round2SameRect', { x: b[0], y: b[1], w: 0.2203, h: b[2], rotate: 90, fill: { color: g.color } });
      });
    });
  });

  chrome(s);
}

// ============================================================== SLIDE  7 ====
// Infographic: grouped column chart plus three "Data Value" cards.
function slide07(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  s.addShape('roundRect', { x: 1.0809, y: 2.0024, w: 5.447, h: 4.7497, rectRadius: 0.1498, fill: { color: C.white }, shadow: shadowCard() });
  text(s, 'Company Value', { x: 2.8741, y: 2.2448, w: 1.8607, h: 0.3366, fontSize: 14, bold: true, color: C.near, align: 'center' });
  s.addShape('line', { x: 1.7461, y: 2.5761, w: 0, h: 3.5687, line: { color: C.rule2, width: 1 } });
  s.addShape('line', { x: 1.7461, y: 6.1396, w: 4.3638, h: 0, line: { color: C.rule2, width: 1 } });

  ['175', '150', '125', '100', '75', '50', '25'].forEach(function (v, i) {
    text(s, v, { x: 1.2461, y: 2.7128 + i * 0.4831, w: 0.5494, h: 0.3029, fontSize: 12, align: 'center' });
  });

  // Each group: [x of first bar, label x, [height, height, height]] — bars grow upward from y=6.1407.
  const groups = [
    { label: 'Data A', labelX: 2.1332, x: 2.0439, h: [1.8793, 2.4545, 1.8793] },
    { label: 'Data B', labelX: 3.5049, x: 3.4156, h: [2.1588, 1.8793, 2.8] },
    { label: 'Data C', labelX: 4.8766, x: 4.7873, h: [2.6419, 2.0708, 2.3208] },
  ];
  const barColors = [C.peach, C.sky, C.coral];
  groups.forEach(function (g) {
    g.h.forEach(function (h, i) {
      s.addShape('round2SameRect', { x: g.x + i * 0.3621, y: 6.1407 - h, w: 0.3008, h: h, fill: { color: barColors[i] } });
    });
    text(s, g.label, { x: g.labelX, y: 6.2625, w: 0.8462, h: 0.3029, fontSize: 12, align: 'center' });
  });

  ['A', 'B', 'C'].forEach(function (letter, i) {
    const y = 2.0024 + i * 1.6815;
    s.addShape('roundRect', { x: 6.8341, y: y, w: 5.4183, h: 1.3868, rectRadius: 0.1992, fill: { color: C.white }, shadow: shadowCard() });
    s.addShape('roundRect', { x: 7.1231, y: y + 0.2456, w: 0.8956, h: 0.8956, rectRadius: 0.1076, fill: { color: C.white }, shadow: shadowCard() });
    pictogram(s, 7.3331, y + 0.4557, 0.4757, 0.4755, C.cream, 'donut');
    text(s, 'Data Value ' + letter, { x: 8.184, y: y + 0.2013, w: 1.6507, h: 0.3366, fontSize: 14, bold: true, color: C.near });
    para(s, L.dataValue, { x: 8.184, y: y + 0.5109, w: 3.7793, h: 0.6301 });
  });

  chrome(s);
}

// ============================================================== SLIDE  8 ====
// Infographic: five overlapping circles connected by chevrons.
function slide08(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  [[2.9364, 3.9371, 45, false], [8.2091, 3.4612, 45, false], [10.006, 3.8568, 135, true], [4.7484, 3.4612, 135, true]].forEach(function (c) {
    s.addShape('chevron', { x: c[0], y: c[1], w: 0.3908, h: 0.459, rectRadius: 0.1263, rotate: c[2], flipH: c[3], fill: { color: C.wash } });
  });

  const nodes = [
    { title: 'First Timeline', big: [0.7604, 1.6012], small: [1.5022, 1.2387], color: C.sky, titleAt: [1.0127, 2.5157, 1.9987], bodyAt: [1.1019, 2.8663], onLight: false, icon: [1.7699, 1.5072, 0.4844, 0.4844, 'donut'] },
    { title: 'Second Timeline', big: [3.0114, 4.2394], small: [3.7532, 3.8769], color: C.coral, titleAt: [3.2003, 5.1539, 2.1257], bodyAt: [3.353, 5.5045], onLight: false, icon: [4.0192, 4.1968, 0.4898, 0.3801, 'cloudCallout'] },
    { title: 'Third Timeline', big: [5.4149, 2.2859], small: [6.1567, 1.9234], color: C.peach, titleAt: [5.7565, 3.2004, 1.8202], bodyAt: [5.7565, 3.551], onLight: true, icon: [6.4467, 2.165, 0.4375, 0.5, 'sun'] },
    { title: 'Fours Timeline', big: [7.8184, 4.241], small: [8.5602, 3.8785], color: C.pink, titleAt: [8.1599, 5.1555, 1.8202], bodyAt: [8.1599, 5.5061], onLight: false, icon: [8.8667, 4.1666, 0.4068, 0.4649, 'flowChartMultidocument'] },
    { title: 'Five Timeline', big: [10.0791, 1.6005], small: [10.8209, 1.2381], color: C.sky, titleAt: [10.4207, 2.515, 1.8202], bodyAt: [10.4207, 2.8657], onLight: false, icon: [11.0983, 1.515, 0.4651, 0.466, 'lightningBolt'] },
  ];
  nodes.forEach(function (n) {
    s.addShape('ellipse', { x: n.big[0], y: n.big[1], w: 2.5034, h: 2.5034, fill: { color: n.color } });
    text(s, n.title, { x: n.titleAt[0], y: n.titleAt[1], w: n.titleAt[2], h: 0.3366, fontSize: 14, bold: true, color: n.onLight ? C.ink : C.white, align: 'center' });
    para(s, L.timeline, { x: n.bodyAt[0], y: n.bodyAt[1], w: 1.8202, h: 0.6301, color: n.onLight ? C.dark : C.white, align: 'center' });
    s.addShape('ellipse', { x: n.small[0], y: n.small[1], w: 1.0198, h: 1.0198, fill: { color: n.color }, line: { color: C.white, width: 2.25 } });
    pictogram(s, n.icon[0], n.icon[1], n.icon[2], n.icon[3], C.white, n.icon[4]);
  });

  chrome(s, false);
}

// ============================================================== SLIDE  9 ====
// Infographic: five ringed icons riding a grey wave.
function slide09(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  // The wave band is five overlapping half-rings sharing one centre line,
  // alternating crest and trough.
  const HUMP_R = 1.215;
  const HUMP_MID = 4.208;
  [2.777, 4.725, 6.669, 8.611, 10.555].forEach(function (cx, i) {
    const crest = i % 2 === 0;
    s.addShape('blockArc', {
      x: cx - HUMP_R, y: HUMP_MID - HUMP_R, w: 2 * HUMP_R, h: 2 * HUMP_R,
      angleRange: crest ? [180, 360] : [0, 180], arcThicknessRatio: 0.4, fill: { color: C.wash },
    });
  });

  const years = [
    { title: 'First Year', ring: [2.1855, 3.6576], color: C.peach, stem: [2.7719, 4.2692], titleAt: [1.61, 5.3925], bodyAt: [1.1837, 5.7651], icon: [2.5397, 4.0584, 0.4369, 0.3495, 'flowChartMultidocument'] },
    { title: 'Second Year', ring: [4.1557, 3.6576], color: C.ice, stem: [4.728, 3.2901], titleAt: [3.6005, 2.0021], bodyAt: [3.1742, 2.3747], icon: [4.5062, 4.0077, 0.4574, 0.4002, 'blockArc'] },
    { title: 'Third Year', ring: [6.105, 3.6576], color: C.coral, stem: [6.668, 4.2331], titleAt: [5.5061, 5.3565], bodyAt: [5.0798, 5.7291], icon: [6.4625, 4.0371, 0.3913, 0.392, 'triangle'] },
    { title: 'Fours Year', ring: [8.0382, 3.6576], color: C.pink, stem: [8.6052, 3.2897], titleAt: [7.4776, 2.0016], bodyAt: [7.0514, 2.3742], icon: [8.3887, 4.0499, 0.4709, 0.314, 'flowChartMagneticDrum'] },
    { title: 'Five Year', ring: [9.995, 3.6601], color: C.peach, stem: [10.5751, 4.1971], titleAt: [9.4132, 5.3205], bodyAt: [8.9869, 5.6931], icon: [10.3376, 4.0584, 0.4338, 0.3471, 'flowChartDisplay'] },
  ];
  years.forEach(function (y) {
    s.addShape('line', { x: y.stem[0], y: y.stem[1], w: 0, h: 0.985, line: { color: y.color, width: 2.25, beginArrowType: 'oval', endArrowType: 'oval' } });
    s.addShape('ellipse', { x: y.ring[0], y: y.ring[1], w: 1.1527, h: 1.1527, fill: { color: y.color }, line: { color: C.white, width: 1.75 } });
    s.addShape('ellipse', { x: y.ring[0] + 0.1357, y: y.ring[1] + 0.1356, w: 0.8814, h: 0.8814, fill: { color: y.color }, line: { color: C.white, width: 1.75 } });
    pictogram(s, y.icon[0], y.icon[1], y.icon[2], y.icon[3], C.white, y.icon[4]);
    text(s, y.title, { x: y.titleAt[0], y: y.titleAt[1], w: 2.2551, h: 0.3366, fontSize: 14, bold: true, color: C.near, align: 'center' });
    para(s, L.year, { x: y.bodyAt[0], y: y.bodyAt[1], w: 3.1076, h: 0.6301, color: C.mid, align: 'center' });
  });

  chrome(s);
}

// ============================================================= SLIDE 10 ====
// Infographic: four "tag" cards with chevron headers.
function slide10(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  const tags = [
    { title: 'First Timeline', x: 1.1973, y: 3.3466, outline: C.cream, color: C.peach, titleW: 1.4592, onLight: true, topIcon: 'rightArrow', botIcon: 'cloudCallout' },
    { title: 'Second Timeline', x: 4.2163, y: 4.5074, outline: C.ice, color: C.sky, titleW: 1.7452, onLight: false, topIcon: 'rightArrow', botIcon: 'star4' },
    { title: 'Third Timeline', x: 7.0688, y: 3.3466, outline: C.coral, color: C.coral, titleW: 1.6303, onLight: false, topIcon: 'rightArrow', botIcon: 'upArrow' },
    { title: 'Fours Timeline', x: 10.0878, y: 4.5074, outline: C.pink, color: C.pink, titleW: 1.6058, onLight: false, topIcon: 'rightArrow', botIcon: 'plus' },
  ];
  tags.forEach(function (t) {
    const bx = t.x + 0.1134; // filled card sits slightly right/above its outline
    const by = t.y - 0.1966;
    s.addShape('roundRect', { x: t.x, y: t.y, w: 1.9834, h: 1.7043, rectRadius: 0.1324, fill: { type: 'none' }, line: { color: t.outline, width: 1.5 } });
    s.addShape('roundRect', { x: bx, y: by, w: 2.0439, h: 1.7777, rectRadius: 0.1338, fill: { color: t.color } });
    s.addShape('chevron', { x: bx, y: by - 0.7888, w: 2.0439, h: 0.5168, rectRadius: 0.167, fill: { color: t.color }, line: { color: C.white, width: 2.25 } });
    s.addShape('ellipse', { x: t.x - 0.4369, y: by - 0.9854, w: 0.9101, h: 0.9101, fill: { color: t.color }, line: { color: C.white, width: 2.25 } });
    s.addShape('ellipse', { x: t.x - 0.2735, y: by - 0.822, w: 0.5832, h: 0.5832, fill: { color: t.color }, line: { color: C.white, width: 2.25 } });
    pictogram(s, t.x - 0.1392, by - 0.6933, 0.3146, 0.3146, C.white, t.topIcon);
    text(s, t.title, { x: t.x + 0.3917, y: by - 0.6986, w: t.titleW, h: 0.3029, fontSize: 12, bold: true, color: t.onLight ? C.ink : C.white });
    para(s, L.tag, { x: t.x + 0.1437, y: t.y + 0.0275, w: 1.9834, h: 0.9078, color: t.onLight ? C.dark : C.white, align: 'center' });
    s.addShape('ellipse', { x: t.x + 0.723, y: t.y + 1.2488, w: 0.815, h: 0.815, fill: { color: t.color }, line: { color: C.white, width: 2.25 } });
    pictogram(s, t.x + 0.9088, t.y + 1.4453, 0.3822, 0.3822, C.white, t.botIcon);
  });

  chrome(s);
}

// ============================================================= SLIDE 11 ====
// Infographic: "Daily Scrum" wheel with six numbered nodes.
function slide11(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  s.addShape('ellipse', { x: 4.2282, y: 2.0428, w: 4.8769, h: 4.8769, fill: { type: 'none' }, line: { color: C.rule2, width: 3.5, dashType: 'dash' } });
  s.addShape('ellipse', { x: 4.7063, y: 2.5209, w: 3.9206, h: 3.9206, fill: { color: C.white }, shadow: shadowCard() });
  pictogram(s, 5.9621, 3.3738, 1.409, 1.4673, C.pink, 'donut');
  text(s, 'Daily Scrum', { x: 5.4416, y: 5.151, w: 2.45, h: 0.4375, fontSize: 20, bold: true, color: C.near, align: 'center' });

  // Six spokes: left column is right-aligned, right column left-aligned.
  const spokes = [
    { n: '03', label: 'Implementation', color: C.coral, side: 'l', labelAt: [1.8753, 2.4258, 1.9518], bodyAt: [1.8753, 2.7624], circle: [4.0153, 2.4573] },
    { n: '02', label: 'Planning', color: C.pink, side: 'l', labelAt: [1.7769, 3.9756, 1.7026], bodyAt: [1.5276, 4.3122], circle: [3.6633, 4.0069] },
    { n: '01', label: 'Sprint Backlog', color: C.coral, side: 'l', labelAt: [2.1246, 5.5255, 1.7026], bodyAt: [1.8753, 5.8621], circle: [4.0104, 5.557] },
    { n: '04', label: 'Review', color: C.pink, side: 'r', labelAt: [9.5061, 2.4258, 1.7026], bodyAt: [9.5061, 2.7624], circle: [8.3698, 2.4573] },
    { n: '05', label: 'Retrospect', color: C.coral, side: 'r', labelAt: [9.8537, 3.9756, 1.7026], bodyAt: [9.8537, 4.3122], circle: [8.7175, 4.0072] },
    { n: '06', label: 'Resolved', color: C.pink, side: 'r', labelAt: [9.5061, 5.5255, 1.7026], bodyAt: [9.5061, 5.8621], circle: [8.3698, 5.557] },
  ];
  spokes.forEach(function (k) {
    const align = k.side === 'l' ? 'right' : 'left';
    text(s, k.label, { x: k.labelAt[0], y: k.labelAt[1], w: k.labelAt[2], h: 0.3366, fontSize: 14, bold: true, color: k.color, align: align });
    para(s, L.timeline, { x: k.bodyAt[0], y: k.bodyAt[1], w: 1.9518, h: 0.6301, align: align });
    s.addShape('ellipse', { x: k.circle[0], y: k.circle[1], w: 0.9481, h: 0.9481, fill: { color: k.color }, shadow: shadowCard() });
    text(s, k.n, { x: k.circle[0] + 0.1856, y: k.circle[1] + 0.2553, w: 0.577, h: 0.4375, fontSize: 20, bold: true, color: C.white, align: 'center' });
  });

  chrome(s, false);
}

// ============================================================= SLIDE 12 ====
// Infographic: six numbered pill cards in two rows.
function slide12(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  // Light grey bracket that threads behind the cards.
  s.addShape('rect', { x: 1.7137, y: 3.0911, w: 9.4894, h: 0.183, fill: { color: C.wash } });
  s.addShape('rect', { x: 1.7137, y: 5.1623, w: 9.4894, h: 0.183, fill: { color: C.wash } });
  s.addShape('rect', { x: 11.0201, y: 3.0911, w: 0.183, h: 2.2542, fill: { color: C.wash } });

  // Top row reads right-to-left (disc on the right); bottom row is mirrored.
  const pills = [
    { n: '01', color: C.peach, card: 1.622, y: 2.6905, disc: 3.6804, textX: 1.7261, numX: 3.9047, align: 'right', numColor: C.ink },
    { n: '02', color: C.sky, card: 5.0398, y: 2.6905, disc: 7.0982, textX: 5.144, numX: 7.3226, align: 'right', numColor: C.white },
    { n: '03', color: C.coral, card: 8.4577, y: 2.6905, disc: 10.5161, textX: 8.5619, numX: 10.7404, align: 'right', numColor: C.white },
    { n: '04', color: C.pink, card: 8.616, y: 4.7619, disc: 8.4577, textX: 9.5926, numX: 8.8654, align: 'left', numColor: C.white, mirror: true },
    { n: '05', color: C.peach, card: 5.1982, y: 4.7619, disc: 5.0398, textX: 6.1748, numX: 5.4476, align: 'left', numColor: C.ink, mirror: true },
    { n: '06', color: C.coral, card: 1.7803, y: 4.7619, disc: 1.622, textX: 2.7569, numX: 2.0297, align: 'left', numColor: C.white, mirror: true },
  ];
  pills.forEach(function (p) {
    const discY = p.y - 0.1055;
    s.addShape('ellipse', { x: p.disc, y: discY, w: 1.1951, h: 1.1951, fill: { color: p.color }, shadow: shadowCard() });
    s.addShape('roundRect', { x: p.card, y: p.y, w: 3.0952, h: 0.9841, rectRadius: 0.492, fill: { color: C.white }, shadow: shadowCard() });
    s.addShape('ellipse', { x: p.disc + (p.mirror ? 0.2638 : 0.0804), y: discY + 0.1722, w: 0.8509, h: 0.8509, fill: { color: p.color } });
    para(s, L.pill, { x: p.textX, y: p.y + 0.1547, w: 2.0144, h: 0.6563, align: p.align });
    text(s, p.n, { x: p.numX, y: p.y + 0.2901, w: 0.563, h: 0.4039, fontSize: 18, bold: true, color: p.numColor, align: 'center' });
  });

  chrome(s);
}

// ============================================================= SLIDE 13 ====
// Infographic: pinwheel of four diamonds with four caption cards.
function slide13(pres) {
  const s = pres.addSlide();
  topRule(s);
  text(s, [{ text: 'Infographic ', options: { color: C.ink } }, { text: 'Section', options: { color: C.coral } }],
    { x: 3.75, y: 1.0663, w: 5.8333, h: 0.7068, fontFace: F.display, fontSize: 36, align: 'center' });
  dot(s, 2.6522, 0.7907, C.pink);

  // Four large diamonds (squares turned 45 degrees) around a soft grey hub.
  [[5.4926, 4.3095, C.sky], [6.6521, 3.121, C.peach], [6.6812, 5.469, C.pink], [7.8407, 4.2805, C.coral]].forEach(function (d) {
    s.addShape('diamond', { x: d[0] - 1.144, y: d[1] - 1.144, w: 2.288, h: 2.288, fill: { color: d[2] }, shadow: shadowCard() });
  });
  [[6.0874, 4.3098, 'F7FCFD'], [6.6672, 3.7156, 'FFEBE3'], [6.6817, 4.8896, C.white], [7.2614, 4.2953, 'FFF2FB']].forEach(function (d) {
    s.addShape('diamond', { x: d[0] - 0.5496, y: d[1] - 0.5496, w: 1.0992, h: 1.0992, fill: { color: d[2] }, shadow: shadowCard() });
  });
  s.addShape('ellipse', { x: 6.1802, y: 3.8086, w: 0.9727, h: 0.9727, fill: { color: C.wash } });

  pictogram(s, 6.505, 2.5145, 0.2941, 0.3792, C.white, 'flowChartInternalStorage');
  pictogram(s, 4.8764, 4.1658, 0.3643, 0.2873, C.white, 'chartX');
  pictogram(s, 8.1528, 4.1088, 0.3527, 0.3432, C.white, 'pie');
  pictogram(s, 6.5034, 5.7554, 0.3263, 0.3036, C.white, 'smileyFace');
  pictogram(s, 6.4825, 4.1067, 0.3682, 0.3765, C.cream, 'lightningBolt');

  const cards = [
    { label: 'Project 03', color: C.tan, x: 9.4456, y: 2.3387, textX: 9.6055, align: 'left' },
    { label: 'Project 04', color: C.coral, x: 9.4456, y: 4.8488, textX: 9.6055, align: 'left' },
    { label: 'Project 02', color: C.sky, x: 1.1998, y: 2.3387, textX: 1.3597, align: 'right' },
    { label: 'Project 01', color: C.pink, x: 1.1998, y: 4.8488, textX: 1.3597, align: 'right' },
  ];
  cards.forEach(function (c) {
    s.addShape('roundRect', { x: c.x, y: c.y, w: 2.6878, h: 1.4023, rectRadius: 0.1226, fill: { color: C.white }, shadow: shadowCard() });
    text(s, c.label, { x: c.align === 'left' ? 9.6055 : 2.4895, y: c.y + 0.2069, w: 1.2381, h: 0.3366, fontSize: 14, bold: true, color: c.color, align: c.align });
    para(s, L.project, { x: c.textX, y: c.y + 0.521, w: 2.368, h: 0.6301, align: c.align });
  });

  chrome(s);
}

// ============================================================= SLIDE 14 ====
// Infographic: stacked triangle with four side labels.
function slide14(pres) {
  const s = pres.addSlide();
  topRule(s);
  text(s, [{ text: 'Infographic ', options: { color: C.ink } }, { text: 'Section', options: { color: C.coral } }],
    { x: 3.75, y: 1.0663, w: 5.8333, h: 0.7068, fontFace: F.display, fontSize: 36, align: 'center' });
  dot(s, 2.6522, 0.7907, C.pink);

  // Apex up, except the middle of the bottom row which points down.
  [[5.726, 2.3681, C.coral, false], [4.7218, 4.3129, C.peach, false],
    [5.726, 4.3129, C.sky, true], [6.7301, 4.3129, C.pink, false]].forEach(function (t) {
    s.addShape('triangle', { x: t[0], y: t[1], w: 1.8812, h: 1.8812, fill: { color: t[2] }, flipV: t[3] });
  });
  pictogram(s, 6.3889, 3.3006, 0.5553, 0.4607, C.white, 'smileyFace');
  pictogram(s, 5.4451, 5.3277, 0.4348, 0.4348, C.white, 'flowChartDocument');
  pictogram(s, 6.4322, 4.8863, 0.4515, 0.4515, C.white, 'chartPlus');
  pictogram(s, 7.458, 5.3277, 0.4255, 0.4255, C.white, 'flowChartPunchedTape');

  const labels = [
    { title: 'Information 02', color: C.sky, side: 'l', y: 2.8923, bar: C.sky },
    { title: 'Information 01', color: C.tan, side: 'l', y: 4.6813, bar: C.peach },
    { title: 'Information 03', color: C.coral, side: 'r', y: 2.8923, bar: C.coral },
    { title: 'Information 04', color: C.pink, side: 'r', y: 4.6813, bar: C.pink },
  ];
  labels.forEach(function (b) {
    const left = b.side === 'l';
    text(s, b.title, { x: left ? 1.9685 : 9.3648, y: b.y, w: 1.9999, h: 0.3366, fontSize: 14, bold: true, color: b.color, align: left ? 'right' : 'left' });
    para(s, L.project, { x: left ? 1.6004 : 9.3648, y: b.y + 0.3141, w: 2.368, h: 0.6301, align: left ? 'right' : 'left' });
    s.addShape('roundRect', { x: left ? 4.0234 : 9.2598, y: b.y + 0.0939, w: 0.0499, h: 0.8188, rectRadius: 0.0083, fill: { color: b.bar } });
  });

  chrome(s);
}

// ============================================================= SLIDE 15 ====
// Infographic: five chevron banners over pointed cards.
function slide15(pres) {
  const s = pres.addSlide();
  topRule(s);
  infographicHeading(s);

  const cols = [
    { x: 1.7489, color: C.peach, onLight: true },
    { x: 3.6517, color: C.sky, onLight: false },
    { x: 5.5544, color: C.coral, onLight: false },
    { x: 7.4572, color: C.pink, onLight: false },
    { x: 9.36, color: C.peach, onLight: true },
  ];
  const glyphs = ['flowChartDocument', 'smileyFace', 'gear6', 'star4', 'pie'];
  cols.forEach(function (c, i) {
    s.addShape('homePlate', { x: c.x - 0.2221, y: 3.6447, w: 2.2352, h: 1.791, rectRadius: 0.6455, rotate: 90, fill: { color: c.color } });
    s.addShape('chevron', { x: c.x, y: 2.6233, w: 2.2243, h: 0.7068, rectRadius: 0.3468, fill: { color: c.color }, line: { color: C.white, width: 1 } });
    pictogram(s, c.x + 0.9587, 2.8083, 0.3368, 0.3368, C.white, glyphs[i]);
    text(s, 'Title Here', { x: c.x + 0.2678, y: 3.617, w: 1.2553, h: 0.3365, fontSize: 14, color: c.onLight ? C.ink : C.white, align: 'center' });
    para(s, L.banner, { x: c.x, y: 3.883, w: 1.791, h: 1.1361, fontSize: 10.5, color: c.onLight ? C.dark : C.white, align: 'center' });
  });

  chrome(s);
}

// ============================================================= SLIDE 16 ====
// "Why We Stand Out" — one highlighted advantage plus two plain ones.
function slide16(pres) {
  const s = pres.addSlide();
  topRule(s);
  vRule(s, 7.0634, 3.7462, 3.7537);
  hRule(s, -0.0317, 3.7462, 13.365);
  hRule(s, -0.0317, 6.1074, 7.0793);

  heading(s, 'Why We ', 'Stand Out', { x: 0.8315, y: 1.5007, w: 5.3571, h: 0.7068 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma port congue massa. Fusce posuere, magna sed pul ultricies, pur Lorem ipsum dolor sit amet, consectetuer adipiscing elit.',
    { x: 0.8315, y: 2.27, w: 5.1763, h: 0.9078 });

  s.addShape('rect', { x: 0, y: 3.9597, w: 6.8095, h: 1.8875, fill: { color: C.cream }, shadow: shadowWide() });
  text(s, '01.', { x: 0.8315, y: 4.5176, w: 0.8508, h: 0.6496, fontSize: 24, bold: true, color: C.pink, lineSpacingMultiple: 1.5 });
  caption(s, 'Competitive Advantage 01', { x: 1.6602, y: 4.2344, w: 2.6475, h: 0.3752, lineSpacingMultiple: 1.5 });
  para(s, 'Lorem ipsum dolor sit amet, Fusce posuere, mag sed pulvin consectetur purus minim pos dolore sed minim dolp. magna sed pul ultricies, pur Lorem ipsum.',
    { x: 1.6602, y: 4.5793, w: 4.5243, h: 0.8711, fontSize: 10.5 });

  [{ n: '02.', color: C.sky, y: 0 }, { n: '03.', color: C.pink, y: 1.3654 }].forEach(function (b, i) {
    text(s, b.n, { x: 7.7661, y: 4.3599 + b.y, w: 0.8508, h: 0.6496, fontSize: 24, bold: true, color: b.color, lineSpacingMultiple: 1.5 });
    caption(s, 'Competitive Advantage 0' + (i + 2), { x: 8.5947, y: 4.2344 + b.y, w: 2.6475, h: 0.3752, lineSpacingMultiple: 1.5 });
    para(s, 'Lorem ipsum dolor sit amet, Fusce posuere, mag sed pulvin consectetur purus minim pos dolore sed',
      { x: 8.5947, y: 4.5793 + b.y, w: 4.0969, h: 0.606, fontSize: 10.5 });
  });

  dot(s, 6.9668, 6.0111, C.pink);
  photo(s, 7.0476, 1.0595, 6.2857, 2.4824);
  chrome(s);
}

// ============================================================= SLIDE 17 ====
// "Project Timeline & Milestones" — three phases on a rule, two milestones.
function slide17(pres) {
  const s = pres.addSlide();
  topRule(s);
  vRule(s, 6.4843, 3.4215, 4.0784);
  hRule(s, 0, 3.4215, 13.3333);
  hRule(s, 0.8253, 2.857, 11.8412, { color: C.mid });

  [{ x: 0.8253, n: '01', color: C.pink }, { x: 5.5723, n: '02', color: C.sky }, { x: 10.3193, n: '03', color: C.pink }].forEach(function (p) {
    caption(s, 'Our Phase ' + p.n, { x: p.x, y: 1.2964, w: 2.1105, h: 0.3752, align: 'center', lineSpacingMultiple: 1.5 });
    para(s, L.phase, { x: p.x, y: 1.6413, w: 2.1885, h: 0.606, fontSize: 10.5, align: 'center' });
    vRule(s, p.x + 1.0852, 2.4685, 0.3884, { color: C.mid });
    dot(s, p.x + 0.9889, 2.4237, p.color);
  });

  [{ n: '01', color: C.sky, y: 0 }, { n: '02', color: C.pink, y: 1.2839 }].forEach(function (m) {
    s.addShape('ellipse', { x: 0.8253, y: 4.5396 + m.y, w: 0.1458, h: 0.1458, fill: { color: m.color } });
    caption(s, 'Our Milestone ' + m.n, { x: 1.2164, y: 4.054 + m.y, w: 2.6475, h: 0.3752, lineSpacingMultiple: 1.5 });
    para(s, L.milestone, { x: 1.2164, y: 4.3989 + m.y, w: 3.8452, h: 0.606, fontSize: 10.5 });
  });

  heading(s, 'Project Timeline ', '& Milestones', { x: 7.774, y: 4.2285, w: 4.7973, h: 1.3126 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma port congue massa. Fusce posuere, magna sed pu ultricies, purus minim purus sed dolor.',
    { x: 7.774, y: 5.564, w: 4.6703, h: 0.9078 });
  chrome(s);
}

// ============================================================= SLIDE 18 ====
// "Meet Our Best Team" — three portrait cards.
function slide18(pres) {
  const s = pres.addSlide();
  topRule(s);
  hRule(s, 0, 3.0344, 13.3333);
  vRule(s, 4.2921, 3.0344, 4.4655);
  vRule(s, 9.0287, 3.0344, 4.4655);

  text(s, [{ text: 'Meet Our ', options: { color: C.ink } }, { text: 'Best Team', options: { color: C.coral } }],
    { x: 3.8172, y: 1.1843, w: 5.6989, h: 0.7068, fontFace: F.display, fontSize: 36, align: 'center' });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma port congue massa. Fusce posuere, magna sed pu ultricies.',
    { x: 3.8291, y: 1.9701, w: 5.548, h: 0.6301, align: 'center' });

  const team = [
    { x: 0.6666, name: 'Serena Savila', color: C.pink },
    { x: 5.3908, name: 'Mikes  Wazows', color: C.sky },
    { x: 10.1397, name: 'Sulli Vanmor', color: C.pink },
  ];
  team.forEach(function (m) {
    photo(s, m.x, 3.3496, 2.5268, 3.1503);
    s.addShape('rect', { x: m.x, y: 5.654, w: 2.5268, h: 0.8459, fill: { color: m.color }, shadow: shadowWide() });
    text(s, m.name, { x: m.x + 0.4013, y: 5.7272, w: 1.7242, h: 0.3752, fontFace: F.medium, fontSize: 12, color: C.ink, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, 'Co-Founder', { x: m.x + 0.5194, y: 6.0184, w: 1.4881, h: 0.3524, fontSize: 10.5, color: C.dark, align: 'center', lineSpacingMultiple: 1.5 });
  });

  dot(s, 4.1921, 2.9381, C.pink);
  dot(s, 8.9323, 5.7956, C.pink);
  chrome(s);
}

// ============================================================= SLIDE 19 ====
// "Stay Connected With Us" — contact tiles left, full-height photo right.
function slide19(pres) {
  const s = pres.addSlide();
  topRule(s);
  hRule(s, 0, 4.0674, 7.121);

  heading(s, 'Stay Connected ', 'With Us', { x: 0.7638, y: 1.3311, w: 4.7973, h: 1.3126 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma port congue massa. Fusce posuere, magna sed ultricies, purus minim purus sed dolor.',
    { x: 0.7638, y: 2.6508, w: 4.6703, h: 0.9078 });

  const contacts = [
    { tile: [0.73, 4.5131], color: C.pink, glyph: 'teardrop', label: '123 Anywhere St., Any City', at: [1.399, 4.6313, 2.3224] },
    { tile: [0.73, 5.6604], color: C.ice, glyph: 'moon', label: '+123-456-7890', at: [1.399, 5.7744, 2.0433] },
    { tile: [3.8967, 4.5131], color: C.ice, glyph: 'donut', label: 'www.yourgreatsite.com', at: [4.4865, 4.6313, 2.0221] },
    { tile: [3.8967, 5.6604], color: C.pink, glyph: 'flowChartDocument', label: 'info@yourgreatsite', at: [4.4865, 5.7744, 1.7202] },
  ];
  contacts.forEach(function (c) {
    s.addShape('rect', { x: c.tile[0], y: c.tile[1], w: 0.5139, h: 0.5139, fill: { color: c.color }, shadow: shadowWide() });
    pictogram(s, c.tile[0] + 0.1309, c.tile[1] + 0.1309, 0.2522, 0.2522, C.white, c.glyph);
    text(s, c.label, { x: c.at[0], y: c.at[1], w: c.at[2], h: 0.2776, fontFace: F.medium, fontSize: 10.5, color: C.mid });
  });

  dot(s, 6.2844, 3.9711, C.pink);
  photo(s, 7.121, 0.887, 6.2122, 6.6129);
  chrome(s);
}

// ============================================================= SLIDE 20 ====
// Closing slide: mirrors the cover with a "Grateful For / Your Attention" lockup.
function slide20(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  // Orange arch, again drawn past the page bottom so only the top curve shows.
  s.addShape('roundRect', { x: 7.8225, y: 2.1774, w: 4.844, h: 7.745, rectRadius: 2.422, fill: { color: C.coral } });

  s.addShape('rect', { x: 0.6666, y: 1.5, w: 4.3834, h: 0.8414, fill: { color: C.ice }, line: { color: C.cyan, width: 1.75, dashType: 'dash' } });
  text(s, 'Grateful For', { x: 0.7592, y: 1.5208, w: 5.0782, h: 0.8414, fontFace: F.display, fontSize: 44, color: C.ink });
  text(s, 'Your Attention', { x: 0.7592, y: 2.403, w: 5.0782, h: 0.8414, fontFace: F.display, fontSize: 44, color: C.ink });
  dot(s, 4.9537, 1.4063, C.cyan);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ma por congue massa. Fusce posuere, magna sed ultricies, purus minim purus sed dolor consectetuer adipiscing.',
    { x: 0.7592, y: 3.2938, w: 5.0782, h: 0.9078 });

  hRule(s, -0.6299, 5.5605, 1.2965);
  dot(s, 0.537, 5.4672, C.pink);
  s.addShape('rect', { x: 1.161, y: 4.6388, w: 2.2112, h: 1.8453, fill: { color: C.white }, shadow: shadowWide() });
  text(s, '35K', { x: 1.465, y: 4.8643, w: 1.4927, h: 0.5722, fontSize: 28, bold: true, color: C.coral });
  caption(s, 'Financial Market', { x: 1.465, y: 5.351, w: 1.7024, h: 0.3029 });
  para(s, L.stat, { x: 1.4667, y: 5.5716, w: 1.7024, h: 0.6301, fontSize: 10.5 });
  photo(s, 3.6195, 4.6388, 2.2112, 1.8453);

  s.addShape('rect', { x: 6.907, y: 3.7816, w: 2.1527, h: 0.6641, fill: { color: C.pink }, shadow: shadowWide() });
  caption(s, 'End Presentation', { x: 7.1321, y: 3.9623, w: 1.7024, h: 0.3029, align: 'center' });
  s.addShape('rect', { x: 8.2278, y: 2.1774, w: 0.7366, h: 0.7366, fill: { color: C.white }, shadow: shadowWide() });
  pictogram(s, 8.4461, 2.3957, 0.3, 0.3, C.sky, 'flowChartConnector');

  chrome(s);
}

// ------------------------------------------------------------------ main ---
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'TESSA_16x9', width: 13.3333, height: 7.5 });
  pres.layout = 'TESSA_16x9';
  pres.author = 'Tessa Consult';
  pres.title = 'Strategic Business Consulting for Growth';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pres); });

  return pres.writeFile({ fileName: path.join(__dirname, '19e935f3-6273-4d3b-9160-88d91507d5f5_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
