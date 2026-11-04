/*
 * Waregh - Restaurant And Bar Presentation Template
 * Standalone pptxgenjs recreation of the 25-slide reference deck (13.333in x 7.5in).
 *
 * Raster photography in the original is replaced by programmatic grey placeholders
 * (`photo()`); vector food/contact icons are rebuilt from native pptxgenjs shapes.
 *
 * Run: node 04eec701-b9bf-43f1-a89d-364cf0ccfae6_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const CREAM = 'F4F3F1'; // page tint / reversed-out ink
const NAVY = '28344E'; // cards, icon plates
const CLAY = 'AA7565'; // accent dots, contact badges, frame stroke
const WHITE = 'FFFFFF';
const INK = '000000';
const GREY = '7F7F7F'; // body copy
const PHOTO = 'CCCCCC'; // image placeholder tone
const PHOTO_DK = '999999'; // darker image placeholder tone
const PHOTO_X = 'DDDDDD'; // faint cross drawn over placeholders

// ------------------------------------------------------------------ fonts
const SERIF = 'Yeseva One'; // display headlines
const LIGHT = 'Josefin Sans Light'; // eyebrows and labels
const XLIGHT = 'Josefin Sans ExtraLight'; // body copy
const MED = 'Josefin Sans Medium'; // web address / emphasis

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const TEXT_INSET = [7.2, 7.2, 3.6, 3.6]; // l, r, b, t in points - matches the source deck

// ------------------------------------------------------------- primitives
function text(s, x, y, w, h, str, o) {
  o = o || {};
  s.addText(str, {
    x: x, y: y, w: w, h: h,
    fontFace: o.font || LIGHT,
    // `shrink` mirrors the source deck's autofit fontScale on a handful of tight boxes
    fontSize: (o.size || 15) * (o.shrink || 1),
    color: o.color || INK,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.lineSpacing,
    margin: TEXT_INSET,
  });
}

function headline(s, x, y, w, h, str, size) {
  text(s, x, y, w, h, str, { font: SERIF, size: size || 50, color: INK });
}

function eyebrow(s, x, y, w, h, str) {
  text(s, x, y, w, h, str, { font: LIGHT, size: 20, color: INK });
}

function label(s, x, y, w, h, str, o) {
  o = o || {};
  text(s, x, y, w, h, str, {
    font: LIGHT, size: 15, color: o.color || INK, align: o.align, shrink: o.shrink,
  });
}

function body(s, x, y, w, h, str, o) {
  o = o || {};
  text(s, x, y, w, h, str, {
    font: XLIGHT, size: 10, color: o.color || GREY,
    align: o.align || 'justify', lineSpacing: 1.5, shrink: o.shrink,
  });
}

/** Full-bleed or partial cream tint inherited from the source deck's layouts. */
function band(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: CREAM } });
}

/** Rounded panel: the navy stat/staff/service cards (and a few white ones). */
function card(s, x, y, w, h, r, color, geom) {
  s.addShape(geom || 'roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: r, fill: { color: color || NAVY },
  });
}

/** 0.157in clay bullet used beside labels throughout the deck. */
function dot(s, x, y) {
  s.addShape('ellipse', { x: x, y: y, w: 0.157, h: 0.157, fill: { color: CLAY } });
}

function rule(s, x, y, w) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: NAVY, width: 1 } });
}

/** Grey stand-in for a photograph, with the reference art's faint diagonal cross. */
function photo(s, x, y, w, h, geom, r, tone) {
  s.addShape(geom || 'rect', {
    x: x, y: y, w: w, h: h, rectRadius: r || undefined, fill: { color: tone || PHOTO },
  });
  if (geom !== 'ellipse') {
    s.addShape('line', { x: x, y: y, w: w, h: h, line: { color: PHOTO_X, width: 0.75 } });
    s.addShape('line', { x: x, y: y, w: w, h: h, flipV: true, line: { color: PHOTO_X, width: 0.75 } });
  }
}

/** Slide 18's photo is one path cut into four staggered vertical bars. */
function photoBars(s, x, y, w, h) {
  const bw = w * 0.2310;
  const gap = w * 0.2564;
  const inset = h * 0.1009;
  [
    [0, inset],
    [gap, 0],
    [gap * 2, 0],
    [gap * 3, inset],
  ].forEach(function (b) {
    photo(s, x + b[0], y + b[1], bw, h - b[1] * 2, 'roundRect', 0.159);
  });
}

/** Tablet device mock-up (slide 22). */
function tabletMockup(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.28, fill: { color: '2A2A2C' } });
  s.addShape('rect', {
    x: x + w * 0.064, y: y + h * 0.106, w: w * 0.864, h: h * 0.783, fill: { color: WHITE },
  });
  s.addShape('ellipse', {
    x: x + w * 0.455, y: y + h * 0.935, w: w * 0.09, h: w * 0.09, fill: { color: '3C3C3E' },
  });
}

/** Laptop device mock-up (slide 23). */
function laptopMockup(s, x, y, w, h) {
  s.addShape('roundRect', {
    x: x + w * 0.096, y: y, w: w * 0.808, h: h * 0.918, rectRadius: 0.06, fill: { color: '1B1B1D' },
  });
  s.addShape('rect', {
    x: x + w * 0.118, y: y + h * 0.049, w: w * 0.766, h: h * 0.838, fill: { color: WHITE },
  });
  s.addShape('roundRect', {
    x: x, y: y + h * 0.918, w: w, h: h * 0.075, rectRadius: 0.03, fill: { color: 'C9CACC' },
  });
}

/** Full cream page plus the corner rules and bead trios of the cover / break / closing slides. */
function ornateFrame(s) {
  band(s, 0, 0, SLIDE_W, SLIDE_H);
  const stroke = { color: CLAY, width: 1 };
  const corner = function (x, y, w, h, flip) {
    s.addShape('custGeom', {
      x: x, y: y, w: w, h: h, flipH: flip, flipV: flip, line: stroke,
      points: [
        { x: 0, y: h },
        { x: 0, y: h * 0.1119 },
        { curve: { type: 'cubic', x1: w * 0.0395, y1: h * 0.1119, x2: w * 0.0725, y2: h * 0.0723 }, x: w * 0.0802, y: h * 0.0198 },
        { x: w * 0.0816, y: 0 },
        { x: w, y: h * 0.0014 },
      ],
    });
  };
  const beads = function (x, y, size, dx, dy) {
    for (let i = 0; i < 3; i++) {
      s.addShape('ellipse', {
        x: x + dx * i, y: y + dy * i, w: size, h: size,
        fill: { color: CREAM }, line: stroke,
      });
    }
  };
  corner(0.703, 0.703, 6.044, 4.289, false);
  corner(7.378, 3.073, 5.262, 3.734, true);
  beads(5.704, 0.63, 0.147, 0.3133, 0);
  beads(0.63, 3.739, 0.147, 0, 0.3133);
  beads(7.613, 6.742, 0.128, 0.2728, 0);
  beads(12.576, 3.49, 0.128, 0, 0.2728);
}

// ------------------------------------------------------------------ icons
/** Rounded bar centred on (cx, cy) and spun about its own centre. */
function bar(s, cx, cy, w, h, rot, color) {
  s.addShape('roundRect', {
    x: cx - w / 2, y: cy - h / 2, w: w, h: h, rotate: rot,
    rectRadius: Math.min(w, h) / 2, fill: { color: color },
  });
}

/** Closed polygon from unit-square coordinates mapped into the icon box. */
function poly(s, x, y, w, h, pts, color) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: color },
    points: pts.map(function (p) { return { x: p[0] * w, y: p[1] * h }; }).concat([{ close: true }]),
  });
}

const icon = {
  // Plate of spaghetti - slides 12, 13, 14, 15, 21
  pasta: function (s, x, y, w, h, color, bg) {
    bg = bg || NAVY;
    const t = h * 0.055;
    [[0.30, -14], [0.47, 0], [0.64, 14]].forEach(function (steam) {
      bar(s, x + w * steam[0], y + h * 0.13, w * 0.05, h * 0.19, steam[1], color);
    });
    s.addShape('ellipse', { x: x + w * 0.08, y: y + h * 0.34, w: w * 0.70, h: h * 0.34, fill: { color: color } });
    s.addShape('ellipse', { x: x + w * 0.60, y: y + h * 0.42, w: w * 0.26, h: h * 0.26, fill: { color: color } });
    s.addShape('ellipse', { x: x + w * 0.06, y: y + h * 0.50, w: w * 0.88, h: h * 0.24, fill: { color: color } });
    // two nested grooves in the mound read as a coiled strand
    s.addShape('ellipse', { x: x + w * 0.13, y: y + h * 0.40, w: w * 0.40, h: h * 0.30, fill: { color: bg } });
    s.addShape('ellipse', { x: x + w * 0.16, y: y + h * 0.42, w: w * 0.34, h: h * 0.26, fill: { color: color } });
    s.addShape('ellipse', { x: x + w * 0.22, y: y + h * 0.48, w: w * 0.22, h: h * 0.16, fill: { color: bg } });
    s.addShape('ellipse', { x: x + w * 0.25, y: y + h * 0.50, w: w * 0.16, h: h * 0.12, fill: { color: color } });
    // plate: outlined ellipse so the pasta shows through, hollowed at the back
    s.addShape('ellipse', {
      x: x, y: y + h * 0.56, w: w, h: h * 0.40,
      fill: { type: 'none' }, line: { color: color, width: t * 72 },
    });
    s.addShape('blockArc', {
      x: x + t, y: y + h * 0.56 + t, w: w - 2 * t, h: h * 0.40 - 2 * t,
      angleRange: [0, 180], arcThicknessRatio: 0.30, fill: { color: bg },
    });
  },
  // Wine glass - slides 12, 13, 15
  wine: function (s, x, y, w, h, color, bg) {
    poly(s, x, y, w, h, [[0.06, 0.06], [0.94, 0.06], [0.84, 0.46], [0.16, 0.46]], color); // bowl
    s.addShape('ellipse', { x: x + w * 0.16, y: y + h * 0.30, w: w * 0.68, h: h * 0.32, fill: { color: color } });
    s.addShape('rect', { x: x + w * 0.15, y: y + h * 0.155, w: w * 0.70, h: h * 0.075, fill: { color: bg || NAVY } });
    s.addShape('rect', { x: x + w * 0.44, y: y + h * 0.50, w: w * 0.12, h: h * 0.38, fill: { color: color } }); // stem
    s.addShape('ellipse', { x: x + w * 0.18, y: y + h * 0.84, w: w * 0.64, h: h * 0.12, fill: { color: color } }); // foot
  },
  // Slice of layer cake - slides 12, 13, 15
  cake: function (s, x, y, w, h, color, bg) {
    poly(s, x, y, w, h, [[0.03, 0.24], [0.62, 0.06], [0.95, 0.46], [0.95, 0.92], [0.40, 1.0], [0.03, 0.80]], color);
    poly(s, x, y, w, h, [[0.07, 0.38], [0.80, 0.60], [0.80, 0.71], [0.07, 0.51]], bg || NAVY);
    poly(s, x, y, w, h, [[0.07, 0.60], [0.80, 0.82], [0.80, 0.93], [0.07, 0.73]], bg || NAVY);
  },
  // Crossed fork and knife - slides 2, 22
  cutlery: function (s, x, y, w, h, color) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const d = 0.707;
    bar(s, cx, cy, w * 0.11, h * 0.92, 45, color); // knife handle
    bar(s, cx + w * 0.24 * d, cy - h * 0.24 * d, w * 0.24, h * 0.42, 45, color); // blade
    bar(s, cx, cy, w * 0.11, h * 0.92, -45, color); // fork handle
    [-1, 0, 1].forEach(function (k) {
      bar(s, cx - w * (0.30 - k * 0.10) * d, cy - h * (0.30 + k * 0.10) * d, w * 0.07, h * 0.30, -45, color);
    });
  },
  // Desk telephone in a clay badge - slide 24
  phone: function (s, x, y, w, h, color, bg) {
    s.addShape('blockArc', {
      x: x + w * 0.02, y: y, w: w * 0.96, h: h * 0.92,
      angleRange: [180, 360], arcThicknessRatio: 0.28, fill: { color: color },
    }); // handset
    poly(s, x, y, w, h, [[0.16, 1.0], [0.16, 0.62], [0.32, 0.42], [0.68, 0.42], [0.84, 0.62], [0.84, 1.0]], color);
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 3; c++) {
        s.addShape('rect', {
          x: x + w * (0.34 + c * 0.13), y: y + h * (0.58 + r * 0.13),
          w: w * 0.08, h: h * 0.09, fill: { color: bg || CLAY },
        }); // keypad
      }
    }
  },
  // Push pin in a clay badge - slide 24
  pin: function (s, x, y, w, h, color) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    bar(s, cx - w * 0.19, cy - h * 0.19, w * 0.46, h * 0.24, 45, color); // head
    bar(s, cx + w * 0.05, cy + h * 0.05, w * 0.30, h * 0.50, 45, color); // body
    bar(s, cx + w * 0.29, cy + h * 0.29, w * 0.05, h * 0.34, 45, color); // needle
  },
};


// Slide 1 - cover
function slide01(s) {
  ornateFrame(s);
  headline(s, 1.26, 4.716, 5.487, 1.717, "Waregh", 96);
  eyebrow(s, 1.26, 6.433, 5.344, 0.438, "Restaurant And Bar Presentation Template");
  dot(s, 7.44, 6.077);
  dot(s, 7.815, 6.077);
  dot(s, 8.19, 6.077);
  label(s, 9.718, 0.944, 2.356, 0.353, "Serve The Best Culinary", { align: 'right' });
  body(s, 7.377, 5.092, 4.696, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auc tor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus");
  label(s, 7.377, 4.678, 1.064, 0.353, "About Us");
  label(s, 8.748, 4.678, 1.366, 0.353, "Our Services");
  label(s, 10.422, 4.678, 1.651, 0.353, "Gallery Portfolio");
  photo(s, 1.26, 1.254, 6.754, 3.108, 'round2DiagRect', 0.629);
  photo(s, 8.171, 1.569, 1.872, 2.478, 'round2DiagRect', 0.319);
  photo(s, 10.201, 1.569, 1.872, 2.478, 'round2DiagRect', 0.315);
}

// Slide 2 - introduction
function slide02(s) {
  band(s, 0, 0.63, 13.333, 3.704);
  headline(s, 0.63, 5.086, 6.574, 1.784, "Your Culinary Adventure Awaits");
  eyebrow(s, 0.63, 4.649, 3.566, 0.438, "Introduction Our Restaurant");
  card(s, 7.363, 2.627, 1.551, 1.549, 0.161);
  card(s, 9.072, 2.63, 3.631, 1.549, 0.164, WHITE);
  body(s, 7.834, 4.709, 4.869, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam");
  text(s, 7.834, 6.212, 2.277, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: CLAY });
  text(s, 9.229, 2.958, 3.316, 0.909, "Nulla pellentesque dignissimt denim sit ametserii saeraseii venenatis urna cursus eget. Ac auctor augue mauris augue", { font: MED, size: 12, color: GREY, lineSpacing: 1.5 });
  icon.cutlery(s, 7.754, 3.017, 0.768, 0.769, CREAM);
  photo(s, 0.63, 0.787, 6.574, 3.389, 'roundRect', 0.155);
  photo(s, 7.361, 0.787, 1.666, 1.666, 'ellipse', 0);
  photo(s, 9.194, 0.787, 1.676, 1.676, 'ellipse', 0);
  photo(s, 11.027, 0.787, 1.676, 1.676, 'ellipse', 0);
}

// Slide 3 - about with stat cards
function slide03(s) {
  band(s, 5.968, 3.642, 7.365, 3.858);
  headline(s, 0.63, 1.697, 4.708, 2.625, "Where Every Flavor Tells A Story");
  eyebrow(s, 0.63, 1.26, 2.882, 0.438, "About Our Restaurant");
  body(s, 0.63, 4.706, 4.708, 1.532, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam vestibulum morbi blandit cursus risus. Tellus id interdum velit");
  card(s, 6.283, 5.059, 2.816, 1.811, 0.159);
  text(s, 6.426, 5.401, 2.53, 0.774, "2,400+", { font: SERIF, size: 40, color: WHITE });
  label(s, 6.426, 6.175, 2.592, 0.353, "About The Number Shown", { color: WHITE });
  card(s, 9.257, 5.059, 2.816, 1.811, 0.159);
  text(s, 9.4, 5.401, 2.53, 0.774, "63.18K", { font: SERIF, size: 40, color: WHITE });
  label(s, 9.4, 6.175, 2.592, 0.353, "About The Number Shown", { color: WHITE });
  photo(s, 6.283, 0.63, 6.42, 4.272, 'round2DiagRect', 0.629);
}

// Slide 4 - about with long copy
function slide04(s) {
  label(s, 5.968, 0.63, 3.244, 0.353, "Where Every Flavor Tells A Story");
  label(s, 5.968, 1.227, 3.244, 0.353, "Where Every Flavor Tells A Story");
  body(s, 5.968, 1.804, 6.519, 1.784, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam vestibulum morbi blandit cursus risus. Tellus id interdum velit. Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii");
  headline(s, 0.63, 1.067, 4.708, 2.625, "Where Every Flavor Tells A Story");
  eyebrow(s, 0.63, 0.63, 2.882, 0.438, "About Our Restaurant");
  photo(s, 0.63, 4.008, 6.01, 2.862, 'roundRect', 0.16);
  photo(s, 6.745, 4.008, 5.958, 2.862, 'roundRect', 0.161);
}

// Slide 5 - about with bullet notes
function slide05(s) {
  band(s, 0, 0, 5.968, 7.5);
  label(s, 1.024, 4.56, 1.584, 0.353, "The Restaurant");
  dot(s, 0.721, 4.658);
  body(s, 2.719, 4.253, 2.619, 1.027, "Nulla pellentesque dignissim enim sit ame venenatis urna cursus eget. Ac auctor aug ue mauris augue neque gravida. Enim nul la aliquet port isertor lacus luctus");
  headline(s, 0.63, 1.067, 4.708, 2.625, "Where Every Flavor Tells A Story");
  eyebrow(s, 0.63, 0.63, 2.882, 0.438, "About Our Restaurant");
  label(s, 1.024, 6.149, 1.584, 0.353, "The Restaurant");
  dot(s, 0.721, 6.246);
  body(s, 2.719, 5.841, 2.619, 1.027, "Nulla pellentesque dignissim enim sit ame venenatis urna cursus eget. Ac auctor aug ue mauris augue neque gravida. Enim nul la aliquet port isertor lacus luctus");
  photo(s, 6.283, 0.63, 4.176, 5.61, 'roundRect', 0.161);
  photo(s, 10.617, 1.26, 2.087, 2.726, 'roundRect', 0.16);
  photo(s, 10.617, 4.144, 2.087, 2.726, 'roundRect', 0.16);
}

// Slide 6 - about with circular gallery
function slide06(s) {
  band(s, 7.365, 0, 5.968, 7.5);
  body(s, 7.995, 5.841, 2.276, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neq ue gravida. Enim nulla aliquet port");
  headline(s, 7.995, 1.067, 4.708, 2.625, "Where Every Flavor Tells A Story");
  eyebrow(s, 7.995, 0.63, 2.882, 0.438, "About Our Restaurant");
  card(s, 7.995, 4.341, 2.275, 1.26, 0.108);
  text(s, 7.995, 4.618, 2.275, 0.707, "2,400+", { font: SERIF, size: 36, color: WHITE, align: 'center' });
  card(s, 10.428, 4.341, 2.275, 1.26, 0.108);
  text(s, 10.428, 4.618, 2.275, 0.707, "63.18K", { font: SERIF, size: 36, color: WHITE, align: 'center' });
  body(s, 10.428, 5.841, 2.276, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neq ue gravida. Enim nulla aliquet port");
  photo(s, 0.63, 0.63, 2.126, 2.126, 'ellipse', 0);
  photo(s, 3.271, 0.63, 3.465, 3.465, 'ellipse', 0);
  photo(s, 0.63, 3.406, 3.465, 3.465, 'ellipse', 0);
  photo(s, 4.609, 4.747, 2.126, 2.126, 'ellipse', 0);
}

// Slide 7 - staff banner
function slide07(s) {
  band(s, 0.63, 5.295, 12.073, 1.575);
  photo(s, 7.75, 2.914, 2.324, 2.324, 'ellipse', 0);
  photo(s, 10.223, 2.914, 2.324, 2.324, 'ellipse', 0);
  photo(s, -0.003, 0, 13.336, 2.606, 'rect', 0);
  headline(s, 0.63, 3.359, 6.632, 1.784, "High Quality Food Made By Chef");
  eyebrow(s, 0.63, 2.921, 2.74, 0.438, "Our Restaurant Staff");
  body(s, 0.945, 5.565, 6.17, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam vestibulum morbi blandit cursus");
  card(s, 7.745, 5.139, 2.322, 1.574, 0.153);
  text(s, 7.879, 5.926, 2.053, 0.438, "Greddy Amor", { font: SERIF, size: 20, color: WHITE });
  text(s, 8.226, 5.488, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  card(s, 10.225, 5.139, 2.322, 1.574, 0.153);
  text(s, 10.36, 5.926, 2.053, 0.438, "Greddy Amor", { font: SERIF, size: 20, color: WHITE });
  text(s, 10.706, 5.488, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
}

// Slide 8 - staff cards
function slide08(s) {
  band(s, 5.884, 0.63, 6.819, 6.24);
  photo(s, 6.199, 0.945, 3.016, 3.016, 'ellipse', 0);
  photo(s, 9.373, 0.945, 3.016, 3.016, 'ellipse', 0);
  card(s, 6.199, 3.646, 3.016, 2.909, 0.156);
  label(s, 0.991, 5.301, 2.228, 0.353, "About The Skill Shown");
  dot(s, 0.708, 5.399);
  headline(s, 0.63, 4.296, 2.526, 0.942, "78,51%");
  body(s, 0.63, 5.841, 4.624, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et");
  text(s, 6.451, 4.399, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 7.027, 3.961, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  headline(s, 0.63, 1.067, 4.624, 2.625, "High Quality Food Made By Chef");
  eyebrow(s, 0.63, 0.63, 2.74, 0.438, "Our Restaurant Staff");
  body(s, 6.451, 5.211, 2.512, 1.027, "Nulla pellentesque dignissim denim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port", { color: CREAM, align: 'center' });
  card(s, 9.373, 3.646, 3.016, 2.909, 0.156);
  text(s, 9.624, 4.399, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 10.2, 3.961, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  body(s, 9.624, 5.211, 2.512, 1.027, "Nulla pellentesque dignissim denim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port", { color: CREAM, align: 'center' });
}

// Slide 9 - staff portraits
function slide09(s) {
  band(s, 8.079, 3.85, 4.624, 3.02);
  photo(s, 0.63, 0.63, 3.331, 5.312, 'roundRect', 0.156);
  photo(s, 4.118, 0.63, 3.331, 5.312, 'roundRect', 0.156);
  label(s, 10.841, 4.299, 1.608, 0.353, "About Our Skill");
  dot(s, 10.458, 4.397);
  headline(s, 8.394, 4.13, 1.838, 0.69, "78.42%", 35);
  body(s, 8.394, 6.021, 3.994, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla");
  card(s, 0.787, 5.312, 3.016, 1.575, 0.156);
  text(s, 1.039, 6.049, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 1.615, 5.611, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  headline(s, 8.079, 1.067, 4.624, 2.625, "High Quality Food Made By Chef");
  eyebrow(s, 8.079, 0.63, 2.74, 0.438, "Our Restaurant Staff");
  label(s, 10.841, 5.214, 1.608, 0.353, "About Our Skill");
  dot(s, 10.458, 5.312);
  headline(s, 8.394, 5.046, 1.838, 0.69, "78.42%", 35);
  card(s, 4.276, 5.312, 3.016, 1.575, 0.156);
  text(s, 4.527, 6.049, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 5.103, 5.611, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
}

// Slide 10 - staff and skills
function slide10(s) {
  band(s, 0, 0, 13.333, 4.019);
  photo(s, 0.63, 0.63, 4.179, 3.074, 'roundRect', 0.158);
  photo(s, 4.967, 0.63, 4.179, 3.074, 'roundRect', 0.162);
  body(s, 7.886, 5.066, 4.817, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. ");
  label(s, 7.886, 4.649, 2.74, 0.353, "What Inside The Restaurant");
  text(s, 7.886, 6.517, 2.24, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: CLAY });
  card(s, 1.145, 2.916, 3.15, 1.417, 0.16);
  text(s, 1.463, 3.583, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 2.039, 3.145, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  headline(s, 0.63, 5.086, 6.632, 1.784, "High Quality Food Made By Chef");
  eyebrow(s, 0.63, 4.649, 2.74, 0.438, "Our Restaurant Staff");
  card(s, 5.482, 2.916, 3.15, 1.417, 0.16);
  text(s, 5.8, 3.583, 2.512, 0.522, "Greddy Amor", { font: SERIF, size: 25, color: WHITE });
  text(s, 6.376, 3.145, 1.361, 0.438, "Our Staff", { font: LIGHT, size: 20, color: WHITE, align: 'center' });
  headline(s, 9.461, 0.63, 2.59, 0.707, "78.62%", 36);
  label(s, 9.461, 1.34, 2.24, 0.353, "About The Skill Shown");
  body(s, 9.461, 1.755, 3.242, 0.269, "Nulla pellentesque dignissim enim sit amet venenatis");
  headline(s, 9.461, 2.307, 2.59, 0.707, "78.62%", 36);
  label(s, 9.461, 3.018, 2.24, 0.353, "About The Skill Shown");
  body(s, 9.461, 3.432, 3.242, 0.269, "Nulla pellentesque dignissim enim sit amet venenatis");
}

// Slide 11 - break time
function slide11(s) {
  ornateFrame(s);
  headline(s, 1.256, 1.257, 7.898, 1.717, "Break Time", 96);
  eyebrow(s, 1.26, 2.974, 5.118, 0.438, "Let’s Take A Break For About 15 Minutes");
  card(s, 1.256, 4.036, 3.682, 2.204, 0.315, NAVY, 'round2DiagRect');
  dot(s, 5.174, 4.573);
  dot(s, 5.174, 5.062);
  dot(s, 5.174, 5.551);
  rule(s, 1.256, 3.726, 10.818);
  label(s, 10.764, 1.257, 1.065, 0.353, "About Us", { align: 'right' });
  dot(s, 11.916, 1.355);
  label(s, 10.476, 1.939, 1.353, 0.353, "Our Services", { align: 'right' });
  dot(s, 11.916, 2.037);
  label(s, 10.174, 2.62, 1.655, 0.353, "Gallery Portfolio", { align: 'right' });
  dot(s, 11.916, 2.718);
  text(s, 1.57, 4.419, 2.301, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: WHITE });
  body(s, 1.57, 4.827, 3.052, 1.027, "Nulla pellentesque dignissim enim sit amet venen atis urna cursus eget. Ac auctor augue mauris au gue neque gravida. Enim nulla aliquet port iserto lacus luctus accumsan tortor posuere. Gravida", { color: WHITE });
  photo(s, 5.568, 4.035, 6.506, 2.203, 'round2DiagRect', 0.315);
}

// Slide 12 - services list
function slide12(s) {
  band(s, 7.94, 0, 5.394, 7.5);
  headline(s, 0.63, 4.529, 6.68, 1.784, "We Are Always Here To Serve You");
  eyebrow(s, 0.63, 4.092, 3.096, 0.438, "Our Restaurant Services");
  text(s, 0.63, 6.517, 2.315, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: CLAY });
  card(s, 8.255, 0.63, 1.417, 1.417, 0.156);
  label(s, 9.987, 0.63, 2.717, 0.353, "Our Restaurant Services 01");
  body(s, 9.987, 1.044, 2.717, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augu mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque");
  card(s, 8.255, 2.902, 1.417, 1.417, 0.156);
  label(s, 9.987, 2.902, 2.717, 0.353, "Our Restaurant Services 02");
  body(s, 9.987, 3.316, 2.717, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augu mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque");
  card(s, 8.255, 5.175, 1.417, 1.417, 0.156);
  label(s, 9.987, 5.175, 2.717, 0.353, "Our Restaurant Services 03");
  body(s, 9.987, 5.589, 2.717, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augu mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque");
  icon.wine(s, 8.789, 3.298, 0.344, 0.648, CREAM);
  icon.cake(s, 8.712, 5.567, 0.5, 0.632, CREAM);
  icon.pasta(s, 8.602, 1.004, 0.722, 0.671, CREAM);
  photo(s, 0.63, 0.63, 6.68, 3.147, 'roundRect', 0.156);
}

// Slide 13 - services with device shot
function slide13(s) {
  band(s, 7.305, 0, 2.205, 7.5);
  card(s, 6.675, 1.258, 1.417, 1.417, 0.16);
  body(s, 0.698, 5.211, 5.347, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien");
  label(s, 0.698, 4.797, 2.188, 0.353, "About The Restaurant");
  headline(s, 0.63, 1.696, 5.348, 2.625, "We Are Always Here To Serve You");
  eyebrow(s, 0.63, 1.258, 3.096, 0.438, "Our Restaurant Services");
  card(s, 6.675, 4.823, 1.417, 1.417, 0.16);
  card(s, 6.675, 3.041, 1.417, 1.417, 0.16);
  icon.wine(s, 7.21, 3.437, 0.344, 0.648, CREAM);
  icon.cake(s, 7.133, 5.215, 0.5, 0.632, CREAM);
  icon.pasta(s, 7.023, 1.633, 0.722, 0.671, CREAM);
  photo(s, 8.25, 0.628, 4.453, 6.242, 'roundRect', 0.159);
}

// Slide 14 - services tiles
function slide14(s) {
  band(s, 7.94, 0, 5.394, 7.5);
  label(s, 8.637, 4.649, 3.983, 0.353, "The Restaurant That Make Delicious Food");
  dot(s, 8.339, 4.747);
  body(s, 8.252, 5.168, 4.452, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus");
  label(s, 8.252, 6.517, 2.376, 0.353, "Serve The Best Culinary");
  headline(s, 0.63, 5.086, 6.68, 1.784, "We Are Always Here To Serve You");
  eyebrow(s, 0.63, 4.649, 3.096, 0.438, "Our Restaurant Services");
  card(s, 0.63, 0.63, 1.772, 1.773, 0.159);
  label(s, 0.728, 1.801, 1.575, 0.353, "Our Services 01", { color: WHITE, align: 'center', shrink: 0.925 });
  card(s, 0.63, 2.556, 1.772, 1.773, 0.159);
  label(s, 0.728, 3.727, 1.575, 0.353, "Our Services 02", { color: WHITE, align: 'center', shrink: 0.925 });
  icon.pasta(s, 1.155, 0.938, 0.722, 0.671, CREAM);
  icon.wine(s, 1.342, 2.885, 0.344, 0.648, CREAM);
  photo(s, 2.559, 0.63, 2.218, 3.699, 'roundRect', 0.158);
  photo(s, 4.934, 0.63, 2.218, 3.699, 'roundRect', 0.158);
  photo(s, 7.31, 0.63, 5.394, 3.699, 'roundRect', 0.158);
}

// Slide 15 - services grid
function slide15(s) {
  band(s, 7.94, 0, 1.772, 7.5);
  headline(s, 0.63, 1.067, 6.68, 1.784, "We Are Always Here To Serve You");
  eyebrow(s, 0.63, 0.63, 3.096, 0.438, "Our Restaurant Services");
  label(s, 10.348, 0.669, 2.355, 0.353, "Restaurant Services 01");
  body(s, 10.348, 1.083, 2.355, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor");
  card(s, 8.256, 0.63, 1.772, 1.773, 0.159);
  label(s, 8.356, 1.801, 1.571, 0.353, "Our Services 01", { color: WHITE, align: 'center', shrink: 0.925 });
  label(s, 10.348, 5.136, 2.355, 0.353, "Restaurant Services 03");
  body(s, 10.348, 5.55, 2.355, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor");
  card(s, 8.256, 5.097, 1.772, 1.773, 0.159);
  label(s, 8.356, 6.268, 1.571, 0.353, "Our Services 03", { color: WHITE, align: 'center', shrink: 0.925 });
  label(s, 10.348, 2.902, 2.355, 0.353, "Restaurant Services 02");
  body(s, 10.348, 3.316, 2.355, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor");
  card(s, 8.256, 2.863, 1.772, 1.773, 0.159);
  label(s, 8.356, 4.035, 1.571, 0.353, "Our Services 02", { color: WHITE, align: 'center', shrink: 0.925 });
  icon.pasta(s, 8.781, 0.938, 0.722, 0.671, CREAM);
  icon.wine(s, 8.968, 3.193, 0.344, 0.648, CREAM);
  icon.cake(s, 8.891, 5.423, 0.5, 0.632, CREAM);
  photo(s, 0.63, 3.166, 3.261, 3.704, 'roundRect', 0.162);
  photo(s, 4.05, 3.166, 3.261, 3.704, 'roundRect', 0.162);
}

// Slide 16 - gallery stat panel
function slide16(s) {
  headline(s, 0.633, 0.967, 7.549, 1.784, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 0.633, 0.529, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  s.addShape('rect', { x: 0.63, y: 3.598, w: 5.303, h: 3.272, fill: { color: CREAM } });
  headline(s, 0.945, 4.05, 2.123, 0.774, "2,600+", 40);
  label(s, 3.068, 4.396, 2.64, 0.353, "About The Number Shown");
  body(s, 0.945, 4.885, 4.674, 1.532, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam vestibulum morbi blandit cursus risus. Tellus id interdum velit");
  label(s, 1.004, 2.93, 3.611, 0.353, "The Restaurant That Make Delicious Food", { shrink: 0.85 });
  dot(s, 0.713, 3.028);
  photo(s, 6.091, 3.598, 2.564, 3.272, 'roundRect', 0.163);
  photo(s, 8.813, 0.63, 3.891, 6.24, 'roundRect', 0.155);
}

// Slide 17 - gallery with contact card
function slide17(s) {
  card(s, 8.771, 0.63, 3.932, 2.193, 0.157);
  label(s, 5.154, 5.934, 2.204, 0.353, "Inside The Restaurant");
  body(s, 5.154, 6.348, 3.617, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. ");
  text(s, 9.086, 0.95, 2.27, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: WHITE });
  label(s, 9.086, 2.15, 2.156, 0.353, "Inside The Restaurant", { color: WHITE });
  body(s, 9.086, 1.364, 3.302, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue", { color: WHITE });
  headline(s, 5.154, 3.891, 7.549, 1.784, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 5.154, 3.453, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  label(s, 9.086, 5.934, 2.204, 0.353, "Inside The Restaurant");
  body(s, 9.086, 6.348, 3.617, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. ");
  photo(s, 0.63, 0.63, 3.894, 6.24, 'roundRect', 0.159);
  photo(s, 4.681, 0.63, 1.887, 2.193, 'roundRect', 0.162);
  photo(s, 6.726, 0.63, 1.887, 2.193, 'roundRect', 0.162);
}

// Slide 18 - gallery bars
function slide18(s) {
  body(s, 7.469, 4.317, 5.34, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa sapien faucibus. Neque erase aliquam vestibulum morbi blandit cursus risus. Tellus id");
  card(s, 7.469, 6.161, 2.52, 0.709, 0.156);
  text(s, 7.706, 6.339, 2.046, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: WHITE, align: 'center', shrink: 0.85 });
  headline(s, 7.469, 1.067, 5.234, 2.625, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 7.469, 0.63, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  photoBars(s, 0.63, 0.63, 6.21, 6.24);
}

// Slide 19 - gallery circles
function slide19(s) {
  band(s, 6.84, 0, 6.494, 7.5);
  body(s, 7.469, 6.094, 5.34, 0.774, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant");
  card(s, 7.469, 4, 4.567, 0.709, 0.156);
  label(s, 7.781, 4.158, 3.943, 0.353, "The Restaurant That Make Delicious Food", { color: WHITE, align: 'center' });
  headline(s, 7.469, 1.067, 5.234, 2.625, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 7.469, 0.63, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  card(s, 7.469, 5.017, 4.567, 0.709, 0.156);
  label(s, 7.781, 5.174, 3.943, 0.353, "The Restaurant That Make Delicious Food", { color: WHITE, align: 'center' });
  photo(s, 0.63, 1.23, 5.04, 5.04, 'ellipse', 0);
  photo(s, 4.436, 0.629, 2.089, 2.089, 'ellipse', 0, PHOTO_DK);
  photo(s, 4.436, 4.781, 2.089, 2.089, 'ellipse', 0, PHOTO_DK);
}

// Slide 20 - gallery trio
function slide20(s) {
  band(s, 10.061, 0, 3.272, 7.5);
  label(s, 0.632, 6.396, 2.204, 0.353, "Inside The Restaurant");
  body(s, 3.151, 6.341, 5.031, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus");
  headline(s, 0.633, 4.339, 7.549, 1.784, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 0.633, 3.901, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  photo(s, 0.63, 0.63, 3.934, 2.96, 'roundRect', 0.167);
  photo(s, 4.721, 0.63, 3.934, 2.96, 'roundRect', 0.167);
  photo(s, 8.813, 0.63, 3.891, 6.235, 'roundRect', 0.159);
}

// Slide 21 - gallery mosaic
function slide21(s) {
  body(s, 0.63, 5.841, 5.322, 1.027, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. Gravida cum sociis natoque seraser penatibus et magnis. Habitant morbi tristique senectus et netus et seraseii malesuada. Consectetur lorem donec massa");
  headline(s, 0.63, 1.067, 5.234, 2.625, "Hundreds Of Flavors Under One Roof");
  eyebrow(s, 0.63, 0.63, 4.06, 0.438, "Restaurant Gallery and Portfolio");
  label(s, 2.362, 4.141, 2.204, 0.353, "Inside The Restaurant");
  body(s, 2.362, 4.555, 3.502, 0.774, "Nulla pellentesque dignissim enim sit amet venenati sser serakieora urnasersersera cursus egetser Ac auctorserase augue mauris augueseras neque gravidaers ser. Enim");
  card(s, 0.63, 4.028, 1.417, 1.417, 0.156);
  icon.pasta(s, 0.978, 4.402, 0.722, 0.671, CREAM);
  photo(s, 6.494, 0.63, 1.816, 2.346, 'roundRect', 0.16);
  photo(s, 8.467, 0.63, 4.236, 2.346, 'roundRect', 0.158);
  photo(s, 10.492, 3.133, 2.211, 3.737, 'roundRect', 0.157);
  photo(s, 6.494, 3.133, 3.841, 3.737, 'roundRect', 0.161);
}

// Slide 22 - tablet mockup
function slide22(s) {
  band(s, 0, 3.166, 13.333, 4.334);
  headline(s, 0.63, 1.067, 6.283, 1.784, "Food That Makes You Say Wow");
  eyebrow(s, 0.63, 0.63, 3.596, 0.438, "Restaurant Mockup Devices ");
  tabletMockup(s, 7.543, 0.63, 5.166, 7.625);
  text(s, 3.028, 3.859, 3.885, 1.212, "Nulla pellentesque dignissim enim sit amet vene natis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor posuere. ", { font: MED, size: 12, color: GREY, lineSpacing: 1.5 });
  label(s, 4.13, 5.36, 2.155, 0.353, "About The Restaurant", { shrink: 0.925 });
  body(s, 4.13, 5.775, 2.782, 0.774, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla");
  icon.cutlery(s, 3.072, 5.606, 0.698, 0.699, NAVY);
  photo(s, 0.63, 3.481, 2.083, 3.383, 'roundRect', 0.156);
  photo(s, 7.854, 1.441, 4.486, 5.962, 'rect', 0);
}

// Slide 23 - laptop mockup
function slide23(s) {
  band(s, 0, 0, 13.333, 3.166);
  laptopMockup(s, 0.63, 0.632, 5.165, 2.989);
  label(s, 9.439, 6.305, 1.568, 0.353, "The Restaurant");
  dot(s, 9.167, 6.403);
  label(s, 9.439, 5.689, 1.568, 0.353, "The Restaurant");
  dot(s, 9.167, 5.787);
  body(s, 9.085, 4.035, 3.618, 1.279, "Nulla pellentesque dignissim enim sit amet venenatis urnaser cursus egetser Ac auctor augue mauris augue neque gravidaer. Enim nulla aliquet portser isertor lacus luctus accumsan tortorer posuere. Gravida cum sociis natoquet serasesr penatibus et mauris augue neque saera");
  label(s, 9.085, 3.621, 2.346, 0.353, "About The Restaurant");
  headline(s, 6.425, 1.067, 6.283, 1.784, "Food That Makes You Say Wow");
  eyebrow(s, 6.425, 0.63, 3.596, 0.438, "Restaurant Mockup Devices ");
  photo(s, 1.24, 0.792, 3.948, 2.469, 'rect', 0);
  photo(s, 0.63, 4.251, 2.615, 2.615, 'ellipse', 0);
  photo(s, 3.396, 4.251, 2.615, 2.615, 'ellipse', 0);
  photo(s, 6.161, 4.251, 2.615, 2.615, 'ellipse', 0);
}

// Slide 24 - contact
function slide24(s) {
  band(s, 0, 3.481, 13.333, 3.074);
  headline(s, 0.63, 1.067, 7.773, 1.784, "You’re In Good Hands With Us");
  eyebrow(s, 0.63, 0.63, 4.353, 0.438, "Get In Touch With Our Restaurant");
  body(s, 3.262, 3.857, 5.141, 0.522, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus");
  label(s, 3.262, 5.554, 1.538, 0.353, "Our Cellphone");
  s.addShape('ellipse', { x: 3.327, y: 4.624, w: 0.866, h: 0.866, fill: { color: CLAY } });
  body(s, 3.262, 5.969, 2.413, 0.269, "Nulla pellent seresque dignissim enim");
  label(s, 5.99, 5.554, 1.538, 0.353, "Our Location");
  s.addShape('ellipse', { x: 6.055, y: 4.624, w: 0.866, h: 0.866, fill: { color: CLAY } });
  body(s, 5.99, 5.969, 2.413, 0.269, "Nulla pellent seresque dignissim enim");
  icon.phone(s, 3.541, 4.89, 0.439, 0.327, CREAM);
  icon.pin(s, 6.281, 4.85, 0.408, 0.408, CREAM);
  photo(s, 0.63, 3.166, 2.317, 3.704, 'round2DiagRect', 0.629);
  photo(s, 9.033, 0, 4.3, 7.5, 'rect', 0);
}

// Slide 25 - thank you
function slide25(s) {
  ornateFrame(s);
  eyebrow(s, 5.263, 5.808, 5.344, 0.438, "Restaurant And Bar Presentation Template");
  headline(s, 5.263, 1.263, 7.44, 1.717, "Thank You", 96);
  eyebrow(s, 5.263, 2.981, 3.371, 0.438, "Thanks For Your Attention");
  card(s, 5.263, 3.904, 4.285, 1.417, 0.158, WHITE);
  label(s, 9.863, 4.261, 1.194, 0.353, "Visit Us At:");
  text(s, 9.863, 4.612, 2.298, 0.353, "www.wareghresto.com", { font: MED, size: 15, color: CLAY });
  body(s, 5.581, 4.255, 3.652, 0.774, "Nulla pellentesque dignissim enim sit amet venenatis urna cursus eget. Ac auctor augue mauris augue neque gravida. Enim nulla aliquet port isertor lacus luctus accumsan tortor");
  photo(s, 1.258, 1.26, 3.375, 5.61, 'round2DiagRect', 0.628);
}

// ------------------------------------------------------------------ build
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WAREGH', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'WAREGH';
pptx.author = 'Waregh';
pptx.title = 'Waregh - Restaurant And Bar Presentation Template';

BUILDERS.forEach(function (build) {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  build(slide);
});

pptx.writeFile({
  fileName: path.join(__dirname, '04eec701-b9bf-43f1-a89d-364cf0ccfae6_grok_final.pptx'),
}).then(function (f) {
  console.log('wrote ' + f);
});
