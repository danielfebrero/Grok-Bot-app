/**
 * Recreation of "Project Action Plan" (16 slides, 10 x 5.625 in) with PptxGenJS.
 * Raster/vector icon art from the original deck is replaced by simple
 * programmatic placeholders (see `icon()` / `imgBox()`).
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  a1: '0F6FC6', a2: '009DD9', a3: '0BD0D9', a4: '10CF9B', a5: '7CCA62', a6: 'A5C249',
  white: 'FFFFFF',
  text: '404040',   // tx1 lumMod 75% / lumOff 25%
  head: '1A1A1A',   // tx1 lumMod 90% / lumOff 10%
  g95: 'F2F2F2', g85: 'D9D9D9', g75: 'BFBFBF', g65: 'A6A6A6', g50: '808080',
};
const ACCENTS = [C.a1, C.a2, C.a3, C.a4, C.a5, C.a6];
// accent1..6 lightened (lumMod 40% / lumOff 60%) - used for the halo rings
const ACCENT_LIGHT = ['91C6F7', '8ADFFF', '94F5FA', '94F7DC', 'CBEAC0', 'DBE7B6'];
// accent1..6 pale tint (lumMod 20% / lumOff 80%) - used for table banding
const ACCENT_PALE = ['C8E3FB', 'C4EFFF', 'C9FAFC', 'CAFBED', 'E5F4E0', 'EDF3DB'];
// accent1..6 darkened (lumMod 75%)
const ACCENT_DARK = ['0B5394', '0076A3', '089CA3', '0C9B74', '55A839', '7E9632'];

/** Linear blend between two hex colours; t = 0 keeps `a`, t = 1 keeps `b`. */
function blend(a, b, t) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const v = Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
    out += ('0' + v.toString(16).toUpperCase()).slice(-2);
  }
  return out;
}
// pptxgenjs has no gradient fill, so shapes the deck fills with an
// accent -> accent-shade gradient use the average of the two stops instead.
const ACCENT_MID = ACCENTS.map(function (c, i) { return blend(c, ACCENT_DARK[i], 0.45); });

const FONT = 'Roboto';
const LOREM = 'Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in';
const LOREM_LONG = 'Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from';
const STD = 'The Standard Lorem Ipsum ';

/* ---------------------------------------------------------------- helpers */

/** Add an auto-shape. */
function sh(slide, type, o) { slide.addShape(type, o); }

/** Add a text run / paragraph block with the deck's zero text insets. */
function tx(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, margin: 0, valign: 'middle' }, o));
}

/** Custom geometry from normalised [0..1] path segments: ['M'|'L',x,y] / ['C',x1,y1,x2,y2,x,y] / ['Z']. */
function poly(slide, segs, o) {
  const pts = [];
  for (const s of segs) {
    if (s[0] === 'Z') { pts.push({ close: true }); }
    else if (s[0] === 'C') {
      pts.push({
        x: o.w * s[5], y: o.h * s[6],
        curve: { type: 'cubic', x1: o.w * s[1], y1: o.h * s[2], x2: o.w * s[3], y2: o.h * s[4] },
      });
    } else {
      pts.push({ x: o.w * s[1], y: o.h * s[2], moveTo: s[0] === 'M' });
    }
  }
  slide.addShape('custGeom', Object.assign({}, o, { points: pts }));
}

/** Normalised path for a w x h rectangle with per-corner radii (inches) [tl, tr, br, bl]. */
function rectPath(w, h, radii) {
  const K = 0.5523; // circle -> cubic-bezier constant
  const r = radii.map(function (v) { return Math.min(v, w / 2, h / 2); });
  const nx = function (v) { return v / w; }, ny = function (v) { return v / h; };
  const segs = [['M', nx(r[0]), 0], ['L', nx(w - r[1]), 0]];
  if (r[1]) segs.push(['C', nx(w - r[1] * (1 - K)), 0, 1, ny(r[1] * (1 - K)), 1, ny(r[1])]);
  segs.push(['L', 1, ny(h - r[2])]);
  if (r[2]) segs.push(['C', 1, ny(h - r[2] * (1 - K)), nx(w - r[2] * (1 - K)), 1, nx(w - r[2]), 1]);
  segs.push(['L', nx(r[3]), 1]);
  if (r[3]) segs.push(['C', nx(r[3] * (1 - K)), 1, 0, ny(h - r[3] * (1 - K)), 0, ny(h - r[3])]);
  segs.push(['L', 0, ny(r[0])]);
  if (r[0]) segs.push(['C', 0, ny(r[0] * (1 - K)), nx(r[0] * (1 - K)), 0, nx(r[0]), 0]);
  segs.push(['Z']);
  return segs;
}

/** Chevron outline with an explicit notch depth (the preset adj is not exposed by pptxgenjs). */
function chevronPath(notch) {
  return [['M', 0, 0], ['L', 1 - notch, 0], ['L', 1, 0.5], ['L', 1 - notch, 1],
    ['L', 0, 1], ['L', notch, 0.5], ['Z']];
}

/** Right-angled "ribbon fold" wedge; `corner` names the square (90 deg) corner. */
const FOLD_PATH = {
  tl: [['M', 0, 0], ['L', 1, 0], ['L', 0, 1], ['Z']],
  tr: [['M', 1, 0], ['L', 1, 1], ['L', 0, 0], ['Z']],
  bl: [['M', 0, 1], ['L', 0, 0], ['L', 1, 1], ['Z']],
  br: [['M', 1, 1], ['L', 1, 0], ['L', 0, 1], ['Z']],
};
function fold(slide, corner, x, y, w, h, transparency) {
  poly(slide, FOLD_PATH[corner], {
    x: x, y: y, w: w, h: h,
    fill: { color: '000000', transparency: transparency }, line: { type: 'none' },
  });
}

/**
 * Org-chart splitter: a downward-opening brace whose centre tick rises to the
 * parent at `tickX` and whose two ends drop to `x`..`x + w`. Drawn as an open
 * path so the rounded corners survive; `h` is the vertical reach.
 */
function splitter(slide, x, y, w, h, tickX, color) {
  const rx = Math.min(0.09, w / 6) / w, ry = 0.42;   // corner rounding, normalised
  const t = (tickX - x) / w;
  poly(slide, [
    ['M', 0, 1], ['C', 0, 1 - ry * 0.55, rx * 0.45, ry, rx, ry],
    ['L', t - rx, ry], ['C', t - rx * 0.45, ry, t, ry * 0.55, t, 0],
    ['C', t, ry * 0.55, t + rx * 0.45, ry, t + rx, ry],
    ['L', 1 - rx, ry], ['C', 1 - rx * 0.45, ry, 1, 1 - ry * 0.55, 1, 1],
  ], { x: x, y: y, w: w, h: h, fill: { type: 'none' }, line: { color: color, width: 0.75 } });
}

/** A ring drawn as a stroked ellipse (outer edge = x,y,w,h). */
function ring(slide, x, y, d, ptWidth, color) {
  const inset = ptWidth / 72 / 2;
  sh(slide, 'ellipse', {
    x: x + inset, y: y + inset, w: d - 2 * inset, h: d - 2 * inset,
    fill: { type: 'none' }, line: { color: color, width: ptWidth },
  });
}

/**
 * Placeholder standing in for one of the deck's line-art icon glyphs:
 * an outlined frame holding three ascending bars, drawn in the caller's colour.
 */
function icon(slide, x, y, size, color) {
  const stroke = Math.max(0.75, size * 2.4);
  sh(slide, 'roundRect', {
    x: x + size * 0.06, y: y + size * 0.12, w: size * 0.88, h: size * 0.76,
    rectRadius: size * 0.1, fill: { type: 'none' }, line: { color: color, width: stroke },
  });
  [[0.22, 0.30], [0.44, 0.44], [0.66, 0.58]].forEach(function (b) {
    sh(slide, 'rect', {
      x: x + b[0] * size, y: y + (0.76 - b[1] * 0.6) * size,
      w: size * 0.11, h: size * b[1] * 0.6,
      fill: { color: color }, line: { type: 'none' },
    });
  });
}

/** Placeholder standing in for a raster picture of the original deck. */
function imgBox(slide, x, y, w, h, label) {
  sh(slide, 'rect', { x, y, w, h, fill: { color: C.g85 }, line: { color: C.g75, width: 0.75 } });
  if (label !== false) {
    tx(slide, '[image]', { x, y, w, h, align: 'center', valign: 'middle', fontSize: 8, color: C.g50 });
  }
}

/** Standard title + sub-headline placeholders (align 'center' or 'left'). */
function heading(slide, align) {
  tx(slide, 'Project Action Plan', {
    x: 0.424, y: 0.171, w: 9.152, h: 0.542, align: align, fontSize: 26, color: C.head,
  });
  tx(slide, 'Enter your sub headline here', {
    x: 0.424, y: 0.727, w: 9.152, h: 0.189, align: align, fontSize: 10, color: C.head,
  });
}

const BULLET = { characterCode: '2022', indent: 13.5 };

/**
 * Soft outer shadow; `angle` is degrees clockwise from +x, matching OOXML `dir`.
 * Always returns a fresh object: pptxgenjs rescales the options in place, so a
 * shared literal would be re-converted (and blow up) on every reuse.
 */
function shadow(blur, offset, angle, opacity) {
  return { type: 'outer', color: '000000', blur: blur, offset: offset, angle: angle, opacity: opacity };
}
const cardShadow = function () { return shadow(11, 4, 90, 0.16); };   // lifted card / circle
const deepShadow = function () { return shadow(14, 13, 45, 0.15); };  // long cast behind a track

/* -------------------------------------------------------- shared geometry */

// Slide 4 / 10 / 15 / 16 outlines lifted from the original custom geometries.
const P_TURN_ARROW = [['M', 0.5911, 0], ['L', 0, 0], ['L', 0, 0.1278], ['L', 0.5911, 0.1278],
  ['L', 0.5911, 1], ['L', 1, 1], ['L', 1, 0.1278],
  ['C', 1, 0.0573, 0.8179, 0, 0.5911, 0], ['Z']];
const P_CHEV_OUTLINE = [['M', 1, 0], ['L', 0.6318, 0], ['L', 0, 0.4993], ['L', 0.6318, 1],
  ['L', 1, 1], ['L', 0.5603, 0.4993], ['L', 1, 0], ['Z']];
const P_TAB = [['M', 1, 0.0515], ['L', 1, 0.8327], ['L', 0.5, 1], ['L', 0, 0.8327], ['L', 0, 0.0515],
  ['C', 0, 0.0231, 0.0689, 0, 0.1539, 0], ['L', 0.8461, 0],
  ['C', 0.9311, 0, 1, 0.0231, 1, 0.0515], ['Z']];
const P_BADGE = [['M', 0.772, 1.0147], ['L', 0.037, 1.0147],
  ['C', 0.0166, 1.0147, 0, 0.9657, 0, 0.9051], ['L', 0, 0.1096],
  ['C', 0, 0.049, 0.0166, 0, 0.037, 0], ['L', 0.772, 0],
  ['C', 0.8146, 0, 0.8554, 0.0488, 0.8861, 0.1361], ['L', 0.9887, 0.4286],
  ['C', 1.0038, 0.4716, 1.0038, 0.5433, 0.9887, 0.5865], ['L', 0.8861, 0.879],
  ['C', 0.8554, 0.9659, 0.8146, 1.0147, 0.772, 1.0147], ['Z']];
const P_PILL_L = [['M', 0.1436, 1], ['C', 0.064, 1, 0, 0.7757, 0, 0.5008],
  ['C', 0, 0.2243, 0.064, 0, 0.1436, 0], ['L', 0.8564, 0],
  ['C', 0.9355, 0, 1, 0.2243, 1, 0.5008], ['C', 1, 0.7757, 0.9355, 1, 0.8564, 1], ['Z']];
const P_BIG_ARROW = [['M', 0.2414, 0], ['L', 0.2414, 0.2614], ['L', 0.7759, 0.2614],
  ['L', 0.7759, 0.0152], ['L', 1, 0.5076], ['L', 0.7759, 1], ['L', 0.7759, 0.7538],
  ['L', 0, 0.7538], ['L', 0, 0.2614], ['Z']];

/* =============================================================== slide 01 */
// Five circles on gray half-rings, with alternating callout cards.
function slide01(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  const items = [
    { x: 0.919, up: true,  color: C.a1 },
    { x: 2.635, up: false, color: C.a2 },
    { x: 4.350, up: true,  color: C.a3 },
    { x: 6.066, up: false, color: C.a4 },
    { x: 7.781, up: true,  color: C.a5 },
  ];
  items.forEach(function (it) {
    // gray half ring behind the circle (bottom for "up" cards, top otherwise)
    sh(s, 'blockArc', {
      x: it.x - 0.307, y: 2.28, w: 1.914, h: 1.914,
      angleRange: it.up ? [10, 170] : [190, 350], arcThicknessRatio: 0.26,
      fill: { color: C.g65 }, line: { type: 'none' },
    });
    sh(s, 'ellipse', {
      x: it.x, y: 2.587, w: 1.3, h: 1.3,
      fill: { color: it.color }, line: { color: C.white, width: 6 },
    });
    icon(s, it.x + 0.336, 2.923, 0.627, C.white);

    // callout card + connector arrow
    const cardY = it.up ? 1.274 : 4.113;
    sh(s, 'roundRect', {
      x: it.x - 0.14, y: cardY, w: 1.58, h: 1.084, rectRadius: 0.083,
      fill: { color: 'F7F7F7' }, line: { color: C.g85, width: 0.75 },
      shadow: shadow(6, 2, 90, 0.12),
    });
    tx(s, STD, {
      x: it.x + 0.066, y: cardY + 0.215, w: 1.167, h: 0.501,
      align: 'center', fontSize: 10.5, bold: true, color: C.head, lineSpacingMultiple: 1.5,
    });
    sh(s, it.up ? 'downArrow' : 'upArrow', {
      x: it.x + 0.493, y: it.up ? 2.358 : 3.828, w: 0.314, h: 0.302,
      fill: { color: C.g85 }, line: { type: 'none' },
    });
  });
}

/* =============================================================== slide 02 */
// Ascending "level" bars with a dashed diagonal arrow.
function slide02(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  const bars = [
    { y: 1.146, w: 4.096, color: C.a1, level: '05', text: 'Sample text goes here, replace with' },
    { y: 1.775, w: 4.944, color: C.a2, level: '04', text: 'This is a sample text. You simply add your own' },
    { y: 2.404, w: 5.583, color: C.a3, level: '03', text: 'This is a sample text. You simply add your own' },
    { y: 3.033, w: 6.389, color: C.a4, level: '02', text: 'This is a sample text. You simply add your own' },
    { y: 3.663, w: 7.222, color: C.a5, level: '01', text: 'This is a sample text. You simply add your own' },
  ];
  bars.forEach(function (b) {
    sh(s, 'roundRect', {
      x: 0.424, y: b.y, w: b.w, h: 0.583, rectRadius: 0.097,
      fill: { color: b.color }, line: { type: 'none' },
    });
    tx(s, b.text, { x: 0.796, y: b.y + 0.199, w: 4.2, h: 0.185, fontSize: 11, color: C.white });
    // translucent badge riding on the right end of the bar
    sh(s, 'roundRect', {
      x: 0.424 + b.w - 0.727, y: b.y - 0.032, w: 0.647, h: 0.647, rectRadius: 0.108,
      fill: { color: C.white, transparency: 72 }, line: { color: C.white, width: 0.75 },
    });
    tx(s, [{ text: 'Level', options: { breakLine: true } }, { text: b.level }], {
      x: 0.424 + b.w - 0.727, y: b.y - 0.032, w: 0.647, h: 0.647,
      align: 'center', valign: 'middle', fontSize: 9, bold: true, color: C.white,
    });
  });

  // summary bar
  sh(s, 'roundRect', {
    x: 0.424, y: 4.292, w: 7.968, h: 0.931, rectRadius: 0.072,
    fill: { color: C.a6 }, line: { type: 'none' },
  });
  sh(s, 'roundRect', {
    x: 0.71, y: 4.465, w: 0.584, h: 0.584, rectRadius: 0.076,
    fill: { color: C.white, transparency: 72 }, line: { color: C.white, width: 0.75 },
  });
  icon(s, 0.788, 4.543, 0.43, C.white);
  tx(s, 'Placeholder text', { x: 1.559, y: 4.664, w: 1.441, h: 0.185, fontSize: 11, color: C.white });
  tx(s, 'Contrary to popular belief, Lorem Ipsum is not simply random text. '
       + 'It has roots in a piece of classical Latin literature from', {
    x: 3.181, y: 4.542, w: 3.721, h: 0.429, fontSize: 9, color: C.white, lineSpacingMultiple: 1.5,
  });

  // dashed diagonal arrow (head at the top-left) + rotated caption
  sh(s, 'line', {
    x: 5.017, y: 0.971, w: 4.15, h: 3.452,
    line: { color: C.g50, width: 1, dashType: 'sysDash', beginArrowType: 'triangle' },
  });
  tx(s, 'You can use this sentence', {
    x: 6.033, y: 2.605, w: 2.116, h: 0.185, rotate: 40, align: 'center',
    fontSize: 11, bold: true, color: C.text,
  });
}

/* =============================================================== slide 03 */
// Four-column data table.
function slide03(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  const head = [C.a1, C.a2, C.a3, C.a4].map(function (c) {
    return {
      text: 'Add. Description',
      options: { fill: { color: c }, color: C.white, bold: true, fontSize: 12, align: 'center' },
    };
  });
  const cellBorder = [
    { type: 'none' }, { type: 'none' },
    { type: 'solid', color: C.g85, pt: 1 }, { type: 'none' },
  ];
  const body = ['The Standard Lorem Ipsum ', 'Description Text', '2027',
    'Contrary to popular belief, Lorem Ipsum is not simply random text. '];
  const rows = [head];
  for (let r = 0; r < 6; r++) {
    rows.push(body.map(function (t) {
      return { text: t, options: { fontSize: 9, color: C.head, align: 'center', border: cellBorder } };
    }));
  }
  s.addTable(rows, {
    x: 0.424, y: 1.268, w: 9.152, colW: [2.288, 2.288, 2.288, 2.288],
    rowH: [0.711, 0.524, 0.524, 0.524, 0.524, 0.524, 0.524],
    valign: 'middle', fontFace: FONT, border: { type: 'none' }, margin: [2, 4, 2, 4],
  });
}

/* =============================================================== slide 04 */
// Six turning arrows fanning out of the centre.
function slide04(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  const arms = [
    { color: C.a3, ax: 3.807, ay: 1.916, ah: 3.709, hx: 3.419, hy: 1.770, side: 'L', label: 'TITLE C', lx: 1.262, ly: 1.854, ix: 2.639, iy: 1.935 },
    { color: C.a2, ax: 3.266, ay: 2.843, ah: 2.782, hx: 2.890, hy: 2.697, side: 'L', label: 'TITLE B', lx: 0.776, ly: 2.781, ix: 2.185, iy: 2.862 },
    { color: C.a1, ax: 2.725, ay: 3.770, ah: 1.855, hx: 2.380, hy: 3.625, side: 'L', label: 'TITLE A', lx: 0.300, ly: 3.709, ix: 1.688, iy: 3.789 },
    { color: C.a4, ax: 5.034, ay: 1.916, ah: 3.709, hx: 5.890, hy: 1.770, side: 'R', label: 'TITLE D', lx: 7.501, ly: 1.854, ix: 6.910, iy: 1.935 },
    { color: C.a5, ax: 5.576, ay: 2.843, ah: 2.782, hx: 6.420, hy: 2.697, side: 'R', label: 'TITLE E', lx: 7.988, ly: 2.781, ix: 7.472, iy: 2.862 },
    { color: C.a6, ax: 6.117, ay: 3.770, ah: 1.855, hx: 6.931, hy: 3.625, side: 'R', label: 'TITLE F', lx: 8.464, ly: 3.709, ix: 7.992, iy: 3.781 },
  ];
  arms.forEach(function (a) {
    const left = a.side === 'L';
    poly(s, P_TURN_ARROW, {
      x: a.ax, y: a.ay, w: 1.158, h: a.ah, flipH: !left,
      fill: { color: a.color }, line: { type: 'none' },
    });
    poly(s, P_CHEV_OUTLINE, {
      x: a.hx, y: a.hy, w: 0.69, h: 0.767, flipH: !left,
      fill: { color: a.color }, line: { type: 'none' },
    });
    tx(s, a.label, {
      x: a.lx, y: a.ly, w: 1.236, h: 0.202, align: left ? 'right' : 'left',
      fontSize: 12, bold: true, color: C.text,
    });
    tx(s, 'Lorem ipsum dolor sit amet, consectetuer', {
      x: a.lx, y: a.ly + 0.195, w: 1.236, h: 0.404, align: left ? 'right' : 'left',
      fontSize: 8, color: C.text, lineSpacingMultiple: 1.5, valign: 'top',
    });
    icon(s, a.ix, a.iy, 0.437, C.head);
  });
  // soft folds where each turning arm meets its vertical shaft
  [[3.418, 4.240, 'tl'], [3.954, 3.316, 'tl'], [4.495, 2.388, 'tl'],
   [6.117, 4.240, 'tr'], [5.570, 3.316, 'tr'], [5.031, 2.388, 'tr']].forEach(function (p) {
    fold(s, p[2], p[0], p[1], 0.476, 0.149, 75);
  });
}

/* =============================================================== slide 05 */
// Org-tree on a soft panel feeding three description cards.
function slide05(pptx) {
  const s = pptx.addSlide();
  sh(s, 'roundRect', {
    x: 0.424, y: 1.146, w: 5.111, h: 4.083, rectRadius: 0.091,
    fill: { color: C.g95 }, line: { type: 'none' },
  });
  heading(s, 'left');

  // root node
  sh(s, 'roundRect', { x: 2.174, y: 1.491, w: 1.677, h: 0.583, rectRadius: 0.044, fill: { color: C.a1 }, line: { type: 'none' } });
  icon(s, 2.835, 1.606, 0.42, C.white);
  sh(s, 'line', { x: 3.013, y: 2.125, w: 0, h: 0.769, line: { color: C.text, width: 1 } });

  // splitter braces (rotated 90 deg, exactly as in the source deck)
  const braceLine = { color: C.text, width: 1 };
  sh(s, 'leftBrace', { x: 2.565, y: 0.980, w: 0.896, h: 3.083, rotate: 90, fill: { type: 'none' }, line: braceLine });
  sh(s, 'leftBrace', { x: 2.593, y: 2.180, w: 0.840, h: 3.417, rotate: 90, fill: { type: 'none' }, line: braceLine });
  sh(s, 'leftBrace', { x: 2.593, y: 3.326, w: 0.840, h: 1.126, rotate: 90, fill: { type: 'none' }, line: braceLine });

  [0.721, 2.304, 3.888].forEach(function (x) {
    sh(s, 'roundRect', { x: x, y: 2.896, w: 1.417, h: 0.583, rectRadius: 0.044, fill: { color: C.a2 }, line: { type: 'none' } });
    icon(s, x + 0.52, 3.0, 0.376, C.white);
  });
  [0.721, 1.888, 3.054, 4.221].forEach(function (x) {
    sh(s, 'roundRect', { x: x, y: 4.301, w: 1.083, h: 0.583, rectRadius: 0.044, fill: { color: C.a3 }, line: { type: 'none' } });
    icon(s, x + 0.36, 4.412, 0.362, C.white);
  });

  // right hand cards + funnel arrows
  [C.a1, C.a2, C.a3].forEach(function (color, i) {
    const y = 1.146 + i * 1.4048;
    sh(s, 'roundRect', { x: 6.417, y: y, w: 3.159, h: 1.273, rectRadius: 0.096, fill: { color: color }, line: { type: 'none' } });
    sh(s, 'rightArrow', {
      x: 5.800, y: 1.5705 + i * 1.4048, w: 0.388, h: 0.424,
      fill: { color: color }, line: { type: 'none' },
    });
    tx(s, LOREM, {
      x: 6.887, y: y + 0.199, w: 2.317, h: 0.875, fontSize: 12, color: C.white,
      bullet: BULLET, lineSpacingMultiple: 1.5,
    });
  });
}

/* =============================================================== slide 06 */
// Four horizontal "step" rows.
function slide06(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  ACCENTS.slice(0, 4).forEach(function (color, i) {
    const dy = i * 1.0315;
    // left colour block (rounded on its left edge) then the pointed tab
    poly(s, rectPath(1.909, 1.008, [0.155, 0, 0, 0.155]), {
      x: 0.4245, y: 1.1375 + dy, w: 1.909, h: 1.008,
      fill: { color: color }, line: { type: 'none' },
    });
    poly(s, P_TAB, {
      x: 2.455, y: 0.136 + dy, w: 1.008, h: 3.012, rotate: 270,
      fill: { color: color }, line: { type: 'none' },
    });
    // translucent icon plate
    sh(s, 'roundRect', {
      x: 0.547, y: 1.248 + dy, w: 0.786, h: 0.786, rectRadius: 0.102,
      fill: { color: C.white, transparency: 70 }, line: { color: C.white, width: 0.75 },
    });
    icon(s, 0.722, 1.425 + dy, 0.433, C.white);
    tx(s, STD, { x: 1.712, y: 1.44 + dy, w: 1.595, h: 0.404, fontSize: 12, bold: true, color: C.white });
    // chevron track
    poly(s, chevronPath(0.0786), {
      x: 4.25, y: 1.224 + dy, w: 5.326, h: 0.837,
      fill: { color: C.g95 }, line: { color: C.white, width: 1 }, shadow: deepShadow(),
    });
    // step badge
    poly(s, P_BADGE, {
      x: 3.496, y: 1.330 + dy, w: 1.722, h: 0.624,
      fill: { color: ACCENT_MID[i] }, line: { type: 'none' },
    });
    tx(s, 'STEP 0' + (i + 1), {
      x: 3.816, y: 1.541 + dy, w: 1.083, h: 0.202,
      align: 'center', fontSize: 12, bold: true, color: C.white,
    });
    sh(s, 'ellipse', { x: 5.333, y: 1.540 + dy, w: 0.204, h: 0.204, fill: { color: color }, line: { type: 'none' } });
    tx(s, LOREM, {
      x: 5.652, y: 1.404 + dy, w: 3.264, h: 0.477,
      fontSize: 10, color: C.text, bullet: BULLET, lineSpacingMultiple: 1.5,
    });
  });
}

/* =============================================================== slide 07 */
// Eight satellite circles orbiting a grey hub.
function slide07(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  const nodes = [
    { x: 3.845, y: 1.169, color: C.a1, lx: 1.333, ly: 1.612, lw: 2.332, al: 'right', tab: [0.420, 1.459, 3.554], flip: true },
    { x: 5.069, y: 1.169, color: C.a2, lx: 6.335, ly: 1.612, lw: 2.860, al: 'left', tab: [6.022, 1.459, 3.554], flip: false },
    { x: 5.935, y: 2.035, color: C.a3, lx: 7.133, ly: 2.478, lw: 2.299, al: 'left', tab: [6.823, 2.325, 2.753], flip: false },
    { x: 5.935, y: 3.261, color: C.a4, lx: 7.133, ly: 3.703, lw: 2.299, al: 'left', tab: [6.823, 3.551, 2.753], flip: false },
    { x: 5.069, y: 4.127, color: C.a5, lx: 6.335, ly: 4.569, lw: 2.860, al: 'left', tab: [6.022, 4.416, 3.554], flip: false },
    { x: 3.845, y: 4.127, color: C.a6, lx: 0.805, ly: 4.569, lw: 2.860, al: 'right', tab: [0.420, 4.416, 3.554], flip: true },
    { x: 2.980, y: 3.261, color: C.a4, lx: 0.567, ly: 3.703, lw: 2.299, al: 'right', tab: [0.420, 3.551, 2.753], flip: true },
    { x: 2.980, y: 2.035, color: C.a3, lx: 0.567, ly: 2.478, lw: 2.299, al: 'right', tab: [0.420, 2.325, 2.753], flip: true },
  ];
  nodes.forEach(function (n) {
    sh(s, 'roundRect', {
      x: n.tab[0], y: n.tab[1], w: n.tab[2], h: 0.507, rectRadius: 0.085,
      fill: { color: 'F7F7F7' }, line: { color: C.g85, width: 1 },
    });
    tx(s, STD, { x: n.lx, y: n.ly, w: n.lw, h: 0.202, align: n.al, fontSize: 12, bold: true, color: C.text });
  });
  // dashed spokes from the hub to each satellite
  const hub = { x: 5.0, y: 3.191 };
  nodes.forEach(function (n) {
    const cx = n.x + 0.543, cy = n.y + 0.543;
    const dx = cx - hub.x, dy = cy - hub.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    const p0 = { x: hub.x + dx / len * 0.74, y: hub.y + dy / len * 0.74 };
    const p1 = { x: cx - dx / len * 0.545, y: cy - dy / len * 0.545 };
    sh(s, 'line', {
      x: Math.min(p0.x, p1.x), y: Math.min(p0.y, p1.y),
      w: Math.abs(p1.x - p0.x), h: Math.abs(p1.y - p0.y),
      flipH: (p1.x - p0.x) * (p1.y - p0.y) < 0,
      line: { color: C.g50, width: 1, dashType: 'sysDash' },
    });
  });
  nodes.forEach(function (n) {
    sh(s, 'ellipse', { x: n.x, y: n.y, w: 1.086, h: 1.087, fill: { color: n.color }, line: { type: 'none' } });
    ring(s, n.x + 0.098, n.y + 0.098, 0.89, 3.6, C.white);
    icon(s, n.x + 0.304, n.y + 0.304, 0.478, C.white);
  });
  sh(s, 'ellipse', { x: 4.268, y: 2.459, w: 1.464, h: 1.464, fill: { color: C.g50 }, line: { type: 'none' } });
  ring(s, 4.333, 2.524, 1.333, 5.4, C.white);
  tx(s, STD, { x: 4.464, y: 3.014, w: 1.073, h: 0.353, align: 'center', fontSize: 10.5, bold: true, color: C.white });
}

/* =============================================================== slide 08 */
// Four full-height colour columns.
function slide08(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  ACCENTS.slice(0, 4).forEach(function (color, i) {
    const x = 0.424 + i * 2.308;
    sh(s, 'rect', { x: x, y: 1.146, w: 2.229, h: 4.083, fill: { color: color }, line: { type: 'none' } });
    sh(s, 'rect', {
      x: x + 0.073, y: 1.222, w: 2.083, h: 4.007,
      fill: { color: C.white, transparency: 85 }, line: { color: C.white, width: 1 },
    });
    tx(s, STD, {
      x: x + 0.367, y: 1.531, w: 1.494, h: 0.404, align: 'center',
      fontSize: 12, bold: true, color: C.white,
    });
    sh(s, 'ellipse', {
      x: x + 0.63, y: 2.099, w: 0.969, h: 0.969,
      fill: { color: C.white, transparency: 78 }, line: { type: 'none' },
    });
    sh(s, 'ellipse', { x: x + 0.717, y: 2.186, w: 0.794, h: 0.794, fill: { color: C.white }, line: { type: 'none' }, shadow: cardShadow() });
    icon(s, x + 0.83, 2.301, 0.566, color);
    tx(s, 'This is a sample text. You simply add your own text and description here. This text is fully editable. ', {
      x: x + 0.245, y: 3.282, w: 1.74, h: 0.982, align: 'center',
      fontSize: 10, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
    });
    sh(s, 'line', { x: x + 0.159, y: 4.878, w: 1.917, h: 0, line: { color: C.white, width: 1.5 } });
    sh(s, 'ellipse', { x: x + 0.833, y: 4.597, w: 0.562, h: 0.562, fill: { color: C.white }, line: { type: 'none' }, shadow: cardShadow() });
    tx(s, '0' + (i + 1), {
      x: x + 0.866, y: 4.597, w: 0.497, h: 0.562, align: 'center',
      fontSize: 14, bold: true, color: C.text,
    });
  });
}

/* =============================================================== slide 09 */
// Three concentric-circle stations linked by thin arcs.
function slide09(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  const cols = [
    { x: 0.879, up: true,  color: C.a1, halo: ACCENT_LIGHT[0] },
    { x: 3.808, up: false, color: C.a2, halo: ACCENT_LIGHT[1] },
    { x: 6.736, up: true,  color: C.a3, halo: ACCENT_LIGHT[2] },
  ];
  cols.forEach(function (c) {
    const panelY = c.up ? 1.670 : 1.196;
    const circY = c.up ? 1.670 : 2.167;
    // pale panel: the two corners next to the circle are fully rounded
    poly(s, rectPath(2.385, 3.356, c.up ? [1.1925, 1.1925, 0, 0] : [0, 0, 1.1925, 1.1925]), {
      x: c.x, y: panelY, w: 2.385, h: 3.356, fill: { color: C.g95 }, line: { type: 'none' },
    });
    sh(s, 'arc', {
      x: c.x - 0.272, y: circY - 0.272, w: 2.928, h: 2.928, rotate: c.up ? 0 : 180,
      angleRange: [172.2, 8.3], fill: { type: 'none' },
      line: c.up
        ? { color: c.color, width: 1, endArrowType: 'triangle' }
        : { color: c.color, width: 1, beginArrowType: 'triangle' },
    });
    sh(s, 'ellipse', { x: c.x, y: circY, w: 2.385, h: 2.385, fill: { color: c.halo }, line: { type: 'none' }, shadow: cardShadow() });
    sh(s, 'ellipse', { x: c.x + 0.284, y: circY + 0.284, w: 1.817, h: 1.817, fill: { color: c.color }, line: { type: 'none' }, shadow: cardShadow() });
    tx(s, 'Placeholder text', {
      x: c.x + 0.521, y: circY + (c.up ? 1.493 : 0.714), w: 1.344, h: 0.177,
      align: 'center', fontSize: 10.5, bold: true, color: C.white,
    });
    icon(s, c.x + 0.872, circY + (c.up ? 0.638 : 1.134), 0.642, C.white);
    tx(s, 'Contrary to popular belief, Lorem Ipsum is not simply random text. '
         + 'It has roots in a piece of classical Latin re', {
      x: c.x + 0.284, y: c.up ? 4.196 : 1.140, w: 1.817, h: 0.884,
      fontSize: 9, color: C.text, bullet: BULLET, lineSpacingMultiple: 1.5,
    });
  });
}

/* =============================================================== slide 10 */
// Eight pills pointing at a central bubble.
const P_LEADER_LONG = [['M', 0.5283, 0], ['L', 0.0165, 0.9519], ['L', 0.0165, 0.8972],
  ['L', 0, 0.9278], ['L', 0, 1], ['L', 0.0377, 1], ['L', 0.0542, 0.9694], ['L', 0.0248, 0.9694],
  ['L', 0.5342, 0.0241], ['L', 1, 0.0241], ['L', 1, 0], ['Z']];
const P_LEADER_SHORT = [['M', 0.0369, 0.8228], ['L', 0.0527, 0.6772], ['L', 0.0211, 0.7278],
  ['L', 0, 0.9177], ['L', 0.0527, 1], ['L', 0.0844, 0.9494], ['L', 0.0439, 0.8861],
  ['L', 0.5782, 0.0696], ['L', 1, 0.0696], ['L', 1, 0], ['L', 0.5729, 0], ['Z']];

function slide10(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  const rows = [
    { left: C.a1, right: C.a2, labelW: 1.816 },
    { left: C.a4, right: C.a3, labelW: 1.457 },
    { left: C.a3, right: C.a4, labelW: 1.816 },
    { left: C.a6, right: C.a5, labelW: 1.541 },
  ];
  rows.forEach(function (r, i) {
    const y = 1.231 + i * 1.0565;
    // left pill
    poly(s, P_PILL_L, { x: 0.752, y: y, w: 2.497, h: 0.719, fill: { color: r.left }, line: { type: 'none' } });
    sh(s, 'ellipse', { x: 2.622, y: y + 0.079, w: 0.561, h: 0.561, fill: { color: C.white, transparency: 82 }, line: { type: 'none' } });
    icon(s, 2.719, y + 0.176, 0.366, C.white);
    tx(s, 'Placeholder text', {
      x: 2.496 - r.labelW, y: y + 0.258, w: r.labelW, h: 0.202,
      align: 'right', fontSize: 12, color: C.white,
    });
    // right pill (mirrored)
    poly(s, P_PILL_L, { x: 6.752, y: y, w: 2.497, h: 0.719, flipH: true, fill: { color: r.right }, line: { type: 'none' } });
    sh(s, 'ellipse', { x: 6.818, y: y + 0.079, w: 0.561, h: 0.561, fill: { color: C.white, transparency: 82 }, line: { type: 'none' } });
    icon(s, 6.915, y + 0.176, 0.366, C.white);
    tx(s, 'Placeholder text', { x: 7.515, y: y + 0.258, w: 1.816, h: 0.202, fontSize: 12, color: C.white });
  });
  // thin arrow leaders converging on the hub
  [{ x: 5.280, y: 1.597, fh: false, fv: false }, { x: 5.280, y: 3.960, fh: false, fv: true },
   { x: 3.248, y: 1.597, fh: true,  fv: false }, { x: 3.248, y: 3.960, fh: true,  fv: true }]
    .forEach(function (a) {
      poly(s, P_LEADER_LONG, { x: a.x, y: a.y, w: 1.472, h: 0.793, flipH: a.fh, flipV: a.fv, fill: { color: C.g50 }, line: { type: 'none' } });
    });
  [{ x: 5.764, y: 2.637, fh: false, fv: false }, { x: 5.764, y: 3.439, fh: false, fv: true },
   { x: 3.248, y: 2.637, fh: true,  fv: false }, { x: 3.248, y: 3.439, fh: true,  fv: true }]
    .forEach(function (a) {
      poly(s, P_LEADER_SHORT, { x: a.x, y: a.y, w: 0.988, h: 0.274, flipH: a.fh, flipV: a.fv, fill: { color: C.g50 }, line: { type: 'none' } });
    });
  sh(s, 'ellipse', { x: 4.207, y: 2.383, w: 1.586, h: 1.584, fill: { color: C.g85 }, line: { type: 'none' } });
  sh(s, 'ellipse', { x: 4.290, y: 2.465, w: 1.420, h: 1.420, fill: { color: C.white }, line: { type: 'none' } });
  tx(s, 'You can use this sentence', {
    x: 4.442, y: 2.990, w: 1.117, h: 0.37, align: 'center', fontSize: 11, color: C.text,
  });
}

/* =============================================================== slide 11 */
// RACI-style matrix: label column + three colour-coded grids.
function slide11(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  const GRID = [
    ['', '', '', '', ''], ['34', '34', '', '34', ''], ['', '', '', '', ''],
    ['I', '', '34', '', ''], ['I', '34', 'I', 'I', ''], ['', '34', '', '', ''],
    ['I', 'I', 'I', 'I', ''],
  ];
  const GRID2 = [
    ['', '', '', '', ''], ['34', '34', '', 'C', ''], ['R', '', '', '', ''],
    ['34', '34', '34', 'C', ''], ['34', 'C', 'C', 'C', ''], ['', '', '', '', ''],
    ['34', 'C', 'C', 'C', 'C'],
  ];
  const GRID3 = [
    ['', '', '', '', ''], ['34', '', '34', '', '34'], ['', '', '', '', ''],
    ['34', '', 'C', '34', ''], ['', '', 'C', '', ''], ['34', '', '', '34', ''],
    ['C', 'C', 'C', '', ''],
  ];

  // label column
  sh(s, 'roundRect', { x: 0.424, y: 1.683, w: 2.256, h: 0.307, rectRadius: 0.15, fill: { color: C.g75 }, line: { type: 'none' } });
  tx(s, 'TEXT HERE', { x: 0.702, y: 1.748, w: 1.701, h: 0.177, align: 'center', fontSize: 10.5, bold: true, color: C.white });
  sh(s, 'round2DiagRect', { x: 0.424, y: 2.101, w: 2.256, h: 0.307, fill: { color: C.a1 }, line: { type: 'none' } });
  tx(s, 'Add Text Here', { x: 0.702, y: 2.171, w: 1.701, h: 0.168, align: 'center', fontSize: 10, bold: true, color: C.white });
  for (let r = 0; r < 7; r++) {
    const y = 2.455 + r * 0.3537;
    sh(s, 'rect', { x: 0.424, y: y, w: 2.256, h: 0.307, fill: { color: C.g95 }, line: { type: 'none' } });
    tx(s, 'Description Text', { x: 0.702, y: y + 0.077, w: 1.701, h: 0.151, align: 'center', fontSize: 9, bold: true, color: C.text });
  }

  // three grids
  const grids = [
    { x: 2.723, color: C.a2, tint: ACCENT_PALE[1], data: GRID },
    { x: 5.021, color: C.a3, tint: ACCENT_PALE[2], data: GRID2 },
    { x: 7.319, color: C.a4, tint: ACCENT_PALE[3], data: GRID3 },
  ];
  grids.forEach(function (g, gi) {
    sh(s, 'round2DiagRect', { x: g.x, y: 2.104, w: 2.256, h: 0.307, fill: { color: g.color }, line: { type: 'none' } });
    tx(s, 'Add Text Here', { x: g.x + 0.277, y: 2.173, w: 1.701, h: 0.168, align: 'center', fontSize: 10, bold: true, color: C.white });
    const dash = { type: 'dash', color: C.g50, pt: 0.75 };
    const rows = g.data.map(function (row) {
      return row.map(function (t, ci) {
        return {
          text: t,
          options: {
            fill: { color: ci % 2 === 0 ? C.g95 : g.tint },
            border: [dash, dash, dash, dash],
            fontSize: 9, color: C.text, align: 'center',
          },
        };
      });
    });
    s.addTable(rows, {
      x: g.x, y: 2.455, w: 2.256, colW: [0.4512, 0.4512, 0.4512, 0.4512, 0.4512],
      rowH: 0.3467, valign: 'middle', fontFace: FONT, margin: [1, 3, 1, 3],
    });
    // column headers (rotated tabs above the grid)
    for (let c = 0; c < 5; c++) {
      const cx = g.x - 0.322 + c * 0.4669;
      sh(s, 'rect', { x: cx, y: 1.391, w: 1.039, h: 0.382, rotate: 270, fill: { color: C.white }, line: { color: g.color, width: 0.75 } });
      tx(s, (gi === 2 && c === 4) ? 'Role #6' : 'Example Text', {
        x: cx + 0.034, y: 1.523, w: 0.971, h: 0.118, rotate: 270,
        align: 'center', fontSize: 7, color: C.text,
      });
    }
    // footer
    sh(s, 'round2SameRect', { x: g.x, y: 4.929, w: 2.256, h: 0.307, rotate: 180, fill: { color: g.color }, line: { type: 'none' } });
    tx(s, STD, { x: g.x + 0.24, y: 5.007, w: 1.776, h: 0.151, align: 'center', fontSize: 9, bold: true, color: C.white });
  });
  sh(s, 'round2SameRect', { x: 0.424, y: 4.929, w: 2.256, h: 0.307, rotate: 180, fill: { color: C.a1 }, line: { type: 'none' } });
  tx(s, STD, { x: 0.664, y: 5.007, w: 1.776, h: 0.151, align: 'center', fontSize: 9, bold: true, color: C.white });
}

/* =============================================================== slide 12 */
// Six chevrons with alternating above/below captions.
function slide12(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  const captions = [
    { x: 0.529, above: true,  w: 2.174 },
    { x: 1.968, above: false, w: 2.235 },
    { x: 3.518, above: true,  w: 2.253 },
    { x: 4.935, above: false, w: 2.185 },
    { x: 6.485, above: true,  w: 2.218 },
    { x: 7.902, above: false, w: 1.674 },
  ];
  ACCENTS.forEach(function (color, i) {
    poly(s, chevronPath(0.271), {
      x: 0.551 + i * 1.4835, y: 2.571, w: 1.417, h: 1.333,
      fill: { color: color }, line: { type: 'none' }, shadow: shadow(7, 9, 0, 0.16),
    });
    icon(s, 1.045 + i * 1.4835, 2.995, 0.52, C.white);
  });
  captions.forEach(function (c, i) {
    const y = c.above ? 1.341 : 4.544;
    tx(s, 'Description Text', { x: c.x, y: y, w: c.w, h: 0.185, fontSize: 11, bold: true, color: ACCENTS[i] });
    tx(s, i === 5 ? 'is simply dummy text of the printing and typesetting'
                  : 'is simply dummy text of the printing and typesetting industry. Lorem', {
      x: c.x, y: y + 0.221, w: i === 5 ? 1.674 : 2.018, h: 0.454,
      fontSize: 9, color: C.text, lineSpacingMultiple: 1.5, valign: 'top',
    });
    // dotted leader between caption and chevron
    sh(s, 'line', {
      x: c.x + 0.587, y: c.above ? 2.125 : 3.905, w: 0, h: 0.429,
      line: { color: C.g50, width: 1, dashType: 'dash', beginArrowType: 'oval', endArrowType: 'oval' },
    });
  });
}

/* =============================================================== slide 13 */
// Four quadrant cards around a cross axis.
function slide13(pptx) {
  const s = pptx.addSlide();
  heading(s, 'center');
  const cards = [
    { x: 1.611, y: 1.455, color: C.a1, ix: 1.127, iy: 1.804, tx: 2.250 },
    { x: 5.083, y: 1.455, color: C.a2, ix: 7.904, iy: 1.804, tx: 5.833 },
    { x: 1.611, y: 3.255, color: C.a4, ix: 1.127, iy: 3.604, tx: 2.250 },
    { x: 5.083, y: 3.255, color: C.a3, ix: 7.904, iy: 3.604, tx: 5.833 },
  ];
  cards.forEach(function (c) {
    sh(s, 'roundRect', { x: c.x, y: c.y, w: 3.306, h: 1.667, rectRadius: 0.11, fill: { color: c.color }, line: { type: 'none' } });
    sh(s, 'roundRect', {
      x: c.x + 0.09, y: c.y + 0.088, w: 3.125, h: 1.49, rectRadius: 0.05,
      fill: { color: C.white, transparency: 85 }, line: { color: C.white, width: 0.75 },
    });
    sh(s, 'ellipse', { x: c.ix, y: c.iy, w: 0.969, h: 0.969, fill: { color: c.color }, line: { type: 'none' }, shadow: cardShadow() });
    sh(s, 'ellipse', { x: c.ix + 0.087, y: c.iy + 0.087, w: 0.794, h: 0.794, fill: { color: C.white }, line: { type: 'none' } });
    icon(s, c.ix + 0.288, c.iy + 0.288, 0.392, c.color);
    tx(s, 'Description Text', { x: c.tx, y: c.y + 0.732, w: 1.917, h: 0.202, align: 'center', fontSize: 12, bold: true, color: C.white });
  });
  // cross axis
  sh(s, 'line', { x: 5.0, y: 1.455, w: 0, h: 3.467, line: { color: C.text, width: 1, beginArrowType: 'triangle', endArrowType: 'triangle' } });
  sh(s, 'line', { x: 1.556, y: 3.188, w: 6.889, h: 0, line: { color: C.text, width: 1, beginArrowType: 'triangle', endArrowType: 'triangle' } });
  [['Text Here', 4.042, 1.195, 1.917], ['Text Here', 4.042, 5.019, 1.917],
   ['Text Here', 0.302, 3.096, 1.262], ['Text Here', 8.436, 3.096, 1.262]].forEach(function (t) {
    tx(s, t[0], { x: t[1], y: t[2], w: t[3], h: 0.185, align: 'center', fontSize: 11, bold: true, color: C.text });
  });
  // circulation arrows (clockwise around the quadrants)
  sh(s, 'rightArrow', { x: 4.598, y: 2.128, w: 0.892, h: 0.397, fill: { color: C.a1 }, line: { type: 'none' } });
  sh(s, 'downArrow', { x: 5.833, y: 2.763, w: 0.397, h: 0.892, fill: { color: C.a2 }, line: { type: 'none' } });
  sh(s, 'leftArrow', { x: 4.510, y: 3.852, w: 0.892, h: 0.397, fill: { color: C.a3 }, line: { type: 'none' } });
  sh(s, 'upArrow', { x: 3.570, y: 2.720, w: 0.397, h: 0.892, fill: { color: C.a4 }, line: { type: 'none' } });
}

/* =============================================================== slide 14 */
// Organisation chart.
function slide14(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  // root
  sh(s, 'roundRect', { x: 4.281, y: 1.174, w: 3.833, h: 0.583, rectRadius: 0.291, fill: { color: C.g50 }, line: { type: 'none' } });
  tx(s, STD, { x: 4.926, y: 1.364, w: 2.543, h: 0.202, align: 'center', fontSize: 12, bold: true, color: C.white });
  sh(s, 'ellipse', { x: 7.565, y: 1.215, w: 0.5, h: 0.5, fill: { color: C.white }, line: { type: 'none' }, shadow: cardShadow() });
  icon(s, 7.641, 1.291, 0.348, C.text);

  const nodes = [
    { x: 3.040, y: 2.002, color: C.a1 }, { x: 7.820, y: 2.002, color: C.a1 },
    { x: 0.645, y: 2.888, color: C.a2 }, { x: 3.040, y: 2.888, color: C.a2 }, { x: 5.435, y: 2.888, color: C.a2 },
    { x: 0.645, y: 3.774, color: C.a3 }, { x: 3.040, y: 3.774, color: C.a3 },
    { x: 0.645, y: 4.661, color: C.a4 }, { x: 3.040, y: 4.661, color: C.a4 }, { x: 5.435, y: 4.661, color: C.a4 },
  ];
  nodes.forEach(function (n) {
    sh(s, 'roundRect', { x: n.x, y: n.y, w: 1.535, h: 0.531, rectRadius: 0.09, fill: { color: n.color }, line: { type: 'none' } });
    tx(s, 'Text Here', { x: n.x + 0.058, y: n.y + 0.182, w: 1.419, h: 0.168, align: 'center', fontSize: 10, color: C.white });
  });

  // level-to-level splitters (parent tick down into a shared bar, then into each child)
  splitter(s, 3.808, 1.685, 4.780, 0.317, 6.198, C.g50);
  splitter(s, 1.412, 2.571, 4.790, 0.317, 3.808, C.g50);
  splitter(s, 1.412, 3.457, 2.396, 0.317, 1.890, C.g50);
  splitter(s, 1.412, 4.344, 4.790, 0.317, 1.890, C.g50);
  // arrow heads where each splitter meets its child, plus the two straight drops
  const drop = { color: C.g50, width: 0.75, endArrowType: 'triangle' };
  [[3.808, 1.940], [8.588, 1.940], [1.412, 2.826], [6.202, 2.826],
   [1.412, 3.712], [3.808, 3.712], [1.412, 4.599], [6.202, 4.599]].forEach(function (a) {
    sh(s, 'line', { x: a[0], y: a[1] - 0.06, w: 0, h: 0.062, line: drop });
  });
  sh(s, 'line', { x: 3.808, y: 2.480, w: 0, h: 0.409, line: drop });
  sh(s, 'line', { x: 3.808, y: 4.446, w: 0, h: 0.214, line: drop });
}

/* =============================================================== slide 15 */
// Four-segment circular process (pinwheel) with side notes.
function slide15(pptx) {
  const s = pptx.addSlide();
  const notes = [
    { x: 0.425, y: 1.392, tx: 0.733, ty: 1.603, lx: 0.584, color: C.a1 },
    { x: 5.155, y: 1.392, tx: 7.333, ty: 1.603, lx: 9.416, color: C.a2 },
    { x: 0.425, y: 3.446, tx: 0.733, ty: 3.657, lx: 0.584, color: C.a4 },
    { x: 5.155, y: 3.446, tx: 7.333, ty: 3.657, lx: 9.416, color: C.a3 },
  ];
  notes.forEach(function (n) {
    sh(s, 'round1Rect', {
      x: n.x, y: n.y, w: 4.42, h: 1.533, rectRadius: 0.766, rotate: 180,
      fill: { color: C.g95 }, line: { color: C.white, width: 1 }, shadow: deepShadow(),
    });
    tx(s, LOREM_LONG, {
      x: n.tx, y: n.ty, w: n.tx < 5 ? 1.999 : 1.851, h: 1.111,
      fontSize: 9, color: C.text, bullet: BULLET, lineSpacingMultiple: 1.5,
    });
    sh(s, 'line', { x: n.lx, y: n.y + 0.058, w: 0, h: 1.417, line: { color: n.color, width: 2.75 } });
  });
  heading(s, 'center');

  sh(s, 'ellipse', { x: 4.210, y: 2.401, w: 1.575, h: 1.575, fill: { color: C.g95 }, line: { type: 'none' } });
  // outer swirl blades
  const blades = [
    { p: P_BLADE_TL, x: 3.083, y: 1.165, w: 2.090, h: 2.132, color: C.a1 },
    { p: P_BLADE_BL, x: 2.972, y: 3.012, w: 2.132, h: 2.089, color: C.a4 },
    { p: P_BLADE_BR, x: 4.819, y: 3.080, w: 2.090, h: 2.132, color: C.a3 },
    { p: P_BLADE_TR, x: 4.889, y: 1.276, w: 2.132, h: 2.089, color: C.a2 },
  ];
  blades.forEach(function (b) { poly(s, b.p, { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: b.color }, line: { type: 'none' } }); });
  const inner = [
    { p: P_INNER_TL, x: 3.578, y: 1.760, w: 1.595, h: 1.503, color: ACCENT_MID[0] },
    { p: P_INNER_BL, x: 3.569, y: 3.012, w: 1.503, h: 1.594, color: ACCENT_MID[3] },
    { p: P_INNER_BR, x: 4.819, y: 3.113, w: 1.595, h: 1.503, color: ACCENT_MID[2] },
    { p: P_INNER_TR, x: 4.922, y: 1.771, w: 1.503, h: 1.594, color: ACCENT_MID[1] },
  ];
  inner.forEach(function (b) { poly(s, b.p, { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: b.color }, line: { type: 'none' } }); });
  sh(s, 'ellipse', { x: 4.293, y: 2.484, w: 1.408, h: 1.408, fill: { color: C.white }, line: { type: 'none' } });
  tx(s, 'You can use this sentence', {
    x: 4.381, y: 2.986, w: 1.232, h: 0.404, align: 'center', fontSize: 12, bold: true, color: C.text,
  });

  const labels = [
    { t: 'Add Text Here', x: 2.999, y: 2.054, w: 1.580, r: -45, sz: 10.5 },
    { t: 'Add Text Here', x: 5.487, y: 4.239, w: 1.335, r: -45, sz: 10.5 },
    { t: 'Add Text Here', x: 3.238, y: 4.260, w: 1.149, r: 45, sz: 10.5 },
    { t: 'Add Text Here', x: 5.479, y: 1.971, w: 1.323, r: 45, sz: 10.5 },
    { t: 'Text Here', x: 3.663, y: 2.293, w: 0.940, r: -45, sz: 9 },
    { t: 'Text Here', x: 5.418, y: 3.951, w: 0.810, r: -45, sz: 9 },
    { t: 'Text Here', x: 5.476, y: 2.283, w: 0.747, r: 45, sz: 9 },
    { t: 'Text Here', x: 3.545, y: 3.913, w: 1.233, r: 45, sz: 9 },
  ];
  labels.forEach(function (l) {
    tx(s, l.t, {
      x: l.x, y: l.y, w: l.w, h: 0.177, rotate: l.r,
      align: 'center', fontSize: l.sz, bold: true, color: C.white,
    });
  });
}

/* =============================================================== slide 16 */
// Two opposing block arrows: entry / exit criteria.
function slide16(pptx) {
  const s = pptx.addSlide();
  heading(s, 'left');
  // soft vertical divider
  sh(s, 'rect', { x: 4.916, y: 2.280, w: 0.168, h: 1.703, fill: { color: C.g95 }, line: { type: 'none' } });

  poly(s, P_BIG_ARROW, { x: 4.079, y: 1.752, w: 3.337, h: 1.535, fill: { color: C.a3 }, line: { type: 'none' } });
  poly(s, P_BIG_ARROW, { x: 2.583, y: 3.029, w: 3.340, h: 1.537, flipH: true, fill: { color: C.a4 }, line: { type: 'none' } });
  fold(s, 'br', 4.079, 1.752, 0.806, 0.407, 85);
  fold(s, 'bl', 5.118, 3.029, 0.806, 0.407, 85);
  tx(s, 'Exit Criteria ', { x: 4.712, y: 2.402, w: 1.617, h: 0.236, align: 'center', fontSize: 14, color: C.white });
  tx(s, 'Entry Criteria ', { x: 3.742, y: 3.680, w: 1.368, h: 0.236, align: 'center', fontSize: 14, color: C.white });
  icon(s, 6.400, 2.320, 0.39, C.white);
  icon(s, 3.140, 3.598, 0.39, C.white);

  const body = 'You can use this sentence in any context where you are referring to some text as an example. '
    + 'For example, if you are writing a paper about writing techniques, you could use the sentence '
    + '"This is a sample text" to';
  tx(s, STD, { x: 0.424, y: 1.443, w: 2.55, h: 0.202, fontSize: 12, bold: true, color: C.a4 });
  tx(s, body, {
    x: 0.424, y: 1.657, w: 2.832, h: 1.47, fontSize: 9, color: C.text,
    bullet: BULLET, lineSpacingMultiple: 2, valign: 'top',
  });
  tx(s, STD, { x: 6.175, y: 3.753, w: 2.55, h: 0.202, fontSize: 12, bold: true, color: C.a3 });
  tx(s, body, {
    x: 6.175, y: 3.995, w: 3.242, h: 1.168, fontSize: 9, color: C.text,
    bullet: BULLET, lineSpacingMultiple: 2, valign: 'top',
  });
  // L-shaped leaders tying each text block to the arrow it describes
  const conn = { color: C.g85, width: 2.25 };
  sh(s, 'line', { x: 1.840, y: 3.810, w: 0.743, h: 0, line: Object.assign({ endArrowType: 'triangle' }, conn) });
  sh(s, 'line', { x: 1.840, y: 3.127, w: 0, h: 0.683, line: Object.assign({ beginArrowType: 'triangle' }, conn) });
  sh(s, 'line', { x: 6.807, y: 2.981, w: 1.060, h: 0, line: Object.assign({ beginArrowType: 'triangle' }, conn) });
  sh(s, 'line', { x: 7.867, y: 2.981, w: 0, h: 0.900, line: Object.assign({ endArrowType: 'triangle' }, conn) });
}

/* ============================================== slide-15 blade geometries */

const P_BLADE_TL = [['M', 1, 0.2842], ['C', 0.8636, 0.0538, 0.8636, 0.0538, 0.8636, 0.0538],
  ['L', 0.8307, 0], ['L', 0.8307, 0.0568],
  ['C', 0.3652, 0.0983, 0, 0.4823, 0, 0.9493], ['C', 0, 0.9662, 0, 0.9831, 0.0016, 1],
  ['L', 0.2367, 0.8664], ['L', 0.4451, 0.9846],
  ['C', 0.4451, 0.9739, 0.4436, 0.9616, 0.4436, 0.9493],
  ['C', 0.4436, 0.722, 0.6113, 0.533, 0.8307, 0.4946], ['L', 0.8307, 0.5684],
  ['L', 0.8793, 0.4885], ['L', 1, 0.2842], ['Z']];
const P_BLADE_BL = [['M', 0.8664, 0.7633], ['L', 0.9846, 0.5549],
  ['C', 0.9739, 0.5549, 0.9616, 0.5564, 0.9493, 0.5564],
  ['C', 0.722, 0.5564, 0.533, 0.3887, 0.4946, 0.1693], ['L', 0.5684, 0.1693],
  ['L', 0.4885, 0.1207], ['L', 0.2842, 0], ['L', 0.0538, 0.1364], ['L', 0, 0.1677],
  ['L', 0.0568, 0.1693], ['C', 0.0983, 0.6348, 0.4823, 1, 0.9493, 1],
  ['C', 0.9662, 1, 0.9831, 1, 1, 0.9984], ['L', 0.8664, 0.7633], ['Z']];
const P_BLADE_BR = [['M', 0.9984, 0], ['L', 0.7633, 0.1336], ['L', 0.5549, 0.0154],
  ['C', 0.5549, 0.0261, 0.5564, 0.0384, 0.5564, 0.0507],
  ['C', 0.5564, 0.278, 0.3887, 0.467, 0.1693, 0.5054], ['L', 0.1693, 0.4316],
  ['L', 0.1207, 0.5115], ['L', 0, 0.7158], ['L', 0.1364, 0.9462], ['L', 0.1677, 1],
  ['L', 0.1693, 0.9432], ['C', 0.6348, 0.9017, 1, 0.5177, 1, 0.0507],
  ['C', 1, 0.0338, 1, 0.0169, 0.9984, 0], ['Z']];
const P_BLADE_TR = [['M', 0.9432, 0.8307], ['C', 0.9017, 0.3652, 0.5177, 0, 0.0507, 0],
  ['C', 0.0338, 0, 0.0169, 0, 0, 0.0016], ['L', 0.1336, 0.2367], ['L', 0.0154, 0.4451],
  ['C', 0.0261, 0.4451, 0.0384, 0.4436, 0.0507, 0.4436],
  ['C', 0.278, 0.4436, 0.467, 0.6113, 0.5054, 0.8307], ['L', 0.4316, 0.8307],
  ['L', 0.5115, 0.8793], ['L', 0.7158, 1], ['L', 0.9462, 0.8636], ['L', 1, 0.8323],
  ['L', 0.9432, 0.8307], ['Z']];
const P_INNER_TL = [['M', 0, 0.8322], ['L', 0.2731, 1],
  ['C', 0.2731, 0.9847, 0.271, 0.9673, 0.271, 0.9499],
  ['C', 0.271, 0.6275, 0.4908, 0.3595, 0.7782, 0.305], ['L', 0.7782, 0.4096],
  ['L', 0.8419, 0.2963], ['L', 1, 0.0065],
  ['C', 0.963, 0.0022, 0.9261, 0, 0.8891, 0],
  ['C', 0.4312, 0, 0.0534, 0.3638, 0, 0.8322], ['Z']];
const P_INNER_BL = [['M', 1, 0.7269], ['C', 0.9847, 0.7269, 0.9673, 0.729, 0.9499, 0.729],
  ['C', 0.6275, 0.729, 0.3595, 0.5092, 0.305, 0.2218], ['L', 0.4096, 0.2218],
  ['L', 0.2963, 0.1581], ['L', 0.0065, 0],
  ['C', 0.0022, 0.037, 0, 0.0739, 0, 0.1109],
  ['C', 0, 0.5688, 0.3638, 0.9466, 0.8322, 1], ['L', 1, 0.7269], ['Z']];
const P_INNER_BR = [['M', 0.7269, 0], ['C', 0.7269, 0.0153, 0.729, 0.0327, 0.729, 0.0501],
  ['C', 0.729, 0.3725, 0.5092, 0.6405, 0.2218, 0.695], ['L', 0.2218, 0.5904],
  ['L', 0.1581, 0.7037], ['L', 0, 0.9935],
  ['C', 0.037, 0.9978, 0.0739, 1, 0.1109, 1],
  ['C', 0.5688, 1, 0.9466, 0.6362, 1, 0.1678], ['L', 0.7269, 0], ['Z']];
const P_INNER_TR = [['M', 0, 0.2731], ['C', 0.0153, 0.2731, 0.0327, 0.271, 0.0501, 0.271],
  ['C', 0.3725, 0.271, 0.6405, 0.4908, 0.695, 0.7782], ['L', 0.5904, 0.7782],
  ['L', 0.7037, 0.8419], ['L', 0.9935, 1],
  ['C', 0.9978, 0.963, 1, 0.9261, 1, 0.8891],
  ['C', 1, 0.4312, 0.6362, 0.0534, 0.1678, 0], ['L', 0, 0.2731], ['Z']];

/* -------------------------------------------------------------------- run */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
    .forEach(function (fn) { fn(pptx); });
  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '1118084b-8829-4491-ad36-710d1653c0db_grok_final.pptx'),
}).then(function (f) { console.log('wrote', f); });
