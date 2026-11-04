/**
 * "Alkes" medical presentation - 30 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt with pptxgenjs only. Photographs of the source deck are replaced by
 * flat dark "PLACE IMAGE HERE" placeholder blocks that keep the original
 * position, size and corner treatment.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette */
const C = {
  red: 'DC0000',
  peach: 'FFCA99',
  orange: 'FB9536',
  teal: '174240',
  yellow: 'E3DE00',
  ink: '202032',
  white: 'FFFFFF',
  photo: '2C2C2E',   // stand-in colour for photographs
  gray: 'D9D9D9',    // white lumMod 85%
  wash: 'F2F2F2',    // white lumMod 95%
  blue: '3F419A'
};
const HEAD = 'Roboto';      // theme major font
const BODY = 'Comfortaa';   // theme minor font
const NONE = { type: 'none' };
// pptxgenjs rewrites the shadow object it is handed, soevery shape needs its own copy.
const shadow = (opacity, offset, angle) =>
  ({ type: 'outer', color: '000000', opacity: opacity, blur: 13, offset: offset || 0, angle: angle || 0 });
const CARD_SHADOW = () => shadow(0.16, 0, 0);        // soft halo under white cards
const DROP_SHADOW = () => shadow(0.25, 7, 90);       // pills and badges
const SOFT_SHADOW = () => shadow(0.25, 0, 90);       // icon bubbles
const TILT_SHADOW = () => shadow(0.3, 7, 45);        // small peach circles

/* ------------------------------------------------------------ shared copy */
const LOREM = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';
const LOREM_LONG = LOREM + 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero.';
const EVERY = 'Every signs sixth midst place above multiply. Our signs moveth Of own grass his dry earth yielding were heaven doesn\u2019t days greater.';
const EVERY_SHORT = 'Every signs sixth midst place above multiply. ';
const AND_IS = 'And is, face, give made fruitful over creeping moving bearing fill cattle fifth saw. God form grass Set it creepeth beginning moved fish fill fill. Fowl sixth abundantly';
const AND_IS_SHORT = 'And is, face, give made fruitful over creeping moving bearing fill cattle fifth saw.';
const QUICK = 'The quick, brown fox jumps over a lazy dog. DJs flock by when MTV ax quiz prog. Junk MTV quiz graced by fox whelps. Bawds jog.';

/* -------------------------------------------------------- generic helpers */
const K = 0.5523;  // circle -> cubic bezier constant

/** Points for a rectangle with an individual radius (inches) per corner. */
function roundedPoints(w, h, r) {
  const tl = r.tl || 0, tr = r.tr || 0, br = r.br || 0, bl = r.bl || 0;
  const p = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }];
  if (tr) p.push({ x: w, y: tr, curve: { type: 'cubic', x1: w - tr + K * tr, y1: 0, x2: w, y2: tr - K * tr } });
  p.push({ x: w, y: h - br });
  if (br) p.push({ x: w - br, y: h, curve: { type: 'cubic', x1: w, y1: h - br + K * br, x2: w - br + K * br, y2: h } });
  p.push({ x: bl, y: h });
  if (bl) p.push({ x: 0, y: h - bl, curve: { type: 'cubic', x1: bl - K * bl, y1: h, x2: 0, y2: h - bl + K * bl } });
  p.push({ x: 0, y: tl });
  if (tl) p.push({ x: tl, y: 0, curve: { type: 'cubic', x1: 0, y1: tl - K * tl, x2: tl - K * tl, y2: 0 } });
  p.push({ close: true });
  return p;
}

/** Scale a normalised path spec (['M',x,y] / ['L',x,y] / ['C',x1,y1,x2,y2,x,y] / ['Z']). */
function scalePath(spec, w, h) {
  return spec.map(c => {
    if (c[0] === 'M') return { x: c[1] * w, y: c[2] * h, moveTo: true };
    if (c[0] === 'L') return { x: c[1] * w, y: c[2] * h };
    if (c[0] === 'C') {
      return { x: c[5] * w, y: c[6] * h, curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } };
    }
    return { close: true };
  });
}

/** Free-form shape from a normalised spec. */
function freeform(s, spec, o) {
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: scalePath(spec, o.w, o.h),
    fill: o.fill ? { color: o.fill, transparency: o.transparency } : NONE,
    line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1 } : NONE,
    shadow: o.shadow, rotate: o.rotate, flipH: o.flipH, flipV: o.flipV
  });
}

/** Rectangle with per-corner radii. */
function corners(s, o) {
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: roundedPoints(o.w, o.h, o.r || {}),
    fill: o.fill ? { color: o.fill, transparency: o.transparency } : NONE,
    line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1 } : NONE,
    shadow: o.shadow
  });
}

function rect(s, o) {
  s.addShape(o.rectRadius ? 'roundRect' : 'rect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.rectRadius,
    fill: o.fill ? { color: o.fill, transparency: o.transparency } : NONE,
    line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1 } : NONE,
    shadow: o.shadow, rotate: o.rotate
  });
}

function oval(s, o) {
  s.addShape('ellipse', {
    x: o.x, y: o.y, w: o.w, h: o.h || o.w,
    fill: o.fill ? { color: o.fill, transparency: o.transparency } : NONE,
    line: o.lineColor ? { color: o.lineColor, width: o.lineWidth || 1 } : NONE,
    shadow: o.shadow
  });
}

/* ------------------------------------------------------------ text helpers */
const R = (text, color, breakLine) => ({ text, options: { color, breakLine } });

function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BODY, fontSize: 12, color: C.ink, valign: 'middle' }, o));
}

/** 12pt justified body copy at 150% line spacing (the deck's default paragraph). */
function body(s, text, o) {
  txt(s, text, Object.assign({ align: 'justify', lineSpacingMultiple: 1.5 }, o));
}

/** 14pt bold sub-heading. */
function sub(s, text, o) {
  txt(s, text, Object.assign({ fontSize: 14, bold: true }, o));
}

/** 32pt bold two-tone section title (Roboto). */
function heading(s, runs, o) {
  s.addText(runs, Object.assign({ fontFace: HEAD, fontSize: 32, bold: true, valign: 'middle' }, o));
}

/** Peach "call to action" pill with centred caption. */
function pill(s, label, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.h / 2,
    fill: o.outline ? NONE : { color: o.fill || C.peach },
    line: o.outline ? { color: C.peach, width: 2.25 } : NONE,
    shadow: o.shadow
  });
  txt(s, label, {
    x: o.x, y: o.y, w: o.w, h: o.h, align: 'center',
    color: o.outline ? C.ink : (o.color || C.white), lineSpacingMultiple: 1.5
  });
}

/**
 * Dark block standing in for a photograph. `size` is the cap height of the
 * "PLACE IMAGE HERE" caption in points; `cx`/`cy` nudge it off the block centre.
 */
function photo(s, o) {
  const fill = o.fill || C.photo;
  if (o.spec) freeform(s, o.spec, { x: o.x, y: o.y, w: o.w, h: o.h, fill: fill, shadow: o.shadow });
  else if (o.r) corners(s, { x: o.x, y: o.y, w: o.w, h: o.h, r: o.r, fill: fill, shadow: o.shadow });
  else if (o.round === 'ellipse') oval(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: fill, shadow: o.shadow });
  else rect(s, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.rectRadius, fill: fill, shadow: o.shadow });
  if (o.size) {
    const bw = Math.max(o.w, o.size / 13), bh = o.size / 15;
    s.addText('PLACE\nIMAGE\nHERE', {
      x: o.x + o.w / 2 - bw / 2 + (o.cx || 0), y: o.y + o.h / 2 - bh / 2 + (o.cy || 0), w: bw, h: bh,
      align: 'center', valign: 'middle', fontFace: HEAD, bold: true,
      fontSize: o.size, color: o.captionColor || C.white, lineSpacing: o.size * 0.95
    });
  }
}

/**
 * Vertical fade from transparent (top) to saturated red (bottom), stacked from
 * translucent slices because pptxgenjs has no gradient fill. `r` re-rounds the
 * panel corners afterwards by masking the overflow with the page background.
 */
/**
 * Approximates the deck's signature red wash. The source gradient runs from a
 * transparent near-white at the top, through red at 60% opacity halfway down,
 * to solid red at the foot; both the colour and the alpha are interpolated.
 * pptxgenjs cannot emit gradient fills, so it is stacked from flat bands.
 */
function redFade(s, o) {
  const bands = 30, mid = 0.52, bh = o.h / bands;
  const hex = v => Math.round(v).toString(16).padStart(2, '0').toUpperCase();
  for (let i = 0; i < bands; i++) {
    const t = (i + 0.5) / bands;
    const alpha = t < mid ? 0.6 * t / mid : 0.6 + 0.4 * (t - mid) / (1 - mid);
    corners(s, {
      x: o.x, y: o.y + i * bh, w: o.w, h: bh,
      r: bandCorners(o.r, i, bands),
      fill: 'DA' + hex(218 * (1 - t)).repeat(2),   // pink washing out to pure red
      transparency: Math.round(100 - 100 * alpha)
    });
  }
}

/** Only the first and last bands of a fade need the panel's rounded corners. */
function bandCorners(r, i, bands) {
  if (!r) return {};
  if (i === 0) return { tl: r.tl, tr: r.tr };
  if (i === bands - 1) return { bl: r.bl, br: r.br };
  return {};
}

/** Grid of small dots used as a decorative accent. */
function dotGrid(s, o) {
  const d = 0.11;
  const stepX = o.cols > 1 ? (o.w - d) / (o.cols - 1) : 0;
  const stepY = o.rows > 1 ? (o.h - d) / (o.rows - 1) : 0;
  for (let r = 0; r < o.rows; r++) {
    for (let c = 0; c < o.cols; c++) {
      oval(s, { x: o.x + c * stepX, y: o.y + r * stepY, w: d, fill: o.color || C.peach });
    }
  }
}

/* --------------------------------------------------------------- vectors */
// Rounded plus / medical cross (the deck's recurring accent mark).
const CROSS = [
  ['M', 0.0, 0.5], ['C', 0.0, 0.4345, 0.0531, 0.3813, 0.1187, 0.3813], ['L', 0.3813, 0.3813],
  ['L', 0.3813, 0.1187], ['C', 0.3813, 0.0531, 0.4345, 0.0, 0.5, 0.0],
  ['C', 0.5655, 0.0, 0.6187, 0.0531, 0.6187, 0.1187], ['L', 0.6187, 0.3813], ['L', 0.8813, 0.3813],
  ['C', 0.9469, 0.3813, 1.0, 0.4345, 1.0, 0.5], ['C', 1.0, 0.5655, 0.9469, 0.6187, 0.8813, 0.6187],
  ['L', 0.6187, 0.6187], ['L', 0.6187, 0.8813], ['C', 0.6187, 0.9469, 0.5655, 1.0, 0.5, 1.0],
  ['C', 0.4345, 1.0, 0.3813, 0.9469, 0.3813, 0.8813], ['L', 0.3813, 0.6187], ['L', 0.1187, 0.6187],
  ['C', 0.0531, 0.6187, 0.0, 0.5655, 0.0, 0.5], ['Z']
];
// Chevron used for the small circular "next / previous" buttons.
const CHEVRON = [
  ['M', 0.166, 0.498], ['L', 0.996, 0.061], ['L', 0.996, 0.0], ['L', 0.887, 0.0], ['L', 0.0, 0.468],
  ['L', 0.0, 0.498], ['L', 0.0, 0.53], ['L', 0.887, 0.998], ['L', 0.996, 0.998], ['L', 0.996, 0.936], ['Z']
];
// "Alkes" logo mark: a shield/heart outline pierced by a cross.
const LOGO = [
  ['M', 0.5, 1.0], ['C', 0.4832, 1.0, 0.4673, 0.9915, 0.4571, 0.9772], ['L', 0.0307, 0.3759],
  ['C', 0.016, 0.3552, 0.016, 0.3265, 0.0307, 0.3057], ['L', 0.1654, 0.1158], ['L', 0.0539, 0.1158],
  ['C', 0.0242, 0.1158, 0.0, 0.0899, 0.0, 0.0579], ['C', 0.0, 0.0259, 0.0242, 0.0, 0.0539, 0.0],
  ['L', 0.2743, 0.0], ['C', 0.2948, 0.0, 0.3135, 0.0125, 0.3226, 0.0322],
  ['C', 0.3317, 0.0519, 0.3296, 0.0755, 0.3171, 0.093], ['L', 0.1414, 0.3408], ['L', 0.5, 0.8465],
  ['L', 0.8586, 0.3408], ['L', 0.6829, 0.093], ['C', 0.6705, 0.0755, 0.6684, 0.0519, 0.6774, 0.0322],
  ['C', 0.6865, 0.0125, 0.7052, 0.0, 0.7257, 0.0], ['L', 0.9461, 0.0],
  ['C', 0.9758, 0.0, 1.0, 0.0259, 1.0, 0.0579], ['C', 1.0, 0.0899, 0.9758, 0.1158, 0.9461, 0.1158],
  ['L', 0.8346, 0.1158], ['L', 0.9693, 0.3057], ['C', 0.984, 0.3265, 0.984, 0.3552, 0.9693, 0.3759],
  ['L', 0.5429, 0.9772], ['C', 0.5327, 0.9915, 0.5168, 1.0, 0.5, 1.0], ['Z']
];

function cross(s, x, y, size, color) {
  freeform(s, CROSS, { x: x, y: y, w: size, h: size, fill: color });
}
function chevron(s, o) {
  freeform(s, CHEVRON, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.color, flipH: o.flipH !== false });
}
/** Circular button with a chevron inside. */
function navButton(s, o) {
  const d = o.d || 0.257;
  oval(s, { x: o.x, y: o.y, w: d, fill: o.fill || C.peach, shadow: TILT_SHADOW() });
  chevron(s, {
    x: o.x + d * (o.flipH === false ? 0.36 : 0.42), y: o.y + d * 0.29,
    w: d * 0.23, h: d * 0.41, color: C.white, flipH: o.flipH
  });
}
function logo(s, x, y, size, color) {
  freeform(s, LOGO, { x: x, y: y, w: size, h: size * 0.931, fill: color });
  cross(s, x + size * 0.275, y + size * 0.139, size * 0.45, color);
}

/* ------------------------------------------------------------------ icons */
/* Flat vector stand-ins for the deck's red medical icon set. Every icon is
 * drawn inside a unit box (0..1) that `icon()` maps onto x/y/w/h.            */

/** Bar of `len` x `thick` centred on (cx, cy) and rotated `ang` degrees. */
function bar(u, cx, cy, len, thick, ang, color, round) {
  rect(u.s, {
    x: u.x + (cx - len / 2) * u.w, y: u.y + (cy - thick / 2) * u.h,
    w: len * u.w, h: thick * u.h, rotate: ang, fill: color,
    rectRadius: round ? thick * u.h / 2 : 0
  });
}
const box = (u, x, y, w, h, color, r) => rect(u.s, {
  x: u.x + x * u.w, y: u.y + y * u.h, w: w * u.w, h: h * u.h, fill: color, rectRadius: r ? r * u.w : 0
});
const disc = (u, cx, cy, d, color) => oval(u.s, {
  x: u.x + (cx - d / 2) * u.w, y: u.y + (cy - d / 2) * u.h, w: d * u.w, h: d * u.h, fill: color
});

const ICONS = {
  kit(u, c) {                                    // first-aid case
    box(u, 0.33, 0.02, 0.34, 0.18, c, 0.05);
    box(u, 0.0, 0.16, 1.0, 0.84, c, 0.12);
    box(u, 0.43, 0.36, 0.14, 0.44, C.white);
    box(u, 0.25, 0.51, 0.5, 0.14, C.white);
  },
  syringe(u, c) {                                // diagonal syringe
    bar(u, 0.13, 0.87, 0.3, 0.05, -45, c);
    bar(u, 0.45, 0.55, 0.48, 0.26, -45, c, true);
    bar(u, 0.62, 0.38, 0.32, 0.07, 45, c);
    bar(u, 0.76, 0.24, 0.24, 0.09, -45, c);
    bar(u, 0.88, 0.12, 0.28, 0.08, 45, c);
  },
  stethoscope(u, c) {                            // ear tubes (U) hooking into a J
    const arc = (x, y, w, h) => u.s.addShape('blockArc', {
      x: u.x + x * u.w, y: u.y + y * u.h, w: w * u.w, h: h * u.h,
      angleRange: [0, 180], arcThicknessRatio: 0.44, fill: { color: c }, line: NONE
    });
    arc(0.0, 0.14, 0.56, 0.56);                  // ear-tube U
    arc(0.4, 0.42, 0.52, 0.5);                   // chest-piece hook
    bar(u, 0.06, 0.2, 0.28, 0.12, 90, c, true);
    bar(u, 0.5, 0.24, 0.36, 0.12, 90, c, true);
    bar(u, 0.88, 0.5, 0.22, 0.12, 90, c, true);
    disc(u, 0.06, 0.07, 0.16, c);
    disc(u, 0.5, 0.07, 0.16, c);
    disc(u, 0.88, 0.36, 0.2, c);
  },
  ribbon(u, c) {                                 // awareness ribbon: loop + crossed legs
    bar(u, 0.66, 0.7, 0.68, 0.19, 52, c, true);
    bar(u, 0.34, 0.7, 0.68, 0.19, -52, c, true);
    u.s.addShape('donut', {
      x: u.x + 0.24 * u.w, y: u.y, w: 0.52 * u.w, h: 0.46 * u.h,
      fill: { color: c }, line: NONE
    });
  },
  dropper(u, c) {                                // blood collection tube
    bar(u, 0.48, 0.52, 0.78, 0.34, -45, c, true);
    bar(u, 0.86, 0.14, 0.3, 0.16, -45, c);
    freeform(u.s, [['M', 0, 0.55], ['L', 0.62, 0], ['L', 0.62, 0.55], ['Z']],
      { x: u.x + 0.28 * u.w, y: u.y + 0.2 * u.h, w: 0.44 * u.w, h: 0.44 * u.h, fill: C.white });
  },
  nurse(u, c) {                                  // head + shoulders with cap
    box(u, 0.31, 0.0, 0.38, 0.15, c, 0.04);
    box(u, 0.47, 0.02, 0.06, 0.11, C.white);
    box(u, 0.42, 0.045, 0.16, 0.045, C.white);
    disc(u, 0.5, 0.36, 0.42, c);
    freeform(u.s, [['M', 0, 1], ['C', 0, 0.12, 1, 0.12, 1, 1], ['Z']],
      { x: u.x, y: u.y + 0.6 * u.h, w: u.w, h: 0.4 * u.h, fill: c });
  },
  heart(u, c) {
    u.s.addShape('heart', { x: u.x, y: u.y, w: u.w, h: u.h, fill: { color: c }, line: NONE });
  },
  thumb(u, c) {
    box(u, 0.3, 0.36, 0.58, 0.56, c, 0.08);
    box(u, 0.04, 0.5, 0.22, 0.42, c, 0.06);
    bar(u, 0.55, 0.24, 0.42, 0.2, -70, c, true);
  },
  chat(u, c) {                                   // two overlapping speech bubbles
    oval(u.s, { x: u.x + 0.28 * u.w, y: u.y, w: 0.72 * u.w, h: 0.6 * u.h, fill: c });
    freeform(u.s, [['M', 0, 0], ['L', 1, 0], ['L', 0.9, 1], ['Z']],
      { x: u.x + 0.74 * u.w, y: u.y + 0.42 * u.h, w: 0.2 * u.w, h: 0.24 * u.h, fill: c });
    oval(u.s, { x: u.x, y: u.y + 0.34 * u.h, w: 0.62 * u.w, h: 0.52 * u.h, fill: c });
    freeform(u.s, [['M', 0, 0], ['L', 1, 0], ['L', 0.1, 1], ['Z']],
      { x: u.x + 0.08 * u.w, y: u.y + 0.74 * u.h, w: 0.2 * u.w, h: 0.26 * u.h, fill: c });
  },
  phone(u, c) {                                  // classic handset silhouette
    freeform(u.s, [['M', 0.2, 0.0], ['C', 0.0, 0.14, 0.0, 0.62, 0.3, 0.9],
      ['C', 0.56, 1.0, 0.86, 1.0, 1.0, 0.82], ['L', 0.72, 0.58], ['L', 0.55, 0.72],
      ['C', 0.4, 0.62, 0.3, 0.5, 0.26, 0.36], ['L', 0.44, 0.24], ['Z']],
      { x: u.x, y: u.y, w: u.w, h: u.h, fill: c });
  },
  calendar(u, c) {
    box(u, 0.0, 0.12, 1.0, 0.88, c, 0.1);
    box(u, 0.2, 0.0, 0.1, 0.2, c, 0.04);
    box(u, 0.7, 0.0, 0.1, 0.2, c, 0.04);
    check(u.s, { x: u.x + 0.2 * u.w, y: u.y + 0.4 * u.h, w: 0.6 * u.w, h: 0.42 * u.h, color: C.white });
  },
  pin(u, c) {                                    // map marker: disc tapering to a point
    freeform(u.s, [['M', 0.5, 1.0], ['C', 0.2, 0.62, 0.0, 0.52, 0.0, 0.36],
      ['C', 0.0, 0.16, 0.22, 0.0, 0.5, 0.0], ['C', 0.78, 0.0, 1.0, 0.16, 1.0, 0.36],
      ['C', 1.0, 0.52, 0.8, 0.62, 0.5, 1.0], ['Z']],
      { x: u.x, y: u.y, w: u.w, h: u.h, fill: c });
    disc(u, 0.5, 0.34, 0.38, C.white);
  },
  mail(u, c) {                                   // envelope with a folded flap
    box(u, 0.0, 0.0, 1.0, 1.0, c, 0.06);
    freeform(u.s, [['M', 0, 0], ['L', 1, 0], ['L', 0.5, 0.62], ['Z']],
      { x: u.x + 0.07 * u.w, y: u.y + 0.1 * u.h, w: 0.86 * u.w, h: 0.7 * u.h, lineColor: C.white, lineWidth: 1.4 });
  }
};
function icon(s, kind, o) {
  ICONS[kind]({ s: s, x: o.x, y: o.y, w: o.w, h: o.h }, o.color);
}

/** White tick mark built from two bars. */
function check(s, o) {
  rect(s, { x: o.x + 0.06 * o.w, y: o.y + 0.42 * o.h, w: 0.42 * o.w, h: 0.16 * o.h, fill: o.color, rotate: 45 });
  rect(s, { x: o.x + 0.3 * o.w, y: o.y + 0.3 * o.h, w: 0.66 * o.w, h: 0.16 * o.h, fill: o.color, rotate: -40 });
}

/** Header logo + wordmark + red slide-number bubble carried by the slide master. */
function chrome(s, num) {
  logo(s, 0.853, 0.501, 0.247, C.peach);
  txt(s, 'ALKES PRESENTATION', { x: 1.132, y: 0.451, w: 1.874, h: 0.269, fontSize: 10, bold: true, color: C.red });
  oval(s, { x: 12.143, y: 0.464, w: 0.305, fill: C.red });
  txt(s, String(num), { x: 12.112, y: 0.409, w: 0.368, h: 0.353, fontSize: 10, align: 'center', color: C.white });
}

/* ================================================================= slides */

// 1 - Cover: full-bleed red with the Alkes wordmark.
function slide01(s) {
  chrome(s, 1);
  photo(s, { x: 0, y: 0, w: 13.333, h: 7.5, size: 82 });
  rect(s, { x: 0, y: 0, w: 13.333, h: 7.5, fill: C.red, transparency: 10 });
  freeform(s, [['M', 0, 0], ['C', K, 0, 1, 1 - K, 1, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 3.146, w: 4.354, h: 4.354, fill: C.white, transparency: 85 });
  freeform(s, [['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['C', 1 - K, 1, 0, K, 0, 0], ['Z']],
    { x: 11.708, y: 0, w: 1.625, h: 1.625, fill: C.white, transparency: 85 });
  s.addText('Alkes', {
    x: 4.338, y: 3.118, w: 4.656, h: 1.279, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 70, bold: true, charSpacing: 3, color: C.white
  });
  pill(s, 'Medical Presentation', { x: 5.391, y: 4.397, w: 2.551, h: 0.468 });
  txt(s, 'Alkes is a multipurpose presentation design made by Neermana Studio lorem ipsum.', {
    x: 4.16, y: 5.979, w: 5.013, h: 0.673, align: 'center', color: C.white, lineSpacingMultiple: 1.5
  });
  cross(s, 10.716, 4.795, 0.34, C.white);
  cross(s, 2.278, 1.952, 0.34, C.white);
  logo(s, 6.274, 2.386, 0.786, C.white);
}

// 2 - Welcome / testimonial card over a dark photo panel.
function slide02(s) {
  chrome(s, 2);
  photo(s, { x: 5.556, y: 0, w: 7.778, h: 7.5, size: 51 });
  rect(s, { x: 1.063, y: 3.138, w: 6.858, h: 3.065, rectRadius: 0.3, fill: C.white, shadow: CARD_SHADOW() });
  heading(s, [R('Welcome To ', C.ink, true), R('Our Clinics', C.red)], { x: 1.063, y: 1.488, w: 3.814, h: 1.178 });
  body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.',
    { x: 3.773, y: 3.617, w: 3.523, h: 1.279 });
  txt(s, 'Luke Skywalker', { x: 3.773, y: 5.053, w: 1.9, h: 0.37, bold: true, lineSpacingMultiple: 1.5 });
  txt(s, 'Board Director of Hospital', { x: 3.773, y: 5.357, w: 2.341, h: 0.37, italic: true, color: C.red, lineSpacingMultiple: 1.5 });
  oval(s, { x: 5.219, y: 1.74, w: 0.674, fill: C.red });
  cross(s, 5.378, 1.899, 0.354, C.white);
  photo(s, { x: 1.688, y: 3.863, w: 1.616, h: 1.615, round: 'ellipse', size: 13 });
  s.addShape('line', { x: 1.429, y: 6.203, w: 2.25, h: 0, line: { color: C.peach, width: 2.25 } });
}

// 3 - Table of content: six icon + copy cells on a white card.
function slide03(s) {
  chrome(s, 3);
  freeform(s, [['M', 0, 0], ['L', 1, 0], ['L', 1, 0.9123], ['C', 1, 0.9607, 0.9697, 1, 0.9323, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 0, w: 8.875, h: 6.854, fill: C.red });
  freeform(s, [['M', 0, 0], ['C', K, 0, 1, 1 - K, 1, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 2.899, w: 3.955, h: 3.955, fill: C.white, transparency: 85 });
  heading(s, 'Table of Content', { x: 1.063, y: 1.13, w: 7.062, h: 0.64, color: C.white });
  rect(s, { x: 1.084, y: 2.104, w: 11.166, h: 4.25, rectRadius: 0.25, fill: C.white, shadow: CARD_SHADOW() });
  const items = [
    ['syringe', 'Welcome Message'], ['kit', 'About Us'], ['stethoscope', 'What We Do'],
    ['ribbon', 'Our Community'], ['dropper', 'Our History'], ['nurse', 'Working Hours']
  ];
  items.forEach((it, i) => {
    const x = 1.745 + (i % 3) * 3.535;
    const y = i < 3 ? 0 : 1.809;
    icon(s, it[0], { x: 1.849 + (i % 3) * 3.536, y: 2.596 + y, w: 0.385, h: 0.385, color: C.red });
    sub(s, it[1], { x: x, y: 3.12 + y, w: 2.773, h: 0.337 });
    body(s, LOREM, { x: x, y: 3.381 + y, w: 2.773, h: 0.673 });
  });
  dotGrid(s, { x: 10.427, y: 1.062, w: 0.901, h: 1.568, cols: 4, rows: 6 });
  // half-disc peeking out from the card's top edge
  freeform(s, [['M', 1, 0], ['L', 1, 1], ['C', 0, 0.94, 0, 0.06, 1, 0], ['Z']],
    { x: 7.014, y: 1.697, w: 0.336, h: 0.676, fill: C.peach });
}

// 4 - Angled dark photo panel + right-aligned quote.
function slide04(s) {
  chrome(s, 4);
  photo(s, {
    x: 0, y: 1.125, w: 7.759, h: 6.379, size: 54, cx: -0.03, cy: -0.11,
    spec: [['M', 0, 0], ['L', 0.7033, 0.1418], ['C', 0.7701, 0.1552, 0.8262, 0.2102, 0.8512, 0.2868],
      ['L', 0.9864, 0.7011], ['C', 1.0061, 0.7615, 1.0043, 0.8292, 0.9815, 0.8879],
      ['L', 0.9381, 1.0], ['L', 0, 1], ['Z']],
    cx: -1.13, cy: -0.7
  });
  freeform(s, CROSS, { x: 8.884, y: 1.875, w: 6.22, h: 6.22, lineColor: 'F2F2F2', lineWidth: 5.5 });
  heading(s, [R('Life is Choices, ', C.ink), R('and They Are Relentless', C.red)],
    { x: 7.759, y: 2.609, w: 4.423, h: 1.178, align: 'right' });
  body(s, 'What the great America poet is saying is that there is no way to avoid what must be done.',
    { x: 8.032, y: 3.869, w: 4.151, h: 0.673, align: 'right' });
  pill(s, 'Continue to next page', { x: 9.773, y: 4.761, w: 2.409, h: 0.449 });
  oval(s, { x: 6.417, y: 2.941, w: 0.611, fill: C.red });
  cross(s, 6.561, 3.086, 0.321, C.white);
}

// 5 - "Do Not Ever Sick": copy at left, two photo columns, floating stat card.
function slide05(s) {
  chrome(s, 5);
  photo(s, {
    x: 5.21, y: 2.875, w: 3.31, h: 4.625, size: 45, cy: 0.03,
    r: { tl: 0.34, tr: 0.34 }
  });
  photo(s, { x: 8.671, y: 0, w: 3.31, h: 6.413, size: 45, cx: 0.09, cy: 0.07, r: { bl: 0.34, br: 0.34 } });
  heading(s, [R('Do Not Ever Sick ', C.ink), R('In Mind or Heart', C.red)], { x: 1.063, y: 1.679, w: 3.814, h: 1.178 });
  body(s, LOREM_LONG, { x: 1.035, y: 2.981, w: 3.842, h: 1.582 });
  oval(s, { x: 1.063, y: 4.948, w: 0.947, fill: C.white, shadow: SOFT_SHADOW() });
  icon(s, 'kit', { x: 1.379, y: 5.267, w: 0.316, h: 0.296, color: C.red });
  sub(s, 'Get Help', { x: 2.351, y: 4.948, w: 2.531, h: 0.337 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer. ', { x: 2.351, y: 5.222, w: 2.531, h: 0.673 });
  rect(s, { x: 5.94, y: 1.5, w: 3.557, h: 1.777, rectRadius: 0.175, fill: C.white, shadow: CARD_SHADOW() });
  txt(s, '5,732', { x: 6.332, y: 1.833, w: 2.773, h: 0.572, fontSize: 28, bold: true, color: C.peach, align: 'center' });
  txt(s, 'Patients we had in this past years lorem ipsum',
    { x: 6.332, y: 2.271, w: 2.773, h: 0.674, align: 'center', lineSpacingMultiple: 1.5 });
  oval(s, { x: 6.523, y: 1.356, w: 0.287, fill: C.red });
}

// 6 - Two photo bands across the top, headline + copy below.
function slide06(s) {
  chrome(s, 6);
  photo(s, {
    x: 0, y: 0, w: 9.0, h: 3.75, size: 58,
    spec: [['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0.0636, 1], ['C', 0.0285, 1, 0, 0.9317, 0, 0.8474], ['Z']]
  });
  photo(s, {
    x: 9.175, y: 0, w: 4.159, h: 3.75, size: 54, cy: -0.08,
    spec: [['M', 0, 0], ['L', 1, 0], ['L', 1, 0.8474], ['C', 1, 0.9317, 0.9384, 1, 0.8624, 1], ['L', 0, 1], ['Z']]
  });
  freeform(s, [['M', 0, 0], ['L', 0.9183, 0], ['L', 0.9347, 0.0042],
    ['C', 0.972, 0.0236, 1, 0.1072, 1, 0.2074], ['L', 1, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 5.323, w: 5.523, h: 2.177, fill: C.wash });
  heading(s, [R('Declare The Past, Diagnose The Present, ', C.ink), R('Foretell The Future of Life', C.red)],
    { x: 1.063, y: 4.557, w: 6.334, h: 1.717 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed',
    { x: 7.987, y: 4.574, w: 4.283, h: 0.976 });
  pill(s, 'Learn more', { x: 8.092, y: 5.807, w: 1.534, h: 0.449 });
  oval(s, { x: 8.754, y: 1.536, w: 0.679, fill: C.red });
  icon(s, 'stethoscope', { x: 8.961, y: 1.742, w: 0.266, h: 0.266, color: C.white });
}

// 7 - Full-height photo with a red fade, two icon bullets at right.
function slide07(s) {
  chrome(s, 7);
  const panel = [['M', 0, 0], ['L', 0.9169, 0], ['C', 0.9628, 0, 1, 0.029, 1, 0.0648],
    ['L', 1, 0.9352], ['C', 1, 0.971, 0.9628, 1, 0.9169, 1], ['L', 0, 1], ['Z']];
  photo(s, { x: 0, y: 0, w: 5.846, h: 7.5, size: 45, spec: panel, captionColor: 'F0C0C0' });
  redFade(s, { x: 0, y: 0, w: 5.846, h: 7.5, r: { tr: 0.44, br: 0.44 } });
  oval(s, { x: 5.608, y: 1.787, w: 0.477, fill: C.red });
  s.addShape('line', { x: 5.846, y: 1.944, w: 0, h: 0.197, line: { color: C.white, width: 1 } });
  heading(s, [R('Flowers Make People ', C.ink), R('Better and Happier', C.red)], { x: 6.667, y: 1.518, w: 5.583, h: 1.178 });
  body(s, LOREM.trim() + ' Maecenas porttitor congue massa ', { x: 6.667, y: 2.696, w: 5.537, h: 0.674 });
  oval(s, { x: 6.667, y: 3.723, w: 0.947, fill: C.white, shadow: SOFT_SHADOW() });
  icon(s, 'ribbon', { x: 7.007, y: 4.048, w: 0.265, h: 0.303, color: C.red });
  sub(s, 'Another side of life', { x: 8.013, y: 3.758, w: 4.245, h: 0.337 });
  body(s, LOREM.trim() + ' Lorem ipsum dolor. ', { x: 8.014, y: 4.023, w: 4.189, h: 0.674 });
  oval(s, { x: 6.667, y: 5.078, w: 0.947, fill: C.red, shadow: DROP_SHADOW() });
  icon(s, 'heart', { x: 6.957, y: 5.398, w: 0.366, h: 0.32, color: C.white });
  sub(s, 'Long relationship', { x: 8.013, y: 5.085, w: 4.245, h: 0.337 });
  body(s, LOREM.trim() + '  Lorem ipsum dolor', { x: 8.014, y: 5.351, w: 4.189, h: 0.674 });
  txt(s, 'Description maunnes consecturer lorem ipsum dolor sit amet.',
    { x: 0.8, y: 4.653, w: 4.245, h: 0.769, align: 'center', fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  pill(s, 'Stay health', { x: 2.142, y: 5.577, w: 1.563, h: 0.449, shadow: DROP_SHADOW() });
}

// 8 - Workout: left copy, centre photo, right callout card and two stats.
function slide08(s) {
  chrome(s, 8);
  photo(s, { x: 4.762, y: 1.109, w: 3.809, h: 5.281, rectRadius: 0.29, size: 45 });
  heading(s, [R('Have a Workout Time, ', C.ink), R('Have a Healthy Life', C.red)], { x: 1.063, y: 1.492, w: 3.08, h: 2.255 });
  body(s, EVERY, { x: 1.063, y: 3.769, w: 3.08, h: 1.582 });
  pill(s, 'Learn more', { x: 1.153, y: 5.559, w: 1.534, h: 0.449 });
  rect(s, { x: 7.412, y: 1.467, w: 4.858, h: 1.777, rectRadius: 0.175, fill: C.white, shadow: CARD_SHADOW() });
  body(s, 'Every signs sixth midst place above multiply. Our signs moveth Of own grass his dry earth yielding were heaven doesn\u2019t.',
    { x: 7.779, y: 1.867, w: 4.123, h: 0.976 });
  cross(s, 10.967, 1.29, 0.354, C.red);
  [['10,287', 3.634], ['1,235', 5.185]].forEach(st => {
    txt(s, st[0], { x: 9.19, y: st[1], w: 2.773, h: 0.572, fontSize: 28, bold: true, color: C.peach });
    txt(s, 'Patients we had in this past years lorem ipsum',
      { x: 9.19, y: st[1] + 0.439, w: 3.08, h: 0.674, lineSpacingMultiple: 1.5 });
  });
}

// 9 - Dark panel with a circular photo cut-out.
function slide09(s) {
  chrome(s, 9);
  photo(s, {
    x: 0, y: 0, w: 3.728, h: 6.406, size: 45, cx: 0.02, cy: 0.34, captionColor: 'FFFFFF',
    spec: [['M', 0, 0], ['L', 1, 0], ['L', 1, 0.9342], ['C', 1, 0.9705, 0.9494, 1, 0.8869, 1], ['L', 0, 1], ['Z']]
  });
  photo(s, { x: 1.834, y: 1.885, w: 3.727, h: 3.729, round: 'ellipse', size: 45, shadow: CARD_SHADOW() });
  heading(s, [R('We Are Capable of Making Healthy and ', C.ink), R('Positive Decisions Today', C.red)],
    { x: 6.667, y: 1.83, w: 5.562, h: 1.717 });
  body(s, 'The quick, brown fox jumps over a lazy dog. DJs flock by when MTV ax quiz prog. Junk MTV quiz graced by fox whelps. Bawds jog, flick quartz, vex nymphs. Waltz, bad nymph, for quick jigs vex! Fox. The quick, brown fox jumps.',
    { x: 6.667, y: 3.624, w: 5.562, h: 1.279 });
  pill(s, 'Learn more', { x: 6.765, y: 5.221, w: 1.534, h: 0.449 });
  cross(s, 5.632, 1.653, 0.354, C.red);
}

// 10 - Three faded photo columns with numbered badges.
function slide10(s) {
  chrome(s, 10);
  heading(s, [R('Help People Recover is ', C.ink), R('Such A Honor', C.red)],
    { x: 2.381, y: 1.005, w: 8.571, h: 0.64, align: 'center' });
  txt(s, AND_IS, { x: 2.381, y: 1.661, w: 8.571, h: 0.674, align: 'center', lineSpacingMultiple: 1.5 });
  const cols = [
    { x: 1.116, cx: -0.23, r: { tl: 0.26 }, n: '1', t: 'Subjudul 1' },
    { x: 4.816, cx: 0, r: {}, n: '2', t: 'Subjudul 2' },
    { x: 8.517, cx: 0, r: { tr: 0.26 }, n: '3', t: 'Subjudul 3' }
  ];
  cols.forEach(c => {
    photo(s, { x: c.x, y: 3.75, w: 3.7, h: 3.75, r: c.r, size: 45, cx: c.cx, captionColor: 'F0C4C4' });
    redFade(s, { x: c.x, y: 3.75, w: 3.7, h: 3.75, r: c.r });
    txt(s, c.t, { x: c.x + 0.36, y: 2.961, w: 2.98, h: 0.337, fontSize: 14, bold: true, align: 'center' });
    oval(s, { x: c.x + 1.626, y: 3.526, w: 0.448, fill: C.red, shadow: DROP_SHADOW() });
    txt(s, c.n, { x: c.x + 1.584, y: 3.582, w: 0.533, h: 0.337, fontSize: 14, bold: true, color: C.white, align: 'center' });
    txt(s, LOREM, { x: c.x + 0.36, y: 6.109, w: 2.98, h: 0.674, align: 'center', color: C.white, lineSpacingMultiple: 1.5 });
  });
}

// 11 - Four service cards, first one highlighted in red.
function slide11(s) {
  chrome(s, 11);
  heading(s, [R('Live A Healthy Life ', C.ink), R('and Recover Quickly', C.red)],
    { x: 2.381, y: 1.005, w: 8.571, h: 0.64, align: 'center' });
  txt(s, AND_IS, { x: 2.381, y: 1.661, w: 8.571, h: 0.674, align: 'center', lineSpacingMultiple: 1.5 });
  const cards = [
    { x: 1.062, red: true, icon: 'nurse', title: 'Meet Good Doctor' },
    { x: 3.916, red: false, icon: 'kit', title: 'Lovely Caregiver\nIn Our Hospital' },
    { x: 6.734, red: false, icon: 'syringe', title: 'A Healthy Environment' },
    { x: 9.553, red: false, icon: 'stethoscope', title: '24/7 Work Ready For Your Needs' }
  ];
  cards.forEach(c => {
    const fg = c.red ? C.white : C.ink;
    corners(s, {
      x: c.x, y: 2.896, w: 2.686, h: 3.462, r: { br: 0.32 },
      fill: c.red ? C.red : C.white, shadow: CARD_SHADOW()
    });
    icon(s, c.icon, { x: c.x + 0.49, y: 3.55, w: 0.37, h: 0.385, color: c.red ? C.white : C.red });
    txt(s, c.title, { x: c.x + 0.365, y: 4.176, w: 2.021, h: 0.572, fontSize: 14, bold: true, color: fg });
    txt(s, EVERY_SHORT, { x: c.x + 0.365, y: 4.7, w: 2.021, h: 0.976, color: fg, lineSpacingMultiple: 1.5 });
    oval(s, { x: c.x + 2.164, y: 5.861, w: 0.276, lineColor: C.peach, lineWidth: 1 });
    chevron(s, { x: c.x + 2.27, y: 5.925, w: 0.052, h: 0.093, color: C.peach });
  });
}

// 12 - Journey timeline: two faded photo cards with year pills.
function journeyCard(s, x) {
  photo(s, { x: x, y: 1.509, w: 3.356, h: 4.429, rectRadius: 0.32, size: 45, cx: -0.03, cy: -0.25, captionColor: 'F0C4C4' });
  redFade(s, { x: x, y: 1.509, w: 3.357, h: 4.43, r: { tl: 0.32, tr: 0.32, bl: 0.32, br: 0.32 } });
  pill(s, '2005', { x: x + 1.136, y: 3.878, w: 1.021, h: 0.449, shadow: DROP_SHADOW() });
  txt(s, 'Another side of life', { x: x + 0.24, y: 4.461, w: 2.812, h: 0.337, fontSize: 14, bold: true, color: C.white, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit',
    { x: x + 0.24, y: 4.782, w: 2.812, h: 0.674, align: 'center', color: C.white, lineSpacingMultiple: 1.5 });
}

function slide12(s) {
  chrome(s, 12);
  heading(s, [R('Our Journey To ', C.ink), R('Save People\u2019s Life', C.red)], { x: 1.051, y: 2.246, w: 3.333, h: 1.717 });
  body(s, QUICK, { x: 1.051, y: 3.975, w: 3.333, h: 1.279 });
  journeyCard(s, 5.02);
  journeyCard(s, 9.093);
  navButton(s, { x: 8.597, y: 3.622 });
  navButton(s, { x: 12.689, y: 3.622 });
  freeform(s, [['M', 0.6393, 0.1471], ['C', 0.5962, 0.0448, 0.5155, 0, 0.4415, 0.0362],
    ['L', 0, 0.2559], ['L', 0, 1], ['L', 1, 1], ['Z']],
    { x: 0, y: 5.906, w: 2.52, h: 1.594, fill: C.red });
  freeform(s, [['M', 1, 0.3345], ['L', 0.6068, 0.0599], ['C', 0.5231, 0, 0.427, 0.0493, 0.3665, 0.1796],
    ['L', 0, 1], ['L', 1, 1], ['Z']],
    { x: 11.438, y: 6.571, w: 1.895, h: 0.964, fill: C.peach });
}

// 13 - Same timeline cards plus a white detail card, giant ALKES watermark.
function slide13(s) {
  chrome(s, 13);
  s.addText('ALKES', {
    x: 1.422, y: 4.546, w: 10.553, h: 3.803, align: 'center',
    fontFace: HEAD, fontSize: 220, bold: true, color: C.wash
  });
  journeyCard(s, 0.884);
  journeyCard(s, 5.02);
  navButton(s, { x: 4.48, y: 3.622 });
  rect(s, { x: 9.093, y: 1.509, w: 3.357, h: 4.43, rectRadius: 0.32, fill: C.white, shadow: CARD_SHADOW() });
  icon(s, 'kit', { x: 9.671, y: 2.191, w: 0.45, h: 0.385, color: C.red });
  txt(s, 'Lovely Caregiver\nIn Our Hospital', { x: 9.567, y: 2.816, w: 2.408, h: 0.572, fontSize: 14, bold: true });
  body(s, EVERY, { x: 9.567, y: 3.423, w: 2.408, h: 1.885 });
}

// 14 - Team: red banner over four portrait tiles.
function slide14(s) {
  chrome(s, 14);
  photo(s, { x: 1.95, y: 1.141, w: 9.433, h: 2.094, rectRadius: 0.18, size: 57, cy: -0.05 });
  rect(s, { x: 1.95, y: 1.141, w: 9.446, h: 2.094, rectRadius: 0.18, fill: C.red, transparency: 10 });
  heading(s, [R('Our Medical Doctor and Staff', C.white, true), R('Who Love to See a Healthy Patient', C.white)],
    { x: 2.683, y: 1.552, w: 7.968, h: 1.178, align: 'center' });
  pill(s, 'Best team in the world', { x: 5.333, y: 3.008, w: 2.667, h: 0.449, shadow: DROP_SHADOW() });
  const team = [
    { x: 1.399, name: 'Sabrina Taka, MBBS', role: 'Doctor Specialist' },
    { x: 4.34, name: 'Carl McManam, MD', role: 'Dentist' },
    { x: 7.285, name: 'John Amidala, DO', role: 'Doctor of Osteopathic' },
    { x: 10.226, name: 'Yeo-Sohan, BEMS', role: 'Eastern Medicine Sp' }
  ];
  team.forEach(m => {
    photo(s, { x: m.x, y: 3.974, w: 1.708, h: 1.708, rectRadius: 0.27, size: 15 });
    oval(s, { x: m.x - 0.23, y: 3.974, w: 0.46, fill: C.peach, shadow: TILT_SHADOW() });
    check(s, { x: m.x - 0.12, y: 4.09, w: 0.24, h: 0.23, color: C.white });
    txt(s, m.name, { x: m.x - 0.296, y: 5.816, w: 2.299, h: 0.337, fontSize: 14, bold: true, color: C.red, align: 'center' });
    txt(s, m.role, { x: m.x - 0.296, y: 6.083, w: 2.299, h: 0.371, italic: true, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

// 15 - Two doctor profile cards with star ratings and action pills.
function slide15(s) {
  chrome(s, 15);
  heading(s, [R('Here is a Profile of ', C.ink), R('The Best Doctor in Hospital', C.red)],
    { x: 0.853, y: 2.353, w: 2.495, h: 2.794 });
  body(s, EVERY, { x: 9.985, y: 2.807, w: 2.408, h: 1.885 });
  [{ x: 3.822, stars: 4 }, { x: 6.861, stars: 5 }].forEach((card, i) => {
    corners(s, { x: card.x, y: 1.492, w: 2.748, h: 4.518, r: { tl: 0.28, tr: 0.28, bl: 0.28, br: 0.28 }, fill: C.wash, shadow: CARD_SHADOW() });
    corners(s, { x: card.x, y: 2.889, w: 2.748, h: 3.119, r: { tl: 0.28, tr: 0.28, bl: 0.28, br: 0.28 }, fill: C.white });
    photo(s, { x: card.x + 0.595, y: 2.132, w: 1.51, h: 1.51, round: 'ellipse', size: 12 });
    for (let k = 0; k < 5; k++) {
      s.addShape('star5', {
        x: card.x + 0.605 + k * 0.318, y: 3.883, w: 0.264, h: 0.257,
        fill: { color: k < card.stars ? C.peach : C.gray }, line: NONE
      });
    }
    txt(s, i === 0 ? 'Julia Aspera, DO' : 'Andy Sakhi, MeD',
      { x: card.x + 0.313, y: 4.247, w: 2.119, h: 0.337, fontSize: 14, bold: true, color: C.red, align: 'center' });
    txt(s, 'Doctor Specialist', { x: card.x + 0.313, y: 4.514, w: 2.119, h: 0.371, italic: true, align: 'center', lineSpacingMultiple: 1.5 });
    pill(s, 'Call', { x: card.x + 0.428, y: 5.108, w: 0.854, h: 0.449, shadow: DROP_SHADOW() });
    pill(s, 'Book', { x: card.x + 1.462, y: 5.108, w: 0.854, h: 0.449, outline: true, shadow: DROP_SHADOW() });
  });
}

// 16 - Board director portrait plus three skill bars.
function slide16(s) {
  chrome(s, 16);
  s.addText('2020', { x: -0.178, y: 4.546, w: 7.745, h: 3.803, fontFace: HEAD, fontSize: 220, bold: true, color: C.wash });
  photo(s, { x: 1.318, y: 1.295, w: 4.545, h: 4.115, rectRadius: 0.36, size: 34, cy: -0.16 });
  rect(s, { x: 1.317, y: 4.894, w: 4.545, h: 1.206, rectRadius: 0.29, fill: C.red });
  txt(s, 'Dr. George Rohan, EdD', { x: 1.698, y: 5.178, w: 2.9, h: 0.337, fontSize: 14, bold: true, color: C.white });
  txt(s, 'Senior Doctor', { x: 1.698, y: 5.445, w: 2.623, h: 0.371, italic: true, color: C.white, lineSpacingMultiple: 1.5 });
  pill(s, 'Call', { x: 4.628, y: 5.256, w: 0.854, h: 0.449, shadow: DROP_SHADOW() });
  heading(s, [R('Meet Our Hospital ', C.ink), R('Board Director Leader', C.red)], { x: 6.777, y: 1.4, w: 5.273, h: 1.178 });
  body(s, LOREM_LONG, { x: 6.777, y: 2.625, w: 5.273, h: 0.976 });
  [['Management', 3.824, 4.315], ['Leadership', 4.651, 3.347], ['Teamwork', 5.477, 4.021]].forEach(bar => {
    txt(s, bar[0], { x: 6.777, y: bar[1], w: 2.042, h: 0.371, bold: true, lineSpacingMultiple: 1.5 });
    rect(s, { x: 6.878, y: bar[1] + 0.442, w: 4.622, h: 0.115, rectRadius: 0.057, fill: C.gray });
    rect(s, { x: 6.873, y: bar[1] + 0.442, w: bar[2], h: 0.115, rectRadius: 0.057, fill: C.peach });
  });
}

// 17 - Photo mosaic over a red lower band.
function slide17(s) {
  chrome(s, 17);
  freeform(s, [['M', 0.968, 0], ['L', 0.032, 0], ['C', 0.0143, 0, 0, 0.0506, 0, 0.1123],
    ['L', 0, 1], ['L', 1, 1], ['L', 1, 0.1123], ['C', 1, 0.0506, 0.9857, 0, 0.968, 0], ['Z']],
    { x: 0.002, y: 3.715, w: 13.333, h: 3.781, fill: C.red });
  freeform(s, [['M', 0, 0], ['C', K, 0, 1, 1 - K, 1, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 5.429, w: 2.071, h: 2.071, fill: C.white, transparency: 85 });
  const tiles = [
    [5.033, 1.123, 1.729, 1.731, 18], [6.959, 1.123, 3.658, 1.731, 23], [10.817, 1.123, 1.729, 1.731, 18],
    [5.033, 3.046, 1.729, 1.731, 19], [6.959, 3.046, 1.729, 3.654, 26], [8.885, 3.046, 3.658, 1.729, 26],
    [5.033, 4.969, 1.729, 1.731, 18], [8.888, 4.969, 1.727, 1.731, 18], [10.817, 4.969, 1.729, 1.731, 18]
  ];
  tiles.forEach(t => photo(s, { x: t[0], y: t[1], w: t[2], h: t[3], rectRadius: 0.22, size: t[4] }));
  heading(s, [R('Diagnose The Present, ', C.ink), R('Foretell The Future of Life', C.red)],
    { x: 0.987, y: 1.258, w: 3.333, h: 2.255 });
  body(s, QUICK, { x: 0.987, y: 4.009, w: 3.333, h: 1.279, color: C.white });
  pill(s, 'Read more', { x: 1.016, y: 5.669, w: 1.563, h: 0.449, shadow: DROP_SHADOW() });
  oval(s, { x: 4.595, y: 1.541, w: 0.876, fill: C.peach, shadow: TILT_SHADOW() });
  txt(s, '2020', { x: 4.595, y: 1.541, w: 0.876, h: 0.876, align: 'center', color: C.white, lineSpacingMultiple: 1.5 });
}

// 18 - Three photos, each with an overlapping "About Us" card.
function slide18(s) {
  chrome(s, 18);
  const cols = [0.847, 4.833, 8.819];
  cols.forEach((x, i) => {
    photo(s, { x: x, y: 1.25, w: 3.687, h: 2.667, rectRadius: 0.28, size: 32, cx: [0, -0.01, 0.04][i] });
    const cx = x + 0.299 - (i ? 0.001 : 0);
    rect(s, { x: cx, y: 3.121, w: 2.282, h: 2.282, rectRadius: 0.32, fill: i === 0 ? C.red : C.white, shadow: CARD_SHADOW() });
    icon(s, ['kit', 'syringe', 'stethoscope'][i], { x: cx + 0.383, y: 3.533, w: 0.4, h: 0.385, color: i === 0 ? C.white : C.red });
    txt(s, 'About Us', { x: cx + 0.278, y: 4.057, w: 1.726, h: 0.337, fontSize: 14, bold: true, color: i === 0 ? C.white : C.ink });
    body(s, 'Lorem hag ipsum dolor sit amet.', { x: cx + 0.278, y: 4.317, w: 1.726, h: 0.674, color: i === 0 ? C.white : C.ink });
  });
  body(s, EVERY + ' ' + LOREM.trim() + ' Maecenas porttitor congue massa. ',
    { x: 0.847, y: 5.705, w: 11.424, h: 0.674 });
}

// 19 - Gallery strip of four tall photos.
function slide19(s) {
  chrome(s, 19);
  [[-2.014, 0], [0.798, 0.07], [6.421, 0]].forEach(t =>
    photo(s, { x: t[0], y: 1.438, w: 2.694, h: 4.146, rectRadius: 0.34, size: 30, cx: t[1] }));
  photo(s, { x: 3.61, y: 1.041, w: 2.694, h: 4.146, rectRadius: 0.34, size: 30, cy: -0.11, shadow: DROP_SHADOW() });
  txt(s, 'We are here since 2005', { x: 3.609, y: 5.386, w: 2.696, h: 0.337, fontSize: 14, bold: true, color: C.red, align: 'center' });
  txt(s, 'Lorem hag ipsum dolor sit amet consectur.',
    { x: 3.609, y: 5.662, w: 2.696, h: 0.674, align: 'center', lineSpacingMultiple: 1.5 });
  heading(s, [R('This is our gallery since ', C.ink), R('we first build the hospital', C.red)],
    { x: 9.619, y: 1.463, w: 2.87, h: 2.255 });
  body(s, EVERY + ' ', { x: 9.619, y: 3.905, w: 2.87, h: 1.582 });
  dotGrid(s, { x: 11.181, y: 5.983, w: 1.568, h: 0.901, cols: 6, rows: 4 });
}

// 20 - Full-bleed photo collage with a floating statement card.
function slide20(s) {
  chrome(s, 20);
  const tiles = [
    [0.003, 0.271, 3.503, 3.339, 27], [3.79, 0.271, 5.806, 6.958, 45, -0.03],
    [9.873, 0.271, 1.75, 3.339, 26], [11.901, 0.271, 1.425, 3.339, 26],
    [0.003, 3.891, 1.477, 3.339, 26], [1.76, 3.891, 1.75, 3.339, 26],
    [9.873, 3.891, 3.46, 3.339, 27]
  ];
  tiles.forEach(t => photo(s, { x: t[0], y: t[1], w: t[2], h: t[3], rectRadius: 0.28, size: t[4], cx: t[5] }));
  rect(s, { x: 0.826, y: 2.609, w: 4.909, h: 2.282, rectRadius: 0.32, fill: C.white, shadow: shadow(0.35) });
  heading(s, [R('Glad to be able to ', C.ink), R('accompany your healthy journey', C.red)],
    { x: 1.274, y: 2.892, w: 4.014, h: 1.717 });
}

// 21 - Appointment calendar on the left, red schedule panel on the right.
const WEEK = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const CAL_ROWS = [
  ['27', '28', '29', '30', '31', '1', '2'],
  ['3', '4', '5', '6', '7', '8', '9'],
  ['10', '11', '12', '13', '14', '15', '16'],
  ['17', '18', '19', '20', '21', '22', '23'],
  ['24', '25', '26', '27', '28', '29', '30'],
  ['31', '1', '2', '3', '4', '5', '6']
];
const COL_X = [1.591, 2.39, 3.191, 3.993, 4.795, 5.563, 6.35];
const ROW_Y = [2.631, 3.211, 3.884, 4.545, 5.218, 5.879];

function slide21(s) {
  chrome(s, 21);
  rect(s, { x: 6.13, y: 0, w: 7.203, h: 7.5, fill: C.red });
  rect(s, { x: 0, y: 0, w: 0.565, h: 7.5, fill: C.wash });
  rect(s, { x: 0.93, y: 1.014, w: 6.523, h: 5.668, rectRadius: 0.37, fill: C.white, shadow: CARD_SHADOW() });
  txt(s, 'January 2021', { x: 2.789, y: 1.472, w: 2.775, h: 0.337, fontFace: HEAD, fontSize: 14, bold: true, color: C.red, align: 'center' });
  chevron(s, { x: 1.548, y: 1.562, w: 0.087, h: 0.155, color: C.ink, flipH: false });
  chevron(s, { x: 6.66, y: 1.562, w: 0.087, h: 0.155, color: C.ink });
  // weekday header strip
  rect(s, { x: 1.421, y: 2.008, w: 5.541, h: 0.396, rectRadius: 0.198, fill: C.gray });
  WEEK.forEach((d, i) => txt(s, d, { x: COL_X[i] - 0.06, y: 2.058, w: 0.52, h: 0.303, align: 'center' }));
  // highlighted ranges: Jan 11-12 and Jan 20-22
  rect(s, { x: 2.138, y: 3.836, w: 1.7, h: 0.396, rectRadius: 0.198, fill: C.gray });
  rect(s, { x: 3.825, y: 4.496, w: 2.289, h: 0.396, rectRadius: 0.198, fill: C.gray });
  oval(s, { x: 1.593, y: 5.83, w: 0.396, fill: C.gray });
  oval(s, { x: 4.78, y: 3.159, w: 0.396, fill: C.red, shadow: shadow(0.4) });
  CAL_ROWS.forEach((week, r) => week.forEach((d, c) => {
    const muted = (r === 0) || (r === 5 && c > 0);
    const today = (r === 1 && c === 4);
    txt(s, d, {
      x: COL_X[c] - 0.05, y: ROW_Y[r], w: 0.5, h: 0.303, align: 'center',
      color: today ? C.white : (muted ? C.gray : C.ink)
    });
  }));
  oval(s, { x: 2.199, y: 3.785, w: 0.149, fill: C.peach });
  oval(s, { x: 3.893, y: 4.442, w: 0.149, fill: C.peach });
  heading(s, [R('Make Appointment', C.white, true), R('Has Never Been Easy Like This', C.white)],
    { x: 8.088, y: 1.014, w: 4.315, h: 1.717 });
  rect(s, { x: 7.818, y: 2.934, w: 4.86, h: 1.312, rectRadius: 0.18, fill: C.peach, shadow: shadow(0.3) });
  const events = [
    ['January 7', 'th', 3.032], ['January 11-12', 'nd', 4.376], ['January 20-22', 'nd', 5.515]
  ];
  events.forEach(e => {
    s.addText([
      { text: e[0], options: { fontSize: 14, bold: true, color: C.white } },
      { text: e[1], options: { fontSize: 14, bold: true, color: C.white, superscript: true } }
    ], { x: 8.123, y: e[2], w: 2.2, h: 0.415, fontFace: BODY, valign: 'middle', lineSpacingMultiple: 1.5 });
    body(s, 'Every signs sixth midst place above multiply. Our signs moveth Of own.',
      { x: 8.095, y: e[2] + 0.33, w: 4.308, h: 0.674, color: C.white });
  });
  dotGrid(s, { x: -0.592, y: 2.966, w: 0.901, h: 1.568, cols: 4, rows: 6, color: C.red });
}

// 22 - Testimonial carousel: red centre card between two grey cards.
function slide22(s) {
  chrome(s, 22);
  heading(s, [R('Our Patients ', C.ink), R('Testimonial', C.red)], { x: 3.226, y: 1.07, w: 6.882, h: 0.64, align: 'center' });
  oval(s, { x: 6.245, y: 1.952, w: 0.141, lineColor: C.ink, lineWidth: 1 });
  oval(s, { x: 6.566, y: 1.921, w: 0.201, fill: C.ink });
  oval(s, { x: 6.948, y: 1.952, w: 0.141, lineColor: C.ink, lineWidth: 1 });
  const side = [
    { x: 1.667, cx: -0.11, title: 'Excellent service', who: 'John Doe, ', role: 'Programmer' },
    { x: 8.826, cx: 0, title: 'Amazing Experience', who: 'Richard, ', role: 'Businessman' }
  ];
  side.forEach(card => {
    photo(s, { x: card.x, y: 2.893, w: 2.841, h: 3.11, rectRadius: 0.31, size: 28, cx: card.cx });
    rect(s, { x: card.x, y: 2.893, w: 2.841, h: 3.11, rectRadius: 0.31, fill: C.wash, shadow: DROP_SHADOW() });
    for (let k = 0; k < 5; k++) {
      s.addShape('star5', {
        x: card.x + 0.382 + k * 0.244, y: 3.288, w: 0.202, h: 0.197,
        fill: { color: k < 4 ? C.peach : C.gray }, line: NONE
      });
    }
    txt(s, card.title, { x: card.x + 0.289, y: 3.538, w: 2.266, h: 0.337, fontFace: HEAD, fontSize: 14, bold: true });
    body(s, 'Every signs sixth midst place above multiply. Our signs moveth Of own grass his dry.',
      { x: card.x + 0.291, y: 3.929, w: 2.265, h: 1.279 });
    s.addText([
      { text: card.who, options: { bold: true, color: C.ink } },
      { text: card.role, options: { italic: true, color: C.peach } }
    ], { x: card.x + 0.289, y: 5.238, w: 2.265, h: 0.371, fontFace: BODY, fontSize: 12, valign: 'middle', lineSpacingMultiple: 1.5 });
  });
  photo(s, { x: 4.891, y: 2.504, w: 3.552, h: 3.889, rectRadius: 0.43, size: 45, cy: 0.05 });
  rect(s, { x: 4.891, y: 2.504, w: 3.551, h: 3.889, rectRadius: 0.43, fill: C.red, transparency: 10 });
  for (let k = 0; k < 5; k++) {
    s.addShape('star5', { x: 5.345 + k * 0.318, y: 3.019, w: 0.264, h: 0.257, fill: { color: C.peach }, line: NONE });
  }
  txt(s, 'Excellent service with lovely caregiver and doctor',
    { x: 5.23, y: 3.272, w: 2.872, h: 0.769, fontFace: HEAD, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  body(s, 'Every signs sixth midst place above multiply. Our signs moveth Of own grass his dry earth yielding were.',
    { x: 5.232, y: 4.106, w: 2.87, h: 1.279, color: C.white });
  s.addText([
    { text: 'Catherine Wilson, ', options: { bold: true, color: C.white } },
    { text: 'Designer', options: { italic: true, color: C.peach } }
  ], { x: 5.23, y: 5.483, w: 2.567, h: 0.303, fontFace: BODY, fontSize: 12, valign: 'middle' });
  navButton(s, { x: 1.026, y: 4.32, flipH: false });
  navButton(s, { x: 12.049, y: 4.32 });
  freeform(s, [['M', 0.1989, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0.0004, 1], ['L', 0, 0.9911],
    ['C', 0, 0.5122, 0.0682, 0.1126, 0.1588, 0.0201], ['Z']],
    { x: 11.034, y: 7.039, w: 2.3, h: 0.461, fill: C.red });
}

// 23 - Break slide.
function slide23(s) {
  chrome(s, 23);
  photo(s, { x: 0, y: 0, w: 13.333, h: 7.5, size: 83 });
  freeform(s, [['M', 0, 0], ['L', 1, 0], ['L', 1, 0.9121], ['C', 1, 0.9607, 0.9784, 1, 0.9518, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 0, w: 10.667, h: 5.856, fill: C.red, transparency: 10 });
  txt(s, 'Let\u2019s take', { x: 1.443, y: 1.413, w: 5.013, h: 0.415, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  s.addText('A BREAK', {
    x: 1.443, y: 1.649, w: 4.656, h: 1.279, fontFace: HEAD, fontSize: 70, bold: true,
    charSpacing: 3, color: C.white, valign: 'middle'
  });
  txt(s, 'This Presentation will continue in 13:00', { x: 1.443, y: 2.76, w: 5.013, h: 0.371, color: C.white, lineSpacingMultiple: 1.5 });
  cross(s, 6.286, 1.554, 0.34, C.white);
  pill(s, 'Continue', { x: 1.443, y: 4.377, w: 1.563, h: 0.449, shadow: DROP_SHADOW() });
  freeform(s, [['M', 1, 0], ['L', 1, 0.7723], ['C', 1, 0.898, 0.8985, 1, 0.7733, 1], ['L', 0, 1],
    ['L', 0.0049, 0.9023], ['C', 0.0561, 0.3955, 0.4821, 0, 1, 0], ['Z']],
    { x: 8.396, y: 3.612, w: 2.27, h: 2.259, fill: C.white, transparency: 85 });
  oval(s, { x: 10.361, y: 1.897, w: 0.611, fill: C.peach });
  cross(s, 10.506, 2.042, 0.321, C.white);
  dotGrid(s, { x: 4.888, y: 5.379, w: 1.568, h: 0.901, cols: 6, rows: 4 });
}

// 24 - Bar chart of daily patients + two stat cards.
function slide24(s) {
  chrome(s, 24);
  s.addChart('bar', [{
    name: 'Series 1',
    labels: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    values: [20, 10, 25, 20, 10, 15, 18]
  }], {
    x: 0.965, y: 1.108, w: 7.31, h: 5.285,
    barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.red], showTitle: true, title: 'Week 1 \u2013 December 2021',
    titleFontFace: BODY, titleFontSize: 14, titleBold: true, titleColor: C.ink,
    showLegend: false, showValue: false,
    catAxisLineShow: true, catAxisLineColor: C.gray,
    valAxisLineShow: false, valGridLine: { style: 'none' },
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12, catAxisLabelColor: C.ink,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12, valAxisLabelColor: C.ink,
    valAxisMaxVal: 30, valAxisMajorUnit: 5,
    chartArea: { fill: { color: C.white } }, plotArea: { fill: { color: C.white } }
  });
  heading(s, [R('Daily Patients ', C.ink), R('Register Data', C.red)], { x: 8.653, y: 1.28, w: 3.685, h: 1.178 });
  [['20 Patients ', 'Average new patients per day', 2.672], ['25 Recovery', 'Average recovery per day', 4.072]].forEach(st => {
    rect(s, { x: 8.653, y: st[2], w: 3.558, h: 1.251, rectRadius: 0.17, fill: C.white, shadow: CARD_SHADOW() });
    txt(s, st[0], { x: 8.947, y: st[2] + 0.19, w: 2.969, h: 0.572, fontSize: 28, bold: true, color: C.peach });
    body(s, st[1], { x: 8.947, y: st[2] + 0.632, w: 2.969, h: 0.371 });
  });
  body(s, 'Every signs sixth midst place above multiply. Our signs moveth.', { x: 8.653, y: 5.546, w: 3.685, h: 0.674 });
}

// 25 - Tablet mock-up with a peach stat tile and four bullet options.
function slide25(s) {
  chrome(s, 25);
  rect(s, { x: 0.646, y: 1.111, w: 5.568, h: 7.66, rectRadius: 0.5, fill: '3A3A3C' });
  photo(s, { x: 0.958, y: 1.341, w: 5.091, h: 7.201, rectRadius: 0.35, size: 54, cy: -0.09 });
  heading(s, [R('Help People Recover is ', C.ink), R('Such A Honor', C.red)], { x: 6.75, y: 1.99, w: 5.583, h: 1.178 });
  body(s, AND_IS + '. They\u2019re night good after image brought in evening she\u2019d from lesser them. ',
    { x: 6.75, y: 3.168, w: 5.583, h: 1.279 });
  [['Description Option 1', 7.125, 4.686, 6.868], ['Description Option 3', 9.915, 4.686, 9.658],
  ['Description Option 2', 7.125, 5.173, 6.868], ['Description Option 4', 9.915, 5.173, 9.658]].forEach(o => {
    cross(s, o[3], o[2] + 0.081, 0.175, C.red);
    sub(s, o[0], { x: o[1], y: o[2], w: 2.416, h: 0.337 });
  });
  rect(s, { x: 4.525, y: 4.024, w: 1.997, h: 1.997, rectRadius: 0.25, fill: C.peach });
  txt(s, '8,276', { x: 4.685, y: 4.44, w: 1.677, h: 0.572, fontSize: 28, bold: true, color: C.white, align: 'center' });
  txt(s, 'Patients recover\nfrom their ill', { x: 4.685, y: 4.961, w: 1.677, h: 0.674, align: 'center', color: C.white, lineSpacingMultiple: 1.5 });
}

// 26 - Mobile app feature list next to two phone mock-ups.
function slide26(s) {
  chrome(s, 26);
  rect(s, { x: 8.873, y: 5.042, w: 4.46, h: 2.458, fill: C.peach });
  // back phone (dark screen), front phone (red screen with logo)
  rect(s, { x: 9.793, y: 1.567, w: 2.387, h: 4.787, rectRadius: 0.36, fill: '1C1C1C' });
  photo(s, { x: 9.906, y: 1.66, w: 2.161, h: 4.601, rectRadius: 0.29, size: 31 });
  rect(s, { x: 8.327, y: 1.167, w: 2.387, h: 4.787, rectRadius: 0.36, fill: '1C1C1C' });
  rect(s, { x: 8.44, y: 1.26, w: 2.161, h: 4.599, rectRadius: 0.29, fill: C.red });
  logo(s, 9.352, 3.029, 0.336, C.white);
  s.addText('Alkes', { x: 8.699, y: 3.306, w: 1.642, h: 0.572, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 28, bold: true, charSpacing: 3, color: C.white });
  txt(s, 'Mobile App', { x: 8.699, y: 3.788, w: 1.642, h: 0.303, align: 'center', color: C.peach });
  heading(s, [R('We Glad To Introduce You To ', C.ink), R('Our Mobile App', C.red)], { x: 1.021, y: 1.447, w: 6.487, h: 1.178 });
  body(s, EVERY, { x: 1.021, y: 2.735, w: 6.725, h: 0.674 });
  const feats = [
    { x: 1.068, y: 3.845, icon: 'thumb', title: 'Easy to use' },
    { x: 4.647, y: 3.845, icon: 'calendar', title: 'Booking' },
    { x: 1.068, y: 5.268, icon: 'chat', title: 'Live Consultation' },
    { x: 4.647, y: 5.268, icon: 'phone', title: 'Emergency Call' }
  ];
  feats.forEach(f => {
    oval(s, { x: f.x, y: f.y, w: 0.947, fill: C.white, shadow: SOFT_SHADOW() });
    icon(s, f.icon, { x: f.x + 0.317, y: f.y + 0.318, w: 0.311, h: 0.311, color: C.red });
    sub(s, f.title, { x: f.x + 1.25, y: f.y + 0.158, w: 2.009, h: 0.337 });
    body(s, 'Lorem ipsum dolor.', { x: f.x + 1.251, y: f.y + 0.438, w: 2.009, h: 0.371 });
  });
}

// 27 - Circular four-step diagram.
function slide27(s) {
  chrome(s, 27);
  const quads = [
    { spec: [['M', 0.0847, 1], ['L', 0, 1], ['C', 0, 0.4477, 0.4477, 0, 1, 0], ['L', 1, 0.0847], ['C', 0.4956, 0.0847, 0.0847, 0.4956, 0.0847, 1], ['Z']], x: 4.252, y: 1.414, fill: C.red },
    { spec: [['M', 0, 1], ['L', 0.0847, 1], ['C', 0.0847, 0.4956, 0.4956, 0.0847, 1, 0.0847], ['L', 1, 0], ['C', 0.4477, 0, 0, 0.4477, 0, 1], ['Z']], x: 6.667, y: 1.414, fill: C.peach, flipH: true },
    { spec: [['M', 0, 1], ['L', 0.0847, 1], ['C', 0.0847, 0.4956, 0.4956, 0.0847, 1, 0.0847], ['L', 1, 0], ['C', 0.4477, 0, 0, 0.4477, 0, 1], ['Z']], x: 6.667, y: 3.829, fill: C.orange, flipH: true, flipV: true },
    { spec: [['M', 0.0847, 1], ['L', 0, 1], ['C', 0, 0.4477, 0.4477, 0, 1, 0], ['L', 1, 0.0847], ['C', 0.4956, 0.0847, 0.0847, 0.4956, 0.0847, 1], ['Z']], x: 4.252, y: 3.829, fill: C.teal, flipV: true }
  ];
  quads.forEach(q => freeform(s, q.spec, { x: q.x, y: q.y, w: 2.415, h: 2.415, fill: q.fill, flipH: q.flipH, flipV: q.flipV }));
  oval(s, { x: 4.8, y: 1.963, w: 3.733, fill: C.white, shadow: CARD_SHADOW() });
  heading(s, [R('A Healthy', C.ink, true), R('Life With Us', C.red)], { x: 5.205, y: 2.777, w: 2.924, h: 1.178, align: 'center' });
  txt(s, 'Every signs sixth midst place above multiply. Our signs moveth Of own.',
    { x: 5.354, y: 3.905, w: 2.62, h: 0.976, align: 'center', lineSpacingMultiple: 1.5 });
  const steps = [
    { badge: [4.703, 1.866], color: C.red, elbow: [4.213, 1.413, true, false, C.blue], icon: ['syringe', 3.579, 1.241, C.red], title: 'Step 1: Lorem Ipsum', tx: 1.737, ty: 1.533, bx: 0.866, align: 'right' },
    { badge: [7.974, 1.866], color: C.peach, elbow: [8.29, 1.413, false, false, C.peach], icon: ['stethoscope', 9.346, 1.245, C.peach], title: 'Step 2: Lorem Ipsum', tx: 9.264, ty: 1.533, bx: 9.236, align: 'left' },
    { badge: [4.703, 5.138], color: C.teal, elbow: [4.213, 5.395, true, true, C.teal], icon: ['kit', 3.56, 4.805, C.teal], title: 'Step 3: Lorem Ipsum', tx: 1.697, ty: 5.093, bx: 0.866, align: 'right' },
    { badge: [7.974, 5.138], color: C.orange, elbow: [8.29, 5.447, false, true, C.orange], icon: ['nurse', 9.346, 4.805, C.orange], title: 'Step 4: Lorem Ipsum', tx: 9.264, ty: 5.093, bx: 9.236, align: 'left' }
  ];
  steps.forEach(st => {
    const e = st.elbow;
    freeform(s, [['M', 0, 1], ['L', 0.6324, 0], ['L', 1, 0]],
      { x: e[0], y: e[1], w: 0.819, h: e[3] ? 0.879 : 0.774, lineColor: e[4], lineWidth: 2.25, flipH: e[2], flipV: e[3] });
    oval(s, { x: st.badge[0], y: st.badge[1], w: 0.653, fill: st.color });
    cross(s, st.badge[0] + 0.155, st.badge[1] + 0.155, 0.343, C.white);
    icon(s, st.icon[0], { x: st.icon[1], y: st.icon[2], w: 0.289, h: 0.289, color: st.icon[3] });
    txt(s, st.title, { x: st.tx, y: st.ty, w: 2.29, h: 0.415, fontSize: 14, bold: true, align: st.align, lineSpacingMultiple: 1.5 });
    txt(s, AND_IS_SHORT, { x: st.bx, y: st.ty + 0.345, w: 3.145, h: 0.976, align: st.align, lineSpacingMultiple: 1.5 });
  });
}

// 28 - Five-step snake process infographic.
function slide28(s) {
  chrome(s, 28);
  heading(s, [R('Process', C.ink), R(' Infographic', C.red)], { x: 3.226, y: 1.07, w: 6.882, h: 0.64, align: 'center' });
  // Wave: vertical risers joined by alternating top / bottom half-loops.
  const TOP = 0.341, BOT = 0.659, RX = 0.1;
  const snake = [['M', 0, 1], ['L', 0, TOP]];
  for (let i = 0; i < 5; i++) {
    const cx = i * 0.2 + RX, up = i % 2 === 0;
    const near = up ? TOP : BOT, far = up ? 0 : 1;
    snake.push(['C', cx - RX, near - K * (near - far), cx - K * RX, far, cx, far]);
    snake.push(['C', cx + K * RX, far, cx + RX, near - K * (near - far), cx + RX, near]);
    if (i < 4) snake.push(['L', cx + RX, up ? BOT : TOP]);
  }
  snake.push(['L', 1, 1]);
  freeform(s, snake, { x: 1.12, y: 2.64, w: 11.094, h: 3.257, lineColor: C.ink, lineWidth: 5 });
  const steps = [
    { x: 1.316, y: 2.902, icon: 'syringe', label: 'Step 1', ty: 4.888, badge: [2.066, 2.476] },
    { x: 3.534, y: 3.817, icon: 'ribbon', label: 'Step 2', ty: 2.307, badge: [4.285, 5.734] },
    { x: 5.753, y: 2.902, icon: 'kit', label: 'Step 3', ty: 4.888, badge: [6.504, 2.476] },
    { x: 7.972, y: 3.817, icon: 'stethoscope', label: 'Step 4', ty: 2.307, badge: [8.723, 5.734] },
    { x: 10.191, y: 2.902, icon: 'nurse', label: 'Step 5', ty: 4.888, badge: [10.942, 2.476] }
  ];
  steps.forEach(st => {
    oval(s, { x: st.x, y: st.y, w: 1.826, fill: C.white, shadow: CARD_SHADOW() });
    oval(s, { x: st.x + 0.22, y: st.y + 0.211, w: 1.373, h: 1.404, fill: C.white, shadow: CARD_SHADOW() });
    icon(s, st.icon, { x: st.x + 0.635, y: st.y + 0.6, w: 0.556, h: 0.556, color: C.red });
    navButton(s, { x: st.badge[0], y: st.badge[1], d: 0.325 });
    txt(s, st.label, { x: st.x, y: st.ty, w: 1.826, h: 0.415, fontSize: 14, bold: true, align: 'center', lineSpacingMultiple: 1.5 });
    txt(s, 'And is, face, give made fruitful over creeping.',
      { x: st.x, y: st.ty + 0.372, w: 1.825, h: 0.976, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

// 29 - Stacked "journey to success" cake with a numbered legend.
function slide29(s) {
  chrome(s, 29);
  heading(s, [R('A Journey ', C.ink), R('To Success', C.red)], { x: 1.163, y: 0.997, w: 6.487, h: 0.64 });
  const tiers = [
    { x: 7.584, y: 4.808, w: 4.587, h: 1.112, side: 1.173, color: C.red, dark: 'A50000' },
    { x: 7.965, y: 4.214, w: 3.824, h: 0.928, side: 0.979, color: C.peach, dark: 'BF9773' },
    { x: 8.392, y: 3.721, w: 2.97, h: 0.72, side: 0.761, color: C.orange, dark: 'BC7028' },
    { x: 8.757, y: 3.329, w: 2.241, h: 0.543, side: 0.572, color: C.teal, dark: '113230' },
    { x: 9.172, y: 3.07, w: 1.41, h: 0.342, side: 0.362, color: C.yellow, dark: 'AAA600' }
  ];
  tiers.forEach(t => {
    freeform(s, [['M', 0, 0], ['C', 0, 0.525, 0, 0.525, 0, 0.525], ['C', 0, 0.787, 0.224, 1, 0.5, 1],
      ['C', 0.776, 1, 1, 0.787, 1, 0.525], ['C', 1, 0, 1, 0, 1, 0], ['L', 0, 0], ['Z']],
      { x: t.x, y: t.y + t.h * 0.5, w: t.w, h: t.side, fill: t.dark });
    oval(s, { x: t.x, y: t.y, w: t.w, h: t.h, fill: t.color });
  });
  s.addShape('line', { x: 9.878, y: 1.192, w: 0, h: 2.049, line: { color: C.ink, width: 3 } });
  freeform(s, [['M', 1, 0.8296], ['C', 0.6653, 0.6571, 0.3327, 1, 0, 0.8296], ['L', 0, 0.1704],
    ['C', 0.3327, 0.3429, 0.6653, 0, 1, 0.1704], ['C', 0.9073, 0.354, 0.9073, 0.354, 0.8609, 0.4558],
    ['C', 0.9073, 0.573, 0.9536, 0.6947, 1, 0.8296], ['Z']],
    { x: 9.897, y: 1.151, w: 0.89, h: 0.811, fill: C.ink });
  const legend = [
    { y: 1.792, color: C.ink, num: '06' }, { y: 2.611, color: C.yellow, num: '05' },
    { y: 3.431, color: C.teal, num: '04' }, { y: 4.251, color: C.orange, num: '03' },
    { y: 5.071, color: C.peach, num: '02' }, { y: 5.891, color: C.red, num: '01' }
  ];
  legend.forEach(l => {
    rect(s, { x: 1.261, y: l.y + 0.067, w: 0.112, h: 0.607, rectRadius: 0.056, fill: l.color });
    txt(s, 'And is, face, give made fruitful over creeping moving bearing fill cattle fifth saw.',
      { x: 1.5, y: l.y, w: 4.576, h: 0.674, lineSpacingMultiple: 1.5 });
    txt(s, l.num, { x: 6.126, y: l.y + 0.057, w: 0.853, h: 0.572, fontSize: 28, bold: true, color: l.color, align: 'center' });
  });
}

// 30 - Thank-you slide with contact details and a phone mock-up.
function slide30(s) {
  chrome(s, 30);
  freeform(s, [['M', 0, 0], ['L', 0.9311, 0], ['C', 0.9691, 0, 1, 0.0274, 1, 0.0612],
    ['L', 1, 0.9388], ['C', 1, 0.9726, 0.9691, 1, 0.9311, 1], ['L', 0, 1], ['Z']],
    { x: 0, y: 0, w: 6.667, h: 7.5, fill: C.red });
  rect(s, { x: 5.473, y: 1.357, w: 2.387, h: 4.787, rectRadius: 0.36, fill: '1C1C1C' });
  photo(s, { x: 5.562, y: 1.443, w: 2.208, h: 4.573, rectRadius: 0.29, size: 36, cy: 0.02 });
  s.addText([
    R('Thank you', C.white, true), R('for watching', C.white, true), R('to this presentation', C.peach)
  ], { x: 0.878, y: 1.099, w: 4.059, h: 2.524, fontFace: HEAD, fontSize: 36, bold: true, charSpacing: 3, valign: 'middle' });
  txt(s, 'If you need anything to explain, please don\u2019t hesitate to ask, I\u2019ll be happy to answer!',
    { x: 0.878, y: 5.364, w: 4.059, h: 0.674, color: C.white, lineSpacingMultiple: 1.5 });
  const contacts = [
    { icon: 'pin', iy: 1.508, iw: 0.232, ih: 0.31, ty: 1.839, text: '575 Market St, Ste 150, San Francisco, CA 94105' },
    { icon: 'phone', iy: 2.851, iw: 0.31, ih: 0.31, ty: 3.182, text: '(+1) 23456789\n(+1) 48359849' },
    { icon: 'mail', iy: 4.226, iw: 0.328, ih: 0.246, ty: 4.526, text: 'contact@yourmail.com\ndoctor.book@yourmail.com' }
  ];
  contacts.forEach(c => {
    icon(s, c.icon, { x: 8.691, y: c.iy, w: c.iw, h: c.ih, color: C.red });
    txt(s, c.text, { x: 8.607, y: c.ty, w: 3.122, h: 0.674, lineSpacingMultiple: 1.5 });
  });
  pill(s, 'www.ourhospital.com', { x: 8.691, y: 5.543, w: 2.409, h: 0.449 });
  // map pin marker sitting on the phone screen
  s.addShape('teardrop', { x: 5.814, y: 2.847, w: 0.563, h: 0.563, rotate: 225, fill: { color: C.red }, line: NONE });
  icon(s, 'pin', { x: 5.99, y: 3.008, w: 0.212, h: 0.283, color: C.white });
  dotGrid(s, { x: 12.56, y: 2.966, w: 0.901, h: 1.568, cols: 4, rows: 6 });
}

/* ------------------------------------------------------------------- main */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'ALKES16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'ALKES16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.author = 'Neermana Studio';
  pptx.title = 'Alkes - Medical Presentation';
  BUILDERS.forEach(fn => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    fn(slide);
  });
  return pptx.writeFile({ fileName: path.join(__dirname, '185b3ebd-f9e8-4e2b-b6e8-f5572579b819_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
