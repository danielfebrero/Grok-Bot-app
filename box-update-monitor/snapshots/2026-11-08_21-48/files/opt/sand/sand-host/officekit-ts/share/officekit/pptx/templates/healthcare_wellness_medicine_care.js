/**
 * Healthcare Deck — recreated with pptxgenjs.
 *
 * Rebuild of "0e2f509a-6ae5-43b2-996a-f6267e04494a.pptx" (12 slides, 13.333in x 7.5in).
 * Everything below is plain pptxgenjs: positions/sizes are inches, colors are hex.
 *
 * Note on artwork: the deck's only bitmap/SVG assets are its small icons. None are
 * embedded here — each is redrawn from native pptxgenjs shapes by `icon()`. The
 * deck's picture placeholders are empty in the source (no photo was ever dropped
 * in), so they render as nothing and are likewise not recreated.
 */

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  green: '96A78D', // accent1 — cards, panels
  mint: 'B6CEB4', // accent2 — pills, leaf corners, icons
  white: 'FFFFFF',
  ink: '262626', // headings on light cards
  body: '808080', // body copy
  grey: 'F0F0F0', // slide-2 background
};

const FONT_HEAD = 'Inter'; // theme major font
const FONT_BODY = 'Plus Jakarta Sans'; // theme minor font

// The two drop shadows used throughout the source deck. These are functions
// because pptxgenjs rewrites the shadow object in place while writing XML,
// so every shape needs its own copy.
const SHADOW = () => ({ type: 'outer', color: '000000', opacity: 0.1, blur: 15, offset: 3, angle: 90 });
const SHADOW_SOFT = () => ({ type: 'outer', color: '000000', opacity: 0.08, blur: 10, offset: 3, angle: 90 });

/* ---------------------------------------------------------------- helpers */

// PowerPoint stores rounded-corner radius as a fraction of the shape's short
// side; pptxgenjs wants it in inches.
const radius = (adj, w, h) => adj * Math.min(w, h);

/** Rounded card / panel. */
function card(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rectRadius: radius(o.adj, o.w, o.h),
    fill: { color: o.fill },
    shadow: o.shadow ? o.shadow() : undefined,
  });
}

/**
 * The decorative "leaf" that sits in the bottom-right corner of every card:
 * a quarter-round notch on the top-right, a concave sweep back to the corner.
 */
function leaf(slide, x, y, w, h, color) {
  slide.addShape('custGeom', {
    x, y, w, h,
    fill: { color },
    points: [
      { x: w, y: 0 },
      { x: w, y: h * 0.61 },
      { curve: { type: 'cubic', x1: w, y1: h * 0.825, x2: w * 0.808, y2: h } , x: w * 0.572, y: h },
      { x: 0, y: h },
      { curve: { type: 'cubic', x1: w * 0.098, y1: h * 0.564, x2: w * 0.426, y2: h * 0.207 }, x: w * 0.862, y: h * 0.039 },
      { close: true },
    ],
  });
}

/** Plain text box: top-anchored, no fill — the source deck's default. */
function textBox(slide, text, o) {
  slide.addText(text, Object.assign({
    valign: 'top',
    fontFace: FONT_BODY,
    fontSize: 12,
    color: C.body,
    align: 'left',
  }, o));
}

/** Section / card heading (14pt, theme major font). */
function heading(slide, text, o) {
  textBox(slide, text, Object.assign({ h: 0.337, fontFace: FONT_HEAD, fontSize: 14, color: C.ink }, o));
}

/** Big slide title: theme major font, black, single-spaced. */
function title(slide, text, o) {
  textBox(slide, text, Object.assign({ fontFace: FONT_HEAD, color: '000000' }, o));
}

/** Body copy: 12pt, grey, 150% line spacing. */
function body(slide, text, o) {
  textBox(slide, text, Object.assign({ lineSpacingMultiple: 1.5 }, o));
}

/** Fully rounded pill button with a centred label. */
function pill(slide, text, o) {
  slide.addText(text, {
    shape: 'roundRect', x: o.x, y: o.y, w: o.w, h: o.h,
    rectRadius: radius(0.5, o.w, o.h),
    fill: { color: o.fill }, shadow: o.shadow ? o.shadow() : undefined,
    align: 'center', valign: 'middle',
    fontFace: FONT_HEAD, fontSize: 12, color: o.color,
  });
}

/* ------------------------------------------------------------------ icons */

/*
 * Icon glyphs, redrawn from native shapes. Each part is
 *   [shapeType, x, y, w, h, mode, extraOptions?]
 * with x/y/w/h in 0..1 units of the icon's box, and mode:
 *   'fill' solid   'line' stroked outline   'hole' knocked out of a solid glyph
 *   'tick' white stroke on top of a solid glyph
 * 'poly' and 'curve' parts carry their points (also 0..1) in extraOptions.pts;
 * a 'roundRect' may override its corner radius with extraOptions.r.
 */
const ICONS = {
  people: [ // a front figure with a second one peeking out behind it
    ['ellipse', 0.53, 0.20, 0.22, 0.30, 'fill'],
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.66, 0.54], [0.78, 0.60], [0.88, 0.78], [0.88, 0.80], [0.66, 0.80]] }],
    ['ellipse', 0.24, 0.20, 0.30, 0.30, 'fill'],
    ['ellipse', 0.31, 0.27, 0.16, 0.16, 'hole'],
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.11, 0.80], [0.11, 0.62], [0.28, 0.54], [0.51, 0.54], [0.67, 0.62], [0.67, 0.80]] }],
    ['poly', 0, 0, 1, 1, 'hole', { pts: [[0.19, 0.72], [0.30, 0.63], [0.49, 0.63], [0.60, 0.72]] }],
  ],
  hand: [ // hand pointing upwards
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.10, 0.50], [0.10, 0.30], [0.24, 0.11], [0.46, 0.06], [0.68, 0.15], [0.77, 0.34], [0.77, 0.50]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.31, 0.86], [0.31, 0.33], [0.38, 0.24], [0.49, 0.23], [0.57, 0.30], [0.58, 0.42], [0.58, 0.60]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.58, 0.60], [0.84, 0.66], [0.92, 0.77], [0.92, 0.97]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.31, 0.68], [0.17, 0.68], [0.07, 0.78], [0.09, 0.91], [0.19, 0.97]] }],
  ],
  plane: [ // paper plane outline with a fold crease
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.97, 0.03], [0.06, 0.30], [0.24, 0.55], [0.26, 0.78], [0.52, 0.80], [0.72, 0.97], [0.97, 0.03]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.97, 0.03], [0.26, 0.78]] }],
  ],
  mic: [ // capsule microphone cradled in a U-shaped stand
    ['ellipse', 0.24, 0.02, 0.52, 0.62, 'line'],
    ['curve', 0, 0, 1, 1, 'line', { pts: [[0.02, 0.52], [0.02, 1.02], [0.98, 1.02], [0.98, 0.52]] }],
    ['rect', 0.36, 0.20, 0.28, 0.06, 'fill'],
    ['rect', 0.36, 0.38, 0.28, 0.06, 'fill'],
  ],
  editCheck: [ // open circle with a pencil across it
    ['ellipse', 0.02, 0.10, 0.78, 0.82, 'line'],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.52, 0.46], [0.90, 0.06], [1.00, 0.16], [0.62, 0.56], [0.50, 0.58], [0.52, 0.46]] }],
  ],
  handshake: [ // two clasped hands
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.00, 0.22], [0.20, 0.20], [0.38, 0.10], [0.52, 0.26], [0.66, 0.10], [0.84, 0.20], [1.00, 0.22]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.00, 0.62], [0.18, 0.60], [0.46, 0.90], [0.74, 0.60], [1.00, 0.62]] }],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.34, 0.46], [0.56, 0.26], [0.82, 0.62]] }],
  ],
  person: [ // figure standing inside a ring
    ['ellipse', 0.36, 0.02, 0.28, 0.28, 'fill'],
    ['roundRect', 0.28, 0.32, 0.44, 0.32, 'fill', { r: 0.30 }],
    ['rect', 0.37, 0.41, 0.26, 0.25, 'hole'],
    ['rect', 0.39, 0.56, 0.09, 0.32, 'fill'],
    ['rect', 0.52, 0.56, 0.09, 0.32, 'fill'],
    ['ellipse', 0.04, 0.62, 0.92, 0.34, 'line'],
  ],
  check: [ // solid disc with a tick
    ['ellipse', 0, 0, 1, 1, 'fill'],
    ['poly', 0, 0, 1, 1, 'tick', { pts: [[0.26, 0.52], [0.44, 0.70], [0.76, 0.34]] }],
  ],
  calculator: [
    ['roundRect', 0.08, 0.02, 0.84, 0.96, 'line'],
    ['rect', 0.22, 0.14, 0.56, 0.20, 'fill'],
    ['rect', 0.22, 0.46, 0.14, 0.12, 'fill'],
    ['rect', 0.43, 0.46, 0.14, 0.12, 'fill'],
    ['rect', 0.64, 0.46, 0.14, 0.12, 'fill'],
    ['rect', 0.22, 0.68, 0.14, 0.12, 'fill'],
    ['rect', 0.43, 0.68, 0.14, 0.12, 'fill'],
    ['rect', 0.64, 0.68, 0.14, 0.12, 'fill'],
  ],
  mail: [ // open envelope with a letter inside
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.02, 0.40], [0.50, 0.06], [0.98, 0.40], [0.98, 0.96], [0.02, 0.96]] }],
    ['rect', 0.20, 0.22, 0.60, 0.44, 'hole'],
    ['rect', 0.30, 0.32, 0.40, 0.09, 'fill'],
    ['rect', 0.30, 0.48, 0.40, 0.09, 'fill'],
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.02, 0.40], [0.50, 0.74], [0.98, 0.40], [0.98, 0.96], [0.02, 0.96]] }],
  ],
  docs: [ // stacked documents
    ['roundRect', 0.00, 0.00, 0.72, 0.80, 'line'],
    ['roundRect', 0.26, 0.20, 0.74, 0.80, 'line'],
    ['rect', 0.40, 0.38, 0.44, 0.07, 'fill'],
    ['rect', 0.40, 0.55, 0.44, 0.07, 'fill'],
    ['rect', 0.40, 0.72, 0.30, 0.07, 'fill'],
  ],
  user: [ // avatar knocked out of a solid disc (hole shows the surface behind)
    ['ellipse', 0.00, 0.00, 1.00, 1.00, 'fill'],
    ['ellipse', 0.11, 0.11, 0.78, 0.78, 'hole'],
    ['ellipse', 0.36, 0.22, 0.28, 0.28, 'fill'],
    ['ellipse', 0.22, 0.58, 0.56, 0.44, 'fill'],
  ],
  head: [ // profile silhouette with two cogs
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.55, 0.02], [0.82, 0.10], [0.97, 0.32], [0.95, 0.56], [0.78, 0.68], [0.80, 0.86], [0.90, 0.97], [0.32, 0.97], [0.34, 0.82], [0.18, 0.78], [0.14, 0.62], [0.02, 0.56], [0.16, 0.36], [0.24, 0.14]] }],
    ['gear6', 0.38, 0.13, 0.36, 0.30, 'hole'],
    ['ellipse', 0.47, 0.21, 0.18, 0.15, 'fill'],
    ['gear6', 0.52, 0.44, 0.26, 0.22, 'hole'],
    ['ellipse', 0.58, 0.50, 0.13, 0.11, 'fill'],
  ],
  badge: [ // ID card in a holder
    ['rect', 0.26, 0.06, 0.48, 0.07, 'fill'],
    ['rect', 0.22, 0.17, 0.56, 0.07, 'fill'],
    ['roundRect', 0.12, 0.27, 0.76, 0.36, 'line'],
    ['ellipse', 0.32, 0.38, 0.16, 0.20, 'fill'],
    ['rect', 0.52, 0.40, 0.26, 0.07, 'fill'],
    ['rect', 0.52, 0.52, 0.26, 0.07, 'fill'],
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.10, 0.65], [0.90, 0.65], [0.94, 0.92], [0.06, 0.92]] }],
  ],
  bars: [ // bar chart with a trend line
    ['rect', 0.00, 0.92, 0.90, 0.05, 'fill'],
    ['rect', 0.00, 0.66, 0.14, 0.24, 'fill'],
    ['rect', 0.19, 0.56, 0.14, 0.34, 'fill'],
    ['rect', 0.38, 0.62, 0.14, 0.28, 'fill'],
    ['rect', 0.57, 0.42, 0.14, 0.48, 'fill'],
    ['rect', 0.76, 0.26, 0.14, 0.64, 'fill'],
    ['poly', 0, 0, 1, 1, 'line', { pts: [[0.04, 0.46], [0.38, 0.18], [0.56, 0.40], [1.00, 0.02]] }],
  ],
  ambulance: [ // cab on the left, box body with a medical cross on the right
    ['poly', 0, 0, 1, 1, 'fill', { pts: [[0.00, 0.52], [0.16, 0.32], [0.40, 0.32], [0.40, 0.72], [0.00, 0.72]] }],
    ['poly', 0, 0, 1, 1, 'hole', { pts: [[0.06, 0.50], [0.19, 0.38], [0.29, 0.38], [0.29, 0.50]] }],
    ['roundRect', 0.38, 0.22, 0.62, 0.50, 'fill', { r: 0.10 }],
    ['rect', 0.62, 0.30, 0.14, 0.36, 'hole'],
    ['rect', 0.51, 0.42, 0.36, 0.12, 'hole'],
    ['ellipse', 0.14, 0.68, 0.20, 0.24, 'fill'],
    ['ellipse', 0.62, 0.68, 0.20, 0.24, 'fill'],
  ],
};

/**
 * Draw one icon glyph inside the box (x, y, w, h) using `color`.
 * `holeColor` is the surface the glyph sits on — cut-outs are painted with it.
 */
function icon(slide, kind, x, y, w, h, color, holeColor) {
  const stroke = Math.max(0.75, Math.min(w, h) * 3.2); // points
  ICONS[kind].forEach(([shape, nx, ny, nw, nh, mode, extra]) => {
    const o = { x: x + nx * w, y: y + ny * h, w: nw * w, h: nh * h };
    if (mode === 'fill') o.fill = { color };
    else if (mode === 'hole') o.fill = { color: holeColor || C.white };
    else if (mode === 'tick') o.line = { color: C.white, width: stroke };
    else o.line = { color, width: stroke };
    if (shape === 'curve') { // cubic stroke; part spans the whole icon box
      const [[sx, sy], [c1x, c1y], [c2x, c2y], [ex, ey]] = extra.pts;
      Object.assign(o, { x, y, w, h, points: [
        { x: sx * w, y: sy * h },
        { curve: { type: 'cubic', x1: c1x * w, y1: c1y * h, x2: c2x * w, y2: c2y * h }, x: ex * w, y: ey * h },
      ] });
      slide.addShape('custGeom', o);
    } else if (shape === 'poly') {
      o.points = extra.pts.map(([px, py]) => ({ x: px * w, y: py * h }));
      if (mode === 'fill' || mode === 'hole') o.points.push({ close: true });
      slide.addShape('custGeom', o);
    } else {
      if (shape === 'roundRect') o.rectRadius = radius((extra && extra.r) || 0.22, o.w, o.h);
      slide.addShape(shape, o);
    }
  });
}

/* ----------------------------------------------------------------- slides */

// Slide 1 — cover
function slide1(pres) {
  const s = pres.addSlide();
  title(s, 'Healthcare Deck', { x: 1.549, y: 1.44, w: 5.562, h: 2.524, fontSize: 72 });
  s.addShape('roundRect', {
    x: 4.287, y: 2.948, w: 2.816, h: 0.802,
    rectRadius: radius(0.5, 2.816, 0.802), fill: { color: C.green },
  });
  // Big angled band on the right (a rotated quadrilateral).
  s.addShape('custGeom', {
    x: 6.237, y: 0.404, w: 7.5, h: 6.692, rotate: 90, fill: { color: C.green },
    points: [
      { x: 0, y: 3.727 }, { x: 0, y: 0 }, { x: 1.998, y: 0 },
      { x: 7.5, y: 2.176 }, { x: 7.5, y: 6.692 }, { close: true },
    ],
  });
  textBox(s, [
    { text: '19.872', options: { color: '000000' } },
    { text: '+', options: { color: C.mint } },
  ], { x: 5.563, y: 3.033, w: 1.77, h: 0.337, fontFace: FONT_HEAD, fontSize: 14 });
  textBox(s, 'Subtitle here', { x: 5.584, y: 3.327, w: 1.77, h: 0.337, fontFace: FONT_HEAD, fontSize: 14, color: C.mint });
  body(s, [
    { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', options: { breakLine: true } },
    { text: 'Ut enim ad minim veniam, quis nostrud' },
  ], { x: 1.552, y: 4.319, w: 5.558, h: 0.97 });
}

// Slide 2 — two year cards
function slide2(pres) {
  const s = pres.addSlide();
  s.background = { color: C.grey };
  title(s, 'Empowering Wellness Through Shared Action', { x: 1.569, y: 1.291, w: 5.891, h: 2.121, fontSize: 40 });

  const cards = [
    { x: 8.052, adj: 0.11279, fill: C.white, shadow: SHADOW, year: '2026', yearColor: C.mint, textColor: C.body, tx: 8.19, tw: 1.823, leaf: C.mint },
    { x: 9.979, adj: 0.09484, fill: C.green, shadow: null, year: '2027', yearColor: C.white, textColor: C.white, tx: 10.116, tw: 1.684, leaf: C.white },
  ];
  cards.forEach(c => {
    card(s, { x: c.x, y: 1.232, w: 1.806, h: 1.983, adj: c.adj, fill: c.fill, shadow: c.shadow });
    textBox(s, c.year, {
      x: c.x + 0.14, y: 1.433, w: 1.075, h: 0.404,
      fontFace: FONT_HEAD, fontSize: 18, bold: true, color: c.yearColor, align: 'justify',
    });
    body(s, [
      { text: 'Lorem ipsum dolo sit amet, elit, sed ', options: { breakLine: true } },
      { text: 'do eiusmod' },
    ], { x: c.tx, y: 1.932, w: c.tw, h: 0.97, color: c.textColor });
    leaf(s, c.x + 1.417, 2.783, 0.389, 0.427, c.leaf);
  });

  heading(s, 'Health Solutions for All', { x: 7.889, y: 4.535, w: 2.923, fontFace: FONT_BODY, color: '000000' });
  body(s, 'Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod tempor incididunt ut labore  dolore magna aliqua. Ut enim ad minim veniam',
    { x: 7.889, y: 5.097, w: 4.087, h: 0.97 });
}

// Slide 3 — two feature tiles + a highlight
function slide3(pres) {
  const s = pres.addSlide();
  title(s, 'Transforming Lives with Modern Medicine', { x: 1.555, y: 1.534, w: 6.383, h: 1.447, fontSize: 40 });

  card(s, { x: 1.549, y: 3.997, w: 2.27, h: 2.27, adj: 0.07499, fill: C.green });
  card(s, { x: 4.009, y: 3.997, w: 2.27, h: 2.27, adj: 0.07075, fill: C.white, shadow: SHADOW });
  icon(s, 'bars', 2.359, 4.366, 0.65, 0.429, C.white);
  icon(s, 'ambulance', 4.884, 4.36, 0.518, 0.429, C.mint);
  body(s, 'Lorem ipsum dolor sit amet elit sed eiusmod tempor labore', { x: 1.625, y: 4.993, w: 2.118, h: 0.97, align: 'center', color: C.white });
  body(s, 'Lorem ipsum dolor sit amet elit sed eiusmod tempor labore', { x: 4.135, y: 4.999, w: 2.017, h: 0.97, align: 'center' });
  leaf(s, 3.43, 5.84, 0.389, 0.427, C.white);
  leaf(s, 5.89, 5.84, 0.389, 0.427, C.mint);

  s.addShape('star5', { x: 8.478, y: 1.547, w: 0.317, h: 0.31, fill: { color: C.green } });
  heading(s, 'Perfect Quality', { x: 8.829, y: 1.534, w: 1.62 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
    { x: 8.401, y: 2.059, w: 3.752, h: 0.97 });
}

// Slide 4 — two strategy cards
function slide4(pres) {
  const s = pres.addSlide();
  title(s, 'Wellness Driven by Compassion and Vision', { x: 6.667, y: 1.487, w: 5.118, h: 2.322, fontSize: 44 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud',
    { x: 6.667, y: 4.539, w: 4.903, h: 0.97 });

  const strategies = [
    { y: 1.225, adj: 0.08953, fill: C.green, shadow: null, label: 'Strategy 01', lx: 3.329, lw: 2.398, color: C.white, bx: 2.692, by: 2.075, glyph: 'hand', gx: 2.836, gy: 1.537, gs: 0.378, leafY: 2.615, leaf: C.white },
    { y: 3.147, adj: 0.11669, fill: C.white, shadow: SHADOW, label: 'Strategy 02', lx: 3.287, lw: 2.562, color: C.ink, bx: 2.681, by: 3.997, glyph: 'people', gx: 2.772, gy: 3.432, gs: 0.438, leafY: 4.537, leaf: C.mint },
  ];
  strategies.forEach(t => {
    card(s, { x: 2.457, y: t.y, w: 3.687, h: 1.811, adj: t.adj, fill: t.fill, shadow: t.shadow });
    heading(s, t.label, { x: t.lx, y: t.y + 0.335, w: t.lw, color: t.color });
    body(s, 'PLACEHOLDER',
      { x: t.bx, y: t.by, w: 3.432, h: 0.667, color: t.color === C.white ? C.white : C.body });
    icon(s, t.glyph, t.gx, t.gy, t.gs, t.gs, t.color === C.white ? C.white : C.mint, t.fill);
    leaf(s, 5.755, t.leafY, 0.389, 0.427, t.leaf);
  });

  pill(s, 'We Are', { x: 6.826, y: 5.667, w: 1.604, h: 0.313, fill: C.mint, color: C.white });
}

// Slide 5 — headline with one floating card
function slide5(pres) {
  const s = pres.addSlide();
  title(s, 'Healing Beyond Medicine, Healing With Care', { x: 1.548, y: 1.444, w: 5.079, h: 1.919, fontSize: 36 });

  card(s, { x: 9.712, y: 1.278, w: 1.673, h: 1.99, adj: 0.10121, fill: C.green });
  icon(s, 'head', 10.374, 1.563, 0.351, 0.441, C.white, C.green);
  body(s, 'Lorem ipsum dolor sit amet elit sed', { x: 9.847, y: 2.078, w: 1.403, h: 0.97, align: 'center', color: C.white });
  leaf(s, 10.996, 2.841, 0.389, 0.427, C.white);

  icon(s, 'badge', 1.641, 4.441, 0.38, 0.384, C.mint);
  heading(s, 'Best Qualification', { x: 2.156, y: 4.464, w: 2.883 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ',
    { x: 1.571, y: 5.018, w: 4.798, h: 0.97 });
}

// Slide 6 — numbered list
function slide6(pres) {
  const s = pres.addSlide();
  title(s, 'Innovative Paths for Modern Healthcare', { x: 1.574, y: 1.529, w: 6.479, h: 1.582, fontSize: 44 });

  card(s, { x: 1.451, y: 4.059, w: 1.673, h: 1.99, adj: 0.10483, fill: C.green });
  icon(s, 'badge', 2.098, 4.32, 0.38, 0.384, C.white, C.green);
  body(s, 'Lorem ipsum dolor sit amet elit sed', { x: 1.587, y: 4.86, w: 1.403, h: 0.97, align: 'center', color: C.white });
  leaf(s, 2.736, 5.675, 0.389, 0.427, C.white);

  ['01.', '02.', '03.'].forEach((num, i) => {
    const y = 1.445 + i * 1.534;
    textBox(s, num, { x: 8.67, y: y + 0.075, w: 0.593, h: 0.368, fontSize: 16, color: C.mint });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { x: 9.263, y: y, w: 2.521, h: 0.97 });
  });
}

// Slide 7 — four practice cards
function slide7(pres) {
  const s = pres.addSlide();
  title(s, 'Compassionate Care for Every Generation', { x: 1.57, y: 1.58, w: 6.486, h: 1.582, fontSize: 44 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.',
    { x: 8.387, y: 1.592, w: 3.477, h: 0.97 });
  pill(s, 'Read More', { x: 8.506, y: 2.739, w: 1.604, h: 0.313, fill: C.mint, color: C.ink });

  const practices = [
    { x: 4.085, y: 3.941, adj: 0.06236, fill: C.white, shadow: SHADOW, label: 'Practice Three', color: C.ink, glyph: 'mic', gx: 4.78, gy: 4.056, gs: 0.43, leaf: C.mint },
    { x: 6.075, y: 3.941, adj: 0.07999, fill: C.green, shadow: null, label: 'Practice One', color: C.white, glyph: 'plane', gx: 6.794, gy: 4.153, gs: 0.381, leaf: C.white },
    { x: 8.056, y: 3.943, adj: 0.07000, fill: C.white, shadow: SHADOW, label: 'Practice Four', color: C.ink, glyph: 'hand', gx: 8.777, gy: 4.127, gs: 0.378, leaf: C.mint },
    { x: 10.046, y: 3.943, adj: 0.07666, fill: C.white, shadow: SHADOW, label: 'Practice Five', color: C.ink, glyph: 'editCheck', gx: 10.751, gy: 4.12, gs: 0.356, leaf: C.mint },
  ];
  practices.forEach(p => {
    card(s, { x: p.x, y: p.y, w: 1.82, h: 2.099, adj: p.adj, fill: p.fill, shadow: p.shadow });
    icon(s, p.glyph, p.gx, p.gy, p.gs, p.gs, p.color === C.white ? C.white : C.mint, p.fill);
    textBox(s, p.label, {
      x: p.x - 0.164, y: p.y + 0.642, w: 2.147, h: 0.336,
      align: 'center', fontFace: FONT_HEAD, fontSize: 12, color: p.color, lineSpacing: 18.8,
    });
    body(s, [
      { text: 'Lorem ipsum dolor', options: { breakLine: true } },
      { text: 'amet adipiscing' },
    ], { x: p.x - 0.164, y: p.y + 1.084, w: 2.147, h: 0.667, align: 'center', color: p.color === C.white ? C.white : C.body });
    leaf(s, p.x + 1.496, p.y + 1.751, 0.321, 0.352, p.leaf);
  });
}

// Slide 8 — attribute cards with badge circles
function slide8(pres) {
  const s = pres.addSlide();
  title(s, 'Transforming Global Health Through Innovation', { x: 1.562, y: 1.582, w: 5.77, h: 1.919, fontSize: 36 });

  const attrs = [
    { x: 7.285, fill: C.green, color: C.white, textColor: C.white, label: 'Integrity', dy: 0, circleFill: C.white, circleShadow: SHADOW, glyph: 'handshake', gx: 8.181, gy: 1.382, gs: 0.362, gColor: C.mint, leaf: C.white, leafX: 9.038 },
    { x: 9.629, fill: C.white, color: C.ink, textColor: C.body, label: 'Adaptability', dy: 0.015, circleFill: C.mint, circleShadow: null, glyph: 'person', gx: 10.502, gy: 1.361, gs: 0.404, gColor: C.white, leaf: C.mint, leafX: 11.383 },
  ];
  attrs.forEach(a => {
    card(s, { x: a.x, y: 1.518, w: 2.143, h: 2.232, adj: 0.0925, fill: a.fill, shadow: SHADOW_SOFT });
    s.addShape('ellipse', { x: a.x + 0.736, y: 1.212 + a.dy, w: 0.704, h: 0.704, fill: { color: a.circleFill }, shadow: a.circleShadow ? a.circleShadow() : undefined });
    icon(s, a.glyph, a.gx, a.gy, a.gs, a.gs, a.gColor, a.circleFill);
    textBox(s, a.label, { x: a.x, y: 2.068 + a.dy, w: 2.143, h: 0.337, align: 'center', fontFace: FONT_HEAD, fontSize: 14, color: a.color });
    body(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod', { x: a.x + 0.056, y: 2.542 + a.dy, w: 2.038, h: 0.97, align: 'center', color: a.textColor });
    leaf(s, a.leafX, 3.323, 0.389, 0.427, a.leaf);
  });

  heading(s, 'Key Attributes for Success', { x: 1.549, y: 4.368, w: 3.461 });
  icon(s, 'check', 1.649, 5.004, 0.29, 0.29, C.mint);
  body(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor incididunt ut labore et dolore magna sed aliqua. Ut enim ad minim veniam, quis',
    { x: 2.086, y: 4.886, w: 3.994, h: 0.97 });
}

// Slide 9 — title with a highlight panel and two ticks
function slide9(pres) {
  const s = pres.addSlide();
  title(s, 'Building a Healthier Future Together', { x: 1.563, y: 1.379, w: 5.025, h: 2.322, fontSize: 44 });

  card(s, { x: 7.28, y: 1.076, w: 3.867, h: 2.573, adj: 0.10256, fill: C.green });
  textBox(s, 'Rewards and Recognition', { x: 7.558, y: 1.342, w: 2.932, h: 0.337, fontFace: FONT_HEAD, fontSize: 14, color: C.white });
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua',
    { x: 7.558, y: 1.857, w: 3.404, h: 0.97, color: C.white });
  pill(s, 'Read More', { x: 7.65, y: 3.004, w: 1.599, h: 0.324, fill: C.white, color: C.mint, shadow: SHADOW });
  leaf(s, 10.49, 2.93, 0.655, 0.718, C.white);

  [['Stage Fright', 1.716, 2.125], ['Confidence ', 3.998, 4.407]].forEach(([label, cx, tx]) => {
    icon(s, 'check', cx, 4.411, 0.286, 0.286, C.mint);
    heading(s, label, { x: tx, y: 4.371, w: 1.591 });
  });
  body(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor incididunt ut labore et dolore magna sed aliqua. Ut enim ad minim veniam, quis nostrud exercitation',
    { x: 1.563, y: 4.99, w: 4.722, h: 0.97 });
}

// Slide 10 — two stat panels
function slide10(pres) {
  const s = pres.addSlide();
  title(s, 'Human Health at the Core', { x: 1.57, y: 1.562, w: 4.797, h: 1.582, fontSize: 44 });

  const stats = [
    { y: 1.369, adj: 0.11953, fill: C.green, shadow: null, stat: '98%', bold: true, statColor: C.white, tx: 7.105, ty: 1.734, bx: 7.105, by: 2.538, bh: 0.675, textColor: C.white, leafY: 3.06, leaf: C.white },
    { y: 3.87, adj: 0.0925, fill: C.white, shadow: SHADOW_SOFT, stat: '123K', bold: false, statColor: C.mint, tx: 7.241, ty: 4.235, bx: 7.241, by: 5.043, bh: 0.667, textColor: C.body, leafY: 5.561, leaf: C.mint },
  ];
  stats.forEach(p => {
    card(s, { x: 6.823, y: p.y, w: 5.083, h: 2.241, adj: p.adj, fill: p.fill, shadow: p.shadow });
    textBox(s, p.stat, { x: p.tx, y: p.ty, w: 3.194, h: 0.707, fontFace: FONT_HEAD, fontSize: 36, bold: p.bold, color: p.statColor });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et',
      { x: p.bx, y: p.by, w: 4.753, h: p.bh, color: p.textColor });
    leaf(s, 11.404, p.leafY, 0.502, 0.551, p.leaf);
  });
}

// Slide 11 — guidance card, bullet rows and a badge
function slide11(pres) {
  const s = pres.addSlide();
  title(s, 'Caring Beyond Hospital Walls', { x: 6.656, y: 1.54, w: 4.999, h: 1.582, fontSize: 44 });

  card(s, { x: 1.55, y: 1.352, w: 3.937, h: 1.77, adj: 0.07948, fill: C.white, shadow: SHADOW });
  heading(s, 'Guidance', { x: 1.817, y: 1.658, w: 1.537 });
  icon(s, 'docs', 3.15, 1.649, 0.355, 0.355, C.mint);
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor ut labore', { x: 1.817, y: 2.098, w: 3.42, h: 0.667 });
  leaf(s, 5.098, 2.695, 0.389, 0.427, C.mint);

  heading(s, 'Expertise in Public Speaking', { x: 1.533, y: 3.876, w: 3.12 });
  [['calculator', 4.595, 0.241, 0.276, 4.387], ['mail', 5.496, 0.276, 0.276, 5.276]].forEach(([glyph, gy, gw, gh, ty]) => {
    icon(s, glyph, 1.636, gy, gw, gh, C.mint);
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { x: 2.206, y: ty, w: 3.553, h: 0.667 });
  });

  // Grouped badge, bottom right.
  s.addShape('roundRect', { x: 9.911, y: 4.691, w: 2.155, h: 0.817, rectRadius: radius(0.15031, 2.155, 0.817), fill: { color: C.green } });
  icon(s, 'user', 10.130, 4.903, 0.368, 0.368, C.white, C.green);
  textBox(s, 'Build Loyalty is A Must', { x: 10.667, y: 4.847, w: 1.324, h: 0.505, color: C.white });
}

// Slide 12 — closing panel
function slide12(pres) {
  const s = pres.addSlide();
  // Panel with the two left corners rounded (a rotated "round2SameRect").
  s.addShape('round2SameRect', { x: 6.631, y: -0.786, w: 4.333, h: 9.072, rotate: 270, fill: { color: C.green } });

  textBox(s, 'Key Attributes for Success', { x: 4.678, y: 2.133, w: 3.461, h: 0.337, fontFace: FONT_HEAD, fontSize: 14, color: C.white });
  s.addShape('line', { x: 7.4, y: 2.305, w: 1.394, h: 0, line: { color: C.white, width: 1.5 } });
  title(s, 'Thank You!', { x: 4.678, y: 2.473, w: 7.447, h: 1.717, fontSize: 96, color: C.white });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud',
    { x: 4.678, y: 4.088, w: 6.817, h: 0.667, color: C.white });
  textBox(s, 'WWW . YOUR WEBSITE . COM', { x: 4.734, y: 5.079, w: 3.461, h: 0.337, fontFace: FONT_HEAD, fontSize: 14, color: C.white });
}

/* ------------------------------------------------------------------ build */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pres.layout = 'DECK';
  pres.author = 'pptxgenjs';
  pres.title = 'Healthcare Deck';

  [slide1, slide2, slide3, slide4, slide5, slide6,
   slide7, slide8, slide9, slide10, slide11, slide12].forEach(fn => fn(pres));

  return pres.writeFile({ fileName: path.join(__dirname, '0e2f509a-6ae5-43b2-996a-f6267e04494a_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
