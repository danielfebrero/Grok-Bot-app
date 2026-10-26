/**
 * Jobcy — Job Board Presentation Template
 * Standalone pptxgenjs recreation (42 slides, 26.66in x 15in / 16:9 widescreen).
 *
 * Raster images in the source deck are replaced by programmatic placeholders:
 *  - `pic()`   -> the template's own "Drag Picture here" pattern boxes
 *  - `phone()` / `laptop()` -> device frames drawn around those same boxes
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const GREEN = '3DB719', DARK = '1E1F21', ORANGE = 'F88B03', GRAY = '808080',
      WHITE = 'FFFFFF', SILVER = 'BFBFBF', LGRAY = 'D9D9D9', HAIR = 'F2F2F2',
      PALE = 'F1F4FB', ICE = 'F5F5F5', GREY2 = 'BEBEBE',
      GREEN2 = '7CC244', OLIVE = '7CA52C', GREEN_D = '2E8913', ORANGE_D = 'BA6802',
      GREEN_M = '5D9233', OLIVE_D = '5D7C21', EEE = 'EEEEEE',
      MINT = 'D8F1D1', PLACEHOLDER = 'EAF7E7';
const HEAD = 'Poppins', BODY = 'Roboto';

const LOREM_LONG = 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices. Donec euismod scelerisque ligula. ';
const LOREM_MED  = 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices. ';
const LOREM_SHORT = 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc';
const LOREM_TINY = 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus';

/* --------------------------------------------------------------- helpers */
const at = (x, y, w, h) => ({ x, y, w, h });

/** filled / outlined preset shape */
function shape(s, kind, box, opt = {}) {
  s.addShape(kind, Object.assign({}, box, opt));
}
const rect  = (s, box, o) => shape(s, 'rect', box, o);
const oval  = (s, box, o) => shape(s, 'ellipse', box, o);
const round = (s, box, radius, o) => shape(s, 'roundRect', box, Object.assign({ rectRadius: radius }, o));

/** flag / bookmark: rectangle with a notch cut out of the bottom edge */
function bookmark(s, x, y, w, h, color) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color },
    points: [{ x: 0, y: 0 }, { x: 0, y: h }, { x: w / 2, y: h - 0.53 }, { x: w, y: h }, { x: w, y: 0 }, { close: true }],
  });
}

/** thin rule: horizontal or vertical hairline */
function rule(s, x, y, w, h, color = HAIR, pt = 1) {
  s.addShape('line', { x, y, w, h, line: { color, width: pt } });
}

/** text block; `runs` is a string or an array of {text, options} */
function tx(s, runs, box, opt = {}) {
  const base = {
    fontFace: BODY, fontSize: 24, color: GRAY, align: 'left', valign: 'top',
    margin: [2.9, 7.2, 2.9, 7.2], isTextBox: true, wrap: true,
  };
  s.addText(runs, Object.assign(base, box, opt));
}

/* -------- recurring text roles (kicker / title / paragraph / heading) ---- */
const kicker = (s, t, x, y, w, o) =>
  tx(s, t, at(x, y, w, 0.44), Object.assign({ fontSize: 20, bold: true, color: GREEN }, o));
const title = (s, runs, x, y, w, h, o) =>
  tx(s, runs, at(x, y, w, h), Object.assign({ fontFace: HEAD, fontSize: 80, bold: true, color: DARK }, o));
const para = (s, t, x, y, w, h, o) =>
  tx(s, t, at(x, y, w, h), Object.assign({ lineSpacingMultiple: 1.5 }, o));
const heading = (s, t, x, y, w, o) =>
  tx(s, t, at(x, y, w, 0.64), Object.assign({ fontSize: 32, bold: true, color: DARK }, o));

/** green pill button with centred caps label */
function button(s, label, x, y, w = 3.45, h = 0.92, fill = GREEN, color = WHITE, bold = true) {
  round(s, at(x, y, w, h), 0.18, { fill: { color: fill } });
  tx(s, label, at(x, y + 0.2, w, 0.5), { align: 'center', bold, color, fontSize: 24 });
}

/* --------------------------------------------------------- picture stand-ins */
/** the deck's own empty picture placeholder (hatched mint box + prompt) */
function pic(s, x, y, w, h, opt = {}) {
  if (opt.plain !== true) rect(s, at(x, y, w, h), { fill: { color: PLACEHOLDER } });
  tx(s, opt.label || 'Drag Picture here', at(x, y + 0.06, w, 0.7),
     { align: 'center', fontFace: 'Arial', fontSize: opt.size || 40, color: opt.labelColor || '000000' });
  if (opt.marker !== false && w > 1.2 && h > 1.2) {          // the tiny "no image yet" badge
    const mw = 0.62, mh = 0.5, mx = x + (w - mw) / 2, my = y + (h - mh) / 2;
    rect(s, at(mx, my, mw, mh), { fill: { color: WHITE }, line: { color: SILVER, width: 0.75 } });
    shape(s, 'triangle', at(mx + 0.1, my + 0.2, 0.32, 0.22), { fill: { color: '7FA8D9' } });
    oval(s, at(mx + 0.42, my + 0.1, 0.1, 0.1), { fill: { color: 'F2B44C' } });
  }
}
/** vertical wash, approximated with stacked translucent bands */
function wash(s, x, y, w, h, color, fromPct, toPct, bands = 22) {
  for (let i = 0; i < bands; i++) {
    const t = bands === 1 ? 0 : i / (bands - 1);
    rect(s, at(x, y + (h * i) / bands, w, h / bands + 0.01),
         { fill: { color, transparency: Math.round(fromPct + (toPct - fromPct) * t) } });
  }
}

/** rotate an offset around a pivot — lets a tilted mockup keep its parts aligned */
function spin(cx, cy, dx, dy, deg) {
  const r = (deg * Math.PI) / 180, c = Math.cos(r), n = Math.sin(r);
  return [cx + dx * c - dy * n, cy + dx * n + dy * c];
}

/** phone mockup: dark rounded body around an empty picture placeholder */
function phone(s, cx, cy, w, h, deg = 0, bezel = 0.26) {
  const place = (dx, dy, pw, ph, draw) => {
    const [x, y] = spin(cx, cy, dx, dy, deg);
    draw(x - pw / 2, y - ph / 2, pw, ph);
  };
  place(0, 0, w, h, (x, y, pw, ph) => round(s, at(x, y, pw, ph), 0.09, { fill: { color: '1C1C1C' }, rotate: deg }));
  place(0, 0, w - bezel * 2, h - bezel * 2, (x, y, pw, ph) => {
    rect(s, at(x, y, pw, ph), { fill: { color: PLACEHOLDER }, rotate: deg });
    tx(s, 'Drag Picture here', at(x, y + 0.1, pw, 0.7),
       { align: 'center', fontFace: 'Arial', fontSize: 26, color: '000000', rotate: deg });
  });
  place(0, -h / 2 + bezel + 0.16, w * 0.34, 0.32,
        (x, y, pw, ph) => round(s, at(x, y, pw, ph), 0.4, { fill: { color: '1C1C1C' }, rotate: deg }));
}

/** laptop mockup: screen with an empty placeholder plus a tapered base */
function laptop(s, x, y, w, h) {
  rect(s, at(x, y, w, h), { fill: { color: '1C1C1C' } });
  rect(s, at(x + 0.28, y + 0.28, w - 0.56, h - 0.9), { fill: { color: PLACEHOLDER } });
  tx(s, 'Drag picture here', at(x + 0.28, y + 0.36, w - 0.56, 0.7),
     { align: 'center', fontFace: 'Arial', fontSize: 28, color: '000000' });
  shape(s, 'trapezoid', at(x - 1.15, y + h, w + 2.3, 0.85), { fill: { color: 'C9CDD2' } });
}

/* --------------------------------------------------------------- branding */
/** "Jobcy" wordmark: magnifier disc + word, drawn from primitives */
function logo(s, x, y, scale = 1, wordColor = DARK, tailColor = GREEN) {
  const disc = wordColor === WHITE ? WHITE : GREEN, glyph = wordColor === WHITE ? GREEN : WHITE;
  oval(s, at(x, y, 0.62 * scale, 0.62 * scale), { fill: { color: disc } });
  oval(s, at(x + 0.15 * scale, y + 0.15 * scale, 0.32 * scale, 0.32 * scale), { fill: { color: glyph } });
  rect(s, at(x + 0.43 * scale, y + 0.43 * scale, 0.24 * scale, 0.1 * scale), { fill: { color: disc }, rotate: 45 });
  tx(s, [{ text: 'Job', options: { color: wordColor } }, { text: 'cy', options: { color: tailColor } }],
     at(x + 0.58 * scale, y - 0.04 * scale, 2.4 * scale, 0.72 * scale),
     { fontFace: HEAD, fontSize: 30 * scale, bold: true });
}

/** the same wordmark rotated to run bottom-to-top along a green band */
function logoVertical(s, x, y, scale = 1, wordColor = WHITE) {
  const d = 0.62 * scale;
  oval(s, at(x, y + 1.9 * scale, d, d), { fill: { color: wordColor } });
  oval(s, at(x + 0.15 * scale, y + 2.05 * scale, 0.32 * scale, 0.32 * scale), { fill: { color: GREEN } });
  rect(s, at(x + 0.43 * scale, y + 2.33 * scale, 0.22 * scale, 0.1 * scale), { fill: { color: wordColor }, rotate: 45 });
  tx(s, 'Jobcy', at(x - 0.85 * scale, y + 0.85 * scale, 2.4 * scale, 0.72 * scale),
     { fontFace: HEAD, fontSize: 30 * scale, bold: true, color: wordColor, rotate: 270, align: 'center' });
}

/** top navigation links used on the web-style slides */
function nav(s, y, color = WHITE, bold = true, items) {
  (items || [['Job Alerts', 8.88, 2.15], ['Post Job', 12.17, 1.94], ['Job Dashboard', 15.24, 2.54]])
    .forEach(([t, x, w]) => tx(s, t, at(x, y, w, 0.51), { align: 'center', bold, color, fontSize: 24 }));
}

/** rounded REGISTER chip with a small user-card glyph */
function registerChip(s, x, y, w = 3.02, h = 0.8, color = WHITE, fill = null) {
  if (fill) round(s, at(x, y, w, h), 0.5, { fill: { color: fill } });
  const gx = x + 0.36, gy = y + 0.2;
  round(s, at(gx, gy, 0.42, 0.36), 0.06, { fill: { type: 'none' }, line: { color, width: 1.5 } });
  oval(s, at(gx + 0.09, gy + 0.08, 0.13, 0.13), { line: { color, width: 1.5 }, fill: { type: 'none' } });
  rule(s, gx + 0.25, gy + 0.11, 0.1, 0, color, 1.5);
  rule(s, gx + 0.25, gy + 0.19, 0.1, 0, color, 1.5);
  tx(s, 'REGISTER', at(x + 0.87, y + 0.15, 1.77, 0.51), { bold: true, color, fontSize: 24 });
}

/* --------------------------------------------------------- outline icons */
/* Line-art pictograms rebuilt from primitives; `c1`/`c2` are the two accents. */
const ICONS = {
  search(s, x, y, k, c1) {
    oval(s, at(x, y, 0.6 * k, 0.6 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.52 * k, y + 0.52 * k, 0.3 * k, 0.3 * k, c1, 2.4 * k);
  },
  pin(s, x, y, k, c1) {                                    // teardrop map marker
    shape(s, 'teardrop', at(x + 0.06 * k, y, 0.62 * k, 0.62 * k),
          { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' }, rotate: 135 });
    oval(s, at(x + 0.24 * k, y + 0.18 * k, 0.24 * k, 0.24 * k), { line: { color: c1, width: 2.2 * k }, fill: { type: 'none' } });
  },
  folder(s, x, y, k, c1, c2) {
    rect(s, at(x + 0.02 * k, y + 0.06 * k, 0.4 * k, 0.14 * k), { fill: { color: c2 } });
    round(s, at(x, y + 0.18 * k, 0.9 * k, 0.64 * k), 0.05, { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.2 * k, y + 0.38 * k, 0.42 * k, 0, c2, 2.2 * k);
    rule(s, x + 0.2 * k, y + 0.52 * k, 0.28 * k, 0, c2, 2.2 * k);
  },
  chart(s, x, y, k, c1, c2) {                              // bar chart with a rising arrow
    [[0.06, 0.24], [0.26, 0.42], [0.46, 0.58]].forEach(([dx, hh]) =>
      rect(s, at(x + dx * k, y + 0.82 * k - hh * k, 0.13 * k, hh * k), { line: { color: c1, width: 2 * k }, fill: { type: 'none' } }));
    rule(s, x, y + 0.84 * k, 0.78 * k, 0, c1, 2.4 * k);
    rule(s, x + 0.1 * k, y + 0.36 * k, 0.3 * k, -0.18 * k, c2, 2.4 * k);
    rule(s, x + 0.4 * k, y + 0.18 * k, 0.3 * k, -0.14 * k, c2, 2.4 * k);
    shape(s, 'triangle', at(x + 0.6 * k, y - 0.06 * k, 0.2 * k, 0.2 * k), { fill: { color: c2 }, rotate: 45 });
  },
  doc(s, x, y, k, c1, c2) {                                // sheet with lines + magnifier
    rect(s, at(x + 0.04 * k, y, 0.62 * k, 0.82 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    [0.18, 0.34, 0.5].forEach(dy => rule(s, x + 0.16 * k, y + dy * k, 0.38 * k, 0, c2, 2 * k));
    oval(s, at(x + 0.38 * k, y + 0.42 * k, 0.36 * k, 0.36 * k), { line: { color: c2, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.7 * k, y + 0.74 * k, 0.16 * k, 0.14 * k, c2, 2.4 * k);
  },
  people(s, x, y, k, c1, c2) {                             // person silhouette + tick badge
    oval(s, at(x + 0.08 * k, y, 0.26 * k, 0.26 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    s.addShape('custGeom', {
      x: x + 0.02 * k, y: y + 0.3 * k, w: 0.4 * k, h: 0.5 * k, line: { color: c1, width: 2.4 * k }, fill: { type: 'none' },
      points: [{ x: 0, y: 0.5 * k }, { x: 0, y: 0.18 * k }, { x: 0.2 * k, y: 0 }, { x: 0.4 * k, y: 0.18 * k }, { x: 0.4 * k, y: 0.5 * k }],
    });
    oval(s, at(x + 0.42 * k, y + 0.36 * k, 0.42 * k, 0.42 * k), { line: { color: c2, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.51 * k, y + 0.58 * k, 0.09 * k, 0.09 * k, c2, 2.2 * k);
    rule(s, x + 0.6 * k, y + 0.67 * k, 0.16 * k, -0.2 * k, c2, 2.2 * k);
  },
  briefcase(s, x, y, k, c1, c2) {
    rect(s, at(x + 0.26 * k, y, 0.34 * k, 0.18 * k), { line: { color: c2, width: 2.4 * k }, fill: { type: 'none' } });
    round(s, at(x, y + 0.18 * k, 0.86 * k, 0.62 * k), 0.05, { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x, y + 0.42 * k, 0.86 * k, 0, c1, 2.4 * k);
  },
  check(s, x, y, k, c1) {
    oval(s, at(x, y, 0.6 * k, 0.64 * k), { fill: { color: c1 } });
    rule(s, x + 0.15 * k, y + 0.32 * k, 0.1 * k, 0.11 * k, WHITE, 2.6 * k);
    rule(s, x + 0.25 * k, y + 0.43 * k, 0.21 * k, -0.24 * k, WHITE, 2.6 * k);
  },
  plane(s, x, y, k, c1, c2) {                              // paper plane
    s.addShape('custGeom', {
      x, y, w: 0.72 * k, h: 0.66 * k, line: { color: c1, width: 2.4 * k }, fill: { type: 'none' },
      points: [{ x: 0.72 * k, y: 0 }, { x: 0, y: 0.34 * k }, { x: 0.3 * k, y: 0.42 * k },
               { x: 0.42 * k, y: 0.66 * k }, { close: true }],
    });
    rule(s, x + 0.3 * k, y + 0.42 * k, 0.42 * k, -0.42 * k, c2, 2 * k);
  },
  award(s, x, y, k, c1, c2) {                              // sheet with a star seal
    rect(s, at(x + 0.04 * k, y, 0.62 * k, 0.82 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    [0.18, 0.34].forEach(dy => rule(s, x + 0.16 * k, y + dy * k, 0.38 * k, 0, c1, 2 * k));
    shape(s, 'star5', at(x + 0.36 * k, y + 0.4 * k, 0.44 * k, 0.44 * k), { line: { color: c2, width: 2.2 * k }, fill: { type: 'none' } });
  },
  phone(s, x, y, k, c1) {
    round(s, at(x + 0.2 * k, y, 0.36 * k, 0.72 * k), 0.06, { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.3 * k, y + 0.6 * k, 0.16 * k, 0, c1, 2.4 * k);
  },
  mail(s, x, y, k, c1, c2) {
    rect(s, at(x, y + 0.12 * k, 0.8 * k, 0.56 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x, y + 0.12 * k, 0.4 * k, 0.3 * k, c2, 2.2 * k);
    rule(s, x + 0.8 * k, y + 0.12 * k, -0.4 * k, 0.3 * k, c2, 2.2 * k);
  },
  gear(s, x, y, k, c1, c2) {
    shape(s, 'gear6', at(x, y, 0.7 * k, 0.7 * k), { line: { color: c1, width: 2.2 * k }, fill: { type: 'none' } });
    oval(s, at(x + 0.25 * k, y + 0.25 * k, 0.2 * k, 0.2 * k), { line: { color: c2, width: 2.2 * k }, fill: { type: 'none' } });
  },
  trophy(s, x, y, k, c1, c2) {
    s.addShape('custGeom', {
      x: x + 0.16 * k, y, w: 0.38 * k, h: 0.44 * k, line: { color: c1, width: 2.4 * k }, fill: { type: 'none' },
      points: [{ x: 0, y: 0 }, { x: 0.38 * k, y: 0 }, { x: 0.3 * k, y: 0.44 * k }, { x: 0.08 * k, y: 0.44 * k }, { close: true }],
    });
    rule(s, x + 0.16 * k, y + 0.1 * k, -0.14 * k, 0.16 * k, c2, 2.2 * k);
    rule(s, x + 0.54 * k, y + 0.1 * k, 0.14 * k, 0.16 * k, c2, 2.2 * k);
    rule(s, x + 0.2 * k, y + 0.68 * k, 0.3 * k, 0, c1, 2.4 * k);
    rule(s, x + 0.35 * k, y + 0.44 * k, 0, 0.24 * k, c1, 2.4 * k);
  },
  blueprint(s, x, y, k, c1, c2) {                          // rolled drawing
    round(s, at(x, y + 0.06 * k, 0.2 * k, 0.72 * k), 0.08, { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rect(s, at(x + 0.2 * k, y, 0.62 * k, 0.7 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rect(s, at(x + 0.36 * k, y + 0.16 * k, 0.28 * k, 0.4 * k), { line: { color: c2, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x + 0.36 * k, y + 0.36 * k, 0.28 * k, 0, c2, 2 * k);
  },
  money(s, x, y, k, c1, c2) {                              // price tag with $
    rect(s, at(x, y + 0.06 * k, 0.66 * k, 0.5 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    tx(s, '$', at(x + 0.06 * k, y + 0.1 * k, 0.5 * k, 0.4 * k), { align: 'center', color: c2, fontSize: 18 * k, bold: true });
    shape(s, 'arc', at(x + 0.3 * k, y + 0.44 * k, 0.44 * k, 0.4 * k), { line: { color: c2, width: 2.4 * k } });
  },
  board(s, x, y, k, c1, c2) {                              // presentation board
    rect(s, at(x, y + 0.1 * k, 0.8 * k, 0.56 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    rule(s, x, y + 0.1 * k, 0.8 * k, 0, c1, 2.4 * k);
    tx(s, '$', at(x + 0.06 * k, y + 0.18 * k, 0.3 * k, 0.36 * k), { align: 'center', color: c2, fontSize: 15 * k, bold: true });
    [[0.4, 0.16], [0.54, 0.26], [0.68, 0.34]].forEach(([dx, hh]) =>
      rect(s, at(x + dx * k, y + 0.6 * k - hh * k, 0.08 * k, hh * k), { fill: { color: c2 } }));
    rule(s, x + 0.4 * k, y + 0.66 * k, 0, 0.16 * k, c1, 2.2 * k);
  },
  server(s, x, y, k, c1, c2) {                             // stacked drives
    [0, 0.26, 0.52].forEach((dy, i) => {
      rect(s, at(x, y + 0.06 * k + dy * k, 0.8 * k, 0.2 * k), { line: { color: i === 1 ? c2 : c1, width: 2.2 * k }, fill: { type: 'none' } });
      oval(s, at(x + 0.62 * k, y + 0.13 * k + dy * k, 0.07 * k, 0.07 * k), { fill: { color: i === 1 ? c2 : c1 } });
    });
  },
  pen(s, x, y, k, c1, c2) {                                // signed sheet
    rect(s, at(x, y, 0.62 * k, 0.78 * k), { line: { color: c1, width: 2.4 * k }, fill: { type: 'none' } });
    [0.16, 0.3].forEach(dy => rule(s, x + 0.12 * k, y + dy * k, 0.36 * k, 0, c1, 2 * k));
    rule(s, x + 0.3 * k, y + 0.72 * k, 0.5 * k, -0.5 * k, c2, 2.6 * k);
    shape(s, 'triangle', at(x + 0.7 * k, y + 0.08 * k, 0.18 * k, 0.18 * k), { fill: { color: c2 }, rotate: 45 });
  },
  layers(s, x, y, k, c1, c2) {                             // stacked sheets
    [0, 0.2, 0.4].forEach((dy, i) =>
      shape(s, 'diamond', at(x, y + dy * k, 0.86 * k, 0.4 * k),
            { line: { color: i === 0 ? c2 : c1, width: 2.2 * k }, fill: { type: 'none' } }));
  },
  teacher(s, x, y, k, c1, c2) {                            // person beside a board
    oval(s, at(x + 0.04 * k, y, 0.22 * k, 0.22 * k), { fill: { color: c2 } });
    rect(s, at(x, y + 0.26 * k, 0.3 * k, 0.5 * k), { fill: { color: c2 } });
    rect(s, at(x + 0.38 * k, y, 0.44 * k, 0.76 * k), { line: { color: c1, width: 2.2 * k }, fill: { type: 'none' } });
    [0.14, 0.32, 0.5].forEach(dy => rule(s, x + 0.46 * k, y + dy * k, 0.28 * k, 0, c1, 2 * k));
  },
  bank(s, x, y, k, c1, c2) {                               // classical facade
    shape(s, 'triangle', at(x, y, 0.86 * k, 0.24 * k), { fill: { color: c2 } });
    [0.1, 0.35, 0.6].forEach(dx => rect(s, at(x + dx * k + 0.05 * k, y + 0.28 * k, 0.11 * k, 0.42 * k), { fill: { color: c1 } }));
    rect(s, at(x, y + 0.72 * k, 0.86 * k, 0.1 * k), { fill: { color: c1 } });
  },
  building(s, x, y, k, c1, c2) {                           // office block with a map pin
    rect(s, at(x, y + 0.08 * k, 0.34 * k, 0.72 * k), { line: { color: c1, width: 2.2 * k }, fill: { type: 'none' } });
    [0.2, 0.38, 0.56].forEach(dy => rule(s, x + 0.08 * k, y + dy * k, 0.18 * k, 0, c1, 2 * k));
    shape(s, 'teardrop', at(x + 0.4 * k, y + 0.3 * k, 0.4 * k, 0.4 * k),
          { line: { color: c2, width: 2.2 * k }, fill: { type: 'none' }, rotate: 135 });
  },
  bag(s, x, y, k, c1, c2) {                                // shopping / hire bag
    s.addShape('custGeom', {
      x, y: y + 0.2 * k, w: 0.76 * k, h: 0.62 * k, line: { color: c1, width: 2.4 * k }, fill: { type: 'none' },
      points: [{ x: 0.08 * k, y: 0 }, { x: 0.68 * k, y: 0 }, { x: 0.76 * k, y: 0.62 * k }, { x: 0, y: 0.62 * k }, { close: true }],
    });
    shape(s, 'arc', at(x + 0.2 * k, y, 0.36 * k, 0.4 * k), { line: { color: c2, width: 2.4 * k } });
  },
};
function icon(s, name, x, y, k = 1, c1 = GREEN, c2 = ORANGE) { ICONS[name](s, x, y, k, c1, c2); }

/** the three round social buttons under team member cards */
function socials(s, x, y, d = 0.84, gap = 1.13) {
  ['f', '𝕏', '◎'].forEach((g, i) => {
    oval(s, at(x + i * gap, y, d, d), { fill: { color: GREEN } });
    tx(s, g, at(x + i * gap, y + 0.19, d, 0.45), { align: 'center', color: WHITE, fontSize: 22, bold: true });
  });
}

/* =================================================================== slides */
const SLIDES = [];
const slide = fn => SLIDES.push(fn);

/* 1 — hero: "Find The Best Job For You Here" + search bar */
slide(s => {
  s.addShape('custGeom', {
    x: 14.86, y: 0, w: 11.8, h: 12.71, fill: { color: GREEN },
    points: [{ x: 0, y: 0 }, { x: 11.8, y: 0 }, { x: 11.8, y: 12.71 }, { x: 2.0, y: 12.71 },
             { x: 0, y: 10.71, curve: { type: 'cubic', x1: 0.9, y1: 12.71, x2: 0, y2: 11.81 } }, { close: true }],
  });
  pic(s, 14.86, 0, 11.8, 12.71, { plain: true });
  logo(s, 1.51, 1.15);
  title(s, [{ text: 'Find The ' }, { text: 'Best Job ', options: { color: GREEN } }, { text: 'For You Here' }],
        1.51, 3.72, 12.24, 3.33, { fontSize: 96 });
  tx(s, 'JOB BOARD PRESENTATION TEMPLATE', at(1.51, 7.39, 10.45, 0.57),
     { fontSize: 28, bold: true, charSpacing: 3, color: GREEN });

  round(s, at(1.51, 9.47, 17.09, 1.79), 0.28, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 20, offset: 4, angle: 90 } });
  [['Job or Position', 2.14, 3.0, 'search'], ['Location', 6.37, 7.41, 'pin'], ['Category', 10.6, 11.64, 'folder']]
    .forEach(([label, bx, tx0, ic]) => {
      round(s, at(bx, 9.9, 3.57, 0.92), 0.18, { fill: { type: 'none' }, line: { color: HAIR, width: 1.5 } });
      icon(s, ic, bx + 0.32, 10.14, 0.72, GREEN, GREEN);
      tx(s, label, at(tx0, 10.1, 2.47, 0.51), { color: LGRAY });
    });
  round(s, at(15.07, 9.9, 2.96, 0.92), 0.3, { fill: { color: ORANGE } });
  tx(s, 'SEARCH', at(15.07, 10.11, 2.96, 0.51), { align: 'center', bold: true, color: WHITE });

  round(s, at(15.65, 2.85, 4.01, 1.79), 0.14, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.14, blur: 18, offset: 4, angle: 90 } });
  tx(s, 'Hiring designer', at(16.1, 3.23, 2.94, 0.57), { fontSize: 28, bold: true, color: DARK });
  tx(s, 'Job or Position', at(16.1, 3.83, 2.47, 0.51));
});

/* 2 — "Browse Jobs" hero over a dimmed photo */
slide(s => {
  pic(s, 0, 0, 26.66, 15);
  rect(s, at(0, 0, 26.66, 15), { fill: { color: '000000', transparency: 64 } });
  logo(s, 1.47, 0.79, 1, WHITE);
  nav(s, 0.85);
  registerChip(s, 22.48, 0.7, 3.02, 0.8, WHITE, GREEN);
  title(s, 'Browse Jobs', 8.09, 4.72, 10.48, 2.04, { fontSize: 115, color: WHITE, align: 'center' });
  tx(s, 'JOB BOARD PRESENTATION TEMPLATE', at(8.85, 6.66, 8.96, 0.57),
     { align: 'center', fontSize: 28, bold: true, charSpacing: 3, color: GREEN });

  round(s, at(4.25, 7.92, 17.94, 2.12), 0.1, { fill: { color: 'F2F2F2', transparency: 70 } });
  round(s, at(4.47, 8.2, 13.5, 1.56), 0.05, { fill: { color: WHITE } });
  round(s, at(17.97, 8.2, 4.0, 1.56), 0.05, { fill: { color: ORANGE } });
  icon(s, 'search', 5.03, 8.85, 0.44, GREEN, GREEN);
  tx(s, 'Job title or keywords', at(5.59, 8.73, 3.32, 0.51), { color: SILVER });
  rule(s, 9.63, 8.59, 0, 0.78, HAIR, 1);
  icon(s, 'pin', 10.14, 8.74, 0.6, GREEN, GREEN);
  tx(s, 'All Location', at(10.74, 8.73, 2.16, 0.51), { color: SILVER });
  tx(s, '▾', at(12.85, 8.72, 0.4, 0.5), { color: SILVER, fontSize: 20 });
  rule(s, 13.71, 8.59, 0, 0.78, HAIR, 1);
  icon(s, 'folder', 14.14, 8.74, 0.55, GREEN, GREEN);
  tx(s, 'All Category', at(14.89, 8.73, 2.16, 0.51), { color: SILVER });
  tx(s, '▾', at(17.01, 8.72, 0.4, 0.5), { color: SILVER, fontSize: 20 });
  tx(s, 'Find jobs', at(18.89, 8.73, 2.16, 0.51), { align: 'center', bold: true, color: WHITE });
});

/* 3 — "LOOKING FOR A JOB?" carousel banner */
slide(s => {
  pic(s, 0, 2.07, 26.66, 12.93);
  rect(s, at(0, 2.07, 26.66, 12.93), { fill: { color: DARK, transparency: 30 } });
  logo(s, 0.96, 0.69);
  nav(s, 0.75, GRAY, false);
  registerChip(s, 22.93, 0.6, 3.02, 0.8, WHITE, GREEN);
  title(s, 'LOOKING FOR A  JOB?', 5.19, 6.18, 16.28, 1.72, { fontSize: 96, color: WHITE, align: 'center' });
  tx(s, 'JOB BOARD PRESENTATION TEMPLATE', at(8.01, 7.9, 10.65, 0.57),
     { align: 'center', fontSize: 28, bold: true, charSpacing: 3, color: GREEN });
  round(s, at(11.86, 9.74, 2.95, 0.92), 0.16, { fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
  tx(s, 'GET STARTED', at(11.86, 9.91, 2.95, 0.51), { align: 'center', bold: true, color: WHITE });
  [[1.06, '‹'], [25.2, '›']].forEach(([x, g]) => {
    oval(s, at(x, 7.76, 0.74, 0.74), { fill: { color: DARK } });
    tx(s, g, at(x, 7.82, 0.74, 0.6), { align: 'center', color: WHITE, fontSize: 28, bold: true });
  });
  [0, 1, 2, 3, 4].forEach(i =>
    shape(s, 'diamond', at(12.31 + i * 0.47, 12.9, 0.17, 0.17), { fill: { color: i === 1 ? GREEN : GREY2 } }));
});

/* 4 — "Find your dream job" + full-height green band */
slide(s => {
  rect(s, at(0, 0, 5.81, 15), { fill: { color: GREEN } });
  pic(s, 1.9, 1.65, 11.43, 13.36);
  kicker(s, 'Welcome Messages', 15.0, 2.37, 3.13);
  title(s, 'Find your dream job talented employed', 15.0, 3.0, 10.02, 4.14);
  para(s, LOREM_LONG, 15.0, 7.5, 10.02, 3.0);
  para(s, 'Donec eget massa a diam condimentum pretium. Aliquam erat volutpat. Integer ut tincidunt orci. Etiam tristique, elit ut consectetur iaculis, metus lectus mattis justo, vel mollis eros neque quis augue. ',
       15.0, 10.8, 10.02, 2.4);
});

/* 5 — pull quote card */
slide(s => {
  pic(s, 14.31, 0, 12.35, 15);
  s.addShape('custGeom', {
    x: 1.77, y: 3.26, w: 15.39, h: 9.2, fill: { color: WHITE },
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 24, offset: 5, angle: 90 },
    points: [{ x: 0, y: 0 }, { x: 15.39, y: 0 }, { x: 15.39, y: 6.89 }, { x: 12.53, y: 6.89 },
             { x: 12.53, y: 9.2 }, { x: 0, y: 9.2 }, { close: true }],
  });
  tx(s, '❝', at(2.6, 3.9, 2.6, 2.4), { color: PALE, fontSize: 140, bold: true, fontFace: HEAD });
  tx(s, 'Far and away the best prize that life has to offer is the chance to work hard at work worth doing”',
     at(3.57, 5.2, 12.77, 4.95), { fontSize: 72, bold: true, color: DARK });
});

/* 6 — About us with overlapping photos and a floating icon card */
slide(s => {
  pic(s, 1.27, 1.57, 7.2, 9.72);
  pic(s, 4.68, 7.54, 8.23, 5.89);
  round(s, at(9.46, 6.07, 2.45, 2.71), 0.16, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.14, blur: 18, offset: 4, angle: 90 } });
  icon(s, 'doc', 10.07, 6.81, 1.35, GREEN, ORANGE);
  kicker(s, 'About Us', 14.35, 1.75, 2.33);
  title(s, 'We are looking a new fresh talent to join outr team', 14.35, 2.29, 11.06, 4.14);
  para(s, LOREM_LONG, 14.35, 7.02, 11.06, 3.06);
  button(s, 'GET STARTED', 14.35, 11.01);
});

/* 7 — "Offering Vacancies" + stat card */
slide(s => {
  pic(s, 0, 6.18, 26.66, 8.82);
  kicker(s, 'About Us', 1.81, 1.86, 2.33);
  title(s, 'Offering Vacancies From Famous Companies', 1.81, 2.55, 15.01, 2.79);
  round(s, at(18.52, 2.29, 6.16, 6.14), 0.16, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 22, offset: 5, angle: 90 } });
  icon(s, 'people', 20.84, 3.47, 1.8, GREEN, ORANGE);
  tx(s, '864.058', at(19.42, 5.41, 4.35, 1.45), { align: 'center', fontSize: 80, bold: true, color: GREEN });
  tx(s, 'Active Workers', at(19.63, 7.01, 3.93, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
});

/* 8 — three numbered steps beside stacked photos */
slide(s => {
  pic(s, 0, 0, 9.73, 12.85);
  pic(s, 4.81, 6.26, 7.63, 8.74);
  round(s, at(13.92, 7.98, 11.4, 3.53), 0.1, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 22, offset: 4, angle: 90 } });
  kicker(s, 'About Us', 14.48, 1.17, 2.33);
  title(s, 'We connect  job vacancy and companies', 14.48, 1.86, 9.73, 4.14);
  [['01', 6.68, 'Get your job posting & company info  out to the wold', 6.74, 9.67],
   ['02', 8.44, 'Search for the best candidates for you request', 8.5, 8.85],
   ['03', 12.1, 'Provide our professional with resources and full support', 12.16, 8.85]]
    .forEach(([n, oy, label, ly, lw]) => {
      oval(s, at(14.48, oy, 0.79, 0.79), { fill: { color: GREEN } });
      tx(s, n, at(14.48, oy + 0.14, 0.79, 0.51), { align: 'center', bold: true, color: WHITE });
      tx(s, label, at(15.65, ly, lw, 0.57), { fontSize: 28, bold: true, italic: true, color: DARK });
    });
  para(s, LOREM_SHORT, 15.65, 9.2, 8.95, 1.85);
});

/* 9 — "Connect people" + 4-step white card */
slide(s => {
  pic(s, 14.47, -0.03, 12.19, 15);
  kicker(s, 'About Us', 1.51, 1.45, 2.33);
  title(s, 'Connect people with the right jobs for them', 1.51, 2.15, 10.74, 4.14);
  para(s, 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices. Donec euismod scelerisque ligula. ',
       1.51, 6.54, 11.39, 1.85);
  round(s, at(1.51, 9.23, 18.13, 4.32), 0.1, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 22, offset: 4, angle: 90 } });
  [['Register', 2.51, 3.33, 'people'], ['Search', 7.01, 7.83, 'doc'],
   ['Apply', 11.58, 12.41, 'plane'], ['Succeed', 16.16, 16.98, 'award']]
    .forEach(([t, tx0, ix, ic], i) => {
      icon(s, ic, ix, 10.45, 1.1, GREEN, ORANGE);
      tx(s, t, at(tx0, 11.97, 2.58, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
      if (i) rule(s, [6.01, 10.59, 15.15][i - 1], 10.16, 0, 2.45, HAIR, 1);
    });
});

/* 10 — "How It Works", three circled icons */
slide(s => {
  kicker(s, 'Working Process', 12.17, 2.01, 2.33, { align: 'center' });
  title(s, 'How It Works', 7.8, 2.5, 11.06, 1.45, { align: 'center' });
  para(s, 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices. Donec euismod scelerisque ligula',
       3.52, 4.11, 19.62, 1.25, { align: 'center' });
  [['Create an Account', 4.62, 3.47, 3.16, 'award', 5.5],
   ['Search Jobs', 12.04, 10.88, 10.57, 'doc', 12.88],
   ['Save & Apply', 19.46, 18.3, 17.99, 'briefcase', 20.32]]
    .forEach(([t, ox, hx, px, ic, ix]) => {
      oval(s, at(ox, 7.21, 2.58, 2.56), { fill: { color: PALE } });
      icon(s, ic, ix, 8.05, 1.05, GREEN, ORANGE);
      tx(s, t, at(hx, 10.46, 4.89, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
      para(s, 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec', px, 11.2, 5.51, 1.25, { align: 'center' });
    });
});

/* 11 — "Our Leaders", three portraits */
slide(s => {
  kicker(s, 'Team', 1.65, 1.37, 2.33);
  title(s, 'Our Leaders', 1.65, 2.07, 9.1, 1.45);
  para(s, LOREM_MED, 12.85, 2.07, 11.88, 1.85);
  [['Mikkel Svane', 'CEO.Founder', 1.65], ['Morten Primdahi', 'Founder Jobcy', 9.74],
   ['Alexander aghaspure', 'Chief People', 17.84]].forEach(([n, role, x]) => {
    pic(s, x, 4.84, 6.69, 6.24);
    rule(s, x, 12.03, 6.69, 0, LGRAY, 1);
    tx(s, n, at(x, 12.38, 5.0, 0.64), { fontSize: 32, bold: true, color: DARK });
    tx(s, role, at(x, 13.23, 3.1, 0.51));
  });
});

/* 12 — "Our Amazing Team", two member cards */
slide(s => {
  [['Morten Primdahi', 'Founder Jobcy', 1.75, 2.77, 3.66], ['Samanta Smith', 'Acounting', 9.57, 10.59, 11.48]]
    .forEach(([n, role, cx, nx, sx]) => {
      s.addShape('custGeom', {
        x: cx, y: 1.74, w: 6.93, h: 10.78, fill: { color: WHITE },
        shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 22, offset: 4, angle: 90 },
        points: [{ x: 0, y: 1.07 }, { x: 6.93, y: 6.49, curve: { type: 'cubic', x1: 0, y1: 4.08, x2: 2.79, y2: 6.52 } },
                 { x: 6.93, y: 10.78 }, { x: 0, y: 10.78 }, { close: true }],
      });
      pic(s, cx, 1.74, 6.93, 6.52);
      tx(s, n, at(nx, 8.88, 4.89, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
      tx(s, role, at(nx + 0.89, 9.73, 3.1, 0.51), { align: 'center' });
      socials(s, sx, 10.77);
    });
  kicker(s, 'Team', 18.33, 1.75, 2.33);
  title(s, 'Our Amazing Team', 18.33, 2.45, 5.96, 4.14);
  para(s, LOREM_MED, 18.33, 6.88, 6.58, 3.06);
  button(s, 'GET STARTED', 18.33, 10.76);
});

/* 13 — profile with skill bars */
slide(s => {
  s.addShape('custGeom', {
    x: 9.32, y: 0, w: 17.34, h: 6.05, fill: { color: MINT },
    points: [{ x: 17.34, y: 0 }, { x: 17.34, y: 5.2 },
             { x: 0, y: 0, curve: { type: 'cubic', x1: 17.34, y1: 5.2, x2: 3.89, y2: 9.15 } }, { close: true }],
  });
  s.addShape('custGeom', {
    x: 0, y: 0, w: 19.38, h: 3.21, fill: { color: GREEN },
    points: [{ x: 0, y: 2.54 },
             { x: 11.14, y: 2.08, curve: { type: 'cubic', x1: 0, y1: 2.54, x2: 6.08, y2: 4.31 } },
             { x: 19.38, y: 0, curve: { type: 'cubic', x1: 16.19, y1: -0.15, x2: 19.38, y2: 0 } },
             { x: 0, y: 0 }, { close: true }],
  });
  logo(s, 0.96, 0.81, 1.27, WHITE, WHITE);
  pic(s, 14.77, 0, 11.87, 15);
  kicker(s, 'Team', 1.46, 4.39, 2.33);
  title(s, 'Horland Benjamin', 1.46, 5.08, 10.86, 1.45);
  para(s, LOREM_MED, 1.46, 6.64, 11.38, 1.85);
  [['Software Company', '90%', 9.22, 10.19, 10.55], [' Financial Services', '80%', 11.19, 12.16, 8.86]]
    .forEach(([label, pct, ty, by, bw]) => {
      tx(s, label, at(1.46, ty, 4.89, 0.64), { fontSize: 32, bold: true, color: DARK });
      tx(s, pct, at(12.01, ty, 1.32, 0.64), { align: 'right', fontSize: 32, bold: true, color: DARK });
      rect(s, at(1.46, by, 11.87, 0.35), { fill: { color: HAIR } });
      rect(s, at(1.46, by + 0.1, bw, 0.16), { fill: { color: GREEN } });
    });
});

/* 14 — "Job Categories" list card */
slide(s => {
  rect(s, at(0, 0, 6.89, 8.3), { fill: { color: GREEN } });
  logoVertical(s, 0.18, 0.6, 1.0);
  pic(s, 2.55, 1.37, 12.51, 9.13);
  rect(s, at(12.18, 6.59, 13.23, 7.31), { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 24, offset: 5, angle: 90 } });
  kicker(s, 'Service', 16.3, 1.37, 2.33);
  title(s, 'Job Categories', 16.3, 1.9, 9.1, 1.45);
  para(s, LOREM_MED, 16.3, 3.61, 9.1, 1.85);
  para(s, LOREM_SHORT + '.', 2.55, 11.33, 8.61, 1.85);
  [['Accounting & finance', '120 job vacancy'], ['Software development', '100 job vacancy'],
   ['Business development', '18 job vacancy'], ['Human resource', '1 job vacancy'],
   ['Marketing', '78 job vacancy']].forEach(([name, n], i) => {
    const y = 7.73 + i * 1.065;
    oval(s, at(13.32, y + 0.23, 0.17, 0.17), { fill: { color: GREEN } });
    tx(s, name, at(13.75, y, 6.0, 0.64), { fontSize: 32, bold: true, color: DARK });
    tx(s, n, at(20.37, y + 0.07, 4.1, 0.51), { align: 'right' });
  });
});

/* 15 — "Experience Categories": 2x4 grid of cards */
slide(s => {
  kicker(s, 'Service', 11.76, 1.55, 3.13, { align: 'center' });
  title(s, 'Experience Categories', 6.12, 2.18, 14.42, 1.45, { align: 'center' });
  const CARDS = [
    ['Project Management', '74 offer', 'blueprint'], ['Cost Estimating', '579 offer', 'money'],
    ['Diaster Recovery', '60 offer', 'server'], ['Drafting Engineering', '120 offer', 'pen'],
    ['Advertising Business', '124 offer', 'chart'], ['Document Control', '24 offer', 'folder'],
    ['Cost Analysis', '387 offer', 'board'],
  ];
  CARDS.forEach(([name, offer, ic], i) => {
    const x = 1.33 + (i % 4) * 6.12, y = 4.99 + Math.floor(i / 4) * 4.58;
    round(s, at(x, y, 5.6, 4.06), 0.13, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 20, offset: 4, angle: 90 } });
    icon(s, ic, x + 0.93, y + 0.91, 1.1, GREEN, ORANGE);
    tx(s, offer, at(x + 2.5, y + 0.74, 2.4, 0.5), { align: 'right' });
    tx(s, name, at(x + 0.93, y + 2.18, 3.59, 1.18), { fontSize: 32, bold: true, color: DARK });
  });
  round(s, at(19.7, 9.57, 5.6, 4.06), 0.13, { fill: { color: GREEN } });
  tx(s, '+', at(20.62, 10.0, 0.9, 0.9), { color: WHITE, fontSize: 54, bold: true });
  tx(s, 'We have Much Category', at(20.62, 12.03, 2.86, 0.91), { color: WHITE });
});

/* 16 — "Popular Categories" with a vertical card list */
slide(s => {
  rect(s, at(0, 0, 10.81, 15), { fill: { color: GREEN } });
  for (let r = 0; r < 7; r++) for (let c = 0; c < 6; c++)   // decorative dot grid on the green panel
    oval(s, at(1.95 + c * 0.265, 0.77 + r * 0.222, 0.11, 0.11), { fill: { color: GREEN_D } });
  pic(s, 2.5, 1.25, 8.31, 12.5);
  rect(s, at(10.81, 1.25, 5.04, 12.5), { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 24, offset: 5, angle: 90 } });
  [['Design & Art', '(6 open positions)', 3.53, 2.17, 'layers'],
   ['Teaching', '(3 open positions)', 7.4, 6.15, 'teacher'],
   ['Banking', '(1 open positions)', 11.27, 10.03, 'bank']].forEach(([t, sub, ty, iy, ic]) => {
    icon(s, ic, 12.86, iy, 1.05, GREEN, ORANGE);
    tx(s, t, at(11.47, ty, 3.71, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
    tx(s, sub, at(11.72, ty + 0.77, 3.23, 0.51), { align: 'center' });
  });
  kicker(s, 'Service', 17.45, 2.85, 2.33);
  title(s, 'Popular Categories', 17.45, 3.48, 7.27, 2.79);
  para(s, LOREM_LONG, 17.45, 6.32, 7.86, 3.06);
  button(s, 'GET STARTED', 17.45, 10.16);
});

/* 17 — CV / resume layout */
slide(s => {
  rect(s, at(1.06, 0, 8.78, 14.99), { fill: { color: ICE } });
  rect(s, at(0, 0, 1.06, 15), { fill: { color: GREEN } });
  rect(s, at(26.36, 0, 0.3, 15), { fill: { color: GREEN } });
  logoVertical(s, 0.22, 0.55, 0.85);
  shape(s, 'donut', at(2.76, 1.47, 5.38, 5.38), { fill: { color: GREEN } });
  pic(s, 3.06, 1.77, 4.79, 4.79, { size: 28 });
  tx(s, 'JAMES BILLY', at(2.52, 7.97, 5.87, 1.01), { align: 'center', fontFace: HEAD, fontSize: 54, bold: true, color: DARK });
  rule(s, 1.97, 9.1, 6.97, 0, SILVER, 1);
  tx(s, 'Graphics Designer', at(3.06, 9.53, 4.79, 0.64), { align: 'center', fontSize: 32 });
  socials(s, 3.9, 11.33);

  heading(s, 'About Me', 11.95, 1.03, 6.15);
  para(s, LOREM_LONG, 11.95, 1.67, 12.7, 1.85);
  heading(s, 'Basic Steps', 11.95, 4.06, 6.15);
  [['1. ', 'Sed sollicitudin ipsum quis nunc sollicitudin ultrices.'],
   ['2. ', 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec,.'],
   ['3. ', 'Donec viverra eleifend lacus, vitae ullamcorper metus']].forEach(([n, t], i) =>
    tx(s, [{ text: n, options: { color: GREEN } }, { text: t }], at(11.95, 4.81 + i * 0.645, 9.42, 0.5)));
  heading(s, 'Education', 11.95, 7.25, 6.15);
  [['Masters In Fine Arts ', '2002 - 2004'], ['Tombers Collage ', '2012 - 2015'],
   ['Diploma In Fine Arts ', '2014 - 2015']].forEach(([t, yr], i) => {
    oval(s, at(11.95, 8.17 + i * 0.695, 0.2, 0.2), { fill: { color: GREEN } });
    tx(s, [{ text: t }, { text: yr, options: { color: GREEN } }], at(12.36, 8.02 + i * 0.695, 9.42, 0.5));
  });
  heading(s, 'Work & Experience', 11.95, 10.57, 6.15);
  [['Web Designer 06/06/2017 - 06/06/2018', 'TBC Studio', 11.47],
   ['Ceo Founder 06/04/2018 - 06/04/2019', 'Wiggle CRC', 12.91]].forEach(([t, org, y], i) => {
    oval(s, at(11.95, y, 0.51, 0.51), { fill: { type: 'none' }, line: { color: GREY2, width: 1 } });
    tx(s, t, at(12.69, y, 9.42, 0.5));
    tx(s, org, at(12.69, y + 0.62, 2.58, 0.5), { color: GREEN });
    if (!i) rule(s, 12.21, 11.98, 0, 0.93, GREY2, 1);
  });
});

/* 18 — candidate profile sheet */
slide(s => {
  pic(s, 0, 0, 26.66, 7.2);
  wash(s, 0, 0, 26.66, 7.2, WHITE, 80, 0);
  oval(s, at(2.1, 3.04, 3.09, 3.09), { fill: { color: GREEN } });
  pic(s, 2.3, 3.24, 2.69, 2.69, { size: 24 });
  title(s, 'Candidate', 6.2, 3.04, 9.11, 1.45);
  tx(s, 'Start Using Wireshark to Hack', at(6.2, 4.73, 9.42, 0.5));
  [['pin', 6.2, '115 Rue de Reuilly Paris', 6.49, 4.02],
   ['phone', 10.75, '+015 123 456 789', 11.16, 3.19],
   ['mail', 14.57, 'candidate@jobcy.com', 14.99, 4.02]].forEach(([ic, ix, t, tx0, w]) => {
    icon(s, ic, ix, 5.68, 0.5, GREEN, ORANGE);
    tx(s, t, at(tx0, 5.64, w, 0.5));
  });
  heading(s, 'Qualification:', 2.1, 8.31, 6.15);
  para(s, LOREM_LONG, 2.1, 9.16, 12.48, 1.85);
  heading(s, 'Job Overview', 16.41, 8.29, 6.15);
  [['Age:', '25-30', 2.1, 11.64], ['Gender:', 'Both', 2.1, 12.96]].forEach(([k, v, x, y]) => {
    tx(s, k, at(x, y, 1.93, 0.5), { color: DARK });
    tx(s, v, at(x + 6.47, y - 0.05, 1.56, 0.5), { align: 'right' });
    rule(s, x, y + 0.78, 8.03, 0, LGRAY, 1);
  });
  [['Categories:', 'Accounting / Finance', 9.51], ['Offered Salary:', '$5000 / month', 10.66],
   ['Experience:', '3 Years', 11.81], ['Languages:', 'English, Turkish, Japanese', 12.96]]
    .forEach(([k, v, y]) => {
      tx(s, k, at(16.41, y, 5.03, 0.5), { color: DARK });
      tx(s, v, at(20.8, y - 0.05, 3.63, 0.91), { align: 'right' });
      rule(s, 16.41, y + 0.78, 8.03, 0, LGRAY, 1);
    });
});

/* 19 — "How to get the Job?" with green swoosh */
slide(s => {
  pic(s, 0, 0, 14.17, 15);
  s.addShape('custGeom', {
    x: 0, y: 4.54, w: 16.76, h: 10.46, fill: { color: GREEN },
    points: [{ x: 0, y: 2.17 },
             { x: 9.63, y: 6.83, curve: { type: 'cubic', x1: 0, y1: 2.17, x2: 5.26, y2: -1.13 } },
             { x: 16.76, y: 10.45, curve: { type: 'cubic', x1: 14.0, y1: 10.61, x2: 16.76, y2: 10.45 } },
             { x: 0, y: 10.46 }, { close: true }],
  });
  round(s, at(1.36, 2.68, 4.64, 5.86), 0.22, { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.14, blur: 20, offset: 4, angle: 90 } });
  pic(s, 2.81, 3.13, 1.73, 1.73, { size: 14, marker: false });
  tx(s, 'M.Marshall', at(1.9, 5.28, 3.57, 0.64), { align: 'center', fontSize: 32, bold: true, color: DARK });
  tx(s, 'Accounting / Finance', at(1.86, 6.1, 3.63, 0.5), { align: 'center' });
  button(s, 'HIRE ME', 2.56, 7.09, 2.25, 0.78, GREEN, WHITE);
  round(s, at(11.67, 7.68, 1.41, 1.47), 0.16, { fill: { color: ORANGE } });
  icon(s, 'people', 12.06, 8.1, 0.8, WHITE, WHITE);
  kicker(s, 'Special Slides', 14.85, 2.07, 2.33);
  title(s, 'How to get the Job ?', 14.85, 2.68, 10.07, 2.79);
  para(s, 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices.',
       14.85, 5.9, 10.07, 1.85);
  ['Improve your skill', 'Send the Proposal', 'Interview'].forEach((t, i) => {
    const y = 8.49 + i * 1.465;
    oval(s, at(14.85, y, 0.95, 0.95), { fill: { color: GREEN } });
    tx(s, String(i + 1), at(14.98, y + 0.19, 0.7, 0.64), { align: 'center', fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
    tx(s, t, at(16.02, y + 0.16, 4.65, 0.64), { fontFace: HEAD, fontSize: 32, bold: true, color: DARK });
  });
});

/* 20 — "Most job applicants at jobcy" */
slide(s => {
  rect(s, at(18.37, 0, 8.29, 13.34), { fill: { color: GREEN } });
  pic(s, 16.79, 1.66, 8.29, 13.34);
  pic(s, 1.58, 8.19, 6.36, 5.16);
  pic(s, 9.13, 8.19, 6.36, 5.16);
  rect(s, at(9.13, 8.19, 6.36, 5.16), { fill: { color: DARK, transparency: 30 } });
  kicker(s, 'Portfolio', 1.58, 1.9, 2.33);
  title(s, [{ text: 'Most job applicants at ' }, { text: 'jobcy' }], 1.58, 2.53, 11.98, 2.79);
  para(s, LOREM_LONG, 1.58, 5.52, 14.02, 1.85);
  tx(s, '250', at(10.57, 9.15, 3.4, 2.04), { align: 'center', fontSize: 115, bold: true, color: WHITE });
  tx(s, 'Interactive Designer', at(10.35, 11.29, 3.83, 1.18), { align: 'center', fontSize: 32, bold: true, color: WHITE });
  rect(s, at(14.18, 6.85, 0.97, 2.67), { fill: { color: WHITE }, rotate: 270, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 14, offset: 3, angle: 90 } });
  tx(s, 'Job Vacancy', at(13.56, 7.94, 2.21, 0.5), { align: 'right' });
});

/* 21 — "Our Portfolio Jobcy" */
slide(s => {
  pic(s, 0, 0, 8.13, 15);
  pic(s, 8.12, 0, 7.38, 7.5);
  pic(s, 8.12, 7.5, 7.38, 7.5);
  kicker(s, 'Portfolio', 16.99, 1.71, 2.33);
  title(s, [{ text: 'Our Portfolio ' }, { text: 'Jobcy' }], 16.99, 2.34, 8.17, 2.79);
  para(s, 'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, vitae ullamcorper metus sed sollicitudin ipsum.',
       16.99, 5.33, 8.6, 1.85);
  [['Project Management', 7.95, 9.24, 'blueprint'], ['Cost Estimating', 11.15, 12.44, 'money']]
    .forEach(([t, ty, py, ic]) => {
      icon(s, ic, 16.99, ty + 0.03, 1.05, GREEN, ORANGE);
      tx(s, t, at(18.39, ty, 3.59, 1.18), { fontSize: 32, bold: true, color: DARK });
      para(s, LOREM_TINY, 16.99, py, 8.35, 1.25);
    });
});

/* 22 — "Our Gallery Jobcy" with two stat cards */
slide(s => {
  pic(s, 0, 0, 10.32, 15);
  wash(s, 0, 0, 10.32, 15, '000000', 88, 0);
  kicker(s, 'Portfolio', 0.78, 8.17, 2.33, { color: WHITE });
  title(s, [{ text: 'Our Gallery ' }, { text: 'Jobcy' }], 0.78, 8.8, 7.08, 2.79, { color: WHITE });
  para(s, LOREM_SHORT, 0.78, 11.78, 7.08, 2.46, { color: LGRAY });
  [['864.058', 'Active Workers', 'Fusce vehicula dolor arcu, sit amet blandit.', 11.71, 8.1, 5.48, 10.01],
   ['120.509', 'Companies', 'Fusce vehicula dolor arcu, sit amet blandit ollis nec', 19.03, 1.71, 5.05, 2.4]]
    .forEach(([n, h, body, cx, cy, ch, ny]) => {
      rect(s, at(cx, cy, 6.29, ch), { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 20, offset: 4, angle: 90 } });
      tx(s, n, at(cx + 0.62, ny, 3.24, 0.84), { fontFace: HEAD, fontSize: 44, bold: true, color: GREEN });
      tx(s, h, at(cx + 0.62, ny + 0.9, 4.53, 0.64), { fontSize: 32, bold: true, color: DARK });
      para(s, body, cx + 0.62, ny + 1.61, 5.11, 1.25);
    });
  pic(s, 11.71, 1.71, 6.29, 7.58);        // photo slots overlap the cards from above
  pic(s, 19.03, 6.0, 6.29, 7.58);
});

/* 23 — clustered bar chart */
slide(s => {
  logo(s, 1.49, 1.17);
  s.addChart('bar', [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 1.8, 2.8] },
  ], {
    x: 1.49, y: 3.22, w: 12.54, h: 10.55, barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [GREEN, ORANGE], showLegend: false, showValue: false,
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 20, catAxisLabelColor: '595959',
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 20, valAxisLabelColor: '595959',
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineShow: true, valAxisLineShow: false, border: { pt: 0, color: 'FFFFFF' },
  });
  kicker(s, 'Charts', 16.05, 4.41, 2.33);
  title(s, [{ text: 'Chart ' }, { text: 'jobcy' }, { text: ' analysis slide' }], 16.05, 4.94, 8.17, 2.79);
  heading(s, 'Report User', 16.05, 8.31, 6.57);
  para(s, LOREM_LONG + 'Maecenas eu varius risus, eu aliquet arcu. ', 16.05, 9.17, 9.12, 3.06);
});

/* 24 — 3-D pie chart on a white card */
slide(s => {
  rect(s, at(1.26, 1.98, 24.14, 11.67), { fill: { color: WHITE }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 26, offset: 5, angle: 90 } });
  logo(s, 12.21, 0.74);
  s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] }], {
    x: 13.24, y: 3.2, w: 11.55, h: 9.87, chartColors: [GREEN, ORANGE, GREEN2, PALE],
    showLegend: true, legendPos: 'b', legendFontFace: BODY, legendFontSize: 18,
    showPercent: false, dataBorder: { pt: 0, color: 'FFFFFF' },
  });
  kicker(s, 'Charts', 2.43, 3.2, 2.33);
  title(s, [{ text: 'Chart ' }, { text: 'jobcy' }, { text: ' analysis slide' }], 2.43, 4.08, 8.17, 2.79);
  para(s, LOREM_SHORT, 2.43, 7.18, 10.2, 1.85);
  [['1200+', 'Visitor', 2.43, 2.61], ['780', 'User', 7.78, 2.1]].forEach(([n, l, x, lw]) => {
    tx(s, n, at(x, 9.84, 2.61, 1.11), { fontSize: 60, bold: true, color: GREEN });
    tx(s, l, at(x, 11.21, lw, 0.64), { fontSize: 32, bold: true, color: DARK });
  });
});

/* 25 — 2x2 ribbon infographic */
slide(s => {
  kicker(s, 'Infographics', 12.17, 1.42, 2.33, { align: 'center' });
  title(s, [{ text: 'jobcy' }, { text: ' Infographics slide' }], 5.37, 1.86, 15.92, 1.45, { align: 'center' });
  para(s, LOREM_SHORT, 4.47, 3.57, 17.71, 1.25, { align: 'center' });
  [['01', 'Project Management', GREEN, GREEN_D, 3.49, 6.01, 'blueprint'],
   ['02', 'Advertising Business', ORANGE, ORANGE_D, 14.0, 6.01, 'chart'],
   ['03', 'Document Control', GREEN2, GREEN_M, 3.49, 10.64, 'folder'],
   ['04', 'Cost Analysis', OLIVE, OLIVE_D, 14.0, 10.64, 'board']]
    .forEach(([n, label, c, shade, x, y, ic]) => {
      bookmark(s, x + 7.27, y + 0.82, 1.21, 2.44, shade);   // dark tail, mostly hidden behind the plate
      s.addShape('custGeom', {              // grey plate, rounded on the top-right
        x, y, w: 9.17, h: 2.32, fill: { color: EEE },
        points: [{ x: 0, y: 0 }, { x: 0, y: 2.32 }, { x: 9.17, y: 2.32 }, { x: 9.17, y: 0.97 },
                 { x: 8.2, y: 0, curve: { type: 'cubic', x1: 9.17, y1: 0.43, x2: 8.74, y2: 0 } }, { close: true }],
      });
      bookmark(s, x, y, 1.61, 3.27, c);                     // bright tail under the number
      tx(s, n, at(x + 0.18, y + 0.43, 1.26, 0.84), { align: 'center', fontSize: 44, bold: true, color: WHITE });
      tx(s, label, at(x + 2.29, y + 0.44, 5.1, 0.64), { fontSize: 32, bold: true, color: DARK });
      tx(s, 'Fusce vehicula dolor arcu, ', at(x + 2.29, y + 1.1, 4.43, 0.5));
      icon(s, ic, x + 7.51, y + 0.44, 1.05, GREEN, ORANGE);
    });
});

/* 26 — arrow-band infographic beside a text column */
slide(s => {
  kicker(s, 'Infographics', 1.83, 2.58, 2.33);
  title(s, [{ text: 'Infographics  ' }, { text: 'jobcy' }, { text: ' analysis slide' }], 1.83, 3.33, 9.27, 4.14);
  para(s, LOREM_LONG + 'Maecenas eu varius risus, eu aliquet arcu. ', 1.83, 7.54, 9.12, 3.06);
  button(s, 'READ MORE', 1.83, 11.38, 3.45, 0.92, GREEN, WHITE, false);
  [['01', 'Project Management', GREEN, GREEN_D, 2.32, 'blueprint'],
   ['02', 'Document Control', ORANGE, ORANGE_D, 4.99, 'folder'],
   ['03', 'Advertising Business', GREEN2, GREEN_M, 7.66, 'chart'],
   ['04', 'Cost Analysis', OLIVE, OLIVE_D, 10.34, 'board']]
    .forEach(([n, label, c, shade, y, ic]) => {
      s.addShape('custGeom', {           // grey slanted band
        x: 14.53, y: y + 0.26, w: 9.97, h: 1.77, fill: { color: EEE },
        points: [{ x: 0.72, y: 0 }, { x: 0, y: 1.25 }, { x: 0.17, y: 1.77 }, { x: 9.0, y: 1.77 },
                 { x: 9.97, y: 0 }, { close: true }],
      });
      s.addShape('custGeom', {           // coloured outline of the band
        x: 15.17, y: y + 0.26, w: 9.33, h: 1.77, fill: { color: c },
        points: [{ x: 0.08, y: 0 }, { x: 0, y: 0.13 }, { x: 9.04, y: 0.13 }, { x: 8.16, y: 1.77 },
                 { x: 8.36, y: 1.77 }, { x: 9.33, y: 0 }, { close: true }],
      });
      s.addShape('custGeom', {           // coloured underline wedge
        x: 14.0, y: y + 1.5, w: 5.8, h: 0.77, fill: { color: c },
        points: [{ x: 0.48, y: 0 }, { x: 0, y: 0.77 }, { x: 5.34, y: 0.77 }, { x: 5.8, y: 0 }, { close: true }],
      });
      s.addShape('custGeom', {           // pentagon badge (flat bottom, peaked top)
        x: 12.88, y, w: 1.81, h: 2.27, fill: { color: c },
        points: [{ x: 0.9, y: 0 }, { x: 0, y: 0.52 }, { x: 0, y: 2.27 }, { x: 1.81, y: 2.27 },
                 { x: 1.81, y: 0.52 }, { close: true }],
      });
      s.addShape('custGeom', {           // fold shadow under the pentagon
        x: 13.41, y: y + 1.18, w: 1.28, h: 1.1, fill: { color: shade },
        points: [{ x: 1.28, y: 0 }, { x: 0, y: 1.1 }, { x: 1.28, y: 1.1 }, { close: true }],
      });
      tx(s, n, at(13.15, y + 0.6, 1.26, 0.84), { align: 'center', fontSize: 44, bold: true, color: WHITE });
      tx(s, label, at(15.54, y + 0.64, 6.57, 0.64), { fontSize: 32, bold: true, color: DARK });
      icon(s, ic, 21.66, y + 0.72, 1.05, GREEN, ORANGE);
    });
});

/* 27 — "Break slide" divider */
slide(s => {
  pic(s, 1.96, 2.78, 12.73, 9.88);
  rect(s, at(11.17, 6.56, 15.49, 8.44), { fill: { color: GREEN } });
  logo(s, 1.96, 0.84);
  nav(s, 0.9, GRAY, false, [['Job Alerts', 10.7, 2.15], ['Post Job', 13.98, 1.94], ['Job Dashboard', 17.06, 2.54]]);
  registerChip(s, 22.31, 0.75, 2.64, 0.8, GRAY);
  title(s, 'Break slide', 12.76, 7.82, 11.59, 2.42, { fontSize: 138, color: WHITE });
  tx(s, 'Job Board Presentation Template', at(12.76, 10.43, 8.49, 0.64), { fontSize: 32, bold: true, color: WHITE });
  para(s, LOREM_MED, 12.76, 11.24, 11.59, 1.85, { color: LGRAY });
});

/* 28-31 — SWOT quartet (same skeleton, mirrored left/right) */
const SWOT = [
  ['Strength ', 1.73, 14.34, [['Al Job Recommendation', 8.57], ['Varieties of job openings', 11.02]], 2.37, 3.05, 6.03],
  ['Weakness ', 14.38, 1.66, [['Inadequate filters', 8.57], ['Job Recommendations', 11.02]], 2.37, 3.05, 6.03],
  ['Opportunity ', 1.73, 14.29, [['Social Media', 9.09], ['Application Process', 11.53]], 1.63, 2.31, 6.42],
  ['Threat ', 14.7, 1.97, [['Social Platform', 8.57], ['Happyer', 11.02]], 2.37, 3.05, 6.03],
];
SWOT.forEach(([word, tx0, picX, items, ky, ty, py]) => slide(s => {
  pic(s, picX, 1.49, 10.32, 12.01);
  kicker(s, 'Swot', tx0, ky, 2.33);
  title(s, [{ text: word }, { text: 'Jobcy' }, { text: ' Analysis Slide' }], tx0, ty, 10.5, ty > 2.4 ? 2.79 : 4.14);
  para(s, LOREM_MED, tx0, py, 10.4, 1.85);
  items.forEach(([label, y]) => {
    icon(s, 'check', tx0, y, 1, GREEN);
    tx(s, label, at(tx0 + 0.84, y, 6.42, 0.64), { fontSize: 32, bold: true, color: DARK });
    para(s, LOREM_TINY, tx0 + 0.84, y + 0.68, 9.6, 1.25);
  });
}));

/* 32 — phone mockups */
slide(s => {
  kicker(s, 'Mockup', 1.5, 2.48, 2.33);
  title(s, 'Get Easy Order with Our Apps', 1.5, 3.16, 10.45, 2.79);
  heading(s, 'Good Design', 1.5, 6.47, 4.53);
  para(s, LOREM_SHORT, 1.5, 7.29, 10.45, 1.85);
  [['450+', 1.5], ['98k', 6.72]].forEach(([n, x]) => {
    tx(s, n, at(x, 9.8, 2.54, 1.21), { fontSize: 66, bold: true, color: GREEN });
    para(s, 'Fusce vehicula dolor arcu, sit amet blandit', x, 11.01, 4.42, 1.25);
  });
  [[17.4, 7.79], [24.76, 7.2]].forEach(([cx, cy]) => phone(s, cx, cy, 5.59, 11.21, 30));
});

/* 33 — single phone mockup, "Easy register jobcy with our apps" */
slide(s => {
  rect(s, at(0, 0, 5.94, 15), { fill: { color: GREEN } });
  phone(s, 4.26, 7.5, 16.95, 8.45, 0, 0.3);
  kicker(s, 'Mockup', 14.55, 1.8, 2.33);
  title(s, [{ text: 'Easy register ' }, { text: 'jobcy' }, { text: ' with our apps' }], 14.55, 2.49, 10.44, 4.14);
  para(s, LOREM_LONG + 'Maecenas eu varius risus, eu aliquet arcu. Curabitur fermentum suscipit est, tincidunt mattis lorem luctus id. ',
       14.55, 7.06, 10.44, 3.06);
  button(s, 'WWW.JOBCY.COM', 14.55, 11.01, 4.03, 0.92, GREEN, WHITE, false);
});

/* 34 — laptop mockup, "Check Jobcy New Site" */
slide(s => {
  laptop(s, 14.6, 3.77, 9.4, 6.6);
  kicker(s, 'Mockup', 2.17, 2.24, 2.33);
  title(s, [{ text: 'Check ' }, { text: 'Jobcy' }, { text: ' New Site' }], 2.17, 2.93, 8.28, 2.79);
  para(s, LOREM_MED, 2.17, 5.97, 9.57, 1.85);
  [['Good Design', 8.42, 'pen'], ['Best Quality', 10.81, 'bag']].forEach(([t, y, ic]) => {
    icon(s, ic, 2.17, y, 0.95, GREEN, ORANGE);
    heading(s, t, 3.47, y, 4.53);
    para(s, 'fusce vehicula dolor arcu, sit amet blandit dolor mollis nec donec viverra.', 3.47, y + 0.56, 7.79, 1.25);
  });
});

/* 35 — "Our Best Offers" pricing table */
slide(s => {
  kicker(s, 'Pricing', 11.76, 1.28, 3.13, { align: 'center' });
  title(s, 'Our Best Offers', 8.32, 1.91, 10.02, 1.45, { align: 'center' });
  [['Standard', 1.77, false, 'trophy'], ['Premium', 9.89, true, 'people'], ['Enterprise', 18.01, false, 'gear']]
    .forEach(([plan, x, featured, ic]) => {
      if (featured) {
        s.addShape('custGeom', {
          x, y: 3.94, w: 6.88, h: 9.75, fill: { color: WHITE },
          shadow: { type: 'outer', color: '000000', opacity: 0.14, blur: 22, offset: 4, angle: 90 },
          points: [{ x: 0.31, y: 0 }, { x: 6.57, y: 0 },
                   { x: 6.88, y: 0.31, curve: { type: 'cubic', x1: 6.74, y1: 0, x2: 6.88, y2: 0.14 } },
                   { x: 6.88, y: 9.75 }, { x: 0, y: 9.75 }, { x: 0, y: 0.31 },
                   { x: 0.31, y: 0, curve: { type: 'cubic', x1: 0, y1: 0.14, x2: 0.14, y2: 0 } }, { close: true }],
        });
      } else {
        shape(s, 'round2SameRect', at(x, 3.94, 6.88, 9.75), { fill: { color: PALE }, rectRadius: 0.31 });
      }
      oval(s, at(x + 2.65, 4.53, 1.57, 1.57), { fill: { type: 'none' }, line: { color: LGRAY, width: 1 } });
      icon(s, ic, x + 3.09, 4.97, 0.8, GREEN, ORANGE);
      tx(s, plan, at(x + 1.87, 6.46, 3.13, 0.64), { align: 'center', fontSize: 32, bold: true, color: GREEN });
      tx(s, '$490', at(x + 1.5, 7.55, 1.74, 0.64), { align: 'center', fontSize: 32, strike: 'sngStrike' });
      tx(s, '$390', at(x + 2.96, 7.26, 2.42, 1.21), { align: 'center', fontSize: 66, color: DARK });
      [['10 Job Posting ', 9.07], ['Urgent Job', 10.22], ['Job displayed for 10 days', 11.37],
       ['Premium support 24/7', 12.52]].forEach(([t, y], i) => {
        tx(s, t, at(x + 1.24, y, 4.4, 0.51), { align: 'center' });
        if (i < 3) rule(s, x + 0.94, y + 0.77, 5.0, 0, LGRAY, 1);
      });
    });
  rect(s, at(9.28, 4.59, 2.64, 0.74), { fill: { color: GREEN } });
  tx(s, 'FEATURED', at(9.65, 4.72, 1.94, 0.5), { align: 'center', color: WHITE });
  shape(s, 'triangle', at(9.44, 5.17, 0.3, 0.63), { fill: { color: '8BD475' }, rotate: 90, flipV: true });
});

/* 36 — "Get In Touch" contact card */
slide(s => {
  pic(s, 0, 0, 26.66, 8.59);
  s.addShape('custGeom', {
    x: 2.04, y: 5.25, w: 22.84, h: 8.28, fill: { color: WHITE },
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 26, offset: 5, angle: 90 },
    points: [{ x: 0, y: 0 }, { x: 10.97, y: 0 }, { x: 10.97, y: 3.34 }, { x: 22.84, y: 3.34 },
             { x: 22.84, y: 8.28 }, { x: 0, y: 8.28 }, { close: true }],
  });
  kicker(s, 'Contact Us', 3.31, 6.8, 3.13);
  title(s, 'Get In Touch', 3.31, 7.43, 7.82, 1.45);
  [['Address', '2045 W Grand Ave Ste, Chicago', 3.31, 'building', 10.92, 11.81, 0.91],
   ['Phone Number', '(+68)1221 09876', 10.3, 'phone', 10.88, 11.77, 0.5],
   ['Email', 'contact@yoursite.com', 17.29, 'mail', 10.92, 11.81, 0.5]]
    .forEach(([h, v, x, ic, hy, vy, vh]) => {
      icon(s, ic, x, 9.52, 1.15, GREEN, ORANGE);
      tx(s, h, at(x, hy, 3.59, 0.64), { fontSize: 32, bold: true, color: DARK });
      tx(s, v, at(x, vy, 4.4, vh));
    });
});

/* 37 — "Icon Slide" divider */
slide(s => {
  pic(s, 0, 0, 26.66, 15);
  s.addShape('custGeom', {
    x: 0, y: 6.68, w: 18.03, h: 6.2, fill: { color: WHITE, transparency: 30 },
    points: [{ x: 0, y: 0 }, { x: 15.22, y: 0 },
             { x: 18.03, y: 3.1, curve: { type: 'cubic', x1: 16.77, y1: 0, x2: 18.03, y2: 1.39 } },
             { x: 15.22, y: 6.2, curve: { type: 'cubic', x1: 18.03, y1: 4.81, x2: 16.77, y2: 6.2 } },
             { x: 0, y: 6.2 }, { close: true }],
  });
  s.addShape('custGeom', {
    x: 0, y: 7.04, w: 17.62, h: 5.49, fill: { color: WHITE },
    shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 22, offset: 4, angle: 90 },
    points: [{ x: 0, y: 0 }, { x: 14.87, y: 0 },
             { x: 17.62, y: 2.75, curve: { type: 'cubic', x1: 16.39, y1: 0, x2: 17.62, y2: 1.23 } },
             { x: 14.87, y: 5.49, curve: { type: 'cubic', x1: 17.62, y1: 4.26, x2: 16.39, y2: 5.49 } },
             { x: 0, y: 5.49 }, { close: true }],
  });
  logo(s, 1.55, 7.73);
  title(s, [{ text: 'Icon ' }, { text: 'Slide', options: { color: GREEN } }], 1.56, 8.92, 14.51, 3.45, { fontSize: 199 });
});

/* 38-41 — glyph reference sheets: 11 x 6 grids of pictograms */
const GLYPH_SHEETS = [
  '☎▤✍⚑✉➤✎✒⌫▭↰|▲▤⇗ⓘ⚲☰⊙➣◉⚠⇘|☆✔✖ⓘ♥♦❞⌂❐⚲✥|⚐⚙⚒★☂☏☁☾❀♣♪|▦⚖✈⚓◉◷☝▤⚡▣◆|◍Ⓐ⚽✏♦⚔⚑▥➤✍❖',,
  '☎▥✍⚐✉➤✏✎⌦⬭↰|▼▤⇗ⓘ⚲☰⊙➤◎⚠⇘|☆✔✖ⓘ♣♠❝⌂⧉⚲✦|⚑⚙⚒✩☂☏☁☽❁♫♩|▤⚖✈⚓◉◷☞▩⚡▤◇|◉Ⓐ⚽✐♠⚕⚐▣▨➔✎',,
  '♿✕⊞✖✗¥▶▷▶⊕⊖|♦∪◁∨⚲◑Ⅴ⚠◐⚑◉|⇧☂⌃↺✳▤⋮⌄⇩↕♠|▣▯✤✥★☆➤▥➣∪▦|▤▧⌄⋯✖✔✚✜◷≡◆|⚙▶⌐✧✥☎▦▩☏⊤⊥',,
  '⌗▤✎▮℗℘▣▶▧✈▷|▦▤▥▨▭◇◆⚒◈⌫⌦|⊟▤Ⅿ▣⊞▯▬☾✥♫♪|⚲⚳☺✍◐⊟▤→⊖✦✧|⇪⚭❀⇧₽∈▤♦▥▤←|ⓘ▣◀⊕▤⚒◍Ⓐ✺⇧⚡',
];
GLYPH_SHEETS.forEach(sheet => slide(s => {
  sheet.split('|').forEach((row, r) => Array.from(row).forEach((g, c) => {
    tx(s, g, at(1.25 + c * 2.25, 1.95 + r * 2.02, 1.4, 1.2),
       { align: 'center', valign: 'middle', fontFace: 'DejaVu Sans', fontSize: 40, color: '000000' });
  }));
}));

/* 42 — "Thank You" closing slide */
slide(s => {
  pic(s, 0, 2.28, 26.66, 9.31);
  wash(s, 0, 2.27, 26.66, 9.34, WHITE, 69, 0);
  logo(s, 1.29, 0.84);
  nav(s, 0.9, GRAY, true, [['Job Alerts', 16.47, 2.15], ['Post Job', 19.75, 1.94], ['Job Dashboard', 22.83, 2.54]]);
  title(s, [{ text: 'Thank ' }, { text: 'You', options: { color: GREEN } }], 4.36, 5.78, 17.94, 3.45,
        { fontSize: 199, align: 'center' });
  para(s, LOREM_MED, 5.24, 12.38, 16.18, 1.25, { align: 'center' });
});

/* ------------------------------------------------------------------ build */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'JOBCY', width: 26.66, height: 15 });
pptx.layout = 'JOBCY';
pptx.author = 'Jobcy';
pptx.title = 'Job Board Presentation Template';
SLIDES.forEach(build => build(pptx.addSlide()));
pptx.writeFile({ fileName: path.join(__dirname, '0117beb1-62ed-47c5-be13-7a7f442b03d5_grok_final.pptx') })
  .then(f => console.log('wrote', f));
