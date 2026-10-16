/**
 * Harmony Yoga & Meditation deck - rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9).  Raster photos in the source deck are
 * re-created here as flat grey placeholder rectangles.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  dark:   '424530', // accent1 - headings, logo
  orange: 'E09132', // accent2 - highlights
  tan:    'A58E74', // accent3
  cream:  'FFEFCD', // accent4
  olive:  '51563F', // accent5
  white:  'FFFFFF',
  page:   'FFFBF5', // accent4 @20% over white = slide background
  rule:   'C9BBAC', // hair rules of the page frame
  grid:   'D9D9D9', // light separators
  grid2:  'BFBFBF', // table separators
  body:   '808080', // body copy
  body2:  '595959', // body copy on the photo slides
  ink:    '404040',
  photo:  'CCCCCC', // image placeholder fill
  label:  'AFAFAF', // "[image]" caption
  device: '2B2B2B'  // device mock-up bezels
};

const F_HEAD = 'Pridi';   // theme major font
const F_BODY = 'Manrope'; // theme minor font

// pptxgenjs rewrites the shadow object in place while serialising, so hand each
// shape its own copy.
const soft = () => ({ type: 'outer', angle: 90, blur: 14, offset: 2, color: '000000', opacity: 0.07 });

/* ------------------------------------------------------------------- copy */

const T = {
  full:  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore.',
  ab:    'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo.',
  hero:  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium totam aperiam, eaque ipsa quae ab illo.',
  short: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem.',
  mid:   'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium.',
  line:  'Sed ut perspiciatis unde omnis iste natus error sit',
  dolor: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque.',
  laud:  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium.',
  totam: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium totam.',
  essay: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, ' +
         'totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta ' +
         'sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia ' +
         'consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui ' +
         'dolorem ipsum quia dolor sit amet, consectetur, adipisci velit. But I must explain to you how all this.'
};

/* ---------------------------------------------------------------- helpers */

// Page title - 44pt Pridi in accent1 (the deck's standard section heading).
function heading(s, text, x, y, w, h, opt = {}) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: F_HEAD, fontSize: 44, color: C.dark, charSpacing: 0, valign: 'top'
  }, opt));
}

// Running body copy - 12pt Manrope, 150% leading.
function copy(s, text, x, y, w, h, opt = {}) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: F_BODY, fontSize: 12, color: C.body,
    lineSpacingMultiple: 1.5, valign: 'top'
  }, opt));
}

function label(s, text, x, y, w, h, opt = {}) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: F_BODY, fontSize: 12, color: C.body, valign: 'top'
  }, opt));
}

function hline(s, x, y, w, color, width) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width } });
}

function vline(s, x, y, h, color, width) {
  s.addShape('line', { x, y, w: 0, h, line: { color, width } });
}

// Free-form polygon given in absolute slide inches.
function quad(s, pts, fill, opt = {}) {
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const x = Math.min(...xs), y = Math.min(...ys);
  s.addShape('custGeom', Object.assign({
    x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y, fill: { color: fill },
    points: pts.map(p => ({ x: p[0] - x, y: p[1] - y })).concat([{ close: true }])
  }, opt));
}

// Shrink a polygon towards its centroid (used for screen insets).
function inset(pts, f) {
  const cx = pts.reduce((a, p) => a + p[0], 0) / pts.length;
  const cy = pts.reduce((a, p) => a + p[1], 0) / pts.length;
  return pts.map(p => [cx + (p[0] - cx) * f, cy + (p[1] - cy) * f]);
}

// Bilinear sub-quad of [topLeft, topRight, bottomRight, bottomLeft], u/v in 0..1.
function subQuad(q, u0, v0, u1, v1) {
  const at = (u, v) => [0, 1].map(k =>
    q[0][k] * (1 - u) * (1 - v) + q[1][k] * u * (1 - v) + q[2][k] * u * v + q[3][k] * (1 - u) * v);
  return [at(u0, v0), at(u1, v0), at(u1, v1), at(u0, v1)];
}

// Stand-in for a photograph: flat grey block, optionally captioned.
function photo(s, x, y, w, h, opt = {}) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.photo } });
  if (opt.caption !== false && w > 0.9 && h > 0.6) {
    s.addText('[image]', {
      x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle',
      fontFace: F_BODY, fontSize: 10, color: C.label
    });
  }
}

function photoCircle(s, x, y, d) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: C.photo } });
}

// Small white pictograms that sit inside the coloured badges of the original.
function glyph(s, kind, cx, cy, size, color = C.white) {
  const ln = { color, width: Math.max(1, size * 4) };
  const dot = (x, y, d) => s.addShape('ellipse', { x: x - d / 2, y: y - d / 2, w: d, h: d, fill: { color } });
  const box = (x, y, w, h, r) => s.addShape(r ? 'roundRect' : 'rect',
    Object.assign({ x: x - w / 2, y: y - h / 2, w, h, fill: { color } }, r ? { rectRadius: r } : {}));
  const outline = (shape, x, y, w, h, r) => s.addShape(shape,
    Object.assign({ x: x - w / 2, y: y - h / 2, w, h, fill: { type: 'none' }, line: ln }, r ? { rectRadius: r } : {}));
  const glyphText = (ch, font, scale, dy) => s.addText(ch, {
    x: cx - size, y: cy - size + (dy || 0) * size, w: size * 2, h: size * 2, margin: 0,
    align: 'center', valign: 'middle', fontFace: font, fontSize: size * scale, bold: true, color
  });
  switch (kind) {
    case 'person':   // upright figure, arms down
    case 'meditate': // lotus pose - the deck's signature pictogram
      dot(cx, cy - size * 0.32, size * 0.26);
      box(cx, cy + size * 0.02, size * 0.52, size * 0.34, size * 0.16);
      box(cx, cy + size * 0.28, size * 0.78, size * 0.12, size * 0.06);
      break;
    case 'warrior': // lunging figure
      dot(cx - size * 0.06, cy - size * 0.34, size * 0.22);
      s.addShape('line', { x: cx - size * 0.44, y: cy - size * 0.18, w: size * 0.72, h: 0, line: ln });
      s.addShape('line', { x: cx - size * 0.10, y: cy - size * 0.20, w: size * 0.02, h: size * 0.30, line: ln });
      s.addShape('line', { x: cx - size * 0.40, y: cy + size * 0.10, w: size * 0.32, h: size * 0.28,
        line: ln, flipV: true });
      s.addShape('line', { x: cx - size * 0.08, y: cy + size * 0.10, w: size * 0.34, h: size * 0.28, line: ln });
      s.addShape('line', { x: cx + size * 0.26, y: cy + size * 0.38, w: size * 0.14, h: 0, line: ln });
      break;
    case 'star': // standing figure, arms and legs spread
      dot(cx, cy - size * 0.34, size * 0.22);
      s.addShape('line', { x: cx - size * 0.42, y: cy - size * 0.14, w: size * 0.84, h: 0, line: ln });
      s.addShape('line', { x: cx, y: cy - size * 0.18, w: 0, h: size * 0.34, line: ln });
      s.addShape('line', { x: cx - size * 0.26, y: cy + size * 0.16, w: size * 0.26, h: size * 0.26,
        line: ln, flipV: true });
      s.addShape('line', { x: cx, y: cy + size * 0.16, w: size * 0.26, h: size * 0.26, line: ln });
      break;
    case 'dance': // figure mid-step with a music note
      dot(cx - size * 0.10, cy - size * 0.34, size * 0.22);
      s.addShape('line', { x: cx - size * 0.42, y: cy - size * 0.22, w: size * 0.56, h: size * 0.10,
        line: ln, flipV: true });
      s.addShape('line', { x: cx - size * 0.12, y: cy - size * 0.18, w: size * 0.04, h: size * 0.30, line: ln });
      s.addShape('line', { x: cx - size * 0.34, y: cy + size * 0.12, w: size * 0.24, h: size * 0.28,
        line: ln, flipV: true });
      s.addShape('line', { x: cx - size * 0.08, y: cy + size * 0.12, w: size * 0.22, h: size * 0.28, line: ln });
      s.addText('\u266B', { x: cx + size * 0.06, y: cy - size * 0.30, w: size * 0.56, h: size * 0.56,
        margin: 0, align: 'center', valign: 'middle', fontFace: F_BODY, fontSize: size * 26, color });
      break;
    case 'facebook':   glyphText('f', F_HEAD, 66, -0.02); break;
    case 'tiktok': // stylised eighth note
      box(cx + size * 0.04, cy - size * 0.06, size * 0.11, size * 0.56);
      s.addShape('ellipse', { x: cx - size * 0.22, y: cy + size * 0.10, w: size * 0.30, h: size * 0.24,
        fill: { color } });
      s.addShape('line', { x: cx + size * 0.10, y: cy - size * 0.32, w: size * 0.20, h: size * 0.18,
        line: { color, width: size * 8 } });
      break;
    case 'quote': // pair of solid comma marks
      [-0.24, 0.04].forEach(o => quad(s, [
        [cx + size * o, cy - size * 0.24], [cx + size * (o + 0.20), cy - size * 0.24],
        [cx + size * (o + 0.20), cy + size * 0.10], [cx + size * (o + 0.06), cy + size * 0.26],
        [cx + size * o, cy + size * 0.26]
      ], color));
      break;
    case 'instagram':
      outline('roundRect', cx, cy, size * 0.7, size * 0.7, size * 0.2);
      outline('ellipse', cx, cy, size * 0.34, size * 0.34);
      dot(cx + size * 0.22, cy - size * 0.22, size * 0.08);
      break;
    case 'sun':
      dot(cx, cy, size * 0.28);
      [0, 60, 120, 180, 240, 300].forEach(a => {
        const r = size * 0.38, t = a * Math.PI / 180;
        dot(cx + r * Math.cos(t), cy + r * Math.sin(t), size * 0.13);
      });
      break;
    case 'phone': // handset, tilted
      s.addShape('roundRect', { x: cx - size * 0.16, y: cy - size * 0.40, w: size * 0.32, h: size * 0.80,
        fill: { color }, rectRadius: size * 0.14, rotate: 330 });
      break;
    case 'web':
      outline('rect', cx, cy + size * 0.02, size * 0.86, size * 0.62);
      box(cx, cy - size * 0.16, size * 0.86, size * 0.10);
      box(cx - size * 0.14, cy + size * 0.10, size * 0.40, size * 0.08);
      break;
    case 'mail':
      outline('rect', cx, cy, size * 0.86, size * 0.62);
      s.addShape('line', { x: cx - size * 0.43, y: cy - size * 0.31, w: size * 0.43, h: size * 0.31, line: ln });
      s.addShape('line', { x: cx, y: cy, w: size * 0.43, h: size * 0.31, line: ln, flipV: true });
      break;
    case 'note': // clipboard with a pencil
      outline('rect', cx - size * 0.04, cy + size * 0.04, size * 0.58, size * 0.74);
      box(cx - size * 0.04, cy - size * 0.32, size * 0.26, size * 0.13, size * 0.04);
      [0.02, 0.16, 0.30].forEach(o => box(cx - size * 0.04, cy + size * o, size * 0.34, size * 0.06));
      s.addShape('line', { x: cx - size * 0.16, y: cy + size * 0.16, w: size * 0.46, h: size * 0.46,
        line: { color, width: size * 6 }, flipV: true });
      break;
    case 'group': // small crowd seen through a magnifier
      [-0.30, 0.30].forEach(o => {
        dot(cx + size * o, cy - size * 0.14, size * 0.20);
        box(cx + size * o, cy + size * 0.16, size * 0.36, size * 0.28, size * 0.10);
      });
      dot(cx - size * 0.02, cy - size * 0.20, size * 0.22);
      box(cx - size * 0.02, cy + size * 0.14, size * 0.40, size * 0.32, size * 0.12);
      outline('ellipse', cx + size * 0.02, cy - size * 0.06, size * 0.54, size * 0.54);
      s.addShape('line', { x: cx + size * 0.24, y: cy + size * 0.18, w: size * 0.20, h: size * 0.20,
        line: { color, width: size * 7 } });
      break;
    case 'coins': // two overlapping stacks of coins
      [[-0.14, -0.20], [0.14, 0.10]].forEach(([ox, oy]) => {
        [0.20, 0.04, -0.12].forEach(o => s.addShape('ellipse',
          { x: cx + size * ox - size * 0.30, y: cy + size * (oy + o) - size * 0.09,
            w: size * 0.60, h: size * 0.20, fill: { color }, line: { color: C.orange, width: size * 3 } }));
      });
      break;
  }
}

/* -------------------------------------------------- shared page furniture */

// Logo lock-up in the top-left corner: line-art meditating figure + wordmark.
function logo(s) {
  const ring = { color: C.orange, width: 1 };
  const oval = (x, y, w, h) => s.addShape('ellipse', { x, y, w, h, fill: { type: 'none' }, line: ring });
  oval(0.777, 0.444, 0.056, 0.056);           // head
  s.addShape('line', { x: 0.717, y: 0.512, w: 0.088, h: 0.086, line: ring, flipH: true }); // left arm
  s.addShape('line', { x: 0.805, y: 0.512, w: 0.088, h: 0.086, line: ring });              // right arm
  oval(0.625, 0.597, 0.186, 0.140);           // left wing
  oval(0.786, 0.597, 0.186, 0.140);           // right wing
  oval(0.762, 0.583, 0.088, 0.167);           // folded legs
  [0.700, 0.900].forEach(x => s.addShape('ellipse',
    { x, y: 0.528, w: 0.018, h: 0.018, fill: { color: C.orange } }));
  s.addText('Harmony', { x: 1.028, y: 0.400, w: 1.10, h: 0.30, margin: 0, wrap: false,
    fontFace: F_HEAD, fontSize: 17, color: C.dark });
  s.addText('yoga and meditation', { x: 1.418, y: 0.664, w: 0.60, h: 0.07, margin: 0, wrap: false,
    fontFace: F_BODY, fontSize: 3.5, color: C.olive });
}

// Every slide starts from the master: cream page + hair-line frame + logo + pill.
function page(pres) {
  const s = pres.addSlide();
  s.background = { color: C.page };
  return s;
}

// The hair-line frame plus the "Contact Us" pill that the slide master paints.
function frame(s) {
  vline(s, 0.602, 0.863, 6.637, C.rule, 0.5);
  hline(s, -0.009, 6.899, 13.342, C.rule, 0.5);
  hline(s, 2.116, 0.601, 9.578, C.rule, 0.5);
  vline(s, 0.602, 0.000, 0.424, C.rule, 0.5);
  hline(s, 0.000, 0.601, 0.514, C.rule, 0.5);
  vline(s, 12.736, 0.921, 6.579, C.rule, 0.5);
  vline(s, 12.736, 0.000, 0.340, C.rule, 0.5);
  hline(s, 13.033, 0.601, 0.300, C.rule, 0.5);
  logo(s);
  s.addShape('roundRect', { x: 11.832, y: 0.475, w: 1.079, h: 0.286, fill: { color: C.dark }, rectRadius: 0.143 });
  s.addText('Contact Us', { x: 11.832, y: 0.475, w: 1.079, h: 0.286, align: 'center', valign: 'middle',
    margin: 0, fontFace: F_BODY, fontSize: 8, color: C.white, charSpacing: 1 });
}

// Facebook / TikTok / Instagram badges (cover + break + closing slides).
function socialRow(s, x, y) {
  const d = 0.413;
  [['facebook', C.dark], ['tiktok', C.orange], ['instagram', C.tan]].forEach(([kind, fill], i) => {
    const bx = x + i * 0.640, cx = bx + d / 2, cy = y + d / 2;
    s.addShape('ellipse', { x: bx, y, w: d, h: d, fill: { color: fill }, shadow: soft() });
    s.addShape('ellipse', { x: cx - 0.137, y: cy - 0.137, w: 0.274, h: 0.274,
      fill: { type: 'none' }, line: { color: C.white, width: 0.5, dashType: 'sysDot' } });
    glyph(s, kind, cx, cy, 0.2);
  });
}

/* ------------------------------------------------------------ slide 1 - cover */

function slideCover(pres) {
  const s = page(pres);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.photo } });
  frame(s);
  heading(s, 'Feel Your Body and Soul with Us', 1.292, 1.466, 6.903, 2.121, { fontSize: 60 });
  copy(s, T.hero, 1.292, 3.835, 4.571, 0.978, { color: C.body2 });
  socialRow(s, 1.412, 5.247);
}

/* ------------------------------------------------- slide 2 - table of content */

function slideAgenda(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Table of Content', 1.277, 1.519, 6.972, 0.841, { valign: 'bottom' });
  // two columns x four rows; numbering runs 1-4 down the left, 5-8 down the right
  [0, 1].forEach(col => {
    const tx = col ? 7.098 : 1.277, lx = col ? 7.207 : 1.386, bx = col ? 11.677 : 5.856;
    [0, 1, 2, 3].forEach(row => {
      const ty = 2.685 + row * 1.105;
      copy(s, T.line, tx, ty, 4.132, 0.379);
      s.addShape('ellipse', { x: bx, y: ty + 0.002, w: 0.379, h: 0.375, fill: { color: C.orange }, shadow: soft() });
      s.addText(String(col * 4 + row + 1), { x: bx, y: ty + 0.002, w: 0.379, h: 0.375, align: 'center',
        valign: 'middle', margin: 0, fontFace: F_BODY, fontSize: 10.5, color: C.white });
      if (row < 3) hline(s, lx, 3.427 + row * 1.105, 4.837, C.grid, 0.5);
    });
  });
}

/* ------------------------------------------------------ slide 3 - about company */

function slideAbout(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 1.764, 1.859, 3.063, 2.729);
  [1.859, 3.309, 4.758].forEach(y => photo(s, 4.981, y, 1.454, 1.279));
  photo(s, 1.764, 4.758, 1.454, 1.279);
  photo(s, 3.372, 4.758, 1.454, 1.279);
  heading(s, 'About Company', 7.283, 1.821, 4.287, 1.582);
  label(s, 'Harmony Yoga and Meditation', 7.283, 3.700, 4.287, 0.438,
    { fontFace: F_HEAD, fontSize: 20, color: C.tan });
  copy(s, T.full, 7.283, 4.202, 4.287, 0.978);
  s.addShape('roundRect', { x: 7.355, y: 5.662, w: 1.383, h: 0.416, fill: { color: C.orange }, rectRadius: 0.208 });
  s.addText('Read more!', { x: 7.355, y: 5.662, w: 1.383, h: 0.416, align: 'center', valign: 'middle',
    margin: 0, fontFace: F_BODY, fontSize: 10.5, color: C.white });
}

/* ------------------------------------------------------- slide 4 - performance */

function slidePerformance(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 6.000, 1.859, 3.519, 4.179);
  heading(s, 'Enhancing Performance and Recovery', 1.277, 2.139, 4.275, 2.322);
  copy(s, T.ab, 1.277, 4.779, 4.275, 0.978);
  [0, 1, 2].forEach(i => {
    const oy = 2.108 + i * 1.435, ox = i === 2 ? 9.212 : 9.249;
    s.addShape('ellipse', { x: ox, y: oy, w: 0.539, h: 0.539, fill: { color: C.orange }, shadow: soft() });
    s.addText(`${i + 1}.`, { x: ox + 0.045, y: oy + 0.084, w: 0.45, h: 0.37, align: 'center', margin: 0,
      fontFace: F_BODY, fontSize: 16, bold: true, color: C.white });
    copy(s, T.short, ox + 0.718, oy - 0.084, 2.09, 0.978);
    if (i < 2) hline(s, 10.093, 3.230 + i * 1.436, 1.944, C.grid, 0.5);
  });
}

/* ------------------------------------------------------- slide 5 - morning yoga */

function slideMorning(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 0.000, 1.475, 6.463, 4.946);
  vline(s, 0.602, 0.863, 6.637, C.rule, 0.5); // frame rule redrawn above the photo
  heading(s, 'The Ultimate Morning Yoga Routine', 7.032, 1.626, 4.884, 2.322);
  copy(s, T.full, 7.032, 4.235, 4.884, 0.978);
  label(s, '90%', 7.032, 5.680, 0.904, 0.505,
    { fontFace: F_HEAD, fontSize: 24, color: C.orange, align: 'center' });
  copy(s, T.dolor, 8.139, 5.594, 3.777, 0.675);
}

/* ----------------------------------------------------------- slide 6 - balance */

function slideBalance(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 9.830, 1.859, 2.227, 4.178);
  photo(s, 6.056, 1.859, 3.615, 3.051);
  heading(s, 'Building Stability and Balance', 1.277, 2.139, 4.275, 2.322);
  copy(s, T.ab, 1.277, 4.779, 4.275, 0.978);
  s.addShape('rect', { x: 6.056, y: 5.060, w: 3.615, h: 0.977, fill: { color: C.orange }, shadow: soft() });
  glyph(s, 'warrior', 6.441, 5.549, 0.36);
  copy(s, T.short, 6.743, 5.211, 2.782, 0.675, { color: C.white });
}

/* -------------------------------------------------------- slide 7 - techniques */

function slideTechniques(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 8.098, 1.519, 3.959, 2.092);
  heading(s, 'Techniques for Focus', 1.277, 1.774, 6.671, 0.841);
  copy(s, T.full, 1.277, 2.680, 6.671, 0.675);
  [0, 1, 2, 3].forEach(i => {
    const x = 1.301 + i * 2.725;
    s.addShape('rect', { x, y: 4.074, w: 2.556, h: 2.329, fill: { color: C.white }, shadow: soft() });
    label(s, `Techniques ${i + 1}`, x + 0.243, 4.424, 2.069, 0.337,
      { fontFace: F_HEAD, fontSize: 14, color: C.dark });
    copy(s, T.short, x + 0.243, 4.803, 2.069, 0.978);
    hline(s, x + 0.366, 6.053, 0.25, C.orange, 0.5);
  });
}

/* ----------------------------------------------------- slide 8 - strengthening */

function slideStrength(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 1.764, 1.859, 4.671, 4.179);
  heading(s, 'Strengthening Your Body Naturally', 7.032, 2.100, 4.884, 2.322);
  s.addShape('ellipse', { x: 7.143, y: 4.905, w: 0.805, h: 0.805, fill: { color: C.orange }, shadow: soft() });
  glyph(s, 'warrior', 7.545, 5.307, 0.38);
  copy(s, T.totam, 8.199, 4.818, 3.717, 0.978);
}

/* ---------------------------------------------------------- slide 9 - benefits */

function slideBenefits(pres) {
  const s = page(pres);
  frame(s);
  photo(s, 6.326, 2.934, 2.776, 3.051);
  photo(s, 9.280, 2.934, 2.776, 3.051);
  heading(s, 'Benefits and Precautions', 1.277, 1.672, 10.780, 0.841);
  [['90%', 'Benefit', 2.694, 3.297], ['10%', 'Precautions', 4.643, 5.246]].forEach(([pct, name, hy, by]) => {
    label(s, pct, 1.277, hy, 0.904, 0.505, { fontFace: F_HEAD, fontSize: 24, color: C.orange, align: 'center' });
    label(s, name, 2.374, hy, 3.506, 0.505, { fontFace: F_HEAD, fontSize: 24, color: C.dark });
    copy(s, T.full, 1.277, by, 4.604, 0.978);
  });
}

/* -------------------------------------------------------- slide 10 - pranayama */

function slidePranayama(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Pranayama Techniques in Yoga', 1.226, 1.709, 10.830, 0.841);
  copy(s, T.essay, 1.108, 2.747, 5.414, 2.493, { color: '000000', align: 'justify' });
  copy(s, T.essay, 6.812, 2.747, 5.414, 2.493, { color: '000000', align: 'justify' });
  ['30 min', '8k+', '30 min', '8k+', '30 min'].forEach((big, i) => {
    const x = 1.108 + i * 2.374;
    label(s, big, x, 5.581, 1.503, 0.370,
      { fontFace: F_HEAD, fontSize: 16, color: C.orange, align: 'center' });
    label(s, 'Lorem ipsum', x + 0.119, 5.936, 1.265, 0.286,
      { fontSize: 11, color: C.body, align: 'center' });
    if (i) vline(s, x - 0.436, 5.646, 0.493, C.grid, 0.5);
  });
}

/* ------------------------------------------------------ slide 11 - break slide */

function slideBreak(pres) {
  const s = page(pres);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.photo } });
  frame(s);
  heading(s, 'Break Slide', 5.128, 2.554, 6.774, 1.582, { fontSize: 88 });
  copy(s, T.hero, 5.128, 4.015, 6.774, 0.675, { color: C.body2 });
  socialRow(s, 5.211, 5.101);
}

/* ----------------------------------------------------------- slide 12 - gallery */

function slideGallery(pres) {
  const s = page(pres);
  frame(s);
  [[5.053, 1.757, 3.232], [8.481, 1.757, 3.489], [1.364, 4.098, 3.489],
   [5.053, 4.098, 3.232], [8.481, 4.098, 3.489]].forEach(([x, y, w]) => photo(s, x, y, w, 2.161));
  heading(s, 'Our Gallery Activities ', 1.287, 1.697, 3.566, 1.582);
  copy(s, T.short, 1.287, 3.228, 3.566, 0.675);
}

/* ------------------------------------------------------- slide 13 - instructors */

function slideInstructors(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Meet Our Best Instructor', 1.287, 1.827, 10.631, 0.841, { align: 'center' });
  const names = ['Zenith', 'Cynthia', 'Daniel', 'Wedgas', 'Nola', 'Bams'];
  names.forEach((name, i) => {
    const x = 1.364 + (i % 3) * 3.613, y = 2.930 + Math.floor(i / 3) * 1.690;
    s.addShape('rect', { x, y, w: 3.379, h: 1.450, fill: { color: C.white }, shadow: soft() });
    photoCircle(s, x + 0.281, y + 0.237, 0.976);
    label(s, name, x + 1.446, y + 0.344, 1.652, 0.404, { fontFace: F_HEAD, fontSize: 18, color: C.tan });
    s.addShape('roundRect', { x: x + 1.505, y: y + 0.837, w: 0.938, h: 0.269,
      fill: { color: C.orange }, rectRadius: 0.134 });
    s.addText('More Info', { x: x + 1.505, y: y + 0.837, w: 0.938, h: 0.269, align: 'center', valign: 'middle',
      margin: 0, fontFace: F_BODY, fontSize: 8, color: C.white });
  });
}

/* --------------------------------------------- slide 14 - arrow-banner infographic */

// One banner: a tall block whose right edge carries a triangular point.
// Odd banners point near the top, even banners near the bottom.
function banner(s, x, fill, pointAtTop) {
  const w = 3.004, h = 3.744, body = 2.533;
  const t = pointAtTop ? 0.107 : 0.893;
  s.addShape('custGeom', {
    x, y: 2.692, w, h, fill: { color: fill }, shadow: soft(),
    points: [
      { x: 0, y: 0 }, { x: body, y: 0 },
      { x: body, y: h * (t - 0.0732) }, { x: w, y: h * t }, { x: body, y: h * (t + 0.0732) },
      { x: body, y: h }, { x: 0, y: h }, { close: true }
    ]
  });
}

function slideBanners(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Harmony Infographic', 1.226, 1.460, 10.830, 0.841);
  [3, 2, 1, 0].forEach(i => banner(s, 1.364 + i * 2.534, i % 2 ? C.cream : C.orange, i % 2 === 0));
  ['meditate', 'warrior', 'star', 'dance'].forEach((icon, i) => {
    const x = 1.728 + i * 2.539;
    s.addShape('ellipse', { x, y: 3.280, w: 0.976, h: 0.977, fill: { color: C.dark }, shadow: soft() });
    glyph(s, icon, x + 0.488, 3.769, 0.42);
    copy(s, T.mid, x - 0.086, 4.567, 1.982, 1.281, { color: i % 2 ? C.body2 : C.white });
  });
}

/* -------------------------------------------------- slide 15 - device mock-ups */

function slideDevices(pres) {
  const s = page(pres);
  frame(s);
  // tablet
  s.addShape('roundRect', { x: 1.639, y: 1.347, w: 3.480, h: 4.806,
    fill: { color: C.device }, rectRadius: 0.20, shadow: soft() });
  photo(s, 1.826, 1.663, 3.130, 4.190);
  // phone, overlapping the tablet on the right
  s.addShape('roundRect', { x: 4.425, y: 2.160, w: 1.514, h: 3.180,
    fill: { color: C.device }, rectRadius: 0.18, shadow: soft() });
  photo(s, 4.472, 2.233, 1.417, 3.035, { caption: false });
  heading(s, 'Connecting Body and Mind', 7.065, 2.043, 4.657, 1.582);
  copy(s, T.full, 7.065, 3.817, 4.657, 0.978);
  s.addShape('ellipse', { x: 7.155, y: 5.211, w: 0.608, h: 0.608, fill: { color: C.orange }, shadow: soft() });
  glyph(s, 'sun', 7.459, 5.515, 0.30);
  copy(s, T.dolor, 7.954, 5.178, 3.769, 0.675);
}

/* ------------------------------------------------------ slide 16 - testimonials */

function slideTestimonials(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Let\u2019s Know Our Customer\u2019s Opinion ', 1.226, 1.501, 10.830, 0.841);
  [['Rahayu Lestari', 1.241, 3.933, 4.414, 4.529], ['John Andrean', 6.789, 9.481, 9.963, 10.077]]
    .forEach(([name, px, ox, tx, lx]) => {
      photo(s, px, 2.664, 2.928, 3.731);
      s.addShape('ellipse', { x: ox, y: 3.101, w: 0.472, h: 0.472, fill: { color: C.orange }, shadow: soft() });
      glyph(s, 'quote', ox + 0.236, 3.337, 0.20);
      copy(s, T.laud, tx, 3.531, 2.130, 1.281);
      vline(s, lx, 4.884, 0.459, C.grid, 0.5);
      label(s, name, tx, 5.416, 2.130, 0.337, { fontFace: F_HEAD, fontSize: 14, color: '989E74' });
      label(s, 'Position', tx, 5.689, 2.130, 0.269, { fontSize: 10, italic: true, color: C.body });
    });
}

/* ------------------------------------------------------ slide 17 - schedule grid */

function slideSchedule(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Harmony Schedule Infographic', 1.226, 1.489, 10.830, 0.841);
  const cols = [[1.444, 1.704], [3.435, 3.089], [6.809, 3.089], [10.185, 1.704]];
  ['Time', 'Activity', 'Notes', 'Status'].forEach((head, i) => {
    label(s, head, cols[i][0], 2.621, cols[i][1], 0.370,
      { fontFace: F_HEAD, fontSize: 16, color: C.dark, align: 'center' });
  });
  for (let row = 0; row < 5; row++) {
    cols.forEach(([x, w]) => label(s, 'Lorem ipsum', x, 3.339 + row * 0.667, w, 0.269,
      { fontSize: 10, color: C.body, valign: 'middle' }));
  }
  [3.293, 6.667, 10.041].forEach(x => vline(s, x, 2.542, 3.875, C.grid2, 0.5));
  for (let r = 0; r < 5; r++) hline(s, 1.364, 3.139 + r * 0.668, 10.593, C.grid2, 0.5);
}

/* --------------------------------------------------------- slide 18 - spiritual */

function slideSpiritual(pres) {
  const s = page(pres);
  frame(s);
  // laptop seen in three-quarter perspective: keyboard deck, then the lid
  const deck = [[2.26, 4.49], [5.38, 3.45], [7.10, 4.32], [4.20, 6.30]];
  const lid  = [[4.45, 1.28], [1.33, 2.32], [2.26, 4.49], [5.38, 3.45]];
  quad(s, deck, 'DCDCDC', { shadow: soft() });
  quad(s, subQuad(deck, 0.06, 0.08, 0.88, 0.60), '727272');  // keyboard
  quad(s, subQuad(deck, 0.36, 0.68, 0.68, 0.92), 'CBCBCB');  // trackpad
  quad(s, lid, C.device);
  quad(s, inset(lid, 0.955), C.photo);            // screen = image placeholder
  s.addText('[image]', { x: 2.20, y: 2.72, w: 2.30, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F_BODY, fontSize: 10, color: C.label });
  heading(s, 'The Spiritual Side of Yoga', 7.788, 2.150, 4.320, 1.582);
  copy(s, T.full, 7.788, 3.923, 4.320, 0.978);
  label(s, '75%', 7.889, 5.284, 0.654, 0.303, { bold: true, color: C.ink });
  label(s, 'Body movement', 9.335, 5.284, 2.552, 0.303, { bold: true, color: C.body, align: 'right' });
  hline(s, 7.889, 5.746, 3.999, 'E6E6E6', 2.5);
  hline(s, 8.520, 5.746, 3.367, C.orange, 2.5);
}

/* ------------------------------------------------- slide 19 - circle infographic */

function slideCircles(pres) {
  const s = page(pres);
  frame(s);
  heading(s, 'Harmony Infographic', 1.226, 1.705, 10.830, 0.841);
  // circle, its glyph, and the text block that sits opposite it
  const cells = [
    { cx: 1.210, cy: 3.163, fill: C.dark,   icon: 'note',   tx: 1.226, ty: 4.946 },
    { cx: 4.121, cy: 4.750, fill: C.orange, icon: 'coins',  tx: 4.108, ty: 3.159 },
    { cx: 7.050, cy: 3.163, fill: C.tan,    icon: 'group',  tx: 6.990, ty: 4.946 },
    { cx: 9.979, cy: 4.750, fill: C.olive,  icon: 'star',   tx: 9.872, ty: 3.159 }
  ];
  cells.forEach(c => {
    s.addShape('ellipse', { x: c.cx, y: c.cy, w: 1.257, h: 1.258, fill: { color: c.fill }, shadow: soft() });
    glyph(s, c.icon, c.cx + 0.629, c.cy + 0.629, 0.52);
    label(s, 'INFOGRAPHIC 1', c.tx, c.ty, 2.235, 0.303, { color: C.orange, charSpacing: 2 });
    copy(s, T.mid, c.tx, c.ty + 0.266, 2.235, 0.978);
  });
}

/* ------------------------------------------------------------ slide 20 - thanks */

function slideThanks(pres) {
  const s = page(pres);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.photo } });
  frame(s);
  heading(s, 'Thanks For Attention', 1.292, 1.494, 4.805, 2.121, { fontSize: 60, align: 'right' });
  copy(s, T.laud, 1.176, 3.688, 4.571, 0.675, { color: C.body2, align: 'right' });
  label(s, 'More Information', 7.667, 2.762, 4.444, 0.505, { fontFace: F_HEAD, fontSize: 24, color: C.dark });
  [['phone', '123 - 4567890', 7.835, 3.578, 8.572, 3.678],
   ['web', 'www.harmonyyoga.com', 8.299, 4.462, 9.036, 4.562],
   ['mail', 'info@harmonyyoga.com', 8.796, 5.346, 9.533, 5.446]].forEach(([icon, text, ox, oy, tx, ty]) => {
    s.addShape('ellipse', { x: ox, y: oy, w: 0.536, h: 0.537, fill: { color: C.orange }, shadow: soft() });
    glyph(s, icon, ox + 0.268, oy + 0.268, 0.26);
    label(s, text, tx, ty, 2.624, 0.337, { fontSize: 14, color: C.body2, valign: 'middle' });
  });
}

/* -------------------------------------------------------------------- build */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'HARMONY', width: 13.333, height: 7.5 });
  pres.layout = 'HARMONY';
  pres.author = 'Harmony Yoga';
  pres.title = 'Harmony - Yoga and Meditation';
  [slideCover, slideAgenda, slideAbout, slidePerformance, slideMorning, slideBalance,
   slideTechniques, slideStrength, slideBenefits, slidePranayama, slideBreak, slideGallery,
   slideInstructors, slideBanners, slideDevices, slideTestimonials, slideSchedule,
   slideSpiritual, slideCircles, slideThanks].forEach(fn => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '044edffb-03c2-4a36-b2c6-20bb31464b6b_grok_final.pptx')
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
