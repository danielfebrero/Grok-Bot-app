// Recreation of "Panorama - Photography Presentation Template" (40 slides, 16:9)
// Generated with pptxgenjs only; raster photos are replaced by flat colour placeholders.
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const RED         = '760000';
const RED_DK      = 'A20000';
const RED_BRIGHT  = 'D60000';
const RED_MID     = 'B40000';
const RED_DEEP    = '920000';
const RED_SHADOW  = '420000';
const RED_9A      = '9A0000';
const RED_C8      = 'C80000';
const RED_VIVID   = 'FF0101';
const RED_LIGHT   = 'FF3B3B';
const INK         = '262626';
const INK_60      = '404040';
const INK_45      = '595959';
const GREY        = '808080';
const GREY_LT     = 'A6A6A6';
const SILVER      = 'BFBFBF';
const SILVER_LT   = 'D9D9D9';
const SNOW        = 'F2F2F2';
const WHITE       = 'FFFFFF';
const MIST        = 'ECECEC';
const PAPER       = 'F9F9F9';
const GREY_97     = '979797';
const GREY_AD     = 'ADADAD';
const GREY_D3     = 'D3D3D3';

// ---------------------------------------------------------------- fonts
const SERIF  = 'Playfair Display';
const SANS   = 'Lato';
const UI     = 'Poppins Medium';
const UI_SB  = 'Poppins SemiBold';
const UI_LT  = 'Poppins Light';
const UI_RG  = 'Poppins';

const PHOTO = 'E4E0DD';        // stand-in tone for a real photograph
const EMPTY_PHOTO = 'FCFCFC';  // empty picture frame (unfilled in the template)

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pptx.layout = 'W16x9';
pptx.author = 'Panorama Photography Studio';
pptx.title = 'Photography Presentation Template';

// ---------------------------------------------------------------- helpers

// Every body copy block in the deck is a prefix of this single filler paragraph.
const LOREM = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. " +
  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an " +
  "unknown printer took a galley of type and scrambled it to make a type specimen book. " +
  "It has survived not only five centuries, but also the leap into electronic typesetting, " +
  "remaining essentially unchanged. It was popularised in the 1960s with the release of " +
  "Letraset sheets containing Lorem Ipsum passages, and more recently with desktop " +
  "publishing software like Aldus PageMaker including versions of Lorem Ipsum.";

/** First `words` words of LOREM, closed with a period (and a trailing space by default). */
function lorem(words, trailingSpace) {
  const t = LOREM.split(' ').slice(0, words).join(' ').replace(/[,.]$/, '') + '.';
  return trailingSpace === false ? t : t + ' ';
}

/** Flat rectangle stand-in for a photograph, optionally clipped to a polygon. */
function photo(s, x, y, w, h, o) {
  o = o || {};
  const box = { x: x, y: y, w: w, h: h };
  if (o.points) s.addShape('custGeom', Object.assign({}, box, { fill: { color: o.color || PHOTO }, points: o.points, rotate: o.rotate, line: o.line }));
  else if (o.shape) s.addShape(o.shape, Object.assign({}, box, { fill: { color: o.color || PHOTO }, rotate: o.rotate, line: o.line }));
  else s.addShape('rect', Object.assign({}, box, { fill: { color: o.color || PHOTO }, rotate: o.rotate, line: o.line }));
  if (o.label && w > 0.9 && h > 0.6) {
    s.addText('[image]', { x: x, y: y + h / 2 - 0.18, w: w, h: 0.36, align: 'center', valign: 'middle',
      fontFace: UI, fontSize: 11, color: 'B0A6A0' });
  }
}

/** Two thin strokes forming an "L" bracket corner ornament. */
function corner(s, x, y, w, h, flipH, flipV, color, width) {
  const line = { color: color, width: width || 1 };
  const hx = flipH ? x : x;                    // horizontal arm spans full width
  const vx = flipH ? x : x + w;                // vertical arm sits on one side
  const hy = flipV ? y + h : y;                // horizontal arm on top or bottom
  s.addShape('line', { x: hx, y: hy, w: w, h: 0, line: line });
  s.addShape('line', { x: vx, y: y, w: 0, h: h, line: line });
}

/**
 * Minimal line-art pictograms built from primitives. The template uses freeform
 * icon artwork; these keep the same silhouette and scale without embedded paths.
 */
function icon(s, kind, x, y, size, color) {
  const line = { color: color, width: 1.25 };
  const cx = x + size / 2, cy = y + size / 2;
  if (kind === 'aperture') {
    s.addShape('ellipse', { x: x, y: y, w: size, h: size, fill: { type: 'none' }, line: { color: color, width: 1.5 } });
    const r1 = size * 0.13, r2 = size * 0.42;
    for (let i = 0; i < 6; i++) {
      const a = (i * 60 + 20) * Math.PI / 180;
      const dx = Math.cos(a), dy = Math.sin(a);
      s.addShape('line', { x: cx + r1 * dx, y: cy + r1 * dy, w: (r2 - r1) * dx, h: (r2 - r1) * dy, line: line });
    }
    s.addShape('ellipse', { x: cx - r1, y: cy - r1, w: r1 * 2, h: r1 * 2, fill: { type: 'none' }, line: line });
  } else if (kind === 'tripod') {
    s.addShape('rect', { x: x + size * 0.12, y: y + size * 0.2, w: size * 0.76, h: size * 0.3, fill: { type: 'none' }, line: line });
    s.addShape('ellipse', { x: cx - size * 0.08, y: y + size * 0.29, w: size * 0.16, h: size * 0.14, fill: { type: 'none' }, line: line });
    s.addShape('line', { x: cx - size * 0.12, y: y + size * 0.12, w: size * 0.12, h: size * 0.08, line: line });
    s.addShape('line', { x: cx, y: y + size * 0.12, w: size * 0.12, h: -size * 0.08, line: line });
    s.addShape('line', { x: cx - size * 0.3, y: y + size * 0.92, w: size * 0.3, h: -size * 0.42, line: line });
    s.addShape('line', { x: cx, y: y + size * 0.5, w: size * 0.3, h: size * 0.42, line: line });
    s.addShape('line', { x: cx, y: y + size * 0.5, w: 0, h: size * 0.42, line: line });
  } else if (kind === 'camera') {
    s.addShape('rect', { x: x + size * 0.14, y: y + size * 0.4, w: size * 0.72, h: size * 0.42, fill: { type: 'none' }, line: line });
    s.addShape('ellipse', { x: cx - size * 0.14, y: cy - size * 0.05, w: size * 0.28, h: size * 0.28, fill: { type: 'none' }, line: line });
    s.addShape('rect', { x: x + size * 0.5, y: y + size * 0.22, w: size * 0.14, h: size * 0.18, fill: { type: 'none' }, line: line });
    s.addShape('rect', { x: x + size * 0.14, y: y + size * 0.14, w: size * 0.12, h: size * 0.18, fill: { type: 'none' }, line: line });
    s.addShape('line', { x: x + size * 0.2, y: y + size * 0.32, w: 0, h: size * 0.08, line: line });
  } else if (kind === 'umbrella') {
    s.addShape('pie', { x: x + size * 0.06, y: y + size * 0.14, w: size * 0.88, h: size * 0.72,
      angleRange: [180, 0], fill: { type: 'none' }, line: line });
    s.addShape('line', { x: cx, y: y + size * 0.5, w: 0, h: size * 0.38, line: line });
    s.addShape('line', { x: cx - size * 0.22, y: y + size * 0.98, w: size * 0.22, h: -size * 0.12, line: line });
    s.addShape('line', { x: cx, y: y + size * 0.86, w: size * 0.22, h: size * 0.12, line: line });
  }
}

/**
 * Stand-in for the device mock-up photos. Only the bezel is drawn (a thick dark
 * outline) so whatever is behind shows through the "screen", matching the
 * cut-out PNG mock-ups of the original deck.
 */
function device(s, x, y, w, h, o) {
  o = o || {};
  const r = o.radius === undefined ? Math.min(w, h) * 0.1 : o.radius;
  const b = o.bezel === undefined ? Math.min(w, h) * 0.035 : o.bezel;
  s.addShape('roundRect', { x: x + b / 2, y: y + b / 2, w: w - b, h: h - b, rectRadius: r, rotate: o.rotate,
    fill: { type: 'none' }, line: { color: '2E2E2E', width: b * 72 } });
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: r + b / 2, rotate: o.rotate,
    fill: { type: 'none' }, line: { color: '9C9C9C', width: 1.25 } });
}

/**
 * Striped sphere illustration (slide 33): ellipse bands whose width follows a
 * circle, each backed by a soft grey shadow band offset down and to the right.
 */
function globe(s, cx, cy, r) {
  const bands = [
    { t: -0.95, color: RED_SHADOW },
    { t: -0.72, color: RED_MID },
    { t: -0.40, color: RED_BRIGHT },
    { t: -0.06, color: RED_VIVID },
    { t: 0.31, color: RED_LIGHT },
    { t: 0.64, color: RED_VIVID },
    { t: 0.86, color: RED_BRIGHT },
    { t: 0.99, color: RED_MID }
  ];
  const bh = r * 0.185;
  bands.forEach(function (b) {
    const half = Math.sqrt(Math.max(0.02, 1 - b.t * b.t)) * r;
    const by = cy + b.t * r - bh / 2;
    s.addShape('ellipse', { x: cx - half + 0.09, y: by + 0.13, w: half * 2, h: bh, fill: { color: SILVER_LT } });
    s.addShape('ellipse', { x: cx - half, y: by, w: half * 2, h: bh, fill: { color: b.color } });
  });
}

/** The four keyword tabs ("Innovative / Creative / Interactive / Professional"). */
const TAGS = [
  { label: 'Innovative', w: 1.019, dx: 0.0, dy: 0.0 },
  { label: 'Creative', w: 0.866, dx: 1.126, dy: 1.05 },
  { label: 'Interactive', w: 1.042, dx: 2.176, dy: 2.187 },
  { label: 'Professional', w: 1.142, dx: 3.314, dy: 3.375 }
];

function tagRow(s, x, y, color) {
  TAGS.forEach(function (t) {
    s.addText(t.label, { x: x + t.dx, y: y, w: t.w, h: 0.286, valign: 'top', wrap: false,
      align: 'center', fontSize: 11, fontFace: UI, color: color });
  });
}

/** Same tabs rotated 90 degrees, stacked down a vertical rail centred on `cx`. */
function tagColumn(s, cx, y, color, diamonds) {
  TAGS.forEach(function (t) {
    s.addText(t.label, { x: cx - t.w / 2, y: y + t.dy, w: t.w, h: 0.286, rotate: 90,
      valign: 'top', wrap: false, align: 'center', fontSize: 11, fontFace: UI, color: color });
  });
  if (diamonds) {
    [0.669, 1.686, 2.863].forEach(function (dy) {
      s.addShape('diamond', { x: cx - 0.035, y: y + dy, w: 0.071, h: 0.071, fill: { color: diamonds } });
    });
  }
}

/** Footer lock-up: "Panorama - Photography Studio" plus its short rule. */
function footer(s, align) {
  const right = align === 'right';
  s.addText([{ text: 'Panorama ', options: { color: GREY_LT } },
             { text: '– Photography Studio', options: { color: SILVER } }],
    { x: right ? 9.61 : 0.497, y: 6.948, w: 3.253, h: 0.303, valign: 'top',
      align: right ? 'right' : 'left', fontSize: 12, fontFace: UI_SB });
  s.addShape('line', { x: right ? 12.864 : 0, y: 7.099, w: 0.457, h: 0, flipH: true,
    line: { color: SILVER_LT, width: 1.5 } });
}

/** Page number in the corner, e.g. "04." */
function pageNo(s, num, x, y, color) {
  s.addText(num + '.', { x: x, y: y, w: 0.461, h: 0.399, valign: 'middle', align: 'center',
    fontSize: 13, fontFace: UI, color: color });
}

/** Facebook / Instagram / Twitter glyph trio, drawn from primitives. */
function socialGlyph(s, kind, x, y, size, color) {
  const line = { color: color, width: 1.1 };
  if (kind === 'fb') {
    s.addText('f', { x: x, y: y, w: size, h: size, align: 'center', valign: 'middle',
      fontFace: UI_SB, fontSize: size * 62, color: color });
  } else if (kind === 'ig') {
    s.addShape('roundRect', { x: x + size * 0.06, y: y + size * 0.06, w: size * 0.88, h: size * 0.88,
      rectRadius: size * 0.24, fill: { type: 'none' }, line: line });
    s.addShape('ellipse', { x: x + size * 0.29, y: y + size * 0.29, w: size * 0.42, h: size * 0.42,
      fill: { type: 'none' }, line: line });
  } else if (kind === 'tw') {
    s.addShape('triangle', { x: x + size * 0.08, y: y + size * 0.25, w: size * 0.84, h: size * 0.5,
      rotate: 90, fill: { color: color } });
  }
}

/** Vertical stack of social glyphs used on the dark side rails. */
function socialRail(s, x, y, h, color) {
  ['fb', 'ig', 'tw'].forEach(function (g, i) {
    socialGlyph(s, g, x, y + (i * h) / 3, h / 3.6, color);
  });
}

/** Horizontal row of social glyphs (title / closing slides). */
function socialRow(s, x, y, w, h, color) {
  const g = ['fb', 'ig', 'tw'];
  g.forEach(function (kind, i) {
    socialGlyph(s, kind, x + i * h * 1.45, y, h, color);
  });
}

// Reusable typography presets (spread into the option objects below)
const TS1 = { valign: 'top', align: 'justify', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY };
const TS2 = { valign: 'top', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY };
const TS3 = { valign: 'top', align: 'center', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER };
const TS4 = { valign: 'middle', wrap: false, align: 'left', paraSpaceBefore: 6, fontSize: 23, fontFace: UI_SB, color: SNOW, charSpacing: 3 };
const TS5 = { valign: 'top', align: 'justify', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER };
const TS6 = { valign: 'top', fontSize: 13, fontFace: UI, color: INK_45 };
const TS7 = { valign: 'middle', margin: [3.6, 0, 3.6, 0], lineSpacingMultiple: 1.5, fontSize: 10, fontFace: SANS, color: SNOW };
const TS8 = { valign: 'middle', wrap: false, align: 'left', fontSize: 9, fontFace: SANS, color: GREY };
const TS9 = { valign: 'top', wrap: false, fontSize: 12, fontFace: UI, color: INK_45 };
const TS10 = { valign: 'middle', align: 'center', fontSize: 13, fontFace: UI, color: RED_DK };
const TS11 = { valign: 'middle', wrap: false, fontSize: 32, fontFace: UI_SB, color: INK_60 };
const TS12 = { valign: 'top', fontSize: 14, fontFace: UI, color: INK_45 };
const TS13 = { valign: 'top', wrap: false, align: 'center', fontSize: 14, fontFace: UI, color: SNOW };
const TS14 = { valign: 'top', fontSize: 14, fontFace: UI, color: INK_45, charSpacing: 1.5 };

function slide01(pptx) {
  const s = pptx.addSlide();
  corner(s, 7.474, 0.525, 1.246, 0.934, false, false, WHITE, 1);
  corner(s, 2.861, 0.525, 1.246, 0.934, true, false, WHITE, 1);
  corner(s, 7.474, 4.511, 1.246, 0.934, false, true, WHITE, 1);
  corner(s, 2.861, 4.511, 1.246, 0.934, true, true, WHITE, 1);
  s.addText('PANORAMA',
    { x: 5.329, y: 2.236, w: 6.304, h: 1.447, valign: 'top', wrap: false, align: 'right', fontSize: 80, fontFace: SERIF, color: WHITE });
  s.addText('Photography Presentation Template',
    { x: 7.338, y: 3.498, w: 4.295, h: 0.37, valign: 'top', wrap: false, align: 'right', fontSize: 16, fontFace: UI, color: SNOW });
  s.addText('PLACEHOLDER',
    { x: -1.239, y: 1.45, w: 3.405, h: 0.505, rotate: 270, valign: 'top', fontSize: 12, fontFace: SANS, color: SILVER_LT });
  s.addText('PLACEHOLDER',
    { x: 10.68, y: 6.581, w: 2.632, h: 0.707, valign: 'top', fontSize: 12, fontFace: SANS, color: SILVER_LT });
  s.addText('FOLLOW US',
    { x: 0.211, y: 6.749, w: 1.944, h: 0.37, valign: 'top', fontSize: 16, fontFace: UI, color: WHITE, charSpacing: 3 });
  socialRow(s, 2.33, 6.796, 1.7, 0.273, WHITE);
}

function slide02(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 12.415, y: 0, w: 0.918, h: 4.533, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 0, y: 2.15, w: 0.761, h: 0, line: { color: SILVER_LT, width: 1.5 } });
  s.addShape('diamond', { x: 0.747, y: 2.077, w: 0.157, h: 0.145, line: { color: SILVER_LT, width: 1.5 } });
  s.addText([{ text: 'Welcome to ' }, { text: 'Panorama ', options: { bold: true, breakLine: true } }, { text: 'Photography Studio' }],
    { x: 1.247, y: 1.734, w: 3.971, h: 1.01, valign: 'top', fontSize: 27, fontFace: SERIF, color: INK });
  s.addText(lorem(60, false),
    { ...TS1, x: 1.247, y: 3.09, w: 4.553, h: 1.731 });
  s.addText('EXPLORE',
    { x: 1.247, y: 5.166, w: 1.536, h: 0.6, fill: { color: RED }, valign: 'middle', align: 'center', fontSize: 14, fontFace: SANS, color: WHITE, charSpacing: 3 });
  s.addText([{ text: '0' }, { text: '2' }, { text: '.' }],
    { ...TS10, x: 12.662, y: 6.835, w: 0.461, h: 0.399 });
}

function slide03(pptx) {
  const s = pptx.addSlide();
  s.addShape('custGeom', { x: 9.526, y: 0, w: 3.808, h: 7.5, rotate: 180, flipH: true, fill: { color: RED }, points: [{ x: 2.135, y: 0, moveTo: true }, { x: 3.808, y: 0 }, { x: 3.808, y: 7.5 }, { x: 0, y: 7.5 }, { x: 2.135, y: 0 }, { close: true }] });
  corner(s, 11.412, 0.266, 1.246, 0.934, false, false, SILVER, 1);
  corner(s, 6.378, 6.287, 1.246, 0.934, true, true, SILVER_LT, 1);
  s.addText([{ text: 'A Brief History of ' }, { text: 'Photography ', options: { bold: true } }, { text: 'and the People Who Made It ' }, { text: 'Succed', options: { bold: true } }],
    { x: 0.698, y: 1.787, w: 4.983, h: 1.313, valign: 'top', fontSize: 24, fontFace: SERIF, color: INK });
  s.addText('The History of Panorama',
    { x: 0.698, y: 3.631, w: 2.626, h: 0.337, valign: 'top', wrap: false, fontSize: 14, fontFace: UI, color: INK_60 });
  s.addText(lorem(60, false),
    { ...TS1, x: 0.698, y: 3.967, w: 4.983, h: 1.731 });
}

function slide04(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 3.646, h: 7.515, rotate: 180, fill: { color: RED } });
  corner(s, 0.784, 3.75, 2.862, 2.471, true, true, WHITE, 1);
  s.addShape('line', { x: 1.123, y: 1.276, w: 2.523, h: 0, flipH: true, line: { color: SILVER, width: 0.75 } });
  s.addText(lorem(60, false),
    { ...TS1, x: 8.039, y: 3.601, w: 4.658, h: 1.731 });
  s.addText([{ text: 'Make ' }, { text: 'Yourself', options: { bold: true } }, { text: ' Together with Nature' }],
    { x: 8.039, y: 2.168, w: 4.658, h: 1.01, valign: 'top', fontSize: 27, fontFace: SERIF, color: INK });
  s.addText([{ text: '0' }, { text: '4' }, { text: '.' }],
    { x: 0.214, y: 0.253, w: 0.471, h: 0.399, valign: 'middle', align: 'center', fontSize: 13, fontFace: UI, color: SILVER });
  s.addText(' Photography Stuido',
    { x: 0.064, y: 2.176, w: 1.331, h: 0.471, rotate: 270, valign: 'top', align: 'right', fontSize: 11, fontFace: UI, color: SILVER_LT });
  tagRow(s, 8.248, 0.392, SILVER);
  footer(s, 'right');
}

function slide05(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 9.687, y: 0, w: 3.646, h: 7.515, rotate: 180, fill: { color: RED } });
  s.addText(lorem(91, false),
    { ...TS1, x: 1.033, y: 2.601, w: 6.785, h: 1.731 });
  s.addText([{ text: 'Photography is One Way to Express ' }, { text: 'Yourself', options: { bold: true } }],
    { x: 1.016, y: 1.184, w: 5.455, h: 1.01, valign: 'top', fontSize: 27, fontFace: SERIF, color: INK });
  tagRow(s, 3.362, 0.244, SILVER);
  s.addShape('line', { x: 0, y: 0.387, w: 3.18, h: 0, flipH: true, line: { color: SILVER_LT } });
}

function slide06(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 10.046, h: 7.515, rotate: 180, fill: { color: RED } });
  corner(s, 3.287, 0.288, 5.913, 2.065, false, false, SILVER, 0.5);
  s.addShape('line', { x: 10.539, y: 0, w: 0, h: 7.5, line: { color: SILVER_LT } });
  s.addText([{ text: 'Art In ' }, { text: 'Photography', options: { italic: true } }],
    { x: 7.435, y: 2.599, w: 2.415, h: 1.043, valign: 'top', align: 'right', fontSize: 28, fontFace: SERIF, color: WHITE });
  s.addText(lorem(43),
    { ...TS5, x: 7.435, y: 4.02, w: 2.415, h: 2.564 });
  s.addText('Beauty Model',
    { x: 1.384, y: 0.127, w: 1.668, h: 0.37, valign: 'top', align: 'right', fontSize: 16, fontFace: SERIF, color: SILVER_LT });
  s.addText('About Us',
    { x: 11.233, y: 3.531, w: 1.478, h: 0.438, rotate: 270, valign: 'top', align: 'center', fontSize: 20, fontFace: UI, color: INK_45 });
  s.addText(lorem(12),
    { x: 10.863, y: 5.686, w: 2.22, h: 0.897, valign: 'top', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY_LT });
  s.addText(' Photography Stuido',
    { x: 11.307, y: 1.346, w: 1.331, h: 0.471, rotate: 270, valign: 'top', align: 'right', fontSize: 11, fontFace: UI, color: GREY_LT });
}

function slide07(pptx) {
  const s = pptx.addSlide();
  s.addShape('custGeom', { x: 0, y: 0, w: 7.75, h: 7.515, rotate: 180, fill: { color: RED }, points: [{ x: 2.827, y: 0, moveTo: true }, { x: 7.75, y: 0 }, { x: 7.75, y: 7.515 }, { x: 2.827, y: 7.515 }, { x: 0, y: 7.515 }, { x: 2.827, y: 0 }, { close: true }] });
  corner(s, 11.647, 0.536, 1.246, 0.934, false, false, SILVER_LT, 1);
  s.addText([{ text: 'The ' }, { text: 'Wonders', options: { italic: true } }, { text: ' of the World in Just One Shot' }],
    { x: 0.811, y: 4.853, w: 4.208, h: 1.717, valign: 'top', fontSize: 32, fontFace: SERIF, color: WHITE });
  s.addText(lorem(25),
    { ...TS1, x: 6.114, y: 5.394, w: 2.969, h: 1.175 });
  s.addText('Stunning coastal areas',
    { x: 6.114, y: 4.853, w: 2.812, h: 0.37, valign: 'top', wrap: false, fontSize: 16, fontFace: UI, color: INK_60 });
  s.addText(lorem(36),
    { ...TS1, x: 9.553, y: 4.853, w: 2.969, h: 1.731 });
}

function slide08(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 9.128, y: 7.168, w: 4.192, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 0, y: 0.33, w: 4.192, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 6.89, y: 5.39, w: 2.353, h: 2.11, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 4.206, y: 0, w: 2.461, h: 2.11, rotate: 180, fill: { color: RED } });
  corner(s, 4.206, 4.459, 1.394, 0.877, true, true, SILVER_LT, 1);
  corner(s, 7.849, 1.233, 1.394, 0.877, false, false, SILVER_LT, 1);
  s.addText([{ text: 'The ' }, { text: 'camera', options: { bold: true } }, { text: ' is an instrument that ' }, { text: 'teaches', options: { bold: true } }, { text: ' people how to see without a ' }, { text: 'camera', options: { bold: true } }],
    { x: 0.802, y: 2.185, w: 2.93, h: 3.13, valign: 'top', paraSpaceBefore: 6, fontSize: 30, fontFace: SERIF, color: INK });
  s.addText(lorem(25),
    { ...TS2, x: 9.717, y: 2.363, w: 2.969, h: 1.175 });
  s.addText('Potrait Photography',
    { x: 9.717, y: 1.924, w: 2.814, h: 0.337, valign: 'top', fontSize: 14, fontFace: UI, color: INK_60 });
  s.addText(lorem(25),
    { ...TS2, x: 9.717, y: 4.401, w: 2.969, h: 1.175 });
  s.addText('Creative With Technique',
    { x: 9.717, y: 3.962, w: 2.814, h: 0.337, valign: 'top', fontSize: 14, fontFace: UI, color: INK_60 });
  footer(s, 'left');
  s.addText([{ text: '0' }, { text: '8' }, { text: '.' }],
    { ...TS10, x: 12.589, y: 0.268, w: 0.534, h: 0.399 });
}

function slide09(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 7.849, y: 0, w: 5.484, h: 4.507, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 5.949, y: 0, w: 0, h: 7.5, line: { color: SILVER_LT } });
  corner(s, 6.667, 2.986, 4.84, 3.365, false, true, SILVER_LT, 1);
  s.addShape('line', { x: 8.562, y: 1.167, w: 2.523, h: 0, flipH: true, line: { color: SILVER, width: 0.75 } });
  s.addText([{ text: '0' }, { text: '9' }, { text: '.' }],
    { x: 12.662, y: 0.268, w: 0.461, h: 0.399, valign: 'middle', align: 'center', fontSize: 13, fontFace: UI, color: SILVER });
  s.addText('Photography Stuido',
    { x: 10.844, y: 1.926, w: 1.331, h: 0.471, rotate: 90, flipH: true, valign: 'top', fontSize: 11, fontFace: UI, color: SILVER_LT });
  tagColumn(s, 6.327, 1.889, SILVER, SILVER);
  s.addText(lorem(60),
    { ...TS1, x: 0.674, y: 3.75, w: 4.703, h: 1.731 });
  s.addText([{ text: 'Explore Nature and Beauty ' }, { text: 'Photography', options: { bold: true } }],
    { x: 0.674, y: 2.019, w: 4.703, h: 1.043, valign: 'top', fontSize: 28, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 0.797, y: 3.374, w: 1.559, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
}

function slide10(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 3.466, y: 7.057, w: 9.329, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 6.973, y: 0.452, w: 5.822, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 12.794, y: 0.452, w: 0, h: 6.605, flipV: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0.589, y: 3.95, w: 3.093, h: 3.255, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 3.942, y: 0.253, w: 3.093, h: 3.255, rotate: 180, fill: { color: RED } });
  s.addText(lorem(60),
    { ...TS1, x: 7.992, y: 3.504, w: 3.698, h: 2.286 });
  s.addText([{ text: 'Alone we can do so little, ' }, { text: 'together', options: { bold: true } }, { text: ' we can do so much' }],
    { x: 7.992, y: 1.73, w: 3.698, h: 1.313, valign: 'top', fontSize: 24, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 8.079, y: 3.233, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('Owner & Founder',
    { x: 1.827, y: 4.584, w: 1.532, h: 0.286, valign: 'top', wrap: false, align: 'right', fontSize: 11, fontFace: UI, color: SILVER });
  s.addText('Valerie Tumewu',
    { x: 1.324, y: 5.006, w: 2.036, h: 0.37, valign: 'top', wrap: false, align: 'right', fontSize: 16, fontFace: UI_SB, color: WHITE });
  s.addText(lorem(12),
    { x: 0.912, y: 5.497, w: 2.447, h: 0.897, valign: 'top', align: 'right', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER_LT });
  s.addText('Photographer',
    { x: 4.219, y: 0.966, w: 1.277, h: 0.286, valign: 'top', wrap: false, fontSize: 11, fontFace: UI, color: SILVER });
  s.addText('Kamila Maciejewska',
    { x: 4.219, y: 1.386, w: 2.537, h: 0.37, valign: 'top', wrap: false, fontSize: 16, fontFace: UI_SB, color: WHITE });
  s.addText(lorem(12),
    { x: 4.219, y: 1.873, w: 2.447, h: 0.897, valign: 'top', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER_LT });
  s.addText([{ text: 'Panorama ' }, { text: 'People', options: { bold: true } }],
    { x: 12.039, y: 3.603, w: 1.52, h: 0.303, rotate: 90, fill: { color: WHITE }, valign: 'top', wrap: false, align: 'right', fontSize: 12, fontFace: SERIF, color: INK_60 });
}

function slide11(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 0, y: 0.549, w: 4.763, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 8.57, y: 0.549, w: 4.75, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addText([{ text: 'Creative', options: { bold: true, fontSize: 27, breakLine: true } }, { text: 'Team Placeholder', options: { fontSize: 20 } }],
    { x: 4.817, y: 0.118, w: 3.698, h: 0.892, valign: 'top', align: 'center', fontFace: SERIF, color: INK });
  s.addText('Make-Up Artist',
    { x: 1.67, y: 5.255, w: 1.343, h: 0.286, valign: 'top', wrap: false, fontSize: 11, fontFace: UI, color: INK_45 });
  s.addText('Tetiana Sadrina',
    { x: 1.67, y: 4.902, w: 1.895, h: 0.353, valign: 'top', wrap: false, fontSize: 15, fontFace: UI_SB, color: INK_60 });
  s.addText(lorem(12),
    { ...TS2, x: 1.67, y: 5.608, w: 2.727, h: 0.62 });
  s.addText('General Manager',
    { x: 5.518, y: 5.255, w: 1.555, h: 0.286, valign: 'top', wrap: false, fontSize: 11, fontFace: UI, color: SILVER });
  s.addText('Monica Takada',
    { x: 5.518, y: 4.902, w: 1.843, h: 0.353, valign: 'top', wrap: false, fontSize: 15, fontFace: UI_SB, color: SNOW });
  s.addText(lorem(12),
    { x: 5.518, y: 5.609, w: 2.297, h: 0.897, valign: 'top', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER_LT });
  s.addText('Event Planner',
    { x: 8.575, y: 5.254, w: 1.254, h: 0.286, valign: 'top', wrap: false, fontSize: 11, fontFace: UI, color: INK_45 });
  s.addText('Anaya Hakim',
    { x: 8.57, y: 4.902, w: 1.636, h: 0.353, valign: 'top', wrap: false, fontSize: 15, fontFace: UI_SB, color: INK_60 });
  s.addText(lorem(12),
    { ...TS2, x: 8.57, y: 5.608, w: 2.896, h: 0.62 });
  footer(s, 'left');
  s.addText([{ text: '11' }, { text: '.' }],
    { ...TS10, x: 12.662, y: 6.835, w: 0.461, h: 0.399 });
}

function slide12(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 3.466, y: 7.057, w: 9.329, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 2.397, y: 0.452, w: 10.397, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 12.794, y: 0.452, w: 0, h: 6.605, flipV: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 4.589, h: 7.5, rotate: 180, fill: { color: RED } });
  corner(s, 1.827, 2.165, 4.84, 4.243, true, true, SNOW, 1);
  s.addShape('line', { x: 2.541, y: 1.208, w: 2.523, h: 0, flipH: true, line: { color: SILVER_LT, width: 0.75 } });
  s.addText(lorem(43),
    { ...TS1, x: 7.814, y: 4.053, w: 3.698, h: 1.453 });
  s.addText([{ text: 'Daniel', options: { bold: true } }, { text: ' Leonal ', options: { breakLine: true } }, { text: 'Monteiro' }],
    { x: 7.814, y: 2.28, w: 3.698, h: 1.178, valign: 'top', fontSize: 32, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 7.9, y: 3.783, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('Editor Professional',
    { x: 7.814, y: 1.994, w: 1.618, h: 0.286, valign: 'top', wrap: false, fontSize: 11, fontFace: UI, color: GREY });
  s.addText([{ text: 'Panorama ' }, { text: 'People', options: { bold: true } }],
    { x: 12.039, y: 3.603, w: 1.52, h: 0.303, rotate: 90, fill: { color: WHITE }, valign: 'top', wrap: false, align: 'right', fontSize: 12, fontFace: SERIF, color: INK_60 });
  s.addText(lorem(18),
    { ...TS5, x: 0.253, y: 0.713, w: 1.994, h: 1.453 });
}

function slide13(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 0, y: 0.33, w: 5.384, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 9.274, y: 0, w: 4.059, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText('Brandon Ladera Jackson',
    { x: 5.384, y: 6.107, w: 3.379, h: 1.111, valign: 'top', align: 'right', fontSize: 30, fontFace: SERIF, color: INK });
  s.addText('Art Director',
    { x: 7.677, y: 5.821, w: 1.085, h: 0.286, valign: 'top', wrap: false, align: 'right', fontSize: 11, fontFace: UI, color: GREY });
  s.addText('Olivia Jenita Jhonson',
    { x: 9.08, y: 0.439, w: 3.379, h: 1.178, valign: 'top', align: 'right', fontSize: 32, fontFace: SERIF, color: SNOW });
  s.addText('Designer',
    { x: 11.57, y: 0.152, w: 0.889, h: 0.286, valign: 'top', wrap: false, align: 'right', fontSize: 11, fontFace: UI, color: SILVER });
  s.addText(lorem(60),
    { ...TS1, x: 0.833, y: 3.591, w: 3.698, h: 2.286 });
  s.addText([{ text: 'Our Creative Team Member of ' }, { text: 'Panorama', options: { bold: true } }],
    { x: 0.833, y: 1.465, w: 3.698, h: 1.464, valign: 'top', fontSize: 27, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 0.933, y: 3.32, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  footer(s, 'left');
}

function slide14(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 2.397, y: 0.261, w: 10.647, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 13.044, y: 0.261, w: 0, h: 7.239, flipV: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 4.027, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText('Professional Services From Panorama',
    { x: 0.202, y: 1.893, w: 2.842, h: 1.414, valign: 'top', fontSize: 26, fontFace: SERIF, color: WHITE });
  s.addText(lorem(43),
    { ...TS5, x: 0.202, y: 3.599, w: 2.842, h: 2.008 });
  s.addText(lorem(12),
    { ...TS2, x: 9.662, y: 1.194, w: 2.628, h: 0.897 });
  s.addText('Creative with Technique',
    { ...TS9, x: 9.662, y: 0.891, w: 2.255, h: 0.303 });
  s.addShape('rect', { x: 8.344, y: 1.037, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'tripod', 8.551, 1.102, 0.768, WHITE);
  s.addText(lorem(12),
    { ...TS2, x: 9.662, y: 2.698, w: 2.628, h: 0.897 });
  s.addText('Perfect Shots',
    { ...TS9, x: 9.662, y: 2.395, w: 1.313, h: 0.303 });
  s.addText(lorem(12),
    { ...TS2, x: 9.662, y: 4.201, w: 2.628, h: 0.897 });
  s.addText('Outdoor Photography',
    { ...TS9, x: 9.662, y: 3.898, w: 2.053, h: 0.303 });
  s.addText(lorem(12),
    { ...TS2, x: 9.662, y: 5.711, w: 2.628, h: 0.897 });
  s.addText('Indoor Photography',
    { ...TS9, x: 9.662, y: 5.408, w: 1.899, h: 0.303 });
  s.addShape('rect', { x: 8.343, y: 2.541, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'aperture', 8.533, 2.732, 0.515, WHITE);
  s.addShape('rect', { x: 8.343, y: 4.044, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'camera', 8.47, 4.167, 0.641, WHITE);
  s.addShape('rect', { x: 8.343, y: 5.554, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'umbrella', 8.521, 5.628, 0.749, WHITE);
}

function slide15(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 1.465, y: 7.053, w: 9.329, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 0.397, y: 0.448, w: 10.397, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('line', { x: 10.794, y: 0.448, w: 0, h: 6.605, flipV: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 9.863, y: 0, w: 3.47, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText('01.',
    { x: 10.486, y: 1.411, w: 0.616, h: 0.627, fill: { color: WHITE }, line: { color: WHITE }, valign: 'middle', align: 'center', fontSize: 11, fontFace: UI, color: INK_60 });
  s.addText('02.',
    { x: 10.486, y: 3.437, w: 0.616, h: 0.627, fill: { color: WHITE }, line: { color: WHITE }, valign: 'middle', align: 'center', fontSize: 11, fontFace: UI, color: INK_60 });
  s.addText('03.',
    { x: 10.486, y: 5.462, w: 0.616, h: 0.627, fill: { color: WHITE }, line: { color: WHITE }, valign: 'middle', align: 'center', fontSize: 11, fontFace: UI, color: INK_60 });
  s.addText(lorem(22),
    { ...TS2, x: 0.934, y: 1.609, w: 3.522, h: 0.897 });
  s.addText('01. Perfect Shots',
    { ...TS9, x: 0.934, y: 1.273, w: 1.54, h: 0.303 });
  s.addText(lorem(22),
    { ...TS2, x: 0.927, y: 3.509, w: 3.522, h: 0.897 });
  s.addText('02. Creative with Technique',
    { ...TS9, x: 0.927, y: 3.173, w: 2.525, h: 0.303 });
  s.addText(lorem(22),
    { ...TS2, x: 0.931, y: 5.462, w: 3.522, h: 0.897 });
  s.addText('03. Outdoor Photography',
    { x: 0.931, y: 5.126, w: 2.398, h: 0.303, valign: 'top', fontSize: 12, fontFace: UI, color: INK_45 });
}

function slide16(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 3.068, y: 7.053, w: 6.931, h: 0, line: { color: SILVER_LT } });
  s.addShape('line', { x: 3.068, y: 0.448, w: 7.725, h: 0, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 3.562, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText([{ text: 'Excellent Services Provide For ' }, { text: 'Photography', options: { bold: true } }],
    { x: 4.642, y: 3.556, w: 4.049, h: 0.909, valign: 'top', align: 'center', fontSize: 24, fontFace: SERIF, color: INK });
  s.addText(lorem(43),
    { ...TS1, x: 4.642, y: 5.074, w: 4.049, h: 1.453 });
  s.addShape('line', { x: 6.035, y: 4.707, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText(lorem(12),
    { ...TS3, x: 0.421, y: 2.294, w: 2.628, h: 0.897 });
  s.addText('Perfect Shots',
    { ...TS13, x: 0.985, y: 1.846, w: 1.499, h: 0.337 });
  icon(s, 'aperture', 1.167, 0.515, 1.136, WHITE);
  s.addText(lorem(12),
    { ...TS3, x: 0.512, y: 6.087, w: 2.628, h: 0.897 });
  s.addText('Outdoor Photography',
    { ...TS13, x: 0.646, y: 5.635, w: 2.362, h: 0.337 });
  icon(s, 'camera', 1.167, 4.168, 1.136, WHITE);
}

function slide17(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 3.068, y: 7.053, w: 6.931, h: 0, line: { color: SILVER_LT } });
  s.addShape('line', { x: 3.068, y: 0.448, w: 7.725, h: 0, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 3.219, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 10.114, y: 0, w: 3.219, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText('Indoor Photography',
    { x: 0.734, y: 0.778, w: 1.87, h: 0.735, fill: { color: RED }, valign: 'middle', align: 'right', fontFace: UI, color: WHITE });
  s.addText([{ text: 'Explore', options: { bold: true } }, { text: ' Our Services' }],
    { x: 4.8, y: 1.855, w: 3.733, h: 1.178, valign: 'top', align: 'center', fontSize: 32, fontFace: SERIF, color: INK });
  s.addText(lorem(43),
    { ...TS1, x: 5.009, y: 3.915, w: 3.316, h: 1.731 });
  s.addShape('line', { x: 6.035, y: 3.485, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText([{ text: 'Perfect ', options: { breakLine: true } }, { text: 'Shots' }],
    { x: 10.757, y: 0.778, w: 1.87, h: 0.735, fill: { color: RED }, valign: 'middle', fontFace: UI, color: WHITE });
  s.addShape('rect', { x: 8.713, y: 5.707, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'aperture', 8.903, 5.899, 0.515, WHITE);
  s.addShape('rect', { x: 3.725, y: 5.736, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'umbrella', 3.903, 5.81, 0.749, WHITE);
}

function slide18(pptx) {
  const s = pptx.addSlide();
  corner(s, 6.093, 0.255, 1.246, 0.934, false, false, WHITE, 1);
  corner(s, 1.479, 0.255, 1.246, 0.934, true, false, WHITE, 1);
  corner(s, 6.093, 4.242, 1.246, 0.934, false, true, WHITE, 1);
  corner(s, 1.479, 4.242, 1.246, 0.934, true, true, WHITE, 1);
  s.addText('BREAK SLIDE',
    { x: 5.159, y: 1.592, w: 6.474, h: 1.313, valign: 'top', wrap: false, align: 'right', fontSize: 72, fontFace: SERIF, color: WHITE });
  s.addText('It’s time for a break',
    { x: 9.237, y: 2.72, w: 2.397, h: 0.37, valign: 'top', wrap: false, align: 'right', fontSize: 16, fontFace: UI, color: SILVER_LT });
}

function slide19(pptx) {
  const s = pptx.addSlide();
  s.addText([{ text: 'Behind the Shot :', options: { fontSize: 24, breakLine: true } }, { text: 'Storm in the ', options: { fontSize: 36 } }, { text: 'Desert', options: { italic: true, fontSize: 36 } }],
    { x: 8.41, y: 2.949, w: 4.543, h: 1.111, valign: 'top', wrap: false, align: 'right', fontFace: SERIF, color: SNOW });
  s.addText([{ text: 'What formula does ' }, { text: 'Panorama Studio ', options: { bold: true } }, { text: 'use for good shooting results ?' }],
    { x: 5.034, y: 5.253, w: 3.517, h: 1.485, valign: 'middle', fontSize: 22, fontFace: SERIF, color: INK_60 });
  s.addText(lorem(42),
    { ...TS1, x: 8.962, y: 5.269, w: 3.517, h: 1.453 });
  s.addShape('rect', { x: 0.002, y: 1.828, w: 4.444, h: 5.69, rotate: 180, fill: { color: RED } });
  s.addText('Portfolios',
    { x: 0.125, y: 2.158, w: 1.552, h: 0.471, valign: 'top', wrap: false, fontSize: 22, fontFace: SERIF, color: SNOW });
  s.addText(lorem(12),
    { ...TS5, x: 0.125, y: 2.717, w: 2.029, h: 0.897 });
}

function slide20(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 9.13, y: 0.33, w: 4.192, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('custGeom', { x: 3.111, y: 4.649, w: 2.181, h: 2.167, rotate: 180, flipH: true, fill: { color: SNOW }, points: [{ x: 0, y: 2.167, moveTo: true }, { x: 0.044, y: 0 }, { x: 2.181, y: 2.167 }, { x: 0, y: 2.167 }, { close: true }] });
  s.addShape('triangle', { x: 3.43, y: 1.089, w: 2.181, h: 2.167, rotate: 338.71, flipH: true, fill: { color: SNOW } });
  s.addShape('custGeom', { x: 0.615, y: 0.009, w: 4.104, h: 4.145, rotate: 177.3, fill: { color: RED }, points: [{ x: 0, y: 2.06, moveTo: true }, { x: 2.032, y: 0 }, { x: 4.104, y: 2.143 }, { x: 2.004, y: 4.145 }, { x: 0, y: 2.06 }, { close: true }] });
  s.addText([{ text: 'Photographers', options: { bold: true } }, { text: ' and Photo Editors on the Passion That Drives Their Work ' }],
    { x: 8.181, y: 3.094, w: 4.366, h: 1.313, valign: 'top', align: 'right', fontSize: 24, fontFace: SERIF, color: INK });
  s.addText(lorem(43),
    { ...TS1, x: 7.986, y: 4.7, w: 4.561, h: 1.175 });
  footer(s, 'left');
  s.addText([{ text: '20' }, { text: '.' }],
    { ...TS10, x: 12.662, y: 6.835, w: 0.461, h: 0.399 });
  s.addText('Photography also includes moving images captured on film or video.',
    { x: 1.29, y: 2.074, w: 2.754, h: 0.808, valign: 'top', align: 'center', fontSize: 14, fontFace: UI_LT, color: WHITE });
  icon(s, 'aperture', 2.409, 1.28, 0.515, SNOW);
}

function slide21(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 2.208, y: 0.207, w: 11.124, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 2.319, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText(lorem(60),
    { ...TS1, x: 7.877, y: 3.873, w: 4.324, h: 2.008 });
  s.addText([{ text: 'Photography', options: { bold: true } }, { text: ' is the ' }, { text: 'Quickest', options: { italic: true } }, { text: ', ', options: { breakLine: true } }, { text: 'Most Exact Tool Ever Invented to Record ', options: { breakLine: true } }, { text: 'Our Lives' }],
    { x: 7.877, y: 1.316, w: 4.455, h: 1.717, valign: 'top', fontSize: 24, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 7.995, y: 3.49, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  footer(s, 'right');
  s.addText('Perfect their almost hallucinatory understanding of the visual world.',
    { x: 3.875, y: 5.949, w: 3.014, h: 0.808, valign: 'top', align: 'center', fontSize: 14, fontFace: UI, color: INK_60 });
}

function slide22(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 12.486, y: 0, w: 0.847, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 12.91, y: 0, w: 0, h: 7.5, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 12.486, y: 1.476, w: 0.847, h: 4.547, rotate: 180, fill: { color: RED } });
  s.addText([{ text: 'Today', options: { bold: true } }, { text: ', photography is the only art that seriously maintains this attention to the stuff that ' }, { text: 'matters', options: { bold: true } }, { text: '.' }],
    { x: 1.851, y: 0.303, w: 4.667, h: 1.043, valign: 'top', align: 'center', fontFace: SERIF, color: INK });
  s.addText(lorem(25),
    { x: 1.344, y: 6.577, w: 5.681, h: 0.62, valign: 'top', align: 'center', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY });
  tagColumn(s, 12.91, 1.889, WHITE, WHITE);
}

function slide23(pptx) {
  const s = pptx.addSlide();
  s.addShape('custGeom', { x: 1.695, y: -0.007, w: 5.581, h: 3.364, rotate: 180, fill: { color: RED }, points: [{ x: 0, y: 3.364, moveTo: true }, { x: 2.747, y: 0 }, { x: 2.961, y: 0.001 }, { x: 5.581, y: 3.357 }, { x: 0, y: 3.364 }, { close: true }] });
  tagRow(s, 8.248, 0.379, SILVER);
  s.addText([{ text: 'Photographers', options: { bold: true } }, { text: ' Use Their Cameras As Tools of Exploration' }],
    { x: 8.101, y: 1.901, w: 4.561, h: 1.313, valign: 'top', align: 'right', fontSize: 24, fontFace: SERIF, color: INK });
  s.addText(lorem(73, false),
    { ...TS1, x: 8.101, y: 3.507, w: 4.561, h: 2.008 });
  s.addText([{ text: '23' }, { text: '.' }],
    { ...TS10, x: 12.662, y: 6.835, w: 0.461, h: 0.399 });
  s.addText('A photograph is an image captured on film, paper or most commonly now in digital memory.',
    { x: 3.108, y: 1.138, w: 2.754, h: 1.043, valign: 'top', align: 'center', fontSize: 14, fontFace: UI_LT, color: WHITE });
  icon(s, 'tripod', 4.246, 0.186, 0.768, WHITE);
}

function slide24(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 0, y: 0.33, w: 5.384, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 8.722, y: 0, w: 4.611, h: 7.5, rotate: 180, fill: { color: RED } });
  footer(s, 'left');
  s.addText(lorem(16),
    { x: 0.905, y: 3.737, w: 2.8, h: 0.62, valign: 'top', align: 'center', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY });
  s.addText('Choose Your Location',
    { x: 0.978, y: 3.376, w: 2.655, h: 0.303, valign: 'top', align: 'center', fontSize: 12, fontFace: UI, color: INK_45 });
  s.addText(lorem(16),
    { x: 0.905, y: 5.574, w: 2.8, h: 0.62, valign: 'top', align: 'center', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: GREY });
  s.addText('Use A Tripod and Camera',
    { x: 1.08, y: 5.271, w: 2.355, h: 0.303, valign: 'top', wrap: false, align: 'center', fontSize: 12, fontFace: UI, color: INK_45 });
  s.addText('What makes a good photo ? Choise',
    { x: 9.328, y: 1.864, w: 3.4, h: 1.515, valign: 'top', align: 'right', fontSize: 28, fontFace: SERIF, color: WHITE });
  s.addText(lorem(50, false),
    { ...TS5, x: 9.328, y: 3.627, w: 3.4, h: 2.008 });
}

function slide25(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 12.91, y: 2.209, w: 0, h: 5.288, line: { color: SILVER_LT } });
  s.addShape('custGeom', { x: 0, y: 0, w: 13.292, h: 7.5, rotate: 180, fill: { color: SNOW }, points: [{ x: 0, y: 0, moveTo: true }, { x: 13.292, y: 0 }, { x: 13.292, y: 7.5 }, { x: 6.986, y: 7.5 }, { x: 0, y: 0 }, { close: true }] });
  s.addShape('diamond', { x: 2.367, y: 3.72, w: 3.694, h: 3.287, rotate: 180, fill: { color: RED } });
  s.addShape('diamond', { x: 6.646, y: 3.75, w: 2.815, h: 2.505, rotate: 180, fill: { color: RED } });
  icon(s, 'camera', 7.503, 4.424, 1.043, WHITE);
  socialRail(s, 12.741, 0.495, 1.338, SILVER_LT);
  s.addText([{ text: 'That advice may sound crazy coming from a photographer, ', options: { breakLine: true } }, { text: 'but it’s true' }],
    { x: 2.837, y: 4.881, w: 2.754, h: 1.043, valign: 'top', align: 'center', fontSize: 14, fontFace: UI_LT, color: WHITE });
}

function slide26(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 12.486, y: 0, w: 0.847, h: 7.516, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 12.91, y: 0, w: 0, h: 5.288, line: { color: SILVER } });
  s.addShape('ellipse', { x: 12.8, y: 5.738, w: 0.22, h: 0.22, fill: { color: SNOW }, line: { color: WHITE } });
  s.addShape('ellipse', { x: 12.797, y: 6.091, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addShape('ellipse', { x: 12.802, y: 6.443, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addShape('ellipse', { x: 12.8, y: 6.796, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addText('Strenght Analysis Slide',
    { x: 6.763, y: 1.557, w: 3.941, h: 1.212, valign: 'top', fontSize: 33, fontFace: SERIF, color: INK });
  s.addText(lorem(60, false),
    { ...TS1, x: 6.763, y: 3.191, w: 4.561, h: 1.731 });
  s.addText('EXPLORE',
    { x: 9.787, y: 5.343, w: 1.536, h: 0.6, fill: { color: RED }, valign: 'middle', align: 'center', fontSize: 14, fontFace: SANS, color: WHITE, charSpacing: 3 });
}

function slide27(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 0.847, h: 7.516, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 0.424, y: -0.016, w: 0, h: 5.288, line: { color: SILVER_LT } });
  socialRail(s, 0.276, 5.727, 1.338, SILVER_LT);
  s.addText('Weakness Analysis Slide',
    { x: 2.078, y: 1.708, w: 4.233, h: 1.212, valign: 'top', align: 'right', fontSize: 33, fontFace: SERIF, color: INK });
  s.addText(lorem(30),
    { ...TS1, x: 2.078, y: 4.617, w: 4.233, h: 1.175 });
  s.addText('98',
    { x: 4.006, y: 3.258, w: 1.2, h: 0.858, valign: 'top', align: 'right', fontSize: 45, fontFace: UI_SB, color: INK_60 });
  s.addShape('rect', { x: 5.417, y: 3.301, w: 0.894, h: 0.897, fill: { color: RED } });
  icon(s, 'aperture', 5.607, 3.5, 0.515, WHITE);
  s.addText('“ Lorem Ipsum is simply',
    { x: 3.726, y: 4.026, w: 1.48, h: 0.252, valign: 'middle', wrap: false, align: 'right', fontSize: 9, fontFace: SANS, color: GREY });
}

function slide28(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 9.13, y: 0.33, w: 4.192, h: 0, flipH: true, line: { color: SILVER_LT } });
  s.addShape('rect', { x: 0, y: 0, w: 6.667, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addText('Opportunities Analysis Slide',
    { x: 9.466, y: 2.225, w: 3.542, h: 1.144, valign: 'top', fontSize: 31, fontFace: SERIF, color: INK });
  s.addText(lorem(42, false),
    { ...TS1, x: 9.466, y: 3.859, w: 3.542, h: 1.453 });
  footer(s, 'right');
  s.addText(lorem(12),
    { ...TS3, x: 0.647, y: 2.294, w: 2.628, h: 0.897 });
  s.addText('Creative with Technique',
    { ...TS13, x: 0.755, y: 1.796, w: 2.595, h: 0.337 });
  s.addText(lorem(12),
    { ...TS3, x: 0.739, y: 6.087, w: 2.628, h: 0.897 });
  s.addText('Outdoor Photography',
    { ...TS13, x: 0.872, y: 5.635, w: 2.362, h: 0.337 });
  icon(s, 'camera', 1.393, 4.168, 1.136, WHITE);
  icon(s, 'tripod', 1.601, 0.38, 1.136, WHITE);
}

function slide29(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 6.25, h: 7.52, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 12.486, y: 0, w: 0.847, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 12.91, y: 0, w: 0, h: 5.288, line: { color: SILVER } });
  s.addShape('ellipse', { x: 12.8, y: 5.738, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addShape('ellipse', { x: 12.797, y: 6.091, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addShape('ellipse', { x: 12.8, y: 6.796, w: 0.22, h: 0.22, fill: { color: WHITE }, line: { color: WHITE } });
  s.addShape('ellipse', { x: 12.797, y: 6.443, w: 0.22, h: 0.22, line: { color: GREY_LT } });
  s.addText('Threat Analysis Slide',
    { x: 7.063, y: 1.669, w: 3.941, h: 1.212, valign: 'top', fontSize: 33, fontFace: SERIF, color: INK });
  s.addText(lorem(43),
    { ...TS1, x: 7.063, y: 3.303, w: 4.561, h: 1.175 });
  s.addText('84',
    { ...TS11, x: 7.063, y: 4.9, w: 0.78, h: 0.639 });
  s.addText('167',
    { ...TS11, x: 8.605, y: 4.9, w: 0.858, h: 0.639 });
  s.addText('209',
    { ...TS11, x: 10.224, y: 4.891, w: 1.012, h: 0.639 });
  s.addText('“ Lorem Ipsum is simply',
    { ...TS8, x: 7.063, y: 5.579, w: 1.48, h: 0.252 });
  s.addText('“ Lorem Ipsum is simply',
    { ...TS8, x: 8.603, y: 5.579, w: 1.48, h: 0.252 });
  s.addText('“ Lorem Ipsum is simply',
    { ...TS8, x: 10.144, y: 5.579, w: 1.48, h: 0.252 });
}

function slide30(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 7.083, y: 0, w: 6.25, h: 7.5, rotate: 180, fill: { color: RED } });
  device(s, 10.693, 0.884, 3.38, 6.033, { rotate: 341.7 });
  device(s, 7.404, 3.27, 3.38, 4.896, { rotate: 341.7 });
  device(s, 5.899, -0.654, 3.38, 4.029, { rotate: 341.7 });
  s.addText(lorem(25),
    { ...TS2, x: 1.847, y: 4.115, w: 4.053, h: 0.897 });
  s.addText([{ text: 'At this point, What other camera gear and ' }, { text: 'accessories', options: { bold: true } }, { text: ' do you need?' }],
    { x: 0.736, y: 1.095, w: 4.283, h: 1.986, valign: 'top', fontSize: 28, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 0.854, y: 3.316, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addShape('rect', { x: 0.854, y: 3.957, w: 0.841, h: 0.844, fill: { color: RED } });
  icon(s, 'camera', 1.01, 4.115, 0.529, WHITE);
  s.addText('Camera',
    { ...TS12, x: 1.847, y: 3.784, w: 2.11, h: 0.337 });
  s.addText(lorem(25),
    { ...TS2, x: 1.847, y: 5.629, w: 4.053, h: 0.897 });
  s.addText('Other Equipment',
    { ...TS12, x: 1.847, y: 5.292, w: 2.546, h: 0.337 });
  s.addShape('rect', { x: 0.854, y: 5.561, w: 0.841, h: 0.844, fill: { color: RED } });
  icon(s, 'umbrella', 1.095, 5.679, 0.576, WHITE);
}

function slide31(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 5.836, h: 7.516, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 1.123, y: 6.211, w: 2.523, h: 0, flipH: true, line: { color: SILVER_LT, width: 0.75 } });
  corner(s, 0.784, 1.272, 2.471, 2.862, true, false, WHITE, 1);
  device(s, 1.13, 1.922, 6.75, 3.72, { radius: 0.1, bezel: 0.11 });
  s.addShape('trapezoid', { x: 0.55, y: 5.6, w: 7.9, h: 0.29, flipV: true,
    fill: { color: 'C9CBCE' }, line: { color: '9A9A9A', width: 1 } });
  s.addText([{ text: 'The first steps on your photographic ' }, { text: 'journey', options: { bold: true } }],
    { x: 8.344, y: 2.154, w: 3.941, h: 1.717, valign: 'top', fontSize: 32, fontFace: SERIF, color: INK });
  s.addText(lorem(38),
    { ...TS1, x: 8.344, y: 4.167, w: 4.132, h: 1.175 });
}

function slide32(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 8.979, y: 0, w: 4.354, h: 7.5, rotate: 180, fill: { color: RED } });
  s.addShape('line', { x: 8.209, y: 1.46, w: 3.586, h: 0, line: { color: SILVER_LT } });
  s.addShape('line', { x: 12.566, y: 1.979, w: 0, h: 5.521, line: { color: SILVER_LT } });
  s.addShape('line', { x: 6.72, y: 0, w: 0, h: 7.5, line: { color: SILVER_LT } });
  device(s, 7.098, 1.979, 5.117, 5.521, { radius: 0.18, bezel: 0.2 });
  s.addText(lorem(50, false),
    { ...TS1, x: 1.203, y: 3.287, w: 4.318, h: 1.453 });
  s.addText([{ text: 'The ' }, { text: 'Fundamental ', options: { bold: true } }, { text: 'Camera Settings You Should Know' }],
    { x: 1.205, y: 1.467, w: 4.125, h: 1.313, valign: 'top', fontSize: 24, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 1.303, y: 3.031, w: 1.455, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText(lorem(12),
    { ...TS2, x: 1.203, y: 5.413, w: 4.318, h: 0.62 });
  s.addText('Technically a bit more complex',
    { ...TS12, x: 1.203, y: 5.014, w: 3.558, h: 0.337 });
}

function slide33(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 6.667, h: 7.5, rotate: 180, fill: { color: SNOW } });
  s.addText(lorem(30, false),
    { ...TS2, x: 0.705, y: 4.011, w: 3.493, h: 1.175 });
  s.addText([{ text: 'We are always with you in the journey of your ' }, { text: 'projects', options: { bold: true } }],
    { x: 0.705, y: 2.191, w: 3.493, h: 1.212, valign: 'top', fontSize: 22, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 0.803, y: 3.755, w: 1.455, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addShape('line', { x: 8.366, y: 2.048, w: 2.068, h: 0, flipH: true, line: { color: GREY, width: 0.75 } });
  s.addShape('line', { x: 8.646, y: 2.518, w: 1.788, h: 0, flipH: true, line: { color: GREY_97, width: 0.75 } });
  s.addShape('line', { x: 8.812, y: 3.08, w: 1.622, h: 0, flipH: true, line: { color: GREY_AD, width: 0.75 } });
  s.addShape('line', { x: 8.979, y: 3.737, w: 1.455, h: 0, flipH: true, line: { color: GREY_D3, width: 0.75 } });
  s.addShape('line', { x: 8.919, y: 4.338, w: 1.515, h: 0, flipH: true, line: { color: GREY_D3, width: 0.75 } });
  s.addShape('line', { x: 8.812, y: 4.857, w: 1.622, h: 0, flipH: true, line: { color: GREY_AD, width: 0.75 } });
  s.addShape('line', { x: 8.621, y: 5.363, w: 1.813, h: 0, flipH: true, line: { color: GREY_97, width: 0.75 } });
  s.addShape('line', { x: 8.366, y: 5.772, w: 2.068, h: 0, flipH: true, line: { color: GREY, width: 0.75 } });
  s.addText('Creative with Technique',
    { ...TS6, x: 10.787, y: 1.871, w: 2.546, h: 0.32 });
  s.addText('Perfect Shots',
    { ...TS6, x: 10.787, y: 2.342, w: 2.213, h: 0.32 });
  s.addText('Outdoor Photography',
    { ...TS6, x: 10.787, y: 2.904, w: 2.546, h: 0.32 });
  s.addText('Indoor Photography',
    { ...TS6, x: 10.787, y: 3.564, w: 2.213, h: 0.32 });
  s.addText('Other Equipment',
    { ...TS6, x: 10.787, y: 4.124, w: 2.213, h: 0.32 });
  s.addText('Technically Complex',
    { ...TS6, x: 10.787, y: 4.68, w: 2.213, h: 0.32 });
  s.addText('Landscape Photography',
    { ...TS6, x: 10.787, y: 5.186, w: 2.546, h: 0.32 });
  s.addText('Editing Photo',
    { ...TS6, x: 10.787, y: 5.687, w: 2.213, h: 0.32 });
  globe(s, 6.53, 3.74, 1.92);
}

function slide34(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 8.245, y: 0, w: 5.088, h: 7.5, rotate: 180, fill: { color: PAPER } });
  s.addShape('custGeom', { x: 6.204, y: 4.21, w: 5.802, h: 2.026, rotate: 180, fill: { color: RED_SHADOW }, points: [{ x: 2.901, y: 2.026, moveTo: true }, { x: 0, y: 1.306 }, { x: 2.916, y: 0 }, { x: 5.802, y: 1.306 }, { x: 2.901, y: 2.026 }, { close: true }] });
  s.addShape('custGeom', { x: 6.929, y: 3.451, w: 4.351, h: 1.682, rotate: 180, fill: { color: RED_SHADOW }, points: [{ x: 2.176, y: 1.682, moveTo: true }, { x: 0, y: 1.137 }, { x: 2.143, y: 0 }, { x: 4.351, y: 1.137 }, { x: 2.176, y: 1.682 }, { close: true }] });
  s.addShape('custGeom', { x: 7.65, y: 2.7, w: 2.904, h: 1.309, rotate: 180, fill: { color: RED_SHADOW }, points: [{ x: 1.449, y: 1.309, moveTo: true }, { x: 0, y: 0.947 }, { x: 1.423, y: 0 }, { x: 2.904, y: 0.943 }, { x: 1.449, y: 1.309 }, { close: true }] });
  s.addShape('custGeom', { x: 8.385, y: 1.944, w: 1.434, h: 1.013, rotate: 180, fill: { color: RED_SHADOW }, points: [{ x: 0.714, y: 1.013, moveTo: true }, { x: 0, y: 0.835 }, { x: 0.708, y: 0 }, { x: 1.434, y: 0.836 }, { x: 0.714, y: 1.013 }, { close: true }] });
  s.addShape('custGeom', { x: 5.679, y: 4.93, w: 6.852, h: 1.531, fill: { color: RED_MID }, points: [{ x: 0.525, y: 0, moveTo: true }, { x: 3.426, y: 0.7 }, { x: 6.327, y: 0 }, { x: 6.852, y: 0.676 }, { x: 3.426, y: 1.531 }, { x: 0, y: 0.676 }, { x: 0.525, y: 0 }, { close: true }] });
  s.addShape('custGeom', { x: 6.384, y: 3.996, w: 5.442, h: 1.344, fill: { color: RED_MID }, points: [{ x: 4.897, y: 0, moveTo: true }, { x: 5.442, y: 0.702 }, { x: 2.721, y: 1.344 }, { x: 0, y: 0.702 }, { x: 0.545, y: 0 }, { x: 2.721, y: 0.513 }, { close: true }] });
  s.addShape('custGeom', { x: 7.122, y: 3.066, w: 3.967, h: 1.175, fill: { color: RED_MID }, points: [{ x: 3.438, y: 0, moveTo: true }, { x: 3.967, y: 0.68 }, { x: 1.983, y: 1.175 }, { x: 0, y: 0.68 }, { x: 0.528, y: 0 }, { x: 1.983, y: 0.336 }, { close: true }] });
  icon(s, 'aperture', 8.563, 1.19, 1.085, RED_MID);
  s.addShape('custGeom', { x: 7.845, y: 2.121, w: 2.521, h: 1.01, fill: { color: RED_MID }, points: [{ x: 1.98, y: 0, moveTo: true }, { x: 2.521, y: 0.696 }, { x: 1.26, y: 1.01 }, { x: 0, y: 0.696 }, { x: 0.541, y: 0 }, { x: 1.26, y: 0.18 }, { close: true }] });
  icon(s, 'aperture', 9.105, 1.19, 1.941, RED_DEEP);
  s.addShape('custGeom', { x: 9.105, y: 3.066, w: 1.983, h: 1.175, fill: { color: RED_DEEP }, points: [{ x: 1.455, y: 0, moveTo: true }, { x: 1.983, y: 0.68 }, { x: 0, y: 1.175 }, { x: 0, y: 1.175 }, { x: 0, y: 0.336 }, { x: 0, y: 0.336 }, { x: 1.455, y: 0 }, { close: true }] });
  s.addShape('custGeom', { x: 9.105, y: 3.996, w: 2.721, h: 1.344, fill: { color: RED_DEEP }, points: [{ x: 2.176, y: 0, moveTo: true }, { x: 2.721, y: 0.702 }, { x: 0, y: 1.344 }, { x: 0, y: 0.513 }, { x: 2.176, y: 0 }, { close: true }] });
  s.addShape('custGeom', { x: 9.105, y: 4.93, w: 3.426, h: 1.531, fill: { color: RED_DEEP }, points: [{ x: 2.901, y: 0, moveTo: true }, { x: 3.426, y: 0.676 }, { x: 0, y: 1.531 }, { x: 0, y: 0.7 }, { x: 2.901, y: 0 }, { close: true }] });
  s.addText('02',
    { ...TS4, x: 8.804, y: 2.439, w: 0.672, h: 0.488 });
  s.addText('01',
    { ...TS4, x: 8.804, y: 1.44, w: 0.596, h: 0.488 });
  s.addText('03',
    { ...TS4, x: 8.804, y: 3.547, w: 0.682, h: 0.488 });
  s.addText('04',
    { ...TS4, x: 8.804, y: 4.656, w: 0.7, h: 0.488 });
  s.addText('05',
    { ...TS4, x: 8.804, y: 5.723, w: 0.693, h: 0.488 });
  s.addText([{ text: '3D ' }, { text: 'Pyramid', options: { bold: true } }, { text: ' with 5 Level for Photography Data' }],
    { x: 0.604, y: 1.19, w: 4.644, h: 1.043, valign: 'top', fontSize: 28, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 0.721, y: 2.564, w: 1.455, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('Creative with Technique',
    { ...TS12, x: 0.604, y: 2.974, w: 3.558, h: 0.337 });
  s.addText(lorem(30, false),
    { ...TS2, x: 0.604, y: 3.31, w: 4.644, h: 0.897 });
  s.addText('Indoor Photography',
    { ...TS12, x: 0.604, y: 4.655, w: 3.558, h: 0.337 });
  s.addText(lorem(30, false),
    { ...TS2, x: 0.604, y: 4.991, w: 4.644, h: 0.897 });
}

function slide35(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 5.088, h: 7.5, rotate: 180, fill: { color: PAPER } });
  s.addShape('roundRect', { x: 0.642, y: 5.209, w: 1.996, h: 0.826, flipH: true, fill: { color: RED_BRIGHT }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 2.262, y: 4.054, w: 1.996, h: 0.826, flipH: true, fill: { color: RED_C8 }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 3.881, y: 2.9, w: 1.996, h: 0.826, flipH: true, fill: { color: RED_9A }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 5.5, y: 1.745, w: 1.996, h: 0.826, flipH: true, fill: { color: RED }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 2.828, y: 6.114, w: 5.01, h: 0.826, fill: { color: RED_BRIGHT }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 4.446, y: 4.959, w: 5.01, h: 0.826, fill: { color: RED_C8 }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 6.063, y: 3.786, w: 5.01, h: 0.826, fill: { color: RED_9A }, rectRadius: 0.413 });
  s.addShape('roundRect', { x: 7.681, y: 2.65, w: 5.01, h: 0.826, fill: { color: RED }, rectRadius: 0.413 });
  s.addText(lorem(25, false),
    { ...TS7, x: 3.758, y: 6.114, w: 3.867, h: 0.825 });
  s.addText(lorem(25, false),
    { ...TS7, x: 5.375, y: 4.96, w: 3.867, h: 0.825 });
  s.addText(lorem(25, false),
    { ...TS7, x: 6.993, y: 3.805, w: 3.867, h: 0.825 });
  s.addText(lorem(25, false),
    { ...TS7, x: 8.611, y: 2.651, w: 3.867, h: 0.825 });
  icon(s, 'tripod', 1.142, 5.292, 0.648, WHITE);
  icon(s, 'aperture', 2.651, 4.198, 0.515, WHITE);
  icon(s, 'camera', 4.231, 2.992, 0.603, WHITE);
  icon(s, 'umbrella', 5.978, 1.826, 0.607, WHITE);
  s.addText([{ text: 'Four step prosess for good ' }, { text: 'photography', options: { bold: true } }],
    { x: 0.388, y: 0.491, w: 4.135, h: 1.919, valign: 'top', fontSize: 36, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 3.578, y: 2.111, w: 1.042, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  footer(s, 'right');
}

function slide36(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 6.3, y: 0.32, w: 7.033, h: 6.86, rotate: 180, fill: { color: RED } });
  s.addText([{ text: 'Panorama', options: { bold: true } }, { text: ' histogram chart sales analystic' }],
    { x: 1.012, y: 1.607, w: 4.644, h: 1.043, valign: 'top', fontSize: 28, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 1.129, y: 2.981, w: 1.455, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('Data Overview',
    { ...TS12, x: 1.012, y: 3.39, w: 3.558, h: 0.337 });
  s.addText(lorem(30, false),
    { ...TS2, x: 1.012, y: 3.727, w: 4.644, h: 0.897 });
  s.addText('73+',
    { ...TS11, x: 1.129, y: 5.061, w: 0.952, h: 0.639 });
  s.addText('“ Lorem Ipsum is simply',
    { ...TS8, x: 1.129, y: 5.74, w: 1.48, h: 0.252 });
  s.addText('198+',
    { ...TS11, x: 3.317, y: 5.061, w: 1.161, h: 0.639 });
  s.addText('“ Lorem Ipsum is simply',
    { ...TS8, x: 3.317, y: 5.74, w: 1.48, h: 0.252 });

  // Pareto: clustered column plus a cumulative-percentage line (chartEx in the original)
  // Pareto sorts the categories from largest to smallest, hence this order
  const paretoCats = ['Innovative', 'Professional', 'Creative', 'Interactive'];
  s.addChart([
    { type: pptx.charts.BAR,  data: [{ name: 'Score', labels: [paretoCats], values: [79, 70, 65, 30] }] },
    { type: pptx.charts.LINE, data: [{ name: 'Cumulative', labels: [paretoCats], values: [32, 61, 88, 100] }],
      options: { secondaryValAxis: true, secondaryCatAxis: true, chartColors: [SILVER], lineDataSymbol: 'none', lineSize: 2 } }
  ], {
    x: 6.733, y: 1.634, w: 6.307, h: 4.233,
    chartColors: [SILVER, RED_BRIGHT, SILVER, SILVER],
    barGapWidthPct: 15, showLegend: false, valAxisMaxVal: 90, valAxisMinVal: 0,
    catAxisLabelColor: SNOW, catAxisLabelFontFace: UI, catAxisLabelFontSize: 12, catAxisLineShow: false,
    valAxisLabelColor: SNOW, valAxisLabelFontFace: UI, valAxisLabelFontSize: 12, valAxisLineShow: false,
    valAxisMajorUnit: 10, valGridLine: { color: '8E3232', style: 'solid', size: 1 },
    valAxes: [
      { valAxisMaxVal: 90, valAxisMinVal: 0, valAxisMajorUnit: 10, valAxisLabelColor: SNOW,
        valAxisLabelFontFace: UI, valAxisLabelFontSize: 12, valGridLine: { color: '8E3232', style: 'solid', size: 1 } },
      { valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 10, valAxisLabelColor: SNOW,
        valAxisLabelFontFace: UI, valAxisLabelFontSize: 12, valAxisLabelFormatCode: '0"%"',
        valGridLine: { style: 'none' } }
    ],
    catAxes: [
      { catAxisLabelColor: SNOW, catAxisLabelFontFace: UI, catAxisLabelFontSize: 12 },
      { catAxisHidden: true, catAxisLabelColor: RED }
    ],
    plotArea: { fill: { color: RED } }, chartArea: { fill: { color: RED } }
  });
}

function slide37(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0.32, w: 6.667, h: 6.86, rotate: 180, fill: { color: RED } });
  s.addText(lorem(41),
    { ...TS1, x: 7.665, y: 3.409, w: 4.324, h: 1.175 });
  s.addText([{ text: 'Sales yearly stock chart ' }, { text: 'analysis', options: { bold: true } }],
    { x: 7.665, y: 1.475, w: 4.455, h: 1.313, valign: 'top', fontSize: 36, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 7.783, y: 3.089, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addShape('rect', { x: 7.665, y: 4.985, w: 4.455, h: 0.98, rotate: 180, fill: { color: RED } });
  icon(s, 'tripod', 7.97, 5.151, 0.648, WHITE);
  s.addText(lorem(21),
    { x: 8.546, y: 5.023, w: 3.575, h: 0.897, valign: 'top', lineSpacingMultiple: 1.5, fontSize: 11, fontFace: SANS, color: SILVER_LT });

  // Volume columns with a Close price line on a secondary axis
  const stockCats = ['1/5/2002', '1/6/2002', '1/7/2002', '1/8/2002', '1/9/2002'];
  s.addChart([
    { type: pptx.charts.BAR,  data: [{ name: 'Volume', labels: [stockCats], values: [70, 120, 150, 135, 148] }] },
    { type: pptx.charts.LINE, data: [{ name: 'Close', labels: [stockCats], values: [25, 38, 50, 35, 43] }],
      options: { secondaryValAxis: true, secondaryCatAxis: true, chartColors: [SILVER_LT], lineDataSymbol: 'none', lineSize: 2 } }
  ], {
    x: 0.214, y: 1.284, w: 6.238, h: 4.933,
    chartColors: [SILVER], barGapWidthPct: 150,
    showLegend: true, legendPos: 'b', legendColor: WHITE, legendFontFace: UI, legendFontSize: 12,
    catAxisLabelColor: WHITE, catAxisLabelFontFace: UI, catAxisLabelFontSize: 12,
    valAxes: [
      { valAxisMaxVal: 160, valAxisMinVal: 0, valAxisMajorUnit: 20, valAxisLabelColor: WHITE,
        valAxisLabelFontFace: UI, valAxisLabelFontSize: 12, valGridLine: { color: '8E3232', style: 'solid', size: 1 } },
      { valAxisMaxVal: 6, valAxisMinVal: 0, valAxisMajorUnit: 1, valAxisLabelColor: WHITE,
        valAxisLabelFontFace: UI, valAxisLabelFontSize: 12, valGridLine: { style: 'none' } }
    ],
    catAxes: [
      { catAxisLabelColor: WHITE, catAxisLabelFontFace: UI, catAxisLabelFontSize: 12 },
      { catAxisHidden: true, catAxisLabelColor: RED }
    ],
    plotArea: { fill: { color: RED } }, chartArea: { fill: { color: RED } }
  });
}

function slide38(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0.55, y: 1.52, w: 3.86, h: 5.3, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 4.737, y: 1.52, w: 3.86, h: 5.3, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 8.923, y: 1.52, w: 3.86, h: 5.3, rotate: 180, fill: { color: RED } });
  s.addShape('roundRect', { x: 1.23, y: 5.535, w: 2.5, h: 0.89, rotate: 180, fill: { color: SNOW }, rectRadius: 0.148 });
  s.addShape('roundRect', { x: 5.417, y: 5.535, w: 2.5, h: 0.89, rotate: 180, fill: { color: SNOW }, rectRadius: 0.148 });
  s.addShape('roundRect', { x: 9.603, y: 5.535, w: 2.5, h: 0.89, rotate: 180, fill: { color: SNOW }, rectRadius: 0.148 });
  s.addShape('custGeom', { x: 8.448, y: 1.575, w: 2.578, h: 0.853, rotate: 146.47, fill: { color: SNOW }, points: [{ x: 0, y: 0.174, moveTo: true }, { x: 2.578, y: 0 }, { x: 2, y: 0.839 }, { x: 1.056, y: 0.853 }, { x: 0, y: 0.174 }, { close: true }] });
  s.addText('Pricing Tables',
    { x: 5.096, y: 0.26, w: 3.141, h: 0.639, valign: 'top', align: 'center', fontSize: 32, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 6.035, y: 1.129, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('$ 250',
    { x: 1.465, y: 2.265, w: 2.03, h: 0.842, valign: 'top', align: 'center', fontSize: 44, fontFace: UI_SB, color: WHITE });
  s.addText('$ 350',
    { x: 5.652, y: 2.265, w: 2.03, h: 0.842, valign: 'top', align: 'center', fontSize: 44, fontFace: UI_SB, color: WHITE });
  s.addText('$ 400',
    { x: 9.838, y: 2.265, w: 2.03, h: 0.842, valign: 'top', align: 'center', fontSize: 44, fontFace: UI_SB, color: WHITE });
  s.addText('Per Month',
    { x: 1.465, y: 3.107, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI_LT, color: SILVER_LT });
  s.addText('Per Month',
    { x: 5.652, y: 3.107, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI_LT, color: SILVER_LT });
  s.addText('Per Month',
    { x: 9.838, y: 3.107, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI_LT, color: SILVER_LT });
  s.addShape('line', { x: 0.971, y: 3.75, w: 3.018, h: 0, flipH: true, line: { color: SILVER, width: 0.75 } });
  s.addShape('line', { x: 5.158, y: 3.758, w: 3.018, h: 0, flipH: true, line: { color: SILVER, width: 0.75 } });
  s.addShape('line', { x: 9.344, y: 3.758, w: 3.018, h: 0, flipH: true, line: { color: SILVER, width: 0.75 } });
  s.addText(lorem(19, false),
    { ...TS3, x: 0.971, y: 4.12, w: 3.018, h: 0.897 });
  s.addText(lorem(19, false),
    { ...TS3, x: 5.096, y: 4.12, w: 3.018, h: 0.897 });
  s.addText(lorem(19, false),
    { ...TS3, x: 9.346, y: 4.12, w: 3.018, h: 0.897 });
  s.addText('GET STARTED',
    { x: 1.465, y: 5.793, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI, color: INK_60, charSpacing: 1 });
  s.addText('GET STARTED',
    { x: 5.652, y: 5.793, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI, color: INK_60, charSpacing: 1 });
  s.addText('GET STARTED',
    { x: 9.838, y: 5.793, w: 2.03, h: 0.37, valign: 'top', align: 'center', fontSize: 16, fontFace: UI, color: INK_60, charSpacing: 1 });
  s.addText('RECOMENDED',
    { x: 8.563, y: 1.889, w: 2.03, h: 0.303, rotate: 323.15, valign: 'top', align: 'center', fontSize: 12, fontFace: UI, color: INK_60, charSpacing: 1 });
}

function slide39(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 4.94, w: 13.333, h: 2.56, rotate: 180, fill: { color: RED } });
  s.addShape('rect', { x: 0.797, y: 3.589, w: 11.74, h: 3.28, rotate: 180, fill: { color: MIST } });
  s.addText('Pleace feel free to call us on ( phone cell ) or contact us by ( email ), if you require any further information.',
    { ...TS1, x: 1.325, y: 5.546, w: 4.655, h: 0.62 });
  s.addText([{ text: 'Contact ' }, { text: 'Us', options: { bold: true } }],
    { x: 1.325, y: 4.292, w: 4.455, h: 0.707, valign: 'top', fontSize: 36, fontFace: SERIF, color: INK });
  s.addShape('line', { x: 1.443, y: 5.226, w: 1.263, h: 0, flipH: true, line: { color: RED_DK, width: 0.75 } });
  s.addText('Office Hours',
    { ...TS14, x: 7.02, y: 3.887, w: 1.68, h: 0.337 });
  s.addText([{ text: 'Monday – Friday', options: { breakLine: true } }, { text: '08:00 – 16:00 ' }],
    { ...TS2, x: 7.02, y: 4.223, w: 1.973, h: 0.62 });
  s.addText('Our Address',
    { ...TS14, x: 9.79, y: 3.887, w: 1.68, h: 0.337 });
  s.addText('5678 Town Square City Parkway, CA 56789 United State',
    { ...TS2, x: 9.79, y: 4.223, w: 1.973, h: 0.897 });
  s.addText('Get In Touch ',
    { ...TS14, x: 6.987, y: 5.327, w: 1.68, h: 0.337 });
  s.addText([{ text: '( 01 ) 56788292001', options: { breakLine: true } }, { text: '( 01 ) 10029288756', options: { breakLine: true } }, { text: '( 0271 ) 54321' }],
    { ...TS2, x: 6.987, y: 5.663, w: 2.311, h: 0.897 });
  s.addText('Follow Us',
    { ...TS14, x: 9.79, y: 5.331, w: 1.68, h: 0.337 });
  s.addText([{ text: '@PanoramaStd.', options: { breakLine: true } }, { text: 'email@panorama.net', options: { breakLine: true } }, { text: 'www.panorama-std.net' }],
    { ...TS2, x: 9.789, y: 5.668, w: 2.311, h: 0.897 });
}

function slide40(pptx) {
  const s = pptx.addSlide();
  corner(s, 10.614, 0.545, 1.246, 0.934, false, false, WHITE, 1);
  corner(s, 6.001, 0.545, 1.246, 0.934, true, false, WHITE, 1);
  corner(s, 10.614, 4.531, 1.246, 0.934, false, true, WHITE, 1);
  corner(s, 6.001, 4.531, 1.246, 0.934, true, true, WHITE, 1);
  s.addText('THANK YOU',
    { x: 2.869, y: 2.189, w: 6.516, h: 1.447, valign: 'top', wrap: false, fontSize: 80, fontFace: SERIF, color: WHITE });
  s.addText('Thank You For Your Attention',
    { x: 4.283, y: 3.573, w: 3.435, h: 0.37, valign: 'top', wrap: false, fontSize: 16, fontFace: UI, color: SNOW });
  s.addText('PLACEHOLDER',
    { x: -1.239, y: 1.45, w: 3.405, h: 0.505, rotate: 270, valign: 'top', fontSize: 12, fontFace: SANS, color: WHITE });
  s.addText([{ text: 'Created by NewCoral', options: { breakLine: true } }, { text: '2020' }],
    { x: 10.468, y: 6.691, w: 2.632, h: 0.639, valign: 'top', align: 'right', fontSize: 16, fontFace: UI_SB, color: WHITE });
}

// ---------------------------------------------------------------- build
const builders = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40];
builders.forEach(function (fn) { fn(pptx); });
pptx.writeFile({ fileName: path.join(__dirname, '0bc7d3c2-dec0-4eab-80d7-846a8988dc0c_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
