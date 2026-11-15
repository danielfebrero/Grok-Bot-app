/**
 * "Recharging the Mind & Body" — yoga / wellness deck, rebuilt with pptxgenjs.
 *
 *   node 09265ec3-223f-4247-b5b0-a27167afa683_grok_final.js
 *
 * Slide size 13.333 x 7.5 in (16:9). Raster photographs of the original are not
 * embedded; they are drawn as labelled placeholder rectangles.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '262626'; // tx1 lum 85% — headings / dark body
const GREEN = '17361C'; // accent1 — brand dark green
const GREEN_DK = '112815'; // accent1 lum 75%
const GREEN_SOFT = '67BE75'; // accent1 lum 50/50 — the soft green glow
const PEACH = 'FFEBE0'; // accent2
const PEACH_DK = 'FFCCB0'; // accent2 lum 90% — hairline outlines
const PEACH_LT = 'FFEBE1'; // accent2 lum 98/2 — the soft peach glow
const GREY = '808080'; // bg1 lum 50% — paragraph text
const GREY_LT = 'A6A6A6'; // bg1 lum 65% — nav bar
const GREY_FT = 'B8B8B8'; // bg2 lum 75% — footer url
const WHITE = 'FFFFFF';

const DISPLAY = 'Marcellus'; // theme major font
const BODY = 'Tenor Sans'; // theme minor font

const HAIRLINE = { color: PEACH_DK, width: 0.75 };
const CARD_SHADOW = { type: 'outer', color: '000000', opacity: 0.1, blur: 17, offset: 4, angle: 90 };

/* ------------------------------------------------------------------ helpers */

/** Plain text box: top-anchored, no fill, like every text box in the source deck. */
function text(slide, body, opts) {
  slide.addText(body, Object.assign({ fontFace: BODY, color: INK, valign: 'top' }, opts));
}

/** Paragraph copy: 11 pt, 1.5 line spacing, grey. */
function para(slide, body, x, y, w, h, color) {
  text(slide, body, { x, y, w, h, fontSize: 11, color: color || GREY, lineSpacingMultiple: 1.5 });
}

/** Section / slide heading in the display face. */
function heading(slide, body, x, y, w, h, size) {
  text(slide, body, { x, y, w, h, fontSize: size, fontFace: DISPLAY, color: INK });
}

/**
 * Soft radial glow, faked with nested ellipses whose stacked alphas ramp
 * linearly from 0 at the rim to `peak` in the middle (pptxgenjs has no
 * gradient fills). `RINGS` concentric ovals share the same bounding box centre.
 */
const RINGS = 22;
// `transparency` is an integer percent and each layer is rounded to 8-bit when
// composited, so a deep stack always lands a little darker than requested.
const STACK_GAIN = 1.12;

function glow(slide, x, y, w, h, color, peak) {
  let covered = 0; // opacity actually accumulated so far, rim -> centre
  for (let i = 0; i < RINGS; i++) {
    const scale = 1 - i / RINGS;
    const target = ((peak / STACK_GAIN) * (i + 1)) / RINGS;
    const transparency = Math.round(100 - (100 * (target - covered)) / (1 - covered));
    covered += ((100 - transparency) / 100) * (1 - covered);
    slide.addShape('ellipse', {
      x: x + (w * (1 - scale)) / 2,
      y: y + (h * (1 - scale)) / 2,
      w: w * scale,
      h: h * scale,
      fill: { color, transparency },
      line: { type: 'none' },
    });
  }
}

/** custGeom point list for a rectangle with per-corner radii (inches). */
function roundedPath(w, h, r) {
  const tl = r.tl || 0;
  const tr = r.tr || 0;
  const br = r.br || 0;
  const bl = r.bl || 0;
  const arc = (rad, stAng) => ({ x: 0, y: 0, curve: { type: 'arc', hR: rad, wR: rad, stAng, swAng: 90 } });
  const pts = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }];
  if (tr) pts.push(arc(tr, 270));
  pts.push({ x: w, y: h - br });
  if (br) pts.push(arc(br, 0));
  pts.push({ x: bl, y: h });
  if (bl) pts.push(arc(bl, 90));
  pts.push({ x: 0, y: tl });
  if (tl) pts.push(arc(tl, 180));
  pts.push({ close: true });
  return pts;
}

/** Arch = rectangle whose two top corners are rounded (the deck's signature shape). */
function arch(slide, o) {
  slide.addShape('custGeom', {
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h,
    rotate: o.rotate,
    points: roundedPath(o.w, o.h, { tl: o.r, tr: o.r }),
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line || { type: 'none' },
    shadow: o.shadow,
  });
}

/** Rounded rectangle using the source deck's `adj` value (0..50000). */
function roundRect(slide, o) {
  slide.addShape('roundRect', {
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h,
    rectRadius: (o.adj / 100000) * Math.min(o.w, o.h),
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line || { type: 'none' },
    shadow: o.shadow,
  });
}

/** Stand-in for a photograph that is not embedded in this rebuild. */
function imagePlaceholder(slide, x, y, w, h, label) {
  slide.addShape('rect', { x, y, w, h, fill: { color: 'FAFAFA' }, line: { color: 'E2E2E2', width: 0.75 } });
  text(slide, label || '[image]', {
    x,
    y: y + h / 2 - 0.2,
    w,
    h: 0.4,
    fontSize: 12,
    color: 'AAAAAA',
    align: 'center',
    valign: 'middle',
  });
}

/* -------------------------------------------------------------------- icons */
/* Small line/solid glyphs built from native shapes (0.2"–0.35" on the slide). */

function iconHeart(slide, x, y, s, color, filled) {
  slide.addShape('heart', {
    x: x + s * 0.11,
    y: y + s * 0.11,
    w: s * 0.84,
    h: s * 0.76,
    fill: filled ? { color } : { type: 'none' },
    line: { color, width: 1.5 },
  });
}

function iconHeartPulse(slide, x, y, w, h, color) {
  iconHeart(slide, x, y, Math.min(w, h), color, false);
  slide.addShape('line', { x, y: y + h * 0.45, w, h: 0, line: { color, width: 1.5 } });
}

function iconPerson(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x: x + w * 0.28, y, w: w * 0.4, h: w * 0.4, fill: { type: 'none' }, line: { color, width: 1.5 } });
  slide.addShape('round1Rect', {
    x,
    y: y + h * 0.5,
    w,
    h: h * 0.4,
    rectRadius: h * 0.18,
    fill: { type: 'none' },
    line: { color, width: 1.5 },
  });
}

function iconArrow(slide, x, y, s, color) {
  slide.addShape('line', {
    x,
    y: y + s / 2,
    w: s,
    h: 0,
    line: { color, width: 1.25, endArrowType: 'triangle' },
  });
}

function iconChat(slide, x, y, s, color) {
  slide.addShape('ellipse', { x: x + s * 0.08, y, w: s * 0.92, h: s * 0.92, fill: { type: 'none' }, line: { color, width: 1.25 } });
  slide.addShape('triangle', { x, y: y + s * 0.7, w: s * 0.26, h: s * 0.3, rotate: 200, fill: { color }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + s * 0.34, y: y + s * 0.27, w: s * 0.34, h: s * 0.38, fill: { color }, line: { type: 'none' } });
}

function iconGlobe(slide, x, y, s, color) {
  slide.addShape('ellipse', { x, y, w: s, h: s, fill: { type: 'none' }, line: { color, width: 1.25 } });
  slide.addShape('ellipse', { x: x + s * 0.31, y, w: s * 0.38, h: s, fill: { type: 'none' }, line: { color, width: 1 } });
  slide.addShape('line', { x, y: y + s / 2, w: s, h: 0, line: { color, width: 1 } });
  slide.addShape('line', { x: x + s * 0.06, y: y + s * 0.25, w: s * 0.88, h: 0, line: { color, width: 0.75 } });
  slide.addShape('line', { x: x + s * 0.06, y: y + s * 0.75, w: s * 0.88, h: 0, line: { color, width: 0.75 } });
}

function iconThumbUp(slide, x, y, s, color) {
  slide.addShape('round2SameRect', { x: x + s * 0.32, y: y + s * 0.18, w: s * 0.68, h: s * 0.68, fill: { color }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: x + s * 0.35, y, w: s * 0.22, h: s * 0.45, rectRadius: s * 0.11, fill: { color }, line: { type: 'none' } });
  slide.addShape('roundRect', { x, y: y + s * 0.4, w: s * 0.26, h: s * 0.46, rectRadius: s * 0.07, fill: { color }, line: { type: 'none' } });
}

function iconHeadset(slide, x, y, w, h, color) {
  // headband arc + two ear cups + mic boom
  slide.addShape('blockArc', {
    x: x + w * 0.1, y, w: w * 0.8, h: h * 0.95,
    angleRange: [180, 0], arcThicknessRatio: 0.34,
    fill: { color }, line: { type: 'none' },
  });
  slide.addShape('roundRect', { x, y: y + h * 0.4, w: w * 0.22, h: h * 0.45, rectRadius: w * 0.1, fill: { color }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: x + w * 0.78, y: y + h * 0.4, w: w * 0.22, h: h * 0.45, rectRadius: w * 0.1, fill: { color }, line: { type: 'none' } });
  slide.addShape('rect', { x: x + w * 0.5, y: y + h * 0.85, w: w * 0.33, h: h * 0.13, fill: { color }, line: { type: 'none' } });
}

function iconMail(slide, x, y, w, h, color) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: h * 0.12, fill: { type: 'none' }, line: { color, width: 1.25 } });
  slide.addShape('line', { x, y, w: w / 2, h: h * 0.62, line: { color, width: 1.25 } });
  slide.addShape('line', { x: x + w / 2, y: y + h * 0.62, w: w / 2, h: -h * 0.62, line: { color, width: 1.25 } });
}

/* ------------------------------------------------------- repeated furniture */

/** Brand mark: interlocking "b" and "q" rings around a small four-point star. */
function logo(slide) {
  const ring = { w: 0.22, h: 0.22, fill: { type: 'none' }, line: { color: GREEN, width: 1.25 } };
  const stem = { w: 0.022, h: 0.155, fill: { color: GREEN }, line: { type: 'none' } };
  slide.addShape('ellipse', Object.assign({ x: 0.54, y: 0.377 }, ring));
  slide.addShape('ellipse', Object.assign({ x: 0.637, y: 0.377 }, ring));
  slide.addShape('rect', Object.assign({ x: 0.643, y: 0.377 }, stem));
  slide.addShape('rect', Object.assign({ x: 0.741, y: 0.442 }, stem));
  slide.addShape('ellipse', { x: 0.652, y: 0.439, w: 0.096, h: 0.096, fill: { color: WHITE }, line: { color: GREEN, width: 1 } });
  slide.addShape('star4', { x: 0.667, y: 0.454, w: 0.066, h: 0.066, fill: { color: GREEN }, line: { type: 'none' } });
}

/** Brand mark, footer rule, page number and url — present on every slide. */
function chrome(slide, num, showUrl) {
  logo(slide);
  slide.addShape('line', { x: 0, y: 7.01, w: 10.7, h: 0, line: { color: PEACH_DK, width: 1 } });
  text(slide, num, { x: 12.481, y: 0.366, w: 0.446, h: 0.337, fontSize: 14 });
  if (showUrl !== false) {
    text(slide, 'www.yourcompany.com', {
      x: 10.642,
      y: 6.831,
      w: 2.286,
      h: 0.303,
      fontSize: 12,
      fontFace: DISPLAY,
      color: GREY_FT,
      align: 'right',
    });
  }
}

/** Peach + green background glows, positioned per layout. */
function backdrop(slide, peachBox, greenBox) {
  glow(slide, peachBox[0], peachBox[1], peachBox[2], peachBox[3], PEACH_LT, 0.95);
  glow(slide, greenBox[0], greenBox[1], greenBox[2], greenBox[3], GREEN_SOFT, 0.4);
}

/* ------------------------------------------------------------ slide content */

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunta labore dolore.';
const LOREM_BIO = 'Lorem ipsum dolor adispicing consectetur adipis elite, eiusmod in es dolore.';
const QUOTE = '\u201cLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.\u201d';

function slide01(pres) {
  const s = pres.addSlide();
  backdrop(s, [5.911, 1.531, 3.635, 4.039], [2.764, 0, 7.003, 7.003]);
  chrome(s, '01');

  s.addShape('ellipse', { x: 2.63, y: 5.29, w: 1.485, h: 1.485, rotate: 345, fill: { color: PEACH_DK }, line: { type: 'none' } });

  heading(s, 'Recharging the Mind & Body', 1.126, 1.935, 4.123, 2.827, 54);
  heading(s, 'Yoga as a tool for resilience, calm, and clarity', 1.126, 1.513, 4.493, 0.337, 14);
  para(s, LOREM_SHORT, 9.018, 2.792, 3.189, 0.626);

  arch(s, { x: 12.296, y: 4.973, w: 1.3, h: 0.929, rotate: 270, r: 0.057, line: HAIRLINE });

  // top navigation
  const nav = [
    ['About', 8.012, 0.975],
    ['Services', 8.867, 0.975],
    ['Contact', 9.846, 1.057],
    ['Testimonials', 10.835, 1.346],
  ];
  nav.forEach(([label, x, w]) => text(s, label, { x, y: 0.395, w, h: 0.303, fontSize: 12, color: GREY_LT }));

  // "Explore Now" pill
  roundRect(s, { x: 9.103, y: 3.73, w: 2.428, h: 0.511, adj: 9770, fill: GREEN, line: { color: GREEN, width: 0.75 } });
  text(s, 'Explore Now', { x: 9.179, y: 3.817, w: 1.835, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: WHITE });

  // tilted "Join Now" badge
  s.addShape('ellipse', { x: 2.689, y: 5.349, w: 1.367, h: 1.367, rotate: 345, fill: { color: GREEN }, line: { type: 'none' } });
  text(s, 'Join Now', {
    x: 2.903, y: 5.679, w: 0.94, h: 0.707, rotate: 345,
    fontSize: 18, fontFace: DISPLAY, color: WHITE, align: 'center',
  });
}

function slide02(pres) {
  const s = pres.addSlide();
  backdrop(s, [0, 3.461, 3.635, 4.039], [6.479, 0.988, 5.202, 5.202]);
  chrome(s, '02');

  heading(s, 'Embracing Inner Peace', 1.549, 1.447, 5.318, 1.582, 44);
  arch(s, { x: 8.817, y: -0.287, w: 2.827, h: 6.294, rotate: 270, r: 1.4135, line: HAIRLINE });
  arch(s, { x: 9.803, y: 3.146, w: 2.362, h: 4.787, rotate: 270, r: 1.181, line: HAIRLINE });

  [
    ['500+', 'Participants Energized', 3.362],
    ['9.3%', 'Boost in Well-being', 5.031],
  ].forEach(([figure, label, y]) => {
    text(s, figure, { x: 1.549, y, w: 1.499, h: 0.505, fontSize: 24, color: GREEN });
    text(s, label, { x: 1.549, y: y + 0.492, w: 3.194, h: 0.337, fontSize: 14 });
    para(s, LOREM_MED, 1.549, y + 0.847, 4.84, 0.626);
  });
}

function slide03(pres) {
  const s = pres.addSlide();
  backdrop(s, [9.698, 0.703, 3.635, 4.039], [0.406, 0.208, 6.207, 6.207]);
  chrome(s, '03');

  heading(s, 'Breath, Stretch, and Connect', 7.708, 1.729, 5.318, 1.582, 44);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 7.708, 3.814, 4.52, 0.904);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod incididunt ut et dolore.', 7.708, 4.818, 4.52, 0.626);

  arch(s, { x: 3.862, y: 2.406, w: 3.065, h: 5.167, r: 1.5325, line: HAIRLINE });
  arch(s, { x: 0.722, y: 1.253, w: 3.065, h: 6.32, r: 1.5325, line: HAIRLINE });
}

function slide04(pres) {
  const s = pres.addSlide();
  backdrop(s, [0, 1.265, 3.635, 4.039], [7.063, 0.76, 6.207, 6.207]);
  chrome(s, '04');

  heading(s, 'Yoga for Everyday', 1.282, 1.554, 5.607, 0.841, 44);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididun labore dolore magna aliqua. enim ad minim veniam, quis nostrud.', 7.032, 1.559, 4.805, 0.904);

  roundRect(s, { x: 1.44, y: 3.004, w: 4.977, h: 1.719, adj: 7374, line: HAIRLINE });
  roundRect(s, { x: 6.992, y: 3.004, w: 4.977, h: 1.719, adj: 7374, line: HAIRLINE });

  iconHeartPulse(s, 1.482, 5.098, 0.29, 0.261, GREEN);
  iconPerson(s, 7.063, 5.098, 0.239, 0.265, GREEN);

  heading(s, 'Personalized Yoga Sessions', 1.828, 5.06, 4.204, 0.337, 14);
  heading(s, 'Mind-Body Wellness Programs', 7.378, 5.063, 4.204, 0.337, 14);
  para(s, LOREM_CARD, 1.405, 5.502, 4.485, 0.627);
  para(s, LOREM_CARD, 6.959, 5.502, 4.485, 0.627);
}

function slide05(pres) {
  const s = pres.addSlide();
  backdrop(s, [0.018, 3.461, 3.635, 4.039], [4.058, 0.079, 4.493, 4.493]);
  chrome(s, '05');

  arch(s, { x: 1.068, y: 1.519, w: 2.848, h: 5.981, r: 1.424, fill: GREEN });
  arch(s, { x: 4.041, y: 1.519, w: 2.848, h: 5.981, r: 1.424, fill: WHITE, line: HAIRLINE, shadow: CARD_SHADOW });
  s.addShape('ellipse', { x: 4.288, y: 1.791, w: 2.354, h: 2.354, fill: { type: 'none' }, line: HAIRLINE });

  heading(s, 'The Faces Behind Our Success', 7.632, 2.124, 5.131, 1.447, 40);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Enim minim veniam, quis nostrud exercitation ullamco laboris aliquip.', 7.632, 3.946, 4.391, 1.175);
  text(s, 'Learn More', {
    x: 7.632, y: 5.387, w: 1.604, h: 0.303,
    fontSize: 12, fontFace: DISPLAY, color: GREEN, underline: { style: 'sng' },
  });
  iconArrow(s, 8.901, 5.424, 0.229, GREEN);

  // two instructor cards: the left one sits on the dark arch, the right on the white arch
  const people = [
    { name: 'Lucia Vanka', role: 'Yoga Instructor', x: 1.473, bioX: 1.234, nameY: 4.258, fg: WHITE, sub: WHITE, icon: 2.218 },
    { name: 'Rubika Hilea', role: 'Wellness Program', x: 4.446, bioX: 4.207, nameY: 4.253, fg: INK, sub: GREY, icon: 5.067 },
  ];
  people.forEach((p) => {
    text(s, p.name, { x: p.x, y: p.nameY, w: 2.039, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: p.fg, align: 'center' });
    text(s, p.role, { x: p.x, y: 4.539, w: 2.039, h: 0.371, fontSize: 12, color: p.sub, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, LOREM_BIO, { x: p.bioX, y: 5.241, w: 2.516, h: 0.897, fontSize: 11, color: p.sub, align: 'center', lineSpacingMultiple: 1.5 });
    iconChat(s, p.icon, 6.361, 0.214, p.fg === WHITE ? WHITE : GREEN);
    iconGlobe(s, p.icon + 0.336, 6.361, 0.214, p.fg === WHITE ? WHITE : GREEN);
  });
}

function slide06(pres) {
  const s = pres.addSlide();
  backdrop(s, [0, 1.024, 4.155, 4.616], [6.83, 1.319, 4.751, 5.411]);
  chrome(s, '06');

  heading(s, 'Celebrity Unity in Stillness', 1.08, 1.599, 7.031, 0.774, 40);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 1.08, 2.741, 5.711, 0.626);

  arch(s, { x: 9.868, y: 0.55, w: 2.379, h: 4.64, rotate: 270, r: 0.1454, line: HAIRLINE });
  arch(s, { x: 0.542, y: 3.496, w: 2.379, h: 3.596, rotate: 270, r: 0.1308, line: HAIRLINE });
  roundRect(s, { x: 3.588, y: 4.104, w: 2.587, h: 2.379, adj: 5885, line: HAIRLINE });
  roundRect(s, { x: 6.231, y: 4.104, w: 5.711, h: 2.379, adj: 5885, line: HAIRLINE });
}

function slide07(pres) {
  const s = pres.addSlide();
  backdrop(s, [0.038, 3.43, 3.635, 4.039], [5.528, 0.703, 6.207, 6.207]);
  chrome(s, '07');

  heading(s, 'Our Clients Have Said', 1.257, 1.979, 4.138, 1.582, 44);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim minim veniam.', 1.303, 5.25, 4.334, 0.897);
  iconThumbUp(s, 1.427, 4.617, 0.315, GREEN);
  heading(s, 'Great Feedback That Inspires', 1.83, 4.632, 3.381, 0.337, 14);

  // dark testimonial card
  roundRect(s, { x: 6.349, y: 1.974, w: 2.833, h: 3.504, adj: 3373, fill: GREEN });
  iconHeart(s, 6.667, 2.21, 0.35, WHITE);
  text(s, 'Sima Lucida', { x: 6.62, y: 2.738, w: 2.353, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: WHITE });
  text(s, 'Business Manager', { x: 6.62, y: 3.147, w: 2.353, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: WHITE });
  text(s, QUOTE, { x: 6.621, y: 3.556, w: 2.441, h: 0.897, fontSize: 11, italic: true, color: WHITE, lineSpacingMultiple: 1.5 });
  roundRect(s, { x: 6.727, y: 4.697, w: 0.528, h: 0.37, adj: 16667, fill: WHITE });
  text(s, '9.7', { x: 6.749, y: 4.731, w: 0.483, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: GREEN_DK, align: 'center' });
  text(s, 'Meditation Guide', { x: 7.255, y: 4.727, w: 1.718, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: WHITE });

  // light testimonial card
  roundRect(s, { x: 9.596, y: 2.77, w: 2.833, h: 3.504, adj: 6431, fill: WHITE, line: HAIRLINE, shadow: CARD_SHADOW });
  iconHeart(s, 9.938, 2.994, 0.35, GREEN);
  text(s, 'Eva Angeline', { x: 9.862, y: 3.534, w: 2.353, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK });
  text(s, 'Supervisor', { x: 9.862, y: 3.943, w: 2.353, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: GREEN_DK });
  text(s, QUOTE, { x: 9.864, y: 4.352, w: 2.441, h: 0.897, fontSize: 11, italic: true, color: GREY, lineSpacingMultiple: 1.5 });
  roundRect(s, { x: 9.969, y: 5.542, w: 0.528, h: 0.37, adj: 16667, fill: GREEN });
  text(s, '9.4', { x: 9.992, y: 5.576, w: 0.483, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: WHITE, align: 'center' });
  text(s, 'Studio Coordinator', { x: 10.497, y: 5.576, w: 1.718, h: 0.303, fontSize: 12, fontFace: DISPLAY, color: INK });

  roundRect(s, { x: 11.08, y: 2.25, w: 1.049, h: 1.048, adj: 12243, line: HAIRLINE });
}

function slide08(pres) {
  const s = pres.addSlide();
  backdrop(s, [9.698, 2.625, 3.635, 4.039], [1.171, 0.647, 6.207, 6.207]);
  chrome(s, '08');

  imagePlaceholder(s, 0, 1.203, 5.711, 5.411, '[image] laptop mockup');

  heading(s, 'Cultivating a Mindful Culture', 6.667, 1.406, 5.918, 1.447, 40);
  text(s, '230+', { x: 6.705, y: 3.719, w: 1.958, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: GREEN });
  heading(s, 'Yoga Session Completed', 6.705, 4.1, 2.468, 0.337, 14);
  text(s, '100K', { x: 9.321, y: 3.719, w: 1.397, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK });
  heading(s, 'Certified Instructors', 9.322, 4.1, 2.265, 0.337, 14);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Veniam nostrud exercitation lae commodo consequat consectetur adipiscing tempor incididunt. ', 6.705, 4.721, 4.881, 1.179);
}

function slide09(pres) {
  const s = pres.addSlide();
  backdrop(s, [0, 3.461, 3.635, 4.039], [5.064, 0.12, 3.637, 3.637]);
  chrome(s, '09');

  roundRect(s, { x: 1.521, y: 1.19, w: 4.136, h: 5.754, adj: 50000, line: HAIRLINE });
  heading(s, 'Great Impressions, Long-Lasting Impact', 6.667, 1.488, 4.937, 2.121, 40);

  roundRect(s, { x: 6.155, y: 3.978, w: 5.94, h: 2.121, adj: 50000, fill: GREEN });
  s.addShape('line', { x: 8.108, y: 3.978, w: 0, h: 2.121, line: { color: WHITE, width: 1 } });
  iconHeadset(s, 7.096, 4.715, 0.338, 0.304, WHITE);
  text(s, '+123 456 789', { x: 6.511, y: 5.153, w: 1.508, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: WHITE, align: 'center' });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna.', 8.302, 4.586, 3.51, 0.904, WHITE);
}

function slide10(pres) {
  const s = pres.addSlide();
  backdrop(s, [0, 1.265, 3.635, 4.039], [0.674, 0.181, 6.207, 6.207]);
  chrome(s, '10', false);

  // green contact bar, right-hand corners rounded
  s.addShape('custGeom', {
    x: 0, y: 5.25, w: 7.977, h: 1.313,
    points: roundedPath(7.977, 1.313, { tr: 0.108, br: 0.108 }),
    fill: { color: GREEN }, line: { type: 'none' },
  });

  heading(s, 'Thank You', 0.935, 2.457, 6.382, 1.447, 80);
  text(s, 'For Your Attention.', { x: 1.145, y: 3.682, w: 5.736, h: 0.64, fontSize: 32, fontFace: DISPLAY, color: GREEN });
  text(s, 'Hi Folks, Get In Touch!', { x: 0.738, y: 5.738, w: 2.513, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: WHITE });
  iconMail(s, 3.792, 5.798, 0.284, 0.222, WHITE);
  text(s, 'yourcompany@gmail.com', { x: 4.126, y: 5.738, w: 3.281, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: WHITE });

  // hairline echo of the bar, slightly larger
  s.addShape('custGeom', {
    x: 0, y: 5.196, w: 8.029, h: 1.413,
    points: roundedPath(8.029, 1.413, { tr: 0.11, br: 0.11 }),
    fill: { type: 'none' }, line: { color: PEACH, width: 0.75 },
  });

  // big sweeping arc on the right edge
  s.addShape('custGeom', {
    x: 7.544, y: -0.174, w: 5.853, h: 7.854,
    points: [
      { x: 0.337, y: 0, moveTo: true },
      { x: 5.853, y: 0 },
      { x: 5.853, y: 7.854 },
      { x: 5.587, y: 7.848 },
      { x: 0, y: 1.965, curve: { type: 'cubic', x1: 2.475, y1: 7.69, x2: 0, y2: 5.116 } },
      { x: 0.265, y: 0.213, curve: { type: 'cubic', x1: 0, y1: 1.355, x2: 0.093, y2: 0.767 } },
      { close: true },
    ],
    fill: { type: 'none' }, line: HAIRLINE,
  });
}

/* --------------------------------------------------------------------- main */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK_16x9', width: 13.333, height: 7.5 });
pres.layout = 'DECK_16x9';
pres.title = 'Recharging the Mind & Body';

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10].forEach((build) => build(pres));

pres
  .writeFile({ fileName: path.join(__dirname, '09265ec3-223f-4247-b5b0-a27167afa683_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
