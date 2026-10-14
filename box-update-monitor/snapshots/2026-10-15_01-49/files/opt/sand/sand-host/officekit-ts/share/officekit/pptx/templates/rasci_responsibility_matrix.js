/*
 * RASCI Framework and Matrix — deck rebuilt with pptxgenjs.
 * Slide size 20 x 11.25 in (widescreen, 16:9 @ 2x).
 * Run:  node 0d797f4e-bac1-4ee8-8e3f-825db18e92e1_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const W = 20;          // slide width  (in)
const H = 11.25;       // slide height (in)

/* ------------------------------------------------------------------ *
 * Palette (theme "GraficsoTheme02")
 * ------------------------------------------------------------------ */
const RED = 'E74C40';   // accent1  Responsible
const ORANGE = 'F7952E'; // accent2  Accountable
const GREEN = '8BB749';  // accent3  Supportive
const TEAL = '5FB7A2';   // accent4
const CYAN = '17A7C0';   // accent5  Consulted
const BLUE = '1D6FA9';   // accent6  Informed
const NAVY = '44546A';   // dk2
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GRAY = '595959';   // text 65% luminance
const GRAY50 = '808080';

const ROLE5 = [RED, ORANGE, GREEN, CYAN, BLUE];   // R A S C I
const ROLE5T = [RED, ORANGE, GREEN, TEAL, BLUE];  // variant using accent4 for C

const FONT = 'Lato';      // theme major + minor
const BEBAS = 'Bebas Neue';
const MONT = 'Montserrat';
const MONT_BLACK = 'Montserrat Black';

/* PowerPoint lumMod/lumOff tinting, so tint tables stay out of the source. */
function tint(hex, lumMod, lumOff) {
  const n = parseInt(hex, 16);
  const r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
  let h = 0, s = 0;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    if (mx === r) h = ((g - b) / d + (g < b ? 6 : 0));
    else if (mx === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  const l2 = Math.max(0, Math.min(1, l * (lumMod / 100) + (lumOff || 0) / 100));
  const hue = (p, q, t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l2 < 0.5 ? l2 * (1 + s) : l2 + s - l2 * s, p = 2 * l2 - q;
  const to = v => Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase();
  return s === 0 ? to(l2) + to(l2) + to(l2) : to(hue(p, q, h + 1 / 3)) + to(hue(p, q, h)) + to(hue(p, q, h - 1 / 3));
}
const t80 = c => tint(c, 20, 80);   // very light
const t60 = c => tint(c, 40, 60);   // light
const t40 = c => tint(c, 60, 40);   // medium
const dk25 = c => tint(c, 75, 0);   // darker
const dk50 = c => tint(c, 50, 0);

/* ------------------------------------------------------------------ *
 * Small drawing helpers
 * ------------------------------------------------------------------ */

/** Solid banded emulation of a linear gradient (pptxgenjs has no gradient fill). */
function gradientBands(slide, c1, c2, angleDeg, box) {
  const b = Object.assign({ x: 0, y: 0, w: W, h: H }, box || {});
  const n = 44;
  const a = angleDeg * Math.PI / 180, dx = Math.cos(a), dy = Math.sin(a);
  const len = Math.abs(b.w * dx) + Math.abs(b.h * dy);   // extent along gradient
  const across = Math.abs(b.w * dy) + Math.abs(b.h * dx); // extent across it
  const cx = b.x + b.w / 2, cy = b.y + b.h / 2, step = len / n;
  const mix = (p) => {
    const A = parseInt(c1, 16), B = parseInt(c2, 16), o = [];
    for (let s = 16; s >= 0; s -= 8) o.push(Math.round(((A >> s) & 255) * (1 - p) + ((B >> s) & 255) * p));
    return o.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  for (let i = 0; i < n; i++) {
    const t = (i + 0.5) / n, off = (t - 0.5) * len;
    slide.addShape('rect', {
      x: cx + off * dx - step / 2, y: cy + off * dy - across / 2,
      w: step + 0.02, h: across, rotate: angleDeg,
      fill: { color: mix(t) }, line: { type: 'none' },
    });
  }
}

/** Rounded pill with a left-to-right colour blend (used for the RASCI badges). */
function gradientPill(slide, x, y, w, h, c1, c2, text, fontSize) {
  const r = h / 2;
  slide.addShape('ellipse', { x, y, w: h, h, fill: { color: c1 }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + w - h, y, w: h, h, fill: { color: c2 }, line: { type: 'none' } });
  gradientBands(slide, c1, c2, 0, { x: x + r, y, w: w - h, h });
  slide.addText(text, { x, y, w, h, align: 'center', valign: 'middle', fontFace: FONT, fontSize, bold: true, color: WHITE });
}

/** Grid of small dots — replaces the halftone dot-pattern images. */
function dotGrid(slide, x, y, cols, rows, color, step, dia) {
  step = step || 0.5; dia = dia || 0.14;
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++)
      slide.addShape('ellipse', {
        x: x + c * step, y: y + r * step, w: dia, h: dia,
        fill: { color: color || WHITE }, line: { type: 'none' },
      });
}

/** Diagonal hair-lines — replaces the striped texture images. */
function diagonalLines(slide, x, y, size, color, count) {
  count = count || 14;
  for (let i = 0; i < count; i++) {
    const off = (i + 1) * (size / (count + 1));
    slide.addShape('line', {
      x: x + off - size / 2, y: y, w: size, h: 0, rotate: -45,
      line: { color: color || WHITE, width: 1, transparency: 55 },
    });
  }
}

/** Neutral stand-in for a raster image / icon from the source deck. */
function imagePlaceholder(slide, x, y, w, h, opt) {
  const o = opt || {};
  slide.addShape(o.round ? 'ellipse' : 'roundRect', {
    x, y, w, h, rectRadius: Math.min(w, h) * 0.12,
    fill: { color: o.fill || 'EFEFEF', transparency: o.transparency == null ? 0 : o.transparency },
    line: { color: o.line || o.fill || 'D0D0D0', width: 1, dashType: 'dash' },
  });
  if (o.label !== false) {
    slide.addText(o.label || '[image]', {
      x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center',
      fontFace: FONT, fontSize: Math.max(7, Math.min(11, w * 6)), color: o.labelColor || GRAY50,
    });
  }
}

/**
 * The line icon that appears next to every role in the deck: a person
 * (head + shoulders) framed by an L-shaped flow arrow, with a small emblem
 * standing in for the gear / checklist / speech-bubble variants.
 */
function roleIcon(slide, x, y, size, color) {
  const s = size, ink = { color: BLACK, width: 1.1 };
  // flow bracket: down the left, along the bottom, then up-and-over on the right
  slide.addShape('line', { x: x + 0.06 * s, y: y + 0.46 * s, w: 0, h: 0.42 * s, line: ink });
  slide.addShape('line', { x: x + 0.06 * s, y: y + 0.88 * s, w: 0.30 * s, h: 0, line: ink });
  slide.addShape('line', { x: x + 0.64 * s, y: y + 0.10 * s, w: 0.30 * s, h: 0, line: ink });
  slide.addShape('line', { x: x + 0.94 * s, y: y + 0.10 * s, w: 0, h: 0.30 * s, line: ink });
  slide.addShape('triangle', { x: x + 0.88 * s, y: y + 0.38 * s, w: 0.12 * s, h: 0.10 * s, rotate: 180, fill: { color: BLACK }, line: { type: 'none' } });
  // person
  slide.addShape('ellipse', { x: x + 0.46 * s, y: y + 0.26 * s, w: 0.26 * s, h: 0.26 * s, fill: { color: t40(color) }, line: ink });
  slide.addShape('blockArc', { x: x + 0.34 * s, y: y + 0.53 * s, w: 0.50 * s, h: 0.50 * s, fill: { color }, line: ink });
  // emblem (gear / list / bubble in the original artwork)
  slide.addShape('ellipse', { x: x + 0.16 * s, y: y + 0.26 * s, w: 0.24 * s, h: 0.24 * s, fill: { color: t60(color) }, line: ink });
  slide.addShape('ellipse', { x: x + 0.24 * s, y: y + 0.34 * s, w: 0.08 * s, h: 0.08 * s, fill: { color: WHITE }, line: { type: 'none' } });
}

/** Master furniture: rule, "Infograficso" wordmark and the two nav arrows. */
function footer(slide) {
  slide.addShape('line', { x: 0, y: 10.531, w: 1.297, h: 0, line: { color: BLACK, width: 5.6 } });
  slide.addText('Infograficso', {
    x: 1.462, y: 10.355, w: 1.6, h: 0.337, fontFace: FONT, fontSize: 14, bold: true, color: BLACK,
  });
  slide.addShape('leftArrow', { x: 17.75, y: 10.46, w: 0.30, h: 0.13, fill: { color: BLACK }, line: { type: 'none' } });
  slide.addShape('rightArrow', { x: 18.14, y: 10.46, w: 0.30, h: 0.13, fill: { color: BLACK }, line: { type: 'none' } });
}

/** Centred title + subtitle pair used by the "Title Slide" layout. */
function titleBlock(slide, title, subtitle, color) {
  slide.addText(title, {
    x: 2.5, y: 0.447, w: 15, h: 0.909, align: 'center', valign: 'top',
    fontFace: FONT, fontSize: 48, bold: true, color: color || BLACK,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 2.5, y: 1.356, w: 15, h: 0.55, align: 'center', valign: 'top',
      fontFace: FONT, fontSize: 24, color: color || GRAY50,
    });
  }
}

/** Plain text box (top-left anchored like PowerPoint text boxes). */
function tb(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: FONT, fontSize: 16, color: GRAY, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  }, o));
}

/** Bulleted list with a blank line between items (as authored in the deck). */
function bullets(slide, items, o) {
  const runs = [];
  items.forEach((it, i) => {
    if (i) runs.push({ text: '', options: { breakLine: true, bullet: false } });
    runs.push({ text: it, options: { bullet: { indent: 20 }, breakLine: true } });
  });
  tb(slide, runs, o);
}

/** Five 0.1" dots showing which of the RASCI cards is active. */
function pagerDots(slide, x, y, activeIndex, color) {
  for (let i = 0; i < 5; i++)
    slide.addShape('ellipse', {
      x: x + i * 0.166, y, w: 0.1, h: 0.1,
      fill: { color: i === activeIndex ? color : tint(WHITE, 95) }, line: { type: 'none' },
    });
}

/** Isometric-style capital used for the deck's 3D RASCI letters. */
function letter3d(slide, x, y, w, h, ch, color) {
  const steps = 7, d = Math.min(w, h) * 0.16 / steps, size = h * 62;
  for (let i = steps; i > 0; i--) {
    slide.addText(ch, {
      x: x + i * d, y: y + i * d, w, h, align: 'center', valign: 'middle',
      fontFace: MONT, bold: true, fontSize: size, color: dk25(color),
    });
  }
  slide.addText(ch, { x, y, w, h, align: 'center', valign: 'middle', fontFace: MONT, bold: true, fontSize: size, color: t40(color) });
}

/* ------------------------------------------------------------------ *
 * Shared copy
 * ------------------------------------------------------------------ */
const ROLE_NAMES = ['Responsible', 'Accountable', 'Supportive', 'Consulted', 'Informed'];
const RASCI = ['R', 'A', 'S', 'C', 'I'];

const SUB_MODEL = 'The RASCI model ensure effective collaboration and accountability within the team or organization.';
const SUB_TOOL = 'A project management tool used to clarify roles and responsibilities within a team.';

// One-line summary per role (used on slides 6, 11, 13, 15, 17)
const ROLE_LEAD = [
  'This person or group is responsible for completing the task',
  'This person ultimately answerable for task success.',
  'These are the individuals who assist in task execution.',
  'These are the individuals provides expertise or input.',
  'These are the individuals kept updated on task progress.',
];

// Three short bullets per role (slides 6, 11, 15, 17)
const ROLE_BULLETS = [
  ['Executes assigned tasks.', 'Ensures completion to standards.', 'Takes ownership of outcomes.'],
  ['Ultimately responsible for success.', 'Provides oversight and direction.', 'Ensures alignment with goals.'],
  ['Assists in task execution.', 'Provides resources and assistance.', 'Helps overcome obstacles.'],
  ['Offers expertise and input.', 'Collaborates for informed decisions.', 'Provides guidance as needed.'],
  ['Kept updated on progress.', 'Provides feedback as necessary.', 'Receives regular project updates.'],
];

// Longer descriptions (slides 10, 14)
const ROLE_LONG = [
  'This person is responsible for executing the task or activity. They are the "doer" and are directly involved in the completion of the task.',
  'Individuals who need to stay informed about the progress and outcome of a task without being directly involved in its execution are categorized as "Informed" in the RASCI model.',
  'Individuals in this role provide assistance and resources to the person who is responsible for the task. They support the execution but are not directly responsible for the outcome.',
  'Individuals in the "Consulted" role provide input and expertise before decisions are made or actions are taken. Their feedback is crucial, but they are not responsible for the final decision or execution.',
  'Individuals who need to be kept in the loop or informed about the progress and outcome of the task. They are not directly involved in the execution, but they need to stay informed about the developments.',
];

const FRAMEWORK_BLURB = 'The RASCI framework is a tool used in project management and organizational development to clarify roles and responsibilities within a team or group.';

/* Five-column card geometry reused by many slides */
const COL_X = [1.576, 5.001, 8.426, 11.851, 15.276];
const COL_W = 3.15;
const TXT_X = [1.773, 5.198, 8.623, 12.048, 15.472];
const TXT_W = 2.756;

/* ------------------------------------------------------------------ *
 * Slide 1 — cover
 * ------------------------------------------------------------------ */
function slide01(pptx) {
  const s = pptx.addSlide();
  coverBackdrop(s);

  const anim = [
    ['A', 6.92, 3.046, 0.632, RED], ['N', 7.883, 3.309, 0.518, ORANGE],
    ['I', 8.73, 2.94, 0.30, GREEN], ['M', 9.296, 3.309, 0.63, TEAL],
    ['A', 10.225, 2.94, 0.632, CYAN], ['T', 11.063, 3.309, 0.484, t60(BLUE)],
    ['E', 11.794, 3.309, 0.436, dk25(RED)], ['D', 12.525, 3.046, 0.554, GREEN],
  ];
  anim.forEach(([ch, x, y, w, c]) => s.addText(ch, {
    x, y, w, h: 0.525, align: 'center', valign: 'middle',
    fontFace: MONT_BLACK, bold: true, fontSize: 52, color: c,
  }));

  s.addText('RASCI', { x: 6.002, y: 3.986, w: 8.033, h: 2.794, align: 'center', valign: 'middle', fontFace: MONT_BLACK, bold: true, fontSize: 160, charSpacing: 6, color: WHITE });
  s.addText('FRAMEWORK AND MATRIX', { x: 5.844, y: 6.863, w: 8.312, h: 0.64, align: 'center', valign: 'middle', fontFace: MONT, bold: true, fontSize: 32, charSpacing: 6, color: WHITE });

  ribbon(s, 7.27, 8.42, 5.45, 0.99, 'Free updates', 40);
}

/** Cyan→blue diagonal wash plus the deck's floating circles and textures. */
function coverBackdrop(s) {
  gradientBands(s, CYAN, BLUE, 45);
  diagonalLines(s, 0.4, 0.4, 3.9, WHITE, 16);                 // top-left hatch
  s.addShape('ellipse', { x: 5.107, y: -1.152, w: 3, h: 3, fill: { color: t40(ORANGE) }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 5.257, y: -1.002, w: 2.7, h: 2.7, fill: { color: ORANGE }, line: { type: 'none' } });
  s.addShape('ellipse', { x: -0.787, y: 4.677, w: 2.121, h: 2.121, fill: { color: RED }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 5.731, y: 9.814, w: 2.121, h: 2.121, fill: { color: GREEN }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 17.518, y: 9.515, w: 1.359, h: 1.359, fill: { color: t60(CYAN) }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 17.518, y: 4.056, w: 2.894, h: 2.894, fill: { color: dk50(RED) }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 17.668, y: 4.206, w: 2.594, h: 2.594, fill: { color: RED }, line: { type: 'none' } });
  [[15.591, -1.522, 3], [16.19, -0.923, 1.801], [-1.411, 7.874, 3], [-0.812, 8.474, 1.801]]
    .forEach(([x, y, d]) => s.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: WHITE, width: 1 } }));
  dotGrid(s, 14.351, 8.773, 5, 3, WHITE, 0.581, 0.174);
  dotGrid(s, 2.297, 7.732, 5, 2, WHITE, 0.698, 0.209);
  dotGrid(s, 13.834, 2.273, 5, 1, WHITE, 0.581, 0.174);
  diagonalLines(s, 11.353, 9.576, 2.121, WHITE, 9);
}

/** Red banner whose left and right ends are notched inwards. */
function ribbon(s, x, y, w, h, text, size) {
  const cap = h * 0.55;
  s.addShape('chevron', { x, y, w: cap * 2, h, fill: { color: RED }, line: { type: 'none' } });
  s.addShape('chevron', { x: x + w - cap * 2, y, w: cap * 2, h, rotate: 180, fill: { color: RED }, line: { type: 'none' } });
  s.addShape('rect', { x: x + cap, y, w: w - cap * 2, h, fill: { color: RED }, line: { type: 'none' } });
  s.addText(text, { x, y, w, h, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: size, color: WHITE });
}

/* ------------------------------------------------------------------ *
 * Slide 2 — five white cards, 3D letters on top
 * ------------------------------------------------------------------ */
function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: 'F8F8F8' };
  titleBlock(s, 'RASCI Model', SUB_MODEL);

  const copy = [
    ['Takes ownership of completing the task or activity.', 'Executes the necessary actions to accomplish the task.', 'Ensures that the work is done according to established standards and timelines.'],
    ['Ultimately answerable for the success or failure of the task.', 'Provides direction and guidance to the responsible person or team.', 'Makes decisions and takes responsibility for the outcomes of the task.'],
    ['Assists the responsible person or team in completing the task.', 'Provides resources, tools, or assistance as needed.', 'Helps overcome obstacles or challenges encountered during task execution.'],
    ['Provides expertise, insights, or advice related to the task.', 'Offers input during the decision-making process.', 'Collaborates with the responsible person or team to ensure the task is completed effectively.'],
    ['Kept up-to-date on the progress, decisions, and outcomes of the task.', "Receives regular updates or reports regarding the task's status.", 'May provide feedback on the task, but is not directly involved in its execution.'],
  ];
  ROLE_NAMES.forEach((name, i) => {
    s.addShape('roundRect', { x: COL_X[i], y: 3.463, w: COL_W, h: 6.101, rectRadius: 0.31, fill: { color: WHITE }, line: { type: 'none' } });
    letter3d(s, COL_X[i] + 0.85, 2.75, 1.4, 1.2, RASCI[i], ROLE5T[i]);
    s.addText(name, { x: TXT_X[i], y: 4.35, w: TXT_W, h: 0.404, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: BLACK });
    bullets(s, copy[i], { x: TXT_X[i], y: 4.991, w: TXT_W, h: 3.9 });
    pagerDots(s, COL_X[i] + 1.193, 9.253, i, ROLE5T[i]);
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 3 — company-event ring diagram
 * ------------------------------------------------------------------ */
function slide03(pptx) {
  const s = pptx.addSlide();
  s.addText([
    { text: 'RASCI Model for a ', options: { color: BLACK } },
    { text: 'Organizing a Company Event', options: { color: RED } },
  ], { x: 1.579, y: 1.184, w: 7.194, h: 2.423, fontFace: FONT, fontSize: 48, bold: true, valign: 'top', margin: 0 });

  // five ring segments arranged around the hub
  const rings = [
    [9.0, 2.743, BLUE, 0], [6.654, 4.295, RED, -72], [7.47, 6.91, ORANGE, -144],
    [10.12, 6.91, GREEN, 144], [11.225, 4.295, CYAN, 72],
  ];
  rings.forEach(([x, y, c, rot]) => {
    s.addShape('blockArc', { x, y, w: 2.3, h: 2.3, rotate: rot, fill: { color: c }, line: { type: 'none' } });
    s.addShape('ellipse', { x: x + 0.45, y: y + 0.45, w: 1.4, h: 1.4, fill: { color: 'F0F0F0' }, line: { color: WHITE, width: 3 } });
  });
  [[7.10, 4.74, RED], [11.44, 4.74, CYAN], [7.92, 7.36, ORANGE], [9.30, 3.19, BLUE], [10.57, 7.36, GREEN]]
    .forEach(([x, y, c]) => roleIcon(s, x + 0.25, y + 0.25, 0.9, c));

  s.addShape('ellipse', { x: 8.525, y: 4.717, w: 2.949, h: 2.95, fill: { color: 'F0F0F0' }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 8.73, y: 4.922, w: 2.539, h: 2.539, fill: { color: 'FAFAFA' }, line: { color: WHITE, width: 2 } });
  roleIcon(s, 9.515, 5.179, 0.969, NAVY);
  s.addText('RASCI\nModel', { x: 9.0, y: 6.232, w: 2, h: 0.909, align: 'center', fontFace: FONT, fontSize: 24, bold: true, color: NAVY, lineSpacing: 28 });

  // callouts: [head, name, colour, body, x, y, align]
  const callouts = [
    ['[I] Informed', 'Sales Team', BLUE, 'Informed about the event details to ensure alignment with sales strategies and opportunities for client engagement.', 12.015, 1.701, 'left'],
    ['[C] Consulted', 'Public Relations (PR) Specialist', CYAN, "Consulted for input on the event's publicity strategy, media coverage, and overall public image.", 13.924, 4.502, 'left'],
    ['[S] Supportive', 'Administrative Assistant', GREEN, 'Provides support to the Event Coordinator by handling administrative tasks, such as scheduling and communication.', 13.346, 7.717, 'left'],
    ['Responsible [R]', 'Event Coordinator', RED, 'Responsible for planning, coordinating logistics, and executing the details of the event.', 1.927, 4.502, 'right'],
    ['Accountable [A]', 'Marketing Director', ORANGE, "Accountable for the overall success of the event, ensuring it aligns with the company's goals and brand image.", 2.654, 7.717, 'right'],
  ];
  callouts.forEach(([head, who, c, body, x, y, align]) => {
    s.addText([
      { text: head, options: { bold: true, color: BLACK, breakLine: true } },
      { text: who, options: { italic: true, underline: true, color: c } },
    ], { x, y, w: 4.0, h: 1.03, align, fontFace: FONT, fontSize: 20, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
    tb(s, body, { x, y: y + 1.16, w: 4.0, h: 1.4, align });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 4 — five tinted cards
 * ------------------------------------------------------------------ */
function slide04(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', SUB_MODEL);
  const tints = [ORANGE, GREEN, TEAL, CYAN, BLUE];
  const copy = [
    ['Executes the tasks or activities assigned to them.', 'Completes the work according to specifications and deadlines.', 'Takes ownership of the outcome and quality of the work performed.'],
    ['Ultimately answerable for the success or failure of the project or task.', 'Provides oversight and direction to those responsible for executing the work.', 'Ensures that the project or task aligns with organizational goals and objectives.'],
    ['Assists the responsible party in executing their tasks effectively.', 'Provides resources, tools, or assistance necessary for task completion.', 'Helps address obstacles or challenges encountered during the project or task.'],
    ['Provides expertise, knowledge, or input relevant to the project or task.', 'Offers advice, guidance, or recommendations based on their area of expertise.', 'Collaborates with the accountable and responsible parties to ensure informed decision-making.'],
    ['Kept informed about the progress, decisions, and outcomes of the project or task.', 'Receives updates, reports, or communications regarding project developments.', 'Provides feedback or input as necessary but is not directly involved in project execution.'],
  ];
  ROLE_NAMES.forEach((name, i) => {
    gradientBands(s, t80(tints[i]), WHITE, 90, { x: COL_X[i], y: 2.584, w: COL_W, h: 6.98 });
    s.addText(name, { x: COL_X[i] + 0.475, y: 2.807, w: 2.2, h: 0.404, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: BLACK });
    roleIcon(s, COL_X[i] + 1.05, 3.3, 1.05, tints[i]);
    bullets(s, copy[i], { x: TXT_X[i], y: 4.644, w: TXT_W, h: 4.4 });
    pagerDots(s, COL_X[i] + 1.193, 9.3, i, tints[i]);
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 5 — website-launch teardrop diagram
 * ------------------------------------------------------------------ */
function slide05(pptx) {
  const s = pptx.addSlide();
  s.addText([
    { text: 'RASCI\nModel for a ', options: { color: BLACK } },
    { text: 'New Website Launching', options: { color: RED } },
  ], { x: 1.579, y: 1.439, w: 4.541, h: 3.231, fontFace: FONT, fontSize: 48, bold: true, valign: 'top', margin: 0, lineSpacing: 54 });

  s.addText([
    { text: 'In this scenario:', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    ...[['R (Responsible):', ' Web Developer'], ['A (Accountable):', ' Project Manager'],
    ['S (Supportive):', ' Graphic Designer'], ['C (Consulted):', ' UX Designer'],
    ['I (Informed):', ' Marketing Team']].flatMap(([b, t]) => ([
      { text: b, options: { bold: true } }, { text: t, options: { breakLine: true } }])),
  ], { x: 1.579, y: 6.716, w: 4.22, h: 2.356, fontFace: FONT, fontSize: 20, color: BLACK, valign: 'top', margin: 0 });

  // petals + knobs, clockwise from the top
  const petals = [
    [7.813, 1.688, 1.811, 2.377, RED, 315, 8.119, 2.559, 'R', 8.452, 2.806],
    [9.559, 2.807, 1.981, 1.981, ORANGE, 0, 9.865, 3.282, 'A', 10.182, 3.514],
    [10.282, 4.723, 2.377, 1.811, GREEN, 45, 10.588, 5.029, 'S', 10.961, 5.292],
    [9.559, 6.469, 1.981, 1.981, TEAL, 90, 9.865, 6.775, 'C', 10.195, 7.047],
    [7.813, 7.192, 1.811, 2.377, CYAN, 135, 8.119, 7.498, 'I', 8.539, 7.804],
  ];
  petals.forEach(([x, y, w, h, c, rot, kx, ky, ch, tx, ty]) => {
    s.addShape('teardrop', { x, y, w, h, rotate: rot, fill: { color: c }, line: { type: 'none' } });
    s.addShape('ellipse', { x: kx, y: ky, w: 1.2, h: 1.2, fill: { color: 'F2F2F2' }, line: { color: WHITE, width: 2 } });
    s.addText(ch, { x: tx, y: ty, w: 0.6, h: 0.707, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 36, bold: true, color: c });
  });

  s.addShape('ellipse', { x: 6.596, y: 3.505, w: 4.239, h: 4.24, fill: { color: WHITE, transparency: 70 }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 7.576, y: 4.486, w: 2.278, h: 2.278, fill: { color: 'F2F2F2' }, line: { color: WHITE, width: 2 } });
  roleIcon(s, 8.231, 4.704, 0.969, NAVY);
  s.addText('RASCI\nMODEL', { x: 7.715, y: 5.723, w: 2, h: 0.707, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: NAVY, lineSpacing: 22 });

  const rows = [
    ['Responsible : ', 'Web Developer', 'Responsible for coding and developing the new website. ensuring the implementation of functionalities.', 9.882, 1.105],
    ['Accountable : ', 'Project Manager', 'Accountable for the overall success of the website launch, ensuring timelines are met, and coordinating various tasks.', 11.968, 3.034],
    ['Supportive : ', 'Graphics Designer', 'Supports the Web Developer by creating visually appealing graphics and images for the website.', 13.009, 4.962],
    ['Consulted : ', 'UI/UX Designer', "Consulted for input on the website's user interface and overall user experience to enhance usability.", 11.968, 6.891],
    ['Informed : ', 'Marketing Team', 'Informed about the progress of the website development to plan marketing strategies and campaigns aligned with the launch.', 9.882, 8.819],
  ];
  rows.forEach(([lead, who, body, x, y]) => {
    s.addText([
      { text: lead, options: { bold: true } },
      { text: who, options: { bold: true, italic: true, underline: true } },
    ], { x, y, w: 5.906, h: 0.404, fontFace: FONT, fontSize: 18, color: BLACK, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
    tb(s, body, { x, y: y + 0.393, w: 6.0, h: 0.7 });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 6 — five solid colour columns
 * ------------------------------------------------------------------ */
function slide06(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', SUB_TOOL);
  const labels = ['responsible', 'accountable', 'supportive', 'confirmed', 'informed'];
  const leads = [ROLE_LEAD[0], ROLE_LEAD[1], ROLE_LEAD[2], ROLE_LEAD[3], ROLE_LEAD[4]];
  ROLE5.forEach((c, i) => {
    const x = 1.575 + i * 3.37;
    s.addShape('round2DiagRect', { x, y: 3.343, w: 3.37, h: 6.355, rectRadius: 0, fill: { color: c }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + 0.897, y: 2.587, w: 1.575, h: 1.575, rectRadius: 0.21, fill: { color: WHITE, transparency: 75 }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + 1.019, y: 2.709, w: 1.332, h: 1.332, rectRadius: 0.18, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
    s.addText(RASCI[i].toLowerCase(), { x: x + 1.019, y: 2.789, w: 1.332, h: 1.332, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: 96, color: c });
    s.addText(labels[i], { x: x + 0.394, y: 4.515, w: 2.583, h: 0.572, align: 'center', fontFace: BEBAS, fontSize: 28, color: WHITE });
    tb(s, leads[i], { x: TXT_X[i], y: 5.231, w: TXT_W, h: 0.95, color: WHITE });
    bullets(s, ROLE_BULLETS[i], { x: TXT_X[i], y: 6.354, w: TXT_W, h: 2.3, color: WHITE });
    roleIcon(s, x + 0.117, 8.855, 0.79, tint(c, 60, 40));
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 7 — arrow flow
 * ------------------------------------------------------------------ */
function slide07(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Framework', SUB_TOOL);
  s.addText("Let's consider a project to develop a new mobile application for a restaurant: This example illustrates how the RASCI model can be applied to a real-world project, ensuring clear roles, responsibilities, and communication channels for all involved stakeholders.",
    { x: 3.802, y: 2.338, w: 12.397, h: 1.01, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: '0D0D0D', valign: 'top' });

  const notes = [
    [['Software developer: ', 'Writes code.'], ['Graphic designer: ', 'Designs UI.'], ['Quality assurance tester: ', 'Tests functionality.']],
    [['Project manager: ', 'Oversees project schedule and budget.'], ['Restaurant owner: ', 'Ensures mobile app meets business needs.'], ['Chief technology officer: ', 'Aligns technical aspects with company strategy.']],
    [['IT department: ', 'Assists with technical. Implementation.'], ['Marketing department: ', 'Provides guidance on marketing strategy.'], ['Customer support team: ', 'Resolves issues during launch and ongoing use.']],
    [['Branding consultant: ', 'Ensures brand consistency.'], ['Marketing team: ', 'Aligns content with marketing strategies.'], ['Head chef: ', 'Provides functionality requirements.']],
    [['Stakeholders', ' receive updates on progress and milestones.'], ['', 'They provide feedback to meet expectations.'], ['', 'They stay informed about project changes and decisions.']],
  ];
  ROLE5.forEach((c, i) => {
    const x = 1.547 + i * 3.4288;
    s.addShape('rightArrow', { x, y: 4.325, w: 3.15, h: 2.003, fill: { color: c }, line: { type: 'none' } });
    s.addShape('rightArrow', { x: x + 0.06, y: 4.4, w: 3.02, h: 1.86, fill: { color: tint(c, 115, 0) }, line: { type: 'none' } });
    s.addText(ROLE_NAMES[i], { x: x + 0.3, y: 5.05, w: 2.2, h: 0.5, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 18, color: WHITE });
    const runs = [];
    notes[i].forEach(([b, t], k) => {
      if (k) runs.push({ text: '', options: { breakLine: true } });
      if (b) runs.push({ text: b, options: { bold: true, breakLine: !t } });
      if (t) runs.push({ text: t, options: { breakLine: true } });
    });
    tb(s, runs, { x, y: 6.673, w: 3.15, h: 3.1, color: BLACK });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 8 — chevrons with oversized letters
 * ------------------------------------------------------------------ */
function slide08(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', SUB_MODEL);
  const cols = [ORANGE, GREEN, TEAL, CYAN, BLUE];
  const copy = [
    ['Executes tasks or activities assigned to them.', 'Completes work according to specified standards and deadlines.'],
    ['Ultimately answerable for the success or failure of the project or task.', 'Provides oversight and direction to those responsible for executing the work.'],
    ['Provides expertise, knowledge, or input relevant to the project or task.', 'Collaborates with the accountable and responsible parties to ensure informed decision-making.'],
    ['Assists the responsible party in executing their tasks effectively.', 'Provides resources, tools, or assistance necessary for task completion.'],
    ['Kept informed about the progress, decisions, and outcomes of the project or task.', 'Provides feedback or input as necessary but is not directly involved in project execution.'],
  ];
  ROLE_NAMES.forEach((name, i) => {
    const x = 1.575 + i * 3.37;
    s.addShape('chevron', { x: x + 2.515, y: 2.85, w: 2.341, h: 3.827, fill: { color: tint(WHITE, 94) }, line: { type: 'none' } });
    s.addShape('chevron', { x: x + 1.977, y: 2.453, w: 3.368, h: 4.635, fill: { color: WHITE }, line: { type: 'none' } });
    s.addShape('chevron', { x: x + 2.307, y: 3.512, w: 0.614, h: 2.543, fill: { color: t40(cols[i]) }, line: { type: 'none' } });
    s.addText(RASCI[i], { x: x + 1.6, y: 3.908, w: 1.71, h: 1.717, fontFace: MONT, bold: true, fontSize: 96, color: cols[i], valign: 'top' });
    roleIcon(s, x - 0.02, 4.411, 0.9, cols[i]);
    s.addText(name, { x, y: 6.115, w: 2.4, h: 0.337, fontFace: FONT, fontSize: 20, bold: true, color: BLACK, valign: 'top', margin: 0 });
    bullets(s, copy[i], { x, y: 6.894, w: TXT_W, h: 2.7, margin: 0 });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 9 — marketing campaign, gradient backdrop with tag shapes
 * ------------------------------------------------------------------ */
function slide09(pptx) {
  const s = pptx.addSlide();
  gradientBands(s, RED, BLUE, -45);
  titleBlock(s, 'RASCI Model', 'A team working on a marketing campaign for a new product launch', WHITE);

  s.addShape('roundRect', { x: 1.575, y: 3.929, w: 16.838, h: 5.853, rectRadius: 0.33, fill: { color: WHITE }, line: { type: 'none' } });

  const copy = [
    ['Marketing Specialist', 'The ', ' is responsible for creating the content for the advertising campaign.'],
    ['Marketing Manager', 'The ', ' ensures campaign success without creating content, overseeing strategy and outcomes.'],
    ['Graphic Designer', 'The ', ' supports the Marketing Specialist by creating visual elements for the campaign, such as images and graphics.'],
    ['Product Manager', 'The ', ' provides crucial input on key features for the campaign, aligning marketing with core selling points.'],
    ['Sales Team', 'The ', ' is informed about campaign progress to align efforts with the marketing message and anticipate customer inquiries.'],
  ];
  ROLE5.forEach((c, i) => {
    const x = 2.577 + i * 2.983;
    s.addShape('rect', { x: x + 1.9, y: 3.142, w: 0.72, h: 1.0, fill: { color: dk50(c) }, line: { type: 'none' } });
    s.addShape('homePlate', { x: x + 0.129, y: 5.302, w: 2.007, h: 1.23, rotate: 90, fill: { color: dk25(c) }, line: { type: 'none' } });
    s.addShape('rect', { x, y: 3.142, w: 2.253, h: 2.16, fill: { color: c }, line: { type: 'none' } });
    s.addText(ROLE_NAMES[i].toLowerCase(), { x, y: 3.423, w: 2.253, h: 0.505, align: 'center', fontFace: BEBAS, fontSize: 24, color: WHITE });
    s.addText(RASCI[i].toLowerCase(), { x, y: 3.871, w: 2.253, h: 1.582, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: 88, color: WHITE });
    roleIcon(s, x + 0.75, 5.5, 0.75, dk50(c));
    tb(s, [
      { text: copy[i][1] }, { text: copy[i][0], options: { underline: true } }, { text: copy[i][2] },
    ], { x: x - 0.12, y: 6.955, w: 2.5, h: 2.0, align: 'center' });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 10 — gauge cards over a navy band
 * ------------------------------------------------------------------ */
function slide10(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', 'Framework for defining and clarifying roles in project management and organizational development.');
  s.addShape('rect', { x: 0, y: 5.799, w: W, h: 5.451, fill: { color: NAVY }, line: { type: 'none' } });

  const cols = [RED, ORANGE, GREEN, TEAL, CYAN];
  const labels = ['responsible', 'accountable', 'supportive', 'consulted', 'informed'];
  cols.forEach((c, i) => {
    const x = 1.585 + i * 3.4225;
    s.addShape('blockArc', { x: x + 0.595, y: 2.266, w: 1.96, h: 1.96, fill: { color: c }, line: { type: 'none' } });
    s.addShape('roundRect', { x, y: 3.282, w: COL_W, h: 6.461, rectRadius: 0.16, fill: { color: tint(WHITE, 96) }, line: { color: WHITE, width: 4 } });
    s.addShape('ellipse', { x: x + 0.949, y: 2.629, w: 1.252, h: 1.252, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
    s.addText(RASCI[i], { x: x + 1.207, y: 2.78, w: 0.736, h: 0.909, align: 'center', fontFace: FONT, fontSize: 48, bold: true, color: c });
    roleIcon(s, x + 1.245, 4.23, 1.181, c);
    s.addText(labels[i], { x: x + 0.197, y: 5.752, w: TXT_W, h: 0.64, align: 'center', fontFace: BEBAS, fontSize: 32, color: BLACK });
    tb(s, ROLE_LONG[i], { x: x + 0.197, y: 6.489, w: TXT_W, h: 3.1, align: 'center' });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 11 — org-chart style cards on a radial wash
 * ------------------------------------------------------------------ */
function slide11(pptx) {
  const s = pptx.addSlide();
  gradientBands(s, RED, BLUE, -30);
  gradientPill(s, 7.5, 1.356, 5, 1, GREEN, CYAN, 'RASCI MODEL', 32);

  const rule = { color: tint(WHITE, 75), width: 1 };
  s.addShape('line', { x: 9.982, y: 2.292, w: 0, h: 0.668, line: rule });
  s.addShape('line', { x: 3.151, y: 2.975, w: 13.661, h: 0, line: rule });
  [3.151, 6.567, 9.982, 13.397, 16.812].forEach(x => s.addShape('line', { x, y: 2.975, w: 0, h: 0.766, line: rule }));

  const names = ['Responsible', 'Accountable', 'Supportive', 'Confirmed', 'Informed'];
  ROLE5T.forEach((c, i) => {
    const x = 1.576 + i * 3.4165;
    s.addShape('roundRect', { x, y: 3.695, w: COL_W, h: 6.0, rectRadius: 0.3, fill: { color: WHITE }, line: { color: t80(c), width: 1 } });
    s.addText(names[i], { x: x + 0.475, y: 4.084, w: 2.2, h: 0.438, align: 'center', fontFace: FONT, fontSize: 20, bold: true, color: BLACK });
    roleIcon(s, x + 0.982, 4.697, 1.181, ROLE5[i]);
    tb(s, ROLE_LEAD[i], { x: x + 0.21, y: 6.053, w: TXT_W, h: 0.95 });
    bullets(s, ROLE_BULLETS[i], { x: x + 0.21, y: 7.178, w: TXT_W, h: 2.3 });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 12 — 2x3 colour blocks around an intro column
 * ------------------------------------------------------------------ */
function slide12(pptx) {
  const s = pptx.addSlide();
  s.addText('RASCI Model', { x: 1.579, y: 1.184, w: 4.929, h: 0.808, fontFace: FONT, fontSize: 48, bold: true, color: BLACK, valign: 'top', margin: 0 });
  s.addText(FRAMEWORK_BLURB, { x: 1.579, y: 2.58, w: 5.118, h: 1.346, fontFace: FONT, fontSize: 20, bold: true, italic: true, color: GRAY, valign: 'top', margin: 0 });
  s.addText([
    { text: 'Example Project:', options: { bold: true, color: GRAY, breakLine: true } },
    { text: 'Launching a New Product', options: { bold: true, underline: true, color: GREEN } },
  ], { x: 1.579, y: 4.267, w: 5.118, h: 0.673, fontFace: FONT, fontSize: 20, valign: 'top', margin: 0 });

  const blocks = [
    ['r', 'RESPONSIBLE', GREEN, 7.195, 1.0, ['The graphic designer is responsible for designing the logo for the new product.', 'The product manager is responsible for developing product specifications.', 'The research analyst is responsible for conducting market research.'], 9.489, 1.435],
    ['a', 'ACCOUNTABLE', TEAL, 12.805, 1.0, ['The marketing department is accountable for the overall success of the product launch.', 'The project manager is accountable for ensuring that tasks are completed on time and within budget.'], 15.114, 1.435],
    ['s', 'SUPPORTIVE', CYAN, 1.575, 5.62, ['The project manager supports the graphic designer, product manager, and research analyst by providing resources and guidance as needed.', 'Other departments may support the project by providing data, insights, or assistance with tasks related to the product launch.'], 3.878, 6.072],
    ['c', 'CONSULTED', BLUE, 7.19, 5.62, ['The marketing department provides input on logo design.', 'Sales and engineering teams are offers insights on product specifications.', 'Sales and customer support teams are provide feedback on market research.'], 9.489, 6.072],
    ['I', 'INFORMED', NAVY, 12.805, 5.62, ['Stakeholders, including company executives, investors, and relevant departments, are kept informed about project progress, changes, and decisions to ensure alignment with business objectives and market needs.'], 15.114, 6.072],
  ];
  blocks.forEach(([ch, label, c, x, y, items, tx, ty]) => {
    s.addShape('round2DiagRect', { x, y, w: 5.61, h: 4.63, rectRadius: 0, fill: { color: c }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + 0.138, y: y + 0.144, w: 1.933, h: 1.933, rectRadius: 0.96, fill: { color: WHITE, transparency: 80 }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + 0.287, y: y + 0.293, w: 1.634, h: 1.634, rectRadius: 0.82, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
    s.addText(ch, { x: x + 0.287, y: y + 0.36, w: 1.634, h: 1.634, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: 96, color: c });
    s.addText(label, { x: x + 0.12, y: y + 2.22, w: 1.969, h: 0.572, align: 'center', fontFace: BEBAS, fontSize: 28, color: WHITE });
    roleIcon(s, x + 0.6, y + 3.48, 0.85, WHITE);
    bullets(s, items, { x: tx, y: ty, w: 2.953, h: 3.9, color: WHITE });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 13 — dashed timeline of circles
 * ------------------------------------------------------------------ */
function slide13(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', SUB_TOOL);
  const dash = { color: GRAY, width: 1.5, dashType: 'dash' };
  s.addShape('roundRect', { x: 1.593, y: 3.544, w: 16.814, h: 6.115, rectRadius: 0.2, fill: { type: 'none' }, line: dash });
  s.addShape('rect', { x: 3.373, y: 3.149, w: 13.522, h: 0.79, fill: { color: WHITE }, line: { type: 'none' } });

  ROLE5.forEach((c, i) => {
    const x = 2.167 + i * 3.252;
    s.addShape('ellipse', { x: x + 0.482, y: 2.697, w: 1.694, h: 1.694, fill: { color: dk25(c) }, line: { type: 'none' } });
    s.addShape('ellipse', { x: x + 0.532, y: 2.747, w: 1.594, h: 1.594, fill: { color: c }, line: { type: 'none' } });
    s.addShape('ellipse', { x: x + 0.64, y: 2.83, w: 1.15, h: 1.15, fill: { color: t40(c), transparency: 45 }, line: { type: 'none' } });
    s.addText(RASCI[i], { x: x + 0.829, y: 2.929, w: 1.0, h: 1.111, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 60, bold: true, color: WHITE });
    if (i < 4) s.addShape('line', { x: x + 2.505, y: 3.544, w: 0.9, h: 0, line: dash });
    s.addText(ROLE_NAMES[i].toUpperCase(), { x, y: 4.734, w: 2.657, h: 0.64, align: 'center', fontFace: BEBAS, fontSize: 32, color: c });
    tb(s, ROLE_LEAD[i], { x, y: 5.626, w: 2.657, h: 0.95 });
    bullets(s, ROLE_BULLETS[i].slice(0, 2), { x, y: 6.805, w: 2.657, h: 1.8 });
  });
  s.addShape('roundRect', { x: 6.063, y: 9.149, w: 7.874, h: 0.984, rectRadius: 0.49, fill: { color: tint(WHITE, 95) }, line: { type: 'none' } });
  s.addText('This framework helps ensure clear communication, accountability, and effective collaboration within a team or organization.',
    { x: 6.063, y: 9.149, w: 7.874, h: 0.984, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 16, color: GRAY });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 14 — RASCI list on a teal wash
 * ------------------------------------------------------------------ */
function slide14(pptx) {
  const s = pptx.addSlide();
  gradientBands(s, BLUE, CYAN, -45);
  s.addText('RASCI', { x: 1.579, y: 1.439, w: 4.5, h: 1.346, fontFace: FONT, fontSize: 80, bold: true, color: WHITE, valign: 'top', margin: 0 });
  s.addText(FRAMEWORK_BLURB, { x: 1.579, y: 5.767, w: 4.5, h: 3.299, fontFace: FONT, fontSize: 28, bold: true, italic: true, color: WHITE, valign: 'top', margin: 0 });

  const cols = [RED, ORANGE, GREEN, CYAN, BLUE];
  cols.forEach((c, i) => {
    const y = 1.466 + i * 1.686;
    s.addShape('roundRect', { x: 6.938, y, w: 11.5, h: 1.575, rectRadius: 0.13, fill: { color: WHITE }, line: { type: 'none' } });
    letter3d(s, 7.237, y + 0.2, 1.378, 1.18, RASCI[i], c);
    s.addText(ROLE_NAMES[i], { x: 9.063, y: y + 0.534, w: 2.6, h: 0.505, fontFace: FONT, fontSize: 24, bold: true, color: c, valign: 'top' });
    s.addShape('line', { x: 11.529, y: y + 0.196, w: 0, h: 1.181, line: { color: tint(WHITE, 85), width: 2 } });
    tb(s, ROLE_LONG[i], { x: 11.908, y: y + 0.333, w: 6.179, h: 1.2 });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 15 — light cards with circular badges
 * ------------------------------------------------------------------ */
function slide15(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Model', SUB_TOOL);
  const cols = [RED, ORANGE, GREEN, TEAL, CYAN];
  const labels = ['responsible', 'accountable', 'Supportive', 'consulted', 'informed'];
  cols.forEach((c, i) => {
    s.addShape('roundRect', { x: COL_X[i], y: 3.436, w: COL_W, h: 6.299, rectRadius: 0.81, fill: { color: tint(WHITE, 95) }, line: { type: 'none' } });
    s.addShape('ellipse', { x: COL_X[i] + 0.675, y: 2.601, w: 1.8, h: 1.8, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
    s.addShape('ellipse', { x: COL_X[i] + 0.825, y: 2.752, w: 1.5, h: 1.5, fill: { color: c }, line: { type: 'none' } });
    s.addText(RASCI[i], { x: COL_X[i] + 1.075, y: 2.9, w: 1.0, h: 1.2, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 40, color: WHITE });
    s.addText(labels[i], { x: TXT_X[i], y: 4.733, w: TXT_W, h: 0.64, align: 'center', fontFace: BEBAS, fontSize: 32, color: c });
    tb(s, ROLE_LEAD[i], { x: TXT_X[i], y: 5.521, w: TXT_W, h: 0.95 });
    bullets(s, ROLE_BULLETS[i], { x: TXT_X[i], y: 6.645, w: TXT_W, h: 2.3 });
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 16 — six cards with coloured header bars
 * ------------------------------------------------------------------ */
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: 'F8F8F8' };
  s.addText('RASCI Model', { x: 1.579, y: 1.184, w: 4.929, h: 0.808, fontFace: FONT, fontSize: 48, bold: true, color: BLACK, valign: 'top', margin: 0 });
  s.addText(FRAMEWORK_BLURB, { x: 1.579, y: 2.635, w: 5.118, h: 2.02, fontFace: FONT, fontSize: 24, bold: true, italic: true, color: GRAY, valign: 'top', margin: 0 });

  const cards = [
    ['r', 'RESPONSIBLE', ORANGE, 7.288, 1.0, 0], ['a', 'ACCOUNTABLE', GREEN, 12.99, 1.0, 1],
    ['s', 'SUPPORTAIVE', TEAL, 1.586, 5.683, 2], ['c', 'CONSULTED', CYAN, 7.288, 5.683, 3],
    ['I', 'INFORMED', BLUE, 12.99, 5.683, 4],
  ];
  cards.forEach(([ch, label, c, x, y, i]) => {
    s.addShape('roundRect', { x, y, w: 5.5, h: 4.5, rectRadius: 0.31, fill: { color: WHITE }, line: { type: 'none' } });
    s.addShape('round2SameRect', { x, y, w: 5.5, h: 0.984, rectRadius: 0.31, fill: { color: c }, line: { type: 'none' } });
    s.addText(label, { x, y, w: 5.5, h: 0.984, align: 'center', valign: 'middle', fontFace: MONT, fontSize: 24, bold: true, color: WHITE });
    s.addText(ch, { x: x + 0.249, y: y + 1.25, w: 0.928, h: 1.616, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: 96, color: t80(c) });
    tb(s, ROLE_LEAD[i], { x: x + 1.601, y: y + 1.324, w: 3.7, h: 0.7 });
    bullets(s, ROLE_BULLETS[i], { x: x + 1.601, y: y + 2.095, w: 3.5, h: 2.0 });
    roleIcon(s, x + 0.16, y + 3.37, 0.98, c);
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 17 — leaf-shaped tinted cards
 * ------------------------------------------------------------------ */
function slide17(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Framework', SUB_TOOL);
  ROLE5.forEach((c, i) => {
    const x = 1.593 + i * 3.4163;
    s.addShape('round2DiagRect', { x, y: 2.957, w: COL_W, h: 7.0, rectRadius: 1.4, fill: { color: t80(c) }, line: { type: 'none' } });
    s.addShape('ellipse', { x, y: 2.957, w: 1.146, h: 1.146, fill: { color: c }, line: { type: 'none' } });
    s.addText(RASCI[i], { x, y: 3.03, w: 1.146, h: 1.146, align: 'center', valign: 'middle', fontFace: BEBAS, fontSize: 48, color: WHITE });
    s.addText(ROLE_NAMES[i], { x: x + 0.197, y: 4.349, w: 2.3, h: 0.438, fontFace: FONT, fontSize: 20, bold: true, color: c, valign: 'top' });
    tb(s, ROLE_LEAD[i], { x: x + 0.18, y: 5.032, w: TXT_W, h: 0.95, color: BLACK });
    bullets(s, ROLE_BULLETS[i], { x: x + 0.18, y: 6.244, w: TXT_W, h: 2.3, color: BLACK });
    roleIcon(s, x + 0.185, 8.955, 0.79, tint(c, 40, 60));
  });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slides 18 & 19 — website-launch worked example (two spreads)
 * ------------------------------------------------------------------ */
function roleCard(s, x, label, color, rows, bulletMode) {
  s.addShape('roundRect', { x, y: 1.234, w: 4.134, h: 8.708, rectRadius: 0.27, fill: { color: WHITE }, line: { type: 'none' } });
  s.addShape('ellipse', { x: x + 0.716, y: 1.834, w: 2.0, h: 2.0, fill: { color: t80(color), transparency: 50 }, line: { type: 'none' } });
  roleIcon(s, x + 0.4, 1.882, 1.97, color);
  s.addText(label, { x: x + 0.196, y: 4.277, w: 3.2, h: 0.505, fontFace: FONT, fontSize: 24, bold: true, color, valign: 'top' });
  const runs = [];
  rows.forEach(([head, body], i) => {
    if (i) runs.push({ text: '', options: { breakLine: true } });
    if (bulletMode && i) {
      runs.push({ text: body, options: { bullet: { indent: 20 }, breakLine: true } });
    } else {
      if (head) runs.push({ text: head, options: { bold: true, underline: true } });
      runs.push({ text: body, options: { breakLine: true } });
    }
  });
  tb(s, runs, { x: x + 0.196, y: 5.03, w: 3.74, h: 4.7 });
}

function slide18(pptx) {
  const s = pptx.addSlide();
  s.background = { color: 'FAFAFA' };
  gradientPill(s, 1.575, 1.234, 2.953, 0.611, RED, BLUE, 'RASCI MODEL', 20);
  s.addText([
    { text: 'RASCI Model for a', options: { color: BLACK, breakLine: true } },
    { text: 'New Website Launching', options: { color: RED } },
  ], { x: 1.575, y: 2.786, w: 7.5, h: 1.616, fontFace: FONT, fontSize: 48, bold: true, valign: 'top', margin: 0 });
  tb(s, 'This detailed example demonstrates how the RASCI framework can be applied to a website development project, clarifying roles and responsibilities to ensure smooth collaboration and project success.',
    { x: 1.575, y: 4.845, w: 7.5, h: 1.01, fontSize: 18, margin: [3.6, 0, 3.6, 0] });
  s.addText([
    { text: 'In this scenario:', options: { breakLine: true } }, { text: '', options: { breakLine: true } },
    ...[['R (Responsible):', ' Marketing Team, UX/UI Designer, Web Developer, Content Creator, Quality Assurance Team.'],
    ['A (Accountable): ', 'Project Manager'],
    ['S (Supportive):', ' UX/UI Designer, Web Developer, Content Creator, Quality Assurance Team.'],
    ['C (Consulted):', ' Marketing Team, Graphic Designer, IT Department, Sales Team.'],
    ['I (Informed):', ' Stockholders']].flatMap(([b, t]) => ([
      { text: b, options: { bold: true } }, { text: t, options: { breakLine: true } }])),
  ], { x: 1.575, y: 6.736, w: 7.478, h: 3.029, fontFace: FONT, fontSize: 18, color: BLACK, valign: 'top', margin: 0 });

  roleCard(s, 9.789, 'RESPONSIBLE', RED, [
    ['Marketing Team :', ' Website objectives and target audience'],
    ['UX/UI Designer :', ' Develop website wireframes and design'],
    ['Web Developer :', ' Write code for website functionality'],
    ['Content Creator :', ' Create and optimize website content'],
    ['Quality Assurance Team :', ' Conduct website testing and quality assurance'],
  ]);
  roleCard(s, 14.268, 'ACCOUNTATBLE', ORANGE, [
    ['Project Manager : ', ''],
    ['', 'Oversee the entire website development process.'],
    ['', 'Ensure alignment with project timelines and objectives.'],
    ['', 'Coordinate communication between different teams and stakeholders.'],
    ['', 'Ultimately responsible for the successful delivery of the website project.'],
  ], true);
  s.addShape('homePlate', { x: 18.75, y: 5.36, w: 1.07, h: 0.53, fill: { color: tint(WHITE, 85) }, line: { type: 'none' } });
  footer(s);
}

function slide19(pptx) {
  const s = pptx.addSlide();
  s.background = { color: 'FAFAFA' };
  roleCard(s, 1.611, 'SUPPORTIVE', GREEN, [
    ['UX/UI Designer :', ' Support the design process by providing feedback and guidance.'],
    ['Web Developer :', ' Receive support from the UX/UI designer and graphic designer during implementation.'],
    ['Content Creator :', ' Receive support from the marketing team for content strategy and messaging.'],
    ['Quality Assurance Team :', ' Receive support from the development team for bug fixes and code adjustments.'],
  ]);
  roleCard(s, 6.092, 'CONSULTED', TEAL, [
    ['Marketing Team :', ' Provide input on branding, messaging, and content strategy.'],
    ['Graphic Designer :', ' Offer design expertise and creative input.'],
    ['IT Department :', ' Provide technical guidance and support.'],
    ['Sales Team :', ' Provide insights into customer needs and preferences.'],
  ]);
  roleCard(s, 10.574, 'INFORMED', CYAN, [
    ['Stakeholders : ', ''],
    ['', 'Receive regular updates on project progress and milestones.'],
    ['', 'Provide feedback and approval at key stages of website development.'],
    ['', 'Informed about any changes or decisions affecting the project timeline or scope.'],
  ], true);
  s.addText('By applying the RASCI framework in this example, Companies can ensure clarity of roles and responsibilities, effective collaboration between teams, and accountability for the successful development and launch of their new software product.',
    { x: 15.056, y: 5.029, w: 3.54, h: 3.13, fontFace: FONT, fontSize: 18, bold: true, color: GRAY, valign: 'top' });
  s.addShape('homePlate', { x: 0, y: 5.36, w: 1.07, h: 0.53, fill: { color: tint(WHITE, 85) }, line: { type: 'none' } });
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slides 20 & 21 — Roles / Responsibilities / Tasks board
 * ------------------------------------------------------------------ */
const BOARD_HEADS = [['Roles', 3.463, 4.134], ['Responsibilities', 7.691, 4.134], ['Tasks or Activity', 11.903, 6.693]];

function boardSlide(pptx, opts) {
  const s = pptx.addSlide();
  gradientBands(s, RED, BLUE, -30);
  if (opts.title) titleBlock(s, opts.title, opts.subtitle, WHITE);
  BOARD_HEADS.forEach(([label, x, w]) => {
    s.addShape('roundRect', { x, y: 2.201, w, h: 0.8, rectRadius: 0.18, fill: { color: WHITE }, line: { type: 'none' } });
    s.addText(label, { x, y: 2.201, w, h: 0.8, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 18, bold: true, color: BLACK });
  });
  opts.rows.forEach((row, i) => {
    const y = 3.101 + i * 2.389;
    const [name, color, roleTitle, roleBody, resp, tasks] = row;
    s.addShape('roundRect', { x: 1.579, y, w: 1.791, h: 2.3, rectRadius: 0.19, fill: { color }, line: { type: 'none' } });
    letter3d(s, 1.845, y + 0.13, 1.26, 1.05, RASCI[opts.startIndex + i], color);
    s.addText(name, { x: 1.642, y: y + 1.352, w: 1.666, h: 0.404, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: WHITE });

    s.addShape('roundRect', { x: 3.463, y, w: 4.134, h: 2.3, rectRadius: 0.19, fill: { color: WHITE }, line: { type: 'none' } });
    s.addText([
      { text: roleTitle, options: { bold: true, underline: true, color, breakLine: true } },
      { text: roleBody, options: { color: GRAY } },
    ], { x: 3.463, y: y + 0.12, w: 4.134, h: 2.1, fontFace: FONT, fontSize: 16, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });

    s.addShape('roundRect', { x: 7.691, y, w: 4.134, h: 2.3, rectRadius: 0.19, fill: { color: WHITE }, line: { type: 'none' } });
    tb(s, resp.join('\n'), { x: 7.691, y: y + 0.12, w: 4.134, h: 2.1 });

    s.addShape('roundRect', { x: 11.903, y, w: 6.693, h: 2.3, rectRadius: 0.19, fill: { color: WHITE }, line: { type: 'none' } });
    tasks.forEach((t, k) => {
      tb(s, t, { x: 12.1, y: y + 0.13 + k * 0.633, w: 6.299, h: 0.6 });
      if (k) s.addShape('line', { x: 12.1, y: y + 0.13 + k * 0.633, w: 6.299, h: 0, line: { color: 'D9D9D9', width: 0.75 } });
    });
  });
  s.addShape('homePlate', { x: 9.679, y: opts.arrowY, w: 0.642, h: 1.414, rotate: 90, fill: { color: WHITE, transparency: 80 }, line: { type: 'none' } });
  if (opts.footnote) {
    s.addText(opts.footnote, { x: 1.642, y: 8.696, w: 16.955, h: 0.9, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: WHITE, valign: 'top' });
  }
  footer(s);
  return s;
}

function slide20(pptx) {
  boardSlide(pptx, {
    title: 'RASCI Framework', startIndex: 0, arrowY: 10.096,
    subtitle: [{ text: 'The RASCI framework to the development and launch process of ' },
      { text: 'software product', options: { underline: true } }, { text: '.' }],
    rows: [
      ['Responsible', RED, 'The Development Team',
        'This team comprises software engineers, designers, and testers responsible for building the software product according to the specifications provided by the product management.',
        ['Responsibilities include coding, designing user interfaces, writing documentation, testing the software for bugs and errors, and ensuring the product meets quality standards.'],
        ['Software engineers write code based on the product requirements.', 'Designers create user interfaces and graphic elements for the software.', 'Testers perform various tests, including functional testing, performance testing, and user acceptance testing.']],
      ['Accountable', ORANGE, 'Product Manager',
        'The product manager is ultimately accountable for the success of the software product. They oversee the entire product development and launch process.',
        ['Responsibilities include defining product requirements, prioritizing features, setting deadlines, and ensuring alignment with company goals.'],
        ['Defining the product features and specifications based on market research and customer feedback.', 'Prioritizing features and creating a roadmap for the product development cycle.', 'Monitoring progress, resolving conflicts, and making strategic decisions to keep the project on track.']],
      ['Supportive', GREEN, 'IT and Operations Teams',
        'The IT and operations teams provide support services necessary for the development, deployment, and maintenance of the software product.',
        ['IT support ensures the availability of development tools, infrastructure, and security measures.', 'Operations support facilitates the deployment of the software to production environments and manages ongoing maintenance.'],
        ['IT team: Setting up development environments, managing version control systems, and ensuring data security.', 'Operations team: Deploying software updates, monitoring system performance, and providing technical support to users.']],
    ],
  });
}

function slide21(pptx) {
  boardSlide(pptx, {
    startIndex: 3, arrowY: -0.386,
    footnote: 'By applying the RASCI framework in this example, Companies can ensure clarity of roles and responsibilities, effective collaboration between teams, and accountability for the successful development and launch of their new software product.',
    rows: [
      ['Consulted', CYAN, 'Marketing and Customer Support Teams',
        'The marketing team and customer support team provide valuable input throughout the product development and launch process.',
        ['The marketing team advises on market trends, target audience preferences, and promotional strategies.', 'The customer support team offers insights into common user issues, feature requests, and user experience improvements.'],
        ['Marketing team: Providing market research data, advising on product positioning, and collaborating on marketing campaigns.', 'Customer support team: Sharing customer feedback, participating in usability testing, and suggesting product enhancements.']],
      ['Informed', BLUE, 'Executive Leadership Team',
        'The executive leadership team is kept informed about the progress, challenges, and achievements related to the development and launch of the software product.',
        ['They may not be directly involved in day-to-day activities but provide strategic guidance and resources as needed.'],
        ['Providing regular progress updates to the executive leadership team through status reports, presentations, and meetings.', 'Seeking approval for major decisions, budget allocations, and resource allocation requests.']],
    ],
  });
}

/* ------------------------------------------------------------------ *
 * Slides 22-26 — tabbed detail pages, one per RASCI letter
 * ------------------------------------------------------------------ */
const TAB_X = [1.575, 4.971, 8.366, 11.762, 15.157];

function tabbedSlide(pptx, i, color, intro, whoLine, who, bulletsList) {
  const s = pptx.addSlide();
  s.addText([
    { text: 'RASCI ', options: { bold: true } }, { text: 'Framework', options: { bold: true, italic: true } },
  ], { x: 1.575, y: 0.821, w: 5.859, h: 0.909, fontFace: FONT, fontSize: 48, color: BLACK, valign: 'top' });
  s.addText('RASCI Framework for defining and clarifying roles in project management and organizational development. (With Example)',
    { x: 8.366, y: 0.821, w: 10.224, h: 0.909, fontFace: FONT, fontSize: 24, bold: true, italic: true, color: BLACK, valign: 'top' });

  TAB_X.forEach((x, k) => {
    const on = k === i;
    s.addShape('round2SameRect', { x, y: 2.632, w: 3.268, h: 2.0, rectRadius: 0.2, fill: { color: on ? color : tint(WHITE, 95) }, line: { type: 'none' } });
    s.addShape('ellipse', { x: x + 0.228, y: 2.843, w: 0.5, h: 0.5, fill: { color: on ? WHITE : tint(WHITE, 95) }, line: { color: on ? color : GRAY50, width: 1 } });
    s.addText('+', { x: x + 0.228, y: 2.843, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 16, color: on ? color : GRAY50 });
    letter3d(s, x + 1.004, 2.795, 1.26, 1.0, RASCI[k], on ? color : 'BFBFBF');
    s.addText(ROLE_NAMES[k], { x: x + 0.534, y: 3.979, w: 2.2, h: 0.438, align: 'center', fontFace: FONT, fontSize: 20, bold: true, color: on ? WHITE : GRAY50 });
  });

  s.addShape('round2SameRect', { x: 1.575, y: 4.577, w: 16.85, h: 5.232, rectRadius: 0.3, fill: { color }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 2.203, y: 5.253, w: 2.479, h: 2.479, fill: { color: t80(ORANGE), transparency: 50 }, line: { type: 'none' } });
  roleIcon(s, 3.187, 5.508, 1.97, WHITE);
  s.addText(`${ROLE_NAMES[i]} (${RASCI[i]})`, { x: 2.077, y: 8.227, w: 4.199, h: 0.539, align: 'center', fontFace: FONT, fontSize: 32, bold: true, color: WHITE });
  s.addText(who, { x: 2.077, y: 8.89, w: 4.199, h: 0.4, align: 'center', fontFace: FONT, fontSize: 16, bold: true, underline: true, color: WHITE });

  s.addText(intro, { x: 7.234, y: 5.112, w: 10.5, h: 0.7, fontFace: FONT, fontSize: 18, bold: true, italic: true, color: WHITE, valign: 'top', margin: 0 });
  s.addShape('roundRect', { x: 7.234, y: 6.085, w: 6.495, h: 0.591, rectRadius: 0.29, fill: { color: t80(color) }, line: { type: 'none' } });
  s.addText([
    { text: 'Example Task : ', options: { bold: true } },
    { text: 'Develop and launch a new software product. ', options: { bold: true, underline: true } },
  ], { x: 7.234, y: 6.085, w: 6.495, h: 0.591, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 16, color: BLACK });
  s.addText(whoLine, { x: 7.234, y: 6.863, w: 10.329, h: 0.37, fontFace: FONT, fontSize: 16, bold: true, color: WHITE, valign: 'top', margin: [0, 0, 0, 0] });
  tb(s, bulletsList.flatMap((t, k) => ([
    ...(k ? [{ text: '', options: { breakLine: true } }] : []),
    { text: t, options: { bullet: { indent: 20 }, breakLine: true } }])),
    { x: 7.234, y: 7.453, w: 10.329, h: 2.1, color: WHITE, margin: [0, 0, 0, 0] });
  footer(s);
}

function slide22(pptx) {
  tabbedSlide(pptx, 0, RED,
    'This person or group is responsible for completing the task or activity. They are the ones who actually perform the work.',
    'Responsible (R) : The Development Team', 'The Development Team',
    ['The development team comprises software engineers, designers, and testers responsible for building the software product according to the specifications provided by the product management team.',
      'Responsibilities include coding, designing user interfaces, writing documentation, testing the software for bugs and errors, and ensuring the product meets quality standards.']);
}
function slide23(pptx) {
  tabbedSlide(pptx, 1, ORANGE,
    'This person ultimately owns the task or activity. They are answerable for the completion and success of the task. There should be only one person accountable for each task.',
    'Accountable (A) : The Product Manager', 'The Product Manager',
    ['The product manager is ultimately accountable for the success of the software product. They define the product vision, strategy, and roadmap, and they oversee the entire product development and launch process.',
      'Responsibilities include defining product requirements, prioritizing features, setting deadlines, and ensuring alignment with company goals.']);
}
function slide24(pptx) {
  tabbedSlide(pptx, 2, GREEN,
    'These are the individuals or groups who assist the responsible person or group in completing the task. They may provide resources, guidance, or assistance as needed.',
    'Supportive (S) : IT and Operation Teams', 'IT and Operation Teams',
    ['The IT and operations teams provide support services necessary for the development, deployment, and maintenance of the software product. ',
      'IT support ensures the availability of development tools, infrastructure, and security measures.',
      'Operations support facilitates the deployment of the software to production environments and manages ongoing maintenance.']);
}
function slide25(pptx) {
  tabbedSlide(pptx, 3, CYAN,
    'These are the individuals or groups whose input or expertise is sought before a decision or action is taken. They provide valuable insights and feedback but do not necessarily perform the task.',
    'Consulted (C) : Marketing and Customer Support Teams', 'Marketing and Customer Support Teams',
    ['The marketing team and customer support team provide valuable input throughout the product development and launch process.',
      'The marketing team advises on market trends, target audience preferences, and promotional strategies.',
      'The customer support team offers insights into common user issues, feature requests, and user experience improvements.']);
}
function slide26(pptx) {
  tabbedSlide(pptx, 4, BLUE,
    'These are the individuals or groups who are kept informed about the progress, decisions, or outcomes of the task or activity. They may not be directly involved in the task but need to be aware of its status.',
    'Informed (I) : Executive Leadership Team', 'Executive Leadership Team',
    ['The executive leadership team is kept informed about the progress, challenges, and achievements related to the development and launch of the software product.',
      'They may not be directly involved in day-to-day activities but provide strategic guidance and resources as needed.']);
}

/* ------------------------------------------------------------------ *
 * Slides 27 & 28 — task matrix with a colour-coded header row
 * ------------------------------------------------------------------ */
const SCENARIO_LEGEND = [
  ['Responsible', RED, 'The individuals or teams responsible for completing each task.'],
  ['Accountable', ORANGE, "The person who is ultimately accountable for the task's success."],
  ['Consulted', GREEN, 'People who need to provide input or expertise during the task.'],
  ['Supported', CYAN, 'Teams or individuals providing support to the responsible party.'],
  ['Informed', BLUE, 'Stakeholders who need to be kept informed about the progress or outcome of the task.'],
];

/** Left-hand "In this scenario:" column shared by slides 27 and 28. */
function scenarioSidebar(s, y0) {
  s.addText('RASCI matrix of a project to redesign a company website',
    { x: 1.6, y: y0, w: 4.134, h: 0.673, fontFace: FONT, fontSize: 20, bold: true, color: BLACK, valign: 'top', margin: 0 });
  s.addShape('line', { x: 1.6, y: y0 + 1.182, w: 4.134, h: 0, line: { color: RED, width: 3 } });
  s.addText([
    { text: 'In this scenario:', options: { bold: true, italic: true, underline: true, breakLine: true } },
    { text: '', options: { breakLine: true } },
    ...SCENARIO_LEGEND.flatMap(([name, c, body]) => ([
      { text: name + ': ', options: { bold: true, color: c } },
      { text: body, options: { color: BLACK, breakLine: true } },
      { text: '', options: { breakLine: true } },
    ])),
  ], { x: 1.579, y: y0 + 1.779, w: 4.134, h: 4.847, fontFace: FONT, fontSize: 16, valign: 'top', margin: 0 });
}

const TASK_HEAD = ['Tasks', 'Responsible\n(R)', 'Accountable\n(A)', 'Supportive\n(S)', 'Consulted\n(C)', 'Informed\n(I)'];

function taskTable(s, rows, y, rowH) {
  const head = TASK_HEAD.map((t, i) => ({
    text: t,
    options: { fill: { color: i ? tint(ROLE5[i - 1], 30, 70) : tint(WHITE, 92) }, bold: true, fontSize: 18, align: 'center', color: BLACK },
  }));
  const body = rows.map((r, ri) => r.map((t, ci) => ({
    text: t,
    options: {
      fill: { color: ri % 2 ? WHITE : tint(WHITE, 98) },
      fontSize: 16, italic: ci > 0, align: ci ? 'center' : 'left', color: BLACK,
    },
  })));
  s.addTable([head].concat(body), {
    x: 6.312, y, w: 12.093, colW: [3.304, 1.758, 1.758, 1.758, 1.758, 1.758],
    rowH, valign: 'middle', fontFace: FONT, border: { type: 'solid', color: 'FFFFFF', pt: 1 },
    margin: [2, 6, 2, 6],
  });
}

function slide27(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Matrix', 'A tool used in project management to define and clarify the roles and responsibilities of individuals or groups involved in a project or process. ');
  scenarioSidebar(s, 2.932);
  taskTable(s, [
    ['Define Project Requirement', 'Marketing', 'Project Manager', 'IT', 'Design Team', 'Executive Team'],
    ['Create Wireframe', 'Design Team', 'Project Manager', 'IT, Marketing', '', 'Executive Team'],
    ['Develop Website', 'Development', 'Project\nManager', 'IT, Design', '', 'Executive Team'],
    ['Test Website', 'Quality Assurance', 'Project\nManager', 'Development', 'Design Team', 'Executive Team'],
    ['Deploy Website', 'IT', 'Project\nManager', 'Development', '', 'Executive Team'],
    ['Monitor Website Performance', 'IT', 'Project\nManager', 'Development', 'Marketing', 'Executive Team'],
  ], 2.932, 0.946);
  footer(s);
}

function slide28(pptx) {
  const s = pptx.addSlide();
  s.addText('RASCI Matrix', { x: 1.579, y: 1.184, w: 4.733, h: 0.808, fontFace: FONT, fontSize: 48, bold: true, color: BLACK, valign: 'top', margin: 0 });
  scenarioSidebar(s, 2.932);
  const PM = 'Project Manager', ST = 'Stakeholders, Project Manager';
  taskTable(s, [
    ['Define website objectives and target audience', 'Marketing', PM, 'Sales Team, UX/UI Designer', 'IT Department', ST],
    ['Develop website wireframes and design', 'UX/UI Designer', PM, 'IT Department', 'Marketing, Graphic Designer', ST],
    ['Code website functionality and backend', 'Web Developer', 'Project\nManager', 'Quality Assurance Team', 'UX/UI Designer, IT Department', ST],
    ['Create and optimize website content', 'Content Creator', 'Project\nManager', 'Graphic Designer, UX/UI Designer', 'Marketing, SEO Specialist', ST],
    ['Implement SEO strategies and keywords', 'SEO Specialist', 'Project\nManager', 'Web Developer, IT Department', 'Marketing, Content Creator', ST],
    ['Set up website hosting and domain', 'IT Administrator', 'Project\nManager', '-', 'IT Department', ST],
    ['Perform website testing and quality assurance', 'Quality Assurance Team', PM, 'IT Department', 'Web Developer, UX/UI Designer', ST],
    ['Launch website and monitor performance', 'Web Developer, IT Administrator', PM, '-', 'Marketing, Graphic Designer', ST],
  ], 1.247, 0.973);
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slides 29-32 — RASCI grids (activities x teams)
 * ------------------------------------------------------------------ */
const GRID_FILL = { R: t40(RED), A: t40(ORANGE), S: t40(GREEN), C: t40(CYAN), I: t40(BLUE), '': tint(WHITE, 95) };
const LEGEND = [['R = Responsible', RED, 'Specifies who is responsible for executing each activity.'],
['A = Accountable', ORANGE, 'Indicates who is ultimately answerable for the success of the activity.'],
['S = Supported', GREEN, 'Identifies those who provide assistance, resources, or guidance to the responsible person or team.'],
['C = Consulted', CYAN, 'Lists individuals or teams whose input or expertise is sought before decisions are made.'],
['I = Informed', BLUE, 'Includes stakeholders or individuals who need to be kept informed about the progress, decisions, or outcomes of the activity.']];

/** Legend strip: five tinted chips with an explanation beneath each. */
function legendStrip(s, y, gap, fontSize, withText) {
  LEGEND.forEach(([label, c, body], i) => {
    const x = 1.576 + i * gap;
    s.addShape('round2DiagRect', { x, y, w: 3.15, h: 0.394, rectRadius: 0, fill: { color: t40(c) }, line: { type: 'none' } });
    s.addText(label, { x, y, w: 3.15, h: 0.394, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 16, bold: true, color: BLACK });
    if (withText) s.addText(body, { x, y: y + 0.584, w: 3.15, h: 1.4, align: 'center', fontFace: FONT, fontSize, color: GRAY, valign: 'top', margin: 0 });
  });
}

/**
 * Grid renderer. `rows[0]` = team names, remaining rows are
 * [activity, ...cells] where each cell is "text|KEY" (KEY picks the fill).
 */
function rasciGrid(s, opts) {
  const heads = opts.heads, rows = opts.rows;
  const nCols = heads.length + 1;
  const headRow = [{ text: '', options: { fill: { color: WHITE } } }].concat(
    heads.map(t => ({ text: t, options: { fill: { color: tint(NAVY, 128) }, color: WHITE, bold: true, align: 'center' } })));
  const noteRow = [
    { text: 'Activities', options: { fill: { color: tint(NAVY, 128) }, color: WHITE, bold: true, align: 'left' } },
    { text: 'Fill for each activity : Responsibility (R), Accountable (A), Supportive (S), ' + (opts.confirmed ? 'Confirmed' : 'Consulted') + ' (C), Informed (I)', options: { fill: { color: WHITE }, bold: true, align: 'center', colspan: heads.length } },
  ];
  const body = rows.map(r => [{ text: r[0], options: { fill: { color: tint(WHITE, 97) }, bold: true, align: 'left' } }].concat(
    r.slice(1).map(cell => {
      const [txt, key] = String(cell).split('|');
      return { text: txt, options: { fill: { color: GRID_FILL[(key || '').charAt(0)] || tint(WHITE, 95) }, align: 'center' } };
    })));
  const colW = [opts.firstW].concat(new Array(heads.length).fill(opts.cellW));
  s.addTable([headRow, noteRow].concat(body), {
    x: 1.576, y: opts.y, w: opts.firstW + heads.length * opts.cellW, colW,
    rowH: [0.787, 0.591].concat(rows.map(() => opts.rowH)),
    valign: 'middle', fontFace: FONT, fontSize: opts.fontSize || 16,
    color: BLACK, border: { type: 'solid', color: 'FFFFFF', pt: 1 }, margin: [1, 4, 1, 4],
  });
}

function slide29(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Matrix', [
    { text: 'RASCI matrix specifically for ' }, { text: 'Graphis Design Business', options: { underline: true } },
  ]);
  rasciGrid(s, {
    y: 2.729, firstW: 2.995, cellW: 1.539, rowH: 0.787, confirmed: true,
    heads: ['Project Manager', 'Marketing Team', 'Sales Team', 'Production Team', 'IT Team', 'Graphics Designer', 'Customer Support', 'Senior Management', 'Stakeholders'],
    rows: [
      ['Client consultation and requirements gathering', 'A|A', 'S|S', 'C|C', '', '', 'R|R', 'S|S', 'I|I', 'I|I'],
      ['Design conceptualization and creation', 'A|A', 'C|C', 'C|C', '', 'S|S', 'R|R', 'S|S', 'I|I', 'I|I'],
      ['Presentation and review with client', 'A|A', 'C|C', 'C|C', '', '', 'R|R', 'S|S', 'I|I', 'I|I'],
      ['Revisions and finalization', 'A|A', 'C|C', 'C|C', '', '', 'R|R', 'S|S', 'I|I', 'I|I'],
      ['File preparation and delivery', 'A|A', '', 'S|S', 'C|C', 'C|C', 'R|R', '', 'I|I', 'I|I'],
    ],
  });
  legendStrip(s, 8.235, 3.425, 16, true);
  footer(s);
}

function slide30(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Matrix', [
    { text: 'RASCI matrix for assigning Responsibility of a ' },
    { text: 'Project to Design a Company Website', options: { underline: true } },
  ]);
  const R = 'Responsible (R)|R', A = 'Accountable (A)|A', I = 'Informed (I)|I', CI = 'Consulted (C) Informed (I)|C';
  rasciGrid(s, {
    y: 2.798, firstW: 3.742, cellW: 1.638, rowH: 0.787,
    heads: ['Analyst', 'UI/UX', 'Project Manager', 'IT Expert', 'Software Developer', 'Tester', 'System Administrator', 'User'],
    rows: [
      ['Planning and Analysis', R, '', A, 'Consulted (C)|C', I, '', '', CI],
      ['Designing and Prototyping', 'Supportive (S)|S', R, A, '', I, '', '', CI],
      ['Coding', '', '', A, '', R, 'Supportive (S) Informed (I)|S', I, I],
      ['Review and Testing', '', '', A, '', '', R, I, CI],
      ['Development', '', '', A, '', 'Supportive (S)|S', 'Supportive (S)|S', R, I],
      ['Maintenance', '', 'Supportive (S)|S', A, '', 'Supportive (S)|S', 'Supportive (S)|S', R, CI],
    ],
  });
  legendStrip(s, 9.094, 3.425, 16, false);
  footer(s);
}

function slide31(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Matrix', [
    { text: 'RASCI matrix specifically for ' }, { text: 'Project Financial Management', options: { underline: true } },
  ]);
  rasciGrid(s, {
    y: 2.973, firstW: 2.995, cellW: 1.539, rowH: 0.787, confirmed: true,
    heads: ['Finance Team', 'Project Manager', 'Project Team', 'Business Owner', 'Department Heads', 'Accounting Team', 'Data Analysis', 'Senior Management', 'Stakeholders'],
    rows: [
      ['Budget Creation', 'R|R', 'AC|A', 'I|I', '', 'C|C', '', '', 'S|S', 'S|S'],
      ['Expenses Tracking and Reporting', 'A|A', 'RC|R', '', '', 'C|C', 'S|S', '', 'I|I', 'I|I'],
      ['Cost Analysis and Variance Analysis', 'R|R', 'AR|A', 'S|S', 'A|A', 'C|C', 'C|C', 'S|S', 'I|I', 'I|I'],
      ['Forecasting and Financial Projections', 'R|R', 'C|C', 'I|I', 'A|A', '', '', 'C|C', 'S|S', 'I|I'],
    ],
  });
  legendStrip(s, 7.916, 3.425, 16, true);
  footer(s);
}

function slide32(pptx) {
  const s = pptx.addSlide();
  titleBlock(s, 'RASCI Matrix', [
    { text: 'RASCI matrix specifically for ' }, { text: 'Drop Shipping Business', options: { underline: true } },
  ]);
  rasciGrid(s, {
    y: 2.432, firstW: 2.069, cellW: 1.356, rowH: 0.591, confirmed: true, fontSize: 14,
    heads: ['Procurement Team', 'Project Manager', 'Sales Team', 'Operation Team', 'Finance Team', 'Marketing Team', 'IT Team', 'Graphic Designer', 'Customer Support Team', 'Senior Management', 'Stakeholders'],
    rows: [
      ['Source and select suppliers', 'R|R', 'A|A', 'C|C', 'S|S', 'S|S', '', '', '', '', 'I|I', 'I|I'],
      ['Set up online store', '', 'A|A', 'C|C', 'S|S', '', 'C|C', 'R|R', 'S|S', '', 'I|I', 'I|I'],
      ['Integrate payment gateway', '', 'A|A', 'S|S', 'S|S', 'C|C', '', 'R|R', '', '', 'I|I', 'I|I'],
      ['Develop product listings', '', 'A|A', 'S|S', '', '', 'R|R', 'C|C', 'S|S', '', 'I|I', 'I|I'],
      ['Handling customers inquiries', '', 'A|A', 'C|C', 'S|S', '', 'C|C', 'S|S', '', 'R|R', 'I|I', 'I|I'],
      ['Process Customer Orders', 'S|S', 'A|A', 'C|C', 'R', '', '', 'C|C', '', 'S|S', 'I|I', 'I|I'],
      ['Manage product returns and refunds', '', 'A|A', 'C|C', 'C|C', 'S|S', '', 'S|S', '', 'R|R', 'I|I', 'I|I'],
      ['Monitor supplier performance', 'R|R', 'A|A', 'S|S', 'C|C', 'C|C', '', 'S|S', '', '', 'I|I', 'I|I'],
    ],
  });
  legendStrip(s, 8.725, 3.459, 12, true);
  footer(s);
}

/* ------------------------------------------------------------------ *
 * Slide 33 — closing
 * ------------------------------------------------------------------ */
function slide33(pptx) {
  const s = pptx.addSlide();
  coverBackdrop(s);
  ribbon(s, 7.27, 2.71, 5.45, 0.9, 'Free updates available', 36);
  s.addText('THANK', { x: 4.868, y: 3.852, w: 10.301, h: 3.13, align: 'center', valign: 'middle', fontFace: MONT_BLACK, bold: true, fontSize: 180, color: WHITE });
  s.addText('YOU', { x: 8.59, y: 6.863, w: 2.819, h: 1.313, align: 'center', valign: 'middle', fontFace: MONT_BLACK, bold: true, fontSize: 72, color: WHITE });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'RASCI', width: W, height: H });
  pptx.layout = 'RASCI';
  pptx.author = 'Infograficso';
  pptx.title = 'RASCI Framework and Matrix';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
    slide31, slide32, slide33].forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0d797f4e-bac1-4ee8-8e3f-825db18e92e1_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
