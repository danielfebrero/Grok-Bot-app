/**
 * WITOKU — Film Production Presentation Template (42 slides, 13.333in x 7.5in)
 *
 * Standalone pptxgenjs recreation of the reference deck.
 * Raster photographs in the original are replaced by flat colour placeholders
 * (see `photo`), which mirror the original crop boxes and dominant tone.
 *
 *   node <thisfile>.js   ->  writes <thisfile>.pptx next to the script
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  orange: 'CD6E35',   // accent1 — warm slide backgrounds
  green:  '0A4D42',   // dk1     — deep teal slide backgrounds
  pink:   'E3BEBB',   // accent2 — blush slide backgrounds
  cream:  'F8F8F8',   // lt2     — text on dark backgrounds
  ink:    '262626',   // tx2 85% — text on light backgrounds
  black:  '000000',
  grey:   'D9D9D9',
  white:  'FFFFFF',
};

const D = 'Bebas Neue';     // display face (all headlines / labels)
const BODY = 'Gudea';       // body face

// Flat stand-ins for the two photo treatments used throughout the template
// (the originals are blue and red "replace this image" placeholder plates).
const BLUE = '0726B8';
const RED  = 'FD2A07';

const W = 13.3333333, H = 7.5;   // 12192000 x 6858000 EMU (16:9)

/* ---------------------------------------------------------------- helpers */

// Full-bleed background colour.
function bg(s, color) {
  s.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color } });
}

// Plain rectangle.
function box(s, x, y, w, h, o = {}) {
  const opts = { x, y, w, h, line: { type: 'none' } };
  if (o.fill) opts.fill = { color: o.fill };
  if (o.line) opts.line = { color: o.line[0], width: o.line[1] };
  s.addShape('rect', opts);
}

// Straight connector (vertical when w === 0, horizontal when h === 0).
function rule(s, x, y, w, h, color, width) {
  s.addShape('line', { x, y, w, h, line: { color, width } });
}

/**
 * Text block. `body` is either a string (single style, "\n" splits paragraphs)
 * or an array of [text, runOptions] pairs for mixed runs.
 */
function t(s, body, x, y, w, h, o = {}) {
  const opts = {
    x, y, w, h,
    fontFace: o.face || BODY,
    fontSize: o.size === undefined ? 9 : o.size,
    color: o.color || C.ink,
    bold: !!o.bold,
    italic: !!o.italic,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: 0,
    wrap: true,
  };
  if (o.ls) opts.lineSpacingMultiple = o.ls;
  if (o.rot) opts.rotate = o.rot;
  if (typeof body === 'string') {
    s.addText(body, opts);
  } else {
    s.addText(body.map(([text, ro = {}]) => ({
      text,
      options: {
        fontFace: ro.face || opts.fontFace,
        fontSize: ro.size === undefined ? opts.fontSize : ro.size,
        color: ro.color || opts.color,
        bold: ro.bold === undefined ? opts.bold : ro.bold,
        breakLine: !!ro.brk,
      },
    })), opts);
  }
}

// Placeholder standing in for a photograph in the reference deck.
function photo(s, x, y, w, h, tone) {
  s.addShape('rect', { x, y, w, h, fill: { color: tone }, line: { type: 'none' } });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: Math.max(8, Math.min(14, h * 3)), color: 'FFFFFF',
    margin: 0,
  });
}

// The rotated identity rail down the left edge, present on slides 1-30.
function rail(s, color, page, lineColor) {
  t(s, 'WITOKU', -0.017, 6.203, 0.827, 0.202,
    { size: 12, bold: true, color, valign: 'middle', rot: 270 });
  t(s, 'FILM PRODUCTION PRESENTATION TEMPLATE', -1.604, 3.59, 4.002, 0.202,
    { size: 12, bold: true, color, align: 'center', valign: 'middle', rot: 270 });
  t(s, page, 0.082, 0.996, 0.63, 0.202,
    { size: 12, bold: true, color, align: 'right', valign: 'middle', rot: 270 });
  rule(s, 0.794, 0, 0, H, lineColor, 1.5);
}

// Letter-spaced web address used as a footer accent.
function wwwTag(s, x, y, color) {
  t(s, 'W   W   W   .   W   I   T   O   K   U   .   C   O   M', x, y, 2.3, 0.135,
    { size: 8, color });
}

// Numbered pill used by the marketing-strategies list (slide 23).
function numBadge(s, x, y, size, n) {
  s.addShape('roundRect', {
    x, y, w: size, h: size, rectRadius: 0.06,
    fill: { color: C.green }, line: { type: 'none' },
  });
  t(s, String(n), x, y, size, size,
    { face: D, size: 16, color: C.cream, align: 'center', valign: 'middle' });
}
function slide01(s) {
  bg(s, C.orange);
  rail(s, C.cream, '01/30', C.cream);
  t(s, 'FILMMAKING IS OUR PASSION, STORYTELLING IS OUR ART', 1.246, 0.687, 2.171, 0.539, { face: D, size: 16, color: C.cream });
  t(s, 'PICTURE IS WORTH A MILLION WORDS, A MOVIE IS WORTH A LIFETIME', 5.478, 0.687, 2.598, 0.539, { face: D, size: 16, color: C.cream });
  photo(s, 9.158, 0, 4.175, 7.5, BLUE);
  t(s, 'WITOKU presentation', 1.246, 1.913, 11.643, 5.587, { face: D, size: 166, color: C.cream, align: 'justify', valign: 'middle' });
}

function slide02(s) {
  bg(s, C.green);
  rail(s, C.cream, '02/30', C.cream);
  t(s, 'Interactively proactive commerce to process this outside the box into thinking pursue scalable into customers for into based services through its star shape envisioned for a good behave.', 6.316, 1.362, 2.638, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'Welcome To witoku', 6.316, 0.914, 2.283, 0.269, { face: D, size: 16, color: C.cream });
  t(s, 'Introduction Company', 9.824, 0.914, 2.283, 0.269, { face: D, size: 16, color: C.cream });
  t(s, 'Interactively proactive commerce to process this outside the box into thinking pursue scalable into customers for into based services through its star shape envisioned for a good behave.', 9.824, 1.362, 2.638, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'Message', 3.556, 5.375, 5.394, 1.935, { face: D, size: 115, color: C.cream });
  t(s, 'Welcome', 1.317, 3.364, 7.12, 2.322, { face: D, size: 138, color: C.cream });
  photo(s, 1.317, 0, 4.127, 3.174, RED);
  photo(s, 9.571, 3.174, 3.762, 4.326, BLUE);
}

function slide03(s) {
  bg(s, C.pink);
  rail(s, C.ink, '03/30', C.black);
  t(s, 'Background Company ', 5.662, 0.263, 7.526, 4.376, { face: D, size: 128, color: C.ink });
  t(s, 'Collaboratively administrate for empowered markets via plug and play networks dynamic procrastinated B2C users. After installed based on benefits dramatic with special ability to interest and bring the team.', 5.662, 6.079, 2.795, 0.884, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, [['We Make Your Stories ', {brk: 1}], ['Come To OUR Life']], 5.662, 5.182, 2.088, 0.606, { face: D, size: 18, color: C.ink, align: 'justify' });
  photo(s, 1.494, 0, 3.701, 7.5, RED);
}

function slide04(s) {
  rail(s, C.ink, '04/30', C.black);
  [[1.119, 0.357, 'Business Objective', 12.165, 2.322, 138, 3.765], [9.333, 3.423, 'Your Title Name Goes Here', 1.969, 0.236, 14, 5.046], [9.333, 4.704, 'Your Title Name Goes Here', 1.969, 0.236, 14, 6.327]].forEach(([x, y, a, b, c, d, e]) => {
    t(s, a, x, y, b, c, { face: D, size: d, color: C.ink });
    t(s, 'Leverage agile frameworks robust any synopsis high level multimedia a based frameworks.', 9.333, e, 2.52, 0.43, { color: C.ink, align: 'justify', ls: 1.5 });
  });
  t(s, 'Your Title Name Goes Here', 9.333, 5.985, 1.969, 0.236, { face: D, size: 14, color: C.ink });
  photo(s, 1.119, 2.68, 6.733, 4.82, BLUE);
}

function slide05(s) {
  bg(s, C.orange);
  rail(s, C.cream, '05/30', C.cream);
  [[0.542, 'Business Model', 10.807, 2.322, 138, 4.922], [4.594, 'Your Title Name Goes Here', 1.929, 0.236, 14, 6.3]].forEach(([y, a, b, c, d, e]) => {
    t(s, a, 1.418, y, b, c, { face: D, size: d, color: C.cream });
    t(s, 'Leverage agile frameworks robust on synopsis high level multimedia a based agile involve to frameworks provides robust good behave.', 1.418, e, 2.48, 0.657, { color: C.cream, align: 'justify', ls: 1.5 });
  });
  t(s, 'Your Title Name Goes Here', 1.418, 5.972, 1.929, 0.236, { face: D, size: 14, color: C.cream });
  photo(s, 7.476, 4.593, 5.857, 2.907, RED);
}

function slide06(s) {
  bg(s, C.green);
  rail(s, C.cream, '06/30', C.cream);
  t(s, 'Types of Film Projects', 1.09, 0.266, 6.181, 6.967, { face: D, size: 138, color: C.cream });
  t(s, 'Interactively proactive action commerce processing centric outside making on boxes into thinking expertise and on cross-media base for a base growth strategies. Seamlessly due and visualize made at proactive media a good based growth for growth to became ones.', 7.362, 0.608, 1.693, 2.02, { color: C.cream, align: 'justify', ls: 1.5 });
  [[7.875, 5.448, 'TYPES FILM PROJECT NUMBER ONE'], [10.913, 5.468, 'TYPES FILM PROJECT NUMBER THREE '], [7.875, 6.146, 'TYPES FILM PROJECT NUMBER TWO'], [10.913, 6.166, 'TYPES FILM PROJECT NUMBER FOUR']].forEach(([x, y, a]) => {
    t(s, a, x, y, 2.244, 0.236, { face: D, size: 14, color: C.cream });
    box(s, x - 0.513, y - 0.02, 0.276, 0.276, { fill: C.pink });
  });
  photo(s, 9.658, 0, 3.675, 3.798, BLUE);
}

function slide07(s) {
  bg(s, C.pink);
  rail(s, C.ink, '07/30', C.cream);
  t(s, 'Source Of Income', 1.709, 0.708, 6.84, 4.645, { face: D, size: 138, color: C.ink });
  t(s, 'We Make Your Stories Come To Life', 1.709, 5.93, 1.142, 0.808, { face: D, size: 16, color: C.ink });
  [3.953, 6.552].forEach((x) => {
    t(s, 'Your Title Name', x, 5.875, 1.299, 0.236, { face: D, size: 14, color: C.ink });
    t(s, 'Capitalize on low hanging fruit to identify a ballpark to value thing.', x, 6.135, 1.378, 0.657, { color: C.ink, ls: 1.5 });
  });
  photo(s, 8.845, 0, 4.488, 7.5, RED);
}

function slide08(s) {
  rail(s, C.ink, '08/30', C.black);
  t(s, 'Investment & Budgeting', 1.615, 0.668, 11.094, 1.616, { face: D, size: 96, color: C.ink });
  [[10.474, 4.901, 'Collaboratively the empowered markets via plug play networks procrastinate benefits.', 1.693, 0.657, 6.074], [1.615, 6.157, 'Collaboratively administrate empowered markets via plug and play networks. Dynamic procrastinate B2C users after installed base benefits. Dramatic visualize. Predominate extensible.', 2.677, 0.884, 4.59], [10.474, 6.384, 'Collaboratively the empowered markets via plug play networks procrastinate benefits.', 1.693, 0.657, 3.107]].forEach(([x, y, a, b, c, d]) => {
    t(s, a, x, y, b, c, { color: C.ink, align: 'justify', ls: 1.5 });
    t(s, 'Your Title Name', 10.474, d, 1.299, 0.236, { face: D, size: 14, color: C.ink });
  });
  t(s, 'Collaboratively the empowered markets via plug play networks procrastinate benefits.', 10.474, 3.418, 1.693, 0.657, { color: C.ink, align: 'justify', ls: 1.5 });
  photo(s, 1.615, 3.107, 5.167, 2.451, BLUE);
}

function slide09(s) {
  bg(s, C.orange);
  rail(s, C.cream, '09/30', C.cream);
  t(s, 'Collaboratively administrate for empowered markets via plugin and play networks dynamic at procrastinate B2C users. ', 6.182, 4.854, 3.46, 0.43, { color: C.cream, align: 'justify', ls: 1.5 });
  rule(s, 6.182, 6.01, 2.712, 0, C.green, 5);
  rule(s, 6.182, 6.432, 2.712, 0, C.green, 5);
  rule(s, 6.182, 6.833, 2.712, 0, C.green, 5);
  [[6.008, 2.258, 'PUBLIC RELATIONS', 5.722], [6.431, 2.441, 'GRAPHIC DESIGN', 6.141], [6.833, 2.592, 'SOCIAL MARKETING', 6.569]].forEach(([y, a, b, c]) => {
    rule(s, 6.182, y, a, 0, C.pink, 5);
    t(s, b, 6.182, c, 1.304, 0.151, { color: C.cream });
  });
  t(s, '85%', 9.291, 5.916, 0.351, 0.151, { color: C.cream });
  t(s, '90%', 9.291, 6.338, 0.351, 0.151, { color: C.cream });
  t(s, '95%', 9.291, 6.738, 0.351, 0.151, { color: C.cream });
  t(s, 'Suzy chesterr', 6.182, 3.511, 3.46, 0.673, { face: D, size: 40, color: C.cream });
  t(s, 'LEADER COMPANY A', 6.182, 4.309, 1.969, 0.269, { face: D, size: 16, color: C.cream });
  t(s, 'Our', 1.39, 0.056, 3.735, 2.794, { face: D, size: 166, color: C.cream });
  t(s, 'Leader', 6.182, 0.056, 6.555, 2.794, { face: D, size: 166, color: C.cream, align: 'right' });
  wwwTag(s, 1.39, 6.755, C.cream);
  photo(s, 10.473, 2.901, 2.86, 4.599, RED);
}

function slide10(s) {
  bg(s, C.green);
  rail(s, C.cream, '10/30', C.cream);
  t(s, 'Renata kloe', 6.188, 2.919, 1.693, 0.404, { face: D, size: 24, color: C.cream });
  t(s, 'Your job here', 6.188, 3.319, 1.102, 0.236, { face: D, size: 14, color: C.cream });
  t(s, 'PLACEHOLDER', 6.188, 3.75, 2.362, 0.657, { color: C.cream, ls: 1.5 });
  t(s, 'Kyle phillipe', 1.792, 2.919, 1.693, 0.404, { face: D, size: 24, color: C.cream });
  t(s, 'Your job here', 1.792, 3.319, 1.102, 0.236, { face: D, size: 14, color: C.cream });
  t(s, 'PLACEHOLDER', 1.792, 3.75, 2.362, 0.657, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'Our Team', 4.517, 5.178, 8.379, 2.322, { face: D, size: 138, color: C.cream, align: 'right' });
  wwwTag(s, 1.792, 6.808, C.cream);
  photo(s, 6.188, 0, 3.604, 2.62, RED);
  photo(s, 1.792, 0, 3.604, 2.62, BLUE);
}

function slide11(s) {
  bg(s, C.pink);
  rail(s, C.ink, '11/30', C.black);
  [1.266, 3.839, 6.412].forEach((y) => {
    t(s, 'Bring table win win survival to ensure dominan at the end of the day, with going forward with team good.', 6.614, y, 2.008, 0.657, { color: C.ink, align: 'justify', ls: 1.5 });
    t(s, 'YOUR NAME HERE', 6.614, y - 0.835, 1.417, 0.269, { face: D, size: 16, color: C.ink });
    t(s, 'YOUR JOB HERE', 6.614, y - 0.449, 1.142, 0.236, { face: D, size: 14, color: C.ink });
  });
  t(s, 'Our Team', 1.869, 2.605, 3.86, 4.645, { face: D, size: 138, color: C.ink });
  photo(s, 9.854, 0, 3.479, 2.355, BLUE);
  photo(s, 9.854, 2.573, 3.479, 2.355, RED);
  photo(s, 9.854, 5.145, 3.479, 2.355, BLUE);
}

function slide12(s) {
  rail(s, C.ink, '12/30', C.black);
  t(s, 'Creative Process', 1.35, 3.629, 5.316, 3.871, { face: D, size: 115, color: C.ink });
  t(s, 'Interactively proactive on commerce process centric outside making box into thinking pursue scalable customer into making for based services multimedia for based into expertise nd cross-media base for in base strategies. ', 1.35, 1.553, 4.567, 0.657, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'We Make Your Stories Come To Life', 1.35, 0.822, 2.835, 0.269, { face: D, size: 16, color: C.ink });
  photo(s, 8.976, 0, 4.357, 3.62, BLUE);
  photo(s, 8.976, 4.473, 4.357, 3.027, RED);
}

function slide13(s) {
  bg(s, C.orange);
  rail(s, C.cream, '13/30', C.cream);
  [[3.38, 'YOUR TITLE NAME HERE', 1.467, 0.236, 14, 1.848], [5.373, 'SCRIPT DEVELOPMENT', 11.119, 1.935, 115, 3.736]].forEach(([y, a, b, c, d, e]) => {
    t(s, a, 1.421, y, b, c, { face: D, size: d, color: C.cream });
    t(s, 'Interactively proactive this commerce of process centric star outside the boxes into thinking pursue scalable and customers for into state based services through stated its shape. Proactively envisioned on for.', 1.421, e, 2.992, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
  });
  t(s, 'YOUR TITLE NAME HERE', 1.421, 1.491, 1.467, 0.236, { face: D, size: 14, color: C.cream });
  t(s, 'We Make Your Stories Come To Life', 1.421, 0.574, 2.795, 0.269, { face: D, size: 16, color: C.cream });
  photo(s, 8.343, 0, 4.99, 5.194, RED);
}

function slide14(s) {
  bg(s, C.green);
  rail(s, C.cream, '14/30', C.cream);
  t(s, 'Casting & Crewing', 6.082, 0.494, 5.85, 3.871, { face: D, size: 115, color: C.cream });
  [6.082, 9.408].forEach((x) => {
    t(s, 'Interactively proactive on commerce process centric outside making box into thinking pursue scalable customer into making for based services multimedia for based in expertise.', x, 6.122, 2.598, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
    t(s, 'Your Title Name Goes Here', x, 5.689, 2.008, 0.236, { face: D, size: 14, color: C.cream, align: 'justify' });
  });
  t(s, 'We Make Your Stories Come To Life', 6.082, 4.859, 3.465, 0.337, { face: D, size: 20, color: C.cream });
  photo(s, 1.401, 0, 3.995, 7.5, BLUE);
}

function slide15(s) {
  bg(s, C.pink);
  rail(s, C.ink, '15/30', C.black);
  t(s, 'Interactively proactive on commerce process centric outside making box into thinking pursue scalable the customer into for making based services multimedia for based make expertise and cross-media base for in base strategies. Seamlessly visualize made quality intellectual capital in base without into the on sharing into growth coordinate proactive media good based growth for markets pursue.', 9.202, 2.851, 3.185, 1.566, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Filming and production', 1.402, 0.567, 11.635, 1.717, { face: D, size: 102, color: C.ink });
  t(s, 'We Make Your Stories Come To Life', 9.591, 6.36, 2.795, 0.269, { face: D, size: 16, color: C.ink, align: 'right' });
  photo(s, 1.402, 2.851, 3.909, 3.778, RED);
}

function slide16(s) {
  rail(s, C.ink, '16/30', C.black);
  t(s, 'Time For Break', 2.622, 3.238, 5.315, 3.871, { face: D, size: 115, color: C.ink });
  t(s, 'We Make Your Stories Come To Life', 2.622, 0.782, 3.465, 0.337, { face: D, size: 20, color: C.ink });
  t(s, 'Interactively proactive commerce process centric outside the box into thinking pursue the scalable customer into for based services through its shape. Proactively envisioned for a multimedia based in expertise and cross-media for good on growth strategies. ', 2.622, 1.393, 3.504, 0.884, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Break for 30 minutes', 2.622, 2.551, 2.52, 0.269, { face: D, size: 16, color: C.ink });
  photo(s, 8.922, 0, 3.995, 7.5, BLUE);
}

function slide17(s) {
  bg(s, C.orange);
  rail(s, C.cream, '17/30', C.cream);
  t(s, 'Production', 1.497, 0, 8.977, 2.322, { face: D, size: 138, color: C.cream });
  rule(s, 3.148, 3.717, 8.159, 0, C.cream, 1.5);
  [[1.497, 'Filming Tecniques', 1.808, 1.34, '01'], [4.554, 'Post-Production', 4.782, 1.507, '02'], [7.611, 'Editing and Sound', 7.922, 1.34, '03'], [10.668, 'Visual Effects', 10.977, 1.34, '04']].forEach(([x, a, b, c, d]) => {
    box(s, x, 2.922, 1.963, 1.59, { fill: C.green });
    t(s, a, b, 3.902, c, 0.185, { face: D, size: 11, color: C.cream, align: 'center' });
    t(s, d, x + 0.7, 3.347, 0.562, 0.539, { face: D, size: 32, color: C.cream, align: 'center' });
  });
  [...Array(4).keys()].forEach((i) => t(s, 'Interactively proactive on commerce process centric outside the box into thinking pursue this event.', 1.497 + i * 3.057, 4.628, 1.963, 0.657, { color: C.cream, align: 'justify', ls: 1.5 }));
  t(s, 'Planning', 5.421, 5.884, 7.209, 1.616, { face: D, size: 96, color: C.cream, align: 'right' });
  wwwTag(s, 1.497, 6.941, C.cream);
}

function slide18(s) {
  bg(s, C.green);
  rail(s, C.cream, '18/30', C.cream);
  [[0.917, 0, 'Filming Techniques', 12.362, 2.322, 138, 'right', 5.731], [1.295, 5.287, 'Your Title Name Here', 2.461, 0.236, 14, 'justify', 3.353]].forEach(([x, y, a, b, c, d, e, f]) => {
    t(s, a, x, y, b, c, { face: D, size: d, color: C.cream, align: e });
    t(s, 'Collaboratively administrate empowered markets it plug and play networks dynamic procrastinate B2C users. After installed base benefits dramatic with special products with coordinate include proactive ecommerce good process centric outside proactively good envisioned multimedia generation.', 1.295, f, 2.756, 1.339, { color: C.cream, align: 'justify', ls: 1.5 });
  });
  t(s, 'Your Title Name Here', 1.295, 2.909, 2.461, 0.236, { face: D, size: 14, color: C.cream, align: 'justify' });
  t(s, 'We Make Your Stories Come To Life', 5.407, 2.909, 3.859, 0.337, { face: D, size: 20, color: C.cream, align: 'center' });
  photo(s, 10.622, 2.479, 2.711, 5.021, RED);
}

function slide19(s) {
  bg(s, C.pink);
  rail(s, C.ink, '19/30', C.cream);
  [[6.408, 0.515, 'Company Growth', 1.969, 0.337, 20, 8.951], [1.77, 3.528, 'Post-Production', 6.213, 3.231, 96, 6.408]].forEach(([x, y, a, b, c, d, e]) => {
    t(s, a, x, y, b, c, { face: D, size: d, color: C.ink });
    t(s, 'Leverage agile frameworks robust on synopsis high multimedia based agile frameworks robust into media into full multimedia molecule for molecule in a market growth strategies.', e, 1.29, 2.008, 1.111, { color: C.ink, align: 'justify', ls: 1.5 });
  });
  photo(s, 9.116, 3.519, 4.218, 3.25, RED);
  photo(s, 1.664, 0, 4.316, 2.917, BLUE);
}

function slide20(s) {
  rail(s, C.ink, '20/30', C.black);
  t(s, 'Collaboratively good administrate empowered markets via plug and play networks. Dynamic procrastinate B2C users. After installed base benefits dramatic witinh special with products with coordinate proactive ecommerce good processed centric outside proactively to envisioned multimedia collaboration.', 7.852, 6.154, 4.409, 0.884, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Editing Film Productions', 7.852, 5.09, 2.711, 0.337, { face: D, size: 20, color: C.ink });
  t(s, 'Editing and \nSound Design', 1.582, 0.372, 5.799, 2.962, { face: D, size: 88, color: C.ink });
  t(s, 'Collaboratively good administrate empowered markets via plug and play networks. Dynamic procrastinate B2C users. After installed base benefits dramatic with special with products with coordinate.', 1.582, 3.386, 3.976, 0.657, { color: C.ink, align: 'justify', ls: 1.5 });
  photo(s, 7.852, 0.596, 5.481, 3.447, BLUE);
  photo(s, 1.582, 5.09, 3.976, 2.41, RED);
}

function slide21(s) {
  bg(s, C.orange);
  rail(s, C.cream, '21/30', C.cream);
  t(s, 'Proactively envisioned multimedia based expertise and cross media growth strategies. Seamlessly visualize quality intellectual capital without superior. Collaboration. Leverage agile frameworks to provide a robust. Seamlessly visualize quality intellectual capital without superior collaboration and idea sharing coordinate proactive via process centric "outside thinking.', 7.715, 5.278, 5.195, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'Visual Effects and CGI', 1.217, 0.352, 12.008, 1.935, { face: D, size: 115, color: C.cream });
  t(s, 'Collaboratively good administrate empowered markets via plug and play networks. Dynamic creating procrastinate B2C users. After installed base benefits dramatic with special with products.', 1.217, 2.639, 1.811, 1.339, { color: C.cream, align: 'justify', ls: 1.5 });
  wwwTag(s, 1.217, 6.028, C.cream);
  photo(s, 7.715, 2.639, 5.195, 2.221, BLUE);
}

function slide22(s) {
  bg(s, C.green);
  rail(s, C.cream, '22/30', C.cream);
  t(s, 'Marketing & Distribution', 5.462, 0.382, 7.522, 3.871, { face: D, size: 115, color: C.cream, align: 'right' });
  [[6.312, 'MARKETING'], [10.544, 'DISTRIBUTION']].forEach(([x, a]) => {
    t(s, 'Collaboratively empowered markets this plug and play networks dynamic procrastinate B2C users. After installed base benefits dramatic with special products with coordinate.', x, 5.509, 2.441, 0.884, { color: C.cream, align: 'justify', ls: 1.5 });
    t(s, a, x, 5.128, 2.126, 0.269, { face: D, size: 16, color: C.cream });
  });
  t(s, 'We Make Your Stories Come To Life', 1.09, 0.653, 1.417, 1.01, { face: D, size: 20, color: C.cream });
  photo(s, 1.09, 5.128, 4.077, 2.372, RED);
}

function slide23(s) {
  bg(s, C.pink);
  rail(s, C.ink, '23/30', C.black);
  t(s, 'Collaboratively administrate on empowered markets it plug and play networks dynamic procrastinate B2C users. ', 1.187, 1.452, 2.323, 0.657, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'We Make Your Stories Come To Life', 5.345, 1.452, 2.283, 0.673, { face: D, size: 20, color: C.ink });
  [...Array(4).keys()].forEach((i) => t(s, 'Your Title Name Goes Here', 1.82, 4.995 + i * -0.713, 2.008, 0.236, { face: D, size: 14, color: C.ink }));
  photo(s, 9.188, 0, 4.146, 7.5, BLUE);
  t(s, 'Film Marketing strategies', 1.187, 5.977, 11.352, 1.481, { face: D, size: 88, color: C.ink });
  [[2.768, 1], [3.48, 2], [4.193, 3], [4.906, 4]].forEach(([y, a]) => numBadge(s, 1.187, y, 0.413, a));
}

function slide24(s) {
  rail(s, C.ink, '24/30', C.black);
  [1.437, 4.427, 7.417, 10.406].forEach((x) => box(s, x, 1.24, 2.283, 1.646, { fill: C.grey }));
  [1.437, 4.427, 7.417, 10.406].forEach((x) => {
    t(s, 'PLACEHOLDER', x, 3.637, 2.283, 0.43, { color: C.ink, align: 'justify', ls: 1.5 });
    t(s, 'YOUR TITLE NAME GOES HERE', x, 3.325, 2.126, 0.236, { face: D, size: 14, color: C.ink });
  });
  t(s, 'DISTRIBUTION', 1.437, 4.807, 5.273, 1.346, { face: D, size: 80, color: C.ink });
  t(s, 'We Make Your Stories Come To Life', 1.437, 6.664, 3.504, 0.337, { face: D, size: 20, color: C.ink });
  t(s, 'PARTNERSHIP', 7.125, 5.891, 5.565, 1.346, { face: D, size: 80, color: C.ink, align: 'right' });
  [1.437, 7.417].forEach((x) => {
    photo(s, x, 1.24, 2.283, 1.646, RED);
    photo(s, x + 2.99, 1.24, 2.283, 1.646, BLUE);
  });
}

function slide25(s) {
  bg(s, C.orange);
  rail(s, C.cream, '25/30', C.cream);
  t(s, 'Financial Performance', 1.481, 4.488, 5.46, 2.693, { face: D, size: 80, color: C.cream });
  [4.956, 6.165].forEach((y) => {
    t(s, 'Your Title Name Here', 7.628, y, 1.811, 0.236, { face: D, size: 14, color: C.cream });
    t(s, 'Collaboratively administrate empowered markets plug and play networks.', 7.628, y + 0.305, 2.165, 0.43, { color: C.cream, align: 'justify', ls: 1.5 });
    t(s, 'Your Title Name Here', 10.481, y, 1.811, 0.236, { face: D, size: 14, color: C.cream });
    t(s, 'Collaboratively administrate empowered markets plug and play networks.', 10.481, y + 0.305, 2.165, 0.43, { color: C.cream, align: 'justify', ls: 1.5 });
  });
  t(s, 'Collaboratively administrate any copy empowered markets it plug and play networks dynamic procrastinate B2C users. After installed base benefits dramatic with special products.', 1.481, 2.124, 2.008, 1.112, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'We Make Your Stories Come To Life', 1.481, 1.291, 2.205, 0.673, { face: D, size: 20, color: C.cream });
  rule(s, 6.471, 0.8, 0, 2.824, C.cream, 0.25);
  rule(s, 6.419, 3.569, 6.226, 0, C.cream, 0.25);
  ['1,400', '1,200', '1,000', '800', '600', '400', '200', '0'].forEach((v, i) => t(s, v, 5.93, 0.706 + i * 0.395, 0.404, 0.151, { color: C.cream }));
  ['2023', '2024', '2025', '2026', '2027'].forEach((v, i) => t(s, v, 6.793 + i * 1.232, 3.669, 0.605, 0.151, { color: C.cream, bold: true, align: 'center' }));
  [[7.095, 1.682, 1.887, C.green], [6.692, 2.08, 1.489, C.pink], [9.56, 1.795, 1.774, C.green], [9.157, 1.586, 1.983, C.pink], [10.793, 2.464, 1.105, C.green], [10.389, 2.372, 1.196, C.pink], [12.025, 3.09, 0.478, C.green], [11.622, 3.182, 0.387, C.pink], [8.328, 1.015, 2.554, C.green], [7.924, 1.437, 2.132, C.pink]].forEach(([x, y, h, fill]) => box(s, x, y, 0.403, h, { fill }));
  [['759', 6.715, 1.803], ['954', 7.118, 1.402], ['1093', 7.947, 1.163], ['1297', 8.351, 0.745], ['1020', 9.175, 1.313], ['897', 9.583, 1.529], ['603', 10.412, 2.113], ['294', 10.816, 2.203], ['195', 11.645, 2.91], ['238', 12.048, 2.824]].forEach(([v, x, y]) => t(s, v, x, y, 0.358, 0.151, { color: C.cream, align: 'center' }));
}

function slide26(s) {
  bg(s, C.green);
  rail(s, C.cream, '26/30', C.cream);
  t(s, 'Revenue Tracking & Reporting', 1.436, 0.66, 11.602, 1.346, { face: D, size: 80, color: C.cream });
  t(s, 'Collaboratively administrate empowered markets it plug and play networks dynamic procrastinate B2C users. After installed base benefits dramatic.', 1.436, 5.961, 2.668, 0.657, { color: C.cream, align: 'justify', ls: 1.5 });
  [[2.627, 2.624, 0.286, 2.905, 3.461], [3.425, 4.018, 0.278, 3.703, 2.624], [4.222, 3.11, 0.278, 4.492, 4.331]].forEach(([y, a, b, c, d]) => {
    box(s, 1.992, y, a, b, { fill: C.orange });
    box(s, 1.992, c, d, b, { fill: C.pink });
  });
  rule(s, 1.981, 2.42, 0, 2.617, C.cream, 1.5);
  rule(s, 1.981, 5.037, 4.762, 0, C.cream, 1.5);
  ['Items Three', 'Items Two', 'Items One'].forEach((v, i) => t(s, v, 1.436, 4.361 + i * -0.794, 0.454, 0.337, { size: 10, color: C.cream, valign: 'middle' }));
  ['0', '1', '2', '3', '4', '5', '6', '7'].forEach((v, i) => t(s, v, 1.981 + i * 0.615, 5.201, 0.338, 0.135, { size: 8, color: C.cream, align: 'center', valign: 'middle' }));
  t(s, 'Collaboratively administrate empowered markets it plug and play networks dynamic procrastinate B2C users. After installed base benefits dramatic.', 7.384, 5.961, 2.669, 0.657, { color: C.cream, align: 'justify', ls: 1.5 });
  [[2.627, 2.624, 0.286, 2.905, 4.449], [3.425, 3.425, 0.278, 3.703, 3.74], [4.222, 3.11, 0.278, 4.492, 2.323]].forEach(([y, a, b, c, d]) => {
    box(s, 7.941, y, a, b, { fill: C.orange });
    box(s, 7.941, c, d, b, { fill: C.pink });
  });
  rule(s, 7.93, 2.42, 0, 2.617, C.cream, 1.5);
  rule(s, 7.93, 5.037, 4.762, 0, C.cream, 1.5);
  ['Items Three', 'Items Two', 'Items One'].forEach((v, i) => t(s, v, 7.384, 4.361 + i * -0.794, 0.454, 0.337, { size: 10, color: C.cream, valign: 'middle' }));
  ['0', '1', '2', '3', '4', '5', '6', '7'].forEach((v, i) => t(s, v, 7.93 + i * 0.615, 5.201, 0.338, 0.135, { size: 8, color: C.cream, align: 'center', valign: 'middle' }));
}

function slide27(s) {
  bg(s, C.pink);
  rail(s, C.ink, '27/30', C.black);
  t(s, 'Return on Investment', 4.817, 0.477, 7.96, 1.212, { face: D, size: 72, color: C.ink, align: 'right' });
  t(s, 'Collaboratively administrate empowered markets it plug and play networking dynamic procrastinate B2C users. After installed any base benefits dramatic with create special products with coordinate.', 11.202, 5.225, 1.575, 1.566, { face: 'PT Sans', color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Your Title Name', 1.344, 5.821, 1.496, 0.236, { face: D, size: 14, color: C.ink });
  t(s, 'Collaboratively administrate empowered markets plug and play networks dynamic on procrastinate B2C users. .', 1.344, 6.133, 2.301, 0.657, { face: 'PT Sans', color: C.ink, align: 'justify', ls: 1.5 });
  rule(s, 1.892, 2.605, 0, 2.568, C.black, 0.25);
  rule(s, 1.834, 5.111, 1.811, 0, C.black, 0.25);
  ['1,000', '800', '600', '400', '200', '0'].forEach((v, i) => t(s, v, 1.344, 2.787 + i * 0.442, 0.394, 0.151, { color: C.ink, align: 'right' }));
  box(s, 2.75, 3.404, 0.452, 1.707, { fill: C.green });
  box(s, 2.139, 3.843, 0.452, 1.268, { fill: C.orange });
  t(s, '601', 2.165, 3.571, 0.401, 0.151, { color: C.ink, align: 'center' });
  t(s, '796', 2.776, 3.122, 0.401, 0.151, { color: C.ink, align: 'center' });
  t(s, 'Data 1', 2.168, 5.27, 0.394, 0.151, { color: C.ink });
  t(s, 'Data 2', 2.78, 5.27, 0.394, 0.151, { color: C.ink });
  t(s, 'Collaboratively administrate empowered markets plug and play networks dynamic on procrastinate B2C users. .', 4.27, 6.133, 2.301, 0.657, { face: 'PT Sans', color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Your Title Name', 4.27, 5.821, 1.496, 0.236, { face: D, size: 14, color: C.ink });
  rule(s, 4.817, 2.605, 0, 2.568, C.black, 0.25);
  rule(s, 4.759, 5.111, 1.811, 0, C.black, 0.25);
  ['1,000', '800', '600', '400', '200', '0'].forEach((v, i) => t(s, v, 4.27, 2.787 + i * 0.442, 0.394, 0.151, { color: C.ink, align: 'right' }));
  box(s, 5.676, 3.89, 0.452, 1.221, { fill: C.green });
  box(s, 5.065, 3.363, 0.452, 1.748, { fill: C.orange });
  t(s, '801', 5.09, 3.122, 0.401, 0.151, { color: C.ink, align: 'center' });
  t(s, '706', 5.701, 3.6, 0.401, 0.151, { color: C.ink, align: 'center' });
  t(s, 'Data 1', 5.094, 5.27, 0.394, 0.151, { color: C.ink });
  t(s, 'Data 2', 5.705, 5.27, 0.394, 0.151, { color: C.ink });
  t(s, 'Your Title Name', 7.197, 5.821, 1.496, 0.236, { face: D, size: 14, color: C.ink });
  t(s, 'Collaboratively administrate empowered markets plug and play networks dynamic on procrastinate B2C users. .', 7.197, 6.133, 2.301, 0.657, { face: 'PT Sans', color: C.ink, align: 'justify', ls: 1.5 });
  rule(s, 7.759, 2.605, 0, 2.568, C.black, 0.25);
  rule(s, 7.701, 5.111, 1.811, 0, C.black, 0.25);
  ['1,000', '800', '600', '400', '200', '0'].forEach((v, i) => t(s, v, 7.197, 2.787 + i * 0.442, 0.394, 0.151, { color: C.ink, bold: true, align: 'right' }));
  box(s, 8.617, 3.228, 0.452, 1.882, { fill: C.green });
  box(s, 8.006, 3.491, 0.452, 1.62, { fill: C.orange });
  t(s, '701', 8.032, 3.197, 0.401, 0.151, { color: C.ink, bold: true, align: 'center' });
  t(s, '896', 8.643, 2.958, 0.401, 0.151, { color: C.ink, bold: true, align: 'center' });
  t(s, 'Data 1', 8.035, 5.27, 0.394, 0.151, { color: C.ink, bold: true });
  t(s, 'Data 2', 8.646, 5.27, 0.394, 0.151, { color: C.ink, bold: true });
}

function slide28(s) {
  rail(s, C.ink, '28/30', C.black);
  t(s, 'Collaboratively administrate empowered markets it plug and play networks dynamic procrastinate for B2C users. After installed base benefits dramatic special products.', 1.762, 2.335, 1.496, 1.339, { color: C.ink, align: 'justify', ls: 1.5 });
  t(s, 'Upcoming Project', 1.762, 0.691, 6.286, 1.111, { face: D, size: 66, color: C.ink });
  t(s, 'See You On The Next Project!', 1.762, 6.405, 3.543, 0.404, { face: D, size: 24, color: C.ink });
  [[4.698, 'monday, 23 January 2025'], [7.627, 'Sunday, 11 April 2025'], [10.556, 'Friday, 21 August 2025']].forEach(([x, a]) => {
    t(s, 'The Name Of Project Film', x, 5.327, 2.008, 0.202, { face: D, size: 12, color: C.ink });
    t(s, a, x, 4.966, 2.008, 0.269, { face: D, size: 16, color: C.ink, valign: 'middle' });
  });
  photo(s, 4.698, 2.335, 2.008, 2.099, RED);
  photo(s, 7.627, 2.335, 2.008, 2.099, BLUE);
  photo(s, 10.556, 2.335, 2.008, 2.099, RED);
}

function slide29(s) {
  bg(s, C.orange);
  rail(s, C.cream, '29/30', C.cream);
  t(s, 'Image Gallery', 1.429, 0.174, 3.81, 3.231, { face: D, size: 96, color: C.cream });
  [[5.476, 0, 9.524], [9.524, 3.921, 5.476]].forEach(([x, y, a]) => {
    photo(s, x, y, 3.81, 3.579, BLUE);
    photo(s, a, y, 3.81, 3.579, RED);
  });
  photo(s, 1.429, 3.921, 3.81, 3.579, BLUE);
}

function slide30(s) {
  bg(s, C.green);
  rail(s, C.cream, '30/30', C.cream);
  t(s, 'Proactively envisioned for quality on based multimedia based expertise and a cross-media making good growth strategies.', 10.245, 1.743, 2.244, 0.657, { color: C.cream, align: 'justify', ls: 1.5 });
  t(s, 'See You On The\nNext Project!', 1.638, 5.018, 1.142, 1.616, { face: D, size: 24, color: C.cream });
  t(s, 'Thanks\nFor watching', 1.638, 0.456, 7.485, 3.231, { face: D, size: 96, color: C.cream });
  photo(s, 4.45, 4.151, 8.883, 3.349, RED);
}


/* ------------------------------------------------------- icon sheet pages */

// Slides 31-42 are pictogram reference sheets: a 10-column grid of small black
// glyphs. The originals are custom vector paths; here each cell is drawn with a
// native pptxgenjs shape from a repeating palette so the pages read the same.
const GLYPHS = [
  'ellipse', 'rect', 'roundRect', 'triangle', 'diamond', 'star5', 'hexagon',
  'plus', 'pie', 'chevron', 'octagon', 'heart', 'sun', 'moon', 'donut',
  'trapezoid', 'pentagon', 'teardrop', 'star4', 'lightningBolt', 'cloud',
  'frame', 'halfFrame', 'cube', 'can', 'smileyFace', 'blockArc', 'arc',
  'rightArrow', 'upArrow', 'leftRightArrow', 'bentArrow', 'noSmoking',
  'star6', 'parallelogram', 'wave',
];

const ICON_COL0 = 1.383, ICON_DX = 1.1761, ICON_W = 0.41, ICON_H = 0.38;

// rows: [yCenter, count] per grid row.
function iconSheet(s, seed, rows) {
  let k = seed;
  rows.forEach(([cy, count]) => {
    for (let i = 0; i < count; i++) {
      const cx = ICON_COL0 + i * ICON_DX;
      s.addShape(GLYPHS[k++ % GLYPHS.length], {
        x: cx - ICON_W / 2, y: cy - ICON_H / 2, w: ICON_W, h: ICON_H,
        fill: { color: C.black }, line: { type: 'none' },
      });
    }
  });
}

// Even grid: `n` rows of 10, first row centred at y0, `dy` apart.
function evenRows(n, y0, dy, lastCount = 10) {
  return Array.from({ length: n }, (_, i) => [y0 + i * dy, i === n - 1 ? lastCount : 10]);
}

const ICON_PAGES = [
  { seed:  0, rows: evenRows(7, 0.899, 0.949) },   // 31
  { seed:  7, rows: evenRows(7, 0.879, 0.957) },   // 32
  { seed: 14, rows: evenRows(7, 0.882, 0.953) },   // 33
  { seed: 21, rows: evenRows(7, 0.893, 0.950) },   // 34
  { seed: 28, rows: evenRows(7, 0.905, 0.948) },   // 35
  { seed: 35, rows: evenRows(7, 0.879, 0.953) },   // 36
  { seed:  3, rows: evenRows(7, 0.892, 0.953) },   // 37
  { seed: 11, rows: evenRows(7, 0.905, 0.952) },   // 38
  { seed: 19, rows: evenRows(8, 0.878, 0.814) },   // 39 — denser, 8 rows
  { seed: 27, rows: evenRows(7, 0.905, 0.948) },   // 40
  { seed: 33, rows: evenRows(7, 0.883, 0.953) },   // 41
  { seed:  5, rows: evenRows(7, 0.904, 0.957, 7) },// 42 — final row is short
];

/* -------------------------------------------------------------- assembling */

const CONTENT_SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: W, height: H });
  pres.layout = 'WIDE';
  pres.author = 'WITOKU';
  pres.title = 'Film Production Presentation Template';

  CONTENT_SLIDES.forEach((fn) => fn(pres.addSlide()));
  ICON_PAGES.forEach((page) => iconSheet(pres.addSlide(), page.seed, page.rows));

  const out = path.join(__dirname, path.basename(__filename, '.js') + '.pptx');
  return pres.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => { console.error(err); process.exit(1); });
