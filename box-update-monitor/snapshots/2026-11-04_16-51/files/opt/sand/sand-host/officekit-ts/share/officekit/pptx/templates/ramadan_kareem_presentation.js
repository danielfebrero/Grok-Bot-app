/**
 * Ramadhan Presentation Template — rebuilt with pptxgenjs.
 *
 * A faithful, code-readable reconstruction of the 15-slide reference deck
 * (13.333in x 7.5in / 16:9).  Every photograph in the original is replaced by a
 * flat colour placeholder shape that keeps the photo's outline, position and
 * dominant colour (see `photo()`).
 *
 * Run: node 195e00d4-f046-41d9-8b03-466c7e7e16d7_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const C = {
  dark: '1B3029', // deep green — brand colour, card fills, dark backgrounds
  white: 'FFFFFF',
  black: '000000',
  gray: '7F7F7F', // "disabled" schedule icons
  // dominant colours of the reference photographs (used by the placeholders)
  photoLight: '6C7A89',
  photoMid: '4C5A69',
  photoDark: '2C3A49',
};

const F = {
  head: 'Open Sans', // headings, buttons, italic captions
  body: 'Roboto', // paragraph copy, calendar
  quote: 'Noto Serif', // the pull-quote on the break slide
};

/* Stroke widths in points. `hairline` is the thinnest stroke PowerPoint draws
 * (the reference leaves <a:ln w> unset on stars and schedule rules); pptxgenjs
 * substitutes 1pt for a missing width, so ask for a sub-pixel value instead. */
const LINE = { hairline: 0.01, thin: 1, card: 1.5 };

/* Lorem copy, built up in the same increments the reference deck uses. */
const L_SENTENCE = 'Lorem Ipsum\u00A0is simply dummy text of the printing and typesetting industry. ';
const L_SHORT = L_SENTENCE.trim() + " Lorem Ipsum has been the industry's standard dummy text ever since the 1500s";
const L_MED = L_SHORT + ', when an unknown printer took a galley of type and scrambled it to make a type specimen book. ';
const L_LONG = L_MED + 'It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. ';
const L_FULL = L_LONG + 'It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, ' +
  'and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.';

const TEMPLATE_LABEL = 'Ramadhan Presentation Template ';

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** Default corner radius PowerPoint gives a `roundRect` (adj = 16667). */
const rr = (w, h) => 0.16667 * Math.min(w, h);

/**
 * pptxgenjs writes `angleRange:[a,b]` straight out as <a:gd name="adj1|adj2"
 * fmla="val a*60000"/>, which is exactly how the corner-radius adjust values of
 * an arch (`round2SameRect`) are expressed. adj1 = radius as % of the short side.
 */
const adj = (adj1, adj2 = 0) => [adj1 / 60000, adj2 / 60000];
const ARCH = adj(50000); // fully-rounded top corners

/** Text box. The reference uses top-anchored, auto-growing ("resize") boxes. */
function T(slide, text, opts) {
  return slide.addText(text, Object.assign({
    fontFace: F.head, fontSize: 18, color: C.black, valign: 'top', fit: 'resize',
  }, opts));
}

/** 10pt Roboto body copy at 1.5 line spacing — the deck's workhorse paragraph. */
function body(slide, text, opts) {
  return T(slide, text, Object.assign({ fontFace: F.body, fontSize: 10, lineSpacingMultiple: 1.5 }, opts));
}

/** Bold display heading (Open Sans). */
function heading(slide, text, opts) {
  return T(slide, text, Object.assign({ bold: true, fontSize: 44 }, opts));
}

/** Small filled bullet dot. */
function dot(slide, x, y, size, color) {
  slide.addShape('ellipse', { x, y, w: size, h: size, fill: { color } });
}

/**
 * Stand-in for a raster photo from the reference deck: same geometry, filled
 * with the photo's dominant colour. `shape` keeps the original crop outline
 * (plain rectangle, seven-point star or arch).
 */
function photo(slide, o) {
  slide.addShape(o.shape || 'rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.color },
    angleRange: o.shape === 'round2SameRect' ? ARCH : undefined,
  });
}

/** Full-bleed photo placeholder covering the whole slide. */
function photoBackdrop(slide, color) {
  photo(slide, { x: 0, y: 0, w: 13.333, h: 7.5, color });
}

/* Office "Checkmark" / "Close" icons, traced from the reference SVGs on a
 * 96x96 grid so they can be drawn as native custom-geometry shapes. */
const ICON_CHECK = [[86.1, 15.8], [34.9, 64.2], [10.3, 39], [1.8, 47.1], [34.5, 80.7], [43.1, 72.7], [94.2, 24.2]];
const ICON_CROSS = [[83.4, 21.1], [74.9, 12.6], [48, 39.5], [21.1, 12.6], [12.6, 21.1], [39.5, 48],
  [12.6, 74.9], [21.1, 83.4], [48, 56.5], [74.9, 83.4], [83.4, 74.9], [56.5, 48]];

function icon(slide, outline, x, y, size, color) {
  slide.addShape('custGeom', {
    x, y, w: size, h: size,
    fill: { color },
    points: outline.map(([px, py]) => ({ x: (px / 96) * size, y: (py / 96) * size })).concat([{ close: true }]),
  });
}

/* ------------------------------------------------------------------ *
 * Slide 1 — cover
 * ------------------------------------------------------------------ */

function slideCover(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);

  T(s, TEMPLATE_LABEL, { x: 0.819, y: 1.679, w: 7.264, h: 0.404, color: C.white });
  T(s, 'Ramadhan Kareem ', { x: 0.708, y: 2.083, w: 7.264, h: 2.794, fontSize: 80, bold: true, color: C.white });

  // "Learn More" pill (adj 50000 => fully rounded ends)
  s.addShape('roundRect', { x: 0.819, y: 5.287, w: 1.972, h: 0.403, fill: { color: C.white }, rectRadius: 0.403 / 2 });
  T(s, 'Learn More', { x: 0.819, y: 5.287, w: 1.972, h: 0.404, bold: true, color: C.dark, align: 'center' });

  body(s, L_SHORT, { x: 0.708, y: 5.764, w: 6.403, h: 0.58, fontFace: F.head, italic: true, color: C.white });
}

/* ------------------------------------------------------------------ *
 * Slide 2 — "Get Ready Ramadhan", three feature cards on white
 * ------------------------------------------------------------------ */

function slideGetReady(pres) {
  const s = pres.addSlide();

  heading(s, 'Get Ready Ramadhan', { x: 4.403, y: 0.75, w: 8.139, h: 0.909, fontSize: 48, color: C.dark, align: 'right' });
  body(s, L_FULL, { x: 1.222, y: 1.793, w: 11.319, h: 1.083, align: 'right' });

  [1.271, 5.132, 8.931].forEach((x) => {
    s.addShape('roundRect', { x, y: 3.431, w: 3.611, h: 3.319, fill: { color: C.dark }, rectRadius: rr(3.611, 3.319) });
    s.addShape('roundRect', {
      x: x + 0.25, y: 3.66, w: 3.111, h: 2.86,
      line: { color: C.white, width: LINE.card }, rectRadius: rr(3.111, 2.86),
    });
    T(s, 'Get Ready Ramadhan', { x: x + 0.354, y: 3.981, w: 2.819, h: 0.37, fontSize: 16, bold: true, color: C.white, align: 'right' });
    [[4.589, 4.764], [5.303, 5.479]].forEach(([ty, dy]) => {
      body(s, L_SENTENCE, { x: x + 0.25, y: ty, w: 2.687, h: 0.578, color: C.white, align: 'right' });
      dot(s, x + 2.937, dy, 0.236, C.white);
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 3 — "Introduction To Ramadhan", two star cut-outs
 * ------------------------------------------------------------------ */

function slideIntroduction(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);

  // white star outlines sitting just behind/offset from the star photos
  s.addShape('star7', { x: 8.474, y: 0.218, w: 4.115, h: 4.115, line: { color: C.white, width: LINE.hairline } });
  s.addShape('star7', { x: 4.825, y: 2.717, w: 4.565, h: 4.565, line: { color: C.white, width: LINE.hairline } });

  heading(s, 'Introduction To Ramadhan', { x: 0.593, y: 0.792, w: 6.514, h: 1.717, fontSize: 48, color: C.white });
  body(s, L_LONG, { x: 0.593, y: 2.808, w: 4.383, h: 1.588, color: C.white });
  body(s, L_MED, { x: 0.593, y: 4.487, w: 4.383, h: 1.083, color: C.white });

  T(s, TEMPLATE_LABEL, { x: 9.885, y: 4.661, w: 2.704, h: 1.01, color: C.white, align: 'right' });
  body(s, L_SHORT, { x: 9.539, y: 5.735, w: 3.051, h: 1.084, fontFace: F.head, italic: true, color: C.white, align: 'right' });

  photo(s, { shape: 'star7', x: 5.127, y: 3.048, w: 3.961, h: 3.961, color: C.photoDark });
  photo(s, { shape: 'star7', x: 8.714, y: 0.492, w: 3.635, h: 3.635, color: C.photoMid });
}

/* ------------------------------------------------------------------ *
 * Slide 4 — "Preparing for Ramadhan", left photo band + three chips
 * ------------------------------------------------------------------ */

function slidePreparingLeftPhoto(pres) {
  const s = pres.addSlide();

  heading(s, 'Preparing for Ramadhan', { x: 4.629, y: 1.166, w: 6.837, h: 1.582, color: C.dark });
  body(s, L_FULL, { x: 4.629, y: 2.956, w: 7.788, h: 1.335 });

  [4.629, 6.894, 9.16].forEach((x) => {
    s.addShape('roundRect', { x, y: 4.752, w: 2.121, h: 1.06, fill: { color: C.dark }, rectRadius: rr(2.121, 1.06) });
    body(s, L_SENTENCE, { x: x + 0.144, y: 4.842, w: 1.977, h: 0.83, color: C.white });
  });

  photo(s, { x: 0, y: 0, w: 4.083, h: 7.5, color: C.photoLight });
}

/* ------------------------------------------------------------------ *
 * Slide 5 — "Preparing for Ramadhan", three bullet cards on a photo
 * ------------------------------------------------------------------ */

function slidePreparingCards(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);

  heading(s, 'Preparing for Ramadhan', { x: 2.078, y: 0.864, w: 9.176, h: 0.841, color: C.white, align: 'center' });
  body(s, L_FULL, { x: 1.023, y: 1.731, w: 11.287, h: 1.083, color: C.white, align: 'center' });

  // the middle card is lifted slightly above its neighbours
  [[1.644, 3.783], [5.238, 3.374], [8.831, 3.783]].forEach(([x, y]) => {
    s.addShape('roundRect', {
      x, y, w: 2.858, h: 3.261,
      fill: { color: C.dark }, rectRadius: rr(2.858, 3.261),
      shadow: { type: 'outer', blur: 7, offset: 3, angle: 45, color: C.white, opacity: 0.55 },
    });
    s.addShape('roundRect', {
      x: x + 0.198, y: y + 0.227, w: 2.463, h: 2.81,
      line: { color: C.white, width: LINE.card }, rectRadius: rr(2.463, 2.81),
    });
    [[0.669, 0.845], [1.657, 1.833]].forEach(([dyText, dyDot]) => {
      body(s, L_SENTENCE, { x: x + 0.72, y: y + dyText, w: 2.039, h: 0.83, color: C.white });
      dot(s, x + 0.431, y + dyDot, 0.236, C.white);
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 6 — "Daily Routine Ramadhan", arch photo + arch caption card
 * ------------------------------------------------------------------ */

function slideDailyRoutineArch(pres) {
  const s = pres.addSlide();
  photo(s, { shape: 'round2SameRect', x: 7.847, y: 0.885, w: 4.514, h: 5.729, color: C.photoLight });

  heading(s, 'Daily Routine Ramadhan', { x: 0.972, y: 0.937, w: 6.837, h: 1.582, color: C.dark });

  [[3.25, 2.953], [4.288, 3.991], [5.327, 5.03]].forEach(([dotY, textY]) => {
    dot(s, 0.972, dotY, 0.236, C.dark);
    body(s, L_SHORT, { x: 1.372, y: textY, w: 4.514, h: 0.83 });
  });

  // small arch caption card overlapping the photo
  s.addShape('round2SameRect', { x: 6.559, y: 4.525, w: 2.538, h: 2.538, fill: { color: C.dark }, angleRange: ARCH });
  s.addShape('round2SameRect', {
    x: 6.718, y: 4.683, w: 2.221, h: 2.221,
    line: { color: C.white, width: LINE.thin }, angleRange: ARCH,
  });
  T(s, TEMPLATE_LABEL, { x: 6.95, y: 5.179, w: 1.517, h: 0.656, fontSize: 11, color: C.white });
  body(s, L_SENTENCE, { x: 6.95, y: 5.835, w: 1.988, h: 0.832, fontFace: F.head, italic: true, color: C.white });
}

/* ------------------------------------------------------------------ *
 * Slide 7 — "Daily Routine Ramadhan", two outline cards + photo panel
 * ------------------------------------------------------------------ */

function slideDailyRoutineCards(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);
  photo(s, { x: 6.531, y: 0, w: 6.031, h: 3.75, color: C.photoMid });

  heading(s, 'Daily Routine Ramadhan', { x: 0.597, y: 1.084, w: 4.99, h: 1.582, color: C.white });
  body(s, L_SHORT, { x: 0.597, y: 2.666, w: 5.213, h: 0.58, fontFace: F.head, italic: true, color: C.white });

  [0.597, 3.347].forEach((x) => {
    s.addShape('roundRect', {
      x, y: 3.968, w: 2.463, h: 2.81,
      line: { color: C.white, width: LINE.card }, rectRadius: rr(2.463, 2.81),
    });
    [[4.502, 4.678], [5.489, 5.665]].forEach(([textY, dotY]) => {
      body(s, L_SENTENCE, { x: x + 0.546, y: textY, w: 2.039, h: 0.83, color: C.white });
      dot(s, x + 0.256, dotY, 0.236, C.white);
    });
  });

  body(s, L_LONG, { x: 6.53, y: 4.196, w: 6.031, h: 1.083, color: C.white });
  body(s, L_MED, { x: 6.53, y: 5.489, w: 6.031, h: 0.83, color: C.white });
}

/* ------------------------------------------------------------------ *
 * Slide 8 — "Daily Routine Ramadhan", wide photo band + dark card
 * ------------------------------------------------------------------ */

function slideDailyRoutineBand(pres) {
  const s = pres.addSlide();
  photo(s, { x: 0.738, y: 0, w: 11.857, h: 4.365, color: C.photoLight });

  s.addShape('roundRect', { x: 0.875, y: 3.5, w: 5.126, h: 3.532, fill: { color: C.dark }, rectRadius: rr(5.126, 3.532) });
  s.addShape('roundRect', {
    x: 1.229, y: 3.778, w: 4.416, h: 2.913,
    line: { color: C.white, width: LINE.card }, rectRadius: rr(4.416, 2.913),
  });
  body(s, L_MED, { x: 1.544, y: 4.074, w: 3.787, h: 1.335, color: C.white });
  body(s, L_SHORT, { x: 1.542, y: 5.459, w: 3.651, h: 0.83, color: C.white });

  heading(s, 'Daily Routine Ramadhan', { x: 6.891, y: 4.708, w: 4.99, h: 1.582 });
  body(s, L_SHORT, { x: 6.891, y: 6.29, w: 5.213, h: 0.58, fontFace: F.head, italic: true });
}

/* ------------------------------------------------------------------ *
 * Slide 9 — "Break Slides" section divider
 * ------------------------------------------------------------------ */

function slideBreak(pres) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape('round2SameRect', {
    x: 0.919, y: 1.011, w: 11.495, h: 5.478,
    line: { color: C.white, width: LINE.thin }, angleRange: adj(38590),
  });
  T(s, TEMPLATE_LABEL, { x: 3.035, y: 2.24, w: 7.264, h: 0.404, color: C.white, align: 'center' });
  heading(s, 'Break Slides', { x: 1.483, y: 2.804, w: 10.368, h: 1.717, fontSize: 96, color: C.white, align: 'center' });
  body(s, L_SHORT, { x: 3.465, y: 4.68, w: 6.403, h: 0.58, fontFace: F.head, italic: true, color: C.white, align: 'center' });
}

/* ------------------------------------------------------------------ *
 * Slide 10 — pull quote
 * ------------------------------------------------------------------ */

function slideQuote(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);

  T(s, L_SHORT, {
    x: 0.744, y: 2.126, w: 9.244, h: 3.248,
    fontFace: F.quote, fontSize: 32, italic: true, color: C.dark,
    align: 'center', lineSpacingMultiple: 1.5,
  });
  T(s, TEMPLATE_LABEL, { x: 10.472, y: 4.826, w: 2.117, h: 0.909, fontSize: 16, align: 'right' });
  body(s, L_SHORT, { x: 9.539, y: 5.735, w: 3.051, h: 1.084, fontFace: F.head, italic: true, align: 'right' });
}

/* ------------------------------------------------------------------ *
 * Slide 11 — "Event Schedule Ramadhan" rule-drawn table
 * ------------------------------------------------------------------ */

/* Hand-drawn table: connector lines + free text boxes (as in the original). */
const SCHEDULE_H_RULES = [[2.128, 3.0, 8.656], [3.489, 3.812, 7.295], [2.128, 4.489, 8.656],
  [2.128, 5.291, 8.656], [3.489, 6.136, 7.295]];
const SCHEDULE_V_RULES = [3.489, 5.448, 9.26];
const SCHEDULE_HEAD = [['Date', 2.128, 2.541, 1.337], ['Hour', 3.8, 2.541, 1.337],
  ['Activities ', 6.266, 2.555, 2.175], ['y/n ', 9.581, 2.555, 0.883]];
/* [date, dateX, hour, hourX, rowY, activityX, activityY, iconX, iconY, icon, iconColor] */
const SCHEDULE_ROWS = [
  { date: '14/3', dateX: 2.128, hour: '00.00', hourX: 3.776, y: 3.239, actX: 6.191, actY: 3.16, iconX: 9.801, iconY: 3.205, mark: ICON_CHECK, markColor: C.white },
  { hour: '03.00', hourX: 3.798, y: 3.931, actX: 6.191, actY: 3.822, iconX: 9.803, iconY: 3.895, mark: ICON_CROSS, markColor: C.white },
  { date: '15/3', dateX: 2.152, hour: '06.00', hourX: 3.8, y: 4.618, actX: 6.191, actY: 4.618, iconX: 9.799, iconY: 4.684, mark: ICON_CHECK, markColor: C.gray },
  { date: '16/3', dateX: 2.128, hour: '18.00', hourX: 3.776, y: 5.558, actX: 6.167, actY: 5.435, iconX: 9.797, iconY: 5.491, mark: ICON_CROSS, markColor: C.gray },
  { hour: '20.00', hourX: 3.776, y: 6.267, actX: 6.167, actY: 6.267, iconX: 9.797, iconY: 6.368, mark: ICON_CROSS, markColor: C.gray },
];

function slideSchedule(pres) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  T(s, TEMPLATE_LABEL, { x: 3.035, y: 0.687, w: 7.264, h: 0.404, color: C.white, align: 'center' });
  heading(s, 'Event Schedule Ramadhan', { x: 2.423, y: 1.012, w: 8.487, h: 0.841, color: C.white, align: 'center' });

  SCHEDULE_H_RULES.forEach(([x, y, w]) => s.addShape('line', { x, y, w, h: 0, line: { color: C.white, width: LINE.hairline } }));
  SCHEDULE_V_RULES.forEach((x) => s.addShape('line', { x, y: 2.541, w: 0, h: 4.444, line: { color: C.white, width: LINE.hairline } }));

  const cell = (text, x, y, w) => T(s, text, { x, y, w, h: 0.438, fontSize: 20, bold: true, color: C.white, align: 'center' });
  SCHEDULE_HEAD.forEach(([text, x, y, w]) => cell(text, x, y, w));

  SCHEDULE_ROWS.forEach((r) => {
    if (r.date) cell(r.date, r.dateX, r.y, 1.337);
    cell(r.hour, r.hourX, r.y, 1.337);
    body(s, L_SENTENCE, { x: r.actX, y: r.actY, w: 2.657, h: 0.578, color: C.white });
    dot(s, r.actX - 0.331, r.actY + 0.176, 0.185, C.white);
    icon(s, r.mark, r.iconX, r.iconY, 0.439, r.markColor);
  });
}

/* ------------------------------------------------------------------ *
 * Slide 12 — "Event Schedule Ramadhan" March calendar inside an arch
 * ------------------------------------------------------------------ */

const CAL_WEEKDAYS = [['Su', 6.993, 2.363, 0.539], ['Mo', 7.712, 2.363, 0.569], ['Tu', 8.433, 2.36, 0.539],
  ['We', 9.149, 2.363, 0.576], ['Th', 9.877, 2.363, 0.539], ['Fr', 10.632, 2.369, 0.539], ['Sa', 11.387, 2.36, 0.539]];
/* [label, x, y] — day cells keep the original per-cell nudges. */
const CAL_DAYS = [
  ['1 ', 11.486, 2.893],
  ['2 ', 7.071, 3.467], ['3 ', 7.759, 3.464], ['4 ', 8.515, 3.464], ['5 ', 9.27, 3.464],
  ['6 ', 10.051, 3.466], ['7 ', 10.78, 3.466], ['8 ', 11.486, 3.466],
  ['9 ', 7.064, 4.065], ['10 ', 7.569, 4.065], ['11 ', 8.325, 4.036], ['12 ', 9.088, 4.036],
  ['13 ', 9.807, 4.058], ['14 ', 10.535, 4.034], ['15 ', 11.237, 4.034],
  ['16 ', 6.888, 4.655], ['17 ', 7.596, 4.655], ['18 ', 8.353, 4.655], ['19 ', 9.115, 4.65],
  ['20 ', 9.827, 4.65], ['21 ', 10.555, 4.625], ['22 ', 11.257, 4.625],
  ['23 ', 6.888, 5.271], ['24 ', 7.596, 5.271], ['25 ', 8.353, 5.271], ['26 ', 9.115, 5.266],
  ['27 ', 9.88, 5.266], ['28 ', 10.608, 5.241], ['29 ', 11.31, 5.241],
  ['30 ', 6.883, 5.813], ['31 ', 7.591, 5.813],
];

function slideCalendar(pres) {
  const s = pres.addSlide();

  s.addShape('round2SameRect', { x: 6.051, y: 0.369, w: 6.762, h: 6.762, fill: { color: C.dark }, angleRange: ARCH });

  T(s, 'March  ', { x: 7.928, y: 1.409, w: 3.07, h: 0.505, fontFace: F.body, fontSize: 24, bold: true, color: C.white, align: 'center' });
  CAL_WEEKDAYS.forEach(([label, x, y, w]) =>
    T(s, label, { x, y, w, h: 0.404, fontFace: F.body, bold: true, color: C.white, align: 'center' }));
  CAL_DAYS.forEach(([label, x, y]) =>
    T(s, label, {
      x, y, w: label.trim().length > 1 ? 0.688 : 0.34, h: 0.572,
      fontFace: F.body, fontSize: 28, color: C.white, align: 'center',
    }));

  s.addShape('round2SameRect', {
    x: 6.314, y: 0.632, w: 6.237, h: 6.237,
    line: { color: C.white, width: LINE.thin }, angleRange: ARCH,
  });

  heading(s, 'Event Schedule Ramadhan', { x: 0.679, y: 2.369, w: 5.058, h: 1.582 });
  body(s, L_LONG, { x: 0.679, y: 4.172, w: 4.926, h: 1.335 });
}

/* ------------------------------------------------------------------ *
 * Slide 13 — "Event Ramadhan", bullets + two photo panels
 * ------------------------------------------------------------------ */

function slideEvent(pres) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  heading(s, 'Event Ramadhan', { x: 0.838, y: 1.444, w: 5.058, h: 1.582, color: C.white });
  [[3.647, 3.35], [4.623, 4.326], [5.599, 5.302]].forEach(([dotY, textY]) => {
    dot(s, 0.838, dotY, 0.236, C.white);
    body(s, L_SHORT, { x: 1.237, y: textY, w: 3.479, h: 0.83, color: C.white });
  });
  body(s, L_SHORT, { x: 9.143, y: 5.696, w: 3.055, h: 1.083, color: C.white, align: 'right' });

  photo(s, { x: 5.722, y: 2.206, w: 3.055, h: 4.698, color: C.photoLight });
  photo(s, { x: 9.143, y: 0, w: 3.055, h: 5.508, color: C.photoMid });
}

/* ------------------------------------------------------------------ *
 * Slide 14 — "Ramadhan Gallery", four star photos
 * ------------------------------------------------------------------ */

/* [outlineX, outlineY, outlineSize, photoX, photoY, photoSize, photoColor] */
const GALLERY_STARS = [
  { ox: 3.99, oy: 0.231, ow: 3.961, oh: 4.014, x: 4.159, y: 0.452, size: 3.635, color: C.photoLight },
  { ox: 0.968, oy: 3.205, ow: 4.248, oh: 4.248, x: 1.111, y: 3.349, size: 3.961, color: C.photoMid },
  { ox: 0.457, oy: 0.303, ow: 3.118, oh: 3.118, x: 0.579, y: 0.452, size: 2.873, color: C.photoLight },
  { ox: 5.487, oy: 4.158, ow: 2.89, oh: 2.89, x: 5.651, y: 4.309, size: 2.611, color: C.photoDark },
];

function slideGallery(pres) {
  const s = pres.addSlide();

  GALLERY_STARS.forEach((g) =>
    s.addShape('star7', { x: g.ox, y: g.oy, w: g.ow, h: g.oh, line: { color: C.dark, width: LINE.hairline } }));

  heading(s, 'Ramadhan Gallery ', { x: 8.786, y: 2.263, w: 3.58, h: 1.582, align: 'right' });
  body(s, L_MED, { x: 8.786, y: 3.894, w: 3.58, h: 1.335, align: 'right' });

  GALLERY_STARS.forEach((g) => photo(s, { shape: 'star7', x: g.x, y: g.y, w: g.size, h: g.size, color: g.color }));
}

/* ------------------------------------------------------------------ *
 * Slide 15 — "Thank You"
 * ------------------------------------------------------------------ */

function slideThankYou(pres) {
  const s = pres.addSlide();
  photoBackdrop(s, C.photoLight);

  T(s, TEMPLATE_LABEL, { x: 3.035, y: 2.555, w: 7.264, h: 0.404, color: C.white, align: 'center' });
  heading(s, 'Thank You', { x: 2.188, y: 2.959, w: 8.958, h: 1.717, fontSize: 96, color: C.white, align: 'center' });
  body(s, L_SHORT, { x: 2.827, y: 4.541, w: 7.68, h: 0.58, fontFace: F.head, italic: true, color: C.white, align: 'center' });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_16x9';
  pres.title = 'Ramadhan Presentation Template';

  [
    slideCover, slideGetReady, slideIntroduction, slidePreparingLeftPhoto, slidePreparingCards,
    slideDailyRoutineArch, slideDailyRoutineCards, slideDailyRoutineBand, slideBreak, slideQuote,
    slideSchedule, slideCalendar, slideEvent, slideGallery, slideThankYou,
  ].forEach((buildSlide) => buildSlide(pres));

  return pres;
}

build()
  .writeFile({ fileName: path.join(__dirname, '195e00d4-f046-41d9-8b03-466c7e7e16d7_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
