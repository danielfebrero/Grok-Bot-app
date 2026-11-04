/**
 * PAYCASH — Finance Presentation Template
 * Standalone pptxgenjs recreation of the 20-slide reference deck (20" x 11.25").
 *
 * Raster images in the source deck are replaced by programmatic placeholders
 * (`imageArea`) and the small vector icons are rebuilt from native shapes.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const BG = '0C0C1E';        // page background (source uses a 0D0D27->0B0B15 gradient)
const CARD = '0D0D27';      // dark card fill
const GLASS = '181829';     // BG seen through the cover's 5%-white glass panel
const GREEN = '01DA8C';     // primary accent
const GREEN_DEEP = '01BB79';// filled panels
const GREEN_ITAL = '0DCF69';// "Paycash" wordmark on the cover slides
const WHITE = 'FFFFFF';
const GREY = 'D9D9D9';      // white @ lumMod 85%
const GREY_LT = 'F2F2F2';   // white @ lumMod 95%

const POPPINS_SB = 'Poppins SemiBold';
const POPPINS_MD = 'Poppins Medium';
const INTER = 'Inter';
const INTER_MD = 'Inter Medium';
const PLAYFAIR = 'Playfair Display';

const BODY = { fontFace: INTER, fontSize: 20, color: GREY, lineSpacingMultiple: 1.5 };
const KICKER = { fontFace: INTER_MD, fontSize: 16, color: WHITE, charSpacing: 2 };

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** Text box: the source deck anchors every text frame to the top. */
function text(slide, body, opts) {
  slide.addText(body, Object.assign({ valign: 'top' }, opts));
}

function rect(slide, o) {
  slide.addShape('rect', o);
}

function roundRect(slide, o) {
  slide.addShape('roundRect', o);
}

/**
 * Placeholder standing in for a picture frame in the reference deck.
 * The source frames are empty, so the placeholder is tinted to whatever
 * surface it sits on (`tone`) and keeps the original position/size/radius.
 */
function imageArea(slide, x, y, w, h, radius, tone) {
  const c = tone || BG;
  const o = { x, y, w, h, fill: { color: c }, line: { color: c, width: 0.75 } };
  if (radius) { o.rectRadius = radius; slide.addShape('roundRect', o); }
  else slide.addShape('rect', o);
}

/** Two-tone heading: accent word + rest of the sentence. */
function heading(slide, o) {
  text(slide, [
    { text: o.green, options: { color: GREEN } },
    { text: o.rest, options: { color: WHITE } },
  ], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: POPPINS_MD, fontSize: o.size || 54, align: o.align || 'left',
  });
}

/** Thin right-pointing arrow used beside every section kicker. */
function arrowIcon(slide, x, y, w, h, color, rotate) {
  const P = [ // outline traced from the source freeform, normalised to 0..1
    [0.000, 0.500], [0.065, 0.589], [0.778, 0.589], [0.585, 0.851], [0.588, 0.976],
    [0.677, 0.976], [0.981, 0.563], [0.981, 0.437], [0.677, 0.024], [0.588, 0.028],
    [0.588, 0.149], [0.778, 0.411], [0.065, 0.411],
  ];
  slide.addShape('custGeom', {
    x, y, w, h, fill: { color: color || GREEN }, line: { type: 'none' }, rotate: rotate || 0,
    points: P.map(([px, py]) => ({ x: px * w, y: py * h })).concat([{ close: true }]),
  });
}

/** Section kicker: "-> ABOUT US". `x`/`y` locate the arrow. */
function kicker(slide, x, y, label) {
  arrowIcon(slide, x, y + 0.102, 0.226, 0.166);
  text(slide, label, Object.assign({ x: x + 0.455, y, w: 4.083, h: 0.37 }, KICKER));
}

/** Three stacked chevrons (the ">>>" motif inside the INVEST NOW pill). */
function chevrons(slide, x, y, size) {
  const s = size || 0.363;
  const P = [[0.376, 0.794], [0.670, 0.499], [0.376, 0.205], [0.331, 0.249],
             [0.581, 0.499], [0.331, 0.750]];
  [0, 40, 65].forEach((transparency, i) => {
    slide.addShape('custGeom', {
      x: x + (2 - i) * 0.1635, y, w: s, h: s,
      fill: { color: WHITE, transparency }, line: { type: 'none' },
      points: P.map(([px, py]) => ({ x: px * s, y: py * s })).concat([{ close: true }]),
    });
  });
}

/** Pill-shaped call to action. `solid` = green pill, otherwise outlined. */
function investButton(slide, x, y, solid) {
  roundRect(slide, {
    x, y, w: 3.57, h: 0.924, rectRadius: 0.462,
    fill: solid ? { color: GREEN } : { type: 'none' },
    line: solid ? { type: 'none' } : { color: GREEN, width: 1 },
  });
  roundRect(slide, {
    x: x + 2.289, y: y + 0.111, w: 1.118, h: 0.703, rectRadius: 0.3515,
    fill: solid ? { color: WHITE, transparency: 60 } : { color: GREEN }, line: { type: 'none' },
  });
  text(slide, 'INVEST NOW', Object.assign({ x: x + 0.357, y: y + 0.277, w: 1.932, h: 0.37 }, KICKER, { charSpacing: 1 }));
  chevrons(slide, x + 2.503, y + 0.281);
}

/* ------------------------------------------------------------------ *
 * Icon set (rebuilt from the deck's vector icons)
 * ------------------------------------------------------------------ */

/** Icon helpers take a unit box (x, y, size) and place primitives inside it. */
function iconRect(slide, x, y, s, fx, fy, fw, fh, c) {
  rect(slide, { x: x + fx * s, y: y + fy * s, w: fw * s, h: fh * s, fill: { color: c }, line: { type: 'none' } });
}

function bankIcon(slide, x, y, s, color) {
  const c = color || GREEN;
  slide.addShape('triangle', { x: x + 0.10 * s, y: y + 0.12 * s, w: 0.80 * s, h: 0.20 * s, fill: { color: c }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 0.455 * s, y: y + 0.21 * s, w: 0.09 * s, h: 0.09 * s, fill: { color: BG }, line: { type: 'none' } });
  iconRect(slide, x, y, s, 0.10, 0.32, 0.80, 0.05, c);
  iconRect(slide, x, y, s, 0.14, 0.37, 0.72, 0.03, c);
  [0.18, 0.345, 0.51, 0.675].forEach(f => iconRect(slide, x, y, s, f, 0.40, 0.075, 0.36, c));
  iconRect(slide, x, y, s, 0.14, 0.76, 0.72, 0.04, c);
  iconRect(slide, x, y, s, 0.08, 0.80, 0.84, 0.07, c);
}

function moneyIcon(slide, x, y, s, color) {
  const c = color || GREEN;
  // Two banknotes fanned out behind the front note
  slide.addShape('parallelogram', { x: x + 0.20 * s, y: y + 0.20 * s, w: 0.66 * s, h: 0.10 * s, rotate: 352, fill: { color: c }, line: { type: 'none' } });
  slide.addShape('parallelogram', { x: x + 0.14 * s, y: y + 0.29 * s, w: 0.72 * s, h: 0.09 * s, rotate: 356, fill: { color: c }, line: { type: 'none' } });
  rect(slide, { x: x + 0.06 * s, y: y + 0.42 * s, w: 0.88 * s, h: 0.38 * s, fill: { color: c }, line: { type: 'none' } });
  slide.addShape('octagon', { x: x + 0.11 * s, y: y + 0.47 * s, w: 0.78 * s, h: 0.28 * s, rectRadius: 0.2, fill: { color: BG }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 0.42 * s, y: y + 0.52 * s, w: 0.16 * s, h: 0.18 * s, fill: { color: c }, line: { type: 'none' } });
  [0.19, 0.75].forEach(f => {
    slide.addShape('ellipse', { x: x + f * s, y: y + 0.58 * s, w: 0.06 * s, h: 0.06 * s, fill: { color: c }, line: { type: 'none' } });
  });
}

function cityIcon(slide, x, y, s, color) {
  const c = color || GREEN;
  // Three towers, each with a grid of lit windows
  [[0.06, 0.34, 0.28], [0.36, 0.46, 0.28], [0.66, 0.14, 0.26]].forEach(([fx, fy, fw]) => {
    iconRect(slide, x, y, s, fx, fy, fw, 0.90 - fy, c);
    for (let r = 0; r < 4; r++) {
      for (let col = 0; col < 2; col++) {
        const wy = fy + 0.06 + r * 0.14;
        if (wy < 0.82) iconRect(slide, x, y, s, fx + 0.05 + col * 0.11, wy, 0.06, 0.06, BG);
      }
    }
  });
}

function phoneIcon(slide, x, y, s, color) {
  const c = color || GREEN;
  // Handset: an arc with a rounded earpiece at each end
  slide.addShape('blockArc', { x: x + 0.06 * s, y: y + 0.20 * s, w: 0.88 * s, h: 0.52 * s,
    angleRange: [200, 340], arcThicknessRatio: 0.5, fill: { color: c }, line: { type: 'none' } });
  [0.04, 0.74].forEach(f => {
    slide.addShape('ellipse', { x: x + f * s, y: y + 0.38 * s, w: 0.22 * s, h: 0.18 * s, fill: { color: c }, line: { type: 'none' } });
  });
  slide.addShape('trapezoid', { x: x + 0.16 * s, y: y + 0.48 * s, w: 0.68 * s, h: 0.34 * s, fill: { color: c }, line: { type: 'none' } });
  for (let r = 0; r < 3; r++) {
    for (let col = 0; col < 3; col++) {
      iconRect(slide, x, y, s, 0.34 + col * 0.11, 0.55 + r * 0.08, 0.055, 0.045, BG);
    }
  }
}

/** Downward-pointing triangle (pptxgenjs `triangle` only points up). */
function wedgeDown(slide, x, y, w, h, color) {
  slide.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: [{ x: 0, y: 0 }, { x: w, y: 0 }, { x: w / 2, y: h }, { close: true }],
  });
}

function emailIcon(slide, x, y, s, color) {
  const c = color || GREEN;
  slide.addShape('triangle', { x: x + 0.06 * s, y: y + 0.28 * s, w: 0.88 * s, h: 0.26 * s, fill: { color: c }, line: { type: 'none' } });
  rect(slide, { x: x + 0.06 * s, y: y + 0.50 * s, w: 0.88 * s, h: 0.32 * s, fill: { color: c }, line: { type: 'none' } });
  wedgeDown(slide, x + 0.10 * s, y + 0.50 * s, 0.80 * s, 0.26 * s, BG);
  // The "@" card tucked inside the envelope
  rect(slide, { x: x + 0.26 * s, y: y + 0.14 * s, w: 0.48 * s, h: 0.44 * s, fill: { color: c }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 0.31 * s, y: y + 0.19 * s, w: 0.38 * s, h: 0.34 * s, fill: { color: BG }, line: { type: 'none' } });
  text(slide, '@', { x: x + 0.26 * s, y: y + 0.16 * s, w: 0.48 * s, h: 0.40 * s, align: 'center', valign: 'middle', margin: 0, fontFace: INTER, fontSize: Math.round(s * 24), color: c });
}

function quoteMark(slide, x, y, s, color) {
  text(slide, '\u201D', {
    x, y: y - 0.06 * s, w: s, h: s * 1.3, align: 'center',
    fontFace: POPPINS_SB, fontSize: Math.round(s * 96), color: color || WHITE,
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/** Slides 1 and 20 share the cover layout: white word + green word, 150pt. */
function coverSlide(slide, lead) {
  text(slide, [
    { text: lead.white, options: { color: WHITE } },
    { text: lead.green, options: { color: GREEN } },
  ], { x: lead.x, y: 1.736, w: lead.w, h: 2.625, fontFace: POPPINS_SB, fontSize: 150, align: 'center' });

  text(slide, 'Finance Presentation Template',
    { x: 4.725, y: 4.027, w: 10.551, h: 0.64, fontFace: INTER, fontSize: 32, color: WHITE, align: 'center' });

  roundRect(slide, {
    x: 1.523, y: 5.625, w: 16.953, h: 4.375, rectRadius: 0.833,
    fill: { color: WHITE, transparency: 95 }, line: { color: WHITE, transparency: 70, width: 2.5 },
  });
  imageArea(slide, 1.756, 5.785, 11.685, 4.055, 0.772, GLASS);

  text(slide, [
    { text: 'Paycash', options: { color: GREEN_ITAL, italic: true, fontFace: PLAYFAIR } },
    { text: ' Financial for Your Future Needs', options: { color: WHITE, fontFace: INTER } },
  ], { x: 14.03, y: 6.719, w: 3.97, h: 0.909, fontSize: 24 });

  investButton(slide, 14.03, 7.982, true);
}

function slide01(slide) { coverSlide(slide, { white: 'PAY', green: 'CASH', x: 3.759, w: 12.482 }); }
function slide20(slide) { coverSlide(slide, { white: 'THANK ', green: 'YOU', x: 2.14, w: 15.719 }); }

function slide02(slide) {
  imageArea(slide, 10.23, 0, 5.542, 11.25);
  imageArea(slide, 15.875, 0, 4.125, 6.667);
  imageArea(slide, 15.875, 6.776, 4.125, 4.474);
  kicker(slide, 1.455, 2.607, 'ABOUT US');
  heading(slide, { x: 1.362, y: 3.112, w: 7.574, h: 1.919, green: 'Paycash', rest: ' financial management ' });
  text(slide, 'ANYTHING WRITE HERE', { x: 1.362, y: 5.625, w: 6.492, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you would bring to the table if you were hired for this position bells and whistles. We need to socialize the comms with thesa wider finance stakeholder community.',
    Object.assign({ x: 1.362, y: 6.197, w: 7.388, h: 2.566 }, BODY));
}

function slide03(slide) {
  imageArea(slide, 10.576, 0, 9.424, 11.25);
  kicker(slide, 1.559, 2.247, 'ABOUT US');
  heading(slide, { x: 1.466, y: 2.752, w: 5.679, h: 2.827, green: 'Building', rest: ' A solid financial foundation' });
  text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you would bring to the table if you were hired for this position bells and whistles. We need to socialize the comms with the are wider stakeholder community we need to this forcing function or get six alpha pups in here.',
    Object.assign({ x: 1.466, y: 5.931, w: 7.388, h: 3.071 }, BODY));

  rect(slide, { x: 11.306, y: 4.866, w: 5.264, h: 5.404, fill: { color: GREEN_DEEP }, line: { type: 'none' } });
  bankIcon(slide, 11.948, 5.799, 0.709, WHITE);
  text(slide, 'ANYTHING WRITE HERE', { x: 11.948, y: 6.823, w: 4.234, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, 'The last person we talked to said this would ready action item and what do you feel us you would bring table.',
    Object.assign({ x: 11.948, y: 7.277, w: 4.122, h: 2.061 }, BODY, { color: WHITE }));
}

function slide04(slide) {
  imageArea(slide, 0, 1.628, 9.274, 7.995);
  rect(slide, { x: 3.095, y: 1.157, w: 4.196, h: 8.936, fill: { type: 'none' }, line: { color: GREEN, width: 2.5 } });
  kicker(slide, 11.038, 1.827, 'ABOUT US');
  heading(slide, { x: 10.946, y: 2.332, w: 7.574, h: 1.919, green: 'Taking', rest: ' control of your finances' });
  text(slide, 'ANYTHING WRITE HERE', { x: 11.038, y: 4.759, w: 4.234, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you would bring to the table if you were hired for this position.',
    Object.assign({ x: 11.038, y: 5.331, w: 7.388, h: 1.557 }, BODY));
  text(slide, [
    { text: 'Paycash', options: { color: GREEN, italic: true, fontFace: PLAYFAIR } },
    { text: ' Financial for Your Future Needs', options: { color: WHITE, fontFace: INTER } },
  ], { x: 11.038, y: 7.322, w: 4.424, h: 0.909, fontSize: 24 });
  text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you bring.',
    Object.assign({ x: 11.038, y: 8.371, w: 7.388, h: 1.052 }, BODY));
}

function slide05(slide) {
  imageArea(slide, 0, 5.625, 8.812, 4.251);
  kicker(slide, 1.559, 2.046, 'ABOUT US');
  heading(slide, { x: 1.466, y: 2.551, w: 7.12, h: 1.919, green: 'Creating', rest: ' a plan for success' });

  // Framed "Paycash" call-out
  rect(slide, { x: 2.211, y: 6.791, w: 6.003, h: 1.919, fill: { color: GREEN_DEEP }, line: { type: 'none' } });
  rect(slide, { x: 2.004, y: 6.57, w: 6.417, h: 2.361, fill: { type: 'none' }, line: { color: WHITE, width: 2.5 } });
  text(slide, [
    { text: 'Paycash', options: { italic: true, fontFace: PLAYFAIR } },
    { text: ' Financial for Your Future Needs', options: { fontFace: INTER } },
  ], { x: 3.0, y: 7.296, w: 4.424, h: 0.909, fontSize: 24, color: WHITE });

  // Two icon + rule + copy blocks
  [{ y: 2.42, icon: bankIcon }, { y: 6.185, icon: moneyIcon }].forEach(block => {
    block.icon(slide, 10.0, block.y, 0.709, GREEN);
    text(slide, 'ANYTHING WRITE HERE', { x: 10.0, y: block.y + 0.894, w: 3.626, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    slide.addShape('line', { x: 13.942, y: block.y + 1.08, w: 4.287, h: 0, line: { color: GREEN, width: 0.75 } });
    text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you would bring to the table if you were hired for this position bells and whistles. ',
      Object.assign({ x: 10.0, y: block.y + 1.433, w: 8.515, h: 1.557 }, BODY));
  });
}

function slide06(slide) {
  imageArea(slide, 9.625, 4.961, 5.125, 6.289);
  kicker(slide, 1.476, 2.176, 'ABOUT US');
  heading(slide, { x: 1.383, y: 2.681, w: 7.574, h: 1.919, green: 'Building', rest: ' a better financial future' });
  text(slide, 'The last person we talked to said this would beasi ready action item, and what do you feel you would bring to the table if you were hired for this position bells and whistles. We need to socialize the comat with the are stakeholder community need.',
    Object.assign({ x: 1.383, y: 4.961, w: 6.992, h: 2.566 }, BODY));
  investButton(slide, 1.476, 8.15, false);

  text(slide, 'ANYTHING WRITE HERE', { x: 9.625, y: 2.176, w: 4.234, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, 'The last person we talked to said this would be ready action item and what do you feel you would bring to the table if you are were  hired for this position bells and whistles socialize.',
    Object.assign({ x: 9.625, y: 2.748, w: 8.992, h: 1.557 }, BODY));

  // Framed green card
  rect(slide, { x: 13.11, y: 5.909, w: 5.264, h: 4.396, fill: { color: GREEN_DEEP }, line: { type: 'none' } });
  rect(slide, { x: 12.867, y: 5.719, w: 5.75, h: 4.773, fill: { type: 'none' }, line: { color: WHITE, width: 2.5 } });
  bankIcon(slide, 13.753, 6.589, 0.709, WHITE);
  text(slide, 'ANYTHING WRITE HERE', { x: 13.753, y: 7.614, w: 4.234, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, 'PLACEHOLDER',
    Object.assign({ x: 13.753, y: 8.067, w: 4.122, h: 1.557 }, BODY, { color: WHITE }));
}

function slide07(slide) {
  imageArea(slide, 1.254, 1.875, 5.142, 9.375);
  rect(slide, { x: 0.936, y: 1.585, w: 4.196, h: 8.936, fill: { type: 'none' }, line: { color: GREEN, width: 2.5 } });
  kicker(slide, 7.56, 2.089, 'ABOUT US');
  heading(slide, { x: 7.467, y: 2.594, w: 6.95, h: 1.919, green: 'Road', rest: ' to financial freedom' });
  text(slide, 'ANYTHING WRITE HERE', { x: 7.56, y: 5.013, w: 4.234, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
  text(slide, [
    { text: 'The last person we talked to said this would be ready action item, and what do you feel you huni would bring to the table if you were hired for this position bells and whistles to socialize.', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'Start advertising on social media usabiltiy game plan. Deliverables please advise here.' },
  ], Object.assign({ x: 7.56, y: 5.585, w: 6.691, h: 3.576 }, BODY));

  [2.374, 5.823].forEach(y0 => {
    text(slide, '$10K', { x: 14.98, y: y0 + 0.223, w: 3.461, h: 0.64, fontFace: POPPINS_MD, fontSize: 32, color: GREEN });
    text(slide, 'ANYTHING WRITE HERE', { x: 14.98, y: y0 + 0.964, w: 3.626, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    text(slide, 'The last person we talked to said this would be ready action item and what.',
      Object.assign({ x: 14.98, y: y0 + 1.503, w: 3.766, h: 1.557 }, BODY));
  });
}

function slide08(slide) {
  imageArea(slide, 10.356, 0.792, 8.178, 9.666);
  rect(slide, { x: 10.229, y: 1.631, w: 4.196, h: 7.865, fill: { type: 'none' }, line: { color: GREEN, width: 2.5 } });
  kicker(slide, 1.559, 1.817, 'ABOUT US');
  heading(slide, { x: 1.466, y: 2.322, w: 7.574, h: 1.919, green: 'Financial', rest: ' planning for beginners' });

  [4.752, 7.338].forEach(y0 => {
    text(slide, 'ANYTHING WRITE HERE', {
      x: 1.466, y: y0, w: 6.732, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE,
      bullet: { characterCode: '2713', indent: 27 },
    });
    text(slide, 'The last person we talked to said this would bearu ready action item, and what do you feel you would bring to the table you were hired position.',
      Object.assign({ x: 1.466, y: y0 + 0.539, w: 6.95, h: 1.557 }, BODY));
  });
}

function slide09(slide) {
  imageArea(slide, 1.544, 1.604, 16.913, 4.021, 0.625);
  rect(slide, { x: 2.965, y: 1.224, w: 3.472, h: 4.781, fill: { type: 'none' }, line: { color: GREEN, width: 2.5 } });
  kicker(slide, 1.637, 7.004, 'ABOUT US');
  heading(slide, { x: 1.544, y: 7.509, w: 7.574, h: 1.919, green: 'Protecting', rest: ' your assets finance' });

  rect(slide, { x: 9.146, y: 7.004, w: 3.549, h: 2.782, fill: { color: GREEN_DEEP }, line: { type: 'none' } });
  rect(slide, { x: 9.0, y: 6.844, w: 3.84, h: 3.104, fill: { type: 'none' }, line: { color: WHITE, width: 1.5 } });
  text(slide, [
    { text: 'Paycash', options: { italic: true, fontFace: PLAYFAIR } },
    { text: ' Financial for Your Future Needs', options: { fontFace: INTER } },
  ], { x: 9.665, y: 7.403, w: 2.509, h: 1.986, fontSize: 28, color: WHITE });

  text(slide, '85%', { x: 13.631, y: 6.975, w: 4.234, h: 0.707, fontFace: POPPINS_MD, fontSize: 36, color: GREEN });
  text(slide, 'The last person we talked to said this would be ready action item and what do you feel you hunian would bring to the table.',
    Object.assign({ x: 13.631, y: 7.754, w: 4.589, h: 2.061 }, BODY));
}

function slide10(slide) {
  [7.751, 11.834, 15.917].forEach(x => imageArea(slide, x, 0, 4.083, 11.25));
  kicker(slide, 1.33, 2.361, 'OUR TEAM');
  heading(slide, { x: 1.237, y: 2.866, w: 5.888, h: 1.717, green: 'Professional', rest: ' Our Team', size: 48 });
  text(slide, 'The last person we talked to said would beasi ready action item, and what do you feel you would bring to the table if you were hired for this position bells and wita whistles. We need to socialize',
    Object.assign({ x: 1.237, y: 4.85, w: 5.795, h: 2.566 }, BODY));
  investButton(slide, 1.33, 7.965, false);

  // Three staggered name cards over the portraits
  [{ x: 7.751, y: 6.944 }, { x: 11.835, y: 8.107 }, { x: 15.918, y: 9.337 }].forEach(card => {
    rect(slide, { x: card.x, y: card.y, w: 4.082, h: 1.919, fill: { color: GREEN_DEEP, transparency: 30 }, line: { type: 'none' } });
    text(slide, 'YOUR NAME HERE', { x: card.x - 0.001, y: card.y + 0.481, w: 4.083, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE, align: 'center' });
    text(slide, 'Job Position', { x: card.x - 0.001, y: card.y + 0.892, w: 4.083, h: 0.547, fontFace: INTER, fontSize: 20, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

function slide11(slide) {
  text(slide, 'OUR TEAM', Object.assign({ x: 7.959, y: 1.259, w: 4.083, h: 0.37, align: 'center' }, KICKER));
  heading(slide, { x: 4.733, y: 1.764, w: 10.534, h: 1.01, green: 'Professional', rest: ' Our Team', align: 'center' });

  [1.152, 5.618, 10.084, 14.55].forEach(x => {
    rect(slide, { x, y: 3.737, w: 4.298, h: 6.118, fill: { type: 'none' }, line: { color: WHITE, width: 1.5 } });
    imageArea(slide, x + 0.168, 3.908, 3.962, 4.177);
    text(slide, 'YOUR NAME HERE', { x, y: 8.398, w: 4.298, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE, align: 'center' });
    text(slide, 'Job Position', { x, y: 8.809, w: 4.298, h: 0.547, fontFace: INTER, fontSize: 20, color: GREEN, align: 'center', lineSpacingMultiple: 1.5 });
    arrowIcon(slide, x + 3.919, 9.522, 0.226, 0.166, GREEN, 45);
  });
}

function slide12(slide) {
  imageArea(slide, 11.346, 0.75, 7.125, 10.5);
  // Decorative pills from the source layout
  [[16.372, 6.024, 2.325, 4.655, 45], [15.921, 2.415, 2.325, 4.655, 45], [11.677, 2.639, 1.956, 3.916, 323.9]]
    .forEach(([x, y, w, h, rot]) => {
      roundRect(slide, { x, y, w, h, rectRadius: Math.min(w, h) / 2, rotate: rot, fill: { color: GREEN }, line: { type: 'none' } });
    });

  heading(slide, { x: 1.529, y: 2.176, w: 9.222, h: 1.111, green: 'Mario', rest: ' Erlando', size: 60 });
  text(slide, 'JOB POSITION', Object.assign({ x: 1.529, y: 3.343, w: 4.943, h: 0.37 }, KICKER));
  text(slide, 'The last person we talked to said  would beasi ready action item and what do you feel you would.',
    Object.assign({ x: 1.529, y: 4.071, w: 7.804, h: 1.052 }, BODY));

  [{ y: 5.757, pct: '90%', bar: 6.007 }, { y: 6.991, pct: '80%', bar: 5.158 }, { y: 8.225, pct: '85%', bar: 5.698 }]
    .forEach(s => {
      text(slide, 'YOUR SKILL HERE', { x: 1.526, y: s.y, w: 5.235, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
      text(slide, s.pct, { x: 6.825, y: s.y + 0.002, w: 2.205, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE, align: 'right' });
      roundRect(slide, { x: 1.526, y: s.y + 0.646, w: 7.501, h: 0.204, rectRadius: 0.102, fill: { type: 'none' }, line: { color: GREEN, width: 1 } });
      roundRect(slide, { x: 1.604, y: s.y + 0.7, w: s.bar, h: 0.095, rectRadius: 0.0475, fill: { color: GREEN }, line: { type: 'none' } });
    });
}

function slide13(slide) {
  kicker(slide, 1.559, 3.012, 'OUR SERVICES');
  heading(slide, { x: 1.466, y: 3.517, w: 6.162, h: 1.919, green: 'Paycash', rest: ' Best Service' });
  text(slide, 'The last person we talked to said would beasi ready action item, and what do you feel you would bring to the table if you were hired for this position bells and wita whistles. We need to socialize',
    Object.assign({ x: 1.466, y: 5.672, w: 5.795, h: 2.566 }, BODY));

  [['01', 8.332, 1.372], ['02', 13.519, 1.372], ['03', 8.332, 5.72], ['04', 13.519, 5.72]].forEach(([num, x, y]) => {
    roundRect(slide, { x, y, w: 5.015, h: 4.154, rectRadius: 0.216, fill: { type: 'none' }, line: { color: GREEN, width: 1 } });
    text(slide, num, { x: x + 0.535, y: y + 0.721, w: 1.013, h: 0.572, fontFace: POPPINS_MD, fontSize: 28, color: GREEN });
    slide.addShape('line', { x: x + 1.43, y: y + 1.007, w: 2.863, h: 0, line: { color: WHITE, width: 0.75 } });
    text(slide, 'ANYTHING WRITE HERE', { x: x + 0.535, y: y + 1.562, w: 3.945, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    text(slide, 'PLACEHOLDER',
      Object.assign({ x: x + 0.535, y: y + 2.015, w: 3.945, h: 1.557 }, BODY));
  });
}

function slide14(slide) {
  imageArea(slide, 0, 6.354, 20, 4.896);
  kicker(slide, 1.497, 2.283, 'OUR SERVICES');
  heading(slide, { x: 1.404, y: 2.787, w: 6.162, h: 1.919, green: 'Paycash', rest: ' Best Service' });

  [7.892, 13.37].forEach(x => {
    roundRect(slide, { x, y: 2.152, w: 5.226, h: 5.488, rectRadius: 0.141, fill: { color: CARD }, line: { color: GREEN, width: 1 } });
    bankIcon(slide, x + 0.61, 3.148, 0.709, GREEN);
    text(slide, 'ANYTHING WRITE HERE', { x: x + 0.61, y: 4.043, w: 3.626, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    text(slide, 'The last person we talked to said this would be ready hisa action item, and what do you feel you would bring.',
      Object.assign({ x: x + 0.61, y: 4.582, w: 4.258, h: 2.061 }, BODY));
  });
}

function slide15(slide) {
  text(slide, 'OUR SERVICES', Object.assign({ x: 8.21, y: 1.012, w: 3.58, h: 0.37, align: 'center' }, KICKER));
  heading(slide, { x: 4.127, y: 1.517, w: 11.746, h: 1.01, green: 'Paycash', rest: ' Best Service', align: 'center' });

  [1.538, 7.387, 13.236].forEach(x => {
    roundRect(slide, { x, y: 3.881, w: 5.226, h: 5.852, rectRadius: 0.125, fill: { color: CARD }, line: { type: 'none' } });
    text(slide, 'ANYTHING WRITE HERE', { x, y: 4.51, w: 5.226, h: 0.505, fontFace: POPPINS_MD, fontSize: 24, color: WHITE, align: 'center' });
    imageArea(slide, x + 0.623, 5.381, 3.979, 1.667, 0.833, CARD);
    roundRect(slide, { x: x + 2.74, y: 6.663, w: 1.118, h: 0.703, rectRadius: 0.3515, fill: { color: GREEN }, line: { color: WHITE, width: 3 } });
    chevrons(slide, x + 2.954, 6.833);
    text(slide, 'PLACEHOLDER',
      Object.assign({ x: x + 0.394, y: 7.624, w: 4.438, h: 1.557 }, BODY, { align: 'center' }));
  });
}

function slide16(slide) {
  kicker(slide, 1.559, 3.012, 'OUR SERVICES');
  heading(slide, { x: 1.466, y: 3.517, w: 6.162, h: 1.919, green: 'Paycash', rest: ' Best Service' });
  text(slide, 'The last person we talked to said this would be ready action item, and what do you feel you would bring to the table if you were hired for this position bells and whistles. We need to socialize the with thesa wide finance.',
    Object.assign({ x: 1.466, y: 5.672, w: 6.367, h: 2.566 }, BODY));

  [1.271, 4.388, 7.505].forEach(y => {
    roundRect(slide, { x: 9.542, y, w: 8.992, h: 2.473, rectRadius: 0.158, fill: { color: CARD }, line: { color: WHITE, width: 1 } });
    slide.addShape('ellipse', { x: 9.031, y: y + 0.368, w: 1.022, h: 1.022, fill: { color: GREEN }, line: { type: 'none' } });
    bankIcon(slide, 9.345, y + 0.683, 0.394, WHITE);
    text(slide, 'ANYTHING WRITE HERE', { x: 10.731, y: y + 0.442, w: 3.626, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    text(slide, 'The last person we talked to said this would be ready hisa action item, and what do you feel here.',
      Object.assign({ x: 10.731, y: y + 0.981, w: 7.254, h: 1.052 }, BODY));
  });
}

function slide17(slide) {
  text(slide, 'PRICING PLAN', Object.assign({ x: 7.984, y: 1.136, w: 4.032, h: 0.404, align: 'center' }, KICKER, { fontSize: 18 }));
  heading(slide, { x: 2.516, y: 1.685, w: 14.969, h: 1.01, green: 'Choose', rest: ' Your Pricing Plan', align: 'center' });

  const plans = [
    { x: 2.106, y: 3.603, name: 'BASIC PLAN', price: '$20' },
    { x: 7.815, y: 3.594, name: 'PREMIUM PLAN', price: '$50' },
    { x: 13.524, y: 3.585, name: 'ADVANCED PLAN', price: '$80' },
  ];
  plans.forEach(p => {
    rect(slide, { x: p.x, y: p.y, w: 4.64, h: 6.51, fill: { color: CARD }, line: { color: GREEN, width: 2 } });
    text(slide, p.name, { x: p.x + 0.724, y: p.y + 0.473, w: 3.192, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE, align: 'center' });
    text(slide, p.price, { x: p.x + 1.07, y: p.y + 0.979, w: 2.056, h: 0.841, fontFace: INTER_MD, fontSize: 44, color: GREEN });
    text(slide, '/Month', { x: p.x + 2.464, y: p.y + 1.196, w: 1.106, h: 0.502, fontFace: INTER, fontSize: 18, color: WHITE, lineSpacingMultiple: 1.5 });

    for (let i = 0; i < 5; i++) {
      const y = p.y + 2.257 + i * 0.571;
      slide.addShape('ellipse', { x: p.x + 0.735, y, w: 0.302, h: 0.302, fill: { color: GREEN }, line: { type: 'none' } });
      slide.addShape('custGeom', {
        x: p.x + 0.817, y: y + 0.098, w: 0.138, h: 0.105, fill: { color: WHITE }, line: { type: 'none' },
        points: [[1, 0.159], [0.881, 0], [0.347, 0.686], [0.12, 0.388], [0, 0.546], [0.346, 1]]
          .map(([px, py]) => ({ x: px * 0.138, y: py * 0.105 })).concat([{ close: true }]),
      });
      text(slide, 'Your Title Service', { x: p.x + 1.203, y: y - 0.153, w: 2.987, h: 0.502, fontFace: INTER, fontSize: 18, color: GREY_LT, lineSpacingMultiple: 1.5 });
    }

    rect(slide, { x: p.x + 1.014, y: p.y + 5.467, w: 2.612, h: 0.724, fill: { color: GREEN }, line: { type: 'none' } });
    text(slide, 'CHOOSE PLAN', { x: p.x + 1.014, y: p.y + 5.644, w: 2.612, h: 0.37, fontFace: POPPINS_MD, fontSize: 16, color: WHITE, align: 'center', charSpacing: 1 });
  });
}

function slide18(slide) {
  kicker(slide, 1.455, 1.899, 'TESTIMONIALS');
  heading(slide, { x: 1.362, y: 2.404, w: 9.242, h: 1.01, green: 'What', rest: ' Are The Saying' });

  [1.292, 7.181, 13.071].forEach(x => {
    rect(slide, { x, y: 4.54, w: 5.638, h: 4.521, fill: { color: CARD }, line: { color: GREEN, width: 1 } });
    imageArea(slide, x + 0.564, 5.209, 1.023, 1.023, 0.5115, CARD);
    text(slide, 'YOUR NAME HERE', { x: x + 1.881, y: 5.27, w: 3.408, h: 0.404, fontFace: POPPINS_MD, fontSize: 18, color: WHITE });
    text(slide, 'Job Position', { x: x + 1.881, y: 5.624, w: 3.408, h: 0.547, fontFace: INTER, fontSize: 20, color: GREEN, lineSpacingMultiple: 1.5 });
    text(slide, 'Leverage agile frameworks provide a robusta educat synopsis for high level overviews ativa is roaches its corporate strategy borative.',
      Object.assign({ x: x + 0.564, y: 6.526, w: 4.508, h: 1.865 }, BODY, { fontSize: 18 }));
    quoteMark(slide, x + 3.891, 8.491, 1.181, WHITE);
  });
}

function slide19(slide) {
  imageArea(slide, 10.0, 0, 10.0, 5.625);
  imageArea(slide, 12.146, 5.7, 7.854, 5.55);
  rect(slide, { x: 11.5, y: 6.415, w: 1.042, h: 4.418, fill: { color: GREEN }, line: { type: 'none' } });

  kicker(slide, 1.737, 1.882, 'CONTACT US');
  heading(slide, { x: 1.644, y: 2.387, w: 7.833, h: 1.01, green: 'More', rest: ' Information' });

  const rows = [
    { y: 3.903, icon: cityIcon, title: 'OFFICE ADDRESS', lines: ['7 East 68th Street, New York, NY 10065, United States'], w: 3.938 },
    { y: 5.899, icon: phoneIcon, title: 'OUR PHONE', lines: ['+08 8983 9843 9201', '+08 7685 3629 0045'], w: 4.261 },
    { y: 7.894, icon: emailIcon, title: 'EMAIL SUPPORT', lines: ['Info@paycash.com', 'contact@paycash.com'], w: 4.261 },
  ];
  rows.forEach(r => {
    r.icon(slide, 1.645, r.y + 0.002, 0.591, GREEN);
    text(slide, r.title, { x: 2.583, y: r.y, w: 4.358, h: 0.438, fontFace: POPPINS_MD, fontSize: 20, color: WHITE });
    text(slide, r.lines.map((t, i) => ({ text: t, options: { breakLine: i < r.lines.length - 1 } })),
      Object.assign({ x: 2.583, y: r.y + 0.517, w: r.w, h: 0.972 }, BODY, { fontSize: 18 }));
  });
}

/* ------------------------------------------------------------------ *
 * Deck assembly
 * ------------------------------------------------------------------ */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'PAYCASH', width: 20, height: 11.25 });
  pptx.layout = 'PAYCASH';

  pptx.defineSlideMaster({ title: 'PAYCASH_MASTER', background: { color: BG } });

  SLIDES.forEach(builder => {
    const slide = pptx.addSlide({ masterName: 'PAYCASH_MASTER' });
    logoMark(slide);
    builder(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0db22f93-8854-4d8b-8422-133d8f4d658f_grok_final.pptx') });
}

/** Two interlocking arcs + wordmark: the PAYCASH lockup on every slide. */
function logoMark(slide) {
  slide.addShape('arc', {
    x: 0.697, y: 0.491, w: 0.343, h: 0.343, rotate: 135, flipH: true,
    angleRange: [134.9, 312.8], line: { color: GREEN, width: 4.5 },
  });
  slide.addShape('arc', {
    x: 0.920, y: 0.511, w: 0.249, h: 0.249, rotate: 315, flipH: true,
    angleRange: [131.4, 312.8], line: { color: GREEN, width: 4.5 },
  });
  text(slide, 'PAYCASH', { x: 1.354, y: 0.477, w: 1.953, h: 0.37, fontFace: POPPINS_SB, fontSize: 16, color: WHITE });
}

build().then(f => console.log('wrote', f));
