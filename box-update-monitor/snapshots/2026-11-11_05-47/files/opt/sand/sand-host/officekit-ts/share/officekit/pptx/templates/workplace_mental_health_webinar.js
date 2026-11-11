/*
 * "From Stress to Strength — Mental Health at Work Webinar" (32 slides, 26.667 x 15 in).
 * Rebuilt with pptxgenjs only; run `node <this file>` to write the .pptx next to it.
 *
 * The source deck contains no embedded raster images -- every picture frame in it is an
 * empty placeholder -- so nothing here needs an image substitute.  Vector icon glyphs are
 * approximated with the ICON() / PERSON() helpers below.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 26.667;
const SLIDE_H = 15;

// ---------------------------------------------------------------- palette --
const BROWN  = '3E3630';   // primary dark background
const PINK   = 'F7CECB';   // accent / headline colour
const BLUSH  = 'FADFDD';   // pale pink card fill
const TAUPE  = '83715C';   // muted brown, footers & rules
const WHITE  = 'FFFFFF';
const BLACK  = '000000';
const GREY   = '7A787C';   // body copy on dark
const SLATE  = '5F5F5F';   // body copy on light
const SILVER = 'ABABAB';
const MIST   = 'D0CECF';
const LINE   = 'E6E6E6';   // hairline rules on white
const METAL  = '808080';   // device mock-up chassis

// ------------------------------------------------------------------ fonts --
const BODY  = 'Instrument Sans Regular';   // deck default (16 pt, white, all-caps off)
const HEAD  = 'Instrument Sans Bold';      // theme minor font, used for headings
const MED   = 'Instrument Sans Medium';    // big statistic numbers

// ---------------------------------------------------------------- helpers --
function newSlide(pptx, bg) {
  const s = pptx.addSlide();
  s.background = { color: bg };
  return s;
}

/** Text box. The deck's boxes are centre-anchored with a 4 pt inset. */
function T(s, x, y, w, h, text, opts) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: BODY, fontSize: 16, color: WHITE, margin: 4, valign: 'middle',
  }, opts));
}

/** Filled rectangle (transparency 0-100). */
function R(s, x, y, w, h, color, transparency) {
  s.addShape('rect', { x, y, w, h, fill: { color, transparency: transparency || 0 }, line: { type: 'none' } });
}

/** Rounded rectangle; r is the corner adjust as a fraction of the short side. */
function RR(s, x, y, w, h, color, r) {
  s.addShape('roundRect', { x, y, w, h, fill: { color }, line: { type: 'none' }, rectRadius: r * Math.min(w, h) });
}

function EL(s, x, y, w, h, color) {
  s.addShape('ellipse', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/** Horizontal hairline. */
function LN(s, x, y, w, color, pt) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: pt || 1 } });
}

/** Stand-in for the deck's small vector pictograms. */
function ICON(s, x, y, w, h, color) {
  RR(s, x + w * 0.08, y + h * 0.08, w * 0.84, h * 0.84, color, 0.28);
}

/** Person glyph used by the slide-15 pictogram: head above a domed pair of shoulders. */
function PERSON(s, x, y, w, h, color) {
  EL(s, x + w * 0.257, y + h * 0.03, w * 0.486, w * 0.486, color);
  const sy = y + h * 0.64;
  const sh = h * 0.36;
  s.addShape('custGeom', {
    x, y: sy, w, h: sh, fill: { color }, line: { type: 'none' },
    points: [
      { x: 0, y: sh },
      { curve: { type: 'arc', wR: w / 2, hR: sh, stAng: 180, swAng: 180 }, x: w, y: sh },
      { close: true },
    ],
  });
}

/** Repeated on every slide: copyright, deck name and page number. */
function footer(s, color) {
  const c = color || TAUPE;
  T(s, 1.286, 13.503, 1.814, 0.389, '\u00a9 2025', { color: c });
  T(s, 3.312, 13.225, 2.718, 0.667, 'Mindwell Webinar  /  Confidential', { color: c });
  T(s, 23.567, 13.503, 1.814, 0.389, '01', { color: c, align: 'right' });
}

// --------------------------------------------------- repeated copy blocks --
const GLOBALLYAROUNDOF =
  'Globally, around 25% of adults suffer from mental health conditions such as anxiety, depression, or burnout.';
const CHRONICSTRESSLEA =
  'Chronic stress leads to fatigue, headaches, and long-term illness.';
const POORMENTALHEALTH =
  'Poor mental health weakens focus and productivity.';
const CURSUSMIPRETIUM =
  'Cursus mi pretium tellus duis convallis tempus leo. Mattis scelerisque maximus eget fermentum odio phasellus non. Porttitor ullamcorper rutrum gravida cras eleifend turpis fames. Tempus leo eu aenean sed diam urna tempor lorem ipsum.';
const SURVEYSHOWSSTRES =
  'Survey shows stress increases with career stage, leading to reduced focus, higher absenteeism, and risk of burnout.';
const CURSUSMIPRETIU2 =
  'Cursus mi pretium tellus duis convallis tempus leo. Mattis scelerisque maximus eget fermentum odio phasellus non.';
const WEONBOARDEDLICEN =
  'We onboarded licensed psychologists and counselors across specialties and developed ethical SOPs.';
const GLOBALLYAROUND2 =
  'Globally, around 25% of adults suffer from mental health conditions.';
const EXISTINGPRODUCTS =
  'Existing products in the market lack the modern features';
const PRACTICALSOLUTIO =
  'Practical Solutions for Employees & Leaders';
const HALFOFDISENGAGED =
  'Half of disengaged staff link it to stress';
const CURSUSMIPRETIU3 =
  'Cursus mi pretium tellus duis convallis tempus leo.';
const GLOBALLYAROUND3 =
  'Globally, around 25% of adults suffer from mental health conditions such as anxiety, depression, or burnout. Yet, only a fraction receive consistent, quality care. The burden is heavy, both emotionally and economically.';
const GLOBALLYAROUND4 =
  'Globally, around 25% of adults suffer from mental health conditions such as anxiety, depression, or burnout. Yet, only a fraction receive consistent.';

// ------------------------------------------------------------------ slides --
function slide1(pptx) {
  const s = newSlide(pptx, BROWN);
  T(s, 1.286, 1.108, 17.657, 10.311, 'From Stress to Strength Mental Health at Work Webinar', { fontSize: 177, color: PINK, charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  footer(s);
}

function slide2(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 1.108, 9.916, 1.264, 'Table of Contents', { fontSize: 68, color: PINK });
  T(s, 9.388, 3.722, 2.932, 0.556, '18:00 - 18:20', { fontSize: 26, color: PINK });
  T(s, 13.439, 3.583, 10.7, 0.833, 'Why Mental Health Matters at Work', { fontSize: 42 });
  T(s, 13.439, 4.417, 5.865, 1, EXISTINGPRODUCTS, { fontSize: 26, color: GREY });
  T(s, 9.388, 5.972, 2.932, 0.556, '18:20 - 19:00', { fontSize: 26, color: PINK });
  T(s, 13.439, 5.833, 11.73, 0.833, 'Common Problems & Workplace Impact', { fontSize: 42 });
  T(s, 9.388, 7.222, 2.932, 0.556, '19:00 - 20:00', { fontSize: 26, color: PINK });
  T(s, 13.439, 7.083, 11.73, 0.833, 'Data & Insights On Stress', { fontSize: 42 });
  T(s, 9.388, 8.472, 2.932, 0.556, '20:00 - 20:30', { fontSize: 26, color: PINK });
  T(s, 13.439, 8.333, 11.941, 0.833, PRACTICALSOLUTIO, { fontSize: 42 });
  T(s, 13.439, 9.167, 5.865, 1, EXISTINGPRODUCTS, { fontSize: 26, color: GREY });
  T(s, 9.388, 10.722, 2.932, 0.556, '19:00 - 20:00', { fontSize: 26, color: PINK });
  T(s, 13.439, 10.583, 11.73, 0.833, 'Data & Insights On Stress', { fontSize: 42 });
  LN(s, 9.388, 5.632, 15.992, TAUPE);
  LN(s, 9.388, 6.868, 15.992, TAUPE);
  LN(s, 9.388, 8.132, 15.992, TAUPE);
  LN(s, 9.388, 10.375, 15.992, TAUPE);
}

function slide3(pptx) {
  const s = newSlide(pptx, WHITE);
  footer(s, SILVER);
  T(s, 3.027, 1.161, 6.434, 3.569, 'Why Mental Health at Work Matters', { fontSize: 68, color: PINK });
  T(s, 13.439, 1.161, 11.941, 5.875, 'Prioritizing mental health at work leads to healthier employees, stronger teams, and more sustainable performance.', { fontSize: 68, color: BROWN });
  T(s, 13.439, 9.359, 10.652, 2.778, 'Cursus mi pretium tellus duis convallis tempus leo. Mattis scelerisque maximus eget fermentum odio phasellus non. Porttitor ullamcorper rutrum gravida cras eleifend turpis fames. Tempus leo eu aenean sed diam urna tempor. Phasellus non purus est efficitur laoreet mauris pharetra. Turpis fames primis vulputate ornare sagittis vehicula praesent.', { fontSize: 26, color: GREY });
  ICON(s, 13.439, 8.17, 0.828, 0.773, PINK);
}

function slide4(pptx) {
  const s = newSlide(pptx, BROWN);
  LN(s, 1.286, 5.374, 24.095, TAUPE);
  LN(s, 1.286, 9.64, 24.095, TAUPE);
  footer(s);
  T(s, 1.286, 1.108, 5.616, 0.556, 'Overview of Today’s Speaker(s)', { fontSize: 26, color: TAUPE });
  T(s, 16.751, 1.108, 8.63, 1.264, 'Lars Holm', { fontSize: 68, color: PINK });
  T(s, 16.751, 2.65, 5.53, 1.444, [{ text: 'Psychologist & Researcher ', options: { breakLine: true } }, { text: 'Uppsala University', options: { breakLine: true } }, { text: 'Sweden', options: { color: TAUPE } }], { fontSize: 26 });
  T(s, 9.6, 6.007, 8.63, 1.264, 'Ibrahim Coneate', { fontSize: 68, color: PINK });
  T(s, 9.6, 7.549, 5.53, 1.444, [{ text: 'Wellbeing Consultant ', options: { breakLine: true } }, { text: 'Stockholm Health Institute ', options: { breakLine: true } }, { text: 'Ghana', options: { color: TAUPE } }], { fontSize: 26 });
  T(s, 6.312, 10.905, 12.921, 1.264, 'Eriksson Lindström Karlsson', { fontSize: 68, color: PINK });
  T(s, 6.312, 12.447, 5.53, 1.444, [{ text: 'Professor of Clinical Psychology ', options: { breakLine: true } }, { text: 'Lund University', options: { breakLine: true } }, { text: 'Finland', options: { color: TAUPE } }], { fontSize: 26 });
}

function slide5(pptx) {
  const s = newSlide(pptx, BROWN);
  R(s, 7.905, 2.697, 18.761, 4.748, TAUPE);
  R(s, 0, 7.444, 18.761, 4.748, WHITE);
  footer(s);
  T(s, 1.286, 1.108, 5.616, 0.556, 'Overview of Today’s Speaker(s)', { fontSize: 26, color: TAUPE });
  T(s, 9.6, 3.859, 8.63, 1.264, 'Ibrahim Coneate', { fontSize: 68, color: PINK });
  T(s, 1.286, 8.714, 8.645, 2.417, 'Eriksson Lindström Karlsson', { fontSize: 68, color: BROWN });
  T(s, 11.96, 9.922, 5.53, 1, [{ text: 'Professor of Clinical Psychology ', options: { breakLine: true } }, { text: '  /  Lund University' }], { fontSize: 26, color: BROWN });
  T(s, 11.96, 8.714, 4.772, 0.833, '18:20 - 19:00', { fontSize: 42, color: BLACK });
  T(s, 19.516, 5.282, 5.53, 1, [{ text: 'Wellbeing Consultant ', options: { breakLine: true } }, { text: '  /  Stockholm Health Institute ' }], { fontSize: 26 });
  T(s, 19.516, 4.074, 4.772, 0.833, '18:20 - 19:00', { fontSize: 42 });
}

function slide6(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 1.108, 9.916, 2.417, 'Eriksson Lindström Karlsson', { fontSize: 68, color: PINK });
  T(s, 1.286, 3.803, 9.009, 0.556, 'Professor of Clinical Psychology  / Lund University', { fontSize: 26 });
  LN(s, 1.286, 7.493, 9.916, TAUPE);
  T(s, 3.523, 6.181, 7.679, 1, 'Author of “The Science of Resilient Minds” (International Best-Seller)', { fontSize: 26 });
  T(s, 1.286, 6.181, 1.814, 0.556, '2018', { fontSize: 26, color: TAUPE });
  T(s, 3.523, 7.805, 7.679, 1, 'Experience in psychotherapy & academic research lorem ipsum', { fontSize: 26 });
  T(s, 1.286, 7.805, 1.814, 0.556, '15+ years', { fontSize: 26, color: TAUPE });
  LN(s, 1.286, 9.117, 9.916, TAUPE);
  T(s, 3.523, 9.428, 7.679, 1, [{ text: 'Keynote Speaker at APA ', options: { breakLine: true } }, { text: 'Annual Conference 2024' }], { fontSize: 26 });
  T(s, 1.286, 9.428, 1.814, 0.556, '2023', { fontSize: 26, color: TAUPE });
}

function slide7(pptx) {
  const s = newSlide(pptx, PINK);
  R(s, 11.414, 4.837, 13.967, 3.533, BLUSH);
  R(s, 11.414, 1.109, 13.967, 3.533, BLUSH);
  R(s, 11.414, 8.565, 13.967, 3.533, BLUSH);
  footer(s);
  R(s, 11.414, 4.641, 13.967, 0.195, PINK);
  T(s, 12.108, 1.876, 1.503, 0.389, '[ 01 ]', { color: TAUPE });
  T(s, 13.611, 1.667, 10.128, 2.417, 'Decreased Innovation and Creativity', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 12.108, 5.604, 1.503, 0.389, '[ 02 ]', { color: TAUPE });
  T(s, 13.611, 5.395, 10.128, 2.417, 'Strained Workplace Relationships', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 12.108, 9.332, 1.503, 0.389, '[ 03 ]', { color: TAUPE });
  T(s, 13.611, 9.123, 10.961, 2.417, HALFOFDISENGAGED, { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
}

function slide8(pptx) {
  const s = newSlide(pptx, PINK);
  R(s, 11.414, 4.837, 13.967, 3.533, BLUSH);
  R(s, 11.414, 1.109, 13.967, 3.533, BLUSH);
  R(s, 11.414, 8.565, 13.967, 3.533, BLUSH);
  footer(s);
  R(s, 11.414, 4.641, 13.967, 0.195, PINK);
  T(s, 12.108, 1.708, 1.503, 0.389, '[ 01 ]', { color: TAUPE });
  T(s, 12.108, 2.514, 10.569, 0.833, 'Decreased Innovation and Creativity', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 12.108, 3.486, 9.687, 0.556, 'Stress and anxiety lower focus, slowing task completion.', { fontSize: 26, color: GREY });
  T(s, 12.108, 5.436, 1.503, 0.389, '[ 02 ]', { color: TAUPE });
  T(s, 12.108, 6.242, 10.569, 0.833, 'Strained Workplace Relationships', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 12.108, 7.214, 12.578, 0.556, CHRONICSTRESSLEA, { fontSize: 26, color: GREY });
  T(s, 12.108, 9.159, 1.503, 0.389, '[ 03 ]', { color: TAUPE });
  T(s, 12.108, 9.964, 12.134, 0.833, HALFOFDISENGAGED, { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 12.108, 10.936, 12.578, 0.556, 'Employees feel disconnected and undervalued, increasing turnover risk.', { fontSize: 26, color: GREY });
}

function slide9(pptx) {
  const s = newSlide(pptx, BROWN);
  R(s, 1.286, 1.108, 9.916, 11.045, TAUPE);
  footer(s);
  T(s, 2.299, 1.872, 7.89, 4.722, POORMENTALHEALTH, { fontSize: 68, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 2.299, 7.862, 7.89, 1.889, 'Poor mental health weakens focus and productivity, leading to mistakes, lower efficiency, and reduced workplace performance lorem ipsum.', { fontSize: 26 });
  T(s, 2.299, 10.445, 7.89, 0.389, 'DATE');
  T(s, 2.299, 10.834, 7.89, 0.556, 'January 28', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 12.762, 1.108, 12.619, 2.278, 'Stress and anxiety lower focus, slowing task completion. Turpis fames primis vulputate ornare sagittis vehicula praesent.', { fontSize: 42 });
  T(s, 12.762, 4.82, 12.619, 7.333, [{ text: CURSUSMIPRETIUM, options: { breakLine: true } }, { text: '', options: { breakLine: true } }, { text: 'Phasellus non purus est efficitur laoreet mauris pharetra. Turpis fames primis vulputate ornare sagittis vehicula praesent.' }], { fontSize: 42, color: GREY });
}

function slide10(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s, WHITE);
  T(s, 1.286, 1.108, 13.967, 5.361, 'The Hidden Costs of Poor Mental Health in Workplaces', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 13.511, 7.261, 11.869, 5.167, 'Beyond personal struggles, poor mental health leads to lost efficiency, high turnover, and billions in economic losses every year. Poor mental health weakens focus and productivity, leading to mistakes, lower efficiency, and reduced workplace performance lorem ipsum.', { fontSize: 42 });
}

function slide11(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  LN(s, 1.286, 1.942, 24.095, TAUPE);
  T(s, 1.286, 1.108, 3.478, 0.556, 'Survey Insights', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 4.764, 1.108, 2.259, 0.556, 'Trends', { fontSize: 26, color: TAUPE });
  T(s, 7.023, 1.108, 2.259, 0.556, 'Key Factors', { fontSize: 26, color: TAUPE });
  T(s, 1.286, 2.712, 9.916, 0.556, [{ text: 'Employees report experiencing ' }, { text: 'high workplace stress', options: { fontFace: 'Instrument Sans Bold' } }], { fontSize: 26, color: TAUPE });
  T(s, 1.286, 3.189, 5.445, 3.111, '58.2', { fontSize: 177, charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  T(s, 6.029, 4.896, 5.865, 0.833, '%', { fontSize: 42 });
  R(s, 1.286, 6.879, 9.916, 1.831, TAUPE);
  R(s, 1.286, 8.822, 20.044, 1.831, PINK);
  R(s, 1.286, 10.764, 14.179, 1.831, WHITE);
  T(s, 1.703, 7.018, 5.865, 0.833, 'Early Career', { fontSize: 42 });
  T(s, 1.703, 8.961, 5.865, 0.833, 'Mid Career', { fontSize: 42, color: BROWN });
  T(s, 1.703, 10.903, 5.865, 0.833, 'Senior Career', { fontSize: 42, color: BROWN });
  T(s, 8.098, 7.222, 2.826, 0.556, '0-2 yrs', { fontSize: 26, align: 'right' });
  T(s, 18.225, 9.1, 2.826, 0.556, '3-7 yrs', { fontSize: 26, color: BROWN, align: 'right' });
  T(s, 12.36, 11.042, 2.826, 0.556, '8+ yrs', { fontSize: 26, color: BROWN, align: 'right' });
  T(s, 16.132, 2.712, 9.249, 3, SURVEYSHOWSSTRES, { fontSize: 42 });
  T(s, 19.516, 7.378, 5.865, 0.833, '46%', { fontSize: 42, align: 'right' });
  T(s, 19.516, 9.321, 5.865, 0.833, '62%', { fontSize: 42, align: 'right' });
  T(s, 19.516, 11.264, 5.865, 0.833, '71%', { fontSize: 42, align: 'right' });
}

function slide12(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  LN(s, 15.465, 7.529, 9.916, SILVER);
  LN(s, 15.465, 10.12, 9.916, SILVER);
  T(s, 20.082, 5.126, 5.299, 1.986, '92%', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, align: 'right', charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 15.465, 5.563, 4.617, 0.556, 'Client Satisfaction Rate', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 20.082, 7.831, 5.299, 1.986, '81%', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, align: 'right', charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 15.465, 8.046, 4.617, 1, 'Stress Reduction After 4 Sessions', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 17.849, 10.423, 7.532, 1.986, '3.6/Client', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, align: 'right', charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 15.465, 10.86, 4.617, 0.556, 'Average Sessions', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 5.903, 5.126, 5.299, 1.986, '42', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, align: 'right', charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 1.286, 5.563, 4.617, 0.556, 'Partner Companies ', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 5.903, 7.831, 5.299, 1.986, '4.8 / 5', { fontSize: 110, fontFace: 'Instrument Sans Medium', color: PINK, align: 'right', charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 1.286, 8.046, 4.617, 0.556, 'Average Therapist Rating', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  LN(s, 1.286, 7.543, 9.916, SILVER);
  T(s, 15.465, 1.108, 9.916, 3, SURVEYSHOWSSTRES, { fontSize: 42 });
  T(s, 1.286, 1.108, 5.616, 0.556, 'Overview of Today’s Speaker(s)', { fontSize: 26, color: TAUPE });
}

function slide13(pptx) {
  const s = newSlide(pptx, BROWN);
  R(s, 15.465, 1.108, 9.916, 1.015, PINK, 88);
  footer(s);
  R(s, 15.465, 2.123, 9.916, 10.753, PINK);
  R(s, 13.333, 5.373, 9.916, 7.504, WHITE);
  T(s, 17.951, 2.262, 5.299, 3.111, '85%', { fontSize: 177, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'right', charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  T(s, 16.123, 2.832, 2.274, 1.444, 'Client Satisfaction Rate', { fontSize: 26, color: BROWN });
  T(s, 15.755, 6.067, 5.299, 3.111, '56%', { fontSize: 177, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'right', charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  T(s, 13.927, 6.638, 2.274, 1.444, 'Client Satisfaction Rate', { fontSize: 26, color: BROWN });
  T(s, 1.286, 1.108, 9.115, 3.569, POORMENTALHEALTH, { fontSize: 68, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 1.286, 5.373, 10.537, 5.167, CURSUSMIPRETIUM, { fontSize: 42, color: LINE });
}

function slide14(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  R(s, 1.286, 5.029, 15.992, 1.831, PINK);
  R(s, 17.279, 5.029, 8.102, 1.831, WHITE);
  T(s, 1.286, 6.86, 7.89, 3.111, '672M', { fontSize: 177, charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  T(s, 17.279, 6.86, 7.89, 3.111, '52M', { fontSize: 177, charSpacing: -5.31, lineSpacingMultiple: 0.8 });
  T(s, 1.703, 5.749, 5.865, 0.833, '76%', { fontSize: 42, color: BROWN });
  T(s, 17.695, 5.749, 5.865, 0.833, '32%', { fontSize: 42, color: BROWN });
  T(s, 1.286, 1.108, 11.941, 2.278, SURVEYSHOWSSTRES, { fontSize: 42 });
  T(s, 1.286, 9.971, 5.616, 1, CURSUSMIPRETIU3, { fontSize: 26 });
  T(s, 17.279, 9.971, 5.616, 1, CURSUSMIPRETIU3, { fontSize: 26 });
}

function slide15(pptx) {
  const s = newSlide(pptx, BROWN);
  // 75 person glyphs: 5 rows x 15, coloured by stress bucket
  const picto = [
    { y: 2.514, runs: [[PINK, 15]] },
    { y: 3.811, runs: [[PINK, 15]] },
    { y: 5.107, runs: [[PINK, 15]] },
    { y: 6.404, runs: [[PINK, 12], [WHITE, 3]] },
    { y: 7.7, runs: [[WHITE, 8], [TAUPE, 7]] },
  ];
  picto.forEach(row => {
    let i = 0;
    row.runs.forEach(([c, n]) => { for (let k = 0; k < n; k++, i++) PERSON(s, 7.363 + i * 1.237, row.y, 0.7, 0.7, c); });
  });
  footer(s);
  T(s, 7.363, 1.108, 3.526, 0.639, 'The Insights', { fontSize: 32, fontFace: 'Work Sans Regular Regular', color: '8B8B8B', lineSpacingMultiple: 1.1 });
  LN(s, 7.363, 1.921, 18.018, MIST, 3);
  ICON(s, 7.363, 11.237, 0.639, 0.639, PINK);
  T(s, 8.14, 11.278, 4.431, 0.556, 'Experiencing high stress', { fontSize: 26, color: GREY });
  ICON(s, 7.363, 12.245, 0.639, 0.639, WHITE);
  T(s, 8.14, 12.286, 3.527, 0.556, 'Occasional stress', { fontSize: 26, color: GREY });
  T(s, 8.14, 13.294, 3.358, 0.556, 'Low or no stress', { fontSize: 26, color: GREY });
  ICON(s, 7.363, 13.253, 0.639, 0.639, TAUPE);
  T(s, 14.514, 13.294, 3.054, 0.556, '= 10 People', { fontSize: 26, color: GREY });
  ICON(s, 13.869, 13.322, 0.431, 0.431, PINK);
  T(s, 13.447, 13.26, 0.421, 0.556, '1', { fontSize: 26, color: GREY });
  T(s, 1.286, 1.108, 4.99, 3, [{ text: '1 out of 3 ', options: { color: PINK } }, { text: 'employees experience high stress at work', options: { color: WHITE } }], { fontSize: 42, color: BLACK });
  T(s, 1.286, 4.439, 4.549, 2.333, CURSUSMIPRETIU2, { fontSize: 26, color: GREY });
}

function slide16(pptx) {
  const s = newSlide(pptx, WHITE);
  R(s, 1.286, 1.109, 7.89, 10.989, BROWN);
  footer(s);
  T(s, 1.944, 8.908, 6.787, 2.333, 'Cursus mi pretium tellus duis convallis tempus leo. Mattis scelerisque maximus eget fermentum odio phasellus non. Porttitor ullamcorper rutrum gravida cras eleifend turpis fames. ', { fontSize: 26, color: GREY });
  T(s, 1.944, 1.964, 6.787, 3.722, WEONBOARDEDLICEN, { fontSize: 42, color: PINK });
}

function slide17(pptx) {
  const s = newSlide(pptx, BROWN);
  R(s, 1.286, 6.18, 11.963, 1.444, TAUPE);
  R(s, 13.417, 6.18, 11.963, 1.444, TAUPE);
  R(s, 1.286, 7.763, 11.963, 1.444, TAUPE);
  R(s, 13.417, 7.763, 11.963, 1.444, TAUPE);
  R(s, 1.286, 9.346, 11.963, 1.444, TAUPE);
  R(s, 13.417, 9.346, 11.963, 1.444, TAUPE);
  R(s, 1.286, 10.93, 11.963, 1.444, TAUPE);
  R(s, 1.286, 6.18, 1.814, 1.444, '332D28');
  R(s, 1.286, 7.763, 1.814, 1.444, '332D28');
  R(s, 1.286, 9.346, 1.814, 1.444, '332D28');
  R(s, 1.286, 10.93, 1.814, 1.444, '332D28');
  R(s, 13.417, 6.18, 1.814, 1.444, '332D28');
  R(s, 13.417, 7.763, 1.814, 1.444, '332D28');
  R(s, 13.417, 9.346, 1.814, 1.444, '332D28');
  footer(s);
  T(s, 1.286, 1.108, 11.777, 2.278, WEONBOARDEDLICEN, { fontSize: 42, color: PINK });
  T(s, 1.286, 4.222, 7.89, 1.444, CURSUSMIPRETIU2, { fontSize: 26, color: GREY });
  ICON(s, 1.873, 6.55, 0.641, 0.705, PINK);
  ICON(s, 1.9, 8.261, 0.586, 0.449, PINK);
  ICON(s, 1.873, 11.396, 0.641, 0.513, PINK);
  ICON(s, 1.937, 9.754, 0.513, 0.63, PINK);
  ICON(s, 14.02, 6.575, 0.653, 0.655, PINK);
  ICON(s, 14.068, 8.158, 0.513, 0.655, PINK);
  ICON(s, 14.068, 9.738, 0.513, 0.66, PINK);
  T(s, 3.655, 6.624, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 3.655, 8.207, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 3.655, 9.791, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 3.655, 11.374, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 15.786, 6.624, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 15.786, 8.207, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
  T(s, 15.786, 9.791, 7.89, 0.556, 'Cursus mi pretium tellus duis', { fontSize: 26 });
}

function slide18(pptx) {
  const s = newSlide(pptx, PINK);
  R(s, 15.465, 8.345, 9.916, 2.491, BROWN);
  footer(s);
  T(s, 15.465, 1.108, 8.702, 3.569, 'How Poor Mental Health Affects Work Performance', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 15.465, 5.095, 9.405, 2.333, CURSUSMIPRETIUM, { fontSize: 26, color: BLACK });
  T(s, 17.655, 8.744, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold' });
  T(s, 17.655, 9.438, 7.35, 1, CHRONICSTRESSLEA, { fontSize: 26, color: SILVER });
  ICON(s, 16.102, 9.258, 0.87, 0.667, PINK);
  RR(s, 7.233, 3.322, 6.021, 12.537, METAL, 0.106);
  R(s, 13.273, 7.279, 0.061, 0.836, METAL);
  R(s, 7.159, 7.446, 0.061, 0.669, METAL);
  R(s, 7.159, 6.61, 0.061, 0.669, METAL);
  RR(s, 7.296, 3.385, 5.896, 12.412, BLACK, 0.099);
  RR(s, 7.526, 3.615, 5.436, 11.952, BLACK, 0.085);
}

function slide19(pptx) {
  const s = newSlide(pptx, WHITE);
  R(s, 0, 0, 26.667, 6.211, BROWN);
  footer(s, WHITE);
  T(s, 1.286, 0.689, 13.852, 2.417, POORMENTALHEALTH, { fontSize: 68, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 18.657, 1.175, 6.723, 1.444, WEONBOARDEDLICEN, { fontSize: 26 });
  RR(s, 7.947, 3.701, 10.847, 8.336, METAL, 0.061);
  R(s, 7.873, 4.951, 0.05, 0.692, METAL);
  RR(s, 7.997, 3.752, 10.746, 8.235, BLACK, 0.057);
  RR(s, 8.349, 4.103, 10.043, 7.532, BLACK, 0.031);
  R(s, 8.912, 3.631, 0.553, 0.05, METAL);
  R(s, 9.636, 3.631, 0.553, 0.05, METAL);
}

function slide20(pptx) {
  const s = newSlide(pptx, BROWN);
  RR(s, 7.675, 1.108, 6.139, 12.783, METAL, 0.106);
  R(s, 13.833, 5.143, 0.062, 0.852, METAL);
  R(s, 7.6, 5.313, 0.062, 0.682, METAL);
  R(s, 7.6, 4.461, 0.062, 0.682, METAL);
  RR(s, 7.739, 1.172, 6.011, 12.655, BLACK, 0.099);
  RR(s, 7.973, 1.407, 5.542, 12.187, BLACK, 0.085);
  R(s, 15.465, 1.108, 9.916, 2.491, WHITE);
  R(s, 15.465, 3.739, 9.916, 2.491, WHITE);
  R(s, 15.465, 6.369, 9.916, 2.491, WHITE);
  R(s, 15.465, 8.999, 9.916, 2.491, WHITE);
  footer(s);
  T(s, 1.286, 1.108, 6.077, 2.278, POORMENTALHEALTH, { fontSize: 42, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 17.655, 1.507, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 17.655, 2.201, 7.35, 1, CHRONICSTRESSLEA, { fontSize: 26, color: SILVER });
  ICON(s, 16.102, 2.021, 0.87, 0.667, PINK);
  T(s, 17.655, 4.137, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 17.655, 4.832, 7.35, 1, CHRONICSTRESSLEA, { fontSize: 26, color: SILVER });
  T(s, 17.655, 6.768, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 17.655, 7.462, 7.35, 1, CHRONICSTRESSLEA, { fontSize: 26, color: SILVER });
  T(s, 17.655, 9.398, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 17.655, 10.092, 7.35, 1, CHRONICSTRESSLEA, { fontSize: 26, color: SILVER });
  T(s, 1.286, 3.664, 5.7, 1.889, CURSUSMIPRETIU2, { fontSize: 26, color: GREY });
  ICON(s, 16.199, 7.128, 0.604, 0.743, PINK);
  ICON(s, 16.199, 4.613, 0.675, 0.743, PINK);
  ICON(s, 16.168, 9.885, 0.739, 0.72, PINK);
}

function slide21(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 11.414, 1.108, 13.427, 8.181, 'Our webinar highlighted the challenges of workplace stress, shared key data-driven insights, and presented practical solutions to support healthier employees and stronger organizations.', { fontSize: 68, color: PINK });
  T(s, 11.414, 9.567, 9.655, 1.889, GLOBALLYAROUND3, { fontSize: 26, color: MIST });
  T(s, 1.286, 1.192, 5.616, 0.389, 'Summary', { color: TAUPE });
}

function slide22(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 1.108, 9.916, 1.264, 'Table of Contents', { fontSize: 68, color: PINK });
  T(s, 1.286, 2.65, 7.89, 1.889, 'Globally, around 25% of adults suffer from mental health conditions such as anxiety, depression, or burnout. Yet, only a fraction receive consistent, quality care.', { fontSize: 26, color: MIST });
  T(s, 11.414, 11.657, 1.814, 0.389, 'Preface');
  T(s, 11.414, 12.185, 7.89, 0.667, GLOBALLYAROUND4);
  T(s, 11.414, 13.225, 3.067, 0.667, [{ text: 'None Report Producer,', options: { breakLine: true } }, { text: 'Date' }]);
  T(s, 11.414, 1.243, 1.814, 0.556, '01', { fontSize: 26, color: PINK });
  T(s, 13.439, 1.108, 11.73, 0.833, 'Common Problems & Workplace Impact', { fontSize: 42 });
  T(s, 13.439, 2.081, 11.73, 0.833, 'Data & Insights On Stress', { fontSize: 42 });
  T(s, 13.439, 3.053, 11.941, 0.833, PRACTICALSOLUTIO, { fontSize: 42 });
  T(s, 11.414, 2.22, 1.814, 0.556, '02', { fontSize: 26, color: PINK });
  T(s, 11.414, 3.192, 1.814, 0.556, '03', { fontSize: 26, color: PINK });
  T(s, 13.439, 4.025, 11.73, 0.833, 'Overview of Top Ranked Themes', { fontSize: 42 });
  T(s, 13.439, 4.997, 11.73, 0.833, 'Report Data Sources', { fontSize: 42 });
  T(s, 13.439, 5.97, 11.941, 0.833, 'Stats', { fontSize: 42 });
  T(s, 13.439, 6.942, 11.941, 0.833, 'Creadits - Report Contributors', { fontSize: 42 });
  T(s, 11.414, 4.164, 1.814, 0.556, '04', { fontSize: 26, color: PINK });
  T(s, 11.414, 5.136, 1.814, 0.556, '05', { fontSize: 26, color: PINK });
  T(s, 11.414, 6.108, 1.814, 0.556, '06', { fontSize: 26, color: PINK });
  T(s, 11.414, 7.081, 1.814, 0.556, '07', { fontSize: 26, color: PINK });
}

function slide23(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 1.108, 13.852, 2.417, 'Most current solutions require advanced technical expertise', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 1.286, 3.942, 15.992, 3, 'Organizations generate massive amounts of data daily but lack the tools to process and extract meaningful insights from it. Relying on human analysis alone leads to delays, mistakes, and missed opportunities in fast-paced markets.', { fontSize: 42 });
  R(s, 18.23, 0, 8.437, 15, WHITE);
  T(s, 1.286, 8.236, 3.839, 0.556, 'Lack of Trust in AI', { fontSize: 26 });
  T(s, 1.286, 9.125, 4.051, 2.333, 'Many businesses fear bias, errors, or ethical risks, which slows adoption despite AI’s potential benefits.', { fontSize: 26, color: MIST, italic: true });
  T(s, 6.138, 9.125, 4.733, 1.889, 'Tools and data are scattered across platforms, creating complexity and inefficiency in workflows.', { fontSize: 26, color: MIST, italic: true });
  T(s, 11.884, 9.125, 4.521, 2.333, 'There’s a global shortage of skilled AI professionals, making it hard for companies to build and maintain solutions.', { fontSize: 26, color: MIST, italic: true });
  T(s, 19.41, 1.132, 6.077, 1.556, [{ text: 'The Problem ', options: { breakLine: true } }, { text: 'We\'re Solving' }], { fontSize: 42, color: BROWN });
  ICON(s, 19.41, 4.392, 0.681, 0.622, BROWN);
  ICON(s, 19.55, 5.431, 0.462, 0.59, BROWN);
  ICON(s, 19.525, 3.315, 0.513, 0.66, BROWN);
  T(s, 20.398, 3.367, 4.749, 0.556, 'Manual decision-making', { fontSize: 26, color: BROWN });
  T(s, 20.398, 4.425, 3.51, 0.556, 'Existing AI tools', { fontSize: 26, color: BROWN });
  T(s, 20.398, 5.448, 4.983, 0.556, 'High cost adopting AI', { fontSize: 26, color: BROWN });
  T(s, 19.41, 9.447, 6.077, 1.264, '500', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 19.41, 10.711, 4.749, 0.556, 'Customers impacted', { fontSize: 26, color: BROWN });
  T(s, 19.41, 11.878, 6.077, 1.264, '$2 Million', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 19.41, 13.142, 4.749, 0.556, 'Connected account value', { fontSize: 26, color: BROWN });
  T(s, 19.41, 6.621, 6.406, 1.444, 'I often find myself lost in the current interface. A more straightforward layout would save a lot of time.', { fontSize: 26, color: GREY });
  LN(s, 19.41, 4.184, 5.971, LINE, 3);
  LN(s, 19.41, 5.214, 5.971, LINE, 3);
  T(s, 6.138, 8.236, 4.521, 0.556, 'Fragmented AI Ecosystem', { fontSize: 26 });
  T(s, 11.884, 8.236, 4.521, 0.556, 'Limited AI Talent', { fontSize: 26 });
}

function slide24(pptx) {
  const s = newSlide(pptx, WHITE);
  EL(s, 2.228, 1.143, 10.965, 10.965, BROWN);
  footer(s);
  EL(s, 10.309, 1.602, 1.326, 1.326, PINK);
  EL(s, 10.295, 10.324, 1.326, 1.326, PINK);
  EL(s, 2.221, 8.35, 1.326, 1.326, PINK);
  T(s, 4.121, 5.417, 7.178, 2.417, 'Voice of Customer', { fontSize: 68, align: 'center' });
  T(s, 11.913, 1.987, 3.111, 0.556, 'Understand', { fontSize: 26, color: BROWN });
  T(s, 1.286, 8.735, 1.814, 0.556, 'Act', { fontSize: 26, color: BROWN });
  T(s, 11.759, 10.709, 1.814, 0.556, 'Respond', { fontSize: 26, color: BROWN });
  ICON(s, 10.632, 1.921, 0.681, 0.622, BROWN);
  ICON(s, 10.727, 10.692, 0.462, 0.59, BROWN);
  ICON(s, 2.627, 8.683, 0.513, 0.66, BROWN);
  T(s, 14.605, 3.166, 10.809, 4.722, 'Presented practical solutions to support healthier employees and stronger organizations.', { fontSize: 68, color: BLACK });
  T(s, 14.605, 8.304, 9.655, 1.889, GLOBALLYAROUND3, { fontSize: 26, color: SILVER });
}

function slide25(pptx) {
  const s = newSlide(pptx, PINK);
  footer(s);
  T(s, 1.286, 1.108, 13.426, 2.417, POORMENTALHEALTH, { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 1.286, 6.997, 4.235, 0.833, 'Listen', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 1.286, 7.969, 4.235, 2.333, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  EL(s, 1.286, 5.677, 1.042, 1.042, WHITE);
  T(s, 1.45, 5.781, 0.713, 0.833, '1', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'center' });
  T(s, 6.244, 6.997, 4.235, 0.833, 'Understand', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 6.244, 7.969, 4.235, 2.333, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  EL(s, 6.251, 5.677, 1.042, 1.042, WHITE);
  T(s, 6.415, 5.781, 0.713, 0.833, '2', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'center' });
  T(s, 11.23, 6.997, 4.235, 0.833, 'Act', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 11.23, 7.969, 4.235, 2.333, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  EL(s, 11.216, 5.677, 1.042, 1.042, WHITE);
  T(s, 11.38, 5.781, 0.713, 0.833, '3', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'center' });
  T(s, 16.181, 6.997, 4.235, 0.833, 'Respond', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 16.181, 7.969, 4.235, 2.333, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  EL(s, 16.181, 5.677, 1.042, 1.042, WHITE);
  T(s, 16.345, 5.781, 0.713, 0.833, '4', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'center' });
  T(s, 21.145, 6.997, 4.235, 0.833, 'Golarize', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 21.145, 7.969, 4.235, 2.333, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  EL(s, 21.145, 5.677, 1.042, 1.042, WHITE);
  T(s, 21.31, 5.781, 0.713, 0.833, '5', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK, align: 'center' });
}

function slide26(pptx) {
  const s = newSlide(pptx, WHITE);
  footer(s);
  R(s, 15.992, 0, 10.674, 15, PINK);
  T(s, 17.49, 4.317, 7.89, 1.986, '672 Million', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 17.49, 6.303, 5.616, 0.556, 'Monthly Tickets', { fontSize: 26, color: SLATE });
  T(s, 17.49, 7.833, 7.89, 1.986, '24.5 min', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 17.49, 9.819, 5.616, 0.556, 'Average Handle Time (AHT)', { fontSize: 26, color: SLATE });
  T(s, 17.49, 11.35, 7.89, 1.986, '68%', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 17.49, 13.336, 5.616, 0.556, 'Negative CSAT', { fontSize: 26, color: SLATE });
  T(s, 17.49, 1.108, 7.726, 0.556, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 17.49, 1.803, 7.35, 1.444, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  LN(s, 17.49, 7.521, 7.89, BROWN);
  LN(s, 17.49, 10.863, 7.89, BROWN);
}

function slide27(pptx) {
  const s = newSlide(pptx, WHITE);
  footer(s);
  T(s, 1.286, 1.108, 21.762, 1.264, 'Making AI Accessible, Scalable, and Impactful', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 1.286, 3.519, 3.839, 0.556, 'Customer Insights', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 1.286, 4.213, 5.234, 1.889, 'We centralize and clean unstructured data, turning it into actionable insights in real time lorem ipsum.', { fontSize: 26, color: SLATE });
  T(s, 7.363, 3.519, 5.234, 0.556, 'Business & Market Insights', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 7.363, 4.213, 5.234, 1.889, 'AI-driven systems accelerate analysis and reduce human error in complex decision-making lorem.', { fontSize: 26, color: SLATE });
  T(s, 13.439, 3.519, 3.839, 0.556, 'Insights Consumers', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 13.439, 4.213, 5.234, 1.444, 'Intuitive interfaces allow non-technical teams to build and deploy AI models easily.', { fontSize: 26, color: SLATE });
  T(s, 19.516, 3.519, 3.839, 0.556, 'Business Strategy', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 19.516, 4.213, 5.234, 1.889, 'Our solution connects seamlessly with existing business systems, reducing setup time and cost.', { fontSize: 26, color: SLATE });
  LN(s, 1.286, 6.574, 5.234, LINE, 3);
  LN(s, 7.363, 6.574, 5.234, LINE, 3);
  LN(s, 13.439, 6.574, 5.234, LINE, 3);
  LN(s, 19.516, 6.574, 5.234, LINE, 3);
  T(s, 1.286, 7.046, 5.234, 0.389, 'INTEGRATED AI PLATFORM', { color: BROWN });
  T(s, 1.286, 7.435, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 7.363, 7.046, 5.234, 0.389, 'ETHICAL AND TRANSPARENT AI FRAMEWORK', { color: BROWN });
  T(s, 7.363, 7.435, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 13.439, 7.046, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 13.439, 7.435, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 19.516, 7.046, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 19.516, 7.435, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 1.286, 8.452, 5.234, 0.389, 'AI TALENT SUPPORT HUB', { color: BROWN });
  T(s, 1.286, 8.841, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 7.363, 8.452, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 7.363, 8.841, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 13.439, 8.452, 5.234, 0.389, 'USER-FRIENDLY AI TOOLS', { color: BROWN });
  T(s, 13.439, 8.841, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 19.516, 8.452, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 19.516, 8.841, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 1.286, 9.857, 5.234, 0.389, 'ROI TRACKING DASHBOARD', { color: BROWN });
  T(s, 1.286, 10.246, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 7.363, 9.857, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 7.363, 10.246, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 13.439, 9.857, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 13.439, 10.246, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
  T(s, 19.516, 9.857, 5.234, 0.389, 'TITLE', { color: BROWN });
  T(s, 19.516, 10.246, 5.234, 0.556, 'Lorem Ipsum Dolor', { fontSize: 26, color: SLATE });
}

function slide28(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  R(s, 9.454, 1.108, 7.89, 5.076, WHITE);
  R(s, 17.622, 6.462, 7.89, 5.076, PINK);
  T(s, 10.02, 1.633, 6.546, 1.556, 'Unclear ROI and Business Value', { fontSize: 42, color: BLACK });
  T(s, 10.02, 3.771, 6.758, 1.889, GLOBALLYAROUND4, { fontSize: 26, color: SILVER });
  T(s, 18.056, 6.819, 6.97, 1.986, '6.18%', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 18.056, 9, 4.617, 0.556, 'Lack of Trust in AI', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 18.056, 9.556, 6.97, 1.444, 'At Tastebud, our primary goal is to deliver high-quality, delicious meals that offer convenience for busy professionals.', { fontSize: 26, color: GREY });
}

function slide29(pptx) {
  const s = newSlide(pptx, WHITE);
  footer(s);
  R(s, 15.465, 1.108, 9.916, 6.392, PINK);
  T(s, 16.372, 1.72, 8.102, 0.833, 'Business Insights', { fontSize: 42, fontFace: 'Instrument Sans Bold', color: BLACK });
  T(s, 16.372, 2.692, 8.102, 1.444, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  T(s, 16.372, 4.438, 8.102, 1.722, [{ text: 'Lorem ipsum dolor dodumaskasda', options: { bullet: true, breakLine: true } }, { text: 'Lorem ipsum dolor dodumaskasda', options: { bullet: true, breakLine: true } }, { text: 'Lorem ipsum dolor dodumaskasda', options: { bullet: true } }], { fontSize: 26, color: SLATE, indent: 26, paraSpaceBefore: 10 });
  T(s, 1.286, 1.108, 13.06, 2.417, 'We centralize and clean unstructured data', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 1.286, 4.136, 12.785, 2.278, 'AI-driven systems accelerate analysis and reduce human error in complex decision-making. Lorem ipsum dolor', { fontSize: 42, color: GREY });
  T(s, 7.363, 7.769, 7.89, 1.986, '78.9%', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 7.363, 10.024, 7.726, 0.556, 'Cross-Industry Use Case Library', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 7.363, 10.719, 7.35, 1.444, GLOBALLYAROUNDOF, { fontSize: 26, color: SLATE });
  R(s, 1.286, 8.004, 5.224, 4.16, PINK);
  T(s, 1.662, 8.452, 3.675, 1, 'Strained Workplace Relationships', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 1.662, 9.826, 4.472, 1.889, 'Globally, around 25% of adults suffer from mental health conditions such as anxiety, depression', { fontSize: 26, color: SLATE });
}

function slide30(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 1.108, 4.958, 1.556, 'Meet the Keynote Speaker', { fontSize: 42, fontFace: 'Instrument Sans Bold' });
  T(s, 7.363, 5.715, 3.639, 0.389, 'Principal', { color: SILVER });
  T(s, 7.363, 5.159, 3.292, 0.556, 'Lars Holm', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 13.439, 5.715, 3.639, 0.389, 'Partner', { color: SILVER });
  T(s, 13.439, 5.159, 3.292, 0.556, 'Wade Warren', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 19.516, 5.715, 3.639, 0.389, 'Executive Venture Partner', { color: SILVER });
  T(s, 19.516, 5.159, 3.292, 0.556, 'Anna Berg', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 7.363, 11.301, 3.639, 0.389, 'General Partner', { color: SILVER });
  T(s, 7.363, 10.745, 3.292, 0.556, 'Mike Simmons', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 13.439, 11.301, 3.639, 0.389, 'Venture Partner', { color: SILVER });
  T(s, 13.439, 10.745, 3.292, 0.556, 'Benedict Johnson', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 19.516, 11.301, 3.639, 0.389, 'Entrepreneur in Residence', { color: SILVER });
  T(s, 19.516, 10.745, 3.292, 0.556, 'David Chen', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: PINK });
}

function slide31(pptx) {
  const s = newSlide(pptx, WHITE);
  RR(s, 10.261, 1.108, 6.139, 12.783, METAL, 0.106);
  R(s, 16.419, 5.143, 0.062, 0.852, METAL);
  R(s, 10.186, 5.313, 0.062, 0.682, METAL);
  R(s, 10.186, 4.461, 0.062, 0.682, METAL);
  RR(s, 10.325, 1.172, 6.011, 12.655, BLACK, 0.099);
  RR(s, 10.559, 1.407, 5.542, 12.187, BLACK, 0.085);
  footer(s);
  RR(s, 10.519, 8.194, 5.629, 5.395, WHITE, 0.07);
  T(s, 10.969, 8.831, 4.729, 1.444, 'Built-in analytics show clear KPIs and value from every AI project launched.', { fontSize: 26, color: SLATE });
  T(s, 10.969, 10.912, 4.729, 1.986, '68%', { fontSize: 110, fontFace: 'Instrument Sans Bold', color: BROWN, charSpacing: -3.3, lineSpacingMultiple: 0.9 });
  T(s, 10.969, 12.704, 4.729, 0.389, 'Integrated AI Platform', { color: SLATE });
  R(s, 17.49, 3.837, 7.89, 2.233, WHITE);
  EL(s, 17.907, 4.259, 1.389, 1.389, PINK);
  ICON(s, 18.292, 4.613, 0.619, 0.681, BROWN);
  T(s, 19.574, 4.343, 5.093, 0.556, 'Unified Data Management', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 19.574, 4.898, 5.319, 0.667, GLOBALLYAROUND2, { color: SLATE });
  R(s, 17.49, 6.383, 7.89, 2.233, WHITE);
  T(s, 19.574, 6.889, 5.093, 0.556, 'Automated Decision Engine', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 19.574, 7.444, 5.319, 0.667, GLOBALLYAROUND2, { color: SLATE });
  R(s, 17.49, 8.93, 7.89, 2.233, WHITE);
  T(s, 19.574, 9.435, 5.093, 0.556, 'User Friendly AI Tools', { fontSize: 26, fontFace: 'Instrument Sans Bold', color: BROWN });
  T(s, 19.574, 9.991, 5.319, 0.667, GLOBALLYAROUND2, { color: SLATE });
  EL(s, 17.907, 6.806, 1.389, 1.389, PINK);
  ICON(s, 18.325, 7.16, 0.554, 0.681, BROWN);
  EL(s, 17.907, 9.352, 1.389, 1.389, PINK);
  ICON(s, 18.264, 9.788, 0.675, 0.517, BROWN);
  T(s, 1.286, 3.837, 8.195, 3.569, 'Simplifying the user interface for new casual users', { fontSize: 68, fontFace: 'Instrument Sans Bold', color: PINK });
  T(s, 1.286, 7.684, 7.89, 3, 'Users have expressed that the current interface is not intuitive, leading to confusion and inefficiencies.', { fontSize: 42 });
}

function slide32(pptx) {
  const s = newSlide(pptx, BROWN);
  footer(s);
  T(s, 1.286, 5.039, 9.222, 2.417, [{ text: 'Let\'s Get In Touch', options: { breakLine: true } }, { text: '—Contact Us' }], { fontSize: 68, color: PINK });
  T(s, 11.414, 1.09, 2.932, 0.389, 'OUR STUDIO', { color: TAUPE });
  T(s, 11.414, 1.601, 3.184, 0.944, 'Launch Tech Inc. 123 Tech Boulevard Innovation City, Techland 54321 United States');
  T(s, 15.465, 1.09, 2.932, 0.389, 'ONLINE', { color: TAUPE });
  T(s, 15.465, 1.591, 3.184, 0.667, [{ text: 'w/   www.reallygreatsite.com', options: { breakLine: true } }, { text: 'e/    info@studiogreatsite.com' }]);
  T(s, 19.516, 1.09, 2.932, 0.389, 'SOCIAL', { color: TAUPE });
  T(s, 19.516, 1.591, 3.839, 0.667, [{ text: 'IG /   reallygreatsite', options: { breakLine: true } }, { text: 'FB /    reallygreatsite' }]);
}


// ------------------------------------------------------------------- main --
const BUILDERS = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32];

function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  BUILDERS.forEach(fn => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '183d5ed9-d302-45c8-adcc-6eb6272f6918_grok_final.pptx') });
}

main().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
