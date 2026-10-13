/**
 * Recreation of "NEW BUILDING / CONSTRUCTION" deck (30 slides, 13.333 x 7.5 in)
 * with pptxgenjs.  Raster photos in the original are replaced by flat grey
 * placeholder shapes of the same geometry.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */
const NAVY = '004F8A', DKNAVY = '002745', CYAN = '00B0F0', DKCYAN = '0084B4',
      ORANGE = 'FF9933', DKORANGE = 'E57300', YELLOW = 'FFC000', DKYELLOW = 'BF9000',
      WHITE = 'FFFFFF', SLATE = '44546A', GREY_BG = 'F2F2F2', LTORANGE = 'FFC285',
      SKY = '5DD4FF', SAND = 'FFD966', SHADE = '994C00';
const PHOTO = 'CECECE', PHOTO_TXT = 'A8A8A8';
const SANS = 'Open Sans', HEAD = 'Montserrat', POPPINS = 'Poppins', POPPINS_L = 'Poppins Light';
const SHADOW = { type: 'outer', blur: 20, offset: 10, angle: 45, color: '000000', opacity: 0.2 };

/* -------------------------------------------------------- shared strings */
const L1 = 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec';
const L2 = 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec sit amet  consec tetur adla dolor';
const L2S = L2.slice(0, -6);   // slides 15/16 drop the trailing word
const L3 = 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec sit amet  consec tetur adla dolor sit consec tetur cons';
const L_DONUT = 'Lorem ipsum dolor sit amet, adipis sit dolor cons sitadipis sit dolor';
const L_CARD = 'Lorem ipsum dolor amet, adipis ectetur adipis, adipis cingetur adipis cing amet, adipis sit ectetur, adipis cing';
const QUOTE3 = ['\u201CLorem ipsum dolor ', 'amet  consec tetur adla sssipsum dolor sit ', 'consec sit\u201D'];
const SERENITY = 'A wonderful serenity has taken possession of my entire soul, like';
const EIUSMOD = 'eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad';
const NULLAM = 'Nullam Eu Tempor Purus. Nunc A Leo Magna, Sit Amet Consequat Risus. ';
const QUIS = 'Quis tellus id nulla mollis iaculis. Nam nec auctor arcu. Aenean molestie';
const PODCAST = 'Podcasting operational change management inside of workflows to establish a framework.';
const UTWISI = 'Ut Wisi Enim Ad Minim Veniam, Quis Nostrud Exerci Tation Ullamcorperut Wisi Enim Ad Minim Veniam, Quis Nostrud Exerci Tation';
const IPSUM_DUMMY = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley";

/* ---------------------------------------------------------- tiny helpers */
const shp = (s, kind, o) => s.addShape(kind, o);
const box = (s, x, y, w, h, color, extra) =>
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, extra || {}));

/** polygon points (fractions of w/h) -> pptxgenjs custGeom point list */
function poly(w, h, pts) {
  const out = pts.map(p => ({ x: +(p[0] * w).toFixed(3), y: +(p[1] * h).toFixed(3) }));
  out.push({ close: true });
  return out;
}

/** grey stand-in for a photograph */
function photo(s, x, y, w, h, pts) {
  const o = { x, y, w, h, fill: { color: PHOTO }, line: { type: 'none' } };
  if (pts) { o.points = poly(w, h, pts); s.addShape('custGeom', o); } else { s.addShape('rect', o); }
  if (w >= 1.5 && h >= 0.8) {
    s.addText('[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center',
      fontFace: SANS, fontSize: 12, color: PHOTO_TXT });
  }
}
const photos = (s, list) => list.forEach(p => photo(s, p[0], p[1], p[2], p[3], p[4]));

/** thin ring drawn with the "donut" preset (rectRadius sets ring thickness) */
function ring(s, x, y, w, h, color, thick) {
  s.addShape('donut', { x, y, w, h, fill: { color }, rectRadius: thick });
}

/** the decorative corner frame that appears on every slide */
function frame(s, labelColor) {
  box(s, 0.263, 0.232, 3.295, 0.453, CYAN);
  shp(s, 'rtTriangle', { x: 0, y: 0, w: 1.147, h: 1.147, rotate: 90, fill: { color: NAVY } });
  shp(s, 'rtTriangle', { x: 12.186, y: 6.353, w: 1.147, h: 1.147, rotate: 270, fill: { color: NAVY } });
  shp(s, 'rtTriangle', { x: 12.76, y: 0.157, w: 0.417, h: 0.417, rotate: 180, fill: { color: CYAN } });
  shp(s, 'rtTriangle', { x: 0.157, y: 6.926, w: 0.417, h: 0.417, fill: { color: CYAN } });
  s.addText('Building Photograph', { x: 0.958, y: 0.315, w: 2.476, h: 0.286, wrap: false,
    fontFace: SANS, fontSize: 11, color: WHITE, charSpacing: 3 });
  s.addText('NEW BUILDING', { x: 11.799, y: 4.976, w: 2.316, h: 0.438, rotate: 90, wrap: false,
    fontFace: SANS, fontSize: 20, bold: true, color: labelColor || WHITE });
}

/** the two big translucent triangles in the orange background */
function bigTriangles(s) {
  const f = { color: SHADE, transparency: 90 };
  shp(s, 'rtTriangle', { x: 0, y: 4.485, w: 3.015, h: 3.015, fill: f });
  shp(s, 'rtTriangle', { x: 10.318, y: 0, w: 3.015, h: 3.015, rotate: 180, fill: f });
}

/** navy plate carrying an orange Montserrat headline */
function titleTag(s, text, x, y, w, h, size, align) {
  s.addText(text, { x, y, w, h, wrap: false, align: align || 'left', fill: { color: NAVY },
    fontFace: HEAD, fontSize: size, bold: true, color: ORANGE, shadow: SHADOW });
}
const infoTitle = s => titleTag(s, 'INFOGRAPHIC SECTION', 3.78, 1.17, 5.773, 0.64, 32, 'center');

/** body copy in the orange slides: white, 11pt, 1.5 line spacing */
function body(s, text, x, y, w, h, o) {
  s.addText(text, Object.assign({ x, y, w, h, fontFace: SANS, fontSize: 11, color: WHITE,
    align: 'justify', lineSpacingMultiple: 1.5 }, o || {}));
}

/** small "chevron" bullet marker */
const chev = (s, x, y, d, color) =>
  shp(s, 'chevron', { x, y, w: d, h: d, fill: { color: color || NAVY } });

/** two–tone headline such as "Chat Client" */
const twoTone = (s, a, b, x, y, w, o) => s.addText(
  [{ text: a, options: { color: WHITE } }, { text: b, options: { color: NAVY } }],
  Object.assign({ x, y, w, h: 0.438, wrap: false, fontFace: SANS, fontSize: 20, bold: true }, o || {}));

/* ===================================================================== */
/*  slide builders                                                       */
/* ===================================================================== */

/* --- 1 & 30 : title / closing --------------------------------------- */
function coverSlide(s, headline, headWidth) {
  photo(s, 3.21, 0, 10.123, 7.5, [[0.185, 0], [1, 0], [0.815, 1], [0, 1]]);
  bigTriangles(s);
  frame(s);
  s.addText(headline, { x: 0.731, y: 2.677, w: headWidth, h: 1.212, wrap: false, fill: { color: NAVY },
    fontFace: HEAD, fontSize: 66, bold: true, color: ORANGE, shadow: SHADOW });
  box(s, 0.731, 4.129, 7.979, 0.7, CYAN, { shadow: SHADOW });
  s.addText([
    { text: 'Slide ', options: { fontSize: 14 } },
    { text: 'Presentations', options: { fontSize: 12 } },
    { text: ' Design About Building', options: { fontSize: 14 } }
  ], { x: 0.994, y: 4.252, w: 7.453, h: 0.454, align: 'center', lineSpacingMultiple: 1.5,
       fontFace: SANS, color: WHITE, charSpacing: 6 });
  body(s, L1, 0.717, 5.111, 2.438, 0.934);
  box(s, 0.824, 6.139, 0.519, 0.05, NAVY);
  s.addText('NEW SLIDE', { x: 0.6, y: 1.429, w: 3.501, h: 0.841, wrap: false,
    fontFace: SANS, fontSize: 44, bold: true, color: WHITE });
}

/* --- 2 : GREATEST BUILDING (timeline) -------------------------------- */
function slide02(s) {
  photos(s, [[0.842, 5.221, 2.6, 1.594], [6.88, 5.221, 2.6, 1.594]]);
  bigTriangles(s);
  box(s, 0.842, 3.099, 11.657, 1.838, LTORANGE);
  s.addShape('line', { x: 1.572, y: 3.587, w: 10.927, h: 0, line: { color: WHITE, width: 1 } });
  frame(s);
  titleTag(s, 'GREATEST BUILDING', 4.101, 1.193, 5.132, 0.64, 32, 'center');
  s.addText('Lorem ipsum dolor sit amet, adipis sit dolor cons sit amet, adipis ectetur adipis sit amet, adipis ectectetur adipis cing amet, adipis sit ectetur adipis, adipis cingetur adipis cing amet, adipis sit ectetur adipis, adipis cing',
    { x: 2.078, y: 2.119, w: 9.19, h: 0.627, align: 'center', lineSpacingMultiple: 1.5,
      fontFace: SANS, fontSize: 11, color: WHITE });

  const marks = [['+6758 K', 1.38, 1.31], ['+7543 K', 3.11, 3.02], ['+5432 K', 4.99, 4.9],
                 ['+2123 K', 6.872, 6.77], ['+9087 K', 8.752, 8.666], ['+6543 K', 10.633, 10.562]];
  marks.forEach(([label, mx, tx]) => {
    box(s, mx, 3.491, 0.193, 0.193, WHITE);
    s.addText([
      { text: label, options: { fontSize: 11, bold: true, breakLine: true } },
      { text: 'Lorem ipsum dolor sit amet  consec', options: { fontSize: 10.5 } }
    ], { x: tx, y: 3.84, w: 1.529, h: 0.749, fontFace: SANS, color: WHITE, lineSpacingMultiple: 1.5 });
  });

  [3.861, 9.9].forEach(x => {
    box(s, x, 5.221, 2.6, 1.594, NAVY, { shadow: SHADOW });
    s.addText(QUOTE3.map((t, i) => ({ text: t, options: { breakLine: i < 2 } })),
      { x: x + 0.214, y: 5.598, w: 2.173, h: 0.841, align: 'center',
        fontFace: SANS, fontSize: 11, bold: true, italic: true, color: ORANGE });
    s.addShape('custGeom', { x: x + 2.253, y: 5.36, w: 0.183, h: 0.125, fill: { color: ORANGE },
      points: poly(0.183, 0.125, [[0.5, 1], [1, 0.13], [0.89, 0], [0.5, 0.66], [0.11, 0], [0, 0.13]]) });
  });
}

/* --- 3 : HIGH BUILDING ---------------------------------------------- */
function slide03(s) {
  bigTriangles(s);
  box(s, 2.583, 0, 2.042, 7.5, NAVY);
  const par = [[0.106, 0], [1, 0], [0.894, 1], [0, 1]];
  photos(s, [[0.704, 1.147, 5.802, 2.451, par], [0.704, 3.902, 5.802, 2.451, par]]);
  frame(s);
  titleTag(s, 'HIGH BUILDING', 7.326, 1.428, 3.889, 0.64, 32);
  [2.816, 3.738, 4.66, 5.582].forEach((y, i) => {
    chev(s, 7.326, y, 0.348);
    body(s, L1, 7.941, y - 0.155, 4.079, 0.656);
  });
}

/* --- 4 : ABOUT BUILDING --------------------------------------------- */
function slide04(s) {
  photo(s, 0, 0, 13.333, 3.75);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'ABOUT BUILDING', 0.678, 3.397, 4.367, 0.64, 32);
  box(s, 5.498, 2.264, 3.152, 4.089, CYAN, { shadow: SHADOW });
  box(s, 8.96, 2.272, 3.152, 4.089, NAVY, { shadow: SHADOW });

  [['78%', 4.397], ['89%', 5.489]].forEach(([pct, y]) => {
    ring(s, 0.678, y, 0.88, 0.88, NAVY, 0.043);
    s.addText(pct, { x: 0.678, y, w: 0.88, h: 0.88, align: 'center', valign: 'middle',
      fontFace: SANS, fontSize: 14, bold: true, color: WHITE });
    s.addText('Lorem ipsum dolor sit amet, adipis sit dolor cons sitadipis sit dolor',
      { x: 1.78, y, w: 3.265, h: 0.841, fontFace: SANS, fontSize: 11, color: WHITE, lineSpacingMultiple: 2 });
  });

  // two feature cards
  const card = (icon, title, titleColor, textColor, btnFill, cx, cy, iconX) => {
    s.addText(title, { x: cx, y: cy, w: 2.52, h: 0.505, align: 'center', wrap: false,
      fontFace: SANS, fontSize: 24, bold: true, color: titleColor });
    s.addText(L_CARD, { x: cx, y: cy + 0.816, w: 2.52, h: 1.212, align: 'center',
      lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 11, color: textColor });
    box(s, cx + 0.435, cy + 2.255, 1.649, 0.453, btnFill, { shadow: SHADOW });
    s.addText('Learn More', { x: cx + 0.435, y: cy + 2.255, w: 1.649, h: 0.453, align: 'center',
      valign: 'middle', fontFace: SANS, fontSize: 11, bold: true, color: WHITE });
    s.addShape('line', { x: cx + 0.144, y: cy + 0.674, w: 2.232, h: 0, line: { color: titleColor, width: 1 } });
    shp(s, 'ellipse', { x: iconX, y: cy - 0.632, w: 0.44, h: 0.44, fill: { type: 'none' },
      line: { color: titleColor, width: 2 } });
  };
  card('chart', 'Statistic', WHITE, WHITE, NAVY, 5.814, 3.221, 6.83);
  card('award', 'Best View', ORANGE, ORANGE, ORANGE, 9.283, 3.281, 10.341);
}

/* --- 5 : BEST OF TOWER ---------------------------------------------- */
function slide05(s) {
  photos(s, [[0, 0, 3.44, 7.5], [9.889, 0, 3.444, 7.5]]);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'BEST OF TOWER', 4.598, 1.211, 4.138, 0.64, 32, 'center');
  [[4.125, 'Chat ', 'Client', 3.993, 1.752, 4.008, 4.104],
   [7.088, 'Target ', 'Market', 6.956, 2.206, 6.971, 7.067]].forEach(([bx, a, b, tx, tw, px, btn]) => {
    shp(s, 'rect', { x: bx, y: 2.528, w: 0.75, h: 0.75, fill: { type: 'none' }, line: { color: NAVY, width: 1.5 } });
    shp(s, 'ellipse', { x: bx + 0.24, y: 2.72, w: 0.27, h: 0.27, fill: { color: WHITE } });
    twoTone(s, a, b, tx, 3.571, tw);
    body(s, 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec sit amet  consec tetur adla dolor sit consec tetur cons',
      px, 4.103, 2.395, 1.489);
    box(s, btn, 5.88, 1.649, 0.453, NAVY, { shadow: SHADOW });
    s.addText('Check View', { x: btn, y: 5.88, w: 1.649, h: 0.453, align: 'center', valign: 'middle',
      fontFace: SANS, fontSize: 11, bold: true, color: ORANGE });
  });
}

/* --- 6 : TYPE BUILDING ---------------------------------------------- */
function slide06(s) {
  box(s, 0, 1.223, 5.768, 5.03, CYAN);
  [0, 2.525, 5.05].forEach(y => photos(s, [[1.326, y, 2.705, 2.45], [4.104, y, 2.705, 2.45]]));
  bigTriangles(s);
  frame(s);
  titleTag(s, 'TYPE BUILDING', 7.505, 1.329, 3.88, 0.64, 32);
  [['01 /', 'West Building', 7.446, 2.548, 8.471, 2.493],
   ['02 /', 'East Building', 7.473, 3.795, 8.497, 3.74],
   ['03 /', 'North Building', 7.499, 5.043, 8.524, 4.988]].forEach(([num, name, nx, ny, tx, ty]) => {
    s.addText(num, { x: nx, y: ny, w: 1.017, h: 0.64, wrap: false, fontFace: SANS, fontSize: 32, bold: true, color: WHITE });
    s.addText([
      { text: name, options: { fontSize: 20, bold: true, color: NAVY, breakLine: true } },
      { text: L1, options: { fontSize: 11, color: WHITE } }
    ], { x: tx, y: ty, w: 3.715, h: 1.161, align: 'justify', lineSpacingMultiple: 1.5, fontFace: SANS });
  });
}

/* --- 7 : THE STRUCTURE BUILD ---------------------------------------- */
function slide07(s) {
  photo(s, 0, 3.75, 13.333, 3.75);
  box(s, 1.726, 1.842, 9.863, 3.211, CYAN, { fill: { color: CYAN, transparency: 10 }, shadow: SHADOW });
  bigTriangles(s);
  frame(s);
  titleTag(s, 'THE STRUCTURE BUILD', 3.806, 1.449, 5.721, 0.64, 32, 'center');
  [['$ 758.686.857', 2.474], ['$ 874.635.904', 7.167]].forEach(([price, gx]) => {
    chev(s, gx - 0.063, 2.853, 0.253);
    s.addText([
      { text: price, options: { fontSize: 20, bold: true, color: NAVY, breakLine: true } },
      { text: 'Special of Build', options: { fontSize: 16, bold: true, italic: true, color: WHITE, charSpacing: 3, breakLine: true } },
      { text: 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec dolor sit amet ', options: { fontSize: 11, color: WHITE } }
    ], { x: gx + 0.429, y: 2.675, w: 3.213, h: 1.843, align: 'justify', lineSpacingMultiple: 1.5, fontFace: SANS });
  });
}

/* --- 8 : THREE BUILDING --------------------------------------------- */
function slide08(s) {
  box(s, 0, 2.483, 13.333, 2.533, NAVY);
  photos(s, [[6.667, 2.483, 1.889, 5.017], [8.69, 1.242, 1.889, 5.017], [10.714, 0, 1.889, 5.017]]);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'THREE BUILDING', 0.671, 1.319, 4.243, 0.64, 32);
  body(s, 'Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit', 5.401, 1.278, 2.841, 0.656);
  s.addText('\u201C' + L3 + '\u201D', { x: 0.582, y: 3.094, w: 5.354, h: 1.313, align: 'justify',
    lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 16, bold: true, italic: true, color: ORANGE });
  ring(s, 0.662, 5.532, 0.88, 0.88, NAVY, 0.043);
  s.addText('$56', { x: 0.662, y: 5.532, w: 0.88, h: 0.88, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 14, bold: true, color: WHITE });
  s.addText(L_DONUT, { x: 1.764, y: 5.532, w: 2.795, h: 0.841, fontFace: SANS, fontSize: 11,
    color: WHITE, lineSpacingMultiple: 2 });
  box(s, 4.694, 5.726, 1.294, 0.533, NAVY, { shadow: SHADOW });
  s.addText('Download', { x: 4.694, y: 5.726, w: 1.294, h: 0.533, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 11, bold: true, color: ORANGE });
}

/* --- 9 : BUILDING PHOTO --------------------------------------------- */
function slide09(s) {
  box(s, 0, 1.837, 13.333, 0.674, CYAN);
  photos(s, [[0, 3.75, 6.667, 3.75], [0.574, 1.145, 2.103, 2.059],
             [5.615, 1.145, 2.103, 2.059], [10.657, 1.145, 2.103, 2.059]]);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'BUILDING PHOTO', 7.441, 4.192, 4.395, 0.64, 32);
  [3.094, 8.136].forEach(x => {
    box(s, x, 1.145, 2.103, 2.059, NAVY, { shadow: SHADOW });
    s.addText(['\u201CLorem ipsum', 'amet  consec te', 'tur adla sssips', 'um dolor sit\u201D']
        .map((t, i) => ({ text: t, options: { breakLine: i < 3 } })),
      { x: x + 0.272, y: 1.753, w: 1.559, h: 0.841, align: 'center',
        fontFace: SANS, fontSize: 11, bold: true, italic: true, color: ORANGE });
    s.addShape('custGeom', { x: x + 1.765, y: 1.322, w: 0.183, h: 0.125, fill: { color: ORANGE },
      points: poly(0.183, 0.125, [[0.5, 1], [1, 0.13], [0.89, 0], [0.5, 0.66], [0.11, 0], [0, 0.13]]) });
  });
  body(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem um dolor sit amet  consec tetur lorem ipsum dolor sit  consec tetur sit amet ',
    7.783, 5.182, 4.27, 0.934);
  body(s, 'Lorem ipsum dolor sit amet  consec tetur ipsum dolor sit  consec tetur', 7.783, 6.189, 4.204, 0.656);
  box(s, 7.474, 5.281, 0.138, 0.736, NAVY);
  box(s, 7.474, 6.318, 0.138, 0.445, CYAN);
}

/* --- 10 : PREMIUM BUILD --------------------------------------------- */
function slide10(s) {
  photos(s, [[2.032, 1.147, 3.411, 2.656], [5.442, 3.786, 3.411, 2.656], [8.853, 1.147, 3.411, 2.656]]);
  bigTriangles(s);
  frame(s);
  s.addText('PREMIUM BUILD', { x: -0.814, y: 3.43, w: 4.278, h: 0.64, rotate: 270, align: 'center', wrap: false,
    fill: { color: NAVY }, fontFace: HEAD, fontSize: 32, bold: true, color: ORANGE, shadow: SHADOW });
  [['California', 2.032, 3.786], ['San Francisco', 5.442, 1.147], ['Las Vegas', 8.853, 3.786]].forEach(([name, cx, cy]) => {
    box(s, cx, cy, 3.411, 2.656, CYAN);
    box(s, cx + 0.347, cy + 0.433, 0.285, 0.285, NAVY);
    s.addText([
      { text: name, options: { fontSize: 20, bold: true, color: NAVY, breakLine: true } },
      { text: L1, options: { fontSize: 11, color: WHITE } }
    ], { x: cx + 0.787, y: cy + 0.128, w: 2.236, h: 1.885, align: 'justify', lineSpacingMultiple: 2, fontFace: SANS });
    box(s, cx + 0.895, cy + 2.111, 0.564, 0.101, NAVY);
  });
}

/* --- 11 : MAKE A BUILDING ------------------------------------------- */
function slide11(s) {
  box(s, 0, 3.75, 4.726, 3.75, NAVY);
  photos(s, [[0, 0.837, 2.966, 5.783], [3.081, 0.837, 2.966, 5.783]]);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'MAKE A BUILDING', 6.91, 1.41, 4.544, 0.64, 32);
  [[6.91, 2.595, NAVY, ORANGE], [9.749, 2.595, CYAN, WHITE],
   [6.91, 4.569, NAVY, ORANGE], [9.749, 4.569, CYAN, WHITE]].forEach(([x, y, dot, glyph]) => {
    shp(s, 'ellipse', { x: x + 0.048, y, w: 0.494, h: 0.494, fill: { color: dot } });
    shp(s, 'rect', { x: x + 0.197, y: y + 0.157, w: 0.197, h: 0.155, fill: { color: glyph } });
    s.addText(SERENITY, { x, y: y + 0.588, w: 2.313, h: 0.934, align: 'justify',
      lineSpacingMultiple: 1.5, fontFace: POPPINS, fontSize: 11, color: WHITE });
  });
}

/* --- 12 : BREAK SLIDE ------------------------------------------------ */
function slide12(s) {
  photo(s, 0.263, 0.232, 12.768, 7);
  bigTriangles(s);
  frame(s);
  s.addText('BREAK SLIDE', { x: 0, y: 2.468, w: 8.085, h: 1.447, wrap: false, align: 'center',
    fill: { color: NAVY }, fontFace: HEAD, fontSize: 80, bold: true, color: ORANGE, shadow: SHADOW });
  box(s, 0.263, 3.916, 7.822, 1.105, GREY_BG, { fill: { color: GREY_BG, transparency: 20 } });
  s.addText('Lorem ipsum dolor sit amet  consec tetur adlorem um dolor sit amet  consec tetur lorem ipsum dolor sit  consec tetur sit amet um dolor sit amet  consec ',
    { x: 0.844, y: 4.14, w: 6.661, h: 0.656, align: 'justify', lineSpacingMultiple: 1.5,
      fontFace: SANS, fontSize: 11, color: '262626' });
  box(s, 8.085, 2.468, 2.232, 2.552, ORANGE, { fill: { color: ORANGE, transparency: 20 } });
  chev(s, 8.92, 3.463, 0.563);
}

/* --- 13 : GREAT BUILDING -------------------------------------------- */
function slide13(s) {
  box(s, 0, 2.806, 9.4, 4.694, CYAN);
  box(s, 9.4, 1.842, 3.933, 3.832, NAVY);
  // phone mock-up (a raster image in the original)
  shp(s, 'roundRect', { x: 8.66, y: 0.68, w: 2.89, h: 6.08, rectRadius: 0.42, fill: { color: 'F1F2F4' },
    line: { color: 'D6D6D6', width: 1 } });
  photo(s, 8.82, 1.40, 2.58, 4.66);
  shp(s, 'roundRect', { x: 9.60, y: 1.04, w: 0.75, h: 0.055, rectRadius: 0.03, fill: { color: '3C3C3C' } });
  shp(s, 'ellipse', { x: 9.34, y: 1.02, w: 0.09, h: 0.09, fill: { color: '3C3C3C' } });
  shp(s, 'ellipse', { x: 9.86, y: 6.14, w: 0.48, h: 0.48, fill: { color: 'F1F2F4' },
    line: { color: 'C4C4C4', width: 1 } });
  bigTriangles(s);
  frame(s);
  titleTag(s, 'GREAT BUILDING', 0.882, 1.53, 4.267, 0.64, 32);
  s.addShape('line', { x: 1.057, y: 3.241, w: 0, h: 3.615, line: { color: NAVY, width: 1.5 } });
  [['45 ', 3.345], ['66 ', 4.36], ['78 ', 5.375], ['92 ', 6.391]].forEach(([n, y]) => {
    ring(s, 0.908, y, 0.321, 0.321, NAVY, 0.0515);
    twoTone(s, n, 'Month', 1.463, y - 0.059, 1.859, { charSpacing: 3 });
    body(s, L2, 3.535, y - 0.064, 4.365, 0.656);
  });
}

/* --- 14 : five photo cards ------------------------------------------ */
function slide14(s) {
  [1.133, 3.451, 5.769, 8.087, 10.405].forEach(x => photo(s, x, 1.724, 1.792, 2.252));
  bigTriangles(s);
  box(s, 0, 3.978, 13.333, 2.252, CYAN);
  ['#4758', '#7589', '#9085', '#6574', '#2357'].forEach((tag, i) => {
    const x = [1.133, 3.451, 5.769, 8.087, 10.405][i];
    s.addText(tag, { x: x + 0.115, y: 4.39, w: 1.561, h: 0.438, align: 'center', fill: { color: NAVY },
      fontFace: POPPINS, fontSize: 20, bold: true, color: ORANGE, shadow: SHADOW });
    s.addText('Lorem ipsum dolor sit amet, consectetur', { x, y: 5.066, w: 1.792, h: 0.631,
      align: 'center', lineSpacingMultiple: 1.5, fontFace: POPPINS_L, fontSize: 10.5, color: WHITE });
  });
  frame(s);
  box(s, 1.133, 5.934, 11.064, 0.803, NAVY, { shadow: SHADOW });
  s.addText('Lorem ipsum dolor sit amet, adipis sit dolor cons sit amet, adipis ectetur adipis sit amet, adi',
    { x: 1.41, y: 6.146, w: 10.527, h: 0.379, align: 'center', lineSpacingMultiple: 1.5,
      fontFace: SANS, fontSize: 11, italic: true, color: ORANGE, charSpacing: 3 });
}

/* --- 15 : TIME TO BUILDING ------------------------------------------ */
function slide15(s) {
  box(s, 0, 5.453, 13.333, 2.063, NAVY);
  photos(s, [[1.153, 3.752, 5.321, 3.748], [6.859, 3.75, 5.321, 3.75]]);
  bigTriangles(s);
  frame(s);
  titleTag(s, 'TIME TO BUILDING', 4.364, 1.17, 4.606, 0.64, 32, 'center');
  [['Fast Building | ', '$56.89', 1.147, 2.32, 1.631, 1.647, 3.191],
   ['Normal Building | ', '$33.75', 6.892, 2.368, 7.375, 7.392, 3.666]].forEach(([a, b, cx, cy, tx, hx, hw]) => {
    chev(s, cx, cy, 0.323);
    twoTone(s, a, b, hx, cy - 0.053, hw);
    body(s, L2S, tx, cy + 0.405, 4.365, 0.656);
  });
}

/* --- 16 : team / tool / target -------------------------------------- */
function slide16(s) {
  [0.983, 2.864, 4.745].forEach(y => photo(s, 0, y, 6.255, 1.811));
  bigTriangles(s);
  [[1.362, CYAN], [2.954, DKCYAN], [4.546, CYAN]].forEach(([y, c]) => box(s, 6.898, y, 5.562, 1.592, c));
  frame(s);
  box(s, 5.612, 1.362, 1.286, 4.776, NAVY, { shadow: SHADOW });
  [['Enginering ', 'Team', 1.601, 2.058, 2.542, 1.948],
   ['Tool ', 'Building', 3.192, 3.65, 2.022, 3.539],
   ['Target ', 'Finish', 4.784, 5.241, 2.016, 5.131]].forEach(([a, b, hy, ty, hw, iy]) => {
    box(s, 6.086, iy, 0.34, 0.34, ORANGE);
    twoTone(s, a, b, 7.256, hy, hw);
    body(s, L2S, 7.256, ty, 4.701, 0.656);
  });
}

/* --- 17 : BEST BUILDING --------------------------------------------- */
function slide17(s) {
  box(s, 0, 2.737, 6.895, 4.763, NAVY);
  photo(s, 1.558, 1.579, 3.505, 2.263);
  bigTriangles(s);
  frame(s);
  // laptop mock-up (a raster image in the original)
  shp(s, 'roundRect', { x: 1.38, y: 1.365, w: 3.86, h: 2.67, rectRadius: 0.07, fill: { color: 'E6E6E6' },
    line: { color: '9A9A9A', width: 1 } });
  photo(s, 1.58, 1.60, 3.46, 2.38);
  shp(s, 'trapezoid', { x: 0.78, y: 4.03, w: 5.11, h: 0.20, fill: { color: 'DCDCDC' } });
  shp(s, 'roundRect', { x: 2.95, y: 4.06, w: 0.72, h: 0.05, rectRadius: 0.025, fill: { color: 'B4B4B4' } });
  titleTag(s, 'BEST BUILDING', 6.895, 1.833, 5.719, 0.909, 48);
  s.addText('\u201C' + L3 + '\u201D', { x: 0.627, y: 4.785, w: 5.613, h: 1.313, align: 'center',
    lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 16, bold: true, italic: true, color: ORANGE });
  [['+98', 3.402], ['+75', 4.635], ['+45', 5.868]].forEach(([n, y]) => {
    ring(s, 7.437, y, 0.88, 0.88, NAVY, 0.043);
    s.addText(n, { x: 7.437, y, w: 0.88, h: 0.88, align: 'center', valign: 'middle',
      fontFace: SANS, fontSize: 14, bold: true, color: WHITE });
    s.addText('Lorem ipsum dolor sit amet, adipis sit dolor adipis sit dolor cons sitadipis sit dolor',
      { x: 8.539, y, w: 3.791, h: 0.841, fontFace: SANS, fontSize: 11, color: WHITE, lineSpacingMultiple: 2 });
  });
  box(s, 2.486, 6.377, 1.649, 0.453, ORANGE, { shadow: SHADOW });
  s.addText('Learn More', { x: 2.486, y: 6.377, w: 1.649, h: 0.453, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 11, bold: true, color: WHITE });
}

/* --- 18 : percentage bars ------------------------------------------- */
function slide18(s) {
  photo(s, 0, 0, 7.032, 7.5, [[0, 0], [0.794, 0], [1, 1], [0.206, 1]]);
  bigTriangles(s);
  frame(s);
  [['76%', 4.968, 1.326, NAVY, ORANGE], ['35%', 5.395, 3.058, CYAN, WHITE],
   ['88%', 5.822, 4.79, NAVY, ORANGE]].forEach(([pct, x, y, fill, color]) => {
    box(s, x, y, 6.364, 1.384, fill, { shadow: SHADOW });
    s.addText(pct, { x: x + 0.427, y: y + 0.339, w: 1.231, h: 0.707, wrap: false,
      fontFace: SANS, fontSize: 36, bold: true, color });
    s.addText('Lorem ipsum dolor sit amet  consec tetur adla sssipsum dolor sit amet  consec sit amet  consec tetur adla dolor sit amet ',
      { x: x + 1.905, y: y + 0.225, w: 4.033, h: 0.934, align: 'justify', lineSpacingMultiple: 1.5,
        fontFace: SANS, fontSize: 11, color });
  });
}

/* --- 19 : two doughnut charts --------------------------------------- */
function slide19(s, pres) {
  const donut = (x, colors, values) => s.addChart(pres.ChartType.doughnut,
    [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values }],
    { x, y: 2.663, w: 3.26, h: 2.173, holeSize: 75, chartColors: colors, showLegend: false,
      showValue: false, dataBorder: { pt: 0, color: WHITE }, firstSliceAng: 0 });
  donut(0.878, [NAVY, 'D9D9D9'], [8.2, 3.2]);
  donut(6.669, [YELLOW, GREY_BG], [8.2, 1]);
  s.addText('$32M', { x: 1.715, y: 3.464, w: 1.585, h: 0.572, align: 'center',
    fontFace: SANS, fontSize: 28, bold: true, color: NAVY });
  s.addText('$56M', { x: 7.506, y: 3.464, w: 1.585, h: 0.572, align: 'center',
    fontFace: SANS, fontSize: 28, bold: true, color: YELLOW });
  [['2016 SALES', '$10.000.000', 3.899, 2.606, 4.016, 3.098],
   ['2017 SALES', '$22.000.000', 3.906, 3.717, 4.022, 4.21],
   ['2018 SALES', '$10.000.000', 9.69, 2.606, 9.807, 3.098],
   ['2019 SALES', '$22.000.000', 9.697, 3.717, 9.813, 4.21]].forEach(([label, amount, lx, ly, ax, ay]) => {
    s.addText(label, { x: lx, y: ly, w: 2.071, h: 0.438, fontFace: SANS, fontSize: 20, color: '595959' });
    s.addText(amount, { x: ax, y: ay, w: 2.339, h: 0.404, align: 'center', fill: { color: ORANGE },
      fontFace: SANS, fontSize: 18, bold: true, color: WHITE, charSpacing: 3 });
  });
  s.addText('A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with serenity has taken possession of my entire soul, like. of spring which I enjoy with serenity has taken possession of.',
    { x: 1.749, y: 5.516, w: 10.474, h: 0.533, align: 'center', lineSpacingMultiple: 1.2,
      fontFace: SANS, fontSize: 11, color: '000000' });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 20 : six dashed-ring bullets ----------------------------------- */
function slide20(s) {
  const rows = [[2.897, YELLOW], [4.208, CYAN], [5.47, ORANGE]];
  [[1.971, 3.085, 0], [7.344, 8.457, -0.018]].forEach(([rx, tx, dy]) => {
    rows.forEach(([y0, color], i) => {
      const y = y0 + (i === 2 ? dy : 0), cx = rx + (i === 2 ? dy : 0);
      shp(s, 'ellipse', { x: cx, y, w: 0.958, h: 0.967, fill: { type: 'none' },
        line: { color, width: 1.5, dashType: 'dash' } });
      shp(s, 'ellipse', { x: cx + 0.079, y: y + 0.082, w: 0.785, h: 0.787, fill: { color } });
      shp(s, 'ellipse', { x: cx + 0.305, y: y + 0.317, w: 0.334, h: 0.334, fill: { type: 'none' },
        line: { color: WHITE, width: 2 } });
      s.addText(EIUSMOD, { x: tx + (i === 2 ? dy : 0), y: y + 0.152, w: 2.942, h: 0.978,
        align: 'justify', lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 12, color: 'A6A6A6' });
    });
  });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 21 : pinwheel of arrows ---------------------------------------- */
function slide21(s) {
  /* One arrow = a long shaft plus two short bars forming the head.  The
     four arrows are the same layout rotated by 90 degrees around (6.705, 4.697). */
  const CX = 6.705, CY = 4.697;
  const ARROW = [                       // [dx, dy, w, h, rotation]
    [ 0.788, -0.820, 0.531, 1.892,  43.22],   // shaft
    [ 0.672, -0.335, 1.140, 0.266,  -1.78],   // head bar "/"
    [ 0.311, -0.673, 1.140, 0.266,  88.22]];  // head bar "\\"
  [[CYAN, 21], [YELLOW, 20], [SLATE, 21], [ORANGE, 20]].forEach(([color, transparency], k) => {
    ARROW.forEach(([dx, dy, w, h, rot]) => {
      let px = dx, py = dy;
      for (let i = 0; i < k; i++) { const t = px; px = py; py = -t; }   // rotate -90 deg
      shp(s, 'roundRect', { x: CX + px - w / 2, y: CY + py - h / 2, w, h,
        rectRadius: Math.min(w, h) / 2, rotate: rot - 90 * k, fill: { color, transparency } });
    });
  });
  [[1.329, 3.579, 1.575, 3.111, YELLOW],
   [9.578, 3.579, 9.824, 3.111, CYAN],
   [1.389, 4.997, 1.635, 5.872, SLATE],
   [9.638, 4.997, 9.884, 5.872, ORANGE]].forEach(([tx, ty, bx, by, color]) => {
    s.addText(NULLAM, { x: tx, y: ty, w: 2.84, h: 0.75, bullet: { code: '2022', indent: 13.5 },
      lineSpacingMultiple: 1.2, fontFace: SANS, fontSize: 11, color: '222A35' });
    s.addText('More Info', { x: bx, y: by, w: 1.694, h: 0.319, align: 'center', valign: 'middle',
      shape: 'roundRect', rectRadius: 0.08, fill: { color }, fontFace: SANS, fontSize: 12, color: WHITE });
  });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 22 : avatar + three subtitles ---------------------------------- */
function slide22(s) {
  // avatar badge with yellow corner brackets
  shp(s, 'ellipse', { x: 3.023, y: 2.414, w: 1.314, h: 1.314, fill: { color: NAVY } });
  shp(s, 'ellipse', { x: 3.49, y: 2.6, w: 0.38, h: 0.38, fill: { color: '595959' } });
  shp(s, 'ellipse', { x: 3.25, y: 3.06, w: 0.86, h: 0.75, fill: { color: '595959' } });
  [[2.693, 2.083, 1, 1], [4.304, 2.083, -1, 1], [2.693, 3.694, 1, -1], [4.304, 3.694, -1, -1]]
    .forEach(([x, y, sx, sy]) => {
      const L = 0.363, T = 0.075;
      box(s, sx > 0 ? x : x + L - T, y, T, L, YELLOW);
      box(s, x, sy > 0 ? y : y + L - T, L, T, YELLOW);
    });
  s.addText('Subtitle Here', { x: 2.267, y: 4.624, w: 2.825, h: 0.303, align: 'center', margin: 0,
    fontFace: SANS, fontSize: 18, bold: true, color: '181717' });
  s.addText('PLACEHOLDER',
    { x: 2.267, y: 5.167, w: 2.825, h: 1.02, align: 'center', margin: 0, lineSpacingMultiple: 1.5,
      fontFace: SANS, fontSize: 14, color: '181717' });
  [2.14, 3.643, 5.147].forEach(y => {
    s.addText('Subtitle Here', { x: 6.642, y, w: 2.825, h: 0.303, margin: 0,
      fontFace: SANS, fontSize: 18, bold: true, color: '181717' });
    s.addText(PODCAST, { x: 6.642, y: y + 0.543, w: 5.125, h: 0.577, margin: 0,
      lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 12, color: '181717' });
  });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 23 : four year pockets ----------------------------------------- */
function slide23(s) {
  /* card body: rounded top corners, semicircular notch bitten out of the top */
  const pocket = (w, h) => [
    { x: 0, y: h }, { x: 0.135 * w, y: h },
    { curve: { type: 'cubic', x1: 0.135 * w, y1: 0.881 * h, x2: 0.299 * w, y2: 0.784 * h }, x: 0.5 * w, y: 0.784 * h },
    { curve: { type: 'cubic', x1: 0.701 * w, y1: 0.784 * h, x2: 0.865 * w, y2: 0.881 * h }, x: 0.865 * w, y: h },
    { x: w, y: h }, { x: w, y: 0.099 * h },
    { curve: { type: 'cubic', x1: w, y1: 0.044 * h, x2: 0.925 * w, y2: 0 }, x: 0.833 * w, y: 0 },
    { x: 0.167 * w, y: 0 },
    { curve: { type: 'cubic', x1: 0.075 * w, y1: 0, x2: 0, y2: 0.044 * h }, x: 0, y: 0.099 * h },
    { close: true }];
  const W = 2.478, H = 3.146;
  [['2017', 1.204, SKY, CYAN], ['2018', 4.098, SAND, YELLOW],
   ['2019', 6.992, SAND, YELLOW], ['2020', 9.886, SKY, CYAN]]
    .forEach(([year, x, tabColor, bodyColor]) => {
      shp(s, 'round2SameRect', { x: x + 0.235, y: 2.242, w: 2.008, h: 1.765,
        rectRadius: 0.55, fill: { color: tabColor } });
      s.addShape('custGeom', { x, y: 3.005, w: W, h: H, flipV: true,
        fill: { color: bodyColor }, points: pocket(W, H) });
      s.addText(year, { x, y: 2.557, w: 2.464, h: 0.505, align: 'center',
        fontFace: SANS, fontSize: 24, color: WHITE });
      s.addText('Keyword Here', { x: x + 0.579, y: 4.097, w: 1.315, h: 0.303, align: 'center', wrap: false,
        fontFace: SANS, fontSize: 12, color: WHITE });
      s.addText('Lorem ipsum dolor sit amet, lacus nulla', { x: x + 0.418, y: 4.519, w: 1.638, h: 0.707,
        align: 'center', fontFace: SANS, fontSize: 12, color: WHITE });
    });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 24 : data table ------------------------------------------------- */
function slide24(s) {
  const head = ['The Styles', 'Style', 'Style', 'Style'];
  const data = [['One', '672', '-', '987'], ['Two', '-', '782', '-'], ['Three', '-', '438', '363'],
                ['Four', '564', '-', '-'], ['Five', '-', '236', '152'], ['Six', '-', '22', '321']];
  const border = [{ type: 'solid', color: WHITE, pt: 3 }, { type: 'solid', color: WHITE, pt: 3 },
                  { type: 'solid', color: WHITE, pt: 3 }, { type: 'solid', color: WHITE, pt: 3 }];
  const cell = (text, fill, color) => ({ text, options: { fill: { color }, color, align: 'center',
    valign: 'middle', fontFace: SANS, fontSize: 12, margin: 0.02, border } });
  const rows = [head.map((t, i) => cell(t, 0, 0)).map((c, i) => ({
      text: head[i], options: { fill: { color: i === 0 ? YELLOW : GREY_BG }, color: '111111',
        align: 'center', valign: 'middle', fontFace: SANS, fontSize: 12, border } }))];
  data.forEach((r, ri) => rows.push(r.map((t, ci) => ({
    text: t,
    options: { fill: { color: ci === 0 ? YELLOW : (ri % 2 ? GREY_BG : WHITE) },
      color: ci === 0 ? '151314' : '24282B', align: 'center', valign: 'middle',
      fontFace: SANS, fontSize: 12, border } }))));
  s.addTable(rows, { x: 1.562, y: 2.392, colW: [2.552, 2.552, 2.552, 2.552], rowH: 0.445 });
  s.addText([
    { text: 'Cras mattis consectetur purus sit amet fermentum. Sed posuere consectetur est at lobortis. Nullam consectetur est at lobortis. Nullam Sed posuere consectetur est at Sed posuere fermentum. Sed posuere consectetur est at lobortis. Nullam ', options: { breakLine: true } },
    { text: 'consectetur est at lobortis. Nullam Sed posuere consectetur' }
  ], { x: 2.021, y: 5.721, w: 9.534, h: 0.882, align: 'center', valign: 'middle',
       lineSpacingMultiple: 1.5, fontFace: SANS, fontSize: 10, italic: true, color: '222A35' });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 25 : four peaks ------------------------------------------------- */
function slide25(s) {
  [['23%', 'Insert text', 1.074, 2.542, 1.428, 2.915, NAVY, '003B68', NAVY],
   ['75%', 'Insert text', 2.568, 2.291, 1.425, 3.172, CYAN, DKCYAN, NAVY],
   ['35%', 'Insert text', 4.057, 3.446, 1.430, 2.013, ORANGE, DKORANGE, NAVY],
   ['27%', 'Insert text', 5.551, 3.115, 1.428, 2.344, YELLOW, DKYELLOW, YELLOW]]
    .forEach(([pct, label, x, y, w, h, light, dark, labelColor], i) => {
      shp(s, 'rtTriangle', { x: x + w / 2, y, w: w / 2, h, fill: { color: light } });
      shp(s, 'rtTriangle', { x, y, w: w / 2, h, flipH: true, fill: { color: dark } });
      s.addText(pct, { x: x + w / 2 - 0.37, y: 4.867, w: 0.745, h: 0.415, align: 'center', wrap: false,
        fontFace: SANS, fontSize: 18.7, bold: true, color: WHITE });
      s.addText(label, { x: [1.064, 2.569, 4.073, 5.578][i], y: 5.548, w: 1.42, h: 0.37, align: 'center',
        wrap: false, fontFace: SANS, fontSize: 16, bold: true, color: labelColor });
    });
  s.addShape('line', { x: 0.829, y: 5.46, w: 6.286, h: 0, line: { color: '808080', width: 4 } });
  [0.829, 7.115].forEach(x => shp(s, 'ellipse', { x: x - 0.075, y: 5.385, w: 0.15, h: 0.15, fill: { color: '808080' } }));
  [[7.752, 2.994], [10.237, 2.994], [7.752, 4.099], [10.237, 4.099]].forEach(([x, y]) =>
    s.addText(QUIS, { x, y, w: 2.333, h: 0.903, bullet: { code: '2022', indent: 13.5 }, lineSpacingMultiple: 1.5,
      fontFace: SANS, fontSize: 11, color: '222A35' }));
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 26 : four halo circles ------------------------------------------ */
function slide26(s) {
  [[2.351, YELLOW], [4.856, ORANGE], [7.363, NAVY], [9.877, NAVY]].forEach(([x, color], i) => {
    shp(s, 'ellipse', { x, y: 2.824, w: 1.405, h: 1.405, fill: { color, transparency: 80 } });
    shp(s, 'ellipse', { x: x + 0.101, y: 2.925, w: 1.203, h: 1.203, fill: { color, transparency: 65 } });
    shp(s, 'ellipse', { x: x + 0.207, y: 3.031, w: 0.992, h: 0.992, fill: { color } });
    shp(s, 'roundRect', { x: x + 0.45, y: 3.29, w: 0.5, h: 0.36, rectRadius: 0.05,
      fill: { type: 'none' }, line: { color: WHITE, width: 1.5 } });
    box(s, x + 0.63, 3.65, 0.14, 0.09, WHITE);      // monitor stand
    box(s, x + 0.55, 3.74, 0.30, 0.05, WHITE);      // monitor foot
  });
  [3.756, 6.26, 8.767].forEach(x =>
    s.addShape('line', { x, y: 3.52, w: 1.1, h: 0, line: { color: 'D9D9D9', width: 1 } }));
  [[1.963, YELLOW], [4.471, ORANGE], [7.008, CYAN], [9.522, CYAN]].forEach(([x, color]) =>
    s.addText([
      { text: 'TITLE GOES HERE', options: { fontSize: 16, bold: true, color, breakLine: true } },
      { text: '', options: { breakLine: true } },
      { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', options: { fontSize: 11, color: '000000' } }
    ], { x, y: 4.575, w: 2.12, h: 1.229, align: 'center', valign: 'middle', margin: 0, fontFace: SANS }));
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 27 : line chart -------------------------------------------------- */
function slide27(s, pres) {
  const labels = ['Jan', 'Feb', 'Mar', 'April', 'May', 'Jun'];
  s.addChart(pres.ChartType.line, [
    { name: 'Name 1', labels, values: [105, 125, 165, 225, 465, 299] },
    { name: 'Name 2', labels, values: [500, 420, 410, 405, 200, 575] },
    { name: 'Name 3', labels, values: [300, 150, 400, 700, 600, 500] },
    { name: 'Column2', labels, values: [null, null, null, null, null, null] }
  ], { x: 1.452, y: 2.138, w: 5.688, h: 4.37, chartColors: [ORANGE, CYAN, YELLOW, CYAN],
       lineSize: 3, lineDataSymbol: 'circle', lineDataSymbolSize: 5, lineSmooth: false,
       showLegend: true, legendPos: 'b', legendFontFace: SANS, legendFontSize: 10,
       catAxisLabelFontFace: SANS, catAxisLabelFontSize: 10, valAxisLabelFontFace: SANS, valAxisLabelFontSize: 10,
       valGridLine: { color: GREY_BG, size: 1 }, catGridLine: { style: 'none' },
       chartArea: { border: { pt: 1, color: '000000' }, roundedCorners: false },
       plotArea: { border: { pt: 0.5, color: GREY_BG } } });
  [['Graph 1', 2.325, 2.762], ['Graph 2', 3.911, 4.334], ['Graph 3', 5.326, 5.696]].forEach(([t, hy, by]) => {
    s.addText(t, { x: 7.853, y: hy, w: 2.486, h: 0.37, fontFace: SANS, fontSize: 16, bold: true, color: '51647E' });
    s.addText(UTWISI, { x: 7.898, y: by, w: 3.939, h: 0.809, fontFace: SANS, fontSize: 12, color: '222A35' });
  });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 28 : ribbon light bulb ------------------------------------------ */
function slide28(s) {
  /* the bulb is a stack of flat quadrilateral "ribbons"; each entry is
     [x, y, w, h, colour, corner points as fractions of w/h]              */
  const RIBBONS = [
    [8.640, 2.270, 1.517, 1.643, ORANGE,   [[0.579, 0], [0, 0.537], [0, 1], [1, 0]]],
    [9.519, 2.270, 1.708, 1.080, 'D1D1D1', [[0, 0], [1, 1], [0.809, 0.351], [0.374, 0]]],
    [8.839, 2.649, 2.388, 1.879, CYAN,     [[0.863, 0], [0, 0.913], [0.038, 1], [1, 0.373]]],
    [8.640, 3.152, 2.587, 0.761, '909090', [[0, 0], [1, 0.427], [0.99, 0.891], [0, 1]]],
    [9.103, 3.477, 2.124, 1.548, YELLOW,   [[1, 0], [0, 0.885], [0.023, 1], [0.987, 0.228]]],
    [8.839, 4.364, 1.923, 0.769, CYAN,     [[0, 0], [1, 0.534], [0.942, 1], [0.047, 0.213]]],
    [9.103, 4.846, 1.454, 0.678, 'BFBFBF', [[0, 0], [1, 0.847], [0.946, 1], [0.034, 0.263]]],
    [9.296, 4.774, 1.466, 0.750, YELLOW,   [[0.063, 1], [0.924, 0.478], [1, 0], [0, 0.881]]],
    [9.296, 5.421, 1.261, 0.103, '595959', [[0, 0.138], [1, 0], [0.938, 1], [0.073, 1]]]];
  RIBBONS.forEach(([x, y, w, h, color, pts]) =>
    s.addShape('custGeom', { x, y, w, h, fill: { color }, points: poly(w, h, pts) }));
  // screw base: five stacked rounded bars
  [[9.372, 5.524, 1.107, 0.105], [9.372, 5.679, 1.107, 0.107], [9.372, 5.833, 1.107, 0.107],
   [9.372, 5.973, 1.095, 0.065], [9.407, 6.031, 1.030, 0.154], [9.575, 6.101, 0.696, 0.182]]
    .forEach(([x, y, w, h]) => shp(s, 'roundRect', { x, y, w, h, rectRadius: h / 2, fill: { color: NAVY } }));

  [['SHOW MORE', 1.979, YELLOW], ['SHOW MORE', 3.853, DKNAVY]].forEach(([t, x, fill]) =>
    s.addText(t, { x, y: 2.649, w: 1.573, h: 0.37, wrap: false, fill: { color: fill },
      fontFace: SANS, fontSize: 16, color: WHITE }));
  [3.258, 4.403].forEach(y =>
    s.addText(IPSUM_DUMMY, { x: 1.912, y, w: 5.633, h: 1.145, lineSpacing: 20.2,
      fontFace: SANS, fontSize: 11, color: '404040' }));
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* --- 29 : three half-moon years -------------------------------------- */
function slide29(s) {
  const moon = (w, h) => [
    { x: w, y: 0 },
    { curve: { type: 'cubic', x1: w, y1: 0.552 * h, x2: 0.776 * w, y2: h }, x: 0.5 * w, y: h },
    { curve: { type: 'cubic', x1: 0.224 * w, y1: h, x2: 0, y2: 0.552 * h }, x: 0, y: 0 },
    { x: 0.5 * w, y: 0 }, { close: true }];
  [['2018', 2.641, CYAN, 2.657, 2.36], ['2019', 5.581, YELLOW, 5.571, 5.274],
   ['2020', 8.344, YELLOW, 8.377, 8.08]].forEach(([year, gx, color, tx, hx], i) => {
    const gy = i === 2 ? 2.533 : 2.528;
    s.addShape('custGeom', { x: gx - 0.244, y: gy + 0.594, w: 2.239, h: 1.119,
      fill: { color }, points: moon(2.239, 1.119) });
    shp(s, 'ellipse', { x: gx + 0.281, y: gy, w: 1.189, h: 1.189, fill: { color: WHITE },
      shadow: { type: 'outer', blur: 40, offset: 21, angle: 92, color: '000000', opacity: 0.39 } });
    s.addText(year, { x: gx + 0.281, y: gy + 0.342, w: 1.189, h: 0.505, align: 'center',
      fontFace: SANS, fontSize: 24, color: '262626' });
    s.addText('Your Text Here', { x: hx, y: 4.633, w: 2.806, h: 0.406, align: 'center',
      fontFace: SANS, fontSize: 18, color: '262626' });
    s.addText([
      { text: 'Ut wisi enim ad minim veniam' },
      { text: ', quis nostrud exerci tation' }
    ], { x: tx, y: 5.039, w: 2.212, h: 0.89, align: 'center', lineSpacingMultiple: 1.2,
         fontFace: SANS, fontSize: 12, color: '808080' });
  });
  frame(s, 'BFBFBF');
  infoTitle(s);
}

/* ===================================================================== */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pres.layout = 'W16x9';
  pres.author = 'pptxgenjs';
  pres.title = 'NEW BUILDING';

  const orange = () => { const s = pres.addSlide(); s.background = { color: ORANGE }; return s; };
  const grey = () => { const s = pres.addSlide(); s.background = { color: GREY_BG }; return s; };

  coverSlide(orange(), 'CONSTRUCTION', 8.066);
  slide02(orange());
  slide03(orange());
  slide04(orange());
  slide05(orange());
  slide06(orange());
  slide07(orange());
  slide08(orange());
  slide09(orange());
  slide10(orange());
  slide11(orange());
  slide12(orange());
  slide13(orange());
  slide14(orange());
  slide15(orange());
  slide16(orange());
  slide17(orange());
  slide18(orange());
  slide19(grey(), pres);
  slide20(grey());
  slide21(grey());
  slide22(grey());
  slide23(grey());
  slide24(grey());
  slide25(grey());
  slide26(grey());
  slide27(grey(), pres);
  slide28(grey());
  slide29(grey());
  coverSlide(orange(), 'THANK FOR YOU', 8.376);

  return pres.writeFile({ fileName: path.join(__dirname, '0ca9b9a5-ef2c-4a6f-9355-09baf609f2c9_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
