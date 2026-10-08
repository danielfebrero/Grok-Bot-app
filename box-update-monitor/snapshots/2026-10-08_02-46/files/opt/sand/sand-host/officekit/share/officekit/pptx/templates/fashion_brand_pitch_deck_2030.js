/**
 * LuxeLane 2030 Pitchdeck - 30 slides, 10 x 5.625 in (16:9).
 *
 * Recreated with pptxgenjs only.  The deck uses a single two-colour palette
 * (deep navy on cream), the Epilogue type family, and a repeating chrome of
 * three 6pt captions along the top edge plus a small ring-and-dot glyph.
 *
 * The original file contains empty picture placeholders (no embedded raster
 * data anywhere in the package); they are reproduced here as flat cream
 * rectangles by `photo()`, drawn first so the navy artwork sits on top.
 *
 *   node <thisfile>.js   ->  writes the .pptx next to this script
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const NAVY = '062540'; // dk1 / accent1 - panels, headings
const CREAM = 'FAF3E6'; // lt1 - slide background, reversed-out text

const SEMI = 'Epilogue SemiBold';
const BOOK = 'Epilogue';

/* -------------------------------------------------------------- text styles */

const TITLE = { size: 30, face: SEMI, color: NAVY };
const TITLE_SM = { size: 27, face: SEMI, color: NAVY };
const TITLE_LG = { size: 39, face: SEMI, color: NAVY };
const SUBTITLE = { size: 12, face: SEMI, color: NAVY };
const PRICE = { size: 12, face: BOOK, color: NAVY };
const LABEL = { size: 11, face: SEMI, color: NAVY };
const LABEL_ON_NAVY = { size: 11, face: SEMI, color: CREAM };
const BODY = { size: 9, face: BOOK, color: NAVY };
const BODY_ON_NAVY = { size: 9, face: BOOK, color: CREAM };
const CAPTION = { size: 6, face: BOOK, color: NAVY };
const CAPTION_BOLD = { size: 6, face: SEMI, color: NAVY };
const CAPTION_ON_NAVY = { size: 6, face: SEMI, color: CREAM };
const CAPTION_LIGHT_ON_NAVY = { size: 6, face: BOOK, color: CREAM };

/* ------------------------------------------------------------------ helpers */

// Every text frame in the source deck shares these body properties:
// top-anchored, left-aligned, 5.4pt / 2.7pt insets, "resize shape to fit text".
const TEXT_FRAME = {
  align: 'left',
  valign: 'top',
  margin: [5.4, 5.4, 2.7, 2.7], // [left, right, bottom, top] in points
  fit: 'resize',
  wrap: true,
  isTextBox: true,
};

function txt(slide, x, y, w, h, text, style) {
  slide.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: style.face, fontSize: style.size, color: style.color,
    ...TEXT_FRAME,
  });
}

// Solid navy block used behind captions and as the deck's graphic accent.
function panel(slide, x, y, w, h) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: NAVY } });
}

// Image placeholder.  The source deck ships with no media at all - every
// picture frame is an unfilled `pic` placeholder painted in the background
// cream - so each one is reproduced as a flat cream block laid down before
// the navy artwork, exactly as the original renders.
function photo(slide, x, y, w, h) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: CREAM } });
}

// Hairline divider used in the "Usage of Funds" cost table.
function rule(slide, x, y, w) {
  slide.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: NAVY, width: 0.75 } });
}

// Page glyph: a hairline navy ring with a filled navy dot inside it.
const DOT_SIZE = 0.1103;
const DOT_CORE = 0.0665;

function dot(slide, x, y) {
  slide.addShape('ellipse', {
    x: x, y: y, w: DOT_SIZE, h: DOT_SIZE,
    fill: { type: 'none' }, line: { color: NAVY, width: 0.25 },
  });
  const inset = (DOT_SIZE - DOT_CORE) / 2;
  slide.addShape('ellipse', {
    x: x + inset, y: y + inset, w: DOT_CORE, h: DOT_CORE,
    fill: { color: NAVY },
  });
}

// The running header repeated verbatim on all 30 slides.
const CHROME = [
  { x: 0.4437, w: 0.5275, text: 'LuxeLane' },
  { x: 2.5938, w: 0.3474, text: '2030' },
  { x: 7.4062, w: 0.5643, text: 'Pitchdeck' },
];

function chrome(slide) {
  CHROME.forEach((c) => txt(slide, c.x, 0.1983, c.w, 0.1767, c.text, CAPTION));
}

/* ------------------------------------------------------------------- slides */

// 1. LuxeLane
function slide1(s) {
  photo(s, 1.1081, 1.1969, 6.0825, 2.3506);
  panel(s, 5.775, 1.0314, 3.7875, 1.5897);
  txt(s, 0.4479, 4.1234, 2.2536, 0.5806, "LuxeLane", TITLE);
  txt(s, 0.4562, 4.6515, 1.9286, 0.2777, "Fashion company", SUBTITLE);
  txt(s, 0.4464, 1.1969, 0.4486, 0.2272, "2030", BODY);
  txt(s, 7.4063, 1.2703, 1.5312, 0.2524, "Fashion Product", LABEL_ON_NAVY);
  txt(s, 7.4062, 1.5229, 0.9043, 0.2272, "Sold Out!", BODY_ON_NAVY);
  txt(s, 7.4062, 1.8909, 1.5312, 0.2524, "Fashion Product", LABEL_ON_NAVY);
  txt(s, 7.4063, 2.1435, 0.9102, 0.2272, "Sold Out!", BODY_ON_NAVY);
  txt(s, 5.1337, 4.5296, 2.0581, 0.3787, "The Best Company to look for Fashion.", BODY);
  dot(s, 9.4375, 4.7424);
  chrome(s);
}

// 2. Vision and Missions
function slide2(s) {
  photo(s, 0, 0.8787, 2.3, 2.715);
  photo(s, 7.4062, 0.8787, 2.1562, 2.715);
  panel(s, 7.3187, 0.7958, 1.7347, 2.6417);
  txt(s, 2.5938, 2.5082, 2.6103, 1.0855, "Vision and Missions", TITLE);
  txt(s, 5.0071, 0.8787, 1.4216, 0.2524, "Our Vision", LABEL);
  txt(s, 5.0071, 1.1414, 2.1562, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue.", BODY);
  txt(s, 0.4437, 4.2501, 1.6779, 0.2524, "Our Missions", LABEL);
  txt(s, 0.4437, 4.5278, 2.8219, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 5.0062, 4.5278, 2.8219, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  dot(s, 9.4375, 4.9225);
  panel(s, 0, 0.7958, 1.9083, 0.8757);
  chrome(s);
}

// 3. Problems Encounter
function slide3(s) {
  photo(s, 5, 0.8037, 5, 2.971);
  txt(s, 0.4453, 1.3717, 2.4093, 1.0855, "Problems Encounter", TITLE);
  txt(s, 0.4453, 4.165, 2.625, 0.2524, "Economic Instability", LABEL);
  txt(s, 5.0062, 4.165, 2.9688, 0.2524, "Consumer Preferences", LABEL);
  txt(s, 0.4437, 4.4427, 3.1979, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 5.0062, 4.4427, 3.1979, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  panel(s, 0.5609, 2.6224, 0.9684, 0.326);
  txt(s, 0.6332, 2.6485, 0.8301, 0.2777, "Based on Data on 03.03.30", CAPTION_LIGHT_ON_NAVY);
  dot(s, 9.4375, 4.8213);
  panel(s, 4.934, 0.7443, 4.8772, 2.8257);
  chrome(s);
}

// 4. Solutions Offered
function slide4(s) {
  photo(s, 2.6012, 2.4361, 4.5687, 3.1889);
  photo(s, 7.5076, 1.1004, 1.8475, 1.1101);
  panel(s, 2.4825, 2.8984, 4.5687, 2.7266);
  panel(s, 0.1224, 3.545, 2.8401, 1.2525);
  txt(s, 2.6012, 1.082, 2.3988, 1.0855, "Solutions Offered", TITLE);
  txt(s, 7.4138, 3.8199, 1.6564, 0.2524, "Solutions 03", LABEL);
  txt(s, 0.4375, 3.7893, 1.2602, 0.2608, "Solutions 01", LABEL_ON_NAVY);
  txt(s, 0.4437, 4.067, 1.7143, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY_ON_NAVY);
  txt(s, 7.4138, 4.0976, 1.7143, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  txt(s, 7.4138, 2.6208, 1.6479, 0.2524, "Solutions 02", LABEL);
  txt(s, 7.4138, 2.8984, 1.7143, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  txt(s, 0.4437, 1.5529, 1.7143, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  dot(s, 9.4375, 5.125);
  chrome(s);
}

// 5. Our Products
function slide5(s) {
  photo(s, 5, 1.545, 1.8475, 2.1237);
  photo(s, 7.4062, 1.545, 1.8475, 2.1237);
  panel(s, 7.3275, 2.2338, 1.7063, 1.8475);
  panel(s, 4.9212, 2.2338, 1.7063, 1.8475);
  txt(s, 0.445, 2.271, 3.0218, 0.5806, "Our Products", TITLE);
  txt(s, 4.9635, 3.7544, 1.229, 0.2608, "Products 01", LABEL_ON_NAVY);
  txt(s, 7.3724, 3.7544, 1.229, 0.2608, "Products 02", LABEL_ON_NAVY);
  txt(s, 0.4437, 2.9779, 2.7906, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  dot(s, 9.4375, 5.125);
  chrome(s);
}

// 6. Unique Value Proposition
function slide6(s) {
  photo(s, 0, 0.825, 2.3925, 4.425);
  photo(s, 7.8863, 3.5738, 1.6762, 1.6762);
  panel(s, 7.7394, 3.2523, 1.6762, 1.6904);
  txt(s, 2.6037, 1.0864, 2.7075, 0.9845, "Unique Value Proposition", TITLE_SM);
  txt(s, 2.6037, 2.7014, 1.8557, 0.2524, "Gain Creators", LABEL);
  txt(s, 2.6037, 3.8746, 2.5947, 0.2524, "Products & Services", LABEL);
  txt(s, 2.6037, 2.9791, 2.7947, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 2.6037, 4.1523, 2.7947, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 7.8191, 3.3325, 0.8352, 0.1767, "Our best-selling", CAPTION_ON_NAVY);
  dot(s, 9.4375, 0.825);
  chrome(s);
}

// 7. Our Product Prototypes
function slide7(s) {
  photo(s, 2.5938, 3.3447, 2.2891, 1.3594);
  photo(s, 0.4375, 3.3447, 2.0156, 1.9053);
  panel(s, 2.5938, 3.6224, 2.1289, 1.3594);
  txt(s, 5.0075, 0.8435, 2.8247, 1.0855, "Our Product Prototypes", TITLE);
  txt(s, 0.4453, 2.2066, 2.3523, 0.2524, "More Information", LABEL);
  txt(s, 0.4487, 2.4843, 2.8247, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 5, 2.4843, 2.8247, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 5, 3.3447, 3.4587, 0.2524, "Why We Use That Material?", LABEL);
  txt(s, 5.0034, 3.6224, 2.8247, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 2.6245, 4.7546, 1.1915, 0.1767, "Prototype Ripstop Nylon", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.8429);
  chrome(s);
}

// 8. Case Study
function slide8(s) {
  photo(s, 5.0804, 1.845, 2.3988, 3.1559);
  photo(s, 0.5325, 3.335, 2.1488, 1.6659);
  panel(s, 5, 2.02, 2.3988, 3.23);
  txt(s, 4.9325, 0.8866, 3.8787, 0.5806, "Case Study", TITLE);
  txt(s, 0.445, 0.898, 1.7321, 0.2524, "Study Case 01", LABEL);
  txt(s, 0.445, 2.204, 1.7617, 0.2524, "Study Case 02", LABEL);
  txt(s, 0.445, 1.1757, 2.6225, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  txt(s, 0.445, 2.4817, 2.6225, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.", BODY);
  panel(s, 0.445, 3.51, 2.1488, 1.74);
  txt(s, 0.456, 5.0366, 0.4959, 0.1767, "Type-02", CAPTION_ON_NAVY);
  txt(s, 5.0101, 5.0077, 0.4959, 0.1767, "Type-00", CAPTION_ON_NAVY);
  txt(s, 2.0794, 5.0366, 0.5577, 0.1767, "$80.00", CAPTION_BOLD);
  txt(s, 6.8568, 5.0077, 0.542, 0.1767, "$76.00", CAPTION_BOLD);
  dot(s, 9.4375, 5.125);
  chrome(s);
}

// 9. Target Market
function slide9(s) {
  photo(s, 5.0862, 1.0505, 3.8356, 3.5241);
  txt(s, 0.4375, 1.6805, 1.9426, 1.0855, "Target Market", TITLE);
  panel(s, 5, 1.253, 3.6953, 3.66);
  txt(s, 0.4427, 3.4208, 2.0185, 0.2524, "Target Market", LABEL);
  txt(s, 2.599, 3.4208, 2.0185, 0.2524, "Target Market", LABEL);
  txt(s, 0.4427, 3.7134, 2.1562, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 2.599, 3.7134, 2.1562, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 5.014, 4.6694, 1.0074, 0.1767, "Our Target Markting", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.7228);
  chrome(s);
}

// 10. Market Size
function slide10(s) {
  photo(s, 0.52, 0.7629, 2.2456, 2.7296);
  photo(s, 5.0825, 1.8904, 2.4062, 1.6021);
  txt(s, 0.445, 4.1118, 2.6851, 0.5806, "Market Size", TITLE);
  txt(s, 5, 1.24, 2.4888, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 5, 0.9548, 1.6354, 0.2524, "Market Size", LABEL);
  txt(s, 5, 4.4653, 2.0613, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 5, 4.1801, 1.6354, 0.2524, "Market Size", LABEL);
  panel(s, 5, 1.9967, 2.4062, 1.7611);
  panel(s, 0.4375, 1.0442, 2.0737, 2.7136);
  txt(s, 0.4528, 3.5368, 0.5643, 0.1767, "Market 01", CAPTION_ON_NAVY);
  txt(s, 5.0083, 3.5368, 0.5775, 0.1767, "Market 02", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.8851);
  chrome(s);
}

// 11. Market Trends
function slide11(s) {
  photo(s, 0.505, 1.6881, 2.4172, 1.3884);
  photo(s, 5.3204, 1.6881, 2.9268, 1.3884);
  txt(s, 3.1794, 1.8657, 1.7415, 1.0855, "Market Trends", TITLE);
  txt(s, 5.0075, 3.5716, 2.3523, 0.2524, "More Information", LABEL);
  txt(s, 5.0075, 3.8567, 2.2944, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 0.445, 3.5716, 2.1999, 0.2524, "About this trend", LABEL);
  txt(s, 0.445, 3.8567, 2.2944, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  dot(s, 9.4375, 4.2765);
  panel(s, 5.2562, 1.6281, 2.9268, 1.3533);
  chrome(s);
}

// 12. Competition From Other Firms
function slide12(s) {
  photo(s, 5.075, 1.1072, 1.9688, 2.1684);
  txt(s, 0.445, 1.1072, 4.2675, 1.0855, "Competition From Other Firms", TITLE);
  panel(s, 4.9531, 1.6931, 2.0231, 1.885);
  txt(s, 5.0075, 4.005, 0.9298, 0.2524, "Firms 1", LABEL);
  txt(s, 5.0075, 4.2902, 2.1488, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  txt(s, 7.4062, 1.2272, 0.9588, 0.2524, "Firms 2", LABEL);
  txt(s, 7.4062, 1.5124, 1.9688, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 0.4375, 3.048, 2.0814, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 4.9997, 3.3385, 0.4407, 0.1767, "Firm 01", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.5585);
  chrome(s);
}

// 13. Competitive Advantage
function slide13(s) {
  photo(s, 3.3427, 1.2172, 3.8573, 1.8675);
  txt(s, 0.4375, 3.8517, 4.3021, 1.0855, "Competitive Advantage", TITLE);
  txt(s, 5.0075, 3.9501, 1.7577, 0.2524, "Advantage 02", LABEL);
  txt(s, 5.0075, 4.2353, 2.055, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 7.4188, 1.2172, 2.3523, 0.2524, "More Information", LABEL);
  txt(s, 7.4188, 1.5024, 2.1437, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 0.445, 1.2172, 1.728, 0.2524, "Advantage 01", LABEL);
  txt(s, 0.445, 1.5024, 2.055, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  panel(s, 3.2625, 1.3897, 3.7425, 1.9628);
  txt(s, 3.2654, 3.1302, 0.5577, 0.1767, "Our Team", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.6551);
  chrome(s);
}

// 14. Business Model
function slide14(s) {
  photo(s, 0.5288, 2.7558, 1.6612, 2.2026);
  photo(s, 4.945, 1.2009, 2.1488, 3.7575);
  txt(s, 0.445, 1.2084, 3.8224, 1.0855, "Business Model", TITLE);
  txt(s, 7.4138, 4.1431, 2.3523, 0.2524, "More Information", LABEL);
  txt(s, 7.4138, 4.4283, 2.3988, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 2.6012, 4.1431, 1.1694, 0.2524, "Model 01", LABEL);
  txt(s, 2.6012, 4.4283, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 7.4062, 1.2009, 1.199, 0.2524, "Model 02", LABEL);
  txt(s, 7.4062, 1.4861, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  dot(s, 9.4375, 2.8125);
  panel(s, 4.8864, 1.1431, 1.6612, 3.3863);
  chrome(s);
}

// 15. Monetization Strategy
function slide15(s) {
  photo(s, 0.4375, 2.8389, 1.8566, 1.9552);
  photo(s, 6.565, 1.1797, 2.69, 3.6144);
  panel(s, 0.3792, 2.7747, 1.7353, 1.9552);
  txt(s, 2.6012, 1.1797, 3.2816, 1.0855, "Monetization Strategy", TITLE);
  txt(s, 2.6012, 4.0115, 1.5566, 0.2524, "Strategy 02", LABEL);
  txt(s, 2.6012, 4.2967, 2.5975, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 2.5938, 2.8389, 1.527, 0.2524, "Strategy 01", LABEL);
  txt(s, 2.5938, 3.1241, 2.69, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  dot(s, 0.4498, 1.2681);
  panel(s, 6.5094, 1.1219, 2.4823, 3.3234);
  chrome(s);
}

// 16. Marketing Strategy
function slide16(s) {
  photo(s, 0.75, 0.9813, 3.8984, 3.9592);
  txt(s, 5.015, 2.3984, 2.8366, 1.0855, "Marketing Strategy", TITLE);
  txt(s, 5.0075, 4.2767, 1.5566, 0.2524, "Strategy 02", LABEL);
  txt(s, 5.0075, 4.5619, 2.3988, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  txt(s, 5.0075, 0.9813, 1.527, 0.2524, "Strategy 01", LABEL);
  txt(s, 5.0075, 1.2665, 2.3988, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY);
  dot(s, 9.4375, 4.8302);
  panel(s, 0.8878, 1.4082, 3.8305, 3.6151);
  chrome(s);
}

// 17. Product Development
function slide17(s) {
  photo(s, 0.4978, 3.0553, 1.3259, 1.3275);
  photo(s, 2.6553, 3.0553, 1.3259, 1.3275);
  photo(s, 4.81, 3.0553, 1.3259, 1.3275);
  photo(s, 6.9647, 3.0553, 1.3259, 1.3275);
  txt(s, 0.447, 1.3366, 3.5342, 1.0855, "Product Development", TITLE);
  panel(s, 0.4375, 3.2363, 1.3275, 1.4212);
  panel(s, 2.5938, 3.2363, 1.3275, 1.4212);
  panel(s, 4.7513, 3.2363, 1.3275, 1.4212);
  panel(s, 6.9088, 3.2363, 1.3275, 1.4212);
  txt(s, 7.4138, 1.5278, 1.9719, 0.2524, "About product", LABEL);
  txt(s, 7.4138, 1.7942, 2.0587, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 0.4258, 4.4318, 0.6932, 0.1767, "Prototype 01", CAPTION_ON_NAVY);
  txt(s, 2.5775, 4.4318, 0.7063, 0.1767, "Prototype 02", CAPTION_ON_NAVY);
  txt(s, 4.7383, 4.4318, 0.4302, 0.1767, "Unit 01", CAPTION_ON_NAVY);
  txt(s, 6.8897, 4.4318, 0.4434, 0.1767, "Unit 02", CAPTION_ON_NAVY);
  dot(s, 9.4375, 4.5472);
  chrome(s);
}

// 18. Go-To-Market Strategy
function slide18(s) {
  photo(s, 0.525, 2.9725, 4.2368, 2.0652);
  txt(s, 5.0102, 3.9272, 3.6617, 1.0855, "Go-To-Market Strategy", TITLE);
  txt(s, 0.4375, 2.0049, 1.527, 0.2524, "Strategy 01", LABEL);
  txt(s, 0.4375, 2.2901, 3.0859, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 5.0102, 2.0049, 1.5651, 0.2524, "Strategy 03", LABEL);
  txt(s, 5.0102, 2.2901, 3.3805, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 5.0102, 1.0178, 1.5566, 0.2524, "Strategy 02", LABEL);
  txt(s, 5.0102, 1.303, 3.3805, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  dot(s, 9.4375, 4.9024);
  panel(s, 0.4625, 2.9083, 4.0792, 1.855);
  chrome(s);
}

// 19. Here’s Some Our Achievement
function slide19(s) {
  photo(s, 1.5735, 1.9182, 4.185, 2.7825);
  txt(s, 0.4375, 0.6015, 4.2538, 1.0855, "Here’s Some Our Achievement", TITLE);
  panel(s, 1.486, 2.0234, 3.9853, 3.0307);
  txt(s, 7.4104, 4.2388, 1.2572, 0.2524, "Tokyo - 2010", LABEL);
  txt(s, 7.4104, 4.524, 2.0865, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 7.4104, 3.0785, 1.5202, 0.2524, "Hungary - 2009", LABEL);
  txt(s, 7.4104, 3.3637, 2.0865, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 7.4104, 1.9182, 1.4636, 0.2524, "Orlando - 2029", LABEL);
  txt(s, 7.4104, 2.2034, 2.0865, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  txt(s, 1.5032, 4.7909, 1.0153, 0.1767, "New York City - 2028", CAPTION_ON_NAVY);
  dot(s, 6.0485, 4.9391);
  chrome(s);
}

// 20. Core Team
function slide20(s) {
  photo(s, 2.8512, 1.0275, 1.3887, 1.4175);
  photo(s, 7.4062, 1.0275, 1.3887, 1.4175);
  photo(s, 7.4062, 3.1275, 1.3887, 1.4175);
  txt(s, 0.445, 3.5025, 2.7675, 0.5806, "Core Team", TITLE);
  panel(s, 5, 3.5025, 3.6088, 1.1865);
  panel(s, 0.445, 1.4025, 3.6088, 1.1865);
  panel(s, 5, 1.4025, 3.6088, 1.1865);
  txt(s, 5.2076, 3.6902, 1.2976, 0.2608, "Pevita Nubis", LABEL_ON_NAVY);
  txt(s, 5.2077, 3.9754, 1.9486, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY_ON_NAVY);
  txt(s, 0.6451, 1.586, 1.502, 0.2608, "Azusa Murphy", LABEL_ON_NAVY);
  txt(s, 0.6452, 1.8712, 1.9486, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY_ON_NAVY);
  txt(s, 5.2288, 1.586, 1.2976, 0.2608, "Lisa Nguban", LABEL_ON_NAVY);
  txt(s, 5.2288, 1.8712, 1.9486, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit.", BODY_ON_NAVY);
  txt(s, 0.445, 4.09, 1.9486, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.", BODY);
  dot(s, 9.4375, 4.5786);
  chrome(s);
}

// 21. Advisor and Partner
function slide21(s) {
  photo(s, 1.0187, 0.8868, 3.15, 3.9985);
  txt(s, 5.0075, 0.8868, 4.4066, 1.0855, "Advisor and Partner", TITLE);
  txt(s, 5.0075, 2.75, 2.9077, 0.2524, "Arip Srikan As Advisor", LABEL);
  txt(s, 5.0075, 3.0352, 2.4625, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 5.0075, 4.07, 2.558, 0.2524, "Srikandi As Partner", LABEL);
  txt(s, 5.0075, 4.3552, 2.4625, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 9.4375, 4.7734);
  panel(s, 0.9462, 1.1939, 2.5589, 3.7661);
  chrome(s);
}

// 22. Funding Plans
function slide22(s) {
  photo(s, 3.8338, 0.8887, 2.7437, 3.9985);
  photo(s, 7.4737, 3.5597, 1.3275, 1.3275);
  panel(s, 3.7474, 1.4857, 1.85, 3.4859);
  txt(s, 0.445, 1.0117, 2.8988, 1.0855, "Funding Plans", TITLE);
  txt(s, 7.4138, 0.902, 1.0557, 0.2524, "Plan 02", LABEL);
  txt(s, 7.4138, 1.1872, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 7.4138, 2.2314, 1.0642, 0.2524, "Plan 03", LABEL);
  txt(s, 7.4138, 2.5166, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 0.445, 4.0719, 1.0261, 0.2524, "Plan 01", LABEL);
  txt(s, 0.445, 4.3571, 2.5569, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 0.543, 2.5166);
  chrome(s);
}

// 23. Usage of Funds
function slide23(s) {
  photo(s, 0.4375, 0.9075, 4.2275, 1.905);
  txt(s, 5.0075, 0.9075, 4.115, 1.0855, "Usage of Funds", TITLE);
  txt(s, 5.0075, 2.3034, 3.0025, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 0.445, 3.3571, 1.19, 0.2524, "Usage 1", LABEL);
  txt(s, 0.445, 3.9237, 1.19, 0.2524, "Usage 2", LABEL);
  txt(s, 0.445, 4.4903, 1.19, 0.2524, "Usage 3", LABEL);
  txt(s, 7.415, 3.3571, 0.9965, 0.2777, "$400.00", PRICE);
  txt(s, 7.415, 3.9237, 0.9965, 0.2777, "$400.00", PRICE);
  panel(s, 0.5112, 1.1921, 4.2275, 1.7022);
  txt(s, 7.415, 4.4903, 0.9965, 0.2777, "$400.00", PRICE);
  rule(s, 0.49, 3.6648, 7.8407);
  rule(s, 0.49, 4.2423, 7.8407);
  rule(s, 0.49, 4.798, 7.8407);
  dot(s, 9.4375, 4.574);
  chrome(s);
}

// 24. Financial Outlook
function slide24(s) {
  photo(s, 2.8756, 0.9637, 1.9994, 4.0188);
  photo(s, 5.1244, 0.9637, 1.9607, 2.7941);
  txt(s, 7.4141, 0.9637, 2.3281, 1.0855, "Financial Outlook", TITLE);
  txt(s, 5.0075, 4.1672, 1.9607, 0.2524, "Information 03", LABEL);
  txt(s, 5.0075, 4.4524, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 0.445, 1.2593, 1.9226, 0.2524, "Information 01", LABEL);
  txt(s, 0.445, 1.5445, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  panel(s, 2.8095, 1.5624, 1.85, 3.4859);
  txt(s, 0.445, 4.1672, 1.9523, 0.2524, "Information 02", LABEL);
  txt(s, 0.445, 4.4524, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 9.4375, 4.8345);
  chrome(s);
}

// 25. Methods of Measuring Success
function slide25(s) {
  photo(s, 0.4525, 3.3523, 9.095, 1.8977);
  txt(s, 2.6014, 0.7097, 6.4925, 1.0855, "Methods of Measuring Success", TITLE);
  txt(s, 2.6091, 2.0465, 1.3298, 0.2524, "Method 01", LABEL);
  txt(s, 2.6091, 2.3317, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 7.1716, 2.0465, 1.3679, 0.2524, "Method 03", LABEL);
  txt(s, 7.1716, 2.3317, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 4.8903, 2.0465, 1.3594, 0.2524, "Method 02", LABEL);
  txt(s, 4.8903, 2.3317, 2.1488, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 0.4525, 2.3099);
  panel(s, 0.3925, 3.2807, 6.6465, 1.8977);
  chrome(s);
}

// 26. Exit Strategies
function slide26(s) {
  photo(s, 5.1275, 2.9575, 2.1488, 2.03);
  photo(s, 7.53, 1.17, 2.0325, 3.8175);
  panel(s, 5.1963, 2.875, 2.1487, 2.07);
  txt(s, 0.445, 1.17, 4.9066, 0.7321, "Exit Strategies", TITLE_LG);
  txt(s, 0.445, 4.3237, 1.1305, 0.2524, "Strat 02", LABEL);
  txt(s, 0.445, 4.6088, 2.1488, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit..", BODY);
  txt(s, 2.6012, 4.3237, 1.139, 0.2524, "Strat 03", LABEL);
  txt(s, 2.6012, 4.6088, 2.1488, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit..", BODY);
  txt(s, 2.6012, 2.9346, 1.1009, 0.2524, "Strat 01", LABEL);
  txt(s, 2.6012, 3.2198, 2.1488, 0.3787, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit..", BODY);
  dot(s, 7.1659, 1.17);
  chrome(s);
}

// 27. Testimonials
function slide27(s) {
  photo(s, 0.445, 2.5175, 3.89, 2.3138);
  txt(s, 5.0075, 1.0151, 4.5521, 0.5806, "Testimonials", TITLE);
  txt(s, 0.445, 1.0965, 1.6716, 0.2524, "Testimony 01", LABEL);
  txt(s, 0.445, 1.3817, 2.615, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  panel(s, 0.675, 2.7263, 3.7437, 2.1912);
  txt(s, 5, 2.4739, 1.7013, 0.2524, "Testimony 02", LABEL);
  txt(s, 5, 2.7591, 2.615, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  txt(s, 5, 4.016, 1.7098, 0.2524, "Testimony 03", LABEL);
  txt(s, 5, 4.3012, 2.615, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 9.4375, 4.721);
  chrome(s);
}

// 28. Time For a Quick Stretch
function slide28(s) {
  photo(s, 6.7916, 2.8125, 2.7631, 1.987);
  photo(s, 0.445, 2.8125, 1.7175, 1.987);
  panel(s, 0.6663, 2.7375, 1.5629, 1.8154);
  txt(s, 5.0075, 1.1775, 2.5246, 0.2524, "Did You Know That?", LABEL);
  txt(s, 5, 1.4627, 2.615, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 0.4453, 1.2486);
  panel(s, 6.9413, 2.7469, 2.6753, 1.8154);
  txt(s, 2.6012, 2.8125, 3.6488, 1.0855, [{ text: "Time For a" }, { text: " ", options: { bold: true } }, { text: "Quick Stretch" }], TITLE);
  chrome(s);
}

// 29. Let’s Take a Breath For Now…
function slide29(s) {
  photo(s, 0.445, 1.8258, 5.9118, 2.0726);
  panel(s, 0.3835, 1.8974, 5.5265, 2.0726);
  txt(s, 5.0153, 4.1977, 0.8536, 0.2777, "Funfact!", SUBTITLE);
  txt(s, 5.0078, 4.4829, 2.615, 0.5301, "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porpoasdp porttitor.", BODY);
  dot(s, 0.445, 4.9579);
  txt(s, 5.0078, 0.881, 3.4375, 1.5904, "Let’s Take a Breath For Now…", TITLE);
  chrome(s);
}

// 30. Thank You For Your Attention!
function slide30(s) {
  photo(s, 0.445, 1.3438, 5.4613, 2.2969);
  panel(s, 0.6063, 1.8458, 5.3667, 1.8677);
  txt(s, 0.445, 4.1267, 1.8988, 0.2524, "Phone Number", LABEL);
  txt(s, 0.4375, 4.4119, 1.8047, 0.2272, "+000 222 444 555", BODY);
  txt(s, 4.9997, 4.1267, 1.6977, 0.2524, "Mail Address", LABEL);
  txt(s, 4.9997, 4.4119, 1.6675, 0.2272, "luxelane@mail.com", BODY);
  txt(s, 7.4072, 4.1267, 2.0684, 0.2524, "Website Address", LABEL);
  txt(s, 7.4072, 4.4119, 2.0575, 0.2272, "www.luxelane.com/shop", BODY);
  dot(s, 9.4375, 1.273);
  txt(s, 5, 0.7946, 3.9222, 1.5904, "Thank You For Your Attention!", TITLE);
  chrome(s);
}


/* -------------------------------------------------------------------- build */

const SLIDES = [
  slide1,
  slide2,
  slide3,
  slide4,
  slide5,
  slide6,
  slide7,
  slide8,
  slide9,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
  slide27,
  slide28,
  slide29,
  slide30,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
pptx.layout = 'DECK';
pptx.author = 'LuxeLane';
pptx.title = 'LuxeLane 2030 Pitchdeck';

SLIDES.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: CREAM };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0b33e0fd-e725-4534-a079-33853c5b0ae0_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f));
