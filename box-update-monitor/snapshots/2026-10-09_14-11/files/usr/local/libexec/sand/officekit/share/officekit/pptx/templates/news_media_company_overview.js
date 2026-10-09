/**
 * "Fozza News" template deck — rebuilt with pptxgenjs.
 *
 * Run:  node 05af0e24-8e37-4289-a436-a9b486aef094_grok_final.js
 * Writes 05af0e24-8e37-4289-a436-a9b486aef094_grok_final.pptx next to this file.
 *
 * Slide size 13.333 x 7.5 in (16:9). All coordinates below are inches.
 * Photographs in the original are replaced by programmatic "[image]" placeholders.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const GREEN = '97C67E';   // theme accent1
const CREAM = 'FAF6F5';   // slide background (theme accent3)
const BONE = 'F4F3EE';    // theme accent4
const WHITE = 'FFFFFF';
const GRAY = '595959';    // body copy (tx1 @ 65% lum)
const INK = '171E13';     // logo mark outline
const LOGO = '3D5231';    // logo wordmark
const HEAD = 'Lora';      // major font
const BODY = 'Lato Light';

// ---------------------------------------------------------------- primitives
function newSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  return s;
}

/**
 * Plain text box: no insets, top anchored, 90% line spacing —
 * the defaults every text frame in the source deck inherits from its master.
 */
function text(s, content, x, y, w, h, opts) {
  s.addText(content, Object.assign({
    x: x, y: y, w: w, h: h,
    margin: 0, valign: 'top', wrap: true, isTextBox: true, lineSpacingMultiple: 0.9,
  }, opts || {}));
}

/** Body copy: 10pt Lato Light italic gray, 1.5 line spacing. */
function para(s, content, x, y, w, h, opts) {
  text(s, content, x, y, w, h, Object.assign({
    fontSize: 10, fontFace: BODY, italic: true, color: GRAY, lineSpacingMultiple: 1.5,
  }, opts || {}));
}

/** Sub-heading: 16pt bold green. */
function lead(s, content, x, y, w, h, opts) {
  text(s, content, x, y, w, h, Object.assign({
    fontSize: 16, bold: true, color: GREEN, lineSpacingMultiple: 1.5,
  }, opts || {}));
}

/** Two stacked Lora display lines, as used for every section title. */
function title(s, line1, line2, x, y, w, opts) {
  const o = Object.assign({ fontSize: 44, color: '000000' }, opts || {});
  const big = o.fontSize >= 90;
  const h = o.gap || (big ? 1.025 : 0.636);
  const gap = o.gap || (big ? 1.026 : 0.569);
  const style = { fontSize: o.fontSize, fontFace: HEAD, bold: true, color: o.color };
  text(s, line1, x, y, w, h, style);
  text(s, line2, x, y + gap, w, h, style);
}

/** Green rule (a straight connector in the source deck); hairlines default to 0.5pt. */
function rule(s, x, y, w, h, width) {
  s.addShape('line', { x: x, y: y, w: w, h: h, line: { color: GREEN, width: width || 0.5 } });
}

/** Bulleted list where every paragraph carries the deck's 0.25" hanging bullet. */
function bulletList(s, lines, x, y, w, h) {
  const items = lines.map(function (line, i) {
    return {
      text: line,
      options: {
        fontSize: 10, fontFace: BODY, italic: true, color: GRAY,
        lineSpacingMultiple: 1, paraSpaceBefore: 10, bullet: { code: '2022', indent: 18 },
        breakLine: i < lines.length - 1,
      },
    };
  });
  text(s, items, x, y, w, h);
}

// ---------------------------------------------------------------- chrome
/** "Fozza News" logo mark: two overlapping rounded rectangles + wordmark. */
function logo(s) {
  s.addShape('roundRect', {
    x: 0.426, y: 0.399, w: 0.309, h: 0.17, rectRadius: 0.085,
    fill: { type: 'none' }, line: { color: INK, width: 1.5 },
  });
  s.addShape('roundRect', {
    x: 0.495, y: 0.468, w: 0.309, h: 0.17, rectRadius: 0.085, rotate: 90,
    fill: { type: 'none' }, line: { color: INK, width: 1.5 },
  });
  text(s, 'Fozza', 0.834, 0.41, 0.421, 0.129,
    { fontSize: 11, fontFace: HEAD, bold: true, color: LOGO });
  text(s, 'News', 0.834, 0.548, 0.421, 0.149,
    { fontSize: 11, fontFace: HEAD, bold: true, color: LOGO });
}

/** Top-right "hamburger" mark: six short green dashes in a 2 x 3 grid. */
function menuMark(s) {
  [12.555, 12.765].forEach(function (x) {
    [0.416, 0.554, 0.698].forEach(function (y) {
      s.addShape('line', { x: x, y: y, w: 0.143, h: 0, line: { color: GREEN, width: 4.5 } });
    });
  });
}

/** Vertical "Follow Us" label under a short tick. */
function followUs(s, x, y) {
  rule(s, x, y, 0, 0.314, 1.5);
  text(s, 'Follow Us', x - 0.462, y + 0.755, 0.925, 0.388,
    { fontSize: 14, color: GRAY, rotate: 90, lineSpacingMultiple: 1.5 });
}

/** Four progress dots (2nd filled). `horizontal` lays them out left-to-right. */
function dots(s, x, y, horizontal) {
  for (var i = 0; i < 4; i++) {
    const filled = i === 1;
    s.addShape('ellipse', {
      x: horizontal ? x + i * 0.548 : x,
      y: horizontal ? y : y + i * 0.548,
      w: 0.351, h: 0.351,
      fill: filled ? { color: GREEN } : { type: 'none' },
      line: filled ? null : { color: GREEN, width: 2 },
    });
  }
}

/** Everything that repeats on (nearly) every slide. */
function chrome(s, o) {
  logo(s);
  menuMark(s);
  if (o.followUs) followUs(s, o.followUs[0], o.followUs[1]);
  if (o.dots) dots(s, o.dots[0], o.dots[1], false);
  if (o.dotsRow) dots(s, o.dotsRow[0], o.dotsRow[1], true);
  if (o.page) {
    text(s, o.page, 0.439, 6.214, 0.395, 0.276,
      { fontSize: 14, color: GRAY, lineSpacingMultiple: 1.5 });
  }
  if (o.footer) {
    text(s, 'Template Presentaion', 0.439, 6.936, 2.036, 0.149,
      { fontSize: 12, fontFace: HEAD, bold: true, color: o.footer });
  }
}

/** Green circle badge with a single white letter (SWOT slide). */
function badge(s, letter, x, y, d) {
  s.addText(letter, {
    x: x, y: y, w: d, h: d, shape: 'ellipse', fill: { color: GREEN },
    align: 'center', valign: 'middle', bold: true, fontSize: 18, color: WHITE, margin: 0,
  });
}

// ---------------------------------------------------------------- placeholders
/**
 * Stand-in for one of the deck's picture placeholders: the framed picture glyph
 * (fixed 0.873 x 0.667 in, centred) plus an "[image]" caption, drawn with native
 * shapes over the exact area the photo would occupy. No raster data is embedded.
 */
function imageBox(s, x, y, w, h) {
  const iw = 0.873, ih = 0.667;
  const ix = x + (w - iw) / 2, iy = y + (h - ih) / 2;
  s.addShape('rect', { x: ix, y: iy, w: iw, h: ih, fill: { color: WHITE }, line: { color: '7F7F7F', width: 1 } });
  s.addShape('rect', {
    x: ix + 0.04, y: iy + 0.04, w: iw - 0.08, h: ih - 0.08,
    fill: { type: 'none' }, line: { color: 'C0C0C0', width: 0.75 },
  });
  s.addShape('ellipse', { x: ix + 0.13, y: iy + 0.1, w: 0.15, h: 0.15, fill: { color: 'F8CE7C' } });
  s.addShape('triangle', { x: ix + 0.18, y: iy + 0.29, w: 0.58, h: 0.32, fill: { color: '83B9E8' } });
  text(s, '[image]', x - 1, y + h / 2 - 0.16, w + 2, 0.32,
    { fontSize: 18, color: '111111', align: 'center' });
}

/**
 * Laptop mockup — a raster image in the source deck, redrawn from native shapes.
 * Fractions are measured off that image; the source crops 5% from its bottom,
 * so vertical fractions are scaled by 1/0.95.
 */
function laptopMockup(s, x, y, w, h) {
  const v = function (f) { return y + h * f / 0.9496; };
  s.addShape('roundRect', {
    x: x + w * 0.102, y: v(0.007), w: w * 0.796, h: h * (0.891 - 0.007) / 0.9496,
    rectRadius: 0.05, fill: { color: '17171A' },
  });
  s.addShape('rect', {
    x: x + w * 0.126, y: v(0.06), w: w * 0.749, h: h * (0.834 - 0.06) / 0.9496,
    fill: { color: WHITE },
  });
  s.addShape('ellipse', {
    x: x + w * 0.496, y: v(0.028), w: w * 0.008, h: w * 0.008 * (w / h), fill: { color: '3C3C3C' },
  });
  s.addShape('roundRect', {
    x: x + w * 0.008, y: v(0.892), w: w * 0.983, h: h * 0.036 / 0.9496,
    rectRadius: 0.02, fill: { color: 'E8E8E8' }, line: { color: 'CFCFCF', width: 0.5 },
  });
  s.addShape('trapezoid', {
    x: x + w * 0.021, y: v(0.928), w: w * 0.958, h: h * 0.031 / 0.9496,
    flipV: true, fill: { color: '8C8E90' },
  });
  s.addShape('roundRect', {
    x: x + w * 0.45, y: v(0.9), w: w * 0.1, h: h * 0.018 / 0.9496,
    rectRadius: 0.02, fill: { color: 'CCCCCC' },
  });
}

/** Phone mockup — a raster image in the source deck, redrawn from native shapes. */
function phoneMockup(s, x, y, w, h) {
  s.addShape('roundRect', {
    x: x + w * 0.107, y: y + h * 0.046, w: w * 0.763, h: h * 0.886,
    rectRadius: 0.3, fill: { color: '121214' },
  });
  s.addShape('rect', {
    x: x + w * 0.171, y: y + h * 0.176, w: w * 0.645, h: h * 0.632, fill: { color: WHITE },
  });
  s.addShape('ellipse', {
    x: x + w * 0.535, y: y + h * 0.108, w: w * 0.035, h: w * 0.035 * (w / h),
    fill: { color: '4A4A4C' },
  });
  s.addShape('roundRect', {
    x: x + w * 0.41, y: y + h * 0.116, w: w * 0.1, h: h * 0.008,
    rectRadius: 0.02, fill: { color: '4A4A4C' },
  });
  s.addShape('ellipse', {
    x: x + w * 0.428, y: y + h * 0.842, w: w * 0.12, h: w * 0.12 * (w / h),
    fill: { type: 'none' }, line: { color: '4A4A4C', width: 1 },
  });
}

/**
 * Canada map: each entry is one region outline, flattened [x0,y0,x1,y1,...]
 * in hundredths of an inch relative to the map's top-left corner.
 */
const MAP_ORIGIN = [6.228, 1.613];
const MAP_REGIONS = [
  [288,57,282,57,286,54,283,44,270,48,271,59,266,54,262,59,276,79,285,71,288,57],
  [287,109,276,95,266,104,281,113,287,109],
  [233,54,226,45,220,48,223,39,205,42,196,51,199,58,208,58,197,58,195,65,217,76,193,72,224,83,233,54],
  [263,19,253,7,244,29,260,29,263,19],
  [352,61,345,49,356,48,365,55,352,61],
  [350,34,338,24,315,29,310,17,291,12,294,20,302,21,306,42,341,42,350,34],
  [264,90,252,79,254,55,245,47,240,66,236,50,224,83,194,73,194,80,204,86,205,97,233,99,237,94,252,105,264,90],
  [456,154,448,137,434,133,424,120,452,128,453,105,412,96,417,88,408,86,411,79,392,72,387,77,390,70,374,69,367,59,349,68,343,49,331,55,328,71,333,79,325,63,331,49,321,49,313,61,311,78,322,84,315,82,314,87,337,96,332,96,334,109,325,127,321,109,315,103,310,115,310,102,300,99,305,94,294,73,311,52,298,45,290,52,290,73,282,90,283,97,295,102,284,127,283,115,277,111,269,115,275,121,252,117,237,99,223,103,225,107,237,105,226,108,222,121,216,105,192,99,201,95,199,88,178,69,167,88,206,142,244,159,234,206,281,212,306,172,323,162,334,141,329,137,339,140,355,129,348,113,352,104,337,96,361,95,361,89,376,98,373,106,381,101,397,116,391,133,397,139,377,143,376,156,401,150,425,169,451,169,448,160,430,152,456,154],
  [575,241,563,228,548,225,541,232,547,221,523,224,506,215,506,207,468,178,504,239,480,242,476,252,490,271,513,276,511,258,522,268,565,248,569,254,575,241],
  [392,110,375,110,374,122,383,121,392,110],
  [175,206,185,195,125,172,87,244,107,288,106,308,139,319,181,204,175,206],
  [114,38,106,20,19,96,25,106,103,160,105,149,95,146,91,122,97,108,95,91,104,79,100,74,107,66,99,57,114,38],
  [233,206,185,195,182,203,194,206,181,205,139,319,213,336,228,239,224,242,222,236,230,230,233,206],
  [124,171,24,106,24,124,40,124,34,164,38,181,17,201,18,208,31,205,18,215,24,224,17,229,17,239,23,244,9,237,24,281,29,283,32,278,27,252,36,276,106,308,107,288,87,244,124,171],
  [9,196,15,192,7,184,1,197,3,213,9,196],
  [547,336,531,328,529,319,501,328,496,338,506,337,522,358,545,348,547,336],
  [574,334,547,337,543,347,549,346,535,362,539,369,547,369,553,354,574,334],
  [631,285,630,275,625,278,622,270,619,278,619,266,614,272,606,261,598,268,585,263,582,270,581,246,571,252,574,302,602,289,609,290,609,298,616,281,622,291,631,285],
  [317,250,293,248,291,228,280,225,281,212,234,206,213,336,265,342,269,297,317,250],
  [245,300,245,290,239,288,244,282,254,307,245,300],
  [582,321,569,313,568,331,580,329,582,321],
  [549,294,522,294,538,299,549,294],
  [206,142,167,88,178,67,165,66,166,56,158,61,157,45,149,51,146,44,115,39,99,57,107,66,95,91,97,108,91,122,95,146,105,150,103,160,234,206,244,159,206,142],
  [168,101,160,104,172,109,167,116,161,112,152,117,155,111,141,111,155,108,142,97,168,101],
  [194,163,164,175,150,168,164,166,163,153,176,168,194,163],
  [456,376,453,370,421,372,407,361,396,309,388,310,372,293,368,267,348,267,319,250,269,296,266,338,271,347,313,357,327,345,339,347,347,356,355,356,364,374,397,377,404,384,404,392,388,385,390,405,382,427,419,411,414,403,436,396,456,376],
  [565,248,522,268,511,258,513,276,498,275,477,257,476,245,500,241,502,232,491,224,489,210,468,180,468,197,461,210,446,206,438,183,424,183,408,173,381,180,391,207,388,224,411,249,388,271,402,298,401,308,396,305,404,354,421,372,453,370,457,377,481,372,490,364,501,328,531,311,524,303,513,306,481,346,503,301,554,279,569,254,565,248],
  [220,33,208,17,195,13,174,36,177,50,189,52,221,37,220,33],
  [366,164,339,149,338,141,326,170,335,170,338,177,350,164,361,170,366,164],
  [245,20,235,4,222,14,234,23,228,25,235,32,243,30,245,20],
  [289,19,266,12,270,25,279,27,278,35,285,35,289,19],
  [356,154,347,137,334,133,324,120,352,128,353,105,312,96,317,88,308,86,311,79,292,72,286,77,290,70,274,69,267,59,249,68,243,49,230,55,228,71,233,79,225,63,231,49,221,49,213,61,211,78,222,84,215,82,214,87,237,96,232,96,233,109,225,127,221,109,215,103,210,115,210,102,200,99,205,94,194,73,211,52,198,45,190,52,190,73,182,90,183,97,195,102,184,127,183,115,176,111,169,115,175,121,152,117,137,99,122,103,125,107,137,105,126,108,122,121,116,105,92,99,101,95,99,88,78,69,67,88,106,142,144,159,134,206,181,212,206,172,223,162,233,141,229,137,239,140,255,129,248,113,252,104,237,96,261,95,261,89,276,98,273,106,281,101,297,116,291,133,297,139,277,143,276,156,301,150,325,169,351,169,348,160,330,152,356,154],
  [357,376,353,370,321,372,307,361,297,309,289,310,272,293,268,267,248,267,219,250,170,296,166,338,171,347,214,357,227,345,240,347,248,356,255,356,264,374,298,377,304,384,305,392,288,385,290,405,282,427,319,411,314,403,337,396,357,376],
];

function canadaMap(s) {
  MAP_REGIONS.forEach(function (coords) {
    const pts = [];
    for (var i = 0; i < coords.length; i += 2) {
      pts.push({ x: MAP_ORIGIN[0] + coords[i] / 100, y: MAP_ORIGIN[1] + coords[i + 1] / 100 });
    }
    const xs = pts.map(function (p) { return p.x; });
    const ys = pts.map(function (p) { return p.y; });
    const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
    s.addShape('custGeom', {
      x: x0, y: y0,
      w: Math.max.apply(null, xs) - x0, h: Math.max.apply(null, ys) - y0,
      points: pts.map(function (p) { return { x: p.x - x0, y: p.y - y0 }; }).concat([{ close: true }]),
      fill: { color: GREEN }, line: { color: CREAM, width: 1.5 },
    });
  });
}

// ---------------------------------------------------------------- slides
function slide01(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.725, 1.422], dots: [12.353, 4.494], footer: "FFFFFF" });
  s.addShape('rect', { x: 0, y: 5.129, w: 11.667, h: 2.371, fill: { color: "97C67E" } });
  imageBox(s, 7.884, 3.171, 2.941, 1.779);
  imageBox(s, 4.775, 4.494, 2.941, 1.779);
  imageBox(s, 1.667, 4.494, 2.941, 1.779);
  text(s, "FOZZA", 2.222, 1.227, 5.494, 1.025, { fontSize: 96, fontFace: "Lora", bold: true });
  text(s, "Exploring the Evolving Role of News in Shaping Perspectives, Informing Decisions, and Impacting Societal Narratives in the Digital Age", 7.884, 1.844, 3.783, 1.025, { fontSize: 14, fontFace: "Lato Light", italic: true, color: "595959", lineSpacingMultiple: 1.5 });
  text(s, "NEWS", 2.222, 2.253, 5.494, 1.025, { fontSize: 96, fontFace: "Lora", bold: true });
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In vitae sapien ac diam aliquam pretium id id erat. Morbi porta molestie mauris in egestas.", 7.884, 5.245, 3.508, 0.792, { color: "FFFFFF" });
}

function slide02(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.725, 3.339], dots: [0.438, 4.494], footer: "595959" });
  title(s, "Welcome", "Message", 1.667, 1.497, 3.698);
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In vitae sapien ac diam aliquam pretium id id erat. Morbi porta molestie mauris in egestas.", 2.704, 3.271, 3.508, 0.792);
  s.addShape('rect', { x: 8.347, y: 4.354, w: 4.986, h: 3.146, fill: { color: "97C67E" } });
  para(s, "Donec cursus ante vitae nisl rhoncus semper. Aliquam erat volutpat. Vestibulum sit amet orci a risus mattis cursus in vitae ex. Aliquam non ante sit amet libero rhoncus mattis ut in lorem. Maecenas vulputate ornare massa, non tristique dui tincidunt a.", 2.704, 4.354, 3.508, 1.034);
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In vitae sapien ac diam aliquam pretium id id erat. Morbi porta molestie mauris in egestas.", 7.306, 2.193, 4.293, 0.509);
  imageBox(s, 7.306, 3.339, 2.05, 2.05);
  imageBox(s, 9.549, 3.339, 2.05, 2.05);
  imageBox(s, 11.791, 3.339, 1.542, 2.05);
}

function slide03(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [4.989, 3.339], dots: [0.438, 4.494], footer: "595959" });
  title(s, "Table OF", "Content", 1.667, 1.497, 3.698);
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ", 6.064, 3.339, 2.201, 0.518);
  lead(s, "About Our Fozza", 6.076, 2.986, 1.864, 0.388);
  text(s, "1", 6.076, 2.384, 0.553, 0.636, { fontSize: 44, fontFace: "Lora", bold: true });
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ", 8.956, 3.339, 2.201, 0.518);
  lead(s, "Vision & Mission", 8.969, 2.986, 1.864, 0.388);
  text(s, "2", 8.969, 2.384, 0.553, 0.636, { fontSize: 44, fontFace: "Lora", bold: true });
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ", 6.064, 5.394, 2.201, 0.518);
  lead(s, "Fozza Timeline", 6.076, 5.04, 1.864, 0.388);
  text(s, "3", 6.076, 4.439, 0.553, 0.636, { fontSize: 44, fontFace: "Lora", bold: true });
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ", 8.956, 5.394, 2.201, 0.518);
  lead(s, "Fozza Services", 8.969, 5.04, 1.864, 0.388);
  text(s, "4", 8.969, 4.439, 0.553, 0.636, { fontSize: 44, fontFace: "Lora", bold: true });
  s.addShape('rect', { x: 12.895, y: 2.384, w: 0.439, h: 5.116, fill: { color: "97C67E" } });
  imageBox(s, 1.667, 4.494, 1.997, 1.997);
}

function slide04(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.725, 3.339], dots: [12.556, 4.494], page: "01", footer: "595959" });
  title(s, "About Our", "Fozza", 1.667, 1.497, 3.698);
  para(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In vitae sapien ac diam aliquam pretium id id erat. Morbi porta molestie mauris in egestas.", 2.704, 3.271, 3.508, 0.792);
  para(s, "Donec cursus ante vitae nisl rhoncus semper. Aliquam erat volutpat. Vestibulum sit amet orci a risus mattis cursus in vitae ex. Aliquam non ante sit amet libero rhoncus mattis ut in lorem. Maecenas vulputate ornare massa.", 6.996, 3.271, 3.508, 1.034);
  s.addShape('rect', { x: 2.704, y: 7.233, w: 10.63, h: 0.267, fill: { color: "97C67E" } });
  para(s, "Donec cursus ante vitae nisl rhoncus semper. Aliquam erat volutpat. Vestibulum sit amet orci a risus mattis cursus in vitae ex. Aliquam non ante sit amet libero rhoncus mattis ut in lorem. Maecenas vulputate ornare massa.", 6.996, 4.577, 3.508, 1.034);
  imageBox(s, 2.704, 4.67, 1.821, 1.821);
  imageBox(s, 6.996, 1.156, 1.821, 1.821);
}

function slide05(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [6.686, 1.151], page: "02", footer: "595959" });
  imageBox(s, 7.103, 1.148, 6.231, 6.352);
  title(s, "Our", "Vision", 8.191, 2.358, 3.698, { fontSize: 54, gap: 0.677, color: "FFFFFF" });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 1.456, 2.829, 3.513, 0.741);
  lead(s, "Vision 1", 1.468, 2.365, 1.864, 0.388);
  para(s, "Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Cras dapibus mauris nisi, a iaculis augue pulvinar nec. Praesent nisi mi, malesuada ac nunc ac, dictum pellentesque velit. ", 1.456, 4.615, 3.513, 0.973);
  lead(s, "Vision 2", 1.468, 4.15, 1.864, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 8.191, 4.844, 3.513, 0.741, { color: "FFFFFF" });
}

function slide06(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [4.989, 3.339], dotsRow: [5.88, 6.734], page: "02", footer: "595959" });
  title(s, "Our", "Mission", 1.667, 1.497, 3.698);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 5.867, 2.719, 3.513, 0.741);
  lead(s, "Mission 1", 5.88, 2.254, 1.864, 0.388);
  para(s, "Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Cras dapibus mauris nisi, a iaculis augue pulvinar nec. Praesent nisi mi, malesuada ac nunc ac, dictum pellentesque velit. ", 5.867, 4.504, 3.513, 0.973);
  lead(s, "Mission 2", 5.88, 4.04, 1.864, 0.388);
  s.addShape('rect', { x: 13.167, y: 2.598, w: 0.167, h: 0.973, fill: { color: "97C67E" } });
  imageBox(s, 0, 3.339, 4.111, 2.139);
  imageBox(s, 10.542, 4.504, 2.792, 2.996);
}

function slide07(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dotsRow: [10.154, 0.399], page: "03", footer: "595959" });
  title(s, "Fozza", "Timeline", 1.667, 1.497, 3.698);
  rule(s, 4.056, 3.365, 9.278, 0);
  rule(s, 4.32, 3.365, 0, 4.135);
  s.addShape('ellipse', { x: 4.685, y: 3.159, w: 0.413, h: 0.413, fill: { color: "97C67E" } });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 4.685, 4.287, 3.119, 1.022);
  lead(s, "Fisrt Time", 4.698, 3.822, 1.864, 0.388);
  rule(s, 8.17, 3.365, 0, 4.135);
  s.addShape('ellipse', { x: 8.535, y: 3.159, w: 0.413, h: 0.413, fill: { color: "97C67E" } });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 8.535, 4.287, 3.119, 1.022);
  lead(s, "Second Time", 8.547, 3.822, 1.864, 0.388);
  para(s, "Proin euismod velit elementum urna condimentum, ac interdum tortor consectetur.", 8.535, 2.133, 3.119, 0.476);
  imageBox(s, 4.685, 5.714, 3.119, 1.371);
  imageBox(s, 8.547, 5.714, 3.119, 1.371);
}

function slide08(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 5.727], dotsRow: [10.154, 0.399] });
  title(s, "Fozza", "Timeline", 1.667, 1.497, 3.698);
  rule(s, 0.017, 3.365, 9.278, 0);
  rule(s, 1.583, 3.365, 0, 4.135);
  s.addShape('ellipse', { x: 1.948, y: 3.159, w: 0.413, h: 0.413, fill: { color: "97C67E" } });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 1.948, 4.287, 3.119, 1.022);
  lead(s, "Thirdt Time", 1.96, 3.822, 1.864, 0.388);
  rule(s, 5.433, 3.365, 0, 4.135);
  s.addShape('ellipse', { x: 5.798, y: 3.159, w: 0.413, h: 0.413, fill: { color: "97C67E" } });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 5.798, 4.287, 3.119, 1.022);
  lead(s, "Fourth Time", 5.81, 3.822, 1.864, 0.388);
  para(s, "Proin euismod velit elementum urna condimentum, ac interdum tortor consectetur.", 10.155, 4.292, 1.781, 0.755);
  s.addShape('rect', { x: 12.895, y: 3.365, w: 0.439, h: 4.135, fill: { color: "97C67E" } });
  imageBox(s, 1.948, 5.714, 3.119, 1.371);
  imageBox(s, 5.81, 5.714, 3.119, 1.371);
}

function slide09(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.861, 3.339], dots: [0.438, 3.339], page: "05", footer: "595959" });
  s.addShape('rect', { x: 5.016, y: 3.339, w: 8.318, h: 3.744, fill: { color: "97C67E" } });
  title(s, "Fozza", "Services", 1.667, 1.497, 3.698);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 5.571, 5.31, 2.939, 0.774, { color: "FAF6F5" });
  lead(s, "Business and Strategy Consulting:", 5.571, 4.648, 2.939, 0.571, { color: "FFFFFF", lineSpacingMultiple: 1 });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 8.728, 5.31, 2.939, 0.774, { color: "FAF6F5" });
  lead(s, "Custom Software Development:", 8.728, 4.648, 2.939, 0.571, { color: "FFFFFF", lineSpacingMultiple: 1 });
  para(s, "Proin euismod velit elementum urna condimentum, ac interdum tortor consectetur.", 1.667, 5.346, 2.036, 0.738);
  imageBox(s, 5.571, 1.497, 2.939, 2.939);
  imageBox(s, 8.728, 1.497, 2.939, 2.939);
}

function slide10(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.861, 3.339], dots: [0.438, 3.339], page: "06", footer: "595959" });
  text(s, [
    { text: "Meet The ", options: { fontSize: 44, fontFace: "Lora", bold: true } },
    { text: "Team", options: { fontSize: 44, fontFace: "Lora", bold: true } },
  ], 1.667, 1.497, 4.861, 0.636);
  para(s, "Marketing", 5.101, 3.558, 1.616, 0.237);
  lead(s, "Eleanor Fitzgerald", 5.101, 3.241, 2.2, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 5.101, 3.864, 2.335, 0.702);
  para(s, "Marketing", 5.101, 5.479, 1.616, 0.237);
  lead(s, "Jonathan Patterson", 5.101, 5.161, 2.335, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 5.101, 5.785, 2.335, 0.702);
  para(s, "Marketing", 10.124, 3.558, 1.616, 0.237);
  lead(s, "Kimberly Nguyen", 10.124, 3.241, 2.2, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 10.124, 3.864, 2.335, 0.702);
  para(s, "Marketing", 10.124, 5.479, 1.616, 0.237);
  lead(s, "Muhammad Patel", 10.124, 5.161, 2.335, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 10.124, 5.785, 2.335, 0.702);
  imageBox(s, 2.933, 2.761, 1.806, 1.806);
  imageBox(s, 2.933, 4.685, 1.806, 1.806);
  imageBox(s, 7.956, 4.685, 1.806, 1.806);
  imageBox(s, 7.956, 2.761, 1.806, 1.806);
}

function slide11(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [8.381, 1.53], dots: [0.438, 3.339], page: "06", footer: "595959" });
  text(s, [
    { text: "Meet The ", options: { fontSize: 44, fontFace: "Lora", bold: true } },
    { text: "Leader", options: { fontSize: 44, fontFace: "Lora", bold: true } },
  ], 2.038, 3.212, 2.675, 2.007);
  text(s, "Marketing", 5.121, 3.511, 1.688, 0.388, { fontSize: 14, fontFace: "Lato Light", italic: true, color: "595959", lineSpacingMultiple: 1.5 });
  lead(s, "Jonathan Patterson", 5.121, 3.116, 2.615, 0.388, { fontSize: 20 });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ", 5.121, 5.321, 2.66, 0.702);
  imageBox(s, 8.788, 1.477, 4.545, 4.545);
}

function slide12(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.725, 1.422], dots: [12.353, 4.494], page: "07", footer: "595959" });
  s.addShape('rect', { x: 7.448, y: 4.494, w: 4.187, h: 3.006, fill: { color: "97C67E" } });
  title(s, "Break", "Slide", 2.222, 1.227, 5.494, { fontSize: 96 });
  text(s, "“ The journey of a thousand miles begins with one step. Life is an intricate tapestry woven with threads of challenges, triumphs, and the unexplored. “", 1.662, 4.738, 3.928, 1.34, { fontSize: 14, fontFace: "Lato Light", italic: true, color: "595959", lineSpacingMultiple: 1.5 });
  imageBox(s, 10.966, 2.094, 1.896, 1.451);
  imageBox(s, 8.911, 2.094, 1.896, 1.451);
  imageBox(s, 6.879, 3.715, 3.928, 3.006);
}

function slide13(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [4.681, 3.339], dots: [0.438, 3.339], page: "08", footer: "595959" });
  title(s, "SWOT", "Analysis", 1.667, 1.497, 3.208);
  badge(s, "O", 5.625, 4.421, 0.559);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 5.625, 5.549, 3.119, 0.942);
  lead(s, "Opportunities", 5.637, 5.084, 1.864, 0.388);
  badge(s, "T", 9.435, 4.421, 0.559);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 9.435, 5.549, 3.119, 0.942);
  lead(s, "Threats", 9.448, 5.084, 1.864, 0.388);
  badge(s, "S", 5.625, 1.896, 0.559);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 5.625, 3.024, 3.119, 0.942);
  lead(s, "Strengths", 5.637, 2.56, 1.864, 0.388);
  badge(s, "W", 9.435, 1.896, 0.559);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 9.435, 3.024, 3.119, 0.942);
  lead(s, "Weaknesses", 9.448, 2.56, 1.864, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 1.657, 3.273, 2.246, 1.262);
}

function slide14(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [4.875, 3.339], dots: [12.555, 3.339], page: "09", footer: "595959" });
  title(s, "Gallery", "Fozza", 1.667, 1.497, 3.208);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique", 6.042, 3.633, 2.036, 0.509);
  lead(s, "Build Project", 6.042, 4.189, 1.333, 0.388);
  para(s, "2020", 7.375, 4.301, 0.702, 0.268, { bold: true, align: 'right' });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique", 8.944, 3.633, 2.036, 0.509);
  lead(s, "Enter Project", 8.944, 4.189, 1.333, 0.388);
  para(s, "2021", 10.278, 4.301, 0.702, 0.268, { bold: true, align: 'right' });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique", 1.667, 5.443, 2.036, 0.509);
  lead(s, "New Project", 1.667, 6, 1.333, 0.388);
  para(s, "2019", 3, 6.112, 0.702, 0.268, { bold: true, align: 'right' });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 6.042, 5.843, 4.939, 0.509);
  s.addShape('rect', { x: 12.511, y: 5.952, w: 0.439, h: 1.548, fill: { color: "97C67E" } });
  imageBox(s, 1.667, 3.308, 2.036, 2.036);
  imageBox(s, 6.042, 1.497, 2.036, 2.036);
  imageBox(s, 8.944, 1.497, 2.036, 2.036);
}

function slide15(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dots: [12.555, 5.087], page: "10", footer: "595959" });
  bulletList(s, ["Employee Training Program:", "Leadership Development Package: $2,000"], 6.073, 6.403, 2.87, 0.533);
  text(s, [
    { text: "Starting at ", options: { fontSize: 10, fontFace: "Lato Light", italic: true, color: "595959" } },
    { text: "$500", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959" } },
    { text: " per session", options: { fontSize: 10, fontFace: "Lato Light", italic: true, color: "595959", breakLine: true } },
    { text: "$2,000", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959" } },
  ], 9.292, 6.403, 2.073, 0.533, { align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 10 });
  title(s, "Pricing", "List", 1.667, 1.497, 3.208);
  bulletList(s, ["1 Hour Consultation: ", "Strategy Analysis Package:"], 6.073, 1.327, 2.87, 0.533);
  lead(s, "Business and Strategy Consulting", 6.086, 0.557, 2.388, 0.606, { lineSpacingMultiple: 1 });
  rule(s, 5.323, 0, 0, 7.5);
  text(s, [
    { text: "$150", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959", breakLine: true } },
    { text: "Starting at", options: { fontSize: 10, fontFace: "Lato Light", italic: true, color: "595959" } },
    { text: " $1,000", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959" } },
  ], 9.834, 1.327, 1.531, 0.533, { align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 10 });
  bulletList(s, ["Digital Monthly Marketing Campaigns:", "Online Presence Optimization:"], 6.073, 3.878, 2.87, 0.533);
  lead(s, "Integrated Digital Marketing", 6.086, 3.108, 2.388, 0.606, { lineSpacingMultiple: 1 });
  text(s, [
    { text: "Starting from ", options: { fontSize: 10, fontFace: "Lato Light", italic: true, color: "595959" } },
    { text: "$800 ", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959", breakLine: true } },
    { text: "$500", options: { fontSize: 10, fontFace: "Lato Light", bold: true, italic: true, color: "595959" } },
    { text: " /month", options: { fontSize: 10, fontFace: "Lato Light", italic: true, color: "595959" } },
  ], 9.834, 3.878, 1.531, 0.533, { align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 10 });
  rule(s, 5.323, 2.43, 6.81, 0);
  rule(s, 5.323, 5.07, 6.81, 0);
  lead(s, "Employee Training and Development", 6.086, 5.633, 2.388, 0.606, { lineSpacingMultiple: 1 });
  rule(s, 12.133, 0, 0, 7.5);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 1.657, 3.339, 2.246, 1.262);
}

function slide16(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dotsRow: [10.911, 6.732], page: "11", footer: "595959" });
  title(s, "Unique", "Mockup", 1.667, 1.497, 3.208);
  laptopMockup(s, 4.546, 1.25, 7.804, 4.476);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 1.667, 3.653, 2.242, 1.215);
  lead(s, "Website", 1.679, 3.189, 1.682, 0.388);
  imageBox(s, 5.531, 1.531, 5.855, 3.66);
}

function slide17(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dotsRow: [10.911, 6.732], page: "11", footer: "595959" });
  phoneMockup(s, 5.569, 1.297, 4.749, 7.691);
  title(s, "Unique", "Mockup", 2.757, 1.661, 3.208);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 2.763, 4.595, 2.242, 1.215);
  lead(s, "Mobile", 2.775, 4.026, 1.682, 0.388);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 10.515, 4.549, 2.246, 1.262);
  imageBox(s, 6.373, 2.622, 3.087, 4.878);
}

function slide18(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dotsRow: [10.911, 6.732], page: "11", footer: "595959" });
  title(s, "Unique", "Maps", 2.348, 1.687, 3.208);
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 2.353, 4.62, 2.974, 1.215);
  lead(s, "Map Information", 2.366, 4.052, 1.936, 0.388);
  canadaMap(s);
}

function slide19(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [0.541, 3.339], dots: [12.555, 5.087], page: "11", footer: "595959" });
  title(s, "Client", "Testimonial", 1.667, 1.34, 3.958, { gap: 0.568 });
  s.addShape('rect', { x: 1.667, y: 3.051, w: 4.931, h: 3.085, fill: { color: "97C67E" } });
  s.addShape('rect', { x: 6.736, y: 3.051, w: 4.931, h: 3.085, fill: { color: "97C67E" } });
  text(s, "“", 3.934, 3.229, 0.463, 0.242, { fontFace: "Lora", bold: true });
  para(s, "“Mauris ornare dui odio, non accumsan justo fringilla vitae. Ut eros felis, auctor vitae commodo porta, ultrices a metus. Sed non turpis ut neque pharetra ultricies id sit amet magna. “", 3.934, 3.675, 2.432, 1.252, { color: "FFFFFF" });
  para(s, "Marketing", 3.934, 5.693, 2.432, 0.237, { color: "FAF6F5" });
  lead(s, "Eleanor Fitzgerald", 3.934, 5.376, 2.432, 0.388, { color: "FFFFFF" });
  text(s, "“", 9.004, 3.229, 0.463, 0.242, { fontFace: "Lora", bold: true });
  para(s, "“Mauris ornare dui odio, non accumsan justo fringilla vitae. Ut eros felis, auctor vitae commodo porta, ultrices a metus. Sed non turpis ut neque pharetra ultricies id sit amet magna. “", 9.004, 3.675, 2.432, 1.252, { color: "FFFFFF" });
  para(s, "Marketing", 9.004, 5.693, 2.432, 0.237, { color: "FAF6F5" });
  lead(s, "Estelle Darcy", 9.004, 5.376, 2.432, 0.388, { color: "FFFFFF" });
  imageBox(s, 1.667, 3.051, 2.036, 3.085);
  imageBox(s, 6.736, 3.051, 2.036, 3.085);
}

function slide20(pres) {
  const s = newSlide(pres);
  chrome(s, { followUs: [1.725, 1.422], dots: [12.353, 4.494], footer: "595959" });
  imageBox(s, 5.873, 4.325, 2.76, 2.76);
  title(s, "Thank", "You", 2.222, 1.227, 4.444, { fontSize: 96 });
  s.addShape('rect', { x: 7.361, y: 3.114, w: 4.208, h: 2.632, fill: { color: "97C67E" } });
  text(s, "Get In Touch", 7.911, 3.474, 2.683, 0.447, { fontSize: 18, bold: true, color: "F4F3EE", lineSpacingMultiple: 1.5 });
  text(s, "Phone:", 7.915, 4.1, 0.767, 0.199, { fontSize: 10, bold: true, color: "F4F3EE", lineSpacingMultiple: 1 });
  para(s, "+123-456-7890", 7.911, 4.246, 1.422, 0.259, { color: "F8F8F5" });
  text(s, "Website:", 9.687, 4.1, 0.723, 0.199, { fontSize: 10, bold: true, color: "F4F3EE", lineSpacingMultiple: 1 });
  para(s, "www.yourwebsite.com", 9.683, 4.246, 1.335, 0.259, { color: "F8F8F5" });
  text(s, "Address", 7.915, 4.652, 0.767, 0.199, { fontSize: 10, bold: true, color: "F4F3EE", lineSpacingMultiple: 1 });
  para(s, "+ABC Solutions Inc. 123 Main Street Suite 456 Cityville, State 78901 United States", 7.911, 4.797, 3.101, 0.588, { color: "F8F8F5" });
  para(s, "Aliquam sit amet lorem nulla. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Donec tristique leo sem, ac mollis lorem semper vitae. ", 7.361, 1.902, 4.208, 0.749);
  imageBox(s, 2.984, 4.325, 2.76, 2.76);
}


// ---------------------------------------------------------------- build
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'FOZZA', width: 13.333, height: 7.5 });
  pres.layout = 'FOZZA';
  pres.author = 'Fozza News';
  pres.title = 'Fozza News Template Presentation';
  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pres); });
  return pres.writeFile({
    fileName: path.join(__dirname, '05af0e24-8e37-4289-a436-a9b486aef094_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); },
             function (e) { console.error(e); process.exit(1); });
