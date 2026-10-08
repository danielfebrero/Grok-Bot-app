/*
 * Standalone recreation of the "Empowerment Group" conference deck (51 slides)
 * using only pptxgenjs + the Node standard library.
 *
 * Raster artwork in the source deck is replaced by native pptxgenjs shapes.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const SLIDE_W = 26.665;
const SLIDE_H = 15.0;

const RED = 'FF3F31';
const PALE_RED = 'FFD9D6';   // accent1 lumMod 20% / lumOff 80%
const BLACK = '000000';
const WHITE = 'FFFFFF';
const GRAY = 'A5A5A5';
const CHECK_BG = 'E7E7E7';   // lgCheck pattern background of the "…with background" layouts

const MAJOR = 'Poppins Black';   // theme major latin font
const MINOR = 'Poppins';         // theme minor latin font
const FA_BRAND = 'Font Awesome 5 Brands Regular';
const FA_SOLID = 'Font Awesome 5 Free Solid';

/* ------------------------------------------------------------------ *
 * Free-form outlines (view-box coordinates, converted by `freeform`)
 * ------------------------------------------------------------------ */
const PATHS = {
  // top-right / bottom-left rounded "flag" corner, used on ~30 slides
  corner: { vb: [1084, 1206], d: [
    ['M', 1084, 0], ['L', 1084, 1065],
    ['C', 912, 1155, 717, 1206, 510, 1206],
    ['L', 510, 412], ['L', 0, 412], ['L', 0, 0], ['Z'] ] },

  // title slide: bottom-left quarter round
  titleBL: { vb: [1644, 811], d: [
    ['M', 1644, 0], ['C', 1644, 311, 1530, 594, 1341, 811],
    ['L', 0, 811], ['L', 0, 0], ['Z'] ] },

  // title slide: big right-hand block with the thin ledge
  titleBig: { vb: [7100, 3405], d: [
    ['M', 5527, 0], ['L', 5527, 1818], ['L', 0, 1818], ['L', 0, 1889],
    ['L', 5527, 1889], ['L', 5527, 2336], ['L', 6160, 2336], ['L', 6160, 3405],
    ['C', 6536, 3405, 6873, 3237, 7100, 2971], ['L', 7100, 0], ['Z'] ] },

  // collage frame that sits above the picture area (slides 10, 11, 13)
  collage: { vb: [3977, 1223], d: [
    ['M', 2902, 1223], ['L', 3977, 1223], ['L', 3977, 0],
    ['C', 3570, 52, 3226, 302, 3043, 650],
    ['L', 0, 650], ['L', 0, 721], ['L', 3009, 721],
    ['C', 2940, 874, 2902, 1044, 2902, 1223], ['Z'] ] },

  // wide bottom ledge with a quarter round on the left (slides 12, 29, 37, 44)
  ledgeL: { vb: [3119, 1209], d: [
    ['M', 3119, 592], ['L', 3119, 663], ['L', 797, 663],
    ['C', 620, 941, 334, 1143, 0, 1209],
    ['L', 0, 0], ['L', 991, 0],
    ['C', 991, 215, 936, 416, 839, 592], ['Z'] ] },

  // wide bottom ledge with a quarter round on the right (slides 20, 47)
  ledgeR: { vb: [4038, 1230], d: [
    ['M', 4038, 728], ['L', 1029, 728],
    ['C', 1097, 881, 1135, 1051, 1135, 1230],
    ['L', 0, 1230], ['L', 0, 0],
    ['C', 432, 34, 802, 291, 994, 657], ['L', 4038, 657], ['Z'] ] },

  // gallery slide 23 frames
  gallTop: { vb: [7100, 1981], d: [
    ['M', 6140, 0], ['L', 6140, 1014], ['L', 0, 1014], ['L', 0, 1085],
    ['L', 6140, 1085], ['L', 6140, 1981],
    ['C', 6528, 1981, 6874, 1802, 7100, 1522], ['L', 7100, 0], ['Z'] ] },
  gallBottom: { vb: [7098, 1234], d: [
    ['M', 7098, 752], ['L', 1389, 752],
    ['C', 1549, 544, 1644, 283, 1644, 0], ['L', 0, 0], ['L', 0, 1163],
    ['C', 128, 1209, 267, 1234, 411, 1234],
    ['C', 776, 1234, 1104, 1075, 1330, 823], ['L', 7098, 823], ['Z'] ] },

  // left half-disc used behind the numbered list (slide 36)
  halfDiscL: { vb: [1000, 2000], d: [
    ['M', 1000, 0], ['L', 1000, 2000],
    ['C', 447.7, 2000, 0, 1552.3, 0, 1000],
    ['C', 0, 447.7, 447.7, 0, 1000, 0], ['Z'] ] },

  // tick mark inside the red bullets of slide 37
  check: { vb: [100, 100], d: [
    ['M', 14, 52], ['L', 24, 41], ['L', 40, 57], ['L', 76, 21], ['L', 86, 32],
    ['L', 40, 78], ['Z'] ] },

  // shoulders of the little line-art people on the "Services" slide
  shoulders: { vb: [100, 50], d: [
    ['M', 0, 50], ['C', 0, 8, 100, 8, 100, 50], ['Z'] ] },
};

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */
const K = 0.5523; // cubic-bezier circle constant

function freeform(slide, def, x, y, w, h, opts) {
  const sx = w / def.vb[0];
  const sy = h / def.vb[1];
  const pts = def.d.map(seg => {
    if (seg[0] === 'M') return { x: seg[1] * sx, y: seg[2] * sy, moveTo: true };
    if (seg[0] === 'L') return { x: seg[1] * sx, y: seg[2] * sy };
    if (seg[0] === 'C') return { x: seg[5] * sx, y: seg[6] * sy,
      curve: { type: 'cubic', x1: seg[1] * sx, y1: seg[2] * sy, x2: seg[3] * sx, y2: seg[4] * sy } };
    return { close: true };
  });
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points: pts, line: { type: 'none' } }, opts));
}

function rect(slide, x, y, w, h, color, opts) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, opts));
}

function ellipse(slide, x, y, w, h, opts) {
  slide.addShape('ellipse', Object.assign({ x, y, w, h, line: { type: 'none' } }, opts));
}

function line(slide, x, y, w, h, color, width, dashType) {
  slide.addShape('line', { x, y, w, h, line: { color, width, dashType: dashType || 'solid' } });
}

/** paragraph text helper - every option mirrors the source placeholder */
function text(slide, content, o) {
  const body = Array.isArray(content)
    ? content.map((t, i) => ({ text: t, options: { breakLine: i < content.length - 1 } }))
    : content;
  slide.addText(body, Object.assign({
    fontFace: MINOR, fontSize: 36, color: BLACK, align: 'left', valign: 'top',
    margin: 0, wrap: true,
  }, o));
}

/** big display heading - the source uses cap="all" + 6pt tracking on Poppins Black */
function heading(slide, content, x, y, w, h, o) {
  const up = Array.isArray(content) ? content.map(t => t.toUpperCase()) : content.toUpperCase();
  text(slide, up, Object.assign({
    x, y, w, h, fontFace: MAJOR, fontSize: 95, charSpacing: 6,
    lineSpacing: 104, valign: 'middle',
  }, o));
}

/** small all-caps label (28pt, 6pt tracking) */
function label(slide, content, x, y, w, h, o) {
  const up = Array.isArray(content) ? content.map(t => t.toUpperCase()) : content.toUpperCase();
  text(slide, up, Object.assign({
    x, y, w, h, fontSize: 28, charSpacing: 6, lineSpacing: 38, valign: 'middle',
  }, o));
}

/** rounded outline "pill" in the top-right corner */
function pill(slide, caption, x, y, w) {
  slide.addShape('roundRect', { x, y, w, h: 0.986, rectRadius: 0.493,
    fill: { type: 'none' }, line: { color: BLACK, width: 3 } });
  label(slide, caption, x, y, w - 0.3, 0.986, { align: 'right', lineSpacing: 50.4 });
}

/** thin horizontal rule (black divider under headings, red picture ledges) */
const rule = (slide, x, y, w, color) => rect(slide, x, y, w, 0.098, color || BLACK);

/** decorative corners */
const cornerTR = slide => freeform(slide, PATHS.corner, 22.901, 0, 3.764, 4.194, { fill: { color: RED } });
const cornerBL = slide => freeform(slide, PATHS.corner, 0, 10.806, 3.764, 4.194,
  { fill: { color: RED }, flipH: true, flipV: true });

/** company logo, drawn from primitives instead of the embedded PNG */
function logo(slide) {
  const x = 1.542, y = 1.221, u = 6.378 / 1829; // source PNG is 1829x306 px
  ellipse(slide, x, y, 154 * u, 154 * u, { fill: { color: RED } });
  rect(slide, x + 77 * u, y + 77 * u, 153 * u, 154 * u, RED);
  text(slide, 'EMPOWERMENT GROUP', {
    x: x + 295 * u, y: y, w: 5.6, h: 231 * u,
    fontFace: MAJOR, fontSize: 31, charSpacing: 0.5, valign: 'bottom',
  });
}

/** circled slide number, bottom right (inherited from the slide master) */
function pageNum(slide, n) {
  slide.addShape('ellipse', { x: 24.16, y: 12.795, w: 0.984, h: 0.984,
    fill: { type: 'none' }, line: { color: BLACK, width: 3 } });
  text(slide, String(n), { x: 24.16, y: 12.795, w: 0.984, h: 0.984,
    fontSize: 28, align: 'center', valign: 'middle' });
}

/** full-bleed light "lgCheck" pattern background */
const checkBg = slide => rect(slide, 0, 0, SLIDE_W, SLIDE_H, CHECK_BG);

/* ------------------------------------------------------------------ *
 * Shared body copy
 * ------------------------------------------------------------------ */
const LOREM_COLS = 'Donec tristique sollicitudin nisi et lobortis. Sed vitae ligula rhoncus, ' +
  'eleifend justo eget, consequat augue. Etiam cursus erat ac. Phasellus a imperdiet lectus, ' +
  'sit amet rhoncus dui. Sed ac mattis neque. Aliquam erat volutpat. Aliquam iaculis consequat.';
const LOREM_ABOUT = 'Etiam sit amet cursus nulla. Nunc quis iaculis odio. Curabitur id faucibus ' +
  'neque. Sed ante tortor, tempor eget lectus sed, suscipit scelerisque metus. Vestibulum ' +
  'faucibus neque nunc, vitae ornare';
const LOREM_IPSUM = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. " +
  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an " +
  "unknown printer took.";

/* ------------------------------------------------------------------ *
 * Slide builders
 *
 * NOTE ON PICTURES: every "Picture Placeholder" in the source deck is an
 * empty placeholder (no bitmap is stored in the file) - only the company
 * logo ships as a raster. The logo is redrawn above from native shapes and
 * the empty picture frames are reproduced as their red ledger/frame shapes,
 * exactly as the reference renders.
 * ------------------------------------------------------------------ */
const slides = [];
const S = fn => slides.push(fn);

/* 1 - Title slide */
S(p => {
  const s = p.addSlide();
  rect(s, 3.855, 14.754, 20.786, 0.246, RED);
  freeform(s, PATHS.titleBL, 0, 12.183, 5.707, 2.817, { fill: { color: RED } });
  freeform(s, PATHS.titleBig, 2.014, 0, 24.647, 11.822, { fill: { color: RED } });
  logo(s);
  heading(s, 'conference', 8.608, 2.311, 16.136, 2.41,
    { fontSize: 141, lineSpacing: 169.2, align: 'right' });
  text(s, 'Opening reception Friday, March 18',
    { x: 8.608, y: 4.467, w: 16.136, h: 0.645, fontSize: 38, lineSpacing: 45.6, align: 'right' });
  text(s, ['yourwebsite.com', 'your@name.com', 'yourcompany'],
    { x: 4.359, y: 8.11, w: 4.88, h: 2.23, color: WHITE, lineSpacing: 53 });
  [['\uf245', FA_SOLID, 8.11], ['\uf0e0', FA_SOLID, 8.91], ['\uf39e', FA_BRAND, 9.71]]
    .forEach(([glyph, face, y]) => {
      rect(s, 3.526, y, 0.63, 0.63, WHITE);
      text(s, glyph, { x: 3.526, y, w: 0.63, h: 0.63, fontFace: face, fontSize: 24,
        align: 'center', valign: 'middle' });
    });
  rect(s, 14.675, 5.36, 11.986, 0.097, BLACK);
});

/* 2 - Table of contents */
S(p => {
  const s = p.addSlide();
  rule(s, 0, 5.399, 20.627);
  cornerTR(s); pageNum(s, 2); logo(s);
  heading(s, 'Table of contents', 5.654, 3.894, 15.422, 1.459);
  const items = ['Introduction', 'Stories', 'Infographics', 'Summary', 'Contacts'];
  const pages = ['02', '04', '28', '44', '48'];
  text(s, items, { x: 5.63, y: 5.334, w: 5.599, h: 7.183, fontSize: 54, lineSpacing: 101 });
  text(s, pages, { x: 19.6, y: 5.334, w: 1.171, h: 7.082, fontSize: 54, lineSpacing: 101,
    align: 'right' });
  // dotted leaders between the entry and its page number
  [[10.380, 6.319, 9.342], [8.302, 7.736, 11.420], [10.490, 9.148, 9.233],
   [9.417, 10.564, 10.306], [9.198, 11.981, 10.524]]
    .forEach(([x, y, w]) => line(s, x, y, w, 0, BLACK, 3, 'sysDot'));
});

/* 3 - Agenda */
S(p => {
  const s = p.addSlide();
  pageNum(s, 3); logo(s);
  heading(s, 'Agenda', 5.5, 4.043, 15.72, 1.459, { align: 'center' });
  const rows = [
    [4.118, 5.904, 7.104, 6.092, '3:00 pm', 'Lorem ipsum dolor  sit amet, consectetur adipiscing elit.'],
    [4.118, 9.138, 7.104, 9.327, '3:00 pm', 'Etiam lacinia fermentum est non lobortis. '],
    [14.338, 5.904, 17.324, 6.092, '5:00 pm', 'Vestibulum ante ipsum primis in faucibus orci luctus.'],
    [14.338, 9.138, 17.324, 9.327, '5:00 pm', 'Lorem Ipsum available, but the majority have.'],
  ];
  rows.forEach(([bx, by, tx, ty, time, body]) => {
    rect(s, bx, by, 2.48, 2.48, RED);
    heading(s, time, bx, by, 2.48, 2.48, { fontSize: 44, lineSpacing: 44, align: 'center' });
    text(s, body, { x: tx, y: ty, w: 5.688, h: 2.104, lineSpacing: 50.2 });
  });
  rule(s, 0, 5.483, 16.378);
  cornerTR(s);
});

/* 4 - Heading */
S(p => {
  const s = p.addSlide();
  pageNum(s, 4); logo(s);
  heading(s, 'Next Generation Leadership', 6.125, 4.566, 15.215, 3.029);
  text(s, 'Sed venenatis turpis quis quam aliquet imperdiet. In maximus, magna ut laoreet ' +
    'maximus, erat erat fringilla sem, nec rhoncus dolor eros id nunc curabitur.',
    { x: 6.125, y: 8.148, w: 11.811, h: 2.906, lineSpacing: 50.2 });
  rule(s, 0, 7.684, 14.961);
  cornerTR(s);
});

/* 5 - Heading with background */
S(p => {
  const s = p.addSlide();
  checkBg(s);
  pageNum(s, 5); logo(s);
  heading(s, 'Make the Member Connection', 1.521, 4.566, 14.482, 3.029);
  text(s, 'Sed venenatis turpis quis quam aliquet imperdiet. In maximus, magna ut laoreet ' +
    'maximus, erat erat fringilla sem, nec rhoncus dolor eros id nunc. ',
    { x: 1.521, y: 8.147, w: 12.699, h: 2.104, lineSpacing: 50.2 });
  rule(s, 0, 7.786, 11.457);
  cornerTR(s);
});

/* 6 - Heading with one column layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 6); logo(s);
  heading(s, 'Everything you need', 3.228, 4.566, 10.832, 3.029);
  text(s, ['Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin aliquet eget ' +
    'mauris a mollis. Aenean vel viverra quam.', '',
    'Phasellus ullamcorper augue et metus auctor maximus. Donec imperdiet, ex ut tincidunt ' +
    'congue, elit diam feugiat tellus.'],
    { x: 15.694, y: 4.222, w: 8.269, h: 6.311, lineSpacing: 50.2 });
  rule(s, 0, 7.746, 10.512);
});

/* 7 - Heading with one column layout and background */
S(p => {
  const s = p.addSlide();
  checkBg(s);
  heading(s, 'Moving to Mastery', 3.252, 4.622, 9.528, 2.917);
  pageNum(s, 7); logo(s);
  text(s, ['Nulla malesuada pellentesque posuere. Vivamus tristique ipsum nisi, non semper ' +
    'purus iaculis sit.', '',
    'Nullam tristique, massa a vehicula auctor, sem turpis tempus orci, at hendrerit urna ' +
    'sem eget elit.'],
    { x: 14.534, y: 4.222, w: 7.286, h: 6.277, lineSpacing: 50.2 });
  rule(s, 0, 7.746, 10.236);
  cornerTR(s);
});

/* 8 - Heading with multi column layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 8); logo(s);
  heading(s, 'Back on Top', 3.262, 4.584, 6.646, 3.029);
  text(s, LOREM_COLS, { x: 12.172, y: 4.222, w: 10.273, h: 5.61, lineSpacing: 50.2 });
  rule(s, 0, 7.746, 8.661);
  cornerTR(s); cornerBL(s);
});

/* 9 - Heading with multi column layout and background */
S(p => {
  const s = p.addSlide();
  checkBg(s);
  pageNum(s, 9); logo(s);
  heading(s, 'Get there faster', 10.942, 5.177, 14.202, 1.515);
  text(s, LOREM_COLS, { x: 10.97, y: 7.243, w: 10.273, h: 5.61, lineSpacing: 50.2 });
  rule(s, 10.97, 6.679, 15.709);
  cornerTR(s);
});

/* 10 - Cover slide collage 1 */
S(p => {
  const s = p.addSlide();
  pageNum(s, 10); logo(s);
  heading(s, 'About us', 1.731, 4.593, 8.859, 1.515);
  text(s, 'Etiam sit amet cursus nulla. Nunc quis iaculis odio. Curabitur id faucibus neque. ' +
    'Sed ante tortor, tempor eget lectus sed, suscipit scelerisque metus. Vestibulum faucibus ' +
    'neque nunc, vitae ornare turpis hendrerit quis. Orci varius natoque penatibus.',
    { x: 1.711, y: 6.614, w: 8.879, h: 5.61, lineSpacing: 50.2 });
  rect(s, 12.424, 12.019, 14.241, 0.255, RED);
  freeform(s, PATHS.collage, 12.424, 2.441, 14.241, 4.389, { fill: { color: RED } });
  rule(s, 0.001, 6.233, 8.898);
});

/* 11 - Cover slide collage 2 (mirrored) */
S(p => {
  const s = p.addSlide();
  pageNum(s, 11); logo(s);
  heading(s, 'Why us?', 16.877, 5.79, 8.267, 1.515);
  text(s, 'Duis fermentum, ipsum at feugiat congue, mi risus accumsan mi, ut molestie diam ' +
    'ipsum sed lectus. Nunc vestibulum nisl eget.',
    { x: 16.877, y: 7.838, w: 7.663, h: 3.506, lineSpacing: 50.2 });
  rect(s, 0, 12.161, 14.249, 0.255, RED);
  freeform(s, PATHS.collage, 0, 2.583, 14.249, 4.389, { fill: { color: RED }, flipH: true });
  rule(s, 16.889, 7.402, 9.803);
  cornerTR(s);
});

/* 12 - Left image layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 12); logo(s);
  heading(s, 'Take action', 13.332, 5.771, 11.812, 1.515);
  text(s, 'Ut rutrum risus sed sapien volutpat, viverra efficitur justo tincidunt. ' +
    'Pellentesque vitae augue porttitor, ultrices massa ac, ultricies nunc. Morbi vel magna ' +
    'eu risus tincidunt.',
    { x: 13.332, y: 8.162, w: 11.812, h: 2.805, lineSpacing: 50.2 });
  pill(s, 'Our mission', 19.945, 1.248, 5.199);
  rule(s, 13.353, 7.402, 13.307);
  rect(s, -0.025, 3.95, 10.123, 0.231, RED);
  freeform(s, PATHS.ledgeL, -0.018, 9.915, 10.116, 3.922, { fill: { color: RED } });
});

/* 13 - Right image layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 13); logo(s);
  heading(s, 'Share the Vision', 1.731, 4.589, 9.239, 3.029);
  text(s, LOREM_ABOUT + '.', { x: 1.711, y: 8.17, w: 9.259, h: 4.173, lineSpacing: 50.2 });
  pill(s, 'our vision', 20.297, 1.248, 4.847);
  rect(s, 12.424, 12.019, 14.241, 0.255, RED);
  freeform(s, PATHS.collage, 12.424, 2.441, 14.241, 4.389, { fill: { color: RED } });
  rule(s, 0.021, 7.693, 6.85);
});

/* 14 - Top image layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 14); logo(s);
  heading(s, 'Everything Counts', 5.335, 9.33, 16.266, 1.515);
  text(s, 'PLACEHOLDER' +
    'turpis egestas. Ut maximus felis ac eros ultricies, nec egestas urna dapibus. Ut justo ' +
    'dolor, vestibulum.',
    { x: 5.335, y: 11.373, w: 16.266, h: 2.104, lineSpacing: 50.2 });
  pill(s, 'Objectives', 20.269, 1.238, 4.875);
  rect(s, 5.328, 8.1, 15.676, 0.252, RED);
  rect(s, 5.328, 3.366, 15.676, 0.252, RED);
  rule(s, 0.021, 10.945, 20.984);
});

/* 15 - Bottom image layout */
S(p => {
  const s = p.addSlide();
  pageNum(s, 15); logo(s);
  heading(s, 'Partners in Progress', 4.415, 4.007, 17.822, 1.515);
  text(s, 'Nullam pellentesque faucibus imperdiet. Fusce efficitur magna et augue convallis, ' +
    'id tincidunt ipsum viverra. Phasellus consectetur at.',
    { x: 4.415, y: 6.052, w: 17.776, h: 1.402, lineSpacing: 50.2 });
  pill(s, 'who we are', 20.05, 1.238, 5.094);
  rect(s, 4.454, 13.229, 17.442, 0.252, RED);
  rect(s, 4.454, 8.495, 17.442, 0.252, RED);
  rule(s, 0.021, 5.687, 21.85);
});

/* 16 - What we do */
S(p => {
  const s = p.addSlide();
  pageNum(s, 16); logo(s);
  heading(s, 'what we do?', 5.28, 3.894, 16.158, 1.459, { align: 'center' });
  const cards = [
    [5.332, 6.053, 6.516, 'Virtual Solutions', 6.694, 'Etiam lacinia fermentum est non lobortis ', 4.765, 6.189],
    [15.390, 6.071, 6.049, 'Meetings & Events', 6.707, 'Vestibulum sollicitudin eget ante cursus ', 14.873, 6.189],
    [5.332, 9.073, 6.516, 'Associations', 9.716, 'Curabitur malesuada interdum scelerisque', 4.765, 9.222],
    [15.390, 9.079, 6.049, 'Tradeshows', 9.727, 'Venenatis sodales justo. In hac habitasse', 14.873, 9.222],
  ];
  cards.forEach(([x, y, w, title, by, body, dx, dy]) => {
    label(s, title, x, y, w, 0.533);
    text(s, body, { x, y: by, w, h: 1.442, lineSpacing: 50.2 });
    ellipse(s, dx, dy, 0.247, 0.247, { fill: { color: RED } });
  });
  rule(s, 0, 5.629, 18.425);
  cornerTR(s); cornerBL(s);
});

/* 17 - Goals */
S(p => {
  const s = p.addSlide();
  ellipse(s, 12.741, 7.186, 5.63, 5.63, { fill: { color: RED } });
  ellipse(s, 18.334, 2.648, 5.63, 5.63, { fill: { color: RED } });
  // line-art clock icon
  ellipse(s, 14.56, 8.70, 1.992, 1.992, { fill: { type: 'none' }, line: { color: BLACK, width: 2 } });
  line(s, 15.556, 9.696, 0, -0.62, BLACK, 2);
  line(s, 15.556, 9.696, 0.44, 0.30, BLACK, 2);
  // line-art award icon
  ellipse(s, 20.334, 3.90, 1.631, 1.613, { fill: { type: 'none' }, line: { color: BLACK, width: 2 } });
  ellipse(s, 20.579, 4.135, 1.143, 1.14, { fill: { type: 'none' }, line: { color: BLACK, width: 2 } });
  [[20.305, 5.304], [21.140, 5.300]].forEach(([rx, ry]) =>
    freeform(s, { vb: [100, 100], d: [['M', 20, 0], ['L', 100, 45], ['L', 62, 55],
      ['L', 45, 100], ['L', 0, 25], ['Z']] }, rx, ry, 0.855, 0.936,
      { fill: { type: 'none' }, line: { color: BLACK, width: 2 } }));
  pageNum(s, 17); logo(s);
  heading(s, 'Our goals', 1.731, 5.187, 9.239, 1.515);
  text(s, LOREM_ABOUT + ' turpis hendrerit quis. Orci varius natoque penatibus et magnis dis. ',
    { x: 1.711, y: 7.219, w: 9.259, h: 5.576, lineSpacing: 50.2 });
  label(s, 'Time-bound', 13.321, 11.258, 4.471, 0.428, { align: 'center' });
  label(s, 'Accepted', 18.894, 6.674, 4.508, 0.424, { align: 'center' });
  rule(s, 0, 6.811, 9.961);
});

/* 18 - Values */
S(p => {
  const s = p.addSlide();
  pageNum(s, 18); logo(s);
  heading(s, 'our values', 2.735, 3.881, 21.229, 1.459, { align: 'center' });
  const cols = [
    [2.736, 6.462, 'Aliquam iaculis tortor scelerisque, dictum magna egestas, aliquet sapien. ' +
      'Suspendisse luctus varius risus, at ultrices mauris ultrices.', 2.265],
    [10.106, 6.470, 'Nullam pellentesque faucibus imperdiet. Fusce efficitur magna et augue ' +
      'convallis, id tincidunt ipsum viverra. Phasellus consectetur at leo eget.', 9.642],
    [17.484, 6.479, 'Cras massa felis, porta non felis volutpat, suscipit pulvinar tellus. ' +
      'Sed finibus posuere nunc et auctor. Vestibulum nec orci quis nulla tempor. ', 17.018],
  ];
  cols.forEach(([x, w, body, dx]) => {
    text(s, body, { x, y: 6.153, w, h: 4.207, lineSpacing: 50.3 });
    ellipse(s, dx, 6.386, 0.247, 0.247, { fill: { color: RED } });
  });
  rule(s, 0, 5.629, 17.717);
});

/* 19 - Team */
S(p => {
  const s = p.addSlide();
  pageNum(s, 19); logo(s);
  heading(s, 'Our team', 4.508, 3.881, 17.699, 1.459, { align: 'center' });
  [[3.900, 'Jane Smith', 'Speaker'], [10.769, 'John Smith', 'Volunteer'],
   [17.642, 'Joan Doe', 'Manager']].forEach(([x, name, role]) => {
    label(s, name, x, 10.777, 5.126, 0.533, { align: 'center' });
    text(s, role, { x, y: 11.42, w: 5.126, h: 0.701, align: 'center', lineSpacing: 50.2 });
  });
  rule(s, 0, 5.629, 16.969);
});

/* 20 - Testimonials */
S(p => {
  const s = p.addSlide();
  pageNum(s, 20); logo(s);
  text(s, '\u201CThis conference is a great opportunity to connect with leaders in the ' +
    'industry, as well as peers dealing with the same day-to-day issues we are. Being able ' +
    'PLACEHOLDER' +
    'dealership well worth it.\u201D',
    { x: 15.693, y: 4.33, w: 9.451, h: 5.61, lineSpacing: 50.2 });
  text(s, 'Jane Smith, Digital Manager',
    { x: 15.74, y: 11.174, w: 9.404, h: 0.545, bold: true });
  pill(s, 'Testimonials', 19.749, 1.238, 5.395);
  rect(s, 0, 11.34, 13.344, 0.241, RED);
  freeform(s, PATHS.ledgeR, 0.003, 2.239, 13.34, 4.183, { fill: { color: RED } });
});

/* 21 - Partners */
S(p => {
  const s = p.addSlide();
  pageNum(s, 21); logo(s);
  heading(s, 'Our team', 2.701, 3.881, 21.064, 1.459, { align: 'center' });
  rule(s, 0, 5.629, 16.969);
  cornerTR(s); cornerBL(s);
});

/* 22 - Quotation */
S(p => {
  const s = p.addSlide();
  checkBg(s);
  text(s, 'Jane Smith, General Manager',
    { x: 1.711, y: 8.115, w: 11.03, h: 0.701, lineSpacing: 50.2 });
  pageNum(s, 22); logo(s);
  heading(s, '\u201CRight Time Right Now\u201D', 1.731, 4.645, 11.01, 2.917, { valign: 'top' });
  pill(s, 'quotation', 20.349, 1.238, 4.795);
  rule(s, 0, 7.775, 10.945);
  cornerBL(s);
});

/* 23 - Gallery with 1 image */
S(p => {
  const s = p.addSlide();
  pageNum(s, 23); logo(s);
  freeform(s, PATHS.gallTop, 2.014, 0.009, 24.651, 6.88, { fill: { color: RED } });
  freeform(s, PATHS.gallBottom, 0, 9.04, 24.644, 4.286, { fill: { color: RED } });
});

/* 24, 25, 26 - Galleries with 3 / 4 / 5 images: each frame is a pair of red ledges */
function gallery(p, num, frames, body) {
  const s = p.addSlide();
  pageNum(s, num); logo(s);
  frames.forEach(([x, y, w, h]) => {
    rect(s, x, y, w, 0.236, RED);
    rect(s, x, y + h, w, 0.236, RED);
  });
  text(s, body.t, { x: body.x, y: body.y, w: body.w, h: body.h, lineSpacing: 50.4 });
  cornerTR(s);
}
S(p => gallery(p, 24,
  [[9.053, 3.582, 7.244, 9.203], [1.508, 3.582, 7.244, 4.336], [1.508, 8.451, 7.244, 4.337]],
  { x: 17.467, y: 4.664, w: 6.497, h: 5.61, t: 'Curabitur id tellus nec sem pellentesque ' +
    'fermentum et eget orci. Sed tincidunt bibendum mi, in ullamcorper quam tempor nec. ' +
    'Donec in finibus dui. Fusce a condimentum ligula.' }));
S(p => gallery(p, 25,
  [[1.528, 3.582, 7.244, 4.336], [1.528, 8.451, 7.244, 4.337],
   [9.078, 3.582, 7.244, 4.336], [9.078, 8.451, 7.244, 4.334]],
  { x: 17.467, y: 4.664, w: 7.087, h: 5.61, t: 'Donec luctus diam vel velit pellentesque, ' +
    'sit amet commodo enim tincidunt. Ut et ornare mauris. Etiam venenatis, nibh id ' +
    'sollicitudin gravida, neque elit sodales quam, et vehicula dolor.' }));
S(p => gallery(p, 26,
  [[1.528, 3.582, 7.244, 4.336], [1.528, 8.451, 7.244, 4.337],
   [9.078, 3.582, 7.244, 4.336], [9.078, 8.451, 7.244, 4.334],
   [16.633, 3.582, 7.244, 4.336]],
  { x: 16.904, y: 8.681, w: 7.087, h: 2.805, t: 'Etiam venenatis, nibh id sollicitudin ' +
    'gravida, neque elit sodales quam, et vehicula dolor tellus id velit. ' }));

/* 27 - Progress (100% stacked bar chart, series name shown inside the filled part) */
S(p => {
  const s = p.addSlide();
  pageNum(s, 27); logo(s);
  s.addChart('bar', [
    { name: '75% growth', labels: [''], values: [75] },
    { name: '25% growth', labels: [''], values: [25] },
  ], {
    x: 2.111, y: 6.566, w: 22.438, h: 2.362,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 10, barOverlapPct: 100,
    chartColors: [RED, PALE_RED],
    catAxisHidden: true, valAxisHidden: true,
    showLegend: false, showValue: false,
  });
  // the reference prints the first series name across the bar; LibreOffice stacks
  // rotated data labels, so the caption is drawn as a plain text box instead
  text(s, '75% growth', { x: 2.111, y: 6.566, w: 22.438 * 0.75, h: 2.362,
    fontSize: 66, align: 'center', valign: 'middle' });
  text(s, 'Percent of satisfied customers over the past 30 years.',
    { x: 6.224, y: 9.752, w: 13.605, h: 0.701, lineSpacing: 50.2 });
});

/* 28 - Ratio */
S(p => {
  const s = p.addSlide();
  pageNum(s, 28); logo(s);
  [[3.467, 7.215, '22%', 'Product Launches',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin non fermentum nisi, ac placerat purus.'],
   [15.955, 7.366, '78%', 'Corporate Events',
    'Vestibulum non tellus vestibulum odio bibendum consequat. Sed tempus bibendum dui quis egestas.']]
    .forEach(([x, w, big, cap, body]) => {
      heading(s, big, x, 4.063, w, 2.973, { fontSize: 200, lineSpacing: 212 });
      label(s, cap, x, 7.517, w, 0.533);
      text(s, body, { x, y: 8.202, w, h: 2.906, lineSpacing: 50.4 });
    });
  pill(s, 'ratio', 21.808, 1.231, 3.336);
  line(s, 13.333, 3.502, 0, 7.874, BLACK, 2);
  cornerBL(s);
});

/* 29 - Numeric information */
S(p => {
  const s = p.addSlide();
  pageNum(s, 29); logo(s);
  heading(s, ['new', 'conference'], 13.924, 4.567, 11.22, 3.029);
  text(s, LOREM_IPSUM, { x: 13.924, y: 8.121, w: 11.22, h: 3.506, lineSpacing: 50.2 });
  rect(s, 0, 3.886, 10.866, 8.071, BLACK);
  heading(s, '99%', 0, 3.886, 10.866, 8.071,
    { fontSize: 250, lineSpacing: 250, color: WHITE, align: 'center' });
  rect(s, 0, 3.884, 10.868, 0.231, RED);
  freeform(s, PATHS.ledgeL, 0.007, 9.898, 10.868, 3.922, { fill: { color: RED } });
  rule(s, 13.924, 7.686, 12.756);
});

/* 30 - Column chart */
S(p => {
  const s = p.addSlide();
  pageNum(s, 30); logo(s);
  s.addChart('bar', [{ name: 'Series 1', labels: ['Item1', 'Item2', 'Item3', 'Item4'],
    values: [20, 50, 40, 60] }], {
    x: 1.521, y: 3.956, w: 11.220, h: 8.858,
    barDir: 'col', barGapWidthPct: 72, chartColors: [RED],
    showLegend: false, showValue: false,
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25,
    valAxisLineColor: BLACK, catAxisLineColor: BLACK,
    valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 20,
    catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 20,
    valGridLine: { style: 'none' },
  });
  heading(s, 'Annual income', 13.924, 4.567, 10.119, 3.029);
  text(s, LOREM_IPSUM, { x: 13.924, y: 8.121, w: 10.163, h: 3.506, lineSpacing: 50.2 });
  rule(s, 13.924, 7.686, 12.756);
  cornerTR(s);
});

/* 31 - Bar chart */
S(p => {
  const s = p.addSlide();
  pageNum(s, 31); logo(s);
  s.addChart('bar', [{ name: 'Value 2',
    labels: ['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5', 'Item 6', 'Item 7'],
    values: [1200, 1400, 1800, 1350, 1220, 1800, 1050] }], {
    x: 1.521, y: 3.956, w: 11.220, h: 8.268,
    barDir: 'bar', barGapWidthPct: 42, chartColors: [RED],
    showLegend: false, showValue: false,
    valAxisMaxVal: 2000, valAxisMinVal: 0, valAxisMajorUnit: 500,
    valAxisLineColor: BLACK, catAxisLineColor: BLACK,
    valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 20,
    catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 20,
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
  });
  heading(s, 'Spending comparison', 13.924, 4.004, 10.63, 3.029);
  text(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. " +
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an " +
    "unknown printer took a galley of type and scrambled it to.",
    { x: 13.924, y: 7.575, w: 10.63, h: 4.173, lineSpacing: 50.2 });
  rule(s, 13.924, 7.186, 12.756);
  cornerTR(s);
});

/* 32 - Line chart (source uses a smoothed scatter over 2016-2019) */
S(p => {
  const s = p.addSlide();
  pageNum(s, 32); logo(s);
  s.addChart('scatter', [
    { name: 'X-Axis', values: [2016, 2017, 2018, 2019] },
    { name: 'Values Y', values: [1, 3, 2, 6] },
  ], {
    x: 1.521, y: 3.956, w: 11.220, h: 8.268,
    chartColors: [RED], lineSize: 8, lineSmooth: true,
    lineDataSymbol: 'circle', lineDataSymbolSize: 12,
    showLegend: false, showValue: false,
    catAxisMinVal: 2016, catAxisMaxVal: 2020, catAxisMajorUnit: 1,
    valAxisMinVal: 0, valAxisMaxVal: 7, valAxisMajorUnit: 1,
    valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 20,
    catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 20,
    valAxisLineColor: BLACK, catAxisLineColor: BLACK,
    valGridLine: { color: 'BFBFBF', size: 0.5 }, catGridLine: { color: 'BFBFBF', size: 0.5 },
  });
  heading(s, 'return on investment', 13.924, 4.06, 10.63, 2.917);
  text(s, 'Vestibulum tincidunt varius sem, quis vulputate nulla venenatis et. Vestibulum ' +
    'vel lacus in lacus feugiat accumsan ut vitae velit. ',
    { x: 13.924, y: 7.564, w: 10.63, h: 2.805, lineSpacing: 50.2 });
  rule(s, 13.924, 7.186, 12.756);
});

/* 33 - Area chart */
S(p => {
  const s = p.addSlide();
  pageNum(s, 33); logo(s);
  const years = ['2016', '2017', '2018', '2019', '2020'];
  s.addChart('area', [
    { name: 'Income', labels: years, values: [14, 18, 15, 20, 16] },
    { name: 'Costs', labels: years, values: [10, 14, 8, 12, 10] },
  ], {
    x: 1.521, y: 3.956, w: 12.992, h: 8.268,
    chartColors: [RED, BLACK],
    showLegend: true, legendPos: 'r', legendFontFace: MINOR, legendFontSize: 28,
    showValue: false,
    valAxisLineColor: BLACK, catAxisLineColor: BLACK,
    valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 20,
    catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 20,
  });
  heading(s, 'Growth graph', 16.877, 4.004, 8.267, 3.029);
  text(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. " +
    "Lorem Ipsum has been the industry's standard dummy text since.",
    { x: 16.877, y: 7.915, w: 8.267, h: 3.506, lineSpacing: 50.2 });
  rule(s, 16.877, 7.186, 9.803);
});

/* 34 - Pie chart */
S(p => {
  const s = p.addSlide();
  pageNum(s, 34); logo(s);
  s.addChart('pie', [{ name: 'Sales', labels: ['Goods', 'Services', 'Other'],
    values: [0.35, 0.40, 0.25] }], {
    x: 5.653, y: 3.956, w: 12.992, h: 8.268,
    chartColors: [RED, BLACK, '808080'], dataBorder: { pt: 3, color: WHITE },
    showLegend: true, legendPos: 'r', legendFontFace: MINOR, legendFontSize: 36,
    showValue: true, dataLabelFormatCode: '0%', dataLabelColor: WHITE,
    dataLabelFontFace: MAJOR, dataLabelFontSize: 48,
  });
  pill(s, 'sales', 21.835, 1.179, 3.309);
  cornerBL(s);
});

/* 35 - Market share (three 75%-hole doughnut gauges) */
S(p => {
  const s = p.addSlide();
  pageNum(s, 35); logo(s);
  const gauges = [
    [1.543, 0.25, 'Sed maximus lectus quam, ut egestas quam luctus ac.', 1.542, 5.907],
    [10.379, 0.50, 'Vivamus maximus pharetra diam sit amet semper.', 10.402, 5.883],
    [19.233, 0.75, 'Praesent gravida volutpat ante et feugiat.', 19.238, 5.906],
  ];
  gauges.forEach(([x, pct, body, tx, tw]) => {
    s.addChart('doughnut', [{ name: 'Input actual value', labels: ['Actual', 'Remainder'],
      values: [pct, 1 - pct] }], {
      x, y: 3.363, w: 5.905, h: 5.905,
      chartColors: [RED, WHITE], holeSize: 75, dataBorder: { pt: 8, color: WHITE },
      showLegend: false, showValue: false,
    });
    // the percentage sits inside the ring; drawn as text so it stays on one line
    text(s, Math.round(pct * 100) + '%', { x, y: 3.363, w: 5.905, h: 5.905,
      fontFace: MAJOR, fontSize: 72, align: 'center', valign: 'middle' });
    text(s, body, { x: tx, y: 9.266, w: tw, h: 2.104, lineSpacing: 50.4 });
  });
  pill(s, 'Our pricing', 19.93, 1.21, 5.214);
});

/* 36 - List of items */
S(p => {
  const s = p.addSlide();
  const groups = [
    [1.921, 6.017, 6.724, 3.371, 6.828, 5.883,
     ['01', 'Vivamus cursus, libero quis condimentum dictum, ante mauris.'],
     ['02', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.']],
    [14.268, 18.345, 6.799, 15.909, 19.146, 6.163,
     ['03', 'Etiam et semper arcu. Quisque dictum quis elit in consectetur.'],
     ['04', 'Aliquam neque massa, eleifend ac viverra quis, iaculis id nibh.']],
  ];
  groups.forEach(([dx, lx, lw, nx, tx, tw, a, b]) => {
    freeform(s, PATHS.halfDiscL, dx, 3.968, 4.134, 8.268, { fill: { color: RED } });
    [4.82, 7.696, 8.375, 11.278].forEach(y => line(s, lx, y, lw, 0, RED, 5));
    [[a, 5.557, 5.190], [b, 9.033, 8.747]].forEach(([item, ny, ty]) => {
      text(s, item[0], { x: nx, y: ny, w: 2.687, h: 1.627, fontFace: MAJOR, fontSize: 96,
        lineSpacing: 116, valign: 'middle' });
      text(s, item[1], { x: tx, y: ty, w: tw, h: 2.104, lineSpacing: 50.4, valign: 'middle' });
    });
  });
  pageNum(s, 36); logo(s);
  pill(s, 'LIST OF ITEMS', 19.806, 1.238, 5.338);
});

/* 37 - List with image */
S(p => {
  const s = p.addSlide();
  pageNum(s, 37); logo(s);
  [[3.968, 'Top quality', 5.714,
    'Pellentesque eget nibh sit amet nulla ullamcorper pulvinar sit amet sit amet.', 5.446, 12.795, 4.061],
   [8.123, 'Best price', 9.866,
    'Praesent vel felis vitae metus finibus faucibus. Fusce ac dictum purus.', 9.566, 11.378, 8.190]]
    .forEach(([hy, title, by, body, ry, rw, cy]) => {
      heading(s, title, 3.887, hy, 9.995, 1.471, { fontSize: 96, charSpacing: 0 });
      text(s, body, { x: 3.887, y: by, w: 9.995, h: 1.442, lineSpacing: 50.4 });
      rule(s, -0.015, ry, rw);
      ellipse(s, 2.222, cy, 0.984, 0.981, { fill: { color: RED } });
      freeform(s, PATHS.check, 2.222, cy, 0.984, 0.981, { fill: { color: BLACK } });
    });
  rect(s, 15.694, 11.032, 10.989, 0.231, RED);
  freeform(s, PATHS.ledgeL, 15.694, 1.376, 10.989, 3.922,
    { fill: { color: RED }, flipH: true, flipV: true });
});

/* 38 - List of images */
S(p => {
  const s = p.addSlide();
  pageNum(s, 38); logo(s);
  [[2.406, 3.748, 6.139, 4.336, 9.241, 6.108,
    'Vivamus cursus, libero quis condimentum dictum ante mauris.'],
   [10.289, 3.748, 6.139, 5.308, 10.197, 6.108,
    'Lorem ipsum dolor sit amet, consectetur adipiscing eit.'],
   [18.172, 3.748, 6.139, 4.336, 9.241, 6.122,
    'Vivamus tellus odio, ultrices ac consectetur non, sollicitudin et.']]
    .forEach(([x, y, w, h, ty, tw, body]) => {
      rect(s, x, y, w, 0.236, RED);
      rect(s, x, y + h, w, 0.236, RED);
      text(s, body, { x: x + 0.011, y: ty, w: tw, h: 2.104, lineSpacing: 50.4 });
    });
  pill(s, 'List of Images', 19.3, 1.231, 5.844);
});

/* 39 - Services */
S(p => {
  const s = p.addSlide();
  pageNum(s, 39); logo(s);
  heading(s, 'our services', 3.292, 3.833, 20.125, 1.459, { align: 'center' });
  // line-art icons: slide + audience, speaker at a lectern, panel of people
  const stroke = { fill: { type: 'none' }, line: { color: BLACK, width: 2 } };
  const person = (x, y, r) => {
    ellipse(s, x, y, r * 2, r * 2, stroke);
    freeform(s, PATHS.shoulders, x - r * 0.5, y + r * 2.1, r * 3, r * 1.5, stroke);
  };
  // presentation: screen with rules, four listeners below
  s.addShape('rect', Object.assign({ x: 4.082, y: 6.425, w: 3.225, h: 1.35 }, stroke));
  [0, 1, 2].forEach(i => rect(s, 5.35, 6.72 + i * 0.31, 1.7, 0.09, BLACK));
  s.addShape('rect', Object.assign({ x: 4.35, y: 6.75, w: 0.85, h: 0.7 }, stroke));
  [0, 1, 2, 3].forEach(i => person(4.30 + i * 0.79, 8.00, 0.24));
  // speaker behind a lectern
  person(12.94, 6.42, 0.42);
  s.addShape('trapezoid', Object.assign({ x: 12.45, y: 7.55, w: 1.80, h: 1.05 }, stroke));
  s.addShape('rect', Object.assign({ x: 12.20, y: 8.60, w: 2.308, h: 0.70 }, stroke));
  // panel: whiteboard flanked by four people
  s.addShape('rect', Object.assign({ x: 19.90, y: 6.60, w: 1.90, h: 2.60 }, stroke));
  [18.90, 21.95].forEach(x => { person(x, 6.55, 0.38); person(x, 8.00, 0.38); });
  [[3.292, 4.726, 'Presentation', 'Lorem ipsum dolor sit amet.'],
   [10.959, 4.746, 'Speaker', 'Proin aliquet eget mauris a mollis.'],
   [18.648, 4.726, 'Presentation', 'Lorem ipsum dolor sit.']]
    .forEach(([x, w, cap, body]) => {
      label(s, cap, x, 9.874, w, 0.533, { align: 'center' });
      text(s, body, { x, y: 10.507, w, h: 1.402, align: 'center', lineSpacing: 50.4 });
    });
  rule(s, 0, 5.529, 18.425);
  cornerTR(s); cornerBL(s);
});

/* 40 - Pricing */
S(p => {
  const s = p.addSlide();
  pageNum(s, 40); logo(s);
  const plans = [
    [1.521, 4.714, 6.062, 6.910, 'basic', '$50', 'Proin ut velit et lorem laoreet mollis a eu erat.'],
    [9.028, 3.957, 8.618, 8.435, 'premium', '$250',
      'Fusce accumsan eros eu nibh convallis sodales. Vestibulum vitae libero viva.'],
    [19.082, 4.714, 6.062, 6.910, 'pro', '$100', 'Vestibulum vitae libero. Vivamus eu tellus mattis.'],
  ];
  plans.forEach(([x, y, w, h, name, price, body]) => {
    rect(s, x, y, w, h, BLACK);
    rect(s, x, y, w, 0.236, RED);
    rect(s, x, y + h - 0.236, w, 0.236, RED);
    heading(s, name, x, 5.557, w, 1.515, { fontSize: 90, color: WHITE, align: 'center' });
    text(s, body, { x, y: y + 0.9, w, h: h - 1.8, color: WHITE, align: 'center',
      valign: 'middle', lineSpacing: 50.4 });
    heading(s, price, x, 9.333, w, 1.515, { fontSize: 75, color: WHITE, align: 'center' });
  });
  pill(s, 'Our pricing', 19.972, 1.238, 5.172);
});

/* 41 - Vertical table */
S(p => {
  const s = p.addSlide();
  pageNum(s, 41); logo(s);
  const head = ['Name', 'Property', 'Price'];
  const body = [
    ['Lorem ipsum', 'Integer in ex varius ', '1000.00'],
    ['Dolor', 'Consectetur lorem', '1500.00'],
    ['Consectetur', 'Duis sit amet porta nisi', '3000.00'],
    ['Adipiscing', 'Mauris non ultrices velit', '4500.00'],
    ['Praesent', 'Aliquam mattis aliquam', '4750.00'],
    ['Tristique', 'Praesent vitae tincidunt', '5000.00'],
  ];
  const cellBase = { valign: 'middle', margin: [4, 28, 4, 28] };
  const rows = [head.map((t, i) => ({ text: t.toUpperCase(), options: Object.assign({}, cellBase, {
      fill: { color: RED }, fontFace: MINOR, fontSize: 28, charSpacing: 6, bold: false,
      align: i === 2 ? 'right' : 'left',
      border: [{ type: 'none' }, { type: 'solid', color: WHITE, pt: 2.25 },
               { type: 'none' }, { type: 'solid', color: WHITE, pt: 2.25 }] }) }))];
  const redBorder = [{ type: 'solid', color: RED, pt: 2.25 }, { type: 'solid', color: RED, pt: 2.25 },
                     { type: 'solid', color: RED, pt: 2.25 }, { type: 'solid', color: RED, pt: 2.25 }];
  body.forEach(r => {
    rows.push(r.map((t, i) => ({ text: t, options: Object.assign({}, cellBase, {
      fontFace: MINOR, fontSize: 36, align: i === 2 ? 'right' : 'left', border: redBorder }) })));
  });
  s.addTable(rows, { x: 3.827, y: 3.881, w: 18.901, colW: [4.728, 9.44, 4.733],
    rowH: [1.183, 1.177, 1.177, 1.177, 1.177, 1.177, 1.177] });
  cornerTR(s);
});

/* 42 - Horizontal table */
S(p => {
  const s = p.addSlide();
  pageNum(s, 42); logo(s);
  const data = [
    ['March', 'Integer in ex varius, congue ligula', '1000.00'],
    ['April', 'Duis sit amet porta nisi mauris non', '1500.00'],
    ['May', 'Aliquam mattis aliquam nibh praesent', '3000.00'],
    ['June', 'Donec eros nisi, sagittis non nulla', '4500.00'],
    ['July', 'Praesent placerat elit id erat rutrum ', '4750.00'],
    ['August', 'Vestibulum convallis vehicula', '5000.00'],
  ];
  const cellBase = { valign: 'middle', margin: [4, 30, 4, 30] };
  const rows = data.map(([m, desc, price]) => [
    { text: m.toUpperCase(), options: Object.assign({}, cellBase, {
      fill: { color: RED }, fontFace: MINOR, fontSize: 28, charSpacing: 6,
      border: [{ type: 'solid', color: WHITE, pt: 2.25 }, { type: 'none' },
               { type: 'solid', color: WHITE, pt: 2.25 }, { type: 'none' }] }) },
    { text: desc, options: Object.assign({}, cellBase, { fontFace: MINOR, fontSize: 36,
      border: [{ type: 'solid', color: RED, pt: 2.25 }, { type: 'solid', color: RED, pt: 2.25 },
               { type: 'solid', color: RED, pt: 2.25 }, { type: 'none' }] }) },
    { text: price, options: Object.assign({}, cellBase, { fontFace: MINOR, fontSize: 36,
      align: 'right',
      border: [{ type: 'solid', color: RED, pt: 2.25 }, { type: 'none' },
               { type: 'solid', color: RED, pt: 2.25 }, { type: 'solid', color: RED, pt: 2.25 }] }) },
  ]);
  s.addTable(rows, { x: 3.292, y: 4.343, w: 20.082, colW: [5.023, 11.305, 3.754],
    rowH: [1.181, 1.181, 1.181, 1.181, 1.181, 1.181] });
  pill(s, 'Price table', 20.185, 1.213, 4.959);
  cornerBL(s);
});

/* 43, 44 - Vertical timeline (with and without image column) */
const TIMELINE_BODY = [
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin aliquet eget mauris a mollis. ',
  'Maecenas ornare sed tortor ut eleifend. Sed porttitor diam sed varius rhoncus.',
  'Praesent cursus tincidunt nisl ac placerat. Quisque nec neque porta, commodo risus vitae.',
];
function verticalTimeline(p, num, days, spineY, spineH, dayW, dayX) {
  const s = p.addSlide();
  line(s, 14.71, spineY, 0, spineH, RED, 4);
  pageNum(s, num); logo(s);
  days.forEach((day, i) => {
    const y = 2.28 + i * 3.583;
    rect(s, 13.332, y, 2.756, 2.756, RED);
    heading(s, day, 13.332, y, 2.756, 2.756, { fontSize: 72, lineSpacing: 72, align: 'center' });
    text(s, 'March', { x: dayX, y: y + 1.48, w: dayW, h: 0.701, align: 'center',
      lineSpacing: 50.4, valign: 'middle' });
    text(s, TIMELINE_BODY[i], { x: 16.614, y: y + 0.326, w: 8.53, h: 2.104,
      lineSpacing: 50.4, valign: 'middle' });
  });
  return s;
}
S(p => {
  const s = verticalTimeline(p, 43, ['12', '13', '14'], 4.547, 10.453, 2.261, 13.579);
  heading(s, 'Timeline', 1.987, 4.726, 8.394, 1.459, { fontSize: 82 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent varius nisi a eros. ',
    { x: 1.987, y: 6.428, w: 8.394, h: 2.104, lineSpacing: 50.4 });
  rule(s, 0, 6.071, 7.795);
  cornerBL(s);
});
S(p => {
  const s = verticalTimeline(p, 44, ['15', '16', '17'], -0.02, 9.567, 2.101, 13.659);
  rect(s, 0, 3.884, 10.868, 0.231, RED);
  freeform(s, PATHS.ledgeL, 0, 9.898, 10.876, 3.922, { fill: { color: RED } });
});

/* 45, 46 - Horizontal timeline (with and without image band) */
const HTIMELINE_43 = ['Lorem ipsum dolor sit amet.', 'Proin porttitor orci magna.',
  'Suspendisse euismod mattis.', 'Cras sodales, ante vitae.'];
const HTIMELINE_46 = ['Vivamus non varius nisi.', 'Quisque tortor risus, finibus.',
  'Morbi nec justo sit amet tellus.', 'Nunc eget sapien laoreet.'];
function horizontalTimeline(p, num, days, boxY, bodies) {
  const s = p.addSlide();
  line(s, -0.256, boxY + 1.378, 20.63, 0, RED, 4);
  days.forEach((day, i) => {
    const x = 3.412 + i * 5.4725;
    rect(s, x, boxY, 2.756, 2.756, RED);
    heading(s, day, x, boxY, 2.756, 2.756, { fontSize: 72, lineSpacing: 72, align: 'center' });
    text(s, 'April', { x: x + 0.226, y: boxY + 1.48, w: 2.30, h: 0.701, align: 'center',
      lineSpacing: 50.4, valign: 'middle' });
    text(s, bodies[i], { x: x - 0.427, y: boxY + 3.144, w: 4.134, h: 1.402, lineSpacing: 50.4 });
  });
  pageNum(s, num); logo(s);
  return s;
}
S(p => {
  const s = horizontalTimeline(p, 45, ['10', '11', '12', '13'], 8.681, HTIMELINE_43);
  heading(s, 'timeline', 9.198, 4.005, 13.66, 1.515);
  text(s, 'Aliquam erat volutpat. Proin semper non purus et interdum. Quisque a eleifend magna.',
    { x: 9.198, y: 6.053, w: 13.66, h: 1.402, lineSpacing: 50.4 });
  rule(s, -0.015, 5.621, 15.827);
});
S(p => {
  const s = horizontalTimeline(p, 46, ['14', '15', '16', '17'], 9.279, HTIMELINE_46);
  rect(s, 0, 3.273, SLIDE_W, 0.231, RED);
  rect(s, 0, 8.577, SLIDE_W, 0.231, RED);
  pill(s, 'timeline', 21.07, 1.211, 4.074);
});

/* 47 - Summary */
S(p => {
  const s = p.addSlide();
  pageNum(s, 47); logo(s);
  heading(s, 'Summary', 15.694, 4.584, 8.946, 1.515, { valign: 'top' });
  text(s, 'Our experience in the hospitality industry stretches over a combined 24 years!',
    { x: 15.696, y: 6.413, w: 8.764, h: 2.104, lineSpacing: 50.2 });
  rect(s, 0, 4.458, 13.344, 0.241, RED);
  freeform(s, PATHS.ledgeR, 0.003, 9.617, 13.34, 4.183, { fill: { color: RED }, flipV: true });
  rule(s, 15.75, 6.1, 10.906);
  cornerTR(s);
});

/* 48 - Call to action */
S(p => {
  const s = p.addSlide();
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque nisi lorem, ' +
    'mollis ut leo vel, blandit consequat turpis. Ut ac hendrerit sem. Morbi consequat.',
    { x: 2.87, y: 5.46, w: 21.093, h: 1.402, lineSpacing: 50.2 });
  pageNum(s, 48); logo(s);
  heading(s, 'Discover our conference', 2.903, 3.415, 21.852, 1.515, { valign: 'top' });
  rect(s, 4.472, 13.544, 17.442, 0.252, RED);
  rect(s, 4.472, 7.902, 17.442, 0.252, RED);
  rule(s, 0, 5.105, 23.898);
});

/* 49 - Locations */
S(p => {
  const s = p.addSlide();
  pageNum(s, 49); logo(s);
  heading(s, 'OUR BRANCHES', 5.921, 3.118, 15.019, 1.515, { align: 'center' });
  [[8.444, 7.500, 9.338, 7.699, 2.852, 'New York', false],
   [13.924, 7.714, 14.839, 7.880, 2.237, 'Genova', false],
   [14.137, 10.955, 15.135, 11.211, 4.334, 'Johannesburg', true]]
    .forEach(([mx, my, tx, ty, tw, name, bold]) => {
      ellipse(s, mx, my, 0.753, 0.753, { fill: { color: RED } });
      ellipse(s, mx + 0.187, my + 0.187, 0.378, 0.378, { fill: { color: WHITE } });
      label(s, name, tx, ty, tw, 0.421,
        { fontFace: MAJOR, bold, lineSpacing: 30, fill: { color: WHITE } });
    });
});

/* 50 - Follow us */
S(p => {
  const s = p.addSlide();
  pageNum(s, 50); logo(s);
  heading(s, 'Follow US', 5.136, 4.019, 16.244, 1.515, { fontSize: 96, lineSpacing: 128,
    align: 'center' });
  [[4.860, 4.768, 3.153, '\uf39e', 'Facebook'],
   [11.756, 11.682, 3.153, '\uf16d', 'Instagram'],
   [18.648, 18.711, 2.889, '\uf099', 'Twitter']]
    .forEach(([cx, tx, tw, glyph, name]) => {
      ellipse(s, cx, 6.596, 3.15, 3.15, { fill: { color: RED } });
      text(s, glyph, { x: cx, y: 6.596, w: 3.15, h: 3.15, fontFace: FA_BRAND, fontSize: 120,
        align: 'center', valign: 'middle' });
      text(s, name, { x: tx, y: 10.103, w: tw, h: 0.701, align: 'center', lineSpacing: 50.2 });
    });
  rule(s, 0, 5.629, 17.48);
  cornerTR(s); cornerBL(s);
});

/* 51 - Contacts */
S(p => {
  const s = p.addSlide();
  pageNum(s, 51); logo(s);
  heading(s, 'Thank you', 5.655, 4.019, 15.496, 1.515, { align: 'center' });
  text(s, ['Visit us at:', 'For more info call:', 'Visit our website:'],
    { x: 5.655, y: 6.389, w: 7.087, h: 2.143, bold: true, align: 'right', lineSpacing: 50.4 });
  text(s, ['123 Street, City, State 45678', '1234 56 7890', 'www.companyname.com'],
    { x: 13.093, y: 6.389, w: 8.057, h: 2.104, lineSpacing: 50.4 });
  pill(s, 'contact us', 20.043, 1.211, 5.101);
  rule(s, 0, 5.678, 17.835);
  cornerBL(s);
});

/* ------------------------------------------------------------------ *
 * Build & save
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };
  slides.forEach(fn => fn(pptx));
  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '043278d0-217f-40aa-9208-865186bf72dc_grok_final.pptx'),
}).then(f => console.log('wrote', f));
