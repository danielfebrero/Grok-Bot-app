/* Recreation of the "Quazar Brand Manual & Guidelines 2021" deck with pptxgenjs.
 * Slide size 26.667 x 15 in (widescreen, 4:2.25). All geometry is in inches.
 * Raster photographs in the source deck are stand-in grey boxes; they are
 * reproduced here as programmatic placeholders (see imageBox / paperBox).
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ---------------------------------------------------------------- palette */
const DARK = '00406D';   // Primary Color A - Dark Blue
const CORAL = '00AEEF';  // Primary Color B - Coral Blue
const LIGHT = 'B9E5FA';  // Primary Color C - Light Blue
const GREY = '939598';   // body copy grey
const PALE = 'E6E7E8';   // light neutral
const WHITE = 'FFFFFF';
const PHOTO = '919191';  // placeholder tone used for photographs

const FONT = 'Roboto Regular';
const FONT_LIGHT = 'Roboto Light';
const FONT_BOLD = 'Roboto Bold';
const FONT_THIN = 'Roboto Thin';

const SLIDE_W = 26.6667;
const SLIDE_H = 15;

/* ------------------------------------------------------------ text styles */
const STYLE = {
  t10: { fontSize: 10, lineSpacing: 12, color: DARK },
  t13: { fontSize: 13, lineSpacing: 22, color: DARK },
  t16: { fontSize: 16, lineSpacing: 26, color: GREY },
  t22: { fontSize: 22, lineSpacingMultiple: 1.1, paraSpaceBefore: 2, color: DARK },
  t28: { fontSize: 28, lineSpacingMultiple: 1.1375, color: DARK },
  t34: { fontSize: 34, lineSpacingMultiple: 1.1, charSpacing: -0.61, color: DARK },
  t48: { fontSize: 48, lineSpacingMultiple: 1.1, color: DARK },
  t80: { fontSize: 80, lineSpacingMultiple: 1.1, color: CORAL, fontFace: FONT_THIN },
};

/* paragraph factories: P = one run, PR = several coloured runs on one line */
const P = (s, t, c, o) => ({ s, t, c, o });
const PR = (s, runs) => ({ s, runs });
const GAP = (s) => ({ s: s || 't16', t: '' });

/* ---------------------------------------------------------------- helpers */
function txt(sl, box, paras, opt) {
  const runs = [];
  paras.forEach((pa) => {
    const base = Object.assign({ fontFace: FONT }, STYLE[pa.s], pa.o || {});
    const parts = pa.runs || [[pa.t || '', pa.c]];
    parts.forEach((part, i) => {
      runs.push({
        /* a lone space keeps blank paragraphs at their own font size */
        text: part[0] === '' ? ' ' : part[0],
        options: Object.assign({}, base, {
          color: part[1] || pa.c || base.color,
          breakLine: i === parts.length - 1,
        }, part[2] || {}),
      });
    });
  });
  sl.addText(runs, Object.assign(
    { x: box[0], y: box[1], w: box[2], h: box[3], margin: 0, valign: 'top', align: 'left', isTextBox: true },
    opt || {}));
}

function rect(sl, x, y, w, h, fill, opt) {
  sl.addShape('rect', Object.assign({ x, y, w, h, fill: fill ? { color: fill } : { type: 'none' } }, opt || {}));
}

function hline(sl, x, y, w, color, width, dash) {
  sl.addShape('line', { x, y, w, h: 0, line: { color: color || CORAL, width: width || 1, dashType: dash } });
}

function vline(sl, x, y, h, color, width, dash) {
  sl.addShape('line', { x, y, w: 0, h, line: { color: color || DARK, width: width || 1, dashType: dash } });
}

/* the thin vertical column rules that structure nearly every inner page */
const RULE_A = 1.141, RULE_B = 9.746, RULE_C = 18.35, RULE_D = 25.589;
function rules(sl, xs) { xs.forEach((x) => vline(sl, x, 0, SLIDE_H, DARK, 1)); }

/* grey stand-in for a photograph: flat tone crossed by two light diagonals */
function imageBox(sl, x, y, w, h) {
  rect(sl, x, y, w, h, PHOTO);
  sl.addShape('line', { x, y, w, h, line: { color: PALE, width: 1.5 } });
  sl.addShape('line', { x, y, w, h, flipV: true, line: { color: PALE, width: 1.5 } });
}
/* stand-in for the light paper / stationery mock-up photographs */
function paperBox(sl, x, y, w, h, fill) {
  rect(sl, x, y, w, h, fill || 'FAFBFB', { line: { color: PALE, width: 0.75 } });
}

/* ------------------------------------------------------------ Quazar mark */
/* The logo is a bar, a ring with a tail and an angled swoosh. Everything is
 * expressed as a fraction of the mark size S so it can be stamped anywhere. */
const RING = [
  { x: 0.5, y: 0 },
  { x: 0, y: 0.5, curve: { type: 'cubic', x1: 0.2238, y1: 0, x2: 0, y2: 0.2238 } },
  { x: 0.062, y: 0.7412, curve: { type: 'cubic', x1: 0, y1: 0.5875, x2: 0.0225, y2: 0.6697 } },
  { x: 0.1271, y: 0.833, curve: { type: 'cubic', x1: 0.0802, y1: 0.7743, x2: 0.1021, y2: 0.805 } },
  { x: 0.136, y: 0.8425 },
  { x: 0.2814, y: 1 },
  { x: 0.4923, y: 1 },
  { x: 0.2452, y: 0.7326 },
  { x: 0.1843, y: 0.6399 },
  { x: 0.1548, y: 0.5 , curve: { type: 'cubic', x1: 0.1653, y1: 0.5971, x2: 0.1548, y2: 0.5498 } },
  { x: 0.5, y: 0.1548, curve: { type: 'cubic', x1: 0.1548, y1: 0.3093, x2: 0.3093, y2: 0.1548 } },
  { x: 0.8452, y: 0.5, curve: { type: 'cubic', x1: 0.6907, y1: 0.1548, x2: 0.8452, y2: 0.3093 } },
  { x: 0.8114, y: 0.6483, curve: { type: 'cubic', x1: 0.8452, y1: 0.5531, x2: 0.8328, y2: 0.6033 } },
  { x: 0.9219, y: 0.7682 },
  { x: 1, y: 0.5, curve: { type: 'cubic', x1: 0.9713, y1: 0.6907, x2: 1, y2: 0.5988 } },
  { x: 0.5, y: 0, curve: { type: 'cubic', x1: 1, y1: 0.2238, x2: 0.7761, y2: 0 } },
];
const SWOOSH = [
  { x: 1, y: 0.6339 },
  { x: 0.6216, y: 0.6339 },
  { x: 0.2026, y: 0 },
  { x: 0, y: 0.2838 },
  { x: 0.4725, y: 1 },
  { x: 1, y: 1 },
  { close: true },
];
function scalePath(pts, x, y, w, h) {
  return pts.map((p) => {
    if (p.close) return p;
    const o = { x: x + p.x * w, y: y + p.y * h };
    if (p.curve) o.curve = { type: 'cubic', x1: x + p.curve.x1 * w, y1: y + p.curve.y1 * h, x2: x + p.curve.x2 * w, y2: y + p.curve.y2 * h };
    return o;
  });
}
/* opt: {color, accent, outline} - outline draws the mark as line art only */
function logoMark(sl, x, y, S, opt) {
  const o = opt || {};
  const col = o.color || DARK;
  const acc = o.accent || CORAL;
  const fill = o.outline ? { type: 'none' } : { color: col };
  const line = o.outline ? { color: col, width: 1 } : { width: 0, color: col };
  rect(sl, x, y, S, 0.1419 * S, o.outline ? null : col, { line });
  sl.addShape('custGeom', { x, y: y + 0.2318 * S, w: S, h: S, fill, line, points: scalePath(RING, 0, 0, S, S) });
  sl.addShape('custGeom', {
    x: x + 0.4608 * S, y: y + 0.8455 * S, w: 0.5395 * S, h: 0.3875 * S,
    fill: o.outline ? { type: 'none' } : { color: acc }, line,
    points: scalePath(SWOOSH, 0, 0, 0.5395 * S, 0.3875 * S),
  });
}
/* full lock-up: mark + "QUAZAR" word mark set to the right of it */
function wordmark(sl, x, y, S, opt) {
  const o = opt || {};
  logoMark(sl, x, y, S, o);
  sl.addText('QUAZAR', {
    x: x + 1.262 * S, y: y + 0.254 * S, w: 4.4 * S, h: 0.9 * S, wrap: false,
    fontFace: FONT_BOLD, bold: true, fontSize: 47 * S, charSpacing: 4.9 * S,
    color: o.wordColor || o.color || DARK, margin: 0, valign: 'top', align: 'left',
  });
}

/* chess-rook pictogram (Vision) */
function chessRook(sl, x, y, w, h) {
  const ln = { color: CORAL, width: 1.2 };
  sl.addShape('custGeom', { x, y, w, h: 0.28 * h, fill: { type: 'none' }, line: ln,
    points: [{ x: 0, y: 0 }, { x: 0, y: 0.28 * h }, { x: w, y: 0.28 * h }, { x: w, y: 0 }, { x: 0.68 * w, y: 0 },
      { x: 0.68 * w, y: 0.13 * h }, { x: 0.32 * w, y: 0.13 * h }, { x: 0.32 * w, y: 0 }, { close: true }] });
  sl.addShape('custGeom', { x: x + 0.16 * w, y: y + 0.28 * h, w: 0.68 * w, h: 0.5 * h, fill: { type: 'none' }, line: ln,
    points: [{ x: 0, y: 0 }, { x: 0.25 * w, y: 0.5 * h }, { x: 0.43 * w, y: 0.5 * h }, { x: 0.68 * w, y: 0 }] });
  rect(sl, x + 0.05 * w, y + 0.8 * h, 0.9 * w, 0.2 * h, null, { line: ln });
}
/* trophy pictogram (Mission) */
function trophy(sl, x, y, w, h) {
  const ln = { color: CORAL, width: 1.2 };
  sl.addShape('custGeom', { x: x + 0.2 * w, y, w: 0.6 * w, h: 0.62 * h, fill: { type: 'none' }, line: ln,
    points: [{ x: 0, y: 0 }, { x: 0.6 * w, y: 0 }, { x: 0.45 * w, y: 0.62 * h }, { x: 0.15 * w, y: 0.62 * h }, { close: true }] });
  sl.addShape('custGeom', { x, y: y + 0.05 * h, w: 0.24 * w, h: 0.35 * h, fill: { type: 'none' }, line: ln,
    points: [{ x: 0.24 * w, y: 0 }, { x: 0, y: 0 }, { x: 0.1 * w, y: 0.35 * h }, { x: 0.24 * w, y: 0.35 * h }] });
  sl.addShape('custGeom', { x: x + 0.76 * w, y: y + 0.05 * h, w: 0.24 * w, h: 0.35 * h, fill: { type: 'none' }, line: ln,
    points: [{ x: 0, y: 0 }, { x: 0.24 * w, y: 0 }, { x: 0.14 * w, y: 0.35 * h }, { x: 0, y: 0.35 * h }] });
  vline(sl, x + 0.5 * w, y + 0.62 * h, 0.22 * h, CORAL, 1.2);
  hline(sl, x + 0.28 * w, y + 0.92 * h, 0.44 * w, CORAL, 2);
}

/* small "document" pictogram used beside the cover footnotes */
function docIcon(sl, x, y, w, h, color) {
  sl.addShape('roundRect', { x, y, w, h, rectRadius: 0.05, fill: { type: 'none' }, line: { color, width: 1 } });
  [0.30, 0.45, 0.60].forEach((f, i) => {
    sl.addShape('ellipse', { x: x + 0.21 * w, y: y + f * h, w: 0.085 * w, h: 0.06 * h, fill: { color } });
    hline(sl, x + 0.35 * w, y + f * h + 0.03 * h, (i === 2 ? 0.24 : 0.43) * w, color, 0.75);
  });
}

/* ------------------------------------------- recurring page furniture bits */
function eyebrow(sl, x, label, w) { txt(sl, [x, 0.885, w || 5.729, 0.557], [P('t34', label)]); }
function bigTitle(sl, x, lines, h) {
  txt(sl, [x, 3.7, 4.147, h || 3.513], lines.map((l) => P('t48', l)));
}
function sideNote(sl, x, heading, body) {
  txt(sl, [x, 7.757, 5.248, 6.201], [P('t34', heading), GAP()].concat(body), { valign: 'bottom' });
}

/* ----------------------------------------------------------- copy deck */
const LOREM = {
  index: 'Dis doluptium cor sitate vendi ium de optur, quossitia cus, utem.At acesequi cones quam res reperciam, evendellabo. Nequatiis doluptatqui destia volore ant aut maximilia con ex eum et aligent eve atqui dest quossitia cus, utem ialitis.',
  visionA: 'Dis doluptium cor sitate vendi ium de optur, quossitia cus, utem.At acesequi cones quam res reperciam, evendellabo. Nequatiis doluptatqui destia volore ant aut maximilia con ex eum et aligent evelitis. ',
  visionB: 'At acesequi cones quam res reperciam, evendellabo. Nequatiis doluptatqui destia volore ant aut maximilia con ex eum et aligent evelitis. ',
  visionC: 'Epudit anime reium, omniscipsum binist ersped quam aut quas dolupta tinctat at iusanditate dignimo int volor aut dessist equia nulliciendit fuga. Et in cuptate quo qu mpores modi int dollupt urestiatesim quas entotatem quasima ximoluptur asis quoditas poribus sae consequi simagniq scitat apedi apiendunt quia ersped quas aut.',
  quatem1: 'Quatem iur? Velenes de plab il imenim fugitatent voloreh enisquis omnist ommodi dolupta desciatur, sediorio inctia voluptat odicim sitionectent asincit, core esciminit asimolut quia quatet essimus eum eatempedi dus dendige niminul luptae vendae sandem que que vollesenime..',
  quatem2: 'Quatem iur? Velenes de plab il imenim fugitatent voloreh enisquis omnist ommodi dolupta desciatur, sediorio inctia voluptat odicim sitionectent asincit, core esciminit asimolut quia quatet essimus eum eatempedi dus dendige niminul luptae Quatem iur? Velenes de plab il imenim fugitatent voloreh enisquis omnist.',
  headline: 'Iriducias nobis auten Iriducias nobis autendenihil iducias denihil iducias.',
  people: 'Henihilia atusdae esteri quia sumqui id modi reribus, non modi reribus, non nihilit provit esteri quia sum modi.',
  webGrid: 'Dis doluptium cor sitate vendi ium de optur, quossitia cus, utem.At acesequi cones quam res reperciam, evendellabo. Nequatiis quossitia cus, utem.At acesequi cones quam res reperciam, evendellabo. Nequatiis ',
};

const CONTACT = [
  P('t16', 'Address:', CORAL), P('t16', 'Quazar Studio'), P('t16', 'Main Avenue, 7th'), P('t16', 'New York, USA'),
  GAP(), GAP(), GAP(),
  P('t16', 'Phone/Fax:', CORAL), P('t16', 'Free Toll:\t+ 1 020 7800 800'), P('t16', 'Phone:\t+ 1 0800 123 123'),
  GAP(), P('t16', 'Fax: \t+ 1 0800 123 123'), GAP(), GAP(), GAP(),
  P('t16', 'Online:', CORAL), P('t16', 'info@quazar-studio.com'),
  P('t16', 'www.quazar-studio.com', GREY, { underline: { style: 'sng' } }),
];

const IMAGE_RULES = [P('t16', 'Requirements for ', CORAL), P('t16', 'Quazar Corporate Image System', CORAL),
  GAP(), P('t16', '\u2013', CORAL), GAP(),
  P('t16', '- desaturate colours'), P('t16', '- high contrast'), P('t16', '- sharp images'), P('t16', '- minimalistic look')];

/* ==================================================================== 01 */
/* Cover: rules across the head, oversized outlined Q on the right */
function slide01(sl) {
  hline(sl, 0, 0.833, 26.664, CORAL);
  hline(sl, 0, 2.333, 26.664, CORAL);
  sl.addShape('arc', { x: 16.583, y: 5.352, w: 10.268, h: 10.268, angleRange: [302, 232], line: { color: CORAL, width: 1 } });
  sl.addShape('arc', { x: 18.171, y: 6.94, w: 7.09, h: 7.09, angleRange: [302, 232], line: { color: CORAL, width: 1 } });
  sl.addShape('custGeom', {
    x: 21.313, y: 11.64, w: 5.537, h: 3.36, fill: { type: 'none' }, line: { color: CORAL, width: 1 },
    points: scalePath([{ x: 1, y: 0.7512 }, { x: 0.6216, y: 0.7512 }, { x: 0.2027, y: 0 }, { x: 0, y: 0.3363 }, { x: 0.3695, y: 1 }], 0, 0, 5.537, 3.36),
  });
  rect(sl, 26.039, 0, 0.634, 15, DARK);
  logoMark(sl, 1.955, 1.36, 0.361);
  txt(sl, [2.735, 1.384, 2.744, 0.48], [P('t28', 'Quazar', DARK)], { valign: 'middle' });
  txt(sl, [1.943, 4.406, 10.772, 1.813], [
    P('t48', 'Brand Manual'),
    PR('t48', [['& Guidelines ', DARK], ['2021', CORAL]]),
  ]);
  txt(sl, [1.943, 7.999, 6.807, 1.49], [
    PR('t16', [['The new Brand Guidelines for ', DARK], ['Quazar Ltd. & Co.', CORAL]]),
    P('t16', 'to represent your identity and give a consistent visual look.', DARK),
    P('t16', '\u2014', DARK),
    P('t16', 'Volume VII', CORAL),
  ]);
  [[1.944, '01. Headline here'], [7.713, '02. Headline here']].forEach(([x, head]) => {
    docIcon(sl, x + 0.004, 10.962, 0.293, 0.407, CORAL);
    hline(sl, x + 0.009, 11.785, 5.31, CORAL);
    txt(sl, [x, 12.052, 5.315, 1.056], [P('t16', head, CORAL), P('t16', LOREM.headline, DARK)]);
  });
}

/* ================================================================= 02-03 */
/* Index pages: a lettered card per chapter */
function indexCard(sl, x, letter, title, sub, page) {
  txt(sl, [x, 3.716, 5.717, 1.421], [P('t80', letter)]);
  txt(sl, [x, 6.895, 5.717, 4.722], [P('t28', title), P('t22', sub), GAP(), P('t16', LOREM.index)]);
  txt(sl, [x, 12.384, 5.717, 1.194], [P('t80', page)]);
}
function slide02(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 11.175, 'Index.');
  txt(sl, [1.111, 0.885, 3.267, 2.41], [P('t34', 'Brand Guidelines'), GAP(), P('t34', '2021', CORAL)]);
  logoMark(sl, 1.122, 4.758, 0.496);
  txt(sl, [1.111, 7.233, 2.902, 6.437], CONTACT);
  indexCard(sl, 11.181, 'a', 'Brand Identity', 'Our Commitment', '04');
  indexCard(sl, 19.774, 'b', 'Logo Guidelines', 'Rules and Guides', '12');
}
function slide03(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 2.581, 'Index.');
  indexCard(sl, 2.594, 'c', 'Colors and Typography ', 'Our Look', '12');
  indexCard(sl, 11.181, 'd', 'Corporate Stationery', 'Web and Print Grids', '18');
  indexCard(sl, 19.774, 'e', 'Corporate Images', 'Images and Application', '24');
}

/* ==================================================================== 04 */
function slide04(sl) {
  /* narrowed slightly from the reference 4.61 so the headline breaks after
     "Identity." exactly as the source deck does */
  txt(sl, [1.111, 0.885, 4.2, 5.176], [
    PR('t48', [['Corporate Identity. Have a look when', DARK], [' image meets design.', CORAL]]),
  ]);
  const body = [GAP(), GAP(), P('t16', LOREM.visionA), GAP(), P('t16', LOREM.visionB), GAP(), P('t16', LOREM.visionC)];
  chessRook(sl, 11.206, 5.683, 0.308, 0.447);
  txt(sl, [11.192, 6.897, 5.699, 6.888], [P('t28', 'Vision')].concat(body));
  trophy(sl, 19.784, 5.682, 0.419, 0.405);
  txt(sl, [19.784, 6.897, 5.699, 6.888], [P('t28', 'Mission')].concat(body));
}

/* ==================================================================== 05 */
function slide05(sl) {
  eyebrow(sl, 1.118, 'About us.');
  logoMark(sl, 1.122, 3.836, 0.496);
  txt(sl, [1.109, 6.935, 3.716, 6.967], [
    P('t28', 'Hard facts'), P('t28', '-'),
    P('t16', 'Name: ', CORAL), P('t22', 'Quazar Ltd.'), GAP(), GAP(),
    P('t16', 'Founded:', CORAL), P('t22', '2009'), GAP(), GAP(),
    P('t16', 'Staff:', CORAL), P('t22', '79'), GAP(), GAP(),
    P('t16', 'Customers:', CORAL), P('t22', '165'),
  ]);
  txt(sl, [10.331, 3.695, 15.225, 0.893], [P('t48', 'Make your brand a lovely one.')]);
  hline(sl, 10.331, 5.619, 15.218, CORAL);
  txt(sl, [10.331, 6.935, 15.211, 1.578], [P('t28', 'Quas dolentiur? Et ut mo ent viorisit nobis autendenihil iducias adit urbix quostrum accus et aut dolentiur? Et ut mo ent viorisit et rem. Quas dolentiur? Et ut mo ent viorisit nobis autendenihil iducias adit urbix.')]);
  const colA = 'Quas dolentiur? Et ut mo beate nobis autendenihil iducias adit quostrum accus et aut et rem quidipsa quam beate nobis autendenihil iducias adit quostrum accus et aut et rem quidipsa quam recus.';
  const colB = 'Quas dolentiur? Et ut mo beate nobis urbi autendenihil iducias nobis autendenihil as iducias adit accus et aut et rem quidipsa quam recus. Quas dolentiur? Ints Et ut mo nobis autendenihil iducias adit quostrum accus et aut et rem quidipsa. Et ut mo beate nobis Quas dolentiur?';
  const colC = 'Quas dolentiur? Et ut mo beate nobis urbi autendenihil iducias nobis autendenihil as iducias adit quostrum accus et aut et rem quidipsa quam recus. Quas dolentiur? Ints Et ut mo nobis autendenihil iducias adit quostrum accus et aut et rem quidipsa. Et ut mo beate nobis Quas dolentiur?';
  const colD = 'Ints Et ut mo nobis autendenihil iducias adit quostrum accus et aut et rem quidipsa. Et ut mo beate nobis Quas dolentiur?';
  txt(sl, [10.331, 9.617, 7.365, 4.273], [P('t16', colA), GAP(), P('t16', colB), P('t16', colA)]);
  txt(sl, [18.183, 9.617, 7.365, 4.273], [P('t16', colC), GAP(), P('t16', colD), P('t16', colA)]);
}

/* ==================================================================== 06 */
function slide06(sl) {
  imageBox(sl, -0.003, -0.022, 25.574, 15.044);
  rules(sl, [RULE_B, RULE_C, RULE_A]);
  rect(sl, -0.003, 4.622, 13.691, 9.272, WHITE);
  logoMark(sl, 1.122, 5.848, 0.496);
  txt(sl, [1.122, 7.774, 9.235, 5.116], [
    P('t48', 'Design is not just what it looks like and feels like. Design is how it works.', DARK, { fontFace: FONT_LIGHT }),
    P('t48', 'Design is not just what it looks like and feels like. ', DARK, { fontFace: FONT_LIGHT }),
    P('t48', 'Design is how it works.', CORAL, { fontFace: FONT_LIGHT }),
  ]);
}

/* ==================================================================== 07 */
function slide07(sl) {
  eyebrow(sl, 1.118, 'Our people.');
  eyebrow(sl, 11.201, 'About our studio.');
  [3.53, 9.296].forEach((y) => imageBox(sl, 1.126, y, 5.702, 2.601));
  [6.482, 12.133].forEach((y) => txt(sl, [1.11, y, 5.702, 1.921], [
    P('t22', 'Sarah Saunderson'), P('t16', 'CEO / Company President', CORAL), GAP(), P('t16', LOREM.people),
  ]));
  txt(sl, [11.183, 3.695, 14.373, 0.893], [P('t48', 'An experienced brand identity studio.')]);
  hline(sl, 11.183, 5.619, 14.362, CORAL);
  txt(sl, [11.197, 6.935, 14.345, 6.958], [
    P('t16', 'Uptam eossum que aut que volor simentur, consed quis des venim  ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur et eatusantus se con consequ iatur? Optatur rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur quatem ipienis? ', CORAL),
    GAP(), P('t28', 'We create strong brands.'), GAP(),
    P('t16', 'Uptam eossum que aut que volor simentur, consed quis des venim ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur quatem fu rernatus ad etur as ent ut quis rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ fugitinctur ipienis et eatusantus se con consequ gitinctur quatem fu rernatus ad etur as ent ut quis ipienis?'),
    GAP(), P('t28', 'Build your brand in an successful way.'), GAP(),
    P('t16', 'Uptam eossum que aut que volor simentur, consed quis des venim ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ iatur? Optatur rernatus ad etur quatem fu rernatus ad etur as ent ut quis rernatus ad etur as ent ut quis quatem fugitinctur ipienis et eatusantus se con consequ fugitinctur ipienis et eatusantus se con consequ gitinctur ipienis quatem fu rernatus ad etur as ent ut quis quatem fu rernatus ad etur as ent ut quis?'),
  ]);
}

/* ==================================================================== 08 */
function slide08(sl) {
  eyebrow(sl, 1.118, 'Quazar identity.');
  txt(sl, [1.111, 3.7, 4.147, 5.176], [P('t48', 'Our brand and corporate identity.')]);
  imageBox(sl, 8.309, -0.022, 18.355, 15.044);
  logoMark(sl, 1.126, 13.134, 0.615);
}

/* ================================================================= 09-10 */
/* three-pillar pages */
function pillar(sl, x, head, sub, claim) {
  txt(sl, [x, 3.7, 4.147, 10.185], [
    P('t48', head), GAP(), P('t28', sub), P('t28', claim, CORAL), GAP(),
    P('t16', LOREM.quatem1), GAP(), P('t16', LOREM.quatem2),
  ]);
}
function slide09(sl) {
  rules(sl, [RULE_B, RULE_C, RULE_A]);
  eyebrow(sl, 2.565, 'Quazar identity.');
  pillar(sl, 2.575, 'Identity ', 'Corporate Identity', 'How we are and want to be.');
  pillar(sl, 11.19, 'Design ', 'Corporate Design', 'How we wanna look and feel.');
  pillar(sl, 19.792, 'Culture ', 'Corporate Culture', 'How we care about us and other.');
}
function slide10(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 1.118, 'Design goals.');
  txt(sl, [1.111, 3.7, 4.147, 5.176], [P('t48', 'Project brief, goals and content.')]);
  pillar(sl, 11.19, 'Brief.', 'Project Description', 'How we wanna be in the future.');
  pillar(sl, 19.792, 'Goals. ', 'Project Goals', 'What are our goals to win?');
}

/* ==================================================================== 11 */
function slide11(sl) {
  imageBox(sl, 1.143, -0.022, 17.218, 13.889);
  vline(sl, RULE_C, 0, SLIDE_H, DARK);
  /* oversized outlined Q bleeding off the right edge */
  logoMark(sl, 22.139, -0.425, 10.734, { outline: true, color: CORAL });
  /* pale blue opening quote marks */
  sl.addShape('custGeom', {
    x: 17.103, y: 3.063, w: 3.938, h: 2.473, fill: { color: LIGHT }, line: { width: 0, color: LIGHT },
    points: scalePath([
      { x: 0.5945, y: 0 }, { x: 0.2724, y: 0 }, { x: 0.2721, y: 0.4587 }, { x: 0, y: 1 }, { x: 0.3221, y: 1 }, { x: 0.5945, y: 0.4548 }, { x: 0.5945, y: 0 },
      { x: 1, y: 0, moveTo: true }, { x: 0.6779, y: 0 }, { x: 0.6774, y: 0.4587 }, { x: 0.4055, y: 1 }, { x: 0.7276, y: 1 }, { x: 1, y: 0.4548 }, { x: 1, y: 0 },
    ], 0, 0, 3.938, 2.473),
  });
  eyebrow(sl, 19.785, 'Design Principle.');
  txt(sl, [19.766, 3.7, 4.306, 6.106], [
    P('t48', 'The design process, '),
    P('t48', 'at its best, integrates the aspirations of art, science, and culture.'),
  ]);
  hline(sl, 19.787, 11.933, 3.067, CORAL);
  txt(sl, [19.776, 12.172, 4.147, 0.815], [P('t22', 'Michael Anderson'), P('t16', 'Creative Director', CORAL)]);
}

/* ==================================================================== 12 */
/* Corporate logo + clearspace construction drawing */
function slide12(sl) {
  vline(sl, RULE_B, 0, SLIDE_H, DARK);
  eyebrow(sl, 1.118, 'Corporate Logo.');
  bigTitle(sl, 1.111, ['Corporate logo and clearspace.']);
  sideNote(sl, 1.109, 'The corporate logo.', [
    P('t16', 'The Quazar Masterbrand or Corporate Logo comprises two elements, the logo symbol and logo type. The Logo Symbol is a powerful image evoking the culture of design services - the connection between the strength of communication and the different points that influence.'),
    GAP(),
    P('t16', 'The Logo Type has been carefully chosen for its modern and yet refined, highly legible style, which has been further enhanced by the use of lower case letters. The typeface of Quazar is Cormorant Garamond and has also been chosen to compliment and balance perfectly with the logo symbol.'),
  ]);
  txt(sl, [11.187, 3.007, 5.729, 0.557], [P('t16', 'Quazar Full Logo', CORAL)]);
  [3.851, 7.659].forEach((y) => hline(sl, 11.188, y, 14.38, CORAL, 1.5, 'sysDot'));
  [5.188, 6.642].forEach((y) => hline(sl, 15.217, y, 10.339, CORAL, 1.5, 'sysDot'));
  wordmark(sl, 11.197, 3.842, 3.106);

  /* clearspace diagram */
  txt(sl, [11.187, 8.937, 5.729, 0.557], [P('t16', 'Full Logo Clearspace', CORAL)]);
  const gx = 11.193, gy = 9.789, gw = 8.597, gh = 4.113;
  const HAIR = 'CFD0D2';
  for (let i = 0; i <= 10; i++) hline(sl, gx, gy + (gh / 10) * i, gw, HAIR, 0.75, 'solid');
  for (let i = 0; i <= 21; i++) vline(sl, gx + (gw / 21) * i, gy, gh, HAIR, 0.75, 'solid');
  rect(sl, gx, gy, gw, gh, null, { line: { color: HAIR, width: 0.75 } });
  hline(sl, 11.978, 11.026, 6.778, CORAL, 1.5);
  hline(sl, 11.978, 12.669, 6.778, CORAL, 1.5);
  [[11.573, 10.203], [18.559, 10.203], [11.573, 12.677], [18.559, 12.666]].forEach(([bx, by]) => {
    rect(sl, bx, by, 0.812, 0.812, WHITE, { line: { color: CORAL, width: 1 } });
    txt(sl, [bx + 0.034, by + 0.264, 0.744, 0.385], [PR('t22', [
      ['1', CORAL, { superscript: true }], ['/', CORAL], ['2', CORAL, { subscript: true }], [' x', CORAL],
    ])]);
  });
  vline(sl, 12.103, 11.125, 1.432, CORAL, 1);
  txt(sl, [11.607, 11.626, 0.349, 0.385], [P('t22', 'x', WHITE)]);
  rect(sl, 11.654, 11.722, 0.256, 0.256, CORAL);
  wordmark(sl, 12.379, 11.032, 1.336);
}

/* ==================================================================== 13 */
/* Logo on the four approved backgrounds */
function slide13(sl) {
  vline(sl, RULE_C, 0, SLIDE_H, DARK);
  rect(sl, 0, 0.006, 18.35, 10, DARK);
  rect(sl, 18.367, 0.006, 8.305, 5, LIGHT);
  rect(sl, 18.358, 5.006, 8.314, 5, PALE);
  rect(sl, 18.366, 10.003, 8.306, 5, WHITE);
  wordmark(sl, 5.111, 3.93, 1.751, { color: WHITE, accent: CORAL });
  wordmark(sl, 5.111, 11.427, 1.751, { outline: true, color: DARK });
  wordmark(sl, 20.777, 2.048, 0.746, { color: DARK, accent: WHITE });
  wordmark(sl, 20.777, 7.042, 0.746, { color: GREY, accent: GREY, wordColor: GREY });
  wordmark(sl, 20.777, 12.045, 0.746, { color: DARK, accent: CORAL });
}

/* ==================================================================== 14 */
function slide14(sl) {
  vline(sl, RULE_B, 0, SLIDE_H, DARK);
  eyebrow(sl, 1.118, 'Corporate Typography.');
  bigTitle(sl, 1.111, ['Corporate Typography and Usage.']);
  sl.addText('A2', { x: 1.05, y: 6.95, w: 10, h: 8, fontFace: FONT_LIGHT, fontSize: 503, color: DARK, margin: 0, valign: 'top', align: 'left' });
  eyebrow(sl, 11.201, 'The Font.');
  txt(sl, [11.183, 3.695, 14.373, 2.481], [
    P('t48', 'Roboto Font Family .'), GAP('t28'), P('t28', 'Font Description'), P('t28', 'A modern sans serif font for free.', CORAL),
  ]);
  hline(sl, 11.176, 7.5, 14.373, CORAL);
  const R = (t) => ['Roboto', CORAL];
  txt(sl, [11.197, 9.262, 14.345, 4.779], [
    PR('t16', [['Roboto', CORAL], [' has a dual nature. It has a mechanical skeleton and the forms are largely geometric. At the same time, the font features friendly and open curves. While some grotesks distort their letterforms to force a rigid rhythm, ', GREY], ['Roboto', CORAL], [' doesn\u2019t compromise, allowing letters to be settled into their natural width. This makes for a more natural reading rhythm more commonly found in humanist and serif types. It includes Thin, Light, Regular, Medium, Bold and Black weights with matching oblique styles rather than true italics. It also includes condensed styles in Light, Regular and Bold, also with matching oblique designs.', GREY]]),
    GAP(),
    P('t16', 'The font was designed entirely in-house by Christian Robertson who previously had released an expanded Ubuntu-Title font through his personal type foundry Betatype. The font was officially made available for free download in 2012, on the newly launched Android Design website.'),
    GAP(),
    PR('t16', [['Roboto', CORAL], [' is a neo-grotesque sans-serif typeface family developed by Google as the system font for its mobile operating system Android, and released in 2011 for Android 4.0 "Ice Cream Sandwich". The entire font family has been licensed under the Apache license. In 2014, ', GREY], ['Roboto', CORAL], [' was redesigned for Android 5.0 "Lollipop".', GREY]]),
  ]);
}

/* ==================================================================== 15 */
function fontSpecimen(sl, x, w, name, face) {
  txt(sl, [x, 3.7, w, 10.185], [
    P('t48', name, DARK, { fontFace: face }), GAP(),
    P('t28', 'A pixel-perfect font ', DARK, { fontFace: face }),
    P('t28', 'for our design.', CORAL, { fontFace: face }),
    GAP(), GAP(), GAP(), GAP(), GAP(),
    P('t16', 'Font Letter', CORAL),
    P('t28', 'ABCDEFGHIJKLMNOPQRSTUVWX', DARK, { fontFace: face, charSpacing: 57.95, align: 'justify' }),
    GAP(), GAP(), GAP(), GAP(),
    P('t16', 'Font Figures', CORAL),
    P('t28', '1234567890', DARK, { fontFace: face, charSpacing: 82.5, align: 'justify' }),
  ]);
}
function slide15(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 1.118, 'Corporate Typography.');
  logoMark(sl, 1.122, 3.855, 1.469);
  eyebrow(sl, 11.201, 'Style a.');
  eyebrow(sl, 19.803, 'Style a.');
  fontSpecimen(sl, 11.19, 7.157, 'Roboto Light ', FONT_LIGHT);
  fontSpecimen(sl, 19.792, 6.861, 'Roboto Bold', FONT_BOLD);
}

/* ==================================================================== 16 */
const TONES = [100, 80, 60, 40, 20, 10];
function colorColumn(sl, x, color, name, role) {
  rect(sl, x, 3.842, 3.978, 1.373, color);
  txt(sl, [x - 0.016, 5.529, 3.978, 0.893], [P('t22', name), P('t16', role, CORAL)]);
  txt(sl, [x - 0.016, 7.283, 4.011, 0.434], [P('t16', 'Color Tones', DARK)]);
  TONES.forEach((pct, i) => {
    const y = 7.968 + i * 1.0424;
    rect(sl, x, y, 3.978, 0.261, null, { fill: { color, transparency: 100 - pct } });
    txt(sl, [x - 0.016, y + 0.383, 3.978, 0.434], [P('t16', pct + '%', CORAL)]);
  });
}
function slide16(sl) {
  vline(sl, RULE_B, 0, SLIDE_H, DARK);
  eyebrow(sl, 1.118, 'Primary color system.');
  bigTitle(sl, 1.111, ['Corporate ', 'Color', 'System.']);
  sideNote(sl, 1.109, 'Brand colors and usage.', [
    P('t16', 'Color plays an important role in the Quazar Studio corporate identity program. The colors below are recommendations for various media. '),
    GAP(),
    P('t16', 'A palette of primary colors has been developed, which comprise the \u201cOne Voice\u201d color scheme. Consistent use of these colors will contribute to the cohesive and harmonious look of the Quazar brand identity across all relevant media.'),
    GAP(),
    P('t16', 'Check with your designer or printer when using the corporate colors and that they will be always be consistent.'),
  ]);
  colorColumn(sl, 11.201, DARK, 'Dark Blue', 'Primary Color A.');
  colorColumn(sl, 16.389, CORAL, 'Coral Blue', 'Primary Color B.');
  colorColumn(sl, 21.577, LIGHT, 'Light Blue', 'Primary Color C.');
}

/* ==================================================================== 17 */
const CODES = ['CMYK\t:\tC000\tM020\tY100\tK000', 'Pantone\t:\t297C', 'HKS \t:\t02K',
  'RGB\t         : \tR000\tG000\tB000', 'Web  \t:\t#000000'];
function codePanel(sl, x, w, bg, name, role, headColor, bodyColor) {
  rect(sl, x, 3.842, w, 11.167, bg);
  txt(sl, [x + 1.454, 5.09, 5.248, 1.108], [P('t34', name, headColor), P('t22', role, headColor)]);
  txt(sl, [x + 1.454, 9.048, 5.248, 4.599], [
    P('t22', 'Primary Color B', bodyColor), P('t22', 'Coral Blue', bodyColor), GAP(), GAP(), GAP(),
    P('t16', 'Color Codes', bodyColor), GAP(),
  ].concat(CODES.map((c) => P('t16', c, headColor))));
}
function slide17(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 1.118, 'Color codes.');
  codePanel(sl, 1.127, 8.626, DARK, 'Dark Blue', 'Primary Color A.', WHITE, CORAL);
  codePanel(sl, 9.754, 8.584, CORAL, 'Coral Blue', 'Primary Color B.', WHITE, DARK);
  codePanel(sl, 18.37, 8.304, LIGHT, 'Light Blue', 'Primary Color C.', DARK, CORAL);
}

/* ==================================================================== 18 */
function slide18(sl) {
  eyebrow(sl, 1.118, 'Print communication.');
  bigTitle(sl, 1.111, ['Corporate Stationery', 'System.']);
  sideNote(sl, 1.109, 'Stationery system.', [
    P('t16', 'Stationery is a primary means of communication and it is essential that every application be a consistent reflection of our corporate identity. '),
    GAP(),
    P('t16', 'There is only one approved design format for all corporate and business unit stationery, although there are slight variations in size and content for different regions of  the world. This section illustrates approved layouts for standard U.S. business stationery. '),
    GAP(),
    P('t16', 'It includes specifications for typography, color, printing method, paper stock.'),
  ]);
  rect(sl, 8.31, -0.008, 15.805, 15.017, DARK);
  imageBox(sl, 22.823, -0.022, 3.838, 15.044);
  txt(sl, [9.741, 3.7, 4.147, 4.655], [
    P('t22', 'Letterhead', CORAL), GAP('t16'),
    P('t16', 'This shows the approved layout with the primary elements of the Quazar stationery system for the frontside of letterheads.', WHITE),
    GAP(), P('t16', 'Usage:', CORAL),
    P('t16', 'The corporate letterhead will be used for all official external communication of Quazar company. ', WHITE),
  ]);
  txt(sl, [9.741, 9.459, 4.147, 4.445], [
    P('t22', 'Parameter', WHITE), GAP(),
    P('t16', 'Dimensions:', CORAL), P('t16', 'DIN A4', WHITE), GAP(),
    P('t16', 'Paper Weight:', CORAL), P('t16', '100 g/m2', WHITE), GAP(),
    P('t16', 'Print:', CORAL), P('t16', 'CMYK Offset', WHITE),
  ]);
  letterheadMock(sl, 15.459, 1.651, 8.656, 12.243);
}
/* A4 letterhead mock-up: header band, logo, address block and grey text lines */
function letterheadMock(sl, x, y, w, h) {
  paperBox(sl, x, y, w, h, WHITE);
  rect(sl, x + 0.6 * w, y, 0.4 * w, 0.115 * h, 'F1F2F2');
  wordmark(sl, x + 0.07 * w, y + 0.026 * h, 0.062 * w, { color: DARK, accent: CORAL });
  logoMark(sl, x + 0.885 * w, y + 0.045 * h, 0.05 * w);
  txt(sl, [x + 0.6 * w, y + 0.12 * h, 0.36 * w, 0.1 * h], [
    P('t10', 'Main Avenue, 7th', GREY), P('t10', 'New York, United States', GREY), GAP('t10'),
    P('t10', 'T. 01 23 4567 8901', GREY), P('t10', 'F. 01 23 4567 8901', GREY), GAP('t10'),
    P('t10', 'info@quazar-studio.com', CORAL), P('t10', 'www.quazar-studio.com', CORAL),
  ], { align: 'right' });
  txt(sl, [x + 0.08 * w, y + 0.185 * h, 0.4 * w, 0.12 * h], [
    P('t10', 'July, 2ne 2017', DARK, { bold: true }), GAP('t10'),
    P('t10', 'Mr. Max Jefferson', GREY), P('t10', 'Creative Director', GREY), P('t10', 'London, GB', GREY),
  ]);
  hline(sl, x + 0.55 * w, y + 0.235 * h, 0.05 * w, DARK, 0.75);
  hline(sl, x + 0.55 * w, y + 0.4 * h, 0.05 * w, DARK, 0.75);
  txt(sl, [x + 0.08 * w, y + 0.265 * h, 0.62 * w, 0.05 * h], [P('t10', 'Dear Sir and Madam,', DARK, { bold: true })]);
  for (let b = 0; b < 3; b++) {
    for (let i = 0; i < 5; i++) {
      const yy = y + (0.31 + b * 0.09) * h + i * 0.014 * h;
      hline(sl, x + 0.08 * w, yy, (i === 4 ? 0.35 : 0.62) * w, 'BDBEC0', 0.5);
    }
  }
  txt(sl, [x + 0.08 * w, y + 0.575 * h, 0.3 * w, 0.05 * h], [P('t10', 'Best Regards', GREY)]);
  txt(sl, [x + 0.08 * w, y + 0.625 * h, 0.3 * w, 0.07 * h], [P('t10', 'Maxwell Doe', DARK, { bold: true }), P('t10', 'Director', GREY)]);
  rect(sl, x + 0.75 * w, y + 0.415 * h, 0.19 * w, 0.055 * h, null, { line: { color: DARK, width: 0.5 } });
  txt(sl, [x + 0.08 * w, y + 0.845 * h, 0.4 * w, 0.04 * h], [PR('t10', [['Quazar Studio. ', DARK], ['Brand Builder', CORAL]])]);
  txt(sl, [x + 0.78 * w, y + 0.755 * h, 0.2 * w, 0.09 * h], [P('t16', 'quazar-'), P('t16', 'design.')]);
}

/* ==================================================================== 19 */
function slide19(sl) {
  rules(sl, [RULE_C, RULE_D]);
  imageBox(sl, 0.008, -0.022, 3.838, 15.044);
  rect(sl, 3.845, -0.008, 1.449, 15.017, DARK);
  eyebrow(sl, 9.397, 'Business Cards.');
  txt(sl, [9.433, 1.931, 4.011, 0.434], [P('t16', 'Front Side', CORAL)]);
  txt(sl, [9.433, 8.403, 4.011, 0.434], [P('t16', 'Back Side', CORAL)]);
  cardFront(sl, 9.442, 2.603, 7.444, 4.82);
  cardBack(sl, 9.441, 9.078, 7.445, 4.82);
  txt(sl, [19.773, 3.7, 4.147, 4.865], [
    P('t22', 'Business Cards'), GAP(),
    P('t16', 'This shows the approved layouts with the primary elements of the Quazar stationery system for business cards.'),
    GAP(), P('t16', 'Usage:', CORAL),
    P('t16', 'Corporate business cards will be used for all official contact and external communication of Quazar Studio company. '),
  ]);
  txt(sl, [19.773, 9.459, 4.147, 4.445], [
    P('t22', 'Parameter'), GAP(),
    P('t16', 'Dimensions:', CORAL), P('t16', '85 x 55 mm'), GAP(),
    P('t16', 'Paper Weight:', CORAL), P('t16', '350 g/m2'), GAP(),
    P('t16', 'Print:', CORAL), P('t16', 'CMYK Offset'),
  ]);
}
function cardFront(sl, x, y, w, h) {
  rect(sl, x, y, w, h, WHITE, { line: { color: DARK, width: 1 } });
  wordmark(sl, x + 0.055 * w, y + 0.16 * h, 0.075 * w, { color: DARK, accent: CORAL });
  rect(sl, x, y + 0.735 * h, w, 0.265 * h, DARK);
  txt(sl, [x + 0.05 * w, y + 0.82 * h, 0.5 * w, 0.1 * h], [P('t10', 'www.quazar-studio.com', WHITE)]);
}
function cardBack(sl, x, y, w, h) {
  rect(sl, x, y, w, h, WHITE, { line: { color: DARK, width: 1 } });
  rect(sl, x, y, 0.055 * w, h, DARK);
  rect(sl, x, y + 0.79 * h, 0.055 * w, 0.09 * h, CORAL);
  txt(sl, [x + 0.11 * w, y + 0.16 * h, 0.6 * w, 0.14 * h], [P('t22', 'Quazar Studio.')]);
  txt(sl, [x + 0.11 * w, y + 0.53 * h, 0.28 * w, 0.2 * h], [
    P('t10', 'Mr. Max Jefferson', DARK, { bold: true }), P('t10', 'Creative Director', GREY)]);
  txt(sl, [x + 0.32 * w, y + 0.53 * h, 0.35 * w, 0.24 * h], [
    P('t10', 'Quazar Studio', DARK, { bold: true }), P('t10', 'Main Avenue, 7th', GREY), P('t10', 'New York, USA', GREY)]);
  txt(sl, [x + 0.11 * w, y + 0.74 * h, 0.28 * w, 0.16 * h], [
    P('t10', 'T. 01 23 4567 8901', GREY), P('t10', 'F. 01 23 4567 8901', GREY)]);
  txt(sl, [x + 0.32 * w, y + 0.74 * h, 0.35 * w, 0.16 * h], [
    P('t10', 'info@quazar-studio.com', GREY), P('t10', 'www.quazar-studio.com', CORAL)]);
  logoMark(sl, x + 0.85 * w, y + 0.79 * h, 0.045 * w);
}

/* ==================================================================== 20 */
/* Correct logo placement: three grid sheets with the mark in the allowed spots */
function gridSheet(sl, x, y, w, h, cols, rowsN) {
  paperBox(sl, x, y, w, h, WHITE);
  for (let i = 1; i < cols; i++) vline(sl, x + (w / cols) * i, y + 0.03 * h, h * 0.94, 'D8D9DA', 0.75);
  for (let i = 1; i < rowsN; i++) hline(sl, x + 0.03 * w, y + (h / rowsN) * i, w * 0.94, 'D8D9DA', 0.75);
}
function slide20(sl) {
  vline(sl, RULE_B, 0, SLIDE_H, DARK);
  eyebrow(sl, 1.118, 'Visual basics.');
  bigTitle(sl, 1.111, ['Correct ', 'Logo Placement']);
  sideNote(sl, 1.109, 'Explanation.', [
    P('t16', 'To place the Quazar logo in the correct way please use one of the approved styles that are shown on the right. To place the Quazar logo in other ways is not allowed.'),
  ]);
  rect(sl, 16.892, -0.008, 9.788, 15.017, DARK);
  gridSheet(sl, 11.141, 1.648, 8.653, 12.25, 6, 12);
  logoMark(sl, 11.77, 2.263, 0.81);
  wordmark(sl, 17.691, 13.028, 0.241, { color: DARK, accent: DARK });
  gridSheet(sl, 21.573, 1.651, 3.989, 5.642, 5, 10);
  wordmark(sl, 21.859, 1.942, 0.184, { color: DARK, accent: CORAL });
  gridSheet(sl, 21.573, 8.23, 3.989, 5.642, 5, 10);
  wordmark(sl, 24.798, 8.847, 0.184, { color: DARK, accent: CORAL, rotate: 270 });
}

/* ==================================================================== 21 */
function slide21(sl) {
  imageBox(sl, 0, 0.006, 19.775, 15);
  rect(sl, 19.766, -0.008, 6.914, 15.017, DARK);
  rect(sl, 1.132, 3.012, 14.954, 10.882, WHITE);
  logoMark(sl, 2.473, 4.311, 0.986);
  txt(sl, [2.473, 7.268, 10.855, 5.623], [
    P('t80', 'The height of cultivation always runs to simplicity.', DARK),
    P('t80', 'Design is your answer.', CORAL),
  ]);
}

/* ================================================================= 22-23 */
/* Browser mock-up used for the web grid pages. `filled` = content version. */
function browserFrame(sl, x, y, w, h) {
  rect(sl, x, y, w, h, WHITE, { line: { color: DARK, width: 1 } });
  rect(sl, x, y, w, 0.658, 'EFEFEF', { line: { color: DARK, width: 1 } });
  [0, 1, 2].forEach((i) => sl.addShape('ellipse', {
    x: x + 0.306 + i * 0.332, y: y + 0.266, w: 0.188, h: 0.188,
    fill: { color: i === 0 ? DARK : GREY }, line: { width: 0, color: GREY },
  }));
}
function webPage(sl, filled) {
  browserFrame(sl, 11.204, 1.144, 14.035, 12.744);
  const barFill = filled ? DARK : WHITE;
  const barLine = filled ? DARK : DARK;
  rect(sl, 11.523, 2.101, 10.233, 0.386, barFill, { line: { color: barLine, width: 1 } });
  rect(sl, 21.926, 2.101, 2.996, 0.386, barFill, { line: { color: barLine, width: 1 } });
  if (filled) txt(sl, [11.594, 2.05, 4.147, 0.408], [P('t16', 'www.quazar-limited.com', CORAL, { underline: { style: 'sng' } })], { valign: 'middle' });
  /* navigation strip */
  rect(sl, 11.523, 2.82, 13.404, 1.107, filled ? DARK : WHITE, { line: { color: filled ? DARK : DARK, width: 1 } });
  wordmark(sl, 11.895, 3.209, 0.265, { color: filled ? WHITE : DARK, accent: CORAL });
  txt(sl, [14.317, 3.133, 1.349, 0.408], [P('t16', 'Newsletter', filled ? WHITE : GREY)], { valign: 'middle', align: 'right' });
  rect(sl, 15.78, 3.179, 8.307, 0.386, WHITE, { line: { color: DARK, width: 1 } });
  if (filled) txt(sl, [24.2, 3.15, 0.5, 0.4], [P('t16', 'x', WHITE)]);
  /* hero + two side panels */
  if (filled) {
    imageBox(sl, 11.523, 4.256, 8.443, 4.474);
    rect(sl, 14.583, 5.321, 2.333, 2.354, DARK);
    logoMark(sl, 15.273, 5.912, 0.944, { color: WHITE, accent: CORAL });
    [[20.369, 4.256, 'Brand Studio'], [20.369, 6.664, 'About us']].forEach(([px, py, label]) => {
      imageBox(sl, px, py, 4.56, 2.069);
      txt(sl, [px + 0.596, py + 0.344, 3.369, 1.382], [P('t28', label)], { valign: 'middle', align: 'center' });
    });
  } else {
    rect(sl, 11.523, 4.256, 8.443, 4.474, WHITE, { line: { color: DARK, width: 1 } });
    rect(sl, 14.583, 5.321, 2.333, 2.354, WHITE, { line: { color: DARK, width: 1 } });
    logoMark(sl, 15.273, 5.912, 0.944, { outline: true, color: DARK });
    rect(sl, 20.369, 4.256, 4.56, 2.069, WHITE, { line: { color: DARK, width: 1 } });
    rect(sl, 20.369, 6.664, 4.56, 2.069, WHITE, { line: { color: DARK, width: 1 } });
  }
  /* three teaser columns */
  const COLX = [11.51, 16.126, 20.728];
  COLX.forEach((cx, ci) => {
    if (filled) {
      teaserIcon(sl, ci, cx + 1.875, 9.431);
      txt(sl, [cx + 0.436, 10.477, 3.444, 3.132], [P('t22', 'Headline here'), P('t16', LOREM.webGrid)], { align: 'center' });
    } else {
      rect(sl, cx + 1.875, 9.431, 0.463, 0.463, WHITE, { line: { color: DARK, width: 1 } });
      for (let r = 0; r < 7; r++) rect(sl, cx, 10.389 + r * 0.479, 4.209, 0.319, r === 0 ? CORAL : PALE);
    }
  });
}
/* the three little line-art icons under the hero (piggy bank, page, bank) */
function teaserIcon(sl, kind, x, y) {
  const S = 0.463, ln = { color: DARK, width: 1.25 };
  if (kind === 0) {
    sl.addShape('ellipse', { x, y: y + 0.18 * S, w: 0.8 * S, h: 0.6 * S, fill: { type: 'none' }, line: ln });
    sl.addShape('ellipse', { x: x + 0.55 * S, y: y + 0.3 * S, w: 0.1 * S, h: 0.1 * S, fill: { color: DARK }, line: { width: 0, color: DARK } });
    rect(sl, x + 0.25 * S, y + 0.02 * S, 0.3 * S, 0.16 * S, null, { line: ln });
    rect(sl, x + 0.15 * S, y + 0.75 * S, 0.1 * S, 0.2 * S, null, { line: ln });
    rect(sl, x + 0.5 * S, y + 0.75 * S, 0.1 * S, 0.2 * S, null, { line: ln });
  } else if (kind === 1) {
    rect(sl, x + 0.2 * S, y, 0.6 * S, S, null, { line: ln });
    [0.25, 0.45, 0.65].forEach((f) => hline(sl, x + 0.32 * S, y + f * S, 0.36 * S, DARK, 1));
  } else {
    hline(sl, x, y + 0.95 * S, S, DARK, 1.5);
    [0.1, 0.35, 0.6, 0.85].forEach((f) => vline(sl, x + f * S, y + 0.35 * S, 0.55 * S, DARK, 1.25));
    sl.addShape('triangle', { x, y, w: S, h: 0.32 * S, fill: { type: 'none' }, line: ln });
  }
}
function webGridText(sl, x) {
  eyebrow(sl, x + 0.007, 'Grid system.');
  bigTitle(sl, x, ['Corporate Web Grid', 'System.']);
  sideNote(sl, x - 0.002, 'Web grid system.', [
    P('t16', 'Screen grid', CORAL),
    P('t16', 'This shows approved layouts with a responsive grid for a 16:9 Screen of our website. This will be applicable for websites, landing pages etc.'),
    P('t16', '', CORAL), P('t16', 'Responsive screen grid', CORAL),
    P('t16', 'This is an example of a grid for desktop pc\u00b4s.'),
  ]);
}
function slide22(sl) { vline(sl, RULE_B, 0, SLIDE_H, DARK); webGridText(sl, 1.111); webPage(sl, false); }
function slide23(sl) { vline(sl, RULE_B, 0, SLIDE_H, DARK); webGridText(sl, 1.111); webPage(sl, true); }

/* ==================================================================== 24 */
/* Two phone mock-ups: an empty grid and a filled example */
function phone(sl, x, y, filled) {
  const shell = filled ? { fill: { type: 'none' } } : { fill: { color: WHITE } };
  sl.addShape('roundRect', Object.assign({ x, y, w: 6.296, h: 12.75, rectRadius: 0.75, line: { color: DARK, width: 1 } }, shell));
  sl.addShape('roundRect', Object.assign({ x: x + 0.389, y: y + 0.389, w: 5.522, h: 11.975, rectRadius: 0.5, line: { color: DARK, width: 1 } }, shell));
  /* notch + status bar */
  sl.addShape('roundRect', { x: x + 2.749, y: y + 0.442, w: 0.784, h: 0.104, rectRadius: 0.05, fill: { color: PALE }, line: { width: 0, color: PALE } });
  sl.addShape('ellipse', { x: x + 3.727, y: y + 0.402, w: 0.194, h: 0.194, fill: { color: PALE }, line: { width: 0, color: PALE } });
  txt(sl, [x + 0.75, y + 0.58, 1.2, 0.35], [P('t13', '9:41', DARK)]);
  txt(sl, [x + 4.35, y + 0.58, 1.4, 0.35], [P('t13', '\u2582\u2584\u2586 \u25be \u25ac', DARK)]);
  /* top nav bar with hamburger */
  rect(sl, x + 0.389, y + 1.019, 5.512, 1.028, DARK);
  [1.365, 1.534, 1.703].forEach((dy) => hline(sl, x + 5.15, y + dy, 0.416, WHITE, 1.26));
  if (filled) wordmark(sl, x + 0.388, y + 1.37, 0.265, { color: WHITE, accent: CORAL });
  else rect(sl, x + 0.792, y + 1.365, 1.394, 0.336, WHITE);
  /* hero tile */
  rect(sl, x + 1.345, y + 2.775, 0.854, 0.861, DARK);
  if (filled) logoMark(sl, x + 1.604, y + 2.992, 0.347, { color: WHITE, accent: CORAL });
  /* body area with three cards */
  if (filled) imageBox(sl, x + 0.362, y + 4.371, 5.551, 8.002);
  else rect(sl, x + 0.389, y + 4.371, 5.512, 8.002, DARK);
  const LABELS = ['About us', 'Clients', 'Contact'];
  [4.974, 7.468, 9.975].forEach((dy, i) => {
    rect(sl, x + 1.086, y + dy, 4.19, 1.909, WHITE);
    if (filled) txt(sl, [x + 1.497, y + dy + 0.44, 3.369, 1.028], [P('t28', LABELS[i], CORAL)], { valign: 'middle', align: 'center' });
    else rect(sl, x + 1.81, y + dy + 0.598, 2.734, 0.515, PALE);
  });
}
function slide24(sl) {
  rules(sl, [RULE_A, RULE_B, RULE_C]);
  phone(sl, 2.308, 1.139, false);
  phone(sl, 10.917, 1.139, true);
  eyebrow(sl, 19.799, 'Grid system.');
  bigTitle(sl, 19.792, ['Corporate Mobile Grid', 'System.']);
  sideNote(sl, 19.789, 'Mobile grid system.', [
    P('t16', 'Screen grid', CORAL),
    P('t16', 'This shows approved layouts with a responsive grid for a Mobile Screen of our website. This will be applicable for websites, landing pages etc.'),
    GAP(), P('t16', 'Responsive screen grid', CORAL),
    P('t16', 'This is an example of a grid for mobile phones.'),
  ]);
}

/* ================================================================= 25-28 */
function slide25(sl) {
  vline(sl, RULE_B, 0, SLIDE_H, DARK);
  eyebrow(sl, 1.118, 'Image systems.');
  bigTitle(sl, 1.111, ['Brand ', 'Image', 'System.']);
  sideNote(sl, 1.109, 'Quazar image system.', [
    P('t16', 'Corporate Images are responsible to transfer the values of Quazar to our customers or our potential customers. It is a composite psychological impression that continually changes with the firm\u2019s circumstances, media coverage, performance, pronouncements, etc.'),
    GAP(),
    P('t16', 'Quazar Ltd. use various corporate advertising techniques to enhance their public image in order to improve their desirability as a supplier, employer, customer, borrower, partner, etc.'),
  ]);
  eyebrow(sl, 11.192, 'Colored Images.');
  txt(sl, [11.185, 3.7, 4.147, 3.513], IMAGE_RULES);
  hline(sl, 11.194, 7.926, 5.708, DARK);
  imageBox(sl, 11.194, 9.114, 5.708, 4.762);
  vline(sl, 18.348, 9.114, 5.886, DARK);
  imageBox(sl, 19.785, 0.006, 6.876, 14.994);
}
function slide26(sl) {
  imageBox(sl, -0.006, 0.006, 16.916, 15.003);
  imageBox(sl, 19.787, -0.002, 5.781, 6.465);
  imageBox(sl, 19.787, 9.114, 5.781, 4.762);
  vline(sl, 18.348, 0.005, 6.459, DARK);
  vline(sl, 18.348, 9.12, 5.889, DARK);
  hline(sl, 19.787, 7.789, 5.781, DARK);
}
function slide27(sl) {
  vline(sl, RULE_C, 0, SLIDE_H, DARK);
  rect(sl, 0.007, 0.006, 9.735, 15.003, DARK);
  /* the mark blown up so it bleeds off the left of the page */
  sl.addShape('custGeom', {
    x: -1.147, y: 1.301, w: 15.227, h: 15.227, fill: { color: WHITE }, line: { width: 0, color: WHITE },
    points: scalePath(RING, 0, 0, 15.227, 15.227),
  });
  sl.addShape('custGeom', {
    x: 5.868, y: 6.938, w: 8.21, h: 5.9, fill: { color: CORAL }, line: { width: 0, color: CORAL },
    points: scalePath(SWOOSH, 0, 0, 8.21, 5.9),
  });
  eyebrow(sl, 19.785, 'Image systems.');
  bigTitle(sl, 19.778, ['Brand ', 'Image', 'System.']);
  logoMark(sl, 19.788, 11.621, 0.986);
}
function slide28(sl) {
  eyebrow(sl, 1.118, 'Black & White Images.');
  logoMark(sl, 1.122, 3.998, 0.986);
  txt(sl, [1.109, 9.196, 5.248, 4.762], [P('t34', 'Black and White.'), GAP()].concat(IMAGE_RULES), { valign: 'bottom' });
  imageBox(sl, 11.189, -0.002, 14.378, 6.465);
  hline(sl, 11.193, 7.789, 5.708, DARK);
  hline(sl, 19.787, 7.789, 5.781, DARK);
  imageBox(sl, 11.201, 9.114, 5.699, 4.762);
  imageBox(sl, 19.787, 9.114, 5.781, 4.762);
  vline(sl, 18.348, 9.12, 5.88, DARK);
}

/* ==================================================================== 29 */
function slide29(sl) {
  rules(sl, [RULE_A, RULE_B]);
  eyebrow(sl, 2.571, 'Mood board.');
  bigTitle(sl, 2.564, ['Image ', 'Mood ', 'Board']);
  sideNote(sl, 2.561, 'Quazar mood boards.', [
    P('t16', 'Mood boards and style elements are highly and extremely useful for establishing the aesthetic feel of a design structure. It usually fits into the process somewhere after wireframes and before design process and mockups. Things and designs that can be explored in the mood board include photography style, color palettes, typography, patterns and the overall look and feel. '),
  ]);
  vline(sl, 18.983, 4.508, 10.492, DARK);
  [[9.777, 0, 3.776, 4.214], [13.922, 0.002, 12.741, 4.212], [9.753, 4.541, 8.815, 8.424],
    [19.398, 4.541, 7.265, 5.462], [19.398, 10.33, 7.265, 4.671], [9.777, 13.291, 8.833, 1.702],
  ].forEach(([x, y, w, h]) => imageBox(sl, x, y, w, h));
}

/* ==================================================================== 30 */
/* Three pattern swatches: plus / random dashes / dots */
const PLUS = [
  { x: 0.5765, y: 0 }, { x: 0.4234, y: 0 }, { x: 0.4234, y: 0.4235 }, { x: 0, y: 0.4235 },
  { x: 0, y: 0.5765 }, { x: 0.4234, y: 0.5765 }, { x: 0.4234, y: 1 }, { x: 0.5765, y: 1 },
  { x: 0.5765, y: 0.5765 }, { x: 1, y: 0.5765 }, { x: 1, y: 0.4235 }, { x: 0.5765, y: 0.4235 }, { close: true },
];
/* deterministic pseudo-random so the "random" swatch is stable between runs */
function rnd(seed) { let s = seed; return () => { s = (s * 1103515245 + 12345) % 2147483648; return s / 2147483648; }; }
function patternCard(sl, x, kind, color, title, sub) {
  const y = 5.746, w = 5.705, h = 8.135;
  rect(sl, x, y, w, h, color);
  if (kind === 'plus') {
    for (let r = 0; r < 11; r++) for (let c = 0; c < 9; c++) {
      const px = x + 0.505 + c * 0.559, py = y + 0.659 + r * 0.695;
      sl.addShape('custGeom', { x: px, y: py, w: 0.224, h: 0.224, fill: { color: CORAL }, line: { width: 0, color: CORAL }, points: scalePath(PLUS, 0, 0, 0.224, 0.224) });
    }
  } else if (kind === 'dash') {
    const r = rnd(20210);
    for (let i = 0; i < 250; i++) {
      const px = x + 0.44 + r() * 4.6, py = y + 0.66 + r() * 6.9;
      const a = r() * 180, L = 0.2;
      sl.addShape('line', { x: px, y: py, w: L * Math.cos(a * Math.PI / 180), h: L * Math.abs(Math.sin(a * Math.PI / 180)), flipV: a > 90, line: { color: WHITE, width: 2.5 } });
    }
  } else {
    for (let r2 = 0; r2 < 12; r2++) for (let c = 0; c < 8; c++) {
      sl.addShape('ellipse', { x: x + 0.65 + c * 0.612, y: y + 0.659 + r2 * 0.614, w: 0.126, h: 0.126, fill: { color: DARK }, line: { width: 0, color: DARK } });
    }
  }
  rect(sl, x - 0.103, 11.478, 5.248, 1.861, WHITE);
  txt(sl, [x + 0.399, 11.848, 4.711, 1.117], [P('t34', title), P('t22', sub, CORAL)]);
}
function slide30(sl) {
  rules(sl, [RULE_A, RULE_B, RULE_C]);
  eyebrow(sl, 2.571, 'Design Textures and Patterns.', 8.6);
  txt(sl, [2.567, 3.695, 9.44, 1.117], [P('t28', 'A rough collage of colors, textures and pictures is all it takes to evoke a specific style or feeling. ')]);
  patternCard(sl, 2.599, 'plus', DARK, 'Design Pattern A', 'Plus Pattern.');
  patternCard(sl, 11.192, 'dash', CORAL, 'Design Pattern B', 'Random Pattern.');
  patternCard(sl, 19.803, 'dot', LIGHT, 'Design Pattern C', 'Circle Pattern.');
}

/* ==================================================================== 31 */
/* browser-window icon shown at three tints (dark / coral / 50% light) */
function browserIcon(sl, x, y, w, h, color) {
  const STROKE = 15.8;              // heavy outline, matches the reference weight
  const i = 0.11;                   // half the stroke: keeps the outline inside the box
  const ln = { color, width: STROKE };
  sl.addShape('roundRect', { x: x + i, y: y + i, w: w - 2 * i, h: h - 2 * i, rectRadius: 0.2, fill: { type: 'none' }, line: ln });
  hline(sl, x + i, y + 0.33 * h, w - 2 * i, color, STROKE);
  [0.105, 0.222, 0.34].forEach((f) => sl.addShape('ellipse', {
    x: x + f * w, y: y + 0.175 * h, w: 0.085 * w, h: 0.1 * h, fill: { color }, line: { width: 0, color } }));
  /* the </> glyph */
  const cy = y + 0.637 * h, ah = 0.233 * h, aw = 0.13 * w, chev = { color, width: STROKE };
  sl.addShape('custGeom', { x: x + 0.265 * w, y: cy - ah / 2, w: aw, h: ah, fill: { type: 'none' }, line: chev,
    points: [{ x: aw, y: 0 }, { x: 0, y: ah / 2 }, { x: aw, y: ah }] });
  sl.addShape('custGeom', { x: x + 0.585 * w, y: cy - ah / 2, w: aw, h: ah, fill: { type: 'none' }, line: chev,
    points: [{ x: 0, y: 0 }, { x: aw, y: ah / 2 }, { x: 0, y: ah }] });
  sl.addShape('custGeom', { x: x + 0.435 * w, y: cy - ah / 2, w: 0.115 * w, h: ah, fill: { type: 'none' }, line: chev,
    points: [{ x: 0.115 * w, y: 0 }, { x: 0, y: ah }] });
}
function slide31(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 2.571, 'Iconography system.');
  bigTitle(sl, 2.564, ['Icono-', 'graphy.']);
  txt(sl, [2.561, 7.757, 5.248, 6.201], [
    P('t34', 'Quazar '), P('t34', 'Iconography.'), GAP(),
    P('t16', 'An  icon is a pictogram displayed on a screen  or print layout in order to help the user navigate through the content in a easier way. '),
    GAP(),
    P('t16', 'The icon itself is a small picture or symbol serving as a quick, \u201cintuitive\u201d representation of a software tool, function, feature or a data file. '),
  ], { valign: 'bottom' });
  browserIcon(sl, 11.197, 3.859, 3.556, 2.999, DARK);
  browserIcon(sl, 11.197, 8.674, 3.556, 2.999, CORAL);
  browserIcon(sl, 19.753, 3.859, 3.556, 2.999, LIGHT);
  txt(sl, [11.206, 7.13, 4.011, 0.812], [P('t16', 'Dark Blue Icon', CORAL), P('t16', 'No Background')]);
  txt(sl, [11.206, 11.945, 4.011, 0.812], [P('t16', 'Light Blue Icon', CORAL), P('t16', 'No Background')]);
  txt(sl, [19.78, 7.13, 4.011, 0.812], [P('t16', 'Light Blue Icon 50%', CORAL), P('t16', 'No Background')]);
  txt(sl, [19.784, 8.72, 5.248, 5.238], [
    P('t22', 'Use of Quazar\u00b4s '), P('t22', 'Icon Set.'), GAP(),
    P('t16', 'You can use the approved icon set for your print and web design development.'), GAP(),
    P('t16', 'If icons are missing or you need additional please contact your designer to create new icons in the corporate design style.'),
  ]);
}

/* ==================================================================== 32 */
/* Icon-set sheet: 5 rows x 9 columns of simple line pictograms */
const ICON_X = [8.181, 10.000, 11.847, 14.806, 16.597, 18.431, 21.319, 23.153, 24.986];
const ICON_Y = [5.17, 7.11, 9.07, 11.03, 12.97];
function pictogram(sl, x, y, S, kind, color) {
  const ln = { color, width: 1.6 };
  const none = { type: 'none' };
  switch (kind % 6) {
    case 0: /* framed document */
      rect(sl, x + 0.15 * S, y, 0.7 * S, S, null, { line: ln });
      [0.25, 0.45, 0.65].forEach((f) => hline(sl, x + 0.3 * S, y + f * S, 0.4 * S, color, 1.4));
      break;
    case 1: /* card with picture area */
      rect(sl, x + 0.15 * S, y, 0.7 * S, S, null, { line: ln });
      rect(sl, x + 0.3 * S, y + 0.2 * S, 0.4 * S, 0.3 * S, null, { line: { color, width: 1.2 } });
      hline(sl, x + 0.3 * S, y + 0.72 * S, 0.4 * S, color, 1.4);
      break;
    case 2: /* bar chart in a frame */
      rect(sl, x + 0.15 * S, y, 0.7 * S, S, null, { line: ln });
      [[0.32, 0.35], [0.47, 0.55], [0.62, 0.25]].forEach(([f, hh]) => rect(sl, x + f * S, y + (0.75 - hh) * S, 0.08 * S, hh * S, color));
      break;
    case 3: /* pen nib */
      sl.addShape('triangle', { x: x + 0.25 * S, y, w: 0.5 * S, h: 0.72 * S, fill: none, line: ln });
      vline(sl, x + 0.5 * S, y + 0.3 * S, 0.42 * S, color, 1.4);
      sl.addShape('ellipse', { x: x + 0.42 * S, y: y + 0.45 * S, w: 0.16 * S, h: 0.16 * S, fill: none, line: { color, width: 1.2 } });
      break;
    case 4: /* map pin */
      sl.addShape('ellipse', { x: x + 0.18 * S, y: y, w: 0.64 * S, h: 0.64 * S, fill: none, line: ln });
      sl.addShape('custGeom', { x: x + 0.3 * S, y: y + 0.45 * S, w: 0.4 * S, h: 0.55 * S, fill: none, line: ln,
        points: [{ x: 0, y: 0 }, { x: 0.2 * S, y: 0.55 * S }, { x: 0.4 * S, y: 0 }] });
      break;
    default: /* bank / classical facade */
      sl.addShape('triangle', { x: x + 0.05 * S, y, w: 0.9 * S, h: 0.28 * S, fill: none, line: ln });
      [0.18, 0.42, 0.66].forEach((f) => vline(sl, x + f * S + 0.08 * S, y + 0.32 * S, 0.5 * S, color, 1.4));
      hline(sl, x + 0.05 * S, y + 0.9 * S, 0.9 * S, color, 1.8);
      break;
  }
}
function slide32(sl) {
  rules(sl, [13.536, 20.099]);
  rect(sl, 6.972, 3.278, 6.572, 11.725, DARK);
  eyebrow(sl, 1.123, 'Icon Set Example.');
  txt(sl, [1.116, 3.7, 4.147, 3.513], [P('t48', 'Icono-'), P('t48', 'graphy.')]);
  txt(sl, [1.113, 7.757, 4.054, 6.201], [
    P('t34', 'Quazar '), P('t34', 'Icon Set.'), GAP(),
    P('t16', 'Here you can find an approved icon set for using in your layouts and designs.'), GAP(),
    P('t16', 'If you need additional icons or similar graphics please contact the design studio. They will create and send the new icons to you in a short time..'),
  ], { valign: 'bottom' });
  ICON_Y.forEach((y, r) => ICON_X.forEach((x, c) => {
    const color = c < 3 ? WHITE : (c < 6 ? DARK : CORAL);
    pictogram(sl, x, y, 0.504, r * 3 + c, color);
  }));
}

/* ==================================================================== 33 */
/* Infographic examples: donut, comparison circles, histogram, line chart */
function slide33(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 2.571, 'Infographic system.');
  bigTitle(sl, 2.564, ['Info-', 'graphics', 'System.']);
  txt(sl, [2.561, 7.757, 5.248, 6.201], [
    P('t34', 'Quazar '), P('t34', 'Infographics.'), GAP(),
    P('t16', 'Infographics are graphic visual representations of information, data or knowledge intended to present information quickly and clearly. They can improve cognition by utilizing graphics to enhance the human visual system\u2019s ability to see patterns and trends.'),
  ], { valign: 'bottom' });
  eyebrow(sl, 11.191, 'Examples one.');

  /* round diagrams -- 75% arc over a full ring */
  txt(sl, [11.192, 5.967, 4.011, 0.441], [P('t16', 'Round Diagram', DARK)]);
  sl.addShape('arc', { x: 11.273, y: 6.806, w: 2.389, h: 2.39, angleRange: [0, 359.9], line: { color: DARK, width: 9 } });
  sl.addShape('arc', { x: 11.359, y: 6.892, w: 2.202, h: 2.203, angleRange: [270, 180], line: { color: CORAL, width: 5 } });
  txt(sl, [11.689, 7.867, 1.557, 0.55], [P('t22', '75%')], { align: 'center' });
  sl.addShape('arc', { x: 14.687, y: 7.204, w: 1.58, h: 1.58, angleRange: [0, 359.9], line: { color: DARK, width: 15 } });
  sl.addShape('arc', { x: 14.687, y: 7.204, w: 1.586, h: 1.587, angleRange: [270, 180], line: { color: CORAL, width: 4 } });
  txt(sl, [14.994, 7.867, 0.995, 0.55], [P('t22', '75%')], { align: 'center' });

  /* comparison circles */
  txt(sl, [11.192, 10.68, 4.011, 0.441], [P('t16', 'Comparison Round', DARK)]);
  sl.addShape('ellipse', { x: 12.346, y: 11.595, w: 2.28, h: 2.288, fill: { color: DARK } });
  sl.addShape('ellipse', { x: 11.22, y: 11.458, w: 1.534, h: 1.538, fill: { color: CORAL } });
  txt(sl, [11.207, 12.079, 1.534, 0.542], [P('t22', '+63%', WHITE)], { align: 'center' });
  txt(sl, [12.554, 12.454, 1.853, 0.717], [P('t34', '+27%', WHITE)], { align: 'center' });

  /* stacked histogram */
  txt(sl, [19.91, 5.967, 4.011, 0.441], [P('t16', 'Histogram Thick', DARK)]);
  const BARS = [[2.172, 1.284, 0.592], [2.172, 0.543, 0], [2.172, 0.977, 0.977], [2.172, 1.245, 0.373], [2.172, 0.543, 0]];
  BARS.forEach(([tot, mid, dark], i) => {
    const bx = 19.944 + i * 1.112, top = 6.75, bot = top + tot;
    rect(sl, bx, top, 0.654, tot, LIGHT);
    if (mid) rect(sl, bx, bot - mid, 0.654, mid, CORAL);
    if (dark) rect(sl, bx, bot - dark, 0.654, dark, DARK);
    txt(sl, [19.871 + i * 1.116, 8.939, 0.767, 0.454], [P('t13', String(2019 + i))], { align: 'center' });
  });

  /* line chart with a light grid and three traces */
  txt(sl, [19.91, 10.68, 4.011, 0.441], [P('t16', 'Data with Icons', DARK)]);
  for (let i = 0; i < 9; i++) vline(sl, 20.835 + i * 0.5403, 11.161, 2.693, GREY, 0.75, 'solid');
  for (let i = 0; i < 5; i++) hline(sl, 20.552, 11.542 + i * 0.51, 4.606, GREY, 0.75, 'solid');
  ['+50%', '+25%', '+0%', '-25%', '-50%'].forEach((t, i) => txt(sl, [19.814, 11.43 + i * 0.511, 0.667, 0.215], [P('t10', t)], { align: 'right' }));
  const TRACES = [
    { color: DARK, width: 1, v: [0.62, 0.55, 0.44, 0.48, 0.40, 0.52, 0.44, 0.38, 0.30, 0.35, 0.42] },
    { color: CORAL, width: 1.5, v: [0.44, 0.46, 0.50, 0.44, 0.52, 0.48, 0.55, 0.58, 0.54, 0.50, 0.52] },
    { color: CORAL, width: 1, v: [0.30, 0.35, 0.42, 0.50, 0.46, 0.56, 0.62, 0.70, 0.60, 0.52, 0.48], dash: 'sysDot' },
  ];
  TRACES.forEach((tr) => {
    const x0 = 20.554, w = 4.578, y0 = 11.542, hh = 2.04;
    for (let i = 0; i < tr.v.length - 1; i++) {
      const ax = x0 + (w / (tr.v.length - 1)) * i, bx = x0 + (w / (tr.v.length - 1)) * (i + 1);
      const ay = y0 + hh * (1 - tr.v[i]), by = y0 + hh * (1 - tr.v[i + 1]);
      sl.addShape('line', { x: ax, y: Math.min(ay, by), w: bx - ax, h: Math.abs(by - ay), flipV: by < ay, line: { color: tr.color, width: tr.width, dashType: tr.dash } });
    }
  });
}

/* ==================================================================== 34 */
function slide34(sl) {
  rules(sl, [RULE_B, RULE_C]);
  eyebrow(sl, 2.571, 'Examples two.');

  /* data table */
  txt(sl, [2.581, 5.967, 5.14, 0.441], [P('t16', 'Table for Data', DARK)]);
  [[6.743, 1], [7.542, 3], [8.252, 1], [8.961, 1], [9.671, 3]].forEach(([y, wgt]) => hline(sl, 2.572, y, 5.15, DARK, wgt));
  const ROWS = [['Expenses', '2019', CORAL, 16], ['1. Real Estate of Company', '$22,000', DARK, 13],
    ['2. Real Estate of Company', '$22,000', DARK, 13], ['Total Assets', '$522,000', DARK, 13]];
  [6.884, 7.677, 8.384, 9.092].forEach((y, i) => {
    const [label, val, col, sz] = ROWS[i];
    txt(sl, [2.78, y, 2.5, sz === 16 ? 0.414 : 0.345], [P(sz === 16 ? 't16' : 't13', label, col)], { valign: 'middle' });
    txt(sl, [6.667, y, 0.972, sz === 16 ? 0.414 : 0.345], [P('t13', val)], { valign: 'middle', align: 'right' });
  });

  /* comparison bars */
  txt(sl, [2.581, 10.68, 4.011, 0.441], [P('t16', 'Comparison Bar Chart', DARK)]);
  const REGIONS = [['Europe', 2.293], ['South America', 1.816], ['Australia', 2.683], ['Asia Pacific', 1.93]];
  REGIONS.forEach(([label, len], i) => {
    const y = 11.362 + i * 0.7445;
    rect(sl, 4.242, y, 3.476, 0.298, LIGHT);
    hline(sl, 4.236, y + 0.156, len, DARK, 5);
    txt(sl, [2.569, y - 0.102, 1.667, 0.394], [P('t13', label)], { valign: 'middle' });
  });

  /* thick histogram: paired vertical rules on a ruled ground */
  txt(sl, [11.191, 5.967, 4.011, 0.441], [P('t16', 'Histogram Thick', DARK)]);
  for (let i = 0; i < 9; i++) hline(sl, 11.191, 6.767 + i * 0.848, 5.144, DARK, 1);
  const HIST = [[10.698, 2.86, 9.561, 3.998], [9.439, 4.12, 8.671, 4.887], [10.357, 3.201, 9.594, 3.964],
    [7.898, 5.66, 7.333, 6.226], [7.333, 6.226, 7.757, 5.802]];
  HIST.forEach(([dy, dh, cy, ch], i) => {
    vline(sl, 11.405 + i * 1.1655, dy, dh, DARK, 5);
    vline(sl, 11.462 + i * 1.1655, cy, ch, CORAL, 5);
  });
  ['Jan', 'Feb', 'Mar', 'Apr', 'May'].forEach((m, i) => txt(sl, [11.226 + i * 1.1655, 13.569, 0.415, 0.312], [P('t10', m)], { valign: 'middle', align: 'center' }));

  eyebrow(sl, 20.264, 'Infographic series.');
  txt(sl, [20.323, 3.7, 4.147, 1.278], [P('t48', 'Usage.')]);
  txt(sl, [20.321, 5.946, 5.248, 3.964], [
    P('t34', 'Important:'), GAP(),
    P('t16', 'Please use the approved infographic set for your print, presentation and web design development.'), GAP(),
    P('t16', 'If any infographics are missing or you need additional designs please contact your designer to create new icons in the corporate design style.'),
  ]);
  txt(sl, [20.321, 11.098, 5.248, 2.86], [P('t34', 'Customized '), P('t34', 'Infographic '), P('t34', 'Design.')], { valign: 'bottom' });
}

/* ================================================================= 35-37 */
function slide35(sl) {
  rules(sl, [RULE_A, RULE_B]);
  logoMark(sl, 2.585, 0.954, 0.986);
  txt(sl, [2.564, 3.7, 4.147, 3.513], [
    P('t48', 'Thank you for audience.'), P('t48', '', CORAL), P('t48', 'Quazar.', CORAL)]);
  txt(sl, [2.561, 12.796, 5.248, 1.162], [P('t28', 'www.quazar.com', CORAL)], { valign: 'bottom' });
  imageBox(sl, 11.189, 1.13, 15.467, 13.872);
  logoMark(sl, 16.684, 0.392, 12.592, { color: PALE, accent: PALE });
}
function slide36(sl) {
  rules(sl, [RULE_B, RULE_C]);
  imageBox(sl, -5.718, 1.13, 15.467, 13.872);
  logoMark(sl, -9.964, 0.392, 12.592, { color: PALE, accent: PALE });
  logoMark(sl, 15.913, 7.389, 0.986);
  txt(sl, [19.787, 0.885, 3.267, 2.41], [P('t34', 'Brand Guidelines'), GAP(), P('t34', '2021', CORAL)]);
  txt(sl, [19.787, 4.782, 3.267, 0.963], [P('t34', 'Contact')]);
  txt(sl, [19.787, 7.233, 3.373, 6.437], CONTACT);
}
function slide37(sl) {
  hline(sl, 0, 0.833, 26.664, CORAL);
  hline(sl, 0, 2.333, 26.664, CORAL);
  rect(sl, 26.039, 0, 0.634, 15, DARK);
  logoMark(sl, 1.955, 1.191, 0.634);
  txt(sl, [1.943, 4.406, 10.772, 1.813], [P('t48', 'Thank you '), P('t48', 'for your audience.')]);
  docIcon(sl, 1.948, 10.962, 0.293, 0.407, CORAL);
  hline(sl, 1.953, 11.785, 5.307, CORAL);
  txt(sl, [1.944, 12.052, 5.315, 1.056], [P('t16', '01. Headline here', CORAL), P('t16', LOREM.headline, DARK)]);
  txt(sl, [18.759, 12.696, 5.315, 0.522], [P('t28', 'www.quazar.com', CORAL)], { align: 'right' });
}

/* ================================================================== main */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36, slide37];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'QUAZAR', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'QUAZAR';
  pptx.author = 'Quazar Studio';
  pptx.title = 'Brand Manual & Guidelines 2021';
  BUILDERS.forEach((fn) => {
    const sl = pptx.addSlide();
    sl.background = { color: WHITE };
    fn(sl);
  });
  return pptx.writeFile({ fileName: path.join(__dirname, OUT_NAME) });
}

const OUT_NAME = '13f44317-61ac-44d9-80eb-acbd2cbab0ab_grok_final.pptx';
build().then((f) => console.log('wrote', f));
