/**
 * Recreation of "Sustainable Business Report" (20 slides, 13.333 x 7.5 in)
 * with PptxGenJS only. Every position/size is in inches, every colour is a hex
 * literal taken from the deck's theme palette below.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const C = {
  ink: '171717',        // accent4  - page background
  white: 'FFFFFF',      // bg1
  silver: 'E7E6E6',     // bg2
  indigo: '5149D0',     // accent1
  orchid: 'BF85FE',     // accent2
  lilac: 'D7B5FF',      // accent3
  grey: 'BABABA',       // accent5
  mist: 'F5F6FB',       // accent6
  midGrey: '808080',
  panel: '232323',      // white @5% over the ink background
  hairline: '3D3D3D',
};
const HEAD = 'Noto Sans SemiBold';   // major latin font
const BODY = 'Noto Sans Light';      // minor latin font

/* filler copy reused all over the deck */
const LOREM = {
  full: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod',
  tempor: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  sed: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed',
  sedDo: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do',
  seDo: 'Lorem ipsum dolor sit amet, se do',
  consectetur: 'Lorem ipsum dolor sit amet, consectetur',
  sitDo: 'Lorem ipsum dolor sit do consectetur adipiscing',
  sit: 'Lorem ipsum dolor sit',
  amet: 'Lorem ipsum dolor sit amet',
  magna: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
         'incididunt ut labore et dolore magna',
};

/* ------------------------------------------------------- generic helpers */

/** Body/heading text box. PowerPoint text boxes are top anchored. */
function text(slide, str, o) {
  slide.addText(str, Object.assign({ valign: 'top', fontFace: BODY, color: C.white }, o));
}

/** 14pt translucent caption - the most common paragraph in the deck. */
function caption(slide, str, x, y, w, o) {
  text(slide, str, Object.assign({
    x: x, y: y, w: w, h: 0.69, fontSize: 14, transparency: 30, lineSpacingMultiple: 1.3,
  }, o));
}

function lerp(a, b, t) { return a + (b - a) * t; }

function mixHex(c1, c2, t) {
  const p = i => parseInt(c1.substr(i, 2), 16), q = i => parseInt(c2.substr(i, 2), 16);
  const h = v => Math.round(v).toString(16).padStart(2, '0').toUpperCase();
  return h(lerp(p(0), q(0), t)) + h(lerp(p(2), q(2), t)) + h(lerp(p(4), q(4), t));
}

/** Sample a stop list [[pos, hex, transparency?], ...] at position t. */
function sampleStops(stops, t) {
  if (t <= stops[0][0]) return { color: stops[0][1], transparency: stops[0][2] || 0 };
  const last = stops[stops.length - 1];
  if (t >= last[0]) return { color: last[1], transparency: last[2] || 0 };
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0]) {
      const a = stops[i - 1], b = stops[i], f = (t - a[0]) / (b[0] - a[0]);
      return { color: mixHex(a[1], b[1], f), transparency: lerp(a[2] || 0, b[2] || 0, f) };
    }
  }
}

/* ---- gradients -----------------------------------------------------------
 * PptxGenJS cannot emit <a:gradFill>, so every gradient is painted as a stack
 * of solid slices cut perpendicular to the gradient axis and clipped to the
 * shape outline, which keeps the silhouette exact for rectangles, circles and
 * stadium shapes alike.
 * ------------------------------------------------------------------------ */

/** Sutherland-Hodgman clip of a convex polygon by the half plane n.p <= c. */
function clipHalfPlane(poly, nx, ny, c) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i], b = poly[(i + 1) % poly.length];
    const da = nx * a[0] + ny * a[1] - c, db = nx * b[0] + ny * b[1] - c;
    if (da <= 0) out.push(a);
    if ((da < 0 && db > 0) || (da > 0 && db < 0)) {
      const f = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f]);
    }
  }
  return out;
}

/** Fill a convex polygon (inch coords) with a linear gradient. */
function gradientPoly(slide, poly, angleDeg, stops) {
  const th = angleDeg * Math.PI / 180, dx = Math.cos(th), dy = Math.sin(th);
  const proj = poly.map(p => p[0] * dx + p[1] * dy);
  const lo = Math.min.apply(null, proj), hi = Math.max.apply(null, proj);
  const bands = Math.min(48, Math.max(8, Math.round((hi - lo) * 9)));
  const step = (hi - lo) / bands;
  for (let i = 0; i < bands; i++) {
    const a = lo + i * step;
    const piece = clipHalfPlane(clipHalfPlane(poly, dx, dy, a + step * 1.04), -dx, -dy, -a);
    if (piece.length < 3) continue;
    const xs = piece.map(p => p[0]), ys = piece.map(p => p[1]);
    const bx = Math.min.apply(null, xs), by = Math.min.apply(null, ys);
    const bw = Math.max(Math.max.apply(null, xs) - bx, 0.01);
    const bh = Math.max(Math.max.apply(null, ys) - by, 0.01);
    const pts = piece.map((p, k) => ({ x: p[0] - bx, y: p[1] - by, moveTo: k === 0 }));
    pts.push({ close: true });
    slide.addShape('custGeom', { x: bx, y: by, w: bw, h: bh, points: pts,
      fill: sampleStops(stops, (i + 0.5) / bands) });
  }
}

function rectPoly(x, y, w, h) { return [[x, y], [x + w, y], [x + w, y + h], [x, y + h]]; }

function ellipsePoly(x, y, w, h, n) {
  const pts = [];
  for (let i = 0; i < (n || 48); i++) {
    const a = 2 * Math.PI * i / (n || 48);
    pts.push([x + w / 2 * (1 + Math.cos(a)), y + h / 2 * (1 + Math.sin(a))]);
  }
  return pts;
}

/** Stadium (fully rounded rectangle) outline. */
function pillPoly(x, y, w, h, n) {
  const r = h / 2, steps = n || 20, pts = [];
  for (let i = 0; i <= steps; i++) {        // left cap: top -> left -> bottom
    const a = 1.5 * Math.PI - Math.PI * i / steps;
    pts.push([x + r + r * Math.cos(a), y + r + r * Math.sin(a)]);
  }
  for (let i = 0; i <= steps; i++) {        // right cap: bottom -> right -> top
    const a = 0.5 * Math.PI - Math.PI * i / steps;
    pts.push([x + w - r + r * Math.cos(a), y + r + r * Math.sin(a)]);
  }
  return pts;
}

function gradient(slide, o) {
  gradientPoly(slide, rectPoly(o.x, o.y, o.w, o.h), o.angle, o.stops);
}
function gradientEllipse(slide, o) {
  gradientPoly(slide, ellipsePoly(o.x, o.y, o.w, o.h), o.angle, o.stops);
}
function gradientPill(slide, o) {
  gradientPoly(slide, pillPoly(o.x, o.y, o.w, o.h), o.angle, o.stops);
}

/**
 * Freeform outline. `pathData` holds [[cmd, ...coords]] with coordinates
 * normalised to 0..1 of the shape box; custGeom points are relative to the
 * shape origin, so they are simply scaled by w/h.
 */
function freeform(slide, pathData, o) {
  const pts = [];
  for (const seg of pathData) {
    const k = seg[0];
    const X = i => seg[i] * o.w, Y = i => seg[i] * o.h;
    if (k === 'M') pts.push({ x: X(1), y: Y(2), moveTo: true });
    else if (k === 'L') pts.push({ x: X(1), y: Y(2) });
    else if (k === 'C') pts.push({ curve: { type: 'cubic', x1: X(1), y1: Y(2), x2: X(3), y2: Y(4) }, x: X(5), y: Y(6) });
    else if (k === 'Z') pts.push({ close: true });
  }
  slide.addShape('custGeom', Object.assign({ x: o.x, y: o.y, w: o.w, h: o.h, points: pts },
    o.opts || {}));
}

/** Flat polygon list (closed) used for the world map silhouette. */
function polygons(slide, box, polys, opts) {
  for (const flat of polys) {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) {
      pts.push({ x: flat[i] * box.w, y: flat[i + 1] * box.h, moveTo: i === 0 });
    }
    pts.push({ close: true });
    slide.addShape('custGeom', Object.assign({ x: box.x, y: box.y, w: box.w, h: box.h, points: pts }, opts));
  }
}

/* -------------------------------------------------- shared deck furniture */

/** Slide master decoration: rule, site address and page number. */
function footer(slide, page) {
  slide.addShape('line', { x: 2.667, y: 7.008, w: 9.584, h: 0,
    line: { color: C.silver, width: 1, transparency: 40 } });
  text(slide, 'www.companysite.co', { x: 0.623, y: 6.841, w: 1.937, h: 0.32, fontSize: 13, transparency: 20 });
  text(slide, String(page), { x: 12.251, y: 6.848, w: 0.565, h: 0.32, fontSize: 13, align: 'right', transparency: 30 });
}

/** Small "Company" lock-up: two-bar mark plus wordmark. */
function logo(slide, x, y, markColor) {
  freeform(slide, LOGO, { x: x, y: y, w: 0.346, h: 0.337, opts: { fill: { color: markColor || C.indigo } } });
  text(slide, 'Company', { x: x + 0.438, y: y + 0.08, w: 1.107, h: 0.337, fontSize: 14, fontFace: HEAD });
}

/**
 * Slides 13-15 put their heading on an outline level 2 paragraph, which the
 * master indents by 0.5in. PptxGenJS zeroes marL, so the indent is folded into
 * the geometry instead.
 */
function indentedTitle(slide, str, o) {
  text(slide, str, Object.assign({}, o, { x: o.x + 0.5, w: o.w - 0.5,
    fontFace: HEAD, lineSpacingMultiple: 0.9 }));
}

/** Numbered row used on slides 5 and 15: "01  Heading  lorem" */
function numberedRow(slide, y, num, heading, x0) {
  const x = x0 === undefined ? 5.224 : x0;
  text(slide, num, { x: x, y: y, w: 1.207, h: 0.464, fontSize: 24, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  text(slide, heading, { x: x + 1.109, y: y + 0.001, w: 2.009, h: 0.64, fontSize: 16, fontFace: HEAD });
  caption(slide, LOREM.sed, x + 4.449, y - 0.049, 3.081);
}

/* --------------------------------------------------------- freeform paths */
// long S-curve ribbon with a loop (title/closing slides)
const WAVE = [
  ['M',.808,1],['C',.7994,1,.7911,.9949,.7842,.9847],['C',.7687,.962,.7623,.9183,.7686,.8784],
  ['C',.7714,.8608,.7764,.844,.7835,.8281],['C',.7626,.8267,.7414,.842,.7228,.8567],
  ['C',.7203,.8587,.7177,.8608,.7152,.8628],['C',.6927,.8808,.6695,.8994,.6438,.9019],
  ['C',.6148,.9047,.5792,.8811,.5653,.8215],['C',.554,.7729,.56,.7115,.5817,.653],
  ['C',.5878,.6366,.5945,.6214,.601,.6067],['C',.6114,.583,.6213,.5607,.628,.5352],
  ['C',.6504,.4507,.633,.3109,.5715,.2375],['C',.5261,.1833,.467,.1663,.3856,.184],
  ['C',.3657,.1883,.3453,.1942,.3256,.2],['C',.2795,.2134,.2319,.2273,.1841,.2227],
  ['C',.1025,.2147,.0343,.1463,.0019,.0397],['C',-.0019,.0272,.0002,.011,.0067,.0036],
  ['C',.0132,-.0037,.0216,.0004,.0254,.013],['C',.0532,.1043,.113,.163,.1855,.17],
  ['C',.2306,.1744,.2768,.1609,.3215,.1479],['C',.3415,.1421,.3621,.136,.3826,.1316],
  ['C',.4691,.1128,.5356,.1328,.5858,.1927],['C',.6458,.2643,.6869,.4286,.6523,.5592],
  ['C',.6441,.5902,.6327,.616,.6217,.6409],['C',.6153,.6554,.6092,.6692,.6038,.6837],
  ['C',.598,.6995,.5796,.7546,.5901,.7998],['C',.5987,.8366,.6233,.8513,.6424,.8493],
  ['C',.6634,.8472,.6835,.8311,.7048,.8141],['C',.7073,.812,.7099,.81,.7125,.808],
  ['C',.7386,.7874,.7667,.7694,.7953,.7772],['C',.8004,.7786,.8055,.7811,.8105,.7845],
  ['C',.8666,.7168,.947,.7292,.9956,.8165],['C',1.0012,.8264,1.0015,.8431,.9964,.8538],
  ['C',.9912,.8644,.9826,.865,.9771,.8551],['C',.9407,.7899,.8828,.7756,.8379,.8156],
  ['C',.85,.8364,.8584,.8646,.8591,.8985],['C',.8598,.9344,.848,.9709,.8299,.9891],['C',.8227,.9964,.8152,1,.808,1],
  ['Z'],['M',.8133,.8455],['C',.8035,.8611,.7973,.8774,.7947,.8939],['C',.7917,.9124,.7943,.9334,.8006,.9427],
  ['C',.8067,.9515,.8143,.9454,.8173,.9424],['C',.826,.9336,.8321,.916,.8318,.9004],
  ['C',.8314,.8779,.8236,.8584,.8133,.8455],['Z'],
];

// looped knot ribbon
const KNOT = [
  ['M',.0345,0],['C',.037,0,.0396,.0003,.0422,.0011],['C',.1488,.0304,.2535,.0784,.3504,.1418],
  ['C',.4118,.0838,.4988,.059,.5836,.0802],['C',.6973,.1087,.7828,.2052,.8345,.2811],
  ['C',.9621,.4686,1.0214,.724,.993,.9642],['C',.9903,.987,.9729,1.0028,.954,.9996],
  ['C',.9351,.9963,.922,.9752,.9247,.9524],['C',.9504,.7349,.8967,.5038,.7812,.334],
  ['C',.7366,.2685,.6636,.1854,.5696,.1619],['C',.516,.1484,.4603,.1595,.4168,.1886],
  ['C',.4327,.2007,.4485,.2133,.4639,.2263],['C',.5529,.3011,.5908,.3904,.5679,.4712],
  ['C',.5574,.5082,.5341,.5384,.5022,.5563],['C',.4647,.5773,.4197,.5792,.3818,.5614],
  ['C',.3345,.5392,.2972,.4895,.2795,.4252],['C',.2594,.3526,.2663,.2728,.2979,.2117],
  ['C',.2991,.2094,.3003,.2071,.3016,.2049],['C',.2147,.1501,.1216,.1084,.0268,.0823],
  ['C',.0083,.0772,-.0034,.0548,.0009,.0324],['C',.0045,.0131,.0188,0,.0345,0],['Z'],['M',.4366,.4902],
  ['C',.4491,.4902,.4619,.487,.4731,.4807],['C',.4835,.4749,.497,.4638,.5026,.4442],
  ['C',.5204,.3814,.4473,.3139,.4244,.2946],['C',.4042,.2776,.3833,.2613,.362,.2458],
  ['C',.3601,.2491,.3582,.2524,.3565,.2558],['C',.336,.2955,.3316,.3503,.3449,.3989],
  ['C',.3563,.4402,.3783,.4703,.4068,.4837],['C',.416,.488,.4263,.4902,.4366,.4902],['Z'],
];

// short ribbon with a single loop
const LOOP = [
  ['M',.5971,1],['C',.577,1,.5578,.9735,.5433,.9246],['C',.5224,.8545,.5155,.7526,.5259,.671],
  ['C',.5358,.5939,.5567,.5287,.5883,.4765],['C',.5642,.3891,.5246,.2932,.46,.2212],
  ['C',.322,.0674,.1598,.1168,.0367,.3501],['C',.0259,.3707,.0113,.3625,.004,.332],
  ['C',-.0033,.3016,-.0004,.2602,.0104,.2398],['C',.1462,-.0177,.3251,-.0722,.4773,.0975],
  ['C',.5443,.1722,.5977,.2844,.6321,.4191],['C',.7449,.3022,.8736,.3075,.9852,.4338],
  ['C',.9973,.4475,1.0031,.4862,.9983,.5203],['C',.9934,.5544,.9797,.5709,.9676,.5572],
  ['C',.8693,.446,.7563,.4386,.6562,.5361],['C',.6612,.5676,.6653,.6,.6684,.633],
  ['C',.6772,.7272,.6724,.834,.6563,.9052],['C',.6448,.9557,.6282,.9875,.6084,.9973],
  ['C',.6046,.9991,.6009,1,.5971,1],['Z'],['M',.6129,.5898],['C',.5907,.626,.5764,.6683,.5703,.7163],
  ['C',.5655,.7536,.5693,.8052,.5794,.8391],['C',.5838,.854,.5911,.8708,.6003,.8662],
  ['C',.607,.863,.6123,.8522,.6166,.8334],['C',.6257,.7933,.6282,.7249,.6228,.6669],
  ['C',.621,.6471,.6179,.6205,.6129,.5898],['Z'],
];

// two-bar company logo mark
const LOGO = [
  ['M',0,.3914],['L',.4421,.3914],['L',.4421,1],['L',0,1],['Z'],['M',.5579,0],['L',1,.3559],['L',1,1],['L',.5579,1],
  ['Z'],
];

// north-east arrow glyph
const ARROW = [
  ['M',.2391,.5625],['L',.5891,.9125],['L',.5,1],['L',0,.5],['L',.5,0],['L',.5891,.0875],['L',.2391,.4375],
  ['L',1,.4375],['L',1,.5625],['L',.2391,.5625],['Z'],
];

// rounded speech-bubble callout
const CALLOUT = [
  ['M',.0262,0],['L',.1355,0],['L',.8645,0],['L',.9738,0],['C',.9883,0,1,.025,1,.0559],['L',1,.845],
  ['C',1,.8758,.9883,.9008,.9738,.9008],['L',.8661,.9008],['L',.8645,.9022],['L',.5648,.9022],['L',.5,1],
  ['L',.4352,.9022],['L',.1355,.9022],['L',.1339,.9008],['L',.0262,.9008],['C',.0117,.9008,0,.8758,0,.845],
  ['L',0,.0559],['C',0,.025,.0117,0,.0262,0],['Z'],
];

// open C-shaped connector arm
const TRACK = [
  ['M',.1874,0],['C',.2187,.0004,.2503,.0215,.2807,.0646],['C',.3657,.1853,.4271,.4584,.4451,.7832],
  ['L',.4481,.8697],['L',1,.8697],['L',1,1],['L',.4153,1],['L',.4153,.9326],['L',.4148,.9327],
  ['C',.4097,.5986,.3523,.3067,.2681,.1873],['C',.1839,.0678,.0887,.143,.025,.3794],['L',0,.2864],
  ['C',.0506,.0988,.1184,-.0008,.1874,0],['Z'],
];

// SE-Asia map silhouette: island outlines as normalized polygons inside each box
const MAP_ISLANDS = [
  { x:11.275, y:4.845, w:1.343, h:.953, polys:[
    [.717,.926,.66,.838,.585,.779,.509,.515,.528,.426,.057,.029,0,0,0,.779,.085,.824,.179,.735,.226,.647,.387,.706,.566,.956,.726,1,.717,.926],
    [.849,.265,.755,.353,.585,.368,.717,.456,.877,.324,.906,.221,.849,.265],
    [.934,.132,.906,.132,.962,.25,.934,.132],
  ] },
  { x:7.097, y:2.774, w:.644, h:.503, polys:[
    [.49,.889,.667,.889,.667,.694,.824,.556,.941,.278,.882,0,.765,.028,.647,.083,.588,.139,.51,.056,.275,.028,.059,.194,.059,.417,.118,.806,.235,.833,.235,.944,.373,.972,.373,1,.49,.889],
  ] },
  { x:6.882, y:1.791, w:.784, h:1.065, polys:[
    [.677,.461,.532,.355,.645,.303,.597,.224,.387,.145,.274,.013,.21,.013,.194,.158,.113,.105,.048,.197,0,.224,.032,.25,.048,.329,.129,.355,.129,.474,.081,.566,.21,.526,.516,.5,.613,.579,.629,.711,.726,.829,.71,.961,.758,.987,.806,.961,1,.921,.968,.763,.855,.671,.677,.461],
  ] },
  { x:6.613, y:2.028, w:.865, h:1.794, polys:[
    [.603,.531,.603,.469,.765,.422,.956,.438,.971,.359,.882,.289,.868,.211,.779,.164,.382,.203,.426,.148,.426,.078,.353,.062,.338,.016,.309,0,.015,.078,0,.188,.088,.242,.103,.289,.088,.367,.221,.57,.118,.672,.074,.773,.088,.828,.279,.945,.353,.953,.412,.992,.515,.992,.559,.969,.412,.922,.324,.852,.279,.789,.206,.758,.221,.656,.279,.516,.382,.5,.618,.578,.603,.531],
  ] },
  { x:6.855, y:3.709, w:2.084, h:.716, polys:[
    [.183,.353,.116,.078,.098,.137,.055,.137,.03,.039,0,.02,.067,.588,.116,.745,.213,.902,.183,.667,.183,.353],
    [.957,.176,.915,.137,.866,0,.787,.333,.787,.431,.756,.471,.726,.353,.677,.569,.604,.647,.573,.863,.494,.804,.512,.902,.555,.961,.622,.941,.671,.882,.72,.902,.762,.804,.811,.51,.89,.412,.915,.431,.945,.333,.994,.275,.957,.176],
  ] },
  { x:8.364, y:3.946, w:.14, h:.101, polys:[[.455,1,.909,.714,.909,0,.545,.286,0,.143,.455,1]] },
  { x:6.065, y:1.021, w:.918, h:2.226, polys:[
    [.806,.912,.681,.748,.694,.711,.681,.648,.597,.604,.611,.516,.875,.459,1,.396,.944,.403,.806,.371,.806,.333,.722,.264,.611,.258,.597,.233,.694,.157,.722,.082,.653,.025,.569,.006,.528,.038,.542,.075,.389,.094,.306,.138,.236,.264,.139,.264,.083,.371,0,.428,.097,.478,.236,.585,.222,.673,.264,.698,.389,.692,.486,.642,.542,.66,.694,.918,.708,.994,.806,.912],
  ] },
  { x:7.097, y:1.678, w:.784, h:1.805, polys:[
    [.742,.124,.597,.07,.565,.031,.452,.008,.339,.039,.032,.054,0,.07,.113,.147,.323,.194,.371,.24,.258,.271,.403,.333,.581,.457,.694,.512,.774,.682,.677,.76,.548,.798,.548,.853,.403,.853,.306,.884,.371,.907,.339,.961,.403,.992,.532,.946,.597,.884,.871,.837,.968,.752,.935,.597,.806,.496,.597,.403,.532,.357,.5,.264,.645,.171,.758,.14,.742,.124],
  ] },
  { x:8.708, y:2.264, w:1.042, h:1.598, polys:[
    [.524,.386,.659,.456,.683,.518,.732,.553,.78,.64,.854,.579,.829,.491,.646,.368,.5,.351,.451,.281,.524,.175,.512,.053,.5,.026,.354,.018,.329,.193,.28,.184,.329,.342,.524,.386],
    [.341,.404,.439,.482,.341,.404],[.5,.623,.561,.596,.585,.737,.695,.623,.5,.518,.5,.623],
    [.146,.658,.012,.781,.207,.658,.232,.579,.146,.658],[.683,.702,.768,.667,.683,.702],
    [.963,.877,.939,.711,.866,.702,.744,.781,.659,.763,.537,.816,.5,.886,.659,.833,.744,.965,.854,1,.854,.904,.963,.877],
  ] },
  { x:6.345, y:3.809, w:4.93, h:1.977, polys:[
    [.075,.426,.085,.461,.075,.426],[.284,.532,.271,.546,.284,.532],[.041,.27,.057,.319,.041,.27],
    [.526,.574,.536,.709,.544,.539,.57,.624,.588,.631,.577,.468,.611,.404,.567,.44,.541,.376,.611,.333,.649,.27,.626,.298,.572,.284,.546,.312,.513,.518,.526,.574],
    [.235,.617,.247,.511,.229,.461,.235,.518,.206,.475,.149,.262,.119,.227,.064,.085,.005,.021,.121,.482,.204,.709,.227,.709,.235,.617],
    [.724,.34,.732,.291,.704,.255,.716,.383,.724,.34],[.67,.56,.698,.582,.67,.56],
    [.621,.858,.534,.887,.57,.901,.621,.858],[.758,.532,.716,.553,.773,.596,.758,.532],
    [.42,.837,.381,.823,.405,.78,.237,.73,.227,.766,.253,.809,.387,.872,.441,.872,.42,.837],
    [.925,.454,.871,.56,.848,.411,.796,.397,.778,.433,.802,.489,.838,.496,.835,.518,.804,.532,.817,.589,.856,.603,.936,.688,.946,.809,.925,.858,.977,.851,1,.901,1,.525,.925,.454],
    [.647,.901,.629,.915,.619,.986,.655,.922,.647,.901],
    [.469,.447,.482,.355,.513,.305,.479,.099,.446,.135,.407,.277,.387,.27,.338,.298,.309,.234,.296,.319,.307,.404,.363,.567,.451,.582,.469,.447],
    [.515,.943,.557,.972,.515,.943],[.485,.851,.464,.858,.454,.894,.485,.908,.515,.879,.485,.851],
  ] },
  { x:9.443, y:5.49, w:.36, h:.154, polys:[
    [.107,.727,0,.818,.107,.909,.143,.727,.107,.727],[.536,.182,.286,.545,.25,.636,.357,.909,.964,.182,.536,.182],
  ] },
];

/* ---------------------------------------------------- recurring gradients */
/* accent1 -> accent2 diagonal (down-right), the deck's signature sweep */
const SWEEP = [[0.26, C.indigo], [1, C.orchid]];
const SWEEP7 = [[0.07, C.indigo], [1, C.orchid]];
const SWEEP6 = [[0.06, C.indigo], [1, C.orchid]];
/* dark -> indigo page wash used on the title / closing slides */
const NIGHT = [[0, C.ink], [1, C.indigo]];

/**
 * The swooping ribbons are rotated freeforms; a freeform cannot carry a banded
 * gradient, so each one gets the flat mid tone of the gradient it replaces.
 */
const RIBBON = {
  team: '36317A',      // slide 4  - dark indigo
  scalable: '906BEA',  // slide 5  - bright orchid
  goal: '8A68E7',      // slide 6
  quoteTop: '3B3688',  // slide 12 - upper right
  quoteLow: '8565E5',  // slide 12 - lower left
  vision: '353176',    // slide 15
  future: '4941A5',    // slide 17
  conclusion: '37327D', // slide 19
  track: '936DEA',     // slide 13 - progress arms
};

/* ============================================================== slide 01 */
function slide01(pres) {
  const s = pres.addSlide();
  gradient(s, { x: 0, y: 0, w: 13.334, h: 7.5, angle: 45, stops: NIGHT });
  freeform(s, WAVE, { x: 0.332, y: -1.775, w: 13.972, h: 7.311,
    opts: { rotate: 341.8, fill: { color: C.white, transparency: 91 } } });
  logo(s, 0.8, 3.007);
  text(s, 'SUSTAINABLE BUSINESS REPORT', { x: 0.703, y: 3.917, w: 9.741, h: 2.592,
    fontSize: 74, fontFace: HEAD });
  text(s, '\u00A9 2026', { x: 11.394, y: 5.82, w: 1.449, h: 0.509, fontSize: 20,
    fontFace: HEAD, lineSpacingMultiple: 1.3 });
}

/* ============================================================== slide 02 */
const AGENDA = [
  { n: '01', label: 'Introduction', y: 0.727 },
  { n: '02', label: 'Bussines Analyst', y: 2.038 },
  { n: '03', label: 'Bussines Model', y: 3.199, active: true },
  { n: '04', label: 'Strategic Roadmap', y: 4.399 },
  { n: '05', label: 'Growth Overview', y: 5.597 },
];
function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  gradient(s, { x: 0.8, y: 3.189, w: 11.733, h: 1.21, angle: 45,
    stops: [[0.12, C.indigo], [1, C.orchid]] });
  for (const y of [1.968, 4.388, 5.597]) {
    s.addShape('line', { x: 1.0, y: y, w: 11.333, h: 0,
      line: { color: C.silver, width: 1, transparency: 70 } });
  }
  for (const row of AGENDA) {
    const dim = row.active ? 0 : 30;
    text(s, row.n, { x: 1.329, y: row.y + 0.038, w: 1.444, h: 1.002, fontSize: 48,
      transparency: dim, lineSpacingMultiple: 1.2 });
    text(s, row.label, { x: 3.727, y: row.y, w: 8.317, h: 1.002, fontSize: 48,
      transparency: dim, lineSpacingMultiple: 1.2 });
  }
  freeform(s, ARROW, { x: 11.388, y: 3.49, w: 0.547, h: 0.547,
    opts: { rotate: 143.9, fill: { color: C.white } } });
}

/* ============================================================== slide 03 */
function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  gradient(s, { x: 8.028, y: 0, w: 5.306, h: 7.5, angle: 45,
    stops: [[0.33, C.indigo], [1, C.orchid]] });
  freeform(s, KNOT, { x: 7.352, y: 1.032, w: 8.031, h: 6.467,
    opts: { rotate: 36.2, fill: { color: C.white, transparency: 70 } } });
  logo(s, 0.8, 1.355);
  text(s, 'Welcome to Our Big Company', { x: 0.607, y: 2.142, w: 5.425, h: 2.827,
    fontSize: 60, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  caption(s, LOREM.full, 0.697, 5.34, 4.199);
  text(s, '2026', { x: 12.141, y: 0.382, w: 0.616, h: 0.32, fontSize: 13, transparency: 20 });
  s.addShape('line', { x: 7.908, y: 7.008, w: 4.343, h: 0,
    line: { color: C.silver, width: 1, transparency: 40 } });
  footer(s, 3);
}

/* ============================================================== slide 04 */
const TEAM = [
  { name: 'Adrian Cole', x: 1.542, y: 3.662, h: 2.681, plateY: 5.989, plateH: 0.477 },
  { name: 'Marcus Lee', x: 5.229, y: 3.186, h: 2.681, plateY: 5.458, plateH: 0.531 },
  { name: 'Daniel Foster', x: 8.917, y: 3.662, h: 2.681, plateY: 5.989, plateH: 0.477 },
];
function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 4);
  freeform(s, LOOP, { x: -1.674, y: 0.714, w: 6.651, h: 2.358,
    opts: { rotate: 311.9, fill: { color: RIBBON.team } } });
  text(s, 'Our Best Team', { x: 3.041, y: 0.662, w: 7.073, h: 1.01, fontSize: 60,
    fontFace: HEAD, align: 'center', lineSpacingMultiple: 0.9 });
  caption(s, LOREM.full, 4.194, 1.786, 4.766, { align: 'center' });
  for (const m of TEAM) {
    gradient(s, { x: m.x, y: m.plateY, w: 2.653, h: m.plateH, angle: 20,
      stops: [[0.26, C.indigo, 0], [0.87, C.indigo, 100]] });
    text(s, m.name, { x: m.x + 0.11, y: m.plateY + 0.07, w: 2.009, h: 0.303,
      fontSize: 12, fontFace: HEAD });
  }
}

/* ============================================================== slide 05 */
function slide05(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 5);
  freeform(s, LOOP, { x: 6.067, y: -0.071, w: 7.957, h: 2.821,
    opts: { rotate: 13.7, fill: { color: RIBBON.scalable } } });
  text(s, 'A Scalable Business Model', { x: 0.661, y: 1.022, w: 7.073, h: 1.919,
    fontSize: 60, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  gradient(s, { x: 2.542, y: 4.062, w: 2.153, h: 2.03, angle: 45, stops: SWEEP });
  text(s, '50%', { x: 2.737, y: 4.424, w: 1.222, h: 0.707, fontSize: 36, fontFace: HEAD });
  caption(s, LOREM.amet, 2.745, 5.17, 1.814, { transparency: 20 });
  numberedRow(s, 4.126, '01', 'Long Term Value Creation');
  s.addShape('line', { x: 5.333, y: 5.043, w: 7.2, h: 0,
    line: { color: C.silver, width: 1, transparency: 70 } });
  numberedRow(s, 5.293, '02', 'Responsible Resource ');
}

/* ============================================================== slide 06 */
const GOALS = [
  { n: '01', title: 'Responsibility', y: 1.303 },
  { n: '02', title: 'Social Impact', y: 3.128 },
  { n: '03', title: 'Strong Governance', y: 4.951 },
];
function slide06(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 6);
  freeform(s, LOOP, { x: -1.711, y: -0.109, w: 7.115, h: 2.523,
    opts: { rotate: 322.9, fill: { color: RIBBON.goal } } });
  text(s, 'Sustainable Goal and Commit', { x: 0.803, y: 2.542, w: 5.054, h: 2.555,
    fontSize: 54, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  caption(s, LOREM.sedDo, 0.831, 5.347, 3.747);
  gradient(s, { x: 7.004, y: 1.011, w: 0.201, h: 5.36, angle: 90,
    stops: [[0.17, C.indigo], [1, C.orchid]] });
  GOALS.forEach((g, i) => {
    s.addShape('line', { x: 7.56, y: [1.046, 2.861, 4.676][i], w: 4.774, h: 0,
      line: { color: C.orchid, width: 1, transparency: 70 } });
    text(s, g.n, { x: 7.645, y: g.y + 0.028, w: 0.458, h: 0.37, fontSize: 16, fontFace: HEAD });
    text(s, g.title, { x: 8.28, y: g.y, w: 2.6, h: 0.404, fontSize: 18, fontFace: HEAD });
    caption(s, LOREM.tempor, 7.62, g.y + 0.492, 4.17);
  });
}

/* ============================================================== slide 07 */
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function slide07(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 7);
  s.addShape('rect', { x: 0.96, y: 2.174, w: 11.414, h: 2.736,
    fill: { color: C.white, transparency: 95 }, line: { color: C.white, width: 0.5, transparency: 80 } });
  text(s, 'Analysis This Month', { x: 1.878, y: 0.659, w: 9.577, h: 0.919, fontSize: 54,
    fontFace: HEAD, align: 'center', lineSpacingMultiple: 0.9 });
  s.addChart('line', [{
    name: 'Series 1', labels: MONTHS,
    values: [1.5, 4.6, 2.4, 5.4, 4.23, 5.13, 3.9, 6.5, 8.5, 7.32, 4.7, 8.56],
  }], {
    x: 1.229, y: 2.504, w: 10.875, h: 1.956,
    chartColors: [C.indigo], lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 5,
    lineDataSymbolLineColor: C.white, lineDataSymbolLineSize: 2,
    showLegend: false, catAxisLabelColor: C.white, catAxisLabelFontFace: HEAD,
    catAxisLabelFontSize: 10, catAxisLineShow: false,
    valAxisLabelColor: C.white, valAxisLabelFontSize: 10, valAxisLineShow: false,
    valAxisLabelFormatCode: '"$"#,##0.00', valAxisMajorUnit: 2,
    valGridLine: { color: '333333', size: 0.5 }, catGridLine: { style: 'none' },
  });
  // "Title Here / 80%" speech-bubble marker over the October data point
  freeform(s, CALLOUT, { x: 9.784, y: 3.653, w: 1.76, h: 0.826,
    opts: { flipV: true, fill: { color: mixHex(C.indigo, C.orchid, 0.55) } } });
  text(s, 'Title Here', { x: 9.846, y: 3.76, w: 1.635, h: 0.345, fontSize: 12, align: 'center' });
  text(s, '80%', { x: 10.067, y: 3.999, w: 1.193, h: 0.462, fontSize: 18, fontFace: HEAD, align: 'center' });
  [['01', 0.876, 1.795, 2.691], ['02', 5.099, 6.017, 2.83], ['03', 9.121, 10.04, 2.83]]
    .forEach(([n, nx, tx, tw]) => {
      text(s, n, { x: nx, y: 5.386, w: 0.844, h: 0.774, fontSize: 40, fontFace: HEAD });
      caption(s, LOREM.sitDo, tx, 5.429, tw);
    });
}

/* ============================================================== slide 08 */
function slide08(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 8);
  s.addShape('rect', { x: 0.971, y: 2.163, w: 8.064, h: 4.087,
    fill: { color: C.white, transparency: 95 }, line: { color: C.white, width: 0.5, transparency: 80 } });
  text(s, 'Performance Report', { x: 1.878, y: 0.642, w: 9.577, h: 0.919, fontSize: 54,
    fontFace: HEAD, align: 'center', lineSpacingMultiple: 0.9 });
  s.addChart('line', [{
    name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4', 'Category 5'],
    values: [1, 1, 2, 3, 4],
  }], {
    x: 1.165, y: 2.441, w: 7.732, h: 3.166,
    chartColors: [C.grey], lineSize: 2.25, lineDataSymbol: 'none', showLegend: false,
    catAxisHidden: true, valAxisLabelColor: C.white, valAxisLabelFontSize: 12,
    valAxisLineShow: false, valAxisMaxVal: 5, valAxisMajorUnit: 1,
    valGridLine: { color: '333333', size: 0.75 }, catGridLine: { style: 'none' },
  });
  // year markers: grey stub at 2023, indigo riser + bubble at 2026
  s.addShape('rect', { x: 3.632, y: 4.779, w: 0.1, h: 0.575, fill: { color: C.midGrey } });
  s.addShape('flowChartConnector', { x: 3.607, y: 4.759, w: 0.15, h: 0.15, fill: { color: C.white } });
  s.addShape('rect', { x: 8.054, y: 2.643, w: 0.123, h: 2.788, fill: { color: C.indigo } });
  s.addShape('flowChartConnector', { x: 7.699, y: 2.518, w: 0.843, h: 0.843, fill: { color: C.indigo } });
  text(s, '75%', { x: 7.699, y: 2.518, w: 0.843, h: 0.843, fontSize: 14, fontFace: HEAD,
    align: 'center', valign: 'middle' });
  text(s, '2023', { x: 3.388, y: 5.592, w: 0.588, h: 0.303, fontSize: 12, fontFace: HEAD });
  text(s, '2026', { x: 7.822, y: 5.592, w: 0.588, h: 0.303, fontSize: 12, fontFace: HEAD });
  // right hand "Overview" panel
  gradient(s, { x: 9.329, y: 2.163, w: 3.034, h: 4.087, angle: 45, stops: SWEEP });
  text(s, 'Overview', { x: 9.688, y: 2.395, w: 1.362, h: 0.404, fontSize: 18, fontFace: HEAD });
  text(s, '75%', { x: 9.673, y: 3.155, w: 1.535, h: 0.774, fontSize: 40, fontFace: HEAD });
  caption(s, LOREM.sit, 9.687, 3.874, 2.548, { transparency: 0, h: 0.387 });
  s.addShape('line', { x: 9.673, y: 4.579, w: 2.272, h: 0,
    line: { color: C.silver, width: 1, transparency: 70 } });
  text(s, '28%', { x: 9.673, y: 4.786, w: 1.535, h: 0.774, fontSize: 40, fontFace: HEAD });
  caption(s, LOREM.sit, 9.687, 5.504, 2.548, { transparency: 0, h: 0.387 });
}

/* ============================================================== slide 09 */
const PIE_LEGEND = [
  { x: 1.23, tx: 1.54, y: 4.751, color: C.indigo },
  { x: 1.23, tx: 1.54, y: 5.368, color: C.orchid },
  { x: 3.293, tx: 3.603, y: 4.751, color: C.lilac },
  { x: 3.293, tx: 3.603, y: 5.368, color: C.white },
];
function slide09(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 9);
  text(s, 'Impact Overview', { x: 1.133, y: 1.313, w: 5.43, h: 1.919, fontSize: 60,
    fontFace: HEAD, lineSpacingMultiple: 0.9 });
  caption(s, LOREM.full, 1.133, 3.441, 4.513, { h: 0.693 });
  text(s, 'Title Chart', { x: 8.689, y: 1.108, w: 1.975, h: 0.438, fontSize: 20,
    fontFace: HEAD, align: 'center', color: C.mist });
  s.addChart('pie', [{
    name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2],
  }], {
    x: 6.667, y: 1.887, w: 6.121, h: 4.165, showLegend: false,
    chartColors: [C.indigo, C.orchid, C.lilac, C.white], firstSliceAng: 0,
  });
  for (const l of PIE_LEGEND) {
    s.addShape('rect', { x: l.x, y: l.y, w: 0.207, h: 0.207, fill: { color: l.color } });
    text(s, 'Title Here', { x: l.tx, y: l.y - 0.065, w: 1.138, h: 0.337, fontSize: 14, fontFace: HEAD });
  }
}

/* ============================================================== slide 10 */
const BARS = [
  { label: 'Overall Progress', pct: '80%', y: 4.493, fill: 3.321, labelY: 4.024, lw: 2.196 },
  { label: 'Execution Effectiveness', pct: '67%', y: 5.451, fill: 2.52, labelY: 4.951, lw: 2.754 },
];
const VENN = [
  { x: 7.859, y: 0.743, angle: 270, label: 'Cross-Functional Collaboration', tx: 8.17, ty: 1.963, tw: 2.859 },
  { x: 6.501, y: 2.933, angle: 160, label: 'The Strategic Framework', tx: 6.963, ty: 4.348, tw: 2.191 },
  { x: 9.179, y: 2.933, angle: 40, label: 'Data-Driven Evaluation Process', tx: 10.073, ty: 4.348, tw: 2.191 },
];
/* accent1 fading from 40% down to fully transparent across the disc */
const VENN_WASH = [[0, C.indigo, 60], [0.39, C.indigo, 80], [0.83, C.indigo, 100]];
function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 10);
  text(s, 'A Company Insight ', { x: 0.679, y: 1.373, w: 5.054, h: 1.737, fontSize: 54,
    fontFace: HEAD, lineSpacingMultiple: 0.9 });
  for (const b of BARS) {
    text(s, b.label, { x: 0.718, y: b.labelY, w: b.lw, h: 0.337, fontSize: 14, fontFace: HEAD });
    text(s, b.pct, { x: 4.533, y: b.labelY - 0.031, w: 0.746, h: 0.387, fontSize: 14,
      transparency: 30, lineSpacingMultiple: 1.3 });
    s.addShape('rect', { x: 0.833, y: b.y, w: 4.188, h: 0.125, fill: { color: C.mist, transparency: 90 } });
    gradient(s, { x: 0.833, y: b.y, w: b.fill, h: 0.125, angle: 45, stops: SWEEP });
  }
  for (const v of VENN) {
    gradientEllipse(s, { x: v.x, y: v.y, w: 3.48, h: 3.239, angle: v.angle, stops: VENN_WASH });
    text(s, v.label, { x: v.tx, y: v.ty, w: v.tw, h: 0.507, fontSize: 14, align: 'center',
      transparency: 30, lineSpacingMultiple: 0.9 });
  }
}

/* ============================================================== slide 11 */
const ROADMAP = [
  { month: 'Jan', x: 0.917 }, { month: 'Feb', x: 3.24, active: true },
  { month: 'March', x: 5.562 }, { month: 'Apr', x: 7.885 }, { month: 'May', x: 10.208 },
];
function slide11(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 11);
  text(s, 'This Year Roadmap Project', { x: 3.409, y: 0.63, w: 6.497, h: 1.616, fontSize: 50,
    fontFace: HEAD, align: 'center', lineSpacingMultiple: 0.9 });
  for (const card of ROADMAP) {
    if (card.active) {
      gradient(s, { x: card.x, y: 2.85, w: 2.208, h: 2.683, angle: 45, stops: SWEEP });
    } else {
      s.addShape('rect', { x: card.x, y: 2.85, w: 2.208, h: 2.683,
        fill: { color: C.white, transparency: 95 }, line: { color: C.white, width: 0.5, transparency: 80 } });
    }
    text(s, card.month, { x: card.x + 0.187, y: 3.124, w: 1.833, h: 0.75, fontSize: 32,
      fontFace: HEAD, lineSpacingMultiple: 1.3, paraSpaceBefore: 12 });
    text(s, 'Title Here', { x: card.x + 0.187, y: 4.192, w: 1.138, h: 0.337, fontSize: 14, fontFace: HEAD });
    caption(s, LOREM.seDo, card.x + 0.17, 4.513, 2.038);
  }
  // timeline rail below the cards
  for (let i = 0; i < 4; i++) {
    s.addShape('line', { x: 2.208 + i * 2.334, y: 6.102, w: 1.958, h: 0,
      line: { color: C.silver, width: 1, transparency: 70 } });
  }
  [1.922, 4.255, 6.589, 8.922, 11.255].forEach((x, i) => {
    s.addShape('ellipse', { x: x, y: 6.003, w: 0.198, h: 0.198,
      fill: { color: i === 1 ? C.mist : 'F2F2F2' } });
  });
}

/* ============================================================== slide 12 */
function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 12);
  freeform(s, LOOP, { x: 5.963, y: 1.52, w: 10.476, h: 3.49,
    opts: { rotate: 221.5, fill: { color: RIBBON.quoteTop } } });
  freeform(s, KNOT, { x: -1.241, y: 4.275, w: 5.288, h: 3.321,
    opts: { rotate: 18.6, fill: { color: RIBBON.quoteLow } } });
  logo(s, 0.717, 0.46, C.grey);
  text(s, 'Let\u2019s Driving Sustainable Progress Through Clear Comprehensive.',
    { x: 1.874, y: 2.104, w: 9.586, h: 2.827, fontSize: 54, fontFace: HEAD, align: 'center' });
  text(s, '2026', { x: 12.058, y: 0.557, w: 0.651, h: 0.337, fontSize: 14 });
}

/* ============================================================== slide 13 */
function slide13(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 13);
  indentedTitle(s, 'Next Target Progress', { x: 1.177, y: 0.89, w: 10.455, h: 0.919,
    fontSize: 54, align: 'center' });
  // the two open C-arms that sweep out of the dial towards the side icons
  freeform(s, TRACK, { x: 5.343, y: 2.454, w: 7.031, h: 1.893,
    opts: { fill: { color: RIBBON.track } } });
  freeform(s, TRACK, { x: 0.958, y: 4.215, w: 7.031, h: 1.893,
    opts: { flipH: true, flipV: true, fill: { color: RIBBON.track } } });
  s.addShape('flowChartConnector', { x: 5.378, y: 2.993, w: 2.578, h: 2.578,
    fill: { color: C.ink, transparency: 69 }, line: { color: C.white, width: 0.5, transparency: 70 } });
  s.addChart('doughnut', [{
    name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [70, 30],
  }], {
    x: 4.903, y: 3.085, w: 3.528, h: 2.352, showLegend: false, holeSize: 75,
    chartColors: [mixHex(C.indigo, C.orchid, 0.45), '404040'], firstSliceAng: 0,
  });
  text(s, '70%', { x: 6.226, y: 4.009, w: 0.882, h: 0.505, fontSize: 24, fontFace: HEAD });
  // icon tiles at each end of the track: a target ring and a bar-chart glyph
  s.addShape('rect', { x: 0.924, y: 3.759, w: 0.919, h: 0.919, fill: { color: C.orchid } });
  s.addShape('ellipse', { x: 1.229, y: 4.064, w: 0.31, h: 0.31,
    fill: { type: 'none' }, line: { color: C.white, width: 1.25 } });
  s.addShape('ellipse', { x: 1.329, y: 4.164, w: 0.11, h: 0.11, fill: { color: C.white } });
  s.addShape('rect', { x: 11.486, y: 3.759, w: 0.919, h: 0.919, fill: { color: C.orchid } });
  [[0, 0.12], [0.12, 0.2], [0.24, 0.31]].forEach(([dx, bh]) => {
    s.addShape('rect', { x: 11.79 + dx, y: 4.375 - bh + 0.31, w: 0.08, h: bh, fill: { color: C.white } });
  });
  text(s, 'Title Here', { x: 0.81, y: 5.046, w: 1.138, h: 0.337, fontSize: 14, fontFace: HEAD });
  caption(s, LOREM.seDo, 0.81, 5.399, 2.038);
  text(s, 'Title Here', { x: 11.39, y: 5.064, w: 1.138, h: 0.337, fontSize: 14, fontFace: HEAD, align: 'right' });
  caption(s, LOREM.seDo, 10.495, 5.417, 2.038, { align: 'right' });
}

/* ============================================================== slide 14 */
const STEPS = [
  { n: '01', cx: 1.94, tx: 2.386, subX: 1.886, capX: 1.61 },
  { n: '02', cx: 4.552, tx: 4.955, subX: 4.415, capX: 4.139 },
  { n: '03', cx: 7.011, tx: 7.426, subX: 6.943, capX: 6.667 },
  { n: '04', cx: 9.576, tx: 9.96, subX: 9.471, capX: 9.194 },
];
function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 14);
  indentedTitle(s, 'Sustainable Objective', { x: 1.163, y: 0.86, w: 10.455, h: 0.858,
    fontSize: 50, align: 'center' });
  gradientPill(s, { x: 1.518, y: 2.367, w: 10.297, h: 2.364, angle: 45, stops: SWEEP7 });
  for (const st of STEPS) {
    s.addShape('flowChartConnector', { x: st.cx, y: 2.665, w: 1.814, h: 1.814, fill: { color: C.white } });
    text(s, st.n, { x: st.cx, y: 2.665, w: 1.814, h: 1.814, fontSize: 44, fontFace: HEAD,
      color: C.ink, align: 'center', valign: 'middle' });
    text(s, 'Sub Title Here', { x: st.subX, y: 5.067, w: 1.976, h: 0.37, fontSize: 16,
      fontFace: HEAD, align: 'center' });
    caption(s, LOREM.consectetur, st.capX, 5.54, 2.528, { align: 'center', h: 0.693 });
  }
}

/* ============================================================== slide 15 */
function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 15);
  freeform(s, LOOP, { x: 9.327, y: 0.294, w: 5.93, h: 1.975,
    opts: { rotate: 30.2, fill: { color: RIBBON.vision } } });
  indentedTitle(s, 'Sustainability Company Vision', { x: 4.858, y: 1.196, w: 7.2, h: 1.737,
    fontSize: 54 });
  gradient(s, { x: 5.172, y: 3.963, w: 7.503, h: 1.009, angle: 45, stops: SWEEP });
  numberedRow(s, 4.149, '01', 'Integrated Approach', 5.355);
  numberedRow(s, 5.293, '02', 'Long Term Value Creation', 5.355);
}

/* ============================================================== slide 16 */
function slide16(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 16);
  text(s, 'Project Worldwide', { x: 0.7, y: 1.469, w: 5.43, h: 1.737, fontSize: 54,
    fontFace: HEAD, lineSpacingMultiple: 0.9 });
  caption(s, LOREM.full, 0.7, 3.353, 4.513, { h: 0.693 });
  for (const isle of MAP_ISLANDS) {
    polygons(s, isle, isle.polys, { fill: { color: C.white, transparency: 85 },
      line: { color: C.white, width: 1.5, transparency: 80 } });
  }
  // two callout bubbles wired back to the map with dashed elbows
  s.addShape('bentConnector3', { x: 8.471, y: 2.065, w: 1.715, h: 2.493, rotate: 180, flipV: true,
    line: { color: C.orchid, width: 0.75, dashType: 'dash', beginArrowType: 'oval', endArrowType: 'oval' } });
  s.addShape('bentConnector3', { x: 7.741, y: 4.062, w: 2.73, h: 1.286, rotate: 180, flipV: true,
    line: { color: C.white, width: 0.75, dashType: 'dash', beginArrowType: 'oval', endArrowType: 'oval' } });
  gradientEllipse(s, { x: 10.448, y: 1.236, w: 1.738, h: 1.738, angle: 45, stops: SWEEP });
  text(s, '75%', { x: 10.687, y: 1.752, w: 1.26, h: 0.707, fontSize: 36, fontFace: HEAD, align: 'center' });
  s.addShape('ellipse', { x: 10.65, y: 3.477, w: 1.135, h: 1.135, fill: { color: C.midGrey } });
  text(s, '35%', { x: 10.735, y: 3.792, w: 0.965, h: 0.505, fontSize: 24, fontFace: HEAD, align: 'center' });
  // KPI pair under the body copy
  s.addShape('triangle', { x: 0.914, y: 4.973, w: 0.213, h: 0.221, fill: { color: C.indigo } });
  text(s, '75%', { x: 1.273, y: 4.82, w: 1.238, h: 0.586, fontSize: 32, fontFace: HEAD,
    color: mixHex(C.indigo, C.orchid, 0.35), lineSpacingMultiple: 0.9 });
  text(s, 'Primary Region', { x: 0.8, y: 5.462, w: 1.684, h: 0.337, fontSize: 14, transparency: 30 });
  s.addShape('line', { x: 2.959, y: 4.82, w: 0, h: 0.93, line: { color: C.silver, width: 1, transparency: 70 } });
  s.addShape('triangle', { x: 3.434, y: 4.991, w: 0.213, h: 0.221, rotate: 180, fill: { color: C.white } });
  text(s, '35%', { x: 3.788, y: 4.82, w: 1.254, h: 0.586, fontSize: 32, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  text(s, 'Secondary Region', { x: 3.318, y: 5.462, w: 1.948, h: 0.337, fontSize: 14, transparency: 30 });
}

/* ============================================================== slide 17 */
const FUTURE = [
  { n: '01', title: 'Industry Alignment', y: 1.162 },
  { n: '02', title: 'Growth Potential', y: 3.07 },
  { n: '03', title: 'ESG Ready', y: 4.978 },
];
function slide17(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 17);
  freeform(s, LOOP, { x: -0.739, y: -0.711, w: 8.689, h: 2.895,
    opts: { rotate: 318.1, fill: { color: RIBBON.future } } });
  gradient(s, { x: 6.667, y: 0, w: 4.014, h: 7.5, angle: 45, stops: SWEEP6 });
  text(s, 'Future Target Company', { x: 0.7, y: 1.235, w: 5.105, h: 2.827, fontSize: 60,
    fontFace: HEAD, lineSpacingMultiple: 0.9 });
  caption(s, LOREM.magna, 0.7, 5.029, 4.014, { h: 0.994 });
  FUTURE.forEach((f, i) => {
    if (i > 0) {
      s.addShape('line', { x: 7.048, y: [0, 2.708, 4.616][i], w: 3.188, h: 0,
        line: { color: C.white, width: 1, transparency: 60 } });
    }
    text(s, f.n, { x: 7.048, y: f.y + 0.016, w: 0.458, h: 0.37, fontSize: 16, fontFace: HEAD });
    text(s, f.title, { x: 7.64, y: f.y, w: 2.6, h: 0.404, fontSize: 18, fontFace: HEAD });
    caption(s, LOREM.sed, 7.05, f.y + 0.48, 3.096);
  });
}

/* ============================================================== slide 18 */
const GANTT = [
  { label: 'Foundation Strength ', x: 1.173, w: 2.945, y: 4.063 },
  { label: 'Product Enhancement', x: 3.561, w: 3.036, y: 4.674 },
  { label: 'Growth Acceleration', x: 5.686, w: 2.964, y: 5.301 },
  { label: 'Expansion & Diversification', x: 8.737, w: 3.436, y: 5.656, highlight: true },
];
const YEARS = [
  { y: '2026', lx: 0.691, gx: 1.038 }, { y: '2027', lx: 3.136, gx: 3.483 },
  { y: '2028', lx: 5.763, gx: 6.11 }, { y: '2029', lx: 8.207, gx: 8.554 },
  { y: '2030', lx: 10.802, gx: 11.127 },
];
function slide18(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 18);
  text(s, 'Sustainable Company Roadmap', { x: 2.831, y: 0.655, w: 7.755, h: 1.737,
    fontSize: 54, fontFace: HEAD, align: 'center', lineSpacingMultiple: 0.9 });
  for (const yr of YEARS) {
    text(s, yr.y, { x: yr.lx, y: 3.266, w: 0.946, h: 0.429, fontSize: 14, fontFace: HEAD,
      align: 'center', valign: 'middle', transparency: 20 });
    s.addShape('line', { x: yr.gx, y: 3.845, w: 0, h: yr.y === '2030' ? 2.3 : 2.396,
      line: { color: C.silver, width: 1, transparency: 80 } });
  }
  for (const bar of GANTT) {
    if (bar.highlight) gradient(s, { x: bar.x, y: bar.y, w: bar.w, h: 0.531, angle: 45, stops: SWEEP6 });
    else s.addShape('rect', { x: bar.x, y: bar.y, w: bar.w, h: 0.531,
      fill: { color: C.white, transparency: 85 }, line: { color: C.white, width: 0.5, transparency: 80 } });
    text(s, bar.label, { x: bar.x, y: bar.y, w: bar.w, h: 0.531, fontSize: 14,
      align: 'center', valign: 'middle' });
  }
}

/* ============================================================== slide 19 */
const CONCLUSION = [
  { n: '01', label: 'Clear Business Impact', y: 4.012, lx: 5.225, lw: 2.737 },
  { n: '02', label: 'Operational Efficiency', y: 4.817, lx: 5.201, lw: 2.69 },
  { n: '03', label: 'Smarter, Data-Driven Decisions', y: 5.54, lx: 4.621, lw: 3.771, active: true },
];
function slide19(pres) {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  footer(s, 19);
  freeform(s, LOOP, { x: 7.34, y: -0.222, w: 8.756, h: 2.917,
    opts: { rotate: 212.5, fill: { color: RIBBON.conclusion } } });
  text(s, 'About Company Conclution', { x: 0.699, y: 0.997, w: 7.231, h: 1.979,
    fontSize: 62, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  gradient(s, { x: 0.804, y: 5.388, w: 11.729, h: 0.723, angle: 45, stops: SWEEP });
  s.addShape('line', { x: 0.8, y: 4.582, w: 11.733, h: 0,
    line: { color: C.silver, width: 1, transparency: 70 } });
  for (const row of CONCLUSION) {
    text(s, row.n, { x: 0.908, y: row.y, w: 0.489, h: 0.404, fontSize: 18, fontFace: HEAD });
    text(s, row.label, { x: row.lx, y: row.y, w: row.lw, h: 0.404, fontSize: 18 });
    s.addShape('line', { x: 11.93, y: row.y + 0.075, w: 0.26, h: 0.26, flipV: true,
      line: { color: C.white, width: 1.25, endArrowType: 'triangle' } });
  }
}

/* ============================================================== slide 20 */
function slide20(pres) {
  const s = pres.addSlide();
  gradient(s, { x: 0, y: 0, w: 13.334, h: 7.5, angle: 45, stops: NIGHT });
  freeform(s, WAVE, { x: 0.275, y: -0.096, w: 14.628, h: 8.744,
    opts: { rotate: 18.2, flipV: true, fill: { color: C.white, transparency: 91 } } });
  logo(s, 5.894, 1.569);
  text(s, 'THANKS FOR YOUR ATTENTION', { x: 2.009, y: 2.375, w: 9.316, h: 2.625,
    fontSize: 75, fontFace: HEAD, align: 'center' });
  caption(s, LOREM.full, 4.41, 5.153, 4.513, { align: 'center', h: 0.693 });
}

/* ==================================================================== main */
function build() {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';                    // 13.333 x 7.5 in
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Sustainable Business Report';
  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pres));
  return pres.writeFile({
    fileName: path.join(__dirname, '113bdde4-b1bc-4385-9aff-b0aa62dc5359_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f));
