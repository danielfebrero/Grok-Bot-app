/**
 * Recreates the "Perfume" presentation template (20 slides, 10 x 5.625 in)
 * using pptxgenjs only.  Photographic images from the source deck are
 * replaced with light-grey "[image]" placeholder blocks.
 *
 *   node 0adeeea2-9d64-42d0-8bd9-9bce2afe6fb6_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const BLUE = '295F98'; // headings / accents
const GREY = '7F7F7F'; // body copy
const SAND = 'DDD9CD'; // card panels
const CREAM = 'EAE6DA'; // slide background
const IVORY = 'F4F2EC'; // giant section numerals
const KHAKI = 'CDC2A5'; // strong accent panels / rules
const BROWN = '766842'; // muted button label + page numbers
const WHITE = 'FFFFFF';
const DECOR = 'EBE7DB'; // decorative corner blobs
const PHOTO = 'F2F2F2'; // image placeholder fill
const PHOTO_TX = '9A9A9A'; // image placeholder caption
const ICE = 'E8EAED'; // check-mark glyph

const BOLD_FONT = 'Bricolage Grotesque SemiBold';
const LIGHT_FONT = 'Bricolage Grotesque Light';

/* Google-Slides text-box insets used throughout the deck (points). */
const INSET = [5.4, 5.4, 2.7, 2.7]; // [left, right, bottom, top]

/* The three drop shadows used by the original deck.  These are factories
 * because pptxgenjs rewrites the shadow object it is handed (points -> EMU),
 * so every shape needs its own copy. */
const SOFT = () => ({ type: 'outer', color: GREY, opacity: 0.2, blur: 30, offset: 20, angle: 330 });
const LIFT = (color) => ({ type: 'outer', color: color || '434343', opacity: 0.2, blur: 33.75, offset: 1.5, angle: 90 });
const DROP = () => ({ type: 'outer', color: '000000', opacity: 0.149, blur: 19, offset: 11, angle: 359 });

/* ------------------------------------------------- custom geometry outlines */
/* Every path is stored normalised to the unit square:                        */
/*   ['m', x, y]                      move-to                                 */
/*   ['l', x, y]                      line-to                                 */
/*   ['c', x1, y1, x2, y2, x, y]      cubic bezier                            */
/*   ['z']                            close                                   */

const BOTTLE = [['m',.8649,.2905],['l',.9495,.1609],['c',.9839,.1082,.9609,.0421,.8981,.0133],
  ['c',.8791,.0045,.8577,-.0001,.8359,-.0001],['l',.1641,-.0001],['c',.0925,-.0001,.0345,.0486,.0345,.1087],
  ['c',.0345,.1269,.04,.1449,.0505,.1609],['l',.1351,.2905],['c',.0627,.361,-.0055,.459,.0003,.5824],
  ['c',.0055,.6952,.055,.7988,.1395,.8739],['c',.2311,.9554,.3555,.9989,.4995,.9999],['l',.5003,.9999],
  ['c',.6443,.9989,.769,.9554,.8603,.8739],['c',.9448,.7986,.9943,.695,.9996,.5824],
  ['c',1.0055,.459,.9372,.361,.8649,.2905],['z'],['m',.1255,.596],['l',.3443,.596],
  ['c',.358,.5961,.369,.6055,.3689,.617],['c',.3688,.6207,.3677,.6242,.3654,.6274],['l',.235,.8075],
  ['c',.1691,.7512,.1333,.6741,.1255,.5961],['z'],['m',.5,.8958],['c',.4375,.8953,.3832,.885,.3368,.8673],
  ['l',.4714,.6816],['c',.5139,.6226,.4915,.5458,.4212,.51],['c',.398,.4982,.3714,.492,.3443,.4919],
  ['l',.1377,.4919],['c',.1764,.3852,.2865,.3071,.3397,.2694],['c',.3502,.2619,.3585,.256,.3653,.2507],
  ['c',.4163,.2107,.4266,.1703,.4279,.1485],['l',.3027,.1485],['c',.3005,.1542,.2948,.1636,.2807,.1747],
  ['c',.2762,.1781,.2686,.1836,.2597,.1899],['c',.25,.1968,.2387,.2048,.2264,.2139],['l',.1592,.1109],
  ['c',.1581,.1095,.1581,.1077,.1592,.1063],['c',.1601,.1048,.162,.1039,.164,.104],['l',.8359,.104],
  ['c',.839,.104,.8415,.1061,.8415,.1086],['c',.8415,.1094,.8412,.1102,.8408,.1109],['l',.7735,.2139],
  ['c',.7611,.2048,.7499,.1968,.7403,.1899],['c',.7314,.1837,.7237,.1782,.7193,.1747],
  ['c',.7051,.1636,.6995,.1542,.6973,.1486],['l',.572,.1486],['c',.5734,.1704,.5836,.2108,.6347,.2508],
  ['c',.6414,.2561,.6497,.262,.6602,.2695],['c',.7134,.3072,.8235,.3853,.8622,.492],['l',.6541,.492],
  ['c',.572,.4921,.5055,.5482,.5056,.6172],['c',.5056,.6399,.5131,.6622,.5271,.6817],['l',.662,.868],
  ['c',.616,.8853,.562,.8954,.5,.8959],['z'],['m',.7641,.8083],['l',.6331,.6274],
  ['c',.6261,.6175,.63,.6047,.6417,.5989],['c',.6455,.597,.6498,.596,.6541,.596],['l',.8745,.596],
  ['c',.8666,.6745,.8306,.7519,.7641,.8084],['z']];

const ARCH_TALL = [['m',0,0],['l',1,0],['l',1,.6757],['c',1,.8548,.7761,1,.5,1],['c',.2239,1,0,.8548,0,.6757],
  ['c',0,.4505,0,.2252,0,0],['z']];

const BRACKET = [['m',.9996,.9999],['l',.5712,.9999],['c',.2558,.9995,.0003,.733,-.0001,.4041],
  ['l',-.0001,-.0001],['l',.2731,-.0001],['l',.2731,.4041],['c',.2735,.5756,.4067,.7145,.5712,.715],
  ['l',.9999,.715],['z']];

const HILL = [['m',0,.9644],['c',0,.9644,.1565,1.0961,.2648,.853],['c',.3731,.6099,.4812,.7264,.554,.7821],
  ['c',.6268,.8378,.7351,.7922,.7876,.5035],['c',.8402,.2149,1.0415,.5146,.9923,0],['l',0,0],['z']];

const TIMELINE = [['m',0,.0003],['l',.8975,0],['c',.9384,0,.9715,.1096,.9715,.2447],['l',.9715,.2448],
  ['l',.9715,.2448],['c',.9715,.3799,.9384,.4894,.8975,.4894],['l',.1026,.4894],
  ['c',.0617,.4894,.0286,.5989,.0285,.734],['l',.0285,.7341],['l',.0285,.7552],
  ['c',.0285,.8904,.0617,1,.1026,1],['l',.1026,1],['l',1,.9989]];

const ICON_STYLE = [['m',.192,.755],['c',.235,.722,.282,.697,.335,.678],['c',.388,.659,.442,.65,.5,.65],
  ['c',.558,.65,.612,.659,.665,.678],['c',.718,.697,.765,.722,.808,.755],['c',.837,.721,.859,.682,.876,.639],
  ['c',.892,.595,.9,.549,.9,.5],['c',.9,.389,.861,.295,.783,.217],['c',.705,.139,.611,.1,.5,.1],
  ['c',.389,.1,.295,.139,.217,.217],['c',.139,.295,.1,.389,.1,.5],['c',.1,.549,.108,.595,.124,.639],
  ['c',.141,.682,.163,.721,.192,.755],['z'],['m',.5,.55],['c',.451,.55,.409,.533,.376,.499],
  ['c',.342,.466,.325,.424,.325,.375],['c',.325,.326,.342,.284,.376,.251],['c',.409,.217,.451,.2,.5,.2],
  ['c',.549,.2,.591,.217,.624,.251],['c',.658,.284,.675,.326,.675,.375],['c',.675,.424,.658,.466,.624,.499],
  ['c',.591,.533,.549,.55,.5,.55],['z'],['m',.5,1],['c',.431,1,.366,.987,.305,.961],
  ['c',.244,.934,.191,.899,.146,.854],['c',.101,.809,.066,.756,.039,.695],['c',.013,.634,0,.569,0,.5],
  ['c',0,.431,.013,.366,.039,.305],['c',.066,.244,.101,.191,.146,.146],['c',.191,.101,.244,.066,.305,.039],
  ['c',.366,.013,.431,0,.5,0],['c',.569,0,.634,.013,.695,.039],['c',.756,.066,.809,.101,.854,.146],
  ['c',.899,.191,.934,.244,.961,.305],['c',.987,.366,1,.431,1,.5],['c',1,.569,.987,.634,.961,.695],
  ['c',.934,.756,.899,.809,.854,.854],['c',.809,.899,.756,.934,.695,.961],['c',.634,.987,.569,1,.5,1],['z'],
  ['m',.5,.9],['c',.544,.9,.586,.894,.625,.881],['c',.664,.868,.7,.849,.732,.825],
  ['c',.7,.801,.664,.782,.625,.769],['c',.586,.756,.544,.75,.5,.75],['c',.456,.75,.414,.756,.375,.769],
  ['c',.336,.782,.3,.801,.268,.825],['c',.3,.849,.336,.868,.375,.881],['c',.414,.894,.456,.9,.5,.9],['z'],
  ['m',.5,.45],['c',.522,.45,.54,.443,.554,.429],['c',.568,.415,.575,.397,.575,.375],
  ['c',.575,.353,.568,.335,.554,.321],['c',.54,.307,.522,.3,.5,.3],['c',.478,.3,.46,.307,.446,.321],
  ['c',.432,.335,.425,.353,.425,.375],['c',.425,.397,.432,.415,.446,.429],['c',.46,.443,.478,.45,.5,.45],['z']];

const ICON_EMOTION = [['m',.488,.676],['c',.549,.676,.601,.658,.644,.62],['c',.687,.582,.709,.537,.709,.484],
  ['c',.709,.436,.693,.396,.661,.363],['c',.629,.33,.59,.314,.544,.314],['c',.503,.314,.468,.326,.44,.351],
  ['c',.411,.376,.397,.407,.397,.444],['c',.397,.46,.4,.475,.407,.49],['c',.414,.505,.423,.519,.435,.531],
  ['l',.51,.46],['c',.508,.458,.506,.456,.504,.454],['c',.503,.451,.502,.448,.502,.445],
  ['c',.502,.436,.506,.429,.514,.423],['c',.522,.418,.532,.415,.544,.415],['c',.562,.415,.576,.422,.588,.436],
  ['c',.599,.449,.605,.466,.605,.485],['c',.605,.511,.594,.533,.571,.551],['c',.549,.569,.522,.578,.489,.578],
  ['c',.448,.578,.413,.562,.385,.53],['c',.356,.498,.342,.46,.342,.414],['c',.342,.39,.347,.366,.356,.344],
  ['c',.366,.322,.38,.303,.397,.286],['l',.322,.215],['c',.294,.241,.273,.271,.258,.305],
  ['c',.243,.339,.235,.375,.235,.412],['c',.235,.486,.26,.548,.309,.599],['c',.358,.651,.418,.676,.488,.676],
  ['z'],['m',.158,1],['l',.158,.785],['c',.108,.742,.069,.691,.041,.633],['c',.014,.575,0,.514,0,.45],
  ['c',0,.325,.046,.219,.138,.131],['c',.23,.044,.342,0,.473,0],['c',.583,0,.68,.031,.765,.092],
  ['c',.849,.153,.904,.233,.93,.331],['l',.998,.588],['c',1.002,.603,.999,.618,.989,.631],
  ['c',.978,.644,.964,.65,.947,.65],['l',.842,.65],['l',.842,.8],['c',.842,.828,.831,.851,.811,.871],
  ['c',.79,.89,.765,.9,.736,.9],['l',.631,.9],['l',.631,1],['l',.526,1],['l',.526,.8],['l',.736,.8],
  ['l',.736,.55],['l',.878,.55],['l',.828,.356],['c',.808,.28,.765,.219,.7,.171],
  ['c',.634,.124,.558,.1,.473,.1],['c',.372,.1,.285,.134,.213,.201],['c',.141,.269,.105,.351,.105,.448],
  ['c',.105,.498,.116,.545,.137,.59],['c',.159,.635,.189,.675,.229,.71],['l',.263,.74],['l',.263,1],
  ['l',.158,1],['z']];

const ICON_LEAF = [['m',.5,1],['c',.439,.992,.379,.976,.319,.951],['c',.259,.925,.205,.889,.158,.841],
  ['c',.111,.794,.073,.734,.044,.661],['c',.015,.589,0,.502,0,.4],['l',0,.35],['l',.05,.35],
  ['c',.092,.35,.136,.355,.181,.366],['c',.226,.377,.268,.393,.308,.415],['c',.318,.343,.34,.27,.376,.194],
  ['c',.411,.119,.452,.054,.5,0],['c',.548,.054,.589,.119,.624,.194],['c',.66,.27,.682,.343,.692,.415],
  ['c',.732,.393,.774,.377,.819,.366],['c',.864,.355,.908,.35,.95,.35],['l',1,.35],['l',1,.4],
  ['c',1,.502,.985,.589,.956,.661],['c',.927,.734,.889,.794,.842,.841],['c',.795,.889,.741,.925,.682,.951],
  ['c',.622,.976,.562,.992,.5,1],['z'],['m',.498,.898],['c',.488,.759,.447,.655,.374,.584],
  ['c',.301,.513,.211,.469,.102,.452],['c',.112,.595,.154,.701,.229,.771],['c',.305,.841,.394,.883,.498,.898],
  ['z'],['m',.5,.58],['c',.512,.562,.528,.543,.546,.523],['c',.564,.504,.581,.487,.598,.472],
  ['c',.596,.425,.586,.375,.569,.324],['c',.552,.272,.529,.222,.5,.172],['c',.471,.222,.448,.272,.431,.324],
  ['c',.414,.375,.404,.425,.402,.472],['c',.419,.487,.437,.504,.455,.523],['c',.473,.543,.488,.562,.5,.58],
  ['z'],['m',.598,.875],['c',.628,.865,.66,.85,.694,.831],['c',.727,.812,.758,.786,.787,.753],
  ['c',.816,.72,.84,.679,.861,.63],['c',.881,.581,.893,.522,.898,.452],['c',.819,.464,.75,.49,.691,.531],
  ['c',.632,.571,.587,.622,.555,.685],['c',.565,.712,.574,.741,.581,.772],['c',.588,.804,.593,.838,.598,.875],
  ['z']];

const ARCH_SMALL = [['m',.5,0],['c',.7761,0,1,.1851,1,.4135],['c',1,.609,1,.8045,1,1],['l',0,1],['l',0,.4135],
  ['c',0,.1851,.2239,0,.5,0],['z']];

const CHECK = [['m',.3497,1],['l',0,.526],['l',.0874,.4075],['l',.3497,.763],['l',.9126,0],['l',1,.1185],
  ['l',.3497,1],['z']];

const WAVE_BLOB = [['m',-.0001,.9998],['l',.9999,.9998],['c',.9999,.9998,.9514,.8953,.9151,.6815],
  ['c',.8787,.4678,.8507,.3807,.712,.3742],['c',.5734,.3676,.5574,-.064,.4482,.0079],
  ['c',.339,.0799,.3891,.3911,.2505,.5461],['c',.1119,.7012,.0174,.6053,-.0001,.9998],['z']];

const WAVE_TOP = [['m',.9497,0],['c',.9819,.0007,.9999,.0394,.9999,.0394],['l',.9999,.6237],['l',1,.6237],
  ['l',1,1],['l',0,1],['l',0,.7309],['l',0,.7309],['l',0,.0397],['c',.0575,.001,.0638,.0756,.0958,.0867],
  ['c',.1279,.0978,.1432,.0704,.1943,.0342],['c',.2455,-.002,.2787,.0728,.3324,.1226],
  ['c',.3862,.1724,.4002,.117,.4475,.0866],['c',.4949,.0562,.5114,.1446,.5562,.1198],
  ['c',.6009,.095,.615,-.0129,.7263,.048],['c',.8375,.1089,.858,.0591,.9129,.0148],
  ['c',.9267,.0038,.939,-.0002,.9497,0],['z']];

/* Decorative corner shapes, traced from the template's PNG ornaments. */

const BLOB_ARC = [[.983,.002],[.992,.002],[.997,.059],[.996,.208],[.979,.342],[.951,.455],[.908,.568],
  [.87,.641],[.812,.724],[.753,.789],[.68,.85],[.614,.892],[.537,.93],[.409,.971],[.266,.994],[.141,.998],
  [0,.989],[.003,.965],[.016,.786],[.187,.794],[.361,.774],[.441,.752],[.518,.718],[.633,.638],[.684,.582],
  [.724,.524],[.783,.389],[.817,.212],[.816,.032]];

const BLOB_CORNER = [[0,0],[.848,.002],[.921,.095],[.97,.195],[.995,.302],[.998,.37],[.991,.395],[.949,.45],
  [.876,.491],[.572,.591],[.537,.611],[.5,.657],[.498,.755],[.465,.834],[.402,.909],[.339,.955],[.238,.993],
  [.145,.995],[.063,.966],[0,.914]];

const BLOB_WAVE = [[.018,0],[.999,0],[.999,.785],[.999,.99],[.966,.906],[.902,.83],[.842,.795],[.749,.774],
  [.701,.747],[.677,.719],[.652,.667],[.597,.427],[.582,.399],[.549,.372],[.521,.372],[.487,.392],[.341,.566],
  [.263,.625],[.223,.639],[.17,.639],[.137,.628],[.101,.604],[.076,.576],[.042,.514],[.007,.372],[0,.201]];

/* ------------------------------------------------------------------ helpers */

/** Scale a normalised command list to a w x h box (inches). */
function geom (cmds, w, h) {
  return cmds.map(c => {
    if (c[0] === 'z') return { close: true };
    if (c[0] === 'm') return { x: c[1] * w, y: c[2] * h, moveTo: true };
    if (c[0] === 'l') return { x: c[1] * w, y: c[2] * h };
    return { x: c[5] * w, y: c[6] * h, curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } };
  });
}

/** Scale a normalised polygon to a w x h box, optionally mirrored. */
function poly (pts, w, h, flip) {
  const out = pts.map(([x, y], i) => ({ x: (flip ? 1 - x : x) * w, y: y * h, moveTo: i === 0 }));
  out.push({ close: true });
  return out;
}

/** Rectangle whose top-right corner is rounded by r inches. */
function topRightRounded (w, h, r) {
  return [
    { x: 0, y: 0, moveTo: true }, { x: w - r, y: 0 },
    { x: w, y: r, curve: { type: 'cubic', x1: w - r * 0.45, y1: 0, x2: w, y2: r * 0.45 } },
    { x: w, y: h }, { x: 0, y: h }, { close: true }
  ];
}

/** A text box that matches the deck's defaults (top-anchored, tight insets). */
function text (slide, body, o) {
  slide.addText(body, Object.assign({
    valign: 'top', margin: INSET, fontFace: LIGHT_FONT, fontSize: 9, color: GREY, align: 'left'
  }, o));
}

/** Section heading: 45 pt semibold blue. */
function heading (slide, body, x, y, w, h, align) {
  text(slide, body, { x, y, w, h, fontFace: BOLD_FONT, fontSize: 45, color: BLUE, align: align || 'left' });
}

/** Card title: 15 pt semibold blue. */
function cardTitle (slide, body, x, y, w, h, o) {
  text(slide, body, Object.assign({ x, y, w, h, fontFace: BOLD_FONT, fontSize: 15, color: BLUE }, o));
}

/** Body copy: 9 pt light grey, 150 % leading. */
function copy (slide, body, x, y, w, h, o) {
  text(slide, body, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.5 }, o));
}

/** Oversized translucent section numeral ("01." ... "05."). */
function numeral (slide, label, x, y) {
  text(slide, label, {
    x, y, w: 3.5321, h: 2.5875, fontFace: BOLD_FONT, fontSize: 149, color: IVORY, shadow: DROP()
  });
}

/** "6" + superscript "th" + " century" style runs. */
function ordinal (num, rest) {
  return [
    { text: num, options: {} },
    { text: 'th', options: { superscript: true } },
    { text: rest, options: {} }
  ];
}

/**
 * Placeholder standing in for a photograph in the reference deck:
 * a light grey block in the picture's exact frame, labelled "[image]".
 */
function imageBox (slide, o) {
  const opts = { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: PHOTO } };
  if (o.rectRadius !== undefined) opts.rectRadius = o.rectRadius;
  if (o.points) opts.points = o.points;
  if (o.shadow) opts.shadow = o.shadow;
  slide.addShape(o.shape || 'rect', opts);
  slide.addText('[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.13, w: o.w, h: 0.26, align: 'center', valign: 'middle',
    margin: 0, fontFace: LIGHT_FONT, fontSize: 9, color: PHOTO_TX
  });
}

/** Filled panel with rounded corners (radius given in inches). */
function panel (slide, x, y, w, h, r, color, o) {
  slide.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: r, fill: { color } }, o));
}

/* ---------------------------------------------------------- lorem fragments */

const L_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
  'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat.\u00a0';
const L_TAIL = 'Labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
  'laboris nisi ut aliquip ex ea commodo consequat.\u00a0\u00a0';
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';
const L_TEMPOR = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const L_DOLORE = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
  'labore et dolore.';
const L_SEDDO = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do.';
const L_TINY = 'Lorem ipsum dolor sit amet, ';
const L_ULLAMCO = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
  'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.';
const L_LABORE = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
  'labore.';
const L_ADIPISCING = 'Lorem ipsum dolor sit amet, consectetur adipiscing.';

/* ------------------------------------------------------------ slide 1 title */

function slide01 (s) {
  text(s, 'Perfume', { x: 4.5312, y: 2.8956, w: 5.194, h: 1.5273, fontFace: BOLD_FONT, fontSize: 86, color: BLUE });
  text(s, 'Presentation Template', { x: 4.5312, y: 4.2289, w: 5.194, h: 0.4796, fontSize: 24, color: BLUE });

  const outline = { color: IVORY, width: 0.75, transparency: 39 };
  s.addShape('rect', { x: 1.5281, y: 0.3021, w: 1.1615, h: 0.4577, fill: { color: SAND }, line: outline });
  s.addShape('custGeom', {
    x: 0.2747, y: 0.9272, w: 3.6912, h: 4.3958, fill: { color: SAND }, line: outline,
    points: geom(BOTTLE, 3.6912, 4.3958)
  });
}

/* ------------------------------------------------- slide 2 introduction 01. */

function slide02 (s) {
  imageBox(s, { x: 2.3459, y: 0.3021, w: 3.2565, h: 5.0208, shape: 'custGeom', shadow: SOFT(),
    points: geom(ARCH_TALL, 3.2565, 5.0208) });

  text(s, 'Introduction', { x: -1.8191, y: 2.396, w: 5.0208, h: 0.8331, rotate: -90,
    fontFace: BOLD_FONT, fontSize: 45, color: BLUE });
  copy(s, L_LONG, -1.0909, 2.4339, 5.0208, 0.7573, { rotate: -90 });

  s.addShape('custGeom', { x: 8.5576, y: 0, w: 1.4424, h: 1.3831, fill: { color: SAND },
    points: geom(BRACKET, 1.4424, 1.3831) });

  panel(s, 6.1501, 1.9853, 3.5752, 3.3376, 0.1942, SAND);
  text(s, '7000 BC', { x: 6.5, y: 2.2395, w: 3.2253, h: 0.6311, fontFace: BOLD_FONT, fontSize: 33,
    color: BLUE, fill: { color: SAND } });
  copy(s, [{ text: L_LONG.trim(), options: { breakLine: true } }, { text: '', options: { breakLine: true } },
    { text: L_TAIL }], 6.5, 2.9482, 3.0217, 2.0453, { fill: { color: SAND } });

  numeral(s, '01.', 5.1763, -0.6022);
}

/* ------------------------------------------------------ slide 3 timeline    */

const HISTORY_TOP = [
  { x: 0.2747, label: '7000 BC', w: 1.0989 },
  { x: 1.8142, label: '2000 BC', w: 1.0989 },
  { x: 3.3537, label: ordinal('6', ' century'), w: 1.2929 },
  { x: 4.8931, label: ordinal('7', ' century'), w: 1.3455 },
  { x: 6.4326, label: ordinal('11', ' century'), w: 1.2929 },
  { x: 7.972, label: '1370', w: 1.0989 }
];
const HISTORY_BOTTOM = [
  { x: 0.9887, label: ordinal('16', ' century'), w: 1.3906 },
  { x: 2.4593, label: ordinal('17', ' century'), w: 1.3 },
  { x: 3.9299, label: ordinal('18', ' century'), w: 1.3455 },
  { x: 5.4005, label: ordinal('19', ' century'), w: 1.3906 },
  { x: 6.8711, label: '1808', w: 1.0989 },
  { x: 8.3417, label: ordinal('19', ' century'), w: 1.3 }
];

function slide03 (s) {
  heading(s, 'The History of Perfume', 5.1667, 0.3021, 4.5586, 1.5904);
  s.addShape('custGeom', { x: 0, y: 0, w: 4.7611, h: 1.5904, fill: { color: SAND },
    points: geom(HILL, 4.7611, 1.5904) });
  s.addShape('custGeom', { x: -0.0012, y: 2.2968, w: 9.9994, h: 3.0261,
    line: { color: 'AB9969', width: 1.823 }, points: geom(TIMELINE, 9.9994, 3.0261) });

  [[HISTORY_TOP, 2.2009, 2.5603, 2.9006], [HISTORY_BOTTOM, 3.6868, 4.1347, 4.475]].forEach(row => {
    const [items, dotY, labelY, copyY] = row;
    items.forEach(it => {
      s.addShape('ellipse', { x: it.x, y: dotY, w: 0.1919, h: 0.1919, fill: { color: 'AB9969' } });
      text(s, it.label, { x: it.x, y: labelY, w: it.w, h: 0.3029, fontFace: BOLD_FONT, fontSize: 14, color: BLUE });
      copy(s, L_TINY, it.x, copyY, 1.0989, 0.5301);
    });
  });
}

/* ------------------------------------------- slide 4 importance of perfume  */

const IMPORTANCE = [
  { y: 2.0331, title: 'Enhancing personal style', icon: ICON_STYLE, iw: 0.319, ix: 5.5558 },
  { y: 3.1966, title: 'Evoking emotions', icon: ICON_EMOTION, iw: 0.3032, ix: 5.5677 },
  { y: 4.36, title: 'Cultural significance', icon: ICON_LEAF, iw: 0.319, ix: 5.5558 }
];

function slide04 (s) {
  s.addShape('round2SameRect', { x: 0.2747, y: 0, w: 4.319, h: 4.6146, rotate: 180, fill: { color: SAND },
    rectRadius: 2.1595 });
  heading(s, 'Importance of Perfume', 5.2704, 0.3021, 4.1218, 1.5904);

  IMPORTANCE.forEach(c => {
    panel(s, 5.2704, c.y, 4.4549, 0.9629, 0.135, SAND);
    cardTitle(s, c.title, 6.0524, c.y + 0.1713, 3.6729, 0.3282, { fill: { color: SAND } });
    copy(s, L_SHORT, 6.0524, c.y + 0.4887, 3.673, 0.2271, { fill: { color: SAND } });
    s.addShape('custGeom', { x: c.ix, y: c.y + 0.1713, w: c.iw, h: 0.319, fill: { color: BLUE },
      points: geom(c.icon, c.iw, 0.319) });
  });

  numeral(s, '02.', 2.0929, 3.6642);
  imageBox(s, { x: 0.7224, y: 0.5954, w: 3.4237, h: 3.4237, shape: 'ellipse', shadow: SOFT() });
}

/* --------------------------------------------- slide 5 fragrance families   */

const FAMILIES = [
  { x: 0.2747, badge: 0.4375, n: '1', name: 'Floral', archY: 2.0729, textY: 4.245 },
  { x: 2.2217, badge: 2.3443, n: '2', name: 'Woody', archY: 3.3125, textY: 2.1064 },
  { x: 4.1686, badge: 4.3417, n: '3', name: 'Oriental', archY: 2.0729, textY: 4.245 },
  { x: 6.1156, badge: 6.2759, n: '4', name: 'Fresh', archY: 3.3125, textY: 2.1064 },
  { x: 8.0625, badge: 8.2732, n: '5', name: 'Gourmand', archY: 2.0729, textY: 4.245 }
];

function slide05 (s) {
  FAMILIES.forEach(f => imageBox(s, { x: f.x, y: f.archY, w: 1.6628, h: 2.0104, shape: 'custGeom',
    shadow: SOFT(), points: geom(ARCH_SMALL, 1.6628, 2.0104) }));

  heading(s, 'Fragrance Families', 0.2747, 0.3021, 9.4505, 0.8331, 'center');
  copy(s, L_LONG, 0.9531, 1.2081, 8.0938, 0.5301, { align: 'center' });

  FAMILIES.forEach(f => {
    s.addShape('ellipse', { x: f.badge, y: f.archY + 0.0032, w: 0.3299, h: 0.3299, fill: { color: BLUE } });
    text(s, f.n, { x: f.badge, y: f.archY + 0.0167, w: 0.3299, h: 0.3029, align: 'center',
      fontFace: BOLD_FONT, fontSize: 14, color: IVORY });
    cardTitle(s, f.name, f.x, f.textY, 1.6628, 0.3282, { align: 'center' });
    copy(s, L_SHORT, f.x, f.textY + 0.3206, 1.6628, 0.7573, { align: 'center' });
  });

  [[0, false], [8.5521, true]].forEach(([x, flip]) => s.addShape('custGeom', {
    x, y: 0, w: 1.4479, h: 1.4885, fill: { color: DECOR }, shadow: LIFT(),
    points: poly(BLOB_CORNER, 1.4479, 1.4885, flip)
  }));
}

/* ----------------------------------------------- slide 6 anatomy of perfume */

const ANATOMY = [
  { y: 2.9299, name: 'Top' },
  { y: 3.8612, name: 'Middle' },
  { y: 4.7928, name: 'Base notes' }
];

function slide06 (s) {
  imageBox(s, { x: 0, y: 0, w: 3.9375, h: 5.625 });
  panel(s, 2.0938, 0.705, 6.6979, 1.9513, 0.2735, SAND);
  heading(s, 'Anatomy of a Perfume', 3.9375, 0.8854, 4.2253, 1.5904);

  ANATOMY.forEach(r => {
    cardTitle(s, r.name, 4.4002, r.y, 1.3724, 0.3282);
    copy(s, L_SHORT, 5.7725, r.y, 2.6084, 0.5301);
  });

  s.addShape('custGeom', { x: 8.5873, y: 4.2123, w: 1.4424, h: 1.3831, rotate: 90, fill: { color: SAND },
    points: geom(BRACKET, 1.4424, 1.3831) });

  [3.6427, 4.5941].forEach(y => s.addShape('line', { x: 4.3989, y, w: 3.9296, h: 0,
    line: { color: KHAKI, width: 1 } }));
}

/* --------------------------------------------- slide 7 how perfume is made  */

const STEPS = [
  { x: 0.7014, name: 'Extraction' },
  { x: 2.9409, name: 'Blending' },
  { x: 5.1805, name: 'Aging' },
  { x: 7.42, name: 'Bottling' }
];

function slide07 (s) {
  heading(s, 'How Perfume is Made', 0.2747, 0.3021, 4.4544, 1.5904);
  copy(s, L_LONG, 4.5924, 1.0797, 5.1328, 0.7573);
  panel(s, 0.2747, 2.0331, 9.4505, 3.2898, 0.2112, SAND);

  STEPS.forEach(st => {
    imageBox(s, { x: st.x, y: 2.315, w: 1.3888, h: 1.3888, shape: 'roundRect', rectRadius: 0.1794, shadow: SOFT() });
    text(s, st.name, { x: st.x, y: 3.9051, w: 1.8786, h: 0.3787, fontFace: BOLD_FONT, fontSize: 18, color: BLUE });
    copy(s, L_TEMPOR, st.x, 4.2837, 1.8786, 0.7573);
  });
}

/* --------------------------------------------------- slide 8 ingredients    */

const INGREDIENTS = [
  { x: 5.0, y: 0.3021, w: 1.5729, name: 'Flowers' },
  { x: 6.6337, y: 1.1472, w: 1.458, name: 'Spices' },
  { x: 8.1523, y: 2.2944, w: 1.5729, name: 'Woods' }
];

function slide08 (s) {
  s.addShape('round2SameRect', { x: 0.4957, y: -0.4957, w: 3.6293, h: 4.6207, rotate: 90,
    fill: { color: SAND }, rectRadius: 1.8147 });
  heading(s, 'Ingredients in Perfume', 1.7834, 3.7325, 4.4544, 1.5904);

  INGREDIENTS.forEach(it => {
    text(s, it.name, { x: it.x, y: it.y, w: it.w, h: 0.3787, fontFace: BOLD_FONT, fontSize: 18, color: BLUE });
    copy(s, L_TEMPOR, it.x, it.y + 0.3504, it.w, 0.9845);
  });

  copy(s, L_LONG, 6.2378, 4.0992, 3.4875, 1.2117);
  imageBox(s, { x: 1.1638, y: 0.3493, w: 2.9307, h: 2.9307, shape: 'ellipse', shadow: SOFT() });
  s.addShape('custGeom', { x: 7.4828, y: 0, w: 2.5172, h: 1.0788, fill: { color: DECOR }, shadow: LIFT(),
    points: poly(BLOB_WAVE, 2.5172, 1.0788) });
  numeral(s, '03.', 3.6133, 1.5997);
}

/* ------------------------------------------------ slide 9 science of scent  */

const SCENT_PANELS = [
  { y: 1.1842, w: 4.5417, h: 4.4408, r: 0.581 },
  { y: 2.4551, w: 3.4271, h: 3.1699, r: 0.508 },
  { y: 3.4688, w: 2.1387, h: 2.1562, r: 0.42 }
];
const SCENT_ROWS = [
  { y: 2.1403, h: 0.3282, title: 'How we smell' },
  { y: 3.3025, h: 0.3282, title: 'How scents affect us' },
  { y: 4.4646, h: 0.5806, title: 'How we can improve our sense of smell' }
];
const SCENT_LEADERS = [
  { x: 3.9483, y: 2.2977, w: 1.5101 },
  { x: 2.7639, y: 3.4688, w: 2.6944 },
  { x: 1.7526, y: 4.6372, w: 3.7057 }
];
const SCENT_STATS = [
  { x: 2.6735, y: 1.4087, v: '17%' },
  { x: 1.5725, y: 2.683, v: '38%' },
  { x: 0.2747, y: 3.7937, v: '45%' }
];

function slide09 (s) {
  SCENT_PANELS.forEach(p => s.addShape('custGeom', { x: 0, y: p.y, w: p.w, h: p.h, fill: { color: CREAM },
    shadow: LIFT(), points: topRightRounded(p.w, p.h, p.r) }));
  heading(s, 'The Science of Scent', 5.0, 0.3021, 4.3276, 1.5904);

  SCENT_ROWS.forEach(r => {
    cardTitle(s, r.title, 5.5895, r.y, 4.1358, r.h);
    copy(s, L_DOLORE, 5.5895, r.y + 0.3282, 4.1358, 0.5301);
  });
  SCENT_LEADERS.forEach(l => s.addShape('line', { x: l.x, y: l.y, w: l.w, h: 0,
    line: { color: KHAKI, width: 1.5, beginArrowType: 'oval' } }));
  SCENT_STATS.forEach(st => text(s, st.v, { x: st.x, y: st.y, w: 1.8639, h: 0.6311,
    fontFace: BOLD_FONT, fontSize: 33, color: BLUE }));
}

/* ------------------------------------------------------ slide 10 brands     */

const BRANDS = [
  { y: 2.3847, name: 'Chenil' },
  { y: 3.3632, name: 'Doir' },
  { y: 4.3416, name: 'Tom Ferd' }
];

function slide10 (s) {
  panel(s, 1.0405, 2.0345, 5.1836, 3.2884, 0.2282, SAND);
  heading(s, 'Famous Perfume Brands', 0.226, 0.3021, 5.1836, 1.5904);

  BRANDS.forEach(b => {
    cardTitle(s, b.name, 1.7272, b.y, 3.8104, 0.3282);
    copy(s, L_SEDDO, 1.7272, b.y + 0.3282, 3.8104, 0.3029);
    s.addShape('ellipse', { x: 1.577, y: b.y + 0.1101, w: 0.1042, h: 0.1042, fill: { color: 'AB9969' } });
  });

  imageBox(s, { x: 5.9149, y: 0.3021, w: 3.8104, h: 5.0208, shape: 'round2SameRect', rectRadius: 1.9052 });
}

/* ------------------------------------------------------ slide 11 trends     */

const TRENDS = [
  { x: 1.5086, name: 'Sustainability' },
  { x: 4.3152, name: 'Niche perfumes' },
  { x: 7.1218, name: 'Personalization' }
];

function slide11 (s) {
  heading(s, 'Perfume Trends', 0.226, 0.3021, 5.1836, 0.8331);
  TRENDS.forEach(t => {
    panel(s, t.x, 1.3026, 2.6034, 4.0203, 0.2482, SAND);
    cardTitle(s, t.name, t.x, 3.7265, 2.6034, 0.3282, { align: 'center', fill: { color: SAND } });
    copy(s, L_DOLORE, t.x + 0.3034, 4.138, 1.9968, 0.909, { align: 'center', fill: { color: SAND } });
  });
  TRENDS.forEach(t => imageBox(s, { x: t.x + 0.3034, y: 1.5799, w: 1.9967, h: 1.9791 }));
}

/* ------------------------------------------ slide 12 choosing the perfume   */

const CHOOSING = [
  'Consider your preferences', 'Consider the occasion', 'Consider your skin type',
  'Layer your fragrance', 'Choose a concentration'
];

/** "Learn more" / "Close" button pair used on slides 12 and 17. */
function buttons (s, y) {
  panel(s, 1.2333, y, 1.2336, 0.3488, 0.0581, CREAM);
  panel(s, 0.2747, y, 1.0961, 0.3488, 0.0581, BLUE);
  text(s, 'Learn more', { x: 0.2747, y: y + 0.0608, w: 1.0961, h: 0.2272, align: 'center', color: IVORY });
  text(s, 'Close', { x: 1.3708, y: y + 0.0608, w: 1.0961, h: 0.2272, align: 'center', color: BROWN });
}

function slide12 (s) {
  heading(s, 'Choosing the Right Perfume', 0.2747, 1.2221, 4.5757, 1.5904);
  panel(s, 5.0, 0.3021, 4.7253, 5.0208, 0.2867, SAND);
  copy(s, L_LONG, 0.2747, 2.8125, 4.5757, 0.9845);
  buttons(s, 4.9741);
  text(s, '5++', { x: 0.5819, y: 3.8608, w: 1.7357, h: 0.5806, fontFace: BOLD_FONT, fontSize: 30, color: BLUE });
  copy(s, 'Tips Choosing the Right Perfume', 0.5819, 4.3204, 2.1473, 0.3029);

  CHOOSING.forEach((title, i) => {
    const y = 0.6798 + i * 0.9086;
    cardTitle(s, title, 5.7972, y, 3.9281, 0.3282, { fill: { color: SAND } });
    copy(s, L_SHORT, 5.7972, y + 0.3281, 3.7445, 0.3029, { fill: { color: SAND } });
    s.addShape('ellipse', { x: 5.4535, y: y + 0.0572, w: 0.2137, h: 0.2137, fill: { color: BLUE } });
    s.addShape('custGeom', { x: 5.4953, y: y + 0.1161, w: 0.1301, h: 0.096, fill: { color: ICE },
      points: geom(CHECK, 0.1301, 0.096) });
  });
}

/* ------------------------------------------------ slide 13 application tips */

function slide13 (s) {
  imageBox(s, { x: 0.2747, y: 0.3021, w: 3.6211, h: 5.0208, shape: 'roundRect', rectRadius: 0.2386, shadow: SOFT() });
  panel(s, 5.5412, 2.3793, 4.1841, 2.9436, 0.2043, SAND);
  heading(s, 'Application Tips', 4.375, 0.3021, 5.3503, 0.8331);
  s.addShape('ellipse', { x: 3.5815, y: 3.8006, w: 1.4185, h: 1.4044, fill: { color: KHAKI } });
  numeral(s, '04.', 2.0091, 3.6958);
  copy(s, L_LONG, 4.375, 1.2922, 5.3503, 0.7573);

  ['01.', '02.', '03.'].forEach((num, i) => {
    const y = 2.6805 + i * 0.8551;
    text(s, num, { x: 5.8021, y, w: 0.4604, h: 0.3282, fontFace: BOLD_FONT, fontSize: 15, color: BLUE });
    cardTitle(s, 'Pulse points', 6.2241, y, 3.5011, 0.3282);
    copy(s, L_ADIPISCING, 6.2241, y + 0.3282, 3.5011, 0.3029);
  });
}

/* ------------------------------------------ slide 14 marketing & packaging  */

const PACKAGING = [
  { cardX: 2.5026, cardY: 2.5729, cardW: 7.2226, textX: 3.8346, imgX: 2.2373,
    name: 'Importance of bottle design', bold: false },
  { cardX: 2.3368, cardY: 4.0568, cardW: 7.3884, textX: 3.1576, imgX: 1.5089,
    name: 'Branding in attracting customers', bold: true }
];

function slide14 (s) {
  PACKAGING.slice().reverse().forEach(p => {
    panel(s, p.cardX, p.cardY, p.cardW, 1.2661, 0.1621, SAND);
    cardTitle(s, p.name, p.textX, p.cardY + 0.2039, 5.7469, 0.3282, { bold: p.bold, fill: { color: SAND } });
    copy(s, L_ULLAMCO, p.textX, p.cardY + 0.5321, 5.7468, 0.4545, { fill: { color: SAND } });
  });

  panel(s, 0.2747, 0.3021, 9.4505, 1.9949, 0.1281, SAND);
  heading(s, 'Marketing and Packaging', 0.492, 0.5043, 4.7816, 1.5904);
  s.addShape('custGeom', { x: 0, y: 4.2419, w: 1.4424, h: 1.3831, rotate: 180, fill: { color: CREAM },
    points: geom(BRACKET, 1.4424, 1.3831) });

  imageBox(s, { x: 5.2736, y: 0.3021, w: 4.4516, h: 1.9949, shape: 'roundRect', rectRadius: 0.1257 });
  PACKAGING.forEach(p => imageBox(s, { x: p.imgX, y: p.cardY, w: 1.2661, h: 1.2661, shape: 'roundRect',
    rectRadius: 0.1996, shadow: SOFT() }));
}

/* --------------------------------------------- slide 15 niche vs designer   */

const COMPARE = [
  { x: 0.2747, textX: 0.6851, card: SAND, pill: IVORY, body: GREY, name: 'Niche Parfumes', price: '$345' },
  { x: 5.1798, textX: 5.5902, card: KHAKI, pill: CREAM, body: WHITE, name: 'Designer Parfumes', price: '$134' }
];

function slide15 (s) {
  heading(s, 'Niche vs. Designer Perfumes', 0.2747, 0.3021, 9.4505, 0.8331, 'center');
  COMPARE.forEach(c => panel(s, c.x, 1.2978, 4.5455, 4.0252, 0.1653, c.card));

  COMPARE.forEach(c => {
    panel(s, c.textX, 1.677, 3.7247, 0.7247, 0.0927, c.pill, { shadow: DROP() });
    text(s, c.name, { x: c.textX, y: 1.85, w: 3.7247, h: 0.3787, align: 'center', fontFace: BOLD_FONT,
      fontSize: 18, color: BLUE });
    ['Production', 'Exclusivity'].forEach((label, i) => {
      cardTitle(s, label, c.textX, 2.5643 + i, 3.7247, 0.3282, { align: 'center' });
      copy(s, L_LABORE, c.textX, 2.8698 + i, 3.7247, 0.5301, { align: 'center', color: c.body });
    });
    text(s, c.price, { x: c.textX, y: 4.4922, w: 3.7247, h: 0.7573, align: 'center', fontFace: BOLD_FONT,
      fontSize: 41, color: BLUE });
  });
}

/* ---------------------------------------- slide 16 men / women / unisex     */

const AUDIENCE = ['For Men', 'For Women', 'Unisex'];

function slide16 (s) {
  AUDIENCE.forEach((name, i) => {
    const y = 0.3021 + i * 1.7429;
    panel(s, 5.1034, y, 4.6219, 1.535, 0.1615, SAND);
    cardTitle(s, name, 6.8713, y + 0.2247, 2.668, 0.3282, { fill: { color: SAND } });
    copy(s, L_LABORE, 6.8713, y + 0.5529, 2.6681, 0.6816, { fill: { color: SAND } });
  });

  s.addShape('round1Rect', { x: 0, y: 2.5787, w: 4.8287, h: 3.0463, fill: { color: KHAKI }, rectRadius: 0.732 });
  text(s, 'Perfume for Men, Women, and Unisex', { x: 0.2747, y: 2.9752, w: 4.5539, h: 2.3477,
    fontFace: BOLD_FONT, fontSize: 45, color: WHITE });

  s.addShape('custGeom', { x: -0.1426, y: -0.2604, w: 2.9214, h: 2.537, fill: { color: DECOR }, shadow: LIFT(),
    points: poly(BLOB_ARC, 2.9214, 2.537) });
  AUDIENCE.forEach((_, i) => imageBox(s, { x: 5.3006, y: 0.4291 + i * 1.7429, w: 1.3736, h: 1.2809 }));
}

/* --------------------------------------------- slide 17 perfume & emotions  */

function slide17 (s) {
  heading(s, 'Perfume and Emotions', 0.2747, 0.6813, 4.7253, 1.5904);
  panel(s, -0.4213, 2.3511, 10.1466, 2.9718, 0.3126, SAND);
  copy(s, [{ text: L_LONG, options: { breakLine: true } }, { text: '', options: { breakLine: true } },
    { text: 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea ' +
      'commodo consequat.\u00a0' }], 0.2747, 2.6509, 4.1685, 1.6661);
  buttons(s, 4.8125);
  imageBox(s, { x: 5.0, y: 0.3021, w: 4.1685, h: 4.6416, shape: 'roundRect', rectRadius: 0.2471 });
}

/* --------------------------------------------------- slide 18 the future    */

const FUTURE = [
  { x: 3.927, y: 0.3021, badgeX: 5.2735, badgeY: 0.2707, n: '01', name: 'AI in fragrace creation' },
  { x: 5.9497, y: 1.6631, badgeX: 7.2963, badgeY: 1.6302, n: '02', name: 'Biodegradle packaging' },
  { x: 7.9724, y: 3.0241, badgeX: 9.319, badgeY: 2.9897, n: '03', name: 'Sustainable practices' }
];

function slide18 (s) {
  heading(s, 'Future of Perfume Industry', 0.2747, 2.9752, 2.9444, 2.3477);

  FUTURE.forEach(c => {
    panel(s, c.x, c.y, 1.7528, 2.3018, 0.1489, SAND);
    cardTitle(s, c.name, c.x + 0.1381, c.y + 0.513, 1.6146, 0.5806, { fill: { color: SAND } });
    copy(s, L_TEMPOR, c.x + 0.1382, c.y + 1.1118, 1.6146, 0.909, { fill: { color: SAND } });
  });

  s.addShape('custGeom', { x: 2.9791, y: 4.0777, w: 4.4532, h: 1.5473, fill: { color: CREAM }, shadow: LIFT(),
    points: geom(WAVE_BLOB, 4.4532, 1.5473) });
  s.addShape('custGeom', { x: -0.1426, y: -0.2604, w: 2.9214, h: 2.537, fill: { color: DECOR }, shadow: LIFT(),
    points: poly(BLOB_ARC, 2.9214, 2.537) });
  s.addShape('custGeom', { x: 8.5576, y: 0, w: 1.4424, h: 1.3831, fill: { color: CREAM }, shadow: LIFT(),
    points: geom(BRACKET, 1.4424, 1.3831) });

  FUTURE.forEach(c => {
    s.addShape('ellipse', { x: c.badgeX, y: c.badgeY, w: 0.4921, h: 0.4921, fill: { color: BLUE } });
    text(s, c.n, { x: c.badgeX, y: c.badgeY + 0.0489, w: 0.4921, h: 0.3787, align: 'center',
      fontFace: BOLD_FONT, fontSize: 18, color: WHITE });
  });
}

/* ------------------------------------------------------ slide 19 fun facts  */

function slide19 (s) {
  imageBox(s, { x: 2.5105, y: 1.2553, w: 7.2147, h: 2.4773, shape: 'roundRect', rectRadius: 0.1561 });
  heading(s, 'Fun Facts About Perfume', 0.2747, 3.7325, 4.7253, 1.5904);
  copy(s, L_LONG, 5.1747, 4.214, 4.5505, 0.9845);
  panel(s, 0.2747, 0.3021, 3.9016, 1.9242, 0.1634, SAND);
  numeral(s, '05.', 0.4595, -0.0385);
}

/* ------------------------------------------------------ slide 20 thank you  */

function slide20 (s) {
  s.addShape('custGeom', { x: 0, y: 1.0012, w: 10, h: 4.6238, fill: { color: CREAM },
    shadow: LIFT(GREY), points: geom(WAVE_TOP, 10, 4.6238) });
  text(s, 'Thank You', { x: 0.3455, y: 2.8125, w: 9.3798, h: 1.5273, align: 'center', fontFace: BOLD_FONT,
    fontSize: 86, color: BLUE });
  text(s, 'For Your Attention', { x: 0.3455, y: 4.1458, w: 9.3798, h: 0.4796, align: 'center',
    fontSize: 24, color: BLUE });
}

/* ------------------------------------------------------------------- build  */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

function build () {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'PERFUME', width: 10, height: 5.625 });
  pptx.layout = 'PERFUME';
  pptx.title = 'Perfume';

  pptx.defineSlideMaster({ title: 'COVER', background: { color: CREAM } });
  pptx.defineSlideMaster({
    title: 'NUMBERED',
    background: { color: CREAM },
    slideNumber: { x: 9.5517, y: 5.3743, w: 0.4013, h: 0.1767, align: 'right',
      fontFace: LIGHT_FONT, fontSize: 6, color: BROWN }
  });

  BUILDERS.forEach((builder, i) => {
    builder(pptx.addSlide({ masterName: i === 0 ? 'COVER' : 'NUMBERED' }));
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0adeeea2-9d64-42d0-8bd9-9bce2afe6fb6_grok_final.pptx') });
}

build().then(f => console.log('wrote ' + f)).catch(err => { console.error(err); process.exit(1); });
