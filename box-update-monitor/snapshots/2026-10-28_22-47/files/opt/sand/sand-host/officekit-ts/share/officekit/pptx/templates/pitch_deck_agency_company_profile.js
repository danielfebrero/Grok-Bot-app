/**
 * CreateBiz — Pitch Deck Agency Presentation (24 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs from the reference deck.
 *
 * Note: the reference deck contains no raster media at all — every "Picture
 * Placeholder" in it is empty and renders as blank, so nothing is drawn for them.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const BLUE = '1460F8';        // theme accent1
const BLUE_LT = 'D0DFFE';     // accent1 lumMod 20% / lumOff 80%
const DARK = '262626';        // tx1 lumMod 85%
const GREY = '808080';        // tx1 lumMod 50%
const WHITE = 'FFFFFF';

const HEAD = 'Montserrat Medium';   // theme major latin font
const BODY = 'Open Sans';           // theme minor latin font

// pptxgenjs mutates the shadow object it is given, so hand out a fresh copy each time.
const cardShadow = () => ({ type: 'outer', blur: 10, offset: 0, angle: 0, color: '000000', opacity: 0.08 });

/* --------------------------------------------------- custom-geometry helpers */
const KAPPA = 0.5523;

/** Turn a normalised path (0..1 coords) into pptxgenjs custGeom points. */
function pts(nPath, w, h) {
    return nPath.map(s => {
        if (s[0] === 'M') return { x: s[1] * w, y: s[2] * h, moveTo: true };
        if (s[0] === 'L') return { x: s[1] * w, y: s[2] * h };
        if (s[0] === 'Z') return { close: true };
        return { x: s[5] * w, y: s[6] * h, curve: { type: 'cubic', x1: s[1] * w, y1: s[2] * h, x2: s[3] * w, y2: s[4] * h } };
    });
}

function freeform(slide, nPath, o) {
    slide.addShape('custGeom', Object.assign({ points: pts(nPath, o.w, o.h) }, o));
}

/** Circle of radius r centred in the unit box; ccw reverses the winding. */
function circlePath(r, ccw) {
    const c = 0.5, k = KAPPA * r, s = ccw ? -1 : 1;
    const q = [[0, -r], [r, 0], [0, r], [-r, 0]];   // top, right, bottom, left
    const t = [[k, 0], [0, k], [-k, 0], [0, -k]];   // clockwise tangents
    const order = ccw ? [0, 3, 2, 1] : [0, 1, 2, 3];
    const out = [['M', c + q[order[0]][0], c + q[order[0]][1]]];
    for (let i = 0; i < 4; i++) {
        const a = order[i], b = order[(i + 1) % 4];
        out.push(['C', c + q[a][0] + s * t[a][0], c + q[a][1] + s * t[a][1],
                       c + q[b][0] - s * t[b][0], c + q[b][1] - s * t[b][1],
                       c + q[b][0], c + q[b][1]]);
    }
    out.push(['Z']);
    return out;
}

/** Ring / donut: outer circle plus a counter-wound inner circle. */
function ringPath(innerR) {
    return circlePath(0.5, false).concat(circlePath(innerR, true));
}

/** Rectangle with per-corner elliptical radii, each given as [rx, ry] or 0. */
function roundedPath(tl, tr, br, bl) {
    const rx = c => (c ? c[0] : 0), ry = c => (c ? c[1] : 0);
    const p = [['M', rx(tl), 0], ['L', 1 - rx(tr), 0]];
    if (tr) p.push(['C', 1 - rx(tr) + KAPPA * rx(tr), 0, 1, ry(tr) - KAPPA * ry(tr), 1, ry(tr)]);
    p.push(['L', 1, 1 - ry(br)]);
    if (br) p.push(['C', 1, 1 - ry(br) + KAPPA * ry(br), 1 - rx(br) + KAPPA * rx(br), 1, 1 - rx(br), 1]);
    p.push(['L', rx(bl), 1]);
    if (bl) p.push(['C', rx(bl) - KAPPA * rx(bl), 1, 0, 1 - ry(bl) + KAPPA * ry(bl), 0, 1 - ry(bl)]);
    p.push(['L', 0, ry(tl)]);
    if (tl) p.push(['C', 0, ry(tl) - KAPPA * ry(tl), rx(tl) - KAPPA * rx(tl), 0, rx(tl), 0]);
    p.push(['Z']);
    return p;
}

/* -------------------------------------------------- decorative shape library */
// Big "C" blob that bleeds off a corner (slides 1, 3, 4, 13, 21, 23, 24).
const BLOB = [
    ['M', 0.01523, 0], ['L', 0.30144, 0], ['L', 0.28778, 0.04424],
    ['C', 0.28031, 0.08096, 0.27638, 0.11899, 0.27638, 0.15793],
    ['C', 0.27638, 0.46950, 0.52756, 0.72207, 0.83740, 0.72207],
    ['C', 0.87613, 0.72207, 0.91394, 0.71813, 0.95046, 0.71061],
    ['L', 1, 0.69780], ['L', 1, 0.98384], ['L', 0.92302, 0.99565],
    ['C', 0.89487, 0.99853, 0.86630, 1, 0.83740, 1],
    ['C', 0.37492, 1, 0, 0.62299, 0, 0.15793],
    ['C', 0, 0.12887, 0.00146, 0.10014, 0.00432, 0.07183], ['Z'],
];
// Tall hairline crescent on the title slide.
const CRESCENT_TALL = [
    ['M', 0.03524, 0], ['C', 0.56806, 0, 1, 0.13694, 1, 0.30587], ['L', 1, 0.69413],
    ['C', 1, 0.86306, 0.56806, 1, 0.03524, 1], ['L', 0, 0.99944], ['L', 0.06340, 0.99842],
    ['C', 0.54988, 0.98276, 0.92952, 0.85250, 0.92952, 0.69413], ['L', 0.92952, 0.30587],
    ['C', 0.92952, 0.14750, 0.54988, 0.01724, 0.06340, 0.00158], ['L', 0, 0.00056], ['Z'],
];
// Wide hairline crescent behind the "Break Slide" panel.
const CRESCENT_WIDE = [
    ['M', 0.92197, 0], ['L', 1, 0],
    ['C', 0.49081, 0, 0.07803, 0.22386, 0.07803, 0.5],
    ['C', 0.07803, 0.77614, 0.49081, 1, 1, 1], ['L', 0.92197, 1],
    ['C', 0.41278, 1, 0, 0.77614, 0, 0.5],
    ['C', 0, 0.22386, 0.41278, 0, 0.92197, 0], ['Z'],
];
// Quarter-ring sweep in the top-left of slide 9.
const SWOOSH = [
    ['M', 0.28143, 0], ['L', 0, 0], ['L', 0.00942, 0.08391],
    ['C', 0.08811, 0.60672, 0.42834, 1, 0.83614, 1],
    ['C', 0.86527, 1, 0.89405, 0.99799, 0.92242, 0.99408],
    ['L', 1, 0.97798], ['L', 1, 0.58826], ['L', 0.95008, 0.60571],
    ['C', 0.91328, 0.61595, 0.87517, 0.62133, 0.83614, 0.62133],
    ['C', 0.56294, 0.62133, 0.33500, 0.35785, 0.28228, 0.00760], ['Z'],
];
// Full-height blue panel with a rounded bite taken out of its left edge (slide 18).
const NOTCH_PANEL = [
    ['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['L', 0, 0.86876], ['L', 0.39098, 0.86876],
    ['C', 0.49895, 0.86876, 0.58647, 0.83651, 0.58647, 0.79672], ['L', 0.58647, 0.20328],
    ['C', 0.58647, 0.16349, 0.49895, 0.13124, 0.39098, 0.13124], ['L', 0, 0.13124], ['Z'],
];
const RING = ringPath(0.2921);                          // logo mark & header dots
const CARD_L = roundedPath([0.04241, 0.11146], 0, 0, [0.04241, 0.11146]);  // slide 2 panel
const CARD_L6 = roundedPath([0.07027, 0.11146], 0, 0, [0.07027, 0.11146]); // slide 6 panel
const CARD_L20 = roundedPath([0.08165, 0.11146], 0, 0, [0.08165, 0.11146]);// slide 20 panel
const CARD_PILL = roundedPath.apply(null, Array(4).fill([0.04344, 0.16667])); // slide 3 stat bar

/* ------------------------------------------------------------ text utilities */
// `fit:'resize'` mirrors the reference's <a:spAutoFit/> on every text box.
const TXT = { fontFace: BODY, color: GREY, fontSize: 11, valign: 'top', fit: 'resize' };

function text(slide, runs, o) {
    slide.addText(runs, Object.assign({}, TXT, o));
}

/**
 * Two-tone heading: dark first half, blue second half.
 * `o.nowrap` mirrors the reference boxes that carry wrap="none".
 */
function heading(slide, x, y, w, h, dark, blue, o) {
    o = o || {};
    text(slide, [
        { text: dark, options: { color: DARK } },
        { text: blue, options: { color: BLUE } },
    ], { x, y, w, h, fontSize: o.size || 28, fontFace: HEAD, align: o.align || 'left', wrap: !o.nowrap });
}

/** Grey 11pt paragraph at 150% line spacing. */
function para(slide, x, y, w, h, str, align) {
    text(slide, str, { x, y, w, h, lineSpacingMultiple: 1.5, align: align || 'left', wrap: true });
}

/** 14pt dark sub-heading (blue on a few slides). */
function label(slide, x, y, w, str, color, align) {
    text(slide, str, { x, y, w, h: 0.337, fontSize: 14, fontFace: HEAD, color: color || DARK, align: align || 'left' });
}

/** Big blue statistic ("10+", "11k", "98%", "01" …). */
function stat(slide, x, y, w, str, color, align, size) {
    text(slide, str, { x, y, w, h: size === 24 ? 0.505 : 0.438, fontSize: size || 20, fontFace: HEAD, color: color || BLUE, align: align || 'left' });
}

/** Blue pill button with centred white caption. */
function learnMore(slide, x, y) {
    slide.addShape('roundRect', { x, y, w: 1.863, h: 0.49, rectRadius: 0.49 / 6, fill: { color: BLUE } });
    text(slide, 'Learn More', { x: x + 0.205, y: y + 0.102, w: 1.454, h: 0.286, color: WHITE, align: 'center', charSpacing: 3 });
}

/** White card with the deck's soft drop shadow. */
function card(slide, x, y, w, h, radiusFrac) {
    slide.addShape('roundRect', {
        x, y, w, h, rectRadius: (radiusFrac || 1 / 6) * Math.min(w, h),
        fill: { color: WHITE }, shadow: cardShadow(),
    });
}

/* number + title + body, laid out left-to-right (used on slides 3, 4, 6, 19, 20, 24) */
function statBlock(slide, o) {
    stat(slide, o.x, o.y + 0.263, o.numW, o.num);
    label(slide, o.tx, o.y, o.titleW || 2.0, o.title);
    para(slide, o.tx, o.y + 0.337, o.bodyW || 2.214, 0.627, o.body);
}

/* blue circle + white number + title + body (slides 11-14, 22) */
function numberRow(slide, o) {
    slide.addShape('ellipse', { x: o.cx, y: o.cy, w: 0.78, h: 0.78, fill: { color: BLUE } });
    stat(slide, o.cx + (o.num.length > 2 ? 0.068 : 0.113), o.cy + 0.135, o.num.length > 2 ? 0.64 : 0.551, o.num, WHITE, 'center', 24);
    label(slide, o.tx, o.ty, 2.4, o.title);
    para(slide, o.tx, o.ty + 0.337, 2.667, 0.627, o.body);
}

/* ------------------------------------------------------------ page furniture */
/** Header logo, page number, footer bar and scroll indicator — on every slide. */
function chrome(slide, page, o) {
    o = o || {};
    const logoColor = o.logoOnDark ? WHITE : DARK;
    const arrowColor = o.logoOnDark ? WHITE : BLUE;
    const scrollColor = o.scrollOnDark ? WHITE : BLUE;

    // logo: ring + short arrow + wordmark
    freeform(slide, RING, { x: 0.331, y: 0.270, w: 0.254, h: 0.254, fill: { color: BLUE_LT } });
    slide.addShape('line', { x: 0.426, y: 0.395, w: 0.216, h: 0, line: { color: arrowColor, width: 1, beginArrowType: 'triangle' } });
    text(slide, [
        { text: 'Create', options: { color: logoColor } },
        { text: 'Biz', options: { color: o.logoOnDark ? WHITE : BLUE } },
    ], { x: 0.696, y: 0.268, w: 0.965, h: 0.286, fontFace: HEAD, wrap: false });

    // six-dot cluster, top right
    [[12.933, 0.270], [12.742, 0.270], [12.552, 0.270],
     [12.742, 0.460], [12.933, 0.460], [12.933, 0.650]].forEach(d => {
        freeform(slide, RING, { x: d[0], y: d[1], w: 0.088, h: 0.088, fill: { color: o.dotsOnDark ? WHITE : BLUE } });
    });

    // scroll indicator: two half-circles enclosing an up and a down arrow
    slide.addShape('arc', { x: 0.5155, y: 3.267, w: 0.254, h: 0.254, angleRange: [180, 360], line: { color: scrollColor, width: 1 } });
    slide.addShape('line', { x: 0.6425, y: 3.394, w: 0, h: 0.283, line: { color: scrollColor, width: 1, beginArrowType: 'triangle' } });
    slide.addShape('line', { x: 0.6425, y: 3.822, w: 0, h: 0.283, line: { color: scrollColor, width: 1, endArrowType: 'triangle' } });
    slide.addShape('arc', { x: 0.5155, y: 3.979, w: 0.254, h: 0.254, angleRange: [0, 180], line: { color: scrollColor, width: 1 } });

    // footer strip + page number
    slide.addShape('rect', { x: 0, y: 6.946, w: 2.92, h: 0.286, fill: { color: o.footerBlue ? BLUE : BLUE_LT } });
    text(slide, 'Your Website Here', { x: 0.331, y: 6.946, w: 2.258, h: 0.286, color: o.footerBlue ? WHITE : GREY, align: 'center', charSpacing: 3, wrap: false });
    text(slide, 'Page ' + String(page).padStart(2, '0'), { x: 11.999, y: 6.946, w: 1.0, h: 0.286, fontFace: HEAD, color: o.pageOnDark ? WHITE : DARK, align: 'right' });
}

/* ------------------------------------------------------------- shared bodies */
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing';
const LOREM_TINY = 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt utlabore et dolore magna aliqua. Ut enim ad minim veniam nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const LOREM_SIDE = 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enima minim veniam quis nostrud exercitation ullamco laboris nisi aliqui commodo consequat. Duis aute irure dolor reprehe in voluptate';
const LOREM_ROW = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ';
const LOREM_TOC = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';

/* ==================================================================== SLIDES */
const slides = [];

/* 1 — title */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.405, y: 0, w: 2.929, h: 2.912, fill: { color: BLUE_LT } });
    freeform(s, CRESCENT_TALL, { x: 10.405, y: 0.909, w: 1.802, h: 5.683, fill: { color: BLUE } });
    heading(s, 1.683, 2.378, 5.175, 1.313, 'Create', 'Biz', { size: 72, nowrap: true });
    text(s, 'Pitch Deck Agency Presentation', { x: 1.683, y: 3.522, w: 4.343, h: 0.337, fontSize: 14, charSpacing: 3, wrap: false });
    para(s, 1.683, 4.218, 5.984, 0.904, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut commodo consequat. ');
    chrome(s, 1);
});

/* 2 — table of content */
slides.push((s) => {
    freeform(s, CARD_L, { x: 4.739, y: 3.404, w: 8.595, h: 3.270, fill: { color: WHITE }, shadow: cardShadow() });
    heading(s, 1.248, 0.958, 3.598, 0.572, 'Table Of ', 'Content', { nowrap: true });
    para(s, 1.248, 1.712, 6.669, 0.904, LOREM_LONG.replace('quis nostrud', 'nostrud'));
    [['About Us', 5.883, 4.045], ['Our Team', 9.429, 4.045],
     ['Our Service', 5.883, 5.127], ['Our Portfolio', 9.429, 5.127]].forEach(t => {
        text(s, t[0], { x: t[1], y: t[2], w: 1.3, h: 0.286, fontFace: HEAD, color: BLUE, wrap: false });
        para(s, t[1], t[2] + 0.280, 3.123, 0.627, LOREM_TOC);
    });
    chrome(s, 2);
});

/* 3 — welcome message */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.413, y: 4.580, w: 2.929, h: 2.912, rotate: 90, fill: { color: BLUE } });
    freeform(s, CARD_PILL, { x: 1.176, y: 2.953, w: 7.592, h: 1.818, fill: { color: WHITE }, shadow: cardShadow() });
    heading(s, 4.676, 0.774, 3.982, 0.572, 'Welcome ', 'Message', { align: 'center', nowrap: true });
    para(s, 1.422, 1.491, 10.490, 0.904, LOREM_LONG + 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ', 'center');
    label(s, 9.034, 2.899, 1.124, 'About Us');
    para(s, 9.034, 3.236, 3.123, 0.627, LOREM_TOC);
    statBlock(s, { x: 1.709, y: 3.380, numW: 0.656, num: '10+', tx: 2.462, title: 'Year Of Existence', body: LOREM_SHORT });
    statBlock(s, { x: 5.159, y: 3.380, numW: 0.793, num: '98%', tx: 6.016, title: 'Happy Client', body: LOREM_SHORT });
    chrome(s, 3);
});

/* 4 — about us */
slides.push((s) => {
    freeform(s, BLOB, { x: 0, y: 0, w: 2.929, h: 2.912, flipH: true, fill: { color: BLUE } });
    card(s, 1.224, 1.078, 8.193, 2.672);
    heading(s, 1.941, 1.657, 2.465, 1.515, 'About Our Pitch Deck ', 'Agency');
    para(s, 4.901, 1.684, 3.941, 1.460, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi aliquip commodo consequat. ');
    statBlock(s, { x: 3.485, y: 4.566, numW: 0.656, num: '10+', tx: 4.237, title: 'Year Of Existence', body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do ', bodyW: 3.225 });
    statBlock(s, { x: 3.485, y: 5.721, numW: 0.582, num: '11k', tx: 4.237, title: 'Project Done', body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do ', bodyW: 3.225 });
    chrome(s, 4);
});

/* 5 — history */
slides.push((s) => {
    slide5Card(s);
    heading(s, 1.318, 5.376, 4.059, 1.043, 'History Of Our ', 'Pitch Deck Agency');
    para(s, 6.043, 5.307, 6.308, 1.182, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, seeiusmod tempor incididunt labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi aliquip ex ea commodo consequat. Duis aute irure reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.');
    chrome(s, 5);
});
function slide5Card(s) {
    s.addShape('rect', { x: 1.428, y: 1.469, w: 4.429, h: 2.363, fill: { color: WHITE }, shadow: cardShadow() });
    [['2014 - 2019', 2.040, 1.895, 1.266], ['2020 - 2025', 2.924, 2.779, 1.338]].forEach(r => {
        label(s, 1.833, r[1], r[3], r[0]);
        para(s, 3.286, r[2], 2.254, 0.627, LOREM_SHORT);
    });
}

/* 6 — the best agency in town */
slides.push((s) => {
    freeform(s, CARD_L6, { x: 0, y: 0.918, w: 8.961, h: 5.650, flipH: true, fill: { color: WHITE }, shadow: cardShadow() });
    heading(s, 1.257, 1.807, 4.407, 1.043, 'The Best Pitch Deck ', 'Agency In Town');
    para(s, 1.257, 3.152, 6.669, 1.182, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex commodo consequat. Duis aute irure dolor reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ');
    statBlock(s, { x: 1.257, y: 4.730, numW: 0.656, num: '10+', tx: 2.010, title: 'Year Of Existence', body: LOREM_SHORT, bodyW: 2.235 });
    statBlock(s, { x: 4.938, y: 4.730, numW: 0.582, num: '11k', tx: 5.691, title: 'Project Done', body: LOREM_SHORT, bodyW: 2.235 });
    chrome(s, 6);
});

/* 7 — vision and mission */
slides.push((s) => {
    s.addShape('rect', { x: 0, y: 0, w: 4.123, h: 2.761, fill: { color: BLUE } });
    heading(s, 8.218, 1.674, 3.410, 1.043, 'Our Vision and ', 'Mission');
    para(s, 8.218, 3.063, 3.933, 1.737, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi aliquip commodo consequat. Duis aute irure dolor reprehe in voluptate velit esse cillum dolore');
    learnMore(s, 8.325, 5.337);
    card(s, 1.309, 4.616, 5.628, 1.737);
    [['Our Vision', 2.254, 1.238, 1.755], ['Our Mission', 4.680, 1.384, 4.254]].forEach(c => {
        label(s, c[1], 5.003, c[2], c[0], DARK, 'center');
        para(s, c[3], 5.340, 2.235, 0.627, LOREM_SHORT, 'center');
    });
    chrome(s, 7, { logoOnDark: true });
});

/* 8 — the problem we face */
slides.push((s) => {
    s.addShape('rect', { x: 7.778, y: 3.525, w: 5.556, h: 2.363, fill: { color: WHITE }, shadow: cardShadow() });
    s.addShape('rect', { x: 0, y: 4.707, w: 4.682, h: 2.793, fill: { color: BLUE_LT } });
    heading(s, 1.418, 0.978, 3.004, 1.043, 'The Problem ', 'We Face');
    learnMore(s, 10.052, 1.255);
    para(s, 4.787, 0.909, 4.684, 1.182, LOREM_CARD);
    [['01', 4.088, 4.138, 3.993, 0.493], ['02', 4.900, 4.950, 4.805, 0.551]].forEach(r => {
        stat(s, 8.397, r[1], r[4], r[0]);
        label(s, 8.952, r[2], 1.496, 'Our Problem');
        para(s, 10.587, r[3], 2.254, 0.627, LOREM_SHORT);
    });
    chrome(s, 8, { footerBlue: true });
});

/* 9 — our excellent solution */
slides.push((s) => {
    freeform(s, SWOOSH, { x: 0, y: 0, w: 2.906, h: 2.138, flipH: true, fill: { color: BLUE_LT } });
    card(s, 3.792, 1.042, 10.251, 2.225);
    heading(s, 4.503, 1.633, 3.090, 1.043, 'Our Excellent ', 'Solution');
    para(s, 8.069, 1.563, 4.684, 1.182, LOREM_CARD);
    [['01', 3.750, 3.899, 0.493], ['02', 5.190, 5.342, 0.551]].forEach(r => {
        s.addShape('ellipse', { x: 5.462, y: r[1], w: 1.269, h: 1.269, fill: { color: BLUE } });
        stat(s, 7.305, r[2] + 0.263, r[3], r[0]);
        label(s, 8.010, r[2], 1.515, 'Our Solution');
        para(s, 8.010, r[2] + 0.337, 4.140, 0.627, LOREM_ROW);
    });
    chrome(s, 9);
});

/* 10 — case study */
slides.push((s) => {
    card(s, 4.935, 1.042, 9.748, 3.237);
    heading(s, 1.304, 5.659, 3.090, 0.572, 'Case ', 'Study');
    para(s, 4.697, 5.355, 7.419, 1.182, LOREM_LONG.replace('nisi ut aliquip ex ea', 'nisi aliquip ex') + 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ');
    [1.596, 2.761].forEach(y => {
        label(s, 5.628, y, 1.745, 'Our Case Study', BLUE);
        para(s, 5.628, y + 0.337, 4.140, 0.627, LOREM_ROW);
    });
    chrome(s, 10);
});

/* 11 — what we do? */
slides.push((s) => {
    heading(s, 5.167, 0.774, 3.000, 0.572, 'What We ', 'Do?', { align: 'center', nowrap: true });
    para(s, 1.422, 1.491, 10.490, 0.627, LOREM_LONG.trim(), 'center');
    [['01', 'Project Management', 3.098, 3.491, 3.400], ['02', 'Digital Marketing', 4.839, 5.229, 5.137]].forEach(r => {
        card(s, 1.571, r[2], 4.509, 1.567);
        numberRow(s, { num: r[0], cx: 2.013, cy: r[3], tx: 3.097, ty: r[4], title: r[1], body: LOREM_TINY });
    });
    chrome(s, 11);
});

/* 12 — our excellent service */
slides.push((s) => {
    heading(s, 1.233, 1.674, 3.032, 1.043, 'Our Excellent ', 'Service');
    para(s, 1.233, 3.063, 3.405, 1.737, LOREM_SIDE);
    learnMore(s, 1.339, 5.337);
    [['01', 'Project Management', 1.124, 1.540, 1.448], ['02', 'Digital Marketing', 2.946, 3.360, 3.268],
     ['03', 'Pitch Deck Agency', 4.769, 5.183, 5.091]].forEach(r => {
        card(s, 6.048, r[2], 3.672, 1.608);
        numberRow(s, { num: r[0], cx: 5.654, cy: r[3], tx: 6.737, ty: r[4], title: r[1], body: LOREM_TINY });
    });
    chrome(s, 12);
});

/* 13 — our services */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.413, y: 4.580, w: 2.929, h: 2.912, rotate: 90, fill: { color: BLUE_LT } });
    card(s, 2.749, 4.284, 9.340, 2.225);
    para(s, 6.825, 4.806, 4.684, 1.182, LOREM_CARD);
    heading(s, 3.393, 5.111, 3.032, 0.572, 'Our ', 'Services');
    [['01', 'Project Management', 1.167, 1.075], ['02', 'Digital Marketing', 2.415, 2.323]].forEach(r => {
        numberRow(s, { num: r[0], cx: 8.339, cy: r[2], tx: 9.423, ty: r[3], title: r[1], body: LOREM_TINY });
    });
    chrome(s, 13);
});

/* 14 — service we provide */
slides.push((s) => {
    s.addShape('rect', { x: 9.898, y: 0, w: 3.436, h: 2.740, fill: { color: BLUE_LT } });
    heading(s, 1.418, 5.407, 2.920, 1.043, 'Service We ', 'Provide');
    learnMore(s, 10.052, 5.684);
    para(s, 4.574, 5.338, 4.684, 1.182, LOREM_CARD);
    card(s, 8.007, 1.238, 3.782, 2.995, 0.14247);
    [['01', 'Project Management', 1.742, 1.650], ['02', 'Digital Marketing', 2.950, 2.858]].forEach(r => {
        numberRow(s, { num: r[0], cx: 7.494, cy: r[2], tx: 8.720, ty: r[3], title: r[1], body: LOREM_TINY });
    });
    chrome(s, 14);
});

/* 15 — break slide */
slides.push((s) => {
    s.addShape('rect', { x: -0.001, y: 1.554, w: 9.763, h: 4.393, fill: { color: WHITE }, shadow: cardShadow() });
    freeform(s, CRESCENT_WIDE, { x: 7.779, y: 1.326, w: 2.626, h: 4.842, fill: { color: BLUE } });
    heading(s, 1.233, 2.235, 4.138, 0.909, 'Break ', 'Slide', { size: 48 });
    para(s, 1.233, 3.273, 5.720, 0.904, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enima minim veniam nostrud exercitation ullamco laboris nisi commodo consequat. ');
    learnMore(s, 1.339, 4.553);
    chrome(s, 15);
});

/* 16 — our best team */
slides.push((s) => {
    card(s, 1.324, 3.232, 10.685, 3.198, 0.14247);
    [2.191, 5.241, 8.292].forEach(x => {
        s.addShape('ellipse', { x, y: 1.800, w: 2.850, h: 2.850, fill: { color: BLUE } });
    });
    heading(s, 5.073, 0.774, 3.187, 0.572, 'Our Best ', 'Team', { align: 'center', nowrap: true });
    [['Steven Coleman', 2.695, 1.843, 2.283], ['Adam Stewart', 5.849, 1.636, 5.333],
     ['Jerry Gravelle', 8.938, 1.557, 7.665 + 0.718]].forEach(p => {
        label(s, p[1], 5.133, p[2], p[0], DARK, 'center');
        para(s, p[3], 5.469, 2.667, 0.627, LOREM_TINY, 'center');
    });
    chrome(s, 16);
});

/* 17 — meet our great team */
slides.push((s) => {
    s.addShape('rect', { x: 0, y: 0, w: 2.920, h: 7.5, fill: { color: BLUE } });
    heading(s, 8.793, 1.674, 2.877, 1.043, 'Meet Our ', 'Great Team');
    para(s, 8.793, 3.063, 3.405, 1.737, LOREM_SIDE);
    learnMore(s, 8.899, 5.337);
    chrome(s, 17, { logoOnDark: true, scrollOnDark: true });
});

/* 18 — team member profile */
slides.push((s) => {
    freeform(s, NOTCH_PANEL, { x: 10.569, y: 0, w: 2.764, h: 7.5, fill: { color: BLUE } });
    heading(s, 1.179, 1.180, 3.482, 0.572, 'Steven ', 'Coleman', { nowrap: true });
    para(s, 1.178, 1.977, 6.669, 1.182, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim minim veniam, nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ');
    card(s, 1.324, 4.233, 8.104, 1.801, 0.14247);
    [['Project Manager', 3.410, 1.864], ['Digital Marketing', 6.305, 1.945]].forEach(r => {
        label(s, r[1], 4.633, r[2], r[0], BLUE);
        para(s, r[1], 4.960, 2.733, 0.627, LOREM_TINY);
    });
    chrome(s, 18, { pageOnDark: true, dotsOnDark: true });
});

/* 19 — our best work */
slides.push((s) => {
    s.addShape('rect', { x: 1.634, y: 1.744, w: 11.699, h: 3.958, fill: { color: WHITE }, shadow: cardShadow() });
    heading(s, 5.086, 0.774, 3.161, 0.572, 'Our Best ', 'Work', { align: 'center', nowrap: true });
    para(s, 1.422, 6.100, 10.490, 0.627, LOREM_LONG.trim(), 'center');
    [['10+', 'Year Of Existence', 2.115, 0.656], ['11k', 'Project Done', 3.242, 0.582],
     ['98%', 'Happy Client', 4.364, 0.793]].forEach(r => {
        statBlock(s, { x: 8.218, y: r[2], numW: r[3], num: r[0], tx: 9.072, title: r[1], body: LOREM_ROW, bodyW: 3.860 });
    });
    chrome(s, 19);
});

/* 20 — our excellent portfolio */
slides.push((s) => {
    freeform(s, CARD_L20, { x: 5.621, y: 0.918, w: 7.712, h: 5.650, fill: { color: WHITE }, shadow: cardShadow() });
    heading(s, 8.121, 1.669, 3.198, 1.043, 'Our Excellent ', 'Portfolio');
    para(s, 8.121, 3.032, 4.169, 1.460, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi utaliquip commodo aute irure dolor in reprehenderit');
    statBlock(s, { x: 8.121, y: 4.867, numW: 0.656, num: '10+', tx: 8.975, title: 'Project Done', body: 'Lorem ipsum dolor sit consectet adipiscing elit, sed do eiusmod tempor', bodyW: 3.315 });
    chrome(s, 20);
});

/* 21 — our portfolio */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.413, y: 4.580, w: 2.929, h: 2.912, rotate: 90, fill: { color: BLUE_LT } });
    heading(s, 1.321, 1.051, 3.198, 0.572, 'Our ', 'Portfolio');
    para(s, 1.321, 1.798, 6.669, 0.904, LOREM_LONG);
    s.addShape('rect', { x: 9.013, y: 1.197, w: 2.861, h: 3.682, fill: { color: WHITE }, shadow: cardShadow() });
    stat(s, 10.150, 3.309, 0.582, '11k', BLUE, 'center');
    label(s, 9.681, 3.742, 1.513, 'Project Done', DARK, 'center');
    para(s, 9.205, 4.078, 2.466, 0.627, 'Lorem ipsum dolor consectetur adipiscing elit, sed do', 'center');
    chrome(s, 21);
});

/* 22 — pricing plan */
slides.push((s) => {
    heading(s, 1.660, 1.674, 2.877, 1.043, 'Our Great ', 'Pricing Plan');
    para(s, 1.660, 3.063, 3.405, 1.737, LOREM_SIDE);
    learnMore(s, 1.766, 5.337);
    [['$15', 'Package One', 1.124, 1.540, 1.448], ['$25', 'Package Two', 2.946, 3.360, 3.268],
     ['$35', 'Package Three', 4.769, 5.183, 5.091]].forEach(r => {
        card(s, 6.351, r[2], 3.672, 1.608);
        numberRow(s, { num: r[0], cx: 5.956, cy: r[3], tx: 7.040, ty: r[4], title: r[1], body: LOREM_TINY });
    });
    chrome(s, 22);
});

/* 23 — contact us */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.405, y: -0.003, w: 2.929, h: 2.916, fill: { color: BLUE_LT } });
    heading(s, 1.418, 5.642, 3.198, 0.572, 'Contact ', 'Us');
    s.addShape('rect', { x: 7.293, y: 1.409, w: 4.878, h: 2.363, fill: { color: WHITE }, shadow: cardShadow() });
    [['Phone Number', 1.980, 1.835, 1.743], ['Our Address', 2.864, 2.719, 1.441]].forEach(r => {
        label(s, 7.789, r[1], r[3], r[0]);
        para(s, 9.548, r[2], 2.254, 0.627, LOREM_SHORT);
    });
    learnMore(s, 10.052, 5.684);
    para(s, 4.574, 5.338, 4.684, 1.182, LOREM_CARD);
    chrome(s, 23);
});

/* 24 — thank you */
slides.push((s) => {
    freeform(s, BLOB, { x: 10.413, y: 4.580, w: 2.929, h: 2.912, rotate: 90, fill: { color: BLUE_LT } });
    heading(s, 1.224, 1.160, 5.645, 1.313, 'Thank ', 'You', { size: 72, nowrap: true });
    text(s, 'For Watching This Presentation', { x: 1.224, y: 2.304, w: 4.313, h: 0.337, fontSize: 14, charSpacing: 3, wrap: false });
    para(s, 1.224, 2.830, 5.984, 0.904, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut commodo consequat. ');
    card(s, 3.245, 4.828, 4.661, 1.593);
    statBlock(s, { x: 3.603, y: 5.143, numW: 0.793, num: '98%', tx: 4.458, title: 'Happy Client', body: LOREM_TOC, bodyW: 3.132 });
    chrome(s, 24);
});

/* ====================================================================== build */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

slides.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '10f5b55c-9c37-436d-b0b6-96631571e00f_grok_final.pptx') })
    .then(f => console.log('wrote ' + f));
