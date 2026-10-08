/**
 * Recreation of "Business Summary" presentation template (30 slides, 16:9 / 13.333" x 7.5")
 * with pptxgenjs only.  Raster artwork from the source deck is replaced by programmatic
 * placeholders (light grey boxes / simple vector marks).
 *
 * Run:  node 106af8d9-b440-46f9-a86e-e8f94f0d0bbe_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BLUE = '00B0F0';
const NAVY = '1E2350';
const GREEN = '59EB7F';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const INK = '262626';   // tx1 lum 85/15
const GREY40 = '404040';
const GREY50 = '7F7F7F';
const GREY51 = '808080';
const GREY35 = '595959';
const GREY85 = 'D9D9D9';
const GREY91 = 'E8E8E8';
const SAND = 'E5DED0';
const LEMON = 'F7F9A1';
const LIME = 'CBFF49';
const PH_FILL = 'AAB1BB';   // stand-in for the deck's embedded photographs
const PH_TEXT = 'FFFFFF';
const PH_SOFT = 'E4E4E4';   // lighter stand-in where the photo sits on white

/* ------------------------------------------------------------------- fonts */
const SEMI = 'Poppins SemiBold';
const LIGHT = 'Poppins Light';
const MED = 'Poppins Medium';
const REG = 'Poppins';

/* margins (points) in pptxgenjs order [left, right, bottom, top] */
const M0 = [0, 0, 0, 0];
const M_GS = [7.2, 7.2, 7.2, 7.2];   // the "Google Shape" boxes use 0.1" all round
const M_SM = [2, 2, 2, 2];

/* ------------------------------------------------------------------ helpers */

/** Text box. `p` is [x, y, w, h]; `t` is a string or a pptxgenjs run array. */
function txt(slide, p, t, o = {}) {
    slide.addText(t, Object.assign({
        x: p[0], y: p[1], w: p[2], h: p[3],
        fontFace: LIGHT, fontSize: 18, color: BLACK,
        valign: 'top', align: 'left', isTextBox: true,
    }, o));
}

/** Filled / stroked shape. */
function shp(slide, kind, p, o = {}) {
    slide.addShape(kind, Object.assign({ x: p[0], y: p[1], w: p[2], h: p[3] }, o));
}

function rect(slide, p, color, o = {}) {
    shp(slide, 'rect', p, Object.assign({ fill: { color } }, o));
}

function ell(slide, p, color, o = {}) {
    shp(slide, 'ellipse', p, Object.assign(color ? { fill: { color } } : {}, o));
}

/** Rounded rectangle whose corner radius is given as a fraction of min(w,h). */
function rrect(slide, p, color, adj = 0.16667, o = {}) {
    shp(slide, 'roundRect', p, Object.assign({
        fill: color ? { color } : undefined,
        rectRadius: adj * Math.min(p[2], p[3]),
    }, o));
}

/** Pill (fully rounded) rectangle. */
function pill(slide, p, color, o = {}) {
    rrect(slide, p, color, 0.5, o);
}

/** Horizontal / vertical rule. */
function line(slide, p, color, width = 3, o = {}) {
    shp(slide, 'line', p, Object.assign({ line: Object.assign({ color, width }, o.line) }, o));
}

/** Drop shadow descriptor (blur & offset in points). */
function sh(blur, offset, angle, opacity) {
    return { type: 'outer', blur, offset, angle, color: BLACK, opacity };
}
const SH_CARD = sh(17, 4, 90, 0.10);     // soft card shadow used on slides 26/27/28
const SH_PANEL = sh(33.6, 16.7, 63, 0.042); // chart cards, slides 14/15
const SH_SOFT = sh(22, 9, 90, 0.09);     // slide 8 white cards

/**
 * Free-form path from a compact command list.  Coordinates are fractions of
 * the shape box: ['M',x,y] ['L',x,y] ['C',x1,y1,x2,y2,x,y] ['Z'].
 */
function poly(slide, p, spec, o = {}) {
    const [, , w, h] = p;
    const pts = spec.map((c) => {
        if (c[0] === 'Z') return { close: true };
        if (c[0] === 'C') {
            return {
                x: c[5] * w, y: c[6] * h,
                curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h },
            };
        }
        return { x: c[1] * w, y: c[2] * h, moveTo: c[0] === 'M' };
    });
    shp(slide, 'custGeom', p, Object.assign({ points: pts }, o));
}

/* the "capsule with a bump" silhouette used on slides 13 and 27 */
const CAPSULE = [
    ['M', 0.2584, 0],
    ['C', 0.3209, 0, 0.3782, 0.0428, 0.4228, 0.1142],
    ['L', 0.4399, 0.1442], ['L', 0.4425, 0.1439], ['L', 0.8160, 0.1439],
    ['C', 0.9176, 0.1439, 1, 0.3034, 1, 0.5],
    ['C', 1, 0.6966, 0.9176, 0.8561, 0.8160, 0.8561],
    ['L', 0.4425, 0.8561], ['L', 0.4399, 0.8558], ['L', 0.4228, 0.8858],
    ['C', 0.3782, 0.9572, 0.3209, 1, 0.2584, 1],
    ['C', 0.1157, 1, 0, 0.7761, 0, 0.5],
    ['C', 0, 0.2239, 0.1157, 0, 0.2584, 0],
    ['Z'],
];

/**
 * Rectangle whose top-left and bottom-right corners are rounded off with a
 * large radius (slide 18's grey slab).  rx / ry are fractions of w / h.
 */
function round2Diag(rx, ry) {
    const k = 0.4477;  // 1 - bezier circle constant
    return [
        ['M', rx, 0], ['L', 1, 0], ['L', 1, 1 - ry],
        ['C', 1, 1 - ry * k, 1 - rx * k, 1, 1 - rx, 1],
        ['L', 0, 1], ['L', 0, ry],
        ['C', 0, ry * k, rx * k, 0, rx, 0],
        ['Z'],
    ];
}

/**
 * Stand-in for a raster photograph embedded in the source deck: a flat block
 * of colour at the picture's position, tagged so it reads as artwork.
 */
function photoBox(slide, p, o = {}) {
    shp(slide, o.shape || 'rect', p, Object.assign({
        fill: { color: o.color || PH_FILL }, rectRadius: o.radius,
    }, o.rotate ? { rotate: o.rotate } : {}, o.flipH ? { flipH: true } : {}));
    txt(slide, [p[0], p[1] + p[3] / 2 - 0.18, p[2], 0.36], '[image]', {
        fontSize: 12, color: o.label || PH_TEXT, align: 'center', valign: 'middle',
    });
}

/**
 * The template's own "drop a picture here" frames.  They hold no artwork in
 * the source file and paint nothing, so nothing is drawn for them either -
 * the call documents where a picture would go.
 */
function photoSlot() { /* intentionally empty - unfilled template placeholder */ }

/** Stand-in for a small line-art icon. */
function icon(slide, p, color) {
    const [x, y, w, h] = p;
    shp(slide, 'roundRect', p, { line: { color, width: 1.25 }, rectRadius: Math.min(w, h) * 0.22 });
    ell(slide, [x + w * 0.31, y + h * 0.31, w * 0.38, h * 0.38], color);
}

/** Check mark drawn as a filled tick. */
function tick(slide, p, color) {
    poly(slide, p, [
        ['M', 0.02, 0.48], ['L', 0.16, 0.33], ['L', 0.38, 0.55],
        ['L', 0.85, 0.05], ['L', 1, 0.20], ['L', 0.38, 0.86], ['Z'],
    ], { fill: { color } });
}

/** Tick inside a thin circle. */
function tickCircle(slide, p, color) {
    const [x, y, w, h] = p;
    shp(slide, 'ellipse', p, { line: { color, width: 1 } });
    tick(slide, [x + w * 0.22, y + h * 0.28, w * 0.56, h * 0.44], color);
}

/** "AGREE / GO BACK" pill pair used on slides 6 and 7. */
function agreeGoBack(slide, x, y, agreeColor) {
    pill(slide, [x + 1.043, y, 0.965, 0.28], null, { line: { color: BLACK, width: 1 } });
    pill(slide, [x + 0.057, y + 0.009, 0.831, 0.28], BLUE, { line: { color: BLACK, width: 1 } });
    txt(slide, [x, y + 0.012, 0.965, 0.269], 'AGREE',
        { fontSize: 10, align: 'center', color: agreeColor });
    txt(slide, [x + 1.021, y + 0.015, 0.965, 0.269], 'GO BACK',
        { fontSize: 10, align: 'center' });
}

const CHART_BASE = {
    showLegend: false, showTitle: false, showValue: false,
    chartArea: { fill: { color: WHITE, transparency: 100 }, border: { pt: 0, color: WHITE } },
};

/* Doughnut-style KPI used on slide 14: a 2-slice pie masked by a white disc. */
function ringChart(slide, chartPos, discPos) {
    slide.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }],
        Object.assign({
            x: chartPos[0], y: chartPos[1], w: chartPos[2], h: chartPos[3],
            chartColors: [BLUE, 'F2F2F2'],
            dataBorder: { pt: 1.5, color: WHITE }, firstSliceAng: 0,
        }, CHART_BASE));
    ell(slide, discPos, WHITE);
}

/* ================================================================= slides == */

function slide01(s) {                                   // title
    photoSlot(s, [0, 0, 13.333, 7.5]);
    rect(s, [0, 0, 13.333, 7.5], BLACK, { fill: { color: BLACK, transparency: 23 } });
    txt(s, [0.511, 2.079, 6.045, 2.524], [
        { text: 'Business', options: { color: BLUE, breakLine: true } },
        { text: 'Summary', options: { color: WHITE } },
    ], { fontFace: SEMI, fontSize: 72 });
    txt(s, [0.511, 4.772, 4.491, 0.438], 'Presentation Template', { fontSize: 20, color: WHITE });
    txt(s, [0.511, 5.343, 5.956, 0.633],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor '
        + 'incididunt ut labore et dolore magna aliqua. Ut enim ad minim.',
        { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
}

function slide02(s) {                                   // vision + numbered tiles
    txt(s, [0.69, 1.379, 5.381, 1.447], 'Vision of The Brand Guideline',
        { fontFace: SEMI, fontSize: 40 });
    txt(s, [0.69, 2.858, 5.953, 0.586],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vestibulum, '
        + 'leo eu bibendum convallis, nisi ex faucibus velit, ',
        { fontSize: 12, color: INK, lineSpacingMultiple: 1.2 });

    ['Market Understanding', 'Brand Purpose', 'Brand Value'].forEach((label, i) => {
        const y = 4.056 + i * 0.772;
        txt(s, [1.358, y + 0.003, 4.301, 0.37], label, { fontSize: 16, color: INK });
        ell(s, [0.809, y, 0.377, 0.377], BLUE, { shadow: sh(40, 28, 90, 0.12) });
        poly(s, [0.942, y + 0.096, 0.11, 0.187], [
            ['M', 0.05, 0], ['L', 1, 0.5], ['L', 0.05, 1], ['L', 0, 0.87],
            ['L', 0.66, 0.5], ['L', 0, 0.13], ['Z'],
        ], { fill: { color: WHITE } });
    });

    /* three photo tiles, each tagged with a number chip */
    photoSlot(s, [7.461, 0.812, 5.872, 2.532]);
    photoSlot(s, [7.463, 3.759, 2.771, 3.068]);
    photoSlot(s, [10.562, 3.758, 2.771, 3.068]);
    rect(s, [7.461, 0.812, 0.888, 0.727], NAVY);
    rect(s, [7.461, 3.753, 0.888, 0.727], NAVY);
    rect(s, [10.562, 3.758, 0.888, 0.727], GREEN);
    txt(s, [7.493, 0.974, 0.824, 0.404], '01', { color: WHITE, align: 'center' });
    txt(s, [7.493, 3.915, 0.824, 0.404], '03', { color: WHITE, align: 'center' });
    txt(s, [10.594, 3.92, 0.824, 0.404], '02', { color: WHITE, align: 'center' });
}

function slide03(s) {                                   // blue panel with 3 missions
    rect(s, [0, 0, 4.906, 7.5], BLUE);
    photoSlot(s, [4.906, 0, 3.887, 7.5]);

    const missions = [
        ['Training and development ', 1.098, 0.605],
        ['Quality and Realibilityy', 3.091, 0.562],
        ['Appreciate Employees', 5.04, 0.652],
    ];
    missions.forEach(([title, y, iconSize]) => {
        icon(s, [0.83, y, iconSize, iconSize], WHITE);
        txt(s, [0.821, y + 0.777, 3.427, 0.231], title,
            { fontSize: 16, color: WHITE, valign: 'middle', margin: M0 });
        txt(s, [0.821, y + 1.05, 3.207, 0.505],
            'Lorem ipsum dolor sit amet, consectetur adipiscing', { fontSize: 12, color: WHITE });
    });

    txt(s, [9.147, 0.955, 4.312, 2.794], 'Several Mission of Brand Guideline',
        { fontFace: SEMI, fontSize: 40 });
    txt(s, [9.147, 3.894, 3.729, 0.828],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vestibulum, leo eu bibendum',
        { fontSize: 12, color: INK, lineSpacingMultiple: 1.2 });
    txt(s, [9.147, 5.148, 1.69, 0.559], '+ 4193 ',
        { fontFace: SEMI, fontSize: 24, color: INK, lineSpacingMultiple: 1.2 });
    txt(s, [9.147, 5.713, 3.14, 0.586],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit',
        { fontSize: 12, color: INK, lineSpacingMultiple: 1.2 });
}

function slide04(s) {                                   // typeface specimen
    rrect(s, [1.324, 1.785, 3.101, 4.505], BLUE, 0.11964);
    pill(s, [-0.096, 2.377, 0.193, 2.548], BLUE);
    txt(s, [1.656, 0.906, 2.527, 0.64], 'Typeface', { fontFace: SEMI, fontSize: 32, align: 'center' });
    [2.265, 3.517, 4.769].forEach((y, i) => {
        txt(s, [1.613, y, 2.57, 1.01], 'Jost',
            { fontFace: SEMI, fontSize: 54, color: WHITE, align: 'center', italic: i === 2 });
    });
    txt(s, [5.979, 1.327, 5.699, 2.423], 'AaBb', { fontFace: SEMI, fontSize: 138 });
    txt(s, [6.097, 4.237, 5.58, 1.632],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis '
        + 'rhoncus mattis. Orci varius natoque penatibus et magnis dis parturient montes, nascetur '
        + 'ridiculus mus. Pellentesque habitant morbi tristique senectus et netus et malesuada.',
        { fontSize: 14, lineSpacingMultiple: 1.3 });
}

function slide05(s) {                                   // innovation / icon mosaic
    txt(s, [6.721, 1.434, 5.432, 2.121], 'Innovation: Driving Digital Transformation',
        { fontFace: SEMI, fontSize: 40 });
    txt(s, [6.721, 3.742, 4.763, 1.475],
        'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. '
        + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero',
        { fontSize: 14, lineSpacingMultiple: 1.5 });

    rrect(s, [6.859, 5.632, 1.359, 0.433], BLUE, 0.16667);
    rrect(s, [8.465, 5.632, 1.481, 0.433], null, 0.16667, { line: { color: BLUE, width: 1.5 } });
    txt(s, [6.859, 5.59, 1.359, 0.425], 'See more',
        { fontSize: 14, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });
    txt(s, [8.465, 5.59, 1.481, 0.415], 'Learn More',
        { fontSize: 14, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });

    /* icon tiles */
    [[0.952, 2.964], [4.7, 2.964], [4.7, 4.177], [3.458, 5.39]].forEach(([x, y]) => {
        rrect(s, [x, y, 1.103, 1.115], BLUE, 0.16667);
    });
    rrect(s, [0.908, 4.177, 3.61, 1.115], GREY91, 0.16667, { shadow: sh(30, 0, 0, 0.10) });
    txt(s, [1.1, 4.342, 3.227, 0.415], 'Title here',
        { fontSize: 14, align: 'center', lineSpacingMultiple: 1.5 });
    txt(s, [0.864, 4.712, 3.697, 0.415], 'Lorem ipsum dolor sit amet',
        { fontSize: 14, align: 'center', lineSpacingMultiple: 1.5 });
    [[5.016, 3.285], [5.016, 4.498], [3.773, 5.711], [1.267, 3.285]].forEach(([x, y]) => {
        icon(s, [x, y, 0.472, 0.472], WHITE);
    });
    photoSlot(s, [2.186, 0.995, 2.367, 3.084]);
}

/* slides 6 & 7 share a "sticker headline + dashed timeline" language */
function stickerHeadline(s, blue, black, label, textW) {
    pill(s, [black[0], black[1], black[2], 0.973], BLACK, { line: { color: BLACK, width: 1 } });
    pill(s, [blue[0], blue[1], blue[2], 0.973], BLUE, { line: { color: BLACK, width: 1 } });
    txt(s, [blue[0] + 0.113, blue[1] - 0.12, textW, 1.212], label,
        { fontFace: SEMI, fontSize: 60, align: 'center', color: '0D0D0D', margin: M_GS });
}

function slide06(s) {                                   // "Our Best Specialities"
    photoSlot(s, [0.407, 2.82, 2.869, 4.108]);
    photoSlot(s, [3.962, 1.024, 2.866, 4.108]);

    stickerHeadline(s, [7.402, 1.092, 2.952], [7.396, 1.188, 2.927], 'Best', 2.651);
    txt(s, [10.748, 1.082, 2.021, 1.212], 'Our', { fontFace: SEMI, fontSize: 60, margin: M_GS });
    txt(s, [7.197, 2.168, 6.137, 1.212], 'Specialities', { fontFace: SEMI, fontSize: 60, margin: M_GS });
    agreeGoBack(s, 10.798, 3.843, BLACK);

    shp(s, 'line', [3.1, 4.305, 1.293, 0], { line: { color: BLACK, width: 1.5, dashType: 'dash' } });
    ell(s, [3.521, 4.22, 0.17, 0.17], BLUE, { line: { color: BLACK, width: 1 } });
    ell(s, [8.931, 6.395, 0.17, 0.17], SAND, { line: { color: BLACK, width: 1 } });
    ell(s, [10.945, 6.395, 0.17, 0.17], LEMON, { line: { color: BLACK, width: 1 } });

    [[0.409, 1.247], [3.887, 5.492]].forEach(([x, y]) => {
        txt(s, [x, y, 2.444, 0.404], 'Your Specialities 001', { fontSize: 12, margin: M_GS });
        txt(s, [x, y + 0.404, 2.869, 0.837],
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
            { fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
    });

    [['+12,3%', 'Profit permonth', 7.297, 7.005, 6.077],
     ['+12,3%', 'Profit permonth', 9.558, 9.266, 6.08],
     ['125+', 'Client Permonth', 11.974, 11.682, 6.08]].forEach(([big, small, bx, sx, y]) => {
        txt(s, [bx, y, 1.325, 0.404], big);
        txt(s, [sx, y + 0.318, 1.909, 0.329], small,
            { fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
    });
}

function slide07(s) {                                   // "Our Finished Project" timeline
    txt(s, [0.971, 1.052, 5.786, 1.212], 'Our Finished',
        { fontFace: SEMI, fontSize: 60, margin: M_GS });
    stickerHeadline(s, [1.052, 2.31, 3.686], [1.057, 2.438, 3.648], 'Project', 3.447);
    agreeGoBack(s, 4.895, 2.69, WHITE);

    [['project 001', 0.792], ['Project 002', 3.957],
     ['Project 003', 7.234], ['Project 004', 10.515]].forEach(([name, x]) => {
        txt(s, [x, 5.332, 2.198, 0.404], name);
        txt(s, [x - 0.232, 5.732, 2.43, 0.835],
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
            { fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
    });

    photoSlot(s, [8.841, 1.235, 3.915, 2.408]);
    shp(s, 'line', [7.954, 4.343, 2.887, 0], { line: { color: BLACK, width: 1.5, dashType: 'dash' } });
    shp(s, 'line', [7.89, 2.375, 0, 1.883], { line: { color: BLACK, width: 1.5, dashType: 'dash' } });

    rrect(s, [10.721, 3.857, 2.035, 0.973], SAND, 0.10269, { line: { color: BLACK, width: 1 } });
    txt(s, [11.263, 4.028, 1.304, 0.404], '11.200+', { fontSize: 12, align: 'center', margin: M_GS });
    txt(s, [10.819, 4.235, 1.748, 0.454], 'Project On this Year',
        { fontSize: 10, align: 'center', lineSpacingMultiple: 1.5, margin: M_GS });

    rrect(s, [7.293, 1.235, 1.348, 1.289], BLUE, 0.10269, { line: { color: BLACK, width: 1 } });
    ell(s, [7.401, 1.345, 0.404, 0.404], WHITE, { line: { color: BLACK, width: 1 } });
    tick(s, [7.492, 1.467, 0.223, 0.157], BLUE);
    txt(s, [7.302, 1.892, 1.304, 0.404], '98,9%', { fontSize: 12, align: 'right', margin: M_GS });
    txt(s, [6.853, 2.081, 1.748, 0.454], 'Client Happy',
        { fontSize: 10, align: 'right', lineSpacingMultiple: 1.5, margin: M_GS });

    ell(s, [7.805, 4.258, 0.17, 0.17], BLUE, { line: { color: BLACK, width: 1 } });
    ell(s, [9.248, 4.258, 0.17, 0.17], WHITE, { line: { color: BLACK, width: 1 } });
    ell(s, [12.638, 3.765, 0.236, 0.236], LEMON, { line: { color: BLACK, width: 1 } });
    ell(s, [7.812, 2.947, 0.17, 0.17], SAND, { line: { color: BLACK, width: 1 } });
}

function slide08(s) {                                   // strategic networking
    rect(s, [1.0, 2.937, 5.523, 1.493], WHITE, { shadow: SH_SOFT });
    txt(s, [1.0, 1.019, 5.894, 1.447], [
        { text: 'The Art of Strategic ', options: { color: INK } },
        { text: 'Networking', options: { color: BLUE } },
    ], { fontFace: SEMI, fontSize: 40 });
    txt(s, [1.338, 3.198, 3.477, 0.438], 'Active Online Presence', { fontSize: 20, color: GREY40 });
    txt(s, [1.338, 3.568, 4.846, 0.602],
        'Leverage agile frameworks to provide a robust synopsis for high level overviews.',
        { fontSize: 12, color: GREY51, lineSpacingMultiple: 1.3 });

    [['Diversify Your Network', 1.005, 0.968], ['Offer Help and Value', 4.143, 4.139]]
        .forEach(([label, tx, bx], i) => {
            txt(s, [tx, 5.586 + i * 0.003, 2.261, 0.341], label,
                { fontSize: 12, color: INK, lineSpacingMultiple: 1.3 });
            txt(s, [bx, 5.931 + i * 0.003, 2.67, 0.626],
                'survival strategies to ensure proactive domination. ',
                { fontSize: 12, color: GREY51, lineSpacingMultiple: 1.3 });
        });

    /* sitemap glyph */
    rect(s, [1.24, 4.92, 0.3, 0.15], BLUE);
    rect(s, [1.151, 5.24, 0.19, 0.16], BLUE);
    rect(s, [1.445, 5.24, 0.19, 0.16], BLUE);
    shp(s, 'line', [1.39, 5.07, 0, 0.12], { line: { color: BLUE, width: 1.25 } });
    shp(s, 'line', [1.245, 5.17, 0.295, 0], { line: { color: BLUE, width: 1.25 } });
    /* waving-person glyph */
    ell(s, [4.3, 4.908, 0.14, 0.14], BLUE);
    rect(s, [4.33, 5.07, 0.09, 0.33], BLUE);
    shp(s, 'line', [4.42, 5.11, 0.09, -0.17], { line: { color: BLUE, width: 2 } });

    photoSlot(s, [7.838, 0.786, 4.572, 5.928]);
    rect(s, [7.167, 2.937, 1.398, 1.493], WHITE, { shadow: SH_SOFT });
    icon(s, [7.431, 3.194, 0.87, 0.979], BLUE);
}

function slide09(s) {                                   // About Us split
    rect(s, [7.228, 0, 6.106, 7.5], BLUE);
    txt(s, [0.894, 1.064, 3.555, 0.325], 'IT Solutions & Technology', { fontSize: 13.3, color: BLUE });
    txt(s, [0.804, 1.423, 4.994, 1.178], 'About Us', { fontFace: SEMI, fontSize: 64 });
    txt(s, [0.804, 2.694, 5.11, 0.887],
        'Lorem ipsum dolor sit amet, consectetuer adipiscing elitai Maecenas porttitor congue massa. '
        + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus ma lesuada',
        { fontSize: 10.7, color: GREY51, lineSpacingMultiple: 1.5 });

    [['Our Section About One', 3.96, 3.862, '1'], ['Our Section About Two', 5.395, 5.297, '2']]
        .forEach(([label, ty, cy, num]) => {
            txt(s, [0.804, ty, 2.894, 0.37], label, { fontSize: 16 });
            txt(s, [0.804, ty + 0.413, 4.719, 0.617],
                'Lorem ipsum dolor sit amet, consectetuer adipiscing elitar. Maecenas porttitor cong '
                + 'ue massa. Fusce posuere, magna',
                { fontSize: 10.7, color: GREY51, lineSpacingMultiple: 1.5 });
            rrect(s, [3.743, cy, 0.465, 0.452], BLUE, 0.09172);
            txt(s, [3.743, cy, 0.465, 0.452], num,
                { fontSize: 16, align: 'center', valign: 'middle', color: WHITE });
        });

    [0.75, 2.021].forEach((y) => {
        txt(s, [7.926, y, 1.615, 0.707], '278+', { fontFace: SEMI, fontSize: 36, color: WHITE });
        txt(s, [7.926, y + 0.673, 1.522, 0.325], 'Happy Clients', { fontSize: 13.3, color: WHITE });
    });
    txt(s, [10.028, 0.772, 2.412, 0.909], 'IT Solutions & Technology',
        { fontFace: SEMI, fontSize: 24, color: WHITE });
    txt(s, [10.028, 1.755, 2.549, 0.887],
        'Lorem ipsum dolor sit amet, con sectetuer adipiscing elitai Maec',
        { fontSize: 10.7, color: WHITE, lineSpacingMultiple: 1.5 });
    txt(s, [7.926, 3.192, 3.093, 0.37], 'IT Solutions Presentation', { fontSize: 16, color: WHITE });
    txt(s, [7.926, 3.691, 4.339, 1.156],
        'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor cong ue massa. '
        + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuadaar bero, sit amet '
        + 'commodo magna eros quis urna.',
        { fontSize: 10.7, color: WHITE, lineSpacingMultiple: 1.5 });
    txt(s, [7.926, 5.269, 4.339, 1.313], 'Fusce posuere, magna sedo raka pulvinar ultricies',
        { fontFace: SEMI, fontSize: 24, color: WHITE });
}

function slide10(s) {                                   // designer profile + skill bars
    rect(s, [10.681, 2.789, 2.644, 3.974], BLUE);       // from the slide layout
    photoSlot(s, [6.667, 0, 5.572, 7.5]);
    txt(s, [1.306, 1.75, 3.48, 1.043], 'Let\u2019s Meet Our Designer Here',
        { fontFace: SEMI, fontSize: 28, color: INK });
    txt(s, [1.306, 2.718, 2.8, 0.375], 'Professional Creative Agency',
        { fontSize: 12, color: INK, align: 'justify', lineSpacingMultiple: 1.5 });
    txt(s, [1.306, 3.409, 1.974, 0.466], 'Podoko Dune',
        { fontSize: 16, color: INK, align: 'justify', lineSpacingMultiple: 1.5 });
    txt(s, [1.306, 3.764, 2.625, 0.349], 'Our Designer',
        { fontSize: 11, color: INK, italic: true, lineSpacingMultiple: 1.5 });
    txt(s, [1.306, 4.092, 5.025, 0.934],
        'Aliquet nibh praesent tristique magna sit amet. Risus pretium micasa lo quam vulputate '
        + 'dignissim. Enim nec dui nunc mattis enim ut tellus elem consec entum nec dui ultrices.',
        { fontSize: 11, color: INK, align: 'justify', lineSpacingMultiple: 1.5 });

    [['90%', 5.109, 5.372, 5.493, 2.244, 2.644, 3.5],
     ['80%', 5.75, 6.006, 6.14, 2.207, 2.944, 3.2]].forEach(([pct, ly, py, by, tx, fx, fw]) => {
        rect(s, [tx, by, 3.9, 0.05], GREY91);
        rect(s, [fx, by, fw, 0.05], BLUE);
        txt(s, [1.37, ly, 2.102, 0.251], 'Type Your Skill Here',
            { fontSize: 11, color: INK, italic: true, margin: M0, lineSpacingMultiple: 1.5 });
        txt(s, [1.38, py, 0.579, 0.274], pct,
            { fontSize: 12, color: INK, margin: M0, align: 'justify', lineSpacingMultiple: 1.5 });
    });

    rect(s, [0.674, 5.692, 0.05, 1.07], BLUE);
    txt(s, [-0.407, 4.443, 2.227, 0.252], 'Creative And Innovative Solution',
        { fontSize: 9, color: INK, rotate: 270, wrap: false });
}

function slide11(s) {                                   // team grid
    ['Nala Hamprey', 'Travis Bill', 'Jodie Smith', 'Mash Harper'].forEach((name, i) => {
        const x = 0.818 + i * 3.068;
        photoSlot(s, [x, 3.157, 2.493, 2.303], { shape: 'roundRect', radius: 0.277 });
        pill(s, [x + 0.397, 5.23, 1.7, 0.485], BLUE);
        txt(s, [x + 0.397, 5.273, 1.7, 0.37], name,
            { fontSize: 16, color: WHITE, align: 'center', wrap: false });
        txt(s, [x - 0.252, 6.06, 3.027, 0.73],
            'Lorem ipsum dolor sit amet consectetur adipiscing vel eros',
            { fontSize: 13, color: '747474', align: 'center', lineSpacingMultiple: 1.5 });
    });
    txt(s, [3.292, 0.639, 6.75, 0.858], 'Meet Amazing Team',
        { fontFace: SEMI, fontSize: 45, color: '0D0D0D', align: 'center', wrap: false });
    txt(s, [1.33, 1.843, 10.675, 0.757],
        'PLACEHOLDER'
        + 'PLACEHOLDER'
        + 'eget duisali inenainisi atasanila eget duisalil.',
        { fontSize: 13, color: '3A3A3A', align: 'center', lineSpacingMultiple: 1.5 });
}

function slide12(s) {                                   // marketing journey
    photoSlot(s, [0, 0, 4.856, 3.75]);
    photoSlot(s, [4.845, 3.75, 3.632, 3.75]);
    rect(s, [0, 3.75, 4.856, 3.75], BLUE);
    poly(s, [0.553, 3.197, 3.75, 4.856], [
        ['M', 0, 0], ['L', 1, 0], ['L', 1, 0.4048], ['L', 0.993, 0.4194],
        ['C', 0.8572, 0.6747, 0.4922, 0.8855, 0.0105, 0.9977],
        ['L', 0, 1], ['Z'],
    ], { fill: { color: BLUE }, rotate: 90, flipH: true });

    txt(s, [5.392, 1.051, 6.921, 2.322], 'From Awareness to Action: The Marketing Journey',
        { fontFace: SEMI, fontSize: 44, color: INK });
    txt(s, [0.652, 4.614, 3.551, 0.303], 'Your Text Here', { fontSize: 12, color: WHITE });
    txt(s, [0.652, 5.051, 4.011, 1.111],
        'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. '
        + 'Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur '
        + 'ridiculus mus. Donec quam felis.',
        { fontSize: 12, color: WHITE });
    txt(s, [0.606, 6.58, 2.439, 0.303], '2025 Template', { fontSize: 12, color: WHITE, align: 'right' });

    rect(s, [8.234, 4.577, 4.303, 1.46], WHITE, { shadow: sh(47.8, 3, 45, 0.15) });
    txt(s, [8.624, 4.954, 2.403, 0.707],
        'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean.',
        { fontSize: 12, color: GREY35 });
    txt(s, [10.608, 4.954, 1.565, 0.707], '890+', { fontFace: SEMI, fontSize: 36, color: BLUE });
}

function slide13(s) {                                   // agenda capsules
    poly(s, [6.887, 1.369, 4.722, 2.44], CAPSULE, { fill: { color: BLUE } });
    poly(s, [6.887, 3.69, 4.722, 2.44], CAPSULE, { fill: { color: BLUE }, flipH: true });
    photoSlot(s, [7.052, 1.531, 2.117, 2.117]);
    photoSlot(s, [9.357, 3.852, 2.117, 2.117]);

    txt(s, [1.579, 2.051, 4.128, 1.717], 'About The Agenda', { fontFace: SEMI, fontSize: 48, color: INK });
    shp(s, 'line', [1.579, 3.852, 3.058, 0], { line: { color: BLUE, width: 3 } });
    txt(s, [1.579, 4.384, 2.705, 0.337], 'Today\u2019s Discussion', { fontSize: 14, color: INK });
    txt(s, [1.579, 4.83, 4.526, 0.633],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
        { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });

    [['Agenda 1', 9.378, 2.111], ['Agenda 2', 7.397, 4.433]].forEach(([label, x, y]) => {
        txt(s, [x, y, 1.407, 0.337], label, { fontSize: 14, color: WHITE });
        txt(s, [x, y + 0.336, 2.117, 0.62], 'Lorem ipsum dolor sit amet, consectetur',
            { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
    });
}

function slide14(s) {                                   // sales performance donuts
    photoSlot(s, [7.469, 0, 5.875, 7.5]);
    txt(s, [0.84, 0.905, 3.658, 0.337], 'Subtitle Here', { fontSize: 14, bold: true });
    txt(s, [0.84, 1.242, 6.382, 0.808], 'Sales Performance',
        { fontFace: SEMI, fontSize: 42, bold: true });

    [['Point One', 0.84, 0.654, 1.431, 1.535], ['Point Two', 3.907, 3.768, 4.545, 4.602]]
        .forEach(([label, cardX, chartX, discX, pctX]) => {
            rect(s, [cardX, 2.309, 2.476, 2.699], WHITE, { shadow: SH_PANEL });
            ringChart(s, [chartX, 2.46, 2.81, 1.873], [discX, 2.752, 1.251, 1.251]);
            txt(s, [pctX, 3.198, 1.081, 0.404], '60%',
                { fontFace: MED, bold: true, align: 'center' });
            txt(s, [cardX, 4.403, 2.476, 0.404], label, { bold: true, align: 'center' });
        });

    txt(s, [0.829, 5.234, 5.408, 0.555], 'Your text description here',
        { bold: true, color: '3F3F3F', lineSpacingMultiple: 1.5 });
    txt(s, [0.829, 5.761, 6.071, 0.808],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
        + 'ut labore et dolore magna',
        { fontSize: 14, color: GREY50, lineSpacingMultiple: 1.5 });
}

function slide15(s) {                                   // financial summary: pie + bar
    txt(s, [4.838, 1.063, 3.658, 0.337], 'Subtitle Here', { fontSize: 14, bold: true, align: 'center' });
    txt(s, [2.911, 1.415, 7.512, 0.909], 'Financial Summary',
        { fontFace: SEMI, fontSize: 48, bold: true, align: 'center' });

    /* left card - pie with a leader line */
    rect(s, [1.255, 2.66, 3.651, 3.962], WHITE, { shadow: SH_PANEL });
    s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] }],
        Object.assign({
            x: 1.723, y: 2.66, w: 2.715, h: 3.064,
            chartColors: [BLUE, 'BFBFBF', BLACK, 'BFBFBF'],
            dataBorder: { pt: 1.5, color: WHITE }, firstSliceAng: 0,
        }, CHART_BASE));
    shp(s, 'line', [1.87, 4.999, 1.031, 0], { line: { color: GREY85, width: 1 } });
    shp(s, 'line', [1.776, 4.999, 0.093, 0.271], { line: { color: GREY85, width: 1 } });
    txt(s, [1.344, 5.371, 0.958, 0.464], '70%',
        { fontFace: MED, align: 'center', lineSpacingMultiple: 1.2 });
    txt(s, [1.344, 5.987, 3.459, 0.384], 'Lorem ipsum dolor sit amet, ',
        { fontSize: 14, color: GREY51, align: 'center', lineSpacingMultiple: 1.2 });

    /* middle column */
    txt(s, [5.585, 3.52, 1.957, 0.37], 'Subtitle Here', { fontSize: 16, bold: true, align: 'center' });
    txt(s, [5.359, 4.069, 2.408, 0.949], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit',
        { fontSize: 14, color: GREY51, align: 'center', lineSpacingMultiple: 1.2 });
    txt(s, [5.718, 5.269, 1.69, 0.586], '3239 +',
        { fontFace: MED, fontSize: 24, bold: true, align: 'center', lineSpacingMultiple: 1.2 });

    /* right card - column chart */
    rect(s, [8.441, 2.66, 3.651, 3.962], WHITE, { shadow: SH_PANEL });
    txt(s, [8.648, 3.026, 1.22, 0.64], '38%', { fontFace: MED, fontSize: 32, bold: true });
    txt(s, [9.695, 2.973, 2.463, 0.666], 'Lorem ipsum dolor sit amet, consectetuer.',
        { fontSize: 14, color: GREY51, lineSpacingMultiple: 1.2 });
    s.addChart('bar', [{ name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] }],
        Object.assign({
            x: 8.729, y: 3.891, w: 3.141, h: 2.457,
            barDir: 'col', chartColors: [BLUE], barGapWidthPct: 122,
            catAxisHidden: true, valAxisLineShow: false,
            valAxisMaxVal: 5, valAxisMajorUnit: 0.5,
            valAxisLabelFontFace: REG, valAxisLabelFontSize: 12, valAxisLabelColor: GREY35,
            valGridLine: { style: 'solid', size: 0.75, color: 'D9D9D9' },
        }, CHART_BASE));
}

function slide16(s) {                                   // vision & mission
    photoSlot(s, [3.432, 0.774, 3.334, 3.333]);
    photoSlot(s, [0.765, 3.758, 2.668, 2.667]);

    rrect(s, [3.72, 4.329, 1.868, 1.868], BLUE, 0.15924);
    txt(s, [3.757, 4.611, 1.795, 1.414],
        'Lorem ipsum dolor amet, adipiscing do eiusmod tempor incididunt ut',
        { fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.3 });
    rrect(s, [0.997, 1.41, 2.203, 2.203], BLUE, 0.15924);
    txt(s, [1.395, 1.792, 1.408, 0.505], [
        { text: '10', options: { fontSize: 24 } },
        { text: '/mbps', options: { fontSize: 16 } },
    ], { fontFace: LIGHT, bold: true, color: WHITE, align: 'center' });
    txt(s, [1.092, 2.179, 2.042, 1.151],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit eiusmod ut labore.',
        { fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.3 });

    txt(s, [7.58, 2.027, 4.984, 1.934], 'Vision &Mission',
        { fontFace: SEMI, fontSize: 66, color: BLUE, lineSpacingMultiple: 0.8 });
    ell(s, [7.621, 4.28, 0.997, 0.997], BLUE);
    icon(s, [7.79, 4.45, 0.66, 0.66], WHITE);
    txt(s, [8.877, 4.132, 3.944, 0.425], 'Stability Comes From Here',
        { fontSize: 14, bold: true, lineSpacingMultiple: 1.5 });
    txt(s, [8.891, 4.593, 3.944, 1.151],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
        + 'ut labore et dolore magna aliqua. Ut enim ad minim veniam',
        { fontSize: 12, color: GREY35, lineSpacingMultiple: 1.3 });
}

function slide17(s) {                                   // big project list
    photoSlot(s, [1.7, 1.389, 4.763, 6.111]);
    txt(s, [7.614, 1.887, 4.499, 1.313], [
        { text: 'Big Project Our ', options: { color: INK } },
        { text: 'Bank In 2024', options: { color: BLUE } },
    ], { fontFace: SEMI, fontSize: 36, bold: true });

    ['01', '02', '03'].forEach((n, i) => {
        const y = 3.678 + i * 1.119;
        tickCircle(s, [7.726, y + 0.089, 0.23, 0.225], BLUE);
        txt(s, [8.012, y, 1.451, 0.303], 'Big Project ' + n,
            { fontFace: REG, fontSize: 12, bold: true, color: BLUE, align: 'justify' });
        txt(s, [8.012, y + 0.231, 4.239, 0.609],
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
            { fontFace: REG, fontSize: 10.5, color: INK, align: 'justify', lineSpacingMultiple: 1.5 });
    });

    rrect(s, [6.046, 3.274, 0.834, 0.834], BLUE, 0.16667, { shadow: sh(30, 15, 45, 0.20) });
    txt(s, [6.046, 3.274, 0.834, 0.834], '50+',
        { fontFace: 'Open Sans', fontSize: 20, bold: true, color: WHITE,
          align: 'center', valign: 'middle' });
    rrect(s, [1.283, 5.154, 0.834, 0.834], BLUE, 0.16667, { shadow: sh(30, 15, 45, 0.20) });
    rect(s, [1.609, 5.39, 0.25, 0.362], WHITE);
    rect(s, [1.657, 5.45, 0.09, 0.14], BLUE);
}

function slide18(s) {                                   // bitcoin guide
    /* grey slab: a 6.433 x 5.501 round2DiagRect turned 270 deg and mirrored,
       written out here in its final on-slide orientation */
    poly(s, [7.487, 0.24, 5.501, 6.433], round2Diag(0.3943, 0.3372), { fill: { color: GREY51 } });
    /* the template's picture prompt sits on top of it */
    rect(s, [7.835, 0, 5.501, 6.433], 'FCFEF6');
    txt(s, [7.835, 0.04, 5.501, 0.42], 'Drag Picture here',
        { fontSize: 20, color: GREY40, align: 'center' });
    rect(s, [10.19, 3.02, 0.8, 0.58], WHITE, { line: { color: GREY51, width: 0.75 } });
    poly(s, [10.26, 3.16, 0.66, 0.38], [
        ['M', 0, 1], ['L', 0.38, 0.18], ['L', 0.68, 0.66], ['L', 0.82, 0.46],
        ['L', 1, 1], ['Z'],
    ], { fill: { color: '5B9BD5' } });
    ell(s, [10.72, 3.1, 0.13, 0.13], 'FFC000');

    txt(s, [1.021, 1.216, 1.577, 0.27], 'ABOUT US',
        { fontSize: 10, bold: true, charSpacing: 1.5, color: INK });
    txt(s, [0.976, 1.637, 6.425, 1.767], [
        { text: 'THE STEP-BY-STEP GUIDE TO EVERYTHING BITCOIN &', options: { breakLine: true } },
        { text: 'CRYPTOCURRENCY' },
    ], { fontFace: SEMI, fontSize: 33, bold: true, color: INK });
    txt(s, [0.976, 3.502, 6.259, 0.985],
        'Fusce vehicula dolor arcu, sit amet blandit dolor mollis nec. Donec viverra eleifend lacus, '
        + 'vitae ullamcorper metus. Sed sollicitudin ipsum quis nunc sollicitudin ultrices sit amet '
        + 'blandit dolor. ',
        { fontSize: 12, color: INK, lineSpacingMultiple: 1.5 });
    txt(s, [0.976, 4.72, 2.719, 0.37], 'TYPES OF INVESTMENTS',
        { fontSize: 16, bold: true, color: INK });

    [['Total Crypto Overload \u2013 Cryptocurrency Basics and Beyond', 5.33, 0.303],
     ['How to Buy, Sell, and Store Bitcoin and Cryptocurrency', 5.677, 0.303],
     ['PLACEHOLDER', 6.024, 0.505]]
        .forEach(([t, y, h]) => {
            ell(s, [0.976, y + 0.084, 0.086, 0.085], BLUE);
            txt(s, [1.182, y, 5.207, h], t, { fontSize: 12, italic: true, color: INK });
        });
}

function slide19(s) {                                   // medical check up
    photoSlot(s, [0, 0, 13.333, 6.182]);
    pill(s, [12.175, 0.467, 1.227, 0.338], BLUE);
    txt(s, [12.175, 0.51, 1.04, 0.252], 'MEDICOZE', { fontSize: 9, color: WHITE, align: 'center' });

    shp(s, 'round1Rect', [0.76, 2.779, 5.906, 4.042], {
        fill: { color: WHITE }, flipH: true, rectRadius: 0.08556 * 4.042, shadow: sh(45, 22, 25, 0.20),
    });
    shp(s, 'rtTriangle', [0.755, 2.779, 0.63, 0.599], { fill: { color: BLUE }, flipV: true });
    txt(s, [1.621, 3.499, 4.184, 0.909], 'We Always Believe in Patient Healing',
        { fontSize: 24, color: INK, align: 'center' });
    txt(s, [1.76, 4.68, 3.907, 1.403],
        'Lorem ipsum dolor sit amet, consectetur adipiscinge eoeoe Pellentesque sceleri larakani '
        + 'malesuadloo libero aeo woooe ellentesque. Morbi orcieretti dui,galang ferment eget lectus',
        { fontSize: 11, align: 'center', lineSpacing: 24 });

    rrect(s, [8.362, 3.942, 5.374, 2.879], BLUE, 0.08429, { shadow: sh(45, 22, 25, 0.20) });
    txt(s, [8.982, 4.272, 3.232, 0.337], '#MedicalHealth', { fontSize: 14, color: WHITE });
    txt(s, [8.982, 4.664, 3.358, 1.582], 'Medical Check Up',
        { fontFace: SEMI, fontSize: 44, color: WHITE });
}

function slide20(s) {                                   // brand strategy
    photoBox(s, [7.08, 1.111, 5.836, 5.833]);
    rrect(s, [0.518, 0.957, 1.689, 0.694], BLUE, 0.20616);
    txt(s, [0.518, 0.957, 1.689, 0.694], '   motivated',
        { fontSize: 16, color: WHITE, valign: 'middle', margin: M_SM });
    txt(s, [0.417, 2.08, 6.667, 3.035], [
        { text: 'Brand', options: { breakLine: true } },
        { text: 'Strategy' },
    ], { fontFace: SEMI, fontSize: 120, color: BLUE, charSpacing: -9.59,
         lineSpacingMultiple: 0.7, margin: M_SM });
    txt(s, [0.622, 5.278, 5.2, 0.41],
        'Cras justo odio, dapibus ac facilisis in, egestas eget quam. Etiam porta sem malesuada '
        + 'magna mollis euismod.',
        { fontFace: REG, fontSize: 9, color: '010203', charSpacing: 0.72,
          lineSpacingMultiple: 1.2, margin: M_SM });
}

function slide21(s) {                                   // logo story
    rect(s, [6.254, 0, 7.08, 7.5], BLUE);
    rect(s, [0, -0.041, 0.563, 3.772], WHITE);
    rect(s, [0, 2.952, 0.563, 2.057], BLUE);

    rect(s, [1.482, 1.486, 3.787, 3.541], BLUE);
    [[2.877, 1.932], [2.399, 2.888], [3.356, 2.888]].forEach(([x, y]) => {
        shp(s, 'cube', [x, y, 0.957, 0.955], { line: { color: WHITE, width: 2.5 } });
    });
    txt(s, [2.228, 4.014, 2.295, 0.841], 'CUBE',
        { fontFace: SEMI, fontSize: 44, color: WHITE, align: 'center', margin: M_GS });
    txt(s, [1.482, 5.199, 3.787, 0.37], 'Logo Preview',
        { fontSize: 16, italic: true, align: 'center', margin: M_GS });

    txt(s, [6.784, 1.614, 5.424, 1.941], 'Story About Brand Logo',
        { fontFace: SEMI, fontSize: 54.7, color: WHITE, margin: M_GS });
    txt(s, [6.784, 3.75, 6.321, 1.851],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis '
        + 'rhoncus mattis. Orci varius natoque penatibus et magnis dis parturient montes, nascetur '
        + 'ridiculus mus. Pellentesque habitant morbi tristique senectus et netus et malesuada.',
        { fontSize: 16, color: WHITE, lineSpacingMultiple: 1.3, margin: M_GS });
}

function slide22(s) {                                   // powerful first impression
    rect(s, [0.553, 0, 12.769, 3.75], BLUE);            // from the slide layout
    rect(s, [0.003, 3.75, 0.561, 0.561], BLUE);
    photoSlot(s, [7.706, 1.121, 4.485, 5.258]);

    txt(s, [1.485, 1.499, 4.162, 1.33], [
        { text: 'Powerful First', options: { breakLine: true } },
        { text: 'Impression.' },
    ], { fontFace: SEMI, fontSize: 40, color: WHITE, lineSpacingMultiple: 0.9 });
    shp(s, 'line', [1.633, 3.037, 0.699, 0], { line: { color: BLUE, width: 1 } });
    txt(s, [1.485, 4.291, 4.485, 0.633],
        'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. ',
        { fontSize: 11, lineSpacingMultiple: 1.5 });
    txt(s, [1.485, 5.097, 5.629, 0.904],
        'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum '
        + 'has been the industry\u2019s standard dummy text of Ipsum\u00a0is simply the printing and '
        + 'industry\u2019s standard.',
        { fontSize: 11, lineSpacingMultiple: 1.5 });
    txt(s, [0.566, 6.939, 0.635, 0.286], '22',
        { fontFace: 'Open Sans', fontSize: 11, color: GREEN, align: 'right' });
    shp(s, 'line', [0.758, 7.225, 0, 0.275], { line: { color: LIME, width: 1 } });
}

function slide23(s) {                                   // introduction
    photoSlot(s, [0.358, 0.729, 3.854, 3.854]);
    photoSlot(s, [9.951, 3.888, 2.766, 2.766]);
    txt(s, [4.854, 1.59, 4.752, 0.909], 'Introduction', { fontFace: SEMI, fontSize: 48, color: INK });
    shp(s, 'line', [4.854, 2.701, 3.841, 0], { line: { color: BLUE, width: 3 } });
    txt(s, [4.854, 3.698, 2.705, 0.337], 'Who We Are', { fontSize: 14, color: INK });
    txt(s, [4.854, 4.143, 5.089, 0.911],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
        + 'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercit',
        { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });
    txt(s, [4.854, 5.1, 5.089, 0.633],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
        + 'ut labore et dolore mag.',
        { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });

    ell(s, [2.017, 3.75, 2.203, 2.203], BLUE);
    ell(s, [2.9, 3.964, 0.434, 0.434], WHITE);
    ell(s, [3.032, 4.084, 0.17, 0.177], null, { line: { color: BLUE, width: 1.25 } });
    ell(s, [3.09, 4.145, 0.055, 0.055], BLUE);
    txt(s, [2.319, 4.46, 1.646, 0.337], 'Subtitle Here',
        { fontSize: 14, color: WHITE, align: 'center' });
    txt(s, [2.091, 4.822, 2.055, 0.897], 'Lorem ipsum dolor sit amet, consectetur adipiscing.',
        { fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
}

function slide24(s) {                                   // about the vision
    shp(s, 'donut', [1.299, 0.926, 5.027, 5.027], { fill: { color: BLUE }, rectRadius: 0.03133 * 5.027 });
    photoSlot(s, [1.548, 1.194, 4.512, 4.492]);
    photoSlot(s, [1.355, 3.833, 2.741, 2.741]);

    txt(s, [7.007, 1.281, 4.252, 1.717], 'About The Vision',
        { fontFace: SEMI, fontSize: 48, color: INK });
    shp(s, 'line', [7.286, 2.998, 1.927, 0], { line: { color: BLUE, width: 3 } });
    ['01', '02', '03'].forEach((n, i) => {
        const y = 3.674 + i * 0.853;
        ell(s, [7.386, y + 0.018, 0.606, 0.606], BLUE);
        txt(s, [7.296, y + 0.114, 0.788, 0.438], n,
            { fontSize: 20, color: WHITE, align: 'center' });
        txt(s, [8.191, y, 3.548, 0.633],
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
            { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });
    });
}

function slide25(s) {                                   // building a better tomorrow
    photoSlot(s, [4.909, 3.728, 8.406, 2.854], { flipH: true });
    txt(s, [1.019, 1.251, 4.896, 1.447], 'Building a Better Tomorrow',
        { fontFace: SEMI, fontSize: 40, color: INK });
    shp(s, 'line', [1.198, 2.828, 2.798, 0], { line: { color: BLUE, width: 3 } });

    /* person + cog, and twin cogs */
    ell(s, [8.143, 1.649, 0.21, 0.21], BLUE);
    poly(s, [8.143, 1.89, 0.32, 0.187], [
        ['M', 0, 1], ['C', 0, 0.3, 1, 0.3, 1, 1], ['Z'],
    ], { fill: { color: BLUE } });
    ell(s, [8.42, 1.86, 0.222, 0.222], null, { line: { color: BLUE, width: 3 } });
    ell(s, [11.147, 1.588, 0.2, 0.2], null, { line: { color: BLUE, width: 3 } });
    ell(s, [11.36, 1.83, 0.286, 0.262], null, { line: { color: BLUE, width: 3.5 } });

    [6.521, 9.448].forEach((x) => {
        txt(s, [x, 1.854, 1.628, 0.337], 'Subtitle Here', { fontSize: 14, color: INK });
        txt(s, [x, 2.191, 2.321, 0.897],
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
            { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });
    });

    ell(s, [1.019, 3.728, 2.854, 2.854], BLUE);
    txt(s, [1.969, 4.1, 0.979, 0.572], '94%', { fontSize: 28, color: WHITE, align: 'center' });
    txt(s, [1.492, 4.569, 1.934, 0.337], 'Company Growth',
        { fontSize: 14, color: WHITE, align: 'center' });
    txt(s, [1.446, 5.019, 2.025, 1.189],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
        { fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
}

function slide26(s) {                                   // our service
    const cards = [
        ['Service 1', 2.304, 0.854, WHITE, INK, 2.511, 0.987, 4.135, 0.63],
        ['Service 2', 1.542, 2.834, BLUE, WHITE, 1.7, 2.967, 3.369, 0.62],
        ['Service 3', 2.304, 4.814, WHITE, INK, 2.511, 4.947, 4.135, 0.63],
    ];
    cards.forEach(([label, x, y, bg, fg, px, py]) => {
        pill(s, [x, y, 4.458, 1.832], bg, bg === WHITE ? { shadow: SH_CARD } : {});
    });
    cards.forEach(([label, x, y, bg, fg, px, py, tx, bh]) => {
        photoSlot(s, [px, py, 1.567, 1.567]);
        txt(s, [tx, y + 0.438, 1.427, 0.337], label, { fontSize: 14, color: fg });
        txt(s, [tx, y + 0.774, 2.329, bh], 'Lorem ipsum dolor sit amet, mini consectetur minim.',
            { fontSize: 11, color: fg, lineSpacingMultiple: 1.5 });
    });

    txt(s, [7.771, 2.253, 3.771, 0.774], 'Our Service', { fontFace: SEMI, fontSize: 40, color: INK });
    shp(s, 'line', [7.771, 3.188, 3.088, 0], { line: { color: BLUE, width: 3 } });
    txt(s, [7.771, 3.64, 2.042, 0.337], 'Tailored Solution', { fontSize: 14, color: INK });
    txt(s, [7.771, 4.072, 3.771, 1.175],
        'Lorem ipsum dolor sit amet, consectetur minim adipiscing elit, sed do eiusmod tempor '
        + 'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercit.',
        { fontSize: 11, color: INK, lineSpacingMultiple: 1.5 });
}

function slide27(s) {                                   // meet our team
    txt(s, [3.889, 1.019, 5.556, 0.774], 'Meet Our Team',
        { fontFace: SEMI, fontSize: 40, color: INK, align: 'center' });
    shp(s, 'line', [4.746, 1.843, 3.841, 0], { line: { color: BLUE, width: 3 } });

    const people = [
        ['Lionell Doe', 'General Manager', 1.693, 2.478, WHITE, INK, 1.849, 2.613, 3.781, 3.206],
        ['Anthony Doe', 'CEO', 4.459, 4.435, WHITE, INK, 4.617, 4.571, 6.566, 5.163],
        ['Edward Doe', 'Founder', 7.681, 2.478, BLUE, WHITE, 7.804, 2.613, 9.62, 3.208],
    ];
    people.forEach(([, , x, y, bg]) => {
        poly(s, [x, y, 3.959, 2.046], CAPSULE,
            Object.assign({ fill: { color: bg } }, bg === WHITE ? { shadow: SH_CARD } : {}));
    });
    people.forEach(([name, role, , , bg, fg, px, py, tx, ty]) => {
        photoSlot(s, [px, py, 1.775, 1.775]);
        txt(s, [tx, ty, 1.712, 0.337], name, { fontSize: 14, color: fg });
        txt(s, [tx, ty + 0.304, 1.531, 0.286], role, { fontSize: 11, color: fg });
    });
}

function slide28(s) {                                   // pricing
    photoSlot(s, [6.932, 3.075, 6.389, 3.469], { flipH: true });
    txt(s, [6.952, 0.909, 4.54, 1.447], 'Value For Your Investment',
        { fontFace: SEMI, fontSize: 40, color: INK });
    shp(s, 'line', [6.952, 2.471, 3.381, 0], { line: { color: BLUE, width: 3 } });

    const plans = [
        ['Business Plan', '$65.9', 1.239, WHITE, INK, BLUE, WHITE, BLUE, WHITE, 0.63],
        ['Premium Plan', '$85.9', 3.832, BLUE, WHITE, WHITE, BLUE, WHITE, INK, 0.62],
    ];
    plans.forEach(([name, price, y, cardBg, cardFg, discBg, discFg, btnBg, btnFg, bh], i) => {
        rrect(s, [1.841, y, 4.286, 2.429], cardBg, 0.11438, i === 0 ? { shadow: SH_CARD } : {});
        ell(s, [4.232, y + 0.377, 1.676, 1.676], discBg);
        txt(s, [4.261, y + 0.835, 1.619, 0.64], price,
            { fontFace: SEMI, fontSize: 32, color: discFg, align: 'center' });
        txt(s, [4.969, y + 1.309, 0.812, 0.286], '/month', { fontSize: 11, color: discFg });
        txt(s, [2.085, y + 0.466, 1.929, 0.337], name, { fontSize: 14, color: cardFg });
        txt(s, [2.085, y + 0.843, 1.885, bh], 'Lorem ipsum dolor sit amet, consectetur.',
            { fontSize: 11, color: cardFg, lineSpacingMultiple: 1.5 });
        pill(s, [2.164, y + 1.613, 1.274, 0.351], btnBg);
        txt(s, [2.191, y + 1.654, 1.221, 0.269], 'Choose Plan',
            { fontSize: 10, color: btnFg, align: 'center' });
    });
}

function slide29(s) {                                   // mockup showcase
    photoSlot(s, [3.535, 2.128, 3.375, 1.977], { rotate: 19.26 });
    photoBox(s, [2.0, 1.55, 4.9, 4.4], { color: PH_SOFT, label: GREY51 });
    txt(s, [7.551, 1.776, 5.46, 1.447], 'Mockup Showcase',
        { fontFace: SEMI, fontSize: 40, color: INK, wrap: false });
    txt(s, [7.551, 3.056, 2.587, 0.337], 'Our Business Concept', { fontSize: 14, color: INK });
    txt(s, [7.551, 3.393, 4.988, 0.911],
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
        + 'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercit..',
        { fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5 });
    pill(s, [7.514, 5.471, 4.385, 0.257], GREY85);
    pill(s, [7.514, 5.471, 3.896, 0.257], BLUE);
    txt(s, [7.514, 4.809, 1.202, 0.572], 'Succes Rate', { fontSize: 14, color: INK });
    txt(s, [8.436, 4.809, 1.288, 0.572], '97%', { fontFace: SEMI, fontSize: 28, color: BLUE });

    ell(s, [1.481, 0.951, 1.675, 1.675], BLUE);
    icon(s, [2.051, 1.22, 0.54, 0.54], WHITE);
    txt(s, [1.59, 1.777, 1.462, 0.572], 'Add Some Text Here',
        { fontSize: 14, color: WHITE, align: 'center' });
}

function slide30(s) {                                   // thank you
    photoSlot(s, [0.002, 0, 6.171, 6.825], { flipH: true });
    txt(s, [6.819, 2.114, 5.027, 1.111], 'Thank You', { fontFace: SEMI, fontSize: 60, color: BLUE });
    txt(s, [6.886, 3.11, 4.535, 0.64], 'For Your Attention',
        { fontFace: SEMI, fontSize: 32, color: INK });
    shp(s, 'line', [7.023, 3.864, 3.841, 0], { line: { color: INK, width: 3 } });

    pill(s, [3.978, 4.503, 7.868, 1.547], BLUE);
    const contact = [
        ['Phone :', 4.288, 1.077, ['012 \u2013 3456 7890', '012 \u2013 3456 7891'], 1.596],
        ['Email :', 6.429, 0.889, ['companyname@email.com', 'nameofcompany@email.com'], 2.514],
    ];
    contact.forEach(([head, x, hw, lines, lw]) => {
        txt(s, [x, 4.817, hw, 0.337], head, { fontSize: 14, color: WHITE });
        lines.forEach((l, i) => {
            txt(s, [x, 5.176 + i * 0.276, lw, 0.286], l, { fontSize: 11, color: WHITE });
        });
    });
    txt(s, [9.481, 4.817, 1.22, 0.337], 'Address :', { fontSize: 14, color: WHITE });
    txt(s, [9.481, 5.107, 1.953, 0.62], 'Dahlia 616B, Surabaya, East Java, Indonesia',
        { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
}

/* ==================================================================== build */

const BUILDERS = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
    pptx.layout = 'W16x9';
    pptx.theme = { headFontFace: SEMI, bodyFontFace: LIGHT };
    pptx.title = 'Business Summary';

    BUILDERS.forEach((fn) => {
        const slide = pptx.addSlide();
        slide.background = { color: WHITE };
        fn(slide);
    });

    return pptx.writeFile({
        fileName: path.join(__dirname, '106af8d9-b440-46f9-a86e-e8f94f0d0bbe_grok_final.pptx'),
    });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => {
    console.error(e);
    process.exit(1);
});
