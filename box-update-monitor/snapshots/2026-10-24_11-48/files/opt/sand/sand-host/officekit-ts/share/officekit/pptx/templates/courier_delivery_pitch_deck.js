/**
 * QuickCourier "Speedy Delivery Service" deck - 25 slides, 20 x 11.25 in.
 * Rebuilt with pptxgenjs only. Raster photos in the original are replaced by
 * flat "[image]" placeholder cards of the same geometry.
 */
'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */
const BLUE = '5D8EFA';      // brand blue (theme dk2) - also the dark slide bg
const BLACK = '000000';
const WHITE = 'FFFFFF';
const SAND = 'D5D4C7';      // theme accent1
const PILL = 'D5D4C7';      // rounded button fill
const PH_FILL = 'F2F2F2';   // photo placeholder card
const PH_TEXT = '999999';
const AXIS = 'D9D9D9';      // chart category-axis rule

const HEAD = 'Lexend Deca SemiBold';   // theme major font
const BODY = 'Plus Jakarta Sans';      // theme minor font

const SW = 20;
const SH = 11.25;

/* --------------------------------------------------------------- helpers */
const pres = new pptxgen();
pres.defineLayout({ name: 'DECK', width: SW, height: SH });
pres.layout = 'DECK';

/** plain text box (top anchored, like every TextBox in the source deck) */
function T(s, text, x, y, w, h, o) {
    s.addText(text, Object.assign({
        x: x, y: y, w: w, h: h,
        fontFace: BODY, fontSize: 16, color: BLACK,
        valign: 'top', align: 'left', margin: [7.2, 7.2, 3.6, 3.6] // l, r, b, t
    }, o || {}));
}

/** section / slide title */
function title(s, text, x, y, w, h, o) {
    T(s, text, x, y, w, h, Object.assign({ fontFace: HEAD, fontSize: 60, lineSpacingMultiple: 1.0 }, o || {}));
}

/** 0.21" bullet / anchor dot */
function dot(s, x, y, fill, line) {
    s.addShape('ellipse', {
        x: x, y: y, w: 0.21, h: 0.21,
        fill: { color: fill },
        line: line ? { color: line, width: 1 } : { type: 'none' }
    });
}

/**
 * Photo placeholder card: flat rounded rectangle + "[image]" caption.
 * The source deck fills these frames with square "Your Photo Here" stock art,
 * whose caption scales with the frame - hence the size-proportional font.
 */
function photo(s, x, y, w, h, adj) {
    const radius = ((adj || 5215) / 100000) * Math.min(w, h);
    s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: radius, fill: { color: PH_FILL }, line: { type: 'none' } });
    s.addText('[image]', {
        x: x, y: y, w: w, h: h, align: 'center', valign: 'middle',
        fontFace: 'Lexend Deca Light', fontSize: Math.round(5.4 * Math.max(w, h)), color: PH_TEXT
    });
}

/** the thin outlined arc that sweeps across most slides */
function arc(s, x, y, w, h, a1, a2, o) {
    s.addShape('arc', Object.assign({
        x: x, y: y, w: w, h: h,
        angleRange: [a1, a2],
        line: { color: BLACK, width: 1 }
    }, o || {}));
}

/** outlined circle with a diagonal arrow inside (slides 1, 4, 23) */
function arrowCircle(s, x, y, color) {
    s.addShape('ellipse', { x: x, y: y, w: 1.268, h: 1.268, fill: { type: 'none' }, line: { color: color, width: 1 } });
    s.addShape('line', {
        x: x + 0.395, y: y + 0.379, w: 0.479, h: 0.511, flipV: true,
        line: { color: color, width: 1, endArrowType: 'arrow' }
    });
}

/**
 * The brand mark: { black disc + accent disc + pupil + arrow } inside a brace pair.
 * `flip` mirrors the whole group (used where the arrow points left).
 */
function mark(s, x, y, o) {
    o = o || {};
    const ink = o.ink || BLACK;          // brace / arrow / black disc
    const accent = o.accent || BLUE;     // the coloured disc
    const edge = o.edge || WHITE;        // outline around the coloured disc
    const W = 2.208;
    const put = (dx, w) => x + (o.flip ? W - dx - w : dx);
    s.addShape('bracePair', { x: x, y: y, w: W, h: 0.853, fill: { type: 'none' }, line: { color: ink, width: 1 } });
    s.addShape('ellipse', { x: put(0.362, 0.676), y: y + 0.088, w: 0.676, h: 0.676, fill: { color: ink }, line: { type: 'none' } });
    s.addShape('ellipse', { x: put(0.722, 0.676), y: y + 0.088, w: 0.676, h: 0.676, fill: { color: accent }, line: { color: edge, width: 1 } });
    s.addShape('ellipse', { x: put(1.003, 0.153), y: y + 0.352, w: 0.153, h: 0.153, fill: { color: ink }, line: { type: 'none' } });
    s.addShape('line', {
        x: put(1.076, 0.817), y: y + 0.433, w: 0.817, h: 0,
        line: o.flip
            ? { color: ink, width: 1, beginArrowType: 'arrow' }
            : { color: ink, width: 1, endArrowType: 'arrow' }
    });
}

/** sand coloured call-to-action button */
function pill(s, x, y, w, label, tx, ty, tw) {
    s.addShape('roundRect', { x: x, y: y, w: w, h: 0.879, rectRadius: 0.24592 * 0.879, fill: { color: PILL }, line: { type: 'none' } });
    T(s, label, tx, ty, tw, 0.572, { fontFace: HEAD, fontSize: 28 });
}

/** header bar shared by every slide: logo, wordmark, nav links, page number */
function chrome(s, page, dark, pageX, pageW) {
    const ink = dark ? WHITE : BLACK;
    const disc = dark ? SAND : BLUE;
    const ring = dark ? BLUE : SAND;
    dot(s, 1.104, 0.56, disc, ring);
    dot(s, 1.24, 0.56, disc, ring);
    T(s, 'QuickCourier', 1.638, 0.483, 2.356, 0.37, { color: ink });
    T(s, 'Business', 6.77, 0.48, 1.844, 0.37, { align: 'center', color: ink });
    T(s, 'Delivery', 8.75, 0.48, 1.988, 0.37, { align: 'center', color: ink });
    T(s, 'Presentation', 10.874, 0.48, 2.356, 0.37, { align: 'center', color: ink });
    T(s, page, pageX || 17.866, 0.48, pageW || 1.03, 0.37, { align: 'right', color: ink });
}

/** heading + paragraph pair preceded by a dot - the deck's workhorse block */
function bullet(s, b) {
    const ink = b.color || BLACK;
    dot(s, b.dx, b.dy, ink === WHITE ? WHITE : BLUE);
    T(s, b.head, b.x, b.hy, b.hw, 0.37, { fontFace: b.headFont || HEAD, bold: !!b.bold, fontSize: b.headSize || 16, color: ink });
    T(s, b.text, b.x, b.by, b.bw, b.bh || 0.64, { fontFace: b.bodyFont || BODY, color: ink });
}

function newSlide(dark) {
    const s = pres.addSlide();
    if (dark) s.background = { color: BLUE };
    return s;
}

/* ------------------------------------------------ recurring copy strings */
const QUOTE_LONG = '"Delivering Excellence to Your Doorstep: Speed, Precision, and a Passion for Perfect Deliveries, Because We Understand Your Expectations and Value Your Trust"';
const QUOTE_SHORT = '"PLACEHOLDER"';
const LOREM_A = 'Odales ut eu sem integer vitae justo eget. Pellentesque dignissim enim sit amet venenatis urna cursus eget nunc. Ultrices tincidunta rcu non sodales neque sodales ut. ';
const LOREM_B = 'Vitae. Et sollicitudin ac orci phasellus egestas. Sapien nec sagittis aliquam malesuada bibendum arcu';
const SOL_1 = 'Justo eget magna fermentum iaculis eu non. Quam adipiscing vitae';
const SOL_2 = 'Eget dolor morbi non arcu risus quis varius quam. Non curabitur gravida';
const SOL_3 = 'Proin sagittis nisl rhoncus mattis rhoncus. Purus non enim praesent';
const SOL_4 = 'arcu ac tortor. Sapien nec sagittis aliqu am malesuada bibendum';

/* ============================================================== slide 01 */
function slide01() {
    const s = newSlide(false);
    photo(s, 14.963, 1.88, 3.92, 4.55);
    photo(s, 11.319, 4.866, 4.972, 5.342);
    chrome(s, '2030', false);
    title(s, 'Speedy Delivery Service  ', 1.126, 2.495, 9.757, 2.794, { fontSize: 80 });
    arrowCircle(s, 17.01, 7.688, BLACK);
    mark(s, 6.195, 4.21);
    arc(s, 10.502, 3.018, 5.603, 5.714, 136.44, 305.73);
    T(s, QUOTE_LONG, 1.126, 5.97, 7.108, 0.909);
    [['129K+', 1.282, 1.53], ['362+', 3.754, 1.236], ['98%', 5.932, 1.099]].forEach(function (v) {
        T(s, v[0], v[1], 8.01, v[2], 0.572, { fontFace: HEAD, fontSize: 28, align: 'center' });
    });
    [['Parcel Delivered', 1.013, 2.069], ['Customers', 3.521, 1.701], ['Reviews', 5.859, 1.244]].forEach(function (v) {
        T(s, v[0], v[1], 8.64, v[2], 0.37, { align: 'center' });
    });
    dot(s, 14.858, 3.474, BLUE, BLACK);
    dot(s, 11.24, 7.76, BLUE, BLACK);
}

/* ============================================================== slide 02 */
function slide02() {
    const s = newSlide(false);
    photo(s, 6.279, 1.851, 5.501, 7.347);
    chrome(s, '02', false, 17.984, 0.912);
    title(s, 'Table Of Content', 1.104, 2.044, 4.433, 2.121);
    arc(s, 7.286, 4.781, 5.603, 5.714, 217.78, 56.12, { rotate: 90 });
    dot(s, 11.675, 5.314, BLUE, BLACK);
    dot(s, 7.668, 9.095, BLUE, BLACK);
    mark(s, 1.252, 4.671);
    T(s, QUOTE_SHORT, 1.104, 6.069, 4.537, 0.909);
    const toc = [
        ['03', 'Introduction', 2.356, 'Lorem ipsum dolor sit amet, consectetur adipisc', 3.747, 1.916],
        ['07', 'Target Market', 2.356, 'Pretium viverra suspendisse potenti nullam ac. ', 3.747, 4.161],
        ['18', 'Human Resources Plan', 3.067, 'Ultrices tincidunt arcu non d ales neque sodales', 3.521, 6.405],
        ['24', 'Contact Information', 3.067, 'Lacinia at quis risus. Massa massa ultricies mi quis hendrerit', 3.747, 8.65]
    ];
    toc.forEach(function (r) {
        T(s, r[0], 13.982, r[5] + 0.267, 0.738, 0.572, { fontFace: HEAD, fontSize: 28, bold: true });
        T(s, r[1], 14.917, r[5], r[2], 0.37, { fontFace: HEAD });
        T(s, r[3], 14.94, r[5] + 0.446, r[4], 0.64);
    });
    pill(s, 1.104, 7.34, 3.093, 'Track Parcel', 1.335, 7.493, 2.63);
}

/* ============================================================== slide 03 */
function slide03() {
    const s = newSlide(false);
    photo(s, 11.571, 1.889, 7.325, 4.846, 8163);
    chrome(s, '03', false);
    title(s, 'Executive Overview', 13.075, 8.117, 5.64, 2.121, { align: 'right' });
    arc(s, 10.55, 2.066, 5.603, 5.714, 219.68, 50.27, { rotate: 270, flipH: true });
    dot(s, 15.439, 6.572, BLUE, BLACK);
    dot(s, 11.519, 2.632, BLUE, BLACK);
    mark(s, 11.298, 9.216, { flip: true });
    T(s, 'Our delivery service, FastTrack Deliveries, is committed by offer unparalleled speed, reliability, and convenience.', 1.104, 2.592, 7.118, 0.64, { fontFace: HEAD });
    T(s, 'Odales ut eu sem integer vitae justo eget. Pellentesque dignissim enim sit amet venenatis urna cursus eget nunc. Ultrices tincidunta rcu non sodales neque sodales ut. Urna cursus eget nunc sceleri sque. Nec dui nunc mattis enim ut tellus elementum', 1.104, 3.593, 7.118, 1.178);
    T(s, 'Vitae. Et sollicitudin ac orci phasellus egestas. Sapien nec sagittis aliquam malesuada bibendum arcu', 1.104, 5.122, 6.924, 0.64);
    bullet(s, { dx: 1.104, dy: 7.131, x: 1.546, hy: 6.817, hw: 2.356, head: 'Our Services', bold: true, by: 7.263, bw: 4.756, text: SOL_1 });
    bullet(s, { dx: 1.104, dy: 8.971, x: 1.546, hy: 8.658, hw: 3.699, head: 'Competitive Advantage', bold: true, by: 9.103, bw: 4.502, text: SOL_3 });
    pill(s, 10.412, 6.103, 3.093, 'Send Parcel', 10.644, 6.256, 2.63);
}

/* ============================================================== slide 04 */
function slide04() {
    const s = newSlide(true);
    photo(s, 3.669, 1.984, 5.701, 4.785);
    photo(s, 1.104, 4.655, 6.438, 5.614);
    chrome(s, '04', true);
    title(s, 'Problem Statement', 11.647, 7.949, 5.64, 2.121, { color: WHITE });
    arc(s, 4.341, 4.019, 5.603, 5.714, 8.25, 127.96, { rotate: 270, flipH: true, line: { color: WHITE, width: 1 } });
    mark(s, 11.927, 6.769, { ink: WHITE, accent: SAND, edge: WHITE });
    arrowCircle(s, 1.717, 2.553, WHITE);
    dot(s, 9.265, 4.984, WHITE, SAND);
    dot(s, 7.437, 9.541, WHITE, SAND);
    bullet(s, { color: WHITE, dx: 11.387, dy: 2.817, x: 11.829, hy: 2.503, hw: 2.356, head: 'Problem 01', by: 2.949, bw: 4.756, text: SOL_1 });
    bullet(s, { color: WHITE, dx: 11.387, dy: 4.657, x: 11.829, hy: 4.343, hw: 3.699, head: 'Problem 02', by: 4.789, bw: 4.502, text: SOL_3 });
}

/* ============================================================== slide 05 */
function slide05() {
    const s = newSlide(false);
    photo(s, 13.742, 2.508, 5.009, 7.364);
    chrome(s, '05', false);
    title(s, [{ text: 'Solution', options: { breakLine: true } }, { text: 'Statement' }], 1.104, 2.458, 5.64, 2.121);
    arc(s, 12.587, 1.322, 5.603, 5.714, 126.93, 324.29, { rotate: 270, flipH: true });
    dot(s, 17.565, 2.419, BLUE, BLACK);
    dot(s, 13.664, 6.371, BLUE, BLACK);
    mark(s, 5.658, 2.643);
    [
        { dx: 1.104, dy: 5.995, x: 1.624, hy: 5.734, hw: 2.356, head: 'Solution 01', by: 6.18, bw: 4.756, text: SOL_1 },
        { dx: 1.104, dy: 7.835, x: 1.624, hy: 7.574, hw: 3.699, head: 'Solution 03', by: 8.02, bw: 4.502, text: SOL_3 },
        { dx: 7.307, dy: 5.995, x: 7.749, hy: 5.682, hw: 2.356, head: 'Solution 02', by: 6.127, bw: 4.756, text: SOL_2 },
        { dx: 7.307, dy: 7.835, x: 7.749, hy: 7.522, hw: 3.699, head: 'Solution 04', by: 7.967, bw: 4.502, text: SOL_4 }
    ].forEach(function (b) { bullet(s, b); });
}

/* ============================================================== slide 06 */
function slide06() {
    const s = newSlide(false);
    chrome(s, '06', false);
    title(s, 'Market Analysis', 13.397, 1.755, 4.98, 2.121);
    mark(s, 13.397, 4.184);
    T(s, 'Vitae. Et sollicitudin ac orci phasellus egestas. Sapien nec sagittis aliquam malesuada', 13.397, 5.953, 5.117, 0.64);
    pill(s, 13.397, 7.328, 3.687, 'Socialgraphic', 13.729, 7.481, 3.022);
    pill(s, 13.397, 8.67, 3.687, 'Behavioral', 13.925, 8.824, 2.63);

    const cats = ['Product 1', 'Product 2', 'Product 3', 'Product 4'];
    s.addChart(pres.ChartType.line, [
        { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
        { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] }
    ], {
        x: 1.209, y: 3.19, w: 11.629, h: 7.342,
        chartColors: ['222222', BLUE, BLUE],
        lineSize: 2.25, lineDataSymbol: 'none', lineSmooth: false,
        showLegend: false, showTitle: false,
        valAxisMaxVal: 6, valAxisMinVal: 0, valAxisMajorUnit: 1,
        valAxisLineShow: false, catAxisLineShow: true, catAxisLineColor: AXIS,
        valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
        catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
        catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
        catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
        catAxisLabelColor: BLACK, valAxisLabelColor: BLACK
    });

    [['2,4M', 2.543, 1.827, 1.236], ['3,5M', 5.256, 1.807, 1.236], ['4,2M', 7.969, 1.807, 1.236]].forEach(function (v) {
        T(s, v[0], v[1], v[2], v[3], 0.572, { fontFace: HEAD, fontSize: 28, align: 'center' });
    });
    [['Product 1', 2.31, 2.457], ['Product 2', 5.023, 2.437], ['Product 3', 7.736, 2.437]].forEach(function (v) {
        T(s, v[0], v[1], v[2], 1.701, 0.37, { align: 'center' });
    });
}

/* ============================================================== slide 07 */
function slide07() {
    const s = newSlide(false);
    photo(s, 5.024, 1.818, 4.595, 5.676);
    photo(s, 2.508, 4.118, 4.454, 5.965);
    chrome(s, '07', false);
    arc(s, 1.292, 2.179, 5.263, 5.368, 155.11, 327.54, { rotate: 270, flipH: true });
    title(s, 'Target Market', 11.209, 1.957, 4.433, 2.121);
    mark(s, 15.924, 2.657);
    T(s, QUOTE_SHORT, 11.344, 4.546, 4.537, 0.909);
    pill(s, 11.392, 6.071, 3.093, 'Track Parcel', 11.623, 6.225, 2.63);
    bullet(s, { dx: 8.993, dy: 8.597, x: 9.434, hy: 8.284, hw: 2.356, head: 'Socialgraphic', by: 8.729, bw: 4.454, text: 'Eget dolor morbi non arcu risus quis varius quam. Non curabitur' });
    bullet(s, { dx: 14.38, dy: 8.597, x: 14.821, hy: 8.284, hw: 3.699, head: 'Behavioral', by: 8.729, bw: 4.14, text: 'arcu ac tortor. Sapien nec sagittaliqu am malesuada bibendum' });
    dot(s, 2.446, 6.974, BLUE, BLACK);
    dot(s, 4.918, 2.346, BLUE, BLACK);
}

/* ============================================================== slide 08 */
function slide08() {
    const s = newSlide(true);
    photo(s, 1.104, 1.79, 16.309, 4.315);
    chrome(s, '08', true);
    title(s, 'Competitive Analysis', 1.104, 6.639, 5.64, 2.121, { color: WHITE });
    arc(s, 12.722, 1.703, 5.603, 5.714, 303.61, 138.76, { rotate: 270, flipH: true, line: { color: WHITE, width: 1 } });
    mark(s, 1.104, 9.354, { ink: WHITE, accent: SAND, edge: WHITE });
    dot(s, 13.035, 6.0, WHITE, SAND);
    dot(s, 17.308, 2.326, WHITE, SAND);
    [
        { dx: 7.893, dy: 7.362, x: 8.26, hy: 6.991, hw: 2.899, head: 'Benchmarking', by: 7.547, bw: 4.657, text: 'Ultrices in iaculis nunc sed. Eleifend mi in nulla posuere sollicitudin.' },
        { dx: 7.893, dy: 9.309, x: 8.26, hy: 8.905, hw: 3.687, head: 'Brand Analysis', by: 9.46, bw: 4.657, text: 'Aliquam eleifend mi in nulla posuere sollicitudin aliquam. ' },
        { dx: 13.074, dy: 9.388, x: 13.441, hy: 8.984, hw: 4.184, head: 'Emerging Competition', by: 9.54, bw: 4.657, text: 'Orci nulla pellentesque dignissim enim sit amet venenatis urna. ' }
    ].forEach(function (b) { b.color = WHITE; b.headSize = 18; bullet(s, b); });
}

/* ============================================================== slide 09 */
function slide09() {
    const s = newSlide(false);
    photo(s, 2.724, 6.371, 7.938, 3.98);
    chrome(s, '09', false);
    title(s, 'SWOT Analysis', 11.936, 6.438, 5.64, 2.121);
    arc(s, 1.943, 4.704, 5.603, 5.714, 114.41, 314.31, { rotate: 270, flipH: true });
    dot(s, 7.187, 6.277, BLUE, BLACK);
    dot(s, 2.619, 9.369, BLUE, BLACK);
    mark(s, 11.936, 9.262);
    [
        { dx: 7.397, dy: 2.148, x: 7.916, hy: 1.887, hw: 2.356, head: 'Solution 01', by: 2.333, bw: 4.756, text: SOL_1 },
        { dx: 7.397, dy: 3.988, x: 7.916, hy: 3.727, hw: 3.699, head: 'Solution 03', by: 4.173, bw: 4.502, text: SOL_3 },
        { dx: 13.6, dy: 2.148, x: 14.042, hy: 1.835, hw: 2.356, head: 'Solution 02', by: 2.28, bw: 4.756, text: SOL_2 },
        { dx: 13.6, dy: 3.988, x: 14.042, hy: 3.675, hw: 3.699, head: 'Solution 04', by: 4.12, bw: 4.502, text: SOL_4 }
    ].forEach(function (b) { bullet(s, b); });
    pill(s, 1.45, 5.993, 3.093, 'Track Parcel', 1.681, 6.147, 2.63);
}

/* ============================================================== slide 10 */
function slide10() {
    const s = newSlide(false);
    photo(s, 10.247, 2.352, 7.325, 5.603);
    chrome(s, '10', false);
    title(s, 'Business Model', 4.318, 2.645, 5.64, 2.121, { align: 'right' });
    arc(s, 13.237, 1.861, 5.603, 5.714, 32.11, 212.17, { rotate: 270, flipH: true });
    dot(s, 14.452, 2.247, BLUE, BLACK);
    dot(s, 17.466, 6.984, BLUE, BLACK);
    mark(s, 3.905, 3.865, { flip: true });
    T(s, 'Odales ut eu sem integer vitae justo eget. Pellentesque dignissim enim sit amet venenatis urna cursus eget nunc. Ultrices tincidunta rcu non sodales neque sodales ut. Urna cursus eget nunc sceleri sque. ', 1.182, 5.986, 8.229, 0.909);
    T(s, LOREM_B, 1.182, 7.305, 7.518, 0.64, { bold: true });
    pill(s, 1.264, 8.867, 3.093, 'Send Parcel', 1.495, 9.021, 2.63);
    dot(s, 10.31, 9.098, BLUE);
    T(s, 'Revenue Streams', 10.758, 8.833, 4.2, 0.37, { fontFace: HEAD, bold: true });
    T(s, 'Lorem ipsum dolor sit amet', 10.758, 9.268, 4.66, 0.37);
    dot(s, 14.562, 9.109, BLUE);
    T(s, 'Pricing Strategy', 14.899, 8.821, 4.2, 0.37, { fontFace: HEAD, bold: true });
    T(s, 'accumsan. Consectetur', 14.899, 9.226, 4.66, 0.37);
}

/* ============================================================== slide 11 */
function slide11() {
    const s = newSlide(false);
    photo(s, 10.138, 5.019, 7.092, 4.521);
    photo(s, 10.843, 1.996, 7.065, 3.293);
    chrome(s, '11', false);
    title(s, 'Product or Service Offering', 1.069, 2.239, 7.268, 2.121);
    arc(s, 13.142, 2.962, 5.603, 5.714, 26.92, 137.46, { rotate: 270, flipH: true });
    mark(s, 6.405, 2.26);
    pill(s, 9.191, 8.312, 3.093, 'Send Parcel', 9.422, 8.466, 2.63);
    dot(s, 17.802, 3.643, BLUE, BLACK);
    dot(s, 17.125, 8.207, BLUE, BLACK);
    const cols = [
        { x: 1.069, head: 'Basic Product', hw: 3.252, body: 'Lorem ipsum dolor sit amet, consectetur adipisci ng elit, sed do eiusmod tempor', by: 6.21, price: '$50', px: 1.069, pw: 2.186, py: 7.475, foot: 'Eu sem integer vitae', fx: 1.041, fw: 3.569 },
        { x: 5.396, head: 'Premium Product', hw: 3.564, body: 'Est placerat in egestas erat. Nec ullamcorper sit amet risus. Eu lobortis elementum', by: 6.196, price: '$100', px: 5.43, pw: 1.581, py: 7.456, foot: 'justo eget magna Id', fx: 5.402, fw: 3.45 }
    ];
    cols.forEach(function (c) {
        T(s, c.head, c.x, 5.705, c.hw, 0.37, { fontFace: HEAD });
        T(s, c.body, c.x === 1.069 ? 1.069 : 5.394, c.by, 3.986, 0.909);
        T(s, c.price, c.px, c.py, c.pw, 0.572, { fontFace: HEAD, fontSize: 28 });
        T(s, c.foot, c.fx, 8.026, c.fw, 0.37);
    });
}

/* ============================================================== slide 12 */
function slide12() {
    const s = newSlide(true);
    chrome(s, '12', true);
    title(s, 'Marketing Strategy', 1.104, 1.764, 5.64, 2.121, { color: WHITE });
    mark(s, 5.842, 2.858, { ink: WHITE, accent: SAND, edge: WHITE });
    [
        { dy: 4.781, hy: 4.468, hw: 2.356, head: 'Product 01', by: 4.913, bw: 4.204, text: 'Laculis eu non. Quam adipisng vitae' },
        { dy: 6.229, hy: 5.915, hw: 3.699, head: 'Product 02', by: 6.361, bw: 4.502, text: 'Rhoncus. Purus non enim praeent' },
        { dy: 7.676, hy: 7.363, hw: 2.356, head: 'Product 01', by: 7.808, bw: 4.756, text: 'Justo eget magna fermentum', headFont: BODY, bodyFont: HEAD },
        { dy: 9.124, hy: 8.81, hw: 3.699, head: 'Product 02', by: 9.256, bw: 4.502, text: 'Proin sagittis nisl rhoncus mat' }
    ].forEach(function (b) {
        b.dx = 1.197; b.x = 1.638; b.bh = 0.37; b.bold = true; b.color = WHITE;
        bullet(s, b);
    });

    const cats = ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4'];
    s.addChart(pres.ChartType.bar, [
        { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
        { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] }
    ], {
        x: 9.148, y: 2.206, w: 9.748, h: 8.204,
        barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
        chartColors: ['E5DDD2', BLACK, SAND],
        showLegend: false, showTitle: false,
        valAxisMaxVal: 6, valAxisMinVal: 0, valAxisMajorUnit: 1,
        valAxisLineShow: false, catAxisLineShow: true, catAxisLineColor: AXIS,
        valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
        catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
        catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
        catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
        catAxisLabelColor: WHITE, valAxisLabelColor: WHITE
    });
}

/* ============================================================== slide 13 */
function slide13() {
    const s = newSlide(false);
    photo(s, 8.992, 5.314, 8.294, 4.212);
    chrome(s, '13', false);
    title(s, 'Operational Workflow', 12.084, 2.068, 5.64, 2.121, { align: 'right' });
    mark(s, 10.0, 3.204, { flip: true });
    [
        { dx: 1.243, dy: 2.848, hx: 1.713, hy: 2.478, head: 'Strategic', bx: 1.731, by: 2.924, text: 'Lorem ipsum dolor sit amet, conse ctetur adipisci ng elit' },
        { dx: 1.214, dy: 5.514, hx: 1.713, hy: 5.225, head: 'Budget', bx: 1.731, by: 5.764, text: 'Fermentum odio eu feugiat pretium nibh ipsum consequat nisl' },
        { dx: 1.161, dy: 8.181, hx: 1.731, hy: 7.822, head: 'Responsibility', bx: 1.731, by: 8.26, text: 'ipsum nunc aliquet bibendum enima facili sis gravida. Velit ege' }
    ].forEach(function (r) {
        dot(s, r.dx, r.dy, BLUE);
        T(s, r.head, r.hx, r.hy, 3.564, 0.37, { fontFace: HEAD });
        T(s, r.text, r.bx, r.by, 3.986, 0.64);
    });
    arc(s, 13.183, 4.669, 5.603, 5.714, 315.26, 217.03, { rotate: 270, flipH: true });
    dot(s, 14.224, 5.148, BLUE, BLACK);
    dot(s, 13.921, 9.421, BLUE, BLACK);
}

/* ============================================================== slide 14 */
function slide14() {
    const s = newSlide(false);
    photo(s, 10.0, 3.101, 4.072, 6.612);
    photo(s, 14.583, 1.94, 4.072, 6.612);
    chrome(s, '14', false);
    title(s, 'Technology Plan', 1.104, 7.169, 5.64, 2.121);
    arc(s, 11.785, 4.58, 5.603, 5.714, 324.33, 68.64, { rotate: 270, flipH: true });
    mark(s, 4.094, 8.637);
    T(s, 'The driver app is a crucial component for the efficiency and effectiveness of a delivery service. ', 1.104, 2.773, 7.474, 0.64, { fontFace: HEAD });
    T(s, LOREM_A, 1.104, 3.565, 7.118, 0.909);
    T(s, LOREM_B, 1.104, 4.692, 6.924, 0.64);
    dot(s, 17.077, 8.427, BLUE, BLACK);
    dot(s, 12.841, 9.609, BLUE, BLACK);
}

/* ============================================================== slide 15 */
function slide15() {
    const s = newSlide(false);
    photo(s, 7.126, 3.011, 4.55, 7.347);
    chrome(s, '115', false);
    title(s, 'Product Development Plan', 12.8, 5.434, 6.584, 3.13);
    mark(s, 12.89, 8.952);
    arc(s, 6.25, 1.833, 5.603, 5.714, 43.02, 234.12, { rotate: 90 });
    dot(s, 7.02, 6.684, BLUE, BLACK);
    dot(s, 11.231, 2.951, BLUE, BLACK);
    [
        { dy: 3.011, hy: 2.706, hw: 2.286, head: 'Roadsmap', by: 3.161, bw: 4.327, text: 'Lorem ipsum dolor sit amet, consec tetur adipisci ng' },
        { dy: 5.598, hy: 5.279, hw: 2.286, head: 'Milestones', by: 5.769, bw: 4.161, text: 'Sed faucibus turpis in eu mi bibendum nequ e egestas. ' },
        { dy: 8.052, hy: 7.787, hw: 2.284, head: 'Development', by: 8.281, bw: 4.035, text: 'Neque. Bibendum arcu vitae elemen tum curabitur vitae nunc. ' }
    ].forEach(function (r) {
        dot(s, 1.12, r.dy, BLUE);
        T(s, r.head, 1.493, r.hy, r.hw, 0.37, { fontFace: HEAD });
        T(s, r.text, 1.493, r.by, r.bw, 0.64);
    });
    pill(s, 12.875, 3.964, 3.093, 'Send Parcel', 13.106, 4.118, 2.63);
    T(s, '362+', 12.979, 2.011, 1.236, 0.572, { fontFace: HEAD, fontSize: 28, align: 'center' });
    T(s, '98%', 15.157, 2.011, 1.099, 0.572, { fontFace: HEAD, fontSize: 28, align: 'center' });
    T(s, 'Customers', 12.747, 2.64, 1.701, 0.37, { align: 'center' });
    T(s, 'Reviews', 15.084, 2.64, 1.244, 0.37, { align: 'center' });
}

/* ============================================================== slide 16 */
function slide16() {
    const s = newSlide(true);
    photo(s, 2.435, 2.013, 7.128, 7.473);
    chrome(s, '04', true);
    title(s, 'Lets Meet The Founders', 10.438, 2.074, 6.52, 2.121, { color: WHITE });
    arc(s, 1.37, 4.635, 5.603, 5.714, 217.99, 45.63, { rotate: 270, flipH: true, line: { color: WHITE, width: 1 } });
    mark(s, 10.559, 4.42, { ink: WHITE, accent: SAND, edge: WHITE });
    dot(s, 2.329, 5.174, WHITE, SAND);
    dot(s, 6.086, 9.385, WHITE, SAND);
    T(s, 'Amara Kensington', 10.438, 5.735, 3.564, 0.438, { fontFace: HEAD, fontSize: 20, bold: true, color: WHITE });
    T(s, 'Executive Officer (CEO)', 10.438, 6.138, 4.095, 0.37, { color: WHITE });
    T(s, '+23 Years', 10.438, 6.876, 3.564, 0.438, { fontFace: HEAD, fontSize: 20, bold: true, color: WHITE });
    T(s, 'Work at company', 10.438, 7.279, 4.095, 0.37, { color: WHITE });
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipisci ng elit, sed do eiusmod tempor incididunt ut labo re et dolore magna aliqua. ', 10.438, 8.118, 6.143, 0.909, { color: WHITE });
}

/* ============================================================== slide 17 */
function slide17() {
    const s = newSlide(false);
    [1.261, 7.359, 13.458].forEach(function (x) { photo(s, x, 4.502, 5.476, 3.445, 6136); });
    chrome(s, '03', false);
    title(s, 'Key Team Members', 1.451, 1.706, 7.299, 2.121);
    arc(s, 10.345, 3.475, 5.603, 5.714, 130.31, 229.65, { rotate: 270, flipH: true });
    mark(s, 6.337, 2.766);
    dot(s, 10.874, 4.397, BLUE, BLACK);
    dot(s, 15.189, 4.409, BLUE, BLACK);
    [
        { nx: 2.217, rx: 1.951, dx: 2.267, dw: 3.464, name: 'Kai Ellington', role: 'Financial Officer (CFO)', note: 'Ut porttitor leo a diam sollici tudin tempor id sed' },
        { nx: 8.315, rx: 8.05, dx: 8.436, dw: 3.322, name: 'Seraphina Blake', role: 'Technology Officer (CTO)', note: 'Laoreet non curabitur gravida arcu ac tortor dignissim' },
        { nx: 14.413, rx: 14.148, dx: 14.534, dw: 3.322, name: 'Rion Bennett', role: 'Marketing Officer (CMO)', note: 'Feugiat vivamus at aug ue eget arcu dictum vari' }
    ].forEach(function (c) {
        T(s, c.name, c.nx, 8.274, 3.564, 0.438, { fontFace: HEAD, fontSize: 20, align: 'center' });
        T(s, c.role, c.rx, 8.727, 4.095, 0.37, { align: 'center' });
        T(s, c.note, c.dx, 9.346, c.dw, 0.64, { align: 'center' });
    });
}

/* ============================================================== slide 18 */
function slide18() {
    const s = newSlide(false);
    photo(s, 2.935, 2.264, 7.065, 4.013, 6402);
    photo(s, 2.231, 5.148, 4.857, 4.66);
    chrome(s, '03', false);
    title(s, 'Human Resources Plan', 10.644, 2.69, 6.704, 2.121);
    mark(s, 14.737, 2.77);
    arc(s, 1.406, 3.042, 5.603, 5.714, 44.28, 153.15, { rotate: 90 });
    dot(s, 2.83, 3.319, BLUE, BLACK);
    dot(s, 2.126, 7.854, BLUE, BLACK);
    pill(s, 5.623, 8.661, 3.093, 'Send Parcel', 5.855, 8.814, 2.63);
    [
        { dx: 10.738, dy: 6.407, hx: 11.007, hy: 6.057, hw: 2.63, head: 'Demand Forecasting', bx: 11.021, by: 6.539, bw: 3.7, text: 'Orci nulla pellentesque digsim enim sit amet venenatis' },
        { dx: 15.227, dy: 6.393, hx: 15.439, hy: 6.057, hw: 2.427, head: 'Supply Forecasting', bx: 15.439, by: 6.539, bw: 3.7, text: 'Ultrices in iaculis nunc sed. Eleifend mi in nulla' },
        { dx: 10.738, dy: 8.404, hx: 11.055, hy: 8.046, hw: 2.63, head: 'Determining HRR', bx: 11.055, by: 8.536, bw: 3.224, text: 'proin fermentum leo. In nibh mauris cursus mattis' },
        { dx: 15.227, dy: 8.404, hx: 15.547, hy: 8.028, hw: 2.63, head: 'Action Planning', bx: 15.547, by: 8.518, bw: 2.908, text: 'sed odio morbi quis. Vulp utate eu scelerisque' }
    ].forEach(function (r) {
        dot(s, r.dx, r.dy, BLUE);
        T(s, r.head, r.hx, r.hy, r.hw, 0.37, { fontFace: HEAD });
        T(s, r.text, r.bx, r.by, r.bw, 0.64);
    });
}

/* ============================================================== slide 19 */
function slide19() {
    const s = newSlide(false);
    chrome(s, '03', false);
    title(s, 'Funding Requirements', 11.777, 2.057, 7.325, 2.121);
    mark(s, 11.913, 4.461);

    const cats = ['Service 1', 'Service 2', 'Service 3', 'Service 4'];
    s.addChart(pres.ChartType.bar, [
        { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
        { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] }
    ], {
        x: 1.104, y: 1.453, w: 10.045, h: 8.87,
        barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 150, barOverlapPct: 100,
        chartColors: ['5C8DFA', BLACK, 'D0CFC2'],
        showLegend: false, showTitle: false,
        valAxisLineShow: false, catAxisLineShow: false,
        valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
        catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
        catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
        catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
        catAxisLabelColor: BLACK, valAxisLabelColor: BLACK
    });

    T(s, 'Dignissim diam quis enim lobortis scelerisque fermentum dui. Praesent semper feugiat nibh sed pulvinar proin gravida hrerit lectus. Vitae auctor eu augue ut lectus arcu', 11.777, 5.858, 7.102, 0.909);
    T(s, 'eget magna fermentum iaculis eu. Velit ut tortor pretium viverra suspendisse. Amet nisl purus in mollis ', 11.777, 6.995, 7.102, 0.64);
    T(s, '99%', 11.794, 8.035, 2.501, 0.572, { fontFace: HEAD, fontSize: 28 });
    T(s, 'High quality service', 11.794, 8.591, 2.873, 0.37);
}

/* ============================================================== slide 20 */
function slide20() {
    const s = newSlide(true);
    photo(s, 1.254, 5.625, 8.746, 4.59);
    chrome(s, '04', true);
    title(s, 'Financial Plan', 1.314, 2.063, 4.919, 2.121, { color: WHITE });
    arc(s, 5.682, 4.557, 5.603, 5.714, 32.37, 231.86, { rotate: 270, flipH: true, line: { color: WHITE, width: 1 } });
    mark(s, 1.45, 4.267, { ink: WHITE, accent: SAND, edge: WHITE });
    dot(s, 6.17, 5.52, WHITE, SAND);
    dot(s, 9.895, 9.659, WHITE, SAND);
    [
        { dy: 2.562, hy: 2.248, hw: 2.356, head: 'Standard', by: 2.693, bw: 4.756, text: 'Lorem ipsum dolor sit amet consectetur adi piscing elit pellentesque. At risus viverra adi piscing at in tellus integer feugiat' },
        { dy: 5.375, hy: 5.062, hw: 3.699, head: 'Basic', by: 5.507, bw: 4.502, text: 'scelerisque. Arcu cursus vitae congue ma uris rhoncus. Maecenas pharetra conval lis posuere morbi leo urna' },
        { dy: 8.189, hy: 7.875, hw: 3.699, head: 'Premium', by: 8.321, bw: 4.502, text: 'iaculis nunc sed augue lacus viverra vitae. Felis donec et odio pellentesque diam vo  utpat commodo. Fringilla' }
    ].forEach(function (b) {
        b.dx = 12.869; b.x = 13.311; b.bh = 0.909; b.color = WHITE;
        bullet(s, b);
    });
    [['$10', 6.458, 1.53], ['$50', 8.576, 1.236], ['$100', 10.543, 1.244]].forEach(function (v) {
        T(s, v[0], v[1], 2.248, v[2], 0.572, { fontFace: HEAD, fontSize: 28, align: 'center', color: WHITE });
    });
    [['Standard', 6.189, 2.069], ['Basic', 8.343, 1.701], ['Premium', 10.543, 1.244]].forEach(function (v) {
        T(s, v[0], v[1], 2.878, v[2], 0.37, { align: 'center', color: WHITE });
    });
}

/* ============================================================== slide 21 */
function slide21() {
    const s = newSlide(false);
    photo(s, 2.489, 6.66, 16.309, 3.812);
    chrome(s, '03', false);
    title(s, 'Company Timeline', 1.082, 1.524, 5.64, 2.121);
    mark(s, 1.236, 3.839);
    arc(s, 1.601, 5.193, 5.603, 5.714, 42.62, 240.65, { rotate: 90 });
    dot(s, 6.77, 6.553, BLUE, BLACK);
    dot(s, 2.384, 10.005, BLUE, BLACK);
    [
        { dx: 5.762, dy: 3.847, x: 6.059, hy: 3.569, by: 3.964, bw: 3.521, head: '2014-2017', text: 'Felis imperdiet proin ferment um leo vel orci. Eleme' },
        { dx: 8.613, dy: 2.106, x: 8.989, hy: 1.841, by: 2.236, bw: 3.298, head: '2017- 2020', text: 'Lorem ipsum dolor sit amet, consectetur adipiscing' },
        { dx: 11.859, dy: 3.847, x: 12.22, hy: 3.569, by: 3.964, bw: 3.521, head: '2020-2022', text: 'platea dictumst quisque elle ntes que divitae turpis' },
        { dx: 14.924, dy: 2.095, x: 15.238, hy: 1.841, by: 2.236, bw: 3.298, head: '2022- 2030', text: 'Odales ut eu sem integer vitae justo eget. ' }
    ].forEach(function (r) {
        dot(s, r.dx, r.dy, BLUE);
        T(s, r.head, r.x, r.hy, 3.298, 0.37, { fontFace: HEAD });
        T(s, r.text, r.x, r.by, r.bw, 0.64);
    });
}

/* ============================================================== slide 22 */
function slide22() {
    const s = newSlide(false);
    photo(s, 11.571, 4.476, 7.325, 4.846);
    chrome(s, '03', false);
    title(s, 'Strategic Partnerships', 12.913, 1.898, 5.64, 2.121, { align: 'right' });
    mark(s, 11.571, 1.981, { flip: true });
    arc(s, 10.55, 4.653, 5.603, 5.714, 219.68, 50.27, { rotate: 270, flipH: true });
    dot(s, 15.439, 9.159, BLUE, BLACK);
    dot(s, 11.519, 5.22, BLUE, BLACK);
    pill(s, 10.412, 8.69, 3.093, 'Track Parcel', 10.644, 8.843, 2.63);
    bullet(s, { dx: 1.153, dy: 2.958, x: 1.588, hy: 2.65, hw: 2.877, head: 'Brand Vision', by: 3.168, bw: 5.639, text: 'Lorem ipsum dolor sit amet, consectetur adipisci ng elit, sed do eiusmod tempor' });
    bullet(s, { dx: 1.153, dy: 4.678, x: 1.588, hy: 4.317, hw: 2.877, head: 'Business Goals', by: 4.855, bw: 5.639, text: 'Id eu nisl nunc mi ipsum. Vitae elementum curabitur vitae nunc. Commodo' });
    T(s, 'Strategic partnerships businesses or organizations that aim to achieve mutually beneficial goals. ', 1.104, 6.6, 7.118, 0.64, { fontFace: HEAD });
    T(s, LOREM_A, 1.104, 7.393, 7.118, 0.909);
    T(s, LOREM_B, 1.104, 8.519, 6.924, 0.64);
}

/* ============================================================== slide 23 */
function slide23() {
    const s = newSlide(false);
    photo(s, 1.928, 4.866, 4.972, 5.342, 6402);
    photo(s, 5.565, 1.88, 3.92, 4.55);
    chrome(s, '03', false);
    title(s, 'Customer Testimonials', 10.01, 2.098, 5.64, 2.121);
    mark(s, 10.257, 4.461);
    bullet(s, { dx: 10.257, dy: 7.217, x: 10.699, hy: 6.903, hw: 2.356, head: 'Arabella Hayes', by: 7.349, bw: 4.756, text: SOL_1 });
    bullet(s, { dx: 10.257, dy: 9.057, x: 10.699, hy: 8.744, hw: 3.699, head: 'Theodore Sterling', by: 9.189, bw: 4.502, text: SOL_3 });
    arrowCircle(s, 7.612, 7.688, BLACK);
    arc(s, 1.104, 3.018, 5.603, 5.714, 136.44, 305.73);
    dot(s, 5.46, 3.474, BLUE, BLACK);
    dot(s, 1.842, 7.76, BLUE, BLACK);
}

/* ============================================================== slide 24 */
function slide24() {
    const s = newSlide(true);
    photo(s, 10.607, 2.576, 6.773, 5.892);
    chrome(s, '04', true);
    title(s, 'Contact Information', 1.159, 2.576, 5.64, 2.121, { color: WHITE });
    mark(s, 1.24, 5.072, { ink: WHITE, accent: SAND, edge: WHITE });
    arc(s, 13.172, 3.593, 5.603, 5.714, 314.97, 150.69, { rotate: 270, flipH: true, line: { color: WHITE, width: 1 } });
    dot(s, 17.274, 3.864, WHITE, SAND);
    dot(s, 13.858, 8.341, WHITE, SAND);
    [
        { dx: 0.939, dy: 7.21, x: 1.314, hy: 6.833, vy: 7.421, head: 'Phone', text: '(+12) 3456-7890-1234' },
        { dx: 0.94, dy: 8.784, x: 1.314, hy: 8.407, vy: 8.995, head: 'Location', text: '123 Main Street Your City' },
        { dx: 5.746, dy: 7.21, x: 6.051, hy: 6.914, vy: 7.501, head: 'Mail', text: 'quickcourier@gmail.com' },
        { dx: 5.748, dy: 8.784, x: 6.051, hy: 8.407, vy: 8.995, head: 'Website', text: 'www.quickcourier.com' }
    ].forEach(function (r) {
        dot(s, r.dx, r.dy, WHITE);
        T(s, r.head, r.x, r.hy, 3.697, 0.438, { fontFace: HEAD, fontSize: 20, color: WHITE });
        T(s, r.text, r.x, r.vy, 4.167, 0.37, { color: WHITE });
    });
}

/* ============================================================== slide 25 */
function slide25() {
    const s = newSlide(false);
    photo(s, 1.104, 1.981, 17.792, 4.315);
    chrome(s, '2030', false);
    arc(s, 1.544, 2.079, 5.603, 5.714, 298.42, 61.64, { rotate: 270, flipH: true });
    title(s, 'Thank You', 9.864, 7.264, 8.283, 1.582, { fontSize: 88, align: 'right' });
    mark(s, 8.53, 7.629, { flip: true });
    T(s, QUOTE_LONG, 11.039, 8.895, 7.108, 0.909, { align: 'right' });
    dot(s, 1.747, 6.191, BLUE, BLACK);
    dot(s, 6.767, 6.191, BLUE, BLACK);
}

/* ------------------------------------------------------------------ main */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
 slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
 slide19, slide20, slide21, slide22, slide23, slide24, slide25].forEach(function (fn) { fn(); });

pres.writeFile({ fileName: path.join(__dirname, '044b90e2-29be-4c90-8e80-1ca7746fcba1_grok_final.pptx') })
    .then(function (f) { console.log('wrote ' + f); })
    .catch(function (e) { console.error(e); process.exit(1); });
