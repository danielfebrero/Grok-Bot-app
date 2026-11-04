/**
 * "Exploring the World of Interior Design" — 20-slide deck rebuilt with pptxgenjs.
 *
 * Run:  node 0cec86d1-2d4e-4202-baa1-6af0b5c7bb81_grok_final.js
 * Out:  0cec86d1-2d4e-4202-baa1-6af0b5c7bb81_grok_final.pptx (next to this file)
 *
 * The source deck contains line-art SVG icons and empty picture placeholders.
 * Icons are redrawn as outlined placeholder boxes at their original geometry;
 * the picture placeholders carry no picture in the source and render as nothing,
 * so they are not reproduced.
 */

'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
    sage: '808275', // accent1 - primary olive
    ink: '3F3D42', // accent2 - near-black text
    sand: 'CBC1B5', // accent3
    cream: 'EEE9E6', // accent4
    mist: 'E0E2DE', // accent6 - pale sage panels
    teal: 'BAC9C4',
    grey: '7F7F7F',
    white: 'FFFFFF',
    black: '000000',
};

const HEAD = 'Figtree'; // theme major font
const BODY = 'Inter'; // theme minor font

const SLIDE_W = 26.6666667;
const SLIDE_H = 15;

/* ---------------------------------------------------------------- helpers */

const NOLINE = { type: 'none' };

/** Flat filled rectangle. */
function rect(s, x, y, w, h, color) {
    s.addShape('rect', { x, y, w, h, fill: { color }, line: NOLINE });
}

/** Filled circle / oval. */
function oval(s, x, y, w, h, color) {
    s.addShape('ellipse', { x, y, w, h, fill: { color }, line: NOLINE });
}

/** "Rectangle: Top Corners Rounded" — a half-round arch, drawn as custom geometry. */
function arch(s, x, y, w, h, color) {
    const r = w / 2;
    s.addShape('custGeom', {
        x, y, w, h,
        fill: { color },
        line: NOLINE,
        points: [
            { x: 0, y: h, moveTo: true },
            { x: 0, y: r },
            { x: w, y: r, curve: { type: 'arc', hR: r, wR: r, stAng: 180, swAng: 180 } },
            { x: w, y: h },
            { close: true },
        ],
    });
}

/**
 * Free-form petal / leaf shapes (slides 12 and 16).
 * `d` is an SVG-like path in the original design-unit grid (pw x ph);
 * 'M' = moveTo, 'C' = cubic bezier, 'Z' = close.
 */
function freeform(s, x, y, w, h, pw, ph, d, color) {
    const sx = (v) => (v / pw) * w;
    const sy = (v) => (v / ph) * h;
    const pts = [];
    for (let i = 0; i < d.length;) {
        const op = d[i++];
        if (op === 'M') pts.push({ x: sx(d[i++]), y: sy(d[i++]), moveTo: true });
        else if (op === 'C') {
            const [x1, y1, x2, y2, ex, ey] = d.slice(i, i + 6); i += 6;
            pts.push({ x: sx(ex), y: sy(ey), curve: { type: 'cubic', x1: sx(x1), y1: sy(y1), x2: sx(x2), y2: sy(y2) } });
        } else if (op === 'Z') pts.push({ close: true });
    }
    s.addShape('custGeom', { x, y, w, h, fill: { color }, line: NOLINE, points: pts });
}

/** Text box. Defaults mirror the deck: Inter 26pt, top aligned, wrapping. */
function txt(s, text, x, y, w, h, o) {
    s.addText(text, Object.assign({
        x, y, w, h,
        fontFace: BODY, fontSize: 26, color: C.black,
        align: 'left', valign: 'top', wrap: true,
    }, o));
}

/** Stand-in for one of the deck's line-art icons: a thin outlined frame at its original box. */
function icon(s, x, y, w, h, color) {
    s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.14, fill: NOLINE, line: { color, width: 1 } });
}

/** Top-right wordmark: "Inter" + shower-head glyph. */
function logo(s, x, y, color) {
    txt(s, 'Inter    o', x, y + 0.035, 2.13, 0.525,
        { fontFace: HEAD, fontSize: 28, bold: true, color, charSpacing: 3, lineSpacingMultiple: 0.9 });
    icon(s, x + 1.052, y, 0.751, 0.751, color);
}

/** Top-left "Website.com" tag (present on every slide). */
function siteTag(s, color) {
    txt(s, 'Website.com', 1.21, 0.967, 2.235, 0.404, { fontSize: 18, bold: true, color });
}

/** Small caps kicker above a headline. */
function kicker(s, text, x, y, w, color, align) {
    txt(s, text, x, y, w, 0.505,
        { fontSize: 24, bold: true, color: color || C.ink, charSpacing: 3, align: align || 'left' });
}

/** Big Figtree headline. `text` may be an array of pptxgenjs runs. */
function headline(s, text, x, y, w, h, size, color, align) {
    txt(s, text, x, y, w, h,
        { fontFace: HEAD, fontSize: size, color: color || C.ink, lineSpacingMultiple: 0.9, align: align || 'left' });
}

/** Bold "Text Here" / "Tittle Here" caption. */
function caption(s, text, x, y, w, h, color, o) {
    txt(s, text, x, y, w, h || 0.539,
        Object.assign({ fontSize: 26, bold: true, color: color || C.black, charSpacing: 3 }, o));
}

/** Body paragraph at the deck's standard 120% leading. */
function para(s, text, x, y, w, h, color, o) {
    txt(s, text, x, y, w, h, Object.assign({ color: color || C.black, lineSpacingMultiple: 1.2 }, o));
}

/** Caption + paragraph pair, the deck's most repeated block. */
function note(s, n) {
    caption(s, n.label || 'Text Here', n.x, n.y, n.lw || 3.797, 0.539, n.color);
    para(s, n.text, n.bx === undefined ? n.x : n.bx, n.by, n.bw, n.bh, n.color);
}

/* ------------------------------------------------------------ slide bodies */

function slide1(s) {
    rect(s, 0, 0.005, 16, 14.991, C.sage);
    arch(s, 14.677, 1.5, 6.983, 11.333, C.mist);
    logo(s, 23.413, 0.892, C.ink);
    siteTag(s, C.white);
    kicker(s, 'UNLOCKING CREATIVITY IN LIVING SPACES', 1.873, 2.922, 9.721, C.white);
    headline(s, 'Exploring the World of Interior Design', 1.873, 3.858, 10.981, 5.327, 115, C.white);
    rect(s, 2.006, 10.854, 4.002, 1.333, C.white);
    txt(s, 'Star here', 2.088, 11.235, 3.838, 0.572,
        { fontSize: 28, bold: true, color: C.sage, charSpacing: 3, align: 'center' });
}

function slide2(s) {
    rect(s, 13.78, 12.833, 12.887, 2.167, C.sage);
    rect(s, 0, 0, 5.332, 12.833, C.mist);
    kicker(s, 'BLENDING COMFORT, STYLE, AND FUNCTIONALITY', 11.934, 2.856, 10.735);
    headline(s, 'Curating the Perfect Ensemble', 11.892, 3.585, 11.618, 3.009, 96);
    logo(s, 23.413, 0.892, C.ink);
    icon(s, 15.322, 7.969, 1.044, 1.044, '231F20');
    note(s, {
        x: 17.247, y: 7.969, by: 8.698, bw: 7.509, bh: 2.157,
        text: 'Interior design is the art and science of enhancing the interior of a space to achieve a healthier and more aesthetically pleasing environment for the people using. ',
    });
    siteTag(s, C.ink);
}

function slide3(s) {
    rect(s, 0, 8.167, 7.997, 6, C.sage);
    kicker(s, 'COLORS, TEXTURES, AND SHAPES', 15.901, 2.89, 8.997);
    headline(s, 'Understanding the Building Blocks', 15.859, 3.619, 8.606, 4.1, 88);
    icon(s, 16.0, 8.804, 1.062, 1.062, C.black);
    note(s, {
        x: 17.933, y: 8.95, bx: 17.913, by: 9.786, bw: 6.985, bh: 2.682,
        text: 'By strategically combining these elements, designers can evoke various moods, enhance spatial perception, and create harmonious environments that resonate with occupants.',
    });
    icon(s, 1.987, 2.537, 1.034, 1.034, '231F20');
    para(s, 'Interior design revolves around key elements such as color schemes, textures, and shapes.',
        1.873, 4.252, 4.667, 2.157);
    icon(s, 1.953, 8.95, 1.209, 1.209, C.white);
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ',
        1.873, 10.601, 4.887, 2.157, C.white);
    logo(s, 23.413, 0.892, C.ink);
    siteTag(s, C.ink);
}

function slide4(s) {
    arch(s, 19.335, 6.167, 7.332, 8.833, 'E7E6E6'); // arch belongs to the slide layout
    rect(s, 3.995, 10.167, 17.525, 4.833, C.mist);
    kicker(s, 'BALANCE, PROPORTION, AND UNITY', 8.547, 2.246, 8.997);
    headline(s, 'Guiding Principles for Design Harmony', 8.505, 2.975, 11.912, 4.463, 96);
    logo(s, 23.413, 0.892, C.ink);
    rect(s, 3.995, 8.167, 13.337, 4.667, C.sage);
    icon(s, 4.764, 8.933, 0.943, 0.943, C.white);
    note(s, {
        x: 6.135, y: 9.014, lw: 4.866, by: 9.91, bw: 10.427, bh: 2.157, color: C.white,
        text: 'Achieving the right balance ensures visual stability, while proportion ensures the harmonious relationship between different elements. Unity ties everything together, creating a cohesive and pleasing design.',
    });
    siteTag(s, C.white);
}

function slide5(s) {
    rect(s, 18.002, -0.002, 8.665, 15, C.mist);
    logo(s, 23.413, 0.892, C.ink);
    icon(s, 19.307, 1.96, 0.901, 0.901, C.black);
    note(s, {
        x: 19.239, y: 3.439, by: 4.334, bw: 5.431, bh: 1.632, color: C.ink,
        text: 'Interior design goes beyond aesthetics; it also prioritizes functionality and ergonomics. ',
    });
    icon(s, 19.307, 8.653, 0.901, 0.901, '231F20');
    note(s, {
        x: 19.239, y: 10.107, by: 11.003, bw: 6.304, bh: 1.632, color: C.ink,
        text: 'From furniture placement to traffic flow, every aspect is meticulously planned to enhance usability.',
    });
    rect(s, 5.998, -0.002, 2.304, 13.502, C.white);
    kicker(s, 'CREATING SPACES THAT WORK', 7.877, 2.492, 7.821);
    headline(s, 'Designing for Comfort and Efficiency', 7.793, 3.221, 8.833, 4.463, 96);
    note(s, {
        x: 7.877, y: 8.91, by: 9.806, bw: 7.225, bh: 1.632, color: C.ink,
        text: 'Designers consider how people interact with spaces and aim to optimize layouts for comfort, efficiency, and ease of use. ',
    });
    siteTag(s, C.white);
}

function slide6(s) {
    rect(s, 19.356, 10.853, 7.312, 4.167, C.mist);
    kicker(s, 'FROM CLASSIC TO CONTEMPORARY', 9.945, 2.957, 9.516);
    headline(s, 'Exploring Design Diversity', 9.774, 3.686, 10.96, 3.009, 96);
    logo(s, 23.413, 0.892, C.ink);
    rect(s, 8.708, 7.508, 14.628, 4.659, C.sage);
    icon(s, 9.875, 8.347, 0.963, 0.963, C.white);
    note(s, {
        x: 11.22, y: 8.347, lw: 4.866, by: 9.003, bw: 10.951, bh: 2.157, color: C.white,
        text: 'Interior design encompasses a wide range of styles and themes, from traditional to modern, minimalist to eclectic. Each style has its own characteristics and influences, allowing designers to tailor spaces to suit various preferences',
    });
    siteTag(s, C.ink);
}

function slide7(s) {
    arch(s, 14.017, 1.371, 7.662, 13.637, C.mist);
    para(s, 'Designers use different lighting techniques, such as ambient, task, and accent lighting, to create layers of illumination that enhance aesthetics',
        1.886, 6.794, 10.947, 1.632, C.ink);
    kicker(s, 'ENHANCING ATMOSPHERE AND FUNCTIONALITY', 1.886, 2.282, 10.397);
    headline(s, 'Illuminating Spaces with Light', 1.803, 3.011, 11.218, 3.009, 96);
    logo(s, 23.413, 0.892, C.ink);
    // two identical sage cards
    const CARD_TEXT = 'Lighting plays a crucial role in interior design, affecting the mood, and ambiance';
    [{ cx: 9.853, ix: 10.602, tx: 11.906 }, { cx: 1.33, ix: 1.912, tx: 3.237 }].forEach(function (c) {
        rect(s, c.cx, 9.291, 8.002, 3.989, C.sage);
        icon(s, c.ix, 10.141, 0.751, 0.751, C.white);
        note(s, { x: c.tx, y: 10.141, lw: 4.866, by: 10.797, bw: 4.866, bh: 1.632, color: C.white, text: CARD_TEXT });
    });
    siteTag(s, C.ink);
}

function slide8(s) {
    rect(s, 9.998, 10.167, 12.003, 4.833, C.mist);
    logo(s, 23.413, 0.892, C.ink);
    kicker(s, 'BALANCING FORM AND FUNCTION', 15.23, 2.24, 7.436);
    headline(s, 'Choosing the Right Materials', 15.147, 2.969, 9.524, 3.009, 96);
    rect(s, 3.351, 2.167, 9.316, 3.873, C.sage);
    icon(s, 3.982, 2.959, 0.751, 0.751, C.white);
    note(s, {
        x: 5.216, y: 2.959, lw: 4.866, by: 3.615, bw: 6.822, bh: 1.632, color: C.white,
        text: 'From hardwood floors to eco-friendly paints, each choice reflects the desired look and feel of the interior.',
    });
    siteTag(s, C.ink);
    rect(s, 15.332, 7.5, 10.003, 4.667, C.sage);
    icon(s, 16.061, 8.266, 0.943, 0.943, C.white);
    note(s, {
        x: 17.432, y: 8.348, lw: 4.866, by: 9.243, bw: 7.175, bh: 2.157, color: C.white,
        text: 'Materials and finishes contribute to the tactile and visual experience of a space. Designers carefully select materials based on durability',
    });
}

function slide9(s) {
    arch(s, 13.997, 1.349, 6.668, 11.5, C.mist);
    kicker(s, 'Eco-Friendly Approaches to Interior', 1.889, 2.267, 8.997);
    headline(s, 'Designing for a Greener Future', 1.806, 2.996, 7.948, 4.1, 88);
    logo(s, 23.413, 0.892, C.ink);
    rect(s, 17.335, 8.842, 8.011, 4.021, C.sage);
    icon(s, 17.862, 9.548, 0.943, 0.943, C.white);
    note(s, {
        x: 19.233, y: 9.63, lw: 4.866, by: 10.525, bw: 5.627, bh: 1.632, color: C.white,
        text: 'Sustainable design principles are increasingly integrated into interior design projects. ',
    });
    icon(s, 1.99, 8.383, 1.209, 1.209, '231F20');
    note(s, {
        x: 1.889, y: 10.079, lw: 4.866, by: 10.734, bw: 6.611, bh: 2.157, color: C.ink,
        text: 'Designers prioritize energy efficiency, waste reduction, and the use of eco-friendly materials to minimize environmental impact',
    });
    siteTag(s, C.ink);
}

function slide10(s) {
    rect(s, 17.335, 8.167, 9.332, 6.833, C.mist);
    kicker(s, 'FROM CONCEPT TO COMPLETION', 13.899, 2.949, 8.997);
    headline(s, 'Navigating the Design Process', 13.816, 3.679, 9.521, 2.767, 88);
    logo(s, 23.413, 0.892, C.ink);
    siteTag(s, C.ink);
    icon(s, 19.334, 9.871, 0.935, 0.935, '231F20');
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ',
        19.209, 11.479, 5.656, 1.632);
    rect(s, 7.998, 8.167, 10.003, 4.667, C.sage);
    icon(s, 9.115, 8.933, 0.943, 0.943, C.white);
    note(s, {
        x: 10.486, y: 9.014, lw: 4.866, by: 9.91, bw: 6.399, bh: 2.157, color: C.white,
        text: 'Designers oversee everything from initial concept development and space planning to procurement and installation, ensuring',
    });
}

function slide11(s) {
    rect(s, 18.636, 0.057, 8.025, 14.999, C.white); // white gradient wash in the source
    kicker(s, 'FROM CLASSIC TO CONTEMPORARY', 1.886, 2.623, 8.997);
    headline(s, [
        { text: 'Charting the', options: { breakLine: true } },
        { text: 'Evolution' },
    ], 1.803, 3.352, 7.343, 2.767, 88);
    logo(s, 23.413, 0.892, C.ink);

    // pie legend (right edge)
    [{ y: 10.554, c: C.ink }, { y: 11.514, c: C.mist }, { y: 12.474, c: C.sage }].forEach(function (l) {
        txt(s, 'Text Here', 23.087, l.y, 1.675, 0.442,
            { fontSize: 16, color: C.grey, lineSpacingMultiple: 1.2 });
        oval(s, 22.666, l.y + 0.101, 0.281, 0.281, l.c);
    });

    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
        1.886, 6.877, 7.177, 1.632, C.ink);

    // three key figures
    [{ x: 2.007, tx: 2.547, lx: 2.565, y: 9.56, v: '2,463', vw: 2.368, c: C.ink },
    { x: 5.401, tx: 5.94, lx: 5.959, y: 9.56, v: '6,181', vw: 2.215, c: C.sage },
    { x: 2.007, tx: 2.547, lx: 2.565, y: 11.706, v: '4,463', vw: 2.368, c: C.mist }].forEach(function (k) {
        txt(s, k.v, k.tx, k.y, k.vw, 0.915,
            { fontSize: 44, bold: true, color: C.ink, lineSpacingMultiple: 1.2 });
        txt(s, 'Text Little Here', k.lx, k.y + 1.006, 2.368, 0.434,
            { fontSize: 18, bold: true, color: C.ink, lineSpacingMultiple: 1.2 });
        oval(s, k.x, k.y + 0.317, 0.281, 0.281, k.c);
    });

    // Source frame is 10.312/1.958/11.802x11.084; nudged so the rendered circle lands
    // where the original does (the source explodes each slice a few percent outward).
    s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [6.2, 4.2, 2.4] }], {
        x: 10.31, y: 1.98, w: 11.72, h: 11.03,
        chartColors: [C.sage, C.mist, C.ink],
        dataBorder: { pt: 21, color: C.white }, // stands in for the slice explosion gaps
        showLegend: false, showValue: false, showTitle: false,
        firstSliceAng: 0,
        chartArea: { fill: { color: C.white, transparency: 100 } },
    });

    // value callouts sitting on top of the pie
    [{ v: '60%', vx: 17.409, vy: 6.699, lx: 17.335, ly: 7.873, lh: 0.677, c: C.white },
    { v: '40%', vx: 12.396, vy: 8.587, lx: 12.321, ly: 9.761, lh: 0.36, c: C.black },
    { v: '20%', vx: 12.711, vy: 3.572, lx: 12.636, ly: 4.746, lh: 0.677, c: C.white }].forEach(function (p) {
        txt(s, p.v, p.vx, p.vy, 2.855, 1.1,
            { fontSize: 54, bold: true, color: p.c, align: 'center', lineSpacingMultiple: 1.2 });
        txt(s, 'Text Here', p.lx, p.ly, 3.004, p.lh,
            { fontSize: 14, bold: true, color: p.c, align: 'center', lineSpacingMultiple: 1.2 });
    });
    siteTag(s, C.ink);
}

/* Slide 12 — four interlocking petals numbered 01-04. */
const PETALS_4 = [
    { x: 13.624, y: 5.342, w: 4.18, h: 4.71, pw: 532, ph: 599, color: C.sage,
        d: ['M', 415, 511, 'C', 532, 394, 532, 204, 415, 87, 'C', 359, 31, 284, 0, 203, 0,
            'C', 127, 0, 56, 28, 0, 79, 'C', 56, 139, 86, 217, 86, 299, 'C', 86, 382, 56, 459, 0, 519,
            'C', 56, 571, 127, 599, 203, 599, 'C', 284, 599, 359, 568, 415, 511, 'Z'] },
    { x: 12.871, y: 9.565, w: 4.707, h: 3.956, pw: 599, ph: 503, color: C.grey,
        d: ['M', 299, 86, 'C', 217, 86, 139, 56, 79, 0, 'C', 28, 55, 0, 127, 0, 203,
            'C', 0, 283, 31, 359, 88, 415, 'C', 144, 472, 219, 503, 299, 503, 'C', 380, 503, 455, 472, 511, 415,
            'C', 568, 359, 599, 283, 599, 203, 'C', 599, 127, 571, 55, 520, 0, 'C', 460, 56, 382, 86, 299, 86, 'Z'] },
    { x: 9.405, y: 8.808, w: 3.953, h: 4.713, pw: 503, ph: 599, color: C.ink,
        d: ['M', 88, 511, 'C', 144, 568, 219, 599, 299, 599, 'C', 376, 599, 447, 571, 503, 519,
            'C', 447, 459, 416, 382, 416, 299, 'C', 416, 217, 447, 139, 503, 79, 'C', 447, 28, 376, 0, 299, 0,
            'C', 219, 0, 144, 31, 88, 87, 'C', 31, 144, 0, 219, 0, 299, 'C', 0, 379, 31, 455, 88, 511, 'Z'] },
    { x: 9.178, y: 5.342, w: 4.934, h: 3.953, pw: 628, ph: 503, color: C.teal,
        d: ['M', 328, 0, 'C', 248, 0, 173, 31, 117, 87, 'C', 2, 201, 0, 385, 108, 503,
            'C', 168, 447, 246, 416, 328, 416, 'C', 411, 416, 489, 447, 549, 503, 'C', 600, 447, 628, 375, 628, 299,
            'C', 628, 219, 597, 144, 540, 87, 'C', 484, 31, 409, 0, 328, 0, 'Z'] },
];

/** Two hard-broken lorem lines. Built fresh per call — pptxgenjs mutates run options. */
function lorem2() {
    return [
        { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', options: { breakLine: true } },
        { text: 'do eiusmod tempor incididunt ut' },
    ];
}

function slide12(s) {
    rect(s, 0.005, 0.039, 26.661, 5.498, C.white); // white gradient wash in the source
    headline(s, 'Color Psychology Demystified', 3.666, 2.373, 19.336, 1.434, 88, C.ink, 'center');
    kicker(s, 'UNDERSTANDING THE IMPACT OF COLOR', 8.835, 1.643, 8.997, C.ink, 'center');
    logo(s, 23.413, 0.892, C.ink);

    PETALS_4.forEach(function (p) { freeform(s, p.x, p.y, p.w, p.h, p.pw, p.ph, p.d, p.color); });

    [{ n: '01', x: 10.65, y: 6.645, w: 1.989, c: C.black },
    { n: '04', x: 10.407, y: 10.491, w: 1.949, c: C.white },
    { n: '02', x: 14.642, y: 7.024, w: 2.145, c: C.white },
    { n: '03', x: 14.161, y: 10.87, w: 2.127, c: C.white }].forEach(function (b) {
        txt(s, b.n, b.x, b.y, b.w, 1.346, {
            fontFace: HEAD, fontSize: 80, bold: true, color: b.c,
            align: 'center', margin: 0, paraSpaceBefore: 10, lineSpacingMultiple: 1,
        });
    });

    [{ x: 1.887, tx: 1.887, y: 6.249, a: 'left', lw: 2.803 },
    { x: 1.907, tx: 1.907, y: 10.22, a: 'left', lw: 2.783 },
    { x: 18.733, tx: 22.042, y: 6.249, a: 'right', lw: 2.783 },
    { x: 18.733, tx: 22.042, y: 10.22, a: 'right', lw: 2.783 }].forEach(function (t) {
        caption(s, 'Tittle Here', t.tx, t.y, t.lw, 0.582, C.ink, { align: t.a, lineSpacingMultiple: 1.2 });
        para(s, lorem2(), t.x, t.y + 0.798, 6.092, 1.632, C.ink, { align: t.a });
    });
    siteTag(s, C.ink);
}

function slide13(s) {
    rect(s, 17.335, 8.167, 9.332, 6.833, C.mist);
    kicker(s, 'ENHANCING AMBIANCE AND VISUAL INTEREST', 7.902, 2.926, 9.681);
    headline(s, 'Illuminating Spaces with Light', 7.819, 3.655, 10.183, 2.767, 88);
    logo(s, 23.413, 0.892, C.ink);
    siteTag(s, C.ink);
    rect(s, 3.33, 8.188, 14.005, 4.656, C.sage);
    icon(s, 4.309, 9.109, 0.751, 0.751, C.white);
    note(s, {
        x: 5.909, y: 9.109, lw: 4.866, by: 9.765, bw: 10.072, bh: 2.157, color: C.white,
        text: 'Shed light on the importance of lighting design in interior spaces, from natural and artificial lighting to fixtures and accents. Discover how lighting can transform atmospheres and accentuate architectural features.',
    });
}

const WEEK = ['Mon', 'Thu', 'Wed', 'Tue', 'Fri', 'Sat', 'Sun'];

function slide14(s) {
    rect(s, 0, 2.167, 10.667, 12.833, C.sage);
    logo(s, 23.413, 0.892, C.ink);

    s.addChart('bar', [
        { name: 'A', labels: WEEK, values: [175, 150, 200, 120, 230, 250, 145] },
        { name: 'B', labels: WEEK, values: [200, 170, 250, 180, 160, 110, 260] },
    ], {
        x: 11.862, y: 2.609, w: 13.475, h: 5.995,
        barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 132,
        chartColors: [C.sand, C.mist],
        showLegend: false, showValue: false, showTitle: false,
        catAxisLabelFontFace: BODY, catAxisLabelFontSize: 14, catAxisLabelColor: C.black,
        catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
        catAxisLineColor: 'D9D9D9',
        valAxisLabelFontFace: BODY, valAxisLabelFontSize: 14, valAxisLabelColor: C.black,
        valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
        valAxisLineShow: false, valGridLine: { color: 'D9D9D9', size: 1 },
        chartArea: { fill: { color: C.white, transparency: 100 } },
    });

    [{ ox: 1.546, oc: C.mist, gx: 2.196, v: '37%', vw: 1.937 },
    { ox: 5.545, oc: C.sand, gx: 6.179, v: '37,5%', vw: 2.598 }].forEach(function (k) {
        oval(s, k.ox, 11.625, 0.457, 0.457, k.oc);
        txt(s, 'Nam aliquam', k.gx, 11.562, 2.505, 0.582, { color: C.white, lineSpacingMultiple: 1.2 });
        txt(s, k.v, k.gx, 12.263, k.vw, 1.1,
            { fontSize: 54, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
    });

    [{ ix: 1.302, tx: 2.539, y: 3.487 }, { ix: 1.271, tx: 2.55, y: 7.461 }].forEach(function (b) {
        icon(s, b.ix, b.y, 0.751, 0.751, C.white);
        caption(s, 'Tittle Here', b.tx, b.y, 2.783, 0.619, C.white, { fontSize: 28, lineSpacingMultiple: 1.2 });
        para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
            b.tx, b.y + 0.798, 5.774, 1.632, C.white);
    });

    kicker(s, 'FROM ANTIQUES TO AVANT-GARDE', 12.57, 10.077, 8.997);
    headline(s, 'Tracking Furniture Trends', 12.487, 10.806, 11.055, 2.767, 88);
    siteTag(s, C.ink);
}

function slide15(s) {
    rect(s, 14, 10.167, 12.667, 4.833, C.mist);
    logo(s, 23.413, 0.892, C.ink);
    kicker(s, 'INTEGRATING TEXTURES AND MATERIALS', 16.547, 3.601, 8.997);
    headline(s, 'Adding Depth and Dimension', 16.463, 4.33, 8.859, 2.767, 88);
    icon(s, 1.965, 2.969, 0.91, 0.91, C.ink);
    note(s, {
        x: 1.882, y: 4.345, lw: 4.866, by: 5.0, bw: 5.028, bh: 2.157, color: C.ink,
        text: 'Start to learn how to layer textures for depth and dimension to bring more life to your interior design.',
    });
    icon(s, 1.965, 8.739, 0.91, 0.91, C.ink);
    note(s, {
        x: 1.876, y: 10.285, lw: 4.866, by: 10.94, bw: 5.028, bh: 1.632, color: C.ink,
        text: 'Lighting plays a crucial role in interior design, affecting the mood, and ambiance',
    });
    siteTag(s, C.ink);
    rect(s, 13.583, 8.182, 10.003, 4.656, C.sage);
    icon(s, 14.582, 9.041, 0.751, 0.751, C.white);
    note(s, {
        x: 15.906, y: 9.103, lw: 4.866, by: 9.759, bw: 6.664, bh: 2.157, color: C.white,
        text: 'Explore the role of texture and material selection in interior design, from soft fabrics and natural finishes to sleek metals and tactile surfaces. ',
    });
}

/* Slide 16 — six-petal pinwheel. */
const PETAL_TALL = ['M', 52, 231, 'C', 0, 320, 31, 435, 120, 487, 'C', 210, 539, 325, 508, 377, 418,
    'C', 394, 388, 402, 355, 402, 323, 'C', 402, 323, 402, 323, 402, 323, 'C', 402, 0, 402, 0, 402, 0,
    'C', 122, 161, 122, 161, 122, 161, 'C', 122, 161, 122, 161, 122, 161, 'C', 94, 177, 69, 201, 52, 231, 'Z'];
const PETAL_TALL_FLIP = ['M', 350, 308, 'C', 402, 218, 371, 103, 281, 51, 'C', 191, 0, 77, 30, 25, 120,
    'C', 8, 150, 0, 183, 0, 215, 'C', 0, 215, 0, 215, 0, 215, 'C', 0, 539, 0, 539, 0, 539,
    'C', 280, 377, 280, 377, 280, 377, 'C', 280, 377, 280, 377, 280, 377, 'C', 308, 361, 332, 338, 350, 308, 'Z'];

const PETALS_6 = [
    { x: 10.002, y: 8.856, w: 3.338, h: 4.477, pw: 402, ph: 539, color: C.mist, d: PETAL_TALL },
    { x: 13.339, y: 4.376, w: 3.338, h: 4.48, pw: 402, ph: 539, color: C.sage, d: PETAL_TALL_FLIP },
    { x: 13.339, y: 7.292, w: 4.664, h: 3.115, pw: 562, ph: 375, color: C.sand,
        d: ['M', 375, 375, 'C', 478, 375, 562, 291, 562, 188, 'C', 562, 84, 478, 0, 375, 0,
            'C', 340, 0, 308, 9, 280, 26, 'C', 280, 26, 280, 26, 280, 26, 'C', 0, 188, 0, 188, 0, 188,
            'C', 280, 349, 280, 349, 280, 349, 'C', 280, 349, 280, 349, 280, 349, 'C', 308, 366, 340, 375, 375, 375, 'Z'] },
    { x: 13.339, y: 8.856, w: 3.338, h: 4.477, pw: 402, ph: 539, color: C.grey,
        d: ['M', 25, 418, 'C', 77, 508, 191, 539, 281, 487, 'C', 371, 435, 402, 320, 350, 231,
            'C', 332, 201, 308, 177, 280, 161, 'C', 280, 161, 280, 161, 280, 161, 'C', 0, 0, 0, 0, 0, 0,
            'C', 0, 323, 0, 323, 0, 323, 'C', 0, 323, 0, 323, 0, 323, 'C', 0, 355, 8, 388, 25, 418, 'Z'] },
    { x: 8.663, y: 7.289, w: 4.676, h: 3.115, pw: 563, ph: 375, color: C.ink,
        d: ['M', 188, 0, 'C', 84, 0, 0, 84, 0, 188, 'C', 0, 291, 84, 375, 188, 375,
            'C', 222, 375, 255, 366, 283, 349, 'C', 283, 349, 283, 349, 283, 349, 'C', 563, 188, 563, 188, 563, 188,
            'C', 283, 26, 283, 26, 283, 26, 'C', 283, 26, 283, 26, 283, 26, 'C', 255, 9, 222, 0, 188, 0, 'Z'] },
    { x: 10.002, y: 4.376, w: 3.338, h: 4.48, pw: 402, ph: 539, color: C.cream,
        d: ['M', 377, 120, 'C', 325, 30, 210, 0, 120, 51, 'C', 31, 103, 0, 218, 52, 308,
            'C', 69, 338, 94, 361, 122, 377, 'C', 122, 377, 122, 377, 122, 377, 'C', 402, 539, 402, 539, 402, 539,
            'C', 402, 215, 402, 215, 402, 215, 'C', 402, 215, 402, 215, 402, 215, 'C', 402, 183, 394, 150, 377, 120, 'Z'] },
];

function slide16(s) {
    rect(s, 0, 0.028, 26.661, 5.498, C.white); // white gradient wash in the source
    headline(s, 'Embracing Eco-Friendly Design', 3.059, 2.211, 20.549, 1.555, 96, C.ink, 'center');
    kicker(s, 'PROMOTING SUSTAINABILITY IN INTERIORS', 8.469, 1.667, 9.728, C.ink, 'center');
    logo(s, 23.327, 0.833, C.ink);
    siteTag(s, C.ink);

    PETALS_6.forEach(function (p) { freeform(s, p.x, p.y, p.w, p.h, p.pw, p.ph, p.d, p.color); });

    [{ x: 3.247, tx: 3.247, y: 4.91, a: 'left' },
    { x: 1.21, tx: 1.21, y: 7.878, a: 'left' },
    { x: 2.577, tx: 2.577, y: 10.846, a: 'left' },
    { x: 18.542, tx: 20.699, y: 4.91, a: 'right' },
    { x: 20.542, tx: 22.699, y: 7.878, a: 'right' },
    { x: 19.208, tx: 21.365, y: 10.846, a: 'right' }].forEach(function (t) {
        caption(s, 'Tittle Here', t.tx, t.y, 2.803, 0.582, C.ink, { align: t.a, lineSpacingMultiple: 1.2 });
        para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ',
            t.x, t.y + 0.798, 4.959, 1.107, C.ink, { align: t.a });
    });

    [{ x: 11.202, y: 5.513, c: C.ink }, { x: 14.39, y: 5.513, c: 'E7E6E6' },
    { x: 9.421, y: 8.256, c: 'E7E6E6' }, { x: 16.085, y: 8.259, c: C.ink },
    { x: 11.265, y: 10.869, c: C.ink }, { x: 14.399, y: 10.89, c: 'E7E6E6' }].forEach(function (g) {
        icon(s, g.x, g.y, 1.182, 1.182, g.c);
    });
}

function slide17(s) {
    rect(s, 0.025, 10.167, 16.644, 4.833, C.mist);
    kicker(s, 'EMBRACING INNOVATION AND CREATIVITY', 8.563, 2.486, 9.441);
    headline(s, 'Reflecting on the Journey', 8.549, 3.255, 8.787, 3.009, 96);
    logo(s, 23.413, 0.892, C.ink);
    siteTag(s, C.ink);
    icon(s, 19.856, 3.155, 1.082, 1.082, C.ink);
    note(s, {
        x: 19.918, y: 4.939, lw: 4.866, by: 5.595, bw: 5.917, bh: 1.632, color: C.ink,
        text: 'Summarize key takeaways from the exploration of interior design and its multifaceted aspects. ',
    });
    icon(s, 19.981, 8.482, 1.082, 1.082, C.ink);
    note(s, {
        x: 19.918, y: 10.261, lw: 4.866, by: 10.917, bw: 5.917, bh: 1.632, color: C.ink,
        text: 'Lighting plays a crucial role in interior design, affecting the mood, and ambiance',
    });
    rect(s, 2.023, 8.852, 8.002, 3.989, C.sage);
    icon(s, 2.605, 9.702, 0.751, 0.751, C.white);
    note(s, {
        x: 3.929, y: 9.702, lw: 4.866, by: 10.358, bw: 5.449, bh: 1.632, color: C.white,
        text: 'Look ahead to emerging trends and innovations shaping the future of interior.',
    });
}

function slide18(s) {
    rect(s, 14, 0, 12.667, 7.5, C.sage);
    rect(s, 0, 7.5, 7.998, 7.5, C.mist);
    logo(s, 23.327, 0.838, C.white);
    siteTag(s, C.ink);
    kicker(s, 'PROMOTING ENVIRONMENTAL RESPONSIBILITY', 15.212, 2.518, 10.107, C.white);
    headline(s, 'Designing with a Purpose', 15.198, 3.287, 9.026, 3.009, 96, C.white);
    caption(s, 'Text Here', 15.231, 10.229, 5.822, 0.619, C.ink, { fontSize: 28, lineSpacingMultiple: 1.2 });
    para(s, 'Embrace the principles of sustainable and eco-friendly design, from energy-efficient practices to environmentally friendly materials. ',
        15.209, 11.027, 8.499, 1.632, C.ink);
    icon(s, 15.165, 8.667, 0.989, 0.989, C.ink);
}

function slide19(s) {
    rect(s, 22.687, 0, 3.98, 15, C.sage);
    kicker(s, 'REACH US ANYTIME, ANYWHERE', 11.897, 2.288, 8.121);
    headline(s, "Let's Connect", 11.801, 3.057, 8.89, 1.555, 96);

    [{ y: 6.194, label: 'Phone', lw: 2.298, value: '+1 (555) 123-4567', vw: 3.865, iw: 0.667, ih: 0.667, iy: 6.513, ix: 12.365, vy: 6.91, ly: 6.218 },
    { y: 8.794, label: 'Email', lw: 2.362, value: 'info@interiordesigncompany.com', vw: 6.787, iw: 0.679, ih: 0.502, iy: 9.195, ix: 12.359, vy: 9.556, ly: 8.863 },
    { y: 11.535, label: 'Website', lw: 3.003, value: 'www.interiordesigncompany.com', vw: 6.787, iw: 0.679, ih: 0.502, iy: 11.936, ix: 12.359, vy: 12.187, ly: 11.537 }].forEach(function (r) {
        oval(s, 12.046, r.y, 1.304, 1.304, C.sage);
        icon(s, r.ix, r.iy, r.iw, r.ih, C.white);
        caption(s, r.label, 13.946, r.ly, r.lw, 0.619, C.ink, { fontSize: 28, lineSpacingMultiple: 1.2 });
        txt(s, r.value, 13.946, r.vy, r.vw, 0.582, { color: C.ink, lineSpacingMultiple: 1.2 });
    });
    siteTag(s, C.ink);
}

function slide20(s) {
    rect(s, 7.998, 0, 14.67, 10.833, C.mist);
    siteTag(s, C.ink);
    kicker(s, 'A BIG THANK YOU FROM OUR TEAM', 11.256, 3.158, 7.759);
    txt(s, 'Gratitude & Appreciation', 11.164, 3.927, 8.964, 2.699,
        { fontFace: HEAD, fontSize: 96, color: C.ink, lineSpacingMultiple: 0.8 });
    logo(s, 23.327, 0.833, C.ink);
    rect(s, 6.665, 8.833, 14.67, 4, C.sage);
    icon(s, 7.302, 9.689, 0.751, 0.751, C.white);
    note(s, {
        x: 8.626, y: 9.689, lw: 4.866, by: 10.345, bw: 11.964, bh: 1.632, color: C.white,
        text: "We extend our heartfelt gratitude for your time and attention. Thank you for considering us for your interior design needs. Let's create spaces that inspire together!",
    });
}

/* ------------------------------------------------------------------ build */

const SLIDES = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

function build() {
    const pres = new pptxgen();
    pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
    pres.layout = 'DECK';
    pres.title = 'Exploring the World of Interior Design';
    pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

    SLIDES.forEach(function (fn) {
        const s = pres.addSlide();
        s.background = { color: C.white };
        fn(s);
    });

    return pres.writeFile({
        fileName: path.join(__dirname, '0cec86d1-2d4e-4202-baa1-6af0b5c7bb81_grok_final.pptx'),
    });
}

build().then(function (f) { console.log('wrote', f); }, function (e) { console.error(e); process.exit(1); });
