/**
 * "Handmade Jewelry" deck — rebuilt with pptxgenjs.
 * 15 slides, 13.333in x 7.5in (16:9).
 *
 * The reference deck's picture placeholders are empty (decorative frames with
 * no picture inside), so they are left out except the full-bleed photo slot on
 * "Materials Used", which becomes a labelled rectangle. The one real raster
 * asset — a ring icon — is redrawn with native shapes.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const RUST = '853728'; // brand red-brown: dark slide bg, headings, rules
const CREAM = 'D7C8B6'; // light slide bg
const GOLD = 'DAB671'; // accent
const WHITE = 'FFFFFF';
const INK = '404040'; // body copy on cream (tx1 @ 75% lumMod)
const GREY = 'A5A5A5'; // chart accent3
const YELLOW = 'FFC000'; // chart accent4
const LEGEND = '595959';

const HEAD = 'Playfair Display Black';
const BODY = 'Lora';

const RULE = 1.5; // pt — every hairline rule / circle outline
const OVAL_D = 1.391; // in — the decorative outline circles

// pptxgenjs has no "no fill"/"no line", so use a fully transparent colour.
const NONE = { color: WHITE, transparency: 100 };

/* ------------------------------------------------- lorem ipsum, by length */

const LOREM_1 =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent imperdiet quis ' +
    'eros sed pellentesque. Quisque sed pretium purus, nec luctus lorem.';
const LOREM_2 =
    LOREM_1 + ' Morbi molestie tincidunt hendrerit. Maecenas nisi massa, tempor in ' +
    'convallis eu, accumsan sit amet purus. ';
const LOREM_3 = LOREM_2 + 'Aliquam erat volutpat. Donec venenatis fermentum lacus id congue. ';
const LOREM_4 =
    LOREM_3 + 'Etiam porttitor massa nec turpis malesuada, molestie venenatis turpis blandit. ';
const LOREM_5 = LOREM_4 + 'Donec accumsan sollicitudin ante, non sagittis sapien dignissim et';
const LOREM_6 =
    LOREM_5 + '. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere ' +
    'cubilia curae; Duis at neque in felis semper tincidunt. Suspendisse bibendum tempor gravida.';

/* ------------------------------------------------------------------ helpers */

/** Straight connector. Horizontal rules pass h:0, vertical rules pass w:0. */
function rule(slide, x, y, w, h, color) {
    slide.addShape('line', { x, y, w, h, line: { color, width: RULE } });
}

/** Decorative outline circle. */
function circle(slide, x, y, color, d = OVAL_D) {
    slide.addShape('ellipse', { x, y, w: d, h: d, fill: NONE, line: { color, width: RULE } });
}

/** Rounded rectangle; `adj` is the OOXML corner adjust value (0..50000). */
function card(slide, x, y, w, h, adj, fill, line) {
    slide.addShape('roundRect', {
        x, y, w, h,
        rectRadius: (adj / 100000) * Math.min(w, h),
        fill: fill ? { color: fill } : NONE,
        line: line ? { color: line, width: RULE } : NONE,
    });
}

/** Display heading — Playfair Display Black, single line. */
function head(slide, x, y, w, text, color, size = 40, h = 0.774, wrap = false) {
    slide.addText(text, {
        x, y, w, h,
        fontFace: HEAD, fontSize: size, bold: true, color,
        align: 'left', valign: 'top', wrap, fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6],
    });
}

/** Body copy — Lora 12pt. */
function copy(slide, x, y, w, h, text, color, align = 'left') {
    slide.addText(text, {
        x, y, w, h,
        fontFace: BODY, fontSize: 12, color,
        align, valign: 'top', wrap: true, fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6],
    });
}

/** Pill button: outlined capsule + centred caption. */
function pill(slide, x, y, w, label, color, capX, capW) {
    card(slide, x, y, w, 0.499, 50000, null, color);
    copy(slide, capX, y + 0.098, capW, 0.303, label, color === GOLD ? WHITE : INK, 'center');
}

/**
 * Stand-in for a photo slot. The reference keeps its picture frames empty, so
 * only the one slot that actually renders (the big panel on "Materials Used")
 * is drawn, as a plain light rectangle.
 */
function imagePanel(slide, x, y, w, h) {
    slide.addShape('rect', { x, y, w, h, fill: { color: WHITE }, line: NONE });
    slide.addText('[image]', {
        x, y, w, h,
        fontFace: BODY, fontSize: 12, color: 'BFBFBF', align: 'center', valign: 'middle',
    });
}

/** Brand mark: three overlapping dots (freeform in the reference). */
function logo(slide, x, y) {
    const d = 0.277; // dot diameter, relative to the 0.502in mark
    [[0, 0], [0, 0.225], [0.225, 0.225]].forEach(([dx, dy]) => {
        slide.addShape('ellipse', { x: x + dx, y: y + dy, w: d, h: d, fill: { color: GOLD }, line: NONE });
    });
}

/** Ring icon (redrawn from the deck's 384px ring/gem PNG). 1in box. */
function ringIcon(slide, x, y) {
    slide.addShape('ellipse', {
        x: x + 0.221, y: y + 0.299, w: 0.557, h: 0.557,
        fill: NONE, line: { color: RUST, width: 4.4 },
    });
    const gem = [[0.365, 0.178], [0.418, 0.120], [0.582, 0.120], [0.635, 0.178], [0.500, 0.332]];
    slide.addShape('custGeom', {
        x, y, w: 1, h: 1, fill: { color: RUST }, line: NONE,
        points: gem.map(([px, py]) => ({ x: px, y: py })).concat([{ close: true }]),
    });
}

/** One of the four doughnuts on the Market Trends slide. */
function doughnut(slide, x, y, labels, values) {
    slide.addChart('doughnut', [{ name: 'Sales', labels, values }], {
        x, y, w: 3.429, h: 2.861,
        holeSize: 75,
        chartColors: [RUST, GOLD, GREY, YELLOW],
        dataBorder: { pt: 1.5, color: WHITE },
        showValue: false,
        showLegend: true,
        legendPos: 'b',
        legendFontSize: 12,
        legendColor: LEGEND,
        legendFontFace: 'Calibri',
        chartArea: { fill: NONE },
        plotArea: { fill: NONE },
    });
}

/** The vertical hairline at x=0.783 that runs down most interior slides. */
function spine(slide) {
    rule(slide, 0.783, -0.127, 0, 7.635, RUST);
}

/* ------------------------------------------------------------------- slides */

const slides = [];

// 1 — title
slides.push((s) => {
    s.background = { color: RUST };
    rule(s, 11.683, -0.097, 0, 7.635, GOLD);
    rule(s, 1.651, 0, 0, 7.635, GOLD);
    rule(s, 0, 1.587, 13.333, 0, GOLD);
    head(s, 4.893, 3.397, 3.671, 'Handmade', WHITE, 48, 0.909);
    head(s, 4.893, 4.071, 2.579, 'Jewelry', GOLD, 48, 0.909);
    copy(s, 4.893, 4.972, 3.671, 0.909, LOREM_1, WHITE);
    pill(s, 4.982, 6.048, 2.49, 'Product by Liceria Jewelry', GOLD, 5.042, 2.371);
    copy(s, 5.449, 3.085, 1.03, 0.303, 'Your Logo', WHITE);
    logo(s, 4.982, 2.845);
    rule(s, 0, 6.992, 13.333, 0, GOLD);
    [[2.295, 5.354], [8.29, 2.134], [-0.385, 2.651], [12.286, 4.072]].forEach(([x, y]) => circle(s, x, y, GOLD));
});

// 2 — welcome
slides.push((s) => {
    rule(s, -0.005, 6.833, 13.333, 0, RUST);
    rule(s, 0, 1.302, 13.333, 0, RUST);
    head(s, 1.176, 2.108, 3.279, 'Welcome to', RUST);
    head(s, 1.176, 2.71, 4.79, 'Our Presentation', RUST);
    copy(s, 1.176, 3.502, 4.621, 2.524, LOREM_6, INK);
    spine(s);
    [[5.966, 1.699], [12.187, 4.056]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 3 — definition
slides.push((s) => {
    head(s, 7.038, 1.489, 3.575, 'Definition of', RUST);
    head(s, 7.038, 2.091, 5.198, 'Handmade Jewelry', RUST);
    copy(s, 1.098, 4.477, 5.189, 1.717, LOREM_5, INK);
    rule(s, -0.003, 0.746, 13.333, 0, RUST);
    rule(s, 0.638, -0.135, 0, 7.635, RUST);
    rule(s, 6.664, -0.067, 0, 7.635, RUST);
    rule(s, 12.695, -0.067, 0, 7.635, RUST);
    rule(s, 0, 6.786, 13.333, 0, RUST);
    [[5.009, 2.927], [11.756, 3.871]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 4 — history timeline: three gold cards over a rust band
slides.push((s) => {
    rule(s, 11.934, -0.067, 0, 7.635, RUST);
    rule(s, -0.003, 0.746, 13.333, 0, RUST);
    head(s, 1.468, 1.425, 4.062, 'History of Our', RUST);
    head(s, 1.468, 2.028, 5.198, 'Handmade Jewelry', RUST);
    s.addShape('rect', { x: 0, y: 4.586, w: 13.333, h: 2.914, fill: { color: RUST }, line: NONE });
    const cards = [
        { year: '2022', cardX: 1.468, cardY: 3.302, yearX: 2.381, yearW: 1.119, iconX: 2.453, textX: 1.708 },
        { year: '2023', cardX: 4.667, cardY: 3.297, yearX: 5.622, yearW: 1.103, iconX: 5.686, textX: 4.941 },
        { year: '2024', cardX: 7.865, cardY: 3.297, yearX: 8.777, yearW: 1.117, iconX: 8.848, textX: 8.104 },
    ];
    cards.forEach((c) => card(s, c.cardX, c.cardY, 2.968, 3.603, 21480, GOLD));
    spine(s);
    cards.forEach((c) => {
        s.addText(c.year, {
            x: c.yearX, y: 3.655, w: c.yearW, h: 0.572,
            fontFace: HEAD, fontSize: 28, bold: true, color: RUST,
            align: 'center', valign: 'top', wrap: false, fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6],
        });
        ringIcon(s, c.iconX, 4.269);
        copy(s, c.textX, 5.265, 2.489, 1.313, LOREM_1, INK, 'center');
    });
});

// 5 — materials
slides.push((s) => {
    rule(s, 0, 6.198, 13.333, 0, RUST);
    rule(s, 0, 1.302, 13.333, 0, RUST);
    head(s, 1.345, 1.957, 2.723, 'Materials', RUST);
    head(s, 1.345, 2.575, 1.48, 'Used', RUST);
    copy(s, 1.355, 3.423, 3.633, 2.121, LOREM_4, INK);
    spine(s);
    imagePanel(s, 5.551, 0.595, 7.302, 6.309);
    [[12.158, 3.978], [4.855, 1.758]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 6 — artisan techniques
slides.push((s) => {
    rule(s, 0, 1.302, 13.333, 0, RUST);
    head(s, 1.345, 1.957, 2.155, 'Artisan', RUST);
    head(s, 1.345, 2.575, 3.277, 'Techniques', RUST);
    copy(s, 1.355, 3.423, 4.058, 1.919, LOREM_4, INK);
    spine(s);
    rule(s, 0, 6.782, 13.333, 0, RUST);
    pill(s, 1.355, 5.562, 1.37, 'See More', RUST, 1.418, 1.244);
    [[5.413, 2.046], [11.283, 4.965]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 7 — design inspiration: rust card over gold card
slides.push((s) => {
    rule(s, 0.01, 1.586, 13.333, 0, RUST);
    head(s, 1.366, 2.358, 2.06, 'Design', RUST);
    head(s, 1.366, 2.976, 3.194, 'Inspiration', RUST);
    card(s, 6.281, 0.773, 6.407, 2.977, 19380, RUST);
    card(s, 6.281, 3.961, 6.407, 2.977, 19380, GOLD);
    spine(s);
    copy(s, 1.366, 3.75, 3.206, 2.524, LOREM_4, INK);
    [{ y: 1.209, ink: WHITE }, { y: 4.395, ink: RUST }].forEach(({ y, ink }) => {
        head(s, 7.76, y, 3.982, 'Inspiration 01', ink, 32, 0.64, true);
        copy(s, 7.76, y + 0.655, 4.387, 1.515, LOREM_3, ink);
    });
});

// 8 — market trends: four doughnut charts
slides.push((s) => {
    head(s, 1.393, 2.437, 2.132, 'Market', RUST);
    head(s, 1.393, 3.055, 2.104, 'Trends', RUST);
    copy(s, 1.393, 3.829, 3.633, 2.121, LOREM_4, INK);
    const QTR3 = ['1st Qtr', '2nd Qtr', '3rd Qtr'];
    doughnut(s, 5.825, 0.795, QTR3.concat('4th Qtr'), [8.2, 3.2, 1.4]);
    doughnut(s, 9.254, 0.782, QTR3, [6, 3.2, 6]);
    doughnut(s, 5.825, 3.763, QTR3, [4, 7, 1.4]);
    doughnut(s, 9.254, 3.75, QTR3, [6, 8, 1.4]);
    spine(s);
    rule(s, 5.442, -0.067, 0, 7.635, RUST);
    rule(s, -0.048, 1.698, 5.49, 0, RUST);
    rule(s, -0.032, 6.611, 5.49, 0, RUST);
});

// 9 — advantages
slides.push((s) => {
    rule(s, 0, 1.302, 13.333, 0, RUST);
    head(s, 7.27, 1.786, 3.883, 'Advantages of', RUST);
    head(s, 7.27, 2.404, 5.198, 'Handmade Jewelry', RUST);
    copy(s, 1.25, 4.925, 5.115, 1.515, LOREM_4, INK);
    card(s, 7.27, 3.462, 7.251, 3.3, 19380, RUST);
    copy(s, 7.778, 4.017, 5.115, 1.111, LOREM_2, WHITE);
    copy(s, 7.778, 5.128, 5.115, 1.111, LOREM_2, WHITE);
    spine(s);
    circle(s, 5.546, 3.445, RUST);
});

// 10 — ethical practices
slides.push((s) => {
    rule(s, 0, 1.302, 13.333, 0, RUST);
    spine(s);
    head(s, 1.268, 1.649, 2.062, 'Ethical', RUST);
    head(s, 1.268, 2.268, 2.644, 'Practices', RUST);
    copy(s, 1.268, 3.042, 5.115, 1.515, LOREM_4, INK);
    card(s, 1.268, 4.731, 4.743, 2.202, 19380, RUST);
    copy(s, 1.649, 5.075, 4.163, 1.515, LOREM_3, WHITE);
    [[6.627, 5.833], [12.295, 1.728]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 11 — break slide
slides.push((s) => {
    rule(s, 0, 2.406, 13.333, 0, RUST);
    head(s, 1.227, 4.488, 5.235, 'Break Slide', RUST, 66, 1.212);
    spine(s);
    s.addShape('rect', { x: -0.879, y: 6.012, w: 6.407, h: 2.977, fill: { color: RUST }, line: NONE });
    [[6.259, 5.571], [5.238, 0.538]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 12 — case studies
slides.push((s) => {
    rule(s, 0, 1.079, 13.333, 0, RUST);
    head(s, 1.268, 1.253, 1.448, 'Case', RUST);
    head(s, 1.268, 1.871, 2.165, 'Studies', RUST);
    copy(s, 1.268, 2.645, 5.115, 1.515, LOREM_4, INK);
    card(s, 1.268, 4.335, 5.115, 2.202, 19380, RUST);
    copy(s, 1.649, 4.678, 4.49, 1.515, LOREM_3, WHITE);
    spine(s);
    circle(s, 6.95, 5.841, RUST);
});

// 13 — handmade jewelry
slides.push((s) => {
    rule(s, -0.013, 1.556, 13.333, 0, RUST);
    head(s, 1.45, 2.247, 3.093, 'Handmade', RUST);
    head(s, 1.45, 2.865, 2.183, 'Jewelry', RUST);
    copy(s, 1.45, 3.75, 5.115, 1.515, LOREM_4, INK);
    spine(s);
    pill(s, 1.499, 5.445, 1.37, 'See More', RUST, 1.562, 1.244);
    [[5.958, 5.607], [11.648, 2.511]].forEach(([x, y]) => circle(s, x, y, RUST));
});

// 14 — conclusion
slides.push((s) => {
    rule(s, 0, 6.278, 13.333, 0, RUST);
    rule(s, 0, 1.762, 13.333, 0, RUST);
    head(s, 1.383, 3.56, 3.215, 'Conclusion', RUST);
    copy(s, 1.383, 4.334, 5.115, 1.515, LOREM_4, INK);
    spine(s);
    circle(s, 11.595, 4.286, RUST, 1.403);
    circle(s, 6.401, 2.295, RUST);
});

// 15 — thank you
slides.push((s) => {
    s.background = { color: RUST };
    rule(s, 0, 1.587, 13.333, 0, GOLD);
    rule(s, 12.429, -0.067, 0, 7.635, GOLD);
    rule(s, 0.667, -0.067, 0, 7.635, GOLD);
    head(s, 8.369, 2.413, 2.279, 'Thank', WHITE, 48, 0.909);
    head(s, 8.369, 3.087, 1.441, 'You', GOLD, 48, 0.909);
    copy(s, 8.369, 3.988, 3.671, 0.909, LOREM_1, WHITE);
    pill(s, 8.458, 5.063, 2.49, 'www.yourwebsite.com', GOLD, 8.518, 2.371);
    copy(s, 8.925, 2.101, 1.03, 0.303, 'Your Logo', WHITE);
    logo(s, 8.458, 1.861);
    rule(s, 0, 6.802, 13.333, 0, GOLD);
    [[6.632, 4.703], [0.861, 3.241]].forEach(([x, y]) => circle(s, x, y, GOLD));
});

/* --------------------------------------------------------------------- main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 40 / 3, height: 7.5 }); // 12192000 x 6858000 EMU
pptx.layout = 'WIDE';
pptx.title = 'Handmade Jewelry';

slides.forEach((build) => {
    const slide = pptx.addSlide();
    slide.background = { color: CREAM };
    build(slide);
});

pptx.writeFile({
    fileName: path.join(__dirname, '0e057d7a-ab7c-45ba-ac36-fcde2cc797f5_grok_final.pptx'),
});
