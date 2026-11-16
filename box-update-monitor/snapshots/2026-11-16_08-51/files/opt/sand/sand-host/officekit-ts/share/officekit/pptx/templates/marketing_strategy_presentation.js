/**
 * "Lemida – Marketing Strategy Presentation" rebuilt with pptxgenjs.
 *
 * 25 slides, 20 x 11.25 in (16:9 widescreen).  Every shape, colour and string
 * below is a plain literal so the deck's design can be read straight off the
 * source.  The original file contains no photographs; the only bitmaps are tiny
 * SVG/PNG glyph icons, which are re-created here as simple coloured placeholder
 * marks (see `glyph`).  The empty picture placeholders of the original template
 * carry a 5% pattern fill that renders as nothing, so they are not drawn.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BG = '0C1016';        // slide background
const CARD = '16222F';      // dark card panel
const GREEN = '1BB36E';
const BLUE = '2A60ED';
const SKY = '088EEF';
const PURPLE = '7367EF';
const GOLD = 'FFC000';      // review stars
const W = 'FFFFFF';
const T95 = 'F2F2F2';       // white lumMod 95%
const T85 = 'D9D9D9';       // white lumMod 85% - body copy
const T65 = 'A6A6A6';       // white lumMod 65% - nav links
const GREY = '7F7F7F';
const GREY_DK = '404040';

/* -------------------------------------------------------------- typography */
const HEAD = 'Plus Jakarta Sans SemiBold';
const HEAVY = 'Plus Jakarta Sans ExtraBold';
const BODY = 'Poppins';
const BODY_SB = 'Poppins SemiBold';
const BODY_XB = 'Poppins ExtraBold';

/* --------------------------------------------------------- recurring copy */
const LOREM = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna';
const LOREM_URNA = LOREM + ' eros quis urna.';
const LOREM_ER = 'Lorem ipsum dolor sit amet, consectetuer adipiscinger elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, ';
const LOREM_SHORT = 'Lorem ipsum dolor sit ameta consectetuer';
const LOREM_SECT = 'Lorem ipsum dolor sit amet, consect etuer adip iscing elit. Maecenas port titor congue massare';
const LOREM_CARD = 'PLACEHOLDER';
const LOREM_CARD2 = 'Lorem ipsum dolor sit ame consectetuer adipiscing el';
const LOREM_SVC = 'Lorem ipsum dolor sit amet, consect etuer adip iscing.';
const LOREM_TL = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ';
const LOREM_SWOT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maec enas porttitor congue massa. Fusce posuere, magna sed pulvi inar ultricies, purus lectus malesuada libero';
const QUOTE_BEST = "The best marketing doesn't feel like marketing.";
const QUOTE_CULTURE = 'Your culture is your brand.';
const CARD_TITLE = 'Our Title Section Here';
const EYEBROW = 'Marketing Strategy';

/* ---------------------------------------------------------------- helpers */

// Rounded rectangle: `adj` is the OOXML adjust fraction (radius / shorter side).
function rr(s, x, y, w, h, fill, adj = 0.1222, extra = {}) {
    s.addShape('roundRect', Object.assign(
        { x, y, w, h, rectRadius: adj * Math.min(w, h), fill: { color: fill } }, extra));
}

function oval(s, x, y, d, fill, extra = {}) {
    s.addShape('ellipse', Object.assign({ x, y, w: d, h: d, fill: { color: fill } }, extra));
}

// Text frame (transparent, top anchored) - the deck's text boxes auto-fit height.
function tx(s, text, x, y, w, h, o = {}) {
    s.addText(text, Object.assign({
        x, y, w, h, fontFace: HEAD, fontSize: 18, color: W, valign: 'top',
    }, o));
}

// 14pt Poppins body copy at 150% line spacing.
function para(s, text, x, y, w, h, o = {}) {
    tx(s, text, x, y, w, h, Object.assign(
        { fontFace: BODY, fontSize: 14, color: T85, lineSpacingMultiple: 1.5 }, o));
}

function eyebrow(s, x, y, text = EYEBROW, o = {}) {
    tx(s, text, x, y, 3.828, 0.438, Object.assign({ fontSize: 20, color: GREEN }, o));
}

// Pill button / search field.
function pill(s, text, x, y, w, h, fill, o = {}) {
    s.addText(text, Object.assign({
        shape: 'roundRect', rectRadius: Math.min(w, h) / 2, x, y, w, h,
        fill: { color: fill }, align: 'center', valign: 'middle',
        fontFace: BODY_SB, fontSize: 16, color: W,
    }, o));
}

// Placeholder for one of the deck's small SVG glyph icons: a compact mark
// whose footprint matches the icon it stands in for.
function glyph(s, x, y, size, color) {
    const d = size * 0.62;
    s.addShape('roundRect', {
        x: x + (size - d) / 2, y: y + (size - d) / 2, w: d, h: d,
        fill: { color }, rectRadius: d * 0.25,
    });
}

// Rounded tile with a centred glyph placeholder inside it.
function iconTile(s, x, y, w, h, tile, mark, gs = w * 0.62) {
    rr(s, x, y, w, h, tile, 0.2066);
    glyph(s, x + (w - gs) / 2, y + (h - gs) / 2, gs, mark);
}

// Tick mark used in the feature lists (replaces a white check-mark icon).
function check(s, x, y, sz) {
    s.addShape('custGeom', {
        x, y, w: sz, h: sz, line: { color: W, width: 2.25 },
        points: [{ x: sz * 0.12, y: sz * 0.52 }, { x: sz * 0.4, y: sz * 0.82 }, { x: sz * 0.9, y: sz * 0.18 }],
    });
}

// The "half ring + dot" brand mark.  arc/dot default to the white nav version.
function logoMark(s, x, y, w = 0.41, h = 0.405, arc = W, dot = W) {
    s.addShape('blockArc', { x, y, w: w * 0.818, h: h * 0.828, fill: { color: arc }, rotate: 315 });
    oval(s, x + w * 0.409, y + h * 0.401, w * 0.591, dot);
}

function wordmark(s, x, y) {
    logoMark(s, x, y);
    tx(s, 'LEMIDA', x + 0.493, y - 0.033, 1.509, 0.438, { fontFace: HEAVY, fontSize: 20 });
}

const NAV_W = { 'Presentation': 2.111, 'Marketing': 2.111, 'Business Stategy': 2.616 };

function navLinks(s, items, align = 'left', size = 18) {
    items.forEach(([label, x, y]) => tx(s, label, x, y, NAV_W[label], size === 20 ? 0.438 : 0.404,
        { fontSize: size, color: T65, align }));
}

// Header shared by the majority of the body slides.
function header(s) {
    wordmark(s, 0.959, 0.84);
    wordmark(s, 17.301, 0.84);
    navLinks(s, [['Presentation', 5.228, 0.807], ['Marketing', 8.437, 0.807],
        ['Business Stategy', 11.381, 0.807]], 'center');
}

// Dark review card: five stars, a title and a caption.
function reviewCard(s, x, y = 2.561) {
    rr(s, x, y, 3.836, 3.982, CARD);
    for (let i = 0; i < 5; i++) {
        s.addShape('star5', { x: x + 0.621 + i * 0.4855, y: y + 0.517, w: 0.305, h: 0.305, fill: { color: GOLD } });
    }
    tx(s, CARD_TITLE, x + 0.503, y + 1.178, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_CARD, x + 0.546, y + 2.309, 2.833, 1.128);
}

// Solid colour card with the same title/caption pairing.
function colourCard(s, x, y, fill) {
    rr(s, x, y, 3.796, 2.76, fill);
    tx(s, CARD_TITLE, x + 0.472, y + 0.427, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_CARD2, x + 0.506, y + 1.438, 2.833, 0.774);
}

// SWOT letter block: dark backing panel, big letter tile, copy and two buttons.
function swotBlock(s, x, letter, letterFill, btn1, btn2) {
    rr(s, x, 5.07, 5.34, 4.935, CARD);
    s.addText(letter, {
        shape: 'roundRect', rectRadius: 0.13866 * 4.042, x: x + 0.491, y: 2.561, w: 4.359, h: 4.042,
        fill: { color: letterFill }, align: 'center', valign: 'middle',
        fontFace: BODY_XB, fontSize: 199, color: W,
    });
    para(s, [
        { text: 'Lorem ipsum dolor sit amet, consectet uer ' },
        { text: 'adipiscing elit. Maecenas porttitor congue massa. ', options: { fontFace: BODY_SB } },
        { text: 'Fusce posuere' },
    ], x + 0.768, 7.12, 4.241, 1.128);
    pill(s, 'Read More', x + 0.845, 8.748, 1.641, 0.56, btn1, { fontSize: 14 });
    pill(s, 'Learn More', x + 2.965, 8.748, 1.641, 0.56, btn2, { fontSize: 14 });
}

// Play button (white circle + triangle) used on the testimonial cards.
function playButton(s, x, y, triFill) {
    oval(s, x, y, 0.518, W);
    s.addShape('triangle', { x: x + 0.163, y: y + 0.174, w: 0.248, h: 0.214, fill: { color: triFill }, rotate: 90 });
}

/* World map outlines for slide 24 - [x, y] pairs in slide inches. */
const MAP_CONTINENTS = [
    [[8.12,4.82], [8.12,5.25], [7.93,5.48], [8.03,5.60], [8.04,5.76], [7.92,5.96],
     [7.82,5.97], [7.89,6.10], [7.70,6.17], [7.78,6.39], [7.25,6.55], [7.17,6.39],
     [7.24,6.34], [7.09,6.17], [7.14,6.10], [6.97,5.80], [7.11,5.60], [7.01,5.57],
     [7.08,5.48], [6.91,5.28], [6.96,5.08], [6.67,5.02], [6.35,5.12], [6.02,4.77],
     [6.02,4.44], [6.09,4.21], [6.21,4.20], [6.30,3.97], [7.00,3.81], [6.97,3.99],
     [7.30,4.13], [7.34,4.02], [7.73,4.08], [7.97,4.69]],
    [[9.17,2.44], [9.02,2.45], [9.13,2.30], [9.11,2.02], [8.99,2.04], [8.99,2.32],
     [8.71,2.25], [8.75,2.31], [8.52,2.35], [8.71,2.39], [8.35,2.67], [8.38,3.10],
     [8.59,3.10], [8.54,3.27], [8.32,3.23], [8.28,3.33], [8.34,3.48], [8.46,3.52],
     [8.47,3.64], [8.75,3.75], [8.52,3.87], [7.85,3.65], [7.57,3.76], [7.53,3.88],
     [7.87,3.91], [7.85,4.02], [7.73,4.08], [8.18,4.87], [8.36,4.70], [8.52,4.67],
     [8.68,4.42], [8.53,4.32], [8.41,4.37], [8.29,4.14], [8.36,4.12], [8.45,4.31],
     [8.52,4.26], [8.65,4.36], [8.63,4.31], [8.84,4.30], [9.02,4.44], [9.17,4.47],
     [9.09,4.65], [9.17,4.74]],
    [[3.59,4.76], [3.61,4.66], [3.73,4.67], [3.67,4.73], [3.80,4.89], [3.97,4.89],
     [4.11,4.77], [4.16,4.89], [4.45,4.85], [4.72,5.04], [4.87,5.04], [4.86,5.27],
     [5.21,5.33], [5.44,5.49], [5.23,5.76], [5.15,6.08], [5.00,6.10], [4.93,6.26],
     [4.86,6.24], [4.76,6.55], [4.54,6.43], [4.66,6.63], [4.46,6.81], [4.38,6.73],
     [4.37,6.94], [4.27,6.94], [4.36,7.08], [4.23,7.36], [4.10,7.25], [4.01,7.00],
     [4.11,6.88], [4.03,6.63], [4.15,6.47], [4.12,6.18], [4.23,6.17], [4.16,6.04],
     [4.20,5.96], [4.06,5.80], [3.94,5.80], [3.95,5.63], [3.75,5.40], [3.75,5.28],
     [3.87,5.22], [3.85,5.13], [3.97,5.10], [3.89,4.94], [3.74,4.98], [3.68,4.77]],
    [[7.95,3.68], [8.32,3.82], [8.18,3.61], [8.34,3.48], [8.28,3.34], [8.32,3.24],
     [8.51,3.23], [8.51,3.10], [8.39,3.10], [8.37,2.75], [8.51,2.55], [8.51,2.35],
     [8.40,2.39], [8.50,2.31], [8.17,2.48], [8.23,2.32], [8.09,2.32], [8.18,2.47],
     [7.99,2.48], [8.07,2.55], [7.87,2.51], [7.96,2.57], [7.90,2.61], [7.69,2.46],
     [7.76,2.40], [8.02,2.45], [8.00,2.35], [7.47,2.11], [7.18,2.26], [6.76,2.74],
     [6.92,3.03], [6.91,2.93], [7.00,2.87], [7.07,3.06], [7.31,2.81], [7.22,2.75],
     [7.37,2.45], [7.50,2.48], [7.31,2.67], [7.37,2.82], [7.59,2.80], [7.39,2.83],
     [7.50,2.94], [7.30,2.94], [7.38,3.01], [7.28,3.09], [6.86,3.01], [6.87,3.11],
     [6.67,3.25], [6.67,3.74], [6.67,3.55], [6.98,3.57], [7.13,3.69], [7.13,3.92],
     [7.25,3.65], [7.15,3.65], [7.09,3.52], [7.36,3.90], [7.46,3.85], [7.40,3.76],
     [7.56,3.74], [7.59,3.48], [7.72,3.44], [7.76,3.54], [7.95,3.40], [7.92,3.54],
     [8.07,3.68]],
    [[2.58,4.14], [2.28,3.77], [2.28,3.60], [2.35,3.57], [2.29,3.40], [2.44,3.46],
     [2.40,3.35], [2.14,3.20], [2.16,3.09], [2.02,3.01], [2.01,2.90], [1.46,2.80],
     [1.40,2.89], [1.37,2.81], [0.95,3.15], [0.95,3.07], [1.20,2.87], [0.83,2.81],
     [0.85,2.69], [1.06,2.56], [0.72,2.48], [0.96,2.47], [0.89,2.34], [0.76,2.31],
     [1.00,2.27], [1.06,2.16], [1.26,2.16], [1.78,2.24], [1.94,2.37], [1.96,2.24],
     [2.19,2.24], [2.16,2.16], [2.58,2.31]],
    [[5.96,1.40], [5.89,1.62], [5.96,1.77], [5.85,1.78], [5.93,1.83], [5.83,1.98],
     [5.67,1.98], [5.85,2.11], [5.66,2.11], [5.63,2.23], [5.81,2.24], [5.46,2.33],
     [5.37,2.48], [5.19,2.48], [5.17,2.87], [4.90,2.78], [4.77,2.48], [4.85,2.40],
     [4.74,2.34], [4.91,2.21], [4.80,2.29], [4.70,2.23], [4.87,2.20], [4.84,2.12],
     [4.66,2.12], [4.81,2.01], [4.68,1.99], [4.62,1.87], [4.62,1.33], [4.79,1.34],
     [4.80,1.42], [4.93,1.39], [4.79,1.24], [5.13,1.45], [5.00,1.33], [5.15,1.33],
     [5.03,1.22], [5.66,1.15], [6.02,1.32], [5.41,1.27], [5.64,1.41], [5.66,1.27],
     [5.74,1.40], [5.81,1.34]],
    [[9.88,1.82], [9.88,4.57], [9.73,4.47], [9.56,4.57], [9.46,4.66], [9.42,4.67],
     [9.42,2.11], [9.46,2.21], [9.62,2.39], [9.62,2.30], [9.53,2.21], [9.56,2.11],
     [9.45,2.06], [9.41,1.99], [9.56,1.97], [9.64,1.89]],
    [[10.85,2.03], [10.85,3.71], [10.70,3.71], [10.68,3.80], [10.80,3.80], [10.72,3.94],
     [10.82,4.05], [10.87,4.16], [10.77,4.25], [10.76,4.33], [10.66,4.37], [10.56,4.46],
     [10.44,4.46], [10.44,1.98], [10.49,1.97], [10.66,2.00], [10.74,1.99]],
    [[11.46,5.87], [11.52,5.87], [11.52,5.65], [11.61,5.64], [11.65,5.85], [11.95,6.17],
     [11.76,6.76], [11.52,6.70], [11.39,6.42], [11.35,6.50], [11.40,6.52], [11.17,6.38],
     [10.92,6.39], [10.89,6.51], [10.61,6.51], [10.67,6.47], [10.54,6.08], [10.58,5.98],
     [10.84,5.95], [10.82,5.85], [11.01,5.85], [10.91,5.73], [11.15,5.79], [11.11,5.71],
     [11.22,5.71], [11.20,5.64], [11.33,5.72], [11.37,5.86]],
    [[10.44,1.78], [10.44,1.90], [10.29,1.96], [10.21,2.08], [10.29,2.10], [10.39,1.99],
     [10.44,1.98], [10.44,4.46], [10.27,4.52], [10.41,4.81], [10.35,4.84], [10.27,4.98],
     [10.14,4.85], [10.15,1.68], [10.33,1.67], [10.21,1.77]],
    [[10.14,1.80], [10.14,4.85], [10.06,4.79], [10.06,4.96], [10.14,5.02], [10.14,5.15],
     [10.07,5.12], [10.09,5.01], [9.97,4.92], [10.06,4.75], [10.01,4.70], [9.87,4.70],
     [9.90,4.62], [9.88,4.57], [9.88,1.82], [9.95,1.81], [9.96,1.85], [10.04,1.81],
     [10.12,1.86]],
    [[2.97,4.47], [2.95,4.42], [2.73,4.23], [2.84,4.39], [2.76,4.42], [2.68,4.29],
     [2.68,2.37], [2.71,2.39], [2.75,2.38], [2.87,2.42], [2.86,2.31], [2.97,2.31]],
    [[9.42,2.11], [9.42,4.67], [9.40,4.70], [9.41,4.78], [9.38,4.89], [9.35,4.94],
     [9.28,4.93], [9.21,4.81], [9.21,2.36], [9.26,2.36], [9.31,2.43], [9.43,2.43],
     [9.37,2.34], [9.27,2.29], [9.21,2.29], [9.21,2.07], [9.24,2.07], [9.16,2.21],
     [9.39,2.10], [9.30,2.11], [9.38,2.09]],
    [[3.29,2.36], [3.39,2.35], [3.29,2.24], [3.29,2.17], [3.36,2.10], [3.42,2.19],
     [3.37,2.25], [3.50,2.24], [3.46,2.35], [3.50,2.32], [3.49,2.66], [3.38,2.67],
     [3.40,2.77], [3.34,2.81], [3.39,2.93], [3.50,2.97], [3.50,4.16], [3.44,4.19],
     [3.37,4.15], [3.29,4.23]],
    [[12.24,2.28], [12.14,2.28], [12.10,2.16], [11.88,2.17], [11.77,2.07], [11.75,2.08],
     [11.73,2.10], [11.73,2.84], [11.79,2.84], [11.84,2.82], [11.95,2.88], [11.99,2.80],
     [12.09,2.65], [12.15,2.67], [12.18,2.73], [12.24,2.72]],
    [[3.29,4.67], [3.24,4.66], [3.16,4.63], [3.16,2.36], [3.29,2.36], [3.29,4.23],
     [3.24,4.29], [3.20,4.41], [3.29,4.49]],
];

/* ------------------------------------------------------------ slide bodies */

function slide01(s) {
    oval(s, 1.615, 3.413, 1.631, GREEN);
    tx(s, 'Lemida', 2.323, 3.68, 7.677, 2.423, { fontSize: 138 });
    oval(s, 1.369, 5.546, 0.754, BLUE);
    tx(s, 'Marketing Strategy Presentation', 2.431, 6.118, 6.014, 0.505, { fontSize: 24 });
    logoMark(s, 9.036, 3.84, 0.787, 0.777, GREEN, SKY);
    wordmark(s, 0.959, 0.84);
    navLinks(s, [['Presentation', 4.264, 0.773], ['Marketing', 7.472, 0.773],
        ['Business Stategy', 10.417, 0.773]], 'left', 20);
    oval(s, 17.629, 1.785, 0.778, GREEN);
    oval(s, 11.982, 9.411, 0.754, SKY);
    oval(s, 18.177, 8.351, 1.038, PURPLE);
    para(s, LOREM_URNA, 1.293, 8.351, 8.873, 1.275, { fontSize: 16 });
    // Blue outline "stadium" behind the artwork panel.
    s.addShape('roundRect', {
        x: 11.905, y: 2.193, w: 6.19, h: 7.44, rectRadius: 0.5 * 6.19,
        fill: { type: 'none' }, line: { color: BLUE, width: 23.25 },
    });
    // Picture placeholder of the title layout (prompt text kept).
    s.addText('IMG PNG FILE', {
        shape: 'rect', x: 12.063, y: 2.318, w: 5.873, h: 7.238, fill: { color: GREEN },
        align: 'center', valign: 'top', fontFace: HEAD, fontSize: 18, color: W,
    });
}

function slide02(s) {
    rr(s, 0.569, 0.46, 3.19, 3.927, GREEN);
    rr(s, 5.337, 6.516, 6.853, 3.927, BLUE);
    wordmark(s, 17.34, 0.84);
    navLinks(s, [['Presentation', 7.918, 0.807], ['Marketing', 11.023, 0.807],
        ['Business Stategy', 13.625, 0.807]]);
    eyebrow(s, 7.918, 2.086);
    tx(s, 'Our About Lemida Business Marketing Strategy Presentation', 7.918, 2.733, 11.132, 1.447, { fontSize: 40 });
    para(s, LOREM, 7.916, 4.634, 11.132, 0.774);
    tx(s, 'Business Marketing Strategy', 6.132, 7.048, 5.568, 1.447, { fontSize: 40 });
    para(s, LOREM_ER, 6.081, 8.705, 5.619, 1.128, { color: T95 });
    tx(s, QUOTE_BEST, 13.431, 7.012, 5.417, 1.043, { fontSize: 28 });
    para(s, LOREM_ER, 13.431, 8.454, 5.619, 1.128);
}

function slide03(s) {
    wordmark(s, 0.959, 0.84);
    navLinks(s, [['Presentation', 4.16, 0.807], ['Marketing', 7.368, 0.807],
        ['Business Stategy', 10.312, 0.807]]);
    eyebrow(s, 1.175, 2.655);
    tx(s, 'Our About Lemida Business Marketing Strategy', 1.175, 3.303, 5.952, 2.322, { fontSize: 44 });
    para(s, LOREM, 1.175, 6.126, 5.952, 1.481);
    para(s, LOREM, 1.175, 8.109, 5.952, 1.481);
    [['1', BLUE, 2.647, 'Our Section About One'],
     ['2', GREEN, 5.107, 'Our Section  About Two'],
     ['3', PURPLE, 7.984, 'Our Section About Three']].forEach(([n, col, y, title]) => {
        s.addText(n, {
            shape: 'roundRect', rectRadius: 0.1222 * 0.518, x: 8.42, y, w: 0.578, h: 0.518,
            fill: { color: col }, align: 'center', valign: 'middle',
            fontFace: BODY_XB, fontSize: 20, color: W,
        });
        tx(s, title, 9.301, y - 0.031, 3.627, 0.438, { fontSize: 20 });
        para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, ', 9.295, y + 0.56, 4.428, 1.128);
    });
}

function slide04(s) {
    tx(s, 'The best way to predict the future is to create it', 1.321, 7.636, 7.317, 2.524, { fontSize: 48 });
    eyebrow(s, 1.321, 6.767);
    rr(s, 0.881, 2.309, 6.853, 3.927, BLUE);
    tx(s, 'Business Marketing Strategy', 1.676, 2.84, 5.568, 1.447, { fontSize: 40 });
    para(s, LOREM_ER, 1.625, 4.497, 5.619, 1.128, { color: T95 });
    tx(s, 'Marketing is really just about sharing your passion.', 9.854, 5.452, 9.119, 1.313, { fontSize: 36 });
    tx(s, 'Our Section Business About', 9.854, 7.636, 4.317, 0.438, { fontSize: 20 });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adip iscing elit. Maecenas porttitor congue massare Fusce posuere, magna sed pulvinar ultriciesw, purus lectus malesuada libero', 9.854, 8.309, 5.079, 1.481);
    rr(s, 15.219, 7.204, 3.46, 2.618, CARD);
    tx(s, '763+', 15.616, 7.586, 2.667, 1.313, { fontSize: 72, fontFace: BODY_XB, color: GREEN, align: 'center' });
    tx(s, 'Sale Marketing', 15.616, 8.91, 2.667, 0.505, { fontSize: 24, align: 'center' });
}

function slide05(s) {
    // Colour blocks that come from the "2_Title Slide" layout.
    rr(s, 2.572, 1.002, 3.46, 6.21, PURPLE);
    rr(s, 9.201, 6.991, 3.46, 2.618, SKY);
    rr(s, 15.541, 6.991, 3.46, 2.618, GREEN);
    wordmark(s, 17.34, 0.84);
    navLinks(s, [['Presentation', 7.09, 0.807], ['Marketing', 10.196, 0.807],
        ['Business Stategy', 12.797, 0.807]]);
    eyebrow(s, 7.09, 2.05);
    tx(s, 'The aim of marketing is to make selling superfluous.', 7.09, 2.919, 10.417, 1.717, { fontSize: 48 });
    para(s, LOREM, 7.09, 5.228, 11.987, 0.774);
    tx(s, QUOTE_BEST, 0.831, 7.778, 5.417, 1.043, { fontSize: 28 });
    para(s, LOREM_ER, 0.831, 9.093, 5.619, 1.128);
}

function slide06(s) {
    header(s);
    eyebrow(s, 1.325, 2.557);
    tx(s, "Marketing is a contest for people's attention.", 1.325, 3.427, 8.833, 1.717, { fontSize: 48 });
    para(s, LOREM, 1.325, 5.891, 8.46, 1.128);
    para(s, LOREM, 1.325, 7.767, 8.46, 1.128);
    rr(s, 14.848, 2.185, 3.836, 4.509, BLUE);
    rr(s, 11.025, 4.025, 3.144, 3.2, GREEN);
    rr(s, 14.587, 7.294, 3.29, 2.614, PURPLE);
    tx(s, QUOTE_CULTURE, 15.31, 3.011, 3.057, 1.043, { fontSize: 28 });
    para(s, LOREM_SHORT, 15.31, 4.352, 3.161, 0.774);
    playButton(s, 17.59, 5.373, BLUE);
    tx(s, CARD_TITLE, 11.425, 4.672, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_SHORT, 11.468, 5.703, 2.701, 0.774);
    tx(s, CARD_TITLE, 15.04, 7.622, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_SHORT, 15.083, 8.654, 2.701, 0.774);
}

function slide07(s) {
    // Two icon cards from the "5_Title Slide" layout.
    rr(s, 5.647, 1.029, 3.506, 4.171, SKY);
    rr(s, 5.647, 6.05, 3.506, 4.171, BLUE);
    wordmark(s, 17.104, 1.062);
    navLinks(s, [['Marketing', 10.666, 1.015], ['Business Stategy', 13.304, 1.029]]);
    eyebrow(s, 10.666, 2.241);
    tx(s, 'Marketing is really just about sharing your passion.', 10.666, 3.11, 7.674, 2.524, { fontSize: 48 });
    para(s, LOREM, 10.678, 6.31, 7.674, 1.128);
    pill(s, '   Your Text Here…', 10.678, 8.412, 7.575, 0.778, W, { align: 'left', fontFace: BODY, color: T65 });
    pill(s, 'More', 16.956, 8.506, 1.13, 0.595, GREEN);
    [[1.737, SKY], [6.746, BLUE]].forEach(([y, mark]) => {
        iconTile(s, 6.273, y, 0.711, 0.595, W, mark, 0.441);
        tx(s, CARD_TITLE, 6.156, y + 0.868, 2.482, 0.909, { fontSize: 24 });
        para(s, LOREM_SHORT, 6.198, y + 1.899, 2.701, 0.774);
    });
}

function slide08(s) {
    // Green band from the "4_Title Slide" layout (rotated round-corner block).
    s.addShape('round2SameRect', { x: 14.505, y: 3.552, w: 4.409, h: 6.58, fill: { color: GREEN }, rotate: 270 });
    tx(s, QUOTE_BEST, 1.317, 7.364, 7.064, 2.524, { fontSize: 48 });
    eyebrow(s, 1.317, 6.495);
    wordmark(s, 17.324, 0.839);
    navLinks(s, [['Presentation', 8.106, 0.807], ['Marketing', 11.212, 0.807],
        ['Business Stategy', 13.813, 0.807]]);
    [8.265, 14.129].forEach((x) => {
        iconTile(s, x, 2.227, 0.711, 0.595, SKY, W, 0.441);
        tx(s, 'Our Section Business', x + 0.897, 2.227, 3.56, 0.438, { fontSize: 20 });
        para(s, LOREM_SECT, x + 0.897, 2.9, 4.1, 1.128);
    });
    rr(s, 8.265, 5.201, 3.506, 4.171, BLUE);
    iconTile(s, 8.993, 5.945, 0.711, 0.595, W, BLUE, 0.441);
    tx(s, CARD_TITLE, 8.875, 6.812, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_SHORT, 8.918, 7.844, 2.701, 0.774);
}

function slide09(s) {
    oval(s, 18.307, 2.51, 0.778, GREEN);
    oval(s, 11.77, 9.755, 0.754, SKY);
    oval(s, 18.177, 8.351, 1.038, PURPLE);
    wordmark(s, 17.301, 0.84);
    navLinks(s, [['Presentation', 4.218, 0.807], ['Marketing', 7.338, 0.807],
        ['Business Stategy', 10.282, 0.807]]);
    eyebrow(s, 4.207, 2.291);
    tx(s, "Marketing is a contest for people's attention.", 4.207, 3.16, 8.833, 1.717, { fontSize: 48 });
    para(s, LOREM, 4.207, 5.625, 8.46, 1.128);
    rr(s, 13.633, 3.042, 3.836, 5.309, CARD);
    for (let i = 0; i < 5; i++) {
        s.addShape('star5', { x: 14.167 + i * 0.4845, y: 3.746, w: 0.305, h: 0.305, fill: { color: GOLD } });
    }
    tx(s, QUOTE_CULTURE, 14.143, 4.534, 3.057, 1.043, { fontSize: 28 });
    para(s, LOREM_SHORT, 14.143, 6.061, 3.161, 0.774);
    playButton(s, 16.424, 6.956, CARD);
    iconTile(s, 4.238, 7.571, 0.936, 0.784, BLUE, W, 0.43);
    tx(s, 'Our Section Business', 5.522, 7.571, 3.56, 0.438, { fontSize: 20 });
    para(s, LOREM_SECT, 5.522, 8.244, 7.065, 0.774);
}

function slide10(s) {
    s.addShape('round2SameRect', { x: 10.524, y: 5.625, w: 8.234, h: 4.435, fill: { color: BLUE }, flipV: true });
    wordmark(s, 0.959, 0.84);
    navLinks(s, [['Marketing', 3.858, 0.77], ['Business Stategy', 6.495, 0.784]]);
    eyebrow(s, 1.325, 2.557);
    tx(s, "The customer's perception is your reality.", 1.325, 3.427, 7.247, 2.827, { fontSize: 54 });
    para(s, LOREM, 1.325, 6.849, 6.941, 1.128);
    pill(s, 'Read More', 1.364, 8.693, 2.017, 0.65, GREEN);
    pill(s, 'Learn More', 3.786, 8.693, 2.017, 0.65, PURPLE);
    tx(s, QUOTE_BEST, 11.241, 6.536, 6.593, 1.043, { fontSize: 28 });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultrici es, purus lectus malesuada libero, sit amet', 11.241, 7.977, 6.838, 1.128);
}

function slide11(s) {
    s.addShape('round2DiagRect', { x: 3.011, y: 3.741, w: 3.506, h: 4.171, fill: { color: BLUE } });
    iconTile(s, 3.739, 4.484, 0.711, 0.595, W, BLUE, 0.441);
    tx(s, CARD_TITLE, 3.621, 5.352, 2.482, 0.909, { fontSize: 24 });
    para(s, LOREM_SHORT, 3.664, 6.384, 2.701, 0.774);
    iconTile(s, 1.028, 7.607, 1.131, 0.947, GREEN, W, 0.52);
    iconTile(s, 7.075, 2.448, 1.131, 0.947, PURPLE, W, 0.52);
    wordmark(s, 17.324, 0.839);
    navLinks(s, [['Presentation', 8.106, 0.807], ['Marketing', 11.212, 0.807],
        ['Business Stategy', 13.813, 0.807]]);
    eyebrow(s, 9.531, 2.443);
    tx(s, 'The aim of marketing is to make selling superfluous.', 9.531, 3.312, 8.773, 1.717, { fontSize: 48 });
    para(s, LOREM, 9.542, 5.897, 8.773, 1.128);
    [9.616, 14.291].forEach((x) => {
        iconTile(s, x, 7.777, 0.711, 0.595, SKY, W, 0.441);
        tx(s, 'Section Business', x + 0.897, 7.777, 2.678, 0.438, { fontSize: 20 });
        para(s, 'Lorem ipsum dolor sit amet, consect etuer adip iscing et. Maecenas', x + 0.897, 8.45, 3.128, 1.128);
    });
}

function slide12(s) {
    oval(s, 3.584, 3.751, 1.631, GREEN);
    tx(s, 'Break Slides', 4.026, 4.018, 11.947, 2.423, { fontSize: 138, align: 'center' });
    oval(s, 2.913, 8.55, 1.305, BLUE);
    tx(s, 'Marketing Strategy Presentation', 6.993, 6.709, 6.014, 0.505, { fontSize: 24, align: 'center' });
    logoMark(s, 15.388, 3.003, 2.057, 2.031, GREEN, SKY);
    header(s);
    oval(s, 18.15, 7.754, 0.797, PURPLE);
    oval(s, 1.041, 3.608, 0.821, PURPLE);
    oval(s, 13.997, 7.561, 0.591, GREEN);
    para(s, LOREM, 3.254, 8.815, 13.492, 0.774, { align: 'center' });
    iconTile(s, 9.435, 2.661, 1.131, 0.947, PURPLE, W, 0.52);
}

function slide13(s) {
    rr(s, 1.143, 1.434, 8.778, 8.381, CARD, 0.05494);
    tx(s, 'Nov 2023 08:27:12', 2.052, 2.474, 2.67, 0.438, { fontSize: 20, fontFace: BODY_SB });
    s.addShape('triangle', { x: 5.167, y: 2.501, w: 0.282, h: 0.243, fill: { color: GREEN } });
    tx(s, '1.87%', 5.532, 2.474, 1.232, 0.438, { fontSize: 20, fontFace: BODY_XB, color: GREEN });
    // Sky-blue volatile series (poly-line, mirrored exactly as in the original).
    const spark = [[0, 0.198], [0.036, 0.001], [0.086, 0.269], [0.118, 0.198], [0.154, 0.245],
        [0.177, 0.037], [0.237, 0.43], [0.261, 0.388], [0.297, 0.567], [0.342, 0.484],
        [0.38, 0.758], [0.41, 0.358], [0.47, 0.71], [0.517, 0.579], [0.56, 0.728],
        [0.592, 0.65], [0.641, 0.793], [0.692, 0.71], [0.735, 0.799], [0.767, 0.746],
        [0.825, 0.859], [0.84, 0.817], [0.874, 0.913], [0.9, 0.716], [0.947, 0.996],
        [0.968, 0.889], [1, 0.99]];
    s.addShape('custGeom', {
        x: 2.052, y: 3.557, w: 6.547, h: 2.331, flipH: true, line: { color: SKY, width: 2.25 },
        points: spark.map(([px, py]) => ({ x: px * 6.547, y: py * 2.331 })),
    });
    // Smooth purple trend line.
    const cw = 6.334, ch = 2.562;
    s.addShape('custGeom', {
        x: 1.954, y: 2.811, w: cw, h: ch, line: { color: PURPLE, width: 1 },
        points: [
            { x: 0, y: ch },
            { x: 0.1744 * cw, y: 0.5621 * ch, curve: { type: 'cubic', x1: 0.0479 * cw, y1: 0.8099 * ch, x2: 0.0957 * cw, y2: 0.6198 * ch } },
            { x: 0.4718 * cw, y: 0.6536 * ch, curve: { type: 'cubic', x1: 0.253 * cw, y1: 0.5044 * ch, x2: 0.3765 * cw, y2: 0.7157 * ch } },
            { x: 0.7462 * cw, y: 0.1895 * ch, curve: { type: 'cubic', x1: 0.5671 * cw, y1: 0.5915 * ch, x2: 0.6581 * cw, y2: 0.2985 * ch } },
            { x: cw, y: 0, curve: { type: 'cubic', x1: 0.8342 * cw, y1: 0.0806 * ch, x2: 0.9547 * cw, y2: 0.0338 * ch } },
        ],
    });
    oval(s, 8.502, 3.855, 0.194, GREEN);
    tx(s, '$32.191,40', 6.059, 5.779, 2.67, 0.64, { fontSize: 32, fontFace: BODY_SB, color: GREEN });
    para(s, 'Lorem ipsum dolor sit amet, cons ectetuer adipiscing elit. Maecena porttitor congue massa. Fusce posuere, magna', 6.0, 6.741, 3.714, 1.481);
    tx(s, '$29.191,40', 1.887, 6.96, 2.67, 0.64, { fontSize: 32, fontFace: BODY_SB });
    tx(s, 'Des 2023 08:27:12', 1.929, 7.666, 2.67, 0.438, { fontSize: 20, fontFace: BODY_SB });
    iconTile(s, 8.386, 2.001, 1.131, 0.947, SKY, W, 0.52);
    wordmark(s, 17.324, 1.405);
    navLinks(s, [['Marketing', 11.212, 1.361], ['Business Stategy', 13.813, 1.373]]);
    eyebrow(s, 11.212, 2.734);
    tx(s, 'Sales Marketing Report', 11.212, 3.603, 7.645, 2.322, { fontSize: 66 });
    para(s, LOREM, 11.267, 6.625, 6.941, 1.128);
    pill(s, 'Download ', 11.306, 8.51, 2.192, 0.665, GREEN);
    pill(s, 'View Report', 14.077, 8.51, 2.192, 0.665, PURPLE);
}

function slide14(s) {
    header(s);
    rr(s, 0.89, 2.561, 3.836, 7.444, BLUE);
    iconTile(s, 1.427, 3.332, 0.711, 0.595, W, BLUE, 0.441);
    tx(s, QUOTE_CULTURE, 1.352, 4.392, 3.057, 1.043, { fontSize: 28 });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, mar gna sed pulvinar ultriciesas purus lectus', 1.352, 5.733, 3.057, 2.188, { align: 'justify' });
    playButton(s, 3.616, 8.318, BLUE);
    reviewCard(s, 5.605);
    colourCard(s, 5.645, 7.245, GREEN);
    eyebrow(s, 10.99, 2.71);
    tx(s, 'Our Services Lemida Business Marketing Strategy', 10.99, 3.422, 7.261, 2.524, { fontSize: 48 });
    para(s, LOREM_URNA, 10.99, 6.544, 7.658, 1.128);
    tx(s, 'Content Here', 10.99, 8.29, 3.018, 0.438, { fontSize: 20 });
    para(s, 'Lorem ipsum dolor sit amet, conse etuer adipiscing elit. Maecenas', 10.99, 8.84, 3.943, 0.774);
    [['Lorem Ipsum Dolor', 8.397, 2.74], ['Fusce posuere, magna ', 9.087, 3.174]].forEach(([label, y, w]) => {
        check(s, 15.174, y, 0.298);
        tx(s, label, 15.474, y - 0.059, w, 0.404);
    });
}

function slide15(s) {
    header(s);
    eyebrow(s, 1.192, 2.71);
    tx(s, 'Our Services Lemida Business Marketing Strategy', 1.192, 3.422, 7.348, 2.524, { fontSize: 48 });
    para(s, LOREM_URNA, 1.191, 6.703, 7.658, 1.128);
    pill(s, '   Your Text Here…', 1.191, 8.54, 7.575, 0.778, W, { align: 'left', fontFace: BODY, color: T65 });
    pill(s, 'More', 7.469, 8.634, 1.13, 0.595, GREEN);
    rr(s, 10.162, 2.44, 8.647, 7.608, CARD);
    [[11.253, 3.301, SKY], [15.177, 3.301, PURPLE], [11.253, 6.355, GREEN], [15.177, 6.355, BLUE]]
        .forEach(([x, y, col]) => {
            iconTile(s, x, y, 0.711, 0.595, col, W, 0.441);
            tx(s, 'Services Business', x - 0.048, y + 1.008, 2.678, 0.438, { fontSize: 20 });
            para(s, LOREM_SVC, x - 0.048, y + 1.681, 3.128, 0.774);
        });
}

function slide16(s) {
    header(s);
    eyebrow(s, 8.086, 1.905, EYEBROW, { align: 'center' });
    tx(s, 'Timeline Marketing Strategy', 4.147, 2.721, 11.706, 1.111, { fontSize: 60, align: 'center' });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation. dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.', 2.777, 4.398, 14.984, 0.774, { align: 'center' });
    rr(s, 1.407, 6.818, 17.186, 0.315, CARD, 0.5);
    [['Your Section One', 1.673, 2.802, 6.375, SKY], ['Your Section Two', 6.124, 7.257, 6.418, GREEN],
     ['Your Section Three', 10.574, 11.714, 6.397, PURPLE], ['Your Section Four', 15.028, 16.162, 6.377, BLUE]]
        .forEach(([label, tX, iX, iY, col]) => {
            iconTile(s, iX, iY, 1.031, 1.031, col, W, 0.597);
            tx(s, label, tX, 8.062, 3.298, 0.438, { fontSize: 20, align: 'center' });
            para(s, 'Lorem ipsum dolor sit ameat, consectetuer adipisc ing elit. ', tX, 8.747, 3.298, 0.774, { align: 'center' });
        });
}

function slide17(s) {
    header(s);
    swotBlock(s, 13.36, 'S', BLUE, GREEN, PURPLE);
    eyebrow(s, 1.192, 2.863);
    tx(s, 'Strengths Analysis Slides', 1.192, 3.576, 6.375, 2.221, { fontSize: 63 });
    para(s, LOREM_SWOT + ', sit amet commod o magna eros quis urna.', 1.191, 6.31, 6.375, 1.481);
    reviewCard(s, 8.791);
    colourCard(s, 8.83, 7.245, GREEN);
    [['Lorem Ipsum Dolor', 1.364, 8.47, 1.664, 8.41, 2.74],
     ['Fusce posuere, ', 1.364, 9.159, 1.664, 9.1, 2.351],
     ['Fusce posuere, ', 4.539, 8.424, 4.84, 8.364, 2.351],
     ['Lorem Ipsum Dolor', 4.511, 9.159, 4.811, 9.1, 2.74]].forEach(([label, cx, cy, tX, tY, w]) => {
        check(s, cx, cy, 0.298);
        tx(s, label, tX, tY, w, 0.404);
    });
}

function slide18(s) {
    header(s);
    swotBlock(s, 0.809, 'W', GREEN, SKY, PURPLE);
    reviewCard(s, 7.373);
    colourCard(s, 7.413, 7.245, BLUE);
    eyebrow(s, 12.474, 2.863);
    tx(s, 'Weaknessess Analysis Slides', 12.474, 3.576, 6.375, 2.221, { fontSize: 63 });
    para(s, LOREM_SWOT, 12.473, 6.31, 6.375, 1.128);
    iconTile(s, 12.616, 8.248, 0.711, 0.595, PURPLE, W, 0.441);
    tx(s, 'Section Business', 13.513, 8.129, 2.678, 0.438, { fontSize: 20 });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscin elit. Maec enas porttitor congue massa. ', 13.513, 8.803, 5.334, 0.774);
}

function slide19(s) {
    header(s);
    swotBlock(s, 13.36, 'O', PURPLE, GREEN, SKY);
    reviewCard(s, 8.791);
    colourCard(s, 8.83, 7.245, BLUE);
    eyebrow(s, 1.192, 2.863);
    tx(s, 'Opportunities Analysis Slides', 1.192, 3.576, 6.375, 2.221, { fontSize: 63 });
    // Two progress bars.
    [['Progress One', 6.313, 6.762, 6.568, '80%', 7.449, 5.019, PURPLE],
     ['Progress Two', 8.198, 8.647, 8.453, '90%', 9.385, 5.505, GREEN]]
        .forEach(([label, ly, by, py, pct, barY, fillW, col]) => {
            tx(s, label, 1.264, ly, 1.902, 0.404);
            para(s, 'tempor incididunt ut labore', 1.264, by, 3.426, 0.421, { align: 'justify' });
            tx(s, pct, 6.785, py, 0.8, 0.404);
            s.addShape('roundRect', {
                x: 1.352, y: barY, w: 6.149, h: 0.257, rectRadius: 0.1285,
                fill: { color: CARD }, line: { color: col, width: 3.25 },
            });
            rr(s, 1.371, barY, fillW, 0.257, col, 0.5);
        });
}

function slide20(s) {
    header(s);
    swotBlock(s, 0.809, 'T', SKY, GREEN, PURPLE);
    reviewCard(s, 7.373);
    colourCard(s, 7.413, 7.245, PURPLE);
    eyebrow(s, 12.474, 2.863);
    tx(s, 'Threatss Analysis Slides', 12.474, 3.576, 6.375, 2.221, { fontSize: 63 });
    para(s, LOREM_SWOT, 12.473, 6.31, 6.375, 1.128);
    para(s, LOREM_SWOT, 12.473, 8.017, 6.375, 1.128);
}

function slide21(s) {
    header(s);
    eyebrow(s, 8.086, 1.905, EYEBROW, { align: 'center' });
    tx(s, 'Our Timeline Marketing Strategy', 2.755, 2.57, 14.489, 1.111, { fontSize: 60, align: 'center' });
    // Spine and drop lines.
    s.addShape('line', { x: 0.918, y: 7.351, w: 18.164, h: 0, line: { color: GREY_DK, width: 6 } });
    [[4.551, 5.895], [11.816, 5.895], [8.184, 7.677], [15.449, 7.677]].forEach(([x, y]) =>
        s.addShape('line', { x, y, w: 0, h: 1.13, line: { color: GREY, width: 2.25 } }));
    // Milestone dots.
    [[0.593, SKY], [4.225, SKY], [7.858, GREEN], [11.491, PURPLE], [15.124, BLUE], [18.757, SKY]]
        .forEach(([x, col]) => oval(s, x, 7.026, 0.651, col));
    // Year chips and their captions.
    [['2023', 3.621, 5.11, SKY], ['2025', 10.881, 5.11, PURPLE],
     ['2024', 7.259, 8.797, GREEN], ['2026', 14.519, 8.797, BLUE]].forEach(([year, x, y, col]) => {
        s.addText(year, {
            shape: 'roundRect', rectRadius: 0.13866 * 0.871, x, y, w: 1.86, h: 0.871,
            fill: { color: col }, align: 'center', valign: 'middle',
            fontFace: BODY_XB, fontSize: 28, color: W,
        });
    });
    [['Section One', 3.16, 8.251, 2.726, 8.899], ['Section Two', 6.92, 4.68, 6.487, 5.329],
     ['Section Three', 14.002, 4.68, 13.569, 5.329], ['Section Four', 10.452, 8.244, 10.018, 8.893]]
        .forEach(([label, tX, tY, bX, bY]) => {
            tx(s, label, tX, tY, 2.699, 0.505, { fontSize: 24, align: 'center' });
            para(s, LOREM_TL, bX, bY, 3.488, 0.774, { align: 'center' });
        });
}

function slide22(s) {
    // Dark "S" connectors that chain the circles together.
    const cw = 1.821, ch = 3.124;
    const link = [
        { x: 0, y: 0 }, { x: cw, y: 0 }, { x: 0.9756 * cw, y: 0.0158 * ch },
        { x: 0.7197 * cw, y: 0.5 * ch, curve: { type: 'cubic', x1: 0.8175 * cw, y1: 0.1397 * ch, x2: 0.7197 * cw, y2: 0.3109 * ch } },
        { x: 0.9756 * cw, y: 0.9842 * ch, curve: { type: 'cubic', x1: 0.7197 * cw, y1: 0.6891 * ch, x2: 0.8175 * cw, y2: 0.8603 * ch } },
        { x: cw, y: ch }, { x: 0, y: ch }, { x: 0.0244 * cw, y: 0.9842 * ch },
        { x: 0.2803 * cw, y: 0.5 * ch, curve: { type: 'cubic', x1: 0.1825 * cw, y1: 0.8603 * ch, x2: 0.2803 * cw, y2: 0.6891 * ch } },
        { x: 0.0244 * cw, y: 0.0158 * ch, curve: { type: 'cubic', x1: 0.2803 * cw, y1: 0.3109 * ch, x2: 0.1825 * cw, y2: 0.1397 * ch } },
        { close: true },
    ];
    [[8.988, 5.46, 297.12], [4.09, 5.508, 240], [14.006, 5.508, 242.18]].forEach(([x, y, rot]) =>
        s.addShape('custGeom', { x, y, w: cw, h: ch, rotate: rot, fill: { color: CARD }, points: link }));
    // Four concentric-ring nodes with their glyphs.
    [[1.248, 6.761, BLUE, 2.436, 7.94], [5.876, 4.142, GREEN, 7.061, 5.331],
     [10.788, 6.761, SKY, 11.969, 7.918], [15.253, 4.324, PURPLE, 16.411, 5.522]]
        .forEach(([x, y, col, gx, gy]) => {
            oval(s, x, y, 3.181, col);
            oval(s, x + 0.743, y + 0.743, 1.694, '000000', { fill: { color: '000000', transparency: 80 } });
            oval(s, x + 0.962, y + 0.962, 1.257, col);
            glyph(s, gx, gy, 0.822, W);
        });
    wordmark(s, 0.959, 0.84);
    wordmark(s, 17.301, 0.84);
    eyebrow(s, 8.086, 1.148, EYEBROW, { align: 'center' });
    tx(s, 'Chain Process Diagram ', 2.755, 1.719, 14.489, 1.313, { fontSize: 72, align: 'center' });
    [['Section One', 1.468, 4.589, 1.094, 5.238], ['Section Three', 10.938, 4.589, 10.565, 5.238],
     ['Section Two', 6.058, 8.231, 5.684, 8.88], ['Section Four', 15.528, 8.231, 15.155, 8.88]]
        .forEach(([label, tX, tY, bX, bY]) => {
            tx(s, label, tX, tY, 2.699, 0.505, { fontSize: 24, align: 'center' });
            para(s, LOREM_TL, bX, bY, 3.488, 0.774, { align: 'center' });
        });
}

function slide23(s) {
    header(s);
    eyebrow(s, 8.086, 1.987, EYEBROW, { align: 'center' });
    tx(s, 'Stastitics diagram', 4.147, 2.525, 11.706, 1.01, { fontSize: 54, align: 'center' });
    // Mountain chart: front face + shaded right face, chip label, year caption.
    const bars = [
        { pct: '82%', col: SKY, x: 1.471, y: 4.504, h: 3.393, chipX: 2.457, chipY: 3.665, year: '2023', yx: 2.716, yw: 1.143 },
        { pct: '46%', col: PURPLE, x: 4.155, y: 6.129, h: 1.768, chipX: 5.093, chipY: 5.217, year: '2024', yx: 5.399, yw: 1.175 },
        { pct: '70%', col: GREEN, x: 6.840, y: 5.472, h: 2.425, chipX: 7.826, chipY: 4.560, year: '2025', yx: 8.077, yw: 1.161 },
        { pct: '31%', col: PURPLE, x: 9.525, y: 6.740, h: 1.157, chipX: 10.498, chipY: 5.873, year: '2026', yx: 10.779, yw: 1.153 },
        { pct: '75%', col: GREEN, x: 12.210, y: 4.827, h: 3.070, chipX: 13.196, chipY: 3.937, year: '2027', yx: 13.507, yw: 1.110 },
        { pct: '72%', col: SKY, x: 14.894, y: 5.196, h: 2.701, chipX: 15.845, chipY: 4.366, year: '2028', yx: 16.180, yw: 1.159 },
    ];
    bars.forEach((b) => {
        s.addShape('triangle', { x: b.x, y: b.y, w: 3.635, h: b.h, fill: { color: b.col } });
        s.addShape('rtTriangle', {
            x: b.x + 1.830, y: b.y + 0.022, w: 1.804, h: b.h - 0.024,
            fill: { color: '000000', transparency: 80 },
        });
        s.addText(b.pct, {
            shape: 'roundRect', rectRadius: 0.16667 * 0.496, x: b.chipX, y: b.chipY, w: 1.663, h: 0.496,
            fill: { color: b.col }, align: 'center', valign: 'middle',
            fontFace: BODY_XB, fontSize: 20, bold: true, color: W,
        });
        s.addShape('triangle', {
            x: b.chipX + 0.509, y: b.chipY + 0.491, w: 0.645, h: 0.165,
            fill: { color: b.col }, rotate: 180,
        });
        tx(s, b.year, b.yx, 8.266, b.yw, 0.505,
            { fontSize: 24, fontFace: BODY_XB, color: T95, align: 'center' });
    });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris', 1.642, 9.377, 16.716, 0.871, { fontSize: 16, align: 'center' });
}

function slide24(s) {
    // World map: continent outlines simplified from the original vector art,
    // given as absolute slide-inch polygons.
    MAP_CONTINENTS.forEach((pts) => s.addShape('custGeom', {
        x: 0, y: 0, w: 20, h: 11.25, fill: { color: CARD },
        points: pts.map(([px, py]) => ({ x: px, y: py })).concat([{ close: true }]),
    }));
    // Location pins - the original uses a radial "glow" gradient, approximated
    // here with three concentric discs of decreasing transparency.
    [[2.671, 2.807, GREEN], [3.779, 4.854, PURPLE], [10.366, 3.045, GREEN], [4.842, 1.426, SKY],
     [7.072, 2.445, PURPLE], [11.106, 5.762, SKY], [7.405, 5.434, SKY]]
        .forEach(([x, y, col]) => {
            [[1.0, 80], [0.68, 45], [0.42, 0]].forEach(([k, alpha]) => {
                const d = 0.583 * k;
                oval(s, x + (0.583 - d) / 2, y + (0.583 - d) / 2, d, col,
                    { fill: { color: col, transparency: alpha } });
            });
        });
    eyebrow(s, 1.251, 6.921);
    tx(s, 'World Map Sales Report', 1.251, 7.634, 6.375, 2.221, { fontSize: 63 });
    tx(s, QUOTE_BEST, 8.109, 7.359, 5.417, 1.043, { fontSize: 28 });
    para(s, LOREM_ER, 8.109, 8.8, 5.619, 1.128);
    [[1.114, 2.122, 2.795, SKY], [4.236, 5.245, 5.918, GREEN], [7.692, 8.7, 9.373, PURPLE]]
        .forEach(([tileY, titleY, bodyY, col], i) => {
            const x = i === 2 ? 15.716 : 15.634;
            const tX = i === 2 ? 15.621 : 15.539;
            rr(s, x, tileY, 0.583, 0.595, col, 0.2066);
            tx(s, 'Section Here', tX, titleY, 2.678, 0.438, { fontSize: 20 });
            para(s, LOREM_SVC, tX, bodyY, 3.128, 0.774);
        });
}

function slide25(s) {
    rr(s, 0.89, 2.143, 18.269, 8.105, CARD);
    tx(s, 'Thank You', 4.026, 4.431, 11.947, 2.423, { fontSize: 138, align: 'center' });
    header(s);
    tx(s, 'For Attention', 6.696, 6.927, 7.144, 1.212, { fontSize: 66, color: GREEN, align: 'center' });
    oval(s, 14.137, 1.584, 1.631, BLUE);
    oval(s, 2.913, 8.55, 1.305, BLUE);
    logoMark(s, 9.427, 3.387, 1.195, 1.18, GREEN, SKY);
    oval(s, 18.15, 7.754, 0.797, PURPLE);
    oval(s, 1.041, 3.608, 0.821, PURPLE);
    oval(s, 14.521, 9.952, 0.591, GREEN);
}

/* -------------------------------------------------------------------- build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
    slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19,
    slide20, slide21, slide22, slide23, slide24, slide25];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LEMIDA_20x11_25', width: 20, height: 11.25 });
pptx.layout = 'LEMIDA_20x11_25';
pptx.author = 'Lemida';
pptx.title = 'Lemida - Marketing Strategy Presentation';

BUILDERS.forEach((build) => {
    const slide = pptx.addSlide();
    slide.background = { color: BG };
    build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0d31aee5-6094-46e3-8295-5501516820d6_grok_final.pptx') })
    .then((f) => console.log('wrote', f));
