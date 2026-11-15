/**
 * "Romance - Wedding Proposal" deck, rebuilt with pptxgenjs.
 *
 * 32 slides, 13.333 x 7.5 in (16:9).  Everything is plain pptxgenjs calls:
 * colours, fonts, boxes and copy live in the tables at the top, the icon
 * library draws the little heart glyphs out of native shapes, and each slide
 * is one small builder function.  Photographs in the original are replaced by
 * grey "[image]" placeholder rectangles.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const PURPLE = 'BD31E3'; // accent1
const PINK = 'F44AA3'; // accent2
const INK = '404040';
const GREY = '595959';
const WHITE = 'FFFFFF';
const MAPGREY = 'D9D9D9';
const PLACEHOLDER = 'E4E4E4';

const SERIF = 'Frank Ruhl Libre';
const SANS = 'Lato';

const NOLINE = { type: 'none' };
const fill = (color) => ({ fill: { color }, line: NOLINE });

/* ------------------------------------------------------- boilerplate copy */

const LEAD =
    'A wedding is a ceremony where two or more people are united in marriage. ' +
    'Most wedding ceremonies involve an exchange of marriage vows';

const T = {
    vows: LEAD + '.',
    gift: LEAD + ' by a couple, presentation of a gift.',
    proclamation: LEAD + ' by a couple, presentation of a gift, and a public proclamation of marriage.',
    proclaim: LEAD + ' by a couple, presentation of a gift, and a public proclamation.',
    authority: LEAD + ' by a couple, presentation of a gift, and a public proclamation of marriage by authority figure or celebrant.',
    anAuthority: LEAD + ' by a couple, presentation of a gift, and a public proclamation of marriage by an authority figure or celebrant.',
    brief: LEAD + ' by a couple, a gift, and a public proclamation of marriage.',
    coupleGift: LEAD + ' by a couple, and a gift.',
    story: 'Wedding is a ceremony of two or more people united in marriage, most of it  involve an exchange of marriage vows by a couple, presentation of a gift.',
    united: 'Wedding is a ceremony of two or more people united in marriage.',
    infographics: 'Our Infographics Collection',
    subtitle: 'Subtitle Here',
    marriage: 'Marriage Story',
};

/* --------------------------------------------------------- text primitives
 * Every text frame in the source deck is a top-anchored auto-fit rectangle
 * with the default PowerPoint insets, so the helpers only vary by role.
 */

function text(slide, str, frame, opt) {
    const o = opt || {};
    slide.addText(str, {
        x: frame[0], y: frame[1], w: frame[2], h: frame[3],
        fontFace: o.face || SANS,
        fontSize: o.size || 18,
        color: o.color || GREY,
        bold: !!o.bold,
        italic: !!o.italic,
        charSpacing: o.spc,
        align: o.align || 'left',
        valign: 'top',
        lineSpacingMultiple: o.lnSpc,
        wrap: true,
    });
}

/** 28pt slide headline. */
function heading(slide, str, box, opt) {
    const o = opt || {};
    text(slide, str, box, { face: SERIF, size: 28, bold: true, color: o.color || PURPLE, lnSpc: 1.1, align: o.align });
}

/** 18pt bold serif sub-head. */
function subhead(slide, str, box, opt) {
    const o = opt || {};
    text(slide, str, box, { face: SERIF, bold: true, color: o.color || PURPLE, spc: 0.5, lnSpc: 1.2, align: o.align });
}

/** 11pt Lato paragraph on double leading. */
function para(slide, str, box, opt) {
    const o = opt || {};
    text(slide, str, box, { size: 11, color: o.color || GREY, spc: 0.3, lnSpc: 2, align: o.align });
}

/** Sub-head + paragraph pair, the deck's most common building block. */
function note(slide, head, bodyText, headBox, bodyBox, opt) {
    subhead(slide, head, headBox, opt);
    para(slide, bodyText, bodyBox, opt);
}

function box(slide, b, color) {
    slide.addShape('rect', { x: b[0], y: b[1], w: b[2], h: b[3], ...fill(color || PURPLE) });
}

/** Purple tag in the top right corner carrying a white heart. */
function badge(slide) {
    box(slide, [11.8, 0, 1, 1.043]);
    slide.addShape('heart', { x: 11.995, y: 0.23, w: 0.584, h: 0.584, ...fill(WHITE) });
}

/** Rounded pill with centred white serif caption. */
function pill(slide, str, b, textBox, opt) {
    const o = opt || {};
    slide.addShape('roundRect', { x: b[0], y: b[1], w: b[2], h: b[3], rectRadius: 0.1, ...fill(PURPLE) });
    text(slide, str, textBox, { face: SERIF, size: o.size || 18, bold: true, color: WHITE, spc: o.spc, align: 'center', lnSpc: o.lnSpc });
}

/**
 * Stand-in for one of the two photographs embedded in the original deck.
 * (The remaining picture frames in the source are empty layout placeholders,
 * so they stay blank here too.)
 */
function imagePlaceholder(slide, b) {
    slide.addShape('rect', { x: b[0], y: b[1], w: b[2], h: b[3], fill: { color: PLACEHOLDER }, line: { color: 'C8C8C8', width: 1 } });
    text(slide, '[image]', [Math.max(b[0], 0.1), b[1] + b[3] / 2 - 0.2, b[2], 0.4], { size: 12, color: '9A9A9A', align: 'center' });
}

/* ------------------------------------------------------------ icon library
 * The source deck uses a set of flat "heart" pictograms.  Each one is rebuilt
 * here from native shapes inside the unit box of the icon.
 */

function icon(slide, kind, b, color) {
    const [X, Y, W, H] = b;
    const tint = color === WHITE ? PURPLE : WHITE; // outline / detail colour
    const at = (fx, fy, fw, fh) => ({ x: X + fx * W, y: Y + fy * H, w: fw * W, h: fh * H });
    const shape = (name, fx, fy, fw, fh, c, extra) =>
        slide.addShape(name, { ...at(fx, fy, fw, fh), ...fill(c === undefined ? color : c), ...(extra || {}) });
    /** Outlined heart: a tinted heart with the body colour punched out of it. */
    const heart = (cx, cy, rel) => {
        const s = rel * W;
        slide.addShape('heart', { x: X + cx * W - s / 2, y: Y + cy * H - s / 2, w: s, h: s, ...fill(tint) });
        const i = s * 0.54;
        slide.addShape('heart', { x: X + cx * W - i / 2, y: Y + cy * H - i / 2 + s * 0.05, w: i, h: i, ...fill(color) });
    };
    /** Straight stroke between two unit-box points (pptxgenjs lines cannot run right-to-left). */
    const stroke = (x1, y1, x2, y2, c, weight) => slide.addShape('line', {
        x: X + Math.min(x1, x2) * W, y: Y + y1 * H,
        w: Math.abs(x2 - x1) * W, h: (y2 - y1) * H,
        flipH: x2 < x1,
        line: { color: c, width: (weight || 0.06) * W * 72 },
    });

    switch (kind) {
        case 'camera':
            shape('rect', 0.2, 0, 0.34, 0.22);
            shape('roundRect', 0, 0.14, 1, 0.86, color, { rectRadius: 0.09 });
            shape('ellipse', 0.83, 0.3, 0.1, 0.12, tint);
            heart(0.46, 0.62, 0.44);
            break;
        case 'gift':
            shape('donut', 0.12, 0, 0.4, 0.3, color, { rectRadius: 0 });
            shape('donut', 0.48, 0, 0.4, 0.3);
            shape('roundRect', 0.02, 0.24, 0.96, 0.2, color, { rectRadius: 0.2 });
            shape('rect', 0.09, 0.44, 0.82, 0.56);
            heart(0.5, 0.72, 0.4);
            break;
        case 'house':
            shape('triangle', 0, 0, 1, 0.55);
            shape('rect', 0.17, 0.4, 0.66, 0.6);
            heart(0.5, 0.68, 0.34);
            break;
        case 'envelope':
            shape('roundRect', 0, 0.04, 1, 0.96, color, { rectRadius: 0.08 });
            stroke(0.05, 0.14, 0.5, 0.52, tint); // flap
            stroke(0.95, 0.14, 0.5, 0.52, tint);
            heart(0.5, 0.52, 0.4);
            break;
        case 'calendar':
            shape('roundRect', 0.2, 0, 0.09, 0.18, color, { rectRadius: 0.5 });
            shape('roundRect', 0.71, 0, 0.09, 0.18, color, { rectRadius: 0.5 });
            shape('rect', 0, 0.08, 1, 0.22);
            shape('rect', 0, 0.28, 1, 0.72);
            heart(0.5, 0.64, 0.38);
            break;
        case 'rings': {
            // two interlocking bands with a solitaire above the right one
            const band = (fx) => {
                shape('ellipse', fx, 0.36, 0.6, 0.6, tint);
                shape('ellipse', fx + 0.115, 0.475, 0.37, 0.37, color);
            };
            band(0);
            band(0.4);
            shape('triangle', 0.55, 0.06, 0.3, 0.17, tint);
            shape('custGeom', 0, 0, 1, 1, tint, {
                points: [
                    { x: 0.55 * W, y: 0.225 * H }, { x: 0.85 * W, y: 0.225 * H },
                    { x: 0.7 * W, y: 0.44 * H }, { close: true },
                ],
            });
            break;
        }
        case 'gem':
            shape('custGeom', 0, 0, 1, 1, color, {
                points: [
                    { x: 0.24 * W, y: 0 }, { x: 0.76 * W, y: 0 }, { x: W, y: 0.34 * H },
                    { x: 0.5 * W, y: H }, { x: 0, y: 0.34 * H }, { close: true },
                ],
            });
            shape('rect', 0.02, 0.32, 0.96, 0.025, tint); // girdle
            [[0.24, 0.34], [0.5, 0.34], [0.76, 0.34]].forEach((f) => {
                shape('rect', f[0] - 0.012, 0, 0.024, f[1], tint); // crown facets
            });
            stroke(0.02, 0.345, 0.5, 1, tint, 0.02); // pavilion edges
            stroke(0.98, 0.345, 0.5, 1, tint, 0.02);
            break;
        case 'books':
            shape('rect', 0.1, 0.04, 0.22, 0.44);
            shape('rect', 0.39, 0, 0.22, 0.48);
            shape('rect', 0.68, 0.14, 0.2, 0.34);
            shape('roundRect', 0, 0.46, 1, 0.54, color, { rectRadius: 0.1 });
            heart(0.5, 0.74, 0.42);
            break;
        case 'book':
            shape('roundRect', 0, 0, 0.86, 0.15, color, { rectRadius: 0.4 });
            shape('roundRect', 0, 0.17, 1, 0.83, color, { rectRadius: 0.08 });
            heart(0.5, 0.58, 0.44);
            break;
        case 'ringbox':
            shape('rect', 0.56, 0.02, 0.44, 0.46);
            shape('donut', 0.22, 0.14, 0.32, 0.34);
            shape('triangle', 0.3, 0.0, 0.16, 0.12);
            shape('roundRect', 0, 0.44, 0.78, 0.56, color, { rectRadius: 0.1 });
            break;
        default:
            throw new Error('unknown icon ' + kind);
    }
}

/** Hand drawn looking connector used on the process slide. */
function curvedArrow(slide, b, down) {
    const [x, y, w, h] = b;
    slide.addShape('custGeom', {
        x, y, w, h,
        line: { color: PURPLE, width: 1.5, endArrowType: 'triangle' },
        points: down
            ? [{ x: 0, y: 0 }, { curve: { type: 'quadratic', x1: w * 0.9, y1: h * 0.45 }, x: w, y: h }]
            : [{ x: 0, y: h }, { curve: { type: 'quadratic', x1: w * 0.9, y1: h * 0.55 }, x: w, y: 0 }],
    });
}

/* ------------------------------------------------------------------ slides */

/** 1 & 30 - opening / closing statement slides. */
function coverSlide(slide, title) {
    slide.addShape('rect', { x: 0.502, y: 0.471, w: 12.329, h: 6.557, fill: { color: WHITE, transparency: 20 }, line: NOLINE });
    text(slide, title, [3.384, 1.932, 6.565, 1.582], { face: SERIF, size: 88, bold: true, color: PURPLE, align: 'center' });
    text(slide, 'Wedding Proposal', [3.182, 3.445, 6.969, 1.01], { face: SERIF, size: 54, bold: true, color: INK, align: 'center' });
    para(slide, T.anAuthority, [2.765, 4.94, 7.985, 1.152], { align: 'center' });
    slide.addShape('heart', { x: 2.765, y: 3.727, w: 0.445, h: 0.445, ...fill(PINK) });
    slide.addShape('heart', { x: 10.179, y: 3.71, w: 0.445, h: 0.445, ...fill(PINK) });
}

const SLIDES = [];

SLIDES[1] = (s) => coverSlide(s, 'Romance');

SLIDES[2] = (s) => {
    s.addShape('ellipse', { x: 3.04, y: 2.875, w: 1.75, h: 1.75, ...fill(PURPLE) });
    icon(s, 'rings', [3.415, 3.243, 1.0, 1.0], PURPLE);
    heading(s, 'Wedding Is The Most Beautiful Things', [6.584, 0.587, 4.43, 1.126]);
    note(s, 'Wedding Idea', T.anAuthority, [7.273, 2.303, 1.901, 0.449], [7.319, 2.768, 5.219, 1.522]);
    note(s, T.marriage, T.gift, [6.796, 4.796, 2.147, 0.449], [6.824, 5.261, 5.096, 1.152]);
    badge(s);
};

SLIDES[3] = (s) => {
    badge(s);
    heading(s, 'Wedding + Photography Package', [3.072, 0.546, 7.189, 0.608], { align: 'center' });
    const captions = ['Picture One', 'Picture Two', 'Picture Three', 'Picture Four'];
    const shots = [0.435, 3.593, 6.778, 9.965];
    captions.forEach((cap, i) => {
        subhead(s, cap, [shots[i] + 0.5, i < 2 ? 4.767 : 4.795, 1.901, 0.449], { align: 'center' });
        para(s, T.vows, [shots[i], i < 2 ? 5.24 : 5.25, 2.901, 1.522], { align: 'center' });
    });
};

SLIDES[4] = (s) => {
    badge(s);
    heading(s, 'A hundred carts would be too few to carry all my love for you.', [0.8, 0.701, 5.24, 1.644]);
    icon(s, 'envelope', [0.865, 2.851, 0.65, 0.52], PURPLE);
    note(s, T.subtitle, T.vows, [1.859, 2.726, 1.901, 0.449], [1.891, 3.189, 4.294, 1.152]);
    icon(s, 'gift', [0.898, 4.781, 0.62, 0.62], PURPLE);
    note(s, T.subtitle, T.proclamation, [1.891, 4.722, 1.901, 0.449], [1.923, 5.171, 4.294, 1.522]);
};

SLIDES[5] = (s) => {
    badge(s);
    heading(s, 'Gravitation is not responsible for people falling in love.', [6.881, 0.488, 4.505, 1.644]);
    icon(s, 'camera', [7.292, 2.769, 0.76, 0.608], PURPLE);
    note(s, T.subtitle, T.gift, [8.443, 2.641, 1.901, 0.449], [8.475, 3.105, 3.825, 1.522]);
    icon(s, 'calendar', [6.884, 5.202, 0.76, 0.684], PURPLE);
    note(s, T.subtitle, T.gift, [8.0, 5.088, 1.901, 0.449], [8.032, 5.551, 4.575, 1.152]);
};

SLIDES[6] = (s) => {
    badge(s);
    heading(s, 'Being deeply loved by someone gives you strength.', [0.577, 4.816, 3.68, 1.644]);
    note(s, T.subtitle, T.gift, [4.978, 4.767, 1.901, 0.449], [4.996, 5.229, 3.574, 1.522]);
    note(s, T.subtitle, T.gift, [9.259, 4.779, 1.901, 0.449], [9.276, 5.238, 3.68, 1.522]);
};

SLIDES[7] = (s) => {
    badge(s);
    heading(s, 'The highest happiness on earth is the happiness of marriage.', [0.626, 0.574, 5.843, 1.126]);
    para(s, T.anAuthority, [0.67, 2.019, 5.096, 1.522]);
    note(s, T.subtitle, T.story, [0.67, 3.895, 1.901, 0.449], [0.673, 4.354, 5.527, 0.782]);
    note(s, T.subtitle, T.story, [0.67, 5.385, 1.901, 0.449], [0.673, 5.85, 6.169, 0.782]);
};

SLIDES[8] = (s) => {
    badge(s);
    heading(s, 'Every heart sings a song, incomplete, until another heart whispers back.', [2.729, 0.546, 7.876, 1.126], { align: 'center' });
    note(s, T.subtitle, T.story, [0.484, 2.049, 1.901, 0.449], [0.502, 2.498, 2.984, 1.522]);
    note(s, T.subtitle, T.story, [10.929, 2.049, 1.901, 0.449], [9.847, 2.519, 2.984, 1.522], { align: 'right' });
    pill(s, 'All You Need Is Love', [4.684, 2.077, 3.965, 0.766], [4.934, 2.205, 3.466, 0.505], { size: 24 });
};

SLIDES[9] = (s) => {
    box(s, [4.001, 0.848, 2.7, 2.7]);
    icon(s, 'house', [4.949, 1.186, 0.817, 0.735], WHITE);
    para(s, T.united, [4.247, 2.086, 2.227, 1.152], { color: WHITE, align: 'center' });
    badge(s);
    heading(s, 'You are every reason, every hope and every dream I\u2019ve ever had.', [7.26, 1.043, 4.422, 1.644]);
    pill(s, T.subtitle, [7.372, 3.075, 2.104, 0.576], [7.516, 3.124, 1.816, 0.449], { spc: 0.5, lnSpc: 1.2 });
    para(s, T.anAuthority, [7.319, 3.817, 5.096, 1.522]);
    para(s, T.gift, [7.307, 5.461, 5.096, 1.152]);
};

SLIDES[10] = (s) => {
    badge(s);
    icon(s, 'ringbox', [0.722, 0.714, 0.585, 0.585], PURPLE);
    subhead(s, 'Wedding Proposal Ideas', [1.536, 0.67, 2.464, 0.76], {});
    para(s, T.authority, [0.58, 1.51, 3.852, 1.893]);
    icon(s, 'gem', [0.639, 4.217, 0.635, 0.566], PURPLE);
    subhead(s, 'Wedding Proposal Ideas', [1.537, 4.115, 2.463, 0.76], {});
    para(s, T.authority, [0.581, 4.955, 3.852, 1.893]);
    heading(s, 'True love stories don\u2019t have endings.', [5.596, 6.558, 7.171, 0.608], { color: WHITE, align: 'center' });
};

SLIDES[11] = (s) => {
    badge(s);
    heading(s, 'Every heart sings a song, incomplete, until another heart whispers back.', [2.708, 0.546, 7.918, 1.126], { align: 'center' });
    [0.487, 4.742, 8.996].forEach((x, i) => {
        subhead(s, 'Wedding Gift', [x + 0.974, 5.241, 1.901, 0.449], { align: 'center' });
        para(s, T.story, [x, i === 1 ? 5.7 : 5.694, 3.85, 1.152], { align: 'center' });
    });
};

SLIDES[12] = (s) => {
    badge(s);
    heading(s, 'Love doesn\u2019t just sit there, like bread; remade all the time, made new.', [8.632, 1.54, 3.953, 2.163]);
    box(s, [0.62, 3.987, 3.507, 2.823]);
    subhead(s, 'Bridesmaid Here', [1.036, 4.305, 2.677, 0.449], { color: WHITE, align: 'center' });
    para(s, T.story, [0.872, 4.754, 3.012, 1.522], { color: WHITE, align: 'center' });
    box(s, [4.556, 0.658, 3.507, 2.823]);
    subhead(s, 'Bridesmaid Here', [4.971, 0.976, 2.677, 0.449], { color: WHITE, align: 'center' });
    para(s, T.story, [4.807, 1.425, 3.012, 1.522], { color: WHITE, align: 'center' });
    note(s, T.marriage, T.authority, [8.632, 4.199, 2.168, 0.449], [8.646, 4.667, 3.85, 1.893]);
};

SLIDES[13] = (s) => {
    badge(s);
    heading(s, 'I fell in love, and you smiled because you knew.', [0.848, 0.773, 5.676, 1.126]);
    icon(s, 'books', [0.981, 2.439, 0.485, 0.692], PURPLE);
    note(s, T.marriage, T.proclaim, [1.827, 2.329, 2.168, 0.449], [1.841, 2.797, 4.133, 1.522]);
    icon(s, 'book', [0.977, 4.883, 0.52, 0.65], PURPLE);
    note(s, T.marriage, T.proclaim, [1.841, 4.737, 2.168, 0.449], [1.855, 5.205, 4.133, 1.522]);
    // purple dome bleeding off the top edge
    s.addShape('ellipse', { x: 7.111, y: -0.512, w: 2.571, h: 2.571, ...fill(PURPLE) });
    box(s, [7.111, 0, 2.571, 0.774]);
    icon(s, 'gem', [7.84, 0.464, 1.142, 1.017], WHITE);
};

SLIDES[14] = (s) => {
    badge(s);
    heading(s, 'Wedding Is The Most Beautiful Things', [0.755, 4.747, 3.101, 1.644]);
    note(s, T.marriage, T.vows, [4.469, 4.804, 2.168, 0.449], [4.483, 5.272, 3.832, 1.152]);
    note(s, T.marriage, T.vows, [8.674, 4.828, 2.168, 0.449], [8.688, 5.296, 3.832, 1.152]);
};

SLIDES[15] = (s) => {
    badge(s);
    heading(s, 'There is no more lovely, friendly, relationship or company than a good marriage.', [6.557, 1.152, 5.967, 1.644]);
    icon(s, 'envelope', [6.666, 3.426, 0.65, 0.52], PURPLE);
    note(s, T.marriage, T.vows, [7.672, 3.28, 2.168, 0.449], [7.715, 3.748, 4.7, 1.152]);
    icon(s, 'gift', [6.68, 5.384, 0.65, 0.649], PURPLE);
    note(s, T.marriage, T.vows, [7.657, 5.32, 2.168, 0.449], [7.7, 5.788, 4.7, 1.152]);
};

SLIDES[16] = (s) => {
    badge(s);
    heading(s, 'Every heart sings a song, incomplete, until another heart whispers back.', [2.526, 0.546, 7.918, 1.126], { align: 'center' });
    para(s, T.authority, [1.532, 5.845, 9.907, 0.782], { align: 'center' });
};

SLIDES[17] = (s) => {
    badge(s);
    heading(s, 'Love doesn\u2019t just sit there, like bread; remade all the time, made new.', [0.675, 0.587, 7.068, 1.126]);
    const cells = [[0.675, 2.239, 0.703, 2.707], [3.92, 2.239, 3.949, 2.707], [0.643, 4.639, 0.657, 5.107], [3.888, 4.639, 3.902, 5.107]];
    cells.forEach((c) => note(s, T.marriage, T.vows, [c[0], c[1], 2.168, 0.449], [c[2], c[3], 2.997, 1.522]));
};

SLIDES[18] = (s) => {
    badge(s);
    heading(s, 'There is no remedy for love but to love more.', [6.926, 0.48, 4.417, 1.126]);
    note(s, T.marriage, T.authority, [6.932, 2.293, 2.677, 0.449], [6.957, 2.753, 5.307, 1.522]);
    box(s, [6.984, 4.846, 5.28, 2.654]);
    note(s, T.marriage, T.authority, [7.375, 5.141, 2.677, 0.449], [7.399, 5.601, 4.735, 1.522], { color: WHITE });
};

SLIDES[19] = (s) => {
    box(s, [1.373, 3.62, 3.833, 3.88]);
    badge(s);
    heading(s, 'But if you can\u2019t handle me at my worst, then you don\u2019t deserve me at my best.', [2.708, 0.546, 7.918, 1.126], { align: 'center' });
    para(s, T.authority, [1.663, 2.116, 10.008, 0.782], { align: 'center' });
    note(s, T.marriage, T.authority, [1.691, 3.995, 2.677, 0.449], [1.715, 4.469, 3.199, 2.633], { color: WHITE });
};

SLIDES[20] = (s) => {
    box(s, [6.477, 4.664, 2.909, 2.836]);
    icon(s, 'house', [7.522, 5.025, 0.817, 0.735], WHITE);
    para(s, T.united, [6.818, 5.978, 2.227, 1.152], { color: WHITE, align: 'center' });
    badge(s);
    heading(s, 'You want the rest of your life to start as soon as possible', [0.699, 0.623, 4.522, 1.644]);
    icon(s, 'camera', [0.816, 3.087, 0.729, 0.583], PURPLE);
    note(s, T.marriage, T.vows, [1.912, 3.01, 2.168, 0.449], [1.94, 3.478, 3.99, 1.152]);
    icon(s, 'calendar', [0.816, 5.122, 0.729, 0.656], PURPLE);
    note(s, T.marriage, T.vows, [1.912, 5.025, 2.168, 0.449], [1.94, 5.493, 3.99, 1.152]);
};

SLIDES[21] = (s) => {
    badge(s);
    box(s, [4.873, 3.912, 3.588, 3.588]);
    icon(s, 'house', [6.039, 4.266, 1.256, 1.13], WHITE);
    para(s, T.vows, [5.15, 5.577, 3.033, 1.522], { color: WHITE, align: 'center' });
    heading(s, 'Tminister to each other in silent unspeakable memories at the moment of the last parting?', [2.708, 0.546, 7.918, 1.644], { align: 'center' });
    para(s, T.authority, [1.663, 2.512, 10.008, 0.782], { align: 'center' });
};

SLIDES[22] = (s) => {
    box(s, [0, 0, 2.81, 7.5]);
    imagePlaceholder(s, [-2.819, 1.119, 9.149, 5.604]);
    badge(s);
    heading(s, 'We loved with a love that was more than love.', [6.046, 0.71, 5.146, 1.126]);
    quote(s, '\u201C' + T.authority + '\u201D', '- Tiara Spandex', [6.046, 2.245, 5.754, 2.36]);
    quote(s, '\u201C' + T.vows + '\u201D', '- Sarah Connor', [6.667, 4.984, 5.754, 1.553]);
};
function quote(s, saying, who, b) {
    s.addText([
        { text: saying, options: { fontSize: 12, italic: true, fontFace: SANS, color: GREY, charSpacing: 0.3, lineSpacingMultiple: 2, breakLine: true } },
        { text: who, options: { fontSize: 18, fontFace: SERIF, color: PURPLE, charSpacing: 0.3, lineSpacingMultiple: 2.5 } },
    ], { x: b[0], y: b[1], w: b[2], h: b[3], valign: 'top', wrap: true });
}

SLIDES[23] = (s) => {
    imagePlaceholder(s, [7.082, 1.997, 2.49, 4.992]);
    imagePlaceholder(s, [10.31, 1.997, 2.49, 4.992]);
    badge(s);
    heading(s, 'There is no remedy for love but to love more.', [0.755, 0.48, 5.145, 1.126]);
    note(s, 'iOS and Android', T.proclamation, [0.761, 2.051, 2.677, 0.449], [0.786, 2.51, 5.471, 1.152]);
    note(s, 'Our Website', T.vows, [0.76, 4.044, 2.677, 0.449], [0.784, 4.504, 5.471, 0.782]);
    s.addShape('roundRect', { x: 0.811, y: 5.936, w: 3.767, h: 0.596, rectRadius: 0.1, ...fill(PURPLE) });
    text(s, 'www.wedding-organizer.com', [0.962, 6.021, 3.466, 0.404], { face: SERIF, color: WHITE, align: 'center' });
};

SLIDES[24] = (s) => {
    badge(s);
    heading(s, T.infographics, [2.708, 0.546, 7.918, 0.608], { align: 'center' });
    const cols = [
        { kind: 'camera', ib: [1.798, 2.162, 1.55, 1.24], head: [1.235, 3.701], body: [0.854, 4.15] },
        { kind: 'gift', ib: [5.969, 2.012, 1.395, 1.394], head: [5.328, 3.701], body: [4.93, 4.161] },
        { kind: 'house', ib: [9.967, 2.007, 1.55, 1.395], head: [9.404, 3.691], body: [9.006, 4.161] },
    ];
    cols.forEach((c) => {
        icon(s, c.kind, c.ib, PURPLE);
        subhead(s, 'Infographic Here', [c.head[0], c.head[1], 2.677, 0.449], { align: 'center' });
        para(s, T.authority, [c.body[0], c.body[1], 3.473, 2.263], { align: 'center' });
    });
};

SLIDES[25] = (s) => {
    badge(s);
    heading(s, T.infographics, [2.708, 0.546, 7.918, 0.608], { align: 'center' });
    const items = [
        { n: '1', oval: [1.488, 1.884], kind: 'camera', ib: [1.157, 2.901, 1.335, 1.068], head: [5.581, 2.088], body: [5.581, 2.551] },
        { n: '2', oval: [3.73, 1.901], kind: 'envelope', ib: [3.4, 2.904, 1.335, 1.068], head: [9.241, 2.032], body: [9.241, 2.496] },
        { n: '3', oval: [1.488, 4.394], kind: 'gift', ib: [1.172, 5.24, 1.305, 1.305], head: [5.581, 4.567], body: [5.581, 5.031] },
        { n: '4', oval: [3.73, 4.373], kind: 'house', ib: [3.395, 5.333, 1.335, 1.201], head: [9.241, 4.512], body: [9.241, 4.976] },
    ];
    items.forEach((it) => {
        icon(s, it.kind, it.ib, PURPLE);
        s.addShape('ellipse', { x: it.oval[0], y: it.oval[1], w: 0.674, h: 0.674, ...fill(PURPLE) });
        text(s, it.n, [it.oval[0], it.oval[1] + 0.036, 0.674, 0.565], { face: SERIF, size: 24, bold: true, color: WHITE, spc: 0.5, align: 'center', lnSpc: 1.2 });
        subhead(s, it.n + '. ' + T.subtitle, [it.head[0], it.head[1], 2.677, 0.449]);
        para(s, T.vows, [it.body[0], it.body[1], 3.216, 1.522]);
    });
};

SLIDES[26] = (s) => {
    badge(s);
    heading(s, T.infographics, [2.708, 0.546, 7.918, 0.608], { align: 'center' });
    const rows = [
        { kind: 'camera', ib: [0.789, 2.203, 1.335, 1.068], head: [2.586, 2.088], body: [2.586, 2.551] },
        { kind: 'envelope', ib: [7.149, 2.196, 1.335, 1.068], head: [8.9, 2.073], body: [8.9, 2.537] },
        { kind: 'gift', ib: [0.822, 4.663, 1.302, 1.302], head: [2.586, 4.623], body: [2.586, 5.087] },
        { kind: 'house', ib: [7.149, 4.763, 1.335, 1.201], head: [8.9, 4.609], body: [8.9, 5.073] },
    ];
    rows.forEach((r) => {
        icon(s, r.kind, r.ib, PURPLE);
        note(s, T.subtitle, T.brief, [r.head[0], r.head[1], 2.677, 0.449], [r.body[0], r.body[1], 3.828, 1.522]);
    });
};

SLIDES[27] = (s) => {
    badge(s);
    heading(s, T.infographics, [2.708, 0.546, 7.918, 0.608], { align: 'center' });
    const steps = [
        { kind: 'camera', ib: [1.612, 1.865, 0.921, 0.737], body: [0.555, 2.799] },
        { kind: 'envelope', ib: [3.943, 4.189, 0.921, 0.737], body: [2.887, 5.055] },
        { kind: 'gift', ib: [6.24, 1.724, 0.899, 0.899], body: [5.183, 2.798] },
        { kind: 'house', ib: [8.606, 4.176, 0.921, 0.829], body: [7.549, 5.18] },
        { kind: 'camera', ib: [10.801, 1.864, 0.921, 0.737], body: [9.744, 2.798] },
    ];
    steps.forEach((st) => {
        icon(s, st.kind, st.ib, PURPLE);
        para(s, T.vows, [st.body[0], st.body[1], 3.035, 1.522], { align: 'center' });
    });
    curvedArrow(s, [1.734, 4.778, 0.665, 1.163], true);
    curvedArrow(s, [4.384, 2.201, 0.663, 1.164], false);
    curvedArrow(s, [6.59, 4.927, 0.665, 1.163], true);
    curvedArrow(s, [8.908, 2.332, 0.663, 1.164], false);
};

SLIDES[28] = (s) => {
    badge(s);
    heading(s, T.infographics, [2.708, 0.546, 7.918, 0.608], { align: 'center' });
    const xs = [0.567, 4.739, 8.885];
    const rows = [1.889, 4.514];
    let n = 0;
    rows.forEach((y, r) => {
        xs.forEach((x) => {
            n += 1;
            const bx = r === 1 ? x + 0.013 : x; // bottom row sits a hair further right
            s.addShape('ellipse', { x: bx + 0.057, y: y + 0.054, w: 0.734, h: 0.734, ...fill(PURPLE) });
            text(s, String(n), [bx + 0.087, y + 0.124, 0.674, 0.565], { face: SERIF, size: 24, bold: true, color: WHITE, spc: 0.5, align: 'center', lnSpc: 1.2 });
            s.addText([
                { text: 'Infographic', options: { breakLine: true } },
                { text: 'Subtitle' },
            ], {
                x: bx + 1.005, y, w: 1.785, h: 0.813, fontFace: SERIF, fontSize: 18, bold: true,
                color: PURPLE, charSpacing: 0.5, lineSpacingMultiple: 1.2, valign: 'top',
            });
            para(s, T.vows, [bx, y + 0.935, 3.856, 1.152]);
        });
    });
};

SLIDES[29] = (s) => {
    // two cards: a purple banner on top of a purple text block
    [[0.755, 0.737], [4.115, 0.737]].forEach((p) => box(s, [p[0], p[1], 3.023, 2.003]));
    box(s, [0.755, 3.143, 3.023, 3.568]);
    box(s, [4.115, 3.143, 3.023, 3.568]);
    badge(s);
    icon(s, 'calendar', [1.709, 1.219, 1.148, 1.034], WHITE);
    icon(s, 'envelope', [5.053, 1.277, 1.148, 0.919], WHITE);
    subhead(s, 'Infographic Here', [0.996, 3.466, 2.53, 0.449], { color: WHITE, align: 'center' });
    para(s, T.coupleGift, [1.07, 3.915, 2.383, 2.263], { color: WHITE, align: 'center' });
    subhead(s, 'Infographic Here', [4.388, 3.511, 2.53, 0.449], { color: WHITE, align: 'center' });
    para(s, T.coupleGift, [4.434, 3.961, 2.383, 2.263], { color: WHITE, align: 'center' });
    heading(s, T.infographics, [7.707, 1.225, 3.525, 1.126]);
    para(s, T.authority, [7.754, 2.75, 4.718, 1.522]);
    para(s, T.authority, [7.754, 4.511, 4.718, 1.522]);
};

SLIDES[30] = (s) => coverSlide(s, 'Thank You');

/* --------------------------------------------------------------- map atlas
 * Slides 31/32 are sheets of flat grey outline maps.  Each entry below is a
 * silhouette given as normalised [x, y] vertices inside its own bounding box;
 * `landmass` scales one into place.  Some maps are drawn from several pieces.
 */

const MAPS = {
    africa: [
        [[0.30, 0.02], [0.42, 0.00], [0.62, 0.02], [0.74, 0.06], [0.86, 0.05], [0.92, 0.10], [0.90, 0.17],
         [0.80, 0.24], [0.78, 0.33], [0.72, 0.40], [0.70, 0.50], [0.64, 0.58], [0.62, 0.68], [0.56, 0.80],
         [0.50, 0.90], [0.44, 1.00], [0.40, 0.90], [0.40, 0.78], [0.34, 0.66], [0.28, 0.58], [0.22, 0.50],
         [0.20, 0.40], [0.14, 0.34], [0.06, 0.28], [0.02, 0.20], [0.06, 0.12], [0.16, 0.06]],
    ],
    world: [
        [[0.04, 0.06], [0.14, 0.02], [0.20, 0.08], [0.18, 0.16], [0.10, 0.20], [0.14, 0.28], [0.12, 0.36],
         [0.16, 0.40], [0.14, 0.50], [0.12, 0.62], [0.08, 0.78], [0.05, 0.96], [0.03, 0.80], [0.05, 0.62],
         [0.06, 0.46], [0.02, 0.34], [0.00, 0.20], [0.02, 0.12]],
        [[0.30, 0.02], [0.38, 0.00], [0.44, 0.06], [0.42, 0.14], [0.46, 0.20], [0.44, 0.30], [0.40, 0.42],
         [0.36, 0.56], [0.32, 0.72], [0.30, 0.60], [0.32, 0.46], [0.30, 0.34], [0.26, 0.22], [0.24, 0.10]],
        [[0.50, 0.06], [0.60, 0.02], [0.72, 0.02], [0.84, 0.06], [0.94, 0.10], [0.98, 0.18], [0.92, 0.26],
         [0.86, 0.32], [0.80, 0.30], [0.74, 0.36], [0.68, 0.34], [0.62, 0.26], [0.56, 0.22], [0.48, 0.16]],
        [[0.46, 0.24], [0.52, 0.22], [0.56, 0.30], [0.54, 0.44], [0.50, 0.56], [0.46, 0.44], [0.44, 0.34]],
        [[0.80, 0.56], [0.90, 0.54], [0.96, 0.62], [0.92, 0.72], [0.84, 0.76], [0.78, 0.70], [0.76, 0.62]],
    ],
    asia: [
        [[0.06, 0.24], [0.14, 0.14], [0.28, 0.08], [0.44, 0.04], [0.58, 0.02], [0.72, 0.04], [0.84, 0.10],
         [0.94, 0.16], [1.00, 0.26], [0.94, 0.34], [0.96, 0.44], [0.88, 0.50], [0.82, 0.46], [0.76, 0.54],
         [0.78, 0.64], [0.70, 0.68], [0.64, 0.62], [0.58, 0.70], [0.54, 0.86], [0.48, 0.72], [0.44, 0.60],
         [0.36, 0.54], [0.28, 0.58], [0.22, 0.50], [0.14, 0.44], [0.06, 0.38], [0.02, 0.30]],
    ],
    southAmerica: [
        [[0.36, 0.02], [0.50, 0.00], [0.64, 0.04], [0.76, 0.06], [0.86, 0.12], [0.84, 0.20], [0.76, 0.28],
         [0.72, 0.38], [0.66, 0.50], [0.62, 0.62], [0.58, 0.74], [0.52, 0.86], [0.46, 1.00], [0.42, 0.86],
         [0.42, 0.72], [0.36, 0.58], [0.30, 0.44], [0.24, 0.32], [0.16, 0.22], [0.12, 0.12], [0.20, 0.06]],
    ],
    europe: [
        [[0.10, 0.36], [0.18, 0.28], [0.14, 0.20], [0.22, 0.14], [0.30, 0.20], [0.34, 0.10], [0.44, 0.06],
         [0.52, 0.00], [0.70, 0.00], [0.88, 0.02], [1.00, 0.06], [0.98, 0.30], [0.92, 0.44], [0.82, 0.48],
         [0.74, 0.44], [0.68, 0.54], [0.60, 0.62], [0.54, 0.74], [0.46, 0.86], [0.38, 0.96], [0.32, 0.86],
         [0.34, 0.72], [0.26, 0.66], [0.18, 0.60], [0.12, 0.50], [0.06, 0.42]],
    ],
    australia: [
        [[0.20, 0.24], [0.28, 0.14], [0.38, 0.16], [0.46, 0.06], [0.54, 0.14], [0.62, 0.10], [0.70, 0.18],
         [0.80, 0.16], [0.90, 0.24], [0.96, 0.36], [0.94, 0.50], [0.88, 0.62], [0.78, 0.72], [0.66, 0.80],
         [0.54, 0.82], [0.42, 0.78], [0.30, 0.70], [0.20, 0.58], [0.14, 0.44], [0.14, 0.32]],
        [[0.72, 0.88], [0.84, 0.86], [0.86, 0.96], [0.76, 0.98]],
    ],
    usa: [
        [[0.20, 0.14], [0.36, 0.10], [0.52, 0.08], [0.70, 0.08], [0.86, 0.12], [0.94, 0.20], [0.90, 0.30],
         [0.94, 0.40], [0.86, 0.50], [0.80, 0.62], [0.74, 0.72], [0.68, 0.66], [0.60, 0.72], [0.52, 0.66],
         [0.44, 0.62], [0.34, 0.58], [0.26, 0.50], [0.20, 0.40], [0.16, 0.28], [0.16, 0.20]],
        [[0.00, 0.62], [0.10, 0.56], [0.14, 0.64], [0.10, 0.74], [0.02, 0.74]],
        [[0.16, 0.86], [0.24, 0.84], [0.26, 0.92], [0.18, 0.94]],
    ],
    germany: [
        [[0.34, 0.02], [0.46, 0.00], [0.56, 0.06], [0.60, 0.14], [0.72, 0.10], [0.78, 0.18], [0.72, 0.28],
         [0.78, 0.36], [0.88, 0.42], [0.84, 0.54], [0.88, 0.64], [0.80, 0.72], [0.74, 0.84], [0.62, 0.94],
         [0.52, 1.00], [0.44, 0.92], [0.38, 0.82], [0.28, 0.76], [0.22, 0.66], [0.26, 0.56], [0.18, 0.48],
         [0.22, 0.36], [0.16, 0.26], [0.24, 0.14]],
    ],
    korea: [
        [[0.44, 0.00], [0.58, 0.04], [0.62, 0.14], [0.56, 0.22], [0.68, 0.28], [0.76, 0.36], [0.72, 0.46],
         [0.80, 0.54], [0.76, 0.66], [0.80, 0.78], [0.70, 0.88], [0.58, 0.96], [0.48, 0.90], [0.44, 0.78],
         [0.36, 0.68], [0.28, 0.58], [0.34, 0.48], [0.26, 0.38], [0.32, 0.28], [0.28, 0.16], [0.34, 0.06]],
        [[0.06, 0.46], [0.18, 0.44], [0.20, 0.52], [0.08, 0.54]],
    ],
    china: [
        [[0.12, 0.30], [0.20, 0.18], [0.34, 0.10], [0.48, 0.04], [0.60, 0.06], [0.70, 0.02], [0.82, 0.08],
         [0.94, 0.14], [1.00, 0.24], [0.94, 0.32], [0.96, 0.42], [0.88, 0.50], [0.82, 0.60], [0.74, 0.68],
         [0.68, 0.80], [0.60, 0.90], [0.54, 0.80], [0.46, 0.74], [0.36, 0.76], [0.28, 0.66], [0.20, 0.56],
         [0.10, 0.50], [0.04, 0.40]],
    ],
    russia: [
        [[0.04, 0.36], [0.10, 0.24], [0.20, 0.16], [0.32, 0.10], [0.44, 0.06], [0.56, 0.02], [0.68, 0.04],
         [0.80, 0.08], [0.90, 0.14], [0.98, 0.22], [0.94, 0.34], [0.98, 0.46], [0.92, 0.58], [0.84, 0.68],
         [0.74, 0.76], [0.62, 0.82], [0.52, 0.76], [0.44, 0.84], [0.34, 0.90], [0.26, 0.80], [0.18, 0.70],
         [0.10, 0.58], [0.04, 0.48], [0.00, 0.42]],
    ],
    france: [
        [[0.34, 0.02], [0.46, 0.00], [0.54, 0.08], [0.66, 0.10], [0.74, 0.18], [0.86, 0.26], [0.88, 0.40],
         [0.80, 0.52], [0.76, 0.66], [0.66, 0.78], [0.54, 0.86], [0.42, 0.82], [0.32, 0.72], [0.22, 0.58],
         [0.16, 0.44], [0.20, 0.30], [0.26, 0.16]],
        [[0.74, 0.88], [0.86, 0.86], [0.88, 0.96], [0.76, 0.98]],
    ],
    india: [
        [[0.24, 0.10], [0.36, 0.02], [0.48, 0.00], [0.60, 0.04], [0.72, 0.02], [0.82, 0.10], [0.76, 0.20],
         [0.80, 0.28], [0.72, 0.36], [0.68, 0.48], [0.62, 0.60], [0.56, 0.74], [0.50, 0.88], [0.44, 0.74],
         [0.40, 0.58], [0.34, 0.44], [0.26, 0.34], [0.18, 0.26], [0.14, 0.16]],
        [[0.48, 0.92], [0.58, 0.92], [0.56, 0.98], [0.48, 0.98]],
        [[0.40, 0.00], [0.52, 0.00], [0.50, 0.04], [0.40, 0.04]],
    ],
    canada: [
        [[0.08, 0.32], [0.16, 0.22], [0.28, 0.14], [0.40, 0.08], [0.52, 0.10], [0.62, 0.04], [0.74, 0.08],
         [0.86, 0.06], [0.96, 0.14], [0.94, 0.26], [0.98, 0.38], [0.92, 0.50], [0.84, 0.62], [0.74, 0.72],
         [0.62, 0.80], [0.52, 0.74], [0.42, 0.80], [0.32, 0.76], [0.24, 0.66], [0.16, 0.54], [0.10, 0.44],
         [0.04, 0.38]],
    ],
};

/**
 * Draws one silhouette from MAPS, scaled into the given box, then rules a few
 * white hairlines across it for the internal borders of the original artwork.
 * The lines run the full width of the box; the parts outside the land are
 * white on white and therefore invisible.
 */
function landmass(slide, name, b) {
    const [x, y, w, h] = b;
    MAPS[name].forEach((outline) => {
        slide.addShape('custGeom', {
            x, y, w, h,
            fill: { color: MAPGREY },
            line: { color: WHITE, width: 0.5 },
            points: outline.map((p) => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }]),
        });
    });
    const n = 3;
    for (let i = 1; i <= n; i += 1) {
        slide.addShape('line', { x: x + (w * i) / (n + 1) - 0.06 * w, y, w: 0.06 * w, h, flipH: true, line: { color: WHITE, width: 0.75 } });
        slide.addShape('line', { x, y: y + (h * i) / (n + 1), w, h: 0.04 * h, line: { color: WHITE, width: 0.75 } });
    }
}

SLIDES[31] = (s) => {
    landmass(s, 'africa', [1.458, 1.049, 2.216, 2.492]);
    landmass(s, 'world', [4.564, 1.193, 3.502, 2.205]);
    landmass(s, 'asia', [9.181, 1.224, 3.028, 2.143]);
    landmass(s, 'southAmerica', [1.568, 3.891, 1.997, 2.789]);
    landmass(s, 'europe', [4.936, 4.222, 2.758, 2.127]);
    landmass(s, 'australia', [9.338, 4.098, 2.712, 2.374]);
};

SLIDES[32] = (s) => {
    landmass(s, 'usa', [0.964, 1.394, 3.088, 2.057]);
    landmass(s, 'germany', [4.452, 1.116, 1.914, 2.614]);
    landmass(s, 'korea', [6.864, 0.847, 2.172, 3.152]);
    landmass(s, 'china', [9.521, 1.283, 2.972, 2.28]);
    landmass(s, 'russia', [0.928, 4.501, 3.16, 1.658]);
    landmass(s, 'france', [4.513, 4.304, 1.793, 2.051]);
    landmass(s, 'india', [6.722, 3.979, 2.455, 2.7]);
    landmass(s, 'canada', [9.438, 4.098, 3.139, 2.462]);
};

/* -------------------------------------------------------------------- main */

function build() {
    const pres = new PptxGenJS();
    pres.defineLayout({ name: 'WIDE16x9', width: 13.3333333, height: 7.5 }); // 12192000 x 6858000 EMU
    pres.layout = 'WIDE16x9';
    pres.author = 'pptxgenjs';
    pres.title = 'Romance - Wedding Proposal';

    for (let i = 1; i <= 32; i += 1) {
        const slide = pres.addSlide();
        slide.background = { color: WHITE };
        SLIDES[i](slide);
    }
    return pres;
}

build()
    .writeFile({ fileName: path.join(__dirname, '0f61b19d-5c80-4151-8725-b55b8b5ef284_grok_final.pptx') })
    .then((f) => console.log('wrote ' + f))
    .catch((e) => { console.error(e); process.exit(1); });
