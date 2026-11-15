/**
 * "Givin Company" brand guidelines deck — 20 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt with pptxgenjs only. Raster photos in the original are drawn here as
 * grey placeholder rectangles (see `photo()`).
 *
 *   node 07eef4ff-d17b-47ab-bf8d-8f926c752e5f_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const DARK = '11141B';
const WHITE = 'FEFEFE';
const ORANGE = 'E76E10';
const YELLOW = 'F6E300';
const GRID = 'DBDBDB';
const PHOTO_FILL = 'E0E0E0';
const PHOTO_EDGE = 'CCCCCC';

const DISPLAY = 'Darker Grotesque Medium'; // headlines / numerals
const BODY = 'Inter Tight'; // paragraphs / nav

const SLIDE_W = 13.3333333; // 12192000 EMU
const SLIDE_H = 7.5;

/* ------------------------------------------------- reusable text presets */

// 110 pt display words that make up the giant two-line slide titles.
const displayText = { fontFace: DISPLAY, fontSize: 110, color: DARK, valign: 'top' };
// 30 pt section sub-heads ("Our Vision", "Point One", ...).
const headingText = { fontFace: DISPLAY, fontSize: 30, color: DARK, valign: 'top' };
// 11 pt justified body copy, 150 % leading.
const bodyText = {
	fontFace: BODY, fontSize: 11, color: DARK,
	align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
};
// 11 pt single-line labels (nav items, captions under names).
const labelText = { fontFace: BODY, fontSize: 11, color: DARK, valign: 'top' };

/* --------------------------------------------------------------- helpers */

function text(slide, content, x, y, w, h, opts) {
	slide.addText(content, Object.assign({ x, y, w, h }, opts));
}

/**
 * Add a flat-coloured shape.
 *
 * A hairline in the fill colour is required: pptxgenjs emits an empty `<a:ln>`
 * when no line is given, and renderers then fall back to the theme's default
 * outline, which halos every shape.
 */
function shape(slide, kind, opts) {
	const color = opts.fill.color;
	return slide.addShape(kind, Object.assign({ line: { color, width: 0.01 } }, opts));
}

/** Display headline word. */
function display(slide, words, x, y, w, h, size) {
	slide.addText(words, Object.assign({}, displayText, { x, y, w, h }, size ? { fontSize: size } : null));
}

/** 30 pt sub-head. */
function heading(slide, words, x, y, w) {
	slide.addText(words, Object.assign({}, headingText, { x, y, w, h: 0.606 }));
}

/** Justified body paragraph. `content` may be a string or a rich-run array. */
function body(slide, content, x, y, w, h) {
	slide.addText(content, Object.assign({}, bodyText, { x, y, w, h }));
}

/** Mixed-size run used three times in the deck (12 pt / 11 pt / 12 pt). */
function mixedRuns(parts) {
	return parts.map(([t, size]) => ({
		text: t,
		options: { fontFace: BODY, fontSize: size, color: DARK },
	}));
}

/* ------------------------------------------------- brand geometry motifs */

/** Square with a square notch bitten out of its lower-left corner. */
function notchSquare(slide, x, y, size, color, rotate) {
	const n = 0.576 * size;
	shape(slide, 'custGeom', {
		x, y, w: size, h: size, fill: { color }, rotate: rotate || 0,
		points: [
			{ x: size, y: 0 }, { x: size, y: size }, { x: 0, y: size },
			{ x: 0, y: n }, { x: n, y: n }, { x: n, y: 0 }, { close: true },
		],
	});
}

/** Plain square (used as the small dark / orange accent tiles). */
function square(slide, x, y, size, color, rotate) {
	shape(slide, 'rect', { x, y, w: size, h: size, fill: { color }, rotate: rotate || 0 });
}

// Outline of the six-pointed asterisk, normalised to a 1 x 1 box.
const ASTERISK = [
	[1, 0.3308], [0.9027, 0.1694], [0.5971, 0.3385], [0.5971, 0], [0.4027, 0],
	[0.4027, 0.3385], [0.0974, 0.1694], [0, 0.3308], [0.3056, 0.5001], [0, 0.6692],
	[0.0974, 0.8309], [0.4027, 0.6615], [0.4027, 1], [0.5971, 1], [0.5971, 0.6615],
	[0.9027, 0.8309], [1, 0.6692], [0.6944, 0.5001],
];

function asterisk(slide, x, y, w, h, color) {
	shape(slide, 'custGeom', {
		x, y, w, h, fill: { color: color || ORANGE },
		points: ASTERISK.map(([px, py]) => ({ x: px * w, y: py * h })).concat([{ close: true }]),
	});
}

/** The "* *" pair motif: two asterisks side by side inside a w x h box. */
function asteriskPair(slide, x, y, w, h) {
	asterisk(slide, x, y, 0.44 * w, h);
	asterisk(slide, x + 0.56 * w, y, 0.44 * w, h);
}

// Corner arrow = vertical bar + horizontal bar + diagonal shaft, in a 1 x 1 box.
const ARROW_PARTS = [
	{ dx: 0, dy: 0.2570, dw: 0.1516, dh: 0.7266 },
	{ dx: 0.2512, dy: 0, dw: 0.7344, dh: 0.1516 },
	{ dx: 0.1516, dy: 0.1516, dw: 0.8484, dh: 0.8488, diagonal: true },
];
const ARROW_DIAGONAL = [[1, 0.8785], [0.8781, 1], [0, 0.1242], [0.1186, 0]];

/** Rotate a point around a centre by `deg` (only right angles are used here). */
function rotateAbout(px, py, cx, cy, deg) {
	const r = (deg * Math.PI) / 180;
	const c = Math.cos(r), s = Math.sin(r);
	return [cx + (px - cx) * c - (py - cy) * s, cy + (px - cx) * s + (py - cy) * c];
}

/** Angular "↖" arrow motif; `rotate` picks the corner it points to. */
function cornerArrow(slide, x, y, w, h, rotate) {
	const rot = rotate || 0;
	const cx = x + w / 2, cy = y + h / 2;
	ARROW_PARTS.forEach(p => {
		const pw = p.dw * w, ph = p.dh * h;
		const [rcx, rcy] = rotateAbout(x + p.dx * w + pw / 2, y + p.dy * h + ph / 2, cx, cy, rot);
		const box = { x: rcx - pw / 2, y: rcy - ph / 2, w: pw, h: ph, fill: { color: ORANGE }, rotate: rot };
		if (p.diagonal) {
			shape(slide, 'custGeom', Object.assign(box, {
				points: ARROW_DIAGONAL.map(([px, py]) => ({ x: px * pw, y: py * ph })).concat([{ close: true }]),
			}));
		} else {
			shape(slide, 'rect', box);
		}
	});
}

/** The four-tile "Givin" logo mark (orange square, yellow notch, two dark chips). */
function logoMark(slide, x, y, size) {
	notchSquare(slide, x + 0.334 * size, y + 0.334 * size, 0.667 * size, YELLOW);
	square(slide, x + 0.749 * size, y, 0.251 * size, DARK);
	square(slide, x, y + 0.734 * size, 0.251 * size, DARK);
	square(slide, x + 0.051 * size, y + 0.051 * size, 0.566 * size, ORANGE);
}

/** Rounded "pill" button with centred label. */
function pillButton(slide, label, x, y, w, h, labelX, labelW) {
	slide.addShape('roundRect', {
		x, y, w, h, fill: { type: 'none' },
		line: { color: DARK, width: 1 }, rectRadius: h / 2,
	});
	text(slide, label, labelX, y + 0.082, labelW, 0.286, Object.assign({}, labelText, { align: 'center' }));
}

/**
 * Grey stand-in for a photo placeholder from the original deck.
 *
 * The original fills each frame with an "insert your picture here" plate —
 * 2048 x 1152 px landscape or 1152 x 2048 px portrait — scaled to cover the
 * frame, with a 512 px-wide caption centred on it. Reproducing that cover
 * scale keeps the caption the same size as in the reference.
 */
const PLATE_CAPTION_PX = 512;
const CAPTION_PT_PER_INCH = 10.71; // bold DM Sans: block width -> point size

function photo(slide, x, y, w, h) {
	shape(slide, 'rect', { x, y, w, h, fill: { color: PHOTO_FILL } });
	slide.addShape('line', { x, y, w, h, line: { color: PHOTO_EDGE, width: 0.5 } });
	slide.addShape('line', { x, y: y + h, w, h: -h, line: { color: PHOTO_EDGE, width: 0.5 } });

	const [plateW, plateH] = w >= h ? [2048, 1152] : [1152, 2048];
	const captionW = PLATE_CAPTION_PX * Math.max(w / plateW, h / plateH);
	const fontSize = Math.round(captionW * CAPTION_PT_PER_INCH * 2) / 2;
	slide.addText('INSERT YOUR\nPICTURE HERE', {
		x, y: y + h / 2 - 0.5, w, h: 1.0,
		fontFace: 'DM Sans', fontSize, bold: true, color: DARK,
		align: 'center', valign: 'middle', lineSpacing: fontSize * 1.1,
	});
}

/* ------------------------------------------------ page frame (all slides) */

/** Faint 11 x 11 measuring grid that sits behind every slide. */
function backgroundGrid(slide) {
	const x0 = 0.002, y0 = 0.002, colW = 1.2110, rowH = 0.6812;
	for (let i = 0; i <= 11; i++) {
		slide.addShape('line', { x: x0 + i * colW, y: y0, w: 0, h: 11 * rowH, line: { color: GRID, width: 0.25 } });
		slide.addShape('line', { x: x0, y: y0 + i * rowH, w: 11 * colW, h: 0, line: { color: GRID, width: 0.25 } });
	}
}

const NAV_ITEMS = [
	['Home', 4.414, 0.714], ['About', 5.704, 0.714],
	['Branding', 6.993, 0.8], ['Contact', 8.369, 0.8],
];

/** Header rules, wordmark, nav links and hamburger — identical on all slides. */
function pageChrome(slide) {
	slide.addShape('line', { x: 0, y: 1.194, w: 13.333, h: 0, line: { color: DARK, width: 0.75 } });
	slide.addShape('line', { x: 2.75, y: 0, w: 0, h: 1.194, line: { color: DARK, width: 0.75 } });
	slide.addShape('line', { x: 10.833, y: 0, w: 0, h: 1.194, line: { color: DARK, width: 0.75 } });

	text(slide, 'Givin Company', 0.989, 0.604, 1.594, 0.353,
		{ fontFace: DISPLAY, fontSize: 15, bold: true, color: DARK, valign: 'top' });
	logoMark(slide, 0.667, 0.667, 0.247);

	NAV_ITEMS.forEach(([label, x, w]) => {
		text(slide, label, x, 0.604, w, 0.286, Object.assign({}, labelText, { align: 'center' }));
	});

	[[12.454, 0.687, 0.2], [12.253, 0.798, 0.401], [12.387, 0.909, 0.267]].forEach(([x, y, w]) => {
		slide.addShape('line', { x, y, w, h: 0, line: { color: ORANGE, width: 2 } });
	});
}

/* ----------------------------------------------------- shared copy blocks */

const COPY = {
	mountains: 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds.',
	mountainsLong: "High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. A hiker's path weaves through the alpine meadows, where wild creatures roam freely and hardy flora cling",
	critter: 'Every rustle of a critter in the underbrush, every gurgle of a hidden stream adds another verse to the enchanting sonnet.',
	critterSanctuary: 'Every rustle of a critter in the underbrush, every gurgle of a hidden stream adds another verse to the enchanting sonnet of this untamed sanctuary.',
	critterShort: 'Every rustle of a critter in the underbrush, every gurgle of a hidden stream adds another verse.',
	critterEnchanting: 'Every rustle of a critter in the underbrush, every gurgle of a hidden stream adds another verse to the enchanting.',
	wilderness: "In the heart of the verdant wilderness, where the symphony of leaves rustling and birdsong intertwine, nature's masterpiece unfolds in breathtaking hues. The sun's warm embrace cascades through the canopy.",
	wildernessLong: "In the heart of the verdant wilderness, where the symphony of leaves rustling and birdsong intertwine, nature's masterpiece unfolds in breathtaking hues. The sun's warm embrace cascades through the canopy, dappling the forest floor with patches of golden light.",
};

/* --------------------------------------------------------- slide builders */

// 1 — Cover: "Gi- Vin" split across three photo plates.
function slide01(s) {
	pillButton(s, 'Architecture Company', 1.138, 6.36, 2.417, 0.468, 1.322, 2.048);
	pillButton(s, 'Brand Guidelines', 9.779, 6.36, 2.417, 0.468, 10.187, 1.6);
	asteriskPair(s, 1.142, 5.602, 0.822, 0.377);

	// Thin rule with a small triangular pointer, under the centre plate.
	shape(s, 'custGeom', {
		x: 5.817, y: 6.698, w: 1.697, h: 0.128, fill: { color: ORANGE },
		points: [
			{ x: 0.9715, y: 0.1177 }, { x: 0.8485, y: 0 }, { x: 0.7255, y: 0.1177 },
			{ x: 0, y: 0.1177 }, { x: 0, y: 0.128 }, { x: 1.697, y: 0.128 },
			{ x: 1.697, y: 0.1177 }, { close: true },
		],
	});
	cornerArrow(s, 11.664, 1.946, 0.531, 0.532, 270);

	photo(s, 10.532, 4.126, 1.663, 1.635);
	photo(s, 1.137, 2.333, 1.917, 1.884);
	photo(s, 4.66, 1.947, 3.979, 4.033);

	display(s, 'Gi-', 2.318, 2.23, 3.479, 3.467, 200);
	display(s, 'Vin', 8.101, 2.23, 3.479, 3.467, 200);
}

// 2 — Introduction.
function slide02(s) {
	display(s, 'Intro-', 0.5, 4.255, 4.083, 1.952);
	display(s, 'duction', 0.5, 5.33, 4.873, 1.952);
	asteriskPair(s, 0.729, 2.792, 1.196, 0.549);
	s.addText(COPY.wildernessLong, Object.assign({}, bodyText, { x: 6.095, y: 2.242, w: 5.14, h: 1.277, fontSize: 12 }));
	cornerArrow(s, 12.136, 4.167, 0.531, 0.532, 0);
	photo(s, 6.095, 4.167, 5.14, 2.667);
}

// 3 — Brand Identity: vision / mission / value.
function slide03(s) {
	display(s, 'Brand', 8.205, 3.997, 4.568, 1.952);
	display(s, 'Identity', 8.205, 5.061, 5.451, 1.952);

	heading(s, 'Our Vision', 5.043, 2.196, 2.12);
	body(s, COPY.mountains, 5.043, 2.826, 2.591, 0.902);
	heading(s, 'Our Mission', 8.163, 2.196, 2.12);
	body(s, COPY.mountains, 8.163, 2.826, 2.591, 0.902);
	heading(s, 'Our Value', 5.043, 3.997, 2.12);
	body(s, COPY.mountains, 5.043, 4.678, 2.591, 0.902);

	asteriskPair(s, 5.043, 6.285, 1.196, 0.549);
	cornerArrow(s, 11.852, 2.362, 0.815, 0.817, 270);
	photo(s, 0.69, 2.333, 3.634, 4.5);
}

// 4 — Brand Messaging.
function slide04(s) {
	display(s, 'Brand', 0.5, 4.12, 3.737, 1.952);
	display(s, 'Messaging', 0.5, 5.12, 6.654, 1.952);

	heading(s, 'Brand Tagline', 3.816, 2.15, 2.464);
	body(s, COPY.mountains, 3.816, 2.848, 3.0, 0.902);
	heading(s, 'Brand Messaging', 7.297, 2.15, 3.214);
	body(s, COPY.wilderness, 7.297, 2.872, 5.39, 0.902);

	asteriskPair(s, 0.729, 2.354, 1.196, 0.549);
	cornerArrow(s, 12.0, 4.475, 0.666, 0.668, 0);
	photo(s, 7.316, 4.475, 3.934, 2.358);
}

// 5 — Brand Applications.
function slide05(s) {
	display(s, 'Brand', 4.58, 1.789, 3.783, 1.952);
	display(s, 'Applications', 4.58, 2.878, 7.783, 1.952);

	heading(s, 'Advertising & Marketing Campaigns', 4.58, 5.204, 6.033);
	body(s, COPY.wilderness, 4.58, 5.98, 5.7, 0.902);

	asteriskPair(s, 11.476, 5.395, 1.196, 0.549);
	cornerArrow(s, 0.667, 2.333, 0.666, 0.668, 180);
	photo(s, 0.694, 4.417, 3.078, 2.417);
}

// 6 — Brand Management: three points across the top.
function slide06(s) {
	display(s, 'Brand', 0.667, 3.955, 3.941, 1.952);
	display(s, 'Management', 0.667, 5.086, 7.703, 1.952);

	[['Point One', 2.13, 2.151, 2.731], ['Point Two', 5.861, 2.151, 2.731], ['Point Three', 9.592, 2.134, 2.749]]
		.forEach(([label, x, hy, by]) => {
			heading(s, label, x, hy, 2.12);
			body(s, COPY.critter, x, by, 3.176, 0.902);
		});

	cornerArrow(s, 0.667, 2.333, 0.666, 0.668, 180);
	asteriskPair(s, 6.707, 4.545, 1.196, 0.549);
	photo(s, 8.829, 4.545, 3.814, 2.289);
}

// 7 — Brand Maintenance.
function slide07(s) {
	display(s, 'Brand', 5.622, 1.704, 3.852, 1.952);
	display(s, 'Maintenance', 5.622, 2.804, 7.711, 1.952);
	body(s, COPY.wildernessLong, 7.229, 5.333, 5.537, 1.179);
	asteriskPair(s, 11.492, 2.337, 1.196, 0.549);
	cornerArrow(s, 5.839, 5.469, 0.666, 0.668, 0);
	photo(s, 0.69, 2.333, 4.191, 4.5);
}

// 8 — Logo Usage: breakdown pill plus three logo-construction studies.
function slide08(s) {
	display(s, 'Logo', 9.052, 1.731, 3.279, 1.952);
	display(s, 'Usage', 9.052, 2.78, 4.368, 1.952);

	s.addShape('roundRect', {
		x: 0.679, y: 3.625, w: 3.956, h: 0.685, fill: { type: 'none' },
		line: { color: DARK, width: 1 }, rectRadius: 0.3425,
	});
	heading(s, 'Logo Breakdown.', 1.295, 3.641, 3.026);
	body(s, COPY.wilderness, 5.396, 2.645, 3.003, 1.735);
	asteriskPair(s, 0.679, 2.337, 1.196, 0.549);

	[['Point One', 2.295, 1.762], ['Point Two', 6.495, 1.762], ['Point Three', 10.591, 2.155]]
		.forEach(([label, x, w]) => {
			heading(s, label, x, 5.063, w);
			body(s, COPY.mountains, x, 5.72, 2.155, 1.179);
		});

	// Full mark, then the same mark stripped back to its parts.
	logoMark(s, 0.8, 5.544, 1.037);
	notchSquare(s, 5.368, 5.893, 0.691, YELLOW);
	square(s, 5.075, 5.6, 0.586, ORANGE);
	square(s, 9.951, 5.543, 0.26, DARK);
	square(s, 9.174, 6.305, 0.26, DARK);
}

// 9 — Logo Misuse: three wrongly-assembled marks.
function slide09(s) {
	display(s, 'Logo', 0.592, 4.268, 3.026, 1.952);
	display(s, 'Misuse', 0.592, 5.329, 4.316, 1.952);

	[['Point One', 2.431, 2.168, 2.859], ['Point Two', 7.365, 2.168, 2.859], ['Point Three', 7.365, 5.029, 5.72]]
		.forEach(([label, x, hy, by]) => {
			heading(s, label, x, hy, 2.12);
			body(s, COPY.mountains, x, by, 2.182, 1.179);
		});

	// Misuse #1 — inverted colours.
	square(s, 1.423, 3.314, 0.26, ORANGE);
	square(s, 0.698, 2.59, 0.586, DARK);
	notchSquare(s, 0.991, 2.882, 0.691, YELLOW, 180);
	// Misuse #2 — rotated 45°.
	notchSquare(s, 5.916, 2.711, 0.691, ORANGE, 225);
	square(s, 6.131, 3.207, 0.26, YELLOW, 45);
	square(s, 6.123, 3.625, 0.26, DARK, 45);
	// Misuse #3 — mirrored stack.
	square(s, 5.661, 5.345, 0.586, ORANGE);
	notchSquare(s, 6.28, 5.931, 0.691, YELLOW, 180);
	square(s, 6.416, 5.538, 0.26, DARK);
	square(s, 5.864, 6.079, 0.26, DARK);

	body(s, mixedRuns([
		['In the heart of the verdant wilderness, where is the ', 12],
		['symphony', 11],
		[' of leaves rustling and birdsong intertwine', 12],
	]), 10.215, 3.931, 2.452, 1.281);
	asteriskPair(s, 10.278, 2.369, 1.196, 0.549);
	cornerArrow(s, 10.333, 6.195, 0.666, 0.668, 0);
}

// 10 — Grid System: construction diagram built from rules and blocks.
function slide10(s) {
	display(s, 'Grid System', 0.562, 1.78, 7.807, 1.952);
	heading(s, 'Precision & Balance', 0.567, 4.369, 3.661);
	body(s, COPY.mountainsLong, 0.594, 5.347, 3.113, 1.457);
	body(s, COPY.mountainsLong, 4.545, 5.347, 3.113, 1.457);
	asteriskPair(s, 4.545, 4.373, 1.196, 0.549);

	// Diagram: colour blocks first, then the ruling lines over the top.
	notchSquare(s, 9.96, 4.065, 2.146, YELLOW);
	square(s, 11.187, 2.993, 0.919, DARK);
	square(s, 8.864, 5.3, 0.911, DARK);
	square(s, 9.051, 3.156, 1.819, ORANGE);
	[8.428, 8.861, 9.048, 9.772, 9.963, 10.866, 11.184, 12.103].forEach(x => {
		if (x > 8.428) s.addShape('line', { x, y: 2.544, w: 0, h: 4.012, line: { color: DARK, width: 0.75 } });
	});
	[2.544, 2.993, 3.156, 3.912, 4.065, 5.3, 6.211].forEach(y => {
		if (y > 2.544) s.addShape('line', { x: 8.428, y, w: 4.226, h: 0, line: { color: DARK, width: 0.75 } });
	});
}

// 11 — Colour Palette: three swatch chips.
function slide11(s) {
	display(s, 'Color Palette', 5.369, 4.718, 7.75, 1.952);

	const swatches = [
		{ hex: '#E76E10', fill: ORANGE, chipX: 0.666, chipY: 2.725, textX: 2.214, titleY: 2.573, bodyY: 2.939 },
		{ hex: '#F6E300', fill: YELLOW, chipX: 4.915, chipY: 2.728, textX: 6.473, titleY: 2.601, bodyY: 2.966 },
		{ hex: '#11141B', fill: DARK, chipX: 9.13, chipY: 2.727, textX: 10.688, titleY: 2.583, bodyY: 2.949, outline: true },
	];
	swatches.forEach(sw => {
		s.addShape('roundRect', {
			x: sw.chipX, y: sw.chipY, w: 1.213, h: 1.332, fill: { color: sw.fill },
			rectRadius: 0.222, line: { color: sw.outline ? WHITE : sw.fill, width: sw.outline ? 1 : 0.01 },
		});
		text(s, sw.hex, sw.textX, sw.titleY, 2.12, 0.438,
			{ fontFace: DISPLAY, fontSize: 20, color: DARK, valign: 'top' });
		body(s, COPY.mountains, sw.textX, sw.bodyY, 2.09, 1.179);
	});

	asteriskPair(s, 0.666, 5.113, 1.196, 0.549);
	body(s, "High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. A hiker's path weaves through the alpine",
		2.478, 4.966, 2.271, 1.457);
}

// 12 — Typography: two typeface specimens plus a pairing guide.
function slide12(s) {
	display(s, 'Typography', 0.563, 2.448, 7.15, 1.952);

	[['Your  First Typeface Name', 0.583, 1.384], ['Your  Second Typeface Name', 2.616, 1.508]]
		.forEach(([name, x, w]) => {
			text(s, name, x, 5.137, w, 0.438, Object.assign({}, labelText, { fontSize: 10 }));
			heading(s, 'Aa Bb Cc', x, 5.538, 1.708);
			text(s, 'abcdefghijklmnopqrstuvwxyz', x, 6.069, 1.508, 0.471,
				{ fontFace: DISPLAY, fontSize: 11, color: DARK, valign: 'top' });
			text(s, '1234567890', x, 6.584, 1.167, 0.286,
				{ fontFace: DISPLAY, fontSize: 11, color: DARK, valign: 'top' });
		});

	text(s, 'Guideline For Pairings', 4.723, 5.138, 2.233, 0.286, labelText);
	text(s, 'Heading :', 4.731, 5.503, 1.194, 0.286, labelText);
	text(s, 'Givin', 5.839, 5.396, 1.015, 0.438,
		{ fontFace: DISPLAY, fontSize: 20, color: DARK, valign: 'top' });
	text(s, 'Sub\u2013Heading :', 4.731, 5.843, 1.194, 0.286, labelText);
	text(s, 'Your Sub Heading Here', 5.839, 5.843, 2.061, 0.286, labelText);
	text(s, 'Body :', 4.731, 6.202, 1.194, 0.286, labelText);
	body(s, 'Every rustle of a critter in the under, every gurgle of a hidden stream.', 5.839, 6.155, 2.641, 0.624);

	heading(s, 'Primary Typeface', 9.387, 2.667, 3.0);
	body(s, COPY.critterEnchanting, 9.387, 3.475, 3.238, 0.905);
	heading(s, 'Secondary Typeface', 9.387, 5.157, 3.662);
	body(s, COPY.critterEnchanting, 9.387, 5.966, 3.384, 0.905);
	asterisk(s, 8.781, 2.743, 0.435, 0.454);
	asterisk(s, 8.752, 5.233, 0.435, 0.454);
}

// 13 — Graphic elements: asterisk, arc and corner arrow catalogued.
function slide13(s) {
	display(s, 'Graphic', 5.769, 4.782, 5.125, 1.952);

	heading(s, 'Graphic Two', 2.439, 2.806, 2.612);
	body(s, COPY.mountains, 2.439, 3.527, 2.595, 0.902);
	heading(s, 'Graphic One', 7.063, 2.806, 2.612);
	body(s, COPY.mountains, 7.083, 3.527, 2.595, 0.902);
	heading(s, 'Graphic Three', 2.439, 5.211, 2.612);
	body(s, COPY.mountains, 2.439, 5.932, 2.595, 0.902);

	asterisk(s, 0.651, 3.051, 1.153, 1.202);
	cornerArrow(s, 0.679, 5.5, 1.096, 1.099, 0);
	shape(s, 'blockArc', {
		x: 5.635, y: 3.381, w: 1.041, h: 1.041, fill: { color: ORANGE },
		angleRange: [180, 1.4], arcThicknessRatio: 0.4546,
	});

	body(s, mixedRuns([
		['In the heart of the verdant wilderness, where is the ', 12],
		['symphony', 11],
		[' of leaves rustling and birdsong intertwine', 12],
	]), 10.409, 3.142, 2.239, 1.281);
	asteriskPair(s, 11.358, 5.58, 1.196, 0.549);
}

// 14 — Imagery: two wide photo plates.
function slide14(s) {
	display(s, 'Imagery', 0.526, 5.083, 5.44, 1.952);
	body(s, mixedRuns([
		['In the ', 12], ['heart', 11],
		[" of the verdant wilderness, where the symphony of leaves rustling and birdsong intertwine, nature's masterpiece unfolds in breathtaking hues. The sun's warm embrace cascades through the canopy.", 12],
	]), 6.591, 5.576, 4.358, 1.277);
	asteriskPair(s, 0.644, 4.451, 1.196, 0.549);
	cornerArrow(s, 11.988, 5.726, 0.666, 0.668, 0);
	photo(s, 2.316, 2.347, 4.885, 2.653);
	photo(s, 7.781, 2.347, 4.885, 2.653);
}

// 15 — Brand Stationery & Collateral.
function slide15(s) {
	display(s, 'Brand Stationery', 0.587, 4.274, 10.24, 1.952);
	display(s, '& Collateral', 0.587, 5.342, 7.156, 1.952);

	heading(s, 'Letterhead And', 3.692, 2.151, 2.756);
	heading(s, 'Envelope Design', 3.692, 2.544, 2.923);
	body(s, COPY.critterShort, 3.692, 3.263, 2.856, 0.902);
	heading(s, 'Business Card', 9.934, 2.151, 2.583);
	heading(s, 'Design Guideline', 9.934, 2.518, 2.833);
	body(s, COPY.critterShort, 9.911, 3.263, 2.856, 0.902);

	asteriskPair(s, 11.167, 5.068, 1.5, 0.688);
	photo(s, 0.667, 2.347, 2.583, 1.829);
	photo(s, 6.885, 2.347, 2.583, 1.829);
}

// 16 — Website & Digital Guidelines.
function slide16(s) {
	display(s, 'Website & Digital', 2.804, 1.723, 10.256, 1.952);
	display(s, 'Guidelines', 2.804, 2.829, 6.506, 1.952);

	heading(s, 'Social Media', 6.491, 4.849, 2.43);
	body(s, COPY.critterSanctuary, 6.491, 5.596, 2.916, 1.179);
	heading(s, 'Website & Ui Ux', 9.766, 4.843, 2.846);
	body(s, COPY.critterSanctuary, 9.766, 5.59, 2.916, 1.179);

	asteriskPair(s, 11.183, 3.675, 1.5, 0.688);
	cornerArrow(s, 0.651, 3.61, 0.666, 0.668, 180);
	photo(s, 0.667, 4.992, 2.243, 1.842);
	photo(s, 3.519, 4.992, 2.243, 1.842);
}

// 17 — Legal Consideration.
function slide17(s) {
	display(s, 'Legal', 4.767, 4.263, 3.613, 1.952);
	display(s, 'Consideration', 4.767, 5.369, 8.612, 1.952);

	heading(s, 'Intellectual', 4.863, 2.182, 2.02);
	heading(s, 'Property Rights', 4.863, 2.532, 2.77);
	body(s, COPY.critterSanctuary, 4.863, 3.18, 3.099, 1.179);
	heading(s, 'Disclaimer Regarding', 8.697, 2.144, 3.556);
	heading(s, 'Brand Usage', 8.697, 2.507, 2.27);
	body(s, COPY.wilderness, 8.697, 3.149, 4.094, 1.179);

	cornerArrow(s, 3.215, 6.167, 0.666, 0.668, 90);
	asteriskPair(s, 11.183, 4.846, 1.5, 0.688);
	photo(s, 0.667, 2.343, 3.215, 3.324);
}

// 18 — Our Company Leader.
function slide18(s) {
	display(s, 'Our', 0.524, 2.3, 2.605, 1.952);
	display(s, 'Company', 0.524, 3.445, 5.939, 1.952);
	display(s, 'Leader', 0.524, 4.56, 4.52, 1.952);

	body(s, COPY.critterSanctuary, 10.216, 3.056, 2.593, 1.457);
	asteriskPair(s, 4.494, 2.94, 1.5, 0.688);
	cornerArrow(s, 10.217, 5.37, 0.666, 0.668, 270);
	photo(s, 6.858, 2.343, 2.666, 3.714);
	s.addText('William Jones', Object.assign({}, headingText, { x: 6.955, y: 6.257, w: 2.473, h: 0.606, align: 'center' }));
}

// 19 — Our Superior Team.
function slide19(s) {
	display(s, 'Our', 7.794, 3.094, 2.605, 1.952);
	display(s, 'Superior', 7.794, 4.191, 5.939, 1.952);
	display(s, 'Team', 7.794, 5.353, 4.52, 1.952);

	body(s, COPY.critter, 4.259, 2.333, 2.908, 0.902);
	asteriskPair(s, 7.943, 2.343, 1.5, 0.688);
	photo(s, 4.23, 3.748, 3.0, 2.172);
	photo(s, 0.667, 2.343, 3.0, 3.577);

	[['Artful Dodger', 'Branch Manager', 0.905, 2.467, 1.202], ['Isabella Linton', 'Marketing Manager', 4.383, 2.568, 4.73]]
		.forEach(([name, role, nameX, nameW, roleX]) => {
			s.addText(name, Object.assign({}, headingText, { x: nameX, y: 6.053, w: nameW, h: 0.606, align: 'center' }));
			text(s, role, roleX, 6.565, 1.873, 0.286, Object.assign({}, labelText, { align: 'center' }));
		});
}

// 20 — Thank You / contact details.
function slide20(s) {
	display(s, 'Thank', 0.655, 3.395, 5.762, 2.625, 150);
	display(s, 'You!', 0.655, 4.766, 3.845, 2.625, 150);

	heading(s, 'Get In Touch', 6.548, 5.23, 2.491);
	[['+11 222 3333 4444', 5.868, 1.707], ['www. givincompany.com', 6.171, 2.052], ['givincompany@email.com', 6.488, 2.136]]
		.forEach(([line, y, w]) => {
			text(s, line, 6.548, y, w, 0.303, Object.assign({}, labelText, { fontSize: 12 }));
		});

	body(s, 'In the heart of the verdant wilderness, where is the symphony of leaves rustling and birdsong intertwine.',
		9.667, 6.014, 2.987, 0.902);
	cornerArrow(s, 0.679, 2.327, 0.666, 0.668, 180);
	asteriskPair(s, 9.702, 5.317, 1.202, 0.551);
	photo(s, 6.294, 2.36, 6.373, 2.406);
}

/* -------------------------------------------------------------- assembly */

const BUILDERS = [
	slide01, slide02, slide03, slide04, slide05, slide06, slide07,
	slide08, slide09, slide10, slide11, slide12, slide13, slide14,
	slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'GIVIN_16x9', width: SLIDE_W, height: SLIDE_H });
	pptx.layout = 'GIVIN_16x9';
	pptx.author = 'Givin Company';
	pptx.title = 'Givin Company — Brand Guidelines';

	BUILDERS.forEach(builder => {
		const slide = pptx.addSlide();
		slide.background = { color: WHITE };
		backgroundGrid(slide);
		pageChrome(slide);
		builder(slide);
	});

	return pptx;
}

build()
	.writeFile({ fileName: path.join(__dirname, '07eef4ff-d17b-47ab-bf8d-8f926c752e5f_grok_final.pptx') })
	.then(f => console.log('wrote', f));
