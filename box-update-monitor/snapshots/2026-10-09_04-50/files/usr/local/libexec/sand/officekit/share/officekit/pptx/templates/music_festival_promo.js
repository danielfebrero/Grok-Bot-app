/**
 * HammerTunes Music Festival — deck rebuilt with pptxgenjs.
 *
 * Run:  node 16ae03a8-5d9d-449b-a5f2-a70b48846500_grok_final.js
 * Out:  16ae03a8-5d9d-449b-a5f2-a70b48846500_grok_final.pptx (next to this file)
 *
 * Raster artwork in the source deck (logo bitmap, UI glyphs) is redrawn here
 * with native pptxgenjs shapes; empty picture placeholders are omitted because
 * they hold no image in the source.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const YELLOW = 'DBF700'; // theme accent1
const TEAL = '049F82'; // theme accent2
const MINT = '00FBCB'; // theme accent3 — used for the offset "glow" shadows
const OLIVE = 'A4B900';
const WHITE = 'FFFFFF';
const INK = '262626'; // slide background / dark text
const GREY = '404040';
const GRID = 'A6A6A6';

const HEAD = 'Michroma'; // theme major font
const BODY = 'Lato'; // theme minor font

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------------------ helpers */

// OOXML "adj" percentages are relative to the shorter side of the shape.
const radius = (adj, w, h) => (adj / 100000) * Math.min(w, h);

// The mint drop-shadow used all over the deck: a crisp 45° offset copy, no blur.
// (0.001 rather than 0 because pptxgenjs treats a falsy blur as "use the default".)
const glow = (dist) => ({ type: 'outer', color: MINT, blur: 0.001, offset: dist, angle: 45, opacity: 1 });

// Every slide sits on the same near-black background (theme tx1 @ 85% lum).
function newSlide(pptx) {
	const s = pptx.addSlide();
	s.background = { color: INK };
	return s;
}

// custGeom from points given in 0..1 shape-relative coordinates ('z' = close).
function poly(slide, x, y, w, h, pts, opts) {
	const points = pts.map((p) => (p === 'z' ? { close: true } : { x: p[0] * w, y: p[1] * h }));
	slide.addShape('custGeom', Object.assign({ x, y, w, h, points }, opts));
}

// Traces a rounded rectangle clockwise from the top-left corner.
// corners = [topLeft, topRight, bottomRight, bottomLeft] radii, in inches.
const ARC = 0.4477; // circular-arc bezier constant
function roundRectPath(ox, oy, w, h, corners, topNotch) {
	const [tl, tr, br, bl] = corners;
	const top = topNotch
		? [
				{ x: ox + topNotch.x, y: oy },
				{ x: ox + topNotch.x + topNotch.r, y: oy + topNotch.depth },
				{ x: ox + topNotch.x + topNotch.w - topNotch.r, y: oy + topNotch.depth },
				{ x: ox + topNotch.x + topNotch.w, y: oy },
			]
		: [];
	return [
		{ moveTo: true, x: ox + tl, y: oy },
		...top,
		{ x: ox + w - tr, y: oy },
		{ curve: { type: 'cubic', x1: ox + w - tr * ARC, y1: oy, x2: ox + w, y2: oy + tr * ARC }, x: ox + w, y: oy + tr },
		{ x: ox + w, y: oy + h - br },
		{ curve: { type: 'cubic', x1: ox + w, y1: oy + h - br * ARC, x2: ox + w - br * ARC, y2: oy + h }, x: ox + w - br, y: oy + h },
		{ x: ox + bl, y: oy + h },
		{ curve: { type: 'cubic', x1: ox + bl * ARC, y1: oy + h, x2: ox, y2: oy + h - bl * ARC }, x: ox, y: oy + h - bl },
		{ x: ox, y: oy + tl },
		{ curve: { type: 'cubic', x1: ox, y1: oy + tl * ARC, x2: ox + tl * ARC, y2: oy }, x: ox + tl, y: oy },
		{ close: true },
	];
}

function corneredRect(slide, x, y, w, h, corners, opts) {
	slide.addShape('custGeom', Object.assign({ x, y, w, h, points: roundRectPath(0, 0, w, h, corners) }, opts));
}

// Rounded rectangle with a rounded-rectangle hole punched through it.
function roundRectRing(slide, x, y, w, h, r, wall, opts, holeNotch) {
	const ir = Math.max(r - wall, 0);
	const points = roundRectPath(0, 0, w, h, [r, r, r, r]).concat(
		roundRectPath(wall, wall, w - 2 * wall, h - 2 * wall, [ir, ir, ir, ir], holeNotch)
	);
	slide.addShape('custGeom', Object.assign({ x, y, w, h, points }, opts));
}

// OOXML star vertices, normalised so the drawn outline exactly fills the box.
// `n` = number of points, `inner` = inner/outer radius ratio (the preset "adj").
function starPoints(n, inner) {
	const raw = [];
	for (let i = 0; i < 2 * n; i++) {
		const r = i % 2 === 0 ? 1 : inner;
		const a = -Math.PI / 2 + (i * Math.PI) / n;
		raw.push([r * Math.cos(a), r * Math.sin(a)]);
	}
	const xs = raw.map((p) => p[0]);
	const ys = raw.map((p) => p[1]);
	const x0 = Math.min(...xs);
	const y0 = Math.min(...ys);
	const sx = Math.max(...xs) - x0;
	const sy = Math.max(...ys) - y0;
	return raw.map((p) => [(p[0] - x0) / sx, (p[1] - y0) / sy]);
}

const STAR7 = starPoints(7, 0.33344).concat(['z']); // preset star7, adj 16672
const STAR5 = starPoints(5, 0.38196); // preset star5, adj 19098

// The outlined 7-point stars scattered across every slide.
function star7(slide, x, y, size, color) {
	poly(slide, x, y, size, size, STAR7, { line: { color, width: 2.5 } });
}

// Faint measuring grid: 6 x 6 hairlines inside a 4.89" square.
function grid(slide, x, y) {
	const SIZE = 4.89;
	const STEP = 0.8635;
	const line = { color: GRID, width: 0.25, transparency: 73 };
	for (let i = 0; i < 6; i++) {
		slide.addShape('line', { x: x + 0.286 + i * STEP, y: y, w: 0, h: SIZE, line });
		slide.addShape('line', { x: x, y: y + 0.286 + i * STEP, w: SIZE, h: 0, line });
	}
}

// Three small squares bottom-right; each slide rotates which one is yellow.
function pageDots(slide, colors) {
	colors.forEach((color, i) => {
		const x = 12.155 + i * 0.2655;
		corneredRect(slide, x, 7.062, 0.2, 0.2, Array(4).fill(0.0612), {
			fill: { color },
			line: { color: INK, width: 1 },
			shadow: glow(2),
		});
	});
}

// "Logoipsum" wordmark: two interlocking arrows plus the lettering.
// The source deck also drops near-black bitmap copies of it on slides 2 and 3;
// those are reproduced with markColor/textColor = INK so they stay a faint ghost.
function logo(slide, x, y, w, markColor, textColor) {
	const u = w / 1.057; // the mark is authored 1.057" wide
	const arrowA = [
		[0.31, 0.04], [0.38, 0.0], [0.74, 0.0], [0.92, 0.55], [0.57, 0.99],
		[0.52, 0.97], [0.52, 0.58], [0.56, 0.53], [0.53, 0.42], [0.0, 0.42], 'z',
	];
	const arrowB = arrowA.map((p) => (p === 'z' ? 'z' : [1 - p[0], 1 - p[1]]));
	poly(slide, x, y, 0.188 * u, 0.156 * u, arrowB, { fill: { color: markColor || YELLOW } });
	poly(slide, x + 0.09 * u, y, 0.188 * u, 0.156 * u, arrowA, { fill: { color: markColor || YELLOW } });
	slide.addText('Logoipsum', {
		x: x + 0.313 * u, y: y - 0.037 * u, w: 0.9 * u, h: 0.238 * u,
		fontFace: 'Poppins', fontSize: 9.2 * u, bold: true, color: textColor || WHITE,
		align: 'left', valign: 'middle', margin: 0, wrap: false,
	});
}

/* -------------------------------------------------------------- glyph icons */
/* Each icon is a native redraw of the small PNG glyph used in the source deck.
 * Icons are authored to fill a unit box; the source PNGs leave a transparent
 * margin, so `icon()` draws them into a centred sub-box of the same footprint.
 * fill = [widthFraction, heightFraction] measured off the original artwork.   */

function icon(draw, slide, x, y, s, color, fill) {
	const [fw, fh] = fill || [0.82, 0.82];
	draw(slide, x + ((1 - fw) / 2) * s, y + ((1 - fh) / 2) * s, fw * s, fh * s, color);
}

function iconPlay(slide, x, y, w, h, color) {
	slide.addShape('ellipse', { x, y, w, h, line: { color, width: 1.3 } });
	poly(slide, x, y, w, h, [[0.36, 0.27], [0.72, 0.5], [0.36, 0.73], 'z'], { fill: { color } });
}

function iconCheckCircle(slide, x, y, w, h, color) {
	slide.addShape('ellipse', { x, y, w, h, line: { color, width: 1.1 } });
	poly(slide, x, y, w, h, [[0.26, 0.52], [0.43, 0.7], [0.75, 0.32]], { line: { color, width: 1.1 } });
}

function iconDoubleCheck(slide, x, y, w, h, color) {
	poly(slide, x, y, w, h, [[0.0, 0.44], [0.26, 1.0]], { line: { color, width: 2.4 } });
	poly(slide, x, y, w, h, [[0.22, 0.56], [0.42, 1.0], [1.0, 0.0]], { line: { color, width: 2.4 } });
	poly(slide, x, y, w, h, [[0.44, 0.3], [0.78, 0.0]], { line: { color, width: 2.4 } });
}

// Square with a break in the top-right corner and an arrow shooting out of it.
function iconExternalLink(slide, x, y, w, h, color) {
	const t = 0.11; // stroke thickness, shape-relative
	poly(slide, x, y, w, h, [
		[0.0, 0.0], [0.47, 0.0], [0.47, t], [t, t], [t, 1 - t], [1 - t, 1 - t],
		[1 - t, 0.55], [1.0, 0.55], [1.0, 1.0], [0.0, 1.0], 'z',
	], { fill: { color } });
	poly(slide, x, y, w, h, [
		[0.56, 0.0], [1.0, 0.0], [1.0, 0.44], [0.87, 0.44], [0.87, 0.23],
		[0.32, 0.78], [0.22, 0.68], [0.77, 0.13], [0.56, 0.13], 'z',
	], { fill: { color } });
}

// Clock face with four tick marks and hands pointing at ~7:30.
function iconClock(slide, x, y, w, h, color) {
	slide.addShape('ellipse', { x, y, w, h, line: { color, width: 1.5 } });
	[[0.5, 0.02, 0.5, 0.14], [0.5, 0.86, 0.5, 0.98], [0.02, 0.5, 0.14, 0.5], [0.86, 0.5, 0.98, 0.5]].forEach((t) => {
		poly(slide, x, y, w, h, [[t[0], t[1]], [t[2], t[3]]], { line: { color, width: 1.5 } });
	});
	poly(slide, x, y, w, h, [[0.5, 0.26], [0.5, 0.54], [0.26, 0.8]], { line: { color, width: 1.5 } });
}

// Flip-chart: framed board with a trend line, on a little easel.
function iconPresentation(slide, x, y, w, h, color) {
	slide.addShape('roundRect', { x, y, w, h: h * 0.72, line: { color, width: 1.6 }, rectRadius: w * 0.09 });
	poly(slide, x, y, w, h, [[0.18, 0.48], [0.36, 0.26], [0.55, 0.44], [0.8, 0.2]], { line: { color, width: 1.5 } });
	poly(slide, x, y, w, h, [[0.35, 1.0], [0.5, 0.74], [0.65, 1.0]], { line: { color, width: 1.5 } });
}

// Three vertical rounded bars of different heights.
function iconBars(slide, x, y, w, h, color) {
	[[0.0, 0.47], [0.36, 0.0], [0.72, 0.28]].forEach(([bx, top]) => {
		slide.addShape('roundRect', {
			x: x + bx * w, y: y + top * h, w: w * 0.28, h: h * (1 - top),
			line: { color, width: 1.6 }, rectRadius: w * 0.14,
		});
	});
}

// Three rounded squares plus a magnifier in the lower-left cell.
function iconGridSearch(slide, x, y, w, h, color) {
	[[0.0, 0.0], [0.58, 0.0], [0.58, 0.58]].forEach(([bx, by]) => {
		slide.addShape('roundRect', { x: x + bx * w, y: y + by * h, w: w * 0.42, h: h * 0.42, line: { color, width: 2 }, rectRadius: w * 0.15 });
	});
	slide.addShape('ellipse', { x: x + 0.02 * w, y: y + 0.58 * h, w: w * 0.32, h: h * 0.32, line: { color, width: 2 } });
	poly(slide, x, y, w, h, [[0.29, 0.87], [0.4, 1.0]], { line: { color, width: 2 } });
}

// Two stacked rounded pages with a plus sign on the front one.
function iconAddPage(slide, x, y, w, h, color) {
	slide.addShape('roundRect', { x: x + 0.22 * w, y: y + 0.22 * h, w: w * 0.78, h: h * 0.78, fill: { color }, rectRadius: w * 0.18 });
	slide.addShape('roundRect', { x: x + 0.16 * w, y: y + 0.16 * h, w: w * 0.68, h: h * 0.68, fill: { color: INK }, rectRadius: w * 0.16 });
	slide.addShape('roundRect', { x, y, w: w * 0.72, h: h * 0.72, fill: { color }, rectRadius: w * 0.16 });
	slide.addShape('rect', { x: x + 0.14 * w, y: y + 0.31 * h, w: w * 0.44, h: h * 0.1, fill: { color: INK } });
	slide.addShape('rect', { x: x + 0.31 * w, y: y + 0.14 * h, w: w * 0.1, h: h * 0.44, fill: { color: INK } });
}

function iconStar(slide, x, y, w, h, opts) {
	poly(slide, x, y, w, h, STAR5.concat(['z']), opts);
}

// 4.5-of-5 rating mark: outlined star with only its left half filled.
function iconHalfStar(slide, x, y, w, h, color) {
	poly(slide, x, y, w, h, STAR5.concat(['z']), { line: { color, width: 1 } });
	poly(slide, x, y, w, h, [STAR5[0], STAR5[9], STAR5[8], STAR5[7], STAR5[6], STAR5[5], 'z'], { fill: { color } });
}

// Support headset: arc over two rounded ear cups.
function iconHeadset(slide, x, y, w, h, color) {
	slide.addShape('arc', { x: x + 0.08 * w, y: y + 0.04 * h, w: w * 0.84, h: h * 1.1, angleRange: [180, 360], line: { color, width: 2 } });
	poly(slide, x, y, w, h, [[0.92, 0.5], [0.92, 1.0], [0.55, 1.0]], { line: { color, width: 2 } });
	slide.addShape('roundRect', { x, y: y + 0.42 * h, w: w * 0.2, h: h * 0.42, line: { color, width: 2 }, rectRadius: w * 0.08 });
	[0.36, 0.62].forEach((ex) => poly(slide, x, y, w, h, [[ex, 0.44], [ex + 0.06, 0.5]], { line: { color, width: 1.8 } }));
	poly(slide, x, y, w, h, [[0.4, 0.62], [0.5, 0.7], [0.6, 0.62]], { line: { color, width: 1.8 } });
}

// Map pin: filled teardrop with a house knocked out of it.
function iconPin(slide, x, y, w, h, color) {
	poly(slide, x, y, w, h, [
		[0.5, 0.0], [0.85, 0.14], [1.0, 0.45], [0.5, 1.0], [0.0, 0.45], [0.15, 0.14], 'z',
	], { fill: { color } });
	slide.addShape('ellipse', { x, y, w, h: h * 0.78, fill: { color } });
	poly(slide, x, y, w, h, [
		[0.5, 0.2], [0.75, 0.4], [0.75, 0.58], [0.25, 0.58], [0.25, 0.4], 'z',
	], { fill: { color: INK } });
	slide.addShape('rect', { x: x + 0.42 * w, y: y + 0.46 * h, w: w * 0.16, h: h * 0.12, fill: { color } });
}

// Globe: circle crossed by two latitudes and one ellipse meridian.
function iconGlobe(slide, x, y, w, h, color) {
	slide.addShape('ellipse', { x, y, w, h, line: { color, width: 1.8 } });
	slide.addShape('ellipse', { x: x + 0.28 * w, y, w: w * 0.44, h, line: { color, width: 1.6 } });
	slide.addShape('line', { x: x + 0.03 * w, y: y + 0.32 * h, w: w * 0.94, h: 0, line: { color, width: 1.6 } });
	slide.addShape('line', { x: x + 0.03 * w, y: y + 0.68 * h, w: w * 0.94, h: 0, line: { color, width: 1.6 } });
}

/* ------------------------------------------------------------------- slides */

const LOREM_LONG =
	'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';

function slide1(pptx) {
	const s = newSlide(pptx);
	grid(s, 5.448, 0.438);

	// Two tall bars on the right: yellow (rounded top) behind teal (rounded bottom).
	corneredRect(s, 7.76, 0.601, 2.441, 6.899, [0.221, 0.221, 0, 0], { fill: { color: YELLOW } });
	corneredRect(s, 10.337, 0.0, 2.441, 6.899, [0, 0, 0.204, 0.204], { fill: { color: TEAL } });

	// "Start Presentation" pill
	corneredRect(s, 0.913, 5.327, 2.147, 0.54, Array(4).fill(0.27), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(3),
	});
	icon(iconPlay, s, 1.079, 5.426, 0.342, INK, [0.833, 0.833]);
	s.addText('Start Presentation', {
		x: 1.421, y: 5.454, w: 1.447, h: 0.286,
		fontFace: BODY, fontSize: 11, color: INK, valign: 'top', wrap: false,
	});

	s.addText(LOREM_LONG, {
		x: 0.855, y: 4.321, w: 5.027, h: 0.62,
		fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	});

	s.addText('HammerTunes Music Festival', {
		x: 0.855, y: 1.703, w: 6.755, h: 1.717,
		fontFace: HEAD, fontSize: 48, color: WHITE, valign: 'top',
	});

	logo(s, 0.602, 0.438, 1.057);
	star7(s, 7.156, 5.718, 1.057, TEAL);
	star7(s, 4.875, 0.324, 0.947, YELLOW);
	pageDots(s, [YELLOW, TEAL, TEAL]);
}

function slide2(pptx) {
	const s = newSlide(pptx);
	grid(s, 6.057, 2.776);
	star7(s, 5.153, 0.863, 1.057, TEAL);
	star7(s, 11.99, 3.716, 0.562, YELLOW);

	s.addText('A Festival Beyond Sound', {
		x: 6.71, y: 1.463, w: 6.308, h: 1.313, fontFace: HEAD, fontSize: 36, color: WHITE, valign: 'top',
	});
	s.addText('90%', { x: 6.726, y: 3.464, w: 1.861, h: 0.572, fontFace: HEAD, fontSize: 28, color: WHITE, valign: 'top' });
	s.addText('Great Subtitle', { x: 6.726, y: 4.036, w: 2.047, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, valign: 'top' });
	s.addText(
		'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam',
		{ x: 6.71, y: 4.595, w: 4.347, h: 0.903, fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top' }
	);

	// "See More" pill
	corneredRect(s, 6.822, 5.812, 1.181, 0.332, Array(4).fill(0.166), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(2),
	});
	s.addText('See More', {
		x: 6.903, y: 5.835, w: 1.019, h: 0.286,
		fontFace: BODY, fontSize: 11, italic: true, color: INK, align: 'center', valign: 'top',
	});

	logo(s, 0.688, 0.525, 1.072, INK, INK); // near-black bitmap copy in the source
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [TEAL, YELLOW, TEAL]);

	// Statistic bubble
	s.addShape('ellipse', {
		x: 5.099, y: 2.776, w: 1.23, h: 1.23,
		fill: { color: TEAL }, line: { color: INK, width: 0 }, shadow: glow(3),
	});
	s.addText('285', { x: 5.189, y: 3.026, w: 1.057, h: 0.438, fontFace: HEAD, fontSize: 20, color: WHITE, align: 'center', valign: 'top' });
	s.addText('Lorem ipsum', { x: 4.958, y: 3.43, w: 1.513, h: 0.286, fontFace: BODY, fontSize: 11, color: WHITE, align: 'center', valign: 'top' });
}

function slide3(pptx) {
	const s = newSlide(pptx);
	grid(s, 8.444, 0.51);

	s.addText('Where Music Meets Magic', {
		x: 1.552, y: 1.126, w: 10.012, h: 0.774, fontFace: HEAD, fontSize: 40, color: WHITE, valign: 'top',
	});
	const para =
		'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, labore et dolor.';
	[1.552, 7.108].forEach((x) => {
		s.addText(para, { x, y: 2.406, w: 4.792, h: 0.907, fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top' });
	});

	// Highlight card
	corneredRect(s, 2.677, 3.75, 3.824, 2.897, Array(4).fill(radius(5976, 3.824, 2.897)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(8.09),
	});
	iconStar(s, 3.134, 4.276, 0.467, 0.447, { line: { color: INK, width: 2.4 } });
	s.addText('Subtitle Here', { x: 2.978, y: 4.962, w: 1.975, h: 0.337, fontFace: HEAD, fontSize: 14, color: GREY, valign: 'top' });
	s.addText('Lorem ipsum dolor sit amet, consectetur adicing elit, sed do eiusmod tempor incididunt. sed do eiusmod tempor', {
		x: 2.978, y: 5.298, w: 3.327, h: 0.897, fontFace: BODY, fontSize: 11, color: GREY, lineSpacingMultiple: 1.5, valign: 'top',
	});

	logo(s, 0.727, 0.473, 1.074, INK, INK); // near-black bitmap copy in the source
	star7(s, 1.966, 5.667, 0.907, TEAL);
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [TEAL, TEAL, YELLOW]);
	star7(s, 11.621, 3.179, 0.706, YELLOW);
}

function slide4(pptx) {
	const s = newSlide(pptx);
	grid(s, 7.155, 2.088);

	// Services 3 & 4 (right half)
	const rightCards = [
		{ x: 7.157, title: 'Service 03', tx: 7.495, icon: iconPresentation, ix: 8.319 },
		{ x: 9.971, title: 'Service 04', tx: 10.309, icon: iconGridSearch, ix: 11.133 },
	];
	rightCards.forEach((c) => {
		s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing', {
			x: c.x, y: 5.509, w: 2.78, h: 0.63, fontFace: BODY, fontSize: 11, color: WHITE,
			align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
		});
		s.addText(c.title, { x: c.tx, y: 5.139, w: 2.105, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'center', valign: 'top' });
	});
	icon(iconGridSearch, s, 11.133, 4.505, 0.456, WHITE, [0.812, 0.812]);
	icon(iconPresentation, s, 8.319, 4.505, 0.456, WHITE, [0.812, 0.729]);

	pageDots(s, [YELLOW, TEAL, TEAL]);

	// Darkening band behind the header
	s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 3.153, fill: { color: '0D0D0D', transparency: 48.5 } });

	logo(s, 0.602, 0.438, 1.057);
	s.addText('Our Service', { x: 3.866, y: 1.026, w: 5.734, h: 0.774, fontFace: HEAD, fontSize: 40, color: WHITE, align: 'center', valign: 'top' });
	s.addText('Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim', {
		x: 3.775, y: 1.955, w: 5.915, h: 0.667, fontFace: BODY, fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
	});
	star7(s, 11.758, 2.858, 0.993, TEAL);

	// Service 01
	icon(iconClock, s, 1.744, 4.505, 0.456, WHITE, [0.812, 0.812]);
	s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing', {
		x: 0.583, y: 5.509, w: 2.78, h: 0.63, fontFace: BODY, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
	});
	s.addText('Service 01', { x: 0.92, y: 5.139, w: 2.105, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'center', valign: 'top' });

	// Service 02 — the highlighted yellow card
	corneredRect(s, 3.592, 2.976, 3.319, 3.153, Array(4).fill(radius(6153, 3.319, 3.153)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(8.09),
	});
	icon(iconBars, s, 5.024, 3.417, 0.456, INK, [0.812, 0.771]);
	s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing Lorem ipsum dolor sit amet, consectetur adipisicing', {
		x: 3.862, y: 4.633, w: 2.78, h: 0.907, fontFace: BODY, fontSize: 11, color: GREY, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
	});
	s.addText('Service 02', { x: 4.199, y: 4.162, w: 2.105, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true, color: INK, align: 'center', valign: 'top' });
	star7(s, 0.575, 3.438, 0.561, YELLOW);
}

function slide5(pptx) {
	const s = newSlide(pptx);
	grid(s, 4.222, 0.861);

	const artists = [
		{ x: 1.322, inner: 1.496, name: 'Fernanda Doe', nx: 0.882, nw: 3.1, dx: 1.785, dw: 1.292, detail: 'Add Detail Here' },
		{ x: 4.145, inner: 4.323, name: 'Ariana Doe', nx: 3.987, nw: 2.535, dx: 4.609, dw: 1.292, detail: 'Add Detail Here' },
		{ x: 6.969, inner: 7.139, name: 'Donsen Doe', nx: 6.687, nw: 2.783, dx: 7.432, dw: 1.292, detail: 'Add Detail Here' },
		{ x: 9.792, inner: 9.967, name: 'Graham Doe', nx: 9.483, nw: 2.838, dx: 10.295, dw: 1.214, detail: 'Add Detil Here' },
	];
	artists.forEach((a) => s.addShape('ellipse', { x: a.x, y: 3.253, w: 2.22, h: 2.22, fill: { color: YELLOW } }));
	artists.forEach((a) => s.addShape('ellipse', { x: a.inner, y: 3.253, w: 1.872, h: 1.872, fill: { color: OLIVE } }));

	s.addText('Your Favorite Artists', { x: 2.519, y: 1.079, w: 8.296, h: 0.774, fontFace: HEAD, fontSize: 40, color: WHITE, align: 'center', valign: 'top' });
	s.addText('Lorem ipsum dolor sit amet, consectetur', { x: 4.878, y: 1.712, w: 3.576, h: 0.286, fontFace: BODY, fontSize: 11, color: INK, align: 'center', valign: 'top' });

	artists.forEach((a) => {
		s.addText(a.detail, { x: a.dx, y: 6.088, w: a.dw, h: 0.286, fontFace: BODY, fontSize: 11, color: WHITE, align: 'center', valign: 'top', wrap: false });
		s.addText(a.name, { x: a.nx, y: 5.751, w: a.nw, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'center', valign: 'top' });
	});

	star7(s, 0.464, 5.751, 0.724, TEAL);
	star7(s, 11.914, 0.724, 0.947, YELLOW);
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [TEAL, YELLOW, TEAL]);
}

function slide6(pptx) {
	const s = newSlide(pptx);
	grid(s, 0.472, 0.951);

	s.addText('Festival Highlights', { x: 0.713, y: 0.951, w: 7.489, h: 0.774, fontFace: HEAD, fontSize: 40, color: WHITE, valign: 'top' });
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [TEAL, TEAL, YELLOW]);

	// Arrow pill
	corneredRect(s, 7.67, 1.228, 1.065, 0.303, Array(4).fill(0.1515), {
		fill: { color: TEAL }, line: { color: INK, width: 0 }, shadow: glow(2),
	});
	s.addShape('line', {
		x: 7.891, y: 1.375, w: 0.622, h: 0,
		line: { color: WHITE, width: 1, endArrowType: 'triangle' },
	});

	// "301+ Moment Capture" card
	corneredRect(s, 10.324, 1.375, 2.183, 1.294, Array(4).fill(radius(11940, 2.183, 1.294)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(8.09),
	});
	s.addText('301+', { x: 10.538, y: 1.637, w: 1.464, h: 0.505, fontFace: HEAD, fontSize: 24, bold: true, color: INK, valign: 'top' });
	s.addText('Moment Capture', { x: 10.538, y: 2.087, w: 1.754, h: 0.303, fontFace: BODY, fontSize: 12, color: INK, valign: 'top' });

	s.addShape('ellipse', { x: 11.877, y: 1.123, w: 0.639, h: 0.639, fill: { color: INK }, line: { color: WHITE, width: 1.5 } });
	icon(iconAddPage, s, 12.003, 1.248, 0.389, WHITE, [0.812, 0.812]);

	star7(s, 7.67, 2.109, 0.944, TEAL);
	star7(s, 1.08, 5.873, 0.908, YELLOW);
}

function slide7(pptx) {
	const s = newSlide(pptx);
	grid(s, 8.099, 0.411);

	s.addText('Ticket Options ', { x: 1.258, y: 1.095, w: 5.572, h: 0.774, fontFace: HEAD, fontSize: 40, color: WHITE, valign: 'top' });
	s.addText('Lorem ipsum dolor sit amet, consectetur', {
		x: 1.462, y: 1.905, w: 4.158, h: 0.352, fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	});

	const tiers = [
		{ name: 'Standard', price: '$450', bar: { y: 2.851, h: 1.031, adj: 9235, fill: YELLOW }, ink: INK, divider: '6E7C00', members: 'Up to 2 Members', rowY: 3.062, nameY: 3.182, priceY: 3.08, linkY: 3.172 },
		{ name: 'Premium', price: '$617', bar: { y: 4.101, h: 1.2, adj: 6144, fill: TEAL }, ink: WHITE, divider: '34FAD5', members: 'Up to 10 Members', rowY: 4.391, nameY: 4.511, priceY: 4.41, linkY: 4.502 },
		{ name: 'Expert', price: '$850', bar: { y: 5.519, h: 1.2, adj: 6144, fill: TEAL }, ink: WHITE, divider: '34FAD5', members: 'Up to 50 Members', rowY: 5.81, nameY: 5.929, priceY: 5.828, linkY: 5.92 },
	];

	tiers.forEach((t) => {
		corneredRect(s, 1.641, t.bar.y, 10.047, t.bar.h, Array(4).fill(radius(t.bar.adj, 10.047, t.bar.h)), {
			fill: { color: t.bar.fill }, line: { color: INK, width: 0 }, shadow: glow(5),
		});
	});

	tiers.forEach((t) => {
		s.addText(t.name, { x: 2.047, y: t.nameY, w: 1.394, h: 0.337, fontFace: HEAD, fontSize: 14, color: t.ink, valign: 'top', wrap: false });
		s.addText(t.price, { x: 3.623, y: t.priceY, w: 1.9, h: 0.572, fontFace: HEAD, fontSize: 28, color: t.ink, valign: 'top' });
		s.addShape('line', { x: 5.809, y: t.rowY, w: 0, h: 0.629, line: { color: t.divider, width: 1, transparency: 60 } });

		const features = [
			[6.383, 'Great Facilities', 'Lorem Ipsum'],
			[8.581, 'Dolor sit amet', t.members],
		];
		features.forEach(([cx, top, bottom]) => {
			[[t.rowY, top], [t.rowY + 0.379, bottom]].forEach(([iy, label]) => {
				icon(iconCheckCircle, s, cx, iy, 0.25, t.ink, [0.833, 0.833]);
				s.addText(label, { x: cx + 0.25, y: iy - 0.024, w: 1.4, h: 0.286, fontFace: BODY, fontSize: 11, italic: true, color: t.ink, valign: 'top', wrap: false });
			});
		});
		icon(iconExternalLink, s, 10.998, t.linkY, 0.388, t.ink, [0.75, 0.75]);
	});

	star7(s, 9.734, 1.328, 1.057, TEAL);
	star7(s, 11.555, 0.568, 0.947, YELLOW);
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [YELLOW, TEAL, TEAL]);
}

function slide8(pptx) {
	const s = newSlide(pptx);
	grid(s, 0.973, 1.305);

	// Teal panel behind the phone (round2SameRect rotated 270° in the source).
	corneredRect(s, 7.414, 1.87, 5.92, 4.085, [0.179, 0, 0, 0.179], {
		fill: { color: TEAL }, line: { color: INK, width: 0 }, shadow: glow(5),
	});

	s.addText('Bringing the Festival to Life:', { x: 1.594, y: 1.425, w: 4.562, h: 2.12, fontFace: HEAD, fontSize: 40, color: WHITE, valign: 'top' });
	s.addText('Explore More', { x: 3.326, y: 5.968, w: 1.071, h: 0.286, fontFace: BODY, fontSize: 11, color: WHITE, valign: 'top', wrap: false });
	s.addText('891+ Subscriber', { x: 3.358, y: 5.678, w: 2.49, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, valign: 'top' });
	s.addText('Project management is the application of knowledge, skills, tools, ', {
		x: 1.591, y: 3.666, w: 3.533, h: 0.63, fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	});

	star7(s, 5.416, 0.368, 1.057, TEAL);
	pageDots(s, [TEAL, YELLOW, TEAL]);
	logo(s, 0.602, 0.438, 1.057);

	// Phone mock-up: three concentric bezel rings left open in the middle so
	// the teal panel behind shows through, exactly as in the source deck.
	[[8.097, 1.622, 0.227], [8.097, 2.05, 0.436], [8.097, 2.589, 0.436], [11.03, 2.204, 0.671]].forEach(([bx, by, bh]) => {
		s.addShape('rect', { x: bx, y: by, w: 0.053, h: bh, fill: { color: '919191' } });
	});
	roundRectRing(s, 8.126, 0.827, 2.93, 5.931, 0.44, 0.04, { fill: { color: '414041' } });
	roundRectRing(s, 8.166, 0.875, 2.843, 5.836, 0.4, 0.023, { fill: { color: '2C2B2C' } });
	roundRectRing(s, 8.189, 0.896, 2.803, 5.794, 0.38, 0.105, { fill: { color: '131313' } },
		{ x: 0.79, w: 1.01, depth: 0.206, r: 0.05 });
	s.addShape('roundRect', { x: 9.404, y: 1.029, w: 0.368, h: 0.058, fill: { color: '414041' }, rectRadius: 0.029 });
	s.addShape('ellipse', { x: 9.856, y: 1.008, w: 0.099, h: 0.099, fill: { color: '414041' } });

	star7(s, 11.702, 1.546, 0.947, YELLOW);

	// "HD Resolution" badge
	corneredRect(s, 6.667, 3.709, 2.046, 0.587, Array(4).fill(radius(18387, 2.046, 0.587)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(4),
	});
	icon(iconDoubleCheck, s, 6.902, 3.824, 0.358, INK, [0.917, 0.5]);
	s.addText('HD Resolution', { x: 7.323, y: 3.859, w: 1.364, h: 0.286, fontFace: BODY, fontSize: 11, italic: true, color: INK, valign: 'top' });

	// Rating badge
	corneredRect(s, 10.499, 5.161, 2.163, 0.901, Array(4).fill(radius(8887, 2.163, 0.901)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(4),
	});
	[10.688, 10.953, 11.217, 11.482].forEach((sx) => iconStar(s, sx, 5.344, 0.25, 0.25, { fill: { color: '000000' } }));
	iconHalfStar(s, 11.746, 5.344, 0.25, 0.25, '000000');
	s.addText('100+ Total Reviews', { x: 10.62, y: 5.637, w: 1.596, h: 0.286, fontFace: BODY, fontSize: 11, color: INK, valign: 'top', wrap: false });
	s.addText('4.8', { x: 12.037, y: 5.3, w: 0.503, h: 0.337, fontFace: 'Montserrat Medium', fontSize: 14, color: INK, valign: 'top', wrap: false });
}

function slide9(pptx) {
	const s = newSlide(pptx);
	grid(s, 7.476, 1.373);

	corneredRect(s, 4.598, 1.319, 2.308, 2.368, Array(4).fill(radius(8134, 2.308, 2.368)), {
		fill: { color: YELLOW }, line: { color: INK, width: 0 }, shadow: glow(4),
	});
	corneredRect(s, 1.644, 3.983, 5.262, 2.478, Array(4).fill(radius(6615, 5.262, 2.478)), {
		fill: { color: TEAL }, line: { color: INK, width: 0 }, shadow: glow(4),
	});

	s.addText([{ text: 'Let\u2019s Connect: We\u2019re Here ', options: { breakLine: true } }, { text: 'to Help' }], {
		x: 7.738, y: 1.238, w: 4.367, h: 2.794, fontFace: HEAD, fontSize: 40, color: WHITE, valign: 'top',
	});
	s.addText('Contact Our Customer Service', { x: 7.738, y: 4.424, w: 3.726, h: 0.337, fontFace: HEAD, fontSize: 14, color: WHITE, valign: 'top' });
	s.addText('Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim', {
		x: 7.738, y: 5.04, w: 3.726, h: 0.981, fontFace: BODY, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	});

	s.addText('Exceptional Client Satisfaction and Service Excellence', {
		x: 4.757, y: 2.397, w: 2.012, h: 1.111, fontFace: HEAD, fontSize: 12, color: INK, valign: 'top',
	});
	s.addShape('ellipse', { x: 6.067, y: 1.379, w: 0.772, h: 0.772, fill: { color: INK } });
	s.addText([{ text: '24', options: { fontSize: 14 } }, { text: 'hr', options: { fontSize: 11 } }], {
		x: 5.95, y: 1.597, w: 1.006, h: 0.337, fontFace: HEAD, color: WHITE, align: 'center', valign: 'top',
	});

	const rows = [
		{ y: 4.42, label: 'Phone:', value: '+(12) 345 678 910', icon: iconHeadset, color: YELLOW, fill: [0.838, 0.75] },
		{ y: 5.03, label: 'Website:', value: 'www.yourgreatsite.com', icon: iconGlobe, color: 'FF4A24', fill: [0.838, 0.838] },
		{ y: 5.64, label: 'Address:', value: 'PUD Street, No. 4, Nganjuk', icon: iconPin, color: 'FF4A24', fill: [0.662, 0.812] },
	];
	rows.forEach((r) => {
		icon(r.icon, s, 2.161, r.y + 0.03, 0.304, r.color, r.fill);
		s.addText(r.label, { x: 2.623, y: r.y + 0.014, w: 1.345, h: 0.337, fontFace: HEAD, fontSize: 14, bold: true, color: WHITE, valign: 'top' });
		s.addText(r.value, { x: 4.221, y: r.y, w: 2.271, h: 0.366, fontFace: BODY, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top' });
	});

	star7(s, 11.933, 0.956, 1.057, TEAL);
	star7(s, 1.388, 6.075, 0.772, YELLOW);
	logo(s, 0.602, 0.438, 1.057);
	pageDots(s, [TEAL, TEAL, YELLOW]);
}

function slide10(pptx) {
	const s = newSlide(pptx);
	grid(s, -0.165, 2.213);

	s.addText([{ text: 'Thank ' }, { text: 'You.', options: { bold: true } }], {
		x: 0.556, y: 4.362, w: 9.444, h: 1.582, fontFace: HEAD, fontSize: 88, color: WHITE, valign: 'top',
	});
	s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore', {
		x: 10.224, y: 4.769, w: 2.553, h: 1.175, fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	});

	star7(s, 5.062, 6.153, 0.764, TEAL);
	pageDots(s, [YELLOW, TEAL, TEAL]);
	logo(s, 0.602, 0.438, 1.057);
	star7(s, 11.065, 1.978, 0.739, YELLOW);
}

/* --------------------------------------------------------------------- main */

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'CUSTOM_16x9', width: SLIDE_W, height: SLIDE_H });
	pptx.layout = 'CUSTOM_16x9';
	pptx.title = 'HammerTunes Music Festival';

	const builders = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10];
	builders.forEach((fn) => fn(pptx));

	return pptx.writeFile({
		fileName: path.join(__dirname, '16ae03a8-5d9d-449b-a5f2-a70b48846500_grok_final.pptx'),
	});
}

build().then((f) => console.log('wrote', f));
