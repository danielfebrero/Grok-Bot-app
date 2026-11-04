/**
 * GLITTERATI — fashion deck (42 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs only. Photographs in the original are replaced by
 * flat placeholder rectangles labelled "[image]".
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const DARK = '262626';   // near-black used for panels and headings
const WHITE = 'FFFFFF';
const GREY = '808080';   // body copy on light slides
const GREY2 = '595959';  // secondary labels
const GREY3 = 'A6A6A6';  // footer url on light slides
const SILVER = 'D9D9D9'; // light grey panel
const SAND = 'F9D59F';   // pale accent block
const GOLD = 'F7C06F';   // price tag / progress bar
const AMBER = 'F1960F';
const AMBER2 = 'B5710B';
const AMBER3 = '794B07';
const PLACEHOLDER = 'CFCFCF'; // stand-in for photographs

/* -------------------------------------------------------------------- fonts */
const HEAD = 'Montserrat SemiBold';
const BODY = 'Lato';
const SERIF = 'Unna';
const UI = 'Roboto';

/* ------------------------------------------------------------- filler texts */
const L1 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip exeasos commodo consequat. Duis aute irure.';
const L2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.';
const L3 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam, quis.';
const M1 = 'Make a type specimen book unknown printer took a galley of type and scrambled it to make.';
const M2 = 'PLACEHOLDER';
const M3 = 'Make a type specimen book unknown printer took a galley of type and scrambled it to make a type.';
const M4 = 'Make a type specimen book unknown printer took a galley of type on.';
const M5 = 'Make a type specimen book unkno printer took a galley of type and scrambled it to make a type.';
const Q_RALPH = '\u201CFashion is not necessarily about labels. It\u2019s not about brands. It\u2019s about something else that comes from within you.\u201D';
const Q_COCO = '\u201CFashion is not something that exists in\u00A0dresses\u00A0only. Fashion is in the sky, in the street, fashion has to do with ideas, the way we live, what is happenin.\u201D';
const Q_KORS = '\u201CI know what women look good in. I don\u2019t think the rules ever change.\u201D ';
const URL = 'WWW.GLITTERATI.COM';

/* ------------------------------------------------------------------ helpers */
const NOLINE = { type: 'none' };

/** Solid rectangle. `alpha` is pptxgenjs transparency (0-100). */
function rect (s, x, y, w, h, color, alpha) {
	s.addShape('rect', { x, y, w, h, fill: { color, transparency: alpha || 0 }, line: NOLINE });
}

/** Hollow rectangular ring of wall thickness `t` (the deck's "Frame" shape). */
function frame (s, x, y, w, h, t, color) {
	s.addShape('custGeom', {
		x, y, w, h, fill: { color }, line: NOLINE,
		points: [
			{ x: 0, y: 0, moveTo: true }, { x: w, y: 0 }, { x: w, y: h }, { x: 0, y: h }, { close: true },
			{ x: t, y: t, moveTo: true }, { x: t, y: h - t }, { x: w - t, y: h - t }, { x: w - t, y: t }, { close: true }
		]
	});
}

/** Frame whose left/right walls are open in the middle (used on the break slide). */
function bracketFrame (s, x, y, w, h, t, leg, color) {
	rect(s, x, y, w, t, color);
	rect(s, x, y + h - t, w, t, color);
	[y, y + h - leg].forEach(yy => { rect(s, x, yy, t, leg, color); rect(s, x + w - t, yy, t, leg, color); });
}

/** Text block. `lines` is a string or an array of strings (one paragraph each). */
function txt (s, lines, o) {
	const arr = Array.isArray(lines) ? lines : [lines];
	const runs = arr.map((t, i) => ({ text: t, options: { breakLine: i < arr.length - 1 } }));
	// margin is [left, right, bottom, top] in points — PowerPoint's default text insets
	s.addText(runs, Object.assign({ valign: 'top', fontFace: HEAD, color: DARK, margin: [7.2, 7.2, 3.6, 3.6] }, o));
}

/** Grey caption footer that appears on every content slide. */
function footer (s, x, align, color, y, extra) {
	txt(s, URL, Object.assign({
		x, y: y === undefined ? 6.831 : y, w: 3.544, h: 0.252,
		fontFace: BODY, fontSize: 9, charSpacing: 6, color: color || GREY3, align: align || 'left'
	}, extra || {}));
}

/** Dress-form logo mark: body outline, stand, head. Normalised to a 1x1 box. */
const FORM_BODY = [['M', 0.96, 0.268], ['C', 0.955, 0.279, 0.942, 0.295, 0.929, 0.297], ['C', 0.854, 0.306, 0.851, 0.347, 0.823, 0.381], ['C', 0.707, 0.517, 0.748, 0.64, 0.834, 0.779], ['C', 0.941, 0.951, 0.9, 0.955, 0.654, 0.982], ['C', 0.495, 1.0, 0.334, 0.994, 0.179, 0.959], ['C', 0.125, 0.946, 0.092, 0.93, 0.112, 0.889], ['C', 0.147, 0.819, 0.173, 0.746, 0.218, 0.678], ['C', 0.265, 0.608, 0.267, 0.542, 0.222, 0.473], ['C', 0.196, 0.433, 0.179, 0.391, 0.152, 0.352], ['C', 0.139, 0.334, 0.115, 0.31, 0.089, 0.305], ['C', 0.0, 0.288, 0.024, 0.242, 0.021, 0.204], ['C', 0.017, 0.165, 0.038, 0.135, 0.107, 0.129], ['C', 0.141, 0.126, 0.174, 0.119, 0.207, 0.113], ['C', 0.242, 0.107, 0.278, 0.102, 0.308, 0.093], ['C', 0.297, 0.112, 0.289, 0.183, 0.295, 0.387], ['C', 0.295, 0.387, 0.286, 0.846, 0.255, 0.893], ['L', 0.292, 0.899], ['C', 0.292, 0.899, 0.332, 0.929, 0.332, 0.303], ['C', 0.332, 0.303, 0.322, 0.144, 0.339, 0.082], ['C', 0.373, 0.066, 0.393, 0.042, 0.386, 0.0], ['L', 0.602, 0.0], ['C', 0.597, 0.037, 0.614, 0.061, 0.643, 0.077], ['L', 0.643, 0.077], ['C', 0.706, 0.161, 0.643, 0.53, 0.643, 0.53], ['L', 0.674, 0.534], ['C', 0.738, 0.418, 0.691, 0.123, 0.686, 0.094], ['C', 0.715, 0.102, 0.747, 0.108, 0.78, 0.113], ['C', 0.799, 0.116, 0.819, 0.119, 0.838, 0.122], ['C', 0.964, 0.144, 1.0, 0.186, 0.96, 0.268]];
const FORM_STAND = [['M', 0.28, 0.755], ['C', 0.236, 0.754, 0.195, 0.749, 0.159, 0.738], ['C', 0.084, 0.716, 0.116, 0.643, 0.197, 0.649], ['C', 0.261, 0.653, 0.327, 0.658, 0.397, 0.663], ['L', 0.397, 0.0], ['L', 0.596, 0.0], ['L', 0.596, 0.662], ['C', 0.665, 0.658, 0.729, 0.654, 0.789, 0.649], ['C', 0.874, 0.644, 0.907, 0.721, 0.828, 0.743], ['C', 0.789, 0.754, 0.745, 0.761, 0.7, 0.762], ['C', 0.682, 0.762, 0.676, 0.777, 0.69, 0.783], ['C', 0.777, 0.825, 0.864, 0.867, 0.948, 0.912], ['C', 0.954, 0.915, 0.959, 0.919, 0.963, 0.923], ['C', 1.0, 0.954, 0.957, 1.0, 0.898, 0.992], ['C', 0.891, 0.991, 0.886, 0.989, 0.88, 0.987], ['C', 0.619, 0.86, 0.366, 0.866, 0.108, 0.99], ['C', 0.103, 0.993, 0.096, 0.994, 0.089, 0.994], ['C', 0.035, 0.998, 0.0, 0.958, 0.029, 0.928], ['C', 0.034, 0.923, 0.039, 0.919, 0.046, 0.916], ['C', 0.128, 0.872, 0.212, 0.832, 0.298, 0.791], ['C', 0.323, 0.779, 0.31, 0.755, 0.28, 0.755]];

/** Draw a normalised path array inside the given box (points are box-relative). */
function pathShape (s, data, x, y, w, h, color) {
	const pts = [];
	data.forEach(seg => {
		if (seg[0] === 'M') pts.push({ x: seg[1] * w, y: seg[2] * h, moveTo: true });
		else if (seg[0] === 'L') pts.push({ x: seg[1] * w, y: seg[2] * h });
		else if (seg[0] === 'Z') pts.push({ close: true });
		else pts.push({
			x: seg[5] * w, y: seg[6] * h,
			curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h }
		});
	});
	s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color }, line: NOLINE });
}

function mannequin (s, x, y, w, h, color) {
	pathShape(s, FORM_BODY, x, y + 0.106 * h, w, 0.545 * h, color);
	pathShape(s, FORM_STAND, x + 0.186 * w, y + 0.662 * h, 0.617 * w, 0.338 * h, color);
	s.addShape('ellipse', { x: x + 0.364 * w, y, w: 0.259 * w, h: 0.092 * h, fill: { color }, line: NOLINE });
}

/**
 * Boxed GLITTERATI lockup: a hollow icon box holding the dress form, with a
 * hollow wordmark bar tucked under it. `markY`/`markH` place the dress form,
 * which is always centred horizontally and has a fixed 0.355 aspect ratio.
 */
function logo (s, o, color) {
	const barWall = o.barH * 0.049;
	const barY = o.y + o.iconH - barWall;
	frame(s, o.x, o.y, o.w, o.iconH, o.iconH * 0.021, color);
	frame(s, o.x, barY, o.w, o.barH, barWall, color);
	txt(s, 'GLITTERATI', { x: o.x, y: barY, w: o.w, h: o.barH, fontSize: o.fontSize, color, align: 'center', valign: 'middle' });
	const mw = o.markH * 0.3548;
	mannequin(s, o.x + (o.w - mw) / 2, o.markY, mw, o.markH, color);
}

/** Vertical GLITTERATI tab that runs down a slide edge. */
function sideTab (s, side, fill, ink) {
	const x = side === 'left' ? 0 : 12.794;
	if (fill) rect(s, x, 0, 0.539, 2.813, fill);
	txt(s, 'GLITTERATI', {
		x: x - 0.855, y: 1.469, w: 2.284, h: 0.404, rotate: side === 'left' ? 90 : 270,
		fontSize: 18, charSpacing: 3, color: ink, align: 'center'
	});
	mannequin(s, x + 0.196, 0.14, 0.184, 0.518, ink);
}

/** Flat stand-in for a photograph. */
function photo (s, x, y, w, h) {
	rect(s, x, y, w, h, PLACEHOLDER);
	txt(s, '[image]', { x, y: y + h / 2 - 0.16, w, h: 0.32, fontFace: BODY, fontSize: 11, color: '8C8C8C', align: 'center' });
}

/**
 * Oversized watermark letter (the S/W/O/T of the SWOT section).
 * `cx`,`cy` are the visual centre of the glyph, `capHeight` its cap height.
 */
function letter (s, ch, cx, cy, capHeight, color) {
	txt(s, ch, {
		x: cx - 5, y: cy - 0.072 * capHeight - capHeight, w: 10, h: capHeight * 2,
		fontSize: Math.round(capHeight / 0.7 * 72), color, align: 'center', valign: 'middle'
	});
}

/* body-copy shorthand ---------------------------------------------------- */
function para (s, text, x, y, w, h, color, align) {
	txt(s, text, { x, y, w, h, fontFace: BODY, fontSize: 10, lineSpacing: 16, color, align: align || 'left' });
}
function heading (s, lines, x, y, w, h, size, color, align, spc) {
	txt(s, lines, { x, y, w, h, fontSize: size, charSpacing: spc === undefined ? 3 : spc, color, align: align || 'left' });
}

/* --------------------------------------------------------------- the deck */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'GLITTERATI', width: 13.333, height: 7.5 });
pptx.layout = 'GLITTERATI';
pptx.author = 'GLITTERATI';
pptx.title = 'GLITTERATI';

const slide = () => { const s = pptx.addSlide(); s.background = { color: WHITE }; return s; };

/* 1 — cover ---------------------------------------------------------------- */
function slide01 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 7.5, DARK);
	rect(s, 6.667, 0.491, 6.189, 6.519, DARK, 25);
	logo(s, { x: 4.625, y: 1.89, w: 4.083, iconH: 2.626, barH: 1.013, fontSize: 44, markY: 2.287, markH: 1.914 }, WHITE);
	txt(s, URL, { x: 4.768, y: 6.427, w: 3.798, h: 0.303, fontFace: BODY, fontSize: 12, charSpacing: 6, color: WHITE, align: 'center' });
}

/* 2 — welcome message ------------------------------------------------------ */
function slide02 () {
	const s = slide();
	rect(s, 0.903, 0, 11.527, 6.614, DARK);
	rect(s, 6.0, 0.886, 3.804, 3.583, SILVER);
	heading(s, ['WEL', 'COME', 'MAS', 'SAGE.'], 7.294, 1.228, 2.147, 2.794, 40, DARK);
	para(s, L1, 7.383, 4.754, 4.284, 0.974, WHITE);
	footer(s, 0.746, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 3 — fashion quotes (dark) ------------------------------------------------ */
function slide03 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 7.5, DARK);
	heading(s, ['FASHION', 'QUOTES.'], 7.865, 1.223, 4.095, 1.919, 54, WHITE);
	txt(s, Q_COCO, { x: 6.004, y: 4.446, w: 4.728, h: 1.447, fontFace: SERIF, fontSize: 20, italic: true, color: WHITE });
	txt(s, '\u2013 Coco Chanel', { x: 6.021, y: 5.894, w: 3.208, h: 0.337, fontFace: SERIF, fontSize: 14, italic: true, color: WHITE });
	footer(s, 9.217, 'right', WHITE);
	sideTab(s, 'right', WHITE, DARK);
}

/* 4 — try the best for customers ------------------------------------------- */
function slide04 () {
	const s = slide();
	rect(s, 0, 2.666, 0.421, 2.168, DARK);
	heading(s, ['TRY ', 'THE BEST', 'FOR', 'CUSTOMERS.'], 0.826, 1.272, 5.329, 3.332, 48, DARK, 'center');
	para(s, L1, 1.261, 4.897, 4.46, 0.974, GREY, 'center');
	footer(s, 1.719, 'center');
	sideTab(s, 'right', DARK, WHITE);
}

/* 5 — our fashion service (4 numbered styles) ------------------------------ */
function slide05 () {
	const s = slide();
	heading(s, ['OUR', 'FASHION', 'SERVICE.'], 5.912, 0.874, 3.825, 2.322, 44, DARK, 'right');
	[
		{ n: '01', label: 'EUROPEAN STYLE', bx: 7.59, y: 3.949, lx: 4.895, lw: 2.428, tx: 4.881 },
		{ n: '02', label: 'ASIAN STYLE', bx: 3.65, y: 3.949, lx: 1.229, lw: 2.154, tx: 0.941 },
		{ n: '03', label: 'AMERICAN STYLE', bx: 7.59, y: 5.454, lx: 4.895, lw: 2.428, tx: 4.881 },
		{ n: '04', label: 'ARABIAN STYLE', bx: 3.65, y: 5.454, lx: 1.229, lw: 2.154, tx: 0.941 }
	].forEach(c => {
		s.addShape('roundRect', { x: c.bx, y: c.y, w: 0.611, h: 0.611, rectRadius: 0.102, fill: { color: DARK }, line: NOLINE });
		txt(s, c.n, { x: c.bx, y: c.y, w: 0.611, h: 0.611, fontFace: UI, fontSize: 20, bold: true, color: WHITE, align: 'center', valign: 'middle' });
		txt(s, c.label, { x: c.lx, y: c.y, w: c.lw, h: 0.37, fontFace: UI, fontSize: 16, bold: true, color: GREY2, align: 'right' });
		txt(s, M1, { x: c.tx, y: c.y + 0.371, w: 2.442, h: 0.634, fontFace: BODY, fontSize: 10, lineSpacing: 12.8, color: GREY, align: 'right' });
	});
	footer(s, 0.746, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 6 — new winter edition --------------------------------------------------- */
function slide06 () {
	const s = slide();
	rect(s, 0, 0, 4.175, 5.912, DARK);
	para(s, L1, 4.698, 0.831, 6.144, 0.75, GREY);
	heading(s, ['NEW', 'WINTER', 'EDITION.'], 7.77, 3.461, 4.175, 2.827, 54, DARK);
	footer(s, -0.032, 'left', WHITE, 3.624, { rotate: 270 });
	sideTab(s, 'right', DARK, WHITE);
}

/* 7 — new spring edition --------------------------------------------------- */
function slide07 () {
	const s = slide();
	heading(s, ['NEW', 'SPRING', 'EDITION.'], 1.332, 4.339, 3.914, 2.524, 48, DARK);
	para(s, L1, 5.807, 5.601, 6.144, 0.75, GREY);
	footer(s, 9.614, 'center', GREY, 2.44, { rotate: 270 });
	sideTab(s, 'right', DARK, WHITE);
}

/* 8 — new office edition --------------------------------------------------- */
function slide08 () {
	const s = slide();
	heading(s, ['NEW', 'OFFICE', 'EDITION.'], 0.972, 0.747, 4.475, 2.827, 54, DARK, 'left', 6);
	para(s, L1, 7.672, 3.263, 4.46, 0.974, GREY);
	footer(s, 0.776, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 9 — meet fashion team ---------------------------------------------------- */
function slide09 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 4.228, DARK);
	txt(s, 'MEET FASHION TEAM.', { x: 2.677, y: 0.696, w: 7.979, h: 0.774, fontSize: 40, charSpacing: 3, color: WHITE, align: 'center' });
	[
		{ cx: 1.116, cy: 3.474, nx: 2.674, nw: 1.134, ny: 3.767, tx: 1.328, ty: 4.53, rx: 1.466, ry: 5.473, n: '01.', name: ['GWEN', 'NORRIS'], role: 'Fashion Designer' },
		{ cx: 5.202, cy: 3.832, nx: 6.76, nw: 1.264, ny: 4.142, tx: 5.544, ty: 4.888, rx: 5.682, ry: 5.831, n: '02.', name: ['NIKLAS', 'LOMO'], role: 'Coach Traine Model' },
		{ cx: 9.115, cy: 3.474, nx: 10.665, nw: 1.231, ny: 3.784, tx: 9.426, ty: 4.564, rx: 9.564, ry: 5.507, n: '03.', name: ['ROBERT', 'FRANCIS'], role: 'Coordinator' }
	].forEach(c => {
		rect(s, c.cx, c.cy, 3.1, 2.803, DARK, 20);
		txt(s, c.n, { x: c.nx, y: c.ny, w: c.nw, h: 0.707, fontSize: 36, charSpacing: 6, color: WHITE, align: 'right' });
		txt(s, c.name, { x: c.tx, y: c.ty, w: 2.479, h: 0.909, fontSize: 24, color: WHITE, align: 'right' });
		txt(s, c.role, { x: c.rx, y: c.ry, w: 2.339, h: 0.325, fontFace: BODY, fontSize: 14, lineSpacing: 16, color: WHITE, align: 'right' });
	});
	footer(s, 4.895, 'center');
	sideTab(s, 'right', WHITE, DARK);
}

/* 10 — new arrival (two product cards) ------------------------------------- */
function slide10 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 7.5, DARK);
	rect(s, 2.402, 0, 3.349, 7.5, WHITE);
	heading(s, ['NEW', 'ARRIVAL.'], 7.649, 0.731, 4.681, 2.121, 60, WHITE, 'center');
	para(s, L2, 7.763, 2.851, 4.454, 0.75, WHITE, 'center');
	[
		{ x: 7.023, solid: true, ink: DARK, name: ['BLUE', 'SHIRT'], price: '75$', tx: 7.649, bx: 7.197 },
		{ x: 10.087, solid: false, ink: WHITE, name: ['GREY', 'SHIRT'], price: '70$', tx: 10.714, bx: 10.262 }
	].forEach(c => {
		if (c.solid) rect(s, c.x, 3.998, 2.89, 2.709, SAND);
		else frame(s, c.x, 3.998, 2.89, 2.709, 0.0755, SAND);
		txt(s, c.name, { x: c.tx, y: 4.198, w: 1.637, h: 0.909, fontSize: 24, color: c.ink, align: 'center' });
		txt(s, c.price, { x: c.tx, y: 5.052, w: 1.637, h: 0.64, fontSize: 32, color: c.ink, align: 'center' });
		txt(s, M5, { x: c.bx, y: 5.691, w: 2.541, h: 0.75, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: c.ink, align: 'center' });
	});
	footer(s, -1.046, 'center', WHITE, 3.624, { rotate: 270 });
	sideTab(s, 'right', WHITE, DARK);
}

/* 11 — bag sale ------------------------------------------------------------ */
function slide11 () {
	const s = slide();
	heading(s, ['BAG SALE', 'UP TO 50%', 'FOR 1 WEEK.'], 7.377, 0.8, 5.186, 2.524, 48, DARK);
	para(s, L1, 7.747, 4.422, 4.46, 0.974, GREY);
	[
		{ x: 1.127, price: '47$', pw: 0.759, label: 'NAVY URBAN BAG', lx: 1.362, lw: 2.358 },
		{ x: 4.193, price: '56$', pw: 0.742, label: 'BROWN URBAN BAG', lx: 4.248, lw: 2.713 }
	].forEach(c => {
		rect(s, c.x, 5.262, 0.824, 0.74, GOLD);
		txt(s, c.price, { x: c.x + 0.033, y: 5.446, w: c.pw, h: 0.37, fontSize: 16, charSpacing: 3, color: DARK, align: 'center' });
		frame(s, c.x, 6.139, 2.828, 0.523, 0.0362, DARK);
		txt(s, c.label, { x: c.lx, y: 6.219, w: c.lw, h: 0.37, fontSize: 16, color: GREY2, align: 'center' });
	});
	footer(s, 9.79, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 12 — style of fashion ---------------------------------------------------- */
function slide12 () {
	const s = slide();
	rect(s, 1.825, 0.993, 11.509, 4.743, DARK);
	heading(s, ['STYLE ', 'OF ', 'FASHION.'], 9.048, 1.122, 3.548, 2.121, 40, WHITE, 'right');
	para(s, L1, 2.586, 2.878, 4.284, 0.974, WHITE, 'right');
	footer(s, 0.746, 'left');
	sideTab(s, 'left', DARK, WHITE);
}

/* 13 — fashion is everything ----------------------------------------------- */
function slide13 () {
	const s = slide();
	heading(s, ['FAS', 'HION', 'IS', 'EVERY', 'THING.'], 0.746, 2.973, 2.636, 3.803, 44, DARK, 'right');
	para(s, L1, 2.483, 0.877, 4.284, 0.974, GREY, 'right');
	footer(s, 8.667, 'right');
	sideTab(s, 'right', DARK, WHITE);
}

/* 14 — fashion quotes (light) ---------------------------------------------- */
function slide14 () {
	const s = slide();
	rect(s, 4.968, 2.89, 3.397, 2.725, DARK);
	txt(s, 'FASHION QUOTES.', { x: 3.109, y: 0.886, w: 7.115, h: 0.774, fontSize: 40, charSpacing: 3, color: DARK, align: 'center' });
	txt(s, Q_RALPH, { x: 5.271, y: 3.521, w: 2.792, h: 1.043, fontFace: SERIF, fontSize: 14, italic: true, color: WHITE, align: 'center' });
	txt(s, '\u2013 Ralph Lauren', { x: 5.652, y: 4.609, w: 2.029, h: 0.337, fontFace: SERIF, fontSize: 14, italic: true, color: WHITE, align: 'center' });
	footer(s, 4.895, 'center');
	sideTab(s, 'right', DARK, WHITE);
}

/* 15 — fashion gallery ----------------------------------------------------- */
function slide15 () {
	const s = slide();
	txt(s, 'FASHION.', { x: 0.738, y: 2.485, w: 4.048, h: 0.909, fontSize: 48, charSpacing: 3, color: DARK, align: 'center' });
	txt(s, 'GALLERY.', { x: 8.66, y: 4.137, w: 4.048, h: 0.909, fontSize: 48, charSpacing: 3, color: DARK });
	footer(s, 4.895, 'center');
	sideTab(s, 'right', DARK, WHITE);
}

/* 16 — break slide --------------------------------------------------------- */
function slide16 () {
	const s = slide();
	rect(s, 7.347, 0.413, 5.594, 6.667, DARK);
	bracketFrame(s, 7.985, 2.115, 4.318, 3.262, 0.095, 1.142, WHITE);
	bracketFrame(s, 8.476, 2.617, 3.336, 2.188, 0.17, 0.64, WHITE);
	txt(s, 'BREAK SLIDE', { x: 7.663, y: 3.257, w: 4.962, h: 0.909, fontSize: 48, color: WHITE, align: 'center' });
	txt(s, URL, { x: 8.245, y: 6.534, w: 3.798, h: 0.303, fontFace: BODY, fontSize: 12, charSpacing: 6, color: WHITE, align: 'center' });
	logo(s, { x: 2.046, y: 0.601, w: 3.673, iconH: 2.278, barH: 0.792, fontSize: 36, markY: 0.825, markH: 1.793 }, DARK);
}

/* 17 — SWOT: strength ------------------------------------------------------ */
function slide17 () {
	const s = slide();
	rect(s, 6.473, 1.007, 5.487, 5.487, DARK);
	frame(s, 6.668, 1.201, 5.099, 5.099, 0.111, SAND);
	rect(s, 0, 0, 5.544, 7.5, SAND, 20);
	letter(s, 'S', 2.765, 3.825, 5.77, WHITE);
	txt(s, 'STRENGTH.', { x: 6.892, y: 2.539, w: 4.649, h: 0.909, fontSize: 48, charSpacing: 3, color: WHITE, align: 'center' });
	para(s, L1, 7.386, 3.669, 3.662, 1.199, WHITE, 'center');
	footer(s, 9.217, 'right');
	sideTab(s, 'right', DARK, WHITE);
}

/* 18 — SWOT: weakness ------------------------------------------------------ */
function slide18 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 4.667, SAND, 20);
	letter(s, 'W', 3.56, 2.375, 3.98, WHITE);
	heading(s, 'WEAKNESS.', 6.667, 1.778, 5.819, 1.111, 60, DARK);
	[
		{ x: 1.94, icon: 1.315, ly: 5.195, by: 5.559, label: 'PREPARATION', glyph: 'bank' },
		{ x: 5.842, icon: 5.088, ly: 5.189, by: 5.553, label: 'BRIEFING', glyph: 'inbox' },
		{ x: 9.744, icon: 9.224, ly: 5.189, by: 5.553, label: 'EXECUTION', glyph: 'cup' }
	].forEach(c => {
		glyph(s, c.glyph, c.icon, 5.189, 0.42, DARK);
		txt(s, c.label, { x: c.x, y: c.ly, w: 2.154, h: 0.37, fontSize: 16, bold: true, color: GREY2 });
		txt(s, M1, { x: c.x, y: c.by, w: 2.481, h: 0.75, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: GREY });
	});
	footer(s, 4.895, 'center');
	sideTab(s, 'right', DARK, WHITE);
}

/* 19 — SWOT: opportunity --------------------------------------------------- */
function slide19 () {
	const s = slide();
	rect(s, 1.724, 1.702, 9.886, 5.798, SAND, 20);
	letter(s, 'O', 6.66, 4.635, 4.83, WHITE);
	heading(s, ['OPPOR', 'TUNITY.'], 0.875, 0.541, 4.25, 2.322, 66, DARK);
	para(s, L1, 9.678, 4.601, 2.779, 1.647, GREY2);
	footer(s, 0.05, 'right', GREY2);
	sideTab(s, 'right', DARK, WHITE);
}

/* 20 — SWOT: threats ------------------------------------------------------- */
function slide20 () {
	const s = slide();
	rect(s, 2.0, 3.011, 5.947, 2.812, DARK);
	rect(s, 7.509, 1.333, 5.825, 6.167, SAND, 20);
	letter(s, 'T', 10.405, 4.47, 4.82, WHITE);
	heading(s, 'THREATS.', 5.175, 0.801, 5.819, 1.212, 66, DARK);
	txt(s, Q_RALPH, { x: 2.537, y: 3.563, w: 4.177, h: 1.447, fontFace: SERIF, fontSize: 20, italic: true, color: WHITE, align: 'right' });
	txt(s, '\u2013 Ralph Lauren', { x: 4.148, y: 5.01, w: 2.532, h: 0.337, fontFace: SERIF, fontSize: 14, italic: true, color: WHITE, align: 'right' });
	footer(s, 0.05, 'right');
	sideTab(s, 'left', DARK, WHITE);
}

/* 21 — new arrival women items --------------------------------------------- */
function slide21 () {
	const s = slide();
	heading(s, ['NEW ARRIVAL', 'WOMEN ITEMS.'], 3.871, 0.548, 6.729, 1.717, 48, DARK, 'left', 6);
	para(s, L1, 0.895, 4.094, 2.619, 1.647, GREY, 'right');
	[
		{ x: 11.632, y: 5.478, tx: 11.782, ty: 5.561, px: 11.765, py: 6.111, name: ['GREY', 'SLAVE'] },
		{ x: 6.216, y: 2.336, tx: 6.366, ty: 2.42, px: 6.349, py: 2.969, name: ['BLACK', 'STRIP'] },
		{ x: 8.708, y: 2.331, tx: 8.859, ty: 2.414, px: 8.842, py: 2.964, name: ['BLUE', 'STRIP'] }
	].forEach(c => {
		rect(s, c.x, c.y, 1.271, 1.276, SAND);
		txt(s, c.name, { x: c.tx, y: c.ty, w: 1.004, h: 0.572, fontSize: 14, color: DARK, align: 'center' });
		txt(s, '60$', { x: c.px, y: c.py, w: 1.004, h: 0.505, fontSize: 24, color: DARK, align: 'center' });
	});
	footer(s, 0.539, 'left');
	sideTab(s, 'left', DARK, WHITE);
}

/* 22 — only 4 items up to 70% ---------------------------------------------- */
function slide22 () {
	const s = slide();
	rect(s, 4.625, 0.417, 2.042, 6.667, DARK);
	txt(s, ['ONLY 4 ITEM', 'UP TO 70%'], { x: 3.063, y: 2.959, w: 5.166, h: 1.582, rotate: 270, fontSize: 44, charSpacing: 3, color: WHITE, align: 'center' });
	[
		{ x: 7.32, y: 1.099, fill: SAND, ink: DARK, name: 'JACKET', price: '85$' },
		{ x: 9.871, y: 1.099, fill: DARK, ink: WHITE, name: 'T-SHIRT', price: '60$', border: SAND },
		{ x: 7.32, y: 3.75, fill: SAND, ink: DARK, name: 'BELT', price: '35$', border: DARK },
		{ x: 9.871, y: 3.75, fill: DARK, ink: WHITE, name: 'JEANS', price: '110$' }
	].forEach(c => {
		rect(s, c.x, c.y, 2.378, 2.458, c.fill);
		txt(s, c.name, { x: c.x + 0.373, y: c.y + 0.193, w: 1.637, h: 0.505, fontSize: 24, color: c.ink, align: 'center' });
		txt(s, c.price, { x: c.x + 0.373, y: c.y + 0.743, w: 1.637, h: 0.64, fontSize: 32, color: c.ink, align: 'center' });
		txt(s, M4, { x: c.x + 0.206, y: c.y + 1.428, w: 1.971, h: 0.75, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: c.ink, align: 'center' });
		if (c.border) frame(s, c.x, c.y, 2.378, 2.458, 0.0663, c.border);
	});
	footer(s, 9.331, 'right', GREY);
	sideTab(s, 'right', DARK, WHITE);
}

/* 23 — current trends ------------------------------------------------------ */
function slide23 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 7.5, DARK);
	rect(s, 0, 2.362, 13.333, 2.775, WHITE);
	heading(s, ['CURRENT', 'TRENDS'], 4.107, 0.647, 5.12, 2.121, 60, WHITE, 'center');
	txt(s, Q_KORS, { x: 0.539, y: 2.982, w: 3.237, h: 1.111, fontFace: SERIF, fontSize: 20, italic: true, color: GREY2, align: 'right' });
	txt(s, '\u2013 Michael Kors', { x: 1.245, y: 4.054, w: 2.532, h: 0.337, fontFace: SERIF, fontSize: 14, italic: true, color: GREY2, align: 'right' });
	para(s, L3, 9.557, 3.263, 2.992, 0.974, GREY);
	footer(s, 4.895, 'center', WHITE);
	sideTab(s, 'right', WHITE, DARK);
}

/* 24 — monochrome style ---------------------------------------------------- */
function slide24 () {
	const s = slide();
	rect(s, 5.553, 3.293, 3.804, 3.164, DARK);
	heading(s, ['MONOCHROME', 'STYLE.'], 6.947, 1.086, 5.719, 1.582, 44, DARK, 'right');
	para(s, L1, 6.168, 3.939, 2.575, 1.872, WHITE, 'right');
	footer(s, 9.123, 'right');
	sideTab(s, 'left', DARK, WHITE);
}

/* 25 — the fashion bloggers (progress bars) -------------------------------- */
function slide25 () {
	const s = slide();
	rect(s, 0, 3.75, 13.333, 3.75, DARK);
	heading(s, ['THE', 'FASHION', 'BLOGGERS.'], 1.321, 0.982, 5.004, 2.524, 48, DARK, 'left', 6);
	para(s, L1, 1.215, 3.932, 4.502, 0.974, WHITE);
	[
		{ y: 5.662, label: ['Hunting ', 'Desain'], ly: 5.181, pct: '85%', py: 5.517, fill: 2.429, color: GOLD },
		{ y: 6.323, label: ['Search Uniq ', 'Desain'], ly: 5.862, pct: '80%', py: 6.198, fill: 2.069, color: AMBER }
	].forEach(b => {
		s.addShape('line', { x: 1.47, y: b.y, w: 3.043, h: 0, line: { color: GREY, width: 18 } });
		s.addShape('line', { x: 1.47, y: b.y, w: b.fill, h: 0, line: { color: b.color, width: 18 } });
		txt(s, b.label.join(''), { x: 1.275, y: b.ly, w: 2.431, h: 0.337, fontFace: BODY, fontSize: 12, color: WHITE });
		txt(s, b.pct, { x: 4.665, y: b.py, w: 1.521, h: 0.32, fontFace: BODY, fontSize: 11, color: WHITE });
	});
	footer(s, 9.217, 'right', WHITE);
	sideTab(s, 'left', DARK, WHITE);
}

/* 26 — our meet team (two profiles with skill bars) ------------------------ */
/* label row, bar row, % row — offsets measured from the block's top-left */
const SKILLS = [
	{ label: 'European Style', pct: '100%', len: 2.845, color: GOLD, ly: 0.0, by: 0.372, py: 0.169, px: 2.972 },
	{ label: 'Latin Style', pct: '80%', len: 2.368, color: AMBER, ly: 0.419, by: 0.792, py: 0.613, px: 2.495 },
	{ label: 'Asian Style', pct: '60%', len: 1.733, color: AMBER2, ly: 0.838, by: 1.210, py: 1.036, px: 1.858 },
	{ label: 'American Style', pct: '85%', len: 2.845, color: AMBER3, ly: 1.266, by: 1.611, py: 1.436, px: 2.972 }
];

function skillBars (s, x, y) {
	SKILLS.forEach(k => {
		s.addShape('line', { x: x + 0.126, y: y + k.by, w: k.len, h: 0, line: { color: k.color, width: 5 } });
		txt(s, k.label, { x, y: y + k.ly, w: 3.863, h: 0.32, fontFace: UI, fontSize: 11, color: GREY2 });
		txt(s, k.pct, { x: x + k.px, y: y + k.py, w: 1.272, h: 0.32, fontFace: UI, fontSize: 11, color: GREY2 });
	});
}

function slide26 () {
	const s = slide();
	[
		{ bar: 1.802, tag: 0.807, num: '01.', nx: 0.807, name: 'FRANCISCA', tx: 1.917, tex: 0.807 },
		{ bar: 9.507, tag: 8.512, num: '02.', nx: 8.532, name: 'ANASTACIA', tx: 9.664, tex: 8.512 }
	].forEach(c => {
		rect(s, c.bar, 3.372, 3.019, 0.866, DARK);
		rect(s, c.tag, 3.372, 0.995, 0.866, SAND);
		txt(s, c.num, { x: c.nx, y: 3.519, w: 1.06, h: 0.572, fontSize: 28, charSpacing: 6, color: GREY2, align: 'center' });
		txt(s, c.name, { x: c.tx, y: 3.553, w: 2.479, h: 0.505, fontSize: 24, color: WHITE });
		txt(s, M3, { x: c.tex, y: 4.458, w: 3.766, h: 0.526, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: GREY });
		skillBars(s, c.tex, 4.984);
	});
	frame(s, 5.224, 1.017, 2.89, 2.961, 0.0805, DARK);
	heading(s, ['OUR', 'MEET', 'TEAM.'], 5.337, 1.235, 2.66, 2.524, 48, DARK, 'center');
	footer(s, 4.895, 'center');
	sideTab(s, 'right', DARK, WHITE);
}

/* 27 — fashion for life ---------------------------------------------------- */
function slide27 () {
	const s = slide();
	rect(s, 5.729, 1.257, 3.804, 3.164, DARK);
	heading(s, ['FASHION', 'FOR', 'LIFE.'], 6.053, 1.779, 3.156, 2.121, 40, WHITE);
	para(s, L1, 7.067, 5.028, 4.46, 0.974, GREY);
	footer(s, 0.746, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 28 — mockup layout ------------------------------------------------------- */
function slide28 () {
	const s = slide();
	rect(s, 0, 1.467, 13.333, 0.05, DARK);
	photo(s, 7.846, 0.682, 5.487, 6.818);
	heading(s, ['MOCKUP', 'LAYOUT.'], 1.019, 0.625, 4.156, 1.717, 48, DARK, 'left', 6);
	[
		{ card: 0.804, cy: 3.161, fx: 1.731, fy: 3.503, gx: 1.966, gy: 3.715, label: 'COACING', lx: 1.086, ly: 4.577, bx: 0.97, by: 4.975, glyph: 'walk' },
		{ card: 3.649, cy: 2.629, fx: 4.569, fy: 2.983, gx: 4.826, gy: 3.179, label: 'TRAINING', lx: 3.93, ly: 4.045, bx: 3.815, by: 4.443, glyph: 'run' },
		{ card: 6.494, cy: 3.166, fx: 7.414, fy: 3.503, gx: 7.706, gy: 3.697, label: 'PREPARATION', lx: 6.775, ly: 4.582, bx: 6.659, by: 4.98, glyph: 'stand' }
	].forEach(c => {
		rect(s, c.card, c.cy, 2.723, 3.128, SILVER);
		frame(s, c.fx, c.fy, 0.883, 0.883, 0.0753, DARK);
		glyph(s, c.glyph, c.gx, c.gy, 0.44, DARK);
		txt(s, c.label, { x: c.lx, y: c.ly, w: 2.154, h: 0.37, fontSize: 16, bold: true, color: GREY2, align: 'center' });
		txt(s, M2, { x: c.bx, y: c.by, w: 2.386, h: 0.974, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: GREY2, align: 'center' });
	});
	footer(s, 0.539, 'left');
	sideTab(s, 'right', DARK, WHITE);
}

/* 29 — summer / winter sale ------------------------------------------------ */
function slide29 () {
	const s = slide();
	rect(s, 0, 0, 13.333, 7.5, DARK);
	[
		{ x: 1.394, bx: 1.762, lines: ['SUMMER', 'SALE', 'UP TO 60%'] },
		{ x: 8.224, bx: 8.592, lines: ['WINTER', 'SALE', 'UP TO 70%'] }
	].forEach(c => {
		heading(s, c.lines, c.x, 2.26, 3.734, 2.121, 40, WHITE, 'center');
		txt(s, M2, { x: c.bx, y: 4.381, w: 2.999, h: 0.75, fontFace: BODY, fontSize: 10, lineSpacing: 16, color: WHITE, align: 'center' });
	});
	footer(s, 9.373, 'right');
	sideTab(s, 'left', DARK, WHITE);
}

/* 30 — thank you ----------------------------------------------------------- */
function slide30 () {
	const s = slide();
	rect(s, 1.596, 1.875, 5.07, 4.23, SAND);
	s.addShape('rect', { x: 1.852, y: 2.089, w: 4.558, h: 3.802, fill: { type: 'none' }, line: { color: WHITE, width: 11.5 } });
	rect(s, 6.667, 1.875, 5.07, 4.23, DARK);
	heading(s, ['THANK', 'YOU'], 2.006, 2.728, 4.25, 2.524, 72, DARK, 'center');
	logo(s, { x: 7.365, y: 2.525, w: 3.673, iconH: 2.278, barH: 0.792, fontSize: 36, markY: 2.749, markH: 1.793 }, WHITE);
	txt(s, URL, { x: 10.489, y: 3.839, w: 3.798, h: 0.303, rotate: 90, fontFace: BODY, fontSize: 12, charSpacing: 6, color: GREY2, align: 'center' });
}

/* ------------------------------------------------- 31-42: icon library grid */
/** Tiny pictograms drawn from primitives; used on the icon-sheet slides. */
function glyph (s, kind, x, y, u, color) {
	const R = (a, b, c, d) => rect(s, x + a * u, y + b * u, c * u, d * u, color);
	const E = (a, b, c, d) => s.addShape('ellipse', { x: x + a * u, y: y + b * u, w: c * u, h: d * u, fill: { color }, line: NOLINE });
	const T = (a, b, c, d, rot) => s.addShape('triangle', { x: x + a * u, y: y + b * u, w: c * u, h: d * u, rotate: rot || 0, fill: { color }, line: NOLINE });
	const F = (a, b, c, d, t) => frame(s, x + a * u, y + b * u, c * u, d * u, t * u, color);
	const RING = (a, b, c) => s.addShape('donut', { x: x + a * u, y: y + b * u, w: c * u, h: c * u, fill: { color }, line: NOLINE });
	// knockouts punched out of a solid glyph; the icon sheets always sit on white
	const KOR = (a, b, c, d) => rect(s, x + a * u, y + b * u, c * u, d * u, WHITE);
	const KOT = (a, b, c, d, rot) => s.addShape('triangle', { x: x + a * u, y: y + b * u, w: c * u, h: d * u, rotate: rot, fill: { color: WHITE }, line: NOLINE });
	switch (kind) {
		case 'bank': T(0.02, 0, 0.96, 0.26); R(0, 0.28, 1, 0.09); R(0.12, 0.44, 0.13, 0.4); R(0.44, 0.44, 0.13, 0.4); R(0.75, 0.44, 0.13, 0.4); R(0, 0.89, 1, 0.11); break;
		case 'inbox':
			s.addShape('roundRect', { x, y, w: u, h: u, rectRadius: 0.16 * u, fill: { color }, line: NOLINE });
			KOT(0.26, 0.42, 0.48, 0.28, 180); KOR(0.42, 0.2, 0.16, 0.24); KOR(0.24, 0.76, 0.52, 0.09);
			break;
		case 'cup': R(0.02, 0.14, 0.62, 0.6); F(0.56, 0.2, 0.4, 0.32, 0.09); R(0, 0.84, 0.7, 0.1); break;
		case 'walk': E(0.36, 0, 0.26, 0.26); R(0.4, 0.3, 0.2, 0.36); R(0.16, 0.34, 0.24, 0.08); R(0.6, 0.34, 0.24, 0.08); R(0.32, 0.66, 0.1, 0.34); R(0.58, 0.66, 0.1, 0.34); break;
		case 'run': E(0.5, 0, 0.24, 0.24); R(0.34, 0.28, 0.24, 0.32); R(0.58, 0.3, 0.3, 0.09); R(0.1, 0.5, 0.26, 0.09); R(0.26, 0.62, 0.12, 0.38); R(0.56, 0.6, 0.12, 0.4); break;
		case 'stand': E(0.38, 0, 0.24, 0.24); R(0.4, 0.28, 0.2, 0.38); R(0.18, 0.32, 0.22, 0.08); R(0.6, 0.32, 0.22, 0.08); R(0.36, 0.68, 0.1, 0.32); R(0.54, 0.68, 0.1, 0.32); break;
		case 'clock': RING(0, 0, 1); R(0.46, 0.2, 0.08, 0.34); R(0.46, 0.46, 0.28, 0.08); break;
		case 'photo': F(0, 0.14, 1, 0.72, 0.1); E(0.36, 0.36, 0.28, 0.28); break;
		case 'roundsq': s.addShape('roundRect', { x, y: y + 0.05 * u, w: u, h: 0.9 * u, rectRadius: 0.2 * u, fill: { type: 'none' }, line: { color, width: 2.4 } }); break;
		case 'star': s.addShape('star5', { x, y, w: u, h: u, fill: { color }, line: NOLINE }); break;
		case 'chat': F(0, 0.06, 1, 0.62, 0.1); T(0.14, 0.6, 0.26, 0.32, 180); break;
		case 'doc': F(0.14, 0, 0.72, 1, 0.09); [0.24, 0.44, 0.64].forEach(o => R(0.28, o, 0.44, 0.08)); break;
		case 'list': [0, 0.26, 0.52, 0.78].forEach(o => { R(0, o, 0.16, 0.16); R(0.26, o + 0.04, 0.74, 0.09); }); break;
		case 'pin': RING(0.15, 0, 0.7); T(0.3, 0.52, 0.4, 0.48, 180); break;
		case 'bar': R(0.04, 0.56, 0.22, 0.44); R(0.39, 0.3, 0.22, 0.7); R(0.74, 0.04, 0.22, 0.96); break;
		case 'arrow': s.addShape('rightArrow', { x, y: y + 0.22 * u, w: u, h: 0.56 * u, fill: { color }, line: NOLINE }); break;
		case 'plus': s.addShape('plus', { x, y, w: u, h: u, fill: { color }, line: NOLINE }); break;
		case 'ring': RING(0, 0, 1); break;
		case 'tri': T(0, 0.06, 1, 0.88); break;
		case 'heart': s.addShape('heart', { x, y, w: u, h: u, fill: { color }, line: NOLINE }); break;
		case 'cloud': s.addShape('cloud', { x, y: y + 0.16 * u, w: u, h: 0.68 * u, fill: { color }, line: NOLINE }); break;
		case 'square': F(0, 0.05, 1, 0.9, 0.11); break;
		case 'dots': [0, 0.36, 0.72].forEach(r => [0, 0.36, 0.72].forEach(c => R(c, r + 0.14, 0.2, 0.2))); break;
		case 'battery': F(0.16, 0.08, 0.68, 0.92, 0.1); R(0.38, 0, 0.24, 0.08); break;
		case 'mail': F(0, 0.16, 1, 0.68, 0.09); T(0.12, 0.26, 0.76, 0.4, 180); break;
		case 'grid': [0, 0.55].forEach(r => [0, 0.55].forEach(c => R(c, r + 0.05, 0.45, 0.4))); break;
		case 'pie': s.addShape('pie', { x, y, w: u, h: u, angleRange: [0, 270], fill: { color }, line: NOLINE }); break;
		case 'phone': F(0.24, 0, 0.52, 1, 0.09); R(0.42, 0.86, 0.16, 0.06); break;
		case 'check':
			s.addShape('rect', { x: x + 0.02 * u, y: y + 0.5 * u, w: 0.36 * u, h: 0.14 * u, rotate: 45, fill: { color }, line: NOLINE });
			s.addShape('rect', { x: x + 0.3 * u, y: y + 0.3 * u, w: 0.72 * u, h: 0.14 * u, rotate: -45, fill: { color }, line: NOLINE });
			break;
		case 'lock': F(0.24, 0, 0.52, 0.5, 0.1); R(0.06, 0.42, 0.88, 0.58); break;
		default: F(0, 0, 1, 1, 0.1);
	}
}

const GLYPHS = ['walk', 'bank', 'doc', 'roundsq', 'stand', 'clock', 'square', 'ring', 'pin', 'star',
	'chat', 'plus', 'bar', 'arrow', 'cup', 'tri', 'inbox', 'heart', 'cloud', 'run',
	'list', 'photo', 'dots', 'battery', 'mail', 'grid', 'pie', 'phone', 'check', 'lock'];

/**
 * Icon-sheet slide: a 10-column grid of small dark pictograms.
 * Every sheet has 7 rows except slide 39, which packs 8; slide 42 is 3 short.
 */
function iconSheet (index, rows, missing) {
	const s = slide();
	const x0 = 1.36, dx = 1.178, size = 0.42;
	const y0 = 0.9, dy = rows === 8 ? 0.814 : 0.95;
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < 10; c++) {
			if (r === rows - 1 && missing && missing.indexOf(c) >= 0) continue;
			const seed = (index * 977 + r * 131 + c * 31) * 2654435761 % 1000003;
			const kind = GLYPHS[seed % GLYPHS.length];
			glyph(s, kind, x0 + c * dx - size / 2, y0 + r * dy - size / 2, size, DARK);
		}
	}
}

/* ------------------------------------------------------------------- build */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
	slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
	slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30
].forEach(fn => fn());

for (let i = 0; i < 12; i++) {
	const n = 31 + i;
	iconSheet(i, n === 39 ? 8 : 7, n === 42 ? [5, 6, 7] : null);
}

pptx.writeFile({ fileName: path.join(__dirname, '1726e86d-cda7-49a5-ab74-12f6ec33acb9_grok_final.pptx') })
	.then(f => console.log('wrote ' + f));
