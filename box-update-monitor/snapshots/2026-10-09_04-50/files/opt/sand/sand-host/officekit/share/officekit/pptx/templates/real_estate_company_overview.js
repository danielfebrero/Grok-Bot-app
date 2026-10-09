/**
 * ESTETIKA — Real Estate Presentation (24 slides, 13.333in x 7.5in)
 * Rebuilt with pptxgenjs only. Run: node <this file>
 */
'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SLIDE_W = 13.3333333; // 12192000 EMU
const SLIDE_H = 7.5;

const GREEN = '57B23E'; // theme accent1
const MINT  = 'DDF0D8'; // accent1 lightened 80%
const DARK  = '262626'; // text 85% black
const GRAY  = '808080'; // text 50% black
const WHITE = 'FFFFFF';

const HEAD = 'Poppins SemiBold'; // theme major font
const BODY = 'Open Sans';        // theme minor font

const SZ_HERO = 66;  // cover wordmark
const SZ_BIG  = 60;  // break / thank-you wordmark
const SZ_HEAD = 28;  // section headings
const SZ_NUM  = 24;  // numbers inside cards
const SZ_LBL  = 14;  // small caps labels & page numbers
const SZ_BODY = 11;  // paragraph copy

const LINE = 1.5;    // 150% paragraph line spacing everywhere

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** Blend two hex colours; t=0 -> a, t=1 -> b. */
function mix(a, b, t) {
	let out = '';
	for (let i = 0; i < 3; i++) {
		const ca = parseInt(a.substr(i * 2, 2), 16);
		const cb = parseInt(b.substr(i * 2, 2), 16);
		out += Math.round(ca + (cb - ca) * t).toString(16).padStart(2, '0');
	}
	return out.toUpperCase();
}

/** Flat colour rectangle. */
function panel(slide, x, y, w, h, color) {
	slide.addShape('rect', { x, y, w, h, fill: { color } });
}

/** Heading built from coloured runs: [['TEXT ', DARK], ['MORE', GREEN]]. */
function heading(slide, x, y, w, h, runs, opts) {
	opts = opts || {};
	slide.addText(runs.map(r => ({ text: r[0], options: { color: r[1] } })), {
		x, y, w, h,
		fontFace: HEAD, fontSize: opts.size || SZ_HEAD,
		align: opts.align || 'left', valign: 'top',
		wrap: opts.wrap !== false,
	});
}

/** Paragraph copy: 11pt Open Sans, 150% leading. */
function body(slide, x, y, w, h, text, opts) {
	opts = opts || {};
	slide.addText(text, {
		x, y, w, h,
		fontFace: BODY, fontSize: SZ_BODY, color: opts.color || GRAY,
		align: opts.align || 'left', valign: 'top',
		lineSpacingMultiple: LINE, wrap: true,
	});
}

/** Small caps label above a paragraph. */
function label(slide, x, y, w, text, color) {
	slide.addText(text, {
		x, y, w, h: 0.337,
		fontFace: HEAD, fontSize: SZ_LBL, color: color || DARK,
		valign: 'top', wrap: false,
	});
}

/** Rounded-square bullet marker. */
function dot(slide, x, y, size, color) {
	slide.addShape('roundRect', {
		x, y, w: size, h: size,
		fill: { color }, rectRadius: size * 0.2,
	});
}

/** Rounded card holding a big number / price. */
function numberCard(slide, x, y, w, h, cardColor, text, textColor) {
	slide.addShape('roundRect', { x, y, w, h, fill: { color: cardColor }, rectRadius: 0.09 });
	slide.addText(text, {
		x, y: y + (h - 0.505) / 2, w, h: 0.505,
		fontFace: HEAD, fontSize: SZ_NUM, color: textColor,
		align: 'center', valign: 'top', wrap: false,
	});
}

/**
 * "ESTETIKA" logo badge — a green pill with a house-in-circle mark.
 * side 'left'  -> pill hugs the left slide edge, wordmark left, mark right
 * side 'right' -> pill hugs the right slide edge, mark left, wordmark right
 */
function badge(slide, side) {
	const W = 1.361, Y = 0.409, H = 0.383;
	const gx = side === 'left' ? 0 : SLIDE_W - W;
	// the pill runs 0.25in past the slide edge so only the rounded end shows
	slide.addShape('roundRect', {
		x: side === 'left' ? gx - 0.25 : gx, y: Y, w: W + 0.25, h: H,
		fill: { color: GREEN }, rectRadius: H / 2,
	});
	// child offsets measured from the pill's inner (on-slide) origin
	const at = (rx, rw) => gx + (side === 'left' ? W - rx - rw : rx);
	slide.addText('ESTETIKA', {
		x: at(0.412, 0.894), y: 0.457, w: 0.894, h: 0.286,
		fontFace: HEAD, fontSize: 11, color: WHITE,
		align: 'center', valign: 'top', wrap: false,
	});
	slide.addShape('ellipse',  { x: at(0.057, 0.270), y: 0.465, w: 0.270, h: 0.270, fill: { color: WHITE } });
	slide.addShape('triangle', { x: at(0.094, 0.196), y: 0.524, w: 0.196, h: 0.082, fill: { color: GREEN } });
	slide.addShape('rect',     { x: at(0.131, 0.122), y: 0.606, w: 0.122, h: 0.068, fill: { color: GREEN } });
	slide.addShape('rect',     { x: at(0.176, 0.030), y: 0.625, w: 0.030, h: 0.049, fill: { color: WHITE } });
}

/**
 * Pale mint trapezoid tabs bleeding off the slide edges.
 * kind 'v' -> top + bottom, 'h' -> left + right, 'l' -> left only
 */
function edgeTabs(slide, kind) {
	const tab = (x, y, w, h, rotate) =>
		slide.addShape('flowChartManualOperation', { x, y, w, h, rotate, fill: { color: MINT } });
	if (kind === 'v') {
		tab(4.563, 0.003, 4.206, 0.452, 0);
		tab(4.563, 7.048, 4.206, 0.452, 180);
	} else {
		tab(-1.173, 3.524, 2.798, 0.452, 270);
		if (kind === 'h') tab(11.709, 3.524, 2.798, 0.452, 90);
	}
}

/** Bottom-right page number ("01." … "23."). */
function pageNumber(slide, text, color) {
	slide.addText(text, {
		x: 12.382, y: 6.706, w: 0.5, h: 0.337,
		fontFace: HEAD, fontSize: SZ_LBL, color: color || DARK,
		align: 'right', valign: 'top', wrap: false,
	});
}

/**
 * Cover / closing wedge: a tall trapezoid whose wide edge is pale green and
 * whose narrow edge is solid green. Painted as vertical slices because
 * pptxgenjs cannot emit a gradient fill.
 * dir 'ltr' -> wide/pale on the left (slide 1)
 * dir 'rtl' -> wide/pale on the right (slide 24)
 */
function gradientWedge(slide, xLeft, dir) {
	const W = 2.342, H = 5.433, TOP = 1.0335, INSET = 0.2155 * H;
	const PALE = mix(WHITE, GREEN, 0.30);
	const N = 60;
	for (let i = 0; i < N; i++) {
		const u0 = i / N, u1 = (i + 1) / N;
		// t: 0 at the wide/pale end, 1 at the narrow/solid end
		const t = dir === 'ltr' ? (u0 + u1) / 2 : 1 - (u0 + u1) / 2;
		const [c0, c1] = dir === 'ltr' ? [u0, u1] : [1 - u0, 1 - u1];
		const sw = W / N + 0.004; // hairline overlap hides seams between slices
		slide.addShape('custGeom', {
			x: xLeft + u0 * W, y: TOP, w: sw, h: H,
			fill: { color: mix(PALE, GREEN, t) },
			points: [
				{ x: 0,  y: c0 * INSET },
				{ x: sw, y: c1 * INSET },
				{ x: sw, y: H - c1 * INSET },
				{ x: 0,  y: H - c0 * INSET },
				{ close: true },
			],
		});
	}
}

/* ------------------------------------------------------------------ *
 * Reusable copy
 * ------------------------------------------------------------------ */
const LOREM_XL = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
	'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
	'ullamco laboris nisi ut aliquip excommodo consequat. Duis aute irure dolor in reprehenderit in ' +
	'voluptate velit esse cillum';
const LOREM_WIDE = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
	'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
	'ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const LOREM_BULLET = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor ' +
	'incididunt ut labore et dolore magna aliqua. ';
const LOREM_SHORT = 'Lorem ipsum dolor amet, consectetur adipiscing';
const LOREM_SERVICE = 'Lorem ipsum dolor sit amet, consectetur adipiscing sediusmod tempor';
const LOREM_TILE = 'Lorem ipsum dolor consectetu adipiscing elit, sed eiusmod tempor incididunt ut labore et';
const LOREM_COVER = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
	'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris';

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 01 — Cover
function slide01(s) {
	edgeTabs(s, 'v');
	gradientWedge(s, 3.2765, 'ltr');
	badge(s, 'right');
	heading(s, 6.848, 2.295, 4.355, 1.212,
		[['ESTET', DARK], ['IKA', GREEN]], { size: SZ_HERO, wrap: false });
	s.addText('Real Estate Presentation', {
		x: 6.889, y: 3.338, w: 3.435, h: 0.337,
		fontFace: BODY, fontSize: SZ_LBL, color: GRAY,
		charSpacing: 3, valign: 'top', wrap: false,
	});
	body(s, 6.889, 4.114, 5.317, 0.904, LOREM_COVER);
	pageNumber(s, '01.');
}

// 02 — Welcome
function slide02(s) {
	panel(s, 1.851, 3.938, 9.631, 0.19, GREEN);
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 4.345, 4.897, 4.642, 0.572,
		[['WELCOME TO ', DARK], ['ESTETIKA', GREEN]], { align: 'center', wrap: false });
	body(s, 1.47, 5.659, 10.394, 0.904,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris ' +
		'nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit ' +
		'esse cillum dolore eu fugiat nulla pariatur. ', { align: 'center' });
	pageNumber(s, '02.');
}

// 03 — Introduction (green side panel right)
function slide03(s) {
	panel(s, 10.043, 0, 3.29, 7.5, GREEN);
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.199, 1.293, 4.773, 1.043,
		[['INTRODUCTION TO THE ', DARK], ['PRESENTATION', GREEN]]);
	body(s, 1.199, 2.589, 5.908, 1.182, LOREM_XL);
	[4.147, 5.247].forEach(y => {
		dot(s, 1.312, y + 0.337, 0.222, GREEN);
		label(s, 2.016, y, 1.177, 'ABOUT US', DARK);
		body(s, 2.016, y + 0.337, 5.092, 0.627, LOREM_BULLET);
	});
	pageNumber(s, '03.', WHITE);
}

// 04 — Our journey
function slide04(s) {
	panel(s, 8.555, 5.149, 2.818, 1.397, GREEN);
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 3.739, 0.869, 5.855, 0.572,
		[['OUR REAL ESTATE ', DARK], ['JOURNEY', GREEN]], { align: 'center' });
	body(s, 1.05, 5.7, 6.667, 0.904,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim minim veniam quis nostrud exercitation ullamco laboris ' +
		'nisi ut aliquip ex ea commodo consequat. ');
	label(s, 9.418, 5.365, 1.152, 'ABOUT US', WHITE);
	body(s, 8.807, 5.702, 2.311, 0.627,
		'Lorem ipsum dolor sit amet, consectetur adipiscing', { color: WHITE, align: 'center' });
	pageNumber(s, '04.');
}

// 05 — Evolution (mint hexagon left, green band)
function slide05(s) {
	s.addShape('custGeom', {
		x: 0, y: 0, w: 1.595, h: 7.5, fill: { color: MINT },
		points: [
			{ x: 0,      y: 0 },
			{ x: 0.0721, y: 0 },
			{ x: 1.595,  y: 0.833 },
			{ x: 1.595,  y: 6.671 },
			{ x: 0.0805, y: 7.5 },
			{ x: 0,      y: 7.5 },
			{ close: true },
		],
	});
	panel(s, 3.959, 4.669, 8.107, 1.787, GREEN);
	edgeTabs(s, 'v');
	badge(s, 'right');
	heading(s, 6.278, 1.293, 5.694, 1.043,
		[['OUR EVOLUTION IN THE REAL ', DARK], ['ESTATE MARKET', GREEN]]);
	body(s, 6.278, 2.589, 5.908, 1.182, LOREM_XL);
	dot(s, 5.711, 5.423, 0.222, WHITE);
	label(s, 6.416, 5.087, 1.177, 'ABOUT US', WHITE);
	body(s, 6.416, 5.423, 5.092, 0.627, LOREM_BULLET, { color: WHITE });
	pageNumber(s, '05.');
}

// 06 — Milestones (wide green band, two ticks)
function slide06(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 7.955, 1.032, 3.785, 1.043,
		[['MILESTONES AND ', DARK], ['SUCCESSES', GREEN]]);
	body(s, 7.955, 2.224, 4.165, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
		'laboris nisi ut aliquip');
	panel(s, 1.367, 4.215, 10.602, 2.125, GREEN);
	[[4.809, 4.606], [5.529, 5.326]].forEach(([yDot, yTxt]) => {
		dot(s, 2.034, yDot, 0.222, WHITE);
		body(s, 2.605, yTxt, 4.909, 0.627,
			'Lorem ipsum dolor amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut',
			{ color: WHITE });
	});
	pageNumber(s, '06.');
}

// 07 — Committed to your goals (green card right)
function slide07(s) {
	panel(s, 5.698, 1.158, 6.194, 3.286, GREEN);
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.361, 1.239, 3.14, 1.515,
		[['COMMITTED TO ', DARK], ['YOUR REAL ', GREEN], ['ESTATE GOALS', DARK]]);
	[[1.829, 1.598], [2.719, 2.487], [3.505, 3.275]].forEach(([yDot, yTxt]) => {
		dot(s, 8.989, yDot, 0.165, WHITE);
		body(s, 9.389, yTxt, 2.033, 0.627, LOREM_SHORT, { color: WHITE });
	});
	body(s, 7.603, 5.149, 4.405, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
		'labore edolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco laboris ' +
		'nisi ut aliquip commodo consequat. ');
	pageNumber(s, '07.');
}

// 08 — Finding your dream home
function slide08(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 8.273, 1.293, 3.378, 1.043,
		[['FINDING YOUR ', DARK], ['DREAM HOME', GREEN]]);
	body(s, 8.273, 2.589, 3.87, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore dolore magna aliqua. Ut enim ad minim venim quis nostrud exercitation ullamco laboris nisi');
	[4.147, 5.247].forEach(y => {
		dot(s, 8.385, y + 0.337, 0.222, GREEN);
		label(s, 9.01, y, 1.177, 'ABOUT US', DARK);
		body(s, 9.01, y + 0.337, 3.132, 0.627, LOREM_SERVICE);
	});
	pageNumber(s, '08.');
}

// 09 — Services (numbered cards on green panel)
function slide09(s) {
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.361, 1.213, 2.902, 1.043,
		[['REAL ESTATE ', DARK], ['SERVICES', GREEN]]);
	body(s, 4.872, 1.144, 3.898, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore dolore magna aliqua. Ut enim ad minim venim quis nostrud exercitation ullamco laboris nisi');
	panel(s, 4.563, 3.105, 6.262, 3.286, GREEN);
	[['01', 'SERVICE ONE', 3.765, 3.653, 1.427],
	 ['02', 'SERVICE TWO', 4.993, 4.881, 1.49]].forEach(([n, ttl, yCard, yTxt, wLbl]) => {
		numberCard(s, 5.171, yCard, 0.738, 0.738, WHITE, n, GREEN);
		label(s, 6.163, yTxt, wLbl, ttl, WHITE);
		body(s, 6.163, yTxt + 0.336, 3.132, 0.627, LOREM_SERVICE, { color: WHITE });
	});
	pageNumber(s, '09.');
}

// 10 — Comprehensive range of offerings
function slide10(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 1.361, 1.213, 6.448, 1.043,
		[['COMPREHENSIVE RANGE OF REAL ', DARK], ['ESTATE OFFERINGS', GREEN]]);
	body(s, 8.756, 1.005, 3.442, 1.46,
		'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut laboreet ' +
		'dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisi ut ' +
		'aliquip ex ea commodo');
	panel(s, 1.317, 3.381, 7.438, 3.048, GREEN);
	[['01', 'SERVICE ONE', 3.915, 3.802, 1.427],
	 ['02', 'SERVICE TWO', 5.156, 5.044, 1.49]].forEach(([n, ttl, yCard, yTxt, wLbl]) => {
		numberCard(s, 3.155, yCard, 0.738, 0.738, WHITE, n, GREEN);
		label(s, 4.207, yTxt, wLbl, ttl, WHITE);
		body(s, 4.207, yTxt + 0.337, 4.156, 0.627,
			'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
			{ color: WHITE });
	});
	pageNumber(s, '10.');
}

// 11 — Property types (green number cards on white)
function slide11(s) {
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.361, 1.371, 4.503, 1.043,
		[['PROPERTY TYPES AND ', DARK], ['SPECIALIZATIONS', GREEN]]);
	[['01', 'SERVICE ONE', 1.474, 1.361, 1.427],
	 ['02', 'SERVICE TWO', 4.216, 4.104, 1.49]].forEach(([n, ttl, xCard, xTxt, wLbl]) => {
		numberCard(s, xCard, 2.772, 0.738, 0.738, GREEN, n, WHITE);
		label(s, xTxt, 3.726, wLbl, ttl, DARK);
		body(s, xTxt, 4.063, 2.277, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing');
	});
	body(s, 1.361, 4.947, 5.019, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elitdo eiusmod tempor incididunt ut labore ' +
		'et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nis ' +
		'alquip ex ea commodo consequat. ');
	pageNumber(s, '11.');
}

// 12 — How we can help (green panel left, mixed-case heading)
function slide12(s) {
	panel(s, 0, 0, 2.986, 7.5, GREEN);
	edgeTabs(s, 'h');
	badge(s, 'right');
	heading(s, 5.873, 1.238, 6.256, 1.043,
		[['How We Can Help You Achieve ', DARK], ['Your Real Estate Goals', GREEN]]);
	body(s, 5.873, 2.457, 6.256, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elsed eiusmod tempor incididunt ut labore et ' +
		'dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation ullamco laboris nisi ut aliquip ' +
		'ex commodo consequat. Duis aute irure reprehenderit in voluptate velit esse cillum dolore eu ' +
		'fugiat nulla pariatur. ');
	[['01', 'SERVICE ONE', 5.986, 5.873, 1.427],
	 ['02', 'SERVICE TWO', 9.346, 9.233, 1.49]].forEach(([n, ttl, xCard, xTxt, wLbl]) => {
		numberCard(s, xCard, 4.066, 0.738, 0.738, GREEN, n, WHITE);
		label(s, xTxt, 5.021, wLbl, ttl, DARK);
		body(s, xTxt, 5.357, 2.896, 0.904,
			'Lorem ipsum dolor amet, consectetur adipiscing elit dos eiusmod incididunt labore et dolore magna aliqua.');
	});
	pageNumber(s, '12.');
}

// 13 — Break slide
function slide13(s) {
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.155, 1.851, 5.68, 1.111,
		[['BREAK ', DARK], ['SLIDES', GREEN]], { size: SZ_BIG, wrap: false });
	body(s, 1.155, 3.046, 6.667, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim minim veniam, nostrud exercitation ullamco laboris nisi ' +
		'ut aliquip ex ea commodo consequat. Duis aute irure inreprehenderit voluptate velit esse cillum ' +
		'dolore eu fugiat nulla pariatur. ');
	dot(s, 1.255, 5.025, 0.222, GREEN);
	label(s, 1.96, 4.689, 2.483, 'WRITE SOMETHING HERE', DARK);
	body(s, 1.96, 5.025, 5.862, 0.627,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod incididunt ut labore et dolore magna aliqua. ');
	pageNumber(s, '13.');
}

// 14 — Meet our agents (title + footer only)
function slide14(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 3.401, 0.869, 6.531, 0.572,
		[['MEET OUR REAL ', DARK], ['ESTATE AGENTS', GREEN]], { align: 'center' });
	body(s, 1.452, 6.004, 10.429, 0.627, LOREM_WIDE, { align: 'center' });
	pageNumber(s, '14.');
}

// 15 — Experienced agents (two green name cards)
function slide15(s) {
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.305, 1.888, 4.662, 1.043,
		[['GET TO KNOW OUR ', DARK], ['EXPERIENCED AGENTS', GREEN]]);
	body(s, 1.305, 3.251, 5.361, 1.46,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris ' +
		'nisi aliquip commodo consequat. Duis aute irure dolor in reprehenderit voluptate velit esse ' +
		'cillum dolore eu fugiat nulla pariatur. ');
	body(s, 1.305, 4.707, 5.361, 0.904,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod veniam, quis nostrud ' +
		'exercitation ullamco laboris nisi aliquip commodo cillum dolore eu fugiat nulla pariatur. ');
	[['MICHAEL SEYMOUR', 1.643, 2.025, 2.023], ['RAYMOND NORTON', 4.133, 4.512, 2.065]]
		.forEach(([name, yCard, yTxt, wLbl]) => {
			panel(s, 10.029, yCard, 3.304, 1.728, GREEN);
			label(s, 10.452, yTxt, wLbl, name, WHITE);
			body(s, 10.452, yTxt + 0.336, 2.458, 0.627,
				'Lorem ipsum dolor consectetur adipiscing elit eiusmod', { color: WHITE });
		});
	pageNumber(s, '15.');
}

// 16 — Expertise and commitment (green panel left)
function slide16(s) {
	panel(s, 0, 0, 2.665, 7.5, GREEN);
	edgeTabs(s, 'h');
	badge(s, 'right');
	heading(s, 8.182, 2.003, 3.741, 1.515,
		[['EXPERTISE AND ', DARK], ['COMMITMENT TO ', GREEN], ['CLIENTS', DARK]]);
	body(s, 8.182, 3.76, 4.152, 1.737,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'laboret dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris ' +
		'nisi ut aliquip commodo consequat. Duis aute irure dolor reprehenderit volupta velit esse cillum ' +
		'dolore eu fugiat nulla pariatur. ');
	pageNumber(s, '16.');
}

// 17 — Trusted advisors (green side panel right)
function slide17(s) {
	panel(s, 10.125, 0, 3.208, 7.5, GREEN);
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.199, 1.29, 5.358, 1.043,
		[['YOUR TRUSTED ADVISORS ', DARK], ['IN REAL ESTATE', GREEN]]);
	body(s, 1.199, 2.589, 5.908, 1.182, LOREM_XL);
	[['MARK ROBINSON', 4.147, 1.813], ['JAMES JOHNSON', 5.247, 1.832]].forEach(([name, y, wLbl]) => {
		dot(s, 1.312, y + 0.337, 0.222, GREEN);
		label(s, 2.016, y, wLbl, name, DARK);
		body(s, 2.016, y + 0.337, 5.092, 0.627, LOREM_BULLET);
	});
	pageNumber(s, '17.', WHITE);
}

// 18 — Design and architecture
function slide18(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 1.445, 1.178, 3.277, 1.043,
		[['DESIGN AND ', DARK], ['ARCHITECTURE', GREEN]]);
	body(s, 5.222, 1.109, 6.667, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
		'laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure do reprehenderit in voluptate ' +
		'velit esse cillum dolore eu fugiat nulla pariatur. ');
	pageNumber(s, '18.');
}

// 19 — Elevating your property's aesthetic
function slide19(s) {
	edgeTabs(s, 'v');
	badge(s, 'right');
	heading(s, 6.837, 1.342, 4.238, 1.043,
		[['Elevating Your ', DARK], ["Property's Aesthetic", GREEN]]);
	body(s, 6.837, 2.6, 5.413, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris ' +
		'PLACEHOLDER');
	panel(s, 3.959, 4.669, 8.107, 1.787, GREEN);
	[[4.762, 5.333], [8.444, 9.015]].forEach(([xDot, xTxt]) => {
		dot(s, xDot, 5.452, 0.222, WHITE);
		body(s, xTxt, 5.111, 2.41, 0.904, LOREM_TILE, { color: WHITE });
	});
	pageNumber(s, '19.');
}

// 20 — Interior and exterior trends
function slide20(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 1.393, 1.194, 4.909, 1.043,
		[['INTERIOR AND EXTERIOR ', DARK], ['DESIGN TRENDS', GREEN]]);
	body(s, 1.393, 2.429, 5.354, 0.904,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut');
	panel(s, 4.079, 4.669, 7.986, 1.787, GREEN);
	[[4.809, 5.38], [8.444, 9.015]].forEach(([xDot, xTxt]) => {
		dot(s, xDot, 5.452, 0.222, WHITE);
		body(s, xTxt, 5.111, 2.41, 0.904, LOREM_TILE, { color: WHITE });
	});
	pageNumber(s, '20.');
}

// 21 — Enhancing home value (green tick panel top-left)
function slide21(s) {
	edgeTabs(s, 'v');
	badge(s, 'right');
	panel(s, 0, 1.238, 3.713, 3.048, GREEN);
	[[1.919, 1.687], [2.681, 2.449], [3.442, 3.21]].forEach(([yDot, yTxt]) => {
		dot(s, 0.726, yDot, 0.165, WHITE);
		body(s, 1.126, yTxt, 2.033, 0.627, LOREM_SHORT, { color: WHITE });
	});
	heading(s, 1.445, 5.145, 4.571, 1.043, [['ENHANCING THE VALUE OF YOUR HOME', DARK]]);
	body(s, 6.54, 5.076, 5.349, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris ' +
		'nisi aliquip commodo aute irure do reprehenderit in velit fugiat nulla pariatur. ');
	pageNumber(s, '21.');
}

// 22 — Pricing (three package cards)
function slide22(s) {
	edgeTabs(s, 'h');
	badge(s, 'left');
	heading(s, 3.916, 0.869, 5.502, 1.043,
		[['REAL ESTATE PRICING AND ', DARK], ['FINANCING', GREEN]], { align: 'center' });
	[['$15', 'PACKAGE ONE', 1.152], ['$25', 'PACKAGE TWO', 4.897], ['$35', 'PACKAGE THREE', 8.642]]
		.forEach(([price, name, x]) => {
			panel(s, x, 2.34, 3.539, 3.236, GREEN);
			numberCard(s, x + 1.276, 2.851, 0.987, 0.738, WHITE, price, GREEN);
			s.addText(name, {
				x, y: 3.958, w: 3.539, h: 0.337,
				fontFace: HEAD, fontSize: SZ_LBL, color: WHITE,
				align: 'center', valign: 'top', wrap: false,
			});
			body(s, x + 0.413, 4.295, 2.703, 0.904,
				'Lorem ipsum dolor sit consectetur adipiscing elit dos eiusmod incididunt labore et dolore magna.',
				{ color: WHITE, align: 'center' });
		});
	body(s, 1.452, 6.004, 10.429, 0.627, LOREM_WIDE, { align: 'center' });
	pageNumber(s, '22.');
}

// 23 — Contact information
function slide23(s) {
	edgeTabs(s, 'v');
	badge(s, 'left');
	heading(s, 1.393, 1.194, 4.909, 0.572,
		[['CONTACT ', DARK], ['INFORMATION', GREEN]]);
	body(s, 1.393, 2.013, 6.667, 1.182,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
		'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
		'laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure do reprehenderit in voluptate ' +
		'velit esse cillum dolore eu fugiat nulla pariatur. ');
	panel(s, 0, 4.469, 9.143, 1.787, GREEN);
	[['PHONE NUMBER', 0.688, 1.21, 1.701], ['OUR ADDRESS', 4.91, 5.433, 1.531]]
		.forEach(([name, xDot, xTxt, wLbl]) => {
			dot(s, xDot, 5.223, 0.222, WHITE);
			label(s, xTxt, 4.886, wLbl, name, WHITE);
			body(s, xTxt, 5.223, 3.202, 0.627,
				'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod', { color: WHITE });
		});
	pageNumber(s, '23.');
}

// 24 — Thank you (closing wedge mirrors the cover)
function slide24(s) {
	edgeTabs(s, 'l');
	gradientWedge(s, 7.7145, 'rtl');
	badge(s, 'left');
	heading(s, 1.151, 2.396, 4.988, 1.111,
		[['THANK ', DARK], ['YOU', GREEN]], { size: SZ_BIG, wrap: false });
	s.addText('Real Estate Presentation', {
		x: 1.193, y: 3.338, w: 3.435, h: 0.337,
		fontFace: BODY, fontSize: SZ_LBL, color: GRAY,
		charSpacing: 3, valign: 'top', wrap: false,
	});
	body(s, 1.193, 4.114, 5.317, 0.904, LOREM_COVER);
	s.addText('24.', {
		x: 0.452, y: 6.706, w: 0.5, h: 0.337,
		fontFace: HEAD, fontSize: SZ_LBL, color: DARK, valign: 'top', wrap: false,
	});
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
const BUILDERS = [
	slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
	slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
	slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
];

const pres = new pptxgen();
pres.defineLayout({ name: 'ESTETIKA', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'ESTETIKA';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.author = 'ESTETIKA';
pres.title = 'ESTETIKA — Real Estate Presentation';

BUILDERS.forEach(build => {
	const slide = pres.addSlide();
	slide.background = { color: WHITE };
	build(slide);
});

pres.writeFile({
	fileName: path.join(__dirname, '1640887d-44bf-4eb1-82fc-72629fa739ab_grok_final.pptx'),
}).then(f => console.log('wrote ' + f));
