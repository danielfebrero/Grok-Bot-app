/*
 * "Photograph" — photographer pitch-deck (18 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs only. The sage "Graphikal PlaceHolder" panels are the
 * template's empty picture placeholders and are drawn as flat colour blocks;
 * the pin / check-box icons and device mock-ups are built from native shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / typography
 * ------------------------------------------------------------------ */
const SAGE = 'A5AA94'; // brand green
const INK = '595959'; // headings and body copy
const PAPER = 'F2F2F2'; // slide background
const CARD = 'F3F3F3'; // polaroid frame (40% white hatch over the paper)
const WHITE = 'FFFFFF';
const MUTED = 'BFBFBF'; // dimmed feature rows
const PIN = '4D4D4D'; // map markers
const DARK = '404040'; // device bezel
const STEEL = '808080'; // device side buttons

const SERIF = 'Merriweather'; // headings
const BODY = 'Yeseva One'; // body copy
const SCRIPT = 'Sacramento'; // hand-written accents

const CARD_SHADOW = { type: 'outer', blur: 13, offset: 15, angle: 45, color: '000000', opacity: 0.1 };
const NO_LINE = { type: 'none' };
const NO_FILL = { type: 'none' };

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

// The template's polaroid frames all use roundRect adj=2493 (of the short side).
const CORNER = 0.02493;

/* ------------------------------------------------------------------ *
 * Small building blocks
 * ------------------------------------------------------------------ */

// Rectangle helper that degrades to a plain rect when the radius is zero,
// because PowerPoint's roundRect would otherwise fall back to its default adj.
function box(s, x, y, w, h, r, opts) {
	s.addShape(r > 0 ? 'roundRect' : 'rect', Object.assign({
		x: x, y: y, w: w, h: h, rectRadius: r,
	}, opts));
}

// White polaroid frame with a soft drop shadow.
function card(s, x, y, w, h, rot, r) {
	box(s, x, y, w, h, r === undefined ? CORNER * Math.min(w, h) : r, {
		rotate: rot || 0, fill: { color: CARD }, line: NO_LINE, shadow: CARD_SHADOW,
	});
}

// Sage image placeholder plus its caption, exactly as the template ships it.
// The caption never wraps — in the original it spills past narrow frames.
function photo(s, x, y, w, h, rot, r) {
	box(s, x, y, w, h, r || 0, { rotate: rot || 0, fill: { color: SAGE }, line: NO_LINE });
	photoIcon(s, x + w / 2, y + h / 2, rot || 0);
	s.addText('Graphikal PlaceHolder', {
		x: x, y: y, w: w, h: h, rotate: rot || 0, wrap: false,
		align: 'center', valign: 'top', fontFace: BODY, fontSize: 11, bold: true, color: PAPER,
	});
}

/* The little framed-landscape glyph an empty picture frame shows. Its size is
 * fixed (it never scales with the frame) and it always sits dead centre. */
const ICON_W = 0.86;
const ICON_H = 0.67;
const ICON_PARTS = [
	// [shape, x, y, w, h, fill, line] as fractions of the glyph box
	['rect', 0.00, 0.00, 1.00, 1.00, 'FFFFFF', 'A6A6A6'],
	['ellipse', 0.159, 0.163, 0.158, 0.204, 'F7DC8F', 'E8A33D'],
	['triangle', 0.110, 0.575, 0.490, 0.310, '83BEEC', '4A90D9'],
	['triangle', 0.330, 0.400, 0.560, 0.485, '83BEEC', '4A90D9'],
];
function photoIcon(s, cx, cy, rot) {
	ICON_PARTS.forEach(function (p) {
		const w = p[3] * ICON_W, h = p[4] * ICON_H;
		// centre of this part before the frame's own rotation is applied
		const px = cx + (p[1] + p[3] / 2 - 0.5) * ICON_W;
		const py = cy + (p[2] + p[4] / 2 - 0.5) * ICON_H;
		const c = rotateAbout(px, py, cx, cy, rot);
		s.addShape(p[0], {
			x: c[0] - w / 2, y: c[1] - h / 2, w: w, h: h, rotate: rot,
			fill: { color: p[5] }, line: { color: p[6], width: 0.75 },
		});
	});
}
function rotateAbout(px, py, cx, cy, deg) {
	const a = deg * Math.PI / 180, dx = px - cx, dy = py - cy;
	return [cx + dx * Math.cos(a) - dy * Math.sin(a), cy + dx * Math.sin(a) + dy * Math.cos(a)];
}

/* Text presets used across the deck. */
function heading(s, t, x, y, w, h, o) {
	s.addText(t, Object.assign({
		x: x, y: y, w: w, h: h, fontFace: SERIF, fontSize: 32, bold: true, color: INK,
		valign: 'top', fit: 'resize',
	}, o || {}));
}
function subhead(s, t, x, y, w, h, o) {
	s.addText(t, Object.assign({
		x: x, y: y, w: w, h: h, fontFace: SERIF, fontSize: 14, bold: true, color: INK,
		valign: 'top', fit: 'resize',
	}, o || {}));
}
function para(s, t, x, y, w, h, o) {
	s.addText(t, Object.assign({
		x: x, y: y, w: w, h: h, fontFace: BODY, fontSize: 11, color: INK,
		valign: 'top', lineSpacingMultiple: 2, fit: 'resize',
	}, o || {}));
}
// Hand-written caption pinned under a tilted polaroid.
function script(s, t, x, y, w, h, rot, size, bold) {
	s.addText(t, {
		x: x, y: y, w: w, h: h, rotate: rot || 0, align: 'center', valign: 'top', fit: 'resize',
		fontFace: SCRIPT, fontSize: size || 20, bold: !!bold, color: SAGE,
	});
}

// Running header, repeated on most slides.
function tagline(s) {
	s.addText('See, Not Remember The Memories', {
		x: 0.683, y: 0.719, w: 4.716, h: 0.286, fit: 'resize',
		fontFace: BODY, fontSize: 11, color: SAGE, valign: 'top',
	});
}

// Camera-settings strip: along the bottom bar, or turned up the right edge.
const EXPOSURE = [
	{ t: 'ISO 200', w: 0.684, hx: 0.657, vx: 12.149, vy: 4.491 },
	{ t: '1/60', w: 0.498, hx: 1.702, vx: 12.242, vy: 3.539 },
	{ t: 'F 1.0', w: 0.503, hx: 2.553, vx: 12.240, vy: 2.686 },
];
function exposure(s, color, vertical) {
	EXPOSURE.forEach(function (e) {
		s.addText(e.t, {
			x: vertical ? e.vx : e.hx, y: vertical ? e.vy : 6.478, w: e.w, h: 0.286,
			rotate: vertical ? 270 : 0, wrap: false, align: 'center', valign: 'top', fit: 'resize',
			fontFace: BODY, fontSize: 11, color: color,
		});
	});
}

function pageNum(s, n, color) {
	s.addText('Page ' + n, {
		x: 11.167, y: 6.421, w: 1.506, h: 0.399, align: 'right', valign: 'middle',
		fontFace: SERIF, fontSize: 11, bold: true, color: color,
	});
}

/* Background furniture that lives on the template's layouts. */
function bottomBar(s) {
	s.addShape('rect', { x: 0, y: 6.477, w: SLIDE_W, h: 1.023, fill: { color: SAGE }, line: NO_LINE });
}
function leftBar(s) {
	s.addShape('rect', { x: 0, y: 0, w: 0.76, h: SLIDE_H, fill: { color: SAGE }, line: NO_LINE });
}
function rightPanel(s) {
	s.addShape('rect', { x: 8.939, y: 0, w: 4.394, h: SLIDE_H, fill: { color: SAGE }, line: NO_LINE });
}
// Slanted sage wedge filling the right third (slides 2 and 9).
function rightWedge(s) {
	s.addShape('custGeom', {
		x: 7.154, y: 0, w: 6.179, h: 7.522, fill: { color: SAGE }, line: NO_LINE,
		points: [
			{ x: 2.637, y: 0 }, { x: 6.179, y: 0.011 },
			{ x: 6.179, y: 7.511 }, { x: 0, y: 7.522 }, { close: true },
		],
	});
}
// Slanted sage band across the top (slide 10).
function topWedge(s) {
	s.addShape('custGeom', {
		x: 0, y: 0, w: SLIDE_W, h: 1.554, fill: { color: SAGE }, line: NO_LINE,
		points: [
			{ x: 0, y: 0 }, { x: SLIDE_W, y: 0 },
			{ x: SLIDE_W, y: 1.554 }, { x: 0, y: 0.815 }, { close: true },
		],
	});
}

/* Map-pin marker: a disc with a downward spike and a punched-out centre.
 * Fractions are of the square frame the original icon occupied. */
function mapPin(s, x, y, size, hole) {
	s.addShape('ellipse', { x: x + 0.287 * size, y: y + 0.155 * size, w: 0.426 * size, h: 0.426 * size, fill: { color: PIN }, line: NO_LINE });
	s.addShape('triangle', { x: x + 0.296 * size, y: y + 0.450 * size, w: 0.408 * size, h: 0.395 * size, rotate: 180, fill: { color: PIN }, line: NO_LINE });
	s.addShape('ellipse', { x: x + 0.422 * size, y: y + 0.282 * size, w: 0.170 * size, h: 0.170 * size, fill: { color: hole || PAPER }, line: NO_LINE });
}

/* Ticked check-box icon used in the pricing table. */
function checkBox(s, x, y, size, color) {
	const p = size * 0.16, k = size - 2 * p; // the glyph sits inside a padded square
	s.addShape('rect', { x: x + p, y: y + p, w: k, h: k, fill: NO_FILL, line: { color: color, width: 1.25 } });
	s.addShape('line', { x: x + p + k * 0.2, y: y + p + k * 0.5, w: k * 0.2, h: k * 0.22, line: { color: color, width: 1.5 } });
	s.addShape('line', { x: x + p + k * 0.4, y: y + p + k * 0.26, w: k * 0.4, h: k * 0.46, flipV: true, line: { color: color, width: 1.5 } });
}

/* ------------------------------------------------------------------ *
 * Slide 1 — title
 * ------------------------------------------------------------------ */
function slide01(s) {
	bottomBar(s);
	s.addText('Photograph', {
		x: 0.634, y: 2.996, w: 4.716, h: 1.01, wrap: false, valign: 'top', fit: 'resize',
		fontFace: SERIF, fontSize: 54, bold: true, color: INK,
	});
	s.addText('Photographer Pitch Deck Presentation', {
		x: 0.624, y: 3.936, w: 4.716, h: 0.303, align: 'center', valign: 'top', fit: 'resize',
		fontFace: BODY, fontSize: 12, bold: true, charSpacing: 3, color: INK,
	});
	card(s, 6.402, 1.901, 3.592, 4.677, 352.93);
	exposure(s, PAPER, false);
	s.addText('Justin Rolland', {
		x: 7.4, y: 5.765, w: 2.267, h: 0.572, rotate: 352.8, wrap: false, valign: 'top', fit: 'resize',
		fontFace: SCRIPT, fontSize: 28, bold: true, color: SAGE,
	});
	tagline(s);
	photo(s, 6.567, 2.15, 3.106, 3.39, 353.29);
	card(s, 8.796, 0.822, 3.592, 4.677, 6.76);
	dateStamp(s);
	photo(s, 9.055, 1.083, 3.106, 3.39, 7.07);
}

// "Monday, 18 Sept 23~" — same run mix on slides 1 and 18.
function dateStamp(s) {
	s.addText([
		{ text: 'Monday, ', options: { fontSize: 20 } },
		{ text: '18', options: { fontSize: 16, bold: true } },
		{ text: ' Sept ', options: { fontSize: 20 } },
		{ text: '23~', options: { fontSize: 16, bold: true } },
	], {
		x: 9.532, y: 4.766, w: 2.388, h: 0.438, rotate: 6.8, wrap: false, fit: 'resize',
		align: 'right', valign: 'top', fontFace: SCRIPT, color: SAGE,
	});
}

/* ------------------------------------------------------------------ *
 * Slide 2 — table of contents (6 polaroid tiles)
 * ------------------------------------------------------------------ */
const TOC = [
	{ label: 'About Us', n: '02', x: 5.680, y: 0.80 },
	{ label: 'Portfolio', n: '06', x: 7.908, y: 0.80 },
	{ label: 'Teams', n: '09', x: 10.136, y: 0.80 },
	{ label: 'Services', n: '11', x: 3.453, y: 4.12 },
	{ label: 'Others', n: '13', x: 5.680, y: 4.12 },
	{ label: 'Contacts', n: '16', x: 7.908, y: 4.12 },
];
function slide02(s) {
	rightWedge(s);
	s.addText([
		{ text: 'Table', options: { breakLine: true } },
		{ text: 'Of Contents' },
	], { x: 0.76, y: 1.675, w: 4.653, h: 1.919, align: 'right', valign: 'top', fit: 'resize', fontFace: SERIF, fontSize: 54, bold: true, color: INK });
	TOC.forEach(function (t) {
		card(s, t.x, t.y, 1.982, 2.58);
		s.addText(t.label, { x: t.x + 0.116, y: t.y + 2.109, w: 1.301, h: 0.37, valign: 'top', fit: 'resize', fontFace: SCRIPT, fontSize: 16, color: SAGE });
		s.addText(t.n, { x: t.x + 0.616, y: t.y + 2.109, w: 1.301, h: 0.37, align: 'right', valign: 'top', fit: 'resize', fontFace: SCRIPT, fontSize: 16, bold: true, color: SAGE });
	});
	pageNum(s, 2, PAPER);
	TOC.forEach(function (t) { photo(s, t.x + 0.14, t.y + 0.148, 1.702, 1.86); });
}

/* ------------------------------------------------------------------ *
 * Slide 3 — company introduction
 * ------------------------------------------------------------------ */
function slide03(s) {
	leftBar(s);
	card(s, 8.526, 1.757, 2.908, 4.041, 17.73);
	photo(s, 8.817, 2.046, 2.529, 2.94, 18.37);
	card(s, 7.126, 1.269, 2.908, 4.041);
	s.addText('Hello Everyone!', { x: 7.339, y: 4.695, w: 2.497, h: 0.438, align: 'center', valign: 'top', fit: 'resize', fontFace: SCRIPT, fontSize: 20, italic: true, color: SAGE });
	para(s, [
		{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', options: { breakLine: true } },
		{ text: ' ', options: { breakLine: true } },
		{ text: 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident.' },
	], 1.778, 2.898, 4.543, 2.641, { align: 'right' });
	s.addText([
		{ text: 'Company', options: { breakLine: true } },
		{ text: 'Introductions' },
	], { x: 1.802, y: 1.242, w: 4.52, h: 1.178, align: 'right', valign: 'top', fit: 'resize', fontFace: SERIF, fontSize: 32, bold: true, color: INK });
	subhead(s, 'Jennifer Stallion', 1.802, 5.808, 4.52, 0.337, { align: 'right' });
	s.addText('Founder of Photon Company Ltd.', { x: 1.802, y: 6.017, w: 4.52, h: 0.303, align: 'right', valign: 'top', fit: 'resize', fontFace: BODY, fontSize: 12, color: INK });
	photo(s, 7.331, 1.556, 2.497, 2.913);
	exposure(s, SAGE, true);
	pageNum(s, 3, SAGE);
}

/* ------------------------------------------------------------------ *
 * Slide 4 — scattered polaroids around a centred statement
 * ------------------------------------------------------------------ */
const MOMENTS = [
	{ label: 'First Born', cx: 6.903, cy: -1.081, cr: 343.85, px: 6.996, py: -0.904, pr: 344.39, lx: 7.229, ly: 1.313 },
	{ label: null, cx: 3.796, cy: 5.135, cr: 7.11, px: 3.975, py: 5.312, pr: 7.42 },
	{ label: 'Vacations', cx: 11.091, cy: 1.068, cr: 7.56, px: 11.281, py: 1.237, pr: 7.75, lx: 10.937, ly: 3.498 },
	{ label: 'Maternal', cx: 0.450, cy: 2.160, cr: 349.27, px: 0.569, py: 2.329, pr: 349.65, lx: 0.668, ly: 4.579 },
	{ label: 'Birthday', cx: 9.049, cy: 4.354, cr: 354.78, px: 9.185, py: 4.513, pr: 355.61, lx: 9.156, ly: 6.790 },
];
// "Wedding Day" is stacked above everything else in the source deck.
const WEDDING = { label: 'Wedding Day', cx: 1.815, cy: -0.178, cr: 10.09, px: 2.029, py: -0.029, pr: 10.64, lx: 1.610, ly: 2.243 };

function moment(s, m) {
	card(s, m.cx, m.cy, 2.283, 2.972, m.cr);
	if (m.label) script(s, m.label, m.lx, m.ly, 2.283, 0.438, m.cr);
}
function slide04(s) {
	MOMENTS.forEach(function (m) { moment(s, m); });
	heading(s, 'Remember A Moment In A Blink Of A Lens', 4.033, 2.425, 5.267, 1.178, { align: 'center' });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut', 3.766, 3.574, 5.762, 1.16, { align: 'center' });
	MOMENTS.forEach(function (m) { photo(s, m.px, m.py, 1.967, 2.148, m.pr); });
	moment(s, WEDDING);
	photo(s, WEDDING.px, WEDDING.py, 1.967, 2.148, WEDDING.pr);
	pageNum(s, 4, SAGE);
}

/* ------------------------------------------------------------------ *
 * Slide 5 — client testimonies
 * ------------------------------------------------------------------ */
const QUOTE = '\u201CLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.\u201D';
const TESTIMONIALS = [
	{ name: 'Jeferson Breakson', cx: 0.876, cy: 0.976, cr: 349.90, px: 0.981, py: 1.139, pr: 350.96, lx: 1.061, ly: 3.452, lr: 349.75, qx: 3.168, qy: 1.866 },
	{ name: 'Christin Aguelera', cx: 3.766, cy: 3.671, cr: 7.41, px: 3.950, py: 3.835, pr: 8.03, lx: 3.594, ly: 6.151, lr: 7.39, qx: 0.547, qy: 4.756 },
	{ name: 'Jonathan Maxwell', cx: 9.525, cy: 3.595, cr: 6.12, px: 9.680, py: 3.737, pr: 6.73, lx: 9.408, ly: 6.071, lr: 5.90, qx: 6.386, qy: 4.409 },
];
function slide05(s) {
	TESTIMONIALS.forEach(function (t) { card(s, t.cx, t.cy, 2.314, 3.013, t.cr); });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut eni.', 8.106, 2.37, 4.543, 0.79, { align: 'right' });
	heading(s, 'Client Testimonies', 8.473, 1.147, 4.1, 1.178, { align: 'right' });
	TESTIMONIALS.forEach(function (t) {
		script(s, t.name, t.lx, t.ly, 2.314, 0.438, t.lr, 20, true);
		para(s, QUOTE, t.qx, t.qy, 3.218, 1.16, { align: 'center' });
	});
	pageNum(s, 5, SAGE);
	TESTIMONIALS.forEach(function (t) { photo(s, t.px, t.py, 2.011, 2.193, t.pr); });
}

/* ------------------------------------------------------------------ *
 * Slide 6 — portfolio opener
 * ------------------------------------------------------------------ */
function slide06(s) {
	heading(s, 'Photography Portfolio', 0.672, 2.358, 3.467, 1.178);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 3.536, 4.195, 1.16);
	pageNum(s, 6, SAGE);
	card(s, 6.082, -1.011, 3.179, 4.139, 346.96);
	card(s, 6.075, 3.998, 3.179, 4.139, 320.86);
	card(s, 9.320, 0.465, 3.179, 4.139, 9.16);
	exposure(s, SAGE, false);
	tagline(s);
	photo(s, 6.071, 4.299, 2.757, 3.008, 321.39);
	photo(s, 6.212, -0.729, 2.757, 3.008, 347.94);
	photo(s, 9.597, 0.682, 2.757, 3.008, 8.97);
}

/* ------------------------------------------------------------------ *
 * Slide 7 — detailed portfolio, maternal
 * ------------------------------------------------------------------ */
function slide07(s) {
	pageNum(s, 7, SAGE);
	heading(s, 'Detailed Portfolio Maternal Photoshoot', 0.672, 1.601, 5.062, 1.178);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 2.92, 4.606, 1.16);
	card(s, 0.76, 4.356, 5.451, 1.508, 0, 0);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', 1.813, 4.953, 4.195, 0.628, { align: 'right', lineSpacingMultiple: 1.5 });
	subhead(s, 'Clyde Park, West Virginia NV - 45466', 1.489, 4.585, 4.52, 0.337, { align: 'right' });
	mapPin(s, 1.033, 4.753, 0.678, CARD);
	photo(s, 6.667, 1.601, 3.836, 1.982);
	photo(s, 6.667, 3.965, 5.906, 1.982);
	card(s, 10.425, 1.292, 2.253, 2.933, 7.26);
	photo(s, 10.618, 1.467, 1.937, 2.103, 7.64);
	exposure(s, SAGE, false);
	tagline(s);
	script(s, 'By Christopher Sin', 10.241, 3.699, 2.314, 0.37, 7.39, 16, true);
}

/* ------------------------------------------------------------------ *
 * Slide 8 — company outing portfolio (2x2 grid)
 * ------------------------------------------------------------------ */
const OUTING = [
	{ cx: 6.642, cy: 0.542, px: 6.844, py: 0.711 },
	{ cx: 9.243, cy: 0.542, px: 9.420, py: 0.688 },
	{ cx: 6.667, cy: 3.845, px: 6.844, py: 4.014 },
	{ cx: 9.267, cy: 3.845, px: 9.420, py: 3.992 },
];
function slide08(s) {
	pageNum(s, 8, SAGE);
	OUTING.forEach(function (o) { card(s, o.cx, o.cy, 2.258, 2.94); });
	card(s, 0.76, 3.166, 4.52, 2.694, 0, 0);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', 0.989, 3.763, 4.195, 0.628, { lineSpacingMultiple: 1.5 });
	subhead(s, 'Client: Graphikal Pty. Ltd.', 0.989, 3.395, 4.52, 0.337);
	s.addText([
		{ text: 'Shoot By:', options: { fontFace: BODY, italic: true, breakLine: true } },
		{ text: 'Haikal Farhan', options: { fontFace: SERIF, bold: true } },
	], { x: 2.2, y: 5.051, w: 2.323, h: 0.572, valign: 'top', fit: 'resize', fontSize: 14, color: INK });
	heading(s, 'Portfolio of Company Outing Photoshoot', 0.672, 1.83, 5.062, 1.178);
	OUTING.forEach(function (o) { photo(s, o.px, o.py, 1.939, 2.119); });
	photo(s, 1.103, 4.574, 0.945, 0.956);
	exposure(s, SAGE, true);
}

/* ------------------------------------------------------------------ *
 * Slide 9 — photographer profile with skill bars
 * ------------------------------------------------------------------ */
const SKILLS = [
	{ label: 'Skill Number One', y: 4.395, fill: 3.876 },
	{ label: 'Skill Number Two', y: 5.098, fill: 4.973 },
	{ label: 'Skill Number Three', y: 5.801, fill: 4.517 },
];
function slide09(s) {
	rightWedge(s);
	heading(s, 'James Came Around', 0.672, 1.812, 5.062, 0.64);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 2.533, 4.606, 1.16);
	subhead(s, 'Our Professional Photographer', 0.672, 1.509, 4.52, 0.337);
	SKILLS.forEach(function (k) {
		s.addShape('rect', { x: 0.76, y: k.y, w: 4.973, h: 0.19, fill: NO_FILL, line: { color: SAGE, width: 0.75 } });
		s.addShape('rect', { x: 0.76, y: k.y, w: k.fill, h: 0.19, fill: { color: SAGE }, line: NO_LINE });
		s.addText(k.label, {
			x: 0.672, y: k.y - 0.419, w: 5.062, h: 0.419, valign: 'top', lineSpacingMultiple: 2, fit: 'resize',
			fontFace: BODY, fontSize: 11, bold: true, color: INK,
		});
	});
	card(s, 7.188, 1.21, 3.902, 5.08);
	pageNum(s, 9, PAPER);
	exposure(s, PAPER, true);
	photo(s, 7.464, 1.5, 3.351, 3.661);
	s.addText([
		{ text: '$450/', options: { fontSize: 24 } },
		{ text: 'Sessions', options: { fontSize: 12, italic: true } },
	], { x: 7.838, y: 5.458, w: 2.601, h: 0.505, align: 'center', valign: 'top', fit: 'resize', fontFace: SERIF, bold: true, color: INK });
}

/* ------------------------------------------------------------------ *
 * Slide 10 — the team, with ring gauges
 * ------------------------------------------------------------------ */
const CREW = [
	{ role: 'Lighting Guy', cx: 7.080, cy: 0.605, cr: 3.74, px: 7.244, py: 0.778, pr: 3.28, lx: 7.027, ly: 2.969, lr: 3.12 },
	{ role: 'Helper', cx: 9.436, cy: 0.631, cr: 349.42, px: 9.537, py: 0.811, pr: 349.64, lx: 9.643, ly: 2.988, lr: 350.03 },
	{ role: 'Model', cx: 6.939, cy: 3.679, cr: 353.01, px: 7.059, py: 3.850, pr: 352.84, lx: 7.082, ly: 6.065, lr: 353.40 },
	{ role: 'Area Clearer', cx: 9.242, cy: 3.386, cr: 0, px: 9.391, py: 3.562, pr: 0, lx: 9.266, ly: 5.791, lr: 0 },
];
const GAUGES = [
	{ label: 'Data 1', x: 2.150, lx: 1.943 },
	{ label: 'Data 2', x: 3.704, lx: 3.498 },
	{ label: 'Data 3', x: 5.259, lx: 5.053 },
];
function slide10(s) {
	topWedge(s);
	CREW.forEach(function (c) { card(s, c.cx, c.cy, 2.227, 2.899, c.cr); });
	pageNum(s, 10, SAGE);
	heading(s, 'People That Makes The Project Works', 1.01, 2.391, 5.267, 1.178, { align: 'right' });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut', 0.515, 3.54, 5.762, 1.16, { align: 'right' });
	GAUGES.forEach(function (g) {
		// 75% arc: a sage pie with a background-coloured disc punched out of it.
		s.addShape('pie', {
			x: g.x, y: 4.934, w: 1.017, h: 1.017, flipH: true,
			angleRange: [269.4, 183.16], fill: { color: SAGE }, line: NO_LINE,
		});
		s.addShape('ellipse', { x: g.x + 0.057, y: 4.991, w: 0.903, h: 0.903, fill: { color: PAPER }, line: NO_LINE });
		s.addText('75%', { x: g.x + 0.057, y: 4.991, w: 0.903, h: 0.903, align: 'center', valign: 'middle', fontFace: SERIF, fontSize: 12, bold: true, color: INK });
		s.addText(g.label, { x: g.lx, y: 5.949, w: 1.43, h: 0.337, align: 'center', valign: 'top', fit: 'resize', fontFace: SERIF, fontSize: 14, bold: true, color: SAGE });
	});
	CREW.forEach(function (c) {
		photo(s, c.px, c.py, 1.929, 2.07, c.pr);
		script(s, c.role, c.lx, c.ly, 2.216, 0.37, c.lr, 16, true);
	});
}

/* ------------------------------------------------------------------ *
 * Slide 11 — services, diagonal cascade of polaroids (all tilted 15deg)
 * ------------------------------------------------------------------ */
const SERVICES = [
	{ label: 'Maternal', cx: 7.312, cy: -0.583, px: 7.549, py: -0.407, lx: 7.007, ly: 1.818 },
	{ label: 'Wedding', cx: 6.463, cy: 2.585, px: 6.700, py: 2.761, lx: 6.158, ly: 4.987 },
	{ label: null, cx: 5.614, cy: 5.753, px: 5.839, py: 5.929 },
	{ label: 'Vacation', cx: 10.331, cy: -1.637, px: 10.567, py: -1.461, lx: 10.026, ly: 0.765 },
	{ label: 'First Born', cx: 9.482, cy: 1.531, px: 9.718, py: 1.707, lx: 9.177, ly: 3.933 },
	{ label: 'Birthday', cx: 8.634, cy: 4.699, px: 8.857, py: 4.875, lx: 8.328, ly: 7.101 },
];
function slide11(s) {
	rightPanel(s);
	heading(s, 'Our List Of Photography Services', 0.672, 2.194, 3.467, 1.717);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 3.873, 4.195, 1.16);
	pageNum(s, 11, PAPER);
	exposure(s, SAGE, false);
	tagline(s);
	SERVICES.forEach(function (v) {
		card(s, v.cx, v.cy, 2.274, 2.961, 15);
		if (v.label) script(s, v.label, v.lx, v.ly, 2.274, 0.438, 15, 20, true);
	});
	SERVICES.forEach(function (v) { photo(s, v.px, v.py, 1.941, 2.138, 15); });
}

/* ------------------------------------------------------------------ *
 * Slides 12 & 13 share a text column; 12 shows a photo mosaic, 13 a map.
 * ------------------------------------------------------------------ */
function textColumn(s, title, sub, n) {
	exposure(s, SAGE, false);
	tagline(s);
	heading(s, title, 0.672, 2.328, 3.467, 1.178);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 4.007, 4.195, 1.16);
	subhead(s, sub, 0.672, 3.506, 4.52, 0.337);
	pageNum(s, n, SAGE);
}

function slide12(s) {
	textColumn(s, 'Indoor Studio Photoshoot', 'Also, We Can Shoot In Our Indoor Studio', 12);
	photo(s, 6.667, 1.165, 5.906, 1.620, 0, 0.149);
	photo(s, 6.667, 2.937, 2.784, 3.398, 0, 0.123);
	photo(s, 9.615, 2.937, 2.958, 1.620, 0, 0.149);
	photo(s, 9.615, 4.715, 2.958, 1.620, 0, 0.149);
}

/* Continental-US silhouette, traced from the reference and normalised to the
 * map frame: the northern border left-to-right, then the coast right-to-left. */
const US_OUTLINE = [
	[0.000, 0.301], [0.011, 0.276], [0.021, 0.198], [0.032, 0.167], [0.042, 0.126],
	[0.053, 0.011], [0.074, 0.028], [0.095, 0.003], [0.127, 0.017], [0.159, 0.031],
	[0.191, 0.041], [0.223, 0.052], [0.254, 0.062], [0.286, 0.071], [0.318, 0.078],
	[0.350, 0.085], [0.382, 0.091], [0.413, 0.096], [0.445, 0.099], [0.477, 0.100],
	[0.509, 0.102], [0.530, 0.106], [0.541, 0.119], [0.562, 0.120], [0.583, 0.139],
	[0.604, 0.136], [0.615, 0.139], [0.625, 0.175], [0.647, 0.156], [0.657, 0.190],
	[0.678, 0.180], [0.700, 0.184], [0.721, 0.198], [0.731, 0.277], [0.753, 0.372],
	[0.774, 0.354], [0.795, 0.330], [0.816, 0.280], [0.837, 0.277], [0.859, 0.219],
	[0.880, 0.194], [0.912, 0.182], [0.933, 0.163], [0.943, 0.089], [0.965, 0.068],
	[0.975, 0.088], [0.996, 0.140], [0.996, 0.160], [0.975, 0.184], [0.965, 0.307],
	[0.943, 0.317], [0.922, 0.352], [0.901, 0.420], [0.890, 0.598], [0.869, 0.621],
	[0.859, 0.641], [0.848, 0.992], [0.837, 0.987], [0.816, 0.959], [0.795, 0.917],
	[0.774, 0.837], [0.753, 0.815], [0.731, 0.833], [0.710, 0.813], [0.689, 0.815],
	[0.668, 0.816], [0.657, 0.816], [0.647, 0.874], [0.625, 0.863], [0.604, 0.856],
	[0.583, 0.856], [0.562, 0.847], [0.541, 0.857], [0.519, 0.884], [0.498, 0.890],
	[0.488, 0.909], [0.477, 0.996], [0.456, 0.989], [0.435, 0.963], [0.413, 0.900],
	[0.392, 0.843], [0.371, 0.840], [0.350, 0.866], [0.329, 0.839], [0.307, 0.779],
	[0.286, 0.747], [0.265, 0.744], [0.244, 0.755], [0.212, 0.747], [0.180, 0.724],
	[0.148, 0.693], [0.117, 0.665], [0.085, 0.658], [0.064, 0.591], [0.042, 0.567],
	[0.021, 0.495], [0.011, 0.407], [0.000, 0.313],
];
/* Interior state boundaries, kept to a readable skeleton. */
const US_STATE_LINES = [
	[0.075, 0.020, 0.085, 0.610], [0.150, 0.030, 0.160, 0.690], [0.245, 0.058, 0.250, 0.750],
	[0.330, 0.080, 0.335, 0.845], [0.415, 0.096, 0.420, 0.900], [0.495, 0.100, 0.500, 0.890],
	[0.560, 0.120, 0.575, 0.850], [0.640, 0.150, 0.650, 0.865], [0.700, 0.185, 0.715, 0.815],
	[0.030, 0.220, 0.180, 0.170], [0.045, 0.400, 0.340, 0.230], [0.070, 0.520, 0.500, 0.330],
	[0.100, 0.640, 0.560, 0.430], [0.240, 0.740, 0.640, 0.560], [0.330, 0.840, 0.700, 0.700],
	[0.620, 0.230, 0.800, 0.290], [0.700, 0.420, 0.870, 0.400], [0.660, 0.560, 0.860, 0.560],
	[0.700, 0.700, 0.830, 0.760],
];
const MAP = { x: 5.191, y: 1.15, w: 7.548, h: 4.714 };

function slide13(s) {
	textColumn(s, 'Our Available Locations', 'We can Shot In Three Different Area', 13);
	s.addShape('custGeom', {
		x: MAP.x, y: MAP.y, w: MAP.w, h: MAP.h, fill: NO_FILL, line: { color: SAGE, width: 0.5 },
		points: US_OUTLINE.map(function (p) { return { x: p[0] * MAP.w, y: p[1] * MAP.h }; }).concat([{ close: true }]),
	});
	US_STATE_LINES.forEach(function (g) {
		s.addShape('custGeom', {
			x: MAP.x, y: MAP.y, w: MAP.w, h: MAP.h, fill: NO_FILL, line: { color: SAGE, width: 0.5 },
			points: [{ x: g[0] * MAP.w, y: g[1] * MAP.h }, { x: g[2] * MAP.w, y: g[3] * MAP.h }],
		});
	});
	[[9.623, 2.822], [8.593, 3.553], [9.181, 2.205]].forEach(function (p) { mapPin(s, p[0], p[1], 1.175); });
}

/* ------------------------------------------------------------------ *
 * Slide 14 — two phone mock-ups
 * ------------------------------------------------------------------ */
// Side buttons, relative to the phone body's top-left corner.
const PHONE_BUTTONS = [[2.319, 1.376, 0.640], [-0.101, 0.649, 0.239], [-0.179, 1.139, 0.396], [-0.179, 1.614, 0.396]];
// Corner ticks on the bezel.
const PHONE_CORNERS = [[0.001, 4.803], [0.004, 0.410], [2.591, 4.803], [2.591, 0.410]];

function phone(s, bx, by, sx, sy) {
	PHONE_BUTTONS.forEach(function (b) {
		s.addShape('roundRect', { x: bx + b[0], y: by + b[1], w: b[2], h: 0.092, rotate: 90, rectRadius: 0.046, fill: { color: STEEL }, line: NO_LINE });
	});
	s.addShape('roundRect', { x: bx, y: by, w: 2.653, h: 5.427, rectRadius: 0.328, fill: { color: SAGE }, line: { color: STEEL, width: 0.75 } });
	s.addShape('roundRect', { x: bx + 0.051, y: by + 0.049, w: 2.550, h: 5.330, rectRadius: 0.253, fill: { color: DARK }, line: NO_LINE });
	photo(s, sx, sy, 2.426, 5.206, 0, 0.236);
	s.addShape('round2SameRect', { x: bx + 0.859, y: by + 0.230, w: 0.929, h: 0.180, rectRadius: 0.09, fill: { color: DARK }, line: NO_LINE });
	PHONE_CORNERS.forEach(function (c) {
		s.addShape('rect', { x: bx + c[0], y: by + c[1], w: 0.058, h: 0.048, fill: { color: '262626' }, line: NO_LINE });
	});
}
function conceptColumn(s, n) {
	heading(s, 'Our Photography Concept', 0.672, 1.759, 3.467, 1.717);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.', 0.672, 4.007, 4.195, 1.16);
	subhead(s, 'Put The Subtitle Here', 0.672, 3.506, 4.52, 0.337);
	exposure(s, SAGE, false);
	tagline(s);
	pageNum(s, n, SAGE);
}
function slide14(s) {
	conceptColumn(s, 14);
	phone(s, 6.124, 1.337, 6.237, 1.442);
	phone(s, 9.115, 0.719, 9.228, 0.830);
}

/* ------------------------------------------------------------------ *
 * Slide 15 — tablet mock-up (landscape)
 * ------------------------------------------------------------------ */
const TABLET_BUTTONS = [[5.5965, 1.1165, 0.107, 0.320], [6.089, 0.797, 0.267, 0.107], [6.367, 0.797, 0.267, 0.107]];
const TABLET_LENS = [[5.728, 3.636, 0.079, '000000'], [5.728, 3.452, 0.079, '181717'], [5.752, 3.313, 0.033, '000000'], [5.751, 3.476, 0.033, '1D2787']];

function slide15(s) {
	conceptColumn(s, 15);
	TABLET_BUTTONS.forEach(function (b) {
		s.addShape('roundRect', { x: b[0], y: b[1], w: b[2], h: b[3], rectRadius: 0.05, fill: { color: 'A6A6A6' }, line: NO_LINE });
	});
	s.addShape('roundRect', { x: 5.6265, y: 0.8275, w: 6.948, h: 5.325, rectRadius: 0.363, fill: { color: SAGE }, line: NO_LINE });
	s.addShape('roundRect', { x: 5.6815, y: 0.8825, w: 6.838, h: 5.217, rectRadius: 0.323, fill: { color: DARK }, line: NO_LINE });
	TABLET_LENS.forEach(function (o) {
		s.addShape('ellipse', { x: o[0], y: o[1], w: o[2], h: o[2], fill: { color: o[3] }, line: NO_LINE });
	});
	photo(s, 5.83, 1.053, 6.532, 4.885, 0, 0.264);
}

/* ------------------------------------------------------------------ *
 * Slide 16 — pricing table
 * ------------------------------------------------------------------ */
const FEATURES = [
	'Lorem ipsum dolor sit amet, consectetur',
	'Lorem ipsum dolor sit amet',
	'Lorem ipsum sit amet, consectetur',
	'Lorem ipsum dolor sit amet, consectetur',
];
const PLANS = [
	{
		title: 'Special Offers', price: '$699/', card: { x: 0.762, y: 0.812, bg: WHITE },
		band: { x: 0.760, y: 2.634, w: 3.933, color: INK }, onBand: WHITE,
		priceBox: { x: 1.144, y: 1.603, w: 3.164 }, titleX: 1.299, titleY: 1.341,
		bullets: { x: 1.246, y: 3.936 }, featX: 2.542, featY: 3.004, iconX: 1.802, iconY: 2.858,
		active: 4,
	},
	{
		title: 'Normal package', price: '$499/', card: { x: 4.689, y: 0.797, bg: PAPER },
		band: { x: 4.687, y: 2.619, w: 3.931, color: SAGE }, onBand: PAPER,
		priceBox: { x: 5.065, y: 1.588, w: 3.176 }, titleX: 5.226, titleY: 1.326,
		bullets: { x: 5.173, y: 3.921 }, featX: 6.469, featY: 2.989, iconX: 5.728, iconY: 2.843,
		active: 2,
	},
	{
		title: 'Sad Day Service', price: '$99/', card: { x: 8.622, y: 0.797, bg: PAPER },
		band: { x: 8.620, y: 2.619, w: 3.931, color: SAGE }, onBand: PAPER,
		priceBox: { x: 9.159, y: 1.588, w: 2.853 }, titleX: 9.159, titleY: 1.326,
		bullets: { x: 9.106, y: 3.921 }, featX: 10.402, featY: 2.989, iconX: 9.661, iconY: 2.843,
		active: 1,
	},
];
function slide16(s) {
	PLANS.forEach(function (p) {
		s.addShape('rect', { x: p.card.x, y: p.card.y, w: 3.931, h: 5.353, fill: { color: p.card.bg }, line: NO_LINE, shadow: CARD_SHADOW });
		s.addShape('rect', { x: p.band.x, y: p.band.y, w: p.band.w, h: 1.0, fill: { color: p.band.color }, line: NO_LINE });
		s.addText([
			{ text: p.price, options: { fontSize: 40 } },
			{ text: 'Sessions', options: { fontSize: 20 } },
		], { x: p.priceBox.x, y: p.priceBox.y, w: p.priceBox.w, h: 0.673, align: 'center', valign: 'top', margin: 0, fit: 'resize', fontFace: SERIF, bold: true, color: INK });
		s.addText(p.title, { x: p.titleX, y: p.titleY, w: 2.853, h: 0.185, align: 'center', valign: 'top', margin: 0, fit: 'resize', fontFace: BODY, fontSize: 11, bold: true, color: INK });
		s.addText(FEATURES.map(function (f, i) {
			return { text: f, options: { color: i < p.active ? INK : MUTED, breakLine: true } };
		}), {
			x: p.bullets.x + 0.06, y: p.bullets.y, w: 2.737, h: 1.851, align: 'center', valign: 'top', margin: 0,
			fontFace: BODY, fontSize: 11, lineSpacingMultiple: 2.5, fit: 'resize',
			bullet: { characterCode: '2022', indent: 13.5 },
		});
		s.addText('Features:', { x: p.featX, y: p.featY, w: 1.445, h: 0.236, valign: 'top', margin: 0, fit: 'resize', fontFace: SERIF, fontSize: 14, bold: true, color: p.onBand });
		checkBox(s, p.iconX, p.iconY, 0.509, p.onBand);
	});
	pageNum(s, 16, SAGE);
}

/* ------------------------------------------------------------------ *
 * Slide 17 — contact
 * ------------------------------------------------------------------ */
const CONTACT = [
	{ head: 'Studio Hours', hx: 1.178, bx: 1.189, lines: ['Monday - Friday', '8AM-  7PM'] },
	{ head: 'Address', hx: 3.871, bx: 3.882, lines: ['Crypto Off Road N69', 'Baltimore, CA - 45466'] },
	{ head: 'Phone Numbers', hx: 6.658, bx: 6.669, lines: ['+01 (234) 567 8910', '+10 (987) 654 3210'] },
	{ head: 'Email & Website', hx: 9.328, bx: 9.339, lines: ['graphikalcreative@gmail.com', 'www.yourwebsitedomain.com'] },
];
function slide17(s) {
	card(s, 0.76, 4.132, 11.812, 1.604, 0, 0);
	CONTACT.forEach(function (c) {
		s.addText(c.head, { x: c.hx, y: 4.437, w: 2.236, h: 0.269, valign: 'top', margin: 0, fit: 'resize', fontFace: SERIF, fontSize: 16, bold: true, color: INK });
		s.addText(c.lines.map(function (l, i) { return { text: l, options: { breakLine: i === 0 } }; }), {
			x: c.bx, y: 4.868, w: 2.236, h: 0.524, valign: 'top', margin: 0, fit: 'resize',
			fontFace: BODY, fontSize: 11, color: INK, lineSpacingMultiple: 1.5,
		});
	});
	exposure(s, SAGE, false);
	pageNum(s, 17, SAGE);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', 0.657, 2.785, 4.195, 0.628, { lineSpacingMultiple: 1.5 });
	subhead(s, 'Let\u2019s Hope That Our Services Will Satisfy Your Needs', 0.672, 2.231, 4.01, 0.572);
	heading(s, 'Contact Us!', 0.672, 1.444, 5.062, 0.64);
	card(s, 7.106, 0.307, 2.274, 2.961, 354.61);
	script(s, 'Contact Us', 7.22, 2.78, 2.274, 0.438, 354.61, 20, true);
	photo(s, 7.252, 0.416, 1.97, 2.303, 354.69);
	card(s, 9.205, 0.876, 2.274, 2.961, 3.31);
	script(s, 'Let\u2019s Work', 9.135, 3.353, 2.274, 0.438, 3.31, 20, true);
	photo(s, 9.397, 0.987, 1.97, 2.303, 3.2);
}

/* ------------------------------------------------------------------ *
 * Slide 18 — thank you (mirrors the title slide)
 * ------------------------------------------------------------------ */
function slide18(s) {
	bottomBar(s);
	card(s, 6.402, 1.901, 3.592, 4.677, 352.93);
	s.addText('Thank You!', {
		x: 7.538, y: 5.765, w: 1.992, h: 0.572, rotate: 352.8, wrap: false, valign: 'top', fit: 'resize',
		fontFace: SCRIPT, fontSize: 28, bold: true, color: SAGE,
	});
	s.addText('Thank You!', {
		x: 0.634, y: 2.996, w: 4.569, h: 1.01, wrap: false, valign: 'top', fit: 'resize',
		fontFace: SERIF, fontSize: 54, bold: true, color: INK,
	});
	s.addText('Hope We Can Shoot You (Just Pun)', {
		x: 0.624, y: 3.936, w: 4.579, h: 0.303, align: 'center', valign: 'top', fit: 'resize',
		fontFace: BODY, fontSize: 12, bold: true, charSpacing: 3, color: INK,
	});
	exposure(s, PAPER, false);
	tagline(s);
	photo(s, 6.567, 2.15, 3.106, 3.39, 353.29);
	card(s, 8.796, 0.822, 3.592, 4.677, 6.76);
	dateStamp(s);
	photo(s, 9.055, 1.083, 3.106, 3.39, 7.07);
}

/* ------------------------------------------------------------------ *
 * Assemble
 * ------------------------------------------------------------------ */
const SLIDES = [
	slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
	slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
];

function build() {
	const pres = new PptxGenJS();
	pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
	pres.layout = 'WIDE';
	pres.title = 'Photograph \u2014 Photographer Pitch Deck Presentation';
	pres.author = 'Justin Rolland';

	SLIDES.forEach(function (fn) {
		const s = pres.addSlide();
		s.background = { color: PAPER };
		fn(s);
	});

	return pres;
}

const outFile = path.join(__dirname, '100a2546-3f9e-4337-b0c4-f44f624b70e0_grok_final.pptx');
build().writeFile({ fileName: outFile }).then(function () {
	console.log('wrote ' + outFile);
});
