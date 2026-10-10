/*
 * "Meeting Plan" - Strategic Business Alignment Meeting deck (15 slides, 13.333 x 7.5 in).
 * Rebuilt with pptxgenjs only. Raster photos in the original are drawn here as
 * grey "[image]" placeholder blocks at the same position / size.
 *
 * Run:  node 0f567ad5-13d6-40f9-bcd4-d3d037e4779a_grok_final.js
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */

const RED = 'DB1717';
const RED_DARK = 'A41111';
const ORANGE = 'FF9305';
const YELLOW = 'FADA1F';
const PINK = 'FACCC9';
const PEACH = 'F9C0AB';
const SKIN = 'FFC0B1';
const NAVY = '102A54';
const PAPER = 'F3F6FD';
const PAPER_BLUE = 'CCE6FF';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const INK = '0D0D0D';
const CHARCOAL = '262626';
const GREY = '404040';
const GREY_MID = '595959';
const GREY_LIGHT = 'F2F2F2';

const MAJOR = 'Poppins Medium'; // theme major latin font (+mj-lt)
const MINOR = 'Poppins Light'; // theme minor latin font (+mn-lt)

const IMG_FILL = 'CCCCCC'; // stand-in colour for the original photographs
const IMG_LABEL = '9C9C9C';

// Every card / pill in the deck uses the same soft drop shadow.
// pptxgenjs rewrites the object it is handed, so build a fresh one per shape.
function softShadow () {
	return { type: 'outer', color: '7F7F7F', opacity: 0.11, blur: 37, offset: 0, angle: 0 };
}

/* ------------------------------------------------------------------ helpers */

/** Plain text box: top aligned, theme body font, 12 pt, black - like the deck's default. */
function txt (slide, content, opts) {
	slide.addText(content, Object.assign({
		fontFace: MINOR, fontSize: 12, color: BLACK, valign: 'top', isTextBox: true
	}, opts));
}

/** Section heading: array of [text, colour] pairs, 32 pt in the display font. */
function heading (slide, parts, opts) {
	const runs = parts.map(([text, color, breakLine]) => ({ text, options: { color, breakLine: !!breakLine } }));
	txt(slide, runs, Object.assign({ fontFace: MAJOR, fontSize: 32, color: BLACK }, opts));
}

/** Fully rounded "pill" tag. */
function pill (slide, x, y, w, h, color, shadow) {
	slide.addShape('roundRect', {
		x, y, w, h, rectRadius: Math.min(w, h) / 2,
		fill: { color }, line: { type: 'none' },
		shadow: shadow ? softShadow() : undefined
	});
}

/** Soft rounded card. `adj` is the OOXML corner adjust value (fraction of the short side). */
function card (slide, x, y, w, h, color, adj, shadow) {
	slide.addShape('roundRect', {
		x, y, w, h, rectRadius: (adj === undefined ? 0.16667 : adj) * Math.min(w, h),
		fill: { color }, line: { type: 'none' },
		shadow: shadow ? softShadow() : undefined
	});
}

function dot (slide, x, y, d, color) {
	slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' } });
}

/**
 * Grey stand-in for one of the original photographs.
 * `flat` optionally squares off one side ('left' | 'right') of the rounded box.
 */
function imagePlaceholder (slide, x, y, w, h, radius, flat) {
	const r = radius || 0;
	if (r) {
		slide.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color: IMG_FILL }, line: { type: 'none' } });
	} else {
		slide.addShape('rect', { x, y, w, h, fill: { color: IMG_FILL }, line: { type: 'none' } });
	}
	if (flat === 'left') slide.addShape('rect', { x, y, w: r, h, fill: { color: IMG_FILL }, line: { type: 'none' } });
	if (flat === 'right') slide.addShape('rect', { x: x + w - r, y, w: r, h, fill: { color: IMG_FILL }, line: { type: 'none' } });
	txt(slide, '[image]', {
		x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', valign: 'middle',
		fontSize: 12, color: IMG_LABEL
	});
}

/**
 * The recurring "radar" decoration: five hairline concentric circles filled at a
 * few percent black. `size` is the bounding box of the whole motif.
 */
const RING_RADII = [0.1826, 0.2616, 0.3406, 0.4195, 0.4985];
function decoRings (slide, x, y, size, alphaPct) {
	const cx = x + size / 2;
	const cy = y + size / 2;
	RING_RADII.forEach(f => {
		const d = 2 * f * size;
		slide.addShape('ellipse', {
			x: cx - d / 2, y: cy - d / 2, w: d, h: d,
			fill: { type: 'none' },
			line: { color: BLACK, width: 0.0031 * size * 72, transparency: 100 - (alphaPct || 6) }
		});
	});
}

/** Quarter-circle arcs anchored on the bottom-right slide corner (slide 15). */
function cornerArcs (slide, cx, cy, radii, color) {
	radii.forEach(r => {
		slide.addShape('ellipse', {
			x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
			fill: { type: 'none' }, line: { color, width: 2.5 }
		});
	});
}

/** "Starlaight" mark: three concentric rings of dots forming an open circle. */
function logo (slide, x, y, color) {
	const cx = x + 0.2257 / 2;
	const cy = y + 0.2305 / 2;
	[[8, 0.088, 0.044], [12, 0.056, 0.026], [16, 0.104, 0.020]].forEach(([n, r, d]) => {
		for (let i = 0; i < n; i++) {
			const a = (2 * Math.PI * i) / n;
			dot(slide, cx + r * Math.cos(a) - d / 2, cy + r * Math.sin(a) - d / 2, d, color);
		}
	});
}

/**
 * Straight limb / stick drawn between two points.
 * A vertical bar of length L rotated by atan2(-dx, dy) points along (dx, dy).
 */
function limb (slide, x1, y1, x2, y2, thick, color, radius) {
	const dx = x2 - x1;
	const dy = y2 - y1;
	const len = Math.hypot(dx, dy);
	const deg = (Math.atan2(-dx, dy) * 180) / Math.PI;
	slide.addShape('roundRect', {
		x: (x1 + x2) / 2 - thick / 2, y: (y1 + y2) / 2 - len / 2, w: thick, h: len,
		rectRadius: radius === undefined ? thick / 2 : radius, rotate: deg,
		fill: { color }, line: { type: 'none' }
	});
}

/** Header strip repeated on every slide. */
function header (slide, onRed) {
	const c = onRed ? WHITE : BLACK;
	logo(slide, 0.6024, 0.4813, onRed ? WHITE : RED);
	txt(slide, 'Starlaight Corp', { x: 0.8281, y: 0.4451, w: 1.5078, h: 0.3029, color: c });
	txt(slide, 'Strategic Business Alignment Meeting', { x: 2.5616, y: 0.4409, w: 2.2096, h: 0.5049, color: c });
	txt(slide, 'July 30, 2025', { x: 9.2491, y: 0.4451, w: 1.2662, h: 0.3029, color: c, bold: true });
	txt(slide, 'www.starlaightcorp.com', { x: 10.6348, y: 0.4451, w: 2.2506, h: 0.3029, color: c });
}

/* -------------------------------------------------------------------- icons */

/** Head + shoulders, the building block of the people glyphs. */
function figure (slide, cx, cy, scale, color) {
	dot(slide, cx - 0.06 * scale, cy - 0.20 * scale, 0.12 * scale, color);
	slide.addShape('roundRect', {
		x: cx - 0.115 * scale, y: cy - 0.06 * scale, w: 0.23 * scale, h: 0.28 * scale,
		rectRadius: 0.085 * scale, fill: { color }, line: { type: 'none' }
	});
}

/** Two shoppers carrying bags (slide 6, card 1). */
function shoppersIcon (slide, cx, cy, color) {
	figure(slide, cx - 0.12, cy + 0.02, 1, color);
	figure(slide, cx + 0.13, cy - 0.03, 0.95, color);
	[[cx - 0.28, cy + 0.10], [cx + 0.02, cy + 0.10], [cx + 0.27, cy + 0.05]].forEach(([bx, by]) => {
		slide.addShape('rect', { x: bx, y: by, w: 0.12, h: 0.13, fill: { color }, line: { type: 'none' } });
	});
}

/** Two colleagues with a dollar sign and a rising arrow (slide 6, card 2 / slide 9). */
function peopleIcon (slide, cx, cy, color) {
	figure(slide, cx - 0.16, cy - 0.02, 0.9, color);
	figure(slide, cx + 0.05, cy - 0.05, 1.05, color);
	txt(slide, '$', { x: cx + 0.16, y: cy - 0.30, w: 0.26, h: 0.26, fontSize: 13, bold: true, color, align: 'center' });
	slide.addShape('rect', { x: cx + 0.05, y: cy + 0.19, w: 0.30, h: 0.045, rotate: -22, fill: { color }, line: { type: 'none' } });
	slide.addShape('triangle', { x: cx + 0.28, y: cy + 0.06, w: 0.12, h: 0.11, rotate: 68, fill: { color }, line: { type: 'none' } });
}

/** Analyst standing beside a growing bar chart (slide 6, card 3). */
function analystIcon (slide, cx, cy, color) {
	figure(slide, cx - 0.13, cy, 1.25, color);
	[[0.06, 0.12], [0.17, 0.19]].forEach(([dx, h]) => {
		slide.addShape('rect', { x: cx + dx, y: cy + 0.22 - h, w: 0.075, h, fill: { color }, line: { type: 'none' } });
	});
	slide.addShape('rect', { x: cx + 0.285, y: cy - 0.08, w: 0.06, h: 0.30, fill: { color }, line: { type: 'none' } });
	slide.addShape('triangle', { x: cx + 0.25, y: cy - 0.20, w: 0.13, h: 0.13, fill: { color }, line: { type: 'none' } });
}

/** Three team members linked around a cog (slide 6, card 4). */
function teamIcon (slide, cx, cy, color) {
	figure(slide, cx, cy - 0.16, 0.7, color);
	figure(slide, cx - 0.25, cy + 0.19, 0.7, color);
	figure(slide, cx + 0.25, cy + 0.19, 0.7, color);
	slide.addShape('gear6', { x: cx - 0.10, y: cy - 0.03, w: 0.20, h: 0.20, fill: { color }, line: { type: 'none' } });
	[[-0.19, -0.05], [0.19, -0.05], [-0.09, 0.20], [0.09, 0.20], [0, -0.14]].forEach(([dx, dy]) => {
		dot(slide, cx + dx - 0.017, cy + dy, 0.034, color);
	});
}

/** Single standing figure (infographic wheel, slide 10). */
function personIcon (slide, cx, cy, color) {
	dot(slide, cx - 0.037, cy - 0.20, 0.075, color);
	slide.addShape('roundRect', { x: cx - 0.094, y: cy - 0.105, w: 0.188, h: 0.30, rectRadius: 0.06, fill: { color }, line: { type: 'none' } });
}

function cartIcon (slide, cx, cy, color) {
	slide.addShape('rect', { x: cx - 0.20, y: cy - 0.10, w: 0.41, h: 0.06, fill: { color }, line: { type: 'none' } });
	slide.addShape('rect', { x: cx - 0.16, y: cy - 0.10, w: 0.06, h: 0.20, fill: { color }, line: { type: 'none' } });
	dot(slide, cx - 0.10, cy + 0.10, 0.07, color);
	dot(slide, cx + 0.06, cy + 0.10, 0.07, color);
}

function tagIcon (slide, cx, cy, color) {
	slide.addShape('round1Rect', { x: cx - 0.16, y: cy - 0.14, w: 0.32, h: 0.28, rectRadius: 0.09, rotate: -45, fill: { color }, line: { type: 'none' } });
	dot(slide, cx - 0.09, cy - 0.05, 0.06, RED);
}

function phoneIcon (slide, cx, cy, color) {
	slide.addShape('roundRect', { x: cx - 0.09, y: cy - 0.15, w: 0.18, h: 0.30, rectRadius: 0.07, rotate: -25, fill: { color }, line: { type: 'none' } });
}

/** Ribbon rosette with a tick (slide 4). */
function badgeIcon (slide, cx, cy, color) {
	for (let i = 0; i < 12; i++) {
		const a = (i * Math.PI) / 6;
		dot(slide, cx + 0.165 * Math.cos(a) - 0.045, cy + 0.165 * Math.sin(a) - 0.045, 0.09, color);
	}
	dot(slide, cx - 0.155, cy - 0.155, 0.31, color);
	txt(slide, '\u2713', { x: cx - 0.16, y: cy - 0.155, w: 0.32, h: 0.3, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: RED });
}

/** Filled circle with a tick, used for the check lists. */
function tickIcon (slide, x, y, d, ring, mark) {
	dot(slide, x, y, d, ring);
	txt(slide, '\u2713', { x: x - 0.09, y: y - 0.06, w: d + 0.18, h: d + 0.12, align: 'center', valign: 'middle', fontSize: 19, bold: true, color: mark });
}

function bellIcon (slide, cx, cy, color) {
	slide.addShape('roundRect', { x: cx - 0.10, y: cy - 0.11, w: 0.20, h: 0.18, rectRadius: 0.085, fill: { color }, line: { type: 'none' } });
	slide.addShape('rect', { x: cx - 0.13, y: cy + 0.03, w: 0.26, h: 0.045, fill: { color }, line: { type: 'none' } });
	dot(slide, cx - 0.03, cy + 0.07, 0.06, color);
}

function bulbIcon (slide, cx, cy, color) {
	dot(slide, cx - 0.085, cy - 0.14, 0.17, color);
	slide.addShape('rect', { x: cx - 0.04, y: cy + 0.02, w: 0.08, h: 0.09, fill: { color }, line: { type: 'none' } });
}

function calendarIcon (slide, cx, cy, color) {
	slide.addShape('roundRect', { x: cx - 0.11, y: cy - 0.09, w: 0.22, h: 0.20, rectRadius: 0.04, fill: { color }, line: { type: 'none' } });
	slide.addShape('rect', { x: cx - 0.08, y: cy - 0.14, w: 0.035, h: 0.07, fill: { color }, line: { type: 'none' } });
	slide.addShape('rect', { x: cx + 0.045, y: cy - 0.14, w: 0.035, h: 0.07, fill: { color }, line: { type: 'none' } });
}

function coinsIcon (slide, cx, cy, color) {
	[-0.09, -0.02, 0.05].forEach(dy => {
		slide.addShape('ellipse', { x: cx - 0.11, y: cy + dy, w: 0.22, h: 0.075, fill: { color }, line: { type: 'none' } });
	});
}

/** Camera aperture: a disc cut into six blades by white slashes (slide 13). */
function apertureIcon (slide, cx, cy, d, color) {
	dot(slide, cx - d / 2, cy - d / 2, d, color);
	for (let i = 0; i < 6; i++) {
		const a = (i * Math.PI) / 3 + Math.PI / 6;
		limb(slide, cx + d * 0.16 * Math.cos(a), cy + d * 0.16 * Math.sin(a),
			cx + d * 0.55 * Math.cos(a - 1.05), cy + d * 0.55 * Math.sin(a - 1.05), d * 0.09, WHITE, 0);
	}
}

/** Dart planted in a tilted ring target (slide 13). */
function targetArrowIcon (slide, cx, cy, d, color) {
	[[1.0, 0.62], [0.60, 0.37]].forEach(([sx, sy]) => {
		slide.addShape('ellipse', {
			x: cx - (d * sx) / 2, y: cy - (d * sy) / 2 + d * 0.1, w: d * sx, h: d * sy,
			rotate: -20, fill: { type: 'none' }, line: { color, width: d * 9 }
		});
	});
	limb(slide, cx + d * 0.05, cy - d * 0.02, cx - d * 0.45, cy - d * 0.45, d * 0.22, color, d * 0.11);
	slide.addShape('triangle', { x: cx - d * 0.62, y: cy - d * 0.66, w: d * 0.28, h: d * 0.26, rotate: -45, fill: { color }, line: { type: 'none' } });
}

function clockIcon (slide, cx, cy, color) {
	slide.addShape('ellipse', { x: cx - 0.115, y: cy - 0.115, w: 0.23, h: 0.23, fill: { type: 'none' }, line: { color, width: 2 } });
	slide.addShape('rect', { x: cx - 0.008, y: cy - 0.075, w: 0.016, h: 0.085, fill: { color }, line: { type: 'none' } });
	slide.addShape('rect', { x: cx - 0.008, y: cy - 0.008, w: 0.065, h: 0.016, fill: { color }, line: { type: 'none' } });
}

function handsetIcon (slide, cx, cy, color) {
	dot(slide, cx - 0.142, cy - 0.142, 0.284, color);
	slide.addShape('roundRect', { x: cx - 0.045, y: cy - 0.085, w: 0.09, h: 0.17, rectRadius: 0.035, rotate: -30, fill: { color: WHITE }, line: { type: 'none' } });
}

function shareIcon (slide, cx, cy, color) {
	dot(slide, cx - 0.157, cy - 0.157, 0.313, color);
	dot(slide, cx - 0.015, cy - 0.09, 0.055, WHITE);
	dot(slide, cx - 0.075, cy + 0.005, 0.055, WHITE);
	dot(slide, cx + 0.025, cy + 0.035, 0.055, WHITE);
}

/* --------------------------------------------------------------- slide  1 */

function slide01 (pptx) {
	const s = pptx.addSlide();
	s.background = { color: RED };

	// circle inherited from the title layout
	s.addShape('ellipse', { x: 3.5205, y: 1.1848, w: 2.0181, h: 2.0181, fill: { color: WHITE, transparency: 85 }, line: { type: 'none' } });

	imagePlaceholder(s, 0, 1.7899, 4.8993, 4.7378, 0.6437, 'left');

	[[-0.4948, 6.9514, 1.0972], [8.0753, -4.6984, 9.5714], [11.4149, 5.7609, 3.5308]].forEach(([x, y, d]) => {
		s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: WHITE, transparency: 85 }, line: { type: 'none' } });
	});

	header(s, true);

	txt(s, 'Strategic Business Alignment Meeting', { x: 5.7656, y: 1.7899, w: 4.8693, h: 0.4039, fontSize: 18, color: WHITE });
	txt(s, 'Meeting Plan', {
		x: 5.7656, y: 2.5257, w: 6.5382, h: 2.8007,
		fontFace: MAJOR, fontSize: 110, color: WHITE, lineSpacingMultiple: 0.7
	});
	txt(s, 'PLACEHOLDER', {
		x: 5.7656, y: 5.1152, w: 3.8028, h: 0.6782, color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});

	s.addShape('roundRect', {
		x: 5.8438, y: 6.0357, w: 2.5838, h: 0.4921, rectRadius: 0.246,
		fill: { color: WHITE, transparency: 85 }, line: { type: 'none' }
	});
	txt(s, 'Start Presentation', {
		x: 5.9974, y: 6.0966, w: 2.2766, h: 0.3702, align: 'center', fontFace: MAJOR, fontSize: 16, color: WHITE
	});
}

/* --------------------------------------------------------------- slide  2 */

const AGENDA_CARDS = [
	{ y: 1.5222, tag: 'Opening Remarks & Goals', tagW: 2.538, tagColor: RED_DARK, tagText: WHITE },
	{ y: 3.3425, tag: 'Business Performance Review', tagW: 2.8179, tagColor: ORANGE, tagText: BLACK },
	{ y: 5.1793, tag: 'Marketing & Sales Strategy', tagW: 2.6096, tagColor: YELLOW, tagText: BLACK }
];
const AGENDA_BODY = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue';

function slide02 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.3432, 4.484, 6.2391);
	imagePlaceholder(s, 0.95, 4.2, 6.5333, 2.547, 0.2578);
	header(s, false);

	heading(s, [['Scheduling ', BLACK, true], ['The Meeting Agenda', RED]], { x: 0.8281, y: 1.5222, w: 4.9031, h: 1.1781 });
	txt(s, 'PLACEHOLDER', {
		x: 0.8281, y: 2.8212, w: 5.9524, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	AGENDA_CARDS.forEach(c => {
		card(s, 8.3477, c.y, 4.3354, 1.5677, WHITE, 0.16667, true);
		pill(s, 8.5779, c.y + 0.2526, c.tagW, 0.3366, c.tagColor, true);
		txt(s, c.tag, { x: 8.6648, y: c.y + 0.2694, w: c.tagW - 0.174, h: 0.3029, color: c.tagText });
		txt(s, AGENDA_BODY, { x: 8.4935, y: c.y + 0.685, w: 4.0438, h: 0.6782, color: GREY, lineSpacingMultiple: 1.5 });
	});
}

/* --------------------------------------------------------------- slide  3 */

const PARTICIPANTS = [
	{ x: 0.9236, y: 4.2191, w: 2.8247, label: 'CEO / Business Owner', tw: 2.1595, fill: RED, text: WHITE, dotColor: WHITE },
	{ x: 3.9273, y: 4.2414, w: 2.3644, label: 'Marketing Director', tw: 1.7354, fill: WHITE, text: BLACK, dotColor: ORANGE },
	{ x: 0.9236, y: 4.9162, w: 2.3264, label: 'Finance Manager', tw: 1.7354, fill: WHITE, text: BLACK, dotColor: RED },
	{ x: 3.5031, y: 4.9162, w: 2.562, label: 'Operations Manager', tw: 1.9424, fill: WHITE, text: BLACK, dotColor: YELLOW },
	{ x: 0.9236, y: 5.6133, w: 2.2361, label: 'Sales Team Lead', tw: 1.6196, fill: WHITE, text: BLACK, dotColor: ORANGE }
];

function slide03 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, 8.5492, 4.5061, 6.055);
	imagePlaceholder(s, 8.85, 1.9327, 3.8707, 4.1727, 0.3918);
	header(s, false);

	heading(s, [['Selecting ', BLACK], ['Participants', RED]], { x: 0.8281, y: 1.8469, w: 5.3385, h: 0.6395 });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', {
		x: 0.8281, y: 2.7688, w: 7.3052, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	PARTICIPANTS.forEach(p => {
		pill(s, p.x, p.y, p.w, 0.4921, p.fill, true);
		txt(s, p.label, { x: p.x + 0.4221, y: p.y + 0.0946, w: p.tw, h: 0.3029, color: p.text });
		dot(s, p.x + 0.3199, p.y + 0.2051, 0.082, p.dotColor);
	});
}

/* --------------------------------------------------------------- slide  4 */

function slide04 (pptx) {
	const s = pptx.addSlide();
	imagePlaceholder(s, 0.6312, 1.8056, 4.5142, 3.4543, 0.2287);
	decoRings(s, -2.1177, -2.1911, 4.9356);
	header(s, false);

	heading(s, [['Decision Making ', BLACK, true], ['And ', BLACK], ['Problem Solving', RED]], { x: 5.852, y: 1.6889, w: 5.156, h: 1.1781 });
	txt(s, 'Stakeholders On Business Priorities', { x: 5.852, y: 3.202, w: 4.7344, h: 0.4039, fontFace: MAJOR, fontSize: 18 });
	txt(s, 'PLACEHOLDER', {
		x: 5.852, y: 3.6833, w: 7.2036, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	[
		{ y: 5.0963, label: 'Sending Invitation And Agenda In Advance', fill: WHITE, shadow: true, dotColor: ORANGE },
		{ y: 5.7818, label: 'Sharing Relevant Documents Or Reports', fill: ORANGE, shadow: false, dotColor: BLACK }
	].forEach(r => {
		pill(s, 6.0004, r.y, 4.375, 0.4921, r.fill, r.shadow);
		txt(s, r.label, { x: 6.4224, y: r.y + 0.0946, w: 3.7568, h: 0.3029 });
		dot(s, 6.3202, r.y + 0.2051, 0.082, r.dotColor);
	});

	// red "top corners rounded" callout in the lower-left corner
	s.addShape('round2SameRect', {
		x: 0.6312, y: 5.4938, w: 2.9854, h: 2.0062, rectRadius: 0.245,
		fill: { color: RED }, line: { type: 'none' }
	});
	badgeIcon(s, 1.22, 5.969, WHITE);
	txt(s, 'Opening Remarks & Goals', { x: 0.9221, y: 6.2903, w: 2.4036, h: 0.3029, fontFace: MAJOR, color: WHITE });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
		x: 0.9221, y: 6.5932, w: 2.4036, h: 0.6301, fontSize: 11, color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});
}

/* --------------------------------------------------------------- slide  5 */

function slide05 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.4355, 3.9984, 6.973);
	imagePlaceholder(s, 0.8281, 3.8667, 5.5527, 2.6577, 0.1759);
	header(s, false);

	heading(s, [['Communication And ', BLACK], ['Pre-meeting Materials', RED]], { x: 0.8281, y: 1.4916, w: 5.5527, h: 1.1781 });
	txt(s, 'Sending Invitation And Agenda In Advance', { x: 0.8281, y: 2.8621, w: 2.6719, h: 0.6395, fontSize: 16 });
	txt(s, 'PLACEHOLDER', {
		x: 7.0675, y: 1.5901, w: 5.5527, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	card(s, 7.0675, 3.0496, 5.2833, 2.0062, RED, 0.10852, false);
	pill(s, 7.287, 3.3839, 3.6068, 0.3366, RED_DARK, true);
	txt(s, 'Sharing Relevant Documents Or Report', { x: 7.3739, y: 3.409, w: 3.4435, h: 0.3029, color: WHITE });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada', {
		x: 7.3739, y: 3.8881, w: 4.7491, h: 0.9078, fontSize: 11, color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});

	pill(s, 7.1083, 5.4201, 2.538, 0.3366, ORANGE, false);
	txt(s, 'Opening Remarks & Goals', { x: 7.1952, y: 5.4369, w: 2.3642, h: 0.3029, color: WHITE });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna', {
		x: 7.0675, y: 5.8942, w: 4.836, h: 0.6301, fontSize: 11, color: GREY, lineSpacingMultiple: 1.5
	});
}

/* --------------------------------------------------------------- slide  6 */

const ROOM_CARDS = [
	{ x: 7.502, y: 1.7387, textY: 2.656, iconY: 2.243, fill: WHITE, text: GREY, icon: RED, glyph: shoppersIcon },
	{ x: 10.0935, y: 1.7387, textY: 2.656, iconY: 2.243, fill: RED, text: GREY_LIGHT, icon: WHITE, glyph: peopleIcon },
	{ x: 7.502, y: 4.153, textY: 5.1775, iconY: 4.693, fill: WHITE, text: GREY, icon: RED, glyph: analystIcon },
	{ x: 10.0935, y: 4.153, textY: 5.1775, iconY: 4.741, fill: WHITE, text: GREY, icon: RED, glyph: teamIcon }
];
const ROOM_BULLET = 'PLACEHOLDER';

function slide06 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.3432, 3.75, 6.973);
	header(s, false);

	heading(s, [['Physical ', BLACK], ['Meeting Room', RED]], { x: 0.8281, y: 1.4358, w: 5.8385, h: 0.6395 });
	txt(s, 'PLACEHOLDER', {
		x: 0.8281, y: 2.1658, w: 6.0052, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	card(s, 0.8281, 3.5667, 5.2833, 2.7984, WHITE, 0.10852, true);
	pill(s, 1.2968, 3.9384, 3.9506, 0.4921, ORANGE, false);
	txt(s, 'Benefits Of Physical Meeting', { x: 1.5609, y: 3.9993, w: 3.4224, h: 0.3702, align: 'center', fontFace: MAJOR, fontSize: 16, color: WHITE });
	[4.6298, 5.3151].forEach(y => {
		txt(s, ROOM_BULLET, {
			x: 1.2968, y, w: 4.3459, h: 0.6782, color: GREY, lineSpacingMultiple: 1.5,
			bullet: { characterCode: '006F', indent: 13.5 }
		});
	});

	ROOM_CARDS.forEach(c => {
		card(s, c.x, c.y, 2.4117, 2.2121, c.fill, 0.10852, true);
		c.glyph(s, c.x + 1.206, c.iconY, c.icon);
		txt(s, 'Lorem dolor sit amet consectetur adipiscing elit sed', {
			x: c.x + 0.2972, y: c.textY, w: 1.8174, h: 0.9812,
			align: 'center', color: c.text, lineSpacingMultiple: 1.5
		});
	});
}

/* --------------------------------------------------------------- slide  7 */

function slide07 (pptx) {
	const s = pptx.addSlide();
	imagePlaceholder(s, 0.8281, 3.8667, 11.651, 3.05, 0.2019);
	decoRings(s, 10.6348, -1.9412, 4.6398);
	header(s, false);

	heading(s, [['Follow Up And ', BLACK], ['Action Item', RED]], { x: 0.8281, y: 1.5158, w: 6.2552, h: 0.6395 });
	pill(s, 0.9115, 2.3129, 3.8598, 0.3366, RED, false);
	txt(s, 'Business Dashboard Or Performance Data', { x: 0.9984, y: 2.3297, w: 3.7169, h: 0.3029, color: WHITE });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', {
		x: 0.8281, y: 2.785, w: 10.9219, h: 0.6782, color: GREY, lineSpacingMultiple: 1.5
	});
}

/* --------------------------------------------------------------- slide  8 */

const MODERATOR_POINTS = [
	{ y: 5.0972, w: 4.2199, h: 0.3857, label: 'Understanding Of Current Business', ring: 'FACECE', mark: RED, cy: 5.1638, d: 0.2699 },
	{ y: 5.6185, w: 3.8453, h: 0.3366, label: 'Aligned Goals Across All Departments', ring: 'FFD49B', mark: ORANGE, cy: 5.6852, d: 0.2459 },
	{ y: 6.1398, w: 3.8453, h: 0.3857, label: 'Specific Action Plans With Owners ', ring: 'FEF8D2', mark: YELLOW, cy: 6.2065, d: 0.2459 }
];

function slide08 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.3432, 1.9375, 8.7855);
	imagePlaceholder(s, 0.656, 1.3542, 6.1458, 6.1458, 0);
	decoRings(s, 10.8012, -2.0505, 4.5953);
	header(s, false);

	heading(s, [['Introduce The Moderator Who Will ', BLACK], ['Guide Us', RED]], { x: 6.5112, y: 1.6296, w: 5.8385, h: 1.1781 });
	pill(s, 6.5112, 3.0389, 3.3906, 0.4921, RED, false);
	txt(s, 'Mrs. Isabella Adams', { x: 6.8088, y: 3.083, w: 2.7954, h: 0.4039, align: 'center', fontFace: MAJOR, fontSize: 18, color: WHITE });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna', {
		x: 6.5112, y: 3.7763, w: 6.1285, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	MODERATOR_POINTS.forEach(p => {
		tickIcon(s, 6.6336, p.cy, p.d, p.ring, p.mark);
		txt(s, p.label, { x: 7.1797, y: p.y, w: p.w, h: p.h, fontSize: 14, lineSpacingMultiple: 1.3 });
	});
}

/* --------------------------------------------------------------- slide  9 */

/** Flat laptop mock-up that framed the screenshot in the original layout. */
function laptop (pptx, s) {
	s.addShape('roundRect', { x: 1.347, y: 3.061, w: 4.556, h: 3.04, rectRadius: 0.11, fill: { color: '111111' }, line: { type: 'none' } });
	s.addShape('rect', { x: 1.472, y: 3.24, w: 4.306, h: 2.69, fill: { color: WHITE }, line: { type: 'none' } });
	s.addShape('roundRect', { x: 0.819, y: 6.104, w: 5.639, h: 0.188, rectRadius: 0.09, fill: { color: 'D6D8DC' }, line: { type: 'none' } });
	s.addShape('roundRect', { x: 3.02, y: 6.104, w: 0.62, h: 0.075, rectRadius: 0.035, fill: { color: 'B7BABF' }, line: { type: 'none' } });
}

function slide09 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.3432, 2.1153, 8.6078, 4);
	laptop(pptx, s);
	imagePlaceholder(s, 1.5, 3.25, 4.2708, 2.6667, 0);
	header(s, false);

	card(s, 0.6024, 1.7952, 2.4117, 2.2121, ORANGE, 0.10852, true);
	peopleIcon(s, 1.75, 2.30, WHITE);
	txt(s, 'Lorem dolor sit amet consectetur adipiscing elit sed', {
		x: 0.8996, y: 2.7124, w: 1.8174, h: 0.9812, align: 'center', color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});

	heading(s, [['Meeting Plan ', BLACK], ['Mockup Design ', RED]], { x: 7.1615, y: 1.7952, w: 3.6719, h: 1.1781 });
	txt(s, 'PLACEHOLDER', {
		x: 7.1615, y: 3.1705, w: 4.8438, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	card(s, 7.2396, 4.56, 5.426, 1.6525, RED, 0.10852, false);
	pill(s, 7.4827, 4.7684, 3.6068, 0.3366, RED_DARK, true);
	txt(s, 'Sharing Relevant Documents Or Report', { x: 7.5696, y: 4.7935, w: 3.4435, h: 0.3029, color: WHITE });
	tickIcon(s, 7.5696, 5.4202, 0.3108, WHITE, RED);
	txt(s, '6.18%', {
		x: 8.0175, y: 5.2328, w: 2.1868, h: 0.7713, fontSize: 44, bold: true, color: WHITE, lineSpacingMultiple: 0.9
	});
	txt(s, 'The European languages are members of the same family. ', {
		x: 9.7788, y: 5.1893, w: 2.6437, h: 0.6782, color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});
}

/* -------------------------------------------------------------- slide  10 */

const WHEEL_SEGMENTS = [
	{ angles: [225, 315], color: ORANGE },
	{ angles: [315, 45], color: RED },
	{ angles: [45, 135], color: PINK },
	{ angles: [135, 225], color: YELLOW }
];
const WHEEL_BUBBLES = [
	{ x: 9.9052, y: 3.0561, d: 0.9329, color: ORANGE, icon: personIcon },
	{ x: 8.9789, y: 3.9854, d: 0.9305, color: YELLOW, icon: cartIcon },
	{ x: 10.8327, y: 3.9854, d: 0.9329, color: RED, icon: tagIcon },
	{ x: 9.9052, y: 4.9123, d: 0.9329, color: PINK, icon: phoneIcon }
];
const WHEEL_TAGS = [
	{ x: 0.9646, y: 2.8038, color: RED, text: WHITE, bodyY: 3.433, bodyX: 0.8575 },
	{ x: 4.3952, y: 2.8038, color: ORANGE, text: BLACK, bodyY: 3.433, bodyX: 4.3454 },
	{ x: 0.9646, y: 4.6912, color: YELLOW, text: BLACK, bodyY: 5.25, bodyX: 0.8575 },
	{ x: 4.3952, y: 4.6912, color: PINK, text: BLACK, bodyY: 5.1681, bodyX: 4.3454 }
];
const WHEEL_BODY = 'PLACEHOLDER';

function slide10 (pptx) {
	const s = pptx.addSlide();
	header(s, false);
	heading(s, [['Infographic ', BLACK], ['Section', RED]], { x: 4.3347, y: 1.3507, w: 4.6639, h: 0.6395, align: 'center' });

	WHEEL_SEGMENTS.forEach(seg => {
		s.addShape('blockArc', {
			x: 8.4371, y: 2.525, w: 3.8692, h: 3.8692,
			angleRange: seg.angles, arcThicknessRatio: 0.2274,
			fill: { color: seg.color }, line: { type: 'none' }
		});
	});
	WHEEL_BUBBLES.forEach(b => {
		dot(s, b.x, b.y, b.d, b.color);
		b.icon(s, b.x + b.d / 2, b.y + b.d / 2, WHITE);
	});

	WHEEL_TAGS.forEach(t => {
		pill(s, t.x, t.y, 1.9791, 0.3702, t.color, false);
		txt(s, 'Subtitle Here', { x: t.x + 0.1882, y: t.y, w: 1.6027, h: 0.3702, align: 'center', fontSize: 16, color: t.text });
		txt(s, WHEEL_BODY, { x: t.bodyX, y: t.bodyY, w: 3.1116, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5 });
	});
}

/* -------------------------------------------------------------- slide  11 */

/** Upward arrow-shaped column. */
function arrowBar (s, x, y, w, h, color) {
	s.addShape('custGeom', {
		x, y, w, h, fill: { color }, line: { type: 'none' },
		points: [{ x: w / 2, y: 0 }, { x: w, y: 0.324 }, { x: w, y: h }, { x: 0, y: h }, { x: 0, y: 0.324 }, { close: true }]
	});
}

/** Skewed plinth under each column; `top` / `bot` are the two slanted edges. */
function arrowBase (s, x, y, w, h, color, top, bot) {
	s.addShape('custGeom', {
		x, y, w, h, fill: { color }, line: { type: 'none' },
		points: [{ x: top[0] * w, y: 0 }, { x: top[1] * w, y: 0 }, { x: bot[1] * w, y: h }, { x: bot[0] * w, y: h }, { close: true }]
	});
}

const GROWTH_BARS = [
	{ x: 6.6905, y: 2.793, w: 1.1361, h: 3.7015, color: RED, icon: bellIcon, iconY: 3.52 },
	{ x: 7.9698, y: 3.5476, w: 1.1362, h: 2.9469, color: ORANGE, icon: bulbIcon, iconY: 4.27 },
	{ x: 9.2491, y: 1.3921, w: 1.1363, h: 5.1023, color: YELLOW, icon: calendarIcon, iconY: 2.11 },
	{ x: 10.5285, y: 4.1318, w: 1.1362, h: 2.3627, color: PINK, icon: coinsIcon, iconY: 4.75 }
];
const GROWTH_BASES = [
	{ x: 6.1919, w: 1.6346, color: RED, top: [0.305, 1.0], bot: [0.0, 0.898] },
	{ x: 7.8036, w: 1.4685, color: ORANGE, top: [0.113, 0.887], bot: [0.0, 1.0] },
	{ x: 9.2491, w: 1.6346, color: YELLOW, top: [0.0, 0.695], bot: [0.102, 1.0] },
	{ x: 10.5285, w: 1.9669, color: PINK, top: [0.0, 0.578], bot: [0.253, 1.0] }
];
const GROWTH_CARDS = [
	{ x: 1.1379, y: 4.8677, color: RED, year: '2025', yearW: 0.3472, badgeText: WHITE },
	{ x: 4.2147, y: 4.8677, color: ORANGE, year: '2026', yearW: 0.3778, badgeText: WHITE },
	{ x: 9.968, y: 2.5704, color: YELLOW, year: '2027', yearW: 1.2564, badgeText: BLACK }
];

function slide11 (pptx) {
	const s = pptx.addSlide();

	GROWTH_BARS.forEach(b => {
		arrowBar(s, b.x, b.y, b.w, b.h, b.color);
		b.icon(s, b.x + b.w / 2, b.iconY, WHITE);
	});
	GROWTH_BASES.forEach(b => arrowBase(s, b.x, 6.4945, b.w, 1.0055, b.color, b.top, b.bot));

	GROWTH_CARDS.forEach(c => {
		card(s, c.x, c.y, 2.6884, 1.1796, WHITE, 0.15401, true);
		dot(s, c.x + 0.2122, c.y - 0.1768, 0.3712, c.color);
		txt(s, '$231,489', { x: c.x + 0.1915, y: c.y + 0.3155, w: 2.1982, h: 0.5722, fontFace: MAJOR, fontSize: 28, color: c.color === YELLOW ? BLACK : c.color });
		txt(s, c.year, { x: c.x + 0.2927, y: c.y + 0.8317, w: c.yearW, h: 0.1346, fontSize: 8, margin: 0 });
		pill(s, c.x + 1.9974, c.y + 0.6768, 0.4753, 0.2232, c.color, false);
		txt(s, '+21', { x: c.x + 1.9974, y: c.y + 0.6493, w: 0.4753, h: 0.2777, align: 'center', valign: 'middle', fontSize: 10.5, color: c.badgeText });
	});

	header(s, false);
	heading(s, [['Infographic ', BLACK], ['Section', RED]], { x: 0.7382, y: 2.3187, w: 4.6639, h: 0.6395 });
	txt(s, 'PLACEHOLDER', {
		x: 0.7382, y: 3.0271, w: 4.8438, h: 0.6782, color: GREY, lineSpacingMultiple: 1.5
	});
}

/* -------------------------------------------------------------- slide  12 */

function oval (s, x, y, w, h, color) {
	s.addShape('ellipse', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/** Concentric-ring archery target: colour / white / colour / white centre. */
function bullseye (s, x, y, w, h, color) {
	oval(s, x, y, w, h, color);
	oval(s, x + w * 0.180, y + h * 0.180, w * 0.640, h * 0.640, WHITE);
	oval(s, x + w * 0.310, y + h * 0.310, w * 0.380, h * 0.380, color);
	oval(s, x + w * 0.415, y + h * 0.415, w * 0.170, h * 0.170, WHITE);
}

/** Arrow: dark shaft with a triangular head at (x2, y2). */
function arrow (s, x1, y1, x2, y2) {
	limb(s, x1, y1, x2, y2, 0.032, CHARCOAL, 0);
	const deg = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI + 90;
	s.addShape('triangle', { x: x2 - 0.075, y: y2 - 0.075, w: 0.15, h: 0.15, rotate: deg, fill: { color: CHARCOAL }, line: { type: 'none' } });
}

/** Sheet of paper: white card ruled with pale blue lines. */
function sheet (s, x, y, w, h, rot) {
	s.addShape('rect', { x, y, w, h, rotate: rot, fill: { color: PAPER }, line: { color: PAPER_BLUE, width: 0.75 } });
	for (let i = 1; i <= 3; i++) {
		s.addShape('rect', {
			x: x + w * 0.15, y: y + (h * i) / 4, w: w * 0.7, h: 0.018, rotate: rot,
			fill: { color: PAPER_BLUE }, line: { type: 'none' }
		});
	}
}

/** Flat-style archer illustration standing on a rule (left half of slide 12). */
function archerIllustration (s) {
	oval(s, 1.1666, 2.0725, 4.3898, 2.3928, PAPER);
	s.addShape('ellipse', {
		x: 1.55, y: 2.15, w: 3.55, h: 2.15,
		fill: { type: 'none' }, line: { color: '9AA3B5', width: 0.75, dashType: 'sysDot' }
	});

	[[1.1888, 2.3615, 0.39, 0.49, -14], [2.4023, 1.5621, 0.43, 0.47, 12],
		[2.193, 3.1532, 0.42, 0.46, -18], [4.5997, 3.6971, 0.34, 0.63, 16],
		[4.6184, 1.7865, 0.42, 0.49, 10]].forEach(p => sheet(s, p[0], p[1], p[2], p[3], p[4]));

	bullseye(s, 1.7008, 2.0359, 0.8595, 0.8595, ORANGE);
	bullseye(s, 3.6393, 1.9307, 0.8049, 0.8049, YELLOW);
	bullseye(s, 1.1864, 3.2441, 0.5739, 1.1629, RED);
	bullseye(s, 4.9788, 2.2046, 0.7218, 1.4624, PINK);

	// four arms, drawn behind the body: two raised, two reaching sideways
	[[3.10, 3.35, 2.42, 2.52], [3.50, 3.35, 4.10, 2.40],
		[2.95, 3.60, 1.72, 3.72], [3.65, 3.62, 4.72, 3.02]].forEach(a => limb(s, a[0], a[1], a[2], a[3], 0.185, SKIN));
	arrow(s, 3.32, 2.79, 2.14, 2.44);
	arrow(s, 3.32, 2.30, 4.34, 2.42);
	arrow(s, 2.72, 3.44, 1.50, 3.72);
	arrow(s, 3.90, 2.68, 4.92, 2.94);

	// shoulders / torso, three white shirt stripes, neck, head and hair
	s.addShape('roundRect', { x: 2.83, y: 3.30, w: 1.06, h: 1.16, rectRadius: 0.26, fill: { color: SKIN }, line: { type: 'none' } });
	[3.5963, 3.7374, 3.8787].forEach(y => {
		s.addShape('rect', { x: 2.8769, y, w: 0.971, h: 0.10, fill: { color: WHITE }, line: { type: 'none' } });
	});
	s.addShape('rect', { x: 3.22, y: 3.16, w: 0.18, h: 0.22, fill: { color: SKIN }, line: { type: 'none' } });
	oval(s, 3.0762, 2.7329, 0.4561, 0.5704, SKIN);
	s.addShape('roundRect', { x: 3.055, y: 2.5573, w: 0.505, h: 0.2723, rectRadius: 0.125, fill: { color: CHARCOAL }, line: { type: 'none' } });
	dot(s, 3.34, 2.90, 0.045, CHARCOAL);

	dot(s, 4.3429, 3.8073, 0.099, PEACH);
	dot(s, 2.3997, 4.1281, 0.099, RED);
	dot(s, 1.5863, 3.02, 0.099, ORANGE);
	dot(s, 2.9454, 1.9868, 0.099, PINK);
	s.addShape('rect', { x: 0.7216, y: 4.4582, w: 5.054, h: 0.014, fill: { color: CHARCOAL }, line: { type: 'none' } });
}

const STRATEGY_ROWS = [
	{ y: 1.3817, pct: '70%', name: 'Strategy 01', fill: RED, text: WHITE, body: WHITE },
	{ y: 3.0599, pct: '84%', name: 'Strategy 02', fill: ORANGE, text: WHITE, body: WHITE },
	{ y: 4.7382, pct: '91%', name: 'Strategy 03', fill: YELLOW, text: BLACK, body: GREY }
];
const STRATEGY_BODY = 'Lorem ipsum dolor sit amet  sed do eiusmod tempor incididunt ut';

function slide12 (pptx) {
	const s = pptx.addSlide();
	header(s, false);
	archerIllustration(s);

	heading(s, [['Infographic ', BLACK], ['Section', RED]], { x: 0.7511, y: 4.9747, w: 4.6639, h: 0.6395 });
	txt(s, 'Lorem ipsum dolor sit amet  sed do eiusmod tempor incididunt ut labore et dolore magna aliquaLorem ipsum dolor sit amet  sed do eiusmod tempor incididunt ut labore et dolore magna aliqua', {
		x: 0.7511, y: 5.8066, w: 5.825, h: 0.8695, color: GREY_MID, lineSpacingMultiple: 1.3
	});

	s.addShape('line', {
		x: 10.0238, y: 1.595, w: 0, h: 5.905,
		line: { color: BLACK, width: 1, dashType: 'dash' }
	});

	STRATEGY_ROWS.forEach(r => {
		card(s, 7.381, r.y, 5.2857, 1.3801, r.fill, 0.04011, false);
		txt(s, r.pct, { x: 7.5042, y: r.y + 0.2411, w: 2.0076, h: 0.9189, fontSize: 54, color: r.text, lineSpacingMultiple: 0.9 });
		txt(s, r.name, { x: 9.4784, y: r.y + 0.2133, w: 1.8884, h: 0.3702, fontSize: 16, color: r.text });
		txt(s, STRATEGY_BODY, { x: 9.4784, y: r.y + 0.5802, w: 2.9502, h: 0.6075, color: r.body, lineSpacingMultiple: 1.3 });
	});
}

/* -------------------------------------------------------------- slide  13 */

const REWARD_STATS = [
	{ numX: 1.1047, numY: 2.7538, num: '120+', labX: 2.4028, labY: 2.7612, title: 'Product', bodyY: 3.1315, bodyW: 2.1599, rewY: 3.3318 },
	{ numX: 5.0145, numY: 2.7765, num: '80+', labX: 6.3126, labY: 2.784, title: 'Market', bodyY: 3.1542, bodyW: 2.1718, rewY: 3.3546 }
];
const REWARD_CARDS = [
	{ x: 1.2181, y: 4.1974, amount: '$31m', label: 'Management', labelW: 1.9901 },
	{ x: 4.9375, y: 4.1924, amount: '$22m', label: 'Investor Relations', labelW: 1.9963 }
];
const PROBLEM_BODY = 'PLACEHOLDER';

function slide13 (pptx) {
	const s = pptx.addSlide();
	header(s, false);

	card(s, 9.298, 2.1231, 3.265, 4.3421, RED, 0.08353, false);
	[2.9261, 4.3283, 5.0454].forEach(y => {
		s.addShape('line', { x: 9.5675, y, w: 2.7261, h: 0, line: { color: WHITE, width: 0.5, transparency: 61 } });
	});
	txt(s, 'Problem', { x: 9.4588, y: 2.3753, w: 2.9435, h: 0.4039, align: 'center', fontSize: 18, bold: true, color: WHITE });
	txt(s, PROBLEM_BODY, { x: 9.496, y: 3.0929, w: 2.869, h: 1.0569, align: 'center', color: GREY_LIGHT, lineSpacingMultiple: 1.2 });
	txt(s, 'Solutions', { x: 9.8089, y: 4.4856, w: 2.2432, h: 0.4039, align: 'center', fontSize: 18, bold: true, color: WHITE });
	txt(s, PROBLEM_BODY, { x: 9.496, y: 5.1886, w: 2.869, h: 1.0569, align: 'center', color: GREY_LIGHT, lineSpacingMultiple: 1.2 });

	heading(s, [['Infographic ', BLACK], ['Section', RED]], { x: 1.1667, y: 1.7672, w: 4.6639, h: 0.6395 });

	REWARD_STATS.forEach(r => {
		txt(s, r.num, { x: r.numX, y: r.numY, w: 1.4164, h: 0.7068, fontSize: 36, bold: true, color: BLACK });
		txt(s, 'Rewards', { x: r.numX, y: r.rewY, w: 1.4164, h: 0.3702, fontSize: 16, color: INK });
		txt(s, r.title, { x: r.labX, y: r.labY, w: 2.5415, h: 0.3702, fontSize: 16, color: INK });
		txt(s, 'Lorem dolor sit amet consectetur adipiscing', {
			x: r.labX, y: r.bodyY, w: r.bodyW, h: 0.5722, color: GREY_MID, lineSpacingMultiple: 1.2
		});
	});

	REWARD_CARDS.forEach(c => {
		card(s, c.x, c.y, 3.1528, 1.3375, WHITE, 0.16667, true);
		txt(s, c.amount, { x: c.x + 0.2377, y: c.y + 0.2674, w: 1.6456, h: 0.5722, fontSize: 28, color: INK });
		txt(s, c.label, { x: c.x + 0.2377, y: c.y + 0.7671, w: c.labelW, h: 0.3299, color: GREY_MID, lineSpacingMultiple: 1.2 });
	});
	apertureIcon(s, 3.71, 4.91, 0.72, RED);
	targetArrowIcon(s, 7.42, 4.99, 0.66, ORANGE);

	txt(s, 'PLACEHOLDER', {
		x: 1.1391, y: 6.0065, w: 6.9512, h: 0.8145, color: GREY_MID, lineSpacingMultiple: 1.2
	});
}

/* -------------------------------------------------------------- slide  14 */

function slide14 (pptx) {
	const s = pptx.addSlide();
	decoRings(s, -3.3432, 3.75, 6.973);
	imagePlaceholder(s, 0.8901, 1.7565, 4.1724, 4.9935, 0.2762);
	header(s, false);

	heading(s, [['Get Our Contact ', BLACK], ['Information\u2019s', RED]], { x: 5.6302, y: 1.665, w: 7.1927, h: 0.6395 });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', {
		x: 5.6302, y: 2.4289, w: 7.1927, h: 0.9812, color: GREY, lineSpacingMultiple: 1.5
	});

	dot(s, 5.7863, 4.1277, 0.324, RED);
	clockIcon(s, 5.9483, 4.2897, WHITE);
	txt(s, 'Office Hours', { x: 6.1686, y: 3.9776, w: 2.4562, h: 0.4667, fontFace: MAJOR, fontSize: 16, lineSpacingMultiple: 1.5 });
	txt(s, [
		{ text: 'Monday \u2013 Saturday', options: { breakLine: true } },
		{ text: '08.00 AM \u2013 08.00 PM' }
	], { x: 6.1686, y: 4.4146, w: 2.4562, h: 0.6732, color: GREY_MID, lineSpacingMultiple: 1.5 });

	card(s, 8.8132, 4.0331, 2.967, 1.1102, RED, 0.13645, false);
	txt(s, 'Address', { x: 9.1048, y: 4.1062, w: 2.4562, h: 0.4667, fontSize: 16, color: WHITE, lineSpacingMultiple: 1.5 });
	txt(s, '123 Street Name, City Name', { x: 9.1048, y: 4.5432, w: 2.4562, h: 0.3753, color: WHITE, lineSpacingMultiple: 1.5 });

	handsetIcon(s, 5.9283, 5.6638, ORANGE);
	txt(s, 'Get In Touch', { x: 6.1289, y: 5.369, w: 2.4562, h: 0.4667, fontFace: MAJOR, fontSize: 16, lineSpacingMultiple: 1.5 });
	txt(s, [
		{ text: '(+62) 000 0000 0000', options: { breakLine: true } },
		{ text: '(0725) 00000' }
	], { x: 6.1289, y: 5.8061, w: 2.4562, h: 0.6732, color: GREY_MID, lineSpacingMultiple: 1.5 });

	shareIcon(s, 8.8, 5.679, YELLOW);
	txt(s, 'Follow Us', { x: 9.065, y: 5.369, w: 2.4562, h: 0.4667, fontFace: MAJOR, fontSize: 16, lineSpacingMultiple: 1.5 });
	txt(s, [
		{ text: 'www.yoursitehere.com', options: { breakLine: true } },
		{ text: 'office@ yoursitehere.com' }
	], { x: 9.065, y: 5.8061, w: 2.4562, h: 0.6782, color: GREY_MID, lineSpacingMultiple: 1.5 });
}

/* -------------------------------------------------------------- slide  15 */

function slide15 (pptx) {
	const s = pptx.addSlide();
	s.background = { color: RED };

	cornerArcs(s, 13.3333, 7.5, [3.126, 2.558, 1.990], 'E03636');
	[[8.0753, -4.6984, 9.5714], [-0.7693, 6.4525, 1.9151], [5.6735, -1.1164, 1.9151]].forEach(([x, y, d]) => {
		s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: WHITE, transparency: 85 }, line: { type: 'none' } });
	});

	header(s, true);

	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed', {
		x: 0.4948, y: 1.5697, w: 6.1719, h: 0.6782, color: GREY_LIGHT, lineSpacingMultiple: 1.5
	});
	s.addShape('roundRect', {
		x: 0.4948, y: 2.4957, w: 2.5838, h: 0.4921, rectRadius: 0.246,
		fill: { color: WHITE, transparency: 85 }, line: { type: 'none' }
	});
	txt(s, 'End Presentation', { x: 0.6484, y: 2.5566, w: 2.2766, h: 0.3702, align: 'center', fontFace: MAJOR, fontSize: 16, color: WHITE });

	txt(s, 'Strategic Business Alignment Meeting', { x: 0.6024, y: 4.35, w: 4.8693, h: 0.4039, fontSize: 18, color: WHITE });
	txt(s, 'Thank You', {
		x: 0.6024, y: 5.0858, w: 12.0573, h: 2.1174,
		fontFace: MAJOR, fontSize: 158, color: WHITE, lineSpacingMultiple: 0.7
	});
}

/* --------------------------------------------------------------------- run */

function build () {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333, height: 7.5 });
	pptx.layout = 'WIDE_16x9';
	pptx.author = 'Starlaight Corp';
	pptx.title = 'Meeting Plan';
	pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };

	[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
		slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach(fn => fn(pptx));

	return pptx.writeFile({
		fileName: path.join(__dirname, '0f567ad5-13d6-40f9-bcd4-d3d037e4779a_grok_final.pptx')
	});
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
