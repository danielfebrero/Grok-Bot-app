#!/usr/bin/env node
/**
 * "Brand Agency Studio" brand-proposal deck (20 slides, 13.333 x 7.5 in)
 * rebuilt with pptxgenjs only.
 *
 * Photographs in the original are reproduced as flat placeholder rectangles
 * (`imageFrame`) at the exact position/size of the original picture frames.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme: "Brand Proposal" colour scheme + "Custom 3" font scheme
 * ------------------------------------------------------------------ */

const HEAD = 'Playfair Display SemiBold'; // theme major latin font
const BODY = 'Inter'; // theme minor latin font

const C = {
	page: 'FFFFFF', // bg1
	ink: '262626', // tx1  lumMod 85%  – headings & labels
	body: '404040', // tx1  lumMod 75%  – running copy
	rule: 'D7C8BC', // accent3          – header rule
	brand: 'C45A47', // accent1          – filled panels, bullets
	clay: 'B86A55', // accent2          – oversized statistics
	sand: 'CFCAC4', // accent4          – table rules, chart line
	hair: 'ECEAE7', // accent4 lum 40%  – hairline separators
	chip: 'EBEAE7', // accent5          – button chip
	grid: 'D9D9D9', // tx1  lumMod 15%  – chart gridlines
	axis: '595959', // tx1  lumMod 65%  – chart axis labels
	white: 'FFFFFF',
};

/* Every picture frame in the deck carries the same tx1 @ 10% alpha wash. */
const FRAME_FILL = { color: '000000', transparency: 90 };

/* ------------------------------------------------------------------ *
 * Small helpers – every one of these is a thin wrapper over addText /
 * addShape so positions and styles stay visible at the call site.
 * ------------------------------------------------------------------ */

/** Text box. `p` = [x, y, w, h] in inches. */
function text(slide, body, p, o = {}) {
	slide.addText(body, {
		x: p[0],
		y: p[1],
		w: p[2],
		h: p[3],
		fontFace: o.font || BODY,
		fontSize: o.size || 14,
		color: o.color || C.body,
		bold: o.bold || false,
		italic: o.italic || false,
		align: o.align || 'left',
		valign: o.valign || 'top',
		lineSpacingMultiple: o.line,
		paraSpaceAfter: o.after,
		wrap: true,
	});
}

/** Playfair display heading (section titles are 40 pt unless overridden). */
function heading(slide, body, p, o = {}) {
	text(slide, body, p, { font: HEAD, size: 40, color: C.ink, after: 12, ...o });
}

/** Small Playfair caps label used above paragraphs of copy. */
function label(slide, body, p, o = {}) {
	text(slide, body, p, { font: HEAD, size: 16, color: C.ink, line: 1.3, after: 12, ...o });
}

/** Body copy: Inter 14 pt, 130% leading, 12 pt space after. */
function copy(slide, body, p, o = {}) {
	text(slide, body, p, { size: 14, color: C.body, line: 1.3, after: 12, ...o });
}

/** Oversized clay-coloured statistic (88+, 95%, 01/02/03 ...). */
function stat(slide, body, p, o = {}) {
	text(slide, body, p, { font: HEAD, size: 80, color: C.clay, line: 1.3, after: 12, ...o });
}

/** Placeholder standing in for a photograph. */
function imageFrame(slide, x, y, w, h) {
	slide.addShape('rect', { x, y, w, h, fill: FRAME_FILL, line: { type: 'none' } });
}

/** Solid accent panel. */
function panel(slide, x, y, w, h, color = C.brand) {
	slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

function dot(slide, x, y, d, color = C.brand) {
	slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' } });
}

function rule(slide, x, y, w, h, color, pt) {
	slide.addShape('line', { x, y, w, h, line: { color, width: pt } });
}

/* ------------------------------------------------------------------ *
 * Slide master chrome: hairline, two running heads, three social icons
 * ------------------------------------------------------------------ */

function chrome(slide) {
	rule(slide, 0.667, 0.755, 12.0, 0, C.rule, 0.5);
	text(slide, 'Brand Agency Studio', [0.667, 0.312, 2.855, 0.303], {
		font: HEAD,
		size: 12,
		color: C.ink,
	});
	text(slide, 'Branding & Creative Development', [4.988, 0.312, 3.355, 0.303], {
		font: HEAD,
		size: 12,
		color: C.ink,
		align: 'center',
	});
	socialIcons(slide);
}

/* The three logo glyphs are line art in the original; redrawn here from
 * native shapes rather than embedded raster/SVG data. */
function socialIcons(slide) {
	const stroke = { color: C.ink, width: 1.25 };
	const hair = { color: C.ink, width: 0.75 };
	const outline = (kind, x, y, w, h, extra = {}) =>
		slide.addShape(kind, { x, y, w, h, fill: { type: 'none' }, line: stroke, ...extra });

	// Instagram — rounded square, lens, flash dot
	outline('roundRect', 11.441, 0.33, 0.267, 0.267, { rectRadius: 0.07 });
	outline('ellipse', 11.518, 0.407, 0.113, 0.113);
	slide.addShape('ellipse', { x: 11.638, y: 0.372, w: 0.036, h: 0.036, fill: { color: C.ink } });

	// LinkedIn — "in" wordmark drawn as glyph strokes: i-dot, i-stem, n-shoulder
	const bar = (x, y, w, h) => slide.addShape('rect', { x, y, w, h, fill: { color: C.ink } });
	bar(11.914, 0.334, 0.05, 0.05);
	bar(11.914, 0.404, 0.05, 0.168);
	bar(11.996, 0.404, 0.05, 0.168);
	bar(12.128, 0.436, 0.05, 0.136);
	bar(11.996, 0.404, 0.182, 0.05);

	// Dribbble — ball crossed by two elliptical seams
	outline('ellipse', 12.443, 0.368, 0.224, 0.224);
	outline('ellipse', 12.491, 0.368, 0.128, 0.224, { line: hair, rotate: 35 });
	outline('ellipse', 12.443, 0.416, 0.224, 0.128, { line: hair, rotate: -20 });
}

/* ------------------------------------------------------------------ *
 * Shared copy — the template repeats the same "wonderful serenity"
 * lorem passage throughout, so it is named once.
 * ------------------------------------------------------------------ */

const L = {
	short: 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings',
	soul: 'A wonderful serenity has taken possession of my entire soul',
	spring:
		'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my heart.',
	alone:
		'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine.',
	happy:
		'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence.',
	stroke: 'PLACEHOLDER',
};

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Cover
function slide01(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	imageFrame(s, 0.994, 1.05, 11.345, 3.569);
	text(s, 'BRAND', [5.369, 2.906, 6.531, 1.717], {
		font: HEAD,
		size: 96,
		bold: true,
		italic: true,
		color: C.white,
		align: 'right',
		valign: 'bottom',
		after: 12,
	});
	heading(s, 'EXPERIENCE', [0.863, 4.869, 9.929, 1.717], { size: 96 });
}

// 2 — Break / company value journey
function slide02(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BREAK', [0.82, 3.094, 3.93, 1.313], { size: 72, bold: true });
	heading(s, 'COMPANY\u2019S VALUE JOURNEY ', [8.821, 2.004, 3.625, 1.919], {
		size: 36,
		bold: true,
		valign: 'bottom',
	});
	copy(s, L.spring, [8.821, 4.2, 3.625, 1.297]);
	imageFrame(s, 5.057, 1.108, 3.219, 5.285);
}

// 3 — Break slide
function slide03(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(
		s,
		['BREAK', 'SLIDE'].map((t) => ({ text: t, options: { breakLine: true } })),
		[6.844, 3.723, 5.822, 3.016],
		{ size: 96, bold: true, align: 'right', line: 0.85 }
	);
	copy(s, L.short, [1.136, 5.403, 3.543, 0.808], { line: undefined, after: undefined });
	imageFrame(s, 1.231, 1.289, 4.103, 3.574);
}

// 4 — Brand overview
function slide04(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND OVERVIEW', [0.82, 0.993, 5.847, 0.774]);
	copy(
		s,
		[L.spring, L.alone, L.happy, L.stroke].map((t) => ({ text: t, options: { breakLine: true } })),
		[0.881, 2.255, 3.807, 4.252]
	);
	imageFrame(s, 5.4, 2.359, 2.358, 4.391);
	imageFrame(s, 8.022, 2.359, 2.358, 3.359);
	imageFrame(s, 8.846, 4.359, 3.821, 2.391);
}

// 5 — Inside the brand
function slide05(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'INSIDE THE BRAND', [0.82, 0.993, 6.303, 0.774]);
	copy(s, L.spring, [0.881, 2.432, 4.678, 0.99]);
	copy(
		s,
		[
			'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.',
			'I am alone, and feel the charm of existence in this spot. ',
			L.happy,
		].map((t) => ({ text: t, options: { breakLine: true } })),
		[7.536, 1.2, 4.678, 2.246]
	);
	imageFrame(s, 0.974, 3.75, 6.381, 3.0);
	imageFrame(s, 7.631, 3.75, 4.728, 3.0);
}

// 6 — Brand visual / concept
function slide06(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	panel(s, 5.19, 4.752, 2.953, 2.744);
	heading(s, 'BRAND VISUAL', [2.485, 1.01, 5.077, 0.774]);
	heading(s, 'CONCEPT', [0.85, 4.506, 3.461, 0.774], { bold: true });
	copy(s, L.short, [9.749, 4.752, 2.662, 1.297], { align: 'center' });
	copy(s, L.short, [0.877, 5.571, 3.637, 0.99]);
	imageFrame(s, 0.0, 1.308, 2.282, 2.744);
	imageFrame(s, 4.374, 2.679, 2.538, 3.192);
	imageFrame(s, 7.175, 2.679, 2.282, 3.628);
	imageFrame(s, 8.256, 1.308, 5.077, 2.744);
}

// 7 — The minds behind the brand
function slide07(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	panel(s, 10.0, 2.383, 2.667, 2.734);
	heading(s, 'THE MINDS BEHIND THE BRAND', [0.82, 0.993, 4.765, 2.121]);
	copy(s, `${L.spring} ${L.alone} ${L.happy}`, [0.858, 4.291, 4.765, 2.215]);
	imageFrame(s, 6.085, 0.75, 4.422, 2.693);
	imageFrame(s, 6.085, 3.712, 2.624, 3.038);
	imageFrame(s, 9.0, 4.436, 4.333, 2.314);
}

// 8 — Brand visual gallery
function slide08(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND VISUAL GALLERY', [0.82, 2.69, 3.321, 2.121]);
	copy(
		s,
		'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.\u00a0',
		[7.193, 1.314, 5.374, 0.99]
	);
	imageFrame(s, 2.767, 5.6, 4.233, 1.9);
	imageFrame(s, 4.333, 1.436, 2.667, 3.899);
	imageFrame(s, 7.253, 2.744, 5.414, 4.756);
}

// 9 — Brand project / 88+
function slide09(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND PROJECT', [0.82, 0.993, 4.147, 1.447]);
	stat(s, '88+', [0.853, 2.908, 4.348, 1.667], { bold: true });
	label(s, 'COMPANY PORTFOLIO', [0.893, 4.757, 4.308, 0.417]);
	copy(s, `${L.spring} `, [0.9, 5.269, 4.308, 0.99]);
	imageFrame(s, 7.311, 1.107, 2.078, 2.873);
	imageFrame(s, 9.66, 2.156, 1.873, 1.824);
	imageFrame(s, 5.667, 4.244, 3.722, 3.256);
	imageFrame(s, 9.66, 4.245, 3.007, 2.169);
}

// 10 — Key brand features (numbered list beside a tall photo)
function slide10(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'KEY BRAND FEATURES', [5.107, 0.964, 7.482, 0.774], { after: 6 });

	// The first label carries 130% leading, the other two are set solid.
	const rows = [
		{ n: '01', nx: 4.928, ny: 2.178, x: 6.43, y: 2.499, lh: 0.417, lead: 1.3, title: 'OVERVIEW' },
		{ n: '02', nx: 4.961, ny: 3.646, x: 6.424, y: 3.967, lh: 0.37, title: 'OTHERS ELEMENTS' },
		{ n: '03', nx: 4.952, ny: 5.114, x: 6.424, y: 5.435, lh: 0.37, title: 'PROPOSAL IDEA' },
	];
	rows.forEach((r) => {
		stat(s, r.n, [r.nx, r.ny, 1.298, 1.041], { size: 48, align: 'center' });
		label(s, r.title, [r.x, r.y, 5.346, r.lh], { line: r.lead, after: r.lead && 12 });
		copy(s, L.short, [r.x, r.y + 0.417, 5.346, 0.684]);
	});

	imageFrame(s, 0.667, 0.75, 3.769, 6.75);
}

// 11 — Who we are (zig-zag numbered list)
function slide11(pptx) {
	const s = pptx.addSlide();
	chrome(s);

	const rows = [
		{ n: '01', nx: 7.035, ny: 1.053, tx: 8.772, ty: 1.393, align: 'right' },
		{ n: '02', nx: 11.095, ny: 3.075, tx: 7.272, ty: 3.419, align: 'left' },
		{ n: '03', nx: 7.059, ny: 5.102, tx: 8.772, ty: 5.457, align: 'right' },
	];
	rows.forEach((r) => {
		stat(s, r.n, [r.nx, r.ny, 1.298, 1.041], { size: 48, align: 'center' });
		copy(s, L.short, [r.tx, r.ty, 3.442, 0.99], { align: r.align });
	});

	rule(s, 7.369, 2.943, 4.738, 0, C.hair, 0.75);
	rule(s, 7.369, 4.95, 4.738, 0, C.hair, 0.75);

	heading(s, 'WHO WE ARE', [1.036, 4.644, 5.333, 0.774]);
	copy(s, L.short, [1.023, 5.705, 5.346, 0.684]);
	imageFrame(s, 1.128, 1.179, 5.538, 3.205);
}

// 12 — Core capabilities (bulleted cards)
function slide12(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'CORE CAPABILITIES', [1.07, 0.978, 4.984, 1.447], { bold: true });
	copy(s, 'A wonderful serenity has taken possession of my entire soul.', [1.11, 2.649, 4.944, 0.572], {
		line: undefined,
		after: undefined,
	});

	const rows = [
		{ y: 1.001, x: 7.489, cx: 7.5, title: 'SIGNATURE STRENGTH' },
		{ y: 2.963, x: 7.489, cx: 7.5, title: 'STRATEGIC ADVANTAGE' },
		{ y: 4.925, x: 7.494, cx: 7.505, title: 'DEFINED EXCELLENCE' },
	];
	rows.forEach((r) => {
		dot(s, 7.033, r.y + 0.066, 0.284);
		label(s, r.title, [r.x, r.y, 4.792, 0.417]);
		copy(s, `${L.spring} `, [r.cx, r.y + 0.584, 4.984, 0.99]);
	});

	imageFrame(s, 1.217, 3.75, 4.623, 3.0);
}

// 13 — How we work (full-bleed accent panel)
function slide13(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	panel(s, 6.022, 1.367, 7.311, 6.133);
	heading(s, 'HOW WE WORK TO DELIVER EXCEPTIONAL RESULTS', [0.82, 2.353, 5.003, 2.794]);
	copy(
		s,
		[
			`${L.spring} ${L.alone}\u00a0`,
			'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents.',
			'I should be incapable of drawing a single stroke at the present moment and yet I feel that I never was a greater artist than now.',
		].map((t) => ({ text: t, options: { breakLine: true } })),
		[7.034, 2.545, 5.288, 3.777],
		{ color: C.white }
	);
}

// 14 — Executive summary (three labelled blocks)
function slide14(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'EXECUTIVE SUMMARY', [0.919, 1.127, 5.206, 1.447]);

	label(s, 'SUMMARY OF GOALS', [0.941, 3.514, 4.992, 0.417]);
	copy(s, `${L.spring} ${L.alone} ${L.happy}`, [0.949, 4.157, 4.984, 2.215]);

	label(s, 'QUICK SNAPSHOT OF THE BRAND VISION', [7.407, 1.292, 5.006, 0.414]);
	copy(s, `${L.spring} ${L.alone} `, [7.429, 1.875, 4.984, 1.603]);

	label(s, 'INTRODUCTION TO THE BRAND CONCEPT', [7.42, 4.186, 4.994, 0.414]);
	copy(s, `${L.spring} ${L.alone} `, [7.431, 4.77, 4.984, 1.603]);
}

// 15 — Smoothed line chart + 95% callout
function slide15(pptx) {
	const s = pptx.addSlide();
	chrome(s);

	s.addChart(
		pptx.ChartType.line,
		[
			{ name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [3.9, 2.4, 3.2, 1] },
			{ name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [1, 3.3, 1.8, 4.5] },
			{ name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [1, 2, 4.2, 3.2] },
		],
		{
			x: 0.679,
			y: 2.756,
			w: 7.41,
			h: 3.994,
			chartColors: [C.sand, C.brand, C.rule],
			lineSize: 2.25,
			lineSmooth: true,
			lineDataSymbol: 'none',
			showLegend: false,
			catAxisHidden: true,
			valAxisMaxVal: 5,
			valAxisMajorUnit: 1,
			valAxisLineShow: false,
			valAxisLabelColor: C.axis,
			valAxisLabelFontFace: BODY,
			valAxisLabelFontSize: 12,
			valGridLine: { color: C.grid, size: 0.75 },
		}
	);

	heading(s, 'UNDERSTANDING OUR BRAND DIRECTION', [0.691, 1.127, 11.951, 0.774], { align: 'center' });
	stat(s, '95%', [8.611, 3.496, 2.938, 1.667], { bold: true });
	copy(s, L.soul, [8.8, 5.326, 3.743, 0.684]);
}

// 16 — Cost breakdown table
function slide16(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND PROPOSAL COST BREAKDOWN', [0.691, 1.127, 11.951, 0.774], {
		bold: true,
		align: 'center',
	});

	const head = ['DATE', 'NAME', 'QUANTITY', 'PRICE', 'TOTAL'];
	const body = [
		['05/02/2025', 'Brand Research & Analysis', '1', '$500', '$500'],
		// The date cell of row 2 carries a trailing empty paragraph in the source,
		// which lifts its text slightly above the rest of the row.
		[['07/02/2025', ''], 'Logo & Visual Identity Design', '3', '$700', '$2,100'],
		['09/02/2025', 'Brand Guideline Development', '1', '$1,200', '$1,200'],
		['15/02/2025', 'Marketing Collateral Design', '4', '$350', '$1,400'],
	];
	const none = { type: 'none' };
	const under = [none, none, { pt: 1, color: C.sand }, none];
	const paras = (v) =>
		Array.isArray(v) ? v.map((t) => ({ text: t, options: { breakLine: true } })) : v;

	const rows = [
		head.map((t) => ({ text: t, options: { fontFace: HEAD, fontSize: 16, color: C.ink, border: under } })),
	];
	body.forEach((r, i) => {
		const border = i === body.length - 1 ? [none, none, none, none] : under;
		rows.push(
			r.map((v) => ({ text: paras(v), options: { fontFace: BODY, fontSize: 14, color: C.body, border } }))
		);
	});

	s.addTable(rows, {
		x: 0.679,
		y: 2.702,
		w: 11.976,
		colW: [2.678, 3.083, 2.214, 2.048, 1.952],
		rowH: 0.81,
		align: 'center',
		valign: 'middle',
	});
}

// 17 — Bar chart + summary card
function slide17(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND PERFORMANCE CHART', [1.244, 1.127, 10.846, 0.774], { align: 'center' });

	const years = ['2020', '2022', '2023', '2024', '2025'];
	s.addChart(
		pptx.ChartType.bar,
		[
			{ name: 'Series 1', labels: years, values: [2, 3, 2, 3, 3.4] },
			{ name: 'Series 2', labels: years, values: [2.4, 4, 3.4, 2, 1.5] },
			// Third (empty) series: present in the source workbook and it is what
			// leaves the visible gap on the right of every cluster.
			{ name: 'Column1', labels: years, values: [] },
		],
		{
			x: 4.59,
			y: 2.333,
			w: 7.718,
			h: 4.417,
			layout: { x: 0.021825510408241575, y: 0, w: 0.95453624921185265, h: 0.90718158544076377 },
			chartColors: [C.brand, C.clay, C.rule],
			barGapWidthPct: 48,
			barOverlapPct: -7,
			showLegend: false,
			valAxisHidden: true,
			valGridLine: { style: 'none' },
			catAxisLineColor: C.grid,
			catAxisLineSize: 0.75,
			catAxisLabelColor: C.ink,
			catAxisLabelFontFace: BODY,
			catAxisLabelFontSize: 12,
		}
	);

	// Left summary card
	panel(s, 1.255, 2.711, 2.478, 4.039);
	text(s, 'BRAND HERE', [1.345, 2.964, 2.298, 0.308], {
		font: HEAD,
		color: C.white,
		align: 'center',
	});
	rule(s, 1.597, 3.399, 1.794, 0, C.white, 0.25);
	text(s, '$65,955', [1.151, 3.48, 2.686, 0.808], {
		font: HEAD,
		size: 40,
		bold: true,
		color: C.white,
		align: 'center',
		line: 1.3,
		after: 12,
	});
	text(s, 'Monthly Reach', [1.345, 4.281, 2.298, 0.215], { size: 8, color: C.white, align: 'center' });
	text(s, 'A wonderful serenity has taken', [1.52, 4.639, 1.948, 0.601], {
		size: 12,
		color: C.white,
		align: 'center',
		line: 1.3,
		after: 12,
	});
	text(s, 'Text Here', [1.462, 5.348, 0.799, 0.215], { font: HEAD, size: 8, color: C.white });
	text(s, 'Text Here', [2.883, 5.348, 0.799, 0.236], { font: HEAD, size: 8, color: C.white });
	text(s, '30+', [1.454, 5.576, 0.592, 0.345], { color: C.white, line: 1.3, after: 12 });
	text(s, '15K', [3.049, 5.576, 0.592, 0.345], { color: C.white, line: 1.3, after: 12 });

	s.addShape('rect', {
		x: 1.55,
		y: 6.117,
		w: 1.888,
		h: 0.38,
		fill: { color: C.chip },
		line: { type: 'none' },
	});
	text(s, 'VIEW DETAILS', [1.55, 6.117, 1.888, 0.38], {
		font: HEAD,
		color: C.ink,
		align: 'center',
		valign: 'middle',
	});
}

// 18 — Performance infographic (two stat panels)
function slide18(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND PERFORMANCE INFOGRAPHIC', [0.754, 1.127, 11.826, 0.774], { align: 'center' });

	const cards = [
		{ px: 1.062, tx: 1.519, cx: 1.521, vx: 1.227, title: 'BRAND AWARENESS GROWTH', value: '85,4%' },
		{ px: 4.733, tx: 5.117, cx: 5.139, vx: 4.899, title: 'CUSTOMER ENGAGEMENT RATE', value: '94,8%' },
	];
	cards.forEach((c) => {
		panel(s, c.px, 2.821, 3.385, 3.667);
		label(s, c.title, [c.tx, 3.156, 2.977, 0.766], { color: C.white });
		copy(s, 'A wonderful serenity has taken possession', [c.cx, 4.039, 2.807, 0.684], { color: C.white });
		text(s, c.value, [c.vx, 4.877, 2.807, 1.276], {
			font: HEAD,
			size: 60,
			bold: true,
			color: C.white,
			align: 'center',
			line: 1.3,
			after: 12,
		});
	});

	copy(
		s,
		[L.spring, L.alone].map((t) => ({ text: t, options: { breakLine: true } })),
		[8.524, 3.463, 3.748, 2.384]
	);
}

// 19 — Brand development timeline
function slide19(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'BRAND DEVELOPMENT', [1.349, 1.115, 6.384, 1.447]);

	rule(s, 8.405, 1.232, 0, 4.685, C.hair, 0.75);

	const steps = [
		{ year: '2021/22', yx: 8.76, yy: 1.047, dx: 8.299, dy: 1.218, tx: 8.759, ty: 1.471, tw: 3.225 },
		{ year: '2022/23', yx: 8.76, yy: 2.575, dx: 8.299, dy: 2.742, tx: 8.759, ty: 3.017, tw: 3.225 },
		{ year: '2023/24', yx: 8.749, yy: 4.1, dx: 8.288, dy: 4.267, tx: 8.748, ty: 4.542, tw: 3.225 },
		{ year: '2024/25', yx: 8.76, yy: 5.625, dx: 8.299, dy: 5.791, tx: 8.759, ty: 6.066, tw: 3.214 },
	];
	steps.forEach((p) => {
		text(s, p.year, [p.yx, p.yy, 2.293, 0.453], {
			font: HEAD,
			size: 18,
			color: C.clay,
			valign: 'middle',
			line: 1.3,
			after: 12,
		});
		dot(s, p.dx, p.dy, 0.211);
		copy(s, L.soul, [p.tx, p.ty, p.tw, 0.684]);
	});

	imageFrame(s, 1.506, 3.231, 5.629, 3.519);
}

// 20 — Contact
function slide20(pptx) {
	const s = pptx.addSlide();
	chrome(s);
	heading(s, 'CONTACT', [1.349, 1.115, 4.573, 0.774], { after: undefined });
	text(
		s,
		'WE CREATE SHOW STOPPING BRANDS FOR BUSINESSES THAT WANT TO LEAD THE PACK.',
		[2.117, 2.536, 3.611, 0.909],
		{ font: HEAD, size: 16, color: C.ink }
	);

	contactIcons(s);
	const contacts = [
		{ y: 2.636, w: 2.261, value: '038 \u2013 378 \u2013 3273' },
		{ y: 3.752, w: 4.108, value: '2345 Street Name, City Name' },
		{ y: 4.881, w: 2.261, value: '@yourbrand' },
		{ y: 5.949, w: 2.825, value: 'www.brand-example.com' },
	];
	contacts.forEach((c) => copy(s, c.value, [8.275, c.y, c.w, 0.378]));

	imageFrame(s, 2.25, 3.75, 3.38, 3.0);
}

/* Phone handset, address book, instagram and globe glyphs, drawn from
 * native shapes at the sizes of the original SVG graphics. */
function contactIcons(slide) {
	const stroke = { color: C.ink, width: 1.1 };
	const outline = (kind, x, y, w, h, extra = {}) =>
		slide.addShape(kind, { x, y, w, h, fill: { type: 'none' }, line: stroke, ...extra });

	// Phone — handset as an open ring cut away at the top right
	outline('blockArc', 7.67, 2.62, 0.4, 0.4, { rotate: 215, adj: [10800000, 17500000, 20000] });

	// Address book — front card with a person bust, second card tucked behind
	outline('roundRect', 7.7, 3.77, 0.36, 0.43, { rectRadius: 0.06 });
	outline('roundRect', 7.637, 3.752, 0.36, 0.43, { rectRadius: 0.06 });
	outline('ellipse', 7.777, 3.855, 0.088, 0.088);
	outline('arc', 7.745, 3.968, 0.152, 0.14, { rotate: 180 });

	// Instagram — rounded square, lens, flash dot
	outline('roundRect', 7.713, 4.912, 0.316, 0.316, { rectRadius: 0.08 });
	outline('ellipse', 7.8, 4.999, 0.142, 0.142);
	slide.addShape('ellipse', { x: 7.945, y: 4.955, w: 0.04, h: 0.04, fill: { color: C.ink } });

	// Globe — sphere crossed by two meridians and two parallels
	outline('ellipse', 7.653, 5.921, 0.435, 0.435);
	outline('ellipse', 7.79, 5.921, 0.162, 0.435);
	slide.addShape('line', { x: 7.653, y: 6.138, w: 0.435, h: 0, line: stroke });
	slide.addShape('line', { x: 7.69, y: 6.032, w: 0.36, h: 0, line: stroke });
	slide.addShape('line', { x: 7.69, y: 6.245, w: 0.36, h: 0, line: stroke });
}

/* ------------------------------------------------------------------ */

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'BRAND_16x9', width: 13.3333333, height: 7.5 });
	pptx.layout = 'BRAND_16x9';
	pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
	pptx.title = 'Brand Agency Studio — Branding & Creative Development';

	[
		slide01, slide02, slide03, slide04, slide05,
		slide06, slide07, slide08, slide09, slide10,
		slide11, slide12, slide13, slide14, slide15,
		slide16, slide17, slide18, slide19, slide20,
	].forEach((fn) => fn(pptx));

	return pptx;
}

build().writeFile({
	fileName: path.join(__dirname, '0e5cb4c7-9602-4e2b-9ec0-b753e2d37e7a_grok_final.pptx'),
});
