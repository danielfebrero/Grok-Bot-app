/*
 * "Agenda Planner" deck — rebuilt with pptxgenjs.
 *
 *   node 0d83c6c8-9383-4bee-a468-b8e1c09bd367_grok_final.js
 *
 * 16 slides, 13.333 x 7.5 in (16:9).  Raster artwork in the source deck is
 * replaced by native pptxgenjs shapes (see `monitorMock`, `meetingScene`).
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette */

const C = {
	blue: '318DF7', blueDk: '0868D6', blueTint: 'D6E8FD',
	yellow: 'FFD630', yellowDk: 'E3B600', yellowTint: 'FFF7D6',
	pink: 'F05691', pinkDk: 'E11462', pinkTint: 'FCDDE9',
	purple: '390249', purpleLt: '5A0373', purpleTint: 'F5D5FE',
	// the deck fills most panels with a 2-stop gradient; pptxgenjs is solid-fill
	// only, so each gradient collapses to the average of its two stops
	blueMid: '2380E9', yellowMid: 'F1C618', pinkMid: 'E93579', purpleMid: '4A025E',
	white: 'FFFFFF', black: '000000',
	grey: '808080',          // tx1 lum 50% — all the muted body copy
	greyLine: 'E7E6E6',      // lt2 — hairlines, ghost cards
	greyDk: 'AFABAB',
	gridLine: 'D9D9D9',
};

const HEAD = 'Sora';   // theme major font
const BODY = 'Mulish'; // theme minor font

/* Shadows lifted from the original effectLst entries (blur/dist in pt).
   pptxgenjs rewrites the object it is handed, so every shape gets a copy. */
const SHADOWS = {
	card: { color: C.black, opacity: 0.10, blur: 20, offset: 3, angle: 135 },
	soft: { color: C.black, opacity: 0.10, blur: 15, offset: 3, angle: 45 },
	tag: { color: C.black, opacity: 0.30, blur: 15, offset: 3, angle: 135 },
	badge: { color: C.black, opacity: 0.15, blur: 20, offset: 3, angle: 135 },
	hero: { color: C.black, opacity: 0.15, blur: 25, offset: 10, angle: 135 },
	ring: { color: C.black, opacity: 0.20, blur: 35, offset: 29, angle: 135 },
	left: { color: C.black, opacity: 0.10, blur: 35, offset: 3, angle: 180 },
};
const sh = name => Object.assign({ type: 'outer' }, SHADOWS[name]);

/* ------------------------------------------------------------ primitives */

const R = { def: 0.16667, icon: 0.19424 };   // PowerPoint default + the deck's icon-tile adj
const rect = (s, x, y, w, h, o) => s.addShape('rect', Object.assign({ x, y, w, h }, o));
const ellipse = (s, x, y, w, h, o) => s.addShape('ellipse', Object.assign({ x, y, w, h }, o));

/* `adj` is the OOXML corner adjustment: a fraction of the shape's short side
   (PowerPoint's own default is 0.16667).  pptxgenjs wants it in inches. */
const rrect = (s, x, y, w, h, adj, o) =>
	s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: adj * Math.min(w, h) }, o));

/** Text box.  Defaults mirror the deck: Mulish 18pt, black, top aligned. */
function txt(s, body, o) {
	s.addText(body, Object.assign({ fontFace: BODY, fontSize: 18, color: C.black, valign: 'top' }, o));
}

/** Multi-coloured headline: rich([['Important ', C.blue], ['Deadlines', null]], 40). */
function rich(parts, size, face) {
	return parts.map(([text, color]) => ({
		text,
		options: { color: color || C.black, fontSize: size, fontFace: face || HEAD },
	}));
}

/** "Page N" marker that sits in the bottom-right corner of every slide. */
function pageNum(s, n) {
	txt(s, 'Page ' + n, { x: 11.553, y: 6.826, w: 1.419, h: 0.337, fontSize: 14, align: 'right' });
}

/** Hollow ring — the decorative "Circle: Hollow" donuts of the cover slides. */
function ring(s, x, y, d) {
	const t = d * 0.09898;   // ring thickness from the source donut's adj value
	ellipse(s, x + t / 2, y + t / 2, d - t, d - t,
		{ fill: { type: 'none' }, line: { color: C.blue, width: t * 72 }, shadow: sh('ring') });
}

/* ------------------------------------------------------------- glyph set */
/* Small line icons rebuilt from primitives; `d` is the glyph box size. */

function glyph(s, kind, x, y, d, color) {
	const u = d / 12, at = (a, b, w, h, o) => rect(s, x + a * u, y + b * u, w * u, h * u, o);
	const fill = { color }, stroke = { fill: { type: 'none' }, line: { color, width: 0.75 } };
	if (kind === 'calendar') {
		rrect(s, x + u, y + 2 * u, 10 * u, 9 * u, 0.08, stroke);
		at(1, 2, 10, 2.4, { fill }); at(3, 0.8, 0.9, 2, { fill }); at(8.1, 0.8, 0.9, 2, { fill });
		[0, 1, 2].forEach(c => [0, 1].forEach(r => at(2.4 + c * 2.7, 5.6 + r * 2.2, 1.6, 1.2, { fill })));
	} else if (kind === 'form') {
		rrect(s, x + 2 * u, y + u, 8 * u, 10 * u, 0.08, stroke);
		[0, 1, 2].forEach(i => at(3.4, 3.2 + i * 2.4, 5.2, 0.9, { fill }));
		at(4.5, 0.2, 3, 1.6, { fill });
	} else if (kind === 'target') {
		s.addShape('donut', { x: x + u, y: y + u, w: 10 * u, h: 10 * u, fill });
		s.addShape('donut', { x: x + 3.6 * u, y: y + 3.6 * u, w: 4.8 * u, h: 4.8 * u, fill });
	} else if (kind === 'clock') {
		s.addShape('donut', { x: x + 0.5 * u, y: y + 0.5 * u, w: 11 * u, h: 11 * u, fill });
		at(5.6, 3, 0.8, 3.5, { fill }); at(6, 5.6, 3, 0.8, { fill });
	} else if (kind === 'check') {
		s.addShape('custGeom', {
			x, y, w: d, h: d, line: { color, width: 1.6 },
			points: [{ x: 0.22 * d, y: 0.52 * d }, { x: 0.42 * d, y: 0.72 * d }, { x: 0.8 * d, y: 0.28 * d }],
		});
	} else if (kind === 'cart') {
		s.addShape('custGeom', {
			x, y, w: d, h: d, line: { color, width: 1.2 },
			points: [{ x: 0.1 * d, y: 0.2 * d }, { x: 0.26 * d, y: 0.2 * d }, { x: 0.4 * d, y: 0.66 * d }, { x: 0.86 * d, y: 0.66 * d }, { x: 0.94 * d, y: 0.32 * d }, { x: 0.3 * d, y: 0.32 * d }],
		});
		ellipse(s, x + 0.4 * d, y + 0.76 * d, 0.13 * d, 0.13 * d, { fill });
		ellipse(s, x + 0.74 * d, y + 0.76 * d, 0.13 * d, 0.13 * d, { fill });
	} else if (kind === 'gear') {
		s.addShape('gear6', { x, y, w: d, h: d, fill });
	} else if (kind === 'cloud') {
		s.addShape('cloud', { x, y: y + 0.15 * d, w: d, h: 0.7 * d, fill });
	} else if (kind === 'case') {
		rrect(s, x + 0.5 * u, y + 3 * u, 11 * u, 7 * u, 0.08, { fill });
		rrect(s, x + 4 * u, y + u, 4 * u, 2.4 * u, 0.1, { fill: { type: 'none' }, line: { color, width: 1.2 } });
	} else if (kind === 'people') {
		ellipse(s, x + 3.5 * u, y + 1.5 * u, 5 * u, 5 * u, { fill });
		s.addShape('round2SameRect', { x: x + 1.5 * u, y: y + 7 * u, w: 9 * u, h: 4 * u, fill });
	} else if (kind === 'phone') {
		s.addShape('custGeom', {
			x, y, w: d, h: d, fill,
			points: [{ x: 0.1 * d, y: 0.14 * d }, { x: 0.34 * d, y: 0.1 * d }, { x: 0.46 * d, y: 0.4 * d }, { x: 0.3 * d, y: 0.52 * d }, { x: 0.5 * d, y: 0.78 * d }, { x: 0.66 * d, y: 0.64 * d }, { x: 0.92 * d, y: 0.78 * d }, { x: 0.84 * d, y: 0.96 * d }, { x: 0.2 * d, y: 0.62 * d }, { close: true }],
		});
	} else if (kind === 'mail') {
		rrect(s, x, y + 0.2 * d, d, 0.6 * d, 0.08, { fill: { type: 'none' }, line: { color, width: 1 } });
		s.addShape('custGeom', {
			x, y: y + 0.2 * d, w: d, h: 0.6 * d, line: { color, width: 1 },
			points: [{ x: 0, y: 0 }, { x: 0.5 * d, y: 0.36 * d }, { x: d, y: 0 }],
		});
	} else if (kind === 'globe') {
		s.addShape('donut', { x, y, w: d, h: d, fill });
		ellipse(s, x + 0.32 * d, y, 0.36 * d, d, { fill: { type: 'none' }, line: { color, width: 0.75 } });
		rect(s, x, y + 0.46 * d, d, 0.08 * d, { fill });
	} else if (kind === 'arrow') {
		s.addShape('rightArrow', { x, y: y + 0.28 * d, w: d, h: 0.44 * d, fill });
	}
}

/** Colour chip (rounded square) with a white glyph on top. */
function iconChip(s, x, y, size, chipColor, kind, glyphColor) {
	rrect(s, x, y, size, size, R.def, { fill: { color: chipColor }, shadow: sh('badge') });
	glyph(s, kind, x + size * 0.22, y + size * 0.22, size * 0.56, glyphColor || C.white);
}

/* ============================================================== SLIDE  1 */

function slide01(s) {
	txt(s, 'Conversion Funnel.', { x: 0.278, y: 0.159, w: 1.944, h: 0.337, fontSize: 14 });
	txt(s, 'www.agendaplanner.com', { x: 5.194, y: 0.159, w: 2.944, h: 0.337, fontSize: 14, color: C.purple, align: 'center' });
	hamburger(s);

	ring(s, 0.361, 5.928, 5.077);
	ring(s, 9.015, -2.379, 5.077);
	ring(s, -1.632, 3.665, 2.743);
	ring(s, 12.577, 2.698, 1.446);

	txt(s, 'Agenda', { x: 2.379, y: 1.989, w: 6.569, h: 2.036, fontSize: 115, fontFace: HEAD, color: C.blue, wrap: false });
	txt(s, 'Planner', { x: 4.618, y: 3.475, w: 6.336, h: 2.036, fontSize: 115, fontFace: HEAD, color: C.blue, wrap: false });

	rrect(s, 4.487, 3.468, 3.44, 0.523, 0.5, { fill: { color: C.white }, rotate: 1, shadow: sh('hero') });
	txt(s, 'Presentation Template', { x: 4.622, y: 3.528, w: 2.842, h: 0.404, rotate: 1, wrap: false });
	ellipse(s, 7.467, 3.569, 0.362, 0.362, { fill: { type: 'none' }, line: { color: C.black, width: 1.25 } });
	glyph(s, 'arrow', 7.545, 3.648, 0.205, C.black);

	pageNum(s, 1);
}

/** Three short blue rules — the "menu" mark in the top-right of the covers. */
function hamburger(s) {
	[[12.750, 0.258, 0.181], [12.709, 0.331, 0.263], [12.750, 0.397, 0.181]].forEach(([x, y, w]) =>
		s.addShape('line', { x, y, w, h: 0, line: { color: C.blue, width: 2 } }));
}

/* ============================================================== SLIDE  2 */

function slide02(s) {
	txt(s, rich([
		['Welcome to Project Planning, ', C.blue], ['structured approach that ', null],
		['organizes, ', C.blue], ['every stage of a project. ', null],
	], 32), { x: 0.754, y: 0.623, w: 5.96, h: 2.255 });

	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
		{ x: 0.754, y: 3.097, w: 5.254, h: 0.572, fontSize: 14, color: C.grey, align: 'justify' });

	const cards = [
		{ x: 0.754, num: '15', color: C.blue, week: 'Week 1', icon: 'target' },
		{ x: 3.381, num: '29', color: C.yellow, week: 'Week 2', icon: 'calendar' },
		{ x: 6.008, num: '7', color: C.pink, week: 'Week 2', icon: 'form' },
	];
	cards.forEach(c => {
		rrect(s, c.x, 4.42, 2.203, 2.457, R.def, { fill: { color: C.white }, shadow: sh('card') });
		txt(s, 'Agenda Date', { x: c.x + 0.295, y: 4.706, w: 1.623, h: 0.337, fontSize: 14, align: 'center' });
		txt(s, c.num, { x: c.x + 0.29, y: 4.925, w: 1.623, h: 1.447, fontSize: 80, fontFace: HEAD, color: c.color, align: 'center' });
		txt(s, c.week, { x: c.x + 0.295, y: 6.254, w: 1.623, h: 0.337, fontSize: 14, align: 'center' });
		iconChip(s, c.x + 0.407, 5.787, 0.392, c.color, c.icon);
	});

	pageNum(s, 2);
}

/* ============================================================== SLIDE  3 */

function slide03(s) {
	[2.333, 4.270, 6.210].forEach(y =>
		s.addShape('line', { x: 0, y, w: 10.857, h: 0, line: { color: C.greyLine, width: 2.25, dashType: 'sysDash' } }));

	rect(s, 0, 0, 11.476, 0.799, { fill: { color: C.white }, shadow: sh('card') });
	s.addShape('round1Rect', {
		x: 11.246, y: 0, w: 2.087, h: 1.083, flipH: true, flipV: true,
		fill: { color: C.purpleMid }, shadow: sh('card'),
	});
	txt(s, 'December', { x: 11.470, y: 0.076, w: 1.752, h: 0.438, fontSize: 20, fontFace: HEAD, color: C.white });
	txt(s, '2030', { x: 11.252, y: 0.435, w: 1.677, h: 0.572, fontSize: 28, fontFace: HEAD, color: C.white, align: 'center' });

	// day strip
	[['12', 0.368], ['13', 1.724], ['14', 3.048], ['15', 4.367], ['16', 5.687], ['17', 6.978], ['18', 8.275], ['19', 9.600]]
		.forEach(([day, x]) => {
			txt(s, 'Mon', { x, y: 0.215, w: 0.58, h: 0.303, fontSize: 12, color: C.grey });
			txt(s, day, { x: x + 0.379, y: 0.113, w: 0.62, h: 0.572, fontSize: 28, fontFace: HEAD, wrap: false });
		});

	// vertical time chips on the right rail
	[['12:00', 1.643], ['13:00', 3.579], ['14:00', 5.520]].forEach(([label, y]) => {
		rrect(s, 11.246, y, 0.46, 1.381, R.def, { fill: { color: C.purple } });
		txt(s, label, { x: 11.257, y: y + 0.181, w: 0.438, h: 1.019, fontSize: 14, fontFace: HEAD, color: C.white, align: 'center', vert: 'vert' });
	});

	// three highlighted agenda blocks (coloured card + purple stub on its left)
	const blocks = [
		{ y: 1.643, cx: 7.951, sx: 6.863, tx: 6.630, label: 'Agenda 1', color: C.pinkMid },
		{ y: 3.611, cx: 1.885, sx: 0.798, tx: 0.564, label: 'Agenda 2', color: C.yellowMid },
		{ y: 5.575, cx: 5.356, sx: 4.269, tx: 4.035, label: 'Agenda 3', color: C.blueMid },
	];
	blocks.forEach(b => {
		s.addShape('round2SameRect', { x: b.sx, y: b.y + 0.078, w: 0.706, h: 1.468, rotate: 270, fill: { color: C.purpleMid }, shadow: sh('card') });
		txt(s, 'Lorem ipsum dolor sit amet', { x: b.tx, y: b.y + 0.576, w: 1.277, h: 0.471, fontSize: 11, color: C.white });
		rrect(s, b.cx, b.y, 2.547, 1.325, 0.072, { fill: { color: b.color } });
		txt(s, b.label, { x: b.cx + 0.218, y: b.y + 0.174, w: 1.534, h: 0.438, fontSize: 20, fontFace: HEAD, color: C.white });
		txt(s, 'Lorem ipsum dolor sit amet', { x: b.cx + 0.223, y: b.y + 0.579, w: 2.037, h: 0.572, fontSize: 14, color: C.white });
	});

	// three empty white agenda cards
	[[0.857, 1.667], [4.947, 3.639], [8.193, 5.572]].forEach(([x, y]) => {
		rrect(s, x, y, 2.095, 1.325, 0.079, { fill: { color: C.white }, shadow: sh('card') });
		txt(s, 'Agenda #', { x: x + 0.213, y: y + 0.192, w: 1.669, h: 0.438, fontSize: 20, fontFace: HEAD });
		txt(s, 'Lorem ipsum dolor sit amet.', { x: x + 0.213, y: y + 0.561, w: 1.669, h: 0.572, fontSize: 14, color: C.grey });
	});

	pageNum(s, 3);
}

/* ============================================================== SLIDE  4 */

function slide04(s) {
	txt(s, rich([['Important ', C.blue], ['Deadlines and ', null], ['Key', C.blue], [' Reminders', null]], 40),
		{ x: 0.633, y: 0.809, w: 4.452, h: 2.121 });

	// timeline rail: pale full-height rule with a solid blue segment on top
	s.addShape('line', { x: 0.849, y: 3.179, w: 0, h: 3.466, line: { color: C.blueTint, width: 3 } });
	s.addShape('line', { x: 0.850, y: 3.556, w: 0, h: 2.006, line: { color: C.blue, width: 3 } });

	const items = [
		{ ix: 1.186, iy: 3.364, tx: 1.880, ty: 3.233, title: 'Market', body: 'Global demand for poultry is rising, now.', color: C.yellow, tint: C.yellowTint, icon: 'calendar' },
		{ ix: 1.220, iy: 4.541, tx: 1.915, ty: 4.471, title: 'Trends', body: 'Consumers prefer healthier, sustainable.', color: C.blue, tint: C.blueTint, icon: 'form' },
		{ ix: 1.199, iy: 5.775, tx: 1.893, ty: 5.705, title: 'Trends', body: 'Consumers prefer healthier, sustainable.', color: C.pink, tint: C.pinkTint, icon: 'target' },
	];
	items.forEach(it => {
		rrect(s, it.ix, it.iy, 0.614, 0.614, R.def, { fill: { color: it.tint } });
		rrect(s, it.ix + 0.086, it.iy + 0.086, 0.443, 0.443, R.def, { fill: { color: it.color } });
		glyph(s, it.icon, it.ix + 0.146, it.iy + 0.146, 0.323, C.white);
		txt(s, it.title, { x: it.tx, y: it.ty, w: 2.499, h: 0.438, fontSize: 20, fontFace: HEAD, align: 'justify' });
		txt(s, it.body, { x: it.tx + 0.010, y: it.ty + 0.378, w: 2.499, h: 0.609, fontSize: 14, color: C.grey, align: 'justify', lineSpacingMultiple: 1.1 });
	});

	// big December "13" calendar card
	rrect(s, 8.265, 1.893, 3.946, 3.946, 0.05, { fill: { color: C.white }, shadow: sh('soft') });
	s.addShape('round2SameRect', { x: 8.265, y: 1.893, w: 3.946, h: 0.908, fill: { color: C.yellowMid }, shadow: sh('card') });
	txt(s, 'December', { x: 8.510, y: 2.061, w: 3.455, h: 0.572, fontSize: 28, fontFace: HEAD, align: 'center' });
	txt(s, '13', { x: 8.694, y: 2.699, w: 3.087, h: 3.45, fontSize: 199, fontFace: HEAD, align: 'center', wrap: false });

	// floating pink note + blue ribbon
	rrect(s, 7.694, 4.152, 1.0, 1.0, R.def, { fill: { color: C.white }, shadow: sh('card') });
	rrect(s, 7.853, 4.313, 0.683, 0.683, R.icon, { fill: { color: C.pink }, shadow: sh('card') });
	glyph(s, 'form', 7.945, 4.405, 0.5, C.white);
	glyph(s, 'people', 6.487, 5.842, 0.295, C.black);

	rrect(s, 9.791, 5.657, 2.804, 0.537, 0.117, { fill: { color: C.blueMid }, rotate: -6, shadow: sh('soft') });
	txt(s, 'Important Agenda!', { x: 9.878, y: 5.723, w: 2.628, h: 0.404, rotate: -6, color: C.white, fontFace: HEAD, align: 'center' });

	pageNum(s, 4);
}

/* ============================================================== SLIDE  5 */

function slide05(s) {
	txt(s, rich([['Detailed Agenda', C.blue], [' for the ', null], ['Day', C.blue]], 40),
		{ x: 0.778, y: 0.507, w: 5.010, h: 1.447 });

	// two check-list rows
	[{ y: 2.055, on: false }, { y: 2.744, on: true }].forEach(row => {
		rect(s, 0.880, row.y + 0.193, 0.324, 0.324, { fill: { color: row.on ? C.blueMid : C.greyLine } });
		glyph(s, 'check', 0.951, row.y + 0.264, 0.181, row.on ? C.white : C.greyDk);
		txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem ',
			{ x: 1.373, y: row.y + 0.100, w: 3.828, h: 0.572, fontSize: 14, color: row.on ? C.black : C.grey });
	});

	// three stacked schedule cards sliding off the right edge
	const cards = [
		{ y: -0.331, time: '10:00', color: C.yellow, dark: false, tab: C.yellow },
		{ y: 1.628, time: '11:00', color: C.pink, dark: false, tab: C.pink },
		{ y: 3.587, time: '12:00', color: C.white, dark: true, tab: C.purpleTint },
	];
	cards.forEach(c => {
		const x = c.dark ? 5.961 : 7.257, w = c.dark ? 7.372 : 6.076, h = c.dark ? 1.851 : 1.526;
		s.addShape('round1Rect', { x, y: c.y, w, h, rectRadius: 0.155, flipH: true, flipV: true,
			fill: { color: c.dark ? C.purpleMid : C.white }, shadow: c.dark ? undefined : sh('card') });
		s.addShape('round2SameRect', { x: c.dark ? 6.311 : 7.545, y: c.y, w: c.dark ? 2.136 : 1.760, h: c.dark ? 0.183 : 0.151, flipV: true, fill: { color: c.tab } });
		txt(s, c.time, { x: c.dark ? 6.311 : 7.401, y: c.y + (c.dark ? 0.44 : 0.384), w: c.dark ? 2.516 : 2.362, h: c.dark ? 1.010 : 0.909,
			fontSize: c.dark ? 54 : 48, fontFace: HEAD, color: c.dark ? C.white : c.color, align: 'center' });
		txt(s, 'Agenda 1', { x: c.dark ? 8.988 : 9.752, y: c.y + (c.dark ? 0.276 : 0.227), w: c.dark ? 4.042 : 3.332, h: c.dark ? 0.505 : 0.438,
			fontSize: c.dark ? 24 : 20, fontFace: HEAD, color: c.dark ? C.white : C.black });
		txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem ',
			{ x: c.dark ? 8.988 : 9.752, y: c.y + (c.dark ? 0.800 : 0.660), w: c.dark ? 4.042 : 3.332, h: c.dark ? 0.707 : 0.640,
				fontSize: c.dark ? 18 : 16, color: c.dark ? C.white : C.grey });
	});

	pageNum(s, 5);
}

/* ============================================================== SLIDE  6 */

function slide06(s) {
	txt(s, rich([['Special Guest ', C.blue], ['Speaker Sessions and Highlights', null]], 40),
		{ x: 8.167, y: 0.829, w: 4.292, h: 2.794 });

	// four speaker slots, each labelled with a name ribbon
	const people = [
		{ x: 0.875, name: 'Larasha', color: C.purpleLt, ny: 3.483 },
		{ x: 4.399, name: 'Jooshica', color: C.blue, ny: 2.223 },
		{ x: 4.399, name: 'Queensha', color: C.pink, ny: 5.962 },
		{ x: 0.875, name: 'Annastasia', color: C.yellow, ny: 7.180 },
	];
	people.forEach(p => {
		s.addShape('round2SameRect', { x: p.x + 0.492, y: p.ny - 0.131, w: 0.357, h: 1.218, rotate: 90, fill: { color: p.color } });
		rrect(s, p.x + 0.260, p.ny, 1.687, 0.410, 0.194, { fill: { color: p.color }, shadow: sh('tag') });
		txt(s, p.name, { x: p.x + 0.400, y: p.ny + 0.020, w: 1.407, h: 0.370, fontSize: 16, color: p.color === C.yellow ? C.purple : C.white, align: 'center' });
	});

	// 75% stat card
	s.addShape('round2SameRect', { x: 10.205, y: 2.598, w: 1.492, h: 4.764, rotate: 270, fill: { color: C.white }, shadow: sh('card') });
	s.addShape('round2SameRect', { x: 8.878, y: 4.234, w: 1.760, h: 0.117, flipV: true, fill: { color: C.blue } });
	txt(s, '75%', { x: 8.878, y: 4.470, w: 2.021, h: 1.111, fontSize: 60, fontFace: HEAD, color: C.blue });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.',
		{ x: 10.900, y: 4.576, w: 2.264, h: 0.808, fontSize: 14, color: C.grey });

	// Join Now button + arrow chip
	rrect(s, 8.369, 6.240, 1.673, 0.601, 0.206, { fill: { color: C.purple } });
	txt(s, 'Join Now', { x: 8.467, y: 6.339, w: 1.448, h: 0.404, color: C.white, fontFace: HEAD, align: 'center' });
	rrect(s, 10.140, 6.240, 0.601, 0.601, 0.187, { fill: { color: C.purple }, line: { color: C.white, width: 1, dashType: 'sysDash' } });
	glyph(s, 'arrow', 10.239, 6.342, 0.4, C.white);

	pageNum(s, 6);
}

/* ============================================================== SLIDE  7 */

function slide07(s) {
	s.addShape('round1Rect', { x: 5.129, y: 0, w: 8.204, h: 5.889, rectRadius: 0.54, flipH: true, flipV: true, fill: { color: C.purpleMid } });
	monitorMock(s);

	txt(s, rich([['Preview', C.yellow], [' of Upcoming ', C.white], ['Events and Initiatives', C.yellow]], 44),
		{ x: 6.667, y: 0.489, w: 5.747, h: 2.322, align: 'right' });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
		{ x: 7.059, y: 3.009, w: 5.355, h: 0.337, fontSize: 14, color: C.white, align: 'right' });

	[{ x: 7.583, num: '15', color: C.blue, chip: C.pink, icon: 'form', ix: 7.127 },
	 { x: 11.384, num: '23', color: C.yellow, chip: C.purple, icon: 'calendar', ix: 10.928 }].forEach(card => {
		rrect(s, card.x, 3.863, 3.016, 3.016, 0.109, { fill: { color: C.white }, shadow: sh('card') });
		txt(s, 'Important', { x: card.x + 0.154, y: 4.044, w: 2.709, h: 0.337, fontSize: 14, fontFace: HEAD, color: C.grey, align: 'center' });
		txt(s, 'Upcoming Event', { x: card.x + 0.154, y: 4.276, w: 2.709, h: 0.438, fontSize: 20, fontFace: HEAD, align: 'center' });
		txt(s, card.num, { x: card.x + 0.32, y: 4.495, w: 2.36, h: 2.423, fontSize: 138, fontFace: HEAD, color: card.color, align: 'center' });
		rrect(s, card.ix, 5.497, 0.912, 0.912, R.def, { fill: { color: C.white }, shadow: sh('card') });
		rrect(s, card.ix + 0.145, 5.643, 0.623, 0.623, R.icon, { fill: { color: card.chip }, shadow: sh('card') });
		glyph(s, card.icon, card.ix + 0.245, 5.743, 0.42, C.white);
	});

	pageNum(s, 7);
}

/** Desktop-monitor mock-up standing in for the photo in the source deck.
    The photo is shot at an angle, so the panels are drawn as quads. */
function monitorMock(s) {
	const quad = (pts, color) => {
		const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
		const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
		s.addShape('custGeom', {
			x, y, w: Math.max.apply(null, xs) - x, h: Math.max.apply(null, ys) - y,
			points: pts.map(p => ({ x: p[0] - x, y: p[1] - y })).concat([{ close: true }]),
			fill: { color },
		});
	};
	quad([[0.47, 0.39], [5.98, 1.34], [6.88, 4.86], [1.28, 4.86]], '2E2E2E');   // bezel
	quad([[0.78, 0.72], [5.78, 1.58], [6.55, 4.55], [1.44, 4.55]], 'F2F2F2');   // screen
	quad([[1.28, 4.86], [6.88, 4.86], [6.98, 5.46], [1.45, 5.62]], 'CFCFCF');   // chin
	quad([[2.42, 5.60], [3.66, 5.60], [3.72, 6.55], [2.30, 6.55]], 'C4C4C4');   // neck
	quad([[2.18, 6.42], [4.55, 6.42], [5.05, 6.90], [1.72, 6.90]], 'DADADA');   // foot
	txt(s, '[image]', { x: 1.5, y: 2.4, w: 4.5, h: 0.4, fontSize: 12, color: C.greyDk, align: 'center' });
}

/* ============================================================== SLIDE  8 */

function slide08(s) {
	txt(s, rich([['Program Flow ', C.blue], ['and Discussions', null]], 40),
		{ x: 1.664, y: 0.630, w: 10.005, h: 0.774, align: 'center' });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ',
		{ x: 3.067, y: 1.404, w: 7.198, h: 0.572, fontSize: 14, align: 'center' });

	// year strip: ghost cells with two highlighted rounded cards
	[0.000, 4.778, 6.917, 11.672].forEach(x =>
		rect(s, x, 2.56, 2.139, 0.73, { line: { color: C.greyLine, width: 1 } }));
	[[0.095, 0.912], [4.873, 5.690], [7.011, 7.829], [11.766, 12.584]].forEach(([tx, bx], i) => {
		const y = 2.70;
		txt(s, '2025', { x: tx, y, w: 0.95, h: 0.438, fontSize: 20, fontFace: HEAD, color: C.greyLine });
		txt(s, 'Lorem ipsum dolor sit amet.', { x: bx, y, w: 1.177, h: 0.419, fontSize: 10.5, color: C.greyLine, lineSpacingMultiple: 0.9 });
	});
	[2.139, 9.033].forEach(x => {
		rrect(s, x, 2.458, 2.639, 0.906, R.def, { fill: { color: C.white }, line: { color: C.greyLine, width: 1 }, shadow: sh('card') });
		txt(s, '2025', { x: x + 0.137, y: 2.634, w: 1.219, h: 0.572, fontSize: 28, fontFace: HEAD, color: C.blue });
		txt(s, 'Lorem ipsum dolor sit amet.', { x: x + 1.253, y: 2.688, w: 1.34, h: 0.464, fontSize: 12, lineSpacingMultiple: 0.9 });
	});

	// clustered bar chart (native chart, same data / colours as the original)
	rrect(s, 2.139, 3.684, 6.064, 3.187, 0.079, { fill: { color: C.white }, shadow: sh('card') });
	const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
	s.addChart('bar', [
		{ name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
	], {
		x: 2.409, y: 3.919, w: 5.523, h: 2.705,
		barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
		chartColors: [C.blue, C.yellow, C.pink],
		showLegend: false, showTitle: false,
		catAxisLabelFontSize: 12, catAxisLabelFontFace: BODY, catAxisLineShow: true,
		valAxisLabelFontSize: 12, valAxisLabelFontFace: BODY, valAxisLineShow: false,
		valGridLine: { style: 'solid', color: C.gridLine, size: 0.75 },
		catGridLine: { style: 'none' },
	});

	// "567+" callout pinned over the chart
	rrect(s, 6.002, 4.081, 0.800, 0.442, 0.158, { fill: { color: C.white }, rotate: -6, shadow: sh('card') });
	txt(s, '567+', { x: 5.950, y: 4.189, w: 0.772, h: 0.337, fontSize: 14, fontFace: HEAD, rotate: -6, align: 'center' });
	rrect(s, 6.476, 3.992, 0.214, 0.214, 0.158, { fill: { color: C.pink }, rotate: -6, shadow: sh('card') });
	s.addShape('arc', { x: 6.720, y: 4.176, w: 0.860, h: 0.475, rotate: 8, line: { color: C.grey, width: 2.25, dashType: 'sysDash' } });

	// 85% purple stat panel
	rrect(s, 9.006, 3.684, 3.402, 3.187, 0.088, { fill: { color: C.purpleMid }, shadow: sh('card') });
	txt(s, '85%', { x: 9.358, y: 5.017, w: 2.540, h: 1.107, fontSize: 72, fontFace: HEAD, color: C.white, lineSpacingMultiple: 0.8 });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
		{ x: 9.358, y: 5.929, w: 2.689, h: 0.572, fontSize: 14, color: C.white });
	rrect(s, 11.696, 3.970, 1.0, 1.0, R.def, { fill: { color: C.white }, shadow: sh('card') });
	rrect(s, 11.855, 4.131, 0.683, 0.683, R.icon, { fill: { color: C.purple }, shadow: sh('card') });
	glyph(s, 'calendar', 11.955, 4.231, 0.48, C.white);

	pageNum(s, 8);
}

/* ============================================================== SLIDE  9 */

const PLOT = { x: 1.19, y: 2.435, w: 11.164, h: 3.825 };   // 50k at top → -10k at bottom
const toY = v => PLOT.y + (50000 - v) / 60000 * PLOT.h;

/** Sum of gaussian humps — the smooth silhouettes used by the source artwork. */
const wave = (t, peaks) => peaks.reduce((a, [c, amp, wd]) => a + amp * Math.exp(-((t - c) ** 2) / (2 * wd * wd)), 0);

function trace(peaks, samples) {
	const pts = [];
	for (let i = 0; i <= samples; i++) {
		const t = i / samples;
		pts.push({ x: PLOT.w * t, y: toY(wave(t, peaks)) - PLOT.y });
	}
	return pts;
}

function slide09(s) {
	txt(s, rich([['Tracking Progress ', C.blue], ['Through ', null], ['Project ', C.blue], ['Milestoneshart', null]], 36),
		{ x: 0.782, y: 0.487, w: 10.836, h: 1.313 });

	// range selector
	rrect(s, 9.554, 1.544, 1.459, 0.516, R.def, { fill: { color: C.white }, line: { color: C.greyLine, width: 1 }, shadow: sh('card') });
	[['12 months', 6.852, 1.452, C.grey, 11], ['30 days', 8.308, 1.250, C.grey, 11],
	 ['1 weeks', 9.554, 1.459, C.blue, 14], ['24 hours', 11.023, 1.331, C.grey, 11]].forEach(([label, x, w, color, size], i) => {
		if (i !== 2) rect(s, x, 1.592, w, 0.395, { fill: { color: C.white }, line: { color: C.greyLine, width: 0.75 } });
		txt(s, label, { x, y: 1.64, w, h: 0.3, fontSize: size, color, align: 'center' });
	});

	// two filled silhouettes, faded towards the baseline, then the trend line
	const pink = [[0.190, 31500, 0.055], [0.47, 8000, 0.11], [0.775, 25400, 0.05]];
	const yellow = [[0.180, 13800, 0.062], [0.47, 3500, 0.11], [0.780, 10500, 0.055]];
	const blue = [[0.200, 17800, 0.052], [0.47, 8500, 0.075], [0.780, 11500, 0.058]];
	area(s, pink, C.pink);
	area(s, yellow, C.yellow);
	whiteVeil(s);

	// gridlines + axis labels sit on top of the fill, as in the original
	[['50k', 2.435], ['40k', 3.067], ['30k', 3.711], ['20k', 4.344], ['10k', 4.999], ['0', 5.631]].forEach(([label, y]) => {
		s.addShape('line', { x: 1.19, y: y + 0.143, w: 11.164, h: 0, line: { color: C.gridLine, width: 0.75 } });
		txt(s, label, { x: 0.714, y, w: 0.467, h: 0.286, fontSize: 11, color: C.grey, align: 'center', wrap: false });
	});
	['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'].forEach((d, i) =>
		txt(s, d, { x: 0.90 + i * 1.744, y: 6.482, w: 1.7, h: 0.286, fontSize: 11, color: C.grey, align: 'center', wrap: false }));

	s.addShape('custGeom', {
		x: PLOT.x, y: PLOT.y, w: PLOT.w, h: PLOT.h, points: trace(blue, 60),
		line: { color: C.blue, width: 2, dashType: 'dash' },
	});
	pageNum(s, 9);
}

/** Filled silhouette, closed down to the baseline. */
function area(s, peaks, color) {
	const pts = trace(peaks, 72);
	pts.push({ x: PLOT.w, y: PLOT.h }, { x: 0, y: PLOT.h }, { close: true });
	s.addShape('custGeom', { x: PLOT.x, y: PLOT.y, w: PLOT.w, h: PLOT.h, points: pts, fill: { color } });
}

/** Stack of translucent white strips reproducing the colour→white fade. */
function whiteVeil(s, bands = 40) {
	const band = PLOT.h / bands;
	for (let b = 0; b < bands; b++) {
		rect(s, PLOT.x, PLOT.y + b * band, PLOT.w, band + 0.004,
			{ fill: { color: C.white, transparency: Math.round(100 - 80 * (b / (bands - 1)) ** 1.8) } });
	}
}

/* ============================================================= SLIDE  10 */

function slide10(s) {
	txt(s, rich([['Infographic ', C.blue], ['Section.', null]], 40), { x: 0.868, y: 0.730, w: 3.608, h: 1.447 });

	const cards = [
		{ x: 0.820, y: 3.853, color: C.blue, body: C.blueMid, title: 'Agenda 1', date: '3 Dec ', icon: 'form', light: false },
		{ x: 3.847, y: 3.115, color: C.yellow, body: C.yellowMid, title: 'Agenda 2', date: '5 Dec ', icon: 'target', light: false },
		{ x: 6.874, y: 1.896, color: C.white, body: C.purpleMid, title: 'Agenda 3', date: '7 Dec ', icon: 'calendar', light: true },
		{ x: 9.901, y: 3.853, color: C.pink, body: C.pinkMid, title: 'Agenda 4', date: '8 Dec ', icon: 'form', light: false },
	];
	cards.forEach(c => {
		rrect(s, c.x, c.y, 2.613, 2.917, 0.122, { fill: { color: c.light ? C.white : c.body }, shadow: c.light ? sh('card') : undefined });
		const ink = c.light ? C.black : C.white;
		txt(s, c.title, { x: c.x + 0.036, y: c.y + 0.800, w: 2.540, h: 0.438, fontSize: 20, fontFace: HEAD, color: ink, align: 'center' });
		txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
			{ x: c.x + 0.209, y: c.y + 1.303, w: 2.223, h: 0.808, fontSize: 14, color: ink, align: 'center' });
		// date badge hanging off the top edge
		const bx = c.x + 0.912, by = c.y - 0.343;
		rrect(s, bx, by, 1.517, 0.679, 0.201, { fill: { color: c.light ? C.purpleMid : C.white }, shadow: sh('card') });
		txt(s, c.date, { x: bx + 0.243, y: by + 0.120, w: 1.033, h: 0.438, fontSize: 20, fontFace: HEAD, color: c.light ? C.white : C.black, align: 'center' });
		rrect(s, c.x + 0.737, c.y - 0.499, 0.421, 0.407, 0.244, { fill: { color: c.light ? C.white : c.color }, shadow: sh('badge') });
		glyph(s, c.icon, c.x + 0.800, c.y - 0.437, 0.295, c.light ? C.purple : C.white);
	});

	txt(s, 'Subtitle Here', { x: 10.016, y: 1.219, w: 2.814, h: 0.438, fontSize: 20 });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
		{ x: 10.016, y: 1.662, w: 2.814, h: 1.043, fontSize: 14, color: C.grey });

	pageNum(s, 10);
}

/* ============================================================= SLIDE  11 */

function slide11(s) {
	txt(s, rich([['Infographic ', C.blue], ['Section.', null]], 40), { x: 0.868, y: 0.730, w: 3.608, h: 1.447 });

	// left agenda list
	const list = [
		{ y: 2.508, title: 'The opening agenda', time: 'Join with us at 10:30 – 12:30', color: C.pink },
		{ y: 3.591, title: 'The main event', time: 'Join with us at 09:30 – 15:00', color: C.purple },
		{ y: 4.674, title: 'The coffee break', time: 'Join with us at 07:30 – 9:00', color: C.yellow },
		{ y: 5.756, title: 'The closing agenda', time: 'Join with us at 15:00 – 17:00', color: C.blue },
	];
	list.forEach(it => {
		glyph(s, 'calendar', 1.225, it.y + 0.079, 0.374, it.color);
		txt(s, it.title, { x: 1.714, y: it.y, w: 3.608, h: 0.438, fontSize: 20, fontFace: HEAD, color: it.color });
		txt(s, it.time, { x: 1.714, y: it.y + 0.364, w: 3.608, h: 0.337, fontSize: 14 });
	});

	// 2x2 quadrant of date tiles; the purple "11 Dec" tile is oversized
	const tiles = [
		{ x: 8.603, y: 0.831, size: 3.127, color: C.purpleMid, num: '11', nx: 8.946, ny: 1.532, nw: 1.386, nh: 1.212, nsz: 66,
			dx: 9.971, dy: 1.880, dw: 1.250, dsz: 36, bx: 9.191, by: 2.550, bw: 1.914, bsz: 12,
			icon: 'cart', ix: 11.221, iy: 1.845, isz: 1.245, gsz: 0.484 },
		{ x: 5.899, y: 1.535, size: 2.423, color: C.pinkMid, num: '12', nx: 6.251, ny: 2.034, nw: 1.000, nh: 0.909, nsz: 48,
			dx: 7.077, dy: 2.278, dw: 0.901, dsz: 28, bx: 6.352, by: 2.853, bw: 1.657, bsz: 10,
			icon: 'gear', ix: 5.378, iy: 2.260, isz: 0.901, gsz: 0.342 },
		{ x: 5.899, y: 4.246, size: 2.423, color: C.blueMid, num: '10', nx: 6.257, ny: 4.745, nw: 1.000, nh: 0.909, nsz: 48,
			dx: 7.084, dy: 4.989, dw: 0.901, dsz: 28, bx: 6.358, by: 5.564, bw: 1.657, bsz: 10,
			icon: 'cloud', ix: 5.349, iy: 4.991, isz: 0.901, gsz: 0.397 },
		{ x: 8.603, y: 4.230, size: 2.423, color: C.yellowMid, num: '13', nx: 8.842, ny: 4.745, nw: 1.000, nh: 0.909, nsz: 48,
			dx: 9.668, dy: 4.989, dw: 0.901, dsz: 28, bx: 8.942, by: 5.564, bw: 1.657, bsz: 10,
			icon: 'case', ix: 10.688, iy: 5.063, isz: 0.901, gsz: 0.406 },
	];
	tiles.forEach(t => {
		rrect(s, t.x, t.y, t.size, t.size, 0.103, { fill: { color: t.color } });
		txt(s, t.num, { x: t.nx, y: t.ny, w: t.nw, h: t.nh, fontSize: t.nsz, fontFace: HEAD, color: C.white, align: 'center' });
		txt(s, 'Dec', { x: t.dx, y: t.dy, w: t.dw, h: t.nh * 0.58, fontSize: t.dsz, color: C.white });
		txt(s, ['Lorem ipsum dolor', 'sit amet, consect', 'adipiscing elit. '].map(text => ({
			text, options: { bullet: { characterCode: '2713', indent: 13 }, breakLine: true },
		})), { x: t.bx, y: t.by, w: t.bw, h: 0.65, fontSize: t.bsz, color: C.white });
		rrect(s, t.ix, t.iy, t.isz, t.isz, R.def, { fill: { color: C.white }, shadow: sh('left') });
		glyph(s, t.icon, t.ix + (t.isz - t.gsz) / 2, t.iy + (t.isz - t.gsz) / 2, t.gsz, t.color);
	});

	pageNum(s, 11);
}

/* ============================================================= SLIDE  12 */

function slide12(s) {
	txt(s, rich([['Infographic ', C.blue], ['Section.', null]], 40),
		{ x: 3.142, y: 0.730, w: 7.049, h: 0.774, align: 'center' });

	// top time strip: three flat cells + one raised card
	[{ x: -0.001, w: 3.764, tx: 1.085, bx: 2.091, time: '11:00', color: C.yellow },
	 { x: 3.762, w: 3.118, tx: 4.202, bx: 5.208, time: '11:00', color: C.pink },
	 { x: 9.731, w: 3.601, tx: 10.386, bx: 11.392, time: '11:00', color: C.blue }].forEach(cell => {
		rect(s, cell.x, 2.048, cell.w, 0.836, { line: { color: C.greyLine, width: 1 } });
		txt(s, cell.time, { x: cell.tx, y: 2.213, w: 1.087, h: 0.505, fontSize: 24, fontFace: HEAD, color: cell.color });
		txt(s, 'Lorem ipsum dolor sit amet.', { x: cell.bx, y: 2.248, w: 1.232, h: 0.434, fontSize: 11, color: C.grey, lineSpacingMultiple: 0.9 });
	});
	rrect(s, 6.829, 1.898, 3.118, 1.021, R.def, { fill: { color: C.white }, line: { color: C.greyLine, width: 1 }, shadow: sh('card') });
	txt(s, '12:00', { x: 7.122, y: 2.122, w: 1.330, h: 0.572, fontSize: 28, fontFace: HEAD, color: C.purple });
	txt(s, 'Lorem ipsum dolor sit amet.', { x: 8.351, y: 2.176, w: 1.303, h: 0.464, fontSize: 12, lineSpacingMultiple: 0.9 });

	// four bullet columns
	const bullets = ['Lorem ipsum dolor', 'Sit amet consectetuer', 'Adipiscing elit aenean', 'Commodo ligula eget', 'Dolor aenean massa'];
	[{ x: 0.673, color: C.yellow, button: C.yellowMid, ink: C.grey },
	 { x: 3.790, color: C.pink, button: C.pinkMid, ink: C.grey },
	 { x: 6.857, color: C.purple, button: C.purpleMid, ink: C.black },
	 { x: 9.974, color: C.blue, button: C.blueMid, ink: C.grey }].forEach(col => {
		txt(s, 'Todays Agenda?', { x: col.x, y: 3.462, w: 2.687, h: 0.438, fontSize: 20, fontFace: HEAD, color: col.color });
		txt(s, bullets.map(text => ({ text, options: { bullet: { characterCode: '2713' }, breakLine: true } })),
			{ x: col.x, y: 3.915, w: 2.687, h: 1.827, fontSize: 14, color: col.ink, lineSpacingMultiple: 1.5 });
		rrect(s, col.x, 6.000, 2.687, 0.482, R.def, { fill: { color: col.button } });
		txt(s, 'Lorem ipsum dolor sit', { x: col.x + 0.164, y: 6.073, w: 2.358, h: 0.337, fontSize: 14, fontFace: HEAD, color: C.white, align: 'center' });
	});

	pageNum(s, 12);
}

/* ============================================================= SLIDE  13 */

function slide13(s) {
	txt(s, rich([['Infographic ', C.blue], ['Section.', null]], 40), { x: 0.858, y: 0.594, w: 3.714, h: 1.447 });

	// dashed "journey" curves running behind the calendar cards
	s.addShape('custGeom', {
		x: 0.774, y: 1.164, w: 11.323, h: 4.134, line: { color: C.gridLine, width: 3, dashType: 'sysDash' },
		points: [
			{ x: 0.0, y: 4.134 },
			{ x: 3.0, y: 2.6, curve: { type: 'cubic', x1: 0.9, y1: 4.1, x2: 1.9, y2: 2.7 } },
			{ x: 5.3, y: 1.15, curve: { type: 'cubic', x1: 4.1, y1: 2.5, x2: 4.3, y2: 1.1 } },
			{ x: 8.0, y: 2.2, curve: { type: 'cubic', x1: 6.4, y1: 1.2, x2: 7.6, y2: 1.4 } },
			{ x: 11.323, y: 0.0, curve: { type: 'cubic', x1: 8.6, y1: 3.2, x2: 10.6, y2: 1.4 } },
		],
	});
	s.addShape('custGeom', {
		x: 6.079, y: 2.311, w: 2.084, h: 2.839, line: { color: C.gridLine, width: 3, dashType: 'sysDash' },
		points: [
			{ x: 2.084, y: 0.0 },
			{ x: 1.0, y: 1.5, curve: { type: 'cubic', x1: 2.0, y1: 0.8, x2: 1.4, y2: 1.1 } },
			{ x: 0.0, y: 2.839, curve: { type: 'cubic', x1: 0.6, y1: 1.9, x2: 0.1, y2: 2.2 } },
		],
	});

	const cards = [
		{ x: 1.397, y: 2.582, size: 2.443, num: '10', numSize: 88, month: 20, tone: C.blueMid, body: C.white, bx: 1.001, bs: 0.792, is: 0.541 },
		{ x: 4.680, y: 1.515, size: 2.839, num: '12', numSize: 96, month: 24, tone: C.purple, body: C.purpleMid, bx: 4.219, bs: 0.921, is: 0.629 },
		{ x: 8.142, y: 1.325, size: 1.906, num: '13', numSize: 66, month: 14, tone: C.yellowMid, body: C.white, bx: 7.833, bs: 0.618, is: 0.422 },
		{ x: 9.890, y: 3.442, size: 2.443, num: '17', numSize: 88, month: 20, tone: C.pinkMid, body: C.white, bx: 9.494, bs: 0.792, is: 0.541 },
	];
	cards.forEach(c => {
		const hdr = c.size * 0.261;
		const dark = c.body !== C.white;
		rrect(s, c.x, c.y, c.size, c.size, 0.155, { fill: { color: c.body }, shadow: dark ? undefined : sh('card') });
		s.addShape('round2SameRect', { x: c.x, y: c.y, w: c.size, h: hdr, fill: { color: c.tone } });
		txt(s, 'December', { x: c.x + c.size * 0.128, y: c.y + hdr * 0.16, w: c.size * 0.72, h: 0.44, fontSize: c.month, fontFace: HEAD, color: C.white, align: 'center' });
		txt(s, c.num, { x: c.x, y: c.y + hdr + c.size * 0.04, w: c.size, h: c.size - hdr - 0.05, fontSize: c.numSize, fontFace: HEAD, color: dark ? C.white : C.black, align: 'center' });
		rrect(s, c.bx, c.y + c.size * 0.553, c.bs, c.bs, R.def, { fill: { color: C.white }, shadow: sh('card') });
		const off = (c.bs - c.is) / 2;
		rrect(s, c.bx + off, c.y + c.size * 0.553 + off, c.is, c.is, R.icon, { fill: { color: dark ? C.purpleMid : c.tone }, shadow: sh('card') });
		glyph(s, 'calendar', c.bx + off + c.is * 0.21, c.y + c.size * 0.553 + off + c.is * 0.21, c.is * 0.58, C.white);
	});

	[{ x: 0.471, y: 5.400, w: 2.851, lead: 'Start ', date: '10 Dec', body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' },
	 { x: 5.728, y: 5.268, w: 2.851, lead: 'Main Event! ', date: '12 Dec', body: 'Lorem ipsum dolor sit amet!' },
	 { x: 10.286, y: 1.270, w: 2.686, lead: 'End ', date: '10 Dec', body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' }]
		.forEach(cap => {
			txt(s, rich([[cap.lead, C.blue], [cap.date, null]], 20), { x: cap.x, y: cap.y, w: cap.w, h: 0.438, wrap: false });
			txt(s, cap.body, { x: cap.x, y: cap.y + 0.364, w: cap.w, h: 0.572, fontSize: 14, color: C.grey });
		});

	pageNum(s, 13);
}

/* ============================================================= SLIDE  14 */

function slide14(s) {
	meetingScene(s);
	txt(s, rich([['Infographic ', C.blue], ['Section.', null]], 40),
		{ x: 8.731, y: 0.733, w: 3.707, h: 1.447, align: 'right' });

	const list = [
		{ y: 1.018, title: 'The opening agenda', time: 'Join with us at 10:30 – 12:30', color: C.pink },
		{ y: 2.552, title: 'The main event', time: 'Join with us at 09:30 – 15:00', color: C.purple },
		{ y: 4.067, title: 'The coffee break', time: 'Join with us at 07:30 – 9:00', color: C.yellow },
		{ y: 5.596, title: 'The closing agenda', time: 'Join with us at 15:00 – 17:00', color: C.blue },
	];
	list.forEach(it => {
		rrect(s, 0.914, it.y + 0.098, 0.785, 0.785, R.def, { fill: { color: C.white }, shadow: sh('card') });
		glyph(s, 'clock', 1.088, it.y + 0.271, 0.438, it.color);
		txt(s, it.title, { x: 1.933, y: it.y, w: 4.400, h: 0.572, fontSize: 28, fontFace: HEAD, color: it.color });
		txt(s, it.time, { x: 1.933, y: it.y + 0.482, w: 4.400, h: 0.404 });
	});

	pageNum(s, 14);
}

/** Flat-illustration stand-in: two people at a laptop table on a grey blob. */
function meetingScene(s) {
	// organic backdrop blob: rounded left flank, tall lobe on the right
	s.addShape('custGeom', {
		x: 6.40, y: 1.58, w: 6.60, h: 5.65, fill: { color: 'E4E4E4' },
		points: [
			{ x: 0.80, y: 1.70 },
			{ x: 3.90, y: 1.45, curve: { type: 'cubic', x1: 1.55, y1: 1.00, x2: 2.90, y2: 1.60 } },
			{ x: 4.65, y: 0.05, curve: { type: 'cubic', x1: 4.35, y1: 1.32, x2: 4.10, y2: 0.05 } },
			{ x: 5.75, y: 1.70, curve: { type: 'cubic', x1: 5.30, y1: 0.05, x2: 5.65, y2: 0.75 } },
			{ x: 6.60, y: 3.60, curve: { type: 'cubic', x1: 5.90, y1: 2.30, x2: 6.45, y2: 2.70 } },
			{ x: 4.60, y: 5.62, curve: { type: 'cubic', x1: 6.72, y1: 4.75, x2: 6.05, y2: 5.62 } },
			{ x: 1.05, y: 5.50, curve: { type: 'cubic', x1: 2.95, y1: 5.62, x2: 1.85, y2: 5.62 } },
			{ x: 0.80, y: 1.70, curve: { type: 'cubic', x1: -0.60, y1: 5.00, x2: -0.40, y2: 2.35 } },
			{ close: true },
		],
	});

	// chairs sit behind the two figures
	[{ x: 6.62, flip: false }, { x: 12.10, flip: true }].forEach(ch => {
		s.addShape('round2SameRect', { x: ch.x, y: 4.75, w: 0.26, h: 1.45, fill: { color: 'A8B2BE' } });
		rect(s, ch.x - 0.05, 6.15, 1.05, 0.14, { fill: { color: 'A8B2BE' } });
		rect(s, ch.x + (ch.flip ? 0.72 : 0.06), 6.25, 0.13, 0.85, { fill: { color: 'A8B2BE' } });
	});

	// seated figures — the left one faces right, the right one faces left
	seated(s, { hair: 7.42, face: 7.78, torso: 7.15, shirt: C.pink, hip: 7.06, knee: 8.20, shin: 7.83, dir: 1 });
	seated(s, { hair: 11.90, face: 11.45, torso: 11.15, shirt: '1B1B43', hip: 12.31, knee: 10.85, shin: 11.18, dir: -1 });

	// table + two laptops
	rect(s, 7.94, 5.10, 3.32, 0.13, { fill: { color: '8B98A8' } });
	[8.12, 10.95].forEach(x => rect(s, x, 5.23, 0.12, 1.60, { fill: { color: '9AA5B1' } }));
	[{ x: 8.62, flip: false }, { x: 10.28, flip: true }].forEach(l => {
		s.addShape('parallelogram', { x: l.x, y: 4.42, w: 0.55, h: 0.66, flipH: l.flip, fill: { color: 'BFC6CF' } });
		rect(s, l.x - 0.16, 5.02, 0.86, 0.09, { fill: { color: '8B94A0' } });
	});
	rect(s, 6.30, 7.10, 6.75, 0.30, { fill: { color: 'B6B6B6' } });   // floor line
	txt(s, '[illustration]', { x: 6.30, y: 6.60, w: 6.75, h: 0.3, fontSize: 12, color: '9A9A9A', align: 'center' });
}

/** One seated person: head, torso, arm reaching forward, thigh, shin, shoe. */
function seated(s, p) {
	ellipse(s, p.hair, 3.58, 0.62, 0.68, { fill: { color: '1B1B43' } });
	ellipse(s, p.face, 3.72, 0.38, 0.52, { fill: { color: 'F0B49A' } });
	s.addShape('round2SameRect', { x: p.torso, y: 4.25, w: 1.05, h: 1.15, fill: { color: p.shirt } });
	rect(s, Math.min(p.torso + 0.5, p.torso + 0.5 + p.dir * 0.9), 4.62, 0.9, 0.28, { fill: { color: p.shirt } });
	rect(s, Math.min(p.hip, p.knee), 5.25, Math.abs(p.knee - p.hip), 0.52, { fill: { color: '2E7DEB' } });
	rect(s, p.shin, 5.55, 0.34, 1.30, { fill: { color: '1C5FBF' } });
	rect(s, p.shin - 0.06, 6.78, 0.48, 0.16, { fill: { color: '4A5568' } });
}

/* ============================================================= SLIDE  15 */

function slide15(s) {
	txt(s, rich([['Closing Session ', C.blue], ['with Key ', null], ['Insights', C.blue], [' and ', null], ['Learnings', C.blue]], 40),
		{ x: 6.271, y: 0.669, w: 6.361, h: 2.121, align: 'right' });
	txt(s, 'Subtitle Here', { x: 0.464, y: 1.303, w: 2.814, h: 0.438, fontSize: 20, align: 'right' });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
		{ x: 0.464, y: 1.746, w: 2.814, h: 1.043, fontSize: 14, color: C.grey, align: 'right' });

	// four staggered testimonial cards
	[{ x: 0.702, y: 3.612, color: C.pinkMid, star: C.pinkTint },
	 { x: 3.740, y: 2.243, color: C.yellowMid, star: C.yellowTint },
	 { x: 6.779, y: 2.927, color: C.blueMid, star: C.blueTint },
	 { x: 9.818, y: 3.612, color: C.purpleMid, star: C.purpleTint }].forEach(card => {
		rrect(s, card.x, card.y, 2.814, 3.219, 0.07, { fill: { color: card.color } });
		txt(s, 'Emily T., Head of Product', { x: card.x + 0.238, y: card.y + 0.377, w: 2.338, h: 0.707, fontFace: HEAD, color: C.white });
		txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem.',
			{ x: card.x + 0.238, y: card.y + 1.767, w: 2.338, h: 0.808, fontSize: 14, color: C.white });
		for (let i = 0; i < 5; i++) {
			s.addShape('star5', { x: card.x + 0.343 + i * 0.170, y: card.y + 2.628, w: 0.170, h: 0.159, fill: { color: card.star } });
		}
		txt(s, '4,9', { x: card.x + 1.176, y: card.y + 2.574, w: 0.440, h: 0.269, fontSize: 10, color: C.white });
	});

	pageNum(s, 15);
}

/* ============================================================= SLIDE  16 */

function slide16(s) {
	txt(s, 'Conversion Funnel.', { x: 0.278, y: 0.159, w: 1.944, h: 0.337, fontSize: 14 });
	txt(s, 'www.agendaplanner.com', { x: 5.194, y: 0.159, w: 2.944, h: 0.337, fontSize: 14, color: C.purple, align: 'center' });
	hamburger(s);

	txt(s, 'Thank You — Let’s Stay Connected!',
		{ x: 0.687, y: 0.956, w: 10.284, h: 2.794, fontSize: 80, fontFace: HEAD, color: C.blue });

	ring(s, 5.829, 4.962, 5.077);
	ring(s, 11.705, 3.332, 2.743);
	ring(s, 4.383, 4.339, 1.446);

	txt(s, 'Get in Touch', { x: 1.519, y: 5.566, w: 1.997, h: 0.438, fontSize: 20, underline: { style: 'sng' }, wrap: false });
	[{ y: 6.004, text: '+123 456 789 012 345', icon: 'phone' },
	 { y: 6.307, text: 'companybranding@gmail.com', icon: 'mail' },
	 { y: 6.587, text: 'www.companybranding.co', icon: 'globe' }].forEach(row => {
		glyph(s, row.icon, 0.660, row.y + 0.070, 0.172, C.purple);
		txt(s, row.text, { x: 0.938, y: row.y, w: 2.601, h: 0.303, fontSize: 12 });
	});

	pageNum(s, 16);
}

/* -------------------------------------------------------------- assemble */

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'AGENDA_16x9', width: 13.333, height: 7.5 });
	pptx.layout = 'AGENDA_16x9';
	pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
	pptx.author = 'Agenda Planner';
	pptx.title = 'Agenda Planner';

	const builders = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
		slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

	builders.forEach(fn => {
		const s = pptx.addSlide();
		s.background = { color: C.white };
		fn(s);
	});

	return pptx.writeFile({ fileName: path.join(__dirname, '0d83c6c8-9383-4bee-a468-b8e1c09bd367_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
