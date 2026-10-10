/**
 * "Phone Mockup — Infographic Presentation Template" (20 slides, 13.333in x 7.5in)
 * rebuilt with pptxgenjs only. Photographs are replaced by native-shape stand-ins.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const BLUE = '1B76FF';   // accent1
const ORANGE = 'F36426'; // accent2
const GREY3 = '969696';  // accent3
const GREY4 = '808080';  // accent4
const GREY6 = '4D4D4D';  // accent6
const INK = '000000';
const WHITE = 'FFFFFF';
const CARD = 'F2F2F2';
const BODY_GREY = '595959';
const DARK_GREY = '404040';
const NEAR_BLACK = '262626';
const SCREEN_TEXT = 'F4F5F6';
const FONT = 'Archivo';        // theme major font
const FONT_BODY = 'Archivo Light'; // theme minor font

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* wallpaper used on the mockup phone screens (pink -> cream -> blue) */
const WALLPAPER = ['E8407F', 'F2C0A8', 'F3EDE6', '7C86D8', '2F6BE0'];

/* the headline that repeats on most content slides */
const HEADLINE = [
	['Get professional results fast.\u00a0', false],
	['Showcase your app', true],
	['\u00a0with our stunning\u00a0', false],
	['Phone Mockups', true],
	['\u00a0today!', false],
];

const LOREM_CARD = [
	['00,01% dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ', { color: GREY6 }],
	['Cum sociis natoque penatibus et magnis', { color: BLUE, italic: true }],
	[' ', { color: BLUE, italic: true }],
	['dis parturient montes, nascetur ridiculus mus. ', { color: GREY6 }],
];

const SERENITY = 'A wonderful serenity has taken possession of my entire soul.';
const SERENITY_LONG = 'A wonderful serenity has taken possession of my entire soul, ' +
	'like these sweet mornings of spring which I enjoy with my whole heart.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor.';

/* ---------------------------------------------------------------- helpers */

/* pptxgenjs draws a 1pt grey stroke unless a shape opts out explicitly */
const NO_LINE = { type: 'none' };

const hex = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();

/** blend a list of evenly spaced hex stops at position t (0..1) */
function stopAt(stops, t) {
	if (stops.length === 1) return stops[0];
	const span = 1 / (stops.length - 1);
	const i = Math.min(stops.length - 2, Math.floor(t / span));
	const f = (t - i * span) / span;
	const a = stops[i], b = stops[i + 1];
	let out = '';
	for (let k = 0; k < 3; k++) {
		const ca = parseInt(a.substr(k * 2, 2), 16);
		const cb = parseInt(b.substr(k * 2, 2), 16);
		out += hex(ca + (cb - ca) * f);
	}
	return out;
}

function rotatePoint(px, py, cx, cy, deg) {
	const r = (deg * Math.PI) / 180;
	const dx = px - cx, dy = py - cy;
	return [cx + dx * Math.cos(r) - dy * Math.sin(r), cy + dx * Math.sin(r) + dy * Math.cos(r)];
}

/* snap to whole EMU so neighbouring strips abut exactly — semi-transparent
   strips would otherwise double up along a seam and print a visible line */
const snap = (v) => Math.round(v * 914400) / 914400;

/** how far a rounded corner of radius r bites into the edge, d along from it */
function cornerInset(d, r) {
	if (!r || d >= r) return 0;
	return r - Math.sqrt(Math.max(0, r * r - (r - d) * (r - d)));
}

/**
 * Linear gradient across a rectangle, painted as a run of solid slices.
 * dir defaults to the long axis and `rot` spins the band about its own centre.
 *
 * With `radius` the slices are rounded rectangles that each run from their own
 * start edge to the far end of the band, so every later slice paints over the
 * previous one and only the two extreme slices contribute a rounded corner —
 * which is exactly how a rounded panel with a linear fill looks. Straight
 * strips can't do that: they would leave the corner arcs uncovered.
 */
function gradBand(slide, o) {
	const dir = o.dir || (o.w >= o.h ? 'h' : 'v');
	const len = dir === 'h' ? o.w : o.h;
	const steps = o.steps || Math.max(8, Math.min(90, Math.round(len * 16)));
	const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
	const rounded = o.radius > 0 && !o.transparency;
	// overlapping slices would double-darken a translucent pass, so butt them
	const bleed = o.transparency ? 0 : 0.012;
	const edge = (i) => snap((len * i) / steps);

	for (let i = 0; i < steps; i++) {
		const e0 = edge(i), e1 = edge(i + 1);
		const tail = rounded ? len : e1;
		const cut = rounded ? 0 : cornerInset(Math.min(e0, len - e1), o.radius);
		let sx, sy, sw, sh;
		if (dir === 'h') { sx = o.x + e0; sy = o.y + cut; sw = tail - e0 + bleed; sh = o.h - 2 * cut; }
		else { sx = o.x + cut; sy = o.y + e0; sw = o.w - 2 * cut; sh = tail - e0 + bleed; }
		const fill = { color: stopAt(o.colors, (e0 + e1) / 2 / len) };
		if (o.transparency) fill.transparency = o.transparency;
		const opts = { x: sx, y: sy, w: sw, h: sh, fill: fill, line: NO_LINE };
		if (rounded) opts.rectRadius = o.radius;
		if (o.rot) {
			const [nx, ny] = rotatePoint(sx + sw / 2, sy + sh / 2, cx, cy, o.rot);
			opts.x = nx - sw / 2; opts.y = ny - sh / 2; opts.rotate = o.rot;
		}
		slide.addShape(rounded ? 'roundRect' : 'rect', opts);
	}
}

/** rounded pill/bar carrying a horizontal gradient (caps keep the end colours) */
function gradPill(slide, o) {
	const r = o.h / 2;
	const cap = Math.min(2 * r, o.w);
	slide.addShape('roundRect', { x: o.x, y: o.y, w: cap, h: o.h, rectRadius: r, fill: { color: o.colors[0] }, line: NO_LINE });
	slide.addShape('roundRect', { x: o.x + o.w - cap, y: o.y, w: cap, h: o.h, rectRadius: r, fill: { color: o.colors[o.colors.length - 1] }, line: NO_LINE });
	if (o.w > 2 * r) gradBand(slide, { x: o.x + r, y: o.y, w: o.w - 2 * r, h: o.h, colors: o.colors, dir: 'h' });
}

/**
 * Corner-to-corner (45-degree) linear ramp filling a rectangle. The colour only
 * depends on dx + dy, so the panel is painted with strips laid along its
 * anti-diagonals, each trimmed to the chord that line cuts through the panel.
 * The strip ends still poke out by half a strip width, so a thin frame in the
 * surrounding colour trims them — draw this before whatever sits next to it.
 */
function gradDiag(slide, o) {
	const reach = o.w + o.h;
	const steps = o.steps || Math.max(24, Math.min(300, Math.round(reach / 0.045)));
	const sw = reach / steps;
	for (let i = 0; i < steps; i++) {
		const a = (reach * (i + 0.5)) / steps;
		const lo = Math.max(0, a - o.h), hi = Math.min(o.w, a);
		const len = (hi - lo) * Math.SQRT2 + sw;
		slide.addShape('rect', {
			x: o.x + (lo + hi) / 2 - (sw * Math.SQRT1_2 + 0.01) / 2,
			y: o.y + a - (lo + hi) / 2 - len / 2,
			w: sw * Math.SQRT1_2 + 0.01, h: len, rotate: 45,
			fill: { color: stopAt(o.colors, a / reach) }, line: NO_LINE,
		});
	}
	const pad = sw, mask = o.mask || WHITE;
	[[o.x - pad, o.y - pad, o.w + 2 * pad, pad], [o.x - pad, o.y + o.h, o.w + 2 * pad, pad],
		[o.x - pad, o.y - pad, pad, o.h + 2 * pad], [o.x + o.w, o.y - pad, pad, o.h + 2 * pad]].forEach(([x, y, w, h]) => {
		slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: mask }, line: NO_LINE });
	});
}

/** the same 45-degree ramp inside an ellipse, cut to chords of the outline */
function ellipseDiag(slide, o) {
	const a = o.w / 2, b = o.h / 2, cx = o.x + a, cy = o.y + b;
	const reach = o.w + o.h;
	const steps = o.steps || Math.max(40, Math.min(320, Math.round(reach * 18)));
	const sw = reach / steps;
	slide.addShape('ellipse', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: stopAt(o.colors, 0.5) }, line: NO_LINE });

	const A = 1 / (a * a) + 1 / (b * b);
	// where the line dx + dy = c meets the ellipse (centred coordinates)
	const halfSpread = (c) => Math.sqrt(Math.max(0, A - (c * c) / (a * a * b * b))) / A;

	for (let i = 0; i < steps; i++) {
		const t = (i + 0.5) / steps;
		const c = (t - 0.5) * reach;
		// clip to the shorter of the strip's two chords so nothing overhangs
		const dx = Math.min(halfSpread(c - sw / 2), halfSpread(c + sw / 2));
		if (dx <= 0) continue;
		const mx = c / (b * b) / A;              // chord midpoint x
		slide.addShape('rect', {
			x: cx + mx - (sw * Math.SQRT1_2 + 0.01) / 2,
			y: cy + (c - mx) - dx * Math.SQRT2,
			w: sw * Math.SQRT1_2 + 0.01, h: 2 * dx * Math.SQRT2, rotate: 45,
			fill: { color: stopAt(o.colors, t) }, line: NO_LINE,
		});
	}
}

/** stand-in for a raster image: soft grey plate with an optional caption */
function imgBox(slide, o) {
	slide.addShape(o.shape || 'rect', {
		x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rot || 0,
		fill: { color: o.fill || 'D9D9D9', transparency: o.transparency || 0 },
		line: o.lineColor ? { color: o.lineColor, width: 1 } : NO_LINE,
	});
	if (o.label !== null) {
		slide.addText(o.label || '[image]', {
			x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rot || 0, margin: 0,
			align: 'center', valign: 'middle', fontFace: FONT,
			fontSize: o.labelSize || 10, color: o.labelColor || '8A8A8A',
		});
	}
}

/** phone mockup: rounded body, inset screen, optional gradient wallpaper */
function phone(slide, o) {
	const bezel = o.bezel === undefined ? 0.09 : o.bezel;
	const body = o.body || '2B2B2B';
	const screenRadius = (o.radius || 0.3) - bezel / 2;
	slide.addShape('roundRect', {
		x: o.x - bezel, y: o.y - bezel, w: o.w + 2 * bezel, h: o.h + 2 * bezel,
		rotate: o.rot || 0, rectRadius: o.radius || 0.3,
		fill: { color: body }, line: o.bodyLine ? { color: o.bodyLine, width: 1.5 } : NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 14, offset: 3, angle: 90, opacity: 0.12 },
	});
	const screen = Array.isArray(o.screen) ? o.screen : [o.screen || 'F7F7F7'];
	slide.addShape('roundRect', {
		x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rot || 0, rectRadius: screenRadius,
		fill: { color: stopAt(screen, 0.5) }, line: NO_LINE,
	});
	if (screen.length > 1) {
		gradBand(slide, {
			x: o.x, y: o.y, w: o.w, h: o.h, colors: screen,
			dir: 'v', radius: screenRadius, rot: o.rot || 0,
		});
	}
	// the notch hangs down into the screen: a wide bar (iPhone) or a dot (Android)
	if (o.notch) {
		const drop = o.notch === 'drop';
		const nw = o.w * (drop ? 0.075 : 0.42);
		const nh = drop ? nw : o.w * 0.075;
		const [nx, ny] = rotatePoint(o.x + o.w / 2, o.y + (drop ? nh * 0.9 : nh / 2), o.x + o.w / 2, o.y + o.h / 2, o.rot || 0);
		slide.addShape(drop ? 'ellipse' : 'roundRect', {
			x: nx - nw / 2, y: ny - nh / 2, w: nw, h: nh, rotate: o.rot || 0,
			rectRadius: nh / 2, fill: { color: body }, line: NO_LINE,
		});
	}
	if (o.label) {
		slide.addText(o.label, {
			x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rot || 0, margin: 0,
			align: 'center', valign: 'middle', fontFace: FONT,
			fontSize: o.labelSize || 14, color: o.labelColor || 'A6A6A6',
		});
	}
}

/** the padlock glyph drawn on the phone lock screens */
function padlock(slide, x, y, s, color) {
	slide.addShape('ellipse', { x: x + s * 0.18, y: y, w: s * 0.64, h: s * 0.64, fill: { type: 'none' }, line: { color: color, width: s * 9 } });
	slide.addShape('roundRect', { x: x, y: y + s * 0.42, w: s, h: s * 0.58, rectRadius: s * 0.12, fill: { color: color }, line: NO_LINE });
}

/** circular button with an arrow, used as a "read more" affordance */
function arrowButton(slide, x, y, d) {
	ellipseDiag(slide, { x: x, y: y, w: d, h: d, colors: [ORANGE, BLUE], steps: 30 });
	slide.addShape('triangle', { x: x + d * 0.36, y: y + d * 0.28, w: d * 0.3, h: d * 0.44, rotate: 90, fill: { color: WHITE }, line: NO_LINE });
}

/* master furniture repeated on slides 1-19 (page tag, big number, footer) */
function chrome(slide, num, opts) {
	const o = opts || {};
	slide.addText('Phone Mockup', {
		x: 0.445, y: 0.369, w: 2.416, h: 0.177, margin: 0,
		fontFace: FONT, fontSize: 10.5, color: o.tagColor || 'BFBFBF',
	});
	slide.addText(String(num), {
		x: 10.326, y: 0, w: 2.666, h: 1.716, margin: 0, align: 'right', valign: 'middle',
		fontFace: FONT, fontSize: 96, color: o.numColor || 'D9D9D9', transparency: 82,
	});
	slide.addShape('roundRect', {
		x: 9.624, y: 6.894, w: 2.861, h: 0.27, rectRadius: 0.135,
		fill: { color: o.pillColor || 'F5F5F5' }, line: NO_LINE,
	});
	slide.addText('Copyright \u00a9. All rights.', {
		x: 9.767, y: 6.941, w: 2.297, h: 0.177, margin: 0, align: 'right',
		fontFace: FONT, fontSize: 10.5, color: o.footColor || 'BFBFBF',
	});
	gradPill(slide, { x: 12.207, y: 6.893, w: 0.824, h: 0.268, colors: [ORANGE, BLUE] });
	for (let i = 0; i < 3; i++) {
		slide.addShape('ellipse', { x: 12.35 + i * 0.204, y: 6.959, w: 0.137, h: 0.137, fill: { color: WHITE }, line: NO_LINE });
	}
}

/** the 5-run marketing headline, at the requested size / position */
function headline(slide, o) {
	slide.addText(HEADLINE.map(([t, b]) => ({ text: t, options: { bold: b } })), {
		x: o.x, y: o.y, w: o.w, h: o.h, margin: 0, align: o.align || 'left',
		fontFace: FONT, fontSize: o.size, color: INK,
	});
}

/** grey rounded card holding the "00,01% dolor…" paragraph */
function loremCard(slide, x, y) {
	slide.addShape('roundRect', {
		x: x, y: y, w: 4.739, h: 2.204, rectRadius: 0.105,
		fill: { color: CARD }, line: NO_LINE,
	});
	slide.addText(LOREM_CARD.map(([t, op]) => ({ text: t, options: op })), {
		x: x + 0.303, y: y + 0.301, w: 4.135, h: 1.601, margin: 7.2,
		fontFace: FONT_BODY, fontSize: 14, lineSpacingMultiple: 1.3,
	});
}

/** QR code + the two app-store badges + caption */
function downloadApp(slide, x, y) {
	slide.addText('Download app:', {
		x: x + 0.012, y: y, w: 2.881, h: 0.36, margin: 0,
		fontFace: FONT, fontSize: 14, color: NEAR_BLACK, lineSpacingMultiple: 1.3,
	});
	imgBox(slide, { x: x, y: y + 0.603, w: 1.143, h: 1.198, fill: 'FFFFFF', lineColor: '111111', label: '[QR]', labelColor: '333333' });
	imgBox(slide, { x: x + 1.848, y: y + 0.96, w: 1.137, h: 0.351, fill: '111111', label: 'Google Play', labelSize: 8, labelColor: WHITE });
	imgBox(slide, { x: x + 1.848, y: y + 1.369, w: 1.137, h: 0.351, fill: '111111', label: 'App Store', labelSize: 8, labelColor: WHITE });
}

/** small gradient pictogram used in the "features" lists */
function pictogram(slide, x, y, w, h) {
	gradBand(slide, { x: x, y: y, w: w, h: h, colors: [BLUE, ORANGE], dir: 'h', steps: 8, radius: Math.min(w, h) * 0.22 });
}

/* ------------------------------------------------------------- the slides */

function slide01(s) {
	chrome(s, 1);
	gradDiag(s, { x: 0.768, y: 4.6, w: 10.218, h: 1.457, colors: [BLUE, ORANGE] });
	s.addText('Phone Mockup', {
		x: 1.125, y: 4.723, w: 9.081, h: 1.212, margin: 0,
		fontFace: FONT, fontSize: 72, bold: true, color: WHITE,
	});
	s.addShape('rect', {
		x: 8.868, y: 6.042, w: 3.37, h: 0.5, fill: { color: WHITE }, line: NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 12, offset: 2, angle: 90, opacity: 0.13 },
	});
	s.addText('Infographic Presentation Template', {
		x: 9.032, y: 6.191, w: 3.043, h: 0.202, margin: 0, fontFace: FONT, fontSize: 12, color: INK,
	});
	s.addText('Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor congue massa. ', {
		x: 0.987, y: 3.414, w: 3.452, h: 0.569, margin: 0,
		fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.5,
	});
}

function slide02(s) {
	chrome(s, 2);
	phone(s, { x: 6.361, y: 3.085, w: 1.5, h: 4.426, body: '3A3A3A', screen: WHITE, radius: 0.24, notch: 'bar' });
	phone(s, { x: 10.914, y: 3.076, w: 1.5, h: 4.424, body: '3A3A3A', screen: WHITE, radius: 0.24, notch: 'bar' });
	phone(s, { x: 8.015, y: 2.211, w: 2.745, h: 5.825, body: '2B2B2B', screen: WHITE, radius: 0.34, notch: 'bar' });
	headline(s, { x: 1.122, y: 1.238, w: 4.196, h: 2.356, size: 28 });
	downloadApp(s, 1.11, 4.461);
}

function slide03(s) {
	chrome(s, 3);
	// four outline phones tumbling across the left half
	[[1.644, -0.494], [0.736, 3.519], [-2.094, -1.398], [4.515, 4.424]].forEach(([x, y]) => {
		phone(s, { x: x, y: y, w: 2.354, h: 5.198, rot: 330, body: WHITE, bodyLine: '111111', screen: WHITE, radius: 0.32, notch: 'drop' });
	});
	loremCard(s, 4.248, 3.885);
	headline(s, { x: 5.884, y: 1.411, w: 5.543, h: 1.885, size: 28 });
	s.addText('100+ Sources', {
		x: 9.735, y: 4.405, w: 1.318, h: 0.275, margin: 0, fontFace: FONT, fontSize: 14, bold: true, color: INK,
	});
	s.addText(SERENITY_LONG, {
		x: 9.735, y: 4.764, w: 2.846, h: 1.024, margin: 0,
		fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.3,
	});
}

function slide04(s) {
	chrome(s, 4);
	phone(s, { x: 5.49, y: 3.077, w: 2.354, h: 5.198, body: '1F1F1F', screen: WHITE, radius: 0.32, notch: 'drop' });
	phone(s, { x: 3.507, y: 3.668, w: 2.354, h: 5.198, body: '1F1F1F', screen: WHITE, radius: 0.32, notch: 'drop' });
	phone(s, { x: 7.472, y: 3.668, w: 2.354, h: 5.198, body: '1F1F1F', screen: WHITE, radius: 0.32, notch: 'drop' });
	headline(s, { x: 2.923, y: 0.913, w: 7.487, h: 1.414, size: 28, align: 'center' });

	// left stat card — "World / 85% / Lorem"
	s.addShape('roundRect', { x: 2.312, y: 4.022, w: 2.271, h: 2.862, rectRadius: 0.194, fill: { color: CARD }, line: NO_LINE });
	s.addText('World', { x: 2.69, y: 4.232, w: 1.516, h: 0.37, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 16, color: GREY6 });
	imgBox(s, { x: 2.923, y: 4.651, w: 1.05, h: 1.058, shape: 'ellipse', fill: '5A5A62', label: null });
	s.addText('85%', { x: 2.661, y: 5.782, w: 1.573, h: 0.64, margin: 7.2, align: 'center', wrap: false, fontFace: FONT, fontSize: 32, italic: true, color: GREY6 });
	s.addText('Lorem', { x: 2.858, y: 6.289, w: 1.18, h: 0.303, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 12, color: GREY6 });

	// right stat card — "This Week / 89322 / Accounts"
	s.addShape('roundRect', { x: 9.208, y: 2.846, w: 2.483, h: 2.917, rectRadius: 0.164, fill: { color: CARD }, line: NO_LINE });
	s.addText('This Week', { x: 9.869, y: 3.24, w: 1.741, h: 0.37, margin: 7.2, wrap: false, fontFace: FONT_BODY, fontSize: 16, color: GREY6 });
	s.addShape('roundRect', { x: 9.49, y: 4.022, w: 1.924, h: 0.073, rectRadius: 0.036, fill: { color: WHITE }, line: NO_LINE });
	gradPill(s, { x: 9.49, y: 4.022, w: 1.3, h: 0.073, colors: [BLUE, ORANGE] });
	s.addText('89322', { x: 9.663, y: 4.448, w: 1.573, h: 0.64, margin: 7.2, align: 'center', wrap: false, fontFace: FONT, fontSize: 32, color: GREY6 });
	s.addText('Accounts', { x: 9.824, y: 5.048, w: 1.25, h: 0.337, margin: 7.2, align: 'center', wrap: false, fontFace: FONT, fontSize: 14, color: GREY6 });
}

function slide05(s) {
	chrome(s, 5);
	s.addShape('ellipse', { x: 3.932, y: 1.015, w: 5.47, h: 5.47, fill: { type: 'none' }, line: { color: 'DCDCDC', width: 1.5 } });
	phone(s, { x: 5.287, y: 0.747, w: 2.777, h: 6.005, body: '1A1A1A', screen: CARD, radius: 0.36, label: 'Image placeholder', labelColor: 'BDBDBD' });

	// six radial callouts: pictogram + connector dot + caption
	const rows = [
		{ tx: 1.056, ty: 1.859, align: 'right', dot: [4.337, 2.027], ic: [3.636, 1.868, 0.409, 0.481] },
		{ tx: 0.676, ty: 3.501, align: 'right', dot: [3.825, 3.663], ic: [3.196, 3.533, 0.457, 0.433] },
		{ tx: 1.066, ty: 5.142, align: 'right', dot: [4.337, 5.304], ic: [3.575, 5.183, 0.481, 0.418] },
		{ tx: 10.004, ty: 1.859, align: 'left', dot: [8.811, 2.027], ic: [9.307, 1.868, 0.481, 0.481] },
		{ tx: 10.395, ty: 3.501, align: 'left', dot: [9.334, 3.663], ic: [9.772, 3.507, 0.365, 0.486] },
		{ tx: 10.015, ty: 5.142, align: 'left', dot: [9.294, 5.211], ic: [9.294, 5.211, 0.529, 0.361] },
	];
	rows.forEach((r) => {
		s.addShape('ellipse', { x: r.dot[0], y: r.dot[1], w: 0.175, h: 0.175, fill: { color: GREY4 }, line: NO_LINE });
		pictogram(s, r.ic[0], r.ic[1], r.ic[2], r.ic[3]);
		s.addText(SERENITY, {
			x: r.tx, y: r.ty, w: 2.263, h: 0.498, margin: 0, align: r.align,
			fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.3,
		});
	});
}

function slide06(s) {
	chrome(s, 6);
	phone(s, { x: 8.107, y: 2.354, w: 2.477, h: 5.385, body: '1F1F1F', screen: WALLPAPER, radius: 0.34, notch: 'bar', label: 'Image Placeholder', labelColor: 'D8D8D8' });
	phone(s, { x: 11.294, y: 0.688, w: 2.477, h: 5.385, body: '1F1F1F', screen: WALLPAPER, radius: 0.34, notch: 'bar', label: 'Image Placeholder', labelColor: 'D8D8D8' });
	headline(s, { x: 0.834, y: 2.673, w: 6.045, h: 1.414, size: 28 });
	s.addText(SERENITY_LONG, {
		x: 0.834, y: 4.424, w: 4.197, h: 0.888, margin: 0,
		fontFace: FONT_BODY, fontSize: 14, color: BODY_GREY, lineSpacingMultiple: 1.3,
	});
	gradPill(s, { x: 0.834, y: 6.032, w: 2.471, h: 0.443, colors: [ORANGE, BLUE] });
	s.addText('Click Here to Our Website', {
		x: 0.834, y: 6.032, w: 2.471, h: 0.443, margin: 0, align: 'center', valign: 'middle',
		fontFace: FONT_BODY, fontSize: 12, color: WHITE,
	});
	arrowButton(s, 3.446, 6.032, 0.443);
	s.addText('Learn More\u2026', {
		x: 4.161, y: 6.164, w: 1.105, h: 0.177, margin: 0, align: 'center',
		fontFace: FONT_BODY, fontSize: 10.5, color: BODY_GREY,
	});
}

function slide07(s) {
	chrome(s, 7);
	imgBox(s, { x: 5.109, y: 0.279, w: 3.115, h: 6.943, fill: 'E8E8E8', label: 'Drop in here your images', labelSize: 16, labelColor: 'A0A0A0' });
	// the large blue device lying diagonally across the middle
	phone(s, { x: 5.25, y: 0.82, w: 2.83, h: 6.0, rot: 125, body: '3B6FCF', screen: '4C82E4', radius: 0.5, bezel: 0.14 });
	imgBox(s, { x: 6.4, y: 3.2, w: 0.68, h: 0.68, fill: 'FFFFFF', lineColor: 'BBBBBB', label: '[img]', labelSize: 8, labelColor: '9AB6F0' });

	s.addText('Clients 2025', { x: 9.26, y: 1.449, w: 2.986, h: 0.337, margin: 7.2, fontFace: FONT_BODY, fontSize: 14, bold: true, color: BLUE });
	s.addText('Special project ', { x: 9.26, y: 1.679, w: 2.986, h: 0.513, margin: 7.2, fontFace: FONT_BODY, fontSize: 20, bold: true, color: NEAR_BLACK });
	s.addShape('line', { x: 9.363, y: 2.357, w: 0.612, h: 0, line: { color: BLUE, width: 2 } });
	s.addText('Lorem ipsum dolor sit amet, consectetur elit.', {
		x: 9.26, y: 2.568, w: 2.986, h: 0.696, margin: 7.2,
		fontFace: FONT_BODY, fontSize: 14, color: BODY_GREY, lineSpacingMultiple: 1.3,
	});
	loremCard(s, 1.087, 3.847);
}

function slide08(s) {
	chrome(s, 8);
	phone(s, { x: 1.644, y: 1.016, w: 4.921, h: 10.54, body: '2E5C8A', screen: WHITE, radius: 0.62, bezel: 0.14, notch: 'bar' });
	padlock(s, 3.9, 2.022, 0.32, SCREEN_TEXT);
	s.addText('9:41', { x: 2.6, y: 2.351, w: 2.92, h: 1.212, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 66, color: SCREEN_TEXT });
	s.addText('Wednesday, July 22', { x: 2.6, y: 3.468, w: 2.92, h: 0.337, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 14, color: SCREEN_TEXT });

	headline(s, { x: 7.706, y: 1.145, w: 4.575, h: 1.616, size: 24 });
	s.addShape('roundRect', { x: 7.706, y: 3.189, w: 3.819, h: 3.167, rectRadius: 0.245, fill: { color: CARD }, line: NO_LINE });
	s.addText('200+', { x: 8.155, y: 3.553, w: 2.92, h: 0.539, margin: 0, fontFace: FONT, fontSize: 32, color: INK });
	s.addText('Learn More', { x: 8.155, y: 4.147, w: 2.92, h: 0.236, margin: 0, fontFace: FONT, fontSize: 14, italic: true, color: BLUE });
	s.addText('PLACEHOLDER' +
		'laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi ' +
		'architecto beatae vitae dicta sunt explicabo.', {
		x: 8.155, y: 4.572, w: 2.92, h: 1.419, margin: 0,
		fontFace: FONT_BODY, fontSize: 11, color: DARK_GREY, lineSpacingMultiple: 1.3,
	});
}

function slide09(s) {
	ellipseDiag(s, { x: -4.598, y: -1.019, w: 9.537, h: 9.537, colors: [ORANGE, BLUE] });
	chrome(s, 9);
	phone(s, { x: 1.512, y: 1.21, w: 2.334, h: 5.079, rot: 345, body: '1F1F1F', screen: WALLPAPER, radius: 0.34, notch: 'bar' });
	headline(s, { x: 7.595, y: 1.274, w: 4.508, h: 1.616, size: 24 });
	imgBox(s, { x: 7.595, y: 3.662, w: 1.137, h: 0.351, fill: '111111', label: 'Google Play', labelSize: 8, labelColor: WHITE });
	imgBox(s, { x: 7.595, y: 4.071, w: 1.137, h: 0.351, fill: '111111', label: 'App Store', labelSize: 8, labelColor: WHITE });
	s.addText(SERENITY_LONG, {
		x: 7.595, y: 4.844, w: 3.708, h: 0.761, margin: 0,
		fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.3,
	});
	gradPill(s, { x: 7.595, y: 5.829, w: 1.67, h: 0.397, colors: [ORANGE, BLUE] });
	s.addText('Download Now!', {
		x: 7.595, y: 5.829, w: 1.67, h: 0.397, margin: 0, align: 'center', valign: 'middle',
		fontFace: FONT_BODY, fontSize: 10, color: WHITE,
	});
}

function slide10(s) {
	chrome(s, 10);
	// three flat phone silhouettes; only the first one has a live screen
	[2.125, 5.458, 8.792].forEach((gx, i) => {
		s.addShape('roundRect', {
			x: gx + 0.05, y: 1.045, w: 2.648, h: 5.411, rectRadius: 0.36,
			fill: { color: BODY_GREY }, line: NO_LINE,
			shadow: { type: 'outer', color: '000000', blur: 16, offset: 3, angle: 90, opacity: 0.12 },
		});
		[[0.003, 2.147, 0.1, 0.412], [0.0, 2.638, 0.1, 0.412], [0.01, 1.752, 0.079, 0.236]].forEach(([dx, y, w, h]) => {
			s.addShape('roundRect', { x: gx + dx - w, y: y, w: w * 2, h: h, rectRadius: 0.03, fill: { color: GREY4 }, line: NO_LINE });
		});
		s.addShape('roundRect', { x: gx + 2.672, y: 2.294, w: 0.146, h: 0.627, rectRadius: 0.03, fill: { color: GREY4 }, line: NO_LINE });
		if (i === 0) {
			gradDiag(s, { x: gx + 0.214, y: 1.209, w: 2.321, h: 5.083, colors: [ORANGE, BLUE], mask: BODY_GREY });
		}
	});
	padlock(s, 3.35, 2.266, 0.24, SCREEN_TEXT);
	s.addText('9:41', { x: 2.55, y: 1.991, w: 1.9, h: 0.774, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 40, color: SCREEN_TEXT });
	s.addText('Wednesday, July 22', { x: 2.55, y: 2.801, w: 1.9, h: 0.252, margin: 7.2, align: 'center', wrap: false, fontFace: FONT_BODY, fontSize: 9, color: SCREEN_TEXT });
	s.addText('30%', { x: 9.158, y: 4.901, w: 2.015, h: 0.919, margin: 7.2, fontFace: FONT, fontSize: 54, color: BODY_GREY, lineSpacingMultiple: 0.9 });
	s.addText('Your Brand Name', { x: 9.158, y: 5.69, w: 2.244, h: 0.263, margin: 7.2, fontFace: FONT, fontSize: 12, color: BODY_GREY });
}

function slide11(s) {
	gradDiag(s, { x: 0, y: 0, w: 3.342, h: 7.5, colors: [ORANGE, BLUE] });
	chrome(s, 11);
	// white wireframe handset outline
	s.addShape('roundRect', { x: 1.692, y: 1.34, w: 3.342, h: 6.818, rectRadius: 0.45, fill: { color: WHITE }, line: { color: 'D9D9D9', width: 1 } });
	s.addShape('roundRect', { x: 1.863, y: 1.512, w: 3.006, h: 6.479, rectRadius: 0.38, fill: { type: 'none' }, line: { color: 'D9D9D9', width: 1 } });
	s.addShape('roundRect', { x: 3.141, y: 1.557, w: 0.416, h: 0.068, rectRadius: 0.034, fill: { type: 'none' }, line: { color: 'D9D9D9', width: 1 } });
	s.addShape('ellipse', { x: 3.651, y: 1.543, w: 0.094, h: 0.095, fill: { type: 'none' }, line: { color: 'D9D9D9', width: 1 } });
	[[1.651, 2.261, 0.258], [1.651, 2.736, 0.474], [1.651, 3.359, 0.474]].forEach(([x, y, h]) => {
		s.addShape('roundRect', { x: x, y: y, w: 0.054, h: h, rectRadius: 0.027, fill: { color: WHITE }, line: { color: 'D9D9D9', width: 1 } });
	});
	s.addShape('roundRect', { x: 5.02, y: 2.912, w: 0.054, h: 0.745, rectRadius: 0.027, fill: { color: WHITE }, line: { color: 'D9D9D9', width: 1 } });
	arrowButton(s, 4.68, 2.981, 0.633);

	headline(s, { x: 6.667, y: 1.414, w: 4.508, h: 1.616, size: 24 });
	s.addText([
		{ text: 'Create basic ', options: { color: DARK_GREY } },
		{ text: 'UI wireframe in second', options: { color: BLUE, bold: true } },
	], { x: 6.667, y: 3.149, w: 4.071, h: 0.808, margin: 0, fontFace: FONT, fontSize: 24 });
	s.addText([
		{ text: 'A wonderful serenity has taken possession of my of spring which I enjoy with my whole heart. ', options: { breakLine: true } },
		{ text: '', options: { breakLine: true } },
		{ text: '\u201cI am alone, and feel the charm of existence in this spot, which was created mind for the bliss\u201d', options: { italic: true } },
	], {
		x: 6.667, y: 4.361, w: 4.676, h: 1.724, margin: 0,
		fontFace: FONT_BODY, fontSize: 14, color: GREY4, lineSpacingMultiple: 1.5,
	});
}

function slide12(s) {
	gradDiag(s, { x: 5.583, y: 0, w: 7.75, h: 7.5, colors: [ORANGE, BLUE] });
	chrome(s, 12);
	// four dark handsets rotated 30 degrees, each showing a drop target
	[[5.464, 2.944], [8.256, -1.892], [7.018, 6.152], [9.809, 1.316]].forEach(([x, y]) => {
		phone(s, {
			x: x, y: y, w: 2.36, h: 4.953, rot: 30, body: '0D0D0D', radius: 0.4, notch: 'drop',
			screen: 'B4B4B4', label: 'Drag and Drop Image Here', labelSize: 12, labelColor: '8C8C8C',
		});
	});
	s.addShape('ellipse', {
		x: 5.614, y: 1.09, w: 2.335, h: 2.335, fill: { color: WHITE }, line: NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 20, offset: 3, angle: 90, opacity: 0.18 },
	});
	s.addText([
		{ text: '89', options: { fontSize: 44, bold: true } },
		{ text: 'K ', options: { fontSize: 32, bold: true, breakLine: true } },
		{ text: 'Download', options: { fontSize: 14 } },
	], {
		x: 5.938, y: 1.746, w: 1.687, h: 0.979, margin: 7.2, align: 'center',
		fontFace: FONT, color: '0D0D0D', lineSpacingMultiple: 0.9,
	});
	headline(s, { x: 0.92, y: 1.238, w: 4.196, h: 2.356, size: 28 });
	downloadApp(s, 0.908, 4.461);
}

function slide13(s) {
	gradDiag(s, { x: 0, y: 0, w: 3.25, h: 7.5, colors: [BLUE, ORANGE] });
	chrome(s, 13);
	phone(s, {
		x: 2.583, y: 0.756, w: 2.875, h: 6.054, body: '111111', radius: 0.42, bezel: 0.12, notch: 'drop',
		screen: CARD, label: 'Drag and Drop Image Here', labelColor: 'A6A6A6',
	});
	s.addText('The European languages are members of the same.', {
		x: -0.634, y: 3.406, w: 3.788, h: 0.688, rotate: 270, margin: 7.2,
		fontFace: FONT, fontSize: 14, bold: true, color: WHITE, lineSpacingMultiple: 1.3,
	});
	s.addText('Our application showcase', {
		x: 5.264, y: 3.154, w: 3.996, h: 1.192, rotate: 270, margin: 7.2, align: 'center',
		fontFace: FONT, fontSize: 36, bold: true, color: INK, lineSpacingMultiple: 0.9,
	});
	[[1.859, 8.93, 0.481, 0.481], [3.501, 8.989, 0.365, 0.486], [5.142, 8.906, 0.529, 0.361]].forEach(([ty, ix, iw, ih]) => {
		pictogram(s, ix, ty + 0.009, iw, ih);
		s.addText(SERENITY, {
			x: 9.655, y: ty, w: 2.263, h: 0.498, margin: 0,
			fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.3,
		});
	});
	s.addText('Phone Mockup', { x: 0.445, y: 0.369, w: 2.416, h: 0.177, margin: 0, fontFace: FONT, fontSize: 10.5, color: WHITE });
}

function slide14(s) {
	chrome(s, 14);
	phone(s, { x: 1.097, y: 1.563, w: 2.005, h: 4.374, body: '1F1F1F', screen: WALLPAPER, radius: 0.3, notch: 'bar' });
	phone(s, { x: 10.242, y: 1.563, w: 2.005, h: 4.374, body: '1F1F1F', screen: WALLPAPER, radius: 0.3, notch: 'bar' });

	gradDiag(s, { x: 6.214, y: 1.109, w: 0.905, h: 0.854, colors: [ORANGE, BLUE], steps: 26 });
	s.addText('01', { x: 6.214, y: 1.109, w: 0.905, h: 0.854, margin: 0, align: 'center', valign: 'middle', fontFace: FONT_BODY, fontSize: 18, color: WHITE });
	s.addText('Your amazing statement', { x: 4.079, y: 2.304, w: 5.175, h: 0.404, margin: 7.2, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: BLUE });
	s.addText(LOREM_SHORT, {
		x: 4.079, y: 2.741, w: 5.175, h: 0.691, margin: 7.2, align: 'center',
		fontFace: FONT_BODY, fontSize: 14, color: DARK_GREY, lineSpacingMultiple: 1.3,
	});
	s.addShape('line', { x: 5.129, y: 3.898, w: 3.076, h: 0, line: { color: BLUE, width: 2 } });
	s.addText('5120', { x: 5.109, y: 4.363, w: 3.115, h: 0.841, margin: 7.2, align: 'center', fontFace: FONT, fontSize: 44, color: '1A1E2A' });
	s.addText('User Application', { x: 5.109, y: 5.186, w: 3.115, h: 0.404, margin: 7.2, align: 'center', fontFace: FONT, fontSize: 18, bold: true, color: BLUE });
	s.addText(LOREM_SHORT, {
		x: 4.079, y: 5.7, w: 5.175, h: 0.691, margin: 7.2, align: 'center',
		fontFace: FONT_BODY, fontSize: 14, color: DARK_GREY, lineSpacingMultiple: 1.3,
	});
}

/** donut gauge used three times on slide 15 — arcs are given in degrees */
function donutGauge(s, x, y, arcs, label) {
	s.addShape('ellipse', {
		x: x + 0.357, y: y, w: 1.659, h: 1.659, fill: { color: WHITE }, line: NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 14, offset: 2, angle: 90, opacity: 0.15 },
	});
	s.addShape('triangle', { x: x + 0.06, y: y + 0.53, w: 0.322, h: 0.442, rotate: 278.4, fill: { color: WHITE }, line: NO_LINE });
	arcs.forEach(([a1, a2, color, faded]) => {
		s.addShape('pie', {
			x: x + 0.519, y: y + 0.168, w: 1.324, h: 1.324, angleRange: [a1, a2],
			fill: { color: color, transparency: faded ? 80 : 0 }, line: NO_LINE,
		});
	});
	s.addShape('ellipse', { x: x + 0.587, y: y + 0.236, w: 1.187, h: 1.187, fill: { color: WHITE }, line: NO_LINE });
	s.addText(label, { x: x + 0.606, y: y + 0.474, w: 1.15, h: 0.505, margin: 7.2, align: 'center', fontFace: FONT, color: NEAR_BLACK });
	s.addText('On value', { x: x + 0.606, y: y + 0.878, w: 1.15, h: 0.286, margin: 7.2, align: 'center', fontFace: FONT_BODY, fontSize: 11, color: GREY4 });
}

function slide15(s) {
	chrome(s, 15);
	const ARCS_A = [[223.0, 266.3, BLUE, true], [119.3, 225.5, BLUE, false], [42.4, 110.2, ORANGE, true],
		[2.4, 43.4, ORANGE, false], [312.7, 356.4, GREY3, true], [271.9, 314.6, GREY3, false]];
	const ARCS_B = [[183.9, 266.3, BLUE, true], [36.2, 185.1, BLUE, false], [356.2, 391.0, ORANGE, true],
		[321.2, 356.7, ORANGE, false], [291.8, 315.3, GREY3, true], [271.9, 293.2, GREY3, false]];

	[[2.128, '24', ARCS_A, 2.593], [5.544, '18K', ARCS_B, 6.047], [8.959, '520', ARCS_A, 9.43]].forEach(([px, val, arcs, gx]) => {
		phone(s, {
			x: px, y: 1.306, w: 2.246, h: 4.887, rot: 330, body: '1F1F1F', radius: 0.32, notch: 'bar',
			screen: WALLPAPER, label: 'Image Placeholder', labelColor: 'E3E3E3',
		});
		const label = val === '24'
			? [{ text: '24', options: { fontSize: 24 } }, { text: '%', options: { fontSize: 20 } }]
			: [{ text: val, options: { fontSize: 24 } }];
		donutGauge(s, gx, 1.071, arcs, label);
	});
	[1.705, 5.504, 9.303].forEach((x) => {
		s.addShape('roundRect', {
			x: x, y: 4.77, w: 2.325, h: 1.079, rectRadius: 0.123, fill: { color: WHITE }, line: NO_LINE,
			shadow: { type: 'outer', color: '000000', blur: 14, offset: 2, angle: 90, opacity: 0.15 },
		});
		s.addText('Clients 2025', { x: x + 0.108, y: 4.9, w: 2.108, h: 0.337, margin: 7.2, fontFace: FONT_BODY, fontSize: 14, bold: true, color: BLUE });
		s.addText('Special project ', { x: x + 0.108, y: 5.16, w: 2.108, h: 0.513, margin: 7.2, wrap: false, fontFace: FONT_BODY, fontSize: 20, bold: true, color: NEAR_BLACK });
	});
}

function slide16(s) {
	ellipseDiag(s, { x: -0.561, y: 3.75, w: 14.456, h: 11.166, colors: [ORANGE, BLUE] });
	chrome(s, 16);
	phone(s, { x: 3.251, y: 4.734, w: 2.53, h: 5.42, rot: 345, body: '1A1A1A', screen: 'FFFFFF', radius: 0.36, notch: 'bar' });
	phone(s, { x: 7.573, y: 4.734, w: 2.53, h: 5.42, rot: 15, body: '1A1A1A', screen: 'FFFFFF', radius: 0.36, notch: 'bar' });
	phone(s, { x: 4.986, y: 3.001, w: 3.361, h: 7.2, body: '1A1A1A', screen: 'FFFFFF', radius: 0.45, notch: 'bar' });
	headline(s, { x: 2.98, y: 0.772, w: 7.373, h: 1.414, size: 28, align: 'center' });
}

function slide17(s) {
	gradDiag(s, { x: 0, y: 0, w: 4.584, h: 7.5, colors: [BLUE, ORANGE] });
	chrome(s, 17);
	s.addShape('ellipse', { x: 1.216, y: 4.273, w: 2.237, h: 2.237, fill: { color: GREY4, transparency: 80 }, line: NO_LINE });
	phone(s, { x: -0.241, y: -0.382, w: 2.477, h: 5.385, body: '1F1F1F', screen: WALLPAPER, radius: 0.34, notch: 'bar', label: 'Image Placeholder', labelColor: '9AB6F0' });
	phone(s, { x: 2.947, y: 1.062, w: 2.477, h: 5.385, body: '1F1F1F', screen: WALLPAPER, radius: 0.34, notch: 'bar', label: 'Image Placeholder', labelColor: '9AB6F0' });

	headline(s, { x: 6.552, y: 1.199, w: 5.316, h: 1.885, size: 28 });
	s.addText('A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring', {
		x: 6.552, y: 3.48, w: 5.43, h: 0.582, margin: 0,
		fontFace: FONT_BODY, fontSize: 14, color: BODY_GREY, lineSpacingMultiple: 1.3,
	});
	[[5.352, '96%'], [8.539, '234+']].forEach(([x, big]) => {
		s.addShape('roundRect', {
			x: x, y: 4.59, w: 3.055, h: 1.711, rectRadius: 0.15, fill: { color: WHITE }, line: NO_LINE,
			shadow: { type: 'outer', color: '000000', blur: 16, offset: 2, angle: 90, opacity: 0.13 },
		});
		s.addText(big, { x: x + 0.283, y: 4.906, w: 1.831, h: 0.424, margin: 7.2, fontFace: FONT, fontSize: 24, bold: true, color: BLUE, lineSpacingMultiple: 0.8 });
		s.addText('A wonderful serenity has taken possession of my', {
			x: x + 0.283, y: 5.382, w: 2.488, h: 0.602, margin: 7.2,
			fontFace: FONT_BODY, fontSize: 12, color: BODY_GREY, lineSpacingMultiple: 1.3,
		});
	});
}

function slide18(s) {
	gradDiag(s, { x: 0, y: 0, w: 2.294, h: 7.5, colors: [BLUE, ORANGE] });
	gradDiag(s, { x: 11.039, y: 0, w: 2.294, h: 7.5, colors: [BLUE, ORANGE] });
	chrome(s, 18);
	phone(s, { x: 1.086, y: 1.563, w: 2.005, h: 4.374, body: '1F1F1F', screen: WALLPAPER, radius: 0.3, notch: 'bar', label: 'Drag & drop image here', labelSize: 11, labelColor: 'E8E8E8' });
	phone(s, { x: 10.242, y: 1.563, w: 2.005, h: 4.374, body: '1F1F1F', screen: WALLPAPER, radius: 0.3, notch: 'bar', label: 'Drag & drop image here', labelSize: 11, labelColor: 'E8E8E8' });

	[[1.421, BLUE], [4.184, ORANGE]].forEach(([y, color]) => {
		imgBox(s, { x: 6.355, y: y, w: 0.624, h: 0.624, fill: color, label: null });
	});
	[[2.187, 2.625], [4.95, 5.388]].forEach(([ty, by]) => {
		s.addText('Your amazing statement', { x: 4.079, y: ty, w: 5.175, h: 0.438, margin: 7.2, align: 'center', fontFace: FONT, fontSize: 20, bold: true, color: NEAR_BLACK });
		s.addText(LOREM_SHORT, {
			x: 4.079, y: by, w: 5.175, h: 0.691, margin: 7.2, align: 'center',
			fontFace: FONT_BODY, fontSize: 14, color: BODY_GREY, lineSpacingMultiple: 1.3,
		});
	});
	s.addShape('line', { x: 4.498, y: 3.75, w: 4.337, h: 0, line: { color: BLUE, width: 2 } });
}

function slide19(s) {
	gradDiag(s, { x: 0, y: 0, w: 4.584, h: 7.5, colors: [BLUE, ORANGE] });
	chrome(s, 19);
	// oversized white handset lying at an angle across the left half
	phone(s, { x: 2.545, y: 0.51, w: 3.25, h: 6.9, rot: 113, body: 'F0F0F0', screen: 'FBFBFB', radius: 0.55 });
	s.addText('Drag and drop picture', {
		x: 1.827, y: -0.486, w: 3.914, h: 8.473, margin: 7.2, align: 'center', valign: 'middle',
		fontFace: FONT_BODY, fontSize: 16, color: 'A6A6A6',
	});
	s.addShape('roundRect', { x: 1.485, y: 1.06, w: 1.033, h: 1.033, rectRadius: 0.1, fill: { color: BLUE, transparency: 88 }, line: NO_LINE });

	// floating "This Week" statistics card
	s.addShape('roundRect', {
		x: 3.38, y: 4.703, w: 2.408, h: 1.896, rectRadius: 0.125, fill: { color: WHITE }, line: NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 18, offset: 3, angle: 90, opacity: 0.15 },
	});
	s.addText('This Week', { x: 3.622, y: 4.89, w: 1.022, h: 0.269, margin: 0, fontFace: FONT_BODY, fontSize: 16, color: GREY6 });
	s.addShape('roundRect', { x: 3.622, y: 5.369, w: 1.924, h: 0.073, rectRadius: 0.036, fill: { color: CARD }, line: NO_LINE });
	gradPill(s, { x: 3.622, y: 5.369, w: 1.3, h: 0.073, colors: [BLUE, ORANGE] });
	s.addText('89322', { x: 3.622, y: 5.646, w: 1.573, h: 0.539, margin: 0, fontFace: FONT, fontSize: 32, color: BLUE });
	s.addText('Accounts', { x: 3.622, y: 6.177, w: 1.25, h: 0.236, margin: 0, fontFace: FONT, fontSize: 14, color: GREY6 });

	headline(s, { x: 6.957, y: 1.194, w: 5.055, h: 1.885, size: 28 });
	downloadApp(s, 9.027, 4.505);
}

function slide20(s) {
	// this layout hides the master furniture and paints its own on the gradient
	gradBand(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, colors: [BLUE, ORANGE], dir: 'v', steps: 60 });
	chrome(s, 20, { tagColor: WHITE, numColor: WHITE, pillColor: 'FAFAFA' });
	s.addShape('rect', { x: 0.768, y: 4.6, w: 7.141, h: 1.457, fill: { color: WHITE }, line: NO_LINE });
	s.addText('Thank You!', {
		x: 0.987, y: 4.723, w: 5.571, h: 1.212, margin: 0,
		fontFace: FONT, fontSize: 72, bold: true, color: BLUE,
	});
	s.addShape('rect', {
		x: 6.558, y: 6.042, w: 3.37, h: 0.5, fill: { color: WHITE }, line: NO_LINE,
		shadow: { type: 'outer', color: '000000', blur: 12, offset: 2, angle: 90, opacity: 0.13 },
	});
	s.addText('Infographic Presentation Template', {
		x: 6.722, y: 6.191, w: 3.043, h: 0.202, margin: 0, fontFace: FONT, fontSize: 12, color: INK,
	});
	s.addText('Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor congue massa. ', {
		x: 0.987, y: 3.414, w: 3.452, h: 0.569, margin: 0,
		fontFace: FONT_BODY, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5,
	});
}

/* ------------------------------------------------------------------- main */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
	slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pres = new pptxgen();
pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'WIDE';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT_BODY };
pres.title = 'Phone Mockup — Infographic Presentation Template';

BUILDERS.forEach((build) => {
	const slide = pres.addSlide();
	slide.background = { color: WHITE };
	build(slide);
});

pres.writeFile({ fileName: path.join(__dirname, '09630f77-fb57-4a91-b82a-803f404d776a_grok_final.pptx') })
	.then((f) => console.log('wrote ' + f));
