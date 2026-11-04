/**
 * Ophthalmology Clinic — deck rebuilt with pptxgenjs.
 * Slide size 20 x 11.25 in (16:9). Photographs in the source deck are
 * re-created here as flat grey placeholder blocks, icons as outline tiles.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */
const CYAN = '2BE9FD';       // primary accent
const TEAL = '02C0D4';       // darker accent / kicker text
const WHITE = 'FFFFFF';
const OFFWHITE = 'F2F2F2';
const INK = '000000';        // headline text
const GREY = '595959';       // body text
const PHOTO = '5B5B5B';      // stand-in for dark photographs
const PHOTO_LT = 'C3C3C3';   // stand-in for light photographs
const MOCKUP = 'D9D9D9';     // stand-in for device mock-ups
const HEAD = 'Manrope ExtraBold';
const BODY = 'Mulish';

// 45-degree drop shadows; blur/offset are points. pptxgenjs mutates the options
// object it is handed, so every shape needs its own copy — hence the factories.
const shadow = (blur, offset, color, opacity = 0.4) => ({ type: 'outer', blur, offset, angle: 45, color, opacity });
const cardShadow = () => shadow(20, 10, INK);     // large white content cards
const tileShadow = () => shadow(20, 5, GREY);     // small icon tiles
const panelShadow = () => shadow(28, 15, GREY);   // full-width panels

/* ---------------------------------------------------------------- helpers */
const S = (s, shape, x, y, w, h, o = {}) => s.addShape(shape, Object.assign({ x, y, w, h }, o));
const T = (s, str, x, y, w, h, o = {}) =>
	s.addText(str, Object.assign({ x, y, w, h, valign: 'top', fontFace: BODY, fontSize: 18, color: GREY }, o));

const TITLE = { fontFace: HEAD, fontSize: 37, color: INK };
const HEADING = { fontFace: HEAD, fontSize: 20, bold: true, color: INK };
const LEAD = { lineSpacingMultiple: 1.5 };
const BTN_TEXT = { fontFace: HEAD, fontSize: 18, bold: true, color: WHITE, align: 'center' };
/** Bold heading style at a non-default size. */
const heading = (size, extra) => Object.assign({}, HEADING, { fontSize: size }, extra);

/** Flat grey block standing in for a photograph; radius 0 gives square corners. */
const photo = (s, x, y, w, h, radius = 0.25, color = PHOTO) =>
	radius ? S(s, 'roundRect', x, y, w, h, { fill: { color }, rectRadius: radius })
		: S(s, 'rect', x, y, w, h, { fill: { color } });

/**
 * Line pictograms, each drawn inside a unit square that `icon()` scales.
 * Coordinates below are fractions of the icon box.
 */
const PICTOGRAMS = {
	eye: (o, ln, fill) => { o('ellipse', 0.08, 0.30, 0.84, 0.40, ln); o('ellipse', 0.38, 0.36, 0.24, 0.28, fill); },
	cross: (o, ln) => { o('ellipse', 0.06, 0.06, 0.88, 0.88, ln); o('mathPlus', 0.24, 0.24, 0.52, 0.52, ln); },
	list: (o, ln, fill) => {
		[0.24, 0.46, 0.68].forEach((ty, i) => {
			o('ellipse', 0.10, ty, 0.10, 0.10, fill);
			o('rect', 0.30, ty + 0.03, [0.58, 0.42, 0.30][i], 0.05, fill);
		});
	},
	chat: (o, ln) => { o('roundRect', 0.06, 0.12, 0.88, 0.56, ln); o('heart', 0.34, 0.24, 0.32, 0.30, ln); o('triangle', 0.30, 0.66, 0.22, 0.22, Object.assign({ rotate: 200 }, ln)); },
	mortar: (o, ln) => { o('chord', 0.14, 0.40, 0.72, 0.44, Object.assign({ angleRange: [0, 180] }, ln)); o('rect', 0.46, 0.80, 0.08, 0.12, ln); o('rect', 0.52, 0.14, 0.06, 0.30, Object.assign({ rotate: 20 }, ln)); },
	pill: (o, ln) => { o('roundRect', 0.10, 0.38, 0.80, 0.26, Object.assign({ rotate: -40, rectRadius: 0.13 }, ln)); },
	bed: (o, ln) => { o('rect', 0.08, 0.50, 0.84, 0.26, ln); o('roundRect', 0.14, 0.34, 0.28, 0.16, ln); o('rect', 0.56, 0.18, 0.34, 0.26, ln); o('rect', 0.08, 0.76, 0.06, 0.14, ln); o('rect', 0.86, 0.76, 0.06, 0.14, ln); },
	bulb: (o, ln) => {
		o('ellipse', 0.26, 0.16, 0.48, 0.48, ln); o('rect', 0.40, 0.66, 0.20, 0.18, ln);
		[[0.50, 0.00, 0.00, 0.10], [0.14, 0.16, 0.09, 0.09], [0.77, 0.16, 0.09, 0.09],
			[0.02, 0.40, 0.10, 0.00], [0.88, 0.40, 0.10, 0.00]].forEach(([lx, ly, lw, lh]) => o('line', lx, ly, lw, lh, ln));
	},
	percent: (o, ln, fill, text) => { text('%', 0.05, 0.10, 0.90, 0.80); },
	tree: (o, ln) => {
		o('rect', 0.38, 0.08, 0.24, 0.20, ln);
		[0.06, 0.38, 0.70].forEach(tx => o('rect', tx, 0.68, 0.24, 0.20, ln));
		o('line', 0.50, 0.28, 0.00, 0.22, ln); o('line', 0.18, 0.50, 0.64, 0.00, ln);
		[0.18, 0.82].forEach(tx => o('line', tx, 0.50, 0.00, 0.18, ln));
	},
	shield: (o, ln) => { o('pentagon', 0.14, 0.06, 0.72, 0.88, Object.assign({ rotate: 180 }, ln)); },
	pin: (o, ln, fill) => { o('teardrop', 0.24, 0.06, 0.52, 0.52, Object.assign({ rotate: 135 }, ln)); o('ellipse', 0.42, 0.24, 0.16, 0.16, fill); o('rect', 0.30, 0.80, 0.40, 0.06, fill); },
	phone: (o, ln, fill) => {
		o('chord', 0.10, 0.28, 0.80, 0.44, Object.assign({ angleRange: [180, 360] }, ln));
		[0.06, 0.70].forEach(lx => o('roundRect', lx, 0.46, 0.24, 0.18, Object.assign({ rectRadius: 0.06 }, fill)));
		o('rect', 0.16, 0.74, 0.68, 0.07, fill);
	},
	globe: (o, ln) => { o('ellipse', 0.06, 0.06, 0.88, 0.88, ln); o('ellipse', 0.34, 0.06, 0.32, 0.88, ln); o('line', 0.06, 0.50, 0.88, 0.00, ln); },
	camera: (o, ln, fill) => { o('roundRect', 0.08, 0.08, 0.84, 0.84, ln); o('ellipse', 0.30, 0.30, 0.40, 0.40, ln); o('ellipse', 0.72, 0.16, 0.10, 0.10, fill); },
	note: (o, ln, fill) => { o('ellipse', 0.14, 0.62, 0.32, 0.26, fill); o('rect', 0.42, 0.12, 0.08, 0.62, fill); o('rect', 0.50, 0.12, 0.26, 0.08, fill); },
};

/** Draw pictogram `name` in the size x size box at (x, y), stroked/filled in `color`. */
function icon(s, name, x, y, size, color) {
	const w = Math.max(size * 0.05, 1.1);
	const shape = (sh, fx, fy, fw, fh, opt) => S(s, sh, x + fx * size, y + fy * size, fw * size, fh * size,
		Object.assign({}, opt, opt.rectRadius ? { rectRadius: opt.rectRadius * size } : {}));
	const stroke = { line: { color, width: w } };
	const solid = { fill: { color } };
	const text = (str, fx, fy, fw, fh) => T(s, str, x + fx * size, y + fy * size, fw * size, fh * size,
		{ fontFace: HEAD, fontSize: Math.round(size * 44), bold: true, color, align: 'center', valign: 'middle' });
	PICTOGRAMS[name](shape, stroke, solid, text);
}

/** Round "eye" logo mark; (x,y) is the top-left of its 0.469in disc. */
function logo(s, x, y) {
	const cx = x + 0.2345, cy = y + 0.2345;
	S(s, 'ellipse', x, y, 0.469, 0.469, { fill: { color: CYAN } });
	S(s, 'ellipse', cx - 0.15, cy - 0.085, 0.30, 0.17, { line: { color: WHITE, width: 1.75 } });
	S(s, 'ellipse', cx - 0.04, cy - 0.04, 0.08, 0.08, { fill: { color: WHITE } });
	S(s, 'rect', x - 0.08, cy - 0.018, 0.629, 0.036, { fill: { color: WHITE } });
}

/** Logo mark plus its all-caps section label. */
function kicker(s, x, y, label, w = 2.747) {
	logo(s, x, y);
	T(s, label, x + 0.64, y + 0.032, w, 0.404, { fontFace: HEAD, fontSize: 18, color: TEAL });
}

/** Large "scan the eye" pictogram used on the cyan discs. */
function scanEye(s, x, y) {
	const w = 0.77, h = 0.568, t = 0.036, arm = 0.128;
	[[0, 0, 1, 1], [1, 0, -1, 1], [0, 1, 1, -1], [1, 1, -1, -1]].forEach(([fx, fy, sx, sy]) => {
		const cx = x + fx * (w - t), cy = y + fy * (h - t);
		S(s, 'rect', sx > 0 ? cx : cx - arm + t, cy, arm, t, { fill: { color: OFFWHITE } });
		S(s, 'rect', cx, sy > 0 ? cy : cy - arm + t, t, arm, { fill: { color: OFFWHITE } });
	});
	S(s, 'ellipse', x + 0.175, y + 0.18, 0.42, 0.21, { line: { color: OFFWHITE, width: 2 } });
	S(s, 'ellipse', x + 0.325, y + 0.225, 0.12, 0.12, { fill: { color: OFFWHITE } });
	S(s, 'rect', x, y + h / 2 - 0.018, w, 0.036, { fill: { color: OFFWHITE } });
}

/** Elbow connector: horizontal run plus a vertical drop on its left end. */
function elbow(s, x, y, w, h, hTop) {
	S(s, 'line', x, hTop ? y : y + h, w, 0, { line: { color: CYAN, width: 2.25 } });
	S(s, 'line', x, y, 0, h, { line: { color: CYAN, width: 2.25 } });
}

/** Bezelled screen: light outer frame, dark bezel, grey "photo" inside. */
function screen(s, x, y, w, h, radius, inset) {
	S(s, 'roundRect', x, y, w, h, { fill: { color: MOCKUP }, rectRadius: radius });
	S(s, 'roundRect', x + inset, y + inset, w - 2 * inset, h - 2 * inset,
		{ fill: { color: '111111' }, rectRadius: Math.max(radius - inset, 0.02) });
}

/** Donut progress ring with a percentage in the middle. */
function donut(s, x, y, label) {
	const d = 1.738, inner = 1.168;
	S(s, 'ellipse', x, y, d, d, { fill: { color: OFFWHITE } });
	S(s, 'arc', x, y, d, d, { fill: { color: CYAN }, angleRange: [334.6, 286.6] });
	S(s, 'ellipse', x + (d - inner) / 2, y + (d - inner) / 2, inner, inner, { fill: { color: WHITE } });
	T(s, label, x - 0.178, y + 0.651, 1.381, 0.438, heading(20, { align: 'center' }));
}

const mix = (a, b, t) => {
	const c = i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
	return [c(0), c(2), c(4)].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
};

/** Left-to-right gradient painted as narrow bands (pptxgenjs only emits solid fills). */
function gradient(s, x, y, w, h, from, to, steps = 24) {
	for (let i = 0; i < steps; i++) {
		S(s, 'rect', x + (w * i) / steps, y, w / steps + 0.01, h, { fill: { color: mix(from, to, i / (steps - 1)) } });
	}
}

/** Same idea for a rounded panel: nested roundRects sharing the left edge keep the corners clean. */
function gradientPanel(s, x, y, w, h, from, to, radius, steps = 24) {
	for (let i = steps; i >= 1; i--) {
		S(s, 'roundRect', x, y, (w * i) / steps, h,
			{ fill: { color: mix(from, to, (i - 0.5) / steps) }, rectRadius: radius });
	}
}

/* --------------------------------------------------- shared layout chrome */
/** Cyan-to-teal page background used by the opening / profile / closing pages. */
function cyanBackdrop(s) {
	S(s, 'rect', 0, 0, 10, 11.25, { fill: { color: TEAL } });
	gradient(s, 10, 0, 10, 11.25, TEAL, CYAN, 20);
}

/** Translucent bubbles + sparkles scattered over the cyan backdrops. */
function bubbles(s, discs, sparkles) {
	discs.forEach(([x, y, d, solid]) =>
		S(s, 'ellipse', x, y, d, d, { fill: { color: solid ? CYAN : WHITE, transparency: solid ? 0 : 65 } }));
	sparkles.forEach(([x, y, w, h]) => S(s, 'star4', x, y, w, h, { fill: { color: WHITE } }));
}

/* ------------------------------------------------------------------ pages */
function slide01(s) {
	cyanBackdrop(s);
	bubbles(s, [[15.318, 2.273, 1.817, true], [15.318, 0.456, 0.909], [17.892, 3.182, 0.909],
		[18.044, 1.061, 1.212], [1.578, 1.45, 0.498, true]],
	[[17.381, 2.22, 0.396, 0.472], [16.881, 0.946, 0.396, 0.472], [14.381, 4.215, 0.396, 0.472]]);

	T(s, 'Ophthalmology Clinic', 1.582, 1.91, 16.836, 1.767, { fontFace: HEAD, fontSize: 99, bold: true, color: WHITE });
	T(s, 'Ophthalmology Clinic Presentation Template', 1.582, 4.094, 7.717, 0.522, { fontSize: 25, color: OFFWHITE });

	S(s, 'rect', 0, 6.388, 20, 4.862, { fill: { color: PHOTO_LT } });
	S(s, 'ellipse', 1.582, 8.081, 1.628, 1.628, { fill: { color: CYAN } });
	scanEye(s, 2.011, 8.611);

	elbow(s, 2.396, 7.298, 1.992, 0.62, true);
	elbow(s, 3.585, 8.268, 0.952, 0.62, true);
	elbow(s, 3.585, 8.902, 0.952, 0.62, false);
	elbow(s, 2.396, 9.873, 1.992, 0.62, false);
	[7.203, 8.159, 9.398, 10.378].forEach(y => S(s, 'ellipse', 4.353, y, 0.22, 0.22, { fill: { color: CYAN } }));

	[['Cataract and refractive surgery (LASIK))', 7.064, 5.097],
		['Retina and cornea treatments', 8.001, 4.634],
		['Pediatric ophthalmology', 9.249, 4.634],
		['Contact lenses and optical services', 10.188, 5.097],
	].forEach(([str, y, w]) => T(s, str, 4.903, y, w, 0.404, LEAD));
}

function slide02(s) {
	S(s, 'roundRect', 18.088, 1.93, 0.28, 6.861, { fill: { color: CYAN }, rectRadius: 0.14 });
	S(s, 'ellipse', 18.088, 9.05, 0.27, 0.27, { fill: { color: CYAN } });

	kicker(s, 1.726, 2.176, 'ABOUT US');
	T(s, 'Innovative Eye Care for a Clearer Tomorrow', 1.631, 2.909, 8.013, 1.347, TITLE);
	T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et ' +
		'dolore magna aliqua. Cursus risus at ultrices mi. Duis ultricies lacus sed turpis tincidunt aliquet sed ' +
		'euismod porta lorem mollis aliquam .Bacon ipsum dolor amet short ribs brisket venison rump drumstick pig ' +
		'sausage prosciutto chicken spare ribs salami picanha doner. Kevin capicola sausage, buffalo bresaola venison.',
	1.631, 4.893, 8.588, 2.677, LEAD);

	S(s, 'roundRect', 1.726, 8.135, 3.296, 0.939, { fill: { color: CYAN }, rectRadius: 0.469 });
	T(s, 'LEARN MORE', 1.903, 8.403, 2.942, 0.404, BTN_TEXT);

	photo(s, 11.738, 1.93, 6.115, 7.391);
}

function slide03(s) {
	S(s, 'rect', 7.229, 0, 0.417, 8.039, { fill: { color: CYAN } });
	S(s, 'rect', 7.229, 9.635, 0.417, 1.615, { fill: { color: CYAN } });

	photo(s, 0, 0, 7.254, 11.25, 0);
	S(s, 'roundRect', 6.385, 8.039, 2.104, 1.596, { fill: { color: CYAN }, rectRadius: 0.2 });
	S(s, 'custGeom', 6.7, 8.55, 1.48, 0.6, {
		line: { color: WHITE, width: 3.5 },
		points: [{ x: 0, y: 0.34 }, { x: 0.36, y: 0.34 }, { x: 0.52, y: 0.02 }, { x: 0.8, y: 0.58 }, { x: 0.98, y: 0.34 }, { x: 1.48, y: 0.34 }],
	});

	kicker(s, 10.326, 1.501, 'ABOUT US');
	T(s, 'Comfort, Care, & Personalized Treatment', 10.246, 2.238, 8.011, 1.347, TITLE);
	T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing the elit, sed tempor incididunt ut labore et dolore ' +
		'magna at aliqua. Cursus risus at ultrices mi. Duis ultricies lacus sed turpis tincidunt aliquet sed euismod ' +
		'porta lorem mollis aliquam. turkey shoulder picanha ham pork tri-tip meatball. Doner spare ribs andouille ' +
		'bacon sausage. Ground round jerky brisket shank.', 10.246, 4.21, 8.693, 2.222, LEAD);
	T(s, 'WRITE TITLE HERE', 10.336, 7.08, 4.634, 0.438, heading(20, LEAD));

	[[10.336, 'eye', true], [12.76, 'list', false], [15.184, 'cross', false]].forEach(([x, glyph, filled]) => {
		S(s, 'roundRect', x, 8.152, 1.906, 1.596, filled
			? { fill: { color: CYAN }, rectRadius: 0.166, shadow: tileShadow() }
			: { rectRadius: 0.166, line: { color: CYAN, width: 1.5 } });
		icon(s, glyph, x + 0.453, 8.298, 1.0, filled ? WHITE : CYAN);
	});
}

function slide04(s) {
	T(s, 'Comprehensive Eye Care Under One Roof', 3.367, 1.252, 13.272, 0.774,
		Object.assign({}, TITLE, { align: 'center', fontSize: 40 }));

	const LOREM_L = 'Lorem ipsum dolor sit amet, consectetur adipiscing the elit, sed do eiusmod tempor.';
	const LOREM_S = 'Lorem ipsum dolor sit amet, consectetur the elit, sed do eiusmod tempor.';
	T(s, 'WRITE TITLE HERE', 1.835, 3.269, 3.267, 0.554, heading(20, LEAD));
	T(s, LOREM_L, 1.835, 3.914, 3.826, 1.417, LEAD);
	T(s, 'WRITE TITLE HERE', 14.904, 3.269, 3.261, 0.554, heading(20, { align: 'right', lineSpacingMultiple: 1.5 }));
	T(s, LOREM_L, 14.339, 3.914, 3.826, 1.417, Object.assign({}, LEAD, { align: 'right' }));
	T(s, 'WRITE TITLE HERE', 3.762, 7.937, 3.826, 0.554, heading(20, { align: 'center', lineSpacingMultiple: 1.5 }));
	T(s, LOREM_S, 3.762, 8.581, 3.826, 1.417, Object.assign({}, LEAD, { align: 'center' }));
	T(s, 'WRITE TITLE HERE', 12.419, 7.937, 3.826, 0.554, heading(20, { align: 'center', lineSpacingMultiple: 1.5 }));
	T(s, LOREM_S, 12.419, 8.581, 3.826, 1.417, Object.assign({}, LEAD, { align: 'center' }));

	const wire = (segs) => segs.forEach(([x, y, w, h]) => S(s, 'line', x, y, w, h, { line: { color: CYAN, width: 2.25 } }));
	wire([[13.49, 2.655, 0, 1.683], [12.656, 3.497, 0, 2.8], [12.671, 6.297, 1.632, 0], [14.288, 6.297, 0, 1.161]]);
	wire([[6.502, 2.655, 0, 1.683], [7.337, 3.497, 0, 2.8], [5.713, 6.297, 1.638, 0], [5.72, 6.297, 0, 1.161]]);
	[[14.13, 3.334], [14.13, 7.205], [5.548, 3.333], [5.548, 7.205]]
		.forEach(([x, y]) => S(s, 'ellipse', x, y, 0.328, 0.328, { fill: { color: CYAN } }));

	photo(s, 7.79, 3.411, 4.427, 4.427);
}

function slide05(s) {
	S(s, 'roundRect', 10.435, 4.712, 1.906, 1.489, { rectRadius: 0.155, line: { color: CYAN, width: 1.5 }, shadow: tileShadow() });

	S(s, 'roundRect', 1.896, 1.93, 6.431, 7.391, { fill: { color: WHITE }, rectRadius: 0.425, shadow: cardShadow() });
	S(s, 'roundRect', 1.789, 2.911, 0.3, 5.429, { fill: { color: CYAN }, rectRadius: 0.15 });
	T(s, '150k', 2.655, 2.789, 2.549, 0.916, heading(36, LEAD));
	T(s, 'Lorem ipsum dolor sit amet, elit, sed et tempor ut labore dolor amet short ribs brisket venison rump ' +
		'sausage in spare ribs salami picanha doner kevin.', 2.655, 3.839, 4.872, 1.767, LEAD);
	T(s, '25+', 2.655, 5.994, 2.549, 0.916, heading(36, LEAD));
	T(s, 'Lorem ipsum dolor sit amet, elit, sed et tempor ut labore dolor amet short ribs brisket venison rump ' +
		'sausage spare.', 2.655, 7.044, 4.872, 1.313, LEAD);

	kicker(s, 10.425, 1.662, 'ABOUT US');
	T(s, 'State-of-the-Art Equipment for Accurate Results', 10.345, 2.398, 7.866, 1.347, TITLE);
	photo(s, 10.621, 4.886, 1.536, 1.14, 0.1);
	S(s, 'line', 12.752, 5.215, 1.308, 0, { line: { color: CYAN, width: 1.5 } });
	S(s, 'ellipse', 13.94, 5.072, 0.287, 0.287, { fill: { color: CYAN } });
	T(s, 'Cataract and refractive surgery (LASIK))', 14.45, 4.93, 3.292, 0.774, Object.assign({}, LEAD, { fontSize: 16 }));
	T(s, 'WRITE TITLE HERE', 10.345, 6.747, 4.634, 0.554, heading(20, LEAD));
	T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing the elit, sed do eiusmod tempor incididunt ut labore ' +
		'et dolore magna at aliqua. Cursus risus at ultrices mi. Duis ultricies lacus sed turpis tincidunt aliquet ' +
		'sed euismod porta lorem mollis aliquam.', 10.345, 7.52, 7.996, 1.767, LEAD);
}

function slide06(s) {
	photo(s, 10.0, 1.126, 8.32, 3.162);
	kicker(s, 1.498, 1.406, 'ABOUT US');
	T(s, 'Who We Are and What We Stand For', 1.418, 2.142, 6.328, 1.347, TITLE);

	S(s, 'roundRect', 1.538, 5.128, 17.045, 4.996, { fill: { color: WHITE }, rectRadius: 0.33, shadow: cardShadow() });
	S(s, 'roundRect', 1.417, 6.061, 0.3, 3.129, { fill: { color: CYAN }, rectRadius: 0.15 });
	S(s, 'roundRect', 18.132, 1.494, 0.3, 2.426, { fill: { color: CYAN }, rectRadius: 0.15 });

	const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod et tempor incididunt ut ' +
		'labore. Bacon iamet ribs brisket venison rump drumstick pig sausage prosciutto chicken spare ribs  doner. ' +
		'Kevin capicola sausage, buffalo.';
	T(s, 'MISSION: PROVIDING WORLD-CLASS EYE CARE', 2.742, 5.77, 8.616, 0.554, heading(20, LEAD));
	T(s, COPY, 2.742, 6.443, 14.402, 0.859, LEAD);
	T(s, 'VISION: TO BE THE LEADER IN ADVANCED OPHTHALMIC SERVICES', 2.742, 7.846, 10.215, 0.438, heading(20, LEAD));
	T(s, COPY, 2.742, 8.519, 14.402, 0.859, LEAD);
}

function slide07(s) {
	S(s, 'roundRect', 16.481, 0.761, 2.736, 3.701, { fill: { color: CYAN }, rectRadius: 0.327 });
	S(s, 'ellipse', 18.715, 4.634, 0.503, 0.503, { fill: { color: CYAN } });

	photo(s, 1.433, 1.287, 5.279, 6.46);
	T(s, 'Dr. EDGAR RICO', 1.368, 8.385, 2.747, 0.505, { fontFace: HEAD, fontSize: 24, bold: true, color: INK });
	T(s, 'Job Position', 1.368, 9.054, 2.747, 0.508, LEAD);

	S(s, 'roundRect', 4.597, 8.385, 2.05, 1.178, { rectRadius: 0.204, line: { color: CYAN, width: 1.5 } });
	[5.018, 5.728].forEach(x => S(s, 'roundRect', x, 8.725, 0.498, 0.498,
		{ fill: { color: CYAN }, rectRadius: 0.083, line: { color: OFFWHITE, width: 1 } }));
	icon(s, 'camera', 5.09, 8.797, 0.354, WHITE);
	icon(s, 'note', 5.8, 8.797, 0.354, WHITE);

	[['Dr. ALEX DONNY', 9.64], ['Dr. SOFI MOREY', 12.763], ['Dr. ELDINO VIDALI', 15.885]].forEach(([name, x]) => {
		photo(s, x, 1.287, 2.747, 3.85);
		T(s, name, x, 5.568, 2.747, 0.508, heading(18, LEAD));
		T(s, 'Job Position', x, 6.213, 2.747, 0.463, Object.assign({}, LEAD, { fontSize: 16 }));
	});

	kicker(s, 9.721, 7.644, 'OUR TEAM');
	T(s, 'Meet Our Skilled Ophthalmologists & Specialists', 9.64, 8.381, 9.16, 1.347, TITLE);
}

function slide08(s) {
	cyanBackdrop(s);
	bubbles(s, [[11.224, 9.562, 0.498, true], [1.936, 1.615, 1.234], [2.553, 0.381, 0.617],
		[0.805, 2.232, 0.617], [0.435, 0.463, 0.823, true]],
	[[1.5, 1.579, 0.269, 0.32], [1.839, 0.714, 0.269, 0.32], [0.577, 9.629, 0.269, 0.32]]);
	S(s, 'roundRect', 3.684, 9.549, 7.003, 0.48, { fill: { color: WHITE }, rectRadius: 0.24 });
	S(s, 'roundRect', 3.684, 1.337, 15.134, 7.638, { fill: { color: WHITE }, rectRadius: 0.412 });

	T(s, 'Social Media', 15.144, 9.609, 2.255, 0.438, { fontSize: 20, color: WHITE });
	[17.611, 18.32].forEach((x, i) => {
		S(s, 'roundRect', x, i ? 9.549 : 9.562, 0.498, 0.498, { fill: { color: WHITE }, rectRadius: 0.083 });
		icon(s, i ? 'note' : 'camera', x + 0.072, (i ? 9.549 : 9.562) + 0.072, 0.354, TEAL);
	});

	kicker(s, 9.512, 2.137, 'OPHTHALMOLOGIST', 4.386);
	T(s, 'Dr. Mariana Carrera Dowie', 9.431, 2.874, 8.388, 0.724, TITLE);
	T(s, 'Lorem ipsum dolor sit, consectetur adipiscing the elit, sed do eiusmod tempor ut labore et dolore magna ' +
		'at aliqua. Cursus risus at ultrices. Duis ultricies lacus sed turpis tincidunt aliquet sed euismod mollis.',
	9.431, 4.247, 8.388, 1.313, LEAD);
	T(s, 'WRITE SKILL HERE', 9.431, 6.505, 4.634, 0.438, HEADING);
	T(s, 'Lorem ipsum dolor sit amet, consectetur the elit, sed do eiusmod incididunt ut labore.',
		9.431, 7.142, 5.524, 0.859, LEAD);
	donut(s, 15.661, 6.366, '95%');

	S(s, 'rect', 0.805, 0.821, 7.27, 10.429, { fill: { color: PHOTO_LT } });
}

/** Right-hand stack of white "WRITE TITLE HERE" cards used on pages 9 and 10. */
function textCard(s, x, y, outlined) {
	S(s, 'roundRect', x, y, 4.786, 2.376, outlined
		? { fill: { color: WHITE }, rectRadius: 0.157, line: { color: CYAN, width: 1.5 } }
		: { fill: { color: WHITE }, rectRadius: 0.157, shadow: cardShadow() });
	T(s, 'WRITE TITLE HERE', x + 0.683, y + 0.402, 3.419, 0.438, HEADING);
	T(s, 'Lorem ipsum dolor sit amet, elit, sed et tempor labore.', x + 0.683, y + 1.011, 3.419, 0.963, LEAD);
}

function slide09(s) {
	S(s, 'roundRect', 9.998, 1.805, 3.054, 7.639, { fill: { color: CYAN }, rectRadius: 0.2 });
	photo(s, 6.948, 1.805, 4.933, 7.639);

	kicker(s, 1.461, 2.46, 'OUR SERVICE');
	T(s, 'Cutting-edge Technology', 1.381, 3.196, 4.674, 1.347, TITLE);

	S(s, 'roundRect', 1.487, 5.85, 5.866, 2.94, { fill: { color: WHITE }, rectRadius: 0.194, shadow: cardShadow() });
	S(s, 'roundRect', 1.381, 6.19, 0.285, 2.261, { fill: { color: CYAN }, rectRadius: 0.142 });
	T(s, '150k', 2.247, 6.42, 4.346, 0.707, heading(36));
	T(s, 'Lorem ipsum dolor sit amet,elit, sed tempor labore dolor amet short.', 2.247, 7.257, 4.346, 0.859, LEAD);

	[1.805, 4.437, 7.068].forEach(y => textCard(s, 13.753, y, false));
}

function slide10(s) {
	S(s, 'roundRect', 16.537, 0.733, 2.736, 3.896, { fill: { color: CYAN }, rectRadius: 0.327 });
	S(s, 'ellipse', 18.721, 4.88, 0.398, 0.398, { fill: { color: CYAN } });
	photo(s, 12.164, 1.345, 6.448, 8.56);

	[[1.735, 1.345, 'chat', true], [4.827, 4.437, 'cross', false], [7.919, 7.529, 'mortar', false]]
		.forEach(([ty, cy, glyph, filled]) => {
		S(s, 'roundRect', 1.388, ty, 1.906, 1.596, filled
			? { fill: { color: CYAN }, rectRadius: 0.166, shadow: tileShadow() }
			: { rectRadius: 0.166, line: { color: CYAN, width: 1.5 } });
			icon(s, glyph, 1.841, ty + 0.298, 1.0, filled ? WHITE : CYAN);
		textCard(s, 3.748, cy, !filled);
	});

	S(s, 'roundRect', 10.699, 6.411, 6.642, 3.805, { fill: { color: WHITE }, rectRadius: 0.258 });
	kicker(s, 11.416, 7.155, 'OUR SERVICE', 4.386);
	T(s, 'Affordable & Accessible Services', 11.336, 7.891, 5.367, 1.347, TITLE);
}

function slide11(s) {
	S(s, 'round2SameRect', -3.91, 4.901, 9.268, 1.449, { fill: { color: CYAN }, rectRadius: 0.483, rotate: 90 });
	S(s, 'round2SameRect', -3.438, 5.229, 7.668, 0.791, { fill: { color: WHITE }, rectRadius: 0.24, rotate: 90 });

	kicker(s, 2.706, 2.242, 'OUR SERVICE', 4.386);
	T(s, 'From Diagnosis to Treatment with Excellence', 2.626, 2.978, 7.955, 1.347, TITLE);

	S(s, 'ellipse', 2.626, 6.432, 1.628, 1.628, { fill: { color: CYAN } });
	scanEye(s, 3.055, 6.962);
	elbow(s, 3.44, 5.648, 1.992, 0.62, true);
	elbow(s, 4.628, 6.619, 0.952, 0.62, true);
	elbow(s, 4.628, 7.253, 0.952, 0.62, false);
	elbow(s, 3.44, 8.223, 1.992, 0.62, false);
	[5.553, 6.509, 7.748, 8.728].forEach(y => S(s, 'ellipse', 5.397, y, 0.22, 0.22, { fill: { color: CYAN } }));
	[['Cataract and refractive surgery (LASIK))', 5.345, 4.94],
		['Retina and cornea treatments', 6.286, 4.634],
		['Pediatric ophthalmology', 7.559, 4.634],
		['Contact lenses and optical services', 8.5, 5.097],
	].forEach(([str, y, w]) => T(s, str, 5.947, y, w, 0.404, LEAD));

	S(s, 'roundRect', 12.592, 0.938, 6.274, 9.373, { fill: { color: WHITE }, rectRadius: 0.415, shadow: cardShadow() });
	S(s, 'roundRect', 12.507, 2.911, 0.3, 5.429, { fill: { color: CYAN }, rectRadius: 0.15 });
	[1.789, 3.822, 5.855, 7.888].forEach(y => {
		T(s, 'WRITE TITLE HERE', 13.844, y, 3.419, 0.438, HEADING);
		T(s, 'Lorem ipsum dolor sit amet, elit, sed tempor labore.', 13.844, y + 0.61, 4.335, 0.963, LEAD);
	});
}

function slide12(s) {
	S(s, 'rect', 0, 0, 20, 3.776, { fill: { color: CYAN } });
	S(s, 'rect', 0, 0, 20, 3.643, { fill: { color: PHOTO } });

	logo(s, 9.766, 4.684);
	T(s, 'OUR PORTFOLIO', 7.807, 5.537, 4.386, 0.404, { fontFace: HEAD, fontSize: 18, color: TEAL, align: 'center' });
	T(s, 'Optical Coherence Tomography (OCT)', 2.739, 6.169, 14.522, 0.724, Object.assign({}, TITLE, { align: 'center' }));

	[1.494, 7.607, 13.72].forEach((x, i) => {
		S(s, 'roundRect', x, 7.966, 4.786, 2.376, Object.assign(
			{ fill: { color: WHITE }, rectRadius: 0.157, line: { color: CYAN, width: 1.5 } },
			i === 1 ? { shadow: shadow(38, 0, 'A5A5A5', 0.3) } : {}));
		T(s, 'WRITE TITLE HERE', x + 0.683, 8.368, 3.419, 0.438, heading(20, { align: 'center' }));
		T(s, 'Lorem ipsum dolor sit amet, elit, sed et tempor labore.', x + 0.683, 8.977, 3.419, 0.963,
			Object.assign({}, LEAD, { align: 'center' }));
	});
	S(s, 'roundRect', 8.523, 7.85, 2.954, 0.256, { fill: { color: CYAN }, rectRadius: 0.128 });
}

function slide13(s) {
	S(s, 'roundRect', 1.478, 3.435, 17.045, 3.513, { fill: { color: WHITE }, rectRadius: 0.232, line: { color: CYAN, width: 1.5 } });

	T(s, 'BEST PORTFOLIO', 7.807, 1.093, 4.386, 0.404, { fontFace: HEAD, fontSize: 18, color: TEAL, align: 'center' });
	T(s, 'Laser-assisted Surgeries', 2.403, 1.707, 15.195, 0.724, Object.assign({}, TITLE, { align: 'center' }));

	photo(s, 2.265, 4.092, 7.459, 2.198, 0.12);
	photo(s, 10.276, 4.092, 7.459, 2.198, 0.12);

	S(s, 'roundRect', 1.478, 7.921, 3.129, 0.3, { fill: { color: CYAN }, rectRadius: 0.15 });
	T(s, 'SMART DIAGNOSTIC TOOLS', 1.478, 8.431, 8.616, 0.554, heading(20, LEAD));
	T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod et tempor incididunt ut labore. ' +
		'PLACEHOLDER' +
		'salami picanha doner. Kevin capicola sausage buffalo.', 1.478, 9.194, 17.045, 0.963, LEAD);
}

function slide14(s) {
	S(s, 'roundRect', 1.362, 1.349, 4.483, 4.276, { fill: { color: CYAN }, rectRadius: 0.487 });
	S(s, 'ellipse', 1.362, 5.99, 0.621, 0.621, { fill: { color: CYAN } });

	// Desktop monitor: black bezel + screen, then the chin, neck and base of the stand.
	S(s, 'rect', 2.82, 2.68, 5.68, 3.3, { fill: { color: '111111' } });
	S(s, 'rect', 2.985, 2.87, 5.303, 2.954, { fill: { color: PHOTO } });
	S(s, 'rect', 2.82, 5.98, 5.68, 0.52, { fill: { color: 'D3D3D3' } });
	S(s, 'rect', 5.03, 6.5, 1.26, 0.76, { fill: { color: 'A8A8A8' } });
	S(s, 'roundRect', 4.64, 7.24, 2.67, 0.21, { fill: { color: 'C0C0C0' }, rectRadius: 0.1 });

	donut(s, 1.362, 8.163, '95%');
	T(s, 'WRITE TITLE HERE', 4.259, 8.163, 4.634, 0.438, HEADING);
	T(s, 'Lorem ipsum dolor sit amet, consectetur the elit, sed do eiusmod tempor.', 4.259, 8.799, 4.856, 0.859, LEAD);

	kicker(s, 11.202, 2.136, 'MOCKUP DEVICE');
	T(s, 'Information and Approach to Patients', 11.107, 2.87, 7.342, 1.347, TITLE);
	T(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, magna aliqua. ultrices mi. Duis lacus sed turpis ' +
		'tincidunt aliquet  euismod porta lorem mollis aliquam .Bacon dolor amet short ribs venison rump drumstick ' +
		'pig sausage prosciutto chicken spare ribs picanha doner. Kevin capicola sausage, buffalo bresaola venison ' +
		'eiusmod tempor the incididunt.', 11.107, 4.886, 7.342, 2.677, LEAD);
	S(s, 'roundRect', 11.121, 8.175, 3.377, 0.939, { fill: { color: CYAN }, rectRadius: 0.469 });
	T(s, 'LEARN MORE', 11.379, 8.442, 2.942, 0.404, BTN_TEXT);
}

function slide15(s) {
	gradient(s, 14.224, 0, 2.656, 11.25, TEAL, CYAN, 12);
	S(s, 'rect', 16.88, 0, 3.12, 11.25, { fill: { color: CYAN } });
	S(s, 'roundRect', 11.096, 7.0, 7.701, 2.326, { fill: { color: WHITE }, rectRadius: 0.47, line: { color: CYAN, width: 1 } });

	// Phone in landscape: thin light body edge, dark bezel, grey screen.
	screen(s, 11.1, 1.93, 7.69, 4.33, 0.5, 0.05);
	S(s, 'roundRect', 11.344, 2.156, 7.207, 3.879, { fill: { color: PHOTO }, rectRadius: 0.3 });

	kicker(s, 1.646, 2.301, 'MOCKUP DEVICE');
	T(s, 'Appointment Booking Information', 1.551, 3.035, 7.342, 1.347, TITLE);
	T(s, 'Lorem ipsum dolor sit, consectetur adipiscing elit, sed  magna aliqua mi. Duis ultricies lacus sed turpis ' +
		'tincidunt aliquet  euismod porta lorem mollis aliquam .Bacon ipsum amet short ribs venison rump drumstick ' +
		'pig sausage prosciutto it chicken spare ribs picanha doner. Kevin capicola sausage, buffalo.',
	1.551, 5.12, 7.342, 2.222, LEAD);
	S(s, 'roundRect', 1.565, 8.01, 3.377, 0.939, { fill: { color: CYAN }, rectRadius: 0.469 });
	T(s, 'DOWNLOAD NOW', 1.823, 8.277, 2.942, 0.404, BTN_TEXT);

	[[11.57, 'bed', true], [13.994, 'mortar', false], [16.418, 'pill', false]].forEach(([x, glyph, filled]) => {
		S(s, 'roundRect', x, 7.364, 1.906, 1.596, filled
			? { fill: { color: CYAN }, rectRadius: 0.166, shadow: tileShadow() }
			: { rectRadius: 0.166, line: { color: CYAN, width: 1.5 } });
		icon(s, glyph, x + 0.49, 7.699, 0.927, filled ? WHITE : CYAN);
	});
}

function slide16(s) {
	S(s, 'roundRect', 14.428, 4.928, 5.572, 4.341, { fill: { color: CYAN }, rectRadius: 0.4 });
	S(s, 'roundRect', 13.082, 5.826, 5.572, 4.248, { fill: { color: WHITE }, rectRadius: 0.4, line: { color: CYAN, width: 1.5 } });

	[[1.303, 2.083, 2.366, 1.26, 2.507, 1.4, 2.015, 3.095, 1.88, 3.657, 'bulb', false],
		[6.399, 2.083, 7.462, 1.26, 7.602, 1.4, 7.111, 3.1, 6.924, 3.662, 'percent', true],
		[1.303, 6.876, 2.366, 6.053, 2.507, 6.194, 2.015, 7.897, 1.88, 8.46, 'tree', true],
		[6.399, 6.876, 7.462, 6.053, 7.602, 6.194, 7.111, 7.897, 6.976, 8.46, 'shield', true],
	].forEach(([cx, cy, px, py, ix, iy, tx, ty, bx, by, glyph, outlined]) => {
		S(s, 'roundRect', px, py, 1.646, 3.661, { fill: { color: TEAL }, rectRadius: 0.823 });
		S(s, 'roundRect', cx, cy, 4.314, 3.116, outlined
			? { fill: { color: WHITE }, rectRadius: 0.3, line: { color: CYAN, width: 1.5 } }
			: { fill: { color: WHITE }, rectRadius: 0.3, shadow: shadow(28, 15, '7F7F7F') });
		S(s, 'ellipse', ix, iy, 1.365, 1.365, { fill: { color: CYAN } });
		icon(s, glyph, ix + 0.313, iy + 0.313, 0.74, WHITE);
		T(s, 'WRITE TITLE HERE', tx, ty, 2.813, 0.438, heading(20, { align: 'center' }));
		T(s, 'Lorem ipsum dolor sit elit, sed do eiusmod tempor.', bx, by, 3.16, 0.859,
			Object.assign({}, LEAD, { align: 'center' }));
	});

	kicker(s, 13.102, 1.703, 'INFOGRAPHIC');
	T(s, 'Infographic Section', 13.007, 2.436, 5.704, 0.774, Object.assign({}, TITLE, { fontSize: 40 }));
	photo(s, 13.304, 5.995, 5.128, 3.909, 0.3);
}

function slide17(s) {
	S(s, 'roundRect', 13.009, 4.311, 5.558, 5.044, { fill: { color: CYAN }, rectRadius: 0.575 });
	S(s, 'ellipse', 17.961, 9.559, 0.605, 0.569, { fill: { color: CYAN } });

	kicker(s, 1.528, 1.122, 'PRICING PLAN');
	T(s, 'Customized Treatment Plans', 1.433, 1.855, 12.389, 0.724, TITLE);
	S(s, 'roundRect', 1.538, 3.537, 16.158, 6.591, { fill: { color: WHITE }, rectRadius: 0.556, shadow: panelShadow() });
	S(s, 'roundRect', 1.428, 5.109, 0.282, 3.448, { fill: { color: CYAN }, rectRadius: 0.141 });

	[['BASIC PLAN', 2.58, '105', 4.311, 4.865, 0.707, 4.979, 0.404, 5.249, 6.111, 8.426, 8.694, 2.444],
		['STANDART', 7.556, '350', 4.306, 4.86, 0.916, 4.974, 0.512, 5.244, 6.105, 8.421, 8.689, 2.444],
		['PREMIUM', 12.532, '550', 4.3, 4.854, 0.707, 4.968, 0.404, 5.238, 6.1, 8.415, 8.683, 1.647],
	].forEach(([name, x, price, ny, py, ph, dy, dh, yy, by, bty, gy, nw]) => {
		T(s, name, x, ny, nw, 0.37, Object.assign({}, LEAD, { fontFace: HEAD, fontSize: 16, bold: true, color: '030303' }));
		T(s, price, x + 0.191, py, 1.157, ph, heading(36, LEAD));
		T(s, '$', x - 0.117, dy, 0.562, dh, heading(18, { align: 'center', lineSpacingMultiple: 1.5 }));
		T(s, '/Years', x + 1.058, yy, 0.873, 0.337, Object.assign({}, LEAD, { fontFace: HEAD, fontSize: 14, color: '7F7F7F' }));
		T(s, 'Lorem ipsum dolor sit amet elit, sed  magna aliqua. ultrices mi lacus turpis tincidunt aliquet ribs ' +
			'brisket venison rump drumstick sausage.', x, by, 4.29, 1.767, LEAD);
		S(s, 'roundRect', x, bty, 2.942, 0.939, { fill: { color: CYAN }, rectRadius: 0.469 });
		T(s, 'GET STARTED', x + 0.211, gy, 2.52, 0.404, BTN_TEXT);
	});
}

function slide18(s) {
	S(s, 'roundRect', 2.551, 2.245, 15.957, 6.75, { fill: { color: WHITE }, rectRadius: 0.569, line: { color: CYAN, width: 1.5 } });
	S(s, 'roundRect', 1.501, 1.399, 4.239, 7.595, { fill: { color: CYAN }, rectRadius: 0.82 });
	S(s, 'ellipse', 1.437, 9.222, 0.651, 0.651, { fill: { color: CYAN } });
	T(s, '?', 1.316, 9.261, 0.894, 0.572, { fontFace: HEAD, fontSize: 28, color: WHITE, align: 'center', rotate: 7 });

	photo(s, 1.909, 1.93, 6.21, 7.379, 0.6);
	T(s, 'When I think of art I think of beauty. Beauty is the mystery of life. It is not in the eye it is in the ' +
		'mind. In our minds there is awareness of perfection.', 9.701, 2.99, 7.972, 3.215, TITLE);
	S(s, 'line', 9.885, 7.061, 4.07, 0, { line: { color: TEAL, width: 1.5 } });
	T(s, 'Agnes Martin, Sp.M', 9.775, 7.218, 5.139, 0.505, { fontFace: HEAD, fontSize: 24, bold: true, color: '1F2126' });
	T(s, '( Ophthalmologist )', 9.775, 7.879, 3.613, 0.37, { fontSize: 16 });
}

function slide19(s) {
	gradientPanel(s, 8.966, 1.079, 7.174, 9.091, CYAN, TEAL, 0.477);

	kicker(s, 1.673, 2.862, 'CONTACT ');
	T(s, 'Get in Touch and Start Your Vision Journey Today', 1.578, 3.596, 5.41, 1.969, TITLE);
	T(s, 'Contact our receptionist:', 1.578, 6.742, 2.795, 0.37, { fontSize: 16, color: '2E3138' });
	S(s, 'line', 1.687, 7.2, 4.748, 0, { line: { color: 'D8D8D8', width: 1.5 } });
	T(s, 'Veronika Zayanna, SE', 1.578, 7.356, 5.139, 0.505, { fontFace: HEAD, fontSize: 24, bold: true, color: '1F2126' });
	T(s, '( receptionist )', 1.578, 8.017, 3.613, 0.37, { fontSize: 16 });

	photo(s, 8.966, 1.079, 3.705, 9.091, 0.3);

	[['Location', 'pin', 13.141, 1.417, 5.181, WHITE, 1.915, 2.524, 4.079,
		['5 East 68th Street, New York, NY 10065 , The United States']],
	['Telephone', 'phone', 13.137, 4.315, 5.286, 'F4F6FE', 4.75, 5.418, 3.875,
		['+97 9877 6788 8710', '+97 5465 1234 8810']],
	['Website', 'globe', 13.137, 7.333, 5.286, 'F4F6FE', 7.77, 8.432, 3.681,
		['www.ophthalmology.com', 'info@ophthalmology.com']],
	].forEach(([label, glyph, x, y, w, fill, ty, by, bw, lines]) => {
		S(s, 'roundRect', x, y, w, 2.5, Object.assign({ fill: { color: fill }, rectRadius: 0.227 },
			fill === WHITE ? { shadow: shadow(28, 15, INK) } : {}));
		T(s, label, x + 0.45, ty, 2.858, 0.454, { fontFace: HEAD, fontSize: 21, bold: true, color: '1F2126' });
		T(s, lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), x + 0.45, by, bw, 0.963, LEAD);
		icon(s, glyph, x + 4.35, ty - 0.06, 0.42, INK);
	});
}

function slide20(s) {
	S(s, 'rect', 0, 0, 20, 11.25, { fill: { color: TEAL } });
	bubbles(s, [[2.158, 7.23, 1.278, true], [2.798, 5.951, 0.639], [0.987, 7.869, 0.639],
		[0.667, 6.378, 0.852], [18.828, 10.1, 0.691, true]],
	[[1.707, 7.192, 0.279, 0.332], [2.059, 6.296, 0.279, 0.332], [3.817, 8.595, 0.279, 0.332],
		[16.156, 5.949, 0.278, 0.331], [17.87, 9.865, 0.427, 0.508]]);

	S(s, 'rect', 0, 0, 20, 5.299, { fill: { color: PHOTO } });
	S(s, 'round2SameRect', 4.069, 4.641, 11.862, 0.658, { fill: { color: TEAL }, rectRadius: 0.279 });
	S(s, 'round2SameRect', 5.093, 4.939, 9.814, 0.359, { fill: { color: WHITE }, rectRadius: 0.179 });

	T(s, 'THANK YOU', 4.511, 6.807, 10.978, 2.121,
		{ fontFace: HEAD, fontSize: 120, bold: true, color: WHITE, align: 'center' });
	T(s, 'Ophthalmology Clinic Presentation Template', 4.916, 9.22, 10.168, 0.522,
		{ fontSize: 25, color: OFFWHITE, align: 'center' });
}

/* ------------------------------------------------------------------ build */
const pages = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
	slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE20', width: 20, height: 11.25 });
pptx.layout = 'WIDE20';
pptx.title = 'Ophthalmology Clinic';

pages.forEach(build => {
	const slide = pptx.addSlide();
	slide.background = { color: WHITE };
	build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '10f94dbf-8a78-4eec-aa83-b7c0cc8b67dc_grok_final.pptx') })
	.then(f => console.log('wrote', f));
