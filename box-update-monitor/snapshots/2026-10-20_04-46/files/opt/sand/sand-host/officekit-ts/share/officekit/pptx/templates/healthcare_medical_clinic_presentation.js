/**
 * "Medico" — Clean Medical Presentation Template (32 slides, 13.333 x 7.5 in).
 * Rebuilt from scratch with pptxgenjs. Raster mock-ups in the original deck are
 * replaced by labelled placeholder rectangles; the vector icons are redrawn from
 * simple primitives.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const INDIGO = '2D35A0';
const CYAN = '00B4D8';
const INK = '404040';
const DARK = '181717';
const GREY = '767171';
const MIDGREY = '808080';
const LIGHTGREY = 'A6A6A6';
const SILVER = 'AFABAB';
const RED = 'D43A42';
const NEARBLACK = '001A22';
const WHITE = 'FFFFFF';
const BODYINK = '000000';

const FONT = 'Poppins';
const FONT_MED = 'Poppins Medium';

/* ------------------------------------------------------------------ helpers */
/** roundRect corner radius in inches, from the OOXML "adj" value. */
const rr = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/** Mix two hex colours; t=0 -> a, t=1 -> b. */
function blend(a, b, t) {
	const ch = (s, i) => parseInt(s.substr(i * 2, 2), 16);
	return [0, 1, 2]
		.map(i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0'))
		.join('')
		.toUpperCase();
}

/** The template's soft drop shadow (fresh object per shape — pptxgenjs mutates it). */
const cardShadow = () => ({ type: 'outer', color: '000000', opacity: 0.17, blur: 45, offset: 22, angle: 25 });

/**
 * Smooth diagonal fill: bands rotated 45 degrees and stepped along the gradient
 * axis. The bands overhang the box, so this is only used for full-slide
 * backgrounds where the slide edge does the cropping (`bleed: true`).
 */
function gradientBands(slide, o) {
	const n = o.bands || 26;
	const k = Math.SQRT1_2;
	const up = o.dir === 'bl' ? -1 : 1; // screen-space y direction of the axis
	const axisLen = (o.w + o.h) * k;
	const bandLen = Math.hypot(o.w, o.h);
	const step = axisLen / n;
	for (let i = 0; i < n; i++) {
		const s = -axisLen / 2 + (i + 0.5) * step;
		const cx = o.x + o.w / 2 + s * k;
		const cy = o.y + o.h / 2 + s * k * up;
		slide.addShape('rect', {
			x: cx - step / 2 - 0.03, y: cy - bandLen / 2,
			w: step + 0.06, h: bandLen,
			rotate: up < 0 ? 315 : 45,
			fill: { color: blend(o.from, o.to, i / (n - 1)) }, line: { type: 'none' },
		});
	}
}

/**
 * pptxgenjs has no gradient fill, so gradients are painted as a mosaic of solid
 * tiles. `dir` is 'h' (left->right), 'v' (top->bottom), 'bl' (bottom-left ->
 * top-right) or 'tl' (top-left -> bottom-right).
 *
 * `radius` keeps the rounded corners of a card intact: the two ends of the long
 * axis are drawn as rounded rectangles, and every tile is split into a
 * vertically inset part (clears the corner arcs) plus a horizontally inset part
 * (fills the rows the arcs live in).
 */
function gradient(slide, o) {
	const r = o.radius || 0;
	if (o.bleed) return gradientBands(slide, o);
	const bands = o.bands || 28;
	const diag = o.dir === 'bl' || o.dir === 'tl';
	const cols = o.dir === 'v' ? 1 : bands;
	const rows = o.dir === 'h' ? 1 : diag ? Math.max(3, Math.round((bands * o.h) / o.w)) : bands;
	const at = (u, v) => blend(o.from, o.to,
		o.dir === 'h' ? u : o.dir === 'v' ? v : o.dir === 'tl' ? (u + v) / 2 : (u + 1 - v) / 2);
	const lap = 0.06; // tile overlap, hides antialiased seams
	const rect = (x, y, w, h, color) => slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
	const horiz = o.w >= o.h;

	if (r) {
		const cap = 2 * r / (horiz ? o.w : o.h);
		[0, 1].forEach(end => {
			const t = end ? 1 - cap / 2 : cap / 2;
			slide.addShape('roundRect', {
				x: o.x + (horiz && end ? o.w - 2 * r : 0), y: o.y + (!horiz && end ? o.h - 2 * r : 0),
				w: horiz ? 2 * r : o.w, h: horiz ? o.h : 2 * r, rectRadius: r,
				fill: { color: horiz ? at(t, 0.5) : at(0.5, t) }, line: { type: 'none' },
				shadow: o.shadow ? cardShadow() : undefined,
			});
		});
	} else if (o.shadow) {
		slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: at(0.5, 0.5) }, line: { type: 'none' }, shadow: cardShadow() });
	}

	for (let c = 0; c < cols; c++) {
		for (let w = 0; w < rows; w++) {
			const color = at(cols === 1 ? 0.5 : c / (cols - 1), rows === 1 ? 0.5 : w / (rows - 1));
			const x0 = o.x + (o.w * c) / cols;
			const x1 = Math.min(o.x + o.w, x0 + o.w / cols + lap);
			const y0 = o.y + (o.h * w) / rows;
			const y1 = Math.min(o.y + o.h, y0 + o.h / rows + lap);
			const a = Math.max(y0, o.y + r);
			const b = Math.min(y1, o.y + o.h - r);
			if (b > a) rect(x0, a, x1 - x0, b - a, color);
			if (r && (a > y0 || b < y1)) {
				const p = Math.max(x0, o.x + r);
				const q = Math.min(x1, o.x + o.w - r);
				if (q > p) rect(p, y0, q - p, y1 - y0, color);
			}
		}
	}
}

/** White card with the template's soft drop shadow. */
function card(slide, x, y, w, h, adj, fill) {
	slide.addShape('roundRect', {
		x, y, w, h, rectRadius: rr(adj, w, h),
		fill: { color: fill || WHITE }, line: { type: 'none' }, shadow: cardShadow(),
	});
}

/** Text with the deck defaults (Poppins, top aligned, no auto-shrink). */
function txt(slide, text, o) {
	slide.addText(text, Object.assign({ fontFace: FONT, color: BODYINK, valign: 'top' }, o));
}
const body = (slide, text, o) => txt(slide, text, Object.assign({ fontSize: 9, lineSpacingMultiple: 2 }, o));
const title = (slide, text, o) => txt(slide, text, Object.assign({ fontSize: 40, color: INK }, o));

/* The little ECG heartbeat mark that closes every footer (traced outline). */
const ECG = [
	[0.3152, 1.0], [0.101, 0.2957], [0.02, 0.5635], [0.0, 0.5496], [0.101, 0.2157],
	[0.3152, 0.9183], [0.5938, 0.0], [0.8701, 0.9096], [0.98, 0.5478], [1.0, 0.5635],
	[0.8701, 0.9896], [0.5938, 0.0817],
];
function heartbeat(slide, x, y, w, h, fill, lineColor) {
	slide.addShape('custGeom', {
		x, y, w, h,
		points: ECG.map(p => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }]),
		fill: { color: fill }, line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' },
	});
}

/** Footer: caption on the left, hairline rule, heartbeat mark on the right. */
function footer(slide, o) {
	o = o || {};
	const col = o.color || INDIGO;
	txt(slide, o.label || 'Medical Presentation', {
		x: 0.638, y: 6.661, w: 2.626, h: 0.303,
		fontFace: o.font || FONT, fontSize: 12, bold: !!o.bold, color: col,
	});
	slide.addShape('line', { x: 3.273, y: 6.826, w: 8.721, h: 0, line: { color: o.rule || col, width: 0.5 } });
	heartbeat(slide, 11.985, 6.625, 0.588, 0.375, o.ecgFill || RED, o.ecgLine === null ? null : (o.ecgLine || INDIGO));
}

/* ------------------------------------------------------------------- icons */
/* Each icon is a list of primitives in 0..1 coordinates of its own box.
   'fg' paints with the icon colour, 'bg' punches back to the plate colour.  */
const ICONS = {
	kit: [
		{ s: 'roundRect', x: 0.30, y: 0.00, w: 0.40, h: 0.22, r: 0.05, c: 'fg' },
		{ s: 'roundRect', x: 0.40, y: 0.06, w: 0.20, h: 0.14, r: 0.03, c: 'bg' },
		{ s: 'roundRect', x: 0.00, y: 0.16, w: 1.00, h: 0.84, r: 0.12, c: 'fg' },
		{ s: 'rect', x: 0.43, y: 0.36, w: 0.14, h: 0.44, c: 'bg' },
		{ s: 'rect', x: 0.28, y: 0.51, w: 0.44, h: 0.14, c: 'bg' },
	],
	chart: [
		{ s: 'rect', x: 0.02, y: 0.56, w: 0.24, h: 0.44, c: 'fg' },
		{ s: 'rect', x: 0.38, y: 0.32, w: 0.24, h: 0.68, c: 'fg' },
		{ s: 'rect', x: 0.74, y: 0.08, w: 0.24, h: 0.92, c: 'fg' },
	],
	bed: [
		{ s: 'ellipse', x: 0.04, y: 0.12, w: 0.24, h: 0.24, c: 'fg' },
		{ s: 'roundRect', x: 0.22, y: 0.28, w: 0.70, h: 0.26, r: 0.10, c: 'fg' },
		{ s: 'rect', x: 0.00, y: 0.56, w: 1.00, h: 0.16, c: 'fg' },
		{ s: 'rect', x: 0.00, y: 0.72, w: 0.13, h: 0.24, c: 'fg' },
		{ s: 'rect', x: 0.87, y: 0.72, w: 0.13, h: 0.24, c: 'fg' },
	],
	hand: [
		{ s: 'heart', x: 0.22, y: 0.00, w: 0.56, h: 0.52, c: 'fg' },
		{ s: 'roundRect', x: 0.00, y: 0.58, w: 1.00, h: 0.30, r: 0.14, c: 'fg' },
	],
	snake: [
		{ s: 'roundRect', x: 0.41, y: 0.00, w: 0.18, h: 1.00, r: 0.08, c: 'fg' },
		{ s: 'ellipse', x: 0.16, y: 0.10, w: 0.40, h: 0.22, c: 'fg' },
		{ s: 'ellipse', x: 0.44, y: 0.39, w: 0.40, h: 0.22, c: 'fg' },
		{ s: 'ellipse', x: 0.16, y: 0.68, w: 0.40, h: 0.22, c: 'fg' },
	],
	syringe: [
		{ s: 'roundRect', x: 0.14, y: 0.28, w: 0.56, h: 0.44, r: 0.06, c: 'fg' },
		{ s: 'rect', x: 0.70, y: 0.44, w: 0.30, h: 0.12, c: 'fg' },
		{ s: 'rect', x: 0.00, y: 0.40, w: 0.14, h: 0.20, c: 'fg' },
		{ s: 'rect', x: 0.26, y: 0.36, w: 0.08, h: 0.28, c: 'bg' },
	],
	quote: [
		{ s: 'ellipse', x: 0.00, y: 0.02, w: 0.42, h: 0.46, c: 'fg' },
		{ s: 'ellipse', x: 0.10, y: 0.12, w: 0.20, h: 0.22, c: 'bg' },
		{ s: 'rect', x: 0.20, y: 0.24, w: 0.22, h: 0.74, c: 'fg' },
		{ s: 'ellipse', x: 0.56, y: 0.02, w: 0.42, h: 0.46, c: 'fg' },
		{ s: 'ellipse', x: 0.66, y: 0.12, w: 0.20, h: 0.22, c: 'bg' },
		{ s: 'rect', x: 0.76, y: 0.24, w: 0.22, h: 0.74, c: 'fg' },
	],
};

function icon(slide, name, x, y, w, h, fg, bg) {
	ICONS[name].forEach(p => {
		slide.addShape(p.s, {
			x: x + p.x * w, y: y + p.y * h, w: p.w * w, h: p.h * h,
			rectRadius: p.r ? p.r * Math.min(w, h) * 4 : undefined,
			fill: { color: p.c === 'bg' ? bg : fg }, line: { type: 'none' },
		});
	});
}

/** Circular plate + icon, the recurring "feature bullet" of the deck. */
function iconPlate(slide, o) {
	if (o.gradient) gradient(slide, { x: o.x, y: o.y, w: o.d, h: o.d, from: o.gradient[0], to: o.gradient[1], dir: 'v', bands: 14, radius: o.d / 2 });
	else slide.addShape('ellipse', { x: o.x, y: o.y, w: o.d, h: o.d, fill: { color: o.plate }, line: { type: 'none' } });
	const s = o.iconSize || o.d * 0.5;
	icon(slide, o.icon, o.x + (o.d - s) / 2, o.y + (o.d - s) / 2, s, s * 0.96, o.fg, o.gradient ? blend(o.gradient[0], o.gradient[1], 0.5) : o.plate);
}

/** Stand-in for a photo / product mock-up from the source deck. */
function imagePlaceholder(slide, x, y, w, h, label) {
	slide.addShape('rect', { x, y, w, h, fill: { color: 'FAFAFA' }, line: { color: 'E0E0E0', width: 1 } });
	txt(slide, label || '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', fontSize: 11, color: SILVER });
}

/* ------------------------------------------------------- repeated copy text */
const LOREM_A = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisque larakanwo malesuadloo libero ellentesque. Morbi orci dui,galang fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi orci duadad adaadvefwati,galang fermentum egeti';
const LOREM_B = 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo.';
const LOREM_C = 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo remaei adipiscing elit. Pellentesque sceler conse lectus ornare galang tarabintar fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi orci duadad adaadvefwati,galang fermentum egeti tutwurihandoet';
const LOREM_D = 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo ipsu conse lectus ornare.';
const LOREM_E = 'Lorem ipsum dolor sit amet, consectetur adipiscig eli kitabi bullei Pellentesque scelerisque larakanwo dadalerei malesuadloo libero ellentesque. Morbi orciwkante dui,galang fermentum  lectus orna';
const LOREM_F = 'Lorem ipsum dolor sit amet, consectetur adipiscig eliti Pellentesque scelerisque larakanwo dadalerei darareli malesuadloo libero ellentesque. Morbi orciwkanteloeri dui,galang fermentum eget lectus ornare, viverral gune dignissim risus. Done ellentesque. Porbi orcipawa faan Pellentesque scelerisque larakanwo dadalerei dararelir ';
const LOREM_G = 'Lorem ipsum dolor sit amet, consectetur adipiscig elita Pellentesque scelerisque larakanwo dadalerei dararele malesuadloo libero ellentesque. Morbi orciwkanteloeri';
const LOREM_H = 'Lorem ipsu conse lectus ornare, pellentesq kata viverra ctetuwwlo ipsu conse lectus ornare dolo sit amet, consectetur adipiscing elita katkan kata';
const LOREM_I = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit katakana katak Pellentesque sceleri kataka kaw malesuadloo liberokataka ellentesque. Morbi orci dui,galang fermentum egetawo qolowo lectuawos lowowieo consectetur adipiscing elit katakan lolowanem katakanw kaowaa lowoa';
const LOREM_J = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisque kataka malesuadlooa liberokataka ellentesque. Morbi orci yow lowakana';
const BULLET_DOT = { characterCode: '2022', indent: 22.5 };
const BULLETS = ['Take care of your health.', 'The possibilities are endless.', 'Be happy. Be healthy.']
	.map(t => ({ text: t, options: { breakLine: true, bullet: BULLET_DOT } }));
const PLAN_LINES = ['Be safe and health.', 'Your health is our top priority.', 'For your family\u2019s wellness.'];

/* =================================================================== slides */

/* 1 — cover */
function slide01(pptx) {
	const s = pptx.addSlide();
	title(s, 'Medico', { x: 7.436, y: 2.372, w: 5.434, h: 1.447, fontSize: 80, color: INDIGO });
	txt(s, 'Clean Medical Presentation Template', { x: 7.567, y: 3.946, w: 4.538, h: 0.37, fontSize: 16, color: GREY });
	txt(s, 'Clean Medical Presentation', { x: 7.59, y: 6.557, w: 2.626, h: 0.303, fontSize: 12, bold: true, color: INDIGO });
	s.addShape('line', { x: 6.789, y: 6.708, w: 0.6, h: 0, line: { color: '4472C4', width: 0.75 } });
	gradient(s, { x: 0, y: 7.18, w: 6.78, h: 0.37, from: CYAN, to: INDIGO, dir: 'h' });
}

/* 2 — "We Always Ready For Your Health" */
function slide02(pptx) {
	const s = pptx.addSlide();
	s.addShape('diagStripe', { x: 11.694, y: 0, w: 1.639, h: 1.444, flipH: true, fill: { color: blend(INDIGO, CYAN, 0.45) }, line: { type: 'none' } });
	title(s, 'We Always Ready For Your Health', { x: 1.233, y: 1.619, w: 5.281, h: 1.447 });
	body(s, LOREM_A, { x: 1.233, y: 3.267, w: 5.646, h: 1.272 });
	card(s, 1.233, 4.783, 5.726, 1.253, 16667);
	iconPlate(s, { x: 1.494, y: 4.948, d: 0.911, plate: INDIGO, fg: WHITE, icon: 'kit', iconSize: 0.47 });
	txt(s, 'Medicine Labs', { x: 2.665, y: 4.989, w: 1.88, h: 0.37, fontSize: 16, bold: true, color: INDIGO });
	body(s, LOREM_B, { x: 2.665, y: 5.277, w: 4.029, h: 0.707 });
	footer(s, { bold: true, ecgLine: '4472C4', rule: '4472C4' });
}

/* 3 — "Improving Lives Together" (about us) */
function slide03(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 0.76, y: 0.951, w: 11.812, h: 4.481, from: CYAN, to: INDIGO, dir: 'bl', radius: rr(3404, 11.812, 4.481), bands: 36 });
	txt(s, 'A B O U T  U S', { x: 7.567, y: 1.426, w: 1.514, h: 0.286, fontSize: 11, bold: true, color: WHITE });
	title(s, 'Improving Lives Together', { x: 7.525, y: 1.798, w: 4.448, h: 1.447, color: WHITE });
	body(s, LOREM_C, { x: 7.567, y: 3.435, w: 4.477, h: 1.616, color: WHITE });
	[
		{ x: 1.499, adj: 4324, icon: 'chart', label: 'Health Analysis', lx: 1.73, lw: 2.051, bx: 1.83 },
		{ x: 4.264, adj: 3925, icon: 'kit', label: 'Health Control', lx: 4.538, lw: 1.957, bx: 4.596 },
	].forEach(c => {
		card(s, c.x, 2.158, 2.537, 3.704, c.adj);
		icon(s, c.icon, c.x + 0.99, 2.7, 0.6, 0.573, INDIGO, WHITE);
		txt(s, c.label, { x: c.lx, y: 3.669, w: c.lw, h: 0.37, fontSize: 16, color: INDIGO, align: 'center' });
		body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo rem adipiscing sceler conse lectus ornare.', { x: c.bx, y: 4.161, w: 1.874, h: 1.616, align: 'center' });
	});
	footer(s, { bold: true, ecgLine: '4472C4', rule: '4472C4' });
}

/* 4 — statistics cards */
function slide04(pptx) {
	const s = pptx.addSlide();
	title(s, 'Because Saving Lives Is Serious Business', { x: 1.233, y: 1.191, w: 6.313, h: 1.447 });
	txt(s, 'Wellness Medical', { x: 1.233, y: 3.33, w: 2.35, h: 0.37, fontSize: 16, color: INDIGO });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscig elit. Pellentesque scelerisque larakanwo dadalerei malesuadloo libero ellentesque. Morbi orciwkante dui,galang fermentum eget lectus ornare, viverral dignissim risus. Done ellentesque. Porbi orcipawa', { x: 1.233, y: 3.794, w: 3.232, h: 2.181 });
	[
		{ x: 8.559, y: 1.063, adj: 6829, big: '5630+', label: 'Medical Clinic' },
		{ x: 8.559, y: 3.478, adj: 5487, big: '772K', label: 'Health Check' },
		{ x: 4.771, y: 3.478, adj: 7276, big: '998M', label: 'Medicine Labs' },
	].forEach(c => {
		card(s, c.x, c.y, 3.564, 2.259, c.adj);
		txt(s, c.big, { x: c.x + 0.391, y: c.y + 0.212, w: 1.595, h: 0.572, fontSize: 28, color: INDIGO });
		txt(s, c.label, { x: c.x + 0.391, y: c.y + 0.776, w: 2.006, h: 0.37, fontSize: 16, color: SILVER });
		body(s, LOREM_D, { x: c.x + 0.391, y: c.y + 1.182, w: 2.953, h: 1.01 });
	});
	footer(s);
}

/* 5 — "Say Yes To Your Good Health" */
function slide05(pptx) {
	const s = pptx.addSlide();
	title(s, 'Say Yes To Your Good Health', { x: 1.233, y: 1.619, w: 5.151, h: 1.447 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisquew larakanwo malesuadloo libero ellentesque. Morbi orci dui,galang fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi orci duadad fuange adaadvefwati,galang fermentum egeti amet, consectetur adipiscing elit babayowel', { x: 1.233, y: 3.267, w: 5.151, h: 1.575 });
	gradient(s, { x: 1.279, y: 5.122, w: 2.098, h: 0.579, from: CYAN, to: INDIGO, dir: 'h', radius: rr(16667, 2.098, 0.579), bands: 18 });
	txt(s, 'Learn More', { x: 1.425, y: 5.221, w: 1.806, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
	s.addShape('flowChartConnector', { x: 6.667, y: 3.988, w: 2.098, h: 2.098, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
	txt(s, '876+', { x: 6.874, y: 4.503, w: 1.683, h: 0.707, fontSize: 36, color: INDIGO, align: 'center' });
	txt(s, 'Surgery Patient', { x: 6.813, y: 5.122, w: 1.806, h: 0.337, fontSize: 14, color: GREY, align: 'center' });
	footer(s);
}

/* 6 — "Making You Better Every Time" + price notes */
function slide06(pptx) {
	const s = pptx.addSlide();
	title(s, 'Making You Better Every Time', { x: 5.071, y: 1.143, w: 3.409, h: 2.121 });
	txt(s, 'Wellness Medical', { x: 5.071, y: 3.834, w: 2.35, h: 0.37, fontSize: 16, color: INDIGO });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscig eli kitabi Pellentesque scelerisque larakanwo dadalerei malesuadloo libero ellentesque. Morbi orciwkante dui,galang fermentum  lectus ornare, viverral scelerisque larakanwo dadalerei bero', { x: 5.071, y: 4.298, w: 3.936, h: 1.575 });
	[['$998', 1.272, 1.706], ['$653', 2.735, 3.169]].forEach(p => {
		txt(s, p[0], { x: 9.596, y: p[1], w: 1.397, h: 0.572, fontSize: 28, color: INDIGO });
		body(s, LOREM_B, { x: 9.636, y: p[2], w: 2.243, h: 0.707 });
	});
	gradient(s, { x: 9.775, y: 4.435, w: 2.451, h: 1.432, from: CYAN, to: INDIGO, dir: 'bl', radius: rr(8202, 2.451, 1.432), shadow: true, bands: 20 });
	iconPlate(s, { x: 9.636, y: 4.149, d: 0.572, plate: WHITE, fg: INDIGO, icon: 'kit', iconSize: 0.29 });
	txt(s, 'Medicine Price', { x: 10.243, y: 4.573, w: 1.861, h: 0.37, fontSize: 16, color: WHITE });
	body(s, 'Lorem ipsu conse lectus ornare, ', { x: 10.243, y: 4.796, w: 1.451, h: 0.707, color: WHITE });
	footer(s);
}

/* 7 — full-bleed quote */
function slide07(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 0, y: 0, w: 13.333, h: 7.5, from: CYAN, to: INDIGO, dir: 'bl', bands: 26, bleed: true });
	card(s, 1.774, 1.6, 9.798, 4.423, 2627);
	title(s, 'Do not let your continued success become a burden to your customer\u2019s health', { x: 2.761, y: 2.819, w: 4.887, h: 1.986, fontSize: 28 });
	txt(s, 'Dr. Leony Ira', { x: 8.761, y: 4.986, w: 1.879, h: 0.37, fontSize: 16, color: DARK, align: 'center' });
	body(s, 'Job Position', { x: 9.032, y: 5.212, w: 1.335, h: 0.471, fontSize: 11, color: GREY, align: 'center' });
	s.addShape('flowChartConnector', { x: 2.761, y: 0.53, w: 1.879, h: 1.879, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
	icon(s, 'quote', 3.246, 1.186, 0.909, 0.567, INDIGO, WHITE);
	footer(s, { color: WHITE, ecgFill: WHITE, ecgLine: null });
}

/* 8 — four services list */
function slide08(pptx) {
	const s = pptx.addSlide();
	title(s, 'Making You Better Every Time', { x: 1.233, y: 1.45, w: 3.409, h: 2.121, color: '262626' });
	body(s, LOREM_F, { x: 1.233, y: 3.976, w: 3.604, h: 2.181 });
	[
		{ y: 1.107, icon: 'kit', label: 'Medical Check', ix: 5.392, iy: 1.113, iw: 0.431, ih: 0.412 },
		{ y: 2.504, icon: 'snake', label: 'Medical Clinic', ix: 5.412, iy: 2.434, iw: 0.392, ih: 0.553 },
		{ y: 3.977, icon: 'chart', label: 'Medical Analysis', ix: 5.423, iy: 3.856, iw: 0.392, ih: 0.551 },
		{ y: 5.367, icon: 'bed', label: 'Medical Treat ', ix: 5.421, iy: 5.357, iw: 0.469, ih: 0.456 },
	].forEach(r => {
		icon(s, r.icon, r.ix, r.iy, r.iw, r.ih, INDIGO, WHITE);
		txt(s, r.label, { x: 6.06, y: r.y, w: 2.085, h: 0.37, fontSize: 16, color: MIDGREY });
		body(s, LOREM_D, { x: 6.06, y: r.y + 0.426, w: 3.068, h: 0.707 });
	});
	footer(s);
}

/* 9 — four habit cards + right hand column */
function slide09(pptx) {
	const s = pptx.addSlide();
	title(s, 'Improving Lives Together', { x: 2.846, y: 1.217, w: 7.641, h: 0.774, color: '262626', align: 'center' });
	[
		{ x: 1.003, y: 2.621, w: 3.207, adj: 8794, icon: 'kit', is: [0.406, 0.388], text: 'Make it a\nhabit to eat nutritious food.', tx: 2.359, tw: 1.822 },
		{ x: 4.445, y: 2.621, w: 3.207, adj: 9450, icon: 'bed', is: [0.406, 0.394], text: 'Get enough rest and wake up in the morning.', tx: 5.794, tw: 1.697 },
		{ x: 1.003, y: 4.503, w: 3.16, adj: 8794, icon: 'hand', is: [0.406, 0.404], text: 'Healthcare that is convenient', tx: 2.36, tw: 1.415 },
		{ x: 4.445, y: 4.503, w: 3.207, adj: 10106, icon: 'snake', is: [0.36, 0.508], text: 'Keep wearing a mask when outside', tx: 5.794, tw: 1.411 },
	].forEach(c => {
		card(s, c.x, c.y, c.w, 1.604, c.adj);
		iconPlate(s, { x: c.x + 0.284, y: c.y + 0.221, d: 0.787, plate: 'F2F2F2', fg: INDIGO, icon: c.icon, iconSize: c.is[0] });
		txt(s, c.text, { x: c.tx, y: c.y + 0.262, w: c.tw, h: 0.707, fontFace: FONT_MED, fontSize: 12, color: DARK });
		txt(s, 'Lorem ipsum', { x: c.tx, y: c.y + 1.099, w: 1.411, h: 0.286, fontSize: 11, color: LIGHTGREY });
	});
	txt(s, 'Wellness Medical', { x: 8.394, y: 2.937, w: 2.35, h: 0.37, fontSize: 16, color: INDIGO });
	body(s, [
		{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscig eli kitabi bullei Pellentesque scelerisque larakanwo dadalerei malesuadloo libero ellentesque. Morbi orciwkante dui,galang fermentum  lectus orna viverral scelerisque larakanwo dadalerei bero scelerisque larakan', options: { breakLine: true } },
		{ text: '', options: { breakLine: true } },
		{ text: 'Korem' + LOREM_E.slice(5) },
	], { x: 8.394, y: 3.4, w: 4.179, h: 3.09 });
	footer(s);
}

/* 10 — testimonials on a full-bleed gradient */
function slide10(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 0, y: 0, w: 13.333, h: 7.5, from: '02B1D5', to: '4950AC', dir: 'bl', bands: 26, bleed: true });
	title(s, 'Wellness Patient Testimonial', { x: 0.831, y: 1.586, w: 4.748, h: 1.447, color: WHITE });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscig eli kitabi Pellentesque scelerisque larakanwo dadalerei malesuadloo libero ellentesque. Morbi orciwkante dui,galang fermentum  lectus ornare, viverral scelerisque larakan', { x: 0.831, y: 3.681, w: 5.003, h: 0.969, color: WHITE });
	[
		{ y: 0.927, adj: 5547, quote: '\u201CYou guys are blowing me away with the quality, commitment and passion of your collective work!\u201D', name: 'Lauis Ismael', tx: 9.39, ny: 2.721 },
		{ y: 3.823, adj: 4157, quote: '\u201CThey always treat me with tender loving care. Because they really do care.\u201D', name: 'Michele Ikwan', tx: 9.291, ny: 5.577 },
	].forEach(c => {
		gradient(s, { x: 6.666, y: c.y, w: 5.536, h: 2.522, from: 'D9D9D9', to: WHITE, dir: 'h', radius: rr(c.adj, 5.536, 2.522), shadow: true, bands: 22 });
		txt(s, c.quote, { x: c.tx, y: c.y + 0.384, w: 2.5, h: 1.313, fontSize: 12, lineSpacingMultiple: 1.5 });
		txt(s, c.name, { x: c.tx, y: c.ny, w: 1.38, h: 0.404, fontSize: 12, lineSpacingMultiple: 1.5, color: RED });
	});
	footer(s, { label: 'Medical Presentation', font: 'Inter', bold: true, color: WHITE, ecgFill: WHITE, ecgLine: null });
}

/* 11 — "Improving Lives Together" with wide feature bar */
function slide11(pptx) {
	const s = pptx.addSlide();
	title(s, 'Improving Lives Together', { x: 7.661, y: 1.27, w: 4.448, h: 1.447 });
	body(s, LOREM_C, { x: 7.703, y: 2.907, w: 4.477, h: 1.616 });
	card(s, 5.349, 4.783, 6.829, 1.253, 7703);
	iconPlate(s, { x: 5.735, y: 4.948, d: 0.911, plate: INDIGO, fg: WHITE, icon: 'kit', iconSize: 0.47 });
	txt(s, 'Medicine Labs', { x: 6.867, y: 4.989, w: 1.88, h: 0.37, fontSize: 16, bold: true, color: INDIGO });
	body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo sceler conse lectus', { x: 6.867, y: 5.325, w: 5.242, h: 0.707 });
	footer(s);
}

/* 12 — "Quality Healthcare" with bullet list */
function slide12(pptx) {
	const s = pptx.addSlide();
	title(s, 'Quality Healthcare', { x: 1.233, y: 3.661, w: 5.434, h: 0.774 });
	body(s, LOREM_I, { x: 1.233, y: 4.617, w: 5.976, h: 1.575 });
	txt(s, 'Wellness Medical', { x: 8.372, y: 3.745, w: 2.292, h: 0.37, fontSize: 16, color: INDIGO });
	txt(s, BULLETS, { x: 8.372, y: 4.115, w: 3.967, h: 1.313, fontSize: 16, lineSpacingMultiple: 1.5 });
	footer(s);
}

/* 13 — "We Always Ready" with two progress bars */
function slide13(pptx) {
	const s = pptx.addSlide();
	title(s, 'We Always Ready For Your Health', { x: 1.233, y: 1.369, w: 5.281, h: 1.447 });
	body(s, [
		{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisque larakanwo malesuadloo libero ellentesque. Morbi orci dui,galang fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi orci duadad adaadvefwati,galang fermentum egetia ' },
		{ text: 'consectetur adipiscing elit. Pellentesque low scelerisque kataka malesuadloo liberokatakaloe', options: { color: NEARBLACK } },
	], { x: 1.233, y: 3.225, w: 5.646, h: 1.575 });
	gradient(s, { x: 1.279, y: 5.081, w: 2.098, h: 0.579, from: CYAN, to: INDIGO, dir: 'h', radius: rr(16667, 2.098, 0.579), bands: 18 });
	txt(s, 'Learn More', { x: 1.425, y: 5.179, w: 1.806, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
	body(s, 'Polem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque low scelerisque kataka malesuadloo liberokataka ellentesque. Morbi orcilapo', { x: 7.682, y: 3.733, w: 4.473, h: 0.969, color: NEARBLACK });
	[{ y: 5.162, ly: 4.762, py: 4.76, pct: '73%', fw: 3.22 }, { y: 5.875, ly: 5.474, py: 5.473, pct: '82%', fw: 3.658 }].forEach(b => {
		s.addShape('rect', { x: 7.755, y: b.y, w: 4.262, h: 0.152, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
		gradient(s, { x: 7.755, y: b.y, w: b.fw, h: 0.152, from: CYAN, to: INDIGO, dir: 'h', bands: 22 });
		txt(s, 'Text Here', { x: 7.682, y: b.ly, w: 2.097, h: 0.303, fontSize: 12, color: NEARBLACK });
		txt(s, b.pct, { x: 10.004, y: b.py, w: 2.097, h: 0.303, fontSize: 12, color: NEARBLACK, align: 'right' });
	});
	footer(s);
}

/* 14 — three-column band */
function slide14(pptx) {
	const s = pptx.addSlide();
	footer(s);
	gradient(s, { x: -0.012, y: 4.257, w: 13.345, h: 2.057, from: INDIGO, to: CYAN, dir: 'h', bands: 40 });
	[
		{ x: 1.233, ty: 4.6, label: 'Health Threat', by: 4.996 },
		{ x: 4.959, ty: 4.585, label: 'Medical Check', by: 4.981 },
		{ x: 8.686, ty: 4.6, label: 'Hospital Check', by: 4.996 },
	].forEach(c => {
		txt(s, c.label, { x: c.x, y: c.ty, w: 2.085, h: 0.37, fontFace: FONT_MED, fontSize: 16, color: WHITE });
		body(s, LOREM_J, { x: c.x, y: c.by, w: 3.481, h: 0.969, color: WHITE });
	});
	gradient(s, { x: 0, y: 0.021, w: 13.345, h: 0.5, from: INDIGO, to: CYAN, dir: 'h', bands: 40 });
}

/* 15 — "We Take Care Of Your Healthy Health" + outlined stat card */
function slide15(pptx) {
	const s = pptx.addSlide();
	title(s, 'We Take Care Of Your Healthy Health', { x: 1.233, y: 1.319, w: 6.125, h: 1.447 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisque malesuadloo liberokataka ellentesque. Morbi orci dui,galang fermentum egetawo lectus ornare, viverra dignissim risus. Done agustus ellentesqueI Morbi orci duadad katak ada akarrn akaoi, galang fermentum eget lectus lacinia dolor nequea', { x: 1.233, y: 2.961, w: 6.517, h: 1.272 });
	statCard(s, { x: 8.417, y: 1.379, h: 4.298, border: CYAN, fill: WHITE, rows: [0.292, 1.064, 1.765, 2.185, 2.611] });
	footer(s, { color: WHITE, ecgFill: WHITE, ecgLine: null, label: 'Medical Presentation', font: 'Inter', bold: true });
}

/** Outlined "876+ / Surgery Patient / Wellness Medicine" card (slides 15, 17, 23). */
function statCard(s, o) {
	const f = o.font || FONT;
	const accent = o.accent || INDIGO;
	const label = o.label || INK;
	s.addShape('roundRect', { x: o.x, y: o.y, w: 3.727, h: o.h, rectRadius: rr(4522, 3.727, o.h), fill: { color: o.fill }, line: { color: o.border, width: 1.5 } });
	txt(s, o.value || '876+', { x: o.x + 0.381, y: o.y + o.rows[0], w: 1.799, h: 0.774, fontFace: f, fontSize: 40, bold: true, color: accent });
	txt(s, 'Surgery Patient', { x: o.x + 0.381, y: o.y + o.rows[1], w: 2.085, h: 0.37, fontFace: f, fontSize: 16, bold: true, color: label });
	s.addShape('line', { x: o.x + 0.527, y: o.y + o.rows[2], w: 2.653, h: 0, line: { color: o.rule || accent, width: 1 } });
	txt(s, 'Wellness Medicine', { x: o.x + 0.381, y: o.y + o.rows[3], w: 2.423, h: 0.37, fontFace: f, fontSize: 16, bold: true, color: label });
	body(s, LOREM_H, { x: o.x + 0.381, y: o.y + o.rows[4], w: 3.008, h: 1.313, fontFace: o.bodyFont || FONT });
}

/* 16 — "How Hospitals Use Content Marketing to Attract Patients" */
function slide16(pptx) {
	const s = pptx.addSlide();
	txt(s, 'A B O U T  U S', { x: 1.26, y: 1.528, w: 1.514, h: 0.286, fontSize: 11, bold: true, color: MIDGREY });
	title(s, 'How Hospitals Use Content Marketing to Attract Patients', { x: 1.233, y: 1.962, w: 5.831, h: 2.121 });
	gradient(s, { x: 1.291, y: 4.639, w: 5.004, h: 1.432, from: CYAN, to: INDIGO, dir: 'bl', radius: rr(8202, 5.004, 1.432), shadow: true, bands: 26 });
	iconPlate(s, { x: 1.576, y: 4.896, d: 0.911, plate: WHITE, fg: INDIGO, icon: 'kit', iconSize: 0.47 });
	txt(s, 'Medicine Price', { x: 2.706, y: 4.825, w: 2.267, h: 0.37, fontSize: 16, bold: true, color: WHITE });
	body(s, 'Lorem ipsu conse lectus ornare, pellentese viverra ctetuwwlo conse lectus ornare. Ipsu', { x: 2.706, y: 5.161, w: 3.042, h: 0.707, color: WHITE });
	body(s, LOREM_G, { x: 6.553, y: 4.616, w: 2.992, h: 1.27 });
	body(s, LOREM_G + ' dui', { x: 9.723, y: 4.618, w: 2.992, h: 1.27 });
	footer(s, { bold: true });
}

/* 17 — "Find Medical Check Up Servicess" over a cyan/indigo wash */
function slide17(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 0, y: 0, w: 13.333, h: 7.5, from: '40C6E1', to: '484FAC', dir: 'v', bands: 40 });
	title(s, 'Find Medical Check Up Servicess', { x: 1.233, y: 1.909, w: 5.831, h: 1.447, color: WHITE });
	const col = 'Lorem ipsum dolor sit amet, consectetur adipiscig elita Pellentesque scelerisque larakanwo dadalerei dararele malesuadloo libero ellentesque. Morbiale orciwkanteloeri dui,galang fermentum eget lectusi ornare, viverral gune dignissim risus. Done';
	body(s, col + 'eremeo', { x: 1.233, y: 3.665, w: 3.242, h: 1.878, color: WHITE });
	body(s, col + 'remero', { x: 4.497, y: 3.665, w: 3.242, h: 1.878, color: WHITE });
	statCard(s, { x: 8.417, y: 1.68, h: 3.936, border: INK, fill: WHITE, rule: INK, rows: [0.282, 1.054, 1.702, 2.017, 2.443] });
	footer(s, { bold: true, color: WHITE, ecgFill: WHITE, ecgLine: null });
}

/* 18 — "Improving Lives Together" (right-hand picture frames are empty in the source) */
function slide18(pptx) {
	const s = pptx.addSlide();
	title(s, 'Improving Lives Together', { x: 1.24, y: 1.217, w: 4.448, h: 1.447 });
	body(s, LOREM_C, { x: 1.282, y: 2.828, w: 4.477, h: 1.616 });
	gradient(s, { x: 1.212, y: 4.691, w: 4.793, h: 1.432, from: CYAN, to: INDIGO, dir: 'bl', radius: rr(8202, 4.793, 1.432), shadow: true, bands: 24 });
	iconPlate(s, { x: 1.676, y: 4.949, d: 0.911, plate: WHITE, fg: INDIGO, icon: 'hand', iconSize: 0.41 });
	txt(s, 'Medicine Price', { x: 2.848, y: 4.878, w: 2.382, h: 0.37, fontSize: 16, bold: true, color: WHITE });
	body(s, 'Lorem ipsu conse lectus ornare, pellentesque. ', { x: 2.848, y: 5.214, w: 3.042, h: 0.707, color: WHITE });
	footer(s);
}

/* 19 — contact panel */
function slide19(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 4.331, y: 1.386, w: 8.242, h: 4.727, from: INDIGO, to: CYAN, dir: 'tl', bands: 34, shadow: true });
	txt(s, 'CONTACT US', { x: 4.987, y: 1.838, w: 2.879, h: 0.505, fontSize: 24, bold: true, color: WHITE });
	body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo remaei adipiscing elit. Pellentesque sceler conse lectus ornare galang tarabintar fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi ori', { x: 4.987, y: 2.343, w: 6.93, h: 1.01, color: WHITE });
	s.addShape('line', { x: 5.073, y: 3.543, w: 6.561, h: 0, line: { color: WHITE, width: 1 } });
	[
		{ x: 4.987, y: 3.926, head: 'Office Hours', hw: 1.824, text: 'Lorem ipsu conse lectus ornare, ', tw: 2.541 },
		{ x: 7.908, y: 3.926, head: 'Office Address', hw: 2.157, text: 'Jl. Anila 2 No.00 Sawojajar, Malang Jawa Timur', tw: 3.463 },
		{ x: 4.987, y: 4.888, head: 'Office Contact', hw: 2.157, text: '+62 0808 0808 / +62 8080 8080', tw: 2.541 },
		{ x: 7.908, y: 4.888, head: 'Office Website', hw: 2.157, text: 'www.wellness.com', tw: 3.018 },
	].forEach(c => {
		txt(s, c.head, { x: c.x, y: c.y, w: c.hw, h: 0.37, fontSize: 16, bold: true, color: WHITE });
		body(s, c.text, { x: c.x, y: c.y + 0.315, w: c.tw, h: 0.438, fontSize: 10, color: WHITE });
	});
	footer(s, { bold: true });
}

/* 20 — meet the doctors */
function slide20(pptx) {
	const s = pptx.addSlide();
	title(s, 'Meet the Profesional Doctor', { x: 5.184, y: 1.27, w: 5.614, h: 1.447 });
	gradient(s, { x: 0.76, y: 1.424, w: 3.618, h: 4.661, from: CYAN, to: INDIGO, dir: 'v', radius: rr(5570, 3.618, 4.661), shadow: true, bands: 24 });
	txt(s, 'Dr. Michele Aly', { x: 1.397, y: 4.851, w: 2.207, h: 0.404, fontSize: 18, bold: true, color: WHITE, align: 'center' });
	body(s, 'Doctor Expert', { x: 1.636, y: 5.043, w: 1.73, h: 0.572, fontSize: 14, color: WHITE, align: 'center' });
	[
		{ x: 4.484, name: 'Dr. Luna Mayo' },
		{ x: 7.216, name: 'Dr. Alex Harrys' },
		{ x: 9.947, name: 'Dr. Yuli Simple' },
	].forEach(c => {
		card(s, c.x, 3.131, 2.626, 2.954, 5570);
		txt(s, c.name, { x: c.x + 0.209, y: 5.117, w: 2.207, h: 0.404, fontSize: 18, bold: true, color: INDIGO, align: 'center' });
		body(s, 'Doctor', { x: c.x + 0.448, y: 5.308, w: 1.73, h: 0.572, fontSize: 14, color: INK, align: 'center' });
	});
	footer(s);
}

/* 21 — three service cards */
function slide21(pptx) {
	const s = pptx.addSlide();
	title(s, 'We Always Ready For Your Health', { x: 1.233, y: 1.619, w: 5.281, h: 1.447 });
	body(s, [
		{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ganae scelerisque larakanwo malesuadloo libero ellentesque. Morbi orci dui,galan fermentum eget lectus ornare, viverra dignissim risus. Done ellentesq Porbi orci duadad adaadvefwati,galang fermentum egetia ' },
		{ text: 'consectetur adipiscingi', options: { color: NEARBLACK } },
	], { x: 1.233, y: 3.345, w: 4.727, h: 1.575 });
	gradient(s, { x: 1.279, y: 5.122, w: 2.098, h: 0.579, from: CYAN, to: INDIGO, dir: 'v', radius: rr(16667, 2.098, 0.579), bands: 10 });
	txt(s, 'Learn More', { x: 1.425, y: 5.247, w: 1.806, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
	[
		{ y: 1.581, adj: 9414, icon: 'bed', is: 0.469, label: 'Wellness Medical', lw: 2.336, ty: 1.787, by: 2.123 },
		{ y: 3.04, adj: 9414, icon: 'kit', is: 0.469, label: 'Medicine Labs', lw: 1.942, ty: 3.246, by: 3.582 },
		{ y: 4.496, adj: 11026, icon: 'syringe', is: 0.483, label: 'Voccine Labs', lw: 1.942, ty: 4.739, by: 5.075 },
	].forEach(c => {
		card(s, 6.767, c.y, 5.49, 1.253, c.adj);
		iconPlate(s, { x: 7.028, y: c.y + 0.165, d: 0.911, gradient: [CYAN, INDIGO], fg: WHITE, icon: c.icon, iconSize: c.is });
		txt(s, c.label, { x: 8.2, y: c.ty, w: c.lw, h: 0.37, fontSize: 16, bold: true, color: INDIGO });
		body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuw.', { x: 8.2, y: c.by, w: 4.029, h: 0.404 });
	});
	footer(s);
}

/* 22 — phone mock-up */
function slide22(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 7.541, y: 0, w: 5.792, h: 7.5, from: INDIGO, to: CYAN, dir: 'tl', bands: 22 });
	imagePlaceholder(s, 8.969, 0.9, 2.771, 5.589, '[image]');
	title(s, 'Improving Lives Together', { x: 1.167, y: 1.27, w: 4.448, h: 1.447 });
	body(s, LOREM_C, { x: 1.228, y: 2.828, w: 4.477, h: 1.616 });
	gradient(s, { x: 1.228, y: 4.639, w: 4.793, h: 1.432, from: CYAN, to: INDIGO, dir: 'bl', radius: rr(8202, 4.793, 1.432), shadow: true, bands: 24 });
	iconPlate(s, { x: 1.692, y: 4.896, d: 0.911, plate: WHITE, fg: INDIGO, icon: 'hand', iconSize: 0.41 });
	txt(s, 'Medicine Price', { x: 2.864, y: 4.825, w: 2.215, h: 0.37, fontSize: 16, bold: true, color: WHITE });
	body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetuwwlo conse lectu', { x: 2.864, y: 5.161, w: 3.042, h: 0.707, color: WHITE });
	footer(s);
}

/* 23 — red "Wellness" brand variant */
function slide23(pptx) {
	const s = pptx.addSlide();
	imagePlaceholder(s, 5.106, 0.854, 2.762, 5.57, '[image]');
	// "W" wordmark: two slanted strokes
	s.addShape('triangle', { x: 0.517, y: 0.5, w: 0.28, h: 0.354, flipV: true, fill: { color: 'D15D60' }, line: { type: 'none' } });
	s.addShape('triangle', { x: 0.723, y: 0.5, w: 0.28, h: 0.354, flipV: true, fill: { color: RED }, line: { type: 'none' } });
	txt(s, 'Wellness', { x: 1.003, y: 0.536, w: 1.014, h: 0.303, fontFace: 'Inter', fontSize: 12, bold: true, color: RED });
	title(s, 'Making You Better Every Time', { x: 1.233, y: 1.493, w: 3.409, h: 2.121, fontFace: 'Inter', bold: true, color: RED });
	body(s, LOREM_F, { x: 1.233, y: 4.019, w: 3.604, h: 1.875, fontFace: 'Open Sans' });
	statCard(s, {
		x: 8.477, y: 2.337, h: 3.936, border: CYAN, fill: 'F2F2F2', value: '599+',
		font: 'Inter', bodyFont: 'Open Sans', accent: RED, label: RED, rule: RED,
		rows: [0.229, 1.002, 1.702, 2.122, 2.548],
	});
	footer(s, { label: 'Medical Presentation', font: 'Inter', bold: true, color: RED, rule: 'CB474A', ecgFill: RED, ecgLine: null });
}

/* 24 — desktop monitor mock-up */
function slide24(pptx) {
	const s = pptx.addSlide();
	imagePlaceholder(s, 6.266, 1.329, 5.831, 4.843, '[image]');
	s.addShape('flowChartConnector', { x: 10.51, y: 0.679, w: 2.098, h: 2.098, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
	txt(s, '876+', { x: 10.717, y: 1.194, w: 1.683, h: 0.707, fontSize: 36, bold: true, color: INDIGO, align: 'center' });
	txt(s, 'Surgery Patient', { x: 10.656, y: 1.813, w: 1.806, h: 0.337, fontSize: 14, bold: true, color: GREY, align: 'center' });
	txt(s, 'Wellness Medical', { x: 1.236, y: 1.443, w: 2.35, h: 0.37, fontSize: 16, color: INK });
	body(s, [
		{ text: LOREM_E + 'PLACEHOLDER', options: { breakLine: true } },
		{ text: '', options: { breakLine: true } },
		{ text: 'Korem' + LOREM_E.slice(5) },
	], { x: 1.236, y: 1.907, w: 4.179, h: 3.09 });
	card(s, 1.236, 4.867, 5.726, 1.253, 16667);
	iconPlate(s, { x: 1.497, y: 5.032, d: 0.911, gradient: [CYAN, INDIGO], fg: WHITE, icon: 'kit', iconSize: 0.47 });
	txt(s, 'Medicine Labs', { x: 2.668, y: 5.074, w: 1.88, h: 0.37, fontSize: 16, bold: true, color: INDIGO });
	body(s, LOREM_B, { x: 2.668, y: 5.41, w: 4.029, h: 0.707 });
	footer(s);
}

/* 25 — laptop mock-up + progress bars + bullets */
function slide25(pptx) {
	const s = pptx.addSlide();
	imagePlaceholder(s, 0.824, 1.164, 5.53, 3.201, '[image]');
	[{ y: 5.235, fw: 3.562, ly: 4.834, py: 4.832, pct: '73%' }, { y: 5.947, fw: 3.999, ly: 5.547, py: 5.545, pct: '82%' }].forEach(b => {
		s.addShape('rect', { x: 1.702, y: b.y, w: 4.262, h: 0.152, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
		gradient(s, { x: 1.361, y: b.y, w: b.fw, h: 0.152, from: CYAN, to: INDIGO, dir: 'h', bands: 22 });
		txt(s, 'Text Here', { x: 1.283, y: b.ly, w: 2.097, h: 0.303, fontSize: 12, color: NEARBLACK });
		txt(s, b.pct, { x: 3.951, y: b.py, w: 2.097, h: 0.303, fontSize: 12, color: NEARBLACK, align: 'right' });
	});
	title(s, 'Desktop Mockup', { x: 7.189, y: 1.566, w: 5.281, h: 0.774 });
	body(s, [
		{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque scelerisquer larakanwo malesuadloo libero ellentesque. Morbi orci dui,galang fermentum eget lectus ornare, viverra dignissim risus. Done ellentesque. Porbi orci duadad ganabi adaadvefwati,galang fermentum egetia ' },
		{ text: 'consectetur adipiscing elit. Pellentesque lo', options: { color: NEARBLACK } },
	], { x: 7.189, y: 2.554, w: 5.281, h: 1.575 });
	txt(s, BULLETS, { x: 7.228, y: 4.231, w: 3.967, h: 1.717, fontSize: 16, lineSpacingMultiple: 2 });
	footer(s);
}

/* Shared chart look: Poppins labels, no legend/title, plain grid. */
const CHART_BASE = {
	chartColors: [INDIGO, CYAN, 'BFBFBF'],
	showLegend: false, showTitle: false, showValue: false,
	catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
	catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
	catAxisLineShow: false, valAxisLineShow: false,
	valGridLine: { color: 'BFBFBF', size: 0.75 }, catGridLine: { style: 'none' },
	valAxisMinVal: 0, valAxisMajorUnit: 1,
	dataBorder: { pt: 0, color: 'FFFFFF' },
};

/* 26 — "Wellness Medical Chart" (2-category clustered bars) */
function slide26(pptx) {
	const s = pptx.addSlide();
	gradient(s, { x: 0.76, y: 1.341, w: 5.94, h: 4.481, from: INDIGO, to: CYAN, dir: 'h', radius: rr(3404, 5.94, 4.481), bands: 30 });
	txt(s, 'A B O U T  U S', { x: 1.523, y: 1.864, w: 1.514, h: 0.286, fontSize: 11, bold: true, color: WHITE });
	title(s, 'Wellness Medical Chart', { x: 1.481, y: 2.236, w: 4.448, h: 1.447, color: WHITE });
	body(s, LOREM_C, { x: 1.523, y: 3.873, w: 4.477, h: 1.616, bold: true, color: WHITE });
	s.addChart(pptx.ChartType.bar, [
		{ name: 'Series 1', labels: ['Category 1', 'Category 2'], values: [4.3, 2.5] },
		{ name: 'Series 2', labels: ['Category 1', 'Category 2'], values: [2.4, 4.4] },
		{ name: 'Series 3', labels: ['Category 1', 'Category 2'], values: [2.0, 2.0] },
	], Object.assign({}, CHART_BASE, {
		x: 7.9, y: 1.25, w: 4.6, h: 3.35,
		barGapWidthPct: 117, valAxisMaxVal: 5, catAxisLabelColor: '7F7F7F', valAxisLabelColor: '7F7F7F',
	}));
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elita lowam Pellentesque scelerisqu malesuadloo liberokataka kowa lowa ellentesque. Morbi orcika ta dui,galang fermentum low egetawoa lectus malesuadloo liberokataka kowao', { x: 7.81, y: 4.855, w: 4.763, h: 1.272 });
	footer(s);
}

/* 27 — two pie diagrams (drawn with `pie` shapes, as in the source) */
function slide27(pptx) {
	const s = pptx.addSlide();
	title(s, 'Find Medical Check Up Servicess', { x: 5.154, y: 1.267, w: 6.225, h: 1.447 });
	// each pie is two wedges: [x, y, d, startDeg, endDeg, fill] + its % label box
	[
		{ n: '01.', ny: 0.636, wedges: [
			{ x: 1.582, y: 0.667, d: 2.784, a: [116.69, 270], fill: INDIGO, label: '45%', lx: 1.955, ly: 1.874 },
			{ x: 1.744, y: 0.777, d: 2.646, a: [269.94, 115.78], fill: NEARBLACK, label: '55%', lx: 3.338, ly: 1.874 },
		] },
		{ n: '02.', ny: 3.592, wedges: [
			{ x: 1.637, y: 3.693, d: 2.683, a: [28.66, 270], fill: CYAN, label: '65%', lx: 2.117, ly: 5.204 },
			{ x: 1.675, y: 3.592, d: 2.784, a: [269.94, 28.78], fill: NEARBLACK, label: '35%', lx: 3.338, ly: 4.412 },
		] },
	].forEach(p => {
		txt(s, p.n, { x: 1.076, y: p.ny, w: 0.668, h: 0.37, fontSize: 16, color: INK });
		p.wedges.forEach(w => {
			s.addShape('pie', { x: w.x, y: w.y, w: w.d, h: w.d, angleRange: w.a, fill: { color: w.fill }, line: { type: 'none' } });
			txt(s, w.label, { x: w.lx, y: w.ly, w: 0.721, h: 0.37, fontSize: 16, color: WHITE, align: 'center' });
		});
	});
	[
		{ x: 5.154, ty: 3.88, label: 'Chart One' },
		{ x: 8.757, ty: 3.881, label: 'Chart Two' },
	].forEach(c => {
		txt(s, c.label, { x: c.x, y: c.ty, w: 1.88, h: 0.37, fontSize: 16, color: INDIGO });
		body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscig elita Pellentesque scelerisque larakanwo dadalerei dararele malesuadloo libero ellentesque. Morbi orciwkanteloeri dui,galang fermentum eget lectus ornare, viverral gune dignissim risus. Done ellentesque. Porbi orcipawa faan', { x: c.x, y: 4.32, w: 3.604, h: 1.878 });
	});
	footer(s);
}

/* 28 — "Medical Static" line chart */
function slide28(pptx) {
	const s = pptx.addSlide();
	title(s, 'Quality Healthcare', { x: 1.233, y: 0.907, w: 3.878, h: 1.447 });
	body(s, [
		{ text: LOREM_I, options: { breakLine: true } },
		{ text: '', options: { breakLine: true } },
		{ text: 'Pellentesque sceleri kataka kaw malesuadloo liberokataka ellentesque. Morbi orci dui,galang fermentum egetawo qolowo lectuawos lowowieo consectetur adipiscing elit katakan lolowanem katakanw kaowaa lowoa' },
	], { x: 1.233, y: 2.761, w: 4.323, h: 3.393 });
	txt(s, 'Medical Static ', { x: 6.64, y: 0.936, w: 2.085, h: 0.37, fontSize: 16, color: NEARBLACK });
	s.addChart(pptx.ChartType.line, [
		{ name: 'Series 1', labels: ['1', '2', '3', '4'], values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', labels: ['1', '2', '3', '4'], values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', labels: ['1', '2', '3', '4'], values: [2.0, 2.0, 3.0, 5.0] },
	], Object.assign({}, CHART_BASE, {
		x: 6.5, y: 1.3, w: 6.05, h: 5.15,
		chartColors: ['A6A6A6', INDIGO, '000000'], lineSize: 2,
		lineDataSymbol: 'none', valAxisMaxVal: 6,
		showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 8.5,
		catAxisLabelColor: '595959', valAxisLabelColor: '595959',
		catGridLine: { color: 'D9D9D9', size: 0.75 }, valGridLine: { color: 'D9D9D9', size: 0.75 },
	}));
	footer(s);
}

/* 29 — "Wellness Infographic" full-width bar chart */
function slide29(pptx) {
	const s = pptx.addSlide();
	title(s, 'Wellness Infographic', { x: 1.694, y: 1.086, w: 9.94, h: 0.707, fontSize: 36, align: 'center' });
	s.addChart(pptx.ChartType.bar, [
		{ name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.0, 2.0, 3.0, 5.0] },
	], Object.assign({}, CHART_BASE, {
		x: 0.75, y: 2.2, w: 11.85, h: 3.6,
		barGapWidthPct: 300, valAxisMaxVal: 5,
		catAxisLabelColor: '181717', valAxisLabelColor: '7F7F7F',
		valGridLine: { color: '808080', size: 0.75 },
	}));
	body(s, 'Lorem ipsu conse lectus ornare, pellentesque. viverra ctetur apiscing elit. Pellentesque sceler scelerisq dolor sit amet, conse malesuada Pellentesque', { x: 2.161, y: 5.851, w: 9.011, h: 0.666, align: 'center' });
	footer(s);
}

/* 30 — pricing table (the outer cards are rotated 90 degrees in the source,
   so their on-screen width/height are the swapped values used here) */
function slide30(pptx) {
	const s = pptx.addSlide();
	s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 3.779, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
	[
		{ x: 1.721, plan: 'Basic Package', price: '799', tx: 2.256, sx: 2.58, bx: 1.723, lines: PLAN_LINES },
		{ x: 8.468, plan: 'Gold Package', price: '999', tx: 9.013, sx: 9.337, bx: 8.48, lines: PLAN_LINES.slice(0, 2).concat(['For your family\u2019s wellness..']) },
	].forEach(c => {
		s.addShape('roundRect', { x: c.x, y: 1.88, w: 3.16, h: 4.262, rectRadius: rr(7716, 3.16, 3.16), fill: { color: WHITE }, line: { color: INK, width: 1 }, shadow: cardShadow() });
		txt(s, c.plan, { x: c.tx, y: 2.289, w: 2.085, h: 0.404, fontSize: 18, color: GREY, align: 'center' });
		txt(s, [{ text: '$', options: { fontSize: 20 } }, { text: c.price, options: { fontSize: 40 } }],
			{ x: c.tx, y: 3.005, w: 2.085, h: 0.774, bold: true, color: INDIGO, align: 'center' });
		txt(s, 'Service', { x: c.sx, y: 3.992, w: 1.437, h: 0.337, fontSize: 14, bold: true, color: GREY, align: 'center' });
		txt(s, c.lines.map(l => ({ text: l, options: { breakLine: true } })),
			{ x: c.bx, y: 4.408, w: 3.16, h: 1.212, fontSize: 11, lineSpacingMultiple: 2, align: 'center' });
	});
	// featured plan — taller, gradient filled
	gradient(s, { x: 4.869, y: 1.186, w: 3.597, h: 4.956, from: INDIGO, to: CYAN, dir: 'bl', radius: rr(5997, 3.597, 3.597), shadow: true, bands: 22 });
	txt(s, 'Premium Package', { x: 5.202, y: 1.852, w: 2.954, h: 0.404, fontSize: 18, color: WHITE, align: 'center' });
	txt(s, [{ text: '$', options: { fontSize: 28 } }, { text: '1099', options: { fontSize: 48 } }],
		{ x: 5.404, y: 2.568, w: 2.524, h: 0.909, bold: true, color: WHITE, align: 'center' });
	txt(s, 'Service', { x: 5.956, y: 3.555, w: 1.437, h: 0.337, fontSize: 14, bold: true, color: WHITE, align: 'center' });
	txt(s, PLAN_LINES.concat(['Our clinic, your advantage.']).map(l => ({ text: l, options: { breakLine: true } })),
		{ x: 5.099, y: 3.97, w: 3.16, h: 1.582, fontSize: 11, lineSpacingMultiple: 2, align: 'center', color: WHITE });
	footer(s);
}

/* 31 — world map with two pins */
function slide31(pptx) {
	const s = pptx.addSlide();
	worldMap(s, 1.72, 0.925, 12.179, 6.006);
	[
		{ x: 9.971, y: 1.701, sx: 10.088, sy: 2.545 },
		{ x: 7.477, y: 3.198, sx: 7.601, sy: 4.048 },
	].forEach(p => {
		s.addShape('ellipse', { x: p.sx, y: p.sy, w: 0.349, h: 0.117, fill: { color: SILVER, transparency: 40 }, line: { type: 'none' } });
		s.addShape('ellipse', { x: p.x, y: p.y, w: 0.597, h: 0.62, fill: { color: INDIGO }, line: { type: 'none' } });
		s.addShape('triangle', { x: p.x + 0.13, y: p.y + 0.48, w: 0.34, h: 0.37, flipV: true, fill: { color: INDIGO }, line: { type: 'none' } });
		icon(s, 'snake', p.x + 0.157, p.y + 0.149, 0.283, 0.4, WHITE, INDIGO);
	});
	title(s, 'Wellness Worldmap', { x: 1.233, y: 4.745, w: 3.878, h: 1.447 });
	footer(s);
}

/**
 * Simplified world map — each entry is a closed outline in 0..1 map coordinates
 * (x0,y0, x1,y1, ...), traced from the original vector artwork.
 */
const CONTINENTS = [
	[0.755,0.079, 0.783,0.102, 0.760,0.135, 0.813,0.123, 0.831,0.158, 0.865,0.136, 0.997,0.205, 0.987,0.221, 0.970,0.202, 0.965,0.238, 0.923,0.257, 0.903,0.331, 0.925,0.235, 0.842,0.299, 0.851,0.363, 0.822,0.418, 0.793,0.413, 0.804,0.482, 0.762,0.526, 0.758,0.598, 0.765,0.570, 0.756,0.588, 0.743,0.570, 0.754,0.640, 0.714,0.492, 0.684,0.602, 0.663,0.491, 0.659,0.510, 0.628,0.498, 0.604,0.468, 0.635,0.517, 0.589,0.573, 0.614,0.536, 0.586,0.552, 0.558,0.469, 0.590,0.588, 0.610,0.584, 0.565,0.787, 0.523,0.852, 0.509,0.650, 0.500,0.668, 0.496,0.609, 0.470,0.590, 0.464,0.619, 0.433,0.601, 0.422,0.523, 0.445,0.487, 0.422,0.520, 0.452,0.434, 0.466,0.457, 0.496,0.424, 0.513,0.462, 0.538,0.457, 0.537,0.530, 0.538,0.458, 0.562,0.463, 0.569,0.427, 0.541,0.404, 0.587,0.395, 0.577,0.355, 0.562,0.375, 0.556,0.358, 0.532,0.429, 0.507,0.366, 0.514,0.419, 0.495,0.375, 0.445,0.426, 0.457,0.344, 0.494,0.298, 0.531,0.341, 0.509,0.306, 0.551,0.256, 0.527,0.251, 0.532,0.203, 0.506,0.295, 0.484,0.238, 0.539,0.151, 0.583,0.186, 0.558,0.191, 0.572,0.223, 0.590,0.177, 0.601,0.195, 0.634,0.159, 0.660,0.182, 0.670,0.134, 0.671,0.202, 0.679,0.135, 0.689,0.155],
	[0.209,0.120, 0.216,0.181, 0.225,0.194, 0.230,0.166, 0.242,0.175, 0.207,0.220, 0.210,0.282, 0.246,0.329, 0.264,0.236, 0.314,0.310, 0.270,0.358, 0.288,0.338, 0.303,0.364, 0.286,0.382, 0.275,0.354, 0.244,0.502, 0.234,0.468, 0.206,0.470, 0.197,0.496, 0.172,0.459, 0.142,0.455, 0.110,0.331, 0.128,0.346, 0.090,0.262, 0.055,0.247, 0.044,0.264, 0.049,0.245, 0.014,0.299, 0.031,0.265, 0.006,0.252, 0.020,0.217, -0.000,0.209, 0.021,0.202, 0.004,0.176, 0.031,0.150, 0.069,0.163, 0.085,0.266, 0.075,0.168, 0.107,0.162, 0.098,0.177, 0.112,0.157, 0.168,0.198, 0.195,0.164, 0.203,0.192],
	[0.269,0.575, 0.268,0.594, 0.273,0.577, 0.286,0.590, 0.294,0.584, 0.327,0.623, 0.332,0.660, 0.336,0.648, 0.344,0.665, 0.359,0.661, 0.373,0.689, 0.355,0.773, 0.333,0.793, 0.318,0.850, 0.304,0.851, 0.308,0.874, 0.286,0.893, 0.279,1.000, 0.268,0.993, 0.274,0.975, 0.266,0.990, 0.262,0.984, 0.270,0.980, 0.262,0.981, 0.257,0.944, 0.272,0.752, 0.241,0.672, 0.251,0.605],
	[0.381,0.000, 0.410,0.013, 0.378,0.024, 0.408,0.021, 0.401,0.041, 0.436,0.031, 0.412,0.050, 0.416,0.119, 0.395,0.163, 0.408,0.163, 0.355,0.209, 0.348,0.257, 0.319,0.202, 0.328,0.159, 0.312,0.114, 0.275,0.095, 0.284,0.079, 0.266,0.066, 0.297,0.025, 0.316,0.016, 0.318,0.031, 0.328,0.013, 0.346,0.028, 0.338,0.008],
	[0.862,0.708, 0.894,0.814, 0.883,0.872, 0.873,0.882, 0.868,0.872, 0.865,0.881, 0.857,0.875, 0.849,0.849, 0.846,0.855, 0.848,0.840, 0.842,0.855, 0.831,0.831, 0.789,0.855, 0.779,0.799, 0.783,0.772, 0.800,0.763, 0.817,0.726, 0.827,0.735, 0.834,0.710, 0.846,0.716, 0.842,0.731, 0.857,0.747],
	[0.267,0.006, 0.301,0.016, 0.251,0.060, 0.257,0.075, 0.239,0.086, 0.252,0.092, 0.220,0.094, 0.222,0.077, 0.233,0.085, 0.224,0.067, 0.239,0.070, 0.226,0.045, 0.243,0.055, 0.249,0.037, 0.210,0.028],
	[0.149,0.454, 0.188,0.473, 0.204,0.541, 0.226,0.520, 0.220,0.560, 0.236,0.557, 0.237,0.596, 0.251,0.592, 0.251,0.605, 0.229,0.589, 0.223,0.565, 0.175,0.535, 0.154,0.463, 0.163,0.514, 0.148,0.486],
	[0.230,0.123, 0.228,0.147, 0.241,0.122, 0.242,0.142, 0.259,0.136, 0.299,0.196, 0.291,0.213, 0.279,0.201, 0.283,0.242, 0.250,0.214, 0.267,0.194, 0.276,0.200, 0.251,0.160, 0.219,0.157],
	[0.174,0.124, 0.186,0.172, 0.153,0.178, 0.141,0.163, 0.157,0.158, 0.139,0.157, 0.138,0.138, 0.149,0.126, 0.170,0.146],
	[0.833,0.649, 0.839,0.667, 0.849,0.656, 0.871,0.673, 0.885,0.708, 0.876,0.706, 0.867,0.690, 0.864,0.700, 0.849,0.696, 0.851,0.681, 0.835,0.671],
	[0.790,0.608, 0.797,0.619, 0.787,0.621, 0.785,0.639, 0.773,0.641, 0.785,0.639, 0.793,0.623, 0.796,0.640, 0.784,0.671, 0.773,0.666, 0.767,0.641],
];
function worldMap(s, x, y, w, h) {
	CONTINENTS.forEach(v => {
		const pts = [];
		for (let i = 0; i < v.length; i += 2) pts.push({ x: v[i] * w, y: v[i + 1] * h });
		pts.push({ close: true });
		s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: WHITE }, line: { color: 'BFBFBF', width: 0.75 } });
	});
}

/* 32 — break slide */
function slide32(pptx) {
	const s = pptx.addSlide();
	s.addShape('rect', { x: 0, y: 3.979, w: 8.453, h: 3.521, fill: { color: WHITE }, line: { type: 'none' } });
	title(s, 'Breakslide', { x: 1.139, y: 4.855, w: 6.175, h: 1.447, fontSize: 80, color: INDIGO });
	txt(s, 'Coffee Time Break 30 Minutes', { x: 1.139, y: 6.149, w: 4.538, h: 0.37, fontSize: 16, color: GREY });
}

/* ------------------------------------------------------------------ driver */
function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
	pptx.layout = 'WIDE';
	pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
	[
		slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
		slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
		slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
		slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
	].forEach(fn => fn(pptx));
	return pptx.writeFile({ fileName: path.join(__dirname, '0da3db76-c688-4e6f-b963-1a722511c24f_grok_final.pptx') });
}

build().then(f => console.log('wrote', f));
