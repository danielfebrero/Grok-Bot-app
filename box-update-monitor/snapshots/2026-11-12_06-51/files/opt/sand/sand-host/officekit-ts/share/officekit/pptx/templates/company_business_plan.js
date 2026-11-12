/**
 * Business Plan deck - rebuilt with pptxgenjs.
 *   node 0c0045a8-9f68-443c-a57b-3fa515d71101_grok_final.js
 * writes 0c0045a8-9f68-443c-a57b-3fa515d71101_grok_final.pptx next to this file.
 *
 * Raster artwork of the original deck is replaced by flat colour blocks tagged "[image]".
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '171717'; // theme tx2 / dk2
const PAPER = 'FFFFFF'; // theme bg2 / lt2
const LIME = 'B7FB06'; // theme accent1
const PHOTO = '6B8E8A'; // stand-in for the deck's photography

const HEAD = 'Dela Gothic One'; // theme major font
const BODY = 'Inter Tight'; // theme minor font

const TITLE_PT = 60;
const LABEL_PT = 12;
const TEXT_PT = 10;

/* ---------------------------------------------------------- reusable strings */

const LOREM = {
	cover: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Dolor sit aneanil commodorisui ligula eget dolor. Aenean massa. Cum dolor sit ametisul.',
	welcome:
		'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aeneanil commodorisui ligula eget dolor. Aenean massa. Cum sociis logolsi natoque penatibus et magnis discol parturient montes, nasceturis.',
	history:
		'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aeneanil commodorisui ligula eget dolor. Aenean massa. Cum sociis logolsi natoque penatibus et magnis discol parturient montes.',
	features:
		'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aeneanil commodorisui ligula eget dolor. Aenean massa. Cum sociis logolsi natoque penatibus et magnis discol parturient montes dolor sit amet consetectuer.',
	concept:
		'Lorem ipsum dolor sit amet, consectetuer dolors adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et.',
	swot: 'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet ligula eget.',
	short: 'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet.',
	growth: 'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula dolor sit amet consetectuer adispisc.',
	flexibility:
		'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet consetectuer adispiscing elit ligula eget dolor sit amet consetectuer adispisc..',
	quantifiable: 'Lorem ipsum dolor sitis sul aso amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor.',
	control:
		'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet consetectuer adispiscing dolor sit amet consetectuer.',
	priorities:
		'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet ligula eget dolor ist amet consetectuer adispising ligula eget dolor.',
	chart:
		'Lorem ipsum dolor sit amet, consectetueri ligula adipiscing elitisuli Aenean commodo ligula eget dolor. Aenean dolor sit amet ligula eget dolor ist amet logor.',
	thanks:
		'\u201CLorem ipsum dolor sit amet, consectetuer adipiscing elit. Dolor sit aneanil commodorisui ligula eget dolor. Aenean massa. Cum dolor sit ametisul ligula dolor sit amet consetectuer.\u201D',
};

/* -------------------------------------------------- hand traced vector icons */

// "double corner arrow" glyph used inside the pill buttons (unit square, y down)
const ARROW_GLYPH = [
	[0.62, 0.0], [0.0, 0.0], [0.0, 0.38], [0.352, 0.38], [0.0, 0.731],
	[0.268, 1.0], [0.62, 0.648], [0.62, 1.0], [1.0, 1.0], [1.0, 0.38], [1.0, 0.0],
];

// eight armed asterisk that sits inside the lime badge (unit square, y down)
const ASTERISK_GLYPH = [
	[1.0, 0.631], [1.0, 0.369], [0.816, 0.369], [0.946, 0.239], [0.761, 0.054],
	[0.631, 0.184], [0.631, 0.0], [0.369, 0.0], [0.369, 0.184], [0.239, 0.054],
	[0.054, 0.239], [0.184, 0.369], [0.0, 0.369], [0.0, 0.631], [0.184, 0.631],
	[0.054, 0.761], [0.239, 0.946], [0.369, 0.816], [0.369, 1.0], [0.631, 1.0],
	[0.631, 0.816], [0.761, 0.946], [0.946, 0.761], [0.816, 0.631],
];

function glyph(slide, unitPts, opts) {
	slide.addShape('custGeom', {
		x: opts.x,
		y: opts.y,
		w: opts.w,
		h: opts.h,
		rotate: opts.rotate || 0,
		fill: { color: opts.color },
		line: { width: 0 },
		points: unitPts
			.map(([ux, uy]) => ({ x: +(ux * opts.w).toFixed(4), y: +(uy * opts.h).toFixed(4) }))
			.concat([{ close: true }]),
	});
}

/* ------------------------------------------------------------- text helpers */

const titleText = (slide, text, x, y, w, h, color) =>
	slide.addText(text, { x, y, w, h, fontFace: HEAD, fontSize: TITLE_PT, color, valign: 'top' });

const labelText = (slide, text, x, y, w, h, color, align) =>
	slide.addText(text, {
		x, y, w, h,
		fontFace: HEAD, fontSize: LABEL_PT, color,
		lineSpacingMultiple: 1.5, valign: 'top', align: align || 'left',
	});

const bodyText = (slide, text, x, y, w, h, color, align) =>
	slide.addText(text, {
		x, y, w, h,
		fontFace: BODY, fontSize: TEXT_PT, color,
		lineSpacingMultiple: 1.5, valign: 'top', align: align === undefined ? 'justify' : align,
	});

const BULLET = { characterCode: '2022', indent: 13.5 }; // marL 0.1875" hanging indent

const bulletText = (slide, lines, x, y, w, h, color) =>
	slide.addText(
		lines.map((t) => ({ text: t, options: { bullet: BULLET, breakLine: true } })),
		{
			x, y, w, h,
			fontFace: BODY, fontSize: TEXT_PT, color,
			lineSpacingMultiple: 1.5, valign: 'top', align: 'justify',
		}
	);

const bigNumber = (slide, text, x, y, w, h) =>
	slide.addText(text, { x, y, w, h, fontFace: HEAD, fontSize: 24, color: LIME, valign: 'top' });

/* ------------------------------------------------------------ shape helpers */

// picture stand-in: flat colour block where the original deck holds a photograph
function photo(slide, x, y, w, h) {
	slide.addShape('rect', { x, y, w, h, fill: { color: PHOTO }, line: { width: 0 } });
	slide.addText('[image]', {
		x, y, w, h,
		fontFace: BODY, fontSize: TEXT_PT, color: PAPER, align: 'center', valign: 'middle',
	});
}

// outlined card behind a block of copy (adj = corner radius as a fraction of the short side)
function card(slide, x, y, w, h, color, adj) {
	slide.addShape('roundRect', {
		x, y, w, h,
		fill: { type: 'none' },
		line: { color, width: 0.75 },
		rectRadius: +((adj || 0.141) * Math.min(w, h)).toFixed(4),
	});
}

// lime disc with the dark asterisk in the middle
function badge(slide, x, y) {
	slide.addShape('ellipse', { x, y, w: 0.815, h: 0.815, fill: { color: LIME }, line: { width: 0 } });
	glyph(slide, ASTERISK_GLYPH, { x: x + 0.204, y: y + 0.213, w: 0.406, h: 0.406, color: INK });
}

// pill outline holding two arrow glyphs
function arrowButton(slide, x, y, color, rotate) {
	slide.addShape('roundRect', {
		x, y, w: 1.271, h: 0.611,
		fill: { type: 'none' },
		line: { color, width: 0.75 },
		rectRadius: 0.3055,
	});
	[0.267, 0.692].forEach((dx) =>
		glyph(slide, ARROW_GLYPH, { x: x + dx, y: y + 0.153, w: 0.318, h: 0.318, color, rotate })
	);
}

/* --------------------------------------------------------- slide furniture */

// every slide carries a rounded panel on one edge with the vertical branding
function chrome(slide, side, panelColor) {
	const left = side === 'left';
	slide.addShape('round2SameRect', {
		x: left ? -3.168 : 9.001,
		y: 3.168, w: 7.5, h: 1.165,
		rotate: left ? 90 : 270,
		fill: { color: panelColor },
		line: { width: 0 },
		rectRadius: 0.5825,
	});
	slide.addText('COMPANY', {
		x: left ? -0.178 : 12.04, y: 5.971, w: 1.52, h: 0.337, rotate: 270,
		fontFace: HEAD, fontSize: 14, color: INK, valign: 'top',
	});
	[0.601, 0.826, 1.051].forEach((dy) =>
		slide.addShape('ellipse', {
			x: left ? 0.533 : 12.751, y: dy, w: 0.09, h: 0.09,
			fill: { color: INK }, line: { width: 0 },
		})
	);
	slide.addText('09-04-2045', {
		x: left ? 0.101 : 12.319, y: 3.615, w: 0.963, h: 0.269, rotate: 270,
		fontFace: BODY, fontSize: TEXT_PT, color: INK, align: 'center', valign: 'top', wrap: false,
	});
}

/* ----------------------------------------------------------------- slides */

// 1 - cover
function slide01(s) {
	photo(s, 1.795, 2.09, 10.778, 3.705);
	chrome(s, 'left', PAPER);
	titleText(s, 'AGENCY-', 1.789, 0.894, 5.554, 1.111, LIME);
	bodyText(s, LOREM.cover, 7.803, 1.051, 4.756, 0.585, PAPER);
	arrowButton(s, 1.789, 6.168, PAPER, 90);
	titleText(s, 'BUSINESS PLAN', 3.34, 5.935, 9.233, 1.111, LIME);
	badge(s, 11.434, 4.671);
}

// 2 - welcome
function slide02(s) {
	photo(s, 0.76, 0.987, 4.805, 2.763);
	chrome(s, 'right', PAPER);
	titleText(s, 'WEL-COME', 0.76, 4.779, 3.595, 2.121, LIME);
	bodyText(s, LOREM.welcome, 6.609, 2.497, 4.317, 0.837, PAPER);
	labelText(s, 'THIS IS OUR BUSINESS PLAN', 6.609, 1.46, 2.203, 0.671, PAPER);
	card(s, 6.058, 0.993, 5.326, 2.757, PAPER);
	badge(s, 4.444, 2.65);
	bigNumber(s, '+100K', 6.109, 5.934, 1.601, 0.505);
	bigNumber(s, '91%', 8.342, 5.934, 1.122, 0.505);
	bodyText(s, 'Client User Increasement', 6.109, 6.365, 1.868, 0.324, PAPER);
	bodyText(s, 'Successful Plan', 8.342, 6.365, 1.389, 0.324, PAPER);
	arrowButton(s, 10.113, 6.06, PAPER, 180);
}

// 3 - table of contents
function slide03(s) {
	const leftItems = [
		['HISTORY & FEATURES', '01', 2.662], ['VISION', '02', 1.049], ['MISSION', '03', 1.236],
		['IDENTITY', '04', 1.348], ['TARGET GOAL', '05', 1.771], ['SWOT ANALYSIS', '06', 2.086],
		['OUR GOALS', '07', 1.469], ['PLANNING', '08', 1.373],
	];
	const rightItems = [
		['CONTROLLING', '09', 1.838], ['LEADING STRATEGY', '10', 2.484], ['PRIORITIZE', '11', 1.592],
		['PRODUCT', '12', 1.285], ['TEAM', '13', 0.849], ['CONTACTS', '14', 1.469],
	];
	photo(s, 10.426, 0.987, 2.147, 2.646);
	chrome(s, 'left', LIME);
	titleText(s, 'CONTENT', 2.256, 0.752, 5.733, 1.111, INK);
	titleText(s, 'INDEX-', 8.387, 5.803, 4.186, 1.111, INK);
	const entry = (name, num, x, y, w, numX, numW) => {
		s.addText(name, { x, y, w, h: 0.303, fontFace: HEAD, fontSize: LABEL_PT, color: INK, valign: 'top', wrap: false });
		s.addText(num, {
			x: numX, y, w: numW, h: 0.303,
			fontFace: BODY, fontSize: LABEL_PT, color: INK, align: 'right', valign: 'top', wrap: false,
		});
	};
	leftItems.forEach(([name, num, w], i) => entry(name, num, 2.256, 2.267 + i * 0.5314, w, 5.14, 0.4));
	rightItems.forEach(([name, num, w], i) => entry(name, num, 6.235, 2.273 + i * 0.5304, w, 9.081, 0.37));
	arrowButton(s, 6.328, 6.06, INK, 90);
	badge(s, 10.641, 2.57);
}

// 4 - our business
function slide04(s) {
	photo(s, 1.942, 0.987, 3.916, 2.763);
	chrome(s, 'left', PAPER);
	titleText(s, 'OUR- BUSINESS', 6.462, 1.372, 6.111, 2.121, LIME);
	labelText(s, 'HISTORY OF OUR BUSINESS', 2.353, 4.414, 2.047, 0.671, PAPER);
	bodyText(s, LOREM.history, 2.353, 5.45, 3.187, 1.081, PAPER);
	labelText(s, 'OUR COMPANY UNIQUE FEATURES', 6.462, 4.414, 2.334, 0.671, PAPER);
	bodyText(s, LOREM.features, 6.462, 5.45, 3.505, 1.081, PAPER);
	badge(s, 4.725, 2.656);
	arrowButton(s, 11.302, 6.06, PAPER, 180);
	card(s, 1.942, 3.968, 3.916, 2.924, PAPER);
}

// 5 - the vision
function slide05(s) {
	photo(s, 0.76, 2.556, 4.362, 4.344);
	chrome(s, 'right', PAPER);
	titleText(s, 'THE-', 0.76, 0.993, 3.073, 1.111, LIME);
	titleText(s, 'VISION', 6.667, 5.815, 4.462, 1.111, LIME);
	labelText(s, 'HERE\u2019S A FEW OF OUR VISION', 6.664, 2.556, 1.931, 0.671, PAPER);
	bodyText(s, LOREM.features, 6.664, 3.592, 3.505, 1.081, PAPER);
	arrowButton(s, 6.664, 1.249, PAPER, 180);
	badge(s, 3.941, 5.721);
	photo(s, 9.406, 1.232, 1.993, 1.994);
}

// 6 - the mission
function slide06(s) {
	photo(s, 6.763, 2.527, 5.482, 4.138);
	chrome(s, 'left', PAPER);
	titleText(s, 'THE-', 1.968, 5.789, 3.073, 1.111, LIME);
	titleText(s, 'MISSION', 6.667, 1.271, 5.578, 1.111, LIME);
	labelText(s, 'THESE ARE SOME OF OUR MISSION', 1.968, 1.446, 2.317, 0.671, PAPER);
	bodyText(s, LOREM.features, 1.968, 2.483, 3.828, 1.081, PAPER);
	bulletText(s, [
		'Mission One, lorem ipsum dolor sit amet consetectis.',
		'Mission Two, lorem ipsum dolor sit amet consetectis.',
		'Mission Three, lorem ipsum dolor sit amet consetecs.',
	], 1.968, 3.639, 3.613, 0.837, PAPER);
	arrowButton(s, 5.274, 6.054, PAPER, 180);
	badge(s, 11.175, 5.576);
}

// 7 - our identity
function slide07(s) {
	photo(s, 7.42, 0.993, 3.998, 2.757);
	chrome(s, 'right', PAPER);
	titleText(s, 'OUR- IDENTITY', 0.76, 4.779, 5.906, 2.121, LIME);
	labelText(s, 'OUR CONCEPT', 0.76, 1.563, 1.384, 0.671, PAPER);
	bodyText(s, LOREM.concept, 0.76, 2.53, 3.273, 0.829, PAPER);
	labelText(s, 'OUR AUDIENCES', 7.883, 4.712, 1.663, 0.671, PAPER);
	bulletText(s, [
		'Audiences One, lorem ipsum dolor sit ametsi.',
		'Audiences Two, lorem ipsum dolor sit ametsi.',
		'Audiences Three, lorem ipsum dolor sit ametsi.',
	], 7.883, 5.678, 3.18, 0.829, PAPER);
	badge(s, 10.395, 2.732);
	arrowButton(s, 4.033, 5.067, PAPER, 180);
	card(s, 7.47, 4.231, 3.945, 2.626, PAPER);
	photo(s, 4.849, 0.993, 2.325, 2.757);
}

// 8 - target goals
function slide08(s) {
	photo(s, 6.667, 1.294, 4.751, 2.592);
	chrome(s, 'right', LIME);
	titleText(s, 'TARGET GOALS-', 0.76, 1.629, 4.892, 2.121, INK);
	s.addText('$70.000,00', {
		x: 3.794, y: 5.57, w: 2.317, h: 0.502,
		fontFace: HEAD, fontSize: 18, color: INK, lineSpacingMultiple: 1.5, valign: 'top',
	});
	s.addText('$140.000,00', {
		x: 6.703, y: 5.57, w: 2.317, h: 0.502,
		fontFace: HEAD, fontSize: 18, color: INK, lineSpacingMultiple: 1.5, valign: 'top',
	});
	labelText(s, 'THIS MONTH\u2019S TARGET:', 3.794, 4.899, 1.838, 0.671, INK);
	labelText(s, 'FUTURE TARGET GOALS:', 6.703, 4.899, 2.054, 0.671, INK);
	[3.815, 6.703].forEach((x, i) =>
		s.addText('The target we would like', {
			x, y: 6.072, w: i ? 1.856 : 1.837, h: 0.269,
			fontFace: BODY, fontSize: TEXT_PT, color: INK, valign: 'top',
		})
	);
	badge(s, 10.416, 2.858);
	arrowButton(s, 10.119, 6.009, INK, 180);
	card(s, 3.433, 4.625, 2.808, 1.994, INK);
	photo(s, 0.76, 4.625, 2.325, 1.994);
}

// 9 - swot analysis
function slide09(s) {
	photo(s, 6.957, 0.985, 4.461, 3.568);
	chrome(s, 'right', PAPER);
	titleText(s, 'SWOT-', 0.76, 1.185, 3.984, 1.111, LIME);
	titleText(s, 'ANALY-SIS', 7.026, 4.769, 4.174, 2.121, LIME);
	// heading + copy for each of the four quadrants: name, x, heading y, heading w/h, copy y
	const quadrants = [
		['STRENGTH', 0.777, 2.865, 1.668, 0.379, 3.472],
		['WEAKNESS', 3.885, 2.865, 1.668, 0.379, 3.472],
		['OPPORTUNITY', 0.76, 4.974, 1.895, 0.368, 5.581],
		['THREAT', 3.869, 4.991, 1.668, 0.379, 5.599],
	];
	quadrants.forEach(([name, x, labelY, w, h, copyY]) => {
		labelText(s, name, x, labelY, w, h, PAPER);
		bodyText(s, LOREM.swot, x, copyY, 2.575, 1.081, PAPER);
	});
	arrowButton(s, 5.189, 1.382, PAPER, 180);
	badge(s, 10.416, 3.538);
}

// 10 - company goals
function slide10(s) {
	photo(s, 1.942, 1.392, 3.38, 2.486);
	chrome(s, 'left', PAPER);
	titleText(s, 'COMPANY GOALS-', 6.046, 4.779, 5.906, 2.121, LIME);
	labelText(s, 'STRATEGIC DIRECTION', 6.058, 1.679, 1.89, 0.671, PAPER);
	bodyText(s, LOREM.short, 6.058, 2.675, 2.859, 0.829, PAPER);
	labelText(s, 'FLEXIBILITY & ADAPTIBILITY', 1.957, 4.62, 2.303, 0.671, PAPER);
	bodyText(s, LOREM.flexibility, 1.957, 5.616, 3.366, 1.081, PAPER);
	labelText(s, 'QUANTIFIABLE TARGET', 9.898, 1.679, 1.996, 0.671, PAPER);
	bodyText(s, LOREM.quantifiable, 9.898, 2.675, 2.411, 0.829, PAPER);
	arrowButton(s, 11.299, 6.064, PAPER, 180);
	badge(s, 4.381, 2.885);
	card(s, 9.544, 1.392, 3.026, 2.358, PAPER);
}

// 11 - the planning
function slide11(s) {
	photo(s, 9.039, 0.985, 3.534, 4.344);
	chrome(s, 'left', LIME);
	titleText(s, 'PLANNING', 5.483, 5.789, 6.111, 1.111, INK);
	titleText(s, 'THE-', 1.972, 1.181, 3.084, 1.111, INK);
	labelText(s, 'CLARIFYING BUSINESS GOALS', 5.505, 1.287, 2.218, 0.671, INK);
	bodyText(s, LOREM.short, 5.505, 2.104, 2.859, 0.829, INK);
	labelText(s, 'PLANNING DESCRIPTION', 1.972, 3.56, 1.807, 0.671, INK);
	bodyText(s, LOREM.short, 1.957, 4.501, 2.859, 0.829, INK);
	labelText(s, 'OUTLINING STRATEGIES & TACTICS PLAN', 5.505, 3.56, 3.084, 0.671, INK);
	bodyText(s, LOREM.short, 5.505, 4.501, 2.859, 0.829, INK);
	arrowButton(s, 1.957, 6.009, INK, 90);
	badge(s, 11.589, 4.288);
	card(s, 5.154, 0.993, 3.534, 2.196, INK);
}

// 12 - our control
function slide12(s) {
	photo(s, 8.116, 0.983, 4.457, 2.657);
	chrome(s, 'left', PAPER);
	titleText(s, 'OUR- CONTROL', 1.944, 1.393, 5.573, 2.121, LIME);
	labelText(s, 'THIS IS HOW WE CONTROL THE SYSTEM', 6.682, 4.851, 2.844, 0.671, PAPER);
	bodyText(s, LOREM.control, 6.667, 5.791, 3.789, 0.829, PAPER);
	arrowButton(s, 11.302, 6.046, PAPER, 180);
	badge(s, 11.589, 2.65);
	photo(s, 1.956, 4.556, 3.944, 2.333);
}

// 13 - leading strategy
function slide13(s) {
	photo(s, 8.822, 0.983, 2.563, 2.735);
	chrome(s, 'right', PAPER);
	titleText(s, 'LEAD-', 0.76, 1.14, 3.606, 1.111, LIME);
	titleText(s, 'STRATEGY', 5.05, 5.802, 6.335, 1.111, LIME);
	labelText(s, 'MARKET POSITIONING', 0.76, 2.859, 2.0, 0.671, PAPER);
	bodyText(s, LOREM.short, 0.76, 3.677, 2.863, 0.829, PAPER);
	labelText(s, 'DEFINITION & CLARITY', 0.76, 5.125, 2.0, 0.671, PAPER);
	bodyText(s, LOREM.short, 0.76, 5.943, 2.863, 0.829, PAPER);
	arrowButton(s, 10.115, 4.71, PAPER, 180);
	badge(s, 10.381, 2.65);
	photo(s, 5.05, 0.983, 3.568, 4.337);
}

// 14 - the priorities
function slide14(s) {
	photo(s, 7.022, 2.356, 4.367, 2.703);
	chrome(s, 'right', LIME);
	titleText(s, 'PRIORITIES', 4.172, 0.993, 7.217, 1.111, INK);
	titleText(s, 'THE-', 0.76, 5.789, 3.084, 1.111, INK);
	labelText(s, 'COMPANY PRIORITIES', 4.307, 5.048, 1.715, 0.671, INK);
	bodyText(s, LOREM.priorities, 4.307, 5.866, 4.137, 0.829, INK);
	arrowButton(s, 10.118, 6.009, INK, 180);
	badge(s, 7.256, 4.019);
	photo(s, 0.76, 0.983, 2.806, 4.075);
}

// 15 - charts
function slide15(s) {
	chrome(s, 'left', PAPER);
	titleText(s, 'CHAR-TS', 9.101, 1.831, 3.472, 2.121, LIME);
	labelText(s, 'THE CHART WE USE FOR', 9.101, 4.528, 1.715, 0.671, PAPER);
	bodyText(s, LOREM.chart, 9.101, 5.379, 3.472, 0.829, PAPER);
	const labels = ['Team 1', 'Team 2', 'Team 3', 'Team 4'];
	s.addChart(
		'line',
		[
			{ name: 'Period 1', labels, values: [4.3, 2.5, 3.5, 4.5] },
			{ name: 'Period 2', labels, values: [2.4, 4.4, 1.8, 2.8] },
			{ name: 'Period 3', labels, values: [2, 2, 3, 5] },
		],
		{
			x: 2.222, y: 1.831, w: 5.822, h: 4.308,
			chartColors: [LIME, '8B8B8B', '8ABE03'],
			lineSize: 2, lineDataSymbol: 'none', lineCap: 'round',
			showLegend: true, legendPos: 'b', legendFontFace: BODY, legendFontSize: TEXT_PT, legendColor: PAPER,
			catAxisLabelFontFace: BODY, catAxisLabelFontSize: LABEL_PT, catAxisLabelColor: PAPER,
			catAxisLineColor: PAPER, catAxisLineSize: 0.75, catAxisMajorTickMark: 'none',
			valAxisLabelFontFace: BODY, valAxisLabelFontSize: LABEL_PT, valAxisLabelColor: PAPER,
			valAxisLineShow: false, valAxisMaxVal: 6, valAxisMajorUnit: 1,
			valGridLine: { color: PAPER, size: 0.75 }, catGridLine: { style: 'none' },
		}
	);
	card(s, 1.749, 1.292, 6.768, 5.208, PAPER, 0.109);
	arrowButton(s, 11.302, 4.561, PAPER, 180);
	badge(s, 7.873, 5.857);
}

// 16 - growth strategy
function slide16(s) {
	photo(s, 1.957, 4.476, 3.694, 2.423);
	chrome(s, 'left', PAPER);
	titleText(s, 'GROWTH- STRATEGY', 6.233, 4.779, 6.34, 2.121, LIME);
	[
		['MARKET DEVELOPMENT', 2.415, 1.927],
		['DIVERSIFICATION STRATEGY', 6.233, 2.448],
		['PRODUCT DEVELOPMENT', 9.666, 1.922],
	].forEach(([name, x, w]) => {
		labelText(s, name, x, 1.873, w, 0.682, PAPER);
		bodyText(s, LOREM.growth, x, 2.836, 2.85, 0.837, PAPER);
	});
	arrowButton(s, 11.302, 0.958, PAPER, 180);
	card(s, 1.957, 1.387, 3.694, 2.668, PAPER, 0.109);
	badge(s, 4.655, 5.901);
}

// 17 - marketing plan
function slide17(s) {
	photo(s, 0.76, 3.569, 5.117, 3.33);
	chrome(s, 'right', PAPER);
	titleText(s, 'MARKE-TING', 0.76, 0.993, 4.44, 2.121, LIME);
	titleText(s, '& PLAN', 6.667, 5.934, 4.195, 1.111, LIME);
	labelText(s, 'INBOUND MARKETING', 6.967, 1.314, 2.064, 0.671, PAPER);
	bulletText(s, [
		'Content Marketing With High Quality.',
		'Search Enging Organic (SEO).',
		'Social Media Marketing.',
	], 6.967, 2.118, 2.686, 0.829, PAPER);
	labelText(s, 'OUTBOUND MARKETING', 6.967, 3.783, 2.214, 0.671, PAPER);
	bulletText(s, [
		'Email Marketing Campaign To Clients.',
		'Paid Advertising In Social Media or Ads.',
		'Direct Outreach Via Cold Emails or Calls.',
	], 6.967, 4.587, 2.899, 0.834, PAPER);
	arrowButton(s, 10.118, 1.314, PAPER, 180);
	badge(s, 4.838, 5.901);
	card(s, 6.658, 3.569, 3.403, 2.087, PAPER, 0.173);
}

// 18 - the product
function slide18(s) {
	photo(s, 1.933, 2.068, 4.556, 3.705);
	chrome(s, 'left', PAPER);
	titleText(s, 'THE-', 1.948, 5.918, 3.084, 1.111, LIME);
	titleText(s, 'PRODUCT', 6.667, 0.87, 5.695, 1.111, LIME);
	bigNumber(s, '+100K', 6.667, 6.125, 1.601, 0.505);
	bigNumber(s, '91%', 8.899, 6.125, 1.122, 0.505);
	bodyText(s, 'Client User Increasement', 6.667, 6.556, 1.868, 0.324, PAPER);
	bodyText(s, 'Successful Plan', 8.899, 6.556, 1.389, 0.324, PAPER);
	arrowButton(s, 11.302, 6.192, PAPER, 180);
	labelText(s, '\u201CINNOVATION THROUGH IMAGINATION.\u201D', 1.933, 0.955, 2.9, 0.671, PAPER);
	badge(s, 5.343, 4.646);
	photo(s, 6.667, 2.068, 5.906, 3.705);
}

// 19 - the teams
function slide19(s) {
	const team = [
		['RYAN DAVIS', 'CEO of Company', 1.967, 1.953],
		['EMILY PARKER', 'Account Manager', 4.635, 4.635],
		['ALEX JOHNSON', 'Main Designers', 7.338, 7.337],
		['PAUL WILLIAM', 'Creative Director', 10.04, 10.04],
	];
	chrome(s, 'left', LIME);
	titleText(s, 'THE-', 1.967, 0.87, 2.944, 1.111, INK);
	titleText(s, 'TEAMS', 8.384, 5.918, 4.189, 1.111, INK);
	team.forEach(([name, role, nx, rx]) => {
		labelText(s, name, nx, 5.15, 1.887, 0.368, INK);
		bodyText(s, role, rx, 5.474, 1.868, 0.324, INK);
	});
	arrowButton(s, 1.953, 6.192, INK, 90);
	labelText(s, '\u201CCREATIVY DRIVES INNOVATION.\u201D', 10.04, 0.955, 2.533, 0.671, INK, 'right');
	[1.933, 4.635, 7.338, 10.04].forEach((x) => photo(s, x, 2.068, 2.533, 2.965));
	badge(s, 2.148, 4.009);
}

// 20 - thank you
function slide20(s) {
	photo(s, 7.456, 0.955, 3.918, 5.944);
	chrome(s, 'right', PAPER);
	titleText(s, 'THANK- YOU VERY MUCH', 0.756, 0.955, 5.906, 3.13, LIME);
	bodyText(s, '@social_media', 1.16, 6.428, 1.612, 0.324, PAPER, 'left');
	bodyText(s, 'www.example.com', 2.905, 6.428, 1.612, 0.324, PAPER, 'center');
	bodyText(s, '+11 222 3333 4444', 4.65, 6.428, 1.612, 0.324, PAPER, 'right');
	s.addShape('roundRect', {
		x: 0.756, y: 6.288, w: 5.906, h: 0.611,
		fill: { type: 'none' }, line: { color: PAPER, width: 0.75 }, rectRadius: 0.3055,
	});
	arrowButton(s, 4.306, 3.172, PAPER, 180);
	bodyText(s, LOREM.thanks, 0.772, 4.794, 5.601, 0.585, PAPER);
	badge(s, 10.224, 5.809);
}

/* -------------------------------------------------------------------- build */

const DECK = [
	[slide01, INK], [slide02, INK], [slide03, PAPER], [slide04, INK], [slide05, INK],
	[slide06, INK], [slide07, INK], [slide08, PAPER], [slide09, INK], [slide10, INK],
	[slide11, PAPER], [slide12, INK], [slide13, INK], [slide14, PAPER], [slide15, INK],
	[slide16, INK], [slide17, INK], [slide18, INK], [slide19, PAPER], [slide20, INK],
];

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Business Plan';

DECK.forEach(([build, background]) => {
	const slide = pptx.addSlide();
	slide.background = { color: background };
	build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0c0045a8-9f68-443c-a57b-3fa515d71101_grok_final.pptx') });
