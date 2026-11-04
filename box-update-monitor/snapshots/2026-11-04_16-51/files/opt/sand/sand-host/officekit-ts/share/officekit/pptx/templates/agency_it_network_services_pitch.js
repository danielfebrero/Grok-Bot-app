/**
 * The Agency Business - 40-slide business deck, rebuilt with pptxgenjs.
 * Photographic content in the original is replaced by flat placeholder blocks.
 *
 *   node 02cb9c24-748b-47f3-8414-913b8756dbb3_grok_final.js
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const DARK      = '3F3F3F';
const GOLD      = 'FFC000';
const GRAY      = '7F7F7F';
const INK       = '262626';
const LTGRAY    = 'BFBFBF';
const MIDGRAY   = '595959';
const NAVY      = '203864';
const OFFWHITE  = 'F2F2F2';
const PALE      = 'D8D8D8';
const PH1       = 'C3C3C3';
const PH2       = 'DBDBDB';
const WHITE     = 'FFFFFF';
const PHOTO     = 'A6A6A6';   // stand-in for photographic imagery

// Body/heading faces used throughout the deck
const LATO = 'Lato', POP = 'Poppins Medium', POPSB = 'Poppins SemiBold';
const POPL = 'Poppins Light', MONTSB = 'Montserrat SemiBold';

// Default text-frame insets of the source deck, in points
const INSET = [7.2, 7.2, 3.6, 3.6];   // left, right, bottom, top
const INSET0 = [0, 0, 3.6, 3.6];

// ---------------------------------------------------------------- helpers

/** Plain shape: sh(slide, x, y, w, h, {shape, f:fill, l:line, rot, fh, fv, rr}) */
function sh (s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.shape || 'rect', geom(x, y, w, h, o));
}

/** Straight connector */
function ln (s, x, y, w, h, o) {
  s.addShape('line', geom(x, y, w, h, o));
}

/** Image placeholder: a flat block standing in for a photo */
function img (s, x, y, w, h, o) {
  const g = geom(x, y, w, h, { f: { color: o.f }, l: o.l, rot: o.rot, shape: o.shape });
  if (!o.label) { s.addShape(o.shape || 'rect', g); return; }
  s.addText([{ text: o.label, options: { fontFace: LATO, fontSize: 10, color: 'FFFFFF' } }],
    Object.assign(g, { shape: o.shape || 'rect', align: 'center', valign: 'middle' }));
}

/** Freeform built from a normalised (0..1000) point list */
function poly (s, subpaths, x, y, w, h, o) {
  o = o || {};
  const pts = [];
  subpaths.forEach(sub => {
    for (let i = 0; i < sub.length; i += 2) {
      pts.push({ x: (sub[i] / 1000) * w, y: (sub[i + 1] / 1000) * h, moveTo: i === 0 });
    }
    pts.push({ close: true });
  });
  s.addShape('custGeom', Object.assign(geom(x, y, w, h, o), { points: pts }));
}

/** Text box (optionally drawn inside a shape) */
function tx (s, x, y, w, h, runs, o) {
  o = o || {};
  const body = runs.map(([t, r]) => ({
    text: t,
    options: {
      fontFace: r.face, fontSize: r.sz, color: r.c,
      bold: !!r.b, italic: !!r.i, underline: r.u ? { style: 'sng' } : undefined,
      superscript: !!r.sup, breakLine: !!r.br,
    },
  }));
  const opt = geom(x, y, w, h, o);
  opt.shape = o.shape || 'rect';
  opt.align = o.a || 'left';
  opt.valign = o.v || 'top';
  opt.margin = o.m0 ? INSET0 : INSET;
  opt.wrap = true;
  if (o.ls) opt.lineSpacingMultiple = o.ls;
  if (o.vert) opt.vert = o.vert;
  s.addText(body, opt);
}

/** Shared position / fill / line / rotation plumbing */
function geom (x, y, w, h, o) {
  const g = { x: x, y: y, w: w, h: h };
  g.fill = o.f ? (typeof o.f === 'string' ? { color: o.f } : o.f) : { type: 'none' };
  g.line = o.l || { type: 'none' };
  if (o.rot) g.rotate = o.rot;
  if (o.fh) g.flipH = true;
  if (o.fv) g.flipV = true;
  if (o.rr !== undefined) g.rectRadius = o.rr;
  return g;
}

/** Horizontal alpha ramp, approximated with a stack of translucent bands */
function fade (s, x, y, w, h, color, t0, t1, bands) {
  bands = bands || 12;
  for (let i = 0; i < bands; i++) {
    sh(s, x + (i / bands) * w, y, w / bands + 0.01, h,
      { f: { color: color, transparency: t0 + ((t1 - t0) * (i + 0.5)) / bands } });
  }
}

// ---------------------------------------------------------------- charts
// The source deck ships these as flat pictures; here they are real charts.

const CHART_CATS = ['Analysis 1', 'Analysis 2', 'Analysis 3', 'Analysis 4'];
const CHART_S1 = [4.3, 2.5, 3.5, 4.5];
const CHART_S2 = [2.4, 4.4, 1.8, 2.8];
const CHART_AXIS = { catAxisLabelFontFace: LATO, catAxisLabelFontSize: 9, catAxisLabelColor: GRAY,
  valAxisLabelFontFace: LATO, valAxisLabelFontSize: 9, valAxisLabelColor: GRAY,
  valAxisMaxVal: 5, valAxisMinVal: 0, showLegend: false, chartColors: [OFFWHITE, GOLD],
  barGapWidthPct: 40, dataBorder: { pt: 0, color: 'FFFFFF' },
  chartArea: { fill: { color: 'FFFFFF', transparency: 100 }, border: { pt: 0, color: 'FFFFFF' } } };

/** Clustered column chart (light/gold pair per category) */
function columnChart (s, x, y, w, h) {
  s.addChart('bar', [
    { name: 'Series 1', labels: CHART_CATS, values: CHART_S1 },
    { name: 'Series 2', labels: CHART_CATS, values: CHART_S2 },
  ], Object.assign({ x: x, y: y, w: w, h: h, barDir: 'col',
    valGridLine: { style: 'solid', color: 'BFBFBF', size: 0.5 },
    catGridLine: { style: 'none' } }, CHART_AXIS));
}

/** Horizontal bar chart (navy/gold pair per category) */
function barChart (s, x, y, w, h) {
  s.addChart('bar', [
    { name: 'Series 1', labels: CHART_CATS, values: CHART_S1 },
    { name: 'Series 2', labels: CHART_CATS, values: CHART_S2 },
  ], Object.assign({}, CHART_AXIS, { x: x, y: y, w: w, h: h, barDir: 'bar',
    chartColors: [NAVY, GOLD], catAxisLabelColor: NAVY,
    valGridLine: { style: 'solid', color: 'BFBFBF', size: 0.5 },
    catGridLine: { style: 'none' } }));
}

/** Four-slice pie with a bottom legend */
function pieChart (s, x, y, w, h) {
  s.addChart('pie', [{ name: 'Quarters', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [58.5, 22.7, 10.1, 8.7] }], {
    x: x, y: y, w: w, h: h,
    chartColors: [NAVY, 'D8D8D8', '2E508E', GOLD],
    showLegend: true, legendPos: 'b', legendFontFace: LATO, legendFontSize: 9, legendColor: DARK,
    showValue: false, dataBorder: { pt: 0, color: 'FFFFFF' },
    chartArea: { fill: { color: 'FFFFFF', transparency: 100 }, border: { pt: 0, color: 'FFFFFF' } },
  });
}

// ---------------------------------------------------------------- device mock-ups
// Drawn from primitives rather than embedded as photographs. Proportions are
// taken from the artwork they replace.

/** Laptop: dark screen in a body sitting on a wide silver base */
function laptop (s, x, y, w, h) {
  sh(s, x + 0.093 * w, y, 0.814 * w, 0.935 * h, { f: '141414' });
  sh(s, x + 0.123 * w, y + 0.056 * h, 0.754 * w, 0.82 * h, { f: PH1 });
  sh(s, x, y + 0.935 * h, w, 0.03 * h, { f: 'E2E3E5' });
  sh(s, x + 0.02 * w, y + 0.965 * h, 0.96 * w, 0.025 * h, { f: 'B9BABE' });
  sh(s, x + 0.44 * w, y + 0.965 * h, 0.12 * w, 0.012 * h, { f: '8E9095' });
}

/** Phone: rounded dark body with a tall screen and a speaker notch */
function phone (s, x, y, w, h) {
  sh(s, x, y, w, h, { f: '101010', rr: 0.1 * w });
  sh(s, x + 0.067 * w, y + 0.066 * h, 0.864 * w, 0.894 * h, { f: PH1 });
  sh(s, x + 0.4 * w, y + 0.03 * h, 0.2 * w, 0.014 * h, { f: '3A3A3A', rr: 0.007 * h });
}

/** Desktop monitor: screen block over a trapezoid pedestal */
function monitor (s, x, y, w, h) {
  sh(s, x + 0.018 * w, y, 0.962 * w, 0.72 * h, { f: '232323' });
  sh(s, x + 0.056 * w, y + 0.048 * h, 0.886 * w, 0.6 * h, { f: PH1 });
  sh(s, x + 0.018 * w, y + 0.72 * h, 0.962 * w, 0.085 * h, { f: 'C7C8CA' });
  sh(s, x + 0.4 * w, y + 0.805 * h, 0.2 * w, 0.13 * h, { f: 'DCDDDF' });
  sh(s, x + 0.34 * w, y + 0.935 * h, 0.32 * w, 0.03 * h, { f: 'C7C8CA', rr: 0.015 * h });
}

// ---------------------------------------------------------------- freeform paths
// Outlines traced from the source artwork, normalised to a 0..1000 box.
// Each entry is a list of sub-paths; every sub-path is a flat x,y,x,y,... list.
const g1 = [[365,538,191,327,0,558,365,1000,1000,231,809,0]];
const g2 = [[750,0,250,0,69,138,8,641,69,862,179,984,750,1000,931,862,992,359,931,138,821,16]];
const g3 = [[749,994,824,949,774,638,998,378,671,297,533,12,467,12,329,297,0,388,226,638,190,983,510,852],[476,750,275,857,313,608,143,429,380,390,495,169,607,392,847,429,679,608,715,857,480,748]];
const g4 = [[983,970,775,970,773,822,678,815,670,725,564,723,565,550,685,488,932,225,951,92,843,85,822,0,172,0,156,85,45,92,65,228,126,318,312,489,432,551,433,724,327,726,318,817,226,821,221,970,1,989,983,1000,987,970],[846,115,922,121,890,241,803,333,845,155],[78,121,154,115,197,332,110,240,82,163],[251,365,196,206,192,30,808,30,805,206,769,329,668,464,535,537,531,723,471,723,461,531,275,402],[355,755,644,755,644,817,355,817],[742,970,260,970,260,845,742,845]];
const g5 = [[983,314,935,280,952,194,868,166,847,94,742,84,685,24,589,38,510,0,416,43,316,24,250,88,142,104,126,174,36,213,56,277,1,325,65,393,40,469,133,507,142,569,189,589,44,952,216,858,260,991,295,995,443,656,500,671,551,648,699,995,736,991,780,858,948,954,805,580,859,568,875,498,963,458,942,394,991,362,993,322],[286,928,238,819,112,895,241,594,322,654,411,641],[890,893,764,817,717,928,591,632,698,648,760,586],[951,350,895,390,925,450,833,485,831,549,727,557,684,617,578,597,494,644,423,604,336,624,288,564,189,560,178,491,86,457,120,388,47,332,108,286,80,224,169,191,172,127,276,117,326,54,420,77,505,30,578,71,667,51,714,112,814,116,825,185,914,215,882,288,954,348]];
const g6 = [[500,0,146,147,0,498,146,852,500,1000,753,932,982,634,932,250,633,18],[500,931,196,804,69,502,129,285,386,85,717,128,915,385,871,715,614,915]];
const g7 = [[539,27,462,12,324,300,5,367,222,639,175,967,236,998,500,855,762,998,820,976,778,639,998,377,676,300],[686,588,705,855,490,747,275,855,308,609,139,428,375,389,485,171,592,390,832,428]];
const g8 = [[500,1000,855,854,1000,500,852,146,500,0,148,146,0,500,69,752,368,982],[500,149,748,252,848,500,745,748,500,851,255,748,152,500,200,323,408,162]];
const g9 = [[740,0,256,0,126,69,9,369,32,997,979,989,991,366,924,145,809,18],[90,825,113,338,256,175,823,221,900,415,906,825]];
const g10 = [[798,560,852,368,853,188,428,0,4,190,31,478,108,655,209,799,428,1000,576,882,658,967,734,996,835,992,934,934,992,837,993,724,942,631,853,572],[429,965,208,750,121,616,56,456,30,210,429,31,827,210,784,512,765,560,688,575,746,424,759,247,428,101,95,254,157,544,292,751,426,880,548,771,561,850,441,956],[559,717,430,847,373,795,181,526,124,271,430,133,736,271,693,495,567,695],[774,968,636,913,579,779,636,645,774,588,912,645,970,779,943,874,826,961]];
const g11 = [[926,426,574,426,571,53,500,0,429,53,426,426,20,447,53,571,426,574,429,947,479,998,571,947,574,574,965,564,997,479,945,429]];
const g12 = [[470,1,352,13,154,86,28,207,2,279,29,412,96,490,212,580,273,673,282,705,186,718,186,867,356,880,356,904,281,906,268,927,356,940,364,980,420,1000,615,994,645,971,647,940,734,927,721,906,647,904,647,880,814,871,817,718,720,705,729,672,790,578,936,459,996,348,979,223,923,148,840,84,733,36,608,7],[576,962,410,961,409,878,579,878,578,962],[753,742,753,843,241,843,241,742,686,742],[470,705,476,423,514,427,514,705],[812,498,704,595,656,705,576,705,574,416,529,388,447,392,414,427,414,705,341,705,294,595,110,435,68,362,61,283,84,219,194,113,370,49,473,37,595,43,799,111,872,167,919,234,935,308,921,378,848,471]];
const g13 = [[500,0,137,55,0,200,16,856,359,994,862,945,1000,800,984,144,641,6]];
const g14 = [[514,103,342,4,129,37,5,175,46,346,592,978,801,989,931,925,995,825,954,654]];
const g15 = [[877,37,662,4,490,103,14,711,70,925,273,1000,469,944,990,289,927,75]];
const g16 = [[985,942,959,942,940,718,894,517,824,346,735,212,630,123,514,87,513,21,498,0,484,21,483,87,367,123,263,212,175,346,105,517,60,718,41,942,7,946,4,992,95,1000,109,986,103,946,71,942,89,733,132,546,196,387,279,262,375,179,483,145,485,201,503,214,514,142,622,176,720,259,803,384,868,543,911,731,929,939,897,943,891,983,905,997,989,996,1000,968,989,943]];
const g17 = [[773,0,664,28,583,102,2,109,0,841,324,852,305,915,204,919,211,1000,716,1000,723,919,623,915,603,852,925,844,927,397,980,322,1000,229,969,114,888,31,833,8],[697,945,697,970,230,970,230,945,630,945],[336,915,356,852,570,852,589,915],[897,821,30,821,30,680,897,680],[897,650,30,650,30,132,567,132,546,212,554,289,612,390,712,449,773,458,897,420],[773,427,673,400,583,281,576,229,603,129,720,37,773,30,872,58,963,176,970,229,943,329,825,420]];
const g18 = [[500,1000,856,850,1000,500,856,150,500,0,244,70,17,370,67,748,365,981],[500,333,615,385,667,500,547,661,412,642,339,541,358,412,459,339]];
const g19 = [[908,756,834,624,1000,572,990,404,823,347,910,202,780,84,624,166,596,10,428,0,376,166,218,84,86,211,177,347,1,419,10,596,177,653,90,798,220,916,376,834,404,990,572,1000,624,834,790,914,912,764],[775,830,644,746,553,794,539,930,461,930,431,778,333,745,225,830,170,779,254,648,206,557,70,542,70,465,222,435,255,337,170,229,225,173,334,258,431,226,461,74,539,74,553,210,646,258,775,173,830,229,745,334,778,431,930,461,930,539,794,553,746,646,830,775]];
const g20 = [[500,0,250,67,18,365,67,754,365,982,754,933,982,635,931,246,631,18],[500,825,338,781,187,588,219,338,412,187,666,219,813,412,776,662,584,813]];
const g21 = [[750,0,250,0,69,138,8,359,31,763,119,937,250,1000,750,1000,931,862,992,359,931,138,821,16]];
const g22 = [[750,0,250,0,69,138,8,641,69,862,179,984,750,1000,931,862,990,359,922,138,812,16]];
const g23 = [[250,1000,750,1000,931,862,992,359,931,138,821,16,250,0,69,138,8,624,69,844,179,980]];
const g24 = [[900,502,577,414,577,192,846,261,969,153,454,0,454,59,126,156,46,305,146,463,454,547,454,778,138,675,0,778,462,911,462,1000,577,1000,577,911,800,875,971,760,997,632,931,523],[454,399,277,291,454,192],[715,749,569,783,569,567,754,629,733,739]];
const g25 = [[984,116,915,144,964,18,843,73,614,13,502,114,460,300,244,222,55,48,25,217,88,367,24,347,37,455,148,595,93,603,131,692,264,785,6,901,306,1000,628,890,829,621,984,116]];
const g26 = [[18,479,248,578,348,924,540,793,813,998,990,7,17,421,8,472],[323,526,785,213,405,626,369,789,315,533]];
const g27 = [[548,603,793,603,625,784,418,799,284,723,190,512,228,350,416,201,739,278,850,144,574,5,256,65,19,364,67,750,367,982,629,983,923,767,1000,413,540,415,548,603]];
const g28 = [[674,326,655,19,390,0,328,47,326,326,2,373,9,642,326,674,328,953,390,1000,655,981,674,674,981,655,998,373,936,326]];
const g29 = [[20,0,984,0,996,5,1000,14,1000,1000,0,1000,0,25,3,12,15,1]];
const g30 = [[0,0,202,0,1000,1000,0,997]];
const g31 = [[20,0,984,0,996,5,1000,12,1000,929,986,941,198,941,134,1000,70,941,20,941,6,934,0,919,0,22,3,11,15,1]];
const g32 = [[0,0,1000,0,805,115,736,275,710,500,736,725,805,885,1000,1000,0,1000,195,885,283,619,264,275,195,115,97,54]];
const g33 = [[10,1000,0,990,0,1000],[948,0,1,0,53,1000,1000,1000]];
const g34 = [[500,0,633,18,752,68,854,146,932,248,982,367,1000,500,982,633,932,752,854,854,752,932,633,982,500,1000,367,982,248,932,146,854,68,752,18,633,0,500,18,367,68,248,146,146,248,68,367,18]];
const g35 = [[0,0,1000,0,1000,1000,571,1000,571,658,0,658]];
const g36 = [[0,0,1000,0,1000,970,994,999,8,1000,0,970]];
const g37 = [[0,0,1000,0,741,0,1000,223,1000,1000,579,1000,579,891,0,342]];
const g38 = [[0,0,1000,832,1000,1000,0,1000]];
const g39 = [[12,0,811,0,861,18,906,68,945,146,974,248,993,367,1000,500,993,633,974,752,945,854,906,932,861,982,811,1000,0,998,0,2]];
const g40 = [[0,0,1000,0,1000,1000,0,1000,147,1000,147,147,143,113,129,82,108,54,81,32,48,15,0,3]];
const g41 = [[0,0,1000,0,1000,969,993,996,13,1000,4,991,0,969]];
const g42 = [[999,0,1000,976,993,997,13,1000,4,993,0,976,0,211,999,211]];
const g43 = [[0,0,1000,0,1000,969,993,996,13,1000,4,991,0,969]];
const g44 = [[0,0,1000,0,1000,969,990,999,13,1000,4,991,0,969]];
const g45 = [[0,500,0,1000,1000,1000,1000,500],[0,500,1000,500,489,0]];
const g46 = [[0,0,1000,0,537,1000,0,1000]];
const g47 = [[572,0,1000,0,1000,1000,0,1000]];
const g48 = [[0,0,985,0,996,7,1000,24,1000,976,993,997,985,1000,0,1000]];
const g49 = [[0,0,1000,0,742,0,1000,224,1000,1000,777,1000,0,264]];
const g50 = [[0,0,1000,0,1000,1000,840,1000,840,237,0,237]];
const g51 = [[534,0,647,5,758,21,862,47,952,83,840,317,786,295,681,264,630,254,531,246,487,248,423,261,403,273,391,288,387,306,391,324,406,338,445,354,722,410,813,437,866,462,936,514,971,559,993,613,1000,678,993,736,985,763,959,816,920,864,868,906,766,957,683,981,589,995,483,1000,347,993,215,973,96,942,0,902,120,667,204,702,266,722,363,743,425,751,511,754,570,748,608,736,627,718,631,702,626,684,611,669,587,657,260,583,202,565,150,539,103,506,61,466,46,442,24,387,17,322,24,264,32,236,58,184,76,158,122,113,149,93,214,57,291,29,380,11,480,1]];
const g52 = [[0,0,230,0,291,630,384,0,616,0,709,630,770,0,1000,0,844,1000,579,1000,500,432,421,1000,156,1000]];
const g53 = [[500,261,464,264,431,274,385,301,359,326,337,356,320,392,308,432,302,476,302,524,308,568,320,608,348,660,371,687,415,719,447,732,482,738,518,738,553,732,585,719,629,687,663,644,680,608,692,568,698,524,698,476,692,432,680,392,652,340,629,313,585,281,553,268,518,262],[500,0,592,7,678,29,757,64,827,113,887,172,935,242,955,281,984,364,998,453,998,547,984,636,955,719,935,758,887,828,827,887,757,936,718,955,636,984,547,998,453,998,364,984,282,955,243,936,173,887,113,828,65,758,45,719,16,636,2,547,2,453,16,364,45,281,65,242,113,172,173,113,243,64,282,45,364,16,453,2]];
const g54 = [[0,0,1000,0,1000,261,683,261,683,1000,317,1000,317,261,0,261]];
const g55 = [[408,1000,594,1000,702,964,799,863,881,706,945,504,986,265,1000,0,0,0,15,267,56,506,119,708,202,864,299,964]];
const g56 = [[994,713,860,482,848,472,835,477,828,490,826,565,525,565,526,380,537,370,594,368,600,356,599,341,303,0,294,3,5,334,0,349,3,363,12,370,65,370,73,376,79,508,87,541,100,569,116,591,136,605,157,610,834,610,844,605,851,593,854,531,973,736,854,941,851,879,844,867,834,862,546,862,546,907,828,907,831,982,842,997,854,999,863,990,997,759,997,720]];
const g57 = [[994,597,849,274,836,260,822,267,815,285,812,390,487,390,487,0,0,0,0,260,3,312,12,358,26,398,43,428,64,448,87,455,818,455,829,448,837,431,840,344,969,631,840,918,839,843,834,822,824,809,508,808,508,870,812,870,813,968,820,988,833,1000,843,996,850,986,995,663,999,643,997,608]];
const g58 = [[994,599,849,277,836,263,822,270,815,288,812,393,487,393,487,0,0,0,0,263,3,314,12,360,26,400,43,430,64,450,87,456,818,456,829,450,837,433,840,346,969,632,840,918,839,843,834,822,824,810,508,808,508,870,812,870,813,968,820,989,833,1000,843,996,850,986,995,664,999,644,997,608]];
const g59 = [[0,0,152,0,168,121,201,267,245,394,298,499,360,579,427,629,500,646,573,629,640,579,702,499,755,394,799,267,832,121,848,0,1000,0,962,294,903,528,824,723,728,872,619,967,500,1000,381,967,272,872,176,723,97,528,38,294,1,30]];
const g60 = [[500,0,619,33,728,128,824,277,903,472,962,706,1000,1000,848,1000,832,879,799,733,755,606,702,501,640,421,573,371,500,354,427,371,360,421,298,501,245,606,201,733,168,879,152,1000,0,1000,38,706,97,472,176,277,272,128,381,33]];
const g61 = [[0,0,152,0,168,121,201,267,245,394,298,499,360,579,427,629,500,646,573,629,640,579,702,499,755,394,799,267,832,121,848,0,1000,0,962,294,903,528,824,723,728,872,619,967,500,1000,381,967,272,872,176,723,97,528,38,294,1,30]];
const g62 = [[500,0,619,33,728,128,824,277,903,472,962,706,1000,1000,848,1000,832,879,799,733,755,606,702,501,640,421,573,371,500,354,427,371,360,421,298,501,245,606,201,733,168,879,152,1000,0,1000,38,706,97,472,176,277,272,128,381,33]];
const g63 = [[500,0,633,34,752,129,854,277,932,468,982,693,1000,944,997,1000,853,1000,842,767,806,607,751,471,679,367,594,299,500,275,406,299,321,367,249,471,194,607,158,767,147,1000,3,1000,0,944,18,693,68,468,146,277,248,129,367,34]];
const g64 = [[3,0,147,0,158,233,194,393,249,529,321,633,406,701,500,725,594,701,679,633,751,529,806,393,842,233,853,0,997,0,1000,56,982,307,932,532,854,723,752,871,633,966,500,1000,367,966,248,871,146,723,68,532,18,307,0,56]];
const g65 = [[500,0,633,34,752,129,854,277,932,468,982,693,1000,944,997,1000,853,1000,842,767,806,607,751,471,679,367,594,299,500,275,406,299,321,367,249,471,194,607,158,767,147,1000,3,1000,0,944,18,693,68,468,146,277,248,129,367,34]];
const g66 = [[3,0,147,0,158,233,194,393,249,529,321,633,406,701,500,725,594,701,679,633,751,529,806,393,842,233,853,0,997,0,1000,56,982,307,932,532,854,723,752,871,633,966,500,1000,367,966,248,871,146,723,68,532,18,307,0,56]];
const g67 = [[500,460,353,419,295,319,357,222,505,181,652,222,714,319,670,403,528,458],[295,25,57,166,2,345,414,966,519,999,967,438,1000,319,973,215,783,56,473,0]];
const g68 = [[943,934,1000,797,959,679,316,41,154,7,0,123,877,1000]];
const g69 = [[692,683,317,349,359,268,90,0,45,45,32,316,370,758,644,952,901,988,1000,909,733,641]];
const g70 = [[947,662,300,27,152,7,0,127,877,1000,993,849,963,681]];
const g71 = [[0,0,861,0,920,85,953,167,977,255,991,343,1000,500,987,657,970,745,944,833,908,915,847,1000,0,1000]];
const g72 = [[249,0,1000,0,1000,1000,224,1000,239,1000,248,986,283,915,310,833,329,745,348,573,350,427,334,255,316,167,292,85,258,14]];
const g73 = [[0,0,1000,0,300,1000,0,578]];
const g74 = [[0,645,1000,1000,510,18,462,0,124,182,120,188]];
const g75 = [[0,650,1000,1000,492,0,116,209]];

// ---------------------------------------------------------------- slide 1
function slide1 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g38,-.03,2.54,7.75,4.98,{f:{color:NAVY}});
  sh(s,9.89,-.02,3.48,1.7,{shape:'rtTriangle',f:{color:NAVY},rot:180});

  poly(s,g37,-.02,-.02,13.36,7.52,{f:PH1});
  poly(s,g33,1.38,4.87,6.9,.63,{f:{color:OFFWHITE},rot:28.95,fh:1});
  sh(s,6.36,3.24,6.99,4.29,{shape:'rtTriangle',f:{color:GOLD,transparency:12},fh:1});
  tx(s,.45,4.58,2.41,1.62,[['THE ',{face:'Montserrat SemiBold',sz:30,c:WHITE,br:1}],['AGENCY',{face:'Montserrat SemiBold',sz:30,c:WHITE,br:1}],['BUSINESS',{face:'Montserrat SemiBold',sz:30,c:WHITE}]],{});
  tx(s,.51,6.34,4.31,.63,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:PALE}]],{ls:1.5});
  tx(s,9.13,6.56,3.74,.34,[['Lorem Ipsum is simply dummy text of the printing',{face:'Lato',sz:11,c:MIDGRAY}]],{a:'right',ls:1.5});
  tx(s,10.53,6.22,2.34,.34,[['Creative Innovation',{face:'Poppins Medium',sz:14,c:NAVY}]],{a:'right'});
  poly(s,g12,11.62,5.11,.56,.9,{f:{color:NAVY}});
  poly(s,g13,11.88,5.32,.03,.08,{f:{color:NAVY}});
  poly(s,g14,11.78,5.35,.06,.07,{f:{color:NAVY}});
  poly(s,g15,11.94,5.35,.06,.07,{f:{color:NAVY}});
  poly(s,g16,11.35,4.92,1.08,.56,{f:{color:NAVY}});
}

// ---------------------------------------------------------------- slide 2
function slide2 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,-.02,-.03,6.67,7.54,{f:{color:NAVY},rot:180});
  poly(s,g39,-.03,4.57,5.08,1.66,{f:{color:GOLD}});

  img(s,6.64,-.03,6.69,7.53,{f:PH1});
  img(s,3.47,4.76,1.32,1.28,{f:PH1,shape:'ellipse',l:{color:WHITE,width:1}});
  tx(s,1.17,1.19,3.63,1.04,[['Welcome to the ',{face:'Poppins SemiBold',sz:28,c:WHITE}],['Agency Business ',{face:'Poppins SemiBold',sz:28,c:GOLD}]],{});
  ln(s,1.28,2.55,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.18,2.88,3.6,1.46,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled.',{face:'Lato',sz:11,c:PALE}]],{a:'justify',ls:1.5});
  tx(s,1.17,4.79,1.91,.32,[['Vincent Poernama',{face:'Poppins Medium',sz:13,c:NAVY}]],{});
  tx(s,1.17,5.17,2.06,.87,[['“Lorem Ipsum is simply dummy text of the printing and typesetting industry. “ ',{face:'Lato',sz:10.5,c:MIDGRAY,i:1}]],{ls:1.5});
  sh(s,6.62,4.15,6.71,2.79,{f:{color:NAVY,transparency:12.16}});
  tx(s,7.76,4.53,4.44,1.16,[['An professionals consisting of experts with modern solutions.',{face:'Poppins SemiBold',sz:21,c:WHITE}]],{a:'center'});
  tx(s,7.88,5.85,4.18,.35,[['Lorem Ipsum is simply dummy text of the printing. ',{face:'Lato',sz:11,c:PALE}]],{a:'center',ls:1.5});
  sh(s,9.34,6.35,.21,.21,{shape:'ellipse',f:{color:GOLD},l:{color:GOLD,width:1}});
  sh(s,9.7,6.35,.21,.21,{shape:'ellipse',l:{color:GOLD,width:1}});
  sh(s,10.05,6.35,.21,.21,{shape:'ellipse',l:{color:GOLD,width:1}});
  sh(s,10.41,6.35,.21,.21,{shape:'ellipse',l:{color:GOLD,width:1}});
  tx(s,6.26,5.24,.76,.61,[['<',{face:'Poppins ExtraBold',sz:36,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rot:180,fh:1,rr:.1});
}

// ---------------------------------------------------------------- slide 3
function slide3 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,9.16,-.02,4.18,7.53,{f:{color:GOLD}});

  img(s,-.01,-.03,9.17,7.53,{f:PH1});
  fade(s,4.58,-.03,4.6,7.53,NAVY,0,16.08);
  tx(s,5.36,1.89,3.05,.81,[['Learning the Value of Helping Hand',{face:'Poppins SemiBold',sz:21,c:WHITE}]],{});
  ln(s,5.47,2.92,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,5.34,3.2,3.07,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make.',{face:'Lato',sz:11,c:PALE}]],{a:'justify',ls:1.5});
  tx(s,5.46,5.15,1.43,.46,[['VIEW MORE',{face:'Poppins Medium',sz:10,c:OFFWHITE}]],{v:'middle',a:'center',l:{color:OFFWHITE,width:1}});
  tx(s,8.73,1.3,.91,.73,[['01',{face:'Poppins Medium',sz:16,c:NAVY}]],{v:'middle',a:'center',shape:'parallelogram',f:{color:OFFWHITE}});
  tx(s,9.91,1.37,2.09,.31,[['Outsourching Services',{face:'Poppins Medium',sz:12,c:INK}]],{});
  tx(s,9.91,1.71,2.36,.62,[['Lorem Ipsum is simply dummy text of the printing.',{face:'Lato',sz:11,c:MIDGRAY}]],{ls:1.5});
  tx(s,8.73,5.38,.91,.73,[['03',{face:'Poppins Medium',sz:16,c:NAVY}]],{v:'middle',a:'center',shape:'parallelogram',f:{color:OFFWHITE}});
  tx(s,9.9,5.44,2.17,.31,[['Network System Design',{face:'Poppins Medium',sz:12,c:INK}]],{});
  tx(s,9.91,5.78,2.36,.62,[['Lorem Ipsum is simply dummy text of the printing.',{face:'Lato',sz:11,c:MIDGRAY}]],{ls:1.5});
  tx(s,8.72,3.38,.91,.73,[['02',{face:'Poppins Medium',sz:16,c:NAVY}]],{v:'middle',a:'center',shape:'parallelogram',f:{color:OFFWHITE}});
  tx(s,9.9,3.44,1.93,.31,[['Software Acquisition',{face:'Poppins Medium',sz:12,c:INK}]],{});
  tx(s,9.9,3.79,2.36,.62,[['Lorem Ipsum is simply dummy text of the printing.',{face:'Lato',sz:11,c:MIDGRAY}]],{ls:1.5});
}

// ---------------------------------------------------------------- slide 4
function slide4 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  tx(s,1.17,4.94,3.23,1.41,[['Maintance and Post Sales Services',{face:'Poppins SemiBold',sz:26,c:NAVY}]],{});
  ln(s,1.28,6.51,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,4.95,4.98,2.42,.34,[['Consisting of Expert',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,4.95,5.44,2.58,1.13,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:10.5,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,8.13,4.98,3,.34,[['Implementation of Network',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,8.13,5.44,3.94,1.13,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type.',{face:'Lato',sz:10.5,c:GRAY}]],{a:'justify',ls:1.5});
  img(s,10.95,0,2.39,4.23,{f:PH1});
  img(s,2.52,0,8.29,4.23,{f:PH2});
  img(s,-.01,0,2.37,4.23,{f:PH1});
}

// ---------------------------------------------------------------- slide 5
function slide5 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,10.07,-.03,3.27,7.54,{f:{color:NAVY},rot:180});
  sh(s,1.19,1.2,1.64,1.58,{shape:'ellipse',l:{color:GOLD,width:3}});

  img(s,1.27,3.05,1.48,1.42,{f:PH1,shape:'ellipse'});
  img(s,1.27,4.8,1.48,1.42,{f:PH1,shape:'ellipse'});
  img(s,1.27,1.28,1.48,1.42,{f:PH1,shape:'ellipse'});
  img(s,8.07,1.29,3.96,4.94,{f:PH1});
  sh(s,1.27,3.05,1.48,1.42,{shape:'ellipse',f:{color:NAVY,transparency:30.2}});
  tx(s,1.76,3.41,.49,.71,[['2',{face:'Poppins SemiBold',sz:36,c:WHITE}]],{a:'center'});
  sh(s,1.27,4.78,1.48,1.42,{shape:'ellipse',f:{color:NAVY,transparency:30.2}});
  tx(s,1.76,5.14,.51,.71,[['3',{face:'Poppins SemiBold',sz:36,c:WHITE}]],{a:'center'});
  tx(s,12.32,3.44,.76,.61,[['>',{face:'Poppins ExtraBold',sz:36,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.1});
  tx(s,3.61,1.7,3.61,1.26,[['Stop Optimizing for Programmers and Start for Users',{face:'Poppins SemiBold',sz:23,c:NAVY}]],{});
  ln(s,3.75,3.21,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,3.64,3.51,3.57,2.29,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
}

// ---------------------------------------------------------------- slide 6
function slide6 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  img(s,-.02,0,13.31,3.75,{f:PH1});
  sh(s,-.02,0,13.36,3.75,{f:{color:NAVY,transparency:12.16}});
  sh(s,1.26,2.32,2.75,2.04,{shape:'roundRect',f:{color:OFFWHITE},rr:.05});
  sh(s,1.25,4.6,2.75,2.04,{shape:'roundRect',f:{color:GOLD},rr:.05});
  tx(s,5.98,.9,6.13,.91,[['The Most Suitable Network Solutions in Many Different Areas',{face:'Poppins SemiBold',sz:24,c:WHITE}]],{a:'right'});
  ln(s,10.66,2.09,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.62,2.59,2.02,.34,[['The Agency Vision',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,1.63,3.12,2.02,.97,[['“Our vision provides the best solutions to give customer.“ ',{face:'Lato',sz:12,c:GRAY,i:1}]],{ls:1.5});
  tx(s,1.56,4.91,2.14,.34,[['The Agency Mission',{face:'Poppins Medium',sz:14,c:INK}]],{});
  tx(s,1.56,5.44,2.14,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:MIDGRAY}]],{ls:1.5});
  tx(s,4.42,4.95,2.42,.34,[['Professional Team',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,4.53,5.48,2.19,.9,[['Lorem Ipsum  dummy text of the printing and typesetting industry lorem Ipsum simply. ',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,7.15,4.99,2.42,.34,[['Cloud Computing',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,7.22,5.48,2.26,.9,[['Lorem Ipsum  dummy text of the printing and typesetting industry lorem Ipsum simply. ',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,9.83,4.99,2.42,.34,[['Build Your Network',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,9.91,5.49,2.26,.9,[['Lorem Ipsum  dummy text of the printing and typesetting industry lorem Ipsum simply. ',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  img(s,4.71,2.9,1.85,1.78,{f:PH1,shape:'ellipse'});
  img(s,7.43,2.89,1.85,1.78,{f:PH1,shape:'ellipse'});
  img(s,10.13,2.9,1.85,1.78,{f:PH1,shape:'ellipse'});
}

// ---------------------------------------------------------------- slide 7
function slide7 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,-.01,-.03,4.34,5.22,{f:{color:NAVY},rot:180});
  poly(s,g34,6.75,1.28,2.31,2.28,{l:{color:GOLD,width:2.25}});
  poly(s,g34,6.75,3.94,2.31,2.28,{l:{color:GOLD,width:2.25}});

  tx(s,-.25,3.55,1.88,.44,[['THE AGENCY',{face:'Poppins Medium',sz:20,c:GOLD}]],{rot:-90});
  ln(s,.67,1.07,0,1.49,{l:{color:GOLD,width:1.5}});
  tx(s,9.36,1.47,2.65,.74,[['The Multi Concept Projects',{face:'Poppins SemiBold',sz:19,c:NAVY}]],{});
  tx(s,9.36,2.35,2.7,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s.',{face:'Lato',sz:11,c:GRAY}]],{v:'bottom',a:'justify',ls:1.5});
  tx(s,9.37,4.25,2.65,.74,[['Endpoint Management',{face:'Poppins SemiBold',sz:19,c:NAVY}]],{});
  tx(s,9.37,5.13,2.7,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s.',{face:'Lato',sz:11,c:GRAY}]],{v:'bottom',a:'justify',ls:1.5});
  img(s,1.27,.65,4.56,6.85,{f:PH1});
  img(s,6.81,1.34,2.17,2.15,{f:PH1,shape:'ellipse'});
  img(s,6.81,4.01,2.17,2.15,{f:PH1,shape:'ellipse'});
}

// ---------------------------------------------------------------- slide 8
function slide8 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,7.3,4.9,1.55,1.3,{shape:'roundRect',f:{color:GOLD},rr:.03});
  sh(s,10.6,4.91,1.4,1.3,{shape:'roundRect',f:{color:GOLD},rr:.03});
  sh(s,9.05,4.91,1.4,1.3,{shape:'roundRect',f:{color:OFFWHITE},rr:.03});
  poly(s,g35,3.47,1.45,2.48,3.26,{f:PH1});
  img(s,1.26,.74,2.02,2.84,{f:PH1});
  img(s,1.27,3.75,3.43,2.84,{f:PH2});
  poly(s,g35,3.47,1.45,2.48,3.26,{f:{color:NAVY,transparency:12.16}});
  tx(s,7.18,1.19,4.93,1.04,[['A Community Where Everyone is Welcome',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{});
  ln(s,7.28,2.55,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,7.18,3.21,4.93,1.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries, but also the leap into electronic typesetting.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,7.37,5.09,1.39,.57,[['3726',{face:'Poppins Medium',sz:28,c:NAVY}]],{a:'center'});
  tx(s,7.29,5.72,1.58,.29,[['Clients',{face:'Poppins Light',sz:11,c:DARK}]],{a:'center'});
  tx(s,9.07,5.09,1.39,.57,[['1382',{face:'Poppins Medium',sz:28,c:NAVY}]],{a:'center'});
  tx(s,8.99,5.72,1.58,.29,[['Reviews',{face:'Poppins Light',sz:11,c:DARK}]],{a:'center'});
  tx(s,10.63,5.09,1.39,.57,[['5234',{face:'Poppins Medium',sz:28,c:NAVY}]],{a:'center'});
  tx(s,10.55,5.72,1.58,.29,[['Portfolios',{face:'Poppins Light',sz:11,c:DARK}]],{a:'center'});
  tx(s,7.18,2.88,2.42,.34,[['Join Our Works',{face:'Poppins Medium',sz:14,c:DARK}]],{});
}

// ---------------------------------------------------------------- slide 9
function slide9 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  tx(s,6.08,1.19,6.03,1.01,[['An Professionals Consisting of Experts with Modern Solutions',{face:'Poppins SemiBold',sz:27,c:NAVY}]],{});
  ln(s,6.15,2.42,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,7.86,3.17,4.24,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,7.84,2.83,2.52,.34,[['Professional Services',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,7.86,5.02,4.24,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,7.84,4.68,2.76,.34,[['IT Organoization Setup',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  img(s,-.01,-.01,4.84,7.53,{f:PH1});
  img(s,6.17,2.86,1.47,1.46,{f:PH1,shape:'ellipse'});
  img(s,6.17,4.74,1.47,1.46,{f:PH1,shape:'ellipse'});
}

// ---------------------------------------------------------------- slide 10
function slide10 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,1.29,5.05,3.12,1.28,{shape:'roundRect',f:{color:NAVY},rr:.03});
  sh(s,8.93,5.05,3.12,1.28,{shape:'roundRect',f:{color:NAVY},rr:.03});
  sh(s,5.13,5.05,3.12,1.28,{shape:'roundRect',f:{color:NAVY},rr:.03});

  sh(s,1.53,5.85,.34,.32,{shape:'ellipse',f:{color:GOLD}});
  tx(s,1.53,5.83,.35,.4,[['+',{face:'Poppins Medium',sz:18,c:NAVY}]],{});
  tx(s,2.14,5.63,2.02,.34,[['Nikita Hernandez',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  tx(s,2.14,5.96,1.1,.3,[['Manager',{face:'Poppins Light',sz:12,c:PALE}]],{});
  sh(s,5.38,5.85,.34,.32,{shape:'ellipse',f:{color:GOLD}});
  tx(s,5.38,5.83,.35,.4,[['+',{face:'Poppins Medium',sz:18,c:NAVY}]],{});
  tx(s,6.01,5.63,2.02,.34,[['Raphael  Montana',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  tx(s,6.01,5.96,1.1,.3,[['Director',{face:'Poppins Light',sz:12,c:PALE}]],{});
  sh(s,9.18,5.85,.34,.32,{shape:'ellipse',f:{color:GOLD}});
  tx(s,9.17,5.82,.35,.4,[['+',{face:'Poppins Medium',sz:18,c:NAVY}]],{});
  tx(s,9.8,5.63,2.02,.34,[['Margareth Waber',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  tx(s,9.8,5.96,1.44,.3,[['Head Finance',{face:'Poppins Light',sz:12,c:PALE}]],{});
  ln(s,9.38,1.31,3.96,0,{l:{color:GOLD,width:6},rot:180});
  tx(s,3.56,1.73,6.22,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{v:'bottom',a:'center',ls:1.5});
  tx(s,4.2,1.02,4.93,.57,[['Meet Our Great Team',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{a:'center'});
  poly(s,g31,1.28,2.9,3.12,2.79,{f:PH1,l:{color:WHITE,width:2}});
  poly(s,g31,5.13,2.9,3.12,2.79,{f:PH1,l:{color:WHITE,width:2}});
  poly(s,g31,8.93,2.9,3.12,2.79,{f:PH1,l:{color:WHITE,width:2}});
}

// ---------------------------------------------------------------- slide 11
function slide11 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,-.01,-.02,13.38,4.05,{f:{color:NAVY},rot:180});

  img(s,1.37,2.77,2.6,2.53,{f:PH1,shape:'ellipse',l:{color:WHITE,width:2}});
  img(s,5.35,2.77,2.6,2.53,{f:PH1,shape:'ellipse',l:{color:WHITE,width:2}});
  img(s,9.34,2.77,2.6,2.53,{f:PH1,shape:'ellipse',l:{color:WHITE,width:2}});
  tx(s,4.2,.99,4.93,.57,[['Our Professional Team',{face:'Poppins SemiBold',sz:28,c:WHITE}]],{a:'center'});
  tx(s,3.56,1.73,6.22,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:PALE}]],{v:'bottom',a:'center',ls:1.5});
  tx(s,5.42,5.47,2.49,.37,[['Bruno Emmanuelle',{face:'Montserrat Medium',sz:16,c:DARK}]],{a:'center'});
  tx(s,5.73,4.94,1.84,.35,[['Team Leader',{face:'Poppins',sz:14,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.01});
  tx(s,5,5.88,3.34,.62,[['Lorem Ipsum',{face:'Lato Black',sz:11,c:MIDGRAY}],[' ',{face:'Lato',sz:11,c:MIDGRAY}],['is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,1.49,5.47,2.37,.37,[['Daria Litvinova',{face:'Montserrat Medium',sz:16,c:DARK}]],{a:'center'});
  tx(s,1.75,4.94,1.84,.35,[['Marketing',{face:'Poppins',sz:14,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.01});
  tx(s,1,5.88,3.34,.62,[['Lorem Ipsum',{face:'Lato Black',sz:11,c:MIDGRAY}],[' ',{face:'Lato',sz:11,c:MIDGRAY}],['is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,9.45,5.47,2.37,.37,[['Clarissa Curie',{face:'Montserrat Medium',sz:16,c:DARK}]],{a:'center'});
  tx(s,8.97,5.88,3.34,.62,[['Lorem Ipsum',{face:'Lato Black',sz:11,c:MIDGRAY}],[' ',{face:'Lato',sz:11,c:MIDGRAY}],['is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,9.72,4.94,1.84,.35,[['Law Staff',{face:'Poppins',sz:14,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.01});
}

// ---------------------------------------------------------------- slide 12
function slide12 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,-.01,-.02,6.6,7.53,{shape:'round1Rect',f:{color:NAVY},rr:1.1});
  sh(s,12.06,-.02,1.29,7.53,{f:{color:GOLD}});

  tx(s,10.23,3.46,4.93,.57,[['SINGLE PROFILE',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{a:'center',rot:90});
  tx(s,1.2,1.15,4.46,1.62,[['John Alexandra',{face:'Poppins SemiBold',sz:45,c:WHITE}]],{});
  tx(s,1.2,2.76,3.19,.37,[['President Director',{face:'Montserrat Medium',sz:16,c:OFFWHITE}]],{});
  ln(s,1.3,3.75,4.19,0,{l:{color:GOLD,width:2.25}});
  tx(s,1.2,3.97,4.44,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:PALE}]],{a:'justify',ls:1.5});
  tx(s,1.18,5.25,.79,.57,[['8.8',{face:'Poppins SemiBold',sz:28,c:GOLD}]],{});
  tx(s,1.2,5.83,1.48,.34,[['Total Rating',{face:'Poppins Light',sz:14,c:OFFWHITE}]],{});
  tx(s,4.32,5.25,1.05,.57,[['+150',{face:'Poppins SemiBold',sz:28,c:GOLD}]],{});
  tx(s,4.34,5.83,1.27,.34,[['Experience',{face:'Poppins Light',sz:14,c:OFFWHITE}]],{});
  tx(s,2.76,5.25,1.15,.57,[['+250',{face:'Poppins SemiBold',sz:28,c:GOLD}]],{});
  tx(s,2.78,5.83,1.15,.34,[['Portfolios',{face:'Poppins Light',sz:14,c:OFFWHITE}]],{});
  poly(s,g40,5.57,0,6.48,7.5,{f:PH1});
}

// ---------------------------------------------------------------- slide 13
function slide13 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,-.01,5.46,13.38,2.07,{f:{color:NAVY},rot:180});
  poly(s,g41,1.29,5.46,2.41,1.05,{f:{color:GOLD}});
  poly(s,g42,9.6,5.18,2.41,1.33,{f:{color:GOLD}});
  poly(s,g43,4.05,5.46,2.39,1.05,{f:{color:GOLD}});
  poly(s,g44,6.94,5.46,2.41,1.05,{f:{color:GOLD}});

  tx(s,3.56,1.73,6.22,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{v:'bottom',a:'center',ls:1.5});
  tx(s,3.85,1.02,5.64,.57,[['Services From The Experts',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{a:'center'});
  tx(s,1.48,5.61,2.02,.34,[['Jhonny Wang',{face:'Poppins Medium',sz:14,c:NAVY}]],{a:'center'});
  sh(s,2,6.07,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g25,2.06,6.14,.15,.13,{f:{color:NAVY}});
  sh(s,2.39,6.08,.25,.25,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g26,2.45,6.15,.14,.12,{f:{color:NAVY}});
  poly(s,g27,2.85,6.16,.09,.09,{f:{color:NAVY}});
  poly(s,g28,2.95,6.18,.04,.04,{f:{color:NAVY}});
  sh(s,2.79,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  tx(s,4.25,5.62,2.02,.34,[['Robert Williams',{face:'Poppins Medium',sz:14,c:NAVY}]],{a:'center'});
  sh(s,4.76,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g25,4.82,6.15,.15,.13,{f:{color:NAVY}});
  sh(s,5.16,6.08,.25,.25,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g26,5.21,6.15,.14,.12,{f:{color:NAVY}});
  poly(s,g27,5.61,6.16,.09,.09,{f:{color:NAVY}});
  poly(s,g28,5.71,6.19,.04,.04,{f:{color:NAVY}});
  sh(s,5.55,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  tx(s,7.14,5.62,2.02,.34,[['Matthew White',{face:'Poppins Medium',sz:14,c:NAVY}]],{a:'center'});
  sh(s,7.66,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g25,7.71,6.15,.15,.13,{f:{color:NAVY}});
  sh(s,8.05,6.08,.25,.25,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g26,8.1,6.15,.14,.12,{f:{color:NAVY}});
  poly(s,g27,8.5,6.16,.09,.09,{f:{color:NAVY}});
  poly(s,g28,8.6,6.19,.04,.04,{f:{color:NAVY}});
  sh(s,8.45,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  tx(s,9.8,5.62,2.02,.34,[['Michael Burton',{face:'Poppins Medium',sz:14,c:NAVY}]],{a:'center'});
  sh(s,10.32,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g25,10.37,6.15,.15,.13,{f:{color:NAVY}});
  sh(s,10.71,6.08,.25,.25,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g26,10.76,6.15,.14,.12,{f:{color:NAVY}});
  poly(s,g27,11.16,6.16,.09,.09,{f:{color:NAVY}});
  poly(s,g28,11.26,6.19,.04,.04,{f:{color:NAVY}});
  sh(s,11.11,6.08,.26,.26,{shape:'roundRect',l:{color:NAVY,width:.75},rr:.04});
  poly(s,g29,1.29,2.67,2.39,2.79,{f:PH1});
  poly(s,g29,9.61,2.67,2.41,2.79,{f:PH1});
  poly(s,g29,4.05,2.67,2.39,2.79,{f:PH1});
  poly(s,g29,6.94,2.67,2.41,2.79,{f:PH1});
}

// ---------------------------------------------------------------- slide 14
function slide14 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,7.14,1.31,7.5,4.88,{f:{color:NAVY},rot:-90,fh:1});
  ln(s,1.34,4.83,3.28,0,{l:{color:OFFWHITE,width:2.25}});
  ln(s,1.34,2.93,3.28,0,{l:{color:OFFWHITE,width:2.25}});
  sh(s,1.28,1.18,3.33,1.35,{shape:'roundRect',f:{color:WHITE},l:{color:GOLD,width:3},rr:.03});

  img(s,4.86,-.01,3.61,7.51,{f:PH1});
  sh(s,4.16,5.29,.93,.93,{shape:'flowChartAlternateProcess',f:{color:WHITE},l:{color:GOLD,width:2},fh:1});
  sh(s,4.16,3.35,.93,.93,{shape:'flowChartAlternateProcess',f:{color:WHITE},l:{color:GOLD,width:2},fh:1});
  tx(s,1.48,1.42,2.6,.34,[['Endpoint Management',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'right'});
  tx(s,1.39,1.71,2.7,.62,[['Lorem Ipsum is simply dummy text of the printing typesetting.',{face:'Lato',sz:11,c:GRAY}]],{a:'right',ls:1.5});
  tx(s,1.29,3.41,2.8,.34,[['18',{face:'Poppins Medium',sz:14,c:DARK}],['th',{face:'Poppins Medium',sz:14,c:DARK,sup:1}],[' Years Experience',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'right'});
  tx(s,1.39,3.7,2.7,.62,[['Lorem Ipsum is simply dummy text of the printing typesetting.',{face:'Lato',sz:11,c:GRAY}]],{a:'right',ls:1.5});
  tx(s,1.43,5.3,2.65,.34,[['Professionals Solution',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'right'});
  tx(s,1.39,5.59,2.7,.62,[['Lorem Ipsum is simply dummy text of the printing typesetting.',{face:'Lato',sz:11,c:GRAY}]],{a:'right',ls:1.5});
  sh(s,4.16,1.42,.93,.93,{shape:'flowChartAlternateProcess',f:{color:GOLD},rot:-90});
  tx(s,9.01,3.74,3.1,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'justify',ls:1.5});
  tx(s,9.02,3.35,2.67,.34,[['Manage Your Business',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  tx(s,9.02,2.03,3.1,.91,[['Introducing Special Services',{face:'Poppins SemiBold',sz:24,c:OFFWHITE}]],{});
  poly(s,g45,7.71,3.28,.93,.93,{f:{color:'1F3864'},rot:-90,fh:1});
  tx(s,8.07,3.33,.47,.84,[['<',{face:'Poppins Medium',sz:44,c:GOLD}]],{});
  poly(s,g5,4.43,1.59,.4,.59,{f:{color:NAVY}});
  poly(s,g6,4.5,1.66,.26,.26,{f:{color:NAVY}});
  poly(s,g7,4.54,1.7,.18,.17,{f:{color:NAVY}});
  poly(s,g3,4.54,5.52,.18,.18,{f:{color:NAVY}});
  poly(s,g4,4.37,5.46,.52,.59,{f:{color:NAVY}});
  poly(s,g8,4.53,3.66,.12,.12,{f:{color:NAVY}});
  poly(s,g9,4.49,3.79,.2,.1,{f:{color:NAVY}});
  poly(s,g10,4.34,3.53,.59,.59,{f:{color:NAVY}});
  poly(s,g11,4.73,3.92,.12,.12,{f:{color:NAVY}});
}

// ---------------------------------------------------------------- slide 15
function slide15 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g47,6.67,3.95,6.68,3.55,{f:{color:NAVY}});

  tx(s,8.65,2.96,3.44,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,8.66,2.62,2.34,.34,[['Marketing Business',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,1.2,5.87,3.36,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.2,5.53,2.7,.34,[['18',{face:'Poppins Medium',sz:14,c:DARK}],['th',{face:'Poppins Medium',sz:14,c:DARK,sup:1}],['  Years Experience',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  poly(s,g17,8.72,1.68,.81,.81,{f:{color:GOLD}});
  poly(s,g18,9.06,2.25,.07,.07,{f:{color:GOLD}});
  poly(s,g19,8.79,1.84,.33,.33,{f:{color:GOLD}});
  poly(s,g20,8.89,1.93,.14,.14,{f:{color:GOLD}});
  poly(s,g2,9.08,2.15,.05,.02,{f:{color:GOLD}});
  poly(s,g2,9.15,2.15,.05,.02,{f:{color:GOLD}});
  poly(s,g21,9.22,2.15,.05,.02,{f:{color:GOLD}});
  poly(s,g22,9.29,2.15,.05,.02,{f:{color:GOLD}});
  poly(s,g23,9.35,2.15,.05,.02,{f:{color:GOLD}});
  poly(s,g24,9.27,1.74,.16,.25,{f:{color:GOLD}});
  tx(s,9.38,6.07,2.73,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been.',{face:'Lato',sz:11,c:PALE}]],{a:'center',ls:1.5});
  tx(s,9.57,5.73,2.34,.34,[['Creative Innovation',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{a:'center'});
  tx(s,5.07,4.75,3.41,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  ',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,5.07,4.41,2.34,.34,[['Professional Solution',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  poly(s,g3,5.43,3.55,.24,.23,{f:{color:GOLD}});
  poly(s,g4,5.22,3.47,.67,.76,{f:{color:GOLD}});
  poly(s,g8,1.55,4.78,.16,.17,{f:{color:GOLD}});
  poly(s,g9,1.49,4.96,.28,.14,{f:{color:GOLD}});
  poly(s,g10,1.28,4.59,.82,.82,{f:{color:GOLD}});
  poly(s,g11,1.83,5.14,.17,.17,{f:{color:GOLD}});
  poly(s,g12,10.56,4.86,.37,.6,{f:{color:GOLD}});
  poly(s,g13,10.73,5,.02,.05,{f:{color:GOLD}});
  poly(s,g14,10.67,5.02,.04,.05,{f:{color:GOLD}});
  poly(s,g15,10.78,5.02,.04,.05,{f:{color:GOLD}});
  poly(s,g16,10.39,4.74,.72,.37,{f:{color:GOLD}});
  ln(s,10.74,1.07,2.58,0,{l:{color:GOLD,width:2.25}});
  tx(s,7.77,.77,2.8,.57,[['Our Services',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{});
  poly(s,g46,-.01,0,7.53,3.55,{f:PH1});
}

// ---------------------------------------------------------------- slide 16
function slide16 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  img(s,1.27,.86,4.89,1.71,{f:PH1,l:{color:NAVY,width:1.5}});
  img(s,7.14,2.93,4.89,1.71,{f:PH1,l:{color:NAVY,width:1.5}});
  img(s,1.27,3.89,4.89,1.76,{f:PH1,l:{color:NAVY,width:1.5}});
  poly(s,g3,1.47,2.84,.17,.16,{f:{color:GOLD}});
  poly(s,g4,1.31,2.79,.48,.55,{f:{color:GOLD}});
  tx(s,2.03,3.06,4.13,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  ',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,2.03,2.72,2.34,.34,[['Professional Solution',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,2.02,6.17,4.13,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  ',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,2.03,5.83,2.54,.34,[['Endpoint Management',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,7.86,5.23,4.13,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  ',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,7.87,4.89,2.34,.34,[['18',{face:'Poppins Medium',sz:14,c:DARK}],['th',{face:'Poppins Medium',sz:14,c:DARK,sup:1}],[' Years Experience',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,7.02,1.46,5.06,1.04,[['Business Strategy and Investment.',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{});
  tx(s,5.59,3.9,.57,.42,[['02',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{v:'middle',a:'center',shape:'roundRect',f:{color:NAVY},rr:.01});
  tx(s,11.47,2.94,.57,.42,[['03',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{v:'middle',a:'center',shape:'roundRect',f:{color:NAVY},rr:.01});
  poly(s,g5,1.35,5.87,.4,.59,{f:{color:GOLD}});
  poly(s,g6,1.42,5.94,.26,.26,{f:{color:GOLD}});
  poly(s,g7,1.46,5.98,.18,.17,{f:{color:GOLD}});
  poly(s,g8,7.34,5.04,.12,.13,{f:{color:GOLD}});
  poly(s,g9,7.3,5.17,.21,.11,{f:{color:GOLD}});
  poly(s,g10,7.14,4.9,.62,.62,{f:{color:GOLD}});
  poly(s,g11,7.56,5.32,.13,.13,{f:{color:GOLD}});
  tx(s,5.6,.85,.57,.42,[['01',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{v:'middle',a:'center',shape:'roundRect',f:{color:NAVY},rr:.01});
}

// ---------------------------------------------------------------- slide 17
function slide17 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g48,-.01,.87,6.52,3.98,{f:{color:NAVY}});
  sh(s,1.28,2.38,2.78,4.23,{shape:'roundRect',f:{color:WHITE},rr:.07});
  poly(s,g36,1.27,5.9,2.8,.73,{f:{color:GOLD}});
  sh(s,2.06,2.72,1.23,1.14,{shape:'ellipse',f:{color:OFFWHITE}});
  sh(s,4.42,2.38,2.78,4.23,{shape:'roundRect',f:{color:WHITE},rr:.07});
  poly(s,g36,4.41,5.9,2.8,.73,{f:{color:GOLD}});
  sh(s,5.19,2.72,1.23,1.14,{shape:'ellipse',f:{color:OFFWHITE}});

  poly(s,g17,2.34,2.94,.66,.66,{f:{color:NAVY}});
  poly(s,g18,2.62,3.4,.06,.06,{f:{color:NAVY}});
  poly(s,g19,2.4,3.06,.27,.27,{f:{color:NAVY}});
  poly(s,g20,2.48,3.14,.11,.11,{f:{color:NAVY}});
  poly(s,g2,2.64,3.32,.04,.02,{f:{color:NAVY}});
  poly(s,g2,2.69,3.32,.04,.02,{f:{color:NAVY}});
  poly(s,g21,2.75,3.32,.04,.02,{f:{color:NAVY}});
  poly(s,g22,2.8,3.32,.04,.02,{f:{color:NAVY}});
  poly(s,g23,2.86,3.32,.04,.02,{f:{color:NAVY}});
  poly(s,g24,2.79,2.99,.13,.2,{f:{color:NAVY}});
  tx(s,1.41,4.59,2.53,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,1.5,4.16,2.34,.34,[['Marketing Business',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,4.54,4.59,2.53,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,4.54,4.16,2.53,.34,[['Endpoint Management',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,8.13,1.44,4.02,.91,[['Consulting and Professional Services',{face:'Poppins SemiBold',sz:24,c:NAVY}]],{});
  ln(s,8.24,2.59,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,8.13,3.27,3.98,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,8.13,2.93,3.13,.34,[['Business Capabilities Project',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,8.13,4.89,3.98,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,8.13,4.54,3.13,.34,[['Prototype Integration Model',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  poly(s,g5,5.57,2.99,.43,.63,{f:{color:NAVY}});
  poly(s,g6,5.65,3.06,.28,.28,{f:{color:NAVY}});
  poly(s,g7,5.69,3.11,.2,.19,{f:{color:NAVY}});
  tx(s,1.5,6.1,2.34,.34,[['SERVICE 1',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,4.61,6.1,2.34,.34,[['SERVICE 2',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
}

// ---------------------------------------------------------------- slide 18
function slide18 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,0,1.95,10.41,5.55,{shape:'rtTriangle',f:{color:NAVY}});
  sh(s,9.85,-.02,3.48,1.7,{shape:'rtTriangle',f:{color:NAVY},rot:180});

  poly(s,g49,-.02,-.02,13.36,7.52,{f:PH1});
  tx(s,.45,4.86,4.11,1.45,[['SECTION BREAK',{face:'Montserrat SemiBold',sz:40,c:WHITE}]],{});
  tx(s,.51,6.33,5.64,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s text ever since the 1500s.',{face:'Lato',sz:11,c:LTGRAY}]],{ls:1.5});
  poly(s,g33,1.91,4.24,5.6,.51,{f:{color:GOLD},rot:28.07,fh:1});
}

// ---------------------------------------------------------------- slide 19
function slide19 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,8.17,2.76,3.13,4.6,{f:{color:NAVY},rot:-90,fh:1});

  img(s,1.25,3.75,2.51,2.88,{f:PH1,l:{color:NAVY,width:1.5}});
  img(s,4.12,3.74,2.51,2.88,{f:PH1,l:{color:NAVY,width:1.5}});
  tx(s,1.19,1.01,5.41,1.11,[['The Business Solution to Prevent Spreading',{face:'Poppins SemiBold',sz:30,c:NAVY}]],{});
  ln(s,1.32,2.38,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.22,2.58,5.47,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,2.36,3.75,1.4,.42,[['Portfolio 1',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{v:'middle',shape:'roundRect',f:{color:NAVY},rr:.01});
  tx(s,5.22,3.73,1.4,.42,[['Portfolio 2',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{v:'middle',shape:'roundRect',f:{color:NAVY},rr:.01});
  tx(s,7.92,4.53,3.63,1.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen. ',{face:'Lato',sz:11,c:LTGRAY}]],{a:'justify',ls:1.5});
  tx(s,7.93,4.14,2.67,.34,[['Our Greatest Portfolio',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  img(s,7.43,.88,4.6,2.62,{f:PH1});
}

// ---------------------------------------------------------------- slide 20
function slide20 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,7.32,1.48,7.52,4.56,{f:{color:NAVY},rot:-90,fh:1});

  tx(s,1.19,1.01,5.57,1.11,[['Company Culture Encoutages Diversity',{face:'Poppins SemiBold',sz:30,c:NAVY}]],{});
  ln(s,1.32,2.38,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.18,3.06,2.64,1.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.19,2.72,2.63,.34,[['Innovation Solution',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,3.94,3.06,2.64,1.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,3.94,2.72,2.63,.34,[['Complete Project',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  img(s,1.78,4.93,4.82,1.67,{f:PH1});
  img(s,7.88,.83,4.15,5.78,{f:PH1});
  img(s,0,4.93,1.62,1.67,{f:PH2});
}

// ---------------------------------------------------------------- slide 21
function slide21 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,0,-.01,13.33,2.52,{f:{color:NAVY}});
  sh(s,-.02,4.97,4.46,2.53,{f:{color:NAVY}});

  tx(s,.86,1.02,2.68,.9,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.  Lorem Ipsum has been.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'center',ls:1.5});
  tx(s,.9,.68,2.63,.34,[['Segmentation Issue',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{a:'center'});
  tx(s,.84,6.03,2.68,.9,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.  Lorem Ipsum has been.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'center',ls:1.5});
  tx(s,.88,5.69,2.63,.34,[['History Behind Project',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{a:'center'});
  tx(s,9.77,1.01,2.68,.9,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.  Lorem Ipsum has been.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'center',ls:1.5});
  tx(s,9.81,.67,2.63,.34,[['The Company Review',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{a:'center'});
  tx(s,4.58,2.97,4.17,1.57,[['The Evolution Your Business Growth Solutions',{face:'Poppins SemiBold',sz:29,c:NAVY}]],{a:'center'});
  img(s,4.4,-.02,4.52,2.53,{f:PH1});
  img(s,4.4,4.97,8.93,2.55,{f:PH1});
  img(s,-.01,2.51,4.41,2.47,{f:PH2});
  img(s,8.9,2.51,4.44,2.47,{f:PH2});
}

// ---------------------------------------------------------------- slide 22
function slide22 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,9.39,-.01,3.93,2.19,{f:{color:OFFWHITE}});

  img(s,8.8,2.39,2.14,2.03,{f:PH2});
  img(s,6.42,4.62,2.14,2.03,{f:PH2});
  img(s,8.81,4.62,2.14,2.03,{f:PH2});
  poly(s,g50,6.42,.87,5.62,5.76,{f:PH1});
  img(s,6.42,2.39,2.14,2.03,{f:PH2});
  tx(s,6.43,.87,3.76,.76,[['The Agency Business',{face:'Poppins Medium',sz:18,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.02});
  tx(s,1.27,1.59,4.39,.91,[['The Agency Recognized for Certain Purpose',{face:'Poppins SemiBold',sz:24,c:NAVY}]],{});
  ln(s,1.38,2.84,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.27,3.57,4.21,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.27,3.23,2.63,.34,[['Our Business Workshop',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,1.27,5.01,4.21,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.27,4.67,2.63,.34,[['The Company Review',{face:'Poppins Medium',sz:14,c:DARK}]],{});
}

// ---------------------------------------------------------------- slide 23
function slide23 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,.54,.53,12.26,6.4,{f:{color:NAVY}});

  tx(s,1.2,4.69,4.55,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularized.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'justify',ls:1.5});
  tx(s,2.53,1.49,3.21,.84,[['Our Innovative Amazing Business',{face:'Poppins SemiBold',sz:22,c:WHITE}]],{});
  tx(s,1.05,1.28,1.11,.94,[['10',{face:'Montserrat SemiBold',sz:28,c:OFFWHITE,br:1}],['MAY',{face:'Montserrat Medium',sz:20,c:OFFWHITE}]],{a:'center'});
  ln(s,2.23,1.05,0,1.3,{l:{color:GOLD,width:2.25},rot:180});
  tx(s,2.53,1.13,1.78,.37,[['SEMINAR',{face:'Poppins Medium',sz:16,c:GOLD}]],{});
  tx(s,1.19,4.33,3.62,.35,[['Big Companies and Individual',{face:'Poppins Medium',sz:15,c:OFFWHITE}]],{});
  tx(s,2.53,3.1,2.96,.4,[['Christie Indirania',{face:'Montserrat SemiBold',sz:18,c:OFFWHITE}]],{});
  tx(s,2.53,2.71,1.78,.3,[['SPEAKER',{face:'Poppins Medium',sz:12,c:GOLD}]],{});
  tx(s,2.53,3.5,2.23,.3,[['Professional Agency',{face:'Poppins Medium',sz:12,c:PALE}]],{});
  img(s,6.68,.53,6.14,6.4,{f:PH1});
  img(s,1.3,2.74,1.1,1.09,{f:PH1,shape:'ellipse',l:{color:GOLD,width:1.5}});
}

// ---------------------------------------------------------------- slide 24
function slide24 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  img(s,6.67,.88,5.37,2.73,{f:PH1});
  img(s,1.28,3.88,3.33,2.75,{f:PH1});
  img(s,4.98,3.86,3.33,2.75,{f:PH2});
  img(s,8.7,3.88,3.33,2.76,{f:PH1});
  sh(s,6.67,.86,5.37,2.75,{f:{color:NAVY,transparency:12.16}});
  tx(s,7.18,1.63,4.31,.98,[['All Customers Achieve Their Targets',{face:'Poppins SemiBold',sz:26,c:WHITE}]],{a:'center'});
  ln(s,8.67,2.85,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.18,1.38,4.56,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.18,.98,3.43,.34,[['Lead The Organization Team',{face:'Poppins Medium',sz:14,c:MIDGRAY}]],{});
  tx(s,1.18,2.73,4.56,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.18,2.34,3.67,.34,[['Corporations Chartered In Regions',{face:'Poppins Medium',sz:14,c:MIDGRAY}]],{});
}

// ---------------------------------------------------------------- slide 25
function slide25 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g30,-.01,-.02,6.37,7.54,{f:{color:NAVY}});

  tx(s,1.67,5.72,4.31,.61,[['STRENGHTS',{face:'Poppins SemiBold',sz:30,c:WHITE}]],{a:'center'});
  tx(s,7.29,1.56,5.01,1.14,[['Strenghts Business Analysis Slide',{face:'Poppins SemiBold',sz:31,c:NAVY}]],{});
  tx(s,7.29,3.4,4.81,1.18,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  ln(s,7.38,3.05,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,7.29,5,1.03,.64,[['84+',{face:'Poppins SemiBold',sz:32,c:DARK}]],{v:'middle'});
  tx(s,8.83,5,1.11,.64,[['167+',{face:'Poppins SemiBold',sz:32,c:DARK}]],{v:'middle'});
  tx(s,10.45,5,1.27,.64,[['209+',{face:'Poppins SemiBold',sz:32,c:DARK}]],{v:'middle'});
  tx(s,7.29,5.68,1.48,.25,[['“ Lorem Ipsum is simply',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
  tx(s,8.83,5.68,1.48,.25,[['“ Lorem Ipsum is simply',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
  tx(s,10.37,5.68,1.48,.25,[['“ Lorem Ipsum is simply',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
  poly(s,g51,2.07,1.17,3.5,4.06,{f:PH1,l:{color:GOLD,width:3}});
}

// ---------------------------------------------------------------- slide 26
function slide26 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g30,-.01,-.02,6.37,7.54,{f:{color:NAVY}});

  tx(s,1.67,5.72,4.31,.61,[['WEAKNESS',{face:'Poppins SemiBold',sz:30,c:WHITE}]],{a:'center'});
  tx(s,7.29,1.56,5.01,1.14,[['Weakness Business Analysis Slide',{face:'Poppins SemiBold',sz:31,c:NAVY}]],{});
  tx(s,7.29,3.4,4.81,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  ln(s,7.38,3.05,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,7.41,5.58,1.12,.4,[['279.000',{face:'Poppins Medium',sz:15,c:'32425B'}]],{v:'middle',a:'center'});
  sh(s,7.38,4.43,1.12,1.03,{shape:'ellipse',f:{color:GOLD}});
  sh(s,10.87,4.43,1.12,1.03,{shape:'ellipse',f:{color:GOLD}});
  sh(s,9.1,4.43,1.12,1.03,{shape:'ellipse',f:{color:OFFWHITE}});
  tx(s,9.12,5.58,1.16,.4,[['500.000',{face:'Poppins Medium',sz:15,c:'32425B'}]],{v:'middle',a:'center'});
  tx(s,10.87,5.58,1.15,.4,[['999.450',{face:'Poppins Medium',sz:15,c:'32425B'}]],{v:'middle',a:'center'});
  poly(s,g3,9.55,4.68,.2,.19,{f:{color:NAVY}});
  poly(s,g4,9.37,4.61,.57,.66,{f:{color:NAVY}});
  poly(s,g8,7.83,4.76,.13,.13,{f:{color:NAVY}});
  poly(s,g9,7.79,4.9,.22,.11,{f:{color:NAVY}});
  poly(s,g10,7.62,4.61,.65,.65,{f:{color:NAVY}});
  poly(s,g11,8.06,5.05,.13,.13,{f:{color:NAVY}});
  poly(s,g5,11.21,4.63,.43,.63,{f:{color:NAVY}});
  poly(s,g6,11.29,4.7,.28,.28,{f:{color:NAVY}});
  poly(s,g7,11.33,4.75,.2,.19,{f:{color:NAVY}});
  poly(s,g52,1.3,1.2,5.06,4.12,{f:PH1,l:{color:GOLD,width:3}});
}

// ---------------------------------------------------------------- slide 27
function slide27 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g30,-.01,-.02,6.37,7.54,{f:{color:NAVY}});

  tx(s,1.67,5.72,4.31,.61,[['OPPORTUNITY',{face:'Poppins SemiBold',sz:30,c:WHITE}]],{a:'center'});
  tx(s,7.29,1.56,5.34,1.14,[['Opportunity Business Analysis Slide',{face:'Poppins SemiBold',sz:31,c:NAVY}]],{});
  ln(s,7.38,3.05,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,8.18,3.54,3.9,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  sh(s,7.4,3.44,.62,.63,{shape:'ellipse',f:{color:GOLD},rot:-90});
  sh(s,7.4,4.93,.65,.66,{shape:'ellipse',f:{color:GOLD},rot:-90});
  tx(s,8.18,3.3,2.01,.32,[['Creative Innovation',{face:'Poppins Medium',sz:13,c:DARK}]],{});
  tx(s,8.18,5.01,3.9,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,8.18,4.77,2.1,.32,[['Professional Solution',{face:'Poppins Medium',sz:13,c:DARK}]],{});
  poly(s,g3,7.67,5.12,.11,.11,{f:{color:NAVY}});
  poly(s,g4,7.57,5.08,.31,.35,{f:{color:NAVY}});
  poly(s,g12,7.61,3.62,.21,.33,{f:{color:NAVY}});
  poly(s,g13,7.71,3.7,.01,.03,{f:{color:NAVY}});
  poly(s,g14,7.67,3.72,.02,.03,{f:{color:NAVY}});
  poly(s,g15,7.73,3.72,.02,.03,{f:{color:NAVY}});
  poly(s,g16,7.51,3.56,.4,.21,{f:{color:NAVY}});
  poly(s,g53,1.57,1.22,4.39,4.06,{f:PH1,l:{color:GOLD,width:3}});
}

// ---------------------------------------------------------------- slide 28
function slide28 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g30,-.01,-.02,6.37,7.54,{f:{color:NAVY}});

  tx(s,1.67,5.72,4.31,.61,[['THREATS',{face:'Poppins SemiBold',sz:30,c:WHITE}]],{a:'center'});
  tx(s,7.29,1.56,5.01,1.14,[['Threats Business Analysis Slide',{face:'Poppins SemiBold',sz:31,c:NAVY}]],{});
  ln(s,7.38,3.05,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,7.29,3.43,4.22,.34,[['Our Threats Analysis Report',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,7.29,3.8,4.82,2.01,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  poly(s,g54,1.94,1.21,3.79,4.1,{f:PH1,l:{color:GOLD,width:3}});
}

// ---------------------------------------------------------------- slide 29
function slide29 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,0,0,4.93,7.5,{f:{color:NAVY}});

  img(s,6.71,1.47,4.84,2.9,{f:PH1});
  laptop(s,5.97,1.34,6.32,3.42);
  tx(s,5.97,3.8,3.76,.76,[['The Agency Business',{face:'Poppins Medium',sz:18,c:NAVY}]],{v:'middle',a:'center',shape:'round2DiagRect',f:{color:GOLD}});
  tx(s,1.2,1.8,2.91,1.51,[['Learn to Develop Your Business',{face:'Poppins SemiBold',sz:28,c:OFFWHITE}]],{});
  ln(s,1.32,3.62,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.2,3.97,2.99,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'justify',ls:1.5});
  tx(s,7.38,5.08,3.5,.37,[['CREATIVE INNOVATION',{face:'Poppins SemiBold',sz:16,c:NAVY}]],{a:'center'});
  ln(s,5.69,5.27,1.34,0,{l:{color:GOLD,width:2.25}});
  ln(s,11.19,5.27,1.34,0,{l:{color:GOLD,width:2.25}});
  tx(s,6.24,5.65,5.79,.62,[[' “ Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy. “',{face:'Lato',sz:11,c:GRAY,i:1}]],{a:'center',ls:1.5});
}

// ---------------------------------------------------------------- slide 30
function slide30 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,.54,.54,12.24,6.38,{l:{color:OFFWHITE,width:2}});
  sh(s,.26,1.06,.56,5.38,{f:{color:WHITE}});
  sh(s,1.23,-.01,5.34,7.51,{f:{color:NAVY}});

  img(s,2.33,1.16,3.03,5.17,{f:PH1,shape:'roundRect'});
  img(s,7.5,3.05,1.42,1.41,{f:PH1,shape:'ellipse'});
  img(s,7.5,4.91,1.42,1.41,{f:PH1,shape:'ellipse'});
  phone(s,2.21,1.1,3.26,5.39);
  tx(s,7.5,1.15,4.8,1.04,[['Grow Up Your Business Strategy',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{});
  ln(s,7.63,2.47,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,9.22,3.45,2.88,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,9.22,3.16,2.52,.32,[['Developing Visilibities',{face:'Poppins Medium',sz:13,c:DARK}]],{});
  tx(s,9.22,5.31,2.88,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,9.22,5.02,2.18,.32,[['Market Provit Issue',{face:'Poppins Medium',sz:13,c:DARK}]],{});
  tx(s,-.2,5.42,1.48,.29,[['INNOVATIVE ',{face:'Poppins Medium',sz:11,c:NAVY}]],{a:'center',rot:-90});
  sh(s,.47,4.64,.14,.13,{shape:'diamond',f:{color:GOLD}});
  tx(s,-.2,3.7,1.48,.29,[['PROFESSIONAL',{face:'Poppins Medium',sz:11,c:NAVY}]],{a:'center',rot:-90});
  sh(s,.47,2.73,.14,.13,{shape:'diamond',f:{color:GOLD}});
  tx(s,-.2,1.8,1.48,.29,[['EXPERIENCE',{face:'Poppins Medium',sz:11,c:NAVY}]],{a:'center',rot:-90});
}

// ---------------------------------------------------------------- slide 31
function slide31 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,9.29,0,4.06,7.52,{f:{color:NAVY}});

  img(s,6.78,1.66,5.08,2.85,{f:PH1});
  tx(s,7.96,3.58,2.69,.35,[['Creative Innovation',{face:'Poppins Medium',sz:15,c:NAVY}]],{a:'center'});
  poly(s,g12,8.99,2.45,.57,.92,{f:{color:NAVY}});
  poly(s,g13,9.26,2.67,.03,.08,{f:{color:NAVY}});
  poly(s,g14,9.16,2.7,.06,.08,{f:{color:NAVY}});
  poly(s,g15,9.32,2.7,.06,.08,{f:{color:NAVY}});
  poly(s,g16,8.72,2.26,1.1,.57,{f:{color:NAVY}});
  monitor(s,6.47,1.36,5.64,4.78);
  tx(s,1.27,1.65,4.45,1.04,[['Business Marketing and Innovation ',{face:'Poppins SemiBold',sz:28,c:NAVY}]],{});
  ln(s,1.38,2.99,1.34,0,{l:{color:GOLD,width:3}});
  tx(s,1.27,3.72,4.21,.9,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.27,3.38,2.63,.34,[['Our Great Mockup',{face:'Poppins Medium',sz:14,c:DARK}]],{});
  tx(s,1.28,5.02,1.14,.57,[['1174+',{face:'Poppins SemiBold',sz:28,c:DARK}]],{v:'middle'});
  tx(s,2.78,5.02,1.33,.57,[['2367+',{face:'Poppins SemiBold',sz:28,c:DARK}]],{v:'middle'});
  tx(s,4.44,5.01,1.28,.57,[['4919+',{face:'Poppins SemiBold',sz:28,c:DARK}]],{v:'middle'});
  tx(s,1.28,5.6,.99,.25,[['ON PROGRES',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
  tx(s,2.77,5.6,1.18,.25,[['BIGGEST CLIENT',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
  tx(s,4.44,5.6,1.12,.25,[['PROJECT GOAL',{face:'Lato',sz:9,c:GRAY}]],{v:'middle'});
}

// ---------------------------------------------------------------- slide 32
function slide32 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  tx(s,6.2,2.41,5.81,.47,[['Creative Solution',{face:'Poppins SemiBold',sz:22,c:DARK}]],{v:'bottom',m0:1});
  tx(s,6.2,2.88,5.81,.67,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:12,c:GRAY}]],{a:'justify',ls:1.5,m0:1});
  tx(s,6.2,3.82,5.81,.47,[['Strategical Business',{face:'Poppins SemiBold',sz:22,c:DARK}]],{v:'bottom',m0:1});
  tx(s,6.2,4.29,5.81,.67,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:12,c:GRAY}]],{a:'justify',ls:1.5,m0:1});
  tx(s,6.2,5.24,5.81,.47,[['Market Profitable',{face:'Poppins SemiBold',sz:22,c:DARK}]],{v:'bottom',m0:1});
  tx(s,6.2,5.7,5.81,.67,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:12,c:GRAY}]],{a:'justify',ls:1.5,m0:1});
  ln(s,10.74,1.46,2.58,0,{l:{color:GOLD,width:5}});
  tx(s,6.08,1.27,4.39,.71,[['Our Infographic',{face:'Poppins SemiBold',sz:36,c:NAVY}]],{});
  tx(s,6.1,.96,2.65,.34,[['Business Strategy',{face:'Poppins Medium',sz:14,c:GRAY}]],{});
  poly(s,g55,1.6,5.83,1.75,.71,{f:{color:GOLD}});
  poly(s,g56,1.31,1.27,3.88,2.26,{f:{color:NAVY}});
  poly(s,g57,1.6,3.21,3.59,1.61,{f:{color:GOLD}});
  poly(s,g58,1.6,4.52,3.59,1.62,{f:{color:NAVY}});
  tx(s,1.6,2.69,1.95,.4,[['Professional',{face:'Montserrat SemiBold',sz:18,c:DARK}]],{v:'middle',a:'center',m0:1});
  tx(s,1.6,4.01,1.95,.4,[['Experience',{face:'Montserrat SemiBold',sz:18,c:DARK}]],{v:'middle',a:'center',m0:1});
  tx(s,1.6,5.34,1.95,.4,[['Management',{face:'Montserrat SemiBold',sz:18,c:DARK}]],{v:'middle',a:'center',m0:1});
  poly(s,g5,2.3,5.92,.34,.5,{f:{color:NAVY}});
  poly(s,g6,2.36,5.98,.22,.22,{f:{color:NAVY}});
  poly(s,g7,2.39,6.02,.16,.15,{f:{color:NAVY}});
  poly(s,g3,2.35,1.72,.24,.23,{f:{color:WHITE}});
  poly(s,g4,2.13,1.64,.69,.78,{f:{color:WHITE}});
  poly(s,g8,2.41,3.43,.11,.11,{f:{color:NAVY}});
  poly(s,g9,2.37,3.55,.18,.09,{f:{color:NAVY}});
  poly(s,g10,2.23,3.31,.53,.53,{f:{color:NAVY}});
  poly(s,g11,2.59,3.67,.11,.11,{f:{color:NAVY}});
  poly(s,g17,2.2,4.6,.54,.54,{f:{color:WHITE}});
  poly(s,g18,2.43,4.98,.05,.05,{f:{color:WHITE}});
  poly(s,g19,2.25,4.7,.22,.22,{f:{color:WHITE}});
  poly(s,g20,2.31,4.76,.09,.09,{f:{color:WHITE}});
  poly(s,g2,2.44,4.91,.03,.02,{f:{color:WHITE}});
  poly(s,g2,2.49,4.91,.03,.02,{f:{color:WHITE}});
  poly(s,g21,2.53,4.91,.03,.02,{f:{color:WHITE}});
  poly(s,g22,2.58,4.91,.03,.02,{f:{color:WHITE}});
  poly(s,g23,2.62,4.91,.03,.02,{f:{color:WHITE}});
  poly(s,g24,2.57,4.64,.11,.17,{f:{color:WHITE}});
}

// ---------------------------------------------------------------- slide 33
function slide33 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  tx(s,3.87,.96,5.6,.81,[['The Best Infographics Here',{face:'Poppins SemiBold',sz:28,c:'1F3864'}]],{a:'center',ls:.9});
  tx(s,1.31,5.94,2.46,.59,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5,m0:1});
  tx(s,9.58,2.64,2.46,.59,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5,m0:1});
  tx(s,1.39,2.71,2.34,.44,[['Professional',{face:'Poppins Medium',sz:20,c:DARK}]],{v:'bottom',a:'center',m0:1});
  tx(s,6.86,2.71,2.34,.44,[['Experience',{face:'Poppins Medium',sz:20,c:DARK}]],{v:'bottom',a:'center',m0:1});
  tx(s,4.05,6.01,2.34,.44,[['Marketing',{face:'Poppins Medium',sz:20,c:DARK}]],{v:'bottom',a:'center',m0:1});
  tx(s,9.68,6.01,2.34,.44,[['Management',{face:'Poppins Medium',sz:20,c:DARK}]],{v:'bottom',a:'center',m0:1});
  tx(s,6.87,5.94,2.46,.59,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5,m0:1});
  tx(s,4.03,2.64,2.46,.59,[['Lorem Ipsum is simply dummy text of the printing typesetting industry.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5,m0:1});
  poly(s,g59,1.42,4.74,2.16,.91,{f:{color:GOLD}});
  poly(s,g60,4.18,3.44,2.16,.91,{f:{color:GOLD}});
  poly(s,g61,6.94,4.74,2.16,.91,{f:{color:GOLD}});
  poly(s,g62,9.71,3.44,2.16,.91,{f:{color:GOLD}});
  poly(s,g63,1.4,3.44,2.2,1.17,{f:{color:NAVY}});
  poly(s,g64,4.16,4.48,2.2,1.17,{f:{color:NAVY}});
  poly(s,g65,6.93,3.44,2.2,1.17,{f:{color:NAVY}});
  poly(s,g66,9.69,4.48,2.2,1.17,{f:{color:NAVY}});
  sh(s,3.55,4.48,.66,.13,{f:{color:NAVY}});
  sh(s,6.31,4.48,.66,.13,{f:{color:NAVY}});
  sh(s,9.08,4.48,.66,.13,{f:{color:NAVY}});
  poly(s,g5,10.54,4.17,.5,.74,{f:{color:DARK}});
  poly(s,g6,10.63,4.26,.33,.33,{f:{color:DARK}});
  poly(s,g7,10.68,4.31,.23,.22,{f:{color:DARK}});
  poly(s,g3,2.38,4.24,.23,.22,{f:{color:DARK}});
  poly(s,g4,2.17,4.16,.66,.76,{f:{color:DARK}});
  poly(s,g8,7.9,4.34,.15,.15,{f:{color:DARK}});
  poly(s,g9,7.85,4.5,.25,.13,{f:{color:DARK}});
  poly(s,g10,7.65,4.16,.75,.75,{f:{color:DARK}});
  poly(s,g11,8.16,4.67,.15,.15,{f:{color:DARK}});
  poly(s,g17,4.86,4.14,.81,.81,{f:{color:DARK}});
  poly(s,g18,5.2,4.7,.07,.07,{f:{color:DARK}});
  poly(s,g19,4.93,4.29,.33,.33,{f:{color:DARK}});
  poly(s,g20,5.02,4.39,.14,.14,{f:{color:DARK}});
  poly(s,g2,5.22,4.6,.05,.02,{f:{color:DARK}});
  poly(s,g2,5.29,4.6,.05,.02,{f:{color:DARK}});
  poly(s,g21,5.35,4.6,.05,.02,{f:{color:DARK}});
  poly(s,g22,5.42,4.6,.05,.02,{f:{color:DARK}});
  poly(s,g23,5.49,4.6,.05,.02,{f:{color:DARK}});
  poly(s,g24,5.4,4.2,.16,.25,{f:{color:DARK}});
  ln(s,5.86,1.85,1.34,0,{l:{color:GOLD,width:3}});
}

// ---------------------------------------------------------------- slide 34
function slide34 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g32,6.3,2.38,.72,.31,{f:{color:NAVY}});
  poly(s,g32,7.8,3.71,.72,.31,{f:{color:NAVY}});
  poly(s,g32,5.15,5.04,.72,.31,{f:{color:NAVY}});
  sh(s,2.71,1.34,4.86,1.06,{shape:'roundRect',f:{color:NAVY},rr:.53});
  sh(s,4.57,4,4.86,1.06,{shape:'roundRect',f:{color:NAVY},rr:.53});
  sh(s,1.8,5.34,4.86,1.06,{shape:'roundRect',f:{color:NAVY},rr:.53});
  sh(s,5.76,2.67,4.86,1.06,{shape:'roundRect',f:{color:NAVY},rr:.53});
  sh(s,1.34,1.34,1.06,1.06,{shape:'ellipse',f:{color:GOLD}});
  poly(s,g5,1.66,1.56,.42,.62,{f:{color:DARK}});
  poly(s,g6,1.74,1.63,.27,.27,{f:{color:DARK}});
  poly(s,g7,1.78,1.67,.19,.18,{f:{color:DARK}});
  sh(s,6.96,5.34,1.06,1.06,{shape:'ellipse',f:{color:GOLD}});
  poly(s,g3,7.4,5.63,.18,.17,{f:{color:DARK}});
  poly(s,g4,7.24,5.57,.52,.59,{f:{color:DARK}});
  sh(s,10.93,2.67,1.06,1.06,{shape:'ellipse',f:{color:GOLD}});
  poly(s,g8,11.36,3.03,.12,.13,{f:{color:DARK}});
  poly(s,g9,11.31,3.16,.21,.11,{f:{color:DARK}});
  poly(s,g10,11.15,2.89,.62,.62,{f:{color:DARK}});
  poly(s,g11,11.57,3.31,.13,.13,{f:{color:DARK}});
  sh(s,3.21,4,1.06,1.06,{shape:'ellipse',f:{color:GOLD}});
  poly(s,g17,3.44,4.23,.6,.6,{f:{color:DARK}});
  poly(s,g18,3.69,4.65,.05,.05,{f:{color:DARK}});
  poly(s,g19,3.49,4.35,.25,.25,{f:{color:DARK}});
  poly(s,g20,3.56,4.42,.1,.1,{f:{color:DARK}});
  poly(s,g2,3.7,4.58,.04,.02,{f:{color:DARK}});
  poly(s,g2,3.75,4.58,.04,.02,{f:{color:DARK}});
  poly(s,g21,3.8,4.58,.04,.02,{f:{color:DARK}});
  poly(s,g22,3.85,4.58,.04,.02,{f:{color:DARK}});
  poly(s,g23,3.9,4.58,.04,.02,{f:{color:DARK}});
  poly(s,g24,3.84,4.28,.12,.18,{f:{color:DARK}});
  tx(s,3.17,1.45,3.94,.39,[['Endpoint Management',{face:'Poppins Medium',sz:17,c:WHITE}]],{v:'bottom',m0:1});
  tx(s,3.17,1.8,3.94,.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:10.5,c:LTGRAY}]],{a:'justify',m0:1});
  tx(s,6.3,2.77,3.94,.39,[['18',{face:'Poppins Medium',sz:17,c:WHITE}],['th',{face:'Poppins Medium',sz:17,c:WHITE,sup:1}],['  Years Experience ',{face:'Poppins Medium',sz:17,c:WHITE}]],{v:'bottom',m0:1});
  tx(s,6.3,3.12,3.94,.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:10.5,c:LTGRAY}]],{a:'justify',m0:1});
  tx(s,5.15,4.13,3.94,.39,[['Marketing Business',{face:'Poppins Medium',sz:17,c:WHITE}]],{v:'bottom',m0:1});
  tx(s,5.15,4.47,3.94,.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:10.5,c:LTGRAY}]],{a:'justify',m0:1});
  tx(s,2.29,5.48,3.94,.39,[['Professional Solution',{face:'Poppins Medium',sz:17,c:WHITE}]],{v:'bottom',m0:1});
  tx(s,2.29,5.83,3.94,.45,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry’s.',{face:'Lato',sz:10.5,c:LTGRAY}]],{a:'justify',m0:1});
  ln(s,-.02,5.87,1.23,0,{l:{color:GOLD,width:3.5}});
  tx(s,8.1,1.42,4.02,.62,[['Infographic Slide',{face:'Poppins SemiBold',sz:31,c:NAVY}]],{a:'right'});
  tx(s,9.46,1.11,2.65,.34,[['Business Marketing',{face:'Poppins Medium',sz:14,c:GRAY}]],{a:'right'});
  ln(s,12.35,1.67,.97,0,{l:{color:GOLD,width:3.5}});
}

// ---------------------------------------------------------------- slide 35
function slide35 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  img(s,-.02,.01,6.91,7.48,{f:PH1});
  sh(s,-.02,0,6.92,7.5,{f:{color:NAVY,transparency:12.16}});
  columnChart(s,1.21,3.33,4.83,3.37);
  tx(s,1.18,.8,5,1.18,[['Our Business Chart Analysis',{face:'Poppins SemiBold',sz:32,c:WHITE}]],{});
  ln(s,1.3,2.24,1.04,0,{l:{color:GOLD,width:2.25}});
  tx(s,1.21,2.52,4.8,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy.',{face:'Lato',sz:11,c:LTGRAY}]],{a:'justify',ls:1.5});
  tx(s,8.13,5,3.99,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  pieChart(s,7.67,1.38,4.83,3.42);
  tx(s,8.9,.8,2.35,.44,[['Our Pie Chart',{face:'Poppins SemiBold',sz:20,c:'000000'}]],{a:'center'});
}

// ---------------------------------------------------------------- slide 36
function slide36 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  barChart(s,.95,.86,6.68,5.78);
  ln(s,12.03,1.55,1.29,0,{l:{color:GOLD,width:5}});
  tx(s,8.08,1.27,3.96,.64,[['Our Chart Style',{face:'Poppins SemiBold',sz:32,c:NAVY}]],{});
  tx(s,8.11,.96,2.65,.3,[['Monthly Data Report',{face:'Poppins Medium',sz:12,c:GRAY}]],{});
  tx(s,8.2,2.43,2.43,.4,[['October 2020',{face:'Poppins Medium',sz:18,c:DARK}]],{v:'bottom',m0:1});
  tx(s,8.2,2.86,3.81,1.27,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer.',{face:'Lato',sz:12,c:GRAY}]],{a:'justify',ls:1.5,m0:1});
  tx(s,8.2,4.56,2.43,.4,[['November 2020',{face:'Poppins Medium',sz:18,c:DARK}]],{v:'bottom',m0:1});
  tx(s,8.2,4.96,3.81,1.27,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer.',{face:'Lato',sz:12,c:GRAY}]],{a:'justify',ls:1.5,m0:1});
}

// ---------------------------------------------------------------- slide 37
function slide37 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  sh(s,5.15,2.21,2.98,4.34,{f:{color:NAVY}});
  sh(s,1.56,2.21,2.98,4.34,{f:{color:OFFWHITE}});
  sh(s,8.86,2.21,2.98,4.34,{f:{color:OFFWHITE}});
  tx(s,2.73,.88,7.87,.54,[['Choose Your Pricing Plan',{face:'Poppins SemiBold',sz:26,c:'1F3864'}]],{a:'center'});
  tx(s,3.66,1.42,6.11,.34,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry lorem Ipsum.',{face:'Lato',sz:11,c:GRAY}]],{a:'center',ls:1.5});
  tx(s,9.39,5.9,1.77,.46,[['Get Started',{face:'Poppins Medium',sz:14,c:DARK}]],{v:'middle',a:'center',shape:'roundRect',f:{color:'ECECEC'},rr:.08});
  tx(s,5.75,5.9,1.77,.46,[['Get Started',{face:'Poppins Medium',sz:14,c:NAVY}]],{v:'middle',a:'center',shape:'roundRect',f:{color:GOLD},rr:.08});
  tx(s,5.84,2.39,1.6,.34,[['Standard Plan',{face:'Poppins Medium',sz:14,c:GOLD}]],{a:'center'});
  tx(s,5.86,2.79,1.56,.94,[['$350 ',{face:'Poppins SemiBold',sz:36,c:WHITE,br:1}],['/month ',{face:'Poppins Light',sz:14,c:WHITE}]],{a:'center'});
  ln(s,5.4,3.85,2.48,0,{l:{color:'ECECEC',width:.75}});
  poly(s,g1,5.44,4.2,.27,.22,{f:{color:OFFWHITE}});
  tx(s,5.78,4.14,1.41,.32,[['Unlimited User',{face:'Lato',sz:13,c:WHITE}]],{});
  poly(s,g1,5.44,4.6,.27,.22,{f:{color:OFFWHITE}});
  tx(s,5.77,4.54,1.54,.32,[['High Quality File',{face:'Lato',sz:13,c:WHITE}]],{});
  poly(s,g1,5.44,5,.27,.22,{f:{color:OFFWHITE}});
  tx(s,5.77,4.95,1.89,.32,[['Free Costum Domain',{face:'Lato',sz:13,c:WHITE}]],{});
  poly(s,g1,5.44,5.41,.27,.22,{f:{color:OFFWHITE}});
  tx(s,5.77,5.35,1.88,.32,[['Outstanding Support',{face:'Lato',sz:13,c:WHITE}]],{});
  tx(s,2.11,5.9,1.77,.46,[['Get Started',{face:'Poppins Medium',sz:14,c:DARK}]],{v:'middle',a:'center',shape:'roundRect',f:{color:'ECECEC'},rr:.08});
  tx(s,2.43,2.4,1.2,.34,[['Basic Plan',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,2.33,2.8,1.44,.94,[['$250',{face:'Poppins SemiBold',sz:36,c:NAVY,br:1}],['/month',{face:'Poppins Light',sz:14,c:NAVY}]],{a:'center'});
  ln(s,1.81,3.85,2.48,0,{l:{color:LTGRAY,width:.75}});
  poly(s,g1,1.88,4.2,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,2.22,4.14,1.41,.32,[['Unlimited User',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,1.88,4.6,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,2.22,4.54,1.54,.32,[['High Quality File',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,1.88,5,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,2.22,4.95,1.89,.32,[['Free Costum Domain',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,1.88,5.41,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,2.22,5.35,1.88,.32,[['Outstanding Support',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  tx(s,9.58,2.38,1.56,.34,[['Advance Plan',{face:'Poppins Medium',sz:14,c:DARK}]],{a:'center'});
  tx(s,9.62,2.8,1.48,.94,[['$450',{face:'Poppins SemiBold',sz:36,c:NAVY,br:1}],['/month',{face:'Poppins Light',sz:14,c:NAVY}]],{a:'center'});
  ln(s,9.11,3.85,2.48,0,{l:{color:LTGRAY,width:.75}});
  poly(s,g1,9.17,4.2,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,9.51,4.14,1.41,.32,[['Unlimited User',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,9.17,4.6,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,9.51,4.54,1.54,.32,[['High Quality File',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,9.17,5,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,9.51,4.95,1.89,.32,[['Free Costum Domain',{face:'Lato',sz:13,c:MIDGRAY}]],{});
  poly(s,g1,9.17,5.41,.27,.22,{f:{color:'AFA6A6'}});
  tx(s,9.51,5.35,1.88,.32,[['Outstanding Support',{face:'Lato',sz:13,c:MIDGRAY}]],{});
}

// ---------------------------------------------------------------- slide 38
function slide38 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  tx(s,1.21,4.95,3.7,1.73,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.  It has survived not only five centuries.',{face:'Lato',sz:11,c:GRAY}]],{a:'justify',ls:1.5});
  tx(s,1.21,4.51,3.84,.4,[['Our Information Centre',{face:'Poppins Medium',sz:18,c:DARK}]],{});
  sh(s,5.56,5.02,.5,.5,{shape:'ellipse',f:{color:NAVY}});
  poly(s,g67,5.73,5.15,.16,.24,{f:{color:WHITE}});
  tx(s,6.24,4.95,2.82,.71,[['5678 Budiono Street , IN 019  Jakarta, Indonesia',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  sh(s,5.56,5.93,.5,.5,{shape:'ellipse',f:{color:NAVY}});
  poly(s,g68,5.73,6.08,.06,.06,{f:{color:WHITE}});
  poly(s,g69,5.71,6.09,.19,.19,{f:{color:WHITE}});
  poly(s,g70,5.86,6.2,.06,.06,{f:{color:WHITE}});
  tx(s,6.23,5.9,2.48,.71,[['(0291) 9218-7231',{face:'Lato',sz:13,c:GRAY,br:1}],['082-384-834-117',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,5.44,4.51,1.95,.41,[['Contact Us',{face:'Poppins Medium',sz:18,c:DARK}]],{});
  tx(s,9.18,4.51,1.95,.41,[['Office Hours',{face:'Poppins Medium',sz:18,c:DARK}]],{});
  tx(s,9.19,4.95,1.68,.39,[['Monday - Thusday',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,9.19,5.57,1.68,.39,[['Friday - Saturday',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,9.18,6.18,1.68,.39,[['Sunday',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,10.88,4.95,1.32,.39,[['8 am – 15 pm',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,10.87,5.57,1.32,.39,[['8 am – 12 pm',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  tx(s,10.85,6.18,1.32,.39,[['Closed',{face:'Lato',sz:13,c:GRAY}]],{ls:1.5});
  img(s,0,0,13.33,3.75,{f:PH1});
}

// ---------------------------------------------------------------- slide 39
function slide39 (p) {
  const s = p.addSlide();
  s.background = { color: WHITE };
  poly(s,g72,3.7,0,9.64,7.52,{f:{color:NAVY}});

  tx(s,8.13,2.13,3.96,3.31,[['There are no secrets to success. It is the result of preparation, hard work, and learning from failure',{face:'Merriweather',sz:26,c:OFFWHITE,i:1}]],{ls:1.5});
  tx(s,8.13,1.71,2.76,.42,[['QOUTES THE DAY',{face:'Poppins Medium',sz:14,c:GOLD}]],{ls:1.5});
  ln(s,8.21,5.79,1.42,0,{l:{color:GOLD,width:2.25}});
  poly(s,g71,0,0,7.08,7.52,{f:PH1,l:{color:WHITE,width:2.25}});
}

// ---------------------------------------------------------------- slide 40
function slide40 (p) {
  const s = p.addSlide();
  s.background = { color: NAVY };
  poly(s,g74,-.69,-1.81,11.27,9.64,{f:{color:OFFWHITE},rot:-163.02,fh:1});
  poly(s,g75,1.36,-1.62,9.84,8.59,{f:{color:GOLD},rot:-163.02,fh:1});

  tx(s,8.05,4.53,3.78,1.11,[['Thanks !',{face:'Montserrat SemiBold',sz:60,c:WHITE}]],{});
  tx(s,8.11,6.03,4.77,.62,[['Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry\'s standard dummy text.',{face:'Lato',sz:11,c:LTGRAY}]],{ls:1.5});
  tx(s,8.11,5.63,3.45,.34,[['For Watching Presentation',{face:'Poppins Medium',sz:14,c:OFFWHITE}]],{});
  poly(s,g12,11.97,1.07,.56,.9,{f:{color:GOLD}});
  poly(s,g13,12.23,1.29,.03,.08,{f:{color:GOLD}});
  poly(s,g14,12.14,1.32,.06,.07,{f:{color:GOLD}});
  poly(s,g15,12.3,1.32,.06,.07,{f:{color:GOLD}});
  poly(s,g16,11.71,.89,1.08,.56,{f:{color:GOLD}});
  poly(s,g73,-.02,0,11.36,7.34,{f:PH1});
}

// ---------------------------------------------------------------- build
const SLIDES = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.33, height: 7.5 });
pptx.layout = 'DECK';
pptx.title = 'The Agency Business';
SLIDES.forEach(fn => fn(pptx));

pptx.writeFile({ fileName: path.join(__dirname, '02cb9c24-748b-47f3-8414-913b8756dbb3_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
