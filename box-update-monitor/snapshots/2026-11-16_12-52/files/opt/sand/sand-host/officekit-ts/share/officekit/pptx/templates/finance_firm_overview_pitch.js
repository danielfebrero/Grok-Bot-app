/**
 * Recreation of the "Fiz Creative" finance deck (20 slides, 13.333 x 7.5 in)
 * using pptxgenjs only.  Raster photos in the original are replaced with
 * programmatic "[image]" placeholders per the conversion brief.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const C = {
  bg:      'F8F8F8', // slide background (master bgPr)
  white:   'FFFFFF', // theme accent2 / accent4
  green:   '67BD66', // theme accent1
  black:   '000000', // theme tx1
  body:    '595959', // tx1 lumMod 65% / lumOff 35% -> body copy grey
  mapGrey: 'D9D9D9', // bg1 lumMod 85% -> Spain map regions
};
const F = { head: 'Inter', body: 'Lato', light: 'Lato Light' };

// Soft drop shadow used by every "floating" white card (outerShdw blurRad 292100, 17% black)
const CARD_SHADOW = { type: 'outer', color: C.black, opacity: 0.17, blur: 23, offset: 0, angle: 90 };

/* --------------------------------------------------------------- helpers */
const card   = (s, x, y, w, h, o = {}) =>
  s.addShape('rect', { x, y, w, h, fill: { color: o.fill || C.white }, line: { type: 'none' },
                       shadow: o.shadow === false ? undefined : CARD_SHADOW });
const plain  = (s, x, y, w, h, fill) =>
  s.addShape('rect', { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });

// Text boxes in the source all use zero insets, top anchor and no autofit.
const TXT = { margin: 0, valign: 'top', isTextBox: true, wrap: true };

// Big display numerals / titles (Inter Bold, theme tx1)
const head = (s, txt, x, y, w, h, size, o = {}) =>
  s.addText(txt, { ...TXT, x, y, w, h, fontFace: F.head, fontSize: size, bold: true,
                   color: o.color || C.black, align: o.align || 'left',
                   lineSpacingMultiple: 0.9, ...(o.extra || {}) });

// Small bold sub-heads (Lato Bold 14pt unless overridden); `lh` picks 100% vs 150% leading
const label = (s, txt, x, y, w, h, o = {}) =>
  s.addText(txt, { ...TXT, x, y, w, h, fontFace: F.body, fontSize: o.size || 14, bold: true,
                   color: o.color || C.black, align: o.align || 'left',
                   lineSpacingMultiple: o.lh || 1 });

// Grey italic Lato Light body copy, 150% line spacing
const body = (s, txt, x, y, w, h, o = {}) =>
  s.addText(txt, { ...TXT, x, y, w, h, fontFace: F.light, fontSize: o.size || 10, italic: true,
                   color: o.color || C.body, align: o.align || 'left',
                   lineSpacingMultiple: 1.5, ...(o.extra || {}) });

// Bulleted grey list (used by the "Goals" card on slide 5); 10pt space between items
const bullets = (s, items, x, y, w, h) =>
  s.addText(items.map(t => ({ text: t, options: { bullet: { characterCode: '2022' }, breakLine: true } })),
            { ...TXT, x, y, w, h, fontFace: F.light, fontSize: 10, italic: true, color: C.body,
              lineSpacingMultiple: 1.5, paraSpaceBefore: 10, indentLevel: 0 });

// Filled outline icon drawn from normalised polygon data inside a green tile.
const drawIcon = (s, glyph, x, y, w, h, color) => {
  const pts = [];
  glyph.forEach(sub => {
    for (let i = 0; i < sub.length; i += 2) {
      pts.push({ x: (sub[i] / 1000) * w, y: (sub[i + 1] / 1000) * h, moveTo: i === 0 });
    }
    pts.push({ close: true });
  });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: color || C.white }, line: { type: 'none' } });
};

// Green tile + centred white icon: the deck's signature motif.
const iconTile = (s, glyph, tx, ty, tw, th, ix, iy, iw, ih) => {
  plain(s, tx, ty, tw, th, C.green);
  drawIcon(s, glyph, ix, iy, iw, ih);
};

/**
 * Placeholder standing in for a photo in the reference deck: a light frame with
 * a mountain/sun pictogram and an "[image]" caption, centred on the original
 * picture placeholder's box (the source renders an "Image Here" prompt there).
 */
const imagePlaceholder = (s, px, py, pw, ph) => {
  const w = 0.875, h = 0.685;                       // frame size, as the source renders it
  const x = px + (pw - w) / 2, y = py + (ph - h) / 2;
  s.addShape('rect', { x, y, w, h, fill: { color: C.white }, line: { color: '4D4D4D', width: 1 } });
  s.addShape('rect', { x: x + 0.085, y: y + 0.05, w: w - 0.170, h: h - 0.105,
                       fill: { type: 'none' }, line: { color: '808080', width: 0.75 } });
  s.addShape('triangle', { x: x + 0.235, y: y + 0.260, w: 0.545, h: 0.355,
                           fill: { color: '83BEEC' }, line: { type: 'none' } });
  s.addShape('triangle', { x: x + 0.095, y: y + 0.445, w: 0.230, h: 0.170,
                           fill: { color: '83BEEC' }, line: { type: 'none' } });
  s.addShape('ellipse', { x: x + 0.154, y: y + 0.123, w: 0.136, h: 0.136,
                          fill: { color: 'F7C873' }, line: { type: 'none' } });
  s.addText('[image]', { x: x - 0.2, y: y + 0.24, w: w + 0.4, h: 0.2, fontFace: F.head, fontSize: 13,
                         color: '1A1A1A', align: 'center', margin: 0, valign: 'middle' });
};

const OUT_NAME = '18484cab-4cdc-49eb-8955-66f481df8c7c_grok_final.pptx';

/* ------------------------------------------- vector icon path data
 * Each glyph is a list of closed sub-paths; numbers are x,y pairs in
 * per-mille of the glyph box, so a glyph can be scaled to any size.
 */
const ICONS = {
  money: [
  [971,0,1000,27,999,458,971,477,821,477,812,597,716,721,762,767,755,969,
   696,1000,296,1000,248,982,229,938,229,781,244,741,283,719,248,630,256,573,
   291,523,285,477,29,477,4,464,4,13,29,0,971,0],
  [347,477,343,500,365,514,423,477],
  [580,335,538,310,571,282,584,223,527,164,461,168,428,196,415,254,466,312,
   429,335,228,334,201,292,155,267,155,210,201,185,228,143,772,143,799,185,
   845,210,845,267,799,292,772,334,580,335],
  [500,265,528,239,500,213,472,239,500,265],
  [706,781,294,772,287,938,706,938],
  [767,512,762,477,683,477,684,512,644,518,545,377,488,361,467,386,528,500,
   525,574,478,574,471,510,395,563,330,560,310,592,310,648,334,693,368,711,
   639,718,748,601,767,512],
  [942,54,58,54,58,424,420,424,410,388,195,387,160,330,101,303,98,180,
   160,147,188,93,797,89,822,102,840,147,899,174,899,303,840,330,805,387,
   628,388,653,424,942,424],
  [651,864,607,894,574,854,617,824,651,864]],
  coins: [
  [796,243,704,252,704,91,695,63,640,23,450,3,327,41,296,91,296,252,
   106,253,8,306,8,871,106,924,299,924,370,981,500,1000,630,981,701,924,
   847,931,970,893,1000,843,1000,334,992,306,936,266,796,243],
  [935,594,811,616,782,729,872,723,935,695],
  [500,829,435,821,326,766,263,673,254,617,263,560,326,467,435,412,500,405,
   565,413,674,467,737,560,746,617,737,673,674,766,565,821,500,829],
  [541,351,638,377,638,328,541,351],
  [458,351,362,328,362,377,458,351],
  [638,177,500,203,362,177,362,251,399,281,500,297,601,281,638,251],
  [935,420,844,444,735,441,804,560,935,525],
  [796,299,704,310,704,376,837,388,935,345,899,313,796,299],
  [362,101,399,130,500,147,601,130,638,101,603,70,500,56,362,91],
  [65,334,82,364,130,382,296,376,296,310,266,304,162,301,65,334],
  [65,420,65,525,157,557,196,560,265,441,156,444,65,420],
  [65,695,128,723,218,729,189,616,65,594],
  [65,843,128,870,251,875,296,866,296,819,256,783,65,763,65,843],
  [500,944,634,915,638,856,500,885,362,856,367,915,500,944],
  [796,878,930,848,935,763,744,783,704,819,704,866,796,878],
  [468,559,481,583,569,616,594,665,582,713,533,751,531,782,509,801,477,794,
   467,754,418,723,402,685,441,663,498,701,522,690,527,664,422,610,406,539,
   467,481,472,445,500,431,528,445,533,481,577,507,598,550,570,582,511,534,
   468,559]],
  target: [
  [990,571,846,516,848,395,821,283,768,183,693,100,599,39,489,5,375,3,
   268,32,173,87,94,166,37,266,5,381,0,465,35,626,123,763,184,818,
   119,993,229,997,288,873,425,897,563,873,622,997,730,995,737,975,667,818,
   777,701,940,755,903,644,994,599,990,571],
  [258,861,211,836,154,965,211,966],
  [697,965,640,836,593,861,640,966],
  [737,685,667,569,568,687,505,718,370,724,245,657,170,533,164,390,228,259,
   283,210,413,167,547,197,611,256,592,269,528,224,407,200,293,242,247,284,
   212,340,190,468,229,587,269,636,322,673,444,697,557,655,638,557,598,540,
   566,588,483,643,384,647,330,623,343,597,392,615,474,611,568,527,495,496,
   468,515,407,454,412,429,504,433,508,466,580,497,587,439,562,358,513,305,
   434,278,353,296,289,356,267,482,309,574,282,584,237,491,241,388,296,298,
   382,250,483,254,568,311,614,403,610,509,650,526,658,493,655,391,629,318,
   645,302,692,422,679,539,816,504,818,410,764,237,646,104,566,61,463,35,
   361,39,265,70,179,125,108,203,58,301,33,409,37,516,66,618,118,709,
   192,783,285,836,378,860,473,860,578,831,678,767,746,689],
  [943,588,797,532,715,571,754,657,899,713,866,638,875,620]],
  hand: [
  [784,577,650,594,587,648,587,929,650,983,878,992,952,969,982,929,974,626,
   919,594,784,577],
  [905,720,784,736,663,720,667,750,748,772,851,768,905,745],
  [663,831,718,854,821,858,902,836,905,809,784,825,663,809],
  [784,628,663,650,718,678,851,678,905,650,784,628],
  [784,948,899,931,905,895,784,911,663,895,663,927,784,948],
  [198,456,63,473,0,528,0,929,30,969,104,992,291,992,365,969,395,929,
   387,505,332,473,198,456],
  [319,693,198,709,77,693,80,731,162,752,264,748,318,725],
  [319,599,198,615,77,599,80,634,162,656,264,652,318,629],
  [77,822,132,845,234,849,315,827,318,790,197,806,77,790],
  [198,508,77,529,107,552,198,563,288,552,318,529,198,508],
  [198,948,312,931,318,886,197,902,77,886,83,931,198,948],
  [998,408,996,445,948,488,839,489,739,553,554,575,508,555,494,520,425,502,
   299,337,196,335,51,259,33,215,231,24,268,5,335,4,495,83,520,126,
   830,229,936,310,998,408],
  [441,131,298,52,112,217,260,293],
  [605,415,605,436,773,412,719,365,634,348,603,356,605,415],
  [683,513,771,465,711,465,570,524,683,513],
  [901,446,874,441,788,336,630,296,536,330,527,445,484,469,434,386,349,299,
   489,173,817,291,910,391,923,425,901,446],
  [265,219,226,244,188,219,226,193,265,219]],
  umbrella: [
  [465,0,588,16,699,62,793,132,866,224,913,331,929,451,916,474,881,474,
   844,396,787,373,730,396,703,465,672,477,612,382,571,372,514,396,492,451,
   488,922,437,985,382,1000,326,985,287,947,272,833,292,808,323,820,328,908,
   354,940,397,946,436,908,427,412,379,376,336,376,288,412,269,470,243,477,
   199,396,142,373,85,396,48,474,13,474,0,451,17,331,64,224,136,132,
   230,62,341,16,465,0],
  [787,321,859,341,755,169,643,92,469,52,363,65,266,102,181,163,114,243,
   70,341,166,323,250,373,308,330,376,322,465,373,523,330,591,322,680,373,
   727,334,787,321],
  [845,776,838,830,779,908,753,902,745,866,699,829,694,804,712,785,764,817,
   792,801,787,779,721,751,696,707,702,662,745,624,758,582,791,587,850,692,
   819,714,769,671,751,705,819,737,845,776],
  [772,966,711,958,610,901,552,803,544,744,552,685,610,587,711,530,772,522,
   832,530,933,587,992,685,1000,744,992,803,933,901,832,958,772,966],
  [772,575,895,625,946,744,895,864,772,913,649,864,598,744,649,625,772,575]],
  notes: [
  [1000,365,713,42,685,46,353,305,353,90,330,43,226,3,130,3,53,25,
   10,65,10,334,89,398,219,409,10,571,1,599,219,949,356,998,617,826,
   685,805,811,704,811,674,985,544,1000,365],
  [178,53,74,76,61,92,98,115,209,125,296,92,266,67,178,53],
  [60,193,118,238,210,244,281,217,296,157,178,180,60,157],
  [60,306,117,351,210,357,281,331,296,271,178,299,60,271,60,306],
  [700,104,931,363,666,559,645,525,800,401,787,351,807,297,733,224,687,238,
   630,219,513,309,489,268],
  [737,386,617,479,542,355,641,278,724,285,744,308,737,386],
  [65,598,262,444,289,483,178,578,186,618,163,659,207,735,231,748,278,735,
   333,763,431,691,457,730,248,892],
  [240,590,320,528,400,646,329,701,242,686,226,661,240,590],
  [555,803,497,768,297,923,361,944],
  [742,690,635,615,444,302,306,410,515,716,608,771,649,764],
  [941,423,704,599,765,640,937,512],
  [692,339,662,372,626,344,656,311,692,339],
  [265,631,295,598,330,626,300,659,265,631]],
  web: [
  [1000,396,999,608,929,621,889,717,930,776,776,930,717,889,621,929,608,999,
   392,999,379,929,283,889,224,930,69,772,116,725,128,746,110,768,232,890,
   285,851,407,902,413,966,587,966,596,900,719,850,768,891,890,768,851,715,
   902,593,966,587,966,413,902,407,851,285,891,232,768,110,719,150,644,114,
   647,85,717,111,759,70,780,73,930,224,889,283,929,379,987,380,1000,396],
  [94,683,112,692,126,677,100,596,34,587,34,413,98,407,149,285,110,232,
   232,110,285,149,407,98,413,34,587,34,589,60,608,68,612,2,392,1,
   379,71,283,111,224,70,70,224,111,283,71,379,1,392,0,604,13,620,
   71,621,94,683],
  [861,464,846,609,796,709,729,781,649,831,560,858,468,861,376,841,274,784,
   183,676,139,537,154,391,215,276,286,207,369,162,459,140,551,141,642,166,
   725,216,821,332,861,464],
  [699,499,517,508,517,645,675,657,699,500],
  [325,657,483,645,483,508,301,499,325,657],
  [176,448,267,461,297,320,235,306,176,448],
  [302,465,483,474,483,336,331,325,302,465],
  [517,173,517,302,654,293,594,211,517,173],
  [346,293,483,302,483,173,406,211,346,293],
  [338,690,400,784,483,827,483,680,338,690],
  [662,690,517,680,517,827,600,784,662,690],
  [698,465,669,325,517,336,517,474,698,465],
  [733,461,824,448,765,306,703,320,733,461],
  [630,199,689,287,740,277,630,199],
  [260,277,311,287,370,199,260,277],
  [172,500,185,593,225,679,291,663,266,496,172,482,172,500],
  [370,801,303,696,247,708,370,801],
  [753,708,697,696,630,801,753,708],
  [828,500,828,482,734,496,709,663,775,679,815,593,828,500]],
  pin: [
  [667,210,727,188,751,134,721,82,667,65,606,87,582,137,606,188,667,210],
  [627,116,621,148,642,173,679,178,710,154,691,102,658,96,627,116],
  [900,351,783,261,824,148,804,68,747,19,634,3,533,63,509,121,520,189,
   310,220,196,273,98,354,57,410,67,426,89,423,123,377,200,389,160,558,
   42,545,66,460,29,457,1,572,12,681,45,760,165,895,345,980,445,998,
   645,982,821,905,941,786,924,764,892,794,808,782,842,588,962,575,941,709,
   955,725,978,714,995,652,992,525,900,351],
  [150,350,214,359,280,266,150,350],
  [38,595,56,699,109,794,192,782,158,588,38,575,38,595],
  [262,917,280,925,204,811,132,822,261,916],
  [481,794,243,807,297,892,369,955,481,970],
  [481,600,196,595,230,777,481,763],
  [481,406,239,393,197,560,481,569],
  [481,221,369,236,253,364,481,375],
  [769,290,757,312,786,359,850,350,769,290],
  [738,348,729,365,747,363,738,348],
  [554,99,558,187,667,401,769,203,783,109,728,45,667,31,597,51,554,99],
  [519,375,609,373,519,221],
  [519,569,803,560,761,393,710,398,667,452,626,403,519,406],
  [519,763,770,777,804,595,519,600],
  [631,955,703,892,757,807,519,794,519,969,631,955],
  [868,822,796,811,720,925,868,821],
  [840,557,958,545,930,457,876,377,800,389,840,557]],
  chat: [
  [893,244,950,262,987,311,1000,385,987,554,951,603,896,621,877,708,794,853,
   735,908,546,996,489,990,461,962,450,921,462,870,507,835,587,849,666,819,
   734,753,789,619,779,350,700,209,569,129,416,134,348,169,247,288,218,367,
   202,615,77,617,13,555,0,385,29,283,77,249,146,244,252,98,323,47,
   424,9,526,1,626,23,718,72,796,146,856,244,893,244],
  [963,481,940,559,864,589,846,680,773,818,655,906,524,958,492,937,498,887,
   522,875,601,889,685,855,763,780,813,682,830,569,818,348,733,185,587,91,
   499,78,405,94,329,133,214,268,182,356,170,580,60,559,38,496,37,385,
   60,307,168,283,227,183,296,114,377,66,467,44,561,47,690,102,783,194,
   834,283,940,307,963,385],
  [765,490,729,341,634,235,500,194,366,235,271,341,244,411,241,551,285,662,
   254,839,276,847,411,769,481,785,614,757,717,660,754,573,765,490],
  [401,727,500,744,615,709,697,618,720,558,728,490,697,362,615,271,561,245,
   500,236,385,271,303,362,280,423,278,545,323,656,297,788,401,727],
  [397,490,371,519,345,490,371,461,397,490],
  [526,490,500,519,474,490,500,461,526,490],
  [655,490,629,519,603,490,629,461,655,490]],
  eyeLens: [
  [985,419,882,230,757,96,616,19,467,1,319,46,179,156,38,368,7,445,
   2,528,118,770,243,904,384,981,533,999,681,954,853,808,985,581,1000,500,
   985,419],
  [935,526,915,427,788,236,662,138,529,98,396,114,270,183,157,304,60,500,
   108,616,212,763,338,862,471,902,604,886,730,817,843,696,935,526]],
  eyePupil: [
  [500,0,367,18,248,68,146,146,68,248,18,367,0,500,18,633,68,752,
   146,854,248,932,367,982,500,1000,633,982,752,932,854,854,932,752,982,633,
   1000,500,982,367,932,248,853,147,752,68,633,18,500,0],
  [500,793,578,783,707,707,783,578,793,500,783,422,707,293,578,217,500,207,
   422,217,293,293,217,422,207,500,217,578,293,707,422,783,500,793]]
};

/* Spain map regions (per-mille of the map box) traced from the reference. */
const SPAIN_GREY = [
  [531,904,522,913,512,899,500,899,484,918,468,909,360,913,342,939,303,950,
   284,992,271,1000,232,965,204,881,174,855,181,865,164,853,139,856,130,810,
   150,757,169,752,174,729,200,752,212,747,229,762,246,755,253,733,265,732,
   264,744,273,738,273,704,302,664,325,669,368,712,370,705,395,709,421,698,
   428,705,439,695,458,695,462,703,479,688,492,690,504,722,495,744,537,771,
   536,798,547,822,563,833,531,904],
  [547,240,579,251,588,171,619,139,625,118,641,137,659,127,681,150,695,143,
   704,151,726,147,739,159,727,258,706,283,715,301,708,309,710,342,697,356,
   703,396,685,412,668,400,655,415,662,455,643,469,619,521,615,506,599,505,
   603,493,587,475,576,483,550,450,558,427,569,425,570,399,561,375,538,350,
   525,348,523,335,528,314,540,319,536,294,552,277,547,240],
  [254,19,265,34,286,33,351,55,351,70,335,73,335,82,322,76,308,94,
   289,95,279,105,265,100,256,112,244,99,225,106,216,99,211,116,185,123,
   169,105,182,88,172,90,156,56,167,32,234,33,254,19],
  [419,38,446,54,451,67,430,76,432,89,415,83,400,95,390,116,402,117,
   395,122,402,134,391,140,388,130,383,139,371,113,355,99,335,105,327,84,
   351,70,351,55,403,40,402,53,419,38],
  [475,136,495,139,498,155,480,152,475,136],
  [182,119,206,118,216,99,225,106,236,97,260,111,265,99,280,104,309,93,
   322,75,335,104,355,99,371,113,378,136,386,138,388,130,391,140,402,134,
   395,122,402,117,390,116,399,95,415,83,426,89,450,83,449,104,464,116,
   446,114,441,123,447,130,454,122,451,140,473,156,456,155,452,168,460,229,
   471,240,482,224,481,240,493,241,498,225,512,221,526,228,527,246,538,251,
   546,245,552,277,536,293,540,319,528,314,524,326,531,363,526,356,505,364,
   492,343,469,329,432,337,393,381,391,399,366,415,365,448,349,458,348,473,
   330,466,315,488,314,481,303,491,291,487,290,468,275,476,256,455,244,464,
   225,439,199,467,181,470,183,390,174,368,219,322,228,297,202,282,199,245,
   173,248,167,231,188,201,182,175,163,173,182,119],
  [912,223,920,247,898,278,850,312,837,335,755,367,734,398,748,409,732,429,
   726,427,737,420,720,434,703,420,694,405,703,396,697,356,710,342,708,309,
   715,301,706,284,727,258,739,195,731,129,782,145,789,183,815,175,830,195,
   848,183,878,196,882,185,906,178,926,198,922,209,913,208,912,223],
  [179,467,199,467,225,439,244,464,256,455,275,476,290,468,286,517,295,519,
   293,532,304,527,300,553,319,580,340,575,333,588,341,608,330,604,328,623,
   318,622,327,644,273,704,276,734,264,744,265,732,253,733,246,755,229,762,
   212,747,200,752,174,727,163,733,144,700,150,664,173,626,151,602,147,571,
   129,545,168,544,179,500,167,478,179,467],
  [101,7,105,18,118,0,125,18,131,8,150,33,166,32,156,56,172,89,
   182,88,169,105,182,119,163,173,182,175,188,201,167,231,169,247,156,243,
   154,256,138,266,117,253,105,261,103,251,84,265,89,234,81,220,49,234,
   32,259,29,226,51,195,32,208,48,184,34,190,26,179,36,181,40,155,
   34,149,19,172,16,164,33,131,16,143,13,120,8,113,0,118,0,100,
   13,89,6,85,28,77,29,64,44,68,66,55,68,62,73,52,80,63,
   83,50,72,47,84,40,71,45,72,31,101,7],
  [546,245,528,246,526,228,512,222,498,226,493,242,481,240,482,225,471,240,
   460,229,453,215,456,156,473,156,475,167,481,158,488,177,498,172,542,197,
   559,216,543,229,546,245],
  [423,351,433,372,428,412,448,435,448,470,455,466,458,488,425,496,404,517,
   398,511,417,495,413,483,375,464,367,474,361,461,346,474,349,458,366,449,
   368,420,377,421,397,377,423,351],
  [511,752,532,716,554,707,566,716,574,708,576,673,600,662,609,683,602,700,
   612,718,607,742,627,774,619,790,632,800,615,811,585,809,563,833,547,822,
   535,798,537,771,523,768,511,752],
  [561,65,565,73,581,73,577,101,589,92,590,101,629,118,590,167,578,212,
   587,235,570,253,543,237,559,216,528,185,506,178,506,161,499,159,514,154,
   518,118,533,110,542,77,561,65],
  [795,614,798,630,786,650,774,647,772,634,795,614],
  [918,507,908,514,918,515,913,530,921,536,928,529,938,538,906,595,899,583,
   887,583,878,558,867,571,855,552,886,520,918,507],
  [988,521,963,512,963,495,989,494,1000,517,996,526,988,521],
  [478,47,510,67,524,69,551,54,555,63,533,110,518,117,514,154,499,159,
   507,172,492,178,486,162,480,158,475,167,470,149,452,140,454,122,447,130,
   441,123,446,114,456,121,464,115,450,104,452,84,432,89,430,78,478,47],
  [474,135,495,139,498,155,480,152,474,135],
  [576,483,587,475,604,495,582,500,576,483],
  [599,505,615,506,619,521,624,517,621,504,637,493,643,469,654,468,662,455,
   655,415,668,400,685,412,699,407,720,434,658,574,671,630,700,662,645,713,
   626,774,607,742,610,713,603,711,615,643,598,638,590,623,597,590,576,583,
   570,572,599,505]];

const SPAIN_GREEN = [
  [423,351,432,337,469,329,492,343,505,364,526,356,531,363,538,350,570,399,
   569,425,560,425,550,450,556,466,577,476,582,500,599,505,570,572,576,583,
   597,590,590,623,599,638,612,636,615,667,607,672,596,660,581,667,574,707,
   566,716,555,707,533,716,511,751,495,744,504,721,492,690,479,688,462,703,
   458,695,439,694,428,705,421,698,395,709,370,704,368,711,312,663,328,637,
   318,622,328,623,330,604,341,608,333,588,340,575,319,580,300,553,304,527,
   293,532,295,519,286,517,288,489,316,487,330,466,350,476,361,461,366,473,
   375,464,418,488,398,511,403,517,425,495,456,492,455,466,448,470,453,449,
   425,399,433,372,423,351]];

/* -------------------------------------------------------- shared copy */
const T = {
  xs:    'Lorem ipsum dolor, consectetur adipiscing elit. Aenean nisi elit.',
  sm:    'Lorem ipsum dolor, consectetur adipiscing elit. Aenean nisi elit Donec pretium interdum.',
  md:    'Vestibulum a magna volutpat, mattis lorem vel, ultrices nisl. Curabitur convallis, risus a ' +
         'egestas lacinia, justo nunc vulputate lorem, ac lobortis purus diam blandit leo. Proin nisl ' +
         'ipsum, congue quis eleifend id, pretium aliquam metus. Aliquam rutrum leo vitae leo ultrices euismod. ',
  proin: 'Proin bibendum dolor ut tellus rhoncus, id gravida erat tristique. Vivamus quis nisl id orci ' +
         'posuere pellentesque eget et nisl. Nulla interdum dui at nisi condimentum, ac semper nulla ' +
         'commodo. Curabitur imperdiet magna imperdiet, pellentesque felis quis, pretium mi.',
  swot:  'Vestibulum a magna volutpat, mattis lorem vel, ultrices nisl. Curabitur convallis, risus a ' +
         'egestas lacinia, justo nunc vulputate lorem, ac lobortis purus diam blandit leo. Proin nisl ' +
         'ipsum, congue quis eleifend id.',
  svc:   'Vestibulum a magna volutpat, ma lorem vel, ultrices nisl. Curabitur convallis, risus a egestas ' +
         'lacinia, justo nunc vulputate lorem, ac lobortis purus diam blandit leo.',
  mock:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse dictum enim vel ante ' +
         'scelerisque euismod. Nulla',
  toc:   'Nulla commodo ut risus at maximus. Cras tempor hendrerit odio, at commodo risus iaculis id.',
  price: 'Vestibulum a magna volutpat, mattis lore, ultrices nisl. Curabitur convallis, risus a egestas ' +
         'lacinia, justo nunc vulputate lorem.',
  testi: '\u201C Curabitur convallis, risus a egestas lacinia, justo nunc vulputate lorem, ac lobortis ' +
         'purus diam blandit leo. \u201C',
};

/* ------------------------------------------------------ composite parts */

/** Two-line stacked page title, e.g. "About Our" / "Fiz". */
const pageTitle = (s, l1, l2, x, y, w, size = 44, gap = 0.548) => {
  head(s, l1, x, y, w, size > 50 ? 0.959 : 0.547, size);
  head(s, l2, x, y + gap, w, size > 50 ? 0.959 : 0.547, size);
};

/**
 * The deck's recurring "statistic card": floating white card with a green icon
 * tile poking out of its top edge, a big numeral and a line of grey copy.
 * All child offsets are constant relative to the card's top-left corner.
 */
const statCard = (s, x, y, value, o = {}) => {
  card(s, x, y, 2.318, 2.498);
  iconTile(s, ICONS.money, x + 0.613, y - 0.668, 1.093, 1.178, x + 0.860, y - 0.402, 0.598, 0.647);
  head(s, value, x + 0.293, y + 0.852, 1.733, 0.581, o.size || 48, { align: o.align || 'left' });
  body(s, T.sm, x + 0.293, y + 1.498, 1.733, 0.703);
};

/** Rounded timeline / SWOT card: roundRect with its top-left corner squared off. */
const notchCard = (s, x, y, w, h) => {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.29237 * Math.min(w, h),
                            fill: { color: C.white }, line: { type: 'none' } });
  plain(s, x, y, 1.236, 1.236, C.white);
};

/** Eye pictogram (lens outline + pupil ring) used on the Vision / Leader slides. */
const eyeIcon = (s, x, y) => {
  drawIcon(s, ICONS.eyeLens, x, y, 0.636, 0.395);
  drawIcon(s, ICONS.eyePupil, x + 0.231, y + 0.105, 0.185, 0.185);
};

/* --------------------------------------------------------- slide bodies */

// 1 - Title slide
function slide01(s) {
  imagePlaceholder(s, 8.535, 1.227, 3.344, 5.045);
  plain(s, 1.451, 5.925, 4.612, 0.347, C.black);
  plain(s, 1.451, 3.957, 5.700, 1.678, C.white);
  head(s, 'Fiz', 1.451, 1.227, 4.612, 0.959, 80);
  head(s, 'Creative', 1.451, 2.238, 5.266, 0.959, 80);
  body(s, 'Exploring Effective Financial Management Strategies for Building Long-Term Financial Well-being',
       1.869, 4.253, 3.308, 1.086, { size: 14 });
  iconTile(s, ICONS.coins, 5.595, 3.957, 1.557, 1.678, 5.906, 4.253, 0.935, 1.086);
  statCard(s, 7.329, 3.137, '80%');
  s.addText([{ text: 'Created By: ', options: { bold: true } }, { text: ' Marceline Anderson' }],
            { ...TXT, x: 1.639, y: 6.012, w: 2.106, h: 0.142, fontFace: F.body, fontSize: 10,
              color: C.white, lineSpacingMultiple: 1, paraSpaceBefore: 10 });
  s.addText([{ text: 'Speaker: ', options: { bold: true } }, { text: ' Rosa Maria Aguado' }],
            { ...TXT, x: 4.066, y: 6.012, w: 1.997, h: 0.142, fontFace: F.body, fontSize: 10,
              color: C.white, lineSpacingMultiple: 1, paraSpaceBefore: 10 });
}

// 2 - Welcome Message
function slide02(s) {
  imagePlaceholder(s, 6.5, 1.227, 5.099, 5.045);
  pageTitle(s, 'Welcome', 'Message', 1.734, 1.227, 3.87);
  statCard(s, 7.402, 1.895, '+2K');
  body(s, 'Phasellus eu luctus purus, vitae imperdiet ante. Duis sed facilisis ipsum. Nullam viverra sem ' +
          'sit amet vulputate ullamcorper. In nulla nibh, molestie eu lorem sit amet, viverra congue nunc. ' +
          'Aliquam ac lorem sed magna tincidunt euismod. Suspendisse fermentum tortor elit, sit amet ' +
          'porttitor tortor scelerisque vitae.', 1.734, 3.393, 3.87, 1.19);
  [['20K', 1.734], ['75%', 3.487]].forEach(([v, x]) => {
    head(s, v, x, 5.082, 1.261, 0.422, 32, { color: C.green });
    body(s, T.xs, x, 5.569, 1.481, 0.703);
  });
}

// 3 - Table of contents
function slide03(s) {
  pageTitle(s, 'Table Of', 'Content', 1.734, 1.227, 3.87);
  [['01', 'About Our Fiz'], ['02', 'Vision & Mission'],
   ['03', 'Fiz Timeline'], ['04', 'Fiz Services']].forEach(([num, name], i) => {
    const y = 1.541 + i * 1.3245;
    head(s, num, 7.905, y, 0.788, 0.471, 40, { color: C.green });
    label(s, name, 8.851, y + 0.008, 2.362, 0.187);
    body(s, T.toc, 8.851, y + 0.329, 2.749, 0.43);
  });
  imagePlaceholder(s, 0.133, 3.028, 7.006, 4.306);
}

// 4 - About Our Fiz
function slide04(s) {
  imagePlaceholder(s, 7.485, 0, 4.12, 7.5);
  plain(s, 1.316, 4.363, 7.056, 1.549, C.white);
  pageTitle(s, 'About Our', 'Fiz', 1.734, 1.227, 3.87);
  body(s, T.md, 1.734, 2.955, 4.198, 0.937);
  [['90%', 1.689], ['12K', 3.888]].forEach(([v, x]) => {
    head(s, v, x, 4.619, 1.261, 0.422, 32, { color: C.green });
    body(s, T.xs, x, 5.106, 1.742, 0.458);
  });
  statCard(s, 6.182, 3.933, '+2K');
}

// 5 - Our Great Vision (vision card + goals card)
function slide05(s) {
  imagePlaceholder(s, 9.389, 0.247, 3.736, 6.011);
  pageTitle(s, 'Our Great', 'Vision', 1.734, 1.227, 3.177);
  plain(s, 1.734, 2.796, 5.710, 2.276, C.white);
  card(s, 6.809, 2.447, 4.790, 3.951);
  label(s, 'Vision', 2.173, 3.338, 1.056, 0.187);
  body(s, T.md, 2.173, 3.684, 4.198, 0.937);
  label(s, 'Goals', 7.407, 3.463, 1.056, 0.187);
  bullets(s, ['Vestibulum a magna volutpat, mattis lorem vel, ultrices nisl.',
              'Curabitur convallis, risus a egestas lacinia, justo nun vulputate lorem, ac lobortis purus diam blandit leo.',
              'Proin nisl ipsum, congue quis eleifend id, pretium aliquam metus. Aliquam rutrum leo vitae leo ultrices euismod.',
              'Nulla vel enim sit amet tortor faucibus rhoncus. Suspendisse vestibulum hendrerit turpis.'],
          7.407, 3.808, 3.595, 2.151);
  plain(s, 5.235, 2.230, 1.093, 1.178, C.green);
  eyeIcon(s, 5.464, 2.621);
  iconTile(s, ICONS.target, 10.024, 1.950, 1.093, 1.178, 10.280, 2.237, 0.636, 0.603);
  imagePlaceholder(s, 0.308, 5.563, 6.03, 1.69);
}

// 6 - Our Great Mission
function slide06(s) {
  plain(s, 1.734, 3.997, 7.993, 2.276, C.white);
  pageTitle(s, 'Our Great', 'Mission', 2.469, 1.515, 3.177);
  label(s, 'Mission', 2.469, 4.443, 1.056, 0.187);
  body(s, T.md, 2.469, 4.789, 4.198, 0.937);
  plain(s, 0.641, 2.812, 1.093, 1.178, C.green);
  eyeIcon(s, 0.870, 3.204);
  imagePlaceholder(s, 8.415, 1.221, 4.918, 5.057);
}

/**
 * Slides 7 & 8 share one timeline layout: a green rule with two dated pills,
 * two rounded story cards below and one floating intro card above.
 */
function timelineSlide(s, o) {
  const introX = o.introLeft ? 1.734 : 7.191;
  card(s, introX, 0.822, 5.169, 1.807);
  body(s, T.proin, introX + 0.485, 1.249, 4.198, 0.951);
  if (o.title) pageTitle(s, o.title[0], o.title[1], 1.734, 1.227, 3.177);
  [1.734, 7.191].forEach((x, i) => {
    notchCard(s, x, 3.968, 5.169, 1.806);
    card(s, x + 0.355, 3.810, 2.028, 0.444);
    label(s, o.cards[i].head, x + 0.485, 3.963, 1.458, 0.187);
    body(s, T.md, x + 0.485, 4.417, 4.198, 1.054);
  });
  s.addShape('line', { x: o.ruleLeft, y: 3.321, w: 11.599, h: 0,
                       line: { color: C.green, width: 0.5 } });
  [1.734, 7.191].forEach((x, i) => {
    s.addShape('roundRect', { x: x + 0.355, y: 3.115, w: 1.589, h: 0.428, rectRadius: 0.214,
                              fill: { color: C.green }, line: { type: 'none' } });
    s.addShape('line', { x: x + 0.569, y: 3.543, w: 0, h: 0.267, line: { color: C.green, width: 0.5 } });
    label(s, o.cards[i].years, x + 0.485, 3.206, 1.458, 0.246, { color: C.white });
  });
  imagePlaceholder(s, 1.734, 6.418, 11.599, 1.082);
}

const slide07 = s => timelineSlide(s, { title: ['Fiz', 'Timeline'], introLeft: false, ruleLeft: 1.734,
  cards: [{ head: 'First Time', years: '2012 - 2015' }, { head: 'Second Time', years: '2015 - 2017' }] });

const slide08 = s => timelineSlide(s, { title: null, introLeft: true, ruleLeft: 0,
  cards: [{ head: 'Third Time', years: '2017 - 2020' }, { head: 'Fourth Time', years: '2020 - 2024' }] });

// 9 - Fiz Services
function slide09(s) {
  imagePlaceholder(s, 8.415, 0.246, 4.682, 4.697);
  pageTitle(s, 'Fiz', 'Services', 1.734, 1.227, 3.177);
  [{ x: 1.734, head: 'Life Insurance',     glyph: ICONS.hand,     gx: 2.470, gy: 3.057, gw: 0.421, gh: 0.624 },
   { x: 5.075, head: 'Health Insurance',   glyph: ICONS.umbrella, gx: 5.727, gy: 3.063, gw: 0.596, gh: 0.614 },
   { x: 8.415, head: 'Property Insurance', glyph: ICONS.notes,    gx: 9.080, gy: 3.064, gw: 0.572, gh: 0.610 }
  ].forEach(col => {
    card(s, col.x, 3.33, 3.184, 2.943);
    plain(s, col.x + 0.488, 2.911, 0.917, 0.917, C.green);
    drawIcon(s, col.glyph, col.gx, col.gy, col.gw, col.gh);
    label(s, col.head, col.x + 0.492, 4.248, 1.857, 0.187);
    body(s, T.svc, col.x + 0.492, 4.633, 2.2, 1.201);
  });
}

// 10 - Meet The Team
function slide10(s) {
  [4.781, 7.209, 9.616].forEach(x => imagePlaceholder(s, x, 1.227, 2.333, 5.045));
  pageTitle(s, 'Meet The', 'Team', 1.384, 1.227, 3.177);
  [{ x: 4.843, name: 'Adeline Palmerston', role: 'Director of Finance' },
   { x: 7.271, name: 'Brigitte Schwartz',  role: 'Insurance Specialist' },
   { x: 9.678, name: 'Cahaya Dewi',        role: 'Marketing' }].forEach(p => {
    plain(s, p.x, 5.259, 2.21, 0.956, C.white);
    label(s, p.name, p.x + 0.176, 5.524, 1.857, 0.187, { align: 'center' });
    body(s, p.role, p.x + 0.353, 5.716, 1.504, 0.187, { align: 'center' });
  });
  statCard(s, 1.384, 3.775, '+2K');
}

// 11 - Meet The Leader
function slide11(s) {
  imagePlaceholder(s, 1.692, 0.714, 2.808, 6.072);
  plain(s, 11.469, 2.963, 1.093, 1.178, C.green);
  eyeIcon(s, 11.698, 3.355);
  plain(s, 1.879, 5.666, 2.438, 0.956, C.white);
  label(s, 'James Marfield', 2.169, 5.931, 1.857, 0.187, { align: 'center' });
  body(s, 'Founder', 2.346, 6.124, 1.504, 0.187, { align: 'center' });
  pageTitle(s, 'Meet The', 'Leader', 6.301, 1.811, 3.177);
  card(s, 6.301, 4.124, 5.169, 1.807);
  body(s, T.proin, 6.786, 4.551, 4.198, 0.951);
}

// 12 - Break Slide
function slide12(s) {
  imagePlaceholder(s, 1.694, 4.556, 3.125, 2.944);
  imagePlaceholder(s, 4.984, 4.556, 3.365, 2.944);
  imagePlaceholder(s, 8.514, 1.227, 3.125, 6.273);
  pageTitle(s, 'Break', 'Slide', 1.694, 1.227, 4.369, 80, 1.011);
  statCard(s, 7.27, 3.612, '+2K');
  card(s, 1.694, 3.612, 4.076, 1.597);
  body(s, '\u201CFinance is not just about money, it is about freedom.\u201D - Robert T. Kiyosaki',
       2.234, 4.033, 3.049, 0.755, { size: 14 });
}

// 13 - SWOT Analysis
function slide13(s) {
  imagePlaceholder(s, 7.607, 0.193, 5.532, 7.114);
  pageTitle(s, 'SWOT', 'Analysis', 1.384, 1.144, 3.177);
  [{ x: 1.384, y: 2.819, head: 'Strengths',     letter: 'S', lx: 5.334 },
   { x: 6.781, y: 2.819, head: 'Weaknesses',    letter: 'W', lx: 10.731 },
   { x: 1.384, y: 4.854, head: 'Opportunities', letter: 'O', lx: 5.334 },
   { x: 6.781, y: 4.854, head: 'Threas',        letter: 'T', lx: 10.731 }].forEach(q => {
    notchCard(s, q.x, q.y, 5.169, 1.503);
    card(s, q.x + 0.354, q.y - 0.158, 2.028, 0.444);
    label(s, q.head, q.x + 0.579, q.y - 0.052, 1.458, 0.187);
    body(s, T.swot, q.x + 0.485, q.y + 0.449, 4.198, 1.054);
    plain(s, q.lx, q.y - 0.346, 0.733, 0.683, C.green);
    head(s, q.letter, q.lx, q.y - 0.291, 0.733, 0.547, 44, { color: C.white, align: 'center' });
  });
}

// 14 - Gallery Fiz
function slide14(s) {
  imagePlaceholder(s, 5.444, 1.227, 7.714, 3.818);
  imagePlaceholder(s, 2.056, 5.194, 11.103, 1.745);
  pageTitle(s, 'Gallery', 'Fiz', 1.734, 1.227, 2.735);
  plain(s, 1.734, 3.612, 4.21, 1.964, C.white);
  body(s, T.md, 2.056, 4.072, 3.568, 1.187);
  statCard(s, 6.23, 3.612, '+2K');
}

// 15 - Pricing List
function slide15(s) {
  imagePlaceholder(s, 8.415, 1.221, 4.918, 5.057);
  pageTitle(s, 'Pricing', 'List', 1.734, 1.227, 2.735);
  [{ x: 1.734, plan: 'Advanced', price: '$99.99' },
   { x: 5.569, plan: 'Premium',  price: '$199.99' }].forEach(p => {
    card(s, p.x, 3.104, 3.619, 3.168);
    plain(s, p.x + 0.764, 2.781, 2.083, 0.643, C.green);
    label(s, p.plan, p.x + 0.981, 2.930, 1.65, 0.347, { size: 20, color: C.white, align: 'center' });
    head(s, p.price, p.x + 0.447, 3.835, 2.506, 0.547, 44);
    body(s, T.price, p.x + 0.447, 4.705, 2.725, 0.698);
    plain(s, p.x + 0.948, 5.725, 1.723, 0.547, C.black);
    label(s, 'Book Now', p.x + 1.077, 5.872, 1.458, 0.254, { color: C.white, align: 'center' });
  });
}

/**
 * Laptop mockup (slide 16): silver lid, black bezel, blank screen, and the base
 * built from four thin strips that shade towards the front edge.
 */
function laptopMockup(s, screen) {
  const rr = (x, y, w, h, r, color) =>
    s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color }, line: { type: 'none' } });
  rr(6.415, 2.890, 5.975, 3.995, 0.07, 'C9CBCB');   // silver lid edge
  rr(6.450, 2.925, 5.905, 3.945, 0.06, '000000');   // black bezel
  plain(s, 6.635, 3.155, 5.550, 3.480, screen);     // blank screen
  s.addShape('ellipse', { x: 9.385, y: 3.005, w: 0.055, h: 0.055,
                          fill: { color: '1E2A44' }, line: { type: 'none' } });  // webcam
  rr(5.750, 6.875, 7.300, 0.140, 0.05, 'EBECEC');   // silver deck
  rr(8.940, 6.875, 0.940, 0.060, 0.02, 'DDDDDD');   // trackpad lip
  rr(5.770, 7.010, 7.260, 0.090, 0.04, '8A8A8A');   // front edge, shaded
  rr(5.855, 7.078, 7.090, 0.052, 0.02, '242424');
}

/** Phone mockup (slide 17): purple-grey rim, black body, blank screen, notch pill. */
function phoneMockup(s, screen) {
  const rr = (x, y, w, h, r, color) =>
    s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color }, line: { type: 'none' } });
  rr(1.790, 0.950, 2.733, 5.596, 0.500, '5B4A5C');  // outer rim
  rr(1.800, 0.960, 2.713, 5.576, 0.495, 'A08CA0');  // metallic highlight
  rr(1.818, 0.978, 2.677, 5.540, 0.485, '5A4157');  // inner rim
  rr(1.840, 1.000, 2.635, 5.496, 0.475, '1D1D1B');  // body
  rr(1.923, 1.086, 2.467, 5.344, 0.400, screen);    // blank screen
  rr(2.832, 1.164, 0.610, 0.185, 0.090, '1D1D1B');  // notch pill
  s.addShape('ellipse', { x: 3.330, y: 1.222, w: 0.062, h: 0.062,
                          fill: { color: '272B45' }, line: { type: 'none' } });
}

// 16 - Unique Mockup (laptop)
function slide16(s) {
  laptopMockup(s, C.bg);
  imagePlaceholder(s, 6.628, 3.137, 5.561, 3.493);
  pageTitle(s, 'Unique', 'Mockup', 1.667, 1.227, 2.735);
  [{ y: 3.671, head: 'Website' }, { y: 4.889, head: 'Information' }].forEach(sec => {
    label(s, sec.head, 1.667, sec.y, 2.294, 0.326, { size: 18, lh: 1.5 });
    body(s, T.mock, 1.667, sec.y + 0.5, 3.319, 0.454);
  });
  statCard(s, 9.62, 1.888, '+2K');
}

// 17 - Unique Mockup (phone)
function slide17(s) {
  phoneMockup(s, C.bg);
  imagePlaceholder(s, 1.923, 1.08, 2.456, 5.34);
  pageTitle(s, 'Unique', 'Mockup', 6.01, 1.559, 2.735);
  label(s, 'Mobile', 6.01, 3.212, 2.294, 0.326, { lh: 1.5 });
  body(s, T.mock, 6.01, 3.711, 3.319, 0.454);
  [['90%', 6.01], ['12K', 8.21]].forEach(([v, x]) => {
    head(s, v, x, 4.833, 1.261, 0.422, 32, { color: C.green });
    body(s, T.xs, x, 5.320, 1.742, 0.458);
  });
  iconTile(s, ICONS.money, 3.916, 4.774, 1.093, 1.178, 4.164, 5.040, 0.598, 0.647);
}

// 18 - Unique Spain Map
function slide18(s) {
  const MX = 6.021, MY = 1.455, MW = 6.328, MH = 4.701;
  const region = (poly, color) => {
    const pts = [];
    for (let i = 0; i < poly.length; i += 2) {
      pts.push({ x: (poly[i] / 1000) * MW, y: (poly[i + 1] / 1000) * MH, moveTo: i === 0 });
    }
    pts.push({ close: true });
    s.addShape('custGeom', { x: MX, y: MY, w: MW, h: MH, points: pts,
                             fill: { color }, line: { color: C.bg, width: 1 } });
  };
  SPAIN_GREY.forEach(p => region(p, C.mapGrey));
  SPAIN_GREEN.forEach(p => region(p, C.green));

  plain(s, 1.316, 4.535, 6.406, 1.549, C.white);
  pageTitle(s, 'Unique', 'Spain Map', 1.734, 1.149, 3.87, 44, 0.656);
  body(s, T.md, 1.734, 3.127, 4.198, 0.937);
  [['90%', 1.689], ['12K', 3.888]].forEach(([v, x]) => {
    head(s, v, x, 4.791, 1.261, 0.422, 32, { color: C.green });
    body(s, T.xs, x, 5.278, 1.742, 0.458);
  });
  statCard(s, 9.212, 4.178, '700K', { size: 40 });
}

// 19 - Client Testimonial
function slide19(s) {
  imagePlaceholder(s, 0.236, 3.526, 4.444, 3.779);
  plain(s, 1.403, 0.639, 4.931, 1.944, C.white);
  pageTitle(s, 'Client', 'Testimonial', 1.667, 1.227, 3.583);
  [{ x: 3.875, name: 'Harumi Kobayasi',   sx: 6.333, px: 4.189 },
   { x: 8.180, name: 'Francisco Andrade', sx: 10.639, px: 8.494 }].forEach(t => {
    card(s, t.x, 3.75, 4.153, 2.523);
    label(s, t.name, t.x + 0.314, 5.176, 1.999, 0.187);
    body(s, T.testi, t.x + 0.314, 5.535, 3.525, 0.473);
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', { x: t.sx + i * 0.2063, y: 4.713, w: 0.178, h: 0.178,
                            fill: { color: C.green }, line: { type: 'none' } });
    }
    imagePlaceholder(s, t.px, 2.946, 1.944, 1.944);
  });
  head(s, '\u201C', 9.35, 0.365, 3.583, 0.547, 44, { align: 'right' });
}

// 20 - Thank You
function slide20(s) {
  imagePlaceholder(s, 9.366, 0.241, 3.679, 7.019);
  imagePlaceholder(s, 5.53, 5.625, 3.679, 1.635);
  plain(s, 1.451, 3.637, 5.71, 2.276, C.white);
  pageTitle(s, 'Thank', 'You', 1.451, 1.227, 4.612, 80, 1.011);
  card(s, 6.667, 2.444, 3.679, 3.468);
  label(s, 'Let\u2019s Talk', 7.163, 2.840, 1.774, 0.42, { size: 28 });
  body(s, 'Phasellus nec ex id tortor feugiat temp accumsan lectus. Nullam vitae tortor vitae arcu ' +
          'ornare sagittis sed quis urna.', 7.163, 3.468, 2.703, 0.696);
  [{ glyph: ICONS.web,  y: 4.544, w: 0.181, h: 0.181, ty: 4.478, tw: 1.759, txt: 'www.yourwebsite.com' },
   { glyph: ICONS.pin,  y: 4.905, w: 0.165, h: 0.203, ty: 4.872, tw: 2.155, txt: '123 Anywhere St., Any City, ST 12345' },
   { glyph: ICONS.chat, y: 5.287, w: 0.202, h: 0.181, ty: 5.230, tw: 2.155, txt: '123-456-7890' }
  ].forEach(r => {
    drawIcon(s, r.glyph, r.glyph === ICONS.chat ? 7.147 : 7.159, r.y, r.w, r.h, C.green);
    body(s, r.txt, 7.475, r.ty, r.tw, 0.231);
  });
  body(s, T.md, 1.96, 4.24, 4.198, 0.937);
}

/* ------------------------------------------------------------- assemble */
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
                slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'Marceline Anderson';
pptx.title = 'Fiz Creative';

SLIDES.forEach(build => {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, OUT_NAME) })
    .then(f => console.log('wrote ' + f));
