// Kifa - Kid Fashion and Toys Presentation
// Standalone pptxgenjs re-creation of the source deck (30 slides, 13.333 x 7.5 in).
// Raster photos in the original are replaced by flat colour placeholder shapes.

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette / type
const C = {
  ink:    '3F3F3F',  // theme tx1  - headings
  gray:   '757575',  // theme tx2  - sub-headings
  mute:   '929292',  // theme accent6 - body copy
  pink:   'FF4980',  // theme accent1
  green:  '7DBE48',  // theme accent2
  amber:  'FFAE01',  // theme accent4
  white:  'FFFFFF',
  smoke:  'F2F2F2',  // bg1 lumMod 95% - faint background doodles
  black:  '000000',
};
const HEAD = 'Fredoka';   // theme major font
const BODY = 'Poppins';   // theme minor font
// pptxgenjs rewrites the shadow object it is handed, so hand it a fresh one each time.
const CARD_SHADOW = () => ({ type: 'outer', blur: 20, offset: 0.001, angle: 90, color: C.black, opacity: 0.10 });
const PANEL_SHADOW = () => ({ type: 'outer', blur: 20, offset: 0.001, angle: 90, color: C.black, opacity: 0.05 });

// ---------------------------------------------------------------- freeform outlines
// Each entry lists closed sub-paths as flat x,y pairs in per-mille of the shape box.
const SHAPES = {
  arrowGlyph: [
    [500,0,589,43,589,852,870,718,958,719,1000,756,500,1000,0,754,47,718,131,718,411,852,412,43,500,0]
  ],
  ballGlyph: [
    [659,890,696,665,682,445,734,415,762,362,910,408,919,544,869,702,765,827,659,890],
    [109,343,322,318,554,325,590,421,401,615,190,783,87,580,81,459,109,343],
    [658,395,605,342,658,290,710,342,658,395],
    [807,212,734,270,624,242,540,80,687,123,807,212],
    [500,0,305,39,146,146,39,305,0,500,39,695,146,854,305,961,500,1000,695,961,854,854,961,695,1000,500,
     961,305,854,146,695,39,500,0]
  ],
  balloonGlyph: [
    [398,391,541,423,577,462,572,494,509,537,398,554,288,537,224,494,220,462,256,423,398,391],
    [832,0,934,18,998,75,966,142,855,194,903,334,866,450,902,486,579,644,579,1000,434,1000,434,804,362,804,
     362,1000,218,1000,218,717,112,815,27,814,2,780,133,640,188,610,329,578,522,576,817,446,856,342,
     808,194,710,153,663,87,712,26,832,0]
  ],
  blobLeft1: [
    [1000,0,335,0,173,131,84,226,36,303,12,366,1,442,9,561,43,683,93,805,147,908,208,1000,1000,1000,1000,0]
  ],
  blobLeft2: [
    [0,0,717,0,802,69,899,183,966,311,991,403,1000,500,991,597,966,689,899,817,837,895,722,997,0,1000,0,0]
  ],
  blobLeft3: [
    [0,0,738,0,816,69,880,143,951,267,982,357,1000,500,992,597,968,689,907,817,849,895,742,997,0,1000,0,0]
  ],
  blobLeft4: [
    [169,0,1000,0,1000,1000,95,1000,16,862,1,800,1,760,26,687,159,487,197,386,200,262,169,0]
  ],
  blobOrganic: [
    [359,0,405,13,534,92,597,111,666,106,828,64,888,68,937,100,982,175,998,244,1000,292,989,365,969,437,
     815,877,771,955,745,982,717,998,669,994,534,914,467,888,396,885,236,907,178,895,145,872,104,821,
     51,704,19,599,2,493,6,393,27,325,93,222,238,66,298,18,359,0]
  ],
  blobRight1: [
    [269,0,1000,0,1000,1000,508,1000,509,794,498,744,478,708,431,672,262,607,175,564,94,509,32,448,7,404,
     0,360,8,315,29,270,120,156,250,39,269,0]
  ],
  blobRight2: [
    [308,0,1000,0,1000,1000,191,1000,175,977,92,820,27,642,3,520,4,403,33,303,95,202,182,109,308,0]
  ],
  blobRound: [
    [616,0,691,16,896,116,1000,147,1000,1000,114,1000,45,840,7,698,0,629,3,561,31,466,65,407,134,325,
     383,109,512,24,564,6,616,0]
  ],
  blobWave: [
    [209,0,285,13,357,44,422,93,462,136,518,224,622,446,682,553,730,614,892,761,958,840,990,916,1000,1000,
     0,1000,0,57,95,14,209,0]
  ],
  buildingGlyph: [
    [501,661,115,932,885,932,501,661],
    [925,473,675,689,925,904,925,473],
    [75,472,75,903,325,688,75,472],
    [500,362,500,435,500,362],
    [501,248,630,305,664,432,459,458,461,330,551,337,622,433,621,382,548,295,428,310,404,464,591,503,
     500,546,385,501,339,396,386,291,501,248],
    [250,182,250,560,362,658,500,615,638,658,750,560,750,182,250,182],
    [500,0,650,114,825,114,825,212,1000,365,1000,1000,0,1000,0,365,175,214,175,114,350,114,500,0]
  ],
  cardFooter: [
    [0,0,1000,0,991,624,969,896,946,993,937,1000,63,1000,54,993,31,896,9,624,0,0]
  ],
  cardTab: [
    [0,0,1000,0,999,827,986,922,967,973,951,993,934,1000,66,1000,49,993,33,973,14,922,1,827,0,0]
  ],
  chatGlyph: [
    [50,0,0,48,0,730,37,776,200,778,200,1000,400,778,986,761,1000,56,981,12,50,0],
    [662,387,683,330,735,306,787,330,809,387,787,444,735,468,683,444,662,387],
    [426,387,448,330,500,306,552,330,574,387,552,444,500,468,448,444,426,387],
    [191,387,213,330,265,306,317,330,338,387,317,444,265,468,213,444,191,387]
  ],
  cloudPink: [
    [387,0,435,19,479,72,516,155,546,268,576,181,612,125,670,87,714,74,767,80,807,104,855,156,895,226,
     970,415,992,504,1000,579,995,677,977,785,932,936,903,1000,0,1000,0,94,63,86,116,114,154,157,187,236,
     230,131,273,68,323,17,387,0]
  ],
  cloudTL: [
    [715,0,733,5,833,125,864,138,883,132,902,110,951,0,1000,0,965,122,943,162,915,191,882,210,851,212,
     785,173,813,348,813,436,778,549,731,640,705,673,669,701,630,711,596,705,548,665,444,526,437,608,
     409,757,354,895,328,938,278,976,208,998,146,996,80,966,0,899,0,825,86,901,126,921,168,930,212,927,
     256,910,295,882,319,847,354,757,375,659,395,487,382,367,397,364,397,353,417,356,433,372,499,506,
     594,622,620,635,641,633,673,612,701,578,759,445,770,380,736,155,703,44,703,8,715,0]
  ],
  duckGlyph: [
    [899,324,882,178,788,40,675,0,562,40,481,144,452,251,460,376,504,477,293,527,169,490,58,415,31,429,
     2,684,72,869,146,949,280,993,643,997,785,961,839,884,848,770,826,633,776,550,850,466,945,422,1000,340,
     899,324],
    [786,290,749,242,786,194,824,242,786,290]
  ],
  familyGlyph: [
    [145,286,194,301,234,345,264,471,288,354,364,287,416,293,459,329,500,471,524,355,590,290,645,290,
     694,329,736,471,758,359,793,310,852,286,914,310,950,362,1000,633,985,680,957,677,925,549,925,1000,
     866,1000,866,694,843,694,843,1000,784,1000,784,551,762,661,737,686,711,661,690,551,690,653,725,816,
     690,816,690,1000,631,1000,631,816,607,816,607,1000,548,1000,548,816,513,816,548,649,548,549,526,659,
     501,684,476,659,454,549,454,1000,395,1000,395,694,372,694,372,1000,313,1000,313,547,291,657,266,682,
     240,657,218,547,218,647,254,814,218,814,218,998,160,998,160,814,136,814,136,998,77,998,77,814,42,814,
     77,651,77,549,39,681,10,672,1,620,52,356,92,305,145,286],
    [852,0,902,36,923,122,902,209,852,245,802,209,782,122,802,36,852,0],
    [617,0,667,36,687,122,667,209,617,245,567,209,546,122,567,36,617,0],
    [381,0,431,36,452,122,431,209,381,245,331,209,310,122,331,36,381,0],
    [145,0,195,36,216,122,195,209,145,245,95,209,75,122,95,36,145,0]
  ],
  gearGlyph: [
    [499,676,375,625,323,500,375,375,499,324,624,375,676,500,624,625,499,676],
    [896,390,858,298,894,188,811,104,700,141,608,103,558,0,440,0,389,103,298,141,188,104,104,188,141,298,
     103,391,0,441,0,559,103,610,141,702,104,812,188,896,298,859,391,897,442,1000,560,1000,611,897,702,859,
     812,896,896,812,859,702,897,609,1000,557,1000,440,896,390]
  ],
  grassBump: [
    [41,0,62,4,101,33,154,126,197,268,229,451,249,325,288,168,339,58,376,15,417,0,437,4,477,33,513,89,
     573,268,605,451,650,216,680,126,714,58,772,4,821,7,848,28,873,61,919,163,972,387,998,675,991,1000,
     0,1000,0,16,41,0]
  ],
  greenScribble: [
    [1000,333,975,285,930,267,846,275,767,303,813,229,788,178,736,166,750,168,636,203,660,148,634,101,
     606,89,564,90,473,125,479,74,466,36,448,14,407,0,304,44,204,135,45,334,2,414,0,435,20,454,46,438,
     64,420,56,408,87,362,231,196,322,110,399,70,407,69,342,195,348,191,116,546,154,596,178,602,204,576,
     395,287,480,211,566,165,389,384,289,550,322,594,362,589,575,337,646,278,725,240,437,587,403,664,
     441,699,460,699,485,681,617,508,702,421,830,352,928,332,925,366,721,613,587,808,556,884,542,965,
     550,987,558,980,562,990,566,979,579,1000,601,984,643,865,686,786,813,614,956,444,976,412,963,409,
     989,387,1000,333],
    [8,414,15,392,8,414],
    [16,389,30,371,16,389]
  ],
  greenSquiggle: [
    [999,76,992,54,888,103,768,354,721,107,679,133,636,196,602,69,587,49,568,47,521,82,454,224,412,136,
     368,100,303,123,259,218,234,101,214,56,156,1,124,21,85,93,51,216,16,430,1,606,13,635,47,410,88,244,
     125,194,153,188,188,216,204,218,232,310,186,514,167,730,172,879,208,999,220,993,235,947,268,734,
     279,546,276,348,299,290,375,276,421,330,394,497,377,692,392,820,408,859,425,868,439,859,456,822,
     486,667,493,510,477,344,503,270,544,206,558,214,566,199,586,221,590,213,603,264,564,417,550,573,
     561,697,572,727,584,734,610,673,643,496,647,328,668,284,714,238,744,402,711,562,701,662,707,781,
     730,847,742,840,762,749,778,470,884,233,896,235,982,182,998,130,999,76],
    [231,741,221,793,210,806,195,758,200,643,243,449,246,591,231,741],
    [461,613,443,673,420,696,406,677,406,643,444,424,464,541,461,613],
    [613,480,594,564,582,584,574,555,581,496,619,362,613,480],
    [744,663,733,704,725,660,747,558,744,663]
  ],
  growthGlyph: [
    [810,240,1000,240,1000,469,930,384,559,830,381,615,60,1000,10,940,381,495,559,709,880,324,810,240],
    [142,158,240,200,262,295,282,202,342,164,404,160,469,191,500,298,530,191,594,160,663,166,721,208,
     755,393,715,440,678,246,678,486,630,543,630,472,606,472,606,571,559,628,559,246,515,452,478,442,
     440,246,440,485,380,437,320,486,320,246,284,439,262,458,239,439,202,245,202,345,238,558,202,558,
     202,628,155,686,155,558,131,558,131,714,83,772,83,558,47,558,83,345,83,243,46,442,19,458,0,426,42,205,
     86,171,142,158],
    [619,0,661,21,678,72,661,122,619,143,577,122,559,72,577,21,619,0],
    [380,0,422,21,440,72,422,122,380,143,338,122,321,72,338,21,380,0],
    [143,0,185,21,202,72,185,122,143,143,101,122,83,72,101,21,143,0]
  ],
  hillGhost: [
    [0,0,73,200,125,321,177,399,226,432,250,434,299,417,477,294,591,250,670,241,747,252,818,287,849,315,
     878,351,914,421,954,555,987,780,1000,1000,0,1000,0,0]
  ],
  hillOutline: [
    [376,55,382,44,391,49,437,114,459,160,542,527,558,510,589,452,616,376,696,307,733,298,786,350,865,534,
     874,536,909,509,927,516,940,534,975,632,1000,899,996,928,970,959,949,967,788,974,631,995,557,981,
     492,1000,454,990,349,982,293,959,209,986,161,971,33,954,16,939,0,878,18,574,43,406,108,166,179,49,
     234,6,285,0,312,6,366,31,376,55],
    [971,857,958,696,936,619,912,595,876,625,855,620,777,443,748,407,733,398,718,401,642,464,566,627,
     537,636,511,584,448,264,410,175,373,138,318,105,231,106,174,147,140,196,115,252,60,445,32,618,14,838,
     17,869,32,894,43,900,76,889,126,892,141,879,218,896,294,856,328,856,364,882,385,878,410,893,409,881,
     421,878,439,890,476,883,488,899,513,900,581,862,599,883,611,874,610,882,637,876,649,885,669,872,
     707,875,714,868,717,877,760,863,771,871,793,861,810,868,837,858,862,862,875,851,914,852,930,862,
     946,849,954,858,957,847,971,857]
  ],
  hourglassGlyph: [
    [600,566,761,725,239,725,430,542,446,494,239,276,155,75,846,75,740,308,559,475,559,525,600,566],
    [952,75,1000,75,1000,0,0,0,0,75,46,75,142,308,232,418,346,500,232,582,141,692,46,925,0,925,0,1000,
     1000,1000,1000,925,952,925,857,692,766,582,652,500,766,418,857,308,952,75]
  ],
  leafCorner: [
    [0,0,674,0,1000,202,605,267,578,281,547,317,532,364,532,398,646,838,327,562,299,551,255,551,216,572,
     194,596,4,1000,0,0]
  ],
  mailGlyph: [
    [572,914,762,526,920,526,813,782,572,914],
    [80,526,238,526,428,914,187,782,80,526],
    [428,86,238,474,80,474,187,218,428,86],
    [526,526,709,526,526,888,526,526],
    [474,526,474,888,291,526,474,526],
    [526,112,709,474,526,474,526,112],
    [474,474,291,474,474,112,474,474],
    [920,474,762,474,572,86,813,218,920,474],
    [500,0,146,146,0,500,146,854,500,1000,854,854,1000,500,854,146,500,0]
  ],
  megaphoneGlyph: [
    [638,788,409,869,450,925,626,863,638,788],
    [48,746,149,936,123,999,84,982,3,809,11,758,48,746],
    [694,0,995,657,692,771,648,922,450,991,354,889,210,939,96,696,694,0]
  ],
  personGlyph: [
    [499,533,771,581,965,684,999,747,1000,1000,0,1000,0,755,35,683,251,575,499,533],
    [499,0,676,69,730,143,750,234,730,326,676,400,499,469,322,400,269,326,249,234,269,143,322,69,499,0]
  ],
  phoneGlyph: [
    [579,790,579,860,631,860,631,790,579,790],
    [474,790,474,860,526,860,526,790,474,790],
    [369,790,369,860,421,860,421,790,369,790],
    [579,649,579,719,631,719,631,649,579,649],
    [474,649,474,719,526,719,526,649,474,649],
    [369,649,369,719,421,719,421,649,369,649],
    [579,509,579,579,631,579,631,509,579,509],
    [474,509,474,579,526,579,526,509,474,509],
    [369,509,369,579,421,579,421,509,369,509],
    [330,228,369,368,631,368,639,248,698,243,710,386,887,644,893,1000,107,1000,115,637,330,228],
    [500,0,750,45,980,179,911,426,773,328,762,196,500,139,238,196,225,332,89,426,20,179,250,45,500,0]
  ],
  pinGlyph: [
    [804,274,696,274,696,202,804,202,804,274],
    [804,464,696,464,696,393,804,393,804,464],
    [804,655,696,655,696,583,804,583,804,655],
    [804,845,696,845,696,774,804,774,804,845],
    [554,274,446,274,446,202,554,202,554,274],
    [554,464,446,464,446,393,554,393,554,464],
    [554,655,446,655,446,583,554,583,554,655],
    [554,917,446,917,446,774,554,774,554,917],
    [304,274,196,274,196,202,304,202,304,274],
    [304,464,196,464,196,393,304,393,304,464],
    [304,655,196,655,196,583,304,583,304,655],
    [304,845,196,845,196,774,304,774,304,845],
    [929,917,929,119,875,119,875,48,821,0,179,0,125,119,71,119,71,917,0,917,0,1000,1000,1000,1000,917,
     929,917]
  ],
  quoteGlyph: [
    [1000,0,1000,216,913,238,841,298,792,387,772,496,1000,496,1000,1000,600,1000,600,496,633,302,686,192,
     800,67,895,18,1000,0],
    [400,0,400,216,313,238,241,298,192,387,172,496,400,496,400,1000,0,1000,0,496,33,302,86,192,200,67,
     295,18,400,0]
  ],
  rattleGlyph: [
    [180,722,111,750,84,818,111,887,180,916,249,887,278,818,249,750,180,722],
    [487,128,635,367,873,513,715,556,558,502,336,725,360,843,309,947,216,997,102,982,10,878,13,750,122,648,
     276,665,498,442,444,286,487,128],
    [741,1,880,48,976,165,996,325,944,445,729,341,593,168,554,56,741,1]
  ],
  searchGlyph: [
    [974,848,810,685,737,666,682,610,746,484,757,328,708,188,612,79,481,14,329,4,189,52,80,148,14,279,4,431,
     52,571,148,680,279,746,380,760,503,739,611,681,666,737,685,810,895,998,951,991,989,954,999,899,
     974,848],
    [379,682,261,658,164,593,99,496,75,378,99,260,164,164,261,99,379,75,497,99,594,164,659,260,683,378,
     659,496,594,593,497,658,379,682]
  ],
  shareGlyph: [
    [0,51,594,51,594,128,70,128,70,923,798,923,798,596,868,531,868,996,0,1000,0,51],
    [664,0,1000,308,664,615,664,449,483,473,293,566,143,756,250,399,447,224,664,180,664,0]
  ],
  shirtGlyph: [
    [994,264,821,74,651,0,525,70,351,1,176,77,0,284,169,437,248,371,248,1000,752,1000,752,372,832,439,
     994,264]
  ],
  sparkStar: [
    [200,320,178,249,221,221,272,222,381,269,501,359,521,212,566,86,576,103,584,58,642,12,678,0,704,14,
     730,113,708,266,664,381,762,356,876,362,977,392,1000,444,932,483,674,526,844,882,871,958,856,989,
     787,974,670,865,611,763,568,611,464,958,446,986,399,1000,350,962,348,914,420,616,131,720,50,727,
     13,703,0,667,21,622,85,572,220,518,432,470,286,395,200,320]
  ],
  squiggleCorner: [
    [454,449,440,458,386,590,371,643,362,758,372,790,391,806,407,800,416,783,449,662,460,552,454,449],
    [825,424,812,442,778,539,755,643,756,677,781,696,809,686,830,667,864,584,858,516,832,428,825,424],
    [279,0,343,20,414,70,439,109,482,218,564,123,634,104,685,100,717,107,766,136,844,224,965,85,1000,65,
     1000,213,936,270,886,344,916,490,915,610,875,763,847,822,817,859,782,867,752,854,724,809,704,738,
     702,662,720,555,746,444,783,330,709,279,666,276,557,289,516,337,518,546,502,716,447,920,426,970,
     403,997,386,999,375,990,330,906,315,857,309,795,317,660,341,533,372,435,431,310,379,218,349,216,
     293,190,260,187,229,196,171,236,87,410,24,635,7,620,1,570,29,430,73,276,109,180,150,105,184,61,231,21,
     246,21,279,0]
  ],
  targetGlyph: [
    [481,215,624,251,567,308,421,299,301,380,253,519,320,680,481,747,620,699,701,579,692,433,749,376,
     785,519,761,637,696,734,599,799,481,823,363,799,266,734,201,637,177,519,201,401,266,304,363,239,
     481,215],
    [481,38,708,94,687,171,554,120,426,118,277,169,161,272,90,412,76,519,108,676,195,805,324,892,481,924,
     588,910,728,839,831,723,882,574,879,445,829,313,905,294,962,519,924,706,821,859,668,962,481,1000,
     294,962,141,859,38,706,0,519,38,332,141,179,294,76,481,38],
    [873,0,886,114,1000,127,861,266,795,258,592,461,602,552,557,619,480,646,402,619,354,536,380,442,446,397,
     539,408,742,205,734,139,873,0]
  ],
  checkGlyph: [
    [982,223,881,125,850,117,819,125,383,559,190,366,158,351,112,363,15,459,1,488,3,520,342,868,371,882,
     403,880,985,307,1000,262,982,223]
  ],
  pieGlyph: [
    [479,104,333,141,214,227,134,351,104,500,135,654,220,780,346,865,446,892,525,895,664,860,765,794,
     479,508,479,104],
    [521,104,521,479,895,479,860,336,780,219,663,139,521,104],
    [550,521,794,765,866,651,895,521,550,521]
  ],
  dbGlyph: [
    [792,188,763,224,682,253,500,271,374,263,272,240,237,224,208,188,237,151,272,136,374,112,500,104,
     626,112,728,136,784,169,792,188],
    [708,396,688,375,708,354,729,375,708,396],
    [500,312,294,288,231,262,208,229,211,407,248,438,322,462,500,479,678,462,752,438,789,407,792,229,
     769,262,706,288,500,312],
    [708,604,688,583,708,562,729,583,708,604],
    [500,521,294,496,231,470,208,438,211,615,248,646,322,670,500,688,678,670,752,646,789,615,792,438,
     769,470,706,496,500,521],
    [708,812,688,792,708,771,729,792,708,812],
    [500,729,294,705,231,678,208,646,211,824,248,854,322,878,500,896,678,878,752,854,789,824,792,646,
     769,678,706,705,500,729]
  ],
  cogGlyph: [
    [500,625,452,615,412,588,385,548,375,500,385,452,412,412,452,385,500,375,548,385,588,412,615,452,
     625,500,615,548,588,588,548,615,500,625],
    [781,422,754,357,780,279,721,220,643,246,577,219,542,146,458,146,422,219,357,246,279,220,220,279,
     246,357,219,423,146,458,146,542,219,578,246,643,220,721,279,780,357,754,423,781,459,854,543,854,
     579,781,644,754,722,780,781,721,755,643,782,577,855,541,855,457,781,422]
  ],
  starGlyph: [
    [243,917,310,624,83,427,383,401,500,125,617,401,917,427,690,624,757,917,500,762,243,917]
  ],
};

// ---------------------------------------------------------------- helpers
let pptx, slide;

/** Draw one of the SHAPES outlines, scaled into the given box. */
function outline(name, x, y, w, h, color, opts = {}) {
  const pts = [];   // custGeom points are shape-local inches
  SHAPES[name].forEach(sub => {
    for (let i = 0; i < sub.length; i += 2) {
      pts.push({ x: sub[i] / 1000 * w, y: sub[i + 1] / 1000 * h, moveTo: i === 0 });
    }
    pts.push({ close: true });
  });
  slide.addShape(pptx.ShapeType.custGeom, shapeOpts({ x, y, w, h, points: pts }, color, opts));
}

/** Solid rectangle / rounded rectangle / ellipse shorthand. */
function box(type, x, y, w, h, color, opts = {}) {
  slide.addShape(pptx.ShapeType[type], shapeOpts({ x, y, w, h }, color, opts));
}

/** Merge geometry, fill (with optional transparency) and extra shape options. */
function shapeOpts(geo, color, opts) {
  const o = Object.assign({}, geo, opts);
  o.fill = { color };
  if (o.transparency) { o.fill.transparency = o.transparency; delete o.transparency; }
  if (!o.line) o.line = { type: 'none' };
  if (o.shadow) o.shadow = o.shadow();
  return o;
}

/** Text box. `runs` is a string or an array of [text, {overrides}] pairs. */
function text(runs, x, y, w, h, o = {}) {
  const base = {
    fontFace: o.face || BODY, fontSize: o.size || 18, bold: !!o.bold,
    italic: !!o.italic, color: o.color || C.ink,
  };  // `o` may spread one of the S presets, then override

  const body = typeof runs === 'string'
    ? [{ text: runs, options: base }]
    : runs.map(r => ({ text: r[0], options: Object.assign({}, base, r[1]) }));
  slide.addText(body, {
    x, y, w, h, wrap: o.wrap !== false,
    align: o.align || 'left', valign: o.valign || 'top',
    lineSpacingMultiple: o.line || 1,
    fit: 'none',
  });
}

/** Heading with an accent-coloured span: [['Our SWOT ', 0], ['Analysis', 1]] */
function heading(parts, x, y, w, h, o = {}) {
  text(parts.map(([t, hot]) => [t, hot ? { color: o.hot || C.amber } : {}]),
    x, y, w, h, Object.assign({ face: HEAD, size: 40, bold: true, color: C.ink }, o));
}

/** Tick bullet: coloured disc with a white check, plus a label to its right. */
function tick(x, y, d, discColor, label, tw, o = {}) {
  box('ellipse', x, y, d, d, discColor);
  outline('checkGlyph', x + d * 0.307, y + d * 0.307, d * 0.386, d * 0.386, C.white);
  text(label, x + d * 1.09, y - 0.065, tw, 0.337, Object.assign({}, S.listItem, o));
}

/** Faint doodles that repeat on most slides. */
function cloudTopLeft(flipH)  { outline('cloudTL', flipH ? 11.35 : 0, 0, 1.983, 1.21, C.smoke, { transparency: 45, flipH: !!flipH }); }
function grassBottomLeft(top) { outline('grassBump', 0, top ? 0 : 6.822, 1.238, 0.678, C.green, { flipV: !!top }); }
function squiggleTopRight()   { outline('squiggleCorner', 11.661, 0.268, 1.672, 0.555, C.smoke, { transparency: 45 }); }
function leafBottomRight(col) { outline('leafCorner', 12.095, 6.425, 1.238, 1.075, col, { flipH: true, flipV: true }); }
function star(x, y, s, color, flipH) { outline('sparkStar', x, y, 0.405 * s, 0.399 * s, color, { flipH: !!flipH }); }

// ---------------------------------------------------------------- text styles
const S = {
  body:       { face: BODY, size: 11, color: C.mute },
  bodySm:     { face: BODY, size: 10.5, color: C.mute },
  bodyW:      { face: BODY, size: 11, color: C.white },
  caption:    { face: BODY, size: 11, italic: true, color: C.mute },
  captionW:   { face: BODY, size: 11, italic: true, color: C.white },
  contact:    { face: HEAD, size: 16, color: C.gray },
  h3:         { face: HEAD, size: 18, bold: true, color: C.gray },
  h3W:        { face: HEAD, size: 18, bold: true, color: C.white },
  h4:         { face: HEAD, size: 16, bold: true, color: C.gray },
  listItem:   { face: HEAD, size: 14, color: C.gray },
  name:       { face: HEAD, size: 20, bold: true, color: C.gray },
  rating:     { face: BODY, size: 9, italic: true, color: C.mute },
  stat:       { face: HEAD, size: 36, bold: true, color: C.white },
  statLabel:  { face: BODY, size: 14, color: C.white },
};

// ---------------------------------------------------------------- slide 1
function slide01() {  // Kifa
  outline('hillGhost', 0, 4.616, 6.783, 2.884, C.smoke, { transparency: 70 });
  outline('blobRight1', 7.176, 0, 6.158, 7.5, C.pink);
  grassBottomLeft();
  cloudTopLeft();
  heading([["Ki", 0], ["fa", 1]], 0.583, 1.518, 5.083, 2.895, { size: 166 });
  text("Kid Fashion and Toys Presentation", 0.667, 4.229, 6.378, 0.572, { face: HEAD, size: 28, bold: true, wrap: false });
  star(9.418, 5.41, 1, C.amber);
}

// ---------------------------------------------------------------- slide 2
function slide02() {  // Table Of Content
  outline('blobLeft1', 0.333, 0, 6.119, 7.5, C.pink, { flipH: true });
  leafBottomRight(C.green);
  squiggleTopRight();
  text("3\t\tAbout Us", 8.09, 2.729, 3.423, 0.404, { face: HEAD, size: 18, bold: true, color: C.pink });
  text("12\t\tFashion identity", 8.09, 3.29, 4.223, 0.404, { face: HEAD, size: 18 });
  text("18\t\tProduct Quality", 8.09, 3.852, 4.223, 0.404, { face: HEAD, size: 18 });
  text("22\t\tToys Portfolio", 8.09, 4.413, 4.223, 0.404, { face: HEAD, size: 18 });
  text("27\t\tSafety and Design", 8.09, 4.975, 4.598, 0.404, { face: HEAD, size: 18 });
  text("28\t\tBrand Strategy", 8.09, 5.537, 4.598, 0.404, { face: HEAD, size: 18 });
  text("35\t\tGet In Touch", 8.09, 6.098, 3.86, 0.404, { face: HEAD, size: 18 });
  outline('arrowGlyph', 7.307, 2.743, 0.183, 0.376, C.pink, { rotate: 270 });
  heading([["Table Of ", 0], ["Content", 1]], 6.97, 1.396, 5.083, 0.842, { size: 44 });
}

// ---------------------------------------------------------------- slide 3
function slide03() {  // Welcome to Kifa Kid Fashion & Toys
  outline('blobRound', 8.335, 2.046, 4.999, 5.454, C.pink);
  heading([["Welcome to Kifa ", 0], ["Kid Fashion ", 1], ["& Toys", 0]], 0.562, 1.809, 5.083, 1.447);
  text("Welcome to Kifa, where imagination meets style! Discover a world of vibrant Kid Fashion & Toys that blend comfort with trendsetting designs. Embrace the joy of childhood with Kifa's playful creations.", 0.569, 3.849, 4.807, 1.284, { size: 12, color: C.mute, line: 1.5 });
  outline('cardTab', 6.667, 4.763, 6.01, 1.994, C.white, { shadow: CARD_SHADOW });
  outline('greenSquiggle', 6.052, 1.679, 1.23, 0.22, C.green);
  grassBottomLeft();
  cloudTopLeft();
  star(5.466, 6.12, 1, C.amber);
  slide.addShape(pptx.ShapeType.line, { x: 0.664, y: 5.5, w: 0.667, h: 0, line: { color: C.amber, width: 2, endArrowType: 'arrow' } });
  box('ellipse', 7.354, 5.285, 0.909, 0.909, C.amber);
  outline('chatGlyph', 7.611, 5.559, 0.396, 0.361, C.white);
  text("Where fashion meets fun! Explore our world of trendy kid fashion and delightful toys!", 8.518, 5.285, 3.657, 0.909, { face: HEAD, size: 16, bold: true, italic: true });
}

// ---------------------------------------------------------------- slide 4
function slide04() {  // We Have Positive Goal for Kid’s Needs
  outline('blobLeft2', 0, 0, 5.921, 7.5, C.pink);
  box('roundRect', 0.667, 0.885, 4.021, 2.669, C.white, { shadow: CARD_SHADOW, rectRadius: 0.278 });
  box('roundRect', 2.646, 3.946, 4.021, 2.669, C.white, { shadow: CARD_SHADOW, rectRadius: 0.278 });
  leafBottomRight(C.amber);
  heading([["We Have ", 0], ["Positive Goal ", 1], ["for Kid’s Needs", 0]], 7.271, 1.476, 5.604, 1.447);
  squiggleTopRight();
  text("Holistic Development", 7.271, 3.255, 2.367, 0.37, { ...S.h4, wrap: false });
  text("We Always Prioritizing kid's needs by offering fashion and toys that contribute to their holistic growth and well-being.", 7.285, 3.869, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 7.277, 3.545, 2.032, 0.307, { ...S.rating, line: 1.5 });
  text("Joyful Learning Experience", 7.271, 4.885, 2.998, 0.37, { ...S.h4, wrap: false });
  text("Creating a positive goal centered on providing a joyful learning experience through innovative, educational toys and stylish", 7.285, 5.499, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 7.277, 5.174, 2.032, 0.307, { ...S.rating, line: 1.5 });
  text("Creative and Smart Product", 0.593, 5.422, 1.796, 0.572, S.statLabel);
  text("100%", 0.593, 4.675, 1.796, 0.707, S.stat);
  outline('greenScribble', 4.289, 1.996, 1.038, 1.096, C.green, { rotate: 332.33 });
}

// ---------------------------------------------------------------- slide 5
function slide05() {  // We Bring The Best Quality Product
  outline('hillOutline', 7.628, 0.635, 3.974, 1.941, C.pink);
  outline('greenSquiggle', 6.413, 1.751, 0.968, 0.173, C.green);
  grassBottomLeft();
  cloudTopLeft();
  heading([["We Bring The Best ", 0], ["Quality ", 1], ["Product", 0]], 0.562, 1.699, 5.603, 1.447);
  box('ellipse', 0.653, 3.898, 0.322, 0.322, C.amber);
  outline('checkGlyph', 0.752, 3.997, 0.124, 0.124, C.white);
  text("Kifa is dedicated to delivering the best quality kid's fashion and toys, marked by exceptional craftsmanship and attention to detail.", 1.084, 3.792, 4.603, 0.902, { ...S.body, line: 1.5 });
  box('ellipse', 0.653, 5.015, 0.322, 0.322, C.amber);
  outline('checkGlyph', 0.752, 5.114, 0.124, 0.124, C.white);
  text("Prioritizing children's safety, Kifa products undergo rigorous testing and adhere to the highest safety standards for worry-free enjoyment.", 1.084, 4.91, 4.603, 0.902, { ...S.body, line: 1.5 });
  star(8.893, 5.682, 1.785, C.amber);
}

// ---------------------------------------------------------------- slide 6
function slide06() {  // Our Mission is to Give Joy for Kids
  squiggleTopRight();
  box('roundRect', 3.688, 3.417, 9, 3.354, C.pink, { shadow: CARD_SHADOW, rectRadius: 0.309 });
  outline('greenScribble', 4.247, 1.975, 1.038, 1.096, C.green, { rotate: 332.33 });
  heading([["Our Mission is to ", 0], ["Give Joy ", 1], ["for Kids", 0]], 6.227, 1.401, 4.981, 1.447);
  text("Infuse joy into children's lives through imaginative and stylish kid's fashion and toys, fostering creativity and endless smiles.", 4.291, 5.467, 3.725, 0.656, S.bodyW);
  text("Creative Delight", 4.284, 5.03, 3.308, 0.404, S.h3W);
  box('ellipse', 4.346, 4.051, 0.821, 0.821, C.white, { transparency: 75 });
  text("Dedicated to bringing joy, Kifa prioritizes playful innovation in every product, ensuring a world of happiness for the little ones.", 8.431, 5.467, 3.725, 0.656, S.bodyW);
  text("Playful Innovation", 8.424, 5.03, 3.308, 0.404, S.h3W);
  box('ellipse', 8.486, 4.051, 0.821, 0.821, C.white, { transparency: 75 });
  outline('duckGlyph', 8.682, 4.295, 0.429, 0.332, C.white);
  outline('rattleGlyph', 4.564, 4.268, 0.385, 0.385, C.white);
}

// ---------------------------------------------------------------- slide 7
function slide07() {  // About Kifa Philosophy Here
  outline('blobRight2', 6.667, 0, 6.667, 7.5, C.pink);
  box('roundRect', 5.688, 5, 7, 1.75, C.white, { shadow: CARD_SHADOW, rectRadius: 0.233 });
  grassBottomLeft();
  cloudTopLeft();
  heading([["About Kifa ", 0], ["Philosophy", 1], [" Here", 0]], 0.542, 1.538, 4.75, 1.447);
  text("Discover Kifa's philosophy—a realm where vibrant creativity intertwines with comfort. We believe in nurturing every child's uniqueness through playful, stylish fashion and imaginative, joyous toys.", 0.552, 3.841, 4.615, 0.842, S.body);
  text("General Descriptions", 0.546, 3.42, 3.032, 0.404, S.h3);
  tick(0.663, 5.043, 0.21, C.amber, "Playful Creativity Unleashed", 3.267);
  tick(0.663, 5.461, 0.21, C.amber, "Comfort Meets Style at Kifa", 3.267);
  tick(0.663, 5.879, 0.21, C.amber, "Nurturing Uniqueness with Joy", 3.267);
  heading([["Where fashion meets fun! ", 0], ["Explore our world of ", 1], ["trendy kid fashion and delightful toys!", 0]], 6.218, 5.555, 5.938, 0.639, { size: 16, italic: true, align: 'center', hot: C.green });
}

// ---------------------------------------------------------------- slide 8
function slide08() {  // Fashion Trend at January 2024
  outline('cloudPink', 0, 4.253, 4.992, 3.247, C.pink);
  leafBottomRight(C.green);
  heading([["Fashion Trend ", 0], ["at January 2024", 1]], 7.562, 1.434, 4.75, 1.447, { color: C.amber, hot: C.ink });
  squiggleTopRight();
  text("Trendsetting Styles now", 8.316, 3.59, 3.345, 0.404, S.h3);
  text("5 Stars Rating at 2024", 8.33, 3.932, 2.723, 0.352, { ...S.caption, line: 1.5 });
  box('ellipse', 7.652, 3.659, 0.588, 0.588, C.amber);
  text("Explore Kifa's latest kid's fashion collection, showcasing cutting-edge designs and vibrant colors that capture the essence of current fashion trends for January 2024. Elevate your child's style with Kifa!", 7.588, 4.425, 5, 1.185, { ...S.body, line: 1.5 });
  outline('shirtGlyph', 7.813, 3.836, 0.265, 0.235, C.white);
  star(4.455, 5.394, 1.785, C.green);
  outline('greenScribble', 1.262, 2.327, 1.038, 1.096, C.green, { rotate: 332.33 });
}

// ---------------------------------------------------------------- slide 9
function slide09() {  // Kifa Product Characteristic
  box('roundRect', 0.688, 2.208, 3.762, 4.5, C.white, { shadow: CARD_SHADOW, rectRadius: 0.237 });
  box('roundRect', 4.786, 2.208, 3.762, 4.5, C.white, { shadow: CARD_SHADOW, rectRadius: 0.253 });
  box('roundRect', 8.884, 2.208, 3.762, 4.5, C.white, { shadow: CARD_SHADOW, rectRadius: 0.253 });
  outline('cardFooter', 0.688, 6.396, 3.762, 0.312, C.amber, { shadow: CARD_SHADOW });
  outline('cardFooter', 4.786, 6.396, 3.762, 0.312, C.pink, { shadow: CARD_SHADOW });
  outline('cardFooter', 8.884, 6.396, 3.762, 0.312, C.green, { shadow: CARD_SHADOW });
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Kifa Product ", 0], ["Characteristic", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  text("Kifa products feature innovative designs that inspire creativity and captivate the imagination of children.", 0.846, 5.46, 3.444, 0.656, { ...S.body, align: 'center' });
  text("Innovative Designs", 1.053, 5.029, 3.032, 0.404, { ...S.h3, align: 'center' });
  outline('starGlyph', 1.985, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 2.218, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 2.451, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 2.684, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 2.917, 4.736, 0.235, 0.235, C.amber, { transparency: 60 });
  text("Experience the perfect blend of comfort and elegance in Kifa's kid's fashion and toys, prioritizing both style and ease.", 4.944, 5.46, 3.444, 0.656, { ...S.body, align: 'center' });
  text("Comfortable Elegance", 5.151, 5.029, 3.032, 0.404, { ...S.h3, align: 'center' });
  outline('starGlyph', 6.083, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 6.316, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 6.549, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 6.782, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 7.015, 4.736, 0.235, 0.235, C.amber);
  text("Kifa products embody educational playfulness, encouraging learning through interactive and engaging toy designs.", 9.043, 5.46, 3.444, 0.656, { ...S.body, align: 'center' });
  text("Educational Playfulness", 9.249, 5.029, 3.032, 0.404, { ...S.h3, align: 'center' });
  outline('starGlyph', 10.181, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 10.415, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 10.648, 4.736, 0.235, 0.235, C.amber);
  outline('starGlyph', 10.88, 4.736, 0.235, 0.235, C.amber, { transparency: 60 });
  outline('starGlyph', 11.113, 4.736, 0.235, 0.235, C.amber, { transparency: 60 });
}

// ---------------------------------------------------------------- slide 10
function slide10() {  // Always Use of Quality Materials
  box('ellipse', 1.156, 3, 2.354, 2.354, C.pink, { flipH: true });
  star(0.744, 1.589, 1, C.amber, true);
  leafBottomRight(C.amber);
  squiggleTopRight();
  heading([["Always Use of ", 0], ["Quality Materials", 1]], 6.726, 1.33, 5.604, 1.447);
  text("Premium Craftsmanship", 7.244, 3.276, 2.63, 0.37, { ...S.h4, wrap: false });
  text("Kifa ensures superior quality by employing premium materials, guaranteeing durability and safety in every kid's fashion.", 7.258, 3.874, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 7.251, 3.55, 2.032, 0.307, { ...S.rating, line: 1.5 });
  box('ellipse', 6.823, 3.363, 0.322, 0.322, C.green);
  outline('checkGlyph', 6.922, 3.462, 0.124, 0.124, C.white);
  text("Safety-First Assurance", 7.244, 4.867, 2.588, 0.37, { ...S.h4, wrap: false });
  text("Kifa prioritizes a safety-first approach, rigorously testing and adhering to the highest standards in all products.", 7.258, 5.464, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 7.251, 5.14, 2.032, 0.307, { ...S.rating, line: 1.5 });
  box('ellipse', 6.823, 4.953, 0.322, 0.322, C.green);
  outline('checkGlyph', 6.922, 5.052, 0.124, 0.124, C.white);
  outline('greenSquiggle', 4.639, 3.75, 1.23, 0.22, C.green, { flipH: true });
}

// ---------------------------------------------------------------- slide 11
function slide11() {  // Seasonal Collections and Special Designs
  outline('blobWave', 4.532, 2.854, 8.802, 4.646, C.pink, { flipH: true });
  grassBottomLeft();
  cloudTopLeft();
  star(6.256, 5.244, 1, C.amber);
  heading([["Seasonal Collections and ", 0], ["Special Designs", 1]], 0.833, 1.455, 5.521, 1.447);
  text("Discover Kifa's enchanting Seasonal Collections and Special Designs, curated to bring joy and style with unique, playful elements in kid's fashion and toys throughout the year", 0.844, 3.799, 5.177, 0.656, S.body);
  text("General Descriptions", 0.838, 3.336, 3.032, 0.404, S.h3);
  tick(0.954, 4.748, 0.21, C.amber, "Year-Round Playful Styles", 3.267);
  tick(0.954, 5.165, 0.21, C.amber, "Special Edition Delights", 3.267);
  tick(0.954, 5.583, 0.21, C.amber, "Fashion and Toys, Seasonally Styled", 3.511);
  outline('greenSquiggle', 7.062, 1.687, 1.23, 0.22, C.green);
}

// ---------------------------------------------------------------- slide 12
function slide12() {  // Our Points About Sustainability
  box('roundRect', 6.441, 3.75, 3.021, 3.021, C.white, { shadow: CARD_SHADOW, rectRadius: 0.278 });
  text("Kifa prioritizes sustainability through eco-friendly practices in kid's fashion and toys.", 6.737, 5.774, 2.429, 0.606, { size: 10, color: C.mute, align: 'center' });
  text("Eco-Friendly Initiatives", 6.734, 5.38, 2.436, 0.337, { face: HEAD, size: 14, bold: true, color: C.gray, align: 'center' });
  box('ellipse', 7.468, 4.12, 0.969, 0.969, C.pink);
  squiggleTopRight();
  star(0.473, 0.74, 1, C.amber, true);
  box('roundRect', 9.667, 3.75, 3.021, 3.021, C.pink, { rectRadius: 0.278 });
  text("Kifa's commitment to responsible and thoughtful production.", 9.962, 5.774, 2.429, 0.606, { size: 10, color: C.white, align: 'center' });
  text("Thoughtful Production", 9.959, 5.38, 2.436, 0.337, { face: HEAD, size: 14, bold: true, color: C.white, align: 'center' });
  box('ellipse', 10.693, 4.12, 0.969, 0.969, C.white, { transparency: 75 });
  heading([["Our Points About ", 0], ["Sustainability", 1]], 6.513, 1.46, 4.75, 1.447);
  outline('ballGlyph', 7.728, 4.381, 0.448, 0.448, C.white);
  outline('balloonGlyph', 11.04, 4.301, 0.275, 0.608, C.white);
  outline('greenSquiggle', 4.575, 2.771, 1.23, 0.22, C.green, { flipH: true });
}

// ---------------------------------------------------------------- slide 13
function slide13() {  // Our Focus is On Quality & Durability
  box('roundRect', 6.789, 3.871, 5.867, 2.905, C.pink, { rectRadius: 0.262 });
  box('roundRect', 6.789, 0.724, 5.867, 2.905, C.white, { shadow: CARD_SHADOW, rectRadius: 0.262 });
  heading([["Our Focus is On ", 0], ["Quality & Durability", 1]], 0.812, 4.455, 5.669, 1.447);
  text("Emphasizing durability, Kifa products are crafted to withstand the rigors of play, ensuring a lasting and delightful experience for children.", 8.531, 5.07, 3.823, 0.908, { ...S.bodyW, line: 1.5 });
  text("Durable Delight", 8.545, 4.592, 3.308, 0.404, S.h3W);
  box('ellipse', 7.23, 4.598, 1.127, 1.127, C.white, { transparency: 75 });
  text("Kifa's commitment to excellence is reflected in the meticulous use of high-quality materials for both durable kid's fashion and toys.", 8.531, 1.923, 3.823, 0.908, { ...S.body, line: 1.5 });
  text("Quality Materials", 8.545, 1.445, 3.308, 0.404, S.h3);
  box('ellipse', 7.23, 1.451, 1.127, 1.127, C.pink);
  outline('familyGlyph', 7.484, 1.836, 0.619, 0.357, C.white);
  outline('searchGlyph', 7.55, 4.917, 0.489, 0.489, C.white);
}

// ---------------------------------------------------------------- slide 14
function slide14() {  // Always Use of Quality Materials
  box('ellipse', 10.333, 3.338, 2.62, 2.62, C.pink, { flipH: true });
  grassBottomLeft();
  cloudTopLeft();
  outline('greenSquiggle', 6.196, 1.5, 1.23, 0.22, C.green);
  star(12.262, 6.423, 1, C.amber);
  heading([["Always Use of ", 0], ["Quality Materials", 1]], 0.752, 1.434, 5.604, 1.447);
  text("Premier Craftsmanship", 1.27, 3.36, 2.541, 0.37, { ...S.h4, wrap: false });
  text("Kifa upholds a commitment to excellence through the exclusive use of premium materials in crafting superior-quality", 1.284, 3.957, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 1.277, 3.633, 2.032, 0.307, { ...S.rating, line: 1.5 });
  box('ellipse', 0.849, 3.446, 0.322, 0.322, C.green);
  outline('checkGlyph', 0.948, 3.545, 0.124, 0.124, C.white);
  text("Safety-First Assurance", 1.27, 4.95, 2.588, 0.37, { ...S.h4, wrap: false });
  text("With a focus on safety, Kifa maintains stringent quality standards, ensuring the consistent use of top-tier materials", 1.284, 5.548, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  text("5 Stars Rating at 2024", 1.277, 5.223, 2.032, 0.307, { ...S.rating, line: 1.5 });
  box('ellipse', 0.849, 5.037, 0.322, 0.322, C.green);
  outline('checkGlyph', 0.948, 5.136, 0.124, 0.124, C.white);
  text("Loyal Customers", 10.642, 4.87, 2.001, 0.337, { ...S.statLabel, align: 'center' });
  text("289+", 10.745, 4.102, 1.796, 0.707, { ...S.stat, align: 'center' });
}

// ---------------------------------------------------------------- slide 15
function slide15() {  // Examples of Product Designs
  box('roundRect', 6.772, 2.198, 2.842, 3.302, C.pink, { rectRadius: 0.256 });
  box('ellipse', 7.546, 2.813, 1.293, 1.293, C.white, { transparency: 80, flipH: true });
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Examples of ", 0], ["Product Designs", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  text("Experience the magic through a glimpse of Kifa's product designs—where innovation meets whimsy in a symphony of captivating styles, creating enchanting moments for children in both fashion and toys.", 1.129, 5.939, 11.076, 0.63, { ...S.body, align: 'center', line: 1.5 });
  text("Stay With Us and Creative Charm Unveiled", 6.975, 4.348, 2.436, 0.572, { face: HEAD, size: 14, bold: true, color: C.white, align: 'center' });
  outline('growthGlyph', 7.853, 3.177, 0.679, 0.565, C.white);
}

// ---------------------------------------------------------------- slide 16
function slide16() {  // We Collaboration With Influencers
  outline('blobLeft4', 0, 0, 7.536, 7.5, C.pink, { flipH: true });
  squiggleTopRight();
  star(11.563, 6.748, 1, C.amber);
  text("We Collaboration With Influencers", 0.562, 1.309, 5.25, 1.447, { face: HEAD, size: 40, bold: true, color: C.white });
  box('ellipse', 0.674, 3.955, 0.322, 0.322, C.amber);
  outline('checkGlyph', 0.773, 4.054, 0.124, 0.124, C.white);
  text("Engaging with influencers amplifies Kifa's reach, showcasing our kid's fashion and toys through trusted voices.", 1.147, 3.828, 4.082, 0.902, { ...S.bodyW, line: 1.5 });
  box('ellipse', 0.674, 5.348, 0.322, 0.322, C.amber);
  outline('checkGlyph', 0.773, 5.447, 0.124, 0.124, C.white);
  text("Strategic collaborations bring authenticity, spotlighting Kifa's commitment to style, comfort, and joyful play experiences.", 1.147, 5.222, 4.082, 0.908, { ...S.bodyW, line: 1.5 });
  text("Kids Story Teller Influencer", 6.814, 2.734, 2.429, 0.286, { ...S.caption, align: 'center' });
  text("Thomas Alpha Dune", 6.697, 2.37, 2.664, 0.404, { ...S.h3, align: 'center' });
  box('ellipse', 7.545, 1.316, 0.969, 0.969, C.amber);
  text("Kids Education Influencer", 9.884, 2.734, 2.429, 0.286, { ...S.caption, align: 'center' });
  text("Ruddy Alpha Dune", 9.767, 2.37, 2.664, 0.404, { ...S.h3, align: 'center' });
  box('ellipse', 10.614, 1.316, 0.969, 0.969, C.amber);
  outline('personGlyph', 7.843, 1.602, 0.371, 0.397, C.white);
  outline('personGlyph', 10.913, 1.602, 0.371, 0.397, C.white);
}

// ---------------------------------------------------------------- slide 17
function slide17() {  // Customers Experience
  box('roundRect', 0.661, 3.375, 5.86, 2.898, C.pink, { rectRadius: 0.337 });
  box('roundRect', 6.812, 3.375, 5.86, 2.898, C.white, { shadow: CARD_SHADOW, rectRadius: 0.358 });
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Customers ", 0], ["Experience", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  text("Embark on a delightful customer experience with Kifa, where vibrant kid's fashion and toys intertwine to create moments of joy, comfort, and lasting memories for children and their families.", 1.54, 2.232, 10.254, 0.63, { ...S.body, align: 'center', line: 1.5 });
  text("Jessy Bee, Jan 2024", 2.986, 4.08, 2.23, 0.37, { face: HEAD, size: 16, bold: true, color: C.white, wrap: false });
  text("Kifa's fashion is a hit! Stylish and comfortable—my child loves it. The toys are a joy, educational and durable!", 2.986, 4.782, 3.29, 0.871, { size: 10.5, color: C.white, line: 1.5 });
  text("Fashion Enthusiast", 2.993, 4.375, 2.032, 0.307, { size: 9, color: C.white, line: 1.5 });
  text("Norman Pole, Sep 2024", 9.138, 4.08, 2.521, 0.37, { ...S.h4, wrap: false });
  text("Quality products, trendy fashion, and engaging toys. Highly recommended for happy, playful kids.", 9.138, 4.782, 3.29, 0.871, { ...S.bodySm, line: 1.5 });
  text("Toys Enthusiast", 9.145, 4.375, 2.032, 0.307, { size: 9, color: C.gray, line: 1.5 });
}

// ---------------------------------------------------------------- slide 18
function slide18() {  // Break Time
  box('round2SameRect', 3.636, 4.729, 7.03, 2.771, C.pink, { rectRadius: 0.37 });
  heading([["Break ", 0], ["Time", 1]], 0.543, 1.803, 6.494, 1.582, { size: 88, wrap: false });
  text("It’s Time To Break Now", 0.57, 3.432, 4.699, 0.572, { size: 28, bold: true, color: C.gray, wrap: false });
  grassBottomLeft();
  cloudTopLeft();
  text("Break Time Estimation", 4.316, 5.922, 2.49, 0.37, { face: HEAD, size: 16, bold: true, color: C.white, wrap: false });
  text("30’", 4.316, 5.128, 1.262, 0.774, { face: HEAD, size: 40, bold: true, color: C.white });
  star(6.911, 0.695, 1, C.pink, true);
}

// ---------------------------------------------------------------- slide 19
function slide19() {  // About Kifa’s Toy Categories
  outline('blobLeft3', 0, 0, 6.4, 7.5, C.pink);
  leafBottomRight(C.amber);
  squiggleTopRight();
  heading([["About Kifa’s Toy ", 0], ["Categories", 1]], 8.271, 1.538, 4.75, 1.447);
  text("Explore Kifa's diverse toy categories, offering a spectrum of engaging, educational, and imaginative play experiences for every child's joy and development.", 8.282, 3.841, 4.615, 0.656, S.body);
  text("General Descriptions", 8.275, 3.42, 3.032, 0.404, S.h3);
  tick(8.392, 4.852, 0.21, C.amber, "Educational Play Essentials", 3.267);
  tick(8.392, 5.269, 0.21, C.amber, "Imaginative Adventure Sets", 3.267);
  tick(8.392, 5.687, 0.21, C.amber, "Quality Toys, Endless Joy", 3.267);
}

// ---------------------------------------------------------------- slide 20
function slide20() {  // Toys Appropriate To The Kid’s Age
  box('rect', 6.667, 0, 6.667, 7.5, C.pink);
  grassBottomLeft();
  cloudTopLeft();
  heading([["Toys Appropriate To ", 0], ["The Kid’s ", 1], ["Age", 0]], 0.562, 1.809, 5.083, 1.447);
  text("Kifa's commitment shines through age-appropriate toys, carefully crafted to nurture each child's developmental milestones, creating a magical and tailored play journey from infancy to adolescence.", 0.569, 3.87, 4.807, 1.179, { ...S.body, line: 1.5 });
  slide.addShape(pptx.ShapeType.line, { x: 0.664, y: 5.41, w: 0.667, h: 0, line: { color: C.amber, width: 2, endArrowType: 'arrow' } });
  text("Kifa's toys suit each age", 7.41, 6.444, 2.429, 0.286, { ...S.captionW, align: 'center' });
  text("Age-Tailored Toys", 7.407, 6.066, 2.436, 0.404, { ...S.h3W, align: 'center' });
  box('ellipse', 8.141, 5.023, 0.969, 0.969, C.white, { transparency: 70 });
  text("Kifa's toys fit kids perfectly", 10.16, 6.444, 2.429, 0.286, { ...S.captionW, align: 'center' });
  text("Playful Learning", 10.157, 6.066, 2.436, 0.404, { ...S.h3W, align: 'center' });
  box('ellipse', 10.891, 5.023, 0.969, 0.969, C.white, { transparency: 70 });
  outline('hourglassGlyph', 11.207, 5.268, 0.336, 0.48, C.white);
  outline('gearGlyph', 8.421, 5.303, 0.408, 0.408, C.white);
}

// ---------------------------------------------------------------- slide 21
function slide21() {  // Design Innovation In Kifa’s Store
  box('roundRect', 4.656, 3.438, 8.01, 3.019, C.pink, { shadow: CARD_SHADOW, rectRadius: 0.372 });
  squiggleTopRight();
  outline('greenSquiggle', 4.967, 1.51, 1.23, 0.22, C.green);
  star(0.667, 6.058, 1, C.amber);
  heading([["Design ", 0], ["Innovation", 1], [" In Kifa’s Store", 0]], 7.083, 1.518, 5.083, 1.447);
  text("We boasts design innovations, creating an immersive and delightful shopping experience for customers.", 5.27, 5.327, 3.162, 0.656, S.bodyW);
  text("Innovative Store Layout", 5.263, 4.891, 3.169, 0.404, S.h3W);
  box('ellipse', 5.326, 3.911, 0.821, 0.821, C.white, { transparency: 75 });
  text("Kifa's design innovation extends to creative and engaging product displays, enhancing the overall shopping.", 9.003, 5.327, 3.425, 0.656, S.bodyW);
  text("Best Product Displays", 8.996, 4.891, 3.169, 0.404, S.h3W);
  box('ellipse', 9.059, 3.911, 0.821, 0.821, C.white, { transparency: 75 });
  outline('targetGlyph', 5.541, 4.126, 0.39, 0.39, C.white);
  outline('megaphoneGlyph', 9.259, 4.15, 0.421, 0.341, C.white);
}

// ---------------------------------------------------------------- slide 22
function slide22() {  // Stimulate Kid's Development
  box('roundRect', 8.688, 2.305, 4, 2.89, C.pink, { rectRadius: 0.336 });
  grassBottomLeft();
  cloudTopLeft();
  outline('greenSquiggle', 5.257, 1.429, 1.23, 0.22, C.green);
  star(9.985, 5.783, 1, C.amber);
  heading([["Stimulate Kid's ", 0], ["Development", 1]], 0.611, 1.58, 4.646, 1.447);
  text("Kifa's toys and fashion stimulate holistic development, fostering creativity, cognitive skills, and emotional well-being in every child's joyful journey of growth.", 0.622, 3.841, 4.615, 0.656, S.body);
  text("General Descriptions", 0.615, 3.42, 3.032, 0.404, S.h3);
  tick(0.732, 4.821, 0.21, C.amber, "Inspiring Cognitive Growth", 3.267);
  tick(0.732, 5.27, 0.21, C.amber, "Encouraging Creative Exploration", 3.267);
  tick(0.732, 5.72, 0.21, C.amber, "Holistic Development Sparked", 3.267);
  box('ellipse', 10.183, 2.727, 1.008, 1.008, C.white, { transparency: 80, flipH: true });
  outline('chatGlyph', 10.481, 3.043, 0.413, 0.377, C.white);
  text("Where fashion meets fun! Explore our world of trendy kid fashion and delightful toys!", 8.941, 3.917, 3.493, 0.808, { face: HEAD, size: 14, bold: true, italic: true, color: C.white, align: 'center' });
}

// ---------------------------------------------------------------- slide 23
function slide23() {  // Reliability and Durability of Kifa toys
  box('roundRect', 7.417, 0.771, 5.243, 1.861, C.white, { shadow: CARD_SHADOW, rectRadius: 0.229 });
  box('roundRect', 7.417, 2.819, 5.243, 1.861, C.pink, { rectRadius: 0.229 });
  box('roundRect', 7.417, 4.868, 5.243, 1.861, C.white, { shadow: CARD_SHADOW, rectRadius: 0.229 });
  heading([["Reliability and ", 0], ["Durability", 1], [" of Kifa toys", 0]], 0.611, 1.538, 6.056, 1.447);
  cloudTopLeft();
  text("Kifa guarantees both reliability and durability, offering toys that withstand the test of imaginative and creative play.", 9.009, 3.661, 3.422, 0.656, S.bodyW);
  text("Quality Assured Creativity", 9.009, 3.183, 3.308, 0.404, S.h3W);
  box('ellipse', 7.856, 3.275, 0.955, 0.955, C.white, { transparency: 75 });
  text("Count on Kifa for reliable and durable toys, creating dependable and joyful moments for children.", 9.009, 5.709, 3.422, 0.656, S.body);
  text("Dependable Joyful", 9.009, 5.231, 3.308, 0.404, S.h3);
  box('ellipse', 7.856, 5.324, 0.955, 0.955, C.pink);
  text("Kifa toys are crafted with utmost reliability, ensuring long-lasting companionship for children's playtime.", 9.009, 1.613, 3.422, 0.656, S.body);
  text("Durable Play Companions", 9.009, 1.134, 3.308, 0.404, S.h3);
  box('ellipse', 7.856, 1.227, 0.955, 0.955, C.pink);
  outline('shareGlyph', 8.191, 1.559, 0.317, 0.29, C.white);
  outline('shareGlyph', 8.191, 3.607, 0.317, 0.29, C.white);
  outline('shareGlyph', 8.191, 5.656, 0.317, 0.29, C.white);
}

// ---------------------------------------------------------------- slide 24
function slide24() {  // Important Things in Business
  box('ellipse', 2.587, 2.443, 2.163, 2.157, C.pink);
  text("Analysis", 2.769, 3.587, 1.798, 0.38, { ...S.h3W, valign: 'middle' });
  outline('pieGlyph', 3.451, 3.077, 0.418, 0.418, C.white);
  box('ellipse', 1.507, 4.31, 2.163, 2.157, C.green);
  text("Finance", 1.689, 5.48, 1.798, 0.38, { ...S.h3W, valign: 'middle' });
  outline('dbGlyph', 2.397, 4.917, 0.418, 0.418, C.white);
  box('ellipse', 3.667, 4.31, 2.163, 2.157, C.amber);
  text("Options", 3.85, 5.46, 1.798, 0.38, { ...S.h3W, valign: 'middle' });
  outline('cogGlyph', 4.54, 4.938, 0.418, 0.418, C.white);
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Important Things ", 0], ["in Business", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  text("Our Business Strategy", 7.166, 2.673, 2.47, 0.37, { ...S.h4, wrap: false });
  text("Kifa ensures superior quality by employing premium materials, guaranteeing durability and safety in every kid's fashion.", 7.166, 2.986, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  box('ellipse', 6.761, 2.755, 0.322, 0.322, C.pink);
  outline('checkGlyph', 6.86, 2.853, 0.124, 0.124, C.white);
  text("Our Finance Strategy", 7.166, 4.16, 2.4, 0.37, { ...S.h4, wrap: false });
  text("Kifa ensures superior quality by employing premium materials, guaranteeing durability and safety in every kid's fashion.", 7.166, 4.473, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  box('ellipse', 6.761, 4.241, 0.322, 0.322, C.green);
  outline('checkGlyph', 6.86, 4.34, 0.124, 0.124, C.white);
  text("Optional Things", 7.166, 5.646, 1.811, 0.37, { ...S.h4, wrap: false });
  text("Kifa ensures superior quality by employing premium materials, guaranteeing durability and safety in every kid's fashion.", 7.166, 5.959, 4.865, 0.606, { ...S.bodySm, line: 1.5 });
  box('ellipse', 6.761, 5.728, 0.322, 0.322, C.amber);
  outline('checkGlyph', 6.86, 5.827, 0.124, 0.124, C.white);
}

// ---------------------------------------------------------------- slide 25
function slide25() {  // The Core Message Want To Convey
  outline('greenSquiggle', 6.28, 1.187, 1.23, 0.22, C.green);
  star(5.792, 5.601, 1, C.amber);
  leafBottomRight(C.amber);
  squiggleTopRight();
  heading([["The Core Message Want ", 0], ["To Convey", 1]], 7.542, 1.78, 5.083, 1.447);
  text("At Kifa, our core message is simple yet profound—nurturing the spirit of childhood through vibrant fashion and enchanting toys. We believe in fostering not just joy but also creativity, contributing to the wholesome growth of every child. In each design and every play experience.", 7.549, 3.945, 4.807, 1.463, { ...S.body, line: 1.5 });
  slide.addShape(pptx.ShapeType.line, { x: 7.643, y: 5.659, w: 0.667, h: 0, line: { color: C.amber, width: 2, endArrowType: 'arrow' } });
  box('ellipse', 0.381, 1.78, 2.62, 2.62, C.pink, { flipH: true });
  text("Loyal Customers", 0.691, 3.311, 2.001, 0.337, { ...S.statLabel, align: 'center' });
  text("289+", 0.793, 2.543, 1.796, 0.707, { ...S.stat, align: 'center' });
}

// ---------------------------------------------------------------- slide 26
function slide26() {  // Our SWOT Analysis
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Our SWOT ", 0], ["Analysis", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  slide.addText("W", { shape: pptx.ShapeType.roundRect, x: 6.792, y: 2.438, w: 1.833, h: 1.833, fill: { color: C.amber }, line: { type: 'none' }, rectRadius: 0.306, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 60, bold: true, color: C.white });
  slide.addText("S", { shape: pptx.ShapeType.roundRect, x: 4.708, y: 2.438, w: 1.833, h: 1.833, fill: { color: C.pink }, line: { type: 'none' }, rectRadius: 0.306, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 60, bold: true, color: C.white });
  slide.addText("O", { shape: pptx.ShapeType.roundRect, x: 6.792, y: 4.521, w: 1.833, h: 1.833, fill: { color: C.pink }, line: { type: 'none' }, flipH: true, rectRadius: 0.306, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 60, bold: true, color: C.white });
  slide.addText("T", { shape: pptx.ShapeType.roundRect, x: 4.708, y: 4.521, w: 1.833, h: 1.833, fill: { color: C.green }, line: { type: 'none' }, rectRadius: 0.306, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 60, bold: true, color: C.white });
  text("Aliquet nibh praesent tristique magna sit amet. Risus venti la pretium quam vulputate dign tempor incididu.", 0.924, 3.117, 3.317, 0.904, { ...S.body, align: 'right', line: 1.5 });
  text("01. Strength", 1.766, 2.624, 2.475, 0.516, { ...S.h3, align: 'right', line: 1.5 });
  text("Aliquet nibh praesent tristique magna sit amet. Risus venti la pretium quam vulputate dign tempor incididu.", 0.924, 5.214, 3.317, 0.904, { ...S.body, align: 'right', line: 1.5 });
  text("04. Threats", 1.766, 4.721, 2.475, 0.516, { ...S.h3, align: 'right', line: 1.5 });
  text("Aliquet nibh praesent tristique magna sit amet. Risus venti la pretium quam vulputate dign tempor incididu.", 9.088, 3.117, 3.281, 0.903, { ...S.body, align: 'justify', line: 1.5 });
  text("02. Weakness", 9.084, 2.624, 2.475, 0.516, { ...S.h3, line: 1.5 });
  text("Aliquet nibh praesent tristique magna sit amet. Risus venti la pretium quam vulputate dign tempor incididu.", 9.129, 5.214, 3.281, 0.903, { ...S.body, align: 'justify', line: 1.5 });
  text("03. Opportunities", 9.126, 4.721, 2.475, 0.516, { ...S.h3, line: 1.5 });
}

// ---------------------------------------------------------------- slide 27
function slide27() {  // Professional Team
  box('roundRect', 0.667, 2.5, 2.857, 4.229, C.white, { shadow: CARD_SHADOW, rectRadius: 0.296 });
  box('roundRect', 3.714, 2.5, 2.857, 4.229, C.white, { shadow: CARD_SHADOW, rectRadius: 0.296 });
  box('roundRect', 6.762, 2.5, 2.857, 4.229, C.pink, { rectRadius: 0.296 });
  box('roundRect', 9.809, 2.5, 2.857, 4.229, C.white, { shadow: CARD_SHADOW, rectRadius: 0.296 });
  cloudTopLeft(true);
  grassBottomLeft(true);
  heading([["Professional ", 0], ["Team", 1]], 2.604, 0.949, 8.125, 0.774, { align: 'center' });
  text("Niki J. Alpha", 6.9, 5.339, 2.58, 0.438, { face: HEAD, size: 20, bold: true, color: C.white, align: 'center' });
  text("CEO and Founder", 7.091, 5.776, 2.199, 0.286, { ...S.captionW, align: 'center' });
  text("Karen L. Beta", 9.948, 5.339, 2.58, 0.438, { ...S.name, align: 'center' });
  text("General Manager", 10.139, 5.776, 2.199, 0.286, { ...S.caption, align: 'center' });
  text("Harmony Dune", 0.805, 5.339, 2.58, 0.438, { ...S.name, align: 'center' });
  text("Senior Toys Artist", 0.996, 5.776, 2.199, 0.286, { ...S.caption, align: 'center' });
  text("Sarah Albert", 3.853, 5.339, 2.58, 0.438, { ...S.name, align: 'center' });
  text("Senior Fashion Designer", 4.043, 5.776, 2.199, 0.286, { ...S.caption, align: 'center' });
}

// ---------------------------------------------------------------- slide 28
function slide28() {  // Style, joy, and creative toys we craft childhood magic in every design and play
  box('roundRect', 0.358, 0.5, 12.617, 6.499, C.white, { shadow: PANEL_SHADOW, rectRadius: 0.44 });
  text("Style, joy, and creative toys we craft childhood magic in every design and play", 6.464, 2.289, 5.646, 2.524, { face: HEAD, size: 36, bold: true, color: C.black });
  outline('quoteGlyph', 6.572, 1.579, 0.717, 0.569, C.pink);
  text("Quote By Your Name", 6.491, 5.24, 2.609, 0.37, { size: 16, bold: true, color: C.pink });
  text("Detail About You", 6.51, 5.618, 1.257, 0.269, { size: 10, color: C.gray, wrap: false });
  outline('blobOrganic', 1.152, 2.959, 3.728, 3.184, C.pink, { rotate: 52.4 });
  outline('greenSquiggle', 1.161, 1.808, 1.23, 0.22, C.green, { rotate: 319.56 });
  star(5.144, 5.09, 1, C.amber);
}

// ---------------------------------------------------------------- slide 29
function slide29() {  // Contact Information
  outline('greenSquiggle', 5.582, 1.259, 1.23, 0.22, C.green);
  star(3.938, 6.225, 1, C.pink);
  leafBottomRight(C.amber);
  squiggleTopRight();
  text("Ready to transform your space with Kifa’s Fashion and Toys Solutions? Don't hesitate, contact us now to start the journey toward creating your dream fashion and toys for kids!", 7.324, 3.292, 5.177, 0.908, { ...S.body, line: 1.5 });
  text("Address", 7.94, 4.46, 1.132, 0.37, S.contact);
  text(": 123 SF, Nusantara, Bali.", 9.216, 4.46, 3.115, 0.37, S.contact);
  box('ellipse', 7.401, 4.441, 0.416, 0.4, C.pink);
  text("Phone", 7.94, 5.048, 1.132, 0.37, S.contact);
  text(": (123) 123-4567", 9.216, 5.048, 3.115, 0.37, S.contact);
  box('ellipse', 7.401, 5.029, 0.416, 0.4, C.pink);
  text("Mail", 7.94, 5.636, 1.132, 0.37, S.contact);
  text(": info@yourmail.com", 9.216, 5.636, 3.115, 0.37, S.contact);
  box('ellipse', 7.401, 5.617, 0.416, 0.4, C.pink);
  outline('buildingGlyph', 7.52, 5.728, 0.167, 0.177, C.white);
  outline('mailGlyph', 7.593, 6.376, 0.189, 0.181, C.white);
  outline('phoneGlyph', 7.51, 5.161, 0.19, 0.136, C.white);
  outline('pinGlyph', 7.534, 4.541, 0.139, 0.201, C.white);
  heading([["Contact ", 0], ["Information", 1]], 7.317, 1.398, 4.186, 1.582, { size: 44 });
}

// ---------------------------------------------------------------- slide 30
function slide30() {  // Thanks
  outline('hillGhost', 0, 4.616, 6.783, 2.884, C.smoke, { transparency: 70 });
  grassBottomLeft();
  cloudTopLeft();
  heading([["Than", 0], ["ks", 1]], 0.479, 1.976, 5.562, 2.036, { size: 115 });
  text("And See You Next Time", 0.562, 3.854, 4.341, 0.572, { face: HEAD, size: 28, bold: true, wrap: false });
}

// ---------------------------------------------------------------- build
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'KIFA', width: 13.333, height: 7.5 });
  pptx.layout = 'KIFA';
  pptx.author = 'Kifa';
  pptx.title = 'Kifa - Kid Fashion and Toys Presentation';
  SLIDES.forEach(fn => { slide = pptx.addSlide(); fn(); });
  return pptx.writeFile({ fileName: path.join(__dirname, '05462bcb-f5c0-4c64-b9e5-d81df7feb490_grok_final.pptx') });
}

build().then(() => console.log('wrote 05462bcb-f5c0-4c64-b9e5-d81df7feb490_grok_final.pptx')).catch(err => { console.error(err); process.exit(1); });
