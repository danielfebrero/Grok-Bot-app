// Auto-generated recreation of the EMORA business-profile deck (100 slides, 13.333 x 7.5 in).
// Every slide is built from plain pptxgenjs calls; raster photos are replaced by
// flat placeholder rectangles (see `photo`).
const PptxGenJS = require('pptxgenjs');
const nodePath = require('path');

// ---------------------------------------------------------------- palette
const BLUE = '61CAF7';   // brand cyan
const DARK = '252525';   // near-black headline colour
const GREY = '808080';   // body copy
const WHITE = 'FFFFFF';
const SILVER = 'F2F2F2';
const HAIR = 'D9D9D9';   // hairline borders

// -------------------------------------------------------------- typefaces
const MEB = 'Montserrat ExtraBold';
const OS = 'Open Sans';
const OSSB = 'Open Sans SemiBold';
const MW = 'Merriweather';
const RB = 'Roboto';

// pptxgenjs mutates the options object it is handed, so hand it a fresh one each time.
const shadow = () => ({ type: 'outer', blur: 15, offset: 4, angle: 90, color: '000000', opacity: 0.2 });

// ---------------------------------------------------------- text presets
const OS12GreyJ = {fontFace:OS, fontSize:12, color:GREY, align:'justify', lineSpacingMultiple:1.5};
const MEB44DarkL = {fontFace:MEB, fontSize:44, color:DARK, align:'left', charSpacing:3};
const MEB16BlueL = {fontFace:MEB, fontSize:16, color:BLUE, align:'left'};
const OS12GreyC = {fontFace:OS, fontSize:12, color:GREY, align:'center', lineSpacingMultiple:1.5};
const MEB16BlueC = {fontFace:MEB, fontSize:16, color:BLUE, align:'center'};
const OS12WhiteJ = {fontFace:OS, fontSize:12, color:WHITE, align:'justify', lineSpacingMultiple:1.5};
const OS12GreyL = {fontFace:OS, fontSize:12, color:GREY, align:'left', lineSpacingMultiple:1.5};
const MEB12WhiteC = {fontFace:MEB, fontSize:12, color:WHITE, align:'center'};
const OSSB12GreyC = {fontFace:OSSB, fontSize:12, color:GREY, align:'center', lineSpacingMultiple:1.5};
const MEB24DarkC = {fontFace:MEB, fontSize:24, color:DARK, align:'center'};
const MEB44DarkC = {fontFace:MEB, fontSize:44, color:DARK, align:'center', charSpacing:3};
const MEB24WhiteC = {fontFace:MEB, fontSize:24, color:WHITE, align:'center'};
const MEB16WhiteC = {fontFace:MEB, fontSize:16, color:WHITE, align:'center'};
const OS12GreyR = {fontFace:OS, fontSize:12, color:GREY, align:'right', lineSpacingMultiple:1.5};
const MEB28WhiteC = {fontFace:MEB, fontSize:28, color:WHITE, align:'center'};
const OS12BlueJ = {fontFace:OS, fontSize:12, color:BLUE, align:'justify', lineSpacingMultiple:1.5};
const OS12WhiteC = {fontFace:OS, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5};
const MEB12BlueC = {fontFace:MEB, fontSize:12, color:BLUE, align:'center'};
const MEBdWhiteC = {fontFace:MEB, color:WHITE, align:'center'};
const MEB40BlueC = {fontFace:MEB, fontSize:40, color:BLUE, align:'center'};
const MEB36DarkL = {fontFace:MEB, fontSize:36, color:DARK, align:'left'};
const MEB32WhiteC = {fontFace:MEB, fontSize:32, color:WHITE, align:'center'};
const MEB44DarkR = {fontFace:MEB, fontSize:44, color:DARK, align:'right', charSpacing:3};
const MEB40WhiteC = {fontFace:MEB, fontSize:40, color:WHITE, align:'center'};
const OSSB12WhiteL = {fontFace:OSSB, fontSize:12, color:WHITE, align:'left', lineSpacingMultiple:1.5};
const MEB44WhiteC = {fontFace:MEB, fontSize:44, color:WHITE, align:'center', charSpacing:3};
const MEB16BlueR = {fontFace:MEB, fontSize:16, color:BLUE, align:'right'};
const MEB44WhiteL = {fontFace:MEB, fontSize:44, color:WHITE, align:'left', charSpacing:3};
const MEB12BlueL = {fontFace:MEB, fontSize:12, color:BLUE, align:'left'};
const MEB12GreyR = {fontFace:MEB, fontSize:12, color:GREY, align:'right'};
const MEBdWhiteL = {fontFace:MEB, color:WHITE, align:'left'};
const MEB96BlueC = {fontFace:MEB, fontSize:96, color:BLUE, align:'center', charSpacing:3};
const MEB500BlueL = {fontFace:MEB, fontSize:500, color:BLUE, align:'left', charSpacing:3};
const MEB44WhiteR = {fontFace:MEB, fontSize:44, color:WHITE, align:'right', charSpacing:3};
const MEB16GreyL = {fontFace:MEB, fontSize:16, color:GREY, align:'left'};
const MEB16GreyR = {fontFace:MEB, fontSize:16, color:GREY, align:'right'};
const MEB24BlueC = {fontFace:MEB, fontSize:24, color:BLUE, align:'center'};
const OS10GreyC = {fontFace:OS, fontSize:10, color:GREY, align:'center', lineSpacingMultiple:1.5};
const MEB10WhiteC = {fontFace:MEB, fontSize:10, color:WHITE, align:'center', lineSpacingMultiple:1.5};
const MEB400BlueR = {fontFace:MEB, fontSize:400, color:BLUE, align:'right', charSpacing:3};
const MEB400BlueC = {fontFace:MEB, fontSize:400, color:BLUE, align:'center', charSpacing:3};
const MWdWhiteC = {fontFace:MW, color:WHITE, align:'center', lineSpacingMultiple:1.5};
const RB12GreyJ = {fontFace:RB, fontSize:12, color:GREY, align:'justify', lineSpacingMultiple:1.5};

// --------------------------------------------- copy blocks reused verbatim
const C1 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium.';
const C2 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui.';
const C3 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium.';
const C4 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.';
const C5 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit.';
const C6 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';
const C7 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit';
const C8 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia deserunt mollitia animi.';
const C9 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const C10 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const C11 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam';
const C12 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam.';
const C13 =
  '“Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium.”';
const C14 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia deserunt mollitia animi, id est laborum et dolorum fuga. Et harum quidem rerum facilis est et expedita distinctio. Nam libero tempore, cum soluta nobis est.';
const C15 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae.';
const C16 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed .incididunt ut';
const C17 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum.';
const C18 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const C19 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae.';
const C20 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium.';
const C21 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const C22 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. ';
const C23 =
  'At vero eos et accusamus et iusto odio';
const C24 =
  '“Sed ut perspiciatis unde omnis iste natus error sit voluptatem.';
const C25 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.';
const C26 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit.';
const C27 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem.';
const C28 =
  'Sed ut perspiciatis unde omnis iste natus error sit.';
const C29 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus.';
const C30 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium placeat facere possimus, omnis voluptas assumenda.';
const C31 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum';
const C32 =
  'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas. molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia deserunt mollitia animi, id est laborum et dolorum fuga. Et harum quidem rerum facilis est et expedita distinctio. Nam libero tempore, cum soluta nobis est.';
const C33 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab.';
const C34 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut';
const C35 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore Veritatis.';
const C36 =
  'Sed ut perspiciatis unde omnis iste natus error';
const C37 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi.';
const C38 =
  'PLACEHOLDER';
const C39 =
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt.';

// ------------------------------------------------------------- geometries
// Outlines traced from the original vector artwork; coordinates are
// per-mille of the shape's bounding box and are scaled by `path()`.
const g1 = [[0,0,1000,0,1000,1000,0,1000]];
const g2 = [[0,0,1000,0,1000,1000,0,1000]];
const g3 = [[1000,500,990,601,961,695,854,854,695,961,601,990,500,1000,399,990,305,961,146,854,39,695,10,601,0,500,10,399,39,305,146,146,305,39,399,10,500,0,601,10,695,39,854,146,961,305,990,399,1000,500]];
const g4 = [[1000,500,990,601,961,695,854,854,695,961,601,990,500,1000,399,990,305,961,146,854,39,695,10,601,0,500,10,399,39,305,146,146,305,39,399,10,500,0,601,10,695,39,854,146,961,305,990,399,1000,500]];
const g5 = [[1000,500,990,601,961,695,854,854,695,961,601,990,500,1000,399,990,305,961,146,854,39,695,10,601,0,500,10,399,39,305,146,146,305,39,399,10,500,0,601,10,695,39,854,146,961,305,990,399,1000,500]];
const g6 = [[1000,500,990,601,961,695,854,854,695,961,601,990,500,1000,399,990,305,961,146,854,39,695,10,601,0,500,10,399,39,305,146,146,305,39,399,10,500,0,601,10,695,39,854,146,961,305,990,399,1000,500]];
const g7 = [[500,0,601,10,695,39,780,85,854,146,915,220,961,305,990,399,1000,500,990,601,961,695,915,780,854,854,780,915,695,961,601,990,500,1000,399,990,305,961,220,915,146,854,85,780,39,695,10,601,0,500,10,399,39,305,85,220,146,146,220,85,305,39,399,10,500,0]];
const g8 = [[1000,417,355,27,0,583,916,583,907,840,183,914,0,667,16,873,996,834,1000,500,83,500,160,260,1000,417],[341,167,659,167,341,167]];
const g9 = [[969,362,745,79,408,10,95,248,8,571,255,921,592,990,941,693,969,362],[899,582,423,884,84,479,293,175,577,115,843,320,899,582]];
const g10 = [[1000,500,854,854,500,1000,146,854,0,500,146,146,500,0,596,9,500,200,332,251,206,440,288,712,500,800,668,749,955,294,1000,500]];
const g11 = [[919,193,502,0,52,215,401,656,750,576,562,906,167,639,229,917,718,959,917,489,988,896,919,193],[883,406,480,588,84,334,520,93,883,406]];
const g12 = [[791,0,36,92,92,964,292,1000,500,667,708,1000,964,908,996,166,791,0],[916,792,791,917,441,589,120,880,208,83,880,120,917,792]];
const g13 = [[500,0,146,146,0,500,146,853,500,1000,854,854,1000,500,854,147,500,0],[500,750,323,677,250,500,323,323,500,250,750,500,500,750]];
const g14 = [[167,0,833,0,867,2,898,7,927,15,951,26,972,39,987,54,997,71,1000,89,1000,911,997,929,987,946,972,961,951,974,927,985,898,993,867,998,833,1000,167,1000,133,998,102,993,73,985,49,974,28,961,13,946,3,929,0,911,0,89,3,71,13,54,28,39,49,26,73,15,102,7,133,2,167,0]];
const g15 = [[708,0,726,1,743,8,772,38,799,86,868,240,956,369,979,418,996,482,1000,549,992,614,983,646,954,705,934,730,911,749,868,766,779,786,740,810,722,831,706,856,667,938,652,962,634,981,613,995,591,1000,562,999,494,991,434,979,376,957,339,936,291,893,264,856,245,814,227,746,218,674,207,482,197,429,186,395,160,352,140,333,118,320,93,313,47,317,27,312,18,306,5,285,1,271,4,232,14,208,30,188,49,174,70,165,126,157,253,162,311,152,368,131,536,47,592,22,649,5,708,0]];
const g16 = [[772,192,1000,192,1000,1000,772,1000],[257,192,485,192,485,1000,257,1000],[515,0,743,0,743,808,515,808],[0,0,228,0,228,808,0,808]];
const g17 = [[745,0,786,3,826,15,885,50,938,105,958,137,975,175,996,266,999,379,982,503,970,551,957,582,942,605,924,620,884,631,862,629,817,615,728,569,679,551,625,547,571,562,521,599,493,634,478,663,465,703,452,815,444,855,431,891,397,947,356,983,311,999,258,995,220,980,185,954,163,930,144,898,130,860,119,796,111,659,103,593,85,535,39,438,31,310,1,217,2,172,18,121,44,90,74,74,119,63,165,63,210,71,256,85,393,137,439,149,492,153,532,135,607,61,645,32,684,13,745,0]];
const g18 = [[433,0,540,2,611,9,681,19,747,34,786,49,808,62,830,84,855,122,876,189,882,230,890,396,896,437,914,502,982,649,995,695,1000,739,993,804,978,846,967,865,939,902,921,918,882,945,860,956,803,971,769,975,558,966,493,970,301,999,250,999,199,992,154,975,135,962,119,946,98,905,71,824,53,782,14,750,4,734,0,719,2,695,18,644,47,587,104,499,120,461,125,442,127,419,122,375,109,332,75,245,64,200,61,155,65,133,72,112,83,92,117,60,138,48,185,30,284,11,433,0]];
const g19 = [[537,224,1000,224,1000,1000,537,1000],[0,0,463,0,463,776,0,776]];
const g20 = [[1000,202,931,91,839,142,796,0,194,1,161,142,34,113,45,361,236,587,281,531,474,684,336,1000,664,1000,526,684,719,531,764,587,1000,202],[232,443,60,302,72,142,232,443],[911,342,759,453,910,138,960,225,911,342]];
const g21 = [[228,631,148,679,87,344,4,300,123,233,234,339,182,105,294,98,360,0,515,130,609,70,769,120,823,46,907,154,865,359,768,306,783,513,675,574,661,735,619,676,655,823,541,537,434,741,393,556,195,439,285,555,228,631],[452,720,461,766,452,720],[794,605,819,684,794,605],[835,726,830,774,835,726],[699,819,778,778,746,895,699,819],[984,221,953,268,984,221],[962,269,955,368,878,387,962,269],[590,779,667,917,590,779],[826,832,782,916,826,832],[680,921,740,949,680,921],[960,878,958,958,878,860,960,878]];
const g22 = [[1000,414,854,121,500,0,146,121,0,414,500,1000,1000,414],[500,683,271,604,176,414,271,224,500,146,730,224,825,415,730,605,500,683]];
const g23 = [[854,146,688,37,500,0,312,37,146,146,37,312,0,500,37,688,146,854,312,963,500,1000,688,963,854,854,963,688,1000,499,854,146]];
const g24 = [[439,735,334,766,292,792,275,827,202,835,146,880,88,865,91,785,0,582,28,576,3,527,4,434,15,452,58,390,98,382,185,324,220,214,246,240,241,196,265,205,262,171,309,125,340,116,402,149,395,122,418,65,478,40,487,17,540,42,596,40,565,140,694,235,727,148,732,51,759,1,785,121,806,120,826,150,849,293,904,344,917,414,944,433,992,547,1000,648,975,747,914,844,883,961,833,973,802,1000,768,969,737,995,679,968,646,885,621,880,623,835,593,861,617,769,567,848,521,761,439,735]];
const g25 = [[998,89,962,148,967,206,988,220,944,273,959,273,922,447,937,532,859,700,901,875,892,944,823,787,787,802,763,774,727,775,692,802,683,853,644,825,631,846,578,838,567,894,547,888,523,923,529,997,507,997,432,850,411,845,380,880,318,758,221,766,140,697,102,694,74,623,39,597,0,339,54,34,79,54,77,86,94,16,386,68,526,70,540,51,550,79,633,96,595,149,631,155,664,107,656,130,680,146,731,131,785,324,832,277,841,223,877,210,887,135,938,116,947,4,998,89]];
const g26 = [[583,0,1000,371,906,451,719,611,625,691,578,731,529,771,478,810,426,850,371,888,314,926,254,963,190,1000,0,713,583,0]];
const g27 = [[1000,601,834,732,771,782,706,833,639,887,570,943,500,1000,427,876,357,752,290,628,225,503,164,378,106,252,51,126,0,0,1000,601]];
const g28 = [[1000,268,74,1000,0,683,72,627,144,571,216,516,289,463,362,410,435,358,509,308,583,259,827,101,875,68,923,35,971,0,1000,268]];
const g29 = [[935,316,945,315,945,261,923,263,870,236,776,264,493,193,492,0,281,0,276,442,0,414,13,458,123,576,150,667,166,689,211,716,256,723,274,711,288,657,304,641,351,632,391,646,432,683,556,866,587,959,657,992,725,999,737,981,713,907,711,817,766,799,796,735,809,731,841,755,868,749,893,706,905,617,924,617,936,647,981,631,1000,512,994,482,952,416,935,316]];
const g30 = [[961,878,997,855,996,844,853,705,426,335,524,53,88,0,83,12,91,37,84,54,14,120,2,136,1,152,6,163,48,211,48,221,28,251,24,288,40,325,94,388,103,468,128,501,133,538,166,612,228,681,235,705,237,752,243,761,274,774,459,823,473,844,549,881,585,910,602,936,617,976,637,989,679,997,856,996,913,990,935,982,936,971,920,936,928,910,961,878]];
const g31 = [[247,0,252,97,238,132,210,145,126,137,103,155,95,171,86,236,91,320,145,412,148,445,66,541,58,569,75,663,67,680,44,692,0,701,32,730,170,806,357,888,532,954,597,975,670,987,918,1000,1000,64,247,0]];
const g32 = [[674,1000,672,962,673,940,677,918,685,898,696,881,713,869,736,864,760,866,814,874,827,874,838,873,845,866,849,856,851,843,852,827,846,764,851,729,893,544,1000,99,802,80,563,54,325,26,129,0,0,391,674,1000]];
const g33 = [[899,788,902,732,942,721,972,675,989,722,939,852,894,849,884,816,899,788],[967,374,999,368,978,427,825,581,850,712,758,778,771,842,731,864,725,906,656,976,546,1000,522,990,519,950,433,763,429,721,456,663,435,601,446,596,382,528,401,518,397,467,343,464,310,435,147,464,107,445,64,414,36,365,11,359,32,345,0,328,21,284,7,240,45,172,101,129,125,72,173,24,225,36,275,13,398,0,417,4,405,57,531,109,573,70,672,99,751,98,757,140,730,114,730,127,834,323,899,382,967,374]];
const g34 = [[1000,0,0,1000,1000,1000]];
const g35 = [[730,662,950,596,1000,427,765,25,649,194,615,92,566,176,497,73,469,231,367,90,408,308,280,390,230,84,129,12,0,317,57,393,199,349,186,544,70,599,221,700,178,993,315,858,475,998,593,950,664,676,730,662]];
const g36 = [[990,881,812,861,672,904,742,1000,994,951,990,881],[692,889,729,774,790,742,816,801,864,793,840,755,926,687,778,603,749,458,889,288,788,144,899,180,960,143,738,54,758,26,651,42,669,0,473,84,494,117,297,313,283,390,326,436,385,393,449,515,530,390,508,290,601,190,666,214,584,283,586,355,741,385,629,399,639,459,567,471,550,542,397,551,379,446,347,459,356,561,95,697,160,812,25,819,0,933,23,979,131,985,254,823,365,808,482,958,524,906,432,774,550,865,614,991,604,891,692,889]];
const g37 = [[8,282,20,402,167,343,173,485,361,462,345,544,421,573,397,674,211,678,200,717,304,748,176,839,469,868,111,1000,925,916,870,859,998,750,826,730,503,364,377,332,519,159,291,150,363,3,128,67,28,174,8,282]];
const g38 = [[318,997,780,875,879,778,853,356,1000,147,924,30,441,62,565,161,546,252,212,295,69,220,71,501,279,676,23,861,19,995,318,997]];
const g39 = [[829,52,559,211,433,238,224,174,1,0,1,71,45,246,137,371,648,607,775,898,939,999,1000,914,920,575,976,53,941,16,829,52]];
const g40 = [[15,391,299,523,369,612,359,683,103,873,171,996,364,969,762,761,922,608,999,434,980,316,847,172,627,72,262,1,24,74,15,391]];
const g41 = [[157,980,316,878,765,843,943,694,995,533,993,359,951,195,858,19,780,15,598,250,473,356,340,400,20,418,15,933,69,996,157,980]];
const g42 = [[246,88,164,488,96,654,9,768,14,832,172,926,427,987,733,1000,971,965,1000,768,948,397,883,228,785,95,649,15,471,7,246,88]];
const g43 = [[935,98,907,16,740,13,475,156,287,181,1,373,194,650,389,748,425,881,666,911,838,994,946,885,933,708,998,452,909,253,935,98]];
const g44 = [[66,112,6,183,6,240,176,372,213,554,321,695,293,807,165,980,334,998,535,947,942,598,906,364,1000,189,933,98,577,0,66,112]];
const g45 = [[556,38,566,8,640,1,670,131,726,160,744,241,794,253,838,221,874,278,888,261,994,305,991,527,940,535,908,636,722,768,762,868,830,851,823,900,767,899,672,1000,638,979,642,904,582,876,647,787,575,788,539,764,537,724,486,733,416,844,420,898,403,872,360,896,334,879,376,735,433,742,369,534,293,492,154,566,132,534,37,520,0,451,46,375,45,305,124,210,112,66,154,40,230,44,376,119,394,93,409,125,453,113,468,138,492,40,556,38],[667,772,695,788,667,772],[739,851,747,868,739,851]];
const g46 = [[913,332,922,426,1000,499,933,327,896,0,822,113,799,233,703,290,677,382,640,330,563,405,579,315,532,303,454,400,455,482,421,429,394,560,412,405,369,501,342,446,256,456,260,401,237,449,194,439,194,377,214,418,230,395,235,305,207,298,146,424,67,479,37,590,62,731,0,791,44,959,119,799,156,794,230,879,261,819,328,827,388,878,401,952,471,1000,510,971,551,992,569,941,668,957,743,902,738,832,781,780,818,840,900,812,907,950,927,748,880,651,844,671,834,541,901,426,889,304,913,332]];
const g47 = [[924,4,800,79,615,87,484,134,401,197,327,322,140,452,63,527,6,638,1,728,32,889,113,995,136,983,170,907,188,735,261,597,355,518,410,443,578,382,654,281,878,210,929,178,991,77,999,25,964,0,924,4]];
const g48 = [[327,80,213,1,114,75,58,303,1,361,81,572,333,723,666,752,486,937,525,979,740,994,991,845,925,594,506,478,415,163,327,80]];
const g49 = [[708,7,561,7,360,103,243,215,36,483,0,614,35,879,77,984,117,1000,176,982,257,920,977,240,1000,173,973,111,881,54,708,7]];
const g50 = [[809,47,575,2,464,21,381,132,374,270,170,351,96,516,0,614,386,981,659,963,819,765,954,484,998,317,982,173,889,73,809,47]];
const g51 = [[308,55,18,168,1,256,18,536,88,655,386,917,535,993,671,977,963,632,998,532,981,406,874,218,688,52,572,8,445,4,308,55]];
const g52 = [[833,217,803,70,756,5,646,30,433,149,16,697,0,756,24,807,440,995,540,989,675,910,859,751,988,543,997,432,950,322,833,217]];
const g53 = [[221,99,170,81,70,1,17,18,2,66,35,131,55,216,164,327,414,503,715,869,867,999,854,884,864,855,986,842,1000,828,993,815,920,750,779,657,708,577,688,521,739,469,732,452,554,355,385,160,342,129,221,99]];
const g54 = [[316,167,60,444,0,584,25,630,512,903,802,997,907,991,967,933,984,862,906,718,1000,470,913,213,695,26,565,0,435,42,316,167]];
const g55 = [[55,500,262,848,465,987,569,999,668,959,755,859,960,524,997,361,885,230,755,172,436,55,163,0,67,29,11,113,4,266,55,500]];
const g56 = [[76,129,4,212,18,340,97,493,223,654,537,921,687,990,807,991,898,902,967,737,999,531,979,320,891,139,722,24,455,8,76,129]];
const g57 = [[108,359,23,531,0,682,31,809,105,906,214,971,498,986,655,928,893,784,994,660,999,597,930,450,567,15,480,2,378,45,108,359]];
const g58 = [[59,208,6,371,6,508,119,712,312,834,792,997,911,957,994,722,987,539,915,364,796,210,646,90,481,16,317,3,171,63,59,208]];
const g59 = [[677,17,520,128,518,205,589,312,543,374,109,533,3,791,33,848,690,998,854,929,874,779,802,661,975,286,996,111,906,12,677,17]];
const g60 = [[777,2,728,79,622,111,8,125,50,236,277,457,133,772,324,930,594,993,653,692,872,629,973,530,875,273,999,114,942,53,777,2]];
const g61 = [[500,0,601,10,695,39,780,85,854,146,915,220,961,305,990,399,1000,500,990,601,961,695,915,780,854,854,780,915,695,961,601,990,500,1000,399,990,305,961,220,915,146,854,85,780,39,695,10,601,0,500,10,399,39,305,85,220,146,146,220,85,305,39,399,10,500,0]];
const g62 = [[13,724,65,604,129,492,204,387,291,291,386,205,491,129,604,65,773,1,821,2,868,14,933,54,965,91,994,156,1000,205,1000,795,984,875,965,910,910,965,836,996,795,1000,156,994,71,950,24,891,2,822,13,724]];
const g63 = [[724,986,604,934,492,870,387,795,291,709,205,613,129,508,65,395,1,226,2,178,14,131,54,67,91,35,156,6,205,0,795,0,875,16,910,35,965,90,996,163,1000,205,994,845,950,930,891,976,798,1000,724,986]];
const g64 = [[986,276,934,396,870,508,795,613,709,709,613,795,508,870,395,934,226,999,178,998,131,986,67,946,35,910,6,844,0,793,0,205,16,125,35,91,90,35,163,4,205,0,845,6,930,50,976,109,1000,203,986,276]];
const g65 = [[276,13,396,65,508,129,613,204,709,291,795,386,870,491,934,604,999,773,998,821,986,868,946,933,910,965,844,994,793,1000,205,1000,125,984,91,965,35,910,4,836,0,795,6,156,50,71,109,24,179,2,276,13]];
const g66 = [[1000,1000,0,1000,106,0,894,0]];
const g67 = [[1000,1000,0,1000,139,0,861,0]];
const g68 = [[1000,1000,0,1000,200,0,800,0]];
const g69 = [[1000,1000,0,1000,500,0]];
const g70 = [[0,0,790,0,832,3,872,12,907,25,939,44,964,66,984,91,996,119,1000,149,1000,1000,210,1000,168,997,128,988,93,975,61,956,36,934,16,909,4,881,0,851]];
const g71 = [[500,1000,176,783,99,720,44,651,11,576,0,500,11,424,44,349,99,280,176,217,500,0,824,217,901,280,956,349,989,424,1000,500,989,576,956,651,901,720,824,783,500,1000]];
const g72 = [[500,1000,176,783,99,720,44,651,11,576,0,500,11,424,44,349,99,280,176,217,500,0,824,217,901,280,956,349,989,424,1000,500,989,576,956,651,901,720,824,783,500,1000]];
const g73 = [[0,500,217,176,280,99,349,44,424,11,500,0,576,11,651,44,720,99,783,176,1000,500,783,824,720,901,651,956,576,989,500,1000,424,989,349,956,280,901,217,824,0,500]];
const g74 = [[0,500,217,176,280,99,349,44,424,11,500,0,576,11,651,44,720,99,783,176,1000,500,783,824,720,901,651,956,576,989,500,1000,424,989,349,956,280,901,217,824,0,500]];
const g75 = [[750,0,250,0,0,500,250,1000,750,1000,1000,500]];
const g76 = [[750,0,250,0,0,500,250,1000,750,1000,1000,500]];
const g77 = [[1000,332,961,500,671,509,659,999,237,997,232,508,0,493,3,337,229,331,272,110,471,18,993,3,996,165,698,185,669,324,1000,332]];
const g78 = [[300,782,158,709,106,610,195,603,47,444,33,358,124,382,61,48,491,311,587,36,974,21,709,817,497,966,259,997,0,891,300,782]];
const g79 = [[416,652,336,869,213,993,323,270,425,221,514,254,465,552,620,620,790,464,778,202,550,100,262,167,150,325,206,507,164,581,60,523,4,411,94,153,276,40,501,0,735,31,911,134,1000,348,897,604,645,723,416,652]];
const g80 = [[324,138,433,37,535,6,637,0,821,38,898,84,955,143,989,217,1000,306,991,992,715,1000,685,988,673,390,628,301,571,264,522,254,474,257,409,287,360,336,329,402,317,991,7,991,3,20,307,11,322,33,324,138]];
const g81 = [[996,501,996,967,988,988,977,993,892,999,102,1000,36,997,12,989,0,968,1,23,23,6,78,0,912,0,986,12,1000,37,996,501]];
const g82 = [[498,1000,304,961,146,855,0,502,39,306,147,146,308,38,505,0,698,41,856,150,962,311,1000,506,959,700,852,857,694,962,498,1000]];


// ---------------------------------------------------------------- helpers

// Flat stand-in for a photograph / device mock-up in the original deck.
const TONES = { mid: '7F7F7F', light: 'C3C3C3', device: 'DCDCDC', mapblue: '9BDDFA' };
function photo(s, x, y, w, h, tone) {
  s.addShape('rect', { x, y, w, h, fill: { color: TONES[tone || 'mid'] } });
}

// The little cyan "+" / "x" sparkle that decorates almost every slide.
function cross(s, x, y, rot) {
  const r = rot === undefined ? 45 : rot, d = 0.25, line = { color: BLUE, width: 3 };
  s.addShape('line', { x, y: y + d / 2, w: d, h: 0, line, rotate: r });
  s.addShape('line', { x: x + d / 2, y, w: 0, h: d, line, rotate: r });
}

// Small cyan outline ring, the deck's other recurring accent.
function ring(s, x, y) {
  s.addShape('ellipse', { x, y, w: 0.255, h: 0.255, fill: { type: 'none' }, line: { color: BLUE, width: 3 } });
}

// Draws a traced outline: `pts` is a list of closed sub-paths in per-mille units.
function path(s, opts) {
  const pts = [];
  opts.points.forEach(sub => {
    for (let i = 0; i < sub.length; i += 2) {
      pts.push({ x: (sub[i] / 1000) * opts.w, y: (sub[i + 1] / 1000) * opts.h, moveTo: i === 0 });
    }
    pts.push({ close: true });
  });
  const o = Object.assign({}, opts, { points: pts });
  s.addShape('custGeom', o);
}

// Text block: `style` carries font/size/colour/alignment, `box` the geometry.
function txt(s, text, style, box, extra) {
  s.addText(text, Object.assign({ isTextBox: true, valign: 'top', wrap: true }, style, box, extra || {}));
}

// The single native chart in the deck (slide 93).
function areaChart(s, x, y, w, h) {
  const labels = ['1/1/2002', '1/2/2002', '1/3/2002', '1/4/2002', '1/5/2002'];
  s.addChart('area', [
    { name: 'Series 1', labels, values: [32, 32, 28, 12, 15] },
    { name: 'Series 2', labels, values: [12, 12, 12, 21, 28] },
  ], {
    x, y, w, h, showLegend: false, chartColors: ['4472C4', 'ED7D31'],
    catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 12, catAxisLabelColor: '595959',
    valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 12, valAxisLabelColor: '595959',
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catAxisLineColor: 'D9D9D9',
  });
}

// --------------------------------------------------------------- slides
function slide1(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:0, w:13.33, h:7.5, fill:{color:DARK, transparency:15}});
  txt(s, 'EMORA', MEB96BlueC, {x:2.8, y:2.89, w:7.73, h:1.72});
  cross(s, .94, .79);
  cross(s, 11.11, 6, 0);
  cross(s, 1.78, 3.96, 0);
  cross(s, 10.11, 2.44);
  ring(s, 11.68, .89);
  ring(s, 3.23, 6.25);
}

function slide2(s) {
  s.addShape('rect', {x:9.94, y:2.69, w:3.39, h:4.81, fill:{color:BLUE}});
  txt(s, 'COMPANY\nPROFILE', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C21, OS12GreyJ, {x:1.1, y:2.69, w:4.96, h:1.28});
  txt(s, C1, OS12GreyJ, {x:1.52, y:4.94, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.27, y:5.08, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.52, y:5.73, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.27, y:5.88, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.52, y:4.14, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.27, y:4.29, w:.15, h:.15, fill:{color:BLUE}});
  ring(s, 8.72, 2.04);
  cross(s, 7.45, 3.62);
  photo(s, 8.85, 1.1, 3.39, 5.31, 'light');
  photo(s, 7.15, 3.75, 3.39, 2.65);
}

function slide3(s) {
  s.addShape('rect', {x:0, y:0, w:1.1, h:7.5, fill:{color:BLUE}});
  txt(s, 'WHO\nWE ARE', MEB44DarkR, {x:7.55, y:1.1, w:4.68, h:1.58});
  txt(s, C14, OS12GreyJ, {x:7.28, y:2.95, w:4.96, h:2.19});
  txt(s, C30, OS12BlueJ, {x:7.28, y:5.43, w:4.96, h:.98});
  cross(s, 3.03, 6.28, 0);
  ring(s, 5.93, 1.96);
  photo(s, 0, 1, 6.06, 5.4);
}

function slide4(s) {
  txt(s, 'WHAT WE DO', MEB44DarkC, {x:3.37, y:1.1, w:6.6, h:.84});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  txt(s, C1, OS12GreyJ, {x:1.35, y:2.35, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.1, y:2.5, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.35, y:3.23, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.1, y:3.38, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.35, y:4.11, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.1, y:4.26, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.35, y:4.99, w:4.53, h:.67});
  s.addShape('ellipse', {x:1.1, y:5.14, w:.15, h:.15, fill:{color:BLUE}});
  ring(s, 6.72, 5.71);
  cross(s, 11.54, 2.23);
  photo(s, 6.85, 2.35, 5.39, 4.05);
}

function slide5(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:1.1, y:1.1, w:11.14, h:5.31, fill:{color:WHITE, transparency:15}});
  txt(s, 'ABOUT STORY', MEB44DarkC, {x:3.37, y:2.19, w:6.6, h:.84});
  txt(s, C14, OS12GreyC, {x:2.19, y:3.34, w:8.94, h:1.28});
  s.addShape('rect', {x:2.19, y:5.21, w:8.94, h:.1, fill:{color:BLUE}});
  ring(s, 2.19, .97);
  cross(s, 10.84, 6.28);
}

function slide6(s) {
  txt(s, 'ABOUT\nCOMPANY', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia deserunt mollitia animi, id est laborum et dolorum fuga. Et harum quidem rerum facilis est et expedita distinctio. ', OS12GreyJ, {x:1.1, y:2.81, w:6.58, h:1.58});
  txt(s, C25, OS12GreyJ, {x:1.1, y:5.12, w:2.74, h:1.28});
  txt(s, C25, OS12GreyJ, {x:4.94, y:5.12, w:2.74, h:1.28});
  s.addShape('line', {x:1.17, y:4.76, w:6.42, h:0, line:{color:HAIR, width:1.5}});
  s.addShape('line', {x:4.39, y:5.16, w:0, h:1.2, line:{color:HAIR, width:1.5}, flipV:true});
  ring(s, 8.65, 1.52);
  cross(s, 12.34, 6.28);
  photo(s, 8.78, 1.1, 4.56, 5.31);
}

function slide7(s) {
  s.addShape('rect', {x:0, y:1.5, w:3.14, h:6, fill:{color:BLUE}});
  txt(s, C21, OS12GreyJ, {x:7.22, y:2.11, w:5.02, h:1.28});
  txt(s, 'ABOUT US', MEB44DarkL, {x:7.22, y:1.1, w:4.68, h:.84});
  txt(s, C14, OS12GreyJ, {x:7.22, y:4.21, w:5.02, h:2.19});
  s.addShape('line', {x:7.33, y:3.87, w:4.9, h:0, line:{color:HAIR, width:1.5}});
  cross(s, 5.99, .79);
  ring(s, 5.23, 6.28);
  photo(s, 0, 4.15, 6.12, 2.25);
  photo(s, 0, 0, 6.12, 3.06);
}

function slide8(s) {
  txt(s, 'COMPANY\nPROFILE', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('rect', {x:6.12, y:1.1, w:6.12, h:5.31, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:3.74, y:3.48, w:5.31, h:.55, fill:{color:BLUE}, rotate:90});
  txt(s, C19, OS12GreyJ, {x:7.22, y:1.96, w:4.47, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:1.65, w:1.8, h:.37});
  txt(s, C19, OS12GreyJ, {x:7.22, y:3.41, w:4.47, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:3.11, w:1.8, h:.37});
  txt(s, C19, OS12GreyJ, {x:7.22, y:4.87, w:4.47, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:4.57, w:1.8, h:.37});
  txt(s, C14, OS12GreyJ, {x:1.1, y:2.87, w:4.47, h:2.49});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  ring(s, 11.43, 6.28);
  cross(s, 8.44, .97, 0);
}

function slide9(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:2.3, w:13.33, h:3.75, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:5.52, y:-5.52, w:2.3, h:13.33, fill:{color:BLUE, transparency:15}, rotate:90});
  txt(s, 'ABOUT COMPANY', MEB44WhiteC, {x:3.22, y:1.1, w:6.89, h:.84});
  txt(s, C31, OS12GreyC, {x:1.8, y:3.53, w:9.74, h:1.28});
  ring(s, 1.22, 5.2);
  cross(s, 11.84, 2.69, 0);
}

function slide10(s) {
  photo(s, 7.15, 1.1, 2.69, 6.4);
  photo(s, 10.39, 0, 2.94, 6.4);
  txt(s, 'COMPANY\nVISSION', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C31, OS12GreyJ, {x:1.1, y:2.93, w:4.96, h:2.49});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  ring(s, 10.26, 5.86);
  cross(s, 7.43, .97);
}

function slide11(s) {
  s.addShape('rect', {x:-.01, y:0, w:3.73, h:3.75, fill:{color:BLUE}});
  photo(s, 1.1, 1.1, 5.25, 5.31);
  txt(s, 'COMPANY\nMISSION', MEB44DarkR, {x:7.55, y:1.1, w:4.68, h:1.58});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat.', OS12GreyJ, {x:7.45, y:2.83, w:4.79, h:1.89});
  txt(s, C1, OS12GreyJ, {x:7.8, y:4.88, w:4.44, h:.67});
  s.addShape('ellipse', {x:7.55, y:5.02, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:7.8, y:5.73, w:4.44, h:.67});
  s.addShape('ellipse', {x:7.55, y:5.88, w:.15, h:.15, fill:{color:BLUE}});
  ring(s, 6.22, 1.76);
  cross(s, .97, 5.73, 0);
}

function slide12(s) {
  path(s, {x:3.9, y:2.69, w:.05, h:3.19, points:g1, fill:{color:DARK}, rotate:303.5});
  path(s, {x:9.21, y:2.69, w:.05, h:3.19, points:g1, fill:{color:DARK}, rotate:303.5});
  path(s, {x:4.98, y:4.25, w:3.19, h:.05, points:g2, fill:{color:DARK}, rotate:326.5});
  txt(s, C9, OS12GreyC, {x:9.21, y:2.89, w:2.69, h:.98});
  txt(s, '2021', MEB16BlueC, {x:9.66, y:2.62, w:1.8, h:.37});
  txt(s, C9, OS12GreyC, {x:6.56, y:4.67, w:2.69, h:.98});
  txt(s, '2020', MEB16BlueC, {x:7.01, y:4.4, w:1.8, h:.37});
  txt(s, C9, OS12GreyC, {x:3.91, y:2.89, w:2.69, h:.98});
  txt(s, '2019', MEB16BlueC, {x:4.36, y:2.62, w:1.8, h:.37});
  path(s, {x:4.68, y:4.59, w:1.14, h:1.14, points:g3, fill:{color:BLUE}});
  path(s, {x:9.99, y:4.59, w:1.14, h:1.14, points:g4, fill:{color:BLUE}});
  path(s, {x:2.21, y:3, w:.79, h:.79, points:g5, fill:{color:BLUE}});
  path(s, {x:7.51, y:3, w:.79, h:.79, points:g6, fill:{color:BLUE}});
  txt(s, C9, OS12GreyC, {x:1.26, y:4.67, w:2.69, h:.98});
  txt(s, '2018', MEB16BlueC, {x:1.7, y:4.4, w:1.8, h:.37});
  path(s, {x:2.42, y:3.21, w:.37, h:.37, points:g5, fill:{color:WHITE}});
  path(s, {x:7.72, y:3.21, w:.37, h:.37, points:g5, fill:{color:WHITE}});
  path(s, {x:4.9, y:4.82, w:.69, h:.69, points:g3, fill:{color:WHITE}});
  path(s, {x:10.23, y:4.82, w:.69, h:.69, points:g3, fill:{color:WHITE}});
  txt(s, 'COMPANY JOURNEY', MEB44DarkL, {x:2.9, y:1.1, w:7.54, h:.84});
  s.addShape('rect', {x:0, y:0, w:1.1, h:7.5, fill:{color:BLUE}});
  ring(s, 11.99, 6.14);
  cross(s, 12.11, 1.09, 0);
}

function slide13(s) {
  s.addShape('rect', {x:6.8, y:0, w:6.54, h:7.5, fill:{color:BLUE}});
  txt(s, 'COMPANY\nJOURNEY', MEB44DarkL, {x:1.1, y:1.1, w:11.03, h:1.58});
  s.addShape('ellipse', {x:7.37, y:1.1, w:1.84, h:1.84, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:10.4, y:2.83, w:1.84, h:1.84, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:7.37, y:4.56, w:1.84, h:1.84, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C3, OS12WhiteJ, {x:9.5, y:1.52, w:2.74, h:.98});
  txt(s, C15, OS12GreyJ, {x:1.09, y:2.78, w:5.19, h:1.28});
  s.addShape('rect', {x:1.09, y:4.38, w:5.19, h:2.05, fill:{color:BLUE}});
  txt(s, 'PLACEHOLDER', OS12WhiteJ, {x:1.61, y:4.76, w:4.15, h:1.28});
  cross(s, 6.01, 1.15);
  ring(s, 5.33, 2.25);
  txt(s, '2018', MEB16BlueC, {x:7.82, y:1.83, w:.95, h:.37});
  txt(s, '2019', MEB16BlueC, {x:10.84, y:3.56, w:.95, h:.37});
  txt(s, C3, OS12WhiteJ, {x:7.37, y:3.26, w:2.74, h:.98});
  txt(s, '2020', MEB16BlueC, {x:7.82, y:5.29, w:.95, h:.37});
  txt(s, C3, OS12WhiteJ, {x:9.5, y:4.99, w:2.74, h:.98});
}

function slide14(s) {
  txt(s, 'COMPANY JOURNEY', MEB44DarkL, {x:2.9, y:1.1, w:7.54, h:.84});
  txt(s, C9, OS12GreyC, {x:1.1, y:5.42, w:2.69, h:.98});
  txt(s, '2018', MEB16BlueC, {x:1.54, y:5.15, w:1.8, h:.37});
  txt(s, C9, OS12GreyC, {x:9.55, y:5.42, w:2.69, h:.98});
  txt(s, '2021', MEB16BlueC, {x:10, y:5.15, w:1.8, h:.37});
  txt(s, C9, OS12GreyC, {x:3.92, y:5.42, w:2.69, h:.98});
  txt(s, '2019', MEB16BlueC, {x:4.36, y:5.15, w:1.8, h:.37});
  txt(s, C9, OS12GreyC, {x:6.73, y:5.42, w:2.69, h:.98});
  txt(s, '2020', MEB16BlueC, {x:7.18, y:5.15, w:1.8, h:.37});
  txt(s, C15, OS12GreyC, {x:1.67, y:2.21, w:10, h:.67});
  cross(s, 12.11, 1.09, 0);
  ring(s, 1.09, 1.1);
  path(s, {x:1.67, y:3.29, w:1.54, h:1.54, points:g7, fill:{color:TONES.mid}});
  path(s, {x:4.49, y:3.29, w:1.54, h:1.54, points:g7, fill:{color:TONES.mid}});
  path(s, {x:7.3, y:3.29, w:1.54, h:1.54, points:g7, fill:{color:TONES.mid}});
  path(s, {x:10.12, y:3.29, w:1.54, h:1.54, points:g7, fill:{color:TONES.mid}});
}

function slide15(s) {
  txt(s, 'S', MEB500BlueL, {x:1.1, y:-.51, w:4.68, h:8.52});
  txt(s, 'COMPANY\nSTRENGTH', MEB44DarkL, {x:6.88, y:1.1, w:4.68, h:1.58});
  txt(s, C14, OS12GreyJ, {x:6.88, y:3.25, w:5.35, h:2.19});
  txt(s, 'Section Here', MEB16BlueL, {x:6.88, y:2.88, w:1.8, h:.37});
  txt(s, C30, OS12BlueJ, {x:6.88, y:5.43, w:5.35, h:.98});
  ring(s, 6.2, 1.23);
  cross(s, .97, 6.28);
}

function slide16(s) {
  txt(s, 'W', MEB400BlueR, {x:5.11, y:.33, w:7.46, h:6.83});
  txt(s, 'COMPANY\nWEAKNESS', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C5, OS12GreyJ, {x:1.1, y:3.25, w:4.5, h:1.89});
  txt(s, 'Section Here', MEB16BlueL, {x:1.1, y:2.88, w:1.8, h:.37});
  s.addShape('rect', {x:1.1, y:5.7, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:5.76, w:1.43, h:.3});
  ring(s, 11.98, .98);
  cross(s, 5.86, 5.38);
}

function slide17(s) {
  txt(s, 'COMPANY\nOPPORTUNITY', MEB44DarkL, {x:6.66, y:1.1, w:5.56, h:1.58});
  txt(s, 'O', MEB400BlueC, {x:-.38, y:.33, w:7.46, h:6.83});
  txt(s, C26, OS12GreyJ, {x:6.88, y:3.25, w:5.35, h:2.49});
  txt(s, 'Section Here', MEB16BlueL, {x:6.88, y:2.88, w:1.8, h:.37});
  ring(s, 1.19, 6.1);
  cross(s, 5.61, 1.48);
}

function slide18(s) {
  txt(s, 'T', MEB500BlueL, {x:7.55, y:-.51, w:4.68, h:8.52});
  txt(s, 'COMPANY\nTHREAT', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C5, OS12GreyJ, {x:1.1, y:3.25, w:5.36, h:1.58});
  txt(s, 'Section Here', MEB16BlueL, {x:1.1, y:2.88, w:1.8, h:.37});
  txt(s, C1, OS12GreyJ, {x:1.35, y:4.94, w:5.11, h:.67});
  s.addShape('ellipse', {x:1.1, y:5.09, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:1.35, y:5.72, w:5.11, h:.67});
  s.addShape('ellipse', {x:1.1, y:5.87, w:.15, h:.15, fill:{color:BLUE}});
  ring(s, 7.97, 4.69);
  cross(s, 11.33, 3.12);
}

function slide19(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:0, w:6.67, h:7.5, fill:{color:BLUE, transparency:15}});
  s.addShape('rect', {x:0, y:1.98, w:13.33, h:3.56, fill:{color:DARK, transparency:15}});
  txt(s, 'TARGET\nSERVICE', MEB44WhiteR, {x:2.48, y:2.97, w:3.63, h:1.58});
  txt(s, C22, OS12WhiteJ, {x:7.22, y:3.12, w:5.02, h:1.28});
  s.addShape('rect', {x:5.71, y:3.71, w:1.92, h:.1, fill:{color:BLUE}, rotate:90});
  cross(s, 11.82, 4.84);
  ring(s, 1.75, 2.68);
}

function slide20(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:0, w:13.33, h:4.95, fill:{color:BLUE, transparency:15}});
  s.addShape('rect', {x:1.1, y:3.49, w:2.91, h:2.91, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:5.21, y:3.49, w:2.91, h:2.91, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:9.33, y:3.49, w:2.91, h:2.91, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'TARGET SERVICE', MEB44WhiteC, {x:2.06, y:1.1, w:9.22, h:.84});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui official.', OS12WhiteC, {x:1.1, y:2.21, w:11.14, h:.67});
  txt(s, '01', MEB40BlueC, {x:2.13, y:3.93, w:.85, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:1.61, y:4.7, w:1.88, h:.37});
  txt(s, C6, OS12GreyC, {x:1.61, y:4.99, w:1.88, h:.98});
  txt(s, '02', MEB40BlueC, {x:6.11, y:3.93, w:1.11, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:5.73, y:4.7, w:1.88, h:.37});
  txt(s, C6, OS12GreyC, {x:5.73, y:4.99, w:1.88, h:.98});
  txt(s, '03', MEB40BlueC, {x:10.28, y:3.93, w:.99, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:9.84, y:4.7, w:1.88, h:.37});
  txt(s, C6, OS12GreyC, {x:9.84, y:4.99, w:1.88, h:.98});
  ring(s, 1.37, 6.28);
  cross(s, 12.11, 5.44);
}

function slide21(s) {
  s.addShape('rect', {x:6.67, y:4.81, w:5.57, h:1.59, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:6.67, y:1.1, w:5.57, h:1.59, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:6.67, y:2.96, w:5.57, h:1.59, fill:{color:WHITE}, shadow:shadow()});
  photo(s, 6.67, 4.82, 1.59, 1.59);
  photo(s, 6.67, 1.1, 1.59, 1.59);
  photo(s, 6.67, 2.95, 1.59, 1.59);
  txt(s, 'OUR\nTARGET', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  txt(s, C5, OS12GreyJ, {x:1.1, y:3.26, w:4.47, h:1.89});
  txt(s, C20, OS12GreyL, {x:8.63, y:1.4, w:3.24, h:.98});
  txt(s, C20, OS12GreyL, {x:8.63, y:3.26, w:3.24, h:.98});
  txt(s, C20, OS12GreyL, {x:8.63, y:5.12, w:3.24, h:.98});
  ring(s, 6.54, 1.4);
  cross(s, 7.77, 6.27);
  txt(s, 'Section Here', MEB16BlueL, {x:1.1, y:2.98, w:1.8, h:.37});
}

function slide22(s) {
  s.addShape('rect', {x:0, y:0, w:1.65, h:7.5, fill:{color:BLUE}});
  txt(s, 'COMPANY\nPLAN', MEB44DarkL, {x:7.21, y:1.1, w:4.68, h:1.58});
  s.addShape('rect', {x:1.1, y:1.1, w:1.09, h:1.09, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:2.5, w:1.09, h:1.09, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:3.9, w:1.09, h:1.09, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:5.3, w:1.09, h:1.09, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C32, OS12BlueJ, {x:7.21, y:3.14, w:5.02, h:2.19});
  path(s, {x:1.44, y:4.23, w:.44, h:.44, points:g8, fill:{color:BLUE}});
  path(s, {x:1.44, y:1.48, w:.44, h:.34, points:g9, fill:{color:BLUE}});
  path(s, {x:1.57, y:1.56, w:.18, h:.18, points:g10, fill:{color:BLUE}});
  path(s, {x:1.44, y:5.65, w:.44, h:.4, points:g11, fill:{color:BLUE}});
  path(s, {x:1.43, y:2.83, w:.44, h:.44, points:g12, fill:{color:BLUE}});
  path(s, {x:1.58, y:2.9, w:.15, h:.15, points:g13, fill:{color:BLUE}});
  txt(s, C6, OS12GreyJ, {x:2.32, y:1.46, w:3.8, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:2.32, y:1.16, w:1.8, h:.37});
  txt(s, C6, OS12GreyJ, {x:2.32, y:2.85, w:3.8, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:2.32, y:2.56, w:1.8, h:.37});
  txt(s, C6, OS12GreyJ, {x:2.32, y:4.26, w:3.8, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:2.32, y:3.97, w:1.8, h:.37});
  txt(s, C6, OS12GreyJ, {x:2.32, y:5.66, w:3.8, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:2.32, y:5.37, w:1.8, h:.37});
  s.addShape('rect', {x:7.21, y:5.99, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:7.36, y:6.05, w:1.43, h:.3});
  ring(s, 1.86, .98);
  cross(s, 12.11, 6.11);
}

function slide23(s) {
  photo(s, 7.4, 0, 5.93, 6.4);
  s.addShape('rect', {x:0, y:3.6, w:1.82, h:3.9, fill:{color:BLUE}});
  txt(s, 'BUSINESS\nPLAN', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('ellipse', {x:1.1, y:2.89, w:1.45, h:1.45, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:1.1, y:4.95, w:1.45, h:1.45, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C1, OS12GreyJ, {x:2.82, y:3.3, w:3.49, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:2.82, y:2.96, w:2.12, h:.37});
  txt(s, C1, OS12GreyJ, {x:2.82, y:5.36, w:3.49, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:2.82, y:5.02, w:2.12, h:.37});
  path(s, {x:1.49, y:5.34, w:.66, h:.66, points:g8, fill:{color:BLUE}});
  path(s, {x:1.49, y:3.35, w:.66, h:.51, points:g9, fill:{color:BLUE}});
  path(s, {x:1.68, y:3.47, w:.28, h:.28, points:g10, fill:{color:BLUE}});
  s.addShape('rect', {x:7.71, y:-1.4, w:1.1, h:3.9, fill:{color:BLUE}, rotate:90});
  s.addShape('rect', {x:12.24, y:3.6, w:1.1, h:3.9, fill:{color:BLUE}});
  ring(s, 7.28, 1.64);
  cross(s, 8.13, 6.27);
}

function slide24(s) {
  s.addShape('rect', {x:0, y:0, w:5.14, h:5.46, fill:{color:BLUE}});
  s.addShape('rect', {x:1.1, y:1.1, w:1.96, h:2.1, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.16, y:1.1, w:1.96, h:2.1, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:4.3, w:1.96, h:2.1, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.16, y:4.3, w:1.96, h:2.1, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'PLAN', MEB44DarkL, {x:7.22, y:1.1, w:4.47, h:.84});
  txt(s, C10, OS12GreyJ, {x:7.22, y:2.35, w:5.02, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:2.09, w:1.8, h:.37});
  txt(s, C10, OS12GreyJ, {x:7.22, y:3.47, w:5.02, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:3.21, w:1.8, h:.37});
  txt(s, C10, OS12GreyJ, {x:7.22, y:5.72, w:5.02, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:5.46, w:1.8, h:.37});
  txt(s, C10, OS12GreyJ, {x:7.22, y:4.6, w:5.02, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:4.34, w:1.8, h:.37});
  ring(s, 11.98, 1.1);
  cross(s, 12.11, 6.27);
  path(s, {x:4.7, y:4.91, w:.88, h:.88, points:g8, fill:{color:BLUE}});
  path(s, {x:1.6, y:1.77, w:.97, h:.75, points:g9, fill:{color:BLUE}});
  path(s, {x:1.88, y:1.94, w:.4, h:.4, points:g10, fill:{color:BLUE}});
  path(s, {x:1.6, y:4.91, w:.97, h:.89, points:g11, fill:{color:BLUE}});
  path(s, {x:4.65, y:1.6, w:.97, h:.97, points:g12, fill:{color:BLUE}});
  path(s, {x:4.97, y:1.76, w:.32, h:.32, points:g13, fill:{color:BLUE}});
  ring(s, .96, 5.81);
  cross(s, 5.99, 2.62);
}

function slide25(s) {
  s.addShape('rect', {x:9.95, y:0, w:3.39, h:7.5, fill:{color:BLUE}});
  s.addShape('rect', {x:7.68, y:1.1, w:4.56, h:1.4, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.68, y:3.05, w:4.56, h:1.4, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.68, y:5, w:4.56, h:1.4, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'COMPANY\nPLAN', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('ellipse', {x:7.22, y:1.33, w:.93, h:.93, fill:{color:BLUE}});
  s.addShape('ellipse', {x:7.22, y:5.24, w:.93, h:.93, fill:{color:BLUE}});
  s.addShape('ellipse', {x:7.22, y:3.29, w:.93, h:.93, fill:{color:BLUE}});
  path(s, {x:7.45, y:5.48, w:.45, h:.45, points:g8, fill:{color:WHITE}});
  path(s, {x:7.43, y:1.6, w:.49, h:.38, points:g9, fill:{color:WHITE}});
  path(s, {x:7.57, y:1.69, w:.2, h:.2, points:g10, fill:{color:WHITE}});
  path(s, {x:7.43, y:3.52, w:.49, h:.45, points:g11, fill:{color:WHITE}});
  txt(s, C27, OS12GreyJ, {x:8.39, y:1.6, w:3.49, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:8.39, y:1.32, w:2.12, h:.37});
  txt(s, C27, OS12GreyJ, {x:8.39, y:3.56, w:3.49, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:8.39, y:3.27, w:2.12, h:.37});
  txt(s, C27, OS12GreyJ, {x:8.39, y:5.5, w:3.49, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:8.39, y:5.21, w:2.12, h:.37});
  txt(s, C5, OS12GreyJ, {x:1.1, y:2.78, w:5.02, h:1.58});
  txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab', OS12BlueJ, {x:1.1, y:4.51, w:5.02, h:.98});
  s.addShape('rect', {x:1.1, y:5.99, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.05, w:1.43, h:.3});
}

function slide26(s) {
  s.addShape('rect', {x:5.3, y:3.71, w:1.92, h:.1, fill:{color:BLUE}, rotate:90});
  txt(s, C3, OS12GreyJ, {x:7.4, y:1.44, w:4.84, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.4, y:1.1, w:3.5, h:.37});
  txt(s, C3, OS12GreyJ, {x:7.4, y:2.87, w:4.84, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.4, y:2.53, w:3.82, h:.37});
  txt(s, C3, OS12GreyJ, {x:7.4, y:4.29, w:4.84, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.4, y:3.95, w:3.82, h:.37});
  s.addShape('line', {x:7.4, y:3.76, w:4.84, h:0, line:{color:HAIR, width:1.5}, flipV:true});
  s.addShape('line', {x:7.4, y:2.32, w:4.84, h:0, line:{color:HAIR, width:1.5}});
  txt(s, C3, OS12GreyJ, {x:7.4, y:5.72, w:4.84, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:7.4, y:5.38, w:3.82, h:.37});
  s.addShape('line', {x:7.4, y:5.17, w:4.84, h:0, line:{color:HAIR, width:1.5}, flipV:true});
  txt(s, 'OUR\nAWESOME\nPROJECT', MEB44DarkL, {x:1.1, y:2.24, w:4.68, h:2.32});
  txt(s, 'Section Here', MEB16BlueL, {x:1.1, y:1.87, w:1.8, h:.37});
  txt(s, C11, OS12GreyJ, {x:1.1, y:4.66, w:4.01, h:.98});
  ring(s, 1.1, 1.1);
  cross(s, 1.23, 6.27);
}

function slide27(s) {
  photo(s, 0, 0, 6.12, 7.5);
  txt(s, 'OUR\nAWESOME\nPROJECT', MEB44DarkL, {x:7.22, y:1.43, w:4.68, h:2.32});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:1.06, w:1.8, h:.37});
  txt(s, C32, OS12GreyJ, {x:7.21, y:3.88, w:5.02, h:2.19});
  ring(s, 5.99, 1.1);
  cross(s, 7.26, 6.27);
}

function slide28(s) {
  s.addShape('rect', {x:11.34, y:5.35, w:2, h:2.15, fill:{color:BLUE}});
  photo(s, 7.23, 1.1, 5.01, 5.3);
  txt(s, 'AWESOME\nPROJECT', MEB44DarkL, {x:1.1, y:1.48, w:4.68, h:1.58});
  txt(s, 'Section Here', MEB16BlueL, {x:1.1, y:1.11, w:1.8, h:.37});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.', OS12GreyJ, {x:1.1, y:3.46, w:5.02, h:1.89});
  s.addShape('rect', {x:1.1, y:5.99, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.05, w:1.43, h:.3});
  s.addShape('rect', {x:6.25, y:0, w:2.15, h:2.15, fill:{color:BLUE}});
  s.addShape('rect', {x:7.23, y:5.01, w:1.38, h:1.38, fill:{color:BLUE}});
  path(s, {x:7.58, y:5.44, w:.68, h:.53, points:g9, fill:{color:WHITE}});
  path(s, {x:7.78, y:5.56, w:.28, h:.28, points:g10, fill:{color:WHITE}});
  ring(s, 7.1, 3.06);
  cross(s, 12.11, 4.46);
}

function slide29(s) {
  s.addShape('rect', {x:0, y:0, w:3.61, h:7.51, fill:{color:BLUE}});
  photo(s, 1.1, 4.3, 5.02, 2.1);
  photo(s, 1.1, 1.1, 5.02, 2.1);
  s.addShape('ellipse', {x:3, y:1.54, w:1.23, h:1.23, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:3, y:4.74, w:1.23, h:1.23, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:3.31, y:5.06, w:.59, h:.59, points:g8, fill:{color:BLUE}});
  path(s, {x:3.29, y:1.85, w:.65, h:.59, points:g11, fill:{color:BLUE}});
  txt(s, 'AWESOME\nPROJECT', MEB44DarkL, {x:7.22, y:1.48, w:4.68, h:1.58});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:1.11, w:1.8, h:.37});
  txt(s, C33, OS12GreyJ, {x:7.22, y:3.61, w:5.02, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:3.33, w:2.12, h:.37});
  txt(s, C33, OS12GreyJ, {x:7.22, y:5.43, w:5.02, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.22, y:5.15, w:2.12, h:.37});
  ring(s, 6, 1.41);
  cross(s, 5.82, 6.28);
}

function slide30(s) {
  photo(s, 7.23, 1.1, 6.11, 5.3);
  s.addShape('rect', {x:0, y:0, w:1.1, h:7.51, fill:{color:BLUE}});
  txt(s, 'OUR\nAWESOME\nPROJECT', MEB44DarkL, {x:2.21, y:2.24, w:4.68, h:2.32});
  txt(s, 'Section Here', MEB16BlueL, {x:2.21, y:1.87, w:1.8, h:.37});
  txt(s, C11, OS12GreyJ, {x:2.21, y:4.66, w:3.91, h:.98});
  ring(s, 5.86, 1.1);
  cross(s, 2.26, 6.27);
  ring(s, 8.74, 6.27);
  s.addShape('rect', {x:7.23, y:1.1, w:1.38, h:1.38, fill:{color:BLUE}});
  path(s, {x:7.58, y:1.45, w:.67, h:.67, points:g12, fill:{color:WHITE}});
  path(s, {x:7.8, y:1.56, w:.22, h:.22, points:g13, fill:{color:WHITE}});
  s.addShape('rect', {x:11.18, y:5.35, w:2.15, h:2.15, fill:{color:BLUE}});
}

function slide31(s) {
  s.addShape('rect', {x:1.1, y:1.1, w:4.76, h:2.24, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:4.15, w:4.76, h:2.24, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:5.26, y:1.61, w:1.21, h:1.21, fill:{color:BLUE}});
  s.addShape('rect', {x:5.26, y:4.66, w:1.21, h:1.21, fill:{color:BLUE}});
  txt(s, 'OUR\nSERVICES', MEB44DarkL, {x:7.56, y:1.1, w:4.68, h:1.58});
  txt(s, '01', MEB40WhiteC, {x:5.44, y:1.83, w:.85, h:.77});
  txt(s, '02', MEB40WhiteC, {x:5.26, y:4.88, w:1.2, h:.77});
  txt(s, C15, OS12GreyJ, {x:7.56, y:3.21, w:4.67, h:1.28});
  txt(s, '01. Section Here', MEB16BlueL, {x:7.56, y:2.87, w:2.12, h:.37});
  txt(s, C15, OS12GreyJ, {x:7.56, y:5.12, w:4.67, h:1.28});
  txt(s, '02. Section Here', MEB16BlueL, {x:7.56, y:4.78, w:2.26, h:.37});
  txt(s, C20, OS12GreyR, {x:1.56, y:1.73, w:3.24, h:.98});
  txt(s, C20, OS12GreyR, {x:1.56, y:4.78, w:3.24, h:.98});
  ring(s, 1.56, .97);
  cross(s, 5.13, 6.27);
}

function slide32(s) {
  photo(s, 6.05, 4.02, 2.71, 2.38);
  photo(s, 6.05, 1.1, 2.71, 2.38);
  photo(s, 1.1, 3.04, 4.4, 3.36);
  txt(s, 'OUR\nSERVICES', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C3, OS12GreyJ, {x:9.31, y:1.44, w:2.93, h:.98});
  txt(s, '01. Section Here', MEB16BlueL, {x:9.31, y:1.1, w:2.12, h:.37});
  txt(s, C3, OS12GreyJ, {x:9.31, y:3.43, w:2.93, h:.98});
  txt(s, '02. Section Here', MEB16BlueL, {x:9.31, y:3.09, w:2.31, h:.37});
  txt(s, C3, OS12GreyJ, {x:9.31, y:5.42, w:2.93, h:.98});
  txt(s, '03. Section Here', MEB16BlueL, {x:9.31, y:5.08, w:2.31, h:.37});
  s.addShape('line', {x:9.42, y:4.75, w:2.82, h:0, line:{color:HAIR, width:1.5}});
  s.addShape('line', {x:9.42, y:2.74, w:2.82, h:0, line:{color:HAIR, width:1.5}});
  s.addShape('rect', {x:1.1, y:5.43, w:.97, h:.97, fill:{color:BLUE, transparency:15}});
  txt(s, '01', MEB28WhiteC, {x:1.16, y:5.62, w:.85, h:.57});
  s.addShape('rect', {x:6.05, y:2.5, w:.97, h:.97, fill:{color:BLUE, transparency:15}});
  txt(s, '02', MEB28WhiteC, {x:6.11, y:2.7, w:.85, h:.57});
  s.addShape('rect', {x:6.05, y:5.43, w:.97, h:.97, fill:{color:BLUE, transparency:15}});
  txt(s, '03', MEB28WhiteC, {x:6.11, y:5.62, w:.85, h:.57});
  ring(s, 8.24, .97);
  cross(s, 4.78, 6.27);
}

function slide33(s) {
  s.addShape('rect', {x:9.18, y:4.82, w:4.16, h:2.68, fill:{color:BLUE}});
  photo(s, 6.11, 3.17, 6.13, 3.23);
  txt(s, 'OUR\nSERVICES', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('rect', {x:1.1, y:3.17, w:2.61, h:3.24, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.8, y:3.88, w:2.61, h:2.53, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01', MEB40BlueC, {x:1.98, y:3.77, w:.85, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:1.46, y:4.53, w:1.88, h:.37});
  txt(s, C6, OS12GreyC, {x:1.46, y:4.82, w:1.88, h:.98});
  txt(s, '02', MEB40BlueC, {x:5.68, y:4.13, w:.98, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:5.17, y:4.89, w:1.88, h:.37});
  txt(s, C6, OS12GreyC, {x:5.17, y:5.18, w:1.88, h:.98});
  txt(s, C15, OS12GreyJ, {x:6.1, y:1.56, w:6.14, h:.98});
  s.addShape('line', {x:6.19, y:1.12, w:6.04, h:0, line:{color:HAIR, width:1.5}});
  ring(s, 12.11, 3.77);
  cross(s, 3.14, 3.04);
}

function slide34(s) {
  s.addShape('rect', {x:7.97, y:4.64, w:5.36, h:2.86, fill:{color:BLUE}});
  s.addShape('rect', {x:6.67, y:2.51, w:2.61, h:3.9, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'OUR SERVICES', MEB44DarkC, {x:3.37, y:1.1, w:6.6, h:.84});
  s.addShape('rect', {x:9.62, y:2.51, w:2.61, h:3.9, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01. Section Here', MEB16BlueL, {x:1.1, y:2.51, w:2.12, h:.37});
  txt(s, C15, OS12GreyJ, {x:1.1, y:2.8, w:4.47, h:1.28});
  txt(s, '02. Section Here', MEB16BlueL, {x:1.1, y:4.83, w:2.27, h:.37});
  txt(s, C15, OS12GreyJ, {x:1.1, y:5.12, w:4.47, h:1.28});
  s.addShape('rect', {x:7.37, y:2.99, w:1.21, h:1.21, fill:{color:BLUE}});
  txt(s, '01', MEB40WhiteC, {x:7.5, y:3.21, w:.85, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:6.99, y:4.26, w:1.88, h:.37});
  txt(s, C34, OS12GreyC, {x:6.76, y:4.64, w:2.43, h:1.28});
  s.addShape('rect', {x:10.33, y:2.99, w:1.21, h:1.21, fill:{color:BLUE}});
  txt(s, '02', MEB40WhiteC, {x:10.41, y:3.21, w:.96, h:.77});
  txt(s, 'Section Here', MEB16BlueC, {x:9.95, y:4.26, w:1.88, h:.37});
  txt(s, C34, OS12GreyC, {x:9.72, y:4.64, w:2.43, h:1.28});
  cross(s, 7.19, 6.28);
  ring(s, 11.7, 2.38);
}

function slide35(s) {
  s.addShape('rect', {x:9.6, y:0, w:3.73, h:3.75, fill:{color:BLUE}});
  s.addShape('rect', {x:6.12, y:.29, w:1.1, h:13.33, fill:{color:BLUE}, rotate:90});
  txt(s, 'BUSINESS\nLEADER', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  s.addShape('rect', {x:1.1, y:5.19, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:5.25, w:1.43, h:.3});
  txt(s, C5, OS12GreyJ, {x:1.1, y:3.29, w:5.02, h:1.58});
  txt(s, 'Charli Mackie', MEB16BlueL, {x:1.1, y:3, w:1.86, h:.37});
  ring(s, 6.67, 1.1);
  cross(s, 7.55, 4.86);
}

function slide36(s) {
  s.addShape('rect', {x:0, y:0, w:4.17, h:3.75, fill:{color:BLUE}});
  txt(s, 'EXECUTIVE\nDIRECTOR', MEB44DarkL, {x:7.1, y:1.1, w:4.68, h:1.58});
  txt(s, C5, OS12GreyJ, {x:7.1, y:3.34, w:5.14, h:1.58});
  txt(s, 'Mahir Rawlings', MEB16BlueL, {x:7.1, y:3, w:2.07, h:.37});
  txt(s, C35, OS12BlueJ, {x:7.1, y:5.41, w:5.14, h:.98});
  ring(s, 5.85, 1.1);
  cross(s, 11.96, 5.04);
}

function slide37(s) {
  photo(s, 4.01, 4.3, 2.1, 2.1);
  photo(s, 4.01, 1.1, 2.1, 2.1);
  txt(s, 'Fionn Carr', MEB16BlueR, {x:2, y:1.48, w:1.88, h:.37});
  txt(s, C16, OS12GreyR, {x:1.1, y:1.85, w:2.79, h:.98});
  txt(s, 'Kaylen Webber', MEB16BlueR, {x:1.78, y:4.68, w:2.1, h:.37});
  txt(s, C16, OS12GreyR, {x:1.1, y:5.05, w:2.79, h:.98});
  txt(s, 'OUR\nCREATOR', MEB44DarkL, {x:7.22, y:1.1, w:4.68, h:1.58});
  txt(s, C22, OS12GreyJ, {x:7.22, y:2.82, w:5.02, h:1.28});
  s.addShape('line', {x:1.17, y:3.75, w:4.95, h:0, line:{color:HAIR, width:1.5}});
  txt(s, 'Adobe Illustrator', MEB16GreyL, {x:7.35, y:5.84, w:2.27, h:.37});
  s.addShape('rect', {x:7.35, y:5.26, w:4.89, h:.1, fill:{color:GREY}});
  txt(s, 'Adobe Photoshop', MEB16GreyL, {x:7.35, y:4.8, w:2.35, h:.37});
  s.addShape('rect', {x:7.35, y:6.3, w:4.89, h:.1, fill:{color:GREY}});
  ring(s, 5.61, .97);
  cross(s, 4.28, 6.28);
  txt(s, '84%', MEB16GreyR, {x:11.44, y:4.8, w:.79, h:.37});
  txt(s, '54%', MEB16GreyR, {x:11.44, y:5.82, w:.79, h:.37});
  s.addShape('rect', {x:7.35, y:5.26, w:3.6, h:.1, fill:{color:BLUE}});
  s.addShape('rect', {x:7.35, y:6.3, w:2.69, h:.1, fill:{color:BLUE}});
}

function slide38(s) {
  photo(s, 5.34, 2.63, 3.17, 3.78);
  photo(s, 9.06, 1.1, 3.17, 3.78);
  txt(s, 'OUR\nDIVISION', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C14, OS12GreyJ, {x:1.1, y:2.68, w:3.17, h:3.7});
  txt(s, 'Waseem Haynes', MEB16BlueC, {x:5.83, y:1.1, w:2.2, h:.37});
  txt(s, C16, OS12GreyC, {x:5.54, y:1.47, w:2.79, h:.98});
  txt(s, 'Beatrice Day', MEB16BlueC, {x:9.71, y:5.06, w:1.88, h:.37});
  txt(s, C16, OS12GreyC, {x:9.26, y:5.43, w:2.79, h:.98});
  ring(s, 5.22, 5.91);
  cross(s, 12.11, 1.52);
}

function slide39(s) {
  s.addShape('rect', {x:0, y:0, w:3.75, h:7.5, fill:{color:BLUE}});
  ring(s, 4.41, 1.37);
  cross(s, 4.42, 5.75);
  txt(s, 'BUSINESS\nCREW', MEB44DarkL, {x:5.64, y:1.99, w:4.68, h:1.58});
  txt(s, C5, OS12GreyJ, {x:5.64, y:3.58, w:6.6, h:1.28});
  s.addShape('rect', {x:5.64, y:5.09, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:5.79, y:5.15, w:1.43, h:.3});
  s.addShape('line', {x:11.3, y:1.22, w:.25, h:0, line:{color:BLUE, width:3}, rotate:45});
  s.addShape('line', {x:11.3, y:1.22, w:.25, h:0, line:{color:BLUE, width:3}, rotate:315});
  s.addShape('line', {x:11.64, y:1.22, w:.25, h:0, line:{color:BLUE, width:3}, rotate:45});
  s.addShape('line', {x:11.64, y:1.22, w:.25, h:0, line:{color:BLUE, width:3}, rotate:315});
  cross(s, 11.99, 1.09);
  ring(s, 11.29, 6.15);
  ring(s, 11.64, 6.15);
  ring(s, 11.99, 6.15);
  photo(s, 1.1, 1.1, 1.58, 1.58);
  photo(s, 2.96, 1.1, 1.58, 1.58);
  photo(s, 2.96, 2.96, 1.58, 1.58);
  photo(s, 1.1, 2.96, 1.58, 1.58);
  photo(s, 1.1, 4.82, 1.58, 1.58);
  photo(s, 2.96, 4.82, 1.58, 1.58);
}

function slide40(s) {
  s.addShape('rect', {x:5.62, y:2.13, w:1.1, h:1.1, fill:{color:BLUE}});
  s.addShape('rect', {x:5.62, y:5.85, w:1.1, h:1.1, fill:{color:BLUE}});
  s.addShape('rect', {x:11.69, y:3.94, w:1.1, h:1.1, fill:{color:BLUE}});
  txt(s, 'CREATIVE\nCREW', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, 'Hayden Chester', MEB16BlueL, {x:8.93, y:1.21, w:2.09, h:.37});
  txt(s, C16, OS12GreyL, {x:8.93, y:1.58, w:2.79, h:.98});
  txt(s, 'Derek Levy', MEB16BlueR, {x:7.59, y:3.08, w:1.88, h:.37});
  txt(s, C16, OS12GreyR, {x:6.69, y:3.45, w:2.79, h:.98});
  txt(s, 'Austin Burrows', MEB16BlueL, {x:8.93, y:4.92, w:2.09, h:.37});
  txt(s, C16, OS12GreyL, {x:8.93, y:5.29, w:2.79, h:.98});
  txt(s, '01. Section Here', MEB16BlueL, {x:1.1, y:3.42, w:2.12, h:.37});
  txt(s, C11, OS12GreyJ, {x:1.1, y:3.71, w:3.97, h:.98});
  txt(s, '02. Section Here', MEB16BlueL, {x:1.1, y:5.13, w:2.24, h:.37});
  txt(s, C11, OS12GreyJ, {x:1.1, y:5.42, w:3.97, h:.98});
  ring(s, 6.46, .97);
  cross(s, 8.29, 4.96);
  photo(s, 6.17, 1.1, 2.24, 1.62);
  photo(s, 6.17, 4.78, 2.24, 1.62);
  photo(s, 10, 2.94, 2.24, 1.62);
}

function slide41(s) {
  s.addShape('rect', {x:7.93, y:3.75, w:5.4, h:3.75, fill:{color:BLUE}});
  photo(s, 6.4, 1.1, 2.92, 5.31);
  photo(s, 10.42, 1.1, 2.92, 5.31);
  txt(s, 'OUR\nCREATIVE\nCREATOR', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:2.32});
  txt(s, C5, OS12GreyJ, {x:1.1, y:3.57, w:4.21, h:1.89});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  ring(s, 6.28, 1.34);
  cross(s, 10.29, 3.27);
}

function slide42(s) {
  s.addShape('rect', {x:0, y:0, w:3.6, h:3.83, fill:{color:BLUE}});
  txt(s, 'EXECUTIVE\nPRODUCER', MEB44DarkL, {x:7.22, y:1.1, w:4.68, h:1.58});
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit.', OS12GreyJ, {x:7.22, y:3.67, w:5.02, h:1.58});
  txt(s, 'Damien Guerrero', MEB16BlueL, {x:7.22, y:3.3, w:2.27, h:.37});
  s.addShape('line', {x:7.26, y:6.28, w:.25, h:0, line:{color:BLUE, width:3}, rotate:45});
  s.addShape('line', {x:7.26, y:6.28, w:.25, h:0, line:{color:BLUE, width:3}, rotate:315});
  s.addShape('line', {x:7.6, y:6.28, w:.25, h:0, line:{color:BLUE, width:3}, rotate:45});
  s.addShape('line', {x:7.6, y:6.28, w:.25, h:0, line:{color:BLUE, width:3}, rotate:315});
  cross(s, 7.95, 6.15);
  ring(s, 5.44, .97);
  cross(s, 5.99, 5.63);
  photo(s, 1.1, 1.1, 5.02, 5.31);
}

function slide43(s) {
  txt(s, 'OUR\nCREATOR', MEB44DarkL, {x:1.1, y:1.1, w:4.68, h:1.58});
  txt(s, C12, OS12GreyJ, {x:5.94, y:4.13, w:6.29, h:.67});
  txt(s, 'Chace Baxter', MEB16BlueL, {x:5.94, y:3.75, w:2.27, h:.37});
  txt(s, C12, OS12GreyJ, {x:5.94, y:5.73, w:6.29, h:.67});
  txt(s, 'Burhan Mcclain', MEB16BlueL, {x:5.94, y:5.36, w:2.27, h:.37});
  ring(s, 5.82, 2.7);
  cross(s, 3.81, 3.08);
  photo(s, 0, 3.2, 4.85, 4.3);
  photo(s, 5.94, 0, 7.39, 3.2);
}

function slide44(s) {
  s.addShape('rect', {x:0, y:0, w:2.4, h:3.16, fill:{color:BLUE}});
  txt(s, 'OUR\nMEMBER', MEB44DarkL, {x:8.51, y:1.1, w:4.68, h:1.58});
  txt(s, C8, OS12GreyJ, {x:8.51, y:3.03, w:3.72, h:2.19});
  s.addShape('rect', {x:10.51, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:10.66, y:6.04, w:1.43, h:.3});
  txt(s, C36, OS12GreyC, {x:1.34, y:5.73, w:2.13, h:.67});
  txt(s, 'Rico Pace', MEB16BlueC, {x:1.27, y:5.44, w:2.27, h:.37});
  txt(s, C36, OS12GreyC, {x:5.05, y:5.73, w:2.13, h:.67});
  txt(s, 'Shannon Reed', MEB16BlueC, {x:4.98, y:5.44, w:2.27, h:.37});
  cross(s, 3.58, 4.72);
  ring(s, 7.29, 1.58);
  photo(s, 1.1, 1.1, 2.61, 4.12);
  photo(s, 4.81, 1.1, 2.61, 4.12);
}

function slide45(s) {
  photo(s, 1.1, 1.1, 5.57, 3.15);
  s.addShape('rect', {x:1.1, y:1.1, w:5.57, h:3.15, fill:{color:BLUE, transparency:15}});
  txt(s, '“Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit.”', OS12WhiteJ, {x:1.78, y:1.57, w:4.21, h:2.19});
  s.addShape('rect', {x:7.76, y:1.1, w:4.47, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.76, y:3.07, w:4.47, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.76, y:5.05, w:4.47, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C17, OS12GreyJ, {x:8.14, y:1.28, w:3.72, h:.98});
  txt(s, C17, OS12GreyJ, {x:8.14, y:3.26, w:3.72, h:.98});
  txt(s, C17, OS12GreyJ, {x:8.14, y:5.24, w:3.72, h:.98});
  s.addShape('rect', {x:7.49, y:1.1, w:.55, h:.55, fill:{color:BLUE}});
  s.addShape('rect', {x:7.49, y:3.07, w:.55, h:.55, fill:{color:BLUE}});
  s.addShape('rect', {x:7.49, y:5.05, w:.55, h:.55, fill:{color:BLUE}});
  txt(s, '01', MEBdWhiteC, {x:7.4, y:1.17, w:.72, h:.4});
  txt(s, '02', MEBdWhiteC, {x:7.4, y:3.15, w:.72, h:.4});
  txt(s, '03', MEBdWhiteC, {x:7.4, y:5.12, w:.72, h:.4});
  ring(s, 7.64, 2.03);
  cross(s, 6.85, 6.08);
  txt(s, 'YESTERDAY’S HOME RUNS DON’T WIN TODAY’S GAMES.', MEB36DarkL, {x:1.1, y:4.48, w:5.93, h:1.92});
}

function slide46(s) {
  photo(s, 6.67, 1.1, 6.67, 6.4);
  s.addShape('rect', {x:0, y:3.75, w:13.33, h:3.75, fill:{color:BLUE, transparency:15}});
  s.addShape('rect', {x:1.1, y:4.4, w:5.02, h:2, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C1, OS12GreyJ, {x:1.85, y:5.05, w:3.53, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:1.85, y:4.77, w:2.12, h:.37});
  txt(s, 'PLAY BY THE RULES, BUT BE FEROCIOUS', MEB36DarkL, {x:1.09, y:1.1, w:5.2, h:1.92});
  s.addShape('rect', {x:7.21, y:4.4, w:5.02, h:2, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C1, OS12GreyJ, {x:7.96, y:5.05, w:3.53, h:.98});
  txt(s, 'Section Here', MEB16BlueL, {x:7.96, y:4.77, w:2.12, h:.37});
  ring(s, 6.54, 1.37);
  cross(s, .97, 3.25);
}

function slide47(s) {
  photo(s, 0, 0, 6.67, 7.5);
  s.addShape('rect', {x:0, y:0, w:6.67, h:3.75, fill:{color:BLUE, transparency:15}});
  txt(s, 'THE ONLY WAY AROUND IS THROUGH.', MEB36DarkL, {x:7.75, y:1.1, w:5.57, h:1.92});
  txt(s, C5, OS12WhiteJ, {x:1.03, y:1.1, w:4.54, h:1.89});
  txt(s, C21, OS12GreyJ, {x:7.75, y:3.03, w:4.5, h:1.58});
  txt(s, C1, OS12GreyJ, {x:8.1, y:4.91, w:4.14, h:.67});
  s.addShape('ellipse', {x:7.85, y:5.05, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:8.1, y:5.73, w:4.14, h:.67});
  s.addShape('ellipse', {x:7.85, y:5.88, w:.15, h:.15, fill:{color:BLUE}});
  cross(s, 6.54, 6.65);
  ring(s, 11.98, 2.46);
}

function slide48(s) {
  photo(s, 0, 0, 13.33, 3.75);
  s.addShape('rect', {x:0, y:1.09, w:4.81, h:5.31, fill:{color:BLUE, transparency:15}});
  txt(s, 'NEVER DREAMED ABOUT SUCCESS. JUST WORKED FOR IT', MEB36DarkL, {x:5.9, y:1.09, w:6.33, h:1.92});
  txt(s, C22, OS12WhiteJ, {x:1.09, y:2.5, w:2.63, h:2.49});
  ring(s, 12.1, 3.62);
  cross(s, 5.05, 6.1);
  txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione.', OS12GreyJ, {x:5.9, y:4.36, w:6.46, h:1.58});
}

function slide49(s) {
  photo(s, 9.52, 0, 3.81, 7.5);
  txt(s, 'THE WAY TO GET STARTED IS TO QUIT TALKING AND BEGIN DOING', MEB36DarkL, {x:1.08, y:1.1, w:7.54, h:1.92});
  txt(s, C8, OS12GreyJ, {x:1.08, y:3.25, w:3.55, h:2.19});
  s.addShape('rect', {x:1.08, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.23, y:6.04, w:1.43, h:.3});
  s.addShape('rect', {x:5.71, y:3.03, w:5.71, h:4.47, fill:{color:BLUE, transparency:15}});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in.', OS12WhiteJ, {x:6.8, y:4.17, w:3.55, h:2.19});
  ring(s, 9.4, 1.1);
  cross(s, 4.73, 6.1);
}

function slide50(s) {
  s.addShape('rect', {x:0, y:0, w:4.29, h:7.5, fill:{color:BLUE}});
  photo(s, 1.36, 1.38, 5.85, 3.32);
  photo(s, 1.1, 1.1, 6.39, 5.31, 'device');
  txt(s, 'MOCKUP\nDEVICE', MEB44DarkL, {x:8.03, y:1.1, w:4.2, h:1.58});
  txt(s, C5, OS12GreyJ, {x:8.03, y:2.81, w:4.21, h:1.89});
  txt(s, C35, OS12BlueJ, {x:8.03, y:4.82, w:4.2, h:1.28});
  ring(s, 11.98, 1.1);
  cross(s, 7.08, 6.1);
}

function slide51(s) {
  s.addShape('rect', {x:0, y:2.19, w:5.89, h:3.13, fill:{color:BLUE}});
  path(s, {x:4.52, y:1.17, w:2.73, h:5.13, points:g14, fill:{color:TONES.mid}});
  path(s, {x:1.2, y:1.17, w:2.73, h:5.13, points:g14, fill:{color:TONES.mid}});
  photo(s, 1.09, 1.09, 3, 5.31, 'device');
  photo(s, 4.39, 1.09, 3, 5.31, 'device');
  txt(s, 'MOCKUP\nDEVICE', MEB44DarkL, {x:7.93, y:1.1, w:4.2, h:1.58});
  txt(s, C1, OS12GreyJ, {x:8.19, y:2.81, w:4.05, h:.98});
  s.addShape('ellipse', {x:7.93, y:2.96, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:8.19, y:4.12, w:4.05, h:.98});
  s.addShape('ellipse', {x:7.93, y:4.26, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:8.19, y:5.43, w:4.05, h:.98});
  s.addShape('ellipse', {x:7.93, y:5.57, w:.15, h:.15, fill:{color:BLUE}});
  cross(s, 12.12, 1.14);
  ring(s, 11.48, 2.26);
}

function slide52(s) {
  s.addShape('rect', {x:9.4, y:4.82, w:3.93, h:2.68, fill:{color:BLUE}});
  photo(s, 7.2, 3.26, 4.4, 2.78);
  photo(s, 6.58, 3.13, 5.66, 3.28, 'device');
  txt(s, 'MOCKUP\nDEVICE', MEB44DarkL, {x:1.1, y:1.1, w:4.2, h:1.58});
  txt(s, '01. Section Here', MEB16BlueL, {x:1.1, y:2.84, w:2.12, h:.37});
  txt(s, C11, OS12GreyJ, {x:1.1, y:3.13, w:4.93, h:.98});
  txt(s, '02. Section Here', MEB16BlueL, {x:1.1, y:4.45, w:2.24, h:.37});
  txt(s, C11, OS12GreyJ, {x:1.1, y:4.74, w:4.93, h:.98});
  s.addShape('rect', {x:1.1, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qu.', OS12BlueJ, {x:6.57, y:1.1, w:5.66, h:1.28});
  ring(s, 4.18, 6.15);
  cross(s, 5.35, 1.21);
}

function slide53(s) {
  s.addShape('rect', {x:0, y:0, w:6.21, h:7.5, fill:{color:BLUE}});
  photo(s, 1.22, 3.78, 3.05, 1.67);
  path(s, {x:5.08, y:2.2, w:2.22, h:4.09, points:g14, fill:{color:TONES.mid}});
  photo(s, 1.1, 3.66, 3.31, 2.75, 'device');
  photo(s, 5, 2.12, 2.42, 4.29, 'device');
  txt(s, 'MOCKUP\nDEVICE', MEB44WhiteL, {x:1.1, y:1.1, w:4.2, h:1.58});
  txt(s, C22, OS12GreyJ, {x:8.51, y:1.09, w:3.73, h:1.89});
  txt(s, C28, OS12GreyJ, {x:9.6, y:3.49, w:2.64, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:9.6, y:3.15, w:2.12, h:.37});
  s.addShape('ellipse', {x:8.51, y:3.21, w:.9, h:.9, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:8.75, y:3.5, w:.41, h:.32, points:g9, fill:{color:BLUE}});
  path(s, {x:8.87, y:3.57, w:.17, h:.17, points:g10, fill:{color:BLUE}});
  txt(s, C28, OS12GreyJ, {x:9.6, y:5.73, w:2.64, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:9.6, y:5.39, w:2.12, h:.37});
  s.addShape('ellipse', {x:8.51, y:5.45, w:.9, h:.9, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C28, OS12GreyJ, {x:9.6, y:4.61, w:2.64, h:.67});
  txt(s, 'Section Here', MEB16BlueL, {x:9.6, y:4.27, w:2.12, h:.37});
  s.addShape('ellipse', {x:8.51, y:4.33, w:.9, h:.9, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:8.79, y:5.73, w:.34, h:.34, points:g8, fill:{color:BLUE}});
  path(s, {x:8.76, y:4.61, w:.37, h:.34, points:g11, fill:{color:BLUE}});
  cross(s, .97, 6.35);
  s.addShape('ellipse', {x:4.02, y:3.08, w:.25, h:.25, fill:{type:'none'}, line:{color:WHITE, width:3}});
}

function slide54(s) {
  s.addShape('rect', {x:7.32, y:0, w:6.01, h:3.08, fill:{color:BLUE}});
  path(s, {x:9.39, y:1.15, w:2.71, h:5.13, points:g14, fill:{color:TONES.mid}});
  photo(s, 6.11, 1.48, 2.39, 3.22);
  photo(s, 9.24, 1.09, 2.99, 5.31, 'device');
  photo(s, 6.02, 1.1, 2.6, 3.98, 'device');
  txt(s, C8, OS12GreyJ, {x:1.1, y:2.81, w:3.82, h:2.19});
  txt(s, 'MOCKUP\nDEVICE', MEB44DarkL, {x:1.1, y:1.1, w:4.2, h:1.58});
  s.addShape('rect', {x:1.1, y:5.33, w:7.52, h:1.08, fill:{color:BLUE}});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti.', OS12WhiteJ, {x:1.54, y:5.53, w:6.64, h:.67});
  ring(s, 8.81, 6.07);
  cross(s, 4.56, 2.2);
}

function slide55(s) {
  s.addShape('rect', {x:0, y:2.85, w:13.33, h:2.09, fill:{color:BLUE}});
  photo(s, -1.65, 1.27, 7.12, 4.57);
  photo(s, -2.67, 1.1, 9.17, 5.31, 'device');
  txt(s, 'MOCKUP\nDEVICE', MEB44DarkL, {x:6.67, y:1.1, w:4.2, h:1.58});
  txt(s, C21, OS12WhiteJ, {x:6.71, y:3.26, w:5.57, h:1.28});
  txt(s, C25, OS12GreyJ, {x:6.67, y:5.15, w:5.57, h:.67});
  ring(s, 12.02, 6.15);
  cross(s, 12.02, 1.05);
}

function slide56(s) {
  txt(s, 'BUSINESS\nPICTURE', MEB44DarkL, {x:1.1, y:2.02, w:4.2, h:1.58});
  txt(s, C8, OS12GreyJ, {x:1.08, y:3.6, w:4.11, h:1.89});
  s.addShape('rect', {x:1.22, y:1.79, w:1.92, h:.1, fill:{color:BLUE}});
  s.addShape('rect', {x:1.22, y:5.71, w:1.92, h:.1, fill:{color:BLUE}});
  cross(s, 6.36, 1.05);
  ring(s, 12.02, 6.15);
  path(s, {x:5.74, y:1.1, w:6.49, h:5.31, points:g15, fill:{color:TONES.mid}});
}

function slide57(s) {
  path(s, {x:5.79, y:1.1, w:6.45, h:5.31, points:g16, fill:{color:TONES.mid}});
  s.addShape('rect', {x:0, y:4.31, w:5.3, h:2.09, fill:{color:BLUE}});
  txt(s, 'BUSINESS\nPICTURE', MEB44DarkL, {x:1.1, y:1.1, w:4.2, h:1.58});
  txt(s, C19, OS12GreyJ, {x:1.1, y:2.79, w:3.6, h:1.28});
  txt(s, C19, OS12WhiteJ, {x:1.1, y:4.71, w:3.6, h:1.28});
  cross(s, 8.45, .97);
  ring(s, 5.79, 6.27);
}

function slide58(s) {
  txt(s, 'BUSINESS\nPICTURE', MEB44DarkL, {x:1.1, y:1.1, w:4.2, h:1.58});
  s.addShape('rect', {x:7.68, y:1.1, w:5.65, h:1.58, fill:{color:BLUE}});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.', OS12WhiteJ, {x:8.23, y:1.34, w:4.01, h:.98});
  txt(s, '01. Section Here', MEB16BlueL, {x:8.23, y:3.2, w:2.12, h:.37});
  txt(s, C11, OS12GreyJ, {x:8.23, y:3.49, w:4.01, h:.98});
  txt(s, '01. Section Here', MEB16BlueL, {x:8.23, y:5.14, w:2.12, h:.37});
  txt(s, C11, OS12GreyJ, {x:8.23, y:5.43, w:4.01, h:.98});
  ring(s, 5.07, 5.51);
  cross(s, 6.54, 1.76);
  path(s, {x:1.1, y:3.12, w:6.58, h:3.28, points:g17, fill:{color:TONES.mid}});
}

function slide59(s) {
  s.addShape('rect', {x:0, y:5.55, w:5.5, h:1.95, fill:{color:BLUE}});
  txt(s, 'BUSINESS\nPICTURE', MEB44DarkL, {x:1.1, y:1.1, w:4.2, h:1.58});
  txt(s, C5, OS12GreyJ, {x:1.09, y:2.81, w:5.29, h:1.58});
  s.addShape('rect', {x:1.09, y:4.69, w:1.72, h:1.71, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.66, y:4.69, w:1.72, h:1.71, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('line', {x:3.74, y:4.69, w:0, h:1.71, line:{color:HAIR, width:1.5}, flipV:true});
  path(s, {x:5.09, y:5.22, w:.85, h:.66, points:g9, fill:{color:BLUE}});
  path(s, {x:5.34, y:5.37, w:.35, h:.35, points:g10, fill:{color:BLUE}});
  path(s, {x:1.53, y:5.12, w:.85, h:.85, points:g12, fill:{color:BLUE}});
  path(s, {x:1.81, y:5.26, w:.28, h:.28, points:g13, fill:{color:BLUE}});
  cross(s, 6.43, 1.15);
  ring(s, 7.55, 3.16);
  path(s, {x:6.93, y:-.98, w:8.44, h:9.8, points:g18, fill:{color:TONES.mid}});
}

function slide60(s) {
  s.addShape('rect', {x:0, y:0, w:2.53, h:3.75, fill:{color:BLUE}});
  txt(s, 'BUSINESS\nPICTURE', MEB44DarkL, {x:7.97, y:1.1, w:4.2, h:1.58});
  txt(s, C26, OS12BlueJ, {x:7.97, y:2.68, w:4.26, h:3.1});
  s.addShape('rect', {x:10.51, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:10.66, y:6.04, w:1.43, h:.3});
  cross(s, 1.12, 6.22);
  ring(s, 6.62, 1.06);
  path(s, {x:1.1, y:1.1, w:5.77, h:5.31, points:g19, fill:{color:TONES.mid}});
}

function slide61(s) {
  txt(s, 'COMPANY ACHIEVEMENT', MEB44DarkL, {x:1.97, y:1.1, w:9.4, h:.84});
  s.addShape('ellipse', {x:1.73, y:2.45, w:2.31, h:2.31, fill:{color:BLUE}});
  txt(s, C17, OS12GreyC, {x:1.1, y:5.43, w:3.59, h:.98});
  txt(s, 'Achievement One', MEB16BlueC, {x:1.7, y:5.12, w:2.39, h:.37});
  s.addShape('ellipse', {x:5.51, y:2.45, w:2.31, h:2.31, fill:{color:BLUE}});
  txt(s, C17, OS12GreyC, {x:4.87, y:5.43, w:3.59, h:.98});
  txt(s, 'Achievement Two', MEB16BlueC, {x:5.47, y:5.12, w:2.39, h:.37});
  s.addShape('ellipse', {x:9.29, y:2.45, w:2.31, h:2.31, fill:{color:BLUE}});
  txt(s, C17, OS12GreyC, {x:8.65, y:5.43, w:3.59, h:.98});
  txt(s, 'Achievement Three', MEB16BlueC, {x:9.12, y:5.12, w:2.64, h:.37});
  path(s, {x:2.21, y:3.21, w:1.36, h:1.09, points:g20, fill:{color:WHITE}});
  path(s, {x:5.98, y:3.21, w:1.36, h:1.09, points:g20, fill:{color:WHITE}});
  path(s, {x:9.75, y:3.21, w:1.36, h:1.09, points:g20, fill:{color:WHITE}});
  ring(s, 1.1, 4.63);
  cross(s, 11.93, 2);
}

function slide62(s) {
  txt(s, 'COMPANY\nACHIEVEMENT', MEB44DarkL, {x:1.1, y:1.1, w:5.57, h:1.58});
  txt(s, C2, OS12GreyL, {x:3.1, y:3.56, w:3.29, h:.67});
  txt(s, C2, OS12GreyL, {x:3.1, y:5.43, w:3.29, h:.67});
  s.addShape('ellipse', {x:1.1, y:2.96, w:1.58, h:1.58, fill:{color:BLUE}});
  s.addShape('ellipse', {x:1.1, y:4.82, w:1.58, h:1.58, fill:{color:BLUE}});
  txt(s, 'Achievement One', MEB16BlueL, {x:3.1, y:3.26, w:2.39, h:.37});
  txt(s, 'Achievement Two', MEB16BlueL, {x:3.1, y:5.12, w:2.39, h:.37});
  path(s, {x:1.39, y:3.35, w:.99, h:.79, points:g20, fill:{color:WHITE}});
  path(s, {x:1.39, y:5.21, w:.99, h:.79, points:g20, fill:{color:WHITE}});
  txt(s, C2, OS12GreyL, {x:8.95, y:3.56, w:3.29, h:.67});
  txt(s, C2, OS12GreyL, {x:8.95, y:5.43, w:3.29, h:.67});
  s.addShape('ellipse', {x:6.95, y:2.96, w:1.58, h:1.58, fill:{color:BLUE}});
  s.addShape('ellipse', {x:6.95, y:4.82, w:1.58, h:1.58, fill:{color:BLUE}});
  txt(s, 'Achievement Three', MEB16BlueL, {x:8.95, y:3.26, w:2.57, h:.37});
  txt(s, 'Achievement Four', MEB16BlueL, {x:8.95, y:5.12, w:2.67, h:.37});
  path(s, {x:7.24, y:3.35, w:.99, h:.79, points:g20, fill:{color:WHITE}});
  path(s, {x:7.24, y:5.21, w:.99, h:.79, points:g20, fill:{color:WHITE}});
  ring(s, 6.41, 1.1);
  cross(s, 11.93, 3.05);
  txt(s, C37, OS12GreyL, {x:7.97, y:1.1, w:4.26, h:1.28});
}

function slide63(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:1.09, w:13.33, h:1.94, fill:{color:BLUE}});
  s.addShape('rect', {x:0, y:3.03, w:13.33, h:3.37, fill:{color:WHITE, transparency:15}});
  txt(s, 'COMPANY ACHIEVEMENT', MEB44WhiteL, {x:1.97, y:1.64, w:9.4, h:.84});
  txt(s, C2, OS12GreyC, {x:1.09, y:5.2, w:3.29, h:.67});
  txt(s, 'Achievement One', MEB16BlueC, {x:1.54, y:4.89, w:2.39, h:.37});
  s.addShape('ellipse', {x:2.08, y:3.56, w:1.31, h:1.31, fill:{color:BLUE}});
  txt(s, '01', MEB32WhiteC, {x:2.2, y:3.9, w:1.06, h:.64});
  txt(s, C2, OS12GreyC, {x:5.02, y:5.2, w:3.29, h:.67});
  txt(s, 'Achievement Two', MEB16BlueC, {x:5.47, y:4.89, w:2.39, h:.37});
  s.addShape('ellipse', {x:6.01, y:3.56, w:1.31, h:1.31, fill:{color:BLUE}});
  txt(s, '02', MEB32WhiteC, {x:6.14, y:3.9, w:1.06, h:.64});
  txt(s, C2, OS12GreyC, {x:8.96, y:5.2, w:3.29, h:.67});
  txt(s, 'Achievement Three', MEB16BlueC, {x:9.32, y:4.89, w:2.56, h:.37});
  s.addShape('ellipse', {x:9.95, y:3.56, w:1.31, h:1.31, fill:{color:BLUE}});
  txt(s, '03', MEB32WhiteC, {x:10.07, y:3.9, w:1.06, h:.64});
  ring(s, 1.14, 6.27);
  cross(s, 11.76, 3.33);
}

function slide64(s) {
  s.addShape('rect', {x:0, y:3.9, w:3.37, h:3.6, fill:{color:BLUE}});
  txt(s, C2, OS12GreyC, {x:8.95, y:2.74, w:3.29, h:.67});
  txt(s, 'Achievement One', MEB16BlueC, {x:9.4, y:2.44, w:2.39, h:.37});
  s.addShape('ellipse', {x:9.94, y:1.11, w:1.31, h:1.31, fill:{color:BLUE}});
  txt(s, '01', MEB32WhiteC, {x:10.06, y:1.44, w:1.06, h:.64});
  txt(s, 'COMPANY\nACHIEVEMENT', MEB44DarkL, {x:1.1, y:1.1, w:5.57, h:1.58});
  txt(s, C2, OS12GreyC, {x:8.95, y:5.73, w:3.29, h:.67});
  txt(s, 'Achievement Two', MEB16BlueC, {x:9.4, y:5.43, w:2.39, h:.37});
  s.addShape('ellipse', {x:9.94, y:4.09, w:1.31, h:1.31, fill:{color:BLUE}});
  txt(s, '02', MEB32WhiteC, {x:10.06, y:4.43, w:1.06, h:.64});
  s.addShape('line', {x:7.85, y:1.1, w:0, h:5.31, line:{color:HAIR, width:1.5}, flipV:true});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque.', OS12GreyJ, {x:3.83, y:5.12, w:2.93, h:1.28});
  cross(s, 8.32, .98);
  ring(s, 7.28, 6.15);
  photo(s, 0, 2.84, 6.76, 2.13);
}

function slide65(s) {
  s.addShape('rect', {x:0, y:2.79, w:13.33, h:1.1, fill:{color:BLUE}});
  s.addShape('ellipse', {x:1.1, y:2.25, w:2.18, h:2.18, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'COMPANY ACHIEVEMENT', MEB44DarkL, {x:1.97, y:1.1, w:9.4, h:.84});
  s.addShape('ellipse', {x:4.08, y:2.25, w:2.18, h:2.18, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:7.07, y:2.25, w:2.18, h:2.18, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:10.05, y:2.25, w:2.18, h:2.18, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C4, OS12GreyC, {x:1.1, y:5.12, w:2.18, h:1.28});
  txt(s, 'Achievement', MEB16BlueC, {x:1.16, y:4.79, w:2.05, h:.37});
  txt(s, C4, OS12GreyC, {x:4.08, y:5.12, w:2.18, h:1.28});
  txt(s, 'Achievement', MEB16BlueC, {x:4.15, y:4.79, w:2.05, h:.37});
  txt(s, C4, OS12GreyC, {x:7.07, y:5.12, w:2.18, h:1.28});
  txt(s, 'Achievement', MEB16BlueC, {x:7.13, y:4.79, w:2.05, h:.37});
  txt(s, C4, OS12GreyC, {x:10.05, y:5.12, w:2.18, h:1.28});
  txt(s, 'Achievement', MEB16BlueC, {x:10.12, y:4.79, w:2.05, h:.37});
  path(s, {x:1.5, y:2.85, w:1.38, h:1.1, points:g20, fill:{color:BLUE}});
  path(s, {x:4.48, y:2.85, w:1.38, h:1.1, points:g20, fill:{color:BLUE}});
  path(s, {x:7.47, y:2.85, w:1.38, h:1.1, points:g20, fill:{color:BLUE}});
  path(s, {x:10.45, y:2.85, w:1.38, h:1.1, points:g20, fill:{color:BLUE}});
  cross(s, 1.15, 1.15);
  ring(s, 11.89, 1.1);
}

function slide66(s) {
  path(s, {x:1.09, y:2.8, w:5.51, h:3.61, points:g21, fill:{color:SILVER}});
  txt(s, 'ASIA WORLD\nMAP', MEB44DarkL, {x:1.09, y:1.09, w:5.49, h:1.58});
  path(s, {x:4.62, y:5.63, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:4.75, y:5.76, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:4.33, y:4.61, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:4.46, y:4.74, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:1.81, y:4.52, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:1.94, y:4.65, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:3.04, y:3.14, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:3.17, y:3.27, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  txt(s, C4, OS12GreyL, {x:7.68, y:1.32, w:4.62, h:.67});
  txt(s, 'Country One', MEB16BlueL, {x:7.68, y:1.09, w:2.05, h:.37});
  txt(s, C4, OS12GreyL, {x:7.68, y:3.53, w:4.62, h:.67});
  txt(s, 'Country Three', MEB16BlueL, {x:7.68, y:3.3, w:2.05, h:.37});
  txt(s, C4, OS12GreyL, {x:7.68, y:4.63, w:4.62, h:.67});
  txt(s, 'Country Four', MEB16BlueL, {x:7.68, y:4.4, w:2.05, h:.37});
  txt(s, C4, OS12GreyL, {x:7.68, y:5.74, w:4.62, h:.67});
  txt(s, 'Country Five', MEB16BlueL, {x:7.68, y:5.51, w:2.05, h:.37});
  txt(s, C4, OS12GreyL, {x:7.68, y:2.43, w:4.62, h:.67});
  txt(s, 'Country Two', MEB16BlueL, {x:7.68, y:2.2, w:2.05, h:.37});
  path(s, {x:4.54, y:3.32, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:4.67, y:3.45, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  cross(s, 1.14, 6.11);
  ring(s, 6.88, 1.09);
}

function slide67(s) {
  path(s, {x:5.52, y:1.1, w:6.73, h:5.32, points:g24, fill:{color:SILVER}});
  path(s, {x:9.77, y:2.61, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:9.9, y:2.74, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:10.66, y:4.37, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:10.79, y:4.5, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:6.79, y:4.06, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:6.92, y:4.19, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  txt(s, 'AUSTRALIA\nWORLD MAP', MEB44DarkL, {x:1.09, y:1.09, w:5.49, h:1.58});
  s.addShape('rect', {x:1.09, y:2.75, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'CITY NAME', MEB12WhiteC, {x:1.24, y:2.81, w:1.43, h:.3});
  txt(s, C29, OS12GreyJ, {x:1.09, y:3.17, w:3.78, h:.67});
  s.addShape('rect', {x:1.09, y:4.03, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'CITY NAME', MEB12WhiteC, {x:1.24, y:4.09, w:1.43, h:.3});
  txt(s, C29, OS12GreyJ, {x:1.09, y:4.45, w:3.78, h:.67});
  s.addShape('rect', {x:1.09, y:5.31, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'CITY NAME', MEB12WhiteC, {x:1.24, y:5.37, w:1.43, h:.3});
  txt(s, C29, OS12GreyJ, {x:1.09, y:5.73, w:3.78, h:.67});
  cross(s, 5.72, 1.27);
  ring(s, 8.62, 6.17);
}

function slide68(s) {
  path(s, {x:5, y:1.46, w:7.24, h:4.57, points:g25, fill:{color:SILVER}});
  path(s, {x:11.79, y:1.92, w:.02, h:.07, points:g26, fill:{color:SILVER}});
  path(s, {x:11.97, y:2.24, w:.02, h:.04, points:g27, fill:{color:SILVER}});
  path(s, {x:11.98, y:2.55, w:.05, h:.03, points:g28, fill:{color:SILVER}});
  txt(s, 'AMERICA\nWORLD\nMAP', MEB44DarkL, {x:1.09, y:1.09, w:5.49, h:2.32});
  txt(s, C5, OS12GreyL, {x:1.09, y:3.55, w:3.51, h:2.49});
  path(s, {x:7.32, y:4.09, w:2.07, h:1.95, points:g29, fill:{color:BLUE}});
  path(s, {x:5, y:2.7, w:1.17, h:1.95, points:g30, fill:{color:BLUE}});
  path(s, {x:6.04, y:3.88, w:.96, h:1.12, points:g31, fill:{color:BLUE}});
  path(s, {x:5.52, y:2.81, w:.87, h:1.39, points:g32, fill:{color:BLUE}});
  cross(s, 11.94, 6.11);
  ring(s, 4.4, 2.34);
}

function slide69(s) {
  s.addShape('rect', {x:0, y:2.68, w:13.33, h:.43, fill:{color:BLUE}});
  s.addShape('rect', {x:0, y:4.85, w:13.33, h:.43, fill:{color:BLUE}});
  path(s, {x:7.5, y:1.09, w:4.74, h:5.32, points:g33, fill:{color:SILVER}});
  cross(s, 11.94, 6.11);
  ring(s, 8.57, 4);
  txt(s, 'AFRICA MAP', MEB44DarkL, {x:1.09, y:1.09, w:5.04, h:.84});
  path(s, {x:8.82, y:1.37, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:8.95, y:1.5, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:8.14, y:2.66, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:8.27, y:2.79, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:10.81, y:3.44, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:10.94, y:3.57, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:9.98, y:5.25, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:10.11, y:5.38, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  s.addShape('ellipse', {x:1.56, y:2.28, w:1.21, h:1.21, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'CITY\nNAME', MEB12BlueC, {x:1.45, y:2.64, w:1.43, h:.5});
  txt(s, C23, OS12GreyC, {x:1.07, y:3.57, w:2.18, h:.67});
  s.addShape('ellipse', {x:1.56, y:4.45, w:1.21, h:1.21, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'CITY\nNAME', MEB12BlueC, {x:1.45, y:4.81, w:1.43, h:.5});
  txt(s, C23, OS12GreyC, {x:1.07, y:5.73, w:2.18, h:.67});
  s.addShape('ellipse', {x:4.48, y:2.28, w:1.21, h:1.21, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'CITY\nNAME', MEB12BlueC, {x:4.37, y:2.64, w:1.43, h:.5});
  txt(s, C23, OS12GreyC, {x:3.99, y:3.57, w:2.18, h:.67});
  s.addShape('ellipse', {x:4.48, y:4.45, w:1.21, h:1.21, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'CITY\nNAME', MEB12BlueC, {x:4.37, y:4.81, w:1.43, h:.5});
  txt(s, C23, OS12GreyC, {x:3.99, y:5.73, w:2.18, h:.67});
}

function slide70(s) {
  txt(s, 'EUROPE\nWORLD\nMAP', MEB44DarkL, {x:1.09, y:1.09, w:5.49, h:2.32});
  path(s, {x:5.76, y:2.01, w:0, h:0, points:g34, fill:{color:SILVER}});
  path(s, {x:5.76, y:2.01, w:1.02, h:.47, points:g35, fill:{color:SILVER}});
  path(s, {x:7.04, y:1.09, w:5.2, h:5.32, points:g36, fill:{color:SILVER}});
  path(s, {x:7.39, y:3.29, w:.76, h:1.25, points:g37, fill:{color:SILVER}});
  path(s, {x:6.97, y:3.86, w:.43, h:.46, points:g38, fill:{color:SILVER}});
  path(s, {x:9.25, y:6.17, w:.23, h:.14, points:g39, fill:{color:SILVER}});
  path(s, {x:9.1, y:3.75, w:.11, h:.2, points:g40, fill:{color:SILVER}});
  path(s, {x:7.26, y:3.36, w:.08, h:.09, points:g41, fill:{color:SILVER}});
  path(s, {x:7.31, y:3.5, w:.05, h:.05, points:g42, fill:{color:SILVER}});
  path(s, {x:8.83, y:5.56, w:.11, h:.18, points:g43, fill:{color:SILVER}});
  path(s, {x:8.81, y:5.8, w:.14, h:.28, points:g44, fill:{color:SILVER}});
  photo(s, 7.49, 4.36, 1.4, 1.38, 'mapblue');
  photo(s, 8.52, 1.24, 1.87, 2.18, 'mapblue');
  s.addShape('rect', {x:1.09, y:3.98, w:4.27, h:.91, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.09, y:3.98, w:1.55, h:.91, fill:{color:BLUE}});
  txt(s, 'FRANCE', MEB12WhiteC, {x:1.33, y:4.28, w:1.12, h:.3});
  s.addShape('rect', {x:1.09, y:5.65, w:4.27, h:.91, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.09, y:5.65, w:1.55, h:.91, fill:{color:BLUE}});
  txt(s, 'NORWEGIA', MEB12WhiteC, {x:1.27, y:5.95, w:1.25, h:.3});
  txt(s, C7, OS12GreyL, {x:2.7, y:4.1, w:2.42, h:.67});
  txt(s, C7, OS12GreyL, {x:2.7, y:5.76, w:2.42, h:.67});
  cross(s, 7.36, 1.21);
  ring(s, 11.69, 3.61);
}

function slide71(s) {
  path(s, {x:1.09, y:2.53, w:5.72, h:3.88, points:g45, fill:{color:SILVER}});
  txt(s, 'UKRAINE MAP', MEB44DarkC, {x:3.96, y:1.09, w:5.41, h:.84});
  path(s, {x:1.76, y:3.75, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:1.89, y:3.88, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:5.1, y:4.71, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:5.23, y:4.84, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:4.11, y:3.07, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:4.24, y:3.2, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  txt(s, C37, OS12GreyL, {x:7.56, y:2.53, w:4.68, h:1.28});
  s.addShape('rect', {x:7.67, y:4.27, w:4.57, h:.08, fill:{color:SILVER}});
  s.addShape('rect', {x:7.67, y:4.27, w:2.84, h:.08, fill:{color:BLUE}});
  txt(s, 'City Progress', MEB12BlueL, {x:7.56, y:3.97, w:1.4, h:.3});
  s.addShape('rect', {x:7.67, y:4.95, w:4.57, h:.08, fill:{color:SILVER}});
  s.addShape('rect', {x:7.67, y:4.95, w:2.12, h:.08, fill:{color:BLUE}});
  txt(s, 'City Progress', MEB12BlueL, {x:7.56, y:4.64, w:1.4, h:.3});
  s.addShape('rect', {x:7.67, y:5.62, w:4.57, h:.08, fill:{color:SILVER}});
  s.addShape('rect', {x:7.67, y:5.62, w:3.74, h:.08, fill:{color:BLUE}});
  txt(s, 'City Progress', MEB12BlueL, {x:7.56, y:5.32, w:1.4, h:.3});
  txt(s, '60%', MEB12GreyR, {x:11.66, y:3.97, w:.57, h:.3});
  txt(s, '45%', MEB12GreyR, {x:11.66, y:4.64, w:.57, h:.3});
  txt(s, '86%', MEB12GreyR, {x:11.66, y:5.32, w:.57, h:.3});
  s.addShape('rect', {x:10.51, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:10.66, y:6.04, w:1.43, h:.3});
  cross(s, 1.14, 6.13);
  ring(s, 1.09, 1.2);
}

function slide72(s) {
  path(s, {x:4.73, y:1.93, w:7.35, h:4.47, points:g46, fill:{color:SILVER}});
  path(s, {x:7.05, y:3.26, w:.73, h:.55, points:g47, fill:{color:SILVER}});
  path(s, {x:8.34, y:2.83, w:.35, h:.4, points:g48, fill:{color:SILVER}});
  path(s, {x:9.78, y:2.85, w:.15, h:.11, points:g49, fill:{color:SILVER}});
  path(s, {x:9.51, y:2.94, w:.24, h:.23, points:g50, fill:{color:SILVER}});
  path(s, {x:9.76, y:3.14, w:.09, h:.08, points:g51, fill:{color:SILVER}});
  path(s, {x:10.63, y:2.11, w:.12, h:.18, points:g52, fill:{color:SILVER}});
  path(s, {x:11.26, y:4.71, w:.63, h:.7, points:g53, fill:{color:SILVER}});
  path(s, {x:12.05, y:4.25, w:.09, h:.18, points:g54, fill:{color:SILVER}});
  path(s, {x:12.13, y:4.47, w:.07, h:.13, points:g55, fill:{color:SILVER}});
  path(s, {x:12.17, y:4.63, w:.05, h:.08, points:g56, fill:{color:SILVER}});
  path(s, {x:12.16, y:4.77, w:.09, h:.09, points:g57, fill:{color:SILVER}});
  path(s, {x:12.17, y:4.91, w:.08, h:.14, points:g58, fill:{color:SILVER}});
  path(s, {x:12.11, y:5.1, w:.12, h:.3, points:g59, fill:{color:SILVER}});
  path(s, {x:11.83, y:5.42, w:.36, h:.47, points:g60, fill:{color:SILVER}});
  txt(s, 'RUSSIA WORLD MAP', MEB44DarkC, {x:2.09, y:1.09, w:9.16, h:.84});
  txt(s, C8, OS12GreyL, {x:1.09, y:2.84, w:3.13, h:2.49});
  txt(s, 'Section Here', MEB16BlueL, {x:1.09, y:2.35, w:2.39, h:.37});
  s.addShape('rect', {x:1.09, y:5.98, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:1.24, y:6.04, w:1.43, h:.3});
  path(s, {x:10.28, y:3.22, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:10.41, y:3.35, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:5.33, y:4.66, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:5.46, y:4.79, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  path(s, {x:8.87, y:4.89, w:.43, h:.52, points:g22, fill:{color:BLUE}});
  path(s, {x:9, y:5.02, w:.16, h:.16, points:g23, fill:{color:BLUE}});
  cross(s, 4.01, 6.1);
  ring(s, 1.09, 1.2);
}

function slide73(s) {
  s.addShape('rect', {x:10.62, y:0, w:2.71, h:4.65, fill:{color:BLUE}});
  s.addShape('rect', {x:0, y:0, w:2.71, h:4.65, fill:{color:BLUE}});
  photo(s, 1.1, 2.9, 3.24, 3.51);
  photo(s, 5.05, 2.9, 3.24, 3.51);
  photo(s, 9, 2.9, 3.24, 3.51);
  txt(s, 'CREATIVE PORTFOLIO', MEB44DarkC, {x:3.96, y:1.09, w:5.41, h:1.58});
  s.addShape('rect', {x:1.1, y:5.68, w:3.24, h:.73, fill:{color:BLUE, transparency:15}});
  s.addShape('rect', {x:5.05, y:5.68, w:3.24, h:.73, fill:{color:BLUE, transparency:15}});
  s.addShape('rect', {x:9, y:5.68, w:3.24, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio One', MEB16WhiteC, {x:1.52, y:5.86, w:2.39, h:.37});
  txt(s, 'Portfolio Two', MEB16WhiteC, {x:5.47, y:5.86, w:2.39, h:.37});
  txt(s, 'Portfolio Three', MEB16WhiteC, {x:9.43, y:5.86, w:2.39, h:.37});
  s.addShape('ellipse', {x:1.09, y:1.2, w:.25, h:.25, fill:{type:'none'}, line:{color:WHITE, width:3}});
  cross(s, 11.93, 1.21);
}

function slide74(s) {
  s.addShape('rect', {x:10.81, y:3.05, w:2.53, h:2.42, fill:{color:BLUE}});
  s.addShape('rect', {x:0, y:4.25, w:2.53, h:3.25, fill:{color:BLUE}});
  photo(s, 1.09, 2.11, 2.99, 4.3);
  photo(s, 4.62, 2.11, 2.99, 1.88);
  photo(s, 4.62, 4.53, 2.99, 1.88);
  photo(s, 8.16, 4.53, 4.08, 1.88);
  photo(s, 8.16, 2.11, 4.08, 1.88);
  txt(s, 'CREATIVE PORTFOLIO', MEB44DarkC, {x:2.53, y:1.09, w:8.28, h:.84});
  s.addShape('rect', {x:1.1, y:5.68, w:2.99, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio One', MEB16WhiteC, {x:1.39, y:5.86, w:2.39, h:.37});
  s.addShape('rect', {x:4.63, y:3.26, w:2.99, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio Two', MEB16WhiteC, {x:4.93, y:3.44, w:2.39, h:.37});
  s.addShape('rect', {x:4.63, y:5.68, w:2.99, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio Four', MEB16WhiteC, {x:4.93, y:5.86, w:2.39, h:.37});
  s.addShape('rect', {x:8.16, y:3.26, w:4.08, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio Three', MEB16WhiteC, {x:9.01, y:3.44, w:2.39, h:.37});
  s.addShape('rect', {x:8.16, y:5.68, w:4.08, h:.73, fill:{color:BLUE, transparency:15}});
  txt(s, 'Portfolio Five', MEB16WhiteC, {x:9.01, y:5.86, w:2.39, h:.37});
  ring(s, 1.09, 1.2);
  cross(s, 11.86, 1.56);
}

function slide75(s) {
  s.addShape('rect', {x:5.48, y:0, w:6.76, h:3.41, fill:{color:BLUE}});
  photo(s, 6.57, 1.09, 2.84, 5.47);
  photo(s, 10.5, 1.09, 2.84, 5.47);
  txt(s, 'OUR\nBUSINESS\nPORTFOLIO', MEB44DarkL, {x:1.09, y:1.09, w:5.57, h:2.32});
  txt(s, C12, OS12GreyJ, {x:1.09, y:3.96, w:4.39, h:.98});
  txt(s, 'Portfolio One', MEB16BlueL, {x:1.09, y:3.64, w:2.39, h:.37});
  txt(s, C12, OS12GreyJ, {x:1.09, y:5.43, w:4.39, h:.98});
  txt(s, 'Portfolio Two', MEB16BlueL, {x:1.09, y:5.1, w:2.39, h:.37});
  s.addShape('rect', {x:6.57, y:5.72, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '01', MEB24WhiteC, {x:6.44, y:5.89, w:1.11, h:.5});
  s.addShape('rect', {x:10.5, y:5.72, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '02', MEB24WhiteC, {x:10.37, y:5.89, w:1.11, h:.5});
  ring(s, 5.85, 6.31);
  cross(s, 3.93, 1.14);
}

function slide76(s) {
  s.addShape('rect', {x:0, y:1.74, w:1.58, h:5.76, fill:{color:BLUE}});
  photo(s, 0, 0, 3.16, 3.48);
  photo(s, 3.7, 0, 3.16, 3.48);
  photo(s, 0, 4.02, 6.86, 2.38);
  txt(s, 'COMPANY\nPORTFOLIO', MEB44DarkL, {x:7.43, y:1.09, w:5.57, h:1.58});
  txt(s, C3, OS12GreyJ, {x:7.43, y:3.15, w:4.82, h:.67});
  txt(s, 'Portfolio One', MEB16BlueL, {x:7.43, y:2.89, w:2.39, h:.37});
  txt(s, C3, OS12GreyJ, {x:7.43, y:4.44, w:4.82, h:.67});
  txt(s, 'Portfolio Two', MEB16BlueL, {x:7.43, y:4.18, w:2.39, h:.37});
  txt(s, C3, OS12GreyJ, {x:7.43, y:5.73, w:4.82, h:.67});
  txt(s, 'Portfolio Three', MEB16BlueL, {x:7.43, y:5.47, w:2.39, h:.37});
  s.addShape('rect', {x:2.31, y:2.63, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '01', MEB24WhiteC, {x:2.18, y:2.8, w:1.11, h:.5});
  s.addShape('rect', {x:6.01, y:2.63, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '02', MEB24WhiteC, {x:5.89, y:2.8, w:1.11, h:.5});
  s.addShape('rect', {x:6.01, y:5.56, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '03', MEB24WhiteC, {x:5.89, y:5.73, w:1.11, h:.5});
  ring(s, 11.99, 2.5);
  cross(s, 11.9, 1.24);
}

function slide77(s) {
  s.addShape('rect', {x:0, y:0, w:2.18, h:5.48, fill:{color:BLUE}});
  photo(s, 1.09, 2.18, 1.95, 1.84);
  photo(s, 1.09, 4.57, 1.95, 1.84);
  photo(s, 3.56, 4.57, 3.1, 1.84);
  txt(s, 'CREATIVE PORTFOLIO', MEB44DarkC, {x:2.53, y:1.09, w:8.28, h:.84});
  s.addShape('rect', {x:3.56, y:2.18, w:3.1, h:1.84, fill:{color:BLUE}});
  txt(s, C3, OS12WhiteJ, {x:3.91, y:2.61, w:2.41, h:.98});
  txt(s, 'Portfolio One', MEB16BlueL, {x:7.76, y:2.18, w:2.39, h:.37});
  txt(s, C18, OS12GreyJ, {x:7.76, y:2.5, w:4.48, h:.98});
  txt(s, 'Portfolio Two', MEB16BlueL, {x:7.76, y:3.65, w:2.39, h:.37});
  txt(s, C18, OS12GreyJ, {x:7.76, y:3.96, w:4.48, h:.98});
  txt(s, 'Portfolio Three', MEB16BlueL, {x:7.76, y:5.11, w:2.39, h:.37});
  txt(s, C18, OS12GreyJ, {x:7.76, y:5.43, w:4.48, h:.98});
  ring(s, 7.09, 6.07);
  cross(s, 11.94, 1.14);
  s.addShape('rect', {x:2.42, y:3.39, w:.63, h:.63, fill:{color:BLUE}});
  txt(s, '01', MEBdWhiteC, {x:2.32, y:3.5, w:.82, h:.4});
  s.addShape('rect', {x:2.42, y:5.78, w:.63, h:.63, fill:{color:BLUE}});
  txt(s, '02', MEBdWhiteC, {x:2.32, y:5.89, w:.82, h:.4});
  s.addShape('rect', {x:6.04, y:5.78, w:.63, h:.63, fill:{color:BLUE}});
  txt(s, '03', MEBdWhiteC, {x:5.94, y:5.89, w:.82, h:.4});
}

function slide78(s) {
  photo(s, 1.09, 2.12, 2.55, 4.29);
  photo(s, 3.96, 2.12, 2.55, 4.29);
  txt(s, 'CREATIVE PORTFOLIO', MEB44DarkC, {x:2.53, y:1.09, w:8.28, h:.84});
  s.addShape('rect', {x:6.82, y:2.12, w:2.55, h:4.29, fill:{color:BLUE}});
  s.addShape('rect', {x:9.69, y:2.12, w:2.55, h:4.29, fill:{color:BLUE}});
  txt(s, 'Portfolio One', MEB16WhiteC, {x:6.91, y:3.55, w:2.39, h:.37});
  txt(s, C18, OS12WhiteC, {x:7.18, y:3.83, w:1.85, h:2.19});
  txt(s, 'Portfolio Two', MEB16WhiteC, {x:9.77, y:3.55, w:2.39, h:.37});
  txt(s, C18, OS12WhiteC, {x:10.04, y:3.83, w:1.85, h:2.19});
  s.addShape('rect', {x:7.68, y:2.41, w:.85, h:.85, fill:{color:WHITE}});
  txt(s, '01', MEB24BlueC, {x:7.54, y:2.58, w:1.11, h:.5});
  s.addShape('rect', {x:10.54, y:2.41, w:.85, h:.85, fill:{color:WHITE}});
  txt(s, '02', MEB24BlueC, {x:10.41, y:2.58, w:1.11, h:.5});
  s.addShape('rect', {x:1.09, y:5.56, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '01', MEB24WhiteC, {x:.96, y:5.73, w:1.11, h:.5});
  s.addShape('rect', {x:3.96, y:5.56, w:.85, h:.85, fill:{color:BLUE}});
  txt(s, '02', MEB24WhiteC, {x:3.83, y:5.73, w:1.11, h:.5});
  ring(s, 1.17, 1.16);
  cross(s, 11.94, 1.11);
}

function slide79(s) {
  photo(s, 6.67, 1.09, 5.58, 5.32);
  s.addShape('rect', {x:0, y:4.83, w:2.53, h:2.67, fill:{color:BLUE}});
  s.addShape('rect', {x:1.09, y:3.11, w:2.86, h:3.3, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'CLIENT’S\nWORDS', MEB44DarkL, {x:1.09, y:1.09, w:4.14, h:1.58});
  s.addShape('rect', {x:5.04, y:1.94, w:5.76, h:3.61, fill:{color:BLUE, transparency:15}});
  txt(s, '“Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi.”', OS12WhiteJ, {x:5.79, y:3.11, w:4.26, h:1.28});
  txt(s, C18, OS12GreyJ, {x:1.55, y:3.81, w:1.94, h:1.89});
  ring(s, 5.23, 1.15);
  cross(s, 5.95, 6.1);
}

function slide80(s) {
  txt(s, 'CLIENT’S WORDS', MEB44DarkC, {x:3.27, y:1.09, w:6.78, h:.84});
  s.addShape('rect', {x:1.09, y:3.26, w:5.03, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.09, y:5.06, w:5.03, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C8, OS12GreyC, {x:1.09, y:2.14, w:11.15, h:.67});
  s.addShape('rect', {x:7.21, y:3.26, w:5.03, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.21, y:5.06, w:5.03, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C13, OS12GreyJ, {x:2.8, y:3.45, w:2.96, h:.98});
  txt(s, C13, OS12GreyJ, {x:2.8, y:5.24, w:2.96, h:.98});
  txt(s, C13, OS12GreyJ, {x:8.92, y:3.45, w:2.96, h:.98});
  txt(s, C13, OS12GreyJ, {x:8.92, y:5.24, w:2.96, h:.98});
  ring(s, 11.63, 1.7);
  cross(s, 3.15, 4.89);
  photo(s, 1.09, 3.26, 1.35, 1.35);
  photo(s, 1.09, 5.06, 1.35, 1.35);
  photo(s, 7.21, 5.06, 1.35, 1.35);
  photo(s, 7.21, 3.26, 1.35, 1.35);
}

function slide81(s) {
  photo(s, 1.09, 2.02, 4.48, 4.39, 'light');
  s.addShape('rect', {x:1.09, y:2.02, w:4.48, h:4.39, fill:{color:BLUE, transparency:15}});
  s.addShape('ellipse', {x:2.41, y:1.09, w:1.85, h:1.85, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:2.51, y:1.19, w:1.65, h:1.65, points:g61, fill:{color:TONES.mid}});
  txt(s, '“Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.”', OS12WhiteC, {x:1.67, y:3.46, w:3.33, h:2.19});
  txt(s, 'CLIENT’S\nWORDS', MEB44DarkL, {x:6.67, y:1.09, w:3.7, h:1.58});
  txt(s, C26, OS12GreyJ, {x:6.67, y:3.16, w:5.56, h:2.49});
  ring(s, 11.98, 2.15);
  cross(s, 6.72, 6.11);
}

function slide82(s) {
  s.addShape('rect', {x:0, y:4.01, w:13.33, h:3.49, fill:{color:BLUE}});
  txt(s, 'TESTIMONIALS', MEB44DarkC, {x:3.26, y:1.09, w:6.81, h:.84});
  s.addShape('rect', {x:1.09, y:3.11, w:2.6, h:3.3, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:3.93, y:3.11, w:2.6, h:3.3, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:6.77, y:3.11, w:2.6, h:3.3, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:9.6, y:3.11, w:2.6, h:3.3, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C8, OS12GreyC, {x:1.09, y:2.14, w:11.15, h:.67});
  txt(s, C13, OS12GreyC, {x:1.34, y:4.89, w:2.1, h:1.28});
  txt(s, C13, OS12GreyC, {x:4.18, y:4.89, w:2.1, h:1.28});
  txt(s, C13, OS12GreyC, {x:7.02, y:4.89, w:2.1, h:1.28});
  txt(s, C13, OS12GreyC, {x:9.9, y:4.89, w:2.1, h:1.28});
  ring(s, 1.09, 1.09);
  cross(s, 11.66, 2.97);
  path(s, {x:1.72, y:3.34, w:1.34, h:1.34, points:g61, fill:{color:TONES.mid}});
  path(s, {x:4.57, y:3.34, w:1.34, h:1.34, points:g61, fill:{color:TONES.mid}});
  path(s, {x:7.42, y:3.34, w:1.34, h:1.34, points:g61, fill:{color:TONES.mid}});
  path(s, {x:10.28, y:3.34, w:1.34, h:1.34, points:g61, fill:{color:TONES.mid}});
}

function slide83(s) {
  s.addShape('rect', {x:0, y:3.75, w:13.33, h:3.75, fill:{color:BLUE}});
  s.addShape('ellipse', {x:5, y:2.09, w:3.32, h:3.32, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:5.27, y:2.36, w:1.3, h:1.3, points:g62, fill:{color:BLUE}});
  path(s, {x:5.27, y:3.85, w:1.3, h:1.3, points:g63, fill:{color:DARK}});
  path(s, {x:6.76, y:3.85, w:1.3, h:1.3, points:g64, fill:{color:BLUE}});
  path(s, {x:6.76, y:2.36, w:1.3, h:1.3, points:g65, fill:{color:DARK}});
  s.addShape('ellipse', {x:5, y:2.09, w:3.32, h:3.32, fill:{type:'none'}, line:{color:BLUE, width:6}});
  s.addShape('line', {x:4.1, y:1.58, w:1.39, h:1, line:{color:BLUE, width:3}, flipH:true, flipV:true});
  s.addShape('line', {x:7.84, y:1.58, w:1.22, h:.99, line:{color:BLUE, width:3}, flipH:true});
  s.addShape('line', {x:4.1, y:4.93, w:1.39, h:1.11, line:{color:WHITE, width:3}, flipH:true});
  s.addShape('line', {x:7.84, y:4.93, w:1.22, h:1.11, line:{color:WHITE, width:3}, flipH:true, flipV:true});
  txt(s, '01', MEB24WhiteC, {x:5.46, y:2.86, w:1.11, h:.5});
  txt(s, '02', MEB24WhiteC, {x:6.76, y:2.86, w:1.11, h:.5});
  txt(s, '03', MEB24WhiteC, {x:5.46, y:4.13, w:1.11, h:.5});
  txt(s, '04', MEB24WhiteC, {x:6.76, y:4.13, w:1.11, h:.5});
  s.addShape('rect', {x:8.6, y:1.09, w:3.64, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:5.05, w:3.64, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:8.6, y:5.05, w:3.64, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.1, y:1.09, w:3.64, h:1.35, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01', MEB24DarkC, {x:3.54, y:1.58, w:1.11, h:.5});
  s.addShape('line', {x:3.63, y:1.37, w:0, h:.83, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C7, OS12GreyR, {x:1.36, y:1.28, w:1.98, h:.98});
  txt(s, '03', MEB24DarkC, {x:3.54, y:5.53, w:1.11, h:.5});
  s.addShape('line', {x:3.63, y:5.33, w:0, h:.83, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C7, OS12GreyR, {x:1.36, y:5.24, w:1.98, h:.98});
  txt(s, '02', MEB24DarkC, {x:8.51, y:1.58, w:1.11, h:.5});
  s.addShape('line', {x:9.53, y:1.37, w:0, h:.83, line:{color:BLUE, width:3}, flipH:true, flipV:true});
  txt(s, C7, OS12GreyL, {x:9.83, y:1.28, w:1.98, h:.98});
  txt(s, '04', MEB24DarkC, {x:8.51, y:5.53, w:1.11, h:.5});
  s.addShape('line', {x:9.53, y:5.33, w:0, h:.83, line:{color:BLUE, width:3}, flipH:true, flipV:true});
  txt(s, C7, OS12GreyL, {x:9.83, y:5.24, w:1.98, h:.98});
  ring(s, 1.37, 2.31);
  cross(s, 11.95, 4.63);
}

function slide84(s) {
  s.addShape('line', {x:8.03, y:4.11, w:0, h:.39, line:{color:BLUE, width:3}, flipV:true});
  s.addShape('line', {x:10.78, y:4.5, w:0, h:.39, line:{color:BLUE, width:3}, flipV:true});
  s.addShape('line', {x:5.29, y:4.5, w:0, h:.39, line:{color:BLUE, width:3}, flipV:true});
  s.addShape('line', {x:2.55, y:4.11, w:0, h:.39, line:{color:BLUE, width:3}, flipV:true});
  s.addShape('rect', {x:1.8, y:2.59, w:1.49, h:1.49, fill:{color:BLUE}});
  txt(s, 'INFOGRAPHICS', MEB44DarkC, {x:3.26, y:1.09, w:6.81, h:.84});
  s.addShape('rect', {x:4.55, y:4.92, w:1.49, h:1.49, fill:{color:BLUE}});
  s.addShape('rect', {x:7.29, y:2.59, w:1.49, h:1.49, fill:{color:BLUE}});
  s.addShape('rect', {x:10.04, y:4.92, w:1.49, h:1.49, fill:{color:BLUE}});
  txt(s, C2, OS12GreyC, {x:1.31, y:5.43, w:2.48, h:.98});
  s.addShape('line', {x:1.09, y:4.5, w:11.15, h:0, line:{color:BLUE, width:3}});
  txt(s, C2, OS12GreyC, {x:4.05, y:3.1, w:2.48, h:.98});
  txt(s, C2, OS12GreyC, {x:9.54, y:3.1, w:2.48, h:.98});
  txt(s, C2, OS12GreyC, {x:6.8, y:5.43, w:2.48, h:.98});
  txt(s, '01', MEB24DarkC, {x:1.99, y:4.92, w:1.11, h:.5});
  txt(s, '03', MEB24DarkC, {x:7.48, y:4.92, w:1.11, h:.5});
  txt(s, '02', MEB24DarkC, {x:4.74, y:2.58, w:1.11, h:.5});
  txt(s, '04', MEB24DarkC, {x:10.23, y:2.58, w:1.11, h:.5});
  path(s, {x:10.45, y:5.32, w:.66, h:.66, points:g8, fill:{color:WHITE}});
  path(s, {x:2.18, y:3.05, w:.73, h:.56, points:g9, fill:{color:WHITE}});
  path(s, {x:2.39, y:3.18, w:.3, h:.3, points:g10, fill:{color:WHITE}});
  path(s, {x:4.93, y:5.33, w:.73, h:.67, points:g11, fill:{color:WHITE}});
  path(s, {x:7.67, y:2.97, w:.73, h:.73, points:g12, fill:{color:WHITE}});
  path(s, {x:7.91, y:3.09, w:.24, h:.24, points:g13, fill:{color:WHITE}});
  s.addShape('ellipse', {x:2.36, y:3.89, w:.38, h:.38, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:5.1, y:4.73, w:.38, h:.38, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:7.85, y:3.89, w:.38, h:.38, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:10.59, y:4.73, w:.38, h:.38, fill:{color:WHITE}, shadow:shadow()});
  ring(s, 2.85, 1.26);
  cross(s, 11.94, 5.62);
}

function slide85(s) {
  s.addShape('rect', {x:9.87, y:2.84, w:3.46, h:4.66, fill:{color:BLUE}});
  s.addShape('line', {x:3.3, y:2.84, w:3.88, h:0, line:{color:BLUE, width:3}});
  s.addShape('line', {x:3.3, y:3.81, w:3.88, h:0, line:{color:DARK, width:3}});
  s.addShape('line', {x:3.3, y:4.77, w:3.88, h:0, line:{color:BLUE, width:3}});
  s.addShape('line', {x:3.3, y:5.74, w:3.88, h:0, line:{color:DARK, width:3}});
  path(s, {x:1.09, y:5.33, w:4.51, h:.83, points:g66, fill:{color:DARK}});
  path(s, {x:1.61, y:4.42, w:3.46, h:.83, points:g67, fill:{color:BLUE}});
  path(s, {x:2.14, y:3.51, w:2.41, h:.83, points:g68, fill:{color:DARK}});
  path(s, {x:2.66, y:2.25, w:1.36, h:1.18, points:g69, fill:{color:BLUE}});
  txt(s, 'INFOGRAPHICS', MEB44DarkC, {x:3.26, y:1.09, w:6.81, h:.84});
  s.addShape('rect', {x:7.49, y:2.47, w:4.75, h:.75, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.49, y:3.43, w:4.75, h:.75, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.49, y:4.4, w:4.75, h:.75, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.49, y:5.36, w:4.75, h:.75, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01', MEB24DarkC, {x:7.49, y:2.59, w:1.11, h:.5});
  s.addShape('line', {x:8.6, y:2.59, w:0, h:.5, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C24, OS12GreyL, {x:8.87, y:2.5, w:3.11, h:.67});
  txt(s, '02', MEB24DarkC, {x:7.49, y:3.56, w:1.11, h:.5});
  s.addShape('line', {x:8.6, y:3.56, w:0, h:.5, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C24, OS12GreyL, {x:8.87, y:3.47, w:3.11, h:.67});
  txt(s, '03', MEB24DarkC, {x:7.49, y:4.52, w:1.11, h:.5});
  s.addShape('line', {x:8.6, y:4.52, w:0, h:.5, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C24, OS12GreyL, {x:8.87, y:4.43, w:3.11, h:.67});
  txt(s, '04', MEB24DarkC, {x:7.49, y:5.49, w:1.11, h:.5});
  s.addShape('line', {x:8.6, y:5.49, w:0, h:.5, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C24, OS12GreyL, {x:8.87, y:5.41, w:3.11, h:.67});
  ring(s, 10.7, 1.38);
  cross(s, 1.41, 2.29);
  path(s, {x:3.16, y:5.56, w:.37, h:.37, points:g8, fill:{color:WHITE}});
  path(s, {x:3.14, y:2.89, w:.4, h:.31, points:g9, fill:{color:WHITE}});
  path(s, {x:3.26, y:2.96, w:.17, h:.17, points:g10, fill:{color:WHITE}});
  path(s, {x:3.14, y:4.65, w:.4, h:.37, points:g11, fill:{color:WHITE}});
  path(s, {x:3.14, y:3.71, w:.4, h:.4, points:g12, fill:{color:WHITE}});
  path(s, {x:3.27, y:3.78, w:.13, h:.13, points:g13, fill:{color:WHITE}});
}

function slide86(s) {
  path(s, {x:1.09, y:3.65, w:2.29, h:2.77, points:g70, fill:{color:BLUE}});
  path(s, {x:4.04, y:3.65, w:2.29, h:2.77, points:g70, fill:{color:BLUE}});
  path(s, {x:7, y:3.65, w:2.29, h:2.77, points:g70, fill:{color:BLUE}});
  path(s, {x:9.95, y:3.65, w:2.29, h:2.77, points:g70, fill:{color:BLUE}});
  s.addShape('ellipse', {x:1.71, y:3.12, w:1.05, h:1.05, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:4.66, y:3.12, w:1.05, h:1.05, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:7.62, y:3.14, w:1.05, h:1.05, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:10.57, y:3.12, w:1.05, h:1.05, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'INFOGRAPHICS', MEB44DarkC, {x:3.26, y:1.09, w:6.81, h:.84});
  txt(s, C8, OS12GreyC, {x:1.09, y:2.14, w:11.15, h:.67});
  txt(s, '01', MEB24DarkC, {x:1.69, y:3.4, w:1.11, h:.5});
  txt(s, '02', MEB24DarkC, {x:4.64, y:3.4, w:1.11, h:.5});
  txt(s, '03', MEB24DarkC, {x:7.59, y:3.4, w:1.11, h:.5});
  txt(s, '04', MEB24DarkC, {x:10.54, y:3.4, w:1.11, h:.5});
  txt(s, C4, OS12WhiteC, {x:1.3, y:4.52, w:1.87, h:1.28});
  txt(s, C4, OS12WhiteC, {x:4.26, y:4.52, w:1.87, h:1.28});
  txt(s, C4, OS12WhiteC, {x:7.21, y:4.52, w:1.87, h:1.28});
  txt(s, C4, OS12WhiteC, {x:10.15, y:4.52, w:1.87, h:1.28});
  ring(s, 11.77, 1.65);
  cross(s, 1.14, 3.05);
}

function slide87(s) {
  path(s, {x:3.22, y:3.49, w:1.74, h:2.61, points:g71, fill:{color:BLUE}, rotate:315});
  path(s, {x:1.14, y:1.41, w:1.74, h:2.61, points:g72, fill:{color:BLUE}, rotate:315});
  path(s, {x:.71, y:3.92, w:2.61, h:1.74, points:g73, fill:{color:DARK}, rotate:315});
  path(s, {x:2.79, y:1.84, w:2.61, h:1.74, points:g74, fill:{color:DARK}, rotate:315});
  path(s, {x:3.72, y:4.41, w:.67, h:.67, points:g8, fill:{color:WHITE}});
  path(s, {x:1.68, y:2.51, w:.73, h:.57, points:g9, fill:{color:WHITE}});
  path(s, {x:1.89, y:2.64, w:.31, h:.31, points:g10, fill:{color:WHITE}});
  path(s, {x:1.68, y:4.41, w:.73, h:.67, points:g11, fill:{color:WHITE}});
  path(s, {x:3.69, y:2.42, w:.73, h:.73, points:g12, fill:{color:WHITE}});
  path(s, {x:3.93, y:2.54, w:.24, h:.24, points:g13, fill:{color:WHITE}});
  txt(s, 'INFOGRAPHICS', MEB44DarkL, {x:6.12, y:1.09, w:6.81, h:.84});
  txt(s, C12, OS12GreyJ, {x:6.12, y:3.48, w:6.12, h:.67});
  txt(s, '02. Section Here', MEB16BlueL, {x:6.11, y:3.11, w:2.21, h:.37});
  txt(s, C12, OS12GreyJ, {x:6.12, y:4.61, w:6.12, h:.67});
  txt(s, '03. Section Here', MEB16BlueL, {x:6.11, y:4.24, w:2.21, h:.37});
  txt(s, C12, OS12GreyJ, {x:6.12, y:5.73, w:6.12, h:.67});
  txt(s, '04. Section Here', MEB16BlueL, {x:6.11, y:5.36, w:2.21, h:.37});
  txt(s, C12, OS12GreyJ, {x:6.12, y:2.35, w:6.12, h:.67});
  txt(s, '01. Section Here', MEB16BlueL, {x:6.11, y:1.98, w:2.21, h:.37});
  ring(s, 5.38, 6.15);
  cross(s, 1.38, 1.16);
}

function slide88(s) {
  path(s, {x:7.94, y:1.09, w:2.26, h:1.96, points:g75, fill:{color:BLUE}});
  path(s, {x:7.94, y:3.34, w:2.26, h:1.96, points:g76, fill:{color:DARK}});
  path(s, {x:9.98, y:2.2, w:2.26, h:1.96, points:g76, fill:{color:DARK}});
  path(s, {x:9.98, y:4.45, w:2.26, h:1.96, points:g76, fill:{color:BLUE}});
  txt(s, 'INFOGRAPHICS', MEB44DarkL, {x:1.09, y:1.09, w:6.81, h:.84});
  path(s, {x:10.77, y:5.09, w:.67, h:.67, points:g8, fill:{color:WHITE}});
  path(s, {x:8.7, y:1.79, w:.73, h:.57, points:g9, fill:{color:WHITE}});
  path(s, {x:8.91, y:1.92, w:.31, h:.31, points:g10, fill:{color:WHITE}});
  path(s, {x:8.7, y:3.98, w:.73, h:.67, points:g11, fill:{color:WHITE}});
  path(s, {x:10.74, y:2.82, w:.73, h:.73, points:g12, fill:{color:WHITE}});
  path(s, {x:10.98, y:2.94, w:.24, h:.24, points:g13, fill:{color:WHITE}});
  s.addShape('rect', {x:1.09, y:2.2, w:2.73, h:1.96, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.09, y:4.45, w:2.73, h:1.96, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.11, y:2.2, w:2.73, h:1.96, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.11, y:4.45, w:2.73, h:1.96, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01. Section Here', MEB16BlueL, {x:1.35, y:2.51, w:2.21, h:.37});
  txt(s, C2, OS12GreyC, {x:1.35, y:2.88, w:2.22, h:.98});
  txt(s, '02. Section Here', MEB16BlueL, {x:4.37, y:2.51, w:2.21, h:.37});
  txt(s, C2, OS12GreyC, {x:4.36, y:2.88, w:2.22, h:.98});
  txt(s, '03. Section Here', MEB16BlueL, {x:1.35, y:4.75, w:2.21, h:.37});
  txt(s, C2, OS12GreyC, {x:1.35, y:5.13, w:2.22, h:.98});
  txt(s, C2, OS12GreyC, {x:4.36, y:5.13, w:2.22, h:.98});
  ring(s, 7.38, 6.15);
  cross(s, 1.4, 2.1);
  txt(s, '04. Section Here', MEB16BlueL, {x:4.37, y:4.75, w:2.21, h:.37});
}

function slide89(s) {
  photo(s, 1.1, 2.89, 5.27, 3.52);
  txt(s, 'AWESOME\nCHART', MEB44DarkL, {x:1.09, y:1.09, w:4.14, h:1.58});
  s.addShape('rect', {x:7.46, y:1.09, w:4.78, h:1.3, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.46, y:5.11, w:4.78, h:1.3, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.46, y:3.1, w:4.78, h:1.3, fill:{color:WHITE}, shadow:shadow()});
  txt(s, '01', MEB24DarkC, {x:7.51, y:1.48, w:1.11, h:.5});
  s.addShape('line', {x:8.62, y:1.26, w:0, h:.96, line:{color:BLUE, width:3}, flipV:true});
  txt(s, '02', MEB24DarkC, {x:7.51, y:3.5, w:1.11, h:.5});
  s.addShape('line', {x:8.62, y:3.27, w:0, h:.96, line:{color:BLUE, width:3}, flipV:true});
  txt(s, '03', MEB24DarkC, {x:7.51, y:5.51, w:1.11, h:.5});
  s.addShape('line', {x:8.62, y:5.28, w:0, h:.96, line:{color:BLUE, width:3}, flipV:true});
  txt(s, C10, OS12GreyJ, {x:8.85, y:1.25, w:3.17, h:.98});
  txt(s, C10, OS12GreyJ, {x:8.85, y:3.26, w:3.17, h:.98});
  txt(s, C10, OS12GreyJ, {x:8.85, y:5.27, w:3.17, h:.98});
  ring(s, 6.67, 6.28);
  cross(s, 6.98, .97);
}

function slide90(s) {
  photo(s, 6.18, 2.37, 6.06, 4.04);
  photo(s, 1.09, 4.18, 4, 2.23);
  txt(s, 'AWESOME CHART', MEB44DarkC, {x:2.51, y:1.09, w:8.31, h:.84});
  txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore.', OS12GreyJ, {x:1.09, y:2.47, w:4, h:1.28});
  ring(s, 4.96, 5.87);
  cross(s, 10.87, 1.21);
}

function slide91(s) {
  photo(s, 6.67, 2.71, 5.57, 3.69);
  s.addShape('rect', {x:0, y:0, w:3.04, h:7.5, fill:{color:BLUE}});
  txt(s, 'AWESOME\nCHART', MEB44DarkL, {x:6.67, y:1.09, w:4.14, h:1.58});
  s.addShape('rect', {x:1.1, y:1.1, w:3.99, h:2.5, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.62, y:1.87, w:.96, h:.96, fill:{color:BLUE}});
  txt(s, '01', MEB28WhiteC, {x:4.76, y:2.04, w:.68, h:.57});
  s.addShape('rect', {x:1.1, y:3.9, w:3.99, h:2.5, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.62, y:4.67, w:.96, h:.96, fill:{color:BLUE}});
  txt(s, '02', MEB28WhiteC, {x:4.68, y:4.84, w:.82, h:.57});
  txt(s, C38, OS12GreyR, {x:1.54, y:1.71, w:2.59, h:1.28});
  txt(s, C38, OS12GreyR, {x:1.54, y:4.49, w:2.59, h:1.28});
  ring(s, 6.18, 6.15);
  cross(s, 11.93, 1.21);
}

function slide92(s) {
  photo(s, 1.09, 3.12, 4.93, 3.28);
  txt(s, 'AWESOME CHART', MEB44DarkC, {x:2.51, y:1.09, w:8.31, h:.84});
  txt(s, C8, OS12GreyC, {x:1.09, y:2.14, w:11.15, h:.67});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident.', OS12GreyJ, {x:7.11, y:3.12, w:5.12, h:1.28});
  txt(s, C1, OS12GreyJ, {x:7.44, y:5.73, w:4.8, h:.67});
  s.addShape('ellipse', {x:7.11, y:5.88, w:.15, h:.15, fill:{color:BLUE}});
  txt(s, C1, OS12GreyJ, {x:7.44, y:4.71, w:4.8, h:.67});
  s.addShape('ellipse', {x:7.11, y:4.86, w:.15, h:.15, fill:{color:BLUE}});
  ring(s, 6.67, 6.15);
  cross(s, 1.14, 1.71);
}

function slide93(s) {
  photo(s, 6.33, 2.47, 5.91, 3.99);
  txt(s, 'AWESOME\nCHART', MEB44DarkL, {x:1.09, y:1.09, w:4.14, h:1.58});
  areaChart(s, 6.33,2.47,5.92,3.94);
  s.addShape('rect', {x:8.74, y:-1.32, w:1.09, h:5.92, fill:{color:BLUE}, rotate:90});
  txt(s, C10, OS12WhiteJ, {x:6.61, y:1.3, w:5.35, h:.67});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex.', OS12GreyJ, {x:1.1, y:2.85, w:4.13, h:1.58});
  s.addShape('rect', {x:0, y:4.65, w:5.23, h:1.81, fill:{color:BLUE}});
  txt(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et.', OS12WhiteJ, {x:1.09, y:4.91, w:3.48, h:1.28});
  ring(s, 11.98, 2.85);
  cross(s, 5.74, .97);
}

function slide94(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:2.91, y:-2.92, w:7.5, h:13.34, fill:{color:BLUE, transparency:15}, rotate:90});
  txt(s, 'PRICING TABLE', MEB44WhiteC, {x:2.51, y:1.09, w:8.31, h:.84});
  s.addShape('rect', {x:1.09, y:2.4, w:2.56, h:4, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:3.95, y:2.4, w:2.56, h:4, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:6.81, y:2.4, w:2.56, h:4, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:9.67, y:2.4, w:2.56, h:4, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'Feature One', OSSB12GreyC, {x:1.35, y:3.97, w:2.05, h:.37});
  txt(s, 'Feature Two', OSSB12GreyC, {x:1.35, y:4.37, w:2.05, h:.37});
  txt(s, 'Feature Three', OSSB12GreyC, {x:1.35, y:4.77, w:2.05, h:.37});
  txt(s, 'Feature Four', OSSB12GreyC, {x:1.35, y:5.17, w:2.05, h:.37});
  txt(s, 'Feature Five', OSSB12GreyC, {x:1.35, y:5.57, w:2.05, h:.37});
  txt(s, 'Feature One', OSSB12GreyC, {x:4.21, y:3.97, w:2.05, h:.37});
  txt(s, 'Feature Two', OSSB12GreyC, {x:4.21, y:4.37, w:2.05, h:.37});
  txt(s, 'Feature Three', OSSB12GreyC, {x:4.21, y:4.77, w:2.05, h:.37});
  txt(s, 'Feature Four', OSSB12GreyC, {x:4.21, y:5.17, w:2.05, h:.37});
  txt(s, 'Feature Five', OSSB12GreyC, {x:4.21, y:5.57, w:2.05, h:.37});
  txt(s, 'Feature One', OSSB12GreyC, {x:7.07, y:3.97, w:2.05, h:.37});
  txt(s, 'Feature Two', OSSB12GreyC, {x:7.07, y:4.37, w:2.05, h:.37});
  txt(s, 'Feature Three', OSSB12GreyC, {x:7.07, y:4.77, w:2.05, h:.37});
  txt(s, 'Feature Four', OSSB12GreyC, {x:7.07, y:5.17, w:2.05, h:.37});
  txt(s, 'Feature Five', OSSB12GreyC, {x:7.07, y:5.57, w:2.05, h:.37});
  txt(s, 'Feature One', OSSB12GreyC, {x:9.93, y:3.97, w:2.05, h:.37});
  txt(s, 'Feature Two', OSSB12GreyC, {x:9.93, y:4.37, w:2.05, h:.37});
  txt(s, 'Feature Three', OSSB12GreyC, {x:9.93, y:4.77, w:2.05, h:.37});
  txt(s, 'Feature Four', OSSB12GreyC, {x:9.93, y:5.17, w:2.05, h:.37});
  txt(s, 'Feature Five', OSSB12GreyC, {x:9.93, y:5.57, w:2.05, h:.37});
  s.addShape('rect', {x:1.09, y:2.87, w:2.56, h:.96, fill:{color:BLUE}});
  txt(s, '25%', MEB28WhiteC, {x:1.63, y:3.05, w:1.49, h:.57});
  s.addShape('rect', {x:3.95, y:2.87, w:2.56, h:.96, fill:{color:BLUE}});
  txt(s, '35%', MEB28WhiteC, {x:4.49, y:3.05, w:1.49, h:.57});
  s.addShape('rect', {x:6.81, y:2.87, w:2.56, h:.96, fill:{color:BLUE}});
  txt(s, '40%', MEB28WhiteC, {x:7.35, y:3.05, w:1.49, h:.57});
  s.addShape('rect', {x:9.67, y:2.87, w:2.56, h:.96, fill:{color:BLUE}});
  txt(s, '60%', MEB28WhiteC, {x:10.21, y:3.05, w:1.49, h:.57});
  s.addShape('ellipse', {x:11.95, y:1.79, w:.25, h:.25, fill:{type:'none'}, line:{color:WHITE, width:3}});
  cross(s, 1.14, 1.36);
}

function slide95(s) {
  photo(s, 1.09, 1.09, 11.16, 5.33);
  s.addShape('rect', {x:1.09, y:1.09, w:11.16, h:5.33, fill:{color:BLUE, transparency:15}});
  txt(s, '“Far and away the best prize that life offers is the chance to work hard at work worth doing.”', MWdWhiteC, {x:2.17, y:3.27, w:8.99, h:.96});
  s.addShape('ellipse', {x:10.91, y:5.2, w:.25, h:.25, fill:{type:'none'}, line:{color:WHITE, width:3}});
  cross(s, 2.22, 1.96);
}

function slide96(s) {
  s.addShape('rect', {x:0, y:4.71, w:4.06, h:2.79, fill:{color:BLUE}});
  s.addShape('rect', {x:4.97, y:3.5, w:2.16, h:2.9, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:4.97, y:3.5, w:2.16, h:.52, fill:{color:BLUE}});
  txt(s, C6, OS10GreyC, {x:5.23, y:5.07, w:1.62, h:.83});
  txt(s, 'YOUR NAME', MEB10WhiteC, {x:5.24, y:3.57, w:1.6, h:.33});
  s.addShape('rect', {x:1.09, y:3.5, w:2.16, h:2.9, fill:{color:WHITE}, shadow:shadow()});
  txt(s, C6, OS10GreyC, {x:1.35, y:5.07, w:1.62, h:.83});
  s.addShape('rect', {x:1.09, y:3.5, w:2.16, h:.52, fill:{color:BLUE}});
  txt(s, 'YOUR NAME', MEB10WhiteC, {x:1.36, y:3.57, w:1.6, h:.33});
  txt(s, 'SOCIAL MEDIA', MEB44DarkC, {x:2.51, y:1.09, w:8.31, h:.84});
  path(s, {x:5.96, y:4.37, w:.19, h:.37, points:g77, fill:{color:BLUE}});
  path(s, {x:1.97, y:4.39, w:.4, h:.33, points:g78, fill:{color:BLUE}});
  txt(s, C39, OS12GreyJ, {x:8.22, y:2.5, w:4.01, h:2.8});
  s.addShape('rect', {x:10.53, y:5.99, w:1.72, h:.42, fill:{color:BLUE}});
  txt(s, 'LEARN MORE', MEB12WhiteC, {x:10.67, y:6.05, w:1.43, h:.3});
  s.addShape('rect', {x:2.73, y:2.7, w:2.76, h:3.71, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:3.79, y:3.67, w:.64, h:.82, points:g79, fill:{color:BLUE}});
  txt(s, C6, OS12GreyC, {x:3.07, y:4.8, w:2.08, h:.98});
  s.addShape('rect', {x:2.73, y:2.7, w:2.76, h:.67, fill:{color:BLUE}});
  txt(s, 'YOUR NAME', MEB12WhiteC, {x:3.08, y:2.85, w:2.05, h:.37});
  ring(s, 9.36, 6.15);
  cross(s, 1.14, 1.14);
}

function slide97(s) {
  s.addShape('rect', {x:9.96, y:0, w:3.37, h:7.5, fill:{color:BLUE}});
  s.addShape('line', {x:7.02, y:3.39, w:2.27, h:2.08, line:{color:BLUE, width:3}, flipH:true});
  s.addShape('line', {x:9.97, y:2.85, w:1.49, h:2.76, line:{color:WHITE, width:3}, flipH:true, flipV:true});
  s.addShape('line', {x:7.19, y:2.08, w:2.06, h:.59, line:{color:BLUE, width:3}, flipH:true, flipV:true});
  s.addShape('ellipse', {x:10.66, y:4.82, w:1.59, h:1.59, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('ellipse', {x:5.97, y:1.09, w:1.59, h:1.59, fill:{color:WHITE}, shadow:shadow()});
  txt(s, 'SOCIAL\nMEDIA', MEB44DarkL, {x:1.09, y:1.09, w:3.21, h:1.58});
  txt(s, C14, OS12GreyJ, {x:1.09, y:2.99, w:3.8, h:3.1});
  path(s, {x:11.22, y:5.14, w:.48, h:.95, points:g77, fill:{color:BLUE}});
  path(s, {x:6.44, y:1.47, w:.64, h:.82, points:g79, fill:{color:BLUE}});
  s.addShape('ellipse', {x:8.62, y:1.34, w:2.7, h:2.7, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:9.29, y:2.13, w:1.35, h:1.11, points:g78, fill:{color:BLUE}});
  s.addShape('ellipse', {x:5.98, y:4.32, w:2.09, h:2.09, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:6.83, y:5.07, w:.71, h:.73, points:g80, fill:{color:BLUE}});
  path(s, {x:6.46, y:5.08, w:.23, h:.72, points:g81, fill:{color:BLUE}});
  path(s, {x:6.45, y:4.72, w:.25, h:.24, points:g82, fill:{color:BLUE}});
  ring(s, 9.06, 6.15);
  cross(s, 6.02, 3.46);
}

function slide98(s) {
  s.addShape('rect', {x:1.09, y:3, w:5.04, h:1.31, fill:{color:BLUE}});
  s.addShape('rect', {x:1.09, y:5.1, w:5.04, h:1.31, fill:{color:BLUE}});
  s.addShape('rect', {x:7.21, y:3, w:5.04, h:1.31, fill:{color:BLUE}});
  s.addShape('rect', {x:7.21, y:5.1, w:5.04, h:1.31, fill:{color:BLUE}});
  txt(s, 'SOCIAL\nMEDIA', MEB44DarkL, {x:1.09, y:1.09, w:6.01, h:1.58});
  s.addShape('rect', {x:1.09, y:3, w:1.31, h:1.31, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:1.09, y:5.1, w:1.31, h:1.31, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.21, y:3, w:1.31, h:1.31, fill:{color:WHITE}, shadow:shadow()});
  s.addShape('rect', {x:7.21, y:5.1, w:1.31, h:1.31, fill:{color:WHITE}, shadow:shadow()});
  path(s, {x:7.69, y:5.4, w:.35, h:.7, points:g77, fill:{color:BLUE}});
  path(s, {x:7.6, y:3.31, w:.54, h:.69, points:g79, fill:{color:BLUE}});
  path(s, {x:1.44, y:3.42, w:.57, h:.47, points:g78, fill:{color:BLUE}});
  path(s, {x:1.64, y:5.66, w:.37, h:.38, points:g80, fill:{color:BLUE}});
  path(s, {x:1.44, y:5.66, w:.12, h:.38, points:g81, fill:{color:BLUE}});
  path(s, {x:1.44, y:5.47, w:.13, h:.13, points:g82, fill:{color:BLUE}});
  txt(s, 'Your Account', OSSB12WhiteL, {x:2.63, y:3.17, w:1.75, h:.37});
  txt(s, C7, OS12WhiteJ, {x:2.64, y:3.47, w:2.94, h:.67});
  txt(s, 'Your Account', OSSB12WhiteL, {x:2.63, y:5.27, w:1.75, h:.37});
  txt(s, C7, OS12WhiteJ, {x:2.64, y:5.57, w:2.94, h:.67});
  txt(s, 'Your Account', OSSB12WhiteL, {x:8.75, y:3.17, w:1.75, h:.37});
  txt(s, C7, OS12WhiteJ, {x:8.76, y:3.47, w:2.94, h:.67});
  txt(s, 'Your Account', OSSB12WhiteL, {x:8.75, y:5.27, w:1.75, h:.37});
  txt(s, C7, OS12WhiteJ, {x:8.76, y:5.57, w:2.94, h:.67});
  txt(s, C39, OS12GreyJ, {x:4.39, y:1.26, w:7.86, h:1.28});
  ring(s, 11.99, 4.62);
  cross(s, 3.98, 2.23);
}

function slide99(s) {
  photo(s, 1.09, 2.64, 11.16, 3.78);
  txt(s, 'CONTACT US', MEB44DarkL, {x:1.09, y:1.09, w:6.01, h:.84});
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', RB12GreyJ, {x:8.64, y:1.1, w:3.6, h:.98});
  s.addShape('rect', {x:0, y:4.54, w:13.33, h:2.96, fill:{color:BLUE, transparency:15}});
  txt(s, 'Address', MEBdWhiteL, {x:2.77, y:4.87, w:1.8, h:.4});
  txt(s, 'St. Pangeran Diponegoro 12', OS12WhiteJ, {x:2.77, y:5.27, w:2.46, h:.37});
  txt(s, 'Social Media', MEBdWhiteL, {x:8.26, y:4.87, w:2.07, h:.4});
  txt(s, 'Phone', MEBdWhiteL, {x:6.05, y:4.87, w:1.14, h:.4});
  txt(s, '+123 456 789', OS12WhiteJ, {x:6.05, y:5.27, w:1.4, h:.37});
  txt(s, 'Instagram', OS12WhiteJ, {x:8.26, y:5.27, w:1.4, h:.37});
  txt(s, 'Facebook', OS12WhiteJ, {x:8.26, y:5.65, w:1.4, h:.37});
  txt(s, 'LinkedIn', OS12WhiteJ, {x:9.65, y:5.27, w:.91, h:.37});
  txt(s, 'Twitter', OS12WhiteJ, {x:9.65, y:5.65, w:.91, h:.37});
  ring(s, 7.74, 1.56);
  cross(s, 1.14, 2.15);
}

function slide100(s) {
  photo(s, 0, 0, 13.33, 7.5);
  s.addShape('rect', {x:0, y:0, w:13.33, h:7.5, fill:{color:DARK, transparency:15}});
  txt(s, 'THANKS', MEB96BlueC, {x:2.8, y:2.89, w:7.73, h:1.72});
  cross(s, .94, 6.25);
  cross(s, 11.11, 1.04, 0);
  cross(s, 1.78, 3.08, 0);
  cross(s, 10.11, 4.59);
  ring(s, 11.68, 6.15);
  ring(s, 3.23, .79);
}

// ---------------------------------------------------------------- output
const BUILDERS = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
  slide41, slide42, slide43, slide44, slide45, slide46, slide47, slide48, slide49, slide50,
  slide51, slide52, slide53, slide54, slide55, slide56, slide57, slide58, slide59, slide60,
  slide61, slide62, slide63, slide64, slide65, slide66, slide67, slide68, slide69, slide70,
  slide71, slide72, slide73, slide74, slide75, slide76, slide77, slide78, slide79, slide80,
  slide81, slide82, slide83, slide84, slide85, slide86, slide87, slide88, slide89, slide90,
  slide91, slide92, slide93, slide94, slide95, slide96, slide97, slide98, slide99, slide100,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
BUILDERS.forEach(build => build(pptx.addSlide()));
pptx.writeFile({ fileName: nodePath.join(__dirname, '026f1a5a-be2b-472b-8233-761d8fff077c_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
