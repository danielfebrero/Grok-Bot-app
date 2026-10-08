/*
 * "House Real Estate Presentation" - a 50 slide template rebuilt with pptxgenjs.
 *
 *   node 0cc535a5-421c-443f-969b-10ddf835933e_grok_final.js   ->   0cc535a5-421c-443f-969b-10ddf835933e_grok_final.pptx
 *
 * Photographs from the source deck are stand-ins here: see photo().
 * Everything else - panels, icons, rules, charts - is a native shape.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ------------------------------------------------------------------ palette
const BG     = 'E3E0DD';   // canvas / slide background
const CARD   = 'F6F5F4';   // accent 3 - near-white panel
const CHAR   = '3E3D3D';
const DEEP   = '626E76';
const GREY   = '595959';   // body copy
const INK    = '000000';
const MIST   = 'B4BEC4';
const RULE   = 'D0CFCF';   // hairlines
const RULE2  = 'D9D9D9';
const RULE3  = 'DEDEDE';
const SAND   = 'E0DFDB';   // accent 1
const SAND2  = 'ECEBE9';   // accent 2
const STEEL  = '82939D';   // accent 4 - primary steel blue
const STEEL2 = '8D9CA5';   // accent 5
const STEEL3 = '98A6AE';   // accent 6
const TAUPE  = '7B7A7A';
const WHITE  = 'FFFFFF';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const HEAD = 'Raleway Bold';   // theme major font
const BODY = 'Roboto';         // theme minor font
const BR = '\n';
const NOLINE = { type: 'none' };
const LIFT = { type: 'outer', color: INK, opacity: 0.05, blur: 15 };  // soft card lift

// Filler copy used all over the template; lorem(n) returns its first n characters.
const LOREM =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunci ' +
  'PLACEHOLDER';
const lorem = (n) => (n ? LOREM.slice(0, n) : LOREM);

// Longer filler variants that are not simple prefixes of LOREM.
const LOREM1 =
  'Aliquet nibh biba praesent tristique magna sitami nibh biba pre';
const LOREM10 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'PLACEHOLDER' +
  'magna sitaminas sit amet. Risus pretium quam';
const LOREM11 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunci ' +
  'PLACEHOLDER' +
  'biba present praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium ';
const LOREM12 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'PLACEHOLDER' +
  'magna sitaminas sit amet. Risus pretium quam tristique magna sitaminas ' +
  'sit ';
const LOREM2 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'nibh biba praesent';
const LOREM3 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'nibh biba praesent tristique ';
const LOREM4 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'praesent tristique';
const LOREM5 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'PLACEHOLDER';
const LOREM6 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunci ' +
  'PLACEHOLDER' +
  'biba praesent';
const LOREM7 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'PLACEHOLDER' +
  'sit amet. Risus ';
const LOREM8 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunci ' +
  'PLACEHOLDER' +
  'biba praesent tristique magna ';
const LOREM9 =
  'Aliquet nibh biba praesent tristique magna sitaminas sit amet. Risus ' +
  'pretium quam vulputate dignissim. Enim silka milal necis badui nunc ' +
  'PLACEHOLDER' +
  'sit amet. Risus praesent tristique';

// Repeated captions
const CEOF     = 'CEO Founder';
const EXEC     = 'Executive Manager';
const GALLERY  = 'Gallery';
const GM       = 'General Manager';
const HARRY    = 'Harry Gilbert';
const JONAH    = 'Jonah Britany';
const MOCKUP   = 'Mockup Device Home Real Estate Here';
const SKILL    = 'Skill Here';
const SUBHEAD  = 'Infographic Powerpoint Template';
const TITLE    = 'Title Here';
const TITTLE   = 'Tittle Here';
const VISION   = 'Vision Carey';

// ------------------------------------------------------------------ text styles
// Naming: h/p = heading (Raleway) or body (Roboto), then point size, then a
// colour hint (W white, K black, C card, S steel) and alignment (c/r/j).
const S = {
  p11:      { fontFace: BODY, fontSize: 11, color: GREY, lineSpacingMultiple: 1.5 },
  h32K:     { fontFace: HEAD, fontSize: 32, color: INK },
  h18K:     { fontFace: HEAD, fontSize: 18, color: INK },
  p11W:     { fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 },
  p11c:     { fontFace: BODY, fontSize: 11, color: GREY, align: 'center', lineSpacingMultiple: 1.5 },
  h18Kc:    { fontFace: HEAD, fontSize: 18, color: INK, align: 'center' },
  p11r:     { fontFace: BODY, fontSize: 11, color: GREY, align: 'right', lineSpacingMultiple: 1.5 },
  h18Kr:    { fontFace: HEAD, fontSize: 18, color: INK, align: 'right' },
  h16W:     { fontFace: HEAD, fontSize: 16, color: WHITE, lineSpacingMultiple: 1.5 },
  h32K_2:   { fontFace: HEAD, fontSize: 32, color: INK, bold: true, lineSpacing: 47.12 },
  p16:      { fontFace: BODY, fontSize: 16, color: GREY, lineSpacing: 21.05 },
  h14K:     { fontFace: HEAD, fontSize: 14, color: INK, lineSpacingMultiple: 1.5 },
  h9Kc:     { fontFace: HEAD, fontSize: 9, color: INK, align: 'center', lineSpacingMultiple: 1.5 },
  p11j:     { fontFace: BODY, fontSize: 11, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 },
  h32K_3:   { fontFace: HEAD, fontSize: 32, color: INK, transparency: 95, bold: true },
  h32W:     { fontFace: HEAD, fontSize: 32, color: WHITE, bold: true },
  p11W_2:   { fontFace: BODY, fontSize: 11, color: WHITE, italic: true, lineSpacingMultiple: 1.5 },
  p12W:     { fontFace: BODY, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 },
  h16K:     { fontFace: HEAD, fontSize: 16, color: INK },
  h16W_2:   { fontFace: HEAD, fontSize: 16, color: WHITE },
  h18W:     { fontFace: HEAD, fontSize: 18, color: WHITE },
  p11K:     { fontFace: BODY, fontSize: 11, color: INK, lineSpacingMultiple: 1.5 },
  p9:       { fontFace: BODY, fontSize: 9, color: GREY, lineSpacingMultiple: 1.5 },
  h32S2c:   { fontFace: HEAD, fontSize: 32, color: STEEL2, bold: true, align: 'center' },
  h32Sc:    { fontFace: HEAD, fontSize: 32, color: STEEL, bold: true, align: 'center' },
  h32Kc:    { fontFace: HEAD, fontSize: 32, color: INK, bold: true, align: 'center', margin: [14.256,14.256,7.128,7.128] },
  h14W:     { fontFace: HEAD, fontSize: 14, color: WHITE, lineSpacingMultiple: 1.5 },
  h80Kc:    { fontFace: HEAD, fontSize: 80, color: INK, align: 'center' },
  h80K:     { fontFace: HEAD, fontSize: 80, color: INK },
  p1704Sc:  { fontFace: 'Poppins', fontSize: 17.04, color: STEEL, bold: true, align: 'center', valign: 'middle', margin: [0,0,0,0] },
  h18K_2:   { fontFace: HEAD, fontSize: 18, color: INK, lineSpacingMultiple: 1.5 },
  p11_2:    { fontFace: BODY, fontSize: 11, color: GREY, italic: true, lineSpacingMultiple: 1.5 },
  h16K_2:   { fontFace: HEAD, fontSize: 16, color: INK, lineSpacingMultiple: 1.5 },
  p1704Cc:  { fontFace: 'Poppins', fontSize: 17.04, color: CARD, bold: true, align: 'center', valign: 'middle', margin: [0,0,0,0] },
};

// ------------------------------------------------------------------ outlines
// Icons and other freeform art, stored as closed rings of x,y pairs in percent
// of the shape's own box. art() scales a ring set into a real frame.
const OUT = {
  g0: [[0,0,100,0,100,100,0,100]],
  g1: [[40,14,14,40,40,66,66,40,40,14], [42,0,73,17,77,57,100,80,80,100,2,53,9,15,42,0]],
  g2: [[36,0,100,0,100,46,17,47,62,54,17,62,61,76,17,83,94,86,100,100,0,100,0,30,36,30,36,0]],
  g3: [[100,0,100,100,0,100,100,0]],
  g4: [[50,8,16,24,50,41,84,24,50,8], [50,0,100,24,100,76,50,100,0,76,0,24,50,0]],
  g5: [[0,0,100,100,0,100,0,0]],
  g6: [[69,70,69,89,87,89,87,70,69,70], [41,58,41,89,59,89,59,58,41,58], [13,45,13,89,31,89,31,45,13,45], [0,0,73,0,73,23,100,23,100,100,0,100,0,0]],
  g7: [[50,23,36,49,57,68,34,71,50,88,66,65,45,47,64,40,50,23], [50,0,99,43,85,100,11,95,0,46,11,10,29,17,50,0]],
  g8: [[18,55,18,76,41,76,41,55,18,55], [50,0,100,48,100,98,82,100,82,55,57,55,57,100,2,100,0,48,50,0]],
  g9: [[50,0,96,100,50,26,1,96,20,17,27,40,50,0]],
  g10: [[54,0,92,21,100,54,79,92,46,100,15,85,0,46,15,15,54,0]],
  g11: [[46,0,85,14,100,46,92,78,54,100,22,91,0,54,9,21,46,0]],
  g12: [[65,48,54,59,65,71,76,59,65,48], [65,30,34,70,65,30], [35,30,24,41,35,52,46,41,35,30], [50,0,93,24,95,71,64,98,3,99,5,29,50,0]],
  g13: [[25,0,100,100,7,100,1,51,25,0]],
  g14: [[32,0,73,36,100,100,0,100,32,0]],
  g15: [[89,0,100,1,39,100,0,38,89,0]],
  g16: [[0,0,100,0,71,100,29,65,0,0]],
  g17: [[0,0,93,0,99,49,74,100,0,0]],
  g18: [[61,0,100,63,54,93,0,99,61,0]],
  g19: [[8,0,92,0,99,2,100,4,100,96,99,98,92,100,8,100,1,98,0,96,0,4,2,1,8,0]],
  g20: [[50,0,99,85,96,100,50,26,8,98,1,96,19,55,20,17,27,40,50,0]],
  g21: [[11,0,89,0,95,1,100,4,100,95,99,98,95,99,89,100,11,100,5,99,0,96,0,5,3,1,11,0]],
  g22: [[10,0,90,0,95,1,99,3,100,6,100,100,0,100,0,6,3,2,10,0]],
  g23: [[9,0,91,0,99,2,100,4,100,96,99,98,91,100,9,100,1,98,0,96,0,4,3,1,9,0]],
  g24: [[100,4,100,85,50,100,0,85,0,4,50,0,100,4]],
  g25: [[3,1,0,0,0,100,43,91,72,80,92,67,100,53,96,39,80,24,48,11,3,1]],
  g26: [[100,0,57,9,28,20,8,33,0,47,4,61,29,80,66,93,100,100]],
  g27: [[100,50,98,63,85,85,63,98,50,100,37,98,15,85,2,63,0,50,2,37,15,15,37,2,50,0,63,2,85,15,98,37,100,50]],
  g28: [[39,46,39,0,100,30,100,100,0,100,0,30,39,46], [50,17,50,61,11,45,11,90,89,90,89,36,50,17]],
  g29: [[91,90,100,100,0,100,10,1,90,1,91,90], [82,90,82,10,18,10,18,90,82,90], [32,45,45,55,32,45], [32,25,45,35,32,25], [32,65,45,75,32,65], [55,65,68,75,55,65], [55,45,68,55,55,45], [55,25,68,35,55,25]],
  g30: [[91,89,100,100,0,100,11,1,62,2,64,89,82,89,73,33,91,37,91,89], [18,11,18,89,55,89,55,11,18,11], [27,44,45,56,27,44], [27,22,45,33,27,22]],
  g31: [[95,100,1,98,0,51,22,1,99,2,95,100], [35,89,50,89,50,55,30,36,10,55,10,89,25,89,25,67,35,89], [60,89,90,89,90,11,30,11,60,50,60,89], [70,44,80,56,70,44], [70,67,80,78,70,67], [70,22,80,33,70,22], [50,22,60,33,50,22]],
  g32: [[86,17,100,17,100,100,0,100,0,17,15,2,83,0,86,17], [86,28,14,28,14,89,27,89,27,50,36,89,45,50,55,89,64,50,73,89,86,89,86,28], [23,11,77,17,23,11]],
  g33: [[100,100,0,100,6,2,84,2,100,100], [75,89,85,44,55,44,55,89,65,56,75,89], [75,33,75,11,15,11,15,89,45,89,45,33,75,33], [25,44,35,56,25,44], [25,67,35,78,25,67], [25,22,35,33,25,22]],
  g34: [[0,0,100,0,100,100,0,0]],
  g35: [[100,0,0,100,0,0,100,0]],
  g36: [[81,0,100,50,81,100,0,100,37,0,81,0]],
  g37: [[19,0,0,50,19,100,100,100,63,0,19,0]],
  g38: [[50,11,11,41,34,58,34,95,5,99,7,26,50,0,93,26,99,95,66,95,66,58,89,41,50,11], [10,63,10,89,25,89,25,63,10,63], [75,63,75,89,90,89,90,63,75,63]],
  g39: [[0,21,98,17,99,98,2,99,0,21], [10,26,10,89,90,89,90,26,10,26], [60,74,75,58,60,42,45,58,60,74], [60,84,35,58,60,32,85,58,60,84], [10,0,40,11,10,0]],
  g40: [[89,67,89,11,11,11,11,67,89,67], [0,6,98,2,98,98,2,98,0,6], [50,50,61,39,50,28,39,39,50,50], [50,61,28,39,50,17,72,39,50,61], [17,17,28,28,17,17], [17,78,17,89,83,89,83,78,17,78]],
  g41: [[50,82,19,100,26,63,0,38,35,34,50,0,65,34,100,38,74,63,81,100,50,82], [50,71,69,82,80,45,59,42,50,22,41,42,20,45,31,82,50,71]],
  g42: [[74,0,66,2,59,9,0,99,1,100,60,10,66,4,74,2,81,4,87,10,99,28,100,27,86,7,79,1,74,0]],
  g43: [[2,0,0,5,28,81,45,99,57,99,69,89,80,70,100,11,98,8,73,75,63,88,51,92,40,88,29,75,2,0]],
  g44: [[0,0,100,0,100,67,98,73,71,93,57,99,40,99,26,93,0,70,0,0]],
  g45: [[50,0,85,15,100,50,85,85,50,100,15,85,0,50,15,15,50,0]],
  g46: [[0,0,100,0,99,72,72,93,57,99,41,99,26,93,4,75,0,69,0,0]],
  g47: [[0,0,100,0,100,67,98,73,71,93,57,99,41,99,27,93,0,70,0,0]],
  g48: [[50,0,63,2,85,15,98,37,100,50,98,63,85,85,63,98,50,100,37,98,15,85,2,63,0,50,2,37,15,15,37,2,50,0]],
  g49: [[50,0,37,2,15,15,2,37,0,50,2,63,15,85,37,98,50,100,63,98,85,85,98,63,100,50,98,37,85,15,63,2,50,0], [50,80,42,79,29,71,20,50,29,29,42,21,50,20,58,21,71,29,80,50,71,71,58,79,50,80]],
  g50: [[50,0,31,4,14,15,4,31,0,50,2,63,15,85,37,98,50,100,69,96,86,85,97,69,100,50,98,37,85,15,63,2,50,0], [50,80,29,71,20,50,29,29,50,20,71,29,80,50,71,71,50,80]],
  g51: [[87,16,70,4,50,0,37,2,15,15,2,37,0,50,4,70,16,87,31,97,50,100,63,98,85,85,98,63,100,50,97,32,87,16], [50,80,29,72,22,62,20,50,29,29,50,20,62,22,72,29,80,50,71,71,50,80]],
  g52: [[50,0,37,2,15,15,2,37,0,50,2,63,15,85,37,98,50,100,63,98,85,85,98,63,100,50,98,37,85,15,63,2,50,0], [50,80,29,71,20,50,29,29,50,20,71,29,80,50,71,71,50,80]],
  g53: [[0,90,100,100,0,90], [10,50,20,85,10,50], [35,50,45,85,35,50], [55,50,65,85,55,50], [80,50,90,85,80,50], [0,25,50,0,100,25,100,45,0,45,0,25], [10,31,90,35,50,11,10,31]],
  g54: [[39,0,99,0,99,2,100,15,100,50,100,85,99,98,99,100,1,100,1,98,0,85,0,50,0,15,1,2,1,0,39,0]],
  g55: [[80,20,95,50,80,80,50,95,20,80,5,50,20,20,50,5,80,20]],
  g56: [[75,20,94,50,71,88,0,80,0,20,47,5,75,20]],
  g57: [[83,11,62,1,38,1,17,11,8,20,2,30,1,53,5,64,34,99,66,99,95,64,99,53,98,30,83,11]],
  g58: [[100,27,95,16,83,7,73,3,50,0,27,3,17,7,5,16,0,27,3,38,21,52,31,68,69,69,73,71,72,73,69,75,37,75,30,77,29,80,35,83,67,83,73,86,73,87,70,89,51,91,49,93,49,96,46,98,50,100,53,99,51,96,53,92,71,90,75,88,75,85,70,82,35,82,31,80,32,78,37,76,70,76,74,74,75,71,70,68,36,67,32,66,26,55,20,48,9,40,4,34,2,27,5,19,14,10,28,4,50,1,72,4,81,8,90,13,97,23,98,30,95,36,86,45,81,46,82,47,85,48,88,47,88,45,97,37,100,27]],
  g59: [[100,27,95,16,83,7,73,3,50,0,27,3,17,7,5,16,0,27,3,38,21,52,31,68,69,69,73,71,72,73,69,75,36,75,30,77,29,80,34,83,67,83,73,86,73,87,70,89,51,91,49,93,49,96,46,98,50,100,53,99,51,96,53,92,71,90,75,88,75,85,70,82,35,82,31,80,32,78,36,76,70,76,74,74,75,71,70,68,36,67,32,66,26,55,20,48,8,40,4,34,2,27,5,19,14,10,28,4,50,1,72,4,81,8,90,13,97,23,98,30,95,36,86,45,81,46,82,47,85,48,88,47,88,45,97,37,100,27]],
  g60: [[100,27,95,16,83,7,73,3,50,0,27,3,17,7,5,16,0,27,3,38,21,52,31,68,69,69,73,71,72,73,69,75,37,75,30,77,29,80,35,83,67,83,73,86,73,87,70,89,51,91,49,93,49,96,46,98,50,100,53,99,51,96,53,92,71,90,75,88,75,85,70,82,35,82,31,80,32,78,37,76,70,76,74,74,75,71,70,68,36,67,31,66,26,55,20,48,8,40,4,34,2,27,5,19,14,10,28,4,50,1,72,4,81,8,90,13,97,23,98,30,95,36,86,45,81,46,82,47,85,48,88,47,88,45,97,37,100,27]],
  g61: [[91,56,95,100,5,100,0,33,18,0,83,1,100,37,91,56], [82,60,18,60,18,90,82,90,82,60], [22,10,17,47,90,42,78,10,22,10]],
  g62: [[100,0,100,77,91,97,18,100,5,93,0,78,0,0,100,0]],
  g63: [[100,86,98,100,2,100,2,0,98,0,100,86]],
  g64: [[100,100,0,100,46,2,100,100]],
  g65: [[100,90,100,100,0,100,32,0,68,0,100,90]],
  g66: [[4,0,38,100,100,100,4,0]],
  g67: [[38,94,45,100,100,100,100,0,0,0,38,94]],
  g68: [[7,38,0,45,0,100,100,100,100,0,7,38]],
  g69: [[94,62,100,55,100,0,0,0,0,100,94,62]],
  g70: [[62,6,55,0,0,0,0,100,100,100,62,6]],
  g71: [[100,0,0,0,0,100,100,100,100,0]],
  g72: [[9,90,10,20,58,0,59,28,91,41,100,100,0,100,9,90], [18,90,50,90,50,13,18,26,18,90], [82,90,82,47,59,38,59,90,82,90]],
  g73: [[64,100,62,85,48,56,37,43,1,16,0,11,13,1,100,0]],
  g74: [[0,0,87,1,100,11,99,16,63,43,52,56,38,85,36,100]],
  g75: [[0,100,87,99,100,89,99,84,63,57,52,44,38,15,36,0]],
  g76: [[100,100,13,99,0,89,1,84,37,57,48,44,62,15,64,0]],
  g77: [[100,50,98,63,93,75,85,85,75,93,63,98,50,100,37,98,25,93,15,85,7,75,2,63,0,50,2,37,7,25,15,15,25,7,37,2,50,0,63,2,75,7,85,15,93,25,98,37,100,50]],
  g78: [[50,100,31,96,15,85,4,69,0,50,4,31,15,15,31,4,50,0,69,4,85,15,96,31,100,50,96,69,85,85,69,96,50,100]],
  g79: [[100,50,85,85,50,100,15,85,0,50,15,15,50,0,85,15,100,50]],
  g80: [[50,10,22,22,10,50,22,78,50,90,78,78,90,50,78,22,50,10], [50,0,85,15,100,50,85,85,50,100,15,85,0,50,15,15,50,0], [72,45,56,81,56,59,38,54,53,50,29,39,53,43,26,22,72,45]],
  g81: [[34,75,75,82,50,100,25,82,34,75], [30,30,41,22,30,13,18,21,30,30], [39,37,61,37,39,37], [59,51,82,52,59,51], [50,15,70,0,100,21,79,37,100,52,70,73,50,58,30,73,0,52,21,37,0,21,30,0,50,15], [59,22,82,21,59,22], [41,51,18,52,41,51]],
  g82: [[67,40,63,16,39,14,22,80,78,80,67,40], [15,40,32,6,62,2,100,78,53,100,16,91,0,77,22,57,15,40]],
  g83: [[50,90,78,78,90,50,78,22,50,10,22,22,10,50,22,78,50,90], [50,100,15,85,0,50,15,15,50,0,85,15,100,50,85,85,50,100], [34,56,20,50,74,29,65,75,40,72,34,56]],
  g84: [[11,11,11,89,89,89,89,11,11,11], [11,0,97,3,97,97,3,97,11,0], [28,22,44,26,40,78,22,74,28,22], [61,22,78,26,74,55,56,51,61,22]],
  g85: [[41,86,46,95,1,93,2,1,65,0,90,24,90,46,60,29,60,10,10,10,10,86,41,86], [75,100,50,76,75,52,100,76,75,100], [69,89,89,70,69,89], [61,82,81,63,61,82]],
  g86: [[49,93,45,84,81,84,81,54,93,57,99,54,100,50,99,46,93,43,81,47,81,16,45,16,49,6,40,0,32,6,36,16,0,16,0,47,12,43,18,46,19,50,18,54,12,57,0,54,0,84,36,84,32,93,33,96,38,100,45,99,49,93]],
  g87: [[7,49,16,45,16,81,46,81,43,93,46,99,50,100,54,99,57,93,53,81,84,81,84,45,94,49,100,40,94,32,84,36,84,0,53,0,57,12,54,18,50,19,46,18,43,12,46,0,16,0,16,36,7,32,4,33,0,38,1,45,7,49]],
  g88: [[91,51,81,55,81,19,45,19,49,9,43,0,36,1,32,7,36,19,0,19,0,55,12,51,19,59,12,68,0,64,0,100,36,100,32,88,40,81,49,88,45,100,81,100,81,64,91,68,100,62,99,55,91,51]],
  g89: [[51,9,55,19,19,19,19,55,9,51,0,57,1,64,7,68,19,64,19,100,55,100,51,88,59,81,68,88,64,100,100,100,100,64,88,68,81,59,88,51,100,55,100,19,64,19,68,9,62,0,55,1,51,9]],
  g90: [[50,0,0,100,100,100,50,0]],
  g91: [[25,0,0,100,100,100,75,0,25,0]],
  g92: [[0,100,100,100,83,0,17,0,0,100]],
  g93: [[0,100,100,100,88,0,13,0,0,100]],
  g94: [[0,100,100,100,90,0,10,0,0,100]],
  g95: [[0,100,100,100,92,0,8,0,0,100]],
  g96: [[50,0,15,9,0,32,50,100,100,32,85,9,50,0], [50,50,21,38,28,17,78,24,75,42,50,50], [31,32,45,43,69,32,54,19,31,32]],
  g97: [[18,0,40,19,29,37,43,57,99,78,87,97,53,92,14,57,0,20,18,0]],
  g98: [[4,0,68,33,100,96,75,100,50,52,0,26,4,0]],
  g99: [[2,0,68,34,100,97,84,100,61,49,0,16,2,0]],
  g100: [[54,76,54,89,54,76], [46,76,46,89,46,76], [79,55,82,76,91,55,79,55], [54,55,68,69,70,55,54,55], [30,55,46,66,30,55], [9,55,18,76,21,55,9,55], [68,31,54,45,70,45,68,31], [32,31,30,45,46,45,32,31], [18,24,9,45,21,45,18,24], [82,24,79,45,91,45,82,24], [54,11,64,22,54,11], [46,11,36,22,46,11], [50,0,85,15,100,50,85,85,50,100,15,85,0,50,15,15,50,0]],
};

// ------------------------------------------------------------------ helpers
let SH = null;   // pres.ShapeType   (set in build)
let CT = null;   // pres.ChartType   (set in build)

// Slide on the sand canvas, with the master's "Page NN" chip bottom-right.
function page(pres, no) {
  const s = pres.addSlide();
  s.background = { color: BG };
  box(s, 12.166, 7.029, 0.73, 0.258, CARD);
  box(s, 12.79, 7.029, 0.304, 0.258, STEEL);
  s.addText('Page', { x: 12.2, y: 7.001, w: 0.556, h: 0.28, fontFace: BODY, fontSize: 8, color: GREY, align: 'center', valign: 'middle' });
  s.addText(String(no), { x: 12.685, y: 6.959, w: 0.514, h: 0.399, fontFace: BODY, fontSize: 9, color: WHITE, align: 'center', valign: 'middle' });
  return s;
}

const box = (s, x, y, w, h, fill) =>
  s.addShape(SH.rect, { x, y, w, h, fill: { color: fill }, line: NOLINE });

const card = (s, x, y, w, h, fill) =>
  s.addShape(SH.rect, { x, y, w, h, fill: { color: fill }, line: NOLINE, shadow: LIFT });

const rule = (s, x, y, w, h, line, more) =>
  s.addShape(SH.line, Object.assign({ x, y, w, h, line }, more));

// Elbow connector - pptxgenjs has no enum entry, the preset name works directly.
const bent = (s, x, y, w, h, line, more) =>
  s.addShape('bentConnector3', Object.assign({ x, y, w, h, line }, more));

const tx = (s, x, y, w, h, text, style, more) =>
  s.addText(text, Object.assign({ x, y, w, h, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] }, style, more));

// Freeform art: rings are scaled to the w x h box (custGeom points are
// relative to the shape origin).
function art(s, rings, x, y, w, h, fill, more) {
  const points = [];
  rings.forEach((ring) => {
    for (let i = 0; i < ring.length; i += 2) {
      const p = { x: (ring[i] / 100) * w, y: (ring[i + 1] / 100) * h };
      if (i === 0) p.moveTo = true;
      points.push(p);
    }
    points.push({ close: true });
  });
  s.addShape(SH.custGeom, Object.assign(
    { x, y, w, h, points, line: NOLINE, fill: fill ? { color: fill } : NOLINE }, more));
}

// Stand-in for a photograph in the source deck: a block in the average tone of
// the original artwork, sized to the area the artwork actually covered.
function photo(s, x, y, w, h, tone) {
  box(s, x, y, w, h, tone);
  s.addText('[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, fontFace: BODY, fontSize: 11, color: CARD, align: 'center', valign: 'middle' });
}

// Percentage ring (chart infographic slide).
function donut(s, x, y, w, h, pct, done, rest) {
  s.addChart(CT.doughnut,
    [{ name: 'Sales', labels: ['Done', 'Rest'], values: [pct, 100 - pct] }],
    { x, y, w, h, holeSize: 84, showLegend: false, showValue: false,
      chartColors: [done, rest], dataBorder: { pt: 0, color: BG } });
}

// ---- 1  House Real Estate Presentation
function slide1(pres) {
  const s = page(pres, 1);
  box(s, 2.688, 0, 10.646, 7.5, STEEL);   // layout backdrop
  tx(s, 3.729, 0.958, 9.412, 2.322, 'House Real Estate Presentation', { fontFace: HEAD, fontSize: 66, color: WHITE });
  tx(s, 7.999, 4.505, 3.094, 0.404, 'Presentation Template', S.h18W);
  tx(s, 7.999, 4.897, 4.292, 0.903, LOREM4, S.p11W);
  tx(s, 0.554, 0.611, 2.279, 0.348, 'www.reallygreatsite.com', S.p11);
  rule(s, 0.684, 1.118, 0, 2.163, { color: GREY, width: 1 }, { flipV: true });
  rule(s, 12.667, 1.292, 0, 5.187, { color: CARD, width: 1 }, { flipV: true });
  tx(s, 10.533, 6.577, 2.279, 0.348, 'DESIGN TEMPLATE', { fontFace: BODY, fontSize: 11, color: CARD, align: 'right', lineSpacingMultiple: 1.5 });
}

// ---- 2  Introduce Home Real Estate Here
function slide2(pres) {
  const s = page(pres, 2);
  box(s, 0, 3.909, 2.263, 3.591, STEEL);   // layout backdrop
  box(s, 9.542, 0, 3.792, 4.929, STEEL);   // layout backdrop
  tx(s, 1.097, 0.581, 5.245, 1.178, 'Introduce Home Real Estate Here', S.h32K);
  tx(s, 1.097, 1.912, 5.57, 0.903, lorem(), S.p11);
  tx(s, 1.097, 2.843, 5.57, 0.625, lorem(131), S.p11);
}

// ---- 3  01.
function slide3(pres) {
  const s = page(pres, 3);
  card(s, 9.091, 4.776, 4.243, 1.798, STEEL);
  card(s, 4.818, 4.776, 4.273, 1.798, CARD);
  tx(s, 5.703, 5.652, 2.806, 0.625, lorem(68), { fontFace: BODY, fontSize: 10.5, color: GREY, lineSpacingMultiple: 1.5 });
  tx(s, 5.703, 5.073, 0.877, 0.64, '01.', S.h32K_3);
  tx(s, 6.407, 5.276, 1.809, 0.37, TITLE, S.h16K);
  tx(s, 9.96, 5.652, 2.806, 0.625, lorem(68), { fontFace: BODY, fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.5 });
  tx(s, 9.96, 5.073, 0.877, 0.64, '02.', S.h32K_3);
  tx(s, 10.665, 5.276, 1.809, 0.37, TITLE, S.h16W_2);
  tx(s, 5.703, 0.926, 5.245, 1.178, 'Introduce Home Real Estate Here', S.h32K);
  tx(s, 5.703, 2.59, 5.368, 0.903, lorem(), S.p11);
  tx(s, 5.703, 3.52, 5.368, 0.625, lorem(131), S.p11);
}

// ---- 4  New Concept Home Real Estate Here
function slide4(pres) {
  const s = page(pres, 4);
  box(s, 6.667, 0, 6.667, 5.925, STEEL);   // layout backdrop
  tx(s, 0.875, 1.322, 5.245, 1.178, 'New Concept Home Real Estate Here', S.h32K);
  tx(s, 7.735, 1.459, 5.368, 0.903, lorem(), S.p11W);
  tx(s, 9.431, 3.747, 2.704, 0.404, 'New Concept Ideas', S.h18W);
  tx(s, 9.431, 3.107, 1.264, 0.64, '75%', S.h32W);
  tx(s, 9.431, 4.992, 2.704, 0.404, 'Happy Clients', S.h18W);
  tx(s, 9.431, 4.352, 1.264, 0.64, '125K', S.h32W);
}

// ---- 5  Best Property Home Real Estate Here
function slide5(pres) {
  const s = page(pres, 5);
  box(s, 0, -0.053, 6.667, 3.972, STEEL);   // layout backdrop
  tx(s, 0.875, 0.897, 5.245, 1.178, 'Best Property Home Real Estate Here', { fontFace: HEAD, fontSize: 32, color: WHITE });
  tx(s, 7.735, 1.035, 5.368, 0.903, lorem(), S.p11K);
  card(s, 0.849, 3.323, 4.069, 1.634, CARD);
  card(s, 0.849, 5.358, 4.069, 1.634, CARD);
  tx(s, 1.801, 4.024, 2.724, 0.625, lorem(68), S.p11);
  tx(s, 1.801, 3.63, 2.421, 0.413, 'Property One', S.h14K);
  art(s, OUT.g1, 1.492, 4.031, 0.234, 0.24, STEEL);
  art(s, OUT.g2, 1.243, 3.795, 0.402, 0.508, STEEL);
  art(s, OUT.g3, 1.258, 3.812, 0.099, 0.106, STEEL);
  tx(s, 1.821, 6.059, 2.724, 0.625, lorem(68), S.p11K);
  tx(s, 1.821, 5.665, 2.421, 0.413, 'Property Two', S.h14K);
  art(s, OUT.g4, 1.224, 5.83, 0.516, 0.541, STEEL);
}

// ---- 6  Welcome Message From Our Company
function slide6(pres) {
  const s = page(pres, 6);
  box(s, 7, 0, 3.354, 4.958, STEEL);   // layout backdrop
  box(s, 0, 5.021, 1.863, 2.479, STEEL);   // layout backdrop
  tx(s, 0.955, 1.56, 5.245, 1.178, 'Welcome Message From Our Company', S.h32K);
  tx(s, 0.955, 3.014, 5.712, 0.903, LOREM8, S.p11);
  tx(s, 2.916, 4.484, 3.094, 0.404, 'Our CEO Founder', S.h18K);
  tx(s, 2.916, 4.876, 3.751, 0.903, lorem(136), S.p11);
}

// ---- 7  Our Mission Home Real Estate Here
function slide7(pres) {
  const s = page(pres, 7);
  box(s, 5.204, 3.97, 5.15, 3.53, STEEL);   // layout backdrop
  tx(s, 0.878, 0.817, 4.432, 1.178, 'Our Mission Home Real Estate Here', S.h32K);
  tx(s, 0.878, 2.31, 5.712, 0.903, lorem(), S.p11);
  card(s, 4.411, 4.954, 2.479, 1.292, CARD);
  card(s, 7.069, 4.954, 2.479, 1.292, CARD);
  tx(s, 4.572, 5.009, 0.877, 0.64, '01.', S.h32K_3);
  tx(s, 5.277, 5.213, 2.175, 0.37, 'Mission', S.h16K);
  tx(s, 4.572, 5.53, 2.342, 0.53, LOREM1, S.p9);
  tx(s, 7.239, 5.009, 0.877, 0.64, '02.', S.h32K_3);
  tx(s, 7.943, 5.213, 2.175, 0.37, 'Mission', S.h16K);
  tx(s, 7.239, 5.53, 2.342, 0.53, LOREM1, S.p9);
}

// ---- 8  Our Vision Home Real Estate Here
function slide8(pres) {
  const s = page(pres, 8);
  box(s, 3.818, 4.75, 3.911, 1.981, STEEL);
  tx(s, 0.917, 2.671, 5.802, 1.182, LOREM12, S.p11j);
  tx(s, 0.917, 1.146, 5.29, 1.178, 'Our Vision Home Real Estate Here', S.h32K);
  tx(s, 4.334, 4.912, 0.877, 0.64, '01.', S.h32W);
  tx(s, 5.038, 5.116, 2.175, 0.37, 'Our Vision', S.h16W_2);
  tx(s, 4.334, 5.451, 2.879, 0.903, lorem(103), S.p11W);
}

// ---- 9  About Our Home Real Estate Here
function slide9(pres) {
  const s = page(pres, 9);
  box(s, 0, 3.75, 6.271, 3.75, STEEL);   // layout backdrop
  tx(s, 6.83, 1.68, 5.245, 1.178, 'About Our Home Real Estate Here', S.h32K);
  tx(s, 6.83, 3.134, 5.424, 0.903, LOREM6, S.p11);
  card(s, 6.938, 4.528, 2.479, 1.292, CARD);
  card(s, 9.596, 4.528, 2.479, 1.292, CARD);
  tx(s, 7.099, 4.583, 0.877, 0.64, '01.', S.h32K_3);
  tx(s, 7.803, 4.787, 2.175, 0.37, TITLE, S.h16K);
  tx(s, 7.099, 5.105, 2.342, 0.53, LOREM1, S.p9);
  tx(s, 9.766, 4.583, 0.877, 0.64, '02.', S.h32K_3);
  tx(s, 10.47, 4.787, 2.175, 0.37, TITLE, S.h16K);
  tx(s, 9.766, 5.105, 2.342, 0.53, LOREM1, S.p9);
}

// ---- 10  About Real Estate Service For Customer
function slide10(pres) {
  const s = page(pres, 10);
  box(s, 0, 0, 3.688, 7.5, STEEL);   // layout backdrop
  tx(s, 6.922, 2.029, 5.245, 1.178, 'About Real Estate Service For Customer', S.h32K);
  tx(s, 6.922, 3.575, 5.245, 0.903, lorem(), S.p11);
  tx(s, 6.922, 4.568, 5.245, 0.903, lorem(), S.p11);
  card(s, 1.792, 2.25, 3.833, 1, CARD);
  card(s, 1.792, 4.146, 3.833, 1, CARD);
  tx(s, 2.769, 2.499, 2.421, 0.502, 'Service One Here', S.h18K_2);
  tx(s, 2.769, 4.395, 2.421, 0.502, 'Service Two Here', S.h18K_2);
  art(s, OUT.g5, 2.582, 2.586, 0.057, 0.057, STEEL);
  art(s, OUT.g6, 2.351, 2.579, 0.295, 0.342, STEEL);
  art(s, OUT.g7, 2.303, 4.472, 0.343, 0.348, STEEL);
}

// ---- 11  Make Your Dream House Come True
function slide11(pres) {
  const s = page(pres, 11);
  tx(s, 0.984, 1.716, 5.245, 1.178, 'Make Your Dream House Come True', S.h32K);
  tx(s, 0.984, 3.194, 5.424, 1.181, LOREM11, S.p11);
  card(s, 1.068, 4.896, 5.341, 1.388, CARD);
  card(s, 6.925, 4.896, 5.341, 1.388, STEEL);
  tx(s, 2.211, 5.471, 3.872, 0.625, LOREM3, S.p11);
  tx(s, 2.211, 5.077, 2.421, 0.413, 'Service One Here', S.h14K);
  tx(s, 8.069, 5.471, 3.872, 0.625, LOREM3, S.p11W);
  tx(s, 8.069, 5.077, 2.421, 0.413, 'Service Two Here', S.h14W);
  art(s, OUT.g1, 1.902, 5.478, 0.234, 0.24, STEEL);
  art(s, OUT.g2, 1.653, 5.242, 0.402, 0.508, STEEL);
  art(s, OUT.g3, 1.668, 5.259, 0.099, 0.106, STEEL);
  art(s, OUT.g4, 7.472, 5.242, 0.516, 0.541, CARD);
}

// ---- 12  Pick Your Dream House With Us Now
function slide12(pres) {
  const s = page(pres, 12);
  box(s, 0, 0, 3.688, 4.188, STEEL);   // layout backdrop
  tx(s, 6.005, 1.738, 5.245, 1.178, 'Pick Your Dream House With Us Now', S.h32K);
  tx(s, 6.005, 3.284, 6.099, 0.903, lorem(), S.p11);
  tx(s, 9.896, 5.843, 3.094, 0.404, 'Best House', S.h18K);
  tx(s, 9.896, 6.234, 2.583, 0.625, lorem(63), S.p11);
  art(s, OUT.g8, 10.024, 5.44, 0.318, 0.327, STEEL);
  art(s, OUT.g9, 9.986, 5.348, 0.395, 0.227, STEEL);
}

// ---- 13  We Always Have House Solution For You
function slide13(pres) {
  const s = page(pres, 13);
  box(s, 9.542, 0, 3.792, 7.5, STEEL);   // layout backdrop
  box(s, 1.028, 3.49, 8.513, 2.115, CARD);   // layout backdrop
  photo(s, 8.062, 1.466, 5.271, 5.196, 'C0BEBC');
  tx(s, 1.028, 1.667, 5.245, 1.178, 'We Always Have House Solution For You', S.h32K);
  tx(s, 2.028, 4.276, 2.539, 0.903, LOREM2, S.p11);
  tx(s, 2.028, 3.899, 3.094, 0.404, 'Aspect One', S.h18K);
  tx(s, 5.239, 4.276, 2.539, 0.903, LOREM2, S.p11);
  tx(s, 5.239, 3.899, 3.094, 0.404, 'Aspect Two', S.h18K);
  art(s, OUT.g10, 1.681, 4.096, 0.03, 0.03, STEEL);
  art(s, OUT.g11, 1.776, 4.156, 0.03, 0.03, STEEL);
  art(s, OUT.g12, 1.582, 3.979, 0.323, 0.323, STEEL);
  art(s, OUT.g13, 4.846, 4.05, 0.11, 0.139, STEEL);
  art(s, OUT.g14, 4.991, 3.982, 0.163, 0.09, STEEL);
  art(s, OUT.g15, 4.887, 3.978, 0.136, 0.145, STEEL);
  art(s, OUT.g16, 4.86, 4.21, 0.164, 0.092, STEEL);
  art(s, OUT.g17, 5.057, 4.092, 0.111, 0.143, STEEL);
  art(s, OUT.g18, 4.993, 4.161, 0.133, 0.144, STEEL);
}

// ---- 14  85
function slide14(pres) {
  const s = page(pres, 14);
  box(s, 0, 0, 2.833, 7.5, STEEL);   // layout backdrop
  photo(s, 1.204, 0.855, 2.915, 5.846, 'C9C7C4');
  card(s, 4.623, 1.714, 2.544, 4.156, CARD);
  s.addShape(SH.ellipse, { x: 5.467, y: 2.483, w: 0.867, h: 0.867, fill: { color: SAND }, line: NOLINE });
  s.addShape(SH.arc, { x: 5.44, y: 2.456, w: 0.921, h: 0.921, line: { color: STEEL, width: 7 } });
  tx(s, 5.148, 2.701, 1.505, 0.438, [{ text: '85', options: { fontSize: 20, bold: true } }, { text: '%', options: { fontSize: 11, bold: true } }], { fontFace: HEAD, fontSize: 20, color: GREY, bold: true, align: 'center' });
  s.addShape(SH.ellipse, { x: 5.314, y: 2.696, w: 0.268, h: 0.268, fill: { color: STEEL2 }, shadow: { type: 'outer', color: INK, opacity: 0.1, blur: 20 }, line: NOLINE });
  s.addShape(SH.ellipse, { x: 5.384, y: 2.765, w: 0.13, h: 0.13, fill: { color: WHITE }, shadow: { type: 'outer', color: INK, opacity: 0.1, blur: 20 }, line: NOLINE });
  tx(s, 4.944, 4.056, 1.91, 1.181, LOREM2, S.p11c);
  tx(s, 4.352, 3.679, 3.094, 0.404, 'House Sale', S.h18Kc);
  tx(s, 7.796, 2.082, 4.432, 1.178, 'Talk About Dream House With Us', S.h32K);
  tx(s, 7.796, 3.522, 4.494, 0.903, lorem(179), S.p11);
  tx(s, 7.796, 4.514, 4.494, 0.903, lorem(179), S.p11);
}

// ---- 15  Real Estate Marketing Strategy
function slide15(pres) {
  const s = page(pres, 15);
  card(s, 4.989, 5.115, 8.344, 1.688, CARD);
  tx(s, 6.341, 5.508, 5.641, 0.903, lorem(), S.p11);
  tx(s, 6.341, 3.519, 4.432, 1.178, 'Real Estate Marketing Strategy', S.h32K);
  box(s, 0.714, 1.72, 4.275, 5.78, STEEL);
  tx(s, 1.412, 2.494, 0.877, 0.64, '01.', S.h32W);
  tx(s, 2.116, 2.698, 2.175, 0.37, TITLE, S.h16W_2);
  tx(s, 1.412, 3.033, 2.879, 0.903, lorem(103), S.p11W);
  tx(s, 1.412, 4.028, 0.877, 0.64, '02.', S.h32W);
  tx(s, 2.116, 4.232, 2.175, 0.37, TITLE, S.h16W_2);
  tx(s, 1.412, 4.567, 2.879, 0.903, lorem(103), S.p11W);
  tx(s, 1.412, 5.562, 0.877, 0.64, '03.', S.h32W);
  tx(s, 2.116, 5.766, 2.175, 0.37, TITLE, S.h16W_2);
  tx(s, 1.412, 6.101, 2.879, 0.625, lorem(68), S.p11W);
}

// ---- 16  Reporting Income Home Real Estate
function slide16(pres) {
  const s = page(pres, 16);
  box(s, 9.749, 3.583, 3.584, 2.347, STEEL);   // layout backdrop
  box(s, 0, 0, 4.01, 7.5, STEEL);   // layout backdrop
  tx(s, 7.513, 0.731, 4.432, 1.178, 'Reporting Income Home Real Estate', S.h32K);
  tx(s, 7.513, 2.224, 5.245, 0.903, lorem(), S.p11);
  card(s, 0.67, 4.896, 5.341, 1.388, CARD);
  tx(s, 1.73, 5.474, 3.872, 0.625, LOREM3, S.p11);
  tx(s, 1.73, 5.08, 2.421, 0.413, 'Other Facility', S.h14K);
  art(s, OUT.g8, 1.135, 5.317, 0.475, 0.489, STEEL);
  art(s, OUT.g20, 1.078, 5.18, 0.59, 0.339, STEEL);
}

// ---- 17  The Best Modern Real Estate Homes
function slide17(pres) {
  const s = page(pres, 17);
  box(s, 0, 0, 5.758, 3.152, STEEL);   // layout backdrop
  tx(s, 0.743, 5.461, 4.601, 1.178, 'The Best Modern Real Estate Homes', S.h32K);
  tx(s, 8.426, 5.46, 4.165, 1.181, lorem(), S.p11);
  card(s, 4.995, 2.719, 4.333, 2.232, CARD);
  tx(s, 6.059, 3.572, 2.572, 0.903, LOREM2, S.p11);
  tx(s, 6.059, 3.195, 3.094, 0.404, 'New Style 01', S.h18K);
  art(s, OUT.g13, 5.665, 3.696, 0.11, 0.139, STEEL);
  art(s, OUT.g14, 5.81, 3.628, 0.163, 0.09, STEEL);
  art(s, OUT.g15, 5.706, 3.624, 0.136, 0.145, STEEL);
  art(s, OUT.g16, 5.679, 3.856, 0.164, 0.092, STEEL);
  art(s, OUT.g17, 5.876, 3.738, 0.111, 0.143, STEEL);
  art(s, OUT.g18, 5.812, 3.807, 0.133, 0.144, STEEL);
}

// ---- 18  Our Latest Project Home Real Estate
function slide18(pres) {
  const s = page(pres, 18);
  box(s, 7.121, 1.919, 6.212, 5.581, STEEL);   // layout backdrop
  tx(s, 0.779, 1.317, 4.601, 1.178, 'Our Latest Project Home Real Estate', S.h32K);
  tx(s, 0.779, 3.189, 3.094, 0.404, 'Project One', S.h18K);
  tx(s, 0.779, 3.593, 4.601, 0.903, lorem(185), S.p11);
  tx(s, 0.779, 4.876, 3.094, 0.404, 'Project Two', S.h18K);
  tx(s, 0.779, 5.28, 4.601, 0.903, lorem(185), S.p11);
}

// ---- 19  Home Real Estate Portfolio Here
function slide19(pres) {
  const s = page(pres, 19);
  box(s, 0, 0, 3.788, 5.697, STEEL);   // layout backdrop
  tx(s, 5.58, 0.623, 4.432, 1.178, 'Home Real Estate Portfolio Here', S.h32K);
  tx(s, 10.054, 3.699, 2.889, 0.404, '2022 portfolio', S.h18K);
  tx(s, 10.054, 4.103, 2.889, 1.736, lorem(), S.p11);
  tx(s, 5.58, 1.993, 6.345, 0.625, lorem(169), S.p11);
}

// ---- 20  Meet Our The Best Team Here
function slide20(pres) {
  const s = page(pres, 20);
  box(s, 8.229, 0, 5.104, 7.5, STEEL);
  tx(s, 0.883, 2.725, 5.644, 0.904, LOREM7, S.p11j);
  tx(s, 9.849, 1.498, 2.462, 0.348, GM, S.p11W_2);
  tx(s, 9.849, 1.144, 2.462, 0.466, VISION, S.h16W);
  tx(s, 9.849, 1.814, 2.778, 0.625, lorem(65), S.p11W);
  tx(s, 9.849, 3.456, 2.462, 0.348, CEOF, S.p11W_2);
  tx(s, 9.849, 3.102, 2.462, 0.466, JONAH, S.h16W);
  tx(s, 9.849, 3.772, 2.778, 0.625, lorem(65), S.p11W);
  tx(s, 9.849, 5.414, 2.462, 0.348, EXEC, S.p11W_2);
  tx(s, 9.849, 5.061, 2.462, 0.466, HARRY, S.h16W);
  tx(s, 9.849, 5.73, 2.778, 0.625, lorem(65), S.p11W);
  tx(s, 0.883, 1.312, 4.432, 1.178, 'Meet Our The Best Team Here', S.h32K);
}

// ---- 21  Meet Our Best Team Here
function slide21(pres) {
  const s = page(pres, 21);
  box(s, 0, 0, 13.333, 5.438, STEEL);   // layout backdrop
  tx(s, 2.396, 0.997, 8.542, 0.64, 'Meet Our Best Team Here', { fontFace: HEAD, fontSize: 32, color: WHITE, align: 'center' });
  tx(s, 1.078, 2.459, 2.462, 0.348, GM, S.p11W_2);
  tx(s, 1.078, 2.106, 2.462, 0.466, VISION, S.h16W);
  tx(s, 1.078, 2.775, 2.778, 0.625, lorem(65), S.p11W);
  tx(s, 5.266, 2.459, 2.462, 0.348, CEOF, S.p11W_2);
  tx(s, 5.266, 2.106, 2.462, 0.466, JONAH, S.h16W);
  tx(s, 5.266, 2.775, 2.778, 0.625, lorem(65), S.p11W);
  tx(s, 9.478, 2.459, 2.462, 0.348, EXEC, S.p11W_2);
  tx(s, 9.478, 2.106, 2.462, 0.466, HARRY, S.h16W);
  tx(s, 9.478, 2.775, 2.778, 0.625, lorem(65), S.p11W);
}

// ---- 22  About Our General Manager Here
function slide22(pres) {
  const s = page(pres, 22);
  box(s, 10.413, 0, 2.921, 5.952, STEEL);   // layout backdrop
  tx(s, 1.143, 1.169, 5.245, 1.178, 'About Our General Manager Here', S.h32K);
  tx(s, 1.143, 3.41, 5.81, 0.625, lorem(161), S.p11);
  tx(s, 1.143, 3.005, 2.462, 0.348, GM, S.p11_2);
  tx(s, 1.143, 2.652, 3.623, 0.466, VISION, S.h16K_2);
  card(s, 1.241, 4.328, 1.264, 0.921, CARD);
  tx(s, 1.184, 4.359, 1.379, 0.64, '80%', S.h32S2c);
  tx(s, 1.237, 4.862, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 2.68, 4.476, 4.273, 0.625, lorem(113), S.p11);
  card(s, 1.241, 5.41, 1.264, 0.921, CARD);
  tx(s, 1.184, 5.441, 1.379, 0.64, '90%', S.h32S2c);
  tx(s, 1.237, 5.943, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 2.68, 5.558, 4.273, 0.625, lorem(113), S.p11);
}

// ---- 23  About Our CEO Founder Here
function slide23(pres) {
  const s = page(pres, 23);
  card(s, 6.396, 3.25, 5.958, 3.229, CARD);
  card(s, 0.979, 3.25, 5.083, 3.229, STEEL);
  tx(s, 6.667, 3.623, 5.245, 1.178, 'About Our CEO Founder Here', { fontFace: HEAD, fontSize: 32, color: INK, align: 'right' });
  tx(s, 6.875, 5.443, 4.997, 0.625, lorem(137), S.p11r);
  tx(s, 7.842, 5.081, 2.462, 0.348, CEOF, { fontFace: BODY, fontSize: 11, color: GREY, italic: true, align: 'right', lineSpacingMultiple: 1.5 });
  tx(s, 8.249, 4.977, 3.623, 0.466, JONAH, { fontFace: HEAD, fontSize: 16, color: INK, align: 'right', lineSpacingMultiple: 1.5 });
  card(s, 1.542, 3.742, 1.264, 0.921, CARD);
  tx(s, 1.484, 3.773, 1.379, 0.64, '80%', S.h32Sc);
  tx(s, 1.538, 4.275, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 2.98, 3.889, 2.79, 0.625, lorem(68), S.p11W);
  card(s, 1.542, 5.094, 1.264, 0.921, CARD);
  tx(s, 1.484, 5.125, 1.379, 0.64, '90%', S.h32Sc);
  tx(s, 1.538, 5.627, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 2.98, 5.242, 2.79, 0.625, lorem(68), S.p11W);
}

// ---- 24  About Our Executive Manager Here
function slide24(pres) {
  const s = page(pres, 24);
  tx(s, 0.971, 1.169, 5.245, 1.178, 'About Our Executive Manager Here', S.h32K);
  tx(s, 0.971, 3.41, 4.461, 0.903, lorem(161), S.p11);
  tx(s, 0.971, 3.005, 2.462, 0.348, GM, S.p11_2);
  tx(s, 0.971, 2.652, 3.623, 0.466, VISION, S.h16K_2);
  card(s, 6.896, 3.167, 5.083, 3.229, STEEL);
  card(s, 7.458, 3.658, 1.264, 0.921, CARD);
  tx(s, 7.401, 3.689, 1.379, 0.64, '80%', S.h32Sc);
  tx(s, 7.454, 4.191, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 8.897, 3.806, 2.79, 0.625, lorem(68), S.p11W);
  card(s, 7.458, 5.011, 1.264, 0.921, CARD);
  tx(s, 7.401, 5.042, 1.379, 0.64, '90%', S.h32Sc);
  tx(s, 7.454, 5.544, 1.272, 0.302, SKILL, S.h9Kc);
  tx(s, 8.897, 5.158, 2.79, 0.625, lorem(68), S.p11W);
}

// ---- 25  Break Time
function slide25(pres) {
  const s = page(pres, 25);
  box(s, 2.444, 2.465, 10.889, 4.217, STEEL);   // layout backdrop
  tx(s, 1.879, 0.605, 9.576, 1.447, 'Break Time', S.h80Kc);
}

// ---- 26  Break Time
function slide26(pres) {
  const s = page(pres, 26);
  box(s, 2.007, 1.624, 4.484, 4.484, STEEL);
  card(s, 0.53, 4.387, 8.535, 2.591, CARD);
  tx(s, 0.883, 4.959, 7.828, 1.447, 'Break Time', S.h80Kc);
}

// ---- 27  Mockup Device Home Real Estate Here
function slide27(pres) {
  const s = page(pres, 27);
  box(s, 0, 0, 3.168, 7.5, STEEL);
  box(s, 3.332, 0, 3.168, 7.5, STEEL);
  tx(s, 7.248, 2.749, 5.102, 1.182, LOREM10, S.p11j);
  tx(s, 7.248, 4.027, 5.102, 0.627, lorem(128), S.p11j);
  card(s, 8.997, 5.179, 3.354, 1.215, CARD);
  tx(s, 9.278, 5.441, 2.772, 0.627, lorem(63), { fontFace: BODY, fontSize: 11, color: INK, align: 'justify', lineSpacingMultiple: 1.5 });
  photo(s, 1.522, 0.948, 3.634, 5.613, 'CBC9C7');
  tx(s, 7.248, 1.042, 5.245, 1.178, MOCKUP, S.h32K);
}

// ---- 28  Mockup Device Home Real Estate Here
function slide28(pres) {
  const s = page(pres, 28);
  box(s, 0, 4.75, 13.333, 2.75, STEEL);
  card(s, 1.751, 3.862, 5.341, 1.388, CARD);
  tx(s, 2.812, 4.44, 3.872, 0.625, LOREM3, S.p11);
  tx(s, 2.812, 4.046, 2.421, 0.413, TITTLE, S.h14K);
  art(s, OUT.g8, 2.217, 4.283, 0.475, 0.489, STEEL);
  art(s, OUT.g20, 2.16, 4.147, 0.59, 0.339, STEEL);
  photo(s, 6.393, 2.249, 6.475, 3.754, 'C4C2C0');
  tx(s, 0.996, 2.284, 5.102, 1.182, LOREM10, S.p11j);
  tx(s, 0.996, 0.576, 5.245, 1.178, MOCKUP, S.h32K);
}

// ---- 29  Mockup Device Home Real Estate Here
function slide29(pres) {
  const s = page(pres, 29);
  box(s, 4.82, 4.499, 8.513, 2.115, CARD);   // layout backdrop
  box(s, 0, 0, 4.366, 7.5, STEEL);   // layout backdrop
  tx(s, 7.065, 2.775, 4.89, 1.182, LOREM9, S.p11j);
  tx(s, 7.066, 1.26, 5.245, 1.178, MOCKUP, S.h32K);
  tx(s, 7.284, 5.279, 2.539, 0.903, LOREM2, S.p11);
  tx(s, 7.284, 4.902, 3.094, 0.404, TITTLE, S.h18K);
  tx(s, 10.495, 5.279, 2.539, 0.903, LOREM2, S.p11);
  tx(s, 10.495, 4.902, 3.094, 0.404, TITTLE, S.h18K);
  art(s, OUT.g10, 6.937, 5.099, 0.03, 0.03, STEEL);
  art(s, OUT.g11, 7.032, 5.159, 0.03, 0.03, STEEL);
  art(s, OUT.g12, 6.838, 4.982, 0.323, 0.323, STEEL);
  art(s, OUT.g13, 10.101, 5.053, 0.11, 0.139, STEEL);
  art(s, OUT.g14, 10.247, 4.985, 0.163, 0.09, STEEL);
  art(s, OUT.g15, 10.142, 4.98, 0.136, 0.145, STEEL);
  art(s, OUT.g16, 10.116, 5.213, 0.164, 0.092, STEEL);
  art(s, OUT.g17, 10.313, 5.095, 0.111, 0.143, STEEL);
  art(s, OUT.g18, 10.249, 5.163, 0.133, 0.144, STEEL);
  photo(s, 1.782, 1.382, 4.188, 6.118, 'CAC8C5');
}

// ---- 30  25K
function slide30(pres) {
  const s = page(pres, 30);
  box(s, 9.667, 0, 3.667, 7.5, STEEL);   // layout backdrop
  photo(s, 7.628, 0.948, 3.634, 5.613, 'CBC9C7');
  rule(s, 0.733, 2.653, 0.514, 0, { color: SAND, width: 4 });
  card(s, 1.241, 4.25, 1.264, 0.921, CARD);
  tx(s, 1.184, 4.281, 1.379, 0.64, '25K', S.h32S2c);
  tx(s, 1.237, 4.783, 1.272, 0.302, TITTLE, S.h9Kc);
  tx(s, 2.68, 4.398, 4.273, 0.625, lorem(113), S.p11);
  card(s, 1.241, 5.332, 1.264, 0.921, CARD);
  tx(s, 1.184, 5.363, 1.379, 0.64, '90%', S.h32S2c);
  tx(s, 1.237, 5.865, 1.272, 0.302, TITTLE, S.h9Kc);
  tx(s, 2.68, 5.479, 4.273, 0.625, lorem(113), S.p11);
  tx(s, 1.109, 2.622, 5.102, 0.903, LOREM5, S.p11j);
  tx(s, 1.109, 1.248, 5.245, 1.178, MOCKUP, S.h32K);
}

// ---- 31  Mockup Device Home Real Estate Here
function slide31(pres) {
  const s = page(pres, 31);
  box(s, 2.161, 0, 3.57, 7.5, STEEL);   // layout backdrop
  photo(s, 0.705, 2.066, 6.475, 3.754, 'C4C2C0');
  tx(s, 7.705, 2.082, 5.075, 1.178, MOCKUP, S.h32K);
  tx(s, 7.705, 3.522, 5.075, 0.903, lorem(179), S.p11);
  tx(s, 7.705, 4.514, 5.075, 0.903, lorem(179), S.p11);
}

// ---- 32  Gallery 2023
function slide32(pres) {
  const s = page(pres, 32);
  tx(s, 2.433, 5.178, 8.468, 1.447, 'Gallery 2023', S.h80Kc);
}

// ---- 33  Gallery
function slide33(pres) {
  const s = page(pres, 33);
  box(s, 0, 1.229, 5.317, 6.271, STEEL);   // layout backdrop
  tx(s, 0.672, 3.748, 4.646, 2.794, [GALLERY, '2023'].join(BR), { fontFace: HEAD, fontSize: 80, color: WHITE });
  tx(s, 5.774, 3.861, 2.539, 1.181, LOREM3, S.p11);
  tx(s, 5.774, 3.484, 3.094, 0.404, 'House Aspect', S.h18K);
}

// ---- 34  Gallery
function slide34(pres) {
  const s = page(pres, 34);
  box(s, 10.979, 0, 2.354, 6.333, STEEL);   // layout backdrop
  box(s, 0, 1.104, 3.889, 1.42, CARD);   // layout backdrop
  box(s, 10.979, 0, 2.354, 6.333, STEEL);
  card(s, 0, 1.104, 3.889, 1.42, CARD);
  tx(s, 0.608, 1.501, 2.952, 0.625, lorem(68), S.p11K);
  tx(s, 4.496, 4.253, 5.876, 2.794, [GALLERY, '2023'].join(BR), S.h80K);
}

// ---- 35  Gallery
function slide35(pres) {
  const s = page(pres, 35);
  box(s, 8.899, 4.404, 4.434, 3.096, STEEL);   // layout backdrop
  tx(s, 5.302, 0.421, 5.876, 2.794, [GALLERY, '2023'].join(BR), S.h80K);
  tx(s, 10.1, 5.312, 2.539, 1.181, LOREM3, S.p11W);
  tx(s, 10.1, 4.935, 3.094, 0.404, 'House Aspect', S.h18W);
}

// ---- 36  Gallery
function slide36(pres) {
  const s = page(pres, 36);
  box(s, 11.212, 0, 2.121, 6.141, STEEL);   // layout backdrop
  box(s, 0, 2.313, 3.785, 4.146, STEEL);   // layout backdrop
  tx(s, 5.195, 4.419, 5.876, 2.794, [GALLERY, '2023'].join(BR), S.h80K);
}

// ---- 37  Business Timeline
function slide37(pres) {
  const s = page(pres, 37);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Business Timeline', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g24, 1.226, 2.014, 0.678, 1.912, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 1.937, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 0.717, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 0.865, 2.162, 1.4, 1.402, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 0.916, 2.213, 1.299, 1.301, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 0.975, 2.272, 1.181, 1.183, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  rule(s, 1.584, 4.359, 10.182, 0, { color: RULE, width: 3 });
  s.addShape(SH.ellipse, { x: 1.439, y: 4.233, w: 0.251, h: 0.251, fill: { color: STEEL }, line: { color: 'F9F9F8', width: 2.25, transparency: 20 } });
  s.addShape(SH.ellipse, { x: 11.641, y: 4.233, w: 0.251, h: 0.251, fill: { color: STEEL3 }, line: { color: 'EAEDEF', width: 2.25, transparency: 20 } });
  s.addShape(SH.ellipse, { x: 3.48, y: 4.233, w: 0.251, h: 0.251, fill: { color: SAND2 }, line: { color: 'FBFBFB', width: 2.25, transparency: 20 } });
  s.addShape(SH.ellipse, { x: 5.52, y: 4.233, w: 0.251, h: 0.251, fill: { color: STEEL }, line: { color: 'FDFDFD', width: 2.25, transparency: 20 } });
  s.addShape(SH.ellipse, { x: 7.56, y: 4.233, w: 0.251, h: 0.251, fill: { color: STEEL }, line: { color: 'E6E9EB', width: 2.25, transparency: 20 } });
  s.addShape(SH.ellipse, { x: 9.6, y: 4.233, w: 0.251, h: 0.251, fill: { color: STEEL }, line: { color: 'E8EBED', width: 2.25, transparency: 20 } });
  art(s, OUT.g24, 3.266, 4.795, 0.678, 1.912, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 3.977, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 2.757, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 2.905, 5.157, 1.4, 1.402, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 2.958, 5.206, 1.299, 1.301, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 3.017, 5.265, 1.181, 1.183, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  art(s, OUT.g24, 5.306, 2.014, 0.678, 1.912, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 6.018, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 4.798, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 4.946, 2.162, 1.4, 1.402, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 4.998, 2.215, 1.299, 1.301, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 5.058, 2.274, 1.181, 1.183, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  art(s, OUT.g24, 7.347, 4.795, 0.678, 1.912, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 8.058, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 6.838, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 6.986, 5.157, 1.4, 1.402, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 7.039, 5.206, 1.299, 1.301, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 7.098, 5.265, 1.181, 1.183, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  art(s, OUT.g24, 9.387, 2.014, 0.678, 1.912, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 10.098, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 8.878, 2.1, 0.476, 1.526, MIST, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 9.026, 2.162, 1.4, 1.402, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 9.079, 2.215, 1.299, 1.301, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 9.138, 2.274, 1.181, 1.183, BG, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  art(s, OUT.g24, 11.427, 4.795, 0.678, 1.912, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g25, 12.138, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g26, 10.918, 5.095, 0.476, 1.526, TAUPE, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 3, offset: 1, angle: 45 } });
  art(s, OUT.g27, 11.066, 5.157, 1.4, 1.402, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, angle: 45 } });
  art(s, OUT.g27, 11.119, 5.206, 1.299, 1.301, CARD, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 5, offset: 1, angle: 45 } });
  art(s, OUT.g27, 11.178, 5.265, 1.181, 1.183, BG, { flipV: true, shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 6, offset: 3, angle: 45 } });
  tx(s, 2.207, 4.158, 0.757, 0.403, '2017', S.p1704Sc, { flipH: true, rectRadius: 0.202, fill: { color: BG } });
  tx(s, 4.247, 4.158, 0.757, 0.403, '2018', S.p1704Cc, { flipH: true, rectRadius: 0.202, fill: { color: BG } });
  tx(s, 6.287, 4.158, 0.757, 0.403, '2019', S.p1704Sc, { flipH: true, rectRadius: 0.202, fill: { color: BG } });
  tx(s, 8.299, 4.158, 0.813, 0.403, '2020', S.p1704Cc, { flipH: true, rectRadius: 0.202, fill: { color: BG } });
  tx(s, 10.368, 4.158, 0.757, 0.403, '2021', S.p1704Sc, { flipH: true, rectRadius: 0.202, fill: { color: BG } });
  rule(s, 1.565, 3.859, 0, 0.394, { color: SAND, width: 1, dashType: 'dash' });
  rule(s, 3.605, 4.464, 0, 0.394, { color: SAND2, width: 1, dashType: 'dash' });
  rule(s, 5.646, 3.859, 0, 0.394, { color: CARD, width: 1, dashType: 'dash' });
  rule(s, 7.686, 4.464, 0, 0.394, { color: STEEL, width: 1, dashType: 'dash' });
  rule(s, 9.726, 3.859, 0, 0.394, { color: STEEL2, width: 1, dashType: 'dash' });
  rule(s, 11.766, 4.464, 0, 0.394, { color: STEEL3, width: 1, dashType: 'dash' });
  s.addShape(SH.rect, { x: 5.474, y: 2.692, w: 0.347, h: 0.347, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g28, 5.518, 2.706, 0.261, 0.289, STEEL);
  s.addShape(SH.rect, { x: 3.434, y: 5.682, w: 0.347, h: 0.347, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g29, 3.448, 5.711, 0.318, 0.289, CARD);
  s.addShape(SH.rect, { x: 11.595, y: 5.682, w: 0.347, h: 0.347, fill: { color: STEEL3, transparency: 100 }, line: NOLINE });
  art(s, OUT.g30, 11.609, 5.726, 0.318, 0.261, CARD);
  s.addShape(SH.rect, { x: 1.391, y: 2.689, w: 0.347, h: 0.347, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g31, 1.42, 2.733, 0.289, 0.261, STEEL);
  s.addShape(SH.rect, { x: 9.555, y: 2.692, w: 0.347, h: 0.347, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g32, 9.569, 2.735, 0.318, 0.261, STEEL);
  s.addShape(SH.rect, { x: 7.514, y: 5.682, w: 0.347, h: 0.347, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g33, 7.543, 5.726, 0.289, 0.261, CARD);
  tx(s, 2.521, 3.342, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 2.521, 2.965, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 6.605, 3.342, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 6.605, 2.965, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 8.65, 5.131, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 8.65, 4.753, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 10.688, 3.342, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 10.688, 2.965, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 0.479, 5.131, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 0.479, 4.753, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 4.564, 5.131, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 4.564, 4.753, 2.165, 0.404, TITTLE, S.h18Kc);
}

// ---- 38  Step Infographic
function slide38(pres) {
  const s = page(pres, 38);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Step Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g34, 6.117, 2.491, 1.088, 1.085, DEEP);
  art(s, OUT.g35, 6.117, 3.824, 1.1, 1.1, CHAR);
  art(s, OUT.g34, 6.117, 5.15, 1.088, 1.088, DEEP);
  box(s, 6.129, 6.482, 1.088, 1.018, CHAR);
  art(s, OUT.g36, 6.117, 1.403, 2.977, 1.088, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 16, offset: 9, angle: 45 } });
  art(s, OUT.g37, 4.239, 2.735, 2.979, 1.088, CARD, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 15, offset: 9, angle: 135 } });
  art(s, OUT.g36, 6.117, 4.061, 2.977, 1.088, STEEL, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 16, offset: 9, angle: 45 } });
  art(s, OUT.g37, 4.239, 5.398, 2.979, 1.085, CARD, { shadow: { type: 'outer', color: INK, opacity: 0.4, blur: 15, offset: 9, angle: 135 } });
  s.addShape(SH.rect, { x: 7.459, y: 1.8, w: 0.294, h: 0.294, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g38, 7.483, 1.824, 0.245, 0.233, CARD);
  s.addShape(SH.rect, { x: 5.581, y: 3.132, w: 0.294, h: 0.294, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g39, 5.606, 3.157, 0.245, 0.233, STEEL);
  s.addShape(SH.rect, { x: 5.581, y: 5.793, w: 0.294, h: 0.294, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g40, 5.618, 5.83, 0.221, 0.221, STEEL);
  s.addShape(SH.rect, { x: 7.437, y: 4.436, w: 0.339, h: 0.339, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g41, 7.445, 4.443, 0.322, 0.307, CARD);
  tx(s, 9.324, 1.817, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 9.324, 1.439, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 9.324, 4.472, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 9.324, 4.095, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 1.845, 3.146, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.845, 2.769, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 1.845, 5.802, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.845, 5.424, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 39  Process Infographic
function slide39(pres) {
  const s = page(pres, 39);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Process Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g42, 1.248, 1.775, 3.539, 2.73, RULE2);
  art(s, OUT.g43, 5.875, 3.461, 1.404, 0.465, RULE2);
  art(s, OUT.g42, 8.539, 1.774, 3.539, 2.73, RULE2, { flipH: true });
  art(s, OUT.g44, 1.771, 4.818, 1.746, 2.375, WHITE, { line: { color: RULE2 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  s.addShape(SH.ellipse, { x: 2.271, y: 6.298, w: 0.722, h: 0.722, fill: { color: STEEL }, line: { color: SAND } });
  art(s, OUT.g45, 2.423, 6.456, 0.417, 0.403, WHITE, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g46, 4.372, 4.818, 1.76, 2.375, WHITE, { line: { color: RULE2 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  s.addShape(SH.ellipse, { x: 4.885, y: 6.298, w: 0.722, h: 0.722, fill: { color: CARD }, line: { color: SAND2 } });
  art(s, OUT.g45, 5.038, 6.456, 0.415, 0.403, WHITE, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g47, 7.01, 4.818, 1.748, 2.375, WHITE, { line: { color: RULE2 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  s.addShape(SH.ellipse, { x: 7.523, y: 6.298, w: 0.722, h: 0.722, fill: { color: STEEL }, line: { color: CARD } });
  art(s, OUT.g45, 7.688, 6.449, 0.403, 0.403, WHITE, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g44, 9.637, 4.818, 1.746, 2.375, WHITE, { line: { color: RULE2 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  s.addShape(SH.ellipse, { x: 10.15, y: 6.298, w: 0.722, h: 0.722, fill: { color: CARD }, line: NOLINE });
  art(s, OUT.g45, 10.308, 6.449, 0.403, 0.403, WHITE, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g48, 1.709, 1.91, 1.906, 1.919, CARD, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g49, 1.758, 1.972, 1.808, 1.796, STEEL, { line: { color: SAND } });
  art(s, OUT.g48, 4.275, 1.91, 1.917, 1.919, STEEL, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g50, 4.348, 1.972, 1.796, 1.796, CARD, { line: { color: SAND2 } });
  art(s, OUT.g48, 6.888, 1.91, 1.919, 1.919, CARD, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g51, 6.962, 1.972, 1.796, 1.796, STEEL, { line: { color: CARD } });
  art(s, OUT.g48, 9.552, 1.91, 1.905, 1.919, STEEL, { line: { color: RULE3 }, shadow: { type: 'outer', color: INK, opacity: 0.15, blur: 3, offset: 1, angle: 90 } });
  art(s, OUT.g52, 9.6, 1.972, 1.796, 1.796, CARD, { line: { color: CARD } });
  s.addShape(SH.rect, { x: 7.671, y: 2.691, w: 0.373, h: 0.373, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g53, 7.702, 2.722, 0.311, 0.311, STEEL);
  s.addShape(SH.rect, { x: 2.46, y: 2.691, w: 0.373, h: 0.373, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g31, 2.492, 2.738, 0.311, 0.28, STEEL);
  s.addShape(SH.rect, { x: 5.059, y: 2.677, w: 0.373, h: 0.373, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g32, 5.075, 2.724, 0.342, 0.28, CARD);
  s.addShape(SH.rect, { x: 10.311, y: 2.691, w: 0.373, h: 0.373, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g33, 10.342, 2.738, 0.311, 0.28, CARD);
  art(s, OUT.g54, -1.392, 4.396, 16.118, 0.439, 'F2F2F2');
  art(s, OUT.g55, 3.897, 4.586, 0.06, 0.062, WHITE);
  art(s, OUT.g56, 6.559, 4.586, 0.048, 0.062, WHITE);
  art(s, OUT.g55, 9.209, 4.586, 0.06, 0.062, WHITE);
  tx(s, 1.795, 5.414, 1.746, 0.625, lorem(36), S.p11c);
  tx(s, 1.795, 5.037, 1.746, 0.404, TITTLE, S.h18Kc);
  tx(s, 4.401, 5.414, 1.746, 0.625, lorem(36), S.p11c);
  tx(s, 4.401, 5.037, 1.746, 0.404, TITTLE, S.h18Kc);
  tx(s, 7.006, 5.414, 1.746, 0.625, lorem(36), S.p11c);
  tx(s, 7.006, 5.037, 1.746, 0.404, TITTLE, S.h18Kc);
  tx(s, 9.611, 5.414, 1.746, 0.625, lorem(36), S.p11c);
  tx(s, 9.611, 5.037, 1.746, 0.404, TITTLE, S.h18Kc);
}

// ---- 40  Lamp Infographic
function slide40(pres) {
  const s = page(pres, 40);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Lamp Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g57, 0.842, 2.122, 1.834, 2.093, STEEL);
  art(s, OUT.g58, 0.676, 1.944, 2.223, 4.051, STEEL, { line: { color: STEEL, width: 1 } });
  art(s, OUT.g57, 3.764, 2.122, 1.836, 2.093, CARD);
  art(s, OUT.g59, 3.559, 1.944, 2.223, 4.051, CARD, { line: { color: CARD, width: 1 } });
  art(s, OUT.g57, 6.667, 2.122, 1.835, 2.093, STEEL);
  art(s, OUT.g58, 6.475, 1.944, 2.223, 4.051, STEEL, { line: { color: STEEL, width: 1 } });
  art(s, OUT.g57, 9.566, 2.122, 1.836, 2.093, CARD);
  art(s, OUT.g60, 9.387, 1.944, 2.223, 4.051, CARD, { line: { color: CARD, width: 1 } });
  art(s, OUT.g29, 7.468, 3.052, 0.258, 0.234, CARD);
  art(s, OUT.g32, 10.355, 3.064, 0.258, 0.211, STEEL);
  art(s, OUT.g33, 4.565, 3.064, 0.234, 0.211, STEEL);
  art(s, OUT.g61, 1.643, 3.052, 0.258, 0.234, CARD);
  tx(s, 1.901, 6.078, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 1.901, 5.701, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 4.818, 6.078, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 4.818, 5.701, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 7.748, 6.078, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 7.748, 5.701, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 10.601, 6.078, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 10.601, 5.701, 2.165, 0.404, TITTLE, S.h18K);
}

// ---- 41  Pencil Infographic
function slide41(pres) {
  const s = page(pres, 41);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Pencil Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g62, 6.02, 6.775, 1.092, 0.446, RULE);
  art(s, OUT.g63, 5.982, 6.481, 1.167, 0.248, INK);
  art(s, OUT.g64, 6.394, 1.315, 0.346, 0.217, INK);
  art(s, OUT.g65, 6.02, 1.57, 1.092, 0.546, RULE);
  box(s, 6.02, 3.278, 1.092, 1.004, CARD);
  box(s, 6.02, 5.399, 1.092, 1.037, CARD);
  box(s, 6.02, 4.345, 1.092, 1.004, STEEL);
  box(s, 6.02, 2.191, 1.092, 1.038, STEEL);
  art(s, OUT.g66, 6.559, 1.318, 0.181, 0.214, INK);
  bent(s, 7.325, 1.908, 1.227, 0.803, { color: STEEL, width: 1, dashType: 'dash' }, { flipV: true });
  bent(s, 4.799, 3.076, 1.027, 0.703, { color: CARD, width: 1, dashType: 'dash' });
  bent(s, 7.325, 4.997, 1.227, 0.953, { color: CARD, width: 1, dashType: 'dash' }, { flipV: true });
  bent(s, 4.799, 4.891, 1.027, 0.783, { color: STEEL, width: 1, dashType: 'dash' }, { flipV: true });
  s.addShape(SH.rect, { x: 6.416, y: 5.768, w: 0.3, h: 0.3, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g29, 6.429, 5.793, 0.275, 0.25, STEEL);
  s.addShape(SH.rect, { x: 6.416, y: 3.63, w: 0.3, h: 0.3, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g33, 6.441, 3.667, 0.25, 0.225, STEEL);
  s.addShape(SH.rect, { x: 6.416, y: 4.697, w: 0.3, h: 0.3, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g30, 6.429, 4.735, 0.275, 0.225, WHITE);
  s.addShape(SH.rect, { x: 6.427, y: 2.572, w: 0.278, h: 0.278, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g28, 6.462, 2.583, 0.208, 0.231, WHITE);
  tx(s, 8.684, 1.785, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.684, 1.407, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 8.684, 4.873, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.684, 4.496, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 2.491, 2.952, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 2.491, 2.575, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 2.485, 5.55, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 2.485, 5.173, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 42  Layered Infographic
function slide42(pres) {
  const s = page(pres, 42);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Layered Infographic', S.h32K_2);
  tx(s, 0.834, 1.191, 9.646, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g67, 4.207, 4.373, 2.46, 1.169, STEEL);
  box(s, 4.207, 3.212, 1.16, 1.16, DEEP);
  art(s, OUT.g68, 5.496, 1.913, 1.171, 2.46, CARD);
  box(s, 6.667, 1.913, 1.16, 1.171, CHAR);
  art(s, OUT.g69, 6.667, 4.373, 1.16, 2.46, CARD);
  box(s, 5.496, 5.672, 1.171, 1.16, CHAR);
  art(s, OUT.g70, 6.667, 3.212, 2.46, 1.16, STEEL);
  art(s, OUT.g71, 7.966, 4.373, 1.171, 1.169, DEEP);
  s.addShape(SH.rect, { x: 8.411, y: 4.829, w: 0.264, h: 0.264, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g72, 8.422, 4.843, 0.242, 0.218, WHITE);
  s.addShape(SH.rect, { x: 5.927, y: 6.132, w: 0.264, h: 0.264, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g28, 5.96, 6.143, 0.198, 0.22, WHITE);
  s.addShape(SH.rect, { x: 7.126, y: 2.363, w: 0.264, h: 0.264, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g31, 7.148, 2.396, 0.22, 0.198, WHITE);
  s.addShape(SH.rect, { x: 4.659, y: 3.649, w: 0.264, h: 0.264, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g33, 4.681, 3.682, 0.22, 0.198, WHITE);
  tx(s, 8.055, 2.374, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.055, 1.997, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 9.387, 4.833, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 9.387, 4.456, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 1.748, 3.669, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.748, 3.291, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 3.08, 6.129, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 3.08, 5.751, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 43  Option Infographic
function slide43(pres) {
  const s = page(pres, 43);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Option Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g73, 8.011, 2.683, 0.876, 1.62, null, { line: { color: 'EEEEEE', width: 1.65, transparency: 50 } });
  art(s, OUT.g74, 4.428, 2.683, 0.876, 1.62, null, { line: { color: 'EEEEEE', width: 1.65, transparency: 50 } });
  art(s, OUT.g75, 4.428, 4.303, 0.876, 1.62, null, { line: { color: 'EEEEEE', width: 1.65, transparency: 50 } });
  art(s, OUT.g76, 8.011, 4.303, 0.876, 1.62, null, { line: { color: 'EEEEEE', width: 1.65, transparency: 50 } });
  art(s, OUT.g77, 4.969, 2.6, 3.438, 3.438, null, { fill: { color: RULE, transparency: 50 } });
  art(s, OUT.g78, 5.125, 2.771, 3.064, 3.064, WHITE);
  art(s, OUT.g77, 5.177, 2.838, 2.931, 2.931, null, { fill: { color: 'BFBFBF', transparency: 50 } });
  tx(s, 5.205, 3.983, 2.877, 0.64, 'Best Option', { fontFace: HEAD, fontSize: 32, color: INK, align: 'center' });
  art(s, OUT.g79, 8.85, 3.819, 1.001, 1.001, STEEL);
  art(s, OUT.g79, 8.85, 2.198, 1.001, 1.001, CARD, { rotate: 315 });
  art(s, OUT.g79, 8.849, 5.439, 1.001, 1.001, CARD, { rotate: 315 });
  art(s, OUT.g79, 9.691, 2.539, 0.319, 0.319, CARD, { rotate: 315, line: { color: WHITE, width: 7 } });
  art(s, OUT.g79, 9.691, 4.16, 0.319, 0.319, STEEL, { rotate: 315, line: { color: WHITE, width: 7 } });
  art(s, OUT.g79, 9.691, 5.78, 0.319, 0.319, CARD, { rotate: 315, line: { color: WHITE, width: 7 } });
  s.addShape(SH.rect, { x: 9.175, y: 5.764, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g80, 9.204, 5.793, 0.292, 0.292, STEEL);
  s.addShape(SH.rect, { x: 9.175, y: 4.144, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g81, 9.197, 4.195, 0.306, 0.267, CARD);
  s.addShape(SH.rect, { x: 9.175, y: 2.523, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g82, 9.205, 2.567, 0.291, 0.275, STEEL);
  s.addShape(SH.rect, { x: 4.428, y: 4.319, w: 0.316, h: 0.005, line: { color: 'EEEEEE', width: 1.65, transparency: 50 } });
  art(s, OUT.g79, 3.464, 2.198, 1.001, 1.001, STEEL, { rotate: 315 });
  art(s, OUT.g79, 3.464, 3.819, 1.001, 1.001, CARD, { rotate: 315 });
  art(s, OUT.g79, 3.464, 5.439, 1.001, 1.001, STEEL, { rotate: 315 });
  art(s, OUT.g79, 3.323, 2.539, 0.319, 0.319, STEEL, { rotate: 315, line: { color: WHITE, width: 7 } });
  art(s, OUT.g79, 3.323, 4.16, 0.319, 0.319, CARD, { rotate: 315, line: { color: WHITE, width: 7 } });
  art(s, OUT.g79, 3.323, 5.78, 0.319, 0.319, STEEL, { rotate: 315, line: { color: WHITE, width: 7 } });
  s.addShape(SH.rect, { x: 3.79, y: 5.764, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g83, 3.819, 5.793, 0.292, 0.292, CARD);
  s.addShape(SH.rect, { x: 3.79, y: 4.144, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g84, 3.834, 4.188, 0.263, 0.263, STEEL);
  s.addShape(SH.rect, { x: 3.79, y: 2.523, w: 0.35, h: 0.35, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g85, 3.834, 2.553, 0.292, 0.307, CARD);
  tx(s, 10.318, 2.559, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 10.318, 2.181, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 10.318, 4.187, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 10.318, 3.81, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 10.318, 5.815, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 10.318, 5.438, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 0.796, 2.559, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 0.796, 2.181, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 0.796, 4.187, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 0.796, 3.81, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 0.796, 5.815, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 0.796, 5.438, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 44  Puzzle Infographic
function slide44(pres) {
  const s = page(pres, 44);
  rule(s, 4.149, 3.26, 0.722, 0, { color: CARD, width: 2.25, dashType: 'sysDash' }, { flipH: true });
  s.addShape(SH.ellipse, { x: 3.986, y: 3.167, w: 0.21, h: 0.204, fill: { color: CARD }, line: { color: CARD } });
  rule(s, 8.489, 5.331, 1.026, 0, { color: CARD, width: 2.25, dashType: 'sysDash' }, { flipH: true });
  s.addShape(SH.ellipse, { x: 9.41, y: 5.227, w: 0.21, h: 0.204, fill: { color: CARD }, line: { color: CARD } });
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Puzzle Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  rule(s, 4.061, 5.924, 1.538, 0, { color: STEEL, width: 2.25, dashType: 'sysDash' }, { flipH: true });
  s.addShape(SH.ellipse, { x: 3.956, y: 5.822, w: 0.21, h: 0.204, fill: { color: STEEL }, line: { color: STEEL } });
  rule(s, 7.684, 2.631, 0.916, 0.009, { color: STEEL, width: 2.25, dashType: 'sysDash' }, { flipH: true, flipV: true });
  s.addShape(SH.ellipse, { x: 8.6, y: 2.538, w: 0.21, h: 0.204, fill: { color: STEEL }, line: { color: STEEL } });
  art(s, OUT.g86, 7.514, 3.109, 2, 2.374, CARD, { line: { color: RULE, width: 5.75 } });
  art(s, OUT.g87, 5.522, 5.124, 2.379, 1.997, STEEL, { line: { color: RULE, width: 5.75 } });
  art(s, OUT.g88, 5.901, 1.459, 2, 1.992, STEEL, { line: { color: RULE, width: 5.75 } });
  art(s, OUT.g89, 3.848, 3.076, 2.053, 2.048, CARD, { line: { color: RULE, width: 5.75 } });
  s.addShape(SH.rect, { x: 6.549, y: 5.96, w: 0.326, h: 0.326, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g53, 6.576, 5.987, 0.272, 0.272, CARD);
  s.addShape(SH.rect, { x: 8.351, y: 4.133, w: 0.326, h: 0.326, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g28, 8.392, 4.147, 0.245, 0.272, STEEL);
  s.addShape(SH.rect, { x: 4.711, y: 3.937, w: 0.326, h: 0.326, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g29, 4.725, 3.964, 0.299, 0.272, STEEL);
  s.addShape(SH.rect, { x: 6.738, y: 2.292, w: 0.326, h: 0.326, fill: { color: WHITE, transparency: 100 }, line: NOLINE });
  art(s, OUT.g31, 6.765, 2.333, 0.272, 0.245, CARD);
  tx(s, 8.996, 2.504, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.996, 2.126, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 9.819, 5.211, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 9.819, 4.833, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 1.635, 3.136, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.635, 2.759, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 1.614, 5.804, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.614, 5.427, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 45  Chart Infographic
function slide45(pres) {
  const s = page(pres, 45);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Chart Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  donut(s, 0.431, 2.319, 2.646, 2.51, 50, STEEL, RULE);
  tx(s, 0.854, 3.205, 1.799, 0.739, '50%', S.h32Kc);
  donut(s, 3.804, 2.319, 2.646, 2.51, 75, CARD, 'D5D2D0');
  tx(s, 4.228, 3.205, 1.799, 0.739, '75%', S.h32Kc);
  donut(s, 6.883, 2.319, 2.646, 2.51, 25, STEEL, 'D5D2D0');
  tx(s, 7.306, 3.205, 1.799, 0.739, '25%', S.h32Kc);
  donut(s, 10.256, 2.319, 2.646, 2.51, 80, CARD, 'D5D2D0');
  tx(s, 10.679, 3.205, 1.799, 0.739, '80%', S.h32Kc);
  tx(s, 0.672, 5.684, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 0.672, 5.307, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 4.045, 5.684, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 4.045, 5.307, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 7.124, 5.684, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 7.124, 5.307, 2.165, 0.404, TITTLE, S.h18Kc);
  tx(s, 10.497, 5.684, 2.165, 0.625, lorem(52), S.p11c);
  tx(s, 10.497, 5.307, 2.165, 0.404, TITTLE, S.h18Kc);
}

// ---- 46  Pyramid Infographic
function slide46(pres) {
  const s = page(pres, 46);
  tx(s, 0.768, 0.51, 10.228, 0.703, 'Pyramid Infographic', S.h32K_2);
  tx(s, 0.786, 1.191, 9.726, 0.324, SUBHEAD, S.p16);
  art(s, OUT.g90, 6.057, 2.192, 0.895, 0.773, STEEL);
  art(s, OUT.g91, 5.757, 2.966, 1.789, 0.775, CARD);
  art(s, OUT.g92, 5.162, 3.741, 2.686, 0.773, STEEL);
  art(s, OUT.g93, 4.861, 4.514, 3.582, 0.775, CARD);
  art(s, OUT.g94, 4.266, 5.289, 4.477, 0.777, STEEL);
  art(s, OUT.g95, 3.966, 6.066, 5.372, 0.773, CARD);
  s.addShape(SH.ellipse, { x: 7.702, y: 3.186, w: 0.151, h: 0.151, fill: { color: CARD }, line: NOLINE });
  s.addShape(SH.ellipse, { x: 5.756, y: 2.406, w: 0.151, h: 0.151, fill: { color: STEEL }, flipH: true, line: NOLINE });
  s.addShape(SH.ellipse, { x: 8.623, y: 4.756, w: 0.151, h: 0.151, fill: { color: CARD }, line: NOLINE });
  s.addShape(SH.ellipse, { x: 4.86, y: 4.006, w: 0.151, h: 0.151, fill: { color: STEEL }, flipH: true, line: NOLINE });
  s.addShape(SH.ellipse, { x: 9.505, y: 6.325, w: 0.151, h: 0.151, fill: { color: CARD }, line: NOLINE });
  s.addShape(SH.ellipse, { x: 3.817, y: 5.578, w: 0.151, h: 0.151, fill: { color: STEEL }, flipH: true, line: NOLINE });
  tx(s, 8.008, 3.138, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.008, 2.76, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 8.921, 4.701, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 8.921, 4.323, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 9.824, 6.277, 2.165, 0.625, lorem(52), S.p11);
  tx(s, 9.824, 5.9, 2.165, 0.404, TITTLE, S.h18K);
  tx(s, 3.43, 2.357, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 3.43, 1.98, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 2.561, 3.954, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 2.561, 3.577, 2.165, 0.404, TITTLE, S.h18Kr);
  tx(s, 1.504, 5.541, 2.165, 0.625, lorem(52), S.p11r);
  tx(s, 1.504, 5.163, 2.165, 0.404, TITTLE, S.h18Kr);
}

// ---- 47  Our Real Estate Customer Here
function slide47(pres) {
  const s = page(pres, 47);
  box(s, 6.258, 2.903, 5.581, 1.516, STEEL);   // layout backdrop
  box(s, 6.258, 5.065, 5.581, 1.516, CARD);   // layout backdrop
  tx(s, 7.661, 3.546, 3.872, 0.625, LOREM3, S.p11W);
  tx(s, 7.661, 3.152, 2.421, 0.413, 'Customer One Here', S.h14W);
  tx(s, 7.661, 5.705, 3.872, 0.625, LOREM3, S.p11K);
  tx(s, 7.661, 5.311, 2.421, 0.413, 'Customer Two Here', S.h14K);
  tx(s, 1.156, 0.922, 4.432, 1.178, 'Our Real Estate Customer Here', S.h32K);
  tx(s, 7.186, 1.059, 4.653, 0.903, lorem(179), S.p11);
}

// ---- 48  Our Real Estate Customer Here
function slide48(pres) {
  const s = page(pres, 48);
  box(s, 6.922, 1.215, 5.341, 1.388, CARD);   // layout backdrop
  box(s, 6.922, 3.056, 5.341, 1.388, STEEL);   // layout backdrop
  box(s, 6.922, 4.897, 5.341, 1.388, CARD);   // layout backdrop
  tx(s, 8.201, 1.79, 3.872, 0.625, LOREM3, S.p11);
  tx(s, 8.201, 1.396, 2.421, 0.413, 'Customer One Here', S.h14K);
  tx(s, 8.201, 3.631, 3.872, 0.625, LOREM3, S.p11W);
  tx(s, 8.201, 3.237, 2.421, 0.413, 'Customer Two Here', S.h14W);
  tx(s, 8.201, 5.472, 3.872, 0.625, LOREM3, S.p11);
  tx(s, 8.201, 5.078, 2.421, 0.413, 'Customer Three Here', S.h14K);
  tx(s, 1.071, 2.009, 4.432, 1.178, 'Our Real Estate Customer Here', S.h32K);
  tx(s, 1.071, 3.595, 4.991, 0.903, lorem(179), S.p11);
  tx(s, 1.071, 4.588, 4.991, 0.903, lorem(179), S.p11);
}

// ---- 49  Contact Information
function slide49(pres) {
  const s = page(pres, 49);
  tx(s, 1.092, 1.473, 4.432, 1.178, 'Contact Information', S.h32K);
  tx(s, 1.092, 2.784, 3.096, 0.625, lorem(76), S.p11);
  box(s, 0, 3.75, 6.208, 3.75, STEEL);
  tx(s, 1.694, 4.84, 3.83, 0.372, '12 Your Street Name, Your City Name 1234', S.p12W);
  art(s, OUT.g96, 1.375, 4.666, 0.208, 0.331, WHITE, { flipH: true });
  tx(s, 1.725, 4.489, 2.462, 0.471, 'ADDRESS', S.h16W);
  tx(s, 1.694, 5.633, 3.83, 0.372, '+123 456 7890 / 067 837 736 7263', S.p12W);
  tx(s, 1.725, 5.282, 2.462, 0.471, 'PHONE NUMBER', S.h16W);
  tx(s, 1.694, 6.425, 3.83, 0.372, 'www.yourcompanysite.com', S.p12W);
  tx(s, 1.725, 6.075, 2.462, 0.471, 'WEBSITE', S.h16W);
  art(s, OUT.g97, 1.375, 5.493, 0.249, 0.249, WHITE);
  art(s, OUT.g98, 1.511, 5.522, 0.083, 0.082, WHITE);
  art(s, OUT.g99, 1.514, 5.467, 0.136, 0.133, WHITE);
  art(s, OUT.g100, 1.349, 6.259, 0.29, 0.287, WHITE);
}

// ---- 50  Contact Information
function slide50(pres) {
  const s = page(pres, 50);
  tx(s, 8.902, 2.169, 4.432, 1.178, 'Contact Information', S.h32K);
  tx(s, 8.902, 3.395, 3.81, 0.903, lorem(155), S.p11);
  box(s, 1.376, 4.833, 11.957, 1.946, STEEL);
  tx(s, 2.39, 5.796, 3.83, 0.372, '12 Your Street Name, Your City Name 1234', S.p12W);
  art(s, OUT.g96, 2.071, 5.622, 0.208, 0.331, WHITE, { flipH: true });
  tx(s, 2.421, 5.445, 2.462, 0.471, 'ADDRESS', S.h16W);
  tx(s, 6.535, 5.796, 3.83, 0.372, '+123 456 7890 / 067 837 736 7263', S.p12W);
  tx(s, 6.567, 5.445, 2.462, 0.471, 'PHONE NUMBER', S.h16W);
  art(s, OUT.g97, 6.216, 5.656, 0.249, 0.249, WHITE);
  art(s, OUT.g98, 6.352, 5.685, 0.083, 0.082, WHITE);
  art(s, OUT.g99, 6.355, 5.63, 0.136, 0.133, WHITE);
  tx(s, 10.129, 5.796, 3.83, 0.372, 'www.yourcompanysite.com', S.p12W);
  tx(s, 10.16, 5.445, 2.462, 0.471, 'WEBSITE', S.h16W);
  art(s, OUT.g100, 9.784, 5.63, 0.29, 0.287, WHITE);
}

// ------------------------------------------------------------------ build
function build() {
  const pres = new PptxGenJS();
  SH = pres.ShapeType;
  CT = pres.ChartType;
  pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'WIDE';
  pres.title = 'House Real Estate Presentation';
  [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
   slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
   slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
   slide41, slide42, slide43, slide44, slide45, slide46, slide47, slide48, slide49, slide50]
    .forEach((fn) => fn(pres));
  return pres.writeFile({ fileName: path.join(__dirname, '0cc535a5-421c-443f-969b-10ddf835933e_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
