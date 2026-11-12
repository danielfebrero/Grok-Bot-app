/*
 * RodeiX - Virtual Reality & Metaverse presentation template (45 slides, 13.33 x 7.5 in).
 *
 * Everything is plain pptxgenjs: the deck's colours, fonts, geometry and copy live in the
 * lookup tables below and in one builder function per slide.
 *
 *   F .... flat fill swatches (the source deck's gradients are reduced to a mid colour)
 *   T .... run/paragraph text styles
 *   A .... freeform outlines for the decorative wave art, as normalised 0..1 polylines
 *   L1.. . repeated body copy
 *
 * Raster photos in the original are redrawn here as labelled placeholder plates.
 *
 * Run:  node <this file>          ->  writes the .pptx next to the script
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---- palette -------------------------------------------------------------
const BG = '270060';      // deep indigo page colour
const PINK = 'F9337A';    // brand magenta
const VIOLET = 'C204D6';
const WHITE = 'FFFFFF';
const GOLD = 'FFC000';    // rating stars
const BLACK = '000000';
const NOLINE = { type: 'none' };
const SHADOW = { type: 'outer', color: BLACK, opacity: 0.2, blur: 12, offset: 3, angle: 90 };


// ---- fill swatches -------------------------------------------------------
const F = {
  c1:{color:BG}, c2:{color:PINK,transparency:50}, c3:{color:'7D156B',transparency:16},
  c4:{color:'9F1D6F',transparency:32}, c5:{color:'741369',transparency:41},
  c6:{color:'6E1169',transparency:46}, c7:{color:'931A6D',transparency:20}, c8:{color:BG,transparency:92},
  c9:{color:PINK,transparency:76}, c10:{color:PINK}, c11:{color:PINK,transparency:58},
  c12:{color:'86176C',transparency:8}, c13:{color:BG,transparency:24}, c14:{color:WHITE},
  c15:{color:BG,transparency:50}, c16:{color:BG,transparency:22}, c17:{color:'B22271',transparency:34},
  c18:{color:GOLD}, c19:{color:'901A6D'}, c20:{color:'44546A'}, c21:{color:'4472C4'}, c22:{color:'A5A5A5'},
  c23:{color:PINK,transparency:36}, c24:{color:'44546A',transparency:80}, c25:{color:WHITE,transparency:52},
  c26:{color:PINK,transparency:22}, c27:{color:PINK,transparency:26}, c28:{color:BLACK,transparency:80},
  c29:{color:PINK,transparency:66}, c30:{color:BG,transparency:34}
};

// ---- text styles ---------------------------------------------------------
const T = {
  a18: {ff:'Archivo',sz:18,c:WHITE},
  r14CN: {ff:'Orbitron',sz:14,c:WHITE,al:'center',nw:1},
  r96bN: {ff:'Orbitron',sz:96,b:1,c:WHITE,nw:1},
  r11bpN: {ff:'Orbitron',sz:11,b:1,c:PINK,nw:1},
  a44b: {ff:'Archivo',sz:44,b:1,c:WHITE},
  o10JS: {ff:'Open Sans',sz:10,c:WHITE,al:'justify',ls:1.5},
  a14bC: {ff:'Archivo',sz:14,b:1,c:WHITE,al:'center'},
  a40b: {ff:'Archivo',sz:40,b:1,c:WHITE},
  a12b: {ff:'Archivo',sz:12,b:1,c:WHITE},
  a28b: {ff:'Archivo',sz:28,b:1,c:WHITE},
  a29b: {ff:'Archivo',sz:29,b:1,c:WHITE},
  o10: {ff:'Open Sans',sz:10,c:WHITE},
  a10_5bR: {ff:'Archivo',sz:10.5,b:1,c:WHITE,al:'right'},
  a32b: {ff:'Archivo',sz:32,b:1,c:WHITE},
  a14bCM: {ff:'Archivo',sz:14,b:1,c:WHITE,al:'center',v:'middle'},
  a11b: {ff:'Archivo',sz:11,b:1,c:WHITE},
  a11bip: {ff:'Archivo',sz:11,b:1,i:1,c:PINK},
  a40bp: {ff:'Archivo',sz:40,b:1,c:PINK},
  a32bp: {ff:'Archivo',sz:32,b:1,c:PINK},
  a16bC: {ff:'Archivo',sz:16,b:1,c:WHITE,al:'center'},
  r66bN: {ff:'Orbitron',sz:66,b:1,c:WHITE,nw:1},
  o10CS: {ff:'Open Sans',sz:10,c:WHITE,al:'center',ls:1.5},
  r11bCpN: {ff:'Orbitron',sz:11,b:1,c:PINK,al:'center',nw:1},
  a24bC: {ff:'Archivo',sz:24,b:1,c:WHITE,al:'center'},
  a24bCp: {ff:'Archivo',sz:24,b:1,c:PINK,al:'center'},
  a14b: {ff:'Archivo',sz:14,b:1,c:WHITE},
  a24b: {ff:'Archivo',sz:24,b:1,c:WHITE},
  a16bR: {ff:'Archivo',sz:16,b:1,c:WHITE,al:'right'},
  o10RS: {ff:'Open Sans',sz:10,c:WHITE,al:'right',ls:1.5},
  a16b: {ff:'Archivo',sz:16,b:1,c:WHITE},
  o10S: {ff:'Open Sans',sz:10,c:WHITE,ls:1.5},
  r11bRpN: {ff:'Orbitron',sz:11,b:1,c:PINK,al:'right',nw:1},
  a32bR: {ff:'Archivo',sz:32,b:1,c:WHITE,al:'right'},
  r11bN: {ff:'Orbitron',sz:11,b:1,c:WHITE,nw:1},
  a16bCN: {ff:'Archivo',sz:16,b:1,c:WHITE,al:'center',nw:1},
  a12bCM: {ff:'Archivo',sz:12,b:1,c:WHITE,al:'center',v:'middle'},
  a36b: {ff:'Archivo',sz:36,b:1,c:WHITE},
  a16bCM: {ff:'Archivo',sz:16,b:1,c:WHITE,al:'center',v:'middle'},
  a14bN: {ff:'Archivo',sz:14,b:1,c:WHITE,nw:1},
  a36bC: {ff:'Archivo',sz:36,b:1,c:WHITE,al:'center'},
  w14bN: {ff:'Work Sans',sz:14,b:1,c:WHITE,nw:1},
  a28bCM: {ff:'Archivo',sz:28,b:1,c:WHITE,al:'center',v:'middle'},
  a18b: {ff:'Archivo',sz:18,b:1,c:WHITE},
  a60b: {ff:'Archivo',sz:60,b:1,c:WHITE},
  a199bC: {ff:'Archivo',sz:199,b:1,c:WHITE,al:'center'},
  a60bR: {ff:'Archivo',sz:60,b:1,c:WHITE,al:'right'},
  o10iJS: {ff:'Open Sans',sz:10,i:1,c:WHITE,al:'justify',ls:1.5},
  a28bC: {ff:'Archivo',sz:28,b:1,c:WHITE,al:'center'},
  a18bC: {ff:'Archivo',sz:18,b:1,c:WHITE,al:'center'},
  a37b: {ff:'Archivo',sz:37,b:1,c:WHITE},
  a42b: {ff:'Archivo',sz:42,b:1,c:WHITE},
  o10_5: {ff:'Open Sans',sz:10.5,c:WHITE},
  r66bCN: {ff:'Orbitron',sz:66,b:1,c:WHITE,al:'center',nw:1},
  r96bCN: {ff:'Orbitron',sz:96,b:1,c:WHITE,al:'center',nw:1},
  r80bCN: {ff:'Orbitron',sz:80,b:1,c:WHITE,al:'center',nw:1},
};

// ---- decorative wave outlines (normalised, integer-scaled by 200) --------
const A = {
  a1: [[0,1,200,0,200,198,188,198,179,188,136,77,120,49,112,41,95,43,58,71,43,74,24,53,0,1]],
  a2: [[0,1,200,0,200,198,184,198,172,188,119,84,97,53,48,18,0,1]],
  a3: [[0,1,200,0,200,198,172,188,97,53,48,18,0,1]],
  a4: [[0,1,200,0,200,198,188,198,179,188,136,77,120,49,112,41,103,40,95,43,58,71,43,74,36,70,24,53,0,1]],
  a5: [[0,28,77,1,102,5,129,57,161,163,200,199,0,200,0,28]],
  a6: [[0,1,200,0,200,198,194,200,188,198,179,188,168,165,136,77,120,49,112,41,103,40,95,43,58,71,50,75,43,74,36,70,24,53,0,1]],
  a7: [[0,1,200,0,200,198,184,198,172,188,158,165,119,84,97,53,84,41,48,18,0,1]],
  a8: [[0,1,200,0,200,198,179,188,120,49,95,43,43,74,24,53,0,1]],
  a9: [[0,1,200,0,200,198,172,188,97,53,0,1]],
  a10: [[0,28,102,5,161,163,200,199,0,200,0,28]],
  a11: [[0,1,200,0,200,198,184,198,172,188,97,53,48,18,0,1]],
  a12: [[200,200,200,0,0,0,0,26,29,10,54,3,69,5,84,14,114,51,178,171,199,199]],
  a13: [[200,200,200,0,0,0,32,17,80,58,164,165,194,196]],
  a14: [[200,200,200,0,0,0,33,18,81,59,165,166,196,197]],
  a15: [[200,200,200,0,0,0,84,67,194,195]],
  a16: [[0,10,0,200,200,200,109,19,65,0,4,9]],
  a17: [[0,1,200,0,200,198,188,198,179,188,136,77,120,49,112,41,103,40,95,43,58,71,50,75,43,74,36,70,24,53,0,1]],
  a18: [[0,1,200,0,200,198,179,188,120,49,95,43,43,74,0,1]],
  a19: [[0,28,102,5,200,199,0,200,0,28]],
  a20: [[15,0,185,0,197,6,200,200,0,200,0,13,5,4,15,0]],
  a21: [[0,1,200,0,200,198,188,198,179,188,136,77,120,49,112,41,103,40,95,43,58,71,43,74,24,53,0,1]],
  a22: [[200,200,200,0,0,0,0,26,54,3,69,5,84,14,114,51,178,171,199,199]],
  a23: [[200,200,200,0,0,0,80,58,194,196]],
  a24: [[200,200,200,0,0,0,81,59,196,197]],
  a25: [[200,200,200,0,0,0,0,26,54,3,84,14,114,51,199,199]],
  a26: [[0,1,200,0,200,198,194,200,188,198,179,188,136,81,120,51,112,41,103,36,94,33,84,34,55,40,38,40,26,34,16,25,0,1]],
  a27: [[0,1,200,0,200,198,194,200,188,198,179,188,168,165,136,81,120,51,112,41,103,36,94,33,84,34,55,40,38,40,26,34,16,25,0,1]],
  a28: [[0,1,200,0,200,198,184,198,172,188,158,165,119,84,97,53,48,18,0,1]],
  a29: [[0,28,19,26,77,1,102,5,129,57,161,163,200,199,0,200,0,28]],
  a30: [[200,200,200,0,0,0,0,26,29,10,54,3,69,5,84,14,99,30,114,51,178,171,199,199]],
  a31: [[200,200,200,0,0,0,39,25,84,67,167,169,194,195]],
  a32: [[0,10,0,200,200,200,146,73,109,19,65,0,4,9]],
  a33: [[200,0,199,200,0,186,0,29]],
  a34: [[0,28,77,1,102,5,161,163,200,199,0,200,0,28]],
  a35: [[172,48,62,84,0,200,172,48]],
  a36: [[8,0,183,0,200,57,194,199,5,198,8,0]],
  a37: [[196,83,36,198,96,1,196,83]],
  a38: [[198,83,36,198,95,2,198,83]],
  a39: [[0,0,73,0,200,200,14,200,0,0]],
  a40: [[119,93,200,93,53,0,0,0,11,65,137,200,119,93]],
  a41: [[137,105,200,105,41,2,0,0,16,51,153,200,137,105]],
  a42: [[156,131,200,133,0,0,177,200,156,131]],
  a43: [[99,77,200,77,77,0,0,0,16,90,120,200,99,77]],
  a44: [[200,100,164,177,100,200,36,177,0,110,23,37,100,0,164,23,200,100]],
  a45: [[200,100,177,164,100,200,23,164,0,100,23,37,100,0,177,37,200,100]],
  a46: [[200,100,100,200,0,100,100,0,200,100]],
  a47: [[200,99,100,200,0,89,100,0,200,99]],
  a48: [[200,100,177,164,110,200,36,177,0,100,36,23,100,0,164,23,200,100]],
  a49: [[200,100,164,177,100,200,36,177,0,105,23,36,90,0,164,23,200,100]],
  a50: [[200,100,164,177,100,200,23,164,0,100,36,23,100,0,164,23,200,100]],
  a51: [[200,100,1,123,80,2,200,100]],
  a52: [[0,0,3,0,4,112,6,121,14,130,25,132,31,128,36,121,39,107,39,49,45,33,56,25,68,27,75,33,79,43,81,141,84,150,92,159,102,161,109,157,113,150,116,136,118,72,123,62,129,56,142,54,152,62,158,78,160,178,164,186,169,192,176,194,183,192,191,183,194,173,194,112,200,95,198,117,196,182,188,195,176,200,165,195,157,182,155,80,152,71,145,62,134,60,128,64,123,71,120,85,119,143,114,159,103,167,91,165,84,159,79,149,77,51,75,42,67,33,60,31,53,33,47,38,44,46,42,114,39,125,33,133,21,138,13,136,4,125,0,108]],
  a53: [[24,6,41,21,14,35,24,6],[182,165,177,195,81,195,84,165],[24,0,0,27,41,41,44,135,39,6,152,20,96,140,153,122,153,159,64,160,54,179,83,180,61,194,43,141,65,200,198,186,159,159,159,120,199,41,174,44,193,46,163,110,133,119,170,67,120,106,168,49,135,0]],
  a54: [[6,0,40,198,81,86,123,198,156,86,195,199,160,2,119,117,77,2,44,117,6,0]],
  a55: [[8,0,44,194,100,86,151,198,193,1,146,117,94,2,54,117,8,0]],
  a56: [[11,0,57,194,130,86,191,199,143,7,70,117,11,0]],
  a57: [[103,7,128,22,109,54,84,37,103,7],[176,50,190,77,165,91,176,50],[55,91,74,109,55,130,36,110,55,91],[79,0,92,51,31,106,58,171,10,172,2,200,121,199,94,154,63,171,77,99,110,65,136,106,194,84,186,27,161,21,181,30,137,99,113,58,153,21,79,0]],
  a58: [[153,0,175,9,175,174,3,157,29,200,198,196,24,191,181,178,183,25,199,174,200,26,171,0]],
  a59: [[66,0,0,31,6,200,13,31,200,27,82,0]],
  a60: [[11,0,11,200,191,198,189,0]],
  a61: [[11,1,11,200,191,198,189,1]],
  a62: [[182,0,7,12,11,195,152,191,182,0]],
  a63: [[183,0,7,10,11,195,152,191,183,0]],
  a64: [[24,0,24,200,155,187,135,0]],
  a65: [[100,0,0,100,101,200,200,100,43,140,100,0]],
  a66: [[72,53,93,65,72,78,72,53],[125,53,125,78,104,65,125,53],[43,39,66,81,39,97,16,67,43,39],[154,40,181,67,158,97,131,81,154,40],[66,87,66,113,45,100,66,87],[131,88,152,100,131,113,131,88],[98,68,125,84,98,132,71,116,98,68],[125,123,125,148,104,135,125,123],[72,123,93,136,72,148,72,123],[157,104,181,136,150,161,130,120,157,104],[98,139,125,155,98,195,71,167,98,139],[98,0,66,42,2,47,33,100,15,151,43,165,16,133,39,103,66,119,49,164,98,200,131,158,195,153,163,100,174,41,127,24,98,62,71,46,88,8,125,18,98,0]],
  a67: [[194,35,194,54,6,54,6,35],[86,0,83,24,1,25,13,93,18,65,182,65,185,200,200,26,117,24,114,0]],
  a68: [[116,133,124,155,76,155,85,133],[15,0,0,130,63,133,39,198,85,180,73,163,127,163,94,179,160,200,137,133,199,132,187,79,182,103,18,103,15,0]],
  a69: [[180,0,136,97,58,51,9,195,63,102,141,148,180,0]],
  a70: [[99,0,200,200,0,200,99,0]],
  a71: [[0,0,0,200,200,200,0,0]],
  a72: [[100,0,200,23,100,200,0,23,100,0]],
  a73: [[100,0,171,7,200,26,175,170,142,193,100,200,58,193,25,170,0,26,29,7,100,0]],
  a74: [[100,0,171,8,200,27,162,179,100,200,38,179,0,27,29,8,100,0]],
  a75: [[0,0,33,21,100,29,167,21,200,0,175,166,159,183,111,199,67,195,28,171,20,134]],
  a76: [[100,0,139,8,171,29,192,61,200,100,189,146,167,174,137,193,100,200,63,193,33,174,11,146,0,100,8,61,29,29,61,8,100,0]],
  a77: [[200,0,184,200,178,164,160,141,100,118,40,141,22,164,16,200,0,0,44,66,100,87,156,66,200,0]],
  a78: [[100,0,171,29,200,100,176,165,134,194,100,200,66,194,24,165,0,100,29,29,100,0]],
  a79: [[200,0,162,175,100,200,38,175,0,0,36,21,100,29,164,21,199,3]],
  a80: [[200,0,177,200,159,147,100,118,41,147,23,200,0,0,40,56,100,77,160,56,200,0]],
  a81: [[100,0,176,35,200,98,167,174,100,200,33,174,0,98,24,35,100,0]],
  a82: [[0,0,100,21,200,0,107,198,0,0]],
  a83: [[139,0,167,11,196,57,200,98,186,143,154,180,133,192,83,200,28,181,22,193,0,128,61,113,55,124,110,134,142,122,166,96,176,43,152,9,121,3,139,0]],
  a84: [[139,0,168,12,196,57,200,97,186,143,154,180,133,192,83,200,28,181,22,193,0,128,33,151,92,169,139,158,172,127,189,82,189,60,162,13,121,2,139,0]],
  a85: [[104,0,137,4,172,19,178,7,200,72,139,87,145,76,108,65,58,78,26,123,24,157,47,191,79,197,44,196,29,185,7,151,0,106,4,81,19,49,41,24,70,7,104,0]],
  a86: [[104,0,137,4,172,19,178,7,200,72,167,49,108,31,83,33,42,56,17,94,11,140,37,187,79,198,61,200,29,185,7,151,4,81,19,49,41,24,70,7,104,0]],
  a87: [[133,65,157,98,137,134,103,88,133,65],[66,58,115,66,95,88,141,153,85,116,106,185,42,91,66,58],[39,0,24,24,55,31,23,75,20,30,0,55,96,200,199,58,179,35,177,82,145,38,175,29,161,7,122,66,79,60,39,0]],
  a88: [[57,20,98,117,24,173,57,20],[143,0,20,27,18,199,101,151,186,197,191,52,175,173,107,133,112,39,179,34,143,0]],
  a89: [[179,0,199,93,199,107,181,197,179,200,177,197,161,121,164,110,179,182,196,100,179,18,139,200,122,121,124,110,140,182,179,0],[139,0,157,79,155,90,139,18,100,200,82,121,85,110,100,182,139,0],[100,0,118,79,115,90,100,18,61,200,43,121,45,110,61,182,100,0],[21,0,39,79,36,90,21,18,4,99,21,182,60,1,62,2,78,79,76,90,60,18,21,200,0,100,21,0]],
  a90: [[0,0,200,0,158,47,144,100,158,153,200,200,0,200,42,153,56,100,42,47,5,3]],
  a91: [[54,200,32,187,13,165,9,147,13,129,29,109,56,97,84,96,98,101,103,108,99,121,82,131,60,132,31,125,14,113,2,96,1,77,7,64,18,53,32,44,72,33,87,33,101,38,108,46,104,54,91,59,75,59,62,53,53,43,51,34,58,20,68,13,88,5,128,0,200,6]],
  a92: [[167,106,200,200,194,60,0,0,167,106]],
  a93: [[146,200,200,8,0,0,146,200]],
  a94: [[0,115,200,200,87,0,0,115]],
  a95: [[106,200,200,77,0,0,106,200]],
  a96: [[0,1,200,0,200,198,188,198,179,188,136,81,112,41,94,33,38,40,16,25,0,1]],
  a97: [[182,0,8,4,8,196,192,196,182,0],[18,12,185,12,100,125,18,12],[9,175,9,23,66,100,9,175],[182,188,15,188,72,109,182,188],[191,175,134,100,191,23,191,175]],
  a98: [[98,200,200,71,98,0,0,71,98,200],[98,19,173,71,98,181,27,71,98,19]],
  a99: [[197,2,147,23,188,167,136,181,133,75,130,181,64,75,12,179,49,9,3,22,4,199,197,180,197,2]],
  a100: [[100,200,200,100,100,0,0,100,100,200],[100,183,40,119,175,135,100,183],[100,17,179,127,23,131,100,17]],
  a101: [[175,5,181,125,100,164,19,125,26,3,32,183,181,169,175,5]],
};

// ---- repeated copy -------------------------------------------------------
const L1 = 'Rodeix Template',
      L2 = 'Section One',
      L3 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ',
      L4 = 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit, sed do eius mod tempor incididunt.',
      L5 = 'Section Two',
      L6 = 'Lorem ipsum dolorao sit amet, consectetur adipiscing elit, ',
      L7 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim',
      L8 = 'consectetur adipiscing elit',
      L9 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna',
      L10 = 'Virtual Reality & Metaverse  Presentation Template',
      L11 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.',
      L12 = 'Section Three',
      L13 = 'Lorem ipsum dolor sit ame conse ctetur adipiscing elit sed do eius mod tem.',
      L14 = 'Fransiska',
      L15 = 'Lorem ipsum dolor sit amet, consectetur adipiscingaw',
      L16 = 'Lorem ipsum dolor sit amet, consectetur adipi cing elit, sed do eiusmod tempor incididunt ut',
      L17 = 'Steven Carlos',
      L18 = 'Dohn Joe',
      L19 = 'Progres one',
      L20 = 'Progres two',
      L21 = 'Our Service 01',
      L22 = 'Our Service 02',
      L23 = 'Lorem ipsum dolor sit amet, consectetur adipiscingaw adipiscinga',
      L24 = 'IMG PNG FILE',
      L25 = 'The first million words are the hardest',
      L26 = 'The metaverse will be the real world in the near future”',
      L27 = 'Virtual lives become erotic when blind to boundaries',
      L28 = 'Progres Analysis ',
      L29 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minimLorem ipsum dolor sit amet, consectetur.',
      L30 = 'Lorem ipsum dolor sit amet',
      L31 = 'Our Service 03',
      L32 = 'Section 01',
      L33 = 'Section 02',
      L34 = 'Section 03',
      L35 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqa',
      L36 = 'Scarlet Pantom',
      L37 = 'Lorem ipsum dolor sit amet, consectetur adipis cing elit, sed do eiusmod tempor incididunt ut la',
      L38 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation';

// ---- helpers -------------------------------------------------------------
function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BG };
  return s;
}

/** Text box. `st` is a T.* style; opt may add fill/line/shape/rotate. */
function txt(s, text, st, x, y, w, h, opt) {
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: h,
    fontFace: st.ff, fontSize: st.sz, bold: !!st.b, italic: !!st.i, color: st.c,
    align: st.al || 'left', valign: st.v || 'top', margin: 0,
    lineSpacingMultiple: st.ls, wrap: st.nw ? false : true, fill: { type: 'none' },
  }, opt || {}));
}

/** Multi-run text box: runs are [text, style, startsNewParagraph?]. */
function rich(s, runs, x, y, w, h, opt) {
  const first = runs[0][1];
  const body = runs.map(function (r, i) {
    const st = r[1];
    return {
      text: r[0],
      options: {
        fontFace: st.ff, fontSize: st.sz, bold: !!st.b, italic: !!st.i, color: st.c,
        align: st.al || 'left', lineSpacingMultiple: st.ls, breakLine: !!r[2] && i > 0,
      },
    };
  });
  s.addText(body, Object.assign({
    x: x, y: y, w: w, h: h, margin: 0, valign: first.v || 'top',
    wrap: first.nw ? false : true, fill: { type: 'none' },
  }, opt || {}));
}

/** Preset autoshape. */
function shp(s, kind, x, y, w, h, opt) {
  s.addShape(kind, Object.assign({ x: x, y: y, w: w, h: h, line: NOLINE }, opt || {}));
}

/** Connector / rule drawn as a straight or bent line. */
function conn(s, kind, x, y, w, h, line, opt) {
  s.addShape(kind, Object.assign({
    x: x, y: y, w: w, h: h, fill: { type: 'none' },
    line: Object.assign({ color: WHITE, width: 1 }, line),
  }, opt || {}));
}

/** Decorative freeform: `poly` is a list of flat [x0,y0,x1,y1,...] rings in 0..1 space. */
function art(s, poly, x, y, w, h, fill, opt) {
  const K = 200;                       // rings are stored as integers scaled by K
  poly.forEach(function (ring) {
    const pts = [];
    for (let i = 0; i < ring.length; i += 2) {
      pts.push({ x: (ring[i] / K) * w, y: (ring[i + 1] / K) * h, moveTo: i === 0 });
    }
    pts.push({ close: true });
    s.addShape('custGeom', Object.assign({
      x: x, y: y, w: w, h: h, points: pts, fill: fill, line: NOLINE,
    }, opt || {}));
  });
}

/** The soft magenta "bokeh" dots scattered over most slides. */
function bubbles(s, list) {
  list.forEach(function (b) {
    s.addShape('ellipse', { x: b[0], y: b[1], w: b[2], h: b[2], line: NOLINE,
      fill: { color: '8E1A6B', transparency: 35 } });
  });
}

/** Rocket mark + "RodeiX" wordmark lock-up. */
function logo(s, x, y, color) {
  s.addShape('triangle', { x: x, y: y + 0.05, w: 0.22, h: 0.22, rotate: 45,
    fill: { color: color || PINK }, line: NOLINE });
  s.addText('RodeiX', { x: x + 0.17, y: y - 0.01, w: 1.2, h: 0.4, margin: 0,
    fontFace: 'Orbitron', fontSize: 18, bold: true, italic: true,
    color: color || PINK, valign: 'middle', wrap: false, fill: { type: 'none' } });
}

/** Small pictogram from the original deck -> flat pink glyph plate. */
function mark(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: Math.min(w, h) * 0.2,
    fill: { color: PINK, transparency: 15 }, line: NOLINE });
}

/** Full photograph from the original deck -> labelled placeholder plate. */
function photo(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.12,
    fill: { color: '3A1470' }, line: { color: PINK, width: 1.5 } });
  s.addText('[image]', { x: x, y: y + h / 2 - 0.2, w: w, h: 0.4, margin: 0,
    fontFace: 'Open Sans', fontSize: 12, color: 'BFA6E0', align: 'center' });
}

/** Slides 40-43 are pure icon walls; each cell is one preset silhouette. */
function iconWall(s, cells) {
  cells.forEach(function (c) {
    s.addShape(c[4], { x: c[0], y: c[1], w: c[2], h: c[3], fill: { color: WHITE }, line: NOLINE });
  });
}


// ---- shared background art (one function per master layout) -------------

function deco2(s) {
  shp(s, 'rect', 0,0,13.33,7.5, {fill:F.c1})
  shp(s, 'ellipse', 8.54,2.52,1.29,1.29, {fill:F.c2})
  shp(s, 'ellipse', 2.41,1.4,.66,.66, {fill:F.c3})
  shp(s, 'ellipse', 6.8,1.17,.25,.25, {fill:F.c3})
  art(s, A.a1, 5.48,-.01,7.87,2.38, F.c4)
  art(s, A.a2, 7.67,-.01,5.67,2.25, F.c2)
  art(s, A.a2, 7.98,-.01,5.36,2.18, F.c5)
  art(s, A.a2, 9.17,-.02,4.17,2.03, F.c6)
  art(s, A.a3, 10.53,0,2.82,1.37, F.c6)
  art(s, A.a4, -.01,3.96,9.41,3.54, F.c7, {rotate:180})
  art(s, A.a2, .01,4.16,6.77,3.35, F.c2, {rotate:180})
  art(s, A.a2, 0,4.15,6.59,3.36, F.c5, {rotate:180})
  art(s, A.a2, 0,4.68,5.94,2.85, F.c8, {rotate:180})
  art(s, A.a5, -.01,5.73,2.84,1.77, F.c2)
  bubbles(s, [
    [1.09,3.04,.25], [3.47,5.47,.332], [8.87,5.98,.452], [10.7,1.62,.454], [11.85,4.41,.25]
  ])
}

function deco1(s) {
  shp(s, 'rect', 0,0,13.33,7.5)
}

function deco3(s) {
  art(s, A.a6, 0,-.01,13.33,2.37, F.c4, {flipH:1})
  art(s, A.a7, .02,-.01,9.6,2.24, F.c2, {flipH:1})
  art(s, A.a7, .02,-.01,9.07,2.17, F.c5, {flipH:1})
  art(s, A.a2, .02,-.02,7.06,2.02, F.c6, {flipH:1})
  art(s, A.a2, 0,-.01,4.77,1.37, F.c6, {flipH:1})
  txt(s, L24, T.a18, 0,0,6.67,7.5)
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
}

function deco4(s) {
  txt(s, L24, T.a18, 6.67,0,6.65,7.5)
  art(s, A.a8, 0,6.25,3.35,1.27, F.c7, {rotate:180})
  art(s, A.a9, .01,6.32,2.41,1.2, F.c2, {rotate:180})
  art(s, A.a9, .01,6.31,2.34,1.2, F.c5, {rotate:180})
  art(s, A.a9, .01,6.5,2.11,1.02, F.c8, {rotate:180})
  art(s, A.a10, 0,6.88,1.01,.63, F.c2)
  art(s, A.a1, 0,-.01,6.96,1.24, F.c4, {flipH:1})
  art(s, A.a2, .01,-.01,5.01,1.17, F.c2, {flipH:1})
  art(s, A.a2, .01,-.01,4.74,1.13, F.c5, {flipH:1})
  art(s, A.a11, .01,-.02,3.68,1.05, F.c6, {flipH:1})
  art(s, A.a9, 0,-.01,2.49,.71, F.c6, {flipH:1})
  art(s, A.a12, 8.07,2.24,7.5,3.03, F.c7, {rotate:270,flipV:1})
  art(s, A.a13, 9.58,.94,4.69,2.81, F.c2, {rotate:270,flipV:1})
  art(s, A.a14, 9.66,.87,4.54,2.8, F.c5, {rotate:270,flipV:1})
  art(s, A.a15, 10.48,.66,3.51,2.19, F.c8, {rotate:270,flipV:1})
  art(s, A.a16, 11.8,.31,1.85,1.22, F.c2, {rotate:90,flipV:1})
}

function deco5(s) {
  art(s, A.a1, 0,-.01,7.15,1.27, F.c4, {flipH:1})
  art(s, A.a2, .01,-.01,5.14,1.2, F.c2, {flipH:1})
  art(s, A.a2, .01,-.01,4.86,1.16, F.c5, {flipH:1})
  art(s, A.a11, .01,-.02,3.78,1.08, F.c6, {flipH:1})
  art(s, A.a3, 0,-.01,2.56,.73, F.c6, {flipH:1})
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
  shp(s, 'roundRect', .59,2.2,5.82,4.77, {fill:F.c9})
  shp(s, 'roundRect', .73,2.06,5.82,4.77, {fill:F.c10})
  shp(s, 'roundRect', .88,1.91,5.82,4.77)
}

function deco6(s) {
  art(s, A.a17, 3.05,-.01,10.28,1.5, F.c4)
  art(s, A.a2, 5.91,-.01,7.4,1.42, F.c2)
  art(s, A.a2, 6.32,-.01,6.99,1.37, F.c5)
  art(s, A.a2, 7.87,-.02,5.44,1.28, F.c6)
  art(s, A.a11, 9.65,-.01,3.68,.87, F.c6)
  art(s, A.a18, -.01,6.25,3.01,1.27, F.c7, {rotate:180})
  art(s, A.a9, 0,6.32,2.16,1.2, F.c2, {rotate:180})
  art(s, A.a9, 0,6.31,2.1,1.2, F.c5, {rotate:180})
  art(s, A.a9, 0,6.5,1.9,1.02, F.c8, {rotate:180})
  art(s, A.a19, -.01,6.88,.91,.63, F.c2)
  shp(s, 'roundRect', 8.12,.85,4.53,6.06, {rectRadius:.37})
}

function deco7(s) {
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
  art(s, A.a20, .88,1.19,5.14,6.31, F.c1)
}

function deco8(s) {
  art(s, A.a21, 5.03,-.01,8.3,1.18, F.c4)
  art(s, A.a2, 7.34,-.01,5.97,1.12, F.c2)
  art(s, A.a2, 7.67,-.01,5.65,1.08, F.c5)
  art(s, A.a2, 8.92,-.02,4.39,1.01, F.c6)
  art(s, A.a3, 10.36,-.01,2.97,.68, F.c6)
  art(s, A.a18, -.01,6.25,3.01,1.27, F.c7, {rotate:180})
  art(s, A.a9, 0,6.32,2.16,1.2, F.c2, {rotate:180})
  art(s, A.a9, 0,6.31,2.1,1.2, F.c5, {rotate:180})
  art(s, A.a9, 0,6.5,1.9,1.02, F.c8, {rotate:180})
  art(s, A.a19, -.01,6.88,.91,.63, F.c2)
  shp(s, 'roundRect', 7.24,1.17,5.28,3.2, {rectRadius:.26})
  shp(s, 'roundRect', 7.24,4.6,5.28,2.51, {rectRadius:.2})
}

function deco9(s) {
  art(s, A.a1, 0,-.01,7.15,1.27, F.c4, {flipH:1})
  art(s, A.a2, .01,-.01,5.14,1.2, F.c2, {flipH:1})
  art(s, A.a2, .01,-.01,4.86,1.16, F.c5, {flipH:1})
  art(s, A.a11, .01,-.02,3.78,1.08, F.c6, {flipH:1})
  art(s, A.a3, 0,-.01,2.56,.73, F.c6, {flipH:1})
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
  shp(s, 'roundRect', .61,3.22,5.53,3.66)
}

function deco10(s) {
  txt(s, L24, T.a18, 7.93,.61,3.77,4.03)
  art(s, A.a8, 0,6.25,3.35,1.27, F.c7, {rotate:180})
  art(s, A.a9, .01,6.32,2.41,1.2, F.c2, {rotate:180})
  art(s, A.a9, .01,6.31,2.34,1.2, F.c5, {rotate:180})
  art(s, A.a9, .01,6.5,2.11,1.02, F.c8, {rotate:180})
  art(s, A.a10, 0,6.88,1.01,.63, F.c2)
  art(s, A.a1, 0,-.01,6.96,1.24, F.c4, {flipH:1})
  art(s, A.a2, .01,-.01,5.01,1.17, F.c2, {flipH:1})
  art(s, A.a2, .01,-.01,4.74,1.13, F.c5, {flipH:1})
  art(s, A.a11, .01,-.02,3.68,1.05, F.c6, {flipH:1})
  art(s, A.a9, 0,-.01,2.49,.71, F.c6, {flipH:1})
  art(s, A.a12, 8.07,2.24,7.5,3.03, F.c7, {rotate:270,flipV:1})
  art(s, A.a13, 9.58,.94,4.69,2.81, F.c2, {rotate:270,flipV:1})
  art(s, A.a14, 9.66,.87,4.54,2.8, F.c5, {rotate:270,flipV:1})
  art(s, A.a15, 10.48,.66,3.51,2.19, F.c8, {rotate:270,flipV:1})
  art(s, A.a16, 11.8,.31,1.85,1.22, F.c2, {rotate:90,flipV:1})
  shp(s, 'roundRect', 7.81,2.23,4,4.14, {fill:F.c2,rectRadius:.68})
}

function deco11(s) {
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
  shp(s, 'rect', 0,0,6.49,7.5)
}

function deco12(s) {
  shp(s, 'rect', 0,0,13.33,7.5, {fill:F.c1})
  shp(s, 'ellipse', 9.17,2.52,1.29,1.29, {fill:F.c2})
  art(s, A.a12, 8.07,2.24,7.5,3.03, F.c7, {rotate:270,flipV:1})
  art(s, A.a13, 9.58,.94,4.69,2.81, F.c2, {rotate:270,flipV:1})
  art(s, A.a14, 9.66,.87,4.54,2.8, F.c5, {rotate:270,flipV:1})
  art(s, A.a15, 10.48,.66,3.51,2.19, F.c8, {rotate:270,flipV:1})
  art(s, A.a16, 11.8,.31,1.85,1.22, F.c2, {rotate:90,flipV:1})
  art(s, A.a12, -2.25,2.24,7.5,3.03, F.c7, {rotate:90,flipV:1})
  art(s, A.a13, -.95,3.75,4.69,2.81, F.c2, {rotate:90,flipV:1})
  art(s, A.a14, -.88,3.83,4.54,2.8, F.c5, {rotate:90,flipV:1})
  art(s, A.a15, -.67,4.65,3.51,2.19, F.c8, {rotate:90,flipV:1})
  art(s, A.a16, -.32,5.96,1.85,1.22, F.c2, {rotate:270,flipV:1})
  bubbles(s, [
    [1.09,3.04,.25], [2.41,1.4,.663], [3.47,5.47,.332], [8.87,5.98,.452], [9.96,.91,.454], [6.8,1.17,.25],
    [11.85,4.41,.25]
  ])
}

function deco13(s) {
  shp(s, 'rect', 0,0,13.33,7.5, {fill:F.c1})
  art(s, A.a12, 8.07,2.24,7.5,3.03, F.c7, {rotate:270,flipV:1})
  art(s, A.a13, 9.58,.94,4.69,2.81, F.c2, {rotate:270,flipV:1})
  art(s, A.a14, 9.66,.87,4.54,2.8, F.c5, {rotate:270,flipV:1})
  art(s, A.a15, 10.48,.66,3.51,2.19, F.c8, {rotate:270,flipV:1})
  art(s, A.a16, 11.8,.31,1.85,1.22, F.c2, {rotate:90,flipV:1})
  art(s, A.a12, -2.24,2.24,7.5,3.03, F.c7, {rotate:90,flipH:1,flipV:1})
  art(s, A.a13, -.95,.94,4.69,2.81, F.c2, {rotate:90,flipH:1,flipV:1})
  art(s, A.a14, -.87,.87,4.54,2.8, F.c5, {rotate:90,flipH:1,flipV:1})
  art(s, A.a15, -.67,.66,3.51,2.19, F.c8, {rotate:90,flipH:1,flipV:1})
  art(s, A.a16, -.32,.31,1.85,1.22, F.c2, {rotate:270,flipH:1,flipV:1})
}

function deco14(s) {
  shp(s, 'rect', 0,0,13.33,5.08)
  art(s, A.a6, -.07,5.6,13.4,1.89, F.c7, {rotate:180})
  art(s, A.a7, -.05,5.71,9.65,1.79, F.c2, {rotate:180})
  art(s, A.a7, -.05,5.7,9.38,1.79, F.c5, {rotate:180})
  art(s, A.a2, -.05,5.98,8.45,1.52, F.c8, {rotate:180})
  art(s, A.a5, -.07,6.55,4.04,.94, F.c2)
}

function deco15(s) {
  shp(s, 'roundRect', 9.09,.96,2.68,2.61)
  shp(s, 'roundRect', 5.32,.96,2.68,2.61)
  shp(s, 'roundRect', 1.56,.96,2.68,2.61)
  art(s, A.a22, 9.39,2.24,5.62,3.03, F.c7, {rotate:270,flipH:1,flipV:1})
  art(s, A.a23, 10.52,3.75,3.52,2.81, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a24, 10.58,3.83,3.4,2.8, F.c5, {rotate:270,flipH:1,flipV:1})
  art(s, A.a15, 11.2,4.65,2.63,2.19, F.c8, {rotate:270,flipH:1,flipV:1})
  art(s, A.a16, 12.18,5.96,1.39,1.22, F.c2, {rotate:90,flipH:1,flipV:1})
  art(s, A.a22, -1.68,2.24,5.62,3.03, F.c7, {rotate:90,flipV:1})
  art(s, A.a23, -.71,3.75,3.52,2.81, F.c2, {rotate:90,flipV:1})
  art(s, A.a24, -.65,3.83,3.4,2.8, F.c5, {rotate:90,flipV:1})
  art(s, A.a15, -.5,4.65,2.63,2.19, F.c8, {rotate:90,flipV:1})
  art(s, A.a16, -.23,5.96,1.39,1.22, F.c2, {rotate:270,flipV:1})
}

function deco16(s) {
  shp(s, 'roundRect', 3.73,.36,3.54,6.54)
  art(s, A.a25, -.98,2.24,3.3,3.03, F.c7, {rotate:90,flipV:1})
  art(s, A.a23, -.41,3.75,2.07,2.81, F.c2, {rotate:90,flipV:1})
  art(s, A.a24, -.38,3.83,2,2.8, F.c5, {rotate:90,flipV:1})
  art(s, A.a15, -.29,4.65,1.55,2.19, F.c8, {rotate:90,flipV:1})
  art(s, A.a16, -.14,5.96,.81,1.22, F.c2, {rotate:270,flipV:1})
  art(s, A.a25, -.81,2.24,2.72,3.03, F.c11, {rotate:90,flipH:1,flipV:1})
  art(s, A.a23, -.34,.94,1.7,2.81, F.c2, {rotate:90,flipH:1,flipV:1})
  art(s, A.a24, -.31,.87,1.64,2.8, F.c5, {rotate:90,flipH:1,flipV:1})
  art(s, A.a15, -.24,.66,1.27,2.19, F.c8, {rotate:90,flipH:1,flipV:1})
  art(s, A.a16, -.11,.31,.67,1.22, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a1, 8.8,2.97,7.44,1.62, F.c4, {rotate:270,flipH:1,flipV:1})
  art(s, A.a2, 9.88,4.05,5.36,1.53, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a2, 10.05,4.22,5.07,1.48, F.c5, {rotate:270,flipH:1,flipV:1})
  art(s, A.a11, 10.67,4.83,3.94,1.38, F.c6, {rotate:270,flipH:1,flipV:1})
  art(s, A.a3, 11.53,5.71,2.66,.94, F.c6, {rotate:270,flipH:1,flipV:1})
}

function deco18(s) {
  art(s, A.a26, .01,1.64,13.33,5.87, F.c12, {rotate:180,flipH:1})
  art(s, A.a7, 3.72,1.95,9.6,5.56, F.c2, {rotate:180,flipH:1})
  art(s, A.a7, 3.99,1.93,9.33,5.57, F.c5, {rotate:180,flipH:1})
  art(s, A.a2, 4.91,2.82,8.41,4.72, F.c8, {rotate:180,flipH:1})
  art(s, A.a5, 9.32,4.57,4.02,2.94, F.c2, {flipH:1})
  art(s, A.a6, -.04,-.01,13.39,2.63, F.c4, {rotate:180,flipV:1})
  art(s, A.a7, -.02,-.01,9.64,2.49, F.c2, {rotate:180,flipV:1})
  art(s, A.a7, -.02,-.01,9.11,2.41, F.c5, {rotate:180,flipV:1})
  art(s, A.a2, -.02,-.02,7.09,2.25, F.c6, {rotate:180,flipV:1})
  art(s, A.a2, -.05,-.01,4.79,1.52, F.c6, {rotate:180,flipV:1})
}

function deco19(s) {
  art(s, A.a27, -.03,1.64,13.37,5.87, F.c12, {rotate:180})
  art(s, A.a7, -.01,1.95,9.62,5.56, F.c2, {rotate:180})
  art(s, A.a7, -.01,1.93,9.36,5.57, F.c5, {rotate:180})
  art(s, A.a2, -.01,2.82,8.43,4.72, F.c8, {rotate:180})
  art(s, A.a5, -.03,4.57,4.03,2.94, F.c2)
  art(s, A.a6, -.02,-.01,13.37,2.63, F.c4, {rotate:180,flipH:1,flipV:1})
  art(s, A.a7, 3.7,-.01,9.62,2.49, F.c2, {rotate:180,flipH:1,flipV:1})
  art(s, A.a7, 4.22,-.01,9.1,2.41, F.c5, {rotate:180,flipH:1,flipV:1})
  art(s, A.a2, 6.25,-.02,7.08,2.25, F.c6, {rotate:180,flipH:1,flipV:1})
  art(s, A.a2, 8.56,-.01,4.78,1.52, F.c6, {rotate:180,flipH:1,flipV:1})
}

function deco20(s) {
  art(s, A.a27, -.01,0,13.37,5.87, F.c12)
  art(s, A.a7, 3.71,0,9.62,5.56, F.c2)
  art(s, A.a7, 3.97,0,9.36,5.57, F.c5)
  art(s, A.a2, 4.9,-.03,8.43,4.72, F.c8)
  art(s, A.a5, 9.32,0,4.03,2.94, F.c2, {rotate:180})
  art(s, A.a6, -.02,4.86,13.37,2.63, F.c4, {flipH:1,flipV:1})
  art(s, A.a7, 0,5,9.62,2.49, F.c2, {flipH:1,flipV:1})
  art(s, A.a7, 0,5.08,9.1,2.41, F.c5, {flipH:1,flipV:1})
  art(s, A.a2, 0,5.25,7.08,2.25, F.c6, {flipH:1,flipV:1})
  art(s, A.a2, -.02,5.97,4.78,1.52, F.c6, {flipH:1,flipV:1})
}

function deco21(s) {
  art(s, A.a26, -.01,0,13.33,5.87, F.c12, {flipH:1})
  art(s, A.a7, .01,0,9.6,5.56, F.c2, {flipH:1})
  art(s, A.a7, .01,0,9.33,5.57, F.c5, {flipH:1})
  art(s, A.a2, .01,-.03,8.41,4.72, F.c8, {flipH:1})
  art(s, A.a5, -.01,0,4.02,2.94, F.c2, {rotate:180,flipH:1})
  art(s, A.a6, -.05,4.89,13.39,2.63, F.c4, {flipV:1})
  art(s, A.a7, 3.68,5.03,9.64,2.49, F.c2, {flipV:1})
  art(s, A.a7, 4.21,5.11,9.11,2.41, F.c5, {flipV:1})
  art(s, A.a2, 6.23,5.28,7.09,2.25, F.c6, {flipV:1})
  art(s, A.a2, 8.55,6,4.79,1.52, F.c6, {flipV:1})
}

function deco22(s) {
  art(s, A.a27, -.07,-.04,13.48,2.52, F.c12)
  art(s, A.a7, 3.68,-.04,9.7,2.38, F.c2)
  art(s, A.a7, 3.95,-.04,9.43,2.39, F.c5)
  art(s, A.a28, 4.88,-.06,8.5,2.02, F.c8)
  art(s, A.a29, 9.34,-.04,4.06,1.26, F.c2, {rotate:180})
  shp(s, 'roundRect', 5.74,1.29,2.28,4.73)
  shp(s, 'roundRect', 8.17,1.29,2.28,4.73)
  shp(s, 'roundRect', 10.59,1.29,2.28,4.73)
  art(s, A.a27, -.07,6.02,13.41,1.5, F.c12, {rotate:180})
  art(s, A.a7, -.05,6.11,9.65,1.42, F.c2, {rotate:180})
  art(s, A.a7, -.05,6.1,9.38,1.43, F.c5, {rotate:180})
  art(s, A.a2, -.05,6.33,8.46,1.21, F.c8, {rotate:180})
  art(s, A.a5, -.07,6.78,4.04,.75, F.c2)
}

function deco23(s) {
  shp(s, 'roundRect', .74,.66,5.81,1.94, {fill:F.c2})
  shp(s, 'roundRect', .74,2.84,5.81,1.94, {fill:F.c2})
  shp(s, 'roundRect', .74,5.02,5.81,1.94, {fill:F.c2})
  shp(s, 'roundRect', 1.04,.88,2.23,1.44)
  shp(s, 'roundRect', 1.04,3.09,2.23,1.44)
  shp(s, 'roundRect', 1.04,5.31,2.23,1.44)
  art(s, A.a7, 3.68,6.17,9.7,1.36, F.c2, {flipV:1})
  art(s, A.a7, 4.21,6.21,9.17,1.31, F.c5, {flipV:1})
  art(s, A.a2, 6.25,6.3,7.13,1.22, F.c6, {flipV:1})
  art(s, A.a2, 8.58,6.69,4.82,.83, F.c6, {flipV:1})
}

function deco24(s) {
  shp(s, 'roundRect', 6.83,1.61,5.81,1.94, {fill:F.c2})
  shp(s, 'roundRect', 6.83,4,5.81,1.94, {fill:F.c2})
  shp(s, 'roundRect', 7.13,1.87,1.49,1.41)
  shp(s, 'roundRect', 7.13,4.26,1.49,1.41)
  art(s, A.a27, -.03,5.66,13.37,1.87, F.c12, {rotate:180})
  art(s, A.a7, -.01,5.76,9.62,1.77, F.c2, {rotate:180})
  art(s, A.a7, -.01,5.76,9.36,1.77, F.c5, {rotate:180})
  art(s, A.a2, -.01,6.04,8.43,1.5, F.c8, {rotate:180})
  art(s, A.a5, -.03,6.59,4.03,.93, F.c2)
  art(s, A.a6, -.02,-.01,13.37,1.62, F.c4, {rotate:180,flipH:1,flipV:1})
  art(s, A.a7, 3.7,-.01,9.62,1.53, F.c2, {rotate:180,flipH:1,flipV:1})
  art(s, A.a7, 4.22,-.01,9.1,1.48, F.c5, {rotate:180,flipH:1,flipV:1})
  art(s, A.a2, 6.25,-.02,7.08,1.38, F.c6, {rotate:180,flipH:1,flipV:1})
  art(s, A.a2, 8.56,-.01,4.78,.93, F.c6, {rotate:180,flipH:1,flipV:1})
}

function deco25(s) {
  shp(s, 'rect', 0,0,13.33,4.48)
}

function deco26(s) {
  shp(s, 'roundRect', 9.79,.96,3.28,5.82)
  shp(s, 'roundRect', 6.36,.96,3.28,5.82)
  art(s, A.a27, -.03,5.66,13.37,1.87, F.c12, {rotate:180})
  art(s, A.a7, -.01,5.76,9.62,1.77, F.c2, {rotate:180})
  art(s, A.a7, -.01,5.76,9.36,1.77, F.c5, {rotate:180})
  art(s, A.a2, -.01,6.04,8.43,1.5, F.c8, {rotate:180})
  art(s, A.a5, -.03,6.59,4.03,.93, F.c2)
}

function deco27(s) {
  art(s, A.a30, -3.65,2.24,12.21,3.03, F.c7, {rotate:90,flipH:1,flipV:1})
  art(s, A.a13, -1.54,.94,7.64,2.81, F.c2, {rotate:90,flipH:1,flipV:1})
  art(s, A.a14, -1.41,.87,7.39,2.8, F.c5, {rotate:90,flipH:1,flipV:1})
  art(s, A.a31, -1.09,.66,5.72,2.19, F.c8, {rotate:90,flipH:1,flipV:1})
  art(s, A.a32, -.52,.31,3.01,1.22, F.c2, {rotate:270,flipH:1,flipV:1})
  shp(s, 'rect', .73,.91,3.57,6.59)
  art(s, A.a7, 3.68,6.17,9.7,1.36, F.c2, {flipV:1})
  art(s, A.a7, 4.21,6.21,9.17,1.31, F.c5, {flipV:1})
  art(s, A.a2, 6.25,6.3,7.13,1.22, F.c6, {flipV:1})
  art(s, A.a2, 8.58,6.69,4.82,.83, F.c6, {flipV:1})
}

function deco28(s) {
  art(s, A.a1, 6.44,.63,7.44,6.3, F.c4, {rotate:270,flipH:1,flipV:1})
  art(s, A.a2, 7.65,1.83,5.36,5.96, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a2, 7.9,2.08,5.07,5.77, F.c5, {rotate:270,flipH:1,flipV:1})
  art(s, A.a2, 8.67,2.83,3.94,5.38, F.c6, {rotate:270,flipH:1,flipV:1})
  art(s, A.a11, 10.16,4.36,2.66,3.64, F.c6, {rotate:270,flipH:1,flipV:1})
  art(s, A.a33, 7.81,1.18,4.64,3.5, F.c1)
  art(s, A.a26, -.01,-.02,13.33,1.3, F.c12, {flipH:1})
  art(s, A.a7, .01,-.02,9.6,1.23, F.c2, {flipH:1})
  art(s, A.a7, .01,-.02,9.33,1.23, F.c5, {flipH:1})
  art(s, A.a2, .01,-.03,8.41,1.04, F.c8, {flipH:1})
  art(s, A.a5, -.01,-.02,4.02,.65, F.c2, {rotate:180,flipH:1})
}

function deco29(s) {
  art(s, A.a26, -.01,0,13.33,5.87, F.c12, {flipH:1})
  art(s, A.a7, .01,0,9.6,5.56, F.c2, {flipH:1})
  art(s, A.a7, .01,0,9.33,5.57, F.c5, {flipH:1})
  art(s, A.a2, .01,-.03,8.41,4.72, F.c8, {flipH:1})
  art(s, A.a5, -.01,0,4.02,2.94, F.c2, {rotate:180,flipH:1})
  shp(s, 'rect', 1.25,1.85,4.62,3.21)
  art(s, A.a8, 9.96,6.25,3.37,1.27, F.c7, {rotate:180,flipH:1})
  art(s, A.a9, 10.9,6.32,2.43,1.2, F.c2, {rotate:180,flipH:1})
  art(s, A.a9, 10.97,6.32,2.36,1.2, F.c5, {rotate:180,flipH:1})
  art(s, A.a9, 11.2,6.51,2.13,1.02, F.c8, {rotate:180,flipH:1})
  art(s, A.a10, 12.32,6.88,1.02,.63, F.c2, {flipH:1})
}


// ---- slides --------------------------------------------------------------

function slide1(pptx) {
  const s = newSlide(pptx);
  deco2(s);
  txt(s, L10, T.r14CN, 3.75,4.24,5.84,.34)
  mark(s, 5.97,1.87,1.07,1.07)
  txt(s, 'RodeiX', T.r96bN, 3.93,2.57,5.46,1.72)
}

function slide2(pptx) {
  const s = newSlide(pptx);
  deco1(s);
  shp(s, 'rect', 0,0,13.33,7.5, {fill:F.c13})
  art(s, A.a1, 5.48,-.01,7.87,2.38, F.c4)
  art(s, A.a2, 7.67,-.01,5.67,2.25, F.c2)
  art(s, A.a2, 7.98,-.01,5.36,2.18, F.c5)
  art(s, A.a2, 9.17,-.02,4.17,2.03, F.c6)
  art(s, A.a3, 10.53,-.01,2.82,1.37, F.c6)
  art(s, A.a4, -.01,3.96,9.41,3.54, F.c7, {rotate:180})
  art(s, A.a2, .01,4.16,6.77,3.35, F.c2, {rotate:180})
  art(s, A.a2, 0,4.15,6.59,3.36, F.c5, {rotate:180})
  art(s, A.a2, 0,4.68,5.94,2.85, F.c8, {rotate:180})
  art(s, A.a5, -.01,5.73,2.84,1.77, F.c2)
  shp(s, 'ellipse', 8.54,2.52,1.29,1.29, {fill:F.c2})
  txt(s, L10, T.r14CN, 3.75,4.24,5.84,.34)
  mark(s, 5.97,1.87,1.07,1.07)
  txt(s, 'RodeiX', T.r96bN, 3.93,2.57,5.46,1.72)
  bubbles(s, [
    [1.09,3.04,.25], [2.41,1.4,.663], [3.47,5.47,.332], [8.87,5.98,.452], [10.7,1.62,.454], [6.8,1.17,.25],
    [11.85,4.41,.25]
  ])
}

function slide3(pptx) {
  const s = newSlide(pptx);
  deco3(s);
  txt(s, L1, T.r11bpN, 7.57,1.59,1.65,.29)
  txt(s, 'Welcome to RodeiX Virtuality', T.a44b, 7.57,1.93,5.19,1.58)
  logo(s, 11.64,.46)
  txt(s, L11, T.o10JS, 7.57,3.63,4.81,.83)
  shp(s, 'roundRect', 7.7,4.85,2.05,.61, {fill:F.c10,rectRadius:.19})
  txt(s, 'LEARN MORE', T.a14bC, 7.57,4.98,2.32,.34)
  bubbles(s, [
    [6.71,6.66,.452], [6.18,1.47,.25], [12.64,5.06,.181]
  ])
}

function slide4(pptx) {
  const s = newSlide(pptx);
  deco4(s);
  txt(s, L1, T.r11bpN, 1.34,1.77,1.65,.29)
  txt(s, 'A coward talks to everyone but YOU.', T.a40b, 1.34,2.05,6.17,1.45)
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit', T.o10JS, 1.35,4.05,5.32,1.08)
  txt(s, L2, T.a12b, 1.34,3.75,1.51,.3)
  bubbles(s, [
    [5.41,6.42,.452], [7.24,1.41,.169], [1.31,5.7,.169], [3.65,1.05,.095]
  ])
}

function slide5(pptx) {
  const s = newSlide(pptx);
  deco5(s);
  txt(s, L1, T.r11bpN, 7.41,1.63,1.65,.29)
  txt(s, 'In this twenty-first century, reality is virtual and truth is broader.', T.a28b, 7.41,1.91,5.42,1.51)
  txt(s, L7, T.o10JS, 7.42,3.92,5.13,.58)
  txt(s, L2, T.a12b, 7.41,3.62,1.51,.3)
  logo(s, 11.64,.39)
  txt(s, L7, T.o10JS, 7.42,5.01,5.13,.58)
  txt(s, L5, T.a12b, 7.41,4.71,1.51,.3)
  shp(s, 'roundRect', 4.76,1.53,1.91,.8, {fill:F.c10,rectRadius:.25})
  logo(s, 5.04,1.72, WHITE)
  bubbles(s, [
    [8.01,6.46,.226], [10.98,1.22,.169], [7.38,1,.095]
  ])
}

function slide6(pptx) {
  const s = newSlide(pptx);
  deco6(s);
  art(s, A.a1, 6.16,4.84,7.17,2.68, F.c12, {rotate:180,flipH:1})
  art(s, A.a2, 8.16,4.98,5.16,2.53, F.c2, {rotate:180,flipH:1})
  art(s, A.a2, 8.3,4.97,5.02,2.54, F.c5, {rotate:180,flipH:1})
  art(s, A.a2, 8.8,5.38,4.52,2.15, F.c8, {rotate:180,flipH:1})
  art(s, A.a34, 11.17,6.17,2.16,1.34, F.c2, {flipH:1})
  txt(s, L1, T.r11bpN, 1.24,1.63,1.65,.29)
  txt(s, L25, T.a29b, 1.24,1.91,5.42,1.08)
  logo(s, .41,.45)
  txt(s, L7, T.o10JS, 1.26,3.08,4.91,.58)
  shp(s, 'roundRect', 7.84,.63,1.91,.8, {fill:F.c10,rectRadius:.25})
  logo(s, 8.12,.82, WHITE)
  txt(s, L8, T.o10, 1.24,4.06,2.07,.28)
  shp(s, 'rect', 1.33,4.49,4.31,.08, {fill:F.c10})
  txt(s, L19, T.a12b, 1.24,3.79,1.95,.3)
  shp(s, 'rect', 5.03,4.49,1.02,.08, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 5.45,3.98,.69,.28)
  txt(s, L8, T.o10, 1.24,5.02,2.07,.28)
  shp(s, 'rect', 1.33,5.44,4.43,.08, {fill:F.c10})
  txt(s, L20, T.a12b, 1.24,4.75,1.95,.3)
  shp(s, 'rect', 5.36,5.44,.69,.08, {fill:F.c14})
  txt(s, '95.2%', T.a10_5bR, 5.45,4.94,.69,.28)
  bubbles(s, [
    [5,6.62,.452], [7.24,1.41,.169], [.6,5.49,.169], [3.65,1.05,.095]
  ])
}

function slide7(pptx) {
  const s = newSlide(pptx);
  deco7(s);
  art(s, A.a6, 0,-.01,11.61,2.06, F.c4, {flipH:1})
  art(s, A.a2, .02,-.01,8.36,1.95, F.c2, {flipH:1})
  art(s, A.a2, .02,-.01,7.9,1.89, F.c5, {flipH:1})
  art(s, A.a2, .02,-.02,6.15,1.76, F.c6, {flipH:1})
  art(s, A.a2, 0,-.01,4.16,1.19, F.c6, {flipH:1})
  txt(s, L1, T.r11bpN, 6.97,1.88,1.65,.29)
  txt(s, 'Did you enjoy killing me, your honour', T.a32b, 6.97,2.16,5.42,1.18)
  txt(s, L3, T.o10JS, 7.57,3.93,4.59,.58)
  txt(s, L2, T.a12b, 7.56,3.62,1.51,.3)
  logo(s, 11.64,.39)
  txt(s, L3, T.o10JS, 7.57,5.09,4.59,.58)
  txt(s, L5, T.a12b, 7.56,4.79,1.51,.3)
  txt(s, '1', T.a14bCM, 7.11,3.67,.4,.4, {shape:'ellipse',fill:F.c10})
  txt(s, '2', T.a14bCM, 7.11,4.82,.4,.4, {shape:'ellipse',fill:F.c10})
  bubbles(s, [
    [8.01,6.46,.226], [10.98,1.22,.169], [7.38,1,.095]
  ])
}

function slide8(pptx) {
  const s = newSlide(pptx);
  deco8(s);
  txt(s, L1, T.r11bpN, 1.24,1.81,1.65,.29)
  txt(s, L26, T.a28b, 1.24,2.09,5.42,1.04)
  txt(s, L3, T.o10JS, 1.8,3.74,4.59,.58)
  txt(s, L2, T.a12b, 1.79,3.44,1.51,.3)
  txt(s, L3, T.o10JS, 1.8,4.9,4.59,.58)
  txt(s, L5, T.a12b, 1.79,4.6,1.51,.3)
  logo(s, .41,.45)
  mark(s, 1.32,3.52,.47,.47)
  mark(s, 1.32,4.73,.47,.47)
  bubbles(s, [
    [5,6.62,.452], [6.7,2,.169], [.6,5.49,.169], [3.65,1.05,.095]
  ])
}

function slide9(pptx) {
  const s = newSlide(pptx);
  deco9(s);
  txt(s, L1, T.r11bpN, 6.78,1.05,1.65,.29)
  txt(s, L27, T.a28b, 6.78,1.33,5.42,1.04)
  txt(s, L3, T.o10JS, 7.33,2.99,4.59,.58)
  txt(s, L2, T.a12b, 7.32,2.68,1.51,.3)
  txt(s, L3, T.o10JS, 7.33,4.15,4.59,.58)
  txt(s, L5, T.a12b, 7.32,3.85,1.51,.3)
  mark(s, 6.85,2.76,.47,.47)
  mark(s, 6.85,3.97,.47,.47)
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minimLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', T.o10JS, .87,1.69,5,1.08)
  txt(s, 'Lorem Ipsum Dolor Sit Amet', T.a12b, .86,1.39,2.41,.3)
  txt(s, L3, T.o10JS, 7.33,5.31,4.59,.58)
  txt(s, L12, T.a12b, 7.32,5.01,1.51,.3)
  logo(s, 11.64,.39)
  shp(s, 'rect', 6.91,5.38,.07,.07, {fill:F.c10})
  shp(s, 'rect', 7.01,5.33,.07,.12, {fill:F.c10})
  shp(s, 'rect', 7.11,5.28,.07,.17, {fill:F.c10})
  shp(s, 'rect', 7.21,5.23,.07,.22, {fill:F.c10})
  bubbles(s, [
    [7.85,6.49,.226], [12.4,2.99,.169], [4.47,1.15,.095], [.3,2.38,.169], [8.72,.58,.095]
  ])
}

function slide10(pptx) {
  const s = newSlide(pptx);
  deco10(s);
  shp(s, 'roundRect', 8.26,4.59,3.2,1.46, {fill:F.c1})
  shp(s, 'rect', 8.53,5.12,2.52,.05, {fill:F.c10})
  rich(s, [[L28,T.a11b],['RodeiX',T.a11bip]], 8.43,4.75,2.14,.29)
  shp(s, 'rect', 10.75,5.12,.39,.05, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 10.53,4.8,.69,.28)
  shp(s, 'rect', 8.53,5.7,2.52,.05, {fill:F.c10})
  shp(s, 'rect', 10.75,5.7,.39,.05, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 10.53,5.38,.69,.28)
  txt(s, L1, T.r11bpN, 1.22,1.73,1.65,.29)
  rich(s, [['Virtual world ',T.a40b],['exists & now after ',T.a40bp],['metaverse boom virtual ',T.a40b],['world is developing',T.a40bp],[' very fast',T.a40b]], 1.22,2.01,6.37,2.79)
  rich(s, [[L28,T.a11b],['RodeiX',T.a11bip]], 8.43,5.34,2.14,.29)
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minimLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', T.o10JS, 1.22,4.9,5.71,.83)
  bubbles(s, [
    [5.58,6.09,.226], [4.47,1.15,.095], [.3,2.38,.169]
  ])
}

function slide11(pptx) {
  const s = newSlide(pptx);
  deco11(s);
  shp(s, 'roundRect', 4.25,1.43,8.42,4.81, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.7})
  txt(s, L1, T.r11bpN, 5.25,1.95,1.65,.29)
  rich(s, [['Something drew ',T.a32b],['her attention. Her',T.a32bp],[' eyes narro wed as if to bri ng ',T.a32b],['a distant object ',T.a32bp],['into focus.',T.a32b]], 5.25,2.23,6.86,1.72)
  logo(s, 11.64,.39)
  txt(s, L11, T.o10JS, 5.29,4.06,6.52,.58)
  shp(s, 'roundRect', 5.4,4.86,1.84,.58, {fill:F.c10,rectRadius:.18})
  txt(s, 'Read More', T.a16bC, 5.5,4.97,1.62,.37)
  shp(s, 'roundRect', 7.46,4.86,1.84,.58, {fill:F.c1,line:{color:PINK,width:3},rectRadius:.18})
  txt(s, 'View Profile', T.a16bC, 7.56,4.97,1.62,.37)
  art(s, A.a1, 0,-.01,8,1.85, F.c4, {flipH:1})
  art(s, A.a2, .01,-.01,5.76,1.75, F.c2, {flipH:1})
  art(s, A.a2, .01,-.01,5.44,1.7, F.c5, {flipH:1})
  art(s, A.a2, .01,-.02,4.23,1.58, F.c6, {flipH:1})
  art(s, A.a3, 0,-.01,2.86,1.07, F.c6, {flipH:1})
  art(s, A.a1, -.01,5.34,5.78,2.18, F.c7, {rotate:180})
  art(s, A.a2, 0,5.45,4.16,2.06, F.c2, {rotate:180})
  art(s, A.a11, 0,5.45,4.05,2.07, F.c5, {rotate:180})
  art(s, A.a11, 0,5.78,3.65,1.75, F.c8, {rotate:180})
  art(s, A.a10, -.01,6.42,1.74,1.09, F.c2)
  bubbles(s, [
    [5.02,5.56,.226], [12.04,2.84,.169], [11.64,5.56,.169], [8.91,1.77,.095]
  ])
}

function slide12(pptx) {
  const s = newSlide(pptx);
  deco12(s);
  txt(s, L10, T.r14CN, 3.75,4.24,5.84,.34)
  mark(s, 6.13,1.86,1.07,1.07)
  txt(s, 'Break Slides', T.r66bN, 3.27,2.98,6.78,1.21)
}

function slide13(pptx) {
  const s = newSlide(pptx);
  deco13(s);
  shp(s, 'roundRect', .52,4.5,3.71,2.32, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.34})
  shp(s, 'roundRect', 4.81,4.5,3.71,2.32, {fill:F.c10,line:{color:WHITE,width:6},rectRadius:.34})
  shp(s, 'roundRect', 9.1,4.5,3.71,2.32, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.34})
  txt(s, L9, T.o10CS, .91,5.67,2.93,.83)
  txt(s, L2, T.a16bC, 1.62,5.29,1.51,.37)
  mark(s, 2.14,4.74,.47,.47)
  txt(s, L9, T.o10CS, 9.49,5.67,2.93,.83)
  txt(s, L12, T.a16bC, 10.11,5.29,1.69,.37)
  txt(s, L9, T.o10CS, 5.2,5.67,2.93,.83)
  txt(s, L5, T.a16bC, 5.88,5.29,1.57,.37)
  shp(s, 'rect', 10.76,5.05,.08,.08, {fill:F.c10})
  shp(s, 'rect', 10.87,4.99,.08,.13, {fill:F.c10})
  shp(s, 'rect', 10.97,4.94,.08,.19, {fill:F.c10})
  shp(s, 'rect', 11.08,4.89,.08,.24, {fill:F.c10})
  art(s, A.a35, 10.73,4.74,.42,.25, F.c10)
  mark(s, 6.37,4.74,.5,.5)
  txt(s, L1, T.r11bCpN, 5.67,1.18,1.65,.29)
  rich(s, [['I think ',T.a24bC],['that real',T.a24bCp],[' landscapes enter i',T.a24bC],['nto pictures, not that pictures ',T.a24bCp],['will one ',T.a24bC],['day sprout out into ',T.a24bCp],['real trees and grass',T.a24bC]], 2.53,1.46,7.92,1.31)
  txt(s, L11, T.o10CS, 2.38,3.02,8.23,.58)
  bubbles(s, [
    [1.51,3.09,.223], [3.75,.65,.226], [4.7,7,.124], [4.83,3.84,.223], [9.96,.91,.454], [10.7,3.71,.25]
  ])
}

function slide14(pptx) {
  const s = newSlide(pptx);
  deco14(s);
  shp(s, 'rect', 0,.32,13.33,4.91, {fill:F.c15})
  txt(s, L1, T.r11bpN, 1.07,1.7,1.65,.29)
  txt(s, L27, T.a28b, 1.07,1.98,5.42,1.04)
  txt(s, L3, T.o10JS, 1.63,3.64,4.59,.58)
  txt(s, L2, T.a12b, 1.61,3.34,1.51,.3)
  txt(s, L3, T.o10JS, 1.63,4.8,4.59,.58)
  txt(s, L5, T.a12b, 1.61,4.5,1.51,.3)
  mark(s, 1.14,3.41,.47,.47)
  mark(s, 1.14,4.62,.47,.47)
  shp(s, 'roundRect', 7.3,1.91,5.51,3.65, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.21})
  txt(s, L29, T.o10JS, 7.72,2.81,4.76,.83)
  txt(s, L30, T.a14b, 7.71,2.47,3.09,.34)
  txt(s, L29, T.o10JS, 7.72,4.16,4.76,.83)
  txt(s, L30, T.a14b, 7.71,3.82,3.09,.34)
  art(s, A.a36, 10.87,1.88,1.98,.59, F.c10)
  logo(s, 11.15,2.02, WHITE)
}

function slide15(pptx) {
  const s = newSlide(pptx);
  deco1(s);
  shp(s, 'rect', -.01,-.01,13.33,7.5, {fill:F.c16})
  shp(s, 'roundRect', 1.08,4.26,3.71,2.32, {fill:F.c17,rectRadius:.34})
  shp(s, 'roundRect', 5.32,4.26,3.71,2.32, {fill:F.c17,rectRadius:.34})
  txt(s, L9, T.o10CS, 1.48,5.43,2.93,.83)
  txt(s, L2, T.a16bC, 2.19,5.06,1.51,.37)
  mark(s, 2.71,4.5,.47,.47)
  txt(s, L9, T.o10CS, 5.71,5.43,2.93,.83)
  txt(s, L12, T.a16bC, 6.33,5.06,1.69,.37)
  shp(s, 'rect', 6.98,4.81,.08,.08, {fill:F.c14})
  shp(s, 'rect', 7.09,4.75,.08,.13, {fill:F.c14})
  shp(s, 'rect', 7.19,4.7,.08,.19, {fill:F.c14})
  shp(s, 'rect', 7.3,4.65,.08,.24, {fill:F.c14})
  art(s, A.a35, 6.95,4.5,.42,.25, F.c14)
  txt(s, L1, T.r11bpN, 1.08,.67,1.65,.29)
  rich(s, [['Earth Existence in a 3rd density world:',T.a32b],['A Veil of Virtual Reality.',T.a32b,1]], 1.08,.95,8.42,1.18)
  txt(s, L3, T.o10JS, 1.68,2.71,3.42,.83)
  txt(s, L2, T.a12b, 1.67,2.41,1.51,.3)
  txt(s, L3, T.o10JS, 5.74,2.71,3.42,.83)
  txt(s, L5, T.a12b, 5.73,2.41,1.51,.3)
  txt(s, '1', T.a14bCM, 1.23,2.46,.4,.4, {shape:'ellipse',fill:F.c10})
  txt(s, '2', T.a14bCM, 5.29,2.44,.4,.4, {shape:'ellipse',fill:F.c10})
  art(s, A.a1, 8.3,2.46,7.5,2.55, F.c12, {rotate:270,flipV:1})
  art(s, A.a2, 9.42,1.5,5.4,2.41, F.c2, {rotate:270,flipV:1})
  art(s, A.a2, 9.49,1.42,5.25,2.42, F.c5, {rotate:270,flipV:1})
  art(s, A.a2, 9.94,1.35,4.73,2.05, F.c8, {rotate:270,flipV:1})
  art(s, A.a34, 11.55,.49,2.26,1.27, F.c2, {rotate:90,flipV:1})
}

function slide16(pptx) {
  const s = newSlide(pptx);
  deco15(s);
  shp(s, 'ellipse', 2.39,3.07,1.01,1.01, {fill:F.c10})
  shp(s, 'ellipse', 6.16,3.07,1.01,1.01, {fill:F.c10})
  shp(s, 'ellipse', 9.92,3.07,1.01,1.01, {fill:F.c10})
  txt(s, L21, T.a16bC, 1.63,4.76,2.51,.37)
  art(s, A.a37, 2.24,4.45,.2,.19, F.c18)
  art(s, A.a38, 3.32,4.45,.2,.19, F.c18)
  art(s, A.a37, 2.52,4.45,.2,.19, F.c18)
  art(s, A.a37, 2.79,4.45,.2,.19, F.c18)
  art(s, A.a37, 3.06,4.45,.2,.19, F.c18)
  txt(s, L4, T.o10CS, 1.64,5.13,2.51,.83)
  txt(s, L22, T.a16bC, 5.38,4.76,2.51,.37)
  art(s, A.a37, 6,4.45,.2,.19, F.c18)
  art(s, A.a37, 6.27,4.45,.2,.19, F.c18)
  art(s, A.a37, 6.55,4.45,.2,.19, F.c18)
  art(s, A.a37, 6.82,4.45,.2,.19, F.c18)
  txt(s, L4, T.o10CS, 5.39,5.13,2.51,.83)
  txt(s, L31, T.a16bC, 9.15,4.76,2.51,.37)
  art(s, A.a37, 9.77,4.45,.2,.19, F.c18)
  art(s, A.a38, 10.84,4.45,.2,.19, F.c18)
  art(s, A.a37, 10.04,4.45,.2,.19, F.c18)
  art(s, A.a37, 10.32,4.45,.2,.19, F.c18)
  art(s, A.a37, 10.59,4.45,.2,.19, F.c18)
  txt(s, L4, T.o10CS, 9.16,5.13,2.51,.83)
  art(s, A.a37, 7.08,4.45,.2,.19, F.c18)
  mark(s, 2.66,3.34,.47,.47)
  shp(s, 'rect', 10.22,3.65,.08,.08, {fill:F.c14})
  shp(s, 'rect', 10.33,3.6,.08,.13, {fill:F.c14})
  shp(s, 'rect', 10.43,3.55,.08,.19, {fill:F.c14})
  shp(s, 'rect', 10.54,3.49,.08,.24, {fill:F.c14})
  art(s, A.a35, 10.2,3.35,.42,.25, F.c14)
  mark(s, 6.41,3.33,.5,.5)
  shp(s, 'roundRect', 1.93,.55,1.91,.67, {fill:F.c10,rectRadius:.2})
  logo(s, 2.21,.7, WHITE)
  shp(s, 'roundRect', 5.71,.55,1.91,.67, {fill:F.c10,rectRadius:.2})
  logo(s, 5.99,.7, WHITE)
  shp(s, 'roundRect', 9.46,.55,1.91,.67, {fill:F.c10,rectRadius:.2})
  logo(s, 9.75,.7, WHITE)
  bubbles(s, [
    [.69,.46,.188], [2.44,6.37,.332], [8.1,6.74,.452], [8.55,.25,.204], [4.65,1.72,.188], [11.85,4.11,.25]
  ])
}

function slide17(pptx) {
  const s = newSlide(pptx);
  deco16(s);
  txt(s, L1, T.r11bpN, 7.62,1.65,1.65,.29)
  txt(s, 'The Metaverse is the gateway to most digital experiences.', T.a24b, 7.62,1.93,5.13,.91)
  txt(s, L7, T.o10JS, 7.63,2.85,4.91,.58)
  txt(s, L8, T.o10, 7.61,4.08,2.07,.28)
  shp(s, 'rect', 7.7,4.5,4.31,.08, {fill:F.c10})
  txt(s, L19, T.a12b, 7.61,3.81,1.95,.3)
  shp(s, 'rect', 11.4,4.5,1.02,.08, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 11.82,4,.69,.28)
  txt(s, L8, T.o10, 7.61,5.03,2.07,.28)
  shp(s, 'rect', 7.7,5.46,4.43,.08, {fill:F.c10})
  txt(s, L20, T.a12b, 7.61,4.76,1.95,.3)
  shp(s, 'rect', 11.73,5.46,.69,.08, {fill:F.c14})
  txt(s, '95.2%', T.a10_5bR, 11.82,4.95,.69,.28)
  logo(s, 11.33,.38)
  shp(s, 'ellipse', 3.4,1.72,1.01,1.01, {fill:F.c10})
  shp(s, 'ellipse', 3.39,4.47,1.01,1.01, {fill:F.c10})
  txt(s, L21, T.a16bR, .82,1.86,2.51,.37)
  art(s, A.a37, 1.96,1.62,.2,.19, F.c18)
  art(s, A.a38, 3.04,1.62,.2,.19, F.c18)
  art(s, A.a37, 2.23,1.62,.2,.19, F.c18)
  art(s, A.a37, 2.51,1.62,.2,.19, F.c18)
  art(s, A.a37, 2.78,1.62,.2,.19, F.c18)
  txt(s, L4, T.o10RS, .83,2.19,2.51,.83)
  txt(s, L22, T.a16bC, 1.19,4.5,2.51,.37)
  art(s, A.a37, 1.96,4.2,.2,.19, F.c18)
  art(s, A.a37, 2.23,4.2,.2,.19, F.c18)
  art(s, A.a37, 2.51,4.2,.2,.19, F.c18)
  art(s, A.a37, 2.78,4.2,.2,.19, F.c18)
  txt(s, L4, T.o10RS, .83,4.85,2.51,.83)
  art(s, A.a37, 3.04,4.2,.2,.19, F.c18)
  mark(s, 3.67,1.99,.47,.47)
  mark(s, 3.64,4.73,.5,.5)
  bubbles(s, [
    [7.79,6.68,.152], [9.45,.6,.25], [11.11,6.03,.152], [.83,3.71,.189]
  ])
}

function slide18(pptx) {
  const s = newSlide(pptx);
  art(s, A.a39, 6.98,5.59,3.75,1.16, F.c19)
  art(s, A.a40, 4.19,3.85,2.12,1.62, F.c19)
  art(s, A.a41, 3.04,3.15,1.48,1.11, F.c19)
  art(s, A.a42, 2.1,2.56,.99,.72, F.c19)
  art(s, A.a43, 5.59,4.71,2.48,1.93, F.c19)
  art(s, A.a44, 7.49,5.91,.74,.21, F.c19)
  art(s, A.a45, 6.18,5.06,.65,.18, F.c19)
  art(s, A.a46, 4.73,4.26,.51,.14, F.c19)
  art(s, A.a47, 3.56,3.51,.36,.1, F.c19)
  conn(s, 'bentConnector2', 7.38,4.87,1.63,.66, {color:'FA7D02'}, {rotate:90,flipH:1,flipV:1})
  shp(s, 'rect', 8.52,4.13,.05,.5, {fill:F.c19})
  conn(s, 'bentConnector2', 4.63,3.24,1.45,.74, {color:'FA7D02'}, {rotate:90,flipH:1,flipV:1})
  shp(s, 'rect', 5.72,2.63,.05,.5, {fill:F.c19})
  conn(s, 'bentConnector2', 2.04,1.81,1.61,.51, {color:'FA7D02'}, {rotate:90,flipH:1,flipV:1})
  shp(s, 'rect', 3.1,1.01,.05,.5, {fill:F.c19})
  conn(s, 'bentConnector2', 5.52,4.48,.31,1.67, {color:'FA7D02'}, {rotate:90})
  shp(s, 'rect', 4.79,5.22,.05,.5, {fill:F.c19})
  conn(s, 'bentConnector2', 3.15,3.18,.25,.99, {color:'FA7D02'}, {rotate:90})
  shp(s, 'rect', 2.73,3.55,.05,.5, {fill:F.c19})
  art(s, A.a48, 7.34,5.84,1.04,.35, F.c14)
  art(s, A.a49, 6.05,5,.91,.3, F.c14)
  art(s, A.a50, 4.63,4.22,.72,.24, F.c14)
  art(s, A.a46, 3.49,3.47,.5,.16, F.c14)
  art(s, A.a51, 2.49,2.83,.2,.07, F.c14)
  shp(s, 'ellipse', 1.23,1.6,.99,.99, {fill:F.c19})
  txt(s, 'Our Service 04', T.a16bR, 2.1,5.01,2.51,.37)
  txt(s, L4, T.o10RS, 2.11,5.38,2.51,.83)
  txt(s, L22, T.a16bR, .15,3.32,2.51,.37)
  txt(s, L4, T.o10RS, .16,3.69,2.51,.83)
  txt(s, 'Our Service 05', T.a16b, 8.83,3.95,2.51,.37)
  txt(s, L4, T.o10S, 8.84,4.32,2.51,.83)
  txt(s, L31, T.a16b, 5.86,2.38,2.51,.37)
  txt(s, L4, T.o10S, 5.88,2.75,2.51,.83)
  txt(s, L21, T.a16b, 3.25,.73,2.51,.37)
  txt(s, L4, T.o10S, 3.26,1.1,2.51,.83)
  txt(s, L1, T.r11bRpN, 11.27,.49,1.65,.29)
  txt(s, 'Chevron Milestones Diagram', T.a32bR, 8.37,.75,4.54,1.18)
  txt(s, L4, T.o10RS, 8.37,2.08,4.54,.58)
  art(s, A.a22, 9.39,2.24,5.62,3.03, F.c7, {rotate:270,flipH:1,flipV:1})
  art(s, A.a23, 10.52,3.75,3.52,2.81, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a24, 10.58,3.83,3.4,2.8, F.c5, {rotate:270,flipH:1,flipV:1})
  art(s, A.a15, 11.2,4.65,2.63,2.19, F.c8, {rotate:270,flipH:1,flipV:1})
  art(s, A.a16, 12.18,5.96,1.39,1.22, F.c2, {rotate:90,flipH:1,flipV:1})
  logo(s, .41,.45)
  bubbles(s, [
    [6.99,.73,.25], [.68,6.12,.332], [3.93,7,.206], [4.44,2.28,.163], [7.37,3.94,.163]
  ])
}

function slide19(pptx) {
  const s = newSlide(pptx);
  art(s, A.a17, -.01,5.43,10.08,2.08, F.c7, {rotate:180})
  art(s, A.a2, .01,5.54,7.26,1.97, F.c2, {rotate:180})
  art(s, A.a2, .01,5.54,7.06,1.98, F.c5, {rotate:180})
  art(s, A.a2, .01,5.85,6.36,1.67, F.c8, {rotate:180})
  art(s, A.a5, -.01,6.47,3.04,1.04, F.c2)
  art(s, A.a52, .71,0,12.62,7.32, F.c19)
  shp(s, 'ellipse', 1.06,2.7,1.92,1.92, {fill:F.c19})
  shp(s, 'ellipse', 3.53,1.27,1.92,1.92, {fill:F.c19})
  shp(s, 'ellipse', 5.98,3.83,1.92,1.92, {fill:F.c19})
  shp(s, 'ellipse', 8.44,2.42,1.92,1.92, {fill:F.c19})
  shp(s, 'ellipse', 10.89,5,1.92,1.92, {fill:F.c19})
  txt(s, 'Section 1', T.a16bC, .76,1.22,2.51,.37)
  txt(s, L13, T.o10CS, .99,1.68,2.09,.83)
  txt(s, 'Section 2', T.a16bC, 3.22,3.39,2.51,.37)
  txt(s, L13, T.o10CS, 3.45,3.85,2.09,.83)
  txt(s, 'Section 3', T.a16bC, 5.67,2.2,2.51,.37)
  txt(s, L13, T.o10CS, 5.9,2.67,2.09,.83)
  txt(s, 'Section 4', T.a16bC, 8.12,4.5,2.51,.37)
  txt(s, L13, T.o10CS, 8.34,4.96,2.09,.83)
  txt(s, 'Section 5', T.a16bC, 10.66,3.39,2.51,.37)
  txt(s, L13, T.o10CS, 10.89,3.85,2.09,.83)
  art(s, A.a53, 8.89,2.89,.97,.97, F.c14)
  art(s, A.a54, 9.17,3.06,.42,.07, F.c14)
  art(s, A.a55, 9.17,3.17,.34,.07, F.c14)
  art(s, A.a56, 9.16,3.28,.26,.07, F.c14)
  art(s, A.a57, 6.48,4.31,.99,.97, F.c14)
  art(s, A.a58, 11.49,5.49,.73,.94, F.c14)
  art(s, A.a59, 11.48,5.4,.51,.78, F.c14)
  art(s, A.a60, 11.78,5.74,.27,.03, F.c14)
  art(s, A.a60, 11.78,5.81,.27,.03, F.c14)
  art(s, A.a61, 11.78,5.93,.27,.03, F.c14)
  art(s, A.a61, 11.78,6,.27,.03, F.c14)
  art(s, A.a61, 11.78,6.13,.27,.03, F.c14)
  art(s, A.a61, 11.78,6.19,.27,.03, F.c14)
  art(s, A.a62, 11.59,5.71,.19,.15, F.c14)
  art(s, A.a63, 11.59,5.91,.19,.15, F.c14)
  art(s, A.a64, 11.59,6.1,.19,.15, F.c14)
  art(s, A.a65, 4.42,2.06,.19,.18, F.c14)
  art(s, A.a66, 4.01,1.66,1.02,.98, F.c14)
  art(s, A.a67, 1.45,3.12,1.03,.55, F.c14)
  art(s, A.a68, 1.45,3.41,1.03,.73, F.c14)
  art(s, A.a69, 1.67,3.33,.58,.41, F.c14)
  art(s, A.a1, 5.48,-.01,7.87,2.38, F.c4)
  art(s, A.a2, 7.67,-.01,5.67,2.25, F.c2)
  art(s, A.a2, 7.98,-.01,5.36,2.18, F.c5)
  art(s, A.a2, 9.17,-.02,4.17,2.03, F.c6)
  art(s, A.a3, 10.53,-.01,2.82,1.37, F.c6)
  txt(s, L1, T.r11bN, .84,5.55,1.65,.29)
  txt(s, 'Wire Loop Diagram ', T.a32b, .84,5.83,4.75,.64)
  txt(s, L13, T.o10S, .86,6.47,4.2,.58)
  bubbles(s, [
    [2.79,.6,.25], [3.64,5.23,.332], [11.21,1.72,.454], [7.1,1.25,.163]
  ])
}

function slide20(pptx) {
  const s = newSlide(pptx);
  art(s, A.a70, 6.48,1.4,1.37,4.34, F.c3)
  art(s, A.a71, 7.17,1.43,.68,4.3, F.c3)
  art(s, A.a70, 7.49,3.47,1.37,2.26, F.c3)
  art(s, A.a71, 8.18,3.49,.68,2.24, F.c3)
  art(s, A.a70, 8.5,2.63,1.37,3.1, F.c3)
  art(s, A.a71, 9.19,2.66,.68,3.07, F.c3)
  art(s, A.a70, 9.52,4.25,1.37,1.48, F.c3)
  art(s, A.a71, 10.21,4.27,.68,1.47, F.c3)
  art(s, A.a70, 10.53,1.81,1.37,3.92, F.c3)
  art(s, A.a71, 11.22,1.84,.68,3.89, F.c3)
  art(s, A.a70, 11.54,2.28,1.37,3.45, F.c3)
  art(s, A.a71, 12.23,2.31,.68,3.43, F.c3)
  txt(s, '2020', T.a16bCN, 6.76,5.88,.73,.37)
  txt(s, '2021', T.a16bCN, 7.8,5.88,.73,.37)
  txt(s, '2022', T.a16bCN, 8.82,5.88,.73,.37)
  txt(s, '2023', T.a16bCN, 9.84,5.88,.73,.37)
  txt(s, '2024', T.a16bCN, 10.87,5.88,.73,.37)
  txt(s, '2025', T.a16bCN, 11.87,5.88,.73,.37)
  txt(s, '82%', T.a12bCM, 6.79,.98,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 7.02,1.43,.29,.15, {fill:F.c19,rotate:180})
  txt(s, '46%', T.a12bCM, 7.82,2.91,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 8.04,3.36,.29,.15, {fill:F.c19,rotate:180})
  txt(s, '70%', T.a12bCM, 8.83,2.09,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 9.05,2.54,.29,.15, {fill:F.c19,rotate:180})
  txt(s, '75%', T.a12bCM, 10.85,1.25,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 11.08,1.69,.29,.15, {fill:F.c19,rotate:180})
  txt(s, '72%', T.a12bCM, 11.86,1.71,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 12.09,2.16,.29,.15, {fill:F.c19,rotate:180})
  txt(s, '30%', T.a12bCM, 9.83,3.68,.74,.45, {fill:F.c19})
  shp(s, 'triangle', 10.06,4.13,.29,.15, {fill:F.c19,rotate:180})
  logo(s, .41,.45)
  txt(s, L1, T.r11bpN, 1,1.64,1.65,.29)
  txt(s, 'Rodeix Statistic Diagtam', T.a29b, 1,1.92,5.42,.59)
  txt(s, L7, T.o10JS, 1.02,2.56,4.91,.58)
  txt(s, L8, T.o10, 1,3.72,2.07,.28)
  shp(s, 'rect', 1.09,4.15,4.31,.08, {fill:F.c10})
  txt(s, L19, T.a12b, 1,3.45,1.95,.3)
  shp(s, 'rect', 4.79,4.15,1.02,.08, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 5.21,3.64,.69,.28)
  txt(s, L8, T.o10, 1,4.67,2.07,.28)
  shp(s, 'rect', 1.09,5.1,4.43,.08, {fill:F.c10})
  txt(s, L20, T.a12b, 1,4.41,1.95,.3)
  shp(s, 'rect', 5.12,5.1,.69,.08, {fill:F.c14})
  txt(s, '95.2%', T.a10_5bR, 5.21,4.6,.69,.28)
  art(s, A.a6, 1.72,-.01,11.62,1.08, F.c4)
  art(s, A.a2, 4.96,-.01,8.37,1.02, F.c2)
  art(s, A.a2, 5.41,-.01,7.91,.99, F.c5)
  art(s, A.a2, 7.17,-.02,6.15,.92, F.c6)
  art(s, A.a2, 9.19,-.01,4.16,.63, F.c6)
  art(s, A.a6, -.04,5.73,13.39,1.78, F.c4, {rotate:180})
  art(s, A.a7, -.02,5.83,9.64,1.68, F.c2, {rotate:180})
  art(s, A.a7, -.02,5.88,9.11,1.63, F.c5, {rotate:180})
  art(s, A.a2, -.02,6,7.09,1.52, F.c6, {rotate:180})
  art(s, A.a2, -.05,6.48,4.79,1.03, F.c6, {rotate:180})
  bubbles(s, [
    [5.85,6.18,.217], [9.99,.72,.169], [.67,3.59,.169], [3.65,1.05,.095]
  ])
}

function slide21(pptx) {
  const s = newSlide(pptx);
  art(s, A.a6, 1.72,-.01,11.62,1.08, F.c4)
  art(s, A.a2, 4.96,-.01,8.37,1.02, F.c2)
  art(s, A.a2, 5.41,-.01,7.91,.99, F.c5)
  art(s, A.a2, 7.17,-.02,6.15,.92, F.c6)
  art(s, A.a2, 9.19,-.01,4.16,.63, F.c6)
  art(s, A.a72, 1.61,4.54,1.02,1.4, F.c20)
  art(s, A.a73, .43,1.37,3.38,1.55, F.c21)
  art(s, A.a74, 1.05,3.01,2.15,1.45, F.c22)
  art(s, A.a75, .43,1.57,3.38,1.35, F.c19)
  art(s, A.a76, .43,1.37,3.38,.4, F.c23)
  art(s, A.a77, .86,2.69,2.54,.53, F.c24)
  art(s, A.a78, 1.05,3.01,2.15,.4, F.c23)
  art(s, A.a79, 1.06,3.23,2.14,1.23, F.c19)
  art(s, A.a80, 1.46,4.3,1.33,.4, F.c24)
  art(s, A.a81, 1.61,4.54,1.02,.33, F.c23)
  art(s, A.a82, 1.63,4.75,.99,1.2, F.c19)
  shp(s, 'ellipse', 3.78,2.28,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  conn(s, 'straightConnector1', 2.99,2.33,.79,0, {color:WHITE}, {flipH:1})
  shp(s, 'ellipse', 2.88,2.28,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  shp(s, 'ellipse', 3.79,3.79,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  conn(s, 'straightConnector1', 2.69,3.84,1.1,0, {color:WHITE}, {flipH:1})
  shp(s, 'ellipse', 2.58,3.79,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  shp(s, 'ellipse', 3.82,5.39,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  conn(s, 'straightConnector1', 2.31,5.44,1.51,0, {color:WHITE}, {flipH:1})
  shp(s, 'ellipse', 2.2,5.39,.11,.11, {fill:F.c25,line:{color:WHITE},rotate:180})
  txt(s, L32, T.a16b, 4.43,1.77,2.51,.37)
  txt(s, L4, T.o10S, 4.45,2.14,2.51,.83)
  txt(s, L33, T.a16b, 4.45,3.26,2.51,.37)
  txt(s, L4, T.o10S, 4.46,3.63,2.51,.83)
  txt(s, L34, T.a16b, 4.46,4.79,2.51,.37)
  txt(s, L4, T.o10S, 4.46,5.16,2.51,.83)
  txt(s, L1, T.r11bpN, 7.49,1.09,1.65,.29)
  txt(s, 'Infographic Funnel w/ 3 Stages', T.a32b, 7.49,1.37,5.42,1.18)
  txt(s, L3, T.o10JS, 8.05,3.11,4.4,.58)
  txt(s, L2, T.a12b, 8.04,2.81,1.51,.3)
  txt(s, L3, T.o10JS, 8.05,4.27,4.4,.58)
  txt(s, L5, T.a12b, 8.04,3.97,1.51,.3)
  mark(s, 7.57,2.88,.47,.47)
  mark(s, 7.57,4.09,.47,.47)
  txt(s, L3, T.o10JS, 8.05,5.43,4.4,.58)
  txt(s, L12, T.a12b, 8.04,5.13,1.51,.3)
  logo(s, .43,.39)
  shp(s, 'rect', 7.62,5.5,.07,.07, {fill:F.c10})
  shp(s, 'rect', 7.72,5.45,.07,.12, {fill:F.c10})
  shp(s, 'rect', 7.82,5.4,.07,.17, {fill:F.c10})
  shp(s, 'rect', 7.92,5.35,.07,.22, {fill:F.c10})
  art(s, A.a6, -.04,5.73,13.39,1.78, F.c4, {rotate:180})
  art(s, A.a7, -.02,5.83,9.64,1.68, F.c2, {rotate:180})
  art(s, A.a7, -.02,5.88,9.11,1.63, F.c5, {rotate:180})
  art(s, A.a2, -.02,6,7.09,1.52, F.c6, {rotate:180})
  art(s, A.a2, -.05,6.48,4.79,1.03, F.c6, {rotate:180})
  bubbles(s, [
    [5.6,.72,.369], [11.18,2.38,.169], [6.94,3.12,.148], [12.74,6.15,.169], [4.71,6.33,.169]
  ])
}

function slide22(pptx) {
  const s = newSlide(pptx);
  art(s, A.a83, .76,2.27,3.58,3.03, F.c26)
  art(s, A.a84, .76,2.28,3.58,3.02, F.c26)
  art(s, A.a85, 1.83,.89,3.58,3.03, F.c26)
  art(s, A.a86, 1.83,.89,3.58,3.02, F.c26)
  art(s, A.a87, 2.63,2.86,.95,.69, F.c19)
  art(s, A.a88, 2.88,2.63,.45,.27, F.c19)
  art(s, A.a6, 1.72,-.01,11.62,1.08, F.c4)
  art(s, A.a2, 4.96,-.01,8.37,1.02, F.c2)
  art(s, A.a2, 5.41,-.01,7.91,.99, F.c5)
  art(s, A.a2, 7.17,-.02,6.15,.92, F.c6)
  art(s, A.a2, 9.19,-.01,4.16,.63, F.c6)
  txt(s, L1, T.r11bpN, 6.67,1.67,1.65,.29)
  txt(s, 'Curved Right & Left Arrows ', T.a36b, 6.67,1.95,5.42,1.31)
  txt(s, L3, T.o10JS, 7.22,3.85,4.4,.58)
  txt(s, L2, T.a12b, 7.21,3.54,1.51,.3)
  txt(s, L3, T.o10JS, 7.22,5.01,4.4,.58)
  txt(s, L5, T.a12b, 7.21,4.71,1.51,.3)
  mark(s, 6.74,3.62,.47,.47)
  mark(s, 6.74,4.83,.47,.47)
  logo(s, .43,.39)
  art(s, A.a6, -.04,5.58,13.39,1.93, F.c4, {rotate:180})
  art(s, A.a7, -.02,5.68,9.64,1.82, F.c2, {rotate:180})
  art(s, A.a7, -.02,5.74,9.11,1.76, F.c5, {rotate:180})
  art(s, A.a2, -.02,5.87,7.09,1.65, F.c6, {rotate:180})
  art(s, A.a2, -.05,6.39,4.79,1.11, F.c6, {rotate:180})
  bubbles(s, [
    [5.6,.72,.369], [11.42,2.97,.169], [5.36,3.65,.148], [12.74,6.15,.169], [4.71,6.33,.169]
  ])
}

function slide23(pptx) {
  const s = newSlide(pptx);
  art(s, A.a89, .52,4.56,12.3,2.59, F.c10)
  shp(s, 'roundRect', 1.23,5.28,1.15,1.15, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.03,rotate:45})
  shp(s, 'roundRect', 3.66,5.28,1.15,1.15, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.03,rotate:45})
  shp(s, 'roundRect', 6.09,5.28,1.15,1.15, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.03,rotate:45})
  shp(s, 'roundRect', 8.52,5.28,1.15,1.15, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.03,rotate:45})
  shp(s, 'roundRect', 10.94,5.28,1.15,1.15, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.03,rotate:45})
  mark(s, 11.31,5.6,.42,.42)
  mark(s, 1.6,5.6,.42,.42)
  mark(s, 4.03,5.6,.42,.42)
  mark(s, 6.45,5.6,.42,.42)
  mark(s, 8.88,5.6,.42,.42)
  txt(s, '1', T.a16bCM, 1.48,2.25,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.a14bN, 2.06,2.21,1.37,.34)
  txt(s, L6, T.o10JS, 2.06,2.47,2.52,.6)
  txt(s, '2', T.a16bCM, 1.48,3.37,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L5, T.a14bN, 2.06,3.33,1.36,.34)
  txt(s, L6, T.o10JS, 2.06,3.59,2.52,.6)
  txt(s, '4', T.a16bCM, 5.02,2.29,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, 'Section Four', T.a14bN, 5.6,2.25,1.44,.34)
  txt(s, L6, T.o10JS, 5.6,2.51,2.52,.6)
  txt(s, '5', T.a16bCM, 5.02,3.37,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, 'Section Five', T.a14bN, 5.6,3.33,1.4,.34)
  txt(s, L6, T.o10JS, 5.6,3.59,2.52,.6)
  txt(s, '3', T.a16bCM, 8.76,2.25,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L12, T.a14bN, 9.34,2.21,1.55,.34)
  txt(s, L6, T.o10JS, 9.34,2.47,2.52,.6)
  txt(s, '6', T.a16bCM, 8.76,3.41,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, 'Section Six', T.a14bN, 9.34,3.37,1.29,.34)
  txt(s, L6, T.o10JS, 9.34,3.64,2.52,.6)
  txt(s, L1, T.r11bCpN, 5.67,.35,1.65,.29)
  txt(s, 'Rhombus Process Diagram ', T.a36bC, 2.53,.63,7.92,.71)
  txt(s, L11, T.o10CS, 2.38,1.33,8.23,.58)
  art(s, A.a22, -1.68,2.24,5.62,3.03, F.c7, {rotate:90,flipH:1,flipV:1})
  art(s, A.a23, -.71,.94,3.52,2.81, F.c2, {rotate:90,flipH:1,flipV:1})
  art(s, A.a24, -.65,.87,3.4,2.8, F.c5, {rotate:90,flipH:1,flipV:1})
  art(s, A.a15, -.5,.66,2.63,2.19, F.c8, {rotate:90,flipH:1,flipV:1})
  art(s, A.a16, -.23,.31,1.39,1.22, F.c2, {rotate:270,flipH:1,flipV:1})
  art(s, A.a22, 9.35,2.24,5.71,3.03, F.c7, {rotate:270,flipV:1})
  art(s, A.a23, 10.5,.94,3.57,2.81, F.c2, {rotate:270,flipV:1})
  art(s, A.a24, 10.56,.87,3.45,2.8, F.c5, {rotate:270,flipV:1})
  art(s, A.a15, 11.18,.66,2.67,2.19, F.c8, {rotate:270,flipV:1})
  art(s, A.a16, 12.18,.31,1.41,1.22, F.c2, {rotate:90,flipV:1})
}

function slide24(pptx) {
  const s = newSlide(pptx);
  art(s, A.a90, 2.28,3.72,1.17,1.49, F.c27, {rotate:240})
  art(s, A.a90, 4.82,3.74,1.17,1.49, F.c27, {rotate:300})
  art(s, A.a90, 7.37,3.72,1.17,1.49, F.c27, {rotate:240})
  art(s, A.a90, 9.91,3.74,1.17,1.49, F.c27, {rotate:300})
  shp(s, 'ellipse', .54,4.19,2.04,2.04, {fill:F.c10})
  shp(s, 'ellipse', 1.02,4.67,1.09,1.09, {fill:F.c28})
  shp(s, 'ellipse', 1.16,4.81,.81,.81, {fill:F.c10,shadow:SHADOW})
  shp(s, 'ellipse', 3.09,2.72,2.04,2.04, {fill:F.c10})
  shp(s, 'ellipse', 3.57,3.2,1.09,1.09, {fill:F.c28})
  shp(s, 'ellipse', 3.71,3.34,.81,.81, {fill:F.c10,shadow:SHADOW})
  shp(s, 'ellipse', 5.65,4.19,2.04,2.04, {fill:F.c10})
  shp(s, 'ellipse', 6.12,4.67,1.09,1.09, {fill:F.c28})
  shp(s, 'ellipse', 6.26,4.81,.81,.81, {fill:F.c10,shadow:SHADOW})
  shp(s, 'ellipse', 8.2,2.72,2.04,2.04, {fill:F.c10})
  shp(s, 'ellipse', 8.68,3.2,1.09,1.09, {fill:F.c28})
  shp(s, 'ellipse', 8.82,3.34,.81,.81, {fill:F.c10,shadow:SHADOW})
  shp(s, 'ellipse', 10.75,4.19,2.04,2.04, {fill:F.c10})
  shp(s, 'ellipse', 11.23,4.67,1.09,1.09, {fill:F.c28})
  shp(s, 'ellipse', 11.37,4.81,.81,.81, {fill:F.c10,shadow:SHADOW})
  mark(s, 11.51,4.93,.53,.53)
  mark(s, 8.94,3.49,.53,.53)
  mark(s, 3.85,3.48,.53,.53)
  mark(s, 1.3,4.95,.53,.53)
  mark(s, 6.4,4.94,.53,.53)
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididun', T.o10CS, 3.13,1.85,6.72,.34)
  txt(s, '1', T.a16bCM, .56,2.79,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.a14bN, 1.14,2.74,1.37,.34)
  txt(s, L6, T.o10JS, 1.14,3.01,1.76,.83)
  txt(s, '2', T.a16bCM, 2.89,5.15,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.w14bN, 3.47,5.11,1.37,.34)
  txt(s, L6, T.o10JS, 3.47,5.37,1.76,.83)
  txt(s, '4', T.a16bCM, 8.02,5.15,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.a14bN, 8.6,5.11,1.37,.34)
  txt(s, L6, T.o10JS, 8.6,5.37,1.76,.83)
  txt(s, '3', T.a16bCM, 5.29,2.79,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.a14bN, 5.87,2.74,1.37,.34)
  txt(s, L6, T.o10JS, 5.87,3.01,1.76,.83)
  txt(s, '5', T.a16bCM, 10.42,2.79,.52,.52, {shape:'ellipse',fill:F.c10})
  txt(s, L2, T.a14bN, 11,2.74,1.37,.34)
  txt(s, L6, T.o10JS, 11,3.01,1.76,.83)
  txt(s, L1, T.r11bCpN, 5.67,.85,1.65,.29)
  txt(s, 'Chain Process Diagram ', T.a36bC, 2.53,1.13,7.92,.71)
  art(s, A.a6, 0,-.01,13.35,1.86, F.c4)
  art(s, A.a7, 3.71,-.01,9.61,1.76, F.c2)
  art(s, A.a7, 4.24,-.01,9.08,1.71, F.c5)
  art(s, A.a2, 6.26,-.02,7.06,1.59, F.c6)
  art(s, A.a2, 8.57,-.01,4.78,1.08, F.c6)
  art(s, A.a6, -.04,6.59,13.39,.92, F.c4, {rotate:180})
  art(s, A.a7, -.02,6.64,9.64,.87, F.c2, {rotate:180})
  art(s, A.a7, -.02,6.67,9.11,.85, F.c5, {rotate:180})
  art(s, A.a2, -.02,6.73,7.09,.79, F.c6, {rotate:180})
  art(s, A.a2, -.05,6.98,4.79,.53, F.c6, {rotate:180})
}

function slide25(pptx) {
  const s = newSlide(pptx);
  art(s, A.a1, 8.6,2.89,7.62,1.83, F.c4, {rotate:270,flipV:1})
  art(s, A.a2, 9.72,1.88,5.49,1.73, F.c2, {rotate:270,flipV:1})
  art(s, A.a2, 9.9,1.76,5.19,1.67, F.c5, {rotate:270,flipV:1})
  art(s, A.a11, 10.54,1.24,4.04,1.56, F.c6, {rotate:270,flipV:1})
  art(s, A.a3, 11.43,.82,2.73,1.05, F.c6, {rotate:270,flipV:1})
  art(s, A.a91, 3.67,-.33,4.03,9.34, F.c1, {rotate:253})
  art(s, A.a92, 10.33,2.06,1.74,.96, F.c27, {rotate:330,flipH:1})
  art(s, A.a93, 9.63,2.25,2.31,.29, F.c10, {rotate:330,flipH:1})
  art(s, A.a94, 10.54,2.85,.51,.45, F.c27, {rotate:330,flipH:1})
  art(s, A.a95, 10.69,1.97,1.45,1.32, F.c10, {rotate:330,flipH:1})
  txt(s, '3', T.a28bCM, 5.81,3.75,.74,.74, {shape:'ellipse',fill:F.c10})
  txt(s, '1', T.a28bCM, .53,3.19,.74,.74, {shape:'ellipse',fill:F.c10})
  txt(s, '2', T.a28bCM, 4.48,6.29,.74,.74, {shape:'ellipse',fill:F.c10})
  txt(s, '4', T.a28bCM, 10.04,3.44,.74,.74, {shape:'ellipse',fill:F.c10})
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusm od tempor incididunt ut labore et dolore magna aliqua. ', T.o10S, .74,1.76,9.13,.34)
  txt(s, L32, T.a18b, 1.42,3.1,2.68,.4)
  txt(s, L16, T.o10S, 1.42,3.5,3.29,.6)
  txt(s, L34, T.a18b, 5.46,2.53,2.68,.4)
  txt(s, L16, T.o10S, 5.46,2.93,3.29,.6)
  txt(s, L33, T.a18b, 6.22,6.28,2.68,.4)
  txt(s, L16, T.o10S, 6.22,6.68,3.29,.6)
  txt(s, 'Section 04', T.a18b, 9.23,4.59,2.68,.4)
  txt(s, L16, T.o10S, 9.23,4.99,3.29,.6)
  txt(s, L1, T.r11bpN, .73,.79,1.65,.29)
  txt(s, 'Paper airplane and trail infographic', T.a32b, .73,1.07,8.39,.64)
  art(s, A.a6, -.05,-.02,13.39,1.1, F.c4)
  art(s, A.a7, 3.68,-.02,9.64,1.04, F.c2)
  art(s, A.a7, 4.21,-.02,9.11,1, F.c5)
  art(s, A.a2, 6.23,-.03,7.09,.94, F.c6)
  art(s, A.a2, 8.55,-.02,4.79,.63, F.c6)
  bubbles(s, [
    [9.52,1.26,.239], [3.53,2.76,.148], [4.85,5.62,.148], [8.9,4.1,.148], [.83,6.68,.148]
  ])
}

function slide26(pptx) {
  const s = newSlide(pptx);
  deco18(s);
  txt(s, L1, T.r11bpN, 2.52,2.2,1.65,.29)
  txt(s, 'Strengths Analysis Slides', T.a60b, 2.52,2.48,6.62,2.12)
  txt(s, L7, T.o10JS, 2.54,4.66,6.09,.58)
  txt(s, 'S', T.a199bC, 10.66,3.73,2.41,3.45)
  bubbles(s, [
    [9.05,1.26,.315], [.45,5.54,.148], [1.76,3.39,.241], [5.61,5.93,.148], [7.31,2.74,.148]
  ])
}

function slide27(pptx) {
  const s = newSlide(pptx);
  deco19(s);
  txt(s, 'W', T.a199bC, .45,3.73,2.41,3.45)
  txt(s, L1, T.r11bRpN, 9.05,2.2,1.65,.29)
  txt(s, 'Weaknesses Analysis Slides', T.a60bR, 4.16,2.48,6.62,2.12)
  txt(s, L7, T.o10RS, 4.61,4.66,6.09,.58)
  bubbles(s, [
    [4.7,2.48,.315], [12.29,6.08,.148], [12.12,3.63,.241], [5.61,5.93,.148], [7.39,1.64,.148], [2.63,2.33,.148]
  ])
}

function slide28(pptx) {
  const s = newSlide(pptx);
  deco20(s);
  txt(s, 'O', T.a199bC, 10.66,.13,2.41,3.45)
  txt(s, L1, T.r11bpN, 2.52,1.9,1.65,.29)
  txt(s, 'Opportunities Analysis Slides', T.a60b, 2.52,2.18,6.62,2.12)
  txt(s, L7, T.o10JS, 2.54,4.36,6.09,.58)
  bubbles(s, [
    [9.97,5.3,.315], [.47,1.63,.148], [1.76,3.39,.241], [5.61,5.93,.148], [5.31,1.72,.148]
  ])
}

function slide29(pptx) {
  const s = newSlide(pptx);
  deco21(s);
  txt(s, 'T', T.a199bC, .45,.13,2.41,3.45)
  txt(s, L1, T.r11bRpN, 9.05,1.9,1.65,.29)
  txt(s, 'Threats Analysis Slides', T.a60bR, 4.16,2.18,6.62,2.12)
  txt(s, L7, T.o10RS, 4.61,4.36,6.09,.58)
  bubbles(s, [
    [5.02,2.18,.315], [1.73,5.93,.148], [12.05,1.47,.241], [5.64,5.39,.148], [6.59,2.86,.148], [11.6,3.75,.148]
  ])
}

function slide30(pptx) {
  const s = newSlide(pptx);
  deco22(s);
  shp(s, 'rect', 5.74,3.34,7.13,2.97, {fill:F.c15})
  txt(s, L17, T.a14bC, 5.65,5.17,2.51,.34)
  txt(s, 'Team Rodiex One', T.o10CS, 5.96,5.44,1.88,.33)
  txt(s, L18, T.a14bC, 8.08,5.17,2.51,.34)
  txt(s, 'Team Rodiex Two', T.o10CS, 8.4,5.44,1.88,.33)
  txt(s, L14, T.a14bC, 10.41,5.17,2.51,.34)
  txt(s, 'Team Rodiex Three', T.o10CS, 10.72,5.44,1.88,.33)
  txt(s, L1, T.r11bpN, .78,1.35,1.65,.29)
  txt(s, 'Our Team Rodiex', T.a36b, .78,1.63,4.39,.71)
  txt(s, L15, T.o10JS, 1.34,3.41,3.75,.33)
  txt(s, L17, T.a12b, 1.33,3.11,1.51,.3)
  txt(s, L15, T.o10JS, 1.34,4.25,3.75,.33)
  txt(s, L18, T.a12b, 1.33,3.95,1.51,.3)
  mark(s, .85,3.19,.47,.47)
  mark(s, .85,4.08,.47,.47)
  txt(s, L15, T.o10JS, 1.34,5.18,3.75,.33)
  txt(s, L14, T.a12b, 1.33,4.88,1.51,.3)
  shp(s, 'ellipse', 6.03,2.97,.17,.17, {fill:F.c29})
  shp(s, 'rect', .91,5.24,.07,.07, {fill:F.c10})
  shp(s, 'rect', 1.01,5.19,.07,.12, {fill:F.c10})
  shp(s, 'rect', 1.11,5.14,.07,.17, {fill:F.c10})
  shp(s, 'rect', 1.21,5.09,.07,.22, {fill:F.c10})
  txt(s, L35, T.o10JS, .84,2.35,4.24,.58)
  logo(s, .43,.39)
}

function slide31(pptx) {
  const s = newSlide(pptx);
  deco23(s);
  txt(s, L17, T.a14b, 3.47,1.31,2.51,.34)
  txt(s, L15, T.o10JS, 3.47,1.64,2.23,.58)
  art(s, A.a37, 3.56,1.06,.16,.15, F.c18)
  art(s, A.a37, 3.83,1.06,.16,.15, F.c18)
  art(s, A.a37, 4.11,1.06,.16,.15, F.c18)
  art(s, A.a37, 4.38,1.06,.16,.15, F.c18)
  art(s, A.a37, 4.64,1.05,.16,.15, F.c18)
  txt(s, L14, T.a14b, 3.47,3.47,2.51,.34)
  txt(s, L15, T.o10JS, 3.47,3.79,2.23,.58)
  art(s, A.a37, 3.56,3.21,.16,.15, F.c18)
  art(s, A.a37, 3.83,3.21,.16,.15, F.c18)
  art(s, A.a37, 4.11,3.21,.16,.15, F.c18)
  art(s, A.a37, 4.38,3.21,.16,.15, F.c18)
  art(s, A.a37, 4.64,3.2,.16,.15, F.c18)
  txt(s, L36, T.a14b, 3.47,5.68,2.34,.34)
  txt(s, L15, T.o10JS, 3.47,6.01,2.23,.58)
  art(s, A.a37, 3.56,5.42,.16,.15, F.c18)
  art(s, A.a37, 3.83,5.42,.16,.15, F.c18)
  art(s, A.a37, 4.11,5.42,.16,.15, F.c18)
  art(s, A.a37, 4.38,5.42,.16,.15, F.c18)
  art(s, A.a37, 4.64,5.42,.16,.15, F.c18)
  txt(s, L1, T.r11bpN, 7.27,1.4,1.65,.29)
  txt(s, 'All Our Team Rodiex', T.a36b, 7.27,1.68,5.27,.71)
  txt(s, L23, T.o10JS, 7.83,3.47,4.53,.33)
  txt(s, L17, T.a12b, 7.82,3.17,1.51,.3)
  txt(s, L23, T.o10JS, 7.83,4.31,4.5,.33)
  txt(s, L14, T.a12b, 7.82,4.01,1.51,.3)
  mark(s, 7.35,3.25,.47,.47)
  mark(s, 7.35,4.13,.47,.47)
  txt(s, L23, T.o10JS, 7.83,5.24,4.53,.33)
  txt(s, L36, T.a12b, 7.82,4.93,1.51,.3)
  shp(s, 'rect', 7.41,5.3,.07,.07, {fill:F.c10})
  shp(s, 'rect', 7.51,5.25,.07,.12, {fill:F.c10})
  shp(s, 'rect', 7.61,5.2,.07,.17, {fill:F.c10})
  shp(s, 'rect', 7.7,5.15,.07,.22, {fill:F.c10})
  txt(s, L35, T.o10JS, 7.34,2.4,5,.58)
  logo(s, 11.49,.39)
}

function slide32(pptx) {
  const s = newSlide(pptx);
  deco24(s);
  txt(s, 'Testimonials One', T.a14b, 8.94,2.27,2.51,.34)
  rich(s, [['“',T.o10iJS],[L37,T.o10iJS],['”',T.o10iJS]], 8.94,2.6,3.21,.58)
  art(s, A.a37, 9.03,2.01,.16,.15, F.c18)
  art(s, A.a37, 9.31,2.01,.16,.15, F.c18)
  art(s, A.a37, 9.58,2.01,.16,.15, F.c18)
  art(s, A.a37, 9.86,2.01,.16,.15, F.c18)
  art(s, A.a37, 10.11,2.01,.16,.15, F.c18)
  txt(s, 'Testimonials Two', T.a14b, 8.94,4.67,2.51,.34)
  art(s, A.a37, 9.03,4.41,.16,.15, F.c18)
  art(s, A.a37, 9.31,4.41,.16,.15, F.c18)
  art(s, A.a37, 9.58,4.41,.16,.15, F.c18)
  art(s, A.a37, 9.86,4.41,.16,.15, F.c18)
  art(s, A.a37, 10.11,4.41,.16,.15, F.c18)
  rich(s, [['“',T.o10iJS],[L37,T.o10iJS],['”',T.o10iJS]], 8.94,5.01,3.21,.58)
  txt(s, L1, T.r11bpN, 1.24,1.81,1.65,.29)
  txt(s, L26, T.a28b, 1.24,2.09,5.42,1.04)
  txt(s, L3, T.o10JS, 1.8,3.74,4.59,.58)
  txt(s, L2, T.a12b, 1.79,3.44,1.51,.3)
  txt(s, L3, T.o10JS, 1.8,4.9,4.59,.58)
  txt(s, L5, T.a12b, 1.79,4.6,1.51,.3)
  logo(s, .41,.45)
  shp(s, 'ellipse', 5.52,6.07,.28,.28, {fill:F.c3})
  shp(s, 'ellipse', 5.81,1.64,.17,.17, {fill:F.c3})
  shp(s, 'ellipse', .58,4.43,.17,.17, {fill:F.c3})
  shp(s, 'ellipse', 2.65,1.23,.1,.1, {fill:F.c3})
  mark(s, 1.32,3.52,.47,.47)
  mark(s, 1.32,4.73,.47,.47)
}

function slide33(pptx) {
  const s = newSlide(pptx);
  deco25(s);
  shp(s, 'rect', 0,-.02,13.33,5.25, {fill:F.c30})
  shp(s, 'roundRect', .71,4.48,3.71,2.32, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.34})
  shp(s, 'roundRect', 5,4.48,3.71,2.32, {fill:F.c1,line:{color:WHITE,width:6},rectRadius:.34})
  shp(s, 'roundRect', 9.29,4.48,3.71,2.32, {fill:F.c1,line:{color:PINK,width:6},rectRadius:.34})
  txt(s, L9, T.o10CS, 1.1,5.64,2.93,.83)
  txt(s, L2, T.a16bC, 1.81,5.27,1.51,.37)
  mark(s, 2.33,4.71,.47,.47)
  txt(s, L9, T.o10CS, 9.68,5.64,2.93,.83)
  txt(s, L12, T.a16bC, 10.3,5.27,1.69,.37)
  txt(s, L9, T.o10CS, 5.39,5.64,2.93,.83)
  txt(s, L5, T.a16bC, 6.06,5.27,1.57,.37)
  shp(s, 'rect', 10.95,5.02,.08,.08, {fill:F.c10})
  shp(s, 'rect', 11.05,4.96,.08,.13, {fill:F.c10})
  shp(s, 'rect', 11.16,4.91,.08,.19, {fill:F.c10})
  shp(s, 'rect', 11.26,4.86,.08,.24, {fill:F.c10})
  art(s, A.a35, 10.92,4.71,.42,.25, F.c10)
  mark(s, 6.56,4.71,.5,.5)
  art(s, A.a96, -.01,0,7.49,1.68, F.c12, {flipH:1})
  art(s, A.a2, 0,0,5.39,1.59, F.c2, {flipH:1})
  art(s, A.a2, 0,0,5.24,1.59, F.c5, {flipH:1})
  art(s, A.a2, 0,-.01,4.73,1.35, F.c8, {flipH:1})
  art(s, A.a34, -.01,0,2.26,.84, F.c2, {rotate:180,flipH:1})
  txt(s, 'Something drew her attention. Her eyes narro wed as if to bri ng a distant object into focus.', T.a28bC, 2,1.97,9.85,1.04)
  txt(s, L11, T.o10CS, 2.7,3.18,7.95,.58)
  shp(s, 'roundRect', 5.9,1.12,1.91,.67, {fill:F.c10,rectRadius:.2})
  logo(s, 6.19,1.26, WHITE)
  art(s, A.a96, 6.06,0,7.32,1.68, F.c12)
  art(s, A.a2, 8.1,0,5.27,1.59, F.c2)
  art(s, A.a2, 8.25,0,5.12,1.59, F.c5)
  art(s, A.a2, 8.75,-.01,4.61,1.35, F.c8)
  art(s, A.a34, 11.17,0,2.21,.84, F.c2, {rotate:180})
  shp(s, 'ellipse', .72,2.31,.28,.28, {fill:F.c3})
  shp(s, 'ellipse', 12.52,3.37,.17,.17, {fill:F.c3})
  shp(s, 'ellipse', 3.44,1.01,.11,.11, {fill:F.c3})
  shp(s, 'ellipse', 9.58,1.23,.28,.28, {fill:F.c3})
}

function slide34(pptx) {
  const s = newSlide(pptx);
  deco26(s);
  shp(s, 'rect', 6.16,2.44,7.13,4.62, {fill:F.c15})
  txt(s, L18, T.a18bC, 10.21,5.79,2.51,.4)
  txt(s, 'User Rodiex Two', T.o10CS, 10.56,6.19,1.88,.33)
  txt(s, L14, T.a18bC, 6.77,5.79,2.51,.4)
  txt(s, 'User Rodiex One', T.o10CS, 7.12,6.19,1.88,.33)
  art(s, A.a27, -.03,.01,13.37,1.3, F.c12)
  art(s, A.a7, 3.69,.01,9.62,1.23, F.c2)
  art(s, A.a7, 3.95,.01,9.36,1.23, F.c5)
  art(s, A.a2, 4.88,0,8.43,1.04, F.c8)
  art(s, A.a5, 9.3,.01,4.03,.65, F.c2, {rotate:180})
  txt(s, L1, T.r11bpN, .83,1.69,1.65,.29)
  txt(s, L25, T.a29b, .83,1.97,5.42,1.08)
  logo(s, .41,.45)
  txt(s, L7, T.o10JS, .85,3.14,4.91,.58)
  txt(s, L8, T.o10, .83,4.12,2.07,.28)
  shp(s, 'rect', .92,4.55,4.31,.08, {fill:F.c10})
  txt(s, L14, T.a12b, .83,3.85,1.95,.3)
  shp(s, 'rect', 4.62,4.55,1.02,.08, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 5.04,4.04,.69,.28)
  txt(s, L8, T.o10, .83,5.08,2.07,.28)
  shp(s, 'rect', .92,5.5,4.43,.08, {fill:F.c10})
  txt(s, L18, T.a12b, .83,4.81,1.95,.3)
  shp(s, 'rect', 4.95,5.5,.69,.08, {fill:F.c14})
  txt(s, '95.2%', T.a10_5bR, 5.04,5,.69,.28)
  bubbles(s, [
    [4.93,6.23,.226], [5.47,1.28,.169], [.29,4.58,.169], [3.14,.74,.095]
  ])
}

function slide35(pptx) {
  const s = newSlide(pptx);
  deco27(s);
  photo(s, -.06,.29,5.27,7.21)
  txt(s, L1, T.r11bpN, 6.54,1.59,1.65,.29)
  txt(s, 'Openend Rodeix With App Mobile', T.a44b, 6.54,1.93,5.19,1.58)
  logo(s, 11.64,.46)
  txt(s, L11, T.o10JS, 6.54,3.63,4.81,.83)
  shp(s, 'roundRect', 6.67,4.85,2.05,.61, {fill:F.c10,rectRadius:.19})
  txt(s, 'DOWNLOAD', T.a16bC, 6.56,4.97,2.32,.37)
  bubbles(s, [
    [5.64,6.11,.452], [5.58,1.16,.25], [12.64,5.06,.181], [10.59,1.61,.181]
  ])
}

function slide36(pptx) {
  const s = newSlide(pptx);
  deco28(s);
  photo(s, 7.17,.66,5.96,5.66)
  txt(s, L1, T.r11bpN, .83,1.95,1.65,.29)
  txt(s, 'Openend Rodeix In Desktop', T.a37b, .83,2.23,5.42,1.35)
  logo(s, .41,6.68)
  txt(s, L8, T.o10, .83,4.12,2.07,.28)
  shp(s, 'rect', .92,4.55,4.31,.08, {fill:F.c10})
  txt(s, L2, T.a12b, .83,3.85,1.95,.3)
  shp(s, 'rect', 4.62,4.55,1.02,.08, {fill:F.c14})
  txt(s, '92.5%', T.a10_5bR, 5.04,4.04,.69,.28)
  txt(s, L8, T.o10, .83,5.08,2.07,.28)
  shp(s, 'rect', .92,5.5,4.43,.08, {fill:F.c10})
  txt(s, L5, T.a12b, .83,4.81,1.95,.3)
  shp(s, 'rect', 4.95,5.5,.69,.08, {fill:F.c14})
  txt(s, '95.2%', T.a10_5bR, 5.04,5,.69,.28)
  bubbles(s, [
    [4.93,6.23,.226], [5.02,1.21,.169], [.29,4.58,.169], [8.75,.66,.095]
  ])
}

function slide37(pptx) {
  const s = newSlide(pptx);
  deco29(s);
  photo(s, .6,1.64,5.88,3.63)
  txt(s, L1, T.r11bpN, 7.46,1.64,1.65,.29)
  txt(s, 'Eassy Access Rodeix With Laptop', T.a36b, 7.46,1.98,5.19,1.31)
  txt(s, L38, T.o10JS, 7.46,3.53,4.8,.83)
  shp(s, 'roundRect', 7.59,4.75,2.05,.61, {fill:F.c10,rectRadius:.19})
  txt(s, 'View Demo', T.a16bC, 7.47,4.87,2.32,.37)
  logo(s, .41,6.68)
  bubbles(s, [
    [7.36,6.45,.22], [6.53,1.64,.25], [12.64,5.06,.181], [11.51,1.67,.181], [3.15,6.17,.387]
  ])
}

function slide38(pptx) {
  const s = newSlide(pptx);
  deco1(s);
  shp(s, 'rect', 0,1.79,13.33,5.68, {fill:F.c15})
  art(s, A.a26, .01,1.64,13.33,5.87, F.c12, {rotate:180,flipH:1})
  art(s, A.a7, 3.72,1.95,9.6,5.56, F.c2, {rotate:180,flipH:1})
  art(s, A.a7, 3.99,1.93,9.33,5.57, F.c5, {rotate:180,flipH:1})
  art(s, A.a2, 4.91,2.82,8.41,4.72, F.c8, {rotate:180,flipH:1})
  art(s, A.a5, 9.32,4.57,4.02,2.94, F.c2, {flipH:1})
  art(s, A.a6, -.04,-.01,13.39,2.31, F.c4, {rotate:180,flipV:1})
  art(s, A.a7, -.02,-.01,9.64,2.19, F.c2, {rotate:180,flipV:1})
  art(s, A.a7, -.02,-.01,9.11,2.11, F.c5, {rotate:180,flipV:1})
  art(s, A.a2, -.02,-.02,7.09,1.97, F.c6, {rotate:180,flipV:1})
  art(s, A.a2, -.05,-.01,4.79,1.33, F.c6, {rotate:180,flipV:1})
  txt(s, 'Get In Touch Rodiex', T.a42b, 3.08,2.28,6.17,.81)
  txt(s, 'Email	:', T.a16b, 3.78,4.13,1.39,.37)
  txt(s, 'Rodiex@mail.com', T.o10_5, 4.99,4.16,2.12,.28)
  txt(s, 'Phone	:', T.a16b, 3.78,4.81,1.62,.37)
  txt(s, '+012 – 3884 -7221', T.o10_5, 4.99,4.83,2.12,.28)
  txt(s, 'Location 	:', T.a16b, 3.78,5.47,1.47,.37)
  txt(s, 'Addresss Line One', T.o10_5, 4.99,5.5,2.12,.28)
  art(s, A.a97, 3.25,4.16,.42,.31, F.c14)
  art(s, A.a98, 3.37,5.43,.19,.28, F.c14)
  art(s, A.a99, 3.26,5.53,.41,.32, F.c14)
  art(s, A.a100, 3.27,4.74,.35,.36, F.c14)
  art(s, A.a101, 3.25,5.04,.4,.16, F.c14)
  shp(s, 'roundRect', 3.1,1.65,1.91,.67, {fill:F.c10,rectRadius:.2})
  logo(s, 3.38,1.8, WHITE)
  txt(s, L38, T.o10JS, 3.14,3.07,5.67,.83)
  txt(s, 'Rodiex12@mail.com', T.o10_5, 4.99,4.42,2.12,.28)
  txt(s, '+321 – 5432 -5323', T.o10_5, 4.99,5.04,2.12,.28)
  txt(s, 'Addresss Line Lorem Ipsum', T.o10_5, 4.99,5.73,2.12,.28)
  bubbles(s, [
    [8.33,1.7,.315], [.45,5.54,.148], [1.76,3.39,.241], [7.14,5.35,.148], [6.59,3.19,.148]
  ])
}

function slide39(pptx) {
  const s = newSlide(pptx);
  deco12(s);
  txt(s, L10, T.r14CN, 3.75,4.24,5.84,.34)
  mark(s, 6.13,1.86,1.07,1.07)
  txt(s, 'Icons Slides', T.r66bCN, 3.46,2.98,6.41,1.21)
}

function slide40(pptx) {
  const s = newSlide(pptx);
  iconWall(s, [
    [.62,.44,.26,.35,'homePlate'], [1.55,.46,.35,.3,'heart'], [2.51,.41,.52,.41,'ellipse'],
    [3.69,.44,.42,.34,'heart'], [4.81,.43,.22,.43,'upArrow'], [5.86,.42,.39,.38,'rect'],
    [6.96,.42,.37,.51,'heart'], [8.06,.44,.46,.46,'ellipse'], [9.16,.46,.47,.46,'downArrow'],
    [10.14,.41,.41,.41,'parallelogram'], [11.22,.39,.42,.44,'star4'], [12.35,.4,.43,.43,'parallelogram'],
    [.63,1.28,.25,.46,'trapezoid'], [1.56,1.29,.33,.42,'cloud'], [2.57,1.32,.41,.38,'donut'],
    [3.65,1.28,.4,.46,'hexagon'], [4.72,1.29,.44,.44,'diamond'], [5.83,1.34,.47,.34,'flowChartDocument'],
    [6.93,1.28,.46,.46,'cloud'], [8.13,1.24,.32,.54,'flowChartMagneticDisk'], [9.15,1.34,.43,.33,'roundRect'],
    [10.12,1.28,.45,.45,'ellipse'], [11.28,1.22,.3,.57,'pentagon'], [12.35,1.35,.39,.42,'hexagon'],
    [.56,2.27,.37,.31,'cloud'], [1.46,2.26,.54,.33,'hexagon'], [2.52,2.22,.46,.46,'ellipse'],
    [3.58,2.2,.49,.49,'ellipse'], [4.75,2.24,.39,.37,'flowChartDocument'], [5.85,2.21,.44,.42,'upArrow'],
    [6.96,2.23,.4,.38,'flowChartDocument'], [8.06,2.22,.47,.4,'roundRect'], [9.25,2.14,.23,.57,'hexagon'],
    [10.14,2.25,.41,.34,'teardrop'], [11.17,2.2,.51,.45,'ellipse'], [12.4,2.2,.32,.45,'trapezoid'],
    [.59,3.1,.33,.52,'roundRect'], [1.44,3.17,.58,.39,'rect'], [2.55,3.17,.44,.38,'cloud'],
    [3.62,3.12,.48,.48,'rightArrow'], [4.74,3.15,.41,.41,'plus'], [5.88,3.18,.37,.37,'flowChartTerminator'],
    [6.93,3.18,.46,.36,'parallelogram'], [8.08,3.17,.42,.38,'flowChartMagneticDisk'],
    [9.16,3.19,.41,.34,'blockArc'], [10.09,3.13,.5,.46,'pie'], [11.19,3.14,.48,.43,'heart'],
    [12.34,3.2,.42,.3,'roundRect'],
    [.51,4.06,.47,.47,'diamond'], [1.57,4.09,.31,.41,'heart'], [2.6,4.03,.35,.53,'plaque'],
    [3.66,4.12,.4,.35,'triangle'], [4.76,4.1,.37,.4,'trapezoid'], [5.86,4.14,.4,.31,'flowChartMagneticDisk'],
    [6.94,4.09,.43,.41,'ellipse'], [8.13,4.04,.33,.52,'hexagon'], [9.14,4.07,.45,.45,'flowChartDocument'],
    [10.09,4.03,.52,.52,'donut'], [11.22,4.08,.42,.43,'roundRect'], [12.3,4.04,.52,.51,'homePlate'],
    [.5,5.06,.51,.36,'trapezoid'], [1.47,5,.5,.48,'pentagon'], [2.55,5.02,.43,.43,'parallelogram'],
    [3.65,4.94,.39,.58,'trapezoid'], [4.8,4.97,.29,.53,'heart'], [5.84,5.04,.45,.39,'flowChartTerminator'],
    [6.94,5.07,.45,.34,'flowChartTerminator'], [8.09,5.04,.41,.39,'ellipse'], [9.15,5,.44,.46,'diamond'],
    [10.15,5,.39,.47,'triangle'], [11.22,5.03,.41,.4,'can'], [12.29,5.04,.54,.39,'cloud'],
    [.51,5.9,.48,.33,'homePlate'], [1.51,5.86,.43,.42,'heart'], [2.56,5.87,.42,.41,'cloud'],
    [3.67,5.83,.36,.48,'flowChartMagneticDisk'], [4.72,5.84,.45,.45,'ellipse'],
    [5.84,5.88,.45,.37,'flowChartMagneticDisk'], [6.95,5.83,.42,.47,'cloud'],
    [8.05,5.89,.48,.36,'parallelogram'], [9.17,5.8,.39,.54,'plaque'], [10.09,5.78,.51,.58,'hexagon'],
    [11.22,5.83,.4,.49,'trapezoid'], [12.29,5.76,.54,.62,'ellipse'],
    [.58,6.64,.34,.39,'triangle'], [1.5,6.66,.44,.36,'heart'], [2.57,6.66,.4,.36,'halfFrame'],
    [3.61,6.62,.48,.44,'homePlate'], [4.73,6.62,.44,.44,'flowChartTerminator'], [5.86,6.6,.42,.49,'trapezoid'],
    [6.82,6.69,.69,.3,'heart'], [8.11,6.66,.37,.37,'flowChartTerminator'], [9.14,6.61,.45,.47,'ellipse'],
    [10.08,6.67,.53,.34,'triangle'], [11.15,6.69,.54,.31,'roundRect'], [12.3,6.58,.53,.53,'triangle']
  ]);
}

function slide41(pptx) {
  const s = newSlide(pptx);
  iconWall(s, [
    [.36,.41,.44,.44,'gear6'], [1.35,.37,.51,.5,'sun'], [2.46,.41,.49,.43,'ellipse'], [3.61,.39,.53,.46,'heart'],
    [4.79,.38,.49,.49,'cloud'], [5.93,.41,.43,.43,'flowChartTerminator'], [6.99,.36,.53,.53,'diamond'],
    [8.17,.43,.53,.39,'heart'], [9.36,.41,.44,.44,'flowChartTerminator'], [10.45,.36,.38,.53,'pentagon'],
    [11.47,.39,.46,.46,'hexagon'], [12.55,.46,.34,.39,'homePlate'],
    [.41,1.32,.33,.45,'donut'], [1.37,1.32,.47,.45,'flowChartDocument'], [2.55,1.32,.32,.46,'heart'],
    [3.56,1.23,.64,.64,'sun'], [4.8,1.32,.47,.45,'trapezoid'], [5.94,1.34,.41,.41,'pentagon'],
    [7.08,1.29,.36,.53,'hexagon'], [8.25,1.34,.36,.41,'star6'], [9.36,1.33,.44,.44,'irregularSeal1'],
    [10.44,1.36,.41,.38,'triangle'], [11.48,1.37,.44,.36,'ellipse'], [12.45,1.43,.55,.24,'plaque'],
    [.3,2.19,.56,.41,'diamond'], [1.38,2.17,.45,.44,'flowChartDocument'], [2.44,2.13,.53,.53,'cloud'],
    [3.63,2.23,.5,.33,'trapezoid'], [4.77,2.23,.53,.33,'hexagon'], [5.93,2.17,.44,.44,'cloud'],
    [7.04,2.15,.44,.47,'blockArc'], [8.16,2.16,.53,.45,'heart'], [9.34,2.16,.48,.47,'heart'],
    [10.39,2.14,.5,.5,'pie'], [11.51,2.21,.38,.35,'ellipse'], [12.43,2.25,.58,.29,'plaque'],
    [.42,3.09,.33,.37,'hexagon'], [1.44,3.11,.33,.33,'trapezoid'], [2.51,3.07,.39,.41,'blockArc'],
    [3.68,3.08,.4,.4,'pentagon'], [4.85,3.09,.37,.37,'flowChartMagneticDisk'],
    [5.94,3.14,.41,.29,'flowChartDocument'], [7,3.1,.52,.36,'octagon'], [8.23,3.05,.4,.45,'roundRect'],
    [9.35,3.05,.45,.46,'flowChartTerminator'], [10.46,3.08,.35,.4,'rect'], [11.44,3.01,.53,.53,'diamond'],
    [12.55,3.08,.35,.4,'trapezoid'],
    [.4,4.03,.35,.29,'ellipse'], [1.41,3.98,.39,.39,'parallelogram'], [2.49,3.95,.44,.44,'irregularSeal1'],
    [3.63,3.96,.51,.42,'rightArrow'], [4.83,3.99,.41,.36,'star6'], [5.93,3.95,.44,.44,'cloud'],
    [7.11,3.93,.29,.49,'diamond'], [8.26,3.96,.34,.41,'trapezoid'], [9.37,3.96,.42,.42,'ellipse'],
    [10.44,3.97,.4,.4,'rightArrow'], [11.53,4.01,.34,.32,'pentagon'], [12.51,4,.42,.35,'parallelogram'],
    [.38,4.81,.4,.39,'parallelogram'], [1.42,4.8,.37,.37,'roundRect'], [2.56,4.86,.3,.44,'hexagon'],
    [3.66,4.82,.44,.48,'flowChartMagneticDisk'], [4.8,4.81,.47,.42,'flowChartDocument'],
    [5.93,4.81,.44,.36,'flowChartTerminator'], [7.08,4.79,.35,.38,'flowChartTerminator'],
    [8.19,4.84,.47,.32,'flowChartDocument'], [9.35,4.85,.46,.4,'flowChartTerminator'],
    [10.41,4.8,.46,.33,'pentagon'], [11.47,4.78,.46,.39,'flowChartMagneticDisk'],
    [12.49,4.84,.47,.36,'flowChartTerminator'],
    [.35,5.72,.46,.4,'ellipse'], [1.38,5.72,.46,.4,'ellipse'], [2.52,5.77,.38,.31,'ellipse'],
    [3.67,5.71,.42,.43,'parallelogram'], [4.84,5.73,.39,.39,'flowChartDocument'], [5.93,5.74,.44,.38,'octagon'],
    [7.04,5.71,.44,.44,'cloud'], [8.23,5.76,.4,.32,'flowChartMagneticDisk'], [9.39,5.68,.37,.49,'can'],
    [10.41,5.7,.45,.45,'hexagon'], [11.48,5.71,.44,.43,'parallelogram'], [12.48,5.7,.48,.46,'flowChartDocument'],
    [.28,6.58,.61,.41,'pentagon'], [1.44,6.56,.33,.44,'ellipse'], [2.5,6.58,.42,.4,'flowChartDocument'],
    [3.65,6.58,.45,.4,'blockArc'], [4.92,6.55,.23,.46,'plus'], [5.94,6.61,.42,.34,'heart'],
    [7.09,6.62,.33,.33,'can'], [8.26,6.61,.34,.35,'roundRect'], [9.39,6.55,.38,.46,'trapezoid'],
    [10.49,6.58,.29,.4,'plaque'], [11.49,6.52,.43,.52,'ellipse'], [12.52,6.58,.41,.41,'ellipse']
  ]);
}

function slide42(pptx) {
  const s = newSlide(pptx);
  iconWall(s, [
    [.81,.59,.41,.42,'downArrow'], [2,.59,.39,.42,'cloud'], [3.1,.59,.43,.42,'homePlate'],
    [4.21,.62,.46,.36,'roundRect'], [5.38,.59,.46,.39,'heart'], [6.55,.56,.43,.41,'star5'],
    [7.56,.56,.42,.41,'upArrow'], [8.63,.6,.36,.39,'trapezoid'], [9.6,.62,.36,.39,'star6'],
    [10.77,.62,.39,.39,'can'], [11.96,.64,.42,.42,'parallelogram'],
    [.79,1.66,.46,.33,'flowChartDocument'], [2,1.63,.39,.4,'donut'], [3.08,1.63,.46,.36,'roundRect'],
    [4.29,1.63,.3,.36,'ellipse'], [5.4,1.63,.44,.4,'heart'], [6.55,1.63,.43,.36,'donut'],
    [7.68,1.67,.19,.31,'rightArrow'], [8.66,1.68,.3,.31,'rightArrow'], [9.66,1.64,.43,.43,'homePlate'],
    [10.78,1.61,.39,.43,'donut'], [11.98,1.66,.46,.39,'roundRect'],
    [.82,2.77,.39,.39,'parallelogram'], [2.07,2.76,.26,.4,'heart'], [3.18,2.75,.26,.38,'ellipse'],
    [4.21,2.76,.46,.36,'flowChartDocument'], [5.4,2.73,.43,.4,'flowChartMagneticDisk'],
    [6.55,2.76,.43,.36,'flowChartMagneticDisk'], [7.54,2.73,.46,.46,'diamond'], [8.61,2.79,.39,.35,'hexagon'],
    [9.68,2.76,.36,.39,'hexagon'], [10.77,2.79,.41,.33,'pentagon'], [12.05,2.74,.39,.46,'can'],
    [.79,3.94,.46,.4,'cloud'], [1.99,3.94,.43,.36,'heart'], [3.1,3.91,.43,.36,'flowChartMagneticDisk'],
    [4.2,3.91,.48,.36,'homePlate'], [5.35,3.91,.53,.39,'rect'], [6.56,3.91,.39,.39,'can'],
    [7.57,3.91,.39,.39,'can'], [8.58,3.91,.46,.39,'roundRect'], [9.69,3.91,.39,.39,'ellipse'],
    [10.81,3.88,.4,.42,'homePlate'], [11.96,3.96,.49,.33,'roundRect'],
    [.87,5.05,.3,.43,'downArrow'], [1.97,5.02,.46,.39,'cube'], [3.11,5.08,.39,.33,'teardrop'],
    [4.23,5.05,.43,.39,'ellipse'], [5.42,5.05,.39,.39,'donut'], [6.55,5.03,.43,.41,'trapezoid'],
    [7.57,5.07,.39,.39,'cloud'], [8.63,5.08,.36,.36,'pie'], [9.61,5.04,.46,.4,'cloud'],
    [10.73,5,.43,.46,'gear6'], [11.98,5.04,.48,1.42,'trapezoid'],
    [.87,6.14,.39,.39,'ellipse'], [2.05,6.14,.42,.42,'parallelogram'], [3.1,6.14,.46,.36,'rect'],
    [4.25,6.17,.36,.36,'downArrow'], [5.34,6.17,.44,.36,'teardrop'], [6.51,6.17,.39,.39,'cloud'],
    [7.55,6.17,.39,.39,'ellipse'], [8.54,6.14,.46,.36,'flowChartTerminator'], [9.63,6.09,.46,.39,'cloud'],
    [10.73,6.06,.49,.42,'roundRect']
  ]);
}

function slide43(pptx) {
  const s = newSlide(pptx);
  iconWall(s, [
    [.91,.75,.44,.42,'ellipse'], [2.19,.78,.46,.36,'roundRect'], [3.41,.83,.4,.31,'heart'],
    [4.49,.84,.3,.31,'flowChartMagneticDisk'], [5.55,.78,.43,.42,'homePlate'], [6.66,.78,.43,.39,'hexagon'],
    [7.76,.79,.39,.38,'homePlate'], [8.87,.79,.39,.39,'donut'], [9.82,.77,.44,.44,'diamond'],
    [10.88,.82,.46,.36,'ellipse'], [11.96,.82,.46,.39,'cloud'],
    [.91,1.89,.47,.33,'trapezoid'], [2.19,1.86,.46,.36,'ellipse'], [3.48,1.83,.23,.42,'upArrow'],
    [4.41,1.83,.46,.39,'trapezoid'], [5.46,1.82,.43,.4,'triangle'], [6.68,1.85,.39,.39,'ellipse'],
    [7.76,1.85,.39,.39,'ellipse'], [8.87,1.89,.39,.33,'can'], [9.88,1.85,.39,.39,'ellipse'],
    [10.91,1.85,.39,.39,'donut'], [11.94,1.82,.49,.43,'flowChartMagneticDisk'],
    [.93,2.99,.43,.36,'ellipse'], [2.25,2.96,1.54,.41,'halfFrame'], [4.44,2.97,.39,.39,'heart'],
    [5.46,2.97,.49,.39,'homePlate'], [6.68,2.97,.43,.4,'parallelogram'], [7.83,2.97,.33,.39,'flowChartDocument'],
    [8.83,2.97,.49,.4,'roundRect'], [9.79,2.93,.46,.39,'flowChartMagneticDisk'], [10.93,2.96,.39,.39,'rect'],
    [11.96,2.93,.49,.46,'hexagon'],
    [.91,4.09,.46,.36,'cloud'], [2.25,4.05,1.55,.46,'donut'], [4.41,4.13,.46,.35,'hexagon'],
    [5.48,4.05,.46,.43,'triangle'], [6.71,4.08,.37,.37,'cloud'], [7.78,4.05,.43,.46,'can'],
    [8.84,4.06,.46,.41,'gear6'], [9.83,4.09,.42,.42,'pie'], [10.88,4.12,.49,.36,'trapezoid'],
    [12.01,4.12,.39,.39,'triangle'],
    [.91,5.14,.43,.41,'pie'], [2.19,5.11,1.62,.45,'donut'], [4.44,5.1,.39,.43,'pentagon'],
    [5.51,5.14,.39,.43,'heart'], [6.66,5.14,.46,.39,'heart'], [7.78,5.17,.4,.33,'homePlate'],
    [8.88,5.14,.39,.39,'can'], [9.85,5.2,.46,.36,'flowChartDocument'], [10.87,5.13,.46,.46,'ellipse'],
    [12.02,5.14,.32,.39,'flowChartDocument'],
    [.97,6.18,.41,.33,'cloud'], [2.31,6.1,1.39,.44,'frame'], [4.4,6.15,.43,.36,'cloud'],
    [5.47,6.15,.49,.39,'roundRect'], [6.71,6.15,.36,.36,'homePlate'], [7.83,6.19,.39,.33,'flowChartTerminator'],
    [8.85,6.13,.48,.48,'rect'], [9.79,6.05,.64,.64,'rect'], [10.79,6.08,.58,.58,'rect'],
    [11.83,6.06,.61,.61,'rect']
  ]);
}

function slide44(pptx) {
  const s = newSlide(pptx);
  deco2(s);
  txt(s, L10, T.r14CN, 3.75,4.24,5.84,.34)
  mark(s, 5.97,1.87,1.07,1.07)
  txt(s, 'Thanks', T.r96bCN, 3.77,2.57,5.79,1.72)
}

function slide45(pptx) {
  const s = newSlide(pptx);
  deco1(s);
  shp(s, 'rect', 0,0,13.33,7.5, {fill:F.c13})
  art(s, A.a6, .01,-.01,13.33,1.78, F.c4)
  art(s, A.a7, 3.73,-.01,9.6,1.69, F.c2)
  art(s, A.a7, 4.25,-.01,9.07,1.63, F.c5)
  art(s, A.a2, 6.27,-.02,7.06,1.52, F.c6)
  art(s, A.a2, 8.58,-.01,4.77,1.03, F.c6)
  art(s, A.a6, -.01,5.63,13.33,1.88, F.c7, {rotate:180})
  art(s, A.a7, .01,5.73,9.6,1.78, F.c2, {rotate:180})
  art(s, A.a7, .01,5.73,9.33,1.79, F.c5, {rotate:180})
  art(s, A.a2, .01,6.01,8.41,1.51, F.c8, {rotate:180})
  art(s, A.a5, -.01,6.57,4.02,.94, F.c2)
  shp(s, 'ellipse', 9.67,2.9,1.29,1.29, {fill:F.c2})
  shp(s, 'ellipse', 11.85,4.41,.25,.25, {fill:F.c29})
  txt(s, L10, T.r14CN, 3.75,4.47,5.84,.34)
  mark(s, 5.97,2.1,1.07,1.07)
  txt(s, 'Thanks You', T.r80bCN, 2.83,3.03,7.66,1.45)
  bubbles(s, [
    [1.09,3.04,.25], [2.41,1.4,.663], [3.47,5.47,.332], [8.87,5.98,.452], [10.7,1.62,.454], [6.8,1.17,.25]
  ])
}


// ---- build ---------------------------------------------------------------
const SLIDES = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40, slide41, slide42, slide43, slide44, slide45];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'RODEIX', width: 13.333, height: 7.5 });
  pptx.layout = 'RODEIX';
  pptx.author = 'RodeiX';
  pptx.title = 'RodeiX - Virtual Reality & Metaverse Presentation Template';
  SLIDES.forEach(function (fn) { fn(pptx); });
  return pptx.writeFile({ fileName: path.join(__dirname, '072d411c-9eeb-45a3-933a-561908425084_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote', f); },
             function (e) { console.error(e); process.exit(1); });

