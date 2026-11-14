// Recreates "Case Study" deck (26 slides, 13.333 x 7.5 in) with pptxgenjs.
// Raster photos in the original are all a flat grey field; here they are drawn as
// grey placeholder shapes that keep the original silhouette, position and size.

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette / fonts
const C = {
  ink:    '262626',  // tx1 lumMod 85%
  body:   '595959',  // tx1 lumMod 65%
  white:  'FFFFFF',
  purple: '6225C6',  // accent1
  pink:   'FF7BB7',  // accent2
  grey:   'D9D9D9',  // bg1 lumMod 85%
  silver: 'BFBFBF',
  wash:   'F2F2F2',  // bg1 lumMod 95%
  lilac:  'CDB6F2',  // accent1 gradient, light bubble
  violet: 'A077E2',  // accent1 gradient, saturated bubble
  pale:   'BFA2EE',
  mid:    '9F74E5',
  deep:   '4A1C94',
  photo:  'CACACA',  // flat colour of every embedded photo
};

const F = { MONT: 'Montserrat', MED: 'Montserrat Medium', SANS: 'Open Sans' };

// the two vertical gradients used by the deck's bubbles (top -> bottom stops)
const G = {
  bubble: [[0, 'F7F3FD'], [0.74, 'B797EC'], [0.83, 'B797EC'], [1, 'CFB9F2']],
  core:   [[0, 'BFA2EE'], [0.74, '6225C6'], [0.83, '9F74E5'], [0.99, 'BFA2EE']],
};

// soft drop shadow used by nearly every raised card in the deck
const SH  = { type: 'outer', color: '000000', opacity: 0.3, blur: 60.6, offset: 24.6, angle: 135 };
const SH2 = { type: 'outer', color: '000000', opacity: 0.3, blur: 24.4, offset: 6.7,  angle: 135 };

// recurring run styles
const S = {
  title  : { fontFace: F.MONT, fontSize: 28, bold: true,  color: C.ink },
  eyebrow: { fontFace: F.MONT, fontSize: 9,  bold: true,  color: C.ink },
  tag    : { fontFace: F.MONT, fontSize: 9,  color: C.grey },
  page   : { fontFace: F.MONT, fontSize: 18, color: C.ink },
  body   : { fontFace: F.SANS, fontSize: 10, color: C.body, lineSpacingMultiple: 1.5 },
  bodyW  : { fontFace: F.SANS, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5, align: 'center' },
  stat   : { fontFace: F.MONT, fontSize: 16, bold: true, color: C.white, align: 'center' },
  statD  : { fontFace: F.MONT, fontSize: 16, bold: true, color: C.ink },
  label  : { fontFace: F.MONT, fontSize: 10, color: C.white, align: 'center' },
  medB   : { fontFace: F.MED,  fontSize: 10, bold: true, color: C.ink },
  med    : { fontFace: F.MED,  fontSize: 10, color: C.ink },
  lead   : { fontFace: F.MONT, fontSize: 12, bold: true, color: C.ink },
  btn    : { fontFace: F.MONT, fontSize: 9,  bold: true, color: C.white, align: 'center' },
};

// ---------------------------------------------------------------- free-form outlines
// Each entry is a normalised (0..1) outline in an SVG-like mini syntax:
//   "M x y"  move   "L x y"  line   "C x1 y1 x2 y2 x y"  cubic   "Z"  close
const P = {
  logo: 'M 0.641 0.7564 C 0.6824 0.7564 0.7207 0.7638 0.7558 0.7785 C 0.7908 0.7933 0.8227 0.8154 0.8514 0.8449 L 0.9254 0.7833 L 0.9146 0.8012 C 0.8248 0.9211 0.6726 1 0.5 1 C 0.4655 1 0.4318 0.9968 0.3992 0.9908 L 0.3759 0.9843 L 0.3929 0.8999 C 0.406 0.8704 0.4246 0.8449 0.4485 0.8233 C 0.4724 0.8017 0.5007 0.7852 0.5334 0.7737 C 0.5661 0.7622 0.6019 0.7564 0.641 0.7564 Z M 0 0 L 1 0 L 1 0.549 C 1 0.6113 0.986 0.6706 0.9607 0.7246 L 0.9592 0.727 L 0.909 0.6835 C 0.8843 0.6669 0.8574 0.6529 0.8281 0.6416 C 0.7695 0.6189 0.7039 0.6076 0.6314 0.6076 C 0.5661 0.6076 0.5059 0.6171 0.4509 0.6362 C 0.3959 0.6552 0.3478 0.6824 0.3068 0.7176 C 0.2657 0.7528 0.2339 0.7942 0.2111 0.8416 L 0.1908 0.901 L 0.1464 0.8679 C 0.056 0.7863 0 0.6736 0 0.549 Z',
  moon: 'M 1 0 L 1 0.2204 L 0.7969 0.2407 C 0.5926 0.2834 0.4492 0.3834 0.4492 0.5 C 0.4492 0.6166 0.5926 0.7166 0.7969 0.7593 L 1 0.7796 L 1 1 L 0.8087 0.9905 C 0.3472 0.9438 0 0.7419 0 0.5 C 0 0.2581 0.3472 0.0562 0.8087 0.0095 Z',
  tick: 'M 0 0.903 L 0 0.903 L 0 0.903 Z M 0.7347 0.0298 C 0.7629 0.0114 0.8017 0 0.8446 0 C 0.9304 0 1 0.0456 1 0.1019 C 1 0.3541 1 0.6063 1 0.8586 C 1 0.8867 0.9826 0.9122 0.9545 0.9306 L 0.9424 0.9339 L 0.9358 0.9458 C 0.9091 0.9651 0.8711 0.9777 0.8283 0.9791 L 0.1626 0.9999 C 0.1198 1.0012 0.0801 0.991 0.0507 0.9735 L 0 0.903 L 0.0402 0.8296 C 0.067 0.8104 0.1049 0.7978 0.1478 0.7964 L 0.6892 0.7795 L 0.6892 0.1019 C 0.6892 0.0737 0.7066 0.0483 0.7347 0.0298 Z',
  petal: 'M 0.5 0.203 C 0.3382 0.203 0.207 0.3341 0.207 0.4959 C 0.207 0.6577 0.3382 0.7889 0.5 0.7889 C 0.6618 0.7889 0.793 0.6577 0.793 0.4959 C 0.793 0.3341 0.6618 0.203 0.5 0.203 Z M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  swoosh: 'M 0.5 0 C 0.7761 0 1 0.4324 1 0.9659 L 0.9982 1 L 0.0018 1 L 0 0.9659 C 0 0.4324 0.2239 0 0.5 0 Z',
  p1: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p2: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0746 0 0.1667 0 Z',
  p3: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p4: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p5: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p6: 'M 0.0319 0 L 1 0 L 1 0.9152 C 1 0.962 0.9857 1 0.9681 1 L 0 1 L 0 0.0848 C 0 0.038 0.0143 0 0.0319 0 Z',
  p7: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p8: 'M 0 0 L 0.0256 0.0107 C 0.0706 0.0306 0.1189 0.0566 0.157 0.0823 C 0.2334 0.1336 0.2863 0.2692 0.3713 0.2831 C 0.4564 0.2969 0.5818 0.1336 0.6674 0.1654 C 0.753 0.1971 0.8294 0.3257 0.885 0.4734 C 0.9302 0.5934 0.9828 0.8308 0.9975 0.9719 L 1 1 L 0.3596 1 L 0.3476 0.98 C 0.2821 0.8685 0.1983 0.7062 0.1358 0.6465 C 0.0941 0.6067 0.0482 0.6174 0.0069 0.6258 L 0 0.6269 Z',
  p9: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0746 0 0.1667 0 Z',
  p10: 'M 0.6198 0.3918 L 0.6198 0.6954 C 0.6198 0.733 0.566 0.7634 0.4995 0.7634 C 0.4331 0.7634 0.3792 0.733 0.3792 0.6954 L 0.3792 0.3918 C 0.3792 0.3543 0.4331 0.3238 0.4995 0.3238 C 0.566 0.3238 0.6198 0.3543 0.6198 0.3918 Z M 0.6198 0.068 L 0.6198 0.218 C 0.6198 0.2556 0.566 0.286 0.4995 0.286 C 0.4331 0.286 0.3792 0.2556 0.3792 0.218 L 0.3792 0.068 C 0.3792 0.0305 0.4331 0 0.4995 0 C 0.566 0 0.6198 0.0305 0.6198 0.068 Z M 1 0.7173 C 1 0.7347 0.9883 0.7522 0.9648 0.7654 L 0.5851 0.9801 C 0.5616 0.9934 0.5308 1 0.5 1 C 0.4692 1 0.4385 0.9934 0.415 0.9801 L 0.4111 0.9768 L 0.0352 0.7643 C -0.0117 0.7377 -0.0117 0.6946 0.0352 0.6681 C 0.0822 0.6415 0.1584 0.6415 0.2054 0.6681 L 0.501 0.8352 L 0.7946 0.6692 C 0.8416 0.6427 0.9178 0.6427 0.9648 0.6692 C 0.9883 0.6825 1 0.6999 1 0.7173 Z',
  p11: 'M 0.4871 0 L 1 0 L 0.9959 0.0041 C 0.9142 0.085 0.7641 0.2335 0.7266 0.355 C 0.6837 0.4939 0.8131 0.6825 0.7688 0.8069 C 0.7466 0.8691 0.7028 0.9313 0.6479 0.9836 L 0.6298 1 L 0 1 L 0.0199 0.9842 C 0.1126 0.9115 0.2574 0.8064 0.3025 0.702 C 0.3626 0.5628 0.2761 0.3288 0.3109 0.2098 C 0.3414 0.1057 0.398 0.0478 0.4764 0.0053 Z',
  p12: 'M 0 0 L 1 0 L 1 1 L 0 1 Z',
  p13: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0739 1 0.1651 L 1 0.8349 C 1 0.9261 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9261 0 0.8349 L 0 0.1651 C 0 0.0739 0.0746 0 0.1667 0 Z',
  p14: 'M 0.1287 0 L 0.8713 0 C 0.9424 0 1 0.025 1 0.0559 L 1 0.9441 C 1 0.975 0.9424 1 0.8713 1 L 0.1287 1 C 0.0576 1 0 0.975 0 0.9441 L 0 0.0559 C 0 0.025 0.0576 0 0.1287 0 Z',
  p15: 'M 0.156 0 L 0.844 0 C 0.9301 0 1 0.0699 1 0.156 L 1 0.844 C 1 0.9301 0.9301 1 0.844 1 L 0.156 1 C 0.0699 1 0 0.9301 0 0.844 L 0 0.156 C 0 0.0699 0.0699 0 0.156 0 Z',
  p16: 'M 0.5 0 L 1 1 L 0 1 Z',
  p17: 'M 0.1907 0.9908 C 0.0325 1.0526 0.0683 0.7891 0.0508 0.6265 C 0.0333 0.4639 -0.0724 0.0769 0.0857 0.0151 C 0.2439 -0.0467 0.9814 0.091 0.9996 0.2557 C 1.0179 0.4205 0.3488 0.929 0.1907 0.9908 Z',
  p18: 'M 0.3806 0 C 0.5708 -0.0007 0.8067 0.0238 0.9808 0.0703 L 1 0.0758 L 1 0.5838 L 0.9664 0.6203 C 0.8518 0.7427 0.7159 0.8696 0.5948 0.9727 L 0.5623 1 L 0.1369 1 L 0.1321 0.9668 C 0.1232 0.899 0.1168 0.8271 0.1051 0.7676 C 0.0675 0.577 -0.1 0.1311 0.0865 0.0391 C 0.139 0.0132 0.2489 0.0005 0.3806 0 Z',
  p19: 'M 0.0951 0.3369 C -0.0021 0.4273 -0.0149 0.4916 0.0128 0.5447 C 0.0405 0.5978 0.2002 0.582 0.2613 0.6555 C 0.3224 0.729 0.2735 0.9502 0.3793 0.9856 C 0.4851 1.021 0.7965 0.9945 0.896 0.8679 C 0.9955 0.7413 1.0265 0.3704 0.9764 0.2261 C 0.9264 0.0818 0.7425 -0.0159 0.5956 0.0022 C 0.4487 0.0202 0.1922 0.2465 0.0951 0.3369 Z',
  p20: 'M 0.6236 0.0001 C 0.765 -0.0035 0.9295 0.0908 0.9764 0.2261 C 1.0265 0.3704 0.9955 0.7413 0.896 0.8679 C 0.7965 0.9945 0.4851 1.021 0.3793 0.9856 C 0.2735 0.9502 0.3224 0.729 0.2613 0.6555 C 0.2002 0.582 0.0405 0.5978 0.0128 0.5447 C -0.0149 0.4916 -0.0021 0.4273 0.0951 0.3369 C 0.1922 0.2465 0.4487 0.0202 0.5956 0.0022 C 0.6048 0.001 0.6141 0.0003 0.6236 0.0001 Z',
  p21: 'M 0.0156 0 L 0.9844 0 C 0.993 0 1 0.0456 1 0.1019 L 1 0.8981 C 1 0.9544 0.993 1 0.9844 1 L 0.0156 1 C 0.007 1 0 0.9544 0 0.8981 L 0 0.1019 C 0 0.0456 0.007 0 0.0156 0 Z',
  p22: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0746 0 0.1667 0 Z',
  p23: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0746 0 0.1667 0 Z',
  p24: 'M 0.0824 0 L 1 0 L 1 1 L 0.0824 1 C 0.0369 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0369 0 0.0824 0 Z',
  p25: 'M 0.1667 0 L 0.5548 0 C 0.8007 0 1 0.1123 1 0.2507 L 1 0.9061 C 1 0.958 0.9254 1 0.8333 1 L 0.4452 1 C 0.1993 1 0 0.8877 0 0.7493 L 0 0.0939 C 0 0.042 0.0746 0 0.1667 0 Z',
  p26: 'M 0.0229 0 L 0.9771 0 C 0.9898 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9898 1 0.9771 1 L 0.0229 1 C 0.0102 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0102 0 0.0229 0 Z',
  p27: 'M 0 0 L 1 0 L 1 1 L 0 1 Z',
  p28: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p29: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p30: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p31: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p32: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p33: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p34: 'M 0.284 0 L 0.9902 0 L 0.9842 0.0193 C 0.9468 0.1171 0.8126 0.2679 0.8156 0.4149 C 0.8187 0.5619 0.9511 0.8174 0.9966 0.9855 L 1 1 L 0.0437 1 L 0.045 0.9977 C 0.0911 0.9216 0.2741 0.8088 0.269 0.6986 C 0.2627 0.5628 -0.0197 0.3524 0.0011 0.2296 C 0.0166 0.1375 0.1377 0.06 0.2662 0.0068 Z',
  p35: 'M 0.284 0 L 0.9902 0 L 0.9842 0.0193 C 0.9468 0.1171 0.8126 0.2679 0.8156 0.4149 C 0.8187 0.5619 0.9511 0.8174 0.9966 0.9855 L 1 1 L 0.0437 1 L 0.045 0.9977 C 0.0911 0.9216 0.2741 0.8088 0.269 0.6986 C 0.2627 0.5628 -0.0197 0.3524 0.0011 0.2296 C 0.0166 0.1375 0.1377 0.06 0.2662 0.0068 Z',
  p36: 'M 0 0 L 1 0 L 1 1 L 0 1 Z',
  p37: 'M 0 0 L 1 0 L 1 1 L 0 1 Z',
  p38: 'M 0.1667 0 L 0.8333 0 C 0.9254 0 1 0.0746 1 0.1667 L 1 0.8333 C 1 0.9254 0.9254 1 0.8333 1 L 0.1667 1 C 0.0746 1 0 0.9254 0 0.8333 L 0 0.1667 C 0 0.0746 0.0746 0 0.1667 0 Z',
  p39: 'M 0.156 0 L 0.844 0 C 0.9301 0 1 0.0338 1 0.0755 L 1 0.9245 C 1 0.9662 0.9301 1 0.844 1 L 0.156 1 C 0.0699 1 0 0.9662 0 0.9245 L 0 0.0755 C 0 0.0338 0.0699 0 0.156 0 Z',
  p40: 'M 0.5 0 C 0.7761 0 1 0.2239 1 0.5 C 1 0.7761 0.7761 1 0.5 1 C 0.2239 1 0 0.7761 0 0.5 C 0 0.2239 0.2239 0 0.5 0 Z',
  p41: 'M 0.1145 0 L 0.8855 0 C 0.9487 0 1 0.0241 1 0.0539 L 1 0.9461 C 1 0.9759 0.9487 1 0.8855 1 L 0.1145 1 C 0.0513 1 0 0.9759 0 0.9461 L 0 0.0539 C 0 0.0241 0.0513 0 0.1145 0 Z',
};

// mini-parser: turn a normalised outline into pptxgenjs custGeom points
function pts(spec, w, h) {
  return spec.split(/(?=[MLCZ])/).map(seg => {
    const k = seg[0];
    const v = seg.slice(1).trim().split(/\s+/).filter(Boolean).map(Number);
    if (k === 'Z') return { close: true };
    if (k === 'M') return { x: v[0] * w, y: v[1] * h, moveTo: true };
    if (k === 'L') return { x: v[0] * w, y: v[1] * h };
    return { x: v[4] * w, y: v[5] * h,
             curve: { type: 'cubic', x1: v[0] * w, y1: v[1] * h, x2: v[2] * w, y2: v[3] * h } };
  });
}

// ---------------------------------------------------------------- drawing helpers
const box = ([x, y, w, h]) => ({ x, y, w, h });

function shape(s, kind, rect, opt = {}, txt) {
  // pptxgenjs rewrites shadow props in place, so hand it a throw-away copy
  const o = { ...box(rect), ...opt, shadow: opt.shadow && { ...opt.shadow } };
  if (txt !== undefined) s.addText(txt, { ...o, shape: kind });
  else s.addShape(kind, o);
}
const oval  = (s, r, o, t) => shape(s, 'ellipse', r, o, t);
const rect  = (s, r, o, t) => shape(s, 'rect', r, o, t);
const rrect = (s, r, o, t) => shape(s, 'roundRect', r, o, t);
const ln    = (s, r, o)    => s.addShape('line', { ...box(r), line: o, flipH: o.flipH, flipV: o.flipV });

// free-form shape scaled into the given rectangle
function blob(s, spec, [x, y, w, h], opt = {}) {
  shape(s, 'custGeom', [x, y, w, h], { points: pts(spec, w, h), ...opt });
}

// pptxgenjs has no gradient fill, so vertical gradients are painted as a stack of
// horizontal slices, each clipped to the host silhouette and filled with the
// interpolated colour at that height.
function gradStops(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0]) {
      const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
      const k = p1 === p0 ? 0 : (t - p0) / (p1 - p0);
      return [0, 2, 4].map(j => Math.round(parseInt(c0.substr(j, 2), 16) * (1 - k) +
                                           parseInt(c1.substr(j, 2), 16) * k)
                                    .toString(16).padStart(2, '0')).join('').toUpperCase();
    }
  }
  return stops[stops.length - 1][1];
}

// `span(t)` returns the [left, right] edges of the silhouette at relative height t
function gradSlices(s, [x, y, w, h], stops, span, opt = {}) {
  const n = Math.max(6, Math.min(20, Math.round(h * 12)));   // ~1 band per 0.08"
  for (let i = 0; i < n; i++) {
    // overlap each band a little into the next so no hairline seam shows through
    const t0 = i / n, t1 = Math.min(1, (i + 1.35) / n), K = 5, right = [], left = [];
    for (let k = 0; k <= K; k++) {
      const t = t0 + (t1 - t0) * k / K, [l, r] = span(t);
      right.push([r, t]); left.unshift([l, t]);
    }
    const pt = right.concat(left);
    s.addShape('custGeom', { x, y, w, h, ...opt, fill: { color: gradStops(stops, (t0 + t1) / 2) },
      points: pt.map(([px, py], j) => ({ x: px * w, y: py * h, moveTo: j === 0 })).concat([{ close: true }]) });
  }
}

const circleSpan = t => { const r = Math.sqrt(Math.max(0, 0.25 - (t - 0.5) ** 2)); return [0.5 - r, 0.5 + r]; };
const gradOval = (s, r, o) => gradSlices(s, r, o.fill, circleSpan, { shadow: o.shadow });
const gradBlob = (s, spec, r, { fill, ...o } = {}) => gradSlices(s, r, fill, spanner(spec), o);

// flatten a normalised outline to a polyline, then report its half-width at height t
function spanner(spec) {
  const poly = [];
  let cur = null;
  for (const p of pts(spec, 1, 1)) {
    if (p.close) { cur = null; continue; }
    if (p.curve && cur) {
      const { x1, y1, x2, y2 } = p.curve;
      for (let k = 1; k <= 12; k++) {
        const u = k / 12, v = 1 - u;
        poly.push([cur[0], cur[1], v ** 3 * cur[0] + 3 * v * v * u * x1 + 3 * v * u * u * x2 + u ** 3 * p.x,
                                   v ** 3 * cur[1] + 3 * v * v * u * y1 + 3 * v * u * u * y2 + u ** 3 * p.y]);
        cur = [poly[poly.length - 1][2], poly[poly.length - 1][3]];
      }
    } else if (cur && !p.moveTo) {
      poly.push([cur[0], cur[1], p.x, p.y]); cur = [p.x, p.y];
    } else cur = [p.x, p.y];
  }
  return t => {
    let lo = Infinity, hi = -Infinity;
    for (const [ax, ay, bx, by] of poly) {
      if ((ay - t) * (by - t) > 0) continue;
      const x = ay === by ? ax : ax + (bx - ax) * (t - ay) / (by - ay);
      lo = Math.min(lo, x); hi = Math.max(hi, x);
      if (ay === by) { lo = Math.min(lo, bx); hi = Math.max(hi, bx); }
    }
    return hi > lo ? [lo, hi] : [0.5, 0.5];
  };
}

// photo placeholders: flat grey, either a plain rectangle or the original silhouette
const photo     = (s, r, o = {}) => rect(s, r, { fill: C.photo, ...o });
const photoBlob = (s, spec, r, o = {}) => blob(s, spec, r, { fill: C.photo, ...o });

// plain text box: no fill, no line, top-anchored like every text box in the deck
function tx(s, [x, y, w, h], body, opt = {}) {
  s.addText(body, { x, y, w, h, valign: 'top', ...opt });
}

// the little "+" glyph: one round-capped bar drawn twice, crossed at 90 degrees
function plus(s, x, y, w, h, color) {
  rrect(s, [x, y, w, h], { rectRadius: h / 2, fill: color });
  rrect(s, [x, y, w, h], { rectRadius: h / 2, fill: color, rotate: 270 });
}

// page furniture repeated on every slide: logo mark, "CASE STUDY", page number,
// the two gradient bubbles and pink crescent top-right, and the @companyname tag
function chrome(s, page) {
  blob(s, P.logo, [0.383, 0, 0.767, 0.85], { fill: C.purple });
  oval(s, [1.04, 0.204, 0.221, 0.221], { fill: C.pink });
  tx(s, [1.352, 0.475, 1.167, 0.252], 'CASE STUDY', S.eyebrow);
  gradOval(s, [12.498, 0.375, 0.452, 0.452], { fill: G.bubble });
  gradOval(s, [12.598, 0.475, 0.252, 0.252], { fill: G.core });
  blob(s, P.moon, [12.305, 0.475, 0.118, 0.252], { fill: C.pink });
  tx(s, [0.798, 6.416, 0.735, 0.404], page, S.page);
  tx(s, [11.334, 6.567, 1.39, 0.252], '@companyname', { ...S.tag, align: 'right' });
}

// purple stat card: big number, caption, sub-caption, optional pink "+" badge
function statCard(s, x, y, big, mid, sub, opt = {}) {
  const w = 2.125, h = 1.173;
  rrect(s, [x, y, w, h], { rectRadius: 0.224, fill: opt.fill || C.purple, shadow: opt.shadow });
  tx(s, [x + 0.4, y + 0.124, 1.327, 0.37], big, S.stat);
  tx(s, [x + 0.4, y + 0.435, 1.327, 0.269], mid, S.label);
  tx(s, [x + 0.111, y + 0.642, 1.906, 0.326], sub, S.bodyW);
  if (opt.plus) plus(s, x + 1.829, y + 0.176, 0.189, 0.055, C.pink);
}

// ---------------------------------------------------------------- slides
// ------------------------------------------------------------------
// 01. STUDY
function slide01(s) {
  photoBlob(s,P.p6,[1.773,1.552,9.787,3.683]);
  chrome(s,'01');
  tx(s,[5.322,2.242,3.438,1.111],'STUDY',{...S.title,fontSize:60,align:'center'});
  tx(s,[4.509,3.267,4.315,0.252],'CASE STUDY PRESENTATION',{fontFace:F.SANS,fontSize:9,color:C.ink,charSpacing:6,align:'center'});
  blob(s,P.petal,[4.877,3.799,0.507,0.507],{fill:C.pink});
  rrect(s,[3.019,3.947,2.198,2.654],{rectRadius:0.279,fill:C.purple,shadow:SH});
  tx(s,[3.733,4.15,0.769,0.37],'80%',{...S.statD,color:C.white});
  gradOval(s,[3.242,4.138,0.379,0.379],{fill:G.bubble});
  blob(s,P.tick,[3.38,4.227,0.103,0.158],{fill:C.white,rotate:40.3});
  tx(s,[3.19,4.59,1.512,0.303],'Analysis 1',{...S.lead,color:C.white});
  tx(s,[3.19,5.389,1.839,0.973],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',{...S.body,color:C.white});
  ln(s,[3.29,5.349,0.28,0],{color:C.white,width:0.75});
  rrect(s,[5.727,4.147,1.878,2.269],{rectRadius:0.239,fill:C.wash});
  tx(s,[6.338,4.321,0.657,0.337],'90%',{...S.title,fontSize:14});
  gradOval(s,[5.918,4.31,0.324,0.324],{fill:G.bubble});
  blob(s,P.tick,[6.036,4.386,0.088,0.135],{fill:C.white,rotate:40.3});
  tx(s,[5.874,4.697,1.292,0.286],'Analysis 2',{...S.title,fontSize:11});
  tx(s,[5.874,5.38,1.571,0.831],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',S.body);
  ln(s,[5.96,5.346,0.239,0],{color:C.purple,width:0.75});
  rrect(s,[8.149,4.147,1.878,2.269],{rectRadius:0.239,fill:C.wash});
  tx(s,[8.76,4.321,0.657,0.337],'70%',{...S.title,fontSize:14});
  gradOval(s,[8.34,4.31,0.324,0.324],{fill:G.bubble});
  blob(s,P.tick,[8.458,4.386,0.088,0.135],{fill:C.white,rotate:40.3});
  tx(s,[8.295,4.697,1.292,0.286],'Analysis 3',{...S.title,fontSize:11});
  tx(s,[8.295,5.38,1.571,0.831],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',S.body);
  ln(s,[8.381,5.346,0.239,0],{color:C.purple,width:0.75});
  plus(s,11.381,1.5,0.357,0.103,C.purple);
  oval(s,[4.833,2.073,0.767,0.767],{fill:C.white});
  tx(s,[4.833,2.153,0.767,0.572],'ce',{fontFace:F.MED,fontSize:28,color:C.purple,align:'center'});
}

// ------------------------------------------------------------------
// 02. Introduction Case Study
function slide02(s) {
  shape(s,'triangle',[1.702,1.357,4.965,4.28],{line:{color:C.pink,width:6},rotate:180});
  chrome(s,'02');
  gradOval(s,[2.111,2.218,0.379,0.379],{fill:G.bubble});
  plus(s,1.533,6.049,0.357,0.103,C.purple);
  tx(s,[7.762,1.484,2.939,1.043],'Introduction Case Study',S.title);
  tx(s,[7.762,1.281,1.39,0.252],'@companyname',S.tag);
  tx(s,[7.762,2.83,4.602,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accum san libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget que sagittis molestie ut accumsan ipsum. ',S.body);
  rrect(s,[7.914,5.071,1.162,1.173],{rectRadius:0.222,fill:C.purple,shadow:SH});
  rrect(s,[9.314,4.457,0.841,0.85],{rectRadius:0.161,fill:C.wash});
  tx(s,[9.314,4.76,0.841,0.471],'Text Here',{...S.title,fontSize:10.5,align:'center'});
  tx(s,[9.314,4.539,0.841,0.286],'50%',{...S.title,fontSize:11,align:'center'});
  rrect(s,[10.324,5.429,0.841,0.85],{rectRadius:0.161,fill:C.wash});
  tx(s,[10.324,5.731,0.841,0.471],'Text Here',{...S.title,fontSize:10.5,align:'center'});
  tx(s,[10.324,5.51,0.841,0.286],'45%',{...S.title,fontSize:11,align:'center'});
  rrect(s,[11.303,4.457,0.841,0.85],{rectRadius:0.161,fill:C.wash});
  tx(s,[11.303,4.76,0.841,0.471],'Text Here',{...S.title,fontSize:10.5,align:'center'});
  tx(s,[11.303,4.539,0.841,0.286],'60%',{...S.title,fontSize:11,align:'center'});
  tx(s,[8.12,5.343,0.769,0.37],'80%',S.stat);
  tx(s,[7.762,5.652,1.512,0.303],'Analysis 1',{...S.stat,fontSize:12});
  photoBlob(s,P.p16,[1.702,1.837,4.965,4.28]);
  blob(s,P.petal,[5.907,5.036,0.507,0.507],{fill:C.pink});
}

// ------------------------------------------------------------------
// 03. Our Case Study Company
function slide03(s) {
  blob(s,P.p17,[6.908,1.135,8.166,7.377],{line:{color:C.pink,width:6},rotate:4.14});
  photoBlob(s,P.p18,[6.66,1.298,6.674,6.202]);
  chrome(s,'03');
  tx(s,[0.798,1.533,3.773,1.043],'Our Case Study Company',S.title);
  tx(s,[0.798,1.33,1.39,0.252],'@companyname',S.tag);
  gradOval(s,[6.453,1.693,0.658,0.658],{fill:G.bubble});
  plus(s,12.593,6.89,0.357,0.103,C.purple);
  blob(s,P.petal,[-0.253,4.296,0.507,0.507],{fill:C.pink});
  tx(s,[0.798,3.082,5.185,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. Fusce bibendum auctor urna a dictum. Morbi varius varius velit.',S.body);
  rrect(s,[2.594,4.695,4.815,0.844],{rectRadius:0.161,fill:C.white,shadow:SH});
  statCard(s,0.918,4.914,'+133.11','Your Text Here','Sed lacinia ipsum quis',{shadow:SH});
  statCard(s,6.048,5.296,'+123.11','Your Text Here','Sed lacinia ipsum quis',{shadow:SH});
  tx(s,[4.158,4.882,2.882,0.252],'CASE STUDY PRESENTATION',{fontFace:F.SANS,fontSize:9,color:C.ink,charSpacing:3,align:'center'});
  ln(s,[3.555,5.014,0.396,0],{color:C.purple,width:1,endArrowType:'triangle'});
  plus(s,2.748,5.09,0.189,0.055,C.pink);
}

// ------------------------------------------------------------------
// 04. Our Impressive Case Study Company
function slide04(s) {
  blob(s,P.p19,[0.618,1.088,6.277,4.861],{line:{color:C.pink,width:6},rotate:14.06});
  photoBlob(s,P.p20,[0.845,1.081,5.821,4.508]);
  chrome(s,'04');
  tx(s,[7.822,1.533,4.377,1.515],'Our Impressive Case Study Company',S.title);
  tx(s,[7.822,1.33,1.39,0.252],'@companyname',S.tag);
  blob(s,P.moon,[9.854,2.71,0.314,0.674],{fill:C.pink,rotate:210});
  tx(s,[7.822,3.579,4.674,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accum san libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque gittis molestie ut accumsan ipsum. ',S.body);
  gradOval(s,[1.741,3.545,0.658,0.658],{fill:G.bubble});
  plus(s,6.299,1.977,0.486,0.14,C.purple);
  rrect(s,[7.822,5.446,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[7.954,5.612,1.126,0.26],'LEARN MORE',S.btn);
  rrect(s,[9.512,5.446,1.39,0.586],{rectRadius:0.293,line:{color:C.purple,width:0.75}});
  tx(s,[9.593,5.612,1.258,0.252],'EXPLORE NOW',{...S.btn,color:C.purple});
}

// ------------------------------------------------------------------
// 05. Our Enterprise Innovation and Achievement Jo
function slide05(s) {
  photoBlob(s,P.p21,[0.811,4.197,11.787,1.8]);
  chrome(s,'05');
  tx(s,[0.798,1.902,5.485,1.515],[{text:'Our Enterprise Innovation and ',options:S.title},{text:'Achievement Journey',options:{...S.title,color:C.pink}}],{});
  tx(s,[0.798,1.699,1.39,0.252],'@companyname',S.tag);
  rrect(s,[6.705,1.057,2.712,3.901],{rectRadius:0.2,fill:C.white,line:{color:C.pink,width:4.5}});
  gradOval(s,[6.375,0.727,0.658,0.658],{fill:G.bubble});
  statCard(s,5.749,1.915,'2013','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  tx(s,[7.034,3.451,2.191,1.084],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit aet accumsan libero, in suscipit dolor. Curabitur in dolo. ',S.body);
  rrect(s,[9.728,1.057,2.712,3.901],{rectRadius:0.2,fill:C.white,line:{color:C.pink,width:4.5}});
  statCard(s,8.825,2.165,'2025','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  tx(s,[10.057,3.451,2.191,1.084],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit aet accumsan libero, in suscipit dolor. Curabitur in dolo. ',S.body);
  gradOval(s,[10.181,0.727,0.658,0.658],{fill:G.bubble});
  photoBlob(s,P.p22,[8.406,1.282,0.819,0.819]);
  photoBlob(s,P.p23,[11.181,1.282,1.068,1.068]);
  blob(s,P.petal,[0.577,5.728,0.507,0.507],{fill:C.pink});
  plus(s,12.419,5.929,0.357,0.103,C.purple);
}

// ------------------------------------------------------------------
// 06. Our Case Study Agenda
function slide06(s) {
  rrect(s,[3.08,0.71,2.48,6.081],{rectRadius:0.213,fill:C.purple});
  blob(s,P.petal,[5.306,2.932,0.507,0.507],{fill:C.pink});
  chrome(s,'06');
  tx(s,[6.857,0.942,3.455,1.043],'Our Case Study Agenda',S.title);
  tx(s,[6.857,0.739,1.39,0.252],'@companyname',S.tag);
  rrect(s,[0.885,1.582,5.286,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  gradOval(s,[5.818,1.252,0.658,0.658],{fill:G.bubble});
  plus(s,8.556,1.91,0.189,0.055,C.pink);
  tx(s,[2.208,2.287,3.752,0.579],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit aet accumsan libero, in suscipit dolor. ',S.body);
  tx(s,[2.199,1.749,1.327,0.37],'+133.11',S.statD);
  tx(s,[2.199,2.06,1.327,0.269],'Your Text Here',S.med);
  rrect(s,[4.332,3.897,5.286,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  gradOval(s,[9.264,3.568,0.658,0.658],{fill:G.bubble});
  tx(s,[5.654,4.603,3.752,0.579],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit aet accumsan libero, in suscipit dolor. ',S.body);
  tx(s,[5.645,4.065,1.327,0.37],'+15.11',S.statD);
  tx(s,[5.645,4.376,1.327,0.269],'Your Text Here',S.med);
  blob(s,P.tick,[0.918,4.003,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.146,3.892,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  blob(s,P.tick,[0.918,5.031,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.146,4.919,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  tx(s,[6.857,2.488,5.286,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accum san libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum.',S.body);
  ln(s,[3.677,6.042,0.239,0],{color:C.white,width:0.75});
  tx(s,[3.572,6.146,1.746,0.269],'phone number here',{...S.med,color:C.white});
  ln(s,[6.02,6.042,0.239,0],{color:C.purple,width:0.75});
  tx(s,[5.915,6.146,1.746,0.269],'company name',S.med);
  ln(s,[7.978,6.042,0.239,0],{color:C.purple,width:0.75});
  tx(s,[7.873,6.146,1.746,0.269],'your email address',S.med);
  photoBlob(s,P.p7,[1.146,1.812,0.801,0.801]);
  photoBlob(s,P.p7,[4.593,4.128,0.801,0.801]);
  photoBlob(s,P.p24,[10.238,3.888,3.095,1.53]);
}

// ------------------------------------------------------------------
// 07. Our About Case Study Company
function slide07(s) {
  blob(s,P.p8,[-0.62,2.937,10.453,4.748],{line:{color:C.pink,width:6},rotate:3});
  photoBlob(s,P.p8,[0,3.123,9.262,4.377]);
  chrome(s,'07');
  tx(s,[1.352,1.505,3.455,1.043],'Our About Case Study Company',S.title);
  tx(s,[1.352,1.302,1.39,0.252],'@companyname',S.tag);
  statCard(s,1.352,3.074,'2013','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  blob(s,P.petal,[13.08,7.273,0.507,0.507],{fill:C.pink});
  gradOval(s,[5.566,3.509,0.658,0.658],{fill:G.bubble});
  plus(s,9.319,7.43,0.486,0.14,C.purple);
  tx(s,[7.786,1.396,4.812,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. Fusce bibendum auctor urna a dictum. ',S.body);
  tx(s,[7.786,3.013,4.812,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum.',S.body);
  rrect(s,[6.92,4.522,5.678,0.831],{rectRadius:0.415,fill:C.white,shadow:SH});
  rrect(s,[7.091,4.643,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[7.223,4.808,1.126,0.26],'LEARN MORE',S.btn);
  tx(s,[9.45,4.805,2.882,0.252],'CASE STUDY PRESENTATION',{fontFace:F.SANS,fontSize:9,color:C.ink,charSpacing:3,align:'center'});
  ln(s,[8.848,4.938,0.396,0],{color:C.purple,width:1,endArrowType:'triangle'});
}

// ------------------------------------------------------------------
// 08. Our About Case Study Company
function slide08(s) {
  chrome(s,'08');
  tx(s,[4.048,0.906,5.238,1.043],'Our About Case Study Company',{...S.title,align:'center'});
  tx(s,[5.97,0.65,1.39,0.252],'@companyname',{...S.tag,align:'center'});
  blob(s,P.petal,[6.413,-0.258,0.507,0.507],{fill:C.pink});
  tx(s,[0.917,2.546,3.592,3.356],[{text:'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. Fusce bibendum auctor urna a dictum. Morbi varius varius velit, ac imperdiet odio semper in. Nam magna odio, venenatis id posuere commodo, sollicitudin et nulla. ',options:{...S.body,breakLine:true}},{text:'',options:{...S.body,breakLine:true}},{text:'Donec commodo metus sit amet velit tristique fringilla',options:{...S.body,italic:true,color:C.purple}}],{});
  statCard(s,4.967,2.577,'2013','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  tx(s,[4.967,4.026,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[4.967,4.296,2.287,1.589],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. ',S.body);
  plus(s,0.522,2.335,0.517,0.149,C.purple);
  tx(s,[7.763,3.052,5.03,2.851],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. Fusce bibendum auctor urna a dictum. Morbi varius varius velit, ac imperdiet odio semper in. Nam magna odio, venenatis id posuere commodo, sollicitudin et nulla. Donec commodo metus sit amet velit tristique fringilla. Nunc efficitur ipsum eu finibus interdum. Phasellus interdum quis quam eu commodo. Aenean ullamcorper nisl efficitur velit pretium consectetur. Sed eu porta tortor. Suspendisse vel felis enim. Nunc nec lorem nec tellus placerat rhoncus.',S.body);
  blob(s,P.tick,[7.927,2.658,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[8.094,2.546,3.944,0.325],'Sed lacinia ipsum quis enim fermentum viverra. ',{...S.title,fontSize:10,lineSpacingMultiple:1.5});
  gradBlob(s,P.swoosh,[6.993,6.614,1.802,0.933],{fill:G.bubble});
}

// ------------------------------------------------------------------
// 09. Hello, I’m Irene Payne
function slide09(s) {
  rrect(s,[6.736,0.906,3.247,5.764],{rectRadius:0.974,line:{color:C.pink,width:6},rotate:4.87});
  photoBlob(s,P.p25,[6.808,0.997,3.102,5.507]);
  chrome(s,'09');
  tx(s,[1.352,1.505,3.455,1.043],'Hello, I’m Irene Payne',S.title);
  tx(s,[1.352,1.302,1.39,0.252],'@companyname',S.tag);
  tx(s,[1.352,2.549,2.411,0.252],'CEO Of Case Study Company',{...S.tag,color:C.purple});
  gradOval(s,[6.479,0.779,0.658,0.658],{fill:G.bubble});
  blob(s,P.petal,[-0.253,6.416,0.507,0.507],{fill:C.pink});
  plus(s,10.239,6.157,0.517,0.149,C.purple);
  tx(s,[1.352,3.211,4.648,1.589],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accum san libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sgittis molestie ut accumsan ipsum. Fusce bibendum auctorna a dictum. Morbi varius varius velit.',S.body);
  rrect(s,[1.351,5.296,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[1.483,5.461,1.126,0.26],'LEARN MORE',S.btn);
  rrect(s,[3.04,5.296,1.39,0.586],{rectRadius:0.293,line:{color:C.purple,width:0.75}});
  tx(s,[3.122,5.461,1.258,0.252],'EXPLORE NOW',{...S.btn,color:C.purple});
  blob(s,P.tick,[10.736,1.632,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[10.964,1.521,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  blob(s,P.tick,[10.736,2.66,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[10.964,2.549,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  statCard(s,9.209,3.956,'+123.122','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
}

// ------------------------------------------------------------------
// 10. Experience Top-Notch Service at Our Study Ca
function slide10(s) {
  photoBlob(s,P.p26,[1.04,4.448,11.246,1.542]);
  chrome(s,'10');
  tx(s,[5.499,1.03,5.835,1.515],[{text:'Experience Top-Notch Service at ',options:S.title},{text:'Our Study Case ',options:{...S.title,color:C.pink}},{text:'Company',options:S.title}],{});
  tx(s,[5.499,0.827,1.39,0.252],'@companyname',S.tag);
  rrect(s,[2.518,0.953,2.48,3.935],{rectRadius:0.213,fill:C.purple});
  rrect(s,[1.04,1.373,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  plus(s,3.29,4.255,0.333,0.096,C.pink);
  tx(s,[3.588,1.443,1.41,0.326],'20 April 2024',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[3.588,1.758,1.41,0.326],'15 May 2024',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[3.588,2.084,1.41,0.326],'23 June 2024',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[3.541,2.635,1.311,1.336],'Sed lacinia sum quis enim ferme ntum viverra. Ut sit amet accu msan libero.',{...S.bodyW,align:'justify'});
  tx(s,[2.096,1.828,1.005,0.37],'+133.11',S.statD);
  gradOval(s,[2.999,1.509,0.243,0.243],{fill:G.bubble});
  tx(s,[1.271,2.484,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[1.399,3.67,0.239,0],{color:C.purple,width:0.75});
  tx(s,[1.293,3.729,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[1.26,2.699,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  rrect(s,[5.433,3.974,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[6.488,4.429,1.005,0.37],'+141.11',S.statD);
  gradOval(s,[7.392,4.11,0.243,0.243],{fill:G.bubble});
  tx(s,[5.664,5.085,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[5.792,6.271,0.239,0],{color:C.purple,width:0.75});
  tx(s,[5.686,6.33,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[5.653,5.3,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  rrect(s,[8.232,3.974,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[9.288,4.429,1.005,0.37],'+124.11',S.statD);
  gradOval(s,[10.192,4.11,0.243,0.243],{fill:G.bubble});
  tx(s,[8.464,5.085,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[8.591,6.271,0.239,0],{color:C.purple,width:0.75});
  tx(s,[8.486,6.33,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[8.453,5.3,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  tx(s,[5.501,2.749,6.785,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Cur abitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut.',S.body);
  photoBlob(s,P.p3,[1.287,1.646,0.731,0.731]);
  photoBlob(s,P.p3,[5.679,4.247,0.731,0.731]);
  photoBlob(s,P.p3,[8.479,4.247,0.731,0.731]);
}

// ------------------------------------------------------------------
// 11. Our Study Case Company is Committed to Deliv
function slide11(s) {
  photoBlob(s,P.p27,[1.15,3.93,5.78,1.877]);
  chrome(s,'11');
  tx(s,[1.04,1.653,6.561,1.515],[{text:'Our Study Case Company is ',options:S.title},{text:'Committed to ',options:{...S.title,color:C.pink}},{text:'Delivering the Best ',options:S.title},{text:'Service.',options:{...S.title,color:C.purple}}],{});
  tx(s,[1.04,1.449,1.39,0.252],'@companyname',S.tag);
  rrect(s,[7.6,1.469,4.815,4.513],{rectRadius:0.336,fill:C.wash});
  gradOval(s,[0.826,5.478,0.658,0.658],{fill:G.bubble});
  plus(s,6.667,3.855,0.517,0.149,C.purple);
  rrect(s,[7.881,0.943,2.023,2.641],{rectRadius:0.281,fill:C.purple,shadow:SH});
  tx(s,[8.083,1.915,1.327,0.286],[{text:'Service',options:{...S.med,fontSize:11,color:C.white}},{text:' One',options:{...S.medB,fontSize:11,color:C.white}}],{});
  tx(s,[8.072,2.184,1.641,1.084],'Sed lacinia ipsum uisn im ferme ntum iverra. Ut sit amet accum san a ante vestib.',{...S.bodyW,align:'justify'});
  plus(s,9.533,1.171,0.189,0.055,C.pink);
  gradOval(s,[8.141,1.144,0.641,0.641],{fill:G.bubble});
  tx(s,[8.141,1.296,0.641,0.337],'01.',{...S.label,fontFace:F.MED,fontSize:14});
  rrect(s,[7.881,3.776,2.023,2.641],{rectRadius:0.281,fill:C.purple,shadow:SH});
  tx(s,[8.083,4.748,1.327,0.286],[{text:'Service',options:{...S.med,fontSize:11,color:C.white}},{text:' Two',options:{...S.medB,fontSize:11,color:C.white}}],{});
  tx(s,[8.072,5.017,1.641,1.084],'Sed lacinia ipsum uisn im ferme ntum iverra. Ut sit amet accum san a ante vestib.',{...S.bodyW,align:'justify'});
  plus(s,9.533,4.004,0.189,0.055,C.pink);
  gradOval(s,[8.141,3.977,0.641,0.641],{fill:G.bubble});
  tx(s,[8.141,4.129,0.641,0.337],'02.',{...S.label,fontFace:F.MED,fontSize:14});
  rrect(s,[10.139,2.411,2.023,2.641],{rectRadius:0.281,fill:C.purple,shadow:SH});
  tx(s,[10.341,3.383,1.327,0.286],[{text:'Service',options:{...S.med,fontSize:11,color:C.white}},{text:' Three',options:{...S.medB,fontSize:11,color:C.white}}],{});
  tx(s,[10.33,3.652,1.641,1.084],'Sed lacinia ipsum uisn im ferme ntum iverra. Ut sit amet accum san a ante vestib.',{...S.bodyW,align:'justify'});
  plus(s,11.792,2.639,0.189,0.055,C.pink);
  gradOval(s,[10.399,2.612,0.641,0.641],{fill:G.bubble});
  tx(s,[10.399,2.764,0.641,0.337],'03.',{...S.label,fontFace:F.MED,fontSize:14});
}

// ------------------------------------------------------------------
// 12. Some Of The Excellent Services In Our Compan
function slide12(s) {
  blob(s,P.petal,[0.711,1.965,0.507,0.507],{fill:C.pink});
  rrect(s,[0.964,2.218,11.459,3.773],{rectRadius:0.144,fill:C.white,line:{color:C.grey,width:0.75}});
  rrect(s,[5.335,4.749,0.507,2.484],{rectRadius:0.254,fill:C.purple,rotate:90});
  chrome(s,'12');
  tx(s,[3.749,0.727,5.835,1.043],[{text:'Some Of The Excellent Services In ',options:{...S.title,align:'center'}},{text:'Our Company',options:{...S.title,color:C.pink,align:'center'}}],{});
  tx(s,[5.991,0.477,1.39,0.252],'@companyname',{...S.tag,align:'center'});
  blob(s,P.moon,[12.28,2.089,0.244,0.524],{fill:C.pink,rotate:226.25,flipH:true});
  plus(s,3.416,5.943,0.333,0.096,C.purple);
  rrect(s,[1.753,1.205,0.507,2.484],{rectRadius:0.254,fill:C.purple});
  rrect(s,[1.352,1.86,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[2.44,2.266,1.005,0.37],'+112.11',S.statD);
  tx(s,[1.583,2.921,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[1.71,4.107,0.239,0],{color:C.purple,width:0.75});
  tx(s,[1.605,4.166,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[1.572,3.136,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  rrect(s,[4.115,4.067,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[5.204,4.473,1.005,0.37],'+201.11',S.statD);
  tx(s,[4.347,5.128,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[4.474,6.314,0.239,0],{color:C.purple,width:0.75});
  tx(s,[4.368,6.373,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[4.335,5.343,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  rrect(s,[6.624,3.644,0.507,1.753],{rectRadius:0.254,fill:C.purple,rotate:180});
  rrect(s,[6.855,2.099,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[7.943,2.505,1.005,0.37],'+42.11',S.statD);
  tx(s,[7.086,3.16,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[7.213,4.346,0.239,0],{color:C.purple,width:0.75});
  tx(s,[7.108,4.405,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[7.075,3.375,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  gradOval(s,[6.57,3.596,0.232,0.232],{fill:G.bubble});
  rrect(s,[10.789,3.676,0.507,2.484],{rectRadius:0.254,fill:C.purple,rotate:90});
  rrect(s,[9.569,2.994,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[10.657,3.4,1.005,0.37],'+13.11',S.statD);
  tx(s,[9.8,4.055,1.327,0.269],'Your Text Here',S.medB);
  ln(s,[9.927,5.241,0.239,0],{color:C.purple,width:0.75});
  tx(s,[9.822,5.3,1.746,0.269],'company name',{...S.med,color:C.silver});
  tx(s,[9.789,4.27,1.883,0.831],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  photoBlob(s,P.p1,[1.605,2.095,0.734,0.734]);
  photoBlob(s,P.p1,[4.368,4.302,0.734,0.734]);
  photoBlob(s,P.p1,[7.108,2.334,0.734,0.734]);
  photoBlob(s,P.p1,[9.822,3.229,0.734,0.734]);
}

// ------------------------------------------------------------------
// 13. Why Does Our Study Case Company Stands Out f
function slide13(s) {
  blob(s,P.petal,[2.437,3.024,0.844,0.844],{fill:C.pink});
  rrect(s,[8.144,1.58,0.645,4.161],{rectRadius:0.323,line:{color:C.grey,width:0.75}});
  rrect(s,[8.629,5.839,0.236,0.417],{rectRadius:0.118,fill:C.purple,rotate:90});
  blob(s,P.petal,[8.13,4.621,0.507,0.507],{fill:C.pink});
  chrome(s,'13');
  tx(s,[1.04,1.472,6.561,1.515],[{text:'Why Does Our Study Case Company ',options:S.title},{text:'Stands Out from the ',options:{...S.title,color:C.pink}},{text:'Rest?',options:S.title}],{});
  tx(s,[1.04,1.268,1.39,0.252],'@companyname',S.tag);
  rrect(s,[3.08,3.206,2.48,3.585],{rectRadius:0.213,fill:C.purple});
  rrect(s,[0.861,3.519,5.286,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  gradOval(s,[5.689,3.637,0.329,0.329],{fill:G.bubble});
  tx(s,[1.892,4.01,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[1.04,4.295,4.826,0.579],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  tx(s,[1.892,3.801,1.69,0.269],'142.299 Followers',{...S.medB,color:C.purple});
  blob(s,P.tick,[8.429,1.871,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[9.053,1.82,1.871,0.269],'Have very good service',S.med);
  blob(s,P.tick,[8.429,2.262,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[9.053,2.211,3.073,0.438],'More than 1000 clients have trusted our company',S.med);
  blob(s,P.tick,[8.429,2.793,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[9.053,2.742,3.073,0.269],'Find multiple case studies the fast way',S.med);
  blob(s,P.tick,[8.429,3.23,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[9.053,3.179,3.073,0.438],'Have a lot of followers on every social media account',S.med);
  plus(s,5.43,6.834,0.517,0.149,C.pink);
  rrect(s,[6.588,4.123,1.808,2.668],{rectRadius:0.219,fill:C.white,shadow:SH});
  tx(s,[6.77,4.771,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[6.759,4.986,1.492,1.084],'Sed lacinia ipsum quisnim ferme ntum viverra. Ut sit amet accum san.',S.body);
  tx(s,[6.776,4.369,1.005,0.37],'+42.11',{...S.statD,color:C.purple});
  tx(s,[6.77,6.234,1.39,0.252],'@companyname',S.tag);
  tx(s,[3.319,5.238,2.134,0.579],'Sed lacinia ipsum quis enim fermentum.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[9.056,4.043,3.556,1.589],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut',S.body);
  gradBlob(s,P.swoosh,[10.028,6.98,1.078,0.52],{fill:G.bubble});
  tx(s,[3.315,5.83,2.134,0.326],'Sed lacinia ipsum quis.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[3.319,6.141,2.134,0.326],'Sed lacinia ipsum quis.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  photoBlob(s,P.p28,[1.042,3.66,0.619,0.619]);
}

// ------------------------------------------------------------------
// 14. Our Problem Statement To Be Solved
function slide14(s) {
  rrect(s,[1.533,2.451,1.859,1.859],{rectRadius:0.31,fill:C.purple});
  photoBlob(s,P.p9,[1.703,2.621,1.521,1.521]);
  chrome(s,'14');
  tx(s,[4.038,1.053,5.257,1.043],'Our Problem Statement To Be Solved',{...S.title,align:'center'});
  tx(s,[5.972,0.85,1.39,0.252],'@companyname',S.tag);
  gradBlob(s,P.p10,[3.817,2.853,0.287,0.508],{fill:G.bubble,rotate:90,flipV:true});
  tx(s,[4.406,2.917,1.327,0.269],'Problem 1',S.medB);
  tx(s,[4.406,3.13,2.048,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero.',S.body);
  rrect(s,[6.929,2.451,1.859,1.859],{rectRadius:0.31,fill:C.purple});
  gradBlob(s,P.p10,[9.213,2.853,0.287,0.508],{fill:G.bubble,rotate:90,flipV:true});
  tx(s,[9.802,2.917,1.327,0.269],'Problem 2',S.medB);
  tx(s,[9.802,3.13,2.048,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero.',S.body);
  blob(s,P.petal,[11.124,4.476,0.507,0.507],{fill:C.pink});
  rrect(s,[1.533,4.653,4.571,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  tx(s,[2.555,4.896,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[2.555,5.153,3.289,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  rrect(s,[6.929,4.653,4.571,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  tx(s,[7.951,4.896,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[7.951,5.153,3.289,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  gradOval(s,[1.403,5.951,0.329,0.329],{fill:G.bubble});
  gradBlob(s,P.swoosh,[-0.279,6.694,1.078,0.52],{fill:G.bubble,rotate:90});
  plus(s,1.473,2.59,0.517,0.149,C.pink);
  photoBlob(s,P.p9,[7.099,2.621,1.521,1.521]);
}

// ------------------------------------------------------------------
// 15. Our Problem Statement To Be Solved
function slide15(s) {
  blob(s,P.p11,[2.975,-0.486,7.767,8.865],{line:{color:C.pink,width:6},rotate:6.73});
  photoBlob(s,P.p11,[3.289,0,7.181,7.511]);
  chrome(s,'15');
  tx(s,[1.074,1.665,3.697,1.515],[{text:'Our Problem ',options:S.title},{text:'Statement To Be ',options:{...S.title,color:C.pink}},{text:'Solved',options:S.title}],{});
  tx(s,[1.074,1.462,1.39,0.252],'@companyname',S.tag);
  gradBlob(s,P.swoosh,[7.602,6.567,1.932,0.933],{fill:G.bubble});
  rrect(s,[8.344,1.216,2.125,1.173],{rectRadius:0.224,fill:C.purple,shadow:SH});
  tx(s,[8.631,1.34,1.552,0.37],[{text:'99% ',options:S.stat},{text:'Done',options:{...S.stat,fontSize:12}}],{});
  tx(s,[8.455,1.858,1.906,0.326],'Sed lacinia ipsum quis',S.bodyW);
  tx(s,[8.745,1.651,1.327,0.269],'Problem 01',S.label);
  plus(s,10.174,1.392,0.189,0.055,C.pink);
  rrect(s,[7.423,3.164,2.125,1.173],{rectRadius:0.224,fill:C.purple,shadow:SH});
  tx(s,[7.737,3.288,1.497,0.37],[{text:'89% ',options:S.stat},{text:'Done',options:{...S.stat,fontSize:12}}],{});
  tx(s,[7.534,3.805,1.906,0.326],'Sed lacinia ipsum quis',S.bodyW);
  tx(s,[7.823,3.598,1.327,0.269],'Problem 02',S.label);
  plus(s,9.253,3.34,0.189,0.055,C.pink);
  rrect(s,[9.492,4.874,2.125,1.173],{rectRadius:0.224,fill:C.purple,shadow:SH});
  tx(s,[9.808,4.998,1.497,0.37],[{text:'95% ',options:S.stat},{text:'Done',options:{...S.stat,fontSize:12}}],{});
  tx(s,[9.602,5.516,1.906,0.326],'Sed lacinia ipsum quis',S.bodyW);
  tx(s,[9.892,5.309,1.327,0.269],'Probblem 03',S.label);
  plus(s,11.321,5.05,0.189,0.055,C.pink);
  tx(s,[1.086,3.479,3.556,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. ',S.body);
  rrect(s,[1.011,5.389,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[1.142,5.555,1.126,0.26],'LEARN MORE',S.btn);
  rrect(s,[2.7,5.389,1.39,0.586],{rectRadius:0.293,line:{color:C.purple,width:0.75}});
  tx(s,[2.781,5.555,1.258,0.252],'EXPLORE NOW',{...S.btn,color:C.purple});
}

// ------------------------------------------------------------------
// 16. The Best Team in Our Case Study Company
function slide16(s) {
  blob(s,P.petal,[5.228,0.477,0.507,0.507],{fill:C.pink});
  chrome(s,'16');
  tx(s,[6.667,1.433,5.015,1.043],'The Best Team in Our Case Study Company',S.title);
  tx(s,[6.667,1.23,1.39,0.252],'@companyname',S.tag);
  rrect(s,[3.08,0.71,2.48,6.081],{rectRadius:0.213,fill:C.purple});
  rrect(s,[0.885,1.582,3.008,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  tx(s,[1.988,1.884,1.757,0.303],'Deborah Baldwin',S.lead);
  tx(s,[1.988,2.133,1.757,0.252],'CEO/Founder',{...S.tag,color:C.pink});
  rrect(s,[1.04,2.61,2.706,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[1.116,2.635,1.296,0.269],'142K Followers',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[2.45,2.694,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[2.602,2.649,1.04,0.269],'Learn More',{...S.medB,color:C.white});
  plus(s,3.83,1.52,0.307,0.088,C.pink);
  rrect(s,[4.26,3.259,3.008,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  tx(s,[5.363,3.562,1.757,0.303],'Christina Ries',S.lead);
  tx(s,[5.363,3.811,1.757,0.252],'Manager',{...S.tag,color:C.pink});
  rrect(s,[4.414,4.288,2.706,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[4.491,4.313,1.296,0.269],'191K Followers',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[5.825,4.371,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[5.976,4.327,1.04,0.269],'Learn More',{...S.medB,color:C.white});
  plus(s,7.204,3.197,0.307,0.088,C.pink);
  rrect(s,[7.773,5.046,3.008,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  tx(s,[8.877,5.348,1.757,0.303],'Nelson Baxley',S.lead);
  tx(s,[8.877,5.598,1.757,0.252],'Founder',{...S.tag,color:C.pink});
  rrect(s,[7.928,6.075,2.706,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[8.005,6.1,1.296,0.269],'122K Followers',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[9.338,6.158,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[9.49,6.114,1.04,0.269],'Learn More',{...S.medB,color:C.white});
  plus(s,10.718,4.984,0.307,0.088,C.pink);
  blob(s,P.tick,[0.918,4.003,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.146,3.892,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  blob(s,P.tick,[0.918,5.031,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.146,4.919,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  gradBlob(s,P.swoosh,[3.354,5.858,1.932,0.933],{fill:G.bubble});
  tx(s,[7.928,3.148,4.57,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellent esque id magna eget neque sagittis molestie ut accumsan ipsum.',S.body);
  photoBlob(s,P.p4,[1.046,1.773,0.734,0.734]);
  photoBlob(s,P.p4,[4.421,3.45,0.734,0.734]);
  photoBlob(s,P.p4,[7.934,5.237,0.734,0.734]);
}

// ------------------------------------------------------------------
// 17. The Team That Drives the Formation of this C
function slide17(s) {
  rrect(s,[6.17,1.863,6.135,3.773],{rectRadius:0.215,fill:C.white,line:{color:C.grey,width:0.75}});
  blob(s,P.petal,[10.074,3.653,0.507,0.507],{fill:C.pink});
  gradBlob(s,P.swoosh,[8.219,5.137,1.019,0.492],{fill:G.bubble});
  chrome(s,'17');
  tx(s,[0.767,2.012,4.842,1.515],[{text:'The Team That Drives ',options:S.title},{text:'the Formation of this ',options:{...S.title,color:C.pink}},{text:'Company',options:S.title}],{});
  tx(s,[0.767,1.809,1.39,0.252],'@companyname',S.tag);
  tx(s,[0.767,3.766,5.162,0.831],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus.',S.body);
  rrect(s,[0.767,4.993,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[0.899,5.158,1.126,0.26],'LEARN MORE',S.btn);
  rrect(s,[10.904,5.113,0.307,1.088],{rectRadius:0.153,fill:C.purple});
  rrect(s,[6.667,1.194,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[6.894,1.779,1.757,0.303],'Kevin Greer',S.lead);
  tx(s,[6.894,2.028,1.757,0.252],'CEO/Founder',{...S.tag,color:C.pink});
  tx(s,[6.894,2.301,1.937,1.336],'Sed lacinia ipsum quis nim ferme ntum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  rrect(s,[8.731,4.256,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[8.958,4.841,1.946,0.303],'Lorraine Wormley',S.lead);
  tx(s,[8.958,5.09,1.757,0.252],'Manager',{...S.tag,color:C.pink});
  tx(s,[8.958,5.363,1.937,1.336],'Sed lacinia ipsum quis nim ferme ntum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  rrect(s,[10.253,1.189,2.333,2.76],{rectRadius:0.206,fill:C.white,shadow:SH});
  tx(s,[10.48,1.774,1.757,0.303],'John Turpin',S.lead);
  tx(s,[10.48,2.023,1.757,0.252],'Marketing',{...S.tag,color:C.pink});
  tx(s,[10.48,2.296,1.937,1.336],'Sed lacinia ipsum quis nim ferme ntum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. ',S.body);
  plus(s,5.868,5.668,0.307,0.088,C.purple);
  rrect(s,[6.08,2.35,0.179,0.439],{rectRadius:0.089,fill:C.purple});
  photoBlob(s,P.p5,[6.936,0.777,0.834,0.834]);
  photoBlob(s,P.p5,[10.522,0.772,0.834,0.834]);
  photoBlob(s,P.p5,[9,3.839,0.834,0.834]);
}

// ------------------------------------------------------------------
// 18. The Driving Force Behind Our Achievements
function slide18(s) {
  blob(s,P.petal,[8.364,4.703,0.507,0.507],{fill:C.pink});
  gradOval(s,[2.738,2.609,0.831,0.831],{fill:G.bubble});
  oval(s,[9.107,2.082,2.345,2.345],{fill:C.purple});
  chrome(s,'18');
  tx(s,[3.778,1.053,5.899,1.043],[{text:'The Driving Force Behind ',options:{...S.title,align:'center'}},{text:'Our Achievements',options:{...S.title,color:C.pink,align:'center'}}],{});
  tx(s,[6.033,0.85,1.39,0.252],'@companyname',{...S.tag,align:'center'});
  rrect(s,[4.89,2.72,1.043,3.268],{rectRadius:0.521,line:{color:C.grey,width:0.75},rotate:270});
  tx(s,[4.909,4.087,1.946,0.303],'Lorraine Wormley',S.lead);
  tx(s,[4.909,4.336,1.757,0.252],'CEO/Founder',{...S.tag,color:C.pink});
  rrect(s,[4.909,4.697,1.485,0.358],{rectRadius:0.179,fill:C.purple,shadow:SH2});
  tx(s,[4.957,4.732,1.39,0.252],'@companyname',{...S.label,fontSize:9});
  tx(s,[0.786,4.367,2.09,1.589],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan.',S.body);
  blob(s,P.tick,[4.509,2.645,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[4.737,2.533,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  blob(s,P.tick,[6.925,2.645,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[7.152,2.533,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  plus(s,10.786,6.65,0.333,0.096,C.purple);
  tx(s,[1.831,6.397,2.391,0.325],'write your phone number here',{...S.body,fontFace:F.MED,color:C.silver});
  photoBlob(s,P.p29,[1.793,2.299,1.451,1.451]);
  photoBlob(s,P.p30,[3.881,3.918,0.884,0.884]);
  photoBlob(s,P.p31,[5.07,5.544,1.324,1.321]);
  photoBlob(s,P.p32,[7.648,4.714,1.073,1.073]);
  photoBlob(s,P.p33,[9.226,2.2,2.109,2.109]);
}

// ------------------------------------------------------------------
// 19. Deborah Baldwin
function slide19(s) {
  blob(s,P.p34,[2.042,-0.259,6.371,8.315],{line:{color:C.pink,width:6}});
  photoBlob(s,P.p35,[2.425,0,5.987,7.5]);
  chrome(s,'19');
  rrect(s,[1.868,5.46,2.811,0.404],{rectRadius:0.202,fill:C.white,shadow:SH});
  rrect(s,[1.921,5.498,2.706,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[1.998,5.522,1.296,0.269],'142K Followers',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[3.332,5.581,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[3.483,5.536,1.04,0.269],'Learn More',{...S.medB,color:C.white});
  rrect(s,[1.553,4.787,2.185,0.404],{rectRadius:0.202,fill:C.white,shadow:SH});
  rrect(s,[1.606,4.825,2.082,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[1.634,4.853,1.296,0.269],'Team Work',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[2.906,4.908,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[3.05,4.863,0.51,0.269],'88%',{...S.medB,color:C.white});
  rrect(s,[1.868,4.121,2.217,0.404],{rectRadius:0.202,fill:C.white,shadow:SH});
  rrect(s,[1.921,4.159,2.111,0.332],{rectRadius:0.166,fill:C.purple});
  tx(s,[1.998,4.183,1.296,0.269],'Leadership',{...S.stat,fontFace:F.MED,fontSize:10});
  ln(s,[3.244,4.242,0.076,0.153],{color:C.white,width:1,flipV:true});
  tx(s,[3.483,4.197,0.549,0.269],'90%',{...S.medB,color:C.white});
  gradBlob(s,P.swoosh,[1.227,1.267,1.932,0.933],{fill:G.bubble});
  plus(s,1.23,2.596,0.486,0.14,C.pink);
  tx(s,[8.361,1.838,4.363,0.572],'Deborah Baldwin',S.title);
  tx(s,[8.361,1.635,1.39,0.252],'@companyname',S.tag);
  tx(s,[8.361,2.411,1.757,0.303],'CEO/Founder',{...S.tag,fontSize:12,color:C.pink});
  tx(s,[8.356,3.094,4.495,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. ',S.body);
  rrect(s,[8.356,4.994,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[8.488,5.16,1.126,0.26],'LEARN MORE',S.btn);
}

// ------------------------------------------------------------------
// 20. BREAK SLIDES
function slide20(s) {
  photoBlob(s,P.p36,[3.163,1.513,7.863,5.987]);
  chrome(s,'20');
  gradBlob(s,P.swoosh,[10.06,2.817,1.932,0.933],{fill:G.bubble});
  plus(s,2.92,1.443,0.486,0.14,C.purple);
  rrect(s,[0.798,2.433,5.016,1.521],{rectRadius:0.76,fill:C.white,shadow:SH});
  tx(s,[1.119,2.836,4.375,0.707],'BREAK SLIDES',{...S.title,fontSize:36,align:'center'});
  rrect(s,[3.406,3.79,1.485,0.358],{rectRadius:0.179,fill:C.purple,shadow:SH2});
  tx(s,[3.454,3.826,1.39,0.252],'@companyname',{...S.label,fontSize:9});
  rrect(s,[1.768,3.786,1.485,0.358],{rectRadius:0.179,fill:C.pink,shadow:SH2});
  tx(s,[1.815,3.821,1.39,0.252],'phone number',{...S.label,fontSize:9});
  gradBlob(s,P.swoosh,[11.469,3.869,0.523,0.252],{fill:G.bubble,rotate:180});
  blob(s,P.petal,[10.773,3.497,0.507,0.507],{fill:C.pink});
  blob(s,P.tick,[0.922,5.102,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.15,4.99,1.629,0.831],'Sed lacinia ipsum quis enim fermentum viverra. ',S.body);
  rrect(s,[10.937,5.768,0.179,0.439],{rectRadius:0.089,fill:C.purple,rotate:90});
}

// ------------------------------------------------------------------
// 21. Our Company's Best Portfolio
function slide21(s) {
  photoBlob(s,P.p37,[1.15,3.485,3.015,2.506]);
  photoBlob(s,P.p12,[8.377,3.043,3.015,2.948]);
  photoBlob(s,P.p12,[8.377,0,3.015,2.948]);
  chrome(s,'21');
  tx(s,[1.04,1.905,4.036,1.043],'Our Company\'s Best Portfolio',S.title);
  tx(s,[1.04,1.701,1.39,0.252],'@companyname',S.tag);
  gradBlob(s,P.swoosh,[7.412,6.567,1.932,0.933],{fill:G.bubble});
  plus(s,3.922,3.415,0.486,0.14,C.purple);
  statCard(s,7.315,0.781,'+123.122','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  statCard(s,10.725,3.965,'+143.244','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
  tx(s,[4.752,3.415,3.228,1.841],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus vitae, consequat ex. Pellentesque id magna eget neque sagittis molestie ut accumsan ipsum. ',S.body);
  rrect(s,[6.359,5.633,1.485,0.358],{rectRadius:0.179,fill:C.purple,shadow:SH2});
  tx(s,[6.407,5.668,1.39,0.252],'@companyname',{...S.label,fontSize:9});
  rrect(s,[4.721,5.628,1.485,0.358],{rectRadius:0.179,fill:C.pink,shadow:SH2});
  tx(s,[4.768,5.663,1.39,0.252],'phone number',{...S.label,fontSize:9});
}

// ------------------------------------------------------------------
// 22. Exploring Our Company's Best Portfolio
function slide22(s) {
  blob(s,P.petal,[4.582,0.475,0.507,0.507],{fill:C.pink});
  chrome(s,'22');
  tx(s,[8.025,1.274,3.9,1.515],[{text:'Exploring Our ',options:S.title},{text:'Company\'s Best ',options:{...S.title,color:C.pink}},{text:'Portfolio',options:S.title}],{});
  tx(s,[8.025,1.07,1.39,0.252],'@companyname',S.tag);
  rrect(s,[4.835,0.71,2.48,6.081],{rectRadius:0.213,fill:C.purple});
  rrect(s,[0.885,1.582,5.286,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  rrect(s,[5.395,2.215,1.485,0.358],{rectRadius:0.179,fill:C.pink,shadow:SH2});
  tx(s,[5.443,2.25,1.39,0.252],'@companyname',{...S.label,fontSize:9});
  blob(s,P.tick,[0.997,3.921,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.386,3.87,1.871,0.269],'Have very good service',S.med);
  blob(s,P.tick,[0.997,4.312,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.386,4.261,2.537,0.438],'More than 1000 clients have trusted our company',S.med);
  blob(s,P.tick,[0.997,4.843,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.386,4.793,2.799,0.438],'Find multiple case studies the fast way',S.med);
  blob(s,P.tick,[0.997,5.371,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.386,5.32,3.073,0.438],'Have a lot of followers on every social media account',S.med);
  tx(s,[4.987,3.815,2.161,0.579],'Sed lacinia ipsum quis enim fermentum.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[4.983,4.541,1.75,0.579],'Sed lacinia ipsum quis.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  tx(s,[4.987,5.267,1.75,0.579],'Sed lacinia ipsum quis.',{...S.body,color:C.white,bullet:{code:'2022',indent:13.5}});
  plus(s,7.932,3.675,0.486,0.14,C.purple);
  photoBlob(s,P.p2,[0.964,1.666,1.352,1.352]);
  photoBlob(s,P.p2,[2.379,1.666,1.352,1.352]);
  photoBlob(s,P.p2,[3.814,1.666,1.352,1.352]);
  photoBlob(s,P.p2,[6.639,4.797,1.352,1.352]);
  photoBlob(s,P.p38,[8.246,3.812,2.337,2.337]);
  statCard(s,9.937,4.772,'+133.11','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
}

// ------------------------------------------------------------------
// 23. Our Company's Best Portfolio
function slide23(s) {
  blob(s,P.petal,[6.962,1.074,0.507,0.507],{fill:C.pink});
  rrect(s,[5.347,3.375,7.251,2.523],{rectRadius:0.217,fill:C.purple});
  chrome(s,'23');
  rrect(s,[1.04,3.9,5.286,1.521],{rectRadius:0.29,fill:C.white,shadow:SH});
  blob(s,P.tick,[1.388,4.864,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.615,4.753,1.36,0.326],'Sed lacinia ipsum',S.body);
  blob(s,P.tick,[3.5,4.864,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[3.728,4.753,1.36,0.326],'Sed lacinia ipsum',S.body);
  rrect(s,[5.363,4.792,1.485,0.358],{rectRadius:0.179,fill:C.pink,shadow:SH2});
  tx(s,[5.411,4.827,1.39,0.252],'phone number',{...S.label,fontSize:9});
  gradBlob(s,P.swoosh,[2.717,6.618,1.932,0.933],{fill:G.bubble});
  tx(s,[1.04,1.754,4.036,1.043],'Our Company\'s Best Portfolio',S.title);
  tx(s,[1.04,1.55,1.39,0.252],'@companyname',S.tag);
  photoBlob(s,P.p13,[1.352,3.214,1.36,1.373]);
  photoBlob(s,P.p13,[3.464,3.214,1.36,1.373]);
  photoBlob(s,P.p14,[7.173,1.301,1.787,4.12]);
  photoBlob(s,P.p14,[9.172,1.301,1.787,4.12]);
  rrect(s,[10.404,1.969,1.81,2.811],{rectRadius:0.227,fill:C.white,shadow:SH});
  plus(s,12.112,1.788,0.486,0.14,C.purple);
  tx(s,[10.586,2.186,1.552,2.346],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus.',S.body);
}

// ------------------------------------------------------------------
// 24. Diving into Our Company's Best Portfolio
function slide24(s) {
  blob(s,P.petal,[5.783,3.167,0.507,0.507],{fill:C.pink});
  rrect(s,[1.15,3.501,6.135,2.487],{rectRadius:0.229,fill:C.white,line:{color:C.grey,width:0.75}});
  rrect(s,[6.088,1.052,6.135,4.064],{rectRadius:0.312,fill:C.white,shadow:SH});
  chrome(s,'24');
  tx(s,[1.04,1.653,3.9,1.515],[{text:'Diving into Our ',options:S.title},{text:'Company\'s',options:{...S.title,color:C.pink}},{text:' Best Portfolio',options:S.title}],{});
  tx(s,[1.04,1.449,1.39,0.252],'@companyname',S.tag);
  tx(s,[3.131,3.881,1.005,0.37],'+201.11',S.statD);
  tx(s,[3.142,4.322,1.327,0.269],'Your Text Here',S.medB);
  tx(s,[3.131,4.537,2.652,0.579],'Sed lacinia ipsum quisnim fermentum viverra. Ut sit amet accum san.',S.body);
  rrect(s,[3.142,5.303,1.485,0.358],{rectRadius:0.179,fill:C.purple,shadow:SH2});
  tx(s,[3.189,5.338,1.39,0.252],'@companyname',{...S.label,fontSize:9});
  plus(s,5.692,0.9,0.486,0.14,C.purple);
  tx(s,[6.405,3.241,1.751,1.336],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. ',S.body);
  photoBlob(s,P.p39,[8.317,1.273,1.698,3.51]);
  photoBlob(s,P.p15,[6.289,1.273,1.698,1.698]);
  photoBlob(s,P.p15,[10.346,1.273,1.698,1.698]);
  photoBlob(s,P.p40,[1.547,4.056,1.377,1.377]);
  statCard(s,9.618,3.414,'+133.11','Your Text Here','Sed lacinia ipsum quis',{shadow:SH,plus:true});
}

// ------------------------------------------------------------------
// 25. Our Industries Analysis
function slide25(s) {
  rrect(s,[5.554,0.71,1.761,6.081],{rectRadius:0.267,fill:C.purple});
  gradBlob(s,P.swoosh,[6.026,6.029,3.15,1.521],{fill:G.bubble});
  chrome(s,'25');
  tx(s,[8.186,1.435,3.149,1.043],'Our Industries Analysis',S.title);
  tx(s,[8.186,1.232,1.39,0.252],'@companyname',S.tag);
  rrect(s,[7.054,3.002,5.155,3.266],{rectRadius:0.261,fill:C.white,shadow:SH});
  shape(s,'rect',[7.576,3.804,0.348,1.901],{fill:C.wash});
  tx(s,[7.487,5.771,1.098,0.269],'Analysis 01',{...S.title,fontSize:10});
  shape(s,'rect',[8.717,4.137,0.348,1.568],{fill:C.wash});
  tx(s,[8.627,5.771,1.098,0.269],'Analysis 01',{...S.title,fontSize:10});
  shape(s,'rect',[9.857,3.891,0.348,1.814],{fill:C.wash});
  tx(s,[9.768,5.771,1.098,0.269],'Analysis 01',{...S.title,fontSize:10});
  shape(s,'rect',[10.997,3.326,0.348,2.379],{fill:C.wash});
  tx(s,[10.908,5.771,1.098,0.269],'Analysis 01',{...S.title,fontSize:10});
  rrect(s,[7.498,3.3,0.888,0.41],{rectRadius:0.205,fill:C.white,shadow:SH});
  tx(s,[7.575,3.366,0.747,0.278],'+133.11',{...S.title,fontSize:10.5,align:'center'});
  rrect(s,[8.634,3.644,0.888,0.41],{rectRadius:0.205,fill:C.white,shadow:SH});
  tx(s,[8.711,3.71,0.747,0.278],'+123.11',{...S.stat,fontSize:10.5,color:C.pink});
  rrect(s,[9.754,3.339,0.888,0.41],{rectRadius:0.205,fill:C.white,shadow:SH});
  tx(s,[9.831,3.405,0.747,0.278],'+127.11',{...S.title,fontSize:10.5,align:'center'});
  rrect(s,[10.943,2.794,0.888,0.41],{rectRadius:0.205,fill:C.white,shadow:SH});
  tx(s,[11.02,2.86,0.747,0.278],'+185.11',{...S.stat,fontSize:10.5,color:C.pink});
  shape(s,'rect',[7.576,4.167,0.348,1.571],{fill:C.deep});
  shape(s,'rect',[9.85,4.054,0.348,1.646],{fill:C.mid});
  tx(s,[0.767,2.794,3.078,1.589],'Sed lacinia ipsum quis enim fermentum viverra. Ut sit amet accumsan libero, in suscipit dolor. Curabitur in dolor nunc. Integer sit amet luctus est, ac accumsan lacus. Aliquam a ante vestibulum, iaculis metus.',S.body);
  rrect(s,[0.767,4.993,1.39,0.586],{rectRadius:0.293,fill:C.purple});
  tx(s,[0.899,5.158,1.126,0.26],'LEARN MORE',S.btn);
  blob(s,P.tick,[0.872,2.37,0.088,0.135],{fill:C.purple,rotate:40.3});
  tx(s,[1.26,2.319,1.871,0.269],'Have very good service',S.med);
  shape(s,'rect',[11.002,3.503,0.348,2.197],{fill:C.pale});
  shape(s,'rect',[8.717,4.312,0.348,1.388],{fill:C.purple});
  photoBlob(s,P.p41,[4.078,1.435,2.273,4.833]);
}

// ------------------------------------------------------------------
// 26. THANKS!
function slide26(s) {
  photoBlob(s,P.p6,[1.773,1.552,9.787,3.683]);
  chrome(s,'26');
  tx(s,[4.526,2.242,4.266,1.111],'THANKS!',{...S.title,fontSize:60,align:'center'});
  tx(s,[4.509,3.267,4.315,0.252],'CASE STUDY PRESENTATION',{fontFace:F.SANS,fontSize:9,color:C.ink,charSpacing:6,align:'center'});
  blob(s,P.petal,[4.877,3.799,0.507,0.507],{fill:C.pink});
  rrect(s,[3.019,3.947,2.198,2.654],{rectRadius:0.279,fill:C.purple,shadow:SH});
  tx(s,[3.733,4.15,0.769,0.37],'80%',{...S.statD,color:C.white});
  gradOval(s,[3.242,4.138,0.379,0.379],{fill:G.bubble});
  blob(s,P.tick,[3.38,4.227,0.103,0.158],{fill:C.white,rotate:40.3});
  tx(s,[3.19,4.59,1.512,0.303],'Analysis 1',{...S.lead,color:C.white});
  tx(s,[3.19,5.389,1.839,0.973],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',{...S.body,color:C.white});
  ln(s,[3.29,5.349,0.28,0],{color:C.white,width:0.75});
  rrect(s,[5.727,4.147,1.878,2.269],{rectRadius:0.239,fill:C.wash});
  tx(s,[6.338,4.321,0.657,0.337],'90%',{...S.title,fontSize:14});
  gradOval(s,[5.918,4.31,0.324,0.324],{fill:G.bubble});
  blob(s,P.tick,[6.036,4.386,0.088,0.135],{fill:C.white,rotate:40.3});
  tx(s,[5.874,4.697,1.292,0.286],'Analysis 2',{...S.title,fontSize:11});
  tx(s,[5.874,5.38,1.571,0.831],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',S.body);
  ln(s,[5.96,5.346,0.239,0],{color:C.purple,width:0.75});
  rrect(s,[8.149,4.147,1.878,2.269],{rectRadius:0.239,fill:C.wash});
  tx(s,[8.76,4.321,0.657,0.337],'70%',{...S.title,fontSize:14});
  gradOval(s,[8.34,4.31,0.324,0.324],{fill:G.bubble});
  blob(s,P.tick,[8.458,4.386,0.088,0.135],{fill:C.white,rotate:40.3});
  tx(s,[8.295,4.697,1.292,0.286],'Analysis 3',{...S.title,fontSize:11});
  tx(s,[8.295,5.38,1.571,0.831],'Sed lacinia ipsum is enim fermentum viv rra. Ut sit amet.',S.body);
  ln(s,[8.381,5.346,0.239,0],{color:C.purple,width:0.75});
  plus(s,11.381,1.5,0.357,0.103,C.purple);
  oval(s,[4.111,1.977,0.767,0.767],{fill:C.white});
  tx(s,[4.111,2.057,0.767,0.572],'ce',{fontFace:F.MED,fontSize:28,color:C.purple,align:'center'});
}
const builders = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'W16', width: 13.333, height: 7.5 });
pptx.layout = 'W16';
pptx.theme = { headFontFace: F.MONT, bodyFontFace: F.SANS };

builders.forEach(fn => fn(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '0433e962-cca1-4f00-84fd-5123abbae46f_grok_final.pptx') })
    .then(f => console.log('wrote', f));
