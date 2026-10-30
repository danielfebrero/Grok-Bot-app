/*
 * Fashion Shop - 51-slide deck rebuilt with pptxgenjs.
 * Run: node <thisfile>.js   ->  writes the .pptx next to this script.
 *
 * Slide canvas: 26.6667 x 15 in (widescreen, oversized template).
 * Palette:  NAVY body copy, RED headings, SKY / ROSE diagonal decoration
 *           on a light BG canvas.
 * Raster artwork in the original deck (photos + the "FASHION" logo bitmap)
 * is redrawn here with native shapes / text - no embedded images.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const NAVY  = '2C4772';   // theme dk1 - body text
const SKY   = '70D6F9';   // theme accent1 - blue diagonals
const ROSE  = 'E6B8C0';   // theme accent2 - pink blocks
const RED   = 'A32735';   // theme accent3 - headings
const BG    = 'E5E6E7';   // theme lt1 - slide background
const WHITE = 'FFFFFF';

const HEAD     = 'Montserrat SemiBold';   // major font
const BODY     = 'Montserrat';            // minor font
const ICONFONT = 'Font Awesome 5 Free Solid';

const W = 26.6667, H = 15;                // slide size (inches)

// ---------------------------------------------------- reusable geometry
// Every polygon is stored as fractions of its own bounding box, so the same
// outline can be stamped at any position / size on the slide.
const POLYGONS = {
  cornerTL:    [[1, 0], [0.158, 0], [0, 0.158], [0, 1]],
  cornerBR:    [[1, 0], [0, 1], [0.842, 1], [1, 0.842]],
  stripe:      [[1, 0], [0.745, 0], [0, 1], [0.255, 1]],
  cornerTLbig: [[1, 0], [0.321, 0], [0, 0.321], [0, 1]],
  cornerBRbig: [[1, 0], [0, 1], [0.68, 1], [1, 0.68]],
  sash:        [[0.943, 0], [1, 0], [1, 0.2815], [0.3224, 1], [0.2539, 1], [0, 1], [0.2539, 0.7306]],
  cutTL10:     [[0.1015, 0], [1, 0], [1, 1], [0, 1], [0, 0.3724]],
  cutTL24:     [[0.242, 0], [1, 0], [1, 1], [0, 1], [0, 0.2172]],
  cutTL20:     [[0.2039, 0], [1, 0], [1, 1], [0, 1], [0, 0.1611]],
  cutTL14:     [[0.1436, 0], [1, 0], [1, 1], [0, 1], [0, 0.3875]],
  cutTL10b:    [[0.1015, 0], [1, 0], [1, 1], [0, 1], [0, 0.5013]],
  cutTL24b:    [[0.242, 0], [1, 0], [1, 1], [0, 1], [0, 0.241]],
  cutBR80:     [[0, 0], [1, 0], [1, 0.7985], [0.7961, 1], [0, 1]],
  cutBR63:     [[0, 0], [1, 0], [1, 0.6276], [0.8985, 1], [0, 1]],
  halfTri:     [[1, 0], [0, 1], [1, 1]],
  cutTLBR:     [[0.2037, 0], [1, 0], [1, 0.7983], [0.796, 1], [0, 1], [0, 0.2016]],
  cutTL20c:    [[0, 0.1611], [0.2039, 0], [1, 0], [1, 1], [0, 1]],
  sashOpen:    [[1, 0], [0.9236, 0], [0, 0.7306], [0, 1], [0.0919, 1], [1, 0.2815]],
};

// Line-art glyphs redrawn from the template's freeform icons.
// Compact path mini-language: M/L = move/line, C = cubic bezier (3 pts),
// Q = quadratic (2 pts), Z = close.  All coordinates are 0..1 of the box.
const GLYPHS = {
  hat:
    'M 0.971 0.636 C 0.925 0.521 0.835 0.393 0.721 0.375 C 0.713 0.166 0.615 0 0.495 0 C 0.371 0 0.269 0.18 0.269 0.4 C 0.269 0.407 0.269 0.411 0.269 0.416 C 0.199 0.431 0.056 0.51 0.003 0.618 C 0 0.625 0 0.625 0 0.625 C 0.001 0.634 0.001 0.634 0.001 0.634 C 0.001 0.638 0.043 0.973 0.222 0.973 C 0.241 0.973 0.26 0.969 0.282 0.962 C 0.306 0.953 0.329 0.944 0.349 0.937 C 0.518 0.874 0.569 0.856 0.727 0.937 C 0.827 0.989 0.892 0.942 0.929 0.894 C 0.971 0.843 0.99 0.771 0.995 0.742 C 1 0.715 0.98 0.658 0.971 0.636 Z M 0.495 0.036 C 0.608 0.036 0.701 0.2 0.701 0.4 C 0.701 0.445 0.697 0.472 0.688 0.481 C 0.676 0.497 0.65 0.483 0.619 0.465 C 0.586 0.447 0.546 0.427 0.501 0.427 C 0.458 0.427 0.419 0.447 0.383 0.465 C 0.348 0.483 0.317 0.499 0.302 0.481 C 0.294 0.47 0.289 0.443 0.289 0.4 C 0.289 0.2 0.382 0.036 0.495 0.036 Z M 0.734 0.903 C 0.569 0.818 0.513 0.84 0.345 0.901 C 0.325 0.91 0.302 0.917 0.278 0.926 C 0.084 0.996 0.03 0.692 0.022 0.636 C 0.072 0.542 0.206 0.467 0.272 0.452 C 0.275 0.479 0.282 0.497 0.291 0.508 C 0.299 0.521 0.31 0.526 0.321 0.526 C 0.34 0.524 0.363 0.512 0.39 0.499 C 0.423 0.481 0.462 0.463 0.501 0.463 C 0.543 0.463 0.58 0.483 0.613 0.501 C 0.65 0.521 0.68 0.537 0.701 0.51 C 0.713 0.494 0.72 0.463 0.721 0.411 C 0.836 0.431 0.911 0.564 0.939 0.622 C 0.966 0.679 0.976 0.721 0.976 0.733 C 0.976 0.735 0.923 1 0.734 0.903 Z',
  sunglasses:
    'M 0.768 0 C 0.237 0 0.237 0 0.237 0 C 0.237 0 0.237 0 0.237 0 C 0.234 0 0.233 0 0.232 0 C 0.12 0 0 0.201 0 0.502 C 0 0.635 0.024 0.762 0.07 0.858 C 0.113 0.95 0.169 1 0.225 1 C 0.234 1 0.245 1 0.254 0.997 C 0.354 0.963 0.442 0.752 0.46 0.505 C 0.464 0.44 0.462 0.375 0.456 0.319 C 0.47 0.3 0.506 0.266 0.544 0.319 C 0.536 0.375 0.535 0.44 0.54 0.505 C 0.558 0.752 0.646 0.963 0.745 0.997 C 0.755 1 0.764 1 0.775 1 C 0.831 1 0.887 0.95 0.93 0.858 C 0.975 0.762 1 0.635 1 0.502 C 1 0.201 0.88 0 0.768 0 Z M 0.439 0.495 C 0.423 0.721 0.343 0.916 0.251 0.947 C 0.192 0.966 0.13 0.92 0.083 0.82 C 0.043 0.734 0.02 0.622 0.02 0.502 C 0.02 0.232 0.13 0.05 0.232 0.05 C 0.276 0.05 0.343 0.059 0.389 0.155 C 0.442 0.263 0.446 0.399 0.439 0.495 Z M 0.45 0.269 C 0.439 0.214 0.423 0.161 0.401 0.118 C 0.387 0.087 0.372 0.065 0.355 0.05 C 0.645 0.05 0.645 0.05 0.645 0.05 C 0.628 0.065 0.613 0.087 0.599 0.118 C 0.577 0.161 0.561 0.214 0.55 0.269 C 0.507 0.217 0.468 0.248 0.45 0.269 Z M 0.917 0.82 C 0.87 0.92 0.808 0.966 0.748 0.947 C 0.657 0.916 0.576 0.721 0.561 0.495 C 0.554 0.399 0.557 0.263 0.611 0.155 C 0.657 0.059 0.722 0.05 0.768 0.05 C 0.87 0.05 0.98 0.232 0.98 0.502 C 0.98 0.622 0.957 0.734 0.917 0.82 Z',
  watch:
    'M 1 0.334 C 0.664 0.219 0.664 0.219 0.664 0.219 C 0.629 0.086 0.568 0 0.5 0 C 0.431 0 0.371 0.086 0.334 0.219 C 0 0.334 0 0.334 0 0.334 C 0 0.666 0 0.666 0 0.666 C 0.334 0.781 0.334 0.781 0.334 0.781 C 0.371 0.914 0.431 1 0.5 1 C 0.568 1 0.629 0.914 0.664 0.781 C 1 0.666 1 0.666 1 0.666 L 1 0.334 Z M 0.018 0.625 C 0.018 0.375 0.018 0.375 0.018 0.375 C 0.323 0.271 0.323 0.271 0.323 0.271 C 0.308 0.34 0.3 0.418 0.3 0.501 C 0.3 0.582 0.308 0.66 0.323 0.729 L 0.018 0.625 Z M 0.5 0.954 C 0.4 0.954 0.318 0.749 0.318 0.501 C 0.318 0.251 0.4 0.046 0.5 0.046 C 0.6 0.046 0.68 0.251 0.68 0.501 C 0.68 0.749 0.6 0.954 0.5 0.954 Z M 0.982 0.625 C 0.677 0.729 0.677 0.729 0.677 0.729 C 0.691 0.663 0.699 0.582 0.699 0.501 C 0.699 0.418 0.691 0.337 0.677 0.271 C 0.982 0.375 0.982 0.375 0.982 0.375 L 0.982 0.625 Z',
  watchHand:
    'M 0.714 0.025 C 0.813 0.425 0.813 0.425 0.813 0.425 C 0.022 0.875 0 0.925 0 0.95 C 0.165 0.983 0.165 0.983 0.165 0.983 C 0.165 0.992 0.165 1 0.165 1 C 0.198 0.95 0.593 0.708 0.956 0.508 C 1 0.483 1 0.483 1 0.483 C 0.89 0 0.89 0 0.89 0 L 0.714 0.025 Z',
  dollar:
    'M 0.983 0.569 C 0.973 0.549 0.96 0.533 0.939 0.519 C 0.923 0.503 0.899 0.491 0.875 0.481 C 0.852 0.471 0.825 0.461 0.798 0.453 C 0.771 0.445 0.744 0.439 0.714 0.433 C 0.694 0.429 0.673 0.425 0.65 0.421 C 0.677 0.269 0.677 0.269 0.677 0.269 C 0.684 0.271 0.687 0.271 0.694 0.273 C 0.737 0.285 0.737 0.285 0.737 0.285 C 0.754 0.289 0.768 0.291 0.778 0.295 C 0.815 0.303 0.845 0.299 0.869 0.291 C 0.882 0.285 0.896 0.277 0.909 0.267 C 0.987 0.198 0.987 0.198 0.987 0.198 C 0.97 0.19 0.97 0.19 0.97 0.19 C 0.939 0.172 0.896 0.154 0.848 0.138 C 0.808 0.124 0.758 0.112 0.704 0.106 C 0.721 0 0.721 0 0.721 0 C 0.593 0 0.593 0 0.593 0 C 0.562 0 0.539 0.006 0.522 0.016 C 0.505 0.026 0.495 0.04 0.492 0.052 C 0.485 0.098 0.485 0.098 0.485 0.098 C 0.424 0.098 0.367 0.106 0.32 0.118 C 0.263 0.132 0.215 0.15 0.178 0.172 C 0.141 0.196 0.111 0.22 0.091 0.248 C 0.071 0.275 0.061 0.305 0.061 0.335 C 0.061 0.373 0.071 0.407 0.094 0.435 C 0.118 0.461 0.145 0.483 0.182 0.501 C 0.219 0.517 0.259 0.531 0.303 0.541 C 0.337 0.549 0.374 0.557 0.407 0.563 C 0.38 0.727 0.38 0.727 0.38 0.727 C 0.367 0.723 0.357 0.721 0.347 0.717 C 0.327 0.711 0.31 0.705 0.293 0.699 C 0.273 0.693 0.256 0.687 0.242 0.683 C 0.205 0.671 0.168 0.671 0.135 0.683 C 0.121 0.687 0.108 0.693 0.098 0.701 C 0 0.784 0 0.784 0 0.784 C 0.017 0.794 0.017 0.794 0.017 0.794 C 0.061 0.818 0.118 0.84 0.178 0.856 C 0.236 0.872 0.293 0.882 0.354 0.89 C 0.337 1 0.337 1 0.337 1 C 0.465 1 0.465 1 0.465 1 C 0.495 1 0.515 0.994 0.535 0.982 C 0.552 0.972 0.562 0.96 0.566 0.946 C 0.572 0.896 0.572 0.896 0.572 0.896 C 0.633 0.892 0.69 0.882 0.741 0.87 C 0.795 0.854 0.842 0.836 0.879 0.812 C 0.919 0.788 0.949 0.76 0.97 0.731 C 0.99 0.699 1 0.667 1 0.631 C 1 0.609 0.993 0.587 0.983 0.569 Z M 0.919 0.719 C 0.902 0.745 0.875 0.768 0.842 0.79 C 0.808 0.81 0.768 0.826 0.717 0.84 C 0.667 0.854 0.609 0.862 0.545 0.864 C 0.522 0.866 0.522 0.866 0.522 0.866 C 0.512 0.944 0.512 0.944 0.512 0.944 C 0.508 0.95 0.505 0.956 0.495 0.96 C 0.488 0.966 0.478 0.968 0.465 0.968 C 0.394 0.968 0.394 0.968 0.394 0.968 C 0.411 0.862 0.411 0.862 0.411 0.862 C 0.387 0.86 0.387 0.86 0.387 0.86 C 0.323 0.854 0.259 0.842 0.202 0.828 C 0.152 0.814 0.108 0.798 0.071 0.78 C 0.141 0.721 0.141 0.721 0.141 0.721 C 0.148 0.717 0.155 0.713 0.162 0.709 C 0.178 0.705 0.199 0.705 0.219 0.711 C 0.232 0.715 0.246 0.721 0.263 0.727 C 0.283 0.733 0.3 0.739 0.32 0.745 C 0.343 0.752 0.37 0.758 0.397 0.762 C 0.428 0.766 0.428 0.766 0.428 0.766 C 0.465 0.539 0.465 0.539 0.465 0.539 C 0.444 0.535 0.444 0.535 0.444 0.535 C 0.404 0.529 0.36 0.521 0.32 0.511 C 0.283 0.503 0.246 0.491 0.215 0.475 C 0.185 0.461 0.162 0.443 0.141 0.419 C 0.125 0.397 0.114 0.369 0.114 0.335 C 0.114 0.309 0.121 0.283 0.141 0.261 C 0.158 0.236 0.182 0.216 0.215 0.196 C 0.249 0.176 0.29 0.16 0.34 0.148 C 0.391 0.136 0.448 0.128 0.508 0.128 C 0.535 0.128 0.535 0.128 0.535 0.128 C 0.545 0.056 0.545 0.056 0.545 0.056 C 0.549 0.05 0.552 0.044 0.562 0.038 C 0.569 0.034 0.579 0.032 0.593 0.032 C 0.66 0.032 0.66 0.032 0.66 0.032 C 0.643 0.132 0.643 0.132 0.643 0.132 C 0.67 0.134 0.67 0.134 0.67 0.134 C 0.727 0.14 0.781 0.15 0.825 0.166 C 0.859 0.178 0.889 0.19 0.916 0.204 C 0.862 0.25 0.862 0.25 0.862 0.25 C 0.855 0.255 0.848 0.261 0.842 0.263 C 0.832 0.267 0.818 0.269 0.798 0.265 C 0.785 0.261 0.774 0.259 0.761 0.255 C 0.714 0.244 0.714 0.244 0.714 0.244 C 0.697 0.24 0.68 0.236 0.66 0.234 C 0.63 0.228 0.63 0.228 0.63 0.228 C 0.593 0.443 0.593 0.443 0.593 0.443 C 0.616 0.447 0.616 0.447 0.616 0.447 C 0.643 0.453 0.67 0.457 0.697 0.463 C 0.724 0.469 0.751 0.475 0.774 0.483 C 0.801 0.489 0.822 0.497 0.845 0.507 C 0.865 0.515 0.882 0.525 0.896 0.537 C 0.912 0.549 0.923 0.563 0.933 0.579 C 0.939 0.593 0.946 0.611 0.946 0.631 C 0.946 0.663 0.936 0.691 0.919 0.719 Z',
  dollarStemTop:
    'M 0.333 0.028 C 0.183 0 0.183 0 0.183 0 C 0 1 0 1 0 1 C 0.183 0.991 0.183 0.991 0.183 0.991 C 0.717 0.934 1 0.755 1 0.462 C 1 0.349 0.933 0.255 0.8 0.189 C 0.683 0.123 0.533 0.075 0.333 0.028 Z M 0.3 0.811 C 0.4 0.217 0.4 0.217 0.4 0.217 C 0.483 0.236 0.567 0.264 0.617 0.302 C 0.7 0.34 0.733 0.396 0.733 0.462 C 0.733 0.642 0.6 0.755 0.3 0.811 Z',
  dollarStemBot:
    'M 0.467 0.071 C 0.35 0.101 0.267 0.141 0.2 0.172 C 0.133 0.222 0.083 0.263 0.05 0.313 C 0.017 0.364 0 0.424 0 0.475 C 0 0.606 0.067 0.707 0.2 0.788 C 0.317 0.859 0.483 0.919 0.667 0.96 C 0.833 1 0.833 1 0.833 1 C 1 0 1 0 1 0 C 0.833 0.02 0.833 0.02 0.833 0.02 C 0.683 0.03 0.567 0.051 0.467 0.071 Z M 0.6 0.768 C 0.517 0.737 0.45 0.707 0.383 0.667 C 0.3 0.616 0.267 0.556 0.267 0.475 C 0.267 0.444 0.267 0.404 0.3 0.374 C 0.317 0.343 0.35 0.313 0.383 0.293 C 0.433 0.263 0.483 0.242 0.567 0.222 C 0.6 0.212 0.65 0.202 0.7 0.192 L 0.6 0.768 Z',
  check:
    'M 0.896 0 L 0.368 0.716 L 0.105 0.361 L 0 0.502 L 0.368 1 L 1 0.141 L 0.896 0 Z M 0.043 0.502 L 0.105 0.418 L 0.368 0.774 L 0.896 0.059 L 0.957 0.141 L 0.368 0.941 L 0.043 0.502 Z',
  star:
    'M 0.71 0.247 L 0.5 0 L 0.29 0.247 L 0 0.382 L 0.16 0.671 L 0.192 1 L 0.5 0.929 L 0.809 1 L 0.841 0.671 L 1 0.382 L 0.71 0.247 Z M 0.807 0.657 L 0.78 0.958 L 0.5 0.894 L 0.221 0.958 L 0.194 0.657 L 0.049 0.398 L 0.309 0.278 L 0.5 0.053 L 0.689 0.278 L 0.952 0.398 L 0.807 0.657 Z',
  starInner:
    'M 0.499 0 L 0.289 0.248 L 0 0.381 L 0.159 0.669 L 0.189 1 L 0.499 0.926 L 0.808 1 L 0.836 0.669 L 1 0.381 L 0.707 0.248 L 0.499 0 Z M 0.781 0.65 L 0.755 0.922 L 0.499 0.863 L 0.243 0.922 L 0.219 0.65 L 0.085 0.413 L 0.326 0.3 L 0.499 0.093 L 0.674 0.3 L 0.915 0.413 L 0.781 0.65 Z',
  arrow:
    'M 0.431 0.076 L 0.606 0.248 L 0 0.855 L 0.145 1 L 0.749 0.395 L 0.924 0.566 L 1 0 L 0.431 0.076 Z M 0.894 0.483 L 0.749 0.338 L 0.145 0.944 L 0.054 0.855 L 0.659 0.248 L 0.515 0.103 L 0.953 0.044 L 0.894 0.483 Z',
  quoteLeft:
    'M 0.631 0.208 C 0.617 0.204 0.603 0.201 0.589 0.199 C 0.632 0.038 0.772 0 0.772 0 C 0.552 0.032 0.32 0.105 0.177 0.257 C 0.008 0.438 0 0.733 0.23 0.88 C 0.271 0.905 0.317 0.924 0.366 0.937 C 0.603 1 0.854 0.887 0.926 0.685 C 1 0.484 0.867 0.27 0.631 0.208 Z',
  quoteRight:
    'M 0.631 0.208 C 0.617 0.204 0.603 0.201 0.588 0.199 C 0.631 0.038 0.77 0 0.77 0 C 0.552 0.032 0.32 0.105 0.177 0.257 C 0.008 0.438 0 0.733 0.23 0.88 C 0.271 0.905 0.317 0.924 0.366 0.937 C 0.603 1 0.852 0.887 0.926 0.685 C 1 0.484 0.867 0.27 0.631 0.208 Z',
  pinDot:
    'M 0.497 1 C 0.225 1 0 0.775 0 0.497 C 0 0.225 0.225 0 0.497 0 C 0.775 0 1 0.225 1 0.497 C 1 0.775 0.775 1 0.497 1 Z M 0.497 0.053 C 0.257 0.053 0.053 0.257 0.053 0.497 C 0.053 0.743 0.257 0.941 0.497 0.941 C 0.743 0.941 0.941 0.743 0.941 0.497 C 0.941 0.257 0.743 0.053 0.497 0.053 Z',
  pinRing:
    'M 0.5 1 C 0.224 1 0 0.776 0 0.5 C 0 0.228 0.224 0 0.5 0 C 0.776 0 1 0.228 1 0.5 C 1 0.776 0.776 1 0.5 1 Z M 0.5 0.016 C 0.232 0.016 0.016 0.232 0.016 0.5 C 0.016 0.768 0.232 0.988 0.5 0.988 C 0.768 0.988 0.988 0.768 0.988 0.5 C 0.988 0.232 0.768 0.016 0.5 0.016 Z',
};

// --------------------------------------------------------------- helpers
// custGeom points are relative to the shape's own box, hence no x/y offset.
const pt = (fx, fy, w, h) => ({ x: +(fx * w).toFixed(4), y: +(fy * h).toFixed(4) });

/** Filled polygon from POLYGONS, placed at x/y/w/h. */
function poly(s, name, x, y, w, h, fill) {
  s.addShape('custGeom', {
    x, y, w, h, fill, line: { type: 'none' },
    points: POLYGONS[name].map(([fx, fy]) => pt(fx, fy, w, h)).concat([{ close: true }]),
  });
}

/** Filled polygon from an inline fraction list. */
function polyPts(s, pts, x, y, w, h, fill) {
  s.addShape('custGeom', {
    x, y, w, h, fill, line: { type: 'none' },
    points: pts.map(([fx, fy]) => pt(fx, fy, w, h)).concat([{ close: true }]),
  });
}

/** Multi-subpath line-art glyph from GLYPHS. */
function icon(s, name, x, y, w, h, fill) {
  const pts = [];
  for (const seg of GLYPHS[name].split(' ').reduce(collectSegments, [])) {
    const [op, nums] = seg;
    if (op === 'Z') { pts.push({ close: true }); continue; }
    const p = [];
    for (let i = 0; i < nums.length; i += 2) p.push(pt(nums[i], nums[i + 1], w, h));
    if (op === 'M') pts.push({ ...p[0], moveTo: true });
    else if (op === 'L') pts.push(p[0]);
    else if (op === 'C') pts.push({ x: p[2].x, y: p[2].y, curve: { type: 'cubic', x1: p[0].x, y1: p[0].y, x2: p[1].x, y2: p[1].y } });
    else if (op === 'Q') pts.push({ x: p[1].x, y: p[1].y, curve: { type: 'quadratic', x1: p[0].x, y1: p[0].y } });
  }
  s.addShape('custGeom', { x, y, w, h, fill, line: { type: 'none' }, points: pts });
}

function collectSegments(acc, token) {
  if (/^[MLCQZ]$/.test(token)) acc.push([token, []]);
  else acc[acc.length - 1][1].push(parseFloat(token));
  return acc;
}

/** Body copy. Defaults match the template's 36pt Montserrat / 54pt leading. */
function txt(s, text, x, y, w, h, o = {}) {
  const opts = {
    x, y, w, h, margin: 0, fontFace: o.fontFace || BODY,
    fontSize: o.fontSize || 36, color: o.color || NAVY,
    align: o.align || 'left', valign: o.valign || 'top',
    lineSpacing: o.lineSpacing === null ? undefined : (o.lineSpacing || 54),
  };
  if (o.bold) opts.bold = true;
  if (o.charSpacing) opts.charSpacing = o.charSpacing;
  if (o.paraSpaceBefore) opts.paraSpaceBefore = o.paraSpaceBefore;
  if (o.vert) opts.vert = o.vert;
  if (o.wrap === false) opts.wrap = false;
  if (o.fill) { opts.fill = o.fill; opts.shape = o.shape || 'rect'; }
  if (o.shape) opts.shape = o.shape;
  if (o.line) { opts.line = o.line; opts.shape = o.shape || 'rect'; }
  s.addText(o.caps ? text.toUpperCase() : text, opts);
}

/** The big red slide heading (100pt SemiBold, all caps, letter-spaced). */
function head(s, text, x, y, w, h) {
  s.addText(text.toUpperCase(), {
    x, y, w, h, margin: 0, fontFace: HEAD, fontSize: 100, color: RED,
    charSpacing: 12, lineSpacing: 106, valign: 'top',
  });
}

/** Pink page-number chip, bottom-left on every slide but the cover. */
function pageNum(s, num) {
  s.addText(String(num), {
    x: 1.1267, y: 13.3958, w: 0.7874, h: 0.7874, margin: 0,
    shape: 'rect', fill: { color: ROSE },
    align: 'center', valign: 'middle', fontFace: BODY, fontSize: 30, color: NAVY,
  });
}

/**
 * Top-right brand bug: an outlined "folded page" mark above the FASHION
 * wordmark. Redrawn with strokes instead of embedding the original bitmap.
 */
function logo(s) {
  const x = 23.57, y = 0.806, w = 1.9685, h = 1.5276;
  const m = { x: x + 0.563, y, w: 0.835, h: 0.837 };
  const stroke = { color: NAVY, width: 2 };
  const stamp = (frac, close) => s.addShape('custGeom', {
    ...m, fill: { type: 'none' }, line: stroke,
    points: frac.map(([fx, fy]) => pt(fx, fy, m.w, m.h)).concat(close ? [{ close: true }] : []),
  });
  stamp([[0.24, 0], [1, 0], [1, 0.72], [0.75, 1], [0, 1], [0, 0.26]], true);  // chamfered square
  stamp([[0.44, 0], [0.44, 1]]);                                             // spine
  stamp([[0.44, 0.70], [1, 0.70]]);                                          // fold base
  stamp([[0.44, 0.70], [0.66, 0.485], [1, 0.485]]);                          // folded corner
  s.addText('FASHION', {
    x, y: y + 0.55 * h, w, h: 0.45 * h, margin: 0, align: 'center', valign: 'bottom',
    fontFace: BODY, fontSize: 20, color: RED, charSpacing: 4,
  });
}

/** Thin navy leader line (table of contents). */
function rule(s, x, y, w) {
  s.addShape('line', { x, y, w, h: 0, line: { color: NAVY, width: 2 } });
}

// ------------------------------------------------------------ data table
// "Light Style 2 - Accent 2": rose header band, alternating rose rows.
function table(s, o) {
  const cellBase = { fontFace: BODY, fontSize: 36, color: NAVY, valign: 'middle', margin: [0.08, 0.39, 0.08, 0.39] };
  const rows = [o.head.map((t, i) => ({
    text: t, options: { ...cellBase, bold: true, align: o.align[i], fill: { color: ROSE } },
  }))];
  o.body.forEach((r, ri) => {
    const banded = ri % 2 === 1;
    rows.push(r.map((t, i) => ({
      text: t,
      options: { ...cellBase, align: o.align[i], fill: banded ? { color: ROSE, transparency: 70 } : { color: BG } },
    })));
  });
  s.addTable(rows, { x: o.x, y: o.y, colW: o.colW, rowH: o.rowH, border: { type: 'none' } });
}

// ----------------------------------------------------------------- charts
const AXIS = {
  catAxisLabelFontFace: BODY, catAxisLabelFontSize: 20, catAxisLabelColor: NAVY,
  catAxisLineColor: NAVY, valAxisLabelFontFace: BODY, valAxisLabelFontSize: 20,
  valAxisLabelColor: NAVY, valAxisLineColor: NAVY, showLegend: false,
};

/** Donut gauge: navy arc for the value, rose for the remainder, % in the hole. */
function ringChart(s, o) {
  s.addChart('doughnut', [{ name: 'Ratio', labels: ['value', 'rest'], values: [o.value, 1 - o.value] }], {
    x: o.x, y: o.y, w: o.w, h: o.h, holeSize: 75,
    chartColors: [NAVY, ROSE], showLegend: false, showValue: false,
    dataBorder: { pt: 0, color: BG },
    layout: { x: 0.131, y: 0.131, w: 0.739, h: 0.741 },   // matches the source deck's inset plot area
  });
  s.addText(Math.round(o.value * 100) + '%', {
    x: o.x, y: o.y, w: o.w, h: o.h, margin: 0, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: o.labelSize, color: NAVY, bold: true,
  });
}

/** Full-width progress bar: rose fill for the achieved share, sky for the rest. */
function progressBar(s, o) {
  const done = o.w * (o.pct / 100);
  s.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: SKY, transparency: 70 }, line: { type: 'none' } });
  s.addShape('rect', { x: o.x, y: o.y, w: done, h: o.h, fill: { color: ROSE }, line: { type: 'none' } });
  s.addText(o.label, {
    x: o.x, y: o.y, w: done, h: o.h, margin: 0, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: o.labelSize, color: RED,
  });
}

function barChart(s, o) {
  s.addChart('bar', [{ name: 'Series 1', labels: o.labels, values: o.values }], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    barDir: o.dir === 'h' ? 'bar' : 'col', barGapWidthPct: o.gap,
    chartColors: o.colors, ...AXIS,
    showValue: true, dataLabelFontFace: HEAD, dataLabelFontSize: 36,
    dataLabelColor: BG, dataLabelPosition: 'ctr', dataLabelFormatCode: o.labelFormat || 'General',
    valAxisMaxVal: o.valMax, valAxisMajorUnit: o.valUnit,
    valGridLine: o.valGrid ? { color: 'D9D9D9', size: 1 } : { style: 'none' },
    catGridLine: { style: 'none' },
  });
}

function lineChart(s, o) {
  s.addChart('line', o.series.map(sr => ({ name: sr.name, labels: o.labels, values: sr.values })), {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: o.series.map(sr => sr.color),
    lineSize: 8, lineDataSymbol: 'none', lineSmooth: false, ...AXIS,
    showLegend: true, legendPos: o.legendPos, legendFontFace: BODY, legendFontSize: 28, legendColor: NAVY,
    valAxisMajorUnit: 1, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
}

function areaChart(s, o) {
  s.addChart('area', o.series.map(sr => ({ name: sr.name, labels: o.labels, values: sr.values })), {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: o.series.map(sr => sr.color), ...AXIS,
    showLegend: true, legendPos: o.legendPos, legendFontFace: BODY, legendFontSize: 28, legendColor: NAVY,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
}

function pieChart(s, o) {
  s.addChart('pie', [{ name: 'Sales', labels: o.labels, values: o.values }], {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: o.colors,
    showLegend: true, legendPos: 'r', legendFontFace: BODY, legendFontSize: 36, legendColor: NAVY,
    showPercent: true, showValue: false, dataLabelFontFace: HEAD, dataLabelFontSize: o.labelSize,
    dataLabelColor: BG, dataLabelPosition: 'ctr', dataLabelFormatCode: '0%',
    dataBorder: { pt: 3, color: BG },
  });
}

// ------------------------------------------------- shared slide furniture
// Every slide is dressed with the same catalogue of corner / stripe motifs.
const DECOR = {
  brRose:   ['cornerBR', 20.566, 8.906, 6.095, 6.097, ROSE],
  brSky:    ['cornerBR', 20.566, 8.906, 6.095, 6.097, SKY],
  tlSky:    ['cornerTL', 0.003, 0.003, 6.092, 6.094, SKY],
  tlRose:   ['cornerTL', 0.003, 0.003, 6.092, 6.094, ROSE],
  stripe:   ['stripe', 3.269, 0.003, 20.123, 15, SKY],
  sashSky:  ['sash', 10.759, 0.003, 15.906, 15, SKY],
  brBig:    ['cornerBRbig', 19.113, 7.455, 7.549, 7.549, ROSE],
};

function decorate(s, keys) {
  keys.forEach(k => {
    const [name, x, y, w, h, color] = DECOR[k];
    poly(s, name, x, y, w, h, { color, transparency: 70 });
  });
}


// ------------------------------------------------------------ -- "Title Slide"
function slide01(s) {
  decorate(s, ['stripe', 'brSky', 'tlSky']);
  poly(s, 'cornerTLbig', 7.217, 0.003, 7.542, 7.545, { color: ROSE, transparency: 70 });
  poly(s, 'cornerBRbig', 11.903, 7.455, 7.542, 7.549, { color: ROSE, transparency: 70 });
  s.addShape('rect', { x: 7.217, y: -0.002, w: 12.227, h: 15.003, fill: { color: WHITE, transparency: 50 }, line: { type: 'none' } });
  logo(s);
  txt(s, 'FASH\nION\nSHOP', 8.51, 0.806, 9.646, 13.378, { fontSize: 169, fontFace: HEAD, valign: 'middle', charSpacing: 40, lineSpacing: 320 });
  txt(s, 'Clothes\nshoes\njewelry', 14.514, 6.61, 2.755, 2.188, { fontSize: 23, color: RED, align: 'right', valign: 'middle', caps: true, charSpacing: 6, lineSpacing: 52 });
  txt(s, 'Your@name.com', 1.153, 0.806, 0.693, 5.514, { vert: 'vert270' });
  txt(s, '\uf0e0', 1.296, 6.888, 0.449, 0.485, { fontSize: 32, fontFace: ICONFONT, align: 'center', lineSpacing: null, vert: 'vert270', wrap: false });
  txt(s, '\uf39e', 25.025, 12.861, 0.28, 0.485, { fontSize: 32, fontFace: ICONFONT, align: 'center', lineSpacing: null, vert: 'vert270', wrap: false });
  txt(s, '\uf245 ', 1.272, 12.926, 0.456, 0.485, { fontSize: 32, fontFace: ICONFONT, align: 'center', lineSpacing: null, vert: 'vert270', wrap: false });
  txt(s, 'Yourwebsite.com', 1.153, 7.925, 0.693, 4.897, { vert: 'vert270' });
  txt(s, 'Companyname', 24.819, 7.925, 0.693, 4.897, { vert: 'vert270' });
}

// ------------------------------------------------------------ -- "Table of contents"
function slide02(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 2);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat.', 3.49, 3.562, 19.686, 1.45);
  txt(s, 'Introduction\nStories\nInfographics\nSummary\nContacts', 6.639, 6.575, 10.63, 3.787);
  txt(s, '02\n04\n28\n44\n48', 17.269, 6.575, 2.757, 3.787, { align: 'right' });
  head(s, 'Table Of Contents', 3.49, 1.682, 19.686, 1.487);
  rule(s, 10.077, 7.116, 9.081);
  rule(s, 8.651, 7.854, 10.507);
  rule(s, 10.187, 8.592, 8.971);
  rule(s, 9.443, 9.346, 9.714);
  rule(s, 9.285, 10.101, 9.872);
}

// ------------------------------------------------------------ -- "Agenda"
function slide03(s) {
  decorate(s, ['brRose', 'tlSky']);
  txt(s, '9:00', 4.784, 5.757, 3.318, 3.318, { fontSize: 70, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '11:00', 11.673, 5.757, 3.318, 3.318, { fontSize: 70, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '14:00', 18.563, 5.757, 3.318, 3.318, { fontSize: 70, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  pageNum(s, 3);
  logo(s);
  txt(s, 'Soleat tamquam omnesque nam et, sed ex nibh epicuri, numquam forensibus quo in. Lorem ipsum dolor sit amet, cons tetuer adipiscing.', 3.49, 3.562, 19.686, 1.45);
  head(s, 'agenda', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Duo ei eirmod vertem singulis. Ridens possim ullamcorper an vel. Ea ferri ridens.', 3.49, 9.567, 5.906, 3.029);
  txt(s, 'Ea sed hinc euismod, posidonium. Veritus mentitum te. Ius oratio elicatissimi.', 17.269, 9.567, 5.906, 3.029);
  txt(s, 'Ne mei novum possit accusam, eum percipit ponderum et. Vim wisi intellegat.', 10.379, 9.567, 5.906, 3.029);
}

// ------------------------------------------------------------ -- "Heading"
function slide04(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 4);
  logo(s);
  txt(s, 'Ex veniam consectetuer vel, ne mel augue nusquam pertinacia. Verear numquam vel in, dico purto principes ne vix. His elit sadipscing in, nam cu justo nostro hend rerit, vis no iuvaret consetetur. Iisque scribentur te eum, duo id wisi postea pertinax. Ne per soluta noster platonem. Est discere lobortis id, no vis sale aliqu ando, no lorem simul omittantur duo. Qui te diam utroque cum ad, et quo stet appetere lobortis.', 3.49, 5.812, 19.686, 4.479);
  head(s, 'enjoy it!', 3.49, 3.932, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Heading with background"
function slide05(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 5);
  logo(s);
  txt(s, 'Nulla integre disputando cum ad, et quo stet appetere. Qui te diam utroque adipiscing, eos ut duis eirmod dissentiunt. Ne altera pertinacia delicatissimi sit, eu tale intellegat nec, ex saepe impedit deleniti est. In sea eros nostro senserit. Et nemore accommodare cum, per an purto elit sententiae, nonumes eligendi adipiscing sea eros nostro eu.', 3.49, 3.562, 19.686, 3.722);
  head(s, 'Fashion', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Heading with one column layout"
function slide06(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 6);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad. Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed.\n\nEuismod tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad aliquip ex ea commodo consequat. Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie feugiat nulla facilisis.', 3.49, 5.812, 19.686, 5.301);
  head(s, 'One column', 3.49, 3.932, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Heading with one column layout and background"
function slide07(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 7);
  logo(s);
  txt(s, 'At vix solum libris, pro at essent mandamus, in nec civibus mandamus. Ei nam illum vidisse sadipscing, quis elaboraret ex eum. Eum defini ebas comprehensam at, te epicuri orem ipsum dol.\n\nOr sit amet, cons tetuer adipis cing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu ipsum dolor sit amet, consectetuer adipiscing elit sed diam. Euismod tincidunt ut laoreet dolore magna. Ut wisi enim ad aliqui.', 3.49, 3.562, 11.024, 9.088);
  head(s, 'Our work', 3.49, 1.682, 11.024, 1.487);
}

// ------------------------------------------------------------ -- "Heading with multi column layout"
function slide08(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 8);
  logo(s);
  txt(s, 'Id regione deserunt quaestio pro, ad affert fuisset mnesarchum vix. Sed ei utamur indoctum. Ne mei novum pos sit accusam, eum percipit ponderum assueverit et. Vim wisi intellegat at. Id regione deserunt mnesarchum vix fuisset.', 3.49, 5.812, 9.252, 5.301);
  txt(s, 'Lorem ipsum dolor sit amet, cons tet er adipiscing elit, sed diam nonum my nibh euis mod tincidunt ut laoreet dol ore ag na aliqu am erat volutpat. Ut wisi enim ad. Lorem ipsum dolor sit amet, elit, sed diam nonummy nibh ore ag na.', 13.923, 5.812, 9.252, 5.301);
  head(s, 'TWO column', 3.49, 3.932, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Heading with multi column layout and background"
function slide09(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 9);
  logo(s);
  txt(s, 'Ne mei novum pos sit accusam, eum percipit ponderum assueverit et. Vim wisi intellegat at. Id regione deserunt mnesarchum vix fuisset.\nLorem diam nonum my nibh.', 3.49, 3.563, 9.252, 3.812);
  txt(s, 'Mod tincidunt ut laoreet dol ore ag na aliqu am erat volutpat. Ut wisi enim ad. Lorem ipsum amet, elit, sed diam  dolor sit amet, cons tet nonummy nibh ore ag na.', 13.923, 3.563, 9.252, 3.812);
  head(s, 'Shopping', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Cover slide collage - 1"
function slide10(s) {
  decorate(s, ['tlRose', 'sashSky', 'brBig']);
  poly(s, 'cornerTLbig', 14.797, 0.003, 7.549, 7.545, { color: ROSE, transparency: 70 });
  pageNum(s, 10);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat.\n\nNostrud exerci tation ullamcorper sus cipit lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, consectetuer adipi cing elit, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat ed.', 3.49, 3.562, 11.024, 9.088);
  head(s, 'Why us?', 3.49, 1.682, 11.024, 1.487);
}

// ------------------------------------------------------------ -- "Cover slide collage - 2"
function slide11(s) {
  decorate(s, ['sashSky', 'brBig']);
  pageNum(s, 11);
  logo(s);
  txt(s, 'Nulla integre disputando cum ad, et quo stet appetere. Qui te diam utroque sent tiae, nonumes eligendi adipiscing usu id. Ius id ridens evertitur adipiscing, eos ut duis eirmod dissentiunt stet appetere.\n\nAt vix solum exerci tation ullamcorper sus cipit lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, consectetuer adipi cing elit, consectetuer adipiscing elit, aliquam erat ed.', 12.975, 3.562, 10.2, 9.088);
  head(s, 'about us?', 12.975, 1.682, 10.2, 1.487);
  poly(s, 'cornerTL', 0, 0, 6.247, 6.247, { color: SKY, transparency: 70 });
  poly(s, 'cornerBRbig', 4.392, 4.561, 7.739, 7.739, { color: ROSE, transparency: 70 });
}

// ------------------------------------------------------------ -- "Right image layout"
function slide12(s) {
  decorate(s, ['tlRose', 'sashSky', 'brBig']);
  pageNum(s, 12);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetu er adipiscing elit, sed diam nonummy nibh euis mod tincidunt dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat.\n\nNostrud exerci tation ullamcorper sus cipit lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, consectetuer adipi cing elit, sed. Lorem ipsum dolor sit met, adipiscing elit, sed diam', 3.49, 3.562, 9.885, 8.331);
  head(s, 'fashion', 3.49, 1.682, 9.885, 1.487);
}

// ------------------------------------------------------------ -- "Left image layout"
function slide13(s) {
  decorate(s, ['sashSky', 'brBig']);
  pageNum(s, 13);
  logo(s);
  txt(s, 'Ne altera pertinacia delicatissimi sit, eu tale intellegat nec, ex saepe impedit um deleniti est. In sea eros nostro senserit. Et nemore accommodare cum, per an purto elit consectetuer sed. Lorem ipsum dolor sit met, tetuer.', 12.975, 5.835, 10.2, 4.479);
  head(s, 'shop', 12.975, 3.955, 10.2, 1.487);
  poly(s, 'cornerBRbig', 4.42, 6.045, 7.731, 7.726, { color: ROSE, transparency: 70 });
  poly(s, 'cornerTLbig', 0, 1.479, 7.727, 7.729, { color: SKY, transparency: 70 });
}

// ------------------------------------------------------------ -- "Top image layout"
function slide14(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 14);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat.', 3.49, 11.082, 19.686, 2.207);
  head(s, 'Clothes', 3.49, 9.201, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Bottom image layout"
function slide15(s) {
  decorate(s, ['tlRose']);
  poly(s, 'cornerBR', 20.566, 1.474, 6.095, 6.097, { color: ROSE, transparency: 70 });
  poly(s, 'stripe', 2.089, 0.003, 20.123, 15, { color: SKY, transparency: 70 });
  pageNum(s, 15);
  logo(s);
  txt(s, 'Qui te diam utroque sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat euis mod tincidunt ut laoreet dolore ag am erat. Lorem ipsum tetuer nonummy nibh.', 3.49, 3.562, 19.686, 2.965);
  head(s, 'Clothes', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "What we do"
function slide16(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 16);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis sus cipit lobortis nisl ut aliquip ex.', 3.49, 3.562, 19.686, 2.207);
  head(s, 'What we do', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Ridens possim er an vel ferri.', 5.445, 10.324, 4.385, 1.45);
  txt(s, 'Ea sed hinc euis mod, posid.', 11.14, 10.324, 4.385, 1.45);
  txt(s, 'Veritus men tit um te us.', 16.835, 10.324, 4.385, 1.45);
  s.addShape('rect', { x: 5.892, y: 6.614, w: 3.491, h: 3.427, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'hat', 6.38, 7.597, 2.526, 1.431, { color: NAVY });
  s.addShape('rect', { x: 11.587, y: 6.614, w: 3.491, h: 3.427, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'sunglasses', 12.075, 7.743, 2.516, 1.036, { color: NAVY });
  s.addShape('rect', { x: 17.282, y: 6.598, w: 3.491, h: 3.427, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'watch', 17.634, 7.665, 2.79, 1.115, { color: NAVY });
  icon(s, 'watchHand', 18.768, 8.045, 0.292, 0.385, { color: NAVY });
}

// ------------------------------------------------------------ -- "Goals"
function slide17(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 17);
  logo(s);
  txt(s, 'Eum definiebas comprehensam at, te probo omnis fuisset qui. Ut duo assum um epicuri ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat nisl ut aliquip ex ag na alique.', 3.49, 3.562, 19.686, 2.272);
  head(s, 'our goals', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Sed ei utamur ind im veni am, qis am aliqu am erat susy et.', 17.663, 7.317, 5.512, 2.272);
  txt(s, 'Trends', 17.663, 6.585, 5.512, 0.693, { bold: true });
  txt(s, 'Lorem ipsum mi im veni am, qis am aliqu am erat susy.', 17.663, 11.066, 5.512, 2.272);
  txt(s, 'Growth', 17.663, 10.334, 5.512, 0.693, { bold: true });
  txt(s, 'Ut wisi enim ad mi im veni am, qis am aliqu am erat susy.', 6.639, 7.317, 5.512, 2.272);
  txt(s, 'Profit', 6.639, 6.585, 5.512, 0.693, { bold: true });
  txt(s, 'Ex veniam consec tetuerim veni am, qis am aliqu amy.', 6.639, 11.066, 5.512, 2.272);
  txt(s, 'Quality', 6.639, 10.334, 5.512, 0.693, { bold: true });
  s.addShape('rect', { x: 3.378, y: 6.616, w: 2.885, h: 2.833, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'dollar', 4.276, 7.12, 1.083, 1.83, { color: NAVY });
  icon(s, 'dollarStemTop', 4.859, 8.139, 0.219, 0.387, { color: NAVY });
  icon(s, 'dollarStemBot', 4.618, 7.533, 0.219, 0.361, { color: NAVY });
  s.addShape('rect', { x: 3.394, y: 10.412, w: 2.885, h: 2.833, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'check', 4.007, 11.117, 1.925, 1.424, { color: NAVY });
  s.addShape('rect', { x: 14.385, y: 6.616, w: 2.889, h: 2.833, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'star', 14.961, 7.205, 1.736, 1.649, { color: NAVY });
  icon(s, 'starInner', 15.338, 7.599, 0.984, 0.938, { color: NAVY });
  s.addShape('rect', { x: 14.38, y: 10.412, w: 2.889, h: 2.833, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'arrow', 15.078, 11.088, 1.488, 1.49, { color: NAVY });
}

// ------------------------------------------------------------ -- "Values"
function slide18(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 18);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed nonummy nibh.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'Our Values', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'PLACEHOLDER', 3.489, 9.564, 7.111, 3.029);
  txt(s, 'Sed ei utamur ind im veni am, sectetuer adip or cing lorem ipsum et con sectetuer adip or', 16.014, 9.564, 7.111, 3.029);
  txt(s, 'PLACEHOLDER', 9.777, 5.086, 7.111, 3.029);
  txt(s, '1', 5.385, 5.343, 3.318, 3.318, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '2', 11.673, 9.275, 3.318, 3.318, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '3', 17.911, 5.343, 3.318, 3.318, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
}

// ------------------------------------------------------------ -- "Team"
function slide19(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 19);
  logo(s);
  txt(s, 'Soleat tamquam omnesque nam et, sed ex nibh epicuri, numquam forensib.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'Our team', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Janie Doe', 3.49, 11.074, 5.906, 0.693, { bold: true, align: 'center' });
  txt(s, 'Janie Doe', 17.269, 11.074, 5.906, 0.693, { bold: true, align: 'center' });
  txt(s, 'Jave Doe', 10.379, 11.074, 5.906, 0.693, { bold: true, align: 'center' });
  txt(s, 'seller', 3.49, 11.882, 5.906, 0.693, { align: 'center' });
  txt(s, 'manager', 17.269, 11.882, 5.906, 0.693, { align: 'center' });
  txt(s, 'manager', 10.379, 11.882, 5.906, 0.693, { align: 'center' });
}

// ------------------------------------------------------------ -- "Testimonials"
function slide20(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 20);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipisуcing elit, sed diam nonummy nibh. Lorem sum dolor sit amet, consectetuer adipiscing elit, sed diam nonum my nibh euismod tinci dunt ut lao reet dolore magna aliquam erat vol pat. Ut wisi enim ad minim ullamcorper suscipit lobor tis nisl ut aliquip ex ea commodo sequat aliquam.', 11.757, 3.562, 11.418, 6.059);
  head(s, 'Testimonials', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Janie Doe', 11.757, 10.317, 11.418, 0.693, { bold: true, align: 'right' });
}

// ------------------------------------------------------------ -- "Partners"
function slide21(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 21);
  logo(s);
  txt(s, 'Nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis sus cipit lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, cons sed diam nonummy nibh.', 3.49, 3.562, 19.686, 2.272);
  head(s, 'Our clients', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Quotation"
function slide22(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 22);
  logo(s);
  txt(s, 'Ne mei novum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad. Lorem ipsum adipiscing elit, sed diam nonummy nibh.\n\nEuismod tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad aliquip ex ea commodo consequat. Duis autem vel eum iriure dolor in hendrerit in feugiat nulla facilisis tincidunt ut laoreet dolore.', 3.49, 4.312, 19.686, 5.301);
  txt(s, 'Janie Doe', 3.49, 10.421, 19.686, 0.693, { bold: true });
  icon(s, 'quoteLeft', 7.45, 2.066, 1.361, 1.601, { color: SKY, transparency: 70 });
  icon(s, 'quoteRight', 8.948, 2.066, 1.361, 1.601, { color: SKY, transparency: 70 });
  icon(s, 'quoteLeft', 16.475, 11.403, 1.361, 1.601, { color: ROSE, transparency: 70 });
  icon(s, 'quoteRight', 17.972, 11.403, 1.361, 1.601, { color: ROSE, transparency: 70 });
}

// ------------------------------------------------------------ -- "Gallery with 1 image"
function slide23(s) {
  decorate(s, ['brSky', 'tlRose']);
  s.addShape('rect', { x: 0, y: 0, w: 11.757, h: 15, fill: { color: WHITE, transparency: 60 }, line: { type: 'none' } });
  pageNum(s, 23);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh euis mod tincidunt ut laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim.\n\nEuismod tincidunt ut laoreet dolore mag na aliquam erat volutpat. Ut wisi enim ad aliquip ex ea commodo conse quat. Duis autem dolor in hend.', 12.975, 4.314, 10.2, 6.816);
  poly(s, 'cornerBR', 5.661, 8.906, 6.095, 6.097, { color: ROSE, transparency: 70 });
}

// ------------------------------------------------------------ -- "Gallery with 3 images"
function slide24(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 24);
  logo(s);
}

// ------------------------------------------------------------ -- "Gallery with 4 images"
function slide25(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 25);
  logo(s);
}

// ------------------------------------------------------------ -- "Gallery with 5 images"
function slide26(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 26);
  logo(s);
}

// ------------------------------------------------------------ -- "Progress"
function slide27(s) {
  decorate(s, ['brRose', 'tlSky']);
  progressBar(s, { x: 3.885, y: 6.431, w: 18.895, h: 4.11, pct: 88, label: '88% growth', labelSize: 66 });
  pageNum(s, 27);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat. lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, cons sed diam.', 3.49, 3.562, 19.686, 2.272);
  head(s, 'Our progress', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Ex veniam consectetuer vel, ne mel augue nusquam pertinacia. Verear num quam ne vix. His elit sadipscing in, nam cu justo nostro rerit, vis.', 3.49, 11.082, 19.686, 1.45);
}

// ------------------------------------------------------------ -- "Ratio"
function slide28(s) {
  decorate(s, ['brRose', 'tlSky']);
  ringChart(s, { x: 3.707, y: 6.742, w: 6.038, h: 6.026, value: 0.25, labelSize: 50 });
  ringChart(s, { x: 10.313, y: 6.742, w: 6.038, h: 6.026, value: 0.56, labelSize: 50 });
  ringChart(s, { x: 16.92, y: 6.742, w: 6.038, h: 6.026, value: 0.8, labelSize: 50 });
  pageNum(s, 28);
  logo(s);
  txt(s, 'Labore euripidis usu no, wisi adipisci conclusionemque ne sit. Augue exno et semper causae virtute. An augue labitur sit. Option dolores ex vis, et decore audire eum. Te has docendi copiosae, mei aliquando comprehensam at. Cum no numquam minimum, noluisse ullamcorper ei usu.', 3.49, 3.562, 19.686, 3.029);
  head(s, 'sales ratio', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Numeric information"
function slide29(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 29);
  logo(s);
  txt(s, 'customer satisfaction', 3.882, 9.334, 18.889, 2.104, { fontSize: 100, fontFace: HEAD, align: 'center', lineSpacing: 150 });
  txt(s, '99%', 3.882, 3.168, 18.889, 6.165, { fontSize: 400, color: RED, fontFace: HEAD, align: 'center', valign: 'bottom', lineSpacing: null });
}

// ------------------------------------------------------------ -- "Column chart"
function slide30(s) {
  decorate(s, ['brRose', 'tlSky']);
  barChart(s, { x: 7.966, y: 4.755, w: 10.733, h: 9.17, dir: 'v', gap: 5, valMax: 100, valUnit: 25, valGrid: true, labelFormat: '[$$-409]#,##0',
    labels: ['2015', '2016', '2017', '2018'],
    values: [50, 90, 75, 80],
    colors: [NAVY, ROSE, NAVY, ROSE] });
  pageNum(s, 30);
  logo(s);
  txt(s, 'An augue labitur sit. Option dolores ex vis, et decore audire eum. Te has doce.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'Annual income', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Bar chart"
function slide31(s) {
  decorate(s, ['brRose', 'tlSky']);
  barChart(s, { x: 6.836, y: 4.556, w: 12.993, h: 9.451, dir: 'h', gap: 8, valMax: 2000, valUnit: 500,
    labels: ['January', 'February', 'March', 'April', 'May', 'June', 'July'],
    values: [1200, 1400, 1800, 1350, 1220, 1800, 1050],
    colors: [NAVY, ROSE, NAVY, ROSE, NAVY, ROSE, NAVY] });
  pageNum(s, 31);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed nonummy nibh.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'profit', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Line chart"
function slide32(s) {
  decorate(s, ['brRose', 'tlSky']);
  lineChart(s, { x: 6.836, y: 3.562, w: 12.993, h: 10.042, legendPos: 'b',
    labels: ['2016', '2017', '2018', '2019'],
    series: [
      { name: ' Income', color: ROSE, values: [4.3, 2.5, 3.5, 4.5] },
      { name: ' Costs', color: NAVY, values: [2.4, 4.4, 0.5, 2.8] },
    ] });
  pageNum(s, 32);
  logo(s);
  head(s, 'our sales', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Area chart"
function slide33(s) {
  decorate(s, ['brRose', 'tlSky']);
  areaChart(s, { x: 6.836, y: 4.556, w: 12.993, h: 9.451, legendPos: 'r',
    labels: ['2016', '2017', '2018', '2019', '2020'],
    series: [
      { name: 'Income', color: NAVY, values: [14, 18, 15, 20, 16] },
      { name: 'Costs', color: ROSE, values: [10, 14, 8, 12, 10] },
    ] });
  pageNum(s, 33);
  logo(s);
  txt(s, 'Dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh ipsum.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'profit', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Pie chart"
function slide34(s) {
  decorate(s, ['brRose', 'tlSky']);
  pieChart(s, { x: 6.245, y: 5.396, w: 14.004, h: 7.593, labelSize: 50,
    labels: ['Goods', 'Services', 'Other'], values: [0.35, 0.4, 0.25], colors: [ROSE, RED, NAVY] });
  pageNum(s, 34);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed nonummy nibh.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'growth graph', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Market share"
function slide35(s) {
  decorate(s, ['brRose', 'tlSky']);
  ringChart(s, { x: 4.469, y: 4.714, w: 7.682, h: 7.667, value: 0.25, labelSize: 80 });
  ringChart(s, { x: 14.514, y: 4.714, w: 7.682, h: 7.667, value: 0.5, labelSize: 80 });
  pageNum(s, 35);
  logo(s);
  txt(s, 'Quo et semper causae virtute. An augue labitur sit. Option dolores ex visum.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'Our Growth', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'An augue labitur sit.', 3.49, 12.305, 9.835, 0.693, { align: 'center' });
  txt(s, 'Option dolores ex visum.', 13.325, 12.305, 9.835, 0.693, { align: 'center' });
}

// ------------------------------------------------------------ -- "List of items"
function slide36(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 36);
  logo(s);
  txt(s, 'Comprehensam at, te probo omnis fuisset um epicuri ut laoreet dolore ag na.', 6.639, 5.087, 5.512, 3.029);
  head(s, 'List of items', 3.49, 3.206, 19.686, 1.487);
  txt(s, 'At vix solum libris, pro at essent mandamus, in nec civibus us. Ei nam illum vidisse.', 17.664, 5.087, 5.512, 3.029);
  txt(s, 'Lorem ipsum dol or sit et, consectetuer ipsum tuer adipis consectetuer.', 6.639, 8.814, 5.512, 2.965);
  txt(s, 'Nulla integre dispu tando cum ad, et quo stet appetere ret con setetur nostro.', 17.664, 8.814, 5.512, 3.029);
  txt(s, '1', 3.489, 5.244, 2.756, 2.756, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '2', 14.513, 5.244, 2.756, 2.756, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '3', 3.489, 8.948, 2.756, 2.756, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
  txt(s, '4', 14.513, 8.948, 2.756, 2.756, { fontSize: 100, color: RED, fontFace: HEAD, align: 'center', valign: 'middle', lineSpacing: null, fill: { color: ROSE } });
}

// ------------------------------------------------------------ -- "List of items-2"
function slide37(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 37);
  logo(s);
  txt(s, 'Est an quis nostro possim, populo patrioque ut mel, ad amet accusam similique cum. Vix virtute accommodare ad, mea atqui prodesset ad, pri ut dicat populo. ', 6.639, 5.087, 16.536, 2.207);
  head(s, 'Our items', 3.49, 3.206, 19.686, 1.487);
  txt(s, 'Has antiopam intellegebat signiferumque te, an duo petentium quaerendum, tantas admodum accusam similique noluisse mea no. Wisi vidisse at quo.', 6.639, 8.08, 16.536, 2.207);
  s.addShape('rect', { x: 3.882, y: 8.08, w: 2.398, h: 2.354, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'check', 4.391, 8.666, 1.6, 1.183, { color: NAVY });
  s.addShape('rect', { x: 3.882, y: 5.087, w: 2.398, h: 2.354, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'check', 4.391, 5.672, 1.6, 1.183, { color: NAVY });
}

// ------------------------------------------------------------ -- "List of images"
function slide38(s) {
  decorate(s, ['stripe', 'brRose', 'tlRose']);
  pageNum(s, 38);
  logo(s);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'Our gallery', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Id regione deser un quaestio pro, adaf fert fuisset archum.', 3.49, 11.074, 5.906, 2.207);
  txt(s, 'Ne mei novum pos sit accusam, eum percipit ponderum.', 17.269, 11.074, 5.906, 2.207);
  txt(s, 'Lorem ipsum dol or sit et, consectetuer adipis cingy lorem.', 10.379, 11.074, 5.906, 2.207);
}

// ------------------------------------------------------------ -- "Services"
function slide39(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 39);
  logo(s);
  txt(s, 'Nulla integre disputando cum ad, et quo stet appetere. Qui te diam utroque sententiae, nonumes eligendi adipiscing usu id. Ius id ridens evertitur duis eirmod dissentiunt ut laoreet dolore.', 6.639, 3.562, 16.536, 2.207);
  head(s, 'services', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam no sententiae, adipiscing usu id. Ius id ridens ever titur my nibh euis mod tincidunt ut laoreet dolore ag na aliqu amtas.', 6.639, 6.576, 16.536, 2.207);
  txt(s, ' Eum definiebas comprehensam at, te probo omnis fuisset qui. Ut duo assum epicuri.  Ei nam, quis elaboraret ex eum. Est discere lobortis id, no vis sale aliquando adipiscing usu.', 6.639, 9.563, 16.536, 2.272);
  s.addShape('rect', { x: 3.915, y: 3.596, w: 2.363, h: 2.319, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'hat', 4.245, 4.261, 1.71, 0.968, { color: NAVY });
  s.addShape('rect', { x: 3.915, y: 6.598, w: 2.363, h: 2.319, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'sunglasses', 4.245, 7.362, 1.703, 0.701, { color: NAVY });
  s.addShape('rect', { x: 3.882, y: 9.601, w: 2.363, h: 2.319, fill: { color: ROSE, transparency: 60 }, line: { type: 'none' } });
  icon(s, 'watch', 4.12, 10.323, 1.888, 0.754, { color: NAVY });
  icon(s, 'watchHand', 4.888, 10.58, 0.197, 0.261, { color: NAVY });
}

// ------------------------------------------------------------ -- "Pricing"
function slide40(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 40);
  logo(s);
  head(s, 'our pricing', 3.49, 1.682, 19.686, 1.487);
  s.addShape('rect', { x: 9.396, y: 4.277, w: 7.873, h: 7.874, fill: { color: ROSE }, line: { type: 'none' } });
  txt(s, 'Premium', 9.396, 4.277, 7.873, 3.228, { fontSize: 43, bold: true, align: 'center', valign: 'middle', lineSpacing: null });
  s.addShape('rect', { x: 3.49, y: 4.277, w: 5.512, h: 5.512, fill: { type: 'none' }, line: { color: ROSE, width: 10 } });
  txt(s, 'Basic', 3.49, 4.277, 5.512, 3.543, { fontSize: 43, bold: true, align: 'center', valign: 'middle', lineSpacing: null });
  s.addShape('rect', { x: 17.658, y: 4.277, w: 5.512, h: 5.512, fill: { type: 'none' }, line: { color: ROSE, width: 10 } });
  txt(s, 'Pro', 17.658, 4.277, 5.512, 3.543, { fontSize: 43, bold: true, align: 'center', valign: 'middle', lineSpacing: null });
  txt(s, '250$', 9.396, 8.44, 7.873, 1.818, { fontSize: 120, color: RED, fontFace: HEAD, bold: true, align: 'center', caps: true, charSpacing: 6, lineSpacing: null });
  txt(s, '50$', 3.663, 6.586, 5.165, 1.818, { fontSize: 120, color: RED, fontFace: HEAD, bold: true, align: 'center', caps: true, charSpacing: 6, lineSpacing: null });
  txt(s, '90$', 17.838, 6.586, 5.165, 1.818, { fontSize: 120, color: RED, fontFace: HEAD, bold: true, align: 'center', caps: true, charSpacing: 6, lineSpacing: null });
  txt(s, 'Fusce accumsan eros eu nib.', 10.215, 6.707, 6.223, 1.303, { fontSize: 43, align: 'center', lineSpacing: null });
}

// ------------------------------------------------------------ -- "Vertical table"
function slide41(s) {
  decorate(s, ['brRose', 'tlSky']);
  table(s, {
    x: 3.882, y: 4.951, colW: [4.728, 9.44, 4.733], rowH: 1.177, align: ['left', 'left', 'right'],
    head: ['Name ', 'Property', 'Price'],
    body: [
      ['Lorem ipsum', 'Integer in ex varius ', '1000.00'],
      ['Dolor', 'Consectetur lorem', '1500.00'],
      ['Consectetur', 'Duis sit amet porta nisi', '3000.00'],
      ['Adipiscing', 'Mauris non ultrices velit', '4500.00'],
      ['Praesent', 'Aliquam mattis aliquam', '4750.00'],
      ['Tristique', 'Praesent vitae tincidunt', '5000.00'],
    ],
  });
  pageNum(s, 41);
  logo(s);
  txt(s, 'Amet ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'fashion shop ', 3.49, 1.682, 19.686, 1.487);
}

// ------------------------------------------------------------ -- "Horizontal table"
function slide42(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 42);
  logo(s);
  txt(s, 'Lorem dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh.', 3.49, 3.562, 19.686, 0.693);
  head(s, 'our shop ', 3.49, 1.682, 19.686, 1.487);
  table(s, {
    x: 3.911, y: 4.931, colW: [4.728, 10.64, 3.533], rowH: 1.181, align: ['left', 'left', 'right'],
    head: ['March', 'Integer in ex varius, congue ligula', '1000.00'],
    body: [
      ['April', 'Duis sit amet porta nisi mauris non', '1500.00'],
      ['May', 'Aliquam mattis aliquam nibh praesent', '3000.00'],
      ['June', 'Donec eros nisi, sagittis non nulla', '4500.00'],
      ['July', 'Praesent placerat elit id erat rutrum ', '4750.00'],
      ['August', 'Vestibulum convallis vehicula', '5000.00'],
    ],
  });
}

// ------------------------------------------------------------ -- "Timeline"
function slide43(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 43);
  logo(s);
  txt(s, 'Nulla integre disputando cum ad, et quo stet appetere. Qui te diam utroque sententiae, nonumes eligendi adipiscing usu id. Ius id ridens evertitur duis eirmod ut laoreet dolore.', 6.639, 3.562, 16.536, 2.272);
  head(s, 'Time schedule', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam no sententiae, adipiscing usu id. Ius id ridens ever titur my nibh euis mod tincidunt ut laoreet dolore ag na aliqu amtas.', 6.639, 6.576, 16.536, 2.207);
  txt(s, 'Eum definiebas comprehensam at, te probo omnis fuisset qui. Ut duo assum epicuri.  Ei nam, quis elaboraret ex eum. Est discere lobortis id, no vis sale aliquando adipiscing usu.', 6.639, 9.563, 16.536, 2.272);
  txt(s, '12', 3.882, 3.562, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.884, 5.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, '13', 3.882, 6.563, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.884, 8.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, '14', 3.882, 9.563, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.884, 11.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
}

// ------------------------------------------------------------ -- "Timeline with image"
function slide44(s) {
  decorate(s, ['tlRose', 'sashSky', 'brBig']);
  pageNum(s, 44);
  txt(s, 'Nulla integre disputando cum ad, et quo stet appetere. Qui te diam utroque sententiae.', 6.247, 3.562, 7.875, 2.272);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam no sententiae.', 6.247, 6.576, 7.875, 2.272);
  txt(s, 'Eum definiebas omnis fuisset qui. Ut duo assum epicuri.  Ei nam, quis ex eum.', 6.247, 9.563, 7.875, 2.272);
  txt(s, '12', 3.49, 3.562, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.491, 5.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, '13', 3.49, 6.563, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.491, 8.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, '14', 3.49, 9.563, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 3.491, 11.016, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  poly(s, 'cornerTLbig', 14.797, 0.003, 7.549, 7.545, { color: ROSE, transparency: 70 });
  logo(s);
  head(s, 'timeline', 3.49, 1.682, 10.632, 1.487);
}

// ------------------------------------------------------------ -- "Horizontal timeline"
function slide45(s) {
  decorate(s, ['brRose', 'tlRose', 'stripe']);
  pageNum(s, 45);
  logo(s);
  head(s, 'Time scale', 3.49, 1.682, 19.686, 1.487);
  txt(s, '12', 5.248, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 5.249, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Id regione deser un quaestio pro, adaf fert fuisset archum. Lorem ipsum dolor sit amet, cons tetuer.', 3.49, 8.075, 5.906, 3.787);
  txt(s, 'Amet ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy.', 3.49, 3.562, 19.686, 0.693);
  txt(s, '14', 19.028, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 19.029, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Eum definiebas com prehensam at, te probo omnis fuisset qui. Ut duo assum epicuri.  Ei nam, elaboraret.', 17.269, 8.075, 5.906, 3.787);
  txt(s, '13', 12.138, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 12.139, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer diam no sententiae, adipis cing usu id. Ius id ridens ever.', 10.379, 8.075, 5.906, 3.787);
}

// ------------------------------------------------------------ -- "Horizontal timeline with image"
function slide46(s) {
  decorate(s, ['tlRose', 'brRose']);
  poly(s, 'stripe', 1.127, 0.003, 20.123, 15, { color: SKY, transparency: 70 });
  pageNum(s, 46);
  logo(s);
  head(s, 'timeline', 3.49, 1.682, 19.686, 1.487);
  txt(s, '12', 5.248, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 5.249, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer.', 3.49, 8.075, 5.906, 1.45);
  txt(s, 'Amet ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy.', 3.49, 3.562, 19.686, 0.693);
  txt(s, '14', 19.028, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 19.029, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Eum definiebas com prehensam at, te.', 17.269, 8.075, 5.906, 1.45);
  txt(s, '13', 12.138, 5.366, 2.389, 2.389, { fontSize: 80, color: RED, bold: true, align: 'center', lineSpacing: null, fill: { color: ROSE } });
  txt(s, 'March', 12.139, 6.82, 2.387, 0.539, { fontSize: 32, align: 'center', lineSpacing: null });
  txt(s, 'Sententiae, adipis cing usu id. Ius id ridever.', 10.379, 8.075, 5.906, 1.45);
}

// ------------------------------------------------------------ -- "Summary"
function slide47(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 47);
  logo(s);
  txt(s, 'Duis autem dolor sit amet, cons tetu er adipiscing elit, sed diam nonummy nibh euis mod tincidunt dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat.\n\nNostrud exerci tation ullamcorper sus cipit lobortis nisl ut aliquip ex. Lorem ipsum dolor sit amet, consectetuer adipi cing elit, sed. Lorem ipsum dolor sit met, adipiscing elit, sed diam', 3.49, 3.562, 9.885, 8.331);
  head(s, 'summary', 3.49, 1.682, 9.885, 1.487);
}

// ------------------------------------------------------------ -- "Call to action"
function slide48(s) {
  decorate(s, ['sashSky', 'brBig']);
  pageNum(s, 48);
  logo(s);
  txt(s, 'Duo ei eirmod verterem singulis. Ridens possim ullamcorper an vel. Ea ferri ride voluptatum qui. Laoreet dolore ag na aliqu am erat volutpat. Ut wisi enim ad minim veni am, quis aliqu am erat.', 12.975, 6.585, 10.2, 3.722);
  head(s, 'enjoy it!', 12.975, 4.705, 10.2, 1.487);
  poly(s, 'cornerTL', 0, 0, 6.247, 6.247, { color: SKY, transparency: 70 });
  poly(s, 'cornerBRbig', 4.434, 4.575, 7.739, 7.739, { color: ROSE, transparency: 70 });
}

// ------------------------------------------------------------ -- "Locations"
function slide49(s) {
  decorate(s, ['brRose', 'tlSky']);
  pageNum(s, 49);
  logo(s);
  head(s, 'Our Branches', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Amet ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy.', 3.49, 3.562, 19.686, 0.693);
  txt(s, 'Lorem', 6.905, 7.261, 2.294, 0.711, { align: 'center', lineSpacing: 52, wrap: false, fill: { color: ROSE } });
  txt(s, 'Dolor', 14.7, 6.584, 2.025, 0.711, { align: 'center', lineSpacing: 52, wrap: false, fill: { color: ROSE } });
  txt(s, 'Ipsum', 9.296, 10.76, 2.254, 0.711, { align: 'center', lineSpacing: 52, wrap: false, fill: { color: ROSE } });
  s.addShape('ellipse', { x: 7.889, y: 6.591, w: 0.332, h: 0.333, fill: { color: NAVY }, line: { type: 'none' } });
  icon(s, 'pinDot', 7.67, 6.374, 0.767, 0.767, { color: NAVY });
  icon(s, 'pinRing', 7.547, 6.249, 1.01, 1.012, { color: NAVY });
  s.addShape('ellipse', { x: 15.549, y: 5.914, w: 0.332, h: 0.333, fill: { color: NAVY }, line: { type: 'none' } });
  icon(s, 'pinDot', 15.331, 5.697, 0.767, 0.767, { color: NAVY });
  icon(s, 'pinRing', 15.207, 5.572, 1.01, 1.012, { color: NAVY });
  s.addShape('ellipse', { x: 10.26, y: 10.09, w: 0.332, h: 0.333, fill: { color: NAVY }, line: { type: 'none' } });
  icon(s, 'pinDot', 10.041, 9.873, 0.767, 0.767, { color: NAVY });
  icon(s, 'pinRing', 9.918, 9.748, 1.01, 1.012, { color: NAVY });
}

// ------------------------------------------------------------ -- "Follow us"
function slide50(s) {
  decorate(s, ['brRose', 'tlSky']);
  txt(s, 'Instagram', 10.773, 11.75, 5.12, 0.661, { bold: true, align: 'center', lineSpacing: 51 });
  txt(s, 'Twitter', 17.662, 11.75, 5.12, 0.661, { bold: true, align: 'center', lineSpacing: 51 });
  txt(s, 'Facebook', 3.882, 11.75, 5.12, 0.661, { bold: true, align: 'center', lineSpacing: 51 });
  txt(s, '\uf39e', 4.673, 7.88, 3.537, 3.537, { fontSize: 88, color: RED, fontFace: ICONFONT, align: 'center', valign: 'middle', lineSpacing: null, wrap: false, fill: { color: ROSE } });
  txt(s, '\uf16d', 11.564, 7.88, 3.537, 3.537, { fontSize: 88, color: RED, fontFace: ICONFONT, align: 'center', valign: 'middle', lineSpacing: null, wrap: false, fill: { color: ROSE } });
  txt(s, '\uf099', 18.454, 7.88, 3.537, 3.537, { fontSize: 88, color: RED, fontFace: ICONFONT, align: 'center', valign: 'middle', lineSpacing: null, wrap: false, fill: { color: ROSE } });
  pageNum(s, 50);
  logo(s);
  head(s, 'follow us', 3.49, 1.682, 19.686, 1.487);
  txt(s, 'Lorem ipsum dolor sit amet, cons tetuer adipiscing elit, sed diam nonummy nibh. Id regione deserunt quaestio pro, ad affert fuisset mnesarchum vix. Sed ei utamur indoctum. Ne mei, eum percipit ponderum assueverit et. Vim wisi intellegat at. Id regione pro, ad affert fuisset percipit.', 3.49, 3.562, 19.686, 3.029);
}

// ------------------------------------------------------------ -- "Contacts"
function slide51(s) {
  decorate(s, ['tlRose', 'sashSky', 'brBig']);
  pageNum(s, 51);
  logo(s);
  txt(s, '123 Street, City, State 45678\n\n1234 56 7890\n\ncompanyname@com\n\nwww.companyname.com', 3.49, 5.074, 9.885, 5.301);
  head(s, 'Contact', 3.49, 3.193, 9.885, 1.487);
  txt(s, '\uf0e0', 2.707, 5.2, 0.449, 0.485, { fontSize: 32, color: RED, fontFace: ICONFONT, align: 'center', lineSpacing: null, wrap: false });
  txt(s, '\uf099', 2.707, 6.707, 0.449, 0.485, { fontSize: 32, color: RED, fontFace: ICONFONT, align: 'center', lineSpacing: null, wrap: false });
  txt(s, '\uf095', 2.707, 9.719, 0.449, 0.485, { fontSize: 32, color: RED, fontFace: ICONFONT, align: 'center', lineSpacing: null, wrap: false });
  txt(s, '\uf015', 2.679, 8.213, 0.505, 0.485, { fontSize: 32, color: RED, fontFace: ICONFONT, align: 'center', lineSpacing: null, wrap: false });
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36,
  slide37, slide38, slide39, slide40, slide41, slide42,
  slide43, slide44, slide45, slide46, slide47, slide48,
  slide49, slide50, slide51,
];

// --------------------------------------------------------------- assemble
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'FASHION', width: W, height: H });
  pptx.layout = 'FASHION';
  pptx.title = 'Fashion Shop';
  pptx.defineSlideMaster({ title: 'MAIN', background: { color: BG } });

  SLIDES.forEach(fn => fn(pptx.addSlide({ masterName: 'MAIN' })));

  return pptx.writeFile({ fileName: path.join(__dirname, '1398797e-5890-4105-a46e-9f4c680ecf51_grok_final.pptx') });
}

build().then(f => console.log('wrote', f));
