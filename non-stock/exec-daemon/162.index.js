"use strict";
exports.id = 162;
exports.ids = [162];
exports.modules = {

/***/ "../../node_modules/.pnpm/@jimp+plugin-resize@1.6.0/node_modules/@jimp/plugin-resize/dist/esm/index.js"
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

// ESM COMPAT FLAG
__webpack_require__.r(__webpack_exports__);

// EXPORTS
__webpack_require__.d(__webpack_exports__, {
  ResizeStrategy: () => (/* reexport */ ResizeStrategy),
  methods: () => (/* binding */ methods)
});

;// ../../node_modules/.pnpm/@jimp+plugin-resize@1.6.0/node_modules/@jimp/plugin-resize/dist/esm/constants.js
/**
 * What resizing algorithm to use.
 */
var ResizeStrategy;
(function (ResizeStrategy) {
    /**
     * Nearest Neighbor resizing is a method used in image processing that assigns the value of the nearest pixel to the output pixel when resizing an image.
     * While fast, it can lead to lower quality outputs with noticeable pixelation, especially at larger scaling factors.
     */
    ResizeStrategy["NEAREST_NEIGHBOR"] = "nearestNeighbor";
    /**
     * Bilinear resizing is an image scaling method that uses the weighted average of the four nearest pixel values, providing smoother gradients than nearest neighbor resizing.
     * It's computationally more intense than nearest neighbor but results in images of higher quality and fewer artifacts.
     */
    ResizeStrategy["BILINEAR"] = "bilinearInterpolation";
    /**
     * Bicubic resizing is an image interpolation method that uses the values of the nearest 16 pixels in the input image to calculate the output pixel value, providing even more smoothness and sharpness than bilinear resizing.
     * Although it's computationally more expensive than bilinear and nearest neighbor, it produces higher quality images, making it ideal for photographic image scaling.
     */
    ResizeStrategy["BICUBIC"] = "bicubicInterpolation";
    /**
     * Hermite resizing is an image resizing method that uses Hermite interpolation, a mathematical formula, to determine the values of output pixels based on a weighted average of the surrounding pixels.
     * Although slower than some simpler techniques like bilinear or nearest neighbor, Hermite resizing can produce higher quality visuals and detailed renderings with less blurring.
     */
    ResizeStrategy["HERMITE"] = "hermiteInterpolation";
    ResizeStrategy["BEZIER"] = "bezierInterpolation";
})(ResizeStrategy || (ResizeStrategy = {}));
//# sourceMappingURL=constants.js.map
// EXTERNAL MODULE: ../../node_modules/.pnpm/zod@3.25.76/node_modules/zod/v3/types.js + 4 modules
var types = __webpack_require__("../../node_modules/.pnpm/zod@3.25.76/node_modules/zod/v3/types.js");
;// ../../node_modules/.pnpm/@jimp+plugin-resize@1.6.0/node_modules/@jimp/plugin-resize/dist/esm/modules/resize.js
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-nocheck
// JavaScript Image Resizer (c) 2012 - Grant Galitz
// Released to public domain 29 July 2013: https://github.com/grantgalitz/JS-Image-Resizer/issues/4
function Resize(widthOriginal, heightOriginal, targetWidth, targetHeight, blendAlpha, interpolationPass, resizeCallback) {
    this.widthOriginal = Math.abs(Math.floor(widthOriginal) || 0);
    this.heightOriginal = Math.abs(Math.floor(heightOriginal) || 0);
    this.targetWidth = Math.abs(Math.floor(targetWidth) || 0);
    this.targetHeight = Math.abs(Math.floor(targetHeight) || 0);
    this.colorChannels = blendAlpha ? 4 : 3;
    this.interpolationPass = Boolean(interpolationPass);
    this.resizeCallback =
        typeof resizeCallback === "function" ? resizeCallback : function () { };
    this.targetWidthMultipliedByChannels = this.targetWidth * this.colorChannels;
    this.originalWidthMultipliedByChannels =
        this.widthOriginal * this.colorChannels;
    this.originalHeightMultipliedByChannels =
        this.heightOriginal * this.colorChannels;
    this.widthPassResultSize =
        this.targetWidthMultipliedByChannels * this.heightOriginal;
    this.finalResultSize =
        this.targetWidthMultipliedByChannels * this.targetHeight;
    this.initialize();
}
Resize.prototype.initialize = function () {
    // Perform some checks:
    if (this.widthOriginal > 0 &&
        this.heightOriginal > 0 &&
        this.targetWidth > 0 &&
        this.targetHeight > 0) {
        this.configurePasses();
    }
    else {
        console.log(this);
        throw new Error("Invalid settings specified for the resizer.");
    }
};
Resize.prototype.configurePasses = function () {
    if (this.widthOriginal === this.targetWidth) {
        // Bypass the width resizer pass:
        this.resizeWidth = this.bypassResizer;
    }
    else {
        // Setup the width resizer pass:
        this.ratioWeightWidthPass = this.widthOriginal / this.targetWidth;
        if (this.ratioWeightWidthPass < 1 && this.interpolationPass) {
            this.initializeFirstPassBuffers(true);
            this.resizeWidth =
                this.colorChannels === 4
                    ? this.resizeWidthInterpolatedRGBA
                    : this.resizeWidthInterpolatedRGB;
        }
        else {
            this.initializeFirstPassBuffers(false);
            this.resizeWidth =
                this.colorChannels === 4 ? this.resizeWidthRGBA : this.resizeWidthRGB;
        }
    }
    if (this.heightOriginal === this.targetHeight) {
        // Bypass the height resizer pass:
        this.resizeHeight = this.bypassResizer;
    }
    else {
        // Setup the height resizer pass:
        this.ratioWeightHeightPass = this.heightOriginal / this.targetHeight;
        if (this.ratioWeightHeightPass < 1 && this.interpolationPass) {
            this.initializeSecondPassBuffers(true);
            this.resizeHeight = this.resizeHeightInterpolated;
        }
        else {
            this.initializeSecondPassBuffers(false);
            this.resizeHeight =
                this.colorChannels === 4 ? this.resizeHeightRGBA : this.resizeHeightRGB;
        }
    }
};
Resize.prototype._resizeWidthInterpolatedRGBChannels = function (buffer, fourthChannel) {
    const channelsNum = fourthChannel ? 4 : 3;
    const ratioWeight = this.ratioWeightWidthPass;
    const outputBuffer = this.widthBuffer;
    let weight = 0;
    let finalOffset = 0;
    let pixelOffset = 0;
    let firstWeight = 0;
    let secondWeight = 0;
    let targetPosition;
    // Handle for only one interpolation input being valid for start calculation:
    for (targetPosition = 0; weight < 1 / 3; targetPosition += channelsNum, weight += ratioWeight) {
        for (finalOffset = targetPosition, pixelOffset = 0; finalOffset < this.widthPassResultSize; pixelOffset += this.originalWidthMultipliedByChannels,
            finalOffset += this.targetWidthMultipliedByChannels) {
            outputBuffer[finalOffset] = buffer[pixelOffset];
            outputBuffer[finalOffset + 1] = buffer[pixelOffset + 1];
            outputBuffer[finalOffset + 2] = buffer[pixelOffset + 2];
            if (fourthChannel)
                outputBuffer[finalOffset + 3] = buffer[pixelOffset + 3];
        }
    }
    // Adjust for overshoot of the last pass's counter:
    weight -= 1 / 3;
    let interpolationWidthSourceReadStop;
    for (interpolationWidthSourceReadStop = this.widthOriginal - 1; weight < interpolationWidthSourceReadStop; targetPosition += channelsNum, weight += ratioWeight) {
        // Calculate weightings:
        secondWeight = weight % 1;
        firstWeight = 1 - secondWeight;
        // Interpolate:
        for (finalOffset = targetPosition,
            pixelOffset = Math.floor(weight) * channelsNum; finalOffset < this.widthPassResultSize; pixelOffset += this.originalWidthMultipliedByChannels,
            finalOffset += this.targetWidthMultipliedByChannels) {
            outputBuffer[finalOffset + 0] =
                buffer[pixelOffset + 0] * firstWeight +
                    buffer[pixelOffset + channelsNum + 0] * secondWeight;
            outputBuffer[finalOffset + 1] =
                buffer[pixelOffset + 1] * firstWeight +
                    buffer[pixelOffset + channelsNum + 1] * secondWeight;
            outputBuffer[finalOffset + 2] =
                buffer[pixelOffset + 2] * firstWeight +
                    buffer[pixelOffset + channelsNum + 2] * secondWeight;
            if (fourthChannel)
                outputBuffer[finalOffset + 3] =
                    buffer[pixelOffset + 3] * firstWeight +
                        buffer[pixelOffset + channelsNum + 3] * secondWeight;
        }
    }
    // Handle for only one interpolation input being valid for end calculation:
    for (interpolationWidthSourceReadStop =
        this.originalWidthMultipliedByChannels - channelsNum; targetPosition < this.targetWidthMultipliedByChannels; targetPosition += channelsNum) {
        for (finalOffset = targetPosition,
            pixelOffset = interpolationWidthSourceReadStop; finalOffset < this.widthPassResultSize; pixelOffset += this.originalWidthMultipliedByChannels,
            finalOffset += this.targetWidthMultipliedByChannels) {
            outputBuffer[finalOffset] = buffer[pixelOffset];
            outputBuffer[finalOffset + 1] = buffer[pixelOffset + 1];
            outputBuffer[finalOffset + 2] = buffer[pixelOffset + 2];
            if (fourthChannel)
                outputBuffer[finalOffset + 3] = buffer[pixelOffset + 3];
        }
    }
    return outputBuffer;
};
Resize.prototype._resizeWidthRGBChannels = function (buffer, fourthChannel) {
    const channelsNum = fourthChannel ? 4 : 3;
    const ratioWeight = this.ratioWeightWidthPass;
    const ratioWeightDivisor = 1 / ratioWeight;
    const nextLineOffsetOriginalWidth = this.originalWidthMultipliedByChannels - channelsNum + 1;
    const nextLineOffsetTargetWidth = this.targetWidthMultipliedByChannels - channelsNum + 1;
    const output = this.outputWidthWorkBench;
    const outputBuffer = this.widthBuffer;
    const trustworthyColorsCount = this.outputWidthWorkBenchOpaquePixelsCount;
    let weight = 0;
    let amountToNext = 0;
    let actualPosition = 0;
    let currentPosition = 0;
    let line = 0;
    let pixelOffset = 0;
    let outputOffset = 0;
    let multiplier = 1;
    let r = 0;
    let g = 0;
    let b = 0;
    let a = 0;
    do {
        for (line = 0; line < this.originalHeightMultipliedByChannels;) {
            output[line++] = 0;
            output[line++] = 0;
            output[line++] = 0;
            if (fourthChannel) {
                output[line++] = 0;
                trustworthyColorsCount[line / channelsNum - 1] = 0;
            }
        }
        weight = ratioWeight;
        do {
            amountToNext = 1 + actualPosition - currentPosition;
            multiplier = Math.min(weight, amountToNext);
            for (line = 0, pixelOffset = actualPosition; line < this.originalHeightMultipliedByChannels; pixelOffset += nextLineOffsetOriginalWidth) {
                r = buffer[pixelOffset];
                g = buffer[++pixelOffset];
                b = buffer[++pixelOffset];
                a = fourthChannel ? buffer[++pixelOffset] : 255;
                // Ignore RGB values if pixel is completely transparent
                output[line++] += (a ? r : 0) * multiplier;
                output[line++] += (a ? g : 0) * multiplier;
                output[line++] += (a ? b : 0) * multiplier;
                if (fourthChannel) {
                    output[line++] += a * multiplier;
                    trustworthyColorsCount[line / channelsNum - 1] += a ? multiplier : 0;
                }
            }
            if (weight >= amountToNext) {
                actualPosition += channelsNum;
                currentPosition = actualPosition;
                weight -= amountToNext;
            }
            else {
                currentPosition += weight;
                break;
            }
        } while (weight > 0 &&
            actualPosition < this.originalWidthMultipliedByChannels);
        for (line = 0, pixelOffset = outputOffset; line < this.originalHeightMultipliedByChannels; pixelOffset += nextLineOffsetTargetWidth) {
            weight = fourthChannel ? trustworthyColorsCount[line / channelsNum] : 1;
            multiplier = fourthChannel
                ? weight
                    ? 1 / weight
                    : 0
                : ratioWeightDivisor;
            outputBuffer[pixelOffset] = output[line++] * multiplier;
            outputBuffer[++pixelOffset] = output[line++] * multiplier;
            outputBuffer[++pixelOffset] = output[line++] * multiplier;
            if (fourthChannel)
                outputBuffer[++pixelOffset] = output[line++] * ratioWeightDivisor;
        }
        outputOffset += channelsNum;
    } while (outputOffset < this.targetWidthMultipliedByChannels);
    return outputBuffer;
};
Resize.prototype._resizeHeightRGBChannels = function (buffer, fourthChannel) {
    const ratioWeight = this.ratioWeightHeightPass;
    const ratioWeightDivisor = 1 / ratioWeight;
    const output = this.outputHeightWorkBench;
    const outputBuffer = this.heightBuffer;
    const trustworthyColorsCount = this.outputHeightWorkBenchOpaquePixelsCount;
    let weight = 0;
    let amountToNext = 0;
    let actualPosition = 0;
    let currentPosition = 0;
    let pixelOffset = 0;
    let outputOffset = 0;
    let caret = 0;
    let multiplier = 1;
    let r = 0;
    let g = 0;
    let b = 0;
    let a = 0;
    do {
        for (pixelOffset = 0; pixelOffset < this.targetWidthMultipliedByChannels;) {
            output[pixelOffset++] = 0;
            output[pixelOffset++] = 0;
            output[pixelOffset++] = 0;
            if (fourthChannel) {
                output[pixelOffset++] = 0;
                trustworthyColorsCount[pixelOffset / 4 - 1] = 0;
            }
        }
        weight = ratioWeight;
        do {
            amountToNext = 1 + actualPosition - currentPosition;
            multiplier = Math.min(weight, amountToNext);
            caret = actualPosition;
            for (pixelOffset = 0; pixelOffset < this.targetWidthMultipliedByChannels;) {
                r = buffer[caret++];
                g = buffer[caret++];
                b = buffer[caret++];
                a = fourthChannel ? buffer[caret++] : 255;
                // Ignore RGB values if pixel is completely transparent
                output[pixelOffset++] += (a ? r : 0) * multiplier;
                output[pixelOffset++] += (a ? g : 0) * multiplier;
                output[pixelOffset++] += (a ? b : 0) * multiplier;
                if (fourthChannel) {
                    output[pixelOffset++] += a * multiplier;
                    trustworthyColorsCount[pixelOffset / 4 - 1] += a ? multiplier : 0;
                }
            }
            if (weight >= amountToNext) {
                actualPosition = caret;
                currentPosition = actualPosition;
                weight -= amountToNext;
            }
            else {
                currentPosition += weight;
                break;
            }
        } while (weight > 0 && actualPosition < this.widthPassResultSize);
        for (pixelOffset = 0; pixelOffset < this.targetWidthMultipliedByChannels;) {
            weight = fourthChannel ? trustworthyColorsCount[pixelOffset / 4] : 1;
            multiplier = fourthChannel
                ? weight
                    ? 1 / weight
                    : 0
                : ratioWeightDivisor;
            outputBuffer[outputOffset++] = Math.round(output[pixelOffset++] * multiplier);
            outputBuffer[outputOffset++] = Math.round(output[pixelOffset++] * multiplier);
            outputBuffer[outputOffset++] = Math.round(output[pixelOffset++] * multiplier);
            if (fourthChannel) {
                outputBuffer[outputOffset++] = Math.round(output[pixelOffset++] * ratioWeightDivisor);
            }
        }
    } while (outputOffset < this.finalResultSize);
    return outputBuffer;
};
Resize.prototype.resizeWidthInterpolatedRGB = function (buffer) {
    return this._resizeWidthInterpolatedRGBChannels(buffer, false);
};
Resize.prototype.resizeWidthInterpolatedRGBA = function (buffer) {
    return this._resizeWidthInterpolatedRGBChannels(buffer, true);
};
Resize.prototype.resizeWidthRGB = function (buffer) {
    return this._resizeWidthRGBChannels(buffer, false);
};
Resize.prototype.resizeWidthRGBA = function (buffer) {
    return this._resizeWidthRGBChannels(buffer, true);
};
Resize.prototype.resizeHeightInterpolated = function (buffer) {
    const ratioWeight = this.ratioWeightHeightPass;
    const outputBuffer = this.heightBuffer;
    let weight = 0;
    let finalOffset = 0;
    let pixelOffset = 0;
    let pixelOffsetAccumulated = 0;
    let pixelOffsetAccumulated2 = 0;
    let firstWeight = 0;
    let secondWeight = 0;
    let interpolationHeightSourceReadStop;
    // Handle for only one interpolation input being valid for start calculation:
    for (; weight < 1 / 3; weight += ratioWeight) {
        for (pixelOffset = 0; pixelOffset < this.targetWidthMultipliedByChannels;) {
            outputBuffer[finalOffset++] = Math.round(buffer[pixelOffset++]);
        }
    }
    // Adjust for overshoot of the last pass's counter:
    weight -= 1 / 3;
    for (interpolationHeightSourceReadStop = this.heightOriginal - 1; weight < interpolationHeightSourceReadStop; weight += ratioWeight) {
        // Calculate weightings:
        secondWeight = weight % 1;
        firstWeight = 1 - secondWeight;
        // Interpolate:
        pixelOffsetAccumulated =
            Math.floor(weight) * this.targetWidthMultipliedByChannels;
        pixelOffsetAccumulated2 =
            pixelOffsetAccumulated + this.targetWidthMultipliedByChannels;
        for (pixelOffset = 0; pixelOffset < this.targetWidthMultipliedByChannels; ++pixelOffset) {
            outputBuffer[finalOffset++] = Math.round(buffer[pixelOffsetAccumulated++] * firstWeight +
                buffer[pixelOffsetAccumulated2++] * secondWeight);
        }
    }
    // Handle for only one interpolation input being valid for end calculation:
    while (finalOffset < this.finalResultSize) {
        for (pixelOffset = 0,
            pixelOffsetAccumulated =
                interpolationHeightSourceReadStop *
                    this.targetWidthMultipliedByChannels; pixelOffset < this.targetWidthMultipliedByChannels; ++pixelOffset) {
            outputBuffer[finalOffset++] = Math.round(buffer[pixelOffsetAccumulated++]);
        }
    }
    return outputBuffer;
};
Resize.prototype.resizeHeightRGB = function (buffer) {
    return this._resizeHeightRGBChannels(buffer, false);
};
Resize.prototype.resizeHeightRGBA = function (buffer) {
    return this._resizeHeightRGBChannels(buffer, true);
};
Resize.prototype.resize = function (buffer) {
    this.resizeCallback(this.resizeHeight(this.resizeWidth(buffer)));
};
Resize.prototype.bypassResizer = function (buffer) {
    // Just return the buffer passed:
    return buffer;
};
Resize.prototype.initializeFirstPassBuffers = function (BILINEARAlgo) {
    // Initialize the internal width pass buffers:
    this.widthBuffer = this.generateFloatBuffer(this.widthPassResultSize);
    if (!BILINEARAlgo) {
        this.outputWidthWorkBench = this.generateFloatBuffer(this.originalHeightMultipliedByChannels);
        if (this.colorChannels > 3) {
            this.outputWidthWorkBenchOpaquePixelsCount = this.generateFloat64Buffer(this.heightOriginal);
        }
    }
};
Resize.prototype.initializeSecondPassBuffers = function (BILINEARAlgo) {
    // Initialize the internal height pass buffers:
    this.heightBuffer = this.generateUint8Buffer(this.finalResultSize);
    if (!BILINEARAlgo) {
        this.outputHeightWorkBench = this.generateFloatBuffer(this.targetWidthMultipliedByChannels);
        if (this.colorChannels > 3) {
            this.outputHeightWorkBenchOpaquePixelsCount = this.generateFloat64Buffer(this.targetWidth);
        }
    }
};
Resize.prototype.generateFloatBuffer = function (bufferLength) {
    // Generate a float32 typed array buffer:
    try {
        return new Float32Array(bufferLength);
    }
    catch (error) {
        console.error(error);
        return [];
    }
};
Resize.prototype.generateFloat64Buffer = function (bufferLength) {
    // Generate a float64 typed array buffer:
    try {
        return new Float64Array(bufferLength);
    }
    catch (error) {
        console.error(error);
        return [];
    }
};
Resize.prototype.generateUint8Buffer = function (bufferLength) {
    // Generate a uint8 typed array buffer:
    try {
        return new Uint8Array(bufferLength);
    }
    catch (error) {
        console.error(error);
        return [];
    }
};
/* harmony default export */ const modules_resize = (Resize);
//# sourceMappingURL=resize.js.map
;// ../../node_modules/.pnpm/@jimp+plugin-resize@1.6.0/node_modules/@jimp/plugin-resize/dist/esm/modules/resize2.js
/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Copyright (c) 2015 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:</p>
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
const operations = {
    nearestNeighbor(src, dst) {
        const wSrc = src.width;
        const hSrc = src.height;
        const wDst = dst.width;
        const hDst = dst.height;
        const bufSrc = src.data;
        const bufDst = dst.data;
        for (let i = 0; i < hDst; i++) {
            for (let j = 0; j < wDst; j++) {
                let posDst = (i * wDst + j) * 4;
                const iSrc = Math.floor((i * hSrc) / hDst);
                const jSrc = Math.floor((j * wSrc) / wDst);
                let posSrc = (iSrc * wSrc + jSrc) * 4;
                bufDst[posDst++] = bufSrc[posSrc++];
                bufDst[posDst++] = bufSrc[posSrc++];
                bufDst[posDst++] = bufSrc[posSrc++];
                bufDst[posDst++] = bufSrc[posSrc++];
            }
        }
    },
    bilinearInterpolation(src, dst) {
        const wSrc = src.width;
        const hSrc = src.height;
        const wDst = dst.width;
        const hDst = dst.height;
        const bufSrc = src.data;
        const bufDst = dst.data;
        const interpolate = function (k, kMin, vMin, kMax, vMax) {
            // special case - k is integer
            if (kMin === kMax) {
                return vMin;
            }
            return Math.round((k - kMin) * vMax + (kMax - k) * vMin);
        };
        const assign = function (pos, offset, x, xMin, xMax, y, yMin, yMax) {
            let posMin = (yMin * wSrc + xMin) * 4 + offset;
            let posMax = (yMin * wSrc + xMax) * 4 + offset;
            const vMin = interpolate(x, xMin, bufSrc[posMin], xMax, bufSrc[posMax]);
            // special case, y is integer
            if (yMax === yMin) {
                bufDst[pos + offset] = vMin;
            }
            else {
                posMin = (yMax * wSrc + xMin) * 4 + offset;
                posMax = (yMax * wSrc + xMax) * 4 + offset;
                const vMax = interpolate(x, xMin, bufSrc[posMin], xMax, bufSrc[posMax]);
                bufDst[pos + offset] = interpolate(y, yMin, vMin, yMax, vMax);
            }
        };
        for (let i = 0; i < hDst; i++) {
            for (let j = 0; j < wDst; j++) {
                const posDst = (i * wDst + j) * 4;
                // x & y in src coordinates
                const x = (j * wSrc) / wDst;
                const xMin = Math.floor(x);
                const xMax = Math.min(Math.ceil(x), wSrc - 1);
                const y = (i * hSrc) / hDst;
                const yMin = Math.floor(y);
                const yMax = Math.min(Math.ceil(y), hSrc - 1);
                assign(posDst, 0, x, xMin, xMax, y, yMin, yMax);
                assign(posDst, 1, x, xMin, xMax, y, yMin, yMax);
                assign(posDst, 2, x, xMin, xMax, y, yMin, yMax);
                assign(posDst, 3, x, xMin, xMax, y, yMin, yMax);
            }
        }
    },
    _interpolate2D(src, dst, options, interpolate) {
        const bufSrc = src.data;
        const bufDst = dst.data;
        const wSrc = src.width;
        const hSrc = src.height;
        const wDst = dst.width;
        const hDst = dst.height;
        // when dst smaller than src/2, interpolate first to a multiple between 0.5 and 1.0 src, then sum squares
        const wM = Math.max(1, Math.floor(wSrc / wDst));
        const wDst2 = wDst * wM;
        const hM = Math.max(1, Math.floor(hSrc / hDst));
        const hDst2 = hDst * hM;
        // ===========================================================
        // Pass 1 - interpolate rows
        // buf1 has width of dst2 and height of src
        const buf1 = Buffer.alloc(wDst2 * hSrc * 4);
        for (let i = 0; i < hSrc; i++) {
            for (let j = 0; j < wDst2; j++) {
                // i in src coords, j in dst coords
                // calculate x in src coords
                // this interpolation requires 4 sample points and the two inner ones must be real
                // the outer points can be fudged for the edges.
                // therefore (wSrc-1)/wDst2
                const x = (j * (wSrc - 1)) / wDst2;
                const xPos = Math.floor(x);
                const t = x - xPos;
                const srcPos = (i * wSrc + xPos) * 4;
                const buf1Pos = (i * wDst2 + j) * 4;
                for (let k = 0; k < 4; k++) {
                    const kPos = srcPos + k;
                    const x0 = xPos > 0 ? bufSrc[kPos - 4] : 2 * bufSrc[kPos] - bufSrc[kPos + 4];
                    const x1 = bufSrc[kPos];
                    const x2 = bufSrc[kPos + 4];
                    const x3 = xPos < wSrc - 2
                        ? bufSrc[kPos + 8]
                        : 2 * bufSrc[kPos + 4] - bufSrc[kPos];
                    buf1[buf1Pos + k] = interpolate(x0, x1, x2, x3, t);
                }
            }
        }
        // this._writeFile(wDst2, hSrc, buf1, "out/buf1.jpg");
        // ===========================================================
        // Pass 2 - interpolate columns
        // buf2 has width and height of dst2
        const buf2 = Buffer.alloc(wDst2 * hDst2 * 4);
        for (let i = 0; i < hDst2; i++) {
            for (let j = 0; j < wDst2; j++) {
                // i&j in dst2 coords
                // calculate y in buf1 coords
                // this interpolation requires 4 sample points and the two inner ones must be real
                // the outer points can be fudged for the edges.
                // therefore (hSrc-1)/hDst2
                const y = (i * (hSrc - 1)) / hDst2;
                const yPos = Math.floor(y);
                const t = y - yPos;
                const buf1Pos = (yPos * wDst2 + j) * 4;
                const buf2Pos = (i * wDst2 + j) * 4;
                for (let k = 0; k < 4; k++) {
                    const kPos = buf1Pos + k;
                    const y0 = yPos > 0
                        ? buf1[kPos - wDst2 * 4]
                        : 2 * buf1[kPos] - buf1[kPos + wDst2 * 4];
                    const y1 = buf1[kPos];
                    const y2 = buf1[kPos + wDst2 * 4];
                    const y3 = yPos < hSrc - 2
                        ? buf1[kPos + wDst2 * 8]
                        : 2 * buf1[kPos + wDst2 * 4] - buf1[kPos];
                    buf2[buf2Pos + k] = interpolate(y0, y1, y2, y3, t);
                }
            }
        }
        // this._writeFile(wDst2, hDst2, buf2, "out/buf2.jpg");
        // ===========================================================
        // Pass 3 - scale to dst
        const m = wM * hM;
        if (m > 1) {
            for (let i = 0; i < hDst; i++) {
                for (let j = 0; j < wDst; j++) {
                    // i&j in dst bounded coords
                    let r = 0;
                    let g = 0;
                    let b = 0;
                    let a = 0;
                    let realColors = 0;
                    for (let y = 0; y < hM; y++) {
                        const yPos = i * hM + y;
                        for (let x = 0; x < wM; x++) {
                            const xPos = j * wM + x;
                            const xyPos = (yPos * wDst2 + xPos) * 4;
                            const pixelAlpha = buf2[xyPos + 3];
                            if (pixelAlpha) {
                                r += buf2[xyPos];
                                g += buf2[xyPos + 1];
                                b += buf2[xyPos + 2];
                                realColors++;
                            }
                            a += pixelAlpha;
                        }
                    }
                    const pos = (i * wDst + j) * 4;
                    bufDst[pos] = realColors ? Math.round(r / realColors) : 0;
                    bufDst[pos + 1] = realColors ? Math.round(g / realColors) : 0;
                    bufDst[pos + 2] = realColors ? Math.round(b / realColors) : 0;
                    bufDst[pos + 3] = Math.round(a / m);
                }
            }
        }
        else {
            // replace dst buffer with buf2
            dst.data = buf2;
        }
    },
    bicubicInterpolation(src, dst, options) {
        const interpolateCubic = function (x0, x1, x2, x3, t) {
            const a0 = x3 - x2 - x0 + x1;
            const a1 = x0 - x1 - a0;
            const a2 = x2 - x0;
            const a3 = x1;
            return Math.max(0, Math.min(255, a0 * (t * t * t) + a1 * (t * t) + a2 * t + a3));
        };
        return this._interpolate2D(src, dst, options, interpolateCubic);
    },
    hermiteInterpolation(src, dst, options) {
        const interpolateHermite = function (x0, x1, x2, x3, t) {
            const c0 = x1;
            const c1 = 0.5 * (x2 - x0);
            const c2 = x0 - 2.5 * x1 + 2 * x2 - 0.5 * x3;
            const c3 = 0.5 * (x3 - x0) + 1.5 * (x1 - x2);
            return Math.max(0, Math.min(255, Math.round(((c3 * t + c2) * t + c1) * t + c0)));
        };
        return this._interpolate2D(src, dst, options, interpolateHermite);
    },
    bezierInterpolation(src, dst, options) {
        // between 2 points y(n), y(n+1), use next points out, y(n-1), y(n+2)
        // to predict control points (a & b) to be placed at n+0.5
        //  ya(n) = y(n) + (y(n+1)-y(n-1))/4
        //  yb(n) = y(n+1) - (y(n+2)-y(n))/4
        // then use std bezier to interpolate [n,n+1)
        //  y(n+t) = y(n)*(1-t)^3 + 3 * ya(n)*(1-t)^2*t + 3 * yb(n)*(1-t)*t^2 + y(n+1)*t^3
        //  note the 3* factor for the two control points
        // for edge cases, can choose:
        //  y(-1) = y(0) - 2*(y(1)-y(0))
        //  y(w) = y(w-1) + 2*(y(w-1)-y(w-2))
        // but can go with y(-1) = y(0) and y(w) = y(w-1)
        const interpolateBezier = function (x0, x1, x2, x3, t) {
            // x1, x2 are the knots, use x0 and x3 to calculate control points
            const cp1 = x1 + (x2 - x0) / 4;
            const cp2 = x2 - (x3 - x1) / 4;
            const nt = 1 - t;
            const c0 = x1 * nt * nt * nt;
            const c1 = 3 * cp1 * nt * nt * t;
            const c2 = 3 * cp2 * nt * t * t;
            const c3 = x2 * t * t * t;
            return Math.max(0, Math.min(255, Math.round(c0 + c1 + c2 + c3)));
        };
        return this._interpolate2D(src, dst, options, interpolateBezier);
    },
};
//# sourceMappingURL=resize2.js.map
;// ../../node_modules/.pnpm/@jimp+plugin-resize@1.6.0/node_modules/@jimp/plugin-resize/dist/esm/index.js





const ResizeOptionsSchema = types/* union */.KC([
    types/* object */.Ik({
        /** the width to resize the image to */
        w: types/* number */.ai().min(0),
        /** the height to resize the image to */
        h: types/* number */.ai().min(0).optional(),
        /** a scaling method (e.g. ResizeStrategy.BEZIER) */
        mode: types/* nativeEnum */.fc(ResizeStrategy).optional(),
    }),
    types/* object */.Ik({
        /** the width to resize the image to */
        w: types/* number */.ai().min(0).optional(),
        /** the height to resize the image to */
        h: types/* number */.ai().min(0),
        /** a scaling method (e.g. ResizeStrategy.BEZIER) */
        mode: types/* nativeEnum */.fc(ResizeStrategy).optional(),
    }),
]);
const ScaleToFitOptionsSchema = types/* object */.Ik({
    /** the width to resize the image to */
    w: types/* number */.ai().min(0),
    /** the height to resize the image to */
    h: types/* number */.ai().min(0),
    /** a scaling method (e.g. Jimp.RESIZE_BEZIER) */
    mode: types/* nativeEnum */.fc(ResizeStrategy).optional(),
});
const ScaleComplexOptionsSchema = types/* object */.Ik({
    /** the width to resize the image to */
    f: types/* number */.ai().min(0),
    /** a scaling method (e.g. Jimp.RESIZE_BEZIER) */
    mode: types/* nativeEnum */.fc(ResizeStrategy).optional(),
});
const methods = {
    /**
     * Resizes the image to a set width and height using a 2-pass bilinear algorithm
     * @example
     * ```ts
     * import { Jimp } from "jimp";
     *
     * const image = await Jimp.read("test/image.png");
     *
     * image.resize({ w: 150 });
     * ```
     */
    resize(image, options) {
        const { mode } = ResizeOptionsSchema.parse(options);
        let w;
        let h;
        if (typeof options.w === "number") {
            w = options.w;
            h = options.h ?? image.bitmap.height * (w / image.bitmap.width);
        }
        else if (typeof options.h === "number") {
            h = options.h;
            w = options.w ?? image.bitmap.width * (h / image.bitmap.height);
        }
        else {
            throw new Error("w must be a number");
        }
        // round inputs
        w = Math.round(w) || 1;
        h = Math.round(h) || 1;
        if (mode && typeof operations[mode] === "function") {
            const dst = {
                data: Buffer.alloc(w * h * 4),
                width: w,
                height: h,
            };
            operations[mode](image.bitmap, dst);
            image.bitmap = dst;
        }
        else {
            const resize = new modules_resize(image.bitmap.width, image.bitmap.height, w, h, true, true, (buffer) => {
                image.bitmap.data = Buffer.from(buffer);
                image.bitmap.width = w;
                image.bitmap.height = h;
            });
            resize.resize(image.bitmap.data);
        }
        return image;
    },
    /**
     * Uniformly scales the image by a factor.
     * @param f the factor to scale the image by
     * @param mode (optional) a scaling method (e.g. Jimp.RESIZE_BEZIER)
     * @example
     * ```ts
     * import { Jimp } from "jimp";
     *
     * const image = await Jimp.read("test/image.png");
     *
     * image.scale(0.5);
     * ```
     */
    scale(image, options) {
        const { f, mode } = typeof options === "number"
            ? { f: options }
            : ScaleComplexOptionsSchema.parse(options);
        const w = image.bitmap.width * f;
        const h = image.bitmap.height * f;
        return this.resize(image, { w, h, mode: mode });
    },
    /**
     * Scale the image to the largest size that fits inside the rectangle that has the given width and height.
     * @param w the width to resize the image to
     * @param h the height to resize the image to
     * @param mode a scaling method (e.g. ResizeStrategy.BEZIER)
     * @example
     * ```ts
     * import { Jimp } from "jimp";
     *
     * const image = await Jimp.read("test/image.png");
     *
     * image.scaleToFit(100, 100);
     * ```
     */
    scaleToFit(image, options) {
        const { h, w, mode } = ScaleToFitOptionsSchema.parse(options);
        const f = w / h > image.bitmap.width / image.bitmap.height
            ? h / image.bitmap.height
            : w / image.bitmap.width;
        return this.scale(image, { f, mode: mode });
    },
};
//# sourceMappingURL=index.js.map

/***/ }

};
;