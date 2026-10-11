exports.id = 295;
exports.ids = [295];
exports.modules = {

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/bitmapimage.js"
(module) {

"use strict";


/** @class BitmapImage */

class BitmapImage {

    /**
     * BitmapImage is a class that hold an RGBA (red, green, blue, alpha) representation of an image. It's shape is borrowed from the Jimp package to make it easy to transfer GIF image frames into Jimp and Jimp images into GIF image frames. Each instance has a `bitmap` property having the following properties:
     * 
     * Property | Description
     * --- | ---
     * bitmap.width | width of image in pixels
     * bitmap.height | height of image in pixels
     * bitmap.data | a Buffer whose every four bytes represents a pixel, each sequential byte of a pixel corresponding to the red, green, blue, and alpha values of the pixel
     *
     * Its constructor supports the following signatures:
     *
     * * new BitmapImage(bitmap: { width: number, height: number, data: Buffer })
     * * new BitmapImage(bitmapImage: BitmapImage)
     * * new BitmapImage(width: number, height: number, buffer: Buffer)
     * * new BitmapImage(width: number, height: number, backgroundRGBA?: number)
     * 
     * When a `BitmapImage` is provided, the constructed `BitmapImage` is a deep clone of the provided one, so that each image's pixel data can subsequently be modified without affecting each other.
     *
     * `backgroundRGBA` is an optional parameter representing a pixel as a single number. In hex, the number is as follows: 0xRRGGBBAA, where RR is the red byte, GG the green byte, BB, the blue byte, and AA the alpha value. An AA of 0x00 is considered transparent, and all non-zero AA values are treated as opaque.
     */

    constructor(...args) {
        // don't confirm the number of args, because a subclass may have
        // additional args and pass them all to the superclass
        if (args.length === 0) {
            throw new Error("constructor requires parameters");
        }
        const firstArg = args[0];
        if (firstArg !== null && typeof firstArg === 'object') {
            if (firstArg instanceof BitmapImage) {
                // copy a provided BitmapImage
                const sourceBitmap = firstArg.bitmap;
                this.bitmap = {
                    width: sourceBitmap.width,
                    height: sourceBitmap.height,
                    data: new Buffer(sourceBitmap.width * sourceBitmap.height * 4)
                };
                sourceBitmap.data.copy(this.bitmap.data);
            }
            else if (firstArg.width && firstArg.height && firstArg.data) {
                // share a provided bitmap
                this.bitmap = firstArg;
            }
            else {
                throw new Error("unrecognized constructor parameters");
            }
        }
        else if (typeof firstArg === 'number' && typeof args[1] === 'number')
        {
            const width = firstArg;
            const height = args[1];
            const thirdArg = args[2];
            this.bitmap = { width, height };

            if (Buffer.isBuffer(thirdArg)) {
                this.bitmap.data = thirdArg;
            }
            else {
                this.bitmap.data = new Buffer(width * height * 4);
                if (typeof thirdArg === 'number') {
                    this.fillRGBA(thirdArg);
                }
            }
        }
        else {
            throw new Error("unrecognized constructor parameters");
        }
    }

    /**
     * Copy a square portion of this image into another image. 
     * 
     * @param {BitmapImage} toImage Image into which to copy the square
     * @param {number} toX x-coord in toImage of upper-left corner of receiving square
     * @param {number} toY y-coord in toImage of upper-left corner of receiving square
     * @param {number} fromX x-coord in this image of upper-left corner of source square
     * @param {number} fromY y-coord in this image of upper-left corner of source square
     * @return {BitmapImage} The present image to allow for chaining.
     */

    blit(toImage, toX, toY, fromX, fromY, fromWidth, fromHeight) {
        if (fromX + fromWidth > this.bitmap.width) {
            throw new Error("copy exceeds width of source bitmap");
        }
        if (toX + fromWidth > toImage.bitmap.width) {
            throw new Error("copy exceeds width of target bitmap");
        }
        if (fromY + fromHeight > this.bitmap.height) {
            throw new Error("copy exceeds height of source bitmap");
        }
        if (toY + fromHeight > toImage.bitmap.height) {
            throw new Erro("copy exceeds height of target bitmap");
        }
        
        const sourceBuf = this.bitmap.data;
        const targetBuf = toImage.bitmap.data;
        const sourceByteWidth = this.bitmap.width * 4;
        const targetByteWidth = toImage.bitmap.width * 4;
        const copyByteWidth = fromWidth * 4;
        let si = fromY * sourceByteWidth + fromX * 4;
        let ti = toY * targetByteWidth + toX * 4;

        while (--fromHeight >= 0) {
            sourceBuf.copy(targetBuf, ti, si, si + copyByteWidth);
            si += sourceByteWidth;
            ti += targetByteWidth;
        }
        return this;
    }

    /**
     * Fills the image with a single color.
     * 
     * @param {number} rgba Color with which to fill image, expressed as a singlenumber in the form 0xRRGGBBAA, where AA is 0x00 for transparent and any other value for opaque.
     * @return {BitmapImage} The present image to allow for chaining.
     */

    fillRGBA(rgba) {
        const buf = this.bitmap.data;
        const bufByteWidth = this.bitmap.height * 4;
        
        let bi = 0;
        while (bi < bufByteWidth) {
            buf.writeUInt32BE(rgba, bi);
            bi += 4;
        }
        while (bi < buf.length) {
            buf.copy(buf, bi, 0, bufByteWidth);
            bi += bufByteWidth;
        }
        return this;
    }

    /**
     * Gets the RGBA number of the pixel at the given coordinate in the form 0xRRGGBBAA, where AA is the alpha value, with alpha 0x00 encoding to transparency in GIFs.
     * 
     * @param {number} x x-coord of pixel
     * @param {number} y y-coord of pixel
     * @return {number} RGBA of pixel in 0xRRGGBBAA form
     */

    getRGBA(x, y) {
        const bi = (y * this.bitmap.width + x) * 4;
        return this.bitmap.data.readUInt32BE(bi);
    }

    /**
     * Gets a set of all RGBA colors found within the image.
     * 
     * @return {Set} Set of all RGBA colors that the image contains.
     */

    getRGBASet() {
        const rgbaSet = new Set();
        const buf = this.bitmap.data;
        for (let bi = 0; bi < buf.length; bi += 4) {
            rgbaSet.add(buf.readUInt32BE(bi, true));
        }
        return rgbaSet;
    }

    /**
     * Converts the image to greyscale using inferred Adobe metrics.
     * 
     * @return {BitmapImage} The present image to allow for chaining.
     */

    greyscale() {
        const buf = this.bitmap.data;
        this.scan(0, 0, this.bitmap.width, this.bitmap.height, (x, y, idx) => {
            const grey = Math.round(
                0.299 * buf[idx] +
                0.587 * buf[idx + 1] +
                0.114 * buf[idx + 2]
            );
            buf[idx] = grey;
            buf[idx + 1] = grey;
            buf[idx + 2] = grey;
        });
        return this;
    }

    /**
     * Reframes the image as if placing a frame around the original image and replacing the original image with the newly framed image. When the new frame is strictly within the boundaries of the original image, this method crops the image. When any of the new boundaries exceed those of the original image, the `fillRGBA` must be provided to indicate the color with which to fill the extra space added to the image.
     * 
     * @param {number} xOffset The x-coord offset of the upper-left pixel of the desired image relative to the present image.
     * @param {number} yOffset The y-coord offset of the upper-left pixel of the desired image relative to the present image.
     * @param {number} width The width of the new image after reframing
     * @param {number} height The height of the new image after reframing
     * @param {number} fillRGBA The color with which to fill space added to the image as a result of the reframing, in 0xRRGGBBAA format, where AA is 0x00 to indicate transparent and a non-zero value to indicate opaque. This parameter is only required when the reframing exceeds the original boundaries (i.e. does not simply perform a crop).
     * @return {BitmapImage} The present image to allow for chaining.
     */

    reframe(xOffset, yOffset, width, height, fillRGBA) {
        const cropX = (xOffset < 0 ? 0 : xOffset);
        const cropY = (yOffset < 0 ? 0 : yOffset);
        const cropWidth = (width + cropX > this.bitmap.width ?
                this.bitmap.width - cropX : width);
        const cropHeight = (height + cropY > this.bitmap.height ?
                this.bitmap.height - cropY : height);
        const newX = (xOffset < 0 ? -xOffset : 0);
        const newY = (yOffset < 0 ? -yOffset : 0);

        let image;
        if (fillRGBA === undefined) {
            if (cropX !== xOffset || cropY != yOffset ||
                    cropWidth !== width || cropHeight !== height)
            {
                throw new GifError(`fillRGBA required for this reframing`);
            }
            image = new BitmapImage(width, height);
        }
        else {
            image = new BitmapImage(width, height, fillRGBA);
        }
        this.blit(image, newX, newY, cropX, cropY, cropWidth, cropHeight);
        this.bitmap = image.bitmap;
        return this;
    }

    /**
     * Scales the image size up by an integer factor. Each pixel of the original image becomes a square of the same color in the new image having a size of `factor` x `factor` pixels.
     * 
     * @param {number} factor The factor by which to scale up the image. Must be an integer >= 1.
     * @return {BitmapImage} The present image to allow for chaining.
     */

    scale(factor) {
        if (factor === 1) {
            return;
        }
        if (!Number.isInteger(factor) || factor < 1) {
            throw new Error("the scale must be an integer >= 1");
        }
        const sourceWidth = this.bitmap.width;
        const sourceHeight = this.bitmap.height;
        const destByteWidth = sourceWidth * factor * 4;
        const sourceBuf = this.bitmap.data;
        const destBuf = new Buffer(sourceHeight * destByteWidth * factor);
        let sourceIndex = 0;
        let priorDestRowIndex;
        let destIndex = 0;
        for (let y = 0; y < sourceHeight; ++y) {
            priorDestRowIndex = destIndex;
            for (let x = 0; x < sourceWidth; ++x) {
                const color = sourceBuf.readUInt32BE(sourceIndex, true);
                for (let cx = 0; cx < factor; ++cx) {
                    destBuf.writeUInt32BE(color, destIndex);
                    destIndex += 4;
                }
                sourceIndex += 4;
            }
            for (let cy = 1; cy < factor; ++cy) {
                destBuf.copy(destBuf, destIndex, priorDestRowIndex, destIndex);
                destIndex += destByteWidth;
                priorDestRowIndex += destByteWidth;
            }
        }
        this.bitmap = {
            width: sourceWidth * factor,
            height: sourceHeight * factor,
            data: destBuf
        };
        return this;
    }

    /**
     * Scans all coordinates of the image, handing each in turn to the provided handler function.
     *
     * @param {function} scanHandler A function(x: number, y: number, bi: number) to be called for each pixel of the image with that pixel's x-coord, y-coord, and index into the `data` buffer. The function accesses the pixel at this coordinate by accessing the `this.data` at index `bi`.
     * @see scanAllIndexes
     */

    scanAllCoords(scanHandler) {
        const width = this.bitmap.width;
        const bufferLength = this.bitmap.data.length;
        let x = 0;
        let y = 0;

        for (let bi = 0; bi < bufferLength; bi += 4) {
            scanHandler(x, y, bi);
            if (++x === width) {
                x = 0;
                ++y;
            }
        }
    }

    /**
     * Scans all pixels of the image, handing the index of each in turn to the provided handler function. Runs a bit faster than `scanAllCoords()`, should the handler not need pixel coordinates.
     *
     * @param {function} scanHandler A function(bi: number) to be called for each pixel of the image with that pixel's index into the `data` buffer. The pixels is found at index 'bi' within `this.data`.
     * @see scanAllCoords
     */

    scanAllIndexes(scanHandler) {
        const bufferLength = this.bitmap.data.length;
        for (let bi = 0; bi < bufferLength; bi += 4) {
            scanHandler(bi);
        }
    }
}

module.exports = BitmapImage;


/***/ },

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gif.js"
(__unused_webpack_module, exports) {

"use strict";


/** @class Gif */

class Gif {

    // width - width of GIF in pixels
    // height - height of GIF in pixels
    // loops - 0 = unending; (n > 0) = iterate n times
    // usesTransparency - whether any frames have transparent pixels
    // colorScope - scope of color tables in GIF
    // frames - array of frames
    // buffer - GIF-formatted data

    /**
     * Gif is a class representing an encoded GIF. It is intended to be a read-only representation of a byte-encoded GIF. Only encoders and decoders should be creating instances of this class.
     * 
     * Property | Description
     * --- | ---
     * width | width of the GIF at its widest
     * height | height of the GIF at its highest
     * loops | the number of times the GIF should loop before stopping; 0 => loop indefinitely
     * usesTransparency | boolean indicating whether at least one frame contains at least one transparent pixel
     * colorScope | the scope of the color tables as encoded within the GIF; either Gif.GlobalColorsOnly (== 1) or Gif.LocalColorsOnly (== 2).
     * frames | a array of GifFrame instances, one for each frame of the GIF
     * buffer | a Buffer holding the encoding's byte data
     * 
     * Its constructor should only ever be called by the GIF encoder or decoder.
     *
     * @param {Buffer} buffer A Buffer containing the encoded bytes
     * @param {GifFrame[]} frames Array of frames found in the encoding
     * @param {object} spec Properties of the encoding as listed above
     */

    constructor(buffer, frames, spec) {
        this.width = spec.width;
        this.height = spec.height;
        this.loops = spec.loops;
        this.usesTransparency = spec.usesTransparency;
        this.colorScope = spec.colorScope;
        this.frames = frames;
        this.buffer = buffer;
    }
}

Gif.GlobalColorsPreferred = 0;
Gif.GlobalColorsOnly = 1;
Gif.LocalColorsOnly = 2;

/** @class GifError */

class GifError extends Error {

    /**
     * GifError is a class representing a GIF-related error
     * 
     * @param {string|Error} messageOrError
     */

    constructor(messageOrError) {
        super(messageOrError);
        if (messageOrError instanceof Error) {
            this.stack = 'Gif' + messageOrError.stack;
        }
    }
}

exports.Gif = Gif;
exports.GifError = GifError;


/***/ },

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifcodec.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";


const Omggif = __webpack_require__("../../node_modules/.pnpm/omggif@1.0.10/node_modules/omggif/omggif.js");
const { Gif, GifError } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gif.js");

// allow circular dependency with GifUtil
function GifUtil() {
    const data = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifutil.js");

    GifUtil = function () {
      return data;
    };

  return data;
}

const { GifFrame } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifframe.js");

const PER_GIF_OVERHEAD = 200; // these are guesses at upper limits
const PER_FRAME_OVERHEAD = 100;

// Note: I experimented with accepting a global color table when encoding and returning the global color table when decoding. Doing this properly greatly increased the complexity of the code and the amount of clock cycles required. The main issue is that each frame can specify any color of the global color table to be transparent within the frame, while this GIF library strives to hide GIF formatting details from its clients. E.g. it's possible to have 256 colors in the global color table and different transparencies in each frame, requiring clients to either provide per-frame transparency indexes, or for arcane reasons that won't be apparent to client developers, encode some GIFs with local color tables that previously decoded with global tables.

/** @class GifCodec */

class GifCodec
{
    // _transparentRGBA - RGB given to transparent pixels (alpha=0) on decode; defaults to null indicating 0x000000, which is fastest

    /**
     * GifCodec is a class that both encodes and decodes GIFs. It implements both the `encode()` method expected of an encoder and the `decode()` method expected of a decoder, and it wraps the `omggif` GIF encoder/decoder package. GifCodec serves as this library's default encoder and decoder, but it's possible to wrap other GIF encoders and decoders for use by `gifwrap` as well. GifCodec will not encode GIFs with interlacing.
     * 
     * Instances of this class are stateless and can be shared across multiple encodings and decodings.
     * 
     * Its constructor takes one option argument:
     * 
     * @param {object} options Optionally takes an objection whose only possible property is `transparentRGB`. Images are internally represented in RGBA format, where A is the alpha value of a pixel. When `transparentRGB` is provided, this RGB value (excluding alpha) is assigned to transparent pixels, which are also given alpha value 0x00. (All opaque pixels are given alpha value 0xFF). The RGB color of transparent pixels shouldn't matter for most applications. Defaults to 0x000000.
     */

    constructor(options = {}) {
        this._transparentRGB = null; // 0x000000
        if (typeof options.transparentRGB === 'number' &&
                options.transparentRGB !== 0)
        {
            this._transparentRGBA = options.transparentRGB * 256;
        }
        this._testInitialBufferSize = 0; // assume no buffer scaling test
    }

    /**
     * Decodes a GIF from a Buffer to yield an instance of Gif. Transparent pixels of the GIF are given alpha values of 0x00, and opaque pixels are given alpha values of 0xFF. The RGB values of transparent pixels default to 0x000000 but can be overridden by the constructor's `transparentRGB` option.
     * 
     * @param {Buffer} buffer Bytes of an encoded GIF to decode.
     * @return {Promise} A Promise that resolves to an instance of the Gif class, representing the encoded GIF.
     * @throws {GifError} Error upon encountered an encoding-related problem with a GIF, so that the caller can distinguish between software errors and problems with GIFs.
     */

    decodeGif(buffer) {
        try {
            let reader;
            try {
                reader = new Omggif.GifReader(buffer);
            }
            catch (err) {
                throw new GifError(err);
            }
            const frameCount = reader.numFrames();
            const frames = [];
            const spec = {
                width: reader.width,
                height: reader.height,
                loops: reader.loopCount()
            };

            spec.usesTransparency = false;
            for (let i = 0; i < frameCount; ++i) {
                const frameInfo =
                        this._decodeFrame(reader, i, spec.usesTransparency);
                frames.push(frameInfo.frame);
                if (frameInfo.usesTransparency) {
                    spec.usesTransparency = true;
                }
            }
            return Promise.resolve(new Gif(buffer, frames, spec));
        }
        catch (err) {
            return Promise.reject(err);
        }
    }

    /**
     * Encodes a GIF from provided frames. Each pixel having an alpha value of 0x00 renders as transparent within the encoding, while all pixels of non-zero alpha value render as opaque.
     * 
     * @param {GifFrame[]} frames Array of frames to encode
     * @param {object} spec An optional object that may provide values for `loops` and `colorScope`, as defined for the Gif class. However, `colorSpace` may also take the value Gif.GlobalColorsPreferred (== 0) to indicate that the encoder should attempt to create only a global color table. `loop` defaults to 0, looping indefinitely. Set `loop` to null to disable looping, playing only once. `colorScope` defaults to Gif.GlobalColorsPreferred.
     * @return {Promise} A Promise that resolves to an instance of the Gif class, representing the encoded GIF.
     * @throws {GifError} Error upon encountered an encoding-related problem with a GIF, so that the caller can distinguish between software errors and problems with GIFs.
     */

    encodeGif(frames, spec = {}) {
        try {
            if (frames === null || frames.length === 0) {
                throw new GifError("there are no frames");
            }
            const dims = GifUtil().getMaxDimensions(frames);

            spec = Object.assign({}, spec); // don't munge caller's spec
            spec.width = dims.maxWidth;
            spec.height = dims.maxHeight;
            if (spec.loops === undefined) {
                spec.loops = 0;
            }
            spec.colorScope = spec.colorScope || Gif.GlobalColorsPreferred;

            return Promise.resolve(this._encodeGif(frames, spec));
        }
        catch (err) {
            return Promise.reject(err);
        }
    }

    _decodeFrame(reader, frameIndex, alreadyUsedTransparency) {
        let info, buffer;
        try {
            info = reader.frameInfo(frameIndex);
            buffer = new Buffer(reader.width * reader.height * 4);
            reader.decodeAndBlitFrameRGBA(frameIndex, buffer);
            if (info.width !== reader.width || info.height !== reader.height) {
                if (info.y) {
                    // skip unused rows
                    buffer = buffer.slice(info.y * reader.width * 4);
                }
                if (reader.width > info.width) {
                    // skip scanstride
                    for (let ii = 0; ii < info.height; ++ii) {
                        buffer.copy(buffer, ii * info.width * 4,
                            (info.x + ii * reader.width) * 4,
                            (info.x + ii * reader.width) * 4 + info.width * 4);
                    }
                }
                // trim buffer to size
                buffer = buffer.slice(0, info.width * info.height * 4);
            }
        }
        catch (err) {
            throw new GifError(err);
        }

        let usesTransparency = false;
        if (this._transparentRGBA === null) {
            if (!alreadyUsedTransparency) {
                for (let i = 3; i < buffer.length; i += 4) {
                    if (buffer[i] === 0) {
                        usesTransparency = true;
                        i = buffer.length;
                    }
                }
            }
        }
        else {
            for (let i = 3; i < buffer.length; i += 4) {
                if (buffer[i] === 0) {
                    buffer.writeUInt32BE(this._transparentRGBA, i - 3);
                    usesTransparency = true; // GIF might encode unused index
                }
            }
        }

        const frame = new GifFrame(info.width, info.height, buffer, {
            xOffset: info.x,
            yOffset: info.y,
            disposalMethod: info.disposal,
            interlaced: info.interlaced,
            delayCentisecs: info.delay
        });
        return { frame, usesTransparency };
    }

    _encodeGif(frames, spec) {
        let colorInfo;
        if (spec.colorScope === Gif.LocalColorsOnly) {
            colorInfo = GifUtil().getColorInfo(frames, 0);
        }
        else {
            colorInfo = GifUtil().getColorInfo(frames, 256);
            if (!colorInfo.colors) { // if global palette impossible
                if (spec.colorScope === Gif.GlobalColorsOnly) {
                    throw new GifError(
                            "Too many color indexes for global color table");
                }
                spec.colorScope = Gif.LocalColorsOnly
            }
        }
        spec.usesTransparency = colorInfo.usesTransparency;

        const localPalettes = colorInfo.palettes;
        if (spec.colorScope === Gif.LocalColorsOnly) {
            const localSizeEst = 2000; //this._getSizeEstimateLocal(localPalettes, frames);
            return _encodeLocal(frames, spec, localSizeEst, localPalettes);
        }

        const globalSizeEst = 2000; //this._getSizeEstimateGlobal(colorInfo, frames);
        return _encodeGlobal(frames, spec, globalSizeEst, colorInfo);
    }

    _getSizeEstimateGlobal(globalPalette, frames) {
        if (this._testInitialBufferSize > 0) {
            return this._testInitialBufferSize;
        }
        let sizeEst = PER_GIF_OVERHEAD + 3*256 /* max palette size*/;
        const pixelBitWidth = _getPixelBitWidth(globalPalette);
        frames.forEach(frame => {
            sizeEst += _getFrameSizeEst(frame, pixelBitWidth);
        });
        return sizeEst; // should be the upper limit
    }

    _getSizeEstimateLocal(palettes, frames) {
        if (this._testInitialBufferSize > 0) {
            return this._testInitialBufferSize;
        }
        let sizeEst = PER_GIF_OVERHEAD;
        for (let i = 0; i < frames.length; ++i ) {
            const palette = palettes[i];
            const pixelBitWidth = _getPixelBitWidth(palette);
            sizeEst += _getFrameSizeEst(frames[i], pixelBitWidth);
        }
        return sizeEst; // should be the upper limit
    }
}
exports.GifCodec = GifCodec;

function _colorLookupLinear(colors, color) {
    const index = colors.indexOf(color);
    return (index === -1 ? null : index);
}

function _colorLookupBinary(colors, color) {
    // adapted from https://stackoverflow.com/a/10264318/650894
    var lo = 0, hi = colors.length - 1, mid;
    while (lo <= hi) {
        mid = Math.floor((lo + hi)/2);
        if (colors[mid] > color)
            hi = mid - 1;
        else if (colors[mid] < color)
            lo = mid + 1;
        else
            return mid;
    }
    return null;
}

function _encodeGlobal(frames, spec, bufferSizeEst, globalPalette) {
    // would be inefficient for frames to lookup colors in extended palette 
    const extendedGlobalPalette = {
        colors: globalPalette.colors.slice(),
        usesTransparency: globalPalette.usesTransparency
    };
    _extendPaletteToPowerOf2(extendedGlobalPalette);
    const options = {
        palette: extendedGlobalPalette.colors,
        loop: spec.loops
    };
    let buffer = new Buffer(bufferSizeEst);
    let gifWriter;
    try {
        gifWriter = new Omggif.GifWriter(buffer, spec.width, spec.height,
                            options);
    }
    catch (err) {
        throw new GifError(err);
    }
    for (let i = 0; i < frames.length; ++i) {
        buffer = _writeFrame(gifWriter, i, frames[i], globalPalette, false);
    }
    return new Gif(buffer.slice(0, gifWriter.end()), frames, spec);
}

function _encodeLocal(frames, spec, bufferSizeEst, localPalettes) {
    const options = {
        loop: spec.loops
    };
    let buffer = new Buffer(bufferSizeEst);
    let gifWriter;
    try {
        gifWriter = new Omggif.GifWriter(buffer, spec.width, spec.height,
                            options);
    }                            
    catch (err) {
        throw new GifError(err);
    }
    for (let i = 0; i < frames.length; ++i) {
        buffer = _writeFrame(gifWriter, i, frames[i], localPalettes[i], true);
    }
    return new Gif(buffer.slice(0, gifWriter.end()), frames, spec);
}

function _extendPaletteToPowerOf2(palette) {
    const colors = palette.colors;
    if (palette.usesTransparency) {
        colors.push(0);
    }
    const colorCount = colors.length;
    let powerOf2 = 2;
    while (colorCount > powerOf2) {
        powerOf2 <<= 1;
    }
    colors.length = powerOf2;
    colors.fill(0, colorCount);
}

function _getFrameSizeEst(frame, pixelBitWidth) {
    let byteLength = frame.bitmap.width * frame.bitmap.height;
    byteLength = Math.ceil(byteLength * pixelBitWidth / 8);
    byteLength += Math.ceil(byteLength / 255); // add block size bytes
    // assume maximum palete size because it might get extended for power of 2
    return (PER_FRAME_OVERHEAD + byteLength + 3 * 256 /* largest palette */);
}

function _getIndexedImage(frameIndex, frame, palette) {
    const colors = palette.colors;
    const colorToIndexFunc = (colors.length <= 8 ? // guess at the break-even
            _colorLookupLinear : _colorLookupBinary);
    const colorBuffer = frame.bitmap.data;
    const indexBuffer = new Buffer(colorBuffer.length/4);
    let transparentIndex = colors.length;
    let i = 0, j = 0;

    while (i < colorBuffer.length) {
        if (colorBuffer[i + 3] !== 0) {
            const color = (colorBuffer.readUInt32BE(i, true) >> 8) & 0xFFFFFF;
            // caller guarantees that the color will be in the palette
            indexBuffer[j] = colorToIndexFunc(colors, color);
        }
        else {
            indexBuffer[j] = transparentIndex;
        }
        i += 4; // skip alpha
        ++j;
    }

    if (palette.usesTransparency) {
        if (transparentIndex === 256) {
            throw new GifError(`Frame ${frameIndex} already has 256 colors` +
                    `and so can't use transparency`);
        }
    }
    else {
        transparentIndex = null;
    }

    return { buffer: indexBuffer, transparentIndex };
}

function _getPixelBitWidth(palette) {
    let indexCount = palette.indexCount;
    let pixelBitWidth = 0;
    --indexCount; // start at maximum index
    while (indexCount) {
        ++pixelBitWidth;
        indexCount >>= 1;
    }
    return (pixelBitWidth > 0 ? pixelBitWidth : 1);
}

function _writeFrame(gifWriter, frameIndex, frame, palette, isLocalPalette) {
    if (frame.interlaced) {
        throw new GifError("writing interlaced GIFs is not supported");
    }
    const frameInfo = _getIndexedImage(frameIndex, frame, palette);
    const options = {
        delay: frame.delayCentisecs,
        disposal: frame.disposalMethod,
        transparent: frameInfo.transparentIndex
    };
    if (isLocalPalette) {
        _extendPaletteToPowerOf2(palette); // ok 'cause palette never used again
        options.palette = palette.colors;
    }
    try {
        let buffer = gifWriter.getOutputBuffer();
        let startOfFrame = gifWriter.getOutputBufferPosition();
        let endOfFrame;
        let tryAgain = true;

        while (tryAgain) {
            endOfFrame = gifWriter.addFrame(frame.xOffset, frame.yOffset,
                    frame.bitmap.width, frame.bitmap.height, frameInfo.buffer, options);
            tryAgain = false;
            if (endOfFrame >= buffer.length - 1) {
                const biggerBuffer = new Buffer(buffer.length * 1.5);
                buffer.copy(biggerBuffer);
                gifWriter.setOutputBuffer(biggerBuffer);
                gifWriter.setOutputBufferPosition(startOfFrame);
                buffer = biggerBuffer;
                tryAgain = true;
            }
        }
        return buffer;
    }
    catch (err) {
        throw new GifError(err);
    }
}


/***/ },

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifframe.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";


const BitmapImage = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/bitmapimage.js");
const { GifError } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gif.js");

/** @class GifFrame */

class GifFrame extends BitmapImage {

    // xOffset - x offset of bitmap on GIF (defaults to 0)
    // yOffset - y offset of bitmap on GIF (defaults to 0)
    // disposalMethod - pixel disposal method when handling partial images
    // delayCentisecs - duration of frame in hundredths of a second
    // interlaced - whether the image is interlaced (defaults to false)

    /**
     * GifFrame is a class representing an image frame of a GIF. GIFs contain one or more instances of GifFrame.
     * 
     * Property | Description
     * --- | ---
     * xOffset | x-coord of position within GIF at which to render the image (defaults to 0)
     * yOffset | y-coord of position within GIF at which to render the image (defaults to 0)
     * disposalMethod | GIF disposal method; only relevant when the frames aren't all the same size (defaults to 2, disposing to background color)
     * delayCentisecs | duration of the frame in hundreths of a second
     * interlaced | boolean indicating whether the frame renders interlaced
     * 
     * Its constructor supports the following signatures:
     * 
     * * new GifFrame(bitmap: {width: number, height: number, data: Buffer}, options?)
     * * new GifFrame(bitmapImage: BitmapImage, options?)
     * * new GifFrame(width: number, height: number, buffer: Buffer, options?)
     * * new GifFrame(width: number, height: number, backgroundRGBA?: number, options?)
     * * new GifFrame(frame: GifFrame)
     * 
     * See the base class BitmapImage for a discussion of all parameters but `options` and `frame`. `options` is an optional argument providing initial values for the above-listed GifFrame properties. Each property within option is itself optional.
     * 
     * Provide a `frame` to the constructor to create a clone of the provided frame. The new frame includes a copy of the provided frame's pixel data so that each can subsequently be modified without affecting each other.
     */

    constructor(...args) {
        super(...args);
        if (args[0] instanceof GifFrame) {
            // copy a provided GifFrame
            const source = args[0];
            this.xOffset = source.xOffset;
            this.yOffset = source.yOffset;
            this.disposalMethod = source.disposalMethod;
            this.delayCentisecs = source.delayCentisecs;
            this.interlaced = source.interlaced;
        }
        else {
            const lastArg = args[args.length - 1];
            let options = {};
            if (typeof lastArg === 'object' && !(lastArg instanceof BitmapImage)) {
                options = lastArg;
            }
            this.xOffset = options.xOffset || 0;
            this.yOffset = options.yOffset || 0;
            this.disposalMethod = (options.disposalMethod !== undefined ?
                    options.disposalMethod : GifFrame.DisposeToBackgroundColor);
            this.delayCentisecs = options.delayCentisecs || 8;
            this.interlaced = options.interlaced || false;
        }
    }

    /**
     * Get a summary of the colors found within the frame. The return value is an object of the following form:
     * 
     * Property | Description
     * --- | ---
     * colors | An array of all the opaque colors found within the frame. Each color is given as an RGB number of the form 0xRRGGBB. The array is sorted by increasing number. Will be an empty array when the image is completely transparent.
     * usesTransparency | boolean indicating whether there are any transparent pixels within the frame. A pixel is considered transparent if its alpha value is 0x00.
     * indexCount | The number of color indexes required to represent this palette of colors. It is equal to the number of opaque colors plus one if the image includes transparency.
     * 
     * @return {object} An object representing a color palette as described above.
     */

    getPalette() {
        // returns with colors sorted low to high
        const colorSet = new Set();
        const buf = this.bitmap.data;
        let i = 0;
        let usesTransparency = false;
        while (i < buf.length) {
            if (buf[i + 3] === 0) {
                usesTransparency = true;
            }
            else {
                // can eliminate the bitshift by starting one byte prior
                const color = (buf.readUInt32BE(i, true) >> 8) & 0xFFFFFF;
                colorSet.add(color);
            }
            i += 4; // skip alpha
        }
        const colors = new Array(colorSet.size);
        const iter = colorSet.values();
        for (i = 0; i < colors.length; ++i) {
            colors[i] = iter.next().value;
        }
        colors.sort((a, b) => (a - b));
        let indexCount = colors.length;
        if (usesTransparency) {
            ++indexCount;
        }
        return { colors, usesTransparency, indexCount };
    }
}

GifFrame.DisposeToAnything = 0;
GifFrame.DisposeNothing = 1;
GifFrame.DisposeToBackgroundColor = 2;
GifFrame.DisposeToPrevious = 3;

exports.GifFrame = GifFrame;


/***/ },

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifutil.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";


/** @namespace GifUtil */

const fs = __webpack_require__("fs");
const ImageQ = __webpack_require__("../../node_modules/.pnpm/image-q@4.0.0/node_modules/image-q/dist/cjs/image-q.cjs");

const BitmapImage = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/bitmapimage.js");
const { GifFrame } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifframe.js");
const { GifError } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gif.js");
const { GifCodec } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifcodec.js");

const INVALID_SUFFIXES = ['.jpg', '.jpeg', '.png', '.bmp'];

const defaultCodec = new GifCodec();

/**
 * cloneFrames() clones provided frames. It's a utility method for cloning an entire array of frames at once.
 * 
 * @function cloneFrames
 * @memberof GifUtil
 * @param {GifFrame[]} frames An array of GifFrame instances to clone
 * @return {GifFrame[]} An array of GifFrame clones of the provided frames.
 */

exports.cloneFrames = function (frames) {
    let clones = [];
    frames.forEach(frame => {

        clones.push(new GifFrame(frame));
    });
    return clones;
}

/**
 * getColorInfo() gets information about the colors used in the provided frames. The method is able to return an array of all colors found across all frames.
 * 
 * `maxGlobalIndex` controls whether the computation short-circuits to avoid doing work that the caller doesn't need. The method only returns `colors` and `indexCount` for the colors across all frames when the number of indexes required to store the colors and transparency in a GIF (which is the value of `indexCount`) is less than or equal to `maxGlobalIndex`. Such short-circuiting is useful when the caller just needs to determine whether any frame includes transparency.
 * 
 * @function getColorInfo
 * @memberof GifUtil
 * @param {GifFrame[]} frames Frames to examine for color and transparency.
 * @param {number} maxGlobalIndex Maximum number of color indexes (including one for transparency) allowed among the returned compilation of colors. `colors` and `indexCount` are not returned if the number of color indexes required to accommodate  all frames exceeds this number. Returns `colors` and `indexCount` by default.
 * @returns {object} Object containing at least `palettes` and `usesTransparency`. `palettes` is an array of all the palettes returned by GifFrame#getPalette(). `usesTransparency` indicates whether at least one frame uses transparency. If `maxGlobalIndex` is not exceeded, the object also contains `colors`, an array of all colors (RGB) found across all palettes, sorted by increasing value, and `indexCount` indicating the number of indexes required to store the colors and the transparency in a GIF.
 * @throws {GifError} When any frame requires more than 256 color indexes.
 */

exports.getColorInfo = function (frames, maxGlobalIndex) {
    let usesTransparency = false;
    const palettes = [];
    for (let i = 0; i < frames.length; ++i) {
        let palette = frames[i].getPalette();
        if (palette.usesTransparency) {
            usesTransparency = true;
        }
        if (palette.indexCount > 256) {
            throw new GifError(`Frame ${i} uses more than 256 color indexes`);
        }
        palettes.push(palette);
    }
    if (maxGlobalIndex === 0) {
        return { usesTransparency, palettes };
    }

    const globalColorSet = new Set();
    palettes.forEach(palette => {

        palette.colors.forEach(color => {

            globalColorSet.add(color);
        });
    });
    let indexCount = globalColorSet.size;
    if (usesTransparency) {
        // odd that GIF requires a color table entry at transparent index
        ++indexCount;
    }
    if (maxGlobalIndex && indexCount > maxGlobalIndex) {
        return { usesTransparency, palettes };
    }
    
    const colors = new Array(globalColorSet.size);
    const iter = globalColorSet.values();
    for (let i = 0; i < colors.length; ++i) {
        colors[i] = iter.next().value;
    }
    colors.sort((a, b) => (a - b));
    return { colors, indexCount, usesTransparency, palettes };
};

/**
 * copyAsJimp() returns a Jimp that contains a copy of the provided bitmap image (which may be either a BitmapImage or a GifFrame). Modifying the Jimp does not affect the provided bitmap image. This method serves as a macro for simplifying working with Jimp.
 *
 * @function copyAsJimp
 * @memberof GifUtil
 * @param {object} Reference to the Jimp package, keeping this library from being dependent on Jimp.
 * @param {bitmapImageToCopy} Instance of BitmapImage (may be a GifUtil) with which to source the Jimp.
 * @return {object} An new instance of Jimp containing a copy of the image in bitmapImageToCopy.
 */
 
exports.copyAsJimp = function (jimp, bitmapImageToCopy) {
    return exports.shareAsJimp(jimp, new BitmapImage(bitmapImageToCopy));
};

/**
 * getMaxDimensions() returns the pixel width and height required to accommodate all of the provided frames, according to the offsets and dimensions of each frame.
 * 
 * @function getMaxDimensions
 * @memberof GifUtil
 * @param {GifFrame[]} frames Frames to measure for their aggregate maximum dimensions.
 * @return {object} An object of the form {maxWidth, maxHeight} indicating the maximum width and height required to accommodate all frames.
 */

exports.getMaxDimensions = function (frames) {
    let maxWidth = 0, maxHeight = 0;
    frames.forEach(frame => {
        const width = frame.xOffset + frame.bitmap.width;
        if (width > maxWidth) {
            maxWidth = width;
        }
        const height = frame.yOffset + frame.bitmap.height;
        if (height > maxHeight) {
            maxHeight = height;
        }
    });
    return { maxWidth, maxHeight };
};

/**
 * Quantizes colors so that there are at most a given number of color indexes (including transparency) across all provided images. Uses an algorithm by Anthony Dekker.
 * 
 * The method treats different RGBA combinations as different colors, so if the frame has multiple alpha values or multiple RGB values for an alpha value, the caller may first want to normalize them by converting all transparent pixels to the same RGBA values.
 * 
 * The method may increase the number of colors if there are fewer than the provided maximum.
 * 
 * @function quantizeDekker
 * @memberof GifUtil
 * @param {BitmapImage|BitmapImage[]} imageOrImages Image or array of images (such as GifFrame instances) to be color-quantized. Quantizing across multiple images ensures color consistency from frame to frame.
 * @param {number} maxColorIndexes The maximum number of color indexes that will exist in the palette after completing quantization. Defaults to 256.
 * @param {object} dither (optional) An object configuring the dithering to apply. The properties are as followings, imported from the [`image-q` package](https://github.com/ibezkrovnyi/image-quantization) without explanation: { `ditherAlgorithm`: One of 'FloydSteinberg', 'FalseFloydSteinberg', 'Stucki', 'Atkinson', 'Jarvis', 'Burkes', 'Sierra', 'TwoSierra', 'SierraLite'; `minimumColorDistanceToDither`: (optional) A number defaulting to 0; `serpentine`: (optional) A boolean defaulting to true; `calculateErrorLikeGIMP`: (optional) A boolean defaulting to false. }
 */

exports.quantizeDekker = function (imageOrImages, maxColorIndexes, dither) {
    maxColorIndexes = maxColorIndexes || 256;
    _quantize(imageOrImages, 'NeuQuantFloat', maxColorIndexes, 0, dither);
}

/**
 * Quantizes colors so that there are at most a given number of color indexes (including transparency) across all provided images. Uses an algorithm by Leon Sorokin. This quantization method differs from the other two by likely never increasing the number of colors, should there be fewer than the provided maximum.
 * 
 * The method treats different RGBA combinations as different colors, so if the frame has multiple alpha values or multiple RGB values for an alpha value, the caller may first want to normalize them by converting all transparent pixels to the same RGBA values.
 * 
 * @function quantizeSorokin
 * @memberof GifUtil
 * @param {BitmapImage|BitmapImage[]} imageOrImages Image or array of images (such as GifFrame instances) to be color-quantized. Quantizing across multiple images ensures color consistency from frame to frame.
 * @param {number} maxColorIndexes The maximum number of color indexes that will exist in the palette after completing quantization. Defaults to 256.
 * @param {string} histogram (optional) Histogram method: 'top-pop' for global top-population, 'min-pop' for minimum-population threshhold within subregions. Defaults to 'min-pop'.
 * @param {object} dither (optional) An object configuring the dithering to apply, as explained for `quantizeDekker()`.
 */

exports.quantizeSorokin = function (imageOrImages, maxColorIndexes, histogram, dither) {
    maxColorIndexes = maxColorIndexes || 256;
    histogram = histogram || 'min-pop';
    let histogramID;
    switch (histogram) {
        case 'min-pop':
        histogramID = 2;
        break;

        case 'top-pop':
        histogramID = 1;
        break

        default:
        throw new Error(`Invalid quantizeSorokin histogram '${histogram}'`);
    }
    _quantize(imageOrImages, 'RGBQuant', maxColorIndexes, histogramID, dither);
}

/**
 * Quantizes colors so that there are at most a given number of color indexes (including transparency) across all provided images. Uses an algorithm by Xiaolin Wu.
 * 
 * The method treats different RGBA combinations as different colors, so if the frame has multiple alpha values or multiple RGB values for an alpha value, the caller may first want to normalize them by converting all transparent pixels to the same RGBA values.
 * 
 * The method may increase the number of colors if there are fewer than the provided maximum.
 * 
 * @function quantizeWu
 * @memberof GifUtil
 * @param {BitmapImage|BitmapImage[]} imageOrImages Image or array of images (such as GifFrame instances) to be color-quantized. Quantizing across multiple images ensures color consistency from frame to frame.
 * @param {number} maxColorIndexes The maximum number of color indexes that will exist in the palette after completing quantization. Defaults to 256.
 * @param {number} significantBits (optional) This is the number of significant high bits in each RGB color channel. Takes integer values from 1 through 8. Higher values correspond to higher quality. Defaults to 5.
 * @param {object} dither (optional) An object configuring the dithering to apply, as explained for `quantizeDekker()`.
 */

exports.quantizeWu = function (imageOrImages, maxColorIndexes, significantBits, dither) {
    maxColorIndexes = maxColorIndexes || 256;
    significantBits = significantBits || 5;
    if (significantBits < 1 || significantBits > 8) {
        throw new Error("Invalid quantization quality");
    }
    _quantize(imageOrImages, 'WuQuant', maxColorIndexes, significantBits, dither);
}

/**
 * read() decodes an encoded GIF, whether provided as a filename or as a byte buffer.
 * 
 * @function read
 * @memberof GifUtil
 * @param {string|Buffer} source Source to decode. When a string, it's the GIF filename to load and parse. When a Buffer, it's an encoded GIF to parse.
 * @param {object} decoder An optional GIF decoder object implementing the `decode` method of class GifCodec. When provided, the method decodes the GIF using this decoder. When not provided, the method uses GifCodec.
 * @return {Promise} A Promise that resolves to an instance of the Gif class, representing the decoded GIF.
 */

exports.read = function (source, decoder) {
    decoder = decoder || defaultCodec;
    if (Buffer.isBuffer(source)) {
        return decoder.decodeGif(source);
    }
    return _readBinary(source)
    .then(buffer => {

        return decoder.decodeGif(buffer);
    });
};

/**
 * shareAsJimp() returns a Jimp that shares a bitmap with the provided bitmap image (which may be either a BitmapImage or a GifFrame). Modifying the image in either the Jimp or the BitmapImage affects the other objects. This method serves as a macro for simplifying working with Jimp.
 *
 * @function shareAsJimp
 * @memberof GifUtil
 * @param {object} Reference to the Jimp package, keeping this library from being dependent on Jimp.
 * @param {bitmapImageToShare} Instance of BitmapImage (may be a GifUtil) with which to source the Jimp.
 * @return {object} An new instance of Jimp that shares the image in bitmapImageToShare.
 */
 
exports.shareAsJimp = function (jimp, bitmapImageToShare) {
    const jimpImage = new jimp(bitmapImageToShare.bitmap.width,
            bitmapImageToShare.bitmap.height, 0);
    jimpImage.bitmap.data = bitmapImageToShare.bitmap.data;
    return jimpImage;
};

/**
 * write() encodes a GIF and saves it as a file.
 * 
 * @function write
 * @memberof GifUtil
 * @param {string} path Filename to write GIF out as. Will overwrite an existing file.
 * @param {GifFrame[]} frames Array of frames to be written into GIF.
 * @param {object} spec An optional object that may provide values for `loops` and `colorScope`, as defined for the Gif class. However, `colorSpace` may also take the value Gif.GlobalColorsPreferred (== 0) to indicate that the encoder should attempt to create only a global color table. `loop` defaults to 0, looping indefinitely, and `colorScope` defaults to Gif.GlobalColorsPreferred.
 * @param {object} encoder An optional GIF encoder object implementing the `encode` method of class GifCodec. When provided, the method encodes the GIF using this encoder. When not provided, the method uses GifCodec.
 * @return {Promise} A Promise that resolves to an instance of the Gif class, representing the encoded GIF.
 */

exports.write = function (path, frames, spec, encoder) {
    encoder = encoder || defaultCodec;
    const matches = path.match(/\.[a-zA-Z]+$/); // prevent accidents
    if (matches !== null &&
            INVALID_SUFFIXES.includes(matches[0].toLowerCase()))
    {
        throw new Error(`GIF '${path}' has an unexpected suffix`);
    }

    return encoder.encodeGif(frames, spec)
    .then(gif => {

        return _writeBinary(path, gif.buffer)
        .then(() => {

            return gif;
        });
    });
};

function _quantize(imageOrImages, method, maxColorIndexes, modifier, dither) {
    const images = Array.isArray(imageOrImages) ? imageOrImages : [imageOrImages];
    const ditherAlgs = [
        'FloydSteinberg',
        'FalseFloydSteinberg',
        'Stucki',
        'Atkinson',
        'Jarvis',
        'Burkes',
        'Sierra',
        'TwoSierra',
        'SierraLite'
    ];

    if (dither) {
        if (ditherAlgs.indexOf(dither.ditherAlgorithm) < 0) {
            throw new Error(`Invalid ditherAlgorithm '${dither.ditherAlgorithm}'`);
        }
        if (dither.serpentine === undefined) {
            dither.serpentine = true;
        }
        if (dither.minimumColorDistanceToDither === undefined) {
            dither.minimumColorDistanceToDither = 0;
        }
        if (dither.calculateErrorLikeGIMP === undefined) {
            dither.calculateErrorLikeGIMP = false;
        }
    }

    const distCalculator = new ImageQ.distance.Euclidean();
    const quantizer = new ImageQ.palette[method](distCalculator, maxColorIndexes, modifier);
    let imageMaker;
    if (dither) {
        imageMaker = new ImageQ.image.ErrorDiffusionArray(
            distCalculator,
            ImageQ.image.ErrorDiffusionArrayKernel[dither.ditherAlgorithm],
            dither.serpentine,
            dither.minimumColorDistanceToDither,
            dither.calculateErrorLikeGIMP
        );
    }
    else {
        imageMaker = new ImageQ.image.NearestColor(distCalculator);
    }

    const inputContainers = [];
    images.forEach(image => {

        const imageBuf = image.bitmap.data;
        const inputBuf = new ArrayBuffer(imageBuf.length);
        const inputArray = new Uint32Array(inputBuf);
        for (let bi = 0, ai = 0; bi < imageBuf.length; bi += 4, ++ai) {
            inputArray[ai] = imageBuf.readUInt32LE(bi, true);
        }
        const inputContainer = ImageQ.utils.PointContainer.fromUint32Array(
                inputArray, image.bitmap.width, image.bitmap.height);
        quantizer.sample(inputContainer);
        inputContainers.push(inputContainer);
    });
    
    const limitedPalette = quantizer.quantizeSync();

    for (let i = 0; i < images.length; ++i) {
        const imageBuf = images[i].bitmap.data;
        const outputContainer = imageMaker.quantizeSync(inputContainers[i], limitedPalette);
        const outputArray = outputContainer.toUint32Array();
        for (let bi = 0, ai = 0; bi < imageBuf.length; bi += 4, ++ai) {
            imageBuf.writeUInt32LE(outputArray[ai], bi);
        }
    }
}

function _readBinary(path) {
    // TBD: add support for URLs
    return new Promise((resolve, reject) => {

        fs.readFile(path, (err, buffer) => {

            if (err) {
                return reject(err);
            }
            return resolve(buffer);
        });
    });
}

function _writeBinary(path, buffer) {
    // TBD: add support for URLs
    return new Promise((resolve, reject) => {

        fs.writeFile(path, buffer, err => {
            
            if (err) {
                return reject(err);
            }
            return resolve();
        });
    });
}


/***/ },

/***/ "../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/index.js"
(module, __unused_webpack_exports, __webpack_require__) {

"use strict";


const BitmapImage = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/bitmapimage.js");
const { Gif, GifError } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gif.js");
const { GifCodec } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifcodec.js");
const { GifFrame } = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifframe.js");
const GifUtil = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/gifutil.js");

module.exports = {
    BitmapImage,
    Gif,
    GifCodec,
    GifFrame,
    GifUtil,
    GifError
};


/***/ },

/***/ "../../node_modules/.pnpm/omggif@1.0.10/node_modules/omggif/omggif.js"
(__unused_webpack_module, exports) {

"use strict";
// (c) Dean McNamee <dean@gmail.com>, 2013.
//
// https://github.com/deanm/omggif
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to
// deal in the Software without restriction, including without limitation the
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
// sell copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
// IN THE SOFTWARE.
//
// omggif is a JavaScript implementation of a GIF 89a encoder and decoder,
// including animation and compression.  It does not rely on any specific
// underlying system, so should run in the browser, Node, or Plask.



function GifWriter(buf, width, height, gopts) {
  var p = 0;

  var gopts = gopts === undefined ? { } : gopts;
  var loop_count = gopts.loop === undefined ? null : gopts.loop;
  var global_palette = gopts.palette === undefined ? null : gopts.palette;

  if (width <= 0 || height <= 0 || width > 65535 || height > 65535)
    throw new Error("Width/Height invalid.");

  function check_palette_and_num_colors(palette) {
    var num_colors = palette.length;
    if (num_colors < 2 || num_colors > 256 ||  num_colors & (num_colors-1)) {
      throw new Error(
          "Invalid code/color length, must be power of 2 and 2 .. 256.");
    }
    return num_colors;
  }

  // - Header.
  buf[p++] = 0x47; buf[p++] = 0x49; buf[p++] = 0x46;  // GIF
  buf[p++] = 0x38; buf[p++] = 0x39; buf[p++] = 0x61;  // 89a

  // Handling of Global Color Table (palette) and background index.
  var gp_num_colors_pow2 = 0;
  var background = 0;
  if (global_palette !== null) {
    var gp_num_colors = check_palette_and_num_colors(global_palette);
    while (gp_num_colors >>= 1) ++gp_num_colors_pow2;
    gp_num_colors = 1 << gp_num_colors_pow2;
    --gp_num_colors_pow2;
    if (gopts.background !== undefined) {
      background = gopts.background;
      if (background >= gp_num_colors)
        throw new Error("Background index out of range.");
      // The GIF spec states that a background index of 0 should be ignored, so
      // this is probably a mistake and you really want to set it to another
      // slot in the palette.  But actually in the end most browsers, etc end
      // up ignoring this almost completely (including for dispose background).
      if (background === 0)
        throw new Error("Background index explicitly passed as 0.");
    }
  }

  // - Logical Screen Descriptor.
  // NOTE(deanm): w/h apparently ignored by implementations, but set anyway.
  buf[p++] = width & 0xff; buf[p++] = width >> 8 & 0xff;
  buf[p++] = height & 0xff; buf[p++] = height >> 8 & 0xff;
  // NOTE: Indicates 0-bpp original color resolution (unused?).
  buf[p++] = (global_palette !== null ? 0x80 : 0) |  // Global Color Table Flag.
             gp_num_colors_pow2;  // NOTE: No sort flag (unused?).
  buf[p++] = background;  // Background Color Index.
  buf[p++] = 0;  // Pixel aspect ratio (unused?).

  // - Global Color Table
  if (global_palette !== null) {
    for (var i = 0, il = global_palette.length; i < il; ++i) {
      var rgb = global_palette[i];
      buf[p++] = rgb >> 16 & 0xff;
      buf[p++] = rgb >> 8 & 0xff;
      buf[p++] = rgb & 0xff;
    }
  }

  if (loop_count !== null) {  // Netscape block for looping.
    if (loop_count < 0 || loop_count > 65535)
      throw new Error("Loop count invalid.")
    // Extension code, label, and length.
    buf[p++] = 0x21; buf[p++] = 0xff; buf[p++] = 0x0b;
    // NETSCAPE2.0
    buf[p++] = 0x4e; buf[p++] = 0x45; buf[p++] = 0x54; buf[p++] = 0x53;
    buf[p++] = 0x43; buf[p++] = 0x41; buf[p++] = 0x50; buf[p++] = 0x45;
    buf[p++] = 0x32; buf[p++] = 0x2e; buf[p++] = 0x30;
    // Sub-block
    buf[p++] = 0x03; buf[p++] = 0x01;
    buf[p++] = loop_count & 0xff; buf[p++] = loop_count >> 8 & 0xff;
    buf[p++] = 0x00;  // Terminator.
  }


  var ended = false;

  this.addFrame = function(x, y, w, h, indexed_pixels, opts) {
    if (ended === true) { --p; ended = false; }  // Un-end.

    opts = opts === undefined ? { } : opts;

    // TODO(deanm): Bounds check x, y.  Do they need to be within the virtual
    // canvas width/height, I imagine?
    if (x < 0 || y < 0 || x > 65535 || y > 65535)
      throw new Error("x/y invalid.")

    if (w <= 0 || h <= 0 || w > 65535 || h > 65535)
      throw new Error("Width/Height invalid.")

    if (indexed_pixels.length < w * h)
      throw new Error("Not enough pixels for the frame size.");

    var using_local_palette = true;
    var palette = opts.palette;
    if (palette === undefined || palette === null) {
      using_local_palette = false;
      palette = global_palette;
    }

    if (palette === undefined || palette === null)
      throw new Error("Must supply either a local or global palette.");

    var num_colors = check_palette_and_num_colors(palette);

    // Compute the min_code_size (power of 2), destroying num_colors.
    var min_code_size = 0;
    while (num_colors >>= 1) ++min_code_size;
    num_colors = 1 << min_code_size;  // Now we can easily get it back.

    var delay = opts.delay === undefined ? 0 : opts.delay;

    // From the spec:
    //     0 -   No disposal specified. The decoder is
    //           not required to take any action.
    //     1 -   Do not dispose. The graphic is to be left
    //           in place.
    //     2 -   Restore to background color. The area used by the
    //           graphic must be restored to the background color.
    //     3 -   Restore to previous. The decoder is required to
    //           restore the area overwritten by the graphic with
    //           what was there prior to rendering the graphic.
    //  4-7 -    To be defined.
    // NOTE(deanm): Dispose background doesn't really work, apparently most
    // browsers ignore the background palette index and clear to transparency.
    var disposal = opts.disposal === undefined ? 0 : opts.disposal;
    if (disposal < 0 || disposal > 3)  // 4-7 is reserved.
      throw new Error("Disposal out of range.");

    var use_transparency = false;
    var transparent_index = 0;
    if (opts.transparent !== undefined && opts.transparent !== null) {
      use_transparency = true;
      transparent_index = opts.transparent;
      if (transparent_index < 0 || transparent_index >= num_colors)
        throw new Error("Transparent color index.");
    }

    if (disposal !== 0 || use_transparency || delay !== 0) {
      // - Graphics Control Extension
      buf[p++] = 0x21; buf[p++] = 0xf9;  // Extension / Label.
      buf[p++] = 4;  // Byte size.

      buf[p++] = disposal << 2 | (use_transparency === true ? 1 : 0);
      buf[p++] = delay & 0xff; buf[p++] = delay >> 8 & 0xff;
      buf[p++] = transparent_index;  // Transparent color index.
      buf[p++] = 0;  // Block Terminator.
    }

    // - Image Descriptor
    buf[p++] = 0x2c;  // Image Seperator.
    buf[p++] = x & 0xff; buf[p++] = x >> 8 & 0xff;  // Left.
    buf[p++] = y & 0xff; buf[p++] = y >> 8 & 0xff;  // Top.
    buf[p++] = w & 0xff; buf[p++] = w >> 8 & 0xff;
    buf[p++] = h & 0xff; buf[p++] = h >> 8 & 0xff;
    // NOTE: No sort flag (unused?).
    // TODO(deanm): Support interlace.
    buf[p++] = using_local_palette === true ? (0x80 | (min_code_size-1)) : 0;

    // - Local Color Table
    if (using_local_palette === true) {
      for (var i = 0, il = palette.length; i < il; ++i) {
        var rgb = palette[i];
        buf[p++] = rgb >> 16 & 0xff;
        buf[p++] = rgb >> 8 & 0xff;
        buf[p++] = rgb & 0xff;
      }
    }

    p = GifWriterOutputLZWCodeStream(
            buf, p, min_code_size < 2 ? 2 : min_code_size, indexed_pixels);

    return p;
  };

  this.end = function() {
    if (ended === false) {
      buf[p++] = 0x3b;  // Trailer.
      ended = true;
    }
    return p;
  };

  this.getOutputBuffer = function() { return buf; };
  this.setOutputBuffer = function(v) { buf = v; };
  this.getOutputBufferPosition = function() { return p; };
  this.setOutputBufferPosition = function(v) { p = v; };
}

// Main compression routine, palette indexes -> LZW code stream.
// |index_stream| must have at least one entry.
function GifWriterOutputLZWCodeStream(buf, p, min_code_size, index_stream) {
  buf[p++] = min_code_size;
  var cur_subblock = p++;  // Pointing at the length field.

  var clear_code = 1 << min_code_size;
  var code_mask = clear_code - 1;
  var eoi_code = clear_code + 1;
  var next_code = eoi_code + 1;

  var cur_code_size = min_code_size + 1;  // Number of bits per code.
  var cur_shift = 0;
  // We have at most 12-bit codes, so we should have to hold a max of 19
  // bits here (and then we would write out).
  var cur = 0;

  function emit_bytes_to_buffer(bit_block_size) {
    while (cur_shift >= bit_block_size) {
      buf[p++] = cur & 0xff;
      cur >>= 8; cur_shift -= 8;
      if (p === cur_subblock + 256) {  // Finished a subblock.
        buf[cur_subblock] = 255;
        cur_subblock = p++;
      }
    }
  }

  function emit_code(c) {
    cur |= c << cur_shift;
    cur_shift += cur_code_size;
    emit_bytes_to_buffer(8);
  }

  // I am not an expert on the topic, and I don't want to write a thesis.
  // However, it is good to outline here the basic algorithm and the few data
  // structures and optimizations here that make this implementation fast.
  // The basic idea behind LZW is to build a table of previously seen runs
  // addressed by a short id (herein called output code).  All data is
  // referenced by a code, which represents one or more values from the
  // original input stream.  All input bytes can be referenced as the same
  // value as an output code.  So if you didn't want any compression, you
  // could more or less just output the original bytes as codes (there are
  // some details to this, but it is the idea).  In order to achieve
  // compression, values greater then the input range (codes can be up to
  // 12-bit while input only 8-bit) represent a sequence of previously seen
  // inputs.  The decompressor is able to build the same mapping while
  // decoding, so there is always a shared common knowledge between the
  // encoding and decoder, which is also important for "timing" aspects like
  // how to handle variable bit width code encoding.
  //
  // One obvious but very important consequence of the table system is there
  // is always a unique id (at most 12-bits) to map the runs.  'A' might be
  // 4, then 'AA' might be 10, 'AAA' 11, 'AAAA' 12, etc.  This relationship
  // can be used for an effecient lookup strategy for the code mapping.  We
  // need to know if a run has been seen before, and be able to map that run
  // to the output code.  Since we start with known unique ids (input bytes),
  // and then from those build more unique ids (table entries), we can
  // continue this chain (almost like a linked list) to always have small
  // integer values that represent the current byte chains in the encoder.
  // This means instead of tracking the input bytes (AAAABCD) to know our
  // current state, we can track the table entry for AAAABC (it is guaranteed
  // to exist by the nature of the algorithm) and the next character D.
  // Therefor the tuple of (table_entry, byte) is guaranteed to also be
  // unique.  This allows us to create a simple lookup key for mapping input
  // sequences to codes (table indices) without having to store or search
  // any of the code sequences.  So if 'AAAA' has a table entry of 12, the
  // tuple of ('AAAA', K) for any input byte K will be unique, and can be our
  // key.  This leads to a integer value at most 20-bits, which can always
  // fit in an SMI value and be used as a fast sparse array / object key.

  // Output code for the current contents of the index buffer.
  var ib_code = index_stream[0] & code_mask;  // Load first input index.
  var code_table = { };  // Key'd on our 20-bit "tuple".

  emit_code(clear_code);  // Spec says first code should be a clear code.

  // First index already loaded, process the rest of the stream.
  for (var i = 1, il = index_stream.length; i < il; ++i) {
    var k = index_stream[i] & code_mask;
    var cur_key = ib_code << 8 | k;  // (prev, k) unique tuple.
    var cur_code = code_table[cur_key];  // buffer + k.

    // Check if we have to create a new code table entry.
    if (cur_code === undefined) {  // We don't have buffer + k.
      // Emit index buffer (without k).
      // This is an inline version of emit_code, because this is the core
      // writing routine of the compressor (and V8 cannot inline emit_code
      // because it is a closure here in a different context).  Additionally
      // we can call emit_byte_to_buffer less often, because we can have
      // 30-bits (from our 31-bit signed SMI), and we know our codes will only
      // be 12-bits, so can safely have 18-bits there without overflow.
      // emit_code(ib_code);
      cur |= ib_code << cur_shift;
      cur_shift += cur_code_size;
      while (cur_shift >= 8) {
        buf[p++] = cur & 0xff;
        cur >>= 8; cur_shift -= 8;
        if (p === cur_subblock + 256) {  // Finished a subblock.
          buf[cur_subblock] = 255;
          cur_subblock = p++;
        }
      }

      if (next_code === 4096) {  // Table full, need a clear.
        emit_code(clear_code);
        next_code = eoi_code + 1;
        cur_code_size = min_code_size + 1;
        code_table = { };
      } else {  // Table not full, insert a new entry.
        // Increase our variable bit code sizes if necessary.  This is a bit
        // tricky as it is based on "timing" between the encoding and
        // decoder.  From the encoders perspective this should happen after
        // we've already emitted the index buffer and are about to create the
        // first table entry that would overflow our current code bit size.
        if (next_code >= (1 << cur_code_size)) ++cur_code_size;
        code_table[cur_key] = next_code++;  // Insert into code table.
      }

      ib_code = k;  // Index buffer to single input k.
    } else {
      ib_code = cur_code;  // Index buffer to sequence in code table.
    }
  }

  emit_code(ib_code);  // There will still be something in the index buffer.
  emit_code(eoi_code);  // End Of Information.

  // Flush / finalize the sub-blocks stream to the buffer.
  emit_bytes_to_buffer(1);

  // Finish the sub-blocks, writing out any unfinished lengths and
  // terminating with a sub-block of length 0.  If we have already started
  // but not yet used a sub-block it can just become the terminator.
  if (cur_subblock + 1 === p) {  // Started but unused.
    buf[cur_subblock] = 0;
  } else {  // Started and used, write length and additional terminator block.
    buf[cur_subblock] = p - cur_subblock - 1;
    buf[p++] = 0;
  }
  return p;
}

function GifReader(buf) {
  var p = 0;

  // - Header (GIF87a or GIF89a).
  if (buf[p++] !== 0x47 ||            buf[p++] !== 0x49 || buf[p++] !== 0x46 ||
      buf[p++] !== 0x38 || (buf[p++]+1 & 0xfd) !== 0x38 || buf[p++] !== 0x61) {
    throw new Error("Invalid GIF 87a/89a header.");
  }

  // - Logical Screen Descriptor.
  var width = buf[p++] | buf[p++] << 8;
  var height = buf[p++] | buf[p++] << 8;
  var pf0 = buf[p++];  // <Packed Fields>.
  var global_palette_flag = pf0 >> 7;
  var num_global_colors_pow2 = pf0 & 0x7;
  var num_global_colors = 1 << (num_global_colors_pow2 + 1);
  var background = buf[p++];
  buf[p++];  // Pixel aspect ratio (unused?).

  var global_palette_offset = null;
  var global_palette_size   = null;

  if (global_palette_flag) {
    global_palette_offset = p;
    global_palette_size = num_global_colors;
    p += num_global_colors * 3;  // Seek past palette.
  }

  var no_eof = true;

  var frames = [ ];

  var delay = 0;
  var transparent_index = null;
  var disposal = 0;  // 0 - No disposal specified.
  var loop_count = null;

  this.width = width;
  this.height = height;

  while (no_eof && p < buf.length) {
    switch (buf[p++]) {
      case 0x21:  // Graphics Control Extension Block
        switch (buf[p++]) {
          case 0xff:  // Application specific block
            // Try if it's a Netscape block (with animation loop counter).
            if (buf[p   ] !== 0x0b ||  // 21 FF already read, check block size.
                // NETSCAPE2.0
                buf[p+1 ] == 0x4e && buf[p+2 ] == 0x45 && buf[p+3 ] == 0x54 &&
                buf[p+4 ] == 0x53 && buf[p+5 ] == 0x43 && buf[p+6 ] == 0x41 &&
                buf[p+7 ] == 0x50 && buf[p+8 ] == 0x45 && buf[p+9 ] == 0x32 &&
                buf[p+10] == 0x2e && buf[p+11] == 0x30 &&
                // Sub-block
                buf[p+12] == 0x03 && buf[p+13] == 0x01 && buf[p+16] == 0) {
              p += 14;
              loop_count = buf[p++] | buf[p++] << 8;
              p++;  // Skip terminator.
            } else {  // We don't know what it is, just try to get past it.
              p += 12;
              while (true) {  // Seek through subblocks.
                var block_size = buf[p++];
                // Bad block size (ex: undefined from an out of bounds read).
                if (!(block_size >= 0)) throw Error("Invalid block size");
                if (block_size === 0) break;  // 0 size is terminator
                p += block_size;
              }
            }
            break;

          case 0xf9:  // Graphics Control Extension
            if (buf[p++] !== 0x4 || buf[p+4] !== 0)
              throw new Error("Invalid graphics extension block.");
            var pf1 = buf[p++];
            delay = buf[p++] | buf[p++] << 8;
            transparent_index = buf[p++];
            if ((pf1 & 1) === 0) transparent_index = null;
            disposal = pf1 >> 2 & 0x7;
            p++;  // Skip terminator.
            break;

          case 0xfe:  // Comment Extension.
            while (true) {  // Seek through subblocks.
              var block_size = buf[p++];
              // Bad block size (ex: undefined from an out of bounds read).
              if (!(block_size >= 0)) throw Error("Invalid block size");
              if (block_size === 0) break;  // 0 size is terminator
              // console.log(buf.slice(p, p+block_size).toString('ascii'));
              p += block_size;
            }
            break;

          default:
            throw new Error(
                "Unknown graphic control label: 0x" + buf[p-1].toString(16));
        }
        break;

      case 0x2c:  // Image Descriptor.
        var x = buf[p++] | buf[p++] << 8;
        var y = buf[p++] | buf[p++] << 8;
        var w = buf[p++] | buf[p++] << 8;
        var h = buf[p++] | buf[p++] << 8;
        var pf2 = buf[p++];
        var local_palette_flag = pf2 >> 7;
        var interlace_flag = pf2 >> 6 & 1;
        var num_local_colors_pow2 = pf2 & 0x7;
        var num_local_colors = 1 << (num_local_colors_pow2 + 1);
        var palette_offset = global_palette_offset;
        var palette_size = global_palette_size;
        var has_local_palette = false;
        if (local_palette_flag) {
          var has_local_palette = true;
          palette_offset = p;  // Override with local palette.
          palette_size = num_local_colors;
          p += num_local_colors * 3;  // Seek past palette.
        }

        var data_offset = p;

        p++;  // codesize
        while (true) {
          var block_size = buf[p++];
          // Bad block size (ex: undefined from an out of bounds read).
          if (!(block_size >= 0)) throw Error("Invalid block size");
          if (block_size === 0) break;  // 0 size is terminator
          p += block_size;
        }

        frames.push({x: x, y: y, width: w, height: h,
                     has_local_palette: has_local_palette,
                     palette_offset: palette_offset,
                     palette_size: palette_size,
                     data_offset: data_offset,
                     data_length: p - data_offset,
                     transparent_index: transparent_index,
                     interlaced: !!interlace_flag,
                     delay: delay,
                     disposal: disposal});
        break;

      case 0x3b:  // Trailer Marker (end of file).
        no_eof = false;
        break;

      default:
        throw new Error("Unknown gif block: 0x" + buf[p-1].toString(16));
        // removed by dead control flow

    }
  }

  this.numFrames = function() {
    return frames.length;
  };

  this.loopCount = function() {
    return loop_count;
  };

  this.frameInfo = function(frame_num) {
    if (frame_num < 0 || frame_num >= frames.length)
      throw new Error("Frame index out of range.");
    return frames[frame_num];
  }

  this.decodeAndBlitFrameBGRA = function(frame_num, pixels) {
    var frame = this.frameInfo(frame_num);
    var num_pixels = frame.width * frame.height;
    var index_stream = new Uint8Array(num_pixels);  // At most 8-bit indices.
    GifReaderLZWOutputIndexStream(
        buf, frame.data_offset, index_stream, num_pixels);
    var palette_offset = frame.palette_offset;

    // NOTE(deanm): It seems to be much faster to compare index to 256 than
    // to === null.  Not sure why, but CompareStub_EQ_STRICT shows up high in
    // the profile, not sure if it's related to using a Uint8Array.
    var trans = frame.transparent_index;
    if (trans === null) trans = 256;

    // We are possibly just blitting to a portion of the entire frame.
    // That is a subrect within the framerect, so the additional pixels
    // must be skipped over after we finished a scanline.
    var framewidth  = frame.width;
    var framestride = width - framewidth;
    var xleft       = framewidth;  // Number of subrect pixels left in scanline.

    // Output indicies of the top left and bottom right corners of the subrect.
    var opbeg = ((frame.y * width) + frame.x) * 4;
    var opend = ((frame.y + frame.height) * width + frame.x) * 4;
    var op    = opbeg;

    var scanstride = framestride * 4;

    // Use scanstride to skip past the rows when interlacing.  This is skipping
    // 7 rows for the first two passes, then 3 then 1.
    if (frame.interlaced === true) {
      scanstride += width * 4 * 7;  // Pass 1.
    }

    var interlaceskip = 8;  // Tracking the row interval in the current pass.

    for (var i = 0, il = index_stream.length; i < il; ++i) {
      var index = index_stream[i];

      if (xleft === 0) {  // Beginning of new scan line
        op += scanstride;
        xleft = framewidth;
        if (op >= opend) { // Catch the wrap to switch passes when interlacing.
          scanstride = framestride * 4 + width * 4 * (interlaceskip-1);
          // interlaceskip / 2 * 4 is interlaceskip << 1.
          op = opbeg + (framewidth + framestride) * (interlaceskip << 1);
          interlaceskip >>= 1;
        }
      }

      if (index === trans) {
        op += 4;
      } else {
        var r = buf[palette_offset + index * 3];
        var g = buf[palette_offset + index * 3 + 1];
        var b = buf[palette_offset + index * 3 + 2];
        pixels[op++] = b;
        pixels[op++] = g;
        pixels[op++] = r;
        pixels[op++] = 255;
      }
      --xleft;
    }
  };

  // I will go to copy and paste hell one day...
  this.decodeAndBlitFrameRGBA = function(frame_num, pixels) {
    var frame = this.frameInfo(frame_num);
    var num_pixels = frame.width * frame.height;
    var index_stream = new Uint8Array(num_pixels);  // At most 8-bit indices.
    GifReaderLZWOutputIndexStream(
        buf, frame.data_offset, index_stream, num_pixels);
    var palette_offset = frame.palette_offset;

    // NOTE(deanm): It seems to be much faster to compare index to 256 than
    // to === null.  Not sure why, but CompareStub_EQ_STRICT shows up high in
    // the profile, not sure if it's related to using a Uint8Array.
    var trans = frame.transparent_index;
    if (trans === null) trans = 256;

    // We are possibly just blitting to a portion of the entire frame.
    // That is a subrect within the framerect, so the additional pixels
    // must be skipped over after we finished a scanline.
    var framewidth  = frame.width;
    var framestride = width - framewidth;
    var xleft       = framewidth;  // Number of subrect pixels left in scanline.

    // Output indicies of the top left and bottom right corners of the subrect.
    var opbeg = ((frame.y * width) + frame.x) * 4;
    var opend = ((frame.y + frame.height) * width + frame.x) * 4;
    var op    = opbeg;

    var scanstride = framestride * 4;

    // Use scanstride to skip past the rows when interlacing.  This is skipping
    // 7 rows for the first two passes, then 3 then 1.
    if (frame.interlaced === true) {
      scanstride += width * 4 * 7;  // Pass 1.
    }

    var interlaceskip = 8;  // Tracking the row interval in the current pass.

    for (var i = 0, il = index_stream.length; i < il; ++i) {
      var index = index_stream[i];

      if (xleft === 0) {  // Beginning of new scan line
        op += scanstride;
        xleft = framewidth;
        if (op >= opend) { // Catch the wrap to switch passes when interlacing.
          scanstride = framestride * 4 + width * 4 * (interlaceskip-1);
          // interlaceskip / 2 * 4 is interlaceskip << 1.
          op = opbeg + (framewidth + framestride) * (interlaceskip << 1);
          interlaceskip >>= 1;
        }
      }

      if (index === trans) {
        op += 4;
      } else {
        var r = buf[palette_offset + index * 3];
        var g = buf[palette_offset + index * 3 + 1];
        var b = buf[palette_offset + index * 3 + 2];
        pixels[op++] = r;
        pixels[op++] = g;
        pixels[op++] = b;
        pixels[op++] = 255;
      }
      --xleft;
    }
  };
}

function GifReaderLZWOutputIndexStream(code_stream, p, output, output_length) {
  var min_code_size = code_stream[p++];

  var clear_code = 1 << min_code_size;
  var eoi_code = clear_code + 1;
  var next_code = eoi_code + 1;

  var cur_code_size = min_code_size + 1;  // Number of bits per code.
  // NOTE: This shares the same name as the encoder, but has a different
  // meaning here.  Here this masks each code coming from the code stream.
  var code_mask = (1 << cur_code_size) - 1;
  var cur_shift = 0;
  var cur = 0;

  var op = 0;  // Output pointer.

  var subblock_size = code_stream[p++];

  // TODO(deanm): Would using a TypedArray be any faster?  At least it would
  // solve the fast mode / backing store uncertainty.
  // var code_table = Array(4096);
  var code_table = new Int32Array(4096);  // Can be signed, we only use 20 bits.

  var prev_code = null;  // Track code-1.

  while (true) {
    // Read up to two bytes, making sure we always 12-bits for max sized code.
    while (cur_shift < 16) {
      if (subblock_size === 0) break;  // No more data to be read.

      cur |= code_stream[p++] << cur_shift;
      cur_shift += 8;

      if (subblock_size === 1) {  // Never let it get to 0 to hold logic above.
        subblock_size = code_stream[p++];  // Next subblock.
      } else {
        --subblock_size;
      }
    }

    // TODO(deanm): We should never really get here, we should have received
    // and EOI.
    if (cur_shift < cur_code_size)
      break;

    var code = cur & code_mask;
    cur >>= cur_code_size;
    cur_shift -= cur_code_size;

    // TODO(deanm): Maybe should check that the first code was a clear code,
    // at least this is what you're supposed to do.  But actually our encoder
    // now doesn't emit a clear code first anyway.
    if (code === clear_code) {
      // We don't actually have to clear the table.  This could be a good idea
      // for greater error checking, but we don't really do any anyway.  We
      // will just track it with next_code and overwrite old entries.

      next_code = eoi_code + 1;
      cur_code_size = min_code_size + 1;
      code_mask = (1 << cur_code_size) - 1;

      // Don't update prev_code ?
      prev_code = null;
      continue;
    } else if (code === eoi_code) {
      break;
    }

    // We have a similar situation as the decoder, where we want to store
    // variable length entries (code table entries), but we want to do in a
    // faster manner than an array of arrays.  The code below stores sort of a
    // linked list within the code table, and then "chases" through it to
    // construct the dictionary entries.  When a new entry is created, just the
    // last byte is stored, and the rest (prefix) of the entry is only
    // referenced by its table entry.  Then the code chases through the
    // prefixes until it reaches a single byte code.  We have to chase twice,
    // first to compute the length, and then to actually copy the data to the
    // output (backwards, since we know the length).  The alternative would be
    // storing something in an intermediate stack, but that doesn't make any
    // more sense.  I implemented an approach where it also stored the length
    // in the code table, although it's a bit tricky because you run out of
    // bits (12 + 12 + 8), but I didn't measure much improvements (the table
    // entries are generally not the long).  Even when I created benchmarks for
    // very long table entries the complexity did not seem worth it.
    // The code table stores the prefix entry in 12 bits and then the suffix
    // byte in 8 bits, so each entry is 20 bits.

    var chase_code = code < next_code ? code : prev_code;

    // Chase what we will output, either {CODE} or {CODE-1}.
    var chase_length = 0;
    var chase = chase_code;
    while (chase > clear_code) {
      chase = code_table[chase] >> 8;
      ++chase_length;
    }

    var k = chase;

    var op_end = op + chase_length + (chase_code !== code ? 1 : 0);
    if (op_end > output_length) {
      console.log("Warning, gif stream longer than expected.");
      return;
    }

    // Already have the first byte from the chase, might as well write it fast.
    output[op++] = k;

    op += chase_length;
    var b = op;  // Track pointer, writing backwards.

    if (chase_code !== code)  // The case of emitting {CODE-1} + k.
      output[op++] = k;

    chase = chase_code;
    while (chase_length--) {
      chase = code_table[chase];
      output[--b] = chase & 0xff;  // Write backwards.
      chase >>= 8;  // Pull down to the prefix code.
    }

    if (prev_code !== null && next_code < 4096) {
      code_table[next_code++] = prev_code << 8 | k;
      // TODO(deanm): Figure out this clearing vs code growth logic better.  I
      // have an feeling that it should just happen somewhere else, for now it
      // is awkward between when we grow past the max and then hit a clear code.
      // For now just check if we hit the max 12-bits (then a clear code should
      // follow, also of course encoded in 12-bits).
      if (next_code >= code_mask+1 && cur_code_size < 12) {
        ++cur_code_size;
        code_mask = code_mask << 1 | 1;
      }
    }

    prev_code = code;
  }

  if (op !== output_length) {
    console.log("Warning, gif stream shorter than expected.");
  }

  return output;
}

// CommonJS.
try { exports.GifWriter = GifWriter; exports.GifReader = GifReader } catch(e) {}


/***/ },

/***/ "../../node_modules/.pnpm/image-q@4.0.0/node_modules/image-q/dist/cjs/image-q.cjs"
(module) {

var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __markAsModule = (target) => __defProp(target, "__esModule", { value: true });
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __reExport = (target, module2, copyDefault, desc) => {
  if (module2 && typeof module2 === "object" || typeof module2 === "function") {
    for (let key of __getOwnPropNames(module2))
      if (!__hasOwnProp.call(target, key) && (copyDefault || key !== "default"))
        __defProp(target, key, { get: () => module2[key], enumerable: !(desc = __getOwnPropDesc(module2, key)) || desc.enumerable });
  }
  return target;
};
var __toCommonJS = /* @__PURE__ */ ((cache) => {
  return (module2, temp) => {
    return cache && cache.get(module2) || (temp = __reExport(__markAsModule({}), module2, 1), cache && cache.set(module2, temp), temp);
  };
})(typeof WeakMap !== "undefined" ? /* @__PURE__ */ new WeakMap() : 0);
var __publicField = (obj, key, value) => {
  __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);
  return value;
};

// src/index.ts
var src_exports = {};
__export(src_exports, {
  applyPalette: () => applyPalette,
  applyPaletteSync: () => applyPaletteSync,
  buildPalette: () => buildPalette,
  buildPaletteSync: () => buildPaletteSync,
  constants: () => constants_exports,
  conversion: () => conversion_exports,
  distance: () => distance_exports,
  image: () => image_exports,
  palette: () => palette_exports,
  quality: () => quality_exports,
  utils: () => utils_exports
});

// src/constants/index.ts
var constants_exports = {};
__export(constants_exports, {
  bt709: () => bt709_exports
});

// src/constants/bt709.ts
var bt709_exports = {};
__export(bt709_exports, {
  Y: () => Y,
  x: () => x,
  y: () => y
});
var Y = /* @__PURE__ */ ((Y2) => {
  Y2[Y2["RED"] = 0.2126] = "RED";
  Y2[Y2["GREEN"] = 0.7152] = "GREEN";
  Y2[Y2["BLUE"] = 0.0722] = "BLUE";
  Y2[Y2["WHITE"] = 1] = "WHITE";
  return Y2;
})(Y || {});
var x = /* @__PURE__ */ ((x2) => {
  x2[x2["RED"] = 0.64] = "RED";
  x2[x2["GREEN"] = 0.3] = "GREEN";
  x2[x2["BLUE"] = 0.15] = "BLUE";
  x2[x2["WHITE"] = 0.3127] = "WHITE";
  return x2;
})(x || {});
var y = /* @__PURE__ */ ((y2) => {
  y2[y2["RED"] = 0.33] = "RED";
  y2[y2["GREEN"] = 0.6] = "GREEN";
  y2[y2["BLUE"] = 0.06] = "BLUE";
  y2[y2["WHITE"] = 0.329] = "WHITE";
  return y2;
})(y || {});

// src/conversion/index.ts
var conversion_exports = {};
__export(conversion_exports, {
  lab2rgb: () => lab2rgb,
  lab2xyz: () => lab2xyz,
  rgb2hsl: () => rgb2hsl,
  rgb2lab: () => rgb2lab,
  rgb2xyz: () => rgb2xyz,
  xyz2lab: () => xyz2lab,
  xyz2rgb: () => xyz2rgb
});

// src/conversion/rgb2xyz.ts
function correctGamma(n) {
  return n > 0.04045 ? ((n + 0.055) / 1.055) ** 2.4 : n / 12.92;
}
function rgb2xyz(r, g, b) {
  r = correctGamma(r / 255);
  g = correctGamma(g / 255);
  b = correctGamma(b / 255);
  return {
    x: r * 0.4124 + g * 0.3576 + b * 0.1805,
    y: r * 0.2126 + g * 0.7152 + b * 0.0722,
    z: r * 0.0193 + g * 0.1192 + b * 0.9505
  };
}

// src/utils/arithmetic.ts
var arithmetic_exports = {};
__export(arithmetic_exports, {
  degrees2radians: () => degrees2radians,
  inRange0to255: () => inRange0to255,
  inRange0to255Rounded: () => inRange0to255Rounded,
  intInRange: () => intInRange,
  max3: () => max3,
  min3: () => min3,
  stableSort: () => stableSort
});
function degrees2radians(n) {
  return n * (Math.PI / 180);
}
function max3(a, b, c) {
  let m = a;
  if (m < b)
    m = b;
  if (m < c)
    m = c;
  return m;
}
function min3(a, b, c) {
  let m = a;
  if (m > b)
    m = b;
  if (m > c)
    m = c;
  return m;
}
function intInRange(value, low, high) {
  if (value > high)
    value = high;
  if (value < low)
    value = low;
  return value | 0;
}
function inRange0to255Rounded(n) {
  n = Math.round(n);
  if (n > 255)
    n = 255;
  else if (n < 0)
    n = 0;
  return n;
}
function inRange0to255(n) {
  if (n > 255)
    n = 255;
  else if (n < 0)
    n = 0;
  return n;
}
function stableSort(arrayToSort, callback) {
  const type = typeof arrayToSort[0];
  let sorted;
  if (type === "number" || type === "string") {
    const ord = /* @__PURE__ */ Object.create(null);
    for (let i = 0, l = arrayToSort.length; i < l; i++) {
      const val = arrayToSort[i];
      if (ord[val] || ord[val] === 0)
        continue;
      ord[val] = i;
    }
    sorted = arrayToSort.sort((a, b) => callback(a, b) || ord[a] - ord[b]);
  } else {
    const ord2 = arrayToSort.slice(0);
    sorted = arrayToSort.sort((a, b) => callback(a, b) || ord2.indexOf(a) - ord2.indexOf(b));
  }
  return sorted;
}

// src/conversion/rgb2hsl.ts
function rgb2hsl(r, g, b) {
  const min = min3(r, g, b);
  const max = max3(r, g, b);
  const delta = max - min;
  const l = (min + max) / 510;
  let s = 0;
  if (l > 0 && l < 1)
    s = delta / (l < 0.5 ? max + min : 510 - max - min);
  let h = 0;
  if (delta > 0) {
    if (max === r) {
      h = (g - b) / delta;
    } else if (max === g) {
      h = 2 + (b - r) / delta;
    } else {
      h = 4 + (r - g) / delta;
    }
    h *= 60;
    if (h < 0)
      h += 360;
  }
  return { h, s, l };
}

// src/conversion/xyz2lab.ts
var refX = 0.95047;
var refY = 1;
var refZ = 1.08883;
function pivot(n) {
  return n > 8856e-6 ? n ** (1 / 3) : 7.787 * n + 16 / 116;
}
function xyz2lab(x2, y2, z) {
  x2 = pivot(x2 / refX);
  y2 = pivot(y2 / refY);
  z = pivot(z / refZ);
  if (116 * y2 - 16 < 0)
    throw new Error("xxx");
  return {
    L: Math.max(0, 116 * y2 - 16),
    a: 500 * (x2 - y2),
    b: 200 * (y2 - z)
  };
}

// src/conversion/rgb2lab.ts
function rgb2lab(r, g, b) {
  const xyz = rgb2xyz(r, g, b);
  return xyz2lab(xyz.x, xyz.y, xyz.z);
}

// src/conversion/lab2xyz.ts
var refX2 = 0.95047;
var refY2 = 1;
var refZ2 = 1.08883;
function pivot2(n) {
  return n > 0.206893034 ? n ** 3 : (n - 16 / 116) / 7.787;
}
function lab2xyz(L, a, b) {
  const y2 = (L + 16) / 116;
  const x2 = a / 500 + y2;
  const z = y2 - b / 200;
  return {
    x: refX2 * pivot2(x2),
    y: refY2 * pivot2(y2),
    z: refZ2 * pivot2(z)
  };
}

// src/conversion/xyz2rgb.ts
function correctGamma2(n) {
  return n > 31308e-7 ? 1.055 * n ** (1 / 2.4) - 0.055 : 12.92 * n;
}
function xyz2rgb(x2, y2, z) {
  const r = correctGamma2(x2 * 3.2406 + y2 * -1.5372 + z * -0.4986);
  const g = correctGamma2(x2 * -0.9689 + y2 * 1.8758 + z * 0.0415);
  const b = correctGamma2(x2 * 0.0557 + y2 * -0.204 + z * 1.057);
  return {
    r: inRange0to255Rounded(r * 255),
    g: inRange0to255Rounded(g * 255),
    b: inRange0to255Rounded(b * 255)
  };
}

// src/conversion/lab2rgb.ts
function lab2rgb(L, a, b) {
  const xyz = lab2xyz(L, a, b);
  return xyz2rgb(xyz.x, xyz.y, xyz.z);
}

// src/distance/index.ts
var distance_exports = {};
__export(distance_exports, {
  AbstractDistanceCalculator: () => AbstractDistanceCalculator,
  AbstractEuclidean: () => AbstractEuclidean,
  AbstractManhattan: () => AbstractManhattan,
  CIE94GraphicArts: () => CIE94GraphicArts,
  CIE94Textiles: () => CIE94Textiles,
  CIEDE2000: () => CIEDE2000,
  CMetric: () => CMetric,
  Euclidean: () => Euclidean,
  EuclideanBT709: () => EuclideanBT709,
  EuclideanBT709NoAlpha: () => EuclideanBT709NoAlpha,
  Manhattan: () => Manhattan,
  ManhattanBT709: () => ManhattanBT709,
  ManhattanNommyde: () => ManhattanNommyde,
  PNGQuant: () => PNGQuant
});

// src/distance/distanceCalculator.ts
var AbstractDistanceCalculator = class {
  constructor() {
    __publicField(this, "_maxDistance");
    __publicField(this, "_whitePoint");
    this._setDefaults();
    this.setWhitePoint(255, 255, 255, 255);
  }
  setWhitePoint(r, g, b, a) {
    this._whitePoint = {
      r: r > 0 ? 255 / r : 0,
      g: g > 0 ? 255 / g : 0,
      b: b > 0 ? 255 / b : 0,
      a: a > 0 ? 255 / a : 0
    };
    this._maxDistance = this.calculateRaw(r, g, b, a, 0, 0, 0, 0);
  }
  calculateNormalized(colorA, colorB) {
    return this.calculateRaw(colorA.r, colorA.g, colorA.b, colorA.a, colorB.r, colorB.g, colorB.b, colorB.a) / this._maxDistance;
  }
};

// src/distance/cie94.ts
var AbstractCIE94 = class extends AbstractDistanceCalculator {
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    const lab1 = rgb2lab(inRange0to255(r1 * this._whitePoint.r), inRange0to255(g1 * this._whitePoint.g), inRange0to255(b1 * this._whitePoint.b));
    const lab2 = rgb2lab(inRange0to255(r2 * this._whitePoint.r), inRange0to255(g2 * this._whitePoint.g), inRange0to255(b2 * this._whitePoint.b));
    const dL = lab1.L - lab2.L;
    const dA = lab1.a - lab2.a;
    const dB = lab1.b - lab2.b;
    const c1 = Math.sqrt(lab1.a * lab1.a + lab1.b * lab1.b);
    const c2 = Math.sqrt(lab2.a * lab2.a + lab2.b * lab2.b);
    const dC = c1 - c2;
    let deltaH = dA * dA + dB * dB - dC * dC;
    deltaH = deltaH < 0 ? 0 : Math.sqrt(deltaH);
    const dAlpha = (a2 - a1) * this._whitePoint.a * this._kA;
    return Math.sqrt((dL / this._Kl) ** 2 + (dC / (1 + this._K1 * c1)) ** 2 + (deltaH / (1 + this._K2 * c1)) ** 2 + dAlpha ** 2);
  }
};
var CIE94Textiles = class extends AbstractCIE94 {
  _setDefaults() {
    this._Kl = 2;
    this._K1 = 0.048;
    this._K2 = 0.014;
    this._kA = 0.25 * 50 / 255;
  }
};
var CIE94GraphicArts = class extends AbstractCIE94 {
  _setDefaults() {
    this._Kl = 1;
    this._K1 = 0.045;
    this._K2 = 0.015;
    this._kA = 0.25 * 100 / 255;
  }
};

// src/distance/ciede2000.ts
var _CIEDE2000 = class extends AbstractDistanceCalculator {
  _setDefaults() {
  }
  static _calculatehp(b, ap) {
    const hp = Math.atan2(b, ap);
    if (hp >= 0)
      return hp;
    return hp + _CIEDE2000._deg360InRad;
  }
  static _calculateRT(ahp, aCp) {
    const aCp_to_7 = aCp ** 7;
    const R_C = 2 * Math.sqrt(aCp_to_7 / (aCp_to_7 + _CIEDE2000._pow25to7));
    const delta_theta = _CIEDE2000._deg30InRad * Math.exp(-(((ahp - _CIEDE2000._deg275InRad) / _CIEDE2000._deg25InRad) ** 2));
    return -Math.sin(2 * delta_theta) * R_C;
  }
  static _calculateT(ahp) {
    return 1 - 0.17 * Math.cos(ahp - _CIEDE2000._deg30InRad) + 0.24 * Math.cos(ahp * 2) + 0.32 * Math.cos(ahp * 3 + _CIEDE2000._deg6InRad) - 0.2 * Math.cos(ahp * 4 - _CIEDE2000._deg63InRad);
  }
  static _calculate_ahp(C1pC2p, h_bar, h1p, h2p) {
    const hpSum = h1p + h2p;
    if (C1pC2p === 0)
      return hpSum;
    if (h_bar <= _CIEDE2000._deg180InRad)
      return hpSum / 2;
    if (hpSum < _CIEDE2000._deg360InRad) {
      return (hpSum + _CIEDE2000._deg360InRad) / 2;
    }
    return (hpSum - _CIEDE2000._deg360InRad) / 2;
  }
  static _calculate_dHp(C1pC2p, h_bar, h2p, h1p) {
    let dhp;
    if (C1pC2p === 0) {
      dhp = 0;
    } else if (h_bar <= _CIEDE2000._deg180InRad) {
      dhp = h2p - h1p;
    } else if (h2p <= h1p) {
      dhp = h2p - h1p + _CIEDE2000._deg360InRad;
    } else {
      dhp = h2p - h1p - _CIEDE2000._deg360InRad;
    }
    return 2 * Math.sqrt(C1pC2p) * Math.sin(dhp / 2);
  }
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    const lab1 = rgb2lab(inRange0to255(r1 * this._whitePoint.r), inRange0to255(g1 * this._whitePoint.g), inRange0to255(b1 * this._whitePoint.b));
    const lab2 = rgb2lab(inRange0to255(r2 * this._whitePoint.r), inRange0to255(g2 * this._whitePoint.g), inRange0to255(b2 * this._whitePoint.b));
    const dA = (a2 - a1) * this._whitePoint.a * _CIEDE2000._kA;
    const dE2 = this.calculateRawInLab(lab1, lab2);
    return Math.sqrt(dE2 + dA * dA);
  }
  calculateRawInLab(Lab1, Lab2) {
    const L1 = Lab1.L;
    const a1 = Lab1.a;
    const b1 = Lab1.b;
    const L2 = Lab2.L;
    const a2 = Lab2.a;
    const b2 = Lab2.b;
    const C1 = Math.sqrt(a1 * a1 + b1 * b1);
    const C2 = Math.sqrt(a2 * a2 + b2 * b2);
    const pow_a_C1_C2_to_7 = ((C1 + C2) / 2) ** 7;
    const G = 0.5 * (1 - Math.sqrt(pow_a_C1_C2_to_7 / (pow_a_C1_C2_to_7 + _CIEDE2000._pow25to7)));
    const a1p = (1 + G) * a1;
    const a2p = (1 + G) * a2;
    const C1p = Math.sqrt(a1p * a1p + b1 * b1);
    const C2p = Math.sqrt(a2p * a2p + b2 * b2);
    const C1pC2p = C1p * C2p;
    const h1p = _CIEDE2000._calculatehp(b1, a1p);
    const h2p = _CIEDE2000._calculatehp(b2, a2p);
    const h_bar = Math.abs(h1p - h2p);
    const dLp = L2 - L1;
    const dCp = C2p - C1p;
    const dHp = _CIEDE2000._calculate_dHp(C1pC2p, h_bar, h2p, h1p);
    const ahp = _CIEDE2000._calculate_ahp(C1pC2p, h_bar, h1p, h2p);
    const T = _CIEDE2000._calculateT(ahp);
    const aCp = (C1p + C2p) / 2;
    const aLp_minus_50_square = ((L1 + L2) / 2 - 50) ** 2;
    const S_L = 1 + 0.015 * aLp_minus_50_square / Math.sqrt(20 + aLp_minus_50_square);
    const S_C = 1 + 0.045 * aCp;
    const S_H = 1 + 0.015 * T * aCp;
    const R_T = _CIEDE2000._calculateRT(ahp, aCp);
    const dLpSL = dLp / S_L;
    const dCpSC = dCp / S_C;
    const dHpSH = dHp / S_H;
    return dLpSL ** 2 + dCpSC ** 2 + dHpSH ** 2 + R_T * dCpSC * dHpSH;
  }
};
var CIEDE2000 = _CIEDE2000;
__publicField(CIEDE2000, "_kA", 0.25 * 100 / 255);
__publicField(CIEDE2000, "_pow25to7", 25 ** 7);
__publicField(CIEDE2000, "_deg360InRad", degrees2radians(360));
__publicField(CIEDE2000, "_deg180InRad", degrees2radians(180));
__publicField(CIEDE2000, "_deg30InRad", degrees2radians(30));
__publicField(CIEDE2000, "_deg6InRad", degrees2radians(6));
__publicField(CIEDE2000, "_deg63InRad", degrees2radians(63));
__publicField(CIEDE2000, "_deg275InRad", degrees2radians(275));
__publicField(CIEDE2000, "_deg25InRad", degrees2radians(25));

// src/distance/cmetric.ts
var CMetric = class extends AbstractDistanceCalculator {
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    const rmean = (r1 + r2) / 2 * this._whitePoint.r;
    const r = (r1 - r2) * this._whitePoint.r;
    const g = (g1 - g2) * this._whitePoint.g;
    const b = (b1 - b2) * this._whitePoint.b;
    const dE = ((512 + rmean) * r * r >> 8) + 4 * g * g + ((767 - rmean) * b * b >> 8);
    const dA = (a2 - a1) * this._whitePoint.a;
    return Math.sqrt(dE + dA * dA);
  }
  _setDefaults() {
  }
};

// src/distance/euclidean.ts
var AbstractEuclidean = class extends AbstractDistanceCalculator {
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    const dR = r2 - r1;
    const dG = g2 - g1;
    const dB = b2 - b1;
    const dA = a2 - a1;
    return Math.sqrt(this._kR * dR * dR + this._kG * dG * dG + this._kB * dB * dB + this._kA * dA * dA);
  }
};
var Euclidean = class extends AbstractEuclidean {
  _setDefaults() {
    this._kR = 1;
    this._kG = 1;
    this._kB = 1;
    this._kA = 1;
  }
};
var EuclideanBT709 = class extends AbstractEuclidean {
  _setDefaults() {
    this._kR = 0.2126 /* RED */;
    this._kG = 0.7152 /* GREEN */;
    this._kB = 0.0722 /* BLUE */;
    this._kA = 1;
  }
};
var EuclideanBT709NoAlpha = class extends AbstractEuclidean {
  _setDefaults() {
    this._kR = 0.2126 /* RED */;
    this._kG = 0.7152 /* GREEN */;
    this._kB = 0.0722 /* BLUE */;
    this._kA = 0;
  }
};

// src/distance/manhattan.ts
var AbstractManhattan = class extends AbstractDistanceCalculator {
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    let dR = r2 - r1;
    let dG = g2 - g1;
    let dB = b2 - b1;
    let dA = a2 - a1;
    if (dR < 0)
      dR = 0 - dR;
    if (dG < 0)
      dG = 0 - dG;
    if (dB < 0)
      dB = 0 - dB;
    if (dA < 0)
      dA = 0 - dA;
    return this._kR * dR + this._kG * dG + this._kB * dB + this._kA * dA;
  }
};
var Manhattan = class extends AbstractManhattan {
  _setDefaults() {
    this._kR = 1;
    this._kG = 1;
    this._kB = 1;
    this._kA = 1;
  }
};
var ManhattanNommyde = class extends AbstractManhattan {
  _setDefaults() {
    this._kR = 0.4984;
    this._kG = 0.8625;
    this._kB = 0.2979;
    this._kA = 1;
  }
};
var ManhattanBT709 = class extends AbstractManhattan {
  _setDefaults() {
    this._kR = 0.2126 /* RED */;
    this._kG = 0.7152 /* GREEN */;
    this._kB = 0.0722 /* BLUE */;
    this._kA = 1;
  }
};

// src/distance/pngQuant.ts
var PNGQuant = class extends AbstractDistanceCalculator {
  calculateRaw(r1, g1, b1, a1, r2, g2, b2, a2) {
    const alphas = (a2 - a1) * this._whitePoint.a;
    return this._colordifferenceCh(r1 * this._whitePoint.r, r2 * this._whitePoint.r, alphas) + this._colordifferenceCh(g1 * this._whitePoint.g, g2 * this._whitePoint.g, alphas) + this._colordifferenceCh(b1 * this._whitePoint.b, b2 * this._whitePoint.b, alphas);
  }
  _colordifferenceCh(x2, y2, alphas) {
    const black = x2 - y2;
    const white = black + alphas;
    return black * black + white * white;
  }
  _setDefaults() {
  }
};

// src/palette/index.ts
var palette_exports = {};
__export(palette_exports, {
  AbstractPaletteQuantizer: () => AbstractPaletteQuantizer,
  ColorHistogram: () => ColorHistogram,
  NeuQuant: () => NeuQuant,
  NeuQuantFloat: () => NeuQuantFloat,
  RGBQuant: () => RGBQuant,
  WuColorCube: () => WuColorCube,
  WuQuant: () => WuQuant
});

// src/palette/paletteQuantizer.ts
var AbstractPaletteQuantizer = class {
  quantizeSync() {
    for (const value of this.quantize()) {
      if (value.palette) {
        return value.palette;
      }
    }
    throw new Error("unreachable");
  }
};

// src/utils/point.ts
var Point = class {
  constructor() {
    __publicField(this, "r");
    __publicField(this, "g");
    __publicField(this, "b");
    __publicField(this, "a");
    __publicField(this, "uint32");
    __publicField(this, "rgba");
    this.uint32 = -1 >>> 0;
    this.r = this.g = this.b = this.a = 0;
    this.rgba = new Array(4);
    this.rgba[0] = 0;
    this.rgba[1] = 0;
    this.rgba[2] = 0;
    this.rgba[3] = 0;
  }
  static createByQuadruplet(quadruplet) {
    const point = new Point();
    point.r = quadruplet[0] | 0;
    point.g = quadruplet[1] | 0;
    point.b = quadruplet[2] | 0;
    point.a = quadruplet[3] | 0;
    point._loadUINT32();
    point._loadQuadruplet();
    return point;
  }
  static createByRGBA(red, green, blue, alpha) {
    const point = new Point();
    point.r = red | 0;
    point.g = green | 0;
    point.b = blue | 0;
    point.a = alpha | 0;
    point._loadUINT32();
    point._loadQuadruplet();
    return point;
  }
  static createByUint32(uint32) {
    const point = new Point();
    point.uint32 = uint32 >>> 0;
    point._loadRGBA();
    point._loadQuadruplet();
    return point;
  }
  from(point) {
    this.r = point.r;
    this.g = point.g;
    this.b = point.b;
    this.a = point.a;
    this.uint32 = point.uint32;
    this.rgba[0] = point.r;
    this.rgba[1] = point.g;
    this.rgba[2] = point.b;
    this.rgba[3] = point.a;
  }
  getLuminosity(useAlphaChannel) {
    let r = this.r;
    let g = this.g;
    let b = this.b;
    if (useAlphaChannel) {
      r = Math.min(255, 255 - this.a + this.a * r / 255);
      g = Math.min(255, 255 - this.a + this.a * g / 255);
      b = Math.min(255, 255 - this.a + this.a * b / 255);
    }
    return r * 0.2126 /* RED */ + g * 0.7152 /* GREEN */ + b * 0.0722 /* BLUE */;
  }
  _loadUINT32() {
    this.uint32 = (this.a << 24 | this.b << 16 | this.g << 8 | this.r) >>> 0;
  }
  _loadRGBA() {
    this.r = this.uint32 & 255;
    this.g = this.uint32 >>> 8 & 255;
    this.b = this.uint32 >>> 16 & 255;
    this.a = this.uint32 >>> 24 & 255;
  }
  _loadQuadruplet() {
    this.rgba[0] = this.r;
    this.rgba[1] = this.g;
    this.rgba[2] = this.b;
    this.rgba[3] = this.a;
  }
};

// src/utils/pointContainer.ts
var PointContainer = class {
  constructor() {
    __publicField(this, "_pointArray");
    __publicField(this, "_width");
    __publicField(this, "_height");
    this._width = 0;
    this._height = 0;
    this._pointArray = [];
  }
  getWidth() {
    return this._width;
  }
  getHeight() {
    return this._height;
  }
  setWidth(width) {
    this._width = width;
  }
  setHeight(height) {
    this._height = height;
  }
  getPointArray() {
    return this._pointArray;
  }
  clone() {
    const clone = new PointContainer();
    clone._width = this._width;
    clone._height = this._height;
    for (let i = 0, l = this._pointArray.length; i < l; i++) {
      clone._pointArray[i] = Point.createByUint32(this._pointArray[i].uint32 | 0);
    }
    return clone;
  }
  toUint32Array() {
    const l = this._pointArray.length;
    const uint32Array = new Uint32Array(l);
    for (let i = 0; i < l; i++) {
      uint32Array[i] = this._pointArray[i].uint32;
    }
    return uint32Array;
  }
  toUint8Array() {
    return new Uint8Array(this.toUint32Array().buffer);
  }
  static fromHTMLImageElement(img) {
    const width = img.naturalWidth;
    const height = img.naturalHeight;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, width, height, 0, 0, width, height);
    return PointContainer.fromHTMLCanvasElement(canvas);
  }
  static fromHTMLCanvasElement(canvas) {
    const width = canvas.width;
    const height = canvas.height;
    const ctx = canvas.getContext("2d");
    const imgData = ctx.getImageData(0, 0, width, height);
    return PointContainer.fromImageData(imgData);
  }
  static fromImageData(imageData) {
    const width = imageData.width;
    const height = imageData.height;
    return PointContainer.fromUint8Array(imageData.data, width, height);
  }
  static fromUint8Array(uint8Array, width, height) {
    switch (Object.prototype.toString.call(uint8Array)) {
      case "[object Uint8ClampedArray]":
      case "[object Uint8Array]":
        break;
      default:
        uint8Array = new Uint8Array(uint8Array);
    }
    const uint32Array = new Uint32Array(uint8Array.buffer);
    return PointContainer.fromUint32Array(uint32Array, width, height);
  }
  static fromUint32Array(uint32Array, width, height) {
    const container = new PointContainer();
    container._width = width;
    container._height = height;
    for (let i = 0, l = uint32Array.length; i < l; i++) {
      container._pointArray[i] = Point.createByUint32(uint32Array[i] | 0);
    }
    return container;
  }
  static fromBuffer(buffer, width, height) {
    const uint32Array = new Uint32Array(buffer.buffer, buffer.byteOffset, buffer.byteLength / Uint32Array.BYTES_PER_ELEMENT);
    return PointContainer.fromUint32Array(uint32Array, width, height);
  }
};

// src/utils/palette.ts
var hueGroups = 10;
function hueGroup(hue, segmentsNumber) {
  const maxHue = 360;
  const seg = maxHue / segmentsNumber;
  const half = seg / 2;
  for (let i = 1, mid = seg - half; i < segmentsNumber; i++, mid += seg) {
    if (hue >= mid && hue < mid + seg)
      return i;
  }
  return 0;
}
var Palette = class {
  constructor() {
    __publicField(this, "_pointContainer");
    __publicField(this, "_pointArray", []);
    __publicField(this, "_i32idx", {});
    this._pointContainer = new PointContainer();
    this._pointContainer.setHeight(1);
    this._pointArray = this._pointContainer.getPointArray();
  }
  add(color) {
    this._pointArray.push(color);
    this._pointContainer.setWidth(this._pointArray.length);
  }
  has(color) {
    for (let i = this._pointArray.length - 1; i >= 0; i--) {
      if (color.uint32 === this._pointArray[i].uint32)
        return true;
    }
    return false;
  }
  getNearestColor(colorDistanceCalculator, color) {
    return this._pointArray[this._getNearestIndex(colorDistanceCalculator, color) | 0];
  }
  getPointContainer() {
    return this._pointContainer;
  }
  _nearestPointFromCache(key) {
    return typeof this._i32idx[key] === "number" ? this._i32idx[key] : -1;
  }
  _getNearestIndex(colorDistanceCalculator, point) {
    let idx = this._nearestPointFromCache("" + point.uint32);
    if (idx >= 0)
      return idx;
    let minimalDistance = Number.MAX_VALUE;
    idx = 0;
    for (let i = 0, l = this._pointArray.length; i < l; i++) {
      const p = this._pointArray[i];
      const distance = colorDistanceCalculator.calculateRaw(point.r, point.g, point.b, point.a, p.r, p.g, p.b, p.a);
      if (distance < minimalDistance) {
        minimalDistance = distance;
        idx = i;
      }
    }
    this._i32idx[point.uint32] = idx;
    return idx;
  }
  sort() {
    this._i32idx = {};
    this._pointArray.sort((a, b) => {
      const hslA = rgb2hsl(a.r, a.g, a.b);
      const hslB = rgb2hsl(b.r, b.g, b.b);
      const hueA = a.r === a.g && a.g === a.b ? 0 : 1 + hueGroup(hslA.h, hueGroups);
      const hueB = b.r === b.g && b.g === b.b ? 0 : 1 + hueGroup(hslB.h, hueGroups);
      const hueDiff = hueB - hueA;
      if (hueDiff)
        return -hueDiff;
      const lA = a.getLuminosity(true);
      const lB = b.getLuminosity(true);
      if (lB - lA !== 0)
        return lB - lA;
      const satDiff = (hslB.s * 100 | 0) - (hslA.s * 100 | 0);
      if (satDiff)
        return -satDiff;
      return 0;
    });
  }
};

// src/utils/index.ts
var utils_exports = {};
__export(utils_exports, {
  HueStatistics: () => HueStatistics,
  Palette: () => Palette,
  Point: () => Point,
  PointContainer: () => PointContainer,
  ProgressTracker: () => ProgressTracker,
  arithmetic: () => arithmetic_exports
});

// src/utils/hueStatistics.ts
var HueGroup = class {
  constructor() {
    __publicField(this, "num", 0);
    __publicField(this, "cols", []);
  }
};
var HueStatistics = class {
  constructor(numGroups, minCols) {
    __publicField(this, "_numGroups");
    __publicField(this, "_minCols");
    __publicField(this, "_stats");
    __publicField(this, "_groupsFull");
    this._numGroups = numGroups;
    this._minCols = minCols;
    this._stats = [];
    for (let i = 0; i <= numGroups; i++) {
      this._stats[i] = new HueGroup();
    }
    this._groupsFull = 0;
  }
  check(i32) {
    if (this._groupsFull === this._numGroups + 1) {
      this.check = () => {
      };
    }
    const r = i32 & 255;
    const g = i32 >>> 8 & 255;
    const b = i32 >>> 16 & 255;
    const hg = r === g && g === b ? 0 : 1 + hueGroup(rgb2hsl(r, g, b).h, this._numGroups);
    const gr = this._stats[hg];
    const min = this._minCols;
    gr.num++;
    if (gr.num > min) {
      return;
    }
    if (gr.num === min) {
      this._groupsFull++;
    }
    if (gr.num <= min) {
      this._stats[hg].cols.push(i32);
    }
  }
  injectIntoDictionary(histG) {
    for (let i = 0; i <= this._numGroups; i++) {
      if (this._stats[i].num <= this._minCols) {
        this._stats[i].cols.forEach((col) => {
          if (!histG[col]) {
            histG[col] = 1;
          } else {
            histG[col]++;
          }
        });
      }
    }
  }
  injectIntoArray(histG) {
    for (let i = 0; i <= this._numGroups; i++) {
      if (this._stats[i].num <= this._minCols) {
        this._stats[i].cols.forEach((col) => {
          if (histG.indexOf(col) === -1) {
            histG.push(col);
          }
        });
      }
    }
  }
};

// src/utils/progressTracker.ts
var _ProgressTracker = class {
  constructor(valueRange, progressRange) {
    __publicField(this, "progress");
    __publicField(this, "_step");
    __publicField(this, "_range");
    __publicField(this, "_last");
    __publicField(this, "_progressRange");
    this._range = valueRange;
    this._progressRange = progressRange;
    this._step = Math.max(1, this._range / (_ProgressTracker.steps + 1) | 0);
    this._last = -this._step;
    this.progress = 0;
  }
  shouldNotify(current) {
    if (current - this._last >= this._step) {
      this._last = current;
      this.progress = Math.min(this._progressRange * this._last / this._range, this._progressRange);
      return true;
    }
    return false;
  }
};
var ProgressTracker = _ProgressTracker;
__publicField(ProgressTracker, "steps", 100);

// src/palette/neuquant/neuquant.ts
var networkBiasShift = 3;
var Neuron = class {
  constructor(defaultValue) {
    __publicField(this, "r");
    __publicField(this, "g");
    __publicField(this, "b");
    __publicField(this, "a");
    this.r = this.g = this.b = this.a = defaultValue;
  }
  toPoint() {
    return Point.createByRGBA(this.r >> networkBiasShift, this.g >> networkBiasShift, this.b >> networkBiasShift, this.a >> networkBiasShift);
  }
  subtract(r, g, b, a) {
    this.r -= r | 0;
    this.g -= g | 0;
    this.b -= b | 0;
    this.a -= a | 0;
  }
};
var _NeuQuant = class extends AbstractPaletteQuantizer {
  constructor(colorDistanceCalculator, colors = 256) {
    super();
    __publicField(this, "_pointArray");
    __publicField(this, "_networkSize");
    __publicField(this, "_network");
    __publicField(this, "_sampleFactor");
    __publicField(this, "_radPower");
    __publicField(this, "_freq");
    __publicField(this, "_bias");
    __publicField(this, "_distance");
    this._distance = colorDistanceCalculator;
    this._pointArray = [];
    this._sampleFactor = 1;
    this._networkSize = colors;
    this._distance.setWhitePoint(255 << networkBiasShift, 255 << networkBiasShift, 255 << networkBiasShift, 255 << networkBiasShift);
  }
  sample(pointContainer) {
    this._pointArray = this._pointArray.concat(pointContainer.getPointArray());
  }
  *quantize() {
    this._init();
    yield* this._learn();
    yield {
      palette: this._buildPalette(),
      progress: 100
    };
  }
  _init() {
    this._freq = [];
    this._bias = [];
    this._radPower = [];
    this._network = [];
    for (let i = 0; i < this._networkSize; i++) {
      this._network[i] = new Neuron((i << networkBiasShift + 8) / this._networkSize | 0);
      this._freq[i] = _NeuQuant._initialBias / this._networkSize | 0;
      this._bias[i] = 0;
    }
  }
  *_learn() {
    let sampleFactor = this._sampleFactor;
    const pointsNumber = this._pointArray.length;
    if (pointsNumber < _NeuQuant._minpicturebytes)
      sampleFactor = 1;
    const alphadec = 30 + (sampleFactor - 1) / 3 | 0;
    const pointsToSample = pointsNumber / sampleFactor | 0;
    let delta = pointsToSample / _NeuQuant._nCycles | 0;
    let alpha = _NeuQuant._initAlpha;
    let radius = (this._networkSize >> 3) * _NeuQuant._radiusBias;
    let rad = radius >> _NeuQuant._radiusBiasShift;
    if (rad <= 1)
      rad = 0;
    for (let i = 0; i < rad; i++) {
      this._radPower[i] = alpha * ((rad * rad - i * i) * _NeuQuant._radBias / (rad * rad)) >>> 0;
    }
    let step;
    if (pointsNumber < _NeuQuant._minpicturebytes) {
      step = 1;
    } else if (pointsNumber % _NeuQuant._prime1 !== 0) {
      step = _NeuQuant._prime1;
    } else if (pointsNumber % _NeuQuant._prime2 !== 0) {
      step = _NeuQuant._prime2;
    } else if (pointsNumber % _NeuQuant._prime3 !== 0) {
      step = _NeuQuant._prime3;
    } else {
      step = _NeuQuant._prime4;
    }
    const tracker = new ProgressTracker(pointsToSample, 99);
    for (let i = 0, pointIndex = 0; i < pointsToSample; ) {
      if (tracker.shouldNotify(i)) {
        yield {
          progress: tracker.progress
        };
      }
      const point = this._pointArray[pointIndex];
      const b = point.b << networkBiasShift;
      const g = point.g << networkBiasShift;
      const r = point.r << networkBiasShift;
      const a = point.a << networkBiasShift;
      const neuronIndex = this._contest(b, g, r, a);
      this._alterSingle(alpha, neuronIndex, b, g, r, a);
      if (rad !== 0)
        this._alterNeighbour(rad, neuronIndex, b, g, r, a);
      pointIndex += step;
      if (pointIndex >= pointsNumber)
        pointIndex -= pointsNumber;
      i++;
      if (delta === 0)
        delta = 1;
      if (i % delta === 0) {
        alpha -= alpha / alphadec | 0;
        radius -= radius / _NeuQuant._radiusDecrease | 0;
        rad = radius >> _NeuQuant._radiusBiasShift;
        if (rad <= 1)
          rad = 0;
        for (let j = 0; j < rad; j++) {
          this._radPower[j] = alpha * ((rad * rad - j * j) * _NeuQuant._radBias / (rad * rad)) >>> 0;
        }
      }
    }
  }
  _buildPalette() {
    const palette = new Palette();
    this._network.forEach((neuron) => {
      palette.add(neuron.toPoint());
    });
    palette.sort();
    return palette;
  }
  _alterNeighbour(rad, i, b, g, r, al) {
    let lo = i - rad;
    if (lo < -1)
      lo = -1;
    let hi = i + rad;
    if (hi > this._networkSize)
      hi = this._networkSize;
    let j = i + 1;
    let k = i - 1;
    let m = 1;
    while (j < hi || k > lo) {
      const a = this._radPower[m++] / _NeuQuant._alphaRadBias;
      if (j < hi) {
        const p = this._network[j++];
        p.subtract(a * (p.r - r), a * (p.g - g), a * (p.b - b), a * (p.a - al));
      }
      if (k > lo) {
        const p = this._network[k--];
        p.subtract(a * (p.r - r), a * (p.g - g), a * (p.b - b), a * (p.a - al));
      }
    }
  }
  _alterSingle(alpha, i, b, g, r, a) {
    alpha /= _NeuQuant._initAlpha;
    const n = this._network[i];
    n.subtract(alpha * (n.r - r), alpha * (n.g - g), alpha * (n.b - b), alpha * (n.a - a));
  }
  _contest(b, g, r, a) {
    const multiplier = 255 * 4 << networkBiasShift;
    let bestd = ~(1 << 31);
    let bestbiasd = bestd;
    let bestpos = -1;
    let bestbiaspos = bestpos;
    for (let i = 0; i < this._networkSize; i++) {
      const n = this._network[i];
      const dist = this._distance.calculateNormalized(n, { r, g, b, a }) * multiplier | 0;
      if (dist < bestd) {
        bestd = dist;
        bestpos = i;
      }
      const biasdist = dist - (this._bias[i] >> _NeuQuant._initialBiasShift - networkBiasShift);
      if (biasdist < bestbiasd) {
        bestbiasd = biasdist;
        bestbiaspos = i;
      }
      const betafreq = this._freq[i] >> _NeuQuant._betaShift;
      this._freq[i] -= betafreq;
      this._bias[i] += betafreq << _NeuQuant._gammaShift;
    }
    this._freq[bestpos] += _NeuQuant._beta;
    this._bias[bestpos] -= _NeuQuant._betaGamma;
    return bestbiaspos;
  }
};
var NeuQuant = _NeuQuant;
__publicField(NeuQuant, "_prime1", 499);
__publicField(NeuQuant, "_prime2", 491);
__publicField(NeuQuant, "_prime3", 487);
__publicField(NeuQuant, "_prime4", 503);
__publicField(NeuQuant, "_minpicturebytes", _NeuQuant._prime4);
__publicField(NeuQuant, "_nCycles", 100);
__publicField(NeuQuant, "_initialBiasShift", 16);
__publicField(NeuQuant, "_initialBias", 1 << _NeuQuant._initialBiasShift);
__publicField(NeuQuant, "_gammaShift", 10);
__publicField(NeuQuant, "_betaShift", 10);
__publicField(NeuQuant, "_beta", _NeuQuant._initialBias >> _NeuQuant._betaShift);
__publicField(NeuQuant, "_betaGamma", _NeuQuant._initialBias << _NeuQuant._gammaShift - _NeuQuant._betaShift);
__publicField(NeuQuant, "_radiusBiasShift", 6);
__publicField(NeuQuant, "_radiusBias", 1 << _NeuQuant._radiusBiasShift);
__publicField(NeuQuant, "_radiusDecrease", 30);
__publicField(NeuQuant, "_alphaBiasShift", 10);
__publicField(NeuQuant, "_initAlpha", 1 << _NeuQuant._alphaBiasShift);
__publicField(NeuQuant, "_radBiasShift", 8);
__publicField(NeuQuant, "_radBias", 1 << _NeuQuant._radBiasShift);
__publicField(NeuQuant, "_alphaRadBiasShift", _NeuQuant._alphaBiasShift + _NeuQuant._radBiasShift);
__publicField(NeuQuant, "_alphaRadBias", 1 << _NeuQuant._alphaRadBiasShift);

// src/palette/neuquant/neuquantFloat.ts
var networkBiasShift2 = 3;
var NeuronFloat = class {
  constructor(defaultValue) {
    __publicField(this, "r");
    __publicField(this, "g");
    __publicField(this, "b");
    __publicField(this, "a");
    this.r = this.g = this.b = this.a = defaultValue;
  }
  toPoint() {
    return Point.createByRGBA(this.r >> networkBiasShift2, this.g >> networkBiasShift2, this.b >> networkBiasShift2, this.a >> networkBiasShift2);
  }
  subtract(r, g, b, a) {
    this.r -= r;
    this.g -= g;
    this.b -= b;
    this.a -= a;
  }
};
var _NeuQuantFloat = class extends AbstractPaletteQuantizer {
  constructor(colorDistanceCalculator, colors = 256) {
    super();
    __publicField(this, "_pointArray");
    __publicField(this, "_networkSize");
    __publicField(this, "_network");
    __publicField(this, "_sampleFactor");
    __publicField(this, "_radPower");
    __publicField(this, "_freq");
    __publicField(this, "_bias");
    __publicField(this, "_distance");
    this._distance = colorDistanceCalculator;
    this._pointArray = [];
    this._sampleFactor = 1;
    this._networkSize = colors;
    this._distance.setWhitePoint(255 << networkBiasShift2, 255 << networkBiasShift2, 255 << networkBiasShift2, 255 << networkBiasShift2);
  }
  sample(pointContainer) {
    this._pointArray = this._pointArray.concat(pointContainer.getPointArray());
  }
  *quantize() {
    this._init();
    yield* this._learn();
    yield {
      palette: this._buildPalette(),
      progress: 100
    };
  }
  _init() {
    this._freq = [];
    this._bias = [];
    this._radPower = [];
    this._network = [];
    for (let i = 0; i < this._networkSize; i++) {
      this._network[i] = new NeuronFloat((i << networkBiasShift2 + 8) / this._networkSize);
      this._freq[i] = _NeuQuantFloat._initialBias / this._networkSize;
      this._bias[i] = 0;
    }
  }
  *_learn() {
    let sampleFactor = this._sampleFactor;
    const pointsNumber = this._pointArray.length;
    if (pointsNumber < _NeuQuantFloat._minpicturebytes)
      sampleFactor = 1;
    const alphadec = 30 + (sampleFactor - 1) / 3;
    const pointsToSample = pointsNumber / sampleFactor;
    let delta = pointsToSample / _NeuQuantFloat._nCycles | 0;
    let alpha = _NeuQuantFloat._initAlpha;
    let radius = (this._networkSize >> 3) * _NeuQuantFloat._radiusBias;
    let rad = radius >> _NeuQuantFloat._radiusBiasShift;
    if (rad <= 1)
      rad = 0;
    for (let i = 0; i < rad; i++) {
      this._radPower[i] = alpha * ((rad * rad - i * i) * _NeuQuantFloat._radBias / (rad * rad));
    }
    let step;
    if (pointsNumber < _NeuQuantFloat._minpicturebytes) {
      step = 1;
    } else if (pointsNumber % _NeuQuantFloat._prime1 !== 0) {
      step = _NeuQuantFloat._prime1;
    } else if (pointsNumber % _NeuQuantFloat._prime2 !== 0) {
      step = _NeuQuantFloat._prime2;
    } else if (pointsNumber % _NeuQuantFloat._prime3 !== 0) {
      step = _NeuQuantFloat._prime3;
    } else {
      step = _NeuQuantFloat._prime4;
    }
    const tracker = new ProgressTracker(pointsToSample, 99);
    for (let i = 0, pointIndex = 0; i < pointsToSample; ) {
      if (tracker.shouldNotify(i)) {
        yield {
          progress: tracker.progress
        };
      }
      const point = this._pointArray[pointIndex];
      const b = point.b << networkBiasShift2;
      const g = point.g << networkBiasShift2;
      const r = point.r << networkBiasShift2;
      const a = point.a << networkBiasShift2;
      const neuronIndex = this._contest(b, g, r, a);
      this._alterSingle(alpha, neuronIndex, b, g, r, a);
      if (rad !== 0)
        this._alterNeighbour(rad, neuronIndex, b, g, r, a);
      pointIndex += step;
      if (pointIndex >= pointsNumber)
        pointIndex -= pointsNumber;
      i++;
      if (delta === 0)
        delta = 1;
      if (i % delta === 0) {
        alpha -= alpha / alphadec;
        radius -= radius / _NeuQuantFloat._radiusDecrease;
        rad = radius >> _NeuQuantFloat._radiusBiasShift;
        if (rad <= 1)
          rad = 0;
        for (let j = 0; j < rad; j++) {
          this._radPower[j] = alpha * ((rad * rad - j * j) * _NeuQuantFloat._radBias / (rad * rad));
        }
      }
    }
  }
  _buildPalette() {
    const palette = new Palette();
    this._network.forEach((neuron) => {
      palette.add(neuron.toPoint());
    });
    palette.sort();
    return palette;
  }
  _alterNeighbour(rad, i, b, g, r, al) {
    let lo = i - rad;
    if (lo < -1)
      lo = -1;
    let hi = i + rad;
    if (hi > this._networkSize)
      hi = this._networkSize;
    let j = i + 1;
    let k = i - 1;
    let m = 1;
    while (j < hi || k > lo) {
      const a = this._radPower[m++] / _NeuQuantFloat._alphaRadBias;
      if (j < hi) {
        const p = this._network[j++];
        p.subtract(a * (p.r - r), a * (p.g - g), a * (p.b - b), a * (p.a - al));
      }
      if (k > lo) {
        const p = this._network[k--];
        p.subtract(a * (p.r - r), a * (p.g - g), a * (p.b - b), a * (p.a - al));
      }
    }
  }
  _alterSingle(alpha, i, b, g, r, a) {
    alpha /= _NeuQuantFloat._initAlpha;
    const n = this._network[i];
    n.subtract(alpha * (n.r - r), alpha * (n.g - g), alpha * (n.b - b), alpha * (n.a - a));
  }
  _contest(b, g, r, al) {
    const multiplier = 255 * 4 << networkBiasShift2;
    let bestd = ~(1 << 31);
    let bestbiasd = bestd;
    let bestpos = -1;
    let bestbiaspos = bestpos;
    for (let i = 0; i < this._networkSize; i++) {
      const n = this._network[i];
      const dist = this._distance.calculateNormalized(n, { r, g, b, a: al }) * multiplier;
      if (dist < bestd) {
        bestd = dist;
        bestpos = i;
      }
      const biasdist = dist - (this._bias[i] >> _NeuQuantFloat._initialBiasShift - networkBiasShift2);
      if (biasdist < bestbiasd) {
        bestbiasd = biasdist;
        bestbiaspos = i;
      }
      const betafreq = this._freq[i] >> _NeuQuantFloat._betaShift;
      this._freq[i] -= betafreq;
      this._bias[i] += betafreq << _NeuQuantFloat._gammaShift;
    }
    this._freq[bestpos] += _NeuQuantFloat._beta;
    this._bias[bestpos] -= _NeuQuantFloat._betaGamma;
    return bestbiaspos;
  }
};
var NeuQuantFloat = _NeuQuantFloat;
__publicField(NeuQuantFloat, "_prime1", 499);
__publicField(NeuQuantFloat, "_prime2", 491);
__publicField(NeuQuantFloat, "_prime3", 487);
__publicField(NeuQuantFloat, "_prime4", 503);
__publicField(NeuQuantFloat, "_minpicturebytes", _NeuQuantFloat._prime4);
__publicField(NeuQuantFloat, "_nCycles", 100);
__publicField(NeuQuantFloat, "_initialBiasShift", 16);
__publicField(NeuQuantFloat, "_initialBias", 1 << _NeuQuantFloat._initialBiasShift);
__publicField(NeuQuantFloat, "_gammaShift", 10);
__publicField(NeuQuantFloat, "_betaShift", 10);
__publicField(NeuQuantFloat, "_beta", _NeuQuantFloat._initialBias >> _NeuQuantFloat._betaShift);
__publicField(NeuQuantFloat, "_betaGamma", _NeuQuantFloat._initialBias << _NeuQuantFloat._gammaShift - _NeuQuantFloat._betaShift);
__publicField(NeuQuantFloat, "_radiusBiasShift", 6);
__publicField(NeuQuantFloat, "_radiusBias", 1 << _NeuQuantFloat._radiusBiasShift);
__publicField(NeuQuantFloat, "_radiusDecrease", 30);
__publicField(NeuQuantFloat, "_alphaBiasShift", 10);
__publicField(NeuQuantFloat, "_initAlpha", 1 << _NeuQuantFloat._alphaBiasShift);
__publicField(NeuQuantFloat, "_radBiasShift", 8);
__publicField(NeuQuantFloat, "_radBias", 1 << _NeuQuantFloat._radBiasShift);
__publicField(NeuQuantFloat, "_alphaRadBiasShift", _NeuQuantFloat._alphaBiasShift + _NeuQuantFloat._radBiasShift);
__publicField(NeuQuantFloat, "_alphaRadBias", 1 << _NeuQuantFloat._alphaRadBiasShift);

// src/palette/rgbquant/colorHistogram.ts
var _ColorHistogram = class {
  constructor(method, colors) {
    __publicField(this, "_method");
    __publicField(this, "_hueStats");
    __publicField(this, "_histogram");
    __publicField(this, "_initColors");
    __publicField(this, "_minHueCols");
    this._method = method;
    this._minHueCols = colors << 2;
    this._initColors = colors << 2;
    this._hueStats = new HueStatistics(_ColorHistogram._hueGroups, this._minHueCols);
    this._histogram = /* @__PURE__ */ Object.create(null);
  }
  sample(pointContainer) {
    switch (this._method) {
      case 1:
        this._colorStats1D(pointContainer);
        break;
      case 2:
        this._colorStats2D(pointContainer);
        break;
    }
  }
  getImportanceSortedColorsIDXI32() {
    const sorted = stableSort(Object.keys(this._histogram), (a, b) => this._histogram[b] - this._histogram[a]);
    if (sorted.length === 0) {
      return [];
    }
    let idxi32;
    switch (this._method) {
      case 1:
        const initialColorsLimit = Math.min(sorted.length, this._initColors);
        const last = sorted[initialColorsLimit - 1];
        const freq = this._histogram[last];
        idxi32 = sorted.slice(0, initialColorsLimit);
        let pos = initialColorsLimit;
        const len = sorted.length;
        while (pos < len && this._histogram[sorted[pos]] === freq) {
          idxi32.push(sorted[pos++]);
        }
        this._hueStats.injectIntoArray(idxi32);
        break;
      case 2:
        idxi32 = sorted;
        break;
      default:
        throw new Error("Incorrect method");
    }
    return idxi32.map((v) => +v);
  }
  _colorStats1D(pointContainer) {
    const histG = this._histogram;
    const pointArray = pointContainer.getPointArray();
    const len = pointArray.length;
    for (let i = 0; i < len; i++) {
      const col = pointArray[i].uint32;
      this._hueStats.check(col);
      if (col in histG) {
        histG[col]++;
      } else {
        histG[col] = 1;
      }
    }
  }
  _colorStats2D(pointContainer) {
    const width = pointContainer.getWidth();
    const height = pointContainer.getHeight();
    const pointArray = pointContainer.getPointArray();
    const boxW = _ColorHistogram._boxSize[0];
    const boxH = _ColorHistogram._boxSize[1];
    const area = boxW * boxH;
    const boxes = this._makeBoxes(width, height, boxW, boxH);
    const histG = this._histogram;
    boxes.forEach((box) => {
      let effc = Math.round(box.w * box.h / area) * _ColorHistogram._boxPixels;
      if (effc < 2)
        effc = 2;
      const histL = {};
      this._iterateBox(box, width, (i) => {
        const col = pointArray[i].uint32;
        this._hueStats.check(col);
        if (col in histG) {
          histG[col]++;
        } else if (col in histL) {
          if (++histL[col] >= effc) {
            histG[col] = histL[col];
          }
        } else {
          histL[col] = 1;
        }
      });
    });
    this._hueStats.injectIntoDictionary(histG);
  }
  _iterateBox(bbox, wid, fn) {
    const b = bbox;
    const i0 = b.y * wid + b.x;
    const i1 = (b.y + b.h - 1) * wid + (b.x + b.w - 1);
    const incr = wid - b.w + 1;
    let cnt = 0;
    let i = i0;
    do {
      fn.call(this, i);
      i += ++cnt % b.w === 0 ? incr : 1;
    } while (i <= i1);
  }
  _makeBoxes(width, height, stepX, stepY) {
    const wrem = width % stepX;
    const hrem = height % stepY;
    const xend = width - wrem;
    const yend = height - hrem;
    const boxesArray = [];
    for (let y2 = 0; y2 < height; y2 += stepY) {
      for (let x2 = 0; x2 < width; x2 += stepX) {
        boxesArray.push({
          x: x2,
          y: y2,
          w: x2 === xend ? wrem : stepX,
          h: y2 === yend ? hrem : stepY
        });
      }
    }
    return boxesArray;
  }
};
var ColorHistogram = _ColorHistogram;
__publicField(ColorHistogram, "_boxSize", [64, 64]);
__publicField(ColorHistogram, "_boxPixels", 2);
__publicField(ColorHistogram, "_hueGroups", 10);

// src/palette/rgbquant/rgbquant.ts
var RemovedColor = class {
  constructor(index, color, distance) {
    __publicField(this, "index");
    __publicField(this, "color");
    __publicField(this, "distance");
    this.index = index;
    this.color = color;
    this.distance = distance;
  }
};
var RGBQuant = class extends AbstractPaletteQuantizer {
  constructor(colorDistanceCalculator, colors = 256, method = 2) {
    super();
    __publicField(this, "_colors");
    __publicField(this, "_initialDistance");
    __publicField(this, "_distanceIncrement");
    __publicField(this, "_histogram");
    __publicField(this, "_distance");
    this._distance = colorDistanceCalculator;
    this._colors = colors;
    this._histogram = new ColorHistogram(method, colors);
    this._initialDistance = 0.01;
    this._distanceIncrement = 5e-3;
  }
  sample(image) {
    this._histogram.sample(image);
  }
  *quantize() {
    const idxi32 = this._histogram.getImportanceSortedColorsIDXI32();
    if (idxi32.length === 0) {
      throw new Error("No colors in image");
    }
    yield* this._buildPalette(idxi32);
  }
  *_buildPalette(idxi32) {
    const palette = new Palette();
    const colorArray = palette.getPointContainer().getPointArray();
    const usageArray = new Array(idxi32.length);
    for (let i = 0; i < idxi32.length; i++) {
      colorArray.push(Point.createByUint32(idxi32[i]));
      usageArray[i] = 1;
    }
    const len = colorArray.length;
    const memDist = [];
    let palLen = len;
    let thold = this._initialDistance;
    const tracker = new ProgressTracker(palLen - this._colors, 99);
    while (palLen > this._colors) {
      memDist.length = 0;
      for (let i = 0; i < len; i++) {
        if (tracker.shouldNotify(len - palLen)) {
          yield {
            progress: tracker.progress
          };
        }
        if (usageArray[i] === 0)
          continue;
        const pxi = colorArray[i];
        for (let j = i + 1; j < len; j++) {
          if (usageArray[j] === 0)
            continue;
          const pxj = colorArray[j];
          const dist = this._distance.calculateNormalized(pxi, pxj);
          if (dist < thold) {
            memDist.push(new RemovedColor(j, pxj, dist));
            usageArray[j] = 0;
            palLen--;
          }
        }
      }
      thold += palLen > this._colors * 3 ? this._initialDistance : this._distanceIncrement;
    }
    if (palLen < this._colors) {
      stableSort(memDist, (a, b) => b.distance - a.distance);
      let k = 0;
      while (palLen < this._colors && k < memDist.length) {
        const removedColor = memDist[k];
        usageArray[removedColor.index] = 1;
        palLen++;
        k++;
      }
    }
    let colors = colorArray.length;
    for (let colorIndex = colors - 1; colorIndex >= 0; colorIndex--) {
      if (usageArray[colorIndex] === 0) {
        if (colorIndex !== colors - 1) {
          colorArray[colorIndex] = colorArray[colors - 1];
        }
        --colors;
      }
    }
    colorArray.length = colors;
    palette.sort();
    yield {
      palette,
      progress: 100
    };
  }
};

// src/palette/wu/wuQuant.ts
function createArray1D(dimension1) {
  const a = [];
  for (let k = 0; k < dimension1; k++) {
    a[k] = 0;
  }
  return a;
}
function createArray4D(dimension1, dimension2, dimension3, dimension4) {
  const a = new Array(dimension1);
  for (let i = 0; i < dimension1; i++) {
    a[i] = new Array(dimension2);
    for (let j = 0; j < dimension2; j++) {
      a[i][j] = new Array(dimension3);
      for (let k = 0; k < dimension3; k++) {
        a[i][j][k] = new Array(dimension4);
        for (let l = 0; l < dimension4; l++) {
          a[i][j][k][l] = 0;
        }
      }
    }
  }
  return a;
}
function createArray3D(dimension1, dimension2, dimension3) {
  const a = new Array(dimension1);
  for (let i = 0; i < dimension1; i++) {
    a[i] = new Array(dimension2);
    for (let j = 0; j < dimension2; j++) {
      a[i][j] = new Array(dimension3);
      for (let k = 0; k < dimension3; k++) {
        a[i][j][k] = 0;
      }
    }
  }
  return a;
}
function fillArray3D(a, dimension1, dimension2, dimension3, value) {
  for (let i = 0; i < dimension1; i++) {
    a[i] = [];
    for (let j = 0; j < dimension2; j++) {
      a[i][j] = [];
      for (let k = 0; k < dimension3; k++) {
        a[i][j][k] = value;
      }
    }
  }
}
function fillArray1D(a, dimension1, value) {
  for (let i = 0; i < dimension1; i++) {
    a[i] = value;
  }
}
var WuColorCube = class {
  constructor() {
    __publicField(this, "redMinimum");
    __publicField(this, "redMaximum");
    __publicField(this, "greenMinimum");
    __publicField(this, "greenMaximum");
    __publicField(this, "blueMinimum");
    __publicField(this, "blueMaximum");
    __publicField(this, "volume");
    __publicField(this, "alphaMinimum");
    __publicField(this, "alphaMaximum");
  }
};
var _WuQuant = class extends AbstractPaletteQuantizer {
  constructor(colorDistanceCalculator, colors = 256, significantBitsPerChannel = 5) {
    super();
    __publicField(this, "_reds");
    __publicField(this, "_greens");
    __publicField(this, "_blues");
    __publicField(this, "_alphas");
    __publicField(this, "_sums");
    __publicField(this, "_weights");
    __publicField(this, "_momentsRed");
    __publicField(this, "_momentsGreen");
    __publicField(this, "_momentsBlue");
    __publicField(this, "_momentsAlpha");
    __publicField(this, "_moments");
    __publicField(this, "_table");
    __publicField(this, "_pixels");
    __publicField(this, "_cubes");
    __publicField(this, "_colors");
    __publicField(this, "_significantBitsPerChannel");
    __publicField(this, "_maxSideIndex");
    __publicField(this, "_alphaMaxSideIndex");
    __publicField(this, "_sideSize");
    __publicField(this, "_alphaSideSize");
    __publicField(this, "_distance");
    this._distance = colorDistanceCalculator;
    this._setQuality(significantBitsPerChannel);
    this._initialize(colors);
  }
  sample(image) {
    const pointArray = image.getPointArray();
    for (let i = 0, l = pointArray.length; i < l; i++) {
      this._addColor(pointArray[i]);
    }
    this._pixels = this._pixels.concat(pointArray);
  }
  *quantize() {
    yield* this._preparePalette();
    const palette = new Palette();
    for (let paletteIndex = 0; paletteIndex < this._colors; paletteIndex++) {
      if (this._sums[paletteIndex] > 0) {
        const sum = this._sums[paletteIndex];
        const r = this._reds[paletteIndex] / sum;
        const g = this._greens[paletteIndex] / sum;
        const b = this._blues[paletteIndex] / sum;
        const a = this._alphas[paletteIndex] / sum;
        const color = Point.createByRGBA(r | 0, g | 0, b | 0, a | 0);
        palette.add(color);
      }
    }
    palette.sort();
    yield {
      palette,
      progress: 100
    };
  }
  *_preparePalette() {
    yield* this._calculateMoments();
    let next = 0;
    const volumeVariance = createArray1D(this._colors);
    for (let cubeIndex = 1; cubeIndex < this._colors; ++cubeIndex) {
      if (this._cut(this._cubes[next], this._cubes[cubeIndex])) {
        volumeVariance[next] = this._cubes[next].volume > 1 ? this._calculateVariance(this._cubes[next]) : 0;
        volumeVariance[cubeIndex] = this._cubes[cubeIndex].volume > 1 ? this._calculateVariance(this._cubes[cubeIndex]) : 0;
      } else {
        volumeVariance[next] = 0;
        cubeIndex--;
      }
      next = 0;
      let temp = volumeVariance[0];
      for (let index = 1; index <= cubeIndex; ++index) {
        if (volumeVariance[index] > temp) {
          temp = volumeVariance[index];
          next = index;
        }
      }
      if (temp <= 0) {
        this._colors = cubeIndex + 1;
        break;
      }
    }
    const lookupRed = [];
    const lookupGreen = [];
    const lookupBlue = [];
    const lookupAlpha = [];
    for (let k = 0; k < this._colors; ++k) {
      const weight = _WuQuant._volume(this._cubes[k], this._weights);
      if (weight > 0) {
        lookupRed[k] = _WuQuant._volume(this._cubes[k], this._momentsRed) / weight | 0;
        lookupGreen[k] = _WuQuant._volume(this._cubes[k], this._momentsGreen) / weight | 0;
        lookupBlue[k] = _WuQuant._volume(this._cubes[k], this._momentsBlue) / weight | 0;
        lookupAlpha[k] = _WuQuant._volume(this._cubes[k], this._momentsAlpha) / weight | 0;
      } else {
        lookupRed[k] = 0;
        lookupGreen[k] = 0;
        lookupBlue[k] = 0;
        lookupAlpha[k] = 0;
      }
    }
    this._reds = createArray1D(this._colors + 1);
    this._greens = createArray1D(this._colors + 1);
    this._blues = createArray1D(this._colors + 1);
    this._alphas = createArray1D(this._colors + 1);
    this._sums = createArray1D(this._colors + 1);
    for (let index = 0, l = this._pixels.length; index < l; index++) {
      const color = this._pixels[index];
      const match = -1;
      let bestMatch = match;
      let bestDistance = Number.MAX_VALUE;
      for (let lookup = 0; lookup < this._colors; lookup++) {
        const foundRed = lookupRed[lookup];
        const foundGreen = lookupGreen[lookup];
        const foundBlue = lookupBlue[lookup];
        const foundAlpha = lookupAlpha[lookup];
        const distance = this._distance.calculateRaw(foundRed, foundGreen, foundBlue, foundAlpha, color.r, color.g, color.b, color.a);
        if (distance < bestDistance) {
          bestDistance = distance;
          bestMatch = lookup;
        }
      }
      this._reds[bestMatch] += color.r;
      this._greens[bestMatch] += color.g;
      this._blues[bestMatch] += color.b;
      this._alphas[bestMatch] += color.a;
      this._sums[bestMatch]++;
    }
  }
  _addColor(color) {
    const bitsToRemove = 8 - this._significantBitsPerChannel;
    const indexRed = (color.r >> bitsToRemove) + 1;
    const indexGreen = (color.g >> bitsToRemove) + 1;
    const indexBlue = (color.b >> bitsToRemove) + 1;
    const indexAlpha = (color.a >> bitsToRemove) + 1;
    this._weights[indexAlpha][indexRed][indexGreen][indexBlue]++;
    this._momentsRed[indexAlpha][indexRed][indexGreen][indexBlue] += color.r;
    this._momentsGreen[indexAlpha][indexRed][indexGreen][indexBlue] += color.g;
    this._momentsBlue[indexAlpha][indexRed][indexGreen][indexBlue] += color.b;
    this._momentsAlpha[indexAlpha][indexRed][indexGreen][indexBlue] += color.a;
    this._moments[indexAlpha][indexRed][indexGreen][indexBlue] += this._table[color.r] + this._table[color.g] + this._table[color.b] + this._table[color.a];
  }
  *_calculateMoments() {
    const area = [];
    const areaRed = [];
    const areaGreen = [];
    const areaBlue = [];
    const areaAlpha = [];
    const area2 = [];
    const xarea = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    const xareaRed = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    const xareaGreen = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    const xareaBlue = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    const xareaAlpha = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    const xarea2 = createArray3D(this._sideSize, this._sideSize, this._sideSize);
    let trackerProgress = 0;
    const tracker = new ProgressTracker(this._alphaMaxSideIndex * this._maxSideIndex, 99);
    for (let alphaIndex = 1; alphaIndex <= this._alphaMaxSideIndex; ++alphaIndex) {
      fillArray3D(xarea, this._sideSize, this._sideSize, this._sideSize, 0);
      fillArray3D(xareaRed, this._sideSize, this._sideSize, this._sideSize, 0);
      fillArray3D(xareaGreen, this._sideSize, this._sideSize, this._sideSize, 0);
      fillArray3D(xareaBlue, this._sideSize, this._sideSize, this._sideSize, 0);
      fillArray3D(xareaAlpha, this._sideSize, this._sideSize, this._sideSize, 0);
      fillArray3D(xarea2, this._sideSize, this._sideSize, this._sideSize, 0);
      for (let redIndex = 1; redIndex <= this._maxSideIndex; ++redIndex, ++trackerProgress) {
        if (tracker.shouldNotify(trackerProgress)) {
          yield {
            progress: tracker.progress
          };
        }
        fillArray1D(area, this._sideSize, 0);
        fillArray1D(areaRed, this._sideSize, 0);
        fillArray1D(areaGreen, this._sideSize, 0);
        fillArray1D(areaBlue, this._sideSize, 0);
        fillArray1D(areaAlpha, this._sideSize, 0);
        fillArray1D(area2, this._sideSize, 0);
        for (let greenIndex = 1; greenIndex <= this._maxSideIndex; ++greenIndex) {
          let line = 0;
          let lineRed = 0;
          let lineGreen = 0;
          let lineBlue = 0;
          let lineAlpha = 0;
          let line2 = 0;
          for (let blueIndex = 1; blueIndex <= this._maxSideIndex; ++blueIndex) {
            line += this._weights[alphaIndex][redIndex][greenIndex][blueIndex];
            lineRed += this._momentsRed[alphaIndex][redIndex][greenIndex][blueIndex];
            lineGreen += this._momentsGreen[alphaIndex][redIndex][greenIndex][blueIndex];
            lineBlue += this._momentsBlue[alphaIndex][redIndex][greenIndex][blueIndex];
            lineAlpha += this._momentsAlpha[alphaIndex][redIndex][greenIndex][blueIndex];
            line2 += this._moments[alphaIndex][redIndex][greenIndex][blueIndex];
            area[blueIndex] += line;
            areaRed[blueIndex] += lineRed;
            areaGreen[blueIndex] += lineGreen;
            areaBlue[blueIndex] += lineBlue;
            areaAlpha[blueIndex] += lineAlpha;
            area2[blueIndex] += line2;
            xarea[redIndex][greenIndex][blueIndex] = xarea[redIndex - 1][greenIndex][blueIndex] + area[blueIndex];
            xareaRed[redIndex][greenIndex][blueIndex] = xareaRed[redIndex - 1][greenIndex][blueIndex] + areaRed[blueIndex];
            xareaGreen[redIndex][greenIndex][blueIndex] = xareaGreen[redIndex - 1][greenIndex][blueIndex] + areaGreen[blueIndex];
            xareaBlue[redIndex][greenIndex][blueIndex] = xareaBlue[redIndex - 1][greenIndex][blueIndex] + areaBlue[blueIndex];
            xareaAlpha[redIndex][greenIndex][blueIndex] = xareaAlpha[redIndex - 1][greenIndex][blueIndex] + areaAlpha[blueIndex];
            xarea2[redIndex][greenIndex][blueIndex] = xarea2[redIndex - 1][greenIndex][blueIndex] + area2[blueIndex];
            this._weights[alphaIndex][redIndex][greenIndex][blueIndex] = this._weights[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xarea[redIndex][greenIndex][blueIndex];
            this._momentsRed[alphaIndex][redIndex][greenIndex][blueIndex] = this._momentsRed[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xareaRed[redIndex][greenIndex][blueIndex];
            this._momentsGreen[alphaIndex][redIndex][greenIndex][blueIndex] = this._momentsGreen[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xareaGreen[redIndex][greenIndex][blueIndex];
            this._momentsBlue[alphaIndex][redIndex][greenIndex][blueIndex] = this._momentsBlue[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xareaBlue[redIndex][greenIndex][blueIndex];
            this._momentsAlpha[alphaIndex][redIndex][greenIndex][blueIndex] = this._momentsAlpha[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xareaAlpha[redIndex][greenIndex][blueIndex];
            this._moments[alphaIndex][redIndex][greenIndex][blueIndex] = this._moments[alphaIndex - 1][redIndex][greenIndex][blueIndex] + xarea2[redIndex][greenIndex][blueIndex];
          }
        }
      }
    }
  }
  static _volumeFloat(cube, moment) {
    return moment[cube.alphaMaximum][cube.redMaximum][cube.greenMaximum][cube.blueMaximum] - moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] - moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - (moment[cube.alphaMaximum][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] - moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
  }
  static _volume(cube, moment) {
    return _WuQuant._volumeFloat(cube, moment) | 0;
  }
  static _top(cube, direction, position, moment) {
    let result;
    switch (direction) {
      case _WuQuant._alpha:
        result = moment[position][cube.redMaximum][cube.greenMaximum][cube.blueMaximum] - moment[position][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] - moment[position][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] + moment[position][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - (moment[position][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] - moment[position][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] - moment[position][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] + moment[position][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
        break;
      case _WuQuant._red:
        result = moment[cube.alphaMaximum][position][cube.greenMaximum][cube.blueMaximum] - moment[cube.alphaMaximum][position][cube.greenMinimum][cube.blueMaximum] - moment[cube.alphaMinimum][position][cube.greenMaximum][cube.blueMaximum] + moment[cube.alphaMinimum][position][cube.greenMinimum][cube.blueMaximum] - (moment[cube.alphaMaximum][position][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMaximum][position][cube.greenMinimum][cube.blueMinimum] - moment[cube.alphaMinimum][position][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMinimum][position][cube.greenMinimum][cube.blueMinimum]);
        break;
      case _WuQuant._green:
        result = moment[cube.alphaMaximum][cube.redMaximum][position][cube.blueMaximum] - moment[cube.alphaMaximum][cube.redMinimum][position][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMaximum][position][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMinimum][position][cube.blueMaximum] - (moment[cube.alphaMaximum][cube.redMaximum][position][cube.blueMinimum] - moment[cube.alphaMaximum][cube.redMinimum][position][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMaximum][position][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMinimum][position][cube.blueMinimum]);
        break;
      case _WuQuant._blue:
        result = moment[cube.alphaMaximum][cube.redMaximum][cube.greenMaximum][position] - moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][position] - moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][position] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][position] - (moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][position] - moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][position] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][position] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][position]);
        break;
      default:
        throw new Error("impossible");
    }
    return result | 0;
  }
  static _bottom(cube, direction, moment) {
    switch (direction) {
      case _WuQuant._alpha:
        return -moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - (-moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
      case _WuQuant._red:
        return -moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - (-moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
      case _WuQuant._green:
        return -moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMaximum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMaximum] - (-moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
      case _WuQuant._blue:
        return -moment[cube.alphaMaximum][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMaximum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMaximum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMaximum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum] - (-moment[cube.alphaMinimum][cube.redMaximum][cube.greenMaximum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMaximum][cube.greenMinimum][cube.blueMinimum] + moment[cube.alphaMinimum][cube.redMinimum][cube.greenMaximum][cube.blueMinimum] - moment[cube.alphaMinimum][cube.redMinimum][cube.greenMinimum][cube.blueMinimum]);
      default:
        return 0;
    }
  }
  _calculateVariance(cube) {
    const volumeRed = _WuQuant._volume(cube, this._momentsRed);
    const volumeGreen = _WuQuant._volume(cube, this._momentsGreen);
    const volumeBlue = _WuQuant._volume(cube, this._momentsBlue);
    const volumeAlpha = _WuQuant._volume(cube, this._momentsAlpha);
    const volumeMoment = _WuQuant._volumeFloat(cube, this._moments);
    const volumeWeight = _WuQuant._volume(cube, this._weights);
    const distance = volumeRed * volumeRed + volumeGreen * volumeGreen + volumeBlue * volumeBlue + volumeAlpha * volumeAlpha;
    return volumeMoment - distance / volumeWeight;
  }
  _maximize(cube, direction, first, last, wholeRed, wholeGreen, wholeBlue, wholeAlpha, wholeWeight) {
    const bottomRed = _WuQuant._bottom(cube, direction, this._momentsRed) | 0;
    const bottomGreen = _WuQuant._bottom(cube, direction, this._momentsGreen) | 0;
    const bottomBlue = _WuQuant._bottom(cube, direction, this._momentsBlue) | 0;
    const bottomAlpha = _WuQuant._bottom(cube, direction, this._momentsAlpha) | 0;
    const bottomWeight = _WuQuant._bottom(cube, direction, this._weights) | 0;
    let result = 0;
    let cutPosition = -1;
    for (let position = first; position < last; ++position) {
      let halfRed = bottomRed + _WuQuant._top(cube, direction, position, this._momentsRed);
      let halfGreen = bottomGreen + _WuQuant._top(cube, direction, position, this._momentsGreen);
      let halfBlue = bottomBlue + _WuQuant._top(cube, direction, position, this._momentsBlue);
      let halfAlpha = bottomAlpha + _WuQuant._top(cube, direction, position, this._momentsAlpha);
      let halfWeight = bottomWeight + _WuQuant._top(cube, direction, position, this._weights);
      if (halfWeight !== 0) {
        let halfDistance = halfRed * halfRed + halfGreen * halfGreen + halfBlue * halfBlue + halfAlpha * halfAlpha;
        let temp = halfDistance / halfWeight;
        halfRed = wholeRed - halfRed;
        halfGreen = wholeGreen - halfGreen;
        halfBlue = wholeBlue - halfBlue;
        halfAlpha = wholeAlpha - halfAlpha;
        halfWeight = wholeWeight - halfWeight;
        if (halfWeight !== 0) {
          halfDistance = halfRed * halfRed + halfGreen * halfGreen + halfBlue * halfBlue + halfAlpha * halfAlpha;
          temp += halfDistance / halfWeight;
          if (temp > result) {
            result = temp;
            cutPosition = position;
          }
        }
      }
    }
    return { max: result, position: cutPosition };
  }
  _cut(first, second) {
    let direction;
    const wholeRed = _WuQuant._volume(first, this._momentsRed);
    const wholeGreen = _WuQuant._volume(first, this._momentsGreen);
    const wholeBlue = _WuQuant._volume(first, this._momentsBlue);
    const wholeAlpha = _WuQuant._volume(first, this._momentsAlpha);
    const wholeWeight = _WuQuant._volume(first, this._weights);
    const red = this._maximize(first, _WuQuant._red, first.redMinimum + 1, first.redMaximum, wholeRed, wholeGreen, wholeBlue, wholeAlpha, wholeWeight);
    const green = this._maximize(first, _WuQuant._green, first.greenMinimum + 1, first.greenMaximum, wholeRed, wholeGreen, wholeBlue, wholeAlpha, wholeWeight);
    const blue = this._maximize(first, _WuQuant._blue, first.blueMinimum + 1, first.blueMaximum, wholeRed, wholeGreen, wholeBlue, wholeAlpha, wholeWeight);
    const alpha = this._maximize(first, _WuQuant._alpha, first.alphaMinimum + 1, first.alphaMaximum, wholeRed, wholeGreen, wholeBlue, wholeAlpha, wholeWeight);
    if (alpha.max >= red.max && alpha.max >= green.max && alpha.max >= blue.max) {
      direction = _WuQuant._alpha;
      if (alpha.position < 0)
        return false;
    } else if (red.max >= alpha.max && red.max >= green.max && red.max >= blue.max) {
      direction = _WuQuant._red;
    } else if (green.max >= alpha.max && green.max >= red.max && green.max >= blue.max) {
      direction = _WuQuant._green;
    } else {
      direction = _WuQuant._blue;
    }
    second.redMaximum = first.redMaximum;
    second.greenMaximum = first.greenMaximum;
    second.blueMaximum = first.blueMaximum;
    second.alphaMaximum = first.alphaMaximum;
    switch (direction) {
      case _WuQuant._red:
        second.redMinimum = first.redMaximum = red.position;
        second.greenMinimum = first.greenMinimum;
        second.blueMinimum = first.blueMinimum;
        second.alphaMinimum = first.alphaMinimum;
        break;
      case _WuQuant._green:
        second.greenMinimum = first.greenMaximum = green.position;
        second.redMinimum = first.redMinimum;
        second.blueMinimum = first.blueMinimum;
        second.alphaMinimum = first.alphaMinimum;
        break;
      case _WuQuant._blue:
        second.blueMinimum = first.blueMaximum = blue.position;
        second.redMinimum = first.redMinimum;
        second.greenMinimum = first.greenMinimum;
        second.alphaMinimum = first.alphaMinimum;
        break;
      case _WuQuant._alpha:
        second.alphaMinimum = first.alphaMaximum = alpha.position;
        second.blueMinimum = first.blueMinimum;
        second.redMinimum = first.redMinimum;
        second.greenMinimum = first.greenMinimum;
        break;
    }
    first.volume = (first.redMaximum - first.redMinimum) * (first.greenMaximum - first.greenMinimum) * (first.blueMaximum - first.blueMinimum) * (first.alphaMaximum - first.alphaMinimum);
    second.volume = (second.redMaximum - second.redMinimum) * (second.greenMaximum - second.greenMinimum) * (second.blueMaximum - second.blueMinimum) * (second.alphaMaximum - second.alphaMinimum);
    return true;
  }
  _initialize(colors) {
    this._colors = colors;
    this._cubes = [];
    for (let cubeIndex = 0; cubeIndex < colors; cubeIndex++) {
      this._cubes[cubeIndex] = new WuColorCube();
    }
    this._cubes[0].redMinimum = 0;
    this._cubes[0].greenMinimum = 0;
    this._cubes[0].blueMinimum = 0;
    this._cubes[0].alphaMinimum = 0;
    this._cubes[0].redMaximum = this._maxSideIndex;
    this._cubes[0].greenMaximum = this._maxSideIndex;
    this._cubes[0].blueMaximum = this._maxSideIndex;
    this._cubes[0].alphaMaximum = this._alphaMaxSideIndex;
    this._weights = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._momentsRed = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._momentsGreen = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._momentsBlue = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._momentsAlpha = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._moments = createArray4D(this._alphaSideSize, this._sideSize, this._sideSize, this._sideSize);
    this._table = [];
    for (let tableIndex = 0; tableIndex < 256; ++tableIndex) {
      this._table[tableIndex] = tableIndex * tableIndex;
    }
    this._pixels = [];
  }
  _setQuality(significantBitsPerChannel = 5) {
    this._significantBitsPerChannel = significantBitsPerChannel;
    this._maxSideIndex = 1 << this._significantBitsPerChannel;
    this._alphaMaxSideIndex = this._maxSideIndex;
    this._sideSize = this._maxSideIndex + 1;
    this._alphaSideSize = this._alphaMaxSideIndex + 1;
  }
};
var WuQuant = _WuQuant;
__publicField(WuQuant, "_alpha", 3);
__publicField(WuQuant, "_red", 2);
__publicField(WuQuant, "_green", 1);
__publicField(WuQuant, "_blue", 0);

// src/image/index.ts
var image_exports = {};
__export(image_exports, {
  AbstractImageQuantizer: () => AbstractImageQuantizer,
  ErrorDiffusionArray: () => ErrorDiffusionArray,
  ErrorDiffusionArrayKernel: () => ErrorDiffusionArrayKernel,
  ErrorDiffusionRiemersma: () => ErrorDiffusionRiemersma,
  NearestColor: () => NearestColor
});

// src/image/imageQuantizer.ts
var AbstractImageQuantizer = class {
  quantizeSync(pointContainer, palette) {
    for (const value of this.quantize(pointContainer, palette)) {
      if (value.pointContainer) {
        return value.pointContainer;
      }
    }
    throw new Error("unreachable");
  }
};

// src/image/nearestColor.ts
var NearestColor = class extends AbstractImageQuantizer {
  constructor(colorDistanceCalculator) {
    super();
    __publicField(this, "_distance");
    this._distance = colorDistanceCalculator;
  }
  *quantize(pointContainer, palette) {
    const pointArray = pointContainer.getPointArray();
    const width = pointContainer.getWidth();
    const height = pointContainer.getHeight();
    const tracker = new ProgressTracker(height, 99);
    for (let y2 = 0; y2 < height; y2++) {
      if (tracker.shouldNotify(y2)) {
        yield {
          progress: tracker.progress
        };
      }
      for (let x2 = 0, idx = y2 * width; x2 < width; x2++, idx++) {
        const point = pointArray[idx];
        point.from(palette.getNearestColor(this._distance, point));
      }
    }
    yield {
      pointContainer,
      progress: 100
    };
  }
};

// src/image/array.ts
var ErrorDiffusionArrayKernel = /* @__PURE__ */ ((ErrorDiffusionArrayKernel2) => {
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["FloydSteinberg"] = 0] = "FloydSteinberg";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["FalseFloydSteinberg"] = 1] = "FalseFloydSteinberg";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["Stucki"] = 2] = "Stucki";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["Atkinson"] = 3] = "Atkinson";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["Jarvis"] = 4] = "Jarvis";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["Burkes"] = 5] = "Burkes";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["Sierra"] = 6] = "Sierra";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["TwoSierra"] = 7] = "TwoSierra";
  ErrorDiffusionArrayKernel2[ErrorDiffusionArrayKernel2["SierraLite"] = 8] = "SierraLite";
  return ErrorDiffusionArrayKernel2;
})(ErrorDiffusionArrayKernel || {});
var ErrorDiffusionArray = class extends AbstractImageQuantizer {
  constructor(colorDistanceCalculator, kernel, serpentine = true, minimumColorDistanceToDither = 0, calculateErrorLikeGIMP = false) {
    super();
    __publicField(this, "_minColorDistance");
    __publicField(this, "_serpentine");
    __publicField(this, "_kernel");
    __publicField(this, "_calculateErrorLikeGIMP");
    __publicField(this, "_distance");
    this._setKernel(kernel);
    this._distance = colorDistanceCalculator;
    this._minColorDistance = minimumColorDistanceToDither;
    this._serpentine = serpentine;
    this._calculateErrorLikeGIMP = calculateErrorLikeGIMP;
  }
  *quantize(pointContainer, palette) {
    const pointArray = pointContainer.getPointArray();
    const originalPoint = new Point();
    const width = pointContainer.getWidth();
    const height = pointContainer.getHeight();
    const errorLines = [];
    let dir = 1;
    let maxErrorLines = 1;
    for (const kernel of this._kernel) {
      const kernelErrorLines = kernel[2] + 1;
      if (maxErrorLines < kernelErrorLines)
        maxErrorLines = kernelErrorLines;
    }
    for (let i = 0; i < maxErrorLines; i++) {
      this._fillErrorLine(errorLines[i] = [], width);
    }
    const tracker = new ProgressTracker(height, 99);
    for (let y2 = 0; y2 < height; y2++) {
      if (tracker.shouldNotify(y2)) {
        yield {
          progress: tracker.progress
        };
      }
      if (this._serpentine)
        dir *= -1;
      const lni = y2 * width;
      const xStart = dir === 1 ? 0 : width - 1;
      const xEnd = dir === 1 ? width : -1;
      this._fillErrorLine(errorLines[0], width);
      errorLines.push(errorLines.shift());
      const errorLine = errorLines[0];
      for (let x2 = xStart, idx = lni + xStart; x2 !== xEnd; x2 += dir, idx += dir) {
        const point = pointArray[idx];
        const error = errorLine[x2];
        originalPoint.from(point);
        const correctedPoint = Point.createByRGBA(inRange0to255Rounded(point.r + error[0]), inRange0to255Rounded(point.g + error[1]), inRange0to255Rounded(point.b + error[2]), inRange0to255Rounded(point.a + error[3]));
        const palettePoint = palette.getNearestColor(this._distance, correctedPoint);
        point.from(palettePoint);
        if (this._minColorDistance) {
          const dist = this._distance.calculateNormalized(originalPoint, palettePoint);
          if (dist < this._minColorDistance)
            continue;
        }
        let er;
        let eg;
        let eb;
        let ea;
        if (this._calculateErrorLikeGIMP) {
          er = correctedPoint.r - palettePoint.r;
          eg = correctedPoint.g - palettePoint.g;
          eb = correctedPoint.b - palettePoint.b;
          ea = correctedPoint.a - palettePoint.a;
        } else {
          er = originalPoint.r - palettePoint.r;
          eg = originalPoint.g - palettePoint.g;
          eb = originalPoint.b - palettePoint.b;
          ea = originalPoint.a - palettePoint.a;
        }
        const dStart = dir === 1 ? 0 : this._kernel.length - 1;
        const dEnd = dir === 1 ? this._kernel.length : -1;
        for (let i = dStart; i !== dEnd; i += dir) {
          const x1 = this._kernel[i][1] * dir;
          const y1 = this._kernel[i][2];
          if (x1 + x2 >= 0 && x1 + x2 < width && y1 + y2 >= 0 && y1 + y2 < height) {
            const d = this._kernel[i][0];
            const e = errorLines[y1][x1 + x2];
            e[0] += er * d;
            e[1] += eg * d;
            e[2] += eb * d;
            e[3] += ea * d;
          }
        }
      }
    }
    yield {
      pointContainer,
      progress: 100
    };
  }
  _fillErrorLine(errorLine, width) {
    if (errorLine.length > width) {
      errorLine.length = width;
    }
    const l = errorLine.length;
    for (let i = 0; i < l; i++) {
      const error = errorLine[i];
      error[0] = error[1] = error[2] = error[3] = 0;
    }
    for (let i = l; i < width; i++) {
      errorLine[i] = [0, 0, 0, 0];
    }
  }
  _setKernel(kernel) {
    switch (kernel) {
      case 0 /* FloydSteinberg */:
        this._kernel = [
          [7 / 16, 1, 0],
          [3 / 16, -1, 1],
          [5 / 16, 0, 1],
          [1 / 16, 1, 1]
        ];
        break;
      case 1 /* FalseFloydSteinberg */:
        this._kernel = [
          [3 / 8, 1, 0],
          [3 / 8, 0, 1],
          [2 / 8, 1, 1]
        ];
        break;
      case 2 /* Stucki */:
        this._kernel = [
          [8 / 42, 1, 0],
          [4 / 42, 2, 0],
          [2 / 42, -2, 1],
          [4 / 42, -1, 1],
          [8 / 42, 0, 1],
          [4 / 42, 1, 1],
          [2 / 42, 2, 1],
          [1 / 42, -2, 2],
          [2 / 42, -1, 2],
          [4 / 42, 0, 2],
          [2 / 42, 1, 2],
          [1 / 42, 2, 2]
        ];
        break;
      case 3 /* Atkinson */:
        this._kernel = [
          [1 / 8, 1, 0],
          [1 / 8, 2, 0],
          [1 / 8, -1, 1],
          [1 / 8, 0, 1],
          [1 / 8, 1, 1],
          [1 / 8, 0, 2]
        ];
        break;
      case 4 /* Jarvis */:
        this._kernel = [
          [7 / 48, 1, 0],
          [5 / 48, 2, 0],
          [3 / 48, -2, 1],
          [5 / 48, -1, 1],
          [7 / 48, 0, 1],
          [5 / 48, 1, 1],
          [3 / 48, 2, 1],
          [1 / 48, -2, 2],
          [3 / 48, -1, 2],
          [5 / 48, 0, 2],
          [3 / 48, 1, 2],
          [1 / 48, 2, 2]
        ];
        break;
      case 5 /* Burkes */:
        this._kernel = [
          [8 / 32, 1, 0],
          [4 / 32, 2, 0],
          [2 / 32, -2, 1],
          [4 / 32, -1, 1],
          [8 / 32, 0, 1],
          [4 / 32, 1, 1],
          [2 / 32, 2, 1]
        ];
        break;
      case 6 /* Sierra */:
        this._kernel = [
          [5 / 32, 1, 0],
          [3 / 32, 2, 0],
          [2 / 32, -2, 1],
          [4 / 32, -1, 1],
          [5 / 32, 0, 1],
          [4 / 32, 1, 1],
          [2 / 32, 2, 1],
          [2 / 32, -1, 2],
          [3 / 32, 0, 2],
          [2 / 32, 1, 2]
        ];
        break;
      case 7 /* TwoSierra */:
        this._kernel = [
          [4 / 16, 1, 0],
          [3 / 16, 2, 0],
          [1 / 16, -2, 1],
          [2 / 16, -1, 1],
          [3 / 16, 0, 1],
          [2 / 16, 1, 1],
          [1 / 16, 2, 1]
        ];
        break;
      case 8 /* SierraLite */:
        this._kernel = [
          [2 / 4, 1, 0],
          [1 / 4, -1, 1],
          [1 / 4, 0, 1]
        ];
        break;
      default:
        throw new Error(`ErrorDiffusionArray: unknown kernel = ${kernel}`);
    }
  }
};

// src/image/spaceFillingCurves/hilbertCurve.ts
function* hilbertCurve(width, height, callback) {
  const maxBound = Math.max(width, height);
  const level = Math.floor(Math.log(maxBound) / Math.log(2) + 1);
  const tracker = new ProgressTracker(width * height, 99);
  const data = {
    width,
    height,
    level,
    callback,
    tracker,
    index: 0,
    x: 0,
    y: 0
  };
  yield* walkHilbert(data, 1 /* UP */);
  visit(data, 0 /* NONE */);
}
function* walkHilbert(data, direction) {
  if (data.level < 1)
    return;
  if (data.tracker.shouldNotify(data.index)) {
    yield { progress: data.tracker.progress };
  }
  data.level--;
  switch (direction) {
    case 2 /* LEFT */:
      yield* walkHilbert(data, 1 /* UP */);
      visit(data, 3 /* RIGHT */);
      yield* walkHilbert(data, 2 /* LEFT */);
      visit(data, 4 /* DOWN */);
      yield* walkHilbert(data, 2 /* LEFT */);
      visit(data, 2 /* LEFT */);
      yield* walkHilbert(data, 4 /* DOWN */);
      break;
    case 3 /* RIGHT */:
      yield* walkHilbert(data, 4 /* DOWN */);
      visit(data, 2 /* LEFT */);
      yield* walkHilbert(data, 3 /* RIGHT */);
      visit(data, 1 /* UP */);
      yield* walkHilbert(data, 3 /* RIGHT */);
      visit(data, 3 /* RIGHT */);
      yield* walkHilbert(data, 1 /* UP */);
      break;
    case 1 /* UP */:
      yield* walkHilbert(data, 2 /* LEFT */);
      visit(data, 4 /* DOWN */);
      yield* walkHilbert(data, 1 /* UP */);
      visit(data, 3 /* RIGHT */);
      yield* walkHilbert(data, 1 /* UP */);
      visit(data, 1 /* UP */);
      yield* walkHilbert(data, 3 /* RIGHT */);
      break;
    case 4 /* DOWN */:
      yield* walkHilbert(data, 3 /* RIGHT */);
      visit(data, 1 /* UP */);
      yield* walkHilbert(data, 4 /* DOWN */);
      visit(data, 2 /* LEFT */);
      yield* walkHilbert(data, 4 /* DOWN */);
      visit(data, 4 /* DOWN */);
      yield* walkHilbert(data, 2 /* LEFT */);
      break;
    default:
      break;
  }
  data.level++;
}
function visit(data, direction) {
  if (data.x >= 0 && data.x < data.width && data.y >= 0 && data.y < data.height) {
    data.callback(data.x, data.y);
    data.index++;
  }
  switch (direction) {
    case 2 /* LEFT */:
      data.x--;
      break;
    case 3 /* RIGHT */:
      data.x++;
      break;
    case 1 /* UP */:
      data.y--;
      break;
    case 4 /* DOWN */:
      data.y++;
      break;
  }
}

// src/image/riemersma.ts
var ErrorDiffusionRiemersma = class extends AbstractImageQuantizer {
  constructor(colorDistanceCalculator, errorQueueSize = 16, errorPropagation = 1) {
    super();
    __publicField(this, "_distance");
    __publicField(this, "_weights");
    __publicField(this, "_errorQueueSize");
    this._distance = colorDistanceCalculator;
    this._errorQueueSize = errorQueueSize;
    this._weights = ErrorDiffusionRiemersma._createWeights(errorPropagation, errorQueueSize);
  }
  *quantize(pointContainer, palette) {
    const pointArray = pointContainer.getPointArray();
    const width = pointContainer.getWidth();
    const height = pointContainer.getHeight();
    const errorQueue = [];
    let head = 0;
    for (let i = 0; i < this._errorQueueSize; i++) {
      errorQueue[i] = { r: 0, g: 0, b: 0, a: 0 };
    }
    yield* hilbertCurve(width, height, (x2, y2) => {
      const p = pointArray[x2 + y2 * width];
      let { r, g, b, a } = p;
      for (let i = 0; i < this._errorQueueSize; i++) {
        const weight = this._weights[i];
        const e = errorQueue[(i + head) % this._errorQueueSize];
        r += e.r * weight;
        g += e.g * weight;
        b += e.b * weight;
        a += e.a * weight;
      }
      const correctedPoint = Point.createByRGBA(inRange0to255Rounded(r), inRange0to255Rounded(g), inRange0to255Rounded(b), inRange0to255Rounded(a));
      const quantizedPoint = palette.getNearestColor(this._distance, correctedPoint);
      head = (head + 1) % this._errorQueueSize;
      const tail = (head + this._errorQueueSize - 1) % this._errorQueueSize;
      errorQueue[tail].r = p.r - quantizedPoint.r;
      errorQueue[tail].g = p.g - quantizedPoint.g;
      errorQueue[tail].b = p.b - quantizedPoint.b;
      errorQueue[tail].a = p.a - quantizedPoint.a;
      p.from(quantizedPoint);
    });
    yield {
      pointContainer,
      progress: 100
    };
  }
  static _createWeights(errorPropagation, errorQueueSize) {
    const weights = [];
    const multiplier = Math.exp(Math.log(errorQueueSize) / (errorQueueSize - 1));
    for (let i = 0, next = 1; i < errorQueueSize; i++) {
      weights[i] = (next + 0.5 | 0) / errorQueueSize * errorPropagation;
      next *= multiplier;
    }
    return weights;
  }
};

// src/quality/index.ts
var quality_exports = {};
__export(quality_exports, {
  ssim: () => ssim
});

// src/quality/ssim.ts
var K1 = 0.01;
var K2 = 0.03;
function ssim(image1, image2) {
  if (image1.getHeight() !== image2.getHeight() || image1.getWidth() !== image2.getWidth()) {
    throw new Error("Images have different sizes!");
  }
  const bitsPerComponent = 8;
  const L = (1 << bitsPerComponent) - 1;
  const c1 = (K1 * L) ** 2;
  const c2 = (K2 * L) ** 2;
  let numWindows = 0;
  let mssim = 0;
  iterate(image1, image2, (lumaValues1, lumaValues2, averageLumaValue1, averageLumaValue2) => {
    let sigxy = 0;
    let sigsqx = 0;
    let sigsqy = 0;
    for (let i = 0; i < lumaValues1.length; i++) {
      sigsqx += (lumaValues1[i] - averageLumaValue1) ** 2;
      sigsqy += (lumaValues2[i] - averageLumaValue2) ** 2;
      sigxy += (lumaValues1[i] - averageLumaValue1) * (lumaValues2[i] - averageLumaValue2);
    }
    const numPixelsInWin = lumaValues1.length - 1;
    sigsqx /= numPixelsInWin;
    sigsqy /= numPixelsInWin;
    sigxy /= numPixelsInWin;
    const numerator = (2 * averageLumaValue1 * averageLumaValue2 + c1) * (2 * sigxy + c2);
    const denominator = (averageLumaValue1 ** 2 + averageLumaValue2 ** 2 + c1) * (sigsqx + sigsqy + c2);
    const ssim2 = numerator / denominator;
    mssim += ssim2;
    numWindows++;
  });
  return mssim / numWindows;
}
function iterate(image1, image2, callback) {
  const windowSize = 8;
  const width = image1.getWidth();
  const height = image1.getHeight();
  for (let y2 = 0; y2 < height; y2 += windowSize) {
    for (let x2 = 0; x2 < width; x2 += windowSize) {
      const windowWidth = Math.min(windowSize, width - x2);
      const windowHeight = Math.min(windowSize, height - y2);
      const lumaValues1 = calculateLumaValuesForWindow(image1, x2, y2, windowWidth, windowHeight);
      const lumaValues2 = calculateLumaValuesForWindow(image2, x2, y2, windowWidth, windowHeight);
      const averageLuma1 = calculateAverageLuma(lumaValues1);
      const averageLuma2 = calculateAverageLuma(lumaValues2);
      callback(lumaValues1, lumaValues2, averageLuma1, averageLuma2);
    }
  }
}
function calculateLumaValuesForWindow(image, x2, y2, width, height) {
  const pointArray = image.getPointArray();
  const lumaValues = [];
  let counter = 0;
  for (let j = y2; j < y2 + height; j++) {
    const offset = j * image.getWidth();
    for (let i = x2; i < x2 + width; i++) {
      const point = pointArray[offset + i];
      lumaValues[counter] = point.r * 0.2126 /* RED */ + point.g * 0.7152 /* GREEN */ + point.b * 0.0722 /* BLUE */;
      counter++;
    }
  }
  return lumaValues;
}
function calculateAverageLuma(lumaValues) {
  let sumLuma = 0;
  for (const luma of lumaValues) {
    sumLuma += luma;
  }
  return sumLuma / lumaValues.length;
}

// src/basicAPI.ts
var setImmediateImpl = typeof setImmediate === "function" ? setImmediate : typeof process !== "undefined" && typeof (process == null ? void 0 : process.nextTick) === "function" ? (callback) => process.nextTick(callback) : (callback) => setTimeout(callback, 0);
function buildPaletteSync(images, {
  colorDistanceFormula,
  paletteQuantization,
  colors
} = {}) {
  const distanceCalculator = colorDistanceFormulaToColorDistance(colorDistanceFormula);
  const paletteQuantizer = paletteQuantizationToPaletteQuantizer(distanceCalculator, paletteQuantization, colors);
  images.forEach((image) => paletteQuantizer.sample(image));
  return paletteQuantizer.quantizeSync();
}
async function buildPalette(images, {
  colorDistanceFormula,
  paletteQuantization,
  colors,
  onProgress
} = {}) {
  return new Promise((resolve, reject) => {
    const distanceCalculator = colorDistanceFormulaToColorDistance(colorDistanceFormula);
    const paletteQuantizer = paletteQuantizationToPaletteQuantizer(distanceCalculator, paletteQuantization, colors);
    images.forEach((image) => paletteQuantizer.sample(image));
    let palette;
    const iterator = paletteQuantizer.quantize();
    const next = () => {
      try {
        const result = iterator.next();
        if (result.done) {
          resolve(palette);
        } else {
          if (result.value.palette)
            palette = result.value.palette;
          if (onProgress)
            onProgress(result.value.progress);
          setImmediateImpl(next);
        }
      } catch (error) {
        reject(error);
      }
    };
    setImmediateImpl(next);
  });
}
function applyPaletteSync(image, palette, { colorDistanceFormula, imageQuantization } = {}) {
  const distanceCalculator = colorDistanceFormulaToColorDistance(colorDistanceFormula);
  const imageQuantizer = imageQuantizationToImageQuantizer(distanceCalculator, imageQuantization);
  return imageQuantizer.quantizeSync(image, palette);
}
async function applyPalette(image, palette, {
  colorDistanceFormula,
  imageQuantization,
  onProgress
} = {}) {
  return new Promise((resolve, reject) => {
    const distanceCalculator = colorDistanceFormulaToColorDistance(colorDistanceFormula);
    const imageQuantizer = imageQuantizationToImageQuantizer(distanceCalculator, imageQuantization);
    let outPointContainer;
    const iterator = imageQuantizer.quantize(image, palette);
    const next = () => {
      try {
        const result = iterator.next();
        if (result.done) {
          resolve(outPointContainer);
        } else {
          if (result.value.pointContainer) {
            outPointContainer = result.value.pointContainer;
          }
          if (onProgress)
            onProgress(result.value.progress);
          setImmediateImpl(next);
        }
      } catch (error) {
        reject(error);
      }
    };
    setImmediateImpl(next);
  });
}
function colorDistanceFormulaToColorDistance(colorDistanceFormula = "euclidean-bt709") {
  switch (colorDistanceFormula) {
    case "cie94-graphic-arts":
      return new CIE94GraphicArts();
    case "cie94-textiles":
      return new CIE94Textiles();
    case "ciede2000":
      return new CIEDE2000();
    case "color-metric":
      return new CMetric();
    case "euclidean":
      return new Euclidean();
    case "euclidean-bt709":
      return new EuclideanBT709();
    case "euclidean-bt709-noalpha":
      return new EuclideanBT709NoAlpha();
    case "manhattan":
      return new Manhattan();
    case "manhattan-bt709":
      return new ManhattanBT709();
    case "manhattan-nommyde":
      return new ManhattanNommyde();
    case "pngquant":
      return new PNGQuant();
    default:
      throw new Error(`Unknown colorDistanceFormula ${colorDistanceFormula}`);
  }
}
function imageQuantizationToImageQuantizer(distanceCalculator, imageQuantization = "floyd-steinberg") {
  switch (imageQuantization) {
    case "nearest":
      return new NearestColor(distanceCalculator);
    case "riemersma":
      return new ErrorDiffusionRiemersma(distanceCalculator);
    case "floyd-steinberg":
      return new ErrorDiffusionArray(distanceCalculator, 0 /* FloydSteinberg */);
    case "false-floyd-steinberg":
      return new ErrorDiffusionArray(distanceCalculator, 1 /* FalseFloydSteinberg */);
    case "stucki":
      return new ErrorDiffusionArray(distanceCalculator, 2 /* Stucki */);
    case "atkinson":
      return new ErrorDiffusionArray(distanceCalculator, 3 /* Atkinson */);
    case "jarvis":
      return new ErrorDiffusionArray(distanceCalculator, 4 /* Jarvis */);
    case "burkes":
      return new ErrorDiffusionArray(distanceCalculator, 5 /* Burkes */);
    case "sierra":
      return new ErrorDiffusionArray(distanceCalculator, 6 /* Sierra */);
    case "two-sierra":
      return new ErrorDiffusionArray(distanceCalculator, 7 /* TwoSierra */);
    case "sierra-lite":
      return new ErrorDiffusionArray(distanceCalculator, 8 /* SierraLite */);
    default:
      throw new Error(`Unknown imageQuantization ${imageQuantization}`);
  }
}
function paletteQuantizationToPaletteQuantizer(distanceCalculator, paletteQuantization = "wuquant", colors = 256) {
  switch (paletteQuantization) {
    case "neuquant":
      return new NeuQuant(distanceCalculator, colors);
    case "rgbquant":
      return new RGBQuant(distanceCalculator, colors);
    case "wuquant":
      return new WuQuant(distanceCalculator, colors);
    case "neuquant-float":
      return new NeuQuantFloat(distanceCalculator, colors);
    default:
      throw new Error(`Unknown paletteQuantization ${paletteQuantization}`);
  }
}
module.exports = __toCommonJS(src_exports);
// Annotate the CommonJS export names for ESM import in node:
0 && (0);
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * cie94.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * ciede2000.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * cmetric.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * common.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * constants.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * ditherErrorDiffusionArray.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * euclidean.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * helper.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * hueStatistics.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * iq.ts - Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * lab2rgb.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * lab2xyz.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * manhattanNeuQuant.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * nearestColor.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * palette.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * pngQuant.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * point.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * pointContainer.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * rgb2hsl.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * rgb2lab.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * rgb2xyz.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * ssim.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * wuQuant.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * xyz2lab.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * xyz2rgb.ts - part of Image Quantization Library
 */
/**
 * @preserve
 * MIT License
 *
 * Copyright 2015-2018 Igor Bezkrovnyi
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to
 * deal in the Software without restriction, including without limitation the
 * rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
 * sell copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
 * THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
 * IN THE SOFTWARE.
 *
 * riemersma.ts - part of Image Quantization Library
 */
/**
 * @preserve TypeScript port:
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * colorHistogram.ts - part of Image Quantization Library
 */
/**
 * @preserve TypeScript port:
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * neuquant.ts - part of Image Quantization Library
 */
/**
 * @preserve TypeScript port:
 * Copyright 2015-2018 Igor Bezkrovnyi
 * All rights reserved. (MIT Licensed)
 *
 * rgbquant.ts - part of Image Quantization Library
 */
//# sourceMappingURL=image-q.cjs.map


/***/ },

/***/ "../../node_modules/.pnpm/@jimp+js-gif@1.6.0/node_modules/@jimp/js-gif/dist/esm/index.js"
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (/* binding */ gif)
/* harmony export */ });
/* harmony import */ var omggif__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__("../../node_modules/.pnpm/omggif@1.0.10/node_modules/omggif/omggif.js");
/* harmony import */ var gifwrap__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__("../../node_modules/.pnpm/gifwrap@0.10.1/node_modules/gifwrap/src/index.js");


function gif() {
    return {
        mime: "image/gif",
        encode: async (bitmap) => {
            const gif = new gifwrap__WEBPACK_IMPORTED_MODULE_1__.BitmapImage(bitmap);
            gifwrap__WEBPACK_IMPORTED_MODULE_1__.GifUtil.quantizeDekker(gif, 256);
            const newFrame = new gifwrap__WEBPACK_IMPORTED_MODULE_1__.GifFrame(bitmap);
            const gifCodec = new gifwrap__WEBPACK_IMPORTED_MODULE_1__.GifCodec();
            const newGif = await gifCodec.encodeGif([newFrame], {});
            return newGif.buffer;
        },
        decode: (data) => {
            const gifObj = new omggif__WEBPACK_IMPORTED_MODULE_0__.GifReader(data);
            const gifData = Buffer.alloc(gifObj.width * gifObj.height * 4);
            gifObj.decodeAndBlitFrameRGBA(0, gifData);
            return {
                data: gifData,
                width: gifObj.width,
                height: gifObj.height,
            };
        },
    };
}
//# sourceMappingURL=index.js.map

/***/ }

};
;