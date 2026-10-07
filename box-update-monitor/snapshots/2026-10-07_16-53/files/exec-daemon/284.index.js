"use strict";
exports.id = 284;
exports.ids = [284];
exports.modules = {

/***/ "../../node_modules/.pnpm/@jimp+js-bmp@1.6.0/node_modules/@jimp/js-bmp/dist/esm/index.js"
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

// ESM COMPAT FLAG
__webpack_require__.r(__webpack_exports__);

// EXPORTS
__webpack_require__.d(__webpack_exports__, {
  BmpCompression: () => (/* reexport */ BmpCompression),
  "default": () => (/* binding */ bmp),
  msBmp: () => (/* binding */ msBmp)
});

;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/header-types.js
var HeaderTypes;
(function (HeaderTypes) {
    HeaderTypes[HeaderTypes["BITMAP_INFO_HEADER"] = 40] = "BITMAP_INFO_HEADER";
    HeaderTypes[HeaderTypes["BITMAP_V2_INFO_HEADER"] = 52] = "BITMAP_V2_INFO_HEADER";
    HeaderTypes[HeaderTypes["BITMAP_V3_INFO_HEADER"] = 56] = "BITMAP_V3_INFO_HEADER";
    HeaderTypes[HeaderTypes["BITMAP_V4_HEADER"] = 108] = "BITMAP_V4_HEADER";
    HeaderTypes[HeaderTypes["BITMAP_V5_HEADER"] = 124] = "BITMAP_V5_HEADER";
})(HeaderTypes || (HeaderTypes = {}));
/* harmony default export */ const header_types = (HeaderTypes);
//# sourceMappingURL=header-types.js.map
;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/mask-color.js
// We have these:
//
// const sample = 0101 0101 0101 0101
// const mask   = 0111 1100 0000 0000
// 256        === 0000 0001 0000 0000
//
// We want to take the sample and turn it into an 8-bit value.
//
// 1. We extract the last bit of the mask:
//
// 0000 0100 0000 0000
//       ^
//
// Like so:
//
// const a = ~mask =    1000 0011 1111 1111
// const b = a + 1 =    1000 0100 0000 0000
// const c = b & mask = 0000 0100 0000 0000
//
// 2. We shift it to the right and extract the bit before the first:
//
// 0000 0000 0010 0000
//             ^
//
// Like so:
//
// const d = mask / c = 0000 0000 0001 1111
// const e = mask + 1 = 0000 0000 0010 0000
//
// 3. We apply the mask and the two values above to a sample:
//
// const f = sample & mask = 0101 0100 0000 0000
// const g = f / c =         0000 0000 0001 0101
// const h = 256 / e =       0000 0000 0000 0100
// const i = g * h =         0000 0000 1010 1000
//                                     ^^^^ ^
//
// Voila, we have extracted a sample and "stretched" it to 8 bits. For samples
// which are already 8-bit, h === 1 and g === i.
function maskColor(maskRed, maskGreen, maskBlue, maskAlpha) {
    const maskRedR = (~maskRed + 1) & maskRed;
    const maskGreenR = (~maskGreen + 1) & maskGreen;
    const maskBlueR = (~maskBlue + 1) & maskBlue;
    const maskAlphaR = (~maskAlpha + 1) & maskAlpha;
    const shiftedMaskRedL = maskRed / maskRedR + 1;
    const shiftedMaskGreenL = maskGreen / maskGreenR + 1;
    const shiftedMaskBlueL = maskBlue / maskBlueR + 1;
    const shiftedMaskAlphaL = maskAlpha / maskAlphaR + 1;
    return {
        shiftRed: (x) => (((x & maskRed) / maskRedR) * 0x100) / shiftedMaskRedL,
        shiftGreen: (x) => (((x & maskGreen) / maskGreenR) * 0x100) / shiftedMaskGreenL,
        shiftBlue: (x) => (((x & maskBlue) / maskBlueR) * 0x100) / shiftedMaskBlueL,
        shiftAlpha: maskAlpha !== 0
            ? (x) => (((x & maskAlpha) / maskAlphaR) * 0x100) / shiftedMaskAlphaL
            : () => 255,
    };
}
//# sourceMappingURL=mask-color.js.map
;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/types.js
var BmpCompression;
(function (BmpCompression) {
    BmpCompression[BmpCompression["NONE"] = 0] = "NONE";
    BmpCompression[BmpCompression["BI_RLE8"] = 1] = "BI_RLE8";
    BmpCompression[BmpCompression["BI_RLE4"] = 2] = "BI_RLE4";
    BmpCompression[BmpCompression["BI_BIT_FIELDS"] = 3] = "BI_BIT_FIELDS";
    BmpCompression[BmpCompression["BI_ALPHA_BIT_FIELDS"] = 6] = "BI_ALPHA_BIT_FIELDS";
})(BmpCompression || (BmpCompression = {}));
//# sourceMappingURL=types.js.map
;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/decoder.js



class BmpDecoder {
    // Header
    flag;
    fileSize;
    reserved1;
    reserved2;
    offset;
    headerSize;
    width;
    height;
    planes;
    bitPP;
    compression;
    rawSize;
    hr;
    vr;
    colors;
    importantColors;
    palette;
    data;
    maskRed;
    maskGreen;
    maskBlue;
    maskAlpha;
    toRGBA;
    pos;
    bottomUp;
    buffer;
    locRed;
    locGreen;
    locBlue;
    locAlpha;
    shiftRed;
    shiftGreen;
    shiftBlue;
    shiftAlpha;
    constructor(buffer, { toRGBA } = { toRGBA: false }) {
        this.buffer = buffer;
        this.toRGBA = !!toRGBA;
        this.pos = 0;
        this.bottomUp = true;
        this.flag = this.buffer.toString('utf-8', 0, (this.pos += 2));
        if (this.flag !== 'BM') {
            throw new Error('Invalid BMP File');
        }
        this.locRed = this.toRGBA ? 0 : 3;
        this.locGreen = this.toRGBA ? 1 : 2;
        this.locBlue = this.toRGBA ? 2 : 1;
        this.locAlpha = this.toRGBA ? 3 : 0;
        this.parseHeader();
        this.parseRGBA();
    }
    parseHeader() {
        this.fileSize = this.readUInt32LE();
        this.reserved1 = this.buffer.readUInt16LE(this.pos);
        this.pos += 2;
        this.reserved2 = this.buffer.readUInt16LE(this.pos);
        this.pos += 2;
        this.offset = this.readUInt32LE();
        // End of BITMAP_FILE_HEADER
        this.headerSize = this.readUInt32LE();
        if (!(this.headerSize in header_types)) {
            throw new Error(`Unsupported BMP header size ${this.headerSize}`);
        }
        this.width = this.readUInt32LE();
        this.height = this.readUInt32LE();
        // negative value are possible here => implies bottom down
        this.height =
            this.height > 0x7fffffff ? this.height - 0x100000000 : this.height;
        this.planes = this.buffer.readUInt16LE(this.pos);
        this.pos += 2;
        this.bitPP = this.buffer.readUInt16LE(this.pos);
        this.pos += 2;
        this.compression = this.readUInt32LE();
        this.rawSize = this.readUInt32LE();
        this.hr = this.readUInt32LE();
        this.vr = this.readUInt32LE();
        this.colors = this.readUInt32LE();
        this.importantColors = this.readUInt32LE();
        // De facto defaults
        if (this.bitPP === 32) {
            this.maskAlpha = 0;
            this.maskRed = 0x00ff0000;
            this.maskGreen = 0x0000ff00;
            this.maskBlue = 0x000000ff;
        }
        else if (this.bitPP === 16) {
            this.maskAlpha = 0;
            this.maskRed = 0x7c00;
            this.maskGreen = 0x03e0;
            this.maskBlue = 0x001f;
        }
        // End of BITMAP_INFO_HEADER
        if (this.headerSize > header_types.BITMAP_INFO_HEADER ||
            this.compression === BmpCompression.BI_BIT_FIELDS ||
            this.compression === BmpCompression.BI_ALPHA_BIT_FIELDS) {
            this.maskRed = this.readUInt32LE();
            this.maskGreen = this.readUInt32LE();
            this.maskBlue = this.readUInt32LE();
        }
        // End of BITMAP_V2_INFO_HEADER
        if (this.headerSize > header_types.BITMAP_V2_INFO_HEADER ||
            this.compression === BmpCompression.BI_ALPHA_BIT_FIELDS) {
            this.maskAlpha = this.readUInt32LE();
        }
        // End of BITMAP_V3_INFO_HEADER
        if (this.headerSize > header_types.BITMAP_V3_INFO_HEADER) {
            this.pos +=
                header_types.BITMAP_V4_HEADER - header_types.BITMAP_V3_INFO_HEADER;
        }
        // End of BITMAP_V4_HEADER
        if (this.headerSize > header_types.BITMAP_V4_HEADER) {
            this.pos += header_types.BITMAP_V5_HEADER - header_types.BITMAP_V4_HEADER;
        }
        // End of BITMAP_V5_HEADER
        if (this.bitPP <= 8 || this.colors > 0) {
            const len = this.colors === 0 ? 1 << this.bitPP : this.colors;
            this.palette = new Array(len);
            for (let i = 0; i < len; i++) {
                const blue = this.buffer.readUInt8(this.pos++);
                const green = this.buffer.readUInt8(this.pos++);
                const red = this.buffer.readUInt8(this.pos++);
                const quad = this.buffer.readUInt8(this.pos++);
                this.palette[i] = {
                    red,
                    green,
                    blue,
                    quad,
                };
            }
        }
        // End of color table
        // Can the height ever be negative?
        if (this.height < 0) {
            this.height *= -1;
            this.bottomUp = false;
        }
        const coloShift = maskColor(this.maskRed, this.maskGreen, this.maskBlue, this.maskAlpha);
        this.shiftRed = coloShift.shiftRed;
        this.shiftGreen = coloShift.shiftGreen;
        this.shiftBlue = coloShift.shiftBlue;
        this.shiftAlpha = coloShift.shiftAlpha;
    }
    parseRGBA() {
        this.data = Buffer.alloc(this.width * this.height * 4);
        switch (this.bitPP) {
            case 1:
                this.bit1();
                break;
            case 4:
                this.bit4();
                break;
            case 8:
                this.bit8();
                break;
            case 16:
                this.bit16();
                break;
            case 24:
                this.bit24();
                break;
            default:
                this.bit32();
        }
    }
    bit1() {
        const xLen = Math.ceil(this.width / 8);
        const mode = xLen % 4;
        const padding = mode !== 0 ? 4 - mode : 0;
        let lastLine;
        this.scanImage(padding, xLen, (x, line) => {
            if (line !== lastLine) {
                lastLine = line;
            }
            const b = this.buffer.readUInt8(this.pos++);
            const location = line * this.width * 4 + x * 8 * 4;
            for (let i = 0; i < 8; i++) {
                if (x * 8 + i < this.width) {
                    const rgb = this.palette[(b >> (7 - i)) & 0x1];
                    this.data[location + i * this.locAlpha] = 0;
                    this.data[location + i * 4 + this.locBlue] = rgb.blue;
                    this.data[location + i * 4 + this.locGreen] = rgb.green;
                    this.data[location + i * 4 + this.locRed] = rgb.red;
                }
                else {
                    break;
                }
            }
        });
    }
    bit4() {
        if (this.compression === BmpCompression.BI_RLE4) {
            this.data.fill(0);
            let lowNibble = false; //for all count of pixel
            let lines = this.bottomUp ? this.height - 1 : 0;
            let location = 0;
            while (location < this.data.length) {
                const a = this.buffer.readUInt8(this.pos++);
                const b = this.buffer.readUInt8(this.pos++);
                //absolute mode
                if (a === 0) {
                    if (b === 0) {
                        //line end
                        lines += this.bottomUp ? -1 : 1;
                        location = lines * this.width * 4;
                        lowNibble = false;
                        continue;
                    }
                    if (b === 1) {
                        // image end
                        break;
                    }
                    if (b === 2) {
                        // offset x, y
                        const x = this.buffer.readUInt8(this.pos++);
                        const y = this.buffer.readUInt8(this.pos++);
                        lines += this.bottomUp ? -y : y;
                        location += y * this.width * 4 + x * 4;
                    }
                    else {
                        let c = this.buffer.readUInt8(this.pos++);
                        for (let i = 0; i < b; i++) {
                            location = this.setPixelData(location, lowNibble ? c & 0x0f : (c & 0xf0) >> 4);
                            if (i & 1 && i + 1 < b) {
                                c = this.buffer.readUInt8(this.pos++);
                            }
                            lowNibble = !lowNibble;
                        }
                        if ((((b + 1) >> 1) & 1) === 1) {
                            this.pos++;
                        }
                    }
                }
                else {
                    //encoded mode
                    for (let i = 0; i < a; i++) {
                        location = this.setPixelData(location, lowNibble ? b & 0x0f : (b & 0xf0) >> 4);
                        lowNibble = !lowNibble;
                    }
                }
            }
        }
        else {
            const xLen = Math.ceil(this.width / 2);
            const mode = xLen % 4;
            const padding = mode !== 0 ? 4 - mode : 0;
            this.scanImage(padding, xLen, (x, line) => {
                const b = this.buffer.readUInt8(this.pos++);
                const location = line * this.width * 4 + x * 2 * 4;
                const first4 = b >> 4;
                let rgb = this.palette[first4];
                this.data[location] = 0;
                this.data[location + 1] = rgb.blue;
                this.data[location + 2] = rgb.green;
                this.data[location + 3] = rgb.red;
                if (x * 2 + 1 >= this.width) {
                    // throw new Error('Something');
                    return false;
                }
                const last4 = b & 0x0f;
                rgb = this.palette[last4];
                this.data[location + 4] = 0;
                this.data[location + 4 + 1] = rgb.blue;
                this.data[location + 4 + 2] = rgb.green;
                this.data[location + 4 + 3] = rgb.red;
            });
        }
    }
    bit8() {
        if (this.compression === BmpCompression.BI_RLE8) {
            this.data.fill(0);
            let lines = this.bottomUp ? this.height - 1 : 0;
            let location = 0;
            while (location < this.data.length) {
                const a = this.buffer.readUInt8(this.pos++);
                const b = this.buffer.readUInt8(this.pos++);
                //absolute mode
                if (a === 0) {
                    if (b === 0) {
                        //line end
                        lines += this.bottomUp ? -1 : 1;
                        location = lines * this.width * 4;
                        continue;
                    }
                    if (b === 1) {
                        //image end
                        break;
                    }
                    if (b === 2) {
                        //offset x,y
                        const x = this.buffer.readUInt8(this.pos++);
                        const y = this.buffer.readUInt8(this.pos++);
                        lines += this.bottomUp ? -y : y;
                        location += y * this.width * 4 + x * 4;
                    }
                    else {
                        for (let i = 0; i < b; i++) {
                            const c = this.buffer.readUInt8(this.pos++);
                            location = this.setPixelData(location, c);
                        }
                        // @ts-ignore
                        const shouldIncrement = b & (1 === 1);
                        if (shouldIncrement) {
                            this.pos++;
                        }
                    }
                }
                else {
                    //encoded mode
                    for (let i = 0; i < a; i++) {
                        location = this.setPixelData(location, b);
                    }
                }
            }
        }
        else {
            const mode = this.width % 4;
            const padding = mode !== 0 ? 4 - mode : 0;
            this.scanImage(padding, this.width, (x, line) => {
                const b = this.buffer.readUInt8(this.pos++);
                const location = line * this.width * 4 + x * 4;
                if (b < this.palette.length) {
                    const rgb = this.palette[b];
                    this.data[location] = 0;
                    this.data[location + 1] = rgb.blue;
                    this.data[location + 2] = rgb.green;
                    this.data[location + 3] = rgb.red;
                }
                else {
                    this.data[location] = 0;
                    this.data[location + 1] = 0xff;
                    this.data[location + 2] = 0xff;
                    this.data[location + 3] = 0xff;
                }
            });
        }
    }
    bit16() {
        const padding = (this.width % 2) * 2;
        this.scanImage(padding, this.width, (x, line) => {
            const loc = line * this.width * 4 + x * 4;
            const px = this.buffer.readUInt16LE(this.pos);
            this.pos += 2;
            this.data[loc + this.locRed] = this.shiftRed(px);
            this.data[loc + this.locGreen] = this.shiftGreen(px);
            this.data[loc + this.locBlue] = this.shiftBlue(px);
            this.data[loc + this.locAlpha] = this.shiftAlpha(px);
        });
    }
    bit24() {
        const padding = this.width % 4;
        this.scanImage(padding, this.width, (x, line) => {
            const loc = line * this.width * 4 + x * 4;
            const blue = this.buffer.readUInt8(this.pos++);
            const green = this.buffer.readUInt8(this.pos++);
            const red = this.buffer.readUInt8(this.pos++);
            this.data[loc + this.locRed] = red;
            this.data[loc + this.locGreen] = green;
            this.data[loc + this.locBlue] = blue;
            this.data[loc + this.locAlpha] = 0;
        });
    }
    bit32() {
        this.scanImage(0, this.width, (x, line) => {
            const loc = line * this.width * 4 + x * 4;
            const px = this.readUInt32LE();
            this.data[loc + this.locRed] = this.shiftRed(px);
            this.data[loc + this.locGreen] = this.shiftGreen(px);
            this.data[loc + this.locBlue] = this.shiftBlue(px);
            this.data[loc + this.locAlpha] = this.shiftAlpha(px);
        });
    }
    scanImage(padding = 0, width = this.width, processPixel) {
        for (let y = this.height - 1; y >= 0; y--) {
            const line = this.bottomUp ? y : this.height - 1 - y;
            for (let x = 0; x < width; x++) {
                const result = processPixel.call(this, x, line);
                if (result === false) {
                    return;
                }
            }
            this.pos += padding;
        }
    }
    readUInt32LE() {
        const value = this.buffer.readUInt32LE(this.pos);
        this.pos += 4;
        return value;
    }
    setPixelData(location, rgbIndex) {
        const { blue, green, red } = this.palette[rgbIndex];
        this.data[location + this.locAlpha] = 0;
        this.data[location + 1 + this.locBlue] = blue;
        this.data[location + 2 + this.locGreen] = green;
        this.data[location + 3 + this.locRed] = red;
        return location + 4;
    }
}
//# sourceMappingURL=decoder.js.map
;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/encoder.js

function createInteger(numbers) {
    return numbers.reduce((final, n) => (final << 1) | n, 0);
}
function createColor(color) {
    return ((color.quad << 24) | (color.red << 16) | (color.green << 8) | color.blue);
}
class BmpEncoder {
    fileSize;
    reserved1;
    reserved2;
    offset;
    width;
    flag;
    height;
    planes;
    bitPP;
    compress;
    hr;
    vr;
    colors;
    importantColors;
    rawSize;
    headerSize;
    data;
    palette;
    extraBytes;
    buffer;
    bytesInColor;
    pos;
    constructor(imgData) {
        this.buffer = imgData.data;
        this.width = imgData.width;
        this.height = imgData.height;
        this.headerSize = header_types.BITMAP_INFO_HEADER;
        // Header
        this.flag = 'BM';
        this.bitPP = imgData.bitPP || 24;
        this.offset = 54;
        this.reserved1 = imgData.reserved1 || 0;
        this.reserved2 = imgData.reserved2 || 0;
        this.planes = 1;
        this.compress = 0;
        this.hr = imgData.hr || 0;
        this.vr = imgData.vr || 0;
        this.importantColors = imgData.importantColors || 0;
        this.colors = Math.min(2 ** (this.bitPP - 1 || 1), imgData.colors || Infinity);
        this.palette = imgData.palette || [];
        if (this.colors && this.bitPP < 16) {
            this.offset += this.colors * 4;
        }
        else {
            this.colors = 0;
        }
        switch (this.bitPP) {
            case 32:
                this.bytesInColor = 4;
                break;
            case 16:
                this.bytesInColor = 2;
                break;
            case 8:
                this.bytesInColor = 1;
                break;
            case 4:
                this.bytesInColor = 1 / 2;
                break;
            case 1:
                this.bytesInColor = 1 / 8;
                break;
            default:
                this.bytesInColor = 3;
                this.bitPP = 24;
        }
        const rowWidth = (this.width * this.bitPP) / 32;
        const rowBytes = Math.ceil(rowWidth);
        this.extraBytes = (rowBytes - rowWidth) * 4;
        // Why 2?
        this.rawSize = this.height * rowBytes * 4 + 2;
        this.fileSize = this.rawSize + this.offset;
        this.data = Buffer.alloc(this.fileSize, 0x1);
        this.pos = 0;
        this.encode();
    }
    encode() {
        this.pos = 0;
        this.writeHeader();
        switch (this.bitPP) {
            case 32:
                this.bit32();
                break;
            case 16:
                this.bit16();
                break;
            case 8:
                this.bit8();
                break;
            case 4:
                this.bit4();
                break;
            case 1:
                this.bit1();
                break;
            default:
                this.bit24();
        }
    }
    writeHeader() {
        this.data.write(this.flag, this.pos, 2);
        this.pos += 2;
        this.writeUInt32LE(this.fileSize);
        // Writing 2 UInt16LE resulted in a weird bug
        this.writeUInt32LE((this.reserved1 << 16) | this.reserved2);
        this.writeUInt32LE(this.offset);
        this.writeUInt32LE(this.headerSize);
        this.writeUInt32LE(this.width);
        this.writeUInt32LE(this.height);
        this.data.writeUInt16LE(this.planes, this.pos);
        this.pos += 2;
        this.data.writeUInt16LE(this.bitPP, this.pos);
        this.pos += 2;
        this.writeUInt32LE(this.compress);
        this.writeUInt32LE(this.rawSize);
        this.writeUInt32LE(this.hr);
        this.writeUInt32LE(this.vr);
        this.writeUInt32LE(this.colors);
        this.writeUInt32LE(this.importantColors);
    }
    bit1() {
        if (this.palette.length && this.colors === 2) {
            this.initColors(1);
        }
        else {
            this.writeUInt32LE(0x00ffffff); // Black
            this.writeUInt32LE(0x00000000); // White
        }
        this.pos += 1; // ?
        let lineArr = [];
        this.writeImage((p, index, x) => {
            let i = index;
            i++;
            const b = this.buffer[i++];
            const g = this.buffer[i++];
            const r = this.buffer[i++];
            const brightness = r * 0.2126 + g * 0.7152 + b * 0.0722;
            lineArr.push(brightness > 127 ? 0 : 1);
            if ((x + 1) % 8 === 0) {
                this.data[p - 1] = createInteger(lineArr);
                lineArr = [];
            }
            else if (x === this.width - 1 && lineArr.length > 0) {
                this.data[p - 1] = createInteger(lineArr) << 4;
                lineArr = [];
            }
            return i;
        });
    }
    bit4() {
        const colors = this.initColors(4);
        let integerPair = [];
        this.writeImage((p, index, x) => {
            let i = index;
            const colorInt = createColor({
                quad: this.buffer[i++],
                blue: this.buffer[i++],
                green: this.buffer[i++],
                red: this.buffer[i++],
            });
            const colorExists = colors.findIndex((c) => c === colorInt);
            if (colorExists !== -1) {
                integerPair.push(colorExists);
            }
            else {
                integerPair.push(0);
            }
            if ((x + 1) % 2 === 0) {
                this.data[p] = (integerPair[0] << 4) | integerPair[1];
                integerPair = [];
            }
            return i;
        });
    }
    bit8() {
        const colors = this.initColors(8);
        this.writeImage((p, index) => {
            let i = index;
            const colorInt = createColor({
                quad: this.buffer[i++],
                blue: this.buffer[i++],
                green: this.buffer[i++],
                red: this.buffer[i++],
            });
            const colorExists = colors.findIndex((c) => c === colorInt);
            if (colorExists !== -1) {
                this.data[p] = colorExists;
            }
            else {
                this.data[p] = 0;
            }
            return i;
        });
    }
    bit16() {
        this.writeImage((p, index) => {
            let i = index + 1;
            const b = this.buffer[i++] / 8; // b
            const g = this.buffer[i++] / 8; // g
            const r = this.buffer[i++] / 8; // r
            const color = (r << 10) | (g << 5) | b;
            this.data[p] = color & 0x00ff;
            this.data[p + 1] = (color & 0xff00) >> 8;
            return i;
        });
    }
    bit24() {
        this.writeImage((p, index) => {
            let i = index + 1;
            this.data[p] = this.buffer[i++]; //b
            this.data[p + 1] = this.buffer[i++]; //g
            this.data[p + 2] = this.buffer[i++]; //r
            return i;
        });
    }
    bit32() {
        this.writeImage((p, index) => {
            let i = index;
            this.data[p + 3] = this.buffer[i++]; // a
            this.data[p] = this.buffer[i++]; // b
            this.data[p + 1] = this.buffer[i++]; // g
            this.data[p + 2] = this.buffer[i++]; // r
            return i;
        });
    }
    writeImage(writePixel) {
        const rowBytes = this.extraBytes + this.width * this.bytesInColor;
        let i = 0;
        for (let y = 0; y < this.height; y++) {
            for (let x = 0; x < this.width; x++) {
                const p = Math.floor(this.pos + (this.height - 1 - y) * rowBytes + x * this.bytesInColor);
                i = writePixel.call(this, p, i, x, y);
            }
        }
    }
    initColors(bit) {
        const colors = [];
        if (this.palette.length) {
            for (let i = 0; i < this.colors; i++) {
                const rootColor = createColor(this.palette[i]);
                this.writeUInt32LE(rootColor);
                colors.push(rootColor);
            }
        }
        else {
            throw new Error(`To encode ${bit}-bit BMPs a pallette is needed. Please choose up to ${this.colors} colors. Colors must be 32-bit integers.`);
        }
        return colors;
    }
    writeUInt32LE(value) {
        this.data.writeUInt32LE(value, this.pos);
        this.pos += 4;
    }
}
//# sourceMappingURL=encoder.js.map
;// ../../node_modules/.pnpm/bmp-ts@1.0.9/node_modules/bmp-ts/dist/esm/index.js


function decode(bmpData, options) {
    return new BmpDecoder(bmpData, options);
}
function encode(imgData) {
    return new BmpEncoder(imgData);
}

//# sourceMappingURL=index.js.map
// EXTERNAL MODULE: ../../node_modules/.pnpm/@jimp+utils@1.6.0/node_modules/@jimp/utils/dist/esm/index.js + 1 modules
var esm = __webpack_require__("../../node_modules/.pnpm/@jimp+utils@1.6.0/node_modules/@jimp/utils/dist/esm/index.js");
;// ../../node_modules/.pnpm/@jimp+js-bmp@1.6.0/node_modules/@jimp/js-bmp/dist/esm/index.js



function esm_encode(image, options = {}) {
    (0,esm/* scan */.SQ)({ bitmap: image }, 0, 0, image.width, image.height, function (_, __, index) {
        const red = image.data[index + 0];
        const green = image.data[index + 1];
        const blue = image.data[index + 2];
        const alpha = image.data[index + 3];
        image.data[index + 0] = alpha;
        image.data[index + 1] = blue;
        image.data[index + 2] = green;
        image.data[index + 3] = red;
    });
    return encode({ ...image, ...options }).data;
}
function esm_decode(data, options) {
    const result = decode(data, options);
    (0,esm/* scan */.SQ)({ bitmap: result }, 0, 0, result.width, result.height, function (_, __, index) {
        // const alpha = result.data[index + 0]!;
        const blue = result.data[index + 1];
        const green = result.data[index + 2];
        const red = result.data[index + 3];
        result.data[index + 0] = red;
        result.data[index + 1] = green;
        result.data[index + 2] = blue;
        result.data[index + 3] = 0xff;
    });
    return result;
}
function msBmp() {
    return {
        mime: "image/x-ms-bmp",
        encode: esm_encode,
        decode: esm_decode,
    };
}
function bmp() {
    return {
        mime: "image/bmp",
        encode: esm_encode,
        decode: esm_decode,
    };
}
//# sourceMappingURL=index.js.map

/***/ }

};
;