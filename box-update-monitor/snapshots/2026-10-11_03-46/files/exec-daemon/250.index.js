exports.id = 250;
exports.ids = [250];
exports.modules = {

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/index.js"
(module, __unused_webpack_exports, __webpack_require__) {

var Parser = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/parser.js");

function getGlobal() {
	return (1,eval)('this');
}

module.exports = {
	create: function(buffer, global) {
		global = global || getGlobal();
		if(buffer instanceof global.ArrayBuffer) {
			var DOMBufferStream = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/dom-bufferstream.js");
			return new Parser(new DOMBufferStream(buffer, 0, buffer.byteLength, true, global));
		} else {
			var NodeBufferStream = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/bufferstream.js");
			return new Parser(new NodeBufferStream(buffer, 0, buffer.length, true));
		}
	}
};


/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/bufferstream.js"
(module) {

function BufferStream(buffer, offset, length, bigEndian) {
	this.buffer = buffer;
	this.offset = offset || 0;
	length = typeof length === 'number' ? length : buffer.length;
	this.endPosition = this.offset + length;
	this.setBigEndian(bigEndian);
}

BufferStream.prototype = {
	setBigEndian: function(bigEndian) {
		this.bigEndian = !!bigEndian;
	},
	nextUInt8: function() {
		var value = this.buffer.readUInt8(this.offset);
		this.offset += 1;
		return value;
	},
	nextInt8: function() {
		var value = this.buffer.readInt8(this.offset);
		this.offset += 1;
		return value;
	},
	nextUInt16: function() {
		var value = this.bigEndian ? this.buffer.readUInt16BE(this.offset) : this.buffer.readUInt16LE(this.offset);
		this.offset += 2;
		return value;
	},
	nextUInt32: function() {
		var value = this.bigEndian ? this.buffer.readUInt32BE(this.offset) : this.buffer.readUInt32LE(this.offset);
		this.offset += 4;
		return value;
	},
	nextInt16: function() {
		var value = this.bigEndian ? this.buffer.readInt16BE(this.offset) : this.buffer.readInt16LE(this.offset);
		this.offset += 2;
		return value;
	},
	nextInt32: function() {
		var value = this.bigEndian ? this.buffer.readInt32BE(this.offset) : this.buffer.readInt32LE(this.offset);
		this.offset += 4;
		return value;
	},
	nextFloat: function() {
		var value = this.bigEndian ? this.buffer.readFloatBE(this.offset) : this.buffer.readFloatLE(this.offset);
		this.offset += 4;
		return value;
	},
	nextDouble: function() {
		var value = this.bigEndian ? this.buffer.readDoubleBE(this.offset) : this.buffer.readDoubleLE(this.offset);
		this.offset += 8;
		return value;
	},
	nextBuffer: function(length) {
		var value = this.buffer.slice(this.offset, this.offset + length);
		this.offset += length;
		return value;
	},
	remainingLength: function() {
		return this.endPosition - this.offset;
	},
	nextString: function(length) {
		var value = this.buffer.toString('utf8', this.offset, this.offset + length);
		this.offset += length;
		return value;
	},
	mark: function() {
		var self = this;
		return {
			openWithOffset: function(offset) {
				offset = (offset || 0) + this.offset;
				return new BufferStream(self.buffer, offset, self.endPosition - offset, self.bigEndian);
			},
			offset: this.offset
		};
	},
	offsetFrom: function(marker) {
		return this.offset - marker.offset;
	},
	skip: function(amount) {
		this.offset += amount;
	},
	branch: function(offset, length) {
		length = typeof length === 'number' ? length : this.endPosition - (this.offset + offset);
		return new BufferStream(this.buffer, this.offset + offset, length, this.bigEndian);
	}
};

module.exports = BufferStream;


/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/date.js"
(module) {

function parseNumber(s) {
	return parseInt(s, 10);
}

//in seconds
var hours = 3600;
var minutes = 60;

//take date (year, month, day) and time (hour, minutes, seconds) digits in UTC
//and return a timestamp in seconds
function parseDateTimeParts(dateParts, timeParts) {
	dateParts = dateParts.map(parseNumber);
	timeParts = timeParts.map(parseNumber);
	var year = dateParts[0];
	var month = dateParts[1] - 1;
	var day = dateParts[2];
	var hours = timeParts[0];
	var minutes = timeParts[1];
	var seconds = timeParts[2];
	var date = Date.UTC(year, month, day, hours, minutes, seconds, 0);
	var timestamp = date / 1000;
	return timestamp;
}

//parse date with "2004-09-04T23:39:06-08:00" format,
//one of the formats supported by ISO 8601, and
//convert to utc timestamp in seconds
function parseDateWithTimezoneFormat(dateTimeStr) {

	var dateParts = dateTimeStr.substr(0, 10).split('-');
	var timeParts = dateTimeStr.substr(11, 8).split(':');
	var timezoneStr = dateTimeStr.substr(19, 6);
	var timezoneParts = timezoneStr.split(':').map(parseNumber);
	var timezoneOffset = (timezoneParts[0] * hours) +
		(timezoneParts[1] * minutes);

	var timestamp = parseDateTimeParts(dateParts, timeParts);
	//minus because the timezoneOffset describes
	//how much the described time is ahead of UTC
	timestamp -= timezoneOffset;

	if(typeof timestamp === 'number' && !isNaN(timestamp)) {
		return timestamp;
	}
}

//parse date with "YYYY:MM:DD hh:mm:ss" format, convert to utc timestamp in seconds
function parseDateWithSpecFormat(dateTimeStr) {
	var parts = dateTimeStr.split(' '),
		dateParts = parts[0].split(':'),
		timeParts = parts[1].split(':');

	var timestamp = parseDateTimeParts(dateParts, timeParts);

	if(typeof timestamp === 'number' && !isNaN(timestamp)) {
		return timestamp;
	}
}

function parseExifDate(dateTimeStr) {
	//some easy checks to determine two common date formats

	//is the date in the standard "YYYY:MM:DD hh:mm:ss" format?
	var isSpecFormat = dateTimeStr.length === 19 &&
		dateTimeStr.charAt(4) === ':';
	//is the date in the non-standard format,
	//"2004-09-04T23:39:06-08:00" to include a timezone?
	var isTimezoneFormat = dateTimeStr.length === 25 &&
		dateTimeStr.charAt(10) === 'T';
	var timestamp;

	if(isTimezoneFormat) {
		return parseDateWithTimezoneFormat(dateTimeStr);
	}
	else if(isSpecFormat) {
		return parseDateWithSpecFormat(dateTimeStr);
	}
}

module.exports = {
	parseDateWithSpecFormat: parseDateWithSpecFormat,
	parseDateWithTimezoneFormat: parseDateWithTimezoneFormat,
	parseExifDate: parseExifDate
};


/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/dom-bufferstream.js"
(module) {

/*jslint browser: true, devel: true, bitwise: false, debug: true, eqeq: false, es5: true, evil: false, forin: false, newcap: false, nomen: true, plusplus: true, regexp: false, unparam: false, sloppy: true, stupid: false, sub: false, todo: true, vars: true, white: true */

function DOMBufferStream(arrayBuffer, offset, length, bigEndian, global, parentOffset) {
	this.global = global;
	offset = offset || 0;
	length = length || (arrayBuffer.byteLength - offset);
	this.arrayBuffer = arrayBuffer.slice(offset, offset + length);
	this.view = new global.DataView(this.arrayBuffer, 0, this.arrayBuffer.byteLength);
	this.setBigEndian(bigEndian);
	this.offset = 0;
	this.parentOffset = (parentOffset || 0) + offset;
}

DOMBufferStream.prototype = {
	setBigEndian: function(bigEndian) {
		this.littleEndian = !bigEndian;
	},
	nextUInt8: function() {
		var value = this.view.getUint8(this.offset);
		this.offset += 1;
		return value;
	},
	nextInt8: function() {
		var value = this.view.getInt8(this.offset);
		this.offset += 1;
		return value;
	},
	nextUInt16: function() {
		var value = this.view.getUint16(this.offset, this.littleEndian);
		this.offset += 2;
		return value;
	},
	nextUInt32: function() {
		var value = this.view.getUint32(this.offset, this.littleEndian);
		this.offset += 4;
		return value;
	},
	nextInt16: function() {
		var value = this.view.getInt16(this.offset, this.littleEndian);
		this.offset += 2;
		return value;
	},
	nextInt32: function() {
		var value = this.view.getInt32(this.offset, this.littleEndian);
		this.offset += 4;
		return value;
	},
	nextFloat: function() {
		var value = this.view.getFloat32(this.offset, this.littleEndian);
		this.offset += 4;
		return value;
	},
	nextDouble: function() {
		var value = this.view.getFloat64(this.offset, this.littleEndian);
		this.offset += 8;
		return value;
	},
	nextBuffer: function(length) {
		//this won't work in IE10
		var value = this.arrayBuffer.slice(this.offset, this.offset + length);
		this.offset += length;
		return value;
	},
	remainingLength: function() {
		return this.arrayBuffer.byteLength - this.offset;
	},
	nextString: function(length) {
		var value = this.arrayBuffer.slice(this.offset, this.offset + length);
		value = String.fromCharCode.apply(null, new this.global.Uint8Array(value));
		this.offset += length;
		return value;
	},
	mark: function() {
		var self = this;
		return {
			openWithOffset: function(offset) {
				offset = (offset || 0) + this.offset;
				return new DOMBufferStream(self.arrayBuffer, offset, self.arrayBuffer.byteLength - offset, !self.littleEndian, self.global, self.parentOffset);
			},
			offset: this.offset,
			getParentOffset: function() {
				return self.parentOffset;
			}
		};
	},
	offsetFrom: function(marker) {
		return this.parentOffset + this.offset - (marker.offset + marker.getParentOffset());
	},
	skip: function(amount) {
		this.offset += amount;
	},
	branch: function(offset, length) {
		length = typeof length === 'number' ? length : this.arrayBuffer.byteLength - (this.offset + offset);
		return new DOMBufferStream(this.arrayBuffer, this.offset + offset, length, !this.littleEndian, this.global, this.parentOffset);
	}
};

module.exports = DOMBufferStream;


/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/exif-tags.js"
(module) {

module.exports = {
	exif : {
		0x0001 : "InteropIndex",
		0x0002 : "InteropVersion",
		0x000B : "ProcessingSoftware",
		0x00FE : "SubfileType",
		0x00FF : "OldSubfileType",
		0x0100 : "ImageWidth",
		0x0101 : "ImageHeight",
		0x0102 : "BitsPerSample",
		0x0103 : "Compression",
		0x0106 : "PhotometricInterpretation",
		0x0107 : "Thresholding",
		0x0108 : "CellWidth",
		0x0109 : "CellLength",
		0x010A : "FillOrder",
		0x010D : "DocumentName",
		0x010E : "ImageDescription",
		0x010F : "Make",
		0x0110 : "Model",
		0x0111 : "StripOffsets",
		0x0112 : "Orientation",
		0x0115 : "SamplesPerPixel",
		0x0116 : "RowsPerStrip",
		0x0117 : "StripByteCounts",
		0x0118 : "MinSampleValue",
		0x0119 : "MaxSampleValue",
		0x011A : "XResolution",
		0x011B : "YResolution",
		0x011C : "PlanarConfiguration",
		0x011D : "PageName",
		0x011E : "XPosition",
		0x011F : "YPosition",
		0x0120 : "FreeOffsets",
		0x0121 : "FreeByteCounts",
		0x0122 : "GrayResponseUnit",
		0x0123 : "GrayResponseCurve",
		0x0124 : "T4Options",
		0x0125 : "T6Options",
		0x0128 : "ResolutionUnit",
		0x0129 : "PageNumber",
		0x012C : "ColorResponseUnit",
		0x012D : "TransferFunction",
		0x0131 : "Software",
		0x0132 : "ModifyDate",
		0x013B : "Artist",
		0x013C : "HostComputer",
		0x013D : "Predictor",
		0x013E : "WhitePoint",
		0x013F : "PrimaryChromaticities",
		0x0140 : "ColorMap",
		0x0141 : "HalftoneHints",
		0x0142 : "TileWidth",
		0x0143 : "TileLength",
		0x0144 : "TileOffsets",
		0x0145 : "TileByteCounts",
		0x0146 : "BadFaxLines",
		0x0147 : "CleanFaxData",
		0x0148 : "ConsecutiveBadFaxLines",
		0x014A : "SubIFD",
		0x014C : "InkSet",
		0x014D : "InkNames",
		0x014E : "NumberofInks",
		0x0150 : "DotRange",
		0x0151 : "TargetPrinter",
		0x0152 : "ExtraSamples",
		0x0153 : "SampleFormat",
		0x0154 : "SMinSampleValue",
		0x0155 : "SMaxSampleValue",
		0x0156 : "TransferRange",
		0x0157 : "ClipPath",
		0x0158 : "XClipPathUnits",
		0x0159 : "YClipPathUnits",
		0x015A : "Indexed",
		0x015B : "JPEGTables",
		0x015F : "OPIProxy",
		0x0190 : "GlobalParametersIFD",
		0x0191 : "ProfileType",
		0x0192 : "FaxProfile",
		0x0193 : "CodingMethods",
		0x0194 : "VersionYear",
		0x0195 : "ModeNumber",
		0x01B1 : "Decode",
		0x01B2 : "DefaultImageColor",
		0x01B3 : "T82Options",
		0x01B5 : "JPEGTables",
		0x0200 : "JPEGProc",
		0x0201 : "ThumbnailOffset",
		0x0202 : "ThumbnailLength",
		0x0203 : "JPEGRestartInterval",
		0x0205 : "JPEGLosslessPredictors",
		0x0206 : "JPEGPointTransforms",
		0x0207 : "JPEGQTables",
		0x0208 : "JPEGDCTables",
		0x0209 : "JPEGACTables",
		0x0211 : "YCbCrCoefficients",
		0x0212 : "YCbCrSubSampling",
		0x0213 : "YCbCrPositioning",
		0x0214 : "ReferenceBlackWhite",
		0x022F : "StripRowCounts",
		0x02BC : "ApplicationNotes",
		0x03E7 : "USPTOMiscellaneous",
		0x1000 : "RelatedImageFileFormat",
		0x1001 : "RelatedImageWidth",
		0x1002 : "RelatedImageHeight",
		0x4746 : "Rating",
		0x4747 : "XP_DIP_XML",
		0x4748 : "StitchInfo",
		0x4749 : "RatingPercent",
		0x800D : "ImageID",
		0x80A3 : "WangTag1",
		0x80A4 : "WangAnnotation",
		0x80A5 : "WangTag3",
		0x80A6 : "WangTag4",
		0x80E3 : "Matteing",
		0x80E4 : "DataType",
		0x80E5 : "ImageDepth",
		0x80E6 : "TileDepth",
		0x827D : "Model2",
		0x828D : "CFARepeatPatternDim",
		0x828E : "CFAPattern2",
		0x828F : "BatteryLevel",
		0x8290 : "KodakIFD",
		0x8298 : "Copyright",
		0x829A : "ExposureTime",
		0x829D : "FNumber",
		0x82A5 : "MDFileTag",
		0x82A6 : "MDScalePixel",
		0x82A7 : "MDColorTable",
		0x82A8 : "MDLabName",
		0x82A9 : "MDSampleInfo",
		0x82AA : "MDPrepDate",
		0x82AB : "MDPrepTime",
		0x82AC : "MDFileUnits",
		0x830E : "PixelScale",
		0x8335 : "AdventScale",
		0x8336 : "AdventRevision",
		0x835C : "UIC1Tag",
		0x835D : "UIC2Tag",
		0x835E : "UIC3Tag",
		0x835F : "UIC4Tag",
		0x83BB : "IPTC-NAA",
		0x847E : "IntergraphPacketData",
		0x847F : "IntergraphFlagRegisters",
		0x8480 : "IntergraphMatrix",
		0x8481 : "INGRReserved",
		0x8482 : "ModelTiePoint",
		0x84E0 : "Site",
		0x84E1 : "ColorSequence",
		0x84E2 : "IT8Header",
		0x84E3 : "RasterPadding",
		0x84E4 : "BitsPerRunLength",
		0x84E5 : "BitsPerExtendedRunLength",
		0x84E6 : "ColorTable",
		0x84E7 : "ImageColorIndicator",
		0x84E8 : "BackgroundColorIndicator",
		0x84E9 : "ImageColorValue",
		0x84EA : "BackgroundColorValue",
		0x84EB : "PixelIntensityRange",
		0x84EC : "TransparencyIndicator",
		0x84ED : "ColorCharacterization",
		0x84EE : "HCUsage",
		0x84EF : "TrapIndicator",
		0x84F0 : "CMYKEquivalent",
		0x8546 : "SEMInfo",
		0x8568 : "AFCP_IPTC",
		0x85B8 : "PixelMagicJBIGOptions",
		0x85D8 : "ModelTransform",
		0x8602 : "WB_GRGBLevels",
		0x8606 : "LeafData",
		0x8649 : "PhotoshopSettings",
		0x8769 : "ExifOffset",
		0x8773 : "ICC_Profile",
		0x877F : "TIFF_FXExtensions",
		0x8780 : "MultiProfiles",
		0x8781 : "SharedData",
		0x8782 : "T88Options",
		0x87AC : "ImageLayer",
		0x87AF : "GeoTiffDirectory",
		0x87B0 : "GeoTiffDoubleParams",
		0x87B1 : "GeoTiffAsciiParams",
		0x8822 : "ExposureProgram",
		0x8824 : "SpectralSensitivity",
		0x8825 : "GPSInfo",
		0x8827 : "ISO",
		0x8828 : "Opto-ElectricConvFactor",
		0x8829 : "Interlace",
		0x882A : "TimeZoneOffset",
		0x882B : "SelfTimerMode",
		0x8830 : "SensitivityType",
		0x8831 : "StandardOutputSensitivity",
		0x8832 : "RecommendedExposureIndex",
		0x8833 : "ISOSpeed",
		0x8834 : "ISOSpeedLatitudeyyy",
		0x8835 : "ISOSpeedLatitudezzz",
		0x885C : "FaxRecvParams",
		0x885D : "FaxSubAddress",
		0x885E : "FaxRecvTime",
		0x888A : "LeafSubIFD",
		0x9000 : "ExifVersion",
		0x9003 : "DateTimeOriginal",
		0x9004 : "CreateDate",
		0x9101 : "ComponentsConfiguration",
		0x9102 : "CompressedBitsPerPixel",
		0x9201 : "ShutterSpeedValue",
		0x9202 : "ApertureValue",
		0x9203 : "BrightnessValue",
		0x9204 : "ExposureCompensation",
		0x9205 : "MaxApertureValue",
		0x9206 : "SubjectDistance",
		0x9207 : "MeteringMode",
		0x9208 : "LightSource",
		0x9209 : "Flash",
		0x920A : "FocalLength",
		0x920B : "FlashEnergy",
		0x920C : "SpatialFrequencyResponse",
		0x920D : "Noise",
		0x920E : "FocalPlaneXResolution",
		0x920F : "FocalPlaneYResolution",
		0x9210 : "FocalPlaneResolutionUnit",
		0x9211 : "ImageNumber",
		0x9212 : "SecurityClassification",
		0x9213 : "ImageHistory",
		0x9214 : "SubjectArea",
		0x9215 : "ExposureIndex",
		0x9216 : "TIFF-EPStandardID",
		0x9217 : "SensingMethod",
		0x923A : "CIP3DataFile",
		0x923B : "CIP3Sheet",
		0x923C : "CIP3Side",
		0x923F : "StoNits",
		0x927C : "MakerNote",
		0x9286 : "UserComment",
		0x9290 : "SubSecTime",
		0x9291 : "SubSecTimeOriginal",
		0x9292 : "SubSecTimeDigitized",
		0x932F : "MSDocumentText",
		0x9330 : "MSPropertySetStorage",
		0x9331 : "MSDocumentTextPosition",
		0x935C : "ImageSourceData",
		0x9C9B : "XPTitle",
		0x9C9C : "XPComment",
		0x9C9D : "XPAuthor",
		0x9C9E : "XPKeywords",
		0x9C9F : "XPSubject",
		0xA000 : "FlashpixVersion",
		0xA001 : "ColorSpace",
		0xA002 : "ExifImageWidth",
		0xA003 : "ExifImageHeight",
		0xA004 : "RelatedSoundFile",
		0xA005 : "InteropOffset",
		0xA20B : "FlashEnergy",
		0xA20C : "SpatialFrequencyResponse",
		0xA20D : "Noise",
		0xA20E : "FocalPlaneXResolution",
		0xA20F : "FocalPlaneYResolution",
		0xA210 : "FocalPlaneResolutionUnit",
		0xA211 : "ImageNumber",
		0xA212 : "SecurityClassification",
		0xA213 : "ImageHistory",
		0xA214 : "SubjectLocation",
		0xA215 : "ExposureIndex",
		0xA216 : "TIFF-EPStandardID",
		0xA217 : "SensingMethod",
		0xA300 : "FileSource",
		0xA301 : "SceneType",
		0xA302 : "CFAPattern",
		0xA401 : "CustomRendered",
		0xA402 : "ExposureMode",
		0xA403 : "WhiteBalance",
		0xA404 : "DigitalZoomRatio",
		0xA405 : "FocalLengthIn35mmFormat",
		0xA406 : "SceneCaptureType",
		0xA407 : "GainControl",
		0xA408 : "Contrast",
		0xA409 : "Saturation",
		0xA40A : "Sharpness",
		0xA40B : "DeviceSettingDescription",
		0xA40C : "SubjectDistanceRange",
		0xA420 : "ImageUniqueID",
		0xA430 : "OwnerName",
		0xA431 : "SerialNumber",
		0xA432 : "LensInfo",
		0xA433 : "LensMake",
		0xA434 : "LensModel",
		0xA435 : "LensSerialNumber",
		0xA480 : "GDALMetadata",
		0xA481 : "GDALNoData",
		0xA500 : "Gamma",
		0xAFC0 : "ExpandSoftware",
		0xAFC1 : "ExpandLens",
		0xAFC2 : "ExpandFilm",
		0xAFC3 : "ExpandFilterLens",
		0xAFC4 : "ExpandScanner",
		0xAFC5 : "ExpandFlashLamp",
		0xBC01 : "PixelFormat",
		0xBC02 : "Transformation",
		0xBC03 : "Uncompressed",
		0xBC04 : "ImageType",
		0xBC80 : "ImageWidth",
		0xBC81 : "ImageHeight",
		0xBC82 : "WidthResolution",
		0xBC83 : "HeightResolution",
		0xBCC0 : "ImageOffset",
		0xBCC1 : "ImageByteCount",
		0xBCC2 : "AlphaOffset",
		0xBCC3 : "AlphaByteCount",
		0xBCC4 : "ImageDataDiscard",
		0xBCC5 : "AlphaDataDiscard",
		0xC427 : "OceScanjobDesc",
		0xC428 : "OceApplicationSelector",
		0xC429 : "OceIDNumber",
		0xC42A : "OceImageLogic",
		0xC44F : "Annotations",
		0xC4A5 : "PrintIM",
		0xC580 : "USPTOOriginalContentType",
		0xC612 : "DNGVersion",
		0xC613 : "DNGBackwardVersion",
		0xC614 : "UniqueCameraModel",
		0xC615 : "LocalizedCameraModel",
		0xC616 : "CFAPlaneColor",
		0xC617 : "CFALayout",
		0xC618 : "LinearizationTable",
		0xC619 : "BlackLevelRepeatDim",
		0xC61A : "BlackLevel",
		0xC61B : "BlackLevelDeltaH",
		0xC61C : "BlackLevelDeltaV",
		0xC61D : "WhiteLevel",
		0xC61E : "DefaultScale",
		0xC61F : "DefaultCropOrigin",
		0xC620 : "DefaultCropSize",
		0xC621 : "ColorMatrix1",
		0xC622 : "ColorMatrix2",
		0xC623 : "CameraCalibration1",
		0xC624 : "CameraCalibration2",
		0xC625 : "ReductionMatrix1",
		0xC626 : "ReductionMatrix2",
		0xC627 : "AnalogBalance",
		0xC628 : "AsShotNeutral",
		0xC629 : "AsShotWhiteXY",
		0xC62A : "BaselineExposure",
		0xC62B : "BaselineNoise",
		0xC62C : "BaselineSharpness",
		0xC62D : "BayerGreenSplit",
		0xC62E : "LinearResponseLimit",
		0xC62F : "CameraSerialNumber",
		0xC630 : "DNGLensInfo",
		0xC631 : "ChromaBlurRadius",
		0xC632 : "AntiAliasStrength",
		0xC633 : "ShadowScale",
		0xC634 : "DNGPrivateData",
		0xC635 : "MakerNoteSafety",
		0xC640 : "RawImageSegmentation",
		0xC65A : "CalibrationIlluminant1",
		0xC65B : "CalibrationIlluminant2",
		0xC65C : "BestQualityScale",
		0xC65D : "RawDataUniqueID",
		0xC660 : "AliasLayerMetadata",
		0xC68B : "OriginalRawFileName",
		0xC68C : "OriginalRawFileData",
		0xC68D : "ActiveArea",
		0xC68E : "MaskedAreas",
		0xC68F : "AsShotICCProfile",
		0xC690 : "AsShotPreProfileMatrix",
		0xC691 : "CurrentICCProfile",
		0xC692 : "CurrentPreProfileMatrix",
		0xC6BF : "ColorimetricReference",
		0xC6D2 : "PanasonicTitle",
		0xC6D3 : "PanasonicTitle2",
		0xC6F3 : "CameraCalibrationSig",
		0xC6F4 : "ProfileCalibrationSig",
		0xC6F5 : "ProfileIFD",
		0xC6F6 : "AsShotProfileName",
		0xC6F7 : "NoiseReductionApplied",
		0xC6F8 : "ProfileName",
		0xC6F9 : "ProfileHueSatMapDims",
		0xC6FA : "ProfileHueSatMapData1",
		0xC6FB : "ProfileHueSatMapData2",
		0xC6FC : "ProfileToneCurve",
		0xC6FD : "ProfileEmbedPolicy",
		0xC6FE : "ProfileCopyright",
		0xC714 : "ForwardMatrix1",
		0xC715 : "ForwardMatrix2",
		0xC716 : "PreviewApplicationName",
		0xC717 : "PreviewApplicationVersion",
		0xC718 : "PreviewSettingsName",
		0xC719 : "PreviewSettingsDigest",
		0xC71A : "PreviewColorSpace",
		0xC71B : "PreviewDateTime",
		0xC71C : "RawImageDigest",
		0xC71D : "OriginalRawFileDigest",
		0xC71E : "SubTileBlockSize",
		0xC71F : "RowInterleaveFactor",
		0xC725 : "ProfileLookTableDims",
		0xC726 : "ProfileLookTableData",
		0xC740 : "OpcodeList1",
		0xC741 : "OpcodeList2",
		0xC74E : "OpcodeList3",
		0xC761 : "NoiseProfile",
		0xC763 : "TimeCodes",
		0xC764 : "FrameRate",
		0xC772 : "TStop",
		0xC789 : "ReelName",
		0xC791 : "OriginalDefaultFinalSize",
		0xC792 : "OriginalBestQualitySize",
		0xC793 : "OriginalDefaultCropSize",
		0xC7A1 : "CameraLabel",
		0xC7A3 : "ProfileHueSatMapEncoding",
		0xC7A4 : "ProfileLookTableEncoding",
		0xC7A5 : "BaselineExposureOffset",
		0xC7A6 : "DefaultBlackRender",
		0xC7A7 : "NewRawImageDigest",
		0xC7A8 : "RawToPreviewGain",
		0xC7B5 : "DefaultUserCrop",
		0xEA1C : "Padding",
		0xEA1D : "OffsetSchema",
		0xFDE8 : "OwnerName",
		0xFDE9 : "SerialNumber",
		0xFDEA : "Lens",
		0xFE00 : "KDC_IFD",
		0xFE4C : "RawFile",
		0xFE4D : "Converter",
		0xFE4E : "WhiteBalance",
		0xFE51 : "Exposure",
		0xFE52 : "Shadows",
		0xFE53 : "Brightness",
		0xFE54 : "Contrast",
		0xFE55 : "Saturation",
		0xFE56 : "Sharpness",
		0xFE57 : "Smoothness",
		0xFE58 : "MoireFilter"
		
	},
	gps : {	
		0x0000 : 'GPSVersionID',
		0x0001 : 'GPSLatitudeRef',
		0x0002 : 'GPSLatitude',
		0x0003 : 'GPSLongitudeRef',
		0x0004 : 'GPSLongitude',
		0x0005 : 'GPSAltitudeRef',
		0x0006 : 'GPSAltitude',
		0x0007 : 'GPSTimeStamp',
		0x0008 : 'GPSSatellites',
		0x0009 : 'GPSStatus',
		0x000A : 'GPSMeasureMode',
		0x000B : 'GPSDOP',
		0x000C : 'GPSSpeedRef',
		0x000D : 'GPSSpeed',
		0x000E : 'GPSTrackRef',
		0x000F : 'GPSTrack',
		0x0010 : 'GPSImgDirectionRef',
		0x0011 : 'GPSImgDirection',
		0x0012 : 'GPSMapDatum',
		0x0013 : 'GPSDestLatitudeRef',
		0x0014 : 'GPSDestLatitude',
		0x0015 : 'GPSDestLongitudeRef',
		0x0016 : 'GPSDestLongitude',
		0x0017 : 'GPSDestBearingRef',
		0x0018 : 'GPSDestBearing',
		0x0019 : 'GPSDestDistanceRef',
		0x001A : 'GPSDestDistance',
		0x001B : 'GPSProcessingMethod',
		0x001C : 'GPSAreaInformation',
		0x001D : 'GPSDateStamp',
		0x001E : 'GPSDifferential',
		0x001F : 'GPSHPositioningError'
	}
};

/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/exif.js"
(module) {

/*jslint browser: true, devel: true, bitwise: false, debug: true, eqeq: false, es5: true, evil: false, forin: false, newcap: false, nomen: true, plusplus: true, regexp: false, unparam: false, sloppy: true, stupid: false, sub: false, todo: true, vars: true, white: true */

function readExifValue(format, stream) {
	switch(format) {
		case 1: return stream.nextUInt8();
		case 3: return stream.nextUInt16();
		case 4: return stream.nextUInt32();
		case 5: return [stream.nextUInt32(), stream.nextUInt32()];
		case 6: return stream.nextInt8();
		case 8: return stream.nextUInt16();
		case 9: return stream.nextUInt32();
		case 10: return [stream.nextInt32(), stream.nextInt32()];
		case 11: return stream.nextFloat();
		case 12: return stream.nextDouble();
		default: throw new Error('Invalid format while decoding: ' + format);
	}
}

function getBytesPerComponent(format) {
	switch(format) {
		case 1:
		case 2:
		case 6:
		case 7:
			return 1;
		case 3:
		case 8:
			return 2;
		case 4:
		case 9:
		case 11:
			return 4;
		case 5:
		case 10:
		case 12:
			return 8;
		default:
			return 0;
	}
}

function readExifTag(tiffMarker, stream) {
	var tagType = stream.nextUInt16(),
		format = stream.nextUInt16(),
		bytesPerComponent = getBytesPerComponent(format),
		components = stream.nextUInt32(),
		valueBytes = bytesPerComponent * components,
		values,
		value,
		c;

	/* if the value is bigger then 4 bytes, the value is in the data section of the IFD
	and the value present in the tag is the offset starting from the tiff header. So we replace the stream
	with a stream that is located at the given offset in the data section. s*/
	if(valueBytes > 4) {
		stream = tiffMarker.openWithOffset(stream.nextUInt32());
	}
	//we don't want to read strings as arrays
	if(format === 2) {
		values = stream.nextString(components);
		//cut off \0 characters
		var lastNull = values.indexOf('\0');
		if(lastNull !== -1) {
			values = values.substr(0, lastNull);
		}
	}
	else if(format === 7) {
		values = stream.nextBuffer(components);
	}
	else if(format !== 0) {
		values = [];
		for(c = 0; c < components; ++c) {
			values.push(readExifValue(format, stream));
		}
	}
	//since our stream is a stateful object, we need to skip remaining bytes
	//so our offset stays correct
	if(valueBytes < 4) {
		stream.skip(4 - valueBytes);
	}

	return [tagType, values, format];
}

function readIFDSection(tiffMarker, stream, iterator) {
	var numberOfEntries = stream.nextUInt16(), tag, i;
	for(i = 0; i < numberOfEntries; ++i) {
		tag = readExifTag(tiffMarker, stream);
		iterator(tag[0], tag[1], tag[2]);
	}
}

function readHeader(stream) {
	var exifHeader = stream.nextString(6);
	if(exifHeader !== 'Exif\0\0') {
		throw new Error('Invalid EXIF header');
	}

	var tiffMarker = stream.mark();
	var tiffHeader = stream.nextUInt16();
	if(tiffHeader === 0x4949) {
		stream.setBigEndian(false);
	} else if(tiffHeader === 0x4D4D) {
		stream.setBigEndian(true);
	} else {
		throw new Error('Invalid TIFF header');
	}
	if(stream.nextUInt16() !== 0x002A) {
		throw new Error('Invalid TIFF data');
	}
	return tiffMarker;
}

module.exports = {
	IFD0: 1,
	IFD1: 2,
	GPSIFD: 3,
	SubIFD: 4,
	InteropIFD: 5,
	parseTags: function(stream, iterator) {
		var tiffMarker;
		try {
			tiffMarker = readHeader(stream);
		} catch(e) {
			return false;	//ignore APP1 sections with invalid headers
		}
		var subIfdOffset, gpsOffset, interopOffset;
		var ifd0Stream = tiffMarker.openWithOffset(stream.nextUInt32()),
			IFD0 = this.IFD0;
		readIFDSection(tiffMarker, ifd0Stream, function(tagType, value, format) {
			switch(tagType) {
				case 0x8825: gpsOffset = value[0]; break;
				case 0x8769: subIfdOffset = value[0]; break;
				default: iterator(IFD0, tagType, value, format); break;
			}
		});
		var ifd1Offset = ifd0Stream.nextUInt32();
		if(ifd1Offset !== 0) {
			var ifd1Stream = tiffMarker.openWithOffset(ifd1Offset);
			readIFDSection(tiffMarker, ifd1Stream, iterator.bind(null, this.IFD1));
		}

		if(gpsOffset) {
			var gpsStream = tiffMarker.openWithOffset(gpsOffset);
			readIFDSection(tiffMarker, gpsStream, iterator.bind(null, this.GPSIFD));
		}

		if(subIfdOffset) {
			var subIfdStream = tiffMarker.openWithOffset(subIfdOffset), InteropIFD = this.InteropIFD;
			readIFDSection(tiffMarker, subIfdStream, function(tagType, value, format) {
				if(tagType === 0xA005) {
					interopOffset = value[0];
				} else {
					iterator(InteropIFD, tagType, value, format);
				}
			});
		}

		if(interopOffset) {
			var interopStream = tiffMarker.openWithOffset(interopOffset);
			readIFDSection(tiffMarker, interopStream, iterator.bind(null, this.InteropIFD));
		}
		return true;
	}
};

/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/jpeg.js"
(module) {

/*jslint browser: true, devel: true, bitwise: false, debug: true, eqeq: false, es5: true, evil: false, forin: false, newcap: false, nomen: true, plusplus: true, regexp: false, unparam: false, sloppy: true, stupid: false, sub: false, todo: true, vars: true, white: true */

module.exports = {
	parseSections: function(stream, iterator) {
		var len, markerType;
		stream.setBigEndian(true);
		//stop reading the stream at the SOS (Start of Stream) marker,
		//because its length is not stored in the header so we can't
		//know where to jump to. The only marker after that is just EOI (End Of Image) anyway
		while(stream.remainingLength() > 0 && markerType !== 0xDA) {
			if(stream.nextUInt8() !== 0xFF) {
				throw new Error('Invalid JPEG section offset');
			}
			markerType = stream.nextUInt8();
			//don't read size from markers that have no datas
			if((markerType >= 0xD0 && markerType <= 0xD9) || markerType === 0xDA) {
				len = 0;
			} else {
				len = stream.nextUInt16() - 2;
			}
			iterator(markerType, stream.branch(0, len));
			stream.skip(len);
		}
	},
	//stream should be located after SOF section size and in big endian mode, like passed to parseSections iterator
	getSizeFromSOFSection: function(stream) {
		stream.skip(1);
		return {
			height: stream.nextUInt16(),
			width: stream.nextUInt16()
		};
	},
	getSectionName: function(markerType) {
		var name, index;
		switch(markerType) {
			case 0xD8: name = 'SOI'; break;
			case 0xC4: name = 'DHT'; break;
			case 0xDB: name = 'DQT'; break;
			case 0xDD: name = 'DRI'; break;
			case 0xDA: name = 'SOS'; break;
			case 0xFE: name = 'COM'; break;
			case 0xD9: name = 'EOI'; break;
			default:
				if(markerType >= 0xE0 && markerType <= 0xEF) {
					name = 'APP';
					index = markerType - 0xE0;
				}
				else if(markerType >= 0xC0 && markerType <= 0xCF && markerType !== 0xC4 && markerType !== 0xC8 && markerType !== 0xCC) {
					name = 'SOF';
					index = markerType - 0xC0;
				}
				else if(markerType >= 0xD0 && markerType <= 0xD7) {
					name = 'RST';
					index = markerType - 0xD0;
				}
				break;
		}
		var nameStruct = {
			name: name
		};
		if(typeof index === 'number') {
			nameStruct.index = index;
		}
		return nameStruct;
	}
};

/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/parser.js"
(module, __unused_webpack_exports, __webpack_require__) {

/*jslint browser: true, devel: true, bitwise: false, debug: true, eqeq: false, es5: true, evil: false, forin: false, newcap: false, nomen: true, plusplus: true, regexp: false, unparam: false, sloppy: true, stupid: false, sub: false, todo: true, vars: true, white: true */

var jpeg = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/jpeg.js"),
	exif = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/exif.js"),
	simplify = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/simplify.js");

function ExifResult(startMarker, tags, imageSize, thumbnailOffset, thumbnailLength, thumbnailType, app1Offset) {
	this.startMarker = startMarker;
	this.tags = tags;
	this.imageSize = imageSize;
	this.thumbnailOffset = thumbnailOffset;
	this.thumbnailLength = thumbnailLength;
	this.thumbnailType = thumbnailType;
	this.app1Offset = app1Offset;
}

ExifResult.prototype = {
	hasThumbnail: function(mime) {
		if(!this.thumbnailOffset || !this.thumbnailLength) {
			return false;
		}
		if(typeof mime !== 'string') {
			return true;
		}
		if(mime.toLowerCase().trim() === 'image/jpeg') {
			return this.thumbnailType === 6;
		}
		if(mime.toLowerCase().trim() === 'image/tiff') {
			return this.thumbnailType === 1;
		}
		return false;
	},
	getThumbnailOffset: function() {
		return this.app1Offset + 6 + this.thumbnailOffset;
	},
	getThumbnailLength: function() {
		return this.thumbnailLength;
	},
	getThumbnailBuffer: function() {
		return this._getThumbnailStream().nextBuffer(this.thumbnailLength);
	},
	_getThumbnailStream: function() {
		return this.startMarker.openWithOffset(this.getThumbnailOffset());
	},
	getImageSize: function() {
		return this.imageSize;
	},
	getThumbnailSize: function() {
		var stream = this._getThumbnailStream(), size;
		jpeg.parseSections(stream, function(sectionType, sectionStream) {
			if(jpeg.getSectionName(sectionType).name === 'SOF') {
				size = jpeg.getSizeFromSOFSection(sectionStream);
			}
		});
		return size;
	}
};

function Parser(stream) {
	this.stream = stream;
	this.flags = {
		readBinaryTags: false,
		resolveTagNames: true,
		simplifyValues: true,
		imageSize: true,
		hidePointers: true,
		returnTags: true
	};
}

Parser.prototype = {
	enableBinaryFields: function(enable) {
		this.flags.readBinaryTags = !!enable;
		return this;
	},
	enablePointers: function(enable) {
		this.flags.hidePointers = !enable;
		return this;
	},
	enableTagNames: function(enable) {
		this.flags.resolveTagNames = !!enable;
		return this;
	},
	enableImageSize: function(enable) {
		this.flags.imageSize = !!enable;
		return this;
	},
	enableReturnTags: function(enable) {
		this.flags.returnTags = !!enable;
		return this;
	},
	enableSimpleValues: function(enable) {
		this.flags.simplifyValues = !!enable;
		return this;
	},
	parse: function() {
		var start = this.stream.mark(),
			stream = start.openWithOffset(0),
			flags = this.flags,
			tags,
			imageSize,
			thumbnailOffset,
			thumbnailLength,
			thumbnailType,
			app1Offset,
			tagNames,
			getTagValue, setTagValue;
		if(flags.resolveTagNames) {
			tagNames = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/exif-tags.js");
		}
		if(flags.resolveTagNames) {
			tags = {};
			getTagValue = function(t) {
				return tags[t.name];
			};
			setTagValue = function(t, value) {
				tags[t.name] = value;
			};
		} else {
			tags = [];
			getTagValue = function(t) {
				var i;
				for(i = 0; i < tags.length; ++i) {
					if(tags[i].type === t.type && tags[i].section === t.section) {
						return tags.value;
					}
				}
			};
			setTagValue = function(t, value) {
				var i;
				for(i = 0; i < tags.length; ++i) {
					if(tags[i].type === t.type && tags[i].section === t.section) {
						tags.value = value;
						return;
					}
				}
			};
		}

		jpeg.parseSections(stream, function(sectionType, sectionStream) {
			var validExifHeaders, sectionOffset = sectionStream.offsetFrom(start);
			if(sectionType === 0xE1) {
				validExifHeaders = exif.parseTags(sectionStream, function(ifdSection, tagType, value, format) {
					//ignore binary fields if disabled
					if(!flags.readBinaryTags && format === 7) {
						return;
					}

					if(tagType === 0x0201) {
						thumbnailOffset = value[0];
						if(flags.hidePointers) {return;}
					} else if(tagType === 0x0202) {
						thumbnailLength = value[0];
						if(flags.hidePointers) {return;}
					} else if(tagType === 0x0103) {
						thumbnailType = value[0];
						if(flags.hidePointers) {return;}
					}
					//if flag is set to not store tags, return here after storing pointers
					if(!flags.returnTags) {
						return;
					}

					if(flags.simplifyValues) {
						value = simplify.simplifyValue(value, format);
					}
					if(flags.resolveTagNames) {
						var sectionTagNames = ifdSection === exif.GPSIFD ? tagNames.gps : tagNames.exif;
						var name = sectionTagNames[tagType];
						if(!name) {
							name = tagNames.exif[tagType];
						}
						if (!tags.hasOwnProperty(name)) {
							tags[name] = value;
						}
					} else {
						tags.push({
							section: ifdSection,
							type: tagType,
							value: value
						});
					}
				});
				if(validExifHeaders) {
					app1Offset = sectionOffset;
				}
			}
			else if(flags.imageSize && jpeg.getSectionName(sectionType).name === 'SOF') {
				imageSize = jpeg.getSizeFromSOFSection(sectionStream);
			}
		});

		if(flags.simplifyValues) {
			simplify.castDegreeValues(getTagValue, setTagValue);
			simplify.castDateValues(getTagValue, setTagValue);
		}

		return new ExifResult(start, tags, imageSize, thumbnailOffset, thumbnailLength, thumbnailType, app1Offset);
	}
};



module.exports = Parser;


/***/ },

/***/ "../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/simplify.js"
(module, __unused_webpack_exports, __webpack_require__) {

var exif = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/exif.js");
var date = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/lib/date.js");

var degreeTags = [{
	section: exif.GPSIFD,
	type: 0x0002,
	name: 'GPSLatitude',
	refType: 0x0001,
	refName: 'GPSLatitudeRef',
	posVal: 'N'
},
{
	section: exif.GPSIFD,
	type: 0x0004,
	name: 'GPSLongitude',
	refType: 0x0003,
	refName: 'GPSLongitudeRef',
	posVal: 'E'
}];
var dateTags = [{
	section: exif.SubIFD,
	type: 0x0132,
	name: 'ModifyDate'
},
{
	section: exif.SubIFD,
	type: 0x9003,
	name: 'DateTimeOriginal'
},
{
	section: exif.SubIFD,
	type: 0x9004,
	name: 'CreateDate'
},
{
	section: exif.SubIFD,
	type: 0x0132,
	name : 'ModifyDate',
}];

module.exports = {
	castDegreeValues: function(getTagValue, setTagValue) {
		degreeTags.forEach(function(t) {
			var degreeVal = getTagValue(t);
			if(degreeVal) {
				var degreeRef = getTagValue({section: t.section, type: t.refType, name: t.refName});
				var degreeNumRef = degreeRef === t.posVal ? 1 : -1;
				var degree = (degreeVal[0] + (degreeVal[1] / 60) + (degreeVal[2] / 3600)) * degreeNumRef;
				setTagValue(t, degree);
			}
		});
	},
	castDateValues: function(getTagValue, setTagValue) {
		dateTags.forEach(function(t) {
			var dateStrVal = getTagValue(t);
			if(dateStrVal) {
				//some easy checks to determine two common date formats
				var timestamp = date.parseExifDate(dateStrVal);
				if(typeof timestamp !== 'undefined') {
					setTagValue(t, timestamp);
				}
			}
		});
	},
	simplifyValue: function(values, format) {
		if(Array.isArray(values)) {
			values = values.map(function(value) {
				if(format === 10 || format === 5) {
					return value[0] / value[1];
				}
				return value;
			});
			if(values.length === 1) {
				values = values[0];
			}
		}
		return values;
	}
};


/***/ },

/***/ "../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/core.js"
(module, __unused_webpack_exports, __webpack_require__) {

"use strict";

const Token = __webpack_require__("../../node_modules/.pnpm/token-types@4.2.1/node_modules/token-types/lib/index.js");
const strtok3 = __webpack_require__("../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/core.js");
const {
	stringToBytes,
	tarHeaderChecksumMatches,
	uint32SyncSafeToken
} = __webpack_require__("../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/util.js");
const supported = __webpack_require__("../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/supported.js");

const minimumBytes = 4100; // A fair amount of file-types are detectable within this range

async function fromStream(stream) {
	const tokenizer = await strtok3.fromStream(stream);
	try {
		return await fromTokenizer(tokenizer);
	} finally {
		await tokenizer.close();
	}
}

async function fromBuffer(input) {
	if (!(input instanceof Uint8Array || input instanceof ArrayBuffer || Buffer.isBuffer(input))) {
		throw new TypeError(`Expected the \`input\` argument to be of type \`Uint8Array\` or \`Buffer\` or \`ArrayBuffer\`, got \`${typeof input}\``);
	}

	const buffer = input instanceof Buffer ? input : Buffer.from(input);

	if (!(buffer && buffer.length > 1)) {
		return;
	}

	const tokenizer = strtok3.fromBuffer(buffer);
	return fromTokenizer(tokenizer);
}

function _check(buffer, headers, options) {
	options = {
		offset: 0,
		...options
	};

	for (const [index, header] of headers.entries()) {
		// If a bitmask is set
		if (options.mask) {
			// If header doesn't equal `buf` with bits masked off
			if (header !== (options.mask[index] & buffer[index + options.offset])) {
				return false;
			}
		} else if (header !== buffer[index + options.offset]) {
			return false;
		}
	}

	return true;
}

async function fromTokenizer(tokenizer) {
	try {
		return _fromTokenizer(tokenizer);
	} catch (error) {
		if (!(error instanceof strtok3.EndOfStreamError)) {
			throw error;
		}
	}
}

async function _fromTokenizer(tokenizer) {
	let buffer = Buffer.alloc(minimumBytes);
	const bytesRead = 12;
	const check = (header, options) => _check(buffer, header, options);
	const checkString = (header, options) => check(stringToBytes(header), options);

	// Keep reading until EOF if the file size is unknown.
	if (!tokenizer.fileInfo.size) {
		tokenizer.fileInfo.size = Number.MAX_SAFE_INTEGER;
	}

	await tokenizer.peekBuffer(buffer, {length: bytesRead, mayBeLess: true});

	// -- 2-byte signatures --

	if (check([0x42, 0x4D])) {
		return {
			ext: 'bmp',
			mime: 'image/bmp'
		};
	}

	if (check([0x0B, 0x77])) {
		return {
			ext: 'ac3',
			mime: 'audio/vnd.dolby.dd-raw'
		};
	}

	if (check([0x78, 0x01])) {
		return {
			ext: 'dmg',
			mime: 'application/x-apple-diskimage'
		};
	}

	if (check([0x4D, 0x5A])) {
		return {
			ext: 'exe',
			mime: 'application/x-msdownload'
		};
	}

	if (check([0x25, 0x21])) {
		await tokenizer.peekBuffer(buffer, {length: 24, mayBeLess: true});

		if (checkString('PS-Adobe-', {offset: 2}) &&
			checkString(' EPSF-', {offset: 14})) {
			return {
				ext: 'eps',
				mime: 'application/eps'
			};
		}

		return {
			ext: 'ps',
			mime: 'application/postscript'
		};
	}

	if (
		check([0x1F, 0xA0]) ||
		check([0x1F, 0x9D])
	) {
		return {
			ext: 'Z',
			mime: 'application/x-compress'
		};
	}

	// -- 3-byte signatures --

	if (check([0xFF, 0xD8, 0xFF])) {
		return {
			ext: 'jpg',
			mime: 'image/jpeg'
		};
	}

	if (check([0x49, 0x49, 0xBC])) {
		return {
			ext: 'jxr',
			mime: 'image/vnd.ms-photo'
		};
	}

	if (check([0x1F, 0x8B, 0x8])) {
		return {
			ext: 'gz',
			mime: 'application/gzip'
		};
	}

	if (check([0x42, 0x5A, 0x68])) {
		return {
			ext: 'bz2',
			mime: 'application/x-bzip2'
		};
	}

	if (checkString('ID3')) {
		await tokenizer.ignore(6); // Skip ID3 header until the header size
		const id3HeaderLen = await tokenizer.readToken(uint32SyncSafeToken);
		if (tokenizer.position + id3HeaderLen > tokenizer.fileInfo.size) {
			// Guess file type based on ID3 header for backward compatibility
			return {
				ext: 'mp3',
				mime: 'audio/mpeg'
			};
		}

		await tokenizer.ignore(id3HeaderLen);
		return fromTokenizer(tokenizer); // Skip ID3 header, recursion
	}

	// Musepack, SV7
	if (checkString('MP+')) {
		return {
			ext: 'mpc',
			mime: 'audio/x-musepack'
		};
	}

	if (
		(buffer[0] === 0x43 || buffer[0] === 0x46) &&
		check([0x57, 0x53], {offset: 1})
	) {
		return {
			ext: 'swf',
			mime: 'application/x-shockwave-flash'
		};
	}

	// -- 4-byte signatures --

	if (check([0x47, 0x49, 0x46])) {
		return {
			ext: 'gif',
			mime: 'image/gif'
		};
	}

	if (checkString('FLIF')) {
		return {
			ext: 'flif',
			mime: 'image/flif'
		};
	}

	if (checkString('8BPS')) {
		return {
			ext: 'psd',
			mime: 'image/vnd.adobe.photoshop'
		};
	}

	if (checkString('WEBP', {offset: 8})) {
		return {
			ext: 'webp',
			mime: 'image/webp'
		};
	}

	// Musepack, SV8
	if (checkString('MPCK')) {
		return {
			ext: 'mpc',
			mime: 'audio/x-musepack'
		};
	}

	if (checkString('FORM')) {
		return {
			ext: 'aif',
			mime: 'audio/aiff'
		};
	}

	if (checkString('icns', {offset: 0})) {
		return {
			ext: 'icns',
			mime: 'image/icns'
		};
	}

	// Zip-based file formats
	// Need to be before the `zip` check
	if (check([0x50, 0x4B, 0x3, 0x4])) { // Local file header signature
		try {
			while (tokenizer.position + 30 < tokenizer.fileInfo.size) {
				await tokenizer.readBuffer(buffer, {length: 30});

				// https://en.wikipedia.org/wiki/Zip_(file_format)#File_headers
				const zipHeader = {
					compressedSize: buffer.readUInt32LE(18),
					uncompressedSize: buffer.readUInt32LE(22),
					filenameLength: buffer.readUInt16LE(26),
					extraFieldLength: buffer.readUInt16LE(28)
				};

				zipHeader.filename = await tokenizer.readToken(new Token.StringType(zipHeader.filenameLength, 'utf-8'));
				await tokenizer.ignore(zipHeader.extraFieldLength);

				// Assumes signed `.xpi` from addons.mozilla.org
				if (zipHeader.filename === 'META-INF/mozilla.rsa') {
					return {
						ext: 'xpi',
						mime: 'application/x-xpinstall'
					};
				}

				if (zipHeader.filename.endsWith('.rels') || zipHeader.filename.endsWith('.xml')) {
					const type = zipHeader.filename.split('/')[0];
					switch (type) {
						case '_rels':
							break;
						case 'word':
							return {
								ext: 'docx',
								mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
							};
						case 'ppt':
							return {
								ext: 'pptx',
								mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
							};
						case 'xl':
							return {
								ext: 'xlsx',
								mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
							};
						default:
							break;
					}
				}

				if (zipHeader.filename.startsWith('xl/')) {
					return {
						ext: 'xlsx',
						mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
					};
				}

				if (zipHeader.filename.startsWith('3D/') && zipHeader.filename.endsWith('.model')) {
					return {
						ext: '3mf',
						mime: 'model/3mf'
					};
				}

				// The docx, xlsx and pptx file types extend the Office Open XML file format:
				// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
				// We look for:
				// - one entry named '[Content_Types].xml' or '_rels/.rels',
				// - one entry indicating specific type of file.
				// MS Office, OpenOffice and LibreOffice may put the parts in different order, so the check should not rely on it.
				if (zipHeader.filename === 'mimetype' && zipHeader.compressedSize === zipHeader.uncompressedSize) {
					const mimeType = await tokenizer.readToken(new Token.StringType(zipHeader.compressedSize, 'utf-8'));

					switch (mimeType) {
						case 'application/epub+zip':
							return {
								ext: 'epub',
								mime: 'application/epub+zip'
							};
						case 'application/vnd.oasis.opendocument.text':
							return {
								ext: 'odt',
								mime: 'application/vnd.oasis.opendocument.text'
							};
						case 'application/vnd.oasis.opendocument.spreadsheet':
							return {
								ext: 'ods',
								mime: 'application/vnd.oasis.opendocument.spreadsheet'
							};
						case 'application/vnd.oasis.opendocument.presentation':
							return {
								ext: 'odp',
								mime: 'application/vnd.oasis.opendocument.presentation'
							};
						default:
					}
				}

				// Try to find next header manually when current one is corrupted
				if (zipHeader.compressedSize === 0) {
					let nextHeaderIndex = -1;

					while (nextHeaderIndex < 0 && (tokenizer.position < tokenizer.fileInfo.size)) {
						await tokenizer.peekBuffer(buffer, {mayBeLess: true});

						nextHeaderIndex = buffer.indexOf('504B0304', 0, 'hex');
						// Move position to the next header if found, skip the whole buffer otherwise
						await tokenizer.ignore(nextHeaderIndex >= 0 ? nextHeaderIndex : buffer.length);
					}
				} else {
					await tokenizer.ignore(zipHeader.compressedSize);
				}
			}
		} catch (error) {
			if (!(error instanceof strtok3.EndOfStreamError)) {
				throw error;
			}
		}

		return {
			ext: 'zip',
			mime: 'application/zip'
		};
	}

	if (checkString('OggS')) {
		// This is an OGG container
		await tokenizer.ignore(28);
		const type = Buffer.alloc(8);
		await tokenizer.readBuffer(type);

		// Needs to be before `ogg` check
		if (_check(type, [0x4F, 0x70, 0x75, 0x73, 0x48, 0x65, 0x61, 0x64])) {
			return {
				ext: 'opus',
				mime: 'audio/opus'
			};
		}

		// If ' theora' in header.
		if (_check(type, [0x80, 0x74, 0x68, 0x65, 0x6F, 0x72, 0x61])) {
			return {
				ext: 'ogv',
				mime: 'video/ogg'
			};
		}

		// If '\x01video' in header.
		if (_check(type, [0x01, 0x76, 0x69, 0x64, 0x65, 0x6F, 0x00])) {
			return {
				ext: 'ogm',
				mime: 'video/ogg'
			};
		}

		// If ' FLAC' in header  https://xiph.org/flac/faq.html
		if (_check(type, [0x7F, 0x46, 0x4C, 0x41, 0x43])) {
			return {
				ext: 'oga',
				mime: 'audio/ogg'
			};
		}

		// 'Speex  ' in header https://en.wikipedia.org/wiki/Speex
		if (_check(type, [0x53, 0x70, 0x65, 0x65, 0x78, 0x20, 0x20])) {
			return {
				ext: 'spx',
				mime: 'audio/ogg'
			};
		}

		// If '\x01vorbis' in header
		if (_check(type, [0x01, 0x76, 0x6F, 0x72, 0x62, 0x69, 0x73])) {
			return {
				ext: 'ogg',
				mime: 'audio/ogg'
			};
		}

		// Default OGG container https://www.iana.org/assignments/media-types/application/ogg
		return {
			ext: 'ogx',
			mime: 'application/ogg'
		};
	}

	if (
		check([0x50, 0x4B]) &&
		(buffer[2] === 0x3 || buffer[2] === 0x5 || buffer[2] === 0x7) &&
		(buffer[3] === 0x4 || buffer[3] === 0x6 || buffer[3] === 0x8)
	) {
		return {
			ext: 'zip',
			mime: 'application/zip'
		};
	}

	//

	// File Type Box (https://en.wikipedia.org/wiki/ISO_base_media_file_format)
	// It's not required to be first, but it's recommended to be. Almost all ISO base media files start with `ftyp` box.
	// `ftyp` box must contain a brand major identifier, which must consist of ISO 8859-1 printable characters.
	// Here we check for 8859-1 printable characters (for simplicity, it's a mask which also catches one non-printable character).
	if (
		checkString('ftyp', {offset: 4}) &&
		(buffer[8] & 0x60) !== 0x00 // Brand major, first character ASCII?
	) {
		// They all can have MIME `video/mp4` except `application/mp4` special-case which is hard to detect.
		// For some cases, we're specific, everything else falls to `video/mp4` with `mp4` extension.
		const brandMajor = buffer.toString('binary', 8, 12).replace('\0', ' ').trim();
		switch (brandMajor) {
			case 'avif':
				return {ext: 'avif', mime: 'image/avif'};
			case 'mif1':
				return {ext: 'heic', mime: 'image/heif'};
			case 'msf1':
				return {ext: 'heic', mime: 'image/heif-sequence'};
			case 'heic':
			case 'heix':
				return {ext: 'heic', mime: 'image/heic'};
			case 'hevc':
			case 'hevx':
				return {ext: 'heic', mime: 'image/heic-sequence'};
			case 'qt':
				return {ext: 'mov', mime: 'video/quicktime'};
			case 'M4V':
			case 'M4VH':
			case 'M4VP':
				return {ext: 'm4v', mime: 'video/x-m4v'};
			case 'M4P':
				return {ext: 'm4p', mime: 'video/mp4'};
			case 'M4B':
				return {ext: 'm4b', mime: 'audio/mp4'};
			case 'M4A':
				return {ext: 'm4a', mime: 'audio/x-m4a'};
			case 'F4V':
				return {ext: 'f4v', mime: 'video/mp4'};
			case 'F4P':
				return {ext: 'f4p', mime: 'video/mp4'};
			case 'F4A':
				return {ext: 'f4a', mime: 'audio/mp4'};
			case 'F4B':
				return {ext: 'f4b', mime: 'audio/mp4'};
			case 'crx':
				return {ext: 'cr3', mime: 'image/x-canon-cr3'};
			default:
				if (brandMajor.startsWith('3g')) {
					if (brandMajor.startsWith('3g2')) {
						return {ext: '3g2', mime: 'video/3gpp2'};
					}

					return {ext: '3gp', mime: 'video/3gpp'};
				}

				return {ext: 'mp4', mime: 'video/mp4'};
		}
	}

	if (checkString('MThd')) {
		return {
			ext: 'mid',
			mime: 'audio/midi'
		};
	}

	if (
		checkString('wOFF') &&
		(
			check([0x00, 0x01, 0x00, 0x00], {offset: 4}) ||
			checkString('OTTO', {offset: 4})
		)
	) {
		return {
			ext: 'woff',
			mime: 'font/woff'
		};
	}

	if (
		checkString('wOF2') &&
		(
			check([0x00, 0x01, 0x00, 0x00], {offset: 4}) ||
			checkString('OTTO', {offset: 4})
		)
	) {
		return {
			ext: 'woff2',
			mime: 'font/woff2'
		};
	}

	if (check([0xD4, 0xC3, 0xB2, 0xA1]) || check([0xA1, 0xB2, 0xC3, 0xD4])) {
		return {
			ext: 'pcap',
			mime: 'application/vnd.tcpdump.pcap'
		};
	}

	// Sony DSD Stream File (DSF)
	if (checkString('DSD ')) {
		return {
			ext: 'dsf',
			mime: 'audio/x-dsf' // Non-standard
		};
	}

	if (checkString('LZIP')) {
		return {
			ext: 'lz',
			mime: 'application/x-lzip'
		};
	}

	if (checkString('fLaC')) {
		return {
			ext: 'flac',
			mime: 'audio/x-flac'
		};
	}

	if (check([0x42, 0x50, 0x47, 0xFB])) {
		return {
			ext: 'bpg',
			mime: 'image/bpg'
		};
	}

	if (checkString('wvpk')) {
		return {
			ext: 'wv',
			mime: 'audio/wavpack'
		};
	}

	if (checkString('%PDF')) {
		await tokenizer.ignore(1350);
		const maxBufferSize = 10 * 1024 * 1024;
		const buffer = Buffer.alloc(Math.min(maxBufferSize, tokenizer.fileInfo.size));
		await tokenizer.readBuffer(buffer, {mayBeLess: true});

		// Check if this is an Adobe Illustrator file
		if (buffer.includes(Buffer.from('AIPrivateData'))) {
			return {
				ext: 'ai',
				mime: 'application/postscript'
			};
		}

		// Assume this is just a normal PDF
		return {
			ext: 'pdf',
			mime: 'application/pdf'
		};
	}

	if (check([0x00, 0x61, 0x73, 0x6D])) {
		return {
			ext: 'wasm',
			mime: 'application/wasm'
		};
	}

	// TIFF, little-endian type
	if (check([0x49, 0x49, 0x2A, 0x0])) {
		if (checkString('CR', {offset: 8})) {
			return {
				ext: 'cr2',
				mime: 'image/x-canon-cr2'
			};
		}

		if (check([0x1C, 0x00, 0xFE, 0x00], {offset: 8}) || check([0x1F, 0x00, 0x0B, 0x00], {offset: 8})) {
			return {
				ext: 'nef',
				mime: 'image/x-nikon-nef'
			};
		}

		if (
			check([0x08, 0x00, 0x00, 0x00], {offset: 4}) &&
			(check([0x2D, 0x00, 0xFE, 0x00], {offset: 8}) ||
				check([0x27, 0x00, 0xFE, 0x00], {offset: 8}))
		) {
			return {
				ext: 'dng',
				mime: 'image/x-adobe-dng'
			};
		}

		buffer = Buffer.alloc(24);
		await tokenizer.peekBuffer(buffer);
		if (
			(check([0x10, 0xFB, 0x86, 0x01], {offset: 4}) || check([0x08, 0x00, 0x00, 0x00], {offset: 4})) &&
			// This pattern differentiates ARW from other TIFF-ish file types:
			check([0x00, 0xFE, 0x00, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x03, 0x01], {offset: 9})
		) {
			return {
				ext: 'arw',
				mime: 'image/x-sony-arw'
			};
		}

		return {
			ext: 'tif',
			mime: 'image/tiff'
		};
	}

	// TIFF, big-endian type
	if (check([0x4D, 0x4D, 0x0, 0x2A])) {
		return {
			ext: 'tif',
			mime: 'image/tiff'
		};
	}

	if (checkString('MAC ')) {
		return {
			ext: 'ape',
			mime: 'audio/ape'
		};
	}

	// https://github.com/threatstack/libmagic/blob/master/magic/Magdir/matroska
	if (check([0x1A, 0x45, 0xDF, 0xA3])) { // Root element: EBML
		async function readField() {
			const msb = await tokenizer.peekNumber(Token.UINT8);
			let mask = 0x80;
			let ic = 0; // 0 = A, 1 = B, 2 = C, 3 = D

			while ((msb & mask) === 0 && mask !== 0) {
				++ic;
				mask >>= 1;
			}

			const id = Buffer.alloc(ic + 1);
			await tokenizer.readBuffer(id);
			return id;
		}

		async function readElement() {
			const id = await readField();
			const lenField = await readField();
			lenField[0] ^= 0x80 >> (lenField.length - 1);
			const nrLen = Math.min(6, lenField.length); // JavaScript can max read 6 bytes integer
			return {
				id: id.readUIntBE(0, id.length),
				len: lenField.readUIntBE(lenField.length - nrLen, nrLen)
			};
		}

		async function readChildren(level, children) {
			while (children > 0) {
				const e = await readElement();
				if (e.id === 0x4282) {
					return tokenizer.readToken(new Token.StringType(e.len, 'utf-8')); // Return DocType
				}

				await tokenizer.ignore(e.len); // ignore payload
				--children;
			}
		}

		const re = await readElement();
		const docType = await readChildren(1, re.len);

		switch (docType) {
			case 'webm':
				return {
					ext: 'webm',
					mime: 'video/webm'
				};

			case 'matroska':
				return {
					ext: 'mkv',
					mime: 'video/x-matroska'
				};

			default:
				return;
		}
	}

	// RIFF file format which might be AVI, WAV, QCP, etc
	if (check([0x52, 0x49, 0x46, 0x46])) {
		if (check([0x41, 0x56, 0x49], {offset: 8})) {
			return {
				ext: 'avi',
				mime: 'video/vnd.avi'
			};
		}

		if (check([0x57, 0x41, 0x56, 0x45], {offset: 8})) {
			return {
				ext: 'wav',
				mime: 'audio/vnd.wave'
			};
		}

		// QLCM, QCP file
		if (check([0x51, 0x4C, 0x43, 0x4D], {offset: 8})) {
			return {
				ext: 'qcp',
				mime: 'audio/qcelp'
			};
		}
	}

	if (checkString('SQLi')) {
		return {
			ext: 'sqlite',
			mime: 'application/x-sqlite3'
		};
	}

	if (check([0x4E, 0x45, 0x53, 0x1A])) {
		return {
			ext: 'nes',
			mime: 'application/x-nintendo-nes-rom'
		};
	}

	if (checkString('Cr24')) {
		return {
			ext: 'crx',
			mime: 'application/x-google-chrome-extension'
		};
	}

	if (
		checkString('MSCF') ||
		checkString('ISc(')
	) {
		return {
			ext: 'cab',
			mime: 'application/vnd.ms-cab-compressed'
		};
	}

	if (check([0xED, 0xAB, 0xEE, 0xDB])) {
		return {
			ext: 'rpm',
			mime: 'application/x-rpm'
		};
	}

	if (check([0xC5, 0xD0, 0xD3, 0xC6])) {
		return {
			ext: 'eps',
			mime: 'application/eps'
		};
	}

	if (check([0x28, 0xB5, 0x2F, 0xFD])) {
		return {
			ext: 'zst',
			mime: 'application/zstd'
		};
	}

	// -- 5-byte signatures --

	if (check([0x4F, 0x54, 0x54, 0x4F, 0x00])) {
		return {
			ext: 'otf',
			mime: 'font/otf'
		};
	}

	if (checkString('#!AMR')) {
		return {
			ext: 'amr',
			mime: 'audio/amr'
		};
	}

	if (checkString('{\\rtf')) {
		return {
			ext: 'rtf',
			mime: 'application/rtf'
		};
	}

	if (check([0x46, 0x4C, 0x56, 0x01])) {
		return {
			ext: 'flv',
			mime: 'video/x-flv'
		};
	}

	if (checkString('IMPM')) {
		return {
			ext: 'it',
			mime: 'audio/x-it'
		};
	}

	if (
		checkString('-lh0-', {offset: 2}) ||
		checkString('-lh1-', {offset: 2}) ||
		checkString('-lh2-', {offset: 2}) ||
		checkString('-lh3-', {offset: 2}) ||
		checkString('-lh4-', {offset: 2}) ||
		checkString('-lh5-', {offset: 2}) ||
		checkString('-lh6-', {offset: 2}) ||
		checkString('-lh7-', {offset: 2}) ||
		checkString('-lzs-', {offset: 2}) ||
		checkString('-lz4-', {offset: 2}) ||
		checkString('-lz5-', {offset: 2}) ||
		checkString('-lhd-', {offset: 2})
	) {
		return {
			ext: 'lzh',
			mime: 'application/x-lzh-compressed'
		};
	}

	// MPEG program stream (PS or MPEG-PS)
	if (check([0x00, 0x00, 0x01, 0xBA])) {
		//  MPEG-PS, MPEG-1 Part 1
		if (check([0x21], {offset: 4, mask: [0xF1]})) {
			return {
				ext: 'mpg', // May also be .ps, .mpeg
				mime: 'video/MP1S'
			};
		}

		// MPEG-PS, MPEG-2 Part 1
		if (check([0x44], {offset: 4, mask: [0xC4]})) {
			return {
				ext: 'mpg', // May also be .mpg, .m2p, .vob or .sub
				mime: 'video/MP2P'
			};
		}
	}

	if (checkString('ITSF')) {
		return {
			ext: 'chm',
			mime: 'application/vnd.ms-htmlhelp'
		};
	}

	// -- 6-byte signatures --

	if (check([0xFD, 0x37, 0x7A, 0x58, 0x5A, 0x00])) {
		return {
			ext: 'xz',
			mime: 'application/x-xz'
		};
	}

	if (checkString('<?xml ')) {
		return {
			ext: 'xml',
			mime: 'application/xml'
		};
	}

	if (check([0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C])) {
		return {
			ext: '7z',
			mime: 'application/x-7z-compressed'
		};
	}

	if (
		check([0x52, 0x61, 0x72, 0x21, 0x1A, 0x7]) &&
		(buffer[6] === 0x0 || buffer[6] === 0x1)
	) {
		return {
			ext: 'rar',
			mime: 'application/x-rar-compressed'
		};
	}

	if (checkString('solid ')) {
		return {
			ext: 'stl',
			mime: 'model/stl'
		};
	}

	// -- 7-byte signatures --

	if (checkString('BLENDER')) {
		return {
			ext: 'blend',
			mime: 'application/x-blender'
		};
	}

	if (checkString('!<arch>')) {
		await tokenizer.ignore(8);
		const str = await tokenizer.readToken(new Token.StringType(13, 'ascii'));
		if (str === 'debian-binary') {
			return {
				ext: 'deb',
				mime: 'application/x-deb'
			};
		}

		return {
			ext: 'ar',
			mime: 'application/x-unix-archive'
		};
	}

	// -- 8-byte signatures --

	if (check([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])) {
		// APNG format (https://wiki.mozilla.org/APNG_Specification)
		// 1. Find the first IDAT (image data) chunk (49 44 41 54)
		// 2. Check if there is an "acTL" chunk before the IDAT one (61 63 54 4C)

		// Offset calculated as follows:
		// - 8 bytes: PNG signature
		// - 4 (length) + 4 (chunk type) + 13 (chunk data) + 4 (CRC): IHDR chunk

		await tokenizer.ignore(8); // ignore PNG signature

		async function readChunkHeader() {
			return {
				length: await tokenizer.readToken(Token.INT32_BE),
				type: await tokenizer.readToken(new Token.StringType(4, 'binary'))
			};
		}

		do {
			const chunk = await readChunkHeader();
			if (chunk.length < 0) {
				return; // Invalid chunk length
			}

			switch (chunk.type) {
				case 'IDAT':
					return {
						ext: 'png',
						mime: 'image/png'
					};
				case 'acTL':
					return {
						ext: 'apng',
						mime: 'image/apng'
					};
				default:
					await tokenizer.ignore(chunk.length + 4); // Ignore chunk-data + CRC
			}
		} while (tokenizer.position + 8 < tokenizer.fileInfo.size);

		return {
			ext: 'png',
			mime: 'image/png'
		};
	}

	if (check([0x41, 0x52, 0x52, 0x4F, 0x57, 0x31, 0x00, 0x00])) {
		return {
			ext: 'arrow',
			mime: 'application/x-apache-arrow'
		};
	}

	if (check([0x67, 0x6C, 0x54, 0x46, 0x02, 0x00, 0x00, 0x00])) {
		return {
			ext: 'glb',
			mime: 'model/gltf-binary'
		};
	}

	// `mov` format variants
	if (
		check([0x66, 0x72, 0x65, 0x65], {offset: 4}) || // `free`
		check([0x6D, 0x64, 0x61, 0x74], {offset: 4}) || // `mdat` MJPEG
		check([0x6D, 0x6F, 0x6F, 0x76], {offset: 4}) || // `moov`
		check([0x77, 0x69, 0x64, 0x65], {offset: 4}) // `wide`
	) {
		return {
			ext: 'mov',
			mime: 'video/quicktime'
		};
	}

	// -- 9-byte signatures --

	if (check([0x49, 0x49, 0x52, 0x4F, 0x08, 0x00, 0x00, 0x00, 0x18])) {
		return {
			ext: 'orf',
			mime: 'image/x-olympus-orf'
		};
	}

	if (checkString('gimp xcf ')) {
		return {
			ext: 'xcf',
			mime: 'image/x-xcf'
		};
	}

	// -- 12-byte signatures --

	if (check([0x49, 0x49, 0x55, 0x00, 0x18, 0x00, 0x00, 0x00, 0x88, 0xE7, 0x74, 0xD8])) {
		return {
			ext: 'rw2',
			mime: 'image/x-panasonic-rw2'
		};
	}

	// ASF_Header_Object first 80 bytes
	if (check([0x30, 0x26, 0xB2, 0x75, 0x8E, 0x66, 0xCF, 0x11, 0xA6, 0xD9])) {
		async function readHeader() {
			const guid = Buffer.alloc(16);
			await tokenizer.readBuffer(guid);
			return {
				id: guid,
				size: Number(await tokenizer.readToken(Token.UINT64_LE))
			};
		}

		await tokenizer.ignore(30);
		// Search for header should be in first 1KB of file.
		while (tokenizer.position + 24 < tokenizer.fileInfo.size) {
			const header = await readHeader();
			let payload = header.size - 24;
			if (_check(header.id, [0x91, 0x07, 0xDC, 0xB7, 0xB7, 0xA9, 0xCF, 0x11, 0x8E, 0xE6, 0x00, 0xC0, 0x0C, 0x20, 0x53, 0x65])) {
				// Sync on Stream-Properties-Object (B7DC0791-A9B7-11CF-8EE6-00C00C205365)
				const typeId = Buffer.alloc(16);
				payload -= await tokenizer.readBuffer(typeId);

				if (_check(typeId, [0x40, 0x9E, 0x69, 0xF8, 0x4D, 0x5B, 0xCF, 0x11, 0xA8, 0xFD, 0x00, 0x80, 0x5F, 0x5C, 0x44, 0x2B])) {
					// Found audio:
					return {
						ext: 'asf',
						mime: 'audio/x-ms-asf'
					};
				}

				if (_check(typeId, [0xC0, 0xEF, 0x19, 0xBC, 0x4D, 0x5B, 0xCF, 0x11, 0xA8, 0xFD, 0x00, 0x80, 0x5F, 0x5C, 0x44, 0x2B])) {
					// Found video:
					return {
						ext: 'asf',
						mime: 'video/x-ms-asf'
					};
				}

				break;
			}

			await tokenizer.ignore(payload);
		}

		// Default to ASF generic extension
		return {
			ext: 'asf',
			mime: 'application/vnd.ms-asf'
		};
	}

	if (check([0xAB, 0x4B, 0x54, 0x58, 0x20, 0x31, 0x31, 0xBB, 0x0D, 0x0A, 0x1A, 0x0A])) {
		return {
			ext: 'ktx',
			mime: 'image/ktx'
		};
	}

	if ((check([0x7E, 0x10, 0x04]) || check([0x7E, 0x18, 0x04])) && check([0x30, 0x4D, 0x49, 0x45], {offset: 4})) {
		return {
			ext: 'mie',
			mime: 'application/x-mie'
		};
	}

	if (check([0x27, 0x0A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00], {offset: 2})) {
		return {
			ext: 'shp',
			mime: 'application/x-esri-shape'
		};
	}

	if (check([0x00, 0x00, 0x00, 0x0C, 0x6A, 0x50, 0x20, 0x20, 0x0D, 0x0A, 0x87, 0x0A])) {
		// JPEG-2000 family

		await tokenizer.ignore(20);
		const type = await tokenizer.readToken(new Token.StringType(4, 'ascii'));
		switch (type) {
			case 'jp2 ':
				return {
					ext: 'jp2',
					mime: 'image/jp2'
				};
			case 'jpx ':
				return {
					ext: 'jpx',
					mime: 'image/jpx'
				};
			case 'jpm ':
				return {
					ext: 'jpm',
					mime: 'image/jpm'
				};
			case 'mjp2':
				return {
					ext: 'mj2',
					mime: 'image/mj2'
				};
			default:
				return;
		}
	}

	if (
		check([0xFF, 0x0A]) ||
		check([0x00, 0x00, 0x00, 0x0C, 0x4A, 0x58, 0x4C, 0x20, 0x0D, 0x0A, 0x87, 0x0A])
	) {
		return {
			ext: 'jxl',
			mime: 'image/jxl'
		};
	}

	// -- Unsafe signatures --

	if (
		check([0x0, 0x0, 0x1, 0xBA]) ||
		check([0x0, 0x0, 0x1, 0xB3])
	) {
		return {
			ext: 'mpg',
			mime: 'video/mpeg'
		};
	}

	if (check([0x00, 0x01, 0x00, 0x00, 0x00])) {
		return {
			ext: 'ttf',
			mime: 'font/ttf'
		};
	}

	if (check([0x00, 0x00, 0x01, 0x00])) {
		return {
			ext: 'ico',
			mime: 'image/x-icon'
		};
	}

	if (check([0x00, 0x00, 0x02, 0x00])) {
		return {
			ext: 'cur',
			mime: 'image/x-icon'
		};
	}

	if (check([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])) {
		// Detected Microsoft Compound File Binary File (MS-CFB) Format.
		return {
			ext: 'cfb',
			mime: 'application/x-cfb'
		};
	}

	// Increase sample size from 12 to 256.
	await tokenizer.peekBuffer(buffer, {length: Math.min(256, tokenizer.fileInfo.size), mayBeLess: true});

	// -- 15-byte signatures --

	if (checkString('BEGIN:')) {
		if (checkString('VCARD', {offset: 6})) {
			return {
				ext: 'vcf',
				mime: 'text/vcard'
			};
		}

		if (checkString('VCALENDAR', {offset: 6})) {
			return {
				ext: 'ics',
				mime: 'text/calendar'
			};
		}
	}

	// `raf` is here just to keep all the raw image detectors together.
	if (checkString('FUJIFILMCCD-RAW')) {
		return {
			ext: 'raf',
			mime: 'image/x-fujifilm-raf'
		};
	}

	if (checkString('Extended Module:')) {
		return {
			ext: 'xm',
			mime: 'audio/x-xm'
		};
	}

	if (checkString('Creative Voice File')) {
		return {
			ext: 'voc',
			mime: 'audio/x-voc'
		};
	}

	if (check([0x04, 0x00, 0x00, 0x00]) && buffer.length >= 16) { // Rough & quick check Pickle/ASAR
		const jsonSize = buffer.readUInt32LE(12);
		if (jsonSize > 12 && buffer.length >= jsonSize + 16) {
			try {
				const header = buffer.slice(16, jsonSize + 16).toString();
				const json = JSON.parse(header);
				// Check if Pickle is ASAR
				if (json.files) { // Final check, assuring Pickle/ASAR format
					return {
						ext: 'asar',
						mime: 'application/x-asar'
					};
				}
			} catch (_) {
			}
		}
	}

	if (check([0x06, 0x0E, 0x2B, 0x34, 0x02, 0x05, 0x01, 0x01, 0x0D, 0x01, 0x02, 0x01, 0x01, 0x02])) {
		return {
			ext: 'mxf',
			mime: 'application/mxf'
		};
	}

	if (checkString('SCRM', {offset: 44})) {
		return {
			ext: 's3m',
			mime: 'audio/x-s3m'
		};
	}

	if (check([0x47], {offset: 4}) && (check([0x47], {offset: 192}) || check([0x47], {offset: 196}))) {
		return {
			ext: 'mts',
			mime: 'video/mp2t'
		};
	}

	if (check([0x42, 0x4F, 0x4F, 0x4B, 0x4D, 0x4F, 0x42, 0x49], {offset: 60})) {
		return {
			ext: 'mobi',
			mime: 'application/x-mobipocket-ebook'
		};
	}

	if (check([0x44, 0x49, 0x43, 0x4D], {offset: 128})) {
		return {
			ext: 'dcm',
			mime: 'application/dicom'
		};
	}

	if (check([0x4C, 0x00, 0x00, 0x00, 0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46])) {
		return {
			ext: 'lnk',
			mime: 'application/x.ms.shortcut' // Invented by us
		};
	}

	if (check([0x62, 0x6F, 0x6F, 0x6B, 0x00, 0x00, 0x00, 0x00, 0x6D, 0x61, 0x72, 0x6B, 0x00, 0x00, 0x00, 0x00])) {
		return {
			ext: 'alias',
			mime: 'application/x.apple.alias' // Invented by us
		};
	}

	if (
		check([0x4C, 0x50], {offset: 34}) &&
		(
			check([0x00, 0x00, 0x01], {offset: 8}) ||
			check([0x01, 0x00, 0x02], {offset: 8}) ||
			check([0x02, 0x00, 0x02], {offset: 8})
		)
	) {
		return {
			ext: 'eot',
			mime: 'application/vnd.ms-fontobject'
		};
	}

	if (check([0x06, 0x06, 0xED, 0xF5, 0xD8, 0x1D, 0x46, 0xE5, 0xBD, 0x31, 0xEF, 0xE7, 0xFE, 0x74, 0xB7, 0x1D])) {
		return {
			ext: 'indd',
			mime: 'application/x-indesign'
		};
	}

	// Increase sample size from 256 to 512
	await tokenizer.peekBuffer(buffer, {length: Math.min(512, tokenizer.fileInfo.size), mayBeLess: true});

	// Requires a buffer size of 512 bytes
	if (tarHeaderChecksumMatches(buffer)) {
		return {
			ext: 'tar',
			mime: 'application/x-tar'
		};
	}

	if (check([0xFF, 0xFE, 0xFF, 0x0E, 0x53, 0x00, 0x6B, 0x00, 0x65, 0x00, 0x74, 0x00, 0x63, 0x00, 0x68, 0x00, 0x55, 0x00, 0x70, 0x00, 0x20, 0x00, 0x4D, 0x00, 0x6F, 0x00, 0x64, 0x00, 0x65, 0x00, 0x6C, 0x00])) {
		return {
			ext: 'skp',
			mime: 'application/vnd.sketchup.skp'
		};
	}

	if (checkString('-----BEGIN PGP MESSAGE-----')) {
		return {
			ext: 'pgp',
			mime: 'application/pgp-encrypted'
		};
	}

	// Check MPEG 1 or 2 Layer 3 header, or 'layer 0' for ADTS (MPEG sync-word 0xFFE)
	if (buffer.length >= 2 && check([0xFF, 0xE0], {offset: 0, mask: [0xFF, 0xE0]})) {
		if (check([0x10], {offset: 1, mask: [0x16]})) {
			// Check for (ADTS) MPEG-2
			if (check([0x08], {offset: 1, mask: [0x08]})) {
				return {
					ext: 'aac',
					mime: 'audio/aac'
				};
			}

			// Must be (ADTS) MPEG-4
			return {
				ext: 'aac',
				mime: 'audio/aac'
			};
		}

		// MPEG 1 or 2 Layer 3 header
		// Check for MPEG layer 3
		if (check([0x02], {offset: 1, mask: [0x06]})) {
			return {
				ext: 'mp3',
				mime: 'audio/mpeg'
			};
		}

		// Check for MPEG layer 2
		if (check([0x04], {offset: 1, mask: [0x06]})) {
			return {
				ext: 'mp2',
				mime: 'audio/mpeg'
			};
		}

		// Check for MPEG layer 1
		if (check([0x06], {offset: 1, mask: [0x06]})) {
			return {
				ext: 'mp1',
				mime: 'audio/mpeg'
			};
		}
	}
}

const stream = readableStream => new Promise((resolve, reject) => {
	// Using `eval` to work around issues when bundling with Webpack
	const stream = eval('require')('stream'); // eslint-disable-line no-eval

	readableStream.on('error', reject);
	readableStream.once('readable', async () => {
		// Set up output stream
		const pass = new stream.PassThrough();
		let outputStream;
		if (stream.pipeline) {
			outputStream = stream.pipeline(readableStream, pass, () => {
			});
		} else {
			outputStream = readableStream.pipe(pass);
		}

		// Read the input stream and detect the filetype
		const chunk = readableStream.read(minimumBytes) || readableStream.read() || Buffer.alloc(0);
		try {
			const fileType = await fromBuffer(chunk);
			pass.fileType = fileType;
		} catch (error) {
			reject(error);
		}

		resolve(outputStream);
	});
});

const fileType = {
	fromStream,
	fromTokenizer,
	fromBuffer,
	stream
};

Object.defineProperty(fileType, 'extensions', {
	get() {
		return new Set(supported.extensions);
	}
});

Object.defineProperty(fileType, 'mimeTypes', {
	get() {
		return new Set(supported.mimeTypes);
	}
});

module.exports = fileType;


/***/ },

/***/ "../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/supported.js"
(module) {

"use strict";


module.exports = {
	extensions: [
		'jpg',
		'png',
		'apng',
		'gif',
		'webp',
		'flif',
		'xcf',
		'cr2',
		'cr3',
		'orf',
		'arw',
		'dng',
		'nef',
		'rw2',
		'raf',
		'tif',
		'bmp',
		'icns',
		'jxr',
		'psd',
		'indd',
		'zip',
		'tar',
		'rar',
		'gz',
		'bz2',
		'7z',
		'dmg',
		'mp4',
		'mid',
		'mkv',
		'webm',
		'mov',
		'avi',
		'mpg',
		'mp2',
		'mp3',
		'm4a',
		'oga',
		'ogg',
		'ogv',
		'opus',
		'flac',
		'wav',
		'spx',
		'amr',
		'pdf',
		'epub',
		'exe',
		'swf',
		'rtf',
		'wasm',
		'woff',
		'woff2',
		'eot',
		'ttf',
		'otf',
		'ico',
		'flv',
		'ps',
		'xz',
		'sqlite',
		'nes',
		'crx',
		'xpi',
		'cab',
		'deb',
		'ar',
		'rpm',
		'Z',
		'lz',
		'cfb',
		'mxf',
		'mts',
		'blend',
		'bpg',
		'docx',
		'pptx',
		'xlsx',
		'3gp',
		'3g2',
		'jp2',
		'jpm',
		'jpx',
		'mj2',
		'aif',
		'qcp',
		'odt',
		'ods',
		'odp',
		'xml',
		'mobi',
		'heic',
		'cur',
		'ktx',
		'ape',
		'wv',
		'dcm',
		'ics',
		'glb',
		'pcap',
		'dsf',
		'lnk',
		'alias',
		'voc',
		'ac3',
		'm4v',
		'm4p',
		'm4b',
		'f4v',
		'f4p',
		'f4b',
		'f4a',
		'mie',
		'asf',
		'ogm',
		'ogx',
		'mpc',
		'arrow',
		'shp',
		'aac',
		'mp1',
		'it',
		's3m',
		'xm',
		'ai',
		'skp',
		'avif',
		'eps',
		'lzh',
		'pgp',
		'asar',
		'stl',
		'chm',
		'3mf',
		'zst',
		'jxl',
		'vcf'
	],
	mimeTypes: [
		'image/jpeg',
		'image/png',
		'image/gif',
		'image/webp',
		'image/flif',
		'image/x-xcf',
		'image/x-canon-cr2',
		'image/x-canon-cr3',
		'image/tiff',
		'image/bmp',
		'image/vnd.ms-photo',
		'image/vnd.adobe.photoshop',
		'application/x-indesign',
		'application/epub+zip',
		'application/x-xpinstall',
		'application/vnd.oasis.opendocument.text',
		'application/vnd.oasis.opendocument.spreadsheet',
		'application/vnd.oasis.opendocument.presentation',
		'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		'application/zip',
		'application/x-tar',
		'application/x-rar-compressed',
		'application/gzip',
		'application/x-bzip2',
		'application/x-7z-compressed',
		'application/x-apple-diskimage',
		'application/x-apache-arrow',
		'video/mp4',
		'audio/midi',
		'video/x-matroska',
		'video/webm',
		'video/quicktime',
		'video/vnd.avi',
		'audio/vnd.wave',
		'audio/qcelp',
		'audio/x-ms-asf',
		'video/x-ms-asf',
		'application/vnd.ms-asf',
		'video/mpeg',
		'video/3gpp',
		'audio/mpeg',
		'audio/mp4', // RFC 4337
		'audio/opus',
		'video/ogg',
		'audio/ogg',
		'application/ogg',
		'audio/x-flac',
		'audio/ape',
		'audio/wavpack',
		'audio/amr',
		'application/pdf',
		'application/x-msdownload',
		'application/x-shockwave-flash',
		'application/rtf',
		'application/wasm',
		'font/woff',
		'font/woff2',
		'application/vnd.ms-fontobject',
		'font/ttf',
		'font/otf',
		'image/x-icon',
		'video/x-flv',
		'application/postscript',
		'application/eps',
		'application/x-xz',
		'application/x-sqlite3',
		'application/x-nintendo-nes-rom',
		'application/x-google-chrome-extension',
		'application/vnd.ms-cab-compressed',
		'application/x-deb',
		'application/x-unix-archive',
		'application/x-rpm',
		'application/x-compress',
		'application/x-lzip',
		'application/x-cfb',
		'application/x-mie',
		'application/mxf',
		'video/mp2t',
		'application/x-blender',
		'image/bpg',
		'image/jp2',
		'image/jpx',
		'image/jpm',
		'image/mj2',
		'audio/aiff',
		'application/xml',
		'application/x-mobipocket-ebook',
		'image/heif',
		'image/heif-sequence',
		'image/heic',
		'image/heic-sequence',
		'image/icns',
		'image/ktx',
		'application/dicom',
		'audio/x-musepack',
		'text/calendar',
		'text/vcard',
		'model/gltf-binary',
		'application/vnd.tcpdump.pcap',
		'audio/x-dsf', // Non-standard
		'application/x.ms.shortcut', // Invented by us
		'application/x.apple.alias', // Invented by us
		'audio/x-voc',
		'audio/vnd.dolby.dd-raw',
		'audio/x-m4a',
		'image/apng',
		'image/x-olympus-orf',
		'image/x-sony-arw',
		'image/x-adobe-dng',
		'image/x-nikon-nef',
		'image/x-panasonic-rw2',
		'image/x-fujifilm-raf',
		'video/x-m4v',
		'video/3gpp2',
		'application/x-esri-shape',
		'audio/aac',
		'audio/x-it',
		'audio/x-s3m',
		'audio/x-xm',
		'video/MP1S',
		'video/MP2P',
		'application/vnd.sketchup.skp',
		'image/avif',
		'application/x-lzh-compressed',
		'application/pgp-encrypted',
		'application/x-asar',
		'model/stl',
		'application/vnd.ms-htmlhelp',
		'model/3mf',
		'image/jxl',
		'application/zstd'
	]
};


/***/ },

/***/ "../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/util.js"
(__unused_webpack_module, exports) {

"use strict";


exports.stringToBytes = string => [...string].map(character => character.charCodeAt(0));

/**
Checks whether the TAR checksum is valid.

@param {Buffer} buffer - The TAR header `[offset ... offset + 512]`.
@param {number} offset - TAR header offset.
@returns {boolean} `true` if the TAR checksum is valid, otherwise `false`.
*/
exports.tarHeaderChecksumMatches = (buffer, offset = 0) => {
	const readSum = parseInt(buffer.toString('utf8', 148, 154).replace(/\0.*$/, '').trim(), 8); // Read sum in header
	if (isNaN(readSum)) {
		return false;
	}

	let sum = 8 * 0x20; // Initialize signed bit sum

	for (let i = offset; i < offset + 148; i++) {
		sum += buffer[i];
	}

	for (let i = offset + 156; i < offset + 512; i++) {
		sum += buffer[i];
	}

	return readSum === sum;
};

/**
ID3 UINT32 sync-safe tokenizer token.
28 bits (representing up to 256MB) integer, the msb is 0 to avoid "false syncsignals".
*/
exports.uint32SyncSafeToken = {
	get: (buffer, offset) => {
		return (buffer[offset + 3] & 0x7F) | ((buffer[offset + 2]) << 7) | ((buffer[offset + 1]) << 14) | ((buffer[offset]) << 21);
	},
	len: 4
};


/***/ },

/***/ "../../node_modules/.pnpm/ieee754@1.2.1/node_modules/ieee754/index.js"
(__unused_webpack_module, exports) {

/*! ieee754. BSD-3-Clause License. Feross Aboukhadijeh <https://feross.org/opensource> */
exports.read = function (buffer, offset, isLE, mLen, nBytes) {
  var e, m
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var nBits = -7
  var i = isLE ? (nBytes - 1) : 0
  var d = isLE ? -1 : 1
  var s = buffer[offset + i]

  i += d

  e = s & ((1 << (-nBits)) - 1)
  s >>= (-nBits)
  nBits += eLen
  for (; nBits > 0; e = (e * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  m = e & ((1 << (-nBits)) - 1)
  e >>= (-nBits)
  nBits += mLen
  for (; nBits > 0; m = (m * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  if (e === 0) {
    e = 1 - eBias
  } else if (e === eMax) {
    return m ? NaN : ((s ? -1 : 1) * Infinity)
  } else {
    m = m + Math.pow(2, mLen)
    e = e - eBias
  }
  return (s ? -1 : 1) * m * Math.pow(2, e - mLen)
}

exports.write = function (buffer, value, offset, isLE, mLen, nBytes) {
  var e, m, c
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var rt = (mLen === 23 ? Math.pow(2, -24) - Math.pow(2, -77) : 0)
  var i = isLE ? 0 : (nBytes - 1)
  var d = isLE ? 1 : -1
  var s = value < 0 || (value === 0 && 1 / value < 0) ? 1 : 0

  value = Math.abs(value)

  if (isNaN(value) || value === Infinity) {
    m = isNaN(value) ? 1 : 0
    e = eMax
  } else {
    e = Math.floor(Math.log(value) / Math.LN2)
    if (value * (c = Math.pow(2, -e)) < 1) {
      e--
      c *= 2
    }
    if (e + eBias >= 1) {
      value += rt / c
    } else {
      value += rt * Math.pow(2, 1 - eBias)
    }
    if (value * c >= 2) {
      e++
      c /= 2
    }

    if (e + eBias >= eMax) {
      m = 0
      e = eMax
    } else if (e + eBias >= 1) {
      m = ((value * c) - 1) * Math.pow(2, mLen)
      e = e + eBias
    } else {
      m = value * Math.pow(2, eBias - 1) * Math.pow(2, mLen)
      e = 0
    }
  }

  for (; mLen >= 8; buffer[offset + i] = m & 0xff, i += d, m /= 256, mLen -= 8) {}

  e = (e << mLen) | m
  eLen += mLen
  for (; eLen > 0; buffer[offset + i] = e & 0xff, i += d, e /= 256, eLen -= 8) {}

  buffer[offset + i - d] |= s * 128
}


/***/ },

/***/ "../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/Mime.js"
(module) {

"use strict";


/**
 * @param typeMap [Object] Map of MIME type -> Array[extensions]
 * @param ...
 */
function Mime() {
  this._types = Object.create(null);
  this._extensions = Object.create(null);

  for (let i = 0; i < arguments.length; i++) {
    this.define(arguments[i]);
  }

  this.define = this.define.bind(this);
  this.getType = this.getType.bind(this);
  this.getExtension = this.getExtension.bind(this);
}

/**
 * Define mimetype -> extension mappings.  Each key is a mime-type that maps
 * to an array of extensions associated with the type.  The first extension is
 * used as the default extension for the type.
 *
 * e.g. mime.define({'audio/ogg', ['oga', 'ogg', 'spx']});
 *
 * If a type declares an extension that has already been defined, an error will
 * be thrown.  To suppress this error and force the extension to be associated
 * with the new type, pass `force`=true.  Alternatively, you may prefix the
 * extension with "*" to map the type to extension, without mapping the
 * extension to the type.
 *
 * e.g. mime.define({'audio/wav', ['wav']}, {'audio/x-wav', ['*wav']});
 *
 *
 * @param map (Object) type definitions
 * @param force (Boolean) if true, force overriding of existing definitions
 */
Mime.prototype.define = function(typeMap, force) {
  for (let type in typeMap) {
    let extensions = typeMap[type].map(function(t) {
      return t.toLowerCase();
    });
    type = type.toLowerCase();

    for (let i = 0; i < extensions.length; i++) {
      const ext = extensions[i];

      // '*' prefix = not the preferred type for this extension.  So fixup the
      // extension, and skip it.
      if (ext[0] === '*') {
        continue;
      }

      if (!force && (ext in this._types)) {
        throw new Error(
          'Attempt to change mapping for "' + ext +
          '" extension from "' + this._types[ext] + '" to "' + type +
          '". Pass `force=true` to allow this, otherwise remove "' + ext +
          '" from the list of extensions for "' + type + '".'
        );
      }

      this._types[ext] = type;
    }

    // Use first extension as default
    if (force || !this._extensions[type]) {
      const ext = extensions[0];
      this._extensions[type] = (ext[0] !== '*') ? ext : ext.substr(1);
    }
  }
};

/**
 * Lookup a mime type based on extension
 */
Mime.prototype.getType = function(path) {
  path = String(path);
  let last = path.replace(/^.*[/\\]/, '').toLowerCase();
  let ext = last.replace(/^.*\./, '').toLowerCase();

  let hasPath = last.length < path.length;
  let hasDot = ext.length < last.length - 1;

  return (hasDot || !hasPath) && this._types[ext] || null;
};

/**
 * Return file extension associated with a mime type
 */
Mime.prototype.getExtension = function(type) {
  type = /^\s*([^;\s]*)/.test(type) && RegExp.$1;
  return type && this._extensions[type.toLowerCase()] || null;
};

module.exports = Mime;


/***/ },

/***/ "../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/lite.js"
(module, __unused_webpack_exports, __webpack_require__) {

"use strict";


let Mime = __webpack_require__("../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/Mime.js");
module.exports = new Mime(__webpack_require__("../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/types/standard.js"));


/***/ },

/***/ "../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/types/standard.js"
(module) {

module.exports = {"application/andrew-inset":["ez"],"application/applixware":["aw"],"application/atom+xml":["atom"],"application/atomcat+xml":["atomcat"],"application/atomdeleted+xml":["atomdeleted"],"application/atomsvc+xml":["atomsvc"],"application/atsc-dwd+xml":["dwd"],"application/atsc-held+xml":["held"],"application/atsc-rsat+xml":["rsat"],"application/bdoc":["bdoc"],"application/calendar+xml":["xcs"],"application/ccxml+xml":["ccxml"],"application/cdfx+xml":["cdfx"],"application/cdmi-capability":["cdmia"],"application/cdmi-container":["cdmic"],"application/cdmi-domain":["cdmid"],"application/cdmi-object":["cdmio"],"application/cdmi-queue":["cdmiq"],"application/cu-seeme":["cu"],"application/dash+xml":["mpd"],"application/davmount+xml":["davmount"],"application/docbook+xml":["dbk"],"application/dssc+der":["dssc"],"application/dssc+xml":["xdssc"],"application/ecmascript":["es","ecma"],"application/emma+xml":["emma"],"application/emotionml+xml":["emotionml"],"application/epub+zip":["epub"],"application/exi":["exi"],"application/express":["exp"],"application/fdt+xml":["fdt"],"application/font-tdpfr":["pfr"],"application/geo+json":["geojson"],"application/gml+xml":["gml"],"application/gpx+xml":["gpx"],"application/gxf":["gxf"],"application/gzip":["gz"],"application/hjson":["hjson"],"application/hyperstudio":["stk"],"application/inkml+xml":["ink","inkml"],"application/ipfix":["ipfix"],"application/its+xml":["its"],"application/java-archive":["jar","war","ear"],"application/java-serialized-object":["ser"],"application/java-vm":["class"],"application/javascript":["js","mjs"],"application/json":["json","map"],"application/json5":["json5"],"application/jsonml+json":["jsonml"],"application/ld+json":["jsonld"],"application/lgr+xml":["lgr"],"application/lost+xml":["lostxml"],"application/mac-binhex40":["hqx"],"application/mac-compactpro":["cpt"],"application/mads+xml":["mads"],"application/manifest+json":["webmanifest"],"application/marc":["mrc"],"application/marcxml+xml":["mrcx"],"application/mathematica":["ma","nb","mb"],"application/mathml+xml":["mathml"],"application/mbox":["mbox"],"application/mediaservercontrol+xml":["mscml"],"application/metalink+xml":["metalink"],"application/metalink4+xml":["meta4"],"application/mets+xml":["mets"],"application/mmt-aei+xml":["maei"],"application/mmt-usd+xml":["musd"],"application/mods+xml":["mods"],"application/mp21":["m21","mp21"],"application/mp4":["mp4s","m4p"],"application/msword":["doc","dot"],"application/mxf":["mxf"],"application/n-quads":["nq"],"application/n-triples":["nt"],"application/node":["cjs"],"application/octet-stream":["bin","dms","lrf","mar","so","dist","distz","pkg","bpk","dump","elc","deploy","exe","dll","deb","dmg","iso","img","msi","msp","msm","buffer"],"application/oda":["oda"],"application/oebps-package+xml":["opf"],"application/ogg":["ogx"],"application/omdoc+xml":["omdoc"],"application/onenote":["onetoc","onetoc2","onetmp","onepkg"],"application/oxps":["oxps"],"application/p2p-overlay+xml":["relo"],"application/patch-ops-error+xml":["xer"],"application/pdf":["pdf"],"application/pgp-encrypted":["pgp"],"application/pgp-signature":["asc","sig"],"application/pics-rules":["prf"],"application/pkcs10":["p10"],"application/pkcs7-mime":["p7m","p7c"],"application/pkcs7-signature":["p7s"],"application/pkcs8":["p8"],"application/pkix-attr-cert":["ac"],"application/pkix-cert":["cer"],"application/pkix-crl":["crl"],"application/pkix-pkipath":["pkipath"],"application/pkixcmp":["pki"],"application/pls+xml":["pls"],"application/postscript":["ai","eps","ps"],"application/provenance+xml":["provx"],"application/pskc+xml":["pskcxml"],"application/raml+yaml":["raml"],"application/rdf+xml":["rdf","owl"],"application/reginfo+xml":["rif"],"application/relax-ng-compact-syntax":["rnc"],"application/resource-lists+xml":["rl"],"application/resource-lists-diff+xml":["rld"],"application/rls-services+xml":["rs"],"application/route-apd+xml":["rapd"],"application/route-s-tsid+xml":["sls"],"application/route-usd+xml":["rusd"],"application/rpki-ghostbusters":["gbr"],"application/rpki-manifest":["mft"],"application/rpki-roa":["roa"],"application/rsd+xml":["rsd"],"application/rss+xml":["rss"],"application/rtf":["rtf"],"application/sbml+xml":["sbml"],"application/scvp-cv-request":["scq"],"application/scvp-cv-response":["scs"],"application/scvp-vp-request":["spq"],"application/scvp-vp-response":["spp"],"application/sdp":["sdp"],"application/senml+xml":["senmlx"],"application/sensml+xml":["sensmlx"],"application/set-payment-initiation":["setpay"],"application/set-registration-initiation":["setreg"],"application/shf+xml":["shf"],"application/sieve":["siv","sieve"],"application/smil+xml":["smi","smil"],"application/sparql-query":["rq"],"application/sparql-results+xml":["srx"],"application/srgs":["gram"],"application/srgs+xml":["grxml"],"application/sru+xml":["sru"],"application/ssdl+xml":["ssdl"],"application/ssml+xml":["ssml"],"application/swid+xml":["swidtag"],"application/tei+xml":["tei","teicorpus"],"application/thraud+xml":["tfi"],"application/timestamped-data":["tsd"],"application/toml":["toml"],"application/trig":["trig"],"application/ttml+xml":["ttml"],"application/ubjson":["ubj"],"application/urc-ressheet+xml":["rsheet"],"application/urc-targetdesc+xml":["td"],"application/voicexml+xml":["vxml"],"application/wasm":["wasm"],"application/widget":["wgt"],"application/winhlp":["hlp"],"application/wsdl+xml":["wsdl"],"application/wspolicy+xml":["wspolicy"],"application/xaml+xml":["xaml"],"application/xcap-att+xml":["xav"],"application/xcap-caps+xml":["xca"],"application/xcap-diff+xml":["xdf"],"application/xcap-el+xml":["xel"],"application/xcap-ns+xml":["xns"],"application/xenc+xml":["xenc"],"application/xhtml+xml":["xhtml","xht"],"application/xliff+xml":["xlf"],"application/xml":["xml","xsl","xsd","rng"],"application/xml-dtd":["dtd"],"application/xop+xml":["xop"],"application/xproc+xml":["xpl"],"application/xslt+xml":["*xsl","xslt"],"application/xspf+xml":["xspf"],"application/xv+xml":["mxml","xhvml","xvml","xvm"],"application/yang":["yang"],"application/yin+xml":["yin"],"application/zip":["zip"],"audio/3gpp":["*3gpp"],"audio/adpcm":["adp"],"audio/amr":["amr"],"audio/basic":["au","snd"],"audio/midi":["mid","midi","kar","rmi"],"audio/mobile-xmf":["mxmf"],"audio/mp3":["*mp3"],"audio/mp4":["m4a","mp4a"],"audio/mpeg":["mpga","mp2","mp2a","mp3","m2a","m3a"],"audio/ogg":["oga","ogg","spx","opus"],"audio/s3m":["s3m"],"audio/silk":["sil"],"audio/wav":["wav"],"audio/wave":["*wav"],"audio/webm":["weba"],"audio/xm":["xm"],"font/collection":["ttc"],"font/otf":["otf"],"font/ttf":["ttf"],"font/woff":["woff"],"font/woff2":["woff2"],"image/aces":["exr"],"image/apng":["apng"],"image/avif":["avif"],"image/bmp":["bmp"],"image/cgm":["cgm"],"image/dicom-rle":["drle"],"image/emf":["emf"],"image/fits":["fits"],"image/g3fax":["g3"],"image/gif":["gif"],"image/heic":["heic"],"image/heic-sequence":["heics"],"image/heif":["heif"],"image/heif-sequence":["heifs"],"image/hej2k":["hej2"],"image/hsj2":["hsj2"],"image/ief":["ief"],"image/jls":["jls"],"image/jp2":["jp2","jpg2"],"image/jpeg":["jpeg","jpg","jpe"],"image/jph":["jph"],"image/jphc":["jhc"],"image/jpm":["jpm"],"image/jpx":["jpx","jpf"],"image/jxr":["jxr"],"image/jxra":["jxra"],"image/jxrs":["jxrs"],"image/jxs":["jxs"],"image/jxsc":["jxsc"],"image/jxsi":["jxsi"],"image/jxss":["jxss"],"image/ktx":["ktx"],"image/ktx2":["ktx2"],"image/png":["png"],"image/sgi":["sgi"],"image/svg+xml":["svg","svgz"],"image/t38":["t38"],"image/tiff":["tif","tiff"],"image/tiff-fx":["tfx"],"image/webp":["webp"],"image/wmf":["wmf"],"message/disposition-notification":["disposition-notification"],"message/global":["u8msg"],"message/global-delivery-status":["u8dsn"],"message/global-disposition-notification":["u8mdn"],"message/global-headers":["u8hdr"],"message/rfc822":["eml","mime"],"model/3mf":["3mf"],"model/gltf+json":["gltf"],"model/gltf-binary":["glb"],"model/iges":["igs","iges"],"model/mesh":["msh","mesh","silo"],"model/mtl":["mtl"],"model/obj":["obj"],"model/step+xml":["stpx"],"model/step+zip":["stpz"],"model/step-xml+zip":["stpxz"],"model/stl":["stl"],"model/vrml":["wrl","vrml"],"model/x3d+binary":["*x3db","x3dbz"],"model/x3d+fastinfoset":["x3db"],"model/x3d+vrml":["*x3dv","x3dvz"],"model/x3d+xml":["x3d","x3dz"],"model/x3d-vrml":["x3dv"],"text/cache-manifest":["appcache","manifest"],"text/calendar":["ics","ifb"],"text/coffeescript":["coffee","litcoffee"],"text/css":["css"],"text/csv":["csv"],"text/html":["html","htm","shtml"],"text/jade":["jade"],"text/jsx":["jsx"],"text/less":["less"],"text/markdown":["markdown","md"],"text/mathml":["mml"],"text/mdx":["mdx"],"text/n3":["n3"],"text/plain":["txt","text","conf","def","list","log","in","ini"],"text/richtext":["rtx"],"text/rtf":["*rtf"],"text/sgml":["sgml","sgm"],"text/shex":["shex"],"text/slim":["slim","slm"],"text/spdx":["spdx"],"text/stylus":["stylus","styl"],"text/tab-separated-values":["tsv"],"text/troff":["t","tr","roff","man","me","ms"],"text/turtle":["ttl"],"text/uri-list":["uri","uris","urls"],"text/vcard":["vcard"],"text/vtt":["vtt"],"text/xml":["*xml"],"text/yaml":["yaml","yml"],"video/3gpp":["3gp","3gpp"],"video/3gpp2":["3g2"],"video/h261":["h261"],"video/h263":["h263"],"video/h264":["h264"],"video/iso.segment":["m4s"],"video/jpeg":["jpgv"],"video/jpm":["*jpm","jpgm"],"video/mj2":["mj2","mjp2"],"video/mp2t":["ts"],"video/mp4":["mp4","mp4v","mpg4"],"video/mpeg":["mpeg","mpg","mpe","m1v","m2v"],"video/ogg":["ogv"],"video/quicktime":["qt","mov"],"video/webm":["webm"]};

/***/ },

/***/ "../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/Deferred.js"
(__unused_webpack_module, exports) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.Deferred = void 0;
class Deferred {
    constructor() {
        this.resolve = () => null;
        this.reject = () => null;
        this.promise = new Promise((resolve, reject) => {
            this.reject = reject;
            this.resolve = resolve;
        });
    }
}
exports.Deferred = Deferred;


/***/ },

/***/ "../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/EndOfFileStream.js"
(__unused_webpack_module, exports) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.EndOfStreamError = exports.defaultMessages = void 0;
exports.defaultMessages = 'End-Of-Stream';
/**
 * Thrown on read operation of the end of file or stream has been reached
 */
class EndOfStreamError extends Error {
    constructor() {
        super(exports.defaultMessages);
    }
}
exports.EndOfStreamError = EndOfStreamError;


/***/ },

/***/ "../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/StreamReader.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.StreamReader = exports.EndOfStreamError = void 0;
const EndOfFileStream_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/EndOfFileStream.js");
const Deferred_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/Deferred.js");
var EndOfFileStream_2 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/EndOfFileStream.js");
Object.defineProperty(exports, "EndOfStreamError", ({ enumerable: true, get: function () { return EndOfFileStream_2.EndOfStreamError; } }));
const maxStreamReadSize = 1 * 1024 * 1024; // Maximum request length on read-stream operation
class StreamReader {
    constructor(s) {
        this.s = s;
        /**
         * Deferred used for postponed read request (as not data is yet available to read)
         */
        this.deferred = null;
        this.endOfStream = false;
        /**
         * Store peeked data
         * @type {Array}
         */
        this.peekQueue = [];
        if (!s.read || !s.once) {
            throw new Error('Expected an instance of stream.Readable');
        }
        this.s.once('end', () => this.reject(new EndOfFileStream_1.EndOfStreamError()));
        this.s.once('error', err => this.reject(err));
        this.s.once('close', () => this.reject(new Error('Stream closed')));
    }
    /**
     * Read ahead (peek) from stream. Subsequent read or peeks will return the same data
     * @param uint8Array - Uint8Array (or Buffer) to store data read from stream in
     * @param offset - Offset target
     * @param length - Number of bytes to read
     * @returns Number of bytes peeked
     */
    async peek(uint8Array, offset, length) {
        const bytesRead = await this.read(uint8Array, offset, length);
        this.peekQueue.push(uint8Array.subarray(offset, offset + bytesRead)); // Put read data back to peek buffer
        return bytesRead;
    }
    /**
     * Read chunk from stream
     * @param buffer - Target Uint8Array (or Buffer) to store data read from stream in
     * @param offset - Offset target
     * @param length - Number of bytes to read
     * @returns Number of bytes read
     */
    async read(buffer, offset, length) {
        if (length === 0) {
            return 0;
        }
        if (this.peekQueue.length === 0 && this.endOfStream) {
            throw new EndOfFileStream_1.EndOfStreamError();
        }
        let remaining = length;
        let bytesRead = 0;
        // consume peeked data first
        while (this.peekQueue.length > 0 && remaining > 0) {
            const peekData = this.peekQueue.pop(); // Front of queue
            if (!peekData)
                throw new Error('peekData should be defined');
            const lenCopy = Math.min(peekData.length, remaining);
            buffer.set(peekData.subarray(0, lenCopy), offset + bytesRead);
            bytesRead += lenCopy;
            remaining -= lenCopy;
            if (lenCopy < peekData.length) {
                // remainder back to queue
                this.peekQueue.push(peekData.subarray(lenCopy));
            }
        }
        // continue reading from stream if required
        while (remaining > 0 && !this.endOfStream) {
            const reqLen = Math.min(remaining, maxStreamReadSize);
            const chunkLen = await this.readFromStream(buffer, offset + bytesRead, reqLen);
            bytesRead += chunkLen;
            if (chunkLen < reqLen)
                break;
            remaining -= chunkLen;
        }
        return bytesRead;
    }
    /**
     * Read chunk from stream
     * @param buffer Target Uint8Array (or Buffer) to store data read from stream in
     * @param offset Offset target
     * @param length Number of bytes to read
     * @returns Number of bytes read
     */
    async readFromStream(buffer, offset, length) {
        const readBuffer = this.s.read(length);
        if (readBuffer) {
            buffer.set(readBuffer, offset);
            return readBuffer.length;
        }
        else {
            const request = {
                buffer,
                offset,
                length,
                deferred: new Deferred_1.Deferred()
            };
            this.deferred = request.deferred;
            this.s.once('readable', () => {
                this.readDeferred(request);
            });
            return request.deferred.promise;
        }
    }
    /**
     * Process deferred read request
     * @param request Deferred read request
     */
    readDeferred(request) {
        const readBuffer = this.s.read(request.length);
        if (readBuffer) {
            request.buffer.set(readBuffer, request.offset);
            request.deferred.resolve(readBuffer.length);
            this.deferred = null;
        }
        else {
            this.s.once('readable', () => {
                this.readDeferred(request);
            });
        }
    }
    reject(err) {
        this.endOfStream = true;
        if (this.deferred) {
            this.deferred.reject(err);
            this.deferred = null;
        }
    }
}
exports.StreamReader = StreamReader;


/***/ },

/***/ "../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/index.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.StreamReader = exports.EndOfStreamError = void 0;
var EndOfFileStream_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/EndOfFileStream.js");
Object.defineProperty(exports, "EndOfStreamError", ({ enumerable: true, get: function () { return EndOfFileStream_1.EndOfStreamError; } }));
var StreamReader_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/StreamReader.js");
Object.defineProperty(exports, "StreamReader", ({ enumerable: true, get: function () { return StreamReader_1.StreamReader; } }));


/***/ },

/***/ "../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/AbstractTokenizer.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.AbstractTokenizer = void 0;
const peek_readable_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/index.js");
/**
 * Core tokenizer
 */
class AbstractTokenizer {
    constructor(fileInfo) {
        /**
         * Tokenizer-stream position
         */
        this.position = 0;
        this.numBuffer = new Uint8Array(8);
        this.fileInfo = fileInfo ? fileInfo : {};
    }
    /**
     * Read a token from the tokenizer-stream
     * @param token - The token to read
     * @param position - If provided, the desired position in the tokenizer-stream
     * @returns Promise with token data
     */
    async readToken(token, position = this.position) {
        const uint8Array = Buffer.alloc(token.len);
        const len = await this.readBuffer(uint8Array, { position });
        if (len < token.len)
            throw new peek_readable_1.EndOfStreamError();
        return token.get(uint8Array, 0);
    }
    /**
     * Peek a token from the tokenizer-stream.
     * @param token - Token to peek from the tokenizer-stream.
     * @param position - Offset where to begin reading within the file. If position is null, data will be read from the current file position.
     * @returns Promise with token data
     */
    async peekToken(token, position = this.position) {
        const uint8Array = Buffer.alloc(token.len);
        const len = await this.peekBuffer(uint8Array, { position });
        if (len < token.len)
            throw new peek_readable_1.EndOfStreamError();
        return token.get(uint8Array, 0);
    }
    /**
     * Read a numeric token from the stream
     * @param token - Numeric token
     * @returns Promise with number
     */
    async readNumber(token) {
        const len = await this.readBuffer(this.numBuffer, { length: token.len });
        if (len < token.len)
            throw new peek_readable_1.EndOfStreamError();
        return token.get(this.numBuffer, 0);
    }
    /**
     * Read a numeric token from the stream
     * @param token - Numeric token
     * @returns Promise with number
     */
    async peekNumber(token) {
        const len = await this.peekBuffer(this.numBuffer, { length: token.len });
        if (len < token.len)
            throw new peek_readable_1.EndOfStreamError();
        return token.get(this.numBuffer, 0);
    }
    /**
     * Ignore number of bytes, advances the pointer in under tokenizer-stream.
     * @param length - Number of bytes to ignore
     * @return resolves the number of bytes ignored, equals length if this available, otherwise the number of bytes available
     */
    async ignore(length) {
        if (this.fileInfo.size !== undefined) {
            const bytesLeft = this.fileInfo.size - this.position;
            if (length > bytesLeft) {
                this.position += bytesLeft;
                return bytesLeft;
            }
        }
        this.position += length;
        return length;
    }
    async close() {
        // empty
    }
    normalizeOptions(uint8Array, options) {
        if (options && options.position !== undefined && options.position < this.position) {
            throw new Error('`options.position` must be equal or greater than `tokenizer.position`');
        }
        if (options) {
            return {
                mayBeLess: options.mayBeLess === true,
                offset: options.offset ? options.offset : 0,
                length: options.length ? options.length : (uint8Array.length - (options.offset ? options.offset : 0)),
                position: options.position ? options.position : this.position
            };
        }
        return {
            mayBeLess: false,
            offset: 0,
            length: uint8Array.length,
            position: this.position
        };
    }
}
exports.AbstractTokenizer = AbstractTokenizer;


/***/ },

/***/ "../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/BufferTokenizer.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.BufferTokenizer = void 0;
const peek_readable_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/index.js");
const AbstractTokenizer_1 = __webpack_require__("../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/AbstractTokenizer.js");
class BufferTokenizer extends AbstractTokenizer_1.AbstractTokenizer {
    /**
     * Construct BufferTokenizer
     * @param uint8Array - Uint8Array to tokenize
     * @param fileInfo - Pass additional file information to the tokenizer
     */
    constructor(uint8Array, fileInfo) {
        super(fileInfo);
        this.uint8Array = uint8Array;
        this.fileInfo.size = this.fileInfo.size ? this.fileInfo.size : uint8Array.length;
    }
    /**
     * Read buffer from tokenizer
     * @param uint8Array - Uint8Array to tokenize
     * @param options - Read behaviour options
     * @returns {Promise<number>}
     */
    async readBuffer(uint8Array, options) {
        if (options && options.position) {
            if (options.position < this.position) {
                throw new Error('`options.position` must be equal or greater than `tokenizer.position`');
            }
            this.position = options.position;
        }
        const bytesRead = await this.peekBuffer(uint8Array, options);
        this.position += bytesRead;
        return bytesRead;
    }
    /**
     * Peek (read ahead) buffer from tokenizer
     * @param uint8Array
     * @param options - Read behaviour options
     * @returns {Promise<number>}
     */
    async peekBuffer(uint8Array, options) {
        const normOptions = this.normalizeOptions(uint8Array, options);
        const bytes2read = Math.min(this.uint8Array.length - normOptions.position, normOptions.length);
        if ((!normOptions.mayBeLess) && bytes2read < normOptions.length) {
            throw new peek_readable_1.EndOfStreamError();
        }
        else {
            uint8Array.set(this.uint8Array.subarray(normOptions.position, normOptions.position + bytes2read), normOptions.offset);
            return bytes2read;
        }
    }
    async close() {
        // empty
    }
}
exports.BufferTokenizer = BufferTokenizer;


/***/ },

/***/ "../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/ReadStreamTokenizer.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.ReadStreamTokenizer = void 0;
const AbstractTokenizer_1 = __webpack_require__("../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/AbstractTokenizer.js");
const peek_readable_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/index.js");
const maxBufferSize = 256000;
class ReadStreamTokenizer extends AbstractTokenizer_1.AbstractTokenizer {
    constructor(stream, fileInfo) {
        super(fileInfo);
        this.streamReader = new peek_readable_1.StreamReader(stream);
    }
    /**
     * Get file information, an HTTP-client may implement this doing a HEAD request
     * @return Promise with file information
     */
    async getFileInfo() {
        return this.fileInfo;
    }
    /**
     * Read buffer from tokenizer
     * @param uint8Array - Target Uint8Array to fill with data read from the tokenizer-stream
     * @param options - Read behaviour options
     * @returns Promise with number of bytes read
     */
    async readBuffer(uint8Array, options) {
        const normOptions = this.normalizeOptions(uint8Array, options);
        const skipBytes = normOptions.position - this.position;
        if (skipBytes > 0) {
            await this.ignore(skipBytes);
            return this.readBuffer(uint8Array, options);
        }
        else if (skipBytes < 0) {
            throw new Error('`options.position` must be equal or greater than `tokenizer.position`');
        }
        if (normOptions.length === 0) {
            return 0;
        }
        const bytesRead = await this.streamReader.read(uint8Array, normOptions.offset, normOptions.length);
        this.position += bytesRead;
        if ((!options || !options.mayBeLess) && bytesRead < normOptions.length) {
            throw new peek_readable_1.EndOfStreamError();
        }
        return bytesRead;
    }
    /**
     * Peek (read ahead) buffer from tokenizer
     * @param uint8Array - Uint8Array (or Buffer) to write data to
     * @param options - Read behaviour options
     * @returns Promise with number of bytes peeked
     */
    async peekBuffer(uint8Array, options) {
        const normOptions = this.normalizeOptions(uint8Array, options);
        let bytesRead = 0;
        if (normOptions.position) {
            const skipBytes = normOptions.position - this.position;
            if (skipBytes > 0) {
                const skipBuffer = new Uint8Array(normOptions.length + skipBytes);
                bytesRead = await this.peekBuffer(skipBuffer, { mayBeLess: normOptions.mayBeLess });
                uint8Array.set(skipBuffer.subarray(skipBytes), normOptions.offset);
                return bytesRead - skipBytes;
            }
            else if (skipBytes < 0) {
                throw new Error('Cannot peek from a negative offset in a stream');
            }
        }
        if (normOptions.length > 0) {
            try {
                bytesRead = await this.streamReader.peek(uint8Array, normOptions.offset, normOptions.length);
            }
            catch (err) {
                if (options && options.mayBeLess && err instanceof peek_readable_1.EndOfStreamError) {
                    return 0;
                }
                throw err;
            }
            if ((!normOptions.mayBeLess) && bytesRead < normOptions.length) {
                throw new peek_readable_1.EndOfStreamError();
            }
        }
        return bytesRead;
    }
    async ignore(length) {
        // debug(`ignore ${this.position}...${this.position + length - 1}`);
        const bufSize = Math.min(maxBufferSize, length);
        const buf = new Uint8Array(bufSize);
        let totBytesRead = 0;
        while (totBytesRead < length) {
            const remaining = length - totBytesRead;
            const bytesRead = await this.readBuffer(buf, { length: Math.min(bufSize, remaining) });
            if (bytesRead < 0) {
                return bytesRead;
            }
            totBytesRead += bytesRead;
        }
        return totBytesRead;
    }
}
exports.ReadStreamTokenizer = ReadStreamTokenizer;


/***/ },

/***/ "../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/core.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.fromBuffer = exports.fromStream = exports.EndOfStreamError = void 0;
const ReadStreamTokenizer_1 = __webpack_require__("../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/ReadStreamTokenizer.js");
const BufferTokenizer_1 = __webpack_require__("../../node_modules/.pnpm/strtok3@6.3.0/node_modules/strtok3/lib/BufferTokenizer.js");
var peek_readable_1 = __webpack_require__("../../node_modules/.pnpm/peek-readable@4.1.0/node_modules/peek-readable/lib/index.js");
Object.defineProperty(exports, "EndOfStreamError", ({ enumerable: true, get: function () { return peek_readable_1.EndOfStreamError; } }));
/**
 * Construct ReadStreamTokenizer from given Stream.
 * Will set fileSize, if provided given Stream has set the .path property/
 * @param stream - Read from Node.js Stream.Readable
 * @param fileInfo - Pass the file information, like size and MIME-type of the corresponding stream.
 * @returns ReadStreamTokenizer
 */
function fromStream(stream, fileInfo) {
    fileInfo = fileInfo ? fileInfo : {};
    return new ReadStreamTokenizer_1.ReadStreamTokenizer(stream, fileInfo);
}
exports.fromStream = fromStream;
/**
 * Construct ReadStreamTokenizer from given Buffer.
 * @param uint8Array - Uint8Array to tokenize
 * @param fileInfo - Pass additional file information to the tokenizer
 * @returns BufferTokenizer
 */
function fromBuffer(uint8Array, fileInfo) {
    return new BufferTokenizer_1.BufferTokenizer(uint8Array, fileInfo);
}
exports.fromBuffer = fromBuffer;


/***/ },

/***/ "../../node_modules/.pnpm/token-types@4.2.1/node_modules/token-types/lib/index.js"
(__unused_webpack_module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", ({ value: true }));
exports.AnsiStringType = exports.StringType = exports.BufferType = exports.Uint8ArrayType = exports.IgnoreType = exports.Float80_LE = exports.Float80_BE = exports.Float64_LE = exports.Float64_BE = exports.Float32_LE = exports.Float32_BE = exports.Float16_LE = exports.Float16_BE = exports.INT64_BE = exports.UINT64_BE = exports.INT64_LE = exports.UINT64_LE = exports.INT32_LE = exports.INT32_BE = exports.INT24_BE = exports.INT24_LE = exports.INT16_LE = exports.INT16_BE = exports.INT8 = exports.UINT32_BE = exports.UINT32_LE = exports.UINT24_BE = exports.UINT24_LE = exports.UINT16_BE = exports.UINT16_LE = exports.UINT8 = void 0;
const ieee754 = __webpack_require__("../../node_modules/.pnpm/ieee754@1.2.1/node_modules/ieee754/index.js");
// Primitive types
function dv(array) {
    return new DataView(array.buffer, array.byteOffset);
}
/**
 * 8-bit unsigned integer
 */
exports.UINT8 = {
    len: 1,
    get(array, offset) {
        return dv(array).getUint8(offset);
    },
    put(array, offset, value) {
        dv(array).setUint8(offset, value);
        return offset + 1;
    }
};
/**
 * 16-bit unsigned integer, Little Endian byte order
 */
exports.UINT16_LE = {
    len: 2,
    get(array, offset) {
        return dv(array).getUint16(offset, true);
    },
    put(array, offset, value) {
        dv(array).setUint16(offset, value, true);
        return offset + 2;
    }
};
/**
 * 16-bit unsigned integer, Big Endian byte order
 */
exports.UINT16_BE = {
    len: 2,
    get(array, offset) {
        return dv(array).getUint16(offset);
    },
    put(array, offset, value) {
        dv(array).setUint16(offset, value);
        return offset + 2;
    }
};
/**
 * 24-bit unsigned integer, Little Endian byte order
 */
exports.UINT24_LE = {
    len: 3,
    get(array, offset) {
        const dataView = dv(array);
        return dataView.getUint8(offset) + (dataView.getUint16(offset + 1, true) << 8);
    },
    put(array, offset, value) {
        const dataView = dv(array);
        dataView.setUint8(offset, value & 0xff);
        dataView.setUint16(offset + 1, value >> 8, true);
        return offset + 3;
    }
};
/**
 * 24-bit unsigned integer, Big Endian byte order
 */
exports.UINT24_BE = {
    len: 3,
    get(array, offset) {
        const dataView = dv(array);
        return (dataView.getUint16(offset) << 8) + dataView.getUint8(offset + 2);
    },
    put(array, offset, value) {
        const dataView = dv(array);
        dataView.setUint16(offset, value >> 8);
        dataView.setUint8(offset + 2, value & 0xff);
        return offset + 3;
    }
};
/**
 * 32-bit unsigned integer, Little Endian byte order
 */
exports.UINT32_LE = {
    len: 4,
    get(array, offset) {
        return dv(array).getUint32(offset, true);
    },
    put(array, offset, value) {
        dv(array).setUint32(offset, value, true);
        return offset + 4;
    }
};
/**
 * 32-bit unsigned integer, Big Endian byte order
 */
exports.UINT32_BE = {
    len: 4,
    get(array, offset) {
        return dv(array).getUint32(offset);
    },
    put(array, offset, value) {
        dv(array).setUint32(offset, value);
        return offset + 4;
    }
};
/**
 * 8-bit signed integer
 */
exports.INT8 = {
    len: 1,
    get(array, offset) {
        return dv(array).getInt8(offset);
    },
    put(array, offset, value) {
        dv(array).setInt8(offset, value);
        return offset + 1;
    }
};
/**
 * 16-bit signed integer, Big Endian byte order
 */
exports.INT16_BE = {
    len: 2,
    get(array, offset) {
        return dv(array).getInt16(offset);
    },
    put(array, offset, value) {
        dv(array).setInt16(offset, value);
        return offset + 2;
    }
};
/**
 * 16-bit signed integer, Little Endian byte order
 */
exports.INT16_LE = {
    len: 2,
    get(array, offset) {
        return dv(array).getInt16(offset, true);
    },
    put(array, offset, value) {
        dv(array).setInt16(offset, value, true);
        return offset + 2;
    }
};
/**
 * 24-bit signed integer, Little Endian byte order
 */
exports.INT24_LE = {
    len: 3,
    get(array, offset) {
        const unsigned = exports.UINT24_LE.get(array, offset);
        return unsigned > 0x7fffff ? unsigned - 0x1000000 : unsigned;
    },
    put(array, offset, value) {
        const dataView = dv(array);
        dataView.setUint8(offset, value & 0xff);
        dataView.setUint16(offset + 1, value >> 8, true);
        return offset + 3;
    }
};
/**
 * 24-bit signed integer, Big Endian byte order
 */
exports.INT24_BE = {
    len: 3,
    get(array, offset) {
        const unsigned = exports.UINT24_BE.get(array, offset);
        return unsigned > 0x7fffff ? unsigned - 0x1000000 : unsigned;
    },
    put(array, offset, value) {
        const dataView = dv(array);
        dataView.setUint16(offset, value >> 8);
        dataView.setUint8(offset + 2, value & 0xff);
        return offset + 3;
    }
};
/**
 * 32-bit signed integer, Big Endian byte order
 */
exports.INT32_BE = {
    len: 4,
    get(array, offset) {
        return dv(array).getInt32(offset);
    },
    put(array, offset, value) {
        dv(array).setInt32(offset, value);
        return offset + 4;
    }
};
/**
 * 32-bit signed integer, Big Endian byte order
 */
exports.INT32_LE = {
    len: 4,
    get(array, offset) {
        return dv(array).getInt32(offset, true);
    },
    put(array, offset, value) {
        dv(array).setInt32(offset, value, true);
        return offset + 4;
    }
};
/**
 * 64-bit unsigned integer, Little Endian byte order
 */
exports.UINT64_LE = {
    len: 8,
    get(array, offset) {
        return dv(array).getBigUint64(offset, true);
    },
    put(array, offset, value) {
        dv(array).setBigUint64(offset, value, true);
        return offset + 8;
    }
};
/**
 * 64-bit signed integer, Little Endian byte order
 */
exports.INT64_LE = {
    len: 8,
    get(array, offset) {
        return dv(array).getBigInt64(offset, true);
    },
    put(array, offset, value) {
        dv(array).setBigInt64(offset, value, true);
        return offset + 8;
    }
};
/**
 * 64-bit unsigned integer, Big Endian byte order
 */
exports.UINT64_BE = {
    len: 8,
    get(array, offset) {
        return dv(array).getBigUint64(offset);
    },
    put(array, offset, value) {
        dv(array).setBigUint64(offset, value);
        return offset + 8;
    }
};
/**
 * 64-bit signed integer, Big Endian byte order
 */
exports.INT64_BE = {
    len: 8,
    get(array, offset) {
        return dv(array).getBigInt64(offset);
    },
    put(array, offset, value) {
        dv(array).setBigInt64(offset, value);
        return offset + 8;
    }
};
/**
 * IEEE 754 16-bit (half precision) float, big endian
 */
exports.Float16_BE = {
    len: 2,
    get(dataView, offset) {
        return ieee754.read(dataView, offset, false, 10, this.len);
    },
    put(dataView, offset, value) {
        ieee754.write(dataView, value, offset, false, 10, this.len);
        return offset + this.len;
    }
};
/**
 * IEEE 754 16-bit (half precision) float, little endian
 */
exports.Float16_LE = {
    len: 2,
    get(array, offset) {
        return ieee754.read(array, offset, true, 10, this.len);
    },
    put(array, offset, value) {
        ieee754.write(array, value, offset, true, 10, this.len);
        return offset + this.len;
    }
};
/**
 * IEEE 754 32-bit (single precision) float, big endian
 */
exports.Float32_BE = {
    len: 4,
    get(array, offset) {
        return dv(array).getFloat32(offset);
    },
    put(array, offset, value) {
        dv(array).setFloat32(offset, value);
        return offset + 4;
    }
};
/**
 * IEEE 754 32-bit (single precision) float, little endian
 */
exports.Float32_LE = {
    len: 4,
    get(array, offset) {
        return dv(array).getFloat32(offset, true);
    },
    put(array, offset, value) {
        dv(array).setFloat32(offset, value, true);
        return offset + 4;
    }
};
/**
 * IEEE 754 64-bit (double precision) float, big endian
 */
exports.Float64_BE = {
    len: 8,
    get(array, offset) {
        return dv(array).getFloat64(offset);
    },
    put(array, offset, value) {
        dv(array).setFloat64(offset, value);
        return offset + 8;
    }
};
/**
 * IEEE 754 64-bit (double precision) float, little endian
 */
exports.Float64_LE = {
    len: 8,
    get(array, offset) {
        return dv(array).getFloat64(offset, true);
    },
    put(array, offset, value) {
        dv(array).setFloat64(offset, value, true);
        return offset + 8;
    }
};
/**
 * IEEE 754 80-bit (extended precision) float, big endian
 */
exports.Float80_BE = {
    len: 10,
    get(array, offset) {
        return ieee754.read(array, offset, false, 63, this.len);
    },
    put(array, offset, value) {
        ieee754.write(array, value, offset, false, 63, this.len);
        return offset + this.len;
    }
};
/**
 * IEEE 754 80-bit (extended precision) float, little endian
 */
exports.Float80_LE = {
    len: 10,
    get(array, offset) {
        return ieee754.read(array, offset, true, 63, this.len);
    },
    put(array, offset, value) {
        ieee754.write(array, value, offset, true, 63, this.len);
        return offset + this.len;
    }
};
/**
 * Ignore a given number of bytes
 */
class IgnoreType {
    /**
     * @param len number of bytes to ignore
     */
    constructor(len) {
        this.len = len;
    }
    // ToDo: don't read, but skip data
    get(array, off) {
    }
}
exports.IgnoreType = IgnoreType;
class Uint8ArrayType {
    constructor(len) {
        this.len = len;
    }
    get(array, offset) {
        return array.subarray(offset, offset + this.len);
    }
}
exports.Uint8ArrayType = Uint8ArrayType;
class BufferType {
    constructor(len) {
        this.len = len;
    }
    get(uint8Array, off) {
        return Buffer.from(uint8Array.subarray(off, off + this.len));
    }
}
exports.BufferType = BufferType;
/**
 * Consume a fixed number of bytes from the stream and return a string with a specified encoding.
 */
class StringType {
    constructor(len, encoding) {
        this.len = len;
        this.encoding = encoding;
    }
    get(uint8Array, offset) {
        return Buffer.from(uint8Array).toString(this.encoding, offset, offset + this.len);
    }
}
exports.StringType = StringType;
/**
 * ANSI Latin 1 String
 * Using windows-1252 / ISO 8859-1 decoding
 */
class AnsiStringType {
    constructor(len) {
        this.len = len;
    }
    static decode(buffer, offset, until) {
        let str = '';
        for (let i = offset; i < until; ++i) {
            str += AnsiStringType.codePointToString(AnsiStringType.singleByteDecoder(buffer[i]));
        }
        return str;
    }
    static inRange(a, min, max) {
        return min <= a && a <= max;
    }
    static codePointToString(cp) {
        if (cp <= 0xFFFF) {
            return String.fromCharCode(cp);
        }
        else {
            cp -= 0x10000;
            return String.fromCharCode((cp >> 10) + 0xD800, (cp & 0x3FF) + 0xDC00);
        }
    }
    static singleByteDecoder(bite) {
        if (AnsiStringType.inRange(bite, 0x00, 0x7F)) {
            return bite;
        }
        const codePoint = AnsiStringType.windows1252[bite - 0x80];
        if (codePoint === null) {
            throw Error('invaliding encoding');
        }
        return codePoint;
    }
    get(buffer, offset = 0) {
        return AnsiStringType.decode(buffer, offset, offset + this.len);
    }
}
exports.AnsiStringType = AnsiStringType;
AnsiStringType.windows1252 = [8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352,
    8249, 338, 141, 381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732,
    8482, 353, 8250, 339, 157, 382, 376, 160, 161, 162, 163, 164, 165, 166, 167, 168,
    169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184,
    185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200,
    201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216,
    217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232,
    233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247,
    248, 249, 250, 251, 252, 253, 254, 255];


/***/ },

/***/ "../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/index.js"
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

"use strict";
// ESM COMPAT FLAG
__webpack_require__.r(__webpack_exports__);

// EXPORTS
__webpack_require__.d(__webpack_exports__, {
  BlendMode: () => (/* reexport */ BlendMode),
  HorizontalAlign: () => (/* reexport */ HorizontalAlign),
  VerticalAlign: () => (/* reexport */ VerticalAlign),
  composite: () => (/* reexport */ composite),
  createJimp: () => (/* binding */ createJimp),
  getExifOrientation: () => (/* reexport */ getExifOrientation)
});

// NAMESPACE OBJECT: ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/utils/composite-modes.js
var composite_modes_namespaceObject = {};
__webpack_require__.r(composite_modes_namespaceObject);
__webpack_require__.d(composite_modes_namespaceObject, {
  add: () => (add),
  darken: () => (darken),
  difference: () => (difference),
  dstOver: () => (dstOver),
  exclusion: () => (exclusion),
  hardLight: () => (hardLight),
  lighten: () => (lighten),
  multiply: () => (multiply),
  names: () => (names),
  overlay: () => (overlay),
  screen: () => (screen),
  srcOver: () => (srcOver)
});

// EXTERNAL MODULE: ../../node_modules/.pnpm/zod@3.25.76/node_modules/zod/v3/types.js + 4 modules
var types = __webpack_require__("../../node_modules/.pnpm/zod@3.25.76/node_modules/zod/v3/types.js");
;// ../../node_modules/.pnpm/@jimp+types@1.6.0/node_modules/@jimp/types/dist/esm/index.js

var Edge;
(function (Edge) {
    Edge[Edge["EXTEND"] = 1] = "EXTEND";
    Edge[Edge["WRAP"] = 2] = "WRAP";
    Edge[Edge["CROP"] = 3] = "CROP";
})(Edge || (Edge = {}));
const JimpClassSchema = types/* object */.Ik({
    bitmap: types/* object */.Ik({
        data: types/* union */.KC([types/* instanceof */.Nl(Buffer), types/* instanceof */.Nl(Uint8Array)]),
        width: types/* number */.ai(),
        height: types/* number */.ai(),
    }),
});
//# sourceMappingURL=index.js.map
// EXTERNAL MODULE: ../../node_modules/.pnpm/@jimp+utils@1.6.0/node_modules/@jimp/utils/dist/esm/index.js + 1 modules
var esm = __webpack_require__("../../node_modules/.pnpm/@jimp+utils@1.6.0/node_modules/@jimp/utils/dist/esm/index.js");
// EXTERNAL MODULE: ../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/core.js
var core = __webpack_require__("../../node_modules/.pnpm/file-type@16.5.4/node_modules/file-type/core.js");
;// ../../node_modules/.pnpm/await-to-js@3.0.0/node_modules/await-to-js/dist/await-to-js.es5.js
/**
 * @param { Promise } promise
 * @param { Object= } errorExt - Additional Information you can pass to the err object
 * @return { Promise }
 */
function to(promise, errorExt) {
    return promise
        .then(function (data) { return [null, data]; })
        .catch(function (err) {
        if (errorExt) {
            Object.assign(err, errorExt);
        }
        return [err, undefined];
    });
}


/* harmony default export */ const await_to_js_es5 = ((/* unused pure expression or super */ null && (to)));
//# sourceMappingURL=await-to-js.es5.js.map

// EXTERNAL MODULE: external "fs"
var external_fs_ = __webpack_require__("fs");
;// ../../node_modules/.pnpm/@jimp+file-ops@1.6.0/node_modules/@jimp/file-ops/dist/esm/index.js


const readFile = external_fs_.promises.readFile;
const writeFile = external_fs_.promises.writeFile;
//# sourceMappingURL=index.js.map
// EXTERNAL MODULE: ../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/lite.js
var lite = __webpack_require__("../../node_modules/.pnpm/mime@3.0.0/node_modules/mime/lite.js");
;// ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/utils/constants.js
/* Horizontal align modes for cover, contain, bit masks */
var HorizontalAlign;
(function (HorizontalAlign) {
    HorizontalAlign[HorizontalAlign["LEFT"] = 1] = "LEFT";
    HorizontalAlign[HorizontalAlign["CENTER"] = 2] = "CENTER";
    HorizontalAlign[HorizontalAlign["RIGHT"] = 4] = "RIGHT";
})(HorizontalAlign || (HorizontalAlign = {}));
/* Vertical align modes for cover, contain, bit masks */
var VerticalAlign;
(function (VerticalAlign) {
    VerticalAlign[VerticalAlign["TOP"] = 8] = "TOP";
    VerticalAlign[VerticalAlign["MIDDLE"] = 16] = "MIDDLE";
    VerticalAlign[VerticalAlign["BOTTOM"] = 32] = "BOTTOM";
})(VerticalAlign || (VerticalAlign = {}));
/**
 * How to blend two images together
 */
var BlendMode;
(function (BlendMode) {
    /**
     * Composite the source image over the destination image.
     * This is the default value. It represents the most intuitive case, where shapes are painted on top of what is below, with transparent areas showing the destination layer.
     */
    BlendMode["SRC_OVER"] = "srcOver";
    /** Composite the source image under the destination image. */
    BlendMode["DST_OVER"] = "dstOver";
    /**
     * Multiply the color components of the source and destination images.
     * This can only result in the same or darker colors (multiplying by white, 1.0, results in no change; multiplying by black, 0.0, results in black).
     * When compositing two opaque images, this has similar effect to overlapping two transparencies on a projector.
     *
     * This mode is useful for coloring shadows.
     */
    BlendMode["MULTIPLY"] = "multiply";
    /**
     * The Add mode adds the color information of the base layers and the blending layer.
     * In digital terms, adding color increases the brightness.
     */
    BlendMode["ADD"] = "add";
    /**
     * Multiply the inverse of the components of the source and destination images, and inverse the result.
     * Inverting the components means that a fully saturated channel (opaque white) is treated as the value 0.0, and values normally treated as 0.0 (black, transparent) are treated as 1.0.
     * This is essentially the same as modulate blend mode, but with the values of the colors inverted before the multiplication and the result being inverted back before rendering.
     * This can only result in the same or lighter colors (multiplying by black, 1.0, results in no change; multiplying by white, 0.0, results in white). Similarly, in the alpha channel, it can only result in more opaque colors.
     * This has similar effect to two projectors displaying their images on the same screen simultaneously.
     */
    BlendMode["SCREEN"] = "screen";
    /**
     * Multiply the components of the source and destination images after adjusting them to favor the destination.
     * Specifically, if the destination value is smaller, this multiplies it with the source value, whereas is the source value is smaller, it multiplies the inverse of the source value with the inverse of the destination value, then inverts the result.
     * Inverting the components means that a fully saturated channel (opaque white) is treated as the value 0.0, and values normally treated as 0.0 (black, transparent) are treated as 1.0.
     *
     * The Overlay mode behaves like Screen mode in bright areas, and like Multiply mode in darker areas.
     * With this mode, the bright areas will look brighter and the dark areas will look darker.
     */
    BlendMode["OVERLAY"] = "overlay";
    /**
     * Composite the source and destination image by choosing the lowest value from each color channel.
     * The opacity of the output image is computed in the same way as for srcOver.
     */
    BlendMode["DARKEN"] = "darken";
    /**
     * Composite the source and destination image by choosing the highest value from each color channel.
     * The opacity of the output image is computed in the same way as for srcOver.
     */
    BlendMode["LIGHTEN"] = "lighten";
    /**
     * Multiply the components of the source and destination images after adjusting them to favor the source.
     * Specifically, if the source value is smaller, this multiplies it with the destination value, whereas is the destination value is smaller, it multiplies the inverse of the destination value with the inverse of the source value, then inverts the result.
     * Inverting the components means that a fully saturated channel (opaque white) is treated as the value 0.0, and values normally treated as 0.0 (black, transparent) are treated as 1.0.
     *
     * The effect of the Hard light mode depends on the density of the superimposed color. Using bright colors on the blending layer will create a brighter effect like the Screen modes, while dark colors will create darker colors like the Multiply mode.
     */
    BlendMode["HARD_LIGHT"] = "hardLight";
    /**
     * Subtract the smaller value from the bigger value for each channel.
     * Compositing black has no effect; compositing white inverts the colors of the other image.
     * The opacity of the output image is computed in the same way as for srcOver.
     * The effect is similar to exclusion but harsher.
     */
    BlendMode["DIFFERENCE"] = "difference";
    /**
     * Subtract double the product of the two images from the sum of the two images.
     * Compositing black has no effect; compositing white inverts the colors of the other image.
     * The opacity of the output image is computed in the same way as for srcOver.
     * The effect is similar to difference but softer.
     */
    BlendMode["EXCLUSION"] = "exclusion";
})(BlendMode || (BlendMode = {}));
//# sourceMappingURL=constants.js.map
;// ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/utils/composite-modes.js
function srcOver(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const r = (src.r * src.a + dst.r * dst.a * (1 - src.a)) / a;
    const g = (src.g * src.a + dst.g * dst.a * (1 - src.a)) / a;
    const b = (src.b * src.a + dst.b * dst.a * (1 - src.a)) / a;
    return { r, g, b, a };
}
function dstOver(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const r = (dst.r * dst.a + src.r * src.a * (1 - dst.a)) / a;
    const g = (dst.g * dst.a + src.g * src.a * (1 - dst.a)) / a;
    const b = (dst.b * dst.a + src.b * src.a * (1 - dst.a)) / a;
    return { r, g, b, a };
}
function multiply(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (sra * dra + sra * (1 - dst.a) + dra * (1 - src.a)) / a;
    const g = (sga * dga + sga * (1 - dst.a) + dga * (1 - src.a)) / a;
    const b = (sba * dba + sba * (1 - dst.a) + dba * (1 - src.a)) / a;
    return { r, g, b, a };
}
function add(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (sra + dra) / a;
    const g = (sga + dga) / a;
    const b = (sba + dba) / a;
    return { r, g, b, a };
}
function screen(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (sra * dst.a +
        dra * src.a -
        sra * dra +
        sra * (1 - dst.a) +
        dra * (1 - src.a)) /
        a;
    const g = (sga * dst.a +
        dga * src.a -
        sga * dga +
        sga * (1 - dst.a) +
        dga * (1 - src.a)) /
        a;
    const b = (sba * dst.a +
        dba * src.a -
        sba * dba +
        sba * (1 - dst.a) +
        dba * (1 - src.a)) /
        a;
    return { r, g, b, a };
}
function overlay(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (2 * dra <= dst.a
        ? 2 * sra * dra + sra * (1 - dst.a) + dra * (1 - src.a)
        : sra * (1 + dst.a) + dra * (1 + src.a) - 2 * dra * sra - dst.a * src.a) /
        a;
    const g = (2 * dga <= dst.a
        ? 2 * sga * dga + sga * (1 - dst.a) + dga * (1 - src.a)
        : sga * (1 + dst.a) + dga * (1 + src.a) - 2 * dga * sga - dst.a * src.a) /
        a;
    const b = (2 * dba <= dst.a
        ? 2 * sba * dba + sba * (1 - dst.a) + dba * (1 - src.a)
        : sba * (1 + dst.a) + dba * (1 + src.a) - 2 * dba * sba - dst.a * src.a) /
        a;
    return { r, g, b, a };
}
function darken(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (Math.min(sra * dst.a, dra * src.a) +
        sra * (1 - dst.a) +
        dra * (1 - src.a)) /
        a;
    const g = (Math.min(sga * dst.a, dga * src.a) +
        sga * (1 - dst.a) +
        dga * (1 - src.a)) /
        a;
    const b = (Math.min(sba * dst.a, dba * src.a) +
        sba * (1 - dst.a) +
        dba * (1 - src.a)) /
        a;
    return { r, g, b, a };
}
function lighten(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (Math.max(sra * dst.a, dra * src.a) +
        sra * (1 - dst.a) +
        dra * (1 - src.a)) /
        a;
    const g = (Math.max(sga * dst.a, dga * src.a) +
        sga * (1 - dst.a) +
        dga * (1 - src.a)) /
        a;
    const b = (Math.max(sba * dst.a, dba * src.a) +
        sba * (1 - dst.a) +
        dba * (1 - src.a)) /
        a;
    return { r, g, b, a };
}
function hardLight(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (2 * sra <= src.a
        ? 2 * sra * dra + sra * (1 - dst.a) + dra * (1 - src.a)
        : sra * (1 + dst.a) + dra * (1 + src.a) - 2 * dra * sra - dst.a * src.a) /
        a;
    const g = (2 * sga <= src.a
        ? 2 * sga * dga + sga * (1 - dst.a) + dga * (1 - src.a)
        : sga * (1 + dst.a) + dga * (1 + src.a) - 2 * dga * sga - dst.a * src.a) /
        a;
    const b = (2 * sba <= src.a
        ? 2 * sba * dba + sba * (1 - dst.a) + dba * (1 - src.a)
        : sba * (1 + dst.a) + dba * (1 + src.a) - 2 * dba * sba - dst.a * src.a) /
        a;
    return { r, g, b, a };
}
function difference(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (sra + dra - 2 * Math.min(sra * dst.a, dra * src.a)) / a;
    const g = (sga + dga - 2 * Math.min(sga * dst.a, dga * src.a)) / a;
    const b = (sba + dba - 2 * Math.min(sba * dst.a, dba * src.a)) / a;
    return { r, g, b, a };
}
function exclusion(src, dst, ops = 1) {
    src.a *= ops;
    const a = dst.a + src.a - dst.a * src.a;
    const sra = src.r * src.a;
    const sga = src.g * src.a;
    const sba = src.b * src.a;
    const dra = dst.r * dst.a;
    const dga = dst.g * dst.a;
    const dba = dst.b * dst.a;
    const r = (sra * dst.a +
        dra * src.a -
        2 * sra * dra +
        sra * (1 - dst.a) +
        dra * (1 - src.a)) /
        a;
    const g = (sga * dst.a +
        dga * src.a -
        2 * sga * dga +
        sga * (1 - dst.a) +
        dga * (1 - src.a)) /
        a;
    const b = (sba * dst.a +
        dba * src.a -
        2 * sba * dba +
        sba * (1 - dst.a) +
        dba * (1 - src.a)) /
        a;
    return { r, g, b, a };
}
const names = [
    srcOver,
    dstOver,
    multiply,
    add,
    screen,
    overlay,
    darken,
    lighten,
    hardLight,
    difference,
    exclusion,
];
//# sourceMappingURL=composite-modes.js.map
;// ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/utils/composite.js




function composite(baseImage, src, x = 0, y = 0, options = {}) {
    if (!(src instanceof baseImage.constructor)) {
        throw new Error("The source must be a Jimp image");
    }
    if (typeof x !== "number" || typeof y !== "number") {
        throw new Error("x and y must be numbers");
    }
    const { mode = BlendMode.SRC_OVER } = options;
    let { opacitySource = 1.0, opacityDest = 1.0 } = options;
    if (typeof opacitySource !== "number" ||
        opacitySource < 0 ||
        opacitySource > 1) {
        opacitySource = 1.0;
    }
    if (typeof opacityDest !== "number" || opacityDest < 0 || opacityDest > 1) {
        opacityDest = 1.0;
    }
    const blendmode = composite_modes_namespaceObject[mode];
    // round input
    x = Math.round(x);
    y = Math.round(y);
    if (opacityDest !== 1.0) {
        baseImage.scan((_, __, idx) => {
            const v = baseImage.bitmap.data[idx + 3] * opacityDest;
            baseImage.bitmap.data[idx + 3] = v;
        });
    }
    src.scan((sx, sy, idx) => {
        const dstIdx = baseImage.getPixelIndex(x + sx, y + sy, Edge.CROP);
        if (dstIdx === -1) {
            // Skip target pixels outside of dst
            return;
        }
        const blended = blendmode({
            r: src.bitmap.data[idx + 0] / 255,
            g: src.bitmap.data[idx + 1] / 255,
            b: src.bitmap.data[idx + 2] / 255,
            a: src.bitmap.data[idx + 3] / 255,
        }, {
            r: baseImage.bitmap.data[dstIdx + 0] / 255,
            g: baseImage.bitmap.data[dstIdx + 1] / 255,
            b: baseImage.bitmap.data[dstIdx + 2] / 255,
            a: baseImage.bitmap.data[dstIdx + 3] / 255,
        }, opacitySource);
        baseImage.bitmap.data[dstIdx + 0] = (0,esm/* limit255 */.QH)(blended.r * 255);
        baseImage.bitmap.data[dstIdx + 1] = (0,esm/* limit255 */.QH)(blended.g * 255);
        baseImage.bitmap.data[dstIdx + 2] = (0,esm/* limit255 */.QH)(blended.b * 255);
        baseImage.bitmap.data[dstIdx + 3] = (0,esm/* limit255 */.QH)(blended.a * 255);
    });
    return baseImage;
}
//# sourceMappingURL=composite.js.map
// EXTERNAL MODULE: ../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/index.js
var exif_parser = __webpack_require__("../../node_modules/.pnpm/exif-parser@0.1.12/node_modules/exif-parser/index.js");
;// ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/utils/image-bitmap.js

/**
 * Obtains image orientation from EXIF metadata.
 *
 * @param img a Jimp image object
 * @returns a number 1-8 representing EXIF orientation,
 *          in particular 1 if orientation tag is missing
 */
function getExifOrientation(img) {
    const _exif = img._exif;
    return (_exif && _exif.tags && _exif.tags.Orientation) || 1;
}
/**
 * Returns a function which translates EXIF-rotated coordinates into
 * non-rotated ones.
 *
 * Transformation reference: http://sylvana.net/jpegcrop/exif_orientation.html.
 *
 * @param img a Jimp image object
 * @returns transformation function for transformBitmap().
 */
function getExifOrientationTransformation(img) {
    const w = img.bitmap.width;
    const h = img.bitmap.height;
    switch (getExifOrientation(img)) {
        case 1: // Horizontal (normal)
            // does not need to be supported here
            return null;
        case 2: // Mirror horizontal
            return function (x, y) {
                return [w - x - 1, y];
            };
        case 3: // Rotate 180
            return function (x, y) {
                return [w - x - 1, h - y - 1];
            };
        case 4: // Mirror vertical
            return function (x, y) {
                return [x, h - y - 1];
            };
        case 5: // Mirror horizontal and rotate 270 CW
            return function (x, y) {
                return [y, x];
            };
        case 6: // Rotate 90 CW
            return function (x, y) {
                return [y, h - x - 1];
            };
        case 7: // Mirror horizontal and rotate 90 CW
            return function (x, y) {
                return [w - y - 1, h - x - 1];
            };
        case 8: // Rotate 270 CW
            return function (x, y) {
                return [w - y - 1, x];
            };
        default:
            return null;
    }
}
/**
 * Transforms bitmap in place (moves pixels around) according to given
 * transformation function.
 *
 * @param img a Jimp image object, which bitmap is supposed to
 *        be transformed
 * @param width bitmap width after the transformation
 * @param height bitmap height after the transformation
 * @param transformation transformation function which defines pixel
 *        mapping between new and source bitmap. It takes a pair of coordinates
 *        in the target, and returns a respective pair of coordinates in
 *        the source bitmap, i.e. has following form:
 *        `function(new_x, new_y) { return [src_x, src_y] }`.
 */
function transformBitmap(img, width, height, transformation) {
    // Underscore-prefixed values are related to the source bitmap
    // Their counterparts with no prefix are related to the target bitmap
    const _data = img.bitmap.data;
    const _width = img.bitmap.width;
    const data = Buffer.alloc(_data.length);
    for (let x = 0; x < width; x++) {
        for (let y = 0; y < height; y++) {
            const [_x, _y] = transformation(x, y);
            const idx = (width * y + x) << 2;
            const _idx = (_width * _y + _x) << 2;
            const pixel = _data.readUInt32BE(_idx);
            data.writeUInt32BE(pixel, idx);
        }
    }
    img.bitmap.data = data;
    img.bitmap.width = width;
    img.bitmap.height = height;
    // @ts-expect-error Accessing private property
    img._exif.tags.Orientation = 1;
}
/**
 * Automagically rotates an image based on its EXIF data (if present).
 * @param img  a Jimp image object
 */
function exifRotate(img) {
    if (getExifOrientation(img) < 2) {
        return;
    }
    const transformation = getExifOrientationTransformation(img);
    const swapDimensions = getExifOrientation(img) > 4;
    const newWidth = swapDimensions ? img.bitmap.height : img.bitmap.width;
    const newHeight = swapDimensions ? img.bitmap.width : img.bitmap.height;
    if (transformation) {
        transformBitmap(img, newWidth, newHeight, transformation);
    }
}
async function attemptExifRotate(image, buffer) {
    try {
        image._exif =
            exif_parser.create(buffer).parse();
        exifRotate(image); // EXIF data
    }
    catch {
        // do nothing
    }
}
//# sourceMappingURL=image-bitmap.js.map
;// ../../node_modules/.pnpm/@jimp+core@1.6.0/node_modules/@jimp/core/dist/esm/index.js








const emptyBitmap = {
    data: Buffer.alloc(0),
    width: 0,
    height: 0,
};
/**
 * Prepare a Buffer object from the arrayBuffer.
 */
function bufferFromArrayBuffer(arrayBuffer) {
    const buffer = Buffer.alloc(arrayBuffer.byteLength);
    const view = new Uint8Array(arrayBuffer);
    for (let i = 0; i < buffer.length; ++i) {
        buffer[i] = view[i];
    }
    return buffer;
}



/**
 * Create a Jimp class that support the given image formats and methods
 */
function createJimp({ plugins: pluginsArg, formats: formatsArg, } = {}) {
    const plugins = pluginsArg || [];
    const formats = (formatsArg || []).map((format) => format());
    const CustomJimp = class Jimp {
        /**
         * The bitmap data of the image
         */
        bitmap = emptyBitmap;
        /**  Default color to use for new pixels */
        background = 0x00000000;
        /** Formats that can be used with Jimp */
        formats = [];
        /** The original MIME type of the image */
        mime;
        constructor(options = emptyBitmap) {
            // Add the formats
            this.formats = formats;
            if ("data" in options) {
                this.bitmap = options;
            }
            else {
                this.bitmap = {
                    data: Buffer.alloc(options.width * options.height * 4),
                    width: options.width,
                    height: options.height,
                };
                if (options.color) {
                    this.background =
                        typeof options.color === "string"
                            ? (0,esm/* cssColorToHex */.by)(options.color)
                            : options.color;
                    for (let i = 0; i < this.bitmap.data.length; i += 4) {
                        this.bitmap.data.writeUInt32BE(this.background, i);
                    }
                }
            }
            // Add the plugins
            for (const methods of plugins) {
                for (const key in methods) {
                    this[key] = (...args) => {
                        const result = methods[key]?.(this, ...args);
                        if (typeof result === "object" && "bitmap" in result) {
                            this.bitmap = result.bitmap;
                            return this;
                        }
                        return result;
                    };
                }
            }
        }
        /**
         * Create a Jimp instance from a URL, a file path, or a Buffer
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * // Read from a file path
         * const image = await Jimp.read("test/image.png");
         *
         * // Read from a URL
         * const image = await Jimp.read("https://upload.wikimedia.org/wikipedia/commons/0/01/Bot-Test.jpg");
         * ```
         */
        static async read(url, options) {
            if (Buffer.isBuffer(url) || url instanceof ArrayBuffer) {
                return this.fromBuffer(url);
            }
            if ((0,external_fs_.existsSync)(url)) {
                return this.fromBuffer(await readFile(url));
            }
            const [fetchErr, response] = await to(fetch(url));
            if (fetchErr) {
                throw new Error(`Could not load Buffer from URL: ${url}`);
            }
            if (!response.ok) {
                throw new Error(`HTTP Status ${response.status} for url ${url}`);
            }
            const [arrayBufferErr, data] = await to(response.arrayBuffer());
            if (arrayBufferErr) {
                throw new Error(`Could not load Buffer from ${url}`);
            }
            const buffer = bufferFromArrayBuffer(data);
            return this.fromBuffer(buffer, options);
        }
        /**
         * Create a Jimp instance from a bitmap.
         * The difference between this and just using the constructor is that this will
         * convert raw image data into the bitmap format that Jimp uses.
         *
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = Jimp.fromBitmap({
         *   data: Buffer.from([
         *     0xffffffff, 0xffffffff, 0xffffffff,
         *     0xffffffff, 0xffffffff, 0xffffffff,
         *     0xffffffff, 0xffffffff, 0xffffffff,
         *   ]),
         *   width: 3,
         *   height: 3,
         * });
         * ```
         */
        static fromBitmap(bitmap) {
            let data;
            if (bitmap.data instanceof Buffer) {
                data = Buffer.from(bitmap.data);
            }
            if (bitmap.data instanceof Uint8Array ||
                bitmap.data instanceof Uint8ClampedArray) {
                data = Buffer.from(bitmap.data.buffer);
            }
            if (Array.isArray(bitmap.data)) {
                data = Buffer.concat(bitmap.data.map((hex) => Buffer.from(hex.toString(16).padStart(8, "0"), "hex")));
            }
            if (!data) {
                throw new Error("data must be a Buffer");
            }
            if (typeof bitmap.height !== "number" ||
                typeof bitmap.width !== "number") {
                throw new Error("bitmap must have width and height");
            }
            return new CustomJimp({
                height: bitmap.height,
                width: bitmap.width,
                data,
            });
        }
        /**
         * Parse a bitmap with the loaded image types.
         *
         * @param buffer Raw image data
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const buffer = await fs.readFile("test/image.png");
         * const image = await Jimp.fromBuffer(buffer);
         * ```
         */
        static async fromBuffer(buffer, options) {
            const actualBuffer = buffer instanceof ArrayBuffer ? bufferFromArrayBuffer(buffer) : buffer;
            const mime = await core.fromBuffer(actualBuffer);
            if (!mime || !mime.mime) {
                throw new Error("Could not find MIME for Buffer");
            }
            const format = formats.find((format) => format.mime === mime.mime);
            if (!format || !format.decode) {
                throw new Error(`Mime type ${mime.mime} does not support decoding`);
            }
            const image = new CustomJimp(await format.decode(actualBuffer, options?.[format.mime]));
            image.mime = mime.mime;
            attemptExifRotate(image, actualBuffer);
            return image;
        }
        /**
         * Nicely format Jimp object when sent to the console e.g. console.log(image)
         * @returns Pretty printed jimp object
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = await Jimp.read("test/image.png");
         *
         * console.log(image);
         * ```
         */
        inspect() {
            return ("<Jimp " +
                (this.bitmap === emptyBitmap
                    ? "pending..."
                    : this.bitmap.width + "x" + this.bitmap.height) +
                ">");
        }
        /**
         * Nicely format Jimp object when converted to a string
         * @returns pretty printed
         */
        toString() {
            return "[object Jimp]";
        }
        /** Get the width of the image */
        get width() {
            return this.bitmap.width;
        }
        /** Get the height of the image */
        get height() {
            return this.bitmap.height;
        }
        /**
         * Converts the Jimp instance to an image buffer
         * @param mime The mime type to export to
         * @param options The options to use when exporting
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         * import { promises as fs } from "fs";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * await image.write("test/output.jpeg", {
         *   quality: 50,
         * });
         * ```
         */
        async getBuffer(mime, options) {
            const format = this.formats.find((format) => format.mime === mime);
            if (!format || !format.encode) {
                throw new Error(`Unsupported MIME type: ${mime}`);
            }
            let outputImage;
            if (format.hasAlpha) {
                // eslint-disable-next-line @typescript-eslint/no-this-alias
                outputImage = this;
            }
            else {
                outputImage = new CustomJimp({
                    width: this.bitmap.width,
                    height: this.bitmap.height,
                    color: this.background,
                });
                composite(outputImage, this);
            }
            return format.encode(outputImage.bitmap, options);
        }
        /**
         * Converts the image to a base 64 string
         *
         * @param mime The mime type to export to
         * @param options The options to use when exporting
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = Jimp.fromBuffer(Buffer.from([
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         * ]));
         *
         * const base64 = image.getBase64("image/jpeg", {
         *   quality: 50,
         * });
         * ```
         */
        async getBase64(mime, options) {
            const data = await this.getBuffer(mime, options);
            return "data:" + mime + ";base64," + data.toString("base64");
        }
        /**
         * Write the image to a file
         * @param path the path to write the image to
         * @param options the options to use when writing the image
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = Jimp.fromBuffer(Buffer.from([
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         *   0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00,
         * ]));
         *
         * await image.write("test/output.png");
         * ```
         */
        async write(path, options) {
            const mimeType = lite.getType(path);
            await writeFile(path, await this.getBuffer(mimeType, options));
        }
        /**
         * Clone the image into a new Jimp instance.
         * @param this
         * @returns A new Jimp instance
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * const clone = image.clone();
         * ```
         */
        clone() {
            return new CustomJimp({
                ...this.bitmap,
                data: Buffer.from(this.bitmap.data),
            });
        }
        /**
         * Returns the offset of a pixel in the bitmap buffer
         * @param x the x coordinate
         * @param y the y coordinate
         * @param edgeHandling (optional) define how to sum pixels from outside the border
         * @returns the index of the pixel or -1 if not found
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * image.getPixelIndex(1, 1); // 2
         * ```
         */
        getPixelIndex(x, y, edgeHandling) {
            let xi;
            let yi;
            if (!edgeHandling) {
                edgeHandling = Edge.EXTEND;
            }
            if (typeof x !== "number" || typeof y !== "number") {
                throw new Error("x and y must be numbers");
            }
            // round input
            x = Math.round(x);
            y = Math.round(y);
            xi = x;
            yi = y;
            if (edgeHandling === Edge.EXTEND) {
                if (x < 0)
                    xi = 0;
                if (x >= this.bitmap.width)
                    xi = this.bitmap.width - 1;
                if (y < 0)
                    yi = 0;
                if (y >= this.bitmap.height)
                    yi = this.bitmap.height - 1;
            }
            if (edgeHandling === Edge.WRAP) {
                if (x < 0) {
                    xi = this.bitmap.width + x;
                }
                if (x >= this.bitmap.width) {
                    xi = x % this.bitmap.width;
                }
                if (y < 0) {
                    yi = this.bitmap.height + y;
                }
                if (y >= this.bitmap.height) {
                    yi = y % this.bitmap.height;
                }
            }
            let i = (this.bitmap.width * yi + xi) << 2;
            // if out of bounds index is -1
            if (xi < 0 || xi >= this.bitmap.width) {
                i = -1;
            }
            if (yi < 0 || yi >= this.bitmap.height) {
                i = -1;
            }
            return i;
        }
        /**
         * Returns the hex color value of a pixel
         * @param x the x coordinate
         * @param y the y coordinate
         * @returns the color of the pixel
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * image.getPixelColor(1, 1); // 0xffffffff
         * ```
         */
        getPixelColor(x, y) {
            if (typeof x !== "number" || typeof y !== "number") {
                throw new Error("x and y must be numbers");
            }
            const idx = this.getPixelIndex(x, y);
            return this.bitmap.data.readUInt32BE(idx);
        }
        /**
         * Sets the hex colour value of a pixel
         *
         * @param hex color to set
         * @param x the x coordinate
         * @param y the y coordinate
         *
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * image.setPixelColor(0xff0000ff, 0, 0);
         * ```
         */
        setPixelColor(hex, x, y) {
            if (typeof hex !== "number" ||
                typeof x !== "number" ||
                typeof y !== "number") {
                throw new Error("hex, x and y must be numbers");
            }
            const idx = this.getPixelIndex(x, y);
            this.bitmap.data.writeUInt32BE(hex, idx);
            return this;
        }
        /**
         * Determine if the image contains opaque pixels.
         *
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffaa });
         * const image2 = new Jimp({ width: 3, height: 3, color: 0xff0000ff });
         *
         * image.hasAlpha(); // false
         * image2.hasAlpha(); // true
         * ```
         */
        hasAlpha() {
            const { width, height, data } = this.bitmap;
            const byteLen = (width * height) << 2;
            for (let idx = 3; idx < byteLen; idx += 4) {
                if (data[idx] !== 0xff) {
                    return true;
                }
            }
            return false;
        }
        /**
         * Composites a source image over to this image respecting alpha channels
         * @param src the source Jimp instance
         * @param x the x position to blit the image
         * @param y the y position to blit the image
         * @param options determine what mode to use
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 10, height: 10, color: 0xffffffff });
         * const image2 = new Jimp({ width: 3, height: 3, color: 0xff0000ff });
         *
         * image.composite(image2, 3, 3);
         * ```
         */
        composite(src, x = 0, y = 0, options = {}) {
            return composite(this, src, x, y, options);
        }
        scan(x, y, w, h, f) {
            return (0,esm/* scan */.SQ)(this, x, y, w, h, f);
        }
        /**
         * Iterate scan through a region of the bitmap
         * @param x the x coordinate to begin the scan at
         * @param y the y coordinate to begin the scan at
         * @param w the width of the scan region
         * @param h the height of the scan region
         * @example
         * ```ts
         * import { Jimp } from "jimp";
         *
         * const image = new Jimp({ width: 3, height: 3, color: 0xffffffff });
         *
         * for (const { x, y, idx, image } of j.scanIterator()) {
         *   // do something with the pixel
         * }
         * ```
         */
        scanIterator(x = 0, y = 0, w = this.bitmap.width, h = this.bitmap.height) {
            if (typeof x !== "number" || typeof y !== "number") {
                throw new Error("x and y must be numbers");
            }
            if (typeof w !== "number" || typeof h !== "number") {
                throw new Error("w and h must be numbers");
            }
            return (0,esm/* scanIterator */.QQ)(this, x, y, w, h);
        }
    };
    return CustomJimp;
}
//# sourceMappingURL=index.js.map

/***/ }

};
;