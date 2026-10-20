const __import_meta_url=require('node:url').pathToFileURL(__filename).href;
"use strict";
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// ../packages/agent/src/tools/core/edit/diff-worker.ts
var diff_worker_exports = {};
__export(diff_worker_exports, {
  default: () => calculateDiff
});
module.exports = __toCommonJS(diff_worker_exports);

// ../node_modules/.pnpm/diff@8.0.3/node_modules/diff/libesm/diff/base.js
var Diff = class {
  diff(oldStr, newStr, options = {}) {
    let callback;
    if (typeof options === "function") {
      callback = options;
      options = {};
    } else if ("callback" in options) {
      callback = options.callback;
    }
    const oldString = this.castInput(oldStr, options);
    const newString = this.castInput(newStr, options);
    const oldTokens = this.removeEmpty(this.tokenize(oldString, options));
    const newTokens = this.removeEmpty(this.tokenize(newString, options));
    return this.diffWithOptionsObj(oldTokens, newTokens, options, callback);
  }
  diffWithOptionsObj(oldTokens, newTokens, options, callback) {
    var _a;
    const done = (value) => {
      value = this.postProcess(value, options);
      if (callback) {
        setTimeout(function() {
          callback(value);
        }, 0);
        return void 0;
      } else {
        return value;
      }
    };
    const newLen = newTokens.length, oldLen = oldTokens.length;
    let editLength = 1;
    let maxEditLength = newLen + oldLen;
    if (options.maxEditLength != null) {
      maxEditLength = Math.min(maxEditLength, options.maxEditLength);
    }
    const maxExecutionTime = (_a = options.timeout) !== null && _a !== void 0 ? _a : Infinity;
    const abortAfterTimestamp = Date.now() + maxExecutionTime;
    const bestPath = [{ oldPos: -1, lastComponent: void 0 }];
    let newPos = this.extractCommon(bestPath[0], newTokens, oldTokens, 0, options);
    if (bestPath[0].oldPos + 1 >= oldLen && newPos + 1 >= newLen) {
      return done(this.buildValues(bestPath[0].lastComponent, newTokens, oldTokens));
    }
    let minDiagonalToConsider = -Infinity, maxDiagonalToConsider = Infinity;
    const execEditLength = () => {
      for (let diagonalPath = Math.max(minDiagonalToConsider, -editLength); diagonalPath <= Math.min(maxDiagonalToConsider, editLength); diagonalPath += 2) {
        let basePath;
        const removePath = bestPath[diagonalPath - 1], addPath = bestPath[diagonalPath + 1];
        if (removePath) {
          bestPath[diagonalPath - 1] = void 0;
        }
        let canAdd = false;
        if (addPath) {
          const addPathNewPos = addPath.oldPos - diagonalPath;
          canAdd = addPath && 0 <= addPathNewPos && addPathNewPos < newLen;
        }
        const canRemove = removePath && removePath.oldPos + 1 < oldLen;
        if (!canAdd && !canRemove) {
          bestPath[diagonalPath] = void 0;
          continue;
        }
        if (!canRemove || canAdd && removePath.oldPos < addPath.oldPos) {
          basePath = this.addToPath(addPath, true, false, 0, options);
        } else {
          basePath = this.addToPath(removePath, false, true, 1, options);
        }
        newPos = this.extractCommon(basePath, newTokens, oldTokens, diagonalPath, options);
        if (basePath.oldPos + 1 >= oldLen && newPos + 1 >= newLen) {
          return done(this.buildValues(basePath.lastComponent, newTokens, oldTokens)) || true;
        } else {
          bestPath[diagonalPath] = basePath;
          if (basePath.oldPos + 1 >= oldLen) {
            maxDiagonalToConsider = Math.min(maxDiagonalToConsider, diagonalPath - 1);
          }
          if (newPos + 1 >= newLen) {
            minDiagonalToConsider = Math.max(minDiagonalToConsider, diagonalPath + 1);
          }
        }
      }
      editLength++;
    };
    if (callback) {
      (function exec() {
        setTimeout(function() {
          if (editLength > maxEditLength || Date.now() > abortAfterTimestamp) {
            return callback(void 0);
          }
          if (!execEditLength()) {
            exec();
          }
        }, 0);
      })();
    } else {
      while (editLength <= maxEditLength && Date.now() <= abortAfterTimestamp) {
        const ret = execEditLength();
        if (ret) {
          return ret;
        }
      }
    }
  }
  addToPath(path, added, removed, oldPosInc, options) {
    const last = path.lastComponent;
    if (last && !options.oneChangePerToken && last.added === added && last.removed === removed) {
      return {
        oldPos: path.oldPos + oldPosInc,
        lastComponent: { count: last.count + 1, added, removed, previousComponent: last.previousComponent }
      };
    } else {
      return {
        oldPos: path.oldPos + oldPosInc,
        lastComponent: { count: 1, added, removed, previousComponent: last }
      };
    }
  }
  extractCommon(basePath, newTokens, oldTokens, diagonalPath, options) {
    const newLen = newTokens.length, oldLen = oldTokens.length;
    let oldPos = basePath.oldPos, newPos = oldPos - diagonalPath, commonCount = 0;
    while (newPos + 1 < newLen && oldPos + 1 < oldLen && this.equals(oldTokens[oldPos + 1], newTokens[newPos + 1], options)) {
      newPos++;
      oldPos++;
      commonCount++;
      if (options.oneChangePerToken) {
        basePath.lastComponent = { count: 1, previousComponent: basePath.lastComponent, added: false, removed: false };
      }
    }
    if (commonCount && !options.oneChangePerToken) {
      basePath.lastComponent = { count: commonCount, previousComponent: basePath.lastComponent, added: false, removed: false };
    }
    basePath.oldPos = oldPos;
    return newPos;
  }
  equals(left, right, options) {
    if (options.comparator) {
      return options.comparator(left, right);
    } else {
      return left === right || !!options.ignoreCase && left.toLowerCase() === right.toLowerCase();
    }
  }
  removeEmpty(array) {
    const ret = [];
    for (let i = 0; i < array.length; i++) {
      if (array[i]) {
        ret.push(array[i]);
      }
    }
    return ret;
  }
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  castInput(value, options) {
    return value;
  }
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  tokenize(value, options) {
    return Array.from(value);
  }
  join(chars) {
    return chars.join("");
  }
  postProcess(changeObjects, options) {
    return changeObjects;
  }
  get useLongestToken() {
    return false;
  }
  buildValues(lastComponent, newTokens, oldTokens) {
    const components = [];
    let nextComponent;
    while (lastComponent) {
      components.push(lastComponent);
      nextComponent = lastComponent.previousComponent;
      delete lastComponent.previousComponent;
      lastComponent = nextComponent;
    }
    components.reverse();
    const componentLen = components.length;
    let componentPos = 0, newPos = 0, oldPos = 0;
    for (; componentPos < componentLen; componentPos++) {
      const component = components[componentPos];
      if (!component.removed) {
        if (!component.added && this.useLongestToken) {
          let value = newTokens.slice(newPos, newPos + component.count);
          value = value.map(function(value2, i) {
            const oldValue = oldTokens[oldPos + i];
            return oldValue.length > value2.length ? oldValue : value2;
          });
          component.value = this.join(value);
        } else {
          component.value = this.join(newTokens.slice(newPos, newPos + component.count));
        }
        newPos += component.count;
        if (!component.added) {
          oldPos += component.count;
        }
      } else {
        component.value = this.join(oldTokens.slice(oldPos, oldPos + component.count));
        oldPos += component.count;
      }
    }
    return components;
  }
};

// ../node_modules/.pnpm/diff@8.0.3/node_modules/diff/libesm/diff/line.js
var LineDiff = class extends Diff {
  constructor() {
    super(...arguments);
    this.tokenize = tokenize;
  }
  equals(left, right, options) {
    if (options.ignoreWhitespace) {
      if (!options.newlineIsToken || !left.includes("\n")) {
        left = left.trim();
      }
      if (!options.newlineIsToken || !right.includes("\n")) {
        right = right.trim();
      }
    } else if (options.ignoreNewlineAtEof && !options.newlineIsToken) {
      if (left.endsWith("\n")) {
        left = left.slice(0, -1);
      }
      if (right.endsWith("\n")) {
        right = right.slice(0, -1);
      }
    }
    return super.equals(left, right, options);
  }
};
var lineDiff = new LineDiff();
function diffLines(oldStr, newStr, options) {
  return lineDiff.diff(oldStr, newStr, options);
}
function tokenize(value, options) {
  if (options.stripTrailingCr) {
    value = value.replace(/\r\n/g, "\n");
  }
  const retLines = [], linesAndNewlines = value.split(/(\n|\r\n)/);
  if (!linesAndNewlines[linesAndNewlines.length - 1]) {
    linesAndNewlines.pop();
  }
  for (let i = 0; i < linesAndNewlines.length; i++) {
    const line = linesAndNewlines[i];
    if (i % 2 && !options.newlineIsToken) {
      retLines[retLines.length - 1] += line;
    } else {
      retLines.push(line);
    }
  }
  return retLines;
}

// ../node_modules/.pnpm/diff@8.0.3/node_modules/diff/libesm/patch/create.js
function structuredPatch(oldFileName, newFileName, oldStr, newStr, oldHeader, newHeader, options) {
  let optionsObj;
  if (!options) {
    optionsObj = {};
  } else if (typeof options === "function") {
    optionsObj = { callback: options };
  } else {
    optionsObj = options;
  }
  if (typeof optionsObj.context === "undefined") {
    optionsObj.context = 4;
  }
  const context = optionsObj.context;
  if (optionsObj.newlineIsToken) {
    throw new Error("newlineIsToken may not be used with patch-generation functions, only with diffing functions");
  }
  if (!optionsObj.callback) {
    return diffLinesResultToPatch(diffLines(oldStr, newStr, optionsObj));
  } else {
    const { callback } = optionsObj;
    diffLines(oldStr, newStr, Object.assign(Object.assign({}, optionsObj), { callback: (diff) => {
      const patch2 = diffLinesResultToPatch(diff);
      callback(patch2);
    } }));
  }
  function diffLinesResultToPatch(diff) {
    if (!diff) {
      return;
    }
    diff.push({ value: "", lines: [] });
    function contextLines(lines) {
      return lines.map(function(entry) {
        return " " + entry;
      });
    }
    const hunks = [];
    let oldRangeStart = 0, newRangeStart = 0, curRange = [], oldLine = 1, newLine = 1;
    for (let i = 0; i < diff.length; i++) {
      const current = diff[i], lines = current.lines || splitLines(current.value);
      current.lines = lines;
      if (current.added || current.removed) {
        if (!oldRangeStart) {
          const prev = diff[i - 1];
          oldRangeStart = oldLine;
          newRangeStart = newLine;
          if (prev) {
            curRange = context > 0 ? contextLines(prev.lines.slice(-context)) : [];
            oldRangeStart -= curRange.length;
            newRangeStart -= curRange.length;
          }
        }
        for (const line of lines) {
          curRange.push((current.added ? "+" : "-") + line);
        }
        if (current.added) {
          newLine += lines.length;
        } else {
          oldLine += lines.length;
        }
      } else {
        if (oldRangeStart) {
          if (lines.length <= context * 2 && i < diff.length - 2) {
            for (const line of contextLines(lines)) {
              curRange.push(line);
            }
          } else {
            const contextSize = Math.min(lines.length, context);
            for (const line of contextLines(lines.slice(0, contextSize))) {
              curRange.push(line);
            }
            const hunk = {
              oldStart: oldRangeStart,
              oldLines: oldLine - oldRangeStart + contextSize,
              newStart: newRangeStart,
              newLines: newLine - newRangeStart + contextSize,
              lines: curRange
            };
            hunks.push(hunk);
            oldRangeStart = 0;
            newRangeStart = 0;
            curRange = [];
          }
        }
        oldLine += lines.length;
        newLine += lines.length;
      }
    }
    for (const hunk of hunks) {
      for (let i = 0; i < hunk.lines.length; i++) {
        if (hunk.lines[i].endsWith("\n")) {
          hunk.lines[i] = hunk.lines[i].slice(0, -1);
        } else {
          hunk.lines.splice(i + 1, 0, "\\ No newline at end of file");
          i++;
        }
      }
    }
    return {
      oldFileName,
      newFileName,
      oldHeader,
      newHeader,
      hunks
    };
  }
}
function splitLines(text) {
  const hasTrailingNl = text.endsWith("\n");
  const result = text.split("\n").map((line) => line + "\n");
  if (hasTrailingNl) {
    result.pop();
  } else {
    result.push(result.pop().slice(0, -1));
  }
  return result;
}

// ../packages/agent/src/tools/core/edit/diff-worker.ts
var CONTEXT = 3;
var CHUNK = 8192;
function withGitDiffPrefix(path, prefix) {
  if (path === "/dev/null") {
    return path;
  }
  return `${prefix}/${path}`;
}
function patch(input) {
  return structuredPatch("a", "b", input.original, input.updated, "", "", { context: CONTEXT }).hunks;
}
function equalExtent(input) {
  const { text, other, from, otherFrom, max, dir } = input;
  let run = 0;
  let step = 128;
  while (run < max) {
    const n = Math.min(step, max - run);
    const start = dir > 0 ? from + run : from - run - n;
    const otherStart = dir > 0 ? otherFrom + run : otherFrom - run - n;
    if (text.slice(start, start + n) !== other.slice(otherStart, otherStart + n)) {
      if (n === 1) break;
      step = n >> 1;
      continue;
    }
    run += n;
    if (step < CHUNK) step <<= 1;
  }
  return run;
}
function trimmedHunks(input) {
  const { original, updated } = input;
  if (original === updated) return [];
  const at = (text, index) => index <= 0 || index >= text.length || text.charCodeAt(index - 1) === 10;
  const limit = Math.min(original.length, updated.length);
  let prefix = equalExtent({
    text: original,
    other: updated,
    from: 0,
    otherFrom: 0,
    max: limit,
    dir: 1
  });
  if (prefix > 0 && original.charCodeAt(prefix - 1) !== 10) {
    prefix = original.lastIndexOf("\n", prefix - 1) + 1;
  }
  let suffix = original.length;
  let other = updated.length;
  const shared = equalExtent({
    text: original,
    other: updated,
    from: suffix,
    otherFrom: other,
    max: limit - prefix,
    dir: -1
  });
  suffix -= shared;
  other -= shared;
  if (!at(original, suffix) || !at(updated, other)) {
    const newline = original.indexOf("\n", suffix);
    const step = newline < 0 ? -1 : newline + 1 - suffix;
    other += step;
    suffix = step < 0 || other > updated.length || !at(updated, other) ? original.length : suffix + step;
  }
  const total = original.length + updated.length;
  if ((prefix + original.length - suffix) * 8 <= total) return patch({ original, updated });
  let prefixLines = 0;
  for (let index = 0; index < prefix; ) {
    const newline = original.indexOf("\n", index);
    if (newline < 0 || newline >= prefix) break;
    prefixLines++;
    index = newline + 1;
  }
  const tail = original.length - suffix;
  if (tail > suffix - prefix) {
    const seen = /* @__PURE__ */ new Set();
    const take = (text, to) => {
      let start = prefix;
      for (let index = prefix; index < to; index++) {
        if (text.charCodeAt(index) === 10) {
          seen.add(text.slice(start, index + 1));
          start = index + 1;
        }
      }
    };
    take(original, suffix);
    take(updated, updated.length - tail);
    for (let newline = original.indexOf("\n", suffix); newline >= 0; newline = original.indexOf("\n", suffix)) {
      if (!seen.has(original.slice(suffix, newline + 1))) break;
      suffix = newline + 1;
    }
  } else {
    suffix = original.length;
  }
  const keep = Math.max(0, prefixLines - CONTEXT);
  let suffixLines = 0;
  for (let index = suffix; suffixLines <= CONTEXT && index < original.length; index++) {
    if (original.charCodeAt(index) === 10) suffixLines++;
  }
  if (keep === 0 && suffixLines <= CONTEXT) return patch({ original, updated });
  const shift = (index, dir) => {
    for (let step = 0; step < CONTEXT; step++) {
      const newline = dir < 0 ? original.lastIndexOf("\n", index - 2) : original.indexOf("\n", index);
      if (newline < 0) return dir < 0 ? 0 : original.length;
      index = newline + 1;
    }
    return index;
  };
  const from = keep === 0 ? 0 : shift(prefix, -1);
  const drop = suffixLines <= CONTEXT ? 0 : original.length - shift(suffix, 1);
  if (from + drop > original.length || from + drop > updated.length || (from + drop) * 8 < total) {
    return patch({ original, updated });
  }
  const hunks = patch({
    original: original.slice(from, original.length - drop),
    updated: updated.slice(from, updated.length - drop)
  });
  for (const hunk of hunks) {
    hunk.oldStart += keep;
    hunk.newStart += keep;
  }
  return hunks;
}
function calculateDiff(params) {
  const measure = params.measure === true;
  const startedAt = measure ? performance.now() : 0;
  const ensureNL = (s) => s === "" || s.endsWith("\n") ? s : `${s}
`;
  const original = ensureNL(params.original);
  const updated = ensureNL(params.new);
  const hunks = original.length < 4096 && updated.length < 4096 ? patch({ original, updated }) : trimmedHunks({ original, updated });
  const formatRange = (start, lineCount) => {
    if (lineCount === 0) {
      return `${start},0`;
    }
    if (lineCount === 1) {
      return `${start}`;
    }
    return `${start},${lineCount}`;
  };
  const trimmedFilePath = params.filePath?.trim();
  const normalizedFilePath = trimmedFilePath && trimmedFilePath.length > 0 ? trimmedFilePath : "file";
  const fileChangeType = params.fileChangeType ?? "modified";
  const oldPath = fileChangeType === "added" ? "/dev/null" : normalizedFilePath;
  const newPath = fileChangeType === "deleted" ? "/dev/null" : normalizedFilePath;
  const diffString = hunks.length === 0 ? "" : [
    `--- ${withGitDiffPrefix(oldPath, "a")}`,
    `+++ ${withGitDiffPrefix(newPath, "b")}`,
    ...hunks.flatMap((h) => [
      `@@ -${formatRange(h.oldStart, h.oldLines)} +${formatRange(h.newStart, h.newLines)} @@`,
      ...h.lines
    ])
  ].join("\n");
  const linesAdded = hunks.reduce(
    (sum, h) => sum + h.lines.filter((l) => l.startsWith("+")).length,
    0
  );
  const linesRemoved = hunks.reduce(
    (sum, h) => sum + h.lines.filter((l) => l.startsWith("-")).length,
    0
  );
  return {
    diffString,
    linesAdded,
    linesRemoved,
    observation: {
      durationMs: measure ? performance.now() - startedAt : 0,
      editLength: linesAdded + linesRemoved,
      oldChars: measure ? params.original.length : 0,
      newChars: measure ? params.new.length : 0
    }
  };
}
