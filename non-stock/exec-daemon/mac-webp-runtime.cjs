"use strict";

/**
 * Webpack parses `import()` even inside `new Function("return import(s)")` and
 * builds a context module that inlines HTML as JavaScript. Keep the native
 * ESM load in this never-bundled helper; mac-webp.ts createRequire()s it.
 */
exports.importFile = specifier => import(specifier);
