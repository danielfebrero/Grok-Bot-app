// Copyright Anysphere Inc.

const binding = require("./build/Release/sand_window_capture.node");

module.exports = {
  // () => "granted" | "denied" | "unsupported"
  screen_capture_access_sync: binding.screen_capture_access_sync,
  // () => boolean
  request_screen_capture_access_sync: binding.request_screen_capture_access_sync,
  // () => Array<{ id, pid, layer, x, y, width, height, owner }>
  list_windows_sync: binding.list_windows_sync,
  // (windowId) => Promise<CaptureResult>
  capture_window_async: binding.capture_window_async,
  // (x, y, width, height, excludeWindowIds?) => Promise<CaptureResult>
  capture_rect_async: binding.capture_rect_async,
};
