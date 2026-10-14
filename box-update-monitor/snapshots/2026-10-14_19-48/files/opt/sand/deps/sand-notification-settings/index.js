// Copyright Anysphere Inc.

const binding = require("./build/Release/sand_notification_settings.node");

module.exports = {
  // () => Promise<{ platform, setting, alertSetting? } | null>
  // The calling app's OS notification permission (macOS, Windows); null
  // elsewhere or when unreadable.
  notification_settings_async: binding.notification_settings_async,
};
