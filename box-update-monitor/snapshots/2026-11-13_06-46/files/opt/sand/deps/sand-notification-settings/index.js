// Copyright Anysphere Inc.

const binding = require("./build/Release/sand_notification_settings.node");

module.exports = {
  // () => Promise<{ platform, setting, alertSetting? } | null>
  // The calling app's OS notification permission (macOS, Windows); null
  // elsewhere or when unreadable.
  notification_settings_async: binding.notification_settings_async,
  // (productName, valueName) => number | string | null
  // A device-management policy value under Software\Policies\Cursor\<productName>
  // (Windows, machine before user): a number for REG_DWORD/REG_QWORD or a
  // string of a whole number, the registry type name for any other value,
  // READ_ERROR_<status> when the value exists but cannot be read, null when
  // unset or off Windows.
  managed_policy_value_sync: binding.managed_policy_value_sync,
};
