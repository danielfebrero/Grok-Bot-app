// Copyright Anysphere Inc.

const binding = require("./build/Release/sand_devicecheck.node");

module.exports = {
  isSupported: binding.isSupported,
  generateToken: binding.generateToken,
  hvVmmPresent: binding.hvVmmPresent,
};
