// Copyright Anysphere Inc.

const binding = require("./build/Release/sand_card_text_recognition.node");

module.exports = {
  recognize_text_async: binding.recognize_text_async,
};
