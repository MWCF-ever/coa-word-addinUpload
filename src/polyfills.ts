// Polyfills for IE11
import "core-js/stable";
import "regenerator-runtime/runtime";

// Promise polyfill
if (!window.Promise) {
  window.Promise = require("promise-polyfill").default;
}

// Fetch polyfill
if (!window.fetch) {
  require("whatwg-fetch");
}

// Object.assign polyfill
if (typeof Object.assign !== "function") {
  Object.assign = require("object-assign");
}