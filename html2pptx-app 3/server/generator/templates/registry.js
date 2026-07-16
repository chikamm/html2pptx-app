"use strict";

/**
 * Known-template registry: each entry is a { name, fingerprint(html),
 * generate(html, options) } triplet. `fingerprint` is a cheap string/regex
 * test run BEFORE we spend any time on the (expensive, Chromium-backed)
 * generic extraction pipeline - if a deck matches a registered in-house
 * template, we hand it to that template's hand-tuned generator instead,
 * which produces much higher fidelity output than generic geometry
 * reconstruction because it understands the *meaning* of each block
 * (e.g. "this is a Q&A card", not just "a rounded rect with two text
 * children").
 *
 * To add support for another recurring internal template: create a new
 * file in this directory exporting { name, fingerprint, generate } and
 * require() it below to register it.
 */

const templates = [];

function registerTemplate(def) {
  if (!def || typeof def.fingerprint !== "function" || typeof def.generate !== "function") {
    throw new Error("registerTemplate: definition needs fingerprint() and generate()");
  }
  templates.push(def);
}

function findMatchingTemplate(html) {
  return templates.find((t) => {
    try {
      return t.fingerprint(html);
    } catch {
      return false;
    }
  }) || null;
}

module.exports = { registerTemplate, findMatchingTemplate, templates };

// Side-effect registrations - each require() call adds one template.
require("./famimaFrappeTemplate");
