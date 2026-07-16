"use strict";

/**
 * End-to-end smoke test using a REAL deck (this project's own earlier
 * conversion job) run through the hybrid pipeline's template-matching
 * path: registry.findMatchingTemplate() should recognize the file as the
 * in-house proposal template and hand it to famimaFrappeTemplate.generate(),
 * which parses the real HTML with cheerio (not hardcoded strings) and
 * reconstructs the deck with native PPTX shapes/text/images.
 */

const fs = require("fs");
const path = require("path");
const { findMatchingTemplate } = require("../server/generator/templates/registry");

const samplePath = path.join(__dirname, "..", "samples", "fruits_frappe.html");
const html = fs.readFileSync(samplePath, "utf-8");

(async () => {
  const template = findMatchingTemplate(html);
  console.log("Matched template:", template ? template.name : "(none - would use generic engine)");
  if (!template) process.exit(1);

  const { pptx, warnings } = await template.generate(html, { allowGenericFallback: false });
  const outPath = path.join(__dirname, "..", "tmp_sample_output.pptx");
  await pptx.writeFile({ fileName: outPath });
  console.log("Wrote:", outPath);
  console.log("Slide count:", pptx.slides.length);
  console.log("Warnings:", warnings);
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
