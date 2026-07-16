"use strict";

const path = require("path");
const express = require("express");
const multer = require("multer");

const { findMatchingTemplate } = require("./generator/templates/registry");
const { extractDom } = require("./extractor/extractDom");
const { generateGenericPptx } = require("./generator/genericGenerator");
const { refineLowConfidenceSlides } = require("./generator/aiFallback");

const PORT = process.env.PORT || 8787;
const AI_FALLBACK_ENABLED = process.env.AI_FALLBACK_ENABLED !== "false" && !!process.env.ANTHROPIC_API_KEY;

const app = express();
app.use(express.json({ limit: "30mb" }));
app.use(express.static(path.join(__dirname, "..", "public")));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });

app.get("/api/health", (req, res) => {
  res.json({ ok: true, aiFallbackEnabled: AI_FALLBACK_ENABLED });
});

/**
 * POST /api/convert
 * Body: either multipart/form-data with a `file` field (an .html file),
 * or application/json { html: "<...>" }.
 * Response: the generated .pptx as a binary download.
 */
app.post("/api/convert", upload.single("file"), async (req, res) => {
  const startedAt = Date.now();
  try {
    const html = req.file ? req.file.buffer.toString("utf-8") : req.body && req.body.html;
    if (!html || typeof html !== "string" || html.trim().length === 0) {
      return res.status(400).json({ error: "No HTML provided. Upload a `file` or send JSON { html }." });
    }

    const warnings = [];
    let pptx;
    let route;

    const template = findMatchingTemplate(html);
    if (template) {
      route = `known-template:${template.name}`;
      const result = await template.generate(html, { allowGenericFallback: process.env.ALLOW_GENERIC_FALLBACK !== "false" });
      pptx = result.pptx;
      warnings.push(...result.warnings);
    } else {
      route = "generic";
      const slidesModel = await extractDom(html, { executablePath: process.env.PUPPETEER_EXECUTABLE_PATH });
      let generic = await generateGenericPptx(slidesModel, {});
      warnings.push(...generic.warnings);

      if (generic.lowConfidenceSlides.length && AI_FALLBACK_ENABLED) {
        const { slidesModel: refinedModel, warnings: aiWarnings } = await refineLowConfidenceSlides(
          slidesModel,
          generic.lowConfidenceSlides,
          {}
        );
        warnings.push(...aiWarnings);
        generic = await generateGenericPptx(refinedModel, {});
        warnings.push(...generic.warnings);
        route += "+ai-refined";
      } else if (generic.lowConfidenceSlides.length) {
        warnings.push(
          `${generic.lowConfidenceSlides.length} slide(s) flagged low-confidence but AI fallback is disabled (set ANTHROPIC_API_KEY to enable).`
        );
      }
      pptx = generic.pptx;
    }

    const buffer = await pptx.write({ outputType: "nodebuffer" });
    const tookMs = Date.now() - startedAt;
    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": 'attachment; filename="converted.pptx"',
      "X-Conversion-Route": route,
      "X-Conversion-Ms": String(tookMs),
      "X-Conversion-Warnings": String(warnings.length),
    });
    res.send(buffer);
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("convert failed:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`html2pptx-app listening on :${PORT} (AI fallback ${AI_FALLBACK_ENABLED ? "ON" : "OFF"})`);
});

module.exports = app;
