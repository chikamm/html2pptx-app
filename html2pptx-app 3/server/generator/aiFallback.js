"use strict";

/**
 * AI fallback layer: the third leg of the hybrid pipeline.
 *
 *   1. Known-template match  -> famimaFrappeTemplate.js (curated, exact)
 *   2. Generic geometry rules -> genericGenerator.js     (works on anything,
 *                                                          occasionally messy)
 *   3. AI refinement          -> THIS FILE, only for the specific slides
 *                                 genericGenerator.js's computeConfidence()
 *                                 flagged as low-confidence (overlapping
 *                                 text boxes / content spilling outside the
 *                                 canvas) - i.e. exactly the cases where a
 *                                 human (or this session's manual QA loop)
 *                                 would normally step in and nudge coordinates
 *                                 by hand.
 *
 * Design choice: instead of asking the model to free-write pptxgenjs code
 * (which would need to be eval'd - a real security and stability risk in a
 * server handling arbitrary uploaded HTML), we ask it to return a corrected
 * array of elements in OUR OWN schema (the same shape extractDom() already
 * produces), via forced tool-use so the response is structurally
 * constrained JSON. That JSON is then run through the exact same
 * renderSlideModelIntoPptx() renderer the generic engine uses - the model
 * is only ever allowed to move/resize/restyle boxes, text and images, never
 * to execute anything.
 *
 * Requires ANTHROPIC_API_KEY. If unset, refineSlide() is a no-op that
 * returns the original slide model untouched and logs a warning - the
 * generic engine's output is still shipped, just without this polish pass.
 */

const Anthropic = require("@anthropic-ai/sdk");

const MODEL = process.env.ANTHROPIC_MODEL || "claude-sonnet-4-5";

const ELEMENT_SCHEMA = {
  type: "object",
  properties: {
    elements: {
      type: "array",
      description: "Corrected drawable elements, same order = same paint order (later = on top).",
      items: {
        type: "object",
        properties: {
          type: { type: "string", enum: ["box", "text", "image"] },
          shape: { type: "string", enum: ["rect", "roundRect", "ellipse"], description: "only for type=box" },
          rect: {
            type: "object",
            properties: {
              x: { type: "number" }, y: { type: "number" }, w: { type: "number" }, h: { type: "number" },
            },
            required: ["x", "y", "w", "h"],
          },
          text: { type: "string", description: "only for type=text" },
          src: { type: "string", description: "only for type=image: pass through the original src unchanged" },
          style: {
            type: "object",
            description: "same fields as the input model's `style` object for that element type",
          },
        },
        required: ["type", "rect"],
      },
    },
  },
  required: ["elements"],
};

function buildSystemPrompt() {
  return [
    "You are fixing the layout of ONE slide that was auto-extracted from an HTML/CSS design into a flat list of primitive elements (box/text/image) with pixel coordinates relative to the slide's top-left corner.",
    "The automatic extraction flagged this slide as low-confidence: some text elements overlap each other or extend outside the slide canvas.",
    "You will receive: the slide's pixel width/height, the full element list (with each element's current rect and style), and optionally a screenshot image of how the ORIGINAL html rendered (the ground truth to match).",
    "Your job: return a corrected element list in the exact same schema, with rect coordinates (and only coordinates/wrapping-relevant style like fontSize) adjusted so nothing overlaps and nothing spills outside the canvas, while keeping every element's role, text content and relative position/order as close to the original as possible.",
    "Do not invent new elements or remove elements that hold real text content. Do not change `text` or `src` values. Only adjust rect/style fields to resolve the specific overlap/out-of-bounds issues.",
    "Call the `return_fixed_slide` tool exactly once with the corrected element array.",
  ].join("\n");
}

/**
 * @param {object} slideModel - one entry from extractDom()'s slides array
 * @param {object} [options]
 * @param {string} [options.apiKey] - defaults to process.env.ANTHROPIC_API_KEY
 * @param {Buffer} [options.screenshotPng] - optional reference screenshot
 * @returns {Promise<{slideModel: object, refined: boolean, reason?: string}>}
 */
async function refineSlide(slideModel, options = {}) {
  const apiKey = options.apiKey || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return { slideModel, refined: false, reason: "ANTHROPIC_API_KEY not set - skipping AI refinement pass." };
  }

  const client = new Anthropic({ apiKey });

  const userContent = [
    {
      type: "text",
      text: `Slide canvas: ${slideModel.widthPx}x${slideModel.heightPx}px.\n\nCurrent elements (JSON):\n${JSON.stringify(
        slideModel.elements,
        null,
        2
      )}`,
    },
  ];
  if (options.screenshotPng) {
    userContent.unshift({
      type: "image",
      source: { type: "base64", media_type: "image/png", data: options.screenshotPng.toString("base64") },
    });
  }

  let response;
  try {
    response = await client.messages.create({
      model: MODEL,
      max_tokens: 8000,
      system: buildSystemPrompt(),
      messages: [{ role: "user", content: userContent }],
      tools: [{ name: "return_fixed_slide", description: "Return the corrected element list.", input_schema: ELEMENT_SCHEMA }],
      tool_choice: { type: "tool", name: "return_fixed_slide" },
    });
  } catch (err) {
    return { slideModel, refined: false, reason: `Anthropic API call failed: ${err.message}` };
  }

  const toolUse = (response.content || []).find((c) => c.type === "tool_use" && c.name === "return_fixed_slide");
  if (!toolUse || !toolUse.input || !Array.isArray(toolUse.input.elements)) {
    return { slideModel, refined: false, reason: "Model response did not include a usable tool_use block." };
  }

  // Merge conservatively: keep original element for any index the model
  // dropped or corrupted (missing type/rect), only accept well-formed ones.
  const fixed = toolUse.input.elements
    .map((el, i) => {
      const orig = slideModel.elements[i] || {};
      if (!el || typeof el !== "object" || !el.rect || !el.type) return orig;
      return { ...orig, ...el, rect: { ...orig.rect, ...el.rect }, style: { ...(orig.style || {}), ...(el.style || {}) } };
    })
    .filter(Boolean);

  return { slideModel: { ...slideModel, elements: fixed }, refined: true };
}

/**
 * Convenience batch helper used by the server pipeline: refines only the
 * slide indices the generic engine flagged, leaves the rest untouched.
 */
async function refineLowConfidenceSlides(slidesModel, lowConfidenceIndices, options = {}) {
  const warnings = [];
  const result = slidesModel.slice();
  for (const idx of lowConfidenceIndices) {
    // eslint-disable-next-line no-await-in-loop
    const { slideModel, refined, reason } = await refineSlide(slidesModel[idx], options);
    result[idx] = slideModel;
    if (!refined) warnings.push(`Slide ${idx + 1}: ${reason}`);
  }
  return { slidesModel: result, warnings };
}

module.exports = { refineSlide, refineLowConfidenceSlides };
