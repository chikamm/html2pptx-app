"use strict";

const pptxgen = require("pptxgenjs");
const { parseColor, alphaToTransparency, parseLinearGradient, gradientToSolidHex, parseBackgroundImageUrl } = require("../utils/colorUtils");
const { resolveImageToDataUri } = require("../utils/imageUtils");

const SAFE_FONT = "Arial";

/** Maps a CSS `font-family` computed string (e.g. `"Noto Sans JP", sans-serif`)
 *  to a font name PowerPoint is likely to have installed. We deliberately do
 *  NOT try to preserve exotic web fonts - per the pptx design guidance, an
 *  unavailable font substitutes unpredictably and can break text fit, so we
 *  normalize to a small safe set and let bold/size/color carry the emphasis. */
function mapFontFamily(cssFontFamily) {
  const f = (cssFontFamily || "").toLowerCase();
  if (f.includes("mincho") || f.includes("serif") && !f.includes("sans")) return "Yu Mincho";
  if (/noto|gothic|meiryo|hiragino|yu gothic|游ゴシック|游明朝|ms pゴシック/i.test(f)) return "Yu Gothic";
  if (f.includes("times") || f.includes("georgia") || f.includes("cambria")) return "Cambria";
  if (f.includes("courier") || f.includes("mono")) return "Courier New";
  return SAFE_FONT;
}

function textAlignToPptx(cssAlign) {
  if (cssAlign === "center") return "center";
  if (cssAlign === "right" || cssAlign === "end") return "right";
  if (cssAlign === "justify") return "justify";
  return "left";
}

/**
 * Computes a single uniform inches-per-px scale so the extracted slide's
 * own pixel geometry maps onto a PowerPoint canvas without distortion,
 * and defines that exact aspect ratio as a custom pptxgenjs layout -
 * rather than forcibly cramming every source design into 16:9.
 */
function computeLayout(pptx, widthPx, heightPx) {
  const MAX_W_IN = 13.333;
  const MAX_H_IN = 7.5;
  const scaleByWidth = MAX_W_IN / widthPx;
  const scaleByHeight = MAX_H_IN / heightPx;
  const scale = Math.min(scaleByWidth, scaleByHeight);
  const widthIn = +(widthPx * scale).toFixed(3);
  const heightIn = +(heightPx * scale).toFixed(3);
  pptx.defineLayout({ name: "SRC_ASPECT", width: widthIn, height: heightIn });
  pptx.layout = "SRC_ASPECT";
  return { scale, widthIn, heightIn };
}

function boxFillAndLine(el, scale) {
  const style = el.style || {};
  let fillColor = null;
  let transparency = undefined;

  if (style.backgroundImage) {
    const grad = parseLinearGradient(style.backgroundImage);
    const solid = grad ? gradientToSolidHex(grad) : null;
    if (solid) fillColor = solid;
  }
  if (!fillColor && style.backgroundColor) {
    const c = parseColor(style.backgroundColor);
    if (c) {
      fillColor = c.hex;
      transparency = alphaToTransparency(c.alpha);
    }
  }

  let line = { type: "none" };
  if (style.borderWidth > 0 && style.borderStyle && style.borderStyle !== "none") {
    const bc = parseColor(style.borderColor);
    if (bc) {
      line = { color: bc.hex, width: Math.max(0.25, style.borderWidth * scale * 72) };
    }
  }

  return {
    fill: fillColor ? { color: fillColor, transparency } : { type: "none" },
    line,
  };
}

function rectRadiusFor(el, w, h, scale) {
  const wanted = el.style.borderRadiusPx * scale;
  return Math.min(w, h) / 2 > wanted ? wanted : Math.min(w, h) / 2;
}

function shapeTypeFor(pptx, el) {
  if (el.shape === "ellipse") return pptx.ShapeType.ellipse;
  if (el.shape === "roundRect") return pptx.ShapeType.roundRect;
  return pptx.ShapeType.rect;
}

function buildRichTextRuns(el, scale) {
  const baseColor = parseColor(el.style.color);
  const baseFont = mapFontFamily(el.style.fontFamily);
  const baseSizePt = Math.max(6, el.style.fontSizePx * scale * 72);
  const runs = el.runs && el.runs.length ? el.runs : [{ text: el.text, style: {} }];

  const out = [];
  runs.forEach((run, i) => {
    if (run.isBreak) {
      if (out.length) out[out.length - 1].options.breakLine = true;
      return;
    }
    if (run.text === "") return;
    const runColor = run.style && run.style.color ? parseColor(run.style.color) : null;
    out.push({
      text: run.text,
      options: {
        fontFace: baseFont,
        fontSize: baseSizePt,
        color: (runColor || baseColor || { hex: "1A1A1A" }).hex,
        bold: !!(run.style && run.style.bold),
        italic: !!(run.style && run.style.italic),
        underline: run.style && run.style.underline ? { style: "sng" } : undefined,
      },
    });
  });
  // last run should not force an extra trailing line break
  if (out.length && out[out.length - 1].options.breakLine) {
    delete out[out.length - 1].options.breakLine;
  }
  return out;
}

async function renderElement(slide, el, scale, opts) {
  const x = el.rect.x * scale;
  const y = el.rect.y * scale;
  const w = Math.max(0.02, el.rect.w * scale);
  const h = Math.max(0.02, el.rect.h * scale);

  if (el.type === "box") {
    const bgUrl = el.style.backgroundImage ? parseBackgroundImageUrl(el.style.backgroundImage) : null;

    if (bgUrl) {
      // CSS `background-image: url(...)` (a raster photo, as opposed to a
      // gradient) - the extractor captures the raw string but only a
      // gradient can be approximated as a fill color, so a url() was
      // previously silently dropped. Render it as a cover-fit image
      // instead, then draw the border on top (if any) as an unfilled
      // outline so rounded corners / borders still work visually.
      try {
        const dataUri = await resolveImageToDataUri(bgUrl, opts.baseUrl);
        if (dataUri) {
          slide.addImage({ data: dataUri, x, y, w, h, sizing: { type: "cover", w, h } });
        }
      } catch (e) {
        opts.warnings.push(`background image skipped (${bgUrl}): ${e.message}`);
      }
      const { line } = boxFillAndLine(el, scale);
      if (line.type !== "none") {
        const outlineOpts = { x, y, w, h, fill: { type: "none" }, line };
        if (el.shape === "roundRect" && el.style.borderRadiusPx) {
          outlineOpts.rectRadius = rectRadiusFor(el, w, h, scale);
        }
        slide.addShape(shapeTypeFor(opts.pptx, el), outlineOpts);
      }
      return;
    }

    const { fill, line } = boxFillAndLine(el, scale);
    if (fill.type === "none" && line.type === "none") return;
    const shapeOpts = { x, y, w, h, fill, line };
    if (el.shape === "roundRect" && el.style.borderRadiusPx) {
      shapeOpts.rectRadius = rectRadiusFor(el, w, h, scale);
    }
    slide.addShape(shapeTypeFor(opts.pptx, el), shapeOpts);
    return;
  }

  if (el.type === "text") {
    if (!el.text || !el.text.trim()) return;
    const runs = buildRichTextRuns(el, scale);
    if (!runs.length) return;
    slide.addText(runs, {
      x, y, w, h,
      align: textAlignToPptx(el.style.textAlign),
      valign: "top",
      fontFace: mapFontFamily(el.style.fontFamily),
      autoFit: false,
      wrap: true,
      margin: 0,
    });
    return;
  }

  if (el.type === "image") {
    try {
      const dataUri = await resolveImageToDataUri(el.src, opts.baseUrl);
      if (dataUri) slide.addImage({ data: dataUri, x, y, w, h, sizing: { type: "cover", w, h } });
    } catch (e) {
      opts.warnings.push(`image skipped (${el.src}): ${e.message}`);
    }
    return;
  }

  if (el.type === "svg") {
    if (!opts.rasterizeSvg) {
      opts.warnings.push("svg element skipped (no rasterizer configured)");
      return;
    }
    try {
      const png = await opts.rasterizeSvg(el.outerHTML, Math.round(w * 96), Math.round(h * 96));
      if (png) slide.addImage({ data: `image/png;base64,${png.toString("base64")}`, x, y, w, h });
    } catch (e) {
      opts.warnings.push(`svg rasterize failed: ${e.message}`);
    }
  }
}

/** Very cheap heuristic auto-QA: flags a slide as "low confidence" (a
 *  candidate for the AI-refinement fallback pass) when a large fraction of
 *  its text boxes overlap each other or spill outside the canvas - the two
 *  defect classes generic geometry extraction is most prone to (e.g. when
 *  the source used CSS transforms, negative margins, or absolutely
 *  positioned overlays the box model doesn't fully capture). */
function computeConfidence(elements, widthPx, heightPx) {
  const texts = elements.filter((e) => e.type === "text");
  let overlapPairs = 0;
  let outOfBounds = 0;
  for (const t of texts) {
    if (t.rect.x < -2 || t.rect.y < -2 || t.rect.x + t.rect.w > widthPx + 2 || t.rect.y + t.rect.h > heightPx + 2) {
      outOfBounds++;
    }
  }
  for (let i = 0; i < texts.length; i++) {
    for (let j = i + 1; j < texts.length; j++) {
      const a = texts[i].rect, b = texts[j].rect;
      const ix = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
      const iy = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
      const interArea = ix * iy;
      const minArea = Math.min(a.w * a.h, b.w * b.h) || 1;
      if (interArea / minArea > 0.5) overlapPairs++;
    }
  }
  const issues = overlapPairs + outOfBounds;
  return { score: issues === 0 ? 1 : Math.max(0, 1 - issues / Math.max(4, texts.length)), overlapPairs, outOfBounds };
}

/**
 * Renders one already-extracted slide model into an existing pptxgenjs
 * slide. Exposed separately from generateGenericPptx() so specialized /
 * known-template generators (server/generator/templates/*) can fall back
 * to generic reconstruction for individual slides they don't recognize,
 * instead of only being able to use the generic engine for a whole deck.
 *
 * @returns {{ confidence: {score:number, overlapPairs:number, outOfBounds:number} }}
 */
async function renderSlideModelIntoPptx(pptx, slide, slideModel, scale, options = {}) {
  const warnings = options.warnings || [];
  const bgUrl = slideModel.backgroundImage ? parseBackgroundImageUrl(slideModel.backgroundImage) : null;

  if (bgUrl) {
    try {
      const dataUri = await resolveImageToDataUri(bgUrl, options.baseUrl);
      if (dataUri) slide.background = { data: dataUri };
    } catch (e) {
      warnings.push(`slide background image skipped (${bgUrl}): ${e.message}`);
    }
  } else {
    const bgGrad = slideModel.backgroundImage ? parseLinearGradient(slideModel.backgroundImage) : null;
    const bgSolidFromGrad = bgGrad ? gradientToSolidHex(bgGrad) : null;
    const bgColor = bgSolidFromGrad || (parseColor(slideModel.backgroundColor) || {}).hex;
    if (bgColor) slide.background = { color: bgColor };
  }

  for (const el of slideModel.elements) {
    // eslint-disable-next-line no-await-in-loop
    await renderElement(slide, el, scale, { pptx, warnings, baseUrl: options.baseUrl, rasterizeSvg: options.rasterizeSvg });
  }

  const confidence = computeConfidence(slideModel.elements, slideModel.widthPx, slideModel.heightPx);
  return { confidence };
}

/**
 * @param {Array} slidesModel - output of extractDom()
 * @param {object} [options]
 * @param {string} [options.baseUrl] - used to resolve relative <img src>
 * @param {function} [options.rasterizeSvg] - async (svgString, wPx, hPx) => Buffer(png)
 * @returns {Promise<{pptx: pptxgen, lowConfidenceSlides: number[], warnings: string[]}>}
 */
async function generateGenericPptx(slidesModel, options = {}) {
  const pptx = new pptxgen();
  pptx.author = "html2pptx-app";
  const warnings = [];
  const lowConfidenceSlides = [];

  if (!slidesModel.length) throw new Error("No slides were extracted from the supplied HTML.");

  // Use the first slide's aspect ratio for the whole deck (mixed aspect
  // ratios within one deck are rare and PPTX only supports one canvas size).
  const { scale } = computeLayout(pptx, slidesModel[0].widthPx, slidesModel[0].heightPx);

  for (let i = 0; i <
