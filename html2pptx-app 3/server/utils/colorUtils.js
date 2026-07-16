"use strict";

/**
 * getComputedStyle() always resolves colors to `rgb(r, g, b)` or
 * `rgba(r, g, b, a)` strings in a real browser (Puppeteer/Chromium),
 * regardless of how the author wrote the CSS (hex, named color, hsl...).
 * These helpers convert that canonical form into what pptxgenjs wants:
 * a bare 6-digit hex with no `#`, plus a separate 0-100 transparency value.
 */

function clamp255(n) {
  return Math.max(0, Math.min(255, Math.round(n)));
}

function toHex2(n) {
  return clamp255(n).toString(16).padStart(2, "0").toUpperCase();
}

/**
 * Parses `rgb(r,g,b)` / `rgba(r,g,b,a)` (and, defensively, `#rrggbb`)
 * into { hex: "RRGGBB", alpha: 0..1 } or null if the color is missing /
 * fully transparent / not parseable (e.g. `transparent`, `none`).
 */
function parseColor(cssColor) {
  if (!cssColor || typeof cssColor !== "string") return null;
  const s = cssColor.trim().toLowerCase();
  if (s === "transparent" || s === "none" || s === "") return null;

  const rgbMatch = s.match(
    /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/
  );
  if (rgbMatch) {
    const [, r, g, b, a] = rgbMatch;
    const alpha = a === undefined ? 1 : parseFloat(a);
    if (alpha <= 0.02) return null; // effectively invisible
    return { hex: toHex2(+r) + toHex2(+g) + toHex2(+b), alpha };
  }

  const hexMatch = s.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hexMatch) {
    let hex = hexMatch[1];
    if (hex.length === 3) {
      hex = hex.split("").map((c) => c + c).join("");
    }
    return { hex: hex.toUpperCase(), alpha: 1 };
  }
  return null;
}

/**
 * pptxgenjs transparency is 0 (opaque) .. 100 (fully transparent) - the
 * inverse of CSS alpha (1 = opaque). This converts alpha -> transparency.
 */
function alphaToTransparency(alpha) {
  return Math.round((1 - alpha) * 100);
}

/**
 * Chrome normalizes `background-image: linear-gradient(...)` /
 * `radial-gradient(...)` computed style into a string like:
 *   "linear-gradient(135deg, rgb(181, 0, 18) 0%, rgb(255, 13, 20) 100%)"
 *   "radial-gradient(at 25% 20%, rgb(249, 115, 22) 0%, rgb(10, 14, 31) 100%)"
 * pptxgenjs has no gradient-fill API (documented limitation), so we
 * extract the stops and let the caller decide how to approximate it -
 * typically either the first stop, or a manually blended midpoint color,
 * or (best fidelity) rendering a tiny gradient PNG and using it as an
 * image fill / background image. Handles both gradient types the same
 * way since only the color stops matter for the solid-color fallback;
 * the leading angle/shape/position token (e.g. `135deg`, `at 25% 20%`)
 * simply fails `parseColor()` and is skipped.
 */
function parseGradient(cssBackgroundImage) {
  if (!cssBackgroundImage) return null;
  const inner = cssBackgroundImage.match(/(?:linear|radial)-gradient\(([^]*)\)/);
  if (!inner) return null;
  // split on commas that are NOT inside rgb(...)/rgba(...)
  const parts = [];
  let depth = 0;
  let cur = "";
  for (const ch of inner[1]) {
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      parts.push(cur.trim());
      cur = "";
    } else {
      cur += ch;
    }
  }
  if (cur.trim()) parts.push(cur.trim());

  let angleDeg = 135;
  const stops = [];
  parts.forEach((p) => {
    const angleMatch = p.match(/^(-?[\d.]+)deg$/);
    if (angleMatch) {
      angleDeg = parseFloat(angleMatch[1]);
      return;
    }
    const posMatch = p.match(/^(.*?)(?:\s+([\d.]+)%)?$/);
    const colorPart = posMatch ? posMatch[1].trim() : p;
    const pos = posMatch && posMatch[2] !== undefined ? parseFloat(posMatch[2]) : null;
    const color = parseColor(colorPart);
    if (color) stops.push({ color, pos });
  });
  if (stops.length < 2) return null;
  // fill in missing positions evenly
  stops.forEach((s, i) => {
    if (s.pos === null) s.pos = (i / (stops.length - 1)) * 100;
  });
  return { angleDeg, stops };
}

// Backward-compatible alias (only gradient parser in this module historically
// handled linear-gradient; now handles radial-gradient too).
const parseLinearGradient = parseGradient;

/** Blends the gradient into a single representative solid hex color
 *  (midpoint-weighted average) - used as the v1 fallback for slide/shape
 *  backgrounds since pptxgenjs cannot express true gradient fills. */
function gradientToSolidHex(gradient) {
  if (!gradient || !gradient.stops || !gradient.stops.length) return null;
  const mid = gradient.stops[Math.floor((gradient.stops.length - 1) / 2)];
  return mid.color.hex;
}

/**
 * Extracts the URL out of a computed `background-image` value that
 * contains a raster image, e.g. `url("https://example.com/photo.jpg")`,
 * or a combined layer like `linear-gradient(...), url(...)`. Chrome's
 * computed style always resolves relative paths to absolute URLs and
 * wraps the value in `url("...")`, so a single regex is sufficient.
 * Returns null if no `url(...)` layer is present (e.g. gradient-only).
 */
function parseBackgroundImageUrl(cssBackgroundImage) {
  if (!cssBackgroundImage || !cssBackgroundImage.includes("url(")) return null;
  const m = cssBackgroundImage.match(/url\((?:"([^"]*)"|'([^']*)'|([^)]*))\)/);
  if (!m) return null;
  return m[1] || m[2] || m[3] || null;
}

module.exports = {
  parseColor,
  alphaToTransparency,
  parseGradient,
  parseLinearGradient,
  gradientToSolidHex,
  parseBackgroundImageUrl,
};
