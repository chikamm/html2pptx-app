"use strict";

const fs = require("fs");
const path = require("path");

const MIME_BY_EXT = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".bmp": "image/bmp",
};

/**
 * Normalizes an <img src> (already resolved to an absolute URL by the
 * browser's `currentSrc`/`src` getter during extraction) into a
 * `data:<mime>;base64,...` string pptxgenjs can embed directly.
 *
 * - `data:` URIs pass straight through.
 * - `http(s)://` URLs are fetched.
 * - `file://` or bare paths are read from disk (useful when the source
 *   HTML was rendered from a local file with relative image paths).
 */
async function resolveImageToDataUri(src) {
  if (!src) return null;
  if (src.startsWith("data:")) return src;

  if (src.startsWith("http://") || src.startsWith("https://")) {
    const res = await fetch(src);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    const mime = res.headers.get("content-type") || guessMime(src);
    return `data:${mime};base64,${buf.toString("base64")}`;
  }

  const localPath = src.startsWith("file://") ? src.slice(7) : src;
  const buf = await fs.promises.readFile(localPath);
  return `data:${guessMime(localPath)};base64,${buf.toString("base64")}`;
}

function guessMime(p) {
  const ext = path.extname(p).toLowerCase();
  return MIME_BY_EXT[ext] || "image/png";
}

module.exports = { resolveImageToDataUri };
